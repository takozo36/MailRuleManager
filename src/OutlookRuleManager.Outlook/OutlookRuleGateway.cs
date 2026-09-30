using System.Runtime.InteropServices;
using OutlookRuleManager.Core;
using static OutlookRuleManager.Outlook.OutlookCom;

namespace OutlookRuleManager.Outlook;

public sealed record StoreInfo(string StoreId, string DisplayName, bool IsDefault)
{
    public override string ToString() => DisplayName;
}

/// <summary>Progress while loading: Rule is one rule that has been read, Total is the number of rules.</summary>
public sealed record LoadProgress(RuleData Rule, int Total);

/// <summary>At save time, Outlook's rules turned out to differ from the state at load time.</summary>
public sealed class RulesChangedException(string message) : Exception(message);

/// <summary>
/// Loads and saves classic Outlook rules.
/// Meant to be called from a thread other than the UI thread (calls take from seconds to tens of seconds).
/// Outlook objects are handled with late binding (dynamic); see OutlookCom.
/// </summary>
public sealed class OutlookRuleGateway : IDisposable
{
    private dynamic? _app;
    private dynamic? _session;

    private dynamic Session
    {
        get
        {
            if (_session is null)
            {
                _app = CreateApplication();
                _session = _app.GetNamespace("MAPI");
            }
            return _session;
        }
    }

    public IReadOnlyList<StoreInfo> ListStores() => Guard(() =>
    {
        string defaultId = (string)Session.DefaultStore.StoreID;
        var list = new List<StoreInfo>();
        foreach (dynamic s in Session.Stores)
        {
            string id = (string)s.StoreID;
            list.Add(new StoreInfo(id, (string)s.DisplayName, id == defaultId));
        }
        return list;
    });

    /// <summary>Which method the last LoadRules used (for display).</summary>
    public string LastLoadNote { get; private set; } = "";

    /// <summary>
    /// Loads all rules.
    /// First reads the rules stream that holds all rules (PR_RW_RULES_STREAM in a hidden message in the Inbox) in one
    /// call and parses it (a few seconds). If it cannot be read or parsed, or the count or names do not match Outlook,
    /// falls back to asking Outlook for each condition and action one by one (about 0.3 s per rule).
    /// Rules are passed to progress as they are read, so the UI can show them while loading.
    /// </summary>
    /// <param name="allowFast">When false, always reads rule by rule (for verification).</param>
    public IReadOnlyList<RuleData> LoadRules(string storeId, IProgress<LoadProgress>? progress, CancellationToken cancel, bool allowFast = true) => Guard(() =>
    {
        var rules = GetRules(storeId);
        int count = (int)rules.Count;

        string? reason = null;
        if (allowFast)
        {
            var fast = TryLoadFast(storeId, rules, count, cancel, out reason);
            if (fast is not null)
            {
                foreach (var r in fast) progress?.Report(new LoadProgress(r, count));
                LastLoadNote = Loc.T("高速読み込み", "fast load");
                return fast;
            }
        }

        var list = new List<RuleData>(count);
        for (int i = 1; i <= count; i++)
        {
            cancel.ThrowIfCancellationRequested();
            var rule = OutlookRuleReader.Read((object)rules[i], i);
            list.Add(rule);
            progress?.Report(new LoadProgress(rule, count));
        }
        LastLoadNote = reason is null
            ? Loc.T("1 件ずつ読み込み", "rule-by-rule load")
            : Loc.T($"1 件ずつ読み込み（高速読み込みできなかった理由: {reason}）", $"rule-by-rule load (fast load not possible: {reason})");
        return list;
    });

    /// <summary>
    /// Loads from the rules stream. Returns null and a reason on failure (does not throw).
    /// The parsed result is used only after the count and the name / enabled flag at each position match Outlook's rules.
    /// </summary>
    private List<RuleData>? TryLoadFast(string storeId, dynamic rules, int count, CancellationToken cancel, out string? reason)
    {
        IReadOnlyList<StreamRule> parsed;
        try
        {
            parsed = RulesStream.Parse(ReadRulesStream(storeId));
        }
        catch (RulesStreamFormatException ex) { reason = ex.Message; return null; }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            reason = Loc.T("ルールのまとめデータを取り出せません: ", "cannot read the rules stream: ") + ex.Message;
            return null;
        }

        if (parsed.Count != count)
        {
            reason = Loc.T($"件数が一致しません（まとめデータ {parsed.Count} 件、Outlook {count} 件）",
                $"rule count mismatch (stream {parsed.Count}, Outlook {count})");
            return null;
        }

        var folders = new Dictionary<string, FolderRef?>(StringComparer.OrdinalIgnoreCase);
        var list = new List<RuleData>(count);
        for (int i = 1; i <= count; i++)
        {
            cancel.ThrowIfCancellationRequested();
            var p = parsed[i - 1];
            dynamic rule = rules[i];
            if ((string)rule.Name != p.Name || (bool)rule.Enabled != p.Enabled)
            {
                reason = Loc.T($"{i} 番目のルールの名前か有効/無効が一致しません", $"name or enabled flag of rule #{i} does not match");
                return null;
            }
            // Whether "on this computer only" means this computer cannot be told from the stream, so ask Outlook
            bool onThisMachine = !p.NeedsMachineCheck || (bool)rule.Conditions.OnLocalMachine.Enabled;
            list.Add(p.ToRuleData(f => ResolveFolder(f, folders), onThisMachine, (bool)rule.IsLocalRule));
        }
        reason = null;
        return list;
    }

    /// <summary>Reads PR_RW_RULES_STREAM from the hidden IPM.RuleOrganizer message in the Inbox.</summary>
    private byte[] ReadRulesStream(string storeId)
    {
        dynamic store = GetStore(storeId);
        dynamic inbox = store.GetDefaultFolder(OlFolderInbox);
        dynamic table = inbox.GetTable("", OlHiddenItems);
        table.Columns.Add("MessageClass");
        table.Columns.Add("EntryID");
        string? entryId = null;
        while (!(bool)table.EndOfTable)
        {
            dynamic row = table.GetNextRow();
            if ((string)row["MessageClass"] == "IPM.RuleOrganizer") { entryId = (string)row["EntryID"]; break; }
        }
        if (entryId is null)
            throw new InvalidOperationException(Loc.T("受信トレイに IPM.RuleOrganizer がありません", "IPM.RuleOrganizer not found in the Inbox"));
        dynamic item = Session.GetItemFromID(entryId, storeId);
        return (byte[])item.PropertyAccessor.GetProperty(RulesStreamProperty);
    }

    private const string RulesStreamProperty = "http://schemas.microsoft.com/mapi/proptag/0x68020102";

    /// <summary>Resolves an EntryID from the rules stream to an Outlook folder (each folder is looked up only once).</summary>
    private FolderRef? ResolveFolder(FolderRef raw, Dictionary<string, FolderRef?> cache)
    {
        if (cache.TryGetValue(raw.EntryId, out var hit)) return hit;
        FolderRef? result = null;
        try
        {
            object? f = raw.StoreId.Length > 0
                ? (object?)Session.GetFolderFromID(raw.EntryId, raw.StoreId)
                : (object?)Session.GetFolderFromID(raw.EntryId);
            if (f is not null) result = OutlookRuleReader.ToFolderRef(f);
        }
        catch
        {
            try
            {
                object? f = Session.GetFolderFromID(raw.EntryId);
                if (f is not null) result = OutlookRuleReader.ToFolderRef(f);
            }
            catch { /* not found = deleted or moved folder */ }
        }
        cache[raw.EntryId] = result;
        return result;
    }

    /// <summary>Tree of mail folders (the top level is the folders directly under the store, such as the Inbox).</summary>
    public IReadOnlyList<FolderNode> LoadFolders(string storeId) => Guard(() =>
    {
        return ReadChildren((object)GetStore(storeId).GetRootFolder());
    });

    private static List<FolderNode> ReadChildren(dynamic parent)
    {
        var list = new List<FolderNode>();
        foreach (dynamic f in parent.Folders)
        {
            if ((int)f.DefaultItemType != OlMailItem) continue;
            list.Add(new FolderNode((string)f.Name, OutlookRuleReader.ToFolderRef((object)f), ReadChildren((object)f)));
        }
        return list;
    }

    /// <summary>
    /// Applies the edits to Outlook and saves them.
    /// Nothing is persisted in Outlook until Rules.Save() is called, so if a step fails, the method throws without saving.
    /// </summary>
    /// <param name="dryRun">When true, skips only the final Rules.Save() (to check that all steps work; nothing is kept in Outlook).</param>
    public void Apply(string storeId, ChangePlan plan, IProgress<string>? progress, bool dryRun = false) => Guard(() =>
    {
        var problems = plan.Validate();
        if (problems.Count > 0) throw new InvalidOperationException(string.Join(Environment.NewLine, problems));

        progress?.Report(Loc.T("Outlook のルールを確認しています…", "Checking Outlook's rules…"));
        var rules = GetRules(storeId);

        // 1. Check that Outlook's rules have not changed since loading (compare names by position)
        int count = (int)rules.Count;
        if (count != plan.Original.Count)
            throw new RulesChangedException(Loc.T(
                $"Outlook 側のルールの数が変わっています（読み込み時 {plan.Original.Count} 件 → 現在 {count} 件）。",
                $"The number of rules in Outlook has changed ({plan.Original.Count} when loaded → {count} now)."));
        var originals = new object[count + 1];
        for (int i = 1; i <= count; i++)
        {
            originals[i] = rules[i];
            string name = (string)((dynamic)originals[i]).Name;
            if (name != plan.Original[i - 1].Name)
                throw new RulesChangedException(Loc.T(
                    $"Outlook 側の {i} 番目のルールが読み込み時と違います（「{plan.Original[i - 1].Name}」→「{name}」）。",
                    $"Rule #{i} in Outlook differs from when it was loaded (\"{plan.Original[i - 1].Name}\" → \"{name}\")."));
        }

        // 2. Name, enabled flag and move-to folder of existing rules
        progress?.Report(Loc.T("変更を反映しています…", "Applying changes…"));
        var byId = new Dictionary<string, object>();
        foreach (var e in plan.FinalOrder.Where(e => !e.IsNew))
        {
            dynamic r = originals[e.Source.Index];
            byId[e.Id] = r;
            Step(e.Name, () =>
            {
                if (e.IsRenamed) r.Name = e.Name;
                if (e.MoveFolderOverride is { } folder) SetProperty((object)r.Actions.MoveToFolder, "Folder", GetFolder(folder));
                if (e.IsEnabledChanged) r.Enabled = e.Enabled;
            });
        }

        // 3. Duplicates (before deleting: touching a removed rule object crashes Outlook)
        foreach (var e in plan.Created)
        {
            progress?.Report(Loc.T($"複製しています: {e.Name}", $"Duplicating: {e.Name}"));
            Step(e.Name, () =>
            {
                dynamic created = rules.Create(e.Name, (int)e.Kind);
                byId[e.Id] = (object)created;
                OutlookRuleWriter.Copy(originals[e.Source.Index], (object)created);
                if (e.MoveFolderOverride is { } folder) SetProperty((object)created.Actions.MoveToFolder, "Folder", GetFolder(folder));
                created.Enabled = e.Enabled;

                // Read it back before saving and check that its content matches the original (with the folder override)
                var copied = OutlookRuleReader.Read((object)created, 0);
                string expected = RuleSignature.Of(e.Conditions, e.Exceptions, e.Actions);
                string actual = copied.ReadError ?? RuleSignature.Of(copied);
                if (actual != expected)
                    throw new InvalidOperationException(Loc.T(
                        $"複製した内容が元のルールと一致しません。{Environment.NewLine}元: {expected}{Environment.NewLine}複製: {actual}",
                        $"The duplicate does not match the original rule.{Environment.NewLine}Original: {expected}{Environment.NewLine}Duplicate: {actual}"));
            });
        }

        // 4. Order. Placing the kept rules from position 1 pushes the rules to delete to the end
        progress?.Report(Loc.T("並び順を反映しています…", "Applying the order…"));
        for (int k = 0; k < plan.FinalOrder.Count; k++)
        {
            dynamic r = byId[plan.FinalOrder[k].Id];
            if ((int)r.ExecutionOrder != k + 1) r.ExecutionOrder = k + 1;
        }

        // 5. Delete. Remove the last rule only after confirming by object identity that it is one to delete
        var toDelete = plan.Deleted.Select(d => originals[d.Index]).ToList();
        while ((int)rules.Count > plan.FinalOrder.Count)
        {
            int last = (int)rules.Count;
            object r = rules[last];
            if (!toDelete.Any(d => ReferenceEquals(d, r)))
                throw new InvalidOperationException(Loc.T(
                    "削除するルールの特定に失敗したため、保存を中止しました（Outlook には何も保存していません）。",
                    "Could not identify the rule to delete, so saving was cancelled (nothing was saved to Outlook)."));
            progress?.Report(Loc.T($"削除しています: {(string)((dynamic)r).Name}", $"Deleting: {(string)((dynamic)r).Name}"));
            rules.Remove(last); // never touch r after this
        }

        // 6. Final check
        for (int k = 0; k < plan.FinalOrder.Count; k++)
            if (!ReferenceEquals((object)rules[k + 1], byId[plan.FinalOrder[k].Id]))
                throw new InvalidOperationException(Loc.T(
                    $"並び順の反映結果が想定と違うため、保存を中止しました（{k + 1} 番目）。Outlook には何も保存していません。",
                    $"The resulting order is not as expected (position {k + 1}), so saving was cancelled. Nothing was saved to Outlook."));

        // 7. Save
        if (dryRun) return;
        progress?.Report(Loc.T("Outlook に保存しています…", "Saving to Outlook…"));
        rules.Save(false);
    });

    private dynamic GetStore(string storeId)
    {
        foreach (dynamic s in Session.Stores)
            if ((string)s.StoreID == storeId) return s;
        throw new InvalidOperationException(Loc.T(
            "アカウント（ストア）が見つかりません。Outlook の構成が変わった可能性があります。",
            "The account (store) was not found. Outlook's configuration may have changed."));
    }

    private dynamic GetRules(string storeId)
    {
        dynamic store = GetStore(storeId);
        try
        {
            return store.GetRules();
        }
        catch (COMException ex)
        {
            string name = (string)store.DisplayName;
            throw new InvalidOperationException(Loc.T($"「{name}」の仕分けルールを取得できません: {ex.Message}", $"Cannot get the rules of \"{name}\": {ex.Message}"), ex);
        }
    }

    private object GetFolder(FolderRef folder) =>
        (object?)Session.GetFolderFromID(folder.EntryId, folder.StoreId)
        ?? throw new InvalidOperationException(Loc.T($"フォルダー「{folder.DisplayPath}」が見つかりません。", $"Folder \"{folder.DisplayPath}\" was not found."));

    private static void Step(string ruleName, Action action)
    {
        try { action(); }
        catch (Exception ex) when (ex is not RulesChangedException)
        {
            throw new InvalidOperationException(Loc.T(
                $"ルール「{ruleName}」の反映に失敗したため、保存を中止しました（Outlook には何も保存していません）。{Environment.NewLine}{ex.Message}",
                $"Applying rule \"{ruleName}\" failed, so saving was cancelled (nothing was saved to Outlook).{Environment.NewLine}{ex.Message}"), ex);
        }
    }

    /// <summary>If Outlook was closed or restarted, drops the connection so that the next call reconnects.</summary>
    private TResult Guard<TResult>(Func<TResult> func)
    {
        try { return func(); }
        catch (COMException ex) when (IsDisconnected(ex))
        {
            Reset();
            throw new InvalidOperationException(Loc.T(
                "Outlook との接続が切れました（Outlook が終了・再起動した可能性があります）。もう一度実行してください。",
                "The connection to Outlook was lost (Outlook may have been closed or restarted). Please try again."), ex);
        }
    }

    private void Guard(Action action) => Guard(() => { action(); return 0; });

    private static bool IsDisconnected(COMException ex) =>
        ex.HResult is unchecked((int)0x800706BA) or unchecked((int)0x800706BE) or unchecked((int)0x80010108);

    private void Reset()
    {
        _session = null;
        _app = null;
    }

    public void Dispose() => Reset();
}
