using System.Runtime.InteropServices;
using OutlookRuleManager.Core;
using static OutlookRuleManager.Outlook.OutlookCom;

namespace OutlookRuleManager.Outlook;

public sealed record StoreInfo(string StoreId, string DisplayName, bool IsDefault)
{
    public override string ToString() => DisplayName;
}

/// <summary>読み込み途中の通知。Rule は読めたルール 1 件、Total は全件数。</summary>
public sealed record LoadProgress(RuleData Rule, int Total);

/// <summary>保存しようとしたら Outlook 側のルールが読み込み時から変わっていた。</summary>
public sealed class RulesChangedException(string message) : Exception(message);

/// <summary>
/// Outlook（クラシック）の仕訳ルールの読み込みと保存。
/// 呼び出しは UI スレッド以外から行う想定（どれも数秒〜数十秒かかる）。
/// Outlook のオブジェクトは遅延バインディング（dynamic）で扱う（OutlookCom を参照）。
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

    /// <summary>直前の LoadRules がどちらの方法で読んだか（画面表示用）。</summary>
    public string LastLoadNote { get; private set; } = "";

    /// <summary>
    /// 全ルールを読み込む。
    /// まず受信トレイの隠しメッセージにある全ルールのまとめデータ（PR_RW_RULES_STREAM）を 1 回で取り出して解析する
    /// （数秒）。取り出せない・解析できない・Outlook 側と件数や名前が合わない場合は、条件・処理を 1 つずつ
    /// Outlook に問い合わせる方法に切り替える（1 件あたり 0.3 秒ほど）。
    /// 読めたルールから順に progress へ渡すので、画面は読み込み途中から表示できる。
    /// </summary>
    /// <param name="allowFast">false なら常に 1 件ずつ読む（検証用）。</param>
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
                LastLoadNote = "高速読み込み";
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
        LastLoadNote = reason is null ? "1 件ずつ読み込み" : $"1 件ずつ読み込み（高速読み込みできなかった理由: {reason}）";
        return list;
    });

    /// <summary>
    /// まとめデータからの読み込み。失敗したら null と理由を返す（例外は投げない）。
    /// 解析結果は、件数・各位置の名前・有効/無効を Outlook のルールと照合してから使う。
    /// </summary>
    private List<RuleData>? TryLoadFast(string storeId, dynamic rules, int count, CancellationToken cancel, out string? reason)
    {
        IReadOnlyList<StreamRule> parsed;
        try
        {
            parsed = RulesStream.Parse(ReadRulesStream(storeId));
        }
        catch (RulesStreamFormatException ex) { reason = ex.Message; return null; }
        catch (Exception ex) when (ex is not OperationCanceledException) { reason = "ルールのまとめデータを取り出せません: " + ex.Message; return null; }

        if (parsed.Count != count)
        {
            reason = $"件数が一致しません（まとめデータ {parsed.Count} 件、Outlook {count} 件）";
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
                reason = $"{i} 番目のルールの名前か有効/無効が一致しません";
                return null;
            }
            // 「このコンピューターのみ」は、このPCか別のPCかをまとめデータだけでは判定できないので Outlook に聞く
            bool onThisMachine = !p.NeedsMachineCheck || (bool)rule.Conditions.OnLocalMachine.Enabled;
            list.Add(p.ToRuleData(f => ResolveFolder(f, folders), onThisMachine, (bool)rule.IsLocalRule));
        }
        reason = null;
        return list;
    }

    /// <summary>受信トレイの隠しメッセージ IPM.RuleOrganizer から PR_RW_RULES_STREAM を取り出す。</summary>
    private byte[] ReadRulesStream(string storeId)
    {
        dynamic store = GetStore(storeId);
        dynamic inbox = store.GetDefaultFolder(6 /* olFolderInbox */);
        dynamic table = inbox.GetTable("", 1 /* olHiddenItems */);
        table.Columns.Add("MessageClass");
        table.Columns.Add("EntryID");
        string? entryId = null;
        while (!(bool)table.EndOfTable)
        {
            dynamic row = table.GetNextRow();
            if ((string)row["MessageClass"] == "IPM.RuleOrganizer") { entryId = (string)row["EntryID"]; break; }
        }
        if (entryId is null) throw new InvalidOperationException("受信トレイに IPM.RuleOrganizer がありません");
        dynamic item = Session.GetItemFromID(entryId, storeId);
        return (byte[])item.PropertyAccessor.GetProperty(RulesStreamProperty);
    }

    private const string RulesStreamProperty = "http://schemas.microsoft.com/mapi/proptag/0x68020102";

    /// <summary>まとめデータ内の EntryID を Outlook のフォルダーに変換する（同じフォルダーは 1 回だけ問い合わせる）。</summary>
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
            catch { /* 見つからない＝削除・移動されたフォルダー */ }
        }
        cache[raw.EntryId] = result;
        return result;
    }

    /// <summary>メールフォルダーのツリー（最上位は受信トレイなどストア直下のフォルダー）。</summary>
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
    /// 編集内容を Outlook に反映して保存する。
    /// Rules.Save() を呼ぶまでは Outlook 側に何も保存されないので、途中で失敗したら保存せずに例外を投げる。
    /// </summary>
    /// <param name="dryRun">true なら最後の Rules.Save() だけを行わない（手順が通るかの確認用。Outlook には何も残らない）。</param>
    public void Apply(string storeId, ChangePlan plan, IProgress<string>? progress, bool dryRun = false) => Guard(() =>
    {
        var problems = plan.Validate();
        if (problems.Count > 0) throw new InvalidOperationException(string.Join(Environment.NewLine, problems));

        progress?.Report("Outlook のルールを確認しています…");
        var rules = GetRules(storeId);

        // 1. 読み込み時から Outlook 側で変更されていないか（位置と名前で照合）
        int count = (int)rules.Count;
        if (count != plan.Original.Count)
            throw new RulesChangedException($"Outlook 側のルールの数が変わっています（読み込み時 {plan.Original.Count} 件 → 現在 {count} 件）。");
        var originals = new object[count + 1];
        for (int i = 1; i <= count; i++)
        {
            originals[i] = rules[i];
            string name = (string)((dynamic)originals[i]).Name;
            if (name != plan.Original[i - 1].Name)
                throw new RulesChangedException($"Outlook 側の {i} 番目のルールが読み込み時と違います（「{plan.Original[i - 1].Name}」→「{name}」）。");
        }

        // 2. 既存ルールの名前・有効/無効・移動先
        progress?.Report("変更を反映しています…");
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

        // 3. 複製（削除より先に行う。削除したルールのオブジェクトに触ると Outlook が異常終了するため）
        foreach (var e in plan.Created)
        {
            progress?.Report($"複製しています: {e.Name}");
            Step(e.Name, () =>
            {
                dynamic created = rules.Create(e.Name, (int)e.Kind);
                byId[e.Id] = (object)created;
                OutlookRuleWriter.Copy(originals[e.Source.Index], (object)created);
                if (e.MoveFolderOverride is { } folder) SetProperty((object)created.Actions.MoveToFolder, "Folder", GetFolder(folder));
                created.Enabled = e.Enabled;

                // 保存前に読み戻して、元のルール（移動先の変更を反映したもの）と中身が同じか確かめる
                var copied = OutlookRuleReader.Read((object)created, 0);
                string expected = RuleSignature.Of(e.Conditions, e.Exceptions, e.Actions);
                string actual = copied.ReadError ?? RuleSignature.Of(copied);
                if (actual != expected)
                    throw new InvalidOperationException($"複製した内容が元のルールと一致しません。{Environment.NewLine}元: {expected}{Environment.NewLine}複製: {actual}");
            });
        }

        // 4. 並び順。残すルールを 1 番から順に置くと、削除するルールは末尾に押し出される
        progress?.Report("並び順を反映しています…");
        for (int k = 0; k < plan.FinalOrder.Count; k++)
        {
            dynamic r = byId[plan.FinalOrder[k].Id];
            if ((int)r.ExecutionOrder != k + 1) r.ExecutionOrder = k + 1;
        }

        // 5. 削除。末尾のルールが削除対象であることをオブジェクトの同一性で確かめてから消す
        var toDelete = plan.Deleted.Select(d => originals[d.Index]).ToList();
        while ((int)rules.Count > plan.FinalOrder.Count)
        {
            int last = (int)rules.Count;
            object r = rules[last];
            if (!toDelete.Any(d => ReferenceEquals(d, r)))
                throw new InvalidOperationException("削除するルールの特定に失敗したため、保存を中止しました（Outlook には何も保存していません）。");
            progress?.Report($"削除しています: {(string)((dynamic)r).Name}");
            rules.Remove(last); // 以後 r には触らない
        }

        // 6. 最終確認
        for (int k = 0; k < plan.FinalOrder.Count; k++)
            if (!ReferenceEquals((object)rules[k + 1], byId[plan.FinalOrder[k].Id]))
                throw new InvalidOperationException($"並び順の反映結果が想定と違うため、保存を中止しました（{k + 1} 番目）。Outlook には何も保存していません。");

        // 7. 保存
        if (dryRun) return;
        progress?.Report("Outlook に保存しています…");
        rules.Save(false);
    });

    private dynamic GetStore(string storeId)
    {
        foreach (dynamic s in Session.Stores)
            if ((string)s.StoreID == storeId) return s;
        throw new InvalidOperationException("アカウント（ストア）が見つかりません。Outlook の構成が変わった可能性があります。");
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
            throw new InvalidOperationException($"「{(string)store.DisplayName}」の仕訳ルールを取得できません: {ex.Message}", ex);
        }
    }

    private object GetFolder(FolderRef folder) =>
        (object?)Session.GetFolderFromID(folder.EntryId, folder.StoreId)
        ?? throw new InvalidOperationException($"フォルダー「{folder.DisplayPath}」が見つかりません。");

    private static void Step(string ruleName, Action action)
    {
        try { action(); }
        catch (Exception ex) when (ex is not RulesChangedException)
        {
            throw new InvalidOperationException($"ルール「{ruleName}」の反映に失敗したため、保存を中止しました（Outlook には何も保存していません）。{Environment.NewLine}{ex.Message}", ex);
        }
    }

    /// <summary>Outlook が終了・再起動していたら接続を捨てて、次回の呼び出しでつなぎ直す。</summary>
    private T Guard<T>(Func<T> func)
    {
        try { return func(); }
        catch (COMException ex) when (IsDisconnected(ex))
        {
            Reset();
            throw new InvalidOperationException("Outlook との接続が切れました（Outlook が終了・再起動した可能性があります）。もう一度実行してください。", ex);
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
