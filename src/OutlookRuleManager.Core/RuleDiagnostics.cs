namespace OutlookRuleManager.Core;

public sealed record Diagnostic(Severity Severity, string Message);

/// <summary>
/// ルール一覧の問題点を調べる。
/// エラー: Outlook でも「エラー」扱いになり動かないもの（移動先フォルダーが消えた、転送先を解決できない等）。
/// 警告: 動くが意図どおりにならないもの（先のルールで処理が中止されて届かない等）。
/// 情報: 知っておくとよいこと（同じ条件のルールが他にもある等）。
/// </summary>
public static class RuleDiagnostics
{
    private static readonly IReadOnlyList<Diagnostic> None = Array.Empty<Diagnostic>();

    /// <param name="entries">現在の並び順どおりのルール一覧</param>
    /// <param name="fileExists">サウンドファイルの存在確認（テストでは差し替える）。null なら確認しない。</param>
    public static IReadOnlyDictionary<string, IReadOnlyList<Diagnostic>> Analyze(
        IReadOnlyList<RuleEntry> entries, Func<string, bool>? fileExists = null)
    {
        var result = new Dictionary<string, IReadOnlyList<Diagnostic>>(entries.Count);
        var signatures = entries.Select(e => Signature(e.Conditions) + "#" + Signature(e.Exceptions)).ToList();
        for (int i = 0; i < entries.Count; i++)
        {
            var list = new List<Diagnostic>();
            CheckRule(entries[i], fileExists, list);
            CheckAgainstEarlier(entries, signatures, i, list);
            result[entries[i].Id] = list.Count == 0 ? None : list.OrderByDescending(d => d.Severity).ToList();
        }
        return result;
    }

    public static Severity? Worst(IReadOnlyList<Diagnostic> diagnostics) =>
        diagnostics.Count == 0 ? null : diagnostics.Max(d => d.Severity);

    private static void CheckRule(RuleEntry e, Func<string, bool>? fileExists, List<Diagnostic> list)
    {
        if (e.Source.ReadError is { } err)
        {
            list.Add(new(Severity.Error, "ルールを読み取れませんでした: " + err));
            return;
        }

        var actions = e.Actions;
        if (actions.Count == 0)
            list.Add(new(Severity.Error, "処理が 1 つも設定されていません。"));

        foreach (var a in actions)
        {
            if (a.Type == ActionType.MoveToFolder && a.FolderMissing)
                list.Add(new(Severity.Error, "移動先のフォルダーが見つかりません（フォルダーが削除・移動された可能性があります）。「移動先を変更」で指定し直すと直ります。"));
            else if (a.Type == ActionType.CopyToFolder && a.FolderMissing)
                list.Add(new(Severity.Error, "コピー先のフォルダーが見つかりません（フォルダーが削除・移動された可能性があります）。Outlook の画面で指定し直してください。"));

            foreach (var v in a.Values.Where(v => !v.Resolved))
                list.Add(new(Severity.Error, $"{RuleText.Label(a.Type)}の宛先「{v.Display}」をアドレス帳で解決できません。"));

            if (a.Type == ActionType.PlaySound && fileExists is not null)
                foreach (var v in a.Values.Where(v => v.Display.Length > 0 && !fileExists(v.Display)))
                    list.Add(new(Severity.Warning, $"サウンドファイルが見つかりません: {v.Display}"));
        }

        foreach (var c in e.Conditions.Concat(e.Exceptions))
            foreach (var v in c.Values.Where(v => !v.Resolved))
                list.Add(new(Severity.Warning, $"{RuleText.Label(c.Type)}の「{v.Display}」をアドレス帳で解決できません。"));

        if (e.Conditions.Any(c => c.Type == ConditionType.Unknown) || e.Actions.Any(a => a.Type == ActionType.Unknown))
            list.Add(new(Severity.Warning, "種類を判別できない条件または処理が含まれています。Outlook の画面で内容を確認してください。"));

        if (!e.CanDuplicate)
        {
            var parts = RuleCapabilities.UncopyableParts(e.Source).Distinct().ToList();
            if (parts.Count > 0)
                list.Add(new(Severity.Info, "Outlook の画面でしか設定できない項目を含むため、このアプリでは複製できません（" + string.Join("、", parts) + "）。"));
        }
    }

    /// <summary>先に実行されるルールとの関係を調べる（届かないルール・同じ条件のルール）。</summary>
    private static void CheckAgainstEarlier(IReadOnlyList<RuleEntry> entries, IReadOnlyList<string> signatures, int index, List<Diagnostic> list)
    {
        var b = entries[index];
        if (!b.Enabled || b.Source.ReadError is not null) return;

        bool shadowReported = false, duplicateReported = false;
        var partialNotes = new List<string>();
        for (int i = 0; i < index; i++)
        {
            var a = entries[i];
            if (!a.Enabled || a.Kind != b.Kind || a.Source.ReadError is not null) continue;
            bool aStops = a.Actions.Any(x => x.Type == ActionType.Stop);

            if (!shadowReported && aStops && Covers(a, b))
            {
                list.Add(new(Severity.Warning,
                    $"先に実行される {i + 1} 番「{a.Name}」が同じメールを処理して「処理を中止」するため、このルールは実行されません。"));
                shadowReported = true;
            }
            else if (!duplicateReported && !aStops && b.Conditions.Count > 0 && signatures[i] == signatures[index])
            {
                list.Add(new(Severity.Info, $"{i + 1} 番「{a.Name}」と条件がまったく同じです。"));
                duplicateReported = true;
            }

            if (!shadowReported && aStops)
                CollectPartialOverlap(a, i, b, partialNotes);
        }

        if (!shadowReported)
            foreach (var note in partialNotes)
                list.Add(new(Severity.Warning, note));
    }

    /// <summary>
    /// 差出人など「どれか 1 つに一致」する条件だけのルール A が先で処理を中止する場合、
    /// B の値のうち A に含まれるもの（そのメールは B まで届かない）を拾う。
    /// </summary>
    private static void CollectPartialOverlap(RuleEntry a, int aIndex, RuleEntry b, List<string> notes)
    {
        if (a.Exceptions.Count > 0) return;
        var aConds = a.Conditions.Where(c => !IsScopeOnly(c.Type)).ToList();
        if (aConds.Count != 1 || !IsAddressList(aConds[0].Type)) return;
        var bCond = b.Conditions.FirstOrDefault(c => c.Type == aConds[0].Type);
        if (bCond is null) return;

        var aKeys = aConds[0].Values.Select(v => v.Key).ToHashSet();
        var overlap = bCond.Values.Where(v => aKeys.Contains(v.Key)).Select(v => v.Display).ToList();
        if (overlap.Count == 0 || overlap.Count == bCond.Values.Count) return; // 全部なら Covers 側で報告済み
        notes.Add($"{RuleText.Label(bCond.Type)}「{string.Join("、", overlap)}」は先の {aIndex + 1} 番「{a.Name}」で処理が中止されるため、このルールには届きません。");
    }

    /// <summary>
    /// A の条件に一致しないメールは B にも一致しない（B に一致するメールは必ず A にも一致する）なら true。
    /// A に例外や判別できない条件があるときは判断しない。
    /// </summary>
    internal static bool Covers(RuleEntry a, RuleEntry b)
    {
        if (a.Exceptions.Count > 0) return false;
        if (a.Conditions.Any(c => c.Type is ConditionType.Unknown or ConditionType.OtherMachine)) return false;
        foreach (var ca in a.Conditions.Where(c => !IsScopeOnly(c.Type)))
        {
            var cb = b.Conditions.FirstOrDefault(c => c.Type == ca.Type);
            if (cb is null || !ConditionCovers(ca, cb)) return false;
        }
        return true;
    }

    private static bool ConditionCovers(RuleCondition a, RuleCondition b)
    {
        var aKeys = a.Values.Select(v => v.Key).ToList();
        var bKeys = b.Values.Select(v => v.Key).ToList();
        if (IsAddressList(a.Type) || a.Type == ConditionType.Category)
            return bKeys.Count > 0 && bKeys.All(aKeys.Contains);
        if (IsWordList(a.Type))
            // B の各語が A のいずれかの語を含んでいれば、B に一致する件名は A にも一致する
            return bKeys.Count > 0 && bKeys.All(bk => aKeys.Any(ak => ak.Length > 0 && bk.Contains(ak, StringComparison.Ordinal)));
        return aKeys.OrderBy(k => k).SequenceEqual(bKeys.OrderBy(k => k));
    }


    private static string Signature(IEnumerable<RuleCondition> conditions) =>
        string.Join("|", conditions
            .Where(c => !IsScopeOnly(c.Type))
            .OrderBy(c => c.Type)
            .Select(c => $"{(int)c.Type}:{string.Join(",", c.Values.Select(v => v.Key).OrderBy(k => k))}"));

    /// <summary>メールの中身ではなく「どこで実行するか」だけを決める条件。</summary>
    private static bool IsScopeOnly(ConditionType t) => t == ConditionType.LocalMachineOnly;

    private static bool IsAddressList(ConditionType t) => t is ConditionType.From or ConditionType.SentTo;

    private static bool IsWordList(ConditionType t) => t is ConditionType.Subject or ConditionType.Body
        or ConditionType.BodyOrSubject or ConditionType.MessageHeader
        or ConditionType.RecipientAddress or ConditionType.SenderAddress;
}
