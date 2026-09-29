using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.Core;

public sealed record Diagnostic(Severity Severity, string Message);

/// <summary>
/// Finds problems in the rule list.
/// Error: rules that Outlook itself treats as broken and that do not run (move-to folder gone, forward address unresolved, ...).
/// Warning: rules that run but not as intended (never reached because an earlier rule stops processing, ...).
/// Info: things worth knowing (another rule has exactly the same conditions, ...).
/// </summary>
public static class RuleDiagnostics
{
    private static readonly IReadOnlyList<Diagnostic> None = Array.Empty<Diagnostic>();

    /// <param name="entries">Rules in their current order.</param>
    /// <param name="fileExists">Checks whether a sound file exists (replaced in tests). No check when null.</param>
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
            list.Add(new(Severity.Error, T("ルールを読み取れませんでした: ", "The rule could not be read: ") + err));
            return;
        }

        var actions = e.Actions;
        if (actions.Count == 0)
            list.Add(new(Severity.Error, T("処理が 1 つも設定されていません。", "The rule has no actions.")));

        foreach (var a in actions)
        {
            if (a.Type == ActionType.MoveToFolder && a.FolderMissing)
                list.Add(new(Severity.Error, T(
                    "移動先のフォルダーが見つかりません（フォルダーが削除・移動された可能性があります）。「移動先を変更」で指定し直すと直ります。",
                    "The move-to folder was not found (it may have been deleted or moved). Use \"Change folder\" to choose it again.")));
            else if (a.Type == ActionType.CopyToFolder && a.FolderMissing)
                list.Add(new(Severity.Error, T(
                    "コピー先のフォルダーが見つかりません（フォルダーが削除・移動された可能性があります）。Outlook の画面で指定し直してください。",
                    "The copy-to folder was not found (it may have been deleted or moved). Choose it again in Outlook.")));

            foreach (var v in a.Values.Where(v => !v.Resolved))
                list.Add(new(Severity.Error, T(
                    $"{RuleText.Label(a.Type)}の宛先「{v.Display}」をアドレス帳で解決できません。",
                    $"{RuleText.Label(a.Type)}: the recipient \"{v.Display}\" cannot be resolved in the address book.")));

            if (a.Type == ActionType.PlaySound && fileExists is not null)
                foreach (var v in a.Values.Where(v => v.Display.Length > 0 && !fileExists(v.Display)))
                    list.Add(new(Severity.Warning, T($"サウンドファイルが見つかりません: {v.Display}", $"Sound file not found: {v.Display}")));
        }

        foreach (var c in e.Conditions.Concat(e.Exceptions))
            foreach (var v in c.Values.Where(v => !v.Resolved))
                list.Add(new(Severity.Warning, T(
                    $"{RuleText.Label(c.Type)}の「{v.Display}」をアドレス帳で解決できません。",
                    $"{RuleText.Label(c.Type)}: \"{v.Display}\" cannot be resolved in the address book.")));

        if (e.Conditions.Any(c => c.Type == ConditionType.Unknown) || e.Actions.Any(a => a.Type == ActionType.Unknown))
            list.Add(new(Severity.Warning, T(
                "種類を判別できない条件または処理が含まれています。Outlook の画面で内容を確認してください。",
                "The rule contains a condition or action of an unknown type. Check it in Outlook.")));

        if (!e.CanDuplicate)
        {
            var parts = RuleCapabilities.UncopyableParts(e.Source).Distinct().ToList();
            if (parts.Count > 0)
                list.Add(new(Severity.Info, T(
                    "Outlook の画面でしか設定できない項目を含むため、このアプリでは複製できません（" + string.Join("、", parts) + "）。",
                    "This rule cannot be duplicated by this app because it contains items that only Outlook can set (" + string.Join(", ", parts) + ").")));
        }
    }

    /// <summary>Checks the relation to earlier rules (unreachable rules, rules with the same conditions).</summary>
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
                list.Add(new(Severity.Warning, T(
                    $"先に実行される {i + 1} 番「{a.Name}」が同じメールを処理して「処理を中止」するため、このルールは実行されません。",
                    $"This rule never runs: rule #{i + 1} \"{a.Name}\" runs first, handles the same messages and stops processing.")));
                shadowReported = true;
            }
            else if (!duplicateReported && !aStops && b.Conditions.Count > 0 && signatures[i] == signatures[index])
            {
                list.Add(new(Severity.Info, T(
                    $"{i + 1} 番「{a.Name}」と条件がまったく同じです。",
                    $"Rule #{i + 1} \"{a.Name}\" has exactly the same conditions.")));
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
    /// When an earlier rule A has a single "matches any of these" condition (e.g. From) and stops processing,
    /// collects the values of B that are also in A (messages from them never reach B).
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
        if (overlap.Count == 0 || overlap.Count == bCond.Values.Count) return; // all of them: already reported via Covers
        notes.Add(T(
            $"{RuleText.Label(bCond.Type)}「{string.Join("、", overlap)}」は先の {aIndex + 1} 番「{a.Name}」で処理が中止されるため、このルールには届きません。",
            $"{RuleText.Label(bCond.Type)} \"{string.Join(", ", overlap)}\": these messages never reach this rule because rule #{aIndex + 1} \"{a.Name}\" stops processing first."));
    }

    /// <summary>
    /// True if every message that matches B also matches A (a message that does not match A cannot match B).
    /// Returns false when A has exceptions or conditions of unknown meaning.
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
            // If each word of B contains one of A's words, any subject matching B also matches A
            return bKeys.Count > 0 && bKeys.All(bk => aKeys.Any(ak => ak.Length > 0 && bk.Contains(ak, StringComparison.Ordinal)));
        return aKeys.OrderBy(k => k).SequenceEqual(bKeys.OrderBy(k => k));
    }

    private static string Signature(IEnumerable<RuleCondition> conditions) =>
        string.Join("|", conditions
            .Where(c => !IsScopeOnly(c.Type))
            .OrderBy(c => c.Type)
            .Select(c => $"{(int)c.Type}:{string.Join(",", c.Values.Select(v => v.Key).OrderBy(k => k))}"));

    /// <summary>Conditions that only decide where the rule runs, not which messages match.</summary>
    private static bool IsScopeOnly(ConditionType t) => t == ConditionType.LocalMachineOnly;

    private static bool IsAddressList(ConditionType t) => t is ConditionType.From or ConditionType.SentTo;

    private static bool IsWordList(ConditionType t) => t is ConditionType.Subject or ConditionType.Body
        or ConditionType.BodyOrSubject or ConditionType.MessageHeader
        or ConditionType.RecipientAddress or ConditionType.SenderAddress;
}
