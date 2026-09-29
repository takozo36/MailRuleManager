using OutlookRuleManager.Core;

namespace OutlookRuleManager.Core.Tests;

/// <summary>テスト用のルールを手短に作るための部品。</summary>
internal static class TestRules
{
    public static FolderRef Folder(string path) => new($@"\\me@example.com\受信トレイ\{path}", "EID-" + path, "SID");

    /// <summary>「差出人が from のいずれか → folder へ移動（→ 処理を中止）」のルール。</summary>
    public static RuleData FromRule(int index, string name, string[] from, string? folder = "A", bool stop = true, bool enabled = true)
    {
        var actions = new List<RuleAction> { RuleAction.ToFolder(ActionType.MoveToFolder, folder is null ? null : Folder(folder)) };
        if (stop) actions.Add(new RuleAction(ActionType.Stop));
        return new RuleData
        {
            Index = index,
            Name = name,
            Enabled = enabled,
            Conditions = [new RuleCondition(ConditionType.From, from.Select(a => RuleValue.Address(null, a)).ToList())],
            Actions = actions,
        };
    }

    public static RuleData SubjectRule(int index, string name, string[] words, bool stop = true) => new()
    {
        Index = index,
        Name = name,
        Conditions = [new RuleCondition(ConditionType.Subject, words.Select(RuleValue.Text).ToList())],
        Actions = stop
            ? [RuleAction.ToFolder(ActionType.MoveToFolder, Folder("S")), new RuleAction(ActionType.Stop)]
            : [RuleAction.ToFolder(ActionType.MoveToFolder, Folder("S"))],
    };

    /// <summary>名前だけ違う単純なルールを count 件（R1, R2, ...）。</summary>
    public static List<RuleData> Simple(int count) =>
        Enumerable.Range(1, count).Select(i => FromRule(i, $"R{i}", [$"user{i}@example.com"])).ToList();

    public static string Names(IEnumerable<RuleEntry> entries) => string.Join(",", entries.Select(e => e.Name));

    public static IReadOnlyList<Diagnostic> DiagnosticsOf(IReadOnlyList<RuleEntry> entries, string name) =>
        RuleDiagnostics.Analyze(entries)[entries.Single(e => e.Name == name).Id];
}
