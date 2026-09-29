using OutlookRuleManager.Core;

// Loc.Current is global state, so tests must not run in parallel
[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace OutlookRuleManager.Core.Tests;

/// <summary>
/// Base class for tests that check Japanese messages. Sets the UI language to Japanese regardless of the machine's
/// Windows language (tests for English texts are in <see cref="LocalizationTests"/>).
/// </summary>
public abstract class JapaneseTestBase
{
    protected JapaneseTestBase() => Loc.Current = UiLanguage.Japanese;
}

/// <summary>Helpers to build test rules concisely.</summary>
internal static class TestRules
{
    public static FolderRef Folder(string path) => new($@"\\me@example.com\Inbox\{path}", "EID-" + path, "SID");

    /// <summary>A rule "from any of these senders → move to folder (→ stop processing)".</summary>
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

    /// <summary>count simple rules that differ only in name (R1, R2, ...).</summary>
    public static List<RuleData> Simple(int count) =>
        Enumerable.Range(1, count).Select(i => FromRule(i, $"R{i}", [$"user{i}@example.com"])).ToList();

    public static string Names(IEnumerable<RuleEntry> entries) => string.Join(",", entries.Select(e => e.Name));

    public static IReadOnlyList<Diagnostic> DiagnosticsOf(IReadOnlyList<RuleEntry> entries, string name) =>
        RuleDiagnostics.Analyze(entries)[entries.Single(e => e.Name == name).Id];
}
