using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleDiagnosticsTests : JapaneseTestBase
{
    private static List<RuleEntry> Entries(params RuleData[] rules) => rules.Select(RuleEntry.FromSource).ToList();

    [Fact]
    public void MissingMoveToFolderIsAnError()
    {
        var entries = Entries(FromRule(1, "壊れ", ["a@example.com"], folder: null));
        var d = DiagnosticsOf(entries, "壊れ");
        Assert.Contains(d, x => x.Severity == Severity.Error && x.Message.Contains("移動先のフォルダーが見つかりません"));
    }

    [Fact]
    public void ChoosingAFolderAgainClearsTheError()
    {
        var ed = new RuleListEditor([FromRule(1, "壊れ", ["a@example.com"], folder: null)]);
        ed.SetMoveFolder(["R1"], Folder("新"));
        Assert.DoesNotContain(DiagnosticsOf(ed.Entries, "壊れ"), x => x.Severity == Severity.Error);
    }

    [Fact]
    public void RuleWithoutActionsIsAnError()
    {
        var rule = new RuleData { Index = 1, Name = "空", Conditions = [new RuleCondition(ConditionType.HasAttachment)] };
        Assert.Contains(DiagnosticsOf(Entries(rule), "空"), x => x.Severity == Severity.Error);
    }

    [Fact]
    public void UnresolvedForwardRecipientIsAnError()
    {
        var rule = new RuleData
        {
            Index = 1,
            Name = "転送",
            Actions = [new RuleAction(ActionType.Forward, [RuleValue.Address("消えた人", "", resolved: false)])],
        };
        Assert.Contains(DiagnosticsOf(Entries(rule), "転送"), x => x.Severity == Severity.Error && x.Message.Contains("消えた人"));
    }

    [Fact]
    public void RuleThatCouldNotBeReadIsAnError()
    {
        var rule = new RuleData { Index = 1, Name = "読めない", ReadError = "RPC エラー" };
        Assert.Contains(DiagnosticsOf(Entries(rule), "読めない"), x => x.Severity == Severity.Error && x.Message.Contains("RPC エラー"));
    }

    [Fact]
    public void MissingSoundFileIsAWarning()
    {
        var rule = new RuleData
        {
            Index = 1,
            Name = "音",
            Actions = [new RuleAction(ActionType.PlaySound, [RuleValue.Text(@"C:\nothing.wav")])],
        };
        var d = RuleDiagnostics.Analyze(Entries(rule), _ => false)["R1"];
        Assert.Contains(d, x => x.Severity == Severity.Warning && x.Message.Contains("nothing.wav"));
    }

    [Fact]
    public void LaterRuleNeverRunsWhenAnEarlierRuleWithTheSameSenderStops()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com", "b@example.com"]),
            FromRule(2, "後", ["A@Example.com"]));
        Assert.Contains(DiagnosticsOf(entries, "後"), x => x.Severity == Severity.Warning && x.Message.Contains("実行されません"));
        Assert.Empty(DiagnosticsOf(entries, "先"));
    }

    [Fact]
    public void WhenTheEarlierRuleDoesNotStop_OnlyInfoAboutIdenticalConditions()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com"], stop: false),
            FromRule(2, "後", ["a@example.com"]));
        var d = DiagnosticsOf(entries, "後");
        Assert.DoesNotContain(d, x => x.Severity >= Severity.Warning);
        Assert.Contains(d, x => x.Severity == Severity.Info && x.Message.Contains("条件がまったく同じ"));
    }

    [Fact]
    public void DisabledRulesAreIgnoredForReachability()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com"], enabled: false),
            FromRule(2, "後", ["a@example.com"]));
        Assert.Empty(DiagnosticsOf(entries, "後"));
    }

    [Fact]
    public void PartialSenderOverlap_WarnsOnlyForTheOverlappingSenders()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com"]),
            FromRule(2, "後", ["a@example.com", "c@example.com"]));
        var d = DiagnosticsOf(entries, "後");
        var w = Assert.Single(d, x => x.Severity == Severity.Warning);
        Assert.Contains("a@example.com", w.Message);
        Assert.DoesNotContain("c@example.com", w.Message);
    }

    [Fact]
    public void SubjectWords_ContainingAnEarlierRulesWordAreNeverReached()
    {
        var entries = Entries(
            SubjectRule(1, "先", ["請求"]),
            SubjectRule(2, "後", ["請求書", "ご請求"]),
            SubjectRule(3, "別", ["見積"]));
        Assert.Contains(DiagnosticsOf(entries, "後"), x => x.Message.Contains("実行されません"));
        Assert.DoesNotContain(DiagnosticsOf(entries, "別"), x => x.Message.Contains("実行されません"));
    }

    [Fact]
    public void NoConclusionWhenTheEarlierRuleHasExceptions()
    {
        var first = FromRule(1, "先", ["a@example.com"]) with
        {
            Exceptions = [new RuleCondition(ConditionType.Subject, [RuleValue.Text("至急")])],
        };
        var entries = Entries(first, FromRule(2, "後", ["a@example.com"]));
        Assert.DoesNotContain(DiagnosticsOf(entries, "後"), x => x.Severity >= Severity.Warning);
    }

    [Fact]
    public void ThisComputerOnlyIsIgnoredWhenComparing()
    {
        var first = FromRule(1, "先", ["a@example.com"]) with
        {
            Conditions =
            [
                new RuleCondition(ConditionType.From, [RuleValue.Address(null, "a@example.com")]),
                new RuleCondition(ConditionType.LocalMachineOnly),
            ],
        };
        var entries = Entries(first, FromRule(2, "後", ["a@example.com"]));
        Assert.Contains(DiagnosticsOf(entries, "後"), x => x.Message.Contains("実行されません"));
    }

    [Fact]
    public void SwappingTheOrderChangesTheReachabilityResult()
    {
        var ed = new RuleListEditor([
            FromRule(1, "広い", ["a@example.com", "b@example.com"]),
            FromRule(2, "狭い", ["a@example.com"], folder: "特別"),
        ]);
        Assert.Contains(DiagnosticsOf(ed.Entries, "狭い"), x => x.Message.Contains("実行されません"));
        ed.Swap("R1", "R2");
        Assert.DoesNotContain(DiagnosticsOf(ed.Entries, "狭い"), x => x.Severity >= Severity.Warning);
    }
}
