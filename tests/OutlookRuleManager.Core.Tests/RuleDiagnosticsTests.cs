using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleDiagnosticsTests
{
    private static List<RuleEntry> Entries(params RuleData[] rules) => rules.Select(RuleEntry.FromSource).ToList();

    [Fact]
    public void 移動先フォルダーが消えたルールはエラー()
    {
        var entries = Entries(FromRule(1, "壊れ", ["a@example.com"], folder: null));
        var d = DiagnosticsOf(entries, "壊れ");
        Assert.Contains(d, x => x.Severity == Severity.Error && x.Message.Contains("移動先のフォルダーが見つかりません"));
    }

    [Fact]
    public void 移動先を指定し直すとエラーが消える()
    {
        var ed = new RuleListEditor([FromRule(1, "壊れ", ["a@example.com"], folder: null)]);
        ed.SetMoveFolder(["R1"], Folder("新"));
        Assert.DoesNotContain(DiagnosticsOf(ed.Entries, "壊れ"), x => x.Severity == Severity.Error);
    }

    [Fact]
    public void 処理のないルールはエラー()
    {
        var rule = new RuleData { Index = 1, Name = "空", Conditions = [new RuleCondition(ConditionType.HasAttachment)] };
        Assert.Contains(DiagnosticsOf(Entries(rule), "空"), x => x.Severity == Severity.Error);
    }

    [Fact]
    public void 転送先を解決できないルールはエラー()
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
    public void 読み取りに失敗したルールはエラー()
    {
        var rule = new RuleData { Index = 1, Name = "読めない", ReadError = "RPC エラー" };
        Assert.Contains(DiagnosticsOf(Entries(rule), "読めない"), x => x.Severity == Severity.Error && x.Message.Contains("RPC エラー"));
    }

    [Fact]
    public void サウンドファイルが無ければ警告()
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
    public void 先のルールが同じ差出人で処理を中止するなら後のルールは届かない()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com", "b@example.com"]),
            FromRule(2, "後", ["A@Example.com"]));
        Assert.Contains(DiagnosticsOf(entries, "後"), x => x.Severity == Severity.Warning && x.Message.Contains("実行されません"));
        Assert.Empty(DiagnosticsOf(entries, "先"));
    }

    [Fact]
    public void 先のルールが処理を中止しないなら届く_条件が同じことだけ知らせる()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com"], stop: false),
            FromRule(2, "後", ["a@example.com"]));
        var d = DiagnosticsOf(entries, "後");
        Assert.DoesNotContain(d, x => x.Severity >= Severity.Warning);
        Assert.Contains(d, x => x.Severity == Severity.Info && x.Message.Contains("条件がまったく同じ"));
    }

    [Fact]
    public void 無効なルールは届かない判定に使わない()
    {
        var entries = Entries(
            FromRule(1, "先", ["a@example.com"], enabled: false),
            FromRule(2, "後", ["a@example.com"]));
        Assert.Empty(DiagnosticsOf(entries, "後"));
    }

    [Fact]
    public void 差出人の一部だけが先のルールと重なる場合はその差出人だけ警告()
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
    public void 件名の語_先のルールの語を含む語なら届かない()
    {
        var entries = Entries(
            SubjectRule(1, "先", ["請求"]),
            SubjectRule(2, "後", ["請求書", "ご請求"]),
            SubjectRule(3, "別", ["見積"]));
        Assert.Contains(DiagnosticsOf(entries, "後"), x => x.Message.Contains("実行されません"));
        Assert.DoesNotContain(DiagnosticsOf(entries, "別"), x => x.Message.Contains("実行されません"));
    }

    [Fact]
    public void 先のルールに例外があれば判断しない()
    {
        var first = FromRule(1, "先", ["a@example.com"]) with
        {
            Exceptions = [new RuleCondition(ConditionType.Subject, [RuleValue.Text("至急")])],
        };
        var entries = Entries(first, FromRule(2, "後", ["a@example.com"]));
        Assert.DoesNotContain(DiagnosticsOf(entries, "後"), x => x.Severity >= Severity.Warning);
    }

    [Fact]
    public void このコンピューターのみ条件は比較で無視する()
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
    public void 並び順を入れ替えると届かない判定も変わる()
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
