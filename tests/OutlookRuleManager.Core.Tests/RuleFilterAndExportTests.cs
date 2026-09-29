using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleFilterAndExportTests
{
    private static readonly IReadOnlyList<Diagnostic> NoDiag = Array.Empty<Diagnostic>();

    [Fact]
    public void 検索は名前_差出人_移動先を対象に空白区切りでAND()
    {
        var e = RuleEntry.FromSource(FromRule(1, "顧客　サンプル商事", ["yamada@sample.example"], folder: @"取引先\サンプル東京"));
        Assert.True(RuleFilter.Matches(e, NoDiag, "サンプル", RuleFilterMode.All));
        Assert.True(RuleFilter.Matches(e, NoDiag, "YAMADA 東京", RuleFilterMode.All));
        Assert.False(RuleFilter.Matches(e, NoDiag, "yamada 大阪", RuleFilterMode.All));
    }

    [Fact]
    public void 全角英数字と全角空白でも検索できる()
    {
        var e = RuleEntry.FromSource(FromRule(1, "社内", ["ito@example.com"]));
        Assert.True(RuleFilter.Matches(e, NoDiag, "ＩＴＯ　社内", RuleFilterMode.All));
    }

    [Fact]
    public void 小さいカナ_ひらがな_半角カナの違いを無視して検索できる()
    {
        var e = RuleEntry.FromSource(FromRule(1, "お店", ["info@example.com"], folder: @"お店\キャンプ用品"));
        Assert.True(RuleFilter.Matches(e, NoDiag, "キヤンプ", RuleFilterMode.All));
        Assert.True(RuleFilter.Matches(e, NoDiag, "きやんぷ", RuleFilterMode.All));
        Assert.True(RuleFilter.Matches(e, NoDiag, "ｷﾔﾝﾌﾟ", RuleFilterMode.All));
    }

    [Fact]
    public void 表示モードで絞り込む()
    {
        var e = RuleEntry.FromSource(FromRule(1, "x", ["a@example.com"], enabled: false));
        var error = new[] { new Diagnostic(Severity.Error, "e") };
        var info = new[] { new Diagnostic(Severity.Info, "i") };
        Assert.True(RuleFilter.Matches(e, NoDiag, "", RuleFilterMode.Disabled));
        Assert.True(RuleFilter.Matches(e, error, "", RuleFilterMode.ErrorsOnly));
        Assert.False(RuleFilter.Matches(e, info, "", RuleFilterMode.Problems));
        Assert.False(RuleFilter.Matches(e, NoDiag, "", RuleFilterMode.Changed));
    }

    [Fact]
    public void フォルダー検索()
    {
        var tree = new FolderNode("受信トレイ", Folder(""), [
            new FolderNode("取引先", Folder("取引先"), [new FolderNode("サンプル商事", Folder(@"取引先\サンプル商事"), [])]),
            new FolderNode("社内", Folder("社内"), []),
        ]);
        var hits = FolderNode.Search([tree], "取引 サンプル").Select(n => n.Name).ToList();
        Assert.Equal(["サンプル商事"], hits);
    }

    [Fact]
    public void CSVは項目をエスケープし_実行順は全体の位置()
    {
        var entries = new[]
        {
            RuleEntry.FromSource(FromRule(1, "名前,カンマ", ["a@example.com"])),
            RuleEntry.FromSource(FromRule(2, "引用\"符", ["b@example.com"])),
        };
        var diags = RuleDiagnostics.Analyze(entries);
        var sw = new StringWriter();
        RuleCsvExporter.Write(sw, entries, diags, new HashSet<string> { "R2" });
        var lines = sw.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length); // 見出し + R2 のみ
        Assert.StartsWith("2,有効,\"引用\"\"符\"", lines[1]);

        sw = new StringWriter();
        RuleCsvExporter.Write(sw, entries, diags);
        Assert.Contains("\"名前,カンマ\"", sw.ToString());
    }

    [Fact]
    public void アドレスの表示は名前とアドレスが同じなら片方だけ()
    {
        Assert.Equal("a@example.com", RuleValue.Address("a@example.com", "a@example.com").Display);
        Assert.Equal("山田 <A@example.com>", RuleValue.Address("山田", "A@example.com").Display);
        Assert.Equal("a@example.com", RuleValue.Address("山田", "A@example.com").Key);
    }

    [Fact]
    public void 表示用パスはストア名を除く()
    {
        Assert.Equal(@"受信トレイ\社内", new FolderRef(@"\\me@example.com\受信トレイ\社内", "", "").DisplayPath);
    }
}
