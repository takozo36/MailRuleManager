using System.Globalization;
using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

/// <summary>Texts shown when the UI language is English, and the language auto-detection.</summary>
public class LocalizationTests
{
    public LocalizationTests() => Loc.Current = UiLanguage.English;

    [Fact]
    public void LanguageIsDetectedFromTheWindowsCulture()
    {
        Assert.Equal(UiLanguage.Japanese, Loc.FromCulture(new CultureInfo("ja-JP")));
        Assert.Equal(UiLanguage.English, Loc.FromCulture(new CultureInfo("en-US")));
        Assert.Equal(UiLanguage.English, Loc.FromCulture(new CultureInfo("de-DE"))); // other languages fall back to English
    }

    [Fact]
    public void LabelsAndSummariesAreInEnglish()
    {
        var e = RuleEntry.FromSource(FromRule(1, "r", ["a@example.com"], folder: "Clients"));
        Assert.Equal("From: a@example.com", RuleText.Summary(e.Conditions));
        Assert.Equal("Move / Stop processing", RuleText.ActionSummary(e.Actions));
        Assert.Equal("Receive", RuleText.KindText(e));
        Assert.Equal("Enabled", RuleText.EnabledText(true));
        Assert.Equal("Warning", RuleText.SeverityText(Severity.Warning));
    }

    [Fact]
    public void ImportanceFollowsTheCurrentLanguage()
    {
        var c = new RuleCondition(ConditionType.Importance, [RuleValue.Importance(2)]);
        Assert.Equal("Importance: High", RuleText.Describe(c));
        Loc.Current = UiLanguage.Japanese;
        Assert.Equal("重要度: 高", RuleText.Describe(c));
    }

    [Fact]
    public void DiagnosticsAreInEnglish()
    {
        var entries = new[]
        {
            RuleEntry.FromSource(FromRule(1, "first", ["a@example.com"])),
            RuleEntry.FromSource(FromRule(2, "second", ["a@example.com"])),
            RuleEntry.FromSource(FromRule(3, "broken", ["b@example.com"], folder: null)),
        };
        var d = RuleDiagnostics.Analyze(entries);
        Assert.Contains(d["R2"], x => x.Message.Contains("never runs") && x.Message.Contains("#1 \"first\""));
        Assert.Contains(d["R3"], x => x.Severity == Severity.Error && x.Message.Contains("move-to folder was not found"));
    }

    [Fact]
    public void DuplicateSuffixAndPlanDescriptionAreInEnglish()
    {
        var ed = new RuleListEditor(Simple(3));
        ed.Duplicate(["R1"], out _);
        Assert.Equal("R1 (copy)", ed.Entries[1].Name);
        ed.Swap("R2", "R3");
        var lines = ed.BuildPlan().Describe();
        Assert.Contains(lines, l => l.StartsWith("Add (duplicate): \"R1 (copy)\""));
        Assert.Contains(lines, l => l.StartsWith("Reorder:"));
    }

    [Fact]
    public void CsvHeaderIsInEnglish()
    {
        var entries = new[] { RuleEntry.FromSource(FromRule(1, "r", ["a@example.com"])) };
        var sw = new StringWriter();
        RuleCsvExporter.Write(sw, entries, RuleDiagnostics.Analyze(entries));
        var lines = sw.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Order,Enabled,Name,Type,Conditions,Exceptions,Actions,Move to folder,Diagnostics", lines[0]);
        Assert.StartsWith("1,Enabled,r,Receive,", lines[1]);
    }

    [Fact]
    public void RenameErrorIsInEnglish()
    {
        var ed = new RuleListEditor(Simple(1));
        var ex = Assert.Throws<ArgumentException>(() => ed.Rename("R1", " "));
        Assert.StartsWith("The rule name is empty.", ex.Message);
    }
}
