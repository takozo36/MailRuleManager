using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleSignatureTests
{
    [Fact]
    public void SameContentMatches_EvenWithDifferentDisplayNamesAndOrder()
    {
        var a = FromRule(1, "original", ["a@example.com", "b@example.com"]);
        var b = a with
        {
            Name = "duplicate",
            Conditions = [new RuleCondition(ConditionType.From, [RuleValue.Address("B", "B@example.com"), RuleValue.Address(null, "a@example.com")])],
            Actions = a.Actions.Reverse().ToList(),
        };
        Assert.Equal(RuleSignature.Of(a), RuleSignature.Of(b));
    }

    [Fact]
    public void DifferentSenderOrFolderDoesNotMatch()
    {
        var a = FromRule(1, "original", ["a@example.com"], folder: "A");
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(2, "x", ["c@example.com"], folder: "A")));
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(3, "y", ["a@example.com"], folder: "B")));
    }

    [Fact]
    public void StopProcessingAndExceptionsAreCompared()
    {
        var a = FromRule(1, "original", ["a@example.com"]);
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(1, "original", ["a@example.com"], stop: false)));
        var withException = a with { Exceptions = [new RuleCondition(ConditionType.Subject, [RuleValue.Text("urgent")])] };
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(withException));
    }

    [Fact]
    public void SignatureDoesNotDependOnTheUiLanguage()
    {
        var rule = FromRule(1, "r", ["a@example.com"]) with
        {
            Conditions = [new RuleCondition(ConditionType.Importance, [RuleValue.Importance(2)])],
        };
        Loc.Current = UiLanguage.Japanese;
        string ja = RuleSignature.Of(rule);
        Loc.Current = UiLanguage.English;
        Assert.Equal(ja, RuleSignature.Of(rule));
    }
}
