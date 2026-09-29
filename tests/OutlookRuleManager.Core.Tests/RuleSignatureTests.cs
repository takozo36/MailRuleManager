using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleSignatureTests
{
    [Fact]
    public void 表示名や並び順が違っても中身が同じなら一致()
    {
        var a = FromRule(1, "元", ["a@example.com", "b@example.com"]);
        var b = a with
        {
            Name = "複製",
            Conditions = [new RuleCondition(ConditionType.From, [RuleValue.Address("Bさん", "B@example.com"), RuleValue.Address(null, "a@example.com")])],
            Actions = a.Actions.Reverse().ToList(),
        };
        Assert.Equal(RuleSignature.Of(a), RuleSignature.Of(b));
    }

    [Fact]
    public void 差出人や移動先が違えば不一致()
    {
        var a = FromRule(1, "元", ["a@example.com"], folder: "A");
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(2, "x", ["c@example.com"], folder: "A")));
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(3, "y", ["a@example.com"], folder: "B")));
    }

    [Fact]
    public void 処理を中止の有無や例外の違いも区別する()
    {
        var a = FromRule(1, "元", ["a@example.com"]);
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(FromRule(1, "元", ["a@example.com"], stop: false)));
        var withException = a with { Exceptions = [new RuleCondition(ConditionType.Subject, [RuleValue.Text("至急")])] };
        Assert.NotEqual(RuleSignature.Of(a), RuleSignature.Of(withException));
    }
}
