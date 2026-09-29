namespace OutlookRuleManager.Core;

/// <summary>
/// Turns the content of a rule (conditions, exceptions, actions) into a string for comparison.
/// Used to read a duplicated rule back before saving and check that it matches the original.
/// Order and display names are ignored; types and comparison keys (addresses, words, folder EntryIDs) are compared.
/// The string is language-neutral (it does not depend on the UI language).
/// </summary>
public static class RuleSignature
{
    public static string Of(RuleData rule) => Of(rule.Conditions, rule.Exceptions, rule.Actions);

    public static string Of(IEnumerable<RuleCondition> conditions, IEnumerable<RuleCondition> exceptions, IEnumerable<RuleAction> actions) =>
        $"conditions[{Conditions(conditions)}] exceptions[{Conditions(exceptions)}] actions[{Actions(actions)}]";

    private static string Conditions(IEnumerable<RuleCondition> conditions) =>
        string.Join(" ", conditions
            .OrderBy(c => c.Type)
            .Select(c => $"{c.Type}({Keys(c.Values)})"));

    private static string Actions(IEnumerable<RuleAction> actions) =>
        string.Join(" ", actions
            .OrderBy(a => a.Type)
            .Select(a => a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder
                ? $"{a.Type}({(a.FolderMissing ? "no-folder" : a.Folder?.EntryId.ToUpperInvariant())})"
                : $"{a.Type}({Keys(a.Values)})"));

    private static string Keys(IEnumerable<RuleValue> values) =>
        string.Join(",", values.Select(v => v.Key).OrderBy(k => k, StringComparer.Ordinal));
}
