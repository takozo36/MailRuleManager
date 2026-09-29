namespace OutlookRuleManager.Core;

/// <summary>
/// Conditions and actions that can be written to a new rule through the Outlook object model.
/// Only rules made up entirely of these types can be duplicated
/// (OutlookRuleWriter in the Outlook project implements copying for exactly this list).
/// </summary>
public static class RuleCapabilities
{
    public static readonly IReadOnlySet<ConditionType> CopyableConditions = new HashSet<ConditionType>
    {
        ConditionType.From,
        ConditionType.SentTo,
        ConditionType.Subject,
        ConditionType.Body,
        ConditionType.BodyOrSubject,
        ConditionType.MessageHeader,
        ConditionType.RecipientAddress,
        ConditionType.SenderAddress,
        ConditionType.Category,
        ConditionType.Importance,
        ConditionType.Account,
        ConditionType.OnlyToMe,
        ConditionType.To,
        ConditionType.Cc,
        ConditionType.ToOrCc,
        ConditionType.NotTo,
        ConditionType.HasAttachment,
        ConditionType.MeetingInviteOrUpdate,
        ConditionType.LocalMachineOnly,
        ConditionType.AnyCategory,
        ConditionType.FromAnyRssFeed,
    };

    public static readonly IReadOnlySet<ActionType> CopyableActions = new HashSet<ActionType>
    {
        ActionType.MoveToFolder,
        ActionType.CopyToFolder,
        ActionType.AssignToCategory,
        ActionType.ClearCategories,
        ActionType.Delete,
        ActionType.DeletePermanently,
        ActionType.Forward,
        ActionType.ForwardAsAttachment,
        ActionType.Redirect,
        ActionType.CcMessage,
        ActionType.Stop,
        ActionType.PlaySound,
        ActionType.NewItemAlert,
        ActionType.DesktopAlert,
        ActionType.NotifyRead,
        ActionType.NotifyDelivery,
        ActionType.MarkAsTask,
    };

    public static bool CanCopy(RuleData rule) =>
        rule.Conditions.All(c => CopyableConditions.Contains(c.Type))
        && rule.Exceptions.All(c => CopyableConditions.Contains(c.Type))
        && rule.Actions.All(a => CopyableActions.Contains(a.Type));

    /// <summary>Names of the parts that cannot be duplicated (for display).</summary>
    public static IEnumerable<string> UncopyableParts(RuleData rule) =>
        rule.Conditions.Where(c => !CopyableConditions.Contains(c.Type)).Select(c => Loc.T("条件: ", "Condition: ") + RuleText.Label(c.Type))
            .Concat(rule.Exceptions.Where(c => !CopyableConditions.Contains(c.Type)).Select(c => Loc.T("例外: ", "Exception: ") + RuleText.Label(c.Type)))
            .Concat(rule.Actions.Where(a => !CopyableActions.Contains(a.Type)).Select(a => Loc.T("処理: ", "Action: ") + RuleText.Label(a.Type)));
}
