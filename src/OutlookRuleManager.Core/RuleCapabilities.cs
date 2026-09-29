namespace OutlookRuleManager.Core;

/// <summary>
/// Outlook のオブジェクトモデルで新しいルールへ書き込める条件・処理の一覧。
/// 複製はここに載っている種類だけで構成されたルールに限る
/// （Outlook 連携側の OutlookRuleWriter はこの一覧どおりにコピー処理を実装している）。
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

    /// <summary>複製できない条件・処理の名前（画面表示用）。</summary>
    public static IEnumerable<string> UncopyableParts(RuleData rule) =>
        rule.Conditions.Where(c => !CopyableConditions.Contains(c.Type)).Select(c => "条件: " + RuleText.Label(c.Type))
            .Concat(rule.Exceptions.Where(c => !CopyableConditions.Contains(c.Type)).Select(c => "例外: " + RuleText.Label(c.Type)))
            .Concat(rule.Actions.Where(a => !CopyableActions.Contains(a.Type)).Select(a => "処理: " + RuleText.Label(a.Type)));
}
