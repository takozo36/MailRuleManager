using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.Core;

/// <summary>Display names of conditions and actions, and the text used by the list and by search.</summary>
public static class RuleText
{
    public static string Label(ConditionType type) => type switch
    {
        ConditionType.From => T("差出人", "From"),
        ConditionType.Subject => T("件名に含む語", "Subject contains"),
        ConditionType.Account => T("アカウント", "Through account"),
        ConditionType.OnlyToMe => T("自分だけが宛先", "Sent only to me"),
        ConditionType.To => T("宛先に自分の名前がある", "My name in To"),
        ConditionType.Importance => T("重要度", "Importance"),
        ConditionType.Sensitivity => T("秘密度", "Sensitivity"),
        ConditionType.FlaggedForAction => T("フラグ付き", "Flagged for action"),
        ConditionType.Cc => T("CC に自分の名前がある", "My name in Cc"),
        ConditionType.ToOrCc => T("宛先か CC に自分の名前がある", "My name in To or Cc"),
        ConditionType.NotTo => T("宛先に自分の名前がない", "My name not in To"),
        ConditionType.SentTo => T("宛先", "Sent to"),
        ConditionType.Body => T("本文に含む語", "Body contains"),
        ConditionType.BodyOrSubject => T("件名か本文に含む語", "Subject or body contains"),
        ConditionType.MessageHeader => T("ヘッダーに含む語", "Header contains"),
        ConditionType.RecipientAddress => T("宛先アドレスに含む語", "Recipient address contains"),
        ConditionType.SenderAddress => T("差出人アドレスに含む語", "Sender address contains"),
        ConditionType.Category => T("分類項目", "Category"),
        ConditionType.OutOfOffice => T("自動応答", "Automatic reply"),
        ConditionType.HasAttachment => T("添付ファイルあり", "Has attachment"),
        ConditionType.SizeRange => T("サイズの範囲", "Size range"),
        ConditionType.DateRange => T("受信日の範囲", "Received date range"),
        ConditionType.FormName => T("フォーム", "Form"),
        ConditionType.Property => T("フォームのプロパティ", "Form property"),
        ConditionType.SenderInAddressBook => T("差出人がアドレス帳にある", "Sender in address book"),
        ConditionType.MeetingInviteOrUpdate => T("会議出席依頼・更新", "Meeting invitation or update"),
        ConditionType.LocalMachineOnly => T("このコンピューターのみ", "This computer only"),
        ConditionType.OtherMachine => T("特定のコンピューターのみ", "Specific computer only"),
        ConditionType.AnyCategory => T("分類項目あり", "Any category"),
        ConditionType.FromRssFeed => T("RSS フィード", "RSS feed"),
        ConditionType.FromAnyRssFeed => T("任意の RSS フィード", "Any RSS feed"),
        _ => T("不明な条件", "Unknown condition"),
    };

    public static string Label(ActionType type) => type switch
    {
        ActionType.MoveToFolder => T("移動", "Move"),
        ActionType.AssignToCategory => T("分類項目を設定", "Assign category"),
        ActionType.Delete => T("削除", "Delete"),
        ActionType.DeletePermanently => T("完全に削除", "Delete permanently"),
        ActionType.CopyToFolder => T("コピー", "Copy"),
        ActionType.Forward => T("転送", "Forward"),
        ActionType.ForwardAsAttachment => T("添付ファイルとして転送", "Forward as attachment"),
        ActionType.Redirect => T("リダイレクト", "Redirect"),
        ActionType.ServerReply => T("サーバーで返信", "Server reply"),
        ActionType.Template => T("テンプレートで返信", "Reply with template"),
        ActionType.FlagForActionInDays => T("フラグを設定", "Flag"),
        ActionType.FlagColor => T("フラグの色", "Flag color"),
        ActionType.FlagClear => T("フラグをクリア", "Clear flag"),
        ActionType.Importance => T("重要度を設定", "Set importance"),
        ActionType.Sensitivity => T("秘密度を設定", "Set sensitivity"),
        ActionType.Print => T("印刷", "Print"),
        ActionType.PlaySound => T("サウンド", "Play sound"),
        ActionType.StartApplication => T("アプリ起動", "Start application"),
        ActionType.MarkRead => T("開封済みにする", "Mark as read"),
        ActionType.RunScript => T("スクリプト実行", "Run script"),
        ActionType.Stop => T("処理を中止", "Stop processing"),
        ActionType.CustomAction => T("カスタム処理", "Custom action"),
        ActionType.NewItemAlert => T("新着通知", "New item alert"),
        ActionType.DesktopAlert => T("デスクトップ通知", "Desktop alert"),
        ActionType.NotifyRead => T("開封確認", "Read receipt"),
        ActionType.NotifyDelivery => T("配信確認", "Delivery receipt"),
        ActionType.CcMessage => T("CC を送信", "Cc"),
        ActionType.Defer => T("配信を遅らせる", "Defer delivery"),
        ActionType.MarkAsTask => T("タスクとしてマーク", "Mark as task"),
        ActionType.ClearCategories => T("分類項目をクリア", "Clear categories"),
        _ => T("不明な処理", "Unknown action"),
    };

    public static string Describe(RuleCondition c) =>
        c.Values.Count == 0 ? Label(c.Type) : $"{Label(c.Type)}: {string.Join(", ", c.Values.Select(v => ValueText(c.Type, v)))}";

    public static string Describe(RuleAction a)
    {
        if (a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder)
            return a.FolderMissing ? $"{Label(a.Type)}: {MissingFolder}" : $"{Label(a.Type)}: {a.Folder?.DisplayPath}";
        return a.Values.Count == 0 ? Label(a.Type) : $"{Label(a.Type)}: {string.Join(", ", a.Values.Select(v => v.Display))}";
    }

    public static string Summary(IEnumerable<RuleCondition> conditions) => string.Join(" / ", conditions.Select(Describe));

    /// <summary>Summary of the actions; the move-to folder is only named, because it has its own column.</summary>
    public static string ActionSummary(IEnumerable<RuleAction> actions) =>
        string.Join(" / ", actions.Select(a => a.Type is ActionType.MoveToFolder ? Label(a.Type) : Describe(a)));

    /// <summary>Move-to folder (empty if the rule has no move action).</summary>
    public static string MoveTarget(IEnumerable<RuleAction> actions)
    {
        var move = actions.FirstOrDefault(a => a.Type == ActionType.MoveToFolder);
        if (move is null) return "";
        return move.FolderMissing ? MissingFolder : move.Folder?.DisplayPath ?? "";
    }

    public static string KindText(RuleEntry e) =>
        (e.Kind == RuleKind.Send ? T("送信", "Send") : T("受信", "Receive"))
        + (e.Source.IsLocalRule ? T("・クライアントのみ", " · client only") : "");

    public static string EnabledText(bool enabled) => enabled ? T("有効", "Enabled") : T("無効", "Disabled");

    public static string SeverityText(Severity s) => s switch
    {
        Severity.Error => T("エラー", "Error"),
        Severity.Warning => T("警告", "Warning"),
        _ => T("情報", "Info"),
    };

    /// <summary>Text searched by the search box (name, conditions, exceptions, actions and move-to folder).</summary>
    public static string SearchText(RuleEntry e) =>
        string.Join("\n",
            e.Name,
            Summary(e.Conditions),
            Summary(e.Exceptions),
            string.Join(" / ", e.Actions.Select(Describe)));

    private static string MissingFolder => T("(フォルダーが見つかりません)", "(folder not found)");

    private static string ValueText(ConditionType type, RuleValue v) =>
        type != ConditionType.Importance ? v.Display : v.Key switch
        {
            "high" => T("高", "High"),
            "low" => T("低", "Low"),
            _ => T("標準", "Normal"),
        };
}
