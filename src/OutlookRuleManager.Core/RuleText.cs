namespace OutlookRuleManager.Core;

/// <summary>条件・処理の日本語表記と、一覧・検索用の文章化。</summary>
public static class RuleText
{
    public static string Label(ConditionType type) => type switch
    {
        ConditionType.From => "差出人",
        ConditionType.Subject => "件名に含む語",
        ConditionType.Account => "アカウント",
        ConditionType.OnlyToMe => "自分だけが宛先",
        ConditionType.To => "宛先に自分の名前がある",
        ConditionType.Importance => "重要度",
        ConditionType.Sensitivity => "秘密度",
        ConditionType.FlaggedForAction => "フラグ付き",
        ConditionType.Cc => "CC に自分の名前がある",
        ConditionType.ToOrCc => "宛先か CC に自分の名前がある",
        ConditionType.NotTo => "宛先に自分の名前がない",
        ConditionType.SentTo => "宛先",
        ConditionType.Body => "本文に含む語",
        ConditionType.BodyOrSubject => "件名か本文に含む語",
        ConditionType.MessageHeader => "ヘッダーに含む語",
        ConditionType.RecipientAddress => "宛先アドレスに含む語",
        ConditionType.SenderAddress => "差出人アドレスに含む語",
        ConditionType.Category => "分類項目",
        ConditionType.OutOfOffice => "自動応答",
        ConditionType.HasAttachment => "添付ファイルあり",
        ConditionType.SizeRange => "サイズの範囲",
        ConditionType.DateRange => "受信日の範囲",
        ConditionType.FormName => "フォーム",
        ConditionType.Property => "フォームのプロパティ",
        ConditionType.SenderInAddressBook => "差出人がアドレス帳にある",
        ConditionType.MeetingInviteOrUpdate => "会議出席依頼・更新",
        ConditionType.LocalMachineOnly => "このコンピューターのみ",
        ConditionType.OtherMachine => "特定のコンピューターのみ",
        ConditionType.AnyCategory => "分類項目あり",
        ConditionType.FromRssFeed => "RSS フィード",
        ConditionType.FromAnyRssFeed => "任意の RSS フィード",
        _ => "不明な条件",
    };

    public static string Label(ActionType type) => type switch
    {
        ActionType.MoveToFolder => "移動",
        ActionType.AssignToCategory => "分類項目を設定",
        ActionType.Delete => "削除",
        ActionType.DeletePermanently => "完全に削除",
        ActionType.CopyToFolder => "コピー",
        ActionType.Forward => "転送",
        ActionType.ForwardAsAttachment => "添付ファイルとして転送",
        ActionType.Redirect => "リダイレクト",
        ActionType.ServerReply => "サーバーで返信",
        ActionType.Template => "テンプレートで返信",
        ActionType.FlagForActionInDays => "フラグを設定",
        ActionType.FlagColor => "フラグの色",
        ActionType.FlagClear => "フラグをクリア",
        ActionType.Importance => "重要度を設定",
        ActionType.Sensitivity => "秘密度を設定",
        ActionType.Print => "印刷",
        ActionType.PlaySound => "サウンド",
        ActionType.StartApplication => "アプリ起動",
        ActionType.MarkRead => "開封済みにする",
        ActionType.RunScript => "スクリプト実行",
        ActionType.Stop => "処理を中止",
        ActionType.CustomAction => "カスタム処理",
        ActionType.NewItemAlert => "新着通知",
        ActionType.DesktopAlert => "デスクトップ通知",
        ActionType.NotifyRead => "開封確認",
        ActionType.NotifyDelivery => "配信確認",
        ActionType.CcMessage => "CC を送信",
        ActionType.Defer => "配信を遅らせる",
        ActionType.MarkAsTask => "タスクとしてマーク",
        ActionType.ClearCategories => "分類項目をクリア",
        _ => "不明な処理",
    };

    public static string Describe(RuleCondition c) =>
        c.Values.Count == 0 ? Label(c.Type) : $"{Label(c.Type)}: {string.Join(", ", c.Values.Select(v => v.Display))}";

    public static string Describe(RuleAction a)
    {
        if (a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder)
            return a.FolderMissing ? $"{Label(a.Type)}: (フォルダーが見つかりません)" : $"{Label(a.Type)}: {a.Folder?.DisplayPath}";
        return a.Values.Count == 0 ? Label(a.Type) : $"{Label(a.Type)}: {string.Join(", ", a.Values.Select(v => v.Display))}";
    }

    public static string Summary(IEnumerable<RuleCondition> conditions) => string.Join(" / ", conditions.Select(Describe));

    /// <summary>移動・コピー以外の処理の要約（移動先は別の列に出すため）。</summary>
    public static string ActionSummary(IEnumerable<RuleAction> actions) =>
        string.Join(" / ", actions.Select(a => a.Type is ActionType.MoveToFolder ? Label(a.Type) : Describe(a)));

    /// <summary>移動先フォルダー（移動処理がなければ空）。</summary>
    public static string MoveTarget(IEnumerable<RuleAction> actions)
    {
        var move = actions.FirstOrDefault(a => a.Type == ActionType.MoveToFolder);
        if (move is null) return "";
        return move.FolderMissing ? "(フォルダーが見つかりません)" : move.Folder?.DisplayPath ?? "";
    }

    public static string KindText(RuleEntry e) =>
        (e.Kind == RuleKind.Send ? "送信" : "受信") + (e.Source.IsLocalRule ? "・クライアントのみ" : "");

    /// <summary>検索対象の文字列（名前・条件・例外・処理・移動先をまとめたもの）。</summary>
    public static string SearchText(RuleEntry e) =>
        string.Join("\n",
            e.Name,
            Summary(e.Conditions),
            Summary(e.Exceptions),
            string.Join(" / ", e.Actions.Select(Describe)));
}
