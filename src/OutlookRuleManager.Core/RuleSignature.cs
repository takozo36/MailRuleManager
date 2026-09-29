namespace OutlookRuleManager.Core;

/// <summary>
/// ルールの中身（条件・例外・処理）を比較用の文字列にする。
/// 複製したルールを保存前に読み戻し、元のルールと同じ内容になっているかの確認に使う。
/// 並び順・表示名の違いは無視し、種類と比較用の値（アドレス・語句・フォルダーの EntryID）で比べる。
/// </summary>
public static class RuleSignature
{
    public static string Of(RuleData rule) => Of(rule.Conditions, rule.Exceptions, rule.Actions);

    public static string Of(IEnumerable<RuleCondition> conditions, IEnumerable<RuleCondition> exceptions, IEnumerable<RuleAction> actions) =>
        $"条件[{Conditions(conditions)}] 例外[{Conditions(exceptions)}] 処理[{Actions(actions)}]";

    private static string Conditions(IEnumerable<RuleCondition> conditions) =>
        string.Join(" ", conditions
            .OrderBy(c => c.Type)
            .Select(c => $"{c.Type}({Keys(c.Values)})"));

    private static string Actions(IEnumerable<RuleAction> actions) =>
        string.Join(" ", actions
            .OrderBy(a => a.Type)
            .Select(a => a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder
                ? $"{a.Type}({(a.FolderMissing ? "フォルダーなし" : a.Folder?.EntryId.ToUpperInvariant())})"
                : $"{a.Type}({Keys(a.Values)})"));

    private static string Keys(IEnumerable<RuleValue> values) =>
        string.Join(",", values.Select(v => v.Key).OrderBy(k => k, StringComparer.Ordinal));
}
