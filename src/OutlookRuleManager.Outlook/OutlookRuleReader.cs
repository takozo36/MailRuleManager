using OutlookRuleManager.Core;
using static OutlookRuleManager.Outlook.OutlookCom;

namespace OutlookRuleManager.Outlook;

/// <summary>Outlook の Rule オブジェクト（遅延バインディング）を RuleData に読み替える。</summary>
internal static class OutlookRuleReader
{
    public static RuleData Read(dynamic rule, int index)
    {
        string name = "(名前を取得できません)";
        try
        {
            name = (string)rule.Name;
            return new RuleData
            {
                Index = index,
                Name = name,
                Enabled = (bool)rule.Enabled,
                Kind = (RuleKind)(int)rule.RuleType,
                IsLocalRule = (bool)rule.IsLocalRule,
                Conditions = ReadConditions((object)rule.Conditions),
                Exceptions = ReadConditions((object)rule.Exceptions),
                Actions = ReadActions((object)rule.Actions),
            };
        }
        catch (Exception ex)
        {
            return new RuleData { Index = index, Name = name, Enabled = false, ReadError = ex.Message };
        }
    }

    private static List<RuleCondition> ReadConditions(dynamic conditions)
    {
        var list = new List<RuleCondition>();
        foreach (dynamic c in conditions)
        {
            if (!(bool)c.Enabled) continue;
            var type = (ConditionType)(int)c.ConditionType;
            list.Add(new RuleCondition(type, SafeValues(() => ConditionValues((object)conditions, type))));
        }
        return list;
    }

    private static IReadOnlyList<RuleValue> ConditionValues(dynamic cs, ConditionType type) => type switch
    {
        ConditionType.From => Recipients((object)cs.From.Recipients),
        ConditionType.SentTo => Recipients((object)cs.SentTo.Recipients),
        ConditionType.Subject => Texts((object?)cs.Subject.Text),
        ConditionType.Body => Texts((object?)cs.Body.Text),
        ConditionType.BodyOrSubject => Texts((object?)cs.BodyOrSubject.Text),
        ConditionType.MessageHeader => Texts((object?)cs.MessageHeader.Text),
        ConditionType.RecipientAddress => Texts((object?)cs.RecipientAddress.Address),
        ConditionType.SenderAddress => Texts((object?)cs.SenderAddress.Address),
        ConditionType.Category => Texts((object?)cs.Category.Categories),
        ConditionType.FormName => Texts((object?)cs.FormName.FormName),
        ConditionType.FromRssFeed => Texts((object?)cs.FromRssFeed.FromRssFeed),
        ConditionType.Importance => [RuleValue.Text(ImportanceText((int)cs.Importance.Importance))],
        ConditionType.Account => [RuleValue.Text((string?)cs.Account.Account?.DisplayName ?? "")],
        ConditionType.SenderInAddressBook => [RuleValue.Text((string?)cs.SenderInAddressList.AddressList?.Name ?? "")],
        _ => Array.Empty<RuleValue>(),
    };

    private static List<RuleAction> ReadActions(dynamic actions)
    {
        var list = new List<RuleAction>();
        foreach (dynamic a in actions)
        {
            if (!(bool)a.Enabled) continue;
            var type = (ActionType)(int)a.ActionType;
            list.Add(type switch
            {
                ActionType.MoveToFolder => RuleAction.ToFolder(type, ReadFolder(() => (object?)actions.MoveToFolder.Folder)),
                ActionType.CopyToFolder => RuleAction.ToFolder(type, ReadFolder(() => (object?)actions.CopyToFolder.Folder)),
                _ => new RuleAction(type, SafeValues(() => ActionValues((object)actions, type))),
            });
        }
        return list;
    }

    private static IReadOnlyList<RuleValue> ActionValues(dynamic acts, ActionType type) => type switch
    {
        ActionType.AssignToCategory => Texts((object?)acts.AssignToCategory.Categories),
        ActionType.Forward => Recipients((object)acts.Forward.Recipients),
        ActionType.ForwardAsAttachment => Recipients((object)acts.ForwardAsAttachment.Recipients),
        ActionType.Redirect => Recipients((object)acts.Redirect.Recipients),
        ActionType.CcMessage => Recipients((object)acts.CC.Recipients),
        ActionType.PlaySound => [RuleValue.Text((string?)acts.PlaySound.FilePath ?? "")],
        ActionType.NewItemAlert => [RuleValue.Text((string?)acts.NewItemAlert.Text ?? "")],
        _ => Array.Empty<RuleValue>(),
    };

    /// <summary>移動先フォルダー。削除済みなどで取れなければ null（＝エラーのルール）。</summary>
    private static FolderRef? ReadFolder(Func<object?> getFolder)
    {
        try
        {
            var f = getFolder();
            return f is null ? null : ToFolderRef(f);
        }
        catch
        {
            return null;
        }
    }

    internal static FolderRef ToFolderRef(dynamic f) => new((string)f.FolderPath, (string)f.EntryID, (string)f.StoreID);

    private static IReadOnlyList<RuleValue> Recipients(dynamic recipients)
    {
        var list = new List<RuleValue>();
        foreach (dynamic r in recipients)
        {
            string? address = null;
            try { address = (string?)r.Address; } catch { /* 未解決の宛先では例外になることがある */ }
            list.Add(RuleValue.Address((string?)r.Name, address, (bool)r.Resolved));
        }
        return list;
    }

    /// <summary>文字列の配列（COM の VARIANT 配列）または単一の文字列を値の一覧にする。</summary>
    private static IReadOnlyList<RuleValue> Texts(object? value) => value switch
    {
        null => Array.Empty<RuleValue>(),
        string s => s.Length == 0 ? Array.Empty<RuleValue>() : [RuleValue.Text(s)],
        Array arr => arr.Cast<object?>().Select(o => o?.ToString() ?? "").Where(s => s.Length > 0).Select(RuleValue.Text).ToList(),
        _ => [RuleValue.Text(value.ToString() ?? "")],
    };


    private static IReadOnlyList<RuleValue> SafeValues(Func<IReadOnlyList<RuleValue>> read)
    {
        try { return read(); }
        catch (Exception ex) { return [new RuleValue($"(値を読み取れません: {ex.Message})", "")]; }
    }

    private static string ImportanceText(int importance) => importance switch
    {
        OlImportanceHigh => "高",
        OlImportanceLow => "低",
        _ => "標準",
    };
}
