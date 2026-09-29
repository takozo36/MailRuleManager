using OutlookRuleManager.Core;
using static OutlookRuleManager.Outlook.OutlookCom;

namespace OutlookRuleManager.Outlook;

/// <summary>
/// 既存ルールの条件・例外・処理を新しいルールへ書き写す（複製用、遅延バインディング）。
/// 扱える種類は RuleCapabilities の一覧と一致させること。
/// </summary>
internal static class OutlookRuleWriter
{
    public static void Copy(dynamic source, dynamic target)
    {
        CopyConditions((object)source.Conditions, (object)target.Conditions, "条件");
        CopyConditions((object)source.Exceptions, (object)target.Exceptions, "例外");
        CopyActions((object)source.Actions, (object)target.Actions);
    }

    private static void CopyConditions(dynamic s, dynamic d, string kind)
    {
        foreach (dynamic c in s)
        {
            if (!(bool)c.Enabled) continue;
            var type = (ConditionType)(int)c.ConditionType;
            switch (type)
            {
                case ConditionType.From:
                    CopyRecipients((object)s.From.Recipients, (object)d.From.Recipients); d.From.Enabled = true; break;
                case ConditionType.SentTo:
                    CopyRecipients((object)s.SentTo.Recipients, (object)d.SentTo.Recipients); d.SentTo.Enabled = true; break;
                case ConditionType.Subject:
                    SetProperty(d.Subject, "Text", s.Subject.Text); d.Subject.Enabled = true; break;
                case ConditionType.Body:
                    SetProperty(d.Body, "Text", s.Body.Text); d.Body.Enabled = true; break;
                case ConditionType.BodyOrSubject:
                    SetProperty(d.BodyOrSubject, "Text", s.BodyOrSubject.Text); d.BodyOrSubject.Enabled = true; break;
                case ConditionType.MessageHeader:
                    SetProperty(d.MessageHeader, "Text", s.MessageHeader.Text); d.MessageHeader.Enabled = true; break;
                case ConditionType.RecipientAddress:
                    SetProperty(d.RecipientAddress, "Address", s.RecipientAddress.Address); d.RecipientAddress.Enabled = true; break;
                case ConditionType.SenderAddress:
                    SetProperty(d.SenderAddress, "Address", s.SenderAddress.Address); d.SenderAddress.Enabled = true; break;
                case ConditionType.Category:
                    SetProperty(d.Category, "Categories", s.Category.Categories); d.Category.Enabled = true; break;
                case ConditionType.Importance:
                    d.Importance.Importance = s.Importance.Importance; d.Importance.Enabled = true; break;
                case ConditionType.Account:
                    SetProperty(d.Account, "Account", s.Account.Account); d.Account.Enabled = true; break;
                case ConditionType.OnlyToMe: d.OnlyToMe.Enabled = true; break;
                case ConditionType.To: d.ToMe.Enabled = true; break;
                case ConditionType.Cc: d.CC.Enabled = true; break;
                case ConditionType.ToOrCc: d.ToOrCc.Enabled = true; break;
                case ConditionType.NotTo: d.NotTo.Enabled = true; break;
                case ConditionType.HasAttachment: d.HasAttachment.Enabled = true; break;
                case ConditionType.MeetingInviteOrUpdate: d.MeetingInviteOrUpdate.Enabled = true; break;
                case ConditionType.LocalMachineOnly: d.OnLocalMachine.Enabled = true; break;
                case ConditionType.AnyCategory: d.AnyCategory.Enabled = true; break;
                case ConditionType.FromAnyRssFeed: d.FromAnyRSSFeed.Enabled = true; break;
                default:
                    throw new NotSupportedException($"{kind}「{RuleText.Label(type)}」は複製できません。");
            }
        }
    }

    private static void CopyActions(dynamic s, dynamic d)
    {
        foreach (dynamic a in s)
        {
            if (!(bool)a.Enabled) continue;
            var type = (ActionType)(int)a.ActionType;
            switch (type)
            {
                case ActionType.MoveToFolder:
                    SetProperty(d.MoveToFolder, "Folder", s.MoveToFolder.Folder); d.MoveToFolder.Enabled = true; break;
                case ActionType.CopyToFolder:
                    SetProperty(d.CopyToFolder, "Folder", s.CopyToFolder.Folder); d.CopyToFolder.Enabled = true; break;
                case ActionType.AssignToCategory:
                    SetProperty(d.AssignToCategory, "Categories", s.AssignToCategory.Categories); d.AssignToCategory.Enabled = true; break;
                case ActionType.Forward:
                    CopyRecipients((object)s.Forward.Recipients, (object)d.Forward.Recipients); d.Forward.Enabled = true; break;
                case ActionType.ForwardAsAttachment:
                    CopyRecipients((object)s.ForwardAsAttachment.Recipients, (object)d.ForwardAsAttachment.Recipients); d.ForwardAsAttachment.Enabled = true; break;
                case ActionType.Redirect:
                    CopyRecipients((object)s.Redirect.Recipients, (object)d.Redirect.Recipients); d.Redirect.Enabled = true; break;
                case ActionType.CcMessage:
                    CopyRecipients((object)s.CC.Recipients, (object)d.CC.Recipients); d.CC.Enabled = true; break;
                case ActionType.PlaySound:
                    d.PlaySound.FilePath = s.PlaySound.FilePath; d.PlaySound.Enabled = true; break;
                case ActionType.NewItemAlert:
                    d.NewItemAlert.Text = s.NewItemAlert.Text; d.NewItemAlert.Enabled = true; break;
                case ActionType.MarkAsTask:
                    d.MarkAsTask.FlagTo = s.MarkAsTask.FlagTo;
                    d.MarkAsTask.MarkInterval = s.MarkAsTask.MarkInterval;
                    d.MarkAsTask.Enabled = true;
                    break;
                case ActionType.ClearCategories: d.ClearCategories.Enabled = true; break;
                case ActionType.Delete: d.Delete.Enabled = true; break;
                case ActionType.DeletePermanently: d.DeletePermanently.Enabled = true; break;
                case ActionType.Stop: d.Stop.Enabled = true; break;
                case ActionType.DesktopAlert: d.DesktopAlert.Enabled = true; break;
                case ActionType.NotifyRead: d.NotifyRead.Enabled = true; break;
                case ActionType.NotifyDelivery: d.NotifyDelivery.Enabled = true; break;
                default:
                    throw new NotSupportedException($"処理「{RuleText.Label(type)}」は複製できません。");
            }
        }
    }

    /// <summary>宛先をアドレスで追加し直して解決する（表示名だと別人に解決されることがあるため）。</summary>
    private static void CopyRecipients(dynamic source, dynamic target)
    {
        foreach (dynamic r in source)
        {
            string? address = null;
            try { address = (string?)r.Address; } catch { }
            target.Add(string.IsNullOrEmpty(address) ? (string)r.Name : address);
        }
        target.ResolveAll();
    }
}
