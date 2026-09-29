using OutlookRuleManager.Core;
using static OutlookRuleManager.Outlook.OutlookCom;

namespace OutlookRuleManager.Outlook;

/// <summary>
/// Copies the conditions, exceptions and actions of an existing rule to a new rule (for duplicates, late bound).
/// The supported types must match the lists in RuleCapabilities.
/// </summary>
internal static class OutlookRuleWriter
{
    public static void Copy(dynamic source, dynamic target)
    {
        CopyConditions((object)source.Conditions, (object)target.Conditions, Loc.T("条件", "Condition"));
        CopyConditions((object)source.Exceptions, (object)target.Exceptions, Loc.T("例外", "Exception"));
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
                    throw new NotSupportedException(Loc.T($"{kind}「{RuleText.Label(type)}」は複製できません。", $"{kind} \"{RuleText.Label(type)}\" cannot be duplicated."));
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
                    throw new NotSupportedException(Loc.T($"処理「{RuleText.Label(type)}」は複製できません。", $"Action \"{RuleText.Label(type)}\" cannot be duplicated."));
            }
        }
    }

    /// <summary>Adds the recipients again by address and resolves them (a display name could resolve to someone else).</summary>
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
