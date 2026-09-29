namespace OutlookRuleManager.Core;

/// <summary>Receive or send rule (values match Outlook's OlRuleType).</summary>
public enum RuleKind
{
    Receive = 0,
    Send = 1,
}

/// <summary>Condition / exception type (values match Outlook's OlRuleConditionType).</summary>
public enum ConditionType
{
    Unknown = 0,
    From = 1,
    Subject = 2,
    Account = 3,
    OnlyToMe = 4,
    To = 5,
    Importance = 6,
    Sensitivity = 7,
    FlaggedForAction = 8,
    Cc = 9,
    ToOrCc = 10,
    NotTo = 11,
    SentTo = 12,
    Body = 13,
    BodyOrSubject = 14,
    MessageHeader = 15,
    RecipientAddress = 16,
    SenderAddress = 17,
    Category = 18,
    OutOfOffice = 19,
    HasAttachment = 20,
    SizeRange = 21,
    DateRange = 22,
    FormName = 23,
    Property = 24,
    SenderInAddressBook = 25,
    MeetingInviteOrUpdate = 26,
    LocalMachineOnly = 27,
    OtherMachine = 28,
    AnyCategory = 29,
    FromRssFeed = 30,
    FromAnyRssFeed = 31,
}

/// <summary>Action type (values match Outlook's OlRuleActionType).</summary>
public enum ActionType
{
    Unknown = 0,
    MoveToFolder = 1,
    AssignToCategory = 2,
    Delete = 3,
    DeletePermanently = 4,
    CopyToFolder = 5,
    Forward = 6,
    ForwardAsAttachment = 7,
    Redirect = 8,
    ServerReply = 9,
    Template = 10,
    FlagForActionInDays = 11,
    FlagColor = 12,
    FlagClear = 13,
    Importance = 14,
    Sensitivity = 15,
    Print = 16,
    PlaySound = 17,
    StartApplication = 18,
    MarkRead = 19,
    RunScript = 20,
    Stop = 21,
    CustomAction = 22,
    NewItemAlert = 23,
    DesktopAlert = 24,
    NotifyRead = 25,
    NotifyDelivery = 26,
    CcMessage = 27,
    Defer = 28,
    MarkAsTask = 29,
    ClearCategories = 30,
}

public enum Severity
{
    Info = 0,
    Warning = 1,
    Error = 2,
}
