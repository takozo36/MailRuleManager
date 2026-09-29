namespace OutlookRuleManager.Core;

/// <summary>
/// A rule being edited. Source is the rule as loaded (for a duplicate, the rule it was copied from).
/// Only the name, the enabled flag and the move-to folder can be overridden; conditions etc. stay as in Source.
/// </summary>
public sealed record RuleEntry
{
    public required string Id { get; init; }
    public required RuleData Source { get; init; }
    /// <summary>True for a duplicate that does not exist in Outlook yet.</summary>
    public bool IsNew { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; }
    /// <summary>New move-to folder; null means unchanged.</summary>
    public FolderRef? MoveFolderOverride { get; init; }

    public static RuleEntry FromSource(RuleData source) => new()
    {
        Id = $"R{source.Index}",
        Source = source,
        Name = source.Name,
        Enabled = source.Enabled,
    };

    public RuleKind Kind => Source.Kind;
    public IReadOnlyList<RuleCondition> Conditions => Source.Conditions;
    public IReadOnlyList<RuleCondition> Exceptions => Source.Exceptions;

    /// <summary>Actions with the move-to folder override applied.</summary>
    public IReadOnlyList<RuleAction> Actions =>
        MoveFolderOverride is null
            ? Source.Actions
            : Source.Actions
                .Select(a => a.Type == ActionType.MoveToFolder ? RuleAction.ToFolder(ActionType.MoveToFolder, MoveFolderOverride) : a)
                .ToList();

    public bool HasMoveAction => Source.Actions.Any(a => a.Type == ActionType.MoveToFolder);

    public bool IsRenamed => !IsNew && Name != Source.Name;
    public bool IsEnabledChanged => !IsNew && Enabled != Source.Enabled;
    public bool IsFolderChanged => MoveFolderOverride is not null;

    /// <summary>New, or renamed / enabled flag changed / folder changed. Reordering is not included.</summary>
    public bool IsModified => IsNew || IsRenamed || IsEnabledChanged || IsFolderChanged;

    /// <summary>Whether this app can duplicate the rule (no conditions or actions that only Outlook's own UI can set).</summary>
    public bool CanDuplicate => Source.ReadError is null && RuleCapabilities.CanCopy(Source);
}
