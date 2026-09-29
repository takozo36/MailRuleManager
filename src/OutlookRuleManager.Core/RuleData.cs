namespace OutlookRuleManager.Core;

/// <summary>
/// One value of a condition or action.
/// Display is for the screen; Key is for comparison (lower-cased e-mail address or word).
/// </summary>
public sealed record RuleValue(string Display, string Key, bool Resolved = true)
{
    public static RuleValue Text(string text) => new(text, text.Trim().ToLowerInvariant());

    /// <summary>Builds a value from a name and an address; shows only one of them when they are the same.</summary>
    public static RuleValue Address(string? name, string? address, bool resolved = true)
    {
        name = name?.Trim() ?? "";
        address = address?.Trim() ?? "";
        string display = name.Length == 0 ? address
            : address.Length == 0 || string.Equals(name, address, StringComparison.OrdinalIgnoreCase) ? name
            : $"{name} <{address}>";
        string key = (address.Length > 0 ? address : name).ToLowerInvariant();
        return new RuleValue(display, key, resolved);
    }

    /// <summary>
    /// Importance value. The key is language-neutral ("low" / "normal" / "high") and the display text
    /// is produced by <see cref="RuleText"/>, so switching the UI language also switches this text.
    /// </summary>
    public static RuleValue Importance(int outlookImportance) => outlookImportance switch
    {
        0 => new RuleValue("low", "low"),
        2 => new RuleValue("high", "high"),
        _ => new RuleValue("normal", "normal"),
    };
}

/// <summary>Reference to an Outlook folder. Path has the form \\StoreName\Inbox\...</summary>
public sealed record FolderRef(string Path, string EntryId, string StoreId)
{
    /// <summary>Path without the leading "\\StoreName\" part, for display.</summary>
    public string DisplayPath
    {
        get
        {
            if (!Path.StartsWith(@"\\", StringComparison.Ordinal)) return Path;
            int sep = Path.IndexOf('\\', 2);
            return sep < 0 ? Path : Path[(sep + 1)..];
        }
    }
}

public sealed record RuleCondition(ConditionType Type, IReadOnlyList<RuleValue> Values)
{
    public RuleCondition(ConditionType type) : this(type, Array.Empty<RuleValue>()) { }
}

/// <summary>
/// One action. For move / copy actions, Folder holds the destination.
/// FolderMissing means the action targets a folder that cannot be found (for example, a deleted folder).
/// </summary>
public sealed record RuleAction(
    ActionType Type,
    IReadOnlyList<RuleValue> Values,
    FolderRef? Folder = null,
    bool FolderMissing = false)
{
    public RuleAction(ActionType type) : this(type, Array.Empty<RuleValue>()) { }

    public static RuleAction ToFolder(ActionType type, FolderRef? folder) =>
        new(type, Array.Empty<RuleValue>(), folder, folder is null);
}

/// <summary>A rule as it was when loaded from Outlook (never changed after loading).</summary>
public sealed record RuleData
{
    /// <summary>Execution order at load time (1-based); equals the position in Outlook's Rules collection.</summary>
    public required int Index { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; } = true;
    public RuleKind Kind { get; init; } = RuleKind.Receive;
    /// <summary>Client-only rule (Outlook's Rule.IsLocalRule).</summary>
    public bool IsLocalRule { get; init; }
    public IReadOnlyList<RuleCondition> Conditions { get; init; } = Array.Empty<RuleCondition>();
    public IReadOnlyList<RuleCondition> Exceptions { get; init; } = Array.Empty<RuleCondition>();
    public IReadOnlyList<RuleAction> Actions { get; init; } = Array.Empty<RuleAction>();
    /// <summary>Error message if the rule could not be read; null when it was read successfully.</summary>
    public string? ReadError { get; init; }
}
