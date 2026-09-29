namespace OutlookRuleManager.Core;

/// <summary>
/// 条件や処理が持つ値 1 件。
/// Display は画面表示用、Key は比較用（アドレスは小文字化したメールアドレス、語句は小文字化した語句）。
/// </summary>
public sealed record RuleValue(string Display, string Key, bool Resolved = true)
{
    public static RuleValue Text(string text) => new(text, text.Trim().ToLowerInvariant());

    /// <summary>名前とアドレスから値を作る。名前とアドレスが同じなら片方だけ表示する。</summary>
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
}

/// <summary>Outlook のフォルダーの参照。Path は \\ストア名\受信トレイ\... 形式。</summary>
public sealed record FolderRef(string Path, string EntryId, string StoreId)
{
    /// <summary>先頭の「\\ストア名\」を除いた表示用パス。</summary>
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
/// 処理 1 件。移動・コピーでは Folder に移動先が入る。
/// FolderMissing は「移動・コピーの処理なのに移動先フォルダーが取得できない」状態（削除されたフォルダーなど）。
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

/// <summary>Outlook から読み込んだ時点のルール 1 件（読み込み後は変更しない）。</summary>
public sealed record RuleData
{
    /// <summary>読み込み時の実行順（1 始まり）。Outlook 側の Rules コレクションの位置と一致する。</summary>
    public required int Index { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; } = true;
    public RuleKind Kind { get; init; } = RuleKind.Receive;
    /// <summary>このコンピューターでのみ実行されるルール（クライアント側ルール）。</summary>
    public bool IsLocalRule { get; init; }
    public IReadOnlyList<RuleCondition> Conditions { get; init; } = Array.Empty<RuleCondition>();
    public IReadOnlyList<RuleCondition> Exceptions { get; init; } = Array.Empty<RuleCondition>();
    public IReadOnlyList<RuleAction> Actions { get; init; } = Array.Empty<RuleAction>();
    /// <summary>読み込み中に例外が出た場合のメッセージ。null なら正常。</summary>
    public string? ReadError { get; init; }
}
