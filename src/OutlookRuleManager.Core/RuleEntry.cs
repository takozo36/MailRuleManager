namespace OutlookRuleManager.Core;

/// <summary>
/// 編集中のルール 1 件。Source は読み込み時のルール（複製で作ったものは複製元）。
/// 名前・有効/無効・移動先フォルダーだけを上書きでき、条件などは Source のまま。
/// </summary>
public sealed record RuleEntry
{
    public required string Id { get; init; }
    public required RuleData Source { get; init; }
    /// <summary>複製で新しく作るルールなら true（Outlook にはまだ存在しない）。</summary>
    public bool IsNew { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; }
    /// <summary>移動先フォルダーの変更。null なら変更なし。</summary>
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

    /// <summary>移動先フォルダーの変更を反映した処理の一覧。</summary>
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

    /// <summary>新規、または名前・有効/無効・移動先のいずれかを変更したもの。並び順の変更は含まない。</summary>
    public bool IsModified => IsNew || IsRenamed || IsEnabledChanged || IsFolderChanged;

    /// <summary>このアプリで複製できるか（Outlook の画面でしか設定できない条件・処理を含まないか）。</summary>
    public bool CanDuplicate => Source.ReadError is null && RuleCapabilities.CanCopy(Source);
}
