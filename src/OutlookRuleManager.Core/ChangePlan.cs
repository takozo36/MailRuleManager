namespace OutlookRuleManager.Core;

/// <summary>
/// 読み込み時の状態と編集後の状態の差分。Outlook へ反映するときの手順の元になる。
/// </summary>
public sealed class ChangePlan
{
    private ChangePlan(IReadOnlyList<RuleData> original, IReadOnlyList<RuleEntry> finalOrder, IReadOnlyList<RuleData> deleted, bool orderChanged)
    {
        Original = original;
        FinalOrder = finalOrder;
        Deleted = deleted;
        OrderChanged = orderChanged;
    }

    public IReadOnlyList<RuleData> Original { get; }
    /// <summary>保存後の並び（先頭が実行順 1）。</summary>
    public IReadOnlyList<RuleEntry> FinalOrder { get; }
    /// <summary>削除するルール（読み込み時の位置の昇順）。</summary>
    public IReadOnlyList<RuleData> Deleted { get; }
    /// <summary>残るルールどうしの前後関係が変わったか、途中に新しいルールが入ったか。</summary>
    public bool OrderChanged { get; }

    public IEnumerable<RuleEntry> Created => FinalOrder.Where(e => e.IsNew);
    public IEnumerable<RuleEntry> Modified => FinalOrder.Where(e => !e.IsNew && e.IsModified);

    public bool HasChanges => OrderChanged || Deleted.Count > 0 || FinalOrder.Any(e => e.IsModified);

    /// <summary>変更件数（並び順の変更は 1 件と数える）。</summary>
    public int ChangeCount => Deleted.Count + FinalOrder.Count(e => e.IsModified) + (OrderChanged ? 1 : 0);

    public static ChangePlan Create(IReadOnlyList<RuleData> original, IReadOnlyList<RuleEntry> current)
    {
        var keptIndexes = current.Where(e => !e.IsNew).Select(e => e.Source.Index).ToHashSet();
        var deleted = original.Where(r => !keptIndexes.Contains(r.Index)).OrderBy(r => r.Index).ToList();

        // 残る既存ルールの並びが読み込み時と同じ順（昇順）で、新規ルールが末尾にしかなければ並びは変わっていない
        var existingOrder = current.Where(e => !e.IsNew).Select(e => e.Source.Index).ToList();
        bool ascending = existingOrder.Zip(existingOrder.Skip(1)).All(p => p.First < p.Second);
        int firstNew = current.ToList().FindIndex(e => e.IsNew);
        bool newOnlyAtEnd = firstNew < 0 || current.Skip(firstNew).All(e => e.IsNew);
        bool orderChanged = !ascending || !newOnlyAtEnd;

        return new ChangePlan(original, current, deleted, orderChanged);
    }

    /// <summary>
    /// 保存に成功したあとの Outlook 側の状態を、読み込み直さずに組み立てる
    /// （全件の読み込みには数分かかるため）。実行順・名前・有効/無効・移動先を反映する。
    /// </summary>
    public IReadOnlyList<RuleData> ToSavedSnapshot() =>
        FinalOrder.Select((e, i) => e.Source with
        {
            Index = i + 1,
            Name = e.Name,
            Enabled = e.Enabled,
            Actions = e.Actions,
        }).ToList();

    /// <summary>保存前に止めるべき問題（Outlook に触る前に判定できるもの）。</summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();
        foreach (var e in Created)
        {
            if (e.Actions.Any(a => a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder && a.FolderMissing))
                problems.Add($"複製「{e.Name}」: 移動先・コピー先のフォルダーが見つかりません。先に移動先を指定してください。");
            if (!e.CanDuplicate)
                problems.Add($"複製「{e.Name}」: このアプリでは複製できない条件・処理が含まれています。");
        }
        foreach (var e in FinalOrder.Where(e => e.Name.Trim().Length == 0))
            problems.Add($"実行順 {IndexOf(e) + 1}: ルール名が空です。");
        return problems;
    }

    /// <summary>確認画面に出す変更内容の説明。</summary>
    public IReadOnlyList<string> Describe()
    {
        var lines = new List<string>();
        foreach (var r in Deleted)
            lines.Add($"削除: 「{r.Name}」");
        foreach (var e in Created)
            lines.Add($"追加(複製): 「{e.Name}」 → 実行順 {IndexOf(e) + 1}");
        foreach (var e in Modified)
        {
            var parts = new List<string>();
            if (e.IsRenamed) parts.Add($"名前「{e.Source.Name}」→「{e.Name}」");
            if (e.IsEnabledChanged) parts.Add(e.Enabled ? "有効にする" : "無効にする");
            if (e.IsFolderChanged) parts.Add($"移動先 → {e.MoveFolderOverride!.DisplayPath}");
            lines.Add($"変更: 「{e.Name}」 {string.Join("、", parts)}");
        }
        if (OrderChanged)
            lines.Add("並び順の変更: " + FinalOrder.Count(e => !e.IsNew && IndexOf(e) + 1 != e.Source.Index) + " 件のルールの実行順が変わります");
        return lines;
    }

    private int IndexOf(RuleEntry e)
    {
        for (int i = 0; i < FinalOrder.Count; i++)
            if (FinalOrder[i].Id == e.Id) return i;
        return -1;
    }
}
