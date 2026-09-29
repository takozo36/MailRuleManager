using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.Core;

/// <summary>
/// Difference between the state at load time and the edited state; the basis of the steps applied to Outlook.
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
    /// <summary>Order after saving (the first one runs first).</summary>
    public IReadOnlyList<RuleEntry> FinalOrder { get; }
    /// <summary>Rules to delete (in ascending order of their original position).</summary>
    public IReadOnlyList<RuleData> Deleted { get; }
    /// <summary>The relative order of the remaining rules changed, or a new rule was inserted in the middle.</summary>
    public bool OrderChanged { get; }

    public IEnumerable<RuleEntry> Created => FinalOrder.Where(e => e.IsNew);
    public IEnumerable<RuleEntry> Modified => FinalOrder.Where(e => !e.IsNew && e.IsModified);

    public bool HasChanges => OrderChanged || Deleted.Count > 0 || FinalOrder.Any(e => e.IsModified);

    /// <summary>Number of changes (a reorder counts as one).</summary>
    public int ChangeCount => Deleted.Count + FinalOrder.Count(e => e.IsModified) + (OrderChanged ? 1 : 0);

    public static ChangePlan Create(IReadOnlyList<RuleData> original, IReadOnlyList<RuleEntry> current)
    {
        var keptIndexes = current.Where(e => !e.IsNew).Select(e => e.Source.Index).ToHashSet();
        var deleted = original.Where(r => !keptIndexes.Contains(r.Index)).OrderBy(r => r.Index).ToList();

        // The order is unchanged if the remaining existing rules are still ascending and new rules are only at the end
        var existingOrder = current.Where(e => !e.IsNew).Select(e => e.Source.Index).ToList();
        bool ascending = existingOrder.Zip(existingOrder.Skip(1)).All(p => p.First < p.Second);
        int firstNew = current.ToList().FindIndex(e => e.IsNew);
        bool newOnlyAtEnd = firstNew < 0 || current.Skip(firstNew).All(e => e.IsNew);
        bool orderChanged = !ascending || !newOnlyAtEnd;

        return new ChangePlan(original, current, deleted, orderChanged);
    }

    /// <summary>
    /// Builds Outlook's state after a successful save without reloading (loading everything can take a while).
    /// Reflects the order, names, enabled flags and move-to folders.
    /// </summary>
    public IReadOnlyList<RuleData> ToSavedSnapshot() =>
        FinalOrder.Select((e, i) => e.Source with
        {
            Index = i + 1,
            Name = e.Name,
            Enabled = e.Enabled,
            Actions = e.Actions,
        }).ToList();

    /// <summary>Problems that must stop the save (those that can be detected before touching Outlook).</summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();
        foreach (var e in Created)
        {
            if (e.Actions.Any(a => a.Type is ActionType.MoveToFolder or ActionType.CopyToFolder && a.FolderMissing))
                problems.Add(T($"複製「{e.Name}」: 移動先・コピー先のフォルダーが見つかりません。先に移動先を指定してください。",
                    $"Duplicate \"{e.Name}\": the move-to or copy-to folder was not found. Choose a folder first."));
            if (!e.CanDuplicate)
                problems.Add(T($"複製「{e.Name}」: このアプリでは複製できない条件・処理が含まれています。",
                    $"Duplicate \"{e.Name}\": it contains conditions or actions that this app cannot duplicate."));
        }
        foreach (var e in FinalOrder.Where(e => e.Name.Trim().Length == 0))
            problems.Add(T($"実行順 {IndexOf(e) + 1}: ルール名が空です。", $"Position {IndexOf(e) + 1}: the rule name is empty."));
        return problems;
    }

    /// <summary>Description of the changes for the confirmation dialog.</summary>
    public IReadOnlyList<string> Describe()
    {
        var lines = new List<string>();
        foreach (var r in Deleted)
            lines.Add(T($"削除: 「{r.Name}」", $"Delete: \"{r.Name}\""));
        foreach (var e in Created)
            lines.Add(T($"追加(複製): 「{e.Name}」 → 実行順 {IndexOf(e) + 1}", $"Add (duplicate): \"{e.Name}\" → position {IndexOf(e) + 1}"));
        foreach (var e in Modified)
        {
            var parts = new List<string>();
            if (e.IsRenamed) parts.Add(T($"名前「{e.Source.Name}」→「{e.Name}」", $"name \"{e.Source.Name}\" → \"{e.Name}\""));
            if (e.IsEnabledChanged) parts.Add(e.Enabled ? T("有効にする", "enable") : T("無効にする", "disable"));
            if (e.IsFolderChanged) parts.Add(T($"移動先 → {e.MoveFolderOverride!.DisplayPath}", $"move to → {e.MoveFolderOverride!.DisplayPath}"));
            lines.Add(T($"変更: 「{e.Name}」 {string.Join("、", parts)}", $"Change: \"{e.Name}\" {string.Join(", ", parts)}"));
        }
        if (OrderChanged)
        {
            int moved = FinalOrder.Count(e => !e.IsNew && IndexOf(e) + 1 != e.Source.Index);
            lines.Add(T($"並び順の変更: {moved} 件のルールの実行順が変わります", $"Reorder: {moved} rules change their position"));
        }
        return lines;
    }

    private int IndexOf(RuleEntry e)
    {
        for (int i = 0; i < FinalOrder.Count; i++)
            if (FinalOrder[i].Id == e.Id) return i;
        return -1;
    }
}
