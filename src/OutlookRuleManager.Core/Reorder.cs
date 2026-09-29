namespace OutlookRuleManager.Core;

/// <summary>並べ替えの計算（一覧を受け取って新しい並びを返すだけの純粋な関数）。</summary>
internal static class Reorder
{
    public static List<RuleEntry> MoveUp(IReadOnlyList<RuleEntry> entries, ISet<string> selected, ISet<string>? visible)
    {
        var (slots, view) = VisibleView(entries, visible);
        // 上から順に、直前が非選択なら入れ替える（連続した選択はまとまって 1 つ上へ動く）
        for (int i = 1; i < view.Count; i++)
            if (selected.Contains(view[i].Id) && !selected.Contains(view[i - 1].Id))
                (view[i - 1], view[i]) = (view[i], view[i - 1]);
        return WriteBack(entries, slots, view);
    }

    public static List<RuleEntry> MoveDown(IReadOnlyList<RuleEntry> entries, ISet<string> selected, ISet<string>? visible)
    {
        var (slots, view) = VisibleView(entries, visible);
        for (int i = view.Count - 2; i >= 0; i--)
            if (selected.Contains(view[i].Id) && !selected.Contains(view[i + 1].Id))
                (view[i], view[i + 1]) = (view[i + 1], view[i]);
        return WriteBack(entries, slots, view);
    }

    public static List<RuleEntry> MoveTo(IReadOnlyList<RuleEntry> entries, ISet<string> selected, int position)
    {
        var moving = entries.Where(e => selected.Contains(e.Id)).ToList();
        var rest = entries.Where(e => !selected.Contains(e.Id)).ToList();
        int insertAt = Math.Clamp(position - 1, 0, rest.Count);
        rest.InsertRange(insertAt, moving);
        return rest;
    }

    public static List<RuleEntry> MoveBefore(IReadOnlyList<RuleEntry> entries, ISet<string> selected, string? targetId)
    {
        // 移動するルール自身の直前が指定されたら、その後ろで最初の「移動しないルール」の直前とみなす
        if (targetId is not null && selected.Contains(targetId))
        {
            int at = entries.ToList().FindIndex(e => e.Id == targetId);
            targetId = entries.Skip(at + 1).FirstOrDefault(e => !selected.Contains(e.Id))?.Id;
        }
        var moving = entries.Where(e => selected.Contains(e.Id)).ToList();
        var rest = entries.Where(e => !selected.Contains(e.Id)).ToList();
        int insertAt = targetId is null ? rest.Count : rest.FindIndex(e => e.Id == targetId);
        if (insertAt < 0) insertAt = rest.Count;
        rest.InsertRange(insertAt, moving);
        return rest;
    }

    /// <summary>表示中のルールの位置（slots）と、その並び（view）を取り出す。</summary>
    private static (List<int> slots, List<RuleEntry> view) VisibleView(IReadOnlyList<RuleEntry> entries, ISet<string>? visible)
    {
        var slots = new List<int>();
        var view = new List<RuleEntry>();
        for (int i = 0; i < entries.Count; i++)
        {
            if (visible is not null && !visible.Contains(entries[i].Id)) continue;
            slots.Add(i);
            view.Add(entries[i]);
        }
        return (slots, view);
    }

    private static List<RuleEntry> WriteBack(IReadOnlyList<RuleEntry> entries, List<int> slots, List<RuleEntry> view)
    {
        var result = entries.ToList();
        for (int k = 0; k < slots.Count; k++) result[slots[k]] = view[k];
        return result;
    }
}
