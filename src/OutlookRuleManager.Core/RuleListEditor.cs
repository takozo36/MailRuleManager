namespace OutlookRuleManager.Core;

/// <summary>
/// In-memory editing of the loaded rules (enable/disable, rename, move-to folder, delete, duplicate, reorder).
/// Changes are applied to Outlook all at once from the result of BuildPlan(). Every edit can be undone one step at a time.
/// </summary>
public sealed class RuleListEditor
{
    private readonly Stack<IReadOnlyList<RuleEntry>> _undo = new();
    private IReadOnlyList<RuleEntry> _entries;
    private int _newCounter;

    public RuleListEditor(IReadOnlyList<RuleData> original)
    {
        Original = original;
        _entries = original.Select(RuleEntry.FromSource).ToList();
    }

    public IReadOnlyList<RuleData> Original { get; }

    /// <summary>Rules in their current order (the first one runs first).</summary>
    public IReadOnlyList<RuleEntry> Entries => _entries;

    public bool CanUndo => _undo.Count > 0;

    public bool Undo()
    {
        if (_undo.Count == 0) return false;
        _entries = _undo.Pop();
        return true;
    }

    public ChangePlan BuildPlan() => ChangePlan.Create(Original, _entries);

    public int IndexOf(string id)
    {
        for (int i = 0; i < _entries.Count; i++)
            if (_entries[i].Id == id) return i;
        return -1;
    }

    // ---- Content changes ----

    public int SetEnabled(IEnumerable<string> ids, bool enabled) =>
        Update(ids, e => e.Enabled == enabled ? e : e with { Enabled = enabled });

    public bool Rename(string id, string newName)
    {
        newName = newName.Trim();
        if (newName.Length == 0) throw new ArgumentException(Loc.T("ルール名が空です。", "The rule name is empty."), nameof(newName));
        return Update([id], e => e.Name == newName ? e : e with { Name = newName }) > 0;
    }

    /// <summary>Changes the move-to folder. Rules without a move action are left as they are.</summary>
    public int SetMoveFolder(IEnumerable<string> ids, FolderRef folder) =>
        Update(ids, e =>
        {
            if (!e.HasMoveAction) return e;
            // Setting the original folder again counts as "no change"
            var original = e.Source.Actions.First(a => a.Type == ActionType.MoveToFolder).Folder;
            var newOverride = original is not null && original.EntryId == folder.EntryId && original.StoreId == folder.StoreId && !e.IsNew
                ? null
                : folder;
            return e.MoveFolderOverride == newOverride ? e : e with { MoveFolderOverride = newOverride };
        });

    public int Delete(IEnumerable<string> ids)
    {
        var set = ids.ToHashSet();
        var next = _entries.Where(e => !set.Contains(e.Id)).ToList();
        int removed = _entries.Count - next.Count;
        return Commit(next) ? removed : 0;
    }

    /// <summary>
    /// Duplicates the given rules and inserts each copy right after its original.
    /// Rules that cannot be duplicated (they contain parts only Outlook's own UI can set) are skipped and returned in skipped.
    /// </summary>
    public IReadOnlyList<string> Duplicate(IEnumerable<string> ids, out IReadOnlyList<RuleEntry> skipped)
    {
        var set = ids.ToHashSet();
        var next = new List<RuleEntry>(_entries.Count + set.Count);
        var created = new List<string>();
        var skip = new List<RuleEntry>();
        foreach (var e in _entries)
        {
            next.Add(e);
            if (!set.Contains(e.Id)) continue;
            if (!e.CanDuplicate) { skip.Add(e); continue; }
            var copy = e with
            {
                Id = $"N{++_newCounter}",
                IsNew = true,
                Name = e.Name + Loc.T(" のコピー", " (copy)"),
            };
            next.Add(copy);
            created.Add(copy.Id);
        }
        skipped = skip;
        if (created.Count > 0) Commit(next);
        return created;
    }

    // ---- Reordering ----

    /// <summary>
    /// Moves the given rules up by one. With visibleIds (while the list is filtered), rules move only among the
    /// visible ones and hidden rules keep their positions.
    /// </summary>
    public bool MoveUp(IEnumerable<string> ids, IEnumerable<string>? visibleIds = null) =>
        Commit(Reorder.MoveUp(_entries, ids.ToHashSet(), visibleIds?.ToHashSet()));

    public bool MoveDown(IEnumerable<string> ids, IEnumerable<string>? visibleIds = null) =>
        Commit(Reorder.MoveDown(_entries, ids.ToHashSet(), visibleIds?.ToHashSet()));

    public bool MoveToTop(IEnumerable<string> ids) => MoveTo(ids, 1);

    public bool MoveToBottom(IEnumerable<string> ids) => MoveTo(ids, int.MaxValue);

    /// <summary>Moves the given rules so that the first of them lands at position (1-based), keeping their order.</summary>
    public bool MoveTo(IEnumerable<string> ids, int position) =>
        Commit(Reorder.MoveTo(_entries, ids.ToHashSet(), position));

    /// <summary>Moves the given rules right before targetId (for drag and drop); to the end when targetId is null.</summary>
    public bool MoveBefore(IEnumerable<string> ids, string? targetId) =>
        Commit(Reorder.MoveBefore(_entries, ids.ToHashSet(), targetId));

    /// <summary>Swaps the positions of two rules.</summary>
    public bool Swap(string id1, string id2)
    {
        int a = IndexOf(id1), b = IndexOf(id2);
        if (a < 0 || b < 0 || a == b) return false;
        var next = _entries.ToList();
        (next[a], next[b]) = (next[b], next[a]);
        return Commit(next);
    }

    // ---- Internals ----

    private int Update(IEnumerable<string> ids, Func<RuleEntry, RuleEntry> change)
    {
        var set = ids.ToHashSet();
        int changed = 0;
        var next = _entries.Select(e =>
        {
            if (!set.Contains(e.Id)) return e;
            var n = change(e);
            if (!ReferenceEquals(n, e) && n != e) changed++;
            return n;
        }).ToList();
        if (changed > 0) Commit(next);
        return changed;
    }

    /// <summary>If the order or content changed, pushes the current state for undo and replaces it.</summary>
    private bool Commit(IReadOnlyList<RuleEntry> next)
    {
        if (next.Count == _entries.Count && next.Zip(_entries).All(p => p.First == p.Second)) return false;
        _undo.Push(_entries);
        _entries = next;
        return true;
    }
}
