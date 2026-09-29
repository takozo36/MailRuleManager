namespace OutlookRuleManager.Core;

/// <summary>
/// 読み込んだルール一覧に対する編集（有効/無効・名前・移動先・削除・複製・並べ替え）をメモリ上で行う。
/// Outlook へは BuildPlan() の結果をまとめて反映する。すべての編集は Undo() で 1 手ずつ戻せる。
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

    /// <summary>現在の並び順どおりのルール一覧（先頭が実行順 1）。</summary>
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

    // ---- 内容の変更 ----

    public int SetEnabled(IEnumerable<string> ids, bool enabled) =>
        Update(ids, e => e.Enabled == enabled ? e : e with { Enabled = enabled });

    public bool Rename(string id, string newName)
    {
        newName = newName.Trim();
        if (newName.Length == 0) throw new ArgumentException("ルール名が空です。", nameof(newName));
        return Update([id], e => e.Name == newName ? e : e with { Name = newName }) > 0;
    }

    /// <summary>移動先フォルダーを変更する。移動の処理を持たないルールは対象外。</summary>
    public int SetMoveFolder(IEnumerable<string> ids, FolderRef folder) =>
        Update(ids, e =>
        {
            if (!e.HasMoveAction) return e;
            // 元の移動先に戻したときは「変更なし」に戻す
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
    /// 選択したルールを複製し、それぞれ元のルールの直後に挿入する。
    /// 複製できない（Outlook の画面でしか設定できない項目を含む）ルールは飛ばし、その一覧を skipped に返す。
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
                Name = e.Name + " のコピー",
            };
            next.Add(copy);
            created.Add(copy.Id);
        }
        skipped = skip;
        if (created.Count > 0) Commit(next);
        return created;
    }

    // ---- 並べ替え ----

    /// <summary>
    /// 選択したルールを 1 つ上へ。visibleIds を渡すと、表示中（検索で絞り込み中）のルールの間だけで動かし、
    /// 非表示のルールの位置は変えない。
    /// </summary>
    public bool MoveUp(IEnumerable<string> ids, IEnumerable<string>? visibleIds = null) =>
        Commit(Reorder.MoveUp(_entries, ids.ToHashSet(), visibleIds?.ToHashSet()));

    public bool MoveDown(IEnumerable<string> ids, IEnumerable<string>? visibleIds = null) =>
        Commit(Reorder.MoveDown(_entries, ids.ToHashSet(), visibleIds?.ToHashSet()));

    public bool MoveToTop(IEnumerable<string> ids) => MoveTo(ids, 1);

    public bool MoveToBottom(IEnumerable<string> ids) => MoveTo(ids, int.MaxValue);

    /// <summary>選択したルールを、先頭が position 番目（1 始まり）になるよう移動する。順序は保つ。</summary>
    public bool MoveTo(IEnumerable<string> ids, int position) =>
        Commit(Reorder.MoveTo(_entries, ids.ToHashSet(), position));

    /// <summary>選択したルールを targetId のルールの直前へ移動する（ドラッグ＆ドロップ用）。targetId が null なら末尾へ。</summary>
    public bool MoveBefore(IEnumerable<string> ids, string? targetId) =>
        Commit(Reorder.MoveBefore(_entries, ids.ToHashSet(), targetId));

    /// <summary>2 つのルールの位置を入れ替える。</summary>
    public bool Swap(string id1, string id2)
    {
        int a = IndexOf(id1), b = IndexOf(id2);
        if (a < 0 || b < 0 || a == b) return false;
        var next = _entries.ToList();
        (next[a], next[b]) = (next[b], next[a]);
        return Commit(next);
    }

    // ---- 内部 ----

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

    /// <summary>並びか内容が変わっていれば現在の状態を Undo 用に積んで差し替える。</summary>
    private bool Commit(IReadOnlyList<RuleEntry> next)
    {
        if (next.Count == _entries.Count && next.Zip(_entries).All(p => p.First == p.Second)) return false;
        _undo.Push(_entries);
        _entries = next;
        return true;
    }
}
