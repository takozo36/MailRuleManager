namespace OutlookRuleManager.Core;

/// <summary>移動先を選ぶためのフォルダーツリーの 1 ノード。</summary>
public sealed record FolderNode(string Name, FolderRef Folder, IReadOnlyList<FolderNode> Children)
{
    /// <summary>自分と子孫を上から順に列挙する。</summary>
    public IEnumerable<FolderNode> Descendants()
    {
        yield return this;
        foreach (var c in Children)
            foreach (var d in c.Descendants())
                yield return d;
    }

    /// <summary>検索語（空白区切りで AND）にパスが一致するフォルダーを返す。</summary>
    public static IEnumerable<FolderNode> Search(IEnumerable<FolderNode> roots, string? query)
    {
        var terms = RuleFilter.SplitTerms(query);
        return roots.SelectMany(r => r.Descendants())
            .Where(n => terms.All(t => RuleFilter.Normalize(n.Folder.DisplayPath).Contains(t, StringComparison.Ordinal)));
    }
}
