namespace OutlookRuleManager.Core;

/// <summary>One node of the folder tree used to choose a move-to folder.</summary>
public sealed record FolderNode(string Name, FolderRef Folder, IReadOnlyList<FolderNode> Children)
{
    /// <summary>Enumerates this node and all descendants, top-down.</summary>
    public IEnumerable<FolderNode> Descendants()
    {
        yield return this;
        foreach (var c in Children)
            foreach (var d in c.Descendants())
                yield return d;
    }

    /// <summary>Folders whose path matches the search terms (space-separated, all must match).</summary>
    public static IEnumerable<FolderNode> Search(IEnumerable<FolderNode> roots, string? query)
    {
        var terms = RuleFilter.SplitTerms(query);
        return roots.SelectMany(r => r.Descendants())
            .Where(n => terms.All(t => RuleFilter.Normalize(n.Folder.DisplayPath).Contains(t, StringComparison.Ordinal)));
    }
}
