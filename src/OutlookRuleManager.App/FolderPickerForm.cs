using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.App;

/// <summary>Dialog to choose a move-to folder, either from the tree or from search results.</summary>
internal sealed class FolderPickerForm : Form
{
    private readonly IReadOnlyList<FolderNode> _roots;
    private readonly TextBox _search = new() { Dock = DockStyle.Top, PlaceholderText = T("フォルダー名で検索（空白区切りで絞り込み）", "Search folders (space-separated terms narrow the results)") };
    private readonly TreeView _tree = new() { Dock = DockStyle.Fill, HideSelection = false };
    private readonly ListBox _results = new() { Dock = DockStyle.Fill, Visible = false, IntegralHeight = false };
    private readonly Button _ok = new() { Text = "OK", DialogResult = DialogResult.OK, Width = 90, Height = 30, Enabled = false };

    public FolderRef? SelectedFolder { get; private set; }

    public FolderPickerForm(IReadOnlyList<FolderNode> roots, FolderRef? current)
    {
        _roots = roots;
        Text = T("移動先フォルダーの選択", "Choose the move-to folder");
        Font = new Font(IsJapanese ? "Yu Gothic UI" : "Segoe UI", 9F);
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(560, 620);
        Padding = new Padding(12);

        var cancel = new Button { Text = T("キャンセル", "Cancel"), DialogResult = DialogResult.Cancel, Width = 100, Height = 30 };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 44, Padding = new Padding(0, 6, 0, 0) };
        buttons.Controls.AddRange([cancel, _ok]);
        var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0) };
        body.Controls.Add(_tree);
        body.Controls.Add(_results);

        Controls.Add(body);
        Controls.Add(buttons);
        Controls.Add(_search);
        AcceptButton = _ok;
        CancelButton = cancel;

        _search.TextChanged += (_, _) => ApplySearch();
        _tree.AfterSelect += (_, e) => SetSelection((e.Node?.Tag as FolderNode)?.Folder);
        _tree.NodeMouseDoubleClick += (_, e) => { if (e.Node.Tag is FolderNode) Accept(); };
        _results.SelectedIndexChanged += (_, _) => SetSelection((_results.SelectedItem as ResultItem)?.Node.Folder);
        _results.DoubleClick += (_, _) => { if (_results.SelectedItem is not null) Accept(); };
        Shown += (_, _) => _search.Focus();
        BuildTree(current); // build after subscribing to AfterSelect, so that the current folder becomes the selection
    }

    private void BuildTree(FolderRef? current)
    {
        _tree.BeginUpdate();
        foreach (var r in _roots) _tree.Nodes.Add(ToTreeNode(r));
        foreach (TreeNode n in _tree.Nodes) n.Expand();
        _tree.EndUpdate();

        if (current is null) return;
        var match = FindNode(_tree.Nodes, current);
        if (match is not null)
        {
            _tree.SelectedNode = match;
            match.EnsureVisible();
        }
    }

    private static TreeNode ToTreeNode(FolderNode f)
    {
        var node = new TreeNode(f.Name) { Tag = f };
        foreach (var c in f.Children) node.Nodes.Add(ToTreeNode(c));
        return node;
    }

    private static TreeNode? FindNode(TreeNodeCollection nodes, FolderRef target)
    {
        foreach (TreeNode n in nodes)
        {
            if (n.Tag is FolderNode f && f.Folder.EntryId == target.EntryId) return n;
            var child = FindNode(n.Nodes, target);
            if (child is not null) return child;
        }
        return null;
    }

    private void ApplySearch()
    {
        bool searching = _search.Text.Trim().Length > 0;
        _tree.Visible = !searching;
        _results.Visible = searching;
        if (!searching)
        {
            SetSelection((_tree.SelectedNode?.Tag as FolderNode)?.Folder);
            return;
        }
        _results.BeginUpdate();
        _results.Items.Clear();
        foreach (var n in FolderNode.Search(_roots, _search.Text)) _results.Items.Add(new ResultItem(n));
        _results.EndUpdate();
        if (_results.Items.Count > 0) _results.SelectedIndex = 0;
        else SetSelection(null);
    }

    private void SetSelection(FolderRef? folder)
    {
        SelectedFolder = folder;
        _ok.Enabled = folder is not null;
    }

    private void Accept()
    {
        if (SelectedFolder is null) return;
        DialogResult = DialogResult.OK;
        Close();
    }

    private sealed record ResultItem(FolderNode Node)
    {
        public override string ToString() => Node.Folder.DisplayPath;
    }
}
