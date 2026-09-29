using OutlookRuleManager.Core;

namespace OutlookRuleManager.App;

internal static class Palette
{
    public static readonly Color ErrorBack = Color.FromArgb(253, 231, 233);
    public static readonly Color WarningBack = Color.FromArgb(255, 246, 219);
    public static readonly Color NewBack = Color.FromArgb(232, 245, 233);
    public static readonly Color ErrorText = Color.FromArgb(196, 43, 28);
    public static readonly Color WarningText = Color.FromArgb(157, 93, 0);
    public static readonly Color ChangedText = Color.FromArgb(0, 95, 184);
    public static readonly Color DisabledText = Color.FromArgb(150, 150, 150);
    public static readonly Color MutedText = Color.FromArgb(96, 96, 96);
    /// <summary>CSV 出力ボタン（Excel の緑）。</summary>
    public static readonly Color ExportAccent = Color.FromArgb(16, 124, 65);
    /// <summary>Outlook へ保存ボタン（Outlook の青）。</summary>
    public static readonly Color SaveAccent = Color.FromArgb(15, 108, 189);
}

/// <summary>
/// Tag に色を持つボタンだけ、その色で塗りつぶして白文字で描く。
/// 標準の描き方だとマウスを乗せたときに水色の強調で上書きされ、白文字が読めなくなるため。
/// </summary>
internal sealed class AccentRenderer : ToolStripProfessionalRenderer
{

    public AccentRenderer() { RoundedEdges = false; }

    protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Tag is not Color accent) { base.OnRenderButtonBackground(e); return; }
        // 押せないときも何のボタンか分かるよう、灰色ではなく同じ色の薄い色にする
        Color back = !e.Item.Enabled ? Blend(accent, 0.18)
            : e.Item.Pressed ? ControlPaint.Dark(accent, 0.1f)
            : e.Item.Selected ? ControlPaint.Light(accent, 0.25f)
            : accent;
        var rect = new Rectangle(Point.Empty, e.Item.Size);
        rect.Inflate(-1, -1);
        using var brush = new SolidBrush(back);
        using var path = RoundedRect(rect, 4);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        e.Graphics.FillPath(brush, path);
    }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    {
        if (e.Item.Tag is Color accent) e.TextColor = e.Item.Enabled ? Color.White : Blend(accent, 0.55);
        base.OnRenderItemText(e);
    }

    /// <summary>白に accent を ratio の割合で混ぜた色。</summary>
    private static Color Blend(Color accent, double ratio) => Color.FromArgb(
        (int)(255 + (accent.R - 255) * ratio),
        (int)(255 + (accent.G - 255) * ratio),
        (int)(255 + (accent.B - 255) * ratio));

    private static System.Drawing.Drawing2D.GraphicsPath RoundedRect(Rectangle r, int radius)
    {
        int d = radius * 2;
        var p = new System.Drawing.Drawing2D.GraphicsPath();
        p.AddArc(r.Left, r.Top, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

internal sealed partial class MainForm
{
    private const string ColOrder = "order";
    private const string ColEnabled = "enabled";
    private const string ColState = "state";
    private const string ColChanged = "changed";
    private const string ColName = "name";
    private const string ColConditions = "conditions";
    private const string ColFolder = "folder";
    private const string ColActions = "actions";
    private const string ColExceptions = "exceptions";
    private const string ColKind = "kind";

    private Font _boldFont = null!;
    private ToolStripComboBox _storeCombo = null!;
    private ToolStripButton _reloadButton = null!;
    private ToolStripTextBox _searchBox = null!;
    private ToolStripComboBox _filterCombo = null!;
    private ToolStripButton _exportButton = null!;
    private ToolStripButton _saveButton = null!;
    private ToolStripButton _btnEnable = null!, _btnDisable = null!, _btnRename = null!, _btnFolder = null!, _btnDuplicate = null!, _btnDelete = null!;
    private ToolStripButton _btnUp = null!, _btnDown = null!, _btnTop = null!, _btnBottom = null!, _btnMoveTo = null!, _btnSwap = null!;
    private ToolStripButton _btnUndo = null!, _btnDiscard = null!;
    private DataGridView _grid = null!;
    private RichTextBox _detail = null!;
    private ToolStripStatusLabel _statusCount = null!, _statusDiag = null!, _statusChanges = null!, _statusMessage = null!;
    private ToolStripProgressBar _progress = null!;

    private void BuildLayout()
    {
        Font = new Font("Yu Gothic UI", 9F);
        _boldFont = new Font(Font, FontStyle.Bold);
        ClientSize = new Size(1400, 820);
        MinimumSize = new Size(900, 500);
        StartPosition = FormStartPosition.CenterScreen;
        KeyPreview = true;

        // ---- 1 段目: アカウント・検索・保存 ----
        var top = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 4, 6, 4), ImageScalingSize = new Size(16, 16) };
        _storeCombo = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = 240 };
        _storeCombo.SelectedIndexChanged += OnStoreChanged;
        _reloadButton = TextButton("読み込み直す", "Outlook から仕訳ルールを読み込み直します (F5)", OnReloadClick);
        _searchBox = new ToolStripTextBox { AutoSize = false, Width = 320, ToolTipText = "名前・差出人・件名・移動先などを検索（空白区切りで絞り込み） Ctrl+F" };
        _searchBox.TextBox.PlaceholderText = "検索（名前・差出人・件名・移動先…）";
        var searchTimer = new System.Windows.Forms.Timer { Interval = 250 };
        searchTimer.Tick += (_, _) => { searchTimer.Stop(); RefreshView(); };
        _searchBox.TextChanged += (_, _) => { searchTimer.Stop(); searchTimer.Start(); };
        _filterCombo = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = 170 };
        // RuleFilterMode の並びと同じ順で並べる
        _filterCombo.Items.AddRange(["すべて表示", "エラー・警告のあるもの", "エラーのあるもの", "無効のもの", "未保存の変更があるもの"]);
        _filterCombo.SelectedIndex = 0;
        _filterCombo.SelectedIndexChanged += (_, _) => RefreshView();
        _exportButton = TextButton("CSV 出力", "一覧を CSV（Excel で開ける形式）に書き出します", OnExport);
        _saveButton = TextButton("Outlook へ保存", "変更を Outlook に保存します (Ctrl+S)", OnSave);
        _saveButton.Alignment = ToolStripItemAlignment.Right;
        AccentButton(_exportButton, Palette.ExportAccent);
        AccentButton(_saveButton, Palette.SaveAccent);
        top.Renderer = new AccentRenderer();
        top.Items.AddRange([
            new ToolStripLabel("アカウント:"), _storeCombo, _reloadButton, new ToolStripSeparator(),
            new ToolStripLabel("検索:"), _searchBox, _filterCombo, new ToolStripSeparator(),
            _exportButton, _saveButton,
        ]);

        // ---- 2 段目: 編集操作 ----
        var edit = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 2, 6, 4) };
        _btnEnable = TextButton("有効にする", "選択したルールを有効にします（Space で切り替え）", OnEnable);
        _btnDisable = TextButton("無効にする", "選択したルールを無効にします（Space で切り替え）", OnDisable);
        _btnRename = TextButton("名前の変更", "名前を変更します (F2)", OnRename);
        _btnFolder = TextButton("移動先を変更", "移動先フォルダーを変更します。エラーのルールもこれで直せます", OnChangeFolder);
        _btnDuplicate = TextButton("複製", "選択したルールを複製して直後に追加します (Ctrl+D)", OnDuplicate);
        _btnDelete = TextButton("削除", "選択したルールを削除します (Delete)", OnDelete);
        _btnUp = TextButton("▲ 上へ", "1 つ上へ (Alt+↑)。検索中は表示中のルールの間で動かします", OnMoveUp);
        _btnDown = TextButton("▼ 下へ", "1 つ下へ (Alt+↓)。検索中は表示中のルールの間で動かします", OnMoveDown);
        _btnTop = TextButton("先頭へ", "実行順の先頭へ移動します", OnMoveTop);
        _btnBottom = TextButton("末尾へ", "実行順の末尾へ移動します", OnMoveBottom);
        _btnMoveTo = TextButton("位置を指定…", "実行順の番号を指定して移動します", OnMoveTo);
        _btnSwap = TextButton("入れ替え", "選択した 2 件の位置を入れ替えます", OnSwap);
        _btnUndo = TextButton("元に戻す", "直前の操作を取り消します (Ctrl+Z)", OnUndo);
        _btnDiscard = TextButton("変更をすべて破棄", "未保存の変更をすべて破棄します", OnDiscard);
        edit.Items.AddRange([
            _btnEnable, _btnDisable, _btnRename, _btnFolder, _btnDuplicate, _btnDelete, new ToolStripSeparator(),
            _btnUp, _btnDown, _btnTop, _btnBottom, _btnMoveTo, _btnSwap, new ToolStripSeparator(),
            _btnUndo, _btnDiscard,
        ]);

        // ---- 一覧 ----
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            VirtualMode = true,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AllowUserToOrderColumns = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            RowHeadersVisible = false,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 30,
            ShowCellToolTips = true,
            AllowDrop = true,
        };
        _grid.RowTemplate.Height = 26;
        _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 228, 247);
        _grid.DefaultCellStyle.SelectionForeColor = SystemColors.ControlText;
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(243, 243, 243);
        _grid.ColumnHeadersDefaultCellStyle.Font = _boldFont;
        typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .SetValue(_grid, true);

        AddColumn(new DataGridViewTextBoxColumn(), ColOrder, "順番", 54, DataGridViewContentAlignment.MiddleRight);
        AddColumn(new DataGridViewCheckBoxColumn(), ColEnabled, "有効", 44, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColState, "状態", 56, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColChanged, "変更", 48, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColName, "名前", 280);
        AddColumn(new DataGridViewTextBoxColumn(), ColConditions, "条件", 340);
        AddColumn(new DataGridViewTextBoxColumn(), ColFolder, "移動先フォルダー", 260);
        AddColumn(new DataGridViewTextBoxColumn(), ColActions, "処理", 200);
        AddColumn(new DataGridViewTextBoxColumn(), ColExceptions, "例外", 150);
        AddColumn(new DataGridViewTextBoxColumn(), ColKind, "種類", 130);

        _grid.CellValueNeeded += OnCellValueNeeded;
        _grid.CellFormatting += OnCellFormatting;
        _grid.CellContentClick += OnCellContentClick;
        _grid.CellToolTipTextNeeded += OnCellToolTipTextNeeded;
        _grid.SelectionChanged += (_, _) => { UpdateCommands(); UpdateDetail(); };
        _grid.CellDoubleClick += (_, e) => { if (e.RowIndex >= 0 && _grid.Columns[e.ColumnIndex].Name == ColFolder) OnChangeFolder(null, EventArgs.Empty); };
        _grid.ContextMenuStrip = BuildContextMenu();
        _grid.CellMouseDown += OnGridCellMouseDown;
        SetUpDragDrop();

        _detail = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = SystemColors.Window,
            Font = new Font("Yu Gothic UI", 9.5F),
            DetectUrls = false,
        };
        var detailPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 6, 10, 6), BackColor = SystemColors.Window };
        detailPanel.Controls.Add(_detail);

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, FixedPanel = FixedPanel.Panel2, SplitterWidth = 6 };
        split.Panel1.Controls.Add(_grid);
        split.Panel2.Controls.Add(detailPanel);

        // ---- ステータスバー ----
        var status = new StatusStrip();
        _statusCount = new ToolStripStatusLabel { BorderSides = ToolStripStatusLabelBorderSides.Right, Padding = new Padding(0, 0, 8, 0) };
        _statusDiag = new ToolStripStatusLabel { BorderSides = ToolStripStatusLabelBorderSides.Right, Padding = new Padding(0, 0, 8, 0), IsLink = false };
        _statusChanges = new ToolStripStatusLabel { BorderSides = ToolStripStatusLabelBorderSides.Right, Padding = new Padding(0, 0, 8, 0) };
        _progress = new ToolStripProgressBar { Visible = false, Width = 180 };
        _statusMessage = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        status.Items.AddRange([_statusCount, _statusDiag, _statusChanges, _progress, _statusMessage]);

        Controls.Add(split);
        Controls.Add(edit);
        Controls.Add(top);
        Controls.Add(status);
        Load += (_, _) => split.SplitterDistance = Math.Max(200, split.Height - 260);

        UpdateCommands();
    }

    private void AddColumn(DataGridViewColumn col, string name, string header, int width, DataGridViewContentAlignment align = DataGridViewContentAlignment.MiddleLeft)
    {
        col.Name = name;
        col.HeaderText = header;
        col.Width = width;
        col.SortMode = DataGridViewColumnSortMode.NotSortable;
        col.DefaultCellStyle.Alignment = align;
        _grid.Columns.Add(col);
    }

    /// <summary>色付きボタンにする（塗りは AccentRenderer が Tag の色で描く）。</summary>
    private void AccentButton(ToolStripButton b, Color accent)
    {
        b.Tag = accent;
        b.Font = _boldFont;
        b.Margin = new Padding(4, 1, 4, 1);
        b.Padding = new Padding(12, 3, 12, 3);
    }

    private static ToolStripButton TextButton(string text, string tip, EventHandler onClick)
    {
        var b = new ToolStripButton(text) { DisplayStyle = ToolStripItemDisplayStyle.Text, ToolTipText = tip, Padding = new Padding(4, 0, 4, 0) };
        b.Click += onClick;
        return b;
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();
        static ToolStripMenuItem Item(string text, EventHandler h, string? keyText = null) =>
            new(text, null, h) { ShortcutKeyDisplayString = keyText };
        var enable = Item("有効にする", OnEnable);
        var disable = Item("無効にする", OnDisable);
        var rename = Item("名前の変更…", OnRename, keyText: "F2");
        var folder = Item("移動先を変更…", OnChangeFolder);
        var dup = Item("複製", OnDuplicate, keyText: "Ctrl+D");
        var del = Item("削除", OnDelete, keyText: "Del");
        var up = Item("上へ", OnMoveUp, keyText: "Alt+↑");
        var down = Item("下へ", OnMoveDown, keyText: "Alt+↓");
        var topItem = Item("先頭へ", OnMoveTop);
        var bottom = Item("末尾へ", OnMoveBottom);
        var moveTo = Item("位置を指定して移動…", OnMoveTo);
        var swap = Item("入れ替え（2 件選択時）", OnSwap);
        menu.Items.AddRange([enable, disable, rename, folder, dup, del, new ToolStripSeparator(), up, down, topItem, bottom, moveTo, swap]);
        menu.Opening += (_, _) =>
        {
            enable.Enabled = _btnEnable.Enabled; disable.Enabled = _btnDisable.Enabled; rename.Enabled = _btnRename.Enabled;
            folder.Enabled = _btnFolder.Enabled; dup.Enabled = _btnDuplicate.Enabled; del.Enabled = _btnDelete.Enabled;
            up.Enabled = down.Enabled = topItem.Enabled = bottom.Enabled = moveTo.Enabled = _btnUp.Enabled;
            swap.Enabled = _btnSwap.Enabled;
        };
        return menu;
    }

    /// <summary>右クリックした行が未選択なら、その行だけを選択してからメニューを出す。</summary>
    private void OnGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right || e.RowIndex < 0) return;
        if (_grid.Rows[e.RowIndex].Selected) return;
        _grid.ClearSelection();
        _grid.CurrentCell = _grid.Rows[e.RowIndex].Cells[ColName];
        _grid.Rows[e.RowIndex].Selected = true;
    }

    // =====================================================================
    // ドラッグ＆ドロップで並べ替え
    // =====================================================================

    private Rectangle _dragBox = Rectangle.Empty;
    private int _dropLineY = -1;

    private void SetUpDragDrop()
    {
        _grid.MouseDown += (_, e) =>
        {
            var hit = _grid.HitTest(e.X, e.Y);
            bool onRow = e.Button == MouseButtons.Left && hit.RowIndex >= 0 && _grid.Columns[hit.ColumnIndex].Name != ColEnabled;
            _dragBox = onRow
                ? new Rectangle(new Point(e.X - SystemInformation.DragSize.Width / 2, e.Y - SystemInformation.DragSize.Height / 2), SystemInformation.DragSize)
                : Rectangle.Empty;
        };
        _grid.MouseUp += (_, _) => _dragBox = Rectangle.Empty;
        _grid.MouseMove += (_, e) =>
        {
            if (e.Button != MouseButtons.Left || _dragBox == Rectangle.Empty || _dragBox.Contains(e.X, e.Y) || _busy || _editor is null) return;
            _dragBox = Rectangle.Empty;
            var ids = SelectedIds();
            if (ids.Count > 0) _grid.DoDragDrop(new DragPayload(ids), DragDropEffects.Move);
        };
        _grid.DragOver += (_, e) =>
        {
            if (e.Data?.GetData(typeof(DragPayload)) is null) { e.Effect = DragDropEffects.None; return; }
            e.Effect = DragDropEffects.Move;
            var p = _grid.PointToClient(new Point(e.X, e.Y));
            AutoScrollWhileDragging(p);
            int y = DropTarget(p).lineY;
            if (y != _dropLineY) { _dropLineY = y; _grid.Invalidate(); }
        };
        _grid.DragLeave += (_, _) => { _dropLineY = -1; _grid.Invalidate(); };
        _grid.DragDrop += (_, e) =>
        {
            _dropLineY = -1;
            _grid.Invalidate();
            if (e.Data?.GetData(typeof(DragPayload)) is not DragPayload payload) return;
            var (targetId, _) = DropTarget(_grid.PointToClient(new Point(e.X, e.Y)));
            Edit(ed => ed.MoveBefore(payload.Ids, targetId), payload.Ids);
        };
        _grid.Paint += (_, e) =>
        {
            if (_dropLineY < 0) return;
            using var pen = new Pen(Palette.ChangedText, 3);
            e.Graphics.DrawLine(pen, 0, _dropLineY, _grid.Width, _dropLineY);
        };
    }

    /// <summary>
    /// ドロップ位置から「どのルールの直前に入れるか」と、挿入線を描く y 座標を求める。
    /// 行の上半分なら その行の直前、下半分なら その行の直後（＝全体の並びで次のルールの直前）。
    /// </summary>
    private (string? targetId, int lineY) DropTarget(Point p)
    {
        var hit = _grid.HitTest(p.X, p.Y);
        if (hit.RowIndex < 0 || _visible.Count == 0)
        {
            // 最終行より下なら末尾へ
            int lastRow = _grid.RowCount - 1;
            var rect = lastRow >= 0 ? _grid.GetRowDisplayRectangle(lastRow, false) : Rectangle.Empty;
            return (null, rect.IsEmpty ? -1 : rect.Bottom);
        }
        var r = _grid.GetRowDisplayRectangle(hit.RowIndex, false);
        bool upper = p.Y < r.Top + r.Height / 2;
        var row = _visible[hit.RowIndex];
        if (upper) return (row.Id, r.Top);
        int pos = _editor!.IndexOf(row.Id);
        string? next = pos + 1 < Entries.Count ? Entries[pos + 1].Id : null;
        return (next, r.Bottom);
    }

    private void AutoScrollWhileDragging(Point p)
    {
        int margin = _grid.RowTemplate.Height;
        int first = _grid.FirstDisplayedScrollingRowIndex;
        if (first < 0) return;
        if (p.Y < _grid.ColumnHeadersHeight + margin && first > 0)
            _grid.FirstDisplayedScrollingRowIndex = first - 1;
        else if (p.Y > _grid.ClientSize.Height - margin && first < _grid.RowCount - 1)
            _grid.FirstDisplayedScrollingRowIndex = first + 1;
    }

    private sealed record DragPayload(List<string> Ids);
}
