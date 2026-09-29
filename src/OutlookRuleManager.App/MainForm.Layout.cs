using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Loc;

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
    /// <summary>Export CSV button (Excel green).</summary>
    public static readonly Color ExportAccent = Color.FromArgb(16, 124, 65);
    /// <summary>Save to Outlook button (Outlook blue).</summary>
    public static readonly Color SaveAccent = Color.FromArgb(15, 108, 189);
}

/// <summary>
/// Paints only the buttons whose Tag holds a color: filled with that color and with white text.
/// With the default renderer, the light-blue hover highlight would paint over them and make white text unreadable.
/// </summary>
internal sealed class AccentRenderer : ToolStripProfessionalRenderer
{
    public AccentRenderer() { RoundedEdges = false; }

    protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Tag is not Color accent) { base.OnRenderButtonBackground(e); return; }
        // When disabled, use a light tint of the same color instead of gray so the button stays recognizable
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

    /// <summary>White mixed with accent at the given ratio.</summary>
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
    private ToolStripLabel _accountLabel = null!, _searchLabel = null!;
    private ToolStripComboBox _storeCombo = null!;
    private ToolStripButton _reloadButton = null!;
    private ToolStripTextBox _searchBox = null!;
    private ToolStripComboBox _filterCombo = null!;
    private ToolStripButton _exportButton = null!;
    private ToolStripButton _saveButton = null!;
    private ToolStripDropDownButton _languageButton = null!;
    private ToolStripMenuItem _langAuto = null!, _langJapanese = null!, _langEnglish = null!;
    private ToolStripButton _btnEnable = null!, _btnDisable = null!, _btnRename = null!, _btnFolder = null!, _btnDuplicate = null!, _btnDelete = null!;
    private ToolStripButton _btnUp = null!, _btnDown = null!, _btnTop = null!, _btnBottom = null!, _btnMoveTo = null!, _btnSwap = null!;
    private ToolStripButton _btnUndo = null!, _btnDiscard = null!;
    private ToolStripMenuItem _miEnable = null!, _miDisable = null!, _miRename = null!, _miFolder = null!, _miDuplicate = null!, _miDelete = null!;
    private ToolStripMenuItem _miUp = null!, _miDown = null!, _miTop = null!, _miBottom = null!, _miMoveTo = null!, _miSwap = null!;
    private DataGridView _grid = null!;
    private RichTextBox _detail = null!;
    private ToolStripStatusLabel _statusCount = null!, _statusDiag = null!, _statusChanges = null!, _statusMessage = null!;
    private ToolStripProgressBar _progress = null!;
    private bool _applyingLanguage;

    /// <summary>Creates the controls. All texts are set by <see cref="ApplyLanguage"/>.</summary>
    private void BuildLayout()
    {
        ClientSize = new Size(1400, 820);
        MinimumSize = new Size(900, 500);
        StartPosition = FormStartPosition.CenterScreen;
        KeyPreview = true;

        // ---- Row 1: account, search, save ----
        var top = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 4, 6, 4), ImageScalingSize = new Size(16, 16) };
        _accountLabel = new ToolStripLabel();
        _storeCombo = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = 240 };
        _storeCombo.SelectedIndexChanged += OnStoreChanged;
        _reloadButton = TextButton(OnReloadClick);
        _searchLabel = new ToolStripLabel();
        _searchBox = new ToolStripTextBox { AutoSize = false, Width = 320 };
        var searchTimer = new System.Windows.Forms.Timer { Interval = 250 };
        searchTimer.Tick += (_, _) => { searchTimer.Stop(); RefreshView(); };
        _searchBox.TextChanged += (_, _) => { searchTimer.Stop(); searchTimer.Start(); };
        _filterCombo = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = false, Width = 190 };
        _filterCombo.SelectedIndexChanged += (_, _) => { if (!_applyingLanguage) RefreshView(); };
        _exportButton = TextButton(OnExport);
        _saveButton = TextButton(OnSave);
        _saveButton.Alignment = ToolStripItemAlignment.Right;
        AccentButton(_exportButton, Palette.ExportAccent);
        AccentButton(_saveButton, Palette.SaveAccent);
        _langAuto = new ToolStripMenuItem("", null, (_, _) => OnLanguageSelected(null));
        _langJapanese = new ToolStripMenuItem("日本語", null, (_, _) => OnLanguageSelected(UiLanguage.Japanese));
        _langEnglish = new ToolStripMenuItem("English", null, (_, _) => OnLanguageSelected(UiLanguage.English));
        _languageButton = new ToolStripDropDownButton { Alignment = ToolStripItemAlignment.Right, DisplayStyle = ToolStripItemDisplayStyle.Text };
        _languageButton.DropDownItems.AddRange([_langAuto, new ToolStripSeparator(), _langJapanese, _langEnglish]);
        top.Renderer = new AccentRenderer();
        // Right-aligned items are laid out from the right edge in the order they are added
        top.Items.AddRange([
            _accountLabel, _storeCombo, _reloadButton, new ToolStripSeparator(),
            _searchLabel, _searchBox, _filterCombo, new ToolStripSeparator(),
            _exportButton, _languageButton, _saveButton,
        ]);

        // ---- Row 2: editing ----
        var edit = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 2, 6, 4) };
        _btnEnable = TextButton(OnEnable);
        _btnDisable = TextButton(OnDisable);
        _btnRename = TextButton(OnRename);
        _btnFolder = TextButton(OnChangeFolder);
        _btnDuplicate = TextButton(OnDuplicate);
        _btnDelete = TextButton(OnDelete);
        _btnUp = TextButton(OnMoveUp);
        _btnDown = TextButton(OnMoveDown);
        _btnTop = TextButton(OnMoveTop);
        _btnBottom = TextButton(OnMoveBottom);
        _btnMoveTo = TextButton(OnMoveTo);
        _btnSwap = TextButton(OnSwap);
        _btnUndo = TextButton(OnUndo);
        _btnDiscard = TextButton(OnDiscard);
        edit.Items.AddRange([
            _btnEnable, _btnDisable, _btnRename, _btnFolder, _btnDuplicate, _btnDelete, new ToolStripSeparator(),
            _btnUp, _btnDown, _btnTop, _btnBottom, _btnMoveTo, _btnSwap, new ToolStripSeparator(),
            _btnUndo, _btnDiscard,
        ]);

        // ---- Rule list ----
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
        typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .SetValue(_grid, true);

        AddColumn(new DataGridViewTextBoxColumn(), ColOrder, 54, DataGridViewContentAlignment.MiddleRight);
        AddColumn(new DataGridViewCheckBoxColumn(), ColEnabled, 56, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColState, 64, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColChanged, 56, DataGridViewContentAlignment.MiddleCenter);
        AddColumn(new DataGridViewTextBoxColumn(), ColName, 280);
        AddColumn(new DataGridViewTextBoxColumn(), ColConditions, 340);
        AddColumn(new DataGridViewTextBoxColumn(), ColFolder, 260);
        AddColumn(new DataGridViewTextBoxColumn(), ColActions, 200);
        AddColumn(new DataGridViewTextBoxColumn(), ColExceptions, 150);
        AddColumn(new DataGridViewTextBoxColumn(), ColKind, 140);

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
            DetectUrls = false,
        };
        var detailPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 6, 10, 6), BackColor = SystemColors.Window };
        detailPanel.Controls.Add(_detail);

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, FixedPanel = FixedPanel.Panel2, SplitterWidth = 6 };
        split.Panel1.Controls.Add(_grid);
        split.Panel2.Controls.Add(detailPanel);

        // ---- Status bar ----
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
    }

    /// <summary>Font for the current language (Yu Gothic UI for Japanese, Segoe UI for English).</summary>
    private static Font UiFont(float size, FontStyle style = FontStyle.Regular) =>
        new(IsJapanese ? "Yu Gothic UI" : "Segoe UI", size, style);

    /// <summary>Sets every text and font of the window for the current language. Also called when the language is switched.</summary>
    private void ApplyLanguage()
    {
        _applyingLanguage = true;
        SuspendLayout();

        Font = UiFont(9F);
        _boldFont = new Font(Font, FontStyle.Bold);
        _grid.ColumnHeadersDefaultCellStyle.Font = _boldFont;
        _exportButton.Font = _saveButton.Font = _boldFont;
        _detail.Font = UiFont(9.5F);
        Text = $"{AppName} v{Version}";

        _accountLabel.Text = T("アカウント:", "Account:");
        SetText(_reloadButton, T("読み込み直す", "Reload"), T("Outlook から仕訳ルールを読み込み直します (F5)", "Reload the rules from Outlook (F5)"));
        _searchLabel.Text = T("検索:", "Search:");
        _searchBox.ToolTipText = T("名前・差出人・件名・移動先などを検索（空白区切りで絞り込み） Ctrl+F",
            "Search names, senders, subjects, folders and more (space-separated terms narrow the results) Ctrl+F");
        _searchBox.TextBox.PlaceholderText = T("検索（名前・差出人・件名・移動先…）", "Search (name, sender, subject, folder…)");

        // Same order as RuleFilterMode
        int filter = Math.Max(0, _filterCombo.SelectedIndex);
        _filterCombo.Items.Clear();
        _filterCombo.Items.AddRange([
            T("すべて表示", "Show all"),
            T("エラー・警告のあるもの", "Errors and warnings"),
            T("エラーのあるもの", "Errors only"),
            T("無効のもの", "Disabled rules"),
            T("未保存の変更があるもの", "Unsaved changes"),
        ]);
        _filterCombo.SelectedIndex = filter;

        SetText(_exportButton, T("CSV 出力", "Export CSV"), T("一覧を CSV（Excel で開ける形式）に書き出します", "Export the list as CSV (opens in Excel)"));
        _saveButton.ToolTipText = T("変更を Outlook に保存します (Ctrl+S)", "Save the changes to Outlook (Ctrl+S)");
        _languageButton.Text = T("表示言語", "Language");
        _languageButton.ToolTipText = T("画面の言語を切り替えます", "Change the display language");
        _langAuto.Text = T("自動（Windows の設定に合わせる）", "Automatic (Windows setting)");
        _langAuto.Checked = AppSettings.Language is null;
        _langJapanese.Checked = AppSettings.Language == UiLanguage.Japanese;
        _langEnglish.Checked = AppSettings.Language == UiLanguage.English;

        SetText(_btnEnable, T("有効にする", "Enable"), T("選択したルールを有効にします（Space で切り替え）", "Enable the selected rules (Space toggles)"));
        SetText(_btnDisable, T("無効にする", "Disable"), T("選択したルールを無効にします（Space で切り替え）", "Disable the selected rules (Space toggles)"));
        SetText(_btnRename, T("名前の変更", "Rename"), T("名前を変更します (F2)", "Rename the rule (F2)"));
        SetText(_btnFolder, T("移動先を変更", "Change folder"), T("移動先フォルダーを変更します。エラーのルールもこれで直せます",
            "Change the move-to folder. This also fixes rules whose folder is missing"));
        SetText(_btnDuplicate, T("複製", "Duplicate"), T("選択したルールを複製して直後に追加します (Ctrl+D)", "Duplicate the selected rules right below them (Ctrl+D)"));
        SetText(_btnDelete, T("削除", "Delete"), T("選択したルールを削除します (Delete)", "Delete the selected rules (Delete)"));
        SetText(_btnUp, T("▲ 上へ", "▲ Up"), T("1 つ上へ (Alt+↑)。検索中は表示中のルールの間で動かします", "Move up (Alt+↑). While searching, moves among the rules shown"));
        SetText(_btnDown, T("▼ 下へ", "▼ Down"), T("1 つ下へ (Alt+↓)。検索中は表示中のルールの間で動かします", "Move down (Alt+↓). While searching, moves among the rules shown"));
        SetText(_btnTop, T("先頭へ", "To top"), T("実行順の先頭へ移動します", "Move to the top of the execution order"));
        SetText(_btnBottom, T("末尾へ", "To bottom"), T("実行順の末尾へ移動します", "Move to the bottom of the execution order"));
        SetText(_btnMoveTo, T("位置を指定…", "Move to…"), T("実行順の番号を指定して移動します", "Move to a given position"));
        SetText(_btnSwap, T("入れ替え", "Swap"), T("選択した 2 件の位置を入れ替えます", "Swap the positions of the two selected rules"));
        SetText(_btnUndo, T("元に戻す", "Undo"), T("直前の操作を取り消します (Ctrl+Z)", "Undo the last operation (Ctrl+Z)"));
        SetText(_btnDiscard, T("変更をすべて破棄", "Discard all changes"), T("未保存の変更をすべて破棄します", "Discard all unsaved changes"));

        _grid.Columns[ColOrder].HeaderText = T("順番", "Order");
        _grid.Columns[ColEnabled].HeaderText = T("有効", "On");
        _grid.Columns[ColState].HeaderText = T("状態", "Status");
        _grid.Columns[ColChanged].HeaderText = T("変更", "Edit");
        _grid.Columns[ColName].HeaderText = T("名前", "Name");
        _grid.Columns[ColConditions].HeaderText = T("条件", "Conditions");
        _grid.Columns[ColFolder].HeaderText = T("移動先フォルダー", "Move to folder");
        _grid.Columns[ColActions].HeaderText = T("処理", "Actions");
        _grid.Columns[ColExceptions].HeaderText = T("例外", "Exceptions");
        _grid.Columns[ColKind].HeaderText = T("種類", "Type");

        _miEnable.Text = T("有効にする", "Enable");
        _miDisable.Text = T("無効にする", "Disable");
        _miRename.Text = T("名前の変更…", "Rename…");
        _miFolder.Text = T("移動先を変更…", "Change folder…");
        _miDuplicate.Text = T("複製", "Duplicate");
        _miDelete.Text = T("削除", "Delete");
        _miUp.Text = T("上へ", "Move up");
        _miDown.Text = T("下へ", "Move down");
        _miTop.Text = T("先頭へ", "Move to top");
        _miBottom.Text = T("末尾へ", "Move to bottom");
        _miMoveTo.Text = T("位置を指定して移動…", "Move to position…");
        _miSwap.Text = T("入れ替え（2 件選択時）", "Swap (two rules selected)");

        ResumeLayout();
        _applyingLanguage = false;
        UpdateStatus();
        UpdateCommands();
        UpdateDetail();
    }

    private static void SetText(ToolStripItem item, string text, string tip)
    {
        item.Text = text;
        item.ToolTipText = tip;
    }

    private void AddColumn(DataGridViewColumn col, string name, int width, DataGridViewContentAlignment align = DataGridViewContentAlignment.MiddleLeft)
    {
        col.Name = name;
        col.Width = width;
        col.SortMode = DataGridViewColumnSortMode.NotSortable;
        col.DefaultCellStyle.Alignment = align;
        _grid.Columns.Add(col);
    }

    /// <summary>Makes a colored button (AccentRenderer fills it with the color in Tag).</summary>
    private static void AccentButton(ToolStripButton b, Color accent)
    {
        b.Tag = accent;
        b.Margin = new Padding(4, 1, 4, 1);
        b.Padding = new Padding(12, 3, 12, 3);
    }

    private static ToolStripButton TextButton(EventHandler onClick)
    {
        var b = new ToolStripButton { DisplayStyle = ToolStripItemDisplayStyle.Text, Padding = new Padding(4, 0, 4, 0) };
        b.Click += onClick;
        return b;
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();
        static ToolStripMenuItem Item(EventHandler h, string? keyText = null) =>
            new("", null, h) { ShortcutKeyDisplayString = keyText };
        _miEnable = Item(OnEnable);
        _miDisable = Item(OnDisable);
        _miRename = Item(OnRename, keyText: "F2");
        _miFolder = Item(OnChangeFolder);
        _miDuplicate = Item(OnDuplicate, keyText: "Ctrl+D");
        _miDelete = Item(OnDelete, keyText: "Del");
        _miUp = Item(OnMoveUp, keyText: "Alt+↑");
        _miDown = Item(OnMoveDown, keyText: "Alt+↓");
        _miTop = Item(OnMoveTop);
        _miBottom = Item(OnMoveBottom);
        _miMoveTo = Item(OnMoveTo);
        _miSwap = Item(OnSwap);
        menu.Items.AddRange([_miEnable, _miDisable, _miRename, _miFolder, _miDuplicate, _miDelete, new ToolStripSeparator(),
            _miUp, _miDown, _miTop, _miBottom, _miMoveTo, _miSwap]);
        menu.Opening += (_, _) =>
        {
            _miEnable.Enabled = _btnEnable.Enabled; _miDisable.Enabled = _btnDisable.Enabled; _miRename.Enabled = _btnRename.Enabled;
            _miFolder.Enabled = _btnFolder.Enabled; _miDuplicate.Enabled = _btnDuplicate.Enabled; _miDelete.Enabled = _btnDelete.Enabled;
            _miUp.Enabled = _miDown.Enabled = _miTop.Enabled = _miBottom.Enabled = _miMoveTo.Enabled = _btnUp.Enabled;
            _miSwap.Enabled = _btnSwap.Enabled;
        };
        return menu;
    }

    /// <summary>If the right-clicked row is not selected, selects only that row before showing the menu.</summary>
    private void OnGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right || e.RowIndex < 0) return;
        if (_grid.Rows[e.RowIndex].Selected) return;
        _grid.ClearSelection();
        _grid.CurrentCell = _grid.Rows[e.RowIndex].Cells[ColName];
        _grid.Rows[e.RowIndex].Selected = true;
    }

    // =====================================================================
    // Reordering by drag and drop
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
    /// From the drop point, works out which rule to insert before, and the y coordinate of the insertion line.
    /// Upper half of a row: before that row. Lower half: right after it (= before the next rule in the whole list).
    /// </summary>
    private (string? targetId, int lineY) DropTarget(Point p)
    {
        var hit = _grid.HitTest(p.X, p.Y);
        if (hit.RowIndex < 0 || _visible.Count == 0)
        {
            // Below the last row: to the end
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
