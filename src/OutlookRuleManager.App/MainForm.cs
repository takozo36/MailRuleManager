using System.Diagnostics;
using OutlookRuleManager.Core;
using OutlookRuleManager.Outlook;

namespace OutlookRuleManager.App;

internal sealed partial class MainForm : Form
{
    private readonly OutlookRuleGateway _gateway = new();
    private readonly Dictionary<string, IReadOnlyList<FolderNode>> _folderCache = new();
    private readonly Dictionary<string, bool> _fileExistsCache = new(StringComparer.OrdinalIgnoreCase);

    private StoreInfo? _store;
    private RuleListEditor? _editor;
    private IReadOnlyDictionary<string, IReadOnlyList<Diagnostic>> _diagnostics = new Dictionary<string, IReadOnlyList<Diagnostic>>();
    private List<RuleEntry> _visible = [];
    private Dictionary<string, int> _positions = new();
    private bool _busy;
    private bool _suppressStoreChange;
    private bool _loadingRules;

    public MainForm()
    {
        BuildLayout();
        var version = FileVersionInfo.GetVersionInfo(Environment.ProcessPath!).FileVersion ?? "";
        if (version.EndsWith(".0.0")) version = version[..^4];
        Text = $"Outlook仕訳ルール管理 v{version}";
        Shown += async (_, _) => await LoadStoresAsync();
        FormClosing += OnFormClosing;
    }

    private IReadOnlyList<RuleEntry> Entries => _editor?.Entries ?? Array.Empty<RuleEntry>();

    // =====================================================================
    // 読み込み
    // =====================================================================

    private async Task LoadStoresAsync()
    {
        IReadOnlyList<StoreInfo> stores;
        try
        {
            SetBusy(true, "Outlook に接続しています…");
            stores = await Task.Run(_gateway.ListStores);
        }
        catch (Exception ex)
        {
            ShowError("Outlook に接続できませんでした。Outlook（クラシック）がインストールされ、起動できる状態か確認してください。", ex);
            return;
        }
        finally
        {
            SetBusy(false, "");
        }

        _suppressStoreChange = true;
        _storeCombo.Items.Clear();
        foreach (var s in stores) _storeCombo.Items.Add(s);
        _storeCombo.SelectedItem = stores.FirstOrDefault(s => s.IsDefault) ?? stores.FirstOrDefault();
        _suppressStoreChange = false;

        await LoadRulesAsync();
    }

    /// <summary>
    /// 仕訳ルールを読み込む。全件で数分かかるので、読めたものから一覧に出していく
    /// （読み込み中は閲覧・検索だけでき、編集は読み込み完了後）。
    /// </summary>
    private async Task LoadRulesAsync()
    {
        if (_storeCombo.SelectedItem is not StoreInfo store) return;
        var loaded = new List<RuleData>();
        bool pending = false;
        var progress = new Progress<LoadProgress>(p =>
        {
            loaded.Add(p.Rule);
            pending = true;
            _progress.Maximum = Math.Max(1, p.Total);
            _progress.Value = Math.Min(loaded.Count, _progress.Maximum);
            _statusMessage.Text = $"仕訳ルールを読み込んでいます… {loaded.Count} / {p.Total}（読み込んだものから表示しています）";
        });
        using var refreshTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        refreshTimer.Tick += (_, _) =>
        {
            if (!pending) return;
            pending = false;
            _editor = new RuleListEditor(loaded.ToList());
            RefreshView();
        };

        IReadOnlyList<RuleData> rules;
        var sw = Stopwatch.StartNew();
        try
        {
            _store = store;
            _editor = null;
            _loadingRules = true;
            SetBusy(true, "仕訳ルールを読み込んでいます…");
            RefreshView(keepSelection: false);
            refreshTimer.Start();
            rules = await Task.Run(() => _gateway.LoadRules(store.StoreId, progress, CancellationToken.None));
        }
        catch (Exception ex)
        {
            _editor = null;
            _statusMessage.Text = "";
            ShowError("仕訳ルールを読み込めませんでした。", ex);
            return;
        }
        finally
        {
            refreshTimer.Stop();
            _loadingRules = false;
            SetBusy(false, null);
        }

        _editor = new RuleListEditor(rules);
        _fileExistsCache.Clear();
        RefreshView();
        _statusMessage.Text = $"{rules.Count} 件を読み込みました（{sw.Elapsed.TotalSeconds:0} 秒）";
        _grid.Focus();
    }
    private async void OnStoreChanged(object? sender, EventArgs e)
    {
        if (_suppressStoreChange || _storeCombo.SelectedItem as StoreInfo == _store) return;
        if (!ConfirmDiscard())
        {
            _suppressStoreChange = true;
            _storeCombo.SelectedItem = _store;
            _suppressStoreChange = false;
            return;
        }
        await LoadRulesAsync();
    }

    private async void OnReloadClick(object? sender, EventArgs e)
    {
        if (!ConfirmDiscard()) return;
        await LoadRulesAsync();
    }

    // =====================================================================
    // 表示
    // =====================================================================

    /// <summary>診断・絞り込みをやり直して一覧を描き直す。</summary>
    private void RefreshView(bool keepSelection = true, IEnumerable<string>? select = null)
    {
        var selectIds = (select ?? (keepSelection ? SelectedIds() : [])).ToHashSet();
        string? currentId = keepSelection && _grid.CurrentCell is { RowIndex: >= 0 } cc && cc.RowIndex < _visible.Count ? _visible[cc.RowIndex].Id : null;
        int firstDisplayed = _grid.FirstDisplayedScrollingRowIndex;

        var entries = Entries;
        _diagnostics = RuleDiagnostics.Analyze(entries, FileExists);
        _positions = entries.Select((e, i) => (e.Id, i)).ToDictionary(p => p.Id, p => p.i + 1);
        var mode = (RuleFilterMode)Math.Max(0, _filterCombo.SelectedIndex);
        _visible = entries.Where(e => RuleFilter.Matches(e, _diagnostics[e.Id], _searchBox.Text, mode)).ToList();

        _grid.SuspendLayout();
        _grid.RowCount = 0;
        _grid.RowCount = _visible.Count;
        _grid.ClearSelection();
        int firstSelected = -1;
        for (int i = 0; i < _visible.Count; i++)
        {
            if (!selectIds.Contains(_visible[i].Id)) continue;
            _grid.Rows[i].Selected = true;
            if (firstSelected < 0) firstSelected = i;
        }
        if (firstDisplayed >= 0 && firstDisplayed < _visible.Count && select is null)
            _grid.FirstDisplayedScrollingRowIndex = firstDisplayed;
        int current = currentId is not null ? _visible.FindIndex(e => e.Id == currentId) : -1;
        if (select is not null || current < 0) current = firstSelected;
        if (current >= 0)
        {
            // CurrentCell を設定すると選択が 1 行に戻るので、設定後に選択を復元する
            _grid.CurrentCell = _grid.Rows[current].Cells[ColName];
            foreach (var i in Enumerable.Range(0, _visible.Count).Where(i => selectIds.Contains(_visible[i].Id)))
                _grid.Rows[i].Selected = true;
        }
        _grid.ResumeLayout();
        _grid.Invalidate();

        UpdateStatus();
        UpdateCommands();
        UpdateDetail();
    }

    private bool FileExists(string path)
    {
        if (!_fileExistsCache.TryGetValue(path, out bool exists))
            _fileExistsCache[path] = exists = File.Exists(path);
        return exists;
    }

    private void UpdateStatus()
    {
        var entries = Entries;
        int errors = entries.Count(e => RuleDiagnostics.Worst(_diagnostics[e.Id]) == Severity.Error);
        int warnings = entries.Count(e => RuleDiagnostics.Worst(_diagnostics[e.Id]) == Severity.Warning);
        _statusCount.Text = _visible.Count == entries.Count ? $"全 {entries.Count} 件" : $"表示 {_visible.Count} / 全 {entries.Count} 件";
        _statusDiag.Text = $"エラー {errors} 件・警告 {warnings} 件";
        _statusDiag.ForeColor = errors > 0 ? Palette.ErrorText : warnings > 0 ? Palette.WarningText : SystemColors.ControlText;

        int changes = _editor?.BuildPlan().ChangeCount ?? 0;
        _statusChanges.Text = changes > 0 ? $"未保存の変更 {changes} 件" : "未保存の変更なし";
        _statusChanges.ForeColor = changes > 0 ? Palette.ChangedText : SystemColors.ControlText;
        _saveButton.Text = changes > 0 ? $"Outlook へ保存 ({changes})" : "Outlook へ保存";
    }

    private void UpdateCommands()
    {
        var sel = SelectedEntries();
        bool loaded = _editor is not null && !_busy;
        bool any = loaded && sel.Count > 0;
        _btnEnable.Enabled = any && sel.Any(e => !e.Enabled);
        _btnDisable.Enabled = any && sel.Any(e => e.Enabled);
        _btnRename.Enabled = loaded && sel.Count == 1;
        _btnFolder.Enabled = any && sel.Any(e => e.HasMoveAction);
        _btnDuplicate.Enabled = any;
        _btnDelete.Enabled = any;
        _btnUp.Enabled = _btnDown.Enabled = _btnTop.Enabled = _btnBottom.Enabled = _btnMoveTo.Enabled = any;
        _btnSwap.Enabled = loaded && sel.Count == 2;
        _btnUndo.Enabled = loaded && _editor!.CanUndo;
        _btnDiscard.Enabled = loaded && _editor!.BuildPlan().HasChanges;
        _saveButton.Enabled = _btnDiscard.Enabled;
        _exportButton.Enabled = loaded;
        _reloadButton.Enabled = !_busy;
        _storeCombo.Enabled = !_busy;
    }

    private void UpdateDetail()
    {
        _detail.Clear();
        WriteDetail();
        _detail.Select(0, 0);
        // AppendText で末尾までスクロールしているので先頭に戻す。
        // その場で送ると描画前のレイアウトで上書きされることがあるため、メッセージキューに積んで後で送る
        BeginInvoke(() => SendMessage(_detail.Handle, WM_VSCROLL, SB_TOP, IntPtr.Zero));
    }

    private void WriteDetail()
    {
        var sel = SelectedEntries();
        if (sel.Count == 0)
        {
            AppendDetail("ルールを選択すると、ここに条件・処理・問題点の詳細を表示します。", Palette.MutedText);
            return;
        }
        if (sel.Count > 1)
        {
            AppendDetail($"{sel.Count} 件を選択中", null, bold: true);
            AppendDetail("");
            foreach (var e in sel.Take(200))
            {
                var worst = RuleDiagnostics.Worst(_diagnostics[e.Id]);
                AppendDetail($"{_positions[e.Id],4}  {e.Name}", worst == Severity.Error ? Palette.ErrorText : worst == Severity.Warning ? Palette.WarningText : null);
            }
            return;
        }

        var r = sel[0];
        AppendDetail($"[{_positions[r.Id]}] {r.Name}", null, bold: true);
        AppendDetail($"{(r.Enabled ? "有効" : "無効")}・{RuleText.KindText(r)}{(r.IsNew ? "・複製（未保存）" : "")}", Palette.MutedText);
        if (r.IsRenamed) AppendDetail($"元の名前: {r.Source.Name}", Palette.MutedText);

        var diags = _diagnostics[r.Id];
        if (diags.Count > 0)
        {
            AppendDetail("");
            AppendDetail("■ 問題点", null, bold: true);
            foreach (var d in diags)
                AppendDetail($"  [{RuleCsvExporter.SeverityText(d.Severity)}] {d.Message}",
                    d.Severity == Severity.Error ? Palette.ErrorText : d.Severity == Severity.Warning ? Palette.WarningText : Palette.MutedText);
        }
        AppendSection("■ 条件", r.Conditions.Select(RuleText.Describe));
        AppendSection("■ 例外", r.Exceptions.Select(RuleText.Describe));
        AppendSection("■ 処理", r.Actions.Select(RuleText.Describe));
    }

    private const int WM_VSCROLL = 0x0115;
    private const int SB_TOP = 6;

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);

    private void AppendSection(string title, IEnumerable<string> lines)
    {
        var list = lines.ToList();
        if (list.Count == 0) return;
        AppendDetail("");
        AppendDetail(title, null, bold: true);
        foreach (var l in list) AppendDetail("  ・" + l);
    }

    private void AppendDetail(string text, Color? color = null, bool bold = false)
    {
        _detail.SelectionStart = _detail.TextLength;
        _detail.SelectionColor = color ?? _detail.ForeColor;
        _detail.SelectionFont = bold ? _boldFont : _detail.Font;
        _detail.AppendText(text + Environment.NewLine);
    }

    // =====================================================================
    // 一覧（仮想モード）
    // =====================================================================

    private void OnCellValueNeeded(object? sender, DataGridViewCellValueEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _visible.Count) return;
        var r = _visible[e.RowIndex];
        e.Value = _grid.Columns[e.ColumnIndex].Name switch
        {
            ColOrder => _positions[r.Id],
            ColEnabled => r.Enabled,
            ColState => RuleDiagnostics.Worst(_diagnostics[r.Id]) switch
            {
                Severity.Error => "エラー",
                Severity.Warning => "警告",
                Severity.Info => "情報",
                _ => "",
            },
            ColChanged => r.IsNew ? "複製" : r.IsModified ? "変更" : "",
            ColName => r.Name,
            ColConditions => RuleText.Summary(r.Conditions),
            ColFolder => RuleText.MoveTarget(r.Actions),
            ColActions => RuleText.ActionSummary(r.Actions),
            ColExceptions => RuleText.Summary(r.Exceptions),
            ColKind => RuleText.KindText(r),
            _ => null,
        };
    }

    private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _visible.Count || e.CellStyle is null) return;
        var r = _visible[e.RowIndex];
        var worst = RuleDiagnostics.Worst(_diagnostics[r.Id]);
        string col = _grid.Columns[e.ColumnIndex].Name;

        if (worst == Severity.Error) e.CellStyle.BackColor = Palette.ErrorBack;
        else if (worst == Severity.Warning) e.CellStyle.BackColor = Palette.WarningBack;
        else if (r.IsNew) e.CellStyle.BackColor = Palette.NewBack;

        if (!r.Enabled) e.CellStyle.ForeColor = Palette.DisabledText;

        if (col == ColState && worst is not null)
        {
            e.CellStyle.ForeColor = worst == Severity.Error ? Palette.ErrorText : worst == Severity.Warning ? Palette.WarningText : Palette.MutedText;
            e.CellStyle.Font = _boldFont;
        }
        else if (col == ColChanged && r.IsModified)
        {
            e.CellStyle.ForeColor = Palette.ChangedText;
            e.CellStyle.Font = _boldFont;
        }
        else if (col == ColName && r.IsRenamed)
        {
            e.CellStyle.ForeColor = Palette.ChangedText;
        }
        else if (col == ColFolder && r.IsFolderChanged)
        {
            e.CellStyle.ForeColor = Palette.ChangedText;
        }
    }

    private void OnCellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (_editor is null || _busy || e.RowIndex < 0 || e.RowIndex >= _visible.Count) return;
        if (_grid.Columns[e.ColumnIndex].Name != ColEnabled) return;
        var r = _visible[e.RowIndex];
        // 選択中の行のチェックを押したら選択中すべて、それ以外ならその行だけを切り替える
        var ids = _grid.Rows[e.RowIndex].Selected ? SelectedIds() : [r.Id];
        _editor.SetEnabled(ids, !r.Enabled);
        RefreshView();
    }

    private void OnCellToolTipTextNeeded(object? sender, DataGridViewCellToolTipTextNeededEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _visible.Count) return;
        var r = _visible[e.RowIndex];
        string col = _grid.Columns[e.ColumnIndex].Name;
        if (col == ColState)
            e.ToolTipText = string.Join(Environment.NewLine, _diagnostics[r.Id].Select(d => $"[{RuleCsvExporter.SeverityText(d.Severity)}] {d.Message}"));
        else if (col is ColConditions or ColActions or ColExceptions or ColFolder or ColName)
            e.ToolTipText = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue?.ToString() ?? "";
    }

    private List<string> SelectedIds() => SelectedEntries().Select(e => e.Id).ToList();

    /// <summary>選択中のルール（実行順の昇順）。</summary>
    private List<RuleEntry> SelectedEntries()
    {
        var rows = new List<int>();
        foreach (DataGridViewRow row in _grid.SelectedRows)
            if (row.Index >= 0 && row.Index < _visible.Count) rows.Add(row.Index);
        return rows.Order().Select(i => _visible[i]).ToList();
    }

    // =====================================================================
    // 編集操作
    // =====================================================================

    private void Edit(Func<RuleListEditor, bool> action, IEnumerable<string>? select = null)
    {
        if (_editor is null || _busy) return;
        if (action(_editor)) RefreshView(select: select);
    }

    private void OnEnable(object? sender, EventArgs e) => Edit(ed => ed.SetEnabled(SelectedIds(), true) > 0);

    private void OnDisable(object? sender, EventArgs e) => Edit(ed => ed.SetEnabled(SelectedIds(), false) > 0);

    private void OnRename(object? sender, EventArgs e)
    {
        var sel = SelectedEntries();
        if (_editor is null || sel.Count != 1) return;
        var name = InputDialog.Show(this, "名前の変更", "新しいルール名:", sel[0].Name);
        if (name is null) return;
        if (name.Trim().Length == 0) { MessageBox.Show(this, "ルール名を入力してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        Edit(ed => ed.Rename(sel[0].Id, name));
    }

    private void OnDelete(object? sender, EventArgs e)
    {
        var sel = SelectedEntries();
        if (_editor is null || sel.Count == 0) return;
        string names = string.Join(Environment.NewLine, sel.Take(15).Select(x => "・" + x.Name)) + (sel.Count > 15 ? $"{Environment.NewLine}…ほか {sel.Count - 15} 件" : "");
        if (MessageBox.Show(this, $"{sel.Count} 件のルールを削除します。{Environment.NewLine}{Environment.NewLine}{names}{Environment.NewLine}{Environment.NewLine}（「Outlook へ保存」を押すまでは Outlook には反映されません）",
                "削除の確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
        // 削除後は、削除した範囲の直後にあったルールを選択しておく
        int lastRow = _visible.FindLastIndex(v => v.Id == sel[^1].Id);
        var next = _visible.Skip(lastRow + 1).FirstOrDefault(v => !sel.Contains(v)) ?? _visible.Take(lastRow).LastOrDefault(v => !sel.Contains(v));
        Edit(ed => ed.Delete(sel.Select(x => x.Id)) > 0, next is null ? [] : [next.Id]);
    }

    private void OnDuplicate(object? sender, EventArgs e)
    {
        if (_editor is null) return;
        var created = _editor.Duplicate(SelectedIds(), out var skipped);
        if (created.Count > 0) RefreshView(select: created);
        if (skipped.Count > 0)
            MessageBox.Show(this,
                "次のルールは Outlook の画面でしか設定できない条件・処理を含むため、複製できません。" + Environment.NewLine + Environment.NewLine +
                string.Join(Environment.NewLine, skipped.Select(s => $"・{s.Name}（{string.Join("、", RuleCapabilities.UncopyableParts(s.Source).Distinct())}）")),
                Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private async void OnChangeFolder(object? sender, EventArgs e)
    {
        var sel = SelectedEntries().Where(x => x.HasMoveAction).ToList();
        if (_editor is null || _store is null || sel.Count == 0) return;

        if (!_folderCache.TryGetValue(_store.StoreId, out var roots))
        {
            try
            {
                SetBusy(true, "フォルダー一覧を読み込んでいます…");
                var storeId = _store.StoreId;
                roots = await Task.Run(() => _gateway.LoadFolders(storeId));
                _folderCache[_store.StoreId] = roots;
            }
            catch (Exception ex)
            {
                ShowError("フォルダー一覧を読み込めませんでした。", ex);
                return;
            }
            finally
            {
                SetBusy(false, "");
            }
        }

        var current = sel[0].Actions.First(a => a.Type == ActionType.MoveToFolder).Folder;
        using var picker = new FolderPickerForm(roots, current);
        if (picker.ShowDialog(this) != DialogResult.OK || picker.SelectedFolder is null) return;
        var folder = picker.SelectedFolder;
        Edit(ed => { ed.SetMoveFolder(sel.Select(x => x.Id), folder); return true; });
    }

    private void OnMoveUp(object? sender, EventArgs e) => Edit(ed => ed.MoveUp(SelectedIds(), _visible.Select(v => v.Id)));

    private void OnMoveDown(object? sender, EventArgs e) => Edit(ed => ed.MoveDown(SelectedIds(), _visible.Select(v => v.Id)));

    private void OnMoveTop(object? sender, EventArgs e) => Edit(ed => ed.MoveToTop(SelectedIds()));

    private void OnMoveBottom(object? sender, EventArgs e) => Edit(ed => ed.MoveToBottom(SelectedIds()));

    private void OnMoveTo(object? sender, EventArgs e)
    {
        var sel = SelectedEntries();
        if (_editor is null || sel.Count == 0) return;
        var input = InputDialog.Show(this, "位置を指定して移動",
            $"移動先の実行順（1～{Entries.Count}）を入力してください。{(sel.Count > 1 ? $"選択した {sel.Count} 件はこの位置から順に並びます。" : "")}",
            _positions[sel[0].Id].ToString());
        if (input is null) return;
        if (!int.TryParse(RuleFilter.Normalize(input).Trim(), out int pos) || pos < 1)
        {
            MessageBox.Show(this, "1 以上の数字を入力してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        Edit(ed => ed.MoveTo(sel.Select(x => x.Id), pos));
    }

    private void OnSwap(object? sender, EventArgs e)
    {
        var sel = SelectedIds();
        if (sel.Count == 2) Edit(ed => ed.Swap(sel[0], sel[1]));
    }

    private void OnUndo(object? sender, EventArgs e) => Edit(ed => ed.Undo());

    private void OnDiscard(object? sender, EventArgs e)
    {
        if (_editor is null) return;
        if (MessageBox.Show(this, "未保存の変更をすべて破棄して、読み込んだ時点の状態に戻します。よろしいですか？", "変更の破棄",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
        _editor = new RuleListEditor(_editor.Original);
        RefreshView();
    }

    // =====================================================================
    // 保存・出力
    // =====================================================================

    private async void OnSave(object? sender, EventArgs e)
    {
        if (_editor is null || _store is null || _busy) return;
        var plan = _editor.BuildPlan();
        if (!plan.HasChanges) return;

        var problems = plan.Validate();
        if (problems.Count > 0)
        {
            MessageBox.Show(this, "次の問題があるため保存できません。" + Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, problems),
                Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string backupPath = BackupPath(_store);
        var lines = plan.Describe().ToList();
        int errorRules = plan.FinalOrder.Count(x => RuleDiagnostics.Worst(_diagnostics[x.Id]) == Severity.Error);
        lines.Add("");
        lines.Add($"保存前の状態を CSV に残します: {backupPath}");
        lines.Add("（CSV は確認用の記録です。元に戻せる完全なバックアップが必要なら、Outlook の");
        lines.Add("  「ルールと通知」→「オプション」→「ルールのエクスポート」を先に実行してください）");
        if (errorRules > 0)
        {
            lines.Add("");
            lines.Add($"※ エラーのあるルールが {errorRules} 件残っています。Outlook が保存を拒否する場合は、");
            lines.Add("   先にそのルールの移動先を直すか、削除してから保存してください。");
        }
        if (!ConfirmDialog.Show(this, "Outlook へ保存", $"次の {plan.ChangeCount} 件の変更を Outlook に保存します。{Environment.NewLine}保存中は Outlook の「ルールと通知」画面を開かないでください。", lines, "保存する"))
            return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);
            var originalEntries = _editor.Original.Select(RuleEntry.FromSource).ToList();
            RuleCsvExporter.WriteFile(backupPath, originalEntries, RuleDiagnostics.Analyze(originalEntries));
        }
        catch (Exception ex)
        {
            if (MessageBox.Show(this, $"保存前の CSV を書き出せませんでした（{ex.Message}）。CSV なしで保存を続けますか？", Text,
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
        }

        var progress = new Progress<string>(m => _statusMessage.Text = m);
        var storeId = _store.StoreId;
        try
        {
            SetBusy(true, "Outlook に保存しています…");
            _progress.Style = ProgressBarStyle.Marquee;
            await Task.Run(() => _gateway.Apply(storeId, plan, progress));
        }
        catch (RulesChangedException ex)
        {
            SetBusy(false, "保存を中止しました");
            if (MessageBox.Show(this, ex.Message + Environment.NewLine + Environment.NewLine +
                    "Outlook 側でルールが変更されたため、保存を中止しました（Outlook には何も保存していません）。" + Environment.NewLine +
                    "最新の状態を読み込み直しますか？（このアプリでの未保存の変更は失われます）",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                await LoadRulesAsync();
            return;
        }
        catch (Exception ex)
        {
            SetBusy(false, "保存に失敗しました");
            ShowError("Outlook への保存に失敗しました。", ex);
            return;
        }
        finally
        {
            _progress.Style = ProgressBarStyle.Continuous;
            SetBusy(false, null);
        }

        // 全件の読み込み直しは数分かかるので、保存した内容から一覧を組み立て直す
        _editor = new RuleListEditor(plan.ToSavedSnapshot());
        RefreshView();
        _statusMessage.Text = $"{plan.ChangeCount} 件の変更を Outlook に保存しました（{DateTime.Now:HH:mm}）";
        MessageBox.Show(this, $"{plan.ChangeCount} 件の変更を Outlook に保存しました。" + Environment.NewLine + Environment.NewLine +
            "Outlook 側の最新の状態を確かめたいときは「読み込み直す」を押してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string BackupPath(StoreInfo store)
    {
        string safe = string.Concat(store.DisplayName.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Outlook仕訳ルール管理", "保存前の記録", $"{safe}_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
    }

    private void OnExport(object? sender, EventArgs e)
    {
        if (_editor is null) return;
        bool filtered = _visible.Count != Entries.Count;
        if (filtered)
        {
            var answer = MessageBox.Show(this, $"表示中の {_visible.Count} 件だけを書き出しますか？{Environment.NewLine}「いいえ」なら全 {Entries.Count} 件を書き出します。",
                "CSV 出力", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (answer == DialogResult.Cancel) return;
            filtered = answer == DialogResult.Yes;
        }
        using var dlg = new SaveFileDialog
        {
            Filter = "CSV ファイル (*.csv)|*.csv",
            FileName = $"仕訳ルール_{DateTime.Now:yyyyMMdd_HHmm}.csv",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            // 実行順は全体での位置を出したいので、全件で書いてから絞り込む
            var rows = filtered ? _visible : Entries.ToList();
            using var writer = new StreamWriter(dlg.FileName, false, new System.Text.UTF8Encoding(true));
            RuleCsvExporter.Write(writer, Entries, _diagnostics, rows.Select(r => r.Id).ToHashSet());
            _statusMessage.Text = $"CSV を書き出しました: {dlg.FileName}";
        }
        catch (Exception ex)
        {
            ShowError("CSV を書き出せませんでした。", ex);
        }
    }

    // =====================================================================
    // 共通
    // =====================================================================

    private bool ConfirmDiscard()
    {
        if (_editor is null || !_editor.BuildPlan().HasChanges) return true;
        return MessageBox.Show(this, "未保存の変更があります。破棄してよろしいですか？", Text,
            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK;
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_busy)
        {
            MessageBox.Show(this, "処理中です。終わるまでお待ちください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            e.Cancel = true;
            return;
        }
        if (!ConfirmDiscard()) e.Cancel = true;
        else _gateway.Dispose();
    }

    private void SetBusy(bool busy, string? message)
    {
        _busy = busy;
        _progress.Visible = busy;
        if (!busy) _progress.Value = 0;
        if (message is not null) _statusMessage.Text = message;
        // ルールの読み込み中は、読めた分の閲覧・検索はできるようにしておく
        UseWaitCursor = busy && !_loadingRules;
        _grid.Enabled = !busy || _loadingRules;
        UpdateCommands();
    }

    private void ShowError(string message, Exception ex)
    {
        MessageBox.Show(this, message + Environment.NewLine + Environment.NewLine + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        // 処理中（読み込み中を含む）は検索以外のショートカットを受け付けない
        if (_busy && keyData is not (Keys.Control | Keys.F) and not Keys.Escape)
            return base.ProcessCmdKey(ref msg, keyData);
        bool gridFocused = _grid.Focused;
        switch (keyData)
        {
            case Keys.Control | Keys.F:
                _searchBox.Focus();
                _searchBox.SelectAll();
                return true;
            case Keys.Escape when _searchBox.Focused && _searchBox.Text.Length > 0:
                _searchBox.Clear();
                return true;
            case Keys.Control | Keys.Z when !_searchBox.Focused:
                OnUndo(null, EventArgs.Empty);
                return true;
            case Keys.Control | Keys.S:
                OnSave(null, EventArgs.Empty);
                return true;
            case Keys.Control | Keys.D:
                OnDuplicate(null, EventArgs.Empty);
                return true;
            case Keys.F2 when gridFocused:
                OnRename(null, EventArgs.Empty);
                return true;
            case Keys.Delete when gridFocused:
                OnDelete(null, EventArgs.Empty);
                return true;
            case Keys.Space when gridFocused:
                var sel = SelectedEntries();
                if (sel.Count > 0) Edit(ed => ed.SetEnabled(sel.Select(x => x.Id), !sel[0].Enabled) > 0);
                return true;
            case Keys.Alt | Keys.Up:
                OnMoveUp(null, EventArgs.Empty);
                return true;
            case Keys.Alt | Keys.Down:
                OnMoveDown(null, EventArgs.Empty);
                return true;
            case Keys.F5:
                OnReloadClick(null, EventArgs.Empty);
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
