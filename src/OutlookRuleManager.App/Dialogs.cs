namespace OutlookRuleManager.App;

/// <summary>1 行の文字入力ダイアログ（名前の変更・位置の指定用）。</summary>
internal sealed class InputDialog : Form
{
    private readonly TextBox _text = new() { Dock = DockStyle.Top };

    private InputDialog(string title, string prompt, string value)
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(520, 110);
        Padding = new Padding(12);

        var label = new Label { Text = prompt, Dock = DockStyle.Top, AutoSize = false, Height = 24 };
        _text.Text = value;
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
        var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Width = 90 };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 36 };
        buttons.Controls.AddRange([cancel, ok]);

        Controls.Add(buttons);
        Controls.Add(_text);
        Controls.Add(label);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public static string? Show(IWin32Window owner, string title, string prompt, string value)
    {
        using var dlg = new InputDialog(title, prompt, value);
        dlg.Shown += (_, _) => { dlg._text.SelectAll(); dlg._text.Focus(); };
        return dlg.ShowDialog(owner) == DialogResult.OK ? dlg._text.Text : null;
    }
}

/// <summary>複数行の説明を見せて OK / キャンセルを選ばせるダイアログ（保存前の確認用）。</summary>
internal sealed class ConfirmDialog : Form
{
    private ConfirmDialog(string title, string heading, IEnumerable<string> lines, string okText)
    {
        Text = title;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(720, 480);
        Padding = new Padding(12);

        var label = new Label { Text = heading, Dock = DockStyle.Top, AutoSize = false, Height = 64 };
        var box = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Dock = DockStyle.Fill,
            Text = string.Join(Environment.NewLine, lines),
            BackColor = SystemColors.Window,
        };
        var ok = new Button { Text = okText, DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(120, 30) };
        var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Width = 100, Height = 30 };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 44, Padding = new Padding(0, 6, 0, 0) };
        buttons.Controls.AddRange([cancel, ok]);

        Controls.Add(box);
        Controls.Add(buttons);
        Controls.Add(label);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public static bool Show(IWin32Window owner, string title, string heading, IEnumerable<string> lines, string okText)
    {
        using var dlg = new ConfirmDialog(title, heading, lines, okText);
        return dlg.ShowDialog(owner) == DialogResult.OK;
    }
}
