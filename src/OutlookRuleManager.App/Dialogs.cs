using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.App;

/// <summary>Single-line text input dialog (for renaming and entering a position).</summary>
internal sealed class InputDialog : Form
{
    private readonly TextBox _text = new() { Dock = DockStyle.Top };

    private InputDialog(string title, string prompt, string value)
    {
        Text = title;
        Font = new Font(IsJapanese ? "Yu Gothic UI" : "Segoe UI", 9F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(520, 110);
        Padding = new Padding(12);

        var label = new Label { Text = prompt, Dock = DockStyle.Top, AutoSize = false, Height = 24 };
        _text.Text = value;
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
        var cancel = new Button { Text = T("キャンセル", "Cancel"), DialogResult = DialogResult.Cancel, Width = 90 };
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

/// <summary>Shows a multi-line description and asks for OK / Cancel (used to confirm saving).</summary>
internal sealed class ConfirmDialog : Form
{
    private ConfirmDialog(string title, string heading, IEnumerable<string> lines, string okText)
    {
        Text = title;
        Font = new Font(IsJapanese ? "Yu Gothic UI" : "Segoe UI", 9F);
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
        var cancel = new Button { Text = T("キャンセル", "Cancel"), DialogResult = DialogResult.Cancel, Width = 100, Height = 30 };
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
