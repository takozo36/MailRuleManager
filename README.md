# Mail Rule Manager for Classic Outlook

Created: 2026-09-29
Updated: 2026-09-30 (renamed the app, added the Download section; previously updated 2026-09-29 for English / Japanese UI and documents)
AI agent: Claude Code
Model: Claude Opus 5.5 / claude-opus-5-5

**English** | [日本語](README.ja.md)

A Windows app to search, bulk-edit, duplicate and reorder rules in **classic Outlook**, and to find
broken rules and rules that never run because an earlier rule stops processing.
The UI is available in English and Japanese (chosen automatically from the Windows display language; can be switched from the menu).

> [!WARNING]
> This app modifies classic Outlook rules.
> Before you use "Save to Outlook" for the first time, back up your rules to an .rwz file in Outlook:
> "Rules and Alerts" → "Options" → "Export Rules".

> [!NOTE]
> The new Outlook is not supported. Classic Outlook must be installed and able to start.
> This is **v0.1.0 (preview)**. See "Requirements and tested environments" below.

## Features

| Area | What you can do |
|---|---|
| View | All rules in one table (order, enabled, status, name, conditions, move-to folder, actions, exceptions, type), with details for the selected rule |
| Search | Names, senders, subjects, folders and more (space-separated terms are ANDed; full-/half-width, hiragana/katakana and small/normal kana are treated as the same) |
| Filter | All / errors and warnings / errors only / disabled / unsaved changes |
| Bulk edit | Enable, disable, delete or change the move-to folder of several rules at once |
| Duplicate | Copies the selected rules right below the originals (named "… (copy)") |
| Reorder | Up, down, to top, to bottom, to a given position, swap two rules, drag and drop |
| Undo | Undo one step at a time (Ctrl+Z) / discard all changes |
| Export | CSV (opens in Excel) |
| Language | English / Japanese ("Language" at the top right; defaults to the Windows display language) |

All edits stay inside the app and are **applied to Outlook only when you click "Save to Outlook"**.

### Problems it finds

| Severity | Description |
|---|---|
| Error | The move-to or copy-to folder no longer exists (deleted or moved) / the rule has no actions / a forward recipient cannot be resolved / the rule cannot be read |
| Warning | The rule never runs because an earlier rule handles the same messages and stops processing (also reported when only some senders overlap) / the sound file is missing / a recipient in a condition cannot be resolved |
| Info | Another rule has exactly the same conditions / the rule contains items only Outlook can set (cannot be duplicated by this app) |

A missing move-to folder can be fixed with "Change folder".

## Requirements and tested environments

| Item | Status |
|---|---|
| Windows 11 x64 + classic Outlook x64 (Microsoft 365, 16.0, Japanese) | Tested |
| POP / IMAP .pst store (about 340 rules) | Loading, diagnostics and editing tested |
| "Save to Outlook" | Every step up to the final save (applying changes, duplicating and verifying by reading back, reordering, deleting) was tested on a real mailbox. **The final save itself is still being verified by the author** |
| English Outlook | Not tested (the English UI of this app is tested) |
| Classic Outlook x86 | Not tested |
| Exchange / Microsoft 365 mailbox | Not tested |
| Multiple Outlook profiles | Not tested (the default profile is used) |
| Windows ARM64 | Not tested |
| New Outlook | Not supported (it has no object model) |

## Download

Download the prebuilt `MailRuleManager.exe` from [Releases](https://github.com/takozo36/MailRuleManager/releases).
No installation is needed; put it in any folder and run it.

- Requires the [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  If it is not installed, Windows shows where to download it when the app starts.
- The exe is not code-signed, so Windows may show "Windows protected your PC" the first time you run it.
  Click "More info" → "Run anyway". If you want to check the file, compare it with the SHA-256 in the release notes
  (`Get-FileHash .\MailRuleManager.exe` in PowerShell), or build it yourself (see "Build" below).
- The current version is a preview. Before you use "Save to Outlook" for the first time, export your rules to an .rwz file.

## Usage

1. Start classic Outlook, then start `MailRuleManager.exe`. It starts loading the rules of the default account.
2. Loading takes **about 2 seconds for 340 rules**. The app reads the data in which Outlook stores all rules at once, checks the count, names and enabled flags against Outlook's rules, then shows them.
   If that data cannot be read (for example, it contains a condition of an unknown type), the app automatically falls back to reading rule by rule (about 0.3 s per rule). The status bar shows which method was used.
3. Edit, then click "Save to Outlook". Review the changes in the confirmation dialog before saving.
4. After saving, the list is rebuilt from what was saved instead of reloading. To check the current state in Outlook, click "Reload" (F5).

### Keyboard shortcuts

| Key | Action |
|---|---|
| Ctrl+F | Go to the search box |
| Space | Toggle enabled for the selected rules |
| F2 | Rename |
| Ctrl+D | Duplicate |
| Delete | Delete |
| Alt+↑ / Alt+↓ | Move up / down |
| Ctrl+Z | Undo |
| Ctrl+S | Save to Outlook |
| F5 | Reload |

## Safety measures and cautions

- Before saving, the state at load time is written as CSV
  (`%LOCALAPPDATA%\MailRuleManager\Backups\`). This is a record for reference and cannot be used to restore.
- **Make a full backup in Outlook** (.rwz export; see the warning above).
- When saving, if Outlook's rules changed since they were loaded (count and names are compared), saving is cancelled.
- Duplicated rules are read back before saving and compared with the originals; if they differ, saving is cancelled.
- If a step fails while saving, nothing is saved to Outlook (Outlook keeps changes only when `Rules.Save()` is called).
- Do not open Outlook's "Rules and Alerts" dialog while saving.
- Outlook may refuse to save while rules with errors remain. In that case, fix their move-to folder or delete them first.

The language setting is stored in `%LOCALAPPDATA%\MailRuleManager\settings.ini`.
If the folder of the previous name (`%LOCALAPPDATA%\OutlookRuleManager\`) exists, it is moved to the new name at startup.

## Known limitations

- Classic Outlook only; the new Outlook is not supported
- Fast loading relies on analysis of an undocumented format (see the technical notes). If it cannot be used, loading falls back to rule by rule (0.2–0.35 s per rule)
- Editable: name, enabled flag, move-to folder, order, deletion and duplication. Condition contents (such as senders) must be edited in Outlook
- Changing the copy-to folder ("copy to folder" action) is not supported
- Rules containing items only Outlook can set (run a script, start an application, ...) cannot be duplicated
- Drag and drop moves one row at a time (use the buttons to move several rows)
- Condition and action names are this app's own wording and may differ from the labels in Outlook

## Build

```bash
dotnet build OutlookRuleManager.slnx
```

```bash
dotnet test tests/OutlookRuleManager.Core.Tests
```

Publish a single exe (distribute only the exe; the .pdb files are not needed):

```bash
dotnet publish src/OutlookRuleManager.App -c Release
```

Output: `src/OutlookRuleManager.App/bin/Release/net8.0-windows/win-x64/publish/MailRuleManager.exe`
(about 400 KB; requires the .NET 8 Desktop Runtime. Add `--self-contained true` to bundle the runtime.)

## Project structure

```
OutlookRuleManager.slnx
src/
  OutlookRuleManager.Core/      Logic independent of the UI and Outlook (editing, reordering, diagnostics, search, CSV, English/Japanese texts)
  OutlookRuleManager.Outlook/   Loading and saving through the Outlook object model (COM, late binding)
  OutlookRuleManager.App/       WinForms UI
tests/
  OutlookRuleManager.Core.Tests/  Unit tests for Core (xUnit)
docs/
  TECHNICAL_NOTES.md / TECHNICAL_NOTES.ja.md   What was measured about the Outlook object model
```

UI texts are written as Japanese/English pairs where they are used (`Loc.T("日本語", "English")`).
Source code comments are in English.

The app itself uses no external libraries (Outlook is called with late binding; the interop assembly is not referenced).
The rules data parser is a C# implementation based on the format analysis of open-source projects (MIT License).
See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) for the sources and the packages used by the tests.

## Disclaimer

This software modifies classic Outlook rules. Before saving changes, export your existing Outlook rules to an `.rwz` file.
The author is not responsible for loss, corruption or unintended modification of Outlook rules, messages, folders or profiles.
Use at your own risk.

This project is not affiliated with or endorsed by Microsoft. Microsoft and Outlook are trademarks of the Microsoft group of companies.

## License

[MIT License](LICENSE)
