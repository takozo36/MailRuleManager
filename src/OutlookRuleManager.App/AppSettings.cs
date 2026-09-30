using System.Globalization;
using OutlookRuleManager.Core;

namespace OutlookRuleManager.App;

/// <summary>
/// Per-user settings stored in %LOCALAPPDATA%\MailRuleManager\settings.ini.
/// Currently only the display language ("language=auto|ja|en").
/// </summary>
internal static class AppSettings
{
    public static string DataDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailRuleManager");

    // Data folder used up to v0.1.0 (before the app was renamed)
    private static string LegacyDataDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookRuleManager");

    private static string FilePath => Path.Combine(DataDirectory, "settings.ini");

    /// <summary>Chosen display language; null means "follow the Windows display language".</summary>
    public static UiLanguage? Language { get; set; }

    public static UiLanguage ResolveLanguage() => Language ?? Loc.FromCulture(CultureInfo.CurrentUICulture);

    public static void Load()
    {
        MigrateLegacyDirectory();
        try
        {
            if (!File.Exists(FilePath)) return;
            foreach (var line in File.ReadAllLines(FilePath))
            {
                var parts = line.Split('=', 2, StringSplitOptions.TrimEntries);
                if (parts.Length == 2 && parts[0] == "language")
                    Language = parts[1] switch { "ja" => UiLanguage.Japanese, "en" => UiLanguage.English, _ => null };
            }
        }
        catch
        {
            // A broken settings file must not stop the app; fall back to defaults
        }
    }

    /// <summary>
    /// Moves the old data folder (settings.ini and Backups) to the new name, only when the new one does not exist yet.
    /// </summary>
    private static void MigrateLegacyDirectory()
    {
        try
        {
            if (Directory.Exists(LegacyDataDirectory) && !Directory.Exists(DataDirectory))
                Directory.Move(LegacyDataDirectory, DataDirectory);
        }
        catch
        {
            // Leave the old folder as is; the app starts with default settings
        }
    }

    public static void Save()
    {
        try
        {
            Directory.CreateDirectory(DataDirectory);
            string value = Language switch { UiLanguage.Japanese => "ja", UiLanguage.English => "en", _ => "auto" };
            File.WriteAllText(FilePath, $"language={value}{Environment.NewLine}");
        }
        catch
        {
            // Not being able to remember the language is not worth an error dialog
        }
    }
}
