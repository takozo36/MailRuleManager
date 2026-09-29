using System.Globalization;
using OutlookRuleManager.Core;

namespace OutlookRuleManager.App;

/// <summary>
/// Per-user settings stored in %LOCALAPPDATA%\OutlookRuleManager\settings.ini.
/// Currently only the display language ("language=auto|ja|en").
/// </summary>
internal static class AppSettings
{
    public static string DataDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookRuleManager");

    private static string FilePath => Path.Combine(DataDirectory, "settings.ini");

    /// <summary>Chosen display language; null means "follow the Windows display language".</summary>
    public static UiLanguage? Language { get; set; }

    public static UiLanguage ResolveLanguage() => Language ?? Loc.FromCulture(CultureInfo.CurrentUICulture);

    public static void Load()
    {
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
