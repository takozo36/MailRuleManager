using System.Globalization;

namespace OutlookRuleManager.Core;

public enum UiLanguage
{
    Japanese,
    English,
}

/// <summary>
/// UI language switch. Every user-facing string in the app is written as a Japanese/English pair
/// next to the code that uses it (<see cref="T"/>), instead of in resource files. With only two
/// languages this keeps both texts side by side, so a missing translation is easy to spot, and it
/// works in a single-file build without satellite assemblies.
/// </summary>
public static class Loc
{
    /// <summary>Current UI language. Defaults to the Windows display language (Japanese if "ja", otherwise English).</summary>
    public static UiLanguage Current { get; set; } = FromCulture(CultureInfo.CurrentUICulture);

    public static bool IsJapanese => Current == UiLanguage.Japanese;

    /// <summary>Picks the text for the current language.</summary>
    public static string T(string japanese, string english) => IsJapanese ? japanese : english;

    public static UiLanguage FromCulture(CultureInfo culture) =>
        culture.TwoLetterISOLanguageName == "ja" ? UiLanguage.Japanese : UiLanguage.English;
}
