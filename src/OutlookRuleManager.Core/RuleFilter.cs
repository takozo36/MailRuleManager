namespace OutlookRuleManager.Core;

public enum RuleFilterMode
{
    All,
    /// <summary>Rules with an error or a warning.</summary>
    Problems,
    ErrorsOnly,
    Disabled,
    /// <summary>Rules with unsaved changes (new, renamed, enabled flag or move-to folder changed).</summary>
    Changed,
}

public static class RuleFilter
{
    /// <summary>
    /// Whether the rule matches the search terms (space-separated, all must match) and the filter mode.
    /// Terms are searched in the name, conditions, exceptions, actions and move-to folder, ignoring the
    /// differences listed in <see cref="Normalize"/>.
    /// </summary>
    public static bool Matches(RuleEntry entry, IReadOnlyList<Diagnostic> diagnostics, string? query, RuleFilterMode mode)
    {
        bool modeOk = mode switch
        {
            RuleFilterMode.Problems => diagnostics.Any(d => d.Severity >= Severity.Warning),
            RuleFilterMode.ErrorsOnly => diagnostics.Any(d => d.Severity == Severity.Error),
            RuleFilterMode.Disabled => !entry.Enabled,
            RuleFilterMode.Changed => entry.IsModified,
            _ => true,
        };
        if (!modeOk) return false;

        var terms = SplitTerms(query);
        if (terms.Count == 0) return true;
        string haystack = Normalize(RuleText.SearchText(entry));
        return terms.All(t => haystack.Contains(t, StringComparison.Ordinal));
    }

    public static IReadOnlyList<string> SplitTerms(string? query) =>
        string.IsNullOrWhiteSpace(query)
            ? Array.Empty<string>()
            : Normalize(query).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

    /// <summary>
    /// Normalizes text for search: full-width alphanumerics to half-width and half-width katakana to full-width (NFKC),
    /// hiragana to katakana, small kana to normal kana (so "キャンプ" and "キヤンプ" match), and lower case.
    /// </summary>
    public static string Normalize(string s)
    {
        var chars = s.Normalize(System.Text.NormalizationForm.FormKC).ToLowerInvariant().ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            char c = chars[i];
            if (c >= 'ぁ' && c <= 'ゖ') c = (char)(c + 0x60); // hiragana -> katakana
            chars[i] = c switch
            {
                'ァ' => 'ア', 'ィ' => 'イ', 'ゥ' => 'ウ', 'ェ' => 'エ', 'ォ' => 'オ',
                'ャ' => 'ヤ', 'ュ' => 'ユ', 'ョ' => 'ヨ', 'ッ' => 'ツ', 'ヮ' => 'ワ',
                'ヵ' => 'カ', 'ヶ' => 'ケ',
                _ => c,
            };
        }
        return new string(chars);
    }
}
