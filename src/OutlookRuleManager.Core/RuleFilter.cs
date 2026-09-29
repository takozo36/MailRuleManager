namespace OutlookRuleManager.Core;

public enum RuleFilterMode
{
    All,
    /// <summary>エラーまたは警告があるもの。</summary>
    Problems,
    ErrorsOnly,
    Disabled,
    /// <summary>未保存の変更があるもの（新規・名前・有効/無効・移動先）。</summary>
    Changed,
}

public static class RuleFilter
{
    /// <summary>
    /// 検索語（空白区切りで AND）と表示モードに一致するか。
    /// 検索語は名前・条件・例外・処理・移動先を対象に、大文字小文字と全角半角の英数字を区別せずに探す。
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
    /// 表記ゆれをそろえる: 全角英数字→半角、半角カナ→全角（NFKC）、ひらがな→カタカナ、
    /// 小さいカナ→大きいカナ（「キャンプ」と「キヤンプ」を同じに扱う）、英字は小文字。
    /// </summary>
    public static string Normalize(string s)
    {
        var chars = s.Normalize(System.Text.NormalizationForm.FormKC).ToLowerInvariant().ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            char c = chars[i];
            if (c >= 'ぁ' && c <= 'ゖ') c = (char)(c + 0x60); // ひらがな → カタカナ
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