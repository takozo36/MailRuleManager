using System.Text;

namespace OutlookRuleManager.Core;

/// <summary>ルール一覧を Excel で開ける CSV（UTF-8 BOM 付き）に書き出す。onlyIds を渡すとその行だけ書く（実行順は全体での位置）。</summary>
public static class RuleCsvExporter
{
    public static readonly string[] Header =
        ["実行順", "有効", "名前", "種類", "条件", "例外", "処理", "移動先フォルダー", "診断"];

    public static void Write(
        TextWriter writer,
        IReadOnlyList<RuleEntry> entries,
        IReadOnlyDictionary<string, IReadOnlyList<Diagnostic>> diagnostics,
        IReadOnlySet<string>? onlyIds = null)
    {
        writer.WriteLine(string.Join(",", Header.Select(Escape)));
        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            if (onlyIds is not null && !onlyIds.Contains(e.Id)) continue;
            var diags = diagnostics.TryGetValue(e.Id, out var d) ? d : Array.Empty<Diagnostic>();
            string[] row =
            [
                (i + 1).ToString(),
                e.Enabled ? "有効" : "無効",
                e.Name,
                RuleText.KindText(e),
                RuleText.Summary(e.Conditions),
                RuleText.Summary(e.Exceptions),
                RuleText.ActionSummary(e.Actions),
                RuleText.MoveTarget(e.Actions),
                string.Join(" / ", diags.Select(x => $"[{SeverityText(x.Severity)}] {x.Message}")),
            ];
            writer.WriteLine(string.Join(",", row.Select(Escape)));
        }
    }

    public static void WriteFile(
        string path,
        IReadOnlyList<RuleEntry> entries,
        IReadOnlyDictionary<string, IReadOnlyList<Diagnostic>> diagnostics)
    {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
        Write(writer, entries, diagnostics);
    }

    public static string SeverityText(Severity s) => s switch
    {
        Severity.Error => "エラー",
        Severity.Warning => "警告",
        _ => "情報",
    };

    internal static string Escape(string value)
    {
        if (value.IndexOfAny([',', '"', '\n', '\r']) < 0) return value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
