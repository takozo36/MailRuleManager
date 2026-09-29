using System.Text;
using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.Core;

/// <summary>
/// Writes the rule list as CSV (UTF-8 with BOM, so Excel opens it correctly) in the current UI language.
/// With onlyIds, only those rows are written (the position column still shows the position in the whole list).
/// </summary>
public static class RuleCsvExporter
{
    public static IReadOnlyList<string> Header =>
    [
        T("実行順", "Order"), T("有効", "Enabled"), T("名前", "Name"), T("種類", "Type"),
        T("条件", "Conditions"), T("例外", "Exceptions"), T("処理", "Actions"),
        T("移動先フォルダー", "Move to folder"), T("診断", "Diagnostics"),
    ];

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
                RuleText.EnabledText(e.Enabled),
                e.Name,
                RuleText.KindText(e),
                RuleText.Summary(e.Conditions),
                RuleText.Summary(e.Exceptions),
                RuleText.ActionSummary(e.Actions),
                RuleText.MoveTarget(e.Actions),
                string.Join(" / ", diags.Select(x => $"[{RuleText.SeverityText(x.Severity)}] {x.Message}")),
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

    internal static string Escape(string value)
    {
        if (value.IndexOfAny([',', '"', '\n', '\r']) < 0) return value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
