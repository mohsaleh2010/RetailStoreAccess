using System.Globalization;
using System.Text;

namespace RetailStore.UI.Components.Shared;

/// <summary>A column of the CSV export of EntityGrid: its header (already in the interface language) and its value.</summary>
public sealed record GridExport<T>(string Header, Func<T, object?> Value);

/// <summary>The CSV file Excel opens (the decision of the owner: no Excel library): UTF-8 with a BOM so Arabic reads
/// right, comma separated, quotes doubled, numbers and dates invariant, and a text that starts like a formula
/// (= + - @) kept as text.</summary>
public static class GridCsv
{
    public static string Build<T>(IEnumerable<T> items, IReadOnlyList<GridExport<T>> columns)
    {
        var sb = new StringBuilder("﻿");
        sb.AppendJoin(',', columns.Select(c => Cell(c.Header))).Append("\r\n");
        foreach (var item in items)
        {
            sb.AppendJoin(',', columns.Select(c => Cell(c.Value(item)))).Append("\r\n");
        }
        return sb.ToString();
    }

    public static string Cell(object? value)
    {
        string text = value switch
        {
            null => "",
            decimal d => d.ToString("0.####", CultureInfo.InvariantCulture),
            double d => d.ToString("0.####", CultureInfo.InvariantCulture),
            float f => f.ToString("0.####", CultureInfo.InvariantCulture),
            DateTime t => t.TimeOfDay == TimeSpan.Zero
                ? t.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : t.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
            bool b => b ? "1" : "0",
            IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? "",
        };
        if (value is string && text.Length > 0 && "=+-@\t\r".Contains(text[0]))
        {
            text = "'" + text;
        }
        return text.IndexOfAny([',', '"', '\r', '\n']) >= 0 ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
    }

    /// <summary>A file name without the characters Windows refuses.</summary>
    public static string FileName(string name)
    {
        var bad = Path.GetInvalidFileNameChars().Concat(['/', '\\', ':', '*', '?', '"', '<', '>', '|']).ToHashSet();
        string clean = new string(name.Select(c => bad.Contains(c) ? '-' : c).ToArray()).Trim();
        return (clean.Length == 0 ? "export" : clean) + ".csv";
    }
}
