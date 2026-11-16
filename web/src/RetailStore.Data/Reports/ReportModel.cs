using System.Globalization;

namespace RetailStore.Data.Reports;

/// <summary>A column of a report (reports_catalog.Col): its caption, its format (money, qty, int, pct, rate or "" for text and
/// dates), whether the report footer adds it up, whether it is a running balance, and its value (a field or the C# of the
/// Access expression, written by tools/gen_web_reports.py).</summary>
public sealed record ReportColumn(string Caption, string Format, bool Total, bool Running, Func<ReportRow, object?> Value);

/// <summary>A summary row of the report footer ("=Sum(...)", in money).</summary>
public sealed record ReportSummary(string Caption, Func<IReadOnlyList<ReportRow>, decimal> Value);

/// <summary>A labelled row of a card report (profit, VAT summary: one record shown as rows).</summary>
public sealed record ReportCardRow(string Caption, string Format, bool Bold, Func<ReportRow, object?> Value);

/// <summary>A report of the report centre (forms.ReportEntry) with its layout. Needs = the letters of modScreens.PrepareReport:
/// P period, D period filtered on DateColumn, C / S / R a customer / supplier / product to choose (query parameter),
/// c / s / r / e / w an optional filter on CustomerID / SupplierID / ProductID / ExpenseTypeID / WarehouseID, W / b an optional
/// warehouse / cash box (query parameter, 0 = all), $ profit figures (REPORTS_PROFIT), A audit trail (AUDIT_LOG), # treasury
/// (CASH_BOX), J accounting (JOURNAL, the journal is updated first), F the financial statements. Link = a web page that shows
/// the report instead (the financial statements and their comparison period).</summary>
public sealed record ReportDef(string Key, string Title, string Query, string Needs, string DateColumn, string Link, string NoData,
                               string Note, IReadOnlyList<ReportColumn> Columns, IReadOnlyList<(string Field, bool Descending)> Sorts,
                               IReadOnlyList<ReportSummary> Summary, IReadOnlyList<ReportCardRow> Card)
{
    /// <summary>modScreens.HasNeed: "C" and "c" differ.</summary>
    public bool Has(char letter) => Needs.Contains(letter, StringComparison.Ordinal);

    public bool Dated => Has('P') || Has('D');
}

public static partial class ReportCatalog
{
    public static ReportDef? Find(string? key) => All.FirstOrDefault(r => r.Key == key);
}

/// <summary>One record of a saved query, by column name (any case). Num / Text / Bool read a value as the report expressions use
/// it; the static helpers are the functions of those expressions (GDate, Space, Format).</summary>
public sealed class ReportRow(IReadOnlyDictionary<string, object?> values)
{
    public IReadOnlyDictionary<string, object?> Values => values;

    public object? Value(string field) => values.TryGetValue(field, out var v) && v is not DBNull ? v : null;

    public decimal Num(string field) => ToNumber(Value(field));

    public string Text(string field) => ToText(Value(field));

    public bool Bool(string field) => Value(field) switch
    {
        null => false,
        bool b => b,
        string s => s.Length > 0 && s != "0",
        var v => ToNumber(v) != 0,
    };

    public static decimal ToNumber(object? v) => v switch
    {
        null or DBNull => 0m,
        decimal d => d,
        double d => double.IsFinite(d) ? (decimal)d : 0m,
        float f => (decimal)f,
        bool b => b ? -1m : 0m,
        string s => decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var x) ? x : 0m,
        IConvertible c => c.ToDecimal(CultureInfo.InvariantCulture),
        _ => 0m,
    };

    public static string ToText(object? v) => v switch
    {
        null or DBNull => "",
        string s => s,
        decimal d => NumberText(d),
        double d => NumberText((decimal)d),
        DateTime t => t.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => v.ToString() ?? "",
    };

    public static string NumberText(decimal d) => d.ToString("0.############", CultureInfo.InvariantCulture);

    public static DateTime? ToDate(object? v) => v switch
    {
        DateTime t => t,
        string s when DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var t) => t,
        _ => null,
    };

    /// <summary>modReports.GDate: always Gregorian, yyyy/mm/dd (hh:nn).</summary>
    public static string GDate(object? v, bool withTime) =>
        ToDate(v) is { } t ? t.ToString(withTime ? "yyyy/MM/dd HH:mm" : "yyyy/MM/dd", CultureInfo.InvariantCulture) : "";

    /// <summary>Space(n): non-breaking spaces, so that the indentation of the account tree shows in a page.</summary>
    public static string Space(decimal n) => new(' ', (int)Math.Max(0, Math.Min(n, 200)));

    /// <summary>Format(x, "00") of a number, Format(x, "hh:nn") of a date.</summary>
    public static string Format(object? v, string pattern) => pattern switch
    {
        "hh:nn" => ToDate(v) is { } t ? t.ToString("HH:mm", CultureInfo.InvariantCulture) : "",
        _ => v is null ? "" : ToNumber(v).ToString(pattern, CultureInfo.InvariantCulture),
    };

    /// <summary>A ratio of the report expressions; never a division by zero.</summary>
    public static decimal Divide(decimal a, decimal b) => b == 0 ? 0 : a / b;

    /// <summary>A cell as the Access report shows it (the Format of its text box).</summary>
    public static string Cell(object? v, string format)
    {
        if (v is null or DBNull)
        {
            return "";
        }
        if (format.Length == 0)
        {
            return v is DateTime t ? GDate(t, t.TimeOfDay != TimeSpan.Zero) : ToText(v);
        }
        decimal d = ToNumber(v);
        return format switch
        {
            "money" => d.ToString("#,##0.00", CultureInfo.InvariantCulture),
            "qty" => d.ToString("#,##0.###", CultureInfo.InvariantCulture),
            "int" => d.ToString("#,##0", CultureInfo.InvariantCulture),
            "pct" => (d * 100).ToString("0.0", CultureInfo.InvariantCulture) + "%",
            "rate" => d.ToString("0.0000", CultureInfo.InvariantCulture),
            _ => ToText(v),
        };
    }

    /// <summary>The order of the Access report (its sorting and grouping): numbers as numbers, dates as dates, texts as texts.</summary>
    public static int Compare(object? a, object? b)
    {
        if (a is null || b is null)
        {
            return a is null ? (b is null ? 0 : -1) : 1;
        }
        if (a is string sa && b is string sb)
        {
            if (ToDate(sa) is { } da && ToDate(sb) is { } db && sa.Length >= 10 && sa[4] == '-')
            {
                return da.CompareTo(db);
            }
            return string.Compare(sa, sb, StringComparison.Ordinal);
        }
        if (a is DateTime ta && b is DateTime tb)
        {
            return ta.CompareTo(tb);
        }
        if (a is string || b is string)
        {
            return string.Compare(ToText(a), ToText(b), StringComparison.Ordinal);
        }
        return ToNumber(a).CompareTo(ToNumber(b));
    }
}
