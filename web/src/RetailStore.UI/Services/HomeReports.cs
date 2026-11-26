using RetailStore.Data;
using RetailStore.Data.Reports;
using RetailStore.Data.Services;

namespace RetailStore.UI.Services;

/// <summary>The reports of the dashboard (template D): the monthly sales and the best sellers (charts), and the products
/// under their reorder level, run by the report centre itself (ReportService.RunAsync: its permissions, its saved queries,
/// its audit line), and kept for the session 10 minutes like the indicators, so going back to the dashboard does not
/// run them again; the Refresh button does.</summary>
public sealed class HomeReports(ReportService reports)
{
    public static readonly TimeSpan KeepFor = TimeSpan.FromMinutes(10);

    public sealed record Result(ReportResult Monthly, ReportResult Best, ReportResult LowStock, DateTime At);

    private Result? kept;
    private (int Employee, bool English, DateTime Day) keptFor;

    /// <summary>The last 12 months, from the first of the month 11 months ago (the monthly chart).</summary>
    public static (DateTime From, DateTime To) Months(DateTime today) => (new DateTime(today.Year, today.Month, 1).AddMonths(-11), today.Date);

    /// <summary>The last 30 days (the best sellers).</summary>
    public static (DateTime From, DateTime To) LastDays(DateTime today) => (today.Date.AddDays(-29), today.Date);

    public async Task<Result> LoadAsync(ActingUser user, DateTime today, bool english, bool force)
    {
        var key = (user.EmployeeID, english, today.Date);
        if (!force && kept is not null && keptFor == key && DateTime.Now - kept.At < KeepFor)
        {
            return kept;
        }
        var (mFrom, mTo) = Months(today);
        var (bFrom, bTo) = LastDays(today);
        var monthly = await reports.RunAsync(user, "MONTHLY_SALES", new ReportChoices(mFrom, mTo), english);
        var best = await reports.RunAsync(user, "BEST_SELLING", new ReportChoices(bFrom, bTo), english);
        var low = await reports.RunAsync(user, "LOW_STOCK", new ReportChoices(), english);
        kept = new Result(monthly, best, low, DateTime.Now);
        keptFor = key;
        return kept;
    }

    /// <summary>The 12 months of the chart, each with its sales and gross profit (0 for a month with no sale).</summary>
    public static List<(string Label, double Sales, double Profit)> MonthlySeries(IReadOnlyList<ReportRow> rows, DateTime today)
    {
        var (from, _) = Months(today);
        var byMonth = rows.ToDictionary(r => ((int)r.Num("SalesYear"), (int)r.Num("SalesMonth")));
        var list = new List<(string, double, double)>();
        for (int i = 0; i < 12; i++)
        {
            var m = from.AddMonths(i);
            byMonth.TryGetValue((m.Year, m.Month), out var r);
            list.Add((m.ToString("yyyy/MM", System.Globalization.CultureInfo.InvariantCulture),
                      (double)(r?.Num("NetSalesTotal") ?? 0), (double)(r?.Num("GrossProfit") ?? 0)));
        }
        return list;
    }
}
