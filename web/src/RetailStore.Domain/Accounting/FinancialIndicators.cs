using System.Globalization;

namespace RetailStore.Domain.Accounting;

/// <summary>The sums of FinancialIndicatorsQuery (one row): the month and the previous month of groups 41 / 51 without the
/// year closing, the cost of the last 365 days, the stock (1400 and 140001-149999) before those days and now, the customers
/// (1300), the credit sales of the last 90 days, the current assets (group 11) and liabilities (group 21).</summary>
public sealed record IndicatorSums(decimal MonthSales, decimal MonthCost, decimal PrevSales, decimal PrevCost, decimal YearCost,
                                   decimal StockStart, decimal StockEnd, decimal Receivables, decimal CreditSales,
                                   decimal CurrentAssets, decimal CurrentLiabilities);

/// <summary>A card of the dashboard: its value, the line under it and its colour (muted, primary, danger, warning, success).</summary>
public sealed record IndicatorCard(string Value, string Sub, string Tone);

/// <summary>modIndicators in C#: the four formulas (null = no meaning, nothing to divide by; tools/indicators_reference.py) and
/// the four cards of frmMain with their texts and colours.</summary>
public static class FinancialIndicators
{
    public const int CollectionSpan = 90;

    public static double? GrossMargin(decimal sales, decimal cost) => sales > 0 ? (double)(sales - cost) / (double)sales : null;

    public static double? StockTurnover(decimal yearCost, decimal stockStart, decimal stockEnd)
    {
        double average = ((double)stockStart + (double)stockEnd) / 2;
        return average > 0 && yearCost > 0 ? (double)yearCost / average : null;
    }

    public static double? CollectionDays(decimal receivables, decimal creditSales, int days = CollectionSpan) =>
        creditSales > 0 ? (double)Math.Max(receivables, 0) / ((double)creditSales / days) : null;

    public static double? CurrentRatio(decimal currentAssets, decimal currentLiabilities) =>
        currentLiabilities > 0 ? (double)currentAssets / (double)currentLiabilities : null;

    /// <summary>The query parameters of SetIndicatorParams: IndEnd (exclusive), IndMonth, IndPrevMonth, IndYear, Ind90.</summary>
    public static (DateTime End, DateTime Month, DateTime PrevMonth, DateTime Year, DateTime Days90) Dates(DateTime asOf)
    {
        var d = asOf.Date;
        var month = new DateTime(d.Year, d.Month, 1);
        return (d.AddDays(1), month, month.AddMonths(-1), d.AddDays(-364), d.AddDays(-(CollectionSpan - 1)));
    }

    private static string F(double v, string format) => v.ToString(format, CultureInfo.InvariantCulture);

    private static string Percent(double? v) => v is { } x ? F(x * 100, "0.0") + "%" : "-";

    /// <summary>modIndicators.FinancialIndicators: margin, turnover, collection, liquidity.</summary>
    public static IndicatorCard[] Cards(IndicatorSums s)
    {
        var cards = new IndicatorCard[4];
        var v = GrossMargin(s.MonthSales, s.MonthCost);
        cards[0] = v is not { } m
            ? new("-", "لا مبيعات هذا الشهر", "muted")
            : new(Percent(m), "الشهر الماضي: " + Percent(GrossMargin(s.PrevSales, s.PrevCost)), m < 0 ? "danger" : "primary");
        v = StockTurnover(s.YearCost, s.StockStart, s.StockEnd);
        cards[1] = v is not { } t
            ? new("-", "لا تكلفة مبيعات أو لا مخزون", "muted")
            : new(F(t, "0.0") + " مرة/سنة", "المخزون يكفي نحو " + F(365 / t, "0") + " يومًا", "primary");
        v = CollectionDays(s.Receivables, s.CreditSales);
        cards[2] = v is not { } c
            ? new("-", "لا مبيعات آجلة في آخر " + CollectionSpan + " يومًا", "muted")
            : new(F(c, "0") + " يومًا", "ديون العملاء ÷ المبيعات الآجلة اليومية", c > 60 ? "danger" : c > 30 ? "warning" : "primary");
        v = CurrentRatio(s.CurrentAssets, s.CurrentLiabilities);
        cards[3] = v is not { } r
            ? new("-", "لا خصوم متداولة", "muted")
            : new(F(r, "0.00"), "الأصول المتداولة ÷ الخصوم المتداولة", r < 1 ? "danger" : r < 1.5 ? "warning" : "success");
        return cards;
    }
}
