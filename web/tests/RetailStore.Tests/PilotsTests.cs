using RetailStore.Data.Reports;
using RetailStore.UI.Components.Shared;
using RetailStore.UI.Services;
using RetailStore.UI.Theme;

namespace RetailStore.Tests;

/// <summary>The logic of the pilot screens of the redesign (phase 4): the periods and the 12 months of the dashboard
/// charts, the chart palette of the theme, and the unit kept on a line of the lines editor.</summary>
public class PilotsTests
{
    private static ReportRow Month(int year, int month, decimal sales, decimal profit) => new(new Dictionary<string, object?>
    {
        ["SalesYear"] = year, ["SalesMonth"] = month, ["NetSalesTotal"] = sales, ["GrossProfit"] = profit,
    });

    [Fact]
    public void The_dashboard_periods_are_the_last_12_months_and_the_last_30_days()
    {
        var today = new DateTime(2026, 11, 26);
        Assert.Equal((new DateTime(2025, 12, 1), today), HomeReports.Months(today));
        Assert.Equal((new DateTime(2026, 10, 28), today), HomeReports.LastDays(today));
    }

    [Fact]
    public void The_monthly_chart_has_every_month_in_order_with_zero_for_no_sale()
    {
        var today = new DateTime(2026, 11, 26);
        var series = HomeReports.MonthlySeries([Month(2026, 3, 1500m, 300m), Month(2025, 12, 200m, 40m)], today);
        Assert.Equal(12, series.Count);
        Assert.Equal("2025/12", series[0].Label);
        Assert.Equal("2026/11", series[11].Label);
        Assert.Equal(200d, series[0].Sales);
        Assert.Equal(1500d, series[3].Sales);
        Assert.Equal(300d, series[3].Profit);
        Assert.Equal(0d, series[1].Sales);
    }

    [Fact]
    public void The_chart_palette_comes_from_the_theme()
    {
        var light = ErpTheme.ChartPalette(false);
        var dark = ErpTheme.ChartPalette(true);
        Assert.Equal(6, light.Length);
        Assert.StartsWith("#1e3a5f", light[0], StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith("#60a5fa", dark[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_line_keeps_what_the_page_gives_it()
    {
        var line = new InvoiceLine<string> { Product = "P1", Quantity = 2, Price = 10, Tag = 7 };
        var r = InvoiceLines.Calculate([line, new InvoiceLine<string>()], 0, false);
        Assert.True(r.Ok);
        Assert.Equal(20m, r.Total);
        Assert.Equal(7, line.Tag);
    }
}
