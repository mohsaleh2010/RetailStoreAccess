using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Data.Reports;
using RetailStore.Data.Services;
using RetailStore.Domain.Accounting;

namespace RetailStore.Tests;

/// <summary>The financial indicators and the report centre of the web (plan 6F, docs/82): modIndicators and modScreens.PrepareReport
/// over the saved queries of Access (SQLite views here, the views of docs/71 on SQL Server).</summary>
public class ReportsTests
{
    private static readonly DateTime Today = DateTime.Today;

    private sealed class Store : IDisposable
    {
        public TestDb Db { get; } = new();
        public SalesService Sales => new(Db, Db.Permissions);
        public JournalService Journal => new(Db, new ServerViewJournalReader());
        public ReportService Reports => new(Db, Db.Permissions, Journal);
        public IndicatorsService Indicators => new(Db, Db.Permissions, Journal);
        public int Ali, Rice, Sugar;

        public Store()
        {
            using var db = Db.CreateDbContext();
            var ali = new Customers { CustomerName = "علي", CreditLimit = 1000, PaymentTermsDays = 30 };
            var rice = new Products { ProductCode = "R1", ProductName = "أرز", SellingPrice = 23, AverageCost = 12, CurrentQuantity = 10, MinimumQuantity = 9 };
            var sugar = new Products { ProductCode = "S1", ProductName = "سكر", SellingPrice = 11.5m, AverageCost = 6, CurrentQuantity = 100 };
            db.AddRange(ali, rice, sugar);
            db.SaveChanges();
            (Ali, Rice, Sugar) = (ali.CustomerID, rice.ProductID, sugar.ProductID);
        }

        public async Task SellAsync()
        {
            Assert.Equal(SaveStatus.Saved, (await Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(1, "CASH", 1, 0, null, "", [new CartLine(Rice, 2, 23)]))).Status);
            Assert.Equal(SaveStatus.Saved, (await Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(Ali, "CREDIT", null, 0, 0, "", [new CartLine(Sugar, 4, 11.5m)]))).Status);
        }

        public void Dispose() => Db.Dispose();
    }

    private static async Task<int> UserWithAsync(TestDb t, params string[] permissions)
    {
        await using var db = t.CreateDbContext();
        var role = new Roles { RoleCode = "R" + permissions.Length, RoleName = "دور" };
        db.Roles.Add(role);
        await db.SaveChangesAsync();
        foreach (var p in permissions)
        {
            db.RolePermissions.Add(new RolePermissions { RoleID = role.RoleID, PermissionKey = p });
        }
        var e = new Employees { EmployeeName = "محاسب", Username = "r" + role.RoleID, RoleID = role.RoleID };
        db.Employees.Add(e);
        await db.SaveChangesAsync();
        return e.EmployeeID;
    }

    [Fact]
    public void The_formulas_are_those_of_mod_indicators()
    {
        // TestIndicators (modIndicators) and tools/indicators_reference.py
        Assert.Equal(0.25, FinancialIndicators.GrossMargin(1000, 750));
        Assert.Null(FinancialIndicators.GrossMargin(0, 10));
        Assert.Equal(4, FinancialIndicators.StockTurnover(1200, 200, 400));
        Assert.Null(FinancialIndicators.StockTurnover(100, 0, 0));
        Assert.Equal(30, FinancialIndicators.CollectionDays(3000, 9000));
        Assert.Null(FinancialIndicators.CollectionDays(100, 0));
        Assert.Equal(0, FinancialIndicators.CollectionDays(-5, 900));
        Assert.Equal(1.5, FinancialIndicators.CurrentRatio(3000, 2000));
        Assert.Null(FinancialIndicators.CurrentRatio(100, 0));
        var cards = FinancialIndicators.Cards(new IndicatorSums(1000, 1100, 0, 0, 1200, 200, 400, 4000, 9000, 900, 1000));
        Assert.Equal(new IndicatorCard("-10.0%", "الشهر الماضي: -", "danger"), cards[0]);
        Assert.Equal(new IndicatorCard("4.0 مرة/سنة", "المخزون يكفي نحو 91 يومًا", "primary"), cards[1]);
        Assert.Equal(("40 يومًا", "warning"), (cards[2].Value, cards[2].Tone));
        Assert.Equal(("0.90", "danger"), (cards[3].Value, cards[3].Tone));
        cards = FinancialIndicators.Cards(new IndicatorSums(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0));
        Assert.Equal(["لا مبيعات هذا الشهر", "لا تكلفة مبيعات أو لا مخزون", "لا مبيعات آجلة في آخر 90 يومًا", "لا خصوم متداولة"],
                     cards.Select(c => c.Sub));
    }

    [Fact]
    public async Task The_indicators_come_from_the_journal_brought_up_to_date()
    {
        using var s = new Store();
        await s.SellAsync();
        IndicatorsService.Forget();
        var v = await s.Indicators.IndicatorsAsync(TestDb.Admin, Today, force: true);
        Assert.True(v.Financial);
        Assert.Equal("", v.JournalMessage);
        await using (var db = s.Db.CreateDbContext())
        {
            Assert.True(await db.JournalEntries.CountAsync() >= 2, "SyncJournal ran on the views of the saved queries");
        }
        // 2 rice at 23 (VAT included: 40 net, cost 24) + 4 sugar at 11.5 (40 net, cost 24): margin 40 %
        Assert.Equal((80m, 48m), (v.Sums!.MonthSales, v.Sums.MonthCost));
        Assert.Equal("40.0%", v.Cards[0].Value);
        Assert.Equal(46m, v.Sums.Receivables);                     // the credit sale with its VAT
        Assert.Equal(46m, v.Sums.CreditSales);

        int user = await UserWithAsync(s.Db, "JOURNAL");
        var hidden = await s.Indicators.IndicatorsAsync(new ActingUser(user, "T"), Today);
        Assert.False(hidden.Financial);
        Assert.All(hidden.Cards, c => Assert.Equal("••••", c.Value));
    }

    [Fact]
    public async Task Every_report_of_the_centre_runs_on_the_saved_queries()
    {
        using var s = new Store();
        await s.SellAsync();
        var choices = new ReportChoices(Today.AddDays(-30), Today, CustomerId: s.Ali, SupplierId: null, ProductId: s.Rice);
        foreach (var r in ReportCatalog.All)
        {
            var c = r.Has('S') ? choices with { SupplierId = 1 } : choices;
            if (r.Has('S'))
            {
                await using var db = s.Db.CreateDbContext();
                if (!await db.Suppliers.AnyAsync())
                {
                    db.Suppliers.Add(new Suppliers { SupplierName = "مورد" });
                    await db.SaveChangesAsync();
                }
                c = c with { SupplierId = await db.Suppliers.Select(x => x.SupplierID).FirstAsync() };
            }
            var result = await s.Reports.RunAsync(TestDb.Admin, r.Key, c, false);
            Assert.True(result.Ok, r.Key + ": " + result.Message);
            Assert.Equal("", result.Warning);
            var running = ReportService.Running(r, result.Rows);
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int col = 0; col < r.Columns.Count; col++)
                {
                    _ = ReportService.CellText(r, result.Rows, running, i, col);
                }
                foreach (var cr in r.Card)
                {
                    _ = ReportRow.Cell(cr.Value(result.Rows[i]), cr.Format);
                }
            }
            foreach (var sum in r.Summary)
            {
                _ = sum.Value(result.Rows);
            }
        }
        var daily = await s.Reports.RunAsync(TestDb.Admin, "DAILY_SALES", choices, false);
        var day = Assert.Single(daily.Rows);
        Assert.Equal(ReportRow.GDate(Today, false), daily.Report!.Columns[0].Value(day));
        Assert.Equal("الفترة: من " + ReportRow.GDate(Today.AddDays(-30), false) + " إلى " + ReportRow.GDate(Today, false), daily.Criteria);

        var statement = await s.Reports.RunAsync(TestDb.Admin, "CUSTOMER_STATEMENT", choices, false);
        Assert.Contains("العميل: علي", statement.Criteria);
        var balance = ReportService.Running(statement.Report!, statement.Rows);
        Assert.Equal(46m, balance[statement.Rows.Count - 1, 5]);   // the running balance of the customer

        var trial = await s.Reports.RunAsync(TestDb.Admin, "TRIAL_BALANCE", choices, false);
        Assert.NotEmpty(trial.Rows);
        Assert.Equal(0m, trial.Rows.Sum(x => x.Num("ClosingBalance")));

        var low = await s.Reports.RunAsync(TestDb.Admin, "LOW_STOCK", new ReportChoices(), false);
        Assert.Equal("أرز", Assert.Single(low.Rows).Text("ProductName"));   // 8 left, minimum 9

        var link = await s.Reports.RunAsync(TestDb.Admin, "INCOME_STATEMENT", choices, false);
        Assert.Equal(("financials", 0), (link.Report!.Link, link.Rows.Count));
        Assert.Contains("أرز", ReportService.Csv(low.Report!, low.Rows, x => x));
    }

    [Fact]
    public async Task The_choices_are_checked_in_the_order_of_access()
    {
        using var s = new Store();
        Assert.Equal("اختر تقريرًا من القائمة أولًا.", (await s.Reports.RunAsync(TestDb.Admin, "NONE", new(), false)).Message);
        int user = await UserWithAsync(s.Db, "REPORTS");
        var u = new ActingUser(user, "T");
        Assert.Equal("هذا التقرير يعرض التكلفة والأرباح ويحتاج صلاحية «تقارير الأرباح والضريبة».",
                     (await s.Reports.RunAsync(u, "PROFIT", new(), false)).Message);
        Assert.Equal("سجل التدقيق يحتاج صلاحية «سجل التدقيق».", (await s.Reports.RunAsync(u, "AUDIT_TRAIL", new(), false)).Message);
        Assert.Equal("تقارير الخزينة تحتاج صلاحية «الخزينة».", (await s.Reports.RunAsync(u, "CASH_BALANCES", new(), false)).Message);
        Assert.Equal("التقارير المحاسبية تحتاج صلاحية «قيود اليومية».", (await s.Reports.RunAsync(u, "JOURNAL", new(), false)).Message);
        Assert.Equal("أدخل تاريخ البداية وتاريخ النهاية.", (await s.Reports.RunAsync(u, "DAILY_SALES", new(), false)).Message);
        Assert.Equal("تاريخ النهاية يجب أن يكون بعد تاريخ البداية.",
                     (await s.Reports.RunAsync(u, "DAILY_SALES", new(Today, Today.AddDays(-1)), false)).Message);
        Assert.Equal("اختر العميل.", (await s.Reports.RunAsync(u, "CUSTOMER_STATEMENT", new(Today, Today), false)).Message);
        Assert.Equal("اختر المنتج.", (await s.Reports.RunAsync(u, "PRODUCT_MOVEMENT", new(Today, Today), false)).Message);
        Assert.Equal("حدد الفترة واختر العميل", ReportService.NeedsHint(ReportCatalog.Find("CUSTOMER_STATEMENT")));
        Assert.Equal("لا يحتاج هذا التقرير أي اختيارات", ReportService.NeedsHint(ReportCatalog.Find("STOCK")));
        Assert.Equal("حدد الفترة، ويمكنك التصفية حسب الاختيار (اختياري)", ReportService.NeedsHint(ReportCatalog.Find("SALES_PERIOD")));
        var ok = await s.Reports.RunAsync(u, "STOCK", new(), false);
        Assert.True(ok.Ok);
        await using var db = s.Db.CreateDbContext();
        Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "REPORT" && a.ObjectName == "STOCK" && a.EmployeeID == user));
    }

    [Fact]
    public void The_expressions_of_the_reports_are_those_of_access()
    {
        var row = new ReportRow(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["AccountLevel"] = 3L, ["AccountName"] = "الصندوق", ["PaymentType"] = "CREDIT", ["SalesYear"] = 2026L, ["SalesMonth"] = 3L,
            ["EntryDate"] = "2026-03-05 14:07:00", ["Debit"] = "10.5", ["Credit"] = 0.5,
        });
        var tree = ReportCatalog.Find("ACCOUNT_TREE")!.Columns.First(c => c.Caption == "الحساب");
        Assert.Equal(new string(' ', 6) + "الصندوق", tree.Value(row));
        var monthly = ReportCatalog.Find("MONTHLY_SALES")!.Columns[0];
        Assert.Equal("2026/03", monthly.Value(row));
        var statement = ReportCatalog.Find("CUSTOMER_STATEMENT")!.Columns;
        Assert.Equal("2026/03/05 14:07", statement[0].Value(row));
        Assert.Equal(10m, statement[5].Value(row));
        Assert.Equal(("1,234.50", "12.5%", "3", "1,234.568"), (ReportRow.Cell(1234.5m, "money"), ReportRow.Cell(0.125, "pct"),
                     ReportRow.Cell(3L, "int"), ReportRow.Cell(1234.5678m, "qty")));
    }
}
