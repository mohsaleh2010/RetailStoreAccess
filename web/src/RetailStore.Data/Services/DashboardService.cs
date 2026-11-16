using Microsoft.EntityFrameworkCore;

namespace RetailStore.Data.Services;

/// <summary>The figures of the dashboard (modDashboard.DashboardRefresh over DashboardQuery, the month profit of ProfitQuery).
/// Financial = DASHBOARD_FINANCIAL: without it the amounts are hidden (null) and only the counts stay, as in Access.</summary>
public sealed record DashboardFigures(bool Financial, decimal? TodaySales, int TodayInvoices, decimal? MonthSales, decimal? MonthVAT,
                                      int MonthInvoices, decimal? MonthNetProfit, int LowStockCount, decimal? CustomerDebt, int DebtorCount,
                                      decimal? SupplierDue, decimal? StockValue, decimal? MonthExpenses, decimal? CashBalance, decimal? BankBalance);

/// <summary>The month profit of ProfitQuery: net sales, cost of sales, stock adjustments, expenses.</summary>
public sealed record ProfitFigures(decimal NetSales, decimal CostOfSales, decimal GrossProfit, decimal InventoryAdjustments, decimal TotalExpenses,
                                   decimal NetProfit);

/// <summary>The dashboard of the web (plan 6E, docs/81): the cards of frmMain (today and month sales, month net profit, products
/// at or under their minimum, customer debts, supplier dues, stock value, month expenses) from the same tables as DashboardQuery,
/// plus the cash in the boxes and the banks.</summary>
public sealed class DashboardService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    /// <summary>ProfitQuery for [from, to]: qryProfitSales (the sales lines less the returns, the cost of the lines less the cost of
    /// the returns put back in stock), qryProfitAdjustments (STOCK_IN / STOCK_OUT / ADJUSTMENT moves) and qryProfitExpenses.</summary>
    public static async Task<ProfitFigures> ProfitAsync(RetailDbContext db, DateTime from, DateTime to)
    {
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var sales = await (from d in db.SalesInvoiceDetails
                           join h in db.SalesInvoices on d.SalesInvoiceID equals h.SalesInvoiceID
                           where h.InvoiceDate >= start && h.InvoiceDate < end
                           select new { d.NetAmount, Cost = d.Quantity * d.UnitCost }).ToListAsync();
        var returns = await (from d in db.SalesReturnDetails
                             join r in db.SalesReturns on d.SalesReturnID equals r.SalesReturnID
                             where r.ReturnDate >= start && r.ReturnDate < end
                             select new { d.NetAmount, Cost = d.ReturnToStock ? d.Quantity * d.UnitCost : 0 }).ToListAsync();
        decimal net = sales.Sum(x => x.NetAmount) - returns.Sum(x => x.NetAmount);
        decimal cost = sales.Sum(x => x.Cost) - returns.Sum(x => x.Cost);
        string[] codes = ["STOCK_IN", "STOCK_OUT", "ADJUSTMENT"];
        var adjustments = await (from t in db.InventoryTransactions
                                 join tt in db.TransactionTypes on t.TransactionTypeID equals tt.TransactionTypeID
                                 where codes.Contains(tt.TypeCode) && t.TransactionDate >= start && t.TransactionDate < end
                                 select t.Quantity * t.UnitCost).ToListAsync();
        decimal adjust = adjustments.Sum();
        decimal expenses = (await db.Expenses.Where(e => e.ExpenseDate >= start && e.ExpenseDate < end).Select(e => e.Amount).ToListAsync()).Sum();
        return new ProfitFigures(net, cost, net - cost, adjust, expenses, net - cost + adjust - expenses);
    }

    /// <summary>DashboardQuery for a day: today (the day) and the month until the day.</summary>
    public async Task<DashboardFigures> FiguresAsync(int employeeId, DateTime today)
    {
        bool financial = await permissions.HasPermissionAsync(employeeId, "DASHBOARD_FINANCIAL");
        await using var db = await factory.CreateDbContextAsync();
        var day = today.Date;
        var month = new DateTime(day.Year, day.Month, 1);
        var end = day.AddDays(1);
        // qrySalesDocuments: the invoices less the returns
        var invoices = await db.SalesInvoices.Where(h => h.InvoiceDate >= month && h.InvoiceDate < end)
            .Select(h => new { h.InvoiceDate, h.TotalAmount, h.Tax }).ToListAsync();
        var returns = await db.SalesReturns.Where(r => r.ReturnDate >= month && r.ReturnDate < end)
            .Select(r => new { r.ReturnDate, r.TotalAmount, r.Tax }).ToListAsync();
        decimal todaySales = invoices.Where(i => i.InvoiceDate >= day).Sum(i => i.TotalAmount) - returns.Where(r => r.ReturnDate >= day).Sum(r => r.TotalAmount);
        decimal monthSales = invoices.Sum(i => i.TotalAmount) - returns.Sum(r => r.TotalAmount);
        decimal monthVat = invoices.Sum(i => i.Tax) - returns.Sum(r => r.Tax);
        int todayCount = invoices.Count(i => i.InvoiceDate >= day);
        decimal expenses = (await db.Expenses.Where(e => e.ExpenseDate >= month && e.ExpenseDate < end).Select(e => e.Amount).ToListAsync()).Sum();
        var debts = await db.Customers.Where(c => c.CurrentBalance > 0).Select(c => c.CurrentBalance).ToListAsync();
        decimal due = (await db.Suppliers.Where(s => s.CurrentBalance > 0).Select(s => s.CurrentBalance).ToListAsync()).Sum();
        decimal stock = (await db.Products.Where(p => p.IsActive && p.CurrentQuantity > 0).Select(p => p.CurrentQuantity * p.AverageCost).ToListAsync()).Sum();
        // LowStockQuery
        int low = await db.Products.CountAsync(p => p.IsActive && p.TrackStock && p.CurrentQuantity <= p.MinimumQuantity);
        decimal? profit = null, cash = null, bank = null;
        if (financial)
        {
            profit = (await ProfitAsync(db, month, day)).NetProfit;
            cash = (await TreasuryService.CashMovesAsync(db, 0, false)).Sum(m => m.AmountIn - m.AmountOut);
            bank = (await db.JournalLines.Where(l => l.AccountCode > 120000 && l.AccountCode < 130000).Select(l => l.Debit - l.Credit).ToListAsync()).Sum();
        }
        decimal? M(decimal v) => financial ? v : null;
        return new DashboardFigures(financial, M(todaySales), todayCount, M(monthSales), M(monthVat), invoices.Count, profit, low, M(debts.Sum()),
                                    debts.Count, M(due), M(stock), M(expenses), cash, bank);
    }
}
