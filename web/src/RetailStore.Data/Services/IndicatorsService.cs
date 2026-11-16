using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Journal;
using RetailStore.Domain.Accounting;

namespace RetailStore.Data.Services;

/// <summary>The financial indicators of the dashboard: Financial = DASHBOARD_FINANCIAL (without it the cards show ••••), the four
/// cards (margin, turnover, collection, liquidity), the sums they come from, when they were computed and what the journal
/// update said.</summary>
public sealed record IndicatorsView(bool Financial, IndicatorCard[] Cards, IndicatorSums? Sums, DateTime ComputedAt, string JournalMessage);

/// <summary>modIndicators on the web (plan 6F, docs/82): FinancialIndicatorsQuery over the journal lines (qryIndicatorLines), so the
/// cards agree with the financial statements; the journal is brought up to date first (SyncJournal) and the figures are kept
/// 10 minutes for the whole site, as in Access (Force = the Refresh button).</summary>
public sealed class IndicatorsService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, JournalService journal)
{
    public const int CacheMinutes = 10;

    private static readonly object Gate = new();
    private static (DateTime AsOf, DateTime At, IndicatorSums Sums, string Message)? cache;

    /// <summary>ForgetIndicators: the next dashboard computes them again.</summary>
    public static void Forget()
    {
        lock (Gate)
        {
            cache = null;
        }
    }

    public async Task<IndicatorsView> IndicatorsAsync(ActingUser user, DateTime asOf, bool force = false)
    {
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "DASHBOARD_FINANCIAL"))
        {
            var hidden = Enumerable.Repeat(new IndicatorCard("••••", "", "muted"), 4).ToArray();
            return new IndicatorsView(false, hidden, null, DateTime.Now, "");
        }
        lock (Gate)
        {
            if (!force && cache is { } c && c.AsOf == asOf.Date && (DateTime.Now - c.At).TotalMinutes < CacheMinutes)
            {
                return new IndicatorsView(true, FinancialIndicators.Cards(c.Sums), c.Sums, c.At, c.Message);
            }
        }
        var sync = await journal.SyncAsync(user);              // the documents saved since the last time
        var sums = await SumsAsync(asOf);
        var now = DateTime.Now;
        lock (Gate)
        {
            cache = (asOf.Date, now, sums, sync.Message);
        }
        return new IndicatorsView(true, FinancialIndicators.Cards(sums), sums, now, sync.Message);
    }

    /// <summary>FinancialIndicatorsQuery: one row of sums over qryIndicatorLines (the journal lines with the group of their
    /// account, Level2Code).</summary>
    public async Task<IndicatorSums> SumsAsync(DateTime asOf)
    {
        var (end, month, prevMonth, year, days90) = FinancialIndicators.Dates(asOf);
        await using var db = await factory.CreateDbContextAsync();
        var lines = await (from l in db.JournalLines
                           join e in db.JournalEntries on l.EntryID equals e.EntryID
                           join a in db.Accounts on l.AccountCode equals a.AccountCode
                           where e.EntryDate < end
                           select new { l.AccountCode, a.Level2Code, e.EntryDate, e.SourceType, l.Debit, l.Credit })
                          .AsNoTracking().ToListAsync();
        decimal monthSales = 0, monthCost = 0, prevSales = 0, prevCost = 0, yearCost = 0, stockStart = 0, stockEnd = 0;
        decimal receivables = 0, creditSales = 0, assets = 0, liabilities = 0;
        foreach (var i in lines)
        {
            bool closing = i.SourceType == "YEAR_CLOSE";
            bool stock = i.AccountCode == 1400 || (i.AccountCode >= 140001 && i.AccountCode <= 149999);
            if (i.Level2Code == 41 && !closing)
            {
                if (i.EntryDate >= month) monthSales += i.Credit - i.Debit;
                else if (i.EntryDate >= prevMonth) prevSales += i.Credit - i.Debit;
            }
            if (i.Level2Code == 51 && !closing)
            {
                if (i.EntryDate >= month) monthCost += i.Debit - i.Credit;
                else if (i.EntryDate >= prevMonth) prevCost += i.Debit - i.Credit;
                if (i.EntryDate >= year) yearCost += i.Debit - i.Credit;
            }
            if (stock)
            {
                stockEnd += i.Debit - i.Credit;
                if (i.EntryDate < year) stockStart += i.Debit - i.Credit;
            }
            if (i.AccountCode == 1300)
            {
                receivables += i.Debit - i.Credit;
                if (i.SourceType == "SALE" && i.EntryDate >= days90) creditSales += i.Debit;
            }
            if (i.Level2Code == 11) assets += i.Debit - i.Credit;
            if (i.Level2Code == 21) liabilities += i.Credit - i.Debit;
        }
        return new IndicatorSums(monthSales, monthCost, prevSales, prevCost, yearCost, stockStart, stockEnd, receivables, creditSales,
                                 assets, liabilities);
    }
}
