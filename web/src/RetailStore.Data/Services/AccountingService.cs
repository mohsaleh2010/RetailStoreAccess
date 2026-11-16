using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Journal;
using RetailStore.Domain.Localization;

namespace RetailStore.Data.Services;

public sealed record JournalEntryRow(int EntryID, string EntryNumber, DateTime EntryDate, string SourceType, string KindName, string? SourceNumber,
                                     string? Description, decimal TotalDebit, decimal TotalCredit, short LineCount);

public sealed record JournalLineView(short LineNumber, int AccountCode, string AccountName, decimal Debit, decimal Credit, string? LineText,
                                     string? CostCenterName);

public sealed record JournalEntryView(JournalEntryRow Entry, string? CurrencyCode, decimal ExchangeRate, decimal ForeignAmount,
                                      List<JournalLineView> Lines);

/// <summary>An account for the choice lists: IsPosting = a sub-account (a main account takes all its sub-accounts).</summary>
public sealed record AccountChoice(int AccountCode, string AccountName, string AccountType, bool IsPosting, int Level);

/// <summary>A line of the account statement (AccountStatementQuery): the opening line first (EntryRef 0), then each journal line
/// with the balance after it (modLedger.FillLedger).</summary>
public sealed record LedgerLine(int EntryRef, DateTime LineDate, string EntryNo, string KindName, string? DocNo, string? Details, string? SubName,
                                decimal Debit, decimal Credit, decimal Balance);

/// <summary>The statement of an account: its lines and the four figures (opening, debit, credit, closing).</summary>
public sealed record LedgerView(int AccountCode, string AccountName, bool IsMain, decimal Opening, decimal Debit, decimal Credit, decimal Closing,
                                int LineCount, List<LedgerLine> Lines);

/// <summary>A row of the trial balance (TrialBalanceQuery).</summary>
public sealed record TrialBalanceRow(int AccountCode, string AccountName, string AccountType, decimal OpeningBalance, decimal PeriodDebit,
                                     decimal PeriodCredit, decimal ClosingBalance);

/// <summary>A row of a financial statement (IncomeStatementQuery / BalanceSheetQuery): RowKind A = an account, H = a heading,
/// C = a class, G = a group, S = a group total, T = a total, R = a result; Prior = the comparison period.</summary>
public sealed record StatementRow(string RowKind, string Caption, int? LineAccount, decimal? CurrentValue, decimal? PriorValue);

/// <summary>A financial statement: its rows in their printed order, the comparison period and the three figures of frmFinancials.</summary>
public sealed record StatementView(string Kind, DateTime From, DateTime To, DateTime CompareFrom, DateTime CompareTo, List<StatementRow> Rows,
                                   List<(string Caption, decimal Amount)> Figures);

/// <summary>Accounting on the web and the API (plan 6E, docs/81): the journal entries and their lines (frmJournal), the update of
/// the journal (modJournal.SyncJournal, JournalService), the account statement (modLedger, AccountStatementQuery), the trial
/// balance (TrialBalanceQuery) and the financial statements (modFinancials: IncomeStatementQuery and BalanceSheetQuery with
/// their comparison period, modQueryParams.SetComparePeriod), read from JournalEntries / JournalLines as Access reads them.</summary>
public sealed class AccountingService(IDbContextFactory<RetailDbContext> factory, JournalService journal)
{
    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    private static string Tr(string arabic, bool english) => english ? Translator.ToEnglish(arabic) : arabic;

    /// <summary>The checks of the period of frmLedger / frmFinancials ("" = fine).</summary>
    public static string PeriodProblem(DateTime? from, DateTime? to)
    {
        if (from is null || to is null)
        {
            return "أدخل تاريخ البداية وتاريخ النهاية.";
        }
        return to.Value.Date < from.Value.Date ? "تاريخ النهاية قبل تاريخ البداية." : "";
    }

    /// <summary>modLedger.BalanceText: 1,250.00 مدين / 300.00 دائن / 0.00.</summary>
    public static string BalanceText(decimal amount)
    {
        string Money(decimal v) => v.ToString("#,##0.00", System.Globalization.CultureInfo.InvariantCulture);
        if (amount > 0)
        {
            return Money(amount) + " مدين";
        }
        return amount < 0 ? Money(-amount) + " دائن" : "0.00";
    }

    /// <summary>modQueryParams.SetComparePeriod: from the first of a month, the same months one step back (a month -> the month
    /// before, Jan 1 - today -> the same dates last year); otherwise the same number of days just before. Both days included.</summary>
    public static (DateTime From, DateTime To) ComparePeriod(DateTime from, DateTime to)
    {
        from = from.Date;
        to = to.Date;
        if (from.Day == 1)
        {
            int k = (to.Year - from.Year) * 12 + to.Month - from.Month + 1;
            var cFrom = from.AddMonths(-k);
            var cTo = to.AddMonths(-k);
            if (cTo >= from)
            {
                cTo = from.AddDays(-1);
            }
            return (cFrom, cTo);
        }
        int days = (to - from).Days + 1;
        var end = from.AddDays(-1);
        return (end.AddDays(1 - days), end);
    }

    /// <summary>modJournal.SyncJournal (JournalService): the entries are rebuilt from the operations.</summary>
    public Task<JournalSyncResult> SyncAsync(ActingUser user) => journal.SyncAsync(user);

    // ------------------------------------------------------------------ journal entries (frmJournal)

    /// <summary>The entries of a period, newest first: a kind (journal source) and a search in the entry or document number or
    /// the description.</summary>
    public async Task<List<JournalEntryRow>> EntriesAsync(DateTime from, DateTime to, string? sourceType, string? search, bool english, int take = 300)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        string s = (search ?? "").Trim();
        var q = db.JournalEntries.AsNoTracking().Where(e => e.EntryDate >= start && e.EntryDate < end);
        if (!string.IsNullOrEmpty(sourceType))
        {
            q = q.Where(e => e.SourceType == sourceType);
        }
        if (s.Length > 0)
        {
            q = q.Where(e => e.EntryNumber.Contains(s) || (e.SourceNumber != null && e.SourceNumber.Contains(s)) ||
                             (e.Description != null && e.Description.Contains(s)));
        }
        var rows = await q.OrderByDescending(e => e.EntryDate).ThenByDescending(e => e.EntryID).Take(take).ToListAsync();
        var kinds = await KindsAsync(db, english);
        return rows.Select(e => new JournalEntryRow(e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, kinds.TryGetValue(e.SourceType, out var k) ? k : e.SourceType,
                                                    e.SourceNumber, e.Description, e.TotalDebit, e.TotalCredit, e.LineCount)).ToList();
    }

    private static async Task<Dictionary<string, string>> KindsAsync(RetailDbContext db, bool english) =>
        (await db.JournalSourceTypes.AsNoTracking().ToListAsync()).ToDictionary(k => k.SourceType, k => Name(k.TypeName, k.TypeNameEn, english));

    /// <summary>The kinds of entries (JournalSourceTypes) in their order.</summary>
    public async Task<List<(string Key, string Name)>> EntryKindsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.JournalSourceTypes.AsNoTracking().OrderBy(k => k.SortOrder).ToListAsync();
        return rows.Select(k => (k.SourceType, Name(k.TypeName, k.TypeNameEn, english))).ToList();
    }

    /// <summary>One entry with its lines, their accounts and cost centres.</summary>
    public async Task<JournalEntryView?> EntryAsync(int entryId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var e = await db.JournalEntries.AsNoTracking().FirstOrDefaultAsync(x => x.EntryID == entryId);
        if (e is null)
        {
            return null;
        }
        var kinds = await KindsAsync(db, english);
        var lines = await (from l in db.JournalLines.AsNoTracking()
                           join a in db.Accounts on l.AccountCode equals a.AccountCode
                           join c in db.CostCenters on l.CostCenterID equals c.CostCenterID into cc
                           from c in cc.DefaultIfEmpty()
                           where l.EntryID == entryId
                           orderby l.LineNumber
                           select new { l, a.AccountName, a.AccountNameEn, Center = c == null ? null : c.CenterName, CenterEn = c == null ? null : c.CenterNameEn })
            .ToListAsync();
        var row = new JournalEntryRow(e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, kinds.TryGetValue(e.SourceType, out var k) ? k : e.SourceType,
                                      e.SourceNumber, e.Description, e.TotalDebit, e.TotalCredit, e.LineCount);
        return new JournalEntryView(row, e.CurrencyCode, e.ExchangeRate, e.ForeignAmount,
                                    lines.Select(x => new JournalLineView(x.l.LineNumber, x.l.AccountCode, Name(x.AccountName, x.AccountNameEn, english),
                                                                          x.l.Debit, x.l.Credit, x.l.LineText,
                                                                          x.Center is null ? null : Name(x.Center, x.CenterEn, english))).ToList());
    }

    // ------------------------------------------------------------------ accounts and the statement (frmLedger)

    public async Task<List<AccountChoice>> AccountsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Accounts.AsNoTracking().Where(a => a.IsActive).ToListAsync();
        return rows.OrderBy(a => a.TreeKey ?? a.AccountCode.ToString(), StringComparer.Ordinal)
            .Select(a => new AccountChoice(a.AccountCode, Name(a.AccountName, a.AccountNameEn, english), a.AccountType, a.IsPosting, a.AccountLevel))
            .ToList();
    }

    /// <summary>The account and its sub-accounts (Level1Code .. Level5Code = the account).</summary>
    private static async Task<HashSet<int>> AccountAndChildrenAsync(RetailDbContext db, int code) =>
        (await db.Accounts.AsNoTracking()
            .Where(a => a.Level1Code == code || a.Level2Code == code || a.Level3Code == code || a.Level4Code == code || a.Level5Code == code ||
                        a.AccountCode == code)
            .Select(a => a.AccountCode).ToListAsync()).ToHashSet();

    /// <summary>modLedger.FillLedger over AccountStatementQuery: the opening balance before the period, then every journal line of
    /// the account (a main account: of all its sub-accounts) in the period, ordered by date and entry number.</summary>
    public async Task<LedgerView?> LedgerAsync(int accountCode, DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var account = await db.Accounts.AsNoTracking().FirstOrDefaultAsync(a => a.AccountCode == accountCode);
        if (account is null)
        {
            return null;
        }
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var codes = await AccountAndChildrenAsync(db, accountCode);
        var codeList = codes.ToList();
        decimal opening = await (from l in db.JournalLines
                                 join e in db.JournalEntries on l.EntryID equals e.EntryID
                                 where codeList.Contains(l.AccountCode) && e.EntryDate < start
                                 select l.Debit - l.Credit).SumAsync();
        var rows = await (from l in db.JournalLines.AsNoTracking()
                          join e in db.JournalEntries on l.EntryID equals e.EntryID
                          join a in db.Accounts on l.AccountCode equals a.AccountCode
                          where codeList.Contains(l.AccountCode) && e.EntryDate >= start && e.EntryDate < end
                          select new { e.EntryID, e.EntryDate, e.EntryNumber, e.SourceType, e.SourceNumber, e.Description, a.AccountName, a.AccountNameEn,
                                       l.Debit, l.Credit, l.LineNumber })
            .ToListAsync();
        var kinds = await KindsAsync(db, english);
        var lines = new List<LedgerLine>
        {
            new(0, start, "-", Tr("رصيد أول المدة", english), null, null, null, opening > 0 ? opening : 0, opening < 0 ? -opening : 0, opening),
        };
        decimal balance = opening, debit = 0, credit = 0;
        foreach (var r in rows.OrderBy(r => r.EntryDate).ThenBy(r => r.EntryNumber, StringComparer.Ordinal).ThenBy(r => r.LineNumber))
        {
            balance += r.Debit - r.Credit;
            debit += r.Debit;
            credit += r.Credit;
            lines.Add(new LedgerLine(r.EntryID, r.EntryDate, r.EntryNumber, kinds.TryGetValue(r.SourceType, out var k) ? k : r.SourceType, r.SourceNumber,
                                     r.Description, Name(r.AccountName, r.AccountNameEn, english), r.Debit, r.Credit, balance));
        }
        return new LedgerView(account.AccountCode, Name(account.AccountName, account.AccountNameEn, english), !account.IsPosting, opening, debit, credit,
                              opening + debit - credit, rows.Count, lines);
    }

    // ------------------------------------------------------------------ the trial balance and the financial statements

    /// <summary>Sum of debit and credit per account of the entries in [start, end) (null = from the first entry), YEAR_CLOSE left out
    /// when asked (qryIncomeMoves / qryCompareMoves).</summary>
    private static async Task<Dictionary<int, (decimal Debit, decimal Credit)>> MovesAsync(RetailDbContext db, DateTime? start, DateTime end,
                                                                                          bool withoutYearClose = false)
    {
        var q = from l in db.JournalLines
                join e in db.JournalEntries on l.EntryID equals e.EntryID
                where e.EntryDate < end
                select new { l.AccountCode, l.Debit, l.Credit, e.EntryDate, e.SourceType };
        if (start is { } s)
        {
            q = q.Where(x => x.EntryDate >= s);
        }
        if (withoutYearClose)
        {
            q = q.Where(x => x.SourceType != "YEAR_CLOSE");
        }
        var rows = await q.GroupBy(x => x.AccountCode).Select(g => new { g.Key, D = g.Sum(x => x.Debit), C = g.Sum(x => x.Credit) }).ToListAsync();
        return rows.ToDictionary(r => r.Key, r => (r.D, r.C));
    }

    /// <summary>TrialBalanceQuery: each account with lines before or in the period: opening, period debit and credit, closing.</summary>
    public async Task<List<TrialBalanceRow>> TrialBalanceAsync(DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var before = await MovesAsync(db, null, start);
        var period = await MovesAsync(db, start, end);
        var accounts = await db.Accounts.AsNoTracking().Where(a => before.Keys.Contains(a.AccountCode) || period.Keys.Contains(a.AccountCode))
            .ToListAsync();
        return accounts.OrderBy(a => a.AccountCode).Select(a =>
        {
            var b = before.TryGetValue(a.AccountCode, out var x) ? x : (0m, 0m);
            var p = period.TryGetValue(a.AccountCode, out var y) ? y : (0m, 0m);
            decimal opening = b.Item1 - b.Item2;
            return new TrialBalanceRow(a.AccountCode, Name(a.AccountName, a.AccountNameEn, english), a.AccountType, opening, p.Item1, p.Item2,
                                       opening + p.Item1 - p.Item2);
        }).ToList();
    }

    /// <summary>IncomeStatementQuery (over qryIncomeAccounts): revenue, cost of sales, gross profit, operating expenses, operating
    /// profit, other revenue and expenses, net profit, for the period and the comparison period (YEAR_CLOSE entries left out).</summary>
    public async Task<StatementView> IncomeStatementAsync(DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var (cFrom, cTo) = ComparePeriod(from, to);
        var current = await MovesAsync(db, from.Date, to.Date.AddDays(1), true);
        var prior = await MovesAsync(db, cFrom, cTo.AddDays(1), true);
        var accounts = await db.Accounts.AsNoTracking().Where(a => a.AccountType == "REVENUE" || a.AccountType == "EXPENSE").ToListAsync();
        var items = new List<(int Block, string Key, StatementRow Row, int Section, decimal Cur, decimal Pri)>();
        foreach (var a in accounts.Where(a => current.ContainsKey(a.AccountCode) || prior.ContainsKey(a.AccountCode)))
        {
            int section = a.Level2Code == 41 ? 1 : a.Level2Code == 51 ? 2 : a.Level2Code == 52 ? 3 : a.AccountType == "REVENUE" ? 4 : 5;
            int sign = a.AccountType == "REVENUE" ? 1 : -1;
            var c = current.TryGetValue(a.AccountCode, out var x) ? x : (0m, 0m);
            var p = prior.TryGetValue(a.AccountCode, out var y) ? y : (0m, 0m);
            decimal cur = sign * (c.Item2 - c.Item1), pri = sign * (p.Item2 - p.Item1);
            items.Add((section * 10 + 1, a.TreeKey ?? "", new StatementRow("A", Name(a.AccountName, a.AccountNameEn, english), a.AccountCode, cur, pri),
                       section, cur, pri));
        }
        decimal Sum(int section, bool cur) => items.Where(i => i.Section == section).Sum(i => cur ? i.Cur : i.Pri);
        decimal Result(int[] sections, int[] signs, bool cur) => sections.Select((s, i) => signs[i] * Sum(s, cur)).Sum();
        var rows = new List<(int Block, string Key, StatementRow Row)>();
        rows.AddRange(items.Select(i => (i.Block, i.Key, i.Row)));
        (int Block, string Caption)[] heads = [(10, "إيرادات النشاط"), (20, "تكلفة المبيعات"), (30, "المصروفات التشغيلية والإدارية"),
                                               (40, "إيرادات أخرى"), (50, "مصروفات أخرى")];
        foreach (var (block, caption) in heads)
        {
            rows.Add((block, "", new StatementRow("H", Tr(caption, english), null, null, null)));
        }
        (int Block, string Caption, int Section)[] totals = [(12, "صافي إيرادات النشاط", 1), (22, "إجمالي تكلفة المبيعات", 2),
                                                             (32, "إجمالي المصروفات التشغيلية والإدارية", 3), (42, "إجمالي الإيرادات الأخرى", 4),
                                                             (52, "إجمالي المصروفات الأخرى", 5)];
        foreach (var (block, caption, section) in totals)
        {
            rows.Add((block, "", new StatementRow("T", Tr(caption, english), null, Sum(section, true), Sum(section, false))));
        }
        (int Block, string Caption, int[] Sections, int[] Signs)[] results =
        [
            (25, "مجمل الربح", [1, 2], [1, -1]),
            (35, "الربح التشغيلي", [1, 2, 3], [1, -1, -1]),
            (60, "صافي الربح (الخسارة)", [1, 2, 3, 4, 5], [1, -1, -1, 1, -1]),
        ];
        foreach (var r in results)
        {
            rows.Add((r.Block, "", new StatementRow("R", Tr(r.Caption, english), null, Result(r.Sections, r.Signs, true), Result(r.Sections, r.Signs, false))));
        }
        var ordered = rows.OrderBy(r => r.Block).ThenBy(r => r.Key, StringComparer.Ordinal).Select(r => r.Row).ToList();
        var figures = new List<(string, decimal)>
        {
            (Tr("صافي إيرادات النشاط", english), Sum(1, true)), (Tr("مجمل الربح", english), Result(results[0].Sections, results[0].Signs, true)),
            (Tr("صافي الربح (الخسارة)", english), Result(results[2].Sections, results[2].Signs, true)),
        };
        return new StatementView("INCOME", from.Date, to.Date, cFrom, cTo, ordered, figures);
    }

    /// <summary>BalanceSheetQuery (over qryBalanceAccounts / qryBalanceItems): assets = liabilities + equity at the end of the
    /// period, the profit of the periods not closed yet in equity (group 32), with the comparison at the end of the comparison
    /// period.</summary>
    public async Task<StatementView> BalanceSheetAsync(DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var (cFrom, cTo) = ComparePeriod(from, to);
        var at = await MovesAsync(db, null, to.Date.AddDays(1));
        var compare = await MovesAsync(db, null, cTo.AddDays(1));
        var accounts = await db.Accounts.AsNoTracking().ToListAsync();
        var byCode = accounts.ToDictionary(a => a.AccountCode);
        decimal Net(Dictionary<int, (decimal Debit, decimal Credit)> m, int code) => m.TryGetValue(code, out var x) ? x.Debit - x.Credit : 0;
        string Cap(Entities.Accounts a) => Name(a.AccountName, a.AccountNameEn, english);
        string Key(int code) => byCode.TryGetValue(code, out var a) ? a.TreeKey ?? "" : "";

        var rows = new List<(int Class, string Group, int Pos, string Key, StatementRow Row)>();
        var items = new List<(int Class, int Group, decimal Cur, decimal Pri)>();
        foreach (var a in accounts.Where(a => a.AccountType is "ASSET" or "LIABILITY" or "EQUITY"))
        {
            decimal cur = Net(at, a.AccountCode), pri = Net(compare, a.AccountCode);
            if (cur == 0 && pri == 0)
            {
                continue;
            }
            int sign = a.AccountType == "ASSET" ? 1 : -1;
            items.Add((a.Level1Code ?? 0, a.Level2Code ?? 0, sign * cur, sign * pri));
            if (a.Level2Code is { } g && byCode.ContainsKey(g))
            {
                rows.Add((a.Level1Code ?? 0, Key(g), 1, a.TreeKey ?? "", new StatementRow("A", Cap(a), a.AccountCode, sign * cur, sign * pri)));
            }
        }
        // the profit of the periods not closed yet (qryProfitAt / qryProfitCompare), in equity under account 32
        var pl = accounts.Where(a => a.AccountType is "REVENUE" or "EXPENSE").Select(a => a.AccountCode).ToList();
        decimal profit = -pl.Sum(c => Net(at, c)), profitPrior = -pl.Sum(c => Net(compare, c));
        items.Add((3, 32, profit, profitPrior));
        if (byCode.ContainsKey(32))
        {
            rows.Add((3, Key(32), 1, "Z", new StatementRow("A", Tr("صافي ربح (خسارة) الفترات غير المقفلة", english), null, profit, profitPrior)));
        }
        foreach (int c in new[] { 1, 2, 3 }.Where(byCode.ContainsKey))
        {
            rows.Add((c, "", 0, "", new StatementRow("C", Cap(byCode[c]), null, null, null)));
        }
        foreach (var g in items.GroupBy(i => i.Group).Where(g => byCode.ContainsKey(g.Key)))
        {
            var ga = byCode[g.Key];
            rows.Add((ga.Level1Code ?? 0, ga.TreeKey ?? "", 0, "", new StatementRow("G", Cap(ga), null, null, null)));
            rows.Add((ga.Level1Code ?? 0, ga.TreeKey ?? "", 2, "", new StatementRow("S", Cap(ga), null, g.Sum(i => i.Cur), g.Sum(i => i.Pri))));
        }
        foreach (var c in items.GroupBy(i => i.Class))
        {
            string caption = c.Key == 1 ? "إجمالي الأصول" : c.Key == 2 ? "إجمالي الخصوم" : "إجمالي حقوق الملكية";
            rows.Add((c.Key, "~", 9, "", new StatementRow("T", Tr(caption, english), null, c.Sum(i => i.Cur), c.Sum(i => i.Pri))));
        }
        decimal le = items.Where(i => i.Class is 2 or 3).Sum(i => i.Cur), lePrior = items.Where(i => i.Class is 2 or 3).Sum(i => i.Pri);
        rows.Add((4, "", 9, "", new StatementRow("T", Tr("إجمالي الخصوم وحقوق الملكية", english), null, le, lePrior)));
        var ordered = rows.OrderBy(r => r.Class).ThenBy(r => r.Group, StringComparer.Ordinal).ThenBy(r => r.Pos).ThenBy(r => r.Key, StringComparer.Ordinal)
            .Select(r => r.Row).ToList();
        decimal assets = items.Where(i => i.Class == 1).Sum(i => i.Cur);
        var figures = new List<(string, decimal)>
        {
            (Tr("إجمالي الأصول", english), assets), (Tr("إجمالي الخصوم وحقوق الملكية", english), le), (Tr("الفرق (صفر = متوازنة)", english), assets - le),
        };
        return new StatementView("BALANCE", from.Date, to.Date, cFrom, cTo, ordered, figures);
    }
}
