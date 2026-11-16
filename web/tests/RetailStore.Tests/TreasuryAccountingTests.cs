using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Data.Services;
using RetailStore.Domain.Journal;

namespace RetailStore.Tests;

/// <summary>The treasury, the banks, the accounting and the dashboard of the web in C# (plan 6E, docs/81): modCash, modBank,
/// modLedger, modFinancials and modDashboard.</summary>
public class TreasuryAccountingTests
{
    private static readonly DateTime Today = DateTime.Today;

    private sealed class Store : IDisposable
    {
        public TestDb Db { get; } = new();
        public TreasuryService Treasury => new(Db, Db.Permissions);
        public AccountingService Accounting => new(Db, new JournalService(Db, new FixedJournalReader(new Dictionary<string, List<JournalSourceLine>>())));
        public DashboardService Dashboard => new(Db, Db.Permissions);
        public int Customer, Bank, Bank2;

        public Store()
        {
            using var db = Db.CreateDbContext();
            var customer = new Customers { CustomerName = "عميل الخزينة" };
            var bank = new Banks { BankName = "بنك الرياض" };
            var bank2 = new Banks { BankName = "البنك الأهلي" };
            db.AddRange(customer, bank, bank2);
            db.SaveChanges();
            (Customer, Bank, Bank2) = (customer.CustomerID, bank.BankID, bank2.BankID);
            // 500 in the main safe from a customer receipt, three days ago
            db.CustomerPayments.Add(new CustomerPayments { PaymentNumber = "RCV-1", CustomerID = Customer, PaymentDate = Today.AddDays(-3), Amount = 500,
                                                           PaymentMethodID = 1, EmployeeID = 1, CashBoxID = 1 });
            db.SaveChanges();
        }

        public void Dispose() => Db.Dispose();
    }

    private static async Task<int> UserWithAsync(TestDb t, params string[] permissions)
    {
        await using var db = t.CreateDbContext();
        var role = new Roles { RoleCode = "T" + permissions.Length, RoleName = "دور" };
        db.Roles.Add(role);
        await db.SaveChangesAsync();
        foreach (var p in permissions)
        {
            db.RolePermissions.Add(new RolePermissions { RoleID = role.RoleID, PermissionKey = p });
        }
        var e = new Employees { EmployeeName = "أمين", Username = "u" + role.RoleID, RoleID = role.RoleID };
        db.Employees.Add(e);
        await db.SaveChangesAsync();
        return e.EmployeeID;
    }

    [Fact]
    public async Task Vouchers_move_the_cash_with_the_checks_of_access()
    {
        using var s = new Store();
        var boxes = await s.Treasury.BoxBalancesAsync(false);
        Assert.Equal(("الخزينة الرئيسية", 500m), (boxes[0].BoxName, boxes[0].Balance));       // MAIN before CASHIER
        Assert.Equal("أدخل المبلغ (أكبر من صفر).",
                     (await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("IN", 1, null, "OTHER", 0, "", "", null))).Message);
        Assert.Equal("اختر بند القبض.", (await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("IN", 1, null, "ADVANCE", 5, "", "", null))).Message);
        Assert.Equal("اختر نوع المصروف.", (await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("OUT", 1, null, "EXPENSE", 5, "", "", null))).Message);
        Assert.Equal("لا يمكن التحويل إلى نفس الصندوق.",
                     (await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("TRANSFER", 1, 1, null, 5, "", "", null))).Message);
        Assert.Equal("رصيد الخزينة الرئيسية (500.00) لا يكفي لصرف 600.00.",
                     (await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("OUT", 1, null, "OTHER", 600, "", "", null))).Message);

        var r = await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("IN", 1, null, "OWNER", 200, "المالك", "إيداع", null));
        Assert.Equal((SaveStatus.Saved, "CIN-000001"), (r.Status, r.Number));
        r = await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("OUT", 1, null, "EXPENSE", 150, "", "إيجار", 1));
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using (var db = s.Db.CreateDbContext())
        {
            var v = await db.CashVouchers.SingleAsync(x => x.CashVoucherID == r.Id);
            var expense = await db.Expenses.SingleAsync(x => x.ExpenseID == v.ExpenseID);
            Assert.Equal((150m, (int?)null, "سند صرف نقدية " + v.VoucherNumber + " - إيجار"), (expense.Amount, expense.CashBoxID, expense.Description));
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "CASH_OUT" && a.Details == "EXPENSE Amount=150"));
        }
        // 500 + 200 - 150: the expense voucher moves the cash once (the expense has no box)
        Assert.Equal(550m, (await s.Treasury.BoxBalancesAsync(false))[0].Balance);
        var statement = await s.Treasury.BoxStatementAsync(1, Today, Today, false);
        Assert.Equal(("رصيد أول المدة", 500m), (statement[0].MoveTypeName, statement[0].Balance));
        Assert.Equal([700m, 550m], statement.Skip(1).Select(l => l.Balance));
        Assert.Equal(["إيداع من المالك", "مصروف (سند صرف)"], statement.Skip(1).Select(l => l.MoveTypeName));
        Assert.Equal("Deposit from the owner", (await s.Treasury.BoxStatementAsync(1, Today, Today, true))[1].MoveTypeName);
        Assert.Equal(2, (await s.Treasury.VouchersAsync(Today, Today, 1, false, false)).Count);
    }

    [Fact]
    public async Task A_transfer_to_a_box_that_confirms_receipt_waits_for_its_custodian()
    {
        using var s = new Store();
        int keeper = await UserWithAsync(s.Db, "CASH_BOX", "CASH_CLOSING");
        await using (var db = s.Db.CreateDbContext())
        {
            (await db.CashBoxes.FindAsync(2))!.ConfirmReceipt = true;
            db.CashBoxUsers.Add(new CashBoxUsers { CashBoxID = 2, EmployeeID = keeper });
            await db.SaveChangesAsync();
        }
        var r = await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("TRANSFER", 1, 2, null, 100, "", "", null));
        Assert.Equal(SaveStatus.Saved, r.Status);
        var boxes = await s.Treasury.BoxBalancesAsync(false);
        Assert.Equal((400m, 0m), (boxes[0].Balance, boxes[1].Balance));                       // in transit
        Assert.Single(await s.Treasury.VouchersAsync(Today, Today, 0, true, false));
        // the keeper of box 2 cannot pay from the main safe (it has no keeper: anyone may) but cannot cancel ...
        var keeperUser = new ActingUser(keeper, "TEST");
        Assert.Equal(SaveStatus.Saved, (await s.Treasury.ReceiveTransferAsync(keeperUser, r.Id)).Status);
        Assert.Equal("هذا التحويل ليس بانتظار الاستلام.", (await s.Treasury.ReceiveTransferAsync(keeperUser, r.Id)).Message);
        boxes = await s.Treasury.BoxBalancesAsync(false);
        Assert.Equal((400m, 100m), (boxes[0].Balance, boxes[1].Balance));
        // box 2 has a keeper: another user without CASH_ALL_BOXES may not pay from it
        int other = await UserWithAsync(s.Db, "CASH_BOX");
        var refused = await s.Treasury.PostCashVoucherAsync(new ActingUser(other, "TEST"), new("OUT", 2, null, "OTHER", 10, "", "", null));
        Assert.Equal("لست من أمناء صندوق الكاشير: لا يمكنك الصرف أو التحويل منه.", refused.Message);
        // a second transfer taken back before its receipt
        var second = await s.Treasury.PostCashVoucherAsync(TestDb.Admin, new("TRANSFER", 1, 2, null, 50, "", "", null));
        Assert.Equal(SaveStatus.Saved, (await s.Treasury.CancelTransferAsync(TestDb.Admin, second.Id)).Status);
        await using (var db = s.Db.CreateDbContext())
        {
            Assert.False(await db.CashVouchers.AnyAsync(v => v.CashVoucherID == second.Id));
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "CASH_TRANSFER_RECEIVE" && a.RecordID == r.Id.ToString()));
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "CASH_TRANSFER_DELETE" && a.RecordID == second.Id.ToString()));
        }
    }

    [Fact]
    public async Task Bank_transactions_follow_bank_tx_problem()
    {
        using var s = new Store();
        BankTxRequest Tx(string type, decimal amount, DateTime? d = null) =>
            new(type, d ?? Today, s.Bank, s.Bank2, 1, 4200, amount, 0, 0, "", "");
        Assert.Equal("تاريخ الحركة بعد اليوم.", (await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("DEPOSIT", 10, Today.AddDays(1)))).Message);
        Assert.Equal("رصيد الصندوق في " + StockService.GDate(Today) + " أقل من المبلغ المودَع.",
                     (await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("DEPOSIT", 900))).Message);
        Assert.Equal("التحويل بين بنكين مختلفين.",
                     (await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("TRANSFER", 10) with { ToBankID = s.Bank })).Message);
        Assert.StartsWith("الحساب المقابل حساب فرعي نشط",
                          (await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("OTHER_IN", 10) with { CounterAccount = 110001 })).Message);
        Assert.Equal("العمولة وضريبتها أقل من مبلغ التسوية.",
                     (await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("SETTLEMENT", 10) with { FeeAmount = 8, FeeVAT = 2 })).Message);

        var r = await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("DEPOSIT", 300));
        Assert.Equal((SaveStatus.Saved, "BNK-000001"), (r.Status, r.Number));
        var row = (await s.Treasury.BankTxAsync(Today, Today, 0, false)).Single();
        Assert.Equal(("إيداع نقدية في البنك", "الخزينة الرئيسية", "إيداع نقدية في البنك"), (row.TxTypeName, row.OtherName, row.Description));
        Assert.Equal(200m, (await s.Treasury.BoxBalancesAsync(false))[0].Balance);         // the deposit left the box
        await using (var db = s.Db.CreateDbContext())
        {
            // the book balance of a bank = its journal lines (Access builds the entry); a cleared transaction is not deleted
            await JournalService.EnsureAccountsAsync(db);                      // 120000 + bank
            var e = new JournalEntries { EntryNumber = "JV-1", EntryDate = Today, SourceType = "BANK_TX", SourceID = r.Id, TotalDebit = 300, TotalCredit = 300 };
            db.JournalEntries.Add(e);
            await db.SaveChangesAsync();
            db.JournalLines.AddRange(new JournalLines { EntryID = e.EntryID, LineNumber = 1, AccountCode = 120000 + s.Bank, Debit = 300 },
                                     new JournalLines { EntryID = e.EntryID, LineNumber = 2, AccountCode = 110001, Credit = 300 });
            var recon = new BankReconciliations { ReconNumber = "REC-1", BankID = s.Bank, EmployeeID = 1 };
            db.BankReconciliations.Add(recon);
            await db.SaveChangesAsync();
            db.BankClearings.Add(new BankClearings { BankID = s.Bank, SourceType = "BANK_TX", SourceID = r.Id, ClearedAmount = 300,
                                                     ReconciliationID = recon.ReconciliationID });
            await db.SaveChangesAsync();
        }
        var banks = await s.Treasury.BankBalancesAsync(false);
        Assert.Equal(300m, banks.Single(b => b.BankID == s.Bank).Balance);
        Assert.Equal("الحركة مطابقة في تسوية بنكية: ألغِ مطابقتها أولًا.", (await s.Treasury.DeleteBankTxAsync(TestDb.Admin, r.Id)).Message);
        var other = await s.Treasury.PostBankTxAsync(TestDb.Admin, Tx("OTHER_OUT", 50) with { CounterAccount = 5610, FeeVAT = 7.5m });
        Assert.Equal(SaveStatus.Saved, other.Status);
        Assert.Equal(SaveStatus.Saved, (await s.Treasury.DeleteBankTxAsync(TestDb.Admin, other.Id)).Status);
        await using (var db = s.Db.CreateDbContext())
        {
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "BANK_TX_DELETE" && a.RecordID == other.Id.ToString()));
        }
        int cashier = await UserWithAsync(s.Db, "CASH_BOX");
        Assert.Equal("لا تملك صلاحية البنوك.", (await s.Treasury.PostBankTxAsync(new ActingUser(cashier, "TEST"), Tx("DEPOSIT", 1))).Message);
    }

    /// <summary>A small business in the journal: capital 1,000 in the safe; a sale of 500 (cost 300); rent 100.</summary>
    private static async Task AddEntriesAsync(TestDb t, DateTime month)
    {
        await using var db = t.CreateDbContext();
        int n = 0;
        async Task Entry(DateTime d, string type, params (int Account, decimal Debit, decimal Credit)[] lines)
        {
            var e = new JournalEntries { EntryNumber = "JV-" + (++n).ToString("000000"), EntryDate = d, SourceType = type, SourceID = n,
                                         SourceNumber = "D-" + n, Description = "قيد " + n, TotalDebit = lines.Sum(l => l.Debit),
                                         TotalCredit = lines.Sum(l => l.Credit), LineCount = (short)lines.Length };
            db.JournalEntries.Add(e);
            await db.SaveChangesAsync();
            short i = 0;
            foreach (var l in lines)
            {
                db.JournalLines.Add(new JournalLines { EntryID = e.EntryID, LineNumber = ++i, AccountCode = l.Account, Debit = l.Debit, Credit = l.Credit });
            }
            await db.SaveChangesAsync();
        }
        await Entry(month.AddMonths(-1), "MANUAL", (110001, 1000, 0), (3200, 0, 1000));
        await Entry(month, "SALE", (110001, 500, 0), (4100, 0, 500));
        await Entry(month, "SALE", (5100, 300, 0), (1400, 0, 300));
        await Entry(month.AddDays(1), "EXPENSE", (530001, 100, 0), (110002, 0, 100));
    }

    [Fact]
    public async Task The_ledger_the_trial_balance_and_the_statements_are_those_of_access()
    {
        using var s = new Store();
        var month = new DateTime(Today.Year, Today.Month, 1);
        await using (var db = s.Db.CreateDbContext())
        {
            await JournalService.EnsureAccountsAsync(db);         // 530001 = rent
        }
        await AddEntriesAsync(s.Db, month);
        var entries = await s.Accounting.EntriesAsync(month, month.AddMonths(1), null, null, false);
        Assert.Equal(3, entries.Count);
        var entry = (await s.Accounting.EntryAsync(entries.Last().EntryID, false))!;
        Assert.Equal(("فاتورة بيع", 2, "الخزينة الرئيسية"), (entry.Entry.KindName, entry.Lines.Count, entry.Lines[0].AccountName));

        // the statement of the main account 1100 takes its boxes 110001 and 110002
        var ledger = (await s.Accounting.LedgerAsync(1100, month, month.AddMonths(1), false))!;
        Assert.Equal((true, 1000m, 500m, 100m, 1400m, 2), (ledger.IsMain, ledger.Opening, ledger.Debit, ledger.Credit, ledger.Closing, ledger.LineCount));
        Assert.Equal(["رصيد أول المدة", "فاتورة بيع", "مصروف"], ledger.Lines.Select(l => l.KindName));
        Assert.Equal([1000m, 1500m, 1400m], ledger.Lines.Select(l => l.Balance));
        Assert.Equal(("1,400.00 مدين", "300.00 دائن", "0.00"),
                     (AccountingService.BalanceText(1400), AccountingService.BalanceText(-300), AccountingService.BalanceText(0)));

        var trial = await s.Accounting.TrialBalanceAsync(month, month.AddMonths(1), false);
        Assert.Equal(trial.Sum(r => r.PeriodDebit), trial.Sum(r => r.PeriodCredit));
        Assert.Equal(0m, trial.Sum(r => r.ClosingBalance));
        Assert.Equal((1000m, 1500m), (trial.Single(r => r.AccountCode == 110001).OpeningBalance, trial.Single(r => r.AccountCode == 110001).ClosingBalance));

        var income = await s.Accounting.IncomeStatementAsync(month, Today, false);
        Assert.Equal(["صافي إيرادات النشاط", "مجمل الربح", "صافي الربح (الخسارة)"], income.Figures.Select(f => f.Caption));
        Assert.Equal([500m, 200m, 100m], income.Figures.Select(f => f.Amount));
        Assert.Equal("إيرادات النشاط", income.Rows[0].Caption);                                   // block 10 first
        Assert.Equal(("A", "المبيعات", 500m), (income.Rows[1].RowKind, income.Rows[1].Caption, income.Rows[1].CurrentValue!.Value));
        Assert.Equal((month.AddMonths(-1), month.AddMonths(-1).AddDays((Today - month).Days)), (income.CompareFrom, income.CompareTo));

        var balance = await s.Accounting.BalanceSheetAsync(month, Today, false);
        Assert.Equal([1400m - 300m, 1100m, 0m], balance.Figures.Select(f => f.Amount));
        var profit = balance.Rows.Single(r => r.Caption == "صافي ربح (خسارة) الفترات غير المقفلة");
        Assert.Equal((100m, 0m), (profit.CurrentValue!.Value, profit.PriorValue!.Value));
        Assert.Equal("الأصول", balance.Rows[0].Caption);
        Assert.Equal("Total assets", (await s.Accounting.BalanceSheetAsync(month, Today, true)).Figures[0].Caption);

        Assert.Equal("تاريخ النهاية قبل تاريخ البداية.", AccountingService.PeriodProblem(Today, Today.AddDays(-1)));
        // modQueryParams.SetComparePeriod
        Assert.Equal((new DateTime(2025, 10, 1), new DateTime(2025, 12, 16)),
                     AccountingService.ComparePeriod(new DateTime(2026, 1, 1), new DateTime(2026, 3, 16)));
        Assert.Equal((new DateTime(2026, 2, 1), new DateTime(2026, 2, 28)),
                     AccountingService.ComparePeriod(new DateTime(2026, 3, 1), new DateTime(2026, 3, 31)));
        Assert.Equal((new DateTime(2026, 3, 5), new DateTime(2026, 3, 9)),
                     AccountingService.ComparePeriod(new DateTime(2026, 3, 10), new DateTime(2026, 3, 14)));
    }

    [Fact]
    public async Task The_dashboard_shows_the_figures_of_dashboard_query_and_hides_the_amounts_without_the_permission()
    {
        using var s = new Store();
        var sales = new SalesService(s.Db, s.Db.Permissions);
        await using (var db = s.Db.CreateDbContext())
        {
            db.Products.Add(new Products { ProductCode = "D1", ProductName = "صنف", SellingPrice = 115, AverageCost = 60, CurrentQuantity = 3,
                                           MinimumQuantity = 5 });
            db.Expenses.Add(new Expenses { ExpenseNumber = "EXP-1", ExpenseDate = Today, ExpenseTypeID = 1, Amount = 40, TotalAmount = 40, EmployeeID = 1 });
            (await db.Customers.FindAsync(s.Customer))!.CurrentBalance = 75;
            await db.SaveChangesAsync();
        }
        int product;
        await using (var db = s.Db.CreateDbContext())
        {
            product = (await db.Products.SingleAsync(p => p.ProductCode == "D1")).ProductID;
        }
        var sale = await sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Customer, "CASH", 1, 0, null, "",
                                                                         [new CartLine(product, 1, 115)]));
        Assert.Equal(SaveStatus.Saved, sale.Status);
        var f = await s.Dashboard.FiguresAsync(1, Today);
        Assert.Equal((115m, 1, 115m, 15m, 1), (f.TodaySales!.Value, f.TodayInvoices, f.MonthSales!.Value, f.MonthVAT!.Value, f.MonthInvoices));
        // net sales 100 - cost 60 - expenses 40
        Assert.Equal((0m, 1, 40m, 120m), (f.MonthNetProfit!.Value, f.LowStockCount, f.MonthExpenses!.Value, f.StockValue!.Value));
        Assert.Equal((75m, 1), (f.CustomerDebt!.Value, f.DebtorCount));
        Assert.Equal(500m + 115m, f.CashBalance!.Value);
        int cashier = await UserWithAsync(s.Db, "SALES_POS");
        var hidden = await s.Dashboard.FiguresAsync(cashier, Today);
        Assert.Equal((false, (decimal?)null, 1, (decimal?)null), (hidden.Financial, hidden.TodaySales, hidden.TodayInvoices, hidden.MonthNetProfit));
    }

    [Fact]
    public async Task The_api_posts_vouchers_and_bank_transactions_and_reads_the_accounting()
    {
        await using var host = new ApiHost();
        var client = await host.AdminAsync();
        var r = await client.PostAsJsonAsync("/api/v1/cash-vouchers", new { voucherType = "IN", cashBoxID = 1, category = "OTHER", amount = 80 });
        Assert.Equal(HttpStatusCode.Created, r.StatusCode);
        Assert.StartsWith("CIN-", (await r.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("number").GetString());
        r = await client.PostAsJsonAsync("/api/v1/cash-vouchers", new { voucherType = "OUT", cashBoxID = 1, category = "OTHER", amount = 900 });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, r.StatusCode);
        var boxes = await client.GetFromJsonAsync<JsonElement>("/api/v1/treasury/boxes");
        Assert.Equal(80m, boxes[0].GetProperty("balance").GetDecimal());
        var lines = await client.GetFromJsonAsync<JsonElement>("/api/v1/treasury/boxes/1/statement");
        Assert.Equal(2, lines.GetArrayLength());
        await using (var db = host.Db.CreateDbContext())
        {
            db.Banks.Add(new Banks { BankName = "بنك" });
            await db.SaveChangesAsync();
        }
        r = await client.PostAsJsonAsync("/api/v1/bank-transactions", new { txType = "DEPOSIT", bankID = 1, cashBoxID = 1, amount = 50 });
        Assert.Equal(HttpStatusCode.Created, r.StatusCode);
        Assert.Equal(1, (await client.GetFromJsonAsync<JsonElement>("/api/v1/bank-transactions")).GetArrayLength());
        Assert.Equal(1, (await client.GetFromJsonAsync<JsonElement>("/api/v1/banks")).GetArrayLength());
        var request = new HttpRequestMessage(HttpMethod.Post, "/api/v1/bank-transactions")
        {
            Content = JsonContent.Create(new { txType = "TRANSFER", bankID = 1, toBankID = 1, amount = 5 }),
        };
        request.Headers.AcceptLanguage.ParseAdd("en");
        var problem = await (await client.SendAsync(request)).Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("Transfer between two different banks.", problem.GetProperty("title").GetString());
        Assert.Equal(HttpStatusCode.OK, (await client.GetAsync("/api/v1/accounting/trial-balance")).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.GetAsync("/api/v1/accounting/income-statement")).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.GetAsync("/api/v1/accounting/balance-sheet")).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.GetAsync("/api/v1/accounting/ledger/1100")).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.GetAsync("/api/v1/journal-entries")).StatusCode);
        var dash = await client.GetFromJsonAsync<JsonElement>("/api/v1/dashboard");
        Assert.True(dash.GetProperty("financial").GetBoolean());
    }
}
