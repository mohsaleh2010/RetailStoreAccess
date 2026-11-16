using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>The treasury, the banks, the accounting and the dashboard (plan 6E, docs/81): the same services as the screens of
/// the web (modCash, modBank, modLedger, modFinancials, modDashboard in Access).</summary>
public static class TreasuryApi
{
    private static IResult Done(HttpContext http, TreasuryResult r, string location) => r.Status switch
    {
        SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + location + r.Id, new PostedDocument(r.Id, r.Number, 0, 0)),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };

    private static IResult Changed(HttpContext http, TreasuryResult r) =>
        r.Status == SaveStatus.Saved ? Results.NoContent() : ApiResults.Refused(http, r.Message);

    private static DateTime MonthStart => DateTime.Today.AddDays(1 - DateTime.Today.Day);

    private static DateTime YearStart => new(DateTime.Today.Year, 1, 1);

    private static async Task<IResult> Period(HttpContext http, DateTime from, DateTime to, Func<Task<object>> read)
    {
        string p = AccountingService.PeriodProblem(from, to);
        return p.Length > 0 ? ApiResults.Refused(http, p) : Results.Ok(await read());
    }

    public static void Map(RouteGroupBuilder secured)
    {
        // ------------------------------------------------------------------ treasury
        var treasury = secured.MapGroup("/treasury").WithTags("Treasury");

        treasury.MapGet("/boxes", async (HttpContext http, TreasuryService service) =>
            Results.Ok(await service.BoxBalancesAsync(ApiResults.English(http))))
        .Needs("CASH_CLOSING", "frmTreasury").WithName("ListCashBoxes")
        .WithSummary("Every box with its cash in, out and balance now (CashBoxBalanceQuery)")
        .Produces<List<BoxBalance>>();

        treasury.MapGet("/boxes/{id:int}/statement", async (int id, HttpContext http, TreasuryService service, DateTime? from, DateTime? to) =>
        {
            var (f, t) = (from ?? DateTime.Today, to ?? DateTime.Today);
            return await Period(http, f, t, async () => await service.BoxStatementAsync(id, f, t, ApiResults.English(http)));
        })
        .Needs("CASH_CLOSING", "frmTreasury").WithName("CashBoxStatement")
        .WithSummary("The movements of a box in a period (CashStatementQuery): the opening line, then each movement with the balance after it")
        .WithDescription("Box 0 = every box; the period defaults to today. Every cash movement of qryCashMovements: sales, returns, " +
                         "receipts, purchases, payments, expenses, cash vouchers, transfers (a transfer waiting for its receipt is not in " +
                         "the receiving box yet), bank deposits and withdrawals, payroll, fixed assets and the opening balance.")
        .Produces<List<BoxStatementLine>>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        var vouchers = secured.MapGroup("/cash-vouchers").WithTags("Treasury");

        vouchers.MapPost("", async (CashVoucherRequest body, HttpContext http, TreasuryService service) =>
            Done(http, await service.PostCashVoucherAsync(UserInfo.Acting(http.User), body), "/cash-vouchers/"))
        .Needs("CASH_BOX", "frmCashVoucher").WithName("PostCashVoucher")
        .WithSummary("A cash voucher: receipt, payment or transfer between boxes (modCash.PostCashVoucher)")
        .WithDescription("VoucherType IN (Category OTHER or OWNER), OUT (Category EXPENSE with ExpenseTypeID, OWNER or OTHER) or " +
                         "TRANSFER (ToCashBoxID). A payment or a transfer needs a box the user may use (its custodian, a box nobody " +
                         "keeps, or CASH_ALL_BOXES) and its balance. An expense voucher also records the expense. A transfer to a box " +
                         "that confirms receipt waits in transit until its custodian receives it. Numbers CIN- / COUT- / TRF-.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        vouchers.MapGet("", async (HttpContext http, TreasuryService service, DateTime? from, DateTime? to, int? cashBoxId, bool? inTransit) =>
            Results.Ok(await service.VouchersAsync(from ?? DateTime.Today, to ?? DateTime.Today, cashBoxId ?? 0, inTransit == true, ApiResults.English(http))))
        .Needs("CASH_BOX", "frmCashVoucher").WithName("ListCashVouchers")
        .WithSummary("The cash vouchers between two days (default today), newest first; inTransit=true: the transfers waiting for their receipt")
        .Produces<List<CashVoucherRow>>();

        vouchers.MapPost("/{id:int}/receive", async (int id, HttpContext http, TreasuryService service) =>
            Changed(http, await service.ReceiveTransferAsync(UserInfo.Acting(http.User), id)))
        .Needs("CASH_CLOSING", "frmCashTransfers").WithName("ReceiveCashTransfer")
        .WithSummary("A custodian of the receiving box confirms the cash of a transfer arrived (modCash.ReceiveCashTransfer)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        vouchers.MapPost("/{id:int}/cancel", async (int id, HttpContext http, TreasuryService service) =>
            Changed(http, await service.CancelTransferAsync(UserInfo.Acting(http.User), id)))
        .Needs("CASH_BOX", "frmCashTransfers").WithName("CancelCashTransfer")
        .WithSummary("The sender takes back a transfer not received yet: the voucher is deleted (modCash.CancelCashTransfer)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ------------------------------------------------------------------ banks
        secured.MapGet("/banks", async (HttpContext http, TreasuryService service) =>
            Results.Ok(await service.BankBalancesAsync(ApiResults.English(http))))
        .Needs("BANKS", "frmBankTx").WithTags("Banks").WithName("ListBankBalances")
        .WithSummary("Every bank with its balance in the books (BankBalanceQuery: its journal lines, account 120000 + bank)")
        .Produces<List<BankBalance>>();

        var bankTx = secured.MapGroup("/bank-transactions").WithTags("Banks");

        bankTx.MapPost("", async (BankTxRequest body, HttpContext http, TreasuryService service) =>
            Done(http, await service.PostBankTxAsync(UserInfo.Acting(http.User), body with { TxDate = body.TxDate ?? DateTime.Today }),
                 "/bank-transactions/"))
        .Needs("BANKS", "frmBankTx").WithName("PostBankTransaction")
        .WithSummary("A bank transaction (modBank.PostBankTx)")
        .WithDescription("TxType DEPOSIT (CashBoxID: cash into the bank, the box must hold it on that day), WITHDRAW (CashBoxID), " +
                         "SETTLEMENT (Mada collections less FeeAmount and FeeVAT), TRANSFER (ToBankID), OTHER_IN / OTHER_OUT " +
                         "(CounterAccount: an active sub-account, not a box, bank, Mada, customers or suppliers; FeeVAT = the input VAT " +
                         "of an OTHER_OUT). TxDate empty = today, never a future day nor a closed period. Number BNK-. The journal is " +
                         "updated after it, as in Access.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        bankTx.MapGet("", async (HttpContext http, TreasuryService service, DateTime? from, DateTime? to, int? bankId) =>
            Results.Ok(await service.BankTxAsync(from ?? MonthStart, to ?? DateTime.Today, bankId ?? 0, ApiResults.English(http))))
        .Needs("BANKS", "frmBankTx").WithName("ListBankTransactions")
        .WithSummary("The bank transactions between two days (default this month), newest first")
        .Produces<List<BankTxRow>>();

        bankTx.MapDelete("/{id:int}", async (int id, HttpContext http, TreasuryService service) =>
            Changed(http, await service.DeleteBankTxAsync(UserInfo.Acting(http.User), id)))
        .Needs("BANKS", "frmBankTx").WithName("DeleteBankTransaction")
        .WithSummary("Deletes a bank transaction not cleared in a reconciliation (modBank.DeleteBankTx)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ------------------------------------------------------------------ accounting
        var journal = secured.MapGroup("/journal-entries").WithTags("Accounting");

        journal.MapGet("", async (HttpContext http, AccountingService service, DateTime? from, DateTime? to, string? sourceType, string? search) =>
            Results.Ok(await service.EntriesAsync(from ?? MonthStart, to ?? DateTime.Today, sourceType, search, ApiResults.English(http))))
        .Needs("JOURNAL", "frmJournal").WithName("ListJournalEntries")
        .WithSummary("The journal entries between two days (default this month), newest first, at most 300")
        .Produces<List<JournalEntryRow>>();

        journal.MapGet("/{id:int}", async (int id, HttpContext http, AccountingService service) =>
            await service.EntryAsync(id, ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("JOURNAL", "frmJournal").WithName("GetJournalEntry")
        .WithSummary("One journal entry with its lines, accounts and cost centres")
        .Produces<JournalEntryView>().ProducesProblem(StatusCodes.Status404NotFound);

        journal.MapPost("/sync", async (HttpContext http, AccountingService service) =>
        {
            var r = await service.SyncAsync(UserInfo.Acting(http.User));
            return Results.Ok(new { r.Added, r.Updated, r.Removed, message = ApiResults.Text(http, r.Message) });
        })
        .Needs("JOURNAL", "frmJournal").WithName("SyncJournal")
        .WithSummary("Rebuilds the journal entries from the operations (modJournal.SyncJournal)")
        .WithDescription("Reads the journal source views of the server (docs/71); message = the operations not balanced or changed in a closed period.");

        var accounting = secured.MapGroup("/accounting").WithTags("Accounting");

        accounting.MapGet("/accounts", async (HttpContext http, AccountingService service) =>
            Results.Ok(await service.AccountsAsync(ApiResults.English(http))))
        .Needs("JOURNAL", "frmLedger").WithName("ListAccounts")
        .WithSummary("The active accounts of the tree in its order")
        .Produces<List<AccountChoice>>();

        accounting.MapGet("/ledger/{accountCode:int}", async (int accountCode, HttpContext http, AccountingService service, DateTime? from, DateTime? to) =>
        {
            var (f, t) = (from ?? YearStart, to ?? DateTime.Today);
            string p = AccountingService.PeriodProblem(f, t);
            if (p.Length > 0)
            {
                return ApiResults.Refused(http, p);
            }
            return await service.LedgerAsync(accountCode, f, t, ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http);
        })
        .Needs("JOURNAL", "frmLedger").WithName("AccountStatement")
        .WithSummary("The statement of an account (AccountStatementQuery): the opening balance, each journal line with the balance after it, and the four figures")
        .WithDescription("A main account takes all its sub-accounts. The period defaults to the start of this year until today.")
        .Produces<LedgerView>().ProducesProblem(StatusCodes.Status404NotFound).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        accounting.MapGet("/trial-balance", async (HttpContext http, AccountingService service, DateTime? from, DateTime? to) =>
        {
            var (f, t) = (from ?? YearStart, to ?? DateTime.Today);
            return await Period(http, f, t, async () => await service.TrialBalanceAsync(f, t, ApiResults.English(http)));
        })
        .Needs("JOURNAL", "frmLedger").WithName("TrialBalance")
        .WithSummary("The trial balance (TrialBalanceQuery): each account with its opening balance, debit and credit of the period, closing balance")
        .Produces<List<TrialBalanceRow>>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        accounting.MapGet("/income-statement", async (HttpContext http, AccountingService service, DateTime? from, DateTime? to) =>
        {
            var (f, t) = (from ?? YearStart, to ?? DateTime.Today);
            return await Period(http, f, t, async () => Rows(await service.IncomeStatementAsync(f, t, ApiResults.English(http))));
        })
        .Needs("REPORTS_PROFIT", "frmFinancials").WithName("IncomeStatement")
        .WithSummary("The income statement (IncomeStatementQuery) with its comparison period (modQueryParams.SetComparePeriod)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        accounting.MapGet("/balance-sheet", async (HttpContext http, AccountingService service, DateTime? from, DateTime? to) =>
        {
            var (f, t) = (from ?? YearStart, to ?? DateTime.Today);
            return await Period(http, f, t, async () => Rows(await service.BalanceSheetAsync(f, t, ApiResults.English(http))));
        })
        .Needs("REPORTS_PROFIT", "frmFinancials").WithName("BalanceSheet")
        .WithSummary("The balance sheet at the end of the period (BalanceSheetQuery), the profit not closed yet in equity, with its comparison")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ------------------------------------------------------------------ dashboard
        secured.MapGet("/dashboard", async (HttpContext http, DashboardService service) =>
            Results.Ok(await service.FiguresAsync(UserInfo.EmployeeId(http.User), DateTime.Today)))
        .WithTags("Dashboard").WithName("Dashboard")
        .WithSummary("The figures of the dashboard (DashboardQuery): today and month sales, month net profit, low stock, debts, dues, stock value, expenses, cash and banks")
        .WithDescription("Every user sees the dashboard (frmMain); without DASHBOARD_FINANCIAL the amounts are null and only the " +
                         "counts are given, as in Access.")
        .Produces<DashboardFigures>();

        secured.MapGet("/dashboard/office-sync", async (DashboardService service) => Results.Ok(await service.OfficeSyncAsync(DateTime.Now)))
        .Needs("OFFICE_SYNC", "frmOfficeSync").WithTags("Dashboard").WithName("OfficeSyncStatus")
        .WithSummary("The office tasks of Access (modOfficeSync): the computer, the period, the last run and its problem, the pending e-invoices, stale or not")
        .WithDescription("Access is kept for the office (docs/84): one Access computer updates the journal and sends the e-invoices of the " +
                         "documents made on the web, the Windows POS and the rep app. NeedsAttention = they stopped, failed, or nobody runs them.")
        .Produces<OfficeSyncStatus>();
    }

    /// <summary>A statement as JSON (the tuples of its figures as objects).</summary>
    private static object Rows(StatementView v) => new
    {
        v.Kind, v.From, v.To, v.CompareFrom, v.CompareTo, v.Rows,
        Figures = v.Figures.Select(f => new { f.Caption, f.Amount }).ToList(),
    };
}
