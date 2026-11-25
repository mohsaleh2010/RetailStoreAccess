using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>A reconciliation started or changed (modBank.OpenReconciliation).</summary>
public sealed record ReconStartRequest(int BankID, DateTime? StatementDate, decimal StatementBalance);

/// <summary>An operation of the bank ticked in, or taken out of, a statement (SourceType + SourceID of qryBankItems).</summary>
public sealed record ReconItemRequest(string SourceType, int SourceID);

/// <summary>The rest of the web (plan 7C, docs/87): sales and purchase returns, the cashier closing, the bank reconciliation and the
/// manual journal entries, with the same services as the screens of the web (modSales, modPurchases, modCash, modBank and
/// modManualEntry in Access).</summary>
public static class OfficeApi
{
    private static IResult Done(HttpContext http, TreasuryResult r, string location) => r.Status switch
    {
        SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + location + r.Id, new PostedDocument(r.Id, r.Number, 0, 0)),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };

    private static IResult Empty(HttpContext http, string problem) =>
        problem.Length == 0 ? Results.NoContent() : ApiResults.Refused(http, problem);

    private static DateTime MonthStart => DateTime.Today.AddDays(1 - DateTime.Today.Day);

    public static void Map(RouteGroupBuilder secured)
    {
        // ------------------------------------------------------------------ returns
        var salesReturns = secured.MapGroup("/sales-returns").WithTags("Sales");

        salesReturns.MapGet("/source/{number}", async (string number, HttpContext http, SalesService service) =>
            await service.ReturnableAsync(number, ApiResults.English(http)) is { } inv
                ? Results.Ok(inv) : ApiResults.Problem(http, StatusCodes.Status404NotFound, "not_found", "لا توجد فاتورة بهذا الرقم."))
        .Needs("SALES_RETURN", "frmSalesReturn").WithName("SalesReturnSource")
        .WithSummary("The sales invoice of a number with its lines to return from (modPOS.LoadInvoiceForReturn)")
        .Produces<ReturnableInvoice>().ProducesProblem(StatusCodes.Status404NotFound);

        salesReturns.MapPost("", async (SalesReturnRequest body, HttpContext http, SalesService service) =>
            SalesApi.Posted(http, await service.PostSalesReturnAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] }),
                            "/sales-returns/"))
        .Needs("SALES_RETURN", "frmSalesReturn", ScreenAction.Add).WithName("PostSalesReturn")
        .WithSummary("A sales return, the credit note of an invoice (modSales.PostSalesReturn)")
        .WithDescription("Lines = the invoice lines with the quantity returned in their unit and ToStock (the goods go back to the " +
                         "warehouse of the invoice at their cost). RefundType CASH (PaymentMethodID; the cash customer is always paid " +
                         "back) or CREDIT (taken off his balance). The reason is required. Number RET-, the ZATCA fields of a credit " +
                         "note (381) and its QR; with e-invoicing on it is PENDING and the office tasks send it.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        salesReturns.MapGet("", async (HttpContext http, SalesService service, DateTime? from, DateTime? to) =>
            Results.Ok(await service.SalesReturnsAsync(from ?? MonthStart, to ?? DateTime.Today, ApiResults.English(http))))
        .Needs("SALES_RETURN", "frmSalesReturn").WithName("ListSalesReturns")
        .WithSummary("The sales returns between two days (default this month), newest first")
        .Produces<List<ReturnRow>>();

        var purchaseReturns = secured.MapGroup("/purchase-returns").WithTags("Purchases");

        purchaseReturns.MapGet("/source/{number}", async (string number, HttpContext http, PurchaseService service) =>
            await service.ReturnableAsync(number, ApiResults.English(http)) is { } inv
                ? Results.Ok(inv) : ApiResults.Problem(http, StatusCodes.Status404NotFound, "not_found", "لا توجد فاتورة شراء بهذا الرقم."))
        .Needs("PURCHASE_RETURN", "frmPurchaseReturn").WithName("PurchaseReturnSource")
        .WithSummary("The purchase invoice of a number (ours or the supplier's) with its lines to return (LoadPurchaseForReturn)")
        .Produces<PurchaseReturnable>().ProducesProblem(StatusCodes.Status404NotFound);

        purchaseReturns.MapPost("", async (PurchaseReturnRequest body, bool? confirm, HttpContext http, PurchaseService service) =>
            SalesApi.Posted(http, await service.PostPurchaseReturnAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] },
                                                                        confirm == true), "/purchase-returns/"))
        .Needs("PURCHASE_RETURN", "frmPurchaseReturn", ScreenAction.Add).WithName("PostPurchaseReturn")
        .WithSummary("A purchase return to the supplier (modPurchases.PostPurchaseReturn)")
        .WithDescription("Lines = the invoice lines with the quantity returned in their unit. The goods leave the warehouse they " +
                         "entered at the cost they came in with (the average cost is recalculated). RefundType CASH (the supplier " +
                         "pays back, PaymentMethodID) or CREDIT (taken off what is owed to him). More than the stock with " +
                         "ALLOW_NEGATIVE_STOCK answers 409 confirm: send it again with ?confirm=true. Number PRT-.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status409Conflict)
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        purchaseReturns.MapGet("", async (HttpContext http, PurchaseService service, DateTime? from, DateTime? to) =>
            Results.Ok(await service.PurchaseReturnsAsync(from ?? MonthStart, to ?? DateTime.Today, ApiResults.English(http))))
        .Needs("PURCHASE_RETURN", "frmPurchaseReturn").WithName("ListPurchaseReturns")
        .WithSummary("The purchase returns between two days (default this month), newest first")
        .Produces<List<ReturnRow>>();

        // ------------------------------------------------------------------ cashier closing
        var closings = secured.MapGroup("/cash-closings").WithTags("Treasury");

        closings.MapGet("/boxes", async (HttpContext http, PosService service) =>
            Results.Ok(await service.ClosingBoxesAsync(UserInfo.EmployeeId(http.User), ApiResults.English(http))))
        .Needs("CASH_CLOSING", "frmCashClosing").WithName("CashClosingBoxes")
        .WithSummary("The boxes the user may close, the main safes and his own box (modCash.ClosingLoad)")
        .Produces<PosClosingBoxes>();

        closings.MapGet("/boxes/{box:int}", async (int box, HttpContext http, TreasuryService service) =>
            await service.ClosingPreviewAsync(box, ApiResults.English(http)) is { } p ? Results.Ok(p) : ApiResults.NotFound(http))
        .Needs("CASH_CLOSING", "frmCashClosing").WithName("CashClosingPreview")
        .WithSummary("What the closing of a box finds now: the last closing, the cash in and out since, the book balance (ClosingBoxChanged)")
        .Produces<ClosingPreview>().ProducesProblem(StatusCodes.Status404NotFound);

        closings.MapPost("", async (CashClosingRequest body, HttpContext http, TreasuryService service) =>
            Done(http, await service.PostCashClosingAsync(UserInfo.Acting(http.User), body), "/cash-closings/"))
        .Needs("CASH_CLOSING", "frmCashClosing", ScreenAction.Add).WithName("PostCashClosing")
        .WithSummary("A cashier closing (modCash.PostCashClosing)")
        .WithDescription("Destination MAIN (TransferAmount to the safe ToCashBoxID, in transit when it confirms receipt), OWNER or " +
                         "KEEP; the shortage or overage is booked with a voucher, a shortage of a rep's custody box is an advance on " +
                         "him. Number CLS-. A closing cannot be changed after it is saved.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ------------------------------------------------------------------ bank reconciliation
        var recons = secured.MapGroup("/bank-reconciliations").WithTags("Banks");

        recons.MapGet("", async (HttpContext http, BankReconService service, int bankId) => Results.Ok(await service.ReconsAsync(bankId)))
        .Needs("BANKS", "frmBankRecon").WithName("ListBankReconciliations")
        .WithSummary("The reconciliations of a bank, the latest statement first")
        .Produces<List<ReconRow>>();

        recons.MapPost("", async (ReconStartRequest body, HttpContext http, BankReconService service) =>
            Done(http, await service.OpenAsync(UserInfo.Acting(http.User), body.BankID, body.StatementDate, body.StatementBalance),
                 "/bank-reconciliations/"))
        .Needs("BANKS", "frmBankRecon").WithName("StartBankReconciliation")
        .WithSummary("Starts the reconciliation of a bank, or changes the date and balance of its open one (modBank.OpenReconciliation)")
        .WithDescription("StatementDate = the day of the bank statement (not after today, not before the last approved one); " +
                         "StatementBalance = the balance the statement shows that day. Number REC-.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        recons.MapGet("/{id:int}", async (int id, HttpContext http, BankReconService service) =>
            await service.ViewAsync(id, ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("BANKS", "frmBankRecon").WithName("BankReconciliation")
        .WithSummary("A reconciliation: the book balance on the statement day, the outstanding operations, the difference, the open and the ticked operations")
        .WithDescription("Read from the journal lines of the bank (qryBankItems): update the journal first (POST /journal-entries/sync).")
        .Produces<ReconView>().ProducesProblem(StatusCodes.Status404NotFound);

        recons.MapPost("/{id:int}/clear", async (int id, ReconItemRequest body, BankReconService service, HttpContext http) =>
            Empty(http, await service.ClearAsync(id, body.SourceType, body.SourceID)))
        .Needs("BANKS", "frmBankRecon").WithName("ClearBankItem")
        .WithSummary("Ticks an operation that appears in the bank statement (modBank.ClearBankItem)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        recons.MapPost("/{id:int}/unclear", async (int id, ReconItemRequest body, BankReconService service, HttpContext http) =>
            Empty(http, await service.UnclearAsync(id, body.SourceType, body.SourceID)))
        .Needs("BANKS", "frmBankRecon").WithName("UnclearBankItem")
        .WithSummary("Takes an operation out of the statement (modBank.UnclearBankItem)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        recons.MapPost("/{id:int}/finish", async (int id, BankReconService service, HttpContext http) =>
            Empty(http, await service.FinishAsync(UserInfo.Acting(http.User), id)))
        .Needs("BANKS", "frmBankRecon", ScreenAction.Edit).WithName("FinishBankReconciliation")
        .WithSummary("Approves the reconciliation when there is no difference (modBank.FinishReconciliation)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        recons.MapPost("/{id:int}/reopen", async (int id, BankReconService service, HttpContext http) =>
            Empty(http, await service.ReopenAsync(UserInfo.Acting(http.User), id)))
        .Needs("BANKS", "frmBankRecon", ScreenAction.Edit).WithName("ReopenBankReconciliation")
        .WithSummary("Opens the last approved reconciliation of the bank again (modBank.ReopenReconciliation)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        recons.MapDelete("/{id:int}", async (int id, BankReconService service, HttpContext http) =>
            Empty(http, await service.DeleteAsync(UserInfo.Acting(http.User), id)))
        .Needs("BANKS", "frmBankRecon").WithName("DeleteBankReconciliation")
        .WithSummary("Deletes the open reconciliation with its ticks, audited (modBank.DeleteReconciliation)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ------------------------------------------------------------------ manual journal entries
        var manual = secured.MapGroup("/manual-entries").WithTags("Accounting");

        manual.MapGet("", async (ManualEntryService service, DateTime? from, DateTime? to) =>
            Results.Ok(await service.EntriesAsync(from ?? MonthStart, to ?? DateTime.Today)))
        .Needs("MANUAL_ENTRY", "frmManualEntry").WithName("ListManualEntries")
        .WithSummary("The manual entries between two days (default this month), newest first")
        .Produces<List<ManualEntryRow>>();

        manual.MapGet("/{id:int}", async (int id, HttpContext http, ManualEntryService service) =>
            await service.EntryAsync(id) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("MANUAL_ENTRY", "frmManualEntry").WithName("ManualEntry")
        .WithSummary("A manual entry with its lines in its currency and its journal entry number")
        .Produces<ManualEntryView>().ProducesProblem(StatusCodes.Status404NotFound);

        manual.MapPost("", async (ManualEntryRequest body, HttpContext http, ManualEntryService service) =>
            Done(http, await service.SaveAsync(UserInfo.Acting(http.User), body with { ManualEntryID = 0, Lines = body.Lines ?? [] }),
                 "/manual-entries/"))
        .Needs("MANUAL_ENTRY", "frmManualEntry", ScreenAction.Add).WithName("AddManualEntry")
        .WithSummary("A manual journal entry (modManualEntry.PostManualEntry), then the journal is updated")
        .WithDescription("Lines: an active sub-account, the amount on one side only (in CurrencyCode at ExchangeRate; empty = the " +
                         "program currency), a text and a cost centre; at least two lines, balanced. Not in a closed period. " +
                         "ReversalOfID = the entry it reverses. Number MJ-; SyncJournal makes its journal entry (source MANUAL).")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        manual.MapPut("/{id:int}", async (int id, ManualEntryRequest body, HttpContext http, ManualEntryService service) =>
        {
            var r = await service.SaveAsync(UserInfo.Acting(http.User), body with { ManualEntryID = id, Lines = body.Lines ?? [] });
            return r.Status == SaveStatus.Saved ? Results.Ok(new PostedDocument(r.Id, r.Number, 0, 0)) : ApiResults.Refused(http, r.Message);
        })
        .Needs("MANUAL_ENTRY", "frmManualEntry", ScreenAction.Edit).WithName("ChangeManualEntry")
        .WithSummary("Changes a manual entry; its journal entry is rebuilt under the same number (audited)")
        .Produces<PostedDocument>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        manual.MapDelete("/{id:int}", async (int id, HttpContext http, ManualEntryService service) =>
        {
            var r = await service.DeleteAsync(UserInfo.Acting(http.User), id);
            return r.Status == SaveStatus.Saved ? Results.NoContent() : ApiResults.Refused(http, r.Message);
        })
        .Needs("MANUAL_ENTRY", "frmManualEntry", ScreenAction.Delete).WithName("DeleteManualEntry")
        .WithSummary("Deletes a manual entry and its journal entry (modManualEntry.RemoveManualEntry), audited")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }
}
