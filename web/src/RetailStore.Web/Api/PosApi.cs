using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>The point of sale on Windows (plan 6G, docs/83) when its database is remote: the snapshot it keeps to sell
/// without the internet, and the sales it made, each sent once (its number on the device).</summary>
public static class PosApi
{
    public static void Map(RouteGroupBuilder secured)
    {
        var pos = secured.MapGroup("/pos").WithTags("Point of sale");

        pos.MapGet("/sync", async (HttpContext http, PosService service) =>
            await service.SnapshotAsync(UserInfo.EmployeeId(http.User), ApiResults.English(http)) is { } s
                ? Results.Ok(s) : ApiResults.NotFound(http))
        .Needs("SALES_POS", "frmPOS").WithName("PosSync")
        .WithSummary("Everything the POS keeps on the PC to sell without the internet (docs/83), in one read")
        .WithDescription("The settings of a sale, the user (his permissions as the sale checks them, his discount limit, " +
                         "warehouse and box), the active products with the price of a new line without a customer, their units " +
                         "and the stock of his warehouse less the reservations, PriceGroups = the prices of each price list " +
                         "where they differ, the active customers (balance, credit, days late), the reps with a credit limit " +
                         "and the payment methods.")
        .Produces<PosSnapshot>();

        pos.MapPost("/sales", async (DeviceSaleRequest body, HttpContext http, PosService service) =>
        {
            var r = await service.PostDeviceSaleAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] });
            return r.Status switch
            {
                SaveStatus.Saved when r.Duplicate => Results.Ok(new PostedDocument(r.InvoiceID, r.InvoiceNumber, 0, 0)),
                SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + "/sales/" + r.InvoiceID,
                                                    new PostedDocument(r.InvoiceID, r.InvoiceNumber, 0, 0)),
                SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
                _ => ApiResults.Refused(http, r.Message),
            };
        })
        .Needs("SALES_POS", "frmPOS").WithName("PostPosSale")
        .WithSummary("A sale made on the POS, with its number on the device (modSales.PostSaleFromCart)")
        .WithDescription("ClientRef = the number of the sale on the device (at most 40 characters): a sale sent again by the " +
                         "same user answers 200 with its invoice and adds nothing. InvoiceNumber = the number the device " +
                         "printed (its code, a dash and digits); TakenAt = when it was made (not after now). The checks of " +
                         "POST /sales; ConfirmStock = the answer to the question on the stock.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).Produces<PostedDocument>()
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ---- returns, receipts and cashier closings (docs/86): made on the program at once, each sent once (its ClientRef)
        pos.MapGet("/returns/{number}", async (string number, HttpContext http, SalesService service) =>
            await service.ReturnableAsync(number, ApiResults.English(http)) is { } inv
                ? Results.Ok(inv) : ApiResults.Problem(http, StatusCodes.Status404NotFound, "not_found", "لا توجد فاتورة بهذا الرقم."))
        .Needs("SALES_RETURN", "frmSalesReturn").WithName("PosReturnSource")
        .WithSummary("The invoice of a number with its lines to return from (modPOS.LoadInvoiceForReturn)")
        .Produces<ReturnableInvoice>().ProducesProblem(StatusCodes.Status404NotFound);

        pos.MapPost("/returns", async (DeviceReturnRequest body, HttpContext http, PosService service) =>
            Document(http, await service.PostDeviceReturnAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] },
                                                               ApiResults.English(http))))
        .Needs("SALES_RETURN", "frmSalesReturn", ScreenAction.Add).WithName("PostPosReturn")
        .WithSummary("A sales return made on the POS, with its number on the device (modSales.PostSalesReturn)")
        .WithDescription("ClientRef = the number of the return on the device (at most 40 characters): a return sent again by the " +
                         "same user answers 200 with its document and adds nothing. Lines = the invoice lines and the quantities " +
                         "returned in their unit; RefundType CASH or CREDIT (the cash customer is always paid back).")
        .Produces<PosDocument>(StatusCodes.Status201Created).Produces<PosDocument>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        pos.MapPost("/payments", async (DevicePaymentRequest body, HttpContext http, PosService service) =>
            Document(http, await service.PostDevicePaymentAsync(UserInfo.Acting(http.User), body, ApiResults.English(http))))
        .Needs("CUSTOMER_PAYMENTS", "frmCustomerPayment").WithName("PostPosPayment")
        .WithSummary("A customer receipt made on the POS, with its number on the device (modSales.PostCustomerPayment)")
        .WithDescription("An amount over the balance answers 409 confirm; send it again with ConfirmOverBalance = true.")
        .Produces<PosDocument>(StatusCodes.Status201Created).Produces<PosDocument>()
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        pos.MapGet("/closing", async (HttpContext http, PosService service) =>
            Results.Ok(await service.ClosingBoxesAsync(UserInfo.EmployeeId(http.User), ApiResults.English(http))))
        .Needs("CASH_CLOSING", "frmCashClosing").WithName("PosClosingBoxes")
        .WithSummary("The boxes the user may close, the main safes and his own box (modCash.ClosingLoad)")
        .Produces<PosClosingBoxes>();

        pos.MapGet("/closing/{box:int}", async (int box, HttpContext http, TreasuryService service) =>
            await service.ClosingPreviewAsync(box, ApiResults.English(http)) is { } p ? Results.Ok(p) : ApiResults.NotFound(http))
        .Needs("CASH_CLOSING", "frmCashClosing").WithName("PosClosingPreview")
        .WithSummary("What the closing of a box finds now: the last closing, the cash in and out since, the book balance (ClosingBoxChanged)")
        .Produces<ClosingPreview>().ProducesProblem(StatusCodes.Status404NotFound);

        pos.MapPost("/closing", async (DeviceClosingRequest body, HttpContext http, PosService service) =>
            Document(http, await service.PostDeviceClosingAsync(UserInfo.Acting(http.User), body, ApiResults.English(http))))
        .Needs("CASH_CLOSING", "frmCashClosing", ScreenAction.Add).WithName("PostPosClosing")
        .WithSummary("A cashier closing made on the POS, with its number on the device (modCash.PostCashClosing)")
        .WithDescription("Destination MAIN (TransferAmount to the safe ToCashBoxID), OWNER or KEEP; the shortage or overage is booked " +
                         "with a voucher, a shortage of a rep's custody box is an advance on him.")
        .Produces<PosDocument>(StatusCodes.Status201Created).Produces<PosDocument>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }

    private static IResult Document(HttpContext http, DeviceDocResult r) => r.Status switch
    {
        SaveStatus.Saved when r.Duplicate => Results.Ok(r.Document),
        SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + "/pos/documents/" + r.Document?.Number, r.Document),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };
}
