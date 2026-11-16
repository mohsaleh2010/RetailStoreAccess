using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>Purchases and supplier payments (plan 6D, docs/80): the same service as the purchase screens of the web, so the
/// same checks, messages and rows as modPurchases.PostPurchaseFromCart / PostSupplierPayment in Access.</summary>
public static class PurchaseApi
{
    public static void Map(RouteGroupBuilder secured)
    {
        var purchases = secured.MapGroup("/purchases").WithTags("Purchases");

        purchases.MapPost("", async (PurchaseRequest body, HttpContext http, PurchaseService service) =>
            SalesApi.Posted(http, await service.PostPurchaseAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] }),
                            "/purchases/"))
        .Needs("PURCHASES", "frmPurchaseInvoice").WithName("PostPurchase")
        .WithSummary("A purchase invoice (modPurchases.PostPurchaseFromCart)")
        .WithDescription("PaymentType CASH or CREDIT; Paid empty = the total of a cash invoice, nothing for a credit one; InvoiceDate " +
                         "empty = now (an earlier day for an invoice entered late, never a future day); ChargeVat = the supplier " +
                         "charged VAT (standard-rated products); WarehouseID 0 = the user's warehouse; the amounts in CurrencyCode " +
                         "(empty = the program currency) at ExchangeRate. Each line: ProductID, Quantity, UnitCost WITHOUT VAT, " +
                         "LineDiscount, UnitID (empty = the smallest unit), NewSellingPrice (a new shelf price, needs PRODUCTS). " +
                         "The goods enter the warehouse at the cost of one piece and update the average cost; a supplier invoice " +
                         "number is refused twice for the same supplier. Number PUR-.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        purchases.MapGet("", async (HttpContext http, PurchaseService service, DateTime? from, DateTime? to, string? search) =>
            Results.Ok(await service.PurchasesAsync(from ?? DateTime.Today, to ?? DateTime.Today, search, ApiResults.English(http))))
        .Needs("PURCHASES", "frmPurchaseView").WithName("ListPurchases")
        .WithSummary("The purchase invoices between two days (default today), newest first, at most 200")
        .Produces<List<PurchaseRow>>();

        purchases.MapGet("/{id:int}", async (int id, HttpContext http, PurchaseService service) =>
            await service.PurchaseAsync(id, ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("PURCHASES", "frmPurchaseView").WithName("GetPurchase")
        .WithSummary("One purchase invoice with its lines in their unit")
        .Produces<PurchaseView>().ProducesProblem(StatusCodes.Status404NotFound);

        purchases.MapGet("/products", async (HttpContext http, PurchaseService service, string? code, int? warehouseId) =>
            Results.Ok(await service.FindProductsAsync(UserInfo.EmployeeId(http.User), code ?? "", ApiResults.English(http), warehouseId ?? 0)))
        .Needs("PURCHASES", "frmPurchaseInvoice").WithName("FindPurchaseProducts")
        .WithSummary("Products by a barcode or code (a unit barcode gives its unit), else by a part of the name, with the last purchase price")
        .WithDescription("modProductUnits.FindProductCode and UnitPurchasePrice: the unit's own purchase price, else the piece price x " +
                         "its factor (without VAT). Available = the stock of the warehouse (0 = the user's).")
        .Produces<List<PurchaseProduct>>();

        var payments = secured.MapGroup("/supplier-payments").WithTags("Purchases");

        payments.MapPost("", async (SupplierPaymentRequest body, bool? confirm, HttpContext http, PurchaseService service) =>
            SalesApi.Posted(http, await service.PostSupplierPaymentAsync(UserInfo.Acting(http.User), body, confirm == true), "/supplier-payments/"))
        .Needs("SUPPLIER_PAYMENTS", "frmSupplierPayment").WithName("PostSupplierPayment")
        .WithSummary("A payment voucher to a supplier (modPurchases.PostSupplierPayment)")
        .WithDescription("Amount in CurrencyCode (empty = the program currency) at ExchangeRate, kept in the program currency. An " +
                         "amount over what is owed answers 409 (code confirm): send the same request with ?confirm=true. Number " +
                         "PAY-; a cash payment leaves the box of the user, a bank transfer the default bank.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status409Conflict)
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        payments.MapGet("", async (HttpContext http, PurchaseService service, DateTime? from, DateTime? to, string? search) =>
            Results.Ok(await service.SupplierPaymentsAsync(from ?? DateTime.Today, to ?? DateTime.Today, search, ApiResults.English(http))))
        .Needs("SUPPLIER_PAYMENTS", "frmSupplierPayment").WithName("ListSupplierPayments")
        .WithSummary("The payment vouchers between two days (default today), newest first, at most 200")
        .Produces<List<SupplierPaymentRow>>();
    }
}
