using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>The answer of a posted document: its id, its number, the total and the change.</summary>
public sealed record PostedDocument(int Id, string Number, decimal Total, decimal Change);

/// <summary>Sales and customer receipts (plan 6C, docs/79): the same service as the POS screen of the web, so the same
/// checks, messages and rows as modSales.PostSaleFromCart / PostCustomerPayment in Access.</summary>
public static class SalesApi
{
    private static IResult Posted(HttpContext http, SaleResult r, string route) => r.Status switch
    {
        SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + route + r.InvoiceID, new PostedDocument(r.InvoiceID, r.InvoiceNumber, r.Total, r.Change)),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };

    public static void Map(RouteGroupBuilder secured)
    {
        var sales = secured.MapGroup("/sales").WithTags("Sales");

        sales.MapPost("", async (SaleRequest body, bool? confirm, HttpContext http, SalesService service) =>
            Posted(http, await service.PostSaleAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] }, confirm == true),
                   "/sales/"))
        .Needs("SALES_POS", "frmPOS").WithName("PostSale")
        .WithSummary("A sales invoice from a cart (modSales.PostSaleFromCart)")
        .WithDescription("PaymentType CASH or CREDIT; Tendered empty = the total of a cash sale, nothing for a credit sale. Each line: " +
                         "ProductID, Quantity, UnitPrice (VAT included when the prices include VAT), LineDiscount, UnitID (empty = the " +
                         "smallest unit). The checks of Access in their order: products and units, the invoice engine, a price " +
                         "change (PRICE_OVERRIDE), the discount limit, payment, the customer (credit, limit, late invoices), the " +
                         "limit of the rep, the stock of the user's warehouse less the reservations. More than the stock with " +
                         "ALLOW_NEGATIVE_STOCK answers 409 (code confirm): send the same request with ?confirm=true. The invoice " +
                         "gets its number INV-, the ZATCA fields and QR; with e-invoicing on it is PENDING and Access sends it.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status409Conflict)
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        sales.MapGet("", async (HttpContext http, SalesService service, DateTime? from, DateTime? to, string? search) =>
            Results.Ok(await service.InvoicesAsync(from ?? DateTime.Today, to ?? DateTime.Today, search, ApiResults.English(http))))
        .Needs("SALES_VIEW", "frmSalesInvoice").WithName("ListSales")
        .WithSummary("The sales invoices between two days (default today), newest first, at most 200")
        .Produces<List<InvoiceRow>>();

        sales.MapGet("/{id:int}", async (int id, HttpContext http, SalesService service) =>
            await service.InvoiceAsync(id, ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("SALES_VIEW", "frmSalesInvoice").WithName("GetSale")
        .WithSummary("One sales invoice as it is printed: the customer, the lines in their unit, the ZATCA QR and status")
        .Produces<InvoiceView>().ProducesProblem(StatusCodes.Status404NotFound);

        sales.MapGet("/products", async (HttpContext http, SalesService service, string? code, int? customerId) =>
            Results.Ok(await service.FindProductsAsync(UserInfo.EmployeeId(http.User), code ?? "", customerId ?? 0, ApiResults.English(http))))
        .Needs("SALES_POS", "frmPOS").WithName("FindSaleProducts")
        .WithSummary("Products by a barcode or code (a unit barcode gives its unit), else by a part of the name, with the price of a new line")
        .WithDescription("modProductUnits.FindProductCode and UnitSalePrice for the customer: the rep's price of the user, the unit's " +
                         "own price or the piece price x its factor, then the customer's price list. Available = the stock of the " +
                         "user's warehouse.")
        .Produces<List<SaleProduct>>();

        var payments = secured.MapGroup("/customer-payments").WithTags("Sales");

        payments.MapPost("", async (PaymentRequest body, bool? confirm, HttpContext http, SalesService service) =>
            Posted(http, await service.PostPaymentAsync(UserInfo.Acting(http.User), body, confirm == true), "/customer-payments/"))
        .Needs("CUSTOMER_PAYMENTS", "frmCustomerPayment").WithName("PostCustomerPayment")
        .WithSummary("A receipt voucher from a customer (modSales.PostCustomerPayment)")
        .WithDescription("Amount in CurrencyCode (empty = the program currency) at ExchangeRate, kept in the program currency. An " +
                         "amount over the balance answers 409 (code confirm): send the same request with ?confirm=true. Number " +
                         "RCV-; a cash receipt goes into the box of the user, a bank transfer to the default bank.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status409Conflict)
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        payments.MapGet("", async (HttpContext http, SalesService service, DateTime? from, DateTime? to, string? search) =>
            Results.Ok(await service.PaymentsAsync(from ?? DateTime.Today, to ?? DateTime.Today, search, ApiResults.English(http))))
        .Needs("CUSTOMER_PAYMENTS", "frmCustomerPayment").WithName("ListCustomerPayments")
        .WithSummary("The receipt vouchers between two days (default today), newest first, at most 200")
        .Produces<List<PaymentRow>>();
    }
}
