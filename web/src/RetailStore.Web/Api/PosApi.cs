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
    }
}
