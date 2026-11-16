using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>Stock and transfers between warehouses (plan 6D, docs/80): the same service as the stock screens of the web
/// (modWarehouses in Access).</summary>
public static class StockApi
{
    private static IResult Done(HttpContext http, StockResult r, Func<IResult> ok) => r.Status switch
    {
        SaveStatus.Saved => ok(),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };

    public static void Map(RouteGroupBuilder secured)
    {
        var stock = secured.MapGroup("/stock").WithTags("Stock");

        stock.MapGet("", async (HttpContext http, StockService service, int? warehouseId, string? search, bool? lowOnly) =>
            Results.Ok(await service.StockAsync(warehouseId ?? 0, search, lowOnly == true, ApiResults.English(http))))
        .Needs("PRODUCTS", "frmInventory").WithName("ListStock")
        .WithSummary("The stock of the active stock products in a warehouse (0 = every warehouse), with the reserved quantity and the value")
        .WithDescription("modWarehouses.AvailableQty per product; Reserved = what open customer orders still need there " +
                         "(modSalesOrders.ReservedQty); Value = quantity x average cost; lowOnly = at or under the minimum. At most 500 rows.")
        .Produces<List<StockRow>>();

        stock.MapGet("/{productId:int}/card", async (int productId, HttpContext http, StockService service, int? warehouseId, DateTime? from,
                                                     DateTime? to) =>
            Results.Ok(await service.ItemCardAsync(productId, warehouseId ?? 0, from ?? DateTime.Today.AddDays(1 - DateTime.Today.Day),
                                                   to ?? DateTime.Today, ApiResults.English(http))))
        .Needs("PRODUCTS", "frmInventory").WithName("ItemCard")
        .WithSummary("The item card of a product (ItemCardQuery): the opening balance, then each move with its other party and the running balance")
        .WithDescription("warehouseId 0 = every warehouse; the period defaults to the start of this month until today.")
        .Produces<List<ItemCardLine>>();

        var transfers = secured.MapGroup("/stock-transfers").WithTags("Stock");

        transfers.MapGet("", async (HttpContext http, StockService service, string? status, int? warehouseId) =>
            Results.Ok(await service.TransfersAsync(status, warehouseId ?? 0, ApiResults.English(http))))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("ListStockTransfers")
        .WithSummary("The transfers between warehouses, newest first: a status (DRAFT, SENT, RECEIVED) and a warehouse on either side")
        .Produces<List<TransferRow>>();

        transfers.MapGet("/{id:int}", async (int id, HttpContext http, StockService service) =>
            await service.TransferAsync(id, UserInfo.EmployeeId(http.User), ApiResults.English(http)) is { } v ? Results.Ok(v) : ApiResults.NotFound(http))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("GetStockTransfer")
        .WithSummary("One transfer with its lines and the free stock of the sending warehouse")
        .Produces<TransferView>().ProducesProblem(StatusCodes.Status404NotFound);

        transfers.MapPost("", async (TransferRequest body, HttpContext http, StockService service) =>
        {
            var r = await service.CreateTransferAsync(UserInfo.Acting(http.User), body with { Lines = body.Lines ?? [] });
            return Done(http, r, () => Results.Created(ApiEndpoints.Prefix + "/stock-transfers/" + r.Id, new PostedDocument(r.Id, r.Number, 0, 0)));
        })
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("CreateStockTransfer")
        .WithSummary("A draft transfer with its lines (modWarehouses.CreateStockTransfer)")
        .WithDescription("FromWarehouseID and ToWarehouseID active and different; each line a stock product with a quantity over zero " +
                         "(in the smallest unit). Number STT-. Send it with POST /stock-transfers/{id}/send.")
        .Produces<PostedDocument>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        transfers.MapPut("/{id:int}/lines", async (int id, List<TransferLineInput> body, HttpContext http, StockService service) =>
            Done(http, await service.SaveTransferLinesAsync(UserInfo.Acting(http.User), id, body ?? []), Results.NoContent))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("SaveStockTransferLines")
        .WithSummary("Replaces the lines of a draft transfer")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        transfers.MapPost("/{id:int}/send", async (int id, bool? confirm, HttpContext http, StockService service) =>
            Done(http, await service.SendTransferAsync(UserInfo.Acting(http.User), id, confirm == true), Results.NoContent))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("SendStockTransfer")
        .WithSummary("Sends a draft (modWarehouses.SendStockTransfer)")
        .WithDescription("Out of the sending warehouse at the average cost; into the receiving one at once, or later when that " +
                         "warehouse confirms receipt (ConfirmReceipt: SENT, in transit on account 1410). A closed period is refused. " +
                         "More than the free stock with ALLOW_NEGATIVE_STOCK answers 409 (code confirm): send it again with ?confirm=true.")
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        transfers.MapPost("/{id:int}/receive", async (int id, HttpContext http, StockService service) =>
            Done(http, await service.ReceiveTransferAsync(UserInfo.Acting(http.User), id), Results.NoContent))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("ReceiveStockTransfer")
        .WithSummary("Confirms that a transfer in transit arrived (modWarehouses.ReceiveStockTransfer)")
        .WithDescription("Only the keeper of the receiving warehouse, a user whose warehouse it is, or the administrator (anyone with " +
                         "the permission when it has no keeper). The goods enter at the cost they left with.")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        transfers.MapPost("/{id:int}/unpost", async (int id, HttpContext http, StockService service) =>
            Done(http, await service.UnpostTransferAsync(UserInfo.Acting(http.User), id), Results.NoContent))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("UnpostStockTransfer")
        .WithSummary("A transfer still in transit goes back to draft (modWarehouses.UnpostStockTransfer)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        transfers.MapDelete("/{id:int}", async (int id, HttpContext http, StockService service) =>
            Done(http, await service.DeleteTransferAsync(UserInfo.Acting(http.User), id), Results.NoContent))
        .Needs("STOCK_TRANSFER", "frmStockTransfer").WithName("DeleteStockTransfer")
        .WithSummary("Deletes a draft transfer and its lines (modWarehouses.DeleteStockTransfer)")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }
}
