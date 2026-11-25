using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>The work of the sales rep (docs/77), for the rep app (plan 4D): always the rep linked to the user
/// (SalesReps.EmployeeID), never another rep's data.</summary>
public static class RepApi
{
    public static void Map(RouteGroupBuilder rep)
    {
        rep.MapGet("/profile", async (HttpContext http, RepService reps) =>
            Results.Ok(await reps.ProfileAsync(UserInfo.EmployeeId(http.User), ApiResults.English(http))))
        .WithName("RepProfile").WithSummary("The rep of the user: his van warehouse, his custody box, his limits")
        .WithDescription("The order warehouse is the warehouse of the user (modWarehouses.CurrentWarehouseID). " +
                         "CustomersBalance = the positive balances of his customers (modRepRoutes.RepCustomersBalance).")
        .Produces<RepProfile>();

        rep.MapGet("/sync", async (HttpContext http, RepService reps, int? orderDays) =>
        {
            var snapshot = await reps.SnapshotAsync(UserInfo.EmployeeId(http.User), ApiResults.English(http),
                                                    Math.Clamp(orderDays ?? 14, 1, 90));
            return snapshot is null ? ApiResults.Refused(http, RepService.NotARep) : Results.Ok(snapshot);
        })
        .Needs("SALES_ORDERS").WithName("RepSync")
        .WithSummary("Everything the rep app keeps on the phone to work offline (docs/78), in one read")
        .WithDescription("The store name, the profile, the customers (CustomerPriceGroups: the price group of each one), the " +
                         "products with the price without a customer and the van stock, PriceGroups = the OrderDefaultPrice " +
                         "of each group where it differs, the visits of today (with REP_ROUTES) and the orders of the last " +
                         "orderDays days (default 14). The phone reads it again after sending what it recorded offline.")
        .Produces<RepSnapshot>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapGet("/customers", async (HttpContext http, RepService reps) =>
            Results.Ok(await reps.CustomersAsync(RepFilter.RepOf(http), ApiResults.English(http))))
        .Needs("REP_ROUTES").WithName("RepCustomers")
        .WithSummary("The active customers of the rep: assigned to him, or on his route (with the weekdays, Sunday = 1)")
        .Produces<List<RepCustomer>>();

        rep.MapGet("/products", async (HttpContext http, RepService reps, int? customerId) =>
        {
            var profile = await reps.ProfileAsync(UserInfo.EmployeeId(http.User), false);
            return Results.Ok(await reps.ProductsAsync(RepFilter.RepOf(http), customerId ?? 0, profile?.VanWarehouseID,
                                                       ApiResults.English(http)));
        })
        .Needs("SALES_ORDERS").WithName("RepProducts")
        .WithSummary("The active products with the price of an order line and the stock of the rep's van")
        .WithDescription("Price = modSalesOrders.OrderDefaultPrice for the customer (customerId): his price list, else the " +
                         "rep's price (RepPrices) or the selling price less the discount of the list. In the smallest unit.")
        .Produces<List<RepProduct>>();

        rep.MapGet("/van-stock", async (HttpContext http, RepService reps) =>
        {
            var profile = await reps.ProfileAsync(UserInfo.EmployeeId(http.User), false);
            if (profile?.VanWarehouseID is not { } van)
            {
                return ApiResults.Refused(http, "لا مخزن سيارة للمندوب. أضفه في شاشة المخازن (النوع سيارة).");
            }
            var stock = await reps.WarehouseStockAsync(van);
            return Results.Ok(stock.Where(s => s.Value != 0).Select(s => new { ProductID = s.Key, Quantity = s.Value }));
        })
        .Needs("REP_ROUTES").WithName("RepVanStock")
        .WithSummary("The stock of the rep's van warehouse (modWarehouses.AvailableQty), products with a quantity")
        .ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapGet("/visits", async (HttpContext http, RepService reps, DateTime? date) =>
            Results.Ok(await reps.VisitsAsync(RepFilter.RepOf(http), date ?? DateTime.Today, ApiResults.English(http))))
        .Needs("REP_ROUTES").WithName("RepVisits").WithSummary("The visits of a day (default today)")
        .Produces<List<RepVisit>>();

        rep.MapPost("/visits/plan", async (PlanRequest body, HttpContext http, RepService reps) =>
        {
            var (problem, added) = await reps.CreateVisitPlanAsync(UserInfo.Acting(http.User), RepFilter.RepOf(http),
                                                                    body.Date ?? DateTime.Today);
            return ApiResults.OfMessage(http, problem, () => Results.Ok(new CreatedId(added, "")));
        })
        .Needs("REP_ROUTES").WithName("PlanRepVisits")
        .WithSummary("The plan of a day from the route of its weekday (modRepRoutes.CreateVisitPlan); Id = visits added")
        .Produces<CreatedId>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapPost("/visits", async (VisitRequest body, HttpContext http, RepService reps) =>
        {
            var (problem, id) = await reps.AddVisitAsync(UserInfo.Acting(http.User), RepFilter.RepOf(http), body.CustomerID,
                                                         body.Date ?? DateTime.Today);
            return ApiResults.OfMessage(http, problem, () => Results.Created(ApiEndpoints.Prefix + "/rep/visits/" + id, new CreatedId(id, "")));
        })
        .Needs("REP_ROUTES").WithName("AddRepVisit").WithSummary("A visit outside the route (modRepRoutes.AddRepVisit)")
        .Produces<CreatedId>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapPut("/visits/{id:int}", async (int id, VisitResultRequest body, HttpContext http, RepService reps) =>
            ApiResults.OfMessage(http, await reps.RecordVisitAsync(UserInfo.Acting(http.User), id, body.Result, body.Notes,
                                                                   RepFilter.RepOf(http)),
                                 () => Results.Ok(new CreatedId(id, ""))))
        .Needs("REP_ROUTES").WithName("RecordRepVisit")
        .WithSummary("The result of a visit (modRepRoutes.RecordRepVisit), audited")
        .Produces<CreatedId>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapDelete("/visits/{id:int}", async (int id, HttpContext http, RepService reps) =>
            ApiResults.OfMessage(http, await reps.DeleteVisitAsync(UserInfo.Acting(http.User), id, RepFilter.RepOf(http)),
                                 () => Results.Ok(new CreatedId(id, ""))))
        .Needs("REP_ROUTES").WithName("DeleteRepVisit")
        .WithSummary("Delete a visit not recorded yet (modRepRoutes.DeleteRepVisit), audited")
        .Produces<CreatedId>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapGet("/orders", async (HttpContext http, RepService reps, DateTime? from, DateTime? to) =>
            Results.Ok(await reps.OrdersAsync(RepFilter.RepOf(http), UserInfo.EmployeeId(http.User), from ?? DateTime.Today.AddDays(-30), to ?? DateTime.Today,
                                              ApiResults.English(http))))
        .Needs("SALES_ORDERS").WithName("RepOrders")
        .WithSummary("The customer orders of the rep, or entered by the user, between two days (default the last 30), with their lines and status")
        .Produces<List<OrderView>>();

        rep.MapGet("/orders/{id:int}", async (int id, HttpContext http, RepService reps) =>
        {
            var list = await reps.OrdersAsync(RepFilter.RepOf(http), UserInfo.EmployeeId(http.User), DateTime.MinValue, DateTime.MaxValue.AddDays(-2),
                                              ApiResults.English(http), id);
            return list.Count == 0 ? ApiResults.NotFound(http) : Results.Ok(list[0]);
        })
        .Needs("SALES_ORDERS").WithName("RepOrder").WithSummary("One order of the rep")
        .Produces<OrderView>().ProducesProblem(StatusCodes.Status404NotFound);

        rep.MapPost("/orders", async (OrderRequest body, HttpContext http, RepService reps) =>
        {
            var (problem, id, existing) = await reps.CreateOrderAsync(UserInfo.Acting(http.User), body.CustomerID, body.WarehouseID,
                                                                      body.RequiredDate, body.Notes, body.Lines ?? [],
                                                                      body.ClientRef, body.OrderDate);
            return ApiResults.OfMessage(http, problem, () => existing
                ? Results.Ok(new CreatedId(id, ""))
                : Results.Created(ApiEndpoints.Prefix + "/rep/orders/" + id, new CreatedId(id, "")));
        })
        .Needs("SALES_ORDERS").WithName("AddRepOrder")
        .WithSummary("A customer order as a draft, with its lines (modSalesOrders.CreateSalesOrder)")
        .WithDescription("Number SO-, the warehouse of the user unless given, the rep of the customer; each line checked " +
                         "(SalesOrderLineProblem), a line without a price at OrderDefaultPrice. The office approves it in " +
                         "Access, which reserves the quantities. ClientRef (the number of the order on the phone, docs/78): " +
                         "sent again by the same user, the order is not added twice (200 with its id); OrderDate = when it " +
                         "was taken on the phone (never later than now).")
        .Produces<CreatedId>(StatusCodes.Status201Created).Produces<CreatedId>().ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        // ---- the invoice and the receipt of the rep (docs/87): made on the program at once, in his open rep day, each sent once
        rep.MapPost("/sales", async (RepSaleRequest body, HttpContext http, RepDocService docs) =>
            RepDocument(http, await docs.PostSaleAsync(UserInfo.Acting(http.User), RepFilter.RepOf(http), body with { Lines = body.Lines ?? [] },
                                                       ApiResults.English(http))))
        .Needs("SALES_POS", "frmPOS").WithName("PostRepSale")
        .WithSummary("A sales invoice of the rep from his van, in his open rep day (modSales.PostSaleFromCart acting as the rep)")
        .WithDescription("The rep day open for the rep in Access (frmRepDay) gives the warehouse (the van), the box of the cash (the " +
                         "custody box) and the rep; without one the invoice is refused. The checks of POST /sales in their order; " +
                         "ConfirmStock = the answer to the question on the stock (409 confirm). ClientRef = the number of the " +
                         "invoice on the phone (at most 40 characters): sent again by the same user it answers 200 with the " +
                         "document and adds nothing. The answer is the document to print (number INV-, totals, QR).")
        .Produces<PosDocument>(StatusCodes.Status201Created).Produces<PosDocument>()
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);

        rep.MapPost("/payments", async (DevicePaymentRequest body, HttpContext http, RepDocService docs) =>
            RepDocument(http, await docs.PostPaymentAsync(UserInfo.Acting(http.User), RepFilter.RepOf(http), body, ApiResults.English(http))))
        .Needs("CUSTOMER_PAYMENTS", "frmCustomerPayment").WithName("PostRepPayment")
        .WithSummary("A receipt from a customer by the rep, into his custody box, in his open rep day (modSales.PostCustomerPayment)")
        .WithDescription("An amount over the balance answers 409 confirm; send it again with ConfirmOverBalance = true. ClientRef as " +
                         "for the invoice.")
        .Produces<PosDocument>(StatusCodes.Status201Created).Produces<PosDocument>()
        .ProducesProblem(StatusCodes.Status409Conflict).ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }

    private static IResult RepDocument(HttpContext http, DeviceDocResult r) => r.Status switch
    {
        SaveStatus.Saved when r.Duplicate => Results.Ok(r.Document),
        SaveStatus.Saved => Results.Created(ApiEndpoints.Prefix + "/rep/documents/" + r.Document?.Number, r.Document),
        SaveStatus.NeedsConfirmation => ApiResults.Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message),
        _ => ApiResults.Refused(http, r.Message),
    };
}
