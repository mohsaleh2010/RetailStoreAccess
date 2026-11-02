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
            var (problem, id) = await reps.CreateOrderAsync(UserInfo.Acting(http.User), body.CustomerID, body.WarehouseID,
                                                            body.RequiredDate, body.Notes, body.Lines ?? []);
            return ApiResults.OfMessage(http, problem, () => Results.Created(ApiEndpoints.Prefix + "/rep/orders/" + id, new CreatedId(id, "")));
        })
        .Needs("SALES_ORDERS").WithName("AddRepOrder")
        .WithSummary("A customer order as a draft, with its lines (modSalesOrders.CreateSalesOrder)")
        .WithDescription("Number SO-, the warehouse of the user unless given, the rep of the customer; each line checked " +
                         "(SalesOrderLineProblem), a line without a price at OrderDefaultPrice. The office approves it in " +
                         "Access, which reserves the quantities.")
        .Produces<CreatedId>(StatusCodes.Status201Created).ProducesProblem(StatusCodes.Status422UnprocessableEntity);
    }
}
