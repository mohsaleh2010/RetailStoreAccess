using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using RetailStore.Data;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;

namespace RetailStore.Tests;

/// <summary>The web host on a test database (docs/77).</summary>
public sealed class ApiHost(int loginsPerMinute = 1000) : WebApplicationFactory<Program>
{
    public TestDb Db { get; } = new();

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.UseEnvironment("Development");
        builder.UseSetting("Api:JwtKey", "test-key-of-the-api-tokens-0123456789-abcdef");
        builder.UseSetting("Api:LoginsPerMinute", loginsPerMinute.ToString());
        builder.ConfigureServices(s =>
        {
            s.RemoveAll<IDbContextFactory<RetailDbContext>>();
            s.AddSingleton<IDbContextFactory<RetailDbContext>>(Db);
        });
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        Db.Dispose();
    }

    public const string AdminPassword = "Store@2026";

    /// <summary>A client with the token of admin (the password set first).</summary>
    public async Task<HttpClient> AdminAsync()
    {
        await new SecurityService(Db).SetUserPasswordAsync(1, AdminPassword, false, TestDb.Admin);
        return await ClientAsync("admin", AdminPassword);
    }

    public async Task<HttpClient> ClientAsync(string user, string password)
    {
        var c = CreateClient();
        var r = await c.PostAsJsonAsync("/api/v1/auth/token", new { username = user, password, device = "TestPhone" });
        Assert.Equal(HttpStatusCode.OK, r.StatusCode);
        var token = (await r.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("accessToken").GetString();
        c.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        return c;
    }
}

public class ApiTests
{
    private static async Task<JsonElement> Json(HttpResponseMessage r) => await r.Content.ReadFromJsonAsync<JsonElement>();

    private static async Task<string> Code(HttpResponseMessage r) => (await Json(r)).GetProperty("code").GetString()!;

    [Fact]
    public async Task Tokens_follow_the_login_of_Access_and_a_new_password_ends_the_old_tokens()
    {
        using var host = new ApiHost();
        var anonymous = host.CreateClient();
        var none = await anonymous.GetAsync("/api/v1/products");
        Assert.Equal(HttpStatusCode.Unauthorized, none.StatusCode);
        Assert.Equal("application/problem+json", none.Content.Headers.ContentType?.MediaType);   // never the login page
        Assert.Equal("unauthorized", await Code(none));

        var wrong = await anonymous.PostAsJsonAsync("/api/v1/auth/token", new { username = "nobody", password = "x" });
        Assert.Equal(HttpStatusCode.Unauthorized, wrong.StatusCode);
        Assert.Equal("login_failed", await Code(wrong));

        // the first login of admin: no password, then it must be set before anything else
        var first = await host.ClientAsync("admin", "");
        var blocked = await first.GetAsync("/api/v1/products");
        Assert.Equal(HttpStatusCode.Forbidden, blocked.StatusCode);
        Assert.Equal("password_change_required", await Code(blocked));
        Assert.Equal(HttpStatusCode.OK, (await first.GetAsync("/api/v1/me")).StatusCode);
        var weak = await first.PostAsJsonAsync("/api/v1/auth/password", new { newPassword = "1", confirmPassword = "1" });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, weak.StatusCode);
        var changed = await first.PostAsJsonAsync("/api/v1/auth/password",
                                                  new { newPassword = ApiHost.AdminPassword, confirmPassword = ApiHost.AdminPassword });
        Assert.Equal(HttpStatusCode.OK, changed.StatusCode);
        Assert.Equal(HttpStatusCode.Unauthorized, (await first.GetAsync("/api/v1/me")).StatusCode);   // the old token ended

        var next = host.CreateClient();
        next.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", (await Json(changed)).GetProperty("accessToken").GetString());
        Assert.Equal(HttpStatusCode.OK, (await next.GetAsync("/api/v1/products")).StatusCode);
        var me = await Json(await next.GetAsync("/api/v1/me"));
        Assert.Contains("PRODUCTS", me.GetProperty("permissions").EnumerateArray().Select(p => p.GetString()));
        Assert.Equal(4, me.GetProperty("screens").GetArrayLength());

        await using var db = host.Db.CreateDbContext();
        var login = await db.AuditLog.Where(l => l.ActionType == "LOGIN").OrderBy(l => l.LogID).FirstAsync();
        Assert.Equal("api", login.Details);
        Assert.StartsWith("API TestPhone", login.ComputerName);
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "LOGIN_FAILED"));

        // a disabled user loses his token at once
        (await db.Employees.FirstAsync(e => e.EmployeeID == 1)).IsActive = false;
        await db.SaveChangesAsync();
        Assert.Equal(HttpStatusCode.Unauthorized, (await next.GetAsync("/api/v1/products")).StatusCode);
    }

    [Fact]
    public async Task A_forged_token_is_refused_and_the_login_is_rate_limited()
    {
        using var host = new ApiHost(loginsPerMinute: 3);
        var c = host.CreateClient();
        c.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "eyJhbGciOiJub25lIn0.eyJzdWIiOiIxIn0.");
        Assert.Equal(HttpStatusCode.Unauthorized, (await c.GetAsync("/api/v1/me")).StatusCode);
        var codes = new List<HttpStatusCode>();
        for (int i = 0; i < 4; i++)
        {
            codes.Add((await c.PostAsJsonAsync("/api/v1/auth/token", new { username = "nobody", password = "x" })).StatusCode);
        }
        Assert.Equal(HttpStatusCode.TooManyRequests, codes[^1]);
    }

    [Fact]
    public async Task Data_screens_save_with_the_checks_questions_and_audit_of_the_screen()
    {
        using var host = new ApiHost();
        var c = await host.AdminAsync();
        var bad = await c.PostAsJsonAsync("/api/v1/products", new Dictionary<string, object?> { ["ProductName"] = "قلم", ["Colour"] = "x" });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, bad.StatusCode);
        Assert.Equal("Colour", (await Json(bad)).GetProperty("field").GetString());

        var missing = await c.PostAsJsonAsync("/api/v1/products", new Dictionary<string, object?> { ["SellingPrice"] = 5 });
        var missingBody = await Json(missing);
        Assert.Equal("ProductName", missingBody.GetProperty("field").GetString());
        Assert.Equal("الحقل «اسم المنتج» مطلوب.", missingBody.GetProperty("title").GetString());

        var body = new Dictionary<string, object?>
        {
            ["ProductName"] = "قلم", ["SellingPrice"] = 2, ["PurchasePrice"] = 3, ["CurrentQuantity"] = 99,
        };
        var ask = await c.PostAsJsonAsync("/api/v1/products", body);
        Assert.Equal(HttpStatusCode.Conflict, ask.StatusCode);
        var question = await Json(ask);
        Assert.Equal("confirm", question.GetProperty("code").GetString());
        Assert.Equal("SellingPrice", question.GetProperty("field").GetString());
        var made = await c.PostAsJsonAsync("/api/v1/products?confirm=true", body);
        Assert.Equal(HttpStatusCode.Created, made.StatusCode);
        var record = (await Json(made)).GetProperty("record");
        Assert.Equal("P00001", record.GetProperty("ProductCode").GetString());
        Assert.Equal(0m, record.GetProperty("CurrentQuantity").GetDecimal());          // locked: ignored
        Assert.Equal(3m, record.GetProperty("AverageCost").GetDecimal());
        int id = record.GetProperty("ProductID").GetInt32();

        var change = await c.PutAsJsonAsync($"/api/v1/products/{id}?confirm=true", new Dictionary<string, object?> { ["SellingPrice"] = "4.5" });
        Assert.Equal(HttpStatusCode.OK, change.StatusCode);
        var now = await Json(await c.GetAsync($"/api/v1/products/{id}"));
        Assert.Equal(4.5m, now.GetProperty("SellingPrice").GetDecimal());
        Assert.Equal("قلم", now.GetProperty("ProductName").GetString());                // not sent: kept
        Assert.Equal(HttpStatusCode.NotFound, (await c.GetAsync("/api/v1/products/999")).StatusCode);

        var english = new HttpRequestMessage(HttpMethod.Post, "/api/v1/customers") { Content = JsonContent.Create(new { VATNumber = "123" }) };
        english.Headers.AcceptLanguage.ParseAdd("en");
        var refused = await c.SendAsync(english);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, refused.StatusCode);
        Assert.DoesNotMatch("[؀-ۿ]", (await Json(refused)).GetProperty("title").GetString()!);

        // a supplier used by a product: deactivated instead of deleted
        var supplier = await Json(await c.PostAsJsonAsync("/api/v1/suppliers", new { SupplierName = "مورد" }));
        int sid = supplier.GetProperty("key").GetInt32();
        await c.PutAsJsonAsync($"/api/v1/products/{id}?confirm=true", new { SupplierID = sid });
        var used = await c.DeleteAsync($"/api/v1/suppliers/{sid}");
        Assert.Equal(HttpStatusCode.Conflict, used.StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await c.DeleteAsync($"/api/v1/suppliers/{sid}?deactivate=true")).StatusCode);
        var active = await Json(await c.GetAsync("/api/v1/suppliers?search=مورد"));
        Assert.Equal(0, active.GetProperty("total").GetInt32());
        var all = await Json(await c.GetAsync("/api/v1/suppliers?search=مورد&includeInactive=true"));
        Assert.Equal(1, all.GetProperty("total").GetInt32());

        var lookups = await Json(await c.GetAsync("/api/v1/products/lookups"));
        Assert.True(lookups.GetProperty("CategoryID").GetArrayLength() > 0);
        Assert.Equal(3, lookups.GetProperty("VATCategory").GetArrayLength());

        await using var db = host.Db.CreateDbContext();
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "ADD" && l.ObjectName == "Products" && l.ComputerName!.StartsWith("API")));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "EDIT" && l.ObjectName == "Products"));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "DEACTIVATE" && l.ObjectName == "Suppliers"));
    }

    [Fact]
    public async Task Lists_have_pages()
    {
        using var host = new ApiHost();
        var c = await host.AdminAsync();
        for (int i = 1; i <= 5; i++)
        {
            var made = await c.PostAsJsonAsync("/api/v1/warehouses", new { WarehouseName = "مخزن " + i, WarehouseType = "BRANCH" });
            Assert.True(made.StatusCode == HttpStatusCode.Created, await made.Content.ReadAsStringAsync());
        }
        var page = await Json(await c.GetAsync("/api/v1/warehouses?search=مخزن&skip=2&take=2"));
        Assert.Equal(6, page.GetProperty("total").GetInt32());                    // with the main warehouse, in the order of the screen
        Assert.Equal(2, page.GetProperty("items").GetArrayLength());
        Assert.Equal("مخزن 2", page.GetProperty("items")[0].GetProperty("WarehouseName").GetString());
    }

    /// <summary>A rep user: his own screens are the visits and the customer orders only.</summary>
    internal static async Task<(int Rep, int Customer, int Van, int Product)> MakeRepAsync(ApiHost host)
    {
        await using var db = host.Db.CreateDbContext();
        var user = new Employees { EmployeeName = "مندوب", Username = "rep1", RoleID = 3, CustomScreens = true, MustChangePassword = false };
        db.Employees.Add(user);
        await db.SaveChangesAsync();
        db.UserScreens.Add(new UserScreens { EmployeeID = user.EmployeeID, ScreenName = "frmRepVisits", CanOpen = true, CanAdd = true, CanEdit = true, CanDelete = true });
        db.UserScreens.Add(new UserScreens { EmployeeID = user.EmployeeID, ScreenName = "frmSalesOrder", CanOpen = true, CanAdd = true });
        var rep = new SalesReps { RepCode = "R1", RepName = "مندوب الشمال", EmployeeID = user.EmployeeID, CreditLimit = 5000 };
        db.SalesReps.Add(rep);
        var list = new PriceLists { ListName = "جملة", DiscountPercent = 0.1m };
        db.PriceLists.Add(list);
        await db.SaveChangesAsync();
        var customer = new Customers { CustomerName = "بقالة النور", SalesRepID = rep.SalesRepID, PriceListID = list.PriceListID, CurrentBalance = 300 };
        var other = new Customers { CustomerName = "عميل المسار" };
        db.Customers.AddRange(customer, other);
        var van = new Warehouses { WarehouseName = "سيارة 1", WarehouseType = "VAN", SalesRepID = rep.SalesRepID };
        db.Warehouses.Add(van);
        var a = new Products { ProductCode = "A1", ProductName = "أرز", SellingPrice = 20, PurchasePrice = 15 };
        var b = new Products { ProductCode = "B1", ProductName = "سكر", SellingPrice = 10, PurchasePrice = 8 };
        var service = new Products { ProductCode = "S1", ProductName = "توصيل", SellingPrice = 5, TrackStock = false };
        db.Products.AddRange(a, b, service);
        await db.SaveChangesAsync();
        db.RepPrices.Add(new RepPrices { SalesRepID = rep.SalesRepID, ProductID = b.ProductID, UnitPrice = 9 });
        db.PriceListItems.Add(new PriceListItems { PriceListID = list.PriceListID, ProductID = a.ProductID, UnitPrice = 18 });
        db.PriceListItems.Add(new PriceListItems { PriceListID = list.PriceListID, ProductID = a.ProductID, UnitPrice = 17,
                                                   ValidFrom = DateTime.Today.AddDays(-3) });
        db.PriceListItems.Add(new PriceListItems { PriceListID = list.PriceListID, ProductID = a.ProductID, UnitPrice = 1,
                                                   ValidFrom = DateTime.Today.AddDays(5) });            // not yet
        db.RouteStops.Add(new RouteStops { SalesRepID = rep.SalesRepID, CustomerID = customer.CustomerID,
                                           VisitWeekday = (byte)RepService.VisitWeekdayOf(DateTime.Today), StopOrder = 1 });
        db.RouteStops.Add(new RouteStops { SalesRepID = rep.SalesRepID, CustomerID = other.CustomerID, VisitWeekday = 1, StopOrder = 2 });
        db.InventoryTransactions.Add(new InventoryTransactions { ProductID = a.ProductID, TransactionTypeID = 1, Quantity = 12, WarehouseID = van.WarehouseID });
        db.InventoryTransactions.Add(new InventoryTransactions { ProductID = a.ProductID, TransactionTypeID = 1, Quantity = -2, WarehouseID = van.WarehouseID });
        await db.SaveChangesAsync();
        await new SecurityService(host.Db).SetUserPasswordAsync(user.EmployeeID, "Rep@2026x", false, TestDb.Admin);
        return (rep.SalesRepID, customer.CustomerID, van.WarehouseID, a.ProductID);
    }

    [Fact]
    public async Task The_rep_works_on_his_own_data_with_his_permissions()
    {
        using var host = new ApiHost();
        var (repId, customerId, vanId, riceId) = await MakeRepAsync(host);
        var c = await host.ClientAsync("rep1", "Rep@2026x");

        var forbidden = await c.GetAsync("/api/v1/customers");                     // not one of his screens
        Assert.Equal(HttpStatusCode.Forbidden, forbidden.StatusCode);
        Assert.Equal("forbidden", await Code(forbidden));

        var profile = await Json(await c.GetAsync("/api/v1/rep/profile"));
        Assert.Equal(repId, profile.GetProperty("salesRepID").GetInt32());
        Assert.Equal(vanId, profile.GetProperty("vanWarehouseID").GetInt32());
        Assert.Equal(300m, profile.GetProperty("customersBalance").GetDecimal());
        Assert.Equal(1, profile.GetProperty("orderWarehouseID").GetInt32());

        var customers = await Json(await c.GetAsync("/api/v1/rep/customers"));
        Assert.Equal(2, customers.GetArrayLength());                              // assigned + on his route
        var products = (await Json(await c.GetAsync($"/api/v1/rep/products?customerId={customerId}"))).EnumerateArray()
            .ToDictionary(p => p.GetProperty("productCode").GetString()!);
        Assert.Equal(17m, products["A1"].GetProperty("price").GetDecimal());       // the list item that starts last
        Assert.Equal(10m, products["A1"].GetProperty("vanQuantity").GetDecimal());
        Assert.Equal(8.1m, products["B1"].GetProperty("price").GetDecimal());      // rep price 9 less the list's 10 %
        var plain = (await Json(await c.GetAsync("/api/v1/rep/products"))).EnumerateArray()
            .ToDictionary(p => p.GetProperty("productCode").GetString()!);
        Assert.Equal(9m, plain["B1"].GetProperty("price").GetDecimal());
        Assert.Equal(20m, plain["A1"].GetProperty("price").GetDecimal());
        var stock = await Json(await c.GetAsync("/api/v1/rep/van-stock"));
        Assert.Equal(10m, stock[0].GetProperty("quantity").GetDecimal());

        // the plan of today, once
        var plan = await Json(await c.PostAsJsonAsync("/api/v1/rep/visits/plan", new { }));
        Assert.Equal(1, plan.GetProperty("id").GetInt32());
        Assert.Equal(0, (await Json(await c.PostAsJsonAsync("/api/v1/rep/visits/plan", new { }))).GetProperty("id").GetInt32());
        var visits = await Json(await c.GetAsync("/api/v1/rep/visits"));
        int visitId = visits[0].GetProperty("repVisitID").GetInt32();
        Assert.Equal("PENDING", visits[0].GetProperty("result").GetString());
        var noResult = await c.PutAsJsonAsync($"/api/v1/rep/visits/{visitId}", new { });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, noResult.StatusCode);
        Assert.StartsWith("اختر نتيجة الزيارة", (await Json(noResult)).GetProperty("title").GetString());
        var twice = await c.PostAsJsonAsync("/api/v1/rep/visits", new { customerID = customerId });
        Assert.Equal("العميل في زيارات هذا اليوم.", (await Json(twice)).GetProperty("title").GetString());

        // an order: checked lines, default prices, the rep of the customer
        var serviceLine = await c.PostAsJsonAsync("/api/v1/rep/orders", new
        {
            customerID = customerId, lines = new[] { new { productID = products["S1"].GetProperty("productID").GetInt32(), quantity = 1 } },
        });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, serviceLine.StatusCode);
        var order = await c.PostAsJsonAsync("/api/v1/rep/orders", new
        {
            customerID = customerId, notes = "من التطبيق",
            lines = new object[] { new { productID = riceId, quantity = 3 }, new { productID = products["B1"].GetProperty("productID").GetInt32(), quantity = 2, unitPrice = 9.5 } },
        });
        Assert.Equal(HttpStatusCode.Created, order.StatusCode);
        int orderId = (await Json(order)).GetProperty("id").GetInt32();
        var view = await Json(await c.GetAsync($"/api/v1/rep/orders/{orderId}"));
        Assert.Equal("SO-000001", view.GetProperty("orderNumber").GetString());
        Assert.Equal("DRAFT", view.GetProperty("status").GetString());
        Assert.Equal(repId, view.GetProperty("salesRepID").GetInt32());
        Assert.Equal(17m, view.GetProperty("lines")[0].GetProperty("unitPrice").GetDecimal());
        Assert.Equal(9.5m, view.GetProperty("lines")[1].GetProperty("unitPrice").GetDecimal());
        Assert.Equal(70m, view.GetProperty("total").GetDecimal());
        Assert.Equal(1, (await Json(await c.GetAsync("/api/v1/rep/orders"))).GetArrayLength());

        // the visit result is read from the order of the day
        Assert.Equal(HttpStatusCode.OK, (await c.PutAsJsonAsync($"/api/v1/rep/visits/{visitId}", new { notes = "تم" })).StatusCode);
        var recorded = await Json(await c.GetAsync("/api/v1/rep/visits"));
        Assert.Equal("ORDER", recorded[0].GetProperty("result").GetString());
        Assert.Equal(HttpStatusCode.UnprocessableEntity, (await c.DeleteAsync($"/api/v1/rep/visits/{visitId}")).StatusCode);
        var extra = await Json(await c.PostAsJsonAsync("/api/v1/rep/visits", new { customerID = 1 }));
        Assert.Equal(HttpStatusCode.OK, (await c.DeleteAsync($"/api/v1/rep/visits/{extra.GetProperty("id").GetInt32()}")).StatusCode);

        // a customer of another rep on this rep's route: the price of the order's rep (SalesRepFor), in the list and the order
        int otherCustomer;
        int sugar = products["B1"].GetProperty("productID").GetInt32();
        await using (var setup = host.Db.CreateDbContext())
        {
            var rep2 = new SalesReps { RepCode = "R2", RepName = "مندوب الجنوب" };
            setup.SalesReps.Add(rep2);
            await setup.SaveChangesAsync();
            setup.RepPrices.Add(new RepPrices { SalesRepID = rep2.SalesRepID, ProductID = sugar, UnitPrice = 7 });
            var c2 = await setup.Customers.FirstAsync(x => x.CustomerName == "عميل المسار");
            c2.SalesRepID = rep2.SalesRepID;
            otherCustomer = c2.CustomerID;
            await setup.SaveChangesAsync();
        }
        var theirs = (await Json(await c.GetAsync($"/api/v1/rep/products?customerId={otherCustomer}"))).EnumerateArray()
            .Single(p => p.GetProperty("productID").GetInt32() == sugar);
        Assert.Equal(7m, theirs.GetProperty("price").GetDecimal());
        var theirOrder = await Json(await c.PostAsJsonAsync("/api/v1/rep/orders",
            new { customerID = otherCustomer, lines = new[] { new { productID = sugar, quantity = 1 } } }));
        Assert.Equal(2, (await Json(await c.GetAsync("/api/v1/rep/orders"))).GetArrayLength());      // entered by him
        await using var db = host.Db.CreateDbContext();
        Assert.Equal(7m, (await db.SalesOrderLines.SingleAsync(l => l.SalesOrderID == theirOrder.GetProperty("id").GetInt32())).UnitPrice);
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "REP_VISIT_PLAN"));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "REP_VISIT_RECORD"));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "REP_VISIT_DELETE"));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "SALES_ORDER_NEW"));
        Assert.Equal(3, await db.AuditLog.CountAsync(l => l.ActionType == "ADD" && l.ObjectName == "SalesOrderLines"));
    }

    [Fact]
    public async Task A_user_who_is_not_a_rep_and_a_visit_of_another_rep_are_refused()
    {
        using var host = new ApiHost();
        await MakeRepAsync(host);
        var admin = await host.AdminAsync();
        var r = await admin.GetAsync("/api/v1/rep/profile");
        Assert.Equal(HttpStatusCode.Forbidden, r.StatusCode);
        Assert.Equal("not_a_rep", await Code(r));

        await using (var db = host.Db.CreateDbContext())
        {
            var otherRep = new SalesReps { RepCode = "R2", RepName = "مندوب آخر" };
            db.SalesReps.Add(otherRep);
            await db.SaveChangesAsync();
            db.RepVisits.Add(new RepVisits { SalesRepID = otherRep.SalesRepID, CustomerID = 1, VisitDate = DateTime.Today });
            await db.SaveChangesAsync();
        }
        var rep = await host.ClientAsync("rep1", "Rep@2026x");
        int foreign;
        await using (var db = host.Db.CreateDbContext())
        {
            foreign = await db.RepVisits.Select(v => v.RepVisitID).MaxAsync();
        }
        Assert.Equal(HttpStatusCode.UnprocessableEntity, (await rep.PutAsJsonAsync($"/api/v1/rep/visits/{foreign}", new { result = "SALE" })).StatusCode);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, (await rep.DeleteAsync($"/api/v1/rep/visits/{foreign}")).StatusCode);
    }

    [Fact]
    public async Task A_sale_and_a_receipt_through_the_api_follow_modSales()
    {
        using var host = new ApiHost();
        var admin = await host.AdminAsync();
        int rice, ali;
        await using (var db = host.Db.CreateDbContext())
        {
            var p = new Products { ProductCode = "R1", Barcode = "628100", ProductName = "أرز", SellingPrice = 23, CurrentQuantity = 1 };
            var c = new Customers { CustomerName = "علي" };
            db.AddRange(p, c);
            await db.SaveChangesAsync();
            (rice, ali) = (p.ProductID, c.CustomerID);
        }
        var found = await Json(await admin.GetAsync("/api/v1/sales/products?code=628100"));
        Assert.Equal(23m, found[0].GetProperty("price").GetDecimal());
        var sale = new { customerID = ali, paymentType = "CREDIT", invoiceDiscount = 0, tendered = 0,
                         lines = new[] { new { productID = rice, quantity = 2, unitPrice = 23 } } };
        var ask = await admin.PostAsJsonAsync("/api/v1/sales", sale);
        Assert.Equal(HttpStatusCode.Conflict, ask.StatusCode);                       // more than the stock: the question
        Assert.Equal("confirm", await Code(ask));
        var made = await admin.PostAsJsonAsync("/api/v1/sales?confirm=true", sale);
        Assert.Equal(HttpStatusCode.Created, made.StatusCode);
        var id = (await Json(made)).GetProperty("id").GetInt32();
        var invoice = await Json(await admin.GetAsync($"/api/v1/sales/{id}"));
        Assert.Equal(46m, invoice.GetProperty("remainingAmount").GetDecimal());
        Assert.Single((await Json(await admin.GetAsync("/api/v1/sales"))).EnumerateArray());

        var pay = await admin.PostAsJsonAsync("/api/v1/customer-payments", new { customerID = ali, amount = 50, paymentMethodID = 1 });
        Assert.Equal(HttpStatusCode.Conflict, pay.StatusCode);
        pay = await admin.PostAsJsonAsync("/api/v1/customer-payments?confirm=true", new { customerID = ali, amount = 50, paymentMethodID = 1 });
        Assert.Equal(HttpStatusCode.Created, pay.StatusCode);
        var cash = await admin.PostAsJsonAsync("/api/v1/customer-payments?lang=en", new { customerID = 1, amount = 5, paymentMethodID = 1 });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, cash.StatusCode);
        Assert.DoesNotContain((await Json(cash)).GetProperty("title").GetString()!, c => c >= '\u0600' && c <= '\u06FF');

        await using var check = host.Db.CreateDbContext();
        Assert.Equal(-4m, (await check.Customers.FindAsync(ali))!.CurrentBalance);
        var cashier = new Employees { EmployeeName = "مخزن", Username = "store1", RoleID = 2 };
        check.Employees.Add(cashier);
        await check.SaveChangesAsync();
        check.RolePermissions.RemoveRange(check.RolePermissions.Where(r => r.RoleID == 2 && r.PermissionKey == "SALES_POS"));
        await check.SaveChangesAsync();
        await new SecurityService(host.Db).SetUserPasswordAsync(cashier.EmployeeID, ApiHost.AdminPassword, false, TestDb.Admin);
        var other = await host.ClientAsync("store1", ApiHost.AdminPassword);
        Assert.Equal(HttpStatusCode.Forbidden, (await other.PostAsJsonAsync("/api/v1/sales", sale)).StatusCode);
    }

    [Fact]
    public async Task The_OpenAPI_document_lists_every_operation_with_its_permission()
    {
        using var host = new ApiHost();
        var c = host.CreateClient();
        var doc = await Json(await c.GetAsync("/api/openapi/v1.json"));
        Assert.Equal("RetailStore API", doc.GetProperty("info").GetProperty("title").GetString());
        Assert.True(doc.GetProperty("components").GetProperty("securitySchemes").TryGetProperty("Bearer", out _));
        var open = new HashSet<string> { "/api/v1/auth/token", "/api/v1/auth/password", "/api/v1/me", "/api/v1/rep/profile" };
        int operations = 0;
        foreach (var path in doc.GetProperty("paths").EnumerateObject())
        {
            Assert.StartsWith("/api/v1/", path.Name);
            foreach (var op in path.Value.EnumerateObject())
            {
                operations++;
                Assert.True(op.Value.TryGetProperty("summary", out _), path.Name);
                Assert.Equal(!open.Contains(path.Name), op.Value.TryGetProperty("x-permission", out _));
                Assert.Equal(path.Name != "/api/v1/auth/token", op.Value.TryGetProperty("security", out _));
            }
        }
        Assert.Equal(3 + 4 * 6 + 13 + 6, operations);
        Assert.Equal(HttpStatusCode.OK, (await c.GetAsync("/api/docs")).StatusCode);
    }
}
