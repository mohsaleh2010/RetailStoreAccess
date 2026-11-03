using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Services;
using RetailStore.Domain.Localization;
using RetailStore.RepApp;
using RetailStore.RepApp.Client;
using RetailStore.RepApp.Offline;
using RetailStore.RepApp.Printing;

namespace RetailStore.Tests;

/// <summary>The rep app (docs/78) on a phone: its own SQLite file, the API of a test host.</summary>
public sealed class RepPhone : IDisposable
{
    private readonly string _file = Path.Combine(Path.GetTempPath(), "rep-" + Guid.NewGuid().ToString("N") + ".db");

    public RepPhone(ApiHost host)
    {
        Store = new LocalStore(_file);
        Api = new RepApiClient(host.CreateClient());
        Session = new RepSession(Store, Api, Platform);
    }

    public LocalStore Store { get; }
    public RepApiClient Api { get; }
    public BasicPlatform Platform { get; } = new();
    public RepSession Session { get; }

    /// <summary>The same phone after the app is closed and opened again: a new session on the same file.</summary>
    public RepSession Reopen() => new(new LocalStore(_file), Api, Platform);

    public void Dispose() => File.Delete(_file);
}

/// <summary>The rep app (docs/78): its offline work and its sync against the real API on a test database.</summary>
public class RepAppTests
{
    private const string Server = "http://localhost";

    private static async Task<RepPhone> LoggedInAsync(ApiHost host)
    {
        var phone = new RepPhone(host);
        await phone.Session.LoadAsync();
        Assert.Equal("", await phone.Session.LoginAsync(Server, "rep1", "Rep@2026x"));
        var first = await phone.Session.SyncAsync();
        Assert.True(first.SnapshotAt is not null, first.Message);
        return phone;
    }

    [Fact]
    public async Task An_order_and_a_visit_taken_offline_reach_the_server_once()
    {
        using var host = new ApiHost();
        var (_, customerId, _, riceId) = await ApiTests.MakeRepAsync(host);
        using var phone = await LoggedInAsync(host);
        var s = phone.Session;
        Assert.Equal(2, s.Snapshot!.Customers.Count);
        Assert.Equal(300m, s.Snapshot.Profile.CustomersBalance);
        Assert.Equal("مندوب الشمال", s.Snapshot.Profile.RepName);

        // no network: everything is kept on the phone
        phone.Platform.HasNetwork = false;
        var draft = s.NewOrder(customerId)!;
        int service = s.Snapshot.Products.Single(p => p.ProductCode == "S1").ProductID;
        Assert.Equal("هذا الصنف لا يتابع المخزون، فلا يُحجز ولا يُسلَّم بإذن.", draft.Add(service, 1));
        Assert.Equal("الكمية أكبر من صفر.", draft.Add(riceId, 0));
        Assert.Equal("", draft.Add(riceId, 2));
        Assert.Equal("", draft.Add(riceId, 1));                      // the same product: one line of 3
        Assert.Single(draft.Lines);
        Assert.Equal(17m, draft.Lines[0].UnitPrice);                  // the item of the customer's list that starts last
        Assert.Equal(51m, draft.Total);
        var (problem, order) = s.SaveOrder(draft);
        Assert.Equal("", problem);
        s.PlanDay(DateTime.Today);
        var visit = Assert.Single(s.Visits(DateTime.Today));          // the route stop of today's weekday
        Assert.True(visit.Waiting);
        s.RecordVisit(visit, DateTime.Today, "", "طلب أرز");
        Assert.Equal("AUTO", s.Visits(DateTime.Today)[0].Result);
        var offline = await s.SyncAsync();
        Assert.Null(offline.SnapshotAt);
        Assert.Equal(3, s.Waiting);

        // the app is closed and opened again: nothing is lost
        var again = phone.Reopen();
        await again.LoadAsync();
        Assert.Equal(3, again.Waiting);
        Assert.NotNull(again.Snapshot);

        // the network is back: the order first, then the plan and the result read from the order of the day
        phone.Platform.HasNetwork = true;
        var report = await again.SyncAsync();
        Assert.Equal((3, 0, ""), (report.Sent, report.Rejected, report.Message));
        await using (var db = host.Db.CreateDbContext())
        {
            var saved = await db.SalesOrders.SingleAsync();
            Assert.Equal(order!.ClientRef, saved.ClientRef);
            Assert.Equal("DRAFT", saved.Status);
            Assert.Equal(order.TakenAt, saved.OrderDate, TimeSpan.FromSeconds(1));
            Assert.Equal(17m, (await db.SalesOrderLines.SingleAsync()).UnitPrice);
            var v = await db.RepVisits.SingleAsync();
            Assert.Equal(("ORDER", "طلب أرز"), (v.Result, v.Notes));
        }
        var row = Assert.Single(again.Orders());
        Assert.StartsWith("SO-", row.OrderNumber);
        Assert.Equal("DRAFT", row.Status);
        Assert.Empty(phone.Store.Outbox());                           // the snapshot shows it: the outbox is empty

        // the answer was lost on the way: the phone sends the same order again, the server does not add it twice
        phone.Store.Enqueue(OutboxKinds.Order, order.ClientRef, order, DateTime.Now);
        Assert.Equal(1, (await again.SyncAsync()).Sent);
        await using (var db = host.Db.CreateDbContext())
        {
            Assert.Equal(1, await db.SalesOrders.CountAsync());
            Assert.Equal(1, await db.AuditLog.CountAsync(l => l.ActionType == "SALES_ORDER_NEW"));
        }

        // the customer's copy on the Bluetooth printer
        Assert.Equal("اختر الطابعة في الإعدادات.", await again.PrintAsync(row.Order, row.OrderNumber));
        again.SetPrinter("Test printer", 58);
        Assert.Equal("", await again.PrintAsync(row.Order, row.OrderNumber));
        var (printer, job) = Assert.Single(phone.Platform.Printed);
        Assert.Equal("Test printer", printer);
        Assert.Equal(EscPos.Initialize, job[..2]);
        Assert.Equal(new byte[] { 0x1D, 0x76, 0x30, 0x00, 48, 0 }, job[2..8]);   // 384 dots = 48 bytes a row
    }

    [Fact]
    public async Task The_prices_of_the_phone_are_those_of_the_server_for_every_customer()
    {
        using var host = new ApiHost();
        await ApiTests.MakeRepAsync(host);
        await using (var db = host.Db.CreateDbContext())
        {
            // a customer of another rep on this rep's route: the prices of his rep
            var rep2 = new RetailStore.Data.Entities.SalesReps { RepCode = "R2", RepName = "مندوب الجنوب" };
            db.SalesReps.Add(rep2);
            await db.SaveChangesAsync();
            var other = await db.Customers.SingleAsync(c => c.CustomerName == "عميل المسار");
            other.SalesRepID = rep2.SalesRepID;
            db.RepPrices.Add(new RetailStore.Data.Entities.RepPrices { SalesRepID = rep2.SalesRepID, UnitPrice = 7,
                                                                       ProductID = (await db.Products.SingleAsync(p => p.ProductCode == "B1")).ProductID });
            await db.SaveChangesAsync();
        }
        using var phone = await LoggedInAsync(host);
        var snap = phone.Session.Snapshot!;
        Assert.Equal(2, snap.PriceGroups.Count);
        var c = await host.ClientAsync("rep1", "Rep@2026x");
        foreach (var customer in snap.Customers)
        {
            var server = await c.GetFromJsonAsync<List<RetailStore.RepApp.Client.RepProduct>>(
                $"/api/v1/rep/products?customerId={customer.CustomerID}", RepApiClient.Json);
            foreach (var p in snap.Products)
            {
                Assert.Equal(server!.Single(x => x.ProductID == p.ProductID).Price, RepPricing.PriceOf(snap, customer.CustomerID, p));
            }
        }
    }

    [Fact]
    public async Task A_refusal_stops_only_its_item_and_an_ended_token_keeps_the_outbox()
    {
        using var host = new ApiHost();
        var (_, customerId, _, riceId) = await ApiTests.MakeRepAsync(host);
        using var phone = await LoggedInAsync(host);
        var s = phone.Session;
        var draft = s.NewOrder(customerId)!;
        draft.Add(riceId, 1);
        var (_, order) = s.SaveOrder(draft);
        s.AddVisit(customerId, DateTime.Today);
        await using (var db = host.Db.CreateDbContext())
        {
            (await db.Customers.SingleAsync(c => c.CustomerID == customerId)).IsActive = false;   // the office stops him
            await db.SaveChangesAsync();
        }
        var report = await s.SyncAsync();
        Assert.Equal((1, 1), (report.Sent, report.Rejected));        // the visit goes, the order is refused
        var refused = Assert.Single(s.Problems());
        Assert.Equal("اختر العميل.", refused.Message);
        Assert.Equal(OutboxKinds.Rejected, s.Orders().Single(o => o.Order.ClientRef == order!.ClientRef).State);

        await using (var db = host.Db.CreateDbContext())
        {
            (await db.Customers.SingleAsync(c => c.CustomerID == customerId)).IsActive = true;
            await db.SaveChangesAsync();
        }
        s.RetryItem(refused.Id);
        Assert.Equal(1, (await s.SyncAsync()).Sent);

        // a new password (or a disabled user) ends the token: log in again, nothing recorded is lost
        s.PlanDay(DateTime.Today);
        int repUser;
        await using (var db = host.Db.CreateDbContext())
        {
            repUser = (await db.Employees.SingleAsync(e => e.Username == "rep1")).EmployeeID;
        }
        await new SecurityService(host.Db).SetUserPasswordAsync(repUser, "Rep@2026y", false, TestDb.Admin);
        var ended = await s.SyncAsync();
        Assert.True(ended.NeedsLogin);
        Assert.False(s.LoggedIn);
        Assert.Equal(1, s.Waiting);

        // another user may not take the phone while the rep's work waits
        Assert.Contains("rep1", await s.LoginAsync(Server, "rep2", "x"));
        Assert.Equal("", await s.LoginAsync(Server, "rep1", "Rep@2026y"));
        Assert.Equal(1, (await s.SyncAsync()).Sent);
    }

    [Fact]
    public async Task The_number_of_an_order_on_a_phone_belongs_to_its_user()
    {
        using var host = new ApiHost();
        var (_, customerId, _, riceId) = await ApiTests.MakeRepAsync(host);
        var rep = await host.ClientAsync("rep1", "Rep@2026x");
        var body = new { customerID = customerId, clientRef = "abc", lines = new[] { new { productID = riceId, quantity = 1 } } };
        Assert.Equal(HttpStatusCode.Created, (await rep.PostAsJsonAsync("/api/v1/rep/orders", body)).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await rep.PostAsJsonAsync("/api/v1/rep/orders", body)).StatusCode);
        var admin = await host.AdminAsync();
        Assert.Equal(HttpStatusCode.Forbidden, (await admin.PostAsJsonAsync("/api/v1/rep/orders", body)).StatusCode);   // not a rep
        await using (var db = host.Db.CreateDbContext())
        {
            Assert.Equal(1, await db.SalesOrders.CountAsync());
        }
        var tooLong = await rep.PostAsJsonAsync("/api/v1/rep/orders", new { customerID = customerId, clientRef = new string('x', 41),
                                                                            lines = body.lines });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, tooLong.StatusCode);
        Assert.Equal(RepService.ClientRefTooLong, (await tooLong.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("title").GetString());
    }

    [Fact]
    public void The_slip_is_sent_as_a_picture_in_bands()
    {
        var dots = new bool[300, 10];
        dots[0, 0] = true;
        dots[0, 9] = true;
        dots[299, 8] = true;
        var job = EscPos.Picture(dots, 3);
        var band = EscPos.Raster(dots, 0, 128);
        Assert.Equal(new byte[] { 0x1D, 0x76, 0x30, 0x00, 2, 0, 128, 0, 0x80, 0x40 }, band[..10]);
        // initialize, three bands (128 + 128 + 44 rows), feed 3, cut
        Assert.Equal(2 + 3 * 8 + 300 * 2 + 3 + 4, job.Length);
        Assert.Equal(new byte[] { 0x1B, 0x64, 3, 0x1D, 0x56, 0x42, 0x00 }, job[^7..]);
        var last = EscPos.Raster(dots, 256, 44);
        Assert.Equal(44, last[6]);
        Assert.Equal(0x80, last[8 + 43 * 2 + 1]);                     // column 8 = the first bit of the second byte

        var order = new PendingOrder("0123456789abcdef", 1, "بقالة النور", new DateTime(2026, 11, 2, 9, 30, 0), null, "صباحًا",
                                     [new PendingOrderLine(1, "أرز", 3, 17), new PendingOrderLine(2, "سكر", 1.5m, 8.1m)]);
        var lines = OrderSlip.Build(order, null, "مندوب الشمال", "متجري", t => t);
        Assert.Contains(lines, l => l.Text == "رقم الطلبية: 01234567");
        Assert.Contains(lines, l => l.Text == "3 x 17.00 = 51.00");
        Assert.Contains(lines, l => l.Text == "1.5 x 8.10 = 12.15");
        Assert.Contains(lines, l => l.Text == "الإجمالي: 63.15");
        Assert.Contains(lines, l => l.Text == "لم تُرسل إلى المكتب بعد");
        var english = OrderSlip.Build(order, "SO-00001", "Rep", "Shop", Translator.ToEnglish);
        Assert.Equal("Customer order", english[1].Text);
        Assert.Equal("Order number: SO-00001", english[2].Text);
        Assert.False(Translator.HasArabic(english[^1].Text));
    }
}
