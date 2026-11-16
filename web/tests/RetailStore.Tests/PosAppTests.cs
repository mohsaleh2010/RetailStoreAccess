using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.Domain.Printing;
using RetailStore.Domain.Sales;
using RetailStore.PosApp;
using RetailStore.PosApp.Connection;
using RetailStore.PosApp.Offline;
using RetailStore.PosApp.Printing;

namespace RetailStore.Tests;

/// <summary>The data of the POS with a switch: Down = the database or the server does not answer.</summary>
public sealed class SwitchableBackend(IPosBackend inner) : IPosBackend
{
    public bool Down { get; set; }

    public Task<string> TestAsync(CancellationToken ct = default) =>
        Down ? Task.FromResult("تعذر الاتصال بقاعدة البيانات.") : inner.TestAsync(ct);

    public Task<BackendResult<BackendLogin>> LoginAsync(string username, string password, string device, CancellationToken ct = default) =>
        Down ? Task.FromResult(BackendResult<BackendLogin>.Unreachable("down")) : inner.LoginAsync(username, password, device, ct);

    public Task<BackendResult<PosSnapshot>> SnapshotAsync(PosCredential who, bool english, CancellationToken ct = default) =>
        Down ? Task.FromResult(BackendResult<PosSnapshot>.Unreachable("down")) : inner.SnapshotAsync(who, english, ct);

    public Task<BackendResult<DeviceSaleResult>> SendSaleAsync(PosCredential who, DeviceSaleRequest sale, bool english,
                                                               CancellationToken ct = default) =>
        Down ? Task.FromResult(BackendResult<DeviceSaleResult>.Unreachable("down")) : inner.SendSaleAsync(who, sale, english, ct);

    public void Dispose()
    {
    }
}

/// <summary>A POS PC (docs/83): its own SQLite file and a way to the data (local: the test database directly; remote:
/// the API of a test host).</summary>
public sealed class PosPc : IDisposable
{
    private readonly string _file = Path.Combine(Path.GetTempPath(), "pos-" + Guid.NewGuid().ToString("N") + ".db");

    public PosPc(IPosBackend data)
    {
        Data = new SwitchableBackend(data);
        Store = new PosStore(_file);
        Session = Reopen();
    }

    public SwitchableBackend Data { get; }
    public PosStore Store { get; }
    public BasicPosPlatform Platform { get; } = new();
    public PosSession Session { get; private set; }

    /// <summary>The same PC after the program is closed and opened again.</summary>
    public PosSession Reopen() => Session = new PosSession(new PosStore(_file), Platform, (_, _) => Data);

    public void Dispose() => File.Delete(_file);
}

public class PosAppTests
{
    private const string Password = "Cash@2026x";

    /// <summary>The store of the tests: a cashier (role 3, discount limit 5 %), rice (10 in stock, a box of 10), sugar
    /// (zero rated), a delivery service, Ali (credit 1000) and a price list for Ali.</summary>
    private sealed record Shop(int Cashier, int Ali, int Rice, int Sugar, int Delivery, int Box);

    private static async Task<Shop> MakeShopAsync(TestDb t)
    {
        await using var db = t.CreateDbContext();
        var cashier = new Employees { EmployeeName = "كاشير", Username = "c1", RoleID = 3, MaxDiscountPercent = 0.05m };
        var list = new PriceLists { ListName = "جملة", DiscountPercent = 0.1m };
        db.AddRange(cashier, list);
        await db.SaveChangesAsync();
        var ali = new Customers { CustomerName = "علي", CreditLimit = 1000, PaymentTermsDays = 30, PriceListID = list.PriceListID };
        var rice = new Products { ProductCode = "R1", Barcode = "628100", ProductName = "أرز", SellingPrice = 23, AverageCost = 12, CurrentQuantity = 10 };
        var sugar = new Products { ProductCode = "S1", ProductName = "سكر", SellingPrice = 11.5m, AverageCost = 6, CurrentQuantity = 100, VATCategory = "Z" };
        var delivery = new Products { ProductCode = "D1", ProductName = "توصيل", SellingPrice = 5, TrackStock = false };
        db.AddRange(ali, rice, sugar, delivery);
        await db.SaveChangesAsync();
        db.ProductUnits.Add(new ProductUnits { ProductID = rice.ProductID, UnitID = 2, UnitFactor = 10, Barcode = "628199", SellingPrice = 220 });
        db.PriceListItems.Add(new PriceListItems { PriceListID = list.PriceListID, ProductID = sugar.ProductID, UnitPrice = 10 });
        var s = await db.Settings.FirstAsync();
        s.VATNumber = "300000000000003";
        s.StoreName = "متجري";
        await db.SaveChangesAsync();
        var security = new SecurityService(t);
        await security.SetUserPasswordAsync(cashier.EmployeeID, Password, false, TestDb.Admin);
        await security.SetUserPasswordAsync(1, ApiHost.AdminPassword, false, TestDb.Admin);
        return new Shop(cashier.EmployeeID, ali.CustomerID, rice.ProductID, sugar.ProductID, delivery.ProductID, 2);
    }

    private static readonly PosConnection Local = PosConnection.Empty with { DeviceCode = "POS1" };

    private static readonly PosConnection Remote = PosConnection.Empty with { Mode = PosConnection.Remote, ServerUrl = "http://localhost",
                                                                              DeviceCode = "POS1" };

    private static async Task<PosPc> ReadyAsync(IPosBackend data, PosConnection connection, string user = "c1", string password = Password)
    {
        var pc = new PosPc(data);
        await pc.Session.LoadAsync();
        Assert.True(pc.Session.CanEditDatabase);                        // the first run: no database yet
        Assert.Equal("", await pc.Session.SaveConnectionAsync(connection, "x"));
        Assert.Equal("", await pc.Session.LoginAsync(user, password));
        return pc;
    }

    [Fact]
    public async Task Sales_made_offline_are_kept_numbered_and_reach_the_server_once()
    {
        using var host = new ApiHost();
        var shop = await MakeShopAsync(host.Db);
        using var pc = await ReadyAsync(new ApiBackend(host.CreateClient(), ""), Remote);
        var s = pc.Session;
        Assert.False(s.User!.IsAdmin);
        Assert.False(s.CanEditDatabase);
        Assert.NotNull(await pc.Platform.GetSecretAsync(PosSession.TokenPrefix + shop.Cashier));

        // online: the sale is kept, then sent at once
        var cart = s.NewCart()!;
        Assert.Equal("", cart.Add(cart.Find("628100").Single().Product.ProductID, null, 2));
        Assert.Equal("", cart.Add(shop.Delivery));
        cart.Tendered = 100;
        var (status, message, first) = s.SaveSale(cart);
        Assert.Equal((SaveStatus.Saved, ""), (status, message));
        Assert.Equal(("POS1-000001", 51m, 49m), (first!.InvoiceNumber, first.Receipt.Total, first.Receipt.Change));
        Assert.Equal((1, 0), ((await s.SyncAsync()).Sent, s.Waiting));

        // the server does not answer: the cashier keeps selling, the stock and the credit of the unsent sales count
        pc.Data.Down = true;
        cart = s.NewCart()!;
        var box = cart.Find("628199").Single();
        Assert.Equal(shop.Box, box.Unit!.UnitID);
        cart.Add(box.Product.ProductID, box.Unit.UnitID);
        var (refused, why, _) = s.SaveSale(cart);
        Assert.Equal(SaveStatus.Refused, refused);
        Assert.Equal("الكمية المتوفرة غير كافية:\n- أرز: المطلوب 10 والمتوفر 8\nالبيع بأكثر من المتوفر يحتاج موافقة مدير النظام.", why);
        cart.Clear();
        cart.Add(shop.Rice, null, 8);
        Assert.Equal(SaveStatus.Saved, s.SaveSale(cart).Status);
        cart = s.NewCart()!;
        cart.ChangeCustomer(shop.Ali);
        cart.Add(shop.Sugar, null, 3);
        Assert.Equal(10m, cart.Items[0].UnitPrice);                        // the item of Ali's price list
        cart.PaymentType = "CREDIT";
        cart.PaymentMethodID = null;
        var (saved, _, credit) = s.SaveSale(cart);
        Assert.Equal(SaveStatus.Saved, saved);
        Assert.Equal(("POS1-000003", 30m), (credit!.InvoiceNumber, credit.Receipt.Remaining));
        var offline = await s.SyncAsync();
        Assert.Equal((false, 2), (offline.Reached, offline.Waiting));
        Assert.True(s.WorkingOffline);
        Assert.Equal(30m, s.Pending().Credit[shop.Ali]);

        // the program is closed and opened again: the cashier logs in without the server, nothing is lost
        var again = pc.Reopen();
        await again.LoadAsync();
        Assert.Equal("تعذر الاتصال بقاعدة البيانات، ولم يدخل هذا المستخدم بكلمة المرور هذه على هذا الجهاز من قبل.",
                     await again.LoginAsync("c1", "wrong"));
        Assert.Equal("", await again.LoginAsync("c1", Password));
        Assert.True(again.WorkingOffline);
        Assert.Equal(2, again.Waiting);

        // the server is back: the two sales in their order, with their numbers, times and QR
        pc.Data.Down = false;
        var report = await again.SyncAsync(refresh: true);
        Assert.Equal((true, 2, 0, 0), (report.Reached, report.Sent, report.Rejected, report.Waiting));
        Assert.False(again.WorkingOffline);
        await using (var db = host.Db.CreateDbContext())
        {
            var invoices = await db.SalesInvoices.OrderBy(i => i.SalesInvoiceID).ToListAsync();
            Assert.Equal(["POS1-000001", "POS1-000002", "POS1-000003"], invoices.Select(i => i.InvoiceNumber));
            var last = invoices[2];
            Assert.Equal((credit.Request.ClientRef, "CREDIT", 30m, shop.Cashier), (last.ClientRef, last.PaymentType, last.RemainingAmount, last.EmployeeID));
            Assert.Equal(credit.Request.TakenAt, last.InvoiceDate);
            Assert.Equal(credit.Receipt.Qr, last.QRCodeData);              // the QR printed is the QR saved
            Assert.Equal(30m, (await db.Customers.FindAsync(shop.Ali))!.CurrentBalance);
            Assert.Equal(0m, (await db.Products.FindAsync(shop.Rice))!.CurrentQuantity);
            Assert.Equal(1, await db.Sequences.Where(q => q.SequenceName == "SALES_INVOICE").Select(q => q.NextValue).SingleAsync());
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "SALE" && a.RecordID == "POS1-000002"
                                                        && a.Details!.Contains(" pos ")));
        }
        Assert.Equal(0m, again.Snapshot!.Products.Single(p => p.ProductID == shop.Rice).Available);   // the snapshot read again
        Assert.Equal(30m, again.Snapshot.Customers.Single(c => c.CustomerID == shop.Ali).CurrentBalance);

        // the answer was lost on the way: the PC sends a sale again, the server adds nothing
        pc.Store.MarkRejected(credit.Id, "lost");
        pc.Store.Retry(credit.Id, false);
        Assert.Equal(1, (await again.SyncAsync()).Sent);
        Assert.Equal(SaleStates.Sent, pc.Store.Sales().Single(x => x.Id == credit.Id).State);
        await using (var db = host.Db.CreateDbContext())
        {
            Assert.Equal(3, await db.SalesInvoices.CountAsync());
        }
    }

    [Fact]
    public async Task A_local_database_gives_the_same_checks_and_messages_as_the_sale_of_the_program()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        using var pc = await ReadyAsync(new DirectBackend(t), Local);
        var s = pc.Session;
        var sales = new SalesService(t, t.Permissions);
        var user = new ActingUser(shop.Cashier, "TEST");

        async Task Same(Action<PosCart> fill, string expected)
        {
            var cart = s.NewCart()!;
            fill(cart);
            var (_, message, _) = s.SaveSale(cart);
            Assert.Equal(expected, message);
            var server = await sales.PostSaleAsync(user, new SaleRequest(cart.CustomerID, cart.PaymentType, cart.PaymentMethodID, cart.InvoiceDiscount,
                cart.Tendered, null, cart.Items.Select(i => new CartLine(i.Product.ProductID, i.Quantity, i.UnitPrice, i.LineDiscount, i.UnitID)).ToList()));
            Assert.Equal(expected, server.Message);
        }

        await Same(c => { c.Add(shop.Rice); c.Items[0].UnitPrice = 20; }, "لا تملك صلاحية تعديل سعر البيع.");
        await Same(c => { c.Add(shop.Rice); c.InvoiceDiscount = 3; }, "الخصم (13.0%) أكبر من الحد المسموح لك (5.0%).");
        await Same(c => { c.Add(shop.Rice); c.Tendered = 10; }, "المبلغ المدفوع (10.00) أقل من إجمالي الفاتورة (23.00).");
        await Same(c => { c.Add(shop.Rice); c.PaymentType = "CREDIT"; },
                   "العميل «عميل نقدي» غير مسموح له بالبيع الآجل.\nللبيع الآجل اختر عميلًا مسجلًا.");
        await Same(c => { c.ChangeCustomer(shop.Ali); c.Add(shop.Sugar, null, 101); c.PaymentType = "CREDIT"; },
                   "تجاوز حد الائتمان: الرصيد 0.00 + المتبقي 1,010.00 أكبر من الحد 1,000.00.");
        await Same(c => c.Add(shop.Rice, null, 11),
                   "الكمية المتوفرة غير كافية:\n- أرز: المطلوب 11 والمتوفر 10\nالبيع بأكثر من المتوفر يحتاج موافقة مدير النظام.");
        await using (var db = t.CreateDbContext())
        {
            Assert.Equal(0, await db.SalesInvoices.CountAsync());
        }

        // a sale of the box unit and a bigger unit price: sent directly, the stock in pieces
        var cart = s.NewCart()!;
        cart.Add(shop.Rice, shop.Box);
        Assert.Equal(220m, cart.Items[0].UnitPrice);
        cart.Add(shop.Delivery);
        Assert.Equal(SaveStatus.Saved, s.SaveSale(cart).Status);
        Assert.Equal(1, (await s.SyncAsync()).Sent);
        await using (var db = t.CreateDbContext())
        {
            var line = await db.SalesInvoiceDetails.OrderBy(d => d.LineNumber).FirstAsync();
            Assert.Equal(10m, line.UnitFactor);
            Assert.Equal(0m, (await db.Products.FindAsync(shop.Rice))!.CurrentQuantity);
            Assert.Equal("POS1-000001", (await db.SalesInvoices.SingleAsync()).InvoiceNumber);
        }
    }

    [Fact]
    public async Task The_prices_of_the_pc_are_those_of_the_program_for_every_customer_and_unit()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        var sales = new SalesService(t, t.Permissions);
        var snapshot = (await new PosService(t, t.Permissions, sales).SnapshotAsync(shop.Cashier, false))!;
        Assert.Single(snapshot.PriceGroups);
        var cart = new PosCart(snapshot);
        foreach (var c in snapshot.Customers)
        {
            foreach (var p in snapshot.Products)
            {
                foreach (int? unit in new int?[] { null }.Concat(p.Units.Select(u => (int?)u.UnitID)))
                {
                    Assert.Equal(await sales.LinePriceAsync(shop.Cashier, p.ProductID, unit, c.CustomerID), cart.PriceFor(p, unit, c.CustomerID));
                }
            }
        }
        Assert.Equal(198m, cart.PriceFor(snapshot.Products.Single(p => p.ProductID == shop.Rice), shop.Box, shop.Ali));   // 220 less 10 %
    }

    [Fact]
    public async Task Only_an_administrator_changes_the_database_and_never_with_unsent_sales()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        using var pc = await ReadyAsync(new DirectBackend(t), Local);
        var s = pc.Session;
        Assert.Equal("إعدادات قاعدة البيانات لمدير النظام فقط.", await s.SaveConnectionAsync(Local with { DeviceCode = "POS2" }, null));
        var cart = s.NewCart()!;
        cart.Add(shop.Sugar);
        pc.Data.Down = true;
        Assert.Equal(SaveStatus.Saved, s.SaveSale(cart).Status);
        s.Logout();
        pc.Data.Down = false;
        Assert.Equal("", await s.LoginAsync("admin", ApiHost.AdminPassword));
        Assert.True(s.CanEditDatabase);
        Assert.Equal("على الجهاز فواتير لم تصل إلى قاعدة البيانات الحالية. أرسلها أولًا، ثم غيّر قاعدة البيانات.",
                     await s.SaveConnectionAsync(Local with { Server = "OTHER" }, null));
        Assert.Equal("", await s.SaveConnectionAsync(Local with { DeviceCode = "pos2" }, null));   // the same database: a new code
        Assert.Equal("POS2", s.Connection!.DeviceCode);
        Assert.Equal(1, (await s.SyncAsync()).Sent);                                             // the cashier's sale, sent for him
        await using (var db = t.CreateDbContext())
        {
            Assert.Equal(shop.Cashier, (await db.SalesInvoices.SingleAsync()).EmployeeID);
        }
        Assert.Equal("", await s.SaveConnectionAsync(Local with { Server = "OTHER" }, null));
        Assert.False(s.LoggedIn);                                                                 // another database: log in again
        Assert.Empty(pc.Store.Users());
    }

    [Fact]
    public void The_connection_is_checked_before_it_is_tried()
    {
        Assert.Equal("", Local.Problem(null));
        Assert.Equal("رمز الجهاز من حرف إلى ثمانية حروف إنجليزية أو أرقام، مثل POS1.", (Local with { DeviceCode = "نقطة" }).Problem(null));
        Assert.Equal("اكتب كلمة مرور مستخدم SQL Server.", (Local with { WindowsLogin = false, User = "pos" }).Problem(""));
        Assert.Equal("الخادم البعيد يحتاج HTTPS. يُقبل http داخل شبكة المحل فقط.",
                     (Remote with { ServerUrl = "http://store.example.com" }).Problem(null));
        Assert.Equal("", (Remote with { ServerUrl = "http://192.168.1.10:5000" }).Problem(null));
        Assert.Equal("https://store.example.com", PosConnection.NormalizeUrl(" store.example.com/ "));
        var sql = (Local with { WindowsLogin = false, User = "pos" }).ConnectionString("p@ss");
        Assert.Contains("Encrypt=True", sql);
        Assert.Contains("Password=p@ss", sql);
        Assert.Contains("Integrated Security=True", Local.ConnectionString(null));
        Assert.Equal("رمز الجهاز هو بادئة فواتير البرنامج. اختر رمزًا آخر.", PosService.DeviceCodeProblem("INV", "INV-"));
    }

    [Fact]
    public async Task The_server_refuses_a_bad_or_used_number_and_another_users_sale()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        var pos = new PosService(t, t.Permissions, new SalesService(t, t.Permissions));
        var cashier = new ActingUser(shop.Cashier, "TEST");
        DeviceSaleRequest Sale(string reference, string number) =>
            new(reference, number, DateTime.Now.AddMinutes(-5), false, 1, "CASH", 1, 0, null, null, [new CartLine(shop.Sugar, 1, 11.5m)]);
        Assert.Equal(PosService.BadNumber, (await pos.PostDeviceSaleAsync(cashier, Sale("a", "INV-000001"))).Message);
        Assert.Equal(PosService.BadNumber, (await pos.PostDeviceSaleAsync(cashier, Sale("a", "POS1000001"))).Message);
        Assert.Equal(PosService.ClientRefTooLong, (await pos.PostDeviceSaleAsync(cashier, Sale(new string('x', 41), "P1-1"))).Message);
        var ok = await pos.PostDeviceSaleAsync(cashier, Sale("a", "p1-000001"));
        Assert.Equal((SaveStatus.Saved, "P1-000001", false), (ok.Status, ok.InvoiceNumber, ok.Duplicate));
        Assert.True((await pos.PostDeviceSaleAsync(cashier, Sale("a", "P1-000001"))).Duplicate);
        Assert.Equal(PosService.ClientRefUsed, (await pos.PostDeviceSaleAsync(TestDb.Admin, Sale("a", "P1-000001"))).Message);
        Assert.Equal("رقم الفاتورة «P1-000001» مستعمل لفاتورة أخرى. اجعل لكل جهاز رمزًا مختلفًا في إعدادات قاعدة البيانات.",
                     (await pos.PostDeviceSaleAsync(cashier, Sale("b", "P1-000001"))).Message);
        // a time in the future is the time it arrives
        var later = await pos.PostDeviceSaleAsync(cashier, Sale("c", "P1-000002") with { TakenAt = DateTime.Now.AddHours(3) });
        await using var db = t.CreateDbContext();
        Assert.True((await db.SalesInvoices.FindAsync(later.InvoiceID))!.InvoiceDate <= DateTime.Now);
    }

    [Fact]
    public async Task A_sale_refused_on_arrival_stops_alone_and_is_sent_again_after_the_fix()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        using var pc = await ReadyAsync(new DirectBackend(t), Local);
        var s = pc.Session;
        pc.Data.Down = true;
        var cart = s.NewCart()!;
        cart.ChangeCustomer(shop.Ali);
        cart.Add(shop.Sugar);
        Assert.Equal(SaveStatus.Saved, s.SaveSale(cart).Status);
        cart = s.NewCart()!;
        cart.Add(shop.Rice);
        Assert.Equal(SaveStatus.Saved, s.SaveSale(cart).Status);
        await using (var db = t.CreateDbContext())
        {
            (await db.Customers.FindAsync(shop.Ali))!.IsActive = false;                         // stopped at the office meanwhile
            await db.SaveChangesAsync();
        }
        pc.Data.Down = false;
        var report = await s.SyncAsync();
        Assert.Equal((1, 1), (report.Sent, report.Rejected));
        var refused = s.Sales().Single(x => x.State == SaleStates.Rejected);
        Assert.Equal("العميل «علي» غير نشط.", refused.Message);
        Assert.Equal("هذا الإجراء لمدير النظام فقط.", s.Delete(refused.Id));
        await using (var db = t.CreateDbContext())
        {
            (await db.Customers.FindAsync(shop.Ali))!.IsActive = true;
            await db.SaveChangesAsync();
        }
        Assert.Equal("", s.Retry(refused.Id, false));
        Assert.Equal((1, 0), ((await s.SyncAsync()).Sent, s.Refused));
    }

    [Fact]
    public async Task The_receipt_is_a_simplified_tax_invoice_with_its_qr_and_opens_the_drawer()
    {
        using var t = new TestDb();
        var shop = await MakeShopAsync(t);
        using var pc = await ReadyAsync(new DirectBackend(t), Local);
        var s = pc.Session;
        var cart = s.NewCart()!;
        cart.Add(shop.Rice, null, 2);
        var (_, _, sale) = s.SaveSale(cart);
        var lines = ReceiptSlip.Build(sale!.Receipt, false, x => x);
        Assert.Equal("متجري", lines[0].Text);
        Assert.Contains(lines, l => l.Text == "فاتورة ضريبية مبسطة");
        Assert.Contains(lines, l => l.Text == "رقم الفاتورة: POS1-000001");
        Assert.Contains(lines, l => l.Text == "الإجمالي شامل الضريبة: 46.00 SAR");
        Assert.Contains(lines, l => l.Text == "لم تُرسل إلى البرنامج بعد");
        Assert.Equal("46.00", ZatcaQr.Decode(sale.Receipt.Qr!)[4]);

        Assert.Equal("اختر الطابعة في الإعدادات.", await s.PrintAsync(sale));
        s.SetPrinter("Receipt printer", 80, true);
        Assert.Equal("", await s.PrintAsync(sale));
        Assert.Equal("", await s.PrintAsync(sale, reprint: true));
        var jobs = pc.Platform.Printed;
        Assert.Equal(EscPos.Initialize.Concat(EscPos.OpenDrawer), jobs[0].Job[..7]);           // a new cash sale opens the drawer
        Assert.Equal(EscPos.Initialize, jobs[1].Job[..2]);
        Assert.Equal(new byte[] { 0x1D, 0x76, 0x30, 0x00, 72, 0 }, jobs[1].Job[2..8]);         // 576 dots = 72 bytes a row
        var qr = ReceiptSlip.Qr(sale.Receipt.Qr!, 5);
        Assert.True(qr.GetLength(0) > 100 && qr.GetLength(0) == qr.GetLength(1));
    }
}
