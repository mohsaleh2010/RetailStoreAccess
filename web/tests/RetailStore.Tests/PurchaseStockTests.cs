using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;

namespace RetailStore.Tests;

/// <summary>modPurchases.PostPurchaseFromCart / PostSupplierPayment and the stock of modWarehouses in C# (plan 6D, docs/80).</summary>
public class PurchaseStockTests
{
    private sealed class Store : IDisposable
    {
        public TestDb Db { get; } = new();
        public PurchaseService Purchases => new(Db, Db.Permissions);
        public StockService Stock => new(Db, Db.Permissions);
        public SalesService Sales => new(Db, Db.Permissions);
        public int Supplier, Rice, Sugar, Branch;

        public Store()
        {
            using var db = Db.CreateDbContext();
            var supplier = new Suppliers { SupplierName = "مورد الأرز", PaymentTermsDays = 15 };
            var rice = new Products { ProductCode = "R1", Barcode = "628100", ProductName = "أرز", SellingPrice = 23, PurchasePrice = 10, AverageCost = 10,
                                      CurrentQuantity = 10 };
            var sugar = new Products { ProductCode = "S1", ProductName = "سكر", SellingPrice = 11.5m, PurchasePrice = 6, VATCategory = "Z" };
            var branch = new Warehouses { WarehouseName = "فرع الشمال", WarehouseType = "BRANCH" };
            db.AddRange(supplier, rice, sugar, branch);
            db.SaveChanges();
            (Supplier, Rice, Sugar, Branch) = (supplier.SupplierID, rice.ProductID, sugar.ProductID, branch.WarehouseID);
            db.Settings.First().VATNumber = "300000000000003";
            // the opening stock of rice in the main warehouse
            db.InventoryTransactions.Add(new InventoryTransactions { ProductID = rice.ProductID, TransactionTypeID = 8, Quantity = 10, UnitCost = 10,
                                                                     QuantityAfter = 10, ReferenceType = "OPENING", WarehouseID = 1,
                                                                     TransactionDate = DateTime.Today.AddDays(-40) });
            db.SaveChanges();
        }

        public void Dispose() => Db.Dispose();
    }

    private static PurchaseRequest Buy(int supplier, string type, decimal? paid, params PurchaseLine[] lines) =>
        new(supplier, null, null, type, 1, true, 0, paid, "", lines);

    [Fact]
    public async Task A_cash_purchase_brings_the_goods_in_at_the_net_cost_and_updates_the_prices()
    {
        using var s = new Store();
        var r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", null,
            new PurchaseLine(s.Rice, 10, 13, 0, null, 25), new PurchaseLine(s.Sugar, 5, 6)));
        Assert.Equal(SaveStatus.Saved, r.Status);
        Assert.Equal(("PUR-000001", 179.5m), (r.InvoiceNumber, r.Total));
        await using var db = s.Db.CreateDbContext();
        var inv = await db.PurchaseInvoices.SingleAsync();
        Assert.Equal((160m, 19.5m, 179.5m, 179.5m, 0m, 1, 1), (inv.TaxableAmount, inv.Tax, inv.TotalAmount, inv.PaidAmount, inv.RemainingAmount,
                                                               inv.CashBoxID!.Value, inv.WarehouseID));
        var rice = (await db.Products.FindAsync(s.Rice))!;
        // a VAT-registered store: the cost is without VAT; (10 x 10 + 10 x 13) / 20 = 11.5
        Assert.Equal((20m, 11.5m, 13m, 25m, s.Supplier), (rice.CurrentQuantity, rice.AverageCost, rice.PurchasePrice, rice.SellingPrice, rice.SupplierID!.Value));
        var moves = await db.InventoryTransactions.Where(t => t.ReferenceType == "PURCHASE").OrderBy(t => t.TransactionID).ToListAsync();
        Assert.Equal([(10m, 13m), (5m, 6m)], moves.Select(m => (m.Quantity, m.UnitCost)));
        Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "PURCHASE" && a.Details == "Total=179.5"));
        Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "PRICE_CHANGE" && a.Details == $"Selling price: {s.Rice}:23->25"));
        var view = (await s.Purchases.PurchaseAsync(inv.PurchaseInvoiceID, false))!;
        Assert.Equal(("مورد الأرز", 2, "المخزن الرئيسي"), (view.SupplierName, view.Lines.Count, view.WarehouseName));
        Assert.Single(await s.Purchases.PurchasesAsync(DateTime.Today, DateTime.Today, "الأرز", false));
    }

    [Fact]
    public async Task A_credit_purchase_owes_the_supplier_and_the_checks_are_those_of_access()
    {
        using var s = new Store();
        var r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CREDIT", 50, new PurchaseLine(s.Rice, 10, 10)) with
        {
            SupplierInvoiceNo = "A-77", InvoiceDate = DateTime.Today.AddDays(-3), ChargeVat = false, WarehouseID = s.Branch,
        });
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using (var db = s.Db.CreateDbContext())
        {
            var inv = await db.PurchaseInvoices.SingleAsync();
            Assert.Equal((DateTime.Today.AddDays(-3), 0m, 50m, 50m, DateTime.Today.AddDays(12)),
                         (inv.InvoiceDate, inv.Tax, inv.PaidAmount, inv.RemainingAmount, inv.DueDate!.Value));
            Assert.Equal(50m, (await db.Suppliers.FindAsync(s.Supplier))!.CurrentBalance);
            Assert.Equal(s.Branch, (await db.InventoryTransactions.SingleAsync(t => t.ReferenceType == "PURCHASE")).WarehouseID);
        }
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CREDIT", 0, new PurchaseLine(s.Rice, 1, 10)) with { SupplierInvoiceNo = "A-77" });
        Assert.Equal("فاتورة المورد رقم A-77 مسجلة مسبقًا لنفس المورد.", r.Message);
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Rice, 1, 10)) with
        {
            InvoiceDate = DateTime.Today.AddDays(1),
        });
        Assert.Equal("تاريخ الفاتورة لا يمكن أن يكون في المستقبل.", r.Message);
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(0, "CASH", null, new PurchaseLine(s.Rice, 1, 10)));
        Assert.Equal("اختر المورد.", r.Message);
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Rice, 1, 10, 0, null, -1)));
        Assert.Equal("سعر البيع الجديد لا يمكن أن يكون سالبًا («أرز»).", r.Message);
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", 1, new PurchaseLine(s.Rice, 1, 10)));
        Assert.Equal("المبلغ المدفوع (1.00) أقل من إجمالي الفاتورة (11.50).", r.Message);
        r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Rice, 1, 10)) with { WarehouseID = 99 });
        Assert.Equal("اختر المخزن (مخزن نشط).", r.Message);
    }

    [Fact]
    public async Task A_carton_brings_pieces_in_and_keeps_its_own_prices()
    {
        using var s = new Store();
        await using (var db = s.Db.CreateDbContext())
        {
            db.ProductUnits.Add(new ProductUnits { ProductID = s.Sugar, UnitID = 3, UnitFactor = 12, Barcode = "C-12", SellingPrice = 130, PurchasePrice = 70 });
            await db.SaveChangesAsync();
        }
        var found = Assert.Single(await s.Purchases.FindProductsAsync(1, "C-12", false));
        Assert.Equal((70m, 130m, "كرتون"), (found.Cost, found.SellingPrice, found.UnitName));
        var r = await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Sugar, 2, 72, 0, found.UnitID, 135)));
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using var check = s.Db.CreateDbContext();
        var sugar = (await check.Products.FindAsync(s.Sugar))!;
        Assert.Equal((24m, 6m, 6m, 11.5m), (sugar.CurrentQuantity, sugar.AverageCost, sugar.PurchasePrice, sugar.SellingPrice));
        var unit = await check.ProductUnits.SingleAsync();
        Assert.Equal((72m, 135m), (unit.PurchasePrice, unit.SellingPrice));
        var line = await check.PurchaseInvoiceDetails.SingleAsync();
        Assert.Equal((12m, found.UnitID), (line.UnitFactor, line.UnitID));
    }

    [Fact]
    public async Task A_new_shelf_price_needs_the_products_permission()
    {
        using var s = new Store();
        int buyer;
        await using (var db = s.Db.CreateDbContext())
        {
            var role = new Roles { RoleCode = "BUY", RoleName = "مشتريات" };
            db.Roles.Add(role);
            await db.SaveChangesAsync();
            db.RolePermissions.Add(new RolePermissions { RoleID = role.RoleID, PermissionKey = "PURCHASES" });
            var e = new Employees { EmployeeName = "مشتريات", Username = "b1", RoleID = role.RoleID };
            db.Employees.Add(e);
            await db.SaveChangesAsync();
            buyer = e.EmployeeID;
        }
        var user = new ActingUser(buyer, "TEST");
        var r = await s.Purchases.PostPurchaseAsync(user, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Rice, 1, 10, 0, null, 30)));
        Assert.Equal("لا تملك صلاحية تعديل أسعار البيع. امسح عمود «سعر البيع الجديد».", r.Message);
        r = await s.Purchases.PostPurchaseAsync(user, Buy(s.Supplier, "CASH", null, new PurchaseLine(s.Rice, 1, 10, 0, null, 23)));
        Assert.Equal(SaveStatus.Saved, r.Status);                     // the same price is no change
        var pay = await s.Purchases.PostSupplierPaymentAsync(user, new SupplierPaymentRequest(s.Supplier, 5, 1, ""));
        Assert.Equal(await s.Db.Permissions.OpenProblemAsync(buyer, "frmSupplierPayment"), pay.Message);
    }

    [Fact]
    public async Task A_payment_lowers_what_is_owed_and_asks_over_it()
    {
        using var s = new Store();
        await s.Purchases.PostPurchaseAsync(TestDb.Admin, Buy(s.Supplier, "CREDIT", 0, new PurchaseLine(s.Rice, 4, 10)));
        var r = await s.Purchases.PostSupplierPaymentAsync(TestDb.Admin, new SupplierPaymentRequest(s.Supplier, 0, 1, ""));
        Assert.Equal("المبلغ يجب أن يكون أكبر من صفر.", r.Message);
        r = await s.Purchases.PostSupplierPaymentAsync(TestDb.Admin, new SupplierPaymentRequest(s.Supplier, 50, 1, ""));
        Assert.Equal((SaveStatus.NeedsConfirmation, "المبلغ (50.00) أكبر من المستحق للمورد (46.00)، وسيصبح المورد مدينًا للمحل.\nهل تريد المتابعة؟"),
                     (r.Status, r.Message));
        r = await s.Purchases.PostSupplierPaymentAsync(TestDb.Admin, new SupplierPaymentRequest(s.Supplier, 40, 1, "دفعة"));
        Assert.Equal((SaveStatus.Saved, "PAY-000001"), (r.Status, r.InvoiceNumber));
        r = await s.Purchases.PostSupplierPaymentAsync(TestDb.Admin, new SupplierPaymentRequest(s.Supplier, 1, 3, "", "USD", 3.75m), true);
        Assert.Equal(3.75m, r.Total);
        await using var db = s.Db.CreateDbContext();
        Assert.Equal(2.25m, (await db.Suppliers.FindAsync(s.Supplier))!.CurrentBalance);
        var first = await db.SupplierPayments.OrderBy(p => p.PaymentID).FirstAsync();
        Assert.Equal((40m, 1), (first.Amount, first.CashBoxID!.Value));
        Assert.Null((await db.SupplierPayments.OrderBy(p => p.PaymentID).LastAsync()).CashBoxID);   // a transfer moves no box
        Assert.Equal(2, (await s.Purchases.SupplierPaymentsAsync(DateTime.Today, DateTime.Today, "الأرز", false)).Count);
    }

    [Fact]
    public async Task The_stock_of_each_warehouse_and_the_item_card()
    {
        using var s = new Store();
        Assert.Equal(10m, (await s.Stock.StockAsync(1, "أرز", false, false)).Single().Quantity);   // only the main warehouse: CurrentQuantity
        var t = await s.Stock.CreateTransferAsync(TestDb.Admin, new TransferRequest(1, s.Branch, "", [new TransferLineInput(s.Rice, 4)]));
        Assert.Equal((SaveStatus.Saved, "STT-000001"), (t.Status, t.Number));
        Assert.Equal(SaveStatus.Saved, (await s.Stock.SendTransferAsync(TestDb.Admin, t.Id)).Status);
        var main = (await s.Stock.StockAsync(1, "R1", false, false)).Single();
        var branch = (await s.Stock.StockAsync(s.Branch, "R1", false, false)).Single();
        var all = (await s.Stock.StockAsync(0, "R1", false, false)).Single();
        Assert.Equal((6m, 4m, 10m, 100m), (main.Quantity, branch.Quantity, all.Quantity, all.Value));
        Assert.Single(await s.Stock.StockAsync(0, "", true, false), r => r.ProductID == s.Sugar);   // at or under the minimum

        var card = await s.Stock.ItemCardAsync(s.Rice, 1, DateTime.Today.AddDays(-1), DateTime.Today, false);
        Assert.Equal(("رصيد أول المدة", 10m, 100m), (card[0].MovementType, card[0].Balance, card[0].NetValue));
        Assert.Equal(("تحويل صادر لمخزن آخر", 4m, 6m, "فرع الشمال"), (card[1].MovementType, card[1].QtyOut, card[1].Balance, card[1].PartyName));
        var branchCard = await s.Stock.ItemCardAsync(s.Rice, s.Branch, DateTime.Today, DateTime.Today, false);
        Assert.Equal((4m, "المخزن الرئيسي"), (branchCard[1].QtyIn, branchCard[1].PartyName));
    }

    [Fact]
    public async Task A_transfer_to_a_warehouse_that_confirms_waits_in_transit()
    {
        using var s = new Store();
        int keeper;
        await using (var db = s.Db.CreateDbContext())
        {
            var role = new Roles { RoleCode = "STK", RoleName = "مخازن" };
            db.Roles.Add(role);
            await db.SaveChangesAsync();
            db.RolePermissions.Add(new RolePermissions { RoleID = role.RoleID, PermissionKey = "STOCK_TRANSFER" });
            var e = new Employees { EmployeeName = "أمين الفرع", Username = "k1", RoleID = role.RoleID };
            db.Employees.Add(e);
            await db.SaveChangesAsync();
            keeper = e.EmployeeID;
            var w = (await db.Warehouses.FindAsync(s.Branch))!;
            (w.ConfirmReceipt, w.KeeperID) = (true, keeper);
            await db.SaveChangesAsync();
        }
        Assert.Equal("المخزن المستلم غير المرسِل.",
                     (await s.Stock.CreateTransferAsync(TestDb.Admin, new TransferRequest(1, 1, "", [new TransferLineInput(s.Rice, 1)]))).Message);
        Assert.Equal("الكمية أكبر من صفر.",
                     (await s.Stock.CreateTransferAsync(TestDb.Admin, new TransferRequest(1, s.Branch, "", [new TransferLineInput(s.Rice, 0)]))).Message);
        var t = await s.Stock.CreateTransferAsync(TestDb.Admin, new TransferRequest(1, s.Branch, "", [new TransferLineInput(s.Rice, 12)]));
        var ask = await s.Stock.SendTransferAsync(TestDb.Admin, t.Id);
        Assert.Equal((SaveStatus.NeedsConfirmation,
                      "الكمية المتوفرة غير كافية:\n- أرز: المطلوب 12 والمتوفر 10\n\nهل تسمح بإتمام العملية (التحويل بين المخازن) رغم ذلك؟ (بصلاحية مدير النظام)"),
                     (ask.Status, ask.Message));
        Assert.Equal(SaveStatus.Saved, (await s.Stock.SaveTransferLinesAsync(TestDb.Admin, t.Id, [new TransferLineInput(s.Rice, 3)])).Status);
        Assert.Equal(SaveStatus.Saved, (await s.Stock.SendTransferAsync(TestDb.Admin, t.Id)).Status);
        var view = (await s.Stock.TransferAsync(t.Id, 1, false))!;
        Assert.Equal(("SENT", true, true, 10m), (view.Status, view.InTransit, view.CanReceive, view.Lines[0].UnitCost));
        Assert.Equal(0m, (await s.Stock.StockAsync(s.Branch, "R1", false, false)).Single().Quantity);   // still on the road

        // the unpost puts it back to draft, then it is sent again and received by the keeper
        Assert.Equal(SaveStatus.Saved, (await s.Stock.UnpostTransferAsync(TestDb.Admin, t.Id)).Status);
        await using (var db = s.Db.CreateDbContext())
        {
            Assert.Equal(10m, (await db.Products.FindAsync(s.Rice))!.CurrentQuantity);
            Assert.False(await db.InventoryTransactions.AnyAsync(x => x.ReferenceType == "TRANSFER"));
        }
        await s.Stock.SendTransferAsync(TestDb.Admin, t.Id);
        Assert.Equal("هذا التحويل ليس بالطريق.", (await s.Stock.ReceiveTransferAsync(TestDb.Admin, 999)).Message);
        Assert.Equal(SaveStatus.Saved, (await s.Stock.ReceiveTransferAsync(new ActingUser(keeper, "TEST"), t.Id)).Status);
        Assert.Equal(3m, (await s.Stock.StockAsync(s.Branch, "R1", false, false)).Single().Quantity);
        Assert.Equal("يُلغى ترحيل التحويل بالطريق فقط. التحويل المستلم يُعاد بتحويل جديد.",
                     (await s.Stock.UnpostTransferAsync(TestDb.Admin, t.Id)).Message);
        Assert.Equal("تُحذف المسودة فقط. ألغِ ترحيل التحويل بالطريق أولًا.", (await s.Stock.DeleteTransferAsync(TestDb.Admin, t.Id)).Message);
        await using var check = s.Db.CreateDbContext();
        var done = await check.StockTransfers.FindAsync(t.Id);
        Assert.Equal(("RECEIVED", keeper), (done!.Status, done.ReceivedByID!.Value));
        Assert.True(await check.AuditLog.AnyAsync(a => a.ActionType == "STOCK_TRANSFER_RECEIVE" && a.RecordID == t.Id.ToString()));
        Assert.True(await check.AuditChanges.AnyAsync(c => c.FieldName == "Status" && c.OldValue == "SENT" && c.NewValue == "RECEIVED"));
    }

    [Fact]
    public async Task A_closed_period_stops_sending_and_a_draft_is_deleted_with_its_lines()
    {
        using var s = new Store();
        var t = await s.Stock.CreateTransferAsync(TestDb.Admin, new TransferRequest(1, s.Branch, "تجربة", [new TransferLineInput(s.Rice, 1)]));
        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.First().ClosedThrough = DateTime.Today;
            await db.SaveChangesAsync();
        }
        string day = DateTime.Today.ToString("yyyy/MM/dd");
        Assert.Equal("التاريخ " + day + " في فترة مقفلة (الإقفال حتى " + day + ").\nلا يُضاف ولا يُعدَّل ولا يُحذف فيها شيء إلا بعد إعادة فتح الفترة.",
                     (await s.Stock.SendTransferAsync(TestDb.Admin, t.Id)).Message);
        Assert.Equal(SaveStatus.Saved, (await s.Stock.DeleteTransferAsync(TestDb.Admin, t.Id)).Status);
        await using var check = s.Db.CreateDbContext();
        Assert.False(await check.StockTransferLines.AnyAsync());
        var log = await check.AuditLog.SingleAsync(a => a.ActionType == "STOCK_TRANSFER_DELETE");
        Assert.Equal("STT-000001", log.RecordLabel);
        Assert.Empty(await s.Stock.TransfersAsync("", 0, false));
    }

    [Fact]
    public async Task Purchases_payments_stock_and_transfers_through_the_api()
    {
        using var host = new ApiHost();
        var admin = await host.AdminAsync();
        int rice, supplier, branch;
        await using (var db = host.Db.CreateDbContext())
        {
            var p = new Products { ProductCode = "R1", Barcode = "628100", ProductName = "أرز", SellingPrice = 23, PurchasePrice = 10 };
            var sp = new Suppliers { SupplierName = "مورد" };
            var w = new Warehouses { WarehouseName = "فرع", WarehouseType = "BRANCH" };
            db.AddRange(p, sp, w);
            await db.SaveChangesAsync();
            (rice, supplier, branch) = (p.ProductID, sp.SupplierID, w.WarehouseID);
        }
        var found = await (await admin.GetAsync("/api/v1/purchases/products?code=628100")).Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(10m, found[0].GetProperty("cost").GetDecimal());
        var made = await admin.PostAsJsonAsync("/api/v1/purchases", new
        {
            supplierID = supplier, paymentType = "CREDIT", chargeVat = true, invoiceDiscount = 0, paid = 0,
            lines = new[] { new { productID = rice, quantity = 10, unitCost = 10 } },
        });
        Assert.Equal(HttpStatusCode.Created, made.StatusCode);
        int id = (await made.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("id").GetInt32();
        var view = await (await admin.GetAsync($"/api/v1/purchases/{id}")).Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(115m, view.GetProperty("remainingAmount").GetDecimal());
        var pay = await admin.PostAsJsonAsync("/api/v1/supplier-payments", new { supplierID = supplier, amount = 200, paymentMethodID = 1 });
        Assert.Equal(HttpStatusCode.Conflict, pay.StatusCode);
        pay = await admin.PostAsJsonAsync("/api/v1/supplier-payments", new { supplierID = supplier, amount = 100, paymentMethodID = 1 });
        Assert.Equal(HttpStatusCode.Created, pay.StatusCode);

        var transfer = await admin.PostAsJsonAsync("/api/v1/stock-transfers", new
        {
            fromWarehouseID = 1, toWarehouseID = branch, lines = new[] { new { productID = rice, quantity = 4 } },
        });
        Assert.Equal(HttpStatusCode.Created, transfer.StatusCode);
        int tid = (await transfer.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("id").GetInt32();
        Assert.Equal(HttpStatusCode.NoContent, (await admin.PostAsync($"/api/v1/stock-transfers/{tid}/send", null)).StatusCode);
        var again = await admin.PostAsync($"/api/v1/stock-transfers/{tid}/send?lang=en", null);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, again.StatusCode);
        Assert.DoesNotContain((await again.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("title").GetString()!, c => c >= '؀' && c <= 'ۿ');
        var stock = await (await admin.GetAsync($"/api/v1/stock?warehouseId={branch}")).Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(4m, stock[0].GetProperty("quantity").GetDecimal());
        var card = await (await admin.GetAsync($"/api/v1/stock/{rice}/card?warehouseId=1")).Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(6m, card[card.GetArrayLength() - 1].GetProperty("balance").GetDecimal());
        Assert.Single((await (await admin.GetAsync("/api/v1/stock-transfers?status=RECEIVED")).Content.ReadFromJsonAsync<JsonElement>()).EnumerateArray());
    }
}
