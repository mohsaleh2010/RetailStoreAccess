using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.Domain.Sales;

namespace RetailStore.Tests;

/// <summary>The invoice engine against the cases of tools/pricing.py, the ZATCA QR against tools/zatca_reference.py.</summary>
public class InvoiceCalculatorTests
{
    [Fact]
    public void Every_pricing_case_of_python_gives_the_same_lines_and_totals()
    {
        foreach (var c in Vectors.Pricing)
        {
            var r = InvoiceCalculator.Calc(c.Lines.Select(l => new CalcLineIn(l.Qty, l.Price, l.Discount, l.Rate)).ToList(), c.InvoiceDiscount,
                                           c.IncludeVat);
            Assert.True(r.Ok, c.Label);
            for (int i = 0; i < c.Out.Length; i++)
            {
                Assert.Equal((c.Out[i].Unit, c.Out[i].Discount, c.Out[i].Net, c.Out[i].Tax, c.Out[i].Total),
                             (r.Lines[i].UnitPrice, r.Lines[i].Discount, r.Lines[i].Net, r.Lines[i].Tax, r.Lines[i].Total));
            }
            Assert.Equal(c.Totals, (r.SubTotal, r.Discount, r.Taxable, r.Tax, r.Total));
        }
    }

    [Fact]
    public void Returns_and_averages_are_those_of_python()
    {
        foreach (var v in Vectors.Returns)
        {
            var r = InvoiceCalculator.ReturnAmounts(v.SoldQty, v.SoldTotal, v.SoldTax, v.PrevQty, v.PrevTotal, v.PrevTax, v.Qty);
            Assert.Equal((v.Net, v.Tax, v.Total), (r.Net, r.Tax, r.Total));
        }
        foreach (var v in Vectors.Averages)
        {
            Assert.Equal(v.Average, InvoiceCalculator.WeightedAverage(v.CurQty, v.CurAvg, v.Qty, v.UnitCost));
        }
    }

    [Fact]
    public void The_problems_are_the_messages_of_modSales()
    {
        Assert.Equal("لا توجد أصناف في الفاتورة.", InvoiceCalculator.Calc([], 0, true).Problem);
        Assert.Equal("الكمية يجب أن تكون أكبر من صفر (السطر 2).",
                     InvoiceCalculator.Calc([new(1, 5, 0, 0.15m), new(0, 5, 0, 0.15m)], 0, true).Problem);
        Assert.Equal("خصم السطر 1 أكبر من قيمته.", InvoiceCalculator.Calc([new(1, 5, 6, 0.15m)], 0, true).Problem);
        Assert.Equal("خصم الفاتورة أكبر من قيمتها.", InvoiceCalculator.Calc([new(1, 5, 0, 0.15m)], 6, true).Problem);
        Assert.Equal("المبلغ المدفوع (10.00) أقل من إجمالي الفاتورة (1,150.00).", InvoiceCalculator.Settle(1150, 10, false).Problem);
        Assert.Equal(new Settlement("", 1150, 0, 50), InvoiceCalculator.Settle(1150, 1200, false));
        Assert.Equal(new Settlement("", 100, 1050, 0), InvoiceCalculator.Settle(1150, 100, true));
    }

    [Fact]
    public void The_qr_is_that_of_zatca_reference()
    {
        foreach (var v in Vectors.ZatcaQr)
        {
            Assert.Equal(v.Qr, ZatcaQr.Build(v.Seller, v.Vat, v.When, v.Total, v.Tax));
            Assert.Equal(v.Seller, ZatcaQr.Decode(v.Qr)[1]);
        }
    }

    [Fact]
    public void Aging_pays_the_linked_invoice_then_the_oldest_due()
    {
        var day = new DateTime(2026, 10, 1);
        var debits = new[]
        {
            new AgingDebit(5, "INVOICE", 1, "INV-1", day, day.AddDays(10), 100),
            new AgingDebit(5, "INVOICE", 2, "INV-2", day.AddDays(1), day.AddDays(5), 50),
        };
        var credits = new[] { new AgingCredit(5, "PAYMENT", 7, day.AddDays(2), 120, 0) };
        var links = new[] { new AgingLink(7, 1, 100) };
        var items = CustomerAging.Compute(debits, credits, links, day.AddDays(20));
        var open = Assert.Single(items);
        Assert.Equal((2, 30m), (open.DocID, open.OpenAmount));
        Assert.Equal(15, CustomerAging.MaxDaysLate(items, day.AddDays(20)));
        Assert.Equal(0, CustomerAging.MaxDaysLate(items, day.AddDays(3)));
    }
}

/// <summary>modSales.PostSaleFromCart and PostCustomerPayment in C# (plan 6C, docs/79) on a test database.</summary>
public class SalesServiceTests
{
    private sealed class Store : IDisposable
    {
        public TestDb Db { get; } = new();
        public SalesService Sales => new(Db, Db.Permissions);
        public int Cash, Ali, Rice, Sugar, Delivery;

        public Store()
        {
            using var db = Db.CreateDbContext();
            var ali = new Customers { CustomerName = "علي", CreditLimit = 1000, PaymentTermsDays = 30 };
            var rice = new Products { ProductCode = "R1", Barcode = "628100", ProductName = "أرز", SellingPrice = 23, AverageCost = 12, CurrentQuantity = 10 };
            var sugar = new Products { ProductCode = "S1", ProductName = "سكر", SellingPrice = 11.5m, AverageCost = 6, CurrentQuantity = 100,
                                       VATCategory = "Z" };
            var delivery = new Products { ProductCode = "D1", ProductName = "توصيل", SellingPrice = 5, TrackStock = false };
            db.AddRange(ali, rice, sugar, delivery);
            db.SaveChanges();
            (Cash, Ali, Rice, Sugar, Delivery) = (1, ali.CustomerID, rice.ProductID, sugar.ProductID, delivery.ProductID);
            db.Settings.First().VATNumber = "300000000000003";
            db.SaveChanges();
        }

        public void Dispose() => Db.Dispose();
    }

    [Fact]
    public async Task A_cash_sale_writes_the_invoice_lines_stock_moves_qr_and_audit()
    {
        using var s = new Store();
        Assert.Equal(1, s.Cash);
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Cash, "CASH", 1, 0, 100, "", [
            new CartLine(s.Rice, 2, 23), new CartLine(s.Sugar, 1, 11.5m), new CartLine(s.Delivery, 1, 5)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
        Assert.Equal(("INV-000001", 62.5m, 37.5m), (r.InvoiceNumber, r.Total, r.Change));
        await using var db = s.Db.CreateDbContext();
        var inv = await db.SalesInvoices.SingleAsync();
        Assert.Equal((62.5m, 6.65m, 55.85m, 62.5m, 0m, 100m), (inv.TotalAmount, inv.Tax, inv.TaxableAmount, inv.PaidAmount, inv.RemainingAmount,
                                                                 inv.AmountTendered));
        Assert.Equal(("SIMPLIFIED", "388", 1, "NOT_SENT", 1, 1), (inv.InvoiceSubType, inv.InvoiceTypeCode, inv.ICV, inv.ZatcaStatus, inv.CashBoxID,
                                                                  inv.WarehouseID));
        Assert.Equal("62.50", ZatcaQr.Decode(inv.QRCodeData!)[4]);
        var lines = await db.SalesInvoiceDetails.OrderBy(d => d.LineNumber).ToListAsync();
        Assert.Equal([0.15m, 0m, 0.15m], lines.Select(l => l.VATRate));
        Assert.Equal(12m, lines[0].UnitCost);
        Assert.Equal(8m, (await db.Products.FindAsync(s.Rice))!.CurrentQuantity);
        var moves = await db.InventoryTransactions.OrderBy(t => t.TransactionID).ToListAsync();
        Assert.Equal(2, moves.Count);                                  // the delivery tracks no stock
        Assert.Equal((-2m, 8m, "SALE", inv.SalesInvoiceID, 12m), (moves[0].Quantity, moves[0].QuantityAfter, moves[0].ReferenceType,
                                                                    moves[0].ReferenceID!.Value, moves[0].UnitCost));
        Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "SALE" && a.RecordID == "INV-000001" && a.Details == "Total=62.5"));
    }

    [Fact]
    public async Task A_credit_sale_needs_a_customer_allowed_and_within_his_limit()
    {
        using var s = new Store();
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Cash, "CREDIT", null, 0, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal("العميل «عميل نقدي» غير مسموح له بالبيع الآجل.\nللبيع الآجل اختر عميلًا مسجلًا.", r.Message);
        r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CREDIT", null, 0, 3, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using (var db = s.Db.CreateDbContext())
        {
            var inv = await db.SalesInvoices.SingleAsync();
            Assert.Equal((3m, 20m, DateTime.Today.AddDays(30)), (inv.PaidAmount, inv.RemainingAmount, inv.DueDate!.Value));
            Assert.Equal(20m, (await db.Customers.FindAsync(s.Ali))!.CurrentBalance);
        }
        r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CREDIT", null, 0, 0, "", [new CartLine(s.Sugar, 90, 11.5m)]));
        Assert.Equal("تجاوز حد الائتمان: الرصيد 20.00 + المتبقي 1,035.00 أكبر من الحد 1,000.00.", r.Message);
    }

    [Fact]
    public async Task A_late_invoice_stops_credit_after_the_block_days()
    {
        using var s = new Store();
        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.First().CreditBlockDays = 10;
            db.SalesInvoices.Add(new SalesInvoices { InvoiceNumber = "OLD-1", InvoiceDate = DateTime.Today.AddDays(-60), CustomerID = s.Ali,
                                                     EmployeeID = 1, TotalAmount = 50, RemainingAmount = 50,
                                                     DueDate = DateTime.Today.AddDays(-30) });
            await db.SaveChangesAsync();
        }
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CREDIT", null, 0, 0, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal("على العميل «علي» فاتورة متأخرة 30 يومًا عن استحقاقها (الحد 10 يومًا).\nحصّل المتأخر أولًا، أو بِع نقدًا.", r.Message);
        r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
    }

    [Fact]
    public async Task Price_changes_and_discounts_need_their_permissions()
    {
        using var s = new Store();
        int cashier;
        await using (var db = s.Db.CreateDbContext())
        {
            var e = new Employees { EmployeeName = "كاشير", Username = "c1", RoleID = 3, MaxDiscountPercent = 0.05m };
            db.Employees.Add(e);
            await db.SaveChangesAsync();
            cashier = e.EmployeeID;
        }
        var user = new ActingUser(cashier, "TEST");
        var r = await s.Sales.PostSaleAsync(user, new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 20)]));
        Assert.Equal("لا تملك صلاحية تعديل سعر البيع.", r.Message);
        r = await s.Sales.PostSaleAsync(user, new SaleRequest(s.Cash, "CASH", 1, 3, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal("الخصم (13.0%) أكبر من الحد المسموح لك (5.0%).", r.Message);
        r = await s.Sales.PostSaleAsync(user, new SaleRequest(s.Cash, "CASH", 1, 1, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
        r = await s.Sales.PostSaleAsync(user, new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 20)]));
        Assert.Equal("لا تملك صلاحية تعديل سعر البيع.", r.Message);
        r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 20)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
    }

    [Fact]
    public async Task More_than_the_stock_asks_the_administrator_and_refuses_the_others()
    {
        using var s = new Store();
        var sale = new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 12, 23)]);
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, sale);
        Assert.Equal(SaveStatus.NeedsConfirmation, r.Status);
        Assert.Equal("الكمية المتوفرة غير كافية:\n- أرز: المطلوب 12 والمتوفر 10\n\nهل تسمح بإتمام العملية (البيع) رغم ذلك؟ (بصلاحية مدير النظام)",
                     r.Message);
        r = await s.Sales.PostSaleAsync(TestDb.Admin, sale, confirmStock: true);
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using var db = s.Db.CreateDbContext();
        Assert.Equal(-2m, (await db.Products.FindAsync(s.Rice))!.CurrentQuantity);
        Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "NEGATIVE_STOCK"));
        var e = new Employees { EmployeeName = "كاشير", Username = "c2", RoleID = 3, MaxDiscountPercent = 0 };
        db.Employees.Add(e);
        await db.SaveChangesAsync();
        r = await s.Sales.PostSaleAsync(new ActingUser(e.EmployeeID, "TEST"), new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal("الكمية المتوفرة غير كافية:\n- أرز: المطلوب 1 والمتوفر -2\nالبيع بأكثر من المتوفر يحتاج موافقة مدير النظام.", r.Message);
    }

    [Fact]
    public async Task A_carton_sells_pieces_and_its_own_barcode_is_found()
    {
        using var s = new Store();
        await using (var db = s.Db.CreateDbContext())
        {
            db.ProductUnits.Add(new ProductUnits { ProductID = s.Sugar, UnitID = 3, UnitFactor = 12, Barcode = "C-12", SellingPrice = 130 });
            await db.SaveChangesAsync();
        }
        var found = Assert.Single(await s.Sales.FindProductsAsync(1, "C-12", s.Cash, false));
        Assert.Equal((s.Sugar, 130m, "كرتون"), (found.ProductID, found.Price, found.UnitName));
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Sugar, 2, 130, 0, found.UnitID)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
        await using var check = s.Db.CreateDbContext();
        var line = await check.SalesInvoiceDetails.SingleAsync();
        Assert.Equal((12m, 72m, found.UnitID), (line.UnitFactor, line.UnitCost, line.UnitID));
        Assert.Equal(76m, (await check.Products.FindAsync(s.Sugar))!.CurrentQuantity);
        Assert.Equal(s.Rice, Assert.Single(await s.Sales.FindProductsAsync(1, "628100", s.Cash, false)).ProductID);
    }

    [Fact]
    public async Task E_invoicing_on_makes_the_invoice_pending_and_a_vat_customer_standard()
    {
        using var s = new Store();
        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.First().EInvoiceEnabled = true;
            (await db.Customers.FindAsync(s.Ali))!.VATNumber = "311111111111113";
            await db.SaveChangesAsync();
        }
        var r = await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CASH", 2, 0, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal(SaveStatus.Saved, r.Status);
        var v = (await s.Sales.InvoiceAsync(r.InvoiceID, false))!;
        Assert.Equal(("STANDARD", "PENDING", "مدى / بطاقة"), (v.InvoiceSubType, v.ZatcaStatus, v.MethodName));
        await using var check = s.Db.CreateDbContext();
        Assert.Null((await check.SalesInvoices.SingleAsync()).CashBoxID);   // a card moves no box
    }

    [Fact]
    public async Task A_receipt_lowers_the_balance_and_asks_over_it()
    {
        using var s = new Store();
        await s.Sales.PostSaleAsync(TestDb.Admin, new SaleRequest(s.Ali, "CREDIT", null, 0, 0, "", [new CartLine(s.Rice, 2, 23)]));
        var r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Cash, 10, 1, ""));
        Assert.Equal("لا تُسجَّل دفعات على العميل النقدي.", r.Message);
        r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Ali, 0, 1, ""));
        Assert.Equal("المبلغ يجب أن يكون أكبر من صفر.", r.Message);
        r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Ali, 50, 1, ""));
        Assert.Equal((SaveStatus.NeedsConfirmation, "المبلغ (50.00) أكبر من الرصيد المستحق (46.00)، وسيصبح للعميل رصيد دائن.\nهل تريد المتابعة؟"),
                     (r.Status, r.Message));
        r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Ali, 40, 1, "دفعة"));
        Assert.Equal((SaveStatus.Saved, "RCV-000001"), (r.Status, r.InvoiceNumber));
        r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Ali, 1, 1, "", "USD", 3.75m));
        Assert.Equal(3.75m, r.Total);
        r = await s.Sales.PostPaymentAsync(TestDb.Admin, new PaymentRequest(s.Ali, 1, 1, "", "XYZ", 2));
        Assert.Equal("العملة XYZ غير موجودة أو غير نشطة.", r.Message);
        await using var db = s.Db.CreateDbContext();
        Assert.Equal(2.25m, (await db.Customers.FindAsync(s.Ali))!.CurrentBalance);
        var first = await db.CustomerPayments.OrderBy(p => p.PaymentID).FirstAsync();
        Assert.Equal((40m, 1, "SAR"), (first.Amount, first.CashBoxID!.Value, first.CurrencyCode));
        Assert.Equal(2, (await s.Sales.PaymentsAsync(DateTime.Today, DateTime.Today, "علي", false)).Count);
    }

    [Fact]
    public async Task A_user_without_the_screen_cannot_sell_or_receive()
    {
        using var s = new Store();
        await using var db = s.Db.CreateDbContext();
        var role = new Roles { RoleCode = "VIEW", RoleName = "عرض" };
        db.Roles.Add(role);
        await db.SaveChangesAsync();
        var e = new Employees { EmployeeName = "زائر", Username = "v1", RoleID = role.RoleID };
        db.Employees.Add(e);
        await db.SaveChangesAsync();
        var user = new ActingUser(e.EmployeeID, "TEST");
        var r = await s.Sales.PostSaleAsync(user, new SaleRequest(s.Cash, "CASH", 1, 0, null, "", [new CartLine(s.Rice, 1, 23)]));
        Assert.Equal(SaveStatus.Refused, r.Status);
        Assert.Equal(await s.Db.Permissions.OpenProblemAsync(e.EmployeeID, "frmPOS"), r.Message);
        Assert.False(await db.SalesInvoices.AnyAsync());
    }
}
