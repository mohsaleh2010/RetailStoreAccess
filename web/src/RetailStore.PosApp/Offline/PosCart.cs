using RetailStore.Data.Services;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.PosApp.Offline;

/// <summary>A line of a receipt as it is printed.</summary>
public sealed record ReceiptLine(string ProductName, string UnitName, decimal Quantity, decimal UnitPrice, decimal Discount, decimal Net,
                                 decimal VATRate, decimal Tax, decimal Total);

/// <summary>A sale made on the POS as it is kept on the PC and printed: the request sent to the program and what the
/// receipt shows (the same amounts the program will save: the invoice engine of the program computed them).</summary>
public sealed record PosReceipt(DeviceSaleRequest Request, string StoreName, string? StoreVAT, string CustomerName, string? CustomerVAT,
                                string? MethodName, ReceiptLine[] Lines, decimal SubTotal, decimal Discount, decimal Taxable, decimal Tax,
                                decimal Total, decimal Paid, decimal Remaining, decimal Tendered, decimal Change, string? Qr,
                                string Currency, string CashierName, string WarehouseName, string? Footer);

/// <summary>A line of the cart (tmpPOSLines): UnitID null = the smallest unit; Factor = pieces in the unit; ListPrice = the
/// price the POS gave the line.</summary>
public sealed class CartItem(PosProduct product, int? unitId, string unitName, decimal factor, decimal quantity, decimal price)
{
    public PosProduct Product { get; } = product;
    public int? UnitID { get; set; } = unitId;
    public string UnitName { get; set; } = unitName;
    public decimal Factor { get; set; } = factor;
    public decimal Quantity { get; set; } = quantity;
    public decimal UnitPrice { get; set; } = price;
    public decimal LineDiscount { get; set; }
    public decimal ListPrice { get; set; } = price;
    public decimal Gross => MoneyMath.RoundMoney(Quantity * UnitPrice) - LineDiscount;
}

/// <summary>What the cart adds up to for the screen: the invoice engine of the program on the lines (or its problem).</summary>
public sealed record CartTotals(string Problem, decimal SubTotal, decimal Discount, decimal Tax, decimal Total);

/// <summary>What the PC still has to send: the pieces of each product and the credit of each customer, taken off the
/// snapshot so a second sale offline sees the first.</summary>
public sealed record PendingEffects(Dictionary<int, decimal> Pieces, Dictionary<int, decimal> Credit)
{
    public static readonly PendingEffects None = new([], []);
}

/// <summary>The cart of the POS without the database (docs/83): the checks of modSales.PostSaleFromCart in the same order
/// and with the same Arabic messages, against the snapshot of the user (his permissions and limits, the prices, the stock
/// of his warehouse less the reservations, the customers' balances) and what the PC has not sent yet. The program checks
/// the sale again when it arrives.</summary>
public sealed class PosCart(PosSnapshot snapshot)
{
    public PosSnapshot Snapshot { get; } = snapshot;
    public List<CartItem> Items { get; } = [];
    public int CustomerID { get; set; } = snapshot.Settings.DefaultCustomerID;
    public string PaymentType { get; set; } = "CASH";
    public int? PaymentMethodID { get; set; } = SalesService.CashMethodId;
    public decimal InvoiceDiscount { get; set; }
    public decimal? Tendered { get; set; }
    public string? Notes { get; set; }

    public PosCustomer? Customer => Snapshot.Customers.FirstOrDefault(c => c.CustomerID == CustomerID);

    public PosProduct? Product(int productId) => Snapshot.Products.FirstOrDefault(p => p.ProductID == productId);

    // ------------------------------------------------------------------ finding a product, prices

    /// <summary>modProductUnits.FindProductCode: a barcode or code of a product (its smallest unit) or of a unit; else the
    /// products whose name or code contains the text.</summary>
    public List<(PosProduct Product, PosUnit? Unit)> Find(string? text, int take = 20)
    {
        string t = (text ?? "").Trim();
        if (t.Length == 0)
        {
            return [];
        }
        var exact = Snapshot.Products.Where(p => p.Barcode == t || p.ProductCode == t).OrderBy(p => p.ProductID).FirstOrDefault();
        if (exact is not null)
        {
            return [(exact, null)];
        }
        foreach (var p in Snapshot.Products)
        {
            if (p.Units.FirstOrDefault(u => u.Barcode == t) is { } u)
            {
                return [(p, u)];
            }
        }
        return Snapshot.Products
            .Where(p => p.ProductName.Contains(t, StringComparison.OrdinalIgnoreCase) || p.ProductCode.Contains(t, StringComparison.OrdinalIgnoreCase)
                        || (p.Barcode?.Contains(t, StringComparison.Ordinal) ?? false))
            .Take(take).Select(p => (p, (PosUnit?)null)).ToList();
    }

    private static bool IsBaseUnit(PosProduct p, int? unitId) => unitId is null || unitId == 0 || unitId == p.UnitID;

    private static PosUnit? UnitOf(PosProduct p, int? unitId) => IsBaseUnit(p, unitId) ? null : p.Units.FirstOrDefault(u => u.UnitID == unitId);

    /// <summary>modProductUnits.UnitSalePrice for the customer: his price list where it differs, else the price of the unit
    /// (the user's rep price, the unit's own price, or the piece price x the factor).</summary>
    public decimal PriceFor(PosProduct p, int? unitId, int customerId)
    {
        var c = Snapshot.Customers.FirstOrDefault(x => x.CustomerID == customerId);
        string key = p.ProductID + ":" + (IsBaseUnit(p, unitId) ? 0 : unitId);
        if (c is { PriceGroup.Length: > 0 } && Snapshot.PriceGroups.TryGetValue(c.PriceGroup, out var prices)
            && prices.TryGetValue(key, out var listed))
        {
            return listed;
        }
        return UnitOf(p, unitId)?.Price ?? p.Price;
    }

    /// <summary>Add a product (more of the same unit adds to its line): "" or the reason.</summary>
    public string Add(int productId, int? unitId = null, decimal quantity = 1)
    {
        var p = Product(productId);
        if (p is null)
        {
            return "أحد الأصناف في الفاتورة لم يعد موجودًا. احذفه من الفاتورة.";
        }
        var unit = UnitOf(p, unitId);
        if (!IsBaseUnit(p, unitId) && unit is null)
        {
            return "الوحدة المختارة لـ «" + p.ProductName + "» لم تعد من وحداته. اختر الوحدة من جديد.";
        }
        int? u = unit?.UnitID;
        var line = Items.FirstOrDefault(i => i.Product.ProductID == productId && i.UnitID == u);
        if (line is not null)
        {
            line.Quantity += quantity;
            return "";
        }
        Items.Add(new CartItem(p, u, unit?.UnitName ?? p.UnitName, unit?.Factor ?? 1, quantity, PriceFor(p, u, CustomerID)));
        return "";
    }

    /// <summary>Another unit for a line: its price is the price of that unit.</summary>
    public string SetUnit(CartItem item, int? unitId)
    {
        var unit = UnitOf(item.Product, unitId);
        if (!IsBaseUnit(item.Product, unitId) && unit is null)
        {
            return "الوحدة المختارة لـ «" + item.Product.ProductName + "» لم تعد من وحداته. اختر الوحدة من جديد.";
        }
        item.UnitID = unit?.UnitID;
        item.UnitName = unit?.UnitName ?? item.Product.UnitName;
        item.Factor = unit?.Factor ?? 1;
        item.UnitPrice = item.ListPrice = PriceFor(item.Product, item.UnitID, CustomerID);
        return "";
    }

    /// <summary>modPOS.RepriceCart: a new customer reprices the lines still at the price the POS gave them.</summary>
    public void ChangeCustomer(int customerId)
    {
        CustomerID = customerId;
        foreach (var i in Items.Where(i => i.UnitPrice == i.ListPrice))
        {
            i.UnitPrice = i.ListPrice = PriceFor(i.Product, i.UnitID, customerId);
        }
    }

    public void Clear()
    {
        Items.Clear();
        InvoiceDiscount = 0;
        Tendered = null;
        Notes = null;
        CustomerID = Snapshot.Settings.DefaultCustomerID;
        PaymentType = "CASH";
        PaymentMethodID = SalesService.CashMethodId;
    }

    // ------------------------------------------------------------------ the calculation and the checks

    private CalcResult Calc() =>
        InvoiceCalculator.Calc(Items.Select(i => new CalcLineIn(i.Quantity, i.UnitPrice, i.LineDiscount,
                                                                i.Product.VATCategory == "S" ? Snapshot.Settings.VATRate : 0)).ToList(),
                               InvoiceDiscount, Snapshot.Settings.PricesIncludeVAT);

    public CartTotals Totals()
    {
        var c = Calc();
        return new CartTotals(c.Problem, c.SubTotal, c.Discount, c.Tax, c.Total);
    }

    /// <summary>The price is not a change (PostSaleFromCart step 1, SalesService.PriceChangedAsync).</summary>
    private bool PriceChanged(CartItem i)
    {
        var p = i.Product;
        if (Customer is { PriceGroup.Length: > 0 } && i.UnitPrice == PriceFor(p, i.UnitID, CustomerID))
        {
            return false;
        }
        if (i.Factor != 1)
        {
            if (i.UnitPrice == (UnitOf(p, i.UnitID)?.OwnPrice ?? 0))
            {
                return false;
            }
            return i.UnitPrice != MoneyMath.RoundMoney(p.SellingPrice * i.Factor)
                   && i.UnitPrice != MoneyMath.RoundMoney((p.RepPrice ?? p.SellingPrice) * i.Factor);
        }
        if (i.UnitPrice == p.SellingPrice || i.UnitPrice == (p.SizePriceM ?? -1) || i.UnitPrice == (p.SizePriceL ?? -1))
        {
            return false;
        }
        return p.RepPrice is not { } rp || rp != i.UnitPrice;
    }

    /// <summary>The checks of PostSaleFromCart before the sale is kept: Refused, NeedsConfirmation (the stock question for
    /// who has ALLOW_NEGATIVE_STOCK) or Saved, with the message; pending = what the PC has not sent yet.</summary>
    public (SaveStatus Status, string Message, CalcResult? Calc, Settlement? Settle) Check(PendingEffects pending)
    {
        var user = Snapshot.User;
        if (user.SellProblem.Length > 0)
        {
            return (SaveStatus.Refused, user.SellProblem, null, null);
        }
        // ---- 1. cart -> calculation engine
        var calc = Calc();
        if (!calc.Ok)
        {
            return (SaveStatus.Refused, calc.Problem, null, null);
        }
        // ---- 2. permissions: price changes and discounts
        if (Items.Any(PriceChanged) && !user.PriceOverride)
        {
            return (SaveStatus.Refused, "لا تملك صلاحية تعديل سعر البيع.", null, null);
        }
        string msg = DiscountProblem(calc);
        if (msg.Length > 0)
        {
            return (SaveStatus.Refused, msg, null, null);
        }
        // ---- 3. customer and payment
        bool isCredit = PaymentType == "CREDIT";
        decimal tend = Tendered ?? (isCredit ? 0 : calc.Total);
        var settle = InvoiceCalculator.Settle(calc.Total, tend, isCredit);
        if (settle.Problem.Length > 0)
        {
            return (SaveStatus.Refused, settle.Problem, null, null);
        }
        msg = CustomerProblem(isCredit, settle.Remaining, pending);
        if (msg.Length == 0)
        {
            msg = RepCreditProblem(settle.Remaining, pending);
        }
        if (msg.Length > 0)
        {
            return (SaveStatus.Refused, msg, null, null);
        }
        var (status, stockMessage) = StockProblem(calc, pending);
        return status == SaveStatus.Saved ? (SaveStatus.Saved, "", calc, settle) : (status, stockMessage, calc, settle);
    }

    /// <summary>modSales.CheckDiscountAllowed with the limit of the snapshot (the rep's when he has one).</summary>
    private string DiscountProblem(CalcResult calc)
    {
        decimal lineDisc = Items.Sum(i => i.LineDiscount);
        if (calc.Gross + lineDisc <= 0)
        {
            return "";
        }
        double pct = (double)(lineDisc + InvoiceDiscount) / (double)(calc.Gross + lineDisc);
        double maxPct = (double)Snapshot.User.MaxDiscountPercent;
        if (pct <= 0 || pct <= maxPct || Snapshot.User.DiscountOverride)
        {
            return "";
        }
        return "الخصم (" + SalesService.Percent(pct) + ") أكبر من الحد المسموح لك (" + SalesService.Percent(maxPct) + ").";
    }

    /// <summary>modSales.CheckCustomer: his balance with the credit of the sales not sent yet.</summary>
    private string CustomerProblem(bool isCredit, decimal remaining, PendingEffects pending)
    {
        var c = Customer;
        if (c is null)
        {
            return "العميل غير موجود.";
        }
        if (!isCredit && remaining <= 0)
        {
            return "";
        }
        if (!c.AllowCredit)
        {
            return "العميل «" + c.CustomerName + "» غير مسموح له بالبيع الآجل.\n" + "للبيع الآجل اختر عميلًا مسجلًا.";
        }
        decimal balance = c.CurrentBalance + pending.Credit.GetValueOrDefault(c.CustomerID);
        if (c.CreditLimit > 0 && balance + remaining > c.CreditLimit)
        {
            return "تجاوز حد الائتمان: الرصيد " + InvoiceCalculator.Money(balance) + " + المتبقي " +
                   InvoiceCalculator.Money(remaining) + " أكبر من الحد " + InvoiceCalculator.Money(c.CreditLimit) + ".";
        }
        int blockDays = Snapshot.Settings.CreditBlockDays;
        if (blockDays > 0 && c.DaysLate > blockDays)
        {
            return "على العميل «" + c.CustomerName + "» فاتورة متأخرة " + c.DaysLate + " يومًا عن استحقاقها " + "(الحد " + blockDays +
                   " يومًا).\n" + "حصّل المتأخر أولًا، أو بِع نقدًا.";
        }
        return "";
    }

    /// <summary>modRepRoutes.RepCreditProblem: the balances of the rep's customers with the credit not sent yet.</summary>
    private string RepCreditProblem(decimal remaining, PendingEffects pending)
    {
        if (Customer?.RepID is not { } rep || remaining <= 0 || Snapshot.Reps.FirstOrDefault(r => r.SalesRepID == rep) is not { } r)
        {
            return "";
        }
        decimal bal = r.CustomersBalance + Snapshot.Customers.Where(c => c.RepID == rep)
            .Sum(c => pending.Credit.GetValueOrDefault(c.CustomerID));
        return bal + remaining > r.CreditLimit
            ? "تجاوز حد ائتمان المندوب «" + r.RepName + "»: أرصدة عملائه " + InvoiceCalculator.Money(bal) + " + المتبقي " +
              InvoiceCalculator.Money(remaining) + " أكبر من الحد " + InvoiceCalculator.Money(r.CreditLimit) + "."
            : "";
    }

    /// <summary>modSales.CheckStockAvailable: the stock of the snapshot less the pieces not sent yet.</summary>
    private (SaveStatus Status, string Message) StockProblem(CalcResult calc, PendingEffects pending)
    {
        string shortList = "";
        foreach (var g in Items.Select((item, n) => (item, pieces: calc.Lines[n].Quantity * item.Factor)).GroupBy(x => x.item.Product.ProductID))
        {
            var p = g.First().item.Product;
            decimal need = p.TrackStock ? g.Sum(x => x.pieces) : 0;
            decimal available = p.Available - pending.Pieces.GetValueOrDefault(p.ProductID);
            if (need > available)
            {
                shortList += "\n- " + p.ProductName + ": المطلوب " + InvoiceCalculator.Plain(need) + " والمتوفر " + InvoiceCalculator.Plain(available);
                if (p.Reserved > 0)
                {
                    shortList += " (بعد " + InvoiceCalculator.Plain(p.Reserved) + " محجوزة لطلبيات العملاء)";
                }
            }
        }
        if (shortList.Length == 0 || Snapshot.Settings.AllowNegativeStock)
        {
            return (SaveStatus.Saved, "");
        }
        if (Snapshot.User.AllowNegativeStock)
        {
            return (SaveStatus.NeedsConfirmation, "الكمية المتوفرة غير كافية:" + shortList + "\n\n" + SalesService.NegativeStockQuestion + " (" +
                                                  "البيع" + ") رغم ذلك؟ (بصلاحية مدير النظام)");
        }
        return (SaveStatus.Refused, "الكمية المتوفرة غير كافية:" + shortList + "\n" + "البيع" + " بأكثر من المتوفر يحتاج موافقة مدير النظام.");
    }

    /// <summary>The sale to keep and send: the request (the device number and time) and the receipt; confirmStock = the
    /// user answered yes to the stock question. Null with the message when a check refuses it.</summary>
    public (SaveStatus Status, string Message, PosReceipt? Receipt) ToReceipt(string invoiceNumber, DateTime takenAt, bool confirmStock,
                                                                             PendingEffects pending)
    {
        var (status, message, calc, settle) = Check(pending);
        if (status == SaveStatus.Refused || (status == SaveStatus.NeedsConfirmation && !confirmStock) || calc is null || settle is null)
        {
            return (status, message, null);
        }
        var at = DateTime.SpecifyKind(new DateTime(takenAt.Ticks - takenAt.Ticks % TimeSpan.TicksPerSecond), DateTimeKind.Unspecified);
        bool isCredit = PaymentType == "CREDIT";
        var request = new DeviceSaleRequest(Guid.NewGuid().ToString("N"), invoiceNumber, at, status == SaveStatus.NeedsConfirmation || confirmStock,
                                            CustomerID, PaymentType, PaymentMethodID, InvoiceDiscount, Tendered ?? (isCredit ? 0 : calc.Total),
                                            string.IsNullOrWhiteSpace(Notes) ? null : Notes.Trim(),
                                            Items.Select(i => new CartLine(i.Product.ProductID, i.Quantity, i.UnitPrice, i.LineDiscount,
                                                                           i.UnitID)).ToArray());
        var s = Snapshot.Settings;
        var c = Customer!;
        string? qr = !string.IsNullOrEmpty(s.VATNumber) && s.CountryCode == "SA"
            ? ZatcaQr.Build(s.StoreName, s.VATNumber, at, calc.Total, calc.Tax) : null;
        var lines = Items.Select((i, n) => new ReceiptLine(i.Product.ProductName, i.UnitName, calc.Lines[n].Quantity, calc.Lines[n].UnitPrice,
                                                           calc.Lines[n].Discount, calc.Lines[n].Net, calc.Lines[n].Rate, calc.Lines[n].Tax,
                                                           calc.Lines[n].Total)).ToArray();
        string? method = Snapshot.Methods.FirstOrDefault(m => m.PaymentMethodID == PaymentMethodID)?.MethodName;
        return (SaveStatus.Saved, "", new PosReceipt(request, s.StoreName, s.VATNumber, c.CustomerName, c.VATNumber, method, lines, calc.SubTotal,
                                                     calc.Discount, calc.Taxable, calc.Tax, calc.Total, settle.Paid, settle.Remaining,
                                                     request.Tendered ?? 0, settle.Change, qr, s.Currency, Snapshot.User.EmployeeName,
                                                     Snapshot.User.WarehouseName, s.ReceiptFooter));
    }
}
