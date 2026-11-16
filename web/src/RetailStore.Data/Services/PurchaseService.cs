using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

/// <summary>A line of a purchase (tmpPurchaseLines): UnitCost typed WITHOUT VAT in the currency of the invoice; UnitID
/// null = the smallest unit; NewSellingPrice = a new shelf price for the unit of the line (null = unchanged).</summary>
public sealed record PurchaseLine(int ProductID, decimal Quantity, decimal UnitCost, decimal LineDiscount = 0, int? UnitID = null,
                                  decimal? NewSellingPrice = null);

/// <summary>A purchase invoice (modPurchases.PostPurchaseFromCart): PaymentType CASH or CREDIT; Paid null = the total of a
/// cash invoice, nothing for a credit one; InvoiceDate null = now; WarehouseID 0 = the user's warehouse; the amounts in
/// CurrencyCode (empty = the program currency) at ExchangeRate.</summary>
public sealed record PurchaseRequest(int SupplierID, string? SupplierInvoiceNo, DateTime? InvoiceDate, string PaymentType, int? PaymentMethodID,
                                     bool ChargeVat, decimal InvoiceDiscount, decimal? Paid, string? Notes, IReadOnlyList<PurchaseLine> Lines,
                                     string? CurrencyCode = null, decimal ExchangeRate = 1, int WarehouseID = 0);

/// <summary>A payment voucher to a supplier (modPurchases.PostSupplierPayment).</summary>
public sealed record SupplierPaymentRequest(int SupplierID, decimal Amount, int PaymentMethodID, string? Notes, string? CurrencyCode = null,
                                            decimal ExchangeRate = 1);

/// <summary>A product found for a purchase line, with the unit of the code and its last purchase price (without VAT).</summary>
public sealed record PurchaseProduct(int ProductID, string ProductCode, string? Barcode, string ProductName, int? UnitID, string UnitName,
                                     decimal Cost, decimal SellingPrice, decimal Available, string VATCategory);

public sealed record PurchaseLineView(short LineNumber, int ProductID, string ProductCode, string ProductName, string UnitName,
                                      decimal Quantity, decimal UnitCost, decimal Discount, decimal NetAmount, decimal VATRate,
                                      decimal Tax, decimal LineTotal);

public sealed record PurchaseView(int PurchaseInvoiceID, string InvoiceNumber, string? SupplierInvoiceNo, DateTime InvoiceDate, int SupplierID,
                                  string SupplierName, string? SupplierVAT, string PaymentType, string? MethodName, decimal SubTotal,
                                  decimal Discount, decimal TaxableAmount, decimal Tax, decimal TotalAmount, decimal PaidAmount,
                                  decimal RemainingAmount, DateTime? DueDate, string? CurrencyCode, decimal ExchangeRate,
                                  decimal ForeignAmount, string? Notes, string EmployeeName, string? WarehouseName, List<PurchaseLineView> Lines);

public sealed record PurchaseRow(int PurchaseInvoiceID, string InvoiceNumber, string? SupplierInvoiceNo, DateTime InvoiceDate, string SupplierName,
                                 string PaymentType, decimal TotalAmount, decimal PaidAmount, decimal RemainingAmount, string? WarehouseName,
                                 string EmployeeName);

public sealed record SupplierPaymentRow(int PaymentID, string PaymentNumber, DateTime PaymentDate, int SupplierID, string SupplierName,
                                        decimal Amount, string? CurrencyCode, decimal ForeignAmount, string? MethodName, string? Notes,
                                        string EmployeeName);

/// <summary>Purchases and supplier payments on the web and the API (plan 6D, docs/80): the C# of
/// modPurchases.PostPurchaseFromCart and PostSupplierPayment with the same checks in the same order, the same Arabic
/// messages and the same rows (the invoice and its lines, the stock moves at the cost of one piece with the new average,
/// the last purchase price and the new shelf prices, the supplier's balance, the audit trail), so Access rebuilds the
/// same journal entry (SyncJournal). A purchase from a receipt note (StockNoteID) stays in Access.</summary>
public sealed class PurchaseService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const int TtPurchase = 1;

    private SalesService Sales => new(factory, permissions);

    /// <summary>modPurchases.CheckSupplier.</summary>
    private static async Task<string> CheckSupplierAsync(RetailDbContext db, int supplierId)
    {
        var s = await db.Suppliers.AsNoTracking().FirstOrDefaultAsync(x => x.SupplierID == supplierId);
        if (s is null)
        {
            return "اختر المورد.";
        }
        return s.IsActive ? "" : "المورد «" + s.SupplierName + "» غير نشط.";
    }

    /// <summary>modPurchases.PurchaseDocDate: today -> the current time; an earlier day (an invoice entered late) -> that day.</summary>
    public static (string Problem, DateTime Date) PurchaseDocDate(DateTime? typed, DateTime now)
    {
        if (typed is not { } t)
        {
            return ("", now);
        }
        if (t.Date > now.Date)
        {
            return ("تاريخ الفاتورة لا يمكن أن يكون في المستقبل.", now);
        }
        return ("", t.Date == now.Date ? now : t.Date);
    }

    /// <summary>modSales.PurchaseUnitCost: the cost of one piece after discounts; a VAT-registered store reclaims the
    /// input VAT (its cost excludes it), an unregistered store bears it.</summary>
    public static decimal PurchaseUnitCost(decimal net, decimal total, decimal qty, bool vatRegistered) =>
        qty <= 0 ? 0 : MoneyMath.RoundMoney((vatRegistered ? net : total) / qty, 4);

    /// <summary>modWarehouses.WarehouseProblem.</summary>
    internal static async Task<string> WarehouseProblemAsync(RetailDbContext db, int warehouseId) =>
        await db.Warehouses.AnyAsync(w => w.IsActive && w.WarehouseID == warehouseId) ? "" : "اختر المخزن (مخزن نشط).";

    /// <summary>The price shown on the screen for a unit (the column «سعر البيع»): its own price, else the piece price x
    /// its factor.</summary>
    private static async Task<decimal> ShownSellingPriceAsync(RetailDbContext db, Products p, int? unitId, decimal factor)
    {
        if (factor == 1)
        {
            return p.SellingPrice;
        }
        decimal own = await SalesService.OwnUnitPriceAsync(db, p, unitId);
        return own > 0 ? own : MoneyMath.RoundMoney(p.SellingPrice * factor);
    }

    // ------------------------------------------------------------------ the purchase invoice

    /// <summary>modPurchases.PostPurchaseFromCart (no receipt note).</summary>
    public async Task<SaleResult> PostPurchaseAsync(ActingUser user, PurchaseRequest buy)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, "frmPurchaseInvoice");
        if (open.Length > 0)
        {
            return SaleResult.Refused(open);
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        string msg = await CheckSupplierAsync(db, buy.SupplierID);
        string baseCurrency = string.IsNullOrWhiteSpace(settings.CurrencyCode) ? "SAR" : settings.CurrencyCode;
        string code = string.IsNullOrWhiteSpace(buy.CurrencyCode) ? baseCurrency : buy.CurrencyCode.Trim();
        decimal fx = string.Equals(code, baseCurrency, StringComparison.OrdinalIgnoreCase) ? 1 : buy.ExchangeRate;
        if (msg.Length == 0)
        {
            msg = await SalesService.CurrencyProblemAsync(db, code, fx, baseCurrency);
        }
        var now = DateTime.Now;
        DateTime docDate = now;
        if (msg.Length == 0)
        {
            (msg, docDate) = PurchaseDocDate(buy.InvoiceDate, now);
        }
        string supplierNo = (buy.SupplierInvoiceNo ?? "").Trim();
        if (msg.Length == 0 && supplierNo.Length > 0
            && await db.PurchaseInvoices.AnyAsync(h => h.SupplierID == buy.SupplierID && h.SupplierInvoiceNo == supplierNo))
        {
            msg = "فاتورة المورد رقم " + supplierNo + " مسجلة مسبقًا لنفس المورد.";
        }
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }
        decimal vatRate = settings.VATRate;
        bool reclaim = !string.IsNullOrEmpty(settings.VATNumber);

        // ---- 1. cart -> calculation engine (purchase prices are typed WITHOUT VAT)
        var ids = buy.Lines.Select(l => l.ProductID).Distinct().ToList();
        var products = await db.Products.AsNoTracking().Where(p => ids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID);
        var calcLines = new List<CalcLineIn>();
        var factors = new List<decimal>();
        var changes = new List<string>();
        foreach (var l in buy.Lines)
        {
            if (!products.TryGetValue(l.ProductID, out var p))
            {
                return SaleResult.Refused("أحد الأصناف في الفاتورة لم يعد موجودًا. احذفه من الفاتورة.");
            }
            if (!p.IsActive)
            {
                return SaleResult.Refused("المنتج «" + p.ProductName + "» غير نشط. فعّله من شاشة المنتجات أولًا.");
            }
            if (l.NewSellingPrice is < 0)
            {
                return SaleResult.Refused("سعر البيع الجديد لا يمكن أن يكون سالبًا («" + p.ProductName + "»).");
            }
            decimal factor = await SalesService.UnitFactorOfAsync(db, p, l.UnitID);
            if (factor <= 0)
            {
                return SaleResult.Refused("الوحدة المختارة لـ «" + p.ProductName + "» لم تعد من وحداته. اختر الوحدة من جديد.");
            }
            if (l.NewSellingPrice is { } np && np != await ShownSellingPriceAsync(db, p, l.UnitID, factor))
            {
                changes.Add(" " + p.ProductID + ":" + InvoiceCalculator.Plain(await ShownSellingPriceAsync(db, p, l.UnitID, factor)) + "->" +
                            InvoiceCalculator.Plain(np));
            }
            decimal rate = buy.ChargeVat && p.VATCategory == "S" ? vatRate : 0;
            calcLines.Add(new CalcLineIn(l.Quantity, MoneyMath.Currency(l.UnitCost * fx), MoneyMath.RoundMoney(l.LineDiscount * fx), rate));
            factors.Add(factor);
        }
        var calc = InvoiceCalculator.Calc(calcLines, MoneyMath.RoundMoney(buy.InvoiceDiscount * fx), false);
        msg = calc.Problem;
        if (msg.Length == 0 && changes.Count > 0 && !await permissions.HasPermissionAsync(user.EmployeeID, "PRODUCTS"))
        {
            msg = "لا تملك صلاحية تعديل أسعار البيع. امسح عمود «سعر البيع الجديد».";
        }
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }
        int warehouse = buy.WarehouseID == 0 ? await RepService.UserWarehouseAsync(db, user.EmployeeID) : buy.WarehouseID;
        msg = await WarehouseProblemAsync(db, warehouse);
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }

        // ---- 2. payment: cash = paid in full; credit = whatever was paid now (default 0)
        bool isCredit = buy.PaymentType == "CREDIT";
        decimal tend = buy.Paid is { } paid ? MoneyMath.RoundMoney(paid * fx) : (isCredit ? 0 : calc.Total);
        var settle = InvoiceCalculator.Settle(calc.Total, tend, isCredit);
        if (settle.Problem.Length > 0)
        {
            return SaleResult.Refused(settle.Problem);
        }

        // ---- 3. post everything in one transaction
        await using var tx = await db.Database.BeginTransactionAsync();
        int? method = buy.PaymentMethodID;
        var supplier = await db.Suppliers.AsNoTracking().FirstAsync(s => s.SupplierID == buy.SupplierID);
        var inv = new PurchaseInvoices
        {
            InvoiceNumber = await NumberSequences.NextNumberAsync(db, "PURCHASE_INVOICE"),
            SupplierInvoiceNo = supplierNo.Length > 0 ? SalesService.Clip(supplierNo, 30) : null,
            InvoiceDate = docDate, SupplierID = buy.SupplierID, EmployeeID = user.EmployeeID, PaymentType = buy.PaymentType,
            PaymentMethodID = method, SubTotal = calc.SubTotal, Discount = calc.Discount, TaxableAmount = calc.Taxable, Tax = calc.Tax,
            TotalAmount = calc.Total, PaidAmount = settle.Paid, RemainingAmount = settle.Remaining, CurrencyCode = code, ExchangeRate = fx,
            ForeignAmount = fx <= 0 ? calc.Total : MoneyMath.RoundMoney(calc.Total / fx),
            DueDate = settle.Remaining > 0 ? docDate.Date.AddDays(supplier.PaymentTermsDays ?? 0) : null,
            CashBoxID = await Sales.CashBoxForAsync(db, user.EmployeeID, method ?? SalesService.CashMethodId, settle.Paid),
            BankID = await SalesService.BankForAsync(db, settings, method, settle.Paid),
            WarehouseID = warehouse, Notes = SalesService.Clip(buy.Notes, 255),
        };
        db.PurchaseInvoices.Add(inv);
        await db.SaveChangesAsync();
        for (int i = 0; i < buy.Lines.Count; i++)
        {
            var l = buy.Lines[i];
            var c = calc.Lines[i];
            db.PurchaseInvoiceDetails.Add(new PurchaseInvoiceDetails
            {
                PurchaseInvoiceID = inv.PurchaseInvoiceID, LineNumber = (short)(i + 1), ProductID = l.ProductID, Quantity = c.Quantity,
                UnitCost = c.UnitPrice, Discount = c.Discount, NetAmount = c.Net, VATRate = c.Rate, Tax = c.Tax, LineTotal = c.Total,
                UnitID = SalesService.IsBaseUnit(products[l.ProductID], l.UnitID) ? null : l.UnitID, UnitFactor = factors[i],
            });
        }
        await db.SaveChangesAsync();
        for (int i = 0; i < buy.Lines.Count; i++)
        {
            var l = buy.Lines[i];
            var c = calc.Lines[i];
            decimal pieces = c.Quantity * factors[i];
            // pieces in, at the cost of one piece (docs/72)
            decimal unitCost = PurchaseUnitCost(c.Net, c.Total, pieces, reclaim);
            await SalesService.ApplyStockMovementAsync(db, user, l.ProductID, pieces, TtPurchase, unitCost, "PURCHASE", inv.PurchaseInvoiceID,
                                                       inv.InvoiceNumber, warehouse, recalcAverage: true);
            if (factors[i] == 1)
            {
                await UpdateProductPricesAsync(db, l.ProductID, c.UnitPrice, l.NewSellingPrice, buy.SupplierID);
            }
            else
            {
                // a bigger unit: its own last price and shelf price
                await UpdateProductPricesAsync(db, l.ProductID, MoneyMath.RoundMoney(c.UnitPrice / factors[i], 4), null, buy.SupplierID);
                var unit = await db.ProductUnits.Where(u => u.ProductID == l.ProductID && u.UnitID == l.UnitID).ToListAsync();
                foreach (var u in unit)
                {
                    u.PurchasePrice = c.UnitPrice;
                    if (l.NewSellingPrice is { } np)
                    {
                        u.SellingPrice = np;
                    }
                }
                await db.SaveChangesAsync();
            }
        }
        if (settle.Remaining != 0)
        {
            await AdjustSupplierBalanceAsync(db, buy.SupplierID, settle.Remaining);
        }
        await tx.CommitAsync();
        Audit.LogAction(db, user, "PURCHASE", "PurchaseInvoices", inv.InvoiceNumber, "Total=" + InvoiceCalculator.Plain(calc.Total));
        if (changes.Count > 0)
        {
            Audit.LogAction(db, user, "PRICE_CHANGE", "Products", inv.InvoiceNumber, "Selling price:" + string.Concat(changes));
        }
        await db.SaveChangesAsync();
        return new SaleResult(SaveStatus.Saved, "", inv.PurchaseInvoiceID, inv.InvoiceNumber, calc.Total, settle.Change);
    }

    /// <summary>modPurchases.UpdateProductPrices: the last purchase price (without VAT), the default supplier when empty,
    /// and the new shelf price if typed.</summary>
    private static async Task UpdateProductPricesAsync(RetailDbContext db, int productId, decimal lastCost, decimal? newSellingPrice, int supplierId)
    {
        var p = await db.Products.FirstAsync(x => x.ProductID == productId);
        p.PurchasePrice = lastCost;
        p.SupplierID ??= supplierId;
        if (newSellingPrice is { } np)
        {
            p.SellingPrice = np;
        }
        p.UpdatedAt = DateTime.Now;
        await db.SaveChangesAsync();
    }

    /// <summary>modSales.AdjustBalance for a supplier (balance = what the shop owes him).</summary>
    private static async Task AdjustSupplierBalanceAsync(RetailDbContext db, int supplierId, decimal delta)
    {
        var s = await db.Suppliers.FirstAsync(x => x.SupplierID == supplierId);
        s.CurrentBalance += delta;
        await db.SaveChangesAsync();
    }

    // ------------------------------------------------------------------ the payment voucher

    /// <summary>modPurchases.PostSupplierPayment: confirmOverBalance = the user answered yes to the question on an amount
    /// over what is owed.</summary>
    public async Task<SaleResult> PostSupplierPaymentAsync(ActingUser user, SupplierPaymentRequest pay, bool confirmOverBalance = false)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, "frmSupplierPayment");
        if (open.Length > 0)
        {
            return SaleResult.Refused(open);
        }
        if (pay.Amount <= 0)
        {
            return SaleResult.Refused("المبلغ يجب أن يكون أكبر من صفر.");
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        string msg = await CheckSupplierAsync(db, pay.SupplierID);
        string baseCurrency = string.IsNullOrWhiteSpace(settings.CurrencyCode) ? "SAR" : settings.CurrencyCode;
        string code = string.IsNullOrWhiteSpace(pay.CurrencyCode) ? baseCurrency : pay.CurrencyCode.Trim();
        decimal rate = string.Equals(code, baseCurrency, StringComparison.OrdinalIgnoreCase) ? 1 : pay.ExchangeRate;
        if (msg.Length == 0)
        {
            msg = await SalesService.CurrencyProblemAsync(db, code, rate, baseCurrency);
        }
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }
        decimal foreign = pay.Amount;
        decimal amount = MoneyMath.RoundMoney(foreign * rate);
        decimal owed = await db.Suppliers.Where(s => s.SupplierID == pay.SupplierID).Select(s => s.CurrentBalance).FirstAsync();
        if (amount > owed && !confirmOverBalance)
        {
            return new SaleResult(SaveStatus.NeedsConfirmation,
                                  "المبلغ (" + InvoiceCalculator.Money(amount) + ") أكبر من المستحق للمورد (" + InvoiceCalculator.Money(owed) +
                                  ")، وسيصبح المورد مدينًا للمحل.\n" + "هل تريد المتابعة؟");
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        var p = new SupplierPayments
        {
            PaymentNumber = await NumberSequences.NextNumberAsync(db, "SUPPLIER_PAYMENT"),
            SupplierID = pay.SupplierID, PaymentDate = DateTime.Now, Amount = amount, CurrencyCode = code, ExchangeRate = rate,
            ForeignAmount = foreign, PaymentMethodID = pay.PaymentMethodID,
            CashBoxID = await Sales.CashBoxForAsync(db, user.EmployeeID, pay.PaymentMethodID, amount),
            BankID = await SalesService.BankForAsync(db, settings, pay.PaymentMethodID, amount), EmployeeID = user.EmployeeID,
            Notes = SalesService.Clip(pay.Notes, 255),
        };
        db.SupplierPayments.Add(p);
        await db.SaveChangesAsync();
        await AdjustSupplierBalanceAsync(db, pay.SupplierID, -amount);
        await tx.CommitAsync();
        Audit.LogAction(db, user, "SUPPLIER_PAYMENT", "SupplierPayments", p.PaymentNumber, "Amount=" + InvoiceCalculator.Plain(amount));
        await db.SaveChangesAsync();
        return new SaleResult(SaveStatus.Saved, "", p.PaymentID, p.PaymentNumber, amount);
    }

    // ------------------------------------------------------------------ what the screens read

    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    /// <summary>modProductUnits.UnitPurchasePrice: the unit's own purchase price, else the piece price x its factor.</summary>
    public static async Task<decimal> UnitPurchasePriceAsync(RetailDbContext db, Products p, int? unitId)
    {
        if (SalesService.IsBaseUnit(p, unitId))
        {
            return p.PurchasePrice;
        }
        var u = await db.ProductUnits.AsNoTracking().Where(x => x.IsActive && x.ProductID == p.ProductID && x.UnitID == unitId).FirstOrDefaultAsync();
        if (u is null)
        {
            return 0;
        }
        return u.PurchasePrice > 0 ? u.PurchasePrice : MoneyMath.RoundMoney(p.PurchasePrice * u.UnitFactor);
    }

    /// <summary>The price of a new purchase line in a unit (the purchase screen).</summary>
    public async Task<(decimal Cost, decimal SellingPrice)> LineCostAsync(int productId, int? unitId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var p = await db.Products.AsNoTracking().FirstOrDefaultAsync(x => x.ProductID == productId);
        if (p is null)
        {
            return (0, 0);
        }
        decimal factor = await SalesService.UnitFactorOfAsync(db, p, unitId);
        return (await UnitPurchasePriceAsync(db, p, unitId), factor > 0 ? await ShownSellingPriceAsync(db, p, unitId, factor) : 0);
    }

    /// <summary>modProductUnits.FindProductCode (a unit barcode gives its unit), then the active products whose name or
    /// code contains the text; with the last purchase price of the unit and the stock of the warehouse.</summary>
    public async Task<List<PurchaseProduct>> FindProductsAsync(int employeeId, string text, bool english, int warehouseId = 0, int take = 20)
    {
        string t = (text ?? "").Trim();
        if (t.Length == 0)
        {
            return [];
        }
        await using var db = await factory.CreateDbContextAsync();
        int warehouse = warehouseId > 0 ? warehouseId : await RepService.UserWarehouseAsync(db, employeeId);
        var found = new List<(Products P, int? Unit)>();
        var exact = await db.Products.AsNoTracking().Where(p => p.IsActive && (p.Barcode == t || p.ProductCode == t)).OrderBy(p => p.ProductID)
            .FirstOrDefaultAsync();
        if (exact is not null)
        {
            found.Add((exact, null));
        }
        else
        {
            var unit = await (from u in db.ProductUnits
                              join p in db.Products on u.ProductID equals p.ProductID
                              where u.IsActive && p.IsActive && u.Barcode == t
                              select new { p, u.UnitID }).FirstOrDefaultAsync();
            if (unit is not null)
            {
                found.Add((unit.p, unit.UnitID));
            }
            else
            {
                var rows = await db.Products.AsNoTracking()
                    .Where(p => p.IsActive && (p.ProductName.Contains(t) || (p.ProductNameEn != null && p.ProductNameEn.Contains(t)) ||
                                               p.ProductCode.Contains(t) || (p.Barcode != null && p.Barcode.Contains(t))))
                    .OrderBy(p => p.ProductName).Take(take).ToListAsync();
                found.AddRange(rows.Select(p => (p, (int?)null)));
            }
        }
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        var list = new List<PurchaseProduct>();
        foreach (var (p, unitId) in found)
        {
            int u = unitId ?? p.UnitID;
            string unitName = units.TryGetValue(u, out var un) ? Name(un.UnitName, un.UnitNameEn, english) : "";
            decimal factor = await SalesService.UnitFactorOfAsync(db, p, unitId);
            list.Add(new PurchaseProduct(p.ProductID, p.ProductCode, p.Barcode, Name(p.ProductName, p.ProductNameEn, english), unitId, unitName,
                                         await UnitPurchasePriceAsync(db, p, unitId), await ShownSellingPriceAsync(db, p, unitId, factor),
                                         await SalesService.AvailableQtyAsync(db, p.ProductID, warehouse), p.VATCategory));
        }
        return list;
    }

    /// <summary>The purchase invoice as it is shown (frmPurchaseView): the supplier, the lines in their unit.</summary>
    public async Task<PurchaseView?> PurchaseAsync(int invoiceId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await db.PurchaseInvoices.AsNoTracking().FirstOrDefaultAsync(x => x.PurchaseInvoiceID == invoiceId);
        if (h is null)
        {
            return null;
        }
        var s = await db.Suppliers.AsNoTracking().FirstAsync(x => x.SupplierID == h.SupplierID);
        var m = h.PaymentMethodID is { } mid ? await db.PaymentMethods.AsNoTracking().FirstOrDefaultAsync(x => x.PaymentMethodID == mid) : null;
        string emp = await db.Employees.Where(e => e.EmployeeID == h.EmployeeID).Select(e => e.EmployeeName).FirstOrDefaultAsync() ?? "";
        var w = await db.Warehouses.AsNoTracking().FirstOrDefaultAsync(x => x.WarehouseID == h.WarehouseID);
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        var lines = await (from d in db.PurchaseInvoiceDetails.AsNoTracking()
                           join p in db.Products on d.ProductID equals p.ProductID
                           where d.PurchaseInvoiceID == invoiceId
                           orderby d.LineNumber
                           select new { d, p.ProductCode, p.ProductName, p.ProductNameEn, BaseUnit = p.UnitID }).ToListAsync();
        return new PurchaseView(h.PurchaseInvoiceID, h.InvoiceNumber, h.SupplierInvoiceNo, h.InvoiceDate, h.SupplierID,
                                Name(s.SupplierName, s.SupplierNameEn, english), s.VATNumber, h.PaymentType,
                                m is null ? null : Name(m.MethodName, m.MethodNameEn, english), h.SubTotal, h.Discount, h.TaxableAmount,
                                h.Tax, h.TotalAmount, h.PaidAmount, h.RemainingAmount, h.DueDate, h.CurrencyCode, h.ExchangeRate,
                                h.ForeignAmount, h.Notes, emp, w is null ? null : Name(w.WarehouseName, w.WarehouseNameEn, english),
                                lines.Select(x =>
                                {
                                    int u = x.d.UnitID ?? x.BaseUnit;
                                    string un = units.TryGetValue(u, out var uu) ? Name(uu.UnitName, uu.UnitNameEn, english) : "";
                                    return new PurchaseLineView(x.d.LineNumber, x.d.ProductID, x.ProductCode, Name(x.ProductName, x.ProductNameEn, english),
                                                                un, x.d.Quantity, x.d.UnitCost, x.d.Discount, x.d.NetAmount, x.d.VATRate, x.d.Tax,
                                                                x.d.LineTotal);
                                }).ToList());
    }

    /// <summary>The purchase invoices of a period, newest first (a search in our number, the supplier's number or the
    /// supplier's name).</summary>
    public async Task<List<PurchaseRow>> PurchasesAsync(DateTime from, DateTime to, string? search, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        string s = (search ?? "").Trim();
        var q = from h in db.PurchaseInvoices.AsNoTracking()
                join sp in db.Suppliers on h.SupplierID equals sp.SupplierID
                join e in db.Employees on h.EmployeeID equals e.EmployeeID
                join w in db.Warehouses on h.WarehouseID equals w.WarehouseID
                where h.InvoiceDate >= start && h.InvoiceDate < end
                select new { h, sp.SupplierName, sp.SupplierNameEn, e.EmployeeName, w.WarehouseName, w.WarehouseNameEn };
        if (s.Length > 0)
        {
            q = q.Where(x => x.h.InvoiceNumber.Contains(s) || (x.h.SupplierInvoiceNo != null && x.h.SupplierInvoiceNo.Contains(s)) ||
                             x.SupplierName.Contains(s) || (x.SupplierNameEn != null && x.SupplierNameEn.Contains(s)));
        }
        var rows = await q.OrderByDescending(x => x.h.InvoiceDate).ThenByDescending(x => x.h.PurchaseInvoiceID).Take(take).ToListAsync();
        return rows.Select(x => new PurchaseRow(x.h.PurchaseInvoiceID, x.h.InvoiceNumber, x.h.SupplierInvoiceNo, x.h.InvoiceDate,
                                                Name(x.SupplierName, x.SupplierNameEn, english), x.h.PaymentType, x.h.TotalAmount, x.h.PaidAmount,
                                                x.h.RemainingAmount, Name(x.WarehouseName, x.WarehouseNameEn, english), x.EmployeeName)).ToList();
    }

    /// <summary>The payment vouchers of a period, newest first.</summary>
    public async Task<List<SupplierPaymentRow>> SupplierPaymentsAsync(DateTime from, DateTime to, string? search, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        string s = (search ?? "").Trim();
        var q = from p in db.SupplierPayments.AsNoTracking()
                join sp in db.Suppliers on p.SupplierID equals sp.SupplierID
                join e in db.Employees on p.EmployeeID equals e.EmployeeID
                join m in db.PaymentMethods on p.PaymentMethodID equals m.PaymentMethodID
                where p.PaymentDate >= start && p.PaymentDate < end
                select new { p, sp.SupplierName, sp.SupplierNameEn, e.EmployeeName, m.MethodName, m.MethodNameEn };
        if (s.Length > 0)
        {
            q = q.Where(x => x.p.PaymentNumber.Contains(s) || x.SupplierName.Contains(s) || (x.SupplierNameEn != null && x.SupplierNameEn.Contains(s)));
        }
        var rows = await q.OrderByDescending(x => x.p.PaymentDate).ThenByDescending(x => x.p.PaymentID).Take(take).ToListAsync();
        return rows.Select(x => new SupplierPaymentRow(x.p.PaymentID, x.p.PaymentNumber, x.p.PaymentDate, x.p.SupplierID,
                                                       Name(x.SupplierName, x.SupplierNameEn, english), x.p.Amount, x.p.CurrencyCode,
                                                       x.p.ForeignAmount, Name(x.MethodName, x.MethodNameEn, english), x.p.Notes,
                                                       x.EmployeeName)).ToList();
    }

    /// <summary>The active suppliers for the choice lists, with what the shop owes each.</summary>
    public async Task<List<(int Id, string Name, decimal Balance)>> SuppliersAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Suppliers.AsNoTracking().Where(s => s.IsActive).OrderBy(s => s.SupplierName).ToListAsync();
        return rows.Select(s => (s.SupplierID, Name(s.SupplierName, s.SupplierNameEn, english), s.CurrentBalance)).ToList();
    }

    /// <summary>The settings the purchase screen needs: the VAT rate, whether the store is VAT-registered, the currency.</summary>
    public async Task<(decimal VatRate, bool VatRegistered, string Currency)> PurchaseSettingsAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        return (s.VATRate, !string.IsNullOrEmpty(s.VATNumber), string.IsNullOrWhiteSpace(s.CurrencyCode) ? "SAR" : s.CurrencyCode);
    }
}
