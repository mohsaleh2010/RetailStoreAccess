using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

/// <summary>A line of the cart (tmpPOSLines): UnitID null = the smallest unit of the product; UnitPrice as typed (VAT
/// included when the prices include VAT); LineDiscount on the same basis.</summary>
public sealed record CartLine(int ProductID, decimal Quantity, decimal UnitPrice, decimal LineDiscount = 0, int? UnitID = null,
                              string? LineNote = null);

/// <summary>A sale (modSales.PostSaleFromCart): PaymentType CASH or CREDIT; Tendered null = the total of a cash sale,
/// nothing for a credit sale.</summary>
public sealed record SaleRequest(int CustomerID, string PaymentType, int? PaymentMethodID, decimal InvoiceDiscount, decimal? Tendered,
                                 string? Notes, IReadOnlyList<CartLine> Lines);

/// <summary>A sale recorded on a POS device (docs/83): its number on the device (ClientRef, kept in
/// SalesInvoices.ClientRef so a sale sent twice is added once), the invoice number the device printed, and the time it was
/// made (not after the time it reaches the server).</summary>
public sealed record DeviceSaleInfo(string ClientRef, string InvoiceNumber, DateTime TakenAt);

public sealed record SaleResult(SaveStatus Status, string Message, int InvoiceID = 0, string InvoiceNumber = "", decimal Total = 0,
                                decimal Change = 0)
{
    public static SaleResult Refused(string message) => new(SaveStatus.Refused, message);
}

/// <summary>A customer receipt voucher (modSales.PostCustomerPayment): Amount in CurrencyCode (empty = the program
/// currency) at ExchangeRate.</summary>
public sealed record PaymentRequest(int CustomerID, decimal Amount, int PaymentMethodID, string? Notes, string? CurrencyCode = null,
                                    decimal ExchangeRate = 1);

/// <summary>A line of a sales return (tmpReturnLines): the invoice line, the quantity returned in the unit of that line,
/// and ToStock = the goods go back to the stock.</summary>
public sealed record ReturnLineRequest(int SalesDetailID, decimal Quantity, bool ToStock = true);

/// <summary>A sales return (modSales.PostSalesReturn): RefundType CASH (paid back, PaymentMethodID) or CREDIT (taken off the
/// customer's balance); the cash customer is always paid back.</summary>
public sealed record SalesReturnRequest(int SalesInvoiceID, string? Reason, string RefundType, int? PaymentMethodID,
                                        ReturnLineRequest[] Lines);

/// <summary>A line of an invoice as the return screen shows it (LoadInvoiceForReturn): sold, returned before, what is left.</summary>
public sealed record ReturnableLine(int SalesDetailID, int ProductID, string ProductName, string UnitName, decimal SoldQty,
                                    decimal ReturnedQty, decimal AvailableQty, decimal UnitPrice, decimal LineTotal, decimal Tax,
                                    decimal ReturnedTotal, decimal ReturnedTax)
{
    /// <summary>modPOS.RecalcReturn: the credit-note value of a quantity of this line (the rule used for posting).</summary>
    public (string Problem, decimal Net, decimal Tax, decimal Total) Amounts(decimal qty) =>
        InvoiceCalculator.ReturnAmounts(SoldQty, LineTotal, this.Tax, ReturnedQty, ReturnedTotal, ReturnedTax, qty);
}

/// <summary>An invoice to return from: CashCustomer = the cash customer (always paid back in cash).</summary>
public sealed record ReturnableInvoice(int SalesInvoiceID, string InvoiceNumber, DateTime InvoiceDate, int CustomerID, string CustomerName,
                                       bool CashCustomer, string PaymentType, decimal TotalAmount, List<ReturnableLine> Lines);

/// <summary>A product found by a code or a name, with the unit of the code and the price of a new line.</summary>
public sealed record SaleProduct(int ProductID, string ProductCode, string? Barcode, string ProductName, int? UnitID, string UnitName,
                                 decimal Price, decimal Available, bool TrackStock, string VATCategory);

public sealed record UnitChoice(int UnitID, string UnitName, decimal Factor);

public sealed record InvoiceLineView(short LineNumber, int ProductID, string ProductCode, string ProductName, string UnitName,
                                     decimal Quantity, decimal UnitPrice, decimal Discount, decimal NetAmount, decimal VATRate,
                                     decimal Tax, decimal LineTotal, string? LineNote);

public sealed record InvoiceView(int SalesInvoiceID, string InvoiceNumber, DateTime InvoiceDate, int CustomerID, string CustomerName,
                                 string? CustomerVAT, string PaymentType, string? MethodName, decimal SubTotal, decimal Discount,
                                 decimal TaxableAmount, decimal Tax, decimal TotalAmount, decimal PaidAmount, decimal RemainingAmount,
                                 decimal AmountTendered, decimal ChangeDue, DateTime? DueDate, string? Notes, string InvoiceSubType,
                                 string? InvoiceUUID, int? ICV, string? QRCodeData, string ZatcaStatus, string EmployeeName,
                                 string? WarehouseName, string? RepName, List<InvoiceLineView> Lines);

public sealed record InvoiceRow(int SalesInvoiceID, string InvoiceNumber, DateTime InvoiceDate, string CustomerName, string PaymentType,
                                decimal TotalAmount, decimal PaidAmount, decimal RemainingAmount, string ZatcaStatus, string EmployeeName);

public sealed record PaymentRow(int PaymentID, string PaymentNumber, DateTime PaymentDate, int CustomerID, string CustomerName,
                                decimal Amount, string? CurrencyCode, decimal ForeignAmount, string? MethodName, string? Notes,
                                string EmployeeName);

/// <summary>Sales and customer receipts on the web and the API (plan 6C, docs/79): the C# of modSales.PostSaleFromCart and
/// PostCustomerPayment, with the same checks in the same order, the same Arabic messages and the same rows (the invoice,
/// its lines, the stock moves, the balance of the customer, the ZATCA fields, the audit trail), so Access rebuilds the
/// same journal entry (SyncJournal) and sends the e-invoice (frmEInvoices). The web works outside a rep day: the rep,
/// warehouse and box of a sale are those of the user (modVanSales.StartActingAsRep stays in Access).</summary>
public sealed class SalesService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const int TtSale = 2;
    public const int CashMethodId = 1;
    public const int BankTransferMethodId = 3;
    public const string NegativeStockQuestion = "هل تسمح بإتمام العملية";

    // ------------------------------------------------------------------ the user's defaults (outside a rep day)

    /// <summary>modCash.CurrentCashBoxID: the user's active box; else the main safe for who has CASH_BOX, else the first
    /// cashier box, else the first active box; 0 = none.</summary>
    public async Task<int> CurrentCashBoxAsync(RetailDbContext db, int employeeId)
    {
        int? v = await (from e in db.Employees
                        join b in db.CashBoxes on e.CashBoxID equals b.CashBoxID
                        where e.EmployeeID == employeeId && b.IsActive
                        select (int?)b.CashBoxID).FirstOrDefaultAsync();
        if (v is null && await permissions.HasPermissionAsync(employeeId, "CASH_BOX"))
        {
            v = await db.CashBoxes.Where(b => b.BoxType == "MAIN" && b.IsActive).MinAsync(b => (int?)b.CashBoxID);
        }
        v ??= await db.CashBoxes.Where(b => b.BoxType == "CASHIER" && b.IsActive).MinAsync(b => (int?)b.CashBoxID);
        v ??= await db.CashBoxes.Where(b => b.IsActive).MinAsync(b => (int?)b.CashBoxID);
        return v ?? 0;
    }

    /// <summary>modCash.CashBoxFor: only cash moves a box.</summary>
    public async Task<int?> CashBoxForAsync(RetailDbContext db, int employeeId, int? methodId, decimal amount)
    {
        if ((methodId ?? 0) != CashMethodId || amount == 0)
        {
            return null;
        }
        int box = await CurrentCashBoxAsync(db, employeeId);
        return box > 0 ? box : null;
    }

    /// <summary>CashBoxFor while acting as a rep: the custody box of the rep day (modVanSales.CurrentCashBoxID).</summary>
    public async Task<int?> CashBoxForAsync(RetailDbContext db, ActingUser user, int? methodId, decimal amount)
    {
        if (user.RepDay is { } day && (methodId ?? 0) == CashMethodId && amount != 0)
        {
            return day.CashBoxID;
        }
        return await CashBoxForAsync(db, user.EmployeeID, methodId, amount);
    }

    /// <summary>modBank.BankFor: a bank transfer goes to the default bank when it is active.</summary>
    public static async Task<int?> BankForAsync(RetailDbContext db, Settings settings, int? methodId, decimal amount)
    {
        if ((methodId ?? 0) != BankTransferMethodId || amount == 0 || settings.DefaultBankID is not { } bank)
        {
            return null;
        }
        return await db.Banks.AnyAsync(b => b.IsActive && b.BankID == bank) ? bank : null;
    }

    /// <summary>modCostCenters.CostCenterFor: the user's active centre, else the active default one.</summary>
    public static async Task<int?> CostCenterForAsync(RetailDbContext db, int employeeId) =>
        await (from e in db.Employees
               join c in db.CostCenters on e.CostCenterID equals c.CostCenterID
               where c.IsActive && e.EmployeeID == employeeId
               select (int?)c.CostCenterID).FirstOrDefaultAsync()
        ?? await db.CostCenters.Where(c => c.IsDefault && c.IsActive).MinAsync(c => (int?)c.CostCenterID);

    // ------------------------------------------------------------------ units and prices (modProductUnits, modPriceLists)

    internal static bool IsBaseUnit(Products p, int? unitId) => unitId is null || unitId == 0 || unitId == p.UnitID;

    /// <summary>modProductUnits.UnitFactorOf: 1 for the smallest unit, 0 when the unit is not one of its active units.</summary>
    public static async Task<decimal> UnitFactorOfAsync(RetailDbContext db, Products p, int? unitId) =>
        IsBaseUnit(p, unitId) ? 1
            : await db.ProductUnits.Where(u => u.IsActive && u.ProductID == p.ProductID && u.UnitID == unitId)
                .Select(u => (decimal?)u.UnitFactor).FirstOrDefaultAsync() ?? 0;

    internal static async Task<decimal> OwnUnitPriceAsync(RetailDbContext db, Products p, int? unitId) =>
        IsBaseUnit(p, unitId) ? 0
            : await db.ProductUnits.Where(u => u.IsActive && u.ProductID == p.ProductID && u.UnitID == unitId)
                .Select(u => (decimal?)u.SellingPrice).FirstOrDefaultAsync() ?? 0;

    /// <summary>modRepRoutes.RepProductPrice.</summary>
    private static async Task<decimal?> RepPriceAsync(RetailDbContext db, int repId, int productId) =>
        repId == 0 ? null
            : await db.RepPrices.Where(r => r.SalesRepID == repId && r.ProductID == productId).Select(r => (decimal?)r.UnitPrice).FirstOrDefaultAsync();

    /// <summary>modPriceLists.CustomerPriceListID and its discount.</summary>
    private static async Task<(int ListId, decimal Discount)> CustomerListAsync(RetailDbContext db, int customerId)
    {
        if (customerId == 0)
        {
            return (0, 0);
        }
        var l = await (from c in db.Customers
                       join pl in db.PriceLists on c.PriceListID equals pl.PriceListID
                       where pl.IsActive && c.CustomerID == customerId
                       select new { pl.PriceListID, pl.DiscountPercent }).FirstOrDefaultAsync();
        return l is null ? (0, 0) : (l.PriceListID, l.DiscountPercent);
    }

    /// <summary>modPriceLists.PriceListItemPrice: among the items valid on the day, the one that starts last.</summary>
    private static async Task<decimal?> ListItemPriceAsync(RetailDbContext db, int listId, int productId, int? unitId, DateTime day)
    {
        var d = day.Date;
        var rows = await db.PriceListItems.AsNoTracking()
            .Where(i => i.PriceListID == listId && i.ProductID == productId && i.UnitID == unitId
                        && (i.ValidFrom == null || i.ValidFrom <= d) && (i.ValidTo == null || i.ValidTo >= d)).ToListAsync();
        return rows.OrderByDescending(i => i.ValidFrom ?? DateTime.MinValue).ThenByDescending(i => i.PriceListItemID)
                   .Select(i => (decimal?)i.UnitPrice).FirstOrDefault();
    }

    /// <summary>modPriceLists.CustomerPrice: the customer's list (the item of the unit, else the piece item x the factor),
    /// else the usual price less the discount of the list.</summary>
    private static async Task<decimal> CustomerPriceAsync(RetailDbContext db, int customerId, Products p, int? unitId, decimal usual)
    {
        var (listId, discount) = await CustomerListAsync(db, customerId);
        if (listId == 0)
        {
            return usual;
        }
        bool baseUnit = IsBaseUnit(p, unitId);
        var item = await ListItemPriceAsync(db, listId, p.ProductID, baseUnit ? null : unitId, DateTime.Today);
        if (item is { } own)
        {
            return own;
        }
        if (!baseUnit)
        {
            decimal factor = await UnitFactorOfAsync(db, p, unitId);
            if (factor > 0 && await ListItemPriceAsync(db, listId, p.ProductID, null, DateTime.Today) is { } piece)
            {
                return MoneyMath.RoundMoney(piece * factor);
            }
        }
        return discount > 0 ? MoneyMath.RoundMoney(usual * (1 - discount)) : usual;
    }

    /// <summary>modProductUnits.UnitSalePrice: the price of a new sale line (the rep's price of the user, the unit's own
    /// price or the piece price x its factor, then the customer's price list).</summary>
    public static async Task<decimal> UnitSalePriceAsync(RetailDbContext db, Products p, int? unitId, int customerId, int repId)
    {
        decimal piece = await RepPriceAsync(db, repId, p.ProductID) ?? p.SellingPrice;
        decimal usual;
        if (IsBaseUnit(p, unitId))
        {
            usual = piece;
        }
        else
        {
            decimal own = await OwnUnitPriceAsync(db, p, unitId);
            usual = own > 0 ? own : MoneyMath.RoundMoney(piece * await UnitFactorOfAsync(db, p, unitId));
        }
        return await CustomerPriceAsync(db, customerId, p, unitId, usual);
    }

    /// <summary>The price of a new line for the POS screen.</summary>
    public async Task<decimal> LinePriceAsync(int employeeId, int productId, int? unitId, int customerId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var p = await db.Products.AsNoTracking().FirstOrDefaultAsync(x => x.ProductID == productId);
        return p is null ? 0 : await UnitSalePriceAsync(db, p, unitId, customerId, await RepService.RepIdOfAsync(db, employeeId));
    }

    /// <summary>The price is not a change (modSales.PostSaleFromCart step 1): the customer's list price
    /// (IsCustomerListPrice), a bigger unit at its listed price (IsListedUnitPrice), the selling price, a size price, or
    /// the rep's price (IsSessionRepPrice).</summary>
    private static async Task<bool> PriceChangedAsync(RetailDbContext db, Products p, int? unitId, decimal factor, decimal price,
                                                      int customerId, int repId)
    {
        if ((await CustomerListAsync(db, customerId)).ListId != 0 && price == await UnitSalePriceAsync(db, p, unitId, customerId, repId))
        {
            return false;
        }
        decimal? repPrice = await RepPriceAsync(db, repId, p.ProductID);
        if (factor != 1)
        {
            // modProductUnits.IsListedUnitPrice
            if (price == await OwnUnitPriceAsync(db, p, unitId))
            {
                return false;
            }
            return price != MoneyMath.RoundMoney(p.SellingPrice * factor)
                   && price != MoneyMath.RoundMoney((repPrice ?? p.SellingPrice) * factor);
        }
        if (price == p.SellingPrice || price == (p.SizePriceM ?? -1) || price == (p.SizePriceL ?? -1))
        {
            return false;
        }
        return repPrice is not { } rp || rp != price;
    }

    // ------------------------------------------------------------------ stock (modWarehouses, modSalesOrders)

    /// <summary>modWarehouses.AvailableQty.</summary>
    public static async Task<decimal> AvailableQtyAsync(RetailDbContext db, int productId, int warehouseId)
    {
        if (warehouseId == RepService.MainWarehouseId && !await db.InventoryTransactions.AnyAsync(t => t.WarehouseID != RepService.MainWarehouseId))
        {
            return await db.Products.Where(p => p.ProductID == productId).Select(p => p.CurrentQuantity).FirstOrDefaultAsync();
        }
        return await db.InventoryTransactions.Where(t => t.WarehouseID == warehouseId && t.ProductID == productId)
            .SumAsync(t => (decimal?)t.Quantity) ?? 0;
    }

    private static readonly string[] OpenOrderStates = ["APPROVED", "PREPARING", "PARTIAL"];

    /// <summary>modSalesOrders.ReservedQty: what the open orders still need of the product in the warehouse
    /// (ordered less delivered by posted delivery notes, qrySalesOrderLineStatus).</summary>
    public static async Task<decimal> ReservedQtyAsync(RetailDbContext db, int productId, int warehouseId)
    {
        var lines = await (from l in db.SalesOrderLines
                           join o in db.SalesOrders on l.SalesOrderID equals o.SalesOrderID
                           where l.ProductID == productId && o.WarehouseID == warehouseId && OpenOrderStates.Contains(o.Status)
                           select new { l.SalesOrderLineID, l.Quantity }).ToListAsync();
        if (lines.Count == 0)
        {
            return 0;
        }
        var ids = lines.Select(l => l.SalesOrderLineID).ToList();
        var delivered = await (from nl in db.StockNoteLines
                               join n in db.StockNotes on nl.StockNoteID equals n.StockNoteID
                               where n.Status != "DRAFT" && nl.SalesOrderLineID != null && ids.Contains(nl.SalesOrderLineID.Value)
                               group nl by nl.SalesOrderLineID into g
                               select new { Line = g.Key!.Value, Qty = g.Sum(x => x.Quantity) }).ToDictionaryAsync(x => x.Line, x => x.Qty);
        return lines.Sum(l => Math.Max(0, l.Quantity - delivered.GetValueOrDefault(l.SalesOrderLineID)));
    }

    /// <summary>modSales.CheckStockAvailable for a sale: Refused, NeedsConfirmation (the user may approve: the question),
    /// or Saved ("" = enough stock, else the list approved by the settings).</summary>
    internal async Task<(SaveStatus Status, string Message, string ShortList)> CheckStockAsync(
        RetailDbContext db, Settings settings, int employeeId, List<(Products P, decimal Pieces)> needs, int warehouseId, string operation)
    {
        string shortList = "";
        foreach (var g in needs.GroupBy(n => n.P.ProductID))
        {
            var p = g.First().P;
            decimal need = p.TrackStock ? g.Sum(x => x.Pieces) : 0;   // made to order: no stock limit
            decimal reserved = await ReservedQtyAsync(db, p.ProductID, warehouseId);
            decimal available = await AvailableQtyAsync(db, p.ProductID, warehouseId) - reserved;
            if (need > available)
            {
                shortList += "\n- " + p.ProductName + ": المطلوب " + InvoiceCalculator.Plain(need) + " والمتوفر " + InvoiceCalculator.Plain(available);
                if (reserved > 0)
                {
                    shortList += " (بعد " + InvoiceCalculator.Plain(reserved) + " محجوزة لطلبيات العملاء)";
                }
            }
        }
        if (shortList.Length == 0)
        {
            return (SaveStatus.Saved, "", "");
        }
        if (settings.AllowNegativeStock)
        {
            return (SaveStatus.Saved, operation + " - allowed by settings:" + shortList, shortList);
        }
        if (await permissions.HasPermissionAsync(employeeId, "ALLOW_NEGATIVE_STOCK"))
        {
            return (SaveStatus.NeedsConfirmation, "الكمية المتوفرة غير كافية:" + shortList + "\n\n" + NegativeStockQuestion + " (" +
                                                  operation + ") رغم ذلك؟ (بصلاحية مدير النظام)", shortList);
        }
        return (SaveStatus.Refused, "الكمية المتوفرة غير كافية:" + shortList + "\n" + operation +
                                    " بأكثر من المتوفر يحتاج موافقة مدير النظام.", shortList);
    }

    // ------------------------------------------------------------------ customer checks (modSales.CheckCustomer, modAging)

    /// <summary>modAging.ComputeAging for one customer: his open documents on the day.</summary>
    public static async Task<List<AgingItem>> CustomerAgingAsync(RetailDbContext db, int customerId, DateTime asOf)
    {
        var settings = await SettingsReader.ReadAsync(db);
        var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == customerId);
        int terms = c.PaymentTermsDays ?? 0;
        DateTime openingOn = settings.OpeningDate ?? c.CreatedAt;
        var debits = (await db.SalesInvoices.AsNoTracking().Where(h => h.CustomerID == customerId && h.RemainingAmount > 0)
                .Select(h => new { h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate, h.DueDate, h.RemainingAmount }).ToListAsync())
            .Select(h => new AgingDebit(customerId, "INVOICE", h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate,
                                        h.DueDate ?? h.InvoiceDate.Date.AddDays(terms), h.RemainingAmount)).ToList();
        if (c.OpeningBalance > 0)
        {
            debits.Add(new AgingDebit(customerId, "OPENING", customerId, "رصيد افتتاحي", openingOn, openingOn, c.OpeningBalance));
        }
        var credits = (await db.SalesReturns.AsNoTracking().Where(r => r.CustomerID == customerId && r.TotalAmount - r.RefundedAmount > 0)
                .Select(r => new { r.SalesReturnID, r.ReturnDate, r.TotalAmount, r.RefundedAmount, r.SalesInvoiceID }).ToListAsync())
            .Select(r => new AgingCredit(customerId, "RETURN", r.SalesReturnID, r.ReturnDate, r.TotalAmount - r.RefundedAmount, r.SalesInvoiceID))
            .ToList();
        var payments = await db.CustomerPayments.AsNoTracking().Where(p => p.CustomerID == customerId)
            .Select(p => new { p.PaymentID, p.PaymentDate, p.Amount, p.SalesInvoiceID }).ToListAsync();
        credits.AddRange(payments.Select(p => new AgingCredit(customerId, "PAYMENT", p.PaymentID, p.PaymentDate, p.Amount, 0)));
        if (c.OpeningBalance < 0)
        {
            credits.Add(new AgingCredit(customerId, "OPENING", customerId, openingOn, -c.OpeningBalance, 0));
        }
        credits.AddRange((await db.Cheques.AsNoTracking().Where(q => q.CustomerID == customerId && q.Direction == "IN" && q.Status != "BOUNCED")
                .Select(q => new { q.ChequeID, q.IssueDate, q.Amount }).ToListAsync())
            .Select(q => new AgingCredit(customerId, "CHEQUE", q.ChequeID, q.IssueDate, q.Amount, 0)));
        var payIds = payments.Select(p => p.PaymentID).ToList();
        var allocations = await db.CustomerAllocations.AsNoTracking().Where(a => payIds.Contains(a.PaymentID))
            .Select(a => new AgingLink(a.PaymentID, a.SalesInvoiceID, a.Amount)).ToListAsync();
        var allocated = allocations.Select(a => a.PaymentID).ToHashSet();
        allocations.AddRange(payments.Where(p => p.SalesInvoiceID is not null && !allocated.Contains(p.PaymentID))
                                     .Select(p => new AgingLink(p.PaymentID, p.SalesInvoiceID!.Value, p.Amount)));
        return CustomerAging.Compute(debits, credits, allocations, asOf);
    }

    /// <summary>modSales.CheckCustomer.</summary>
    private static async Task<string> CheckCustomerAsync(RetailDbContext db, Settings settings, Customers? c, bool isCredit, decimal remaining)
    {
        if (c is null)
        {
            return "العميل غير موجود.";
        }
        if (!c.IsActive)
        {
            return "العميل «" + c.CustomerName + "» غير نشط.";
        }
        if (!isCredit && remaining <= 0)
        {
            return "";
        }
        if (!c.AllowCredit)
        {
            return "العميل «" + c.CustomerName + "» غير مسموح له بالبيع الآجل.\n" + "للبيع الآجل اختر عميلًا مسجلًا.";
        }
        if (c.CreditLimit > 0 && c.CurrentBalance + remaining > c.CreditLimit)
        {
            return "تجاوز حد الائتمان: الرصيد " + InvoiceCalculator.Money(c.CurrentBalance) + " + المتبقي " +
                   InvoiceCalculator.Money(remaining) + " أكبر من الحد " + InvoiceCalculator.Money(c.CreditLimit) + ".";
        }
        int blockDays = settings.CreditBlockDays ?? 0;
        if (blockDays > 0)
        {
            int late = CustomerAging.MaxDaysLate(await CustomerAgingAsync(db, c.CustomerID, DateTime.Today), DateTime.Today);
            if (late > blockDays)
            {
                return "على العميل «" + c.CustomerName + "» فاتورة متأخرة " + late + " يومًا عن استحقاقها " + "(الحد " + blockDays +
                       " يومًا).\n" + "حصّل المتأخر أولًا، أو بِع نقدًا.";
            }
        }
        return "";
    }

    /// <summary>modRepRoutes.RepCreditProblem: the credit of a new sale against the limit of the rep of the invoice.</summary>
    private static async Task<string> RepCreditProblemAsync(RetailDbContext db, int? repId, decimal remaining)
    {
        if (repId is not { } rep || remaining <= 0)
        {
            return "";
        }
        var r = await db.SalesReps.AsNoTracking().FirstOrDefaultAsync(x => x.SalesRepID == rep);
        if (r is null || r.CreditLimit <= 0)
        {
            return "";
        }
        decimal bal = await RepService.CustomersBalanceAsync(db, rep);
        return bal + remaining > r.CreditLimit
            ? "تجاوز حد ائتمان المندوب «" + r.RepName + "»: أرصدة عملائه " + InvoiceCalculator.Money(bal) + " + المتبقي " +
              InvoiceCalculator.Money(remaining) + " أكبر من الحد " + InvoiceCalculator.Money(r.CreditLimit) + "."
            : "";
    }

    /// <summary>modSales.CheckDiscountAllowed: the line and invoice discounts against the user's limit (the rep's when
    /// the user is a rep with a limit), unless DISCOUNT_OVERRIDE.</summary>
    private async Task<string> CheckDiscountAsync(RetailDbContext db, int employeeId, int repId, CalcResult calc,
                                                  IReadOnlyList<CartLine> lines, decimal invoiceDiscount)
    {
        decimal lineDisc = lines.Sum(l => l.LineDiscount);
        if (calc.Gross + lineDisc <= 0)
        {
            return "";
        }
        double pct = (double)(lineDisc + invoiceDiscount) / (double)(calc.Gross + lineDisc);
        if (pct <= 0)
        {
            return "";
        }
        double maxPct = (double)await db.Employees.Where(e => e.EmployeeID == employeeId).Select(e => e.MaxDiscountPercent).FirstOrDefaultAsync();
        if (repId != 0)
        {
            double lim = (double)await db.SalesReps.Where(r => r.SalesRepID == repId).Select(r => r.MaxDiscountPercent).FirstOrDefaultAsync();
            if (lim > 0)
            {
                maxPct = lim;
            }
        }
        if (pct > maxPct && !await permissions.HasPermissionAsync(employeeId, "DISCOUNT_OVERRIDE"))
        {
            return "الخصم (" + Percent(pct) + ") أكبر من الحد المسموح لك (" + Percent(maxPct) + ").";
        }
        return "";
    }

    /// <summary>Format$(x, "0.0%").</summary>
    public static string Percent(double v) =>
        Math.Round(v * 100, 1, MidpointRounding.AwayFromZero).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + "%";

    // ------------------------------------------------------------------ the sale

    /// <summary>modSales.PostSaleFromCart (a POS sale: no delivery note, quotation or restaurant order). confirmStock = the
    /// user answered yes to the question on the stock (ALLOW_NEGATIVE_STOCK).</summary>
    public async Task<SaleResult> PostSaleAsync(ActingUser user, SaleRequest sale, bool confirmStock = false, DeviceSaleInfo? device = null,
                                                string? clientRef = null)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, "frmPOS");
        if (open.Length > 0)
        {
            return SaleResult.Refused(open);
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        decimal vatRate = settings.VATRate;
        int repId = user.RepDay?.SalesRepID ?? await RepService.RepIdOfAsync(db, user.EmployeeID);

        // ---- 1. cart -> calculation engine
        var ids = sale.Lines.Select(l => l.ProductID).Distinct().ToList();
        var products = await db.Products.AsNoTracking().Where(p => ids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID);
        var calcLines = new List<CalcLineIn>();
        var factors = new List<decimal>();
        bool priceChanged = false;
        foreach (var l in sale.Lines)
        {
            if (!products.TryGetValue(l.ProductID, out var p))
            {
                return SaleResult.Refused("أحد الأصناف في الفاتورة لم يعد موجودًا. احذفه من الفاتورة.");
            }
            if (!p.IsActive)
            {
                return SaleResult.Refused("المنتج «" + p.ProductName + "» غير نشط ولا يمكن بيعه.");
            }
            decimal factor = await UnitFactorOfAsync(db, p, l.UnitID);
            if (factor <= 0)
            {
                return SaleResult.Refused("الوحدة المختارة لـ «" + p.ProductName + "» لم تعد من وحداته. اختر الوحدة من جديد.");
            }
            if (await PriceChangedAsync(db, p, l.UnitID, factor, l.UnitPrice, sale.CustomerID, repId))
            {
                priceChanged = true;
            }
            calcLines.Add(new CalcLineIn(l.Quantity, l.UnitPrice, l.LineDiscount, p.VATCategory == "S" ? vatRate : 0));
            factors.Add(factor);
        }
        var calc = InvoiceCalculator.Calc(calcLines, sale.InvoiceDiscount, settings.PricesIncludeVAT);
        if (!calc.Ok)
        {
            return SaleResult.Refused(calc.Problem);
        }

        // ---- 2. permissions: price changes and discounts
        if (priceChanged && !await permissions.HasPermissionAsync(user.EmployeeID, "PRICE_OVERRIDE"))
        {
            return SaleResult.Refused("لا تملك صلاحية تعديل سعر البيع.");
        }
        string msg = await CheckDiscountAsync(db, user.EmployeeID, repId, calc, sale.Lines, sale.InvoiceDiscount);
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }

        // ---- 3. customer and payment
        bool isCredit = sale.PaymentType == "CREDIT";
        decimal tend = sale.Tendered ?? (isCredit ? 0 : calc.Total);
        var settle = InvoiceCalculator.Settle(calc.Total, tend, isCredit);
        if (settle.Problem.Length > 0)
        {
            return SaleResult.Refused(settle.Problem);
        }
        var customer = await db.Customers.AsNoTracking().FirstOrDefaultAsync(c => c.CustomerID == sale.CustomerID);
        msg = await CheckCustomerAsync(db, settings, customer, isCredit, settle.Remaining);
        int? rep = msg.Length == 0 ? user.RepDay?.SalesRepID ?? await RepService.SalesRepForAsync(db, sale.CustomerID, user.EmployeeID) : null;
        if (msg.Length == 0)
        {
            msg = await RepCreditProblemAsync(db, rep, settle.Remaining);
        }
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }
        int warehouse = user.RepDay?.WarehouseID ?? await RepService.UserWarehouseAsync(db, user.EmployeeID);
        var stock = await CheckStockAsync(db, settings, user.EmployeeID,
                                          sale.Lines.Select((l, i) => (products[l.ProductID], calc.Lines[i].Quantity * factors[i])).ToList(),
                                          warehouse, "البيع");
        if (stock.Status == SaveStatus.Refused || (stock.Status == SaveStatus.NeedsConfirmation && !confirmStock))
        {
            return new SaleResult(stock.Status, stock.Message);
        }

        // ---- 4. post everything in one transaction
        await using var tx = await db.Database.BeginTransactionAsync();
        if (stock.ShortList.Length > 0)
        {
            Audit.LogAction(db, user, "NEGATIVE_STOCK", "Products", "",
                            stock.Status == SaveStatus.Saved ? stock.Message : "البيع" + " - approved by user:" + stock.ShortList);
        }
        var now = DateTime.Now;
        if (device is not null && device.TakenAt < now)
        {
            now = device.TakenAt;
        }
        int? method = sale.PaymentMethodID;
        var inv = new SalesInvoices
        {
            InvoiceNumber = device?.InvoiceNumber ?? await NumberSequences.NextNumberAsync(db, "SALES_INVOICE"),
            ClientRef = device?.ClientRef ?? clientRef,
            InvoiceDate = now, CustomerID = sale.CustomerID, EmployeeID = user.EmployeeID, PaymentType = sale.PaymentType,
            PaymentMethodID = method, SubTotal = calc.SubTotal, Discount = calc.Discount, TaxableAmount = calc.Taxable, Tax = calc.Tax,
            TotalAmount = calc.Total, PaidAmount = settle.Paid, RemainingAmount = settle.Remaining,
            DueDate = settle.Remaining > 0 ? now.Date.AddDays(customer!.PaymentTermsDays ?? 0) : null,
            AmountTendered = tend, ChangeDue = settle.Change,
            CashBoxID = await CashBoxForAsync(db, user, method ?? CashMethodId, settle.Paid),
            BankID = await BankForAsync(db, settings, method, settle.Paid),
            CostCenterID = await CostCenterForAsync(db, user.EmployeeID), SalesRepID = rep, WarehouseID = warehouse,
            Notes = Clip(sale.Notes, 255),
        };
        await SetZatcaFieldsAsync(db, settings, inv, string.IsNullOrEmpty(customer!.VATNumber) ? "SIMPLIFIED" : "STANDARD", now);
        db.SalesInvoices.Add(inv);
        await db.SaveChangesAsync();
        for (int i = 0; i < sale.Lines.Count; i++)
        {
            var l = sale.Lines[i];
            var p = products[l.ProductID];
            var c = calc.Lines[i];
            db.SalesInvoiceDetails.Add(new SalesInvoiceDetails
            {
                SalesInvoiceID = inv.SalesInvoiceID, LineNumber = (short)(i + 1), ProductID = l.ProductID, Quantity = c.Quantity,
                UnitPrice = c.UnitPrice, Discount = c.Discount, NetAmount = c.Net, VATCategory = p.VATCategory, VATRate = c.Rate,
                Tax = c.Tax, LineTotal = c.Total, UnitCost = p.AverageCost * factors[i], LineNote = Clip(l.LineNote, 100),
                UnitID = IsBaseUnit(p, l.UnitID) ? null : l.UnitID, UnitFactor = factors[i],
            });
        }
        await db.SaveChangesAsync();
        for (int i = 0; i < sale.Lines.Count; i++)
        {
            var p = products[sale.Lines[i].ProductID];
            await ApplyStockMovementAsync(db, user, p.ProductID, -calc.Lines[i].Quantity * factors[i], TtSale, p.AverageCost, "SALE",
                                          inv.SalesInvoiceID, inv.InvoiceNumber, warehouse, when: now);
        }
        if (settle.Remaining != 0)
        {
            await AdjustBalanceAsync(db, sale.CustomerID, settle.Remaining);
        }
        Audit.LogAction(db, user, "SALE", "SalesInvoices", inv.InvoiceNumber, "Total=" + InvoiceCalculator.Plain(calc.Total) +
                                                                            (device is null ? "" : " pos " + device.ClientRef));
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        // the e-invoice: a PENDING document is sent by Access (modEInvoice, frmEInvoices); it never fails the sale
        return new SaleResult(SaveStatus.Saved, "", inv.SalesInvoiceID, inv.InvoiceNumber, calc.Total, settle.Change);
    }

    /// <summary>modSales.SetZatcaFields: the type, UUID, ICV, the first e-invoice status and the phase-1 QR (Saudi
    /// VAT-registered sellers).</summary>
    private static async Task SetZatcaFieldsAsync(RetailDbContext db, Settings settings, SalesInvoices inv, string subType, DateTime when)
    {
        inv.InvoiceSubType = subType;
        inv.InvoiceTypeCode = "388";
        inv.InvoiceUUID = Guid.NewGuid().ToString().ToLowerInvariant();
        inv.ICV = int.Parse(await NumberSequences.NextNumberAsync(db, "ZATCA_ICV"));
        inv.ZatcaStatus = settings.EInvoiceEnabled || settings.ZatcaPhase == 2 ? "PENDING" : "NOT_SENT";
        string vat = settings.VATNumber ?? "";
        if (vat.Length > 0 && CountryRules.Normalize(settings.CountryCode) == "SA")
        {
            inv.QRCodeData = ZatcaQr.Build(settings.StoreName ?? "", vat, when, inv.TotalAmount, inv.Tax);
        }
    }

    /// <summary>modSales.ApplyStockMovement: the product's quantity and a stock move (no move for a product that does not
    /// track stock). QuantityAfter = the total of every warehouse.</summary>
    public static async Task ApplyStockMovementAsync(RetailDbContext db, ActingUser user, int productId, decimal qty, int typeId,
                                                     decimal unitCost, string refType, int refId, string refNumber, int warehouseId,
                                                     bool recalcAverage = false, DateTime? when = null)
    {
        var p = await db.Products.FirstAsync(x => x.ProductID == productId);
        if (!p.TrackStock)
        {
            return;
        }
        decimal newQty = p.CurrentQuantity + qty;
        if (recalcAverage)
        {
            p.AverageCost = InvoiceCalculator.WeightedAverage(p.CurrentQuantity, p.AverageCost, qty, unitCost);
        }
        p.CurrentQuantity = newQty;
        p.UpdatedAt = DateTime.Now;
        db.InventoryTransactions.Add(new InventoryTransactions
        {
            TransactionDate = when ?? DateTime.Now, ProductID = productId, TransactionTypeID = typeId, Quantity = qty, UnitCost = unitCost,
            QuantityAfter = newQty, ReferenceType = refType, ReferenceID = refId > 0 ? refId : null,
            ReferenceNumber = Clip(refNumber, 20), EmployeeID = user.EmployeeID, WarehouseID = warehouseId,
        });
        await db.SaveChangesAsync();
    }

    /// <summary>modSales.AdjustBalance for a customer.</summary>
    private static async Task AdjustBalanceAsync(RetailDbContext db, int customerId, decimal delta)
    {
        var c = await db.Customers.FirstAsync(x => x.CustomerID == customerId);
        c.CurrentBalance += delta;
        await db.SaveChangesAsync();
    }

    internal static string? Clip(string? s, int n)
    {
        string t = (s ?? "").Trim();
        return t.Length == 0 ? null : (t.Length <= n ? t : t[..n]);
    }

    // ------------------------------------------------------------------ the sales return (credit note)

    public const int TtSalesReturn = 4;

    /// <summary>modPOS.ReturnFindInvoice + LoadInvoiceForReturn: the invoice of a number with its lines (sold, returned
    /// before, what is left to return); null = no invoice has this number.</summary>
    public async Task<ReturnableInvoice?> ReturnableAsync(string invoiceNumber, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        string number = (invoiceNumber ?? "").Trim(), upper = number.ToUpperInvariant();      // Access compares without case
        var h = await db.SalesInvoices.AsNoTracking().FirstOrDefaultAsync(x => x.InvoiceNumber == number || x.InvoiceNumber == upper);
        if (h is null)
        {
            return null;
        }
        var settings = await SettingsReader.ReadAsync(db);
        var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == h.CustomerID);
        var details = await db.SalesInvoiceDetails.AsNoTracking().Where(d => d.SalesInvoiceID == h.SalesInvoiceID)
            .OrderBy(d => d.LineNumber).ToListAsync();
        var ids = details.Select(d => d.SalesDetailID).ToList();
        var returned = await db.SalesReturnDetails.AsNoTracking().Where(r => ids.Contains(r.SalesDetailID))
            .GroupBy(r => r.SalesDetailID).Select(g => new { g.Key, Qty = g.Sum(r => r.Quantity), Total = g.Sum(r => r.LineTotal), Tax = g.Sum(r => r.Tax) })
            .ToDictionaryAsync(g => g.Key);
        var pids = details.Select(d => d.ProductID).Distinct().ToList();
        var products = await db.Products.AsNoTracking().Where(p => pids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID);
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID, u => Name(u.UnitName, u.UnitNameEn, english));
        var lines = details.Select(d =>
        {
            var p = products[d.ProductID];
            var back = returned.GetValueOrDefault(d.SalesDetailID);
            string unit = units.GetValueOrDefault(d.UnitID ?? p.UnitID) ?? "";
            decimal qty = back?.Qty ?? 0;
            return new ReturnableLine(d.SalesDetailID, d.ProductID, Name(p.ProductName, p.ProductNameEn, english), unit, d.Quantity, qty,
                                      d.Quantity - qty, d.UnitPrice, d.LineTotal, d.Tax, back?.Total ?? 0, back?.Tax ?? 0);
        }).ToList();
        return new ReturnableInvoice(h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate, h.CustomerID, Name(c.CustomerName, c.CustomerNameEn, english),
                                     h.CustomerID == settings.DefaultCustomerID, h.PaymentType, h.TotalAmount, lines);
    }

    /// <summary>modSales.PrevReturned: what was returned before from an invoice line.</summary>
    private static async Task<(decimal Qty, decimal Total, decimal Tax)> PrevReturnedAsync(RetailDbContext db, int salesDetailId)
    {
        var rows = await db.SalesReturnDetails.AsNoTracking().Where(r => r.SalesDetailID == salesDetailId)
            .Select(r => new { r.Quantity, r.LineTotal, r.Tax }).ToListAsync();
        return (rows.Sum(r => r.Quantity), rows.Sum(r => r.LineTotal), rows.Sum(r => r.Tax));
    }

    /// <summary>modSales.PostSalesReturn (and SaveReturn: the screen and its ADD action): the credit note of an invoice with
    /// the same checks, order, messages and rows as Access; clientRef = its number on a POS device (docs/86).</summary>
    public async Task<SaleResult> PostSalesReturnAsync(ActingUser user, SalesReturnRequest ret, string? clientRef = null)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, "frmSalesReturn");
        if (open.Length == 0)
        {
            open = await permissions.ActionProblemAsync(user.EmployeeID, "frmSalesReturn", ScreenAction.Add);
        }
        if (open.Length > 0)
        {
            return SaleResult.Refused(open);
        }
        string reason = (ret.Reason ?? "").Trim();
        if (reason.Length == 0)
        {
            return SaleResult.Refused("سبب الإرجاع مطلوب (إلزامي في الإشعار الدائن).");
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        var inv = await db.SalesInvoices.AsNoTracking().FirstOrDefaultAsync(h => h.SalesInvoiceID == ret.SalesInvoiceID);
        if (inv is null)
        {
            return SaleResult.Refused("الفاتورة الأصلية غير موجودة.");
        }
        string refundType = inv.CustomerID == settings.DefaultCustomerID ? "CASH" : ret.RefundType;
        var lines = new List<(SalesInvoiceDetails D, ReturnLineRequest R, decimal Net, decimal Tax, decimal Total, decimal Disc, decimal Factor)>();
        foreach (var r in (ret.Lines ?? []).Where(l => l.Quantity > 0))
        {
            var d = await db.SalesInvoiceDetails.AsNoTracking().FirstOrDefaultAsync(x => x.SalesDetailID == r.SalesDetailID);
            if (d is null || d.SalesInvoiceID != ret.SalesInvoiceID)
            {
                return SaleResult.Refused("سطر لا ينتمي إلى الفاتورة الأصلية.");
            }
            decimal factor = d.UnitFactor <= 0 ? 1 : d.UnitFactor;          // the unit of the invoice line (docs/72)
            var prev = await PrevReturnedAsync(db, d.SalesDetailID);
            var a = InvoiceCalculator.ReturnAmounts(d.Quantity, d.LineTotal, d.Tax, prev.Qty, prev.Total, prev.Tax, r.Quantity);
            if (a.Problem.Length > 0)
            {
                return SaleResult.Refused(a.Problem);
            }
            decimal disc = Math.Max(0, MoneyMath.RoundMoney(r.Quantity * d.UnitPrice) - a.Net);
            lines.Add((d, r, a.Net, a.Tax, a.Total, disc, factor));
        }
        if (lines.Count == 0)
        {
            return SaleResult.Refused("حدد الكمية المرتجعة لصنف واحد على الأقل.");
        }
        decimal sumNet = lines.Sum(l => l.Net), sumTax = lines.Sum(l => l.Tax), sumTotal = lines.Sum(l => l.Total), sumDisc = lines.Sum(l => l.Disc);
        decimal refunded = refundType == "CASH" ? sumTotal : 0;

        await using var tx = await db.Database.BeginTransactionAsync();
        var now = DateTime.Now;
        var row = new SalesReturns
        {
            ReturnNumber = await NumberSequences.NextNumberAsync(db, "SALES_RETURN"), ReturnDate = now, SalesInvoiceID = inv.SalesInvoiceID,
            CostCenterID = inv.CostCenterID, SalesRepID = inv.SalesRepID,
            WarehouseID = inv.WarehouseID > 0 ? inv.WarehouseID : await RepService.UserWarehouseAsync(db, user.EmployeeID),
            CustomerID = inv.CustomerID, EmployeeID = user.EmployeeID, Reason = reason.Length <= 255 ? reason : reason[..255],
            RefundType = refundType, PaymentMethodID = ret.PaymentMethodID, SubTotal = sumNet + sumDisc, Discount = sumDisc,
            TaxableAmount = sumNet, Tax = sumTax, TotalAmount = sumTotal, RefundedAmount = refunded,
            CashBoxID = await CashBoxForAsync(db, user.EmployeeID, ret.PaymentMethodID ?? CashMethodId, refunded),
            BankID = await BankForAsync(db, settings, ret.PaymentMethodID, refunded), ClientRef = clientRef,
        };
        await SetReturnZatcaFieldsAsync(db, settings, row, inv.InvoiceSubType, now);
        db.SalesReturns.Add(row);
        await db.SaveChangesAsync();
        foreach (var l in lines)
        {
            db.SalesReturnDetails.Add(new SalesReturnDetails
            {
                SalesReturnID = row.SalesReturnID, SalesDetailID = l.D.SalesDetailID, ProductID = l.D.ProductID, Quantity = l.R.Quantity,
                UnitPrice = l.D.UnitPrice, Discount = l.Disc, NetAmount = l.Net, VATCategory = l.D.VATCategory, VATRate = l.D.VATRate,
                Tax = l.Tax, LineTotal = l.Total, UnitCost = l.D.UnitCost, ReturnToStock = l.R.ToStock, UnitID = l.D.UnitID,
                UnitFactor = l.Factor,
            });
        }
        await db.SaveChangesAsync();
        foreach (var l in lines.Where(l => l.R.ToStock))
        {
            await ApplyStockMovementAsync(db, user, l.D.ProductID, l.R.Quantity * l.Factor, TtSalesReturn, l.D.UnitCost / l.Factor, "SALES_RETURN",
                                          row.SalesReturnID, row.ReturnNumber, row.WarehouseID, recalcAverage: true, when: now);
        }
        if (refunded - sumTotal != 0)
        {
            await AdjustBalanceAsync(db, inv.CustomerID, refunded - sumTotal);
        }
        Audit.LogAction(db, user, "SALES_RETURN", "SalesReturns", row.ReturnNumber,
                        "Invoice=" + inv.SalesInvoiceID + " Total=" + InvoiceCalculator.Plain(sumTotal) + (clientRef is null ? "" : " pos " + clientRef));
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        // the e-invoice of the credit note: PENDING, sent by the office tasks (docs/84, docs/85)
        return new SaleResult(SaveStatus.Saved, "", row.SalesReturnID, row.ReturnNumber, sumTotal, refunded);
    }

    /// <summary>The sales returns of a period (frmSalesReturn, the list of credit notes).</summary>
    public async Task<List<ReturnRow>> SalesReturnsAsync(DateTime from, DateTime to, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var rows = await (from r in db.SalesReturns
                          join c in db.Customers on r.CustomerID equals c.CustomerID
                          join h in db.SalesInvoices on r.SalesInvoiceID equals h.SalesInvoiceID
                          where r.ReturnDate >= start && r.ReturnDate < end
                          orderby r.ReturnDate descending
                          select new { r.SalesReturnID, r.ReturnNumber, r.ReturnDate, c.CustomerName, c.CustomerNameEn, h.InvoiceNumber, r.TotalAmount, r.RefundedAmount })
            .Take(take).ToListAsync();
        return rows.Select(r => new ReturnRow(r.SalesReturnID, r.ReturnNumber, r.ReturnDate, Name(r.CustomerName, r.CustomerNameEn, english), r.InvoiceNumber,
                                              r.TotalAmount, r.RefundedAmount)).ToList();
    }

    /// <summary>modSales.SetZatcaFields for a credit note (type 381, the subtype of its invoice).</summary>
    private static async Task SetReturnZatcaFieldsAsync(RetailDbContext db, Settings settings, SalesReturns row, string subType, DateTime when)
    {
        row.InvoiceSubType = subType;
        row.InvoiceTypeCode = "381";
        row.InvoiceUUID = Guid.NewGuid().ToString().ToLowerInvariant();
        row.ICV = int.Parse(await NumberSequences.NextNumberAsync(db, "ZATCA_ICV"));
        row.ZatcaStatus = settings.EInvoiceEnabled || settings.ZatcaPhase == 2 ? "PENDING" : "NOT_SENT";
        string vat = settings.VATNumber ?? "";
        if (vat.Length > 0 && CountryRules.Normalize(settings.CountryCode) == "SA")
        {
            row.QRCodeData = ZatcaQr.Build(settings.StoreName ?? "", vat, when, row.TotalAmount, row.Tax);
        }
    }

    // ------------------------------------------------------------------ the receipt voucher

    /// <summary>modSales.PostCustomerPayment: confirmOverBalance = the user answered yes to the question on an amount over
    /// the balance; clientRef = its number on a POS device (docs/86).</summary>
    public async Task<SaleResult> PostPaymentAsync(ActingUser user, PaymentRequest pay, bool confirmOverBalance = false,
                                                   string? clientRef = null)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, "frmCustomerPayment");
        if (open.Length > 0)
        {
            return SaleResult.Refused(open);
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        decimal amount = pay.Amount;
        if (amount <= 0)
        {
            return SaleResult.Refused("المبلغ يجب أن يكون أكبر من صفر.");
        }
        if (pay.CustomerID == settings.DefaultCustomerID)
        {
            return SaleResult.Refused("لا تُسجَّل دفعات على العميل النقدي.");
        }
        string baseCurrency = string.IsNullOrWhiteSpace(settings.CurrencyCode) ? "SAR" : settings.CurrencyCode;
        string code = string.IsNullOrWhiteSpace(pay.CurrencyCode) ? baseCurrency : pay.CurrencyCode.Trim();
        decimal rate = string.Equals(code, baseCurrency, StringComparison.OrdinalIgnoreCase) ? 1 : pay.ExchangeRate;
        string msg = await CurrencyProblemAsync(db, code, rate, baseCurrency);
        if (msg.Length > 0)
        {
            return SaleResult.Refused(msg);
        }
        var customer = await db.Customers.AsNoTracking().FirstOrDefaultAsync(c => c.CustomerID == pay.CustomerID);
        if (customer is null)
        {
            return SaleResult.Refused("العميل غير موجود.");
        }
        decimal foreign = amount;
        amount = MoneyMath.RoundMoney(foreign * rate);
        decimal balance = customer.CurrentBalance;
        if (amount > balance && !confirmOverBalance)
        {
            return new SaleResult(SaveStatus.NeedsConfirmation,
                                  "المبلغ (" + InvoiceCalculator.Money(amount) + ") أكبر من الرصيد المستحق (" + InvoiceCalculator.Money(balance) +
                                  ")، وسيصبح للعميل رصيد دائن.\n" + "هل تريد المتابعة؟");
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        var p = new CustomerPayments
        {
            PaymentNumber = await NumberSequences.NextNumberAsync(db, "CUSTOMER_PAYMENT"),
            CustomerID = pay.CustomerID, PaymentDate = DateTime.Now, Amount = amount, CurrencyCode = code, ExchangeRate = rate,
            ForeignAmount = foreign, SalesRepID = user.RepDay?.SalesRepID ?? await RepService.SalesRepForAsync(db, pay.CustomerID, user.EmployeeID),
            PaymentMethodID = pay.PaymentMethodID, CashBoxID = await CashBoxForAsync(db, user, pay.PaymentMethodID, amount),
            BankID = await BankForAsync(db, settings, pay.PaymentMethodID, amount), EmployeeID = user.EmployeeID,
            Notes = Clip(pay.Notes, 255), ClientRef = clientRef,
        };
        db.CustomerPayments.Add(p);
        await db.SaveChangesAsync();
        await AdjustBalanceAsync(db, pay.CustomerID, -amount);
        Audit.LogAction(db, user, "CUSTOMER_PAYMENT", "CustomerPayments", p.PaymentNumber, "Amount=" + InvoiceCalculator.Plain(amount));
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        return new SaleResult(SaveStatus.Saved, "", p.PaymentID, p.PaymentNumber, amount);
    }

    /// <summary>modCurrency.CurrencyProblem.</summary>
    internal static async Task<string> CurrencyProblemAsync(RetailDbContext db, string code, decimal rate, string baseCurrency)
    {
        if (code.Length == 0)
        {
            return "اختر العملة.";
        }
        if (!await db.Currencies.AnyAsync(c => c.IsActive && c.CurrencyCode == code))
        {
            return "العملة " + code + " غير موجودة أو غير نشطة.";
        }
        if (string.Equals(code, baseCurrency, StringComparison.OrdinalIgnoreCase))
        {
            return rate != 1 ? "معامل عملة البرنامج " + baseCurrency + " يساوي 1." : "";
        }
        return rate <= 0 ? "اكتب معامل التحويل للعملة " + code + " (قيمة وحدة واحدة بـ " + baseCurrency + ")." : "";
    }

    // ------------------------------------------------------------------ what the screens read

    internal static string Name(string arabic, string? english, bool inEnglish) =>
        inEnglish && !string.IsNullOrWhiteSpace(english) ? english : arabic;

    /// <summary>modProductUnits.FindProductCode, then the active products whose name or code contains the text.</summary>
    public async Task<List<SaleProduct>> FindProductsAsync(int employeeId, string text, int customerId, bool english, int take = 20)
    {
        string t = (text ?? "").Trim();
        if (t.Length == 0)
        {
            return [];
        }
        await using var db = await factory.CreateDbContextAsync();
        int repId = await RepService.RepIdOfAsync(db, employeeId);
        int warehouse = await RepService.UserWarehouseAsync(db, employeeId);
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
        var list = new List<SaleProduct>();
        foreach (var (p, unitId) in found)
        {
            int u = unitId ?? p.UnitID;
            string unitName = units.TryGetValue(u, out var un) ? Name(un.UnitName, un.UnitNameEn, english) : "";
            list.Add(new SaleProduct(p.ProductID, p.ProductCode, p.Barcode, Name(p.ProductName, p.ProductNameEn, english), unitId, unitName,
                                     await UnitSalePriceAsync(db, p, unitId, customerId, repId),
                                     await AvailableQtyAsync(db, p.ProductID, warehouse), p.TrackStock, p.VATCategory));
        }
        return list;
    }

    /// <summary>The units of a product: its smallest unit, then its active units (ProductUnits).</summary>
    public async Task<List<UnitChoice>> UnitsAsync(int productId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var p = await db.Products.AsNoTracking().FirstOrDefaultAsync(x => x.ProductID == productId);
        if (p is null)
        {
            return [];
        }
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        string N(int id) => units.TryGetValue(id, out var u) ? Name(u.UnitName, u.UnitNameEn, english) : "";
        var list = new List<UnitChoice> { new(p.UnitID, N(p.UnitID), 1) };
        list.AddRange((await db.ProductUnits.AsNoTracking().Where(u => u.IsActive && u.ProductID == productId).OrderBy(u => u.UnitFactor).ToListAsync())
                      .Select(u => new UnitChoice(u.UnitID, N(u.UnitID), u.UnitFactor)));
        return list;
    }

    /// <summary>The settings a screen needs to compute the invoice as it will be saved.</summary>
    public async Task<(decimal VatRate, bool PricesIncludeVat, int DefaultCustomerID, string Currency)> SaleSettingsAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        return (s.VATRate, s.PricesIncludeVAT, s.DefaultCustomerID, string.IsNullOrWhiteSpace(s.CurrencyCode) ? "SAR" : s.CurrencyCode);
    }

    public async Task<Dictionary<int, string>> VatCategoriesAsync(IEnumerable<int> productIds)
    {
        var ids = productIds.Distinct().ToList();
        await using var db = await factory.CreateDbContextAsync();
        return await db.Products.AsNoTracking().Where(p => ids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID, p => p.VATCategory);
    }

    /// <summary>The invoice as it is printed (qrySalesDocPrint): the customer, the lines in their unit, the QR.</summary>
    public async Task<InvoiceView?> InvoiceAsync(int invoiceId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await db.SalesInvoices.AsNoTracking().FirstOrDefaultAsync(x => x.SalesInvoiceID == invoiceId);
        if (h is null)
        {
            return null;
        }
        var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == h.CustomerID);
        var m = h.PaymentMethodID is { } mid ? await db.PaymentMethods.AsNoTracking().FirstOrDefaultAsync(x => x.PaymentMethodID == mid) : null;
        string emp = await db.Employees.Where(e => e.EmployeeID == h.EmployeeID).Select(e => e.EmployeeName).FirstOrDefaultAsync() ?? "";
        var w = await db.Warehouses.AsNoTracking().FirstOrDefaultAsync(x => x.WarehouseID == h.WarehouseID);
        var r = h.SalesRepID is { } rid ? await db.SalesReps.AsNoTracking().FirstOrDefaultAsync(x => x.SalesRepID == rid) : null;
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        var lines = await (from d in db.SalesInvoiceDetails.AsNoTracking()
                           join p in db.Products on d.ProductID equals p.ProductID
                           where d.SalesInvoiceID == invoiceId
                           orderby d.LineNumber
                           select new { d, p.ProductCode, p.ProductName, p.ProductNameEn, BaseUnit = p.UnitID }).ToListAsync();
        return new InvoiceView(h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate, h.CustomerID, Name(c.CustomerName, c.CustomerNameEn, english),
                               c.VATNumber, h.PaymentType, m is null ? null : Name(m.MethodName, m.MethodNameEn, english), h.SubTotal,
                               h.Discount, h.TaxableAmount, h.Tax, h.TotalAmount, h.PaidAmount, h.RemainingAmount, h.AmountTendered,
                               h.ChangeDue, h.DueDate, h.Notes, h.InvoiceSubType, h.InvoiceUUID, h.ICV, h.QRCodeData, h.ZatcaStatus, emp,
                               w is null ? null : Name(w.WarehouseName, w.WarehouseNameEn, english),
                               r is null ? null : Name(r.RepName, r.RepNameEn, english),
                               lines.Select(x =>
                               {
                                   int u = x.d.UnitID ?? x.BaseUnit;
                                   string un = units.TryGetValue(u, out var uu) ? Name(uu.UnitName, uu.UnitNameEn, english) : "";
                                   return new InvoiceLineView(x.d.LineNumber, x.d.ProductID, x.ProductCode, Name(x.ProductName, x.ProductNameEn, english),
                                                              un, x.d.Quantity, x.d.UnitPrice, x.d.Discount, x.d.NetAmount, x.d.VATRate, x.d.Tax,
                                                              x.d.LineTotal, x.d.LineNote);
                               }).ToList());
    }

    /// <summary>The invoices of a period, newest first (a search in the number or the customer's name).</summary>
    public async Task<List<InvoiceRow>> InvoicesAsync(DateTime from, DateTime to, string? search, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        string s = (search ?? "").Trim();
        var q = from h in db.SalesInvoices.AsNoTracking()
                join c in db.Customers on h.CustomerID equals c.CustomerID
                join e in db.Employees on h.EmployeeID equals e.EmployeeID
                where h.InvoiceDate >= start && h.InvoiceDate < end
                select new { h, c.CustomerName, c.CustomerNameEn, e.EmployeeName };
        if (s.Length > 0)
        {
            q = q.Where(x => x.h.InvoiceNumber.Contains(s) || x.CustomerName.Contains(s) || (x.CustomerNameEn != null && x.CustomerNameEn.Contains(s)));
        }
        var rows = await q.OrderByDescending(x => x.h.InvoiceDate).ThenByDescending(x => x.h.SalesInvoiceID).Take(take).ToListAsync();
        return rows.Select(x => new InvoiceRow(x.h.SalesInvoiceID, x.h.InvoiceNumber, x.h.InvoiceDate, Name(x.CustomerName, x.CustomerNameEn, english),
                                               x.h.PaymentType, x.h.TotalAmount, x.h.PaidAmount, x.h.RemainingAmount, x.h.ZatcaStatus,
                                               x.EmployeeName)).ToList();
    }

    /// <summary>The receipt vouchers of a period, newest first.</summary>
    public async Task<List<PaymentRow>> PaymentsAsync(DateTime from, DateTime to, string? search, bool english, int take = 200, int paymentId = 0)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        string s = (search ?? "").Trim();
        var q = from p in db.CustomerPayments.AsNoTracking()
                join c in db.Customers on p.CustomerID equals c.CustomerID
                join e in db.Employees on p.EmployeeID equals e.EmployeeID
                join m in db.PaymentMethods on p.PaymentMethodID equals m.PaymentMethodID
                where paymentId != 0 ? p.PaymentID == paymentId : p.PaymentDate >= start && p.PaymentDate < end
                select new { p, c.CustomerName, c.CustomerNameEn, e.EmployeeName, m.MethodName, m.MethodNameEn };
        if (s.Length > 0)
        {
            q = q.Where(x => x.p.PaymentNumber.Contains(s) || x.CustomerName.Contains(s) || (x.CustomerNameEn != null && x.CustomerNameEn.Contains(s)));
        }
        var rows = await q.OrderByDescending(x => x.p.PaymentDate).ThenByDescending(x => x.p.PaymentID).Take(take).ToListAsync();
        return rows.Select(x => new PaymentRow(x.p.PaymentID, x.p.PaymentNumber, x.p.PaymentDate, x.p.CustomerID,
                                               Name(x.CustomerName, x.CustomerNameEn, english), x.p.Amount, x.p.CurrencyCode, x.p.ForeignAmount,
                                               Name(x.MethodName, x.MethodNameEn, english), x.p.Notes, x.EmployeeName)).ToList();
    }

    /// <summary>The active customers for the choice lists of the screens (the cash customer first).</summary>
    public async Task<List<(int Id, string Name, decimal Balance, bool AllowCredit)>> CustomersAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        int cash = (await SettingsReader.ReadAsync(db)).DefaultCustomerID;
        var rows = await db.Customers.AsNoTracking().Where(c => c.IsActive).ToListAsync();
        return rows.OrderBy(c => c.CustomerID == cash ? 0 : 1).ThenBy(c => c.CustomerName)
                   .Select(c => (c.CustomerID, Name(c.CustomerName, c.CustomerNameEn, english), c.CurrentBalance, c.AllowCredit)).ToList();
    }

    /// <summary>The active payment methods in their order.</summary>
    public async Task<List<(int Id, string Name)>> PaymentMethodsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.PaymentMethods.AsNoTracking().Where(m => m.IsActive).OrderBy(m => m.SortOrder).ThenBy(m => m.PaymentMethodID).ToListAsync();
        return rows.Select(m => (m.PaymentMethodID, Name(m.MethodName, m.MethodNameEn, english))).ToList();
    }
}
