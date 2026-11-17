using System.Text.RegularExpressions;
using Microsoft.EntityFrameworkCore;
using RetailStore.Domain;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

// The point of sale on Windows (plan 6G, docs/83): what the POS keeps on the PC to sell without the database (the snapshot)
// and a sale made on the PC, sent once (its number on the device). The same records travel through the API (remote
// database) and are read directly (local database), so both ways give the POS the same data.

/// <summary>The settings a sale needs (modSales, modCountry).</summary>
public sealed record PosSettings(string StoreName, string? VATNumber, string CountryCode, decimal VATRate, bool PricesIncludeVAT,
                                 int DefaultCustomerID, string Currency, bool AllowNegativeStock, int CreditBlockDays,
                                 string InvoicePrefix, string? ReceiptFooter, ScaleSettings? Scale = null);

/// <summary>The user at the POS: his permissions as the sale checks them (PostSaleFromCart), his limits, warehouse and box.
/// MaxDiscountPercent = the limit that applies (the rep's when the user is a rep with a limit). ReturnProblem / PaymentProblem
/// / ClosingProblem = why he may not make a return, a receipt or a closing ("" = he may; docs/86).</summary>
public sealed record PosUser(int EmployeeID, string Username, string EmployeeName, bool IsAdmin, string SellProblem, bool PriceOverride,
                             bool DiscountOverride, bool AllowNegativeStock, decimal MaxDiscountPercent, int RepID, int WarehouseID,
                             string WarehouseName, int CashBoxID, string ReturnProblem = "", string PaymentProblem = "",
                             string ClosingProblem = "");

/// <summary>A bigger unit of a product: its barcode, its own price (0 = piece price x factor) and the price of a new line
/// without a customer.</summary>
public sealed record PosUnit(int UnitID, string UnitName, decimal Factor, string? Barcode, decimal OwnPrice, decimal Price);

/// <summary>An active product: Price = the price of a new line without a customer (the user's rep price or the selling
/// price); Available = the stock of the user's warehouse less Reserved, what the open orders reserve.</summary>
public sealed record PosProduct(int ProductID, string ProductCode, string? Barcode, string ProductName, int UnitID, string UnitName,
                                string VATCategory, bool TrackStock, decimal SellingPrice, decimal? SizePriceM, decimal? SizePriceL,
                                decimal? RepPrice, decimal Price, decimal Available, decimal Reserved, PosUnit[] Units);

/// <summary>An active customer: DaysLate = his oldest late document (only when CreditBlockDays is set); PriceGroup = his
/// price list ("" = none); RepID = the rep of his sales (his own active rep, else the rep of the user).</summary>
public sealed record PosCustomer(int CustomerID, string CustomerName, string? VATNumber, string? Mobile, decimal CurrentBalance,
                                 bool AllowCredit, decimal CreditLimit, int PaymentTermsDays, int DaysLate, string PriceGroup, int? RepID);

public sealed record PosRep(int SalesRepID, string RepName, decimal CreditLimit, decimal CustomersBalance);

public sealed record PosMethod(int PaymentMethodID, string MethodName);

/// <summary>Everything the POS keeps to sell without the database. PriceGroups: for each price list, the prices that
/// differ from the price without a customer, keyed "product:unit" (unit 0 = the smallest unit).</summary>
public sealed record PosSnapshot(DateTime ServerTime, PosSettings Settings, PosUser User, List<PosProduct> Products,
                                 List<PosCustomer> Customers, Dictionary<string, Dictionary<string, decimal>> PriceGroups,
                                 List<PosRep> Reps, List<PosMethod> Methods);

/// <summary>A sale made on the POS: its number on the device (ClientRef), the invoice number it printed, the time, and
/// ConfirmStock = the user answered yes to the question on the stock (or the settings allow it).</summary>
public sealed record DeviceSaleRequest(string ClientRef, string InvoiceNumber, DateTime TakenAt, bool ConfirmStock, int CustomerID,
                                       string PaymentType, int? PaymentMethodID, decimal InvoiceDiscount, decimal? Tendered,
                                       string? Notes, CartLine[] Lines)
{
    public SaleRequest ToSale() => new(CustomerID, PaymentType, PaymentMethodID, InvoiceDiscount, Tendered, Notes, Lines ?? []);
}

/// <summary>The answer to a device sale: Duplicate = it was already on the server (sent before), nothing added.</summary>
public sealed record DeviceSaleResult(SaveStatus Status, string Message, int InvoiceID, string InvoiceNumber, bool Duplicate);

/// <summary>A sales return made on the POS (docs/86): ClientRef = its number on the device, so a return sent again after a
/// lost answer is added once.</summary>
public sealed record DeviceReturnRequest(string ClientRef, int SalesInvoiceID, string? Reason, string RefundType, int? PaymentMethodID,
                                         ReturnLineRequest[] Lines)
{
    public SalesReturnRequest ToReturn() => new(SalesInvoiceID, Reason, RefundType, PaymentMethodID, Lines ?? []);
}

/// <summary>A customer receipt made on the POS: ConfirmOverBalance = the answer yes to the question on an amount over the
/// balance.</summary>
public sealed record DevicePaymentRequest(string ClientRef, int CustomerID, decimal Amount, int PaymentMethodID, string? Notes,
                                          bool ConfirmOverBalance = false);

/// <summary>A cashier closing made on the POS.</summary>
public sealed record DeviceClosingRequest(string ClientRef, int CashBoxID, decimal Counted, string Destination, int? ToCashBoxID,
                                          decimal TransferAmount, string? Notes)
{
    public CashClosingRequest ToClosing() => new(CashBoxID, Counted, Destination, ToCashBoxID, TransferAmount, Notes);
}

/// <summary>A line of a printed document: the product (its unit) and the quantity, and the amount.</summary>
public sealed record PosSlipLine(string Text, decimal Quantity, decimal Amount);

/// <summary>A row of the totals of a printed document (its Arabic label, translated when printed).</summary>
public sealed record PosSlipRow(string Label, decimal Amount);

/// <summary>A document the POS made on the program and prints (docs/86): Kind RETURN / PAYMENT / CLOSING, its number, date,
/// title, who made it, the customer or the box, its lines and totals, Note (the invoice of a return, the method of a receipt)
/// and the QR of a Saudi credit note.</summary>
public sealed record PosDocument(string Kind, int Id, string Number, DateTime Date, string Title, string CashierName, string PartyName,
                                 PosSlipLine[] Lines, PosSlipRow[] Rows, string? Note, string? Qr);

/// <summary>The answer to a document of the POS: Duplicate = it was already on the server (sent before), nothing added.</summary>
public sealed record DeviceDocResult(SaveStatus Status, string Message, bool Duplicate, PosDocument? Document);

/// <summary>The boxes the user may close, the main safes, and his own box (ClosingLoad).</summary>
public sealed record PosClosingBoxes(List<ClosingBox> Boxes, List<ClosingBox> Safes, int DefaultBoxID);

/// <summary>The POS on Windows (docs/83): the snapshot and the device sale, used by the API (/api/v1/pos) and directly by a
/// POS connected to the database of the store.</summary>
public sealed class PosService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, SalesService sales)
{
    public const string DocRefUsed = "رقم المستند في الجهاز مستعمل لمستخدم آخر.";
    public const string ClientRefMissing = "رقم الفاتورة في الجهاز مطلوب.";
    public const string ClientRefTooLong = "رقم الفاتورة في الجهاز أطول من 40 حرفًا.";
    public const string ClientRefUsed = "رقم الفاتورة في الجهاز مستعمل لمستخدم آخر.";
    public const string BadNumber = "رقم الفاتورة من الجهاز غير صالح: رمز الجهاز ثم شرطة ثم أرقام.";

    private static readonly Regex NumberPattern = new("^[A-Za-z0-9]{1,8}-[0-9]{1,10}$", RegexOptions.CultureInvariant);

    /// <summary>A device code: one to eight English letters or digits, not the prefix of the invoices of the program.</summary>
    public static string DeviceCodeProblem(string? code, string invoicePrefix)
    {
        string c = (code ?? "").Trim();
        if (!Regex.IsMatch(c, "^[A-Za-z0-9]{1,8}$", RegexOptions.CultureInvariant))
        {
            return "رمز الجهاز من حرف إلى ثمانية حروف إنجليزية أو أرقام، مثل POS1.";
        }
        return string.Equals(c + "-", invoicePrefix, StringComparison.OrdinalIgnoreCase)
            ? "رمز الجهاز هو بادئة فواتير البرنامج. اختر رمزًا آخر."
            : "";
    }

    /// <summary>The invoice number a device gives its n-th sale.</summary>
    public static string DeviceNumber(string code, long n) => code.Trim().ToUpperInvariant() + "-" + n.ToString("000000");

    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    // ------------------------------------------------------------------ the user

    public async Task<PosUser?> UserAsync(int employeeId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await UserAsync(db, employeeId, english);
    }

    private async Task<PosUser?> UserAsync(RetailDbContext db, int employeeId, bool english)
    {
        var e = await db.Employees.AsNoTracking().FirstOrDefaultAsync(x => x.EmployeeID == employeeId && x.IsActive);
        if (e is null)
        {
            return null;
        }
        int repId = await RepService.RepIdOfAsync(db, employeeId);
        decimal maxDiscount = e.MaxDiscountPercent;
        if (repId != 0)
        {
            decimal lim = await db.SalesReps.Where(r => r.SalesRepID == repId).Select(r => r.MaxDiscountPercent).FirstOrDefaultAsync();
            if (lim > 0)
            {
                maxDiscount = lim;
            }
        }
        int warehouse = await RepService.UserWarehouseAsync(db, employeeId);
        var w = await db.Warehouses.AsNoTracking().FirstOrDefaultAsync(x => x.WarehouseID == warehouse);
        string returnProblem = await AddProblemAsync(employeeId, "frmSalesReturn");
        string closingProblem = await AddProblemAsync(employeeId, TreasuryService.ClosingForm);
        if (closingProblem.Length == 0 && !await permissions.HasPermissionAsync(employeeId, "CASH_CLOSING"))
        {
            closingProblem = "ليست لديك صلاحية تصفية يومية الكاشير.";
        }
        return new PosUser(employeeId, e.Username ?? "", e.EmployeeName, e.IsDeveloper || e.RoleID == AppData.AdminRoleId,
                           await permissions.OpenProblemAsync(employeeId, "frmPOS"),
                           await permissions.HasPermissionAsync(employeeId, "PRICE_OVERRIDE"),
                           await permissions.HasPermissionAsync(employeeId, "DISCOUNT_OVERRIDE"),
                           await permissions.HasPermissionAsync(employeeId, "ALLOW_NEGATIVE_STOCK"), maxDiscount, repId, warehouse,
                           w is null ? "" : Name(w.WarehouseName, w.WarehouseNameEn, english),
                           await sales.CurrentCashBoxAsync(db, employeeId), returnProblem,
                           await permissions.OpenProblemAsync(employeeId, "frmCustomerPayment"), closingProblem);
    }

    /// <summary>The screen opens and its ADD action is allowed (modSecurity.CanOpenScreen + CanScreenAction).</summary>
    private async Task<string> AddProblemAsync(int employeeId, string form)
    {
        string open = await permissions.OpenProblemAsync(employeeId, form);
        return open.Length > 0 ? open : await permissions.ActionProblemAsync(employeeId, form, ScreenAction.Add);
    }

    // ------------------------------------------------------------------ the snapshot

    /// <summary>The snapshot of the user: the prices are those the sale accepts (UnitSalePrice), the stock that of his
    /// warehouse less the reservations (CheckStockAvailable).</summary>
    public async Task<PosSnapshot?> SnapshotAsync(int employeeId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var user = await UserAsync(db, employeeId, english);
        if (user is null)
        {
            return null;
        }
        var s = await SettingsReader.ReadAsync(db);
        string prefix = await db.Sequences.Where(q => q.SequenceName == "SALES_INVOICE").Select(q => q.Prefix).FirstOrDefaultAsync() ?? "";
        var settings = new PosSettings(s.StoreName ?? "", s.VATNumber, CountryRules.Normalize(s.CountryCode), s.VATRate, s.PricesIncludeVAT,
                                       s.DefaultCustomerID, string.IsNullOrWhiteSpace(s.CurrencyCode) ? "SAR" : s.CurrencyCode,
                                       s.AllowNegativeStock, s.CreditBlockDays ?? 0, prefix, s.ReceiptFooter,
                                       await ScaleSettingsAsync(db));
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        string UnitName(int id) => units.TryGetValue(id, out var u) ? Name(u.UnitName, u.UnitNameEn, english) : "";
        var stock = await RepService.WarehouseStockAsync(db, user.WarehouseID);
        var productUnits = (await db.ProductUnits.AsNoTracking().Where(u => u.IsActive).ToListAsync())
            .GroupBy(u => u.ProductID).ToDictionary(g => g.Key, g => g.OrderBy(u => u.UnitFactor).ToList());
        var repPrices = user.RepID == 0 ? []
            : await db.RepPrices.AsNoTracking().Where(r => r.SalesRepID == user.RepID).ToDictionaryAsync(r => r.ProductID, r => r.UnitPrice);
        var products = await db.Products.AsNoTracking().Where(p => p.IsActive).OrderBy(p => p.ProductName).ToListAsync();
        var lists = await (from c in db.Customers
                           join pl in db.PriceLists on c.PriceListID equals pl.PriceListID
                           where c.IsActive && pl.IsActive
                           group c by pl.PriceListID into g
                           select new { ListId = g.Key, CustomerID = g.Min(c => c.CustomerID) }).ToListAsync();
        var groups = lists.ToDictionary(l => "L" + l.ListId, _ => new Dictionary<string, decimal>());
        var list = new List<PosProduct>();
        foreach (var p in products)
        {
            decimal piece = await SalesService.UnitSalePriceAsync(db, p, null, 0, user.RepID);
            var own = productUnits.GetValueOrDefault(p.ProductID) ?? [];
            var unitRows = new List<PosUnit>();
            foreach (var u in own)
            {
                unitRows.Add(new PosUnit(u.UnitID, UnitName(u.UnitID), u.UnitFactor, u.Barcode, u.SellingPrice,
                                         await SalesService.UnitSalePriceAsync(db, p, u.UnitID, 0, user.RepID)));
            }
            foreach (var l in lists)
            {
                var prices = groups["L" + l.ListId];
                decimal forList = await SalesService.UnitSalePriceAsync(db, p, null, l.CustomerID, user.RepID);
                if (forList != piece)
                {
                    prices[p.ProductID + ":0"] = forList;
                }
                foreach (var u in unitRows)
                {
                    decimal unitForList = await SalesService.UnitSalePriceAsync(db, p, u.UnitID, l.CustomerID, user.RepID);
                    if (unitForList != u.Price)
                    {
                        prices[p.ProductID + ":" + u.UnitID] = unitForList;
                    }
                }
            }
            decimal reserved = await SalesService.ReservedQtyAsync(db, p.ProductID, user.WarehouseID);
            decimal available = stock.GetValueOrDefault(p.ProductID) - reserved;
            list.Add(new PosProduct(p.ProductID, p.ProductCode, p.Barcode, Name(p.ProductName, p.ProductNameEn, english), p.UnitID,
                                    UnitName(p.UnitID), p.VATCategory, p.TrackStock, p.SellingPrice, p.SizePriceM, p.SizePriceL,
                                    repPrices.TryGetValue(p.ProductID, out var rp) ? rp : null, piece, available, reserved,
                                    unitRows.ToArray()));
        }

        var activeReps = await db.SalesReps.AsNoTracking().Where(r => r.IsActive).ToListAsync();
        var activeRepIds = activeReps.Select(r => r.SalesRepID).ToHashSet();
        var customers = new List<PosCustomer>();
        foreach (var c in await db.Customers.AsNoTracking().Where(c => c.IsActive).OrderBy(c => c.CustomerName).ToListAsync())
        {
            int late = 0;
            if (settings.CreditBlockDays > 0 && c.AllowCredit
                && (c.OpeningBalance > 0 || await db.SalesInvoices.AnyAsync(h => h.CustomerID == c.CustomerID && h.RemainingAmount > 0)))
            {
                late = CustomerAging.MaxDaysLate(await SalesService.CustomerAgingAsync(db, c.CustomerID, DateTime.Today), DateTime.Today);
            }
            string group = c.PriceListID is { } pl && groups.ContainsKey("L" + pl) ? "L" + pl : "";
            int? rep = c.SalesRepID is { } r && activeRepIds.Contains(r) ? r : (user.RepID == 0 ? null : user.RepID);
            customers.Add(new PosCustomer(c.CustomerID, Name(c.CustomerName, c.CustomerNameEn, english), c.VATNumber, c.Mobile,
                                          c.CurrentBalance, c.AllowCredit, c.CreditLimit, c.PaymentTermsDays ?? 0, late, group, rep));
        }
        var reps = new List<PosRep>();
        foreach (var r in activeReps.Where(r => r.CreditLimit > 0))
        {
            reps.Add(new PosRep(r.SalesRepID, Name(r.RepName, r.RepNameEn, english), r.CreditLimit,
                                await RepService.CustomersBalanceAsync(db, r.SalesRepID)));
        }
        var methods = (await sales.PaymentMethodsAsync(english)).Select(m => new PosMethod(m.Id, m.Name)).ToList();
        return new PosSnapshot(DateTime.Now, settings, user, list, customers, groups, reps, methods);
    }

    // ------------------------------------------------------------------ the sale of a device

    /// <summary>A sale made on the POS (modSales.PostSaleFromCart with the device number and time): a sale already sent by
    /// the same user answers with its invoice (nothing added), a number of another sale or user is refused.</summary>
    public async Task<DeviceSaleResult> PostDeviceSaleAsync(ActingUser user, DeviceSaleRequest request)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string number = (request.InvoiceNumber ?? "").Trim().ToUpperInvariant();
        if (reference.Length == 0)
        {
            return new DeviceSaleResult(SaveStatus.Refused, ClientRefMissing, 0, "", false);
        }
        if (reference.Length > 40)
        {
            return new DeviceSaleResult(SaveStatus.Refused, ClientRefTooLong, 0, "", false);
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.SalesInvoices.AsNoTracking().Where(h => h.ClientRef == reference)
                .Select(h => new { h.SalesInvoiceID, h.InvoiceNumber, h.EmployeeID }).FirstOrDefaultAsync();
            if (sent is not null)
            {
                return sent.EmployeeID == user.EmployeeID
                    ? new DeviceSaleResult(SaveStatus.Saved, "", sent.SalesInvoiceID, sent.InvoiceNumber, true)
                    : new DeviceSaleResult(SaveStatus.Refused, ClientRefUsed, 0, "", false);
            }
            string prefix = await db.Sequences.Where(q => q.SequenceName == "SALES_INVOICE").Select(q => q.Prefix).FirstOrDefaultAsync() ?? "";
            if (!NumberPattern.IsMatch(number) || DeviceCodeProblem(number[..number.IndexOf('-')], prefix).Length > 0)
            {
                return new DeviceSaleResult(SaveStatus.Refused, BadNumber, 0, "", false);
            }
            if (await db.SalesInvoices.AnyAsync(h => h.InvoiceNumber == number))
            {
                return new DeviceSaleResult(SaveStatus.Refused, "رقم الفاتورة «" + number + "» مستعمل لفاتورة أخرى. " +
                                                                "اجعل لكل جهاز رمزًا مختلفًا في إعدادات قاعدة البيانات.", 0, "", false);
            }
        }
        var r = await sales.PostSaleAsync(user, request.ToSale(), request.ConfirmStock,
                                          new DeviceSaleInfo(reference, number, request.TakenAt));
        return new DeviceSaleResult(r.Status, r.Message, r.InvoiceID, r.InvoiceNumber, false);
    }

    // ------------------------------------------------------------------ returns, receipts and closings of a device (docs/86)

    /// <summary>The scale barcodes of the settings (LabelSettings, modScale).</summary>
    public static async Task<ScaleSettings> ScaleSettingsAsync(RetailDbContext db) =>
        await db.LabelSettings.AsNoTracking().Where(l => l.LabelSettingID == 1)
            .Select(l => new ScaleSettings(l.ScalePrefixes, l.ScaleItemDigits, l.ScaleValueKind)).FirstOrDefaultAsync() ?? ScaleSettings.None;

    private static string RefProblem(string reference) =>
        reference.Length == 0 ? "رقم المستند في الجهاز مطلوب." : reference.Length > 40 ? "رقم المستند في الجهاز أطول من 40 حرفًا." : "";

    /// <summary>A document already on the server with this number of the device: its id for the same user (sent again), a
    /// refusal for another user, null = new.</summary>
    private static DeviceDocResult? Sent(int? id, int employeeId, int userId) =>
        id is null ? null : employeeId == userId ? new DeviceDocResult(SaveStatus.Saved, "", true, null) : new DeviceDocResult(SaveStatus.Refused, DocRefUsed, false, null);

    /// <summary>A sales return of the POS (modSales.PostSalesReturn), sent once.</summary>
    public async Task<DeviceDocResult> PostDeviceReturnAsync(ActingUser user, DeviceReturnRequest request, bool english)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string problem = RefProblem(reference);
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.SalesReturns.AsNoTracking().Where(r => r.ClientRef == reference)
                .Select(r => new { r.SalesReturnID, r.EmployeeID }).FirstOrDefaultAsync();
            if (Sent(sent?.SalesReturnID, sent?.EmployeeID ?? 0, user.EmployeeID) is { } done)
            {
                return done with { Document = done.Status == SaveStatus.Saved ? await ReturnDocumentAsync(sent!.SalesReturnID, english) : null };
            }
        }
        var r = await sales.PostSalesReturnAsync(user, request.ToReturn(), reference);
        return r.Status == SaveStatus.Saved
            ? new DeviceDocResult(SaveStatus.Saved, "", false, await ReturnDocumentAsync(r.InvoiceID, english))
            : new DeviceDocResult(r.Status, r.Message, false, null);
    }

    /// <summary>A customer receipt of the POS (modSales.PostCustomerPayment), sent once.</summary>
    public async Task<DeviceDocResult> PostDevicePaymentAsync(ActingUser user, DevicePaymentRequest request, bool english)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string problem = RefProblem(reference);
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.CustomerPayments.AsNoTracking().Where(p => p.ClientRef == reference)
                .Select(p => new { p.PaymentID, p.EmployeeID }).FirstOrDefaultAsync();
            if (Sent(sent?.PaymentID, sent?.EmployeeID ?? 0, user.EmployeeID) is { } done)
            {
                return done with { Document = done.Status == SaveStatus.Saved ? await PaymentDocumentAsync(sent!.PaymentID, english) : null };
            }
        }
        var r = await sales.PostPaymentAsync(user, new PaymentRequest(request.CustomerID, request.Amount, request.PaymentMethodID, request.Notes),
                                             request.ConfirmOverBalance, reference);
        return r.Status == SaveStatus.Saved
            ? new DeviceDocResult(SaveStatus.Saved, "", false, await PaymentDocumentAsync(r.InvoiceID, english))
            : new DeviceDocResult(r.Status, r.Message, false, null);
    }

    /// <summary>A cashier closing of the POS (modCash.PostCashClosing), sent once.</summary>
    public async Task<DeviceDocResult> PostDeviceClosingAsync(ActingUser user, DeviceClosingRequest request, bool english)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string problem = RefProblem(reference);
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.CashClosings.AsNoTracking().Where(c => c.ClientRef == reference)
                .Select(c => new { c.ClosingID, c.EmployeeID }).FirstOrDefaultAsync();
            if (Sent(sent?.ClosingID, sent?.EmployeeID ?? 0, user.EmployeeID) is { } done)
            {
                return done with { Document = done.Status == SaveStatus.Saved ? await ClosingDocumentAsync(sent!.ClosingID, english) : null };
            }
        }
        var r = await new TreasuryService(factory, permissions).PostCashClosingAsync(user, request.ToClosing(), reference);
        return r.Status == SaveStatus.Saved
            ? new DeviceDocResult(SaveStatus.Saved, "", false, await ClosingDocumentAsync(r.Id, english))
            : new DeviceDocResult(r.Status, r.Message, false, null);
    }

    /// <summary>The boxes the user may close (ClosingLoad).</summary>
    public async Task<PosClosingBoxes> ClosingBoxesAsync(int employeeId, bool english)
    {
        var (boxes, safes, own) = await new TreasuryService(factory, permissions).ClosingBoxesAsync(employeeId, english);
        return new PosClosingBoxes(boxes, safes, own);
    }

    private static async Task<string> EmployeeNameAsync(RetailDbContext db, int id) =>
        await db.Employees.Where(e => e.EmployeeID == id).Select(e => e.EmployeeName).FirstOrDefaultAsync() ?? "";

    private static async Task<string> CustomerNameAsync(RetailDbContext db, int id, bool english) =>
        await db.Customers.Where(c => c.CustomerID == id).Select(c => english && c.CustomerNameEn != null && c.CustomerNameEn != "" ? c.CustomerNameEn : c.CustomerName)
            .FirstOrDefaultAsync() ?? "";

    /// <summary>The credit note as the POS prints it.</summary>
    public async Task<PosDocument?> ReturnDocumentAsync(int returnId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await db.SalesReturns.AsNoTracking().FirstOrDefaultAsync(x => x.SalesReturnID == returnId);
        if (r is null)
        {
            return null;
        }
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID, u => Name(u.UnitName, u.UnitNameEn, english));
        var rows = await (from d in db.SalesReturnDetails
                          join p in db.Products on d.ProductID equals p.ProductID
                          where d.SalesReturnID == returnId
                          orderby d.ReturnDetailID
                          select new { p.ProductName, p.ProductNameEn, Unit = d.UnitID ?? p.UnitID, d.Quantity, d.LineTotal }).ToListAsync();
        var lines = rows.Select(x => new PosSlipLine(Name(x.ProductName, x.ProductNameEn, english) + " - " + units.GetValueOrDefault(x.Unit, ""),
                                                     x.Quantity, x.LineTotal)).ToArray();
        string invoice = await db.SalesInvoices.Where(h => h.SalesInvoiceID == r.SalesInvoiceID).Select(h => h.InvoiceNumber).FirstAsync();
        var totals = new List<PosSlipRow>
        {
            new("الإجمالي قبل الضريبة", r.TaxableAmount), new("الضريبة", r.Tax), new("الإجمالي شامل الضريبة", r.TotalAmount),
            r.RefundedAmount > 0 ? new PosSlipRow("المبلغ المردود نقدًا", r.RefundedAmount) : new PosSlipRow("خصم من رصيد العميل", r.TotalAmount),
        };
        return new PosDocument("RETURN", r.SalesReturnID, r.ReturnNumber, r.ReturnDate, "إشعار دائن", await EmployeeNameAsync(db, r.EmployeeID),
                               await CustomerNameAsync(db, r.CustomerID, english), lines, totals.ToArray(), invoice + " - " + r.Reason, r.QRCodeData);
    }

    /// <summary>The receipt voucher as the POS prints it.</summary>
    public async Task<PosDocument?> PaymentDocumentAsync(int paymentId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var p = await db.CustomerPayments.AsNoTracking().FirstOrDefaultAsync(x => x.PaymentID == paymentId);
        if (p is null)
        {
            return null;
        }
        decimal balance = await db.Customers.Where(c => c.CustomerID == p.CustomerID).Select(c => c.CurrentBalance).FirstAsync();
        string? method = (await sales.PaymentMethodsAsync(english)).Where(m => m.Id == p.PaymentMethodID).Select(m => m.Name).FirstOrDefault();
        return new PosDocument("PAYMENT", p.PaymentID, p.PaymentNumber, p.PaymentDate, "سند قبض", await EmployeeNameAsync(db, p.EmployeeID),
                               await CustomerNameAsync(db, p.CustomerID, english), [],
                               [new PosSlipRow("المبلغ", p.Amount), new PosSlipRow("الرصيد المستحق", balance)], method, null);
    }

    /// <summary>The cashier closing as the POS prints it.</summary>
    public async Task<PosDocument?> ClosingDocumentAsync(int closingId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var c = await db.CashClosings.AsNoTracking().FirstOrDefaultAsync(x => x.ClosingID == closingId);
        if (c is null)
        {
            return null;
        }
        var box = await db.CashBoxes.AsNoTracking().FirstAsync(b => b.CashBoxID == c.CashBoxID);
        var rows = new List<PosSlipRow>
        {
            new("رصيد البداية", c.OpeningBalance), new("المقبوضات", c.CashIn), new("المدفوعات", c.CashOut), new("الرصيد الدفتري", c.ExpectedBalance),
            new("النقدية الفعلية", c.CountedAmount), new("الفرق", c.Difference), new("المبلغ المرحَّل", c.TransferAmount),
            new("المتبقي في الصندوق (عهدة)", c.KeptAmount),
        };
        return new PosDocument("CLOSING", c.ClosingID, c.ClosingNumber, c.ClosingDate, "تصفية يومية الكاشير", await EmployeeNameAsync(db, c.EmployeeID),
                               Name(box.BoxName, box.BoxNameEn, english), [], rows.ToArray(), c.Notes, null);
    }
}
