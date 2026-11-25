using Microsoft.EntityFrameworkCore;

namespace RetailStore.Data.Services;

/// <summary>A sales invoice made on the rep app (docs/87): ClientRef = its number on the phone, kept in SalesInvoices.ClientRef,
/// so an invoice sent again after a lost answer is added once; ConfirmStock = the answer yes to the question on the stock.</summary>
public sealed record RepSaleRequest(string ClientRef, int CustomerID, string PaymentType, int? PaymentMethodID, decimal InvoiceDiscount,
                                    string? Notes, CartLine[] Lines, bool ConfirmStock = false)
{
    public SaleRequest ToSale() => new(CustomerID, PaymentType, PaymentMethodID, InvoiceDiscount, null, Notes, Lines ?? []);
}

/// <summary>The invoice and the receipt of the rep app (plan 7C, docs/87): made on the program at once (never kept on the phone),
/// as the rep of his open rep day (modVanSales.StartActingAsRep): the van is the warehouse, the custody box takes the cash and he
/// is the rep of the document; then SalesService.PostSaleAsync / PostPaymentAsync with their checks and messages. Each carries
/// the ClientRef of the phone: sent again by the same user = the existing document.</summary>
public sealed class RepDocService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const string NoOpenDay = "لا يوجد يوم مفتوح لك: افتح يوم المندوب في البرنامج (شاشة يوم المندوب)، ثم أعد المحاولة.";

    private SalesService Sales => new(factory, permissions);
    private PosService Pos => new(factory, permissions, Sales);

    /// <summary>The open rep day of the rep (RepDays.Status OPEN, the latest), null = none.</summary>
    public static async Task<RepDayContext?> OpenDayAsync(RetailDbContext db, int repId) =>
        repId == 0 ? null
            : await db.RepDays.AsNoTracking().Where(d => d.SalesRepID == repId && d.Status == "OPEN")
                .OrderByDescending(d => d.OpenedAt).ThenByDescending(d => d.RepDayID)
                .Select(d => new RepDayContext(d.RepDayID, d.SalesRepID, d.WarehouseID, d.CashBoxID)).FirstOrDefaultAsync();

    private static string RefProblem(string reference) =>
        reference.Length == 0 ? "رقم المستند في الجهاز مطلوب." : reference.Length > 40 ? "رقم المستند في الجهاز أطول من 40 حرفًا." : "";

    /// <summary>The user acting as the rep of his open rep day, or the reason he cannot.</summary>
    private async Task<(string Problem, ActingUser User)> ActingAsync(RetailDbContext db, ActingUser user, int repId)
    {
        var day = await OpenDayAsync(db, repId);
        return day is null ? (NoOpenDay, user) : ("", user with { RepDay = day });
    }

    /// <summary>A sales invoice of the rep app, sent once.</summary>
    public async Task<DeviceDocResult> PostSaleAsync(ActingUser user, int repId, RepSaleRequest request, bool english)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string problem = RefProblem(reference);
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        ActingUser acting;
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.SalesInvoices.AsNoTracking().Where(s => s.ClientRef == reference)
                .Select(s => new { s.SalesInvoiceID, s.EmployeeID }).FirstOrDefaultAsync();
            if (sent is not null)
            {
                return sent.EmployeeID == user.EmployeeID
                    ? new DeviceDocResult(SaveStatus.Saved, "", true, await SaleDocumentAsync(sent.SalesInvoiceID, english))
                    : new DeviceDocResult(SaveStatus.Refused, PosService.DocRefUsed, false, null);
            }
            (problem, acting) = await ActingAsync(db, user, repId);
        }
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        var r = await Sales.PostSaleAsync(acting, request.ToSale(), request.ConfirmStock, clientRef: reference);
        return r.Status == SaveStatus.Saved
            ? new DeviceDocResult(SaveStatus.Saved, "", false, await SaleDocumentAsync(r.InvoiceID, english))
            : new DeviceDocResult(r.Status, r.Message, false, null);
    }

    /// <summary>A receipt of the rep app (into the custody box for cash), sent once.</summary>
    public async Task<DeviceDocResult> PostPaymentAsync(ActingUser user, int repId, DevicePaymentRequest request, bool english)
    {
        string reference = (request.ClientRef ?? "").Trim();
        string problem = RefProblem(reference);
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        ActingUser acting;
        await using (var db = await factory.CreateDbContextAsync())
        {
            var sent = await db.CustomerPayments.AsNoTracking().Where(p => p.ClientRef == reference)
                .Select(p => new { p.PaymentID, p.EmployeeID }).FirstOrDefaultAsync();
            if (sent is not null)
            {
                return sent.EmployeeID == user.EmployeeID
                    ? new DeviceDocResult(SaveStatus.Saved, "", true, await Pos.PaymentDocumentAsync(sent.PaymentID, english))
                    : new DeviceDocResult(SaveStatus.Refused, PosService.DocRefUsed, false, null);
            }
            (problem, acting) = await ActingAsync(db, user, repId);
        }
        if (problem.Length > 0)
        {
            return new DeviceDocResult(SaveStatus.Refused, problem, false, null);
        }
        var r = await Sales.PostPaymentAsync(acting, new PaymentRequest(request.CustomerID, request.Amount, request.PaymentMethodID, request.Notes),
                                             request.ConfirmOverBalance, reference);
        return r.Status == SaveStatus.Saved
            ? new DeviceDocResult(SaveStatus.Saved, "", false, await Pos.PaymentDocumentAsync(r.InvoiceID, english))
            : new DeviceDocResult(r.Status, r.Message, false, null);
    }

    /// <summary>The sales invoice as the rep prints it: the lines with their unit, the totals, the paid and remaining amounts,
    /// and the QR of the tax invoice (Saudi Arabia).</summary>
    public async Task<PosDocument?> SaleDocumentAsync(int invoiceId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await db.SalesInvoices.AsNoTracking().FirstOrDefaultAsync(x => x.SalesInvoiceID == invoiceId);
        if (h is null)
        {
            return null;
        }
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID, u => SalesService.Name(u.UnitName, u.UnitNameEn, english));
        var rows = await (from d in db.SalesInvoiceDetails
                          join p in db.Products on d.ProductID equals p.ProductID
                          where d.SalesInvoiceID == invoiceId
                          orderby d.LineNumber
                          select new { p.ProductName, p.ProductNameEn, Unit = d.UnitID ?? p.UnitID, d.Quantity, d.LineTotal }).ToListAsync();
        var lines = rows.Select(x => new PosSlipLine(SalesService.Name(x.ProductName, x.ProductNameEn, english) + " - " + units.GetValueOrDefault(x.Unit, ""),
                                                     x.Quantity, x.LineTotal)).ToArray();
        var totals = new List<PosSlipRow>
        {
            new("الإجمالي قبل الضريبة", h.TaxableAmount), new("الضريبة", h.Tax), new("الإجمالي شامل الضريبة", h.TotalAmount),
            new("المدفوع", h.PaidAmount), new("المتبقي", h.RemainingAmount),
        };
        string cashier = await db.Employees.Where(e => e.EmployeeID == h.EmployeeID).Select(e => e.EmployeeName).FirstOrDefaultAsync() ?? "";
        var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == h.CustomerID);
        string title = h.InvoiceSubType == "STANDARD" ? "فاتورة ضريبية" : "فاتورة ضريبية مبسطة";
        return new PosDocument("SALE", h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate, title, cashier,
                               SalesService.Name(c.CustomerName, c.CustomerNameEn, english), lines, totals.ToArray(),
                               h.PaymentType == "CREDIT" ? "آجل" : "نقدي", h.QRCodeData);
    }
}
