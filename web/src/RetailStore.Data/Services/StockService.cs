using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

/// <summary>The stock of a product in a warehouse (0 = every warehouse): the quantity, what open customer orders reserve
/// there, the average cost and the value.</summary>
public sealed record StockRow(int ProductID, string ProductCode, string? Barcode, string ProductName, string CategoryName, string UnitName,
                              decimal Quantity, decimal Reserved, decimal MinimumQuantity, decimal AverageCost, decimal Value,
                              decimal SellingPrice, bool IsLow);

/// <summary>A line of the item card (ItemCardQuery): the opening line first (MoveID 0), then each move in its order.</summary>
public sealed record ItemCardLine(int MoveID, DateTime MovementDate, string MovementType, string? ReferenceNumber, string? WarehouseName,
                                  string? PartyName, decimal QtyIn, decimal QtyOut, decimal Balance, decimal? MoveCost, decimal NetValue,
                                  string? Notes);

public sealed record TransferLineInput(int ProductID, decimal Quantity);

/// <summary>A new draft transfer (modWarehouses.CreateStockTransfer) with its lines.</summary>
public sealed record TransferRequest(int FromWarehouseID, int ToWarehouseID, string? Notes, IReadOnlyList<TransferLineInput> Lines);

public sealed record TransferLineView(int StockTransferLineID, int ProductID, string ProductCode, string ProductName, string UnitName,
                                      decimal Quantity, decimal UnitCost, decimal Available);

public sealed record TransferView(int StockTransferID, string TransferNumber, DateTime TransferDate, int FromWarehouseID, string FromName,
                                  int ToWarehouseID, string ToName, string Status, bool InTransit, DateTime? ReceivedAt, string? ReceivedBy,
                                  string EmployeeName, string? Notes, bool CanReceive, List<TransferLineView> Lines);

public sealed record TransferRow(int StockTransferID, string TransferNumber, DateTime TransferDate, string FromName, string ToName,
                                 string Status, int LineCount, decimal Quantity, string EmployeeName);

public sealed record StockResult(SaveStatus Status, string Message, int Id = 0, string Number = "")
{
    public static StockResult Refused(string message) => new(SaveStatus.Refused, message);
}

/// <summary>Stock and warehouses on the web and the API (plan 6D, docs/80): the stock of each warehouse
/// (modWarehouses.AvailableQty, less modSalesOrders.ReservedQty), the item card (ItemCardQuery) and the transfers between
/// warehouses (modWarehouses: CreateStockTransfer, SendStockTransfer, ReceiveStockTransfer, UnpostStockTransfer,
/// DeleteStockTransfer) with the same checks, messages, rows and audit as Access, so Access builds the same journal
/// entries (TRANSFER, STOCK_RECEIVE).</summary>
public sealed class StockService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const int TtTransferOut = 9;
    public const int TtTransferIn = 10;
    public const string TransferForm = "frmStockTransfer";

    private SalesService Sales => new(factory, permissions);

    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    /// <summary>modClosing.ClosedPeriodProblem: "" when the day is open.</summary>
    public static string ClosedPeriodProblem(DateTime when, DateTime? closedThrough)
    {
        if (closedThrough is not { } closed || when.Date > closed.Date)
        {
            return "";
        }
        return "التاريخ " + GDate(when) + " في فترة مقفلة (الإقفال حتى " + GDate(closed) + ").\n" +
               "لا يُضاف ولا يُعدَّل ولا يُحذف فيها شيء إلا بعد إعادة فتح الفترة.";
    }

    /// <summary>modReports.GDate.</summary>
    public static string GDate(DateTime d) => d.ToString("yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture);

    // ------------------------------------------------------------------ the stock of the warehouses

    /// <summary>The stock of the active stock products in a warehouse (0 = every warehouse: Products.CurrentQuantity), with
    /// a search in the code, barcode or name; lowOnly = at or under the minimum.</summary>
    public async Task<List<StockRow>> StockAsync(int warehouseId, string? search, bool lowOnly, bool english, int take = 500)
    {
        await using var db = await factory.CreateDbContextAsync();
        string s = (search ?? "").Trim();
        var q = db.Products.AsNoTracking().Where(p => p.IsActive && p.TrackStock);
        if (s.Length > 0)
        {
            q = q.Where(p => p.ProductName.Contains(s) || (p.ProductNameEn != null && p.ProductNameEn.Contains(s)) || p.ProductCode.Contains(s) ||
                             (p.Barcode != null && p.Barcode.Contains(s)));
        }
        var products = await q.OrderBy(p => p.ProductName).ToListAsync();
        var categories = await db.Categories.AsNoTracking().ToDictionaryAsync(c => c.CategoryID);
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        Dictionary<int, decimal>? byWarehouse = null;
        bool onlyMain = !await db.InventoryTransactions.AnyAsync(t => t.WarehouseID != RepService.MainWarehouseId);
        if (warehouseId > 0 && !(warehouseId == RepService.MainWarehouseId && onlyMain))
        {
            byWarehouse = await db.InventoryTransactions.AsNoTracking().Where(t => t.WarehouseID == warehouseId)
                .GroupBy(t => t.ProductID).Select(g => new { g.Key, Qty = g.Sum(t => t.Quantity) }).ToDictionaryAsync(x => x.Key, x => x.Qty);
        }
        var rows = new List<StockRow>();
        foreach (var p in products)
        {
            decimal qty = byWarehouse is null ? p.CurrentQuantity : byWarehouse.GetValueOrDefault(p.ProductID);
            bool low = qty <= p.MinimumQuantity;
            if (lowOnly && !low)
            {
                continue;
            }
            decimal reserved = warehouseId > 0 ? await SalesService.ReservedQtyAsync(db, p.ProductID, warehouseId) : 0;
            rows.Add(new StockRow(p.ProductID, p.ProductCode, p.Barcode, Name(p.ProductName, p.ProductNameEn, english),
                                  categories.TryGetValue(p.CategoryID, out var c) ? Name(c.CategoryName, c.CategoryNameEn, english) : "",
                                  units.TryGetValue(p.UnitID, out var u) ? Name(u.UnitName, u.UnitNameEn, english) : "",
                                  qty, reserved, p.MinimumQuantity, p.AverageCost, MoneyMath.RoundMoney(qty * p.AverageCost), p.SellingPrice, low));
            if (rows.Count >= take)
            {
                break;
            }
        }
        return rows;
    }

    /// <summary>The item card of a product (ItemCardQuery): the opening quantity and value before the period, then each move
    /// of the period in one warehouse (0 = all) with its other party (qryStockMoveParty) and the running balance.</summary>
    public async Task<List<ItemCardLine>> ItemCardAsync(int productId, int warehouseId, DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var all = db.InventoryTransactions.AsNoTracking().Where(t => t.ProductID == productId && (warehouseId == 0 || t.WarehouseID == warehouseId));
        var before = await all.Where(t => t.TransactionDate < start).Select(t => new { t.Quantity, t.UnitCost }).ToListAsync();
        decimal openQty = before.Sum(t => t.Quantity);
        decimal openValue = before.Sum(t => MoneyMath.Currency(t.Quantity * t.UnitCost));
        var moves = await all.Where(t => t.TransactionDate >= start && t.TransactionDate < end)
            .OrderBy(t => t.TransactionDate).ThenBy(t => t.TransactionID).ToListAsync();
        var types = await db.TransactionTypes.AsNoTracking().ToDictionaryAsync(t => t.TransactionTypeID);
        var warehouses = await db.Warehouses.AsNoTracking().ToDictionaryAsync(w => w.WarehouseID);
        var parties = await PartiesAsync(db, moves, warehouses, english);
        var list = new List<ItemCardLine>
        {
            new(0, start, "رصيد أول المدة", null, null, null, openQty > 0 ? openQty : 0, openQty < 0 ? -openQty : 0, openQty, null, openValue, null),
        };
        decimal balance = openQty;
        foreach (var t in moves)
        {
            balance += t.Quantity;
            list.Add(new ItemCardLine(t.TransactionID, t.TransactionDate,
                                      types.TryGetValue(t.TransactionTypeID, out var tt) ? Name(tt.TypeName, tt.TypeNameEn, english) : "",
                                      t.ReferenceNumber,
                                      warehouses.TryGetValue(t.WarehouseID, out var w) ? Name(w.WarehouseName, w.WarehouseNameEn, english) : null,
                                      parties.GetValueOrDefault(t.TransactionID), t.Quantity > 0 ? t.Quantity : 0, t.Quantity < 0 ? -t.Quantity : 0,
                                      balance, t.UnitCost, MoneyMath.Currency(t.Quantity * t.UnitCost), t.Notes));
        }
        return list;
    }

    /// <summary>qryStockMoveParty: the customer, the supplier, the party of a stock voucher or the other warehouse of a
    /// transfer.</summary>
    private static async Task<Dictionary<int, string>> PartiesAsync(RetailDbContext db, List<InventoryTransactions> moves,
                                                                   Dictionary<int, Warehouses> warehouses, bool english)
    {
        var result = new Dictionary<int, string>();
        List<int> Ids(string type) => moves.Where(m => m.ReferenceType == type && m.ReferenceID != null).Select(m => m.ReferenceID!.Value).Distinct().ToList();
        var customers = await db.Customers.AsNoTracking().ToDictionaryAsync(c => c.CustomerID, c => Name(c.CustomerName, c.CustomerNameEn, english));
        var suppliers = await db.Suppliers.AsNoTracking().ToDictionaryAsync(s => s.SupplierID, s => Name(s.SupplierName, s.SupplierNameEn, english));
        var sale = Ids("SALE");
        var saleParty = await db.SalesInvoices.Where(h => sale.Contains(h.SalesInvoiceID)).ToDictionaryAsync(h => h.SalesInvoiceID, h => customers[h.CustomerID]);
        var sret = Ids("SALES_RETURN");
        var sretParty = await db.SalesReturns.Where(h => sret.Contains(h.SalesReturnID)).ToDictionaryAsync(h => h.SalesReturnID, h => customers[h.CustomerID]);
        var pur = Ids("PURCHASE");
        var purParty = await db.PurchaseInvoices.Where(h => pur.Contains(h.PurchaseInvoiceID)).ToDictionaryAsync(h => h.PurchaseInvoiceID, h => suppliers[h.SupplierID]);
        var pret = Ids("PURCHASE_RETURN");
        var pretParty = await db.PurchaseReturns.Where(h => pret.Contains(h.PurchaseReturnID)).ToDictionaryAsync(h => h.PurchaseReturnID, h => suppliers[h.SupplierID]);
        var docs = Ids("STOCK_DOC");
        var docParty = await (from d in db.StockDocuments
                              join p in db.StockParties on d.StockPartyID equals p.StockPartyID
                              where docs.Contains(d.StockDocumentID)
                              select new { d.StockDocumentID, p.PartyName, p.PartyNameEn }).ToDictionaryAsync(x => x.StockDocumentID, x => Name(x.PartyName, x.PartyNameEn, english));
        var notes = Ids("STOCK_NOTE");
        var noteRows = await db.StockNotes.Where(n => notes.Contains(n.StockNoteID)).ToListAsync();
        var trs = Ids("TRANSFER");
        var transfers = await db.StockTransfers.Where(x => trs.Contains(x.StockTransferID)).ToDictionaryAsync(x => x.StockTransferID);
        string W(int id) => warehouses.TryGetValue(id, out var w) ? Name(w.WarehouseName, w.WarehouseNameEn, english) : "";
        foreach (var m in moves.Where(m => m.ReferenceID != null))
        {
            int id = m.ReferenceID!.Value;
            string? party = m.ReferenceType switch
            {
                "SALE" => saleParty.GetValueOrDefault(id),
                "SALES_RETURN" => sretParty.GetValueOrDefault(id),
                "PURCHASE" => purParty.GetValueOrDefault(id),
                "PURCHASE_RETURN" => pretParty.GetValueOrDefault(id),
                "STOCK_DOC" => docParty.GetValueOrDefault(id),
                "STOCK_NOTE" => noteRows.FirstOrDefault(n => n.StockNoteID == id) is { } n
                    ? (n.NoteType == "DELIVERY" ? (n.CustomerID is { } c ? customers.GetValueOrDefault(c) : null)
                                                : (n.SupplierID is { } s ? suppliers.GetValueOrDefault(s) : null))
                    : null,
                "TRANSFER" => transfers.TryGetValue(id, out var x) ? W(m.Quantity < 0 ? x.ToWarehouseID : x.FromWarehouseID) : null,
                _ => null,
            };
            if (party is not null)
            {
                result[m.TransactionID] = party;
            }
        }
        return result;
    }

    /// <summary>The active warehouses (modWarehouses.WarehouseRows) and the user's warehouse (CurrentWarehouseID).</summary>
    public async Task<(List<(int Id, string Name)> Warehouses, int Current)> WarehousesAsync(int employeeId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Warehouses.AsNoTracking().Where(w => w.IsActive).OrderBy(w => w.WarehouseID).ToListAsync();
        return (rows.Select(w => (w.WarehouseID, Name(w.WarehouseName, w.WarehouseNameEn, english))).ToList(),
                await RepService.UserWarehouseAsync(db, employeeId));
    }

    // ------------------------------------------------------------------ transfers between warehouses (docs/58)

    private const string NoPermission = "ليست لديك صلاحية التحويل بين المخازن.";

    private async Task<string> TransferPermissionAsync(int employeeId)
    {
        string open = await permissions.OpenProblemAsync(employeeId, TransferForm);
        if (open.Length > 0)
        {
            return open;
        }
        return await permissions.HasPermissionAsync(employeeId, "STOCK_TRANSFER") ? "" : NoPermission;
    }

    private static Task<bool> IsActiveWarehouseAsync(RetailDbContext db, int id) => db.Warehouses.AnyAsync(w => w.IsActive && w.WarehouseID == id);

    /// <summary>modWarehouses.StockTransferLineProblem.</summary>
    private static async Task<string> TransferLineProblemAsync(RetailDbContext db, int productId, decimal qty)
    {
        var p = await db.Products.AsNoTracking().FirstOrDefaultAsync(x => x.ProductID == productId);
        if (p is null)
        {
            return "اختر الصنف.";
        }
        if (!p.TrackStock)
        {
            return "هذا الصنف لا يتابع المخزون، فلا يُحوَّل.";
        }
        return qty <= 0 ? "الكمية أكبر من صفر." : "";
    }

    /// <summary>modWarehouses.CreateStockTransfer, then its lines (frmStockTransferLines checks each line with
    /// StockTransferLineProblem).</summary>
    public async Task<StockResult> CreateTransferAsync(ActingUser user, TransferRequest req)
    {
        string msg = await TransferPermissionAsync(user.EmployeeID);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        if (!await IsActiveWarehouseAsync(db, req.FromWarehouseID))
        {
            msg = "اختر المخزن المرسِل.";
        }
        else if (!await IsActiveWarehouseAsync(db, req.ToWarehouseID))
        {
            msg = "اختر المخزن المستلم.";
        }
        else if (req.FromWarehouseID == req.ToWarehouseID)
        {
            msg = "المخزن المستلم غير المرسِل.";
        }
        foreach (var l in req.Lines)
        {
            if (msg.Length == 0)
            {
                msg = await TransferLineProblemAsync(db, l.ProductID, l.Quantity);
            }
        }
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        var t = new StockTransfers
        {
            TransferNumber = await NumberSequences.NextNumberAsync(db, "STOCK_TRANSFER"), TransferDate = DateTime.Now,
            FromWarehouseID = req.FromWarehouseID, ToWarehouseID = req.ToWarehouseID, Status = "DRAFT", InTransit = false,
            EmployeeID = user.EmployeeID, Notes = SalesService.Clip(req.Notes, 255),
        };
        db.StockTransfers.Add(t);
        await db.SaveChangesAsync();
        foreach (var l in req.Lines)
        {
            db.StockTransferLines.Add(new StockTransferLines { StockTransferID = t.StockTransferID, ProductID = l.ProductID, Quantity = l.Quantity });
        }
        Audit.LogAction(db, user, "STOCK_TRANSFER_NEW", "StockTransfers", t.StockTransferID.ToString());
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        return new StockResult(SaveStatus.Saved, "", t.StockTransferID, t.TransferNumber);
    }

    /// <summary>The lines of a draft are replaced (the lines grid of frmStockTransfer, each line checked).</summary>
    public async Task<StockResult> SaveTransferLinesAsync(ActingUser user, int transferId, IReadOnlyList<TransferLineInput> lines)
    {
        string msg = await TransferPermissionAsync(user.EmployeeID);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null || t.Status != "DRAFT")
        {
            return StockResult.Refused("التحويل غير موجود أو تم ترحيله.");
        }
        foreach (var l in lines)
        {
            msg = await TransferLineProblemAsync(db, l.ProductID, l.Quantity);
            if (msg.Length > 0)
            {
                return StockResult.Refused(msg);
            }
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        db.StockTransferLines.RemoveRange(await db.StockTransferLines.Where(x => x.StockTransferID == transferId).ToListAsync());
        foreach (var l in lines)
        {
            db.StockTransferLines.Add(new StockTransferLines { StockTransferID = transferId, ProductID = l.ProductID, Quantity = l.Quantity });
        }
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        return new StockResult(SaveStatus.Saved, "", transferId, t.TransferNumber);
    }

    /// <summary>modWarehouses.SendStockTransfer: out of the sending warehouse at the average cost; into the receiving one
    /// now, or later when it confirms receipt (in transit). confirmStock = the user answered yes to the question on the stock.</summary>
    public async Task<StockResult> SendTransferAsync(ActingUser user, int transferId, bool confirmStock = false)
    {
        string msg = await TransferPermissionAsync(user.EmployeeID);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null || t.Status != "DRAFT")
        {
            return StockResult.Refused("التحويل غير موجود أو تم ترحيله.");
        }
        var settings = await SettingsReader.ReadAsync(db);
        msg = await PurchaseService.WarehouseProblemAsync(db, t.FromWarehouseID);
        if (msg.Length == 0)
        {
            msg = await PurchaseService.WarehouseProblemAsync(db, t.ToWarehouseID);
        }
        if (msg.Length == 0)
        {
            msg = ClosedPeriodProblem(DateTime.Today, settings.ClosedThrough);
        }
        var lines = await db.StockTransferLines.Where(x => x.StockTransferID == transferId).OrderBy(x => x.StockTransferLineID).ToListAsync();
        if (msg.Length == 0)
        {
            foreach (var l in lines)
            {
                if (msg.Length == 0)
                {
                    msg = await TransferLineProblemAsync(db, l.ProductID, l.Quantity);
                }
            }
            if (msg.Length == 0 && lines.Count == 0)
            {
                msg = "أضف الأصناف وكمياتها أولًا.";
            }
        }
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        var ids = lines.Select(l => l.ProductID).Distinct().ToList();
        var products = await db.Products.AsNoTracking().Where(p => ids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID);
        const string operation = "التحويل بين المخازن";
        var stock = await Sales.CheckStockAsync(db, settings, user.EmployeeID, lines.Select(l => (products[l.ProductID], l.Quantity)).ToList(),
                                                t.FromWarehouseID, operation);
        if (stock.Status == SaveStatus.Refused || (stock.Status == SaveStatus.NeedsConfirmation && !confirmStock))
        {
            return new StockResult(stock.Status, stock.Message);
        }

        bool transit = await db.Warehouses.Where(w => w.WarehouseID == t.ToWarehouseID).Select(w => w.ConfirmReceipt).FirstOrDefaultAsync();
        await using var tx = await db.Database.BeginTransactionAsync();
        if (stock.ShortList.Length > 0)
        {
            Audit.LogAction(db, user, "NEGATIVE_STOCK", "Products", "",
                            stock.Status == SaveStatus.Saved ? stock.Message : operation + " - approved by user:" + stock.ShortList);
        }
        foreach (var l in lines)
        {
            decimal cost = await db.Products.Where(p => p.ProductID == l.ProductID).Select(p => p.AverageCost).FirstOrDefaultAsync();
            l.UnitCost = cost;
            await db.SaveChangesAsync();
            await SalesService.ApplyStockMovementAsync(db, user, l.ProductID, -l.Quantity, TtTransferOut, cost, "TRANSFER", transferId,
                                                       t.TransferNumber, t.FromWarehouseID);
            if (!transit)
            {
                await SalesService.ApplyStockMovementAsync(db, user, l.ProductID, l.Quantity, TtTransferIn, cost, "TRANSFER", transferId,
                                                           t.TransferNumber, t.ToWarehouseID);
            }
        }
        var now = DateTime.Now;
        t.TransferDate = now;
        if (transit)
        {
            t.Status = "SENT";
            t.InTransit = true;
        }
        else
        {
            t.Status = "RECEIVED";
            t.InTransit = false;
            t.ReceivedAt = now;
            t.ReceivedByID = user.EmployeeID;
        }
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        Audit.LogAction(db, user, "STOCK_TRANSFER_SEND", "StockTransfers", t.TransferNumber,
                        "From=" + t.FromWarehouseID + " To=" + t.ToWarehouseID + " Transit=" + (transit ? "True" : "False"));
        await db.SaveChangesAsync();
        return new StockResult(SaveStatus.Saved, "", transferId, t.TransferNumber);
    }

    /// <summary>modWarehouses.CanReceiveIn: its keeper, a user whose warehouse it is, or the administrator; anyone with the
    /// permission when it has no keeper.</summary>
    public async Task<bool> CanReceiveInAsync(RetailDbContext db, int employeeId, int warehouseId)
    {
        if (!await permissions.HasPermissionAsync(employeeId, "STOCK_TRANSFER"))
        {
            return false;
        }
        int? keeper = await db.Warehouses.Where(w => w.WarehouseID == warehouseId).Select(w => w.KeeperID).FirstOrDefaultAsync();
        if (keeper is null)
        {
            return true;
        }
        if (keeper == employeeId || await RepService.UserWarehouseAsync(db, employeeId) == warehouseId)
        {
            return true;
        }
        return await db.Employees.AnyAsync(e => e.EmployeeID == employeeId && (e.RoleID == Domain.AppData.AdminRoleId || e.IsDeveloper));
    }

    /// <summary>modWarehouses.ReceiveStockTransfer: the goods enter the receiving warehouse now, at the cost they left with.</summary>
    public async Task<StockResult> ReceiveTransferAsync(ActingUser user, int transferId)
    {
        string msg = await permissions.OpenProblemAsync(user.EmployeeID, TransferForm);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null || t.Status != "SENT")
        {
            return StockResult.Refused("هذا التحويل ليس بالطريق.");
        }
        if (!await CanReceiveInAsync(db, user.EmployeeID, t.ToWarehouseID))
        {
            string name = await db.Warehouses.Where(w => w.WarehouseID == t.ToWarehouseID).Select(w => w.WarehouseName).FirstOrDefaultAsync() ?? "";
            return StockResult.Refused("يستلم التحويل أمين " + name + ".");
        }
        var settings = await SettingsReader.ReadAsync(db);
        msg = ClosedPeriodProblem(DateTime.Today, settings.ClosedThrough);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        var before = Audit.Snapshot("StockTransfers", t);
        await using var tx = await db.Database.BeginTransactionAsync();
        foreach (var l in await db.StockTransferLines.AsNoTracking().Where(x => x.StockTransferID == transferId).OrderBy(x => x.StockTransferLineID).ToListAsync())
        {
            await SalesService.ApplyStockMovementAsync(db, user, l.ProductID, l.Quantity, TtTransferIn, l.UnitCost, "TRANSFER", transferId,
                                                       t.TransferNumber, t.ToWarehouseID);
        }
        t.Status = "RECEIVED";
        t.ReceivedAt = DateTime.Now;
        t.ReceivedByID = user.EmployeeID;
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        await AuditEditedAsync(db, user, "STOCK_TRANSFER_RECEIVE", t, before);
        return new StockResult(SaveStatus.Saved, "", transferId, t.TransferNumber);
    }

    /// <summary>modWarehouses.UnpostStockTransfer: a transfer still in transit goes back to draft (its moves out are
    /// deleted, the goods are back in the sending warehouse). A received transfer stays.</summary>
    public async Task<StockResult> UnpostTransferAsync(ActingUser user, int transferId)
    {
        string msg = await TransferPermissionAsync(user.EmployeeID);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null || t.Status != "SENT")
        {
            return StockResult.Refused("يُلغى ترحيل التحويل بالطريق فقط. التحويل المستلم يُعاد بتحويل جديد.");
        }
        var settings = await SettingsReader.ReadAsync(db);
        msg = ClosedPeriodProblem(t.TransferDate, settings.ClosedThrough);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        var before = Audit.Snapshot("StockTransfers", t);
        await using var tx = await db.Database.BeginTransactionAsync();
        foreach (var l in await db.StockTransferLines.AsNoTracking().Where(x => x.StockTransferID == transferId).ToListAsync())
        {
            var p = await db.Products.FirstAsync(x => x.ProductID == l.ProductID);
            p.CurrentQuantity += l.Quantity;
            p.UpdatedAt = DateTime.Now;
        }
        db.InventoryTransactions.RemoveRange(await db.InventoryTransactions
            .Where(x => x.ReferenceType == "TRANSFER" && x.ReferenceID == transferId).ToListAsync());
        t.Status = "DRAFT";
        t.InTransit = false;
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        await AuditEditedAsync(db, user, "STOCK_TRANSFER_UNPOST", t, before);
        return new StockResult(SaveStatus.Saved, "", transferId, t.TransferNumber);
    }

    /// <summary>modWarehouses.DeleteStockTransfer: a draft only; its lines go with it.</summary>
    public async Task<StockResult> DeleteTransferAsync(ActingUser user, int transferId)
    {
        string msg = await TransferPermissionAsync(user.EmployeeID);
        if (msg.Length > 0)
        {
            return StockResult.Refused(msg);
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null || t.Status != "DRAFT")
        {
            return StockResult.Refused("تُحذف المسودة فقط. ألغِ ترحيل التحويل بالطريق أولًا.");
        }
        var snapshot = Audit.Snapshot("StockTransfers", t);
        await using var tx = await db.Database.BeginTransactionAsync();
        db.StockTransferLines.RemoveRange(await db.StockTransferLines.Where(x => x.StockTransferID == transferId).ToListAsync());
        db.StockTransfers.Remove(t);
        await db.SaveChangesAsync();
        // modAudit.AuditDeleted: every field that had a value
        var changes = snapshot.Where(s => !Audit.IsEmpty(s.Value))
                              .Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList();
        await Audit.WriteAsync(db, user, "STOCK_TRANSFER_DELETE", "StockTransfers", transferId.ToString(),
                               Audit.Label(snapshot.Select(s => (s.Field, s.Value))), changes);
        await tx.CommitAsync();
        return new StockResult(SaveStatus.Saved, "", transferId, t.TransferNumber);
    }

    /// <summary>modAudit.AuditEdited: the fields whose value changed.</summary>
    private static async Task AuditEditedAsync(RetailDbContext db, ActingUser user, string action, StockTransfers t,
                                               List<(string Field, string Caption, object? Value)> before)
    {
        var after = Audit.Snapshot("StockTransfers", t);
        var changes = before.Zip(after).Where(x => !Audit.SameValue(x.First.Value, x.Second.Value))
                            .Select(x => new AuditChange(x.First.Field, x.First.Caption, Audit.ValueText(x.First.Value), Audit.ValueText(x.Second.Value)))
                            .ToList();
        await Audit.WriteAsync(db, user, action, "StockTransfers", t.StockTransferID.ToString(),
                               Audit.Label(after.Select(s => (s.Field, s.Value))), changes);
    }

    /// <summary>The transfers, newest first: a status (empty = all) and a warehouse on either side (0 = all).</summary>
    public async Task<List<TransferRow>> TransfersAsync(string? status, int warehouseId, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var q = db.StockTransfers.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status))
        {
            q = q.Where(t => t.Status == status);
        }
        if (warehouseId > 0)
        {
            q = q.Where(t => t.FromWarehouseID == warehouseId || t.ToWarehouseID == warehouseId);
        }
        var rows = await q.OrderByDescending(t => t.TransferDate).ThenByDescending(t => t.StockTransferID).Take(take).ToListAsync();
        var ids = rows.Select(r => r.StockTransferID).ToList();
        var sums = await db.StockTransferLines.Where(l => ids.Contains(l.StockTransferID)).GroupBy(l => l.StockTransferID)
            .Select(g => new { g.Key, N = g.Count(), Q = g.Sum(l => l.Quantity) }).ToDictionaryAsync(x => x.Key);
        var warehouses = await db.Warehouses.AsNoTracking().ToDictionaryAsync(w => w.WarehouseID);
        var employees = await db.Employees.AsNoTracking().ToDictionaryAsync(e => e.EmployeeID, e => e.EmployeeName);
        string W(int id) => warehouses.TryGetValue(id, out var w) ? Name(w.WarehouseName, w.WarehouseNameEn, english) : "";
        return rows.Select(r => new TransferRow(r.StockTransferID, r.TransferNumber, r.TransferDate, W(r.FromWarehouseID), W(r.ToWarehouseID), r.Status,
                                                sums.TryGetValue(r.StockTransferID, out var s) ? s.N : 0, s?.Q ?? 0,
                                                employees.GetValueOrDefault(r.EmployeeID, ""))).ToList();
    }

    /// <summary>One transfer with its lines; Available = the free stock of the sending warehouse; CanReceive = the user may
    /// confirm its receipt.</summary>
    public async Task<TransferView?> TransferAsync(int transferId, int employeeId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.StockTransfers.AsNoTracking().FirstOrDefaultAsync(x => x.StockTransferID == transferId);
        if (t is null)
        {
            return null;
        }
        var warehouses = await db.Warehouses.AsNoTracking().ToDictionaryAsync(w => w.WarehouseID);
        var employees = await db.Employees.AsNoTracking().ToDictionaryAsync(e => e.EmployeeID, e => e.EmployeeName);
        var units = await db.Units.AsNoTracking().ToDictionaryAsync(u => u.UnitID);
        string W(int id) => warehouses.TryGetValue(id, out var w) ? Name(w.WarehouseName, w.WarehouseNameEn, english) : "";
        var lines = await (from l in db.StockTransferLines.AsNoTracking()
                           join p in db.Products on l.ProductID equals p.ProductID
                           where l.StockTransferID == transferId
                           orderby l.StockTransferLineID
                           select new { l, p }).ToListAsync();
        var views = new List<TransferLineView>();
        foreach (var x in lines)
        {
            decimal available = await SalesService.AvailableQtyAsync(db, x.p.ProductID, t.FromWarehouseID)
                                - await SalesService.ReservedQtyAsync(db, x.p.ProductID, t.FromWarehouseID);
            views.Add(new TransferLineView(x.l.StockTransferLineID, x.p.ProductID, x.p.ProductCode, Name(x.p.ProductName, x.p.ProductNameEn, english),
                                           units.TryGetValue(x.p.UnitID, out var u) ? Name(u.UnitName, u.UnitNameEn, english) : "",
                                           x.l.Quantity, x.l.UnitCost, available));
        }
        bool canReceive = t.Status == "SENT" && await CanReceiveInAsync(db, employeeId, t.ToWarehouseID);
        return new TransferView(t.StockTransferID, t.TransferNumber, t.TransferDate, t.FromWarehouseID, W(t.FromWarehouseID), t.ToWarehouseID,
                                W(t.ToWarehouseID), t.Status, t.InTransit, t.ReceivedAt,
                                t.ReceivedByID is { } r ? employees.GetValueOrDefault(r) : null, employees.GetValueOrDefault(t.EmployeeID, ""),
                                t.Notes, canReceive, views);
    }
}
