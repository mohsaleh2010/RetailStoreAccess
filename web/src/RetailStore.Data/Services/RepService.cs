using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;

namespace RetailStore.Data.Services;

/// <summary>The rep linked to the user (SalesReps.EmployeeID), his van warehouse, his custody box and his limits.</summary>
public sealed record RepProfile(int SalesRepID, string RepCode, string RepName, string? Region, int? VanWarehouseID,
                                string? VanWarehouseName, int? CustodyBoxID, string? CustodyBoxName, int OrderWarehouseID,
                                decimal CreditLimit, decimal MaxDiscountPercent, decimal CustomersBalance);

/// <summary>A customer of the rep: assigned to him (Customers.SalesRepID) or on his route.</summary>
public sealed record RepCustomer(int CustomerID, string CustomerName, string? Mobile, string? Phone, string? City,
                                 string? District, string? Address, string? VATNumber, decimal CurrentBalance,
                                 bool AllowCredit, decimal CreditLimit, int? PriceListID, bool Assigned, int[] RouteWeekdays);

/// <summary>A product as the rep sells it: the price of an order line (OrderDefaultPrice) and the stock of his van.</summary>
public sealed record RepProduct(int ProductID, string ProductCode, string? Barcode, string ProductName, int UnitID,
                                string VATCategory, decimal SellingPrice, decimal Price, bool RepPrice, bool TrackStock,
                                decimal VanQuantity);

/// <summary>A visit of the day (RepVisits) with the name of its customer.</summary>
public sealed record RepVisit(int RepVisitID, DateTime VisitDate, int CustomerID, string CustomerName, bool IsPlanned,
                              short StopOrder, string Result, DateTime? VisitedAt, string? Notes);

/// <summary>A line of a new customer order: Price null or 0 = the default price of the line (OrderDefaultPrice).</summary>
public sealed record OrderLineInput(int ProductID, decimal Quantity, decimal? UnitPrice);

public sealed record OrderLineView(int SalesOrderLineID, int ProductID, string ProductName, decimal Quantity, decimal UnitPrice);

/// <summary>Everything the rep app keeps on the phone to work without the internet (docs/78): the store name, the profile,
/// the customers with their price group, the products with the price without a customer and the van stock, the prices of
/// each price group (the prices that differ only), the visits of the day and the orders of the last days.</summary>
public sealed record RepSnapshot(DateTime ServerTime, string StoreName, RepProfile Profile, List<RepCustomer> Customers, List<RepProduct> Products,
                                 Dictionary<string, Dictionary<int, decimal>> PriceGroups, Dictionary<int, string> CustomerPriceGroups,
                                 List<RepVisit> Visits, List<OrderView> Orders, bool CanVisits);

public sealed record OrderView(int SalesOrderID, string OrderNumber, DateTime OrderDate, int CustomerID, string CustomerName,
                               int WarehouseID, int? SalesRepID, DateTime? RequiredDate, string Status, string? Notes,
                               decimal Total, OrderLineView[] Lines, string? ClientRef = null);

/// <summary>The work of a sales rep through the API (docs/77), first used by the rep app (plan 4D): the same rules and
/// messages as modRepRoutes (visits), modSalesOrders (a draft order and its lines), modPriceLists / modRepRoutes (the
/// price of an order line) and modWarehouses (the stock of a warehouse). Each method returns "" or the Arabic message.</summary>
public sealed class RepService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const string NotARep = "هذا المستخدم ليس مندوبًا نشطًا. اربطه بمندوب في شاشة المندوبين.";
    public const string ClientRefTooLong = "رقم الطلبية في الجهاز أطول من 40 حرفًا.";
    public const string ClientRefUsed = "رقم الطلبية في الجهاز مستعمل لمستخدم آخر.";
    public const int MainWarehouseId = 1;
    public static readonly string[] VisitResults = ["ORDER", "SALE", "PAYMENT", "NO_SALE", "CLOSED"];

    // ------------------------------------------------------------------ the rep

    /// <summary>modRepRoutes.SessionRepID outside a rep day: the active rep linked to the user, else 0.</summary>
    public static async Task<int> RepIdOfAsync(RetailDbContext db, int employeeId) =>
        await db.SalesReps.Where(r => r.IsActive && r.EmployeeID == employeeId)
            .Select(r => (int?)r.SalesRepID).MinAsync() ?? 0;

    public async Task<int> RepIdOfAsync(int employeeId)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await RepIdOfAsync(db, employeeId);
    }

    /// <summary>modWarehouses.CurrentWarehouseID outside a rep day: the user's active warehouse, else the first active
    /// main warehouse, else the first active one.</summary>
    public static async Task<int> UserWarehouseAsync(RetailDbContext db, int employeeId)
    {
        int? v = await (from e in db.Employees
                        join w in db.Warehouses on e.WarehouseID equals w.WarehouseID
                        where e.EmployeeID == employeeId && w.IsActive
                        select (int?)w.WarehouseID).FirstOrDefaultAsync();
        v ??= await db.Warehouses.Where(w => w.IsActive && w.WarehouseType == "MAIN").Select(w => (int?)w.WarehouseID).MinAsync();
        v ??= await db.Warehouses.Where(w => w.IsActive).Select(w => (int?)w.WarehouseID).MinAsync();
        return v ?? MainWarehouseId;
    }

    public async Task<RepProfile?> ProfileAsync(int employeeId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        int repId = await RepIdOfAsync(db, employeeId);
        if (repId == 0)
        {
            return null;
        }
        var r = await db.SalesReps.AsNoTracking().FirstAsync(x => x.SalesRepID == repId);
        var van = await db.Warehouses.AsNoTracking().Where(w => w.IsActive && w.WarehouseType == "VAN" && w.SalesRepID == repId)
            .OrderBy(w => w.WarehouseID).FirstOrDefaultAsync();
        var box = await db.CashBoxes.AsNoTracking().Where(b => b.IsActive && b.BoxType == "CUSTODY" && b.SalesRepID == repId)
            .OrderBy(b => b.CashBoxID).FirstOrDefaultAsync();
        return new RepProfile(repId, r.RepCode, Name(r.RepName, r.RepNameEn, english), r.Region, van?.WarehouseID,
                              van is null ? null : Name(van.WarehouseName, van.WarehouseNameEn, english), box?.CashBoxID,
                              box is null ? null : Name(box.BoxName, box.BoxNameEn, english),
                              await UserWarehouseAsync(db, employeeId), r.CreditLimit, r.MaxDiscountPercent,
                              await CustomersBalanceAsync(db, repId));
    }

    /// <summary>modRepRoutes.RepCustomersBalance: the positive balances of the rep's customers.</summary>
    public static async Task<decimal> CustomersBalanceAsync(RetailDbContext db, int repId) =>
        await db.Customers.Where(c => c.CurrentBalance > 0 && c.SalesRepID == repId).SumAsync(c => (decimal?)c.CurrentBalance) ?? 0m;

    private static string Name(string arabic, string? english, bool inEnglish) =>
        inEnglish && !string.IsNullOrWhiteSpace(english) ? english : arabic;

    // ------------------------------------------------------------------ customers, products, stock

    /// <summary>The active customers assigned to the rep, and those on his active route stops.</summary>
    public async Task<List<RepCustomer>> CustomersAsync(int repId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var stops = await db.RouteStops.AsNoTracking().Where(s => s.SalesRepID == repId && s.IsActive)
            .Select(s => new { s.CustomerID, s.VisitWeekday }).ToListAsync();
        var onRoute = stops.Select(s => s.CustomerID).Distinct().ToList();
        var rows = await db.Customers.AsNoTracking()
            .Where(c => c.IsActive && (c.SalesRepID == repId || onRoute.Contains(c.CustomerID)))
            .OrderBy(c => c.CustomerName).ToListAsync();
        return rows.Select(c => new RepCustomer(
            c.CustomerID, Name(c.CustomerName, c.CustomerNameEn, english), c.Mobile, c.Phone, c.City, c.District, c.Address,
            c.VATNumber, c.CurrentBalance, c.AllowCredit, c.CreditLimit, c.PriceListID, c.SalesRepID == repId,
            stops.Where(s => s.CustomerID == c.CustomerID).Select(s => (int)s.VisitWeekday).Distinct().Order().ToArray())).ToList();
    }

    /// <summary>modWarehouses.AvailableQty of every product in a warehouse: Products.CurrentQuantity while every move
    /// is in the main warehouse, else the sum of the moves there.</summary>
    public static async Task<Dictionary<int, decimal>> WarehouseStockAsync(RetailDbContext db, int warehouseId)
    {
        if (warehouseId == MainWarehouseId && !await db.InventoryTransactions.AnyAsync(t => t.WarehouseID != MainWarehouseId))
        {
            return await db.Products.AsNoTracking().ToDictionaryAsync(p => p.ProductID, p => p.CurrentQuantity);
        }
        var sums = await db.InventoryTransactions.AsNoTracking().Where(t => t.WarehouseID == warehouseId)
            .GroupBy(t => t.ProductID).Select(g => new { g.Key, Qty = g.Sum(t => t.Quantity) }).ToListAsync();
        return sums.ToDictionary(s => s.Key, s => s.Qty);
    }

    public async Task<Dictionary<int, decimal>> WarehouseStockAsync(int warehouseId)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await WarehouseStockAsync(db, warehouseId);
    }

    /// <summary>The active products with the price an order line of the customer gets (OrderDefaultPrice with the rep of
    /// the order: the customer's rep, else this one; 0 = no customer: this rep's price, else the selling price) and the
    /// stock of the rep's van. RepPrice = the price came from a rep's price list.</summary>
    public async Task<List<RepProduct>> ProductsAsync(int repId, int customerId, int? vanWarehouseId, bool english, DateTime? today = null)
    {
        await using var db = await factory.CreateDbContextAsync();
        var products = await db.Products.AsNoTracking().Where(p => p.IsActive).OrderBy(p => p.ProductName).ToListAsync();
        int priceRep = customerId == 0 ? repId : await PriceRepAsync(db, customerId, repId);
        var list = await CustomerListAsync(db, customerId);
        var prices = await PricesAsync(db, products, priceRep, list, (today ?? DateTime.Today).Date);
        var van = vanWarehouseId is { } w ? await WarehouseStockAsync(db, w) : [];
        return products.Select(p => new RepProduct(p.ProductID, p.ProductCode, p.Barcode, Name(p.ProductName, p.ProductNameEn, english),
                                                   p.UnitID, p.VATCategory, p.SellingPrice, prices[p.ProductID].Price,
                                                   prices[p.ProductID].Rep, p.TrackStock, van.GetValueOrDefault(p.ProductID))).ToList();
    }

    /// <summary>The rep of an order of the customer (SalesRepFor): the active rep of the customer, else this rep.</summary>
    private static async Task<int> PriceRepAsync(RetailDbContext db, int customerId, int repId) =>
        await (from c in db.Customers
               join r in db.SalesReps on c.SalesRepID equals r.SalesRepID
               where r.IsActive && c.CustomerID == customerId
               select (int?)r.SalesRepID).FirstOrDefaultAsync() ?? repId;

    /// <summary>OrderDefaultPrice of every product for the rep of the order and the customer's list on a day; Rep = the
    /// usual price came from the rep's price list.</summary>
    private static async Task<Dictionary<int, (decimal Price, bool Rep)>> PricesAsync(RetailDbContext db, List<Products> products,
                                                                                      int priceRep, CustomerList? list, DateTime day)
    {
        var repPrices = await db.RepPrices.AsNoTracking().Where(r => r.SalesRepID == priceRep)
            .ToDictionaryAsync(r => r.ProductID, r => r.UnitPrice);
        var items = list is null ? [] : await ListItemsAsync(db, list.PriceListID, day);
        return products.ToDictionary(p => p.ProductID, p =>
        {
            bool rep = repPrices.TryGetValue(p.ProductID, out var rp);
            return (CustomerPrice(list, items, p.ProductID, rep ? rp : p.SellingPrice), rep);
        });
    }

    /// <summary>The data of the rep app (docs/78) in one read: the customers share a price group when they have the same
    /// rep of the order and the same active price list, so the phone prices every customer without asking the server;
    /// a group keeps only the prices that differ from the price without a customer.</summary>
    public async Task<RepSnapshot?> SnapshotAsync(int employeeId, bool english, int orderDays = 14, DateTime? today = null)
    {
        var profile = await ProfileAsync(employeeId, english);
        if (profile is null)
        {
            return null;
        }
        int repId = profile.SalesRepID;
        var day = (today ?? DateTime.Today).Date;
        var customers = await CustomersAsync(repId, english);
        var products = await ProductsAsync(repId, 0, profile.VanWarehouseID, english, day);
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Products.AsNoTracking().Where(p => p.IsActive).ToListAsync();
        var plain = products.ToDictionary(p => p.ProductID, p => p.Price);
        var groups = new Dictionary<string, Dictionary<int, decimal>>();
        var ofCustomer = new Dictionary<int, string>();
        foreach (var c in customers)
        {
            int priceRep = await PriceRepAsync(db, c.CustomerID, repId);
            var list = await CustomerListAsync(db, c.CustomerID);
            string key = "R" + priceRep + "-L" + (list?.PriceListID ?? 0);
            ofCustomer[c.CustomerID] = key;
            if (!groups.ContainsKey(key))
            {
                var prices = await PricesAsync(db, rows, priceRep, list, day);
                groups[key] = prices.Where(x => x.Value.Price != plain.GetValueOrDefault(x.Key))
                                    .ToDictionary(x => x.Key, x => x.Value.Price);
            }
        }
        bool canVisits = await permissions.HasPermissionAsync(employeeId, "REP_ROUTES");
        var visits = canVisits ? await VisitsAsync(repId, day, english) : [];
        var orders = await OrdersAsync(repId, employeeId, day.AddDays(-orderDays), day, english);
        var settings = await SettingsReader.ReadAsync(db);
        string store = english && !string.IsNullOrWhiteSpace(settings.StoreNameEn) ? settings.StoreNameEn! : (settings.StoreName ?? "");
        return new RepSnapshot(DateTime.Now, store, profile, customers, products, groups, ofCustomer, visits, orders, canVisits);
    }

    // ------------------------------------------------------------------ prices (modPriceLists, pieces only)

    private sealed record CustomerList(int PriceListID, decimal DiscountPercent);

    /// <summary>modPriceLists.CustomerPriceListID: the active price list of the customer.</summary>
    private static async Task<CustomerList?> CustomerListAsync(RetailDbContext db, int customerId) =>
        customerId == 0 ? null
            : await (from c in db.Customers
                     join l in db.PriceLists on c.PriceListID equals l.PriceListID
                     where l.IsActive && c.CustomerID == customerId
                     select new CustomerList(l.PriceListID, l.DiscountPercent)).FirstOrDefaultAsync();

    /// <summary>modPriceLists.PriceListItemPrice for the smallest unit of every product: among the items valid on the
    /// day, the one that starts last (then the last one entered).</summary>
    private static async Task<Dictionary<int, decimal>> ListItemsAsync(RetailDbContext db, int priceListId, DateTime day)
    {
        var rows = await db.PriceListItems.AsNoTracking()
            .Where(i => i.PriceListID == priceListId && i.UnitID == null
                        && (i.ValidFrom == null || i.ValidFrom <= day) && (i.ValidTo == null || i.ValidTo >= day))
            .ToListAsync();
        return rows.GroupBy(i => i.ProductID)
            .ToDictionary(g => g.Key, g => g.OrderByDescending(i => i.ValidFrom ?? DateTime.MinValue)
                                            .ThenByDescending(i => i.PriceListItemID).First().UnitPrice);
    }

    /// <summary>modPriceLists.CustomerPrice (smallest unit): the item of the customer's list, else the usual price less
    /// the discount of the list.</summary>
    private static decimal CustomerPrice(CustomerList? list, Dictionary<int, decimal> items, int productId, decimal usual)
    {
        if (list is null)
        {
            return usual;
        }
        if (items.TryGetValue(productId, out var item))
        {
            return item;
        }
        return list.DiscountPercent > 0 ? RoundMoney(usual * (1 - list.DiscountPercent)) : usual;
    }

    /// <summary>modCommon.RoundMoney: half away from zero.</summary>
    public static decimal RoundMoney(decimal value, int digits = 2) => Math.Round(value, digits, MidpointRounding.AwayFromZero);

    /// <summary>modSalesOrders.OrderDefaultPrice: the customer's price list, else the rep's price or the selling price
    /// less the discount of the customer's list.</summary>
    public static async Task<decimal> OrderDefaultPriceAsync(RetailDbContext db, int customerId, int? repId, int productId,
                                                             DateTime? today = null)
    {
        decimal? rep = repId is { } r
            ? await db.RepPrices.Where(x => x.SalesRepID == r && x.ProductID == productId).Select(x => (decimal?)x.UnitPrice).FirstOrDefaultAsync()
            : null;
        decimal usual = rep ?? await db.Products.Where(p => p.ProductID == productId).Select(p => p.SellingPrice).FirstOrDefaultAsync();
        var list = await CustomerListAsync(db, customerId);
        var items = list is null ? [] : await ListItemsAsync(db, list.PriceListID, (today ?? DateTime.Today).Date);
        return CustomerPrice(list, items, productId, usual);
    }

    // ------------------------------------------------------------------ visits (modRepRoutes)

    private async Task<string> CanRoutesAsync(int employeeId) =>
        await permissions.HasPermissionAsync(employeeId, "REP_ROUTES") ? "" : "ليست لديك صلاحية خط السير وزيارات المندوبين.";

    /// <summary>modRepRoutes.VisitWeekdayOf: VBA Weekday(d, vbSunday), Sunday = 1 ... Saturday = 7.</summary>
    public static int VisitWeekdayOf(DateTime d) => (int)d.DayOfWeek + 1;

    public async Task<List<RepVisit>> VisitsAsync(int repId, DateTime day, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var d = day.Date;
        var rows = await (from v in db.RepVisits.AsNoTracking()
                          join c in db.Customers on v.CustomerID equals c.CustomerID
                          where v.SalesRepID == repId && v.VisitDate >= d && v.VisitDate < d.AddDays(1)
                          orderby v.StopOrder, v.RepVisitID
                          select new { v, c.CustomerName, c.CustomerNameEn }).ToListAsync();
        return rows.Select(x => new RepVisit(x.v.RepVisitID, x.v.VisitDate, x.v.CustomerID, Name(x.CustomerName, x.CustomerNameEn, english),
                                             x.v.IsPlanned, x.v.StopOrder, x.v.Result, x.v.VisitedAt, x.v.Notes)).ToList();
    }

    /// <summary>modRepRoutes.CreateVisitPlan: the visits of the day from the route of its weekday; customers already in
    /// the day are kept as they are.</summary>
    public async Task<(string Problem, int Added)> CreateVisitPlanAsync(ActingUser user, int repId, DateTime day)
    {
        string problem = await CanRoutesAsync(user.EmployeeID);
        if (problem.Length > 0)
        {
            return (problem, 0);
        }
        await using var db = await factory.CreateDbContextAsync();
        if (!await db.SalesReps.AnyAsync(r => r.SalesRepID == repId && r.IsActive))
        {
            return ("اختر مندوبًا نشطًا.", 0);
        }
        var d = day.Date;
        int weekday = VisitWeekdayOf(d);
        var stops = await (from s in db.RouteStops
                           join c in db.Customers on s.CustomerID equals c.CustomerID
                           where s.IsActive && s.SalesRepID == repId && s.VisitWeekday == weekday
                           select new { s, CustomerActive = c.IsActive }).ToListAsync();
        if (stops.Count == 0)
        {
            return ("لا عملاء في خط سير المندوب ليوم " + WeekdayName(weekday) + ". أضفهم في «خط السير».", 0);
        }
        var already = await db.RepVisits.Where(v => v.VisitDate == d && v.SalesRepID == repId).Select(v => v.CustomerID).ToListAsync();
        int added = 0;
        foreach (var x in stops.Where(x => x.CustomerActive && !already.Contains(x.s.CustomerID)))
        {
            db.RepVisits.Add(new RepVisits
            {
                VisitDate = d, SalesRepID = repId, CustomerID = x.s.CustomerID, IsPlanned = true, StopOrder = x.s.StopOrder,
                Result = "PENDING", EmployeeID = user.EmployeeID,
            });
            already.Add(x.s.CustomerID);
            added++;
        }
        if (added > 0)
        {
            Audit.LogAction(db, user, "REP_VISIT_PLAN", "RepVisits", "", "Rep=" + repId + " Date=" + d.ToString("yyyy-MM-dd") +
                                                                         " Added=" + added);
            await db.SaveChangesAsync();
        }
        return ("", added);
    }

    /// <summary>modRepRoutes.RouteWeekdayName.</summary>
    public static string WeekdayName(int weekday) => weekday switch
    {
        1 => "الأحد", 2 => "الاثنين", 3 => "الثلاثاء", 4 => "الأربعاء", 5 => "الخميس", 6 => "الجمعة", _ => "السبت",
    };

    /// <summary>modRepRoutes.AddRepVisit: a visit outside the route.</summary>
    public async Task<(string Problem, int VisitID)> AddVisitAsync(ActingUser user, int repId, int customerId, DateTime day)
    {
        string problem = await CanRoutesAsync(user.EmployeeID);
        if (problem.Length > 0)
        {
            return (problem, 0);
        }
        await using var db = await factory.CreateDbContextAsync();
        if (!await db.SalesReps.AnyAsync(r => r.SalesRepID == repId && r.IsActive))
        {
            return ("اختر مندوبًا نشطًا.", 0);
        }
        if (customerId == 0 || !await db.Customers.AnyAsync(c => c.CustomerID == customerId))
        {
            return ("اختر العميل.", 0);
        }
        var d = day.Date;
        if (await db.RepVisits.AnyAsync(v => v.VisitDate == d && v.SalesRepID == repId && v.CustomerID == customerId))
        {
            return ("العميل في زيارات هذا اليوم.", 0);
        }
        short order = (short)((await db.RepVisits.Where(v => v.VisitDate == d && v.SalesRepID == repId)
                                   .MaxAsync(v => (short?)v.StopOrder) ?? 0) + 1);
        var visit = new RepVisits
        {
            VisitDate = d, SalesRepID = repId, CustomerID = customerId, IsPlanned = false, StopOrder = order,
            Result = "PENDING", EmployeeID = user.EmployeeID,
        };
        db.RepVisits.Add(visit);
        await db.SaveChangesAsync();
        return ("", visit.RepVisitID);
    }

    private static async Task<int> DocsOnDayAsync<T>(IQueryable<T> docs, Func<IQueryable<T>, IQueryable<T>> filter) =>
        await filter(docs).CountAsync();

    /// <summary>modRepRoutes.SuggestedVisitResult: what the customer did that day: a sale, else an order, else a payment.</summary>
    public static async Task<string> SuggestedResultAsync(RetailDbContext db, RepVisits v)
    {
        var from = v.VisitDate.Date;
        var to = from.AddDays(1);
        if (await DocsOnDayAsync(db.SalesInvoices, q => q.Where(x => x.CustomerID == v.CustomerID && x.InvoiceDate >= from && x.InvoiceDate < to)) > 0)
        {
            return "SALE";
        }
        if (await DocsOnDayAsync(db.SalesOrders, q => q.Where(x => x.CustomerID == v.CustomerID && x.OrderDate >= from && x.OrderDate < to)) > 0)
        {
            return "ORDER";
        }
        if (await DocsOnDayAsync(db.CustomerPayments, q => q.Where(x => x.CustomerID == v.CustomerID && x.PaymentDate >= from && x.PaymentDate < to)) > 0)
        {
            return "PAYMENT";
        }
        return "";
    }

    /// <summary>modRepRoutes.RecordRepVisit: the result (empty = the suggested one) and the notes, audited.
    /// repId limits the visit to that rep (0 = any).</summary>
    public async Task<string> RecordVisitAsync(ActingUser user, int visitId, string? result, string? notes, int repId = 0)
    {
        string problem = await CanRoutesAsync(user.EmployeeID);
        if (problem.Length > 0)
        {
            return problem;
        }
        await using var db = await factory.CreateDbContextAsync();
        var v = await db.RepVisits.FirstOrDefaultAsync(x => x.RepVisitID == visitId && (repId == 0 || x.SalesRepID == repId));
        if (v is null)
        {
            return "اختر الزيارة.";
        }
        string res = (result ?? "").Trim();
        if (res.Length == 0)
        {
            res = await SuggestedResultAsync(db, v);
        }
        if (res.Length == 0)
        {
            return "اختر نتيجة الزيارة: لا فاتورة ولا طلبية ولا تحصيل للعميل في هذا اليوم.";
        }
        if (!VisitResults.Contains(res))
        {
            return "نتيجة الزيارة غير صحيحة.";
        }
        var before = Audit.Snapshot("RepVisits", v);
        v.Result = res;
        v.VisitedAt = DateTime.Now;
        v.EmployeeID = user.EmployeeID;
        string text = (notes ?? "").Trim();
        if (text.Length > 0)
        {
            v.Notes = text.Length <= 255 ? text : text[..255];
        }
        await db.SaveChangesAsync();
        await WriteEditedAsync(db, user, "REP_VISIT_RECORD", "RepVisits", visitId, before, v);
        return "";
    }

    /// <summary>modRepRoutes.DeleteRepVisit: a visit not recorded yet, audited.</summary>
    public async Task<string> DeleteVisitAsync(ActingUser user, int visitId, int repId = 0)
    {
        string problem = await CanRoutesAsync(user.EmployeeID);
        if (problem.Length > 0)
        {
            return problem;
        }
        await using var db = await factory.CreateDbContextAsync();
        var v = await db.RepVisits.FirstOrDefaultAsync(x => x.RepVisitID == visitId && (repId == 0 || x.SalesRepID == repId));
        if (v is null || v.Result != "PENDING")
        {
            return "تُحذف الزيارة التي لم تُسجَّل نتيجتها فقط.";
        }
        var before = Audit.Snapshot("RepVisits", v);
        db.RepVisits.Remove(v);
        await db.SaveChangesAsync();
        await Audit.WriteAsync(db, user, "REP_VISIT_DELETE", "RepVisits", visitId.ToString(),
                               Audit.Label(before.Select(s => (s.Field, s.Value))),
                               before.Where(s => !Audit.IsEmpty(s.Value))
                                     .Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList());
        return "";
    }

    /// <summary>modAudit.AuditEdited: the fields that changed since the snapshot.</summary>
    private static async Task WriteEditedAsync(RetailDbContext db, ActingUser user, string action, string table, int id,
                                               List<(string Field, string Caption, object? Value)> before, object after)
    {
        var now = Audit.Snapshot(table, after);
        var changes = new List<AuditChange>();
        foreach (var (field, caption, value) in now)
        {
            var old = before.First(b => b.Field == field).Value;
            if (!Audit.SameValue(old, value))
            {
                changes.Add(new AuditChange(field, caption, Audit.ValueText(old), Audit.ValueText(value)));
            }
        }
        if (changes.Count > 0)
        {
            await Audit.WriteAsync(db, user, action, table, id.ToString(), Audit.Label(now.Select(s => (s.Field, s.Value))), changes);
        }
    }

    // ------------------------------------------------------------------ customer orders (modSalesOrders)

    /// <summary>modSalesOrders.SalesOrderLineProblem.</summary>
    public static string LineProblem(Products? p, decimal qty, decimal? price)
    {
        if (p is null)
        {
            return "اختر الصنف.";
        }
        if (!p.TrackStock)
        {
            return "هذا الصنف لا يتابع المخزون، فلا يُحجز ولا يُسلَّم بإذن.";
        }
        if (qty <= 0)
        {
            return "الكمية أكبر من صفر.";
        }
        return (price ?? 0) < 0 ? "السعر لا يكون سالبًا." : "";
    }

    /// <summary>modSalesReps.SalesRepFor outside a rep day: the active rep of the customer, else the active rep of the user.</summary>
    public static async Task<int?> SalesRepForAsync(RetailDbContext db, int customerId, int employeeId)
    {
        int? v = await (from c in db.Customers
                        join s in db.SalesReps on c.SalesRepID equals s.SalesRepID
                        where s.IsActive && c.CustomerID == customerId
                        select (int?)s.SalesRepID).FirstOrDefaultAsync();
        if (v is null)
        {
            int rep = await RepIdOfAsync(db, employeeId);
            v = rep == 0 ? null : rep;
        }
        return v;
    }

    /// <summary>modSalesOrders.CreateSalesOrder and its lines (SalesOrderLineBeforeUpdate) in one transaction: a DRAFT for
    /// an active customer in the user's warehouse (or the one given), number SO-, the rep of the customer; each line
    /// checked, a line without a price at OrderDefaultPrice. The office approves it in Access (the reservation).</summary>
    /// <summary>From the rep app (docs/78): clientRef = the number of the order on the phone. An order sent again with the
    /// same number by the same user is not added twice (Existing = true, its id); orderDate = when it was taken on the
    /// phone (never later than now).</summary>
    public async Task<(string Problem, int OrderID, bool Existing)> CreateOrderAsync(ActingUser user, int customerId, int? warehouseId,
                                                                                     DateTime? requiredDate, string? notes,
                                                                                     IReadOnlyList<OrderLineInput> lines,
                                                                                     string? clientRef = null, DateTime? orderDate = null)
    {
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "SALES_ORDERS"))
        {
            return ("ليست لديك صلاحية طلبيات العملاء.", 0, false);
        }
        await using var db = await factory.CreateDbContextAsync();
        string reference = (clientRef ?? "").Trim();
        if (reference.Length > 40)
        {
            return (ClientRefTooLong, 0, false);
        }
        if (reference.Length > 0)
        {
            var sent = await db.SalesOrders.Where(o => o.ClientRef == reference)
                .Select(o => new { o.SalesOrderID, o.EmployeeID }).FirstOrDefaultAsync();
            if (sent is not null)
            {
                return sent.EmployeeID == user.EmployeeID ? ("", sent.SalesOrderID, true) : (ClientRefUsed, 0, false);
            }
        }
        if (!await db.Customers.AnyAsync(c => c.IsActive && c.CustomerID == customerId))
        {
            return ("اختر العميل.", 0, false);
        }
        int warehouse = warehouseId ?? await UserWarehouseAsync(db, user.EmployeeID);
        if (!await db.Warehouses.AnyAsync(w => w.IsActive && w.WarehouseID == warehouse))
        {
            return ("اختر المخزن.", 0, false);
        }
        if (lines.Count == 0)
        {
            return ("أضف أسطر الطلبية.", 0, false);
        }
        var ids = lines.Select(l => l.ProductID).Distinct().ToList();
        var products = await db.Products.AsNoTracking().Where(p => ids.Contains(p.ProductID)).ToDictionaryAsync(p => p.ProductID);
        foreach (var l in lines)
        {
            string lineProblem = LineProblem(products.GetValueOrDefault(l.ProductID), l.Quantity, l.UnitPrice);
            if (lineProblem.Length > 0)
            {
                return (lineProblem, 0, false);
            }
        }
        int? rep = await SalesRepForAsync(db, customerId, user.EmployeeID);
        await using var tx = await db.Database.BeginTransactionAsync();
        string text = (notes ?? "").Trim();
        var order = new SalesOrders
        {
            OrderNumber = await NumberSequences.NextNumberAsync(db, "SALES_ORDER"),
            OrderDate = orderDate is { } taken && taken < DateTime.Now ? taken : DateTime.Now,
            CustomerID = customerId, WarehouseID = warehouse, SalesRepID = rep, RequiredDate = requiredDate?.Date,
            Status = "DRAFT", EmployeeID = user.EmployeeID, Notes = text.Length == 0 ? null : (text.Length <= 255 ? text : text[..255]),
            ClientRef = reference.Length == 0 ? null : reference,
        };
        db.SalesOrders.Add(order);
        await db.SaveChangesAsync();
        Audit.LogAction(db, user, "SALES_ORDER_NEW", "SalesOrders", order.SalesOrderID.ToString(),
                        reference.Length == 0 ? "api" : "api " + reference);
        await db.SaveChangesAsync();
        foreach (var l in lines)
        {
            decimal price = (l.UnitPrice ?? 0) == 0
                ? await OrderDefaultPriceAsync(db, customerId, rep, l.ProductID)
                : l.UnitPrice!.Value;
            var line = new SalesOrderLines { SalesOrderID = order.SalesOrderID, ProductID = l.ProductID, Quantity = l.Quantity, UnitPrice = price };
            db.SalesOrderLines.Add(line);
            await db.SaveChangesAsync();
            var name = products[l.ProductID].ProductName;
            await Audit.WriteAsync(db, user, "ADD", "SalesOrderLines", line.SalesOrderLineID.ToString(), "",
                [
                    new AuditChange("SalesOrderID", TableFields.Caption("SalesOrderLines", "SalesOrderID"), "", order.OrderNumber + " (" + order.SalesOrderID + ")"),
                    new AuditChange("ProductID", TableFields.Caption("SalesOrderLines", "ProductID"), "", name + " (" + l.ProductID + ")"),
                    new AuditChange("Quantity", TableFields.Caption("SalesOrderLines", "Quantity"), "", Audit.ValueText(l.Quantity)),
                    new AuditChange("UnitPrice", TableFields.Caption("SalesOrderLines", "UnitPrice"), "", Audit.ValueText(price)),
                ]);
        }
        await tx.CommitAsync();
        return ("", order.SalesOrderID, false);
    }

    /// <summary>The orders of a rep (SalesOrders.SalesRepID) or entered by the user, between two days, newest first,
    /// with their lines.</summary>
    public async Task<List<OrderView>> OrdersAsync(int repId, int employeeId, DateTime from, DateTime to, bool english, int orderId = 0)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var orders = await (from o in db.SalesOrders.AsNoTracking()
                            join c in db.Customers on o.CustomerID equals c.CustomerID
                            where (o.SalesRepID == repId || o.EmployeeID == employeeId) && (orderId != 0 ? o.SalesOrderID == orderId : o.OrderDate >= start && o.OrderDate < end)
                            orderby o.OrderDate descending, o.SalesOrderID descending
                            select new { o, c.CustomerName, c.CustomerNameEn }).ToListAsync();
        var ids = orders.Select(x => x.o.SalesOrderID).ToList();
        var lines = await (from l in db.SalesOrderLines.AsNoTracking()
                           join p in db.Products on l.ProductID equals p.ProductID
                           where ids.Contains(l.SalesOrderID)
                           orderby l.SalesOrderLineID
                           select new { l, p.ProductName, p.ProductNameEn }).ToListAsync();
        return orders.Select(x =>
        {
            var mine = lines.Where(l => l.l.SalesOrderID == x.o.SalesOrderID)
                .Select(l => new OrderLineView(l.l.SalesOrderLineID, l.l.ProductID, Name(l.ProductName, l.ProductNameEn, english),
                                               l.l.Quantity, l.l.UnitPrice)).ToArray();
            return new OrderView(x.o.SalesOrderID, x.o.OrderNumber, x.o.OrderDate, x.o.CustomerID,
                                 Name(x.CustomerName, x.CustomerNameEn, english), x.o.WarehouseID, x.o.SalesRepID, x.o.RequiredDate,
                                 x.o.Status, x.o.Notes, RoundMoney(mine.Sum(l => l.Quantity * l.UnitPrice)), mine, x.o.ClientRef);
        }).ToList();
    }
}
