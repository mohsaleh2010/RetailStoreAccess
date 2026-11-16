using System.Globalization;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Journal;
using RetailStore.Data.Reports;

namespace RetailStore.Data.Services;

/// <summary>The choices of the report centre (frmReportCenter): the period and the combo boxes.</summary>
public sealed record ReportChoices(DateTime? From = null, DateTime? To = null, int? CustomerId = null, int? SupplierId = null,
                                   int? ProductId = null, int? WarehouseId = null, int? CashBoxId = null, int? ExpenseTypeId = null);

/// <summary>The lists of the combo boxes of the report centre.</summary>
public sealed record ReportLists(List<(int Id, string Name)> Customers, List<(int Id, string Name)> Suppliers, List<(int Id, string Name)> Products,
                                 List<(int Id, string Name)> Warehouses, List<(int Id, string Name)> CashBoxes, List<(int Id, string Name)> ExpenseTypes);

/// <summary>A report run: its records in the order of the report, the line printed under its title (Criteria), the message
/// that stopped it, and a warning that did not (the journal update).</summary>
public sealed record ReportResult(ReportDef? Report, List<ReportRow> Rows, string Criteria, string Message, string Warning)
{
    public bool Ok => Message.Length == 0;

    public static ReportResult Fail(ReportDef? report, string message) => new(report, [], "", message, "");
}

/// <summary>The report centre of Access on the web (plan 6F, docs/82): modScreens.PrepareReport (its checks in their order and
/// its Arabic messages), then the saved query of the report (SavedQueryRunner: the views of SQL Server, docs/71) with the
/// parameters and filters of the choices, sorted as the Access report sorts it.</summary>
public sealed class ReportService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, JournalService journal)
{
    /// <summary>modScreens.ReportSelected: what the report asks for (lblNeeds).</summary>
    public static string NeedsHint(ReportDef? r)
    {
        if (r is null)
        {
            return "";
        }
        string hint = r.Dated ? "حدد الفترة" : "";
        if (r.Has('C')) hint += (hint.Length > 0 ? " و" : "") + "اختر العميل";
        if (r.Has('S')) hint += (hint.Length > 0 ? " و" : "") + "اختر المورد";
        if (r.Has('R')) hint += (hint.Length > 0 ? " و" : "") + "اختر المنتج";
        if (r.Has('c') || r.Has('s') || r.Has('r') || r.Has('b') || r.Has('e') || r.Has('W') || r.Has('w'))
        {
            hint += (hint.Length > 0 ? "، " : "") + "ويمكنك التصفية حسب الاختيار (اختياري)";
        }
        return hint.Length == 0 ? "لا يحتاج هذا التقرير أي اختيارات" : hint;
    }

    /// <summary>modReports.PeriodText.</summary>
    public static string PeriodText(DateTime from, DateTime to) =>
        "الفترة: من " + ReportRow.GDate(from, false) + " إلى " + ReportRow.GDate(to, false);

    public async Task<ReportResult> RunAsync(ActingUser user, string key, ReportChoices c, bool english)
    {
        var r = ReportCatalog.Find(key);
        if (r is null)
        {
            return ReportResult.Fail(null, "اختر تقريرًا من القائمة أولًا.");
        }
        int emp = user.EmployeeID;
        if (r.Has('$') && !await permissions.HasPermissionAsync(emp, "REPORTS_PROFIT"))
        {
            return ReportResult.Fail(r, "هذا التقرير يعرض التكلفة والأرباح ويحتاج صلاحية «تقارير الأرباح والضريبة».");
        }
        if (r.Has('A') && !await permissions.HasPermissionAsync(emp, "AUDIT_LOG"))
        {
            return ReportResult.Fail(r, "سجل التدقيق يحتاج صلاحية «سجل التدقيق».");
        }
        if (r.Has('#') && !await permissions.HasPermissionAsync(emp, "CASH_BOX"))
        {
            return ReportResult.Fail(r, "تقارير الخزينة تحتاج صلاحية «الخزينة».");
        }
        var parameters = new Dictionary<string, object?>();
        string warning = "";
        if (r.Has('J'))
        {
            if (!await permissions.HasPermissionAsync(emp, "JOURNAL"))
            {
                return ReportResult.Fail(r, "التقارير المحاسبية تحتاج صلاحية «قيود اليومية».");
            }
            warning = (await journal.SyncAsync(user)).Message;
            parameters["AccountCode"] = 0;              // general ledger from here: every account
        }
        var filters = new List<QueryFilter>();
        string criteria = "";
        DateTime from = default, to = default;
        if (r.Dated)
        {
            if (c.From is null || c.To is null)
            {
                return ReportResult.Fail(r, "أدخل تاريخ البداية وتاريخ النهاية.");
            }
            from = c.From.Value.Date;
            to = c.To.Value.Date;
            if (to < from)
            {
                return ReportResult.Fail(r, "تاريخ النهاية يجب أن يكون بعد تاريخ البداية.");
            }
            parameters["PeriodStart"] = from;               // SetPeriod: the end is the next day, exclusive
            parameters["PeriodEnd"] = to.AddDays(1);
            criteria = PeriodText(from, to);
            if (r.Key == "OPENING_BALANCES")
            {
                criteria = "في بداية يوم " + ReportRow.GDate(from, false);   // only the start counts
            }
        }
        if (r.Has('C') && c.CustomerId is null) return ReportResult.Fail(r, "اختر العميل.");
        if (r.Has('C')) parameters["CustomerID"] = c.CustomerId;
        if (r.Has('S') && c.SupplierId is null) return ReportResult.Fail(r, "اختر المورد.");
        if (r.Has('S')) parameters["SupplierID"] = c.SupplierId;
        if (r.Has('R') && c.ProductId is null) return ReportResult.Fail(r, "اختر المنتج.");
        if (r.Has('R')) parameters["ProductID"] = c.ProductId;

        if (r.Has('D'))
        {
            filters.Add(new QueryFilter(r.DateColumn, ">=", from));
            filters.Add(new QueryFilter(r.DateColumn, "<", to.AddDays(1)));
        }
        if (r.Has('c') && c.CustomerId is { } cu) filters.Add(new QueryFilter("CustomerID", "=", cu));
        if (r.Has('s') && c.SupplierId is { } su) filters.Add(new QueryFilter("SupplierID", "=", su));
        if (r.Has('r') && c.ProductId is { } pr) filters.Add(new QueryFilter("ProductID", "=", pr));
        if (r.Has('e') && c.ExpenseTypeId is { } et) filters.Add(new QueryFilter("ExpenseTypeID", "=", et));
        if (r.Has('w') && c.WarehouseId is { } wf) filters.Add(new QueryFilter("WarehouseID", "=", wf));
        if (r.Has('W')) parameters["WarehouseID"] = c.WarehouseId ?? 0;   // 0 = all
        if (r.Has('b')) parameters["CashBoxID"] = c.CashBoxId ?? 0;       // 0 = all boxes

        await using var db = await factory.CreateDbContextAsync();
        criteria = await CriteriaAsync(db, r, c, criteria, english);
        List<ReportRow> rows = [];
        if (r.Link.Length == 0)
        {
            rows = Sort(r, await SavedQueryRunner.RunAsync(db, r.Query, parameters, filters, english));
        }
        Audit.LogAction(db, user, "REPORT", r.Key);
        await db.SaveChangesAsync();
        return new ReportResult(r, rows, criteria, "", warning);
    }

    /// <summary>The sorting of the Access report (ListSpec.sorts); the order of the query stays for equal keys.</summary>
    public static List<ReportRow> Sort(ReportDef r, List<ReportRow> rows)
    {
        if (r.Sorts.Count == 0)
        {
            return rows;
        }
        var indexed = rows.Select((row, i) => (row, i)).ToList();
        indexed.Sort((a, b) =>
        {
            foreach (var (field, desc) in r.Sorts)
            {
                int k = ReportRow.Compare(a.row.Value(field), b.row.Value(field));
                if (k != 0)
                {
                    return desc ? -k : k;
                }
            }
            return a.i.CompareTo(b.i);
        });
        return indexed.Select(x => x.row).ToList();
    }

    /// <summary>The running balances of a report (RunningSum over the report), by column.</summary>
    public static decimal?[,] Running(ReportDef r, IReadOnlyList<ReportRow> rows)
    {
        var sums = new decimal?[rows.Count, r.Columns.Count];
        for (int col = 0; col < r.Columns.Count; col++)
        {
            if (!r.Columns[col].Running)
            {
                continue;
            }
            decimal total = 0;
            for (int i = 0; i < rows.Count; i++)
            {
                total += ReportRow.ToNumber(r.Columns[col].Value(rows[i]));
                sums[i, col] = total;
            }
        }
        return sums;
    }

    /// <summary>A cell of the report: the running balance, else the value of the column, as the report formats it.</summary>
    public static string CellText(ReportDef r, IReadOnlyList<ReportRow> rows, decimal?[,] running, int row, int col) =>
        running[row, col] is { } v ? ReportRow.Cell(v, r.Columns[col].Format)
                                   : ReportRow.Cell(r.Columns[col].Value(rows[row]), r.Columns[col].Format);

    /// <summary>The total of a column in the report footer (Sum).</summary>
    public static string TotalText(ReportDef r, IReadOnlyList<ReportRow> rows, int col) =>
        ReportRow.Cell(rows.Sum(x => ReportRow.ToNumber(r.Columns[col].Value(x))), r.Columns[col].Format);

    /// <summary>The report as CSV (Excel opens it in UTF-8 with its BOM): captions, then the cells as shown.</summary>
    public static string Csv(ReportDef r, IReadOnlyList<ReportRow> rows, Func<string, string> tr)
    {
        static string Q(string s) => "\"" + s.Replace("\"", "\"\"").Replace(' ', ' ') + "\"";
        var sb = new System.Text.StringBuilder();
        if (r.Card.Count > 0)
        {
            foreach (var row in rows.Take(1))
            {
                foreach (var cr in r.Card)
                {
                    sb.Append(Q(tr(cr.Caption))).Append(',').Append(Q(ReportRow.Cell(cr.Value(row), cr.Format))).Append("\r\n");
                }
            }
            return sb.ToString();
        }
        sb.Append(string.Join(",", r.Columns.Select(c => Q(tr(c.Caption))))).Append("\r\n");
        var running = Running(r, rows);
        for (int i = 0; i < rows.Count; i++)
        {
            sb.Append(string.Join(",", r.Columns.Select((_, col) => Q(CellText(r, rows, running, i, col))))).Append("\r\n");
        }
        return sb.ToString();
    }

    /// <summary>modScreens.AddCriteria: the chosen customer, supplier, product, box, expense type and warehouse.</summary>
    private static async Task<string> CriteriaAsync(RetailDbContext db, ReportDef r, ReportChoices c, string criteria, bool english)
    {
        void Add(bool used, string? name, string caption)
        {
            if (!used || string.IsNullOrEmpty(name))
            {
                return;
            }
            criteria += (criteria.Length > 0 ? "    " : "") + caption + ": " + name;
        }
        static string? Pick(string? ar, string? en, bool english) => english && !string.IsNullOrWhiteSpace(en) ? en : ar;
        if ((r.Has('C') || r.Has('c')) && c.CustomerId is { } cu)
        {
            var x = await db.Customers.AsNoTracking().Where(v => v.CustomerID == cu).Select(v => new { v.CustomerName, v.CustomerNameEn }).FirstOrDefaultAsync();
            Add(true, Pick(x?.CustomerName, x?.CustomerNameEn, english), "العميل");
        }
        if ((r.Has('S') || r.Has('s')) && c.SupplierId is { } su)
        {
            var x = await db.Suppliers.AsNoTracking().Where(v => v.SupplierID == su).Select(v => new { v.SupplierName, v.SupplierNameEn }).FirstOrDefaultAsync();
            Add(true, Pick(x?.SupplierName, x?.SupplierNameEn, english), "المورد");
        }
        if ((r.Has('R') || r.Has('r')) && c.ProductId is { } pr)
        {
            Add(true, await db.Products.AsNoTracking().Where(v => v.ProductID == pr).Select(v => v.ProductName).FirstOrDefaultAsync(), "المنتج");
        }
        if (r.Has('b'))
        {
            if (c.CashBoxId is { } bx)
            {
                var x = await db.CashBoxes.AsNoTracking().Where(v => v.CashBoxID == bx).Select(v => new { v.BoxName, v.BoxNameEn }).FirstOrDefaultAsync();
                Add(true, Pick(x?.BoxName, x?.BoxNameEn, english), "الصندوق");
            }
            else
            {
                criteria += "    كل الصناديق";
            }
        }
        if (r.Has('e') && c.ExpenseTypeId is { } et)
        {
            var x = await db.ExpenseTypes.AsNoTracking().Where(v => v.ExpenseTypeID == et).Select(v => new { v.ExpenseTypeName, v.ExpenseTypeNameEn }).FirstOrDefaultAsync();
            Add(true, Pick(x?.ExpenseTypeName, x?.ExpenseTypeNameEn, english), "نوع المصروف");
        }
        if (r.Has('W') || r.Has('w'))
        {
            if (c.WarehouseId is { } wh)
            {
                var x = await db.Warehouses.AsNoTracking().Where(v => v.WarehouseID == wh).Select(v => new { v.WarehouseName, v.WarehouseNameEn }).FirstOrDefaultAsync();
                Add(true, Pick(x?.WarehouseName, x?.WarehouseNameEn, english), "المخزن");
            }
            else
            {
                criteria += "    كل المخازن";
            }
        }
        return criteria.Trim();
    }

    /// <summary>The combo boxes of frmReportCenter: active customers, suppliers, products, warehouses, cash boxes, expense types.</summary>
    public async Task<ReportLists> ListsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        static string Pick(string ar, string? en, bool english) => english && !string.IsNullOrWhiteSpace(en) ? en : ar;
        var customers = (await db.Customers.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.CustomerName)
            .Select(x => new { x.CustomerID, x.CustomerName, x.CustomerNameEn }).ToListAsync())
            .Select(x => (x.CustomerID, Pick(x.CustomerName, x.CustomerNameEn, english))).ToList();
        var suppliers = (await db.Suppliers.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.SupplierName)
            .Select(x => new { x.SupplierID, x.SupplierName, x.SupplierNameEn }).ToListAsync())
            .Select(x => (x.SupplierID, Pick(x.SupplierName, x.SupplierNameEn, english))).ToList();
        var products = (await db.Products.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.ProductName)
            .Select(x => new { x.ProductID, x.ProductCode, x.ProductName }).ToListAsync())
            .Select(x => (x.ProductID, x.ProductCode + " - " + x.ProductName)).ToList();
        var warehouses = (await db.Warehouses.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.WarehouseID)
            .Select(x => new { x.WarehouseID, x.WarehouseName, x.WarehouseNameEn }).ToListAsync())
            .Select(x => (x.WarehouseID, Pick(x.WarehouseName, x.WarehouseNameEn, english))).ToList();
        var boxes = (await db.CashBoxes.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.CashBoxID)
            .Select(x => new { x.CashBoxID, x.BoxName, x.BoxNameEn }).ToListAsync())
            .Select(x => (x.CashBoxID, Pick(x.BoxName, x.BoxNameEn, english))).ToList();
        var types = (await db.ExpenseTypes.AsNoTracking().Where(x => x.IsActive).OrderBy(x => x.ExpenseTypeName)
            .Select(x => new { x.ExpenseTypeID, x.ExpenseTypeName, x.ExpenseTypeNameEn }).ToListAsync())
            .Select(x => (x.ExpenseTypeID, Pick(x.ExpenseTypeName, x.ExpenseTypeNameEn, english))).ToList();
        return new ReportLists(customers, suppliers, products, warehouses, boxes, types);
    }

    /// <summary>modFiscal.FiscalPeriodOf: the fiscal period of a day, else its calendar year (the quick periods of the centre).</summary>
    public async Task<(DateTime Start, DateTime End)> FiscalPeriodOfAsync(DateTime day)
    {
        await using var db = await factory.CreateDbContextAsync();
        var d = day.Date;
        var p = await db.FiscalPeriods.AsNoTracking().Where(x => x.StartDate <= d && x.EndDate >= d)
            .Select(x => new { x.StartDate, x.EndDate }).FirstOrDefaultAsync();
        return p is null ? (new DateTime(d.Year, 1, 1), new DateTime(d.Year, 12, 31)) : (p.StartDate.Date, p.EndDate.Date);
    }

    public static string DateText(DateTime d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
}
