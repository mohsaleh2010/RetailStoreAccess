using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Domain.Localization;

namespace RetailStore.Data.Services;

/// <summary>A cash movement of a box (qryCashMovements): in (+) or out (-).</summary>
public sealed record CashMove(int CashBoxID, DateTime MoveDate, string MoveType, string MoveTypeName, string? DocNumber, string? PartyName,
                              string? Details, decimal AmountIn, decimal AmountOut);

/// <summary>A box and its balance now (CashBoxBalanceQuery).</summary>
public sealed record BoxBalance(int CashBoxID, string BoxName, string BoxType, string BoxTypeName, bool IsActive, decimal TotalIn,
                                decimal TotalOut, decimal Balance, DateTime? LastMoveDate);

/// <summary>A line of the statement of a box (CashStatementQuery): the opening line first (BALANCE_FWD), then the moves with
/// the balance after each.</summary>
public sealed record BoxStatementLine(DateTime MoveDate, string MoveType, string MoveTypeName, string? DocNumber, string? PartyName,
                                      string? Details, string BoxName, decimal AmountIn, decimal AmountOut, decimal Balance);

/// <summary>A cash voucher (modCash.PostCashVoucher): VoucherType IN / OUT / TRANSFER; Category of an IN voucher OTHER / OWNER,
/// of an OUT voucher EXPENSE (with ExpenseTypeID) / OWNER / OTHER; ToCashBoxID for a transfer. The advance of an employee and
/// the commission of a rep (ADVANCE / COMMISSION, modPayroll / modSalesReps) stay in Access.</summary>
public sealed record CashVoucherRequest(string VoucherType, int CashBoxID, int? ToCashBoxID, string? Category, decimal Amount,
                                        string? PartyName, string? Description, int? ExpenseTypeID);

public sealed record CashVoucherRow(int CashVoucherID, string VoucherNumber, DateTime VoucherDate, string VoucherType, string Category,
                                    string BoxName, string? ToBoxName, decimal Amount, string? PartyName, string? Description,
                                    bool InTransit, DateTime? ReceivedAt, string EmployeeName);

/// <summary>A bank transaction (modBank.PostBankTx): TxType DEPOSIT / WITHDRAW / SETTLEMENT / TRANSFER / OTHER_IN / OTHER_OUT.</summary>
public sealed record BankTxRequest(string TxType, DateTime? TxDate, int BankID, int? ToBankID, int? CashBoxID, int? CounterAccount,
                                   decimal Amount, decimal FeeAmount, decimal FeeVAT, string? Reference, string? Description);

public sealed record BankTxRow(int BankTxID, string TxNumber, DateTime TxDate, string TxType, string TxTypeName, string BankName,
                               string? OtherName, decimal Amount, decimal FeeAmount, decimal FeeVAT, string? Reference,
                               string? Description, string EmployeeName);

/// <summary>A bank and its balance in the books (BankBalanceQuery: its journal lines, account 120000 + bank).</summary>
public sealed record BankBalance(int BankID, string BankName, string? AccountNo, string? IBAN, bool IsActive, decimal Balance,
                                 DateTime? LastItemDate);

public sealed record TreasuryResult(SaveStatus Status, string Message, int Id = 0, string Number = "")
{
    public static TreasuryResult Refused(string message) => new(SaveStatus.Refused, message);
}

/// <summary>The treasury and the banks on the web and the API (plan 6E, docs/81): the cash movements of each box
/// (qryCashMovements), the balances and the statement of a box, the cash vouchers (modCash.PostCashVoucher, with the expense of
/// an expense voucher), the transfers between boxes waiting for their receipt (ReceiveCashTransfer / CancelCashTransfer), the
/// bank balances in the books and the bank transactions (modBank.PostBankTx / DeleteBankTx), with the checks, messages, rows
/// and audit of Access. The cash closing and the bank reconciliation stay in Access.</summary>
public sealed class TreasuryService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, JournalService? journal = null)
{
    public const string VoucherForm = "frmCashVoucher";
    public const string BankTxForm = "frmBankTx";
    public static readonly string[] InCategories = ["OTHER", "OWNER"];
    public static readonly string[] OutCategories = ["EXPENSE", "OWNER", "OTHER"];
    public static readonly string[] BankTxTypes = ["DEPOSIT", "WITHDRAW", "SETTLEMENT", "TRANSFER", "OTHER_IN", "OTHER_OUT"];

    private SalesService Sales => new(factory, permissions);

    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    private static string Tr(string arabic, bool english) => english ? Translator.ToEnglish(arabic) : arabic;

    /// <summary>Format$(x, "#,##0.00").</summary>
    private static string Money(decimal v) => v.ToString("#,##0.00", System.Globalization.CultureInfo.InvariantCulture);

    // ------------------------------------------------------------------ the cash movements (qryCashMovements)

    /// <summary>qryCashMovements: every cash movement of the boxes, one branch per document, as in tools/queries.py.</summary>
    public static async Task<List<CashMove>> CashMovesAsync(RetailDbContext db, int boxId, bool english)
    {
        var list = new List<CashMove>();
        bool all = boxId == 0;
        var customers = await db.Customers.AsNoTracking().Select(c => new { c.CustomerID, c.CustomerName, c.CustomerNameEn })
            .ToDictionaryAsync(c => c.CustomerID, c => Name(c.CustomerName, c.CustomerNameEn, english));
        var suppliers = await db.Suppliers.AsNoTracking().Select(s => new { s.SupplierID, s.SupplierName, s.SupplierNameEn })
            .ToDictionaryAsync(s => s.SupplierID, s => Name(s.SupplierName, s.SupplierNameEn, english));
        var boxes = await db.CashBoxes.AsNoTracking().ToDictionaryAsync(b => b.CashBoxID);
        var banks = await db.Banks.AsNoTracking().ToDictionaryAsync(b => b.BankID, b => Name(b.BankName, b.BankNameEn, english));
        var types = await db.ExpenseTypes.AsNoTracking()
            .ToDictionaryAsync(t => t.ExpenseTypeID, t => Name(t.ExpenseTypeName, t.ExpenseTypeNameEn, english));
        string C(int id) => customers.TryGetValue(id, out var n) ? n : "";
        string S(int id) => suppliers.TryGetValue(id, out var n) ? n : "";
        string B(int id) => boxes.TryGetValue(id, out var b) ? Name(b.BoxName, b.BoxNameEn, english) : "";
        void Add(int box, DateTime d, string type, string typeName, string? no, string? party, string? details, decimal amountIn, decimal amountOut)
        {
            list.Add(new CashMove(box, d, type, Tr(typeName, english), no, party, details, amountIn, amountOut));
        }

        foreach (var h in await db.SalesInvoices.AsNoTracking().Where(h => h.CashBoxID != null && h.PaidAmount != 0 && (all || h.CashBoxID == boxId))
                     .Select(h => new { Box = h.CashBoxID!.Value, h.InvoiceDate, h.InvoiceNumber, h.CustomerID, h.Notes, h.PaidAmount }).ToListAsync())
        {
            Add(h.Box, h.InvoiceDate, "SALE", "فاتورة بيع", h.InvoiceNumber, C(h.CustomerID), h.Notes, h.PaidAmount, 0);
        }
        foreach (var r in await db.SalesReturns.AsNoTracking().Where(r => r.CashBoxID != null && r.RefundedAmount != 0 && (all || r.CashBoxID == boxId))
                     .Select(r => new { Box = r.CashBoxID!.Value, r.ReturnDate, r.ReturnNumber, r.CustomerID, r.Reason, r.RefundedAmount }).ToListAsync())
        {
            Add(r.Box, r.ReturnDate, "SALES_RETURN", "مرتجع بيع (رد نقدي)", r.ReturnNumber, C(r.CustomerID), r.Reason, 0, r.RefundedAmount);
        }
        foreach (var p in await db.CustomerPayments.AsNoTracking().Where(p => p.CashBoxID != null && (all || p.CashBoxID == boxId))
                     .Select(p => new { Box = p.CashBoxID!.Value, p.PaymentDate, p.PaymentNumber, p.CustomerID, p.Notes, p.Amount }).ToListAsync())
        {
            Add(p.Box, p.PaymentDate, "CUSTOMER_PAYMENT", "سند قبض من عميل", p.PaymentNumber, C(p.CustomerID), p.Notes, p.Amount, 0);
        }
        foreach (var h in await db.PurchaseInvoices.AsNoTracking().Where(h => h.CashBoxID != null && h.PaidAmount != 0 && (all || h.CashBoxID == boxId))
                     .Select(h => new { Box = h.CashBoxID!.Value, h.InvoiceDate, h.InvoiceNumber, h.SupplierID, h.Notes, h.PaidAmount }).ToListAsync())
        {
            Add(h.Box, h.InvoiceDate, "PURCHASE", "فاتورة شراء", h.InvoiceNumber, S(h.SupplierID), h.Notes, 0, h.PaidAmount);
        }
        foreach (var r in await db.PurchaseReturns.AsNoTracking().Where(r => r.CashBoxID != null && r.RefundedAmount != 0 && (all || r.CashBoxID == boxId))
                     .Select(r => new { Box = r.CashBoxID!.Value, r.ReturnDate, r.ReturnNumber, r.SupplierID, r.Reason, r.RefundedAmount }).ToListAsync())
        {
            Add(r.Box, r.ReturnDate, "PURCHASE_RETURN", "مرتجع شراء (استرداد نقدي)", r.ReturnNumber, S(r.SupplierID), r.Reason, r.RefundedAmount, 0);
        }
        foreach (var p in await db.SupplierPayments.AsNoTracking().Where(p => p.CashBoxID != null && (all || p.CashBoxID == boxId))
                     .Select(p => new { Box = p.CashBoxID!.Value, p.PaymentDate, p.PaymentNumber, p.SupplierID, p.Notes, p.Amount }).ToListAsync())
        {
            Add(p.Box, p.PaymentDate, "SUPPLIER_PAYMENT", "سند صرف لمورد", p.PaymentNumber, S(p.SupplierID), p.Notes, 0, p.Amount);
        }
        foreach (var e in await db.Expenses.AsNoTracking().Where(e => e.CashBoxID != null && (all || e.CashBoxID == boxId))
                     .Select(e => new { Box = e.CashBoxID!.Value, e.ExpenseDate, e.ExpenseNumber, e.ExpenseTypeID, e.Description, e.TotalAmount }).ToListAsync())
        {
            Add(e.Box, e.ExpenseDate, "EXPENSE", "مصروف", e.ExpenseNumber, types.TryGetValue(e.ExpenseTypeID, out var t) ? t : "", e.Description, 0,
                e.TotalAmount);
        }
        foreach (var v in await db.CashVouchers.AsNoTracking().Where(v => all || v.CashBoxID == boxId || v.ToCashBoxID == boxId).ToListAsync())
        {
            if (v.VoucherType == "IN" && (all || v.CashBoxID == boxId))
            {
                string name = v.Category == "OWNER" ? "إيداع من المالك" : v.Category == "OVERAGE" ? "زيادة في الصندوق" : "سند قبض نقدية";
                Add(v.CashBoxID, v.VoucherDate, "CASH_IN", name, v.VoucherNumber, v.PartyName, v.Description, v.Amount, 0);
            }
            else if (v.VoucherType == "OUT" && (all || v.CashBoxID == boxId))
            {
                string name = v.Category switch
                {
                    "OWNER" => "تسوية مع المالك",
                    "EXPENSE" => "مصروف (سند صرف)",
                    "ADVANCE" => "سلفة موظف",
                    "SHORTAGE" => "عجز في الصندوق",
                    _ => "سند صرف نقدية",
                };
                Add(v.CashBoxID, v.VoucherDate, "CASH_OUT", name, v.VoucherNumber, v.PartyName, v.Description, 0, v.Amount);
            }
            else if (v.VoucherType == "TRANSFER" && v.ToCashBoxID is { } to)
            {
                if (all || v.CashBoxID == boxId)
                {
                    Add(v.CashBoxID, v.VoucherDate, "TRANSFER_OUT", "تحويل إلى صندوق آخر", v.VoucherNumber, B(to), v.Description, 0, v.Amount);
                }
                if ((all || to == boxId) && (!v.InTransit || v.ReceivedAt is not null))
                {
                    Add(to, v.InTransit ? v.ReceivedAt!.Value : v.VoucherDate, "TRANSFER_IN", "تحويل من صندوق آخر", v.VoucherNumber, B(v.CashBoxID),
                        v.Description, v.Amount, 0);
                }
            }
        }
        foreach (var t in await db.BankTransactions.AsNoTracking()
                     .Where(t => (t.TxType == "DEPOSIT" || t.TxType == "WITHDRAW") && t.CashBoxID != null && (all || t.CashBoxID == boxId)).ToListAsync())
        {
            string bank = banks.TryGetValue(t.BankID, out var k) ? k : "";
            if (t.TxType == "DEPOSIT")
            {
                Add(t.CashBoxID!.Value, t.TxDate, "BANK_DEPOSIT", "إيداع في البنك", t.TxNumber, bank, t.Description, 0, t.Amount);
            }
            else
            {
                Add(t.CashBoxID!.Value, t.TxDate, "BANK_WITHDRAW", "سحب من البنك", t.TxNumber, bank, t.Description, t.Amount, 0);
            }
        }
        foreach (var r in await db.PayrollRuns.AsNoTracking()
                     .Where(r => r.Status == "POSTED" && r.PaidFrom == "CASHBOX" && r.PaidAmount != 0 && r.CashBoxID != null && r.PaidDate != null &&
                                 (all || r.CashBoxID == boxId)).ToListAsync())
        {
            Add(r.CashBoxID!.Value, r.PaidDate!.Value, "PAYROLL", "صرف الرواتب", r.RunNumber, "-", r.Notes, 0, r.PaidAmount);
        }
        foreach (var a in await db.FixedAssets.AsNoTracking().Where(a => (a.SourceType == "CASHBOX" && a.CashBoxID != null && (all || a.CashBoxID == boxId)) ||
                                                                         (a.Status == "DISPOSED" && a.DisposalTo == "CASHBOX" && a.DisposalProceeds != 0 &&
                                                                          a.DisposalCashBoxID != null && (all || a.DisposalCashBoxID == boxId))).ToListAsync())
        {
            if (a.SourceType == "CASHBOX" && a.CashBoxID is { } buyBox && (all || buyBox == boxId))
            {
                Add(buyBox, a.PurchaseDate, "ASSET", "شراء أصل ثابت", a.AssetCode, a.AssetName, a.Notes, 0, a.Cost + a.InputVAT);
            }
            if (a.Status == "DISPOSED" && a.DisposalTo == "CASHBOX" && a.DisposalProceeds != 0 && a.DisposalCashBoxID is { } sellBox &&
                a.DisposalDate is { } sold && (all || sellBox == boxId))
            {
                Add(sellBox, sold, "ASSET_SALE", "بيع أصل ثابت", a.AssetCode, a.AssetName, a.Notes, a.DisposalProceeds + a.DisposalVAT, 0);
            }
        }
        foreach (var b in boxes.Values.Where(b => b.OpeningBalance != 0 && (all || b.CashBoxID == boxId)))
        {
            Add(b.CashBoxID, b.OpeningDate, "OPENING", "رصيد افتتاحي", "-", Name(b.BoxName, b.BoxNameEn, english), b.Notes, b.OpeningBalance, 0);
        }
        return list;
    }

    /// <summary>modCash.CashBoxBalance: the balance now, or of the movements before a moment.</summary>
    public static async Task<decimal> CashBoxBalanceAsync(RetailDbContext db, int boxId, DateTime? before = null) =>
        (await CashMovesAsync(db, boxId, false)).Where(m => before is null || m.MoveDate < before).Sum(m => m.AmountIn - m.AmountOut);

    /// <summary>CashBoxBalanceQuery: every box with its totals and balance (main safe first, as in Access).</summary>
    public async Task<List<BoxBalance>> BoxBalancesAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var moves = await CashMovesAsync(db, 0, english);
        var totals = moves.GroupBy(m => m.CashBoxID).ToDictionary(g => g.Key, g => (In: g.Sum(m => m.AmountIn), Out: g.Sum(m => m.AmountOut),
                                                                                       Last: g.Max(m => m.MoveDate)));
        var boxes = await db.CashBoxes.AsNoTracking().ToListAsync();
        return boxes.OrderByDescending(b => b.BoxType, StringComparer.Ordinal).ThenBy(b => b.BoxName)
            .Select(b =>
            {
                var t = totals.TryGetValue(b.CashBoxID, out var x) ? x : (In: 0m, Out: 0m, Last: (DateTime?)null);
                string typeName = b.BoxType == "MAIN" ? "خزينة رئيسية" : b.BoxType == "CUSTODY" ? "عهدة مندوب" : "صندوق كاشير";
                return new BoxBalance(b.CashBoxID, Name(b.BoxName, b.BoxNameEn, english), b.BoxType, Tr(typeName, english), b.IsActive, t.In, t.Out,
                                      t.In - t.Out, t.Last);
            }).ToList();
    }

    /// <summary>CashStatementQuery: the opening balance of the period, then each movement of the period with the balance after it
    /// (box 0 = every box).</summary>
    public async Task<List<BoxStatementLine>> BoxStatementAsync(int boxId, DateTime from, DateTime to, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var moves = await CashMovesAsync(db, boxId, english);
        var boxes = await db.CashBoxes.AsNoTracking().ToDictionaryAsync(b => b.CashBoxID, b => Name(b.BoxName, b.BoxNameEn, english));
        decimal opening = moves.Where(m => m.MoveDate < start).Sum(m => m.AmountIn - m.AmountOut);
        var list = new List<BoxStatementLine>
        {
            new(start, "BALANCE_FWD", Tr("رصيد أول المدة", english), "-", null, null, boxId > 0 && boxes.TryGetValue(boxId, out var bn) ? bn : "",
                opening > 0 ? opening : 0, opening < 0 ? -opening : 0, opening),
        };
        decimal balance = opening;
        foreach (var m in moves.Where(m => m.MoveDate >= start && m.MoveDate < end).OrderBy(m => m.MoveDate))
        {
            balance += m.AmountIn - m.AmountOut;
            list.Add(new BoxStatementLine(m.MoveDate, m.MoveType, m.MoveTypeName, m.DocNumber, m.PartyName, m.Details,
                                          boxes.TryGetValue(m.CashBoxID, out var n) ? n : "", m.AmountIn, m.AmountOut, balance));
        }
        return list;
    }

    // ------------------------------------------------------------------ box custodians (docs/56)

    /// <summary>modCash.IsBoxCustodian: the box of the user, a box he keeps, or the custody box of his sales rep.</summary>
    public static async Task<bool> IsBoxCustodianAsync(RetailDbContext db, int employeeId, int boxId) =>
        await db.Employees.AnyAsync(e => e.EmployeeID == employeeId && e.CashBoxID == boxId)
        || await db.CashBoxUsers.AnyAsync(u => u.CashBoxID == boxId && u.EmployeeID == employeeId)
        || await (from b in db.CashBoxes
                  join s in db.SalesReps on b.SalesRepID equals s.SalesRepID
                  where b.CashBoxID == boxId && s.EmployeeID == employeeId
                  select b).AnyAsync();

    /// <summary>modCash.BoxHasCustodians.</summary>
    public static async Task<bool> BoxHasCustodiansAsync(RetailDbContext db, int boxId) =>
        await db.CashBoxUsers.AnyAsync(u => u.CashBoxID == boxId)
        || await db.Employees.AnyAsync(e => e.IsActive && e.CashBoxID == boxId)
        || await (from b in db.CashBoxes
                  join s in db.SalesReps on b.SalesRepID equals s.SalesRepID
                  where b.CashBoxID == boxId && s.EmployeeID != null
                  select b).AnyAsync();

    /// <summary>modCash.CanUseBox: every box with CASH_ALL_BOXES; else the boxes the user keeps and the boxes nobody keeps (the
    /// web works outside a rep day).</summary>
    public async Task<bool> CanUseBoxAsync(RetailDbContext db, int employeeId, int boxId)
    {
        if (await permissions.HasPermissionAsync(employeeId, "CASH_ALL_BOXES"))
        {
            return true;
        }
        return await IsBoxCustodianAsync(db, employeeId, boxId) || !await BoxHasCustodiansAsync(db, boxId);
    }

    private static async Task<string> BoxNameAsync(RetailDbContext db, int? boxId) =>
        boxId is { } id ? await db.CashBoxes.Where(b => b.CashBoxID == id).Select(b => b.BoxName).FirstOrDefaultAsync() ?? "" : "";

    /// <summary>modCash.BoxUseProblem.</summary>
    public async Task<string> BoxUseProblemAsync(RetailDbContext db, int employeeId, int boxId) =>
        await CanUseBoxAsync(db, employeeId, boxId) ? "" : "لست من أمناء " + await BoxNameAsync(db, boxId) + ": لا يمكنك الصرف أو التحويل منه.";

    private static async Task<bool> BoxIsActiveAsync(RetailDbContext db, int? boxId) =>
        boxId is { } id && await db.CashBoxes.AnyAsync(b => b.IsActive && b.CashBoxID == id);

    // ------------------------------------------------------------------ cash vouchers

    /// <summary>modCash.VoucherProblem: the checks of a voucher entered by the user ("" = fine).</summary>
    public static async Task<string> VoucherProblemAsync(RetailDbContext db, CashVoucherRequest v)
    {
        if (v.Amount <= 0)
        {
            return "أدخل المبلغ (أكبر من صفر).";
        }
        if (!await BoxIsActiveAsync(db, v.CashBoxID))
        {
            return "اختر الصندوق.";
        }
        string category = v.Category ?? "";
        switch (v.VoucherType)
        {
            case "TRANSFER":
                if (!await BoxIsActiveAsync(db, v.ToCashBoxID))
                {
                    return "اختر الصندوق المحوَّل إليه.";
                }
                return v.ToCashBoxID == v.CashBoxID ? "لا يمكن التحويل إلى نفس الصندوق." : "";
            case "IN":
                return InCategories.Contains(category) ? "" : "اختر بند القبض.";
            case "OUT":
                if (!OutCategories.Contains(category))
                {
                    return "اختر بند الصرف.";
                }
                return category == "EXPENSE" && v.ExpenseTypeID is null ? "اختر نوع المصروف." : "";
            default:
                return "نوع السند غير معروف: " + v.VoucherType;
        }
    }

    /// <summary>modCash.PostCashVoucher (and InsertVoucher): the permission, VoucherProblem, the box of a payment or transfer, its
    /// balance, then the voucher (and the expense of an expense voucher) in one transaction, then the audit row.</summary>
    public async Task<TreasuryResult> PostCashVoucherAsync(ActingUser user, CashVoucherRequest v)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, VoucherForm);
        if (open.Length > 0)
        {
            return TreasuryResult.Refused(open);
        }
        string add = await permissions.ActionProblemAsync(user.EmployeeID, VoucherForm, ScreenAction.Add);
        if (add.Length > 0)
        {
            return TreasuryResult.Refused(add);
        }
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "CASH_BOX"))
        {
            return TreasuryResult.Refused("ليست لديك صلاحية سندات الخزينة.");
        }
        await using var db = await factory.CreateDbContextAsync();
        string msg = await VoucherProblemAsync(db, v);
        if (msg.Length == 0 && v.VoucherType != "IN")
        {
            msg = await BoxUseProblemAsync(db, user.EmployeeID, v.CashBoxID);
        }
        if (msg.Length == 0 && v.VoucherType != "IN")
        {
            decimal balance = await CashBoxBalanceAsync(db, v.CashBoxID);
            if (v.Amount > balance)
            {
                msg = "رصيد " + await BoxNameAsync(db, v.CashBoxID) + " (" + Money(balance) + ") لا يكفي لصرف " + Money(v.Amount) + ".";
            }
        }
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        var now = DateTime.Now;
        string category = v.VoucherType == "TRANSFER" ? "TRANSFER" : v.Category!;
        await using var tx = await db.Database.BeginTransactionAsync();
        string seq = v.VoucherType switch { "IN" => "CASH_IN", "OUT" => "CASH_OUT", _ => "CASH_TRANSFER" };
        string number = await NumberSequences.NextNumberAsync(db, seq);
        int? center = await SalesService.CostCenterForAsync(db, user.EmployeeID);
        string description = (v.Description ?? "").Trim();
        Expenses? expense = null;
        if (v.VoucherType == "OUT" && category == "EXPENSE")
        {
            // the voucher moves the cash; the expense (without a box) feeds the expense reports
            expense = new Expenses
            {
                ExpenseNumber = await NumberSequences.NextNumberAsync(db, "EXPENSE"), ExpenseDate = now.Date, ExpenseTypeID = v.ExpenseTypeID!.Value,
                Amount = v.Amount, Tax = 0, TotalAmount = v.Amount, PaymentMethodID = SalesService.CashMethodId,
                Description = SalesService.Clip("سند صرف نقدية " + number + (description.Length > 0 ? " - " + description : ""), 255),
                EmployeeID = user.EmployeeID, CostCenterID = center,
            };
            db.Expenses.Add(expense);
            await db.SaveChangesAsync();
        }
        bool inTransit = v.VoucherType == "TRANSFER" &&
                         await db.CashBoxes.Where(b => b.CashBoxID == v.ToCashBoxID).Select(b => b.ConfirmReceipt).FirstOrDefaultAsync();
        string party = (v.PartyName ?? "").Trim();
        var voucher = new CashVouchers
        {
            VoucherNumber = number, VoucherDate = now, VoucherType = v.VoucherType, CashBoxID = v.CashBoxID,
            ToCashBoxID = v.VoucherType == "TRANSFER" ? v.ToCashBoxID : null, Category = category, InTransit = inTransit, Amount = v.Amount,
            PartyName = party.Length > 0 ? SalesService.Clip(party, 100) : null,
            Description = description.Length > 0 ? SalesService.Clip(description, 255) : null, ExpenseID = expense?.ExpenseID,
            EmployeeID = user.EmployeeID, CostCenterID = center,
        };
        db.CashVouchers.Add(voucher);
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        Audit.LogAction(db, user, "CASH_" + v.VoucherType, "CashVouchers", voucher.CashVoucherID.ToString(),
                        category + " Amount=" + Domain.Sales.InvoiceCalculator.Plain(v.Amount));
        await db.SaveChangesAsync();
        return new TreasuryResult(SaveStatus.Saved, "", voucher.CashVoucherID, number);
    }

    /// <summary>The cash vouchers of a period, newest first; inTransitOnly = the transfers waiting for their receipt.</summary>
    public async Task<List<CashVoucherRow>> VouchersAsync(DateTime from, DateTime to, int boxId, bool inTransitOnly, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var q = db.CashVouchers.AsNoTracking().AsQueryable();
        q = inTransitOnly ? q.Where(v => v.VoucherType == "TRANSFER" && v.InTransit && v.ReceivedAt == null)
                          : q.Where(v => v.VoucherDate >= start && v.VoucherDate < end);
        if (boxId > 0)
        {
            q = q.Where(v => v.CashBoxID == boxId || v.ToCashBoxID == boxId);
        }
        var rows = await q.OrderByDescending(v => v.VoucherDate).ThenByDescending(v => v.CashVoucherID).Take(take).ToListAsync();
        var boxes = await db.CashBoxes.AsNoTracking().ToDictionaryAsync(b => b.CashBoxID, b => Name(b.BoxName, b.BoxNameEn, english));
        var ids = rows.Select(r => r.EmployeeID).Distinct().ToList();
        var users = await db.Employees.AsNoTracking().Where(e => ids.Contains(e.EmployeeID)).ToDictionaryAsync(e => e.EmployeeID, e => e.EmployeeName);
        return rows.Select(v => new CashVoucherRow(v.CashVoucherID, v.VoucherNumber, v.VoucherDate, v.VoucherType, v.Category,
                                                   boxes.TryGetValue(v.CashBoxID, out var b) ? b : "",
                                                   v.ToCashBoxID is { } t && boxes.TryGetValue(t, out var tb) ? tb : null, v.Amount, v.PartyName,
                                                   v.Description, v.InTransit, v.ReceivedAt, users.TryGetValue(v.EmployeeID, out var u) ? u : "")).ToList();
    }

    private static Task<CashVouchers?> PendingTransferAsync(RetailDbContext db, int voucherId) =>
        db.CashVouchers.FirstOrDefaultAsync(v => v.CashVoucherID == voucherId && v.VoucherType == "TRANSFER" && v.InTransit && v.ReceivedAt == null);

    /// <summary>modCash.ReceiveCashTransfer: a custodian of the receiving box confirms the cash arrived: it enters the box now.</summary>
    public async Task<TreasuryResult> ReceiveTransferAsync(ActingUser user, int voucherId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var v = await PendingTransferAsync(db, voucherId);
        string msg;
        if (v is null)
        {
            msg = "هذا التحويل ليس بانتظار الاستلام.";
        }
        else if (!await permissions.HasPermissionAsync(user.EmployeeID, "CASH_CLOSING") && !await permissions.HasPermissionAsync(user.EmployeeID, "CASH_BOX"))
        {
            msg = "ليست لديك صلاحية الخزينة.";
        }
        else if (!await CanUseBoxAsync(db, user.EmployeeID, v.ToCashBoxID!.Value))
        {
            msg = "يستلم التحويل أمين " + await BoxNameAsync(db, v.ToCashBoxID) + ".";
        }
        else
        {
            msg = StockService.ClosedPeriodProblem(DateTime.Today, (await SettingsReader.ReadAsync(db)).ClosedThrough);
        }
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        var before = Audit.Snapshot("CashVouchers", v!);
        v!.ReceivedAt = DateTime.Now;
        v.ReceivedByID = user.EmployeeID;
        await db.SaveChangesAsync();
        await AuditEditedAsync(db, user, "CASH_TRANSFER_RECEIVE", v, before);
        return new TreasuryResult(SaveStatus.Saved, "", v.CashVoucherID, v.VoucherNumber);
    }

    /// <summary>modCash.CancelCashTransfer: before its receipt the sender takes a transfer back (the voucher is deleted); a
    /// closing transfer stays.</summary>
    public async Task<TreasuryResult> CancelTransferAsync(ActingUser user, int voucherId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var v = await PendingTransferAsync(db, voucherId);
        string msg;
        if (v is null)
        {
            msg = "هذا التحويل ليس بانتظار الاستلام.";
        }
        else if (!await permissions.HasPermissionAsync(user.EmployeeID, "CASH_BOX"))
        {
            msg = "ليست لديك صلاحية سندات الخزينة.";
        }
        else if (!await CanUseBoxAsync(db, user.EmployeeID, v.CashBoxID))
        {
            msg = await BoxUseProblemAsync(db, user.EmployeeID, v.CashBoxID);
        }
        else if (v.ClosingID is not null)
        {
            msg = "تحويل تصفية الكاشير لا يُلغى: يستلمه أمين الصندوق المستلم.";
        }
        else
        {
            msg = StockService.ClosedPeriodProblem(v.VoucherDate, (await SettingsReader.ReadAsync(db)).ClosedThrough);
        }
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        await AuditDeleteAsync(db, user, "CASH_TRANSFER_DELETE", "CashVouchers", v!.CashVoucherID, v, () => db.CashVouchers.Remove(v));
        return new TreasuryResult(SaveStatus.Saved, "", v.CashVoucherID, v.VoucherNumber);
    }

    /// <summary>modAudit.AuditEdited: the fields whose value changed.</summary>
    private static async Task AuditEditedAsync(RetailDbContext db, ActingUser user, string action, CashVouchers v,
                                               List<(string Field, string Caption, object? Value)> before)
    {
        var after = Audit.Snapshot("CashVouchers", v);
        var changes = before.Zip(after).Where(x => !Audit.SameValue(x.First.Value, x.Second.Value))
                            .Select(x => new AuditChange(x.First.Field, x.First.Caption, Audit.ValueText(x.First.Value), Audit.ValueText(x.Second.Value)))
                            .ToList();
        await Audit.WriteAsync(db, user, action, "CashVouchers", v.CashVoucherID.ToString(), Audit.Label(after.Select(s => (s.Field, s.Value))), changes);
    }

    /// <summary>AuditSnapshot before the DELETE, AuditDeleted after it (every field that had a value).</summary>
    private static async Task AuditDeleteAsync(RetailDbContext db, ActingUser user, string action, string table, int id, object entity, Action remove)
    {
        var snapshot = Audit.Snapshot(table, entity);
        remove();
        await db.SaveChangesAsync();
        var changes = snapshot.Where(s => !Audit.IsEmpty(s.Value)).Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList();
        await Audit.WriteAsync(db, user, action, table, id.ToString(), Audit.Label(snapshot.Select(s => (s.Field, s.Value))), changes);
    }

    /// <summary>The active boxes the user can choose (all of them; the checks say which he may pay from), with their balance.</summary>
    public async Task<List<(int Id, string Name, decimal Balance)>> BoxesAsync(bool english)
    {
        var all = await BoxBalancesAsync(english);
        return all.Where(b => b.IsActive).Select(b => (b.CashBoxID, b.BoxName, b.Balance)).ToList();
    }

    public async Task<List<(int Id, string Name)>> ExpenseTypesAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.ExpenseTypes.AsNoTracking().Where(t => t.IsActive).OrderBy(t => t.ExpenseTypeName).ToListAsync();
        return rows.Select(t => (t.ExpenseTypeID, Name(t.ExpenseTypeName, t.ExpenseTypeNameEn, english))).ToList();
    }

    /// <summary>The box of the user (modCash.CurrentCashBoxID outside a rep day).</summary>
    public async Task<int> UserBoxAsync(int employeeId)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await Sales.CurrentCashBoxAsync(db, employeeId);
    }

    // ------------------------------------------------------------------ banks

    /// <summary>modBank.BankTxTypeName.</summary>
    public static string BankTxTypeName(string txType) => txType switch
    {
        "DEPOSIT" => "إيداع نقدية في البنك",
        "WITHDRAW" => "سحب نقدية من البنك",
        "SETTLEMENT" => "تسوية تحصيلات مدى",
        "TRANSFER" => "تحويل بين البنوك",
        "OTHER_IN" => "وارد آخر للبنك",
        "OTHER_OUT" => "صادر آخر من البنك",
        _ => "",
    };

    /// <summary>BankBalanceQuery: the balance of each bank in the books = its journal lines (account 120000 + bank,
    /// qryBankItemSums).</summary>
    public async Task<List<BankBalance>> BankBalancesAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var sums = await (from l in db.JournalLines
                          join e in db.JournalEntries on l.EntryID equals e.EntryID
                          where l.AccountCode > 120000 && l.AccountCode < 130000
                          group new { l, e } by l.AccountCode into g
                          select new { g.Key, Balance = g.Sum(x => x.l.Debit) - g.Sum(x => x.l.Credit), Last = g.Max(x => x.e.EntryDate) })
            .ToListAsync();
        var byBank = sums.ToDictionary(s => s.Key - 120000);
        var banks = await db.Banks.AsNoTracking().ToListAsync();
        return banks.Select(k => new BankBalance(k.BankID, Name(k.BankName, k.BankNameEn, english), k.AccountNo, k.IBAN, k.IsActive,
                                                 byBank.TryGetValue(k.BankID, out var s) ? s.Balance : 0,
                                                 byBank.TryGetValue(k.BankID, out var t) ? t.Last : null))
            .OrderBy(b => b.BankName).ToList();
    }

    /// <summary>modBank.CounterAccountProblem: an active sub-account, not a box, a bank, Mada, customers or suppliers.</summary>
    public static async Task<string> CounterAccountProblemAsync(RetailDbContext db, int? accountCode)
    {
        if (accountCode is not { } code)
        {
            return "اختر الحساب المقابل.";
        }
        int[] excluded = [1190, 1200, 1300, 2100];
        bool ok = await db.Accounts.AnyAsync(a => a.AccountCode == code && a.IsPosting && a.IsActive && (a.Level3Code ?? 0) != 1100 &&
                                                  (a.Level3Code ?? 0) != 1210 && !excluded.Contains(a.AccountCode));
        return ok ? "" : "الحساب المقابل حساب فرعي نشط، وليس صندوقًا ولا بنكًا ولا مدى ولا العملاء أو الموردين.\n" +
                         "للإيداع والسحب والتحويل وتسوية مدى اختر نوع الحركة الخاص بها.";
    }

    private static async Task<bool> ActiveBankAsync(RetailDbContext db, int? bankId) =>
        bankId is { } id && await db.Banks.AnyAsync(b => b.IsActive && b.BankID == id);

    /// <summary>modBank.BankTxProblem.</summary>
    public async Task<string> BankTxProblemAsync(RetailDbContext db, int employeeId, BankTxRequest t, DateTime today, DateTime? closedThrough)
    {
        string p = "";
        if (t.TxDate is not { } d)
        {
            p = "اكتب تاريخ الحركة.";
        }
        else if (d.Date > today)
        {
            p = "تاريخ الحركة بعد اليوم.";
        }
        else if (!await ActiveBankAsync(db, t.BankID))
        {
            p = "اختر البنك.";
        }
        else if (t.Amount <= 0)
        {
            p = "المبلغ يجب أن يكون أكبر من صفر.";
        }
        else if (t.FeeAmount < 0 || t.FeeVAT < 0)
        {
            p = "العمولة وضريبتها لا تكونان سالبتين.";
        }
        else
        {
            switch (t.TxType)
            {
                case "DEPOSIT":
                case "WITHDRAW":
                    if (t.CashBoxID is null)
                    {
                        p = "اختر الصندوق.";
                    }
                    else if (!await BoxIsActiveAsync(db, t.CashBoxID))
                    {
                        p = "الصندوق غير نشط.";
                    }
                    else if (t.TxType == "DEPOSIT" && !await CanUseBoxAsync(db, employeeId, t.CashBoxID.Value))
                    {
                        p = await BoxUseProblemAsync(db, employeeId, t.CashBoxID.Value);
                    }
                    else if (t.TxType == "DEPOSIT" && await CashBoxBalanceAsync(db, t.CashBoxID.Value, d.Date.AddDays(1)) < t.Amount)
                    {
                        p = "رصيد الصندوق في " + StockService.GDate(d) + " أقل من المبلغ المودَع.";
                    }
                    break;
                case "SETTLEMENT":
                    if (t.FeeAmount + t.FeeVAT >= t.Amount)
                    {
                        p = "العمولة وضريبتها أقل من مبلغ التسوية.";
                    }
                    break;
                case "TRANSFER":
                    if (!await ActiveBankAsync(db, t.ToBankID))
                    {
                        p = "اختر البنك المحوَّل إليه.";
                    }
                    else if (t.ToBankID == t.BankID)
                    {
                        p = "التحويل بين بنكين مختلفين.";
                    }
                    break;
                case "OTHER_IN":
                case "OTHER_OUT":
                    p = await CounterAccountProblemAsync(db, t.CounterAccount);
                    if (p.Length == 0 && t.TxType == "OTHER_OUT" && t.FeeVAT >= t.Amount && t.FeeVAT > 0)
                    {
                        p = "ضريبة المدخلات أقل من المبلغ.";
                    }
                    break;
                default:
                    p = "اختر نوع الحركة.";
                    break;
            }
        }
        if (p.Length == 0)
        {
            p = StockService.ClosedPeriodProblem(t.TxDate!.Value, closedThrough);
        }
        return p;
    }

    /// <summary>The journal after a bank transaction, as modBank calls SyncJournal (its message, "" when every entry is fine).</summary>
    private async Task<string> SyncJournalAsync(ActingUser user) => journal is null ? "" : (await journal.SyncAsync(user)).Message;

    /// <summary>modBank.PostBankTx: BANKS, the add permission of frmBankTx, BankTxProblem, then the transaction, its audit row and
    /// the journal.</summary>
    public async Task<TreasuryResult> PostBankTxAsync(ActingUser user, BankTxRequest t)
    {
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "BANKS"))
        {
            return TreasuryResult.Refused("لا تملك صلاحية البنوك.");
        }
        if ((await permissions.ActionProblemAsync(user.EmployeeID, BankTxForm, ScreenAction.Add)).Length > 0)
        {
            return TreasuryResult.Refused("لا تملك صلاحية الإضافة في شاشة الحركات البنكية.");
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        string msg = await BankTxProblemAsync(db, user.EmployeeID, t, DateTime.Today, settings.ClosedThrough);
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        string number = await NumberSequences.NextNumberAsync(db, "BANK_TX");
        string reference = (t.Reference ?? "").Trim();
        string description = (t.Description ?? "").Trim();
        var row = new BankTransactions
        {
            TxNumber = number, TxDate = t.TxDate!.Value.Date, TxType = t.TxType, BankID = t.BankID,
            ToBankID = t.TxType == "TRANSFER" ? t.ToBankID : null,
            CashBoxID = t.TxType is "DEPOSIT" or "WITHDRAW" ? t.CashBoxID : null,
            CounterAccount = t.TxType is "OTHER_IN" or "OTHER_OUT" ? t.CounterAccount : null,
            Amount = t.Amount, FeeAmount = t.TxType == "SETTLEMENT" ? t.FeeAmount : 0,
            FeeVAT = t.TxType is "SETTLEMENT" or "OTHER_OUT" ? t.FeeVAT : 0,
            Reference = reference.Length > 0 ? SalesService.Clip(reference, 40) : null,
            Description = SalesService.Clip(description.Length > 0 ? description : BankTxTypeName(t.TxType), 255), EmployeeID = user.EmployeeID,
        };
        db.BankTransactions.Add(row);
        await db.SaveChangesAsync();
        Audit.LogAction(db, user, "BANK_TX", "BankTransactions", number, t.TxType + " " +
                        t.Amount.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
        await db.SaveChangesAsync();
        return new TreasuryResult(SaveStatus.Saved, await SyncJournalAsync(user), row.BankTxID, number);
    }

    /// <summary>modBank.DeleteBankTx: not a transaction cleared in a reconciliation, not in a closed period.</summary>
    public async Task<TreasuryResult> DeleteBankTxAsync(ActingUser user, int bankTxId)
    {
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "BANKS") ||
            (await permissions.ActionProblemAsync(user.EmployeeID, BankTxForm, ScreenAction.Delete)).Length > 0)
        {
            return TreasuryResult.Refused("لا تملك صلاحية حذف الحركات البنكية.");
        }
        await using var db = await factory.CreateDbContextAsync();
        var t = await db.BankTransactions.FirstOrDefaultAsync(x => x.BankTxID == bankTxId);
        if (t is null)
        {
            return TreasuryResult.Refused("الحركة غير موجودة.");
        }
        if (await db.BankClearings.AnyAsync(c => c.SourceType == "BANK_TX" && c.SourceID == bankTxId))
        {
            return TreasuryResult.Refused("الحركة مطابقة في تسوية بنكية: ألغِ مطابقتها أولًا.");
        }
        string msg = StockService.ClosedPeriodProblem(t.TxDate, (await SettingsReader.ReadAsync(db)).ClosedThrough);
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        await AuditDeleteAsync(db, user, "BANK_TX_DELETE", "BankTransactions", bankTxId, t, () => db.BankTransactions.Remove(t));
        return new TreasuryResult(SaveStatus.Saved, await SyncJournalAsync(user), bankTxId, t.TxNumber);
    }

    /// <summary>The bank transactions of a period, newest first (bank 0 = every bank).</summary>
    public async Task<List<BankTxRow>> BankTxAsync(DateTime from, DateTime to, int bankId, bool english, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var q = db.BankTransactions.AsNoTracking().Where(t => t.TxDate >= start && t.TxDate < end);
        if (bankId > 0)
        {
            q = q.Where(t => t.BankID == bankId || t.ToBankID == bankId);
        }
        var rows = await q.OrderByDescending(t => t.TxDate).ThenByDescending(t => t.BankTxID).Take(take).ToListAsync();
        var banks = await db.Banks.AsNoTracking().ToDictionaryAsync(b => b.BankID, b => Name(b.BankName, b.BankNameEn, english));
        var boxes = await db.CashBoxes.AsNoTracking().ToDictionaryAsync(b => b.CashBoxID, b => Name(b.BoxName, b.BoxNameEn, english));
        var accounts = await db.Accounts.AsNoTracking().ToDictionaryAsync(a => a.AccountCode, a => Name(a.AccountName, a.AccountNameEn, english));
        var ids = rows.Select(r => r.EmployeeID).Distinct().ToList();
        var users = await db.Employees.AsNoTracking().Where(e => ids.Contains(e.EmployeeID)).ToDictionaryAsync(e => e.EmployeeID, e => e.EmployeeName);
        return rows.Select(t =>
        {
            string? other = t.ToBankID is { } tb && banks.TryGetValue(tb, out var bn) ? bn
                : t.CashBoxID is { } cb && boxes.TryGetValue(cb, out var xn) ? xn
                : t.CounterAccount is { } ca && accounts.TryGetValue(ca, out var an) ? ca + " " + an : null;
            return new BankTxRow(t.BankTxID, t.TxNumber, t.TxDate, t.TxType, Tr(BankTxTypeName(t.TxType), english),
                                 banks.TryGetValue(t.BankID, out var b) ? b : "", other, t.Amount, t.FeeAmount, t.FeeVAT, t.Reference, t.Description,
                                 users.TryGetValue(t.EmployeeID, out var u) ? u : "");
        }).ToList();
    }

    /// <summary>The active banks for the choice lists.</summary>
    public async Task<List<(int Id, string Name)>> BanksAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Banks.AsNoTracking().Where(b => b.IsActive).OrderBy(b => b.BankName).ToListAsync();
        return rows.Select(b => (b.BankID, Name(b.BankName, b.BankNameEn, english))).ToList();
    }

    /// <summary>The counter accounts of other bank transactions (CounterAccountProblem accepts them).</summary>
    public async Task<List<(int Code, string Name)>> CounterAccountsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        int[] excluded = [1190, 1200, 1300, 2100];
        var rows = await db.Accounts.AsNoTracking()
            .Where(a => a.IsPosting && a.IsActive && (a.Level3Code ?? 0) != 1100 && (a.Level3Code ?? 0) != 1210 && !excluded.Contains(a.AccountCode))
            .OrderBy(a => a.AccountCode).ToListAsync();
        return rows.Select(a => (a.AccountCode, Name(a.AccountName, a.AccountNameEn, english))).ToList();
    }
}
