using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Domain.Common;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

/// <summary>A line of a manual entry as typed (tmpManualLines): the amounts in the currency of the entry; an empty line (no
/// account, no amount) is ignored.</summary>
public sealed record ManualLineInput(int? AccountCode, decimal Debit, decimal Credit, string? LineText = null, int? CostCenterID = null);

/// <summary>A manual entry to save (PostManualEntry): ManualEntryID 0 = a new one; ReversalOfID = the entry it reverses.</summary>
public sealed record ManualEntryRequest(int ManualEntryID, DateTime? EntryDate, string? Description, string? Reference, int? ReversalOfID,
                                        IReadOnlyList<ManualLineInput> Lines, string? CurrencyCode = null, decimal ExchangeRate = 1);

/// <summary>A saved manual entry with its lines in its currency (ManualOpen, LINE_AMOUNTS) and its journal entry.</summary>
public sealed record ManualEntryView(int ManualEntryID, string EntryNumber, DateTime EntryDate, string Description, string? Reference,
                                     int? ReversalOfID, string CurrencyCode, decimal ExchangeRate, decimal TotalAmount, bool InClosedPeriod,
                                     string? JournalNumber, List<ManualLineInput> Lines);

public sealed record ManualEntryRow(int ManualEntryID, string EntryNumber, DateTime EntryDate, string Description, decimal TotalAmount);

/// <summary>Manual journal entries on the web and the API (plan 7C, docs/87): modManualEntry.ManualEntryProblem, PostManualEntry,
/// RemoveManualEntry and ReverseManualEntry with the same checks, order and messages; the journal follows at once
/// (SyncJournal, source MANUAL).</summary>
public sealed class ManualEntryService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, JournalService? journal = null)
{
    public const string Form = "frmManualEntry";
    public const string NotFound = "القيد غير موجود.";

    /// <summary>modManualEntry.ManualEntryProblem: the header and the lines ("" = the entry can be saved).</summary>
    public static async Task<string> ProblemAsync(RetailDbContext db, DateTime? entryDate, string? description, IReadOnlyList<ManualLineInput> lines)
    {
        if (entryDate is null)
        {
            return "اكتب تاريخ القيد.";
        }
        if ((description ?? "").Trim().Length == 0)
        {
            return "اكتب بيان القيد.";
        }
        int n = 0;
        decimal debit = 0, credit = 0;
        foreach (var l in lines)
        {
            if (l.AccountCode is null && l.Debit == 0 && l.Credit == 0)
            {
                continue;                                   // an empty line is ignored
            }
            if (l.AccountCode is null)
            {
                return "اختر الحساب في كل سطر فيه مبلغ.";
            }
            if (l.Debit < 0 || l.Credit < 0)
            {
                return "المبالغ لا تكون سالبة: اكتب المبلغ في الجانب الآخر.";
            }
            if (l.Debit > 0 && l.Credit > 0)
            {
                return "السطر يكون مدينًا أو دائنًا، وليس الاثنين.";
            }
            if (l.Debit == 0 && l.Credit == 0)
            {
                return "اكتب المبلغ مدينًا أو دائنًا في كل سطر فيه حساب.";
            }
            int acc = l.AccountCode.Value;
            if (!await db.Accounts.AnyAsync(a => a.AccountCode == acc && a.IsPosting && a.IsActive))
            {
                return "الحساب " + acc + " حساب رئيسي أو معطّل: القيود على الحسابات الفرعية النشطة فقط.";
            }
            n++;
            debit += l.Debit;
            credit += l.Credit;
        }
        if (n < 2)
        {
            return "القيد سطران على الأقل: طرف مدين وطرف دائن.";
        }
        if (debit != credit)
        {
            return "القيد غير متوازن: المدين " + InvoiceCalculator.Money(debit) + " والدائن " + InvoiceCalculator.Money(credit) +
                   " (الفرق " + InvoiceCalculator.Money(Math.Abs(debit - credit)) + ").";
        }
        return "";
    }

    /// <summary>modManualEntry.PostManualEntry (and SaveManualEntry: the screen and its ADD / EDIT action), then SyncJournal.
    /// Id = the saved entry; Message = the journal message when the journal could not follow.</summary>
    public async Task<TreasuryResult> SaveAsync(ActingUser user, ManualEntryRequest e)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, Form);
        if (open.Length == 0)
        {
            open = await permissions.ActionProblemAsync(user.EmployeeID, Form, e.ManualEntryID == 0 ? ScreenAction.Add : ScreenAction.Edit);
        }
        if (open.Length > 0)
        {
            return TreasuryResult.Refused(open);
        }
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "MANUAL_ENTRY"))
        {
            return TreasuryResult.Refused("لا تملك صلاحية القيود اليدوية.");
        }
        await using var db = await factory.CreateDbContextAsync();
        var settings = await SettingsReader.ReadAsync(db);
        string baseCurrency = string.IsNullOrWhiteSpace(settings.CurrencyCode) ? "SAR" : settings.CurrencyCode;
        string code = string.IsNullOrWhiteSpace(e.CurrencyCode) ? baseCurrency : e.CurrencyCode.Trim();
        decimal fx = string.Equals(code, baseCurrency, StringComparison.OrdinalIgnoreCase) ? 1 : e.ExchangeRate;
        var lines = e.Lines ?? [];
        string msg = await ProblemAsync(db, e.EntryDate, e.Description, lines);
        if (msg.Length == 0)
        {
            msg = await SalesService.CurrencyProblemAsync(db, code, fx, baseCurrency);
        }
        if (msg.Length == 0)
        {
            msg = StockService.ClosedPeriodProblem(e.EntryDate!.Value, settings.ClosedThrough);
        }
        ManualEntries? h = null;
        if (msg.Length == 0 && e.ManualEntryID > 0)
        {
            h = await db.ManualEntries.FirstOrDefaultAsync(x => x.ManualEntryID == e.ManualEntryID);
            if (h is not null)
            {
                msg = StockService.ClosedPeriodProblem(h.EntryDate, settings.ClosedThrough);
            }
        }
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        if (e.ManualEntryID > 0 && h is null)
        {
            return TreasuryResult.Refused(NotFound);
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        List<(string Field, string Caption, object? Value)>? before = null;
        if (h is null)
        {
            h = new ManualEntries
            {
                EntryNumber = await NumberSequences.NextNumberAsync(db, "MANUAL_ENTRY"), EmployeeID = user.EmployeeID, ReversalOfID = e.ReversalOfID,
            };
            db.ManualEntries.Add(h);
        }
        else
        {
            before = Audit.Snapshot("ManualEntries", h);
            h.UpdatedAt = DateTime.Now;
            db.ManualEntryLines.RemoveRange(await db.ManualEntryLines.Where(l => l.ManualEntryID == h.ManualEntryID).ToListAsync());
        }
        h.EntryDate = e.EntryDate!.Value.Date;
        h.Description = SalesService.Clip(e.Description, 255)!;
        h.Reference = SalesService.Clip(e.Reference, 50);
        h.CurrencyCode = code;
        h.ExchangeRate = fx;
        await db.SaveChangesAsync();

        decimal total = 0, credits = 0, foreignTotal = 0;
        short n = 0;
        var added = new List<ManualEntryLines>();
        foreach (var l in lines.Where(l => l.AccountCode is not null))
        {
            var row = new ManualEntryLines
            {
                ManualEntryID = h.ManualEntryID, LineNumber = ++n, AccountCode = l.AccountCode!.Value, Debit = MoneyMath.RoundMoney(l.Debit * fx),
                Credit = MoneyMath.RoundMoney(l.Credit * fx), ForeignDebit = l.Debit, ForeignCredit = l.Credit,
                LineText = SalesService.Clip(l.LineText, 150), CostCenterID = l.CostCenterID,
            };
            total += row.Debit;
            credits += row.Credit;
            foreignTotal += l.Debit;
            added.Add(row);
        }
        // balanced in the currency, the conversion may leave a cent: on the largest line of the smaller side
        decimal diff = total - credits;
        if (diff > 0)
        {
            added.Where(r => r.Credit > 0).OrderByDescending(r => r.Credit).First().Credit += diff;
        }
        else if (diff < 0)
        {
            added.Where(r => r.Debit > 0).OrderByDescending(r => r.Debit).First().Debit -= diff;
            total -= diff;
        }
        db.ManualEntryLines.AddRange(added);
        h.TotalAmount = total;
        h.ForeignAmount = foreignTotal;
        await db.SaveChangesAsync();
        if (before is not null)
        {
            var after = Audit.Snapshot("ManualEntries", h);
            var changes = before.Zip(after).Where(p => !Audit.SameValue(p.First.Value, p.Second.Value))
                .Select(p => new AuditChange(p.First.Field, p.First.Caption, Audit.ValueText(p.First.Value), Audit.ValueText(p.Second.Value))).ToList();
            await Audit.WriteAsync(db, user, "MANUAL_ENTRY_EDIT", "ManualEntries", h.ManualEntryID.ToString(), h.EntryNumber, changes);
        }
        Audit.LogAction(db, user, e.ManualEntryID == 0 ? "MANUAL_ENTRY_ADD" : "MANUAL_ENTRY_EDIT", "ManualEntries", h.ManualEntryID.ToString());
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        string sync = journal is null ? "" : (await journal.SyncAsync(user)).Message;      // the journal follows at once
        return new TreasuryResult(SaveStatus.Saved, sync, h.ManualEntryID, h.EntryNumber);
    }

    /// <summary>modManualEntry.RemoveManualEntry (and DeleteManualEntry: the DELETE action), then SyncJournal.</summary>
    public async Task<TreasuryResult> DeleteAsync(ActingUser user, int manualEntryId)
    {
        string open = await permissions.OpenProblemAsync(user.EmployeeID, Form);
        if (open.Length == 0)
        {
            open = await permissions.ActionProblemAsync(user.EmployeeID, Form, ScreenAction.Delete);
        }
        if (open.Length > 0)
        {
            return TreasuryResult.Refused(open);
        }
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "MANUAL_ENTRY"))
        {
            return TreasuryResult.Refused("لا تملك صلاحية القيود اليدوية.");
        }
        await using var db = await factory.CreateDbContextAsync();
        var h = await db.ManualEntries.FirstOrDefaultAsync(x => x.ManualEntryID == manualEntryId);
        if (h is null)
        {
            return TreasuryResult.Refused(NotFound);
        }
        string msg = StockService.ClosedPeriodProblem(h.EntryDate, (await SettingsReader.ReadAsync(db)).ClosedThrough);
        if (msg.Length > 0)
        {
            return TreasuryResult.Refused(msg);
        }
        var snapshot = Audit.Snapshot("ManualEntries", h);
        db.ManualEntryLines.RemoveRange(await db.ManualEntryLines.Where(l => l.ManualEntryID == manualEntryId).ToListAsync());
        db.ManualEntries.Remove(h);
        await db.SaveChangesAsync();
        var changes = snapshot.Where(s => !Audit.IsEmpty(s.Value)).Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList();
        await Audit.WriteAsync(db, user, "MANUAL_ENTRY_DELETE", "ManualEntries", manualEntryId.ToString(),
                               Audit.Label(snapshot.Select(s => (s.Field, s.Value))), changes);
        string sync = journal is null ? "" : (await journal.SyncAsync(user)).Message;
        return new TreasuryResult(SaveStatus.Saved, sync, manualEntryId, h.EntryNumber);
    }

    /// <summary>ManualOpen: the entry with its lines in its own currency (null = none).</summary>
    public async Task<ManualEntryView?> EntryAsync(int manualEntryId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await db.ManualEntries.AsNoTracking().FirstOrDefaultAsync(x => x.ManualEntryID == manualEntryId);
        if (h is null)
        {
            return null;
        }
        var settings = await SettingsReader.ReadAsync(db);
        var rows = await db.ManualEntryLines.AsNoTracking().Where(l => l.ManualEntryID == manualEntryId).OrderBy(l => l.LineNumber).ToListAsync();
        var lines = rows.Select(l => l.ForeignDebit + l.ForeignCredit == 0
                ? new ManualLineInput(l.AccountCode, l.Debit, l.Credit, l.LineText, l.CostCenterID)
                : new ManualLineInput(l.AccountCode, l.ForeignDebit, l.ForeignCredit, l.LineText, l.CostCenterID)).ToList();
        string? jv = await db.JournalEntries.Where(j => j.SourceType == "MANUAL" && j.SourceID == manualEntryId).Select(j => j.EntryNumber)
            .FirstOrDefaultAsync();
        string baseCurrency = string.IsNullOrWhiteSpace(settings.CurrencyCode) ? "SAR" : settings.CurrencyCode;
        return new ManualEntryView(h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, h.Reference, h.ReversalOfID,
                                   string.IsNullOrEmpty(h.CurrencyCode) ? baseCurrency : h.CurrencyCode, h.ExchangeRate == 0 ? 1 : h.ExchangeRate,
                                   h.TotalAmount, StockService.ClosedPeriodProblem(h.EntryDate, settings.ClosedThrough).Length > 0, jv, lines);
    }

    /// <summary>modManualEntry.ReverseManualEntry: a new entry (not saved) with debit and credit swapped, dated today.</summary>
    public static ManualEntryRequest Reversal(ManualEntryView v, DateTime today)
    {
        string text = "عكس القيد " + v.EntryNumber + ": " + v.Description;
        return new ManualEntryRequest(0, today, text.Length <= 255 ? text : text[..255], v.Reference, v.ManualEntryID,
                                      v.Lines.Select(l => l with { Debit = l.Credit, Credit = l.Debit }).ToList(), v.CurrencyCode, v.ExchangeRate);
    }

    /// <summary>The entries of a period (cboFind), the latest first.</summary>
    public async Task<List<ManualEntryRow>> EntriesAsync(DateTime from, DateTime to, int take = 200)
    {
        await using var db = await factory.CreateDbContextAsync();
        var start = from.Date;
        var end = to.Date.AddDays(1);
        var rows = await db.ManualEntries.AsNoTracking().Where(h => h.EntryDate >= start && h.EntryDate < end).ToListAsync();
        return rows.OrderByDescending(h => h.EntryDate).ThenByDescending(h => h.ManualEntryID).Take(take)
            .Select(h => new ManualEntryRow(h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, h.TotalAmount)).ToList();
    }

    /// <summary>The accounts a line may take: active sub-accounts.</summary>
    public async Task<List<(int Code, string Name)>> AccountsAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.Accounts.AsNoTracking().Where(a => a.IsPosting && a.IsActive).OrderBy(a => a.AccountCode).ToListAsync();
        return rows.Select(a => (a.AccountCode, SalesService.Name(a.AccountName, a.AccountNameEn, english))).ToList();
    }

    /// <summary>The active cost centres a line may take.</summary>
    public async Task<List<(int Id, string Name)>> CostCentersAsync(bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.CostCenters.AsNoTracking().Where(c => c.IsActive).OrderBy(c => c.CenterName).ToListAsync();
        return rows.Select(c => (c.CostCenterID, SalesService.Name(c.CenterName, c.CenterNameEn, english))).ToList();
    }

    /// <summary>ManualRecalc: the totals line of the screen.</summary>
    public static string TotalsText(IEnumerable<ManualLineInput> lines)
    {
        decimal debit = lines.Sum(l => l.Debit), credit = lines.Sum(l => l.Credit);
        return "المدين: " + InvoiceCalculator.Money(debit) + "     الدائن: " + InvoiceCalculator.Money(credit) + "     " +
               (debit == credit && debit > 0 ? "متوازن" : "الفرق: " + InvoiceCalculator.Money(Math.Abs(debit - credit)));
    }
}
