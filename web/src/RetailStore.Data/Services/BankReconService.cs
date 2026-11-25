using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Domain.Sales;

namespace RetailStore.Data.Services;

/// <summary>An operation on the account of a bank (qryBankItems): its net amount in the journal (+ in, - out), and whether it was
/// ticked in a statement (ReconciliationID, ClearedAmount = its amount on that day).</summary>
public sealed record BankItem(int BankID, string SourceType, int SourceID, DateTime ItemDate, string? ItemNumber, string? ItemText,
                              decimal ItemAmount, string TypeName, int? ReconciliationID, decimal? ClearedAmount)
{
    public bool IsCleared => ReconciliationID is not null;

    /// <summary>The key of the screen lists (SourceType|SourceID).</summary>
    public string Key => SourceType + "|" + SourceID;
}

/// <summary>A reconciliation in the list of the screen (lstRecons).</summary>
public sealed record ReconRow(int ReconciliationID, string ReconNumber, DateTime StatementDate, decimal StatementBalance, string Status);

/// <summary>The figures of a reconciliation (ReconFigures): the book balance on the statement day, the operations until that day
/// not in a statement yet, the adjusted balance and the difference with the statement.</summary>
public sealed record ReconFiguresView(decimal Book, decimal Outstanding, decimal Adjusted, decimal Difference);

/// <summary>A reconciliation as frmBankRecon shows it: its figures, the open operations until the statement day, the ones ticked
/// in it, and the operations whose amount changed after they were ticked (ChangedClearings).</summary>
public sealed record ReconView(ReconRow Recon, ReconFiguresView Figures, List<BankItem> Open, List<BankItem> Cleared, int Changed);

/// <summary>The bank reconciliation on the web and the API (plan 7C, docs/87): modBank.OpenReconciliation, ClearBankItem,
/// UnclearBankItem, FinishReconciliation, ReopenReconciliation, DeleteReconciliation, ReconFigures and ChangedClearings, with the
/// same checks, order and messages, over the journal lines of the bank (qryBankItemSums / qryBankItems).</summary>
public sealed class BankReconService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions, JournalService? journal = null)
{
    public const string Form = "frmBankRecon";
    public const string NotFoundOrDone = "التسوية غير موجودة أو معتمدة.";

    private static string Name(string arabic, string? english, bool inEnglish) => SalesService.Name(arabic, english, inEnglish);

    /// <summary>modJournal.SyncJournal on opening the screen (BankReconLoad): "" = done (or no journal service).</summary>
    public async Task<string> SyncAsync(ActingUser user) => journal is null ? "" : (await journal.SyncAsync(user)).Message;

    /// <summary>qryBankItems for one bank: the net of each operation on account 120000 + bank, with its clearing.</summary>
    public static async Task<List<BankItem>> ItemsAsync(RetailDbContext db, int bankId, bool english)
    {
        int account = 120000 + bankId;
        var lines = await (from l in db.JournalLines
                           join e in db.JournalEntries on l.EntryID equals e.EntryID
                           where l.AccountCode == account
                           select new { e.SourceType, e.SourceID, e.EntryDate, e.SourceNumber, e.Description, l.Debit, l.Credit })
            .AsNoTracking().ToListAsync();
        var types = await db.JournalSourceTypes.AsNoTracking().ToDictionaryAsync(t => t.SourceType, t => Name(t.TypeName, t.TypeNameEn, english));
        var clearings = await db.BankClearings.AsNoTracking().Where(c => c.BankID == bankId).ToListAsync();
        var byKey = clearings.GroupBy(c => (c.SourceType, c.SourceID)).ToDictionary(g => g.Key, g => g.First());
        return lines.GroupBy(l => (l.SourceType, l.SourceID))
            .Select(g =>
            {
                decimal amount = g.Sum(x => x.Debit) - g.Sum(x => x.Credit);
                var c = byKey.GetValueOrDefault(g.Key);
                return new { g.Key, amount, c, Date = g.Max(x => x.EntryDate), Number = g.Max(x => x.SourceNumber), Text = g.Max(x => x.Description) };
            })
            .Where(x => x.amount != 0 && types.ContainsKey(x.Key.SourceType))
            .Select(x => new BankItem(bankId, x.Key.SourceType, x.Key.SourceID, x.Date, x.Number, x.Text, x.amount, types[x.Key.SourceType],
                                      x.c?.ReconciliationID, x.c?.ClearedAmount))
            .OrderBy(i => i.ItemDate).ThenBy(i => i.ItemNumber, StringComparer.Ordinal).ToList();
    }

    /// <summary>modBank.ReconFigures (null = no such reconciliation).</summary>
    public static async Task<ReconFiguresView?> FiguresAsync(RetailDbContext db, int reconId)
    {
        var r = await db.BankReconciliations.AsNoTracking().FirstOrDefaultAsync(x => x.ReconciliationID == reconId);
        if (r is null)
        {
            return null;
        }
        var cutoff = r.StatementDate.Date.AddDays(1);
        var items = (await ItemsAsync(db, r.BankID, false)).Where(i => i.ItemDate < cutoff).ToList();
        decimal book = items.Sum(i => i.ItemAmount);
        decimal outstanding = items.Where(i => !i.IsCleared).Sum(i => i.ItemAmount);
        return new ReconFiguresView(book, outstanding, book - outstanding, r.StatementBalance - (book - outstanding));
    }

    private static async Task<bool> ActiveBankAsync(RetailDbContext db, int bankId) =>
        await db.Banks.AnyAsync(b => b.IsActive && b.BankID == bankId);

    private async Task<bool> CanAsync(int employeeId, ScreenAction action) =>
        (await permissions.ActionProblemAsync(employeeId, Form, action)).Length == 0;

    /// <summary>modBank.OpenReconciliation: the open reconciliation of the bank (a new one, or the open one with this date and
    /// balance). Id = the reconciliation.</summary>
    public async Task<TreasuryResult> OpenAsync(ActingUser user, int bankId, DateTime? statementDate, decimal statementBalance, DateTime? today = null)
    {
        if (!await permissions.HasPermissionAsync(user.EmployeeID, "BANKS"))
        {
            return TreasuryResult.Refused("لا تملك صلاحية البنوك.");
        }
        await using var db = await factory.CreateDbContextAsync();
        if (!await ActiveBankAsync(db, bankId))
        {
            return TreasuryResult.Refused("اختر البنك.");
        }
        if (statementDate is not { } d)
        {
            return TreasuryResult.Refused("اكتب تاريخ كشف البنك.");
        }
        if (d.Date > (today ?? DateTime.Today))
        {
            return TreasuryResult.Refused("تاريخ الكشف بعد اليوم.");
        }
        var lastDone = await db.BankReconciliations.Where(r => r.Status == "DONE" && r.BankID == bankId)
            .Select(r => (DateTime?)r.StatementDate).MaxAsync();
        if (lastDone is { } last && d.Date < last.Date)
        {
            return TreasuryResult.Refused("آخر تسوية معتمدة لهذا البنك في " + StockService.GDate(last) + ": اختر تاريخًا بعده.");
        }
        var open = await db.BankReconciliations.FirstOrDefaultAsync(r => r.Status == "OPEN" && r.BankID == bankId);
        if (open is null)
        {
            if (!await CanAsync(user.EmployeeID, ScreenAction.Add))
            {
                return TreasuryResult.Refused("لا تملك صلاحية الإضافة في شاشة التسوية البنكية.");
            }
            open = new BankReconciliations
            {
                ReconNumber = await NumberSequences.NextNumberAsync(db, "BANK_RECON"), BankID = bankId, StatementDate = d.Date,
                StatementBalance = statementBalance, Status = "OPEN", EmployeeID = user.EmployeeID,
            };
            db.BankReconciliations.Add(open);
        }
        else
        {
            open.StatementDate = d.Date;
            open.StatementBalance = statementBalance;
        }
        await db.SaveChangesAsync();
        return new TreasuryResult(SaveStatus.Saved, "", open.ReconciliationID, open.ReconNumber);
    }

    private static async Task<BankReconciliations?> OpenReconAsync(RetailDbContext db, int reconId) =>
        await db.BankReconciliations.FirstOrDefaultAsync(r => r.ReconciliationID == reconId && r.Status == "OPEN");

    /// <summary>modBank.ClearBankItem: the operation appears in the bank statement of this reconciliation.</summary>
    public async Task<string> ClearAsync(int reconId, string sourceType, int sourceId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await OpenReconAsync(db, reconId);
        if (r is null)
        {
            return NotFoundOrDone;
        }
        var item = (await ItemsAsync(db, r.BankID, false)).FirstOrDefault(i => i.SourceType == sourceType && i.SourceID == sourceId && !i.IsCleared);
        if (item is null)
        {
            return "العملية غير موجودة أو مطابقة من قبل.";
        }
        db.BankClearings.Add(new BankClearings
        {
            ReconciliationID = reconId, BankID = r.BankID, SourceType = sourceType, SourceID = sourceId, ClearedAmount = item.ItemAmount,
        });
        await db.SaveChangesAsync();
        return "";
    }

    /// <summary>modBank.UnclearBankItem.</summary>
    public async Task<string> UnclearAsync(int reconId, string sourceType, int sourceId)
    {
        await using var db = await factory.CreateDbContextAsync();
        if (await OpenReconAsync(db, reconId) is null)
        {
            return NotFoundOrDone;
        }
        var rows = await db.BankClearings.Where(c => c.ReconciliationID == reconId && c.SourceType == sourceType && c.SourceID == sourceId).ToListAsync();
        db.BankClearings.RemoveRange(rows);
        await db.SaveChangesAsync();
        return "";
    }

    /// <summary>modBank.FinishReconciliation: approved only without a difference.</summary>
    public async Task<string> FinishAsync(ActingUser user, int reconId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await OpenReconAsync(db, reconId);
        if (r is null)
        {
            return NotFoundOrDone;
        }
        if (!await CanAsync(user.EmployeeID, ScreenAction.Edit))
        {
            return "لا تملك صلاحية اعتماد التسوية البنكية.";
        }
        var f = (await FiguresAsync(db, reconId))!;
        if (f.Difference != 0)
        {
            return "يوجد فرق " + InvoiceCalculator.Money(f.Difference) + " بين كشف البنك والدفاتر.\n" +
                   "طابق العمليات الظاهرة في الكشف، وسجّل ما في الكشف ولم يُسجَّل (عمولات، فوائد...) كحركة بنكية.";
        }
        r.Status = "DONE";
        r.BookBalance = f.Book;
        r.Outstanding = f.Outstanding;
        r.Difference = 0;
        r.EmployeeID = user.EmployeeID;
        Audit.LogAction(db, user, "BANK_RECON", "BankReconciliations", reconId.ToString());
        await db.SaveChangesAsync();
        return "";
    }

    /// <summary>modBank.ReopenReconciliation: only the last reconciliation of the bank, and none open.</summary>
    public async Task<string> ReopenAsync(ActingUser user, int reconId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await db.BankReconciliations.FirstOrDefaultAsync(x => x.ReconciliationID == reconId);
        if (r is null)
        {
            return "التسوية غير موجودة.";
        }
        if (!await CanAsync(user.EmployeeID, ScreenAction.Edit))
        {
            return "لا تملك صلاحية تعديل التسوية البنكية.";
        }
        if (await db.BankReconciliations.AnyAsync(x => x.Status == "OPEN" && x.BankID == r.BankID))
        {
            return "توجد تسوية جارية لهذا البنك: اعتمدها أو احذفها أولًا.";
        }
        int lastId = (await db.BankReconciliations.AsNoTracking().Where(x => x.BankID == r.BankID).ToListAsync())
            .OrderByDescending(x => x.StatementDate).ThenByDescending(x => x.ReconciliationID).First().ReconciliationID;
        if (lastId != reconId)
        {
            return "تُعاد فتح آخر تسوية للبنك فقط.";
        }
        r.Status = "OPEN";
        Audit.LogAction(db, user, "BANK_RECON_REOPEN", "BankReconciliations", reconId.ToString());
        await db.SaveChangesAsync();
        return "";
    }

    /// <summary>modBank.DeleteReconciliation: the open one only, with its clearings; audited like Access.</summary>
    public async Task<string> DeleteAsync(ActingUser user, int reconId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await OpenReconAsync(db, reconId);
        if (r is null)
        {
            return "تُحذف التسوية الجارية فقط (أعد فتح المعتمدة أولًا).";
        }
        var snapshot = Audit.Snapshot("BankReconciliations", r);
        db.BankClearings.RemoveRange(await db.BankClearings.Where(c => c.ReconciliationID == reconId).ToListAsync());
        db.BankReconciliations.Remove(r);
        await db.SaveChangesAsync();
        var changes = snapshot.Where(s => !Audit.IsEmpty(s.Value)).Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList();
        await Audit.WriteAsync(db, user, "BANK_RECON_DELETE", "BankReconciliations", reconId.ToString(),
                               Audit.Label(snapshot.Select(s => (s.Field, s.Value))), changes);
        return "";
    }

    /// <summary>The reconciliations of a bank, the latest first (lstRecons), and its open one (BankReconBankChanged).</summary>
    public async Task<List<ReconRow>> ReconsAsync(int bankId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var rows = await db.BankReconciliations.AsNoTracking().Where(r => r.BankID == bankId).ToListAsync();
        return rows.OrderByDescending(r => r.StatementDate).ThenByDescending(r => r.ReconciliationID)
            .Select(r => new ReconRow(r.ReconciliationID, r.ReconNumber, r.StatementDate, r.StatementBalance, r.Status)).ToList();
    }

    /// <summary>BankReconRefresh: the reconciliation with its lists and figures (null = none).</summary>
    public async Task<ReconView?> ViewAsync(int reconId, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        var r = await db.BankReconciliations.AsNoTracking().FirstOrDefaultAsync(x => x.ReconciliationID == reconId);
        if (r is null)
        {
            return null;
        }
        var items = await ItemsAsync(db, r.BankID, english);
        var cutoff = r.StatementDate.Date.AddDays(1);
        var figures = (await FiguresAsync(db, reconId))!;
        int changed = items.Count(i => i.IsCleared && i.ClearedAmount != i.ItemAmount);
        return new ReconView(new ReconRow(r.ReconciliationID, r.ReconNumber, r.StatementDate, r.StatementBalance, r.Status), figures,
                             items.Where(i => !i.IsCleared && i.ItemDate < cutoff).ToList(),
                             items.Where(i => i.ReconciliationID == reconId).ToList(), changed);
    }

    /// <summary>The state line of the screen (BankReconRefresh lblState).</summary>
    public static string StateText(ReconView v)
    {
        string state = "التسوية " + v.Recon.ReconNumber + (v.Recon.Status == "OPEN" ? " جارية" : " معتمدة") +
                       ": اختر العمليات الظاهرة في كشف البنك ثم «مطابقة».";
        if (v.Changed > 0)
        {
            state += "\n" + v.Changed + " عملية مطابقة تغيّر مبلغها في الدفاتر بعد المطابقة: راجعها.";
        }
        return state;
    }
}
