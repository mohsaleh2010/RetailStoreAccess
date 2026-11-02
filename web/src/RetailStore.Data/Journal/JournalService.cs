using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.Domain;
using RetailStore.Domain.Journal;

namespace RetailStore.Data.Journal;

public sealed record JournalSyncResult(int Added, int Updated, int Removed, string Message)
{
    public bool Ok => Message.Length == 0;
}

/// <summary>modJournal.SyncJournal in C#: the entries are rebuilt from the operations, one source view per kind
/// (AppData.JournalSourceViews); the decisions are JournalPlanner's; everything runs in one transaction.</summary>
public sealed class JournalService(IDbContextFactory<RetailDbContext> factory, IJournalSourceReader reader)
{
    public async Task<JournalSyncResult> SyncAsync(ActingUser user)
    {
        await using var db = await factory.CreateDbContextAsync();
        await using var tx = await db.Database.BeginTransactionAsync();
        try
        {
            var settings = await SettingsReader.ReadAsync(db);
            DateTime? closed = settings.ClosedThrough?.Date;
            await EnsureAccountsAsync(db);
            var kinds = await db.JournalSourceTypes.AsNoTracking().ToDictionaryAsync(k => k.SourceType, k => k.TypeName);
            var existing = await db.JournalEntries.AsNoTracking()
                .Select(e => new ExistingEntry(e.EntryID, e.SourceType, e.SourceID, e.Signature, e.EntryDate, e.TotalDebit))
                .ToListAsync();
            var sources = new List<List<JournalSourceLine>>();
            foreach (var view in AppData.JournalSourceViews)
            {
                sources.Add(await reader.ReadAsync(db, view));
            }
            var plan = JournalPlanner.Plan(sources, existing, kinds, closed);

            int temp = 0;
            var added = new List<JournalEntries>();
            foreach (var h in plan.Added)
            {
                var e = new JournalEntries { EntryNumber = "~" + (++temp).ToString("000000") };
                Fill(e, h);
                db.JournalEntries.Add(e);
                added.Add(e);
            }
            await db.SaveChangesAsync();
            for (int i = 0; i < added.Count; i++)
            {
                AddLines(db, added[i].EntryID, plan.Added[i]);
            }
            foreach (var (entryId, h) in plan.Updated)
            {
                await db.JournalLines.Where(l => l.EntryID == entryId).ExecuteDeleteAsync();
                var e = await db.JournalEntries.FirstAsync(x => x.EntryID == entryId);
                Fill(e, h);
                e.UpdatedAt = DateTime.Now;
                AddLines(db, entryId, h);
            }
            await db.SaveChangesAsync();
            foreach (int entryId in plan.Removed)
            {
                await db.JournalLines.Where(l => l.EntryID == entryId).ExecuteDeleteAsync();
                await db.JournalEntries.Where(e => e.EntryID == entryId).ExecuteDeleteAsync();
            }
            // the numbers of the new entries, in date order
            foreach (var e in added.OrderBy(e => e.EntryDate).ThenBy(e => e.EntryID))
            {
                e.EntryNumber = await NumberSequences.NextNumberAsync(db, "JOURNAL");
            }
            await StampCurrenciesAsync(db);
            int changes = plan.Added.Count + plan.Updated.Count + plan.Removed.Count;
            if (changes > 0)
            {
                Audit.LogAction(db, user, "JOURNAL_SYNC", "JournalEntries", "",
                                $"Added={plan.Added.Count} Updated={plan.Updated.Count} Removed={plan.Removed.Count}");
            }
            await db.SaveChangesAsync();
            await tx.CommitAsync();
            return new JournalSyncResult(plan.Added.Count, plan.Updated.Count, plan.Removed.Count, Message(plan));
        }
        catch (Exception ex) when (ex is DbUpdateException or InvalidOperationException)
        {
            await tx.RollbackAsync();
            return new JournalSyncResult(0, 0, 0, "تعذر تحديث القيود: " + ex.Message);
        }
    }

    private static string Message(JournalPlan plan)
    {
        string text = "";
        if (plan.Unbalanced.Count > 0)
        {
            text = "عمليات غير متوازنة لم يُنشأ لها قيد (راجعها):\n" + Cut(string.Join("\n", plan.Unbalanced.Select(x => "  " + x)));
        }
        if (plan.Locked.Count > 0)
        {
            text += (text.Length > 0 ? "\n" : "") +
                    "عمليات في فترة مقفلة تغيرت ولم يتغير قيدها (أعد فتح الفترة إن كان التغيير مقصودًا):\n" +
                    Cut(string.Join("\n", plan.Locked.Select(x => "  " + x)));
        }
        return text;
    }

    private static string Cut(string s) => s.Length <= 700 ? s : s[..700];

    /// <summary>modJournal.FillHeader.</summary>
    private static void Fill(JournalEntries e, PlannedHeader h)
    {
        e.EntryDate = h.EntryDate;
        e.SourceType = h.SourceType;
        e.SourceID = h.SourceID;
        e.SourceNumber = h.SourceNumber;
        e.Description = h.Description;
        e.TotalDebit = h.TotalDebit;
        e.TotalCredit = h.TotalCredit;
        e.Signature = h.Signature;
        e.LineCount = (short)h.Lines.Count;
    }

    /// <summary>The lines of an entry; cost centre 0 in the source = not allocated (Null).</summary>
    private static void AddLines(RetailDbContext db, int entryId, PlannedHeader h)
    {
        foreach (var l in h.Lines)
        {
            db.JournalLines.Add(new JournalLines
            {
                EntryID = entryId, LineNumber = (short)l.LineOrder, AccountCode = l.AccountCode, Debit = l.Debit, Credit = l.Credit,
                LineText = l.LineText, CostCenterID = l.CostCenter == 0 ? null : l.CostCenter,
            });
        }
    }

    /// <summary>modJournal.EnsureAccounts: a sub-account for each cash box (110000 + box), expense type (530000 + type),
    /// bank (120000 + bank) and warehouse but the main one (140000 + warehouse), with their English names, then the
    /// levels of the tree (modAccounts.RebuildAccountTree).</summary>
    public static async Task EnsureAccountsAsync(RetailDbContext db)
    {
        var codes = (await db.Accounts.Select(a => a.AccountCode).ToListAsync()).ToHashSet();
        void Add(int code, string name, string? nameEn, string type, int parent)
        {
            if (codes.Add(code))
            {
                db.Accounts.Add(new Accounts
                {
                    AccountCode = code, AccountName = name, AccountType = type, ParentCode = parent, IsPosting = true, IsSystem = true,
                });
            }
        }
        foreach (var b in await db.CashBoxes.AsNoTracking().ToListAsync())
        {
            Add(110000 + b.CashBoxID, b.BoxName, b.BoxNameEn, "ASSET", 1100);
        }
        foreach (var t in await db.ExpenseTypes.AsNoTracking().ToListAsync())
        {
            Add(530000 + t.ExpenseTypeID, t.ExpenseTypeName, t.ExpenseTypeNameEn, "EXPENSE", 5300);
        }
        foreach (var k in await db.Banks.AsNoTracking().ToListAsync())
        {
            Add(120000 + k.BankID, k.BankName, k.BankNameEn, "ASSET", 1210);
        }
        foreach (var w in await db.Warehouses.AsNoTracking().Where(w => w.WarehouseID != 1).ToListAsync())
        {
            Add(140000 + w.WarehouseID, w.WarehouseName, w.WarehouseNameEn, "ASSET", 11);
        }
        await db.SaveChangesAsync();
        // the English name of a box, bank, expense type or warehouse reaches its account; a name typed on the account is kept
        var english = new Dictionary<int, string>();
        foreach (var b in await db.CashBoxes.AsNoTracking().Where(b => b.BoxNameEn != null).ToListAsync())
        {
            english[110000 + b.CashBoxID] = b.BoxNameEn!;
        }
        foreach (var t in await db.ExpenseTypes.AsNoTracking().Where(t => t.ExpenseTypeNameEn != null).ToListAsync())
        {
            english[530000 + t.ExpenseTypeID] = t.ExpenseTypeNameEn!;
        }
        foreach (var k in await db.Banks.AsNoTracking().Where(k => k.BankNameEn != null).ToListAsync())
        {
            english[120000 + k.BankID] = k.BankNameEn!;
        }
        foreach (var w in await db.Warehouses.AsNoTracking().Where(w => w.WarehouseID != 1 && w.WarehouseNameEn != null).ToListAsync())
        {
            english[140000 + w.WarehouseID] = w.WarehouseNameEn!;
        }
        var accounts = await db.Accounts.ToListAsync();
        foreach (var a in accounts.Where(a => a.AccountNameEn is null && english.ContainsKey(a.AccountCode)))
        {
            a.AccountNameEn = english[a.AccountCode];
        }
        var (places, _) = AccountTree.Build(accounts.ToDictionary(a => a.AccountCode, a => a.ParentCode ?? 0));
        foreach (var a in accounts)
        {
            if (!places.TryGetValue(a.AccountCode, out var p))
            {
                continue;
            }
            if (a.TreeKey != p.TreeKey || a.AccountLevel != p.Level || a.Level1Code != p.LevelCodes[0])
            {
                a.AccountLevel = (byte)p.Level;
                a.TreeKey = p.TreeKey;
                a.Level1Code = p.LevelCodes[0];
                a.Level2Code = p.LevelCodes[1];
                a.Level3Code = p.LevelCodes[2];
                a.Level4Code = p.LevelCodes[3];
                a.Level5Code = p.LevelCodes[4];
            }
        }
        await db.SaveChangesAsync();
    }

    /// <summary>modCurrency.StampJournalCurrencies: the currency, rate and foreign amount of a document entered in another
    /// currency reach its entry.</summary>
    public static async Task StampCurrenciesAsync(RetailDbContext db)
    {
        var docs = new List<(string Type, int Id, string? Code, decimal Rate, decimal Amount)>();
        docs.AddRange((await db.PurchaseInvoices.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.PurchaseInvoiceID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("PURCHASE", d.PurchaseInvoiceID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        docs.AddRange((await db.PurchaseReturns.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.PurchaseReturnID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("PURCHASE_RETURN", d.PurchaseReturnID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        docs.AddRange((await db.SupplierPayments.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.PaymentID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("SUPPLIER_PAYMENT", d.PaymentID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        docs.AddRange((await db.CustomerPayments.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.PaymentID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("CUSTOMER_PAYMENT", d.PaymentID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        docs.AddRange((await db.Expenses.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.ExpenseID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("EXPENSE", d.ExpenseID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        docs.AddRange((await db.ManualEntries.AsNoTracking().Where(d => d.CurrencyCode != null)
            .Select(d => new { d.ManualEntryID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount }).ToListAsync())
            .Select(d => ("MANUAL", d.ManualEntryID, d.CurrencyCode, d.ExchangeRate, d.ForeignAmount)));
        if (docs.Count == 0)
        {
            return;
        }
        var byKey = docs.ToDictionary(d => (d.Type, d.Id));
        string[] types = ["PURCHASE", "PURCHASE_RETURN", "SUPPLIER_PAYMENT", "CUSTOMER_PAYMENT", "EXPENSE", "MANUAL"];
        foreach (var e in await db.JournalEntries.Where(e => types.Contains(e.SourceType)).ToListAsync())
        {
            if (byKey.TryGetValue((e.SourceType, e.SourceID), out var d) &&
                (e.CurrencyCode is null || e.CurrencyCode != d.Code || e.ExchangeRate != d.Rate || e.ForeignAmount != d.Amount))
            {
                e.CurrencyCode = d.Code;
                e.ExchangeRate = d.Rate;
                e.ForeignAmount = d.Amount;
            }
        }
        await db.SaveChangesAsync();
    }
}
