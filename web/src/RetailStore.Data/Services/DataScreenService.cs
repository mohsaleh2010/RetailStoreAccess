using System.Globalization;
using System.Linq.Expressions;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.DataScreens;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;

namespace RetailStore.Data.Services;

/// <summary>A row of the record list: key, the cells (in the columns of the screen), active or not.</summary>
public sealed record ListRow(int Key, string[] Cells, bool Active);

public enum SaveStatus { Saved, Refused, NeedsConfirmation }

/// <summary>The answer of a save or delete: Refused with the message (and the field to fix), or NeedsConfirmation
/// with the question: the screen asks it and calls again with confirmed = true.</summary>
public sealed record SaveResult(SaveStatus Status, string Message = "", string Field = "", int Key = 0)
{
    public static SaveResult Refuse(string message, string field = "") => new(SaveStatus.Refused, message, field);
    public static SaveResult Ask(string question, string field = "") => new(SaveStatus.NeedsConfirmation, question, field);
}

/// <summary>What the screen must lock on the current record (modForms.LockPartnerFields).</summary>
public sealed record RecordState(IReadOnlySet<string> LockedFields);

/// <summary>The data screens of modForms on the web: list and search, new record with its defaults, save with the
/// same checks in the same order as FormBeforeUpdate (permission, required, rules, the screen's own rules, unique
/// values, the sequence number), delete as DeleteRecord (system rows, related records -> deactivate), and the
/// audit trail of every add, change and delete.</summary>
public sealed class DataScreenService(IDbContextFactory<RetailDbContext> factory, PermissionService permissions)
{
    public const int MaxListRows = 1000;

    // ------------------------------------------------------------------ list

    public async Task<List<ListRow>> ListAsync<T>(ScreenSpec<T> spec, string? search, bool showInactive, bool english)
        where T : class, new()
    {
        await using var db = await factory.CreateDbContextAsync();
        IQueryable<T> q = spec.Query(db).AsNoTracking();
        if (spec.ActiveField is not null && !showInactive)
        {
            q = q.Where(e => EF.Property<bool>(e, spec.ActiveField));
        }
        string text = (search ?? "").Trim();
        if (text.Length > 0 && spec.Search.Length > 0)
        {
            q = q.Where(SearchPredicate<T>(spec.Search, text));
        }
        var rows = await spec.OrderBy(q).Take(MaxListRows).ToListAsync();
        return rows.Select(e => new ListRow(spec.KeyOf(e), spec.Columns.Select(c => Cell(spec, e, c.Field, english)).ToArray(),
                                            spec.ActiveField is null || (bool)(Prop(e, spec.ActiveField) ?? true))).ToList();
    }

    /// <summary>The records of a screen for the API: the same filter and search as the list, in its order, a page of
    /// them (take at most MaxListRows) and how many there are.</summary>
    public async Task<(int Total, List<T> Rows)> RecordsAsync<T>(ScreenSpec<T> spec, string? search, bool showInactive,
                                                                  int skip, int take) where T : class, new()
    {
        await using var db = await factory.CreateDbContextAsync();
        IQueryable<T> q = spec.Query(db).AsNoTracking();
        if (spec.ActiveField is not null && !showInactive)
        {
            q = q.Where(e => EF.Property<bool>(e, spec.ActiveField));
        }
        string text = (search ?? "").Trim();
        if (text.Length > 0 && spec.Search.Length > 0)
        {
            q = q.Where(SearchPredicate<T>(spec.Search, text));
        }
        int total = await q.CountAsync();
        var rows = await spec.OrderBy(q).Skip(Math.Max(skip, 0)).Take(Math.Clamp(take, 1, MaxListRows)).ToListAsync();
        return (total, rows);
    }

    private static Expression<Func<T, bool>> SearchPredicate<T>(IEnumerable<string> fields, string text)
    {
        var e = Expression.Parameter(typeof(T), "e");
        var property = typeof(EF).GetMethod(nameof(EF.Property))!.MakeGenericMethod(typeof(string));
        var contains = typeof(string).GetMethod(nameof(string.Contains), [typeof(string)])!;
        Expression? body = null;
        foreach (var f in fields)
        {
            var value = Expression.Call(property, e, Expression.Constant(f));
            var test = Expression.AndAlso(Expression.NotEqual(value, Expression.Constant(null, typeof(string))),
                                          Expression.Call(value, contains, Expression.Constant(text)));
            body = body is null ? test : Expression.OrElse(body, test);
        }
        return Expression.Lambda<Func<T, bool>>(body!, e);
    }

    private static object? Prop(object e, string field) => e.GetType().GetProperty(field)?.GetValue(e);

    /// <summary>The text of a cell: the English name of a master row in English, the text of a value list, numbers.</summary>
    public static string Cell<T>(ScreenSpec<T> spec, T e, string field, bool english) where T : class, new()
    {
        object? v = Prop(e, field);
        if (english && field == spec.NameField && spec.NameEnField is not null)
        {
            v = Prop(e, spec.NameEnField) is string en && en.Trim().Length > 0 ? en : v;
        }
        var f = spec.Field(field);
        if (f?.Options is not null && v is string code)
        {
            return f.Options.FirstOrDefault(o => o.Value == code)?.Text ?? code;
        }
        return v switch
        {
            null => "",
            decimal m => m.ToString(f?.Kind == FieldKind.Money ? "#,##0.00" : "#,##0.###", CultureInfo.InvariantCulture),
            bool b => b ? "✓" : "",
            DateTime d => d.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
            _ => v.ToString() ?? "",
        };
    }

    // ------------------------------------------------------------------ records

    public Task<T> NewAsync<T>(ScreenSpec<T> spec) where T : class, new() => Task.FromResult(new T());

    public async Task<T?> LoadAsync<T>(ScreenSpec<T> spec, int key) where T : class, new()
    {
        await using var db = await factory.CreateDbContextAsync();
        var e = await spec.FindAsync(db, key);
        if (e is not null)
        {
            db.Entry(e).State = EntityState.Detached;
        }
        return e;
    }

    public async Task<List<LookupItem>> LookupAsync(LookupSpec lookup, bool english)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await Lookups.LoadAsync(db, lookup.Table, english);
    }

    /// <summary>The fields locked on this record: the opening balance of a customer or supplier with movements,
    /// the credit and active flags of a system row (the cash customer).</summary>
    public async Task<RecordState> StateAsync<T>(ScreenSpec<T> spec, T e, bool isNew) where T : class, new()
    {
        var locked = new HashSet<string>();
        if (!isNew && e is Customers or Suppliers)
        {
            await using var db = await factory.CreateDbContextAsync();
            if (await PartnerHasMovementsAsync(db, e))
            {
                locked.Add("OpeningBalance");
            }
            if (e is Customers { IsSystem: true })
            {
                locked.Add("AllowCredit");
                locked.Add("IsActive");
            }
        }
        return new RecordState(locked);
    }

    private static async Task<bool> PartnerHasMovementsAsync(RetailDbContext db, object e) => e switch
    {
        Customers c => await db.SalesInvoices.AnyAsync(x => x.CustomerID == c.CustomerID)
                       || await db.SalesReturns.AnyAsync(x => x.CustomerID == c.CustomerID)
                       || await db.CustomerPayments.AnyAsync(x => x.CustomerID == c.CustomerID),
        Suppliers s => await db.PurchaseInvoices.AnyAsync(x => x.SupplierID == s.SupplierID)
                       || await db.PurchaseReturns.AnyAsync(x => x.SupplierID == s.SupplierID)
                       || await db.SupplierPayments.AnyAsync(x => x.SupplierID == s.SupplierID),
        _ => false,
    };

    // ------------------------------------------------------------------ save

    /// <summary>Saves the record typed on the screen. Only the fields of the screen are taken from it: the other
    /// fields keep what the database holds (a balance changed meanwhile by an invoice in Access stays right).</summary>
    public async Task<SaveResult> SaveAsync<T>(ScreenSpec<T> spec, T edited, bool isNew, ActingUser user, bool confirmed = false)
        where T : class, new()
    {
        string denied = await permissions.ActionProblemAsync(user.EmployeeID, spec.Name, isNew ? ScreenAction.Add : ScreenAction.Edit);
        if (denied.Length > 0)
        {
            return SaveResult.Refuse(denied);
        }
        if (await permissions.OpenProblemAsync(user.EmployeeID, spec.Name) is { Length: > 0 } cannotOpen)
        {
            return SaveResult.Refuse(cannotOpen);
        }
        await using var db = await factory.CreateDbContextAsync();
        T target;
        T? before = null;
        if (isNew)
        {
            target = edited;
        }
        else
        {
            target = await spec.FindAsync(db, spec.KeyOf(edited)) ?? throw new InvalidOperationException("Record not found");
            before = CopyOf(spec, target);
            var state = await StateAsync(spec, target, false);
            foreach (var f in spec.Fields.Where(f => f.IsData && !f.Locked && !state.LockedFields.Contains(f.Name)))
            {
                f.Set(target, f.Get(edited));
            }
        }
        if (Required(spec, target) is { } missing)
        {
            return missing;
        }
        if (Limits(spec, target) is { } outOfRange)
        {
            return outOfRange;
        }
        var own = target switch
        {
            Products p => await ValidateProductAsync(db, p, isNew, confirmed),
            Customers or Suppliers => await ValidatePartnerAsync(db, target, isNew, confirmed),
            _ => null,
        };
        if (own is not null)
        {
            return own;
        }
        if (await UniqueAsync(db, spec, target, isNew) is { } duplicate)
        {
            return duplicate;
        }
        await using var tx = await db.Database.BeginTransactionAsync();
        if (spec.Sequence is not null && spec.SequenceField is not null)
        {
            await AssignSequenceAsync(db, spec, target);         // last: a refused save wastes no number
        }
        typeof(T).GetProperty("UpdatedAt")?.SetValue(target, DateTime.Now);
        var changes = Changes(spec, before, target, await LookupTextsAsync(db, spec, before, target));
        if (isNew)
        {
            db.Add(target);
        }
        try
        {
            await db.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException)
        {
            return SaveResult.Refuse("غيّر مستخدم آخر هذا السجل الآن. أعد فتحه ثم عدّل.");
        }
        int key = spec.KeyOf(target);
        if (isNew || changes.Count > 0)
        {
            string label = Audit.Label(spec.Fields.Where(f => f.IsData).Select(f => (f.Name, f.Get(target))));
            await Audit.WriteAsync(db, user, isNew ? "ADD" : "EDIT", spec.Table, key.ToString(), label, changes);
        }
        await tx.CommitAsync();
        return new SaveResult(SaveStatus.Saved, "تم الحفظ", "", key);
    }

    private static T CopyOf<T>(ScreenSpec<T> spec, T e) where T : class, new()
    {
        var copy = new T();
        foreach (var p in typeof(T).GetProperties().Where(p => p.CanWrite))
        {
            p.SetValue(copy, p.GetValue(e));
        }
        return copy;
    }

    /// <summary>modForms.CheckRequired (the sequence field is filled later).</summary>
    private static SaveResult? Required<T>(ScreenSpec<T> spec, T e) where T : class, new()
    {
        foreach (var f in spec.Fields.Where(f => f.Required))
        {
            if (f.Get(e) is null || f.Get(e) is string s && s.Trim().Length == 0)
            {
                return SaveResult.Refuse("الحقل «" + f.Label + "» مطلوب.", f.Name);
            }
        }
        return null;
    }

    /// <summary>The field sizes and the Access validation rules of the fields (>=0, value lists ...).</summary>
    private static SaveResult? Limits<T>(ScreenSpec<T> spec, T e) where T : class, new()
    {
        foreach (var f in spec.Fields.Where(f => f.IsData))
        {
            object? v = f.Get(e);
            if (v is string s && f.MaxLength > 0 && s.Length > f.MaxLength)
            {
                return SaveResult.Refuse("الحقل «" + f.Label + "» أطول من " + f.MaxLength + " حرفًا.", f.Name);
            }
            if (v is string code && f.Options is not null && code.Length > 0 && f.Options.All(o => o.Value != code))
            {
                return SaveResult.Refuse("قيمة غير مقبولة في حقل «" + f.Label + "».", f.Name);
            }
            if (v is IConvertible and not string and not bool and not DateTime && (f.Min is not null || f.Max is not null))
            {
                decimal n = Convert.ToDecimal(v, CultureInfo.InvariantCulture);
                bool low = f.Min is { } min && (f.MinExclusive ? n <= min : n < min);
                bool high = f.Max is { } max && (f.MaxExclusive ? n >= max : n > max);
                if (low || high)
                {
                    return SaveResult.Refuse("القيمة في حقل «" + f.Label + "» خارج الحدود المسموحة.", f.Name);
                }
            }
        }
        return null;
    }

    /// <summary>modForms.ValidateProduct.</summary>
    private static async Task<SaveResult?> ValidateProductAsync(RetailDbContext db, Products p, bool isNew, bool confirmed)
    {
        if (isNew)
        {
            p.CurrentQuantity = 0;
        }
        if (p.CurrentQuantity == 0 && p.AverageCost == 0)
        {
            p.AverageCost = p.PurchasePrice;                     // no stock yet
        }
        string barcode = (p.Barcode ?? "").Trim();
        if (barcode.Length > 0 && await db.ProductUnits.AnyAsync(u => u.Barcode == barcode))
        {
            return SaveResult.Refuse("الباركود " + barcode + " مستخدم لوحدة من وحدات الأصناف.", "Barcode");   // modProductUnits
        }
        if (!confirmed && p.AverageCost > 0 && p.SellingPrice < p.AverageCost)
        {
            return SaveResult.Ask("سعر البيع (" + p.SellingPrice.ToString("#,##0.00", CultureInfo.InvariantCulture) +
                                  ") أقل من التكلفة (" + p.AverageCost.ToString("#,##0.00", CultureInfo.InvariantCulture) +
                                  "). هل تريد الحفظ على أي حال؟", "SellingPrice");
        }
        return null;
    }

    /// <summary>modForms.ValidatePartner (customers and suppliers).</summary>
    private static async Task<SaveResult?> ValidatePartnerAsync(RetailDbContext db, object e, bool isNew, bool confirmed)
    {
        string country = await SettingsReader.CountryAsync(db);
        var (mobileValue, vat, emailValue) = e switch
        {
            Customers c => (c.Mobile, c.VATNumber, c.Email),
            Suppliers s => (s.Mobile, s.VATNumber, s.Email),
            _ => throw new ArgumentException("Not a customer or supplier", nameof(e)),
        };
        string mobile = (mobileValue ?? "").Trim();
        if (!confirmed && mobile.Length > 0 && !CountryRules.MobileFits(mobile, country))
        {
            return SaveResult.Ask("رقم الجوال «" + mobile + "» ليس بصيغة " + CountryRules.CountryName(country) + " (" +
                                  CountryRules.MobileExample(country) + "). هل تريد الحفظ على أي حال؟", "Mobile");
        }
        string taxProblem = CountryRules.TaxNumberProblem(vat, country);
        if (taxProblem.Length > 0)
        {
            return SaveResult.Refuse(taxProblem, "VATNumber");
        }
        string email = (emailValue ?? "").Trim();
        if (email.Length > 0 && !CountryRules.EmailFits(email))
        {
            return SaveResult.Refuse("البريد الإلكتروني غير صحيح.", "Email");
        }
        if (e is Customers { IsSystem: true } cash)
        {
            cash.AllowCredit = false;
            cash.OpeningBalance = 0;
        }
        // without movements the balance is just the opening balance
        if (isNew || !await PartnerHasMovementsAsync(db, e))
        {
            switch (e)
            {
                case Customers c:
                    c.CurrentBalance = c.OpeningBalance;
                    break;
                case Suppliers s:
                    s.CurrentBalance = s.OpeningBalance;
                    break;
            }
        }
        return null;
    }

    /// <summary>modForms.CheckUnique.</summary>
    private static async Task<SaveResult?> UniqueAsync<T>(RetailDbContext db, ScreenSpec<T> spec, T e, bool isNew)
        where T : class, new()
    {
        int key = isNew ? 0 : spec.KeyOf(e);
        foreach (var name in spec.Unique)
        {
            if (Prop(e, name) is not string v)
            {
                continue;
            }
            bool taken = await spec.Query(db).AsNoTracking()
                .Where(x => EF.Property<string>(x, name) == v && EF.Property<int>(x, spec.Key) != key).AnyAsync();
            if (taken)
            {
                return SaveResult.Refuse("القيمة «" + v + "» في حقل «" + (spec.Field(name)?.Label ?? name) +
                                         "» مستخدمة لسجل آخر.", name);
            }
        }
        return null;
    }

    /// <summary>modForms.AssignSequence: an empty sequence field gets the next free number.</summary>
    private static async Task AssignSequenceAsync<T>(RetailDbContext db, ScreenSpec<T> spec, T e) where T : class, new()
    {
        var p = typeof(T).GetProperty(spec.SequenceField!)!;
        if (((string?)p.GetValue(e) ?? "").Trim().Length > 0)
        {
            return;
        }
        string candidate;
        int tries = 0;
        do
        {
            candidate = await NumberSequences.NextNumberAsync(db, spec.Sequence!);
            tries++;
        }
        while (tries < 1000 && await spec.Query(db).AsNoTracking()
                   .AnyAsync(x => EF.Property<string>(x, spec.SequenceField!) == candidate));
        p.SetValue(e, candidate);
    }

    /// <summary>The names of the chosen look-up rows: a combo is written "name (key)" (modAudit.ShownText).</summary>
    private static async Task<Dictionary<(string, string), string>> LookupTextsAsync<T>(RetailDbContext db, ScreenSpec<T> spec,
                                                                                    T? before, T after) where T : class, new()
    {
        var texts = new Dictionary<(string, string), string>();
        foreach (var f in spec.Fields.Where(f => f.Lookup is not null))
        {
            var keys = new[] { before is null ? null : f.Get(before), f.Get(after) }.Where(k => k is not null).ToList();
            if (keys.Count == 0)
            {
                continue;
            }
            foreach (var item in await Lookups.LoadAsync(db, f.Lookup!.Table, false))
            {
                if (keys.Any(k => k!.ToString() == item.Key.ToString()))
                {
                    texts[(f.Name, item.Key.ToString()!)] = item.Text;
                }
            }
        }
        return texts;
    }

    /// <summary>modAudit.AuditFormBefore: the fields of the screen that changed (all filled ones for a new record).</summary>
    private static List<AuditChange> Changes<T>(ScreenSpec<T> spec, T? before, T after,
                                                Dictionary<(string, string), string> lookupTexts) where T : class, new()
    {
        string Shown(FieldSpec<T> f, object? v)
        {
            string text = Audit.ValueText(v);
            if (v is null)
            {
                return text;
            }
            if (f.Options is not null)
            {
                var o = f.Options.FirstOrDefault(x => x.Value == v.ToString());
                return o is null ? text : o.Text + " (" + text + ")";
            }
            if (f.Lookup is not null && lookupTexts.TryGetValue((f.Name, v.ToString()!), out var name))
            {
                return name + " (" + text + ")";
            }
            return text;
        }
        var list = new List<AuditChange>();
        foreach (var f in spec.Fields.Where(f => f.IsData && !Audit.SkippedFields.Contains(f.Name)))
        {
            object? newV = f.Get(after);
            if (before is null)
            {
                if (!Audit.IsEmpty(newV))
                {
                    list.Add(new AuditChange(f.Name, f.Label, "", Shown(f, newV)));
                }
                continue;
            }
            object? oldV = f.Get(before);
            if (!Audit.SameValue(oldV, newV))
            {
                list.Add(new AuditChange(f.Name, f.Label, Shown(f, oldV), Shown(f, newV)));
            }
        }
        return list;
    }

    // ------------------------------------------------------------------ delete

    /// <summary>modForms.DeleteRecord: a system row is never deleted; a record used by operations is refused, and the
    /// screen offers to deactivate it instead (deactivate = true: through the save, so the checks run).</summary>
    public async Task<SaveResult> DeleteAsync<T>(ScreenSpec<T> spec, int key, ActingUser user, bool deactivate = false)
        where T : class, new()
    {
        string denied = await permissions.ActionProblemAsync(user.EmployeeID, spec.Name, ScreenAction.Delete);
        if (denied.Length > 0)
        {
            return SaveResult.Refuse(denied);
        }
        if (await permissions.OpenProblemAsync(user.EmployeeID, spec.Name) is { Length: > 0 } cannotOpen)
        {
            return SaveResult.Refuse(cannotOpen);
        }
        await using var db = await factory.CreateDbContextAsync();
        var e = await spec.FindAsync(db, key);
        if (e is null)
        {
            return SaveResult.Refuse("السجل غير موجود.");
        }
        if (Prop(e, "IsSystem") is true)
        {
            return SaveResult.Refuse("هذا سجل أساسي في النظام ولا يمكن حذفه.");
        }
        if (deactivate && spec.ActiveField is not null)
        {
            db.Entry(e).State = EntityState.Detached;
            typeof(T).GetProperty(spec.ActiveField)!.SetValue(e, false);
            var saved = await SaveAsync(spec, e, false, user, confirmed: true);
            if (saved.Status == SaveStatus.Saved)
            {
                Audit.LogAction(db, user, "DEACTIVATE", spec.Table, key.ToString());
                await db.SaveChangesAsync();
                return saved with { Message = "تم تعطيل السجل" };
            }
            return saved;
        }
        var snapshot = Audit.Snapshot(spec.Table, e);
        db.Remove(e);
        try
        {
            await db.SaveChangesAsync();
        }
        catch (DbUpdateException)
        {
            return spec.ActiveField is not null
                ? SaveResult.Ask("لا يمكن حذف هذا السجل لأنه مستخدم في عمليات سابقة. هل تريد تعطيله بدلًا من ذلك؟ " +
                                 "(يختفي من القوائم ويبقى في التقارير)")
                : SaveResult.Refuse("لا يمكن حذف هذا السجل لأنه مستخدم في عمليات سابقة.");
        }
        var changes = snapshot.Where(s => !Audit.IsEmpty(s.Value))
                              .Select(s => new AuditChange(s.Field, s.Caption, Audit.ValueText(s.Value), "")).ToList();
        await Audit.WriteAsync(db, user, "DELETE", spec.Table, key.ToString(),
                               Audit.Label(snapshot.Select(s => (s.Field, s.Value))), changes);
        return new SaveResult(SaveStatus.Saved, "تم الحذف", "", key);
    }
}
