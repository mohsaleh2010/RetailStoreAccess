using System.Globalization;
using RetailStore.Data.Entities;

namespace RetailStore.Data.Services;

/// <summary>One changed field of an audited operation: field, caption, value before, value after.</summary>
public sealed record AuditChange(string Field, string Caption, string OldValue, string NewValue);

/// <summary>The audit trail of modAudit / modCommon.LogAction: AuditLog (one row per operation) and AuditChanges
/// (the fields). The web writes the same rows as Access, so frmAuditLog shows both.</summary>
public static class Audit
{
    public const int MaxText = 255;
    public static readonly string[] SkippedFields = ["CreatedAt", "UpdatedAt", "PasswordSalt"];
    public static readonly string[] SecretFields = ["PasswordHash", "PasswordSalt"];

    private static string Left(string? s, int n) => s is null ? "" : (s.Length <= n ? s : s[..n]);

    /// <summary>modCommon.LogAction: adds the row; the caller saves (logging must never stop the user's work).</summary>
    public static void LogAction(RetailDbContext db, ActingUser user, string action, string objectName = "",
                                 string recordId = "", string details = "")
    {
        db.AuditLog.Add(new AuditLog
        {
            LogDate = DateTime.Now, EmployeeID = user.EmployeeID, ActionType = Left(action, 30),
            ObjectName = Left(objectName, 50), RecordID = Left(recordId, 30), Details = details,
            ComputerName = Left(user.Computer, 50),
        });
    }

    /// <summary>modAudit.WriteAudit: the operation, then its fields (saved here: AuditChanges needs the LogID).</summary>
    public static async Task WriteAsync(RetailDbContext db, ActingUser user, string action, string table, string recordId,
                                        string label, IReadOnlyList<AuditChange> changes)
    {
        var log = new AuditLog
        {
            LogDate = DateTime.Now, EmployeeID = user.EmployeeID, ActionType = Left(action, 30), ObjectName = Left(table, 50),
            RecordID = recordId.Length > 0 ? Left(recordId, 30) : null, RecordLabel = label.Length > 0 ? Left(label, 100) : null,
            ComputerName = Left(user.Computer, 50),
        };
        db.AuditLog.Add(log);
        await db.SaveChangesAsync();
        short n = 0;
        foreach (var c in changes)
        {
            n++;
            db.AuditChanges.Add(new AuditChanges
            {
                LogID = log.LogID, LineNo = n, FieldName = Left(c.Field, 64), FieldCaption = Left(c.Caption, 100),
                OldValue = c.OldValue.Length > 0 ? Left(c.OldValue, MaxText) : null,
                NewValue = c.NewValue.Length > 0 ? Left(c.NewValue, MaxText) : null,
            });
        }
        await db.SaveChangesAsync();
    }

    /// <summary>modAudit.AuditValueText.</summary>
    public static string ValueText(object? v)
    {
        string s = v switch
        {
            null => "",
            bool b => b ? "نعم" : "لا",
            DateTime d => d.TimeOfDay == TimeSpan.Zero ? d.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                                                       : d.ToString("yyyy/MM/dd HH:mm", CultureInfo.InvariantCulture),
            decimal m => m.ToString("0.############", CultureInfo.InvariantCulture),
            IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
            _ => v.ToString() ?? "",
        };
        return s.Length > MaxText ? s[..(MaxText - 3)] + "..." : s;
    }

    public static bool IsEmpty(object? v) => v is null || (v is string s && s.Length == 0);

    /// <summary>modAudit.SameValue.</summary>
    public static bool SameValue(object? a, object? b)
    {
        if (IsEmpty(a) || IsEmpty(b))
        {
            return IsEmpty(a) && IsEmpty(b);
        }
        if (a is string || b is string)
        {
            return a!.ToString() == b!.ToString();
        }
        return Equals(a, b) || (a is IConvertible && b is IConvertible &&
               Convert.ToDecimal(a, CultureInfo.InvariantCulture) == Convert.ToDecimal(b, CultureInfo.InvariantCulture));
    }

    /// <summary>modAudit.AuditSnapshot: every field of a record (field, caption, value), in the order of the table.</summary>
    public static List<(string Field, string Caption, object? Value)> Snapshot(string table, object entity)
    {
        var type = entity.GetType();
        var list = new List<(string, string, object?)>();
        foreach (var (field, caption) in TableFields.All[table])
        {
            if (SkippedFields.Contains(field) || SecretFields.Contains(field))
            {
                continue;
            }
            list.Add((field, caption, type.GetProperty(field)?.GetValue(entity)));
        }
        return list;
    }

    /// <summary>modAudit.SnapshotLabel / RecordLabel: the first field ending in Name, else Number, else Code.</summary>
    public static string Label(IEnumerable<(string Field, object? Value)> fields)
    {
        var all = fields.ToList();
        foreach (var end in new[] { "Name", "Number", "Code" })
        {
            foreach (var (field, value) in all)
            {
                if (field.Length > end.Length && field.EndsWith(end, StringComparison.Ordinal) && value is not null)
                {
                    return Left(ValueText(value), 100);
                }
            }
        }
        return "";
    }
}
