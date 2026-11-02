using Microsoft.EntityFrameworkCore;
using RetailStore.Domain;
using RetailStore.Domain.Security;

namespace RetailStore.Data.Services;

/// <summary>The result of a login: EmployeeID 0 and the message on failure.</summary>
public sealed record LoginResult(int EmployeeID, string Message, bool MustChangePassword, string EmployeeName)
{
    public bool Ok => EmployeeID > 0;
}

/// <summary>The login and passwords of modSecurity: the same users, roles and passwords as Access.</summary>
public sealed class SecurityService(IDbContextFactory<RetailDbContext> factory)
{
    public const string WrongLogin = "اسم المستخدم أو كلمة المرور غير صحيحة.";

    /// <summary>modSecurity.LoginUser: lock after AppData.MaxFailedLogins wrong passwords for AppData.LockMinutes;
    /// the same message for an unknown user and a wrong password; the built-in administrator (1) enters without a
    /// password only until one is set (then the password screen is forced).</summary>
    public async Task<LoginResult> LoginAsync(string username, string password, string computer, DateTime? now = null,
                                              string channel = "web")
    {
        DateTime at = now ?? DateTime.Now;
        await using var db = await factory.CreateDbContextAsync();
        string name = (username ?? "").Trim();
        var e = await db.Employees.FirstOrDefaultAsync(x => x.Username == name);
        var anonymous = new ActingUser(0, computer);
        if (e is null)
        {
            Audit.LogAction(db, anonymous, "LOGIN_FAILED", "Employees", "", "Unknown user: " + Left(username ?? "", 30));
            await db.SaveChangesAsync();
            return new LoginResult(0, WrongLogin, false, "");
        }
        int id = 0;
        string message = "";
        if (!e.IsActive)
        {
            message = "هذا الحساب معطّل. راجع مدير النظام.";
        }
        else if (e.LockedUntil is not null && e.LockedUntil > at)
        {
            message = "الحساب مقفل مؤقتًا بسبب محاولات دخول خاطئة حتى الساعة " + e.LockedUntil.Value.ToString("HH:mm") + ".";
        }
        else if (e.PasswordHash is null)
        {
            if (e.EmployeeID == 1 && password.Length == 0)
            {
                id = 1;
            }
            else
            {
                message = "لم تُعيَّن كلمة مرور لهذا المستخدم بعد. اطلب من مدير النظام تعيينها.";
            }
        }
        else if (!PasswordHasher.Matches(password, e.PasswordSalt ?? "", e.PasswordHash))
        {
            e.FailedLoginCount++;
            int left = AppData.MaxFailedLogins - e.FailedLoginCount;
            if (left <= 0)
            {
                e.LockedUntil = at.AddMinutes(AppData.LockMinutes);
                e.FailedLoginCount = 0;
                message = "كلمة المرور غير صحيحة. تم قفل الحساب " + AppData.LockMinutes + " دقيقة.";
            }
            else
            {
                message = WrongLogin;          // the screen must not reveal which names exist
            }
            Audit.LogAction(db, anonymous, "LOGIN_FAILED", "Employees", e.EmployeeID.ToString());
        }
        else
        {
            id = e.EmployeeID;
        }
        if (id > 0)
        {
            e.FailedLoginCount = 0;
            e.LockedUntil = null;
            e.LastLoginAt = at;
            Audit.LogAction(db, new ActingUser(id, computer), "LOGIN", "Employees", id.ToString(), channel);
        }
        await db.SaveChangesAsync();
        return id > 0
            ? new LoginResult(id, "", e.MustChangePassword || e.PasswordHash is null, e.EmployeeName)
            : new LoginResult(0, message, false, "");
    }

    /// <summary>modSecurity.NeedsPasswordChange.</summary>
    public async Task<bool> NeedsPasswordChangeAsync(int employeeId)
    {
        await using var db = await factory.CreateDbContextAsync();
        var e = await db.Employees.AsNoTracking().FirstOrDefaultAsync(x => x.EmployeeID == employeeId);
        return e is not null && (e.MustChangePassword || e.PasswordHash is null);
    }

    /// <summary>modSecurity.SetUserPassword: a new salt and digest, the lock cleared. "" on success.</summary>
    public async Task<string> SetUserPasswordAsync(int employeeId, string newPassword, bool mustChangeNext, ActingUser user)
    {
        await using var db = await factory.CreateDbContextAsync();
        var e = await db.Employees.FirstOrDefaultAsync(x => x.EmployeeID == employeeId);
        if (e is null)
        {
            return "المستخدم غير موجود.";
        }
        string problem = PasswordRules.Problem(newPassword, e.Username);
        if (problem.Length > 0)
        {
            return problem;
        }
        string salt = PasswordHasher.NewSalt();
        e.PasswordSalt = salt;
        e.PasswordHash = PasswordHasher.Hash(newPassword, salt);
        e.MustChangePassword = mustChangeNext;
        e.FailedLoginCount = 0;
        e.LockedUntil = null;
        Audit.LogAction(db, user, "PASSWORD_SET", "Employees", employeeId.ToString(), mustChangeNext ? "temporary" : "by user");
        await db.SaveChangesAsync();
        return "";
    }

    /// <summary>modSecurity.ChangeOwnPassword.</summary>
    public async Task<string> ChangeOwnPasswordAsync(ActingUser user, string oldPassword, string newPassword, string confirm,
                                                     bool requireOld = true)
    {
        if (user.EmployeeID == 0)
        {
            return "سجّل الدخول أولًا.";
        }
        string? hashNow;
        string salt;
        await using (var db = await factory.CreateDbContextAsync())
        {
            var e = await db.Employees.AsNoTracking().FirstOrDefaultAsync(x => x.EmployeeID == user.EmployeeID);
            if (e is null)
            {
                return "المستخدم غير موجود.";
            }
            hashNow = e.PasswordHash;
            salt = e.PasswordSalt ?? "";
        }
        if (requireOld && hashNow is not null && !PasswordHasher.Matches(oldPassword, salt, hashNow))
        {
            return "كلمة المرور الحالية غير صحيحة.";
        }
        if (newPassword != confirm)
        {
            return "تأكيد كلمة المرور لا يطابق كلمة المرور الجديدة.";
        }
        if (hashNow is not null && PasswordHasher.Matches(newPassword, salt, hashNow))
        {
            return "اختر كلمة مرور مختلفة عن الحالية.";
        }
        return await SetUserPasswordAsync(user.EmployeeID, newPassword, false, user);
    }

    /// <summary>The stamp of the user's password (API tokens carry it): a new password ends the tokens given before.</summary>
    public static string PasswordStamp(string? hash, string? salt) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes((hash ?? "") + "|" + (salt ?? ""))))[..16];

    private static string Left(string s, int n) => s.Length <= n ? s : s[..n];
}
