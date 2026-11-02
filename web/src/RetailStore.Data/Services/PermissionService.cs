using Microsoft.EntityFrameworkCore;
using RetailStore.Domain;

namespace RetailStore.Data.Services;

public enum ScreenAction { Add, Edit, Delete }

/// <summary>The permissions of modCommon.HasPermission and modSecurity.CanOpenScreen / CanScreenAction, read from the
/// same tables (Roles, RolePermissions, Screens, UserScreens, Employees): a change in Access applies to the web too.</summary>
public sealed class PermissionService(IDbContextFactory<RetailDbContext> factory)
{
    private sealed record UserFacts(int RoleID, bool IsDeveloper, bool CustomScreens);

    private static async Task<UserFacts?> FactsAsync(RetailDbContext db, int employeeId) =>
        await db.Employees.AsNoTracking().Where(e => e.EmployeeID == employeeId && e.IsActive)
            .Select(e => new UserFacts(e.RoleID, e.IsDeveloper, e.CustomScreens)).FirstOrDefaultAsync();

    /// <summary>A user whose screens are set one by one (never an administrator or the programmer).</summary>
    private static bool UsesCustomScreens(UserFacts u) => u.CustomScreens && !u.IsDeveloper && u.RoleID != AppData.AdminRoleId;

    public async Task<bool> IsAdministratorAsync(int employeeId)
    {
        await using var db = await factory.CreateDbContextAsync();
        return (await FactsAsync(db, employeeId))?.RoleID == AppData.AdminRoleId;
    }

    /// <summary>modCommon.HasPermission: the programmer always; a user with his own screens, for a permission that
    /// opens screens: one of those screens is granted to him; otherwise the role has the permission.</summary>
    public async Task<bool> HasPermissionAsync(int employeeId, string permissionKey)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await HasPermissionAsync(db, employeeId, permissionKey);
    }

    private static async Task<bool> HasPermissionAsync(RetailDbContext db, int employeeId, string key)
    {
        var u = await FactsAsync(db, employeeId);
        if (u is null)
        {
            return false;
        }
        if (u.IsDeveloper)
        {
            return true;
        }
        if (u.RoleID != AppData.AdminRoleId && u.CustomScreens && await db.Screens.AnyAsync(s => s.PermissionKey == key))
        {
            return await (from us in db.UserScreens
                          join s in db.Screens on us.ScreenName equals s.ScreenName
                          where us.EmployeeID == employeeId && us.CanOpen && s.PermissionKey == key
                          select us).AnyAsync();
        }
        return await db.RolePermissions.AnyAsync(r => r.RoleID == u.RoleID && r.PermissionKey == key);
    }

    /// <summary>modSecurity.CanOpenScreen: "" when the user may open the screen, else the message.</summary>
    public async Task<string> OpenProblemAsync(int employeeId, string formName)
    {
        await using var db = await factory.CreateDbContextAsync();
        var u = await FactsAsync(db, employeeId);
        if (u is null)
        {
            return "سجّل الدخول أولًا.";
        }
        if (u.IsDeveloper)
        {
            return "";
        }
        var screen = await db.Screens.AsNoTracking().FirstOrDefaultAsync(s => s.ScreenName == formName);
        if (screen is not null && UsesCustomScreens(u))
        {
            bool can = await db.UserScreens.AnyAsync(s => s.EmployeeID == employeeId && s.ScreenName == formName && s.CanOpen);
            return can ? "" : "لا تملك صلاحية فتح شاشة «" + screen.ScreenTitle + "». راجع مدير النظام.";
        }
        if (!AppData.ScreenPermissions.TryGetValue(formName, out var key) || await HasPermissionAsync(db, employeeId, key))
        {
            return "";
        }
        string name = await db.Permissions.Where(p => p.PermissionKey == key).Select(p => p.PermissionName).FirstOrDefaultAsync() ?? key;
        return "لا تملك صلاحية «" + name + "». راجع مدير النظام.";
    }

    public async Task<bool> CanOpenScreenAsync(int employeeId, string formName) => (await OpenProblemAsync(employeeId, formName)).Length == 0;

    /// <summary>modSecurity.CanScreenAction: only a user with his own screen permissions is limited, and only in a
    /// screen that has the action. "" when allowed, else the message.</summary>
    public async Task<string> ActionProblemAsync(int employeeId, string formName, ScreenAction action)
    {
        await using var db = await factory.CreateDbContextAsync();
        var u = await FactsAsync(db, employeeId);
        if (u is null)
        {
            return "سجّل الدخول أولًا.";
        }
        if (!UsesCustomScreens(u))
        {
            return "";
        }
        var screen = await db.Screens.AsNoTracking().FirstOrDefaultAsync(s => s.ScreenName == formName);
        if (screen is null)
        {
            return "";
        }
        bool has = action switch { ScreenAction.Add => screen.HasAdd, ScreenAction.Edit => screen.HasEdit, _ => screen.HasDelete };
        if (!has)
        {
            return "";
        }
        var us = await db.UserScreens.AsNoTracking().FirstOrDefaultAsync(s => s.EmployeeID == employeeId && s.ScreenName == formName);
        bool can = us is not null && us.CanOpen && action switch
        {
            ScreenAction.Add => us.CanAdd, ScreenAction.Edit => us.CanEdit, _ => us.CanDelete,
        };
        string verb = action switch { ScreenAction.Add => "الإضافة والحفظ", ScreenAction.Edit => "التعديل", _ => "الحذف" };
        return can ? "" : "لا تملك صلاحية " + verb + " في شاشة «" + screen.ScreenTitle + "». راجع مدير النظام.";
    }
}
