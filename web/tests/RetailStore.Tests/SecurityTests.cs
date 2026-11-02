using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.Domain;

namespace RetailStore.Tests;

public class LoginTests
{
    [Fact]
    public async Task The_administrator_enters_without_a_password_only_until_one_is_set()
    {
        using var t = new TestDb();
        var first = await t.Security.LoginAsync("admin", "", "TEST");
        Assert.True(first.Ok);
        Assert.True(first.MustChangePassword);
        Assert.Equal("", await t.Security.ChangeOwnPasswordAsync(TestDb.Admin, "", "Store@2026", "Store@2026"));
        Assert.False((await t.Security.LoginAsync("admin", "", "TEST")).Ok);
        var ok = await t.Security.LoginAsync(" admin ", "Store@2026", "TEST");
        Assert.True(ok.Ok);
        Assert.False(ok.MustChangePassword);
        await using var db = t.CreateDbContext();
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "LOGIN" && l.EmployeeID == 1));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "PASSWORD_SET"));
    }

    [Fact]
    public async Task Five_wrong_passwords_lock_the_account_with_the_same_message_as_an_unknown_user()
    {
        using var t = new TestDb();
        await t.Security.SetUserPasswordAsync(1, "Store@2026", false, TestDb.Admin);
        var unknown = await t.Security.LoginAsync("nobody", "x", "TEST");
        Assert.Equal(SecurityService.WrongLogin, unknown.Message);
        var now = new DateTime(2026, 11, 2, 10, 0, 0);
        for (int i = 1; i < AppData.MaxFailedLogins; i++)
        {
            Assert.Equal(SecurityService.WrongLogin, (await t.Security.LoginAsync("admin", "wrong", "TEST", now)).Message);
        }
        var locked = await t.Security.LoginAsync("admin", "wrong", "TEST", now);
        Assert.Contains("قفل", locked.Message);
        Assert.False((await t.Security.LoginAsync("admin", "Store@2026", "TEST", now.AddMinutes(1))).Ok);
        Assert.True((await t.Security.LoginAsync("admin", "Store@2026", "TEST", now.AddMinutes(AppData.LockMinutes + 1))).Ok);
    }

    [Fact]
    public async Task An_inactive_user_cannot_enter_and_a_new_password_must_differ()
    {
        using var t = new TestDb();
        await using (var db = t.CreateDbContext())
        {
            db.Employees.Add(new Employees { EmployeeID = 2, EmployeeName = "كاشير", Username = "cashier", RoleID = 3 });
            await db.SaveChangesAsync();
        }
        Assert.Contains("لم تُعيَّن", (await t.Security.LoginAsync("cashier", "", "TEST")).Message);
        Assert.Equal("", await t.Security.SetUserPasswordAsync(2, "Cash@123", true, TestDb.Admin));
        var login = await t.Security.LoginAsync("cashier", "Cash@123", "TEST");
        Assert.True(login.MustChangePassword);
        var cashier = new ActingUser(2, "TEST");
        Assert.NotEqual("", await t.Security.ChangeOwnPasswordAsync(cashier, "bad", "Cash@456", "Cash@456"));
        Assert.NotEqual("", await t.Security.ChangeOwnPasswordAsync(cashier, "Cash@123", "Cash@123", "Cash@123"));
        Assert.NotEqual("", await t.Security.ChangeOwnPasswordAsync(cashier, "Cash@123", "Cash@456", "Cash@457"));
        Assert.Equal("", await t.Security.ChangeOwnPasswordAsync(cashier, "Cash@123", "Cash@456", "Cash@456"));
        await using (var db = t.CreateDbContext())
        {
            (await db.Employees.FirstAsync(e => e.EmployeeID == 2)).IsActive = false;
            await db.SaveChangesAsync();
        }
        Assert.Contains("معطّل", (await t.Security.LoginAsync("cashier", "Cash@456", "TEST")).Message);
    }
}

public class PermissionTests
{
    private static async Task AddUser(TestDb t, int id, int role, bool custom = false)
    {
        await using var db = t.CreateDbContext();
        db.Employees.Add(new Employees { EmployeeID = id, EmployeeName = "U" + id, Username = "u" + id, RoleID = role, CustomScreens = custom });
        await db.SaveChangesAsync();
    }

    [Fact]
    public async Task Roles_open_the_screens_of_their_permissions()
    {
        using var t = new TestDb();
        await AddUser(t, 2, 3);                                   // cashier: customers, not products
        Assert.True(await t.Permissions.CanOpenScreenAsync(1, "frmProducts"));
        Assert.True(await t.Permissions.CanOpenScreenAsync(2, "frmCustomers"));
        Assert.False(await t.Permissions.CanOpenScreenAsync(2, "frmProducts"));
        Assert.Contains("لا تملك صلاحية", await t.Permissions.OpenProblemAsync(2, "frmSuppliers"));
        Assert.True(await t.Permissions.HasPermissionAsync(2, "SALES_POS"));
        Assert.False(await t.Permissions.HasPermissionAsync(99, "SALES_POS"));
        Assert.Equal("", await t.Permissions.ActionProblemAsync(2, "frmCustomers", ScreenAction.Delete));
    }

    [Fact]
    public async Task A_user_with_his_own_screens_gets_only_what_is_ticked()
    {
        using var t = new TestDb();
        await AddUser(t, 2, 2, custom: true);
        await using (var db = t.CreateDbContext())
        {
            db.UserScreens.Add(new UserScreens { EmployeeID = 2, ScreenName = "frmProducts", CanOpen = true, CanAdd = true });
            await db.SaveChangesAsync();
        }
        Assert.True(await t.Permissions.CanOpenScreenAsync(2, "frmProducts"));
        Assert.False(await t.Permissions.CanOpenScreenAsync(2, "frmCustomers"));
        Assert.Equal("", await t.Permissions.ActionProblemAsync(2, "frmProducts", ScreenAction.Add));
        Assert.Contains("التعديل", await t.Permissions.ActionProblemAsync(2, "frmProducts", ScreenAction.Edit));
        Assert.Contains("الحذف", await t.Permissions.ActionProblemAsync(2, "frmProducts", ScreenAction.Delete));
        Assert.True(await t.Permissions.HasPermissionAsync(2, "PRODUCTS"));
        Assert.False(await t.Permissions.HasPermissionAsync(2, "CUSTOMERS"));
    }
}
