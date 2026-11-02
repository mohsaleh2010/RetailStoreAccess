using Microsoft.EntityFrameworkCore;
using RetailStore.Data.DataScreens;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;

namespace RetailStore.Tests;

public class DataScreenTests
{
    private static readonly ScreenSpec<Products> P = ScreenSpecs.Products;

    private static Products Product(string name, decimal price = 10, decimal cost = 6) =>
        new() { ProductName = name, SellingPrice = price, PurchasePrice = cost };

    [Fact]
    public void The_web_screens_are_the_Access_screens()
    {
        Assert.Equal(["frmProducts", "frmCustomers", "frmSuppliers", "frmWarehouses"], ScreenSpecs.All.Select(s => s.Name));
        Assert.Equal("ProductName", P.Field("ProductName")!.Name);
        Assert.True(P.Field("ProductName")!.Required);
        Assert.False(P.Field("ProductCode")!.Required);            // the sequence fills it
        Assert.True(P.Field("CurrentQuantity")!.Locked);
        Assert.Equal(3, P.Field("VATCategory")!.Options!.Length);
        Assert.Equal("Categories", P.Field("CategoryID")!.Lookup!.Table);
    }

    [Fact]
    public async Task A_new_product_gets_its_code_defaults_cost_and_audit()
    {
        using var t = new TestDb();
        var e = await t.Screens.NewAsync(P);
        Assert.True(e.IsActive && e.TrackStock);
        Assert.Equal("S", e.VATCategory);
        e.ProductName = "TEST قلم";
        e.SellingPrice = 5;
        e.PurchasePrice = 3;
        var r = await t.Screens.SaveAsync(P, e, true, TestDb.Admin);
        Assert.Equal(SaveStatus.Saved, r.Status);
        var saved = await t.Screens.LoadAsync(P, r.Key);
        Assert.Equal("P00001", saved!.ProductCode);
        Assert.Equal(3, saved.AverageCost);
        await using var db = t.CreateDbContext();
        var log = await db.AuditLog.SingleAsync(l => l.ActionType == "ADD");
        Assert.Equal("Products", log.ObjectName);
        Assert.Equal("TEST قلم", log.RecordLabel);
        Assert.Contains(await db.AuditChanges.Where(c => c.LogID == log.LogID).ToListAsync(),
                        c => c.FieldName == "CategoryID" && c.NewValue == "عام (1)");
    }

    [Fact]
    public async Task Required_unique_rules_and_the_question_on_a_price_under_the_cost()
    {
        using var t = new TestDb();
        var missing = await t.Screens.SaveAsync(P, new Products(), true, TestDb.Admin);
        Assert.Equal(SaveStatus.Refused, missing.Status);
        Assert.Equal("ProductName", missing.Field);
        var first = Product("A");
        first.Barcode = "111";
        Assert.Equal(SaveStatus.Saved, (await t.Screens.SaveAsync(P, first, true, TestDb.Admin)).Status);
        var twin = Product("B");
        twin.Barcode = "111";
        var dup = await t.Screens.SaveAsync(P, twin, true, TestDb.Admin);
        Assert.Equal(("Barcode", SaveStatus.Refused), (dup.Field, dup.Status));
        var negative = Product("C", price: -1);
        Assert.Equal("SellingPrice", (await t.Screens.SaveAsync(P, negative, true, TestDb.Admin)).Field);
        var cheap = Product("D", price: 2, cost: 5);
        var ask = await t.Screens.SaveAsync(P, cheap, true, TestDb.Admin);
        Assert.Equal(SaveStatus.NeedsConfirmation, ask.Status);
        Assert.Equal(SaveStatus.Saved, (await t.Screens.SaveAsync(P, cheap, true, TestDb.Admin, confirmed: true)).Status);
    }

    [Fact]
    public async Task An_edit_keeps_the_fields_that_are_not_on_the_screen_and_audits_the_change()
    {
        using var t = new TestDb();
        var r = await t.Screens.SaveAsync(P, Product("A"), true, TestDb.Admin);
        var e = (await t.Screens.LoadAsync(P, r.Key))!;
        await using (var db = t.CreateDbContext())        // meanwhile a purchase in Access changes the stock
        {
            (await db.Products.FirstAsync(x => x.ProductID == r.Key)).CurrentQuantity = 7;
            await db.SaveChangesAsync();
        }
        e.SellingPrice = 12;
        e.CurrentQuantity = 0;                             // locked on the screen: ignored
        Assert.Equal(SaveStatus.Saved, (await t.Screens.SaveAsync(P, e, false, TestDb.Admin)).Status);
        var now = (await t.Screens.LoadAsync(P, r.Key))!;
        Assert.Equal((12m, 7m), (now.SellingPrice, now.CurrentQuantity));
        Assert.NotNull(now.UpdatedAt);
        await using var check = t.CreateDbContext();
        var edit = await check.AuditLog.SingleAsync(l => l.ActionType == "EDIT");
        var change = await check.AuditChanges.SingleAsync(c => c.LogID == edit.LogID);
        Assert.Equal(("10", "12"), (change.OldValue, change.NewValue));
    }

    [Fact]
    public async Task Customers_balance_tax_number_and_mobile()
    {
        using var t = new TestDb();
        var c = new Customers { CustomerName = "TEST عميل", OpeningBalance = 50, VATNumber = "123" };
        var bad = await t.Screens.SaveAsync(ScreenSpecs.Customers, c, true, TestDb.Admin);
        Assert.Equal("VATNumber", bad.Field);
        c.VATNumber = "300000000000003";
        c.Mobile = "12345";
        Assert.Equal(SaveStatus.NeedsConfirmation, (await t.Screens.SaveAsync(ScreenSpecs.Customers, c, true, TestDb.Admin)).Status);
        var ok = await t.Screens.SaveAsync(ScreenSpecs.Customers, c, true, TestDb.Admin, confirmed: true);
        Assert.Equal(50, (await t.Screens.LoadAsync(ScreenSpecs.Customers, ok.Key))!.CurrentBalance);
        var cash = (await t.Screens.LoadAsync(ScreenSpecs.Customers, 1))!;
        Assert.Contains("AllowCredit", (await t.Screens.StateAsync(ScreenSpecs.Customers, cash, false)).LockedFields);
        Assert.Contains("أساسي", (await t.Screens.DeleteAsync(ScreenSpecs.Customers, 1, TestDb.Admin)).Message);
    }

    [Fact]
    public async Task A_record_in_use_is_deactivated_instead_of_deleted()
    {
        using var t = new TestDb();
        var s = await t.Screens.SaveAsync(ScreenSpecs.Suppliers, new Suppliers { SupplierName = "TEST مورد" }, true, TestDb.Admin);
        var p = Product("A");
        p.SupplierID = s.Key;
        var pr = await t.Screens.SaveAsync(P, p, true, TestDb.Admin);
        var ask = await t.Screens.DeleteAsync(ScreenSpecs.Suppliers, s.Key, TestDb.Admin);
        Assert.Equal(SaveStatus.NeedsConfirmation, ask.Status);
        Assert.Equal(SaveStatus.Saved, (await t.Screens.DeleteAsync(ScreenSpecs.Suppliers, s.Key, TestDb.Admin, deactivate: true)).Status);
        Assert.False((await t.Screens.LoadAsync(ScreenSpecs.Suppliers, s.Key))!.IsActive);
        Assert.Empty(await t.Screens.ListAsync(ScreenSpecs.Suppliers, "", false, false));
        Assert.Single(await t.Screens.ListAsync(ScreenSpecs.Suppliers, "مورد", true, false));
        Assert.Equal(SaveStatus.Saved, (await t.Screens.DeleteAsync(P, pr.Key, TestDb.Admin)).Status);
        await using var db = t.CreateDbContext();
        var log = await db.AuditLog.SingleAsync(l => l.ActionType == "DELETE");
        Assert.True(await db.AuditChanges.AnyAsync(c => c.LogID == log.LogID && c.FieldName == "ProductName" && c.OldValue == "A"));
        Assert.True(await db.AuditLog.AnyAsync(l => l.ActionType == "DEACTIVATE"));
    }

    [Fact]
    public async Task The_list_shows_value_list_texts_and_English_names()
    {
        using var t = new TestDb();
        var rows = await t.Screens.ListAsync(ScreenSpecs.Warehouses, "", false, false);
        Assert.Equal(["WH-1", "المخزن الرئيسي", "مخزن رئيسي"], rows.Single().Cells);
        var english = await t.Screens.ListAsync(ScreenSpecs.Warehouses, "", false, true);
        Assert.Equal("Main warehouse", english.Single().Cells[1]);
    }

    [Fact]
    public async Task A_cashier_cannot_save_a_product()
    {
        using var t = new TestDb();
        await using (var db = t.CreateDbContext())
        {
            db.Employees.Add(new Employees { EmployeeID = 2, EmployeeName = "C", Username = "c", RoleID = 3 });
            await db.SaveChangesAsync();
        }
        var r = await t.Screens.SaveAsync(P, Product("A"), true, new ActingUser(2, "TEST"));
        Assert.Equal(SaveStatus.Refused, r.Status);
        Assert.Contains("صلاحية", r.Message);
    }
}
