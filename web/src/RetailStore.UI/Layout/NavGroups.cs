using MudBlazor;
using RetailStore.Data.DataScreens;
using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.UI.Layout;

/// <summary>One entry of the side menu: a data screen (ScreenSpecs) or a work page (WorkPages), with the Access form whose
/// permission opens it.</summary>
public sealed record NavItem(string Route, string FormName, string Caption, string Icon);

/// <summary>A module of the side menu and its entries, in menu order.</summary>
public sealed record NavGroup(string Caption, string Icon, IReadOnlyList<NavItem> Items);

/// <summary>The side menu grouped by module (sales, purchases, stock, treasury, accounting, reports). The screens and
/// their permissions stay those of ScreenSpecs and WorkPages: this class only places each route in a group and gives it
/// a Material icon. A route not placed here goes to the last group, so a new page is never hidden.</summary>
public static class NavGroups
{
    private static readonly (string Caption, string Icon, (string Route, string Icon)[] Items)[] Layout =
    [
        ("المبيعات", Icons.Material.Outlined.PointOfSale,
        [
            ("pos", Icons.Material.Outlined.ShoppingCart),
            ("sales", Icons.Material.Outlined.ReceiptLong),
            ("payments", Icons.Material.Outlined.Payments),
            ("sales-returns", Icons.Material.Outlined.AssignmentReturn),
            ("customers", Icons.Material.Outlined.People),
        ]),
        ("المشتريات", Icons.Material.Outlined.LocalShipping,
        [
            ("purchase", Icons.Material.Outlined.AddShoppingCart),
            ("purchases", Icons.Material.Outlined.Receipt),
            ("supplier-payments", Icons.Material.Outlined.RequestQuote),
            ("purchase-returns", Icons.Material.Outlined.KeyboardReturn),
            ("suppliers", Icons.Material.Outlined.Storefront),
        ]),
        ("المخازن", Icons.Material.Outlined.Warehouse,
        [
            ("products", Icons.Material.Outlined.Inventory2),
            ("warehouses", Icons.Material.Outlined.Store),
            ("stock", Icons.Material.Outlined.Inventory),
            ("transfers", Icons.Material.Outlined.SyncAlt),
        ]),
        ("الخزينة والبنوك", Icons.Material.Outlined.AccountBalanceWallet,
        [
            ("treasury", Icons.Material.Outlined.Savings),
            ("cash-vouchers", Icons.Material.Outlined.Toll),
            ("cash-closing", Icons.Material.Outlined.LockClock),
            ("bank-transactions", Icons.Material.Outlined.AccountBalance),
            ("bank-reconciliation", Icons.Material.Outlined.FactCheck),
        ]),
        ("الحسابات", Icons.Material.Outlined.Calculate,
        [
            ("journal", Icons.Material.Outlined.MenuBook),
            ("manual-entries", Icons.Material.Outlined.EditNote),
            ("ledger", Icons.Material.Outlined.ListAlt),
            ("trial-balance", Icons.Material.Outlined.Balance),
            ("financials", Icons.Material.Outlined.QueryStats),
        ]),
        ("التقارير", Icons.Material.Outlined.Assessment,
        [
            ("reports", Icons.Material.Outlined.Summarize),
        ]),
    ];

    /// <summary>Every screen of the web: the data screens, then the work pages.</summary>
    public static IEnumerable<NavItem> AllItems() =>
        ScreenSpecs.All.Select(s => new NavItem(s.Route, s.Name, s.Caption, IconOf(s.Route)))
            .Concat(WorkPages.All.Select(w => new NavItem(w.Route, w.FormName, w.Caption, IconOf(w.Route))));

    /// <summary>The groups with the screens this user may open (PermissionService.CanOpenScreenAsync), empty groups left
    /// out. The permission of a form is asked once even when two pages share it.</summary>
    public static async Task<List<NavGroup>> ForUserAsync(PermissionService permissions, int employeeId)
    {
        var allowed = new List<NavItem>();
        if (employeeId > 0)
        {
            var answers = new Dictionary<string, bool>();
            foreach (var item in AllItems())
            {
                if (!answers.TryGetValue(item.FormName, out bool can))
                {
                    can = await permissions.CanOpenScreenAsync(employeeId, item.FormName);
                    answers[item.FormName] = can;
                }
                if (can)
                {
                    allowed.Add(item);
                }
            }
        }
        return Group(allowed);
    }

    /// <summary>Places the screens in their groups, in the order of Layout; a screen of no group goes to the last one.</summary>
    public static List<NavGroup> Group(IReadOnlyList<NavItem> items)
    {
        var byRoute = items.ToDictionary(i => i.Route);
        var placed = new HashSet<string>();
        var groups = new List<NavGroup>();
        for (int g = 0; g < Layout.Length; g++)
        {
            var (caption, icon, entries) = Layout[g];
            var list = new List<NavItem>();
            foreach (var (route, _) in entries)
            {
                if (byRoute.TryGetValue(route, out var item))
                {
                    list.Add(item);
                    placed.Add(route);
                }
            }
            if (g == Layout.Length - 1)
            {
                list.AddRange(items.Where(i => !placed.Contains(i.Route)));
            }
            if (list.Count > 0)
            {
                groups.Add(new NavGroup(caption, icon, list));
            }
        }
        return groups;
    }

    private static string IconOf(string route)
    {
        foreach (var (_, _, entries) in Layout)
        {
            foreach (var (r, icon) in entries)
            {
                if (r == route)
                {
                    return icon;
                }
            }
        }
        return Icons.Material.Outlined.Article;
    }
}
