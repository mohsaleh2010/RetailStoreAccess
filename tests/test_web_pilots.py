"""The pilot screens of the web redesign (docs/ERP-Redesign-Prompt.md, phase 4; docs/ui-inventory.md section 8): one screen of
each template built from the shared components, and no business logic changed:
A = DataScreen (every data screen of the web) with DataScreenDialog (EntityDialog), B = /pos, C = /ledger, D = the dashboard
with the charts MONTHLY_SALES / BEST_SELLING and the tables of the last invoices and LOW_STOCK."""
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UI = os.path.join(ROOT, "web", "src", "RetailStore.UI")
KIT = os.path.join(ROOT, "web", "src", "RetailStore.UI.Kit")


def read(*parts):
    """A file of RetailStore.UI, else of RetailStore.UI.Kit (the theme and erp.css, docs/ui-inventory.md section 17)."""
    path = os.path.join(UI, *parts)
    if not os.path.exists(path):
        path = os.path.join(KIT, *parts)
    with open(path, encoding="utf-8-sig") as fh:
        return fh.read()


PILOTS = {
    "A": ("Components", "DataScreen.razor"),
    "A-dialog": ("Components", "DataScreenDialog.razor"),
    "B": ("Pages", "Sales", "PosPage.razor"),
    "C": ("Pages", "Accounting", "LedgerPage.razor"),
    "D": ("Pages", "Home.razor"),
}


class PilotScreensTests(unittest.TestCase):

    def test_every_page_starts_with_page_header_and_has_no_old_markup(self):
        for key, parts in PILOTS.items():
            if key == "A-dialog":
                continue
            with self.subTest(key):
                text = read(*parts)
                self.assertIn("<PageHeader ", text)
                for old in ('class="screen"', 'class="btn', "<table", "<select", 'class="message', "dialog-backdrop"):
                    self.assertNotIn(old, text, "the markup of retailstore.css is gone from the pilot")

    def test_no_hex_colors_and_no_inline_styles(self):
        for key, parts in PILOTS.items():
            with self.subTest(key):
                text = read(*parts)
                self.assertNotRegex(text, r"#[0-9A-Fa-f]{6}\b", "colors come from ErpTheme")
                self.assertNotIn("style=", text.lower())

    def test_template_a_uses_the_grid_and_the_dialog(self):
        screen = read("Components", "DataScreen.razor")
        for part in ("<EntityGrid ", "Dialogs.ShowAsync<DataScreenDialog<T>>", "ConfirmDialog.AskAsync", "<KeyboardShortcuts OnNew"):
            self.assertIn(part, screen)
        dialog = read("Components", "DataScreenDialog.razor")
        self.assertIn("<EntityDialog ", dialog)
        self.assertIn("<LookupField ", dialog)
        self.assertIn("EntityDialog.StayOpen", dialog, "no to a question keeps the record open, as on the Access screen")

    def test_template_a_keeps_the_service_and_its_order(self):
        """Save: the service once, then again only after its question was answered yes; delete: deactivate after its question."""
        dialog = read("Components", "DataScreenDialog.razor")
        save = dialog[dialog.index("private async Task<string> SaveAsync"):]
        self.assertEqual(save.count("Service.SaveAsync(Spec, Record, IsNew"), 2)
        self.assertLess(save.index("false)"), save.index("SaveStatus.NeedsConfirmation"))
        self.assertLess(save.index("ConfirmDialog.AskAsync"), save.index("true)"))
        screen = read("Components", "DataScreen.razor")
        self.assertIn("Service.DeleteAsync(Spec, key, UserInfo.Acting(user), deactivate)", screen)
        self.assertIn("await Delete(key, true)", screen)
        self.assertIn("Permissions.OpenProblemAsync(UserInfo.EmployeeId(user), Spec.Name)", screen)

    def test_the_four_data_pages_use_the_template(self):
        for page in ("ProductsPage", "CustomersPage", "SuppliersPage", "WarehousesPage"):
            with self.subTest(page):
                self.assertIn("<DataScreen T=", read("Pages", page + ".razor"))

    def test_template_b_posts_through_the_sales_service(self):
        pos = read("Pages", "Sales", "PosPage.razor")
        for part in ("<InvoiceLinesEditor ", "<LookupField T=\"CustomerChoice\"", "Sales.PostSaleAsync(", "Sales.LinePriceAsync(",
                     "Sales.FindProductsAsync(", "Sales.UnitsAsync(", "OpenProblemAsync(EmployeeId, \"frmPOS\")",
                     "<KeyboardShortcuts OnSave", "erp-doc-side"):
            with self.subTest(part):
                self.assertIn(part, pos)
        self.assertIn("InvoiceLines.Calculate(", pos, "the totals come from the invoice engine")
        self.assertNotRegex(pos, r"Quantity\s*\*\s*\w*\.?Price", "no amount computed by the screen")
        self.assertIn("l.Price == TagOf(l).ListPrice", pos, "modPOS.RepriceCart: only the lines still at the program price")

    def test_template_c_is_the_statement_of_the_service(self):
        ledger = read("Pages", "Accounting", "LedgerPage.razor")
        for part in ("<FilterPanel ", "<StatCard ", "<EntityGrid T=\"LedgerLine\"", "Export=\"export\"", "Accounting.LedgerAsync(",
                     "AccountingService.PeriodProblem(", "AccountingService.BalanceText(", "OpenProblemAsync(UserInfo.EmployeeId(user), \"frmLedger\")"):
            with self.subTest(part):
                self.assertIn(part, ledger)

    def test_template_d_runs_the_reports_of_the_report_centre(self):
        home = read("Pages", "Home.razor")
        for part in ("<StatCard ", "<MudChart T=\"double\" ChartType=\"ChartType.Bar\"", "ChartType.Donut", "<EntityGrid T=\"InvoiceRow\"",
                     "<EntityGrid T=\"ReportRow\"", "CanOpenScreenAsync(emp, \"frmReportCenter\")", "CanOpenScreenAsync(emp, \"frmSalesInvoice\")",
                     "ErpTheme.ChartPalette(Mode.Dark)", "NavGroups.ForUserAsync("):
            with self.subTest(part):
                self.assertIn(part, home)
        reports = read("Services", "HomeReports.cs")
        for key in ("MONTHLY_SALES", "BEST_SELLING", "LOW_STOCK"):
            self.assertIn(f'reports.RunAsync(user, "{key}"', reports, "through ReportService: its permissions and queries")
        self.assertIn("services.AddScoped<HomeReports>()", read("ServiceCollectionExtensions.cs"))

    def test_the_report_keys_exist(self):
        with open(os.path.join(ROOT, "tools", "forms.py"), encoding="utf-8") as fh:
            forms = fh.read()
        for key in ("MONTHLY_SALES", "BEST_SELLING", "LOW_STOCK"):
            self.assertRegex(forms, r'["\']' + key + r'["\']')

    def test_chart_colors_come_from_the_theme(self):
        theme = read("Theme", "ErpTheme.cs")
        self.assertIn("public static string[] ChartPalette(bool dark)", theme)

    def test_the_css_of_the_templates_is_logical(self):
        css = read("wwwroot", "erp.css")
        part = css[css.index("the page templates"):css.index("@media (max-width")]
        self.assertNotRegex(part, r"#[0-9A-Fa-f]{3,8}\b")
        self.assertNotRegex(part, r"\b(margin|padding)-(left|right)\b")
        self.assertNotRegex(part, r"text-align:\s*(left|right)")
        self.assertNotRegex(part, r"\b(left|right):")

    def test_no_business_logic_changed(self):
        """The pilots do not touch the services: no service file mentions the redesign."""
        services = os.path.join(ROOT, "web", "src", "RetailStore.Data", "Services")
        for f in os.listdir(services):
            with open(os.path.join(services, f), encoding="utf-8-sig") as fh:
                self.assertNotIn("MudBlazor", fh.read(), f)


if __name__ == "__main__":
    unittest.main()
