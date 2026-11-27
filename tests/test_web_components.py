"""The shared components of the web redesign (docs/ERP-Redesign-Prompt.md, phase 3; docs/ui-inventory.md section 7):
every component is there and documented with an example, the colors come from the theme only, the currency from the
settings, the keyboard shortcuts reach the browser, and the gallery page /dev/components lives only in Development.
The logic of the components (formats, lookup, CSV, lines, shortcuts) is tested in xUnit (ComponentsTests.cs)."""
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# the shared components, the theme, their services and erp.css / erp.js live in RetailStore.UI.Kit (docs/ui-inventory.md
# section 17), shared by the web, the Windows POS and the rep app
KIT = os.path.join(ROOT, "web", "src", "RetailStore.UI.Kit")
SHARED = os.path.join(KIT, "Components", "Shared")
UI = os.path.join(ROOT, "web", "src", "RetailStore.UI")


def read(*parts):
    with open(os.path.join(ROOT, *parts), encoding="utf-8-sig") as fh:
        return fh.read()


COMPONENTS = ["PageHeader", "EntityGrid", "EntityDialog", "LookupField", "InvoiceLinesEditor", "FilterPanel",
              "MoneyText", "QtyText", "StatusChip", "ConfirmDialog", "StatCard", "EmptyState", "KeyboardShortcuts"]


class SharedComponentsTests(unittest.TestCase):

    def test_every_component_exists_with_its_example(self):
        for name in COMPONENTS:
            with self.subTest(name):
                text = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", name + ".razor")
                comment = text[text.index("@*"):text.index("*@")]
                self.assertIn("Example", comment, "a short comment with an example of use")
                self.assertIn(name, comment, "the example uses the component")

    def test_no_hex_colors_and_no_inline_styles_in_components(self):
        for f in os.listdir(SHARED):
            with self.subTest(f):
                text = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", f)
                self.assertNotRegex(text, r"#[0-9A-Fa-f]{3,8}\b", "colors come from ErpTheme")
                self.assertNotIn("style=", text.lower())

    def test_the_css_of_the_components_uses_the_palette(self):
        css = read("web", "src", "RetailStore.UI.Kit", "wwwroot", "erp.css")
        part = css[css.index("the shared components"):css.index("@media (max-width")]
        self.assertNotRegex(part, r"#[0-9A-Fa-f]{3,8}\b")
        self.assertNotRegex(part, r"\b(margin|padding)-(left|right)\b", "logical properties for rtl and ltr")
        self.assertNotRegex(part, r"text-align:\s*(left|right)")

    def test_amounts_take_the_currency_of_the_settings(self):
        money = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "MoneyText.razor")
        self.assertIn("Currency.Word", money)
        self.assertNotRegex(money, r"ريال|جنيه|SAR|EGP")
        currency = read("web", "src", "RetailStore.UI.Kit", "Services", "UiCurrency.cs")
        self.assertIn("store.CurrencyCodeAsync()", currency)
        info = read("web", "src", "RetailStore.UI", "Services", "StoreInfo.cs")
        self.assertIn("s.CurrencyCode", info)
        self.assertIn("CountryRules.Normalize(s.CountryCode)", info)

    def test_the_lines_use_the_invoice_engine(self):
        lines = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "InvoiceLine.cs")
        self.assertIn("InvoiceCalculator.Calc(", lines)
        editor = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "InvoiceLinesEditor.razor")
        self.assertNotRegex(editor, r"Quantity\s*\*\s*\w*\.?Price", "no amount computed by the screen")
        self.assertIn("@L.T(Result.Problem)", editor)

    def test_the_extra_column_of_the_lines(self):
        """ExtraColumn (phase 6 batch 2): a column of the page after the discount, a cell of the keyboard, counted in the
        footer; the purchase invoice puts the new selling price in it and keeps it in the Tag of the line."""
        editor = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "InvoiceLinesEditor.razor")
        self.assertIn("public RenderFragment<InvoiceLine<TProduct>>? ExtraColumn", editor)
        self.assertIn("public string? ExtraColumnTitle", editor)
        self.assertIn('<td data-cell class="erp-lines-extra">', editor)
        self.assertIn("(ExtraColumn is null ? 0 : 1)", editor, "the footer spans the extra column")
        self.assertLess(editor.index('erp-lines-disc">@L.T'), editor.index("erp-lines-extra"), "after the discount")
        self.assertIn(".erp-lines-extra", read("web", "src", "RetailStore.UI.Kit", "wwwroot", "erp.css"))
        page = read("web", "src", "RetailStore.UI", "Pages", "Purchases", "PurchasePage.razor")
        self.assertIn('ExtraColumnTitle="@L.T("سعر البيع الجديد")"', page)
        self.assertIn("TagOf(l).NewPrice", page, "the new price reaches PurchaseLine.NewSellingPrice")
        self.assertIn("<ExtraColumn", read("web", "src", "RetailStore.Web", "Components", "Pages", "Dev", "ComponentsGallery.razor"))

    def test_dates_are_gregorian(self):
        panel = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "FilterPanel.razor")
        self.assertIn('Culture="@ErpFormat.DateCulture(L.English)"', panel)
        self.assertIn("GregorianCalendar", read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "ErpFormat.cs"))


class ShortcutsTests(unittest.TestCase):

    def test_the_keys_reach_the_browser(self):
        service = read("web", "src", "RetailStore.UI.Kit", "Services", "ShortcutService.cs")
        js = read("web", "src", "RetailStore.UI.Kit", "wwwroot", "erp.js")
        for key in ("F2", "F3", "F4", "Escape", "Ctrl+P"):
            with self.subTest(key):
                self.assertIn(f'"{key}"', service)
                self.assertIn(f'"{key}"', js)
        self.assertIn("[JSInvokable]", service)
        self.assertIn('invokeMethodAsync("OnKey"', js)
        self.assertIn("dialogOpen()", js, "Esc in a dialog is left to MudBlazor")
        self.assertIn('CloseOnEscapeKey="true"', read("web", "src", "RetailStore.UI", "Layout", "MainLayout.razor"))

    def test_the_lines_keyboard(self):
        editor = read("web", "src", "RetailStore.UI.Kit", "Components", "Shared", "InvoiceLinesEditor.razor")
        js = read("web", "src", "RetailStore.UI.Kit", "wwwroot", "erp.js")
        for method in ("AddLineFromKeyboard", "RemoveLineFromKeyboard"):
            with self.subTest(method):
                self.assertIn(f"public async Task {method}(", editor)
                self.assertIn(f'"{method}"', js)

    def test_services_registered(self):
        ext = read("web", "src", "RetailStore.UI.Kit", "KitServiceCollectionExtensions.cs")
        for name in ("UiCurrency", "ShortcutService", "ErpMudLocalizer"):
            self.assertIn(name, ext)


class GalleryTests(unittest.TestCase):

    PAGE = ("web", "src", "RetailStore.Web", "Components", "Pages", "Dev", "ComponentsGallery.razor")

    def test_only_in_development_and_not_in_the_menu(self):
        page = read(*self.PAGE)
        self.assertIn('@page "/dev/components"', page)
        self.assertIn("@attribute [Authorize]", page)
        self.assertIn("Env.IsDevelopment()", page)
        self.assertIn("Nav.NotFound()", page)
        for f in ("Layout/NavGroups.cs", "Services/WorkPages.cs"):
            self.assertNotIn("dev/components", read("web", "src", "RetailStore.UI", *f.split("/")))

    def test_the_gallery_shows_every_component(self):
        page = read(*self.PAGE) + read("web", "src", "RetailStore.Web", "Components", "Pages", "Dev", "DemoItemDialog.razor")
        for name in COMPONENTS:
            with self.subTest(name):
                self.assertTrue(re.search(r"<" + name + r"\b", page) or name + "." in page, name)


if __name__ == "__main__":
    unittest.main()
