"""The rules of the web redesign made permanent (docs/ERP-Redesign-Prompt.md phase 5; docs/ui-inventory.md section 9):
the skill .claude/skills/mudblazor-erp-ui names screens and components that exist, AGENTS.md points to it, and every page
marked approved (✅) in docs/ui-inventory.md follows the design rules of the skill, so each batch of phase 6 is checked
as soon as its rows are marked."""
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB_SRC = os.path.join(ROOT, "web", "src")
UI = os.path.join(WEB_SRC, "RetailStore.UI")
SHARED = os.path.join(UI, "Components", "Shared")
SKILL = os.path.join(ROOT, ".claude", "skills", "mudblazor-erp-ui", "SKILL.md")
INVENTORY = os.path.join(ROOT, "docs", "ui-inventory.md")

TEMPLATES = {
    "A": "Components/DataScreen.razor",
    "B": "Pages/Sales/PosPage.razor",
    "C": "Pages/Accounting/LedgerPage.razor",
    "D": "Pages/Home.razor",
}
FIELDS = ("MudTextField", "MudNumericField", "MudSelect", "MudDatePicker", "MudDateRangePicker", "MudAutocomplete")
OLD_MARKUP = ('class="screen"', 'class="btn', 'class="message', "<select", "<table", "dialog-backdrop")


def read(path):
    with open(path, encoding="utf-8-sig") as fh:
        return fh.read()


def ui_path(rel):
    """A path of the inventory or the skill: RetailStore.Web/... is under web/src, any other under RetailStore.UI."""
    if rel.startswith("RetailStore."):
        return os.path.join(WEB_SRC, *rel.split("/"))
    return os.path.join(UI, *rel.split("/"))


def open_tags(text, name):
    """The opening tag of each <name ...>: up to the next '<' (lambdas hold '>', never '<')."""
    for m in re.finditer(r"<%s\b" % name, text):
        end = text.find("<", m.end())
        yield text[m.start():end if end > 0 else len(text)]


def approved_pages():
    """The .razor files of the rows marked ✅ in the web table of the inventory (a DataScreen wrapper brings DataScreen)."""
    text = read(INVENTORY)
    table = text[text.index("## 2)"):text.index("## 3)")]
    files = []
    for row in table.splitlines():
        cells = [c.strip() for c in row.strip().strip("|").split("|")]
        if len(cells) < 7 or not cells[6].startswith("✅"):
            continue
        for rel in re.findall(r"`([^`]+\.razor)`", cells[2]):
            files.append(rel)
    return files


def page_problems(rel, text):
    problems = []
    wrapper = "<DataScreen T=" in text
    if not wrapper and "<PageHeader " not in text and not rel.endswith("Dialog.razor"):
        problems.append("does not start with PageHeader")
    for old in OLD_MARKUP:
        if old in text:
            problems.append("old markup " + old)
    if re.search(r"#[0-9A-Fa-f]{6}\b", text):
        problems.append("hex color (colors come from ErpTheme)")
    if "style=" in text.lower():
        problems.append("inline style")
    for kind in set(re.findall(r"Icons\.Material\.(\w+)\.", text)) - {"Outlined"}:
        problems.append("icon set " + kind + " (pages use Icons.Material.Outlined)")
    for word in ("ريال", "جنيه", "\"SAR\"", "\"EGP\""):
        if word in text:
            problems.append("currency written in the page: " + word)
    if re.search(r"Quantity\s*\*\s*\w*\.?Price", text):
        problems.append("an amount computed by the page (InvoiceCalculator)")
    problems += field_problems(text)
    return problems


def field_problems(text):
    problems = []
    text = re.sub(r"@\*.*?\*@", "", text, flags=re.S)  # the example comments of the components
    for name in FIELDS:
        for tag in open_tags(text, name):
            if 'Variant="Variant.Outlined"' not in tag:
                problems.append(name + " without Variant.Outlined: " + tag[:60])
            if 'Margin="Margin.Dense"' not in tag:
                problems.append(name + " without Margin.Dense: " + tag[:60])
    return problems


class SkillTests(unittest.TestCase):

    def setUp(self):
        self.skill = read(SKILL)

    def test_front_matter(self):
        head = self.skill.split("---")[1]
        self.assertIn("name: mudblazor-erp-ui", head)
        self.assertRegex(head, r"description: .*MudBlazor")

    def test_the_templates_exist_and_are_named(self):
        for key, rel in TEMPLATES.items():
            with self.subTest(key):
                self.assertTrue(os.path.isfile(ui_path(rel)), rel)
                self.assertRegex(self.skill, r"\*\*%s\*\*[^\n]*`%s`" % (key, re.escape(rel)))
        self.assertTrue(os.path.isfile(ui_path("Components/DataScreenDialog.razor")))

    def test_every_path_of_the_skill_exists(self):
        for rel in re.findall(r"`((?:Components|Pages|Services|Theme|Layout|wwwroot)/[\w./]+)`", self.skill):
            with self.subTest(rel):
                self.assertTrue(os.path.exists(ui_path(rel)), rel)

    def test_every_shared_component_is_in_the_skill(self):
        components = {f[:-6] for f in os.listdir(SHARED) if f.endswith(".razor")}
        table = self.skill[self.skill.index("## 2)"):self.skill.index("## 3)")]
        named = set(re.findall(r"^\| `(\w+)", table, re.M)) | set(re.findall(r"^\| `\w+` / `(\w+)`", table, re.M))
        self.assertEqual(components, named, "the table of section 2 lists every component of Components/Shared")

    def test_the_rules_of_the_skill(self):
        for rule in ('Variant="Variant.Outlined"', 'Margin="Margin.Dense"', "Spacing=\"3\"", "pa-4", "mb-4", "ErpTheme",
                     "Icons.Material.Outlined", "MoneyText", "InvoiceCalculator", "L.T(", "dotnet build", "1440px", "768px",
                     "Color.Error", "Variant.Filled"):
            with self.subTest(rule):
                self.assertIn(rule, self.skill)

    def test_agents_points_to_the_skill(self):
        agents = read(os.path.join(ROOT, "AGENTS.md"))
        self.assertIn(".claude/skills/mudblazor-erp-ui", agents)
        for rel in TEMPLATES.values():
            self.assertIn(rel.split("/")[-1], agents)
        claude = read(os.path.join(ROOT, "CLAUDE.md"))
        self.assertIn("mudblazor-erp-ui", claude)


class ApprovedPagesTests(unittest.TestCase):

    def test_the_templates_are_approved_in_the_inventory(self):
        pages = approved_pages()
        for rel in ("Pages/Home.razor", "Pages/ProductsPage.razor", "Pages/Sales/PosPage.razor", "Pages/Accounting/LedgerPage.razor"):
            self.assertIn(rel, pages)

    def test_every_approved_page_follows_the_rules(self):
        files = set(approved_pages())
        files |= {"Components/DataScreenDialog.razor"}
        for rel in sorted(files):
            path = ui_path(rel)
            with self.subTest(rel):
                self.assertTrue(os.path.isfile(path), rel)
                self.assertEqual(page_problems(rel, read(path)), [])

    def test_the_shared_fields_are_outlined_and_dense(self):
        for name in sorted(os.listdir(SHARED)):
            if name.endswith(".razor"):
                with self.subTest(name):
                    self.assertEqual(field_problems(read(os.path.join(SHARED, name))), [])

    def test_the_checks_find_a_bad_page(self):
        """The rules really catch what they forbid (a page written the old way)."""
        bad = ('<div class="screen"><MudTextField @bind-Value="x" Label="a" />'
               '<MudIcon Icon="@Icons.Material.Filled.Add" /><span style="color:#FF0000">ريال</span>'
               '@(line.Quantity * line.Price)</div>')
        problems = page_problems("Pages/X.razor", bad)
        for part in ("PageHeader", "old markup", "hex color", "inline style", "icon set Filled", "currency",
                     "InvoiceCalculator", "Variant.Outlined", "Margin.Dense"):
            with self.subTest(part):
                self.assertTrue(any(part in p for p in problems), problems)


if __name__ == "__main__":
    unittest.main()
