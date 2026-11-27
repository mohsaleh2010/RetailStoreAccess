"""Phase 7, batch 7A of the web redesign (docs/ui-inventory.md section 17): the design system moves into the Razor library
RetailStore.UI.Kit (the theme, the shared components, the language / currency / shortcut services, erp.css / erp.js and the
Cairo font), which references RetailStore.Domain only, so the Windows POS (RetailStore.PosApp) and the rep app
(RetailStore.RepApp, which never reaches the database: docs/78) can use it; the web keeps its namespaces through
RetailStore.UI. RetailStore.Pos and RetailStore.App get MudBlazor and the kit; RetailStore.AppPreview shows the screens
of both apps in a browser for development only. The services themselves are tested in xUnit (KitTests.cs)."""
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "web", "src")
KIT = os.path.join(SRC, "RetailStore.UI.Kit")


def read(*parts):
    with open(os.path.join(SRC, *parts), encoding="utf-8-sig") as fh:
        return fh.read()


def project_refs(project):
    return re.findall(r'<ProjectReference Include="\.\.\\([\w.]+)\\', read(project, project + ".csproj"))


class KitTests(unittest.TestCase):

    def test_the_kit_references_the_domain_only(self):
        project = read("RetailStore.UI.Kit", "RetailStore.UI.Kit.csproj")
        self.assertEqual(project_refs("RetailStore.UI.Kit"), ["RetailStore.Domain"])
        self.assertIn('<PackageReference Include="MudBlazor"', project)
        self.assertNotIn("EntityFramework", project)
        self.assertIn("<RootNamespace>RetailStore.UI</RootNamespace>", project, "the namespaces of the web do not change")

    def test_the_kit_holds_the_design_system(self):
        for rel in ("Theme/ErpTheme.cs", "Services/UiLanguage.cs", "Services/UiCurrency.cs", "Services/UiTheme.cs",
                    "Services/ShortcutService.cs", "Services/ErpMudLocalizer.cs", "Services/ICurrencySource.cs",
                    "Components/Shared/PageHeader.razor", "Components/Shared/ErpFormat.cs",
                    "Components/Shared/StatusStyles.cs", "Components/Shared/InvoiceLine.cs", "wwwroot/erp.css", "wwwroot/erp.js"):
            with self.subTest(rel):
                self.assertTrue(os.path.isfile(os.path.join(KIT, *rel.split("/"))), rel)
                self.assertFalse(os.path.exists(os.path.join(SRC, "RetailStore.UI", *rel.split("/"))), "moved, not copied")

    def test_no_file_of_the_kit_reads_the_database(self):
        for folder, _, files in os.walk(KIT):
            if os.sep + "bin" in folder or os.sep + "obj" in folder:
                continue
            for f in files:
                if f.endswith((".cs", ".razor")):
                    with self.subTest(f):
                        text = read(os.path.relpath(os.path.join(folder, f), SRC))
                        self.assertNotIn("RetailStore.Data", text)
                        self.assertNotIn("DbContext", text)

    def test_the_currency_has_a_source_without_the_database(self):
        self.assertIn("UiCurrency(ICurrencySource store", read("RetailStore.UI.Kit", "Services", "UiCurrency.cs"))
        self.assertIn("public interface IStoreInfo : ICurrencySource", read("RetailStore.UI", "Services", "StoreInfo.cs"))
        ext = read("RetailStore.UI", "ServiceCollectionExtensions.cs")
        self.assertIn("services.AddErpUiKit();", ext)
        self.assertIn("AddScoped<ICurrencySource>(sp => sp.GetRequiredService<IStoreInfo>())", ext)

    def test_the_cairo_font_is_inside_the_kit(self):
        css = read("RetailStore.UI.Kit", "wwwroot", "erp.css")
        fonts = re.findall(r"url\((fonts/[\w.-]+\.woff2)\)", css)
        self.assertEqual(len(fonts), 3, "Arabic, Latin and Latin extended")
        for f in fonts:
            with self.subTest(f):
                self.assertTrue(os.path.isfile(os.path.join(KIT, "wwwroot", *f.split("/"))))
        self.assertTrue(os.path.isfile(os.path.join(KIT, "wwwroot", "fonts", "OFL.txt")), "the licence goes with the font")
        self.assertIn("font-family: Cairo;", css)
        for page in (("RetailStore.Web", "Components", "App.razor"), ("RetailStore.Pos", "wwwroot", "index.html"),
                     ("RetailStore.App", "wwwroot", "index.html"), ("RetailStore.AppPreview", "Components", "App.razor")):
            with self.subTest(page[0]):
                self.assertNotIn("fonts.googleapis.com", read(*page), "the POS works without the internet")


class AppsTests(unittest.TestCase):

    def test_the_apps_use_the_kit_and_never_the_web_library(self):
        self.assertIn("RetailStore.UI.Kit", project_refs("RetailStore.PosApp"))
        self.assertIn("RetailStore.UI.Kit", project_refs("RetailStore.RepApp"))
        for app in ("RetailStore.PosApp", "RetailStore.RepApp"):
            with self.subTest(app):
                self.assertNotIn("RetailStore.UI", project_refs(app))
        self.assertNotIn("RetailStore.Data", project_refs("RetailStore.RepApp"), "the phone reaches the program by the API")
        self.assertEqual(project_refs("RetailStore.UI"), ["RetailStore.Data", "RetailStore.UI.Kit"])

    def test_the_apps_register_the_services_of_the_kit(self):
        pos = read("RetailStore.PosApp", "PosAppServices.cs")
        self.assertIn("services.AddErpUiKit();", pos)
        self.assertIn("AddSingleton<ICurrencySource, PosCurrencySource>()", pos)
        rep = read("RetailStore.RepApp", "RepAppServices.cs")
        self.assertIn("services.AddErpUiKit();", rep)
        self.assertIn("new FixedCurrencySource(\"\")", rep)

    def test_the_routes_give_the_language_of_the_session(self):
        for app in ("RetailStore.PosApp", "RetailStore.RepApp"):
            with self.subTest(app):
                routes = read(app, "Routes.razor")
                self.assertIn("@implements IDisposable", routes)
                self.assertIn('L.Set(S.English ? "en" : "ar")', routes)
                self.assertIn("S.Changed += Bind;", routes)
                self.assertIn("S.Changed -= Bind", routes)
        self.assertIn("Currency.Use(S.Snapshot?.Settings.Currency)", read("RetailStore.PosApp", "Routes.razor"))

    def test_the_maui_programs_add_mudblazor_after_the_kit(self):
        for host, call in (("RetailStore.Pos", "AddPosApp("), ("RetailStore.App", "AddRepApp(")):
            with self.subTest(host):
                program = read(host, "MauiProgram.cs")
                self.assertIn("using MudBlazor.Services;", program)
                self.assertLess(program.index(call), program.index("AddMudServices()"), "the Arabic MudLocalizer first")

    def test_the_host_pages_link_mudblazor_and_the_kit(self):
        for host, css in (("RetailStore.Pos", "RetailStore.PosApp/posapp.css"), ("RetailStore.App", "RetailStore.RepApp/repapp.css")):
            with self.subTest(host):
                page = read(host, "wwwroot", "index.html")
                order = ["_content/MudBlazor/MudBlazor.min.css", "_content/RetailStore.UI.Kit/erp.css", "_content/" + css,
                         "_framework/blazor.webview.js", "_content/MudBlazor/MudBlazor.min.js",
                         "_content/RetailStore.UI.Kit/erp.js"]
                positions = [page.index(x) for x in order]
                self.assertEqual(positions, sorted(positions))
        web = read("RetailStore.Web", "Components", "App.razor")
        self.assertIn('_content/RetailStore.UI.Kit/erp.css', web)
        self.assertIn('_content/RetailStore.UI.Kit/erp.js', web)
        self.assertNotIn('_content/RetailStore.UI/', web)


class PreviewTests(unittest.TestCase):

    def test_the_preview_is_for_development_only(self):
        project = read("RetailStore.AppPreview", "RetailStore.AppPreview.csproj")
        self.assertIn("<IsPublishable>false</IsPublishable>", project)
        self.assertEqual(sorted(project_refs("RetailStore.AppPreview")), ["RetailStore.PosApp", "RetailStore.RepApp"])
        with open(os.path.join(ROOT, "web", "RetailStore.slnx"), encoding="utf-8") as fh:
            self.assertIn("RetailStore.AppPreview.csproj", fh.read(), "built and checked with the web")
        for sln in ("RetailStore.Windows.slnx", "RetailStore.Mobile.slnx"):
            with open(os.path.join(ROOT, "web", sln), encoding="utf-8") as fh:
                self.assertNotIn("RetailStore.AppPreview", fh.read(), "never delivered")
        with open(os.path.join(ROOT, "tools", "deploy_web.py"), encoding="utf-8") as fh:
            self.assertNotIn("AppPreview", fh.read())

    def test_the_preview_uses_the_services_of_the_maui_programs(self):
        program = read("RetailStore.AppPreview", "Program.cs")
        for text in ("AddRepApp(", "new BasicPlatform", "AddPosApp(", "new BasicPosPlatform", "AddMudServices()",
                     "AddAdditionalAssemblies(app.Routes.Assembly)"):
            with self.subTest(text):
                self.assertIn(text, program)
        self.assertLess(program.index("AddPosApp("), program.index("AddMudServices()"))
        page = read("RetailStore.AppPreview", "Components", "App.razor")
        self.assertIn("prerender: false", page)
        for link in ("_content/MudBlazor/MudBlazor.min.css", "_content/RetailStore.UI.Kit/erp.css",
                     "_content/MudBlazor/MudBlazor.min.js", "_content/RetailStore.UI.Kit/erp.js"):
            with self.subTest(link):
                self.assertIn(link, page)
        launch = read("RetailStore.AppPreview", "Properties", "launchSettings.json")
        self.assertIn('"--app pos"', launch)
        self.assertIn('"--app rep"', launch)


if __name__ == "__main__":
    unittest.main()
