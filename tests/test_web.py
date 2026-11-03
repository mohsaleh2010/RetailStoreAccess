"""Plan 6A (docs/76): the web system in C# (web/).
  * the generated C# (tools/gen_web.py) is what the generator writes now;
  * the model has every table, field, key, size and relationship of tools/schema.py;
  * the web screens are the Access screens (forms.DATA_SCREENS) with their permissions;
  * every Arabic text of the web has its English (tools/i18n_web.py), none is left outside T("...");
  * no AI model name, no LICENSE_SECRET, .NET 10;
  * with dotnet on the PATH (skip: RETAIL_SKIP_DOTNET=1): the solution builds and its xUnit tests pass, among them
    the C# journal planner against the Python mirror of SyncJournal on the fixture business."""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, HERE)

import forms as F
import gen_web
import i18n
import i18n_web
import queries as Q
import relations as R
from i18n_en import EN
from schema import TABLES, table

WEB = os.path.join(ROOT, "web")


def read(rel):
    with open(os.path.join(ROOT, rel), encoding="utf-8-sig") as fh:
        return fh.read()


def all_web_files(ext=(".cs", ".razor", ".csproj", ".json", ".css")):
    out = []
    for folder, dirs, files in os.walk(WEB):
        dirs[:] = [d for d in dirs if d not in ("bin", "obj")]
        out += [os.path.join(folder, f) for f in files if f.endswith(ext)]
    return out


class GeneratedCodeTests(unittest.TestCase):

    def test_generated_files_are_current(self):
        for path, text in gen_web.build_all().items():
            with self.subTest(path):
                with open(os.path.join(ROOT, path), encoding="utf-8") as fh:
                    self.assertEqual(fh.read(), text if text.endswith("\n") else text + "\n",
                                     "run python3 tools/generate.py")

    def test_every_table_and_field_is_in_the_model(self):
        entities = read("web/src/RetailStore.Data/Generated/Entities.g.cs")
        context = read("web/src/RetailStore.Data/Generated/RetailDbContext.g.cs")
        self.assertEqual(len(re.findall(r"^public partial class ", entities, re.M)), len(TABLES))
        for t in TABLES:
            self.assertIn(f"public DbSet<{t.name}> {t.name} =>", context)
            self.assertIn(f'e.ToTable("{t.name}");', context)
            for f in t.fields:
                self.assertRegex(entities, rf"public \S+ @?{f.name} \{{ get; set; \}}", f"{t.name}.{f.name}")

    def test_types_nulls_and_defaults(self):
        f = next(x for x in table("Products").fields if x.name == "SellingPrice")
        self.assertEqual(gen_web.clr_type(f), "decimal")
        self.assertEqual(gen_web.default_literal(f), "0m")
        f = next(x for x in table("Products").fields if x.name == "SupplierID")
        self.assertEqual(gen_web.clr_type(f), "int?")
        f = next(x for x in table("Products").fields if x.name == "IsActive")
        self.assertEqual((gen_web.clr_type(f), gen_web.default_literal(f)), ("bool", "true"))
        f = next(x for x in table("Products").fields if x.name == "VATCategory")
        self.assertEqual(gen_web.default_literal(f), '"S"')
        f = next(x for x in table("Products").fields if x.name == "CreatedAt")
        self.assertEqual(gen_web.default_literal(f), "DateTime.Now")
        self.assertEqual(gen_web.ident("default"), "@default")

    def test_relationships_are_the_enforced_ones(self):
        context = read("web/src/RetailStore.Data/Generated/RetailDbContext.g.cs")
        rels = R.relations()
        self.assertEqual(context.count(".HasConstraintName("), len(rels))
        for r in rels:
            self.assertIn(f'.HasConstraintName("{r.name}")', context)
        for child, field in R.UNENFORCED:
            self.assertNotIn(f'.HasConstraintName("FK_{child}_{field}")', context)

    def test_rowversion_and_money_only_on_sql_server(self):
        context = read("web/src/RetailStore.Data/Generated/RetailDbContext.g.cs")
        self.assertEqual(context.count('if (sqlServer) e.Property<byte[]>("RowVer").IsRowVersion();'), len(TABLES))
        self.assertIn('HasColumnType("money")', context)
        self.assertNotIn("HasIndex", context.replace("if (!sqlServer) e.HasIndex", ""))

    def test_seed_rows_and_their_english_names(self):
        seed = read("web/src/RetailStore.Data/Generated/SeedData.g.cs")
        for t in TABLES:
            if t.seed_rows:
                self.assertIn(f"db.{t.name}.AddRange(", seed)
        self.assertIn('new Employees { EmployeeID = 1, EmployeeName = "مدير النظام"', seed)
        self.assertIn('WarehouseNameEn = "Main warehouse"', seed)
        self.assertIn('RoleNameEn = "System administrator"', seed)
        self.assertEqual(seed.count("new RolePermissions {"), len(table("RolePermissions").seed_rows))

    def test_journal_sources_and_permissions(self):
        app = read("web/src/RetailStore.Domain/Generated/AppData.g.cs")
        self.assertIn(", ".join(f'"{q}"' for q in Q.JOURNAL_SOURCE_QUERIES), app)
        for form, key in F.SCREEN_PERMISSIONS.items():
            self.assertIn(f'["{form}"] = "{key}",', app)


class WebScreenTests(unittest.TestCase):

    def test_screens_are_access_screens_with_permissions(self):
        names = {s.name for s in F.DATA_SCREENS}
        screens = {row[0] for row in table("Screens").seed_rows}
        for route, name in gen_web.WEB_SCREENS:
            self.assertIn(name, names)
            self.assertIn(name, screens)
            self.assertIn(name, F.SCREEN_PERMISSIONS)
            page = read(f"web/src/RetailStore.UI/Pages/{gen_web.table(name[3:]).name}Page.razor")
            self.assertIn(f'@page "/{route}"', page)
            self.assertIn("@attribute [Authorize]", page)

    def test_fields_rules_and_lookups(self):
        spec = read("web/src/RetailStore.Data/Generated/Screens.g.cs")
        self.assertIn('Name = "ProductName", Kind = FieldKind.Text', spec)
        self.assertRegex(spec, r'Name = "SellingPrice".*Min = 0m')
        self.assertRegex(spec, r'Name = "WarehouseType", Kind = FieldKind.Choice.*new\("VAN", "سيارة مندوب"\)')
        self.assertIn('Sequence = "PRODUCT_CODE", SequenceField = "ProductCode"', spec)
        self.assertIn('new LookupSpec("Currencies", "CurrencyCode", "CurrencyName", null)', spec)
        self.assertEqual(gen_web.rule_limits(next(f for f in table("Settings").fields if f.name == "VATRate")),
                         {"min": "0", "max": "1", "maxExclusive": True})
        self.assertEqual(gen_web.rule_limits(next(f for f in table("Customers").fields if f.name == "VATNumber")), {})

    def test_save_follows_form_before_update(self):
        code = read("web/src/RetailStore.Data/Services/DataScreenService.cs")
        order = ["ActionProblemAsync", "Required(spec, target)", "Limits(spec, target)", "ValidateProductAsync",
                 "UniqueAsync(db, spec, target", "AssignSequenceAsync(db, spec, target)", "Audit.WriteAsync"]
        save = code[code.index("public async Task<SaveResult> SaveAsync"):code.index("private static T CopyOf")]
        self.assertEqual([save.index(x) for x in order], sorted(save.index(x) for x in order))


class WebLanguageTests(unittest.TestCase):

    def test_every_web_text_has_an_english_text_and_none_is_unused(self):
        keys = i18n_web.keys()
        self.assertEqual(sorted(keys - set(EN) - set(i18n_web.WEB_EN)), [], "add these to tools/i18n_web.py")
        self.assertEqual(sorted(set(i18n_web.WEB_EN) - keys), [], "no longer used: remove them from tools/i18n_web.py")
        self.assertEqual(sorted(set(i18n_web.WEB_EN) & set(EN)), [], "already in tools/i18n_en.py")

    def test_english_texts_are_safe(self):
        for ar, en in i18n_web.WEB_EN.items():
            with self.subTest(ar):
                self.assertEqual(en, en.strip())
                self.assertFalse(set(en) & i18n.FORBIDDEN_IN_ENGLISH, en)
                en.encode("ascii")
                self.assertEqual(i18n.fragments(ar), [ar], "a key is one fragment")

    def test_razor_arabic_goes_through_t(self):
        for path in i18n_web.web_sources():
            if path.endswith(".razor"):
                with self.subTest(path):
                    self.assertEqual(i18n_web.razor_outside_t(i18n_web._read(path)), [])

    def test_nothing_arabic_is_left_in_english(self):
        words = i18n_web.dictionary()
        keys = i18n.ordered(words)
        left = [s for s in i18n_web.strings() if i18n.has_arabic(i18n.translate(s, words, keys))]
        self.assertEqual(left, [])


class WebRulesTests(unittest.TestCase):

    def test_no_model_name_no_secret_and_net10(self):
        for path in all_web_files():
            with self.subTest(path):
                with open(path, encoding="utf-8-sig") as fh:
                    text = fh.read()
                self.assertNotRegex(text.lower(), r"\b(claude|opus|sonnet|haiku|gpt-\d|chatgpt|gemini)\b")
                self.assertNotIn("LICENSE_SECRET", text)
                if path.endswith("RetailStore.App.csproj"):
                    self.assertIn("<TargetFrameworks>net10.0-android</TargetFrameworks>", text)
                elif path.endswith(".csproj"):
                    self.assertIn("<TargetFramework>net10.0</TargetFramework>", text)

    def test_sql_server_reads_the_views_of_docs_71(self):
        code = read("web/src/RetailStore.Data/Journal/JournalSources.cs")
        self.assertIn("AppData.JournalSourceViews.Contains(view)", code, "only known view names reach the SQL")
        self.assertIn("FROM [dbo].[", code)

    def test_passwords_and_cookies(self):
        program = read("web/src/RetailStore.Web/Program.cs")
        self.assertIn("options.Cookie.HttpOnly = true;", program)
        self.assertIn("UseAntiforgery", program)
        login = read("web/src/RetailStore.Web/AccountEndpoints.cs")
        self.assertIn("SafeReturn", login)
        hasher = read("web/src/RetailStore.Domain/Security/PasswordHasher.cs")
        self.assertIn("FixedTimeEquals", hasher)


class ApiTests(unittest.TestCase):
    """Plan 6B (docs/77): the REST API /api/v1."""

    API = "web/src/RetailStore.Web/Api/"

    def test_messages_of_the_rep_service_are_those_of_access(self):
        code = read("web/src/RetailStore.Data/Services/RepService.cs")
        vba = "".join(read(f"src/vba/{m}.bas") for m in ("modRepRoutes", "modSalesOrders", "modPriceLists"))
        own = {"هذا المستخدم ليس مندوبًا نشطًا. اربطه بمندوب في شاشة المندوبين.", "أضف أسطر الطلبية.",
               "رقم الطلبية في الجهاز أطول من 40 حرفًا.", "رقم الطلبية في الجهاز مستعمل لمستخدم آخر."}
        messages = set(re.findall(r'"([^"]*[\u0600-\u06FF][^"]*)"', code)) - own
        self.assertGreater(len(messages), 10)
        for m in messages:
            with self.subTest(m):
                self.assertIn(m, vba)

    def test_tokens_are_safe(self):
        security = read(self.API + "ApiSecurity.cs")
        self.assertIn("SecurityAlgorithms.HmacSha256", security)
        self.assertIn("RandomNumberGenerator.GetBytes(64)", security)
        self.assertIn("ValidateLifetime = true", security)
        self.assertIn("SecurityService.PasswordStamp", security, "a new password ends the old tokens")
        self.assertIn("!e.IsActive", security, "a disabled user loses his token")
        settings = json.loads(read("web/src/RetailStore.Web/appsettings.json"))
        self.assertEqual(settings["Api"]["JwtKey"], "", "never a signing key in the repository")
        self.assertIn("App_Data/", read("web/.gitignore"), "the key file of App_Data is never committed")
        setup = read(self.API + "ApiSetup.cs")
        self.assertIn("MapInboundClaims = false", setup)
        self.assertIn("RequireRateLimiting(ApiSetup.TokenLimit)", read(self.API + "ApiEndpoints.cs"))
        self.assertIn("DisableAgent().DisableMcp().DisableTelemetry()", setup, "the API document never leaves the shop")

    def test_errors_of_the_api_are_never_pages(self):
        program = read("web/src/RetailStore.Web/Program.cs")
        self.assertIn("app.UseWhen(context => !ApiSetup.IsApi(context)", program)
        self.assertIn("app.UseWhen(ApiSetup.IsApi, api => api.UseExceptionHandler());", program)

    def test_every_operation_has_its_permission(self):
        """Each mapped operation of the secured group names its permission (.Needs), except the user's own ones."""
        own = {"/auth/token", "/auth/password", "/me", "/profile"}
        for name in ("ApiEndpoints.cs", "RepApi.cs", "SalesApi.cs"):
            code = read(self.API + name)
            for m in re.finditer(r'\.Map(Get|Post|Put|Delete)\("([^"]*)"', code):
                route = m.group(2)
                end = code.find(".Map", m.end())
                body = code[m.end():end if end > 0 else len(code)]
                with self.subTest(f"{name} {m.group(1)} {route}"):
                    self.assertEqual(".Needs(" in body, route not in own)
                    self.assertIn(".WithSummary(", body)


class RepAppTests(unittest.TestCase):
    """Plan 4D (docs/78): the rep app on Android (MAUI Blazor Hybrid), offline first, synced through /api/v1/rep."""

    APP = "web/src/RetailStore.RepApp/"

    def test_the_phone_never_reaches_the_database(self):
        project = read(self.APP + "RetailStore.RepApp.csproj")
        self.assertNotIn("RetailStore.Data", project)
        self.assertNotIn("EntityFrameworkCore", project)
        self.assertIn("RetailStore.RepApp.csproj", read("web/src/RetailStore.App/RetailStore.App.csproj"))

    def test_the_solutions(self):
        main = read("web/RetailStore.slnx")
        self.assertIn("RetailStore.RepApp.csproj", main)
        self.assertNotIn("RetailStore.App.csproj", main, "the web builds on a machine without the MAUI workload")
        self.assertIn("RetailStore.App.csproj", read("web/RetailStore.Mobile.slnx"))

    def test_the_token_is_in_the_secure_storage(self):
        session = read(self.APP + "Offline/RepSession.cs")
        self.assertIn("platform.SetSecretAsync(TokenKey", session)
        self.assertNotRegex(session, r'store\.Set\("[^"]*Token')
        self.assertIn("SecureStorage.Default", read("web/src/RetailStore.App/Platforms/Android/PhonePlatform.cs"))

    def test_orders_are_sent_once_and_first(self):
        sync = read(self.APP + "Offline/SyncService.cs")
        self.assertIn("pending.Where(i => i.Kind == OutboxKinds.Order).Concat(", sync)
        self.assertIn("ClientRef", read(self.APP + "Offline/LocalStore.cs"))
        self.assertEqual(next(f for f in table("SalesOrders").fields if f.name == "ClientRef").size, 40)
        self.assertTrue(any(ix.fields == ["ClientRef"] and ix.unique and ix.ignore_nulls for ix in table("SalesOrders").indexes))
        self.assertIn("o.ClientRef == reference", read("web/src/RetailStore.Data/Services/RepService.cs"))

    def test_the_messages_of_an_order_line_are_those_of_access(self):
        draft = read(self.APP + "Offline/OrderDraft.cs")
        vba = read("src/vba/modSalesOrders.bas")
        for m in set(re.findall(r'"([^"]*[\u0600-\u06FF][^"]*)"', draft)) - {"أضف أسطر الطلبية."}:
            with self.subTest(m):
                self.assertIn(m, vba)


def journal_vectors(path):
    """The source lines of the fixture business and the entries of the Python mirror of SyncJournal."""
    from access_sqlite import AccessOnSqlite, day
    from sim import Store
    db = AccessOnSqlite()
    db.load_fixture()
    lines = {}
    for q in Q.JOURNAL_SOURCE_QUERIES:
        rows = db.con.execute("SELECT SourceType, SourceID, SourceNumber, SourceDate, Party, LineOrder, AccountCode, "
                              f"Debit, Credit, LineText, CostCenter FROM {q}").fetchall()
        lines[q] = [list(r) for r in rows]
    unbalanced = 0
    for q in Q.JOURNAL_SOURCE_QUERIES:
        for d, c in db.con.execute(f"SELECT Sum(Debit), Sum(Credit) FROM {q} GROUP BY SourceType, SourceID"):
            unbalanced += round(d or 0, 4) != round(c or 0, 4)
    kinds = dict(db.con.execute("SELECT SourceType, TypeName FROM JournalSourceTypes").fetchall())
    s = Store(db.con, now=day(0, 9))
    s.sync_journal()
    entries = [dict(SourceType=t, SourceID=i, EntryDate=when, Description=text, TotalDebit=debit, Signature=sig,
                    LineCount=n)
               for t, i, when, text, debit, sig, n in db.con.execute(
                   "SELECT SourceType, SourceID, EntryDate, Description, TotalDebit, Signature, LineCount FROM JournalEntries")]
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"kinds": kinds, "lines": lines, "entries": entries, "unbalanced": unbalanced}, fh, ensure_ascii=False)
    return len(entries)


@unittest.skipIf(shutil.which("dotnet") is None or os.environ.get("RETAIL_SKIP_DOTNET"),
                 "dotnet is not installed (or RETAIL_SKIP_DOTNET is set)")
class DotnetTests(unittest.TestCase):

    def test_solution_builds_and_xunit_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            vectors = os.path.join(tmp, "journal.json")
            self.assertGreater(journal_vectors(vectors), 10)
            env = dict(os.environ, RETAIL_JOURNAL_VECTORS=vectors, DOTNET_CLI_TELEMETRY_OPTOUT="1",
                       DOTNET_NOLOGO="1")
            r = subprocess.run(["dotnet", "test", os.path.join(WEB, "RetailStore.slnx"), "--nologo"],
                               cwd=WEB, env=env, capture_output=True, text=True, timeout=1500)
            tail = (r.stdout + r.stderr)[-4000:]
            self.assertEqual(r.returncode, 0, tail)
            self.assertRegex(r.stdout, r"Passed!\s+- Failed:\s+0")


if __name__ == "__main__":
    unittest.main()
