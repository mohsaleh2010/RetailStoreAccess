"""Phase 5 checks: screen layout, event wiring, bound fields, list/search/report
SQL on the SQLite mirror, and static checks across every VBA module.

Run:  python3 -m unittest discover -s tests -v
"""

import glob
import os
import re
import unittest

from access_sqlite import AccessOnSqlite
from helpers import ROOT, VbaModuleChecks, logical_lines
import forms as F
import gen_forms
import queries as Q
from schema import table

MODELS = F.all_forms()
SRC = os.path.join(ROOT, "src", "vba")

# Fields that the code fills in, so they need no input on the screen.
AUTO_FILLED = {("Expenses", "EmployeeID"), ("Expenses", "ExpenseNumber"),
               ("Products", "ProductCode"), ("Settings", "SettingID")}


def read(name):
    with open(os.path.join(SRC, name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def to_sqlite(sql: str) -> str:
    """Access-only syntax used by row sources -> SQLite."""
    from i18n import resolve_names
    sql = resolve_names(sql).replace(" & ", " || ")

    def like(m):
        pattern = m.group(1).replace("*", "%").replace("[%]", "*")
        return f"LIKE '{pattern}'"
    return re.sub(r"Like '([^']*)'", like, sql)


def overlaps(a, b):
    return a.x < b.x + b.w and b.x < a.x + a.w and a.y < b.y + b.h and b.y < a.y + a.h


class LayoutTests(unittest.TestCase):

    def test_forms_fit_a_1366_x_768_screen(self):
        for m in MODELS:
            self.assertLessEqual(m.width, F.cm(34.5), m.name)
            self.assertLessEqual(m.height, F.cm(19.5), m.name)

    def test_controls_inside_form(self):
        for m in MODELS:
            for c in m.controls:
                with self.subTest(form=m.name, control=c.name):
                    self.assertGreaterEqual(c.x, 0)
                    self.assertGreaterEqual(c.y, 0)
                    self.assertGreater(c.w, 0)
                    self.assertGreater(c.h, 0)
                    self.assertLessEqual(c.x + c.w, m.width)
                    self.assertLessEqual(c.y + c.h, m.height)

    def test_no_overlapping_controls(self):
        for m in MODELS:
            solid = [c for c in m.controls if not c.decorative]
            for i, a in enumerate(solid):
                for b in solid[i + 1:]:
                    self.assertFalse(overlaps(a, b), f"{m.name}: {a.name} overlaps {b.name}")

    def test_control_names_unique(self):
        for m in MODELS:
            names = [c.name.lower() for c in m.controls]
            self.assertEqual(len(names), len(set(names)), m.name)

    def test_labels_are_never_empty(self):
        for m in MODELS:
            for c in m.controls:
                if c.kind in ("label", "icon"):
                    cap = c.props.get("Caption")
                    self.assertTrue(isinstance(cap, F.Sym) or (cap and cap.strip() != "" or cap == " "),
                                    f"{m.name}.{c.name}")


class WiringTests(unittest.TestCase):

    def test_every_event_has_a_procedure_and_vice_versa(self):
        for m in MODELS:
            code = "\n".join(m.code)
            procs = set(re.findall(r"^Private Sub (\w+)\(", code, re.M))
            expected = {f"Form_{e}" for e in m.form_events}
            expected |= {f"{c.name}_{e}" for c in m.controls for e in c.events}
            self.assertEqual(procs, expected, m.name)

    def test_event_procedure_signatures(self):
        sig = {"BeforeUpdate": "(Cancel As Integer)", "Unload": "(Cancel As Integer)", "Open": "(Cancel As Integer)",
               "BeforeInsert": "(Cancel As Integer)",
               "Error": "(DataErr As Integer, Response As Integer)",
               "KeyDown": "(KeyCode As Integer, Shift As Integer)", "DblClick": "(Cancel As Integer)"}
        for m in MODELS:
            code = "\n".join(m.code)
            for name, args in re.findall(r"^Private Sub (\w+)(\(.*\))", code, re.M):
                event = name.split("_")[-1]
                if name.startswith("Form_"):
                    expected = sig.get(event, "()")
                elif event in ("DblClick", "KeyDown"):          # control events with arguments
                    expected = sig[event]
                else:
                    expected = "()"                              # BeforeUpdate on controls unused
                self.assertEqual(args, expected, f"{m.name}.{name}")

    def test_navigation_targets(self):
        built = {m.name for m in MODELS}
        for item in F.NAV_ITEMS:
            if item.target and item.phase <= 5:
                self.assertIn(item.target, built, item.key)
            if item.target and item.target not in built:
                self.assertGreater(item.phase, 5, f"{item.key} has no screen yet")

    def test_tags_fit_and_are_parseable(self):
        for m in MODELS:
            self.assertLess(len(m.tag), 2048, m.name)
            for part in filter(None, m.tag.split("|")):
                self.assertIn("=", part, m.name)


class BoundFieldTests(unittest.TestCase):

    def test_bound_fields_exist(self):
        for s in F.DATA_SCREENS:
            names = {f.name for f in table(s.table).fields}
            m = next(x for x in MODELS if x.name == s.name)
            for c in m.controls:
                if c.source:
                    self.assertIn(c.source, names, f"{s.name}.{c.name}")
                    self.assertEqual(c.name, c.source, "bound controls are named after their field")

    def test_required_fields_are_on_screen(self):
        for s in F.DATA_SCREENS:
            if not s.allow_add and s.kind != "SINGLE":       # edits records made elsewhere (frmEmployeePay)
                continue
            on_screen = {c.source for c in next(x for x in MODELS if x.name == s.name).controls}
            for f in table(s.table).fields:
                if f.required and f.default is None and f.kind not in ("AUTO", "BOOL") \
                        and (s.table, f.name) not in AUTO_FILLED:
                    self.assertIn(f.name, on_screen, f"{s.name}: {f.name} is required")

    def test_code_references_existing_controls(self):
        """frm!Name used by modForms for each table must exist on the screen."""
        text = read("modForms")
        by_table = {s.table: next(x for x in MODELS if x.name == s.name) for s in F.DATA_SCREENS}
        rules = {"ValidateProduct": "Products", "ValidateExpense": "Expenses",
                 "ValidateSettings": "Settings", "UpdatePriceInfo": "Products"}
        for proc, tbl in rules.items():
            body = text.split(f"Private Function {proc}(")[-1] if f"Function {proc}(" in text \
                else text.split(f"Private Sub {proc}(")[-1]
            body = body.split("\nEnd ")[0]
            names = {c.name for c in by_table[tbl].controls}
            fields = {f.name for f in table(tbl).fields}
            for ref in set(re.findall(r"frm!(\w+)", body)):
                self.assertTrue(ref in names or ref in fields, f"{proc}: frm!{ref}")
        for tbl in ("Customers", "Suppliers"):
            body = text.split("Private Function ValidatePartner(")[1].split("\nEnd Function")[0]
            names = {c.name for c in by_table[tbl].controls}
            for ref in set(re.findall(r"frm!(\w+)", body)) - {"AllowCredit"}:
                self.assertIn(ref, names, f"ValidatePartner on {tbl}: frm!{ref}")


class SqlOnMirrorTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.db = AccessOnSqlite()
        cls.db.load_fixture()

    def run_sql(self, sql):
        return self.db.con.execute(to_sqlite(sql)).fetchall()

    def test_combo_and_list_row_sources(self):
        for m in MODELS:
            for c in m.controls:
                rows = c.props.get("RowSource", "")
                if rows.upper().startswith("SELECT"):
                    with self.subTest(form=m.name, control=c.name):
                        self.run_sql(rows)

    def test_list_templates(self):
        for s in F.DATA_SCREENS:
            if s.kind != "LIST":
                continue
            base = s.list_template()
            active = f"{s.active} = True" if s.active else "True"
            all_rows = self.run_sql(base.replace("{ACTIVE}", active).replace("{SEARCH}", "True"))
            search = " OR ".join(f"{f} Like '*TEST*'" for f in s.search)
            found = self.run_sql(base.replace("{ACTIVE}", "True").replace("{SEARCH}", search))
            self.assertLessEqual(len(found), len(all_rows) + 100)
            self.assertEqual(len(all_rows[0]) if all_rows else len(s.list_headers) + 1,
                             len(s.list_headers) + 1, s.name)
        products = next(s for s in F.DATA_SCREENS if s.table == "Products")
        sql = products.list_template().replace("{ACTIVE}", "True").replace(
            "{SEARCH}", " OR ".join(f"{f} Like '*TEST-P2*'" for f in products.search))
        self.assertEqual([r[1] for r in self.run_sql(sql)], ["TEST-P2"])

    def fill(self, kind, text="", num=-1, frm="1900-01-01 00:00:00", to="9999-12-31 00:00:00"):
        k = next(k for k in F.SEARCH_KINDS if k.key == kind)
        sql = (k.sql.replace("{LIKE}", f"'*{text}*'").replace("{NUM}", str(num))
               .replace("{FROM}", f"'{frm}'").replace("{TO}", f"'{to}'"))
        return self.run_sql(sql)

    def test_search_templates(self):
        for k in F.SEARCH_KINDS:
            rows = self.fill(k.key)
            if rows:
                self.assertEqual(len(rows[0]), len(k.widths), k.key)
        ids = self.db.ids
        self.assertEqual([r[0] for r in self.fill("PRODUCT", "TEST-P1")], [ids["P1"]])
        self.assertEqual([r[0] for r in self.fill("PRODUCT", "zzz", num=ids["P3"])], [ids["P3"]])
        self.assertEqual([r[0] for r in self.fill("CUSTOMER", "آجل")], [ids["C2"]])
        self.assertEqual([r[0] for r in self.fill("SALE", "TEST-INV-2")], [ids["INV2"]])
        self.assertEqual(len(self.fill("SALE", "TEST")), 3)
        # date range: only the last 7 days -> INV2 (5 days ago)
        start = self.db.value(Q.Day(7))
        self.assertEqual([r[0] for r in self.fill("SALE", frm=start)], [ids["INV2"]])
        self.assertEqual([r[0] for r in self.fill("PURCHASE", "S-100")], [ids["PUR1"]])

    def test_report_catalogue(self):
        queries = {q.name: q for q in Q.QUERIES}
        self.db.params.update({"CustomerID": self.db.ids["C2"], "SupplierID": self.db.ids["S1"],
                               "ProductID": self.db.ids["P1"]})
        for r in F.REPORTS:
            with self.subTest(r.key):
                self.assertIn(r.query, queries)
                q = queries[r.query]
                params = set(q.params) | set(re.findall(r"Q(?:Date|Long)\('(\w+)'\)", q.sql))
                self.assertEqual("PeriodStart" in params, "P" in r.needs, "period parameter")
                self.assertEqual("CustomerID" in params, "C" in r.needs)
                self.assertEqual("SupplierID" in params, "S" in r.needs)
                self.assertEqual("ProductID" in params, "R" in r.needs)
                cur = self.db.con.execute(f'SELECT * FROM "{r.query}"')
                columns = {d[0] for d in cur.description}
                if "D" in r.needs:
                    self.assertIn(r.date_column, columns)
                for letter, col in (("c", "CustomerID"), ("s", "SupplierID"), ("r", "ProductID")):
                    if letter in r.needs:
                        self.assertIn(col, columns, f"optional filter {col}")
        self.assertEqual(len({r.key for r in F.REPORTS}), len(F.REPORTS))
        self.assertGreaterEqual(len(F.REPORTS), 16)


# --------------------------------------------------------------------------
# Static checks on every VBA module of the project
# --------------------------------------------------------------------------
def all_modules():
    out = {}
    for path in glob.glob(os.path.join(SRC, "*.bas")):
        with open(path, encoding="utf-8") as fh:
            out[os.path.basename(path)[:-4]] = fh.read()
    return out


def procedures(text, public_only=False):
    pat = r"^(Public |Private )?(?:Sub|Function) (\w+)\("
    return {name for scope, name in re.findall(pat, text, re.M)
            if not public_only or scope != "Private "}


class ProjectStaticTests(unittest.TestCase):

    def test_no_local_variable_shadows_a_procedure(self):
        """VBA names are case-insensitive: Dim needs + Function Needs() breaks the call."""
        modules = all_modules()
        public = {p.lower() for t in modules.values() for p in procedures(t, public_only=True)}
        for name, text in modules.items():
            own = {p.lower() for p in procedures(text)}
            for line in logical_lines(text):
                for decl in re.findall(r"\b(?:Dim|ByVal|ByRef)\s+(\w+)", line):
                    self.assertNotIn(decl.lower(), own | public, f"{name}: {decl}")
                for decl in re.findall(r",\s*(?:ByVal\s+|ByRef\s+)?(\w+)(?:\(\))?\s+As\b", line):
                    self.assertNotIn(decl.lower(), own | public, f"{name}: {decl}")

    def test_public_names_unique_across_modules(self):
        seen = {}
        for name, text in all_modules().items():
            for p in procedures(text, public_only=True):
                self.assertNotIn(p.lower(), seen, f"{p} in {name} and {seen.get(p.lower())}")
                seen[p.lower()] = name

    def test_form_code_calls_existing_public_procedures(self):
        public = {p for t in all_modules().values() for p in procedures(t, public_only=True)}
        builtin = {"DoCmd", "End", "Cancel", "Response"}
        for m in MODELS:
            for line in m.code:
                s = line.strip()
                if s.startswith(("Private Sub", "End Sub")):
                    continue
                m2 = re.match(r"^(?:Cancel = Not |Response = )?(\w+)", s)
                name = m2.group(1)
                if name in builtin:
                    continue
                self.assertIn(name, public, f"{m.name}: {s}")

    def test_modules_call_existing_procedures(self):
        """Every Name(...) or statement call to a project-looking procedure exists."""
        modules = all_modules()
        known = set()
        for t in modules.values():
            known |= procedures(t)
            known |= set(re.findall(r"Declare (?:PtrSafe )?(?:Function|Sub) (\w+)", t))   # Windows API
        vba_and_access = set("""Nz DLookup DCount DMax IIf Format Format$ Left Left$ Mid Mid$ Len Trim Trim$
            Replace Split InStr UCase UCase$ CStr CLng CDbl CCur CDec CDate Fix Abs DateSerial DateValue
            DateAdd Date Now Year Month IsNull IsNumeric IsDate IsEmpty Array LBound UBound MsgBox
            InputBox Environ Environ$ String String$ ChrW Debug CurrentDb CreateForm CreateControl
            Forms TempVars Application DBEngine Err Erl CurrentProject RGB Val Dir Dir$ Lines
            Round Int Space StrComp InStrRev LCase LCase$ Asc AscW Chr Chr$ Hex Sgn TimeSerial
            Hour Minute Weekday DatePart DateDiff Time Timer Erase CVar CInt CBool IsArray Choose
            DoEvents IsMissing IsObject TypeName VarType CSng Sqr Exp Log Rnd Second Day Hex$
            StrConv LenB AscB ChrB MidB Filter Join InStrB DMin DSum DAvg Eval Reports
            CreateReport CreateReportControl CreateGroupLevel GetSetting SaveSetting FileDateTime FileLen
            CreateObject GetObject""".split())
        for name, text in modules.items():
            in_type = False
            for line in logical_lines(text):
                if re.match(r"^(Private |Public )?Type \w+", line):
                    in_type = True
                if line == "End Type":
                    in_type = False
                    continue
                if in_type or "Declare " in line or line.startswith(("'", "Attribute", "#")) or not line:
                    continue
                code = re.sub(r'"[^"]*"', '""', line)
                for call in re.findall(r"(?<![\w.!])([A-Z]\w+)\(", code):
                    if call in known or call in vba_and_access:
                        continue
                    if re.search(rf"\b(?:Dim|Private|Public|Const|ByVal|ByRef)\s+{call}\b", text):
                        continue   # array/variable indexing
                    if re.search(rf"\b{call}\s+As\b", text):
                        continue
                    self.fail(f"{name}: unknown procedure {call}( in: {line}")

    def test_procedure_size_limit(self):
        """Access refuses procedures larger than 64 KB."""
        for name, text in all_modules().items():
            for proc in re.split(r"\n(?=(?:Public |Private )?(?:Sub|Function) )", text):
                self.assertLess(len(proc.encode("cp1256")), 60000, f"{name}: {proc[:60]}")


class GeneratedFormsModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modBuildForms"
    vba = gen_forms.build_forms_vba()

    def test_every_form_has_a_builder(self):
        for m in MODELS:
            self.assertIn(f"Private Sub BuildForm_{m.name}()", self.vba)
            self.assertIn(f"    BuildForm_{m.name}\n", self.vba)


class GeneratedAppDataModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modAppData"
    vba = gen_forms.build_appdata_vba()

    def test_search_sql_round_trips(self):
        for k in F.SEARCH_KINDS:
            body = self.vba.split(f'Case "{k.key}"\n')[1].split("        Case")[0]
            parts = re.findall(r'^\s+s = (?:s & )?"(.*)"$', body, re.M)
            self.assertEqual("".join(p.replace('""', '"') for p in parts), k.sql, k.key)


def _static(name):
    class T(VbaModuleChecks, unittest.TestCase):
        module_name = name
        vba = read(name)
    T.__name__ = T.__qualname__ = f"Static_{name}"
    return T


Static_modCommon = _static("modCommon")
Static_modStartup = _static("modStartup")
Static_modForms = _static("modForms")
Static_modScreens = _static("modScreens")



class FormRecordAccessTests(unittest.TestCase):

    def test_no_current_record_reads_through_form_recordset(self):
        # On a new record (or an empty table) frm.Recordset has no current row:
        # Access raises 3021 "No current record". Read and write frm("Field") instead.
        bad = re.compile(r"\.Recordset\.Fields\([^)]*\)\.Value|\.Recordset![A-Za-z]", re.I)
        for path in glob.glob(os.path.join(ROOT, "src", "vba", "*.bas")):
            with open(path, encoding="utf-8") as fh:
                lines = fh.readlines()
            for n, line in enumerate(lines, 1):
                self.assertIsNone(bad.search(line), f"{os.path.basename(path)}:{n}: {line.strip()}")


class WindowFitTests(unittest.TestCase):
    """Replays modForms.FitControls for several window sizes on every screen that has a fit spec."""

    @staticmethod
    def placed(m, dw, dh):
        out = []
        for c in m.controls:
            mx, mw, my, mh = m.fit.get(c.name, (0, 0, 0, 0))
            out.append(F.Control(c.kind, c.name, c.x + dw * mx // 1000, c.y + dh * my // 1000,
                                 c.w + dw * mw // 1000, c.h + dh * mh // 1000, decorative=c.decorative))
        return out

    def test_big_screens_fit_a_laptop_and_have_a_fit_spec(self):
        for m in MODELS:
            if not m.popup and m.height > F.cm(2):
                self.assertTrue(m.fit, m.name)
                self.assertLessEqual(m.height, F.cm(17.0), m.name)

    def test_no_overlap_and_inside_window_at_any_size(self):
        for m in MODELS:
            if not m.fit:
                continue
            min_dh = gen_forms.fit_min_dh(m)
            self.assertLessEqual(min_dh, 0, m.name)
            for dw in (0, F.cm(5), F.cm(20)):
                for dh in (min_dh, 0, F.cm(4), F.cm(15)):
                    ctls = self.placed(m, dw, dh)
                    for c in ctls:
                        with self.subTest(form=m.name, control=c.name, dw=dw, dh=dh):
                            self.assertGreaterEqual(c.h, 1)
                            self.assertLessEqual(c.x + c.w, m.width + dw)
                            self.assertLessEqual(c.y + c.h, m.height + dh)
                    solid = [c for c in ctls if not c.decorative]
                    for i, a in enumerate(solid):
                        for b in solid[i + 1:]:
                            self.assertFalse(overlaps(a, b), f"{m.name} dw={dw} dh={dh}: {a.name} / {b.name}")

    def test_spec_names_exist_and_line_lengths(self):
        for m in MODELS:
            names = {c.name for c in m.controls}
            self.assertTrue(set(m.fit) <= names, m.name)
            for line in gen_forms.fit_code_lines(m) if m.fit else []:
                self.assertLess(len(line), 1000)


class ControlPropertyTests(unittest.TestCase):
    """BuildForms sets only properties the control type has: Access stops the whole screen
    with error 438 otherwise (BackColor on the locked IsSystem check box of frmAccounts)."""

    MISSING = {"AddCheck": {"BackColor", "ForeColor", "FontSize", "FontBold", "Caption", "TextAlign",
                            "Format", "InputMask", "RowSource"},
               "AddRect": {"FontSize", "Caption", "Locked", "ControlSource"},
               "AddList": {"Caption", "Format", "InputMask"}}

    def test_no_missing_properties(self):
        with open(os.path.join(ROOT, "src", "vba", "modBuildForms.bas"), encoding="utf-8") as fh:
            lines = fh.read().splitlines()
        kind = None
        for no, line in enumerate(lines, 1):
            m = re.match(r'\s+Set c = (Add\w+)\(', line)
            if m:
                kind = m.group(1)
                continue
            m = re.match(r'\s+(?:c\.(\w+) =|SetCtlProp c, "(\w+)")', line)
            if m and kind in self.MISSING:
                self.assertNotIn(m.group(1) or m.group(2), self.MISSING[kind], f"line {no}: {line.strip()}")



class SharedTableScreenTests(unittest.TestCase):
    """modForms runs table handlers (Case "<table>" on the TABLE tag). When two screens edit
    one table, a handler written for one of them must check the screen name: UserCurrent of
    frmUsers stopped frmEmployeePay (error 2465, lblPasswordState)."""

    def test_handlers_of_shared_tables_check_the_screen(self):
        import collections
        screens = collections.defaultdict(list)
        for m in MODELS:
            found = re.search(r"\bTABLE=(\w+)", m.tag or "")
            if found:
                screens[found.group(1)].append(m.name)
        shared = {t for t, names in screens.items() if len(names) > 1}
        self.assertIn("Employees", shared)
        with open(os.path.join(ROOT, "src", "vba", "modForms.bas"), encoding="utf-8") as fh:
            text = fh.read()
        for t in shared:
            for case in re.finditer(rf'^        Case [^\n]*"{t}"[^\n]*\n(.*?)(?=^        Case |^    End Select)',
                                    text, re.M | re.S):
                self.assertIn("frm.Name = ", case.group(0), f"{t}: {case.group(0).strip()}")



class ScreensIndexTests(unittest.TestCase):

    def test_screens_index_is_up_to_date(self):
        with open(os.path.join(ROOT, "docs", "dev", "Screens-Index.md"), encoding="utf-8") as fh:
            self.assertEqual(fh.read().rstrip("\n"), gen_forms.build_screens_md().rstrip("\n"),
                             "run: python3 tools/generate.py")


if __name__ == "__main__":
    unittest.main()
