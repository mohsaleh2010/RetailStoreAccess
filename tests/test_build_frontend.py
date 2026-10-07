"""dist/tools/BuildFrontEnd.vbs builds the whole front-end in Access in one step:
every procedure it runs exists as a public procedure of the imported modules, the
build steps run in the installation order, and the script is plain ASCII (WSH)."""
import os
import re
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
SCRIPT = os.path.join(ROOT, "dist", "tools", "BuildFrontEnd.vbs")


def public_procs():
    found = set()
    folder = os.path.join(ROOT, "dist", "vba")
    for name in os.listdir(folder):
        with open(os.path.join(folder, name), encoding="cp1256") as fh:
            found |= set(re.findall(r"^Public (?:Function|Sub) (\w+)", fh.read(), re.M))
    return found


class BuildFrontEndTests(unittest.TestCase):

    def setUp(self):
        with open(SCRIPT, "rb") as fh:
            self.raw = fh.read()
        self.text = self.raw.decode("ascii")

    def test_runs_existing_procedures_in_order(self):
        steps = re.findall(r'RunStep "(\w+)", (True|False)', self.text)
        names = [s for s, _ in steps]
        self.assertEqual([s for s, req in steps if req == "True"],
                         ["BuildSchema", "BuildRelationships", "BuildQueries", "BuildForms", "BuildReports"])
        self.assertLessEqual(set(names), public_procs())
        self.assertLess(names.index("BuildReports"), names.index("LoadDemoData"))
        self.assertLess(names.index("LoadDemoData"), names.index("RunAllTests"))

    def test_imports_every_module_then_compiles(self):
        self.assertIn('fso.BuildPath(fso.GetParentFolderName(here), "vba")', self.text)
        self.assertIn('LCase(fso.GetExtensionName(f.Name)) = "bas"', self.text)
        self.assertEqual(self.text.count('Compile "'), 2)
        self.assertIn("app.IsCompiled", self.text)

    def test_private_procedures_stay_in_their_module(self):
        # A Private Sub/Function called from another module: "Sub or Function not defined"
        # when Access compiles (BoxIsActive of modCash, called by modBank and modAssets).
        folder = os.path.join(ROOT, "dist", "vba")
        code = {}
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding="cp1256") as fh:
                code[name] = "\n".join(re.sub(r'"[^"]*"', '""', l).split("'", 1)[0] for l in fh.read().splitlines())
        procs = {n: set(re.findall(r"^(?:Public |Private )?(?:Function|Sub) (\w+)", t, re.M)) for n, t in code.items()}
        public = set().union(*(re.findall(r"^(?:Public )?(?:Function|Sub) (\w+)", t, re.M) for t in code.values()))
        for owner, text in code.items():
            for proc in re.findall(r"^Private (?:Function|Sub) (\w+)", text, re.M):
                if proc in public:
                    continue
                for other, body in code.items():
                    if other == owner or proc in procs[other] or re.search(rf"\b{proc}\s+As\b", body):
                        continue
                    self.assertIsNone(re.search(rf"(?<![.!\w]){proc}\s*\(", body), f"{owner} {proc} used in {other}")

    def test_no_date_conversion_of_a_zero_default(self):
        # VBA evaluates both sides of And / Or: IsDate(x) And DateValue(Nz(x, 0)) fails with
        # Type mismatch when x is Null (ReopenPeriod, TestJournal).
        folder = os.path.join(ROOT, "dist", "vba")
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding="cp1256") as fh:
                for no, line in enumerate(fh, 1):
                    self.assertNotRegex(line.split("'", 1)[0], r"DateValue\(Nz\([^()]*,\s*0\)\)", f"{name}:{no}")

    def test_iif_branches_take_no_null(self):
        # IIf evaluates both branches: IIf(IsDate(x), SqlDate(x), "Null") fails with "Invalid use
        # of Null" when x is Null (SetClosedThrough, TestJournal). Nz(x, ...) is fine.
        typed = r"(?:SqlDate|Format\$|CLng|CDate|CCur|CDbl|CInt|DateValue|Year|Month|Day)"
        folder = os.path.join(ROOT, "dist", "vba")
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding="cp1256") as fh:
                for no, line in enumerate(fh, 1):
                    code = line.split("'", 1)[0]
                    for m in re.finditer(r"IIf\((?:Not )?(?:IsNull|IsDate|IsEmpty)\(([^()]+(?:\(\))?)\)", code):
                        var = re.escape(m.group(1))
                        rest = code[m.end():]
                        self.assertNotRegex(rest, typed + r"\(" + var + r"\s*[,)]", f"{name}:{no}: {line.strip()}")

    def test_no_blank_column_names(self):
        # AS [ ] - Access: error 3126 "Invalid bracketing of name ' '" (the list of frmCostCenters).
        import sys
        sys.path.insert(0, os.path.join(ROOT, "tools"))
        import forms as F
        import queries as Q
        texts = [m.tag for m in F.all_forms()] + [q.sql for q in Q.QUERIES]
        folder = os.path.join(ROOT, "dist", "vba")
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding="cp1256") as fh:
                texts.append(fh.read())
        for text in texts:
            found = re.search(r"\bAS \[\s*\]", text)
            self.assertIsNone(found, text[max(0, found.start() - 80):found.end()] if found else "")

    def test_no_chained_currentdb_collections(self):
        # CurrentDb.TableDefs(...).Fields / For Each x In CurrentDb.TableDefs: the temporary
        # database object is released while it is read - error 3420 on a fresh front-end.
        bad = re.compile(r"CurrentDb(?:\(\))?\.(?:TableDefs|QueryDefs|Relations|Containers)\b")
        folder = os.path.join(ROOT, "dist", "vba")
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding="cp1256") as fh:
                for no, line in enumerate(fh, 1):
                    code = line.split("'", 1)[0]
                    self.assertIsNone(bad.search(code), f"{name}:{no}: {line.strip()}")

    def test_totalled_reports_have_no_subqueries(self):
        # A report with totals wraps its record source in a GROUP BY; Access then refuses a
        # subquery in the column list of it or of the queries it reads (error 3612, rptCashDaily).
        # A subquery in WHERE (SlowMovingProductsQuery) is accepted.
        import sys
        sys.path.insert(0, os.path.join(ROOT, "tools"))
        import forms as F
        import queries as Q
        import reports_catalog as R
        by_name = {q.name: q for q in Q.QUERIES}
        entries = {r.key: r for r in F.REPORTS}

        def chain(name, seen):
            if name in seen or name not in by_name:
                return seen
            seen.add(name)
            for other in re.findall(r"\b(?:qry\w+|\w+Query)\b", by_name[name].sql):
                chain(other, seen)
            return seen

        for spec in R.LIST_SPECS:
            if not (any(c.total or c.running for c in spec.cols) or spec.summary):
                continue
            for name in chain(entries[spec.key].query, set()):
                columns = re.split(r"\bFROM\b", by_name[name].sql, maxsplit=1)[0]
                self.assertNotRegex(columns, r"\(\s*SELECT", f"{spec.key}: {name}")

    def test_ascii_crlf(self):
        self.assertNotIn(b"\n", self.raw.replace(b"\r\n", b""))


if __name__ == "__main__":
    unittest.main()
