"""Phase 4 checks: query results on the fixture (SQLite mirror), Access-dialect
lint for every saved query, and the generated VBA module.

Run:  python3 -m unittest discover -s tests -v
"""

import re
import unittest

from access_sqlite import AccessOnSqlite
from helpers import VbaModuleChecks
import gen_queries
import queries as Q
from schema import TABLES

SPEC_QUERIES = ["DailySalesQuery", "MonthlySalesQuery", "StockBalanceQuery", "LowStockQuery",
                "CustomerBalanceQuery", "SupplierBalanceQuery", "ProfitQuery",
                "BestSellingProductsQuery", "SlowMovingProductsQuery"]

REPORT_QUERIES = ["DailySalesQuery", "MonthlySalesQuery", "SalesByPeriodQuery",
                  "SalesByProductQuery", "BestSellingProductsQuery",
                  "LeastSellingProductsQuery", "PurchasesQuery", "StockBalanceQuery",
                  "LowStockQuery", "ProductMovementQuery", "CustomerStatementQuery",
                  "SupplierStatementQuery", "ExpensesQuery", "ProfitQuery",
                  "SlowMovingProductsQuery", "StockByCategoryQuery"]


# --------------------------------------------------------------------------
# helpers for the dialect lint
# --------------------------------------------------------------------------
def strip_strings(sql):
    return re.sub(r"'[^']*'", "''", sql)


def top_level_split(text, sep_regex):
    """Split on a regex only at parenthesis depth 0."""
    parts, depth, last = [], 0, 0
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0:
            m = re.compile(sep_regex, re.I).match(text, i)
            if m:
                parts.append(text[last:i])
                last = i = m.end()
                continue
        i += 1
    parts.append(text[last:])
    return parts


def union_branches(sql):
    return top_level_split(strip_strings(sql), r"\s+UNION ALL\s+")


def select_items(branch):
    m = re.search(r"SELECT\s+(.*?)\s+FROM\s", branch, re.S | re.I)
    if not m:
        m = re.search(r"SELECT\s+(.*)$", branch, re.S | re.I)
    return [x.strip() for x in top_level_split(m.group(1), r",")]


def from_clause(branch):
    m = re.search(r"\sFROM\s+(.*?)(?:\s+WHERE\s|\s+GROUP BY\s|\s+ORDER BY\s|\s+HAVING\s|$)",
                  branch, re.S | re.I)
    return m.group(1) if m else ""


class AccessDialectTests(unittest.TestCase):

    def test_names_unique_and_spec_queries_exist(self):
        names = [q.name for q in Q.QUERIES]
        self.assertEqual(len(names), len(set(names)))
        table_names = {t.name.lower() for t in TABLES}
        for n in names:
            self.assertNotIn(n.lower(), table_names)
            self.assertLessEqual(len(n), 64)
        for n in SPEC_QUERIES + REPORT_QUERIES + ["IntegrityCheckQuery", "VatSummaryQuery"]:
            self.assertIn(n, names)

    def test_forbidden_constructs(self):
        forbidden = [r"\bCASE\b", r"\bCOALESCE\b", r"\bLIMIT\b", r"\bTOP\b", r"\|\|",
                     r"COUNT\s*\(\s*DISTINCT", r"\bCROSS JOIN\b", r'"', r"%", r"\bIFNULL\b",
                     r"\[TempVars\]", r"\[Forms\]", r"#"]
        for q in Q.QUERIES:
            body = strip_strings(q.sql)
            for pat in forbidden:
                self.assertIsNone(re.search(pat, body, re.I), f"{q.name}: {pat}")

    def test_joins_are_explicit_and_nested(self):
        for q in Q.QUERIES:
            for branch in union_branches(q.sql):
                body = " ".join(branch.split())
                self.assertIsNone(re.search(r"(?<!INNER)(?<!LEFT)(?<!RIGHT) JOIN\b", body),
                                  f"{q.name}: bare JOIN")
                frm = from_clause(" " + body)
                joins = len(re.findall(r"\bJOIN\b", frm))
                leading = len(frm) - len(frm.lstrip("("))
                self.assertGreaterEqual(leading, max(0, joins - 1),
                                        f"{q.name}: {joins} joins need nested parentheses")

    def test_no_constant_in_outer_join_condition(self):
        for q in Q.QUERIES:
            body = " ".join(strip_strings(q.sql).split())
            for m in re.finditer(r"LEFT JOIN .*? ON (\(.*?\)|[^()]*?)(?= WHERE| GROUP| ORDER| LEFT| INNER|\)|$)",
                                 body):
                self.assertNotIn("''", m.group(1), f"{q.name}: constant in LEFT JOIN ON")

    def test_no_circular_alias(self):
        for q in Q.QUERIES:
            for branch in union_branches(q.sql):
                for item in select_items(branch):
                    m = re.match(r"^(.*)\s+AS\s+(\w+)$", item, re.S | re.I)
                    if not m:
                        continue
                    expr, alias = m.group(1), m.group(2)
                    self.assertIsNone(re.search(rf"(?<![\w.]){alias}\b|\.{alias}\b", expr),
                                      f"{q.name}: alias {alias} used in its own expression")

    def test_union_branches_have_same_width(self):
        for q in Q.QUERIES:
            widths = [len(select_items(b)) for b in union_branches(q.sql)]
            self.assertEqual(len(set(widths)), 1, f"{q.name}: {widths}")

    def test_numeric_nz_is_wrapped_in_ccur(self):
        for q in Q.QUERIES:
            for m in re.finditer(r"(\w*)\(Nz\(", q.sql):
                if "DateDiff" in q.sql[max(0, m.start() - 10):m.start() + 5]:
                    continue
                self.assertEqual(m.group(1), "CCur", f"{q.name}: Nz not wrapped in CCur")
            for m in re.finditer(r"Nz\(([^()]*(?:\([^()]*\))?[^()]*)\)", q.sql):
                self.assertEqual(len(top_level_split(m.group(1), r",")), 2,
                                 f"{q.name}: Nz needs 2 arguments")

    def test_queries_only_use_earlier_queries(self):
        seen = set()
        names = [q.name for q in Q.QUERIES]
        for q in Q.QUERIES:
            for other in names:
                if other != q.name and re.search(rf"\b{other}\b", q.sql):
                    self.assertIn(other, seen, f"{q.name} uses {other} before it is defined")
            seen.add(q.name)

    def test_referenced_columns_exist_in_tables(self):
        cols = {t.name: {f.name for f in t.fields} for t in TABLES}
        for q in Q.QUERIES:
            for branch in union_branches(q.sql):
                alias_to_table = {a: t for t, a in re.findall(r"\b(\w+) AS (\w+)\b", branch)
                                  if t in cols}
                for a, c in re.findall(r"\b(\w+)\.(\w+)\b", branch):
                    if a in alias_to_table:
                        self.assertIn(c, cols[alias_to_table[a]], f"{q.name}: {a}.{c}")

    def test_parameterised_queries_declare_their_params(self):
        for q in Q.QUERIES:
            used = set(re.findall(r"Q(?:Date|Long)\('(\w+)'\)", q.sql))
            self.assertTrue(used <= set(q.params) or not used or set(q.params) >= used,
                            f"{q.name}: {used} vs {q.params}")


class QueryResultTests(unittest.TestCase):
    """Every check from tools/queries.py evaluated on the SQLite mirror."""

    @classmethod
    def setUpClass(cls):
        cls.db = AccessOnSqlite()
        cls.db.load_fixture()

    def test_every_query_runs(self):
        self.db.params.update({"CustomerID": self.db.ids["C2"], "SupplierID": self.db.ids["S1"],
                               "ProductID": self.db.ids["P2"]})
        for q in Q.QUERIES:
            with self.subTest(q.name):
                self.db.con.execute(f'SELECT * FROM "{q.name}"').fetchall()

    def test_checks(self):
        for check in Q.CHECKS:
            with self.subTest(check.label):
                actual, expected = self.db.run_check(check)
                self.assertIsNotNone(actual, check.label)
                self.assertAlmostEqual(float(actual), float(expected), places=2, msg=check.label)


class IntegrityCorruptionTests(unittest.TestCase):

    def test_corruptions_are_detected(self):
        db = AccessOnSqlite()
        db.load_fixture()
        for update, check in Q.CORRUPTIONS:
            db.con.execute(db.sql(update))
            actual, expected = db.run_check(check)
            self.assertAlmostEqual(float(actual), float(expected), places=2, msg=check.label)


class PeriodBoundaryTests(unittest.TestCase):
    """The period includes the first and the last day completely."""

    def test_period_edges(self):
        db = AccessOnSqlite()
        db.load_fixture()
        db.params["PeriodStart"] = db.value(Q.Day(5))       # day of the credit sale, 00:00
        db.params["PeriodEnd"] = db.value(Q.Day(4))         # exclusive -> only that day
        self.assertEqual(db.scalar("SELECT COUNT(*) FROM SalesByPeriodQuery"), 1)
        self.assertEqual(db.scalar("SELECT Sum(GrossAmount) FROM SalesByPeriodQuery"), 460)
        db.params["PeriodEnd"] = db.value(Q.Day(5))         # empty period
        self.assertEqual(db.scalar("SELECT COUNT(*) FROM SalesByPeriodQuery"), 0)
        self.assertEqual(db.scalar("SELECT NetProfit FROM ProfitQuery"), 0)


class GeneratedQueriesModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modBuildQueries"
    vba = gen_queries.build_queries_vba()

    def test_every_query_is_created_in_order(self):
        calls = re.findall(r"^    Q_(\w+)$", self.vba, re.M)
        self.assertEqual(calls, [q.name for q in Q.QUERIES])
        for q in Q.QUERIES:
            self.assertIn(f"Private Sub Q_{q.name}()", self.vba)

    def test_query_sql_round_trips(self):
        """Rebuild each query's SQL from the generated VBA lines and compare."""
        for q in Q.QUERIES:
            body = self.vba.split(f"Private Sub Q_{q.name}()")[1].split("End Sub")[0]
            parts = re.findall(r'^    s = (?:s & )?"(.*)" & vbCrLf$', body, re.M)
            rebuilt = "\n".join(p.replace('""', '"') for p in parts)
            expected = "\n".join(l.rstrip() for l in q.sql.strip().splitlines())
            self.assertEqual(rebuilt, expected.replace("\u2212", "-"), q.name)

    def test_fixture_and_checks_complete(self):
        self.assertEqual(self.vba.count("\n    Ins "), len(Q.FIXTURE))
        self.assertEqual(self.vba.count("\n    Chk "), len(Q.CHECKS) + len(Q.CORRUPTIONS))
        self.assertNotIn("{ref:", self.vba)

    def test_no_name_keyword_as_variable(self):
        self.assertIsNone(re.search(r"\b(?:Dim|For Each)\s+name\b", self.vba, re.I))


class StaticModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modQueryParams"
    with open(gen_queries.__file__.replace("tools/gen_queries.py", "src/vba/modQueryParams.bas"),
              encoding="utf-8") as _fh:
        vba = _fh.read()

    def test_public_api(self):
        for sig in ("Public Sub SetPeriod(", "Public Sub SetQueryParam(",
                    "Public Sub ClearQueryParams(", "Public Function QDate(",
                    "Public Function QLong("):
            self.assertIn(sig, self.vba)


if __name__ == "__main__":
    unittest.main()
