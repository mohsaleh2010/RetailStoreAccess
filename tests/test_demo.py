"""Phase 11 checks: demo data, the test runner, and project-wide call checks.

  * the demo plan covers the specification (20 products, 5 categories, 10
    customers, 5 suppliers, 3 employees, 10 sales, 5 purchases, expenses);
  * replayed on the SQLite mirror with the posting rules (tools/sim.py) it
    leaves no integrity issue, no negative stock, and every report of the
    report centre has data; RemoveDemoData's statements empty it again;
  * modDemoData verifies exactly the values of the replay;
  * every call to a procedure of the project passes an acceptable number of
    arguments (a wrong count only shows up when Access compiles the module).

Run:  python3 -m unittest discover -s tests -v
"""

import glob
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks, logical_lines
import demo_data as DD
import forms as F
import gen_demo
from access_sqlite import day

SPEC_MINIMUM = {"Products": 20, "Categories": 5, "Customers": 10, "Suppliers": 5, "Employees": 3,
                "SalesInvoices": 10, "PurchaseInvoices": 5, "Expenses": 3}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


class PlanTests(unittest.TestCase):

    def test_covers_the_specification(self):
        counts = DD.demo_counts()
        for key, minimum in SPEC_MINIMUM.items():
            self.assertGreaterEqual(counts[key], minimum, key)

    def test_plan_is_chronological(self):
        days = [(-p["day"], p["hour"] if p["day"] else 24) for p in DD.PLAN]
        self.assertEqual(days, sorted(days), "documents must be posted in date order")

    def test_unique_identifiers(self):
        barcodes = [p.barcode for p in DD.PRODUCTS]
        self.assertEqual(len(barcodes), len(set(barcodes)))
        for b in barcodes:
            self.assertEqual(len(b), 13)
            body = [int(d) for d in b[:12]]
            self.assertEqual((10 - sum(d * (3 if i % 2 else 1) for i, d in enumerate(body)) % 10) % 10, int(b[12]))
        for vat in [s.vat for s in DD.SUPPLIERS] + [c.vat for c in DD.CUSTOMERS if c.vat] + [DD.STORE["VATNumber"]]:
            self.assertRegex(vat, r"^3\d{13}3$")

    def test_variety(self):
        ops = [p["op"] for p in DD.PLAN]
        for op in ("sales_return", "purchase_return", "customer_payment", "supplier_payment", "stock_out",
                   "stock_count"):
            self.assertIn(op, ops)
        sales = [p for p in DD.PLAN if p["op"] == "sale"]
        self.assertTrue(any(p.get("credit") for p in sales) and any(not p.get("credit") for p in sales))
        self.assertTrue(any(p.get("discount") for p in sales))
        self.assertTrue(any(p.get("user") for p in sales), "some sales by the demo cashiers")
        self.assertTrue(any(DD.CUSTOMERS[p["customer"] - 1].vat for p in sales if p.get("customer")),
                        "a B2B sale (standard tax invoice, A4)")
        self.assertEqual(DD.PLAN[-1]["day"], 0, "a sale today for the dashboard")


class ReplayTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        cls.c = cls.r.db.con

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def test_no_integrity_issue(self):
        # the mirror stores money as floating point: equal to the halala is equal
        rows = self.q("SELECT IssueCode, ExpectedValue, ActualValue FROM IntegrityCheckQuery")
        real = [r for r in rows if abs((r[1] or 0) - (r[2] or 0)) > 0.005]
        self.assertEqual(real, [])

    def test_no_negative_stock_and_ledger_matches(self):
        self.assertEqual(min(self.r.stock.values()), min(v for v in self.r.stock.values()))
        self.assertGreaterEqual(min(self.r.stock.values()), 0)
        self.assertEqual(self.q("SELECT COUNT(*) FROM StockBalanceQuery WHERE abs(QuantityMismatch) > 0.0001")[0][0], 0)

    def test_alerts_have_something_to_show(self):
        self.assertEqual(self.q("SELECT COUNT(*) FROM LowStockQuery")[0][0], 4)
        self.assertEqual(self.q("SELECT COUNT(*) FROM SlowMovingProductsQuery")[0][0], 1)

    def test_every_report_has_data(self):
        db = self.r.db
        db.set_period(31)
        db.params.update({"CustomerID": self.r.ids["cust"][1], "SupplierID": self.r.ids["sup"][1],
                          "ProductID": self.r.ids["prod"][1]})
        for rep in F.REPORTS:
            rows = self.q(f'SELECT COUNT(*) FROM "{rep.query}"')[0][0]
            with self.subTest(rep.key):
                if rep.key == "INTEGRITY":
                    continue                          # checked with a tolerance above
                if rep.key == "BUDGET_VS_ACTUAL":
                    continue                          # needs a budget, entered by the user (not demo data)
                if rep.key == "AUDIT_TRAIL":
                    continue                          # written by the screens in Access, not by the replay
                self.assertGreater(rows, 0, rep.query)

    def test_dashboard_has_today(self):
        db = self.r.db
        db.params.update({"DashDay": day(0), "DashEnd": day(-1), "DashMonth": day(30)})
        today, invoices = self.q('SELECT TodaySales, TodayInvoices FROM "DashboardQuery"')[0]
        self.assertGreater(today, 0)
        self.assertEqual(invoices, 1)

    def test_generated_module_checks_the_replay(self):
        vba = gen_demo.build_demo_vba()
        self.assertEqual(vba.count("Check PostSaleFromCart("), DD.demo_counts()["SalesInvoices"])
        self.assertEqual(vba.count("Check PostPurchaseFromCart("), DD.demo_counts()["PurchaseInvoices"])
        self.assertIn(f"CCur({gen_demo.money(self.r.sales_total)})", vba)
        for n in self.r.stock:
            self.assertIn(f'ProductValue("{DD.PRODUCTS[n - 1].barcode}", "CurrentQuantity") = '
                          f'{gen_demo.num(self.r.stock[n])}', vba)
        self.assertNotIn("INSERT INTO SalesInvoices", vba, "documents only through the posting functions")
        self.assertNotIn("INSERT INTO PurchaseInvoices", vba)
        self.assertNotIn("INSERT INTO InventoryTransactions", vba)


class RemoveTests(unittest.TestCase):

    def test_remove_statements_empty_the_demo(self):
        r = DD.simulate()
        c = r.db.con
        c.execute("PRAGMA foreign_keys = ON")
        for sql in gen_demo.remove_sql():
            c.execute(sql)                      # RESTRICT relationships: fails if the order is wrong
        for t in DD.DOCUMENT_TABLES:
            self.assertEqual(c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0], 0, t)
        for t, field in (("Products", "Notes"), ("Customers", "Notes"), ("Suppliers", "Notes"),
                         ("Employees", "Notes"), ("Categories", "Description")):
            self.assertEqual(c.execute(f"SELECT COUNT(*) FROM {t} WHERE {field} = 'DEMO'").fetchone()[0], 0, t)
        self.assertEqual(c.execute("SELECT COUNT(*) FROM Customers WHERE CustomerID = 1").fetchone()[0], 1)
        self.assertEqual(c.execute("SELECT COUNT(*) FROM Employees WHERE EmployeeID = 1").fetchone()[0], 1)
        self.assertEqual(c.execute("SELECT NextValue FROM Sequences WHERE SequenceName = 'SALES_INVOICE'")
                         .fetchone()[0], 1)
        self.assertEqual(c.execute("SELECT COUNT(*) FROM IntegrityCheckQuery").fetchone()[0], 0)


# --------------------------------------------------------------------------
# Project-wide: argument counts of calls to the project's own procedures
# --------------------------------------------------------------------------
SIG = re.compile(r"^(?:Public |Private )?(Sub|Function) (\w+)\((.*?)\)(?: As \w+)?\s*$", re.M)


def signatures():
    """name -> (required, total) for every procedure of every module (private ones per module)."""
    public, private = {}, {}
    for path in glob.glob(os.path.join(ROOT, "src", "vba", "*.bas")):
        module = os.path.basename(path)[:-4]
        with open(path, encoding="utf-8") as fh:
            text = "\n".join(logical_lines(fh.read()))
        for kind, name, params in SIG.findall(text):
            ps = [p.strip() for p in params.split(",") if p.strip()]
            req = sum(1 for p in ps if not p.startswith("Optional") and not p.startswith("ParamArray"))
            sig = (req, len(ps))
            line = re.search(rf"^(Public |Private )?(?:Sub|Function) {name}\(", text, re.M)
            if line and line.group(1) == "Private ":
                private.setdefault(module, {})[name.lower()] = sig
            else:
                public[name.lower()] = sig
    return public, private


def split_args(s):
    out, depth, cur, in_str = [], 0, "", False
    for ch in s:
        if ch == '"':
            in_str = not in_str
        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                out.append(cur)
                cur = ""
                continue
        cur += ch
    if cur.strip() or out:
        out.append(cur)
    return out


def calls_in(line, names):
    """(name, argument count) of every call to a known procedure in one logical line."""
    found = []
    code = re.sub(r'"[^"]*"', '""', line)
    for m in re.finditer(r"(?<![\w.!$])([A-Za-z]\w*)\(", code):
        name = m.group(1)
        if name.lower() not in names:
            continue
        depth, j = 1, m.end()
        while j < len(code) and depth:
            depth += {"(": 1, ")": -1}.get(code[j], 0)
            j += 1
        found.append((name, len(split_args(code[m.end():j - 1]))))
    st = re.match(r"^(?:Call\s+)?([A-Za-z]\w*)(?:\s+(.*))?$", code.strip())
    if st and st.group(1).lower() in names and not code.strip().lower().startswith(("dim ", "set ")) \
            and "=" not in (st.group(2) or "").split(",")[0][:1] and not re.match(r"^\w+\s*=", code.strip()) \
            and not re.match(r"^\w+\(", code.strip()):
        found.append((st.group(1), len(split_args(st.group(2))) if st.group(2) else 0))
    return found


class ArgumentCountTests(unittest.TestCase):

    def test_calls_match_signatures(self):
        public, private = signatures()
        problems = []
        for path in glob.glob(os.path.join(ROOT, "src", "vba", "*.bas")):
            module = os.path.basename(path)[:-4]
            known = dict(public)
            known.update(private.get(module, {}))
            with open(path, encoding="utf-8") as fh:
                text = fh.read()
            locals_ = {d.lower() for d in re.findall(r"\b(?:Dim|Private|Public|Const)\s+(\w+)", text)}
            names = {k: v for k, v in known.items() if k not in locals_}
            for line in logical_lines(text):
                s = line.strip()
                if not s or s.startswith(("'", "Attribute", "#")) or re.match(r"^(Public |Private )?(Sub|Function|"
                                                                            r"Declare|Type|Const|Dim) ", s):
                    continue
                for name, n in calls_in(s, names):
                    req, total = names[name.lower()]
                    if not req <= n <= total:
                        problems.append(f"{module}: {name} called with {n} argument(s), expects {req}-{total}: {s}")
        for m in F.all_forms():
            for line in m.code:
                s = line.strip()
                if s.startswith(("Private Sub", "End Sub")):
                    continue
                s = re.sub(r"^(Cancel = Not |Response = )", "", s)
                for name, n in calls_in(s, public):
                    req, total = public[name.lower()]
                    if not req <= n <= total:
                        problems.append(f"{m.name}: {name} called with {n} argument(s), expects {req}-{total}: {s}")
        self.assertEqual(problems, [])

    def test_the_checker_finds_a_wrong_count(self):
        names = {"resultbox": (2, 3), "postx": (1, 1)}
        self.assertEqual(calls_in('Call ResultBox("a", 1)', names), [("ResultBox", 2)])
        self.assertEqual(calls_in('ResultBox "a" & (1 + 2), 3', names), [("ResultBox", 2)])
        self.assertEqual(calls_in('x = PostX(a, b(1, 2))', names), [("PostX", 2)])


class StaticModules(VbaModuleChecks, unittest.TestCase):
    module_name = "modDemoData"
    vba = gen_demo.build_demo_vba()


class StaticTestAll(VbaModuleChecks, unittest.TestCase):
    module_name = "modTestAll"
    vba = read("modTestAll")
