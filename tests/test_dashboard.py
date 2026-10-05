"""Phase 9 checks: the dashboard of frmMain.

  * DashboardQuery and qryDashboardTopProducts on the SQLite mirror with the
    test fixture: every figure is compared with a hand calculation;
  * the dashboard never changes the report-centre parameters;
  * every tile has its click handler and every amount respects the
    DASHBOARD_FINANCIAL permission;
  * static checks on modDashboard.

Run:  python3 -m unittest discover -s tests -v
"""

import os
import re
import unittest

from access_sqlite import AccessOnSqlite, day
from helpers import ROOT, VbaModuleChecks
import forms as F
import queries as Q


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


class DashboardQueryTests(unittest.TestCase):
    """"Today" = 5 days ago (credit sale TEST-INV-2), "month" = the 30 days before it."""

    @classmethod
    def setUpClass(cls):
        cls.db = db = AccessOnSqlite()
        db.load_fixture()
        db.params.update({"DashDay": day(5), "DashEnd": day(4), "DashMonth": day(30)})
        cur = db.con.execute('SELECT * FROM "DashboardQuery"')
        cls.row = dict(zip([d[0] for d in cur.description], cur.fetchone()))

    def test_today(self):
        self.assertAlmostEqual(self.row["TodaySales"], 460)          # 2 x 115 + 10 x 23
        self.assertEqual(self.row["TodayInvoices"], 1)

    def test_month(self):
        # TEST-INV-1 (10 days ago) 1150 + TEST-INV-2 460; the return 2 days ago is after DashEnd
        self.assertAlmostEqual(self.row["MonthSales"], 1610)
        self.assertAlmostEqual(self.row["MonthVAT"], 210)
        self.assertEqual(self.row["MonthInvoices"], 2)
        self.assertAlmostEqual(self.row["MonthExpenses"], 200)       # electricity 6 days ago, rent excluded

    def test_balances_and_stock(self):
        self.assertAlmostEqual(self.row["CustomerDebt"], 164)
        self.assertEqual(self.row["DebtorCount"], 1)                 # the cash customer has no balance
        self.assertAlmostEqual(self.row["SupplierDue"], 2855)
        self.assertAlmostEqual(self.row["StockValue"], 83 * 60 + 186 * 10 + 20 * 10)
        self.assertEqual(self.row["LowStockCount"], 1)

    def test_returns_reduce_sales(self):
        db = self.db
        saved = dict(db.params)
        db.params.update({"DashDay": day(2), "DashEnd": day(1)})     # the day of the credit note
        try:
            today = db.con.execute('SELECT TodaySales, TodayInvoices FROM "DashboardQuery"').fetchone()
        finally:
            db.params.clear()
            db.params.update(saved)
        self.assertAlmostEqual(today[0], -46)
        self.assertEqual(today[1], 0)

    def test_top_products(self):
        rows = self.db.con.execute('SELECT ProductID, NetQty FROM "qryDashboardTopProducts" '
                                   'ORDER BY NetQty DESC').fetchall()
        self.assertEqual(rows, [(self.db.ids["P1"], 12), (self.db.ids["P2"], 10)])

    def test_one_row_without_parameters(self):
        db = AccessOnSqlite()                          # empty database, no parameters set
        self.assertEqual(db.con.execute('SELECT COUNT(*) FROM "DashboardQuery"').fetchone()[0], 1)

    def test_own_parameters_only(self):
        for name in ("DashboardQuery", "qryDashboardTopProducts"):
            q = next(q for q in Q.QUERIES if q.name == name)
            used = set(re.findall(r"QDate\('(\w+)'\)", q.sql))
            self.assertTrue(used <= {"DashDay", "DashMonth", "DashEnd"}, name)
            self.assertEqual(set(q.params), used, name)


class DashboardCodeTests(unittest.TestCase):
    text = read("modDashboard")

    def test_every_tile_has_a_handler(self):
        main = next(m for m in F.all_forms() if m.name == "frmMain")
        code = "\n".join(main.code)
        body = self.text.split("Public Sub DashboardTileClick(")[1].split("\nEnd Sub")[0]
        for i, (key, _) in enumerate(F.DASHBOARD_TILES):
            self.assertIn(f'DashboardTileClick "{key}"', code)
            self.assertRegex(body, rf'Case [^\n]*"{key}"', key)
            self.assertIn(f"SetTile frm, {i + 1},", self.text)

    def test_amounts_hidden_without_permission(self):
        body = self.text.split("Public Sub DashboardRefresh(")[1].split("\nEnd Sub")[0]
        self.assertIn('financial = HasPermission("DASHBOARD_FINANCIAL")', body)
        for field in ("TodaySales", "MonthSales", "CustomerDebt", "SupplierDue", "StockValue", "MonthExpenses"):
            self.assertIn(f"Money(rs!{field}, financial)", body, field)

    def test_launcher_tiles(self):
        main = next(m for m in F.all_forms() if m.name == "frmMain")
        ctl = {c.name: c for c in main.controls}
        targets = {i.key: i.target for i in F.NAV_ITEMS}
        self.assertEqual(len(F.LAUNCH_TILES), 12)
        for key, *_ in F.LAUNCH_TILES:
            btn = ctl[f"btnTile{key}"]
            self.assertTrue(btn.props["Transparent"])
            self.assertEqual(btn.props.get("Tag", ""), targets[key])
            for part in ("boxNav", "icoTile", "lblTile"):
                c = ctl[f"{part}{key}"]
                self.assertEqual((c.x, c.w), (btn.x, btn.w), part + key)
                self.assertTrue(btn.y <= c.y and c.y + c.h <= btn.y + btn.h, part + key)   # under the button
                self.assertTrue(c.decorative)
        self.assertIn('frm.Controls("boxNav" & Mid$(ctl.Name, 8)).BackColor = RGB(205, 210, 218)',
                      read("modSecurityScreens"))

    def test_report_period_is_restored(self):
        body = self.text.split("Private Function MonthNetProfit(")[1].split("\nEnd Function")[0]
        self.assertIn('savedStart = TempVars("PeriodStart")', body)
        self.assertIn('TempVars.Add "PeriodStart", savedStart', body)

    def test_dashboard_refreshes(self):
        main = next(m for m in F.all_forms() if m.name == "frmMain")
        self.assertEqual(main.form_props.get("TimerInterval"), 300000)
        self.assertIn("DashboardRefresh frm", read("modScreens"))
        self.assertIn("If Not g_SilentMode Then LowStockAlert", read("modScreens"))


class StaticModDashboard(VbaModuleChecks, unittest.TestCase):
    module_name = "modDashboard"
    vba = read("modDashboard")
