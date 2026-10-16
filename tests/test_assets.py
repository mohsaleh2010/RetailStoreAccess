"""Fixed assets and depreciation (modAssets, FixedAssets, DepreciationRuns, AssetDepreciations, journal sources
ASSET / ASSET_DISPOSAL / DEPRECIATION, FixedAssetsQuery) on the SQLite mirror with tools/sim.py
(Store.depreciate = RecordDepreciation):
  * buying: the asset account and input VAT against a cash box (its balance too), a bank or the opening
    balances with the depreciation until then;
  * monthly straight line depreciation, never below the salvage value, 5600 / 1790;
  * selling / scrapping: accumulated depreciation and price against the cost, gain 4500 or loss 5650;
  * the VBA keeps the rules, the screens, the permission, the report and the demo asset."""
import datetime as dt
import os
import re
import unittest
from decimal import Decimal as D

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day, TODAY
import forms as F
import demo_data as DD
from schema import ACCOUNT_TREE
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def month_end(months_ago):
    first = TODAY.replace(day=1)
    for _ in range(months_ago):
        first = (first - dt.timedelta(days=1)).replace(day=1)
    nxt = (first + dt.timedelta(days=32)).replace(day=1)
    return (nxt - dt.timedelta(days=1)).isoformat()


class AssetTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.s.sync_journal()
        self.box = self.db.ids["BOXM"]
        self.start = month_end(5)[:8] + "01 00:00:00"             # first day, five months ago

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def asset(self, cost, salvage, months, source="CASHBOX", vat=0, opening=0, **kw):
        aid = self.s.insert("FixedAssets", AssetCode=self.s.next_number("FIXED_ASSET"), AssetName="TEST أصل",
                            AssetAccount=1720, PurchaseDate=self.start, Cost=cost, InputVAT=vat, SalvageValue=salvage,
                            UsefulLifeMonths=months, DepStartDate=self.start, SourceType=source,
                            CashBoxID=self.box if source == "CASHBOX" else None, OpeningAccumDep=opening,
                            Status="ACTIVE", DisposalProceeds=0, DisposalAccumDep=0, EmployeeID=1, **kw)
        self.s.sync_journal()
        return aid

    def balanced(self):
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)

    def test_bought_from_a_cash_box(self):
        box_before = self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box)
        vat_before = self.balance(1500)
        self.asset(1000, 100, 10, vat=150)
        self.assertEqual(self.balance(1720), 1000)
        self.assertAlmostEqual(self.balance(1500), vat_before + 150, places=4)
        self.assertAlmostEqual(self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box),
                               box_before - 1150, places=4)
        self.assertAlmostEqual(self.balance(110000 + self.box),
                               self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box), places=4)
        self.balanced()

    def test_owned_before_the_program(self):
        opening = self.balance(3900)
        self.asset(6000, 600, 36, source="OPENING", opening=1500)
        self.assertEqual(self.balance(1720), 6000)
        self.assertEqual(self.balance(1790), -1500)
        self.assertAlmostEqual(self.balance(3900), opening - 4500, places=4)
        self.balanced()

    def test_monthly_depreciation_until_the_salvage_value(self):
        aid = self.asset(1000, 100, 4)                                  # 225 a month, four months
        for k in range(5, 0, -1):                                       # five months recorded
            self.s.depreciate(month_end(k))
        amounts = [r[0] for r in self.c.execute("SELECT Amount FROM AssetDepreciations WHERE AssetID = ? "
                                                "ORDER BY DepreciationID", (aid,))]
        self.assertEqual(amounts, [225, 225, 225, 225])                 # nothing in the fifth month
        self.assertEqual(self.balance(5600), 900)
        self.assertEqual(self.balance(1790), -900)
        row = self.c.execute("SELECT AccumDep, BookValue, MonthlyDep, DepMonths FROM FixedAssetsQuery WHERE AssetID = ?",
                             (aid,)).fetchone()
        self.assertEqual(row, (900, 100, 225, 4))
        self.balanced()

    def test_rounding_ends_on_the_salvage_value(self):
        aid = self.asset(1000, 0, 3)                                    # 333.33, the last month 333.34
        for k in (5, 4, 3, 2):
            self.s.depreciate(month_end(k))
        amounts = [D(str(r[0])) for r in self.c.execute("SELECT Amount FROM AssetDepreciations WHERE AssetID = ?",
                                                         (aid,))]
        self.assertEqual(sum(amounts), 1000)
        self.assertEqual(amounts[-1], D("333.34"))

    def test_sold_with_a_gain_and_scrapped_with_a_loss(self):
        a1 = self.asset(1000, 0, 10)
        a2 = self.asset(500, 0, 10)
        for k in (5, 4):
            self.s.depreciate(month_end(k))                             # 100 + 100, 50 + 50
        when = month_end(3)[:8] + "15 00:00:00"
        self.c.execute("UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = ?, DisposalProceeds = 900, "
                       "DisposalTo = 'CASHBOX', DisposalCashBoxID = ?, DisposalAccumDep = 200 WHERE AssetID = ?",
                       (when, self.box, a1))
        self.c.execute("UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = ?, DisposalProceeds = 0, "
                       "DisposalTo = 'NONE', DisposalAccumDep = 100 WHERE AssetID = ?", (when, a2))
        self.s.sync_journal()
        self.assertEqual(self.balance(4500), -100)                      # 900 + 200 - 1000
        self.assertEqual(self.balance(5650), 400)                       # 500 - 100
        self.assertEqual(self.balance(1720), 0)
        self.assertEqual(self.balance(1790), 0)
        _, total = self.s.depreciate(month_end(3))                     # the disposal month: nothing
        self.assertEqual(total, 0)
        self.balanced()

    # ---------------------------------------------------------------- docs/65
    def amounts(self, aid):
        return [D(str(r[0])) for r in self.c.execute("SELECT Amount FROM AssetDepreciations WHERE AssetID = ? "
                                                      "ORDER BY DepreciationID", (aid,))]

    def first_day(self, months_ago):
        return month_end(months_ago)[:8] + "01"

    def center_balance(self, account, center):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ? "
                                "AND CostCenterID = ?", account, center), 4)

    def test_declining_balance(self):
        a1 = self.asset(1200, 0, 4, DepMethod="DB", DepRate=0)          # 24 / 4 = 600 % a year: half the book value a month
        a2 = self.asset(1200, 100, 4, DepMethod="DB", DepRate=0)
        a3 = self.asset(1000, 0, 24, DepMethod="DB", DepRate=0.6)
        for k in (5, 4, 3, 2, 1):
            self.s.depreciate(month_end(k))
        self.assertEqual(self.amounts(a1), [600, 300, 150, 150])         # all that is left in the last month
        self.assertEqual(self.amounts(a2), [600, 300, 150, 50])          # never below the salvage value
        self.assertEqual(self.amounts(a3)[:2], [D("50"), D("47.5")])     # 1000 x 60 % / 12, then on 950
        row = self.c.execute("SELECT MethodName, MonthlyDep FROM FixedAssetsQuery WHERE AssetID = ?", (a3,)).fetchone()
        book = 1000 - float(sum(self.amounts(a3)))
        self.assertEqual(row[0], "متناقص")
        self.assertAlmostEqual(row[1], round(book * 0.6 / 12, 2), places=2)
        self.balanced()

    def test_revaluation_surplus_loss_and_depreciation(self):
        aid = self.asset(1000, 0, 10)
        self.s.depreciate(month_end(5))                                  # 100, book value 900
        self.s.revalue(aid, self.first_day(4), 1080, 8)                  # +180 surplus
        self.assertEqual(self.balance(3500), -180)
        self.assertEqual(self.balance(1720), 1180)
        self.s.depreciate(month_end(4))
        self.assertEqual(self.amounts(aid)[-1], 135)                     # 1080 / 8
        self.s.revalue(aid, self.first_day(3), 700, 6)                   # 945 -> 700: -180 surplus, 65 loss
        self.assertEqual(self.balance(3500), 0)
        self.assertEqual(self.balance(5660), 65)
        self.s.depreciate(month_end(3))
        self.assertEqual(self.amounts(aid)[-1], D("116.67"))             # 700 / 6
        self.s.revalue(aid, self.first_day(2), 650, 5)                   # 583.33 -> 650: the 65 loss back, then surplus
        self.assertEqual(self.balance(5660), 0)
        self.assertAlmostEqual(self.balance(3500), -1.67, places=2)
        row = self.c.execute("SELECT BookValue, MonthlyDep, RevalTotal FROM FixedAssetsQuery WHERE AssetID = ?",
                             (aid,)).fetchone()
        self.assertAlmostEqual(row[0], 650, places=2)
        self.assertEqual(row[1], 130)
        self.assertAlmostEqual(row[2], 1080 - 900 + 700 - 945 + 650 - 583.33, places=2)
        retained = self.balance(3300)
        accum = float(self.s.asset_depreciated(aid))
        self.c.execute("UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = ?, DisposalProceeds = 600, "
                       "DisposalTo = 'CASHBOX', DisposalCashBoxID = ?, DisposalAccumDep = ? WHERE AssetID = ?",
                       (self.first_day(1) + " 00:00:00", self.box, accum, aid))
        self.s.sync_journal()
        self.assertEqual(self.balance(1720), 0)
        self.assertEqual(self.balance(1790), 0)
        self.assertEqual(self.balance(3500), 0)                          # the surplus left to retained earnings
        self.assertAlmostEqual(self.balance(3300), retained - 1.67, places=2)
        self.assertAlmostEqual(self.balance(5650), 50, places=2)         # 650 - 600
        self.balanced()

    def test_transfer_between_cost_centres(self):
        c1 = self.s.insert("CostCenters", CenterCode="T1", CenterName="TEST فرع 1", IsDefault=0, IsActive=1)
        c2 = self.s.insert("CostCenters", CenterCode="T2", CenterName="TEST فرع 2", IsDefault=0, IsActive=1)
        aid = self.asset(1000, 0, 10, CostCenterID=c1)
        self.s.depreciate(month_end(5))
        self.s.transfer_asset(aid, self.first_day(4), c2)
        self.s.depreciate(month_end(4))
        self.assertEqual(self.center_balance(1720, c1), 0)
        self.assertEqual(self.center_balance(1720, c2), 1000)
        self.assertEqual(self.center_balance(1790, c1), 0)
        self.assertEqual(self.center_balance(1790, c2), -200)
        self.assertEqual(self.center_balance(5600, c1), 100)             # the month before stays on the old branch
        self.assertEqual(self.center_balance(5600, c2), 100)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalLines AS l INNER JOIN JournalEntries AS e ON "
                                    "l.EntryID = e.EntryID WHERE e.SourceType = 'ASSET' AND l.CostCenterID = ?", c1), 2)
        self.balanced()

    def test_sale_with_output_vat_in_the_vat_return(self):
        aid = self.asset(1000, 0, 10, vat=150)
        box = self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box)
        vat_out = self.balance(2200)
        when = self.first_day(1) + " 00:00:00"
        self.c.execute("UPDATE FixedAssets SET Status = 'DISPOSED', DisposalDate = ?, DisposalProceeds = 800, "
                       "DisposalVAT = 120, DisposalTo = 'CASHBOX', DisposalCashBoxID = ?, DisposalAccumDep = 0 "
                       "WHERE AssetID = ?", (when, self.box, aid))
        self.s.sync_journal()
        self.assertEqual(self.balance(2200), vat_out - 120)
        self.assertEqual(self.balance(5650), 200)                        # the VAT is not part of the gain or loss
        self.assertAlmostEqual(self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box),
                               box + 920, places=4)
        rows = self.c.execute("SELECT Side, Category, Amount, VAT FROM qryVatReturnLines WHERE Amount IN (800, 1000) "
                              "ORDER BY Side").fetchall()
        self.assertIn(("S", "S", 800, 120), rows)
        self.assertIn(("P", "S", 1000, 150), rows)
        self.balanced()

    def test_demo_asset(self):
        step = [p for p in DD.PLAN if p["op"] == "asset"][0]
        self.assertEqual(step["account"], 1720)
        r = DD.simulate()
        self.assertEqual(r.db.con.execute("SELECT COUNT(*) FROM FixedAssetsQuery").fetchone()[0], 1)


class AssetCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modAssets")

    def test_rules(self):
        problem = proc(self.text, "AssetProblem")
        for part in ("بداية الإهلاك قبل شهر الشراء", "القيمة المتبقية والإهلاك السابق لا يتجاوزان التكلفة",
                     'Case "OPENING"', "CounterAccountProblem(CounterAccount)", "CashBoxBalance(", "ClosedPeriodProblem("):
            self.assertIn(part, problem)
        self.assertIn("AccountType = 'ASSET' AND Level2Code = 12 AND AccountCode <> 1790", proc(self.text, "AssetAccountProblem"))
        self.assertIn("على الأصل قيود إهلاك", proc(self.text, "AssetLocked"))
        run = proc(self.text, "RecordDepreciation")
        for part in ("الأشهر تُسجَّل بالترتيب", "RunMonth > Date", "ClosedPeriodProblem(RunMonth)", "d!LineNo = n"):
            self.assertIn(part, run)
        dep = proc(self.text, "DepreciationFor")
        self.assertIn('rs!Status = "ACTIVE" Or Nz(rs!DisposalDate, 0) > RunMonth', dep)
        self.assertIn("If monthly > remaining Then monthly = remaining", dep)
        self.assertIn("If remaining - monthly < 1 Then monthly = remaining", dep)
        self.assertIn("احذف قيود الإهلاك من شهر الاستبعاد أولًا", proc(self.text, "DisposeAsset"))
        self.assertIn("في القيد أصل مستبعد بعده", proc(self.text, "DeleteLastRun"))
        self.assertIn("qryJournalAsset,qryJournalDepreciation", read("modJournal"))
        self.assertIn('"TestAssets"', read("modTestAll"))

    def test_accounts_screens_report(self):
        tree = {r[0]: r for r in ACCOUNT_TREE}
        for code in (1790, 5600, 4500, 5650):
            self.assertTrue(tree[code][5], code)
        for form, prefixes, skip in [("frmAssets", ("Asset", "PrintAssets"), ("AssetChanges",)),
                                     ("frmDepreciation", ("Depreciation",), ()),
                                     ("frmAssetChanges", ("AssetChanges",), ())]:
            self.assertEqual(F.SCREEN_PERMISSIONS[form], "FIXED_ASSETS")
            names = {c.name for c in MODELS[form].controls}
            for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
                if pname.startswith(prefixes) and not (skip and pname.startswith(skip)):
                    for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                        self.assertIn(ctl, names, f"{form} {pname}: {ctl}")
        self.assertIn("FIXED_ASSETS", {r.key for r in F.REPORTS})
        self.assertIn("btnAssets", {c.name for c in MODELS["frmFinancials"].controls})


class Static_modAssets(VbaModuleChecks, unittest.TestCase):
    module_name = "modAssets"
    vba = read("modAssets")


if __name__ == "__main__":
    unittest.main()
