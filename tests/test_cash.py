"""Treasury (modCash, frmTreasury / frmCashVoucher / frmCashClosing):
  * the cash scenario replayed on the SQLite mirror with the posting rules of modCash
    (tools/sim.py): which box a document uses, vouchers, cashier closings, and the
    treasury queries agree with the balances;
  * the posting code is pinned to those rules (every cash document fills CashBoxID);
  * schema upgrade safety (new fields, new sequences and permissions on an existing back-end);
  * every control the screen code uses exists; report centre and navigation wiring."""
import os
import re
import sys
import unittest
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import forms as F
import queries as Q
from access_sqlite import AccessOnSqlite, day
from helpers import VbaModuleChecks
from schema import table
from sim import Store


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


FORMS = {m.name: m for m in F.all_forms()}
MAIN, CASHIER = 1, 2          # seeded boxes


class CashScenarioTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.db = AccessOnSqlite()
        s = cls.s = Store(cls.db.con, now=day(2, 9))
        cls.cashier = s.insert("Employees", EmployeeName="TEST كاشير", Username="tcash", RoleID=3)
        cls.keeper = s.insert("Employees", EmployeeName="TEST أمين صندوق", Username="tkeep", RoleID=3,
                              CashBoxID=MAIN)
        cls.prod = s.insert("Products", ProductCode="TEST-C1", ProductName="TEST صنف", CategoryID=1, UnitID=1,
                            SellingPrice=115, CurrentQuantity=0, AverageCost=0, PurchasePrice=0)
        cls.sup = s.insert("Suppliers", SupplierName="TEST مورد", OpeningBalance=0, CurrentBalance=0)
        s.user = 1                                              # administrator: the main safe
        s.cash_voucher("IN", MAIN, D(5000), "OWNER", party="المالك")
        s.purchase(cls.sup, [(cls.prod, 50, 60)], credit=False)          # 50 x 60 + 15% = 3450 cash
        s.now = day(2, 12)
        s.user = cls.cashier                                    # cashier: the cashier box
        s.sale([(cls.prod, 4)])                                 # 460 cash
        s.sale([(cls.prod, 1)], customer=1, credit=False)       # 115 cash
        s.user = 1
        s.expense(2, 100, 15, "TEST كهرباء")                    # admin, cash -> main safe 115
        s.expense(3, 50, 0, "TEST مياه", method=3)              # transfer -> no box
        s.now = day(2, 22)
        s.user = cls.cashier
        cls.closing1, cls.expected1 = s.cash_closing(CASHIER, D(570), "MAIN", MAIN, D(470))   # 5 short, keep 100
        s.now = day(1, 10)
        s.sale([(cls.prod, 2)])                                 # 230
        s.now = day(1, 22)
        cls.closing2, cls.expected2 = s.cash_closing(CASHIER, D(335), "OWNER", None, D(300))  # 5 over, keep 35

    def bal(self, box):
        return self.s.cash_balance(box)

    def test_which_box_a_document_uses(self):
        s = self.s
        self.assertEqual(s.one("SELECT COUNT(*) FROM SalesInvoices WHERE CashBoxID = ?", CASHIER), 3)
        self.assertEqual(s.one("SELECT CashBoxID FROM PurchaseInvoices"), MAIN)
        self.assertEqual(s.one("SELECT COUNT(*) FROM Expenses WHERE CashBoxID Is Null"), 1, "transfer: no box")
        old = s.user
        s.user = self.keeper
        self.assertEqual(s.cash_box(), MAIN, "Employees.CashBoxID wins")
        s.user = old

    def test_first_closing(self):
        self.assertEqual(self.expected1, D(575))                 # 460 + 115
        row = self.s.c.execute("SELECT OpeningBalance, CashIn, CashOut, Difference, TransferAmount, KeptAmount, "
                               "PeriodStart FROM CashClosings WHERE ClosingID = ?", (self.closing1,)).fetchone()
        self.assertEqual(tuple(row), (0, 575, 0, -5, 470, 100, None))
        kinds = self.s.c.execute("SELECT VoucherType, Category, Amount FROM CashVouchers WHERE ClosingID = ? "
                                 "ORDER BY CashVoucherID", (self.closing1,)).fetchall()
        self.assertEqual(kinds, [("OUT", "SHORTAGE", 5), ("TRANSFER", "TRANSFER", 470)])

    def test_second_closing_starts_from_the_first(self):
        row = self.s.c.execute("SELECT OpeningBalance, CashIn, CashOut, ExpectedBalance, Difference, PeriodStart "
                               "FROM CashClosings WHERE ClosingID = ?", (self.closing2,)).fetchone()
        self.assertEqual(tuple(row), (100, 230, 0, 330, 5, day(2, 22)))
        self.assertEqual(self.s.one("SELECT Category FROM CashVouchers WHERE ClosingID = ? AND VoucherType = 'OUT'",
                                    self.closing2), "OWNER")

    def test_balances(self):
        self.assertEqual(self.bal(CASHIER), D(35))               # the float left after the second closing
        self.assertEqual(self.bal(MAIN), D(5000) - D(3450) - D(115) + D(470))

    def test_queries_agree(self):
        db = self.db
        db.params.update({"PeriodStart": day(2), "PeriodEnd": day(-1), "CashBoxID": CASHIER})
        self.assertEqual(db.scalar("SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery"), 35)
        self.assertEqual(db.scalar("SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue('"
                                   + day(2) + "')"), 100)
        self.assertEqual(db.scalar("SELECT OpeningBalance FROM CashDailyQuery WHERE CashDay = DateValue('"
                                   + day(1) + "')"), 100)
        self.assertEqual(db.scalar("SELECT COUNT(*) FROM CashClosingsQuery"), 2)
        self.assertEqual(db.scalar("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = 2"), 35)
        db.params["CashBoxID"] = 0                               # all boxes: transfers cancel out
        total = db.scalar("SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery")
        self.assertEqual(D(str(total)), self.bal(MAIN) + self.bal(CASHIER))

    def test_expense_voucher_records_an_expense_once(self):
        s = self.s
        before = s.one("SELECT COUNT(*) FROM Expenses")
        main_before = self.bal(MAIN)
        s.user = 1
        v = s.cash_voucher("OUT", MAIN, D(40), "EXPENSE", expense_type=8, party="TEST")
        self.assertEqual(s.one("SELECT COUNT(*) FROM Expenses"), before + 1)
        self.assertEqual(self.bal(MAIN), main_before - 40, "the voucher moves the cash, the expense has no box")
        exp = s.one("SELECT ExpenseID FROM CashVouchers WHERE CashVoucherID = ?", v)
        self.assertIsNone(s.one("SELECT CashBoxID FROM Expenses WHERE ExpenseID = ?", exp))
        s.c.execute("DELETE FROM CashVouchers WHERE CashVoucherID = ?", (v,))
        s.c.execute("DELETE FROM Expenses WHERE ExpenseID = ?", (exp,))


class PostingCodeTests(unittest.TestCase):

    def test_every_cash_document_fills_its_box(self):
        sales, purchases = read("modSales"), read("modPurchases")
        for text, name, amount in [(sales, "PostSaleFromCart", "paid"), (sales, "PostSalesReturn", "refunded"),
                                   (sales, "PostCustomerPayment", "Amount"),
                                   (purchases, "PostPurchaseFromCart", "paid"),
                                   (purchases, "PostPurchaseReturn", "refunded"),
                                   (purchases, "PostSupplierPayment", "Amount")]:
            body = proc(text, name)
            self.assertRegex(body, rf"rs!CashBoxID = CashBoxFor\((Nz\()?PaymentMethodID(, CASH_METHOD_ID\))?, "
                                   rf"{amount}\)", name)

    def test_closing_rules_match_the_replay(self):
        body = proc(read("modCash"), "PostCashClosing")
        for line in ['If Destination = "KEEP" Then TransferAmount = 0',
                     "If Not IsNull(periodStart) Then opening = CashBoxBalance(BoxID, periodStart, True)",
                     'cashIn = CashMovesTotal(BoxID, "IN", periodStart)',
                     "expected = opening + cashIn - cashOut", "diff = Counted - expected",
                     'InsertVoucher db, "OUT", BoxID, Null, "SHORTAGE", -diff',
                     'InsertVoucher db, "IN", BoxID, Null, "OVERAGE", diff',
                     'InsertVoucher db, "TRANSFER", BoxID, ToBoxID, "TRANSFER", TransferAmount',
                     'InsertVoucher db, "OUT", BoxID, Null, "OWNER", TransferAmount',
                     "rs!KeptAmount = Counted - TransferAmount", "ws.BeginTrans", "ws.Rollback"]:
            self.assertIn(line, body)

    def test_vouchers_check_balance_and_permission(self):
        body = proc(read("modCash"), "PostCashVoucher")
        self.assertIn('If Not HasPermission("CASH_BOX") Then', body)
        self.assertIn("If Amount > balance Then", body)
        self.assertIn('If Len(msg) = 0 And VoucherType <> "IN" Then', body)

    def test_box_rule_matches_the_replay(self):
        body = proc(read("modCash"), "CurrentCashBoxID")
        self.assertLess(body.index("e.CashBoxID"), body.index('HasPermission("CASH_BOX")'))
        self.assertLess(body.index("BoxType = 'MAIN'"), body.index("BoxType = 'CASHIER'"))
        self.assertIn("If Nz(PaymentMethodID, 0) <> CASH_METHOD_ID Or Amount = 0 Then Exit Function",
                      proc(read("modCash"), "CashBoxFor"))

    def test_expense_screen_keeps_vouchers_and_boxes_consistent(self):
        body = proc(read("modForms"), "ValidateExpense")
        self.assertIn("FROM CashVouchers WHERE ExpenseID", body)
        self.assertIn("frm!CashBoxID.Value = CurrentCashBoxID()", body)
        self.assertIn("frm!CashBoxID.Value = Null", body)

    def test_in_access_test_is_run(self):
        self.assertIn('"TestCash"', read("modTestAll"))


class SchemaUpgradeTests(unittest.TestCase):

    def test_fields_added_to_existing_tables_are_optional(self):
        for t in ("Employees", "SalesInvoices", "SalesReturns", "CustomerPayments", "SupplierPayments",
                  "PurchaseInvoices", "PurchaseReturns", "Expenses"):
            f = next(f for f in table(t).fields if f.name == "CashBoxID")
            self.assertFalse(f.required, t)
            self.assertEqual(f.fk, "CashBoxes.CashBoxID", t)

    def test_new_sequences_and_permissions_reach_an_existing_back_end(self):
        schema = read("modBuildSchema")
        self.assertTrue(table("Sequences").seed_missing and table("Permissions").seed_missing)
        for seq in ("CASH_IN", "CASH_OUT", "CASH_TRANSFER", "CASH_CLOSING"):
            self.assertIn(f"SeedRow \"[SequenceName] = '{seq}'\"", schema)
        self.assertIn('GrantNewPermission "CASH_BOX", "1,2"', schema)
        self.assertIn('GrantNewPermission "CASH_CLOSING", "1,2,3"', schema)
        self.assertIn('If Not BeginSeed("RolePermissions", False) Then Exit Sub', schema)

    def test_boxes_are_seeded(self):
        self.assertEqual(table("CashBoxes").seed_rows, [(1, "الخزينة الرئيسية", "MAIN"), (2, "صندوق الكاشير", "CASHIER")])


class ScreenTests(unittest.TestCase):

    SCREENS = {"frmTreasury": ("Treasury", "PrintCashDay"),
               "frmCashVoucher": ("Voucher", "SaveCashVoucher"),
               "frmCashClosing": ("Closing", "SaveCashClosing")}

    def test_controls_used_by_the_code_exist(self):
        text = read("modCash")
        procs = re.findall(r"^Public (?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)", text, re.M | re.S)
        bodies = {name: proc(text, name) for name in procs}
        for form, (prefix, extra) in self.SCREENS.items():
            names = {c.name for c in FORMS[form].controls}
            for pname, body in bodies.items():
                if pname.startswith(prefix) or pname == extra:
                    for ctl in set(re.findall(r"frm!(\w+)", body)):
                        self.assertIn(ctl, names, f"{form}: {pname} uses {ctl}")

    def test_expenses_screen(self):
        names = {c.name for c in FORMS["frmExpenses"].controls}
        for ctl in ("CashBoxID", "btnNewType", "btnExpenseTypes"):
            self.assertIn(ctl, names)
        self.assertIn("AddExpenseType", read("modCash"))

    def test_navigation_and_permissions(self):
        self.assertIn("Treasury", {n.key for n in F.NAV_ITEMS})
        self.assertIn("Treasury", {t[0] for t in F.LAUNCH_TILES})
        self.assertEqual(F.SCREEN_PERMISSIONS["frmTreasury"], "CASH_CLOSING")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCashVoucher"], "CASH_BOX")
        perms = {r[0] for r in table("Permissions").seed_rows}
        self.assertTrue({"CASH_BOX", "CASH_CLOSING"} <= perms)

    def test_report_centre(self):
        names = {c.name for c in FORMS["frmReportCenter"].controls}
        self.assertTrue({"cboCashBox", "cboExpenseType"} <= names)
        for r in F.REPORTS:
            if "b" in r.needs:
                self.assertIn("CashBoxID", Q.query(r.query).params, r.key)
            if "e" in r.needs:
                self.assertIn("e.ExpenseTypeID", Q.query(r.query).sql, r.key)
        screens = read("modScreens")
        for letter in ('"b"', '"e"', '"#"'):
            self.assertIn(f"HasNeed(needs, {letter})", screens)
        self.assertIn('"CashBoxID"', read("modQueryParams"))


class Static_modCash(VbaModuleChecks, unittest.TestCase):
    module_name = "modCash"
    vba = read("modCash")


if __name__ == "__main__":
    unittest.main()
