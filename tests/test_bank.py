"""Banks and bank reconciliation (modBank, Banks, BankTransactions, BankReconciliations, BankClearings,
journal sources BANK_OPENING / BANK_TX, qryBankItems, BankBalanceQuery) on the SQLite mirror:
  * each bank has its account 120000 + BankID under 1210, with its opening balance;
  * a bank transfer (payment method 3) of a document is booked in its bank, Mada stays in 1200;
  * deposit / withdraw move the cash box too, the Mada settlement books the fee and its VAT,
    transfers and other transactions; every entry is balanced;
  * the reconciliation: book - operations not in the statement = statement balance;
  * the VBA keeps the rules, the screens, the permission, and the list row sources are built."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
from schema import table, ACCOUNT_TREE
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class BankTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.bank = self.s.insert("Banks", BankName="TEST بنك", OpeningBalance=1000, OpeningDate=day(10), IsActive=1)
        self.bank2 = self.s.insert("Banks", BankName="TEST بنك 2", OpeningBalance=0, OpeningDate=day(10), IsActive=1)
        for b, name in [(self.bank, "TEST بنك"), (self.bank2, "TEST بنك 2")]:       # EnsureAccounts
            self.s.insert("Accounts", AccountCode=120000 + b, AccountName=name, AccountType="ASSET", ParentCode=1210,
                          IsPosting=1, IsSystem=1)
        self.s.sync_journal()
        self.box = self.db.ids["BOXM"]

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def book(self, bank):
        return round(self.s.one("SELECT Balance FROM BankBalanceQuery WHERE BankID = ?", bank), 4)

    def tx(self, kind, amount, **kw):
        no = self.s.next_number("BANK_TX")
        tid = self.s.insert("BankTransactions", TxNumber=no, TxDate=day(1), TxType=kind, BankID=kw.pop("bank", self.bank),
                            Amount=amount, FeeAmount=kw.pop("fee", 0), FeeVAT=kw.pop("fee_vat", 0),
                            Description="TEST", EmployeeID=1, **kw)
        self.s.sync_journal()
        return tid

    def balanced(self):
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)

    def test_opening_balance(self):
        self.assertEqual(self.balance(120000 + self.bank), 1000)
        self.assertEqual(self.book(self.bank), 1000)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType = 'BANK_OPENING'"), 1)

    def test_deposit_and_withdraw_move_the_cash_box(self):
        before = self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box)
        self.tx("DEPOSIT", 300, CashBoxID=self.box)
        self.tx("WITHDRAW", 50, CashBoxID=self.box)
        self.assertEqual(self.book(self.bank), 1250)
        self.assertAlmostEqual(self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box),
                               before - 250, places=4)
        self.assertAlmostEqual(self.balance(110000 + self.box),
                               self.s.one("SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = ?", self.box), places=4)
        self.balanced()

    def test_mada_settlement(self):
        mada = self.balance(1200)
        charges = self.balance(5610)
        vat_in = self.balance(1500)
        self.tx("SETTLEMENT", 100, fee=1.5, fee_vat=0.23)
        self.assertAlmostEqual(self.book(self.bank), 1098.27, places=4)
        self.assertAlmostEqual(self.balance(1200), mada - 100, places=4)
        self.assertAlmostEqual(self.balance(5610), charges + 1.5, places=4)
        self.assertAlmostEqual(self.balance(1500), vat_in + 0.23, places=4)
        self.balanced()

    def test_transfer_and_other(self):
        self.tx("TRANSFER", 200, ToBankID=self.bank2)
        self.tx("OTHER_OUT", 11.5, CounterAccount=5610, fee_vat=1.5)
        self.tx("OTHER_IN", 40, CounterAccount=4200)
        self.assertAlmostEqual(self.book(self.bank), 1000 - 200 - 11.5 + 40, places=4)
        self.assertEqual(self.book(self.bank2), 200)
        self.balanced()

    def test_bank_transfer_documents_go_to_their_bank(self):
        self.s.customer_payment(self.db.ids["C2"], 100)
        pay = self.s.one("SELECT Max(PaymentID) FROM CustomerPayments")
        self.c.execute("UPDATE CustomerPayments SET PaymentMethodID = 3, CashBoxID = Null, BankID = ? WHERE PaymentID = ?",
                       (self.bank, pay))
        mada = self.balance(1200)
        self.s.sync_journal()
        self.assertEqual(self.book(self.bank), 1100)
        self.assertEqual(self.balance(1200), mada)
        self.c.execute("UPDATE CustomerPayments SET BankID = Null WHERE PaymentID = ?", (pay,))   # no default bank
        self.s.sync_journal()
        self.assertEqual(self.book(self.bank), 1000)
        self.assertEqual(self.balance(1200), mada + 100)

    def test_reconciliation(self):
        dep = self.tx("DEPOSIT", 300, CashBoxID=self.box)
        self.tx("OTHER_OUT", 20, CounterAccount=5610)
        rid = self.s.insert("BankReconciliations", ReconNumber="TEST-REC", BankID=self.bank, StatementDate=day(0),
                            StatementBalance=1300, Status="OPEN", EmployeeID=1)
        for st, sid in [("BANK_OPENING", self.bank), ("BANK_TX", dep)]:
            amount = self.s.one("SELECT ItemAmount FROM qryBankItems WHERE BankID = ? AND SourceType = ? AND SourceID = ?",
                                self.bank, st, sid)
            self.s.insert("BankClearings", ReconciliationID=rid, BankID=self.bank, SourceType=st, SourceID=sid,
                          ClearedAmount=amount)
        book = self.s.one("SELECT Sum(ItemAmount) FROM qryBankItems WHERE BankID = ?", self.bank)
        outstanding = self.s.one("SELECT Sum(ItemAmount) FROM qryBankItems WHERE BankID = ? AND IsCleared = 0", self.bank)
        self.assertEqual((book, outstanding), (1280, -20))
        self.assertEqual(1300 - (book - outstanding), 0)               # ReconFigures: no difference
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM qryBankItems WHERE IsCleared = 1 AND ClearedAmount <> ItemAmount"),
                         0)
        self.c.execute("UPDATE BankTransactions SET Amount = 310 WHERE BankTxID = ?", (dep,))        # changed later
        self.s.sync_journal()
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM qryBankItems WHERE IsCleared = 1 AND ClearedAmount <> ItemAmount"),
                         1)                                                                       # ChangedClearings


class BankCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modBank")

    def test_rules(self):
        problem = proc(self.text, "BankTxProblem")
        for part in ('Case "DEPOSIT", "WITHDRAW"', "CashBoxBalance(", "FeeAmount + FeeVAT >= Amount",
                     "CLng(ToBankID) = CLng(BankID)", "CounterAccountProblem(", "ClosedPeriodProblem(TxDate)"):
            self.assertIn(part, problem)
        self.assertIn("NOT IN (1100, 1210) AND AccountCode NOT IN (1190, 1200, 1300, 2100)",
                      proc(self.text, "CounterAccountProblem"))
        self.assertIn("الحركة مطابقة في تسوية بنكية", proc(self.text, "DeleteBankTx"))
        self.assertIn("If diff <> 0 Then", proc(self.text, "FinishReconciliation"))
        self.assertIn("تُعاد فتح آخر تسوية للبنك فقط", proc(self.text, "ReopenReconciliation"))
        self.assertIn("BANK_TRANSFER_METHOD_ID", proc(self.text, "BankFor"))
        for module in ("modSales", "modPurchases"):
            self.assertEqual(read(module).count("rs!BankID = BankFor("), 3, module)
        self.assertIn("frm!BankID.Value = BankFor(", read("modForms"))
        self.assertIn("120000 + k.BankID, k.BankName, 'ASSET', 1210", read("modJournal"))
        self.assertIn("qryJournalBankTx", read("modJournal"))
        self.assertIn('Case "CashBoxes", "Banks"', read("modClosing"))
        self.assertIn('"TestBank"', read("modTestAll"))

    def test_schema_screens(self):
        self.assertIn((1210, "الحسابات البنكية", "ASSET", 11, False, True), [tuple(r) for r in ACCOUNT_TREE])
        for t in ("SalesInvoices", "SalesReturns", "CustomerPayments", "PurchaseInvoices", "PurchaseReturns",
                  "SupplierPayments", "Expenses"):
            self.assertIn("BankID", [f.name for f in table(t).fields], t)
        for screen in ("frmBanks", "frmBankTx", "frmBankRecon"):
            self.assertEqual(F.SCREEN_PERMISSIONS[screen], "BANKS")
        for form, prefixes in [("frmBankTx", ("BankTx", "SaveBankTx", "DeleteSelectedBankTx")),
                               ("frmBankRecon", ("BankRecon",))]:
            names = {c.name for c in MODELS[form].controls}
            for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
                if not pname.startswith(prefixes):
                    continue
                for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                    self.assertIn(ctl, names, f"{form} {pname}: {ctl}")
        self.assertIn("btnBanks", {c.name for c in MODELS["frmTreasury"].controls})
        self.assertIn("BankID", {c.name for c in MODELS["frmExpenses"].controls})

    def test_list_row_sources_are_built(self):
        # a list with a fixed row source gets it in BuildForms (it was left out before)
        build = read("modBuildForms")
        for form, lst in [("frmPeriodClosing", "lstHistory"), ("frmVatReturn", "lstReturns"), ("frmBankTx", "lstTx")]:
            m = re.search(rf'AddList\("{lst}".*?\n(.*?)\n', build[build.index(f'"{form}"'):])
            self.assertTrue(m and m.group(1).strip().startswith("c.RowSource = "), (form, lst))


class Static_modBank(VbaModuleChecks, unittest.TestCase):
    module_name = "modBank"
    vba = read("modBank")


if __name__ == "__main__":
    unittest.main()
