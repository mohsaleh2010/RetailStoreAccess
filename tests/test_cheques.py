"""Received and issued cheques (modCheque, Cheques, frmCheques, journal sources CHEQUE / CHEQUE_STATUS) on
the SQLite mirror:
  * a received cheque pays the customer's balance (ledger, aging, 1300) and waits in 1250; collected it
    reaches the bank, bounced it is owed again (and ages from the bounce day);
  * an issued cheque pays the supplier and waits in 2110 until the bank pays it, or goes back on bounce;
  * every entry is balanced; the VBA keeps the rules; the screen and permission."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day, TODAY
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


class ChequeTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.bank = self.s.insert("Banks", BankName="TEST بنك", OpeningBalance=0, OpeningDate=day(10), IsActive=1)
        self.s.insert("Accounts", AccountCode=120000 + self.bank, AccountName="TEST بنك", AccountType="ASSET",
                      ParentCode=1210, IsPosting=1, IsSystem=1)
        self.s.sync_journal()
        self.cust, self.supp = self.db.ids["C2"], self.db.ids["S1"]

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def party(self, kind, pid):
        q = ("SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = ?" if kind == "C" else
             "SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = ?")
        return round(self.s.one(q, pid), 4)

    def cheque(self, direction, amount, days_ago=2, due_in=0):          # PostCheque
        pid = self.cust if direction == "IN" else self.supp
        cid = self.s.insert("Cheques", ChequeRef=self.s.next_number("CHEQUE"), Direction=direction,
                            CustomerID=pid if direction == "IN" else None,
                            SupplierID=pid if direction == "OUT" else None, ChequeNo=f"TEST-{amount}",
                            BankID=self.bank if direction == "OUT" else None, IssueDate=day(days_ago),
                            DueDate=day(days_ago - due_in), Amount=amount, Status="PENDING", EmployeeID=1)
        self.s.adjust("Customers" if direction == "IN" else "Suppliers",
                      "CustomerID" if direction == "IN" else "SupplierID", pid, -amount)
        self.s.sync_journal()
        return cid

    def status(self, cid, new, direction):                             # SetChequeStatus
        self.c.execute("UPDATE Cheques SET Status = ?, StatusDate = ?, BankID = ? WHERE ChequeID = ?",
                       (new, day(1), self.bank, cid))
        if new == "BOUNCED":
            pid = self.cust if direction == "IN" else self.supp
            amount = self.s.one("SELECT Amount FROM Cheques WHERE ChequeID = ?", cid)
            self.s.adjust("Customers" if direction == "IN" else "Suppliers",
                          "CustomerID" if direction == "IN" else "SupplierID", pid, amount)
        self.s.sync_journal()

    def controls_match(self):
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)
        self.assertAlmostEqual(self.balance(1300), self.s.one("SELECT Sum(CurrentBalance) FROM Customers"), places=4)
        self.assertAlmostEqual(-self.balance(2100), self.s.one("SELECT Sum(CurrentBalance) FROM Suppliers"), places=4)
        for kind, q, key in [("C", "CustomerBalanceQuery", "CustomerID"), ("S", "SupplierBalanceQuery", "SupplierID")]:
            rows, credits = self.s.aging(kind, TODAY.isoformat())
            total = sum(float(r[6]) for r in rows) - sum(float(v) for v in credits.values())
            self.assertAlmostEqual(total, self.s.one(f"SELECT Sum(Balance) FROM {q}"), places=4, msg=kind)

    def test_received_cheque_collected(self):
        before, pending = self.party("C", self.cust), self.balance(1250)
        cid = self.cheque("IN", 100)
        self.assertAlmostEqual(self.party("C", self.cust), before - 100, places=4)
        self.assertAlmostEqual(self.balance(1250), pending + 100, places=4)
        self.controls_match()
        self.status(cid, "COLLECTED", "IN")
        self.assertAlmostEqual(self.balance(1250), pending, places=4)
        self.assertEqual(self.balance(120000 + self.bank), 100)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM qryBankItems WHERE BankID = ? AND SourceType = 'CHEQUE_STATUS'",
                                    self.bank), 1)                       # in the bank reconciliation
        self.controls_match()

    def test_received_cheque_bounced(self):
        before = self.party("C", self.cust)
        aging_before = self.s.aging("C", TODAY.isoformat())
        cid = self.cheque("IN", 100)
        self.assertNotEqual(self.s.aging("C", TODAY.isoformat()), aging_before)
        self.status(cid, "BOUNCED", "IN")
        self.assertAlmostEqual(self.party("C", self.cust), before, places=4)
        self.assertEqual(self.balance(120000 + self.bank), 0)
        self.assertEqual(self.s.aging("C", TODAY.isoformat()), aging_before)      # the invoices are open again
        self.controls_match()

    def test_issued_cheque(self):
        before, notes = self.party("S", self.supp), self.balance(2110)
        cid = self.cheque("OUT", 400)
        self.assertAlmostEqual(self.party("S", self.supp), before - 400, places=4)
        self.assertAlmostEqual(self.balance(2110), notes - 400, places=4)
        self.controls_match()
        self.status(cid, "COLLECTED", "OUT")
        self.assertAlmostEqual(self.balance(2110), notes, places=4)
        self.assertEqual(self.balance(120000 + self.bank), -400)
        self.controls_match()

    def test_issued_cheque_bounced(self):
        before = self.party("S", self.supp)
        cid = self.cheque("OUT", 400)
        self.status(cid, "BOUNCED", "OUT")
        self.assertAlmostEqual(self.party("S", self.supp), before, places=4)
        self.controls_match()

    def test_list(self):
        self.cheque("IN", 100)
        self.cheque("OUT", 50, due_in=20)
        rows = self.c.execute("SELECT DirectionName, PartyName, StatusName FROM ChequesQuery ORDER BY ChequeID").fetchall()
        self.assertEqual([r[0] for r in rows], ["وارد", "صادر"])
        self.assertTrue(all(r[1] for r in rows))
        self.assertEqual({r[2] for r in rows}, {"تحت التحصيل"})


class ChequeCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modCheque")

    def test_rules(self):
        problem = proc(self.text, "ChequeProblem")
        for part in ("لا يُسجَّل شيك على العميل النقدي", "تاريخ الاستحقاق قبل", "اختر البنك المسحوب عليه الشيك",
                     "هذا الشيك مسجل من قبل", "ClosedPeriodProblem(IssueDate)"):
            self.assertIn(part, problem)
        self.assertIn("MoveBalance Direction, CLng(PartyID), -Amount", proc(self.text, "PostCheque"))
        status = proc(self.text, "SetChequeStatus")
        for part in ("لا يُحصَّل قبل استحقاقه", 'If NewStatus = "BOUNCED" Then MoveBalance', "ClosedPeriodProblem(StatusDate)"):
            self.assertIn(part, status)
        self.assertIn("SourceType = 'CHEQUE_STATUS'", proc(self.text, "UndoChequeStatus"))
        self.assertIn("يُحذف الشيك تحت التحصيل فقط", proc(self.text, "DeleteCheque"))
        self.assertIn("qryJournalCheque", read("modJournal"))
        self.assertIn('"TestCheques"', read("modTestAll"))

    def test_schema_screen(self):
        self.assertIn((1250, "شيكات تحت التحصيل", "ASSET", 11, True, True), [tuple(r) for r in ACCOUNT_TREE])
        self.assertIn(("CHEQUE_STATUS", "تحصيل أو ارتداد شيك", 21), table("JournalSourceTypes").seed_rows)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmCheques"], "CHEQUES")
        names = {c.name for c in MODELS["frmCheques"].controls}
        for pname in re.findall(r"^Public Sub (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        self.assertIn("btnCheques", {c.name for c in MODELS["frmTreasury"].controls})


class Static_modCheque(VbaModuleChecks, unittest.TestCase):
    module_name = "modCheque"
    vba = read("modCheque")


if __name__ == "__main__":
    unittest.main()
