"""The VAT return (modVat, frmVatReturn, VatReturns, qryVatReturnTotals, VatReturnQuery, journal
sources VAT_RETURN / VAT_PAYMENT) on the SQLite mirror:
  * the boxes from the document lines agree with the journal: the VAT of box 1 is what the sales
    put in output VAT 2200, the VAT of box 7 what purchases and expenses put in input VAT 1500;
  * the filed return (same steps as FileVatReturn, tools/sim.py): a balanced settlement entry that
    empties 2200 and 1500 into the settlement account 2250, whose balance is the net VAT due; the
    payment empties 2250 from the bank; a refund is carried to the next return;
  * the form: 16 boxes, totals, box 16 = net due;
  * the VBA keeps the rules: filing after the period and in an open period, in order, unfiling
    the last one only with a reason and without payment, the bank account; the screen and permission."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import queries as Q
from schema import table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
ALL_TIME = {"PeriodStart": "2000-01-01 00:00:00", "PeriodEnd": "2100-01-01 00:00:00"}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class VatReturnTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.s = Store(self.db.con, now=day(0, 9))
        self.s.sync_journal()
        self.db.params.update(ALL_TIME)

    def totals(self):
        cur = self.db.con.execute("SELECT * FROM qryVatReturnTotals")
        return dict(zip([d[0] for d in cur.description], cur.fetchone()))

    def balance(self, account):
        return round(self.s.one("SELECT Coalesce(Sum(Debit) - Sum(Credit), 0) FROM JournalLines WHERE AccountCode = ?",
                                account), 4)

    def test_boxes_agree_with_the_journal(self):
        t = self.totals()
        self.assertAlmostEqual(t["SalesStdVAT"], -self.balance(2200), places=4)
        self.assertAlmostEqual(t["PurchStdVAT"], self.balance(1500), places=4)

    def test_filed_return_settles_into_2250_and_payment_clears_it(self):
        t = self.totals()
        vid, net = self.s.file_vat_return(t, "2000-01-01 00:00:00", day(1), day(0), ref="Z-1")
        entry = self.s.one("SELECT EntryID FROM JournalEntries WHERE SourceType = 'VAT_RETURN' AND SourceID = ?", vid)
        self.assertIsNotNone(entry)
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0)
        self.assertEqual(self.balance(2200), 0)
        self.assertEqual(self.balance(1500), 0)
        self.assertAlmostEqual(self.balance(2250), -float(net), places=4)
        self.assertAlmostEqual(float(net), t["SalesStdVAT"] - t["PurchStdVAT"], places=4)
        bank = self.balance(1200)
        self.s.pay_vat_return(vid, day(0), net)
        self.assertIsNotNone(self.s.one("SELECT EntryID FROM JournalEntries WHERE SourceType = 'VAT_PAYMENT' "
                                        "AND SourceID = ?", vid))
        self.assertEqual(self.balance(2250), 0)
        self.assertAlmostEqual(self.balance(1200), bank - float(net), places=4)

    def test_draft_makes_no_entry(self):
        t = self.totals()
        vid, _ = self.s.file_vat_return(t, "2000-01-01 00:00:00", day(1), day(0))
        self.db.con.execute("UPDATE VatReturns SET Status = 'DRAFT' WHERE VatReturnID = ?", (vid,))
        self.assertEqual(self.s.sync_journal(), (0, 0, 1))
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM JournalEntries WHERE SourceType LIKE 'VAT%'"), 0)

    def test_refund_is_carried_to_the_next_return(self):
        t = self.totals()
        due = t["SalesStdVAT"] - t["PurchStdVAT"]
        # corrections larger than the VAT of the period: a refund
        vid1, net1 = self.s.file_vat_return(t, "2000-01-01 00:00:00", day(40), day(39), corrections=-(due + 100))
        self.assertAlmostEqual(float(net1), -100, places=4)
        self.assertAlmostEqual(self.balance(2250), 100, places=4)             # a debit: refundable
        zero = {f: 0 for f in Store.VAT_FIELDS}
        zero.update(SalesStdVAT=300)
        vid2, net2 = self.s.file_vat_return(zero, day(38), day(1), day(0), carried=-net1)
        self.assertAlmostEqual(float(net2), 200, places=4)
        self.assertAlmostEqual(self.balance(2250), -200, places=4)            # the second return's net due

    def test_the_form(self):
        t = self.totals()
        vid, net = self.s.file_vat_return(t, "2000-01-01 00:00:00", day(1), day(0), corrections=10, carried=4)
        self.db.params["VatReturnID"] = vid
        rows = {r[0]: r for r in self.db.con.execute(
            "SELECT BoxNo, Amount, Adjust, VAT, RowKind FROM VatReturnQuery ORDER BY BoxNo").fetchall()}
        self.assertEqual(sorted(rows), list(range(1, 17)))
        self.assertEqual([b for b, *_ in Q.VAT_BOXES], list(range(1, 17)))
        self.assertAlmostEqual(rows[6][1], t["SalesStdAmount"] + t["SalesZeroAmount"] + t["SalesExemptAmount"], places=4)
        self.assertAlmostEqual(rows[12][3], t["PurchStdVAT"], places=4)
        self.assertAlmostEqual(rows[13][3], t["SalesStdVAT"] - t["PurchStdVAT"], places=4)
        self.assertAlmostEqual(rows[16][3], float(net), places=4)
        self.assertAlmostEqual(float(net), t["SalesStdVAT"] - t["PurchStdVAT"] + 10 - 4, places=4)
        self.assertIsNone(rows[16][1])
        self.assertEqual({r[4] for r in rows.values()}, {"L", "T", "N"})

    def test_categories(self):
        # a zero-rated sale line goes to box 3, an untaxed purchase line to box 10
        self.db.con.execute("UPDATE SalesInvoiceDetails SET VATCategory = 'Z', Tax = 0 WHERE SalesDetailID = "
                            "(SELECT Min(SalesDetailID) FROM SalesInvoiceDetails)")
        self.db.con.execute("UPDATE PurchaseInvoiceDetails SET VATRate = 0, Tax = 0 WHERE PurchaseDetailID = "
                            "(SELECT Min(PurchaseDetailID) FROM PurchaseInvoiceDetails)")
        t = self.totals()
        self.assertGreater(t["SalesZeroAmount"], 0)
        self.assertGreater(t["PurchZeroAmount"], 0)


class VatCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modVat")

    def test_captions_match_the_report(self):
        body = proc(self.text, "VatBoxCaption")
        vba = dict((int(n), c) for n, c in re.findall(r'Case (\d+): VatBoxCaption = "(.*)"', body))
        self.assertEqual(vba, {b: c for b, c, *_ in Q.VAT_BOXES})

    def test_rules(self):
        file_ = proc(self.text, "FileVatReturn")
        for part in ("If FiledDate <= periodTo Then", "ClosedPeriodProblem(FiledDate)", "الإقرارات تُعتمد بالترتيب",
                     "WriteVatDraft(", "SyncJournal()", "Status = 'DRAFT'"):
            self.assertIn(part, file_)
        unfile = proc(self.text, "UnfileVatReturn")
        for part in ("اكتب سبب إلغاء الاعتماد", "ألغِ السداد أولًا", "ClosedPeriodProblem(", "ألغِ اعتماد الإقرار اللاحق"):
            self.assertIn(part, unfile)
        pay = proc(self.text, "PayVatReturn")
        for part in ("Amount > due", "VatPaymentAccountProblem(", "ClosedPeriodProblem(PaidDate)"):
            self.assertIn(part, pay)
        self.assertIn("Nz(Level3Code, 0) <> 1100", proc(self.text, "VatPaymentAccountProblem"))
        self.assertIn("VatOverlap(", proc(self.text, "WriteVatDraft"))
        self.assertIn('CanVat(IIf(VatReturnID = 0, "ADD", "EDIT"))', proc(self.text, "SaveVatDraft"))
        journal = read("modJournal")
        self.assertIn("qryJournalVatReturn", journal)
        self.assertIn('Case "VAT_RETURN":       OpenScreen "frmVatReturn", 0, id', proc(journal, "OpenJournalSource"))
        self.assertIn("qryJournalVatReturn", Q.JOURNAL_SOURCE_QUERIES)

    def test_schema_permission_and_screen(self):
        rows = table("JournalSourceTypes").seed_rows
        self.assertIn(("VAT_RETURN", "تسوية إقرار ضريبة القيمة المضافة", 16), rows)
        self.assertIn(("VAT_PAYMENT", "سداد ضريبة القيمة المضافة", 17), rows)
        grants = [r for r in table("RolePermissions").seed_rows if r[1] == "VAT_RETURN"]
        self.assertEqual(grants, [(1, "VAT_RETURN"), (2, "VAT_RETURN")], "administrators and managers")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmVatReturn"], "VAT_RETURN")
        names = {c.name for c in MODELS["frmVatReturn"].controls}
        for pname in re.findall(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(ByVal frm As Access\.Form", self.text, re.M):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        self.assertIn("btnVat", {c.name for c in MODELS["frmFinancials"].controls})

    def test_in_access_tests(self):
        body = proc(read("modJournal"), "TestJournal")
        for part in ("اعتماد الإقرار ينشئ قيد التسوية", "قيد التسوية متوازن", "إلغاء الاعتماد والسداد يحذف القيدين"):
            self.assertIn(part, body)


class Static_modVat(VbaModuleChecks, unittest.TestCase):
    module_name = "modVat"
    vba = read("modVat")


if __name__ == "__main__":
    unittest.main()
