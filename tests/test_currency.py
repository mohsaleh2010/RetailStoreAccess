"""Several currencies (modCurrency, Currencies, CurrencyRates, the currency fields of the documents and of
JournalEntries, SupplierFxBalanceQuery): the amounts stay in SAR, the document keeps its currency, rate and
foreign total. The balances by currency run on the SQLite mirror, the conversion runs in LibreOffice, the
posting code and the screens are checked."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import AccessOnSqlite, day
import forms as F
import vba_harness as H
from relations import index_load, ACCESS_MAX_INDEXES
from schema import TABLES, table
from sim import Store

MODELS = {m.name: m for m in F.all_forms()}
FX_TABLES = ("PurchaseInvoices", "PurchaseReturns", "SupplierPayments", "CustomerPayments", "Expenses",
             "ManualEntries", "JournalEntries")


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


class CurrencySchemaTests(unittest.TestCase):

    def test_tables_and_document_fields(self):
        codes = [r[0] for r in table("Currencies").seed_rows]
        self.assertEqual(codes[0], "SAR")
        for code in ("USD", "EUR", "AED", "KWD"):
            self.assertIn(code, codes)
        rates = {r[1]: r[3] for r in table("CurrencyRates").seed_rows}
        self.assertEqual(rates["USD"], 3.75)
        settings = {f.name: f for f in table("Settings").fields}
        self.assertEqual(settings["CurrencyCode"].default, '"SAR"')           # the program currency
        for t in FX_TABLES:
            fields = {f.name: f for f in table(t).fields}
            self.assertEqual(fields["CurrencyCode"].fk, "Currencies.CurrencyCode", t)
            self.assertEqual(fields["CurrencyCode"].default, '"SAR"', t)     # old documents: SAR
            self.assertEqual(fields["ExchangeRate"].default, "1", t)
            self.assertEqual(fields["ExchangeRate"].rule, ">0", t)
            self.assertIn("ForeignAmount", fields)
        self.assertIn("ForeignTax", {f.name for f in table("Expenses").fields})
        self.assertLessEqual({"ForeignDebit", "ForeignCredit"}, {f.name for f in table("ManualEntryLines").fields})
        self.assertEqual(next(f.fk for f in table("Suppliers").fields if f.name == "CurrencyCode"),
                         "Currencies.CurrencyCode")
        self.assertLessEqual(index_load("Currencies"), ACCESS_MAX_INDEXES - 4)


class CurrencyBalanceTests(unittest.TestCase):

    def setUp(self):
        self.db = AccessOnSqlite()
        self.db.load_fixture()
        self.c = self.db.con
        self.s = Store(self.c, now=day(0, 9))
        self.supplier = self.s.one("SELECT Min(SupplierID) FROM Suppliers")

    def fx_row(self):
        cur = self.c.execute("SELECT * FROM SupplierFxBalanceQuery WHERE SupplierID = ? AND CurrencyCode = 'USD'",
                             (self.supplier,))
        row = cur.fetchone()
        return dict(zip([d[0] for d in cur.description], row)) if row else None

    def test_old_documents_are_sar(self):
        self.assertEqual(self.s.one("SELECT COUNT(*) FROM SupplierFxBalanceQuery"), 0)

    def test_a_dollar_payment_and_its_revaluation(self):
        self.s.insert("SupplierPayments", PaymentNumber="TEST-FX", SupplierID=self.supplier, PaymentDate=day(1, 10),
                      Amount=375, PaymentMethodID=1, EmployeeID=1, CurrencyCode="USD", ExchangeRate=3.75,
                      ForeignAmount=100)
        r = self.fx_row()
        self.assertEqual((r["FxBalance"], r["BookBalance"], r["LastRate"]), (-100, -375, 3.75))
        self.assertEqual((r["RevaluedBalance"], r["FxDifference"]), (-375, 0))
        self.s.insert("CurrencyRates", CurrencyCode="USD", RateDate=day(0), Rate=3.8)
        r = self.fx_row()
        self.assertEqual((r["LastRate"], r["RevaluedBalance"], r["FxDifference"]), (3.8, -380, -5))


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class ConversionRuntimeTests(unittest.TestCase):

    def test_to_and_from_sar(self):
        h = H.Harness()
        try:
            # a Currency value does not come back through UNO: wrappers return a Double
            wrap = ("Option Explicit\n"
                    "Public Function WTo(ByVal a As Double, ByVal r As Double) As Double\n"
                    "    WTo = CDbl(ToBase(CCur(a), r))\nEnd Function\n"
                    "Public Function WFrom(ByVal a As Double, ByVal r As Double) As Double\n"
                    "    WFrom = CDbl(FromBase(CCur(a), r))\nEnd Function\n")
            h.load({"modCurrency": H.read_module("modCurrency"), "modWrap": wrap,
                    "modRound": "Option Explicit\n" + H.HARNESS_ROUND})      # RoundMoney of modCommon
            for amount, rate, sar in [(100, 3.75, 375), (0.1, 3.75, 0.38), (1234.56, 1.0211, 1260.61), (99.99, 1, 99.99)]:
                self.assertAlmostEqual(h.call("modWrap", "WTo", float(amount), float(rate)), sar, places=2)
            self.assertAlmostEqual(h.call("modWrap", "WFrom", 375.0, 3.75), 100, places=2)
            self.assertAlmostEqual(h.call("modWrap", "WFrom", 375.0, 0.0), 375, places=2)   # no rate: unchanged
        finally:
            h.close()


class CurrencyCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modCurrency")

    def test_posting_keeps_sar_and_the_document_currency(self):
        for module, name, parts in [
                ("modPurchases", "PostPurchaseFromCart",
                 ("CurrencyProblem(CurrencyCode, FxRate)", "CCur(CDbl(Nz(rs!UnitCost, 0)) * FxRate)",
                  "ToBase(Nz(rs!LineDiscount, 0), FxRate)", "tend = ToBase(CCur(PaidAmount), FxRate)",
                  "rs!ForeignAmount = FromBase(CalcTotal(\"TOTAL\"), FxRate)")),
                ("modPurchases", "PostSupplierPayment", ("Amount = ToBase(foreign, FxRate)", "rs!ForeignAmount = foreign")),
                ("modSales", "PostCustomerPayment", ("Amount = ToBase(foreign, FxRate)", "rs!ForeignAmount = foreign")),
                ("modPurchases", "PostPurchaseReturn", ("rs!ExchangeRate = Nz(DbValue(\"SELECT ExchangeRate FROM PurchaseInvoices",)),
                ("modManualEntry", "PostManualEntry", ("h!Debit = ToBase(Nz(rs!Debit, 0), FxRate)", "h!ForeignDebit =",
                                                       "diff = total - credits"))]:
            body = proc(read(module), name)
            for part in parts:
                self.assertIn(part, body, f"{name}: {part}")
            if name != "PostPurchaseReturn":
                self.assertIn("If IsBaseCurrency(CurrencyCode) Then FxRate = 1", body, name)

    def test_journal_entries_carry_the_currency(self):
        journal = proc(read("modJournal"), "SyncJournal")
        self.assertLess(journal.index("StampJournalCurrencies"), journal.index("ws.CommitTrans"))
        stamp = proc(self.text, "StampJournalCurrencies")
        sources = {r[0] for r in table("JournalSourceTypes").seed_rows}
        for kind, tbl, key in re.findall(r'"(\w+)\|(\w+)\|(\w+)"', stamp):
            self.assertIn(kind, sources)
            self.assertIn(key, {f.name for f in table(tbl).fields})
            self.assertIn("ForeignAmount", {f.name for f in table(tbl).fields})
        self.assertEqual(len(re.findall(r'"(\w+)\|(\w+)\|(\w+)"', stamp)), 6)

    def test_screens(self):
        for form, save, module in [("frmPurchaseInvoice", "SavePurchase", "modPurchaseScreens"),
                                   ("frmSupplierPayment", "SaveSupplierPayment", "modPurchaseScreens"),
                                   ("frmCustomerPayment", "SavePayment", "modPOS"),
                                   ("frmManualEntry", "SaveManualEntry", "modManualEntry")]:
            names = {c.name for c in MODELS[form].controls}
            self.assertLessEqual({"cboCurrency", "txtRate"}, names, form)
            self.assertIn("CurrencyChoice(frm, code, fx)", proc(read(module), save), save)
        expense = {c.name for c in MODELS["frmExpenses"].controls}
        self.assertLessEqual({"CurrencyCode", "ExchangeRate", "ForeignAmount", "ForeignTax"}, expense)
        self.assertIn("If Not ExpenseCurrencyOK(frm) Then Exit Function", read("modForms"))
        for form in ("frmCurrencies", "frmCurrencyRates"):
            self.assertEqual(F.SCREEN_PERMISSIONS[form], "CURRENCIES")
        self.assertIn("SUPPLIER_FX", {r.key for r in F.REPORTS})
        self.assertIn('"TestCurrency"', read("modTestAll"))


class Static_modCurrency(VbaModuleChecks, unittest.TestCase):
    module_name = "modCurrency"
    vba = read("modCurrency")


if __name__ == "__main__":
    unittest.main()
