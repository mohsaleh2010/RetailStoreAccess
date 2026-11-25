"""Phase 7C (docs/87-Web-Completion.md): the rest of the web and the invoice and receipt of the rep app.
  * the C# of modPurchases.PostPurchaseReturn, of the bank reconciliation of modBank and of modManualEntry answers with the
    messages of Access, in the order of Access;
  * the web pages of the returns, the cashier closing, the bank reconciliation and the manual entries are work pages with the
    permission of their Access form, and their API operations exist;
  * the rep app makes its invoices and receipts on the program at once, as the rep of his open rep day (the van, the custody
    box), each with its ClientRef; they are never kept in the outbox.
The behaviour itself is tested in xUnit (web/tests/RetailStore.Tests/WebCompletionTests.cs), run by tests/test_web.py."""
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, HERE)

import schema


def read(rel):
    with open(os.path.join(ROOT, rel), encoding="utf-8-sig") as fh:
        return fh.read()


SERVICES = "web/src/RetailStore.Data/Services/"


def arabic_literals(code):
    code = re.sub(r"//[^\n]*", "", code)                           # the comments quote Arabic outputs
    return {m.replace("\\n", "") for m in re.findall(r'"([^"]*[؀-ۿ][^"]*)"', code)}


def in_order(test, text, marks):
    test.assertEqual([text.index(x) for x in marks], sorted(text.index(x) for x in marks), marks)


class MessagesTests(unittest.TestCase):

    def test_the_bank_reconciliation_speaks_as_mod_bank(self):
        vba = read("src/vba/modBank.bas")
        messages = arabic_literals(read(SERVICES + "BankReconService.cs"))
        self.assertGreater(len(messages), 15)
        for m in messages:
            with self.subTest(m):
                self.assertIn(m, vba)

    def test_the_manual_entries_speak_as_mod_manual_entry(self):
        vba = "".join(read(f"src/vba/{m}.bas") for m in ("modManualEntry", "modClosing", "modCurrency"))
        messages = arabic_literals(read(SERVICES + "ManualEntryService.cs"))
        self.assertGreater(len(messages), 15)
        for m in messages:
            with self.subTest(m):
                self.assertIn(m, vba)

    def test_the_documents_of_the_rep_speak_as_access(self):
        code = read(SERVICES + "RepDocService.cs")
        own = {re.search(r'NoOpenDay = "([^"]+)"', code).group(1)}
        known = "".join(read(f"src/vba/{m}.bas") for m in ("modSales", "modReports")) + read(SERVICES + "PosService.cs") + read(
            "web/src/RetailStore.UI/Services/WorkPages.cs") + read("tools/queries.py")
        for m in arabic_literals(code) - own:
            with self.subTest(m):
                self.assertIn(m, known)


class OrderTests(unittest.TestCase):

    def test_the_purchase_return_follows_post_purchase_return(self):
        code = read(SERVICES + "PurchaseService.cs")
        ret = code[code.index("public async Task<SaleResult> PostPurchaseReturnAsync"):code.index("PurchaseReturnsAsync(")]
        in_order(self, ret, ["\"frmPurchaseReturn\"", "ScreenAction.Add", "\"PURCHASE_RETURN\")", "سبب الإرجاع مطلوب.", "\"CREDIT\"",
                             "فاتورة الشراء الأصلية غير موجودة.", "سطر لا ينتمي إلى فاتورة الشراء الأصلية.", "PurchaseUnitCost(",
                             "InvoiceCalculator.ReturnAmounts(", "حدد الكمية المرتجعة", "\"الإرجاع للمورد\"", "BeginTransactionAsync",
                             "NextNumberAsync(db, \"PURCHASE_RETURN\")", "PurchaseReturnDetails.Add", "ApplyStockMovementAsync",
                             "AdjustSupplierBalanceAsync", "\"PURCHASE_RETURN\", \"PurchaseReturns\"", "CommitAsync"])
        self.assertIn("-(l.Qty * l.Factor), TtPurchaseReturn, l.Cost", ret, "the pieces leave at the cost they came in with (docs/72)")
        self.assertIn("TT_PURCHASE_RETURN As Long = 3", read("src/vba/modSales.bas"))
        self.assertIn("TtPurchaseReturn = 3", code)

    def test_the_manual_entry_follows_post_manual_entry(self):
        code = read(SERVICES + "ManualEntryService.cs")
        save = code[code.index("public async Task<TreasuryResult> SaveAsync"):code.index("public async Task<TreasuryResult> DeleteAsync")]
        in_order(self, save, ["OpenProblemAsync", "\"MANUAL_ENTRY\")", "ProblemAsync(db", "CurrencyProblemAsync", "ClosedPeriodProblem(e.EntryDate",
                              "ClosedPeriodProblem(h.EntryDate", "BeginTransactionAsync", "NextNumberAsync(db, \"MANUAL_ENTRY\")",
                              "the conversion may leave a cent", "MANUAL_ENTRY_EDIT", "CommitAsync", "journal.SyncAsync"])
        delete = code[code.index("public async Task<TreasuryResult> DeleteAsync"):code.index("public async Task<ManualEntryView?> EntryAsync")]
        in_order(self, delete, ["ScreenAction.Delete", "\"MANUAL_ENTRY\")", "NotFound", "ClosedPeriodProblem", "Audit.Snapshot",
                                "MANUAL_ENTRY_DELETE", "journal.SyncAsync"])
        problem = code[code.index("public static async Task<string> ProblemAsync"):code.index("public async Task<TreasuryResult> SaveAsync")]
        vba = read("src/vba/modManualEntry.bas")
        vba_problem = vba[vba.index("Public Function ManualEntryProblem"):vba.index("Public Function PostManualEntry")]
        order = ["اكتب تاريخ القيد.", "اكتب بيان القيد.", "اختر الحساب في كل سطر فيه مبلغ.", "المبالغ لا تكون سالبة",
                 "السطر يكون مدينًا أو دائنًا", "اكتب المبلغ مدينًا أو دائنًا", "حساب رئيسي أو معطّل", "القيد سطران على الأقل",
                 "القيد غير متوازن"]
        in_order(self, problem, order)
        in_order(self, vba_problem, order)

    def test_the_reconciliation_follows_mod_bank(self):
        code = read(SERVICES + "BankReconService.cs")
        vba = read("src/vba/modBank.bas")
        for csharp, basic, marks in [
            ("OpenAsync", "OpenReconciliation", ["لا تملك صلاحية البنوك.", "اختر البنك.", "اكتب تاريخ كشف البنك.", "تاريخ الكشف بعد اليوم.",
                                                 "آخر تسوية معتمدة لهذا البنك في ", "لا تملك صلاحية الإضافة في شاشة التسوية البنكية."]),
            ("FinishAsync", "FinishReconciliation", ["NotFoundOrDone", "لا تملك صلاحية اعتماد التسوية البنكية.", "يوجد فرق "]),
            ("ReopenAsync", "ReopenReconciliation", ["التسوية غير موجودة.", "لا تملك صلاحية تعديل التسوية البنكية.",
                                                     "توجد تسوية جارية لهذا البنك", "تُعاد فتح آخر تسوية للبنك فقط."]),
        ]:
            with self.subTest(csharp):
                start = code.index(f"public async Task<{'TreasuryResult' if csharp == 'OpenAsync' else 'string'}> {csharp}")
                body = code[start:code.index("\n    }\n", start)]
                in_order(self, body, marks)
                b = vba[vba.index(f"Public Function {basic}"):]
                b = b[:b.index("End Function")]
                in_order(self, b, [m for m in marks if m != "NotFoundOrDone"])
        self.assertIn('"BANK_RECON_DELETE"', code)
        self.assertIn('"BANK_RECON_REOPEN"', code)
        self.assertIn('NextNumberAsync(db, "BANK_RECON")', code)


class PagesAndApiTests(unittest.TestCase):

    PAGES = {"sales-returns": "frmSalesReturn", "purchase-returns": "frmPurchaseReturn", "cash-closing": "frmCashClosing",
             "bank-reconciliation": "frmBankRecon", "manual-entries": "frmManualEntry"}

    def test_every_page_is_a_work_page_with_the_permission_of_its_form(self):
        work = read("web/src/RetailStore.UI/Services/WorkPages.cs")
        screens = {s[0] for s in schema.SCREEN_LIST}
        pages = "".join(read(f"web/src/RetailStore.UI/Pages/Office/{f}") for f in os.listdir(os.path.join(ROOT, "web/src/RetailStore.UI/Pages/Office")))
        for route, form in self.PAGES.items():
            with self.subTest(route):
                self.assertIn(f'new("{route}", "{form}"', work)
                self.assertIn(form, screens)
                self.assertIn(f'@page "/{route}"', pages)
        self.assertEqual(pages.count("OpenProblemAsync"), len(self.PAGES))

    def test_the_operations_of_the_api(self):
        api = read("web/src/RetailStore.Web/Api/OfficeApi.cs") + read("web/src/RetailStore.Web/Api/RepApi.cs")
        for route in ("/sales-returns", "/purchase-returns", "/cash-closings", "/bank-reconciliations", "/manual-entries"):
            with self.subTest(route):
                self.assertIn(f'MapGroup("{route}")', api)
        for name in ("PostSalesReturn", "PostPurchaseReturn", "PostCashClosing", "StartBankReconciliation", "FinishBankReconciliation",
                     "AddManualEntry", "ChangeManualEntry", "DeleteManualEntry", "PostRepSale", "PostRepPayment"):
            with self.subTest(name):
                self.assertIn(f'WithName("{name}")', api)
        self.assertIn("OfficeApi.Map(secured);", read("web/src/RetailStore.Web/Api/ApiEndpoints.cs"))


class RepAppTests(unittest.TestCase):

    def test_the_invoice_and_the_receipt_are_made_as_the_rep_of_his_open_day(self):
        docs = read(SERVICES + "RepDocService.cs")
        self.assertIn('d.Status == "OPEN"', docs)
        self.assertIn("user with { RepDay = day }", docs)
        sales = read(SERVICES + "SalesService.cs")
        for mark in ("user.RepDay?.WarehouseID ?? await RepService.UserWarehouseAsync",
                     "user.RepDay?.SalesRepID ?? await RepService.SalesRepForAsync(db, sale.CustomerID",
                     "user.RepDay?.SalesRepID ?? await RepService.SalesRepForAsync(db, pay.CustomerID",
                     "return day.CashBoxID;", "ClientRef = device?.ClientRef ?? clientRef"):
            with self.subTest(mark):
                self.assertIn(mark, sales)
        # each with its number on the phone, checked before the sale (the same user = the existing document)
        self.assertLess(docs.index("s.ClientRef == reference"), docs.index("Sales.PostSaleAsync"))
        self.assertLess(docs.index("p.ClientRef == reference"), docs.index("Sales.PostPaymentAsync"))

    def test_they_are_never_kept_on_the_phone(self):
        session = read("web/src/RetailStore.RepApp/Offline/RepSession.cs")
        sell = session[session.index("public async Task<DocAnswer> SellAsync"):session.index("public async Task<string> PrintDocumentAsync")]
        self.assertNotIn("Enqueue", sell)
        self.assertIn("api.PostSaleAsync", sell)
        self.assertIn("api.PostPaymentAsync", sell)
        kinds = read("web/src/RetailStore.RepApp/Offline/LocalStore.cs")
        self.assertNotIn('"SALE"', kinds)
        self.assertNotIn('"PAYMENT"', kinds)
        self.assertNotIn("RetailStore.Data", read("web/src/RetailStore.RepApp/RetailStore.RepApp.csproj"))


if __name__ == "__main__":
    unittest.main()
