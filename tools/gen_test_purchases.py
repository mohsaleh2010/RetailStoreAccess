"""Generate modTestPurchases.bas: the in-Access tests of Phase 7.
Expected numbers come from tools/purchases_reference.py (the same values are
checked on the SQLite mirror by tests/test_purchases.py)."""

import re

import pricing as P
import purchases_reference as R
from generate_common import vba_str


def num(v) -> str:
    return R.vba_number(v)


def vector_lines():
    out = []
    for label, cq, ca, q, c in P.AVERAGE_CASES:
        exp = P.weighted_average(cq, ca, q, c)
        out.append(f"    Call Record(WeightedAverage({cq}, {ca}, {q}, {c}) = CCur({num(exp)}), "
                   f"{vba_str('متوسط التكلفة: ' + label)})")
    for net, total, qty, reg in [("100", "115", "3", True), ("100", "115", "3", False), ("0", "0", "0", True)]:
        exp = P.purchase_unit_cost(net, total, qty, reg)
        label = f"تكلفة الوحدة {net}/{total} ÷ {qty} " + ("مسجل بالضريبة" if reg else "غير مسجل")
        out.append(f"    Call Record(PurchaseUnitCost({net}, {total}, {qty}, {'True' if reg else 'False'}) = "
                   f"CCur({num(exp)}), {vba_str(label)})")
    return "\n".join(out)


TEMPLATE = r'''Attribute VB_Name = "modTestPurchases"
'==============================================================================
' modTestPurchases  -  Retail Store Management System (Phase 7 tests)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_test_purchases.py
'
'   TestPurchases  1) average-cost vectors (same as tools/pricing.py)
'                  2) the purchase and inventory cycle inside a transaction that
'                     is rolled back: new supplier -> buy 100 -> stock 100 ->
'                     sell 10 -> stock 90 (spec), second purchase at a new cost,
'                     purchase return, payment voucher, opening balance, stock-out,
'                     stocktake, and every refusal (duplicate supplier invoice,
'                     unpaid cash purchase, future date, over-return, ...)
'                  3) the purchase and inventory screens (nothing is posted)
'==============================================================================
Option Compare Database
Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_report As String

Public Function TestPurchases() As Boolean
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestPurchases  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    g_AutoAnswer = True
    EnsureLocalTables

    TestAverageVectors
    TestCycle
    TestPurchaseScreens

    g_SilentMode = False
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        MsgBox "جميع اختبارات المشتريات والمخزون ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "لم تُترك أي بيانات اختبار.", vbInformation + MSG_RTL, "TestPurchases"
        TestPurchases = True
    Else
        MsgBox "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestPurchases"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestAverageVectors()
@@VECTORS@@
End Sub

'------------------------------------------------------------------------------
' 2) the purchase and inventory cycle, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestCycle()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, integrityBefore As Long
    Dim sid As Long, cat As Long, pa As Long, pb As Long, msg As String, inv1 As Long, inv2 As Long
    Dim saleID As Long, dummy As Long, retID As Long, payID As Long, refNo As String
    Dim countID As Long, count2 As Long, adjusted As Long, netValue As Currency
    On Error GoTo EH
    Set db = CurrentDb
    integrityBefore = DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery")
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "UPDATE Settings SET VATNumber = '300000000000003', VATRate = 0.15, PricesIncludeVAT = True, " & _
               "AllowNegativeStock = False", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    ' an open stocktake of the user blocks a new one: closed here, restored by the rollback
    db.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE Status = 'OPEN'", dbFailOnError

    ' --- 1. new supplier and products
    db.Execute "INSERT INTO Suppliers (SupplierName, VATNumber) VALUES ('TEST مورد المشتريات', '300000000000003')", _
               dbFailOnError
    sid = DbValue("SELECT SupplierID FROM Suppliers WHERE SupplierName = 'TEST مورد المشتريات'")
    db.Execute "INSERT INTO Categories (CategoryName) VALUES ('TEST تصنيف الجرد')", dbFailOnError
    cat = DbValue("SELECT CategoryID FROM Categories WHERE CategoryName = 'TEST تصنيف الجرد'")
    db.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, CategoryID, SellingPrice, MinimumQuantity) " & _
               "VALUES ('TEST-PUR-A', 'TEST منتج أ', 'TESTPUR000A', " & cat & ", @@SALE_PRICE@@, 5)", dbFailOnError
    pa = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PUR-A'")
    db.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, SellingPrice) " & _
               "VALUES ('TEST-PUR-B', 'TEST منتج ب', " & cat & ", 23)", dbFailOnError
    pb = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PUR-B'")
    Call Record(Stock(pa) = 0, "منتج جديد بدون رصيد")

    ' --- 2. purchase 1: 100 x 50 + VAT, credit, 1000 paid now
    PurCartAdd pa, @@P1_QTY@@, @@P1_COST@@
    msg = PostPurchaseFromCart(sid, "TEST-S-1", Date, "CREDIT", 1, True, 0, @@P1_PAID@@, "", inv1)
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(msg) = 0 And inv1 > 0, "حفظ فاتورة شراء آجلة " & msg)
    Call Record(PurInv(inv1, "TotalAmount") = @@P1_TOTAL@@ And PurInv(inv1, "Tax") = @@P1_TAX@@ And _
                PurInv(inv1, "RemainingAmount") = @@P1_REMAINING@@, "الإجمالي @@P1_TOTAL@@ والضريبة @@P1_TAX@@ والمتبقي @@P1_REMAINING@@")
    Call Record(Stock(pa) = @@STOCK_1@@, "بعد شراء 100 أصبح المخزون 100")
    Call Record(AvgCost(pa) = @@AVG_1@@, "متوسط التكلفة @@AVG_1@@ (بدون الضريبة القابلة للاسترداد)")
    Call Record(SupBalance(sid) = @@P1_REMAINING@@, "رصيد المورد = المتبقي من الفاتورة")
    Call Record(DbValue("SELECT PurchasePrice FROM Products WHERE ProductID = " & pa) = @@P1_COST@@ And _
                DbValue("SELECT SupplierID FROM Products WHERE ProductID = " & pa) = sid, "آخر سعر شراء والمورد الافتراضي")
    Call Record(Nz(DbValue("SELECT QuantityAfter FROM InventoryTransactions WHERE ReferenceType = 'PURCHASE' " & _
                "AND ReferenceID = " & inv1), 0) = @@STOCK_1@@, "حركة مخزون شراء بالرصيد بعدها")

    ' --- refusals (nothing may change)
    PurCartAdd pa, 1, @@P1_COST@@
    Call Record(Len(PostPurchaseFromCart(sid, "TEST-S-1", Date, "CREDIT", 1, True, 0, 0, "", dummy)) > 0, _
                "رفض تكرار رقم فاتورة المورد لنفس المورد")
    Call Record(Len(PostPurchaseFromCart(sid, "", Date, "CASH", 1, True, 0, 10, "", dummy)) > 0, _
                "رفض شراء نقدي بمبلغ مدفوع أقل من الإجمالي")
    Call Record(Len(PostPurchaseFromCart(sid, "", Date + 1, "CASH", 1, True, 0, Null, "", dummy)) > 0, _
                "رفض تاريخ فاتورة في المستقبل")
    Call Record(Len(PostPurchaseFromCart(0, "", Date, "CASH", 1, True, 0, Null, "", dummy)) > 0, "رفض فاتورة بدون مورد")
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(PostPurchaseFromCart(sid, "", Date, "CASH", 1, True, 0, Null, "", dummy)) > 0, _
                "رفض فاتورة بدون أصناف")
    Call Record(Stock(pa) = @@STOCK_1@@ And SupBalance(sid) = @@P1_REMAINING@@, "الرفض لم يغير المخزون ولا الرصيد")

    ' --- 3. the spec scenario: sell 10 -> 90
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, @@SALE_QTY@@, SellingPrice, 0, " & _
        "CurrentQuantity FROM Products WHERE ProductID = " & pa, dbFailOnError
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", saleID)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Len(msg) = 0 And Stock(pa) = @@STOCK_AFTER_SALE@@, "بيع 10: المخزون 90 " & msg)
    Call Record(Nz(DbValue("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & saleID), 0) = @@AVG_1@@, _
                "تكلفة البيع = متوسط تكلفة الشراء")

    ' --- 4. purchase 2: 10 x 60, cash, new selling price
    PurCartAdd pa, @@P2_QTY@@, @@P2_COST@@
    db.Execute "UPDATE tmpPurchaseLines SET NewSellingPrice = @@P2_NEW_PRICE@@", dbFailOnError
    msg = PostPurchaseFromCart(sid, "TEST-S-2", Date, "CASH", 1, True, 0, Null, "", inv2)
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    Call Record(Len(msg) = 0 And PurInv(inv2, "TotalAmount") = @@P2_TOTAL@@ And PurInv(inv2, "RemainingAmount") = 0, _
                "شراء نقدي @@P2_TOTAL@@ مدفوع بالكامل " & msg)
    Call Record(Stock(pa) = @@STOCK_2@@ And AvgCost(pa) = @@AVG_2@@, "المخزون 100 ومتوسط التكلفة المرجح @@AVG_2@@")
    Call Record(DbValue("SELECT SellingPrice FROM Products WHERE ProductID = " & pa) = @@P2_NEW_PRICE@@, _
                "تحديث سعر البيع من فاتورة الشراء")
    Call Record(SupBalance(sid) = @@P1_REMAINING@@, "الشراء النقدي لا يغير رصيد المورد")

    ' --- 5. purchase return: 10 units of purchase 1, deducted from the account
    db.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ReturnQty) SELECT PurchaseDetailID, " & _
               "ProductID, @@RETURN_QTY@@ FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " & inv1, dbFailOnError
    Call Record(Len(PostPurchaseReturn(inv1, "", "CREDIT", Null, dummy)) > 0, "رفض مرتجع بدون سبب")
    msg = PostPurchaseReturn(inv1, "TEST عيب مصنعي", "CREDIT", Null, retID)
    Call Record(Len(msg) = 0 And retID > 0, "حفظ مرتجع المشتريات " & msg)
    Call Record(Nz(DbValue("SELECT TotalAmount FROM PurchaseReturns WHERE PurchaseReturnID = " & retID), 0) = @@RETURN_TOTAL@@ And _
                Nz(DbValue("SELECT Tax FROM PurchaseReturns WHERE PurchaseReturnID = " & retID), 0) = @@RETURN_TAX@@, _
                "قيمة المرتجع @@RETURN_TOTAL@@ منها ضريبة @@RETURN_TAX@@")
    Call Record(Stock(pa) = @@STOCK_3@@ And AvgCost(pa) = @@AVG_3@@, "المرتجع أخرج 10 بتكلفة شرائها: المتوسط @@AVG_3@@")
    Call Record(SupBalance(sid) = @@BALANCE_AFTER_RETURN@@, "المرتجع خُصم من رصيد المورد: @@BALANCE_AFTER_RETURN@@")
    db.Execute "UPDATE tmpPurchaseReturnLines SET ReturnQty = 95", dbFailOnError
    Call Record(Len(PostPurchaseReturn(inv1, "TEST", "CREDIT", Null, dummy)) > 0 And Stock(pa) = @@STOCK_3@@, _
                "رفض إرجاع أكثر من المتبقي من الفاتورة")
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError

    ' --- 6. payment voucher
    Call Record(Len(PostSupplierPayment(sid, 0, 1, "", dummy)) > 0, "رفض سند صرف بمبلغ صفر")
    msg = PostSupplierPayment(sid, @@PAYMENT@@, 1, "TEST", payID)
    Call Record(Len(msg) = 0 And SupBalance(sid) = @@BALANCE_FINAL@@, "سند صرف @@PAYMENT@@: الرصيد @@BALANCE_FINAL@@ " & msg)

    ' --- 7. manual movements on product B
    Call Record(Len(PostManualStock(pa, TT_OPENING, 5, 10, "", refNo)) > 0, "رفض رصيد افتتاحي لمنتج له حركات")
    msg = PostManualStock(pb, TT_OPENING, @@B_OPENING_QTY@@, @@B_OPENING_COST@@, "", refNo)
    Call Record(Len(msg) = 0 And Len(refNo) > 0 And Stock(pb) = @@B_OPENING_QTY@@ And AvgCost(pb) = @@B_AVG@@, _
                "رصيد افتتاحي 20 بتكلفة 10 " & msg)
    Call Record(Len(PostManualStock(pb, TT_STOCK_IN, 1, 10, "", refNo)) > 0, "رفض إضافة مخزون بدون سبب")
    Call Record(Len(PostManualStock(pb, TT_SALE, 1, Null, "TEST", refNo)) > 0, "رفض نوع حركة غير يدوي")
    msg = PostManualStock(pb, TT_STOCK_OUT, @@B_OUT_QTY@@, Null, "TEST تالف", refNo)
    Call Record(Len(msg) = 0 And Stock(pb) = @@B_STOCK@@, "خصم تالف 2: المخزون 18 " & msg)
    g_AutoAnswer = False
    Call Record(Len(PostManualStock(pb, TT_STOCK_OUT, 100, Null, "TEST", refNo)) > 0 And Stock(pb) = @@B_STOCK@@, _
                "منع خصم أكثر من المتوفر بدون موافقة المدير")
    g_AutoAnswer = True

    ' --- 8. stocktake of the test category
    msg = CreateStockCount(cat, "TEST", countID)
    Call Record(Len(msg) = 0 And countID > 0, "إنشاء جرد للتصنيف " & msg)
    Call Record(DbValue("SELECT COUNT(*) FROM StockCountDetails WHERE StockCountID = " & countID) = 2 And _
                DbValue("SELECT SystemQuantity FROM StockCountDetails WHERE StockCountID = " & countID & _
                " AND ProductID = " & pa) = @@STOCK_3@@, "الجرد يضم منتجي التصنيف بالكمية المسجلة")
    Call Record(Len(CreateStockCount(cat, "", dummy)) > 0, "منع فتح جردين في نفس الوقت")
    db.Execute "UPDATE StockCountDetails SET ActualQuantity = @@COUNT_A@@ WHERE StockCountID = " & countID & _
               " AND ProductID = " & pa, dbFailOnError
    db.Execute "UPDATE StockCountDetails SET ActualQuantity = @@COUNT_B@@ WHERE StockCountID = " & countID & _
               " AND ProductID = " & pb, dbFailOnError
    msg = PostStockCount(countID, False, adjusted, netValue)
    Call Record(Len(msg) = 0 And adjusted = @@COUNT_LINES@@ And netValue = @@COUNT_VALUE@@, _
                "ترحيل الجرد: تسوية واحدة بقيمة @@COUNT_VALUE@@ " & msg)
    Call Record(Stock(pa) = @@STOCK_FINAL_A@@ And Stock(pb) = @@B_STOCK@@, "المخزون بعد الجرد 88 و 18")
    Call Record(Nz(DbValue("SELECT Quantity FROM InventoryTransactions WHERE ReferenceType = 'STOCK_COUNT' " & _
                "AND ReferenceID = " & countID), 0) = @@COUNT_DIFF_A@@, "حركة تسوية جرد بالفرق")
    Call Record(DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & countID) = "POSTED", "حالة الجرد مُرحّل")
    Call Record(Len(PostStockCount(countID, False, adjusted, netValue)) > 0, "رفض ترحيل جرد مُرحّل مرة ثانية")
    Call Record(Len(CancelStockCount(countID)) > 0, "رفض إلغاء جرد مُرحّل")
    msg = CreateStockCount(cat, "", count2)
    Call Record(Len(msg) = 0 And Len(CancelStockCount(count2)) = 0 And _
                DbValue("SELECT Status FROM StockCounts WHERE StockCountID = " & count2) = "CANCELLED" And _
                Stock(pa) = @@STOCK_FINAL_A@@, "إلغاء جرد مفتوح لا يغير المخزون")

    ' --- the saved queries agree with the posted documents
    Call Record(DbValue("SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = " & sid) = @@BALANCE_FINAL@@, _
                "رصيد المورد من كشف الحساب = @@BALANCE_FINAL@@")
    Call Record(DbValue("SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0 AND ProductID IN (" & _
                pa & ", " & pb & ")") = 0, "الكميات تطابق دفتر حركة المخزون")
    Call Record(DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = integrityBefore, "فحص السلامة لا يجد أي مشكلة جديدة")

    ws.Rollback
    inTrans = False
    db.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    db.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(DbValue("SELECT COUNT(*) FROM Suppliers WHERE SupplierName = 'TEST مورد المشتريات'") = 0, _
                "التراجع عن كل بيانات الاختبار")
    Exit Sub
EH:
    Fail "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
End Sub

'------------------------------------------------------------------------------
' 3) the screens (nothing is posted)
'------------------------------------------------------------------------------
Private Sub TestPurchaseScreens()
    Dim frm As Access.Form, pid As Long
    On Error GoTo EH
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, PurchasePrice, SellingPrice) VALUES " & _
                      "('TEST-PUR-C', 'TEST منتج شاشة المشتريات', 'TESTPUR000C', 40, 55)", dbFailOnError
    pid = DLookup("ProductID", "Products", "ProductCode = 'TEST-PUR-C'")
    DoCmd.OpenForm "frmPurchaseInvoice", acNormal, , , , acHidden
    Set frm = Forms("frmPurchaseInvoice")
    Call Record(PurScanCode(frm, "TESTPUR000C") And DLookup("UnitCost", "tmpPurchaseLines") = 40, _
                "مسح الباركود يضيف المنتج بآخر سعر شراء")
    Call Record(PurScanCode(frm, "2*TEST-PUR-C") And DCount("*", "tmpPurchaseLines") = 1 And _
                DLookup("Quantity", "tmpPurchaseLines") = 3, "الصيغة 2*الكود تزيد الكمية في نفس السطر")
    Call Record(Not PurScanCode(frm, "NO-SUCH-CODE"), "باركود غير موجود يُرفض")
    frm!chkChargeVAT.Value = True
    RecalcPurchase frm
    Call Record(frm!lblTotal.Caption = Format$(138, "#,##0.00") And frm!lblTax.Caption = Format$(18, "#,##0.00"), _
                "3 × 40 + ضريبة 15% = 138.00")
    frm!chkChargeVAT.Value = False
    RecalcPurchase frm
    Call Record(frm!lblTotal.Caption = Format$(120, "#,##0.00"), "مورد غير مسجل بالضريبة: 120.00")
    PurRemoveLine frm!subLines.Form
    Call Record(DCount("*", "tmpPurchaseLines") = 0, "حذف السطر")
    DoCmd.Close acForm, "frmPurchaseInvoice", acSaveNo

    DoCmd.OpenForm "frmInventory", acNormal, , , , acHidden
    Set frm = Forms("frmInventory")
    frm!txtSearch.Value = "TEST-PUR-C"
    InventoryRefresh frm
    Call Record(frm!lstProducts.ListCount = 2, "البحث في شاشة المخزون (عنوان + منتج)")
    DoCmd.Close acForm, "frmInventory", acSaveNo
    DoCmd.OpenForm "frmStockCount", acNormal, , , , acHidden
    Call Record(IsFormOpen("frmStockCount"), "فتح شاشة الجرد")
    DoCmd.Close acForm, "frmStockCount", acSaveNo
    CurrentDb.Execute "DELETE FROM Products WHERE ProductID = " & pid, dbFailOnError
    Exit Sub
EH:
    Fail "شاشات المشتريات: خطأ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmPurchaseInvoice", acSaveNo
    DoCmd.Close acForm, "frmInventory", acSaveNo
    DoCmd.Close acForm, "frmStockCount", acSaveNo
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines"
    CurrentDb.Execute "DELETE FROM Products WHERE ProductCode = 'TEST-PUR-C'"
End Sub

'------------------------------------------------------------------------------
' helpers
'------------------------------------------------------------------------------
Private Sub PurCartAdd(ByVal ProductID As Long, ByVal Qty As Currency, ByVal Cost As Currency)
    CurrentDb.Execute "INSERT INTO tmpPurchaseLines (ProductID, ProductCode, ProductName, Quantity, UnitCost, " & _
        "LineDiscount, SellingPrice) SELECT ProductID, ProductCode, ProductName, " & Trim$(Str$(Qty)) & ", " & _
        Trim$(Str$(Cost)) & ", 0, SellingPrice FROM Products WHERE ProductID = " & ProductID, dbFailOnError
End Sub

Private Function Stock(ByVal ProductID As Long) As Currency
    Stock = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function AvgCost(ByVal ProductID As Long) As Currency
    AvgCost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function SupBalance(ByVal SupplierID As Long) As Currency
    SupBalance = Nz(DbValue("SELECT CurrentBalance FROM Suppliers WHERE SupplierID = " & SupplierID), -9999)
End Function

Private Function PurInv(ByVal InvoiceID As Long, ByVal FieldName As String) As Currency
    PurInv = Nz(DbValue("SELECT " & FieldName & " FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & InvoiceID), -1)
End Function

Private Sub Record(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        Fail Label
    End If
End Sub

Private Sub Fail(ByVal Msg As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & Msg & vbCrLf
    Debug.Print "[X] " & Msg
End Sub
'''


def placeholders():
    e = R.expectations()
    values = {
        "SALE_PRICE": R.SALE_PRICE, "P1_QTY": R.P1_QTY, "P1_COST": R.P1_COST, "P1_PAID": R.P1_PAID,
        "SALE_QTY": R.SALE_QTY, "P2_QTY": R.P2_QTY, "P2_COST": R.P2_COST, "P2_NEW_PRICE": R.P2_NEW_PRICE,
        "RETURN_QTY": R.RETURN_QTY, "PAYMENT": R.PAYMENT, "B_OPENING_QTY": R.B_OPENING_QTY,
        "B_OPENING_COST": R.B_OPENING_COST, "B_OUT_QTY": R.B_OUT_QTY, "COUNT_A": R.COUNT_A,
        "COUNT_B": R.COUNT_B,
    }
    values.update(e)
    return {k: num(v) for k, v in values.items()}


def build_test_purchases_vba() -> str:
    text = TEMPLATE.replace("@@VECTORS@@", vector_lines())
    for key, value in placeholders().items():
        text = text.replace(f"@@{key}@@", value)
    assert not re.search(r"@@[A-Z0-9_]+@@", text), re.findall(r"@@[A-Z0-9_]+@@", text)
    return text
