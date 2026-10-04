"""Generate modTestSales.bas: the in-Access tests of Phase 6.
Vectors come from the Python references (also executed through LibreOffice
in tests/test_vba_runtime.py), so Access must produce identical results."""

import datetime as dt
import re

import pricing as P
import qr_reference as R
import zatca_reference as Z
from generate_common import vba_str

ARABIC_LOCAL = dt.datetime(2026, 10, 4, 12, 15)


def qr_vectors():
    return [
        ("رمز QR لمثال الهيئة الرسمي",
         "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAFBjE1MC4wMA=="),
        ("رمز QR لاسم محل عربي",
         Z.qr_payload("متجر الاختبار", "300000000000003", Z.timestamp(ARABIC_LOCAL), 1150, 150)),
        ("رمز QR طويل (الإصدار 15)", "Q" * 400),
    ]


def long_assign(var, text, width=180):
    parts = [text[i:i + width] for i in range(0, len(text), width)] or [""]
    return [f"    {var} = {'' if i == 0 else var + ' & '}{vba_str(p)}" for i, p in enumerate(parts)]


def vector_lines():
    out = []
    for label, text in qr_vectors():
        out += long_assign("e", R.checksum(R.encode(text.encode())[0]))
        out += long_assign("t", text)
        out.append(f"    Call Record(QRChecksum(t) = e, {vba_str(label)})")
    out.append('    Call Record(BuildZatcaQR("Bobs Records", "310122393500003", DateSerial(2022, 4, 25) + TimeSerial(18, 30, 0), 1000, 150) = _')
    out.append('        "AQxCb2JzIFJlY29yZHMCDzMxMDEyMjM5MzUwMDAwMwMUMjAyMi0wNC0yNVQxNTozMDowMFoEBzEwMDAuMDAFBjE1MC4wMA==", _')
    out.append('        "حمولة QR لمثال الهيئة الرسمي")')
    arabic = Z.qr_payload("متجر الاختبار", "300000000000003", Z.timestamp(ARABIC_LOCAL), 1150, 150)
    out += long_assign("e", arabic)
    out.append('    Call Record(BuildZatcaQR("متجر الاختبار", "300000000000003", DateSerial(2026, 10, 4) + TimeSerial(12, 15, 0), 1150, 150) = e, _')
    out.append('        "حمولة QR بالعربية (UTF-8)")')
    for case in P.CASES:
        label, incl, inv, lines = case
        o = P.calc(*P.case_inputs(case)[:2], incl)
        out.append("    CalcReset")
        for q, p, d, r in lines:
            out.append(f"    CalcAddLine {q}, {p}, {d}, {r}")
        out.append(f'    msg = CalcRun({inv}, {"True" if incl else "False"})')
        checks = [f'CalcTotal("{k}") = {v}' for k, v in (("SUBTOTAL", o.subtotal), ("DISCOUNT", o.discount),
                                                         ("TAX", o.tax), ("TOTAL", o.total))]
        for i, ln in enumerate(o.lines):
            checks.append(f'CalcLine({i}, "UNIT") = {ln.unit_price}')
            checks.append(f'CalcLine({i}, "TAX") = {ln.tax}')
        out.append(f'    Call Record(Len(msg) = 0 And {" And ".join(checks[:4])}, {vba_str("حساب: " + label)})')
        out.append(f'    Call Record({" And ".join(checks[4:])}, {vba_str("أسطر: " + label)})')
    return "\n".join(out)


TEMPLATE = r'''Attribute VB_Name = "modTestSales"
'==============================================================================
' modTestSales  -  Retail Store Management System (Phase 6 tests)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_test_sales.py
'
'   TestSales   1) calculation, ZATCA and QR vectors (same as the Python references)
'               2) the full sales cycle inside a transaction that is rolled back:
'                  opening stock 100 -> cash sale 10 -> stock 90 (spec), credit sale,
'                  receipt voucher, return, and the refusals (unpaid cash sale,
'                  credit for the cash customer, more than the stock, over-return)
'               3) the point-of-sale screen itself (scan, merge, discount, change)
'==============================================================================
Option Compare Database
Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_report As String

Public Function TestSales() As Boolean
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestSales  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    g_AutoAnswer = True
    EnsureLocalTables

    TestVectors
    TestPosting
    TestPOSScreen

    g_SilentMode = False
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        MsgBox "جميع اختبارات المبيعات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "لم تُترك أي بيانات اختبار.", vbInformation + MSG_RTL, "TestSales"
        TestSales = True
    Else
        MsgBox "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestSales"
    End If
End Function

'------------------------------------------------------------------------------
' 1) vectors
'------------------------------------------------------------------------------
Private Sub TestVectors()
    Dim e As String, t As String, msg As String
@@VECTORS@@
    Call Record(ZatcaAmount(1000) = "1000.00" And ZatcaAmount(0.5) = "0.50", "صيغة المبالغ في QR")
    Call Record(ZatcaTimestamp(DateSerial(2026, 10, 4) + TimeSerial(1, 5, 9)) = "2026-10-03T22:05:09Z", "الوقت في QR بتوقيت UTC")
    Call Record(NewUUID() Like "????????-????-4???-????-????????????", "توليد UUID")
End Sub

'------------------------------------------------------------------------------
' 2) the sales cycle, rolled back at the end
'------------------------------------------------------------------------------
Private Sub TestPosting()
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, integrityBefore As Long
    Dim pid As Long, cid As Long, msg As String, cashInv As Long, creditInv As Long, dummy As Long
    Dim retID As Long, payID As Long, invDate As Date
    On Error GoTo EH
    Set db = CurrentDb
    integrityBefore = DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery")
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "UPDATE Settings SET StoreName = 'TEST متجر', VATNumber = '300000000000003', " & _
               "VATRate = 0.15, PricesIncludeVAT = True, AllowNegativeStock = False", dbFailOnError
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    db.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, PurchasePrice, AverageCost, " & _
               "SellingPrice, CurrentQuantity, MinimumQuantity) VALUES ('TEST-POS-1', 'TEST منتج', " & _
               "'TESTPOS0001', 60, 60, 115, 0, 5)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-POS-1'")
    ApplyStockMovement db, pid, 100, TT_OPENING, 60, "MANUAL", 0, "TEST-OPEN"
    Call Record(Stock(pid) = 100, "رصيد افتتاحي 100 قطعة")
    db.Execute "INSERT INTO Customers (CustomerName, AllowCredit, CreditLimit) VALUES " & _
               "('TEST عميل آجل', True, 1000)", dbFailOnError
    cid = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST عميل آجل'")

    ' --- cash sale: 10 x 115 (incl. VAT)
    CartAdd pid, 10
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", cashInv)
    Call Record(Len(msg) = 0 And cashInv > 0, "حفظ فاتورة نقدية " & msg)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Stock(pid) = 90, "بعد بيع 10 قطع أصبح المخزون 90")
    Call Record(Inv(cashInv, "TotalAmount") = 1150 And Inv(cashInv, "Tax") = 150 And _
                Inv(cashInv, "TaxableAmount") = 1000, "الإجمالي 1150 = 1000 + ضريبة 150")
    Call Record(Inv(cashInv, "PaidAmount") = 1150 And Inv(cashInv, "RemainingAmount") = 0, "مدفوعة بالكامل")
    Call Record(Nz(DbValue("SELECT UnitCost FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & cashInv), 0) = 60, _
                "حفظ تكلفة الصنف لحظة البيع")
    Call Record(Nz(DbValue("SELECT QuantityAfter FROM InventoryTransactions WHERE ReferenceType = 'SALE' " & _
                "AND ReferenceID = " & cashInv), 0) = 90, "حركة مخزون بيع بالرصيد بعدها 90")
    invDate = DbValue("SELECT InvoiceDate FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv)
    Call Record(DbValue("SELECT QRCodeData FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv) = _
                BuildZatcaQR("TEST متجر", "300000000000003", invDate, 1150, 150), "رمز QR محفوظ في الفاتورة")
    Call Record(DbValue("SELECT InvoiceSubType FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv) = "SIMPLIFIED" And _
                Nz(DbValue("SELECT ICV FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv), 0) > 0 And _
                Len(Nz(DbValue("SELECT InvoiceUUID FROM SalesInvoices WHERE SalesInvoiceID = " & cashInv), "")) = 36, _
                "فاتورة مبسطة برقم ICV و UUID")

    ' --- refusals
    CartAdd pid, 1
    msg = PostSaleFromCart(1, "CASH", 1, 0, 50, "", dummy)
    Call Record(Len(msg) > 0 And Stock(pid) = 90, "رفض بيع نقدي بمبلغ مدفوع أقل من الإجمالي")
    msg = PostSaleFromCart(1, "CREDIT", 1, 0, 0, "", dummy)
    Call Record(Len(msg) > 0, "رفض البيع الآجل للعميل النقدي")
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CartAdd pid, 500
    g_AutoAnswer = False
    msg = PostSaleFromCart(1, "CASH", 1, 0, Null, "", dummy)
    g_AutoAnswer = True
    Call Record(Len(msg) > 0 And Stock(pid) = 90, "منع البيع بأكثر من المتوفر بدون موافقة المدير")
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError

    ' --- credit sale: 2 x 115 = 230, paid 50
    CartAdd pid, 2
    msg = PostSaleFromCart(cid, "CREDIT", 1, 0, 50, "", creditInv)
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    Call Record(Len(msg) = 0 And Inv(creditInv, "RemainingAmount") = 180, "بيع آجل: المتبقي 180 " & msg)
    Call Record(Balance(cid) = 180, "رصيد العميل الآجل 180")
    Call Record(Stock(pid) = 88, "المخزون 88")

    ' --- receipt voucher 100
    msg = PostCustomerPayment(cid, 100, 1, "TEST", payID)
    Call Record(Len(msg) = 0 And Balance(cid) = 80, "سند قبض 100: الرصيد 80 " & msg)
    Call Record(Len(PostCustomerPayment(1, 10, 1, "", dummy)) > 0, "رفض سند قبض على العميل النقدي")

    ' --- return 1 unit of the credit sale to stock, credited to the account
    db.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    db.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ReturnQty, ReturnToStock) " & _
               "SELECT SalesDetailID, ProductID, 1, True FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & _
               creditInv, dbFailOnError
    Call Record(Len(PostSalesReturn(creditInv, "", "CREDIT", Null, dummy)) > 0, "رفض مرتجع بدون سبب")
    msg = PostSalesReturn(creditInv, "TEST إرجاع", "CREDIT", Null, retID)
    Call Record(Len(msg) = 0 And retID > 0, "حفظ إشعار دائن " & msg)
    Call Record(Nz(DbValue("SELECT TotalAmount FROM SalesReturns WHERE SalesReturnID = " & retID), 0) = 115 And _
                DbValue("SELECT InvoiceTypeCode FROM SalesReturns WHERE SalesReturnID = " & retID) = "381", _
                "قيمة المرتجع 115 ونوعه إشعار دائن 381")
    Call Record(Stock(pid) = 89 And Balance(cid) = -35, "المرتجع أعاد قطعة للمخزون (89) وخصم 115 من الرصيد (-35)")
    db.Execute "UPDATE tmpReturnLines SET ReturnQty = 5", dbFailOnError
    Call Record(Len(PostSalesReturn(creditInv, "TEST", "CREDIT", Null, dummy)) > 0, _
                "رفض إرجاع أكثر من الكمية المتبقية")

    Call Record(DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = integrityBefore, "فحص السلامة لا يجد أي مشكلة جديدة")
    Call Record(CurrentDb.OpenRecordset("SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' " & _
                "AND DocID = " & cashInv)(0) = 1, "بيانات طباعة الفاتورة")

    ws.Rollback
    inTrans = False
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    db.Execute "DELETE FROM tmpReturnLines", dbFailOnError
    Call Record(DbValue("SELECT COUNT(*) FROM Products WHERE ProductCode = 'TEST-POS-1'") = 0, "التراجع عن كل بيانات الاختبار")
    Exit Sub
EH:
    Fail "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
End Sub

'------------------------------------------------------------------------------
' 3) the point-of-sale screen (nothing is posted)
'------------------------------------------------------------------------------
Private Sub TestPOSScreen()
    Dim frm As Access.Form, pid As Long
    On Error GoTo EH
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, Barcode, SellingPrice) VALUES " & _
                      "('TEST-POS-2', 'TEST منتج شاشة', 'TESTPOS0002', 11.5)", dbFailOnError
    pid = DLookup("ProductID", "Products", "ProductCode = 'TEST-POS-2'")
    DoCmd.OpenForm "frmPOS", acNormal, , , , acHidden
    Set frm = Forms("frmPOS")
    Call Record(ScanCode(frm, "TESTPOS0002"), "مسح الباركود يضيف المنتج")
    Call Record(ScanCode(frm, "TESTPOS0002") And DCount("*", "tmpPOSLines") = 1 And _
                DLookup("Quantity", "tmpPOSLines") = 2, "مسح نفس المنتج يزيد الكمية في نفس السطر")
    Call Record(ScanCode(frm, "3*TEST-POS-2") And DLookup("Quantity", "tmpPOSLines") = 5, _
                "الصيغة 3*الكود تضيف 3 قطع (بالكود أيضًا)")
    Call Record(frm!lblTotal.Caption = Format$(57.5, "#,##0.00"), "الإجمالي 5 × 11.50 = 57.50")
    Call Record(Not ScanCode(frm, "NO-SUCH-CODE"), "باركود غير موجود يُرفض")
    frm!txtInvoiceDiscount.Value = 7.5
    RecalcPOS frm
    Call Record(frm!lblTotal.Caption = Format$(50, "#,##0.00") And frm!lblTax.Caption = Format$(6.52, "#,##0.00"), _
                "خصم 7.50 على الفاتورة: الإجمالي 50.00 والضريبة 6.52")
    frm!txtTendered.Value = 100
    RecalcPOS frm
    Call Record(InStr(frm!lblChange.Caption, Format$(50, "#,##0.00")) > 0, "الباقي للعميل 50.00")
    RemoveCurrentLine frm!subLines.Form
    Call Record(DCount("*", "tmpPOSLines") = 0 And frm!lblTotal.Caption = Format$(0, "#,##0.00"), "حذف السطر")
    DoCmd.Close acForm, "frmPOS", acSaveNo
    CurrentDb.Execute "DELETE FROM Products WHERE ProductID = " & pid, dbFailOnError
    Exit Sub
EH:
    Fail "شاشة البيع: خطأ " & Err.Number & ": " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, "frmPOS", acSaveNo
    CurrentDb.Execute "DELETE FROM tmpPOSLines"
    CurrentDb.Execute "DELETE FROM Products WHERE ProductCode = 'TEST-POS-2'"
End Sub

'------------------------------------------------------------------------------
' helpers
'------------------------------------------------------------------------------
Private Sub CartAdd(ByVal ProductID As Long, ByVal Qty As Currency)
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, " & Trim$(Str$(Qty)) & _
        ", SellingPrice, 0, CurrentQuantity FROM Products WHERE ProductID = " & ProductID, dbFailOnError
End Sub

Private Function Stock(ByVal ProductID As Long) As Currency
    Stock = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), -1)
End Function

Private Function Balance(ByVal CustomerID As Long) As Currency
    Balance = Nz(DbValue("SELECT CurrentBalance FROM Customers WHERE CustomerID = " & CustomerID), -9999)
End Function

Private Function Inv(ByVal InvoiceID As Long, ByVal FieldName As String) As Currency
    Inv = Nz(DbValue("SELECT " & FieldName & " FROM SalesInvoices WHERE SalesInvoiceID = " & InvoiceID), -1)
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


def build_test_sales_vba() -> str:
    text = TEMPLATE.replace("@@VECTORS@@", vector_lines())
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text
