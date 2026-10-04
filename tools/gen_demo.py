"""Generate modDemoData.bas from tools/demo_data.py.

  LoadDemoData     posts the demo plan through the real posting functions, in one
                   transaction (all or nothing), then dates every document on its
                   planned day (and rebuilds the ZATCA QR with that date)
  VerifyDemoData   compares the result with the Python replay (stock, average cost,
                   balances, totals, low-stock and slow-moving products, integrity)
  RemoveDemoData   removes the demo before going live (after a backup), only when
                   no real document was entered after it
"""

import re
from decimal import Decimal as D

import demo_data as DD
from generate_common import vba_str

CHUNK = 6          # plan steps per generated procedure (keeps procedures small)


def num(v) -> str:
    return format(D(str(v)).normalize(), "f")


def money(v) -> str:
    return format(D(str(v)).quantize(D("0.01")), "f")


def product_literal(n: int) -> str:
    return vba_str(DD.PRODUCTS[n - 1].barcode)


def step_lines(step, sale_no, purchase_no):
    """VBA lines for one plan step. sale_no / purchase_no: index of this document."""
    op = step["op"]
    out = [f"    ' --- {op} (day -{step['day']})",
           f"    when = DemoWhen({step['day']}, {step['hour']})",
           f"    SetDemoUser {vba_str(step.get('user') or '')}"]
    if op == "purchase":
        out.append('    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError')
        for n, q in step["lines"]:
            out.append(f"    PurLine {product_literal(n)}, {q}, {vba_str(DD.PRODUCTS[n - 1].cost)}")
        sup = vba_str(DD.SUPPLIERS[step["supplier"] - 1].name)
        paid = str(step["paid"]) if step["credit"] else "Null"
        out.append(f"    Check PostPurchaseFromCart(SupplierIDOf({sup}), {vba_str(step['ref'])}, DateValue(when), "
                   f"{vba_str('CREDIT' if step['credit'] else 'CASH')}, 1, True, {money(step.get('discount', 0))}, "
                   f"{paid}, \"\", id), \"فاتورة شراء {purchase_no}\"")
        out.append('    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError')
        out.append('    MoveDoc "PURCHASE", id, when')
        out.append(f"    m_purchases({purchase_no}) = id")
    elif op == "sale":
        out.append('    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError')
        for n, q in step["lines"]:
            out.append(f"    CartLine {product_literal(n)}, {q}")
        cust = (f"CustomerIDOf({vba_str(DD.CUSTOMERS[step['customer'] - 1].name)})" if step.get("customer")
                else "1")
        credit = step.get("credit", False)
        paid = str(step.get("paid", 0)) if credit else "Null"
        out.append(f"    Check PostSaleFromCart({cust}, {vba_str('CREDIT' if credit else 'CASH')}, 1, "
                   f"{money(step.get('discount', 0))}, {paid}, \"\", id), \"فاتورة بيع {sale_no}\"")
        out.append('    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError')
        out.append('    MoveDoc "SALE", id, when')
        out.append(f"    m_sales({sale_no}) = id")
    elif op == "sales_return":
        out.append('    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError')
        for ln, q in step["lines"]:
            out.append(f'    CurrentDb.Execute "INSERT INTO tmpReturnLines (SalesDetailID, ProductID, ReturnQty, '
                       f'ReturnToStock) SELECT SalesDetailID, ProductID, {q}, True FROM SalesInvoiceDetails '
                       f'WHERE SalesInvoiceID = " & m_sales({step["invoice"]}) & " AND LineNumber = {ln}", dbFailOnError')
        out.append(f'    Check PostSalesReturn(m_sales({step["invoice"]}), "تلف في التغليف (تجريبي)", "CREDIT", Null, id), '
                   f'"مرتجع مبيعات"')
        out.append('    CurrentDb.Execute "DELETE FROM tmpReturnLines", dbFailOnError')
        out.append('    MoveDoc "SALES_RETURN", id, when')
    elif op == "purchase_return":
        out.append('    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError')
        for ln, q in step["lines"]:
            out.append(f'    CurrentDb.Execute "INSERT INTO tmpPurchaseReturnLines (PurchaseDetailID, ProductID, ReturnQty) '
                       f'SELECT PurchaseDetailID, ProductID, {q} FROM PurchaseInvoiceDetails WHERE PurchaseInvoiceID = " '
                       f'& m_purchases({step["invoice"]}) & " AND LineNumber = {ln}", dbFailOnError')
        out.append(f'    Check PostPurchaseReturn(m_purchases({step["invoice"]}), "عيب مصنعي (تجريبي)", "CREDIT", Null, id), '
                   f'"مرتجع مشتريات"')
        out.append('    CurrentDb.Execute "DELETE FROM tmpPurchaseReturnLines", dbFailOnError')
        out.append('    MoveDoc "PURCHASE_RETURN", id, when')
    elif op == "supplier_payment":
        sup = vba_str(DD.SUPPLIERS[step["supplier"] - 1].name)
        out.append(f'    Check PostSupplierPayment(SupplierIDOf({sup}), {money(step["amount"])}, 1, "دفعة من الحساب", id), '
                   f'"سند صرف"')
        out.append('    MoveDoc "SUPPLIER_PAYMENT", id, when')
    elif op == "customer_payment":
        cust = vba_str(DD.CUSTOMERS[step["customer"] - 1].name)
        out.append(f'    Check PostCustomerPayment(CustomerIDOf({cust}), {money(step["amount"])}, 1, "دفعة من الحساب", id), '
                   f'"سند قبض"')
        out.append('    MoveDoc "CUSTOMER_PAYMENT", id, when')
    elif op == "expense":
        total = money(D(step["amount"]) + D(step["tax"]))
        out.append(f'    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, '
                   f'TotalAmount, PaymentMethodID, Description, EmployeeID, CreatedAt) VALUES (" & '
                   f'SqlText(NextNumber("EXPENSE")) & ", " & SqlDate(DateValue(when)) & ", {step["type"]}, '
                   f'{money(step["amount"])}, {money(step["tax"])}, {total}, 1, " & SqlText({vba_str(step["text"])}) & '
                   f'", " & CurrentUserID() & ", " & SqlDate(when) & ")", dbFailOnError')
    elif op == "stock_out":
        out.append(f'    Check PostManualStock(ProductIDOf({product_literal(step["product"])}), TT_STOCK_OUT, {step["qty"]}, '
                   f'Null, {vba_str(step["text"])}, refNo), "خصم مخزون"')
        out.append('    MoveManual refNo, when')
    elif op == "stock_count":
        cat = vba_str(DD.CATEGORIES[step["category"] - 1])
        out.append(f'    Check CreateStockCount(CategoryIDOf({cat}), "جرد تجريبي", id), "إنشاء الجرد"')
        for n, d in step["actual"].items():
            out.append(f'    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = SystemQuantity + ({d}) '
                       f'WHERE StockCountID = " & id & " AND ProductID = " & ProductIDOf({product_literal(n)}), '
                       f'dbFailOnError')
        out.append('    Check PostStockCount(id, False, adjusted, netValue), "ترحيل الجرد"')
        out.append('    MoveDoc "STOCK_COUNT", id, when')
    else:
        raise ValueError(op)
    return out


def steps_procs():
    procs, calls = [], []
    sale_no = purchase_no = 0
    lines_by_chunk = []
    for i, step in enumerate(DD.PLAN):
        if step["op"] == "sale":
            sale_no += 1
        if step["op"] == "purchase":
            purchase_no += 1
        if i % CHUNK == 0:
            lines_by_chunk.append([])
        lines_by_chunk[-1] += step_lines(step, sale_no, purchase_no)
    for k, body in enumerate(lines_by_chunk, 1):
        name = f"DemoSteps{k}"
        calls.append(f"    {name}")
        procs.append("\n".join([f"Private Sub {name}()",
                                "    Dim when As Date, id As Long, refNo As String, adjusted As Long, "
                                "netValue As Currency"] + body + ["End Sub"]))
    return "\n".join(calls), "\n\n".join(procs), sale_no, purchase_no


def masters_lines():
    out = []
    for name in DD.CATEGORIES:
        out.append(f'    CurrentDb.Execute "INSERT INTO Categories (CategoryName, Description) VALUES (" & '
                   f'SqlText({vba_str(name)}) & ", \'{DD.DEMO_MARK}\')", dbFailOnError')
    for s in DD.SUPPLIERS:
        out.append(f'    CurrentDb.Execute "INSERT INTO Suppliers (SupplierName, ContactPerson, Mobile, VATNumber, City, '
                   f'Notes, CreatedAt) VALUES (" & SqlText({vba_str(s.name)}) & ", " & SqlText({vba_str(s.contact)}) & '
                   f'", \'{s.mobile}\', \'{s.vat}\', " & SqlText({vba_str(s.city)}) & ", \'{DD.DEMO_MARK}\', " & '
                   f'SqlDate(Date - 60) & ")", dbFailOnError')
    for i, p in enumerate(DD.PRODUCTS, 1):
        created = 150 if i == DD.SLOW_PRODUCT else 60
        out.append(f'    CurrentDb.Execute "INSERT INTO Products (ProductCode, Barcode, ProductName, CategoryID, UnitID, '
                   f'PurchasePrice, SellingPrice, MinimumQuantity, SupplierID, Notes, CreatedAt) VALUES (" & '
                   f'SqlText(NextNumber("PRODUCT_CODE")) & ", \'{p.barcode}\', " & SqlText({vba_str(p.name)}) & ", " & '
                   f'CategoryIDOf({vba_str(DD.CATEGORIES[p.category])}) & ", {p.unit}, {p.cost}, {p.price}, {p.minimum}, '
                   f'" & SupplierIDOf({vba_str(DD.SUPPLIERS[p.supplier].name)}) & ", \'{DD.DEMO_MARK}\', " & '
                   f'SqlDate(Date - {created}) & ")", dbFailOnError')
    for c in DD.CUSTOMERS:
        vat = f"'{c.vat}'" if c.vat else "Null"
        out.append(f'    CurrentDb.Execute "INSERT INTO Customers (CustomerName, Mobile, VATNumber, City, AllowCredit, '
                   f'CreditLimit, Notes, CreatedAt) VALUES (" & SqlText({vba_str(c.name)}) & ", \'{c.mobile}\', {vat}, " & '
                   f'SqlText({vba_str(c.city)}) & ", {"True" if c.credit else "False"}, {c.limit}, \'{DD.DEMO_MARK}\', " & '
                   f'SqlDate(Date - 60) & ")", dbFailOnError')
    for name, job, user, role, disc in DD.EMPLOYEES:
        out.append(f'    CurrentDb.Execute "INSERT INTO Employees (EmployeeName, JobTitle, Username, RoleID, '
                   f'MaxDiscountPercent, Notes) VALUES (" & SqlText({vba_str(name)}) & ", " & SqlText({vba_str(job)}) & '
                   f'", \'{user}\', {role}, {disc}, \'{DD.DEMO_MARK}\')", dbFailOnError')
        out.append(f'    Check SetUserPassword(UserIDOf("{user}"), DEMO_PASSWORD, True), "كلمة مرور {user}"')
    return "\n".join(out)


def verify_lines(r):
    out = []
    for n, p in enumerate(DD.PRODUCTS, 1):
        out.append(f"    Expect ProductValue({product_literal(n)}, \"CurrentQuantity\") = {num(r.stock[n])} And "
                   f"ProductValue({product_literal(n)}, \"AverageCost\") = CCur({num(r.avg[n])}), "
                   f"{vba_str(f'{p.name}: الكمية {num(r.stock[n])} والمتوسط {num(r.avg[n])}')}")
    for n, c in enumerate(DD.CUSTOMERS, 1):
        bal = money(r.customer_balance[n])
        out.append(f"    Expect Nz(DbValue(\"SELECT CurrentBalance FROM Customers WHERE CustomerName = \" & "
                   f"SqlText({vba_str(c.name)})), -1) = CCur({bal}), {vba_str(f'رصيد {c.name} = {bal}')}")
    for n, s in enumerate(DD.SUPPLIERS, 1):
        bal = money(r.supplier_balance[n])
        out.append(f"    Expect Nz(DbValue(\"SELECT CurrentBalance FROM Suppliers WHERE SupplierName = \" & "
                   f"SqlText({vba_str(s.name)})), -1) = CCur({bal}), {vba_str(f'رصيد {s.name} = {bal}')}")
    return "\n".join(out)


TEMPLATE = r'''Attribute VB_Name = "modDemoData"
'==============================================================================
' modDemoData  -  Retail Store Management System (Phase 11: demo data)
'
' GENERATED FILE - do not edit by hand.  Source: tools/demo_data.py, tools/gen_demo.py
'
'   LoadDemoData     enters the demo store: @@COUNTS@@.
'                    Every document goes through the real posting functions, in ONE
'                    transaction (all or nothing), then gets its planned date.
'                    Only on a database without documents.
'   VerifyDemoData   compares the result with the Python replay of the same plan.
'   RemoveDemoData   removes the demo before real use (backup first). Refused
'                    once real documents were entered after the demo.
' Demo users (manager, cashier1, cashier2) have the temporary password
' "@@PASSWORD@@" and must change it at their first login.
'==============================================================================
Option Compare Database
Option Explicit

Private Const DEMO_PASSWORD As String = "@@PASSWORD@@"
Private Const DEMO_MARK As String = "@@MARK@@"

Private m_sales(1 To @@SALES@@) As Long
Private m_purchases(1 To @@PURCHASES@@) As Long
Private m_adminID As Long
Private m_passed As Long
Private m_failed As Long
Private m_report As String

'==============================================================================
' Load
'==============================================================================
Public Function LoadDemoData() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, problem As String, errText As String
    On Error GoTo EH
    Calendar = vbCalGreg
    EnsureTestUser
    m_adminID = CurrentUserID()
    problem = DemoBlocker()
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation + MSG_RTL, "LoadDemoData"
        Exit Function
    End If
    If MsgBox("سيتم إدخال متجر تجريبي كامل:" & vbCrLf & "@@COUNTS@@." & vbCrLf & vbCrLf & _
              "للتدريب والتجربة فقط. قبل الاستخدام الفعلي احذفها بالأمر RemoveDemoData." & vbCrLf & _
              "هل تريد المتابعة؟", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, "LoadDemoData") <> vbYes Then
        Exit Function
    End If
    DoCmd.Hourglass True
    EnsureLocalTables
    ClearWorkTables
    g_SilentMode = True
    g_AutoAnswer = True
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    DemoSettings
    DemoMasters
@@STEP_CALLS@@
    ws.CommitTrans
    inTrans = False
    TempVars.Add "UserID", m_adminID
    LogAction "DEMO_LOADED", "", "", "@@COUNTS@@"
    g_SilentMode = False
    DoCmd.Hourglass False
    LoadDemoData = VerifyDemoData()
    Exit Function
EH:
    errText = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    TempVars.Add "UserID", m_adminID
    ClearWorkTables
    g_SilentMode = False
    DoCmd.Hourglass False
    MsgBox "تعذر إدخال البيانات التجريبية، ولم يُحفظ منها أي شيء:" & vbCrLf & vbCrLf & errText, _
           vbCritical + MSG_RTL, "LoadDemoData"
End Function

Private Function DemoBlocker() As String
    ' Reasons not to load: documents already exist, demo already there, settings differ from the plan.
    Dim t As Variant
    If Not IsAdministrator() Then
        DemoBlocker = "إدخال البيانات التجريبية لمدير النظام فقط."
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "PurchaseInvoices", "InventoryTransactions", "Expenses", _
                        "CustomerPayments", "SupplierPayments", "StockCounts")
        If Nz(DbValue("SELECT COUNT(*) FROM [" & t & "]"), 0) > 0 Then
            DemoBlocker = "توجد بيانات فعلية (" & t & ")." & vbCrLf & _
                          "البيانات التجريبية تُدخل في قاعدة بدون مستندات فقط (نسخة للتدريب)."
            Exit Function
        End If
    Next
    If Nz(DbValue("SELECT COUNT(*) FROM Products WHERE Barcode = " & SqlText("@@FIRST_BARCODE@@")), 0) > 0 Or _
       Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE Username = 'manager' OR Username = 'cashier1' " & _
                  "OR Username = 'cashier2'"), 0) > 0 Then
        DemoBlocker = "يوجد جزء من البيانات التجريبية مسبقًا (منتج أو مستخدم بنفس الاسم)."
        Exit Function
    End If
    If Nz(SettingValue("VATRate"), 0) <> 0.15 Or Not Nz(SettingValue("PricesIncludeVAT"), False) Then
        DemoBlocker = "البيانات التجريبية مبنية على نسبة ضريبة 15% وأسعار شاملة الضريبة." & vbCrLf & _
                      "أعد هذين الإعدادين ثم حاول مرة أخرى."
    End If
End Function

Private Sub DemoSettings()
    ' The demo store details are used only when the store is not set up yet.
    If Not IsNull(SettingValue("VATNumber")) Then Exit Sub
    CurrentDb.Execute "UPDATE Settings SET @@SETTINGS_SQL@@ WHERE SettingID = 1", dbFailOnError
End Sub

Private Sub DemoMasters()
@@MASTERS@@
End Sub

@@STEP_PROCS@@

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function DemoWhen(ByVal DaysAgo As Long, ByVal HourOfDay As Long) As Date
    If DaysAgo = 0 Then
        DemoWhen = DateAdd("n", -2, Now)                     ' today: just now
    Else
        DemoWhen = DateValue(Date - DaysAgo) + TimeSerial(HourOfDay, 0, 0)
    End If
End Function

Private Sub SetDemoUser(ByVal Username As String)
    If Len(Username) = 0 Then
        TempVars.Add "UserID", m_adminID
    Else
        TempVars.Add "UserID", UserIDOf(Username)
    End If
End Sub

Private Sub Check(ByVal Msg As String, ByVal Label As String)
    If Len(Msg) > 0 Then Err.Raise vbObjectError + 811, "LoadDemoData", Label & ": " & Msg
End Sub

Private Function ProductIDOf(ByVal Barcode As String) As Long
    ProductIDOf = DbValue("SELECT ProductID FROM Products WHERE Barcode = " & SqlText(Barcode))
End Function

Private Function CustomerIDOf(ByVal PartyName As String) As Long
    CustomerIDOf = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = " & SqlText(PartyName))
End Function

Private Function SupplierIDOf(ByVal PartyName As String) As Long
    SupplierIDOf = DbValue("SELECT SupplierID FROM Suppliers WHERE SupplierName = " & SqlText(PartyName))
End Function

Private Function CategoryIDOf(ByVal CategoryName As String) As Long
    CategoryIDOf = DbValue("SELECT CategoryID FROM Categories WHERE CategoryName = " & SqlText(CategoryName))
End Function

Private Function UserIDOf(ByVal Username As String) As Long
    UserIDOf = DbValue("SELECT EmployeeID FROM Employees WHERE Username = " & SqlText(Username))
End Function

Private Sub CartLine(ByVal Barcode As String, ByVal Qty As Long)
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, " & _
        "LineDiscount, Available) SELECT ProductID, ProductCode, ProductName, " & Qty & ", SellingPrice, 0, " & _
        "CurrentQuantity FROM Products WHERE Barcode = " & SqlText(Barcode), dbFailOnError
End Sub

Private Sub PurLine(ByVal Barcode As String, ByVal Qty As Long, ByVal Cost As String)
    CurrentDb.Execute "INSERT INTO tmpPurchaseLines (ProductID, ProductCode, ProductName, Quantity, UnitCost, " & _
        "LineDiscount, SellingPrice) SELECT ProductID, ProductCode, ProductName, " & Qty & ", " & Cost & _
        ", 0, SellingPrice FROM Products WHERE Barcode = " & SqlText(Barcode), dbFailOnError
End Sub

Private Sub ClearWorkTables()
    Dim t As Variant
    For Each t In Array("tmpPOSLines", "tmpReturnLines", "tmpPurchaseLines", "tmpPurchaseReturnLines")
        CurrentDb.Execute "DELETE FROM " & t, dbFailOnError
    Next
End Sub

Private Sub MoveDoc(ByVal Kind As String, ByVal DocID As Long, ByVal When As Date)
    ' Puts a posted document (and its stock movements) on its planned date.
    Dim tbl As String, key As String, fld As String, ref As String, rs As DAO.Recordset, vatNo As Variant
    Select Case Kind
        Case "SALE":             tbl = "SalesInvoices": key = "SalesInvoiceID": fld = "InvoiceDate"
        Case "SALES_RETURN":     tbl = "SalesReturns": key = "SalesReturnID": fld = "ReturnDate"
        Case "PURCHASE":         tbl = "PurchaseInvoices": key = "PurchaseInvoiceID": fld = "InvoiceDate"
        Case "PURCHASE_RETURN":  tbl = "PurchaseReturns": key = "PurchaseReturnID": fld = "ReturnDate"
        Case "CUSTOMER_PAYMENT": tbl = "CustomerPayments": key = "PaymentID": fld = "PaymentDate"
        Case "SUPPLIER_PAYMENT": tbl = "SupplierPayments": key = "PaymentID": fld = "PaymentDate"
        Case "STOCK_COUNT":      tbl = "StockCounts": key = "StockCountID": fld = "CountDate"
    End Select
    If Kind = "STOCK_COUNT" Then
        CurrentDb.Execute "UPDATE StockCounts SET CountDate = " & SqlDate(When) & ", PostedAt = " & SqlDate(When) & _
                          " WHERE StockCountID = " & DocID, dbFailOnError
    Else
        CurrentDb.Execute "UPDATE " & tbl & " SET " & fld & " = " & SqlDate(When) & ", CreatedAt = " & SqlDate(When) & _
                          " WHERE " & key & " = " & DocID, dbFailOnError
    End If
    Select Case Kind
        Case "SALE", "SALES_RETURN", "PURCHASE", "PURCHASE_RETURN", "STOCK_COUNT"
            CurrentDb.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(When) & ", CreatedAt = " & _
                              SqlDate(When) & " WHERE ReferenceType = " & SqlText(Kind) & " AND ReferenceID = " & DocID, _
                              dbFailOnError
    End Select
    If Kind = "SALE" Or Kind = "SALES_RETURN" Then
        ' the ZATCA QR code carries the time stamp of the document
        vatNo = SettingValue("VATNumber")
        If Not IsNull(vatNo) Then
            Set rs = CurrentDb.OpenRecordset("SELECT QRCodeData, TotalAmount, Tax FROM " & tbl & " WHERE " & key & _
                                             " = " & DocID, dbOpenDynaset)
            rs.Edit
            rs!QRCodeData = BuildZatcaQR(Nz(SettingValue("StoreName"), ""), CStr(vatNo), When, rs!TotalAmount, rs!Tax)
            rs.Update
            rs.Close
        End If
    End If
End Sub

Private Sub MoveManual(ByVal RefNumber As String, ByVal When As Date)
    CurrentDb.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(When) & ", CreatedAt = " & _
                      SqlDate(When) & " WHERE ReferenceType = 'MANUAL' AND ReferenceNumber = " & SqlText(RefNumber), _
                      dbFailOnError
End Sub

'==============================================================================
' Verify
'==============================================================================
Public Function VerifyDemoData() As Boolean
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== VerifyDemoData  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    Expect DbValue("SELECT COUNT(*) FROM Products WHERE Notes = '" & DEMO_MARK & "'") = @@N_PRODUCTS@@, "@@N_PRODUCTS@@ منتجًا"
    Expect DbValue("SELECT COUNT(*) FROM Categories WHERE Description = '" & DEMO_MARK & "'") = @@N_CATEGORIES@@, "@@N_CATEGORIES@@ تصنيفات"
    Expect DbValue("SELECT COUNT(*) FROM Customers WHERE Notes = '" & DEMO_MARK & "'") = @@N_CUSTOMERS@@, "@@N_CUSTOMERS@@ عملاء"
    Expect DbValue("SELECT COUNT(*) FROM Suppliers WHERE Notes = '" & DEMO_MARK & "'") = @@N_SUPPLIERS@@, "@@N_SUPPLIERS@@ موردين"
    Expect DbValue("SELECT COUNT(*) FROM Employees WHERE Notes = '" & DEMO_MARK & "' AND MustChangePassword = True " & _
                   "AND PasswordHash Is Not Null") = @@N_EMPLOYEES@@, "@@N_EMPLOYEES@@ موظفين بكلمة مرور مؤقتة"
    Expect DbValue("SELECT COUNT(*) FROM SalesInvoices") = @@SALES@@, "@@SALES@@ فواتير بيع"
    Expect DbValue("SELECT COUNT(*) FROM PurchaseInvoices") = @@PURCHASES@@, "@@PURCHASES@@ فواتير شراء"
    Expect DbValue("SELECT COUNT(*) FROM Expenses") = @@N_EXPENSES@@, "@@N_EXPENSES@@ مصروفات"
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM SalesInvoices"), 0) = CCur(@@SALES_TOTAL@@), "إجمالي المبيعات @@SALES_TOTAL@@"
    Expect Nz(DbValue("SELECT Sum(TotalAmount) FROM PurchaseInvoices"), 0) = CCur(@@PURCHASES_TOTAL@@), "إجمالي المشتريات @@PURCHASES_TOTAL@@"
@@VERIFY@@
    Expect DbValue("SELECT COUNT(*) FROM LowStockQuery AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                   "WHERE p.Notes = '" & DEMO_MARK & "'") = @@LOW@@, "@@LOW@@ منتجات منخفضة المخزون (للتنبيه والتقرير)"
    If Nz(SettingValue("SlowMovingDays"), 90) = 90 Then
        Expect DbValue("SELECT COUNT(*) FROM SlowMovingProductsQuery AS s INNER JOIN Products AS p ON s.ProductID = " & _
                       "p.ProductID WHERE p.Notes = '" & DEMO_MARK & "'") = @@SLOW@@, "@@SLOW@@ منتج غير متحرك"
    End If
    Expect DbValue("SELECT COUNT(*) FROM StockCounts WHERE Status = 'POSTED'") = 1, "جرد مُرحّل واحد"
    Expect DbValue("SELECT COUNT(*) FROM IntegrityCheckQuery") = 0, "فحص سلامة البيانات: لا توجد أي مشكلة"
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        TestMsg "البيانات التجريبية مطابقة تمامًا للنتائج المحسوبة مسبقًا (" & m_passed & " فحصًا)." & vbCrLf & _
                "المستخدمون التجريبيون: manager و cashier1 و cashier2، وكلمة المرور المؤقتة: " & DEMO_PASSWORD, _
                vbInformation + MSG_RTL, "VerifyDemoData"
        VerifyDemoData = True
    Else
        TestMsg "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
                vbExclamation + MSG_RTL, "VerifyDemoData"
    End If
End Function

Private Function ProductValue(ByVal Barcode As String, ByVal FieldName As String) As Currency
    ProductValue = Nz(DbValue("SELECT " & FieldName & " FROM Products WHERE Barcode = " & SqlText(Barcode)), -1)
End Function

Private Sub Expect(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        m_failed = m_failed + 1
        m_report = m_report & "- " & Label & vbCrLf
        Debug.Print "[X] " & Label
    End If
End Sub

'==============================================================================
' Remove (before going live)
'==============================================================================
Public Function RemoveDemoData() As Boolean
    Dim loaded As Variant, later As Long, t As Variant, ws As DAO.Workspace, inTrans As Boolean
    Dim db As DAO.Database, msg As String, file As String, errText As String
    On Error GoTo EH
    Calendar = vbCalGreg
    EnsureTestUser
    If Not IsAdministrator() Then
        MsgBox "حذف البيانات التجريبية لمدير النظام فقط.", vbExclamation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    loaded = DMax("LogDate", "AuditLog", "ActionType = 'DEMO_LOADED'")
    If IsNull(loaded) Then
        MsgBox "لم تُدخل بيانات تجريبية في هذه القاعدة.", vbInformation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    For Each t In Array("SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns", "CustomerPayments", _
                        "SupplierPayments", "Expenses")
        later = later + Nz(DbValue("SELECT COUNT(*) FROM [" & t & "] WHERE CreatedAt > " & SqlDate(loaded)), 0)
    Next
    later = later + Nz(DbValue("SELECT COUNT(*) FROM StockCounts WHERE CountDate > " & SqlDate(loaded)), 0)
    If later > 0 Then
        MsgBox "توجد " & later & " مستندات أُدخلت بعد البيانات التجريبية، فلن تُحذف أي بيانات." & vbCrLf & _
               "(الحذف هنا لا يميز بين التجريبي والحقيقي، لذلك يُرفض لحماية بياناتك.)", vbExclamation + MSG_RTL, _
               "RemoveDemoData"
        Exit Function
    End If
    If MsgBox("سيتم حذف كل المستندات والبيانات التجريبية وإعادة ترقيم المستندات من 1." & vbCrLf & _
              "تُحفظ نسخة احتياطية أولًا. هل تريد المتابعة؟", vbQuestion + vbYesNo + vbDefaultButton2 + MSG_RTL, _
              "RemoveDemoData") <> vbYes Then Exit Function
    msg = BackupNow(file, , "_BeforeDemoRemoval")
    If Len(msg) > 0 Then
        MsgBox "تعذرت النسخة الاحتياطية، فلم يُحذف شيء:" & vbCrLf & msg, vbExclamation + MSG_RTL, "RemoveDemoData"
        Exit Function
    End If
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
@@REMOVE_SQL@@
    ws.CommitTrans
    inTrans = False
    LogAction "DEMO_REMOVED", "", "", file
    MsgBox "تم حذف البيانات التجريبية. القاعدة جاهزة لبياناتك الفعلية." & vbCrLf & _
           "النسخة قبل الحذف: " & file, vbInformation + MSG_RTL, "RemoveDemoData"
    RemoveDemoData = True
    Exit Function
EH:
    errText = Err.Description
    On Error Resume Next
    If inTrans Then ws.Rollback
    MsgBox "تعذر الحذف ولم يتغير شيء: " & errText, vbCritical + MSG_RTL, "RemoveDemoData"
End Function
'''


def remove_sql():
    """Statements of RemoveDemoData, also run on the SQLite mirror by tests/test_demo.py."""
    m = DD.DEMO_MARK
    users = ", ".join(f"'{e[2]}'" for e in DD.EMPLOYEES)
    stmts = [f"DELETE FROM {t}" for t in DD.DOCUMENT_TABLES]
    stmts += [
        f"DELETE FROM Products WHERE Notes = '{m}'",
        "UPDATE Products SET CurrentQuantity = 0",
        f"DELETE FROM Customers WHERE Notes = '{m}'",
        "UPDATE Customers SET CurrentBalance = OpeningBalance",
        f"DELETE FROM Suppliers WHERE Notes = '{m}' AND SupplierID NOT IN (SELECT SupplierID FROM Products "
        f"WHERE SupplierID Is Not Null)",
        "UPDATE Suppliers SET CurrentBalance = OpeningBalance",
        f"DELETE FROM Categories WHERE Description = '{m}' AND CategoryID NOT IN (SELECT CategoryID FROM Products)",
        f"DELETE FROM AuditLog WHERE EmployeeID IN (SELECT EmployeeID FROM Employees WHERE Notes = '{m}' "
        f"AND Username IN ({users}))",
        f"DELETE FROM Employees WHERE Notes = '{m}' AND Username IN ({users})",
    ]
    stmts += [f"UPDATE Sequences SET NextValue = 1 WHERE SequenceName = '{s}'" for s in DD.DOCUMENT_SEQUENCES
              if s != "PRODUCT_CODE"]
    stmts.append("UPDATE Sequences SET NextValue = 1 WHERE SequenceName = 'PRODUCT_CODE' AND "
                 "(SELECT COUNT(*) FROM Products) = 0")
    return stmts


def build_demo_vba() -> str:
    r = DD.simulate()
    counts = DD.demo_counts()
    calls, procs, n_sales, n_purchases = steps_procs()
    settings = ", ".join(f"{k} = '{v}'" for k, v in DD.STORE.items())
    assert "'" not in "".join(DD.STORE.values()) and '"' not in "".join(DD.STORE.values())
    low = r.db.con.execute("SELECT COUNT(*) FROM LowStockQuery").fetchone()[0]
    slow = r.db.con.execute("SELECT COUNT(*) FROM SlowMovingProductsQuery").fetchone()[0]
    counts_text = (f"{counts['Products']} منتجًا، {counts['Categories']} تصنيفات، {counts['Customers']} عملاء، "
                   f"{counts['Suppliers']} موردين، {counts['Employees']} موظفين، {counts['SalesInvoices']} فواتير بيع، "
                   f"{counts['PurchaseInvoices']} فواتير شراء، {counts['Expenses']} مصروفات")
    remove = "\n".join(f'    db.Execute {vba_str(sql)}, dbFailOnError' for sql in remove_sql())
    values = {
        "COUNTS": counts_text, "PASSWORD": DD.DEMO_PASSWORD, "MARK": DD.DEMO_MARK,
        "SALES": str(n_sales), "PURCHASES": str(n_purchases), "STEP_CALLS": calls, "STEP_PROCS": procs,
        "MASTERS": masters_lines(), "SETTINGS_SQL": settings, "FIRST_BARCODE": DD.PRODUCTS[0].barcode,
        "N_PRODUCTS": str(counts["Products"]), "N_CATEGORIES": str(counts["Categories"]),
        "N_CUSTOMERS": str(counts["Customers"]), "N_SUPPLIERS": str(counts["Suppliers"]),
        "N_EMPLOYEES": str(counts["Employees"]), "N_EXPENSES": str(counts["Expenses"]),
        "SALES_TOTAL": money(r.sales_total), "PURCHASES_TOTAL": money(r.purchases_total),
        "VERIFY": verify_lines(r), "LOW": str(low), "SLOW": str(slow), "REMOVE_SQL": remove,
    }
    text = TEMPLATE
    for k, v in values.items():
        text = text.replace(f"@@{k}@@", v)
    assert not re.search(r"@@[A-Z_]+@@", text), re.findall(r"@@[A-Z_]+@@", text)
    return text
