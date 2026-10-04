Attribute VB_Name = "modBuildQueries"
'==============================================================================
' modBuildQueries  -  Retail Store Management System (Phase 4: Queries)
'
' GENERATED FILE - do not edit by hand.
' Source of truth: tools/queries.py  ->  python3 tools/generate.py
'
' Public procedures (run from the Immediate window, Ctrl+G):
'   BuildQueries     creates or updates every saved query in this front-end.
'   TestQueries      loads a small known business history inside a transaction,
'                    checks the numbers returned by the queries, then rolls back.
'                    Runs only on a database without transactions.
'   DropQueries      DEVELOPMENT ONLY: deletes the system queries.
'
' Requires: modQueryParams (QDate, QLong, SetPeriod), BuildSchema, BuildRelationships.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MSG_RTL As Long = &H180000           ' vbMsgBoxRight + vbMsgBoxRtlReading
Private Const PERIOD_START_DAYS_AGO As Long = 30
Private Const TEST_SLOW_MOVING_DAYS As Long = 90
Private Const QUERY_NAMES As String = "qrySalesDocuments,qrySalesLineItems,qrySalesLinesInPeriod,DailySalesQuery,qrySalesMonthlyDocs,qrySalesMonthlyCost,MonthlySalesQuery,SalesByPeriodQuery,SalesByProductQuery,BestSellingProductsQuery,LeastSellingProductsQuery,qryPurchaseDocuments,PurchasesQuery,qryProductLedger,qryProductLastSale,StockBalanceQuery,LowStockQuery,ProductMovementQuery,SlowMovingProductsQuery,StockByCategoryQuery,StockCou" & _
    "ntQuery,qryCustomerLedger,qryCustomerLedgerTotals,CustomerBalanceQuery,CustomersWithDebtQuery,CustomerStatementQuery,qrySupplierLedger,qrySupplierLedgerTotals,SupplierBalanceQuery,SupplierStatementQuery,ExpensesQuery,ExpensesByTypeQuery,qryProfitSales,qryProfitAdjustments,qryProfitExpenses,ProfitQuery,qryVatOutput,qryVatInputPurchases,qryVatInputExpenses,VatSummaryQuery,DashboardQuery,qryDashboard" & _
    "TopProducts,qrySalesDocPrint,qryPurchaseDocPrint,qryVoucherPrint,qrySalesInvoiceLineTotals,qryPurchaseInvoiceLineTotals,qrySalesReturnedQty,qryPurchaseReturnedQty,IntegrityCheckQuery"

Private m_db As DAO.Database
Private m_created As Long
Private m_updated As Long
Private m_failed As Long
Private m_passed As Long
Private m_report As String
Private m_ids As Collection
Private m_keys As Collection

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildQueries() As Boolean
    m_created = 0: m_updated = 0: m_failed = 0: m_report = ""
    Debug.Print "=== BuildQueries  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Set m_db = CurrentDb

    CreateAllQueries

    m_db.QueryDefs.Refresh
    Application.RefreshDatabaseWindow
    Debug.Print "--- جديدة: " & m_created & " | محدّثة: " & m_updated & " | فشلت: " & m_failed

    If m_failed = 0 Then
        MsgBox "تم إنشاء الاستعلامات بنجاح." & vbCrLf & vbCrLf & _
               "استعلامات جديدة: " & m_created & vbCrLf & _
               "استعلامات محدّثة: " & m_updated & vbCrLf & vbCrLf & _
               "الخطوة التالية: شغّل TestQueries", vbInformation + MSG_RTL, "BuildQueries"
        BuildQueries = True
    Else
        MsgBox "فشل إنشاء " & m_failed & " استعلام:" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "BuildQueries"
    End If
End Function

Public Function TestQueries() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, blocker As String, slowDays As Long

    On Error GoTo EH
    Calendar = vbCalGreg
    m_passed = 0: m_failed = 0: m_report = ""
    Set m_db = CurrentDb

    blocker = ExistingDataTable()
    If Len(blocker) > 0 Then
        MsgBox "يعمل هذا الاختبار على قاعدة بدون حركات فقط، لأن نتائجه أرقام محددة مسبقًا." & _
               vbCrLf & "يوجد بيانات في الجدول: " & blocker, vbExclamation + MSG_RTL, "TestQueries"
        Exit Function
    End If

    Debug.Print "=== TestQueries  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    slowDays = ScalarLong("SELECT SlowMovingDays FROM Settings")
    Set m_ids = New Collection
    Set m_keys = New Collection
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    m_db.Execute "UPDATE [Settings] SET [SlowMovingDays] = " & TEST_SLOW_MOVING_DAYS, dbFailOnError
    LoadFixture
    SetPeriod DateAdd("d", -PERIOD_START_DAYS_AGO, Date), Date

    CheckQueriesOpen
    RunChecks
    RunCorruptionChecks

    ws.Rollback
    inTrans = False
    CleanUpAfterTest slowDays
    ClearQueryParams

    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed & " (تم التراجع عن بيانات الاختبار)"
    If m_failed = 0 Then
        MsgBox "جميع اختبارات الاستعلامات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "الأرقام مطابقة للحسابات اليدوية، ولم تُترك أي بيانات اختبار.", _
               vbInformation + MSG_RTL, "TestQueries"
        TestQueries = True
    Else
        MsgBox "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestQueries"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    If inTrans Then ws.Rollback
    CleanUpAfterTest slowDays
    ClearQueryParams
    MsgBox errText & vbCrLf & "تم التراجع عن بيانات الاختبار.", vbCritical + MSG_RTL, "TestQueries"
End Function

Public Sub DropQueries()
    ' DEVELOPMENT ONLY - deletes the queries created by BuildQueries.
    Dim db As DAO.Database, qName As Variant, removed As Long
    If InputBox("سيتم حذف استعلامات النظام (البيانات لا تُحذف)." & vbCrLf & _
                "للتأكيد اكتب DELETE", "DropQueries") <> "DELETE" Then Exit Sub
    Set db = CurrentDb
    For Each qName In Split(QUERY_NAMES, ",")
        If QueryExists(db, CStr(qName)) Then
            db.QueryDefs.Delete CStr(qName)
            removed = removed + 1
        End If
    Next
    Application.RefreshDatabaseWindow
    MsgBox "تم حذف " & removed & " استعلام.", vbInformation + MSG_RTL
End Sub

'------------------------------------------------------------------------------
' Building
'------------------------------------------------------------------------------
Private Sub SaveQuery(ByVal QueryName As String, ByVal Description As String, ByVal Sql As String)
    Dim qdf As DAO.QueryDef
    On Error GoTo EH
    If QueryExists(m_db, QueryName) Then
        Set qdf = m_db.QueryDefs(QueryName)
        qdf.SQL = Sql
        m_updated = m_updated + 1
        Debug.Print "  ~ " & QueryName
    Else
        Set qdf = m_db.CreateQueryDef(QueryName, Sql)
        m_created = m_created + 1
        Debug.Print "  + " & QueryName
    End If
    SetProp qdf, "Description", dbText, Description
    Exit Sub
EH:
    Fail QueryName & ": خطأ " & Err.Number & " - " & Err.Description
End Sub

Private Sub SetProp(ByVal obj As Object, ByVal PropName As String, _
                    ByVal PropType As Integer, ByVal PropValue As Variant)
    Dim prp As Object
    On Error Resume Next
    Set prp = obj.Properties(PropName)
    On Error GoTo 0
    If prp Is Nothing Then
        obj.Properties.Append obj.CreateProperty(PropName, PropType, PropValue)
    Else
        prp.Value = PropValue
    End If
End Sub

'------------------------------------------------------------------------------
' Testing
'------------------------------------------------------------------------------
Private Function ExistingDataTable() As String
    Dim t As Variant
    For Each t In Array("SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns", _
                        "CustomerPayments", "SupplierPayments", "Expenses", _
                        "InventoryTransactions", "StockCounts", "Products", "Suppliers")
        If ScalarLong("SELECT COUNT(*) FROM [" & t & "]") > 0 Then
            ExistingDataTable = CStr(t)
            Exit Function
        End If
    Next
    If ScalarLong("SELECT COUNT(*) FROM [Customers]") > 1 Then ExistingDataTable = "Customers"
End Function

Private Sub CheckQueriesOpen()
    Dim qName As Variant, rs As DAO.Recordset, opened As Long
    SetQueryParam "CustomerID", CLng(R("C2"))
    SetQueryParam "SupplierID", CLng(R("S1"))
    SetQueryParam "ProductID", CLng(R("P2"))
    For Each qName In Split(QUERY_NAMES, ",")
        On Error Resume Next
        Set rs = m_db.OpenRecordset("SELECT * FROM [" & qName & "]", dbOpenSnapshot)
        If Err.Number <> 0 Then
            Fail "الاستعلام " & qName & " لا يعمل: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            rs.Close
            opened = opened + 1
        End If
        On Error GoTo 0
    Next
    Call Record(opened = UBound(Split(QUERY_NAMES, ",")) + 1, _
                "كل الاستعلامات (" & opened & ") تعمل بدون أخطاء")
End Sub

Private Sub Chk(ByVal Label As String, ByVal Sql As String, ByVal Expected As Double)
    Dim rs As DAO.Recordset, actual As Variant
    On Error GoTo EH
    Set rs = m_db.OpenRecordset(Sql, dbOpenSnapshot)
    If rs.EOF Then
        actual = Null
    Else
        actual = rs(0).Value
    End If
    rs.Close
    If IsNull(actual) Then
        Fail Label & ": لا توجد نتيجة (المتوقع " & Expected & ")"
    ElseIf Abs(CDbl(actual) - Expected) < 0.005 Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label & " = " & actual
    Else
        Fail Label & ": النتيجة " & actual & " والمتوقع " & Expected
    End If
    Exit Sub
EH:
    Fail Label & ": خطأ " & Err.Number & " - " & Err.Description
End Sub

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

Private Sub Ins(ByVal Key As String, ByVal TableName As String, ByVal PkField As String, _
                ByVal Sql As String)
    m_db.Execute Sql, dbFailOnError
    m_ids.Add ScalarLong("SELECT @@IDENTITY"), Key
    m_keys.Add TableName & "|" & PkField & "|" & Key
End Sub

Private Function R(ByVal Key As String) As String
    R = CStr(m_ids(Key))
End Function

Private Function D(ByVal DaysAgo As Long, ByVal HourOfDay As Long) As String
    ' Access SQL date literal, always Gregorian (Saudi PCs may default to Hijri).
    D = "#" & Format$(DateAdd("d", -DaysAgo, Date) + TimeSerial(HourOfDay, 0, 0), _
                      "yyyy-mm-dd hh:nn:ss") & "#"
End Function

Private Sub CleanUpAfterTest(ByVal SlowMovingDays As Long)
    ' Safety net: normally the rollback already removed everything.
    Dim i As Long, parts() As String
    On Error Resume Next
    If m_keys Is Nothing Then Exit Sub
    If ScalarLong("SELECT COUNT(*) FROM [Products] WHERE Left([ProductCode], 5) = 'TEST-'") > 0 Then
        Debug.Print "تنبيه: لم يشمل التراجع كل الجداول، يتم حذف بيانات الاختبار يدويًا."
        For i = m_keys.Count To 1 Step -1
            parts = Split(m_keys(i), "|")
            m_db.Execute "DELETE FROM [" & parts(0) & "] WHERE [" & parts(1) & "] = " & _
                         m_ids(parts(2)), dbFailOnError
        Next
    End If
    If SlowMovingDays > 0 Then
        m_db.Execute "UPDATE [Settings] SET [SlowMovingDays] = " & SlowMovingDays, dbFailOnError
    End If
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function QueryExists(ByVal db As DAO.Database, ByVal QueryName As String) As Boolean
    Dim qdf As DAO.QueryDef
    For Each qdf In db.QueryDefs
        If StrComp(qdf.Name, QueryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next
End Function

Private Function ScalarLong(ByVal Sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = m_db.OpenRecordset(Sql, dbOpenSnapshot)
    If Not rs.EOF Then ScalarLong = Nz(rs(0), 0)
    rs.Close
End Function

'------------------------------------------------------------------------------
' Generated: test fixture, checks and corruption checks
'------------------------------------------------------------------------------
Private Sub LoadFixture()
    Ins "S1", "Suppliers", "SupplierID", _
        "INSERT INTO [Suppliers] ([SupplierName], [OpeningBalance], [CurrentBalance], [CreatedAt]) VALUES ('TEST مورد', 0, 2855, " & D(120, 0) & ")"
    Ins "C2", "Customers", "CustomerID", _
        "INSERT INTO [Customers] ([CustomerName], [OpeningBalance], [CurrentBalance], [AllowCredit], [CreatedAt]) VALUES ('TEST عميل آجل', 50, 164, True, " & D(90, 0) & ")"
    Ins "CAT2", "Categories", "CategoryID", _
        "INSERT INTO [Categories] ([CategoryName]) VALUES ('TEST مواد غذائية')"
    Ins "P1", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [SupplierID], [CreatedAt]) VALUES ('TEST-P1', 'TEST منتج 1', 1, 1, 60, 60, 115, 83, 5, " & R("S1") & ", " & D(120, 0) & ")"
    Ins "P2", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [SupplierID], [CreatedAt]) VALUES ('TEST-P2', 'TEST منتج 2', " & R("CAT2") & ", 1, 10, 10, 23, 186, 200, " & R("S1") & ", " & D(120, 0) & ")"
    Ins "P3", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [CreatedAt]) VALUES ('TEST-P3', 'TEST منتج 3', 1, 1, 10, 10, 20, 20, 0, " & D(120, 0) & ")"
    Ins "T1", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(60, 8) & ", " & R("P3") & ", 8, 20, 10, 20, 'MANUAL', Null, 'TEST-OPEN', 1)"
    Ins "PUR1", "PurchaseInvoices", "PurchaseInvoiceID", _
        "INSERT INTO [PurchaseInvoices] ([InvoiceNumber], [SupplierInvoiceNo], [InvoiceDate], [SupplierID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount]) VALUES ('TEST-PUR-1', 'S-100', " & D(50, 9) & ", " & R("S1") & ", 1, 'CREDIT', 1, 8000, 0, 8000, 1200, 9200, 5000, 4200)"
    Ins "PUR1L1", "PurchaseInvoiceDetails", "PurchaseDetailID", _
        "INSERT INTO [PurchaseInvoiceDetails] ([PurchaseInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PUR1") & ", 1, " & R("P1") & ", 100, 60, 0, 6000, 0.15, 900, 6900)"
    Ins "PUR1L2", "PurchaseInvoiceDetails", "PurchaseDetailID", _
        "INSERT INTO [PurchaseInvoiceDetails] ([PurchaseInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PUR1") & ", 2, " & R("P2") & ", 200, 10, 0, 2000, 0.15, 300, 2300)"
    Ins "T2", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(50, 9) & ", " & R("P1") & ", 1, 100, 60, 100, 'PURCHASE', " & R("PUR1") & ", 'TEST-PUR-1', 1)"
    Ins "T3", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(50, 9) & ", " & R("P2") & ", 1, 200, 10, 200, 'PURCHASE', " & R("PUR1") & ", 'TEST-PUR-1', 1)"
    Ins "INV0", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue]) VALUES ('TEST-INV-0', " & D(45, 10) & ", 1, 1, 'CASH', 1, 100, 0, 100, 15, 115, 115, 0, 115, 0)"
    Ins "INV0L1", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV0") & ", 1, " & R("P2") & ", 5, 20, 0, 100, 0.15, 15, 115, 10)"
    Ins "T4", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(45, 10) & ", " & R("P2") & ", 2, -5, 10, 195, 'SALE', " & R("INV0") & ", 'TEST-INV-0', 1)"
    Ins "INV1", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue]) VALUES ('TEST-INV-1', " & D(10, 11) & ", 1, 1, 'CASH', 1, 1000, 0, 1000, 150, 1150, 1150, 0, 1200, 50)"
    Ins "INV1L1", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV1") & ", 1, " & R("P1") & ", 10, 100, 0, 1000, 0.15, 150, 1150, 60)"
    Ins "T5", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(10, 11) & ", " & R("P1") & ", 2, -10, 60, 90, 'SALE', " & R("INV1") & ", 'TEST-INV-1', 1)"
    Ins "INV2", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue]) VALUES ('TEST-INV-2', " & D(5, 12) & ", " & R("C2") & ", 1, 'CREDIT', 1, 400, 0, 400, 60, 460, 100, 360, 100, 0)"
    Ins "INV2L1", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV2") & ", 1, " & R("P1") & ", 2, 100, 0, 200, 0.15, 30, 230, 60)"
    Ins "INV2L2", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV2") & ", 2, " & R("P2") & ", 10, 20, 0, 200, 0.15, 30, 230, 10)"
    Ins "T6", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(5, 12) & ", " & R("P1") & ", 2, -2, 60, 88, 'SALE', " & R("INV2") & ", 'TEST-INV-2', 1)"
    Ins "T7", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(5, 12) & ", " & R("P2") & ", 2, -10, 10, 185, 'SALE', " & R("INV2") & ", 'TEST-INV-2', 1)"
    Ins "PAY1", "SupplierPayments", "PaymentID", _
        "INSERT INTO [SupplierPayments] ([PaymentNumber], [SupplierID], [PaymentDate], [Amount], [PaymentMethodID], [EmployeeID]) VALUES ('TEST-PAY-1', " & R("S1") & ", " & D(4, 10) & ", 1000, 3, 1)"
    Ins "RCV1", "CustomerPayments", "PaymentID", _
        "INSERT INTO [CustomerPayments] ([PaymentNumber], [CustomerID], [PaymentDate], [Amount], [PaymentMethodID], [EmployeeID]) VALUES ('TEST-RCV-1', " & R("C2") & ", " & D(3, 10) & ", 200, 1, 1)"
    Ins "CRN1", "SalesReturns", "SalesReturnID", _
        "INSERT INTO [SalesReturns] ([ReturnNumber], [ReturnDate], [SalesInvoiceID], [CustomerID], [EmployeeID], [Reason], [RefundType], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [RefundedAmount]) VALUES ('TEST-CRN-1', " & D(2, 13) & ", " & R("INV2") & ", " & R("C2") & ", 1, 'TEST إرجاع العميل', 'CREDIT', 40, 0, 40, 6, 46, 0)"
    Ins "CRN1L1", "SalesReturnDetails", "ReturnDetailID", _
        "INSERT INTO [SalesReturnDetails] ([SalesReturnID], [SalesDetailID], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost], [ReturnToStock]) VALUES (" & R("CRN1") & ", " & R("INV2L2") & ", " & R("P2") & ", 2, 20, 0, 40, 0.15, 6, 46, 10, True)"
    Ins "T8", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(2, 13) & ", " & R("P2") & ", 4, 2, 10, 187, 'SALES_RETURN', " & R("CRN1") & ", 'TEST-CRN-1', 1)"
    Ins "PRT1", "PurchaseReturns", "PurchaseReturnID", _
        "INSERT INTO [PurchaseReturns] ([ReturnNumber], [ReturnDate], [PurchaseInvoiceID], [SupplierID], [EmployeeID], [Reason], [RefundType], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [RefundedAmount]) VALUES ('TEST-PRT-1', " & D(1, 10) & ", " & R("PUR1") & ", " & R("S1") & ", 1, 'TEST عيب مصنعي', 'CREDIT', 300, 0, 300, 45, 345, 0)"
    Ins "PRT1L1", "PurchaseReturnDetails", "PurchaseReturnDetailID", _
        "INSERT INTO [PurchaseReturnDetails] ([PurchaseReturnID], [PurchaseDetailID], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PRT1") & ", " & R("PUR1L1") & ", " & R("P1") & ", 5, 60, 0, 300, 0.15, 45, 345)"
    Ins "T9", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(1, 10) & ", " & R("P1") & ", 3, -5, 60, 83, 'PURCHASE_RETURN', " & R("PRT1") & ", 'TEST-PRT-1', 1)"
    Ins "T10", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(1, 15) & ", " & R("P2") & ", 6, -1, 10, 186, 'MANUAL', Null, 'TEST-ADJ-1', 1)"
    Ins "CNT1", "StockCounts", "StockCountID", _
        "INSERT INTO [StockCounts] ([CountNumber], [CountDate], [Status], [EmployeeID]) VALUES ('TEST-CNT-1', " & D(1, 18) & ", 'OPEN', 1)"
    Ins "CNT1L1", "StockCountDetails", "StockCountDetailID", _
        "INSERT INTO [StockCountDetails] ([StockCountID], [ProductID], [SystemQuantity], [ActualQuantity], [Difference], [UnitCost], [DifferenceValue]) VALUES (" & R("CNT1") & ", " & R("P1") & ", 83, 80, -3, 60, -180)"
    Ins "CNT1L2", "StockCountDetails", "StockCountDetailID", _
        "INSERT INTO [StockCountDetails] ([StockCountID], [ProductID], [SystemQuantity], [ActualQuantity], [Difference], [UnitCost], [DifferenceValue]) VALUES (" & R("CNT1") & ", " & R("P2") & ", 186, Null, 0, 10, 0)"
    Ins "EXP1", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-1', " & D(6, 0) & ", 2, 200, 30, 230, 1, 'TEST كهرباء', 1)"
    Ins "EXP2", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-2', " & D(40, 0) & ", 1, 1000, 0, 1000, 3, 'TEST إيجار', 1)"
End Sub

Private Sub RunChecks()
    Chk "رصيد المنتج 1 = 100 شراء - 10 - 2 بيع - 5 مرتجع شراء = 83", _
        "SELECT CurrentQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 83
    Chk "رصيد المنتج 1 من دفتر الحركات = 83", _
        "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 83
    Chk "رصيد المنتج 2 = 200 - 5 - 10 + 2 - 1 = 186", _
        "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P2"), 186
    Chk "لا توجد فروقات بين الكمية المسجلة والحركات", _
        "SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0", 0
    Chk "قيمة مخزون المنتج 1 بالتكلفة = 83 × 60", _
        "SELECT StockCostValue FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 4980
    Chk "منخفض المخزون: منتج واحد فقط", _
        "SELECT COUNT(*) FROM LowStockQuery", 1
    Chk "منخفض المخزون: المنتج 2 (186 <= 200)", _
        "SELECT ShortageQty FROM LowStockQuery WHERE ProductID = " & R("P2"), 14
    Chk "طباعة فاتورة الشراء: سطران", _
        "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'PURCHASE' AND DocID = " & R("PUR1"), 2
    Chk "طباعة مرتجع الشراء: رقم السطر الأصلي ورقم الفاتورة الأصلية", _
        "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'RETURN' AND DocID = " & R("PRT1") & " AND LineNumber = 1 AND OriginalNumber = 'TEST-PUR-1' AND RemainingAmount = 345", 1
    Chk "طباعة سند الصرف: المبلغ وطريقة الدفع", _
        "SELECT Amount FROM qryVoucherPrint WHERE DocKind = 'PAYMENT' AND DocID = " & R("PAY1"), 1000
    Chk "طباعة سند القبض", _
        "SELECT COUNT(*) FROM qryVoucherPrint WHERE DocKind = 'RECEIPT' AND DocID = " & R("RCV1"), 1
    Chk "الجرد المفتوح: عجز المنتج 1 = -3 × 60", _
        "SELECT DifferenceValue FROM StockCountQuery WHERE ProductID = " & R("P1"), -180
    Chk "الجرد المفتوح: صنف واحد لم يُعدّ بعد", _
        "SELECT COUNT(*) FROM StockCountQuery WHERE ActualQuantity Is Null", 1
    Chk "غير المتحركة: منتج واحد", _
        "SELECT COUNT(*) FROM SlowMovingProductsQuery", 1
    Chk "غير المتحركة: المنتج 3 بقيمة 200", _
        "SELECT StockCostValue FROM SlowMovingProductsQuery WHERE ProductID = " & R("P3"), 200
    Chk "المخزون حسب التصنيف: تصنيف عام = 83 + 20", _
        "SELECT TotalQuantity FROM StockByCategoryQuery WHERE CategoryID = 1", 103
    Chk "المخزون حسب التصنيف: قيمة تصنيف عام = 4980 + 200", _
        "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = 1", 5180
    Chk "المخزون حسب التصنيف: قيمة التصنيف 2 = 186 × 10", _
        "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = " & R("CAT2"), 1860
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "حركة المنتج 2: رصيد أول المدة = 195", _
        "SELECT NetQty FROM ProductMovementQuery WHERE SortKey = 0", 195
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "حركة المنتج 2: 3 حركات + سطر الرصيد السابق", _
        "SELECT COUNT(*) FROM ProductMovementQuery", 4
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "حركة المنتج 2: الرصيد الختامي = 186", _
        "SELECT Sum(NetQty) FROM ProductMovementQuery", 186
    Chk "المبيعات اليومية: يوم الفاتورة الآجلة = 460", _
        "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(5, 0) & ")", 460
    Chk "المبيعات اليومية: الآجل في نفس اليوم = 460", _
        "SELECT CreditSales FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(5, 0) & ")", 460
    Chk "المبيعات اليومية: يوم الفاتورة النقدية = 1150 نقدًا", _
        "SELECT CashSales FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(10, 0) & ")", 1150
    Chk "المبيعات اليومية: يوم المرتجع = -46", _
        "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(2, 0) & ")", -46
    Chk "المبيعات الشهرية: صافي كل الأشهر = 100 + 1000 + 400 - 40", _
        "SELECT Sum(NetSalesExVAT) FROM MonthlySalesQuery", 1460
    Chk "المبيعات الشهرية: مجمل ربح كل الأشهر = 1460 - 850", _
        "SELECT Sum(GrossProfit) FROM MonthlySalesQuery", 610
    Chk "المبيعات خلال الفترة: فاتورتان ومرتجع", _
        "SELECT COUNT(*) FROM SalesByPeriodQuery", 3
    Chk "المبيعات خلال الفترة: الإجمالي = 1150 + 460 - 46", _
        "SELECT Sum(GrossAmount) FROM SalesByPeriodQuery", 1564
    Chk "المبيعات حسب المنتج: المنتج 1 صافي كمية 12", _
        "SELECT NetQty FROM SalesByProductQuery WHERE ProductID = " & R("P1"), 12
    Chk "المبيعات حسب المنتج: ربح المنتج 1 = 1200 - 720", _
        "SELECT GrossProfit FROM SalesByProductQuery WHERE ProductID = " & R("P1"), 480
    Chk "المبيعات حسب المنتج: المنتج 2 مرتجع 2", _
        "SELECT QtyReturned FROM SalesByProductQuery WHERE ProductID = " & R("P2"), 2
    Chk "المبيعات حسب المنتج: صافي مبيعات المنتج 2 = 200 - 40", _
        "SELECT NetSales FROM SalesByProductQuery WHERE ProductID = " & R("P2"), 160
    Chk "الأكثر مبيعًا: المنتج 1", _
        "SELECT TOP 1 ProductID FROM BestSellingProductsQuery ORDER BY NetQty DESC, NetSales DESC", CDbl(R("P1"))
    Chk "الأقل مبيعًا: المنتج 3 (لم يُبع)", _
        "SELECT TOP 1 ProductID FROM LeastSellingProductsQuery ORDER BY NetQtySold, ProductName", CDbl(R("P3"))
    Chk "المشتريات خلال الفترة: مرتجع الشراء فقط", _
        "SELECT COUNT(*) FROM PurchasesQuery", 1
    Chk "المشتريات خلال الفترة: -345", _
        "SELECT Sum(GrossAmount) FROM PurchasesQuery", -345
    Chk "رصيد العميل الآجل = 50 + 460 - 100 - 46 - 200", _
        "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = " & R("C2"), 164
    Chk "رصيد العميل النقدي = 0", _
        "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = 1", 0
    Chk "العملاء المدينون: عميل واحد", _
        "SELECT COUNT(*) FROM CustomersWithDebtQuery", 1
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "كشف العميل: الرصيد السابق = 50", _
        "SELECT Debit FROM CustomerStatementQuery WHERE SortKey = 0", 50
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "كشف العميل: رصيد سابق + 3 حركات", _
        "SELECT COUNT(*) FROM CustomerStatementQuery", 4
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "كشف العميل: الرصيد الختامي = 164", _
        "SELECT Sum(Debit) - Sum(Credit) FROM CustomerStatementQuery", 164
    Chk "رصيد المورد = 9200 - 5000 - 1000 - 345", _
        "SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = " & R("S1"), 2855
    SetQueryParam "SupplierID", CLng(R("S1"))
    Chk "كشف المورد: الرصيد السابق = 4200", _
        "SELECT Credit FROM SupplierStatementQuery WHERE SortKey = 0", 4200
    SetQueryParam "SupplierID", CLng(R("S1"))
    Chk "كشف المورد: الرصيد الختامي = 2855", _
        "SELECT Sum(Credit) - Sum(Debit) FROM SupplierStatementQuery", 2855
    Chk "المصروفات خلال الفترة: مصروف واحد", _
        "SELECT COUNT(*) FROM ExpensesQuery", 1
    Chk "المصروفات خلال الفترة: 230 شامل الضريبة", _
        "SELECT Sum(TotalAmount) FROM ExpensesQuery", 230
    Chk "المصروفات حسب النوع: الكهرباء 200", _
        "SELECT AmountExVAT FROM ExpensesByTypeQuery", 200
    Chk "الأرباح: صافي المبيعات = 1000 + 400 - 40", _
        "SELECT NetSales FROM ProfitQuery", 1360
    Chk "الأرباح: تكلفة المبيعات = 600 + 220 - 20", _
        "SELECT CostOfSales FROM ProfitQuery", 800
    Chk "الأرباح: مجمل الربح = 1360 - 800", _
        "SELECT GrossProfit FROM ProfitQuery", 560
    Chk "الأرباح: فروقات المخزون = -10 (منتج تالف)", _
        "SELECT InventoryAdjustments FROM ProfitQuery", -10
    Chk "الأرباح: المصروفات = 200", _
        "SELECT TotalExpenses FROM ProfitQuery", 200
    Chk "الأرباح: صافي الربح = 560 - 10 - 200", _
        "SELECT NetProfit FROM ProfitQuery", 350
    Chk "الضريبة: ضريبة المخرجات = 150 + 60 - 6", _
        "SELECT OutputVAT FROM VatSummaryQuery", 204
    Chk "الضريبة: ضريبة المدخلات = -45 (مرتجع شراء) + 30 (مصروف)", _
        "SELECT InputVAT FROM VatSummaryQuery", -15
    Chk "الضريبة: الصافي المستحق = 204 + 15", _
        "SELECT NetVATDue FROM VatSummaryQuery", 219
    Chk "طباعة الفاتورة الآجلة: سطران", _
        "SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = " & R("INV2"), 2
    Chk "طباعة الفاتورة الآجلة: مجموع الأسطر = 460", _
        "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = " & R("INV2"), 460
    Chk "طباعة الإشعار الدائن: سطر واحد بقيمة 46", _
        "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'RETURN' AND DocID = " & R("CRN1"), 46
    Chk "فحص السلامة: لا توجد مشكلات", _
        "SELECT COUNT(*) FROM IntegrityCheckQuery", 0
End Sub

Private Sub RunCorruptionChecks()
    m_db.Execute "UPDATE [Products] SET [CurrentQuantity] = 80 WHERE [ProductID] = " & R("P1"), dbFailOnError
    Chk "فحص السلامة يكتشف تلاعبًا بالكمية", _
        "SELECT COUNT(*) FROM IntegrityCheckQuery WHERE IssueCode = 'STOCK_MISMATCH'", 1
    m_db.Execute "UPDATE [Customers] SET [CurrentBalance] = 0 WHERE [CustomerID] = " & R("C2"), dbFailOnError
    Chk "فحص السلامة يكتشف رصيد عميل خاطئ", _
        "SELECT ExpectedValue FROM IntegrityCheckQuery WHERE IssueCode = 'CUSTOMER_BALANCE'", 164
    m_db.Execute "UPDATE [SalesReturns] SET [CustomerID] = 1 WHERE [SalesReturnID] = " & R("CRN1"), dbFailOnError
    Chk "فحص السلامة يكتشف مرتجعًا لعميل مختلف", _
        "SELECT COUNT(*) FROM IntegrityCheckQuery WHERE IssueCode = 'RETURN_CUSTOMER'", 1
End Sub

'------------------------------------------------------------------------------
' Generated: query definitions (in dependency order)
'------------------------------------------------------------------------------
Private Sub CreateAllQueries()
    Q_qrySalesDocuments
    Q_qrySalesLineItems
    Q_qrySalesLinesInPeriod
    Q_DailySalesQuery
    Q_qrySalesMonthlyDocs
    Q_qrySalesMonthlyCost
    Q_MonthlySalesQuery
    Q_SalesByPeriodQuery
    Q_SalesByProductQuery
    Q_BestSellingProductsQuery
    Q_LeastSellingProductsQuery
    Q_qryPurchaseDocuments
    Q_PurchasesQuery
    Q_qryProductLedger
    Q_qryProductLastSale
    Q_StockBalanceQuery
    Q_LowStockQuery
    Q_ProductMovementQuery
    Q_SlowMovingProductsQuery
    Q_StockByCategoryQuery
    Q_StockCountQuery
    Q_qryCustomerLedger
    Q_qryCustomerLedgerTotals
    Q_CustomerBalanceQuery
    Q_CustomersWithDebtQuery
    Q_CustomerStatementQuery
    Q_qrySupplierLedger
    Q_qrySupplierLedgerTotals
    Q_SupplierBalanceQuery
    Q_SupplierStatementQuery
    Q_ExpensesQuery
    Q_ExpensesByTypeQuery
    Q_qryProfitSales
    Q_qryProfitAdjustments
    Q_qryProfitExpenses
    Q_ProfitQuery
    Q_qryVatOutput
    Q_qryVatInputPurchases
    Q_qryVatInputExpenses
    Q_VatSummaryQuery
    Q_DashboardQuery
    Q_qryDashboardTopProducts
    Q_qrySalesDocPrint
    Q_qryPurchaseDocPrint
    Q_qryVoucherPrint
    Q_qrySalesInvoiceLineTotals
    Q_qryPurchaseInvoiceLineTotals
    Q_qrySalesReturnedQty
    Q_qryPurchaseReturnedQty
    Q_IntegrityCheckQuery
End Sub

Private Sub Q_qrySalesDocuments()
    Dim s As String
    s = "SELECT 'SALE' AS DocType, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.InvoiceDate AS DocDate, h.CustomerID, h.EmployeeID, h.PaymentType," & vbCrLf
    s = s & "       h.TaxableAmount AS NetAmount, h.Tax AS VATAmount, h.TotalAmount AS GrossAmount," & vbCrLf
    s = s & "       h.PaidAmount AS SettledAmount, h.RemainingAmount AS OnAccount" & vbCrLf
    s = s & "FROM SalesInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, r.EmployeeID," & vbCrLf
    s = s & "       r.RefundType, -r.TaxableAmount, -r.Tax, -r.TotalAmount, -r.RefundedAmount," & vbCrLf
    s = s & "       -(r.TotalAmount - r.RefundedAmount)" & vbCrLf
    s = s & "FROM SalesReturns AS r" & vbCrLf
    SaveQuery "qrySalesDocuments", "مستندات البيع: الفواتير (+) والمرتجعات (-) بقيم موقّعة", s
End Sub

Private Sub Q_qrySalesLineItems()
    Dim s As String
    s = "SELECT 'SALE' AS DocType, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.InvoiceDate AS DocDate, h.CustomerID, d.ProductID, d.Quantity AS SignedQty," & vbCrLf
    s = s & "       d.Quantity AS SoldQty, CCur(0) AS ReturnedQty, d.NetAmount AS LineNet," & vbCrLf
    s = s & "       d.Tax AS LineVAT, d.LineTotal AS LineGross, d.Quantity * d.UnitCost AS LineCost" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d" & vbCrLf
    s = s & "     ON h.SalesInvoiceID = d.SalesInvoiceID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, d.ProductID," & vbCrLf
    s = s & "       -d.Quantity, CCur(0), d.Quantity, -d.NetAmount, -d.Tax, -d.LineTotal," & vbCrLf
    s = s & "       IIf(d.ReturnToStock, -d.Quantity * d.UnitCost, 0)" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN SalesReturnDetails AS d" & vbCrLf
    s = s & "     ON r.SalesReturnID = d.SalesReturnID" & vbCrLf
    SaveQuery "qrySalesLineItems", "أسطر البيع والمرتجعات مع التكلفة (أساس تحليل المنتجات والأرباح)", s
End Sub

Private Sub Q_qrySalesLinesInPeriod()
    Dim s As String
    s = "SELECT * FROM qrySalesLineItems" & vbCrLf
    s = s & "WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qrySalesLinesInPeriod", "أسطر البيع والمرتجعات داخل الفترة", s
End Sub

Private Sub Q_DailySalesQuery()
    Dim s As String
    s = "SELECT DateValue(DocDate) AS SaleDate," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'SALE', 1, 0)) AS InvoiceCount," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'RETURN', 1, 0)) AS ReturnCount," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'SALE', NetAmount, 0)) AS SalesExVAT," & vbCrLf
    s = s & "       -Sum(IIf(DocType = 'RETURN', NetAmount, 0)) AS ReturnsExVAT," & vbCrLf
    s = s & "       Sum(NetAmount) AS NetSalesExVAT, Sum(VATAmount) AS NetVAT," & vbCrLf
    s = s & "       Sum(GrossAmount) AS NetSalesTotal," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'SALE' AND PaymentType = 'CASH', GrossAmount, 0)) AS CashSales," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'SALE' AND PaymentType = 'CREDIT', GrossAmount, 0)) AS CreditSales," & vbCrLf
    s = s & "       Sum(SettledAmount) AS CollectedAmount" & vbCrLf
    s = s & "FROM qrySalesDocuments" & vbCrLf
    s = s & "GROUP BY DateValue(DocDate)" & vbCrLf
    s = s & "ORDER BY DateValue(DocDate) DESC" & vbCrLf
    SaveQuery "DailySalesQuery", "المبيعات اليومية: عدد الفواتير والمرتجعات والصافي والنقدي والآجل", s
End Sub

Private Sub Q_qrySalesMonthlyDocs()
    Dim s As String
    s = "SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'SALE', 1, 0)) AS InvoiceCount," & vbCrLf
    s = s & "       Sum(IIf(DocType = 'RETURN', 1, 0)) AS ReturnCount," & vbCrLf
    s = s & "       Sum(NetAmount) AS NetSalesExVAT, Sum(VATAmount) AS NetVAT," & vbCrLf
    s = s & "       Sum(GrossAmount) AS NetSalesTotal" & vbCrLf
    s = s & "FROM qrySalesDocuments" & vbCrLf
    s = s & "GROUP BY Year(DocDate), Month(DocDate)" & vbCrLf
    SaveQuery "qrySalesMonthlyDocs", "تجميع شهري لمستندات البيع", s
End Sub

Private Sub Q_qrySalesMonthlyCost()
    Dim s As String
    s = "SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth, Sum(LineCost) AS MonthCost" & vbCrLf
    s = s & "FROM qrySalesLineItems" & vbCrLf
    s = s & "GROUP BY Year(DocDate), Month(DocDate)" & vbCrLf
    SaveQuery "qrySalesMonthlyCost", "تكلفة البضاعة المباعة شهريًا", s
End Sub

Private Sub Q_MonthlySalesQuery()
    Dim s As String
    s = "SELECT d.SalesYear, d.SalesMonth, d.InvoiceCount, d.ReturnCount, d.NetSalesExVAT," & vbCrLf
    s = s & "       d.NetVAT, d.NetSalesTotal, CCur(Nz(c.MonthCost, 0)) AS CostOfSales," & vbCrLf
    s = s & "       d.NetSalesExVAT - CCur(Nz(c.MonthCost, 0)) AS GrossProfit" & vbCrLf
    s = s & "FROM qrySalesMonthlyDocs AS d LEFT JOIN qrySalesMonthlyCost AS c" & vbCrLf
    s = s & "     ON (d.SalesYear = c.SalesYear AND d.SalesMonth = c.SalesMonth)" & vbCrLf
    s = s & "ORDER BY d.SalesYear DESC, d.SalesMonth DESC" & vbCrLf
    SaveQuery "MonthlySalesQuery", "المبيعات الشهرية مع التكلفة ومجمل الربح", s
End Sub

Private Sub Q_SalesByPeriodQuery()
    Dim s As String
    s = "SELECT s.DocType, s.DocID, s.DocNumber, s.DocDate, s.CustomerID, c.CustomerName," & vbCrLf
    s = s & "       e.EmployeeName, s.PaymentType, s.NetAmount, s.VATAmount, s.GrossAmount," & vbCrLf
    s = s & "       s.SettledAmount, s.OnAccount" & vbCrLf
    s = s & "FROM (qrySalesDocuments AS s INNER JOIN Customers AS c ON s.CustomerID = c.CustomerID)" & vbCrLf
    s = s & "     INNER JOIN Employees AS e ON s.EmployeeID = e.EmployeeID" & vbCrLf
    s = s & "WHERE s.DocDate >= QDate('PeriodStart') AND s.DocDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY s.DocDate" & vbCrLf
    SaveQuery "SalesByPeriodQuery", "فواتير ومرتجعات البيع خلال فترة مع العميل والكاشير", s
End Sub

Private Sub Q_SalesByProductQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName," & vbCrLf
    s = s & "       Sum(l.SoldQty) AS QtySold, Sum(l.ReturnedQty) AS QtyReturned, Sum(l.SignedQty) AS NetQty," & vbCrLf
    s = s & "       Sum(l.LineNet) AS NetSales, Sum(l.LineVAT) AS SalesVAT, Sum(l.LineGross) AS SalesTotal," & vbCrLf
    s = s & "       Sum(l.LineCost) AS CostOfSales, Sum(l.LineNet) - Sum(l.LineCost) AS GrossProfit" & vbCrLf
    s = s & "FROM (qrySalesLinesInPeriod AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID)" & vbCrLf
    s = s & "     INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID" & vbCrLf
    s = s & "GROUP BY p.ProductID, p.ProductCode, p.ProductName, c.CategoryName" & vbCrLf
    s = s & "ORDER BY p.ProductName" & vbCrLf
    SaveQuery "SalesByProductQuery", "المبيعات حسب المنتج خلال فترة (كمية، صافي، تكلفة، ربح)", s
End Sub

Private Sub Q_BestSellingProductsQuery()
    Dim s As String
    s = "SELECT * FROM SalesByProductQuery" & vbCrLf
    s = s & "ORDER BY NetQty DESC, NetSales DESC" & vbCrLf
    SaveQuery "BestSellingProductsQuery", "أفضل المنتجات مبيعًا خلال فترة (حسب صافي الكمية)", s
End Sub

Private Sub Q_LeastSellingProductsQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity," & vbCrLf
    s = s & "       CCur(Nz(s.NetQty, 0)) AS NetQtySold, CCur(Nz(s.NetSales, 0)) AS NetSalesAmount" & vbCrLf
    s = s & "FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)" & vbCrLf
    s = s & "     LEFT JOIN SalesByProductQuery AS s ON p.ProductID = s.ProductID" & vbCrLf
    s = s & "WHERE p.IsActive = True" & vbCrLf
    s = s & "ORDER BY CCur(Nz(s.NetQty, 0)), p.ProductName" & vbCrLf
    SaveQuery "LeastSellingProductsQuery", "أقل المنتجات مبيعًا خلال فترة (تشمل المنتجات التي لم تُبع)", s
End Sub

Private Sub Q_qryPurchaseDocuments()
    Dim s As String
    s = "SELECT 'PURCHASE' AS DocType, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.SupplierInvoiceNo AS SupplierRef, h.InvoiceDate AS DocDate, h.SupplierID," & vbCrLf
    s = s & "       h.PaymentType, h.TaxableAmount AS NetAmount, h.Tax AS VATAmount," & vbCrLf
    s = s & "       h.TotalAmount AS GrossAmount, h.PaidAmount AS SettledAmount," & vbCrLf
    s = s & "       h.RemainingAmount AS OnAccount" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN', r.PurchaseReturnID, r.ReturnNumber, Null, r.ReturnDate, r.SupplierID," & vbCrLf
    s = s & "       r.RefundType, -r.TaxableAmount, -r.Tax, -r.TotalAmount, -r.RefundedAmount," & vbCrLf
    s = s & "       -(r.TotalAmount - r.RefundedAmount)" & vbCrLf
    s = s & "FROM PurchaseReturns AS r" & vbCrLf
    SaveQuery "qryPurchaseDocuments", "مستندات الشراء: الفواتير (+) والمرتجعات (-)", s
End Sub

Private Sub Q_PurchasesQuery()
    Dim s As String
    s = "SELECT d.DocType, d.DocID, d.DocNumber, d.SupplierRef, d.DocDate, d.SupplierID," & vbCrLf
    s = s & "       s.SupplierName, d.PaymentType, d.NetAmount, d.VATAmount, d.GrossAmount," & vbCrLf
    s = s & "       d.SettledAmount, d.OnAccount" & vbCrLf
    s = s & "FROM qryPurchaseDocuments AS d INNER JOIN Suppliers AS s ON d.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE d.DocDate >= QDate('PeriodStart') AND d.DocDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY d.DocDate" & vbCrLf
    SaveQuery "PurchasesQuery", "فواتير ومرتجعات الشراء خلال فترة مع المورد", s
End Sub

Private Sub Q_qryProductLedger()
    Dim s As String
    s = "SELECT ProductID, Sum(Quantity) AS LedgerQty, Max(TransactionDate) AS LastMovementDate" & vbCrLf
    s = s & "FROM InventoryTransactions" & vbCrLf
    s = s & "GROUP BY ProductID" & vbCrLf
    SaveQuery "qryProductLedger", "رصيد كل منتج من دفتر حركة المخزون", s
End Sub

Private Sub Q_qryProductLastSale()
    Dim s As String
    s = "SELECT d.ProductID, Max(h.InvoiceDate) AS LastSaleDate" & vbCrLf
    s = s & "FROM SalesInvoiceDetails AS d INNER JOIN SalesInvoices AS h" & vbCrLf
    s = s & "     ON d.SalesInvoiceID = h.SalesInvoiceID" & vbCrLf
    s = s & "GROUP BY d.ProductID" & vbCrLf
    SaveQuery "qryProductLastSale", "تاريخ آخر بيع لكل منتج", s
End Sub

Private Sub Q_StockBalanceQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName, u.UnitName," & vbCrLf
    s = s & "       p.ProductLocation, p.CurrentQuantity, p.MinimumQuantity, p.AverageCost, p.SellingPrice," & vbCrLf
    s = s & "       p.CurrentQuantity * p.AverageCost AS StockCostValue," & vbCrLf
    s = s & "       p.CurrentQuantity * p.SellingPrice AS StockSalesValue," & vbCrLf
    s = s & "       CCur(Nz(l.LedgerQty, 0)) AS LedgerQuantity," & vbCrLf
    s = s & "       p.CurrentQuantity - CCur(Nz(l.LedgerQty, 0)) AS QuantityMismatch," & vbCrLf
    s = s & "       IIf(p.CurrentQuantity <= p.MinimumQuantity, True, False) AS IsLowStock, p.IsActive" & vbCrLf
    s = s & "FROM ((Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID" & vbCrLf
    s = s & "ORDER BY p.ProductName" & vbCrLf
    SaveQuery "StockBalanceQuery", "المخزون الحالي: الكمية والقيمة بالتكلفة وبسعر البيع ومطابقتها مع الحركات", s
End Sub

Private Sub Q_LowStockQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.Barcode, p.ProductName, c.CategoryName," & vbCrLf
    s = s & "       p.CurrentQuantity, p.MinimumQuantity, p.MinimumQuantity - p.CurrentQuantity AS ShortageQty," & vbCrLf
    s = s & "       s.SupplierName, s.Mobile AS SupplierMobile" & vbCrLf
    s = s & "FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)" & vbCrLf
    s = s & "     LEFT JOIN Suppliers AS s ON p.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE p.IsActive = True AND p.CurrentQuantity <= p.MinimumQuantity" & vbCrLf
    s = s & "ORDER BY p.MinimumQuantity - p.CurrentQuantity DESC, p.ProductName" & vbCrLf
    SaveQuery "LowStockQuery", "المنتجات منخفضة المخزون: CurrentQuantity <= MinimumQuantity", s
End Sub

Private Sub Q_ProductMovementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, t.TransactionID, t.TransactionDate AS MovementDate," & vbCrLf
    s = s & "       tt.TypeName AS MovementType, t.ReferenceNumber," & vbCrLf
    s = s & "       IIf(t.Quantity > 0, t.Quantity, 0) AS QtyIn, IIf(t.Quantity < 0, -t.Quantity, 0) AS QtyOut," & vbCrLf
    s = s & "       t.Quantity AS NetQty, t.UnitCost, t.Notes" & vbCrLf
    s = s & "FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt" & vbCrLf
    s = s & "     ON t.TransactionTypeID = tt.TransactionTypeID" & vbCrLf
    s = s & "WHERE t.ProductID = QLong('ProductID') AND t.TransactionDate >= QDate('PeriodStart') AND t.TransactionDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, 0, QDate('PeriodStart'), 'رصيد أول المدة', Null," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Quantity), 0)) > 0, CCur(Nz(Sum(o.Quantity), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Quantity), 0)) < 0, -CCur(Nz(Sum(o.Quantity), 0)), 0)," & vbCrLf
    s = s & "       CCur(Nz(Sum(o.Quantity), 0)), Null, Null" & vbCrLf
    s = s & "FROM InventoryTransactions AS o" & vbCrLf
    s = s & "WHERE o.ProductID = QLong('ProductID') AND o.TransactionDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, MovementDate, TransactionID" & vbCrLf
    SaveQuery "ProductMovementQuery", "حركة منتج خلال فترة مع رصيد أول المدة (الرصيد التراكمي في التقرير)", s
End Sub

Private Sub Q_SlowMovingProductsQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity," & vbCrLf
    s = s & "       p.AverageCost, p.CurrentQuantity * p.AverageCost AS StockCostValue, ls.LastSaleDate," & vbCrLf
    s = s & "       DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) AS DaysWithoutSale" & vbCrLf
    s = s & "FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)" & vbCrLf
    s = s & "     LEFT JOIN qryProductLastSale AS ls ON p.ProductID = ls.ProductID" & vbCrLf
    s = s & "WHERE p.IsActive = True AND p.CurrentQuantity > 0" & vbCrLf
    s = s & "  AND DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) >=" & vbCrLf
    s = s & "      (SELECT SlowMovingDays FROM Settings)" & vbCrLf
    s = s & "ORDER BY DateDiff('d', Nz(ls.LastSaleDate, p.CreatedAt), Date()) DESC" & vbCrLf
    SaveQuery "SlowMovingProductsQuery", "المنتجات غير المتحركة: لها رصيد ولم تُبع منذ عدد الأيام المحدد في الإعدادات", s
End Sub

Private Sub Q_StockByCategoryQuery()
    Dim s As String
    s = "SELECT c.CategoryID, c.CategoryName, Count(*) AS ProductCount," & vbCrLf
    s = s & "       Sum(p.CurrentQuantity) AS TotalQuantity," & vbCrLf
    s = s & "       Sum(p.CurrentQuantity * p.AverageCost) AS StockCostValue," & vbCrLf
    s = s & "       Sum(p.CurrentQuantity * p.SellingPrice) AS StockSalesValue" & vbCrLf
    s = s & "FROM Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID" & vbCrLf
    s = s & "WHERE p.IsActive = True" & vbCrLf
    s = s & "GROUP BY c.CategoryID, c.CategoryName" & vbCrLf
    s = s & "ORDER BY c.CategoryName" & vbCrLf
    SaveQuery "StockByCategoryQuery", "المخزون حسب التصنيف: عدد المنتجات والكمية والقيمة", s
End Sub

Private Sub Q_StockCountQuery()
    Dim s As String
    s = "SELECT c.StockCountID, c.CountNumber, c.CountDate, c.Status, c.CategoryID, g.CategoryName," & vbCrLf
    s = s & "       d.ProductID, p.ProductCode, p.ProductName, d.SystemQuantity, d.ActualQuantity, d.Difference," & vbCrLf
    s = s & "       d.UnitCost, d.DifferenceValue, d.Notes" & vbCrLf
    s = s & "FROM ((StockCountDetails AS d INNER JOIN StockCounts AS c ON d.StockCountID = c.StockCountID)" & vbCrLf
    s = s & "      INNER JOIN Products AS p ON d.ProductID = p.ProductID)" & vbCrLf
    s = s & "     LEFT JOIN Categories AS g ON c.CategoryID = g.CategoryID" & vbCrLf
    s = s & "ORDER BY c.StockCountID, p.ProductName" & vbCrLf
    SaveQuery "StockCountQuery", "تفاصيل جلسات الجرد: الكمية المسجلة والفعلية والفرق وقيمته", s
End Sub

Private Sub Q_qryCustomerLedger()
    Dim s As String
    s = "SELECT h.CustomerID, h.InvoiceDate AS EntryDate, 'SALE' AS EntryType," & vbCrLf
    s = s & "       'فاتورة بيع' AS EntryTypeName, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.TotalAmount AS Debit, h.PaidAmount AS Credit" & vbCrLf
    s = s & "FROM SalesInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.CustomerID, r.ReturnDate, 'SALES_RETURN', 'مرتجع بيع', r.SalesReturnID, r.ReturnNumber," & vbCrLf
    s = s & "       r.RefundedAmount, r.TotalAmount" & vbCrLf
    s = s & "FROM SalesReturns AS r" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.CustomerID, p.PaymentDate, 'PAYMENT', 'سند قبض', p.PaymentID, p.PaymentNumber," & vbCrLf
    s = s & "       CCur(0), p.Amount" & vbCrLf
    s = s & "FROM CustomerPayments AS p" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT c.CustomerID, c.CreatedAt, 'OPENING', 'رصيد افتتاحي', 0, '-'," & vbCrLf
    s = s & "       IIf(c.OpeningBalance > 0, c.OpeningBalance, 0)," & vbCrLf
    s = s & "       IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0)" & vbCrLf
    s = s & "FROM Customers AS c" & vbCrLf
    s = s & "WHERE c.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qryCustomerLedger", "دفتر حساب العملاء: مدين (عليه) / دائن (له)", s
End Sub

Private Sub Q_qryCustomerLedgerTotals()
    Dim s As String
    s = "SELECT CustomerID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit," & vbCrLf
    s = s & "       Max(EntryDate) AS LastEntryDate" & vbCrLf
    s = s & "FROM qryCustomerLedger" & vbCrLf
    s = s & "GROUP BY CustomerID" & vbCrLf
    SaveQuery "qryCustomerLedgerTotals", "مجاميع حساب كل عميل", s
End Sub

Private Sub Q_CustomerBalanceQuery()
    Dim s As String
    s = "SELECT c.CustomerID, c.CustomerName, c.Mobile, c.CreditLimit, c.AllowCredit, c.IsActive," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) - CCur(Nz(l.TotalCredit, 0)) AS Balance," & vbCrLf
    s = s & "       c.CurrentBalance AS CachedBalance, l.LastEntryDate" & vbCrLf
    s = s & "FROM Customers AS c LEFT JOIN qryCustomerLedgerTotals AS l ON c.CustomerID = l.CustomerID" & vbCrLf
    s = s & "ORDER BY c.CustomerName" & vbCrLf
    SaveQuery "CustomerBalanceQuery", "رصيد كل عميل محسوبًا من الحركات (موجب = عليه للمحل)", s
End Sub

Private Sub Q_CustomersWithDebtQuery()
    Dim s As String
    s = "SELECT * FROM CustomerBalanceQuery" & vbCrLf
    s = s & "WHERE Balance > 0" & vbCrLf
    s = s & "ORDER BY Balance DESC" & vbCrLf
    SaveQuery "CustomersWithDebtQuery", "العملاء الذين عليهم مبالغ مستحقة", s
End Sub

Private Sub Q_CustomerStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM qryCustomerLedger AS l" & vbCrLf
    s = s & "WHERE l.CustomerID = QLong('CustomerID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-'," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)" & vbCrLf
    s = s & "FROM qryCustomerLedger AS o" & vbCrLf
    s = s & "WHERE o.CustomerID = QLong('CustomerID') AND o.EntryDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, EntryDate" & vbCrLf
    SaveQuery "CustomerStatementQuery", "كشف حساب عميل لفترة: رصيد سابق ثم الحركات", s
End Sub

Private Sub Q_qrySupplierLedger()
    Dim s As String
    s = "SELECT h.SupplierID, h.InvoiceDate AS EntryDate, 'PURCHASE' AS EntryType," & vbCrLf
    s = s & "       'فاتورة شراء' AS EntryTypeName, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.PaidAmount AS Debit, h.TotalAmount AS Credit" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.SupplierID, r.ReturnDate, 'PURCHASE_RETURN', 'مرتجع شراء', r.PurchaseReturnID," & vbCrLf
    s = s & "       r.ReturnNumber, r.TotalAmount, r.RefundedAmount" & vbCrLf
    s = s & "FROM PurchaseReturns AS r" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.SupplierID, p.PaymentDate, 'PAYMENT', 'سند صرف', p.PaymentID, p.PaymentNumber," & vbCrLf
    s = s & "       p.Amount, CCur(0)" & vbCrLf
    s = s & "FROM SupplierPayments AS p" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT s.SupplierID, s.CreatedAt, 'OPENING', 'رصيد افتتاحي', 0, '-'," & vbCrLf
    s = s & "       IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0)," & vbCrLf
    s = s & "       IIf(s.OpeningBalance > 0, s.OpeningBalance, 0)" & vbCrLf
    s = s & "FROM Suppliers AS s" & vbCrLf
    s = s & "WHERE s.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qrySupplierLedger", "دفتر حساب الموردين: دائن (للمورد) / مدين (سُدِّد له)", s
End Sub

Private Sub Q_qrySupplierLedgerTotals()
    Dim s As String
    s = "SELECT SupplierID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit," & vbCrLf
    s = s & "       Max(EntryDate) AS LastEntryDate" & vbCrLf
    s = s & "FROM qrySupplierLedger" & vbCrLf
    s = s & "GROUP BY SupplierID" & vbCrLf
    SaveQuery "qrySupplierLedgerTotals", "مجاميع حساب كل مورد", s
End Sub

Private Sub Q_SupplierBalanceQuery()
    Dim s As String
    s = "SELECT s.SupplierID, s.SupplierName, s.ContactPerson, s.Mobile, s.IsActive," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal," & vbCrLf
    s = s & "       CCur(Nz(l.TotalCredit, 0)) - CCur(Nz(l.TotalDebit, 0)) AS Balance," & vbCrLf
    s = s & "       s.CurrentBalance AS CachedBalance, l.LastEntryDate" & vbCrLf
    s = s & "FROM Suppliers AS s LEFT JOIN qrySupplierLedgerTotals AS l ON s.SupplierID = l.SupplierID" & vbCrLf
    s = s & "ORDER BY s.SupplierName" & vbCrLf
    SaveQuery "SupplierBalanceQuery", "رصيد كل مورد محسوبًا من الحركات (موجب = مستحق للمورد)", s
End Sub

Private Sub Q_SupplierStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM qrySupplierLedger AS l" & vbCrLf
    s = s & "WHERE l.SupplierID = QLong('SupplierID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', 'رصيد سابق', '-'," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)" & vbCrLf
    s = s & "FROM qrySupplierLedger AS o" & vbCrLf
    s = s & "WHERE o.SupplierID = QLong('SupplierID') AND o.EntryDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, EntryDate" & vbCrLf
    SaveQuery "SupplierStatementQuery", "كشف حساب مورد لفترة: رصيد سابق ثم الحركات", s
End Sub

Private Sub Q_ExpensesQuery()
    Dim s As String
    s = "SELECT e.ExpenseID, e.ExpenseNumber, e.ExpenseDate, t.ExpenseTypeName, e.Amount, e.Tax," & vbCrLf
    s = s & "       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName" & vbCrLf
    s = s & "FROM ((Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID)" & vbCrLf
    s = s & "      INNER JOIN Employees AS em ON e.EmployeeID = em.EmployeeID)" & vbCrLf
    s = s & "     LEFT JOIN PaymentMethods AS pm ON e.PaymentMethodID = pm.PaymentMethodID" & vbCrLf
    s = s & "WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY e.ExpenseDate" & vbCrLf
    SaveQuery "ExpensesQuery", "المصروفات خلال فترة", s
End Sub

Private Sub Q_ExpensesByTypeQuery()
    Dim s As String
    s = "SELECT t.ExpenseTypeName, Count(*) AS ExpenseCount, Sum(e.Amount) AS AmountExVAT," & vbCrLf
    s = s & "       Sum(e.Tax) AS InputVAT, Sum(e.TotalAmount) AS AmountTotal" & vbCrLf
    s = s & "FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID" & vbCrLf
    s = s & "WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "GROUP BY t.ExpenseTypeName" & vbCrLf
    s = s & "ORDER BY Sum(e.TotalAmount) DESC" & vbCrLf
    SaveQuery "ExpensesByTypeQuery", "المصروفات مجمّعة حسب النوع خلال فترة", s
End Sub

Private Sub Q_qryProfitSales()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(LineNet), 0)) AS PeriodNetSales, CCur(Nz(Sum(LineCost), 0)) AS PeriodCost" & vbCrLf
    s = s & "FROM qrySalesLinesInPeriod" & vbCrLf
    SaveQuery "qryProfitSales", "صافي المبيعات وتكلفتها خلال الفترة", s
End Sub

Private Sub Q_qryProfitAdjustments()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(t.Quantity * t.UnitCost), 0)) AS PeriodAdjustments" & vbCrLf
    s = s & "FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt" & vbCrLf
    s = s & "     ON t.TransactionTypeID = tt.TransactionTypeID" & vbCrLf
    s = s & "WHERE tt.TypeCode IN ('STOCK_IN', 'STOCK_OUT', 'ADJUSTMENT') AND t.TransactionDate >= QDate('PeriodStart') AND t.TransactionDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryProfitAdjustments", "قيمة فروقات المخزون (جرد، إضافة، خصم) خلال الفترة", s
End Sub

Private Sub Q_qryProfitExpenses()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(Amount), 0)) AS PeriodExpenses" & vbCrLf
    s = s & "FROM Expenses" & vbCrLf
    s = s & "WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryProfitExpenses", "المصروفات (بدون ضريبة) خلال الفترة", s
End Sub

Private Sub Q_ProfitQuery()
    Dim s As String
    s = "SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo," & vbCrLf
    s = s & "       s.PeriodNetSales AS NetSales, s.PeriodCost AS CostOfSales," & vbCrLf
    s = s & "       s.PeriodNetSales - s.PeriodCost AS GrossProfit," & vbCrLf
    s = s & "       a.PeriodAdjustments AS InventoryAdjustments, x.PeriodExpenses AS TotalExpenses," & vbCrLf
    s = s & "       s.PeriodNetSales - s.PeriodCost + a.PeriodAdjustments - x.PeriodExpenses AS NetProfit" & vbCrLf
    s = s & "FROM qryProfitSales AS s, qryProfitAdjustments AS a, qryProfitExpenses AS x" & vbCrLf
    SaveQuery "ProfitQuery", "الأرباح: صافي المبيعات - التكلفة = مجمل الربح؛ ثم ± فروقات المخزون - المصروفات = صافي الربح", s
End Sub

Private Sub Q_qryVatOutput()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(LineNet), 0)) AS TaxableSales, CCur(Nz(Sum(LineVAT), 0)) AS OutputVAT" & vbCrLf
    s = s & "FROM qrySalesLinesInPeriod" & vbCrLf
    SaveQuery "qryVatOutput", "ضريبة المخرجات (المبيعات ناقص المرتجعات)", s
End Sub

Private Sub Q_qryVatInputPurchases()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(NetAmount), 0)) AS TaxablePurchases, CCur(Nz(Sum(VATAmount), 0)) AS PurchaseVAT" & vbCrLf
    s = s & "FROM qryPurchaseDocuments" & vbCrLf
    s = s & "WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryVatInputPurchases", "ضريبة المدخلات من المشتريات (ناقص المرتجعات)", s
End Sub

Private Sub Q_qryVatInputExpenses()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(Tax), 0)) AS ExpenseVAT" & vbCrLf
    s = s & "FROM Expenses" & vbCrLf
    s = s & "WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryVatInputExpenses", "ضريبة المدخلات من المصروفات", s
End Sub

Private Sub Q_VatSummaryQuery()
    Dim s As String
    s = "SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo," & vbCrLf
    s = s & "       o.TaxableSales, o.OutputVAT, p.TaxablePurchases, p.PurchaseVAT, e.ExpenseVAT," & vbCrLf
    s = s & "       p.PurchaseVAT + e.ExpenseVAT AS InputVAT," & vbCrLf
    s = s & "       o.OutputVAT - p.PurchaseVAT - e.ExpenseVAT AS NetVATDue" & vbCrLf
    s = s & "FROM qryVatOutput AS o, qryVatInputPurchases AS p, qryVatInputExpenses AS e" & vbCrLf
    SaveQuery "VatSummaryQuery", "ملخص ضريبة القيمة المضافة للفترة (للإقرار الضريبي)", s
End Sub

Private Sub Q_DashboardQuery()
    Dim s As String
    s = "SELECT (SELECT CCur(Nz(Sum(d.GrossAmount), 0)) FROM qrySalesDocuments AS d" & vbCrLf
    s = s & "        WHERE d.DocDate >= QDate('DashDay') AND d.DocDate < QDate('DashEnd')) AS TodaySales," & vbCrLf
    s = s & "       (SELECT Count(*) FROM SalesInvoices AS h" & vbCrLf
    s = s & "        WHERE h.InvoiceDate >= QDate('DashDay') AND h.InvoiceDate < QDate('DashEnd')) AS TodayInvoices," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(d.GrossAmount), 0)) FROM qrySalesDocuments AS d" & vbCrLf
    s = s & "        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthSales," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(d.VATAmount), 0)) FROM qrySalesDocuments AS d" & vbCrLf
    s = s & "        WHERE d.DocDate >= QDate('DashMonth') AND d.DocDate < QDate('DashEnd')) AS MonthVAT," & vbCrLf
    s = s & "       (SELECT Count(*) FROM SalesInvoices AS h" & vbCrLf
    s = s & "        WHERE h.InvoiceDate >= QDate('DashMonth') AND h.InvoiceDate < QDate('DashEnd')) AS MonthInvoices," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(e.Amount), 0)) FROM Expenses AS e" & vbCrLf
    s = s & "        WHERE e.ExpenseDate >= QDate('DashMonth') AND e.ExpenseDate < QDate('DashEnd')) AS MonthExpenses," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(c.CurrentBalance), 0)) FROM Customers AS c WHERE c.CurrentBalance > 0) AS CustomerDebt," & vbCrLf
    s = s & "       (SELECT Count(*) FROM Customers AS c WHERE c.CurrentBalance > 0) AS DebtorCount," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(s.CurrentBalance), 0)) FROM Suppliers AS s WHERE s.CurrentBalance > 0) AS SupplierDue," & vbCrLf
    s = s & "       (SELECT CCur(Nz(Sum(p.CurrentQuantity * p.AverageCost), 0)) FROM Products AS p" & vbCrLf
    s = s & "        WHERE p.IsActive = True AND p.CurrentQuantity > 0) AS StockValue," & vbCrLf
    s = s & "       (SELECT Count(*) FROM LowStockQuery) AS LowStockCount" & vbCrLf
    s = s & "FROM Settings AS st" & vbCrLf
    s = s & "WHERE st.SettingID = 1" & vbCrLf
    SaveQuery "DashboardQuery", "مؤشرات لوحة التحكم في سجل واحد (اليوم، الشهر، الأرصدة، المخزون)", s
End Sub

Private Sub Q_qryDashboardTopProducts()
    Dim s As String
    s = "SELECT l.ProductID, p.ProductName, Sum(l.SignedQty) AS NetQty, Sum(l.LineGross) AS NetSales" & vbCrLf
    s = s & "FROM qrySalesLineItems AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID" & vbCrLf
    s = s & "WHERE l.DocDate >= QDate('DashMonth') AND l.DocDate < QDate('DashEnd')" & vbCrLf
    s = s & "GROUP BY l.ProductID, p.ProductName" & vbCrLf
    s = s & "HAVING Sum(l.SignedQty) > 0" & vbCrLf
    SaveQuery "qryDashboardTopProducts", "صافي الكمية المباعة لكل منتج منذ بداية الشهر (لوحة التحكم)", s
End Sub

Private Sub Q_qrySalesDocPrint()
    Dim s As String
    s = "SELECT 'SALE' AS DocKind, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.InvoiceDate AS DocDate, '' AS OriginalNumber, h.InvoiceSubType, h.PaymentType," & vbCrLf
    s = s & "       h.CustomerID, c.CustomerName, c.VATNumber AS CustomerVAT, c.City AS CustomerCity," & vbCrLf
    s = s & "       c.District AS CustomerDistrict, c.StreetName AS CustomerStreet," & vbCrLf
    s = s & "       c.BuildingNo AS CustomerBuilding, c.PostalCode AS CustomerPostal, e.EmployeeName," & vbCrLf
    s = s & "       h.SubTotal AS DocSubTotal, h.Discount AS DocDiscount, h.TaxableAmount, h.Tax AS DocTax," & vbCrLf
    s = s & "       h.TotalAmount, h.PaidAmount, h.RemainingAmount, h.AmountTendered, h.ChangeDue," & vbCrLf
    s = s & "       d.LineNumber, p.ProductName, p.ProductCode, u.UnitName, d.Quantity, d.UnitPrice," & vbCrLf
    s = s & "       d.Discount AS LineDiscount, d.NetAmount, d.VATRate, d.Tax AS LineTax, d.LineTotal" & vbCrLf
    s = s & "FROM ((((SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d ON h.SalesInvoiceID = d.SalesInvoiceID)" & vbCrLf
    s = s & "       INNER JOIN Products AS p ON d.ProductID = p.ProductID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID)" & vbCrLf
    s = s & "    INNER JOIN Employees AS e ON h.EmployeeID = e.EmployeeID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN', r.SalesReturnID, r.ReturnNumber, r.ReturnDate, o.InvoiceNumber, r.InvoiceSubType," & vbCrLf
    s = s & "       r.RefundType, r.CustomerID, c.CustomerName, c.VATNumber, c.City, c.District, c.StreetName," & vbCrLf
    s = s & "       c.BuildingNo, c.PostalCode, e.EmployeeName, r.SubTotal, r.Discount, r.TaxableAmount, r.Tax," & vbCrLf
    s = s & "       r.TotalAmount, r.RefundedAmount, r.TotalAmount - r.RefundedAmount, CCur(0), CCur(0)," & vbCrLf
    s = s & "       rd.ReturnDetailID, p.ProductName, p.ProductCode, u.UnitName, rd.Quantity, rd.UnitPrice," & vbCrLf
    s = s & "       rd.Discount, rd.NetAmount, rd.VATRate, rd.Tax, rd.LineTotal" & vbCrLf
    s = s & "FROM (((((SalesReturns AS r INNER JOIN SalesReturnDetails AS rd ON r.SalesReturnID = rd.SalesReturnID)" & vbCrLf
    s = s & "        INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = o.SalesInvoiceID)" & vbCrLf
    s = s & "       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID)" & vbCrLf
    s = s & "    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID" & vbCrLf
    SaveQuery "qrySalesDocPrint", "بيانات طباعة فواتير البيع والإشعارات الدائنة (سطر لكل صنف)", s
End Sub

Private Sub Q_qryPurchaseDocPrint()
    Dim s As String
    s = "SELECT 'PURCHASE' AS DocKind, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.InvoiceDate AS DocDate, h.SupplierInvoiceNo, '' AS OriginalNumber, h.PaymentType," & vbCrLf
    s = s & "       '' AS Reason, h.SupplierID, s.SupplierName, s.VATNumber AS SupplierVAT," & vbCrLf
    s = s & "       s.Mobile AS SupplierMobile, e.EmployeeName, h.SubTotal AS DocSubTotal," & vbCrLf
    s = s & "       h.Discount AS DocDiscount, h.TaxableAmount, h.Tax AS DocTax, h.TotalAmount," & vbCrLf
    s = s & "       h.PaidAmount, h.RemainingAmount, d.LineNumber, p.ProductCode, p.ProductName, u.UnitName," & vbCrLf
    s = s & "       d.Quantity, d.UnitCost, d.Discount AS LineDiscount, d.NetAmount, d.VATRate," & vbCrLf
    s = s & "       d.Tax AS LineTax, d.LineTotal" & vbCrLf
    s = s & "FROM ((((PurchaseInvoices AS h INNER JOIN PurchaseInvoiceDetails AS d" & vbCrLf
    s = s & "         ON h.PurchaseInvoiceID = d.PurchaseInvoiceID)" & vbCrLf
    s = s & "       INNER JOIN Products AS p ON d.ProductID = p.ProductID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID)" & vbCrLf
    s = s & "    INNER JOIN Employees AS e ON h.EmployeeID = e.EmployeeID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN', r.PurchaseReturnID, r.ReturnNumber, r.ReturnDate, o.SupplierInvoiceNo," & vbCrLf
    s = s & "       o.InvoiceNumber, r.RefundType, r.Reason, r.SupplierID, s.SupplierName, s.VATNumber," & vbCrLf
    s = s & "       s.Mobile, e.EmployeeName, r.SubTotal, r.Discount, r.TaxableAmount, r.Tax, r.TotalAmount," & vbCrLf
    s = s & "       r.RefundedAmount, r.TotalAmount - r.RefundedAmount, od.LineNumber, p.ProductCode," & vbCrLf
    s = s & "       p.ProductName, u.UnitName, rd.Quantity, rd.UnitCost, rd.Discount, rd.NetAmount," & vbCrLf
    s = s & "       rd.VATRate, rd.Tax, rd.LineTotal" & vbCrLf
    s = s & "FROM ((((((PurchaseReturns AS r INNER JOIN PurchaseReturnDetails AS rd" & vbCrLf
    s = s & "           ON r.PurchaseReturnID = rd.PurchaseReturnID)" & vbCrLf
    s = s & "         INNER JOIN PurchaseInvoiceDetails AS od ON rd.PurchaseDetailID = od.PurchaseDetailID)" & vbCrLf
    s = s & "        INNER JOIN PurchaseInvoices AS o ON r.PurchaseInvoiceID = o.PurchaseInvoiceID)" & vbCrLf
    s = s & "       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID)" & vbCrLf
    s = s & "    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID" & vbCrLf
    SaveQuery "qryPurchaseDocPrint", "بيانات طباعة فواتير الشراء ومرتجعاتها (سطر لكل صنف)", s
End Sub

Private Sub Q_qryVoucherPrint()
    Dim s As String
    s = "SELECT 'RECEIPT' AS DocKind, p.PaymentID AS DocID, p.PaymentNumber AS DocNumber," & vbCrLf
    s = s & "       p.PaymentDate AS DocDate, 1 AS LineNumber, c.CustomerName AS PartyName," & vbCrLf
    s = s & "       c.Mobile AS PartyMobile, p.Amount, m.MethodName, p.Notes, e.EmployeeName," & vbCrLf
    s = s & "       c.CurrentBalance AS PartyBalance" & vbCrLf
    s = s & "FROM ((CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID)" & vbCrLf
    s = s & "      INNER JOIN PaymentMethods AS m ON p.PaymentMethodID = m.PaymentMethodID)" & vbCrLf
    s = s & "     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PAYMENT', p.PaymentID, p.PaymentNumber, p.PaymentDate, 1, s.SupplierName, s.Mobile," & vbCrLf
    s = s & "       p.Amount, m.MethodName, p.Notes, e.EmployeeName, s.CurrentBalance" & vbCrLf
    s = s & "FROM ((SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID)" & vbCrLf
    s = s & "      INNER JOIN PaymentMethods AS m ON p.PaymentMethodID = m.PaymentMethodID)" & vbCrLf
    s = s & "     INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID" & vbCrLf
    SaveQuery "qryVoucherPrint", "بيانات طباعة سندات القبض (من العملاء) وسندات الصرف (للموردين)", s
End Sub

Private Sub Q_qrySalesInvoiceLineTotals()
    Dim s As String
    s = "SELECT SalesInvoiceID, Sum(LineTotal) AS LinesTotal" & vbCrLf
    s = s & "FROM SalesInvoiceDetails" & vbCrLf
    s = s & "GROUP BY SalesInvoiceID" & vbCrLf
    SaveQuery "qrySalesInvoiceLineTotals", "مجموع أسطر كل فاتورة بيع", s
End Sub

Private Sub Q_qryPurchaseInvoiceLineTotals()
    Dim s As String
    s = "SELECT PurchaseInvoiceID, Sum(LineTotal) AS LinesTotal" & vbCrLf
    s = s & "FROM PurchaseInvoiceDetails" & vbCrLf
    s = s & "GROUP BY PurchaseInvoiceID" & vbCrLf
    SaveQuery "qryPurchaseInvoiceLineTotals", "مجموع أسطر كل فاتورة شراء", s
End Sub

Private Sub Q_qrySalesReturnedQty()
    Dim s As String
    s = "SELECT SalesDetailID, Sum(Quantity) AS QtyReturned" & vbCrLf
    s = s & "FROM SalesReturnDetails" & vbCrLf
    s = s & "GROUP BY SalesDetailID" & vbCrLf
    SaveQuery "qrySalesReturnedQty", "الكمية المرتجعة من كل سطر فاتورة بيع", s
End Sub

Private Sub Q_qryPurchaseReturnedQty()
    Dim s As String
    s = "SELECT PurchaseDetailID, Sum(Quantity) AS QtyReturned" & vbCrLf
    s = s & "FROM PurchaseReturnDetails" & vbCrLf
    s = s & "GROUP BY PurchaseDetailID" & vbCrLf
    SaveQuery "qryPurchaseReturnedQty", "الكمية المرتجعة للمورد من كل سطر فاتورة شراء", s
End Sub

Private Sub Q_IntegrityCheckQuery()
    Dim s As String
    s = "SELECT 'STOCK_MISMATCH' AS IssueCode, 'الكمية المسجلة لا تطابق حركات المخزون' AS IssueText," & vbCrLf
    s = s & "       'Products' AS SourceTable, p.ProductID AS RecordID," & vbCrLf
    s = s & "       CCur(Nz(l.LedgerQty, 0)) AS ExpectedValue, CCur(p.CurrentQuantity) AS ActualValue" & vbCrLf
    s = s & "FROM Products AS p LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID" & vbCrLf
    s = s & "WHERE p.CurrentQuantity <> CCur(Nz(l.LedgerQty, 0))" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'NEGATIVE_STOCK', 'رصيد المنتج سالب', 'Products', p.ProductID, CCur(0)," & vbCrLf
    s = s & "       CCur(p.CurrentQuantity)" & vbCrLf
    s = s & "FROM Products AS p" & vbCrLf
    s = s & "WHERE p.CurrentQuantity < 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CUSTOMER_BALANCE', 'رصيد العميل المسجل لا يطابق الحركات', 'Customers', b.CustomerID," & vbCrLf
    s = s & "       CCur(b.Balance), CCur(b.CachedBalance)" & vbCrLf
    s = s & "FROM CustomerBalanceQuery AS b" & vbCrLf
    s = s & "WHERE b.Balance <> b.CachedBalance" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_BALANCE', 'رصيد المورد المسجل لا يطابق الحركات', 'Suppliers', b.SupplierID," & vbCrLf
    s = s & "       CCur(b.Balance), CCur(b.CachedBalance)" & vbCrLf
    s = s & "FROM SupplierBalanceQuery AS b" & vbCrLf
    s = s & "WHERE b.Balance <> b.CachedBalance" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE_TOTAL', 'إجمالي فاتورة البيع لا يساوي مجموع أسطرها', 'SalesInvoices'," & vbCrLf
    s = s & "       h.SalesInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)" & vbCrLf
    s = s & "FROM SalesInvoices AS h LEFT JOIN qrySalesInvoiceLineTotals AS t" & vbCrLf
    s = s & "     ON h.SalesInvoiceID = t.SalesInvoiceID" & vbCrLf
    s = s & "WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.SalesInvoiceID Is Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_TOTAL', 'إجمالي فاتورة الشراء لا يساوي مجموع أسطرها', 'PurchaseInvoices'," & vbCrLf
    s = s & "       h.PurchaseInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h LEFT JOIN qryPurchaseInvoiceLineTotals AS t" & vbCrLf
    s = s & "     ON h.PurchaseInvoiceID = t.PurchaseInvoiceID" & vbCrLf
    s = s & "WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.PurchaseInvoiceID Is Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CREDIT_NOT_ALLOWED', 'بيع آجل لعميل غير مسموح له بالآجل', 'SalesInvoices'," & vbCrLf
    s = s & "       h.SalesInvoiceID, CCur(0), CCur(h.RemainingAmount)" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.RemainingAmount > 0 AND c.AllowCredit = False" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_CUSTOMER', 'عميل المرتجع يختلف عن عميل الفاتورة الأصلية', 'SalesReturns'," & vbCrLf
    s = s & "       r.SalesReturnID, CCur(h.CustomerID), CCur(r.CustomerID)" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN SalesInvoices AS h ON r.SalesInvoiceID = h.SalesInvoiceID" & vbCrLf
    s = s & "WHERE r.CustomerID <> h.CustomerID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_LINE_INVOICE', 'سطر المرتجع من فاتورة غير فاتورة المرتجع', 'SalesReturnDetails'," & vbCrLf
    s = s & "       d.ReturnDetailID, CCur(r.SalesInvoiceID), CCur(o.SalesInvoiceID)" & vbCrLf
    s = s & "FROM (SalesReturnDetails AS d INNER JOIN SalesReturns AS r ON d.SalesReturnID = r.SalesReturnID)" & vbCrLf
    s = s & "     INNER JOIN SalesInvoiceDetails AS o ON d.SalesDetailID = o.SalesDetailID" & vbCrLf
    s = s & "WHERE o.SalesInvoiceID <> r.SalesInvoiceID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_LINE_PRODUCT', 'منتج سطر المرتجع يختلف عن منتج السطر الأصلي', 'SalesReturnDetails'," & vbCrLf
    s = s & "       d.ReturnDetailID, CCur(o.ProductID), CCur(d.ProductID)" & vbCrLf
    s = s & "FROM SalesReturnDetails AS d INNER JOIN SalesInvoiceDetails AS o" & vbCrLf
    s = s & "     ON d.SalesDetailID = o.SalesDetailID" & vbCrLf
    s = s & "WHERE d.ProductID <> o.ProductID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE_RETURN_QTY', 'الكمية المرتجعة أكبر من الكمية المباعة', 'SalesInvoiceDetails'," & vbCrLf
    s = s & "       o.SalesDetailID, CCur(o.Quantity), CCur(q.QtyReturned)" & vbCrLf
    s = s & "FROM SalesInvoiceDetails AS o INNER JOIN qrySalesReturnedQty AS q" & vbCrLf
    s = s & "     ON o.SalesDetailID = q.SalesDetailID" & vbCrLf
    s = s & "WHERE q.QtyReturned > o.Quantity" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_RETURN_QTY', 'الكمية المرتجعة للمورد أكبر من الكمية المشتراة', 'PurchaseInvoiceDetails'," & vbCrLf
    s = s & "       o.PurchaseDetailID, CCur(o.Quantity), CCur(q.QtyReturned)" & vbCrLf
    s = s & "FROM PurchaseInvoiceDetails AS o INNER JOIN qryPurchaseReturnedQty AS q" & vbCrLf
    s = s & "     ON o.PurchaseDetailID = q.PurchaseDetailID" & vbCrLf
    s = s & "WHERE q.QtyReturned > o.Quantity" & vbCrLf
    s = s & "ORDER BY IssueCode, RecordID" & vbCrLf
    SaveQuery "IntegrityCheckQuery", "فحص سلامة البيانات: أي سطر هنا مشكلة يجب مراجعتها (النتيجة الفارغة = سليم)", s
End Sub
