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
Private Const QUERY_NAMES As String = "qrySalesDocuments,qrySalesLineItems,qrySalesLinesInPeriod,DailySalesQuery,qrySalesMonthlyDocs,qrySalesMonthlyCost,MonthlySalesQuery,SalesByPeriodQuery,SalesByProductQuery,BestSellingProductsQuery,LeastSellingProductsQuery,qryPurchaseDocuments,PurchasesQuery,qryProductLedger,qryProductLastSale,StockBalanceQuery,LowStockQuery,ProductMovementQuery,SlowMovingProductsQuery,StockByCategoryQuery,qryCustomerLedger,qryCustomerLedgerTotals,CustomerBalanceQuery,CustomersWithDebtQuery,CustomerStatementQuery,qrySupplierLedger,qrySupplierLedgerTotals,SupplierBalanceQuery,SupplierStatementQuery,ExpensesQuery,ExpensesByTypeQuery,qryProfitSales,qryProfitAdjustments,qryProfitExpenses,ProfitQuery,qryVatOutput,qryVatInputPurchases,qryVatInputExpenses,VatSummaryQuery,qrySalesInvoiceLineTotals,qryPurchaseInvoiceLineTotals,qrySalesReturnedQty,qryPurchaseReturnedQty,IntegrityCheckQuery"

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
    Debug.Print "--- ÃœÌœ…: " & m_created & " | „Õœ¯À…: " & m_updated & " | ›‘· : " & m_failed

    If m_failed = 0 Then
        MsgBox " „ ≈‰‘«¡ «·«” ⁄·«„«  »‰Ã«Õ." & vbCrLf & vbCrLf & _
               "«” ⁄·«„«  ÃœÌœ…: " & m_created & vbCrLf & _
               "«” ⁄·«„«  „Õœ¯À…: " & m_updated & vbCrLf & vbCrLf & _
               "«·ŒÿÊ… «· «·Ì…: ‘€¯· TestQueries", vbInformation + MSG_RTL, "BuildQueries"
        BuildQueries = True
    Else
        MsgBox "›‘· ≈‰‘«¡ " & m_failed & " «” ⁄·«„:" & vbCrLf & vbCrLf & Left$(m_report, 900), _
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
        MsgBox "Ì⁄„· Â–« «·«Œ »«— ⁄·Ï ﬁ«⁄œ… »œÊ‰ Õ—ﬂ«  ›ﬁÿ° ·√‰ ‰ «∆ÃÂ √—ﬁ«„ „Õœœ… „”»ﬁ«." & _
               vbCrLf & "ÌÊÃœ »Ì«‰«  ›Ì «·ÃœÊ·: " & blocker, vbExclamation + MSG_RTL, "TestQueries"
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

    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed & " ( „ «· —«Ã⁄ ⁄‰ »Ì«‰«  «·«Œ »«—)"
    If m_failed = 0 Then
        MsgBox "Ã„Ì⁄ «Œ »«—«  «·«” ⁄·«„«  ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               "«·√—ﬁ«„ „ÿ«»ﬁ… ··Õ”«»«  «·ÌœÊÌ…° Ê·„  ı —ﬂ √Ì »Ì«‰«  «Œ »«—.", _
               vbInformation + MSG_RTL, "TestQueries"
        TestQueries = True
    Else
        MsgBox "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestQueries"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "Œÿ√ €Ì— „ Êﬁ⁄ " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    If inTrans Then ws.Rollback
    CleanUpAfterTest slowDays
    ClearQueryParams
    MsgBox errText & vbCrLf & " „ «· —«Ã⁄ ⁄‰ »Ì«‰«  «·«Œ »«—.", vbCritical + MSG_RTL, "TestQueries"
End Function

Public Sub DropQueries()
    ' DEVELOPMENT ONLY - deletes the queries created by BuildQueries.
    Dim db As DAO.Database, qName As Variant, removed As Long
    If InputBox("”Ì „ Õ–› «” ⁄·«„«  «·‰Ÿ«„ («·»Ì«‰«  ·«  ıÕ–›)." & vbCrLf & _
                "·· √ﬂÌœ «ﬂ » DELETE", "DropQueries") <> "DELETE" Then Exit Sub
    Set db = CurrentDb
    For Each qName In Split(QUERY_NAMES, ",")
        If QueryExists(db, CStr(qName)) Then
            db.QueryDefs.Delete CStr(qName)
            removed = removed + 1
        End If
    Next
    Application.RefreshDatabaseWindow
    MsgBox " „ Õ–› " & removed & " «” ⁄·«„.", vbInformation + MSG_RTL
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
    Fail QueryName & ": Œÿ√ " & Err.Number & " - " & Err.Description
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
            Fail "«·«” ⁄·«„ " & qName & " ·« Ì⁄„·: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            rs.Close
            opened = opened + 1
        End If
        On Error GoTo 0
    Next
    Call Record(opened = UBound(Split(QUERY_NAMES, ",")) + 1, _
                "ﬂ· «·«” ⁄·«„«  (" & opened & ")  ⁄„· »œÊ‰ √Œÿ«¡")
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
        Fail Label & ": ·«  ÊÃœ ‰ ÌÃ… («·„ Êﬁ⁄ " & Expected & ")"
    ElseIf Abs(CDbl(actual) - Expected) < 0.005 Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label & " = " & actual
    Else
        Fail Label & ": «·‰ ÌÃ… " & actual & " Ê«·„ Êﬁ⁄ " & Expected
    End If
    Exit Sub
EH:
    Fail Label & ": Œÿ√ " & Err.Number & " - " & Err.Description
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
        Debug.Print " ‰»ÌÂ: ·„ Ì‘„· «· —«Ã⁄ ﬂ· «·Ãœ«Ê·° Ì „ Õ–› »Ì«‰«  «·«Œ »«— ÌœÊÌ«."
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
        "INSERT INTO [Suppliers] ([SupplierName], [OpeningBalance], [CurrentBalance], [CreatedAt]) VALUES ('TEST „Ê—œ', 0, 2855, " & D(120, 0) & ")"
    Ins "C2", "Customers", "CustomerID", _
        "INSERT INTO [Customers] ([CustomerName], [OpeningBalance], [CurrentBalance], [AllowCredit], [CreatedAt]) VALUES ('TEST ⁄„Ì· ¬Ã·', 50, 164, True, " & D(90, 0) & ")"
    Ins "CAT2", "Categories", "CategoryID", _
        "INSERT INTO [Categories] ([CategoryName]) VALUES ('TEST „Ê«œ €–«∆Ì…')"
    Ins "P1", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [SupplierID], [CreatedAt]) VALUES ('TEST-P1', 'TEST „‰ Ã 1', 1, 1, 60, 60, 115, 83, 5, " & R("S1") & ", " & D(120, 0) & ")"
    Ins "P2", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [SupplierID], [CreatedAt]) VALUES ('TEST-P2', 'TEST „‰ Ã 2', " & R("CAT2") & ", 1, 10, 10, 23, 186, 200, " & R("S1") & ", " & D(120, 0) & ")"
    Ins "P3", "Products", "ProductID", _
        "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID], [UnitID], [PurchasePrice], [AverageCost], [SellingPrice], [CurrentQuantity], [MinimumQuantity], [CreatedAt]) VALUES ('TEST-P3', 'TEST „‰ Ã 3', 1, 1, 10, 10, 20, 20, 0, " & D(120, 0) & ")"
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
        "INSERT INTO [SalesReturns] ([ReturnNumber], [ReturnDate], [SalesInvoiceID], [CustomerID], [EmployeeID], [Reason], [RefundType], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [RefundedAmount]) VALUES ('TEST-CRN-1', " & D(2, 13) & ", " & R("INV2") & ", " & R("C2") & ", 1, 'TEST ≈—Ã«⁄ «·⁄„Ì·', 'CREDIT', 40, 0, 40, 6, 46, 0)"
    Ins "CRN1L1", "SalesReturnDetails", "ReturnDetailID", _
        "INSERT INTO [SalesReturnDetails] ([SalesReturnID], [SalesDetailID], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost], [ReturnToStock]) VALUES (" & R("CRN1") & ", " & R("INV2L2") & ", " & R("P2") & ", 2, 20, 0, 40, 0.15, 6, 46, 10, True)"
    Ins "T8", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(2, 13) & ", " & R("P2") & ", 4, 2, 10, 187, 'SALES_RETURN', " & R("CRN1") & ", 'TEST-CRN-1', 1)"
    Ins "PRT1", "PurchaseReturns", "PurchaseReturnID", _
        "INSERT INTO [PurchaseReturns] ([ReturnNumber], [ReturnDate], [PurchaseInvoiceID], [SupplierID], [EmployeeID], [Reason], [RefundType], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [RefundedAmount]) VALUES ('TEST-PRT-1', " & D(1, 10) & ", " & R("PUR1") & ", " & R("S1") & ", 1, 'TEST ⁄Ì» „’‰⁄Ì', 'CREDIT', 300, 0, 300, 45, 345, 0)"
    Ins "PRT1L1", "PurchaseReturnDetails", "PurchaseReturnDetailID", _
        "INSERT INTO [PurchaseReturnDetails] ([PurchaseReturnID], [PurchaseDetailID], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PRT1") & ", " & R("PUR1L1") & ", " & R("P1") & ", 5, 60, 0, 300, 0.15, 45, 345)"
    Ins "T9", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(1, 10) & ", " & R("P1") & ", 3, -5, 60, 83, 'PURCHASE_RETURN', " & R("PRT1") & ", 'TEST-PRT-1', 1)"
    Ins "T10", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(1, 15) & ", " & R("P2") & ", 6, -1, 10, 186, 'MANUAL', Null, 'TEST-ADJ-1', 1)"
    Ins "EXP1", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-1', " & D(6, 0) & ", 2, 200, 30, 230, 1, 'TEST ﬂÂ—»«¡', 1)"
    Ins "EXP2", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-2', " & D(40, 0) & ", 1, 1000, 0, 1000, 3, 'TEST ≈ÌÃ«—', 1)"
End Sub

Private Sub RunChecks()
    Chk "—’Ìœ «·„‰ Ã 1 = 100 ‘—«¡ - 10 - 2 »Ì⁄ - 5 „— Ã⁄ ‘—«¡ = 83", _
        "SELECT CurrentQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 83
    Chk "—’Ìœ «·„‰ Ã 1 „‰ œ› — «·Õ—ﬂ«  = 83", _
        "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 83
    Chk "—’Ìœ «·„‰ Ã 2 = 200 - 5 - 10 + 2 - 1 = 186", _
        "SELECT LedgerQuantity FROM StockBalanceQuery WHERE ProductID = " & R("P2"), 186
    Chk "·«  ÊÃœ ›—Êﬁ«  »Ì‰ «·ﬂ„Ì… «·„”Ã·… Ê«·Õ—ﬂ« ", _
        "SELECT COUNT(*) FROM StockBalanceQuery WHERE QuantityMismatch <> 0", 0
    Chk "ﬁÌ„… „Œ“Ê‰ «·„‰ Ã 1 »«· ﬂ·›… = 83 ◊ 60", _
        "SELECT StockCostValue FROM StockBalanceQuery WHERE ProductID = " & R("P1"), 4980
    Chk "„‰Œ›÷ «·„Œ“Ê‰: „‰ Ã Ê«Õœ ›ﬁÿ", _
        "SELECT COUNT(*) FROM LowStockQuery", 1
    Chk "„‰Œ›÷ «·„Œ“Ê‰: «·„‰ Ã 2 (186 <= 200)", _
        "SELECT ShortageQty FROM LowStockQuery WHERE ProductID = " & R("P2"), 14
    Chk "€Ì— «·„ Õ—ﬂ…: „‰ Ã Ê«Õœ", _
        "SELECT COUNT(*) FROM SlowMovingProductsQuery", 1
    Chk "€Ì— «·„ Õ—ﬂ…: «·„‰ Ã 3 »ﬁÌ„… 200", _
        "SELECT StockCostValue FROM SlowMovingProductsQuery WHERE ProductID = " & R("P3"), 200
    Chk "«·„Œ“Ê‰ Õ”» «· ’‰Ì›:  ’‰Ì› ⁄«„ = 83 + 20", _
        "SELECT TotalQuantity FROM StockByCategoryQuery WHERE CategoryID = 1", 103
    Chk "«·„Œ“Ê‰ Õ”» «· ’‰Ì›: ﬁÌ„…  ’‰Ì› ⁄«„ = 4980 + 200", _
        "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = 1", 5180
    Chk "«·„Œ“Ê‰ Õ”» «· ’‰Ì›: ﬁÌ„… «· ’‰Ì› 2 = 186 ◊ 10", _
        "SELECT StockCostValue FROM StockByCategoryQuery WHERE CategoryID = " & R("CAT2"), 1860
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "Õ—ﬂ… «·„‰ Ã 2: —’Ìœ √Ê· «·„œ… = 195", _
        "SELECT NetQty FROM ProductMovementQuery WHERE SortKey = 0", 195
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "Õ—ﬂ… «·„‰ Ã 2: 3 Õ—ﬂ«  + ”ÿ— «·—’Ìœ «·”«»ﬁ", _
        "SELECT COUNT(*) FROM ProductMovementQuery", 4
    SetQueryParam "ProductID", CLng(R("P2"))
    Chk "Õ—ﬂ… «·„‰ Ã 2: «·—’Ìœ «·Œ «„Ì = 186", _
        "SELECT Sum(NetQty) FROM ProductMovementQuery", 186
    Chk "«·„»Ì⁄«  «·ÌÊ„Ì…: ÌÊ„ «·›« Ê—… «·¬Ã·… = 460", _
        "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(5, 0) & ")", 460
    Chk "«·„»Ì⁄«  «·ÌÊ„Ì…: «·¬Ã· ›Ì ‰›” «·ÌÊ„ = 460", _
        "SELECT CreditSales FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(5, 0) & ")", 460
    Chk "«·„»Ì⁄«  «·ÌÊ„Ì…: ÌÊ„ «·›« Ê—… «·‰ﬁœÌ… = 1150 ‰ﬁœ«", _
        "SELECT CashSales FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(10, 0) & ")", 1150
    Chk "«·„»Ì⁄«  «·ÌÊ„Ì…: ÌÊ„ «·„— Ã⁄ = -46", _
        "SELECT NetSalesTotal FROM DailySalesQuery WHERE SaleDate = DateValue(" & D(2, 0) & ")", -46
    Chk "«·„»Ì⁄«  «·‘Â—Ì…: ’«›Ì ﬂ· «·√‘Â— = 100 + 1000 + 400 - 40", _
        "SELECT Sum(NetSalesExVAT) FROM MonthlySalesQuery", 1460
    Chk "«·„»Ì⁄«  «·‘Â—Ì…: „Ã„· —»Õ ﬂ· «·√‘Â— = 1460 - 850", _
        "SELECT Sum(GrossProfit) FROM MonthlySalesQuery", 610
    Chk "«·„»Ì⁄«  Œ·«· «·› —…: ›« Ê— «‰ Ê„— Ã⁄", _
        "SELECT COUNT(*) FROM SalesByPeriodQuery", 3
    Chk "«·„»Ì⁄«  Œ·«· «·› —…: «·≈Ã„«·Ì = 1150 + 460 - 46", _
        "SELECT Sum(GrossAmount) FROM SalesByPeriodQuery", 1564
    Chk "«·„»Ì⁄«  Õ”» «·„‰ Ã: «·„‰ Ã 1 ’«›Ì ﬂ„Ì… 12", _
        "SELECT NetQty FROM SalesByProductQuery WHERE ProductID = " & R("P1"), 12
    Chk "«·„»Ì⁄«  Õ”» «·„‰ Ã: —»Õ «·„‰ Ã 1 = 1200 - 720", _
        "SELECT GrossProfit FROM SalesByProductQuery WHERE ProductID = " & R("P1"), 480
    Chk "«·„»Ì⁄«  Õ”» «·„‰ Ã: «·„‰ Ã 2 „— Ã⁄ 2", _
        "SELECT QtyReturned FROM SalesByProductQuery WHERE ProductID = " & R("P2"), 2
    Chk "«·„»Ì⁄«  Õ”» «·„‰ Ã: ’«›Ì „»Ì⁄«  «·„‰ Ã 2 = 200 - 40", _
        "SELECT NetSales FROM SalesByProductQuery WHERE ProductID = " & R("P2"), 160
    Chk "«·√ﬂÀ— „»Ì⁄«: «·„‰ Ã 1", _
        "SELECT TOP 1 ProductID FROM BestSellingProductsQuery ORDER BY NetQty DESC, NetSales DESC", CDbl(R("P1"))
    Chk "«·√ﬁ· „»Ì⁄«: «·„‰ Ã 3 (·„ Ìı»⁄)", _
        "SELECT TOP 1 ProductID FROM LeastSellingProductsQuery ORDER BY NetQtySold, ProductName", CDbl(R("P3"))
    Chk "«·„‘ —Ì«  Œ·«· «·› —…: „— Ã⁄ «·‘—«¡ ›ﬁÿ", _
        "SELECT COUNT(*) FROM PurchasesQuery", 1
    Chk "«·„‘ —Ì«  Œ·«· «·› —…: -345", _
        "SELECT Sum(GrossAmount) FROM PurchasesQuery", -345
    Chk "—’Ìœ «·⁄„Ì· «·¬Ã· = 50 + 460 - 100 - 46 - 200", _
        "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = " & R("C2"), 164
    Chk "—’Ìœ «·⁄„Ì· «·‰ﬁœÌ = 0", _
        "SELECT Balance FROM CustomerBalanceQuery WHERE CustomerID = 1", 0
    Chk "«·⁄„·«¡ «·„œÌ‰Ê‰: ⁄„Ì· Ê«Õœ", _
        "SELECT COUNT(*) FROM CustomersWithDebtQuery", 1
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "ﬂ‘› «·⁄„Ì·: «·—’Ìœ «·”«»ﬁ = 50", _
        "SELECT Debit FROM CustomerStatementQuery WHERE SortKey = 0", 50
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "ﬂ‘› «·⁄„Ì·: —’Ìœ ”«»ﬁ + 3 Õ—ﬂ« ", _
        "SELECT COUNT(*) FROM CustomerStatementQuery", 4
    SetQueryParam "CustomerID", CLng(R("C2"))
    Chk "ﬂ‘› «·⁄„Ì·: «·—’Ìœ «·Œ «„Ì = 164", _
        "SELECT Sum(Debit) - Sum(Credit) FROM CustomerStatementQuery", 164
    Chk "—’Ìœ «·„Ê—œ = 9200 - 5000 - 1000 - 345", _
        "SELECT Balance FROM SupplierBalanceQuery WHERE SupplierID = " & R("S1"), 2855
    SetQueryParam "SupplierID", CLng(R("S1"))
    Chk "ﬂ‘› «·„Ê—œ: «·—’Ìœ «·”«»ﬁ = 4200", _
        "SELECT Credit FROM SupplierStatementQuery WHERE SortKey = 0", 4200
    SetQueryParam "SupplierID", CLng(R("S1"))
    Chk "ﬂ‘› «·„Ê—œ: «·—’Ìœ «·Œ «„Ì = 2855", _
        "SELECT Sum(Credit) - Sum(Debit) FROM SupplierStatementQuery", 2855
    Chk "«·„’—Ê›«  Œ·«· «·› —…: „’—Ê› Ê«Õœ", _
        "SELECT COUNT(*) FROM ExpensesQuery", 1
    Chk "«·„’—Ê›«  Œ·«· «·› —…: 230 ‘«„· «·÷—Ì»…", _
        "SELECT Sum(TotalAmount) FROM ExpensesQuery", 230
    Chk "«·„’—Ê›«  Õ”» «·‰Ê⁄: «·ﬂÂ—»«¡ 200", _
        "SELECT AmountExVAT FROM ExpensesByTypeQuery", 200
    Chk "«·√—»«Õ: ’«›Ì «·„»Ì⁄«  = 1000 + 400 - 40", _
        "SELECT NetSales FROM ProfitQuery", 1360
    Chk "«·√—»«Õ:  ﬂ·›… «·„»Ì⁄«  = 600 + 220 - 20", _
        "SELECT CostOfSales FROM ProfitQuery", 800
    Chk "«·√—»«Õ: „Ã„· «·—»Õ = 1360 - 800", _
        "SELECT GrossProfit FROM ProfitQuery", 560
    Chk "«·√—»«Õ: ›—Êﬁ«  «·„Œ“Ê‰ = -10 („‰ Ã  «·›)", _
        "SELECT InventoryAdjustments FROM ProfitQuery", -10
    Chk "«·√—»«Õ: «·„’—Ê›«  = 200", _
        "SELECT TotalExpenses FROM ProfitQuery", 200
    Chk "«·√—»«Õ: ’«›Ì «·—»Õ = 560 - 10 - 200", _
        "SELECT NetProfit FROM ProfitQuery", 350
    Chk "«·÷—Ì»…: ÷—Ì»… «·„Œ—Ã«  = 150 + 60 - 6", _
        "SELECT OutputVAT FROM VatSummaryQuery", 204
    Chk "«·÷—Ì»…: ÷—Ì»… «·„œŒ·«  = -45 („— Ã⁄ ‘—«¡) + 30 („’—Ê›)", _
        "SELECT InputVAT FROM VatSummaryQuery", -15
    Chk "«·÷—Ì»…: «·’«›Ì «·„” Õﬁ = 204 + 15", _
        "SELECT NetVATDue FROM VatSummaryQuery", 219
    Chk "›Õ’ «·”·«„…: ·«  ÊÃœ „‘ﬂ·« ", _
        "SELECT COUNT(*) FROM IntegrityCheckQuery", 0
End Sub

Private Sub RunCorruptionChecks()
    m_db.Execute "UPDATE [Products] SET [CurrentQuantity] = 80 WHERE [ProductID] = " & R("P1"), dbFailOnError
    Chk "›Õ’ «·”·«„… Ìﬂ ‘›  ·«⁄»« »«·ﬂ„Ì…", _
        "SELECT COUNT(*) FROM IntegrityCheckQuery WHERE IssueCode = 'STOCK_MISMATCH'", 1
    m_db.Execute "UPDATE [Customers] SET [CurrentBalance] = 0 WHERE [CustomerID] = " & R("C2"), dbFailOnError
    Chk "›Õ’ «·”·«„… Ìﬂ ‘› —’Ìœ ⁄„Ì· Œ«ÿ∆", _
        "SELECT ExpectedValue FROM IntegrityCheckQuery WHERE IssueCode = 'CUSTOMER_BALANCE'", 164
    m_db.Execute "UPDATE [SalesReturns] SET [CustomerID] = 1 WHERE [SalesReturnID] = " & R("CRN1"), dbFailOnError
    Chk "›Õ’ «·”·«„… Ìﬂ ‘› „— Ã⁄« ·⁄„Ì· „Œ ·›", _
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
    SaveQuery "qrySalesDocuments", "„” ‰œ«  «·»Ì⁄: «·›Ê« Ì— (+) Ê«·„— Ã⁄«  (-) »ﬁÌ„ „Êﬁ¯⁄…", s
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
    SaveQuery "qrySalesLineItems", "√”ÿ— «·»Ì⁄ Ê«·„— Ã⁄«  „⁄ «· ﬂ·›… (√”«”  Õ·Ì· «·„‰ Ã«  Ê«·√—»«Õ)", s
End Sub

Private Sub Q_qrySalesLinesInPeriod()
    Dim s As String
    s = "SELECT * FROM qrySalesLineItems" & vbCrLf
    s = s & "WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qrySalesLinesInPeriod", "√”ÿ— «·»Ì⁄ Ê«·„— Ã⁄«  œ«Œ· «·› —…", s
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
    SaveQuery "DailySalesQuery", "«·„»Ì⁄«  «·ÌÊ„Ì…: ⁄œœ «·›Ê« Ì— Ê«·„— Ã⁄«  Ê«·’«›Ì Ê«·‰ﬁœÌ Ê«·¬Ã·", s
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
    SaveQuery "qrySalesMonthlyDocs", " Ã„Ì⁄ ‘Â—Ì ·„” ‰œ«  «·»Ì⁄", s
End Sub

Private Sub Q_qrySalesMonthlyCost()
    Dim s As String
    s = "SELECT Year(DocDate) AS SalesYear, Month(DocDate) AS SalesMonth, Sum(LineCost) AS MonthCost" & vbCrLf
    s = s & "FROM qrySalesLineItems" & vbCrLf
    s = s & "GROUP BY Year(DocDate), Month(DocDate)" & vbCrLf
    SaveQuery "qrySalesMonthlyCost", " ﬂ·›… «·»÷«⁄… «·„»«⁄… ‘Â—Ì«", s
End Sub

Private Sub Q_MonthlySalesQuery()
    Dim s As String
    s = "SELECT d.SalesYear, d.SalesMonth, d.InvoiceCount, d.ReturnCount, d.NetSalesExVAT," & vbCrLf
    s = s & "       d.NetVAT, d.NetSalesTotal, CCur(Nz(c.MonthCost, 0)) AS CostOfSales," & vbCrLf
    s = s & "       d.NetSalesExVAT - CCur(Nz(c.MonthCost, 0)) AS GrossProfit" & vbCrLf
    s = s & "FROM qrySalesMonthlyDocs AS d LEFT JOIN qrySalesMonthlyCost AS c" & vbCrLf
    s = s & "     ON (d.SalesYear = c.SalesYear AND d.SalesMonth = c.SalesMonth)" & vbCrLf
    s = s & "ORDER BY d.SalesYear DESC, d.SalesMonth DESC" & vbCrLf
    SaveQuery "MonthlySalesQuery", "«·„»Ì⁄«  «·‘Â—Ì… „⁄ «· ﬂ·›… Ê„Ã„· «·—»Õ", s
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
    SaveQuery "SalesByPeriodQuery", "›Ê« Ì— Ê„— Ã⁄«  «·»Ì⁄ Œ·«· › —… „⁄ «·⁄„Ì· Ê«·ﬂ«‘Ì—", s
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
    SaveQuery "SalesByProductQuery", "«·„»Ì⁄«  Õ”» «·„‰ Ã Œ·«· › —… (ﬂ„Ì…° ’«›Ì°  ﬂ·›…° —»Õ)", s
End Sub

Private Sub Q_BestSellingProductsQuery()
    Dim s As String
    s = "SELECT * FROM SalesByProductQuery" & vbCrLf
    s = s & "ORDER BY NetQty DESC, NetSales DESC" & vbCrLf
    SaveQuery "BestSellingProductsQuery", "√›÷· «·„‰ Ã«  „»Ì⁄« Œ·«· › —… (Õ”» ’«›Ì «·ﬂ„Ì…)", s
End Sub

Private Sub Q_LeastSellingProductsQuery()
    Dim s As String
    s = "SELECT p.ProductID, p.ProductCode, p.ProductName, c.CategoryName, p.CurrentQuantity," & vbCrLf
    s = s & "       CCur(Nz(s.NetQty, 0)) AS NetQtySold, CCur(Nz(s.NetSales, 0)) AS NetSalesAmount" & vbCrLf
    s = s & "FROM (Products AS p INNER JOIN Categories AS c ON p.CategoryID = c.CategoryID)" & vbCrLf
    s = s & "     LEFT JOIN SalesByProductQuery AS s ON p.ProductID = s.ProductID" & vbCrLf
    s = s & "WHERE p.IsActive = True" & vbCrLf
    s = s & "ORDER BY CCur(Nz(s.NetQty, 0)), p.ProductName" & vbCrLf
    SaveQuery "LeastSellingProductsQuery", "√ﬁ· «·„‰ Ã«  „»Ì⁄« Œ·«· › —… ( ‘„· «·„‰ Ã«  «· Ì ·„  ı»⁄)", s
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
    SaveQuery "qryPurchaseDocuments", "„” ‰œ«  «·‘—«¡: «·›Ê« Ì— (+) Ê«·„— Ã⁄«  (-)", s
End Sub

Private Sub Q_PurchasesQuery()
    Dim s As String
    s = "SELECT d.DocType, d.DocID, d.DocNumber, d.SupplierRef, d.DocDate, d.SupplierID," & vbCrLf
    s = s & "       s.SupplierName, d.PaymentType, d.NetAmount, d.VATAmount, d.GrossAmount," & vbCrLf
    s = s & "       d.SettledAmount, d.OnAccount" & vbCrLf
    s = s & "FROM qryPurchaseDocuments AS d INNER JOIN Suppliers AS s ON d.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE d.DocDate >= QDate('PeriodStart') AND d.DocDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY d.DocDate" & vbCrLf
    SaveQuery "PurchasesQuery", "›Ê« Ì— Ê„— Ã⁄«  «·‘—«¡ Œ·«· › —… „⁄ «·„Ê—œ", s
End Sub

Private Sub Q_qryProductLedger()
    Dim s As String
    s = "SELECT ProductID, Sum(Quantity) AS LedgerQty, Max(TransactionDate) AS LastMovementDate" & vbCrLf
    s = s & "FROM InventoryTransactions" & vbCrLf
    s = s & "GROUP BY ProductID" & vbCrLf
    SaveQuery "qryProductLedger", "—’Ìœ ﬂ· „‰ Ã „‰ œ› — Õ—ﬂ… «·„Œ“Ê‰", s
End Sub

Private Sub Q_qryProductLastSale()
    Dim s As String
    s = "SELECT d.ProductID, Max(h.InvoiceDate) AS LastSaleDate" & vbCrLf
    s = s & "FROM SalesInvoiceDetails AS d INNER JOIN SalesInvoices AS h" & vbCrLf
    s = s & "     ON d.SalesInvoiceID = h.SalesInvoiceID" & vbCrLf
    s = s & "GROUP BY d.ProductID" & vbCrLf
    SaveQuery "qryProductLastSale", " «—ÌŒ ¬Œ— »Ì⁄ ·ﬂ· „‰ Ã", s
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
    SaveQuery "StockBalanceQuery", "«·„Œ“Ê‰ «·Õ«·Ì: «·ﬂ„Ì… Ê«·ﬁÌ„… »«· ﬂ·›… Ê»”⁄— «·»Ì⁄ Ê„ÿ«»ﬁ Â« „⁄ «·Õ—ﬂ« ", s
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
    SaveQuery "LowStockQuery", "«·„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰: CurrentQuantity <= MinimumQuantity", s
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
    s = s & "SELECT 0, 0, QDate('PeriodStart'), '—’Ìœ √Ê· «·„œ…', Null," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Quantity), 0)) > 0, CCur(Nz(Sum(o.Quantity), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Quantity), 0)) < 0, -CCur(Nz(Sum(o.Quantity), 0)), 0)," & vbCrLf
    s = s & "       CCur(Nz(Sum(o.Quantity), 0)), Null, Null" & vbCrLf
    s = s & "FROM InventoryTransactions AS o" & vbCrLf
    s = s & "WHERE o.ProductID = QLong('ProductID') AND o.TransactionDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, MovementDate, TransactionID" & vbCrLf
    SaveQuery "ProductMovementQuery", "Õ—ﬂ… „‰ Ã Œ·«· › —… „⁄ —’Ìœ √Ê· «·„œ… («·—’Ìœ «· —«ﬂ„Ì ›Ì «· ﬁ—Ì—)", s
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
    SaveQuery "SlowMovingProductsQuery", "«·„‰ Ã«  €Ì— «·„ Õ—ﬂ…: ·Â« —’Ìœ Ê·„  ı»⁄ „‰– ⁄œœ «·√Ì«„ «·„Õœœ ›Ì «·≈⁄œ«œ« ", s
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
    SaveQuery "StockByCategoryQuery", "«·„Œ“Ê‰ Õ”» «· ’‰Ì›: ⁄œœ «·„‰ Ã«  Ê«·ﬂ„Ì… Ê«·ﬁÌ„…", s
End Sub

Private Sub Q_qryCustomerLedger()
    Dim s As String
    s = "SELECT h.CustomerID, h.InvoiceDate AS EntryDate, 'SALE' AS EntryType," & vbCrLf
    s = s & "       '›« Ê—… »Ì⁄' AS EntryTypeName, h.SalesInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.TotalAmount AS Debit, h.PaidAmount AS Credit" & vbCrLf
    s = s & "FROM SalesInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.CustomerID, r.ReturnDate, 'SALES_RETURN', '„— Ã⁄ »Ì⁄', r.SalesReturnID, r.ReturnNumber," & vbCrLf
    s = s & "       r.RefundedAmount, r.TotalAmount" & vbCrLf
    s = s & "FROM SalesReturns AS r" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.CustomerID, p.PaymentDate, 'PAYMENT', '”‰œ ﬁ»÷', p.PaymentID, p.PaymentNumber," & vbCrLf
    s = s & "       CCur(0), p.Amount" & vbCrLf
    s = s & "FROM CustomerPayments AS p" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT c.CustomerID, c.CreatedAt, 'OPENING', '—’Ìœ «›  «ÕÌ', 0, '-'," & vbCrLf
    s = s & "       IIf(c.OpeningBalance > 0, c.OpeningBalance, 0)," & vbCrLf
    s = s & "       IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0)" & vbCrLf
    s = s & "FROM Customers AS c" & vbCrLf
    s = s & "WHERE c.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qryCustomerLedger", "œ› — Õ”«» «·⁄„·«¡: „œÌ‰ (⁄·ÌÂ) / œ«∆‰ (·Â)", s
End Sub

Private Sub Q_qryCustomerLedgerTotals()
    Dim s As String
    s = "SELECT CustomerID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit," & vbCrLf
    s = s & "       Max(EntryDate) AS LastEntryDate" & vbCrLf
    s = s & "FROM qryCustomerLedger" & vbCrLf
    s = s & "GROUP BY CustomerID" & vbCrLf
    SaveQuery "qryCustomerLedgerTotals", "„Ã«„Ì⁄ Õ”«» ﬂ· ⁄„Ì·", s
End Sub

Private Sub Q_CustomerBalanceQuery()
    Dim s As String
    s = "SELECT c.CustomerID, c.CustomerName, c.Mobile, c.CreditLimit, c.AllowCredit, c.IsActive," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) - CCur(Nz(l.TotalCredit, 0)) AS Balance," & vbCrLf
    s = s & "       c.CurrentBalance AS CachedBalance, l.LastEntryDate" & vbCrLf
    s = s & "FROM Customers AS c LEFT JOIN qryCustomerLedgerTotals AS l ON c.CustomerID = l.CustomerID" & vbCrLf
    s = s & "ORDER BY c.CustomerName" & vbCrLf
    SaveQuery "CustomerBalanceQuery", "—’Ìœ ﬂ· ⁄„Ì· „Õ”Ê»« „‰ «·Õ—ﬂ«  („ÊÃ» = ⁄·ÌÂ ··„Õ·)", s
End Sub

Private Sub Q_CustomersWithDebtQuery()
    Dim s As String
    s = "SELECT * FROM CustomerBalanceQuery" & vbCrLf
    s = s & "WHERE Balance > 0" & vbCrLf
    s = s & "ORDER BY Balance DESC" & vbCrLf
    SaveQuery "CustomersWithDebtQuery", "«·⁄„·«¡ «·–Ì‰ ⁄·ÌÂ„ „»«·€ „” Õﬁ…", s
End Sub

Private Sub Q_CustomerStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM qryCustomerLedger AS l" & vbCrLf
    s = s & "WHERE l.CustomerID = QLong('CustomerID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', '—’Ìœ ”«»ﬁ', '-'," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)" & vbCrLf
    s = s & "FROM qryCustomerLedger AS o" & vbCrLf
    s = s & "WHERE o.CustomerID = QLong('CustomerID') AND o.EntryDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, EntryDate" & vbCrLf
    SaveQuery "CustomerStatementQuery", "ﬂ‘› Õ”«» ⁄„Ì· ·› —…: —’Ìœ ”«»ﬁ À„ «·Õ—ﬂ« ", s
End Sub

Private Sub Q_qrySupplierLedger()
    Dim s As String
    s = "SELECT h.SupplierID, h.InvoiceDate AS EntryDate, 'PURCHASE' AS EntryType," & vbCrLf
    s = s & "       '›« Ê—… ‘—«¡' AS EntryTypeName, h.PurchaseInvoiceID AS DocID, h.InvoiceNumber AS DocNumber," & vbCrLf
    s = s & "       h.PaidAmount AS Debit, h.TotalAmount AS Credit" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.SupplierID, r.ReturnDate, 'PURCHASE_RETURN', '„— Ã⁄ ‘—«¡', r.PurchaseReturnID," & vbCrLf
    s = s & "       r.ReturnNumber, r.TotalAmount, r.RefundedAmount" & vbCrLf
    s = s & "FROM PurchaseReturns AS r" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.SupplierID, p.PaymentDate, 'PAYMENT', '”‰œ ’—›', p.PaymentID, p.PaymentNumber," & vbCrLf
    s = s & "       p.Amount, CCur(0)" & vbCrLf
    s = s & "FROM SupplierPayments AS p" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT s.SupplierID, s.CreatedAt, 'OPENING', '—’Ìœ «›  «ÕÌ', 0, '-'," & vbCrLf
    s = s & "       IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0)," & vbCrLf
    s = s & "       IIf(s.OpeningBalance > 0, s.OpeningBalance, 0)" & vbCrLf
    s = s & "FROM Suppliers AS s" & vbCrLf
    s = s & "WHERE s.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qrySupplierLedger", "œ› — Õ”«» «·„Ê—œÌ‰: œ«∆‰ (··„Ê—œ) / „œÌ‰ (”ıœˆ¯œ ·Â)", s
End Sub

Private Sub Q_qrySupplierLedgerTotals()
    Dim s As String
    s = "SELECT SupplierID, Sum(Debit) AS TotalDebit, Sum(Credit) AS TotalCredit," & vbCrLf
    s = s & "       Max(EntryDate) AS LastEntryDate" & vbCrLf
    s = s & "FROM qrySupplierLedger" & vbCrLf
    s = s & "GROUP BY SupplierID" & vbCrLf
    SaveQuery "qrySupplierLedgerTotals", "„Ã«„Ì⁄ Õ”«» ﬂ· „Ê—œ", s
End Sub

Private Sub Q_SupplierBalanceQuery()
    Dim s As String
    s = "SELECT s.SupplierID, s.SupplierName, s.ContactPerson, s.Mobile, s.IsActive," & vbCrLf
    s = s & "       CCur(Nz(l.TotalDebit, 0)) AS DebitTotal, CCur(Nz(l.TotalCredit, 0)) AS CreditTotal," & vbCrLf
    s = s & "       CCur(Nz(l.TotalCredit, 0)) - CCur(Nz(l.TotalDebit, 0)) AS Balance," & vbCrLf
    s = s & "       s.CurrentBalance AS CachedBalance, l.LastEntryDate" & vbCrLf
    s = s & "FROM Suppliers AS s LEFT JOIN qrySupplierLedgerTotals AS l ON s.SupplierID = l.SupplierID" & vbCrLf
    s = s & "ORDER BY s.SupplierName" & vbCrLf
    SaveQuery "SupplierBalanceQuery", "—’Ìœ ﬂ· „Ê—œ „Õ”Ê»« „‰ «·Õ—ﬂ«  („ÊÃ» = „” Õﬁ ··„Ê—œ)", s
End Sub

Private Sub Q_SupplierStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, l.EntryDate, l.EntryType, l.EntryTypeName, l.DocNumber, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM qrySupplierLedger AS l" & vbCrLf
    s = s & "WHERE l.SupplierID = QLong('SupplierID') AND l.EntryDate >= QDate('PeriodStart') AND l.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', '—’Ìœ ”«»ﬁ', '-'," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) > 0, CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) < 0, CCur(Nz(Sum(o.Credit), 0)) - CCur(Nz(Sum(o.Debit), 0)), 0)" & vbCrLf
    s = s & "FROM qrySupplierLedger AS o" & vbCrLf
    s = s & "WHERE o.SupplierID = QLong('SupplierID') AND o.EntryDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, EntryDate" & vbCrLf
    SaveQuery "SupplierStatementQuery", "ﬂ‘› Õ”«» „Ê—œ ·› —…: —’Ìœ ”«»ﬁ À„ «·Õ—ﬂ« ", s
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
    SaveQuery "ExpensesQuery", "«·„’—Ê›«  Œ·«· › —…", s
End Sub

Private Sub Q_ExpensesByTypeQuery()
    Dim s As String
    s = "SELECT t.ExpenseTypeName, Count(*) AS ExpenseCount, Sum(e.Amount) AS AmountExVAT," & vbCrLf
    s = s & "       Sum(e.Tax) AS InputVAT, Sum(e.TotalAmount) AS AmountTotal" & vbCrLf
    s = s & "FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID" & vbCrLf
    s = s & "WHERE e.ExpenseDate >= QDate('PeriodStart') AND e.ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "GROUP BY t.ExpenseTypeName" & vbCrLf
    s = s & "ORDER BY Sum(e.TotalAmount) DESC" & vbCrLf
    SaveQuery "ExpensesByTypeQuery", "«·„’—Ê›«  „Ã„¯⁄… Õ”» «·‰Ê⁄ Œ·«· › —…", s
End Sub

Private Sub Q_qryProfitSales()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(LineNet), 0)) AS PeriodNetSales, CCur(Nz(Sum(LineCost), 0)) AS PeriodCost" & vbCrLf
    s = s & "FROM qrySalesLinesInPeriod" & vbCrLf
    SaveQuery "qryProfitSales", "’«›Ì «·„»Ì⁄«  Ê ﬂ·› Â« Œ·«· «·› —…", s
End Sub

Private Sub Q_qryProfitAdjustments()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(t.Quantity * t.UnitCost), 0)) AS PeriodAdjustments" & vbCrLf
    s = s & "FROM InventoryTransactions AS t INNER JOIN TransactionTypes AS tt" & vbCrLf
    s = s & "     ON t.TransactionTypeID = tt.TransactionTypeID" & vbCrLf
    s = s & "WHERE tt.TypeCode IN ('STOCK_IN', 'STOCK_OUT', 'ADJUSTMENT') AND t.TransactionDate >= QDate('PeriodStart') AND t.TransactionDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryProfitAdjustments", "ﬁÌ„… ›—Êﬁ«  «·„Œ“Ê‰ (Ã—œ° ≈÷«›…° Œ’„) Œ·«· «·› —…", s
End Sub

Private Sub Q_qryProfitExpenses()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(Amount), 0)) AS PeriodExpenses" & vbCrLf
    s = s & "FROM Expenses" & vbCrLf
    s = s & "WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryProfitExpenses", "«·„’—Ê›«  (»œÊ‰ ÷—Ì»…) Œ·«· «·› —…", s
End Sub

Private Sub Q_ProfitQuery()
    Dim s As String
    s = "SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo," & vbCrLf
    s = s & "       s.PeriodNetSales AS NetSales, s.PeriodCost AS CostOfSales," & vbCrLf
    s = s & "       s.PeriodNetSales - s.PeriodCost AS GrossProfit," & vbCrLf
    s = s & "       a.PeriodAdjustments AS InventoryAdjustments, x.PeriodExpenses AS TotalExpenses," & vbCrLf
    s = s & "       s.PeriodNetSales - s.PeriodCost + a.PeriodAdjustments - x.PeriodExpenses AS NetProfit" & vbCrLf
    s = s & "FROM qryProfitSales AS s, qryProfitAdjustments AS a, qryProfitExpenses AS x" & vbCrLf
    SaveQuery "ProfitQuery", "«·√—»«Õ: ’«›Ì «·„»Ì⁄«  - «· ﬂ·›… = „Ã„· «·—»Õ∫ À„ ± ›—Êﬁ«  «·„Œ“Ê‰ - «·„’—Ê›«  = ’«›Ì «·—»Õ", s
End Sub

Private Sub Q_qryVatOutput()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(LineNet), 0)) AS TaxableSales, CCur(Nz(Sum(LineVAT), 0)) AS OutputVAT" & vbCrLf
    s = s & "FROM qrySalesLinesInPeriod" & vbCrLf
    SaveQuery "qryVatOutput", "÷—Ì»… «·„Œ—Ã«  («·„»Ì⁄«  ‰«ﬁ’ «·„— Ã⁄« )", s
End Sub

Private Sub Q_qryVatInputPurchases()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(NetAmount), 0)) AS TaxablePurchases, CCur(Nz(Sum(VATAmount), 0)) AS PurchaseVAT" & vbCrLf
    s = s & "FROM qryPurchaseDocuments" & vbCrLf
    s = s & "WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryVatInputPurchases", "÷—Ì»… «·„œŒ·«  „‰ «·„‘ —Ì«  (‰«ﬁ’ «·„— Ã⁄« )", s
End Sub

Private Sub Q_qryVatInputExpenses()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(Tax), 0)) AS ExpenseVAT" & vbCrLf
    s = s & "FROM Expenses" & vbCrLf
    s = s & "WHERE ExpenseDate >= QDate('PeriodStart') AND ExpenseDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryVatInputExpenses", "÷—Ì»… «·„œŒ·«  „‰ «·„’—Ê›« ", s
End Sub

Private Sub Q_VatSummaryQuery()
    Dim s As String
    s = "SELECT QDate('PeriodStart') AS PeriodFrom, DateAdd('d', -1, QDate('PeriodEnd')) AS PeriodTo," & vbCrLf
    s = s & "       o.TaxableSales, o.OutputVAT, p.TaxablePurchases, p.PurchaseVAT, e.ExpenseVAT," & vbCrLf
    s = s & "       p.PurchaseVAT + e.ExpenseVAT AS InputVAT," & vbCrLf
    s = s & "       o.OutputVAT - p.PurchaseVAT - e.ExpenseVAT AS NetVATDue" & vbCrLf
    s = s & "FROM qryVatOutput AS o, qryVatInputPurchases AS p, qryVatInputExpenses AS e" & vbCrLf
    SaveQuery "VatSummaryQuery", "„·Œ’ ÷—Ì»… «·ﬁÌ„… «·„÷«›… ··› —… (··≈ﬁ—«— «·÷—Ì»Ì)", s
End Sub

Private Sub Q_qrySalesInvoiceLineTotals()
    Dim s As String
    s = "SELECT SalesInvoiceID, Sum(LineTotal) AS LinesTotal" & vbCrLf
    s = s & "FROM SalesInvoiceDetails" & vbCrLf
    s = s & "GROUP BY SalesInvoiceID" & vbCrLf
    SaveQuery "qrySalesInvoiceLineTotals", "„Ã„Ê⁄ √”ÿ— ﬂ· ›« Ê—… »Ì⁄", s
End Sub

Private Sub Q_qryPurchaseInvoiceLineTotals()
    Dim s As String
    s = "SELECT PurchaseInvoiceID, Sum(LineTotal) AS LinesTotal" & vbCrLf
    s = s & "FROM PurchaseInvoiceDetails" & vbCrLf
    s = s & "GROUP BY PurchaseInvoiceID" & vbCrLf
    SaveQuery "qryPurchaseInvoiceLineTotals", "„Ã„Ê⁄ √”ÿ— ﬂ· ›« Ê—… ‘—«¡", s
End Sub

Private Sub Q_qrySalesReturnedQty()
    Dim s As String
    s = "SELECT SalesDetailID, Sum(Quantity) AS QtyReturned" & vbCrLf
    s = s & "FROM SalesReturnDetails" & vbCrLf
    s = s & "GROUP BY SalesDetailID" & vbCrLf
    SaveQuery "qrySalesReturnedQty", "«·ﬂ„Ì… «·„— Ã⁄… „‰ ﬂ· ”ÿ— ›« Ê—… »Ì⁄", s
End Sub

Private Sub Q_qryPurchaseReturnedQty()
    Dim s As String
    s = "SELECT PurchaseDetailID, Sum(Quantity) AS QtyReturned" & vbCrLf
    s = s & "FROM PurchaseReturnDetails" & vbCrLf
    s = s & "GROUP BY PurchaseDetailID" & vbCrLf
    SaveQuery "qryPurchaseReturnedQty", "«·ﬂ„Ì… «·„— Ã⁄… ··„Ê—œ „‰ ﬂ· ”ÿ— ›« Ê—… ‘—«¡", s
End Sub

Private Sub Q_IntegrityCheckQuery()
    Dim s As String
    s = "SELECT 'STOCK_MISMATCH' AS IssueCode, '«·ﬂ„Ì… «·„”Ã·… ·«  ÿ«»ﬁ Õ—ﬂ«  «·„Œ“Ê‰' AS IssueText," & vbCrLf
    s = s & "       'Products' AS SourceTable, p.ProductID AS RecordID," & vbCrLf
    s = s & "       CCur(Nz(l.LedgerQty, 0)) AS ExpectedValue, CCur(p.CurrentQuantity) AS ActualValue" & vbCrLf
    s = s & "FROM Products AS p LEFT JOIN qryProductLedger AS l ON p.ProductID = l.ProductID" & vbCrLf
    s = s & "WHERE p.CurrentQuantity <> CCur(Nz(l.LedgerQty, 0))" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'NEGATIVE_STOCK', '—’Ìœ «·„‰ Ã ”«·»', 'Products', p.ProductID, CCur(0)," & vbCrLf
    s = s & "       CCur(p.CurrentQuantity)" & vbCrLf
    s = s & "FROM Products AS p" & vbCrLf
    s = s & "WHERE p.CurrentQuantity < 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CUSTOMER_BALANCE', '—’Ìœ «·⁄„Ì· «·„”Ã· ·« Ìÿ«»ﬁ «·Õ—ﬂ« ', 'Customers', b.CustomerID," & vbCrLf
    s = s & "       CCur(b.Balance), CCur(b.CachedBalance)" & vbCrLf
    s = s & "FROM CustomerBalanceQuery AS b" & vbCrLf
    s = s & "WHERE b.Balance <> b.CachedBalance" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_BALANCE', '—’Ìœ «·„Ê—œ «·„”Ã· ·« Ìÿ«»ﬁ «·Õ—ﬂ« ', 'Suppliers', b.SupplierID," & vbCrLf
    s = s & "       CCur(b.Balance), CCur(b.CachedBalance)" & vbCrLf
    s = s & "FROM SupplierBalanceQuery AS b" & vbCrLf
    s = s & "WHERE b.Balance <> b.CachedBalance" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE_TOTAL', '≈Ã„«·Ì ›« Ê—… «·»Ì⁄ ·« Ì”«ÊÌ „Ã„Ê⁄ √”ÿ—Â«', 'SalesInvoices'," & vbCrLf
    s = s & "       h.SalesInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)" & vbCrLf
    s = s & "FROM SalesInvoices AS h LEFT JOIN qrySalesInvoiceLineTotals AS t" & vbCrLf
    s = s & "     ON h.SalesInvoiceID = t.SalesInvoiceID" & vbCrLf
    s = s & "WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.SalesInvoiceID Is Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_TOTAL', '≈Ã„«·Ì ›« Ê—… «·‘—«¡ ·« Ì”«ÊÌ „Ã„Ê⁄ √”ÿ—Â«', 'PurchaseInvoices'," & vbCrLf
    s = s & "       h.PurchaseInvoiceID, CCur(Nz(t.LinesTotal, 0)), CCur(h.TotalAmount)" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h LEFT JOIN qryPurchaseInvoiceLineTotals AS t" & vbCrLf
    s = s & "     ON h.PurchaseInvoiceID = t.PurchaseInvoiceID" & vbCrLf
    s = s & "WHERE h.TotalAmount <> CCur(Nz(t.LinesTotal, 0)) OR t.PurchaseInvoiceID Is Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CREDIT_NOT_ALLOWED', '»Ì⁄ ¬Ã· ·⁄„Ì· €Ì— „”„ÊÕ ·Â »«·¬Ã·', 'SalesInvoices'," & vbCrLf
    s = s & "       h.SalesInvoiceID, CCur(0), CCur(h.RemainingAmount)" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.RemainingAmount > 0 AND c.AllowCredit = False" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_CUSTOMER', '⁄„Ì· «·„— Ã⁄ ÌŒ ·› ⁄‰ ⁄„Ì· «·›« Ê—… «·√’·Ì…', 'SalesReturns'," & vbCrLf
    s = s & "       r.SalesReturnID, CCur(h.CustomerID), CCur(r.CustomerID)" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN SalesInvoices AS h ON r.SalesInvoiceID = h.SalesInvoiceID" & vbCrLf
    s = s & "WHERE r.CustomerID <> h.CustomerID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_LINE_INVOICE', '”ÿ— «·„— Ã⁄ „‰ ›« Ê—… €Ì— ›« Ê—… «·„— Ã⁄', 'SalesReturnDetails'," & vbCrLf
    s = s & "       d.ReturnDetailID, CCur(r.SalesInvoiceID), CCur(o.SalesInvoiceID)" & vbCrLf
    s = s & "FROM (SalesReturnDetails AS d INNER JOIN SalesReturns AS r ON d.SalesReturnID = r.SalesReturnID)" & vbCrLf
    s = s & "     INNER JOIN SalesInvoiceDetails AS o ON d.SalesDetailID = o.SalesDetailID" & vbCrLf
    s = s & "WHERE o.SalesInvoiceID <> r.SalesInvoiceID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'RETURN_LINE_PRODUCT', '„‰ Ã ”ÿ— «·„— Ã⁄ ÌŒ ·› ⁄‰ „‰ Ã «·”ÿ— «·√’·Ì', 'SalesReturnDetails'," & vbCrLf
    s = s & "       d.ReturnDetailID, CCur(o.ProductID), CCur(d.ProductID)" & vbCrLf
    s = s & "FROM SalesReturnDetails AS d INNER JOIN SalesInvoiceDetails AS o" & vbCrLf
    s = s & "     ON d.SalesDetailID = o.SalesDetailID" & vbCrLf
    s = s & "WHERE d.ProductID <> o.ProductID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE_RETURN_QTY', '«·ﬂ„Ì… «·„— Ã⁄… √ﬂ»— „‰ «·ﬂ„Ì… «·„»«⁄…', 'SalesInvoiceDetails'," & vbCrLf
    s = s & "       o.SalesDetailID, CCur(o.Quantity), CCur(q.QtyReturned)" & vbCrLf
    s = s & "FROM SalesInvoiceDetails AS o INNER JOIN qrySalesReturnedQty AS q" & vbCrLf
    s = s & "     ON o.SalesDetailID = q.SalesDetailID" & vbCrLf
    s = s & "WHERE q.QtyReturned > o.Quantity" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_RETURN_QTY', '«·ﬂ„Ì… «·„— Ã⁄… ··„Ê—œ √ﬂ»— „‰ «·ﬂ„Ì… «·„‘ —«…', 'PurchaseInvoiceDetails'," & vbCrLf
    s = s & "       o.PurchaseDetailID, CCur(o.Quantity), CCur(q.QtyReturned)" & vbCrLf
    s = s & "FROM PurchaseInvoiceDetails AS o INNER JOIN qryPurchaseReturnedQty AS q" & vbCrLf
    s = s & "     ON o.PurchaseDetailID = q.PurchaseDetailID" & vbCrLf
    s = s & "WHERE q.QtyReturned > o.Quantity" & vbCrLf
    s = s & "ORDER BY IssueCode, RecordID" & vbCrLf
    SaveQuery "IntegrityCheckQuery", "›Õ’ ”·«„… «·»Ì«‰« : √Ì ”ÿ— Â‰« „‘ﬂ·… ÌÃ» „—«Ã⁄ Â« («·‰ ÌÃ… «·›«—€… = ”·Ì„)", s
End Sub
