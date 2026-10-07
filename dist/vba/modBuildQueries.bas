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
Private Const QUERY_NAMES As String = "qrySalesDocuments,qrySalesLineItems,qrySalesLinesInPeriod,DailySalesQuery,qrySalesMonthlyDocs,qrySalesMonthlyCost,MonthlySalesQuery,SalesByPeriodQuery,SalesByProductQuery,BestSellingProductsQuery,SalesByCategoryQuery,LeastSellingProductsQuery,qryPurchaseDocuments,PurchasesQuery,qryProductLedger,qryProductLastSale,StockBalanceQuery,LowStockQuery,ProductMovementQuery,SlowMovingProductsQuery,StockByC" & _
    "ategoryQuery,StockCountQuery,qryCustomerLedger,qryCustomerLedgerTotals,CustomerBalanceQuery,CustomersWithDebtQuery,CustomerStatementQuery,qrySupplierLedger,qrySupplierLedgerTotals,SupplierBalanceQuery,SupplierStatementQuery,ExpensesQuery,ExpensesByTypeQuery,qryProfitSales,qryProfitAdjustments,qryProfitExpenses,ProfitQuery,qryVatOutput,qryVatInputPurchases,qryVatInputExpenses,VatSummaryQuery,qryVat" & _
    "ReturnLines,qryVatReturnTotals,qryVatReturnHead,VatReturnQuery,DashboardQuery,qryDashboardTopProducts,qrySalesDocPrint,qryPurchaseDocPrint,qryVoucherPrint,qryCashMovements,qryCashBoxTotals,CashBoxBalanceQuery,CashStatementQuery,qryCashDays,qryCashDayOpening,CashDailyQuery,CashClosingsQuery,qryCashClosingPrint,qryCashVoucherPrint,qrySaleCost,qryReturnCost,qryStockCountValue,qryJournalSale,qryJourna" & _
    "lSalesReturn,qryJournalPurchase,qryJournalPurchaseReturn,qryJournalPayments,qryJournalExpense,qryJournalCashVoucher,qryJournalStock,qryJournalOpening,qryManualEntryLines,qryJournalManual,qryYearCloseLines,qryJournalYearClose,qryJournalVatReturn,JournalLinesQuery,qryJournalEntryPrint,qryTrialBefore,qryTrialPeriod,TrialBalanceQuery,qryStatementBefore,AccountStatementQuery,GeneralLedgerQuery,qryTreeR" & _
    "ollup,TrialBalanceTreeQuery,qryIncomeMoves,qryCompareMoves,qryIncomeAccounts,IncomeStatementQuery,qryBalanceAt,qryBalanceCompare,qryBalanceAccounts,qryProfitAt,qryProfitCompare,qryBalanceItems,BalanceSheetQuery,AccountTreeQuery,qrySalesInvoiceLineTotals,qryPurchaseInvoiceLineTotals,qrySalesReturnedQty,qryPurchaseReturnedQty,IntegrityCheckQuery"

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
        Call ResultBox("Ì⁄„· Â–« «·«Œ »«— ⁄·Ï ﬁ«⁄œ… »œÊ‰ Õ—ﬂ«  ›ﬁÿ° ·√‰ ‰ «∆ÃÂ √—ﬁ«„ „Õœœ… „”»ﬁ«." & _
               vbCrLf & "ÌÊÃœ »Ì«‰«  ›Ì «·ÃœÊ·: " & blocker, vbExclamation + MSG_RTL, "TestQueries")
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
        Call ResultBox("Ã„Ì⁄ «Œ »«—«  «·«” ⁄·«„«  ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               "«·√—ﬁ«„ „ÿ«»ﬁ… ··Õ”«»«  «·ÌœÊÌ…° Ê·„  ı —ﬂ √Ì »Ì«‰«  «Œ »«—.", _
               vbInformation + MSG_RTL, "TestQueries")
        TestQueries = True
    Else
        Call ResultBox("‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestQueries")
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
    Call ResultBox(errText & vbCrLf & " „ «· —«Ã⁄ ⁄‰ »Ì«‰«  «·«Œ »«—.", vbCritical + MSG_RTL, "TestQueries")
End Function

Private Sub ResultBox(ByVal Text As String, ByVal Style As Long, Optional ByVal Title As String = "")
    ' Through modCommon.TestMsg when it is installed (RunAllTests collects the results),
    ' otherwise a plain message box: this module is installed before modCommon.
    On Error GoTo Plain
    Application.Run "TestMsg", Text, Style, Title
    Exit Sub
Plain:
    MsgBox Text, Style, Title
End Sub

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
    Ins "BOXM", "CashBoxes", "CashBoxID", _
        "INSERT INTO [CashBoxes] ([BoxName], [BoxType], [OpeningBalance], [OpeningDate]) VALUES ('TEST «·Œ“Ì‰…', 'MAIN', 10000, " & D(60, 0) & ")"
    Ins "BOXC", "CashBoxes", "CashBoxID", _
        "INSERT INTO [CashBoxes] ([BoxName], [BoxType], [OpeningBalance], [OpeningDate]) VALUES ('TEST ’‰œÊﬁ ﬂ«‘Ì—', 'CASHIER', 500, " & D(60, 0) & ")"
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
        "INSERT INTO [PurchaseInvoices] ([InvoiceNumber], [SupplierInvoiceNo], [InvoiceDate], [SupplierID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [CashBoxID]) VALUES ('TEST-PUR-1', 'S-100', " & D(50, 9) & ", " & R("S1") & ", 1, 'CREDIT', 1, 8000, 0, 8000, 1200, 9200, 5000, 4200, " & R("BOXM") & ")"
    Ins "PUR1L1", "PurchaseInvoiceDetails", "PurchaseDetailID", _
        "INSERT INTO [PurchaseInvoiceDetails] ([PurchaseInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PUR1") & ", 1, " & R("P1") & ", 100, 60, 0, 6000, 0.15, 900, 6900)"
    Ins "PUR1L2", "PurchaseInvoiceDetails", "PurchaseDetailID", _
        "INSERT INTO [PurchaseInvoiceDetails] ([PurchaseInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitCost], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal]) VALUES (" & R("PUR1") & ", 2, " & R("P2") & ", 200, 10, 0, 2000, 0.15, 300, 2300)"
    Ins "T2", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(50, 9) & ", " & R("P1") & ", 1, 100, 60, 100, 'PURCHASE', " & R("PUR1") & ", 'TEST-PUR-1', 1)"
    Ins "T3", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(50, 9) & ", " & R("P2") & ", 1, 200, 10, 200, 'PURCHASE', " & R("PUR1") & ", 'TEST-PUR-1', 1)"
    Ins "INV0", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue], [CashBoxID]) VALUES ('TEST-INV-0', " & D(45, 10) & ", 1, 1, 'CASH', 1, 100, 0, 100, 15, 115, 115, 0, 115, 0, " & R("BOXC") & ")"
    Ins "INV0L1", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV0") & ", 1, " & R("P2") & ", 5, 20, 0, 100, 0.15, 15, 115, 10)"
    Ins "T4", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(45, 10) & ", " & R("P2") & ", 2, -5, 10, 195, 'SALE', " & R("INV0") & ", 'TEST-INV-0', 1)"
    Ins "INV1", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue], [CashBoxID]) VALUES ('TEST-INV-1', " & D(10, 11) & ", 1, 1, 'CASH', 1, 1000, 0, 1000, 150, 1150, 1150, 0, 1200, 50, " & R("BOXC") & ")"
    Ins "INV1L1", "SalesInvoiceDetails", "SalesDetailID", _
        "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice], [Discount], [NetAmount], [VATRate], [Tax], [LineTotal], [UnitCost]) VALUES (" & R("INV1") & ", 1, " & R("P1") & ", 10, 100, 0, 1000, 0.15, 150, 1150, 60)"
    Ins "T5", "InventoryTransactions", "TransactionID", _
        "INSERT INTO [InventoryTransactions] ([TransactionDate], [ProductID], [TransactionTypeID], [Quantity], [UnitCost], [QuantityAfter], [ReferenceType], [ReferenceID], [ReferenceNumber], [EmployeeID]) VALUES (" & D(10, 11) & ", " & R("P1") & ", 2, -10, 60, 90, 'SALE', " & R("INV1") & ", 'TEST-INV-1', 1)"
    Ins "INV2", "SalesInvoices", "SalesInvoiceID", _
        "INSERT INTO [SalesInvoices] ([InvoiceNumber], [InvoiceDate], [CustomerID], [EmployeeID], [PaymentType], [PaymentMethodID], [SubTotal], [Discount], [TaxableAmount], [Tax], [TotalAmount], [PaidAmount], [RemainingAmount], [AmountTendered], [ChangeDue], [CashBoxID]) VALUES ('TEST-INV-2', " & D(5, 12) & ", " & R("C2") & ", 1, 'CREDIT', 1, 400, 0, 400, 60, 460, 100, 360, 100, 0, " & R("BOXC") & ")"
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
        "INSERT INTO [CustomerPayments] ([PaymentNumber], [CustomerID], [PaymentDate], [Amount], [PaymentMethodID], [EmployeeID], [CashBoxID]) VALUES ('TEST-RCV-1', " & R("C2") & ", " & D(3, 10) & ", 200, 1, 1, " & R("BOXC") & ")"
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
    Ins "CNT1", "StockCounts", "StockCountID", _
        "INSERT INTO [StockCounts] ([CountNumber], [CountDate], [Status], [EmployeeID]) VALUES ('TEST-CNT-1', " & D(1, 18) & ", 'OPEN', 1)"
    Ins "CNT1L1", "StockCountDetails", "StockCountDetailID", _
        "INSERT INTO [StockCountDetails] ([StockCountID], [ProductID], [SystemQuantity], [ActualQuantity], [Difference], [UnitCost], [DifferenceValue]) VALUES (" & R("CNT1") & ", " & R("P1") & ", 83, 80, -3, 60, -180)"
    Ins "CNT1L2", "StockCountDetails", "StockCountDetailID", _
        "INSERT INTO [StockCountDetails] ([StockCountID], [ProductID], [SystemQuantity], [ActualQuantity], [Difference], [UnitCost], [DifferenceValue]) VALUES (" & R("CNT1") & ", " & R("P2") & ", 186, Null, 0, 10, 0)"
    Ins "EXP1", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID], [CashBoxID]) VALUES ('TEST-EXP-1', " & D(6, 0) & ", 2, 200, 30, 230, 1, 'TEST ﬂÂ—»«¡', 1, " & R("BOXC") & ")"
    Ins "EXP2", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-2', " & D(40, 0) & ", 1, 1000, 0, 1000, 3, 'TEST ≈ÌÃ«—', 1)"
    Ins "EXPV", "Expenses", "ExpenseID", _
        "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseDate], [ExpenseTypeID], [Amount], [Tax], [TotalAmount], [PaymentMethodID], [Description], [EmployeeID]) VALUES ('TEST-EXP-V', " & D(35, 0) & ", 9, 50, 0, 50, 1, 'TEST ‰À—Ì« ', 1)"
    Ins "V1", "CashVouchers", "CashVoucherID", _
        "INSERT INTO [CashVouchers] ([VoucherNumber], [VoucherDate], [VoucherType], [CashBoxID], [Category], [Amount], [PartyName], [ExpenseID], [EmployeeID]) VALUES ('TEST-COT-1', " & D(35, 12) & ", 'OUT', " & R("BOXC") & ", 'EXPENSE', 50, 'TEST „Õ·', " & R("EXPV") & ", 1)"
    Ins "CL1", "CashClosings", "ClosingID", _
        "INSERT INTO [CashClosings] ([ClosingNumber], [ClosingDate], [CashBoxID], [EmployeeID], [OpeningBalance], [CashIn], [CashOut], [ExpectedBalance], [CountedAmount], [Difference], [Destination], [ToCashBoxID], [TransferAmount], [KeptAmount]) VALUES ('TEST-CLS-1', " & D(4, 18) & ", " & R("BOXC") & ", 1, 0, 1865, 280, 1585, 1570, -15, 'MAIN', " & R("BOXM") & ", 1000, 570)"
    Ins "V2", "CashVouchers", "CashVoucherID", _
        "INSERT INTO [CashVouchers] ([VoucherNumber], [VoucherDate], [VoucherType], [CashBoxID], [Category], [Amount], [ClosingID], [EmployeeID]) VALUES ('TEST-COT-2', " & D(4, 18) & ", 'OUT', " & R("BOXC") & ", 'SHORTAGE', 15, " & R("CL1") & ", 1)"
    Ins "V3", "CashVouchers", "CashVoucherID", _
        "INSERT INTO [CashVouchers] ([VoucherNumber], [VoucherDate], [VoucherType], [CashBoxID], [ToCashBoxID], [Category], [Amount], [ClosingID], [EmployeeID]) VALUES ('TEST-TRF-1', " & D(4, 18) & ", 'TRANSFER', " & R("BOXC") & ", " & R("BOXM") & ", 'TRANSFER', 1000, " & R("CL1") & ", 1)"
    Ins "V4", "CashVouchers", "CashVoucherID", _
        "INSERT INTO [CashVouchers] ([VoucherNumber], [VoucherDate], [VoucherType], [CashBoxID], [Category], [Amount], [PartyName], [EmployeeID]) VALUES ('TEST-CIN-1', " & D(2, 9) & ", 'IN', " & R("BOXM") & ", 'OWNER', 2000, 'TEST «·„«·ﬂ', 1)"
    Ins "V5", "CashVouchers", "CashVoucherID", _
        "INSERT INTO [CashVouchers] ([VoucherNumber], [VoucherDate], [VoucherType], [CashBoxID], [Category], [Amount], [PartyName], [EmployeeID]) VALUES ('TEST-COT-3', " & D(1, 12) & ", 'OUT', " & R("BOXM") & ", 'OWNER', 300, 'TEST «·„«·ﬂ', 1)"
    Ins "MJ1", "ManualEntries", "ManualEntryID", _
        "INSERT INTO [ManualEntries] ([EntryNumber], [EntryDate], [Description], [TotalAmount], [EmployeeID]) VALUES ('TEST-MJ-1', " & D(3, 0) & ", 'TEST —Ê« » «·‘Â— «·„” Õﬁ…', 3000, 1)"
    Ins "MJ1A", "ManualEntryLines", "ManualLineID", _
        "INSERT INTO [ManualEntryLines] ([ManualEntryID], [LineNumber], [AccountCode], [Debit], [Credit]) VALUES (" & R("MJ1") & ", 1, 5500, 3000, 0)"
    Ins "MJ1B", "ManualEntryLines", "ManualLineID", _
        "INSERT INTO [ManualEntryLines] ([ManualEntryID], [LineNumber], [AccountCode], [Debit], [Credit], [LineText]) VALUES (" & R("MJ1") & ", 2, 2310, 0, 2500, 'TEST ’«›Ì «·—Ê« »')"
    Ins "MJ1C", "ManualEntryLines", "ManualLineID", _
        "INSERT INTO [ManualEntryLines] ([ManualEntryID], [LineNumber], [AccountCode], [Debit], [Credit]) VALUES (" & R("MJ1") & ", 3, 2320, 0, 500)"
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
    Chk "ÿ»«⁄… ›« Ê—… «·‘—«¡: ”ÿ—«‰", _
        "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'PURCHASE' AND DocID = " & R("PUR1"), 2
    Chk "ÿ»«⁄… „— Ã⁄ «·‘—«¡: —ﬁ„ «·”ÿ— «·√’·Ì Ê—ﬁ„ «·›« Ê—… «·√’·Ì…", _
        "SELECT COUNT(*) FROM qryPurchaseDocPrint WHERE DocKind = 'RETURN' AND DocID = " & R("PRT1") & " AND LineNumber = 1 AND OriginalNumber = 'TEST-PUR-1' AND RemainingAmount = 345", 1
    Chk "ÿ»«⁄… ”‰œ «·’—›: «·„»·€ Êÿ—Ìﬁ… «·œ›⁄", _
        "SELECT Amount FROM qryVoucherPrint WHERE DocKind = 'PAYMENT' AND DocID = " & R("PAY1"), 1000
    Chk "ÿ»«⁄… ”‰œ «·ﬁ»÷", _
        "SELECT COUNT(*) FROM qryVoucherPrint WHERE DocKind = 'RECEIPT' AND DocID = " & R("RCV1"), 1
    Chk "«·Ã—œ «·„› ÊÕ: ⁄Ã“ «·„‰ Ã 1 = -3 ◊ 60", _
        "SELECT DifferenceValue FROM StockCountQuery WHERE ProductID = " & R("P1"), -180
    Chk "«·Ã—œ «·„› ÊÕ: ’‰› Ê«Õœ ·„ Ìı⁄œ¯ »⁄œ", _
        "SELECT COUNT(*) FROM StockCountQuery WHERE ActualQuantity Is Null", 1
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
    Chk "«·„»Ì⁄«  Õ”» «· ’‰Ì›: «·„Ã„Ê⁄ = «·„»Ì⁄«  Õ”» «·„‰ Ã", _
        "SELECT (SELECT Sum(CategoryTotal) FROM SalesByCategoryQuery) - (SELECT Sum(SalesTotal) FROM SalesByProductQuery) FROM Settings", 0
    Chk "«·„»Ì⁄«  Õ”» «· ’‰Ì›: ’«›Ì «·ﬂ„Ì… = 12 + 8", _
        "SELECT Sum(CategoryQty) FROM SalesByCategoryQuery", 20
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
    Chk "«·≈ﬁ—«— «·÷—Ì»Ì: ÷—Ì»… «·„»Ì⁄«  «·Œ«÷⁄… («·Œ«‰… 1) = ÷—Ì»… «·„Œ—Ã« ", _
        "SELECT SalesStdVAT FROM qryVatReturnTotals", 204
    Chk "«·≈ﬁ—«— «·÷—Ì»Ì: ÷—Ì»… «·„‘ —Ì«  Ê«·„’—Ê›«  («·Œ«‰… 7) = ÷—Ì»… «·„œŒ·« ", _
        "SELECT PurchStdVAT FROM qryVatReturnTotals", -15
    Chk "«·≈ﬁ—«— «·÷—Ì»Ì: ’«›Ì «·„»Ì⁄«  «·Œ«÷⁄… («·Œ«‰… 1 „⁄ «· ⁄œÌ·« ) = ’«›Ì «·„»Ì⁄« ", _
        "SELECT SalesStdAmount + SalesStdAdjust + SalesZeroAmount + SalesZeroAdjust + SalesExemptAmount + SalesExemptAdjust FROM qryVatReturnTotals", 1360
    Chk "ÿ»«⁄… «·›« Ê—… «·¬Ã·…: ”ÿ—«‰", _
        "SELECT COUNT(*) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = " & R("INV2"), 2
    Chk "ÿ»«⁄… «·›« Ê—… «·¬Ã·…: „Ã„Ê⁄ «·√”ÿ— = 460", _
        "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'SALE' AND DocID = " & R("INV2"), 460
    Chk "ÿ»«⁄… «·≈‘⁄«— «·œ«∆‰: ”ÿ— Ê«Õœ »ﬁÌ„… 46", _
        "SELECT Sum(LineTotal) FROM qrySalesDocPrint WHERE DocKind = 'RETURN' AND DocID = " & R("CRN1"), 46
    Chk "—’Ìœ ’‰œÊﬁ «·ﬂ«‘Ì— = 500 + 115 + 1150 + 100 + 200 - 230 - 50 - 15 - 1000", _
        "SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = " & R("BOXC"), 770
    Chk "—’Ìœ «·Œ“Ì‰… = 10000 - 5000 + 1000 + 2000 - 300", _
        "SELECT Balance FROM CashBoxBalanceQuery WHERE CashBoxID = " & R("BOXM"), 7700
    Chk "«·’‰«œÌﬁ «·„”Ã·… »œÊ‰ Õ—ﬂ… —’ÌœÂ« ’›—", _
        "SELECT Sum(Balance) FROM CashBoxBalanceQuery WHERE CashBoxID <= 2", 0
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "Õ—ﬂ… ’‰œÊﬁ «·ﬂ«‘Ì—: —’Ìœ √Ê· «·„œ… = 500 + 115 - 50", _
        "SELECT AmountIn FROM CashStatementQuery WHERE SortKey = 0", 565
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "Õ—ﬂ… ’‰œÊﬁ «·ﬂ«‘Ì—: 6 Õ—ﬂ«  ›Ì «·› —…", _
        "SELECT COUNT(*) FROM CashStatementQuery WHERE SortKey = 1", 6
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "Õ—ﬂ… ’‰œÊﬁ «·ﬂ«‘Ì—: —’Ìœ ¬Œ— «·„œ… = 770", _
        "SELECT Sum(AmountIn) - Sum(AmountOut) FROM CashStatementQuery", 770
    SetQueryParam "CashBoxID", CLng(R("BOXM"))
    Chk "Õ—ﬂ… «·Œ“Ì‰…: «· ÕÊÌ· „‰ «·ﬂ«‘Ì— œ«Œ· = 1000", _
        "SELECT AmountIn FROM CashStatementQuery WHERE MoveType = 'TRANSFER_IN'", 1000
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "ÌÊ„Ì… ’‰œÊﬁ «·ﬂ«‘Ì— ÌÊ„ «· ’›Ì…: —’Ìœ √Ê· «·ÌÊ„ = 565 + 1150 + 100 - 230", _
        "SELECT OpeningBalance FROM CashDailyQuery WHERE CashDay = DateValue(" & D(4, 0) & ")", 1585
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "ÌÊ„Ì… ’‰œÊﬁ «·ﬂ«‘Ì— ÌÊ„ «· ’›Ì…: «·„œ›Ê⁄«  = 15 ⁄Ã“ + 1000  ÕÊÌ·", _
        "SELECT Payments FROM CashDailyQuery WHERE CashDay = DateValue(" & D(4, 0) & ")", 1015
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "ÌÊ„Ì… ’‰œÊﬁ «·ﬂ«‘Ì— ÌÊ„ «· ’›Ì…: —’Ìœ ¬Œ— «·ÌÊ„ = 570", _
        "SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue(" & D(4, 0) & ")", 570
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk "ÌÊ„Ì… ’‰œÊﬁ «·ﬂ«‘Ì—: ¬Œ— ÌÊ„ = «·—’Ìœ «·Õ«·Ì", _
        "SELECT ClosingBalance FROM CashDailyQuery WHERE CashDay = DateValue(" & D(3, 0) & ")", 770
    SetQueryParam "CashBoxID", CLng(R("BOXC"))
    Chk " ’›Ì«  «·ﬂ«‘Ì— Œ·«· «·› —…:  ’›Ì… Ê«Õœ… »⁄Ã“ 15", _
        "SELECT Difference FROM CashClosingsQuery", -15
    Chk "ÿ»«⁄… ”‰œ ’—› «·„’—Ê›: ‰Ê⁄ «·„’—Ê›", _
        "SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = " & R("V1") & " AND ExpenseTypeName = '„’—Ê›«  √Œ—Ï'", 1
    Chk "ÿ»«⁄… ”‰œ «· ÕÊÌ·: «·’‰œÊﬁ «·„” ·„", _
        "SELECT COUNT(*) FROM qryCashVoucherPrint WHERE DocID = " & R("V3") & " AND ToBoxName = 'TEST «·Œ“Ì‰…'", 1
    Chk "ﬁÌÊœ Sale: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSale GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ SalesReturn: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalSalesReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ Purchase: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchase GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ PurchaseReturn: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPurchaseReturn GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ Payments: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalPayments GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ Expense: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalExpense GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ CashVoucher: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalCashVoucher GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ Stock: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalStock GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌÊœ Opening: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalOpening GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "ﬁÌœ «·›« Ê—… «·¬Ã·…: 6 √”ÿ— (‰ﬁœÌ° ⁄„Ì·° „»Ì⁄« ° ÷—Ì»…°  ﬂ·›…° „Œ“Ê‰)", _
        "SELECT COUNT(*) FROM qryJournalSale WHERE SourceID = " & R("INV2"), 6
    Chk "ﬁÌœ «·›« Ê—… «·¬Ã·…: «·„ »ﬁÌ ⁄·Ï «·⁄„Ì· 360 ›Ì –„„ «·⁄„·«¡", _
        "SELECT Debit FROM qryJournalSale WHERE SourceID = " & R("INV2") & " AND AccountCode = 1300", 360
    Chk "ﬁÌœ «·›« Ê—… «·¬Ã·…: «· ﬂ·›… = 2◊60 + 10◊10", _
        "SELECT Debit FROM qryJournalSale WHERE SourceID = " & R("INV2") & " AND AccountCode = 5100", 220
    Chk "–„„ «·⁄„·«¡ „‰ «·ﬁÌÊœ = √—’œ… «·⁄„·«¡ (164)", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 1300", 164
    Chk "–„„ «·„Ê—œÌ‰ „‰ «·ﬁÌÊœ = —’Ìœ «·„Ê—œ (2855 œ«∆‰)", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 2100", -2855
    Chk "«·„Œ“Ê‰ „‰ «·ﬁÌÊœ = ﬁÌ„… «·„Œ“Ê‰ »«· ﬂ·›… (7040)", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock) AS x WHERE AccountCode = 1400", 7040
    Chk "’‰œÊﬁ «·ﬂ«‘Ì— „‰ «·ﬁÌÊœ = —’ÌœÂ (770)", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + " & R("BOXC"), 770
    Chk "«·Œ“Ì‰… „‰ «·ﬁÌÊœ = —’ÌœÂ« (7700)", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchase UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPurchaseReturn UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalPayments UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalExpense UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalCashVoucher UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalStock UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalOpening) AS x WHERE AccountCode = 110000 + " & R("BOXM"), 7700
    Chk "÷—Ì»… «·„Œ—Ã«  „‰ «·ﬁÌÊœ = 15 + 150 + 60 - 6", _
        "SELECT Sum(Debit) - Sum(Credit) FROM (SELECT AccountCode, Debit, Credit FROM qryJournalSale UNION ALL SELECT AccountCode, Debit, Credit FROM qryJournalSalesReturn) AS x WHERE AccountCode = 2200", -219
    Chk "„’—Ê› ”‰œ «·‰ﬁœÌ… ·« ÌıﬁÌÛ¯œ „— Ì‰", _
        "SELECT COUNT(*) FROM qryJournalExpense WHERE SourceID = " & R("EXPV"), 0
    Chk "«·ﬁÌœ «·ÌœÊÌ: 3 √”ÿ— „ Ê«“‰… (3000)", _
        "SELECT COUNT(*) FROM qryJournalManual WHERE SourceID = " & R("MJ1") & " AND SourceType = 'MANUAL'", 3
    Chk "ﬁÌÊœ Manual: ﬂ· ﬁÌœ „ Ê«“‰", _
        "SELECT COUNT(*) FROM (SELECT SourceType, SourceID FROM qryJournalManual GROUP BY SourceType, SourceID HAVING Abs(Sum(Debit) - Sum(Credit)) > 0.001) AS x", 0
    Chk "”‰œ ’—› «·„’—Ê› ÌıﬁÌÛ¯œ ⁄·Ï Õ”«» ‰Ê⁄ «·„’—Ê›", _
        "SELECT Debit FROM qryJournalCashVoucher WHERE SourceID = " & R("V1") & " AND AccountCode = 530009", 50
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
    Q_SalesByCategoryQuery
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
    Q_qryVatReturnLines
    Q_qryVatReturnTotals
    Q_qryVatReturnHead
    Q_VatReturnQuery
    Q_DashboardQuery
    Q_qryDashboardTopProducts
    Q_qrySalesDocPrint
    Q_qryPurchaseDocPrint
    Q_qryVoucherPrint
    Q_qryCashMovements
    Q_qryCashBoxTotals
    Q_CashBoxBalanceQuery
    Q_CashStatementQuery
    Q_qryCashDays
    Q_qryCashDayOpening
    Q_CashDailyQuery
    Q_CashClosingsQuery
    Q_qryCashClosingPrint
    Q_qryCashVoucherPrint
    Q_qrySaleCost
    Q_qryReturnCost
    Q_qryStockCountValue
    Q_qryJournalSale
    Q_qryJournalSalesReturn
    Q_qryJournalPurchase
    Q_qryJournalPurchaseReturn
    Q_qryJournalPayments
    Q_qryJournalExpense
    Q_qryJournalCashVoucher
    Q_qryJournalStock
    Q_qryJournalOpening
    Q_qryManualEntryLines
    Q_qryJournalManual
    Q_qryYearCloseLines
    Q_qryJournalYearClose
    Q_qryJournalVatReturn
    Q_JournalLinesQuery
    Q_qryJournalEntryPrint
    Q_qryTrialBefore
    Q_qryTrialPeriod
    Q_TrialBalanceQuery
    Q_qryStatementBefore
    Q_AccountStatementQuery
    Q_GeneralLedgerQuery
    Q_qryTreeRollup
    Q_TrialBalanceTreeQuery
    Q_qryIncomeMoves
    Q_qryCompareMoves
    Q_qryIncomeAccounts
    Q_IncomeStatementQuery
    Q_qryBalanceAt
    Q_qryBalanceCompare
    Q_qryBalanceAccounts
    Q_qryProfitAt
    Q_qryProfitCompare
    Q_qryBalanceItems
    Q_BalanceSheetQuery
    Q_AccountTreeQuery
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

Private Sub Q_SalesByCategoryQuery()
    Dim s As String
    s = "SELECT CategoryName, Count(*) AS ProductCount, Sum(NetQty) AS CategoryQty," & vbCrLf
    s = s & "       Sum(NetSales) AS CategoryNet, Sum(SalesTotal) AS CategoryTotal, Sum(GrossProfit) AS CategoryProfit" & vbCrLf
    s = s & "FROM SalesByProductQuery" & vbCrLf
    s = s & "GROUP BY CategoryName" & vbCrLf
    s = s & "ORDER BY Sum(SalesTotal) DESC" & vbCrLf
    SaveQuery "SalesByCategoryQuery", "«·„»Ì⁄«  Õ”» «· ’‰Ì› Œ·«· › —… (··—”„ «·œ«∆—Ì)", s
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
    s = s & "WHERE p.IsActive = True AND p.TrackStock = True AND p.CurrentQuantity <= p.MinimumQuantity" & vbCrLf
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

Private Sub Q_StockCountQuery()
    Dim s As String
    s = "SELECT c.StockCountID, c.CountNumber, c.CountDate, c.Status, c.CategoryID, g.CategoryName," & vbCrLf
    s = s & "       d.ProductID, p.ProductCode, p.ProductName, d.SystemQuantity, d.ActualQuantity, d.Difference," & vbCrLf
    s = s & "       d.UnitCost, d.DifferenceValue, d.Notes" & vbCrLf
    s = s & "FROM ((StockCountDetails AS d INNER JOIN StockCounts AS c ON d.StockCountID = c.StockCountID)" & vbCrLf
    s = s & "      INNER JOIN Products AS p ON d.ProductID = p.ProductID)" & vbCrLf
    s = s & "     LEFT JOIN Categories AS g ON c.CategoryID = g.CategoryID" & vbCrLf
    s = s & "ORDER BY c.StockCountID, p.ProductName" & vbCrLf
    SaveQuery "StockCountQuery", " ›«’Ì· Ã·”«  «·Ã—œ: «·ﬂ„Ì… «·„”Ã·… Ê«·›⁄·Ì… Ê«·›—ﬁ ÊﬁÌ„ Â", s
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
    s = s & "       e.TotalAmount, pm.MethodName, e.Description, em.EmployeeName, e.ExpenseTypeID" & vbCrLf
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

Private Sub Q_qryVatReturnLines()
    Dim s As String
    s = "SELECT 'S' AS Side, d.VATCategory AS Category, h.InvoiceDate AS DocDate, d.NetAmount AS Amount," & vbCrLf
    s = s & "       CCur(0) AS Adjust, d.Tax AS VAT" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN SalesInvoiceDetails AS d ON h.SalesInvoiceID = d.SalesInvoiceID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'S', d.VATCategory, r.ReturnDate, CCur(0), -d.NetAmount, -d.Tax" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN SalesReturnDetails AS d ON r.SalesReturnID = d.SalesReturnID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'P', IIf(d.VATRate > 0, 'S', 'Z'), h.InvoiceDate, d.NetAmount, CCur(0), d.Tax" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN PurchaseInvoiceDetails AS d ON h.PurchaseInvoiceID = d.PurchaseInvoiceID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'P', IIf(d.VATRate > 0, 'S', 'Z'), r.ReturnDate, CCur(0), -d.NetAmount, -d.Tax" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN PurchaseReturnDetails AS d ON r.PurchaseReturnID = d.PurchaseReturnID" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'P', 'S', e.ExpenseDate, e.Amount, CCur(0), e.Tax" & vbCrLf
    s = s & "FROM Expenses AS e" & vbCrLf
    s = s & "WHERE e.Tax <> 0" & vbCrLf
    SaveQuery "qryVatReturnLines", "√”ÿ— «·≈ﬁ—«— «·÷—Ì»Ì: «·„»Ì⁄«  Ê«·„‘ —Ì«  Ê«·„’—Ê›«  »›∆ Â« «·÷—Ì»Ì…", s
End Sub

Private Sub Q_qryVatReturnTotals()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', Amount, 0)), 0)) AS SalesStdAmount," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', Adjust, 0)), 0)) AS SalesStdAdjust," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'S', VAT, 0)), 0)) AS SalesStdVAT," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'Z', Amount, 0)), 0)) AS SalesZeroAmount," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'Z', Adjust, 0)), 0)) AS SalesZeroAdjust," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'E', Amount, 0)), 0)) AS SalesExemptAmount," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'S' AND Category = 'E', Adjust, 0)), 0)) AS SalesExemptAdjust," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', Amount, 0)), 0)) AS PurchStdAmount," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', Adjust, 0)), 0)) AS PurchStdAdjust," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'S', VAT, 0)), 0)) AS PurchStdVAT," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'Z', Amount, 0)), 0)) AS PurchZeroAmount," & vbCrLf
    s = s & "       CCur(Nz(Sum(IIf(Side = 'P' AND Category = 'Z', Adjust, 0)), 0)) AS PurchZeroAdjust" & vbCrLf
    s = s & "FROM qryVatReturnLines" & vbCrLf
    s = s & "WHERE DocDate >= QDate('PeriodStart') AND DocDate < QDate('PeriodEnd')" & vbCrLf
    SaveQuery "qryVatReturnTotals", "Œ«‰«  «·≈ﬁ—«— «·÷—Ì»Ì ··› —… „Õ”Ê»… „‰ «·„” ‰œ«  (’› Ê«Õœ)", s
End Sub

Private Sub Q_qryVatReturnHead()
    Dim s As String
    s = "SELECT * FROM VatReturns" & vbCrLf
    s = s & "WHERE VatReturnID = QLong('VatReturnID')" & vbCrLf
    SaveQuery "qryVatReturnHead", "«·≈ﬁ—«— «·÷—Ì»Ì «·„Œ «—", s
End Sub

Private Sub Q_VatReturnQuery()
    Dim s As String
    s = "SELECT 1 AS BoxNo, '«·„»Ì⁄«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… (15%)' AS BoxText, v.SalesStdAmount AS Amount, v.SalesStdAdjust AS Adjust, v.SalesStdVAT AS VAT, 'L' AS RowKind, v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 2, '«·„»Ì⁄«  ··„Ê«ÿ‰Ì‰ («·Œœ„«  «·’ÕÌ… «·Œ«’… Ê«· ⁄·Ì„ «·√Â·Ì Ê«·„”ﬂ‰ «·√Ê·)', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 3, '«·„»Ì⁄«  «·„Õ·Ì… «·Œ«÷⁄… ··‰”»… «·’›—Ì…', v.SalesZeroAmount, v.SalesZeroAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 4, '«·’«œ—« ', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 5, '«·„»Ì⁄«  «·„⁄›«…', v.SalesExemptAmount, v.SalesExemptAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 6, '≈Ã„«·Ì «·„»Ì⁄« ', v.SalesStdAmount + v.SalesZeroAmount + v.SalesExemptAmount, v.SalesStdAdjust + v.SalesZeroAdjust + v.SalesExemptAdjust, v.SalesStdVAT, 'T', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 7, '«·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… („⁄ «·„’—Ê›«  »›« Ê—… ÷—Ì»Ì…)', v.PurchStdAmount, v.PurchStdAdjust, v.PurchStdVAT, 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 8, '«·«” Ì—«œ«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì… Ê«·„œ›Ê⁄… ÷—Ì» Â« ›Ì «·Ã„«—ﬂ', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 9, '«·«” Ì—«œ«  «·Œ«÷⁄… ··÷—Ì»… »¬·Ì… «·«Õ ”«» «·⁄ﬂ”Ì', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 10, '«·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·’›—Ì…', v.PurchZeroAmount, v.PurchZeroAdjust, CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 11, '«·„‘ —Ì«  «·„⁄›«…', CCur(0), CCur(0), CCur(0), 'L', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 12, '≈Ã„«·Ì «·„‘ —Ì« ', v.PurchStdAmount + v.PurchZeroAmount, v.PurchStdAdjust + v.PurchZeroAdjust, v.PurchStdVAT, 'T', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 13, '≈Ã„«·Ì ÷—Ì»… «·ﬁÌ„… «·„÷«›… «·„” Õﬁ… ⁄‰ «·› —… «·Õ«·Ì…', Null, Null, v.SalesStdVAT - v.PurchStdVAT, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 14, ' ’ÕÌÕ«  „‰ «·› —«  «·”«»ﬁ…', Null, Null, v.Corrections, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 15, '÷—Ì»… «·ﬁÌ„… «·„÷«›… «·„—ÕÛ¯·… „‰ «·› —«  «·”«»ﬁ… (—’Ìœ œ«∆‰)', Null, Null, v.CarriedCredit, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 16, '’«›Ì «·÷—Ì»… «·„” Õﬁ… (”«·» = „” —œ…)', Null, Null, v.NetDue, 'N', v.VatReturnID, v.ReturnNumber, v.PeriodFrom, v.PeriodTo, v.Status, v.FiledDate, v.FilingRef, v.PaidDate, v.PaidAmount" & vbCrLf
    s = s & "FROM qryVatReturnHead AS v" & vbCrLf
    SaveQuery "VatReturnQuery", "≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›… »Œ«‰«  ‰„Ê–Ã «·ÂÌ∆… (1 ≈·Ï 16)", s
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
    SaveQuery "DashboardQuery", "„ƒ‘—«  ·ÊÕ… «· Õﬂ„ ›Ì ”Ã· Ê«Õœ («·ÌÊ„° «·‘Â—° «·√—’œ…° «·„Œ“Ê‰)", s
End Sub

Private Sub Q_qryDashboardTopProducts()
    Dim s As String
    s = "SELECT l.ProductID, p.ProductName, Sum(l.SignedQty) AS NetQty, Sum(l.LineGross) AS NetSales" & vbCrLf
    s = s & "FROM qrySalesLineItems AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID" & vbCrLf
    s = s & "WHERE l.DocDate >= QDate('DashMonth') AND l.DocDate < QDate('DashEnd')" & vbCrLf
    s = s & "GROUP BY l.ProductID, p.ProductName" & vbCrLf
    s = s & "HAVING Sum(l.SignedQty) > 0" & vbCrLf
    SaveQuery "qryDashboardTopProducts", "’«›Ì «·ﬂ„Ì… «·„»«⁄… ·ﬂ· „‰ Ã „‰– »œ«Ì… «·‘Â— (·ÊÕ… «· Õﬂ„)", s
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
    s = s & "       d.Discount AS LineDiscount, d.NetAmount, d.VATRate, d.Tax AS LineTax, d.LineTotal," & vbCrLf
    s = s & "       h.OrderType, h.TableNo, h.OrderName, d.LineNote" & vbCrLf
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
    s = s & "       rd.Discount, rd.NetAmount, rd.VATRate, rd.Tax, rd.LineTotal, o.OrderType, o.TableNo, o.OrderName," & vbCrLf
    s = s & "       od.LineNote" & vbCrLf
    s = s & "FROM ((((((SalesReturns AS r INNER JOIN SalesReturnDetails AS rd ON r.SalesReturnID = rd.SalesReturnID)" & vbCrLf
    s = s & "        INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = o.SalesInvoiceID)" & vbCrLf
    s = s & "        INNER JOIN SalesInvoiceDetails AS od ON rd.SalesDetailID = od.SalesDetailID)" & vbCrLf
    s = s & "       INNER JOIN Products AS p ON rd.ProductID = p.ProductID)" & vbCrLf
    s = s & "      INNER JOIN Units AS u ON p.UnitID = u.UnitID)" & vbCrLf
    s = s & "     INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID)" & vbCrLf
    s = s & "    INNER JOIN Employees AS e ON r.EmployeeID = e.EmployeeID" & vbCrLf
    SaveQuery "qrySalesDocPrint", "»Ì«‰«  ÿ»«⁄… ›Ê« Ì— «·»Ì⁄ Ê«·≈‘⁄«—«  «·œ«∆‰… (”ÿ— ·ﬂ· ’‰›)", s
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
    SaveQuery "qryPurchaseDocPrint", "»Ì«‰«  ÿ»«⁄… ›Ê« Ì— «·‘—«¡ Ê„— Ã⁄« Â« (”ÿ— ·ﬂ· ’‰›)", s
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
    SaveQuery "qryVoucherPrint", "»Ì«‰«  ÿ»«⁄… ”‰œ«  «·ﬁ»÷ („‰ «·⁄„·«¡) Ê”‰œ«  «·’—› (··„Ê—œÌ‰)", s
End Sub

Private Sub Q_qryCashMovements()
    Dim s As String
    s = "SELECT h.CashBoxID, h.InvoiceDate AS MoveDate, 'SALE' AS MoveType, '›« Ê—… »Ì⁄' AS MoveTypeName," & vbCrLf
    s = s & "       h.InvoiceNumber AS DocNumber, c.CustomerName AS PartyName, h.Notes AS Details," & vbCrLf
    s = s & "       h.PaidAmount AS AmountIn, CCur(0) AS AmountOut, h.EmployeeID" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.CashBoxID, r.ReturnDate, 'SALES_RETURN', '„— Ã⁄ »Ì⁄ (—œ ‰ﬁœÌ)', r.ReturnNumber," & vbCrLf
    s = s & "       c.CustomerName, r.Reason, CCur(0), r.RefundedAmount, r.EmployeeID" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.CashBoxID, p.PaymentDate, 'CUSTOMER_PAYMENT', '”‰œ ﬁ»÷ „‰ ⁄„Ì·', p.PaymentNumber," & vbCrLf
    s = s & "       c.CustomerName, p.Notes, p.Amount, CCur(0), p.EmployeeID" & vbCrLf
    s = s & "FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE p.CashBoxID Is Not Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT h.CashBoxID, h.InvoiceDate, 'PURCHASE', '›« Ê—… ‘—«¡', h.InvoiceNumber," & vbCrLf
    s = s & "       s.SupplierName, h.Notes, CCur(0), h.PaidAmount, h.EmployeeID" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE h.CashBoxID Is Not Null AND h.PaidAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT r.CashBoxID, r.ReturnDate, 'PURCHASE_RETURN', '„— Ã⁄ ‘—«¡ («” —œ«œ ‰ﬁœÌ)', r.ReturnNumber," & vbCrLf
    s = s & "       s.SupplierName, r.Reason, r.RefundedAmount, CCur(0), r.EmployeeID" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE r.CashBoxID Is Not Null AND r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT p.CashBoxID, p.PaymentDate, 'SUPPLIER_PAYMENT', '”‰œ ’—› ·„Ê—œ', p.PaymentNumber," & vbCrLf
    s = s & "       s.SupplierName, p.Notes, CCur(0), p.Amount, p.EmployeeID" & vbCrLf
    s = s & "FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE p.CashBoxID Is Not Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT e.CashBoxID, e.ExpenseDate, 'EXPENSE', '„’—Ê›', e.ExpenseNumber," & vbCrLf
    s = s & "       t.ExpenseTypeName, e.Description, CCur(0), e.TotalAmount, e.EmployeeID" & vbCrLf
    s = s & "FROM Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID" & vbCrLf
    s = s & "WHERE e.CashBoxID Is Not Null" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT v.CashBoxID, v.VoucherDate, 'CASH_IN'," & vbCrLf
    s = s & "       IIf(v.Category = 'OWNER', '≈Ìœ«⁄ „‰ «·„«·ﬂ', IIf(v.Category = 'OVERAGE', '“Ì«œ… ›Ì «·’‰œÊﬁ'," & vbCrLf
    s = s & "           '”‰œ ﬁ»÷ ‰ﬁœÌ…'))," & vbCrLf
    s = s & "       v.VoucherNumber, v.PartyName, v.Description, v.Amount, CCur(0), v.EmployeeID" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'IN'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT v.CashBoxID, v.VoucherDate, 'CASH_OUT'," & vbCrLf
    s = s & "       IIf(v.Category = 'OWNER', ' ”ÊÌ… „⁄ «·„«·ﬂ', IIf(v.Category = 'EXPENSE', '„’—Ê› (”‰œ ’—›)'," & vbCrLf
    s = s & "           IIf(v.Category = 'ADVANCE', '”·›… „ÊŸ›', IIf(v.Category = 'SHORTAGE', '⁄Ã“ ›Ì «·’‰œÊﬁ'," & vbCrLf
    s = s & "           '”‰œ ’—› ‰ﬁœÌ…'))))," & vbCrLf
    s = s & "       v.VoucherNumber, v.PartyName, v.Description, CCur(0), v.Amount, v.EmployeeID" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'OUT'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT v.CashBoxID, v.VoucherDate, 'TRANSFER_OUT', ' ÕÊÌ· ≈·Ï ’‰œÊﬁ ¬Œ—', v.VoucherNumber," & vbCrLf
    s = s & "       b.BoxName, v.Description, CCur(0), v.Amount, v.EmployeeID" & vbCrLf
    s = s & "FROM CashVouchers AS v INNER JOIN CashBoxes AS b ON v.ToCashBoxID = b.CashBoxID" & vbCrLf
    s = s & "WHERE v.VoucherType = 'TRANSFER'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT v.ToCashBoxID, v.VoucherDate, 'TRANSFER_IN', ' ÕÊÌ· „‰ ’‰œÊﬁ ¬Œ—', v.VoucherNumber," & vbCrLf
    s = s & "       b.BoxName, v.Description, v.Amount, CCur(0), v.EmployeeID" & vbCrLf
    s = s & "FROM CashVouchers AS v INNER JOIN CashBoxes AS b ON v.CashBoxID = b.CashBoxID" & vbCrLf
    s = s & "WHERE v.VoucherType = 'TRANSFER'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT b.CashBoxID, b.OpeningDate, 'OPENING', '—’Ìœ «›  «ÕÌ', '-', b.BoxName, b.Notes," & vbCrLf
    s = s & "       b.OpeningBalance, CCur(0), Null" & vbCrLf
    s = s & "FROM CashBoxes AS b" & vbCrLf
    s = s & "WHERE b.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qryCashMovements", "ﬂ· Õ—ﬂ«  «·‰ﬁœÌ… ›Ì «·Œ“Ì‰… Ê«·’‰«œÌﬁ: œ«Œ· (+) ÊŒ«—Ã (-)", s
End Sub

Private Sub Q_qryCashBoxTotals()
    Dim s As String
    s = "SELECT CashBoxID, Sum(AmountIn) AS BoxIn, Sum(AmountOut) AS BoxOut, Max(MoveDate) AS LastMoveDate" & vbCrLf
    s = s & "FROM qryCashMovements" & vbCrLf
    s = s & "GROUP BY CashBoxID" & vbCrLf
    SaveQuery "qryCashBoxTotals", "≈Ã„«·Ì «·œ«Œ· Ê«·Œ«—Ã ·ﬂ· ’‰œÊﬁ", s
End Sub

Private Sub Q_CashBoxBalanceQuery()
    Dim s As String
    s = "SELECT b.CashBoxID, b.BoxName, b.BoxType," & vbCrLf
    s = s & "       IIf(b.BoxType = 'MAIN', 'Œ“Ì‰… —∆Ì”Ì…', '’‰œÊﬁ ﬂ«‘Ì—') AS BoxTypeName, b.IsActive," & vbCrLf
    s = s & "       CCur(Nz(t.BoxIn, 0)) AS TotalIn, CCur(Nz(t.BoxOut, 0)) AS TotalOut," & vbCrLf
    s = s & "       CCur(Nz(t.BoxIn, 0)) - CCur(Nz(t.BoxOut, 0)) AS Balance, t.LastMoveDate" & vbCrLf
    s = s & "FROM CashBoxes AS b LEFT JOIN qryCashBoxTotals AS t ON b.CashBoxID = t.CashBoxID" & vbCrLf
    s = s & "ORDER BY b.BoxType DESC, b.BoxName" & vbCrLf
    SaveQuery "CashBoxBalanceQuery", "√—’œ… «·Œ“Ì‰… Ê«·’‰«œÌﬁ «·¬‰", s
End Sub

Private Sub Q_CashStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, m.MoveDate, m.MoveType, m.MoveTypeName, m.DocNumber, m.PartyName, m.Details," & vbCrLf
    s = s & "       b.BoxName, m.AmountIn, m.AmountOut, m.CashBoxID" & vbCrLf
    s = s & "FROM qryCashMovements AS m INNER JOIN CashBoxes AS b ON m.CashBoxID = b.CashBoxID" & vbCrLf
    s = s & "WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND m.MoveDate >= QDate('PeriodStart') AND m.MoveDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, QDate('PeriodStart'), 'BALANCE_FWD', '—’Ìœ √Ê· «·„œ…', '-', Null, Null, Null," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)) > 0, CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)), 0)," & vbCrLf
    s = s & "       IIf(CCur(Nz(Sum(o.AmountIn), 0)) - CCur(Nz(Sum(o.AmountOut), 0)) < 0, CCur(Nz(Sum(o.AmountOut), 0)) - CCur(Nz(Sum(o.AmountIn), 0)), 0)," & vbCrLf
    s = s & "       QLong('CashBoxID')" & vbCrLf
    s = s & "FROM qryCashMovements AS o" & vbCrLf
    s = s & "WHERE (QLong('CashBoxID') = 0 OR o.CashBoxID = QLong('CashBoxID')) AND o.MoveDate < QDate('PeriodStart')" & vbCrLf
    s = s & "ORDER BY SortKey, MoveDate" & vbCrLf
    SaveQuery "CashStatementQuery", "Õ—ﬂ… «·Œ“Ì‰… / «·’‰œÊﬁ ·› —…: —’Ìœ √Ê· «·„œ… À„ «·Õ—ﬂ«  (0 = ﬂ· «·’‰«œÌﬁ)", s
End Sub

Private Sub Q_qryCashDays()
    Dim s As String
    s = "SELECT DateValue(m.MoveDate) AS CashDay, Sum(m.AmountIn) AS Receipts, Sum(m.AmountOut) AS Payments," & vbCrLf
    s = s & "       Count(*) AS MoveCount" & vbCrLf
    s = s & "FROM qryCashMovements AS m" & vbCrLf
    s = s & "WHERE (QLong('CashBoxID') = 0 OR m.CashBoxID = QLong('CashBoxID')) AND m.MoveDate >= QDate('PeriodStart') AND m.MoveDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "GROUP BY DateValue(m.MoveDate)" & vbCrLf
    SaveQuery "qryCashDays", "„ﬁ»Ê÷«  Ê„œ›Ê⁄«  ﬂ· ÌÊ„ œ«Œ· «·› —…", s
End Sub

Private Sub Q_qryCashDayOpening()
    Dim s As String
    s = "SELECT d.CashDay, Sum(x.AmountIn - x.AmountOut) AS DayOpening" & vbCrLf
    s = s & "FROM qryCashDays AS d, qryCashMovements AS x" & vbCrLf
    s = s & "WHERE (QLong('CashBoxID') = 0 OR x.CashBoxID = QLong('CashBoxID')) AND x.MoveDate < d.CashDay" & vbCrLf
    s = s & "GROUP BY d.CashDay" & vbCrLf
    SaveQuery "qryCashDayOpening", "—’Ìœ √Ê· ﬂ· ÌÊ„ „‰ √Ì«„ «·Õ—ﬂ… (ﬂ· «·Õ—ﬂ«  ﬁ»· –·ﬂ «·ÌÊ„)", s
End Sub

Private Sub Q_CashDailyQuery()
    Dim s As String
    s = "SELECT d.CashDay, CCur(Nz(o.DayOpening, 0)) AS OpeningBalance, d.Receipts, d.Payments," & vbCrLf
    s = s & "       CCur(Nz(o.DayOpening, 0)) + d.Receipts - d.Payments AS ClosingBalance, d.MoveCount" & vbCrLf
    s = s & "FROM qryCashDays AS d LEFT JOIN qryCashDayOpening AS o ON d.CashDay = o.CashDay" & vbCrLf
    s = s & "ORDER BY d.CashDay" & vbCrLf
    SaveQuery "CashDailyQuery", "Õ—ﬂ… «·Œ“Ì‰… «·ÌÊ„Ì…: —’Ìœ √Ê· «·ÌÊ„ Ê«·„ﬁ»Ê÷«  Ê«·„œ›Ê⁄«  Ê—’Ìœ ¬Œ— «·ÌÊ„", s
End Sub

Private Sub Q_CashClosingsQuery()
    Dim s As String
    s = "SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart," & vbCrLf
    s = s & "       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference," & vbCrLf
    s = s & "       IIf(c.Destination = 'MAIN', '«·Œ“Ì‰… «·—∆Ì”Ì…', IIf(c.Destination = 'OWNER', ' ”ÊÌ… „⁄ «·„«·ﬂ'," & vbCrLf
    s = s & "           'Ì»ﬁÏ ›Ì «·’‰œÊﬁ')) AS DestinationName," & vbCrLf
    s = s & "       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes, c.CashBoxID" & vbCrLf
    s = s & "FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)" & vbCrLf
    s = s & "      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)" & vbCrLf
    s = s & "     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID" & vbCrLf
    s = s & "WHERE (QLong('CashBoxID') = 0 OR c.CashBoxID = QLong('CashBoxID')) AND c.ClosingDate >= QDate('PeriodStart') AND c.ClosingDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY c.ClosingDate" & vbCrLf
    SaveQuery "CashClosingsQuery", " ’›Ì«  ÌÊ„Ì… «·ﬂ«‘Ì— Œ·«· › —… (0 = ﬂ· «·’‰«œÌﬁ)", s
End Sub

Private Sub Q_qryCashClosingPrint()
    Dim s As String
    s = "SELECT c.ClosingID, c.ClosingNumber, c.ClosingDate, b.BoxName, e.EmployeeName, c.PeriodStart," & vbCrLf
    s = s & "       c.OpeningBalance, c.CashIn, c.CashOut, c.ExpectedBalance, c.CountedAmount, c.Difference," & vbCrLf
    s = s & "       IIf(c.Destination = 'MAIN', '«·Œ“Ì‰… «·—∆Ì”Ì…', IIf(c.Destination = 'OWNER', ' ”ÊÌ… „⁄ «·„«·ﬂ'," & vbCrLf
    s = s & "           'Ì»ﬁÏ ›Ì «·’‰œÊﬁ')) AS DestinationName," & vbCrLf
    s = s & "       t.BoxName AS ToBoxName, c.TransferAmount, c.KeptAmount, c.Notes" & vbCrLf
    s = s & "FROM ((CashClosings AS c INNER JOIN CashBoxes AS b ON c.CashBoxID = b.CashBoxID)" & vbCrLf
    s = s & "      INNER JOIN Employees AS e ON c.EmployeeID = e.EmployeeID)" & vbCrLf
    s = s & "     LEFT JOIN CashBoxes AS t ON c.ToCashBoxID = t.CashBoxID" & vbCrLf
    SaveQuery "qryCashClosingPrint", "»Ì«‰«  ÿ»«⁄…  ’›Ì… «·ﬂ«‘Ì—", s
End Sub

Private Sub Q_qryCashVoucherPrint()
    Dim s As String
    s = "SELECT v.CashVoucherID AS DocID, v.VoucherNumber, v.VoucherDate, v.VoucherType," & vbCrLf
    s = s & "       IIf(v.VoucherType = 'IN', '”‰œ ﬁ»÷ ‰ﬁœÌ…', IIf(v.VoucherType = 'OUT', '”‰œ ’—› ‰ﬁœÌ…'," & vbCrLf
    s = s & "           '”‰œ  ÕÊÌ· ‰ﬁœÌ…')) AS VoucherTitle," & vbCrLf
    s = s & "       IIf(v.Category = 'OWNER', IIf(v.VoucherType = 'IN', '≈Ìœ«⁄ „‰ «·„«·ﬂ', ' ”ÊÌ… „⁄ «·„«·ﬂ')," & vbCrLf
    s = s & "           IIf(v.Category = 'EXPENSE', '„’—Ê›', IIf(v.Category = 'ADVANCE', '”·›… „ÊŸ›'," & vbCrLf
    s = s & "           IIf(v.Category = 'SHORTAGE', '⁄Ã“ ›Ì «·’‰œÊﬁ', IIf(v.Category = 'OVERAGE', '“Ì«œ… ›Ì «·’‰œÊﬁ'," & vbCrLf
    s = s & "           IIf(v.Category = 'TRANSFER', ' ÕÊÌ· »Ì‰ «·’‰«œÌﬁ', '√Œ—Ï')))))) AS CategoryName," & vbCrLf
    s = s & "       b.BoxName, t.BoxName AS ToBoxName, v.Amount, v.PartyName, v.Description," & vbCrLf
    s = s & "       x.ExpenseTypeName, e.EmployeeName" & vbCrLf
    s = s & "FROM ((((CashVouchers AS v INNER JOIN CashBoxes AS b ON v.CashBoxID = b.CashBoxID)" & vbCrLf
    s = s & "        INNER JOIN Employees AS e ON v.EmployeeID = e.EmployeeID)" & vbCrLf
    s = s & "       LEFT JOIN CashBoxes AS t ON v.ToCashBoxID = t.CashBoxID)" & vbCrLf
    s = s & "      LEFT JOIN Expenses AS ex ON v.ExpenseID = ex.ExpenseID)" & vbCrLf
    s = s & "     LEFT JOIN ExpenseTypes AS x ON ex.ExpenseTypeID = x.ExpenseTypeID" & vbCrLf
    SaveQuery "qryCashVoucherPrint", "»Ì«‰«  ÿ»«⁄… ”‰œ«  ﬁ»÷ Ê’—› Ê ÕÊÌ· «·‰ﬁœÌ…", s
End Sub

Private Sub Q_qrySaleCost()
    Dim s As String
    s = "SELECT SalesInvoiceID, Sum(Quantity * UnitCost) AS SaleCost" & vbCrLf
    s = s & "FROM SalesInvoiceDetails" & vbCrLf
    s = s & "GROUP BY SalesInvoiceID" & vbCrLf
    SaveQuery "qrySaleCost", " ﬂ·›… ﬂ· ›« Ê—… »Ì⁄", s
End Sub

Private Sub Q_qryReturnCost()
    Dim s As String
    s = "SELECT SalesReturnID, Sum(IIf(ReturnToStock, Quantity * UnitCost, 0)) AS ReturnCost" & vbCrLf
    s = s & "FROM SalesReturnDetails" & vbCrLf
    s = s & "GROUP BY SalesReturnID" & vbCrLf
    SaveQuery "qryReturnCost", " ﬂ·›… „« ⁄«œ ··„Œ“Ê‰ „‰ ﬂ· „— Ã⁄ »Ì⁄", s
End Sub

Private Sub Q_qryStockCountValue()
    Dim s As String
    s = "SELECT ReferenceID AS StockCountID, Max(ReferenceNumber) AS CountNumber, Max(TransactionDate) AS CountDate," & vbCrLf
    s = s & "       Sum(Quantity * UnitCost) AS CountValue" & vbCrLf
    s = s & "FROM InventoryTransactions" & vbCrLf
    s = s & "WHERE ReferenceType = 'STOCK_COUNT'" & vbCrLf
    s = s & "GROUP BY ReferenceID" & vbCrLf
    SaveQuery "qryStockCountValue", "ﬁÌ„… ›—Êﬁ«  ﬂ· Ã—œ „ı—Õ¯·", s
End Sub

Private Sub Q_qryJournalSale()
    Dim s As String
    s = "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, 1200), 110000 + h.CashBoxID) AS AccountCode, h.PaidAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.PaidAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, h.RemainingAmount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.RemainingAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, 4100 AS AccountCode, CCur(0) AS Debit, h.TaxableAmount AS Credit, '«·„»Ì⁄« ' AS LineText" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.TaxableAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 2200 AS AccountCode, CCur(0) AS Debit, h.Tax AS Credit, '÷—Ì»… «·„Œ—Ã« ' AS LineText" & vbCrLf
    s = s & "FROM SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE h.Tax <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 5100 AS AccountCode, k.SaleCost AS Debit, CCur(0) AS Credit, ' ﬂ·›… «·»÷«⁄… «·„»«⁄…' AS LineText" & vbCrLf
    s = s & "FROM (SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID" & vbCrLf
    s = s & "WHERE k.SaleCost <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALE' AS SourceType, h.SalesInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, k.SaleCost AS Credit, '«·„Œ“Ê‰' AS LineText" & vbCrLf
    s = s & "FROM (SalesInvoices AS h INNER JOIN Customers AS c ON h.CustomerID = c.CustomerID) INNER JOIN qrySaleCost AS k ON h.SalesInvoiceID = k.SalesInvoiceID" & vbCrLf
    s = s & "WHERE k.SaleCost <> 0" & vbCrLf
    SaveQuery "qryJournalSale", "√”ÿ— ﬁÌÊœ ›Ê« Ì— «·»Ì⁄", s
End Sub

Private Sub Q_qryJournalSalesReturn()
    Dim s As String
    s = "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 4110 AS AccountCode, r.TaxableAmount AS Debit, CCur(0) AS Credit, '„—œÊœ«  «·„»Ì⁄« ' AS LineText" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE r.TaxableAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 2200 AS AccountCode, r.Tax AS Debit, CCur(0) AS Credit, '÷—Ì»… «·„Œ—Ã« ' AS LineText" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE r.Tax <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 3 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, 1200), 110000 + r.CashBoxID) AS AccountCode, CCur(0) AS Debit, r.RefundedAmount AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 4 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, r.TotalAmount - r.RefundedAmount AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE r.TotalAmount - r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 5 AS LineOrder, 1400 AS AccountCode, k.ReturnCost AS Debit, CCur(0) AS Credit, '«·„Œ“Ê‰' AS LineText" & vbCrLf
    s = s & "FROM (SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID" & vbCrLf
    s = s & "WHERE k.ReturnCost <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SALES_RETURN' AS SourceType, r.SalesReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, c.CustomerName AS Party, 6 AS LineOrder, 5100 AS AccountCode, CCur(0) AS Debit, k.ReturnCost AS Credit, ' ﬂ·›… «·»÷«⁄… «·„»«⁄…' AS LineText" & vbCrLf
    s = s & "FROM (SalesReturns AS r INNER JOIN Customers AS c ON r.CustomerID = c.CustomerID) INNER JOIN qryReturnCost AS k ON r.SalesReturnID = k.SalesReturnID" & vbCrLf
    s = s & "WHERE k.ReturnCost <> 0" & vbCrLf
    SaveQuery "qryJournalSalesReturn", "√”ÿ— ﬁÌÊœ „— Ã⁄«  «·»Ì⁄", s
End Sub

Private Sub Q_qryJournalPurchase()
    Dim s As String
    s = "SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 1400 AS AccountCode, h.TaxableAmount AS Debit, CCur(0) AS Credit, '«·„Œ“Ê‰' AS LineText" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE h.TaxableAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 1500 AS AccountCode, h.Tax AS Debit, CCur(0) AS Credit, '÷—Ì»… «·„œŒ·« ' AS LineText" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE h.Tax <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, IIf(h.CashBoxID Is Null, IIf(h.PaymentMethodID Is Null Or h.PaymentMethodID = 1, 1190, 1200), 110000 + h.CashBoxID) AS AccountCode, CCur(0) AS Debit, h.PaidAmount AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE h.PaidAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE' AS SourceType, h.PurchaseInvoiceID AS SourceID, h.InvoiceNumber AS SourceNumber, h.InvoiceDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 2100 AS AccountCode, CCur(0) AS Debit, h.RemainingAmount AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM PurchaseInvoices AS h INNER JOIN Suppliers AS s ON h.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE h.RemainingAmount <> 0" & vbCrLf
    SaveQuery "qryJournalPurchase", "√”ÿ— ﬁÌÊœ ›Ê« Ì— «·‘—«¡", s
End Sub

Private Sub Q_qryJournalPurchaseReturn()
    Dim s As String
    s = "SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, IIf(r.CashBoxID Is Null, IIf(r.PaymentMethodID Is Null Or r.PaymentMethodID = 1, 1190, 1200), 110000 + r.CashBoxID) AS AccountCode, r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 2100 AS AccountCode, r.TotalAmount - r.RefundedAmount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE r.TotalAmount - r.RefundedAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 3 AS LineOrder, 1400 AS AccountCode, CCur(0) AS Debit, r.TaxableAmount AS Credit, '«·„Œ“Ê‰' AS LineText" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE r.TaxableAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'PURCHASE_RETURN' AS SourceType, r.PurchaseReturnID AS SourceID, r.ReturnNumber AS SourceNumber, r.ReturnDate AS SourceDate, s.SupplierName AS Party, 4 AS LineOrder, 1500 AS AccountCode, CCur(0) AS Debit, r.Tax AS Credit, '÷—Ì»… «·„œŒ·« ' AS LineText" & vbCrLf
    s = s & "FROM PurchaseReturns AS r INNER JOIN Suppliers AS s ON r.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE r.Tax <> 0" & vbCrLf
    SaveQuery "qryJournalPurchaseReturn", "√”ÿ— ﬁÌÊœ „— Ã⁄«  «·‘—«¡", s
End Sub

Private Sub Q_qryJournalPayments()
    Dim s As String
    s = "SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, 1200), 110000 + p.CashBoxID) AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE p.Amount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CUSTOMER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 1300 AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM CustomerPayments AS p INNER JOIN Customers AS c ON p.CustomerID = c.CustomerID" & vbCrLf
    s = s & "WHERE p.Amount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, p.Amount AS Debit, CCur(0) AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE p.Amount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_PAYMENT' AS SourceType, p.PaymentID AS SourceID, p.PaymentNumber AS SourceNumber, p.PaymentDate AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, IIf(p.CashBoxID Is Null, IIf(p.PaymentMethodID Is Null Or p.PaymentMethodID = 1, 1190, 1200), 110000 + p.CashBoxID) AS AccountCode, CCur(0) AS Debit, p.Amount AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM SupplierPayments AS p INNER JOIN Suppliers AS s ON p.SupplierID = s.SupplierID" & vbCrLf
    s = s & "WHERE p.Amount <> 0" & vbCrLf
    SaveQuery "qryJournalPayments", "√”ÿ— ﬁÌÊœ ”‰œ«  «·ﬁ»÷ „‰ «·⁄„·«¡ Ê«·’—› ··„Ê—œÌ‰", s
End Sub

Private Sub Q_qryJournalExpense()
    Dim s As String
    s = "SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 1 AS LineOrder, 530000 + e.ExpenseTypeID AS AccountCode, e.Amount AS Debit, CCur(0) AS Credit, t.ExpenseTypeName AS LineText" & vbCrLf
    s = s & "FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID" & vbCrLf
    s = s & "WHERE v.CashVoucherID Is Null AND e.Amount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 2 AS LineOrder, 1500 AS AccountCode, e.Tax AS Debit, CCur(0) AS Credit, '÷—Ì»… «·„œŒ·« ' AS LineText" & vbCrLf
    s = s & "FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID" & vbCrLf
    s = s & "WHERE v.CashVoucherID Is Null AND e.Tax <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'EXPENSE' AS SourceType, e.ExpenseID AS SourceID, e.ExpenseNumber AS SourceNumber, e.ExpenseDate AS SourceDate, t.ExpenseTypeName AS Party, 3 AS LineOrder, IIf(e.CashBoxID Is Null, IIf(e.PaymentMethodID Is Null Or e.PaymentMethodID = 1, 1190, 1200), 110000 + e.CashBoxID) AS AccountCode, CCur(0) AS Debit, e.TotalAmount AS Credit, e.Description AS LineText" & vbCrLf
    s = s & "FROM (Expenses AS e INNER JOIN ExpenseTypes AS t ON e.ExpenseTypeID = t.ExpenseTypeID) LEFT JOIN CashVouchers AS v ON e.ExpenseID = v.ExpenseID" & vbCrLf
    s = s & "WHERE v.CashVoucherID Is Null AND e.TotalAmount <> 0" & vbCrLf
    SaveQuery "qryJournalExpense", "√”ÿ— ﬁÌÊœ «·„’—Ê›«  (⁄œ« «·„”Ã·… »”‰œ ‰ﬁœÌ…)", s
End Sub

Private Sub Q_qryJournalCashVoucher()
    Dim s As String
    s = "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.PartyName AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'IN'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'OVERAGE', 4300, 4200)) AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'IN'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, IIf(v.Category = 'OWNER', 3100, IIf(v.Category = 'ADVANCE', 1600, IIf(v.Category = 'SHORTAGE', 5400, IIf(v.Category = 'EXPENSE' AND x.ExpenseTypeID Is Not Null, 530000 + x.ExpenseTypeID, 5900)))) AS AccountCode, v.Amount - CCur(Nz(x.Tax, 0)) AS Debit, CCur(0) AS Credit, v.Description AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID" & vbCrLf
    s = s & "WHERE v.VoucherType = 'OUT'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.PartyName AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID" & vbCrLf
    s = s & "WHERE v.VoucherType = 'OUT'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 3 AS LineOrder, 1500 AS AccountCode, x.Tax AS Debit, CCur(0) AS Credit, '÷—Ì»… «·„œŒ·« ' AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v LEFT JOIN Expenses AS x ON v.ExpenseID = x.ExpenseID" & vbCrLf
    s = s & "WHERE v.VoucherType = 'OUT' AND x.Tax <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 1 AS LineOrder, 110000 + v.ToCashBoxID AS AccountCode, v.Amount AS Debit, CCur(0) AS Credit, v.Description AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'TRANSFER'" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CASH_VOUCHER' AS SourceType, v.CashVoucherID AS SourceID, v.VoucherNumber AS SourceNumber, v.VoucherDate AS SourceDate, IIf(v.PartyName Is Null, v.Description, v.PartyName) AS Party, 2 AS LineOrder, 110000 + v.CashBoxID AS AccountCode, CCur(0) AS Debit, v.Amount AS Credit, v.Description AS LineText" & vbCrLf
    s = s & "FROM CashVouchers AS v" & vbCrLf
    s = s & "WHERE v.VoucherType = 'TRANSFER'" & vbCrLf
    SaveQuery "qryJournalCashVoucher", "√”ÿ— ﬁÌÊœ ”‰œ«  «·‰ﬁœÌ… (ﬁ»÷ Ê’—› Ê ÕÊÌ·)", s
End Sub

Private Sub Q_qryJournalStock()
    Dim s As String
    s = "SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText" & vbCrLf
    s = s & "FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID" & vbCrLf
    s = s & "WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'STOCK_MOVE' AS SourceType, i.TransactionID AS SourceID, i.ReferenceNumber AS SourceNumber, i.TransactionDate AS SourceDate, p.ProductName AS Party, 2 AS LineOrder, IIf(i.TransactionTypeID = 8, 3900, 5200) AS AccountCode, IIf(i.Quantity * i.UnitCost < 0, -i.Quantity * i.UnitCost, 0) AS Debit, IIf(i.Quantity * i.UnitCost > 0, i.Quantity * i.UnitCost, 0) AS Credit, i.Notes AS LineText" & vbCrLf
    s = s & "FROM InventoryTransactions AS i INNER JOIN Products AS p ON i.ProductID = p.ProductID" & vbCrLf
    s = s & "WHERE i.ReferenceType = 'MANUAL' AND i.Quantity * i.UnitCost <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, ' ”ÊÌ… «·Ã—œ' AS Party, 1 AS LineOrder, 1400 AS AccountCode, IIf(k.CountValue > 0, k.CountValue, 0) AS Debit, IIf(k.CountValue < 0, -k.CountValue, 0) AS Credit, '«·„Œ“Ê‰' AS LineText" & vbCrLf
    s = s & "FROM qryStockCountValue AS k" & vbCrLf
    s = s & "WHERE k.CountValue <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'STOCK_COUNT' AS SourceType, k.StockCountID AS SourceID, k.CountNumber AS SourceNumber, k.CountDate AS SourceDate, ' ”ÊÌ… «·Ã—œ' AS Party, 2 AS LineOrder, 5200 AS AccountCode, IIf(k.CountValue < 0, -k.CountValue, 0) AS Debit, IIf(k.CountValue > 0, k.CountValue, 0) AS Credit, '›—Êﬁ«  «·Ã—œ' AS LineText" & vbCrLf
    s = s & "FROM qryStockCountValue AS k" & vbCrLf
    s = s & "WHERE k.CountValue <> 0" & vbCrLf
    SaveQuery "qryJournalStock", "√”ÿ— ﬁÌÊœ Õ—ﬂ«  «·„Œ“Ê‰ «·ÌœÊÌ… Ê ”ÊÌ«  «·Ã—œ", s
End Sub

Private Sub Q_qryJournalOpening()
    Dim s As String
    s = "SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 1 AS LineOrder, 110000 + b.CashBoxID AS AccountCode, b.OpeningBalance AS Debit, CCur(0) AS Credit, b.BoxName AS LineText" & vbCrLf
    s = s & "FROM CashBoxes AS b" & vbCrLf
    s = s & "WHERE b.OpeningBalance <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'BOX_OPENING' AS SourceType, b.CashBoxID AS SourceID, b.BoxName AS SourceNumber, b.OpeningDate AS SourceDate, b.BoxName AS Party, 2 AS LineOrder, 3900 AS AccountCode, CCur(0) AS Debit, b.OpeningBalance AS Credit, '—’Ìœ «›  «ÕÌ' AS LineText" & vbCrLf
    s = s & "FROM CashBoxes AS b" & vbCrLf
    s = s & "WHERE b.OpeningBalance <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 1 AS LineOrder, 1300 AS AccountCode, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Credit, c.CustomerName AS LineText" & vbCrLf
    s = s & "FROM Customers AS c" & vbCrLf
    s = s & "WHERE c.OpeningBalance <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'CUSTOMER_OPENING' AS SourceType, c.CustomerID AS SourceID, c.CustomerName AS SourceNumber, c.CreatedAt AS SourceDate, c.CustomerName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(c.OpeningBalance < 0, -c.OpeningBalance, 0) AS Debit, IIf(c.OpeningBalance > 0, c.OpeningBalance, 0) AS Credit, '—’Ìœ «›  «ÕÌ' AS LineText" & vbCrLf
    s = s & "FROM Customers AS c" & vbCrLf
    s = s & "WHERE c.OpeningBalance <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 1 AS LineOrder, 2100 AS AccountCode, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Credit, s.SupplierName AS LineText" & vbCrLf
    s = s & "FROM Suppliers AS s" & vbCrLf
    s = s & "WHERE s.OpeningBalance <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'SUPPLIER_OPENING' AS SourceType, s.SupplierID AS SourceID, s.SupplierName AS SourceNumber, s.CreatedAt AS SourceDate, s.SupplierName AS Party, 2 AS LineOrder, 3900 AS AccountCode, IIf(s.OpeningBalance > 0, s.OpeningBalance, 0) AS Debit, IIf(s.OpeningBalance < 0, -s.OpeningBalance, 0) AS Credit, '—’Ìœ «›  «ÕÌ' AS LineText" & vbCrLf
    s = s & "FROM Suppliers AS s" & vbCrLf
    s = s & "WHERE s.OpeningBalance <> 0" & vbCrLf
    SaveQuery "qryJournalOpening", "√”ÿ— ﬁÌÊœ «·√—’œ… «·«›  «ÕÌ… ··’‰«œÌﬁ Ê«·⁄„·«¡ Ê«·„Ê—œÌ‰", s
End Sub

Private Sub Q_qryManualEntryLines()
    Dim s As String
    s = "SELECT h.ManualEntryID, h.EntryNumber, h.EntryDate, h.Description, l.LineNumber AS LineNo," & vbCrLf
    s = s & "       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote" & vbCrLf
    s = s & "FROM ManualEntries AS h INNER JOIN ManualEntryLines AS l ON h.ManualEntryID = l.ManualEntryID" & vbCrLf
    SaveQuery "qryManualEntryLines", "√”ÿ— «·ﬁÌÊœ «·ÌœÊÌ… „⁄ —√” ﬂ· ﬁÌœ", s
End Sub

Private Sub Q_qryJournalManual()
    Dim s As String
    s = "SELECT 'MANUAL' AS SourceType, m.ManualEntryID AS SourceID, m.EntryNumber AS SourceNumber, m.EntryDate AS SourceDate, m.Description AS Party, m.LineNo AS LineOrder, m.LineAccount AS AccountCode, m.LineDebit AS Debit, m.LineCredit AS Credit, m.LineNote AS LineText" & vbCrLf
    s = s & "FROM qryManualEntryLines AS m" & vbCrLf
    s = s & "WHERE m.LineDebit + m.LineCredit <> 0" & vbCrLf
    SaveQuery "qryJournalManual", "√”ÿ— «·ﬁÌÊœ «·ÌœÊÌ…", s
End Sub

Private Sub Q_qryYearCloseLines()
    Dim s As String
    s = "SELECT h.YearClosingID, h.ClosingNumber, h.ClosingDate, h.Notes, l.LineNumber AS LineNo," & vbCrLf
    s = s & "       l.AccountCode AS LineAccount, l.Debit AS LineDebit, l.Credit AS LineCredit, l.LineText AS LineNote" & vbCrLf
    s = s & "FROM FiscalYearClosings AS h INNER JOIN FiscalYearClosingLines AS l ON h.YearClosingID = l.YearClosingID" & vbCrLf
    SaveQuery "qryYearCloseLines", "√”ÿ— ﬁÌÊœ ≈ﬁ›«· «·”‰Ê«  „⁄ —√” ﬂ· ≈ﬁ›«·", s
End Sub

Private Sub Q_qryJournalYearClose()
    Dim s As String
    s = "SELECT 'YEAR_CLOSE' AS SourceType, y.YearClosingID AS SourceID, y.ClosingNumber AS SourceNumber, y.ClosingDate AS SourceDate, y.Notes AS Party, y.LineNo AS LineOrder, y.LineAccount AS AccountCode, y.LineDebit AS Debit, y.LineCredit AS Credit, y.LineNote AS LineText" & vbCrLf
    s = s & "FROM qryYearCloseLines AS y" & vbCrLf
    s = s & "WHERE y.LineDebit + y.LineCredit <> 0" & vbCrLf
    SaveQuery "qryJournalYearClose", "√”ÿ— ﬁÌÊœ ≈ﬁ›«· «·”‰Ê« : «·≈Ì—«œ«  Ê«·„’—Ê›«  ≈·Ï «·√—»«Õ «·„Õ Ã“…", s
End Sub

Private Sub Q_qryJournalVatReturn()
    Dim s As String
    s = "SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 1 AS LineOrder, 2200 AS AccountCode, IIf(v.SalesStdVAT > 0, v.SalesStdVAT, 0) AS Debit, IIf(v.SalesStdVAT < 0, -v.SalesStdVAT, 0) AS Credit, '÷—Ì»… «·„Œ—Ã«  ··› —…' AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND v.SalesStdVAT <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 2 AS LineOrder, 1500 AS AccountCode, IIf(v.PurchStdVAT < 0, -v.PurchStdVAT, 0) AS Debit, IIf(v.PurchStdVAT > 0, v.PurchStdVAT, 0) AS Credit, '÷—Ì»… «·„œŒ·«  ··› —…' AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND v.PurchStdVAT <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 3 AS LineOrder, 2200 AS AccountCode, IIf(v.Corrections > 0, v.Corrections, 0) AS Debit, IIf(v.Corrections < 0, -v.Corrections, 0) AS Credit, ' ’ÕÌÕ«  „‰ «·› —«  «·”«»ﬁ…' AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND v.Corrections <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'VAT_RETURN' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.FiledDate AS SourceDate, v.ReturnNumber AS Party, 4 AS LineOrder, 2250 AS AccountCode, IIf((v.SalesStdVAT - v.PurchStdVAT + v.Corrections) < 0, -(v.SalesStdVAT - v.PurchStdVAT + v.Corrections), 0) AS Debit, IIf((v.SalesStdVAT - v.PurchStdVAT + v.Corrections) > 0, (v.SalesStdVAT - v.PurchStdVAT + v.Corrections), 0) AS Credit, '’«›Ì ÷—Ì»… «·› —…' AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND (v.SalesStdVAT - v.PurchStdVAT + v.Corrections) <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'VAT_PAYMENT' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.PaidDate AS SourceDate, v.ReturnNumber AS Party, 1 AS LineOrder, 2250 AS AccountCode, v.PaidAmount AS Debit, CCur(0) AS Credit, '”œ«œ ÷—Ì»… «·ﬁÌ„… «·„÷«›…' AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND v.PaidAmount <> 0" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'VAT_PAYMENT' AS SourceType, v.VatReturnID AS SourceID, v.ReturnNumber AS SourceNumber, v.PaidDate AS SourceDate, v.ReturnNumber AS Party, 2 AS LineOrder, v.PaidAccount AS AccountCode, CCur(0) AS Debit, v.PaidAmount AS Credit, v.FilingRef AS LineText" & vbCrLf
    s = s & "FROM VatReturns AS v" & vbCrLf
    s = s & "WHERE v.Status = 'FILED' AND v.PaidAmount <> 0" & vbCrLf
    SaveQuery "qryJournalVatReturn", "√”ÿ— ﬁÌÊœ «·≈ﬁ—«— «·÷—Ì»Ì «·„⁄ „œ («· ”ÊÌ…) Ê”œ«œÂ", s
End Sub

Private Sub Q_JournalLinesQuery()
    Dim s As String
    s = "SELECT e.EntryID, e.EntryNumber, e.EntryDate, e.SourceType, t.TypeName, e.SourceID, e.SourceNumber," & vbCrLf
    s = s & "       e.Description, l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)" & vbCrLf
    s = s & "      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)" & vbCrLf
    s = s & "     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode" & vbCrLf
    s = s & "WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "ORDER BY e.EntryDate, e.EntryNumber, l.LineNumber" & vbCrLf
    SaveQuery "JournalLinesQuery", "ﬁÌÊœ «·ÌÊ„Ì… Œ·«· › —… »√”ÿ—Â«", s
End Sub

Private Sub Q_qryJournalEntryPrint()
    Dim s As String
    s = "SELECT e.EntryID, e.EntryNumber, e.EntryDate, t.TypeName, e.SourceNumber, e.Description, e.TotalDebit," & vbCrLf
    s = s & "       l.LineNumber, l.AccountCode, a.AccountName, l.LineText, l.Debit, l.Credit" & vbCrLf
    s = s & "FROM ((JournalEntries AS e INNER JOIN JournalSourceTypes AS t ON e.SourceType = t.SourceType)" & vbCrLf
    s = s & "      INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID)" & vbCrLf
    s = s & "     INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode" & vbCrLf
    SaveQuery "qryJournalEntryPrint", "»Ì«‰«  ÿ»«⁄… ﬁÌœ", s
End Sub

Private Sub Q_qryTrialBefore()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) AS DebitBefore, Sum(l.Credit) AS CreditBefore" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate < QDate('PeriodStart')" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryTrialBefore", "„Ã„Ê⁄ «·Õ”«»«  ﬁ»· «·› —…", s
End Sub

Private Sub Q_qryTrialPeriod()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryTrialPeriod", "Õ—ﬂ… «·Õ”«»«  Œ·«· «·› —…", s
End Sub

Private Sub Q_TrialBalanceQuery()
    Dim s As String
    s = "SELECT a.AccountCode, a.AccountName, a.AccountType," & vbCrLf
    s = s & "       CCur(Nz(b.DebitBefore, 0)) - CCur(Nz(b.CreditBefore, 0)) AS OpeningBalance," & vbCrLf
    s = s & "       CCur(Nz(p.SumDebit, 0)) AS PeriodDebit, CCur(Nz(p.SumCredit, 0)) AS PeriodCredit," & vbCrLf
    s = s & "       CCur(Nz(b.DebitBefore, 0)) - CCur(Nz(b.CreditBefore, 0)) + CCur(Nz(p.SumDebit, 0)) - CCur(Nz(p.SumCredit, 0)) AS ClosingBalance" & vbCrLf
    s = s & "FROM (Accounts AS a LEFT JOIN qryTrialBefore AS b ON a.AccountCode = b.AccountCode)" & vbCrLf
    s = s & "     LEFT JOIN qryTrialPeriod AS p ON a.AccountCode = p.AccountCode" & vbCrLf
    s = s & "WHERE b.AccountCode Is Not Null OR p.AccountCode Is Not Null" & vbCrLf
    s = s & "ORDER BY a.AccountCode" & vbCrLf
    SaveQuery "TrialBalanceQuery", "„Ì“«‰ «·„—«Ã⁄…: —’Ìœ √Ê· «·„œ… ÊÕ—ﬂ… «·› —… Ê«·—’Ìœ «·Œ «„Ì («·„œÌ‰ „ÊÃ»)", s
End Sub

Private Sub Q_qryStatementBefore()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(o.Debit), 0)) - CCur(Nz(Sum(o.Credit), 0)) AS SumBefore" & vbCrLf
    s = s & "FROM (JournalLines AS o INNER JOIN JournalEntries AS f ON o.EntryID = f.EntryID)" & vbCrLf
    s = s & "     INNER JOIN Accounts AS b ON o.AccountCode = b.AccountCode" & vbCrLf
    s = s & "WHERE (b.Level1Code = QLong('AccountCode') OR b.Level2Code = QLong('AccountCode') OR b.Level3Code = QLong('AccountCode') OR b.Level4Code = QLong('AccountCode') OR b.Level5Code = QLong('AccountCode')) AND f.EntryDate < QDate('PeriodStart')" & vbCrLf
    SaveQuery "qryStatementBefore", "—’Ìœ «·Õ”«» «·„Œ «— („⁄ Õ”«»« Â «· «»⁄…) ﬁ»· »œ«Ì… «·› —…", s
End Sub

Private Sub Q_AccountStatementQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, s.AccountCode AS StatementAccount, s.AccountName AS StatementName, e.EntryDate AS LineDate," & vbCrLf
    s = s & "       e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName, e.SourceNumber AS DocNo," & vbCrLf
    s = s & "       e.Description AS Details, a.AccountCode AS SubCode, a.AccountName AS SubName, l.Debit AS LineDebit," & vbCrLf
    s = s & "       l.Credit AS LineCredit" & vbCrLf
    s = s & "FROM (((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)" & vbCrLf
    s = s & "      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)" & vbCrLf
    s = s & "     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType), Accounts AS s" & vbCrLf
    s = s & "WHERE s.AccountCode = QLong('AccountCode') AND (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode')) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, s.AccountCode, s.AccountName, QDate('PeriodStart'), '-', 0, '—’Ìœ √Ê· «·„œ…', Null, Null, Null, Null," & vbCrLf
    s = s & "       IIf(x.SumBefore > 0, x.SumBefore, 0), IIf(x.SumBefore < 0, -x.SumBefore, 0)" & vbCrLf
    s = s & "FROM Accounts AS s, qryStatementBefore AS x" & vbCrLf
    s = s & "WHERE s.AccountCode = QLong('AccountCode')" & vbCrLf
    s = s & "ORDER BY SortKey, LineDate, EntryNo" & vbCrLf
    SaveQuery "AccountStatementQuery", "ﬂ‘› Õ”«» ·› —…: —’Ìœ √Ê· «·„œ… À„ ﬂ· ”ÿ— ﬁÌœ («·Õ”«» «·—∆Ì”Ì Ì‘„· Õ”«»« Â «· «»⁄…)", s
End Sub

Private Sub Q_GeneralLedgerQuery()
    Dim s As String
    s = "SELECT 1 AS SortKey, a.TreeKey AS AccountKey, a.AccountCode AS LedgerCode, a.AccountName AS LedgerName," & vbCrLf
    s = s & "       e.EntryDate AS LineDate, e.EntryNumber AS EntryNo, e.EntryID AS EntryRef, k.TypeName AS KindName," & vbCrLf
    s = s & "       e.SourceNumber AS DocNo, e.Description AS Details, l.Debit AS LineDebit, l.Credit AS LineCredit" & vbCrLf
    s = s & "FROM ((JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID)" & vbCrLf
    s = s & "      INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode)" & vbCrLf
    s = s & "     INNER JOIN JournalSourceTypes AS k ON e.SourceType = k.SourceType" & vbCrLf
    s = s & "WHERE (QLong('AccountCode') = 0 OR (a.Level1Code = QLong('AccountCode') OR a.Level2Code = QLong('AccountCode') OR a.Level3Code = QLong('AccountCode') OR a.Level4Code = QLong('AccountCode') OR a.Level5Code = QLong('AccountCode'))) AND e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 0, b.TreeKey, b.AccountCode, b.AccountName, QDate('PeriodStart'), '-', 0, '—’Ìœ √Ê· «·„œ…', Null, Null," & vbCrLf
    s = s & "       IIf(t.OpeningBalance > 0, t.OpeningBalance, 0), IIf(t.OpeningBalance < 0, -t.OpeningBalance, 0)" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS b ON t.AccountCode = b.AccountCode" & vbCrLf
    s = s & "WHERE QLong('AccountCode') = 0 OR (b.Level1Code = QLong('AccountCode') OR b.Level2Code = QLong('AccountCode') OR b.Level3Code = QLong('AccountCode') OR b.Level4Code = QLong('AccountCode') OR b.Level5Code = QLong('AccountCode'))" & vbCrLf
    s = s & "ORDER BY AccountKey, SortKey, LineDate, EntryNo" & vbCrLf
    SaveQuery "GeneralLedgerQuery", "œ› — «·√” «– ·› —…: ·ﬂ· Õ”«» ›—⁄Ì —’Ìœ √Ê· «·„œ… À„ √”ÿ— ﬁÌÊœÂ (0 = ﬂ· «·Õ”«»« )", s
End Sub

Private Sub Q_qryTreeRollup()
    Dim s As String
    s = "SELECT d.Level1Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit," & vbCrLf
    s = s & "       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode" & vbCrLf
    s = s & "WHERE d.Level1Code Is Not Null" & vbCrLf
    s = s & "GROUP BY d.Level1Code" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT d.Level2Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit," & vbCrLf
    s = s & "       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode" & vbCrLf
    s = s & "WHERE d.Level2Code Is Not Null" & vbCrLf
    s = s & "GROUP BY d.Level2Code" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT d.Level3Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit," & vbCrLf
    s = s & "       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode" & vbCrLf
    s = s & "WHERE d.Level3Code Is Not Null" & vbCrLf
    s = s & "GROUP BY d.Level3Code" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT d.Level4Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit," & vbCrLf
    s = s & "       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode" & vbCrLf
    s = s & "WHERE d.Level4Code Is Not Null" & vbCrLf
    s = s & "GROUP BY d.Level4Code" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT d.Level5Code AS TreeCode, Sum(t.OpeningBalance) AS SumOpening, Sum(t.PeriodDebit) AS SumDebit," & vbCrLf
    s = s & "       Sum(t.PeriodCredit) AS SumCredit, Sum(t.ClosingBalance) AS SumClosing" & vbCrLf
    s = s & "FROM TrialBalanceQuery AS t INNER JOIN Accounts AS d ON t.AccountCode = d.AccountCode" & vbCrLf
    s = s & "WHERE d.Level5Code Is Not Null" & vbCrLf
    s = s & "GROUP BY d.Level5Code" & vbCrLf
    SaveQuery "qryTreeRollup", "√—’œ… „Ì“«‰ «·„—«Ã⁄… „Ã„¯⁄… ⁄·Ï ﬂ· „” ÊÏ „‰ ‘Ã—… «·Õ”«»« ", s
End Sub

Private Sub Q_TrialBalanceTreeQuery()
    Dim s As String
    s = "SELECT a.AccountCode, a.AccountName, IIf(a.AccountType = 'ASSET', '√’Ê·', IIf(a.AccountType = 'LIABILITY', 'Œ’Ê„', IIf(a.AccountType = 'EQUITY', 'ÕﬁÊﬁ „·ﬂÌ…', IIf(a.AccountType = 'REVENUE', '≈Ì—«œ« ', '„’—Ê›« ')))) AS TypeName, a.AccountLevel, a.TreeKey, a.IsPosting," & vbCrLf
    s = s & "       r.SumOpening AS OpeningBalance, r.SumDebit AS PeriodDebit, r.SumCredit AS PeriodCredit," & vbCrLf
    s = s & "       r.SumClosing AS ClosingBalance" & vbCrLf
    s = s & "FROM Accounts AS a INNER JOIN qryTreeRollup AS r ON a.AccountCode = r.TreeCode" & vbCrLf
    s = s & "ORDER BY a.TreeKey" & vbCrLf
    SaveQuery "TrialBalanceTreeQuery", "„Ì“«‰ «·„—«Ã⁄… »«·„” ÊÌ« : ﬂ· Õ”«» —∆Ì”Ì »„Ã„Ê⁄ Õ”«»« Â «· «»⁄…", s
End Sub

Private Sub Q_qryIncomeMoves()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate >= QDate('PeriodStart') AND e.EntryDate < QDate('PeriodEnd') AND e.SourceType <> 'YEAR_CLOSE'" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryIncomeMoves", "Õ—ﬂ… «·Õ”«»«  ›Ì «·› —… »œÊ‰ ﬁÌÊœ ≈ﬁ›«· «·”‰…", s
End Sub

Private Sub Q_qryCompareMoves()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) AS SumDebit, Sum(l.Credit) AS SumCredit" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate >= QDate('CompareStart') AND e.EntryDate < QDate('CompareEnd') AND e.SourceType <> 'YEAR_CLOSE'" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryCompareMoves", "Õ—ﬂ… «·Õ”«»«  ›Ì › —… «·„ﬁ«—‰… »œÊ‰ ﬁÌÊœ ≈ﬁ›«· «·”‰…", s
End Sub

Private Sub Q_qryIncomeAccounts()
    Dim s As String
    s = "SELECT a.AccountCode, a.AccountName, a.TreeKey," & vbCrLf
    s = s & "       IIf(a.Level2Code = 41, 1, IIf(a.Level2Code = 51, 2, IIf(a.Level2Code = 52, 3," & vbCrLf
    s = s & "           IIf(a.AccountType = 'REVENUE', 4, 5)))) AS SectionNo," & vbCrLf
    s = s & "       IIf(a.AccountType = 'REVENUE', 1, -1) * (CCur(Nz(c.SumCredit, 0)) - CCur(Nz(c.SumDebit, 0))) AS CurrentAmount," & vbCrLf
    s = s & "       IIf(a.AccountType = 'REVENUE', 1, -1) * (CCur(Nz(p.SumCredit, 0)) - CCur(Nz(p.SumDebit, 0))) AS PriorAmount" & vbCrLf
    s = s & "FROM (Accounts AS a LEFT JOIN qryIncomeMoves AS c ON a.AccountCode = c.AccountCode)" & vbCrLf
    s = s & "     LEFT JOIN qryCompareMoves AS p ON a.AccountCode = p.AccountCode" & vbCrLf
    s = s & "WHERE a.AccountType IN ('REVENUE', 'EXPENSE') AND (c.AccountCode Is Not Null OR p.AccountCode Is Not Null)" & vbCrLf
    SaveQuery "qryIncomeAccounts", "Õ”«»«  ﬁ«∆„… «·œŒ·: ’«›Ì Õ—ﬂ… ﬂ· Õ”«» ≈Ì—«œ«  √Ê „’—Ê›«  ›Ì «·› —… Ê› —… «·„ﬁ«—‰…", s
End Sub

Private Sub Q_IncomeStatementQuery()
    Dim s As String
    s = "SELECT q.SectionNo * 10 + 1 AS Block, q.TreeKey AS AccountKey, 'A' AS RowKind, q.AccountName AS Caption," & vbCrLf
    s = s & "       q.AccountCode AS LineAccount, q.CurrentAmount AS CurrentValue, q.PriorAmount AS PriorValue" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 10, '', 'H', '≈Ì—«œ«  «·‰‘«ÿ', Null, Null, Null" & vbCrLf
    s = s & "FROM Settings AS z WHERE z.SettingID = 1" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 12, '', 'T', '’«›Ì ≈Ì—«œ«  «·‰‘«ÿ', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 20, '', 'H', ' ﬂ·›… «·„»Ì⁄« ', Null, Null, Null" & vbCrLf
    s = s & "FROM Settings AS z WHERE z.SettingID = 1" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 22, '', 'T', '≈Ã„«·Ì  ﬂ·›… «·„»Ì⁄« ', Null, CCur(Nz(Sum(IIf(q.SectionNo = 2, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 2, 1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 30, '', 'H', '«·„’—Ê›«  «· ‘€Ì·Ì… Ê«·≈œ«—Ì…', Null, Null, Null" & vbCrLf
    s = s & "FROM Settings AS z WHERE z.SettingID = 1" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 32, '', 'T', '≈Ã„«·Ì «·„’—Ê›«  «· ‘€Ì·Ì… Ê«·≈œ«—Ì…', Null, CCur(Nz(Sum(IIf(q.SectionNo = 3, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 3, 1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 40, '', 'H', '≈Ì—«œ«  √Œ—Ï', Null, Null, Null" & vbCrLf
    s = s & "FROM Settings AS z WHERE z.SettingID = 1" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 42, '', 'T', '≈Ã„«·Ì «·≈Ì—«œ«  «·√Œ—Ï', Null, CCur(Nz(Sum(IIf(q.SectionNo = 4, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 4, 1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 50, '', 'H', '„’—Ê›«  √Œ—Ï', Null, Null, Null" & vbCrLf
    s = s & "FROM Settings AS z WHERE z.SettingID = 1" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 52, '', 'T', '≈Ã„«·Ì «·„’—Ê›«  «·√Œ—Ï', Null, CCur(Nz(Sum(IIf(q.SectionNo = 5, 1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 5, 1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 25, '', 'R', '„Ã„· «·—»Õ', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 35, '', 'R', '«·—»Õ «· ‘€Ì·Ì', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 3, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 3, -1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 60, '', 'R', '’«›Ì «·—»Õ («·Œ”«—…)', Null, CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 2, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 3, -1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 4, 1 * q.CurrentAmount, 0) + IIf(q.SectionNo = 5, -1 * q.CurrentAmount, 0)), 0)), CCur(Nz(Sum(IIf(q.SectionNo = 1, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 2, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 3, -1 * q.PriorAmount, 0) + IIf(q.SectionNo = 4, 1 * q.PriorAmount, 0) + IIf(q.SectionNo = 5, -1 * q.PriorAmount, 0)), 0))" & vbCrLf
    s = s & "FROM qryIncomeAccounts AS q" & vbCrLf
    s = s & "ORDER BY Block, AccountKey" & vbCrLf
    SaveQuery "IncomeStatementQuery", "ﬁ«∆„… «·œŒ·: «·≈Ì—«œ«  Ê«· ﬂ«·Ì› Ê«·„’—Ê›«  Ê„Ã„· Ê’«›Ì «·—»Õ° „⁄ › —… «·„ﬁ«—‰…", s
End Sub

Private Sub Q_qryBalanceAt()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) - Sum(l.Credit) AS NetAt" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate < QDate('PeriodEnd')" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryBalanceAt", "—’Ìœ ﬂ· Õ”«» ›Ì ‰Â«Ì… «·› —… („œÌ‰ „ÊÃ»)", s
End Sub

Private Sub Q_qryBalanceCompare()
    Dim s As String
    s = "SELECT l.AccountCode, Sum(l.Debit) - Sum(l.Credit) AS NetCompare" & vbCrLf
    s = s & "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID" & vbCrLf
    s = s & "WHERE e.EntryDate < QDate('CompareEnd')" & vbCrLf
    s = s & "GROUP BY l.AccountCode" & vbCrLf
    SaveQuery "qryBalanceCompare", "—’Ìœ ﬂ· Õ”«» ›Ì ‰Â«Ì… › —… «·„ﬁ«—‰… („œÌ‰ „ÊÃ»)", s
End Sub

Private Sub Q_qryBalanceAccounts()
    Dim s As String
    s = "SELECT a.AccountCode, a.AccountName, a.TreeKey, a.Level1Code, a.Level2Code," & vbCrLf
    s = s & "       IIf(a.AccountType = 'ASSET', 1, -1) * CCur(Nz(b.NetAt, 0)) AS CurrentAmount," & vbCrLf
    s = s & "       IIf(a.AccountType = 'ASSET', 1, -1) * CCur(Nz(c.NetCompare, 0)) AS PriorAmount" & vbCrLf
    s = s & "FROM (Accounts AS a LEFT JOIN qryBalanceAt AS b ON a.AccountCode = b.AccountCode)" & vbCrLf
    s = s & "     LEFT JOIN qryBalanceCompare AS c ON a.AccountCode = c.AccountCode" & vbCrLf
    s = s & "WHERE a.AccountType IN ('ASSET', 'LIABILITY', 'EQUITY') AND (CCur(Nz(b.NetAt, 0)) <> 0 OR CCur(Nz(c.NetCompare, 0)) <> 0)" & vbCrLf
    SaveQuery "qryBalanceAccounts", "Õ”«»«  «·„Ì“«‰Ì…: —’Ìœ ﬂ· Õ”«» √’Ê· √Ê Œ’Ê„ √Ê ÕﬁÊﬁ „·ﬂÌ… (»ÿ»Ì⁄ Â „ÊÃ»)", s
End Sub

Private Sub Q_qryProfitAt()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(b.NetAt), 0)) AS NetProfitSum" & vbCrLf
    s = s & "FROM qryBalanceAt AS b INNER JOIN Accounts AS a ON b.AccountCode = a.AccountCode" & vbCrLf
    s = s & "WHERE a.AccountType IN ('REVENUE', 'EXPENSE')" & vbCrLf
    SaveQuery "qryProfitAt", "’«›Ì —»Õ «·› —«  €Ì— «·„ﬁ›·… Õ Ï ‰Â«Ì… «·› —… („œÌ‰ „ÊÃ»)", s
End Sub

Private Sub Q_qryProfitCompare()
    Dim s As String
    s = "SELECT CCur(Nz(Sum(c.NetCompare), 0)) AS NetCompareSum" & vbCrLf
    s = s & "FROM qryBalanceCompare AS c INNER JOIN Accounts AS a ON c.AccountCode = a.AccountCode" & vbCrLf
    s = s & "WHERE a.AccountType IN ('REVENUE', 'EXPENSE')" & vbCrLf
    SaveQuery "qryProfitCompare", "’«›Ì —»Õ «·› —«  €Ì— «·„ﬁ›·… Õ Ï ‰Â«Ì… › —… «·„ﬁ«—‰… („œÌ‰ „ÊÃ»)", s
End Sub

Private Sub Q_qryBalanceItems()
    Dim s As String
    s = "SELECT q.Level1Code AS ClassNo, q.Level2Code AS GroupCode, q.CurrentAmount AS CurrentValue, q.PriorAmount AS PriorValue" & vbCrLf
    s = s & "FROM qryBalanceAccounts AS q" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 3, 32, -x.NetProfitSum, -y.NetCompareSum" & vbCrLf
    s = s & "FROM qryProfitAt AS x, qryProfitCompare AS y" & vbCrLf
    SaveQuery "qryBalanceItems", "»‰Êœ «·„Ì“«‰Ì… »„Ã„Ê⁄« Â«° Ê„⁄Â« ’«›Ì «·—»Õ €Ì— «·„ﬁ›· ›Ì «·√—»«Õ «·„Õ Ã“… (32)", s
End Sub

Private Sub Q_BalanceSheetQuery()
    Dim s As String
    s = "SELECT q.Level1Code AS ClassNo, g.TreeKey AS GroupKey, 1 AS Pos, q.TreeKey AS AccountKey, 'A' AS RowKind," & vbCrLf
    s = s & "       q.AccountName AS Caption, q.AccountCode AS LineAccount, q.CurrentAmount AS CurrentValue," & vbCrLf
    s = s & "       q.PriorAmount AS PriorValue" & vbCrLf
    s = s & "FROM qryBalanceAccounts AS q INNER JOIN Accounts AS g ON q.Level2Code = g.AccountCode" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 3, g.TreeKey, 1, 'Z', 'A', '’«›Ì —»Õ (Œ”«—…) «·› —«  €Ì— «·„ﬁ›·…', Null, -x.NetProfitSum, -y.NetCompareSum" & vbCrLf
    s = s & "FROM Accounts AS g, qryProfitAt AS x, qryProfitCompare AS y" & vbCrLf
    s = s & "WHERE g.AccountCode = 32" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT c.AccountCode, '', 0, '', 'C', c.AccountName, Null, Null, Null" & vbCrLf
    s = s & "FROM Accounts AS c" & vbCrLf
    s = s & "WHERE c.AccountCode IN (1, 2, 3)" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT g.Level1Code, g.TreeKey, 0, '', 'G', g.AccountName, Null, Null, Null" & vbCrLf
    s = s & "FROM Accounts AS g" & vbCrLf
    s = s & "WHERE g.AccountCode IN (SELECT GroupCode FROM qryBalanceItems)" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT g.Level1Code, g.TreeKey, 2, '', 'S', g.AccountName, Null, Sum(i.CurrentValue), Sum(i.PriorValue)" & vbCrLf
    s = s & "FROM Accounts AS g INNER JOIN qryBalanceItems AS i ON g.AccountCode = i.GroupCode" & vbCrLf
    s = s & "GROUP BY g.Level1Code, g.TreeKey, g.AccountName" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT i.ClassNo, '~', 9, '', 'T', IIf(i.ClassNo = 1, '≈Ã„«·Ì «·√’Ê·', IIf(i.ClassNo = 2, '≈Ã„«·Ì «·Œ’Ê„', '≈Ã„«·Ì ÕﬁÊﬁ «·„·ﬂÌ…')), Null, Sum(i.CurrentValue), Sum(i.PriorValue)" & vbCrLf
    s = s & "FROM qryBalanceItems AS i" & vbCrLf
    s = s & "GROUP BY i.ClassNo" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 4, '', 9, '', 'T', '≈Ã„«·Ì «·Œ’Ê„ ÊÕﬁÊﬁ «·„·ﬂÌ…', Null, CCur(Nz(Sum(i.CurrentValue), 0)), CCur(Nz(Sum(i.PriorValue), 0))" & vbCrLf
    s = s & "FROM qryBalanceItems AS i" & vbCrLf
    s = s & "WHERE i.ClassNo IN (2, 3)" & vbCrLf
    s = s & "ORDER BY ClassNo, GroupKey, Pos, AccountKey" & vbCrLf
    SaveQuery "BalanceSheetQuery", "«·„Ì“«‰Ì… «·⁄„Ê„Ì… ›Ì ‰Â«Ì… «·› —…: «·√’Ê· = «·Œ’Ê„ + ÕﬁÊﬁ «·„·ﬂÌ…° „⁄ › —… «·„ﬁ«—‰…", s
End Sub

Private Sub Q_AccountTreeQuery()
    Dim s As String
    s = "SELECT a.AccountCode, a.AccountName, IIf(a.AccountType = 'ASSET', '√’Ê·', IIf(a.AccountType = 'LIABILITY', 'Œ’Ê„', IIf(a.AccountType = 'EQUITY', 'ÕﬁÊﬁ „·ﬂÌ…', IIf(a.AccountType = 'REVENUE', '≈Ì—«œ« ', '„’—Ê›« ')))) AS TypeName, a.AccountLevel, a.TreeKey," & vbCrLf
    s = s & "       a.ParentCode, IIf(a.IsPosting, '›—⁄Ì', '—∆Ì”Ì') AS KindName, a.IsPosting, a.IsActive" & vbCrLf
    s = s & "FROM Accounts AS a" & vbCrLf
    s = s & "ORDER BY a.TreeKey" & vbCrLf
    SaveQuery "AccountTreeQuery", "‘Ã—… «·Õ”«»« : ﬂ· Õ”«» »„” Ê«Â Ê‰Ê⁄Â ÊÂ· Ìﬁ»· «·ﬁÌÊœ", s
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
