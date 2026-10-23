Attribute VB_Name = "modQuotations"
'==============================================================================
' modQuotations  -  Retail Store Management System (customer quotations, docs/75)
'
' Module 3 of docs/69. A quotation is entered as a draft for a customer, issued with a validity date (its
' value and VAT are worked out by the invoice engine of modSales and it is printed), then turned with one
' click into a customer order (modSalesOrders, docs/62) or a sales invoice, or rejected by the customer.
' An issued quotation past its validity date is expired: it is no longer turned into anything; a new
' quotation is made from it ("CopyQuotation").
'
'   CreateQuotation         a draft for a customer, valid until a date
'   IssueQuotation          DRAFT -> ISSUED: the lines and totals are priced; a price other than the customer's
'                           price (QuotationDefaultPrice) needs PRICE_OVERRIDE
'   UnissueQuotation        back to a draft while it was not turned into anything
'   QuotationToOrder        ISSUED -> ORDERED: a draft order in the user's warehouse with its lines
'   InvoiceQuotation        ISSUED -> INVOICED: the sale is posted by PostSaleFromCart (modSales) with
'                           QuotationID, which accepts the quoted prices and marks it inside its transaction
'   RejectQuotation         REJECTED (a draft or an issued quotation)
'   CopyQuotation           a new draft with the same customer and lines
'   DeleteQuotation         a draft only
' Lines are in the smallest unit of the product (like the orders, docs/72 section 4). Prices are entered as the
' selling price is (VAT included or not by Settings.PricesIncludeVAT).
' Screen: frmQuotation (built by modBuildForms, reached from the customer orders screen); print rptQuotation;
' report QUOTATIONS.
'==============================================================================
Option Compare Database
Option Explicit

Public Const QUOTE_VALID_DAYS As Long = 14              ' the default validity of a new quotation

Private Function QtStatus(ByVal QuotationID As Long) As String
    QtStatus = Nz(DbValue("SELECT Status FROM Quotations WHERE QuotationID = " & QuotationID), "")
End Function

Public Function QuotationStateFor(ByVal Status As String, ByVal ValidUntil As Variant, ByVal Today As Date) As String
    ' The status shown to the user: an issued quotation past its validity date is EXPIRED.
    QuotationStateFor = Status
    If Status <> "ISSUED" Or IsNull(ValidUntil) Then Exit Function
    If DateValue(CDate(ValidUntil)) < Today Then QuotationStateFor = "EXPIRED"
End Function

Private Function QtState(ByVal QuotationID As Long) As String
    QtState = QuotationStateFor(QtStatus(QuotationID), _
                                DbValue("SELECT ValidUntil FROM Quotations WHERE QuotationID = " & QuotationID), Date)
End Function

Public Function QuotationStatusName(ByVal State As String) As String
    Select Case State
        Case "DRAFT": QuotationStatusName = "مسودة"
        Case "ISSUED": QuotationStatusName = "صادر"
        Case "EXPIRED": QuotationStatusName = "منتهي الصلاحية"
        Case "ORDERED": QuotationStatusName = "تحوّل إلى طلبية"
        Case "INVOICED": QuotationStatusName = "تحوّل إلى فاتورة"
        Case "REJECTED": QuotationStatusName = "رفضه العميل"
    End Select
End Function

Private Function QtCustomer(ByVal QuotationID As Long) As Long
    QtCustomer = Nz(DbValue("SELECT CustomerID FROM Quotations WHERE QuotationID = " & QuotationID), 0)
End Function

Private Function QtNumber(ByVal QuotationID As Long) As String
    QtNumber = Nz(DbValue("SELECT QuotationNumber FROM Quotations WHERE QuotationID = " & QuotationID), "")
End Function

Public Function QuotationDefaultPrice(ByVal CustomerID As Long, ByVal ProductID As Long) As Currency
    ' The price of a new line: the customer's price list (modPriceLists), else the price list of the customer's
    ' rep (modRepRoutes) or the selling price, less the discount of the customer's list. Same as an order line.
    Dim v As Variant, usual As Currency
    v = RepProductPrice(Nz(SalesRepFor(CustomerID), 0), ProductID)
    If IsNull(v) Then
        usual = NoteDefaultPrice("DELIVERY", ProductID)
    Else
        usual = CCur(v)
    End If
    QuotationDefaultPrice = CustomerPrice(CustomerID, ProductID, Null, usual)
End Function

Public Function QuotationLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant, ByVal UnitPrice As Variant) As String
    If IsNull(ProductID) Then
        QuotationLineProblem = "اختر الصنف."
    ElseIf Nz(DbValue("SELECT IsActive FROM Products WHERE ProductID = " & CLng(ProductID)), 0) = 0 Then
        QuotationLineProblem = "الصنف غير نشط."
    ElseIf Nz(Qty, 0) <= 0 Then
        QuotationLineProblem = "الكمية أكبر من صفر."
    ElseIf Nz(UnitPrice, 0) < 0 Then
        QuotationLineProblem = "السعر لا يكون سالبًا."
    End If
End Function

Public Function IsQuotationPrice(ByVal QuotationID As Long, ByVal ProductID As Long, ByVal UnitID As Variant, _
                                 ByVal UnitPrice As Currency) As Boolean
    ' PostSaleFromCart: the price of a line of the issued quotation being invoiced is not a price change
    ' (it was checked when the quotation was issued).
    If QuotationID = 0 Then Exit Function
    If Not IsNull(UnitID) Then Exit Function                       ' the quotation is in the smallest unit
    IsQuotationPrice = Nz(DbValue("SELECT COUNT(*) FROM QuotationLines WHERE QuotationID = " & QuotationID & _
                                  " AND ProductID = " & ProductID & " AND UnitPrice = " & Str$(UnitPrice)), 0) > 0
End Function

Private Function QuotationConvertProblem(ByVal QuotationID As Long) As String
    ' Before an order or an invoice is made from a quotation.
    Select Case QtState(QuotationID)
        Case "ISSUED"
        Case "EXPIRED"
            QuotationConvertProblem = "انتهت صلاحية العرض. أنشئ منه عرضًا جديدًا («عرض جديد منه») وأصدره."
        Case "DRAFT"
            QuotationConvertProblem = "أصدر العرض أولًا."
        Case "ORDERED", "INVOICED"
            QuotationConvertProblem = "تحوّل هذا العرض من قبل."
        Case Else
            QuotationConvertProblem = "العرض غير موجود أو رفضه العميل."
    End Select
    If Len(QuotationConvertProblem) > 0 Then Exit Function
    If Not NotePartyIsActive("DELIVERY", QtCustomer(QuotationID)) Then
        QuotationConvertProblem = "العميل غير موجود أو غير نشط."
    End If
End Function

'==============================================================================
' The way of a quotation
'==============================================================================
Public Function CreateQuotation(ByVal CustomerID As Variant, ByVal ValidUntil As Variant, ByVal Terms As String, _
                                ByVal Notes As String, ByRef NewQuotationID As Long) As String
    Dim rs As DAO.Recordset
    NewQuotationID = 0
    If Not HasPermission("QUOTATIONS") Then
        CreateQuotation = "ليست لديك صلاحية عروض الأسعار."
    ElseIf Not NotePartyIsActive("DELIVERY", CustomerID) Then
        CreateQuotation = "اختر العميل."
    ElseIf IsNull(ValidUntil) Then
        CreateQuotation = "اكتب تاريخ صلاحية العرض («ساري حتى»)."
    ElseIf Not IsDate(ValidUntil) Then
        CreateQuotation = "«ساري حتى» ليس تاريخًا."
    End If
    If Len(CreateQuotation) = 0 Then
        If DateValue(CDate(ValidUntil)) < Date Then CreateQuotation = "تاريخ الصلاحية مضى."
    End If
    If Len(CreateQuotation) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("Quotations", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    rs.AddNew
    rs!QuotationNumber = NextNumber("QUOTATION")
    rs!QuoteDate = Now
    rs!CustomerID = CLng(CustomerID)
    rs!ValidUntil = DateValue(CDate(ValidUntil))
    rs!Status = "DRAFT"
    rs!SubTotal = 0
    rs!Tax = 0
    rs!TotalAmount = 0
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Terms)) > 0 Then rs!Terms = Left$(Trim$(Terms), 255)
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewQuotationID = rs!QuotationID
    rs.Close
    LogAction "QUOTATION_NEW", "Quotations", CStr(NewQuotationID)
End Function

Public Function IssueQuotation(ByVal QuotationID As Long) As String
    ' The lines are priced by the invoice engine (CalcRun) and the totals kept on the quotation.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, msg As String
    Dim cust As Long, n As Long, i As Long, ids() As Long, vatRate As Currency, rate As Currency
    Dim priceChanged As Boolean, validUntil As Variant
    If Not HasPermission("QUOTATIONS") Then
        IssueQuotation = "ليست لديك صلاحية عروض الأسعار."
        Exit Function
    End If
    If QtStatus(QuotationID) <> "DRAFT" Then
        IssueQuotation = "يُصدر العرض وهو مسودة."
        Exit Function
    End If
    cust = QtCustomer(QuotationID)
    If Not NotePartyIsActive("DELIVERY", cust) Then
        IssueQuotation = "العميل غير موجود أو غير نشط."
        Exit Function
    End If
    validUntil = DbValue("SELECT ValidUntil FROM Quotations WHERE QuotationID = " & QuotationID)
    If DateValue(CDate(validUntil)) < Date Then
        IssueQuotation = "تاريخ صلاحية المسودة مضى. أنشئ منها عرضًا جديدًا بصلاحية جديدة."
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    vatRate = Nz(SettingValue("VATRate"), 0.15)
    CalcReset
    Set rs = db.OpenRecordset("SELECT l.QuotationLineID, l.ProductID, l.Quantity, l.UnitPrice, p.VATCategory " & _
                              "FROM QuotationLines AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                              "WHERE l.QuotationID = " & QuotationID & " ORDER BY l.QuotationLineID", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(msg) = 0 Then msg = QuotationLineProblem(rs!ProductID, rs!Quantity, rs!UnitPrice)
        If rs!UnitPrice <> QuotationDefaultPrice(cust, rs!ProductID) Then priceChanged = True
        If Nz(rs!VATCategory, "S") = "S" Then rate = vatRate Else rate = 0
        CalcAddLine rs!Quantity, rs!UnitPrice, 0, rate                      ' modSales
        ReDim Preserve ids(0 To n)
        ids(n) = rs!QuotationLineID
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If Len(msg) = 0 And n = 0 Then msg = "أضف الأصناف وكمياتها أولًا."
    If Len(msg) = 0 And priceChanged And Not HasPermission("PRICE_OVERRIDE") Then
        msg = "في العرض سعر يختلف عن سعر العميل، ولا تملك صلاحية تعديل سعر البيع."
    End If
    If Len(msg) = 0 Then msg = CalcRun(0, Nz(SettingValue("PricesIncludeVAT"), True))
    If Len(msg) > 0 Then
        IssueQuotation = msg
        Exit Function
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    For i = 0 To n - 1
        db.Execute "UPDATE QuotationLines SET NetAmount = " & Str$(CalcLine(i, "NET")) & ", Tax = " & _
                   Str$(CalcLine(i, "TAX")) & ", LineTotal = " & Str$(CalcLine(i, "TOTAL")) & _
                   " WHERE QuotationLineID = " & ids(i), dbFailOnError + dbSeeChanges
    Next
    db.Execute "UPDATE Quotations SET Status = 'ISSUED', IssuedAt = Now(), SubTotal = " & Str$(CalcTotal("TAXABLE")) & _
               ", Tax = " & Str$(CalcTotal("TAX")) & ", TotalAmount = " & Str$(CalcTotal("TOTAL")) & _
               " WHERE QuotationID = " & QuotationID, dbFailOnError + dbSeeChanges
    ws.CommitTrans
    inTrans = False
    LogAction "QUOTATION_ISSUE", "Quotations", CStr(QuotationID), "Total=" & CalcTotal("TOTAL")
    Exit Function
EH:
    IssueQuotation = "تعذر إصدار العرض: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function UnissueQuotation(ByVal QuotationID As Long) As String
    ' Back to a draft (its lines can be changed again) while it was not turned into an order or an invoice.
    Dim before As Collection
    If Not HasPermission("QUOTATIONS") Then
        UnissueQuotation = "ليست لديك صلاحية عروض الأسعار."
    ElseIf QtStatus(QuotationID) <> "ISSUED" Then
        UnissueQuotation = "يعود للمسودة العرض الصادر الذي لم يتحوّل."
    End If
    If Len(UnissueQuotation) > 0 Then Exit Function
    Set before = AuditSnapshot("Quotations", "QuotationID", QuotationID)
    CurrentDb.Execute "UPDATE Quotations SET Status = 'DRAFT', IssuedAt = Null WHERE QuotationID = " & QuotationID, _
                      dbFailOnError + dbSeeChanges
    AuditEdited "QUOTATION_UNISSUE", "Quotations", "QuotationID", QuotationID, before
End Function

Public Function RejectQuotation(ByVal QuotationID As Long) As String
    Dim before As Collection, st As String
    st = QtStatus(QuotationID)
    If Not HasPermission("QUOTATIONS") Then
        RejectQuotation = "ليست لديك صلاحية عروض الأسعار."
    ElseIf st <> "DRAFT" And st <> "ISSUED" Then
        RejectQuotation = "تحوّل العرض أو رُفض من قبل."
    End If
    If Len(RejectQuotation) > 0 Then Exit Function
    Set before = AuditSnapshot("Quotations", "QuotationID", QuotationID)
    CurrentDb.Execute "UPDATE Quotations SET Status = 'REJECTED' WHERE QuotationID = " & QuotationID, _
                      dbFailOnError + dbSeeChanges
    AuditEdited "QUOTATION_REJECT", "Quotations", "QuotationID", QuotationID, before
End Function

Public Function DeleteQuotation(ByVal QuotationID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("QUOTATIONS") Then
        DeleteQuotation = "ليست لديك صلاحية عروض الأسعار."
        Exit Function
    End If
    If QtStatus(QuotationID) <> "DRAFT" Then
        DeleteQuotation = "تُحذف المسودة فقط. العرض الصادر يُرفض بدل حذفه."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("Quotations", "QuotationID", QuotationID)
    CurrentDb.Execute "DELETE FROM QuotationLines WHERE QuotationID = " & QuotationID, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM Quotations WHERE QuotationID = " & QuotationID, dbFailOnError + dbSeeChanges
    AuditDeleted "QUOTATION_DELETE", "Quotations", QuotationID, auditBefore
End Function

Public Function CopyQuotation(ByVal QuotationID As Long, ByVal ValidUntil As Variant, ByRef NewQuotationID As Long) As String
    ' A new draft for the same customer with the same lines, prices, terms and notes (an expired or rejected
    ' quotation is quoted again this way).
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean, msg As String, q As DAO.Recordset
    Dim cust As Long, quoteTerms As String, quoteNotes As String
    NewQuotationID = 0
    If Len(QtStatus(QuotationID)) = 0 Then
        CopyQuotation = "اختر العرض."
        Exit Function
    End If
    On Error GoTo EH
    Set db = CurrentDb
    Set q = db.OpenRecordset("SELECT CustomerID, Terms, Notes FROM Quotations WHERE QuotationID = " & QuotationID, _
                             dbOpenSnapshot, dbSeeChanges)
    cust = q!CustomerID
    quoteTerms = Nz(q!Terms, "")
    quoteNotes = Nz(q!Notes, "")
    q.Close                                  ' closed before BeginTrans (error 3246 inside it)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateQuotation(cust, ValidUntil, quoteTerms, quoteNotes, NewQuotationID)
    If Len(msg) = 0 Then
        db.Execute "INSERT INTO QuotationLines (QuotationID, ProductID, Quantity, UnitPrice, NetAmount, Tax, LineTotal) " & _
                   "SELECT " & NewQuotationID & ", ProductID, Quantity, UnitPrice, 0, 0, 0 FROM QuotationLines " & _
                   "WHERE QuotationID = " & QuotationID & " ORDER BY QuotationLineID", dbFailOnError + dbSeeChanges
    End If
    If Len(msg) > 0 Then
        ws.Rollback
        NewQuotationID = 0
        CopyQuotation = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "QUOTATION_COPY", "Quotations", CStr(QuotationID), "New=" & NewQuotationID
    Exit Function
EH:
    CopyQuotation = "تعذر إنشاء العرض: " & Err.Description & " (" & Err.Number & ")"
    NewQuotationID = 0
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' Turning a quotation into an order or an invoice
'==============================================================================
Public Function QuotationToOrder(ByVal QuotationID As Long, ByRef NewOrderID As Long) As String
    ' A draft customer order (modSalesOrders) in the user's warehouse with the lines of the quotation at its
    ' prices; the user approves it there (approving reserves the stock).
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, ln As DAO.Recordset, inTrans As Boolean
    Dim msg As String
    NewOrderID = 0
    If Not HasPermission("QUOTATIONS") Then
        QuotationToOrder = "ليست لديك صلاحية عروض الأسعار."
        Exit Function
    End If
    msg = QuotationConvertProblem(QuotationID)
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity, UnitPrice FROM QuotationLines WHERE QuotationID = " & _
                                         QuotationID, dbOpenSnapshot)
        Do Until rs.EOF
            If Len(msg) = 0 Then msg = SalesOrderLineProblem(rs!ProductID, rs!Quantity, rs!UnitPrice)   ' modSalesOrders
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) > 0 Then msg = msg & vbCrLf & "حوّل العرض إلى فاتورة بدل الطلبية."
    End If
    If Len(msg) > 0 Then
        QuotationToOrder = msg
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateSalesOrder(QtCustomer(QuotationID), CurrentWarehouseID(), Null, "عرض السعر " & QtNumber(QuotationID), _
                           NewOrderID)
    If Len(msg) = 0 Then
        Set rs = db.OpenRecordset("SELECT ProductID, Quantity, UnitPrice FROM QuotationLines WHERE QuotationID = " & _
                                  QuotationID & " ORDER BY QuotationLineID", dbOpenSnapshot, dbSeeChanges)
        Set ln = db.OpenRecordset("SalesOrderLines", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
        Do Until rs.EOF
            ln.AddNew
            ln!SalesOrderID = NewOrderID
            ln!ProductID = rs!ProductID
            ln!Quantity = rs!Quantity
            ln!UnitPrice = rs!UnitPrice
            ln.Update
            rs.MoveNext
        Loop
        ln.Close
        rs.Close
        db.Execute "UPDATE Quotations SET Status = 'ORDERED', SalesOrderID = " & NewOrderID & _
                   " WHERE QuotationID = " & QuotationID, dbFailOnError + dbSeeChanges
    End If
    If Len(msg) > 0 Then
        ws.Rollback
        NewOrderID = 0
        QuotationToOrder = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "QUOTATION_ORDER", "Quotations", CStr(QuotationID), "Order=" & NewOrderID
    Exit Function
EH:
    QuotationToOrder = "تعذر إنشاء الطلبية: " & Err.Description & " (" & Err.Number & ")"
    NewOrderID = 0
    If inTrans Then ws.Rollback
End Function

Public Sub QuotationOrderDeleted(ByVal SalesOrderID As Long)
    ' DeleteSalesOrder (a draft order): the quotation it was made from is issued again.
    CurrentDb.Execute "UPDATE Quotations SET Status = 'ISSUED', SalesOrderID = Null WHERE SalesOrderID = " & SalesOrderID, _
                      dbFailOnError + dbSeeChanges
End Sub

Public Function InvoiceQuotation(ByVal QuotationID As Long, ByVal PaymentType As String, ByRef NewInvoiceID As Long) As String
    ' The sales invoice of the quotation at its prices, from the user's warehouse, paid in cash or on credit.
    Dim db As DAO.Database, rs As DAO.Recordset, cart As DAO.Recordset, method As Variant, msg As String
    NewInvoiceID = 0
    If Not HasPermission("QUOTATIONS") Or Not HasPermission("SALES_POS") Then
        InvoiceQuotation = "تحتاج صلاحيتي عروض الأسعار ونقطة البيع."
        Exit Function
    End If
    msg = QuotationConvertProblem(QuotationID)
    If Len(msg) = 0 And PaymentType <> "CASH" And PaymentType <> "CREDIT" Then msg = "اختر دفع الفاتورة: نقدي أو آجل."
    If Len(msg) > 0 Then
        InvoiceQuotation = msg
        Exit Function
    End If
    EnsureLocalTables                                               ' modPOS
    If Nz(DbValue("SELECT COUNT(*) FROM tmpPOSLines"), 0) > 0 Then
        InvoiceQuotation = "في شاشة البيع فاتورة لم تُحفظ بعد. احفظها أو امسحها أولًا."
        Exit Function
    End If
    If PaymentType = "CASH" Then method = CASH_METHOD_ID Else method = Null

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT l.ProductID, l.Quantity, l.UnitPrice, p.ProductCode, p.ProductName, p.CurrentQuantity " & _
                              "FROM QuotationLines AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                              "WHERE l.QuotationID = " & QuotationID & " ORDER BY l.QuotationLineID", dbOpenSnapshot, _
                              dbSeeChanges)
    Set cart = db.OpenRecordset("tmpPOSLines", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    Do Until rs.EOF                                                 ' one cart line per quotation line, in order
        cart.AddNew
        cart!ProductID = rs!ProductID
        cart!ProductCode = rs!ProductCode
        cart!ProductName = rs!ProductName
        cart!Quantity = rs!Quantity
        cart!UnitPrice = rs!UnitPrice
        cart!LineDiscount = 0                                       ' local tables have no defaults
        cart!Available = Nz(rs!CurrentQuantity, 0)
        cart!UnitFactor = 1
        cart!ListPrice = rs!UnitPrice
        cart.Update
        rs.MoveNext
    Loop
    cart.Close
    rs.Close
    msg = PostSaleFromCart(QtCustomer(QuotationID), PaymentType, method, 0, Null, "عرض السعر " & QtNumber(QuotationID), _
                           NewInvoiceID, QuotationID:=QuotationID)                  ' modSales
    db.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges              ' the cart never keeps the lines
    InvoiceQuotation = msg
    If Len(msg) = 0 Then LogAction "QUOTATION_INVOICE", "Quotations", CStr(QuotationID), "Invoice=" & NewInvoiceID
End Function

Public Sub MarkQuotationInvoiced(ByVal db As DAO.Database, ByVal QuotationID As Long, ByVal SalesInvoiceID As Long)
    ' Inside the transaction of PostSaleFromCart.
    db.Execute "UPDATE Quotations SET Status = 'INVOICED', SalesInvoiceID = " & SalesInvoiceID & _
               " WHERE QuotationID = " & QuotationID, dbFailOnError + dbSeeChanges
End Sub

'==============================================================================
' frmQuotation (OpenArgs = a quotation to show) and its lines frmQuotationLines
'==============================================================================
Public Sub QuotationLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboCustomer.RowSource = NotePartyRows("DELIVERY")              ' the active customers (modStockNotes)
    frm!txtValidUntil.Value = Date + QUOTE_VALID_DAYS
    frm!cboPayType.RowSource = Tr("CASH;دفع نقدي;CREDIT;دفع آجل")
    frm!cboPayType.Value = "CASH"
    QuotationListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(QuotationID) FROM Quotations")
    frm!cboQuote.Value = id
    QuotationShow frm
End Sub

Private Sub QuotationListRefresh(ByVal frm As Access.Form)
    frm!cboQuote.RowSource = Tr("SELECT q.QuotationID, q.QuotationNumber, Format(q.QuoteDate, 'yyyy/mm/dd'), c.CustomerName, " & _
        "Format(q.ValidUntil, 'yyyy/mm/dd'), IIf(q.Status = 'DRAFT', 'مسودة', IIf(q.Status = 'ISSUED', " & _
        "IIf(q.ValidUntil < Date(), 'منتهي', 'صادر'), IIf(q.Status = 'ORDERED', 'طلبية', IIf(q.Status = 'INVOICED', " & _
        "'فاتورة', 'مرفوض')))) FROM Quotations AS q INNER JOIN [@Customers] AS c ON q.CustomerID = c.CustomerID " & _
        "ORDER BY q.QuotationID DESC")
    frm!cboQuote.Requery
End Sub

Public Sub QuotationShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, state As String, rs As DAO.Recordset, info As String, v As Currency
    id = Nz(frm!cboQuote.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM QuotationLines WHERE QuotationID = " & id & " ORDER BY QuotationLineID"
    st = QtStatus(id)
    state = QtState(id)
    If Len(st) = 0 Then
        info = "لا يوجد عرض سعر. اختر العميل وتاريخ الصلاحية ثم «عرض جديد»."
    Else
        Set rs = CurrentDb.OpenRecordset(Tr("SELECT q.QuotationNumber, q.ValidUntil, q.TotalAmount, q.SalesOrderID, " & _
                                            "q.SalesInvoiceID, c.CustomerName FROM Quotations AS q INNER JOIN [@Customers] AS c " & _
                                            "ON q.CustomerID = c.CustomerID WHERE q.QuotationID = ") & id, dbOpenSnapshot)
        info = rs!QuotationNumber & "    " & rs!CustomerName & "    ساري حتى " & GDate(rs!ValidUntil) & "    " & _
               QuotationStatusName(state)
        If st = "ORDERED" Then
            info = info & " " & Nz(DbValue("SELECT OrderNumber FROM SalesOrders WHERE SalesOrderID = " & _
                                           Nz(rs!SalesOrderID, 0)), "")
        ElseIf st = "INVOICED" Then
            info = info & " " & Nz(DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & _
                                           Nz(rs!SalesInvoiceID, 0)), "")
        End If
        If st = "DRAFT" Then
            v = Nz(DbValue("SELECT Sum(Quantity * UnitPrice) FROM QuotationLines WHERE QuotationID = " & id), 0)
            info = info & "    القيمة: " & Format$(v, "#,##0.00")
        Else
            info = info & "    الإجمالي شامل الضريبة: " & Format$(Nz(rs!TotalAmount, 0), "#,##0.00")
        End If
        rs.Close
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    SafeFocus frm!cboQuote                  ' a focused control cannot be disabled
    frm!btnIssue.Enabled = (st = "DRAFT")
    frm!btnPrint.Enabled = (Len(st) > 0 And st <> "DRAFT")       ' the totals are worked out at issue
    frm!btnToOrder.Enabled = (state = "ISSUED")
    frm!btnToInvoice.Enabled = (state = "ISSUED")
    frm!btnUnissue.Enabled = (st = "ISSUED")
    frm!btnReject.Enabled = (st = "DRAFT" Or st = "ISSUED")
    frm!btnDelete.Enabled = (st = "DRAFT")
    frm!btnCopy.Enabled = (Len(st) > 0)
    frm!btnOpenDoc.Enabled = (st = "ORDERED" Or st = "INVOICED")
End Sub

Public Sub QuotationNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateQuotation(frm!cboCustomer.Value, frm!txtValidUntil.Value, Nz(frm!txtTerms.Value, ""), _
                          Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    QuotationListRefresh frm
    frm!cboQuote.Value = id
    QuotationShow frm
End Sub

Public Sub QuotationAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String, newID As Long
    id = Nz(frm!cboQuote.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", IIf(Action = "COPY", "ADD", "EDIT"))) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "ISSUE"
            msg = IssueQuotation(id)
        Case "ORDER"
            If Not AskYesNo("تحويل العرض إلى طلبية عميل (مسودة في مخزنك)؟") Then Exit Sub
            msg = QuotationToOrder(id, newID)
        Case "INVOICE"
            If Not AskYesNo("إنشاء فاتورة البيع من العرض؟ (" & frm!cboPayType.Column(1) & ")") Then Exit Sub
            msg = InvoiceQuotation(id, Nz(frm!cboPayType.Value, ""), newID)
        Case "UNISSUE"
            If Not AskYesNo("إعادة العرض لمسودة؟") Then Exit Sub
            msg = UnissueQuotation(id)
        Case "REJECT"
            If Not AskYesNo("تسجيل أن العميل رفض العرض؟") Then Exit Sub
            msg = RejectQuotation(id)
        Case "DELETE"
            If Not AskYesNo("حذف مسودة العرض؟") Then Exit Sub
            msg = DeleteQuotation(id)
        Case "COPY"
            msg = CopyQuotation(id, frm!txtValidUntil.Value, newID)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    QuotationListRefresh frm
    If Action = "DELETE" Then frm!cboQuote.Value = DbValue("SELECT Max(QuotationID) FROM Quotations")
    If Action = "COPY" Then frm!cboQuote.Value = newID
    QuotationShow frm
    Select Case Action
        Case "ISSUE"
            If AskYesNo("صدر العرض. طباعته الآن؟") Then PrintQuotation id               ' modReports
        Case "ORDER"
            OpenScreen "frmSalesOrder", 0, newID                    ' check the order, then approve it
        Case "INVOICE"
            ShowInfo "صدرت فاتورة البيع " & Nz(DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & _
                                                        newID), "") & "."
    End Select
End Sub

Public Sub QuotationPrint(ByVal frm As Access.Form)
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    PrintQuotation frm!cboQuote.Value                                ' modReports
End Sub

Public Sub QuotationOpenDoc(ByVal frm As Access.Form)
    ' The order the quotation was turned into, or the view of its invoice.
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT SalesOrderID, SalesInvoiceID FROM Quotations WHERE QuotationID = " & _
                                     Nz(frm!cboQuote.Value, 0), dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    If Not IsNull(rs!SalesOrderID) Then
        OpenScreen "frmSalesOrder", 0, rs!SalesOrderID
    ElseIf Not IsNull(rs!SalesInvoiceID) Then
        OpenScreen "frmSalesInvoice", 0, rs!SalesInvoiceID
    End If
    rs.Close
End Sub

Public Sub QuotationLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String
    id = Nz(frm.Parent!cboQuote.Value, 0)
    If QtStatus(id) <> "DRAFT" Then
        msg = "تُعدَّل أسطر المسودة فقط. أعد العرض لمسودة أولًا."
    Else
        msg = QuotationLineProblem(frm!ProductID.Value, frm!Quantity.Value, frm!UnitPrice.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!QuotationID.Value = id
    If Nz(frm!UnitPrice.Value, 0) = 0 Then
        frm!UnitPrice.Value = QuotationDefaultPrice(QtCustomer(id), CLng(frm!ProductID.Value))
    End If
    AuditFormBefore frm, "QuotationLines", "QuotationLineID"
End Sub

Public Sub QuotationLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    QuotationShow frm.Parent
End Sub

Public Sub QuotationRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!QuotationLineID.Value
    If IsNull(id) Then Exit Sub
    If QtStatus(Nz(frm.Parent!cboQuote.Value, 0)) <> "DRAFT" Then
        ShowWarning "تُعدَّل أسطر المسودة فقط. أعد العرض لمسودة أولًا."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("QuotationLines", "QuotationLineID", id)
    CurrentDb.Execute "DELETE FROM QuotationLines WHERE QuotationLineID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "QUOTATION_LINE_DELETE", "QuotationLines", id, auditBefore
    frm.Requery
    QuotationShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Private Sub CheckQt(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                    ByRef Report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "  [OK] " & Title
    Else
        failed = failed + 1
        Report = Report & "- " & Title & vbCrLf
        Debug.Print "  [X]  " & Title
    End If
End Sub

Public Function TestQuotations() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, refNo As String, msg As String, cust As Long
    Dim passed As Long, failed As Long, Report As String, qid As Long, orderID As Long, invID As Long
    Dim q2 As Long, total As Currency, expected As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestQuotations  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "UPDATE Settings SET AllowNegativeStock = False", dbFailOnError + dbSeeChanges
    g_AutoAnswer = False
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-QT-P', 'TEST صنف عروض الأسعار', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-QT-P'")
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive <> 0 AND PriceListID Is Null"), 0)
    CheckQt cust > 0, "يوجد عميل نشط بلا قائمة أسعار", passed, failed, Report
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckQt Len(msg) = 0, "إضافة 10 بتكلفة 6 " & msg, passed, failed, Report

    ' the expiry rule
    CheckQt QuotationStateFor("ISSUED", Date - 1, Date) = "EXPIRED" And QuotationStateFor("ISSUED", Date, Date) = "ISSUED" _
            And QuotationStateFor("ORDERED", Date - 1, Date) = "ORDERED", "العرض الصادر بعد صلاحيته منتهٍ", _
            passed, failed, Report

    ' a draft of 3 at the selling price, issued: priced like an invoice
    CheckQt Len(CreateQuotation(cust, Date - 1, "", "", qid)) > 0, "لا عرض بصلاحية مضت", passed, failed, Report
    msg = CreateQuotation(cust, Date + 14, "TEST الدفع نقدًا", "TEST", qid)
    CheckQt Len(msg) = 0 And QtStatus(qid) = "DRAFT" And Left$(QtNumber(qid), 3) = "QT-", "مسودة عرض QT- " & msg, _
            passed, failed, Report
    CheckQt Len(IssueQuotation(qid)) > 0, "لا يُصدر عرض بلا أسطر", passed, failed, Report
    CheckQt QuotationDefaultPrice(cust, pid) = 20, "السعر المقترح للسطر = سعر البيع", passed, failed, Report
    CurrentDb.Execute "INSERT INTO QuotationLines (QuotationID, ProductID, Quantity, UnitPrice) VALUES (" & qid & ", " & _
                      pid & ", 3, 20)", dbFailOnError + dbSeeChanges
    msg = IssueQuotation(qid)
    If Nz(SettingValue("PricesIncludeVAT"), True) Then expected = 60 Else expected = RoundMoney(60 * (1 + Nz(SettingValue("VATRate"), 0)))
    total = Nz(DbValue("SELECT TotalAmount FROM Quotations WHERE QuotationID = " & qid), 0)
    CheckQt Len(msg) = 0 And QtStatus(qid) = "ISSUED" And total = expected, "الإصدار: الإجمالي " & total & " " & msg, _
            passed, failed, Report
    CheckQt Nz(DbValue("SELECT SubTotal + Tax FROM Quotations WHERE QuotationID = " & qid), 0) = total And _
            Nz(DbValue("SELECT Sum(LineTotal) FROM QuotationLines WHERE QuotationID = " & qid), 0) = total, _
            "الصافي + الضريبة = الإجمالي = مجموع الأسطر", passed, failed, Report
    CheckQt Nz(DbValue("SELECT COUNT(*) FROM qryQuotationPrint WHERE QuotationID = " & qid), 0) = 1, _
            "بيانات طباعة العرض", passed, failed, Report

    ' expired: nothing is made from it
    CurrentDb.Execute "UPDATE Quotations SET ValidUntil = Date() - 1 WHERE QuotationID = " & qid, dbFailOnError + dbSeeChanges
    CheckQt QtState(qid) = "EXPIRED" And Len(QuotationToOrder(qid, orderID)) > 0 And _
            Len(InvoiceQuotation(qid, "CASH", invID)) > 0, "العرض المنتهي لا يتحوّل", passed, failed, Report
    CheckQt Nz(DbValue("SELECT StatusName FROM QuotationListQuery WHERE QuotationID = " & qid), "") = "منتهي", _
            "تقرير العروض: منتهي", passed, failed, Report
    CurrentDb.Execute "UPDATE Quotations SET ValidUntil = Date() + 14 WHERE QuotationID = " & qid, dbFailOnError + dbSeeChanges

    ' to an order, then the draft order deleted: the quotation is issued again
    msg = QuotationToOrder(qid, orderID)
    CheckQt Len(msg) = 0 And QtStatus(qid) = "ORDERED" And _
            Nz(DbValue("SELECT Status FROM SalesOrders WHERE SalesOrderID = " & orderID), "") = "DRAFT" And _
            Nz(DbValue("SELECT Sum(Quantity * UnitPrice) FROM SalesOrderLines WHERE SalesOrderID = " & orderID), 0) = 60 And _
            Nz(DbValue("SELECT SalesOrderID FROM Quotations WHERE QuotationID = " & qid), 0) = orderID, _
            "إلى طلبية: مسودة طلبية بأسطر العرض وأسعاره " & msg, passed, failed, Report
    CheckQt Len(QuotationToOrder(qid, q2)) > 0 And Len(InvoiceQuotation(qid, "CASH", invID)) > 0, _
            "العرض لا يتحوّل مرتين", passed, failed, Report
    msg = DeleteSalesOrder(orderID)
    CheckQt Len(msg) = 0 And QtStatus(qid) = "ISSUED" And _
            IsNull(DbValue("SELECT SalesOrderID FROM Quotations WHERE QuotationID = " & qid)), _
            "حذف مسودة الطلبية يعيد العرض صادرًا " & msg, passed, failed, Report

    ' a quoted price other than the customer's price, then the invoice at it
    msg = UnissueQuotation(qid)
    CurrentDb.Execute "UPDATE QuotationLines SET UnitPrice = 18 WHERE QuotationID = " & qid, dbFailOnError + dbSeeChanges
    msg = msg & IssueQuotation(qid)
    CheckQt Len(msg) = 0 And QtStatus(qid) = "ISSUED", "سعر متفاوض عليه 18 (صلاحية تعديل السعر) " & msg, _
            passed, failed, Report
    CheckQt IsQuotationPrice(qid, pid, Null, 18) And Not IsQuotationPrice(qid, pid, Null, 17) And _
            Not IsQuotationPrice(0, pid, Null, 18), "سعر العرض مقبول في فاتورته فقط", passed, failed, Report
    total = Nz(DbValue("SELECT TotalAmount FROM Quotations WHERE QuotationID = " & qid), 0)
    msg = InvoiceQuotation(qid, "CASH", invID)
    CheckQt Len(msg) = 0 And invID > 0 And QtStatus(qid) = "INVOICED" And _
            Nz(DbValue("SELECT SalesInvoiceID FROM Quotations WHERE QuotationID = " & qid), 0) = invID, _
            "إلى فاتورة: تحوّل إلى فاتورة " & msg, passed, failed, Report
    CheckQt Nz(DbValue("SELECT TotalAmount FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) = total And _
            WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 7 And Nz(DbValue("SELECT COUNT(*) FROM tmpPOSLines"), 0) = 0, _
            "الفاتورة بقيمة العرض، والرصيد 7، وسلة البيع فارغة", passed, failed, Report
    CheckQt Len(UnissueQuotation(qid)) > 0 And Len(RejectQuotation(qid)) > 0 And Len(DeleteQuotation(qid)) > 0, _
            "العرض المحوّل لا يعود ولا يُرفض ولا يُحذف", passed, failed, Report

    ' a new quotation made from it, rejected; a draft deleted
    msg = CopyQuotation(qid, Date + 7, q2)
    CheckQt Len(msg) = 0 And QtStatus(q2) = "DRAFT" And QtCustomer(q2) = cust And _
            Nz(DbValue("SELECT Sum(Quantity * UnitPrice) FROM QuotationLines WHERE QuotationID = " & q2), 0) = 54, _
            "عرض جديد منه بالأسطر نفسها " & msg, passed, failed, Report
    msg = RejectQuotation(q2)
    CheckQt Len(msg) = 0 And QtStatus(q2) = "REJECTED" And Len(DeleteQuotation(q2)) > 0, "رفض العميل " & msg, _
            passed, failed, Report
    msg = CreateQuotation(cust, Date + 3, "", "TEST", q2)
    CheckQt Len(DeleteQuotation(q2)) = 0 And QtStatus(q2) = "", "حذف المسودة", passed, failed, Report
    CheckQt Len(QuotationLineProblem(pid, 0, 5)) > 0 And Len(QuotationLineProblem(Null, 1, 5)) > 0 And _
            Len(QuotationLineProblem(pid, 1, -1)) > 0 And Len(QuotationLineProblem(pid, 1, 5)) = 0, _
            "فحص السطر", passed, failed, Report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckQt False, "خطأ: " & Err.Description, passed, failed, Report
    If inTrans Then ws.Rollback
Done:
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات عروض الأسعار ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestQuotations"
        TestQuotations = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, _
                "TestQuotations"
    End If
End Function
