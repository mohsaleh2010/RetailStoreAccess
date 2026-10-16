Attribute VB_Name = "modRepRoutes"
'==============================================================================
' modRepRoutes  -  Retail Store Management System (routes, visits and rep limits, docs/64, plan 4C)
'
' The route of a rep (RouteStops): the customers he visits on each weekday (VBA Weekday, Sunday = 1),
' in an order. The visits (RepVisits): one per rep, customer and day:
'   CreateVisitPlan     the plan of a day = the active route stops of its weekday (IsPlanned), once
'   AddRepVisit         a visit outside the route
'   RecordRepVisit      the result: ORDER / SALE / PAYMENT / NO_SALE / CLOSED; left empty, the result is
'                       read from the documents of the customer that day (SuggestedVisitResult:
'                       a sale, else an order, else a payment)
'   DeleteRepVisit      a visit not recorded yet (audited)
' The limits of a rep apply when the rep sells (SessionRepID: the rep day the office works for, modVanSales,
' else the active rep linked to the current user):
'   RepPrices           his price list: the selling price of a product in the POS (AddLine), accepted
'                       without PRICE_OVERRIDE (modSales / modPOS), and the default price of an order line
'   MaxDiscountPercent  replaces the discount limit of the user (RepDiscountLimit, 0 = the user's)
'   CreditLimit         the balances of the rep's customers (Customers.SalesRepID) plus the credit of a new
'                       sale may not pass it (RepCreditProblem, for the rep of the invoice; 0 = no limit)
' Reports: RouteExecutionQuery (ROUTE_EXECUTION), RepVisitListQuery (REP_VISITS), RepVanStockQuery
' (REP_VAN_STOCK). Screens: frmRepVisits, frmRouteStops, frmRepPrices (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

'==============================================================================
' The rep who sells, and his limits
'==============================================================================
Public Function SessionRepID() As Long
    If ActingRepDayID() > 0 Then
        SessionRepID = ActingSalesRepID()
        Exit Function
    End If
    SessionRepID = Nz(DbValue("SELECT Min(SalesRepID) FROM SalesReps WHERE IsActive = True AND EmployeeID = " & _
                              CurrentUserID()), 0)
End Function

Public Function RepProductPrice(ByVal SalesRepID As Long, ByVal ProductID As Long) As Variant
    RepProductPrice = Null
    If SalesRepID = 0 Then Exit Function
    RepProductPrice = DbValue("SELECT UnitPrice FROM RepPrices WHERE SalesRepID = " & SalesRepID & _
                              " AND ProductID = " & ProductID)
End Function

Public Function SessionPrice(ByVal ProductID As Long, ByVal SellingPrice As Currency) As Currency
    ' The price of a product in a new sale: the price list of the rep who sells, else the selling price.
    Dim v As Variant
    v = RepProductPrice(SessionRepID(), ProductID)
    If IsNull(v) Then
        SessionPrice = SellingPrice
    Else
        SessionPrice = CCur(v)
    End If
End Function

Public Function IsSessionRepPrice(ByVal ProductID As Long, ByVal UnitPrice As Currency) As Boolean
    Dim v As Variant
    v = RepProductPrice(SessionRepID(), ProductID)
    If Not IsNull(v) Then IsSessionRepPrice = (CCur(v) = UnitPrice)
End Function

Public Function RepDiscountLimit(ByVal UserLimit As Double) As Double
    Dim rep As Long, lim As Double
    RepDiscountLimit = UserLimit
    rep = SessionRepID()
    If rep = 0 Then Exit Function
    lim = Nz(DbValue("SELECT MaxDiscountPercent FROM SalesReps WHERE SalesRepID = " & rep), 0)
    If lim > 0 Then RepDiscountLimit = lim
End Function

Public Function RepCustomersBalance(ByVal SalesRepID As Long) As Currency
    RepCustomersBalance = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Customers WHERE CurrentBalance > 0 " & _
                                     "AND SalesRepID = " & SalesRepID), 0)
End Function

Public Function RepCreditProblem(ByVal SalesRepID As Variant, ByVal Remaining As Currency) As String
    ' The credit of a new sale against the limit of the rep of the invoice.
    Dim limitAmt As Currency, bal As Currency
    If IsNull(SalesRepID) Or Remaining <= 0 Then Exit Function
    limitAmt = Nz(DbValue("SELECT CreditLimit FROM SalesReps WHERE SalesRepID = " & CLng(SalesRepID)), 0)
    If limitAmt <= 0 Then Exit Function
    bal = RepCustomersBalance(CLng(SalesRepID))
    If bal + Remaining > limitAmt Then
        RepCreditProblem = "تجاوز حد ائتمان المندوب «" & _
                           Nz(DbValue(Tr("SELECT s.RepName FROM [@SalesReps] AS s WHERE s.SalesRepID = " & CLng(SalesRepID))), "") & _
                           "»: أرصدة عملائه " & Format$(bal, "#,##0.00") & " + المتبقي " & Format$(Remaining, "#,##0.00") & _
                           " أكبر من الحد " & Format$(limitAmt, "#,##0.00") & "."
    End If
End Function

'==============================================================================
' Visits
'==============================================================================
Private Function CanRoutes() As String
    If Not HasPermission("REP_ROUTES") Then CanRoutes = "ليست لديك صلاحية خط السير وزيارات المندوبين."
End Function

Private Function VisitField(ByVal VisitID As Long, ByVal FieldName As String) As Variant
    VisitField = DbValue("SELECT " & FieldName & " FROM RepVisits WHERE RepVisitID = " & VisitID)
End Function

Public Function VisitWeekdayOf(ByVal VisitDate As Date) As Integer
    VisitWeekdayOf = Weekday(VisitDate, vbSunday)
End Function

Public Function RouteWeekdayName(ByVal VisitWeekday As Integer) As String
    ' 1 = Sunday ... 7 = Saturday, as RouteStops.VisitWeekday (fixed names, not the Windows locale)
    Select Case VisitWeekday
        Case 1
            RouteWeekdayName = Tr("الأحد")
        Case 2
            RouteWeekdayName = Tr("الاثنين")
        Case 3
            RouteWeekdayName = Tr("الثلاثاء")
        Case 4
            RouteWeekdayName = Tr("الأربعاء")
        Case 5
            RouteWeekdayName = Tr("الخميس")
        Case 6
            RouteWeekdayName = Tr("الجمعة")
        Case Else
            RouteWeekdayName = Tr("السبت")
    End Select
End Function

Private Function RepActive(ByVal SalesRepID As Variant) As Boolean
    If IsNull(SalesRepID) Then Exit Function
    RepActive = Nz(DbValue("SELECT IsActive FROM SalesReps WHERE SalesRepID = " & CLng(SalesRepID)), False)
End Function

Public Function CreateVisitPlan(ByVal SalesRepID As Variant, ByVal VisitDate As Variant, ByRef Added As Long) As String
    ' The visits of the day from the route of its weekday; customers already in the day are kept as they are.
    Dim db As DAO.Database, d As Date, stops As Long
    Added = 0
    CreateVisitPlan = CanRoutes()
    If Len(CreateVisitPlan) > 0 Then Exit Function
    If Not RepActive(SalesRepID) Then
        CreateVisitPlan = "اختر مندوبًا نشطًا."
        Exit Function
    End If
    If Not IsDate(VisitDate) Then
        CreateVisitPlan = "اكتب اليوم."
        Exit Function
    End If
    d = DateValue(VisitDate)
    stops = Nz(DbValue("SELECT Count(*) FROM RouteStops WHERE IsActive = True AND SalesRepID = " & CLng(SalesRepID) & _
                       " AND VisitWeekday = " & VisitWeekdayOf(d)), 0)
    If stops = 0 Then
        CreateVisitPlan = "لا عملاء في خط سير المندوب ليوم " & RouteWeekdayName(VisitWeekdayOf(d)) & ". أضفهم في «خط السير»."
        Exit Function
    End If
    Set db = CurrentDb
    db.Execute "INSERT INTO RepVisits (VisitDate, SalesRepID, CustomerID, IsPlanned, StopOrder, Result, EmployeeID) " & _
               "SELECT " & SqlDate(d) & ", s.SalesRepID, s.CustomerID, True, s.StopOrder, 'PENDING', " & CurrentUserID() & _
               " FROM RouteStops AS s INNER JOIN Customers AS c ON s.CustomerID = c.CustomerID " & _
               "WHERE s.IsActive = True AND c.IsActive = True AND s.SalesRepID = " & CLng(SalesRepID) & _
               " AND s.VisitWeekday = " & VisitWeekdayOf(d) & " AND s.CustomerID NOT IN (SELECT v.CustomerID FROM " & _
               "RepVisits AS v WHERE v.VisitDate = " & SqlDate(d) & " AND v.SalesRepID = " & CLng(SalesRepID) & ")", _
               dbFailOnError
    Added = db.RecordsAffected
    If Added > 0 Then LogAction "REP_VISIT_PLAN", "RepVisits", "", "Rep=" & SalesRepID & " Date=" & Format$(d, "yyyy-mm-dd") & _
                                " Added=" & Added
End Function

Public Function AddRepVisit(ByVal SalesRepID As Variant, ByVal CustomerID As Variant, ByVal VisitDate As Variant, _
                            ByRef NewVisitID As Long) As String
    Dim rs As DAO.Recordset, d As Date
    NewVisitID = 0
    AddRepVisit = CanRoutes()
    If Len(AddRepVisit) > 0 Then Exit Function
    If Not RepActive(SalesRepID) Then
        AddRepVisit = "اختر مندوبًا نشطًا."
    ElseIf IsNull(CustomerID) Then
        AddRepVisit = "اختر العميل."
    ElseIf Not IsDate(VisitDate) Then
        AddRepVisit = "اكتب اليوم."
    End If
    If Len(AddRepVisit) > 0 Then Exit Function
    d = DateValue(VisitDate)
    If Nz(DbValue("SELECT Count(*) FROM RepVisits WHERE VisitDate = " & SqlDate(d) & " AND SalesRepID = " & _
                  CLng(SalesRepID) & " AND CustomerID = " & CLng(CustomerID)), 0) > 0 Then
        AddRepVisit = "العميل في زيارات هذا اليوم."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("RepVisits", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!VisitDate = d
    rs!SalesRepID = CLng(SalesRepID)
    rs!CustomerID = CLng(CustomerID)
    rs!IsPlanned = False
    rs!StopOrder = Nz(DbValue("SELECT Max(StopOrder) FROM RepVisits WHERE VisitDate = " & SqlDate(d) & _
                              " AND SalesRepID = " & CLng(SalesRepID)), 0) + 1
    rs!Result = "PENDING"
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    NewVisitID = rs!RepVisitID
    rs.Close
End Function

Private Function DocsOnDay(ByVal TableName As String, ByVal DateField As String, ByVal CustomerID As Long, _
                           ByVal d As Date) As Long
    DocsOnDay = Nz(DbValue("SELECT Count(*) FROM " & TableName & " WHERE CustomerID = " & CustomerID & " AND " & _
                           DateField & " >= " & SqlDate(d) & " AND " & DateField & " < " & SqlDate(d + 1)), 0)
End Function

Public Function SuggestedVisitResult(ByVal VisitID As Long) As String
    ' What the customer did that day: a sale, else an order, else a payment, else nothing known.
    Dim cust As Long, d As Date
    If IsNull(VisitField(VisitID, "RepVisitID")) Then Exit Function
    cust = VisitField(VisitID, "CustomerID")
    d = VisitField(VisitID, "VisitDate")
    If DocsOnDay("SalesInvoices", "InvoiceDate", cust, d) > 0 Then
        SuggestedVisitResult = "SALE"
    ElseIf DocsOnDay("SalesOrders", "OrderDate", cust, d) > 0 Then
        SuggestedVisitResult = "ORDER"
    ElseIf DocsOnDay("CustomerPayments", "PaymentDate", cust, d) > 0 Then
        SuggestedVisitResult = "PAYMENT"
    End If
End Function

Public Function VisitResultValid(ByVal Result As String) As Boolean
    Select Case Result
        Case "ORDER", "SALE", "PAYMENT", "NO_SALE", "CLOSED"
            VisitResultValid = True
    End Select
End Function

Public Function RecordRepVisit(ByVal VisitID As Long, ByVal Result As Variant, ByVal Notes As Variant) As String
    Dim res As String, rs As DAO.Recordset, auditBefore As Collection
    RecordRepVisit = CanRoutes()
    If Len(RecordRepVisit) > 0 Then Exit Function
    If IsNull(VisitField(VisitID, "RepVisitID")) Then
        RecordRepVisit = "اختر الزيارة."
        Exit Function
    End If
    res = Nz(Result, "")
    If Len(res) = 0 Then res = SuggestedVisitResult(VisitID)
    If Len(res) = 0 Then
        RecordRepVisit = "اختر نتيجة الزيارة: لا فاتورة ولا طلبية ولا تحصيل للعميل في هذا اليوم."
        Exit Function
    End If
    If Not VisitResultValid(res) Then
        RecordRepVisit = "نتيجة الزيارة غير صحيحة."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("RepVisits", "RepVisitID", VisitID)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM RepVisits WHERE RepVisitID = " & VisitID, dbOpenDynaset)
    rs.Edit
    rs!Result = res
    rs!VisitedAt = Now
    rs!EmployeeID = CurrentUserID()
    If Len(Nz(Notes, "")) > 0 Then rs!Notes = Left$(CStr(Notes), 255)
    rs.Update
    rs.Close
    AuditEdited "REP_VISIT_RECORD", "RepVisits", "RepVisitID", VisitID, auditBefore
End Function

Public Function DeleteRepVisit(ByVal VisitID As Long) As String
    ' A visit not recorded yet (planned by mistake, or the customer left the route).
    Dim auditBefore As Collection
    DeleteRepVisit = CanRoutes()
    If Len(DeleteRepVisit) > 0 Then Exit Function
    If Nz(VisitField(VisitID, "Result"), "") <> "PENDING" Then
        DeleteRepVisit = "تُحذف الزيارة التي لم تُسجَّل نتيجتها فقط."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("RepVisits", "RepVisitID", VisitID)
    CurrentDb.Execute "DELETE FROM RepVisits WHERE RepVisitID = " & VisitID, dbFailOnError
    AuditDeleted "REP_VISIT_DELETE", "RepVisits", VisitID, auditBefore
End Function

'==============================================================================
' Screen frmRepVisits
'==============================================================================
Public Function VisitResultRows() As String
    VisitResultRows = "ORDER;" & Tr("طلبية") & ";SALE;" & Tr("بيع") & ";PAYMENT;" & Tr("تحصيل") & _
                      ";NO_SALE;" & Tr("بلا نتيجة") & ";CLOSED;" & Tr("المحل مغلق")
End Function

Public Sub RepVisitsLoad(ByVal frm As Access.Form)
    Dim rep As Long
    Calendar = vbCalGreg
    frm!cboRep.RowSource = Tr("SELECT s.SalesRepID, s.RepName FROM [@SalesReps] AS s WHERE s.IsActive = True ORDER BY s.RepName")
    frm!cboCustomer.RowSource = Tr("SELECT c.CustomerID, c.CustomerName FROM [@Customers] AS c WHERE c.IsActive = True " & _
                                   "AND c.IsSystem = False ORDER BY c.CustomerName")
    frm!cboResult.RowSource = VisitResultRows()
    frm!txtDate.Value = Date
    rep = SessionRepID()
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then rep = CLng(frm.OpenArgs)
    End If
    If rep > 0 Then frm!cboRep.Value = rep
    RepVisitsShow frm
End Sub

Private Function ScreenDate(ByVal frm As Access.Form) As Variant
    ScreenDate = Null
    If IsDate(frm!txtDate.Value) Then ScreenDate = DateValue(frm!txtDate.Value)
End Function

Public Sub RepVisitsShow(ByVal frm As Access.Form)
    Dim rep As Long, d As Variant, info As String, where As String, total As Long, done As Long, outside As Long
    rep = Nz(frm!cboRep.Value, 0)
    d = ScreenDate(frm)
    If rep = 0 Or IsNull(d) Then
        frm!lstVisits.RowSource = ""
        frm!lblInfo.Caption = Tr("اختر المندوب واليوم.")
        Exit Sub
    End If
    where = " WHERE SalesRepID = " & rep & " AND VisitDate = " & SqlDate(CDate(d))
    frm!lstVisits.RowSource = Tr("SELECT RepVisitID, StopOrder AS [الترتيب], CustomerName AS [العميل], PlanName AS [الخطة], " & _
        "ResultName AS [النتيجة], Format(VisitedAt, 'hh:nn') AS [الوقت], Notes AS [ملاحظات الزيارة] FROM RepVisitListQuery" & _
        where & " ORDER BY StopOrder, CustomerName")
    frm!lstVisits.Requery
    total = Nz(DbValue("SELECT Count(*) FROM RepVisits" & where), 0)
    done = Nz(DbValue("SELECT Count(*) FROM RepVisits" & where & " AND Result <> 'PENDING'"), 0)
    outside = Nz(DbValue("SELECT Count(*) FROM RepVisits" & where & " AND IsPlanned = False"), 0)
    If total = 0 Then
        info = "لا زيارات في هذا اليوم. «خطة اليوم من خط السير» تضيف عملاء يوم " & RouteWeekdayName(VisitWeekdayOf(CDate(d))) & "."
    Else
        info = "الزيارات: " & total & "    تمت: " & done & "    لم تتم: " & (total - done) & "    خارج خط السير: " & outside
    End If
    frm!lblInfo.Caption = Tr(info)
End Sub

Public Sub RepVisitPicked(ByVal frm As Access.Form)
    Dim id As Long, res As String
    id = Nz(frm!lstVisits.Value, 0)
    If id = 0 Then Exit Sub
    res = Nz(VisitField(id, "Result"), "")
    If res = "PENDING" Then res = SuggestedVisitResult(id)
    If Len(res) = 0 Then
        frm!cboResult.Value = Null
    Else
        frm!cboResult.Value = res
    End If
    frm!txtVisitNotes.Value = VisitField(id, "Notes")
End Sub

Public Sub RepVisitsPlan(ByVal frm As Access.Form)
    Dim msg As String, added As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateVisitPlan(frm!cboRep.Value, frm!txtDate.Value, added)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    RepVisitsShow frm
    If added = 0 Then ShowInfo "عملاء خط السير موجودون في زيارات هذا اليوم."
End Sub

Public Sub RepVisitsAdd(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = AddRepVisit(frm!cboRep.Value, frm!cboCustomer.Value, frm!txtDate.Value, id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!cboCustomer.Value = Null
    RepVisitsShow frm
    frm!lstVisits.Value = id
End Sub

Public Sub RepVisitsRecord(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Sub
    id = Nz(frm!lstVisits.Value, 0)
    If id = 0 Then
        ShowWarning "اختر الزيارة من القائمة."
        Exit Sub
    End If
    msg = RecordRepVisit(id, frm!cboResult.Value, frm!txtVisitNotes.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtVisitNotes.Value = Null
    RepVisitsShow frm
End Sub

Public Sub RepVisitsDelete(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "DELETE") Then Exit Sub
    id = Nz(frm!lstVisits.Value, 0)
    If id = 0 Then
        ShowWarning "اختر الزيارة من القائمة."
        Exit Sub
    End If
    If Not AskYesNo("حذف الزيارة؟") Then Exit Sub
    msg = DeleteRepVisit(id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    RepVisitsShow frm
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestRepRoutes() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, msg As String, rep As Long, c1 As Long, c2 As Long, pid As Long
    Dim passed As Long, failed As Long, report As String, added As Long, v1 As Long, v2 As Long, v3 As Long
    Dim invID As Long, today As Date
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestRepRoutes  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    StopActingAsRep
    today = Date
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CurrentDb.Execute "UPDATE SalesReps SET EmployeeID = Null WHERE EmployeeID = " & CurrentUserID(), dbFailOnError
    CurrentDb.Execute "INSERT INTO SalesReps (RepCode, RepName, CommissionRate, CommissionBase, IsActive, EmployeeID, " & _
                      "CreditLimit, MaxDiscountPercent) VALUES ('TEST-RR', 'TEST مندوب خط السير', 0, 'SALES', True, " & _
                      CurrentUserID() & ", 100, 0.05)", dbFailOnError
    rep = DbValue("SELECT SalesRepID FROM SalesReps WHERE RepCode = 'TEST-RR'")
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, OpeningBalance, CurrentBalance, AllowCredit, CreditLimit, IsActive, " & _
                      "IsSystem, SalesRepID) VALUES ('TEST عميل خط السير 1', 0, 0, True, 0, True, False, " & rep & ")", dbFailOnError
    c1 = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST عميل خط السير 1'")
    CurrentDb.Execute "INSERT INTO Customers (CustomerName, OpeningBalance, CurrentBalance, AllowCredit, CreditLimit, IsActive, " & _
                      "IsSystem, SalesRepID) VALUES ('TEST عميل خط السير 2', 0, 0, True, 0, True, False, " & rep & ")", dbFailOnError
    c2 = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST عميل خط السير 2'")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-RR-P', 'TEST صنف المندوب', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-RR-P'")
    CurrentDb.Execute "INSERT INTO RepPrices (SalesRepID, ProductID, UnitPrice) VALUES (" & rep & ", " & pid & ", 18)", _
                      dbFailOnError

    ' the route: customer 1 on today's weekday, customer 2 on another day
    msg = CreateVisitPlan(rep, today, added)
    CheckRr Len(msg) > 0 And added = 0, "لا خطة بلا عملاء في خط السير لهذا اليوم", passed, failed, report
    CurrentDb.Execute "INSERT INTO RouteStops (SalesRepID, CustomerID, VisitWeekday, StopOrder, IsActive) VALUES (" & _
                      rep & ", " & c1 & ", " & VisitWeekdayOf(today) & ", 1, True)", dbFailOnError
    CurrentDb.Execute "INSERT INTO RouteStops (SalesRepID, CustomerID, VisitWeekday, StopOrder, IsActive) VALUES (" & _
                      rep & ", " & c2 & ", " & (VisitWeekdayOf(today) Mod 7) + 1 & ", 1, True)", dbFailOnError
    msg = CreateVisitPlan(rep, today, added)
    v1 = Nz(DbValue("SELECT RepVisitID FROM RepVisits WHERE SalesRepID = " & rep & " AND CustomerID = " & c1), 0)
    CheckRr Len(msg) = 0 And added = 1 And v1 > 0 And VisitField(v1, "IsPlanned") = True, _
            "خطة اليوم: عملاء يوم الأسبوع فقط " & msg, passed, failed, report
    msg = CreateVisitPlan(rep, today, added)
    CheckRr Len(msg) = 0 And added = 0, "الخطة لا تكرر العميل", passed, failed, report
    msg = AddRepVisit(rep, c2, today, v2)
    CheckRr Len(msg) = 0 And v2 > 0 And VisitField(v2, "IsPlanned") = False, "زيارة خارج خط السير " & msg, _
            passed, failed, report
    CheckRr Len(AddRepVisit(rep, c2, today, v3)) > 0, "لا زيارتان لعميل في يوم", passed, failed, report

    ' the results: an order of customer 1 today, customer 2 closed
    CheckRr Len(RecordRepVisit(v1, Null, Null)) > 0, "لا نتيجة بلا مستند ولا اختيار", passed, failed, report
    CurrentDb.Execute "INSERT INTO SalesOrders (OrderNumber, OrderDate, CustomerID, WarehouseID, Status, EmployeeID) " & _
                      "VALUES ('TEST-SO-RR', Now(), " & c1 & ", 1, 'DRAFT', " & CurrentUserID() & ")", dbFailOnError
    msg = RecordRepVisit(v1, Null, "TEST")
    CheckRr Len(msg) = 0 And VisitField(v1, "Result") = "ORDER" And Not IsNull(VisitField(v1, "VisitedAt")), _
            "النتيجة من مستندات اليوم: طلبية " & msg, passed, failed, report
    msg = RecordRepVisit(v2, "CLOSED", Null)
    CheckRr Len(msg) = 0 And VisitField(v2, "Result") = "CLOSED", "تسجيل «المحل مغلق» " & msg, passed, failed, report
    CheckRr Len(DeleteRepVisit(v1)) > 0, "لا تُحذف زيارة مسجلة", passed, failed, report
    msg = AddRepVisit(rep, c2, today + 1, v3)
    CheckRr Len(msg) = 0 And Len(DeleteRepVisit(v3)) = 0 And IsNull(VisitField(v3, "RepVisitID")), _
            "حذف زيارة لم تتم " & msg, passed, failed, report
    CheckRr Nz(DbValue("SELECT VisitCount FROM RouteExecutionQuery WHERE SalesRepID = " & rep), 0) = 2 And _
            Nz(DbValue("SELECT PlannedCount FROM RouteExecutionQuery WHERE SalesRepID = " & rep), 0) = 1 And _
            Nz(DbValue("SELECT DoneCount FROM RouteExecutionQuery WHERE SalesRepID = " & rep), 0) = 2 And _
            Nz(DbValue("SELECT OrderCount FROM RouteExecutionQuery WHERE SalesRepID = " & rep), 0) = 1 And _
            Nz(DbValue("SELECT ClosedCount FROM RouteExecutionQuery WHERE SalesRepID = " & rep), 0) = 1, _
            "تقرير تنفيذ خط السير", passed, failed, report

    ' the limits of the rep who sells (the rep of the current user)
    CheckRr SessionRepID() = rep And SessionPrice(pid, 20) = 18 And IsSessionRepPrice(pid, 18) And _
            RepDiscountLimit(0.5) = 0.05, "سعر المندوب وحد خصمه", passed, failed, report
    CheckRr Len(RepCreditProblem(rep, 50)) = 0 And Len(RepCreditProblem(rep, 150)) > 0, "حد ائتمان المندوب", _
            passed, failed, report
    CurrentDb.Execute "UPDATE Customers SET CurrentBalance = 90 WHERE CustomerID = " & c1, dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "Available) VALUES (" & pid & ", 'TEST-RR-P', 'TEST صنف المندوب', 1, 18, 0, 0)", dbFailOnError
    msg = PostSaleFromCart(c2, "CREDIT", CASH_METHOD_ID, 0, Null, "TEST", invID)
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CheckRr invID = 0 And InStr(msg, Nz(DbValue(Tr("SELECT s.RepName FROM [@SalesReps] AS s WHERE s.SalesRepID = " & rep)), "#")) > 0, _
            "البيع الآجل يُرفض فوق حد ائتمان المندوب", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckRr False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات خط السير وحدود المندوب ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestRepRoutes"
        TestRepRoutes = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestRepRoutes"
    End If
End Function

Private Sub CheckRr(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                   ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "  [OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "  [X]  " & Title
    End If
End Sub
