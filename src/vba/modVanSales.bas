Attribute VB_Name = "modVanSales"
'==============================================================================
' modVanSales  -  Retail Store Management System (the rep day and van sales, docs/63, plan 4B)
'
' A rep with a van warehouse (Warehouses.WarehouseType = VAN, SalesRepID) and a custody cash box
' (CashBoxes.BoxType = CUSTODY, SalesRepID) works one day at a time:
'   OpenRepDay          the day of a rep (one open day per rep), with the warehouse the van loads from
'   LoadRepDay          a draft stock transfer store -> van (modWarehouses), completed in frmStockTransfer
'   StartActingAsRep    the office enters the rep's documents: while it acts, CurrentWarehouseID = the van,
'                       CurrentCashBoxID = the custody box, SalesRepFor = the rep, CanUseBox allows the box
'                       (the rep logged in with his own van and box needs none of this)
'   CountRepDay         the stock count of the van (modPurchases): its differences are the goods shortage
'   UnloadRepDay        a transfer van -> store with everything left in the van
'   (cash)              the cash closing of the custody box (modCash.PostCashClosing): the cash difference and
'                       the cash handed to the treasury
'   SettleRepDay        the day is SETTLED once its transfers are complete, the van count is posted and the box
'                       is closed after the day was opened; it keeps the count and the closing
' RepDaySummaryQuery adds up the day: the moves of the van, the sales, returns and collections in the box
' between OpenedAt and SettledAt. No journal entry of its own: the documents post theirs.
' Screen: frmRepDay (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Private m_ActDay As Long
Private m_ActWarehouse As Long
Private m_ActBox As Long
Private m_ActRep As Long

'==============================================================================
' Acting as the rep
'==============================================================================
Public Function ActingRepDayID() As Long
    ActingRepDayID = m_ActDay
End Function

Public Function ActingWarehouseID() As Long
    ActingWarehouseID = m_ActWarehouse
End Function

Public Function ActingCashBoxID() As Long
    ActingCashBoxID = m_ActBox
End Function

Public Function ActingSalesRepID() As Long
    ActingSalesRepID = m_ActRep
End Function

Public Function StartActingAsRep(ByVal RepDayID As Long) As String
    Dim rs As DAO.Recordset
    If RepDayStatus(RepDayID) <> "OPEN" Then
        StartActingAsRep = "يوم المندوب غير مفتوح."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT WarehouseID, CashBoxID, SalesRepID FROM RepDays WHERE RepDayID = " & RepDayID, _
                                     dbOpenSnapshot)
    m_ActWarehouse = rs!WarehouseID
    m_ActBox = rs!CashBoxID
    m_ActRep = rs!SalesRepID
    rs.Close
    m_ActDay = RepDayID
End Function

Public Sub StopActingAsRep()
    m_ActDay = 0
    m_ActWarehouse = 0
    m_ActBox = 0
    m_ActRep = 0
End Sub

'==============================================================================
' The day
'==============================================================================
Private Function RepDayStatus(ByVal RepDayID As Long) As String
    RepDayStatus = Nz(DbValue("SELECT Status FROM RepDays WHERE RepDayID = " & RepDayID), "")
End Function

Private Function DayField(ByVal RepDayID As Long, ByVal FieldName As String) As Variant
    DayField = DbValue("SELECT " & FieldName & " FROM RepDays WHERE RepDayID = " & RepDayID)
End Function

Public Function RepVanID(ByVal SalesRepID As Long) As Long
    RepVanID = Nz(DbValue("SELECT Min(WarehouseID) FROM Warehouses WHERE WarehouseType = 'VAN' AND IsActive = True " & _
                          "AND SalesRepID = " & SalesRepID), 0)
End Function

Public Function RepCustodyBoxID(ByVal SalesRepID As Long) As Long
    RepCustodyBoxID = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'CUSTODY' AND IsActive = True " & _
                                 "AND SalesRepID = " & SalesRepID), 0)
End Function

Public Function OpenRepDay(ByVal SalesRepID As Variant, ByVal StoreID As Variant, ByRef NewDayID As Long) As String
    Dim rs As DAO.Recordset, van As Long, box As Long, openNo As Variant
    NewDayID = 0
    If Not HasPermission("REP_DAYS") Then
        OpenRepDay = "ليست لديك صلاحية يوم المندوب."
        Exit Function
    End If
    If IsNull(SalesRepID) Then
        OpenRepDay = "اختر المندوب."
        Exit Function
    End If
    If Nz(DbValue("SELECT IsActive FROM SalesReps WHERE SalesRepID = " & CLng(SalesRepID)), False) = False Then
        OpenRepDay = "المندوب غير موجود أو غير نشط."
        Exit Function
    End If
    van = RepVanID(CLng(SalesRepID))
    box = RepCustodyBoxID(CLng(SalesRepID))
    If van = 0 Then
        OpenRepDay = "ليس للمندوب مخزن سيارة نشط. أنشئه في «المخازن» بنوع «سيارة مندوب» واختر المندوب."
    ElseIf box = 0 Then
        OpenRepDay = "ليس للمندوب صندوق عهدة نشط. أنشئه في «الصناديق» بنوع «عهدة مندوب» واختر المندوب."
    ElseIf Not WarehouseIsActive(StoreID) Then
        OpenRepDay = "اختر المخزن الذي تُحمَّل منه السيارة."
    ElseIf CLng(StoreID) = van Then
        OpenRepDay = "المخزن الذي تُحمَّل منه السيارة غير مخزن السيارة."
    End If
    If Len(OpenRepDay) > 0 Then Exit Function
    openNo = DbValue("SELECT DayNumber FROM RepDays WHERE Status = 'OPEN' AND SalesRepID = " & CLng(SalesRepID))
    If Not IsNull(openNo) Then
        OpenRepDay = "للمندوب يوم مفتوح (" & openNo & "). سوِّه أولًا."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("RepDays", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!DayNumber = NextNumber("REP_DAY")
    rs!DayDate = Date
    rs!SalesRepID = CLng(SalesRepID)
    rs!WarehouseID = van
    rs!CashBoxID = box
    rs!StoreID = CLng(StoreID)
    rs!Status = "OPEN"
    rs!OpenedAt = Now
    rs!EmployeeID = CurrentUserID()
    rs.Update
    rs.Bookmark = rs.LastModified
    NewDayID = rs!RepDayID
    rs.Close
    LogAction "REP_DAY_OPEN", "RepDays", CStr(NewDayID), "Rep=" & SalesRepID
End Function

Private Function DayHasActivity(ByVal RepDayID As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT LoadedValue + SoldCost + ReturnedCost + UnloadedValue + Abs(CountDiffValue) AS M, " & _
                                     "InvoiceCount, ReturnsTotal, Collections FROM RepDaySummaryQuery WHERE RepDayID = " & _
                                     RepDayID, dbOpenSnapshot)
    If Not rs.EOF Then DayHasActivity = (rs!M <> 0 Or rs!InvoiceCount > 0 Or rs!ReturnsTotal <> 0 Or rs!Collections <> 0)
    rs.Close
End Function

Public Function DeleteRepDay(ByVal RepDayID As Long) As String
    ' A day opened by mistake: open, and nothing was loaded, sold, returned or collected in it.
    Dim auditBefore As Collection
    If Not HasPermission("REP_DAYS") Then
        DeleteRepDay = "ليست لديك صلاحية يوم المندوب."
    ElseIf RepDayStatus(RepDayID) <> "OPEN" Then
        DeleteRepDay = "يُحذف اليوم المفتوح فقط."
    ElseIf DayHasActivity(RepDayID) Then
        DeleteRepDay = "في هذا اليوم حركة (تحميل أو بيع أو تحصيل). سوِّه بدل حذفه."
    End If
    If Len(DeleteRepDay) > 0 Then Exit Function
    If ActingRepDayID() = RepDayID Then StopActingAsRep
    Set auditBefore = AuditSnapshot("RepDays", "RepDayID", RepDayID)
    CurrentDb.Execute "DELETE FROM RepDays WHERE RepDayID = " & RepDayID, dbFailOnError
    AuditDeleted "REP_DAY_DELETE", "RepDays", RepDayID, auditBefore
End Function

Public Function LoadRepDay(ByVal RepDayID As Long, ByRef NewTransferID As Long) As String
    ' A draft transfer from the store to the van; its lines are entered in frmStockTransfer, then it is sent.
    NewTransferID = 0
    If RepDayStatus(RepDayID) <> "OPEN" Then
        LoadRepDay = "يوم المندوب غير مفتوح."
        Exit Function
    End If
    LoadRepDay = CreateStockTransfer(DayField(RepDayID, "StoreID"), DayField(RepDayID, "WarehouseID"), _
                                     "تحميل " & DayField(RepDayID, "DayNumber"), NewTransferID)   ' modWarehouses
End Function

Public Function UnloadRepDay(ByVal RepDayID As Long, ByRef NewTransferID As Long) As String
    ' A draft transfer from the van back to the store with everything left in the van.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, ln As DAO.Recordset, inTrans As Boolean
    Dim van As Long, msg As String
    NewTransferID = 0
    If RepDayStatus(RepDayID) <> "OPEN" Then
        UnloadRepDay = "يوم المندوب غير مفتوح."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    If Nz(DbValue("SELECT COUNT(*) FROM StockTransfers WHERE Status = 'DRAFT' AND FromWarehouseID = " & van), 0) > 0 Then
        UnloadRepDay = "للسيارة تحويل مسودة لم يُرسَل. أرسله أو احذفه أولًا."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM qryWarehouseStock WHERE StockQty > 0 AND WarehouseID = " & van), 0) = 0 Then
        UnloadRepDay = "السيارة فارغة: لا شيء يُفرَّغ."
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateStockTransfer(van, DayField(RepDayID, "StoreID"), "تفريغ " & DayField(RepDayID, "DayNumber"), NewTransferID)
    If Len(msg) = 0 Then
        Set rs = db.OpenRecordset("SELECT ProductID, StockQty FROM qryWarehouseStock WHERE StockQty > 0 AND WarehouseID = " & _
                                  van & " ORDER BY ProductID", dbOpenSnapshot)
        Set ln = db.OpenRecordset("StockTransferLines", dbOpenDynaset, dbAppendOnly)
        Do Until rs.EOF
            ln.AddNew
            ln!StockTransferID = NewTransferID
            ln!ProductID = rs!ProductID
            ln!Quantity = rs!StockQty
            ln!UnitCost = 0                                        ' set when it is sent
            ln.Update
            rs.MoveNext
        Loop
        ln.Close
        rs.Close
    End If
    If Len(msg) > 0 Then
        ws.Rollback
        NewTransferID = 0
        UnloadRepDay = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    Exit Function
EH:
    UnloadRepDay = "تعذر إنشاء تحويل التفريغ: " & Err.Description & " (" & Err.Number & ")"
    NewTransferID = 0
    If inTrans Then ws.Rollback
End Function

Public Function CountRepDay(ByVal RepDayID As Long, ByRef CountID As Long) As String
    ' The open count of the van, else a new one.
    Dim van As Long
    CountID = 0
    If RepDayStatus(RepDayID) <> "OPEN" Then
        CountRepDay = "يوم المندوب غير مفتوح."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    CountID = Nz(DbValue("SELECT Max(StockCountID) FROM StockCounts WHERE Status = 'OPEN' AND WarehouseID = " & van), 0)
    If CountID = 0 Then
        CountRepDay = CreateStockCount(Null, "جرد سيارة " & DayField(RepDayID, "DayNumber"), CountID, van)   ' modPurchases
    End If
End Function

Public Function SettleRepDay(ByVal RepDayID As Long) As String
    ' The transfers of the van are complete, its count is posted and the custody box is closed since the day opened.
    Dim van As Long, box As Long, countID As Variant, closingID As Variant, since As String
    Dim advanceTo As Variant, shortage As Currency
    If Not HasPermission("REP_DAYS") Then
        SettleRepDay = "ليست لديك صلاحية يوم المندوب."
        Exit Function
    End If
    If RepDayStatus(RepDayID) <> "OPEN" Then
        SettleRepDay = "يوم المندوب غير مفتوح."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    box = DayField(RepDayID, "CashBoxID")
    since = "(SELECT OpenedAt FROM RepDays WHERE RepDayID = " & RepDayID & ")"
    If Nz(DbValue("SELECT COUNT(*) FROM StockTransfers WHERE Status In ('DRAFT', 'SENT') AND (FromWarehouseID = " & van & _
                  " OR ToWarehouseID = " & van & ")"), 0) > 0 Then
        SettleRepDay = "للسيارة تحويلات لم تكتمل (مسودة أو بالطريق). أرسلها واستلمها أو احذفها أولًا."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM StockCounts WHERE Status = 'OPEN' AND WarehouseID = " & van), 0) > 0 Then
        SettleRepDay = "جرد السيارة مفتوح. أدخل الكميات الفعلية ورحّله أولًا."
        Exit Function
    End If
    countID = DbValue("SELECT Max(StockCountID) FROM StockCounts WHERE Status = 'POSTED' AND WarehouseID = " & van & _
                      " AND PostedAt >= " & since)
    If IsNull(countID) Then
        SettleRepDay = "اجرد السيارة ورحّل الجرد قبل إغلاق اليوم (زر «جرد السيارة»)."
        Exit Function
    End If
    closingID = DbValue("SELECT Max(ClosingID) FROM CashClosings WHERE CashBoxID = " & box & " AND ClosingDate >= " & since)
    If IsNull(closingID) Then
        SettleRepDay = "صفِّ صندوق العهدة قبل إغلاق اليوم (زر «تصفية الصندوق»)."
        Exit Function
    End If
    ' the net goods shortage of the day is an advance on the rep when he has a program user (docs/66)
    advanceTo = DbValue("SELECT r.EmployeeID FROM RepDays AS d INNER JOIN SalesReps AS r ON d.SalesRepID = r.SalesRepID " & _
                        "WHERE d.RepDayID = " & RepDayID)
    shortage = -RoundMoney(Nz(DbValue("SELECT CountDiffValue FROM RepDaySummaryQuery WHERE RepDayID = " & RepDayID), 0))
    If shortage < 0 Or IsNull(advanceTo) Then shortage = 0
    CurrentDb.Execute "UPDATE RepDays SET Status = 'SETTLED', SettledAt = Now(), StockCountID = " & countID & _
                      ", ClosingID = " & closingID & ", GoodsShortage = " & Str$(shortage) & ", AdvanceEmployeeID = " & _
                      IIf(shortage > 0, CLng(Nz(advanceTo, 0)), "Null") & " WHERE RepDayID = " & RepDayID, dbFailOnError
    If ActingRepDayID() = RepDayID Then StopActingAsRep
    LogAction "REP_DAY_SETTLE", "RepDays", CStr(RepDayID), "Count=" & countID & " Closing=" & closingID & _
              " Shortage=" & Format$(shortage, "0.00")
    If shortage > 0 Then SyncJournal
End Function

'==============================================================================
' frmRepDay (OpenArgs = a day to show)
'==============================================================================
Public Sub RepDayLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboRep.RowSource = Tr("SELECT r.SalesRepID, r.RepName FROM [@SalesReps] AS r WHERE r.IsActive = True AND " & _
                              "r.SalesRepID In (SELECT SalesRepID FROM Warehouses WHERE WarehouseType = 'VAN' AND " & _
                              "IsActive = True) ORDER BY r.RepName")
    frm!cboStore.RowSource = WarehouseRows()
    frm!cboStore.Value = MAIN_WAREHOUSE_ID
    RepDayListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(RepDayID) FROM RepDays")
    frm!cboDay.Value = id
    RepDayShow frm
End Sub

Private Sub RepDayListRefresh(ByVal frm As Access.Form)
    frm!cboDay.RowSource = Tr("SELECT d.RepDayID, d.DayNumber, Format(d.DayDate, 'yyyy/mm/dd'), r.RepName, w.WarehouseName, " & _
        "IIf(d.Status = 'OPEN', 'مفتوح', 'تمت تسويته') FROM (RepDays AS d INNER JOIN [@SalesReps] AS r " & _
        "ON d.SalesRepID = r.SalesRepID) INNER JOIN [@Warehouses] AS w ON d.WarehouseID = w.WarehouseID " & _
        "ORDER BY d.RepDayID DESC")
    frm!cboDay.Requery
End Sub

Private Function SummaryItem(ByVal Caption As String, ByVal Amount As Variant) As String
    ' One row of the summary list (a value list: quoted, so commas in the amounts stay in their column).
    SummaryItem = """" & Tr(Caption) & """;""" & Format$(Nz(Amount, 0), "#,##0.00") & """;"
End Function

Private Function SummaryHead(ByVal Caption As String) As String
    SummaryHead = """" & Tr(Caption) & """;"""";"
End Function

Public Function RepDaySummaryRows(ByVal RepDayID As Long) As String
    Dim rs As DAO.Recordset, rows As String, vanValue As Variant, advance As Currency
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM RepDaySummaryQuery WHERE RepDayID = " & RepDayID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    vanValue = DbValue("SELECT Sum(q.StockQty * p.AverageCost) FROM qryWarehouseStock AS q INNER JOIN Products AS p " & _
                       "ON q.ProductID = p.ProductID WHERE q.WarehouseID = " & rs!WarehouseID)
    rows = SummaryHead("البضاعة (بالتكلفة)") & _
           SummaryItem("المحمَّل للسيارة", rs!LoadedValue) & SummaryItem("المباع", rs!SoldCost) & _
           SummaryItem("المرتجع من العملاء", rs!ReturnedCost) & SummaryItem("المفرَّغ", rs!UnloadedValue) & _
           SummaryItem("فرق جرد السيارة (عجز بالسالب)", rs!CountDiffValue) & _
           SummaryItem("الباقي في السيارة الآن", vanValue) & _
           SummaryHead("النقدية") & _
           SummaryItem("فواتير البيع (" & rs!InvoiceCount & ")", rs!SalesTotal) & _
           SummaryItem("المقبوض نقدًا من الفواتير", rs!CashSales) & SummaryItem("التحصيل من العملاء", rs!Collections) & _
           SummaryItem("المرتجعات", rs!ReturnsTotal) & SummaryItem("المردود نقدًا", rs!CashRefunds) & _
           SummaryItem("رصيد صندوق العهدة الآن", CashBoxBalance(rs!CashBoxID))
    If Not IsNull(rs!CashCounted) Then
        rows = rows & SummaryItem("التصفية: المتوقع", rs!CashExpected) & SummaryItem("التصفية: الفعلي", rs!CashCounted) & _
               SummaryItem("التصفية: الفرق (عجز بالسالب)", rs!CashDifference) & _
               SummaryItem("المسلَّم للخزينة", rs!CashHandedOver)
        advance = Nz(DbValue("SELECT Sum(Amount) FROM CashVouchers WHERE Category = 'ADVANCE' AND ClosingID = " & _
                             rs!ClosingID), 0)
        If advance <> 0 Then rows = rows & SummaryItem("عجز النقدية: سلفة على المندوب", advance)
    End If
    If Nz(rs!GoodsShortage, 0) <> 0 Then rows = rows & SummaryItem("عجز البضاعة: سلفة على المندوب", rs!GoodsShortage)
    rs.Close
    RepDaySummaryRows = rows
End Function

Public Sub RepDayShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, info As String, rs As DAO.Recordset
    id = Nz(frm!cboDay.Value, 0)
    st = RepDayStatus(id)
    If Len(st) = 0 Then
        info = "لا يوجد يوم. اختر المندوب ثم «فتح يوم»."
        frm!lstSummary.RowSource = ""
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT DayNumber, DayDate, RepName, WarehouseName, BoxName, StatusName " & _
                                         "FROM RepDaySummaryQuery WHERE RepDayID = " & id, dbOpenSnapshot)
        info = rs!DayNumber & "    " & GDate(rs!DayDate) & "    " & rs!RepName & "    " & rs!WarehouseName & _
               "    " & rs!BoxName & "    " & rs!StatusName
        rs.Close
        frm!lstSummary.RowSource = RepDaySummaryRows(id)
    End If
    frm!lblInfo.Caption = Tr(info)
    If ActingRepDayID() > 0 Then
        frm!lblActing.Caption = Tr("تعمل الآن باسم المندوب: البيع والمرتجع والتحصيل من مخزن السيارة وصندوق العهدة " & _
                                   "حتى تُغلق هذه الشاشة.")
    Else
        frm!lblActing.Caption = " "
    End If
    SafeFocus frm!cboDay                    ' a focused control cannot be disabled
    frm!btnDeleteDay.Enabled = (st = "OPEN")
    frm!btnLoad.Enabled = (st = "OPEN")
    frm!btnSale.Enabled = (st = "OPEN")
    frm!btnReturn.Enabled = (st = "OPEN")
    frm!btnPayment.Enabled = (st = "OPEN")
    frm!btnCount.Enabled = (st = "OPEN")
    frm!btnUnload.Enabled = (st = "OPEN")
    frm!btnCash.Enabled = (st = "OPEN")
    frm!btnSettle.Enabled = (st = "OPEN")
    frm!btnPrint.Enabled = (Len(st) > 0)
End Sub

Public Sub RepDayOpen(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = OpenRepDay(frm!cboRep.Value, frm!cboStore.Value, id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    RepDayListRefresh frm
    frm!cboDay.Value = id
    RepDayShow frm
End Sub

Public Sub RepDayAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String, newID As Long
    id = Nz(frm!cboDay.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    Select Case Action
        Case "LOAD"
            msg = LoadRepDay(id, newID)
        Case "UNLOAD"
            msg = UnloadRepDay(id, newID)
        Case "COUNT"
            msg = StartActingAsRep(id)
            If Len(msg) = 0 Then msg = CountRepDay(id, newID)
        Case "SALE", "RETURN", "PAYMENT", "CASH"
            msg = StartActingAsRep(id)
        Case "SETTLE"
            If Not AskYesNo("إغلاق يوم المندوب؟ لا يُفتح بعد إغلاقه.") Then Exit Sub
            msg = SettleRepDay(id)
        Case "DELETE"
            If Not AskYesNo("حذف يوم المندوب؟") Then Exit Sub
            msg = DeleteRepDay(id)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Action = "DELETE" Then
        RepDayListRefresh frm
        frm!cboDay.Value = DbValue("SELECT Max(RepDayID) FROM RepDays")
    End If
    If Action = "SETTLE" Then RepDayListRefresh frm
    RepDayShow frm
    Select Case Action
        Case "LOAD", "UNLOAD"
            OpenScreen "frmStockTransfer", 0, newID            ' the lines, then «send»
        Case "COUNT"
            OpenScreen "frmStockCount", 7, newID
        Case "SALE"
            OpenScreen "frmPOS", 6
        Case "RETURN"
            OpenScreen "frmSalesReturn", 6
        Case "PAYMENT"
            OpenScreen "frmCustomerPayment", 6
        Case "CASH"
            OpenScreen "frmCashClosing", 0, ActingCashBoxID()
    End Select
End Sub

Public Sub RepDayPrint(ByVal frm As Access.Form)
    If IsNull(frm!cboDay.Value) Then Exit Sub
    OpenReportOrQuery "rptRepDays", "RepDaySummaryQuery", "RepDayID = " & CLng(frm!cboDay.Value)
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestRepDays() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, refNo As String, msg As String, cust As Long
    Dim passed As Long, failed As Long, report As String, rep As Long, van As Long, box As Long, mainBox As Long
    Dim dayID As Long, trID As Long, invID As Long, countID As Long, lines As Long, netValue As Currency
    Dim closingID As Long, expected As Currency, advBefore As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestRepDays  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    StopActingAsRep
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CurrentDb.Execute "UPDATE Settings SET AllowNegativeStock = False", dbFailOnError
    g_AutoAnswer = False
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-RD-P', 'TEST صنف السيارة', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-RD-P'")
    CurrentDb.Execute "INSERT INTO SalesReps (RepCode, RepName, CommissionRate, CommissionBase, IsActive) " & _
                      "VALUES ('TEST-RD', 'TEST مندوب السيارة', 0, 'SALES', True)", dbFailOnError
    rep = DbValue("SELECT SalesRepID FROM SalesReps WHERE RepCode = 'TEST-RD'")
    CurrentDb.Execute "INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType, SalesRepID, ConfirmReceipt, " & _
                      "IsActive) VALUES ('TEST-VAN', 'TEST سيارة', 'VAN', " & rep & ", False, True)", dbFailOnError
    van = DbValue("SELECT WarehouseID FROM Warehouses WHERE WarehouseCode = 'TEST-VAN'")
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive = True"), 0)
    mainBox = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True"), 0)
    msg = OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) > 0 And dayID = 0, "لا يوم لمندوب بلا صندوق عهدة", passed, failed, report
    CurrentDb.Execute "INSERT INTO CashBoxes (BoxName, BoxType, OpeningBalance, IsActive, SalesRepID, ConfirmReceipt) " & _
                      "VALUES ('TEST عهدة السيارة', 'CUSTODY', 0, True, " & rep & ", False)", dbFailOnError
    box = DbValue("SELECT CashBoxID FROM CashBoxes WHERE BoxName = 'TEST عهدة السيارة'")
    CheckRd cust > 0 And mainBox > 0, "يوجد عميل وخزينة رئيسية", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    msg = msg & OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) = 0 And RepDayStatus(dayID) = "OPEN", "فتح يوم المندوب " & msg, passed, failed, report
    CheckRd Len(OpenRepDay(rep, MAIN_WAREHOUSE_ID, lines)) > 0, "يوم مفتوح واحد لكل مندوب", passed, failed, report

    ' load 6 into the van
    msg = LoadRepDay(dayID, trID)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      trID & ", " & pid & ", 6, 0)", dbFailOnError
    msg = msg & SendStockTransfer(trID)
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 6, "تحميل 6 في السيارة " & msg, passed, failed, report

    ' a cash sale of 2 entered by the office in the rep's name
    msg = StartActingAsRep(dayID)
    CheckRd Len(msg) = 0 And CurrentWarehouseID() = van And CurrentCashBoxID() = box And SalesRepFor(cust) = rep, _
            "العمل باسم المندوب: مخزن السيارة وصندوقها والمندوب", passed, failed, report
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "Available) VALUES (" & pid & ", 'TEST-RD-P', 'TEST صنف السيارة', 2, 20, 0, 6)", dbFailOnError
    msg = PostSaleFromCart(cust, "CASH", CASH_METHOD_ID, 0, Null, "TEST", invID)
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 4 And _
            Nz(DbValue("SELECT CashBoxID FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) = box And _
            Nz(DbValue("SELECT SalesRepID FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) = rep, _
            "البيع من السيارة نقدًا لصندوق العهدة " & msg, passed, failed, report

    ' the van count finds 3: one short
    msg = CountRepDay(dayID, countID)
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = 3 WHERE StockCountID = " & countID & _
                      " AND ProductID = " & pid, dbFailOnError
    CheckRd Len(SettleRepDay(dayID)) > 0, "لا يُغلق اليوم وجرد السيارة مفتوح", passed, failed, report
    msg = msg & PostStockCount(countID, False, lines, netValue)
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 3 And netValue = -6, "جرد السيارة: عجز صنف واحد " & msg, _
            passed, failed, report

    ' unload the rest, then close the custody box
    msg = UnloadRepDay(dayID, trID)
    CheckRd Len(msg) = 0 And Nz(DbValue("SELECT Sum(Quantity) FROM StockTransferLines WHERE StockTransferID = " & trID), 0) = 3, _
            "تحويل التفريغ بالباقي في السيارة " & msg, passed, failed, report
    msg = SendStockTransfer(trID)
    If DbValue("SELECT Status FROM StockTransfers WHERE StockTransferID = " & trID) = "SENT" Then
        msg = msg & ReceiveStockTransfer(trID)
    End If
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 0, "تفريغ السيارة " & msg, passed, failed, report
    CheckRd Len(SettleRepDay(dayID)) > 0, "لا يُغلق اليوم قبل تصفية الصندوق", passed, failed, report
    expected = CashBoxBalance(box)
    ' docs/66: the rep has a program user, so his shortages are advances on him
    CurrentDb.Execute "UPDATE SalesReps SET EmployeeID = " & CurrentUserID() & " WHERE SalesRepID = " & rep, dbFailOnError
    advBefore = Nz(DbValue("SELECT AdvanceBalance FROM qryAdvanceTotals WHERE EmployeeID = " & CurrentUserID()), 0)
    msg = PostCashClosing(box, expected - 5, "MAIN", mainBox, expected - 5, "TEST", closingID)
    CheckRd Len(msg) = 0 And expected = Nz(DbValue("SELECT PaidAmount FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0), _
            "تصفية صندوق العهدة بعجز 5 " & msg, passed, failed, report
    CheckRd Nz(DbValue("SELECT COUNT(*) FROM CashVouchers WHERE Category = 'ADVANCE' AND ClosingID = " & closingID & _
                       " AND AdvanceEmployeeID = " & CurrentUserID()), 0) = 1, "عجز صندوق العهدة سلفة على المندوب", _
            passed, failed, report
    msg = SettleRepDay(dayID)
    CheckRd Nz(DbValue("SELECT GoodsShortage FROM RepDays WHERE RepDayID = " & dayID), 0) = 6 And _
            Nz(DbValue("SELECT AdvanceBalance FROM qryAdvanceTotals WHERE EmployeeID = " & CurrentUserID()), 0) = advBefore + 11, _
            "عجز البضاعة والنقدية سلفة على المندوب (6 + 5)", passed, failed, report
    CheckRd Len(msg) = 0 And RepDayStatus(dayID) = "SETTLED" And ActingRepDayID() = 0 And CurrentWarehouseID() <> van, _
            "إغلاق اليوم يوقف العمل باسم المندوب " & msg, passed, failed, report
    CheckRd Nz(DbValue("SELECT LoadedValue FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = 36 And _
            Nz(DbValue("SELECT SalesTotal FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = _
            Nz(DbValue("SELECT TotalAmount FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) And _
            Nz(DbValue("SELECT CountDiffValue FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = -6 And _
            Nz(DbValue("SELECT CashDifference FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = -5, _
            "ملخص اليوم: المحمَّل والمبيعات وفرق الجرد وفرق النقدية", passed, failed, report
    msg = OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) = 0 And Len(DeleteRepDay(dayID)) = 0 And RepDayStatus(dayID) = "", "حذف يوم بلا حركة", _
            passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckRd False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    StopActingAsRep
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات يوم المندوب ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestRepDays"
        TestRepDays = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestRepDays"
    End If
End Function

Private Sub CheckRd(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
