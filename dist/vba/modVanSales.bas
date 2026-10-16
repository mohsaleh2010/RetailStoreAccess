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
        StartActingAsRep = "ÌÊ„ «·„‰œÊ» €Ì— „› ÊÕ."
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
        OpenRepDay = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÌÊ„ «·„‰œÊ»."
        Exit Function
    End If
    If IsNull(SalesRepID) Then
        OpenRepDay = "«Œ — «·„‰œÊ»."
        Exit Function
    End If
    If Nz(DbValue("SELECT IsActive FROM SalesReps WHERE SalesRepID = " & CLng(SalesRepID)), False) = False Then
        OpenRepDay = "«·„‰œÊ» €Ì— „ÊÃÊœ √Ê €Ì— ‰‘ÿ."
        Exit Function
    End If
    van = RepVanID(CLng(SalesRepID))
    box = RepCustodyBoxID(CLng(SalesRepID))
    If van = 0 Then
        OpenRepDay = "·Ì” ··„‰œÊ» „Œ“‰ ”Ì«—… ‰‘ÿ. √‰‘∆Â ›Ì ´«·„Œ«“‰ª »‰Ê⁄ ´”Ì«—… „‰œÊ»ª Ê«Œ — «·„‰œÊ»."
    ElseIf box = 0 Then
        OpenRepDay = "·Ì” ··„‰œÊ» ’‰œÊﬁ ⁄Âœ… ‰‘ÿ. √‰‘∆Â ›Ì ´«·’‰«œÌﬁª »‰Ê⁄ ´⁄Âœ… „‰œÊ»ª Ê«Œ — «·„‰œÊ»."
    ElseIf Not WarehouseIsActive(StoreID) Then
        OpenRepDay = "«Œ — «·„Œ“‰ «·–Ì  ıÕ„Û¯· „‰Â «·”Ì«—…."
    ElseIf CLng(StoreID) = van Then
        OpenRepDay = "«·„Œ“‰ «·–Ì  ıÕ„Û¯· „‰Â «·”Ì«—… €Ì— „Œ“‰ «·”Ì«—…."
    End If
    If Len(OpenRepDay) > 0 Then Exit Function
    openNo = DbValue("SELECT DayNumber FROM RepDays WHERE Status = 'OPEN' AND SalesRepID = " & CLng(SalesRepID))
    If Not IsNull(openNo) Then
        OpenRepDay = "··„‰œÊ» ÌÊ„ „› ÊÕ (" & openNo & "). ”Êˆ¯Â √Ê·«."
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
        DeleteRepDay = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÌÊ„ «·„‰œÊ»."
    ElseIf RepDayStatus(RepDayID) <> "OPEN" Then
        DeleteRepDay = "ÌıÕ–› «·ÌÊ„ «·„› ÊÕ ›ﬁÿ."
    ElseIf DayHasActivity(RepDayID) Then
        DeleteRepDay = "›Ì Â–« «·ÌÊ„ Õ—ﬂ… ( Õ„Ì· √Ê »Ì⁄ √Ê  Õ’Ì·). ”Êˆ¯Â »œ· Õ–›Â."
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
        LoadRepDay = "ÌÊ„ «·„‰œÊ» €Ì— „› ÊÕ."
        Exit Function
    End If
    LoadRepDay = CreateStockTransfer(DayField(RepDayID, "StoreID"), DayField(RepDayID, "WarehouseID"), _
                                     " Õ„Ì· " & DayField(RepDayID, "DayNumber"), NewTransferID)   ' modWarehouses
End Function

Public Function UnloadRepDay(ByVal RepDayID As Long, ByRef NewTransferID As Long) As String
    ' A draft transfer from the van back to the store with everything left in the van.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, ln As DAO.Recordset, inTrans As Boolean
    Dim van As Long, msg As String
    NewTransferID = 0
    If RepDayStatus(RepDayID) <> "OPEN" Then
        UnloadRepDay = "ÌÊ„ «·„‰œÊ» €Ì— „› ÊÕ."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    If Nz(DbValue("SELECT COUNT(*) FROM StockTransfers WHERE Status = 'DRAFT' AND FromWarehouseID = " & van), 0) > 0 Then
        UnloadRepDay = "··”Ì«—…  ÕÊÌ· „”Êœ… ·„ Ìı—”Û·. √—”·Â √Ê «Õ–›Â √Ê·«."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM qryWarehouseStock WHERE StockQty > 0 AND WarehouseID = " & van), 0) = 0 Then
        UnloadRepDay = "«·”Ì«—… ›«—€…: ·« ‘Ì¡ Ìı›—Û¯€."
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateStockTransfer(van, DayField(RepDayID, "StoreID"), " ›—Ì€ " & DayField(RepDayID, "DayNumber"), NewTransferID)
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
    UnloadRepDay = " ⁄–— ≈‰‘«¡  ÕÊÌ· «· ›—Ì€: " & Err.Description & " (" & Err.Number & ")"
    NewTransferID = 0
    If inTrans Then ws.Rollback
End Function

Public Function CountRepDay(ByVal RepDayID As Long, ByRef CountID As Long) As String
    ' The open count of the van, else a new one.
    Dim van As Long
    CountID = 0
    If RepDayStatus(RepDayID) <> "OPEN" Then
        CountRepDay = "ÌÊ„ «·„‰œÊ» €Ì— „› ÊÕ."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    CountID = Nz(DbValue("SELECT Max(StockCountID) FROM StockCounts WHERE Status = 'OPEN' AND WarehouseID = " & van), 0)
    If CountID = 0 Then
        CountRepDay = CreateStockCount(Null, "Ã—œ ”Ì«—… " & DayField(RepDayID, "DayNumber"), CountID, van)   ' modPurchases
    End If
End Function

Public Function SettleRepDay(ByVal RepDayID As Long) As String
    ' The transfers of the van are complete, its count is posted and the custody box is closed since the day opened.
    Dim van As Long, box As Long, countID As Variant, closingID As Variant, since As String
    If Not HasPermission("REP_DAYS") Then
        SettleRepDay = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÌÊ„ «·„‰œÊ»."
        Exit Function
    End If
    If RepDayStatus(RepDayID) <> "OPEN" Then
        SettleRepDay = "ÌÊ„ «·„‰œÊ» €Ì— „› ÊÕ."
        Exit Function
    End If
    van = DayField(RepDayID, "WarehouseID")
    box = DayField(RepDayID, "CashBoxID")
    since = "(SELECT OpenedAt FROM RepDays WHERE RepDayID = " & RepDayID & ")"
    If Nz(DbValue("SELECT COUNT(*) FROM StockTransfers WHERE Status In ('DRAFT', 'SENT') AND (FromWarehouseID = " & van & _
                  " OR ToWarehouseID = " & van & ")"), 0) > 0 Then
        SettleRepDay = "··”Ì«—…  ÕÊÌ·«  ·„  ﬂ „· („”Êœ… √Ê »«·ÿ—Ìﬁ). √—”·Â« Ê«” ·„Â« √Ê «Õ–›Â« √Ê·«."
        Exit Function
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM StockCounts WHERE Status = 'OPEN' AND WarehouseID = " & van), 0) > 0 Then
        SettleRepDay = "Ã—œ «·”Ì«—… „› ÊÕ. √œŒ· «·ﬂ„Ì«  «·›⁄·Ì… Ê—Õ¯·Â √Ê·«."
        Exit Function
    End If
    countID = DbValue("SELECT Max(StockCountID) FROM StockCounts WHERE Status = 'POSTED' AND WarehouseID = " & van & _
                      " AND PostedAt >= " & since)
    If IsNull(countID) Then
        SettleRepDay = "«Ã—œ «·”Ì«—… Ê—Õ¯· «·Ã—œ ﬁ»· ≈€·«ﬁ «·ÌÊ„ (“— ´Ã—œ «·”Ì«—…ª)."
        Exit Function
    End If
    closingID = DbValue("SELECT Max(ClosingID) FROM CashClosings WHERE CashBoxID = " & box & " AND ClosingDate >= " & since)
    If IsNull(closingID) Then
        SettleRepDay = "’›ˆ¯ ’‰œÊﬁ «·⁄Âœ… ﬁ»· ≈€·«ﬁ «·ÌÊ„ (“— ´ ’›Ì… «·’‰œÊﬁª)."
        Exit Function
    End If
    CurrentDb.Execute "UPDATE RepDays SET Status = 'SETTLED', SettledAt = Now(), StockCountID = " & countID & _
                      ", ClosingID = " & closingID & " WHERE RepDayID = " & RepDayID, dbFailOnError
    If ActingRepDayID() = RepDayID Then StopActingAsRep
    LogAction "REP_DAY_SETTLE", "RepDays", CStr(RepDayID), "Count=" & countID & " Closing=" & closingID
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
        "IIf(d.Status = 'OPEN', '„› ÊÕ', ' „   ”ÊÌ Â') FROM (RepDays AS d INNER JOIN [@SalesReps] AS r " & _
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
    Dim rs As DAO.Recordset, rows As String, vanValue As Variant
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM RepDaySummaryQuery WHERE RepDayID = " & RepDayID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
    vanValue = DbValue("SELECT Sum(q.StockQty * p.AverageCost) FROM qryWarehouseStock AS q INNER JOIN Products AS p " & _
                       "ON q.ProductID = p.ProductID WHERE q.WarehouseID = " & rs!WarehouseID)
    rows = SummaryHead("«·»÷«⁄… (»«· ﬂ·›…)") & _
           SummaryItem("«·„Õ„Û¯· ··”Ì«—…", rs!LoadedValue) & SummaryItem("«·„»«⁄", rs!SoldCost) & _
           SummaryItem("«·„— Ã⁄ „‰ «·⁄„·«¡", rs!ReturnedCost) & SummaryItem("«·„›—Û¯€", rs!UnloadedValue) & _
           SummaryItem("›—ﬁ Ã—œ «·”Ì«—… (⁄Ã“ »«·”«·»)", rs!CountDiffValue) & _
           SummaryItem("«·»«ﬁÌ ›Ì «·”Ì«—… «·¬‰", vanValue) & _
           SummaryHead("«·‰ﬁœÌ…") & _
           SummaryItem("›Ê« Ì— «·»Ì⁄ (" & rs!InvoiceCount & ")", rs!SalesTotal) & _
           SummaryItem("«·„ﬁ»Ê÷ ‰ﬁœ« „‰ «·›Ê« Ì—", rs!CashSales) & SummaryItem("«· Õ’Ì· „‰ «·⁄„·«¡", rs!Collections) & _
           SummaryItem("«·„— Ã⁄« ", rs!ReturnsTotal) & SummaryItem("«·„—œÊœ ‰ﬁœ«", rs!CashRefunds) & _
           SummaryItem("—’Ìœ ’‰œÊﬁ «·⁄Âœ… «·¬‰", CashBoxBalance(rs!CashBoxID))
    If Not IsNull(rs!CashCounted) Then
        rows = rows & SummaryItem("«· ’›Ì…: «·„ Êﬁ⁄", rs!CashExpected) & SummaryItem("«· ’›Ì…: «·›⁄·Ì", rs!CashCounted) & _
               SummaryItem("«· ’›Ì…: «·›—ﬁ (⁄Ã“ »«·”«·»)", rs!CashDifference) & _
               SummaryItem("«·„”·Û¯„ ··Œ“Ì‰…", rs!CashHandedOver)
    End If
    rs.Close
    RepDaySummaryRows = rows
End Function

Public Sub RepDayShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, info As String, rs As DAO.Recordset
    id = Nz(frm!cboDay.Value, 0)
    st = RepDayStatus(id)
    If Len(st) = 0 Then
        info = "·« ÌÊÃœ ÌÊ„. «Œ — «·„‰œÊ» À„ ´› Õ ÌÊ„ª."
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
        frm!lblActing.Caption = Tr(" ⁄„· «·¬‰ »«”„ «·„‰œÊ»: «·»Ì⁄ Ê«·„— Ã⁄ Ê«· Õ’Ì· „‰ „Œ“‰ «·”Ì«—… Ê’‰œÊﬁ «·⁄Âœ… " & _
                                   "Õ Ï  ı€·ﬁ Â–Â «·‘«‘….")
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
            If Not AskYesNo("≈€·«ﬁ ÌÊ„ «·„‰œÊ»ø ·« Ìı› Õ »⁄œ ≈€·«ﬁÂ.") Then Exit Sub
            msg = SettleRepDay(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› ÌÊ„ «·„‰œÊ»ø") Then Exit Sub
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
            OpenScreen "frmStockTransfer", 0, newID            ' the lines, then ´sendª
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
    Dim closingID As Long, expected As Currency
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
                      "CurrentQuantity, AverageCost) VALUES ('TEST-RD-P', 'TEST ’‰› «·”Ì«—…', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-RD-P'")
    CurrentDb.Execute "INSERT INTO SalesReps (RepCode, RepName, CommissionRate, CommissionBase, IsActive) " & _
                      "VALUES ('TEST-RD', 'TEST „‰œÊ» «·”Ì«—…', 0, 'SALES', True)", dbFailOnError
    rep = DbValue("SELECT SalesRepID FROM SalesReps WHERE RepCode = 'TEST-RD'")
    CurrentDb.Execute "INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType, SalesRepID, ConfirmReceipt, " & _
                      "IsActive) VALUES ('TEST-VAN', 'TEST ”Ì«—…', 'VAN', " & rep & ", False, True)", dbFailOnError
    van = DbValue("SELECT WarehouseID FROM Warehouses WHERE WarehouseCode = 'TEST-VAN'")
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive = True"), 0)
    mainBox = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True"), 0)
    msg = OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) > 0 And dayID = 0, "·« ÌÊ„ ·„‰œÊ» »·« ’‰œÊﬁ ⁄Âœ…", passed, failed, report
    CurrentDb.Execute "INSERT INTO CashBoxes (BoxName, BoxType, OpeningBalance, IsActive, SalesRepID, ConfirmReceipt) " & _
                      "VALUES ('TEST ⁄Âœ… «·”Ì«—…', 'CUSTODY', 0, True, " & rep & ", False)", dbFailOnError
    box = DbValue("SELECT CashBoxID FROM CashBoxes WHERE BoxName = 'TEST ⁄Âœ… «·”Ì«—…'")
    CheckRd cust > 0 And mainBox > 0, "ÌÊÃœ ⁄„Ì· ÊŒ“Ì‰… —∆Ì”Ì…", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    msg = msg & OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) = 0 And RepDayStatus(dayID) = "OPEN", "› Õ ÌÊ„ «·„‰œÊ» " & msg, passed, failed, report
    CheckRd Len(OpenRepDay(rep, MAIN_WAREHOUSE_ID, lines)) > 0, "ÌÊ„ „› ÊÕ Ê«Õœ ·ﬂ· „‰œÊ»", passed, failed, report

    ' load 6 into the van
    msg = LoadRepDay(dayID, trID)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      trID & ", " & pid & ", 6, 0)", dbFailOnError
    msg = msg & SendStockTransfer(trID)
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 6, " Õ„Ì· 6 ›Ì «·”Ì«—… " & msg, passed, failed, report

    ' a cash sale of 2 entered by the office in the rep's name
    msg = StartActingAsRep(dayID)
    CheckRd Len(msg) = 0 And CurrentWarehouseID() = van And CurrentCashBoxID() = box And SalesRepFor(cust) = rep, _
            "«·⁄„· »«”„ «·„‰œÊ»: „Œ“‰ «·”Ì«—… Ê’‰œÊﬁÂ« Ê«·„‰œÊ»", passed, failed, report
    CurrentDb.Execute "INSERT INTO tmpPOSLines (ProductID, ProductCode, ProductName, Quantity, UnitPrice, LineDiscount, " & _
                      "Available) VALUES (" & pid & ", 'TEST-RD-P', 'TEST ’‰› «·”Ì«—…', 2, 20, 0, 6)", dbFailOnError
    msg = PostSaleFromCart(cust, "CASH", CASH_METHOD_ID, 0, Null, "TEST", invID)
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 4 And _
            Nz(DbValue("SELECT CashBoxID FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) = box And _
            Nz(DbValue("SELECT SalesRepID FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) = rep, _
            "«·»Ì⁄ „‰ «·”Ì«—… ‰ﬁœ« ·’‰œÊﬁ «·⁄Âœ… " & msg, passed, failed, report

    ' the van count finds 3: one short
    msg = CountRepDay(dayID, countID)
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = 3 WHERE StockCountID = " & countID & _
                      " AND ProductID = " & pid, dbFailOnError
    CheckRd Len(SettleRepDay(dayID)) > 0, "·« Ìı€·ﬁ «·ÌÊ„ ÊÃ—œ «·”Ì«—… „› ÊÕ", passed, failed, report
    msg = msg & PostStockCount(countID, False, lines, netValue)
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 3 And netValue = -6, "Ã—œ «·”Ì«—…: ⁄Ã“ ’‰› Ê«Õœ " & msg, _
            passed, failed, report

    ' unload the rest, then close the custody box
    msg = UnloadRepDay(dayID, trID)
    CheckRd Len(msg) = 0 And Nz(DbValue("SELECT Sum(Quantity) FROM StockTransferLines WHERE StockTransferID = " & trID), 0) = 3, _
            " ÕÊÌ· «· ›—Ì€ »«·»«ﬁÌ ›Ì «·”Ì«—… " & msg, passed, failed, report
    msg = SendStockTransfer(trID)
    If DbValue("SELECT Status FROM StockTransfers WHERE StockTransferID = " & trID) = "SENT" Then
        msg = msg & ReceiveStockTransfer(trID)
    End If
    CheckRd Len(msg) = 0 And WarehouseStock(pid, van) = 0, " ›—Ì€ «·”Ì«—… " & msg, passed, failed, report
    CheckRd Len(SettleRepDay(dayID)) > 0, "·« Ìı€·ﬁ «·ÌÊ„ ﬁ»·  ’›Ì… «·’‰œÊﬁ", passed, failed, report
    expected = CashBoxBalance(box)
    msg = PostCashClosing(box, expected - 5, "MAIN", mainBox, expected - 5, "TEST", closingID)
    CheckRd Len(msg) = 0 And expected = Nz(DbValue("SELECT PaidAmount FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0), _
            " ’›Ì… ’‰œÊﬁ «·⁄Âœ… »⁄Ã“ 5 " & msg, passed, failed, report
    msg = SettleRepDay(dayID)
    CheckRd Len(msg) = 0 And RepDayStatus(dayID) = "SETTLED" And ActingRepDayID() = 0 And CurrentWarehouseID() <> van, _
            "≈€·«ﬁ «·ÌÊ„ ÌÊﬁ› «·⁄„· »«”„ «·„‰œÊ» " & msg, passed, failed, report
    CheckRd Nz(DbValue("SELECT LoadedValue FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = 36 And _
            Nz(DbValue("SELECT SalesTotal FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = _
            Nz(DbValue("SELECT TotalAmount FROM SalesInvoices WHERE SalesInvoiceID = " & invID), 0) And _
            Nz(DbValue("SELECT CountDiffValue FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = -6 And _
            Nz(DbValue("SELECT CashDifference FROM RepDaySummaryQuery WHERE RepDayID = " & dayID), 0) = -5, _
            "„·Œ’ «·ÌÊ„: «·„Õ„Û¯· Ê«·„»Ì⁄«  Ê›—ﬁ «·Ã—œ Ê›—ﬁ «·‰ﬁœÌ…", passed, failed, report
    msg = OpenRepDay(rep, MAIN_WAREHOUSE_ID, dayID)
    CheckRd Len(msg) = 0 And Len(DeleteRepDay(dayID)) = 0 And RepDayStatus(dayID) = "", "Õ–› ÌÊ„ »·« Õ—ﬂ…", _
            passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckRd False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    StopActingAsRep
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ÌÊ„ «·„‰œÊ» ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestRepDays"
        TestRepDays = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestRepDays"
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
