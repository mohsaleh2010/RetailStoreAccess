Attribute VB_Name = "modWarehouses"
'==============================================================================
' modWarehouses  -  Retail Store Management System (warehouses, docs/57)
'
' Every stock move carries its warehouse (InventoryTransactions.WarehouseID), and so do the
' documents that move stock (sales, returns, purchases, stock counts, opening stock). The balance
' of a product in a warehouse is the sum of its moves there; Products.CurrentQuantity stays the
' total of every warehouse and the average cost stays one per product.
'
'   CurrentWarehouseID   warehouse of the logged-in user (Employees.WarehouseID), else the main one
'   WarehouseStock       quantity of a product in a warehouse (the sum of its moves there)
'   AvailableQty         what can leave a warehouse; while every move is in the main warehouse
'                        it is Products.CurrentQuantity, exactly as before warehouses
'   WarehouseProblem     "" when the warehouse can be used by a document
'   Transfers (docs/58)  CreateStockTransfer (draft) -> SendStockTransfer: out of the sending warehouse at
'                        the average cost, into the receiving one, or in transit (account 1410) while the
'                        receiving warehouse confirms its receipts -> ReceiveStockTransfer; UnpostStockTransfer
'                        takes back a transfer in transit, DeleteStockTransfer a draft
'   Reports (docs/60)    the item card ItemCardQuery (quantity and value, one warehouse or all), the
'                        valuation of each warehouse, the transfers in transit, the posted count differences;
'                        each warehouse may have its own open stock count (CreateStockCount); TestStockReports
' Screens: frmWarehouses, frmStockTransfer (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Public Const MAIN_WAREHOUSE_ID As Long = 1
Public Const TT_TRANSFER_OUT As Long = 9
Public Const TT_TRANSFER_IN As Long = 10

Public Function CurrentWarehouseID() As Long
    Dim v As Variant
    If ActingRepDayID() > 0 Then                    ' the office works in the name of a rep (modVanSales)
        CurrentWarehouseID = ActingWarehouseID()
        Exit Function
    End If
    v = DbValue("SELECT e.WarehouseID FROM Employees AS e INNER JOIN Warehouses AS w ON e.WarehouseID = w.WarehouseID " & _
                "WHERE e.EmployeeID = " & CurrentUserID() & " AND w.IsActive = True")
    If IsNull(v) Then v = DbValue("SELECT Min(WarehouseID) FROM Warehouses WHERE IsActive = True AND WarehouseType = 'MAIN'")
    If IsNull(v) Then v = DbValue("SELECT Min(WarehouseID) FROM Warehouses WHERE IsActive = True")
    CurrentWarehouseID = Nz(v, MAIN_WAREHOUSE_ID)
End Function

Public Function WarehouseIsActive(ByVal WarehouseID As Variant) As Boolean
    If IsNull(WarehouseID) Then Exit Function
    WarehouseIsActive = Nz(DbValue("SELECT COUNT(*) FROM Warehouses WHERE IsActive = True AND WarehouseID = " & _
                                   CLng(WarehouseID)), 0) > 0
End Function

Public Function WarehouseProblem(ByVal WarehouseID As Variant) As String
    If Not WarehouseIsActive(WarehouseID) Then WarehouseProblem = "«Œ — «·„Œ“‰ („Œ“‰ ‰‘ÿ)."
End Function

Public Function WarehouseName(ByVal WarehouseID As Variant) As String
    If IsNull(WarehouseID) Then Exit Function
    WarehouseName = Nz(DbValue(Tr("SELECT w.WarehouseName FROM [@Warehouses] AS w WHERE w.WarehouseID = " & _
                                  CLng(WarehouseID))), "")
End Function

Public Function OnlyMainWarehouse() As Boolean
    ' True while no stock move was made outside the main warehouse.
    OnlyMainWarehouse = Nz(DbValue("SELECT COUNT(*) FROM InventoryTransactions WHERE WarehouseID <> " & _
                                   MAIN_WAREHOUSE_ID), 0) = 0
End Function

Public Function WarehouseStock(ByVal ProductID As Long, ByVal WarehouseID As Long) As Currency
    WarehouseStock = Nz(DbValue("SELECT Sum(Quantity) FROM InventoryTransactions WHERE ProductID = " & ProductID & _
                                " AND WarehouseID = " & WarehouseID), 0)
End Function

Public Function AvailableQty(ByVal ProductID As Long, Optional ByVal WarehouseID As Long = 0) As Currency
    If WarehouseID = 0 Then WarehouseID = CurrentWarehouseID()
    If WarehouseID = MAIN_WAREHOUSE_ID And OnlyMainWarehouse() Then
        AvailableQty = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & ProductID), 0)
    Else
        AvailableQty = WarehouseStock(ProductID, WarehouseID)
    End If
End Function

Public Function WarehouseRows() As String
    ' The active warehouses for a combo box (through Tr: an English front-end shows the English names).
    WarehouseRows = Tr("SELECT w.WarehouseID, w.WarehouseName FROM [@Warehouses] AS w WHERE w.IsActive = True " & _
                       "ORDER BY w.WarehouseID")
End Function

'==============================================================================
' Transfers between warehouses (docs/58)
'==============================================================================
Public Function WarehouseConfirmsReceipt(ByVal WarehouseID As Variant) As Boolean
    If IsNull(WarehouseID) Then Exit Function
    WarehouseConfirmsReceipt = Nz(DbValue("SELECT ConfirmReceipt FROM Warehouses WHERE WarehouseID = " & _
                                          CLng(WarehouseID)), False)
End Function

Public Function CanReceiveIn(ByVal WarehouseID As Long) As Boolean
    ' Its keeper, a user whose warehouse it is, or the administrator; anyone with the permission when it has no keeper.
    Dim keeper As Variant, uid As Long
    If Not HasPermission("STOCK_TRANSFER") Then Exit Function
    uid = CurrentUserID()
    keeper = DbValue("SELECT KeeperID FROM Warehouses WHERE WarehouseID = " & WarehouseID)
    If IsNull(keeper) Then
        CanReceiveIn = True
    ElseIf CLng(keeper) = uid Or CurrentWarehouseID() = WarehouseID Then
        CanReceiveIn = True
    Else
        CanReceiveIn = Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE EmployeeID = " & uid & _
                                  " AND (RoleID = " & ADMIN_ROLE_ID & " OR IsDeveloper = True)"), 0) > 0
    End If
End Function

Private Function TransferStatus(ByVal StockTransferID As Long) As String
    TransferStatus = Nz(DbValue("SELECT Status FROM StockTransfers WHERE StockTransferID = " & StockTransferID), "")
End Function

Public Function CreateStockTransfer(ByVal FromWarehouseID As Variant, ByVal ToWarehouseID As Variant, _
                                    ByVal Notes As String, ByRef NewTransferID As Long) As String
    Dim rs As DAO.Recordset
    NewTransferID = 0
    If Not HasPermission("STOCK_TRANSFER") Then
        CreateStockTransfer = "·Ì”  ·œÌﬂ ’·«ÕÌ… «· ÕÊÌ· »Ì‰ «·„Œ«“‰."
    ElseIf Not WarehouseIsActive(FromWarehouseID) Then
        CreateStockTransfer = "«Œ — «·„Œ“‰ «·„—”ˆ·."
    ElseIf Not WarehouseIsActive(ToWarehouseID) Then
        CreateStockTransfer = "«Œ — «·„Œ“‰ «·„” ·„."
    ElseIf CLng(FromWarehouseID) = CLng(ToWarehouseID) Then
        CreateStockTransfer = "«·„Œ“‰ «·„” ·„ €Ì— «·„—”ˆ·."
    End If
    If Len(CreateStockTransfer) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("StockTransfers", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!TransferNumber = NextNumber("STOCK_TRANSFER")
    rs!TransferDate = Now
    rs!FromWarehouseID = CLng(FromWarehouseID)
    rs!ToWarehouseID = CLng(ToWarehouseID)
    rs!Status = "DRAFT"
    rs!InTransit = False
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewTransferID = rs!StockTransferID
    rs.Close
    LogAction "STOCK_TRANSFER_NEW", "StockTransfers", CStr(NewTransferID)
End Function

Public Function StockTransferLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant) As String
    If IsNull(ProductID) Then
        StockTransferLineProblem = "«Œ — «·’‰›."
    ElseIf Not ProductTracksStock(CLng(ProductID)) Then
        StockTransferLineProblem = "Â–« «·’‰› ·« Ì «»⁄ «·„Œ“Ê‰° ›·« ÌıÕÊÛ¯·."
    ElseIf Nz(Qty, 0) <= 0 Then
        StockTransferLineProblem = "«·ﬂ„Ì… √ﬂ»— „‰ ’›—."
    End If
End Function

Public Function SendStockTransfer(ByVal StockTransferID As Long) As String
    ' Out of the sending warehouse at the average cost; into the receiving one now, or later when it confirms.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, msg As String
    Dim fromWh As Long, toWh As Long, transferNo As String, transit As Boolean, n As Long, i As Long
    Dim ids() As Long, qtys() As Currency, cost As Currency
    If Not HasPermission("STOCK_TRANSFER") Then
        SendStockTransfer = "·Ì”  ·œÌﬂ ’·«ÕÌ… «· ÕÊÌ· »Ì‰ «·„Œ«“‰."
        Exit Function
    End If
    If TransferStatus(StockTransferID) <> "DRAFT" Then
        SendStockTransfer = "«· ÕÊÌ· €Ì— „ÊÃÊœ √Ê  „  —ÕÌ·Â."
        Exit Function
    End If
    fromWh = DbValue("SELECT FromWarehouseID FROM StockTransfers WHERE StockTransferID = " & StockTransferID)
    toWh = DbValue("SELECT ToWarehouseID FROM StockTransfers WHERE StockTransferID = " & StockTransferID)
    transferNo = DbValue("SELECT TransferNumber FROM StockTransfers WHERE StockTransferID = " & StockTransferID)
    msg = WarehouseProblem(fromWh)
    If Len(msg) = 0 Then msg = WarehouseProblem(toWh)
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(Date)
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity FROM StockTransferLines WHERE StockTransferID = " & _
                                         StockTransferID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            If Len(msg) = 0 Then msg = StockTransferLineProblem(rs!ProductID, rs!Quantity)
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) = 0 And n = 0 Then msg = "√÷› «·√’‰«› Êﬂ„Ì« Â« √Ê·«."
    End If
    If Len(msg) = 0 Then msg = CheckStockAvailable(ids, qtys, n, "«· ÕÊÌ· »Ì‰ «·„Œ«“‰", fromWh)
    If Len(msg) > 0 Then
        SendStockTransfer = msg
        Exit Function
    End If

    On Error GoTo EH
    transit = WarehouseConfirmsReceipt(toWh)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT StockTransferLineID, ProductID, Quantity, UnitCost FROM StockTransferLines " & _
                              "WHERE StockTransferID = " & StockTransferID & " ORDER BY StockTransferLineID", dbOpenDynaset)
    Do Until rs.EOF
        cost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & rs!ProductID), 0)
        rs.Edit
        rs!UnitCost = cost
        rs.Update
        ApplyStockMovement db, rs!ProductID, -rs!Quantity, TT_TRANSFER_OUT, cost, "TRANSFER", StockTransferID, _
                           transferNo, "", False, fromWh
        If Not transit Then
            ApplyStockMovement db, rs!ProductID, rs!Quantity, TT_TRANSFER_IN, cost, "TRANSFER", StockTransferID, _
                               transferNo, "", False, toWh
        End If
        rs.MoveNext
    Loop
    rs.Close
    If transit Then
        db.Execute "UPDATE StockTransfers SET TransferDate = Now(), Status = 'SENT', InTransit = True " & _
                   "WHERE StockTransferID = " & StockTransferID, dbFailOnError
    Else
        db.Execute "UPDATE StockTransfers SET TransferDate = Now(), Status = 'RECEIVED', InTransit = False, " & _
                   "ReceivedAt = Now(), ReceivedByID = " & CurrentUserID() & " WHERE StockTransferID = " & _
                   StockTransferID, dbFailOnError
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_TRANSFER_SEND", "StockTransfers", transferNo, "From=" & fromWh & " To=" & toWh & " Transit=" & transit
    Exit Function
EH:
    SendStockTransfer = " ⁄–—  —ÕÌ· «· ÕÊÌ·: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function ReceiveStockTransfer(ByVal StockTransferID As Long) As String
    ' The receiving warehouse confirms the goods arrived: they enter it now, at the cost they left with.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim toWh As Long, transferNo As String, before As Collection
    If TransferStatus(StockTransferID) <> "SENT" Then
        ReceiveStockTransfer = "Â–« «· ÕÊÌ· ·Ì” »«·ÿ—Ìﬁ."
        Exit Function
    End If
    toWh = DbValue("SELECT ToWarehouseID FROM StockTransfers WHERE StockTransferID = " & StockTransferID)
    If Not CanReceiveIn(toWh) Then
        ReceiveStockTransfer = "Ì” ·„ «· ÕÊÌ· √„Ì‰ " & WarehouseName(toWh) & "."
        Exit Function
    End If
    ReceiveStockTransfer = ClosedPeriodProblem(Date)
    If Len(ReceiveStockTransfer) > 0 Then Exit Function
    transferNo = DbValue("SELECT TransferNumber FROM StockTransfers WHERE StockTransferID = " & StockTransferID)

    On Error GoTo EH
    Set before = AuditSnapshot("StockTransfers", "StockTransferID", StockTransferID)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT ProductID, Quantity, UnitCost FROM StockTransferLines WHERE StockTransferID = " & _
                              StockTransferID & " ORDER BY StockTransferLineID", dbOpenSnapshot)
    Do Until rs.EOF
        ApplyStockMovement db, rs!ProductID, rs!Quantity, TT_TRANSFER_IN, rs!UnitCost, "TRANSFER", StockTransferID, _
                           transferNo, "", False, toWh
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE StockTransfers SET Status = 'RECEIVED', ReceivedAt = Now(), ReceivedByID = " & CurrentUserID() & _
               " WHERE StockTransferID = " & StockTransferID, dbFailOnError
    ws.CommitTrans
    inTrans = False
    AuditEdited "STOCK_TRANSFER_RECEIVE", "StockTransfers", "StockTransferID", StockTransferID, before
    Exit Function
EH:
    ReceiveStockTransfer = " ⁄–— «” ·«„ «· ÕÊÌ·: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function UnpostStockTransfer(ByVal StockTransferID As Long) As String
    ' A transfer still in transit goes back to draft: its moves out are deleted and the goods are back in
    ' the sending warehouse. A received transfer stays (move the goods back with a new transfer).
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, before As Collection
    If Not HasPermission("STOCK_TRANSFER") Then
        UnpostStockTransfer = "·Ì”  ·œÌﬂ ’·«ÕÌ… «· ÕÊÌ· »Ì‰ «·„Œ«“‰."
        Exit Function
    End If
    If TransferStatus(StockTransferID) <> "SENT" Then
        UnpostStockTransfer = "Ìı·€Ï  —ÕÌ· «· ÕÊÌ· »«·ÿ—Ìﬁ ›ﬁÿ. «· ÕÊÌ· «·„” ·„ Ìı⁄«œ » ÕÊÌ· ÃœÌœ."
        Exit Function
    End If
    UnpostStockTransfer = ClosedPeriodProblem(DbValue("SELECT TransferDate FROM StockTransfers WHERE StockTransferID = " & _
                                                      StockTransferID))
    If Len(UnpostStockTransfer) > 0 Then Exit Function

    On Error GoTo EH
    Set before = AuditSnapshot("StockTransfers", "StockTransferID", StockTransferID)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT ProductID, Quantity FROM StockTransferLines WHERE StockTransferID = " & _
                              StockTransferID, dbOpenSnapshot)
    Do Until rs.EOF
        db.Execute "UPDATE Products SET CurrentQuantity = CurrentQuantity + " & Str$(rs!Quantity) & _
                   ", UpdatedAt = Now() WHERE ProductID = " & rs!ProductID, dbFailOnError
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "DELETE FROM InventoryTransactions WHERE ReferenceType = 'TRANSFER' AND ReferenceID = " & StockTransferID, _
               dbFailOnError
    db.Execute "UPDATE StockTransfers SET Status = 'DRAFT', InTransit = False WHERE StockTransferID = " & StockTransferID, _
               dbFailOnError
    ws.CommitTrans
    inTrans = False
    AuditEdited "STOCK_TRANSFER_UNPOST", "StockTransfers", "StockTransferID", StockTransferID, before
    Exit Function
EH:
    UnpostStockTransfer = " ⁄–— ≈·€«¡ «· —ÕÌ·: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function DeleteStockTransfer(ByVal StockTransferID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("STOCK_TRANSFER") Then
        DeleteStockTransfer = "·Ì”  ·œÌﬂ ’·«ÕÌ… «· ÕÊÌ· »Ì‰ «·„Œ«“‰."
        Exit Function
    End If
    If TransferStatus(StockTransferID) <> "DRAFT" Then
        DeleteStockTransfer = " ıÕ–› «·„”Êœ… ›ﬁÿ. √·€ˆ  —ÕÌ· «· ÕÊÌ· »«·ÿ—Ìﬁ √Ê·«."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("StockTransfers", "StockTransferID", StockTransferID)
    CurrentDb.Execute "DELETE FROM StockTransferLines WHERE StockTransferID = " & StockTransferID, dbFailOnError
    CurrentDb.Execute "DELETE FROM StockTransfers WHERE StockTransferID = " & StockTransferID, dbFailOnError
    AuditDeleted "STOCK_TRANSFER_DELETE", "StockTransfers", StockTransferID, auditBefore
End Function

'==============================================================================
' frmStockTransfer (OpenArgs = a transfer to show) and its lines frmStockTransferLines
'==============================================================================
Public Sub TransferLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboFrom.RowSource = WarehouseRows()
    frm!cboTo.RowSource = WarehouseRows()
    frm!cboFrom.Value = CurrentWarehouseID()
    TransferListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(StockTransferID) FROM StockTransfers")
    frm!cboTransfer.Value = id
    TransferShow frm
End Sub

Private Sub TransferListRefresh(ByVal frm As Access.Form)
    frm!cboTransfer.RowSource = Tr("SELECT t.StockTransferID, t.TransferNumber, Format(t.TransferDate, 'yyyy/mm/dd'), " & _
        "f.WarehouseName, w.WarehouseName, IIf(t.Status = 'DRAFT', '„”Êœ…', IIf(t.Status = 'SENT', '»«·ÿ—Ìﬁ', " & _
        "'„ı” ·Û„')) FROM (StockTransfers AS t INNER JOIN [@Warehouses] AS f ON t.FromWarehouseID = f.WarehouseID) " & _
        "INNER JOIN [@Warehouses] AS w ON t.ToWarehouseID = w.WarehouseID ORDER BY t.StockTransferID DESC")
    frm!cboTransfer.Requery
End Sub

Public Sub TransferShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, rs As DAO.Recordset, info As String
    id = Nz(frm!cboTransfer.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM StockTransferLines WHERE StockTransferID = " & id & _
                                     " ORDER BY StockTransferLineID"
    st = TransferStatus(id)
    If Len(st) = 0 Then
        info = "·« ÌÊÃœ  ÕÊÌ·. «Œ — «·„Œ“‰Ì‰ À„ ´ ÕÊÌ· ÃœÌœª."
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT TransferNumber, TransferDate, FromWarehouseID, ToWarehouseID, " & _
                                         "ReceivedAt FROM StockTransfers WHERE StockTransferID = " & id, dbOpenSnapshot)
        info = rs!TransferNumber & "    " & WarehouseName(rs!FromWarehouseID) & "  ->  " & WarehouseName(rs!ToWarehouseID) & _
               "    " & IIf(st = "DRAFT", "„”Êœ…: √÷› «·√’‰«› À„ ´ —ÕÌ·ª", IIf(st = "SENT", "»«·ÿ—Ìﬁ „‰– " & _
               GDate(rs!TransferDate) & ": Ì‰ Ÿ— «·«” ·«„", "„ı” ·Û„ ›Ì " & GDate(Nz(rs!ReceivedAt, rs!TransferDate))))
        rs.Close
        info = info & "    «·ﬁÌ„…: " & Format$(Nz(DbValue("SELECT TransferValue FROM qryStockTransferValue " & _
                                                         "WHERE StockTransferID = " & id), 0), "#,##0.00")
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    SafeFocus frm!cboTransfer               ' a focused control cannot be disabled
    frm!btnSend.Enabled = (st = "DRAFT")
    frm!btnDelete.Enabled = (st = "DRAFT")
    frm!btnReceive.Enabled = (st = "SENT")
    frm!btnUnpost.Enabled = (st = "SENT")
End Sub

Public Sub TransferNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateStockTransfer(frm!cboFrom.Value, frm!cboTo.Value, Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    TransferListRefresh frm
    frm!cboTransfer.Value = id
    TransferShow frm
End Sub

Public Sub TransferAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String
    id = Nz(frm!cboTransfer.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "SEND"
            If Not AskYesNo(" —ÕÌ· «· ÕÊÌ·ø  ıŒ’„ «·√’‰«› „‰ «·„Œ“‰ «·„—”ˆ· »„ Ê”ÿ  ﬂ·› Â«.") Then Exit Sub
            msg = SendStockTransfer(id)
        Case "RECEIVE"
            msg = ReceiveStockTransfer(id)
        Case "UNPOST"
            If Not AskYesNo("≈·€«¡  —ÕÌ· «· ÕÊÌ· Ê≈⁄«œ… «·√’‰«› ··„Œ“‰ «·„—”ˆ·ø") Then Exit Sub
            msg = UnpostStockTransfer(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› „”Êœ… «· ÕÊÌ·ø") Then Exit Sub
            msg = DeleteStockTransfer(id)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    TransferListRefresh frm
    If Action = "DELETE" Then frm!cboTransfer.Value = DbValue("SELECT Max(StockTransferID) FROM StockTransfers")
    TransferShow frm
End Sub

Public Sub TransferLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String
    id = Nz(frm.Parent!cboTransfer.Value, 0)
    If TransferStatus(id) <> "DRAFT" Then
        msg = " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
    Else
        msg = StockTransferLineProblem(frm!ProductID.Value, frm!Quantity.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!StockTransferID.Value = id
    frm!UnitCost.Value = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & CLng(frm!ProductID.Value)), 0)
    AuditFormBefore frm, "StockTransferLines", "StockTransferLineID"
End Sub

Public Sub TransferLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    TransferShow frm.Parent
End Sub

Public Sub TransferRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!StockTransferLineID.Value
    If IsNull(id) Then Exit Sub
    If TransferStatus(Nz(frm.Parent!cboTransfer.Value, 0)) <> "DRAFT" Then
        ShowWarning " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("StockTransferLines", "StockTransferLineID", id)
    CurrentDb.Execute "DELETE FROM StockTransferLines WHERE StockTransferLineID = " & CLng(id), dbFailOnError
    AuditDeleted "STOCK_TRANSFER_LINE_DELETE", "StockTransferLines", id, auditBefore
    frm.Requery
    TransferShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestWarehouses() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, wh2 As Long, pid As Long, refNo As String, msg As String
    Dim passed As Long, failed As Long, report As String, countID As Long, lines As Long, netValue As Currency
    Dim tid As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestWarehouses  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckWh WarehouseIsActive(MAIN_WAREHOUSE_ID), "«·„Œ“‰ «·—∆Ì”Ì „ÊÃÊœ Ê‰‘ÿ", passed, failed, report
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType) VALUES " & _
                      "('TEST-WH', 'TEST „Œ“‰ ›—⁄', 'BRANCH')", dbFailOnError
    wh2 = DbValue("SELECT Max(WarehouseID) FROM Warehouses")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-WH-P', 'TEST ’‰› «·„Œ«“‰', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-WH-P'")

    msg = PostManualStock(pid, TT_STOCK_IN, 10, 5, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckWh Len(msg) = 0, "≈÷«›… 10 ··„Œ“‰ «·—∆Ì”Ì " & msg, passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 4, 5, "TEST", refNo, wh2)
    CheckWh Len(msg) = 0, "≈÷«›… 4 ·„Œ“‰ «·›—⁄ " & msg, passed, failed, report
    CheckWh WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10 And WarehouseStock(pid, wh2) = 4, _
            "—’Ìœ ﬂ· „Œ“‰ = Õ—ﬂ« Â ›ÌÂ", passed, failed, report
    CheckWh Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 14, _
            "—’Ìœ «·’‰› = „Ã„Ê⁄ «·„Œ«“‰", passed, failed, report
    CheckWh AvailableQty(pid, wh2) = 4 And AvailableQty(pid, MAIN_WAREHOUSE_ID) = 10, _
            "«·„ «Õ ··’—› „‰ ﬂ· „Œ“‰", passed, failed, report
    CheckWh Nz(DbValue("SELECT Sum(StockQty) FROM StockByWarehouseQuery WHERE ProductID = " & pid & _
                       " AND WarehouseID = " & wh2), 0) = 4, " ﬁ—Ì— √—’œ… «·„Œ«“‰", passed, failed, report

    ' the user's warehouse
    CurrentDb.Execute "UPDATE Employees SET WarehouseID = " & wh2 & " WHERE EmployeeID = " & CurrentUserID(), dbFailOnError
    CheckWh CurrentWarehouseID() = wh2, "„Œ“‰ «·„” Œœ„", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_OUT, 1, Null, "TEST  «·›", refNo)
    CheckWh Len(msg) = 0 And WarehouseStock(pid, wh2) = 3 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10, _
            "«·Œ’„ „‰ „Œ“‰ «·„” Œœ„ " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE Employees SET WarehouseID = Null WHERE EmployeeID = " & CurrentUserID(), dbFailOnError

    ' a stock count of one warehouse compares with that warehouse only
    CurrentDb.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE Status = 'OPEN'", dbFailOnError
    msg = CreateStockCount(Null, "TEST", countID, wh2)
    CheckWh Len(msg) = 0 And Nz(DbValue("SELECT SystemQuantity FROM StockCountDetails WHERE StockCountID = " & _
            countID & " AND ProductID = " & pid), -1) = 3, "«·Ã—œ Ìﬁ—√ —’Ìœ „Œ“‰Â " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = 2 WHERE StockCountID = " & countID & _
                      " AND ProductID = " & pid, dbFailOnError
    CurrentDb.Execute "UPDATE StockCountDetails SET ActualQuantity = SystemQuantity WHERE StockCountID = " & countID & _
                      " AND ProductID <> " & pid, dbFailOnError
    msg = PostStockCount(countID, False, lines, netValue)
    CheckWh Len(msg) = 0 And WarehouseStock(pid, wh2) = 2 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10, _
            " ”ÊÌ… «·Ã—œ ›Ì „Œ“‰Â ›ﬁÿ " & msg, passed, failed, report

    ' transfers (docs/58): at once, then through transit
    msg = CreateStockTransfer(MAIN_WAREHOUSE_ID, wh2, "TEST", tid)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      tid & ", " & pid & ", 3, 0)", dbFailOnError
    msg = msg & SendStockTransfer(tid)
    CheckWh Len(msg) = 0 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 7 And WarehouseStock(pid, wh2) = 5 And _
            Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 12, _
            "«· ÕÊÌ· ÌŒ’„ „‰ «·„—”ˆ· ÊÌ÷Ì› ··„” ·„ " & msg, passed, failed, report
    CheckWh TransferStatus(tid) = "RECEIVED", "„Œ“‰ »·«  √ﬂÌœ Ì” ·„ ›Ê—«", passed, failed, report
    CurrentDb.Execute "UPDATE Warehouses SET ConfirmReceipt = True WHERE WarehouseID = " & wh2, dbFailOnError
    msg = CreateStockTransfer(MAIN_WAREHOUSE_ID, wh2, "TEST", tid)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      tid & ", " & pid & ", 2, 0)", dbFailOnError
    msg = msg & SendStockTransfer(tid)
    CheckWh Len(msg) = 0 And TransferStatus(tid) = "SENT" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 5 And _
            WarehouseStock(pid, wh2) = 5, "«· ÕÊÌ· »«·ÿ—Ìﬁ ·« ÌœŒ· «·„” ·„ " & msg, passed, failed, report
    msg = UnpostStockTransfer(tid)
    CheckWh Len(msg) = 0 And TransferStatus(tid) = "DRAFT" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 7 And _
            Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 12, _
            "≈·€«¡  —ÕÌ· «· ÕÊÌ· »«·ÿ—Ìﬁ Ì⁄Ìœ «·√’‰«› " & msg, passed, failed, report
    msg = SendStockTransfer(tid) & ReceiveStockTransfer(tid)
    CheckWh Len(msg) = 0 And TransferStatus(tid) = "RECEIVED" And WarehouseStock(pid, wh2) = 7, _
            " √ﬂÌœ «·«” ·«„ ÌœŒ· «·„” ·„ " & msg, passed, failed, report
    CheckWh Len(UnpostStockTransfer(tid)) > 0 And Len(DeleteStockTransfer(tid)) > 0, _
            "«· ÕÊÌ· «·„” ·„ ·« Ìı·€Ï Ê·« ÌıÕ–›", passed, failed, report
    msg = CreateStockTransfer(MAIN_WAREHOUSE_ID, wh2, "TEST", tid)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      tid & ", " & pid & ", 999, 0)", dbFailOnError
    CurrentDb.Execute "UPDATE Settings SET AllowNegativeStock = False", dbFailOnError
    g_AutoAnswer = False
    CheckWh Len(SendStockTransfer(tid)) > 0 And TransferStatus(tid) = "DRAFT", "·« ÌıÕÊÛ¯· √ﬂÀ— „‰ —’Ìœ «·„—”ˆ·", _
            passed, failed, report
    CheckWh Len(DeleteStockTransfer(tid)) = 0 And TransferStatus(tid) = "", "Õ–› «·„”Êœ…", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckWh False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·„Œ«“‰ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestWarehouses"
        TestWarehouses = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestWarehouses"
    End If
End Function

Public Function TestStockReports() As Boolean
    ' The item card and the warehouse reports (docs/60, plan 2D), and one open count per warehouse.
    Dim ws As DAO.Workspace, inTrans As Boolean, wh2 As Long, pid As Long, refNo As String, msg As String
    Dim passed As Long, failed As Long, report As String, countID As Long, tid As Long, branchName As String
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestStockReports  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO Warehouses (WarehouseCode, WarehouseName, WarehouseType) VALUES " & _
                      "('TEST-WH', 'TEST „Œ“‰ ›—⁄', 'BRANCH')", dbFailOnError
    wh2 = DbValue("SELECT Max(WarehouseID) FROM Warehouses")
    branchName = Nz(DbValue(Tr("SELECT w.WarehouseName FROM [@Warehouses] AS w WHERE w.WarehouseID = " & wh2)), "")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-CARD-P', 'TEST ﬂ«—  «·’‰›', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 0, 0)", dbFailOnError
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-CARD-P'")
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 5, "TEST", refNo, MAIN_WAREHOUSE_ID) & _
          PostManualStock(pid, TT_STOCK_IN, 4, 5, "TEST", refNo, wh2)
    CheckWh Len(msg) = 0, "≈÷«›… «·’‰› ··„Œ“‰Ì‰ " & msg, passed, failed, report

    ' the item card: every warehouse, then one, then the opening line of a later period
    SetPeriod Date, Date
    SetQueryParam "ProductID", pid
    SetQueryParam "WarehouseID", 0
    CheckWh Nz(DbValue("SELECT Sum(NetQty) FROM ItemCardQuery"), 0) = 14 And _
            Nz(DbValue("SELECT Sum(NetValue) FROM ItemCardQuery"), 0) = 70 And _
            Nz(DbValue("SELECT Count(*) FROM ItemCardQuery WHERE SortKey = 1"), 0) = 2, _
            "ﬂ«—  «·’‰› ·ﬂ· «·„Œ«“‰: «·ﬂ„Ì… Ê«·ﬁÌ„…", passed, failed, report
    SetQueryParam "WarehouseID", wh2
    CheckWh Nz(DbValue("SELECT Sum(NetQty) FROM ItemCardQuery"), 0) = 4 And _
            Nz(DbValue("SELECT Sum(ValueIn) FROM ItemCardQuery"), 0) = 20, _
            "ﬂ«—  «·’‰› ·„Œ“‰ Ê«Õœ", passed, failed, report
    SetPeriod Date + 1, Date + 1
    SetQueryParam "WarehouseID", 0
    CheckWh Nz(DbValue("SELECT NetQty FROM ItemCardQuery WHERE SortKey = 0"), 0) = 14 And _
            Nz(DbValue("SELECT NetValue FROM ItemCardQuery WHERE SortKey = 0"), 0) = 70 And _
            Nz(DbValue("SELECT Count(*) FROM ItemCardQuery"), 0) = 1, _
            "—’Ìœ √Ê· «·„œ… »«·ﬂ„Ì… Ê«·ﬁÌ„…", passed, failed, report

    ' a transfer shows the other warehouse as its party
    msg = CreateStockTransfer(MAIN_WAREHOUSE_ID, wh2, "TEST", tid)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      tid & ", " & pid & ", 3, 0)", dbFailOnError
    msg = msg & SendStockTransfer(tid)
    SetPeriod Date, Date
    CheckWh Len(msg) = 0 And Nz(DbValue("SELECT PartyName FROM ItemCardQuery WHERE NetQty = -3"), "") = branchName, _
            "«·ÃÂ… ›Ì ﬂ«—  «·’‰›: «·„Œ“‰ «·¬Œ— " & msg, passed, failed, report
    CheckWh Nz(DbValue("SELECT TotalQuantity FROM StockValueByWarehouseQuery WHERE WarehouseID = " & wh2), 0) = 7 And _
            Nz(DbValue("SELECT StockCostValue FROM StockValueByWarehouseQuery WHERE WarehouseID = " & wh2), 0) = 35, _
            " ﬁÌÌ„ «·„Œ“Ê‰ ·ﬂ· „Œ“‰", passed, failed, report

    ' a transfer waiting for its receiver
    CurrentDb.Execute "UPDATE Warehouses SET ConfirmReceipt = True WHERE WarehouseID = " & wh2, dbFailOnError
    msg = CreateStockTransfer(MAIN_WAREHOUSE_ID, wh2, "TEST", tid)
    CurrentDb.Execute "INSERT INTO StockTransferLines (StockTransferID, ProductID, Quantity, UnitCost) VALUES (" & _
                      tid & ", " & pid & ", 1, 0)", dbFailOnError
    msg = msg & SendStockTransfer(tid)
    CheckWh Len(msg) = 0 And Nz(DbValue("SELECT TransferValue FROM TransfersInTransitQuery WHERE StockTransferID = " & _
            tid), 0) = 5, "«· ÕÊÌ·«  »«·ÿ—Ìﬁ " & msg, passed, failed, report

    ' one open count per warehouse
    CurrentDb.Execute "UPDATE StockCounts SET Status = 'CANCELLED' WHERE Status = 'OPEN'", dbFailOnError
    msg = CreateStockCount(Null, "TEST", countID, MAIN_WAREHOUSE_ID)
    CheckWh Len(msg) = 0, "Ã—œ «·„Œ“‰ «·—∆Ì”Ì " & msg, passed, failed, report
    msg = CreateStockCount(Null, "TEST", countID, wh2)
    CheckWh Len(msg) = 0 And Nz(DbValue("SELECT Count(*) FROM StockCounts WHERE Status = 'OPEN'"), 0) = 2, _
            "Ã—œ «·›—⁄ ›Ì ‰›” «·Êﬁ  " & msg, passed, failed, report
    CheckWh Len(CreateStockCount(Null, "TEST", countID, wh2)) > 0, "Ã—œ „› ÊÕ Ê«Õœ ·ﬂ· „Œ“‰", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckWh False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ﬂ«—  «·’‰› Ê ﬁ«—Ì— «·„Œ«“‰ ‰«ÃÕ… (" & passed & " «Œ »«—«).", _
                vbInformation + MSG_RTL, "TestStockReports"
        TestStockReports = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestStockReports"
    End If
End Function

Private Sub CheckWh(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
