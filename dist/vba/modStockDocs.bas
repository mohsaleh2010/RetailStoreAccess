Attribute VB_Name = "modStockDocs"
'==============================================================================
' modStockDocs  -  Retail Store Management System (stock issue and receipt vouchers, docs/59)
'
' A stock party (StockParties) is where goods go to or come from, other than customers and suppliers:
' a department, a production hall, the waste, or another party. Its account: its own, else the account
' of its type (department 5640, production hall 1420 work in progress, waste 5250, other 5900).
'
'   CreateStockDocument   a draft voucher: ISSUE (out of a warehouse to a party) or RECEIPT (from a party)
'   PostStockDocument     moves the stock (types 11 / 12, reference STOCK_DOC): an issue at the average
'                         cost; a receipt at the cost typed on each line, which updates the average cost
'   UnpostStockDocument   takes the moves back (the average cost of a receipt is taken back too); the
'                         voucher is a draft again
'   DeleteStockDocument   a draft only
' Its entry (qryJournalStockDocument): issue = party account / stock of the warehouse, receipt = the reverse.
' Screens: frmStockParties (a data screen), frmStockDocument (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Public Const TT_PARTY_ISSUE As Long = 11
Public Const TT_PARTY_RECEIPT As Long = 12

Public Function StockPartyIsActive(ByVal StockPartyID As Variant) As Boolean
    If IsNull(StockPartyID) Then Exit Function
    If Not IsNumeric(StockPartyID) Then Exit Function
    StockPartyIsActive = Nz(DbValue("SELECT COUNT(*) FROM StockParties WHERE IsActive <> 0 AND StockPartyID = " & _
                                    CLng(StockPartyID)), 0) > 0
End Function

Public Function StockPartyRows() As String
    StockPartyRows = Tr("SELECT p.StockPartyID, p.PartyName FROM [@StockParties] AS p WHERE p.IsActive <> 0 " & _
                        "ORDER BY p.PartyName")
End Function

Private Function StockDocStatus(ByVal StockDocumentID As Long) As String
    StockDocStatus = Nz(DbValue("SELECT Status FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID), "")
End Function

Private Function StockDocType(ByVal StockDocumentID As Long) As String
    StockDocType = Nz(DbValue("SELECT DocType FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID), "")
End Function

Public Function ReceiptDefaultCost(ByVal ProductID As Long) As Currency
    ' the cost a receipt line starts with: the average cost, else the purchase price
    ReceiptDefaultCost = Nz(DbValue("SELECT IIf(AverageCost > 0, AverageCost, PurchasePrice) FROM Products " & _
                                    "WHERE ProductID = " & ProductID), 0)
End Function

Public Function CreateStockDocument(ByVal DocType As String, ByVal WarehouseID As Variant, ByVal StockPartyID As Variant, _
                                    ByVal Notes As String, ByRef NewDocumentID As Long) As String
    Dim rs As DAO.Recordset
    NewDocumentID = 0
    If Not HasPermission("STOCK_DOCS") Then
        CreateStockDocument = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·’—› Ê«·«” ·«„ «·„Œ“‰Ì."
    ElseIf DocType <> "ISSUE" And DocType <> "RECEIPT" Then
        CreateStockDocument = "«Œ — ‰Ê⁄ «·”‰œ: ’—› √Ê «” ·«„."
    ElseIf Not WarehouseIsActive(WarehouseID) Then
        CreateStockDocument = "«Œ — «·„Œ“‰."
    ElseIf Not StockPartyIsActive(StockPartyID) Then
        CreateStockDocument = "«Œ — «·ÃÂ…."
    End If
    If Len(CreateStockDocument) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("StockDocuments", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    rs.AddNew
    rs!DocNumber = NextNumber(IIf(DocType = "ISSUE", "STOCK_ISSUE", "STOCK_RECEIPT"))
    rs!DocType = DocType
    rs!DocDate = Now
    rs!WarehouseID = CLng(WarehouseID)
    rs!StockPartyID = CLng(StockPartyID)
    rs!Status = "DRAFT"
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewDocumentID = rs!StockDocumentID
    rs.Close
    LogAction "STOCK_DOC_NEW", "StockDocuments", CStr(NewDocumentID)
End Function

Public Function StockDocLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant, ByVal UnitCost As Variant) As String
    If IsNull(ProductID) Then
        StockDocLineProblem = "«Œ — «·’‰›."
    ElseIf Not ProductTracksStock(CLng(ProductID)) Then
        StockDocLineProblem = "Â–« «·’‰› ·« Ì «»⁄ «·„Œ“Ê‰° ›·« Ìı’—› Ê·« Ìı” ·„."
    ElseIf Nz(Qty, 0) <= 0 Then
        StockDocLineProblem = "«·ﬂ„Ì… √ﬂ»— „‰ ’›—."
    ElseIf Nz(UnitCost, 0) < 0 Then
        StockDocLineProblem = " ﬂ·›… «·ÊÕœ… ·«  ﬂÊ‰ ”«·»…."
    End If
End Function

Public Function PostStockDocument(ByVal StockDocumentID As Long) As String
    ' An issue leaves the warehouse at the average cost; a receipt enters it at its typed cost.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, msg As String
    Dim wh As Long, docNo As String, docKind As String, party As Variant, partyName As String
    Dim n As Long, ids() As Long, qtys() As Currency, cost As Currency
    If Not HasPermission("STOCK_DOCS") Then
        PostStockDocument = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·’—› Ê«·«” ·«„ «·„Œ“‰Ì."
        Exit Function
    End If
    If StockDocStatus(StockDocumentID) <> "DRAFT" Then
        PostStockDocument = "«·”‰œ €Ì— „ÊÃÊœ √Ê  „  —ÕÌ·Â."
        Exit Function
    End If
    docKind = StockDocType(StockDocumentID)
    wh = DbValue("SELECT WarehouseID FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID)
    party = DbValue("SELECT StockPartyID FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID)
    docNo = DbValue("SELECT DocNumber FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID)
    msg = WarehouseProblem(wh)
    If Len(msg) = 0 Then
        If Not StockPartyIsActive(party) Then msg = "«·ÃÂ… €Ì— „ÊÃÊœ… √Ê €Ì— ‰‘ÿ…."
    End If
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(Date)
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity, UnitCost FROM StockDocumentLines " & _
                                         "WHERE StockDocumentID = " & StockDocumentID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            If Len(msg) = 0 Then msg = StockDocLineProblem(rs!ProductID, rs!Quantity, rs!UnitCost)
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) = 0 And n = 0 Then msg = "√÷› «·√’‰«› Êﬂ„Ì« Â« √Ê·«."
    End If
    If Len(msg) = 0 And docKind = "ISSUE" Then msg = CheckStockAvailable(ids, qtys, n, "«·’—› «·„Œ“‰Ì", wh)
    If Len(msg) > 0 Then
        PostStockDocument = msg
        Exit Function
    End If
    partyName = Nz(DbValue(Tr("SELECT p.PartyName FROM [@StockParties] AS p WHERE p.StockPartyID = ") & CLng(party)), "")

    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT StockDocumentLineID, ProductID, Quantity, UnitCost FROM StockDocumentLines " & _
                              "WHERE StockDocumentID = " & StockDocumentID & " ORDER BY StockDocumentLineID", dbOpenDynaset, dbSeeChanges)
    Do Until rs.EOF
        If docKind = "ISSUE" Then
            cost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & rs!ProductID), 0)
            rs.Edit
            rs!UnitCost = cost
            rs.Update
            ApplyStockMovement db, rs!ProductID, -rs!Quantity, TT_PARTY_ISSUE, cost, "STOCK_DOC", StockDocumentID, _
                               docNo, partyName, False, wh
        Else
            ApplyStockMovement db, rs!ProductID, rs!Quantity, TT_PARTY_RECEIPT, rs!UnitCost, "STOCK_DOC", StockDocumentID, _
                               docNo, partyName, True, wh
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE StockDocuments SET DocDate = Now(), Status = 'POSTED' WHERE StockDocumentID = " & StockDocumentID, _
               dbFailOnError + dbSeeChanges
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_DOC_POST", "StockDocuments", docNo, docKind & " Warehouse=" & wh & " Party=" & party
    Exit Function
EH:
    PostStockDocument = " ⁄–—  —ÕÌ· «·”‰œ: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function UnpostStockDocument(ByVal StockDocumentID As Long) As String
    ' Its moves are deleted: an issue gives the goods back, a receipt takes them out again with the average
    ' cost it brought (WeightedAverage of the opposite quantity). The voucher becomes a draft to correct.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, before As Collection
    Dim docKind As String, wh As Long, n As Long, ids() As Long, qtys() As Currency, cur As DAO.Recordset
    Dim newAvg As Currency
    If Not HasPermission("STOCK_DOCS") Then
        UnpostStockDocument = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·’—› Ê«·«” ·«„ «·„Œ“‰Ì."
        Exit Function
    End If
    If StockDocStatus(StockDocumentID) <> "POSTED" Then
        UnpostStockDocument = "«·”‰œ ·Ì” „ı—ÕÛ¯·«."
        Exit Function
    End If
    UnpostStockDocument = ClosedPeriodProblem(DbValue("SELECT DocDate FROM StockDocuments WHERE StockDocumentID = " & _
                                                      StockDocumentID))
    If Len(UnpostStockDocument) > 0 Then Exit Function
    docKind = StockDocType(StockDocumentID)
    wh = DbValue("SELECT WarehouseID FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID)
    If docKind = "RECEIPT" Then                       ' the received goods leave the warehouse again
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity FROM StockDocumentLines WHERE StockDocumentID = " & _
                                         StockDocumentID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If n > 0 Then UnpostStockDocument = CheckStockAvailable(ids, qtys, n, "≈·€«¡  —ÕÌ· ”‰œ «·«” ·«„", wh)
        If Len(UnpostStockDocument) > 0 Then Exit Function
    End If

    On Error GoTo EH
    Set before = AuditSnapshot("StockDocuments", "StockDocumentID", StockDocumentID)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT ProductID, Quantity, UnitCost FROM InventoryTransactions WHERE ReferenceType = " & _
                              "'STOCK_DOC' AND ReferenceID = " & StockDocumentID, dbOpenSnapshot)
    Do Until rs.EOF
        Set cur = db.OpenRecordset("SELECT ProductID, CurrentQuantity, AverageCost, UpdatedAt FROM Products " & _
                                   "WHERE ProductID = " & rs!ProductID, dbOpenDynaset, dbSeeChanges)
        cur.Edit
        newAvg = cur!AverageCost
        If docKind = "RECEIPT" Then newAvg = WeightedAverage(cur!CurrentQuantity, cur!AverageCost, -rs!Quantity, rs!UnitCost)
        cur!AverageCost = newAvg
        cur!CurrentQuantity = cur!CurrentQuantity - rs!Quantity
        cur!UpdatedAt = Now
        cur.Update
        cur.Close
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "DELETE FROM InventoryTransactions WHERE ReferenceType = 'STOCK_DOC' AND ReferenceID = " & StockDocumentID, _
               dbFailOnError + dbSeeChanges
    db.Execute "UPDATE StockDocuments SET Status = 'DRAFT' WHERE StockDocumentID = " & StockDocumentID, dbFailOnError + dbSeeChanges
    ws.CommitTrans
    inTrans = False
    AuditEdited "STOCK_DOC_UNPOST", "StockDocuments", "StockDocumentID", StockDocumentID, before
    Exit Function
EH:
    UnpostStockDocument = " ⁄–— ≈·€«¡ «· —ÕÌ·: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function DeleteStockDocument(ByVal StockDocumentID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("STOCK_DOCS") Then
        DeleteStockDocument = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·’—› Ê«·«” ·«„ «·„Œ“‰Ì."
        Exit Function
    End If
    If StockDocStatus(StockDocumentID) <> "DRAFT" Then
        DeleteStockDocument = " ıÕ–› «·„”Êœ… ›ﬁÿ. √·€ˆ  —ÕÌ· «·”‰œ √Ê·«."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("StockDocuments", "StockDocumentID", StockDocumentID)
    CurrentDb.Execute "DELETE FROM StockDocumentLines WHERE StockDocumentID = " & StockDocumentID, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM StockDocuments WHERE StockDocumentID = " & StockDocumentID, dbFailOnError + dbSeeChanges
    AuditDeleted "STOCK_DOC_DELETE", "StockDocuments", StockDocumentID, auditBefore
End Function

'==============================================================================
' frmStockDocument (OpenArgs = a voucher to show) and its lines frmStockDocumentLines
'==============================================================================
Public Sub StockDocLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboWarehouse.RowSource = WarehouseRows()
    frm!cboParty.RowSource = StockPartyRows()
    frm!cboWarehouse.Value = CurrentWarehouseID()
    frm!cboType.Value = "ISSUE"
    StockDocListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(StockDocumentID) FROM StockDocuments")
    frm!cboDoc.Value = id
    StockDocShow frm
End Sub

Private Sub StockDocListRefresh(ByVal frm As Access.Form)
    frm!cboDoc.RowSource = Tr("SELECT d.StockDocumentID, d.DocNumber, Format(d.DocDate, 'yyyy/mm/dd'), " & _
        "IIf(d.DocType = 'ISSUE', '’—›', '«” ·«„'), p.PartyName, w.WarehouseName, IIf(d.Status = 'DRAFT', '„”Êœ…', " & _
        "'„ı—ÕÛ¯·') FROM (StockDocuments AS d INNER JOIN [@StockParties] AS p ON d.StockPartyID = p.StockPartyID) " & _
        "INNER JOIN [@Warehouses] AS w ON d.WarehouseID = w.WarehouseID ORDER BY d.StockDocumentID DESC")
    frm!cboDoc.Requery
End Sub

Public Sub StockDocShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, rs As DAO.Recordset, info As String, docKind As String
    id = Nz(frm!cboDoc.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM StockDocumentLines WHERE StockDocumentID = " & id & _
                                     " ORDER BY StockDocumentLineID"
    st = StockDocStatus(id)
    docKind = StockDocType(id)
    If Len(st) = 0 Then
        info = "·« ÌÊÃœ ”‰œ. «Œ — «·‰Ê⁄ Ê«·„Œ“‰ Ê«·ÃÂ… À„ ´”‰œ ÃœÌœª."
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT DocNumber, DocDate, WarehouseID, StockPartyID FROM StockDocuments " & _
                                         "WHERE StockDocumentID = " & id, dbOpenSnapshot)
        info = rs!DocNumber & "    " & IIf(docKind = "ISSUE", "’—› „‰ ", "«” ·«„ ›Ì ") & WarehouseName(rs!WarehouseID) & _
               IIf(docKind = "ISSUE", "  ->  ", "  <-  ") & _
               Nz(DbValue(Tr("SELECT p.PartyName FROM [@StockParties] AS p WHERE p.StockPartyID = " & rs!StockPartyID)), "") & _
               "    " & IIf(st = "DRAFT", "„”Êœ…: √÷› «·√’‰«› À„ ´ —ÕÌ·ª", "„ı—ÕÛ¯· ›Ì " & GDate(rs!DocDate))
        rs.Close
        info = info & "    «·ﬁÌ„…: " & Format$(Nz(DbValue("SELECT DocValue FROM qryStockDocumentValue " & _
                                                         "WHERE StockDocumentID = " & id), 0), "#,##0.00")
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    frm!subLines.Form!UnitCost.Locked = (docKind <> "RECEIPT")    ' an issue leaves at the average cost
    SafeFocus frm!cboDoc                    ' a focused control cannot be disabled
    frm!btnPost.Enabled = (st = "DRAFT")
    frm!btnDelete.Enabled = (st = "DRAFT")
    frm!btnUnpost.Enabled = (st = "POSTED")
End Sub

Public Sub StockDocNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateStockDocument(Nz(frm!cboType.Value, ""), frm!cboWarehouse.Value, frm!cboParty.Value, _
                              Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    StockDocListRefresh frm
    frm!cboDoc.Value = id
    StockDocShow frm
End Sub

Public Sub StockDocAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String
    id = Nz(frm!cboDoc.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "POST"
            If Not AskYesNo(" —ÕÌ· «·”‰œø Ì Õ—ﬂ «·„Œ“Ê‰ ÊÌı‰‘√ «·ﬁÌœ.") Then Exit Sub
            msg = PostStockDocument(id)
        Case "UNPOST"
            If Not AskYesNo("≈·€«¡  —ÕÌ· «·”‰œ Ê≈⁄«œ… «·„Œ“Ê‰ ﬂ„« ﬂ«‰ø") Then Exit Sub
            msg = UnpostStockDocument(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› „”Êœ… «·”‰œø") Then Exit Sub
            msg = DeleteStockDocument(id)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    StockDocListRefresh frm
    If Action = "DELETE" Then frm!cboDoc.Value = DbValue("SELECT Max(StockDocumentID) FROM StockDocuments")
    StockDocShow frm
End Sub

Public Sub StockDocLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String, docKind As String
    id = Nz(frm.Parent!cboDoc.Value, 0)
    docKind = StockDocType(id)
    If StockDocStatus(id) <> "DRAFT" Then
        msg = " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
    Else
        msg = StockDocLineProblem(frm!ProductID.Value, frm!Quantity.Value, frm!UnitCost.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!StockDocumentID.Value = id
    If docKind = "ISSUE" Then
        frm!UnitCost.Value = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & CLng(frm!ProductID.Value)), 0)
    ElseIf Nz(frm!UnitCost.Value, 0) = 0 Then
        frm!UnitCost.Value = ReceiptDefaultCost(CLng(frm!ProductID.Value))
    End If
    AuditFormBefore frm, "StockDocumentLines", "StockDocumentLineID"
End Sub

Public Sub StockDocLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    StockDocShow frm.Parent
End Sub

Public Sub StockDocRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!StockDocumentLineID.Value
    If IsNull(id) Then Exit Sub
    If StockDocStatus(Nz(frm.Parent!cboDoc.Value, 0)) <> "DRAFT" Then
        ShowWarning " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("StockDocumentLines", "StockDocumentLineID", id)
    CurrentDb.Execute "DELETE FROM StockDocumentLines WHERE StockDocumentLineID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "STOCK_DOC_LINE_DELETE", "StockDocumentLines", id, auditBefore
    frm.Requery
    StockDocShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestStockDocs() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, refNo As String, msg As String
    Dim passed As Long, failed As Long, report As String, partyID As Long, docID As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestStockDocs  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckSd StockPartyIsActive(1), "ÃÂ… «·Â«·ﬂ „ÊÃÊœ… Ê‰‘ÿ…", passed, failed, report
    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO StockParties (PartyCode, PartyName, PartyType) VALUES " & _
                      "('TEST-SP', 'TEST ’«·… «·≈‰ «Ã', 'PRODUCTION')", dbFailOnError + dbSeeChanges
    partyID = DbValue("SELECT Max(StockPartyID) FROM StockParties")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-SD-P', 'TEST ’‰› «·”‰œ« ', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 0, 0)", dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-SD-P'")
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckSd Len(msg) = 0, "≈÷«›… 10 » ﬂ·›… 6 " & msg, passed, failed, report

    ' an issue to the production hall at the average cost
    msg = CreateStockDocument("ISSUE", MAIN_WAREHOUSE_ID, partyID, "TEST", docID)
    CurrentDb.Execute "INSERT INTO StockDocumentLines (StockDocumentID, ProductID, Quantity, UnitCost) VALUES (" & _
                      docID & ", " & pid & ", 4, 0)", dbFailOnError + dbSeeChanges
    msg = msg & PostStockDocument(docID)
    CheckSd Len(msg) = 0 And StockDocStatus(docID) = "POSTED" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 6, _
            "”‰œ «·’—› ÌŒ’„ „‰ «·„Œ“‰ " & msg, passed, failed, report
    CheckSd Nz(DbValue("SELECT DocValue FROM qryStockDocumentValue WHERE StockDocumentID = " & docID), 0) = 24, _
            "«·’—› »„ Ê”ÿ «· ﬂ·›…", passed, failed, report
    CheckSd Nz(DbValue("SELECT Sum(Debit) FROM qryJournalStockDocument WHERE SourceID = " & docID & _
                       " AND AccountCode = 1420"), 0) = 24, "ﬁÌœ «·’—› ·’«·… «·≈‰ «Ã ⁄·Ï 1420", passed, failed, report
    msg = UnpostStockDocument(docID)
    CheckSd Len(msg) = 0 And StockDocStatus(docID) = "DRAFT" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10 And _
            Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 10, _
            "≈·€«¡  —ÕÌ· «·’—› Ì⁄Ìœ «·√’‰«› " & msg, passed, failed, report
    CheckSd Len(DeleteStockDocument(docID)) = 0 And StockDocStatus(docID) = "", "Õ–› «·„”Êœ…", passed, failed, report

    ' a receipt at its own cost updates the average; unposting it gives the average back
    msg = CreateStockDocument("RECEIPT", MAIN_WAREHOUSE_ID, partyID, "TEST", docID)
    CurrentDb.Execute "INSERT INTO StockDocumentLines (StockDocumentID, ProductID, Quantity, UnitCost) VALUES (" & _
                      docID & ", " & pid & ", 10, 10)", dbFailOnError + dbSeeChanges
    msg = msg & PostStockDocument(docID)
    CheckSd Len(msg) = 0 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 20 And _
            Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & pid), 0) = 8, _
            "”‰œ «·«” ·«„ Ì÷Ì› ÊÌ⁄Ìœ Õ”«» «·„ Ê”ÿ " & msg, passed, failed, report
    CheckSd Len(DeleteStockDocument(docID)) > 0, "«·”‰œ «·„ı—ÕÛ¯· ·« ÌıÕ–›", passed, failed, report
    msg = UnpostStockDocument(docID)
    CheckSd Len(msg) = 0 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 10 And _
            Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & pid), 0) = 6, _
            "≈·€«¡  —ÕÌ· «·«” ·«„ Ì⁄Ìœ «·ﬂ„Ì… Ê«·„ Ê”ÿ " & msg, passed, failed, report

    ' not more than the stock of the warehouse
    msg = CreateStockDocument("ISSUE", MAIN_WAREHOUSE_ID, 1, "TEST", docID)
    CurrentDb.Execute "INSERT INTO StockDocumentLines (StockDocumentID, ProductID, Quantity, UnitCost) VALUES (" & _
                      docID & ", " & pid & ", 999, 0)", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "UPDATE Settings SET AllowNegativeStock = False", dbFailOnError + dbSeeChanges
    g_AutoAnswer = False
    CheckSd Len(PostStockDocument(docID)) > 0 And StockDocStatus(docID) = "DRAFT", "·« Ìı’—› √ﬂÀ— „‰ —’Ìœ «·„Œ“‰", _
            passed, failed, report
    CheckSd Len(CreateStockDocument("ISSUE", MAIN_WAREHOUSE_ID, Null, "TEST", docID)) > 0, "«·”‰œ ÌÕ «Ã ÃÂ…", _
            passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckSd False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ”‰œ«  «·’—› Ê«·«” ·«„ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, _
                "TestStockDocs"
        TestStockDocs = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestStockDocs"
    End If
End Function

Private Sub CheckSd(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
