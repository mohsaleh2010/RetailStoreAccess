Attribute VB_Name = "modStockNotes"
'==============================================================================
' modStockNotes  -  Retail Store Management System (delivery and receipt notes, docs/61)
'
' A delivery note hands goods to a customer before the invoice; a receipt note takes goods in from a
' supplier before the invoice. The note moves the stock; the invoice made from it does not move it again.
'
'   CreateStockNote    a draft: DELIVERY (to a customer) or RECEIPT (from a supplier)
'   PostStockNote      moves the stock (types 13 / 14, reference STOCK_NOTE): a delivery at the average cost
'                      (entry 1430 goods delivered not invoiced / stock), a receipt at its purchase price,
'                      which updates the average cost (entry stock / 2150 goods received not invoiced)
'   UnpostStockNote    takes the moves back while no invoice was made; the note is a draft again
'   DeleteStockNote    a draft only
'   InvoiceStockNote   fills the sales or purchase cart with the lines of the note and posts the invoice
'                      through PostSaleFromCart / PostPurchaseFromCart with StockNoteID: no stock moves, the
'                      cost of each line is the cost of the note; the invoice clears 1430 / 2150
'                      (qryInvoicedNotes in qryJournalSale / qryJournalPurchase)
' A delivery note made from a customer order (docs/62) carries SalesOrderID and, on each line, SalesOrderLineID:
' posting it checks the order (OrderNoteProblem) and frees its reservation; the order follows its notes
' (SyncSalesOrder, modSalesOrders).
' Screen: frmStockNote (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Public Const TT_DELIVERY_NOTE As Long = 13
Public Const TT_RECEIPT_NOTE As Long = 14

Private Function NoteStatus(ByVal StockNoteID As Long) As String
    NoteStatus = Nz(DbValue("SELECT Status FROM StockNotes WHERE StockNoteID = " & StockNoteID), "")
End Function

Private Function NoteKind(ByVal StockNoteID As Long) As String
    NoteKind = Nz(DbValue("SELECT NoteType FROM StockNotes WHERE StockNoteID = " & StockNoteID), "")
End Function

Public Function NoteOrderID(ByVal StockNoteID As Long) As Long
    ' The customer order a delivery note was made from (docs/62), else 0.
    NoteOrderID = Nz(DbValue("SELECT SalesOrderID FROM StockNotes WHERE StockNoteID = " & StockNoteID), 0)
End Function

Private Function NoteParty(ByVal StockNoteID As Long) As Long
    NoteParty = Nz(DbValue("SELECT IIf(NoteType = 'DELIVERY', CustomerID, SupplierID) FROM StockNotes " & _
                           "WHERE StockNoteID = " & StockNoteID), 0)
End Function

Public Function NotePartyRows(ByVal NoteType As String) As String
    ' the active customers of a delivery note, the active suppliers of a receipt note
    If NoteType = "RECEIPT" Then
        NotePartyRows = Tr("SELECT s.SupplierID, s.SupplierName FROM [@Suppliers] AS s WHERE s.IsActive <> 0 " & _
                           "ORDER BY s.SupplierName")
    Else
        NotePartyRows = Tr("SELECT c.CustomerID, c.CustomerName FROM [@Customers] AS c WHERE c.IsActive <> 0 " & _
                           "ORDER BY c.CustomerName")
    End If
End Function

Public Function NotePartyIsActive(ByVal NoteType As String, ByVal PartyID As Variant) As Boolean
    If IsNull(PartyID) Then Exit Function
    If Not IsNumeric(PartyID) Then Exit Function
    If NoteType = "RECEIPT" Then
        NotePartyIsActive = Nz(DbValue("SELECT COUNT(*) FROM Suppliers WHERE IsActive <> 0 AND SupplierID = " & _
                                       CLng(PartyID)), 0) > 0
    Else
        NotePartyIsActive = Nz(DbValue("SELECT COUNT(*) FROM Customers WHERE IsActive <> 0 AND CustomerID = " & _
                                       CLng(PartyID)), 0) > 0
    End If
End Function

Private Function NotePartyName(ByVal StockNoteID As Long) As String
    If NoteKind(StockNoteID) = "RECEIPT" Then
        NotePartyName = Nz(DbValue(Tr("SELECT s.SupplierName FROM [@Suppliers] AS s WHERE s.SupplierID = ") & _
                                   NoteParty(StockNoteID)), "")
    Else
        NotePartyName = Nz(DbValue(Tr("SELECT c.CustomerName FROM [@Customers] AS c WHERE c.CustomerID = ") & _
                                   NoteParty(StockNoteID)), "")
    End If
End Function

Public Function NoteDefaultPrice(ByVal NoteType As String, ByVal ProductID As Long) As Currency
    ' the price a line starts with: the selling price of a delivery, the purchase price of a receipt
    If NoteType = "RECEIPT" Then
        NoteDefaultPrice = Nz(DbValue("SELECT IIf(PurchasePrice > 0, PurchasePrice, AverageCost) FROM Products " & _
                                      "WHERE ProductID = " & ProductID), 0)
    Else
        NoteDefaultPrice = Nz(DbValue("SELECT SellingPrice FROM Products WHERE ProductID = " & ProductID), 0)
    End If
End Function

Public Function CreateStockNote(ByVal NoteType As String, ByVal WarehouseID As Variant, ByVal PartyID As Variant, _
                                ByVal Notes As String, ByRef NewNoteID As Long) As String
    Dim rs As DAO.Recordset
    NewNoteID = 0
    If Not HasPermission("STOCK_NOTES") Then
        CreateStockNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
    ElseIf NoteType <> "DELIVERY" And NoteType <> "RECEIPT" Then
        CreateStockNote = "«Œ — ‰Ê⁄ «·≈–‰:  ”·Ì„ ·⁄„Ì· √Ê «” ·«„ „‰ „Ê—œ."
    ElseIf Not WarehouseIsActive(WarehouseID) Then
        CreateStockNote = "«Œ — «·„Œ“‰."
    ElseIf Not NotePartyIsActive(NoteType, PartyID) Then
        CreateStockNote = IIf(NoteType = "RECEIPT", "«Œ — «·„Ê—œ.", "«Œ — «·⁄„Ì·.")
    End If
    If Len(CreateStockNote) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("StockNotes", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    rs.AddNew
    rs!NoteNumber = NextNumber(IIf(NoteType = "DELIVERY", "DELIVERY_NOTE", "RECEIPT_NOTE"))
    rs!NoteType = NoteType
    rs!NoteDate = Now
    rs!WarehouseID = CLng(WarehouseID)
    If NoteType = "DELIVERY" Then
        rs!CustomerID = CLng(PartyID)
    Else
        rs!SupplierID = CLng(PartyID)
    End If
    rs!Status = "DRAFT"
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewNoteID = rs!StockNoteID
    rs.Close
    LogAction "STOCK_NOTE_NEW", "StockNotes", CStr(NewNoteID)
End Function

Public Function StockNoteLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant, ByVal UnitPrice As Variant) As String
    If IsNull(ProductID) Then
        StockNoteLineProblem = "«Œ — «·’‰›."
    ElseIf Not ProductTracksStock(CLng(ProductID)) Then
        StockNoteLineProblem = "Â–« «·’‰› ·« Ì «»⁄ «·„Œ“Ê‰° ›·« Ìı”·Û¯„ Ê·« Ìı” ·„ »≈–‰."
    ElseIf Nz(Qty, 0) <= 0 Then
        StockNoteLineProblem = "«·ﬂ„Ì… √ﬂ»— „‰ ’›—."
    ElseIf Nz(UnitPrice, 0) < 0 Then
        StockNoteLineProblem = "«·”⁄— ·« ÌﬂÊ‰ ”«·»«."
    End If
End Function

Public Function PostStockNote(ByVal StockNoteID As Long) As String
    ' A delivery leaves the warehouse at the average cost; a receipt enters it at its purchase price.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, msg As String
    Dim wh As Long, noteNo As String, kind As String, partyName As String, orderID As Long
    Dim n As Long, ids() As Long, qtys() As Currency, cost As Currency
    If Not HasPermission("STOCK_NOTES") Then
        PostStockNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    If NoteStatus(StockNoteID) <> "DRAFT" Then
        PostStockNote = "«·≈–‰ €Ì— „ÊÃÊœ √Ê  „  —ÕÌ·Â."
        Exit Function
    End If
    kind = NoteKind(StockNoteID)
    wh = DbValue("SELECT WarehouseID FROM StockNotes WHERE StockNoteID = " & StockNoteID)
    noteNo = DbValue("SELECT NoteNumber FROM StockNotes WHERE StockNoteID = " & StockNoteID)
    orderID = NoteOrderID(StockNoteID)
    msg = WarehouseProblem(wh)
    If Len(msg) = 0 Then
        If Not NotePartyIsActive(kind, NoteParty(StockNoteID)) Then
            msg = IIf(kind = "RECEIPT", "«·„Ê—œ €Ì— „ÊÃÊœ √Ê €Ì— ‰‘ÿ.", "«·⁄„Ì· €Ì— „ÊÃÊœ √Ê €Ì— ‰‘ÿ.")
        End If
    End If
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(Date)
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity, UnitPrice FROM StockNoteLines " & _
                                         "WHERE StockNoteID = " & StockNoteID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            If Len(msg) = 0 Then msg = StockNoteLineProblem(rs!ProductID, rs!Quantity, rs!UnitPrice)
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) = 0 And n = 0 Then msg = "√÷› «·√’‰«› Êﬂ„Ì« Â« √Ê·«."
    End If
    If Len(msg) = 0 And orderID > 0 Then msg = OrderNoteProblem(StockNoteID)          ' modSalesOrders
    If Len(msg) = 0 And kind = "DELIVERY" Then msg = CheckStockAvailable(ids, qtys, n, "≈–‰ «· ”·Ì„", wh, orderID)
    If Len(msg) > 0 Then
        PostStockNote = msg
        Exit Function
    End If
    partyName = NotePartyName(StockNoteID)

    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT StockNoteLineID, ProductID, Quantity, UnitPrice, UnitCost FROM StockNoteLines " & _
                              "WHERE StockNoteID = " & StockNoteID & " ORDER BY StockNoteLineID", dbOpenDynaset, dbSeeChanges)
    Do Until rs.EOF
        If kind = "DELIVERY" Then
            cost = Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & rs!ProductID), 0)
        Else
            cost = Nz(rs!UnitPrice, 0)                    ' the purchase price without VAT
        End If
        rs.Edit
        rs!UnitCost = cost
        rs.Update
        If kind = "DELIVERY" Then
            ApplyStockMovement db, rs!ProductID, -rs!Quantity, TT_DELIVERY_NOTE, cost, "STOCK_NOTE", StockNoteID, _
                               noteNo, partyName, False, wh
        Else
            ApplyStockMovement db, rs!ProductID, rs!Quantity, TT_RECEIPT_NOTE, cost, "STOCK_NOTE", StockNoteID, _
                               noteNo, partyName, True, wh
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE StockNotes SET NoteDate = Now(), Status = 'POSTED' WHERE StockNoteID = " & StockNoteID, dbFailOnError + dbSeeChanges
    If orderID > 0 Then SyncSalesOrder db, orderID                      ' partly or fully delivered
    ws.CommitTrans
    inTrans = False
    LogAction "STOCK_NOTE_POST", "StockNotes", noteNo, kind & " Warehouse=" & wh
    Exit Function
EH:
    PostStockNote = " ⁄–—  —ÕÌ· «·≈–‰: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function UnpostStockNote(ByVal StockNoteID As Long) As String
    ' Only before its invoice. Its moves are deleted: a delivery gives the goods back, a receipt takes them
    ' out again with the average cost it brought (WeightedAverage of the opposite quantity).
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, before As Collection
    Dim kind As String, wh As Long, n As Long, ids() As Long, qtys() As Currency, cur As DAO.Recordset
    Dim newAvg As Currency, st As String
    If Not HasPermission("STOCK_NOTES") Then
        UnpostStockNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    st = NoteStatus(StockNoteID)
    If st = "INVOICED" Then
        UnpostStockNote = "’œ—  ›« Ê—… Â–« «·≈–‰° ›·« Ìı·€Ï  —ÕÌ·Â. «⁄„· „— Ã⁄« ··›« Ê—…."
        Exit Function
    End If
    If st <> "POSTED" Then
        UnpostStockNote = "«·≈–‰ ·Ì” „ı—ÕÛ¯·«."
        Exit Function
    End If
    UnpostStockNote = ClosedPeriodProblem(DbValue("SELECT NoteDate FROM StockNotes WHERE StockNoteID = " & StockNoteID))
    If Len(UnpostStockNote) > 0 Then Exit Function
    kind = NoteKind(StockNoteID)
    wh = DbValue("SELECT WarehouseID FROM StockNotes WHERE StockNoteID = " & StockNoteID)
    If kind = "RECEIPT" Then                          ' the received goods leave the warehouse again
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity FROM StockNoteLines WHERE StockNoteID = " & _
                                         StockNoteID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If n > 0 Then UnpostStockNote = CheckStockAvailable(ids, qtys, n, "≈·€«¡  —ÕÌ· ≈–‰ «·«” ·«„", wh)
        If Len(UnpostStockNote) > 0 Then Exit Function
    End If

    On Error GoTo EH
    Set before = AuditSnapshot("StockNotes", "StockNoteID", StockNoteID)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT ProductID, Quantity, UnitCost FROM InventoryTransactions WHERE ReferenceType = " & _
                              "'STOCK_NOTE' AND ReferenceID = " & StockNoteID, dbOpenSnapshot)
    Do Until rs.EOF
        Set cur = db.OpenRecordset("SELECT ProductID, CurrentQuantity, AverageCost, UpdatedAt FROM Products " & _
                                   "WHERE ProductID = " & rs!ProductID, dbOpenDynaset, dbSeeChanges)
        cur.Edit
        newAvg = cur!AverageCost
        If kind = "RECEIPT" Then newAvg = WeightedAverage(cur!CurrentQuantity, cur!AverageCost, -rs!Quantity, rs!UnitCost)
        cur!AverageCost = newAvg
        cur!CurrentQuantity = cur!CurrentQuantity - rs!Quantity
        cur!UpdatedAt = Now
        cur.Update
        cur.Close
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "DELETE FROM InventoryTransactions WHERE ReferenceType = 'STOCK_NOTE' AND ReferenceID = " & StockNoteID, _
               dbFailOnError + dbSeeChanges
    db.Execute "UPDATE StockNotes SET Status = 'DRAFT' WHERE StockNoteID = " & StockNoteID, dbFailOnError + dbSeeChanges
    If NoteOrderID(StockNoteID) > 0 Then SyncSalesOrder db, NoteOrderID(StockNoteID)   ' reserved again
    ws.CommitTrans
    inTrans = False
    AuditEdited "STOCK_NOTE_UNPOST", "StockNotes", "StockNoteID", StockNoteID, before
    Exit Function
EH:
    UnpostStockNote = " ⁄–— ≈·€«¡ «· —ÕÌ·: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function DeleteStockNote(ByVal StockNoteID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("STOCK_NOTES") Then
        DeleteStockNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    If NoteStatus(StockNoteID) <> "DRAFT" Then
        DeleteStockNote = " ıÕ–› «·„”Êœ… ›ﬁÿ. √·€ˆ  —ÕÌ· «·≈–‰ √Ê·«."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("StockNotes", "StockNoteID", StockNoteID)
    CurrentDb.Execute "DELETE FROM StockNoteLines WHERE StockNoteID = " & StockNoteID, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM StockNotes WHERE StockNoteID = " & StockNoteID, dbFailOnError + dbSeeChanges
    AuditDeleted "STOCK_NOTE_DELETE", "StockNotes", StockNoteID, auditBefore
End Function

'==============================================================================
' The invoice of a note
'==============================================================================
Public Function NoteInvoiceProblem(ByVal StockNoteID As Long, ByVal NoteType As String, ByVal PartyID As Long, _
                                   ByRef ProductIDs() As Long, ByRef Costs() As Currency, ByVal n As Long, _
                                   ByRef WarehouseID As Long) As String
    ' Called by PostSaleFromCart / PostPurchaseFromCart: the note is posted, of this type and party, and the cart
    ' holds its lines in their order. Gives the cost of each line (that of the note) and the warehouse.
    Dim rs As DAO.Recordset, i As Long
    If NoteStatus(StockNoteID) <> "POSTED" Then
        NoteInvoiceProblem = "«·≈–‰ €Ì— „ı—ÕÛ¯· √Ê ’œ—  ›« Ê— Â."
        Exit Function
    End If
    If NoteKind(StockNoteID) <> NoteType Or NoteParty(StockNoteID) <> PartyID Then
        NoteInvoiceProblem = "«·›« Ê—… ·Ì”  ·‰›” ‰Ê⁄ «·≈–‰ Êÿ—›Â."
        Exit Function
    End If
    If n > 0 Then ReDim Preserve Costs(0 To n - 1)
    Set rs = CurrentDb.OpenRecordset("SELECT ProductID, UnitCost FROM StockNoteLines WHERE StockNoteID = " & _
                                     StockNoteID & " ORDER BY StockNoteLineID", dbOpenSnapshot)
    Do Until rs.EOF
        If i >= n Then Exit Do
        If rs!ProductID <> ProductIDs(i) Then Exit Do
        Costs(i) = Nz(rs!UnitCost, 0)
        i = i + 1
        rs.MoveNext
    Loop
    If Not rs.EOF Or i <> n Then NoteInvoiceProblem = "√”ÿ— «·›« Ê—… ·Ì”  √”ÿ— «·≈–‰."
    rs.Close
    WarehouseID = DbValue("SELECT WarehouseID FROM StockNotes WHERE StockNoteID = " & StockNoteID)
End Function

Public Sub MarkNoteInvoiced(ByVal db As DAO.Database, ByVal StockNoteID As Long, ByVal FieldName As String, _
                            ByVal InvoiceID As Long)
    ' Inside the transaction of the invoice: a note gets one invoice only.
    db.Execute "UPDATE StockNotes SET Status = 'INVOICED', " & FieldName & " = " & InvoiceID & _
               " WHERE Status = 'POSTED' AND StockNoteID = " & StockNoteID, dbFailOnError + dbSeeChanges
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 610, "MarkNoteInvoiced", "«·≈–‰ ·„ Ì⁄œ „ı—ÕÛ¯·«"
    If NoteOrderID(StockNoteID) > 0 Then SyncSalesOrder db, NoteOrderID(StockNoteID)   ' the order may be invoiced now
End Sub

Public Function InvoiceStockNote(ByVal StockNoteID As Long, ByVal PaymentType As String, ByVal SupplierInvoiceNo As String, _
                                 ByVal ChargeVat As Boolean, ByRef NewInvoiceID As Long) As String
    ' The sales invoice of a delivery note, or the purchase invoice of a receipt note, at the prices of the note.
    Dim db As DAO.Database, rs As DAO.Recordset, kind As String, party As Long, noteNo As String, msg As String
    Dim cartTable As String, method As Variant, cart As DAO.Recordset
    NewInvoiceID = 0
    If Not HasPermission("STOCK_NOTES") Then
        InvoiceStockNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    If NoteStatus(StockNoteID) <> "POSTED" Then
        InvoiceStockNote = " ’œ— «·›« Ê—… „‰ ≈–‰ „ı—ÕÛ¯· ·„  ’œ— ›« Ê— Â."
        Exit Function
    End If
    If PaymentType <> "CASH" And PaymentType <> "CREDIT" Then
        InvoiceStockNote = "«Œ — œ›⁄ «·›« Ê—…: ‰ﬁœÌ √Ê ¬Ã·."
        Exit Function
    End If
    kind = NoteKind(StockNoteID)
    party = NoteParty(StockNoteID)
    noteNo = DbValue("SELECT NoteNumber FROM StockNotes WHERE StockNoteID = " & StockNoteID)
    cartTable = IIf(kind = "DELIVERY", "tmpPOSLines", "tmpPurchaseLines")
    EnsureLocalTables                                               ' modPOS
    If Nz(DbValue("SELECT COUNT(*) FROM " & cartTable), 0) > 0 Then
        InvoiceStockNote = IIf(kind = "DELIVERY", "›Ì ‘«‘… «·»Ì⁄ ›« Ê—… ·„  ıÕ›Ÿ »⁄œ. «Õ›ŸÂ« √Ê «„”ÕÂ« √Ê·«.", _
                               "›Ì ‘«‘… ›« Ê—… «·„‘ —Ì«  ›« Ê—… ·„  ıÕ›Ÿ »⁄œ. «Õ›ŸÂ« √Ê «„”ÕÂ« √Ê·«.")
        Exit Function
    End If
    If PaymentType = "CASH" Then method = CASH_METHOD_ID Else method = Null

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT l.ProductID, l.Quantity, l.UnitPrice, p.ProductCode, p.ProductName, p.SellingPrice, " & _
                              "p.CurrentQuantity FROM StockNoteLines AS l INNER JOIN Products AS p ON l.ProductID = p.ProductID " & _
                              "WHERE l.StockNoteID = " & StockNoteID & " ORDER BY l.StockNoteLineID", dbOpenSnapshot)
    Set cart = db.OpenRecordset(cartTable, dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    Do Until rs.EOF                                                 ' one cart line per note line, in order
        cart.AddNew
        cart!ProductID = rs!ProductID
        cart!ProductCode = rs!ProductCode
        cart!ProductName = rs!ProductName
        cart!Quantity = rs!Quantity
        cart!LineDiscount = 0                                       ' local tables have no defaults
        If kind = "DELIVERY" Then
            cart!UnitPrice = Nz(rs!UnitPrice, 0)
            cart!Available = Nz(rs!CurrentQuantity, 0)
        Else
            cart!UnitCost = Nz(rs!UnitPrice, 0)
            cart!SellingPrice = Nz(rs!SellingPrice, 0)
        End If
        cart.Update
        rs.MoveNext
    Loop
    cart.Close
    rs.Close
    If kind = "DELIVERY" Then
        msg = PostSaleFromCart(party, PaymentType, method, 0, Null, "≈–‰ «· ”·Ì„ " & noteNo, NewInvoiceID, _
                               StockNoteID:=StockNoteID)
    Else
        msg = PostPurchaseFromCart(party, SupplierInvoiceNo, Date, PaymentType, method, ChargeVat, 0, Null, _
                                   "≈–‰ «·«” ·«„ " & noteNo, NewInvoiceID, StockNoteID:=StockNoteID)
    End If
    db.Execute "DELETE FROM " & cartTable, dbFailOnError + dbSeeChanges            ' the cart never keeps the lines of a note
    InvoiceStockNote = msg
    If Len(msg) = 0 Then LogAction "STOCK_NOTE_INVOICE", "StockNotes", noteNo, "Invoice=" & NewInvoiceID
End Function

'==============================================================================
' frmStockNote (OpenArgs = a note to show) and its lines frmStockNoteLines
'==============================================================================
Public Sub StockNoteLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboWarehouse.RowSource = WarehouseRows()
    frm!cboWarehouse.Value = CurrentWarehouseID()
    frm!cboType.Value = "DELIVERY"
    frm!cboPayType.Value = "CREDIT"
    frm!chkVat.Value = True
    StockNoteTypeChanged frm
    StockNoteListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(StockNoteID) FROM StockNotes")
    frm!cboNote.Value = id
    StockNoteShow frm
End Sub

Public Sub StockNoteTypeChanged(ByVal frm As Access.Form)
    frm!cboParty.RowSource = NotePartyRows(Nz(frm!cboType.Value, "DELIVERY"))
    frm!cboParty.Value = Null
End Sub

Private Sub StockNoteListRefresh(ByVal frm As Access.Form)
    frm!cboNote.RowSource = Tr("SELECT n.StockNoteID, n.NoteNumber, Format(n.NoteDate, 'yyyy/mm/dd'), " & _
        "IIf(n.NoteType = 'DELIVERY', ' ”·Ì„', '«” ·«„'), IIf(n.NoteType = 'DELIVERY', c.CustomerName, s.SupplierName), " & _
        "w.WarehouseName, IIf(n.Status = 'DRAFT', '„”Êœ…', IIf(n.Status = 'POSTED', '„ı—ÕÛ¯·', '’œ—  ›« Ê— Â')) " & _
        "FROM ((StockNotes AS n INNER JOIN [@Warehouses] AS w ON n.WarehouseID = w.WarehouseID) " & _
        "LEFT JOIN [@Customers] AS c ON n.CustomerID = c.CustomerID) LEFT JOIN [@Suppliers] AS s " & _
        "ON n.SupplierID = s.SupplierID ORDER BY n.StockNoteID DESC")
    frm!cboNote.Requery
End Sub

Public Sub StockNoteShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, rs As DAO.Recordset, info As String, kind As String, invNo As Variant
    id = Nz(frm!cboNote.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM StockNoteLines WHERE StockNoteID = " & id & " ORDER BY StockNoteLineID"
    st = NoteStatus(id)
    kind = NoteKind(id)
    If Len(st) = 0 Then
        info = "·« ÌÊÃœ ≈–‰. «Œ — «·‰Ê⁄ Ê«·„Œ“‰ Ê«·⁄„Ì· √Ê «·„Ê—œ À„ ´≈–‰ ÃœÌœª."
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT NoteNumber, NoteDate, WarehouseID, SalesInvoiceID, PurchaseInvoiceID " & _
                                         "FROM StockNotes WHERE StockNoteID = " & id, dbOpenSnapshot)
        info = rs!NoteNumber & "    " & IIf(kind = "DELIVERY", " ”·Ì„ „‰ ", "«” ·«„ ›Ì ") & WarehouseName(rs!WarehouseID) & _
               IIf(kind = "DELIVERY", "  ->  ", "  <-  ") & NotePartyName(id) & "    "
        If NoteOrderID(id) > 0 Then
            info = info & "«·ÿ·»Ì… " & DbValue("SELECT OrderNumber FROM SalesOrders WHERE SalesOrderID = " & NoteOrderID(id)) & "    "
        End If
        If st = "DRAFT" Then
            info = info & "„”Êœ…: √÷› «·√’‰«› À„ ´ —ÕÌ·ª"
        ElseIf st = "POSTED" Then
            info = info & "„ı—ÕÛ¯· ›Ì " & GDate(rs!NoteDate) & ": √‰‘∆ «·›« Ê—…"
        Else
            If kind = "DELIVERY" Then
                invNo = DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & Nz(rs!SalesInvoiceID, 0))
            Else
                invNo = DbValue("SELECT InvoiceNumber FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & _
                                Nz(rs!PurchaseInvoiceID, 0))
            End If
            info = info & "’œ—  ›« Ê— Â " & Nz(invNo, "")
        End If
        rs.Close
        info = info & "    «· ﬂ·›…: " & Format$(Nz(DbValue("SELECT NoteValue FROM qryStockNoteValue WHERE StockNoteID = " & _
                                                           id), 0), "#,##0.00")
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    SafeFocus frm!cboNote                   ' a focused control cannot be disabled
    frm!btnPost.Enabled = (st = "DRAFT")
    frm!btnDelete.Enabled = (st = "DRAFT")
    frm!btnUnpost.Enabled = (st = "POSTED")
    frm!btnInvoice.Enabled = (st = "POSTED")
    frm!btnOrder.Enabled = (NoteOrderID(id) > 0)
    frm!txtSupplierRef.Enabled = (st = "POSTED" And kind = "RECEIPT")
    frm!chkVat.Enabled = (st = "POSTED" And kind = "RECEIPT")
End Sub

Public Sub StockNoteNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateStockNote(Nz(frm!cboType.Value, ""), frm!cboWarehouse.Value, frm!cboParty.Value, _
                          Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    StockNoteListRefresh frm
    frm!cboNote.Value = id
    StockNoteShow frm
End Sub

Public Sub StockNoteAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String, invoiceID As Long
    id = Nz(frm!cboNote.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "POST"
            If Not AskYesNo(" —ÕÌ· «·≈–‰ø Ì Õ—ﬂ «·„Œ“Ê‰ ÊÌı‰‘√ «·ﬁÌœ.") Then Exit Sub
            msg = PostStockNote(id)
        Case "UNPOST"
            If Not AskYesNo("≈·€«¡  —ÕÌ· «·≈–‰ Ê≈⁄«œ… «·„Œ“Ê‰ ﬂ„« ﬂ«‰ø") Then Exit Sub
            msg = UnpostStockNote(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› „”Êœ… «·≈–‰ø") Then Exit Sub
            msg = DeleteStockNote(id)
        Case "INVOICE"
            If Not AskYesNo("≈‰‘«¡ «·›« Ê—… „‰ Â–« «·≈–‰ »√”⁄«—Âø ·« Ì Õ—ﬂ «·„Œ“Ê‰ „—… À«‰Ì….") Then Exit Sub
            msg = InvoiceStockNote(id, Nz(frm!cboPayType.Value, ""), Nz(frm!txtSupplierRef.Value, ""), _
                                   Nz(frm!chkVat.Value, True), invoiceID)
            If Len(msg) = 0 Then frm!txtSupplierRef.Value = Null
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    StockNoteListRefresh frm
    If Action = "DELETE" Then frm!cboNote.Value = DbValue("SELECT Max(StockNoteID) FROM StockNotes")
    StockNoteShow frm
End Sub

Public Sub StockNoteLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String
    id = Nz(frm.Parent!cboNote.Value, 0)
    If NoteStatus(id) <> "DRAFT" Then
        msg = " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
    Else
        msg = StockNoteLineProblem(frm!ProductID.Value, frm!Quantity.Value, frm!UnitPrice.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!StockNoteID.Value = id
    If Nz(frm!UnitPrice.Value, 0) = 0 Then
        frm!UnitPrice.Value = NoteDefaultPrice(NoteKind(id), CLng(frm!ProductID.Value))
        If NoteKind(id) = "DELIVERY" Then           ' the customer's price list (modPriceLists)
            frm!UnitPrice.Value = CustomerPrice(NoteParty(id), CLng(frm!ProductID.Value), Null, frm!UnitPrice.Value)
        End If
    End If
    AuditFormBefore frm, "StockNoteLines", "StockNoteLineID"
End Sub

Public Sub StockNoteLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    StockNoteShow frm.Parent
End Sub

Public Sub StockNoteRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!StockNoteLineID.Value
    If IsNull(id) Then Exit Sub
    If NoteStatus(Nz(frm.Parent!cboNote.Value, 0)) <> "DRAFT" Then
        ShowWarning " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("StockNoteLines", "StockNoteLineID", id)
    CurrentDb.Execute "DELETE FROM StockNoteLines WHERE StockNoteLineID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "STOCK_NOTE_LINE_DELETE", "StockNoteLines", id, auditBefore
    frm.Requery
    StockNoteShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestStockNotes() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, refNo As String, msg As String
    Dim passed As Long, failed As Long, report As String, noteID As Long, invID As Long, cust As Long, sup As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestStockNotes  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-SN-P', 'TEST ’‰› «·√–Ê‰« ', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-SN-P'")
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive <> 0"), 0)
    sup = Nz(DbValue("SELECT Min(SupplierID) FROM Suppliers WHERE IsActive <> 0"), 0)
    CheckSn cust > 0 And sup > 0, "ÌÊÃœ ⁄„Ì· Ê„Ê—œ ‰‘ÿ«‰", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckSn Len(msg) = 0, "≈÷«›… 10 » ﬂ·›… 6 " & msg, passed, failed, report

    ' a receipt note at the purchase price, then its purchase invoice: no second stock move
    msg = CreateStockNote("RECEIPT", MAIN_WAREHOUSE_ID, sup, "TEST", noteID)
    CurrentDb.Execute "INSERT INTO StockNoteLines (StockNoteID, ProductID, Quantity, UnitPrice) VALUES (" & _
                      noteID & ", " & pid & ", 10, 10)", dbFailOnError + dbSeeChanges
    msg = msg & PostStockNote(noteID)
    CheckSn Len(msg) = 0 And NoteStatus(noteID) = "POSTED" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 20 And _
            Nz(DbValue("SELECT AverageCost FROM Products WHERE ProductID = " & pid), 0) = 8, _
            "≈–‰ «·«” ·«„ Ì÷Ì› ÊÌ⁄Ìœ Õ”«» «·„ Ê”ÿ " & msg, passed, failed, report
    CheckSn Nz(DbValue("SELECT Sum(Credit) FROM qryJournalStockNote WHERE SourceType = 'RECEIPT_NOTE' AND SourceID = " & _
                       noteID & " AND AccountCode = 2150"), 0) = 100, "ﬁÌœ ≈–‰ «·«” ·«„ ⁄·Ï 2150", passed, failed, report
    msg = InvoiceStockNote(noteID, "CREDIT", "TEST-SN-1", True, invID)
    CheckSn Len(msg) = 0 And NoteStatus(noteID) = "INVOICED" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 20 And _
            Nz(DbValue("SELECT TaxableAmount FROM PurchaseInvoices WHERE PurchaseInvoiceID = " & invID), 0) = 100, _
            "›« Ê—… «·‘—«¡ „‰ «·≈–‰ ·«  ÷Ì› „—… À«‰Ì… " & msg, passed, failed, report
    CheckSn Nz(DbValue("SELECT Sum(Debit) FROM qryJournalPurchase WHERE SourceID = " & invID & " AND AccountCode = 2150"), 0) = 100 _
            And Nz(DbValue("SELECT Count(*) FROM qryJournalPurchase WHERE SourceID = " & invID & " AND AccountCode = 1400"), 0) = 0, _
            "ﬁÌœ ›« Ê—… «·‘—«¡ Ìﬁ›· 2150", passed, failed, report
    CheckSn Len(UnpostStockNote(noteID)) > 0, "«·≈–‰ «·–Ì ’œ—  ›« Ê— Â ·« Ìı·€Ï  —ÕÌ·Â", passed, failed, report
    CheckSn Len(InvoiceStockNote(noteID, "CREDIT", "TEST-SN-2", True, invID)) > 0, "›« Ê—… Ê«Õœ… ·ﬂ· ≈–‰", _
            passed, failed, report

    ' a delivery note at the average cost, then its sales invoice
    msg = CreateStockNote("DELIVERY", MAIN_WAREHOUSE_ID, cust, "TEST", noteID)
    CurrentDb.Execute "INSERT INTO StockNoteLines (StockNoteID, ProductID, Quantity, UnitPrice) VALUES (" & _
                      noteID & ", " & pid & ", 5, 20)", dbFailOnError + dbSeeChanges
    msg = msg & PostStockNote(noteID)
    CheckSn Len(msg) = 0 And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 15 And _
            Nz(DbValue("SELECT NoteValue FROM qryStockNoteValue WHERE StockNoteID = " & noteID), 0) = 40, _
            "≈–‰ «· ”·Ì„ ÌŒ’„ »„ Ê”ÿ «· ﬂ·›… " & msg, passed, failed, report
    msg = UnpostStockNote(noteID)
    CheckSn Len(msg) = 0 And NoteStatus(noteID) = "DRAFT" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 20, _
            "≈·€«¡  —ÕÌ· ≈–‰ «· ”·Ì„ " & msg, passed, failed, report
    msg = PostStockNote(noteID) & InvoiceStockNote(noteID, "CASH", "", True, invID)
    CheckSn Len(msg) = 0 And NoteStatus(noteID) = "INVOICED" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 15 And _
            Nz(DbValue("SELECT Sum(Quantity * UnitCost) FROM SalesInvoiceDetails WHERE SalesInvoiceID = " & invID), 0) = 40, _
            "›« Ê—… «·»Ì⁄ „‰ «·≈–‰ » ﬂ·› Â Ê·«  Œ’„ „—… À«‰Ì… " & msg, passed, failed, report
    CheckSn Nz(DbValue("SELECT Sum(Credit) FROM qryJournalSale WHERE SourceID = " & invID & " AND AccountCode = 1430"), 0) = 40, _
            "ﬁÌœ ›« Ê—… «·»Ì⁄ Ìﬁ›· 1430", passed, failed, report
    CheckSn Nz(DbValue("SELECT Count(*) FROM tmpPOSLines"), 0) = 0, "”·… «·»Ì⁄ ›«—€… »⁄œ «·›« Ê—…", passed, failed, report

    msg = CreateStockNote("DELIVERY", MAIN_WAREHOUSE_ID, cust, "TEST", noteID)
    CheckSn Len(DeleteStockNote(noteID)) = 0 And NoteStatus(noteID) = "", "Õ–› «·„”Êœ…", passed, failed, report
    CheckSn Len(CreateStockNote("DELIVERY", MAIN_WAREHOUSE_ID, Null, "TEST", noteID)) > 0, "«·≈–‰ ÌÕ «Ã ⁄„Ì·«", _
            passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckSn False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, _
                "TestStockNotes"
        TestStockNotes = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestStockNotes"
    End If
End Function

Private Sub CheckSn(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
