Attribute VB_Name = "modPurchaseOrders"
'==============================================================================
' modPurchaseOrders  -  Retail Store Management System (purchase orders and the purchase suggestion, docs/74)
'
' Module 4 of docs/69: the customer orders of docs/62, for suppliers. A purchase order is entered as a draft
' (by hand or from the purchase suggestion), approved (its quantities are then "on order"), received with
' receipt notes (modStockNotes, docs/61) made from it, in one go or in parts, and invoiced from those notes.
'
'   CreatePurchaseOrder           a draft for a supplier and the warehouse that receives
'   ApprovePurchaseOrder          DRAFT -> APPROVED: what is left of it is on order (qryOnOrderStock)
'   CreateOrderReceiptNote        a draft receipt note with what is left of each line (PurchaseOrderLineID)
'   SyncPurchaseOrder             after a note of the order is posted, unposted or invoiced: APPROVED (nothing
'                                 received), PARTIAL, RECEIVED (all received), INVOICED (all invoiced)
'   UnapprovePurchaseOrder        back to a draft while it has no receipt note
'   CancelPurchaseOrder           CANCELLED, or CLOSED when part of it was received
'   DeletePurchaseOrder           a draft only
' The purchase suggestion (PurchaseSuggestQuery): the active products that follow their stock, with a reorder
' level (Products.MinimumQuantity), whose stock (all warehouses) plus what is on order is at or below it.
' Suggested quantity = Products.ReorderQuantity, else what brings the stock to twice the level.
'   CreateSuggestedOrders         one draft order per supplier (Products.SupplierID) with the suggested lines
' Screen: frmPurchaseOrder (built by modBuildForms); reports PURCHASE_SUGGEST and PENDING_PURCHASE_ORDERS.
'==============================================================================
Option Compare Database
Option Explicit

Private Function PoStatus(ByVal PurchaseOrderID As Long) As String
    PoStatus = Nz(DbValue("SELECT Status FROM PurchaseOrders WHERE PurchaseOrderID = " & PurchaseOrderID), "")
End Function

Private Function PoIsOpen(ByVal Status As String) As Boolean
    PoIsOpen = (Status = "APPROVED" Or Status = "PARTIAL")
End Function

Private Function PoNoteCount(ByVal PurchaseOrderID As Long, Optional ByVal OnlyDrafts As Boolean = False) As Long
    Dim sql As String
    sql = "SELECT COUNT(*) FROM StockNotes WHERE PurchaseOrderID = " & PurchaseOrderID
    If OnlyDrafts Then sql = sql & " AND Status = 'DRAFT'"
    PoNoteCount = Nz(DbValue(sql), 0)
End Function

Private Function PoReceived(ByVal PurchaseOrderID As Long) As Currency
    PoReceived = Nz(DbValue("SELECT Sum(ReceivedQty) FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderID = " & _
                            PurchaseOrderID), 0)
End Function

Public Function NotePurchaseOrderID(ByVal StockNoteID As Long) As Long
    ' The purchase order a receipt note was made from, else 0.
    NotePurchaseOrderID = Nz(DbValue("SELECT PurchaseOrderID FROM StockNotes WHERE StockNoteID = " & StockNoteID), 0)
End Function

Public Function OnOrderQty(ByVal ProductID As Long) As Currency
    ' What the open purchase orders still bring of a product.
    OnOrderQty = Nz(DbValue("SELECT OpenQty FROM qryOnOrderStock WHERE ProductID = " & ProductID), 0)
End Function

Public Function PurchaseOrderLineReceived(ByVal PurchaseOrderLineID As Variant) As Variant
    ' The received column of the order lines grid.
    PurchaseOrderLineReceived = Null
    If IsNull(PurchaseOrderLineID) Then Exit Function
    PurchaseOrderLineReceived = Nz(DbValue("SELECT ReceivedQty FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderLineID = " & _
                                           CLng(PurchaseOrderLineID)), 0)
End Function

Public Function PurchaseOrderLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant, ByVal UnitCost As Variant) As String
    If IsNull(ProductID) Then
        PurchaseOrderLineProblem = "«Œ — «·’‰›."
    ElseIf Not ProductTracksStock(CLng(ProductID)) Then
        PurchaseOrderLineProblem = "Â–« «·’‰› ·« Ì «»⁄ «·„Œ“Ê‰° ›·« Ìı” ·„ »≈–‰."
    ElseIf Nz(Qty, 0) <= 0 Then
        PurchaseOrderLineProblem = "«·ﬂ„Ì… √ﬂ»— „‰ ’›—."
    ElseIf Nz(UnitCost, 0) < 0 Then
        PurchaseOrderLineProblem = "«·”⁄— ·« ÌﬂÊ‰ ”«·»«."
    End If
End Function

Public Function SuggestQty(ByVal CurrentQty As Currency, ByVal OnOrder As Currency, ByVal MinimumQty As Currency, _
                           ByVal ReorderQty As Currency) As Currency
    ' The quantity PurchaseSuggestQuery suggests (0 = the product is above its reorder level).
    If MinimumQty <= 0 Or CurrentQty + OnOrder > MinimumQty Then Exit Function
    If ReorderQty > 0 Then
        SuggestQty = ReorderQty
    Else
        SuggestQty = 2 * MinimumQty - CurrentQty - OnOrder
    End If
End Function

'==============================================================================
' The way of an order
'==============================================================================
Public Function CreatePurchaseOrder(ByVal SupplierID As Variant, ByVal WarehouseID As Variant, ByVal ExpectedDate As Variant, _
                                   ByVal Notes As String, ByRef NewOrderID As Long) As String
    Dim rs As DAO.Recordset
    NewOrderID = 0
    If Not HasPermission("PURCHASE_ORDERS") Then
        CreatePurchaseOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
    ElseIf Not NotePartyIsActive("RECEIPT", SupplierID) Then
        CreatePurchaseOrder = "«Œ — «·„Ê—œ."
    ElseIf Not WarehouseIsActive(WarehouseID) Then
        CreatePurchaseOrder = "«Œ — «·„Œ“‰."
    End If
    If Len(CreatePurchaseOrder) = 0 And Not IsNull(ExpectedDate) Then
        If Not IsDate(ExpectedDate) Then CreatePurchaseOrder = "„Ê⁄œ «·«” ·«„ ·Ì”  «—ÌŒ«."
    End If
    If Len(CreatePurchaseOrder) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("PurchaseOrders", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    rs.AddNew
    rs!OrderNumber = NextNumber("PURCHASE_ORDER")
    rs!OrderDate = Now
    rs!SupplierID = CLng(SupplierID)
    rs!WarehouseID = CLng(WarehouseID)
    If Not IsNull(ExpectedDate) Then rs!ExpectedDate = DateValue(CDate(ExpectedDate))
    rs!Status = "DRAFT"
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewOrderID = rs!PurchaseOrderID
    rs.Close
    LogAction "PURCHASE_ORDER_NEW", "PurchaseOrders", CStr(NewOrderID)
End Function

Public Function ApprovePurchaseOrder(ByVal PurchaseOrderID As Long) As String
    Dim rs As DAO.Recordset, msg As String, n As Long
    If Not HasPermission("PURCHASE_ORDERS") Then
        ApprovePurchaseOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
        Exit Function
    End If
    If PoStatus(PurchaseOrderID) <> "DRAFT" Then
        ApprovePurchaseOrder = "Ìı⁄ „œ √„— «·‘—«¡ ÊÂÊ „”Êœ…."
        Exit Function
    End If
    If Not NotePartyIsActive("RECEIPT", DbValue("SELECT SupplierID FROM PurchaseOrders WHERE PurchaseOrderID = " & _
                                                PurchaseOrderID)) Then
        ApprovePurchaseOrder = "«·„Ê—œ €Ì— „ÊÃÊœ √Ê €Ì— ‰‘ÿ."
        Exit Function
    End If
    msg = WarehouseProblem(DbValue("SELECT WarehouseID FROM PurchaseOrders WHERE PurchaseOrderID = " & PurchaseOrderID))
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity, UnitCost FROM PurchaseOrderLines WHERE PurchaseOrderID = " & _
                                         PurchaseOrderID, dbOpenSnapshot)
        Do Until rs.EOF
            If Len(msg) = 0 Then msg = PurchaseOrderLineProblem(rs!ProductID, rs!Quantity, rs!UnitCost)
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) = 0 And n = 0 Then msg = "√÷› «·√’‰«› Êﬂ„Ì« Â« √Ê·«."
    End If
    If Len(msg) > 0 Then
        ApprovePurchaseOrder = msg
        Exit Function
    End If
    CurrentDb.Execute "UPDATE PurchaseOrders SET Status = 'APPROVED', ApprovedAt = Now() WHERE PurchaseOrderID = " & _
                      PurchaseOrderID, dbFailOnError + dbSeeChanges
    LogAction "PURCHASE_ORDER_APPROVE", "PurchaseOrders", CStr(PurchaseOrderID)
End Function

Public Function UnapprovePurchaseOrder(ByVal PurchaseOrderID As Long) As String
    ' Back to a draft (its lines can be changed again) while no receipt note was made from it.
    Dim before As Collection
    If Not HasPermission("PURCHASE_ORDERS") Then
        UnapprovePurchaseOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
    ElseIf PoStatus(PurchaseOrderID) <> "APPROVED" Then
        UnapprovePurchaseOrder = "Ì⁄Êœ ··„”Êœ… √„— «·‘—«¡ «·„⁄ „œ."
    ElseIf PoNoteCount(PurchaseOrderID) > 0 Then
        UnapprovePurchaseOrder = "’œ—  „‰ «·√„— √–Ê‰«  «” ·«„. «Õ–› «·„”Êœ«  „‰Â«° √Ê √·€ˆ «·√„—."
    End If
    If Len(UnapprovePurchaseOrder) > 0 Then Exit Function
    Set before = AuditSnapshot("PurchaseOrders", "PurchaseOrderID", PurchaseOrderID)
    CurrentDb.Execute "UPDATE PurchaseOrders SET Status = 'DRAFT', ApprovedAt = Null WHERE PurchaseOrderID = " & _
                      PurchaseOrderID, dbFailOnError + dbSeeChanges
    AuditEdited "PURCHASE_ORDER_UNAPPROVE", "PurchaseOrders", "PurchaseOrderID", PurchaseOrderID, before
End Function

Public Function CancelPurchaseOrder(ByVal PurchaseOrderID As Long) As String
    ' What is left is no longer on order. An order received in part is CLOSED: its notes stay as they are.
    Dim before As Collection, st As String, newStatus As String
    st = PoStatus(PurchaseOrderID)
    If Not HasPermission("PURCHASE_ORDERS") Then
        CancelPurchaseOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
    ElseIf st <> "DRAFT" And Not PoIsOpen(st) Then
        CancelPurchaseOrder = "√„— «·‘—«¡ „” ·„ ﬂ·Â √Ê „·€Ï „‰ ﬁ»·."
    ElseIf PoNoteCount(PurchaseOrderID, True) > 0 Then
        CancelPurchaseOrder = "··√„— ≈–‰ «” ·«„ ·„ Ìı—ÕÛ¯·. —Õ¯·Â √Ê «Õ–›Â √Ê·«."
    End If
    If Len(CancelPurchaseOrder) > 0 Then Exit Function
    newStatus = "CANCELLED"
    If PoReceived(PurchaseOrderID) > 0 Then newStatus = "CLOSED"
    Set before = AuditSnapshot("PurchaseOrders", "PurchaseOrderID", PurchaseOrderID)
    CurrentDb.Execute "UPDATE PurchaseOrders SET Status = '" & newStatus & "' WHERE PurchaseOrderID = " & PurchaseOrderID, _
                      dbFailOnError + dbSeeChanges
    AuditEdited "PURCHASE_ORDER_CANCEL", "PurchaseOrders", "PurchaseOrderID", PurchaseOrderID, before
End Function

Public Function DeletePurchaseOrder(ByVal PurchaseOrderID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("PURCHASE_ORDERS") Then
        DeletePurchaseOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
        Exit Function
    End If
    If PoStatus(PurchaseOrderID) <> "DRAFT" Or PoNoteCount(PurchaseOrderID) > 0 Then
        DeletePurchaseOrder = " ıÕ–› «·„”Êœ… ›ﬁÿ. √·€ˆ √„— «·‘—«¡ «·„⁄ „œ »œ· Õ–›Â."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("PurchaseOrders", "PurchaseOrderID", PurchaseOrderID)
    CurrentDb.Execute "DELETE FROM PurchaseOrderLines WHERE PurchaseOrderID = " & PurchaseOrderID, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM PurchaseOrders WHERE PurchaseOrderID = " & PurchaseOrderID, dbFailOnError + dbSeeChanges
    AuditDeleted "PURCHASE_ORDER_DELETE", "PurchaseOrders", PurchaseOrderID, auditBefore
End Function

'==============================================================================
' Receiving: the notes of an order
'==============================================================================
Public Function CreateOrderReceiptNote(ByVal PurchaseOrderID As Long, ByRef NewNoteID As Long) As String
    ' A draft receipt note in the warehouse of the order with what is left of each line, at the order price.
    ' The user may lower the quantities (a partial receipt) before posting it.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, nl As DAO.Recordset, inTrans As Boolean
    Dim msg As String, o As DAO.Recordset, orderNo As String, sup As Long, wh As Long
    NewNoteID = 0
    If Not HasPermission("STOCK_NOTES") Then
        CreateOrderReceiptNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    If Not PoIsOpen(PoStatus(PurchaseOrderID)) Then
        CreateOrderReceiptNote = "Ìı‰‘√ ≈–‰ «·«” ·«„ „‰ √„— ‘—«¡ „⁄ „œ ·„ Ìı” ·„ ﬂ·Â."
        Exit Function
    End If
    If PoNoteCount(PurchaseOrderID, True) > 0 Then
        CreateOrderReceiptNote = "··√„— ≈–‰ «” ·«„ ·„ Ìı—ÕÛ¯·. «› ÕÂ „‰ ´√–Ê‰«  «·√„—ª."
        Exit Function
    End If
    If Nz(DbValue("SELECT Sum(RemainingQty) FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderID = " & PurchaseOrderID), 0) <= 0 Then
        CreateOrderReceiptNote = "«” ı·„  ﬂ· √’‰«› «·√„—."
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    Set o = db.OpenRecordset("SELECT OrderNumber, SupplierID, WarehouseID FROM PurchaseOrders WHERE PurchaseOrderID = " & _
                             PurchaseOrderID, dbOpenSnapshot)
    orderNo = o!OrderNumber
    sup = o!SupplierID
    wh = o!WarehouseID
    o.Close                                  ' closed before BeginTrans (error 3246 inside it)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateStockNote("RECEIPT", wh, sup, "√„— «·‘—«¡ " & orderNo, NewNoteID)   ' modStockNotes
    If Len(msg) = 0 Then
        db.Execute "UPDATE StockNotes SET PurchaseOrderID = " & PurchaseOrderID & " WHERE StockNoteID = " & NewNoteID, _
                   dbFailOnError + dbSeeChanges
        Set rs = db.OpenRecordset("SELECT PurchaseOrderLineID, ProductID, RemainingQty, UnitCost FROM qryPurchaseOrderLineStatus " & _
                                  "WHERE PurchaseOrderID = " & PurchaseOrderID & " AND RemainingQty > 0 " & _
                                  "ORDER BY PurchaseOrderLineID", dbOpenSnapshot, dbSeeChanges)
        Set nl = db.OpenRecordset("StockNoteLines", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
        Do Until rs.EOF
            nl.AddNew
            nl!StockNoteID = NewNoteID
            nl!ProductID = rs!ProductID
            nl!Quantity = rs!RemainingQty
            nl!UnitPrice = rs!UnitCost
            nl!UnitCost = 0
            nl!PurchaseOrderLineID = rs!PurchaseOrderLineID
            nl.Update
            rs.MoveNext
        Loop
        nl.Close
        rs.Close
    End If
    If Len(msg) > 0 Then
        ws.Rollback
        NewNoteID = 0
        CreateOrderReceiptNote = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "PURCHASE_ORDER_NOTE", "PurchaseOrders", CStr(PurchaseOrderID), "Note=" & NewNoteID
    Exit Function
EH:
    CreateOrderReceiptNote = " ⁄–— ≈‰‘«¡ ≈–‰ «·«” ·«„: " & Err.Description & " (" & Err.Number & ")"
    NewNoteID = 0
    If inTrans Then ws.Rollback
End Function

Public Function PurchaseOrderNoteProblem(ByVal StockNoteID As Long) As String
    ' Before a receipt note of an order is posted: the order is open, the note is for its supplier and warehouse,
    ' each line receives the product of its order line and not more than what is left of it.
    Dim rs As DAO.Recordset, ln As DAO.Recordset, orderID As Long, o As DAO.Recordset, note As DAO.Recordset
    orderID = NotePurchaseOrderID(StockNoteID)
    If orderID = 0 Then Exit Function
    If Not PoIsOpen(PoStatus(orderID)) Then
        PurchaseOrderNoteProblem = "√„— «·‘—«¡ ·Â–« «·≈–‰ ·Ì” „› ÊÕ« ··«” ·«„ („” ·„ √Ê „·€Ï)."
        Exit Function
    End If
    Set o = CurrentDb.OpenRecordset("SELECT SupplierID, WarehouseID FROM PurchaseOrders WHERE PurchaseOrderID = " & orderID, _
                                    dbOpenSnapshot, dbSeeChanges)
    Set note = CurrentDb.OpenRecordset("SELECT SupplierID, WarehouseID FROM StockNotes WHERE StockNoteID = " & StockNoteID, _
                                       dbOpenSnapshot, dbSeeChanges)
    If Nz(note!SupplierID, 0) <> o!SupplierID Or note!WarehouseID <> o!WarehouseID Then
        PurchaseOrderNoteProblem = "«·≈–‰ ·Ì” ·„Ê—œ √„— «·‘—«¡ Ê„Œ“‰Â."
    End If
    note.Close
    o.Close
    If Len(PurchaseOrderNoteProblem) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT PurchaseOrderLineID, ProductID, Sum(Quantity) AS LineQty FROM StockNoteLines " & _
                                     "WHERE StockNoteID = " & StockNoteID & " AND PurchaseOrderLineID Is Not Null " & _
                                     "GROUP BY PurchaseOrderLineID, ProductID", dbOpenSnapshot)
    Do Until rs.EOF
        Set ln = CurrentDb.OpenRecordset("SELECT PurchaseOrderID, ProductID, RemainingQty FROM qryPurchaseOrderLineStatus " & _
                                         "WHERE PurchaseOrderLineID = " & rs!PurchaseOrderLineID, dbOpenSnapshot)
        If ln.EOF Then
            PurchaseOrderNoteProblem = "”ÿ— ›Ì «·≈–‰ ·« Ìÿ«»ﬁ ”ÿ—« ›Ì √„— «·‘—«¡."
        ElseIf ln!PurchaseOrderID <> orderID Or ln!ProductID <> rs!ProductID Then
            PurchaseOrderNoteProblem = "”ÿ— ›Ì «·≈–‰ ·« Ìÿ«»ﬁ ’‰› ”ÿ—Â ›Ì √„— «·‘—«¡."
        ElseIf rs!LineQty > ln!RemainingQty Then
            PurchaseOrderNoteProblem = "«·ﬂ„Ì… ›Ì «·≈–‰ √ﬂ»— „‰ «·»«ﬁÌ ›Ì √„— «·‘—«¡: " & _
                                       DbValue("SELECT ProductName FROM Products WHERE ProductID = " & rs!ProductID) & _
                                       ": «·»«ﬁÌ " & ln!RemainingQty
        End If
        ln.Close
        If Len(PurchaseOrderNoteProblem) > 0 Then Exit Do
        rs.MoveNext
    Loop
    rs.Close
End Function

Public Sub SyncPurchaseOrder(ByVal db As DAO.Database, ByVal PurchaseOrderID As Long)
    ' Inside the transaction that posted, unposted or invoiced a note of the order: its status follows its notes.
    ' A draft, cancelled or closed order keeps its status.
    Dim rs As DAO.Recordset, st As String, newStatus As String
    Dim received As Currency, invoiced As Currency, remaining As Currency
    st = Nz(DbValue("SELECT Status FROM PurchaseOrders WHERE PurchaseOrderID = " & PurchaseOrderID), "")
    If Not PoIsOpen(st) And st <> "RECEIVED" And st <> "INVOICED" Then Exit Sub
    Set rs = db.OpenRecordset("SELECT Sum(ReceivedQty) AS R, Sum(InvoicedQty) AS I, Sum(RemainingQty) AS L " & _
                              "FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderID = " & PurchaseOrderID, dbOpenSnapshot)
    received = Nz(rs!R, 0)
    invoiced = Nz(rs!I, 0)
    remaining = Nz(rs!L, 0)
    rs.Close
    newStatus = PurchaseOrderStatusFor(received, invoiced, remaining)
    If newStatus <> st Then
        db.Execute "UPDATE PurchaseOrders SET Status = '" & newStatus & "' WHERE PurchaseOrderID = " & PurchaseOrderID, _
                   dbFailOnError + dbSeeChanges
    End If
End Sub

Public Function PurchaseOrderStatusFor(ByVal Received As Currency, ByVal Invoiced As Currency, _
                                       ByVal Remaining As Currency) As String
    If Received = 0 Then
        PurchaseOrderStatusFor = "APPROVED"
    ElseIf Remaining > 0 Then
        PurchaseOrderStatusFor = "PARTIAL"
    ElseIf Invoiced >= Received Then
        PurchaseOrderStatusFor = "INVOICED"
    Else
        PurchaseOrderStatusFor = "RECEIVED"
    End If
End Function

'==============================================================================
' The purchase suggestion
'==============================================================================
Public Function CreateSuggestedOrders(ByVal WarehouseID As Variant, ByRef Created As Long, ByRef Skipped As Long) As String
    ' One draft order per active supplier with the suggested lines of his products (Products.SupplierID), at the
    ' last purchase price. A product already in a draft order, or without an active supplier, is skipped.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, ln As DAO.Recordset, inTrans As Boolean
    Dim msg As String, sup As Long, orderID As Long
    Created = 0
    Skipped = 0
    If Not HasPermission("PURCHASE_ORDERS") Then
        CreateSuggestedOrders = "·Ì”  ·œÌﬂ ’·«ÕÌ… √Ê«„— «·‘—«¡."
        Exit Function
    End If
    If Not WarehouseIsActive(WarehouseID) Then
        CreateSuggestedOrders = "«Œ — «·„Œ“‰ «·–Ì Ì” ·„."
        Exit Function
    End If
    On Error GoTo EH
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set rs = db.OpenRecordset("SELECT q.ProductID, q.SupplierID, q.SuggestQty, q.PurchasePrice FROM PurchaseSuggestQuery AS q " & _
                              "WHERE q.SuggestQty > 0 AND q.ProductID NOT IN (SELECT l.ProductID FROM PurchaseOrderLines AS l " & _
                              "INNER JOIN PurchaseOrders AS o ON l.PurchaseOrderID = o.PurchaseOrderID " & _
                              "WHERE o.Status = 'DRAFT') ORDER BY q.SupplierID, q.ProductID", dbOpenSnapshot, dbSeeChanges)
    Set ln = db.OpenRecordset("PurchaseOrderLines", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    Do Until rs.EOF
        If Not NotePartyIsActive("RECEIPT", rs!SupplierID) Then
            Skipped = Skipped + 1
        Else
            If rs!SupplierID <> sup Then
                msg = CreatePurchaseOrder(rs!SupplierID, WarehouseID, Null, "«ﬁ —«Õ «·‘—«¡", orderID)
                If Len(msg) > 0 Then Exit Do
                sup = rs!SupplierID
                Created = Created + 1
            End If
            ln.AddNew
            ln!PurchaseOrderID = orderID
            ln!ProductID = rs!ProductID
            ln!Quantity = rs!SuggestQty
            ln!UnitCost = Nz(rs!PurchasePrice, 0)
            ln.Update
        End If
        rs.MoveNext
    Loop
    ln.Close
    rs.Close
    If Len(msg) > 0 Then
        ws.Rollback
        Created = 0
        CreateSuggestedOrders = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    If Created = 0 And Skipped = 0 Then CreateSuggestedOrders = "·«  ÊÃœ √’‰«› Ê’·  ·Õœ ≈⁄«œ… «·ÿ·» (√Ê ÂÌ ›Ì „”Êœ«  √Ê«„—)."
    LogAction "PURCHASE_SUGGEST", "PurchaseOrders", CStr(Created), "Skipped=" & Skipped
    Exit Function
EH:
    CreateSuggestedOrders = " ⁄–— ≈‰‘«¡ √Ê«„— «·‘—«¡: " & Err.Description & " (" & Err.Number & ")"
    Created = 0
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' frmPurchaseOrder (OpenArgs = an order to show) and its lines frmPurchaseOrderLines
'==============================================================================
Public Function PurchaseOrderStatusName(ByVal Status As String) As String
    Select Case Status
        Case "DRAFT": PurchaseOrderStatusName = "„”Êœ…"
        Case "APPROVED": PurchaseOrderStatusName = "„⁄ „œ°  Õ  «·ÿ·»"
        Case "PARTIAL": PurchaseOrderStatusName = "„” ·„ Ã“∆Ì«"
        Case "RECEIVED": PurchaseOrderStatusName = "„” ·„ ﬂ·Â° ·„ Ìı›Ê — ﬂ·Â"
        Case "INVOICED": PurchaseOrderStatusName = "„” ·„ Ê„›Ê —"
        Case "CANCELLED": PurchaseOrderStatusName = "„·€Ï"
        Case "CLOSED": PurchaseOrderStatusName = "„€·ﬁ »⁄œ «” ·«„ Ã“¡ „‰Â"
    End Select
End Function

Public Sub PurchaseOrderLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboSupplier.RowSource = NotePartyRows("RECEIPT")              ' the active suppliers (modStockNotes)
    frm!cboWarehouse.RowSource = WarehouseRows()
    frm!cboWarehouse.Value = CurrentWarehouseID()
    PurchaseOrderListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(PurchaseOrderID) FROM PurchaseOrders")
    frm!cboOrder.Value = id
    PurchaseOrderShow frm
End Sub

Private Sub PurchaseOrderListRefresh(ByVal frm As Access.Form)
    frm!cboOrder.RowSource = Tr("SELECT o.PurchaseOrderID, o.OrderNumber, Format(o.OrderDate, 'yyyy/mm/dd'), s.SupplierName, " & _
        "w.WarehouseName, IIf(o.Status = 'DRAFT', '„”Êœ…', IIf(o.Status = 'APPROVED', '„⁄ „œ', " & _
        "IIf(o.Status = 'PARTIAL', '„” ·„ Ã“∆Ì«', IIf(o.Status = 'RECEIVED', '„” ·„', " & _
        "IIf(o.Status = 'INVOICED', '„›Ê —', IIf(o.Status = 'CANCELLED', '„·€Ï', '„€·ﬁ')))))) " & _
        "FROM (PurchaseOrders AS o INNER JOIN [@Suppliers] AS s ON o.SupplierID = s.SupplierID) " & _
        "INNER JOIN [@Warehouses] AS w ON o.WarehouseID = w.WarehouseID ORDER BY o.PurchaseOrderID DESC")
    frm!cboOrder.Requery
End Sub

Public Sub PurchaseOrderShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, rs As DAO.Recordset, info As String, tot As DAO.Recordset, notes As Long
    id = Nz(frm!cboOrder.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM PurchaseOrderLines WHERE PurchaseOrderID = " & id & _
                                     " ORDER BY PurchaseOrderLineID"
    st = PoStatus(id)
    notes = PoNoteCount(id)
    If Len(st) = 0 Then
        info = "·« ÌÊÃœ √„— ‘—«¡. «Œ — «·„Ê—œ Ê«·„Œ“‰ À„ ´√„— ÃœÌœª° √Ê ´√Ê«„— „‰ «·«ﬁ —«Õª."
    Else
        Set rs = CurrentDb.OpenRecordset(Tr("SELECT o.OrderNumber, o.WarehouseID, o.ExpectedDate, s.SupplierName " & _
                                            "FROM PurchaseOrders AS o INNER JOIN [@Suppliers] AS s ON o.SupplierID = s.SupplierID " & _
                                            "WHERE o.PurchaseOrderID = ") & id, dbOpenSnapshot)
        info = rs!OrderNumber & "    " & rs!SupplierName & "  ->  " & WarehouseName(rs!WarehouseID) & "    "
        If Not IsNull(rs!ExpectedDate) Then info = info & "„Ê⁄œ «·«” ·«„ " & GDate(rs!ExpectedDate) & "    "
        rs.Close
        Set tot = CurrentDb.OpenRecordset("SELECT Sum(Quantity) AS Q, Sum(ReceivedQty) AS R, Sum(Quantity * UnitCost) AS V " & _
                                          "FROM qryPurchaseOrderLineStatus WHERE PurchaseOrderID = " & id, dbOpenSnapshot)
        info = info & PurchaseOrderStatusName(st) & "    «·„” ·„ " & Nz(tot!R, 0) & " „‰ " & Nz(tot!Q, 0) & _
               "    «·ﬁÌ„… »œÊ‰ «·÷—Ì»…: " & Format$(Nz(tot!V, 0), "#,##0.00")
        tot.Close
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    SafeFocus frm!cboOrder                  ' a focused control cannot be disabled
    frm!btnApprove.Enabled = (st = "DRAFT")
    frm!btnReceive.Enabled = PoIsOpen(st)
    frm!btnNotes.Enabled = (notes > 0)
    frm!btnUnapprove.Enabled = (st = "APPROVED" And notes = 0)
    frm!btnCancel.Enabled = (st = "DRAFT" Or PoIsOpen(st))
    frm!btnDelete.Enabled = (st = "DRAFT")
End Sub

Public Sub PurchaseOrderNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreatePurchaseOrder(frm!cboSupplier.Value, frm!cboWarehouse.Value, frm!txtExpected.Value, _
                              Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    frm!txtExpected.Value = Null
    PurchaseOrderListRefresh frm
    frm!cboOrder.Value = id
    PurchaseOrderShow frm
End Sub

Public Sub PurchaseOrderSuggest(ByVal frm As Access.Form)
    ' Draft orders from the purchase suggestion, for the warehouse chosen on the screen.
    Dim msg As String, created As Long, skipped As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    If Not AskYesNo("≈‰‘«¡ „”Êœ«  √Ê«„— ‘—«¡ „‰ «ﬁ —«Õ «·‘—«¡ („”Êœ… ·ﬂ· „Ê—œ)ø") Then Exit Sub
    msg = CreateSuggestedOrders(frm!cboWarehouse.Value, created, skipped)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    PurchaseOrderListRefresh frm
    frm!cboOrder.Value = DbValue("SELECT Max(PurchaseOrderID) FROM PurchaseOrders")
    PurchaseOrderShow frm
    msg = "√ı‰‘∆  " & created & " „”Êœ… √„— ‘—«¡. —«Ã⁄ «·ﬂ„Ì«  À„ «⁄ „œ ﬂ· √„—."
    If skipped > 0 Then msg = msg & vbCrLf & skipped & " ’‰› »·« „Ê—œ ‰‘ÿ ·„ ÌœŒ·: Õœ¯œ „Ê—œÂ ›Ì ‘«‘… «·„‰ Ã« ."
    ShowInfo msg
End Sub

Public Sub PurchaseOrderAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String, noteID As Long
    id = Nz(frm!cboOrder.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "APPROVE"
            If Not AskYesNo("«⁄ „«œ √„— «·‘—«¡ø  ıÕ”» ﬂ„Ì« Â  Õ  «·ÿ·».") Then Exit Sub
            msg = ApprovePurchaseOrder(id)
        Case "RECEIVE"
            msg = CreateOrderReceiptNote(id, noteID)
        Case "UNAPPROVE"
            If Not AskYesNo("≈⁄«œ… √„— «·‘—«¡ ·„”Êœ…ø") Then Exit Sub
            msg = UnapprovePurchaseOrder(id)
        Case "CANCEL"
            If Not AskYesNo("≈·€«¡ √„— «·‘—«¡ø ·« Ì»ﬁÏ „« ·„ Ìı” ·„ „‰Â  Õ  «·ÿ·».") Then Exit Sub
            msg = CancelPurchaseOrder(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› „”Êœ… √„— «·‘—«¡ø") Then Exit Sub
            msg = DeletePurchaseOrder(id)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    PurchaseOrderListRefresh frm
    If Action = "DELETE" Then frm!cboOrder.Value = DbValue("SELECT Max(PurchaseOrderID) FROM PurchaseOrders")
    PurchaseOrderShow frm
    If Action = "RECEIVE" Then OpenScreen "frmStockNote", 0, noteID     ' check the quantities, then post it
End Sub

Public Sub PurchaseOrderOpenNote(ByVal frm As Access.Form)
    ' The latest receipt note of the order.
    Dim noteID As Variant
    noteID = DbValue("SELECT Max(StockNoteID) FROM StockNotes WHERE PurchaseOrderID = " & Nz(frm!cboOrder.Value, 0))
    If Not IsNull(noteID) Then OpenScreen "frmStockNote", 0, noteID
End Sub

Public Sub PurchaseOrderLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String
    id = Nz(frm.Parent!cboOrder.Value, 0)
    If PoStatus(id) <> "DRAFT" Then
        msg = " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ. √⁄œ √„— «·‘—«¡ ·„”Êœ… √Ê·«."
    Else
        msg = PurchaseOrderLineProblem(frm!ProductID.Value, frm!Quantity.Value, frm!UnitCost.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!PurchaseOrderID.Value = id
    If Nz(frm!UnitCost.Value, 0) = 0 Then frm!UnitCost.Value = NoteDefaultPrice("RECEIPT", CLng(frm!ProductID.Value))
    AuditFormBefore frm, "PurchaseOrderLines", "PurchaseOrderLineID"
End Sub

Public Sub PurchaseOrderLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    PurchaseOrderShow frm.Parent
End Sub

Public Sub PurchaseOrderRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!PurchaseOrderLineID.Value
    If IsNull(id) Then Exit Sub
    If PoStatus(Nz(frm.Parent!cboOrder.Value, 0)) <> "DRAFT" Then
        ShowWarning " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ. √⁄œ √„— «·‘—«¡ ·„”Êœ… √Ê·«."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("PurchaseOrderLines", "PurchaseOrderLineID", id)
    CurrentDb.Execute "DELETE FROM PurchaseOrderLines WHERE PurchaseOrderLineID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "PURCHASE_ORDER_LINE_DELETE", "PurchaseOrderLines", id, auditBefore
    frm.Requery
    PurchaseOrderShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Private Sub CheckPo(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestPurchaseOrders() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, pid2 As Long, msg As String, sup As Long
    Dim passed As Long, failed As Long, Report As String, orderID As Long, noteID As Long, invID As Long
    Dim order2 As Long, created As Long, skipped As Long, cat As Long, unit As Long, refNo As String
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestPurchaseOrders  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPurchaseLines", dbFailOnError + dbSeeChanges
    g_AutoAnswer = False
    sup = Nz(DbValue("SELECT Min(SupplierID) FROM Suppliers WHERE IsActive <> 0"), 0)
    CheckPo sup > 0, "ÌÊÃœ „Ê—œ ‰‘ÿ", passed, failed, Report
    cat = Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1)
    unit = Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1)
    ' a product at its reorder level of 10 (stock 4) and one with a reorder quantity of 50 (stock 0)
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost, MinimumQuantity, ReorderQuantity, SupplierID) VALUES ('TEST-PO-1', " & _
                      "'TEST ’‰› √Ê«„— «·‘—«¡', " & cat & ", " & unit & ", 20, 8, 0, 0, 10, 0, " & sup & ")", _
                      dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PO-1'")
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost, MinimumQuantity, ReorderQuantity, SupplierID) VALUES ('TEST-PO-2', " & _
                      "'TEST ’‰› À«‰Ú', " & cat & ", " & unit & ", 9, 3, 0, 0, 5, 50, " & sup & ")", _
                      dbFailOnError + dbSeeChanges
    pid2 = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-PO-2'")
    msg = PostManualStock(pid, TT_STOCK_IN, 4, 8, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckPo Len(msg) = 0, "—’Ìœ 4 ··’‰› «·√Ê· " & msg, passed, failed, Report

    ' the suggestion: 2 x 10 - 4 = 16, and the reorder quantity 50
    CheckPo SuggestQty(4, 0, 10, 0) = 16 And SuggestQty(11, 0, 10, 0) = 0 And SuggestQty(0, 0, 5, 50) = 50, _
            "«·ﬂ„Ì… «·„ﬁ —Õ…", passed, failed, Report
    CheckPo Nz(DbValue("SELECT SuggestQty FROM PurchaseSuggestQuery WHERE ProductID = " & pid), 0) = 16 And _
            Nz(DbValue("SELECT SuggestQty FROM PurchaseSuggestQuery WHERE ProductID = " & pid2), 0) = 50, _
            "«ﬁ —«Õ «·‘—«¡: 16 Ê50", passed, failed, Report
    msg = CreateSuggestedOrders(MAIN_WAREHOUSE_ID, created, skipped)
    orderID = Nz(DbValue("SELECT Max(l.PurchaseOrderID) FROM PurchaseOrderLines AS l WHERE l.ProductID = " & pid), 0)
    CheckPo Len(msg) = 0 And created >= 1 And PoStatus(orderID) = "DRAFT" And _
            Nz(DbValue("SELECT Sum(Quantity) FROM PurchaseOrderLines WHERE PurchaseOrderID = " & orderID & _
                       " AND ProductID In (" & pid & ", " & pid2 & ")"), 0) = 66, _
            "„”Êœ… √„— ··„Ê—œ »«·ﬂ„Ì«  «·„ﬁ —Õ… " & msg, passed, failed, Report
    CheckPo Nz(DbValue("SELECT UnitCost FROM PurchaseOrderLines WHERE PurchaseOrderID = " & orderID & " AND ProductID = " & _
                       pid), 0) = 8, "»¬Œ— ”⁄— ‘—«¡", passed, failed, Report
    msg = CreateSuggestedOrders(MAIN_WAREHOUSE_ID, created, skipped)
    CheckPo Nz(DbValue("SELECT COUNT(*) FROM PurchaseOrderLines WHERE ProductID = " & pid), 0) = 1, _
            "·« Ì ﬂ—— ’‰› ›Ì „”Êœ… √Œ—Ï", passed, failed, Report
    CurrentDb.Execute "DELETE FROM PurchaseOrderLines WHERE PurchaseOrderID = " & orderID & " AND ProductID <> " & pid & _
                      " AND ProductID <> " & pid2, dbFailOnError + dbSeeChanges

    ' approving puts it on order: the suggestion no longer asks for it
    msg = ApprovePurchaseOrder(orderID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "APPROVED" And OnOrderQty(pid) = 16, "«·«⁄ „«œ: 16  Õ  «·ÿ·» " & msg, _
            passed, failed, Report
    CheckPo IsNull(DbValue("SELECT ProductID FROM PurchaseSuggestQuery WHERE ProductID = " & pid)), _
            "„«  Õ  «·ÿ·» ÌıÕ”» ›Ì «·«ﬁ —«Õ", passed, failed, Report

    ' a partial receipt of 10, then the rest
    msg = CreateOrderReceiptNote(orderID, noteID)
    CheckPo Len(msg) = 0 And Nz(DbValue("SELECT Sum(Quantity) FROM StockNoteLines WHERE StockNoteID = " & noteID), 0) = 66, _
            "≈–‰ «·«” ·«„ »«·»«ﬁÌ „‰ «·√„— " & msg, passed, failed, Report
    CheckPo Len(CancelPurchaseOrder(orderID)) > 0, "·« Ìı·€Ï «·√„— Ê·Â ≈–‰ ·„ Ìı—ÕÛ¯·", passed, failed, Report
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 10 WHERE StockNoteID = " & noteID & " AND ProductID = " & pid, _
                      dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM StockNoteLines WHERE StockNoteID = " & noteID & " AND ProductID = " & pid2, _
                      dbFailOnError + dbSeeChanges
    msg = PostStockNote(noteID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "PARTIAL" And OnOrderQty(pid) = 6 And OnOrderQty(pid2) = 50 And _
            WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 14, "«” ·«„ 10: „” ·„ Ã“∆Ì«° Ê«·»«ﬁÌ 6  Õ  «·ÿ·» " & msg, _
            passed, failed, Report
    msg = InvoiceStockNote(noteID, "CREDIT", "TEST-PO-INV1", True, invID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "PARTIAL" And invID > 0, "›« Ê—… «·≈–‰ «·√Ê· " & msg, passed, failed, Report
    CheckPo Len(UnapprovePurchaseOrder(orderID)) > 0, "·« Ì⁄Êœ ·„”Êœ… »⁄œ «·«” ·«„", passed, failed, Report

    msg = CreateOrderReceiptNote(orderID, noteID)
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 7 WHERE StockNoteID = " & noteID & " AND ProductID = " & pid, _
                      dbFailOnError + dbSeeChanges
    CheckPo Len(PostStockNote(noteID)) > 0, "·« Ìı” ·„ √ﬂÀ— „‰ «·»«ﬁÌ ›Ì «·√„— " & msg, passed, failed, Report
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 6 WHERE StockNoteID = " & noteID & " AND ProductID = " & pid, _
                      dbFailOnError + dbSeeChanges
    msg = PostStockNote(noteID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "RECEIVED" And OnOrderQty(pid) = 0, "«” ·«„ «·»«ﬁÌ: „” ·„ ﬂ·Â " & msg, _
            passed, failed, Report
    msg = UnpostStockNote(noteID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "PARTIAL" And OnOrderQty(pid) = 6, _
            "≈·€«¡  —ÕÌ· «·≈–‰ Ì⁄ÌœÂ  Õ  «·ÿ·» " & msg, passed, failed, Report
    msg = PostStockNote(noteID) & InvoiceStockNote(noteID, "CREDIT", "TEST-PO-INV2", True, invID)
    CheckPo Len(msg) = 0 And PoStatus(orderID) = "INVOICED", "›« Ê—… «·≈–‰ «·À«‰Ì: „” ·„ Ê„›Ê — " & msg, _
            passed, failed, Report
    CheckPo Len(CreateOrderReceiptNote(orderID, noteID)) > 0, "·« ≈–‰ „‰ √„— «” ı·„ ﬂ·Â", passed, failed, Report

    ' cancelling: nothing received = CANCELLED, part received = CLOSED
    msg = CreatePurchaseOrder(sup, MAIN_WAREHOUSE_ID, Date + 7, "TEST", order2)
    CurrentDb.Execute "INSERT INTO PurchaseOrderLines (PurchaseOrderID, ProductID, Quantity, UnitCost) VALUES (" & _
                      order2 & ", " & pid & ", 5, 8)", dbFailOnError + dbSeeChanges
    msg = msg & ApprovePurchaseOrder(order2) & CreateOrderReceiptNote(order2, noteID)
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 2 WHERE StockNoteID = " & noteID, dbFailOnError + dbSeeChanges
    msg = msg & PostStockNote(noteID)
    CheckPo Len(msg) = 0 And OnOrderQty(pid) = 3, "√„— À«‰Ú «” ı·„ „‰Â 2 " & msg, passed, failed, Report
    msg = CancelPurchaseOrder(order2)
    CheckPo Len(msg) = 0 And PoStatus(order2) = "CLOSED" And OnOrderQty(pid) = 0, _
            "«·≈·€«¡ »⁄œ «” ·«„ Ã“¡: „€·ﬁ Ê·« ‘Ì¡  Õ  «·ÿ·» " & msg, passed, failed, Report
    CheckPo Nz(DbValue("SELECT COUNT(*) FROM PendingPurchaseOrdersQuery WHERE PurchaseOrderID = " & order2), 0) = 0, _
            "«·„€·ﬁ ·« ÌŸÂ— ›Ì «·√Ê«„— «·„› ÊÕ…", passed, failed, Report
    msg = CreatePurchaseOrder(sup, MAIN_WAREHOUSE_ID, Null, "TEST", order2)
    CheckPo Len(ApprovePurchaseOrder(order2)) > 0, "·« Ìı⁄ „œ √„— »·« √”ÿ—", passed, failed, Report
    CheckPo Len(DeletePurchaseOrder(order2)) = 0 And PoStatus(order2) = "", "Õ–› „”Êœ… «·√„—", passed, failed, Report
    CheckPo Len(CreatePurchaseOrder(Null, MAIN_WAREHOUSE_ID, Null, "TEST", order2)) > 0, "«·√„— ÌÕ «Ã „Ê—œ«", _
            passed, failed, Report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckPo False, "Œÿ√: " & Err.Description, passed, failed, Report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  √Ê«„— «·‘—«¡ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestPurchaseOrders"
        TestPurchaseOrders = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, _
                "TestPurchaseOrders"
    End If
End Function
