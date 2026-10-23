Attribute VB_Name = "modSalesOrders"
'==============================================================================
' modSalesOrders  -  Retail Store Management System (customer orders, docs/62, plan 4A)
'
' A customer order is entered as a draft, approved (which reserves its quantities in its warehouse, so they
' are not sold to anybody else), prepared, then delivered with delivery notes (modStockNotes, docs/61) made
' from it, in one go or in parts, and invoiced from those notes.
'
'   CreateSalesOrder          a draft for a customer and a warehouse
'   ApproveSalesOrder         DRAFT -> APPROVED: the free stock must cover it; from now on it is reserved
'   PrepareSalesOrder         APPROVED -> PREPARING (the warehouse is picking it)
'   CreateOrderDeliveryNote   a draft delivery note with what is left of each line (SalesOrderLineID)
'   SyncSalesOrder            after a note of the order is posted, unposted or invoiced: APPROVED / PREPARING
'                             (nothing delivered), PARTIAL, DELIVERED (all delivered), INVOICED (all invoiced)
'   UnapproveSalesOrder       back to a draft while it has no delivery note
'   CancelSalesOrder          CANCELLED, or CLOSED when part of it was delivered: the rest is no longer reserved
'   DeleteSalesOrder          a draft only
' Reserved = ordered less delivered (posted notes) on the lines of APPROVED / PREPARING / PARTIAL orders
' (qrySalesOrderLineStatus); CheckStockAvailable (modSales) takes it off the stock of the warehouse.
' Screen: frmSalesOrder (built by modBuildForms).
'==============================================================================
Option Compare Database
Option Explicit

Private Const OPEN_STATES As String = "'APPROVED', 'PREPARING', 'PARTIAL'"

Private Function OrderStatus(ByVal SalesOrderID As Long) As String
    OrderStatus = Nz(DbValue("SELECT Status FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID), "")
End Function

Private Function OrderIsOpen(ByVal Status As String) As Boolean
    OrderIsOpen = (Status = "APPROVED" Or Status = "PREPARING" Or Status = "PARTIAL")
End Function

Private Function OrderNoteCount(ByVal SalesOrderID As Long, Optional ByVal OnlyDrafts As Boolean = False) As Long
    Dim sql As String
    sql = "SELECT COUNT(*) FROM StockNotes WHERE SalesOrderID = " & SalesOrderID
    If OnlyDrafts Then sql = sql & " AND Status = 'DRAFT'"
    OrderNoteCount = Nz(DbValue(sql), 0)
End Function

Private Function OrderDelivered(ByVal SalesOrderID As Long) As Currency
    OrderDelivered = Nz(DbValue("SELECT Sum(DeliveredQty) FROM qrySalesOrderLineStatus WHERE SalesOrderID = " & _
                                SalesOrderID), 0)
End Function

Public Function ReservedQty(ByVal ProductID As Long, ByVal WarehouseID As Long, _
                            Optional ByVal ExceptOrderID As Long = 0) As Currency
    ' What the open orders still need of a product in a warehouse (0 = the user's), except one order.
    If Nz(DbValue("SELECT COUNT(*) FROM SalesOrders WHERE Status In (" & OPEN_STATES & ")"), 0) = 0 Then Exit Function
    If WarehouseID = 0 Then WarehouseID = CurrentWarehouseID()
    ReservedQty = Nz(DbValue("SELECT Sum(RemainingQty) FROM qrySalesOrderLineStatus WHERE ProductID = " & ProductID & _
                             " AND WarehouseID = " & WarehouseID & " AND OrderStatus In (" & OPEN_STATES & _
                             ") AND SalesOrderID <> " & ExceptOrderID), 0)
End Function

Public Function OrderLineDelivered(ByVal SalesOrderLineID As Variant) As Variant
    ' The delivered column of the order lines grid.
    OrderLineDelivered = Null
    If IsNull(SalesOrderLineID) Then Exit Function
    OrderLineDelivered = Nz(DbValue("SELECT DeliveredQty FROM qrySalesOrderLineStatus WHERE SalesOrderLineID = " & _
                                    CLng(SalesOrderLineID)), 0)
End Function

Public Function SalesOrderLineProblem(ByVal ProductID As Variant, ByVal Qty As Variant, ByVal UnitPrice As Variant) As String
    If IsNull(ProductID) Then
        SalesOrderLineProblem = "«Œ — «·’‰›."
    ElseIf Not ProductTracksStock(CLng(ProductID)) Then
        SalesOrderLineProblem = "Â–« «·’‰› ·« Ì «»⁄ «·„Œ“Ê‰° ›·« ÌıÕÃ“ Ê·« Ìı”·Û¯„ »≈–‰."
    ElseIf Nz(Qty, 0) <= 0 Then
        SalesOrderLineProblem = "«·ﬂ„Ì… √ﬂ»— „‰ ’›—."
    ElseIf Nz(UnitPrice, 0) < 0 Then
        SalesOrderLineProblem = "«·”⁄— ·« ÌﬂÊ‰ ”«·»«."
    End If
End Function

Public Function CreateSalesOrder(ByVal CustomerID As Variant, ByVal WarehouseID As Variant, ByVal RequiredDate As Variant, _
                                 ByVal Notes As String, ByRef NewOrderID As Long) As String
    Dim rs As DAO.Recordset
    NewOrderID = 0
    If Not HasPermission("SALES_ORDERS") Then
        CreateSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
    ElseIf Not NotePartyIsActive("DELIVERY", CustomerID) Then
        CreateSalesOrder = "«Œ — «·⁄„Ì·."
    ElseIf Not WarehouseIsActive(WarehouseID) Then
        CreateSalesOrder = "«Œ — «·„Œ“‰."
    End If
    If Len(CreateSalesOrder) = 0 And Not IsNull(RequiredDate) Then
        If Not IsDate(RequiredDate) Then CreateSalesOrder = "„Ê⁄œ «· ”·Ì„ ·Ì”  «—ÌŒ«."
    End If
    If Len(CreateSalesOrder) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SalesOrders", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
    rs.AddNew
    rs!OrderNumber = NextNumber("SALES_ORDER")
    rs!OrderDate = Now
    rs!CustomerID = CLng(CustomerID)
    rs!WarehouseID = CLng(WarehouseID)
    rs!SalesRepID = SalesRepFor(CLng(CustomerID))                    ' modSalesReps
    If Not IsNull(RequiredDate) Then rs!RequiredDate = DateValue(CDate(RequiredDate))
    rs!Status = "DRAFT"
    rs!EmployeeID = CurrentUserID()
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewOrderID = rs!SalesOrderID
    rs.Close
    LogAction "SALES_ORDER_NEW", "SalesOrders", CStr(NewOrderID)
End Function

Public Function ApproveSalesOrder(ByVal SalesOrderID As Long) As String
    ' The free stock of the warehouse (its stock less the other open orders) must cover the order.
    Dim rs As DAO.Recordset, msg As String, n As Long, ids() As Long, qtys() As Currency, wh As Long
    If Not HasPermission("SALES_ORDERS") Then
        ApproveSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
        Exit Function
    End If
    If OrderStatus(SalesOrderID) <> "DRAFT" Then
        ApproveSalesOrder = " ı⁄ „œ «·ÿ·»Ì… ÊÂÌ „”Êœ…."
        Exit Function
    End If
    If Not NotePartyIsActive("DELIVERY", DbValue("SELECT CustomerID FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID)) Then
        ApproveSalesOrder = "«·⁄„Ì· €Ì— „ÊÃÊœ √Ê €Ì— ‰‘ÿ."
        Exit Function
    End If
    wh = DbValue("SELECT WarehouseID FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID)
    msg = WarehouseProblem(wh)
    If Len(msg) = 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT ProductID, Quantity, UnitPrice FROM SalesOrderLines WHERE SalesOrderID = " & _
                                         SalesOrderID, dbOpenSnapshot)
        Do Until rs.EOF
            ReDim Preserve ids(0 To n)
            ReDim Preserve qtys(0 To n)
            ids(n) = rs!ProductID
            qtys(n) = rs!Quantity
            If Len(msg) = 0 Then msg = SalesOrderLineProblem(rs!ProductID, rs!Quantity, rs!UnitPrice)
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
        If Len(msg) = 0 And n = 0 Then msg = "√÷› «·√’‰«› Êﬂ„Ì« Â« √Ê·«."
    End If
    If Len(msg) = 0 Then msg = CheckStockAvailable(ids, qtys, n, "«⁄ „«œ «·ÿ·»Ì… ÊÕÃ“ ﬂ„Ì« Â«", wh, SalesOrderID)
    If Len(msg) > 0 Then
        ApproveSalesOrder = msg
        Exit Function
    End If
    CurrentDb.Execute "UPDATE SalesOrders SET Status = 'APPROVED', ApprovedAt = Now() WHERE SalesOrderID = " & SalesOrderID, _
                      dbFailOnError + dbSeeChanges
    LogAction "SALES_ORDER_APPROVE", "SalesOrders", CStr(SalesOrderID)
End Function

Public Function PrepareSalesOrder(ByVal SalesOrderID As Long) As String
    If Not HasPermission("SALES_ORDERS") Then
        PrepareSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
    ElseIf OrderStatus(SalesOrderID) <> "APPROVED" Then
        PrepareSalesOrder = "Ì»œ√  ÃÂÌ“ «·ÿ·»Ì… «·„⁄ „œ…."
    Else
        CurrentDb.Execute "UPDATE SalesOrders SET Status = 'PREPARING', PreparedAt = Now() WHERE SalesOrderID = " & _
                          SalesOrderID, dbFailOnError + dbSeeChanges
        LogAction "SALES_ORDER_PREPARE", "SalesOrders", CStr(SalesOrderID)
    End If
End Function

Public Function UnapproveSalesOrder(ByVal SalesOrderID As Long) As String
    ' Back to a draft (its lines can be changed again) while no delivery note was made from it.
    Dim before As Collection, st As String
    st = OrderStatus(SalesOrderID)
    If Not HasPermission("SALES_ORDERS") Then
        UnapproveSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
    ElseIf st <> "APPROVED" And st <> "PREPARING" Then
        UnapproveSalesOrder = " ⁄Êœ ··„”Êœ… «·ÿ·»Ì… «·„⁄ „œ… √Ê ﬁÌœ «· ÃÂÌ“."
    ElseIf OrderNoteCount(SalesOrderID) > 0 Then
        UnapproveSalesOrder = "’œ—  „‰ «·ÿ·»Ì… √–Ê‰«   ”·Ì„. «Õ–› «·„”Êœ«  „‰Â«° √Ê √·€ˆ «·ÿ·»Ì…."
    End If
    If Len(UnapproveSalesOrder) > 0 Then Exit Function
    Set before = AuditSnapshot("SalesOrders", "SalesOrderID", SalesOrderID)
    CurrentDb.Execute "UPDATE SalesOrders SET Status = 'DRAFT', ApprovedAt = Null, PreparedAt = Null " & _
                      "WHERE SalesOrderID = " & SalesOrderID, dbFailOnError + dbSeeChanges
    AuditEdited "SALES_ORDER_UNAPPROVE", "SalesOrders", "SalesOrderID", SalesOrderID, before
End Function

Public Function CancelSalesOrder(ByVal SalesOrderID As Long) As String
    ' What is left is no longer reserved. An order delivered in part is CLOSED: its notes stay as they are.
    Dim before As Collection, st As String, newStatus As String
    st = OrderStatus(SalesOrderID)
    If Not HasPermission("SALES_ORDERS") Then
        CancelSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
    ElseIf st <> "DRAFT" And Not OrderIsOpen(st) Then
        CancelSalesOrder = "«·ÿ·»Ì… „ı”·Û¯„… ﬂ·Â« √Ê „·€«… „‰ ﬁ»·."
    ElseIf OrderNoteCount(SalesOrderID, True) > 0 Then
        CancelSalesOrder = "··ÿ·»Ì… ≈–‰  ”·Ì„ ·„ Ìı—ÕÛ¯·. —Õ¯·Â √Ê «Õ–›Â √Ê·«."
    End If
    If Len(CancelSalesOrder) > 0 Then Exit Function
    newStatus = "CANCELLED"
    If OrderDelivered(SalesOrderID) > 0 Then newStatus = "CLOSED"
    Set before = AuditSnapshot("SalesOrders", "SalesOrderID", SalesOrderID)
    CurrentDb.Execute "UPDATE SalesOrders SET Status = '" & newStatus & "' WHERE SalesOrderID = " & SalesOrderID, dbFailOnError + dbSeeChanges
    AuditEdited "SALES_ORDER_CANCEL", "SalesOrders", "SalesOrderID", SalesOrderID, before
End Function

Public Function DeleteSalesOrder(ByVal SalesOrderID As Long) As String
    ' A draft only; its lines go with it.
    Dim auditBefore As Collection
    If Not HasPermission("SALES_ORDERS") Then
        DeleteSalesOrder = "·Ì”  ·œÌﬂ ’·«ÕÌ… ÿ·»Ì«  «·⁄„·«¡."
        Exit Function
    End If
    If OrderStatus(SalesOrderID) <> "DRAFT" Or OrderNoteCount(SalesOrderID) > 0 Then
        DeleteSalesOrder = " ıÕ–› «·„”Êœ… ›ﬁÿ. √·€ˆ «·ÿ·»Ì… «·„⁄ „œ… »œ· Õ–›Â«."
        Exit Function
    End If
    Set auditBefore = AuditSnapshot("SalesOrders", "SalesOrderID", SalesOrderID)
    QuotationOrderDeleted SalesOrderID                      ' its quotation is issued again (modQuotations)
    CurrentDb.Execute "DELETE FROM SalesOrderLines WHERE SalesOrderID = " & SalesOrderID, dbFailOnError + dbSeeChanges
    CurrentDb.Execute "DELETE FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID, dbFailOnError + dbSeeChanges
    AuditDeleted "SALES_ORDER_DELETE", "SalesOrders", SalesOrderID, auditBefore
End Function

'==============================================================================
' Delivery: the notes of an order
'==============================================================================
Public Function CreateOrderDeliveryNote(ByVal SalesOrderID As Long, ByRef NewNoteID As Long) As String
    ' A draft delivery note from the warehouse of the order with what is left of each line, at the order price.
    ' The user may lower the quantities (a partial delivery) before posting it.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, nl As DAO.Recordset, inTrans As Boolean
    Dim msg As String, o As DAO.Recordset, orderNo As String, cust As Long, wh As Long
    NewNoteID = 0
    If Not HasPermission("STOCK_NOTES") Then
        CreateOrderDeliveryNote = "·Ì”  ·œÌﬂ ’·«ÕÌ… √–Ê‰«  «· ”·Ì„ Ê«·«” ·«„."
        Exit Function
    End If
    If Not OrderIsOpen(OrderStatus(SalesOrderID)) Then
        CreateOrderDeliveryNote = "Ìı‰‘√ ≈–‰ «· ”·Ì„ „‰ ÿ·»Ì… „⁄ „œ… ·„  ı”·Û¯„ ﬂ·Â«."
        Exit Function
    End If
    If OrderNoteCount(SalesOrderID, True) > 0 Then
        CreateOrderDeliveryNote = "··ÿ·»Ì… ≈–‰  ”·Ì„ ·„ Ìı—ÕÛ¯·. «› ÕÂ „‰ ´√–Ê‰«  «·ÿ·»Ì…ª."
        Exit Function
    End If
    If Nz(DbValue("SELECT Sum(RemainingQty) FROM qrySalesOrderLineStatus WHERE SalesOrderID = " & SalesOrderID), 0) <= 0 Then
        CreateOrderDeliveryNote = "”ı·ˆ¯„  ﬂ· √’‰«› «·ÿ·»Ì…."
        Exit Function
    End If

    On Error GoTo EH
    Set db = CurrentDb
    Set o = db.OpenRecordset("SELECT OrderNumber, CustomerID, WarehouseID FROM SalesOrders WHERE SalesOrderID = " & _
                             SalesOrderID, dbOpenSnapshot)
    orderNo = o!OrderNumber
    cust = o!CustomerID
    wh = o!WarehouseID
    o.Close                                  ' closed before BeginTrans (error 3246 inside it)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = CreateStockNote("DELIVERY", wh, cust, "«·ÿ·»Ì… " & orderNo, NewNoteID)   ' modStockNotes
    If Len(msg) = 0 Then
        db.Execute "UPDATE StockNotes SET SalesOrderID = " & SalesOrderID & " WHERE StockNoteID = " & NewNoteID, dbFailOnError + dbSeeChanges
        Set rs = db.OpenRecordset("SELECT SalesOrderLineID, ProductID, RemainingQty, UnitPrice FROM qrySalesOrderLineStatus " & _
                                  "WHERE SalesOrderID = " & SalesOrderID & " AND RemainingQty > 0 ORDER BY SalesOrderLineID", _
                                  dbOpenSnapshot, dbSeeChanges)
        Set nl = db.OpenRecordset("StockNoteLines", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
        Do Until rs.EOF
            nl.AddNew
            nl!StockNoteID = NewNoteID
            nl!ProductID = rs!ProductID
            nl!Quantity = rs!RemainingQty
            nl!UnitPrice = rs!UnitPrice
            nl!UnitCost = 0
            nl!SalesOrderLineID = rs!SalesOrderLineID
            nl.Update
            rs.MoveNext
        Loop
        nl.Close
        rs.Close
    End If
    If Len(msg) > 0 Then
        ws.Rollback
        NewNoteID = 0
        CreateOrderDeliveryNote = msg
        Exit Function
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "SALES_ORDER_NOTE", "SalesOrders", CStr(SalesOrderID), "Note=" & NewNoteID
    Exit Function
EH:
    CreateOrderDeliveryNote = " ⁄–— ≈‰‘«¡ ≈–‰ «· ”·Ì„: " & Err.Description & " (" & Err.Number & ")"
    NewNoteID = 0
    If inTrans Then ws.Rollback
End Function

Public Function OrderNoteProblem(ByVal StockNoteID As Long) As String
    ' Before a delivery note of an order is posted: the order is open, the note is for its customer and warehouse,
    ' each line delivers the product of its order line and not more than what is left of it.
    Dim rs As DAO.Recordset, ln As DAO.Recordset, orderID As Long, o As DAO.Recordset, note As DAO.Recordset
    orderID = NoteOrderID(StockNoteID)                              ' modStockNotes
    If orderID = 0 Then Exit Function
    If Not OrderIsOpen(OrderStatus(orderID)) Then
        OrderNoteProblem = "ÿ·»Ì… Â–« «·≈–‰ ·Ì”  „› ÊÕ… ·· ”·Ì„ („ı”·Û¯„… √Ê „·€«…)."
        Exit Function
    End If
    Set o = CurrentDb.OpenRecordset("SELECT CustomerID, WarehouseID FROM SalesOrders WHERE SalesOrderID = " & orderID, _
                                    dbOpenSnapshot, dbSeeChanges)
    Set note = CurrentDb.OpenRecordset("SELECT CustomerID, WarehouseID FROM StockNotes WHERE StockNoteID = " & StockNoteID, _
                                       dbOpenSnapshot, dbSeeChanges)
    If Nz(note!CustomerID, 0) <> o!CustomerID Or note!WarehouseID <> o!WarehouseID Then
        OrderNoteProblem = "«·≈–‰ ·Ì” ·⁄„Ì· «·ÿ·»Ì… Ê„Œ“‰Â«."
    End If
    note.Close
    o.Close
    If Len(OrderNoteProblem) > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT SalesOrderLineID, ProductID, Sum(Quantity) AS LineQty FROM StockNoteLines " & _
                                     "WHERE StockNoteID = " & StockNoteID & " AND SalesOrderLineID Is Not Null " & _
                                     "GROUP BY SalesOrderLineID, ProductID", dbOpenSnapshot)
    Do Until rs.EOF
        Set ln = CurrentDb.OpenRecordset("SELECT SalesOrderID, ProductID, RemainingQty FROM qrySalesOrderLineStatus " & _
                                         "WHERE SalesOrderLineID = " & rs!SalesOrderLineID, dbOpenSnapshot)
        If ln.EOF Then
            OrderNoteProblem = "”ÿ— ›Ì «·≈–‰ ·« Ìÿ«»ﬁ ”ÿ—« ›Ì «·ÿ·»Ì…."
        ElseIf ln!SalesOrderID <> orderID Or ln!ProductID <> rs!ProductID Then
            OrderNoteProblem = "”ÿ— ›Ì «·≈–‰ ·« Ìÿ«»ﬁ ’‰› ”ÿ—Â ›Ì «·ÿ·»Ì…."
        ElseIf rs!LineQty > ln!RemainingQty Then
            OrderNoteProblem = "«·ﬂ„Ì… ›Ì «·≈–‰ √ﬂ»— „‰ «·»«ﬁÌ ›Ì «·ÿ·»Ì…: " & _
                               DbValue("SELECT ProductName FROM Products WHERE ProductID = " & rs!ProductID) & _
                               ": «·»«ﬁÌ " & ln!RemainingQty
        End If
        ln.Close
        If Len(OrderNoteProblem) > 0 Then Exit Do
        rs.MoveNext
    Loop
    rs.Close
End Function

Public Sub SyncSalesOrder(ByVal db As DAO.Database, ByVal SalesOrderID As Long)
    ' Inside the transaction that posted, unposted or invoiced a note of the order: its status follows its notes.
    ' A draft, cancelled or closed order keeps its status.
    Dim rs As DAO.Recordset, st As String, prepared As Boolean, newStatus As String
    Dim delivered As Currency, invoiced As Currency, remaining As Currency
    Set rs = db.OpenRecordset("SELECT Status, PreparedAt FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    st = rs!Status
    prepared = Not IsNull(rs!PreparedAt)
    rs.Close
    If Not OrderIsOpen(st) And st <> "DELIVERED" And st <> "INVOICED" Then Exit Sub
    Set rs = db.OpenRecordset("SELECT Sum(DeliveredQty) AS D, Sum(InvoicedQty) AS I, Sum(RemainingQty) AS R " & _
                              "FROM qrySalesOrderLineStatus WHERE SalesOrderID = " & SalesOrderID, dbOpenSnapshot)
    delivered = Nz(rs!D, 0)
    invoiced = Nz(rs!I, 0)
    remaining = Nz(rs!R, 0)
    rs.Close
    If delivered = 0 Then
        newStatus = "APPROVED"
        If prepared Then newStatus = "PREPARING"
    ElseIf remaining > 0 Then
        newStatus = "PARTIAL"
    ElseIf invoiced >= delivered Then
        newStatus = "INVOICED"
    Else
        newStatus = "DELIVERED"
    End If
    If newStatus <> st Then
        db.Execute "UPDATE SalesOrders SET Status = '" & newStatus & "' WHERE SalesOrderID = " & SalesOrderID, dbFailOnError + dbSeeChanges
    End If
End Sub

'==============================================================================
' frmSalesOrder (OpenArgs = an order to show) and its lines frmSalesOrderLines
'==============================================================================
Public Function OrderStatusName(ByVal Status As String) As String
    Select Case Status
        Case "DRAFT": OrderStatusName = "„”Êœ…"
        Case "APPROVED": OrderStatusName = "„⁄ „œ… Ê„ÕÃÊ“…"
        Case "PREPARING": OrderStatusName = "ﬁÌœ «· ÃÂÌ“"
        Case "PARTIAL": OrderStatusName = "„ı”·Û¯„… Ã“∆Ì«"
        Case "DELIVERED": OrderStatusName = "„ı”·Û¯„… ﬂ·Â«° ·„  ı›Ê — ﬂ·Â«"
        Case "INVOICED": OrderStatusName = "„ı”·Û¯„… Ê„›Ê —…"
        Case "CANCELLED": OrderStatusName = "„·€«…"
        Case "CLOSED": OrderStatusName = "„€·ﬁ… »⁄œ  ”·Ì„ Ã“¡ „‰Â«"
    End Select
End Function

Public Sub SalesOrderLoad(ByVal frm As Access.Form)
    Dim id As Variant
    Calendar = vbCalGreg
    frm!cboCustomer.RowSource = NotePartyRows("DELIVERY")             ' the active customers (modStockNotes)
    frm!cboWarehouse.RowSource = WarehouseRows()
    frm!cboWarehouse.Value = CurrentWarehouseID()
    SalesOrderListRefresh frm
    id = Null
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then id = CLng(frm.OpenArgs)
    End If
    If IsNull(id) Then id = DbValue("SELECT Max(SalesOrderID) FROM SalesOrders")
    frm!cboOrder.Value = id
    SalesOrderShow frm
End Sub

Private Sub SalesOrderListRefresh(ByVal frm As Access.Form)
    frm!cboOrder.RowSource = Tr("SELECT o.SalesOrderID, o.OrderNumber, Format(o.OrderDate, 'yyyy/mm/dd'), c.CustomerName, " & _
        "w.WarehouseName, IIf(o.Status = 'DRAFT', '„”Êœ…', IIf(o.Status = 'APPROVED', '„⁄ „œ…', " & _
        "IIf(o.Status = 'PREPARING', 'ﬁÌœ «· ÃÂÌ“', IIf(o.Status = 'PARTIAL', '„ı”·Û¯„… Ã“∆Ì«', " & _
        "IIf(o.Status = 'DELIVERED', '„ı”·Û¯„…', IIf(o.Status = 'INVOICED', '„›Ê —…', " & _
        "IIf(o.Status = 'CANCELLED', '„·€«…', '„€·ﬁ…'))))))) " & _
        "FROM (SalesOrders AS o INNER JOIN [@Customers] AS c ON o.CustomerID = c.CustomerID) " & _
        "INNER JOIN [@Warehouses] AS w ON o.WarehouseID = w.WarehouseID ORDER BY o.SalesOrderID DESC")
    frm!cboOrder.Requery
End Sub

Public Sub SalesOrderShow(ByVal frm As Access.Form)
    Dim id As Long, st As String, rs As DAO.Recordset, info As String, tot As DAO.Recordset, notes As Long
    id = Nz(frm!cboOrder.Value, 0)
    frm!subLines.Form.RecordSource = "SELECT * FROM SalesOrderLines WHERE SalesOrderID = " & id & " ORDER BY SalesOrderLineID"
    st = OrderStatus(id)
    notes = OrderNoteCount(id)
    If Len(st) = 0 Then
        info = "·«  ÊÃœ ÿ·»Ì…. «Œ — «·⁄„Ì· Ê«·„Œ“‰ À„ ´ÿ·»Ì… ÃœÌœ…ª."
    Else
        Set rs = CurrentDb.OpenRecordset(Tr("SELECT o.OrderNumber, o.WarehouseID, o.RequiredDate, c.CustomerName " & _
                                            "FROM SalesOrders AS o INNER JOIN [@Customers] AS c ON o.CustomerID = c.CustomerID " & _
                                            "WHERE o.SalesOrderID = ") & id, dbOpenSnapshot)
        info = rs!OrderNumber & "    " & rs!CustomerName & "  <-  " & WarehouseName(rs!WarehouseID) & "    "
        If Not IsNull(rs!RequiredDate) Then info = info & "„Ê⁄œ «· ”·Ì„ " & GDate(rs!RequiredDate) & "    "
        rs.Close
        Set tot = CurrentDb.OpenRecordset("SELECT Sum(Quantity) AS Q, Sum(DeliveredQty) AS D, Sum(Quantity * UnitPrice) AS V " & _
                                          "FROM qrySalesOrderLineStatus WHERE SalesOrderID = " & id, dbOpenSnapshot)
        info = info & OrderStatusName(st) & "    «·„ı”·Û¯„ " & Nz(tot!D, 0) & " „‰ " & Nz(tot!Q, 0) & _
               "    «·ﬁÌ„…: " & Format$(Nz(tot!V, 0), "#,##0.00")
        tot.Close
    End If
    frm!lblInfo.Caption = Tr(info)
    frm!subLines.Form.AllowEdits = (st = "DRAFT")
    frm!subLines.Form.AllowAdditions = (st = "DRAFT")
    SafeFocus frm!cboOrder                  ' a focused control cannot be disabled
    frm!btnApprove.Enabled = (st = "DRAFT")
    frm!btnPrepare.Enabled = (st = "APPROVED")
    frm!btnDeliver.Enabled = OrderIsOpen(st)
    frm!btnNotes.Enabled = (notes > 0)
    frm!btnUnapprove.Enabled = ((st = "APPROVED" Or st = "PREPARING") And notes = 0)
    frm!btnCancel.Enabled = (st = "DRAFT" Or OrderIsOpen(st))
    frm!btnDelete.Enabled = (st = "DRAFT")
End Sub

Public Sub SalesOrderNew(ByVal frm As Access.Form)
    Dim msg As String, id As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Sub
    msg = CreateSalesOrder(frm!cboCustomer.Value, frm!cboWarehouse.Value, frm!txtRequired.Value, _
                           Nz(frm!txtNotes.Value, ""), id)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!txtNotes.Value = Null
    frm!txtRequired.Value = Null
    SalesOrderListRefresh frm
    frm!cboOrder.Value = id
    SalesOrderShow frm
End Sub

Public Sub SalesOrderAction(ByVal frm As Access.Form, ByVal Action As String)
    Dim id As Long, msg As String, noteID As Long
    id = Nz(frm!cboOrder.Value, 0)
    If id = 0 Then Exit Sub
    If Not CanScreenAction(frm.Name, IIf(Action = "DELETE", "DELETE", "EDIT")) Then Exit Sub
    If frm!subLines.Form.Dirty Then frm!subLines.Form.Dirty = False
    Select Case Action
        Case "APPROVE"
            If Not AskYesNo("«⁄ „«œ «·ÿ·»Ì… ÊÕÃ“ ﬂ„Ì« Â« ›Ì «·„Œ“‰ø") Then Exit Sub
            msg = ApproveSalesOrder(id)
        Case "PREPARE"
            msg = PrepareSalesOrder(id)
        Case "DELIVER"
            msg = CreateOrderDeliveryNote(id, noteID)
        Case "UNAPPROVE"
            If Not AskYesNo("≈⁄«œ… «·ÿ·»Ì… ·„”Êœ…ø Ìı·€Ï ÕÃ“ ﬂ„Ì« Â«.") Then Exit Sub
            msg = UnapproveSalesOrder(id)
        Case "CANCEL"
            If Not AskYesNo("≈·€«¡ «·ÿ·»Ì…ø Ìı·€Ï ÕÃ“ „« ·„ Ìı”·Û¯„ „‰Â«.") Then Exit Sub
            msg = CancelSalesOrder(id)
        Case "DELETE"
            If Not AskYesNo("Õ–› „”Êœ… «·ÿ·»Ì…ø") Then Exit Sub
            msg = DeleteSalesOrder(id)
    End Select
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    SalesOrderListRefresh frm
    If Action = "DELETE" Then frm!cboOrder.Value = DbValue("SELECT Max(SalesOrderID) FROM SalesOrders")
    SalesOrderShow frm
    If Action = "DELIVER" Then OpenScreen "frmStockNote", 0, noteID     ' check the quantities, then post it
End Sub

Public Sub SalesOrderOpenNote(ByVal frm As Access.Form)
    ' The latest delivery note of the order.
    Dim noteID As Variant
    noteID = DbValue("SELECT Max(StockNoteID) FROM StockNotes WHERE SalesOrderID = " & Nz(frm!cboOrder.Value, 0))
    If Not IsNull(noteID) Then OpenScreen "frmStockNote", 0, noteID
End Sub

Public Sub StockNoteOpenOrder(ByVal frm As Access.Form)
    ' frmStockNote: the customer order or the purchase order (modPurchaseOrders) of the note.
    Dim orderID As Long
    orderID = NoteOrderID(Nz(frm!cboNote.Value, 0))
    If orderID > 0 Then OpenScreen "frmSalesOrder", 0, orderID
    orderID = NotePurchaseOrderID(Nz(frm!cboNote.Value, 0))
    If orderID > 0 Then OpenScreen "frmPurchaseOrder", 0, orderID
End Sub

Public Function OrderDefaultPrice(ByVal SalesOrderID As Long, ByVal ProductID As Long) As Currency
    ' The price list of the order's customer (modPriceLists), else the price list of the order's rep
    ' (modRepRoutes) or the selling price, less the discount of the customer's list.
    Dim v As Variant, usual As Currency
    v = RepProductPrice(Nz(DbValue("SELECT SalesRepID FROM SalesOrders WHERE SalesOrderID = " & SalesOrderID), 0), ProductID)
    If IsNull(v) Then
        usual = NoteDefaultPrice("DELIVERY", ProductID)
    Else
        usual = CCur(v)
    End If
    OrderDefaultPrice = CustomerPrice(Nz(DbValue("SELECT CustomerID FROM SalesOrders WHERE SalesOrderID = " & _
                                                 SalesOrderID), 0), ProductID, Null, usual)
End Function

Public Sub SalesOrderLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim id As Long, msg As String
    id = Nz(frm.Parent!cboOrder.Value, 0)
    If OrderStatus(id) <> "DRAFT" Then
        msg = " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ. √⁄œ «·ÿ·»Ì… ·„”Êœ… √Ê·«."
    Else
        msg = SalesOrderLineProblem(frm!ProductID.Value, frm!Quantity.Value, frm!UnitPrice.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    If frm.NewRecord Then frm!SalesOrderID.Value = id
    If Nz(frm!UnitPrice.Value, 0) = 0 Then frm!UnitPrice.Value = OrderDefaultPrice(id, CLng(frm!ProductID.Value))
    AuditFormBefore frm, "SalesOrderLines", "SalesOrderLineID"
End Sub

Public Sub SalesOrderLineAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    SalesOrderShow frm.Parent
End Sub

Public Sub SalesOrderRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, auditBefore As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!SalesOrderLineID.Value
    If IsNull(id) Then Exit Sub
    If OrderStatus(Nz(frm.Parent!cboOrder.Value, 0)) <> "DRAFT" Then
        ShowWarning " ı⁄œÛ¯· √”ÿ— «·„”Êœ… ›ﬁÿ. √⁄œ «·ÿ·»Ì… ·„”Êœ… √Ê·«."
        Exit Sub
    End If
    If frm.Dirty Then frm.Undo
    Set auditBefore = AuditSnapshot("SalesOrderLines", "SalesOrderLineID", id)
    CurrentDb.Execute "DELETE FROM SalesOrderLines WHERE SalesOrderLineID = " & CLng(id), dbFailOnError + dbSeeChanges
    AuditDeleted "SALES_ORDER_LINE_DELETE", "SalesOrderLines", id, auditBefore
    frm.Requery
    SalesOrderShow frm.Parent
End Sub

'==============================================================================
' Test (inside a transaction that is rolled back)
'==============================================================================
Public Function TestSalesOrders() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, refNo As String, msg As String, cust As Long
    Dim passed As Long, failed As Long, report As String, orderID As Long, noteID As Long, invID As Long
    Dim ids(0 To 0) As Long, qtys(0 To 0) As Currency, order2 As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestSalesOrders  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    EnsureLocalTables
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError + dbSeeChanges
    CurrentDb.Execute "UPDATE Settings SET AllowNegativeStock = False", dbFailOnError + dbSeeChanges
    g_AutoAnswer = False
    CurrentDb.Execute "INSERT INTO Products (ProductCode, ProductName, CategoryID, UnitID, SellingPrice, PurchasePrice, " & _
                      "CurrentQuantity, AverageCost) VALUES ('TEST-SO-P', 'TEST ’‰› «·ÿ·»Ì« ', " & _
                      Nz(DbValue("SELECT Min(CategoryID) FROM Categories"), 1) & ", " & _
                      Nz(DbValue("SELECT Min(UnitID) FROM Units"), 1) & ", 20, 8, 0, 0)", dbFailOnError + dbSeeChanges
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = 'TEST-SO-P'")
    cust = Nz(DbValue("SELECT Min(CustomerID) FROM Customers WHERE IsActive <> 0"), 0)
    CheckSo cust > 0, "ÌÊÃœ ⁄„Ì· ‰‘ÿ", passed, failed, report
    msg = PostManualStock(pid, TT_STOCK_IN, 10, 6, "TEST", refNo, MAIN_WAREHOUSE_ID)
    CheckSo Len(msg) = 0, "≈÷«›… 10 » ﬂ·›… 6 " & msg, passed, failed, report

    ' an order of 6: approving it reserves 6, so only 4 are left for others
    msg = CreateSalesOrder(cust, MAIN_WAREHOUSE_ID, Null, "TEST", orderID)
    CurrentDb.Execute "INSERT INTO SalesOrderLines (SalesOrderID, ProductID, Quantity, UnitPrice) VALUES (" & _
                      orderID & ", " & pid & ", 6, 20)", dbFailOnError + dbSeeChanges
    msg = msg & ApproveSalesOrder(orderID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "APPROVED" And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 6, _
            "«⁄ „«œ «·ÿ·»Ì… ÌÕÃ“ 6 " & msg, passed, failed, report
    ids(0) = pid
    qtys(0) = 5
    CheckSo Len(CheckStockAvailable(ids, qtys, 1, "TEST", MAIN_WAREHOUSE_ID)) > 0, "·« Ìı»«⁄ ·€Ì—Â«: 5 √ﬂÀ— „‰ «·„ «Õ 4", _
            passed, failed, report
    qtys(0) = 4
    CheckSo Len(CheckStockAvailable(ids, qtys, 1, "TEST", MAIN_WAREHOUSE_ID)) = 0, "Ìı»«⁄ «·„ «Õ 4", passed, failed, report
    msg = CreateSalesOrder(cust, MAIN_WAREHOUSE_ID, Null, "TEST", order2)
    CurrentDb.Execute "INSERT INTO SalesOrderLines (SalesOrderID, ProductID, Quantity, UnitPrice) VALUES (" & _
                      order2 & ", " & pid & ", 5, 20)", dbFailOnError + dbSeeChanges
    CheckSo Len(ApproveSalesOrder(order2)) > 0 And OrderStatus(order2) = "DRAFT", "·«  ı⁄ „œ ÿ·»Ì… √ﬂ»— „‰ «·„ «Õ", _
            passed, failed, report
    CheckSo Len(DeleteSalesOrder(order2)) = 0 And OrderStatus(order2) = "", "Õ–› „”Êœ… «·ÿ·»Ì…", passed, failed, report

    ' a partial delivery of 4, then the rest
    CheckSo Len(PrepareSalesOrder(orderID)) = 0 And OrderStatus(orderID) = "PREPARING", "»œ¡ «· ÃÂÌ“", passed, failed, report
    msg = CreateOrderDeliveryNote(orderID, noteID)
    CheckSo Len(msg) = 0 And Nz(DbValue("SELECT Sum(Quantity) FROM StockNoteLines WHERE StockNoteID = " & noteID), 0) = 6, _
            "≈–‰ «· ”·Ì„ »«·»«ﬁÌ „‰ «·ÿ·»Ì… " & msg, passed, failed, report
    CheckSo Len(CancelSalesOrder(orderID)) > 0, "·«  ı·€Ï «·ÿ·»Ì… Ê·Â« ≈–‰ ·„ Ìı—ÕÛ¯·", passed, failed, report
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 4 WHERE StockNoteID = " & noteID, dbFailOnError + dbSeeChanges
    msg = PostStockNote(noteID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "PARTIAL" And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 2 And _
            WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 6, " ”·Ì„ 4: „ı”·Û¯„… Ã“∆Ì« Ê«·„ÕÃÊ“ 2 " & msg, passed, failed, report
    msg = InvoiceStockNote(noteID, "CASH", "", True, invID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "PARTIAL", "›« Ê—… «·≈–‰ «·√Ê· " & msg, passed, failed, report
    CheckSo Len(UnapproveSalesOrder(orderID)) > 0, "·«  ⁄Êœ ·„”Êœ… »⁄œ «· ”·Ì„", passed, failed, report

    msg = CreateOrderDeliveryNote(orderID, noteID)
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 3 WHERE StockNoteID = " & noteID, dbFailOnError + dbSeeChanges
    CheckSo Len(PostStockNote(noteID)) > 0, "·« Ìı”·Û¯„ √ﬂÀ— „‰ «·»«ﬁÌ ›Ì «·ÿ·»Ì… " & msg, passed, failed, report
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 2 WHERE StockNoteID = " & noteID, dbFailOnError + dbSeeChanges
    msg = PostStockNote(noteID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "DELIVERED" And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 0, _
            " ”·Ì„ «·»«ﬁÌ: „ı”·Û¯„… ﬂ·Â« Ê·« ÕÃ“ " & msg, passed, failed, report
    msg = UnpostStockNote(noteID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "PARTIAL" And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 2, _
            "≈·€«¡  —ÕÌ· «·≈–‰ Ì⁄Ìœ «·ÕÃ“ " & msg, passed, failed, report
    msg = PostStockNote(noteID) & InvoiceStockNote(noteID, "CASH", "", True, invID)
    CheckSo Len(msg) = 0 And OrderStatus(orderID) = "INVOICED" And WarehouseStock(pid, MAIN_WAREHOUSE_ID) = 4, _
            "›« Ê—… «·≈–‰ «·À«‰Ì: „ı”·Û¯„… Ê„›Ê —… " & msg, passed, failed, report
    CheckSo Len(CreateOrderDeliveryNote(orderID, noteID)) > 0, "·« ≈–‰ „‰ ÿ·»Ì… ”ı·ˆ¯„  ﬂ·Â«", passed, failed, report

    ' cancelling frees the reservation; an order delivered in part is closed
    msg = CreateSalesOrder(cust, MAIN_WAREHOUSE_ID, Date + 3, "TEST", order2)
    CurrentDb.Execute "INSERT INTO SalesOrderLines (SalesOrderID, ProductID, Quantity, UnitPrice) VALUES (" & _
                      order2 & ", " & pid & ", 3, 20)", dbFailOnError + dbSeeChanges
    msg = msg & ApproveSalesOrder(order2) & CreateOrderDeliveryNote(order2, noteID)
    CurrentDb.Execute "UPDATE StockNoteLines SET Quantity = 1 WHERE StockNoteID = " & noteID, dbFailOnError + dbSeeChanges
    msg = msg & PostStockNote(noteID)
    CheckSo Len(msg) = 0 And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 2, "ÿ·»Ì… À«‰Ì… ”ı·ˆ¯„ „‰Â« 1 " & msg, passed, failed, report
    msg = CancelSalesOrder(order2)
    CheckSo Len(msg) = 0 And OrderStatus(order2) = "CLOSED" And ReservedQty(pid, MAIN_WAREHOUSE_ID) = 0, _
            "«·≈·€«¡ »⁄œ  ”·Ì„ Ã“¡: „€·ﬁ… Ê·« ÕÃ“ " & msg, passed, failed, report
    CheckSo Nz(DbValue("SELECT COUNT(*) FROM PendingOrdersQuery WHERE SalesOrderID = " & order2), 0) = 0, _
            "«·„€·ﬁ… ·«  ŸÂ— ›Ì «·ÿ·»Ì«  «·„› ÊÕ…", passed, failed, report
    CheckSo Len(CreateSalesOrder(Null, MAIN_WAREHOUSE_ID, Null, "TEST", order2)) > 0, "«·ÿ·»Ì…  Õ «Ã ⁄„Ì·«", _
            passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckSo False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ÿ·»Ì«  «·⁄„·«¡ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestSalesOrders"
        TestSalesOrders = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestSalesOrders"
    End If
End Function

Private Sub CheckSo(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
