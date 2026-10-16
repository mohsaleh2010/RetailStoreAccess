Attribute VB_Name = "modCash"
'==============================================================================
' modCash  -  Retail Store Management System (Treasury)
'
' Cash boxes: the main safe (BoxType MAIN) and the cashier boxes (CASHIER).
' A balance is never stored: it is the sum of qryCashMovements, i.e. the cash part
' of sales, sales returns, customer receipts, purchases, purchase returns,
' supplier payments and expenses, plus the cash vouchers of this module.
'
'   CurrentCashBoxID   box of the logged-in user: Employees.CashBoxID, else the main safe
'                      for users with the CASH_BOX permission and the first cashier box for the others
'   CashBoxFor         box a document pays into / out of (Null = not paid in cash)
'   CashBoxBalance     balance of a box, optionally before a moment
'   PostCashVoucher    cash in / cash out / transfer between two boxes; a cash out of
'                      the EXPENSE category also records the expense
'   PostCashClosing    cashier closing: count, shortage / overage, then the cash goes
'                      to the main safe, to the owner, or stays in the box
'   CanUseBox          who pays, transfers and closes from a box (docs/56): its custodians
'                      (CashBoxUsers, Employees.CashBoxID, the sales rep of a custody box),
'                      everyone with CASH_BOX when it has none, every box with CASH_ALL_BOXES
'   ReceiveCashTransfer / CancelCashTransfer   a transfer to a box with ConfirmReceipt stays
'                      in transit (account 1195) until a custodian of that box receives it
' Screens: frmTreasury, frmCashVoucher, frmCashClosing, frmCashTransfers (built by modBuildForms).
' Posting functions return "" on success or an Arabic error message.
'==============================================================================
Option Compare Database
Option Explicit

Public Const CASH_METHOD_ID As Long = 1          ' PaymentMethods: ‰ﬁœÌ

Private Const BOX_ROWS As String = "SELECT q.CashBoxID, q.BoxName FROM [@CashBoxes] AS q WHERE q.IsActive = True"
Private Const VOUCHER_TYPES As String = "IN;”‰œ ﬁ»÷ ‰ﬁœÌ…;OUT;”‰œ ’—› ‰ﬁœÌ…;TRANSFER; ÕÊÌ· »Ì‰ «·’‰«œÌﬁ"
Private Const IN_CATEGORIES As String = "OTHER;ﬁ»÷ ‰ﬁœÌ… (≈Ì—«œ«  √Œ—Ï);OWNER;≈Ìœ«⁄ „‰ «·„«·ﬂ;" & _
                                        "ADVANCE;”œ«œ ”·›… „ÊŸ›"
Private Const OUT_CATEGORIES As String = "EXPENSE;„’—Ê› (Ìı”Ã· ›Ì «·„’—Ê›« );OWNER; ”ÊÌ… / „”ÕÊ»«  «·„«·ﬂ;" & _
                                         "ADVANCE;”·›… „ÊŸ›;COMMISSION;’—› ⁄„Ê·… „‰œÊ»;OTHER;’—› ¬Œ—"
Private Const EMPLOYEE_ROWS As String = "SELECT EmployeeID, EmployeeName FROM Employees WHERE IsActive = True ORDER BY EmployeeName"
Private Const REP_ROWS As String = "SELECT s.SalesRepID, s.RepName FROM [@SalesReps] AS s WHERE s.IsActive = True ORDER BY s.RepName"
Private Const DESTINATIONS As String = "MAIN; —ÕÌ· ≈·Ï «·Œ“Ì‰… «·—∆Ì”Ì…;OWNER; ”·Ì„ ··„«·ﬂ ( ”ÊÌ…);" & _
                                       "KEEP;Ì»ﬁÏ ›Ì «·’‰œÊﬁ"

'==============================================================================
' Boxes and balances
'==============================================================================
Public Function CurrentCashBoxID() As Long
    Dim v As Variant
    If ActingRepDayID() > 0 Then                    ' the office works in the name of a rep (modVanSales)
        CurrentCashBoxID = ActingCashBoxID()
        Exit Function
    End If
    v = DbValue("SELECT e.CashBoxID FROM Employees AS e INNER JOIN CashBoxes AS b ON e.CashBoxID = b.CashBoxID " & _
                "WHERE e.EmployeeID = " & CurrentUserID() & " AND b.IsActive = True")
    ' no box of his own: the main safe for who manages the treasury, the cashier box for the others
    If IsNull(v) And HasPermission("CASH_BOX") Then _
        v = DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True")
    If IsNull(v) Then v = DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'CASHIER' AND IsActive = True")
    If IsNull(v) Then v = DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE IsActive = True")
    CurrentCashBoxID = Nz(v, 0)
End Function

Public Function CashBoxFor(ByVal PaymentMethodID As Variant, ByVal Amount As Currency) As Variant
    ' Only cash moves a box: card / transfer payments and unpaid (credit) amounts do not.
    Dim box As Long
    CashBoxFor = Null
    If Nz(PaymentMethodID, 0) <> CASH_METHOD_ID Or Amount = 0 Then Exit Function
    box = CurrentCashBoxID()
    If box > 0 Then CashBoxFor = box
End Function

Public Function CashBoxBalance(ByVal BoxID As Long, Optional ByVal Before As Variant, _
                               Optional ByVal Inclusive As Boolean = False) As Currency
    ' Balance now, or of the movements before (Inclusive: up to and including) a moment.
    Dim sql As String
    If IsMissing(Before) Then Before = Null
    sql = "SELECT Sum(AmountIn) - Sum(AmountOut) FROM qryCashMovements WHERE CashBoxID = " & BoxID
    If Not IsNull(Before) Then sql = sql & " AND MoveDate " & IIf(Inclusive, "<= ", "< ") & SqlDate(CDate(Before))
    CashBoxBalance = Nz(DbValue(sql), 0)
End Function

Public Function CashMovesTotal(ByVal BoxID As Long, ByVal Direction As String, ByVal After As Variant, _
                               Optional ByVal UpTo As Variant) As Currency
    ' Direction "IN" / "OUT": cash in or out of a box after a moment (Null = from the start)
    ' and before UpTo (missing = up to now).
    Dim sql As String
    If IsMissing(UpTo) Then UpTo = Null
    sql = "SELECT Sum(" & IIf(Direction = "IN", "AmountIn", "AmountOut") & ") FROM qryCashMovements " & _
          "WHERE CashBoxID = " & BoxID
    If Not IsNull(After) Then sql = sql & " AND MoveDate > " & SqlDate(CDate(After))
    If Not IsNull(UpTo) Then sql = sql & " AND MoveDate < " & SqlDate(CDate(UpTo))
    CashMovesTotal = Nz(DbValue(sql), 0)
End Function

Public Function LastClosingDate(ByVal BoxID As Long) As Variant
    LastClosingDate = DbValue("SELECT Max(ClosingDate) FROM CashClosings WHERE CashBoxID = " & BoxID)
End Function

Private Function BoxName(ByVal BoxID As Variant) As String
    If IsNull(BoxID) Then Exit Function
    BoxName = Nz(DbValue(Tr("SELECT b.BoxName FROM [@CashBoxes] AS b WHERE b.CashBoxID = " & CLng(BoxID))), "")
End Function

Public Function BoxIsActive(ByVal BoxID As Variant) As Boolean
    If IsNull(BoxID) Then Exit Function
    BoxIsActive = Nz(DbValue("SELECT COUNT(*) FROM CashBoxes WHERE IsActive = True AND CashBoxID = " & CLng(BoxID)), 0) > 0
End Function

Private Function Money(ByVal Value As Currency) As String
    Money = Format$(Value, "#,##0.00")
End Function

'==============================================================================
' Box custodians (docs/56)
'==============================================================================
Public Function IsBoxCustodian(ByVal BoxID As Variant) As Boolean
    ' The box of the user (Employees.CashBoxID), a box he keeps (CashBoxUsers) or the custody box of his sales rep.
    Dim uid As Long, box As Long
    If IsNull(BoxID) Then Exit Function
    uid = CurrentUserID()
    box = CLng(BoxID)
    IsBoxCustodian = Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE EmployeeID = " & uid & _
                                " AND CashBoxID = " & box), 0) > 0
    If Not IsBoxCustodian Then
        IsBoxCustodian = Nz(DbValue("SELECT COUNT(*) FROM CashBoxUsers WHERE CashBoxID = " & box & _
                                    " AND EmployeeID = " & uid), 0) > 0
    End If
    If Not IsBoxCustodian Then
        IsBoxCustodian = Nz(DbValue("SELECT COUNT(*) FROM CashBoxes AS b INNER JOIN SalesReps AS s ON " & _
                                    "b.SalesRepID = s.SalesRepID WHERE b.CashBoxID = " & box & _
                                    " AND s.EmployeeID = " & uid), 0) > 0
    End If
End Function

Public Function BoxHasCustodians(ByVal BoxID As Long) As Boolean
    BoxHasCustodians = Nz(DbValue("SELECT COUNT(*) FROM CashBoxUsers WHERE CashBoxID = " & BoxID), 0) > 0
    If Not BoxHasCustodians Then
        BoxHasCustodians = Nz(DbValue("SELECT COUNT(*) FROM Employees WHERE IsActive = True AND CashBoxID = " & _
                                      BoxID), 0) > 0
    End If
    If Not BoxHasCustodians Then
        BoxHasCustodians = Nz(DbValue("SELECT COUNT(*) FROM CashBoxes AS b INNER JOIN SalesReps AS s ON " & _
                                      "b.SalesRepID = s.SalesRepID WHERE b.CashBoxID = " & BoxID & _
                                      " AND s.EmployeeID Is Not Null"), 0) > 0
    End If
End Function

Public Function CanUseBox(ByVal BoxID As Variant) As Boolean
    ' Paying, transferring or closing from a box: every box with CASH_ALL_BOXES; otherwise the boxes
    ' the user keeps and the boxes nobody keeps.
    If IsNull(BoxID) Then Exit Function
    If HasPermission("CASH_ALL_BOXES") Then
        CanUseBox = True
        Exit Function
    End If
    If ActingRepDayID() > 0 Then                    ' the custody box of the rep the office works for
        If CLng(BoxID) = ActingCashBoxID() Then
            CanUseBox = True
            Exit Function
        End If
    End If
    CanUseBox = IsBoxCustodian(BoxID)
    If Not CanUseBox Then CanUseBox = Not BoxHasCustodians(CLng(BoxID))
End Function

Public Function BoxUseProblem(ByVal BoxID As Variant) As String
    If Not CanUseBox(BoxID) Then BoxUseProblem = "·”  „‰ √„‰«¡ " & BoxName(BoxID) & ": ·« Ì„ﬂ‰ﬂ «·’—› √Ê «· ÕÊÌ· „‰Â."
End Function

Public Function ClosingAllowed(ByVal BoxID As Variant) As Boolean
    ' A cashier closes his own box; who manages the treasury closes the boxes he may use.
    If IsNull(BoxID) Then Exit Function
    If CLng(BoxID) = CurrentCashBoxID() Or IsBoxCustodian(BoxID) Then
        ClosingAllowed = True
    ElseIf HasPermission("CASH_BOX") Then
        ClosingAllowed = CanUseBox(BoxID)
    End If
End Function

Private Function ClosingBoxes() As String
    ' The active boxes the user may close, as a list for IN (...).
    Dim rs As DAO.Recordset, ids As String
    Set rs = CurrentDb.OpenRecordset("SELECT CashBoxID FROM CashBoxes WHERE IsActive = True", dbOpenSnapshot)
    Do While Not rs.EOF
        If ClosingAllowed(rs!CashBoxID) Then ids = ids & "," & rs!CashBoxID
        rs.MoveNext
    Loop
    rs.Close
    ClosingBoxes = IIf(Len(ids) = 0, "0", Mid$(ids, 2))
End Function

Public Function BoxConfirmsReceipt(ByVal BoxID As Variant) As Boolean
    If IsNull(BoxID) Then Exit Function
    BoxConfirmsReceipt = Nz(DbValue("SELECT ConfirmReceipt FROM CashBoxes WHERE CashBoxID = " & CLng(BoxID)), False)
End Function

'==============================================================================
' Transfers in transit (docs/56)
'==============================================================================
Private Function PendingTransferBox(ByVal CashVoucherID As Long, ByVal FieldName As String) As Variant
    ' The sending (CashBoxID) or receiving (ToCashBoxID) box of a transfer still in transit; Null otherwise.
    PendingTransferBox = DbValue("SELECT " & FieldName & " FROM CashVouchers WHERE CashVoucherID = " & CashVoucherID & _
                                 " AND VoucherType = 'TRANSFER' AND InTransit = True AND ReceivedAt Is Null")
End Function

Public Function ReceiveCashTransfer(ByVal CashVoucherID As Long) As String
    ' A custodian of the receiving box confirms the cash arrived: it enters the box now.
    Dim toBox As Variant, before As Collection, msg As String
    toBox = PendingTransferBox(CashVoucherID, "ToCashBoxID")
    If IsNull(toBox) Then
        msg = "Â–« «· ÕÊÌ· ·Ì” »«‰ Ÿ«— «·«” ·«„."
    ElseIf Not HasPermission("CASH_CLOSING") And Not HasPermission("CASH_BOX") Then
        msg = "·Ì”  ·œÌﬂ ’·«ÕÌ… «·Œ“Ì‰…."
    ElseIf Not CanUseBox(toBox) Then
        msg = "Ì” ·„ «· ÕÊÌ· √„Ì‰ " & BoxName(toBox) & "."
    Else
        msg = ClosedPeriodProblem(Date)
    End If
    If Len(msg) > 0 Then
        ReceiveCashTransfer = msg
        Exit Function
    End If
    Set before = AuditSnapshot("CashVouchers", "CashVoucherID", CashVoucherID)
    CurrentDb.Execute "UPDATE CashVouchers SET ReceivedAt = " & SqlDate(Now) & ", ReceivedByID = " & CurrentUserID() & _
                      " WHERE CashVoucherID = " & CashVoucherID & " AND ReceivedAt Is Null", dbFailOnError
    AuditEdited "CASH_TRANSFER_RECEIVE", "CashVouchers", "CashVoucherID", CashVoucherID, before
End Function

Public Function CancelCashTransfer(ByVal CashVoucherID As Long) As String
    ' Before its receipt the sender can take a transfer back: the voucher is deleted and the cash
    ' returns to the sending box. A closing transfer stays (the closing counts on it).
    Dim fromBox As Variant, before As Collection, msg As String
    fromBox = PendingTransferBox(CashVoucherID, "CashBoxID")
    If IsNull(fromBox) Then
        msg = "Â–« «· ÕÊÌ· ·Ì” »«‰ Ÿ«— «·«” ·«„."
    ElseIf Not HasPermission("CASH_BOX") Then
        msg = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·Œ“Ì‰…."
    ElseIf Not CanUseBox(fromBox) Then
        msg = BoxUseProblem(fromBox)
    ElseIf Not IsNull(DbValue("SELECT ClosingID FROM CashVouchers WHERE CashVoucherID = " & CashVoucherID)) Then
        msg = " ÕÊÌ·  ’›Ì… «·ﬂ«‘Ì— ·« Ìı·€Ï: Ì” ·„Â √„Ì‰ «·’‰œÊﬁ «·„” ·„."
    Else
        msg = ClosedPeriodProblem(DbValue("SELECT VoucherDate FROM CashVouchers WHERE CashVoucherID = " & CashVoucherID))
    End If
    If Len(msg) > 0 Then
        CancelCashTransfer = msg
        Exit Function
    End If
    Set before = AuditSnapshot("CashVouchers", "CashVoucherID", CashVoucherID)
    CurrentDb.Execute "DELETE FROM CashVouchers WHERE CashVoucherID = " & CashVoucherID, dbFailOnError
    AuditDeleted "CASH_TRANSFER_DELETE", "CashVouchers", CashVoucherID, before
End Function

'==============================================================================
' Posting
'==============================================================================
Public Function PostCashVoucher(ByVal VoucherType As String, ByVal BoxID As Long, ByVal ToBoxID As Variant, _
                                ByVal Category As String, ByVal Amount As Currency, ByVal PartyName As String, _
                                ByVal Description As String, ByVal ExpenseTypeID As Variant, _
                                ByRef NewVoucherID As Long) As String
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean, msg As String, balance As Currency

    On Error GoTo EH
    NewVoucherID = 0
    If Not HasPermission("CASH_BOX") Then
        PostCashVoucher = "·Ì”  ·œÌﬂ ’·«ÕÌ… ”‰œ«  «·Œ“Ì‰…."
        Exit Function
    End If
    msg = VoucherProblem(VoucherType, BoxID, ToBoxID, Category, Amount, ExpenseTypeID)
    If Len(msg) = 0 And VoucherType <> "IN" Then msg = BoxUseProblem(BoxID)
    If Len(msg) = 0 And VoucherType <> "IN" Then
        balance = CashBoxBalance(BoxID)
        If Amount > balance Then msg = "—’Ìœ " & BoxName(BoxID) & " (" & Money(balance) & ") ·« Ìﬂ›Ì ·’—› " & Money(Amount) & "."
    End If
    If Len(msg) > 0 Then
        PostCashVoucher = msg
        Exit Function
    End If

    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    NewVoucherID = InsertVoucher(db, VoucherType, BoxID, ToBoxID, Category, Amount, PartyName, Description, _
                                 ExpenseTypeID, 0, Now)
    ws.CommitTrans
    inTrans = False
    LogAction "CASH_" & VoucherType, "CashVouchers", CStr(NewVoucherID), Category & " Amount=" & Amount
    Exit Function

EH:
    PostCashVoucher = " ⁄–— Õ›Ÿ «·”‰œ: " & Err.Description & " (" & Err.Number & ")"
    NewVoucherID = 0
    If inTrans Then ws.Rollback
End Function

Public Function VoucherProblem(ByVal VoucherType As String, ByVal BoxID As Long, ByVal ToBoxID As Variant, _
                               ByVal Category As String, ByVal Amount As Currency, _
                               ByVal ExpenseTypeID As Variant) As String
    ' The checks of a voucher entered by the user ("" = fine). Balances are checked by the caller.
    If Amount <= 0 Then
        VoucherProblem = "√œŒ· «·„»·€ (√ﬂ»— „‰ ’›—)."
    ElseIf Not BoxIsActive(BoxID) Then
        VoucherProblem = "«Œ — «·’‰œÊﬁ."
    ElseIf VoucherType = "TRANSFER" Then
        If Not BoxIsActive(ToBoxID) Then
            VoucherProblem = "«Œ — «·’‰œÊﬁ «·„ÕÊÛ¯· ≈·ÌÂ."
        ElseIf CLng(ToBoxID) = BoxID Then
            VoucherProblem = "·« Ì„ﬂ‰ «· ÕÊÌ· ≈·Ï ‰›” «·’‰œÊﬁ."
        End If
    ElseIf VoucherType = "IN" Then
        If InStr(";" & IN_CATEGORIES & ";", ";" & Category & ";") = 0 Then VoucherProblem = "«Œ — »‰œ «·ﬁ»÷."
    ElseIf VoucherType = "OUT" Then
        If InStr(";" & OUT_CATEGORIES & ";", ";" & Category & ";") = 0 Then
            VoucherProblem = "«Œ — »‰œ «·’—›."
        ElseIf Category = "EXPENSE" And IsNull(ExpenseTypeID) Then
            VoucherProblem = "«Œ — ‰Ê⁄ «·„’—Ê›."
        End If
    Else
        VoucherProblem = "‰Ê⁄ «·”‰œ €Ì— „⁄—Ê›: " & VoucherType
    End If
End Function

Private Function InsertVoucher(ByVal db As DAO.Database, ByVal VoucherType As String, ByVal BoxID As Long, _
                               ByVal ToBoxID As Variant, ByVal Category As String, ByVal Amount As Currency, _
                               ByVal PartyName As String, ByVal Description As String, _
                               ByVal ExpenseTypeID As Variant, ByVal ClosingID As Long, _
                               ByVal VoucherDate As Date) As Long
    ' Inside the caller's transaction. Returns the new CashVoucherID.
    Dim rs As DAO.Recordset, vNo As String, expenseID As Long, seq As String
    Select Case VoucherType
        Case "IN": seq = "CASH_IN"
        Case "OUT": seq = "CASH_OUT"
        Case Else: seq = "CASH_TRANSFER"
    End Select
    vNo = NextNumber(seq)

    If VoucherType = "OUT" And Category = "EXPENSE" Then
        ' the voucher moves the cash; the expense (without a box) feeds the expense reports
        Set rs = db.OpenRecordset("Expenses", dbOpenDynaset, dbAppendOnly)
        rs.AddNew
        rs!ExpenseNumber = NextNumber("EXPENSE")
        rs!ExpenseDate = DateValue(VoucherDate)
        rs!ExpenseTypeID = CLng(ExpenseTypeID)
        rs!Amount = Amount
        rs!Tax = 0
        rs!TotalAmount = Amount
        rs!PaymentMethodID = CASH_METHOD_ID
        rs!Description = Left$("”‰œ ’—› ‰ﬁœÌ… " & vNo & IIf(Len(Description) > 0, " - " & Description, ""), 255)
        rs!EmployeeID = CurrentUserID()
        rs!CostCenterID = CostCenterFor()                  ' modCostCenters
        rs.Update
        rs.Bookmark = rs.LastModified
        expenseID = rs!ExpenseID
        rs.Close
    End If

    Set rs = db.OpenRecordset("CashVouchers", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!VoucherNumber = vNo
    rs!VoucherDate = VoucherDate
    rs!VoucherType = VoucherType
    rs!CashBoxID = BoxID
    If VoucherType = "TRANSFER" Then
        rs!ToCashBoxID = CLng(ToBoxID)
        rs!Category = "TRANSFER"
        rs!InTransit = BoxConfirmsReceipt(ToBoxID)         ' enters the box when its custodian receives it
    Else
        rs!Category = Category
    End If
    rs!Amount = Amount
    If Len(Trim$(PartyName)) > 0 Then rs!PartyName = Left$(Trim$(PartyName), 100)
    If Len(Trim$(Description)) > 0 Then rs!Description = Left$(Trim$(Description), 255)
    If expenseID > 0 Then rs!ExpenseID = expenseID
    If ClosingID > 0 Then rs!ClosingID = ClosingID
    rs!EmployeeID = CurrentUserID()
    rs!CostCenterID = CostCenterFor()                      ' modCostCenters
    rs.Update
    rs.Bookmark = rs.LastModified
    InsertVoucher = rs!CashVoucherID
    rs.Close
End Function

Public Function ShortageEmployeeFor(ByVal BoxID As Long) As Variant
    ' The program user of the rep who keeps a custody box (docs/66): its closing shortage is an advance on him.
    ' Null = a box of the shop, or a rep with no user: the shortage stays a box shortage (5400).
    ShortageEmployeeFor = DbValue("SELECT r.EmployeeID FROM CashBoxes AS b INNER JOIN SalesReps AS r ON b.SalesRepID = " & _
                                  "r.SalesRepID WHERE b.BoxType = 'CUSTODY' AND r.EmployeeID Is Not Null AND b.CashBoxID = " & _
                                  BoxID)
End Function

Public Function PostCashClosing(ByVal BoxID As Long, ByVal Counted As Currency, ByVal Destination As String, _
                                ByVal ToBoxID As Variant, ByVal TransferAmount As Currency, ByVal Notes As String, _
                                ByRef NewClosingID As Long) As String
    ' Cashier closing. Book balance = balance at the last closing + cash in - cash out since then.
    ' Counted - book = overage (+) / shortage (-), booked with a voucher so the box matches the count.
    ' TransferAmount then goes to the main safe (MAIN) or to the owner (OWNER); the rest stays as float.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim periodStart As Variant, opening As Currency, cashIn As Currency, cashOut As Currency
    Dim expected As Currency, diff As Currency, closeDate As Date, closingNo As String
    Dim advanceTo As Variant, voucherID As Long

    On Error GoTo EH
    NewClosingID = 0
    If Not HasPermission("CASH_CLOSING") Then
        PostCashClosing = "·Ì”  ·œÌﬂ ’·«ÕÌ…  ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—."
        Exit Function
    End If
    If Not BoxIsActive(BoxID) Then
        PostCashClosing = "«Œ — «·’‰œÊﬁ."
        Exit Function
    End If
    If Not ClosingAllowed(BoxID) Then
        PostCashClosing = "Ì„ﬂ‰ﬂ  ’›Ì… ’‰«œÌﬁﬂ ›ﬁÿ."
        Exit Function
    End If
    If Counted < 0 Then
        PostCashClosing = "«·‰ﬁœÌ… «·›⁄·Ì… ·«  ﬂÊ‰ ”«·»…."
        Exit Function
    End If
    If Destination = "KEEP" Then TransferAmount = 0
    If TransferAmount < 0 Or TransferAmount > Counted Then
        PostCashClosing = "«·„»·€ «·„—ÕÛ¯· ÌÃ» √‰ ÌﬂÊ‰ »Ì‰ ’›— Ê«·‰ﬁœÌ… «·›⁄·Ì… (" & Money(Counted) & ")."
        Exit Function
    End If
    If Destination = "MAIN" And TransferAmount > 0 Then
        If Not BoxIsActive(ToBoxID) Then
            PostCashClosing = "«Œ — «·Œ“Ì‰… «· Ì  ” ·„ «·„»·€."
            Exit Function
        End If
        If CLng(ToBoxID) = BoxID Then
            PostCashClosing = "·« Ì„ﬂ‰ «· —ÕÌ· ≈·Ï ‰›” «·’‰œÊﬁ."
            Exit Function
        End If
    ElseIf Destination <> "MAIN" And Destination <> "OWNER" And Destination <> "KEEP" Then
        PostCashClosing = "«Œ — ÃÂ… «· —ÕÌ·."
        Exit Function
    End If

    periodStart = LastClosingDate(BoxID)
    If Not IsNull(periodStart) Then opening = CashBoxBalance(BoxID, periodStart, True)
    cashIn = CashMovesTotal(BoxID, "IN", periodStart)
    cashOut = CashMovesTotal(BoxID, "OUT", periodStart)
    expected = opening + cashIn - cashOut
    diff = Counted - expected

    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    closingNo = NextNumber("CASH_CLOSING")
    closeDate = Now
    Set rs = db.OpenRecordset("CashClosings", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ClosingNumber = closingNo
    rs!ClosingDate = closeDate
    rs!CashBoxID = BoxID
    rs!EmployeeID = CurrentUserID()
    rs!PeriodStart = periodStart
    rs!OpeningBalance = opening
    rs!CashIn = cashIn
    rs!CashOut = cashOut
    rs!ExpectedBalance = expected
    rs!CountedAmount = Counted
    rs!Difference = diff
    rs!Destination = Destination
    If Destination = "MAIN" And TransferAmount > 0 Then rs!ToCashBoxID = CLng(ToBoxID)
    rs!TransferAmount = TransferAmount
    rs!KeptAmount = Counted - TransferAmount
    If Len(Trim$(Notes)) > 0 Then rs!Notes = Left$(Trim$(Notes), 255)
    rs.Update
    rs.Bookmark = rs.LastModified
    NewClosingID = rs!ClosingID
    rs.Close

    If diff < 0 Then
        advanceTo = ShortageEmployeeFor(BoxID)
        If IsNull(advanceTo) Then
            InsertVoucher db, "OUT", BoxID, Null, "SHORTAGE", -diff, "", "⁄Ã“  ’›Ì… " & closingNo, Null, NewClosingID, closeDate
        Else
            ' a rep's custody box: the shortage is an advance on the rep (docs/66), deducted from his pay
            voucherID = InsertVoucher(db, "OUT", BoxID, Null, "ADVANCE", -diff, "", "⁄Ã“  ’›Ì… " & closingNo & _
                                      " ”·›… ⁄·Ï «·„‰œÊ»", Null, NewClosingID, closeDate)
            db.Execute "UPDATE CashVouchers SET AdvanceEmployeeID = " & CLng(advanceTo) & " WHERE CashVoucherID = " & _
                       voucherID, dbFailOnError
        End If
    ElseIf diff > 0 Then
        InsertVoucher db, "IN", BoxID, Null, "OVERAGE", diff, "", "“Ì«œ…  ’›Ì… " & closingNo, Null, NewClosingID, closeDate
    End If
    If TransferAmount > 0 Then
        If Destination = "MAIN" Then
            InsertVoucher db, "TRANSFER", BoxID, ToBoxID, "TRANSFER", TransferAmount, "", _
                          " —ÕÌ·  ’›Ì… " & closingNo, Null, NewClosingID, closeDate
        Else
            InsertVoucher db, "OUT", BoxID, Null, "OWNER", TransferAmount, "«·„«·ﬂ", _
                          " ”ÊÌ…  ’›Ì… " & closingNo, Null, NewClosingID, closeDate
        End If
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "CASH_CLOSING", "CashClosings", closingNo, "Expected=" & expected & " Counted=" & Counted
    Exit Function

EH:
    PostCashClosing = " ⁄–— Õ›Ÿ «· ’›Ì…: " & Err.Description & " (" & Err.Number & ")"
    NewClosingID = 0
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' frmTreasury: boxes, the day of the selected box, and the treasury actions
'==============================================================================
Public Sub TreasuryLoad(ByVal frm As Access.Form)
    Dim sql As String
    Calendar = vbCalGreg
    sql = "SELECT CashBoxID, BoxName AS [«·’‰œÊﬁ], BoxTypeName AS [«·‰Ê⁄], Format(q.Balance, '#,##0.00') AS [—’Ìœ «·’‰œÊﬁ] " & _
          "FROM CashBoxBalanceQuery AS q WHERE q.IsActive = True"
    If Not HasPermission("CASH_BOX") Then sql = sql & " AND q.CashBoxID = " & CurrentCashBoxID()   ' a cashier sees his box
    frm!lstBoxes.RowSource = Tr(sql & " ORDER BY q.BoxType DESC, q.BoxName")
    frm!btnCashIn.Enabled = HasPermission("CASH_BOX")
    frm!btnCashOut.Enabled = frm!btnCashIn.Enabled
    frm!btnTransfer.Enabled = frm!btnCashIn.Enabled
    frm!btnBoxes.Enabled = frm!btnCashIn.Enabled
    frm!btnClosing.Enabled = HasPermission("CASH_CLOSING")
    frm!txtDay.Value = Date
    If frm!lstBoxes.ListCount > 1 Then                 ' row 0 = headings
        frm!lstBoxes.Value = CurrentCashBoxID()
        If frm!lstBoxes.ListIndex < 0 Then frm!lstBoxes.Value = frm!lstBoxes.ItemData(1)
    End If
    TreasuryShow frm
End Sub

Public Sub TreasuryActivate(ByVal frm As Access.Form)
    ' Back from a voucher or a closing: balances changed.
    Dim keep As Variant
    keep = frm!lstBoxes.Value
    frm!lstBoxes.Requery
    frm!lstBoxes.Value = keep
    TreasuryShow frm
End Sub

Public Sub TreasuryShow(ByVal frm As Access.Form)
    Dim box As Long, d As Date, opening As Currency, cashIn As Currency, cashOut As Currency
    If IsNull(frm!lstBoxes.Value) Or Not IsDate(frm!txtDay.Value) Then
        frm!lblBoxName.Caption = Tr("·« ÌÊÃœ ’‰œÊﬁ")
        frm!lstMoves.RowSource = Tr("")
        Exit Sub
    End If
    box = CLng(frm!lstBoxes.Value)
    d = DateValue(frm!txtDay.Value)
    opening = CashBoxBalance(box, d)
    cashIn = DayTotal(box, d, "AmountIn")
    cashOut = DayTotal(box, d, "AmountOut")
    frm!lblBoxName.Caption = Tr(BoxName(box) & "  -  " & GDate(d))
    frm!lblOpening.Caption = Tr(Money(opening))
    frm!lblIn.Caption = Tr(Money(cashIn))
    frm!lblOut.Caption = Tr(Money(cashOut))
    frm!lblClosing.Caption = Tr(Money(opening + cashIn - cashOut))
    frm!lblCurrent.Caption = Tr("«·—’Ìœ «·Õ«·Ì ··’‰œÊﬁ: " & Money(CashBoxBalance(box)) & _
                             "    ¬Œ—  ’›Ì…: " & IIf(IsNull(LastClosingDate(box)), "·«  ÊÃœ", GDate(LastClosingDate(box), True)))
    frm!lstMoves.RowSource = Tr("SELECT Format(MoveDate, 'hh:nn') AS [«·Êﬁ ], MoveTypeName AS [«·Õ—ﬂ…], " & _
        "DocNumber AS [«·„” ‰œ], PartyName AS [«·ÃÂ… / «·»Ì«‰], " & _
        "IIf(AmountIn = 0, Null, Format(AmountIn, '#,##0.00')) AS [„ﬁ»Ê÷], " & _
        "IIf(AmountOut = 0, Null, Format(AmountOut, '#,##0.00')) AS [„œ›Ê⁄] " & _
        "FROM qryCashMovements WHERE CashBoxID = " & box & " AND MoveDate >= " & SqlDate(d) & _
        " AND MoveDate < " & SqlDate(DateAdd("d", 1, d)) & " ORDER BY MoveDate")
End Sub

Private Function DayTotal(ByVal BoxID As Long, ByVal d As Date, ByVal FieldName As String) As Currency
    DayTotal = Nz(DbValue("SELECT Sum(" & FieldName & ") FROM qryCashMovements WHERE CashBoxID = " & BoxID & _
                          " AND MoveDate >= " & SqlDate(d) & " AND MoveDate < " & SqlDate(DateAdd("d", 1, d))), 0)
End Function

Public Sub TreasuryDayStep(ByVal frm As Access.Form, ByVal Days As Integer)
    If Not IsDate(frm!txtDay.Value) Then frm!txtDay.Value = Date
    frm!txtDay.Value = DateAdd("d", Days, DateValue(frm!txtDay.Value))
    TreasuryShow frm
End Sub

Public Sub TreasuryToday(ByVal frm As Access.Form)
    frm!txtDay.Value = Date
    TreasuryShow frm
End Sub

Public Sub TreasuryOpen(ByVal frm As Access.Form, ByVal Which As String)
    ' Which: IN / OUT / TRANSFER (voucher), CLOSING
    Dim box As Variant
    box = frm!lstBoxes.Value
    If Which = "CLOSING" Then
        OpenScreen "frmCashClosing", 0, box
    Else
        OpenScreen "frmCashVoucher", 0, Which & "|" & Nz(box, "")
    End If
End Sub

Public Sub PrintCashDay(ByVal frm As Access.Form)
    ' The day book of the selected box on the shown day (docs/55): opening of the day, the receipts and the
    ' payments in detail, the totals, the closing of the day. A day without moves: its opening balance only.
    Dim d As Date, criteria As String
    If IsNull(frm!lstBoxes.Value) Or Not IsDate(frm!txtDay.Value) Then
        ShowWarning "«Œ — «·’‰œÊﬁ Ê«·ÌÊ„."
        Exit Sub
    End If
    d = DateValue(frm!txtDay.Value)
    SetPeriod d, d
    SetQueryParam "CashBoxID", CLng(frm!lstBoxes.Value)
    criteria = "«·ÌÊ„: " & GDate(d) & "    «·’‰œÊﬁ: " & BoxName(frm!lstBoxes.Value)
    If Nz(DbValue("SELECT COUNT(*) FROM CashDayBookQuery"), 0) > 0 Then
        OpenReportOrQuery "rptCashDayBook", "CashDayBookQuery", "", criteria
    Else
        OpenReportOrQuery "rptCashStatement", "CashStatementQuery", "", criteria
    End If
End Sub

'==============================================================================
' frmCashTransfers: the transfers in transit
'==============================================================================
Public Sub TransfersLoad(ByVal frm As Access.Form)
    frm!lstTransfers.RowSource = Tr("SELECT v.CashVoucherID, v.VoucherNumber AS [—ﬁ„ «·”‰œ], " & _
        "Format(v.VoucherDate, 'yyyy/mm/dd hh:nn') AS [«· «—ÌŒ], f.BoxName AS [„‰ ’‰œÊﬁ], t.BoxName AS [≈·Ï ’‰œÊﬁ], " & _
        "Format(v.Amount, '#,##0.00') AS [«·„»·€ «·„ÕÊÛ¯·], v.Description AS [»Ì«‰ «· ÕÊÌ·] " & _
        "FROM (CashVouchers AS v INNER JOIN [@CashBoxes] AS f ON v.CashBoxID = f.CashBoxID) " & _
        "INNER JOIN [@CashBoxes] AS t ON v.ToCashBoxID = t.CashBoxID " & _
        "WHERE v.VoucherType = 'TRANSFER' AND v.InTransit = True AND v.ReceivedAt Is Null ORDER BY v.VoucherDate")
    frm!lstTransfers.Requery
    frm!lblHint.Caption = Tr("«· ÕÊÌ· ÌœŒ· «·’‰œÊﬁ «·„” ·„ ⁄‰œ  √ﬂÌœ «” ·«„Â° ÊÕ Ï –·ﬂ ›Ì Õ”«» ´‰ﬁœÌ… »«·ÿ—Ìﬁª. " & _
                             "≈·€«¡ «· ÕÊÌ· Ì⁄Ìœ «·„»·€ ··’‰œÊﬁ «·„—”ˆ·.")
End Sub

Private Function SelectedTransfer(ByVal frm As Access.Form) As Long
    If IsNull(frm!lstTransfers.Value) Then
        ShowWarning "«Œ — «· ÕÊÌ· „‰ «·ﬁ«∆„…."
    Else
        SelectedTransfer = CLng(frm!lstTransfers.Value)
    End If
End Function

Public Sub TransfersReceive(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = SelectedTransfer(frm)
    If id = 0 Then Exit Sub
    msg = ReceiveCashTransfer(id)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo " „ «” ·«„ «· ÕÊÌ· ÊœŒ· «·„»·€ «·’‰œÊﬁ."
        TransfersLoad frm
    End If
End Sub

Public Sub TransfersCancel(ByVal frm As Access.Form)
    Dim id As Long, msg As String
    id = SelectedTransfer(frm)
    If id = 0 Then Exit Sub
    If Not AskYesNo("≈·€«¡ «· ÕÊÌ· Ê≈⁄«œ… «·„»·€ ··’‰œÊﬁ «·„—”ˆ·ø") Then Exit Sub
    msg = CancelCashTransfer(id)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        TransfersLoad frm
    End If
End Sub

Public Sub TransfersPrint(ByVal frm As Access.Form)
    Dim id As Long
    id = SelectedTransfer(frm)
    If id > 0 Then PrintCashVoucher id
End Sub

'==============================================================================
' frmCashVoucher: OpenArgs = "IN|box" / "OUT|box" / "TRANSFER|box"
'==============================================================================
Public Sub VoucherLoad(ByVal frm As Access.Form)
    Dim parts() As String, kind As String
    frm!cboVoucherType.RowSource = Tr(VOUCHER_TYPES)
    frm!cboBox.RowSource = Tr(BOX_ROWS & " ORDER BY q.BoxType DESC, q.BoxName")
    frm!cboToBox.RowSource = Tr(frm!cboBox.RowSource)
    kind = "OUT"
    If Not IsNull(frm.OpenArgs) Then
        parts = Split(CStr(frm.OpenArgs) & "|", "|")
        If Len(parts(0)) > 0 Then kind = parts(0)
        If IsNumeric(parts(1)) Then frm!cboBox.Value = CLng(parts(1))
    End If
    If IsNull(frm!cboBox.Value) And CurrentCashBoxID() > 0 Then frm!cboBox.Value = CurrentCashBoxID()
    frm!cboVoucherType.Value = kind
    VoucherTypeChanged frm
End Sub

Public Sub VoucherTypeChanged(ByVal frm As Access.Form)
    Dim kind As String
    kind = Nz(frm!cboVoucherType.Value, "OUT")
    Select Case kind
        Case "IN"
            frm!cboCategory.RowSource = Tr(IN_CATEGORIES)
            frm!cboCategory.Value = "OTHER"
            frm!lblBox.Caption = Tr("Ìıﬁ»÷ ›Ì ’‰œÊﬁ")
            frm!lblParty.Caption = Tr("«” ·„‰« „‰")
        Case "OUT"
            frm!cboCategory.RowSource = Tr(OUT_CATEGORIES)
            frm!cboCategory.Value = "EXPENSE"
            frm!lblBox.Caption = Tr("Ìı’—› „‰ ’‰œÊﬁ")
            frm!lblParty.Caption = Tr("Ìı’—› ≈·Ï")
        Case Else
            frm!cboCategory.RowSource = Tr("TRANSFER; ÕÊÌ· »Ì‰ «·’‰«œÌﬁ")
            frm!cboCategory.Value = "TRANSFER"
            frm!lblBox.Caption = Tr("„‰ ’‰œÊﬁ")
            frm!lblParty.Caption = Tr("«·„” ·„")
    End Select
    frm!lblTitle.Caption = Tr(DLookupList(VOUCHER_TYPES, kind))
    frm!cboToBox.Visible = (kind = "TRANSFER")
    frm!lblToBox.Visible = frm!cboToBox.Visible
    frm!cboCategory.Enabled = (kind <> "TRANSFER")
    VoucherCategoryChanged frm
    VoucherBoxChanged frm
End Sub

Public Sub VoucherCategoryChanged(ByVal frm As Access.Form)
    Dim isExpense As Boolean
    isExpense = (Nz(frm!cboCategory.Value, "") = "EXPENSE" And Nz(frm!cboVoucherType.Value, "") = "OUT")
    frm!cboExpenseType.Visible = isExpense
    frm!lblExpenseType.Visible = isExpense
    frm!btnNewExpenseType.Visible = isExpense
    ' an employee advance (or its payback): the employee (modPayroll); a commission: the sales rep (modSalesReps)
    Select Case Nz(frm!cboCategory.Value, "")
        Case "ADVANCE"
            If frm!cboEmployee.RowSource <> EMPLOYEE_ROWS Then
                frm!cboEmployee.RowSource = Tr(EMPLOYEE_ROWS)
                frm!cboEmployee.Value = Null
            End If
            frm!lblEmployee.Caption = Tr("«·„ÊŸ› ’«Õ» «·”·›… *")
            frm!cboEmployee.Visible = True
        Case "COMMISSION"
            If frm!cboEmployee.RowSource <> Tr(REP_ROWS) Then
                frm!cboEmployee.RowSource = Tr(REP_ROWS)
                frm!cboEmployee.Value = Null
            End If
            frm!lblEmployee.Caption = Tr("«·„‰œÊ» *")
            frm!cboEmployee.Visible = True
        Case Else
            frm!cboEmployee.Visible = False
    End Select
    frm!lblEmployee.Visible = frm!cboEmployee.Visible
End Sub

Public Sub VoucherBoxChanged(ByVal frm As Access.Form)
    If IsNull(frm!cboBox.Value) Then
        frm!lblBoxBalance.Caption = Tr(" ")
    Else
        frm!lblBoxBalance.Caption = Tr("—’Ìœ «·’‰œÊﬁ «·¬‰: " & Money(CashBoxBalance(CLng(frm!cboBox.Value))))
    End If
End Sub

Private Function DLookupList(ByVal List As String, ByVal Code As String) As String
    ' Caption of a code in a "CODE;caption;CODE;caption" value list.
    Dim items() As String, i As Long
    items = Split(List, ";")
    For i = 0 To UBound(items) - 1 Step 2
        If items(i) = Code Then
            DLookupList = items(i + 1)
            Exit Function
        End If
    Next
End Function

Public Function SaveCashVoucher(ByVal frm As Access.Form, Optional ByVal PrintAfter As Boolean = False) As Boolean
    Dim msg As String, newID As Long, kind As String
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    kind = Nz(frm!cboVoucherType.Value, "")
    msg = AdvanceEmployeeProblem(kind, Nz(frm!cboCategory.Value, ""), frm!cboEmployee.Value, Nz(frm!txtAmount.Value, 0))
    If Len(msg) = 0 Then
        msg = CommissionVoucherProblem(kind, Nz(frm!cboCategory.Value, ""), frm!cboEmployee.Value, Nz(frm!txtAmount.Value, 0))
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    msg = PostCashVoucher(kind, Nz(frm!cboBox.Value, 0), frm!cboToBox.Value, Nz(frm!cboCategory.Value, ""), _
                          Nz(frm!txtAmount.Value, 0), Nz(frm!txtParty.Value, ""), Nz(frm!txtDescription.Value, ""), _
                          frm!cboExpenseType.Value, newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    If Nz(frm!cboCategory.Value, "") = "ADVANCE" Then
        CurrentDb.Execute "UPDATE CashVouchers SET AdvanceEmployeeID = " & CLng(frm!cboEmployee.Value) & _
                          " WHERE CashVoucherID = " & newID, dbFailOnError
    ElseIf Nz(frm!cboCategory.Value, "") = "COMMISSION" Then
        CurrentDb.Execute "UPDATE CashVouchers SET SalesRepID = " & CLng(frm!cboEmployee.Value) & _
                          " WHERE CashVoucherID = " & newID, dbFailOnError                  ' modSalesReps
    End If
    ShowInfo " „ Õ›Ÿ " & DLookupList(VOUCHER_TYPES, kind) & " —ﬁ„ " & _
             DLookup("VoucherNumber", "CashVouchers", "CashVoucherID = " & newID)
    If PrintAfter Then PrintCashVoucher newID
    frm!txtAmount.Value = Null
    frm!txtParty.Value = Null
    frm!txtDescription.Value = Null
    VoucherBoxChanged frm
    SaveCashVoucher = True
End Function

Public Sub AddExpenseType(ByVal frm As Access.Form, ByVal ControlName As String)
    ' Adds a new expense type to the list and selects it (frmCashVoucher, frmExpenses).
    Dim typeName As String, id As Variant
    typeName = Trim$(InputBox(Tr("«”„ ‰Ê⁄ «·„’—Ê› «·ÃœÌœ:"), Tr("‰Ê⁄ „’—Ê› ÃœÌœ")))
    If Len(typeName) = 0 Then Exit Sub
    id = DbValue("SELECT ExpenseTypeID FROM ExpenseTypes WHERE ExpenseTypeName = " & SqlText(Left$(typeName, 50)))
    If IsNull(id) Then
        CurrentDb.Execute "INSERT INTO ExpenseTypes (ExpenseTypeName) VALUES (" & SqlText(Left$(typeName, 50)) & ")", _
                          dbFailOnError
        id = DbValue("SELECT ExpenseTypeID FROM ExpenseTypes WHERE ExpenseTypeName = " & SqlText(Left$(typeName, 50)))
        LogAction "SAVE", "ExpenseTypes", CStr(id), typeName
    End If
    frm.Controls(ControlName).Requery
    frm.Controls(ControlName).Value = id
End Sub

'==============================================================================
' frmCashClosing: OpenArgs = box
'==============================================================================
Public Sub ClosingLoad(ByVal frm As Access.Form)
    Dim sql As String
    Calendar = vbCalGreg
    sql = BOX_ROWS
    If Not HasPermission("CASH_ALL_BOXES") Then sql = sql & " AND q.CashBoxID IN (" & ClosingBoxes() & ")"
    frm!cboBox.RowSource = Tr(sql & " ORDER BY q.BoxType, q.BoxName")
    frm!cboToBox.RowSource = Tr(BOX_ROWS & " AND q.BoxType = 'MAIN' ORDER BY q.BoxName")
    frm!cboDestination.RowSource = Tr(DESTINATIONS)
    frm!cboDestination.Value = "MAIN"
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then frm!cboBox.Value = CLng(frm.OpenArgs)
    End If
    If Not ClosingAllowed(frm!cboBox.Value) Then frm!cboBox.Value = CurrentCashBoxID()
    frm!cboToBox.Value = DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True")
    ClosingBoxChanged frm
    ClosingDestinationChanged frm
End Sub

Public Sub ClosingBoxChanged(ByVal frm As Access.Form)
    Dim box As Long, since As Variant, opening As Currency, cashIn As Currency, cashOut As Currency
    frm!txtCounted.Value = Null
    frm!txtTransfer.Value = Null
    If IsNull(frm!cboBox.Value) Then
        frm!lstBreakdown.RowSource = Tr("")
        frm!lblPeriod.Caption = Tr(" ")
        frm!lblOpening.Caption = "-": frm!lblCashIn.Caption = "-"
        frm!lblCashOut.Caption = "-": frm!lblExpected.Caption = "-"
        frm!txtExpected.Value = Null
        ClosingRecalc frm
        Exit Sub
    End If
    box = CLng(frm!cboBox.Value)
    since = LastClosingDate(box)
    If Not IsNull(since) Then opening = CashBoxBalance(box, since, True)
    cashIn = CashMovesTotal(box, "IN", since)
    cashOut = CashMovesTotal(box, "OUT", since)
    frm!lblPeriod.Caption = Tr(IIf(IsNull(since), "„‰ »œ«Ì… «·’‰œÊﬁ (·«  ÊÃœ  ’›Ì… ”«»ﬁ…)", _
                                "„‰ ¬Œ—  ’›Ì…: " & GDate(since, True)) & "  Õ Ï «·¬‰")
    frm!lblOpening.Caption = Tr(Money(opening))
    frm!lblCashIn.Caption = Tr(Money(cashIn))
    frm!lblCashOut.Caption = Tr(Money(cashOut))
    frm!lblExpected.Caption = Tr(Money(opening + cashIn - cashOut))
    frm!txtExpected.Value = opening + cashIn - cashOut
    frm!lstBreakdown.RowSource = Tr("SELECT MoveTypeName AS [«·»‰œ], Count(*) AS [«·⁄œœ], " & _
        "Format(Sum(AmountIn), '#,##0.00') AS [„ﬁ»Ê÷], Format(Sum(AmountOut), '#,##0.00') AS [„œ›Ê⁄] " & _
        "FROM qryCashMovements WHERE CashBoxID = " & box & _
        IIf(IsNull(since), "", " AND MoveDate > " & SqlDate(CDate(Nz(since, Now)))) & _
        " GROUP BY MoveTypeName ORDER BY MoveTypeName")
    ClosingRecalc frm
End Sub

Public Sub ClosingCountChanged(ByVal frm As Access.Form)
    ' The whole count goes to the safe unless the user types another amount.
    If Nz(frm!cboDestination.Value, "") = "KEEP" Then
        frm!txtTransfer.Value = 0
    Else
        frm!txtTransfer.Value = frm!txtCounted.Value
    End If
    ClosingRecalc frm
End Sub

Public Sub ClosingDestinationChanged(ByVal frm As Access.Form)
    frm!cboToBox.Visible = (Nz(frm!cboDestination.Value, "") = "MAIN")
    frm!lblToBox.Visible = frm!cboToBox.Visible
    frm!txtTransfer.Enabled = (Nz(frm!cboDestination.Value, "") <> "KEEP")
    ClosingCountChanged frm
End Sub

Public Sub ClosingRecalc(ByVal frm As Access.Form)
    Dim expected As Currency, counted As Currency, diff As Currency, kept As Currency
    If IsNull(frm!txtCounted.Value) Or IsNull(frm!cboBox.Value) Then
        frm!lblDifference.Caption = Tr("√œŒ· «·‰ﬁœÌ… «·„ÊÃÊœ… ›⁄·« ›Ì «·’‰œÊﬁ")
        frm!lblDifference.ForeColor = CLR_MUTED
        frm!lblKept.Caption = Tr(" ")
        Exit Sub
    End If
    expected = Nz(frm!txtExpected.Value, 0)
    counted = Nz(frm!txtCounted.Value, 0)
    diff = counted - expected
    If diff < 0 Then
        frm!lblDifference.Caption = Tr("⁄Ã“: " & Money(-diff))
        frm!lblDifference.ForeColor = CLR_DANGER
    ElseIf diff > 0 Then
        frm!lblDifference.Caption = Tr("“Ì«œ…: " & Money(diff))
        frm!lblDifference.ForeColor = CLR_WARNING
    Else
        frm!lblDifference.Caption = Tr("„ÿ«»ﬁ: ·« ÌÊÃœ ⁄Ã“ Ê·« “Ì«œ…")
        frm!lblDifference.ForeColor = CLR_SUCCESS
    End If
    kept = counted - Nz(frm!txtTransfer.Value, 0)
    frm!lblKept.Caption = Tr("Ì»ﬁÏ ›Ì «·’‰œÊﬁ (⁄Âœ… «·ﬂ«‘Ì—): " & Money(kept))
End Sub

Public Function SaveCashClosing(ByVal frm As Access.Form, Optional ByVal PrintAfter As Boolean = False) As Boolean
    Dim msg As String, newID As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If IsNull(frm!txtCounted.Value) Then
        ShowWarning "√œŒ· «·‰ﬁœÌ… «·›⁄·Ì… «·„ÊÃÊœ… ›Ì «·’‰œÊﬁ."
        SafeFocus frm!txtCounted
        Exit Function
    End If
    If Not AskYesNo("Õ›Ÿ  ’›Ì… " & frm!cboBox.Column(1) & "ø" & vbCrLf & frm!lblDifference.Caption & vbCrLf & _
                    "·« Ì„ﬂ‰  ⁄œÌ· «· ’›Ì… »⁄œ «·Õ›Ÿ.") Then Exit Function
    msg = PostCashClosing(Nz(frm!cboBox.Value, 0), Nz(frm!txtCounted.Value, 0), Nz(frm!cboDestination.Value, ""), _
                          frm!cboToBox.Value, Nz(frm!txtTransfer.Value, 0), Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ «· ’›Ì… " & DLookup("ClosingNumber", "CashClosings", "ClosingID = " & newID)
    If PrintAfter Then PrintCashClosing newID
    frm!txtNotes.Value = Null
    ClosingBoxChanged frm
    SaveCashClosing = True
End Function

'==============================================================================
' Printing
'==============================================================================
Public Sub PrintCashVoucher(ByVal CashVoucherID As Variant)
    If IsNull(CashVoucherID) Then Exit Sub
    PrintCashDocument "rptCashVoucher", "[DocID] = " & CLng(CashVoucherID)
End Sub

Public Sub PrintCashClosing(ByVal ClosingID As Variant)
    If IsNull(ClosingID) Then Exit Sub
    PrintCashDocument "rptCashClosing", "[ClosingID] = " & CLng(ClosingID)
End Sub

Private Sub PrintCashDocument(ByVal ReportName As String, ByVal WhereCondition As String)
    If Not ReportExists(ReportName) Then
        ShowWarning " ﬁ—Ì— «·ÿ»«⁄… €Ì— „ÊÃÊœ: " & ReportName & vbCrLf & "‘€¯· BuildReports."
        Exit Sub
    End If
    On Error GoTo EH
    TempVars.Add "ReportCriteria", ""
    DoCmd.OpenReport ReportName, POS_PRINT_VIEW, , WhereCondition
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError " ⁄–— › Õ «· ﬁ—Ì—: " & Err.Description
End Sub

'==============================================================================
' In-Access test (RunAllTests): vouchers and a closing inside a transaction that is rolled back
'==============================================================================
Public Function TestCash() As Boolean
    Dim ws As DAO.Workspace, box As Long, mainBox As Long, before As Currency, msg As String, id As Long
    Dim passed As Long, failed As Long, report As String, expBefore As Long, inTrans As Boolean
    Dim mainBefore As Currency
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestCash  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    mainBox = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True"), 0)
    box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'CASHIER' AND IsActive = True"), 0)
    CheckCash mainBox > 0 And box > 0, "Œ“Ì‰… —∆Ì”Ì… Ê’‰œÊﬁ ﬂ«‘Ì— ‰‘ÿ«‰", passed, failed, report
    CheckCash IsNull(CashBoxFor(2, 100)), "«·œ›⁄ »«·»ÿ«ﬁ… ·« ÌœŒ· «·’‰œÊﬁ", passed, failed, report
    CheckCash IsNull(CashBoxFor(CASH_METHOD_ID, 0)), "«·„»·€ «·¬Ã· ·« ÌœŒ· «·’‰œÊﬁ", passed, failed, report
    CheckCash Nz(CashBoxFor(CASH_METHOD_ID, 100), 0) = CurrentCashBoxID() And CurrentCashBoxID() > 0, _
              "«·œ›⁄ «·‰ﬁœÌ ÌœŒ· ’‰œÊﬁ «·„” Œœ„", passed, failed, report
    If failed > 0 Then GoTo Done

    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    before = CashBoxBalance(box)
    expBefore = Nz(DbValue("SELECT COUNT(*) FROM Expenses"), 0)
    msg = PostCashVoucher("IN", box, Null, "OWNER", 1000, "TEST", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 1000, "”‰œ ﬁ»÷ Ì“Ìœ «·’‰œÊﬁ 1000 " & msg, passed, failed, report
    msg = PostCashVoucher("OUT", box, Null, "EXPENSE", 100, "TEST", "TEST-CASH", 1, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 900, "”‰œ ’—› „’—Ê› Ì‰ﬁ’ «·’‰œÊﬁ 100 " & msg, passed, failed, report
    CheckCash Nz(DbValue("SELECT COUNT(*) FROM Expenses"), 0) = expBefore + 1, "”‰œ ’—› «·„’—Ê› Ì”Ã· „’—Ê›«", _
              passed, failed, report
    msg = PostCashVoucher("OUT", box, Null, "OTHER", before + 100000, "TEST", "TEST-CASH", Null, id)
    CheckCash Len(msg) > 0, "·« Ìı’—› √ﬂÀ— „‰ —’Ìœ «·’‰œÊﬁ", passed, failed, report
    msg = PostCashVoucher("TRANSFER", box, box, "TRANSFER", 10, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) > 0, "·«  ÕÊÌ· ·‰›” «·’‰œÊﬁ", passed, failed, report
    msg = PostCashVoucher("TRANSFER", box, mainBox, "TRANSFER", 100, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 800, "«· ÕÊÌ· Ì‰ﬁ’ «·’‰œÊﬁ " & msg, passed, failed, report

    ' closing: 5 short, 500 to the main safe, the rest stays as float
    msg = PostCashClosing(box, before + 795, "MAIN", mainBox, 500, "TEST-CASH", id)
    CheckCash Len(msg) = 0, "Õ›Ÿ «· ’›Ì… " & msg, passed, failed, report
    CheckCash Nz(DbValue("SELECT Difference FROM CashClosings WHERE ClosingID = " & id), 0) = -5, "⁄Ã“ «· ’›Ì… = 5", _
              passed, failed, report
    CheckCash CashBoxBalance(box) = before + 295, "«·’‰œÊﬁ »⁄œ «· ’›Ì… = «·„⁄œÊœ - «·„—ÕÛ¯·", passed, failed, report
    CheckCash Nz(DbValue("SELECT COUNT(*) FROM CashVouchers WHERE ClosingID = " & id), 0) = 2, _
              "«· ’›Ì…  ‰‘∆ ”‰œ ⁄Ã“ Ê”‰œ  ÕÊÌ·", passed, failed, report
    ' the day book of today (docs/55): opening + receipts - payments = the balance now, with the closing
    SetPeriod Date, Date
    SetQueryParam "CashBoxID", box
    CheckCash Nz(DbValue("SELECT Max(DayOpening) + Sum(AmountIn) - Sum(AmountOut) FROM CashDayBookQuery"), 0) = _
              CashBoxBalance(box), "ÌÊ„Ì… «·’‰œÊﬁ: —’Ìœ ¬Œ— «·ÌÊ„ = —’Ìœ «·’‰œÊﬁ", passed, failed, report
    CheckCash Nz(DbValue("SELECT Max(ClosingCount) FROM CashDayBookQuery"), 0) >= 1, _
              "ÌÊ„Ì… «·’‰œÊﬁ:  ’›Ì… «·ÌÊ„  ŸÂ—", passed, failed, report

    ' custodians and transfers in transit (docs/56)
    CurrentDb.Execute "INSERT INTO CashBoxUsers (CashBoxID, EmployeeID) VALUES (" & box & ", " & CurrentUserID() & ")", _
                      dbFailOnError
    CheckCash IsBoxCustodian(box) And BoxHasCustodians(box) And CanUseBox(box) And ClosingAllowed(box), _
              "√„Ì‰ «·’‰œÊﬁ Ì’—› „‰Â ÊÌ’›¯ÌÂ", passed, failed, report
    CurrentDb.Execute "UPDATE CashBoxes SET ConfirmReceipt = True WHERE CashBoxID = " & mainBox, dbFailOnError
    mainBefore = CashBoxBalance(mainBox)
    before = CashBoxBalance(box)
    msg = PostCashVoucher("TRANSFER", box, mainBox, "TRANSFER", 50, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before - 50 And CashBoxBalance(mainBox) = mainBefore, _
              " ÕÊÌ· ·’‰œÊﬁ Ìƒﬂœ «·«” ·«„: ÌŒ—Ã „‰ «·„—”ˆ· Ê·« ÌœŒ· «·„” ·„ " & msg, passed, failed, report
    CheckCash Len(ReceiveCashTransfer(id)) = 0 And CashBoxBalance(mainBox) = mainBefore + 50, _
              " √ﬂÌœ «·«” ·«„ ÌœŒ· «·„»·€ «·’‰œÊﬁ", passed, failed, report
    CheckCash Len(ReceiveCashTransfer(id)) > 0, "·« Ìı” ·„ «· ÕÊÌ· „— Ì‰", passed, failed, report
    msg = PostCashVoucher("TRANSFER", box, mainBox, "TRANSFER", 30, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And Len(CancelCashTransfer(id)) = 0 And CashBoxBalance(box) = before - 50 And _
              Nz(DbValue("SELECT COUNT(*) FROM CashVouchers WHERE CashVoucherID = " & id), 0) = 0, _
              "≈·€«¡ «· ÕÊÌ· ﬁ»· «” ·«„Â Ì⁄Ìœ «·„»·€ " & msg, passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckCash False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·Œ“Ì‰… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestCash"
        TestCash = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCash"
    End If
End Function

Private Sub CheckCash(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                      ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "[X]  " & Title
    End If
End Sub
