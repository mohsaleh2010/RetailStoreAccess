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
' Screens: frmTreasury, frmCashVoucher, frmCashClosing (built by modBuildForms).
' Posting functions return "" on success or an Arabic error message.
'==============================================================================
Option Compare Database
Option Explicit

Public Const CASH_METHOD_ID As Long = 1          ' PaymentMethods: نقدي

Private Const BOX_ROWS As String = "SELECT CashBoxID, BoxName FROM CashBoxes WHERE IsActive = True"
Private Const VOUCHER_TYPES As String = "IN;سند قبض نقدية;OUT;سند صرف نقدية;TRANSFER;تحويل بين الصناديق"
Private Const IN_CATEGORIES As String = "OTHER;قبض نقدية (إيرادات أخرى);OWNER;إيداع من المالك;" & _
                                        "ADVANCE;سداد سلفة موظف"
Private Const OUT_CATEGORIES As String = "EXPENSE;مصروف (يُسجل في المصروفات);OWNER;تسوية / مسحوبات المالك;" & _
                                         "ADVANCE;سلفة موظف;OTHER;صرف آخر"
Private Const DESTINATIONS As String = "MAIN;ترحيل إلى الخزينة الرئيسية;OWNER;تسليم للمالك (تسوية);" & _
                                       "KEEP;يبقى في الصندوق"

'==============================================================================
' Boxes and balances
'==============================================================================
Public Function CurrentCashBoxID() As Long
    Dim v As Variant
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
    BoxName = Nz(DbValue("SELECT BoxName FROM CashBoxes WHERE CashBoxID = " & CLng(BoxID)), "")
End Function

Public Function BoxIsActive(ByVal BoxID As Variant) As Boolean
    If IsNull(BoxID) Then Exit Function
    BoxIsActive = Nz(DbValue("SELECT COUNT(*) FROM CashBoxes WHERE IsActive = True AND CashBoxID = " & CLng(BoxID)), 0) > 0
End Function

Private Function Money(ByVal Value As Currency) As String
    Money = Format$(Value, "#,##0.00")
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
        PostCashVoucher = "ليست لديك صلاحية سندات الخزينة."
        Exit Function
    End If
    msg = VoucherProblem(VoucherType, BoxID, ToBoxID, Category, Amount, ExpenseTypeID)
    If Len(msg) = 0 And VoucherType <> "IN" Then
        balance = CashBoxBalance(BoxID)
        If Amount > balance Then msg = "رصيد " & BoxName(BoxID) & " (" & Money(balance) & ") لا يكفي لصرف " & Money(Amount) & "."
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
    PostCashVoucher = "تعذر حفظ السند: " & Err.Description & " (" & Err.Number & ")"
    NewVoucherID = 0
    If inTrans Then ws.Rollback
End Function

Public Function VoucherProblem(ByVal VoucherType As String, ByVal BoxID As Long, ByVal ToBoxID As Variant, _
                               ByVal Category As String, ByVal Amount As Currency, _
                               ByVal ExpenseTypeID As Variant) As String
    ' The checks of a voucher entered by the user ("" = fine). Balances are checked by the caller.
    If Amount <= 0 Then
        VoucherProblem = "أدخل المبلغ (أكبر من صفر)."
    ElseIf Not BoxIsActive(BoxID) Then
        VoucherProblem = "اختر الصندوق."
    ElseIf VoucherType = "TRANSFER" Then
        If Not BoxIsActive(ToBoxID) Then
            VoucherProblem = "اختر الصندوق المحوَّل إليه."
        ElseIf CLng(ToBoxID) = BoxID Then
            VoucherProblem = "لا يمكن التحويل إلى نفس الصندوق."
        End If
    ElseIf VoucherType = "IN" Then
        If InStr(";" & IN_CATEGORIES & ";", ";" & Category & ";") = 0 Then VoucherProblem = "اختر بند القبض."
    ElseIf VoucherType = "OUT" Then
        If InStr(";" & OUT_CATEGORIES & ";", ";" & Category & ";") = 0 Then
            VoucherProblem = "اختر بند الصرف."
        ElseIf Category = "EXPENSE" And IsNull(ExpenseTypeID) Then
            VoucherProblem = "اختر نوع المصروف."
        End If
    Else
        VoucherProblem = "نوع السند غير معروف: " & VoucherType
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
        rs!Description = Left$("سند صرف نقدية " & vNo & IIf(Len(Description) > 0, " - " & Description, ""), 255)
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

Public Function PostCashClosing(ByVal BoxID As Long, ByVal Counted As Currency, ByVal Destination As String, _
                                ByVal ToBoxID As Variant, ByVal TransferAmount As Currency, ByVal Notes As String, _
                                ByRef NewClosingID As Long) As String
    ' Cashier closing. Book balance = balance at the last closing + cash in - cash out since then.
    ' Counted - book = overage (+) / shortage (-), booked with a voucher so the box matches the count.
    ' TransferAmount then goes to the main safe (MAIN) or to the owner (OWNER); the rest stays as float.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim periodStart As Variant, opening As Currency, cashIn As Currency, cashOut As Currency
    Dim expected As Currency, diff As Currency, closeDate As Date, closingNo As String

    On Error GoTo EH
    NewClosingID = 0
    If Not HasPermission("CASH_CLOSING") Then
        PostCashClosing = "ليست لديك صلاحية تصفية يومية الكاشير."
        Exit Function
    End If
    If Not BoxIsActive(BoxID) Then
        PostCashClosing = "اختر الصندوق."
        Exit Function
    End If
    If BoxID <> CurrentCashBoxID() And Not HasPermission("CASH_BOX") Then
        PostCashClosing = "يمكنك تصفية صندوقك فقط."
        Exit Function
    End If
    If Counted < 0 Then
        PostCashClosing = "النقدية الفعلية لا تكون سالبة."
        Exit Function
    End If
    If Destination = "KEEP" Then TransferAmount = 0
    If TransferAmount < 0 Or TransferAmount > Counted Then
        PostCashClosing = "المبلغ المرحَّل يجب أن يكون بين صفر والنقدية الفعلية (" & Money(Counted) & ")."
        Exit Function
    End If
    If Destination = "MAIN" And TransferAmount > 0 Then
        If Not BoxIsActive(ToBoxID) Then
            PostCashClosing = "اختر الخزينة التي تستلم المبلغ."
            Exit Function
        End If
        If CLng(ToBoxID) = BoxID Then
            PostCashClosing = "لا يمكن الترحيل إلى نفس الصندوق."
            Exit Function
        End If
    ElseIf Destination <> "MAIN" And Destination <> "OWNER" And Destination <> "KEEP" Then
        PostCashClosing = "اختر جهة الترحيل."
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
        InsertVoucher db, "OUT", BoxID, Null, "SHORTAGE", -diff, "", "عجز تصفية " & closingNo, Null, NewClosingID, closeDate
    ElseIf diff > 0 Then
        InsertVoucher db, "IN", BoxID, Null, "OVERAGE", diff, "", "زيادة تصفية " & closingNo, Null, NewClosingID, closeDate
    End If
    If TransferAmount > 0 Then
        If Destination = "MAIN" Then
            InsertVoucher db, "TRANSFER", BoxID, ToBoxID, "TRANSFER", TransferAmount, "", _
                          "ترحيل تصفية " & closingNo, Null, NewClosingID, closeDate
        Else
            InsertVoucher db, "OUT", BoxID, Null, "OWNER", TransferAmount, "المالك", _
                          "تسوية تصفية " & closingNo, Null, NewClosingID, closeDate
        End If
    End If
    ws.CommitTrans
    inTrans = False
    LogAction "CASH_CLOSING", "CashClosings", closingNo, "Expected=" & expected & " Counted=" & Counted
    Exit Function

EH:
    PostCashClosing = "تعذر حفظ التصفية: " & Err.Description & " (" & Err.Number & ")"
    NewClosingID = 0
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' frmTreasury: boxes, the day of the selected box, and the treasury actions
'==============================================================================
Public Sub TreasuryLoad(ByVal frm As Access.Form)
    Dim sql As String
    Calendar = vbCalGreg
    sql = "SELECT CashBoxID, BoxName AS [الصندوق], BoxTypeName AS [النوع], Format(Balance, '#,##0.00') AS [الرصيد] " & _
          "FROM CashBoxBalanceQuery WHERE IsActive = True"
    If Not HasPermission("CASH_BOX") Then sql = sql & " AND CashBoxID = " & CurrentCashBoxID()   ' a cashier sees his box
    frm!lstBoxes.RowSource = sql & " ORDER BY BoxType DESC, BoxName"
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
        frm!lblBoxName.Caption = "لا يوجد صندوق"
        frm!lstMoves.RowSource = ""
        Exit Sub
    End If
    box = CLng(frm!lstBoxes.Value)
    d = DateValue(frm!txtDay.Value)
    opening = CashBoxBalance(box, d)
    cashIn = DayTotal(box, d, "AmountIn")
    cashOut = DayTotal(box, d, "AmountOut")
    frm!lblBoxName.Caption = BoxName(box) & "  -  " & GDate(d)
    frm!lblOpening.Caption = Money(opening)
    frm!lblIn.Caption = Money(cashIn)
    frm!lblOut.Caption = Money(cashOut)
    frm!lblClosing.Caption = Money(opening + cashIn - cashOut)
    frm!lblCurrent.Caption = "الرصيد الحالي للصندوق: " & Money(CashBoxBalance(box)) & _
                             "    آخر تصفية: " & IIf(IsNull(LastClosingDate(box)), "لا توجد", GDate(LastClosingDate(box), True))
    frm!lstMoves.RowSource = "SELECT Format(MoveDate, 'hh:nn') AS [الوقت], MoveTypeName AS [الحركة], " & _
        "DocNumber AS [المستند], PartyName AS [الجهة / البيان], " & _
        "IIf(AmountIn = 0, Null, Format(AmountIn, '#,##0.00')) AS [مقبوض], " & _
        "IIf(AmountOut = 0, Null, Format(AmountOut, '#,##0.00')) AS [مدفوع] " & _
        "FROM qryCashMovements WHERE CashBoxID = " & box & " AND MoveDate >= " & SqlDate(d) & _
        " AND MoveDate < " & SqlDate(DateAdd("d", 1, d)) & " ORDER BY MoveDate"
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
    ' Movements of the selected box on the shown day: opening balance, receipts, payments, closing balance.
    Dim d As Date
    If IsNull(frm!lstBoxes.Value) Or Not IsDate(frm!txtDay.Value) Then
        ShowWarning "اختر الصندوق واليوم."
        Exit Sub
    End If
    d = DateValue(frm!txtDay.Value)
    SetPeriod d, d
    SetQueryParam "CashBoxID", CLng(frm!lstBoxes.Value)
    OpenReportOrQuery "rptCashStatement", "CashStatementQuery", "", _
                      "اليوم: " & GDate(d) & "    الصندوق: " & BoxName(frm!lstBoxes.Value)
End Sub

'==============================================================================
' frmCashVoucher: OpenArgs = "IN|box" / "OUT|box" / "TRANSFER|box"
'==============================================================================
Public Sub VoucherLoad(ByVal frm As Access.Form)
    Dim parts() As String, kind As String
    frm!cboVoucherType.RowSource = VOUCHER_TYPES
    frm!cboBox.RowSource = BOX_ROWS & " ORDER BY BoxType DESC, BoxName"
    frm!cboToBox.RowSource = frm!cboBox.RowSource
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
            frm!cboCategory.RowSource = IN_CATEGORIES
            frm!cboCategory.Value = "OTHER"
            frm!lblBox.Caption = "يُقبض في صندوق"
            frm!lblParty.Caption = "استلمنا من"
        Case "OUT"
            frm!cboCategory.RowSource = OUT_CATEGORIES
            frm!cboCategory.Value = "EXPENSE"
            frm!lblBox.Caption = "يُصرف من صندوق"
            frm!lblParty.Caption = "يُصرف إلى"
        Case Else
            frm!cboCategory.RowSource = "TRANSFER;تحويل بين الصناديق"
            frm!cboCategory.Value = "TRANSFER"
            frm!lblBox.Caption = "من صندوق"
            frm!lblParty.Caption = "المستلم"
    End Select
    frm!lblTitle.Caption = DLookupList(VOUCHER_TYPES, kind)
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
    ' an employee advance (or its payback): the employee (modPayroll)
    frm!cboEmployee.Visible = (Nz(frm!cboCategory.Value, "") = "ADVANCE")
    frm!lblEmployee.Visible = frm!cboEmployee.Visible
End Sub

Public Sub VoucherBoxChanged(ByVal frm As Access.Form)
    If IsNull(frm!cboBox.Value) Then
        frm!lblBoxBalance.Caption = " "
    Else
        frm!lblBoxBalance.Caption = "رصيد الصندوق الآن: " & Money(CashBoxBalance(CLng(frm!cboBox.Value)))
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
    End If
    ShowInfo "تم حفظ " & DLookupList(VOUCHER_TYPES, kind) & " رقم " & _
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
    typeName = Trim$(InputBox("اسم نوع المصروف الجديد:", "نوع مصروف جديد"))
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
    If Not HasPermission("CASH_BOX") Then sql = sql & " AND CashBoxID = " & CurrentCashBoxID()
    frm!cboBox.RowSource = sql & " ORDER BY BoxType, BoxName"
    frm!cboToBox.RowSource = BOX_ROWS & " AND BoxType = 'MAIN' ORDER BY BoxName"
    frm!cboDestination.RowSource = DESTINATIONS
    frm!cboDestination.Value = "MAIN"
    If Not IsNull(frm.OpenArgs) Then
        If IsNumeric(frm.OpenArgs) Then frm!cboBox.Value = CLng(frm.OpenArgs)
    End If
    If IsNull(frm!cboBox.Value) Or Not HasPermission("CASH_BOX") Then frm!cboBox.Value = CurrentCashBoxID()
    frm!cboToBox.Value = DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True")
    ClosingBoxChanged frm
    ClosingDestinationChanged frm
End Sub

Public Sub ClosingBoxChanged(ByVal frm As Access.Form)
    Dim box As Long, since As Variant, opening As Currency, cashIn As Currency, cashOut As Currency
    frm!txtCounted.Value = Null
    frm!txtTransfer.Value = Null
    If IsNull(frm!cboBox.Value) Then
        frm!lstBreakdown.RowSource = ""
        frm!lblPeriod.Caption = " "
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
    frm!lblPeriod.Caption = IIf(IsNull(since), "من بداية الصندوق (لا توجد تصفية سابقة)", _
                                "من آخر تصفية: " & GDate(since, True)) & "  حتى الآن"
    frm!lblOpening.Caption = Money(opening)
    frm!lblCashIn.Caption = Money(cashIn)
    frm!lblCashOut.Caption = Money(cashOut)
    frm!lblExpected.Caption = Money(opening + cashIn - cashOut)
    frm!txtExpected.Value = opening + cashIn - cashOut
    frm!lstBreakdown.RowSource = "SELECT MoveTypeName AS [البند], Count(*) AS [العدد], " & _
        "Format(Sum(AmountIn), '#,##0.00') AS [مقبوض], Format(Sum(AmountOut), '#,##0.00') AS [مدفوع] " & _
        "FROM qryCashMovements WHERE CashBoxID = " & box & _
        IIf(IsNull(since), "", " AND MoveDate > " & SqlDate(CDate(Nz(since, Now)))) & _
        " GROUP BY MoveTypeName ORDER BY MoveTypeName"
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
        frm!lblDifference.Caption = "أدخل النقدية الموجودة فعلًا في الصندوق"
        frm!lblDifference.ForeColor = CLR_MUTED
        frm!lblKept.Caption = " "
        Exit Sub
    End If
    expected = Nz(frm!txtExpected.Value, 0)
    counted = Nz(frm!txtCounted.Value, 0)
    diff = counted - expected
    If diff < 0 Then
        frm!lblDifference.Caption = "عجز: " & Money(-diff)
        frm!lblDifference.ForeColor = CLR_DANGER
    ElseIf diff > 0 Then
        frm!lblDifference.Caption = "زيادة: " & Money(diff)
        frm!lblDifference.ForeColor = CLR_WARNING
    Else
        frm!lblDifference.Caption = "مطابق: لا يوجد عجز ولا زيادة"
        frm!lblDifference.ForeColor = CLR_SUCCESS
    End If
    kept = counted - Nz(frm!txtTransfer.Value, 0)
    frm!lblKept.Caption = "يبقى في الصندوق (عهدة الكاشير): " & Money(kept)
End Sub

Public Function SaveCashClosing(ByVal frm As Access.Form, Optional ByVal PrintAfter As Boolean = False) As Boolean
    Dim msg As String, newID As Long
    If Not CanScreenAction(frm.Name, "ADD") Then Exit Function      ' frmUserScreens
    If IsNull(frm!txtCounted.Value) Then
        ShowWarning "أدخل النقدية الفعلية الموجودة في الصندوق."
        SafeFocus frm!txtCounted
        Exit Function
    End If
    If Not AskYesNo("حفظ تصفية " & frm!cboBox.Column(1) & "؟" & vbCrLf & frm!lblDifference.Caption & vbCrLf & _
                    "لا يمكن تعديل التصفية بعد الحفظ.") Then Exit Function
    msg = PostCashClosing(Nz(frm!cboBox.Value, 0), Nz(frm!txtCounted.Value, 0), Nz(frm!cboDestination.Value, ""), _
                          frm!cboToBox.Value, Nz(frm!txtTransfer.Value, 0), Nz(frm!txtNotes.Value, ""), newID)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo "تم حفظ التصفية " & DLookup("ClosingNumber", "CashClosings", "ClosingID = " & newID)
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
        ShowWarning "تقرير الطباعة غير موجود: " & ReportName & vbCrLf & "شغّل BuildReports."
        Exit Sub
    End If
    On Error GoTo EH
    TempVars.Add "ReportCriteria", ""
    DoCmd.OpenReport ReportName, POS_PRINT_VIEW, , WhereCondition
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError "تعذر فتح التقرير: " & Err.Description
End Sub

'==============================================================================
' In-Access test (RunAllTests): vouchers and a closing inside a transaction that is rolled back
'==============================================================================
Public Function TestCash() As Boolean
    Dim ws As DAO.Workspace, box As Long, mainBox As Long, before As Currency, msg As String, id As Long
    Dim passed As Long, failed As Long, report As String, expBefore As Long, inTrans As Boolean
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestCash  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    mainBox = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'MAIN' AND IsActive = True"), 0)
    box = Nz(DbValue("SELECT Min(CashBoxID) FROM CashBoxes WHERE BoxType = 'CASHIER' AND IsActive = True"), 0)
    CheckCash mainBox > 0 And box > 0, "خزينة رئيسية وصندوق كاشير نشطان", passed, failed, report
    CheckCash IsNull(CashBoxFor(2, 100)), "الدفع بالبطاقة لا يدخل الصندوق", passed, failed, report
    CheckCash IsNull(CashBoxFor(CASH_METHOD_ID, 0)), "المبلغ الآجل لا يدخل الصندوق", passed, failed, report
    CheckCash Nz(CashBoxFor(CASH_METHOD_ID, 100), 0) = CurrentCashBoxID() And CurrentCashBoxID() > 0, _
              "الدفع النقدي يدخل صندوق المستخدم", passed, failed, report
    If failed > 0 Then GoTo Done

    On Error GoTo EH
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    before = CashBoxBalance(box)
    expBefore = Nz(DbValue("SELECT COUNT(*) FROM Expenses"), 0)
    msg = PostCashVoucher("IN", box, Null, "OWNER", 1000, "TEST", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 1000, "سند قبض يزيد الصندوق 1000 " & msg, passed, failed, report
    msg = PostCashVoucher("OUT", box, Null, "EXPENSE", 100, "TEST", "TEST-CASH", 1, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 900, "سند صرف مصروف ينقص الصندوق 100 " & msg, passed, failed, report
    CheckCash Nz(DbValue("SELECT COUNT(*) FROM Expenses"), 0) = expBefore + 1, "سند صرف المصروف يسجل مصروفًا", _
              passed, failed, report
    msg = PostCashVoucher("OUT", box, Null, "OTHER", before + 100000, "TEST", "TEST-CASH", Null, id)
    CheckCash Len(msg) > 0, "لا يُصرف أكثر من رصيد الصندوق", passed, failed, report
    msg = PostCashVoucher("TRANSFER", box, box, "TRANSFER", 10, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) > 0, "لا تحويل لنفس الصندوق", passed, failed, report
    msg = PostCashVoucher("TRANSFER", box, mainBox, "TRANSFER", 100, "", "TEST-CASH", Null, id)
    CheckCash Len(msg) = 0 And CashBoxBalance(box) = before + 800, "التحويل ينقص الصندوق " & msg, passed, failed, report

    ' closing: 5 short, 500 to the main safe, the rest stays as float
    msg = PostCashClosing(box, before + 795, "MAIN", mainBox, 500, "TEST-CASH", id)
    CheckCash Len(msg) = 0, "حفظ التصفية " & msg, passed, failed, report
    CheckCash Nz(DbValue("SELECT Difference FROM CashClosings WHERE ClosingID = " & id), 0) = -5, "عجز التصفية = 5", _
              passed, failed, report
    CheckCash CashBoxBalance(box) = before + 295, "الصندوق بعد التصفية = المعدود - المرحَّل", passed, failed, report
    CheckCash Nz(DbValue("SELECT COUNT(*) FROM CashVouchers WHERE ClosingID = " & id), 0) = 2, _
              "التصفية تنشئ سند عجز وسند تحويل", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckCash False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الخزينة ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestCash"
        TestCash = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCash"
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
