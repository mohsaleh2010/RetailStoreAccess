Attribute VB_Name = "modOpening"
'==============================================================================
' modOpening  -  Retail Store Management System
'
' The start-up opening balances (docs/52-Opening-Balances.md, screen frmOpeningBalances,
' permission OPENING_BALANCES):
'   Settings.OpeningDate  the start-up date: every opening balance is an entry of that day
'                         (customers and suppliers: qryCustomerOpening / qrySupplierOpening,
'                         the cash boxes and banks get it as their OpeningDate).
'                         Not after the first document, not in a closed period.
'   OpeningBalances       the general accounts (capital, loans, accruals...): each line is an entry
'                         ACCOUNT_OPENING against the opening balances account 3900.
'   Account 3900          "total opening balances": the other side of every opening balance. Its
'                         balance is zero when the balances entered are balanced; a difference stays
'                         there, visible on the screen by source (qryOpeningTotals).
' The customers, suppliers, cash boxes and banks keep their OpeningBalance field; the screen edits it
' in one grid per part, with the closed period check and the audit trail.
'   OpeningStock          the stock (docs/53-Opening-Stock.md): quantity and unit cost of each product, entered
'                         in one grid or imported from Excel. "Post" records one stock move "opening" (type 8)
'                         per line on the start-up date (journal: stock 1400 against 3900); a posted line is
'                         locked, and can be unposted while it is the only move of its product.
'==============================================================================
Option Compare Database
Option Explicit

Public Const OPENING_ACCOUNT As Long = 3900
Private Const OPENING_SOURCES As String = "'BOX_OPENING','CUSTOMER_OPENING','SUPPLIER_OPENING','BANK_OPENING','ACCOUNT_OPENING'"

'------------------------------------------------------------------------------
' Rules
'------------------------------------------------------------------------------
Public Function OpeningDateValue() As Variant
    ' The start-up date, Null when it was never set.
    Dim v As Variant
    v = SettingValue("OpeningDate")
    If IsDate(v) Then
        OpeningDateValue = DateValue(v)
    Else
        OpeningDateValue = Null
    End If
End Function

Public Function OpeningEntryDate(ByVal CreatedAt As Variant) As Variant
    ' The day of the opening balance of a customer or supplier: the start-up date, else the day of the card.
    OpeningEntryDate = OpeningDateValue()
    If IsNull(OpeningEntryDate) Then OpeningEntryDate = CreatedAt
End Function

Public Function CanOpening(ByVal Action As String) As String
    If Not HasPermission("OPENING_BALANCES") Then
        CanOpening = "لا تملك صلاحية أرصدة بداية التشغيل."
    ElseIf Not CanScreenAction("frmOpeningBalances", Action, True) Then
        CanOpening = "لا تملك صلاحية " & IIf(Action = "ADD", "الإضافة", IIf(Action = "EDIT", "التعديل", "الحذف")) & _
                     " في شاشة أرصدة بداية التشغيل."
    End If
End Function

Public Function FirstDocumentDate() As Variant
    ' The day of the first operation that is not an opening balance (from the journal), Null when none.
    ' An opening stock move (type 8) is an opening balance too.
    FirstDocumentDate = DbValue("SELECT Min(EntryDate) FROM JournalEntries WHERE SourceType NOT IN (" & _
                                OPENING_SOURCES & ") AND NOT (SourceType = 'STOCK_MOVE' AND SourceID IN " & _
                                "(SELECT TransactionID FROM InventoryTransactions WHERE TransactionTypeID = " & _
                                TT_OPENING & "))")
End Function

Public Function OpeningDateProblem(ByVal NewDate As Variant) As String
    ' "" when the start-up date can be NewDate.
    Dim first As Variant, old As Variant
    If Not IsDate(NewDate) Then
        OpeningDateProblem = "أدخل تاريخ بدء التشغيل."
        Exit Function
    End If
    first = FirstDocumentDate()
    If IsDate(first) Then
        If DateValue(NewDate) > DateValue(first) Then
            OpeningDateProblem = "تاريخ بدء التشغيل لا يكون بعد أول مستند (" & GDate(first) & ")."
            Exit Function
        End If
    End If
    OpeningDateProblem = ClosedPeriodProblem(NewDate)
    If Len(OpeningDateProblem) > 0 Then Exit Function
    old = OpeningDateValue()
    If IsDate(old) Then OpeningDateProblem = ClosedPeriodProblem(old)       ' the entries of the old date move
End Function

Public Function OpeningAccountProblem(ByVal AccountCode As Variant) As String
    ' The general accounts only: the others have their opening balance in their own part.
    Dim code As Long
    If IsNull(AccountCode) Then
        OpeningAccountProblem = "اختر الحساب."
        Exit Function
    End If
    code = CLng(AccountCode)
    Select Case code
        Case 1300: OpeningAccountProblem = "رصيد العملاء يُدخل في قسم «العملاء»: رصيد لكل عميل."
        Case 2100: OpeningAccountProblem = "رصيد الموردين يُدخل في قسم «الموردون»: رصيد لكل مورد."
        Case 1400: OpeningAccountProblem = "رصيد المخزون يُدخل في قسم «المخزون»: كمية وتكلفة لكل صنف."
        Case OPENING_ACCOUNT: OpeningAccountProblem = "حساب إجمالي الأرصدة الافتتاحية هو مقابل كل الأرصدة، ولا يُدخل له رصيد."
        Case 110001 To 119999: OpeningAccountProblem = "رصيد الصندوق يُدخل في قسم «الصناديق»."
        Case 120001 To 129999: OpeningAccountProblem = "رصيد البنك يُدخل في قسم «البنوك»."
        Case Else
            If Nz(DbValue("SELECT IsPosting FROM Accounts WHERE AccountCode = " & code), False) = False Then
                OpeningAccountProblem = "اختر حسابًا فرعيًا (حساب تُسجَّل عليه قيود)."
            End If
    End Select
End Function

Public Function SaveOpeningDate(ByVal NewDate As Variant) As String
    ' The start-up date; the cash boxes and banks that had the old date (or none) move with it, and so
    ' does the opening stock posted from the screen.
    Dim old As Variant, before As Collection, cond As String
    SaveOpeningDate = CanOpening("EDIT")
    If Len(SaveOpeningDate) > 0 Then Exit Function
    SaveOpeningDate = SyncJournal()               ' the first document must be known
    If Len(SaveOpeningDate) > 0 Then Exit Function
    SaveOpeningDate = OpeningDateProblem(NewDate)
    If Len(SaveOpeningDate) > 0 Then Exit Function
    old = OpeningDateValue()
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    CurrentDb.Execute "UPDATE Settings SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE SettingID = 1", _
                      dbFailOnError
    cond = "OpeningDate Is Null"
    If IsDate(old) Then cond = cond & " OR OpeningDate = " & SqlDate(old)
    CurrentDb.Execute "UPDATE CashBoxes SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE " & cond, dbFailOnError
    CurrentDb.Execute "UPDATE Banks SET OpeningDate = " & SqlDate(DateValue(NewDate)) & " WHERE " & cond, dbFailOnError
    CurrentDb.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(DateValue(NewDate)) & _
                      " WHERE TransactionID IN (SELECT TransactionID FROM OpeningStock)", dbFailOnError
    AuditEdited "OPENING_DATE", "Settings", "SettingID", 1, before
    SaveOpeningDate = SyncJournal()
End Function

Public Function OpeningBalanceTotal() As Currency
    ' The balance of account 3900 (debit - credit); zero when the opening balances are balanced.
    OpeningBalanceTotal = Nz(DbValue("SELECT Sum(SumDebit - SumCredit) FROM qryOpeningTotals"), 0)
End Function

Public Function OpeningTotalText(ByVal Total As Currency) As String
    Dim acc As String
    acc = "حساب إجمالي الأرصدة الافتتاحية (" & OPENING_ACCOUNT & ")"
    If Total = 0 Then
        OpeningTotalText = acc & ": الرصيد صفر، الأرصدة الافتتاحية متوازنة."
    ElseIf Total > 0 Then
        OpeningTotalText = acc & ": رصيده " & Format$(Total, "#,##0.00") & " مدين." & vbCrLf & _
                           "الأرصدة الدائنة (رأس المال والالتزامات) أكبر من المدينة بهذا المبلغ."
    Else
        OpeningTotalText = acc & ": رصيده " & Format$(-Total, "#,##0.00") & " دائن." & vbCrLf & _
                           "الأرصدة المدينة (الأصول) أكبر من الدائنة بهذا المبلغ: أكمل رأس المال أو الالتزامات."
    End If
End Function

'------------------------------------------------------------------------------
' Screen frmOpeningBalances
'------------------------------------------------------------------------------
Public Sub OpeningLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    frm!txtOpeningDate.Value = OpeningDateValue()
    OpeningShowPart frm, "ACCOUNTS"
    OpeningRefresh frm
End Sub

Public Sub OpeningSaveDate(ByVal frm As Access.Form)
    Dim msg As String
    If Not IsDate(frm!txtOpeningDate.Value) Then
        ShowWarning "أدخل تاريخ بدء التشغيل."
        SafeFocus frm!txtOpeningDate
        Exit Sub
    End If
    If Not AskYesNo("كل الأرصدة الافتتاحية تصبح بتاريخ " & GDate(frm!txtOpeningDate.Value) & ". متابعة؟") Then Exit Sub
    msg = SaveOpeningDate(frm!txtOpeningDate.Value)
    If Len(msg) > 0 Then
        ShowWarning msg
        frm!txtOpeningDate.Value = OpeningDateValue()
        Exit Sub
    End If
    ShowInfo "تم حفظ تاريخ بدء التشغيل."
    OpeningShowPart frm, Nz(frm!txtPart.Value, "ACCOUNTS")
    OpeningRefresh frm
End Sub

Public Sub OpeningRefresh(ByVal frm As Access.Form)
    ' The balance of 3900 by source, after bringing the entries up to date.
    Dim msg As String, total As Currency
    msg = SyncJournal()
    If Len(msg) > 0 Then ShowWarning msg
    frm!lstSources.RowSource = Tr("SELECT TypeName AS [المصدر], Format(SumDebit, '#,##0.00') AS [مدين], " & _
                                  "Format(SumCredit, '#,##0.00') AS [دائن] FROM qryOpeningTotals ORDER BY SourceType")
    total = OpeningBalanceTotal()
    frm!lblTotal.Caption = Tr(OpeningTotalText(total))
    frm!lblTotal.ForeColor = IIf(total = 0, CLR_SUCCESS, CLR_DANGER)
End Sub

Public Sub OpeningShowPart(ByVal frm As Access.Form, ByVal Part As String)
    ' One grid per part: the general accounts, the customers, the suppliers, the cash boxes, the banks.
    Dim grid As String, caps As Variant, hint As String
    Select Case Part
        Case "CUSTOMERS"
            grid = "frmOpeningCustomers"
            caps = Array("العميل", "الرصيد الافتتاحي", "", "الرصيد الحالي")
            hint = "رصيد كل عميل يوم بدء التشغيل: موجب = العميل مدين للمحل (عليه)، سالب = دائن (له)."
        Case "SUPPLIERS"
            grid = "frmOpeningSuppliers"
            caps = Array("المورد", "الرصيد الافتتاحي", "", "الرصيد الحالي")
            hint = "رصيد كل مورد يوم بدء التشغيل: موجب = المحل مدين للمورد (له)، سالب = دائن (عليه)."
        Case "BOXES"
            grid = "frmOpeningBoxes"
            caps = Array("الصندوق", "الرصيد الافتتاحي", "", "تاريخ الرصيد")
            hint = "النقدية في كل صندوق يوم بدء التشغيل."
        Case "BANKS"
            grid = "frmOpeningBanks"
            caps = Array("البنك", "الرصيد الافتتاحي", "", "تاريخ الرصيد")
            hint = "رصيد كل حساب بنكي يوم بدء التشغيل: موجب = لنا في البنك، سالب = سحب على المكشوف."
        Case "STOCK"
            grid = "frmOpeningStock"
            caps = Array("الصنف والمخزن", "الكمية", "تكلفة الوحدة", "القيمة والحالة")
            hint = "كمية كل صنف وتكلفة وحدته يوم بدء التشغيل، ثم «ترحيل الرصيد». «كل الأصناف» يضيف الأصناف " & _
                   "التي ليس لها حركات، و«قالب Excel» يصدّرها لتملأها ثم تستوردها."
        Case Else
            Part = "ACCOUNTS"
            grid = "frmOpeningAccountLines"
            caps = Array("الحساب", "مدين", "دائن", "البيان")
            hint = "رصيد كل حساب عام يوم بدء التشغيل: رأس المال، والقروض، والمستحقات، والأصول الأخرى. " & _
                   "العملاء والموردون والصناديق والبنوك والمخزون في أقسامها."
    End Select
    frm!txtPart.Value = Part
    If frm!subGrid.SourceObject <> grid Then frm!subGrid.SourceObject = grid
    frm!lblCol1.Caption = Tr(caps(0))
    frm!lblCol2.Caption = Tr(caps(1))
    frm!lblCol3.Caption = Tr(IIf(Len(caps(2)) = 0, " ", caps(2)))
    frm!lblCol4.Caption = Tr(caps(3))
    frm!btnStockFill.Visible = (Part = "STOCK")
    frm!btnStockImport.Visible = (Part = "STOCK")
    frm!btnStockTemplate.Visible = (Part = "STOCK")
    frm!btnStockPost.Visible = (Part = "STOCK")
    If IsNull(OpeningDateValue()) Then
        hint = "احفظ تاريخ بدء التشغيل أولًا، ثم أدخل الأرصدة."
        frm!subGrid.Enabled = False
    Else
        frm!subGrid.Enabled = True
    End If
    frm!lblHint.Caption = Tr(hint)
End Sub

'------------------------------------------------------------------------------
' Grids
'------------------------------------------------------------------------------
Public Sub OpeningLineBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    ' A line of the general accounts (OpeningBalances).
    Dim msg As String
    msg = CanOpening(IIf(frm.NewRecord, "ADD", "EDIT"))
    If Len(msg) = 0 And IsNull(OpeningDateValue()) Then msg = "احفظ تاريخ بدء التشغيل أولًا."
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    If Len(msg) = 0 Then msg = OpeningAccountProblem(frm!AccountCode.Value)
    If Len(msg) = 0 Then
        If Nz(frm!Debit.Value, 0) < 0 Or Nz(frm!Credit.Value, 0) < 0 Or _
           (Nz(frm!Debit.Value, 0) <> 0 And Nz(frm!Credit.Value, 0) <> 0) Or _
           Nz(frm!Debit.Value, 0) + Nz(frm!Credit.Value, 0) = 0 Then
            msg = "اكتب المبلغ في المدين أو في الدائن فقط."
        End If
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    frm!Debit.Value = Nz(frm!Debit.Value, 0)
    frm!Credit.Value = Nz(frm!Credit.Value, 0)
    AuditFormBefore frm, "OpeningBalances", "OpeningLineID"
End Sub

Public Sub OpeningPartyBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer, ByVal TableName As String, _
                                    ByVal KeyField As String)
    ' The opening balance of a customer, supplier, cash box or bank.
    Dim msg As String, oldDay As Variant, delta As Currency
    msg = CanOpening("EDIT")
    If Len(msg) = 0 And IsNull(OpeningDateValue()) Then msg = "احفظ تاريخ بدء التشغيل أولًا."
    If Len(msg) = 0 Then
        If TableName = "Customers" Or TableName = "Suppliers" Then
            oldDay = OpeningEntryDate(frm!CreatedAt.Value)
        Else
            oldDay = frm!OpeningDate.Value
        End If
        msg = ClosedPeriodProblem(oldDay)
        If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    frm!OpeningBalance.Value = Nz(frm!OpeningBalance.Value, 0)
    If TableName = "Customers" Or TableName = "Suppliers" Then
        ' the stored current balance follows (the reference stays the balance queries)
        delta = Nz(frm!OpeningBalance.Value, 0) - Nz(frm!OpeningBalance.OldValue, 0)
        frm!CurrentBalance.Value = Nz(frm!CurrentBalance.Value, 0) + delta
    ElseIf IsNull(frm!OpeningDate.Value) Then
        frm!OpeningDate.Value = OpeningDateValue()
    End If
    AuditFormBefore frm, TableName, KeyField
End Sub

Public Sub OpeningGridAfterUpdate(ByVal frm As Access.Form)
    AuditFormAfter frm
    On Error Resume Next                      ' the grid can be opened alone
    frm.Parent!lblTotal.Caption = Tr("تغيرت الأرصدة: اضغط «تحديث الإجمالي».")
    frm.Parent!lblTotal.ForeColor = CLR_MUTED
End Sub

Public Sub OpeningRemoveLine(ByVal frm As Access.Form)
    Dim id As Variant, msg As String, before As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!OpeningLineID.Value
    If IsNull(id) Then Exit Sub
    msg = CanOpening("DELETE")
    If Len(msg) = 0 Then msg = ClosedPeriodProblem(OpeningDateValue())
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Not AskYesNo("حذف رصيد هذا الحساب؟") Then Exit Sub
    If frm.Dirty Then frm.Undo
    Set before = AuditSnapshot("OpeningBalances", "OpeningLineID", id)
    CurrentDb.Execute "DELETE FROM OpeningBalances WHERE OpeningLineID = " & CLng(id), dbFailOnError
    AuditDeleted "OPENING_DELETE", "OpeningBalances", id, before
    frm.Requery
    OpeningGridAfterUpdate frm
End Sub

'------------------------------------------------------------------------------
' The opening stock (OpeningStock, grid frmOpeningStock)
'------------------------------------------------------------------------------
Public Function ProductHasMoves(ByVal ProductID As Long, Optional ByVal WarehouseID As Long = 0) As Boolean
    ' Moves of the product in a warehouse (0 = in any warehouse).
    Dim sql As String
    sql = "SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = " & ProductID
    If WarehouseID > 0 Then sql = sql & " AND WarehouseID = " & WarehouseID
    ProductHasMoves = Nz(DbValue(sql), 0) > 0
End Function

Public Function OpeningStockReady(ByVal Action As String) As String
    ' "" when the opening stock can be changed now: the permission, the start-up date, its period open.
    OpeningStockReady = CanOpening(Action)
    If Len(OpeningStockReady) = 0 And IsNull(OpeningDateValue()) Then
        OpeningStockReady = "احفظ تاريخ بدء التشغيل أولًا."
    End If
    If Len(OpeningStockReady) = 0 Then OpeningStockReady = ClosedPeriodProblem(OpeningDateValue())
End Function

Public Function OpeningStockLineProblem(ByVal LineID As Variant, ByVal ProductID As Variant, ByVal Qty As Variant, _
                                        ByVal UnitCost As Variant, Optional ByVal WarehouseID As Variant = 1) As String
    ' "" when a line not posted yet can hold this product, quantity and cost in this warehouse (docs/57).
    Dim other As Variant, wh As Long
    If IsNull(ProductID) Then
        OpeningStockLineProblem = "اختر الصنف."
        Exit Function
    End If
    wh = Nz(WarehouseID, MAIN_WAREHOUSE_ID)
    If Not ProductTracksStock(CLng(ProductID)) Then
        OpeningStockLineProblem = "هذا الصنف لا يتابع المخزون، فليس له رصيد."
    ElseIf Not WarehouseIsActive(wh) Then
        OpeningStockLineProblem = WarehouseProblem(wh)
    ElseIf ProductHasMoves(CLng(ProductID), wh) Then
        OpeningStockLineProblem = "للصنف حركات في هذا المخزن: الرصيد الافتتاحي يكون قبل أي حركة. صحّح الكمية بالجرد."
    ElseIf Nz(Qty, 0) < 0 Or Nz(UnitCost, 0) < 0 Then
        OpeningStockLineProblem = "الكمية والتكلفة لا تكونان سالبتين."
    Else
        other = DbValue("SELECT OpeningStockID FROM OpeningStock WHERE ProductID = " & CLng(ProductID) & _
                        " AND WarehouseID = " & wh & " AND OpeningStockID <> " & Nz(LineID, 0))
        If Not IsNull(other) Then OpeningStockLineProblem = "الصنف موجود في سطر آخر لنفس المخزن."
    End If
End Function

Public Sub OpeningStockBeforeUpdate(ByVal frm As Access.Form, ByRef Cancel As Integer)
    Dim msg As String
    msg = OpeningStockReady(IIf(frm.NewRecord, "ADD", "EDIT"))
    If Len(msg) = 0 And Not IsNull(frm!TransactionID.Value) Then
        msg = "هذا الرصيد مرحّل. لتعديله ألغِ ترحيله بزر الحذف في سطره، ثم عدّله ورحّله من جديد."
    End If
    If IsNull(frm!WarehouseID.Value) Then frm!WarehouseID.Value = MAIN_WAREHOUSE_ID
    If Len(msg) = 0 Then
        msg = OpeningStockLineProblem(frm!OpeningStockID.Value, frm!ProductID.Value, frm!Quantity.Value, _
                                      frm!UnitCost.Value, frm!WarehouseID.Value)
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Cancel = True
        Exit Sub
    End If
    frm!Quantity.Value = Nz(frm!Quantity.Value, 0)
    frm!UnitCost.Value = Nz(frm!UnitCost.Value, 0)
    AuditFormBefore frm, "OpeningStock", "OpeningStockID"
End Sub

Public Function FillOpeningStock(Optional ByVal WarehouseID As Long = 1) As Long
    ' Adds a line (quantity 0, the known cost) in the warehouse for every active product that tracks
    ' stock and has no line and no move there.
    Dim before As Long
    before = Nz(DbValue("SELECT COUNT(*) FROM OpeningStock"), 0)
    CurrentDb.Execute "INSERT INTO OpeningStock (ProductID, WarehouseID, Quantity, UnitCost) " & _
        "SELECT p.ProductID, " & WarehouseID & ", 0, IIf(p.AverageCost > 0, p.AverageCost, p.PurchasePrice) " & _
        "FROM Products AS p WHERE p.TrackStock = True AND p.IsActive = True " & _
        "AND p.ProductID NOT IN (SELECT ProductID FROM OpeningStock WHERE WarehouseID = " & WarehouseID & ") " & _
        "AND p.ProductID NOT IN (SELECT ProductID FROM InventoryTransactions WHERE WarehouseID = " & WarehouseID & ")", _
        dbFailOnError
    FillOpeningStock = Nz(DbValue("SELECT COUNT(*) FROM OpeningStock"), 0) - before
End Function

Public Function PostOpeningStock(ByRef Posted As Long, ByRef Skipped As Long) As String
    ' One stock move "opening" per line not posted with a quantity, on the start-up date.
    Dim db As DAO.Database, ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean
    Dim refNo As String, pid As Long, tid As Variant, startDay As Date, wh As Long
    Posted = 0
    Skipped = 0
    PostOpeningStock = OpeningStockReady("ADD")
    If Len(PostOpeningStock) > 0 Then Exit Function
    If Nz(DbValue("SELECT COUNT(*) FROM OpeningStock WHERE TransactionID Is Null AND Quantity > 0"), 0) = 0 Then
        PostOpeningStock = "لا توجد أسطر بكمية لم تُرحَّل."
        Exit Function
    End If
    On Error GoTo EH
    startDay = OpeningDateValue()
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    refNo = NextNumber("STOCK_ADJUST")
    Set rs = db.OpenRecordset("SELECT OpeningStockID, ProductID, WarehouseID, Quantity, UnitCost FROM OpeningStock " & _
                              "WHERE TransactionID Is Null AND Quantity > 0 ORDER BY OpeningStockID", dbOpenSnapshot)
    Do Until rs.EOF
        pid = rs!ProductID
        wh = Nz(rs!WarehouseID, MAIN_WAREHOUSE_ID)
        If ProductTracksStock(pid) And Not ProductHasMoves(pid, wh) Then
            ApplyStockMovement db, pid, CCur(rs!Quantity), TT_OPENING, CCur(rs!UnitCost), "MANUAL", 0, refNo, "", True, wh
            tid = DbValue("SELECT Max(TransactionID) FROM InventoryTransactions WHERE ProductID = " & pid & _
                          " AND WarehouseID = " & wh)
            db.Execute "UPDATE InventoryTransactions SET TransactionDate = " & SqlDate(startDay) & _
                       " WHERE TransactionID = " & tid, dbFailOnError
            db.Execute "UPDATE OpeningStock SET TransactionID = " & tid & " WHERE OpeningStockID = " & _
                       rs!OpeningStockID, dbFailOnError
            Posted = Posted + 1
        Else
            Skipped = Skipped + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    ws.CommitTrans
    inTrans = False
    LogAction "OPENING_STOCK_POST", "OpeningStock", refNo, "Posted=" & Posted & " Skipped=" & Skipped
    PostOpeningStock = SyncJournal()
    Exit Function
EH:
    PostOpeningStock = "تعذر ترحيل الرصيد: " & Err.Description & " (" & Err.Number & ")"
    Posted = 0
    If inTrans Then ws.Rollback
End Function

Public Function UnpostOpeningStock(ByVal LineID As Long) As String
    ' Takes back the move of a posted line while it is the only move of its product; the line stays to edit.
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean, tid As Variant, pid As Long
    Dim before As Collection, wh As Long, qty As Currency, remainQty As Currency
    UnpostOpeningStock = OpeningStockReady("DELETE")
    If Len(UnpostOpeningStock) > 0 Then Exit Function
    tid = DbValue("SELECT TransactionID FROM OpeningStock WHERE OpeningStockID = " & LineID)
    If IsNull(tid) Then Exit Function
    pid = DbValue("SELECT ProductID FROM OpeningStock WHERE OpeningStockID = " & LineID)
    wh = Nz(DbValue("SELECT WarehouseID FROM InventoryTransactions WHERE TransactionID = " & tid), MAIN_WAREHOUSE_ID)
    qty = Nz(DbValue("SELECT Quantity FROM InventoryTransactions WHERE TransactionID = " & tid), 0)
    If Nz(DbValue("SELECT COUNT(*) FROM InventoryTransactions WHERE ProductID = " & pid & " AND WarehouseID = " & wh & _
                  " AND TransactionID <> " & tid), 0) > 0 Then
        UnpostOpeningStock = "للصنف حركات بعد رصيده الافتتاحي، فلا يُلغى ترحيله. صحّح الكمية بالجرد."
        Exit Function
    End If
    On Error GoTo EH
    Set db = CurrentDb
    Set before = AuditSnapshot("InventoryTransactions", "TransactionID", tid)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "DELETE FROM InventoryTransactions WHERE TransactionID = " & tid, dbFailOnError
    ' the quantity of the other warehouses stays; with nothing left the average cost starts again
    remainQty = Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) - qty
    If remainQty = 0 Then
        db.Execute "UPDATE Products SET CurrentQuantity = 0, AverageCost = 0, UpdatedAt = Now() WHERE ProductID = " & pid, _
                   dbFailOnError
    Else
        db.Execute "UPDATE Products SET CurrentQuantity = " & Str$(remainQty) & ", UpdatedAt = Now() WHERE ProductID = " & pid, _
                   dbFailOnError
    End If
    db.Execute "UPDATE OpeningStock SET TransactionID = Null WHERE OpeningStockID = " & LineID, dbFailOnError
    ws.CommitTrans
    inTrans = False
    AuditDeleted "OPENING_STOCK_UNPOST", "InventoryTransactions", tid, before
    UnpostOpeningStock = SyncJournal()
    Exit Function
EH:
    UnpostOpeningStock = "تعذر إلغاء الترحيل: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
End Function

Public Function ImportOpeningStockRow(ByVal Code As String, ByVal Qty As Variant, ByVal UnitCost As Variant, _
                                      Optional ByVal WarehouseCode As String = "") As String
    ' One row of the Excel file: "ADDED", "UPDATED", or the reason it was not taken. WarehouseCode: the code
    ' or name of the warehouse (empty = the main warehouse, docs/57).
    Dim pid As Variant, lineID As Variant, cost As Currency, wh As Variant
    Code = Trim$(Code)
    If Len(Code) = 0 Then
        ImportOpeningStockRow = "EMPTY"
        Exit Function
    End If
    If Not IsNumeric(Qty) Then
        ImportOpeningStockRow = "BAD"
        Exit Function
    End If
    If CDbl(Qty) < 0 Then
        ImportOpeningStockRow = "BAD"
        Exit Function
    End If
    pid = DbValue("SELECT ProductID FROM Products WHERE ProductCode = " & SqlText(Code) & " OR Barcode = " & SqlText(Code))
    If IsNull(pid) Then
        ImportOpeningStockRow = "UNKNOWN"
        Exit Function
    End If
    If IsNumeric(UnitCost) Then
        cost = CCur(UnitCost)
    Else
        cost = Nz(DbValue("SELECT IIf(AverageCost > 0, AverageCost, PurchasePrice) FROM Products WHERE ProductID = " & pid), 0)
    End If
    If cost < 0 Then
        ImportOpeningStockRow = "BAD"
        Exit Function
    End If
    If Len(Trim$(WarehouseCode)) = 0 Then
        wh = MAIN_WAREHOUSE_ID
    Else
        wh = DbValue("SELECT WarehouseID FROM Warehouses WHERE WarehouseCode = " & SqlText(Trim$(WarehouseCode)) & _
                     " OR WarehouseName = " & SqlText(Trim$(WarehouseCode)))
        If IsNull(wh) Then
            ImportOpeningStockRow = "UNKNOWN"
            Exit Function
        End If
    End If
    lineID = DbValue("SELECT OpeningStockID FROM OpeningStock WHERE ProductID = " & pid & " AND WarehouseID = " & wh)
    If Not IsNull(lineID) Then
        If Not IsNull(DbValue("SELECT TransactionID FROM OpeningStock WHERE OpeningStockID = " & lineID)) Then
            ImportOpeningStockRow = "POSTED"
            Exit Function
        End If
    End If
    If Len(OpeningStockLineProblem(lineID, pid, Qty, cost, wh)) > 0 Then
        ImportOpeningStockRow = "MOVES"
        Exit Function
    End If
    If IsNull(lineID) Then
        CurrentDb.Execute "INSERT INTO OpeningStock (ProductID, WarehouseID, Quantity, UnitCost) VALUES (" & pid & ", " & _
                          wh & ", " & Str$(CDbl(Qty)) & ", " & Str$(cost) & ")", dbFailOnError
        ImportOpeningStockRow = "ADDED"
    Else
        CurrentDb.Execute "UPDATE OpeningStock SET Quantity = " & Str$(CDbl(Qty)) & ", UnitCost = " & Str$(cost) & _
                          " WHERE OpeningStockID = " & lineID, dbFailOnError
        ImportOpeningStockRow = "UPDATED"
    End If
End Function

Public Function ImportOpeningStock(ByVal FilePath As String) As String
    ' The Excel file: first row = titles; columns code (or barcode), quantity, cost, or the template's
    ' code, name, quantity, cost. Returns the summary.
    Dim db As DAO.Database, rs As DAO.Recordset, qCol As Integer, cCol As Integer, result As String
    Dim added As Long, updated As Long, unknown As Long, refused As Long, codes As String, kind As Integer
    Set db = CurrentDb
    DropImportTables db
    If LCase$(Right$(FilePath, 4)) = ".xls" Then
        kind = acSpreadsheetTypeExcel9
    Else
        kind = acSpreadsheetTypeExcel12Xml
    End If
    On Error GoTo EH
    DoCmd.TransferSpreadsheet acImport, kind, "tmpOpeningImport", FilePath, True
    Set rs = db.OpenRecordset("SELECT * FROM tmpOpeningImport", dbOpenSnapshot)
    If rs.Fields.Count >= 4 Then
        qCol = 2
        cCol = 3
    Else
        qCol = 1
        cCol = 2
    End If
    If rs.Fields.Count < 2 Then
        rs.Close
        DropImportTables db
        ImportOpeningStock = "الملف يحتاج عمودين على الأقل: الكود والكمية، ثم التكلفة."
        Exit Function
    End If
    Do Until rs.EOF
        If rs.Fields.Count > 4 And qCol = 2 Then          ' the template: code, name, quantity, cost, warehouse
            result = ImportOpeningStockRow(CStr(Nz(rs.Fields(0).Value, "")), rs.Fields(qCol).Value, rs.Fields(cCol).Value, _
                                           CStr(Nz(rs.Fields(4).Value, "")))
        ElseIf rs.Fields.Count > cCol Then
            result = ImportOpeningStockRow(CStr(Nz(rs.Fields(0).Value, "")), rs.Fields(qCol).Value, rs.Fields(cCol).Value)
        Else
            result = ImportOpeningStockRow(CStr(Nz(rs.Fields(0).Value, "")), rs.Fields(qCol).Value, Null)
        End If
        Select Case result
            Case "ADDED": added = added + 1
            Case "UPDATED": updated = updated + 1
            Case "UNKNOWN"
                unknown = unknown + 1
                If unknown <= 10 Then codes = codes & "  " & CStr(rs.Fields(0).Value)
            Case "EMPTY"
            Case Else: refused = refused + 1
        End Select
        rs.MoveNext
    Loop
    rs.Close
    DropImportTables db
    LogAction "OPENING_STOCK_IMPORT", "OpeningStock", "", "Added=" & added & " Updated=" & updated & _
              " Unknown=" & unknown & " Refused=" & refused & " File=" & FilePath
    ImportOpeningStock = "أُضيف " & added & " صنفًا، وعُدّل " & updated & "."
    If refused > 0 Then
        ImportOpeningStock = ImportOpeningStock & vbCrLf & "لم يُؤخذ " & refused & _
                             " سطرًا: كمية غير صحيحة، أو صنف له حركات، أو رصيده مرحّل."
    End If
    If unknown > 0 Then
        ImportOpeningStock = ImportOpeningStock & vbCrLf & unknown & " كودًا غير موجود:" & codes
    End If
    Exit Function
EH:
    ImportOpeningStock = "تعذر قراءة الملف: " & Err.Description & " (" & Err.Number & ")"
    DropImportTables db
End Function

Private Sub DropImportTables(ByVal db As DAO.Database)
    ' The imported sheet and the error tables Access adds next to it.
    Dim i As Long
    On Error Resume Next
    For i = db.TableDefs.Count - 1 To 0 Step -1
        If db.TableDefs(i).Name = "tmpOpeningImport" Or db.TableDefs(i).Name Like "*_ImportErrors*" Then
            db.TableDefs.Delete db.TableDefs(i).Name
        End If
    Next i
End Sub

Public Function OpeningStockTemplateSql() As String
    ' The products still waiting for their opening stock, in the columns the import reads.
    OpeningStockTemplateSql = "SELECT p.ProductCode AS [كود الصنف], p.ProductName AS [اسم الصنف], " & _
        "Nz(o.Quantity, 0) AS [الكمية الافتتاحية], " & _
        "Nz(o.UnitCost, IIf(p.AverageCost > 0, p.AverageCost, p.PurchasePrice)) AS [تكلفة الوحدة الافتتاحية], " & _
        "w.WarehouseCode AS [كود المخزن] " & _
        "FROM (Products AS p LEFT JOIN (SELECT * FROM OpeningStock WHERE WarehouseID = 1) AS o " & _
        "ON p.ProductID = o.ProductID), Warehouses AS w " & _
        "WHERE w.WarehouseID = 1 AND p.TrackStock = True AND p.IsActive = True AND o.TransactionID Is Null " & _
        "AND p.ProductID NOT IN (SELECT ProductID FROM InventoryTransactions WHERE WarehouseID = 1) " & _
        "ORDER BY p.ProductName"
End Function

'------------------------------------------------------------------------------
' The stock part of the screen
'------------------------------------------------------------------------------
Public Sub OpeningStockFill(ByVal frm As Access.Form)
    Dim msg As String, n As Long
    msg = OpeningStockReady("ADD")
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    n = FillOpeningStock()
    LogAction "OPENING_STOCK_FILL", "OpeningStock", "", "Added=" & n
    frm!subGrid.Form.Requery
    ShowInfo "أُضيف " & n & " صنفًا بكمية صفر. اكتب الكميات، ثم «ترحيل الرصيد»."
End Sub

Public Sub OpeningStockPost(ByVal frm As Access.Form)
    Dim msg As String, posted As Long, skipped As Long
    If frm!subGrid.Form.Dirty Then frm!subGrid.Form.Dirty = False
    If Not AskYesNo("ترحيل الرصيد الافتتاحي للأصناف بتاريخ " & GDate(OpeningDateValue()) & "؟" & vbCrLf & _
                    "بعد الترحيل لا يُعدَّل الرصيد إلا بإلغاء ترحيله.") Then Exit Sub
    msg = PostOpeningStock(posted, skipped)
    frm!subGrid.Form.Requery
    If Len(msg) > 0 Then
        ShowWarning msg
        If posted = 0 Then Exit Sub
    End If
    ShowInfo "رُحِّل رصيد " & posted & " صنفًا." & IIf(skipped > 0, vbCrLf & "لم يُرحَّل " & skipped & _
             " صنفًا لأن لها حركات أو لا تتابع المخزون.", "")
    OpeningRefresh frm
End Sub

Public Sub OpeningStockImport(ByVal frm As Access.Form)
    Dim msg As String, path As String
    msg = OpeningStockReady("ADD")
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    With Application.FileDialog(3)          ' msoFileDialogFilePicker
        .Title = Tr("اختر ملف Excel")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xls"
        If .Show Then path = .SelectedItems(1)
    End With
    If Len(path) = 0 Then Exit Sub
    msg = ImportOpeningStock(path)
    frm!subGrid.Form.Requery
    ShowInfo msg
End Sub

Public Sub OpeningStockTemplate(ByVal frm As Access.Form)
    Dim db As DAO.Database, path As String
    On Error GoTo EH
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete "qryOpeningStockExport"
    On Error GoTo EH
    db.CreateQueryDef "qryOpeningStockExport", Tr(OpeningStockTemplateSql())
    path = ReportsFolder() & "\OpeningStock_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    DoCmd.OutputTo acOutputQuery, "qryOpeningStockExport", acFormatXLSX, path
    db.QueryDefs.Delete "qryOpeningStockExport"
    ShowInfo "تم حفظ القالب في:" & vbCrLf & path & vbCrLf & "اكتب الكميات والتكلفة، ثم «استيراد من Excel»."
    If Not g_SilentMode Then Application.FollowHyperlink path
    Exit Sub
EH:
    ShowError "تعذر حفظ القالب: " & Err.Description
End Sub

Public Sub OpeningStockRemoveLine(ByVal frm As Access.Form)
    ' A line not posted is deleted; a posted line is unposted (it stays, to be corrected and posted again).
    Dim id As Variant, msg As String, before As Collection
    If frm.NewRecord Then
        If frm.Dirty Then frm.Undo
        Exit Sub
    End If
    id = frm!OpeningStockID.Value
    If IsNull(id) Then Exit Sub
    If frm.Dirty Then frm.Undo
    If Not IsNull(frm!TransactionID.Value) Then
        If Not AskYesNo("إلغاء ترحيل رصيد هذا الصنف؟ تُحذف حركته ويعود رصيده صفرًا.") Then Exit Sub
        msg = UnpostOpeningStock(CLng(id))
    Else
        msg = OpeningStockReady("DELETE")
        If Len(msg) = 0 Then
            If Not AskYesNo("حذف هذا السطر؟") Then Exit Sub
            Set before = AuditSnapshot("OpeningStock", "OpeningStockID", id)
            CurrentDb.Execute "DELETE FROM OpeningStock WHERE OpeningStockID = " & CLng(id), dbFailOnError
            AuditDeleted "OPENING_STOCK_DELETE", "OpeningStock", id, before
        End If
    End If
    If Len(msg) > 0 Then ShowWarning msg
    frm.Requery
    OpeningGridAfterUpdate frm
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Public Function TestOpeningBalances() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String
    Dim first As Variant, startDay As Date, totalBefore As Currency, entry As Variant, msg As String
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestOpeningBalances  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record OpeningAccountProblem(1300) <> "" And OpeningAccountProblem(OPENING_ACCOUNT) <> "" And _
           OpeningAccountProblem(110001) <> "" And OpeningAccountProblem(3200) = "", _
           "الحسابات المسموحة في الأرصدة العامة", passed, failed, report
    Record OpeningDateProblem(Null) <> "", "التاريخ مطلوب", passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    msg = SyncJournal()
    first = FirstDocumentDate()
    If IsDate(first) Then
        Record OpeningDateProblem(DateAdd("d", 1, DateValue(first))) <> "", "لا يكون بعد أول مستند", passed, failed, report
        startDay = DateAdd("d", -1, DateValue(first))
    Else
        startDay = Date
    End If
    If Len(ClosedPeriodProblem(startDay)) > 0 Then
        Debug.Print "[--] يوم البدء في فترة مقفلة: تخطي القيد"
    Else
        Record SaveOpeningDate(startDay) = "" And OpeningDateValue() = startDay, "حفظ تاريخ بدء التشغيل", _
               passed, failed, report
        totalBefore = OpeningBalanceTotal()
        CurrentDb.Execute "INSERT INTO OpeningBalances (AccountCode, Debit, Credit, LineText) VALUES (3200, 0, 1000, 'TEST')", _
                          dbFailOnError
        msg = SyncJournal()
        entry = DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'ACCOUNT_OPENING' AND SourceID = " & _
                        "(SELECT Max(OpeningLineID) FROM OpeningBalances)")
        Record Not IsNull(entry), "قيد الرصيد الافتتاحي للحساب", passed, failed, report
        If Not IsNull(entry) Then
            Record DateValue(DbValue("SELECT EntryDate FROM JournalEntries WHERE EntryID = " & entry)) = startDay And _
                   Nz(DbValue("SELECT Debit FROM JournalLines WHERE EntryID = " & entry & " AND AccountCode = 3900"), 0) = 1000, _
                   "بتاريخ بدء التشغيل ومقابله حساب إجمالي الأرصدة الافتتاحية", passed, failed, report
        End If
        Record OpeningBalanceTotal() = totalBefore + 1000, "إجمالي الأرصدة الافتتاحية يتغير بالرصيد", _
               passed, failed, report
        Record IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType IN ('CUSTOMER_OPENING', " & _
                              "'SUPPLIER_OPENING') AND DateValue(EntryDate) <> " & SqlDate(startDay))), _
               "أرصدة العملاء والموردين بتاريخ بدء التشغيل", passed, failed, report
        TestOpeningStockPart startDay, first, passed, failed, report
    End If
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات أرصدة بداية التشغيل ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, _
                "TestOpeningBalances"
        TestOpeningBalances = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestOpeningBalances"
    End If
End Function

Private Sub TestOpeningStockPart(ByVal startDay As Date, ByVal first As Variant, ByRef passed As Long, _
                                 ByRef failed As Long, ByRef report As String)
    ' The opening stock of a product without moves: post, entry, unpost, import row.
    Dim pid As Variant, lineID As Variant, tid As Variant, posted As Long, skipped As Long, code As String
    pid = DbValue("SELECT TOP 1 ProductID FROM Products WHERE TrackStock = True AND ProductID NOT IN " & _
                  "(SELECT ProductID FROM InventoryTransactions) AND ProductID NOT IN (SELECT ProductID FROM OpeningStock) " & _
                  "ORDER BY ProductID")
    If IsNull(pid) Then
        Debug.Print "[--] لا يوجد صنف بلا حركات: تخطي اختبار رصيد المخزون"
        Exit Sub
    End If
    code = Nz(DbValue("SELECT ProductCode FROM Products WHERE ProductID = " & pid), "")
    Record ImportOpeningStockRow("ZZ-NO-SUCH-CODE-9", 1, 1) = "UNKNOWN", "استيراد: كود غير موجود", passed, failed, report
    Record ImportOpeningStockRow(code, -1, 1) = "BAD", "استيراد: كمية سالبة مرفوضة", passed, failed, report
    Record ImportOpeningStockRow(code, 4, 3) = "ADDED", "استيراد: إضافة سطر للصنف", passed, failed, report
    Record ImportOpeningStockRow(code, 10, 5) = "UPDATED", "استيراد: تعديل سطر الصنف", passed, failed, report
    lineID = DbValue("SELECT OpeningStockID FROM OpeningStock WHERE ProductID = " & pid)
    CurrentDb.Execute "UPDATE OpeningStock SET Quantity = 0 WHERE TransactionID Is Null AND OpeningStockID <> " & lineID, _
                      dbFailOnError                                   ' only the test line is posted
    Record PostOpeningStock(posted, skipped) = "" And posted = 1, "ترحيل الرصيد الافتتاحي للمخزون", _
           passed, failed, report
    tid = DbValue("SELECT TransactionID FROM OpeningStock WHERE OpeningStockID = " & lineID)
    Record Not IsNull(tid), "السطر مربوط بحركة المخزون", passed, failed, report
    If IsNull(tid) Then Exit Sub
    Record DateValue(DbValue("SELECT TransactionDate FROM InventoryTransactions WHERE TransactionID = " & tid)) = startDay _
           And Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), 0) = 10, _
           "الحركة بتاريخ بدء التشغيل والكمية 10", passed, failed, report
    Record Nz(DbValue("SELECT l.Credit FROM JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID " & _
                      "WHERE e.SourceType = 'STOCK_MOVE' AND e.SourceID = " & tid & " AND l.AccountCode = 3900"), 0) = 50, _
           "قيد المخزون مقابل إجمالي الأرصدة الافتتاحية (50)", passed, failed, report
    Record Nz(FirstDocumentDate(), 0) = Nz(first, 0), "رصيد المخزون لا يُعد أول مستند", passed, failed, report
    Record ImportOpeningStockRow(code, 1, 1) = "POSTED", "استيراد: الرصيد المرحّل لا يتغير", passed, failed, report
    Record UnpostOpeningStock(CLng(lineID)) = "" And _
           IsNull(DbValue("SELECT TransactionID FROM OpeningStock WHERE OpeningStockID = " & lineID)) And _
           Nz(DbValue("SELECT CurrentQuantity FROM Products WHERE ProductID = " & pid), -1) = 0, _
           "إلغاء الترحيل يحذف الحركة ويعيد الكمية صفرًا", passed, failed, report
End Sub

Private Sub Record(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
