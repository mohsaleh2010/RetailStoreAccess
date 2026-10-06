Attribute VB_Name = "modJournal"
'==============================================================================
' modJournal  -  Retail Store Management System (journal entries)
'
' Every operation gets ONE automatic journal entry, linked to its origin by
' JournalEntries.SourceType + SourceID. The debit / credit lines come from the
' saved queries qryJournal* (one per kind of operation: the accounting rules).
'
'   EnsureAccounts   sub-accounts of the cash boxes (110000 + box) and of the
'                    expense types (530000 + type)
'   SyncJournal      creates the entries of new operations, rebuilds the entries of
'                    changed ones (same number), deletes the entries of deleted ones;
'                    new entries are numbered (JV-000001) in date order
'   OpenJournalSource opens the origin of an entry (invoice, voucher, expense...)
' Screens: frmJournal (review by period), frmJournalEntry (one entry).
'==============================================================================
Option Compare Database
Option Explicit

Private Const SOURCE_QUERIES As String = "qryJournalSale,qryJournalSalesReturn,qryJournalPurchase," & _
    "qryJournalPurchaseReturn,qryJournalPayments,qryJournalExpense,qryJournalCashVoucher,qryJournalStock," & _
    "qryJournalOpening"

'==============================================================================
' Accounts and synchronisation
'==============================================================================
Public Sub EnsureAccounts()
    CurrentDb.Execute "INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode) " & _
        "SELECT 110000 + b.CashBoxID, b.BoxName, 'ASSET', 1100 FROM CashBoxes AS b " & _
        "WHERE 110000 + b.CashBoxID NOT IN (SELECT AccountCode FROM Accounts)", dbFailOnError
    CurrentDb.Execute "INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode) " & _
        "SELECT 530000 + t.ExpenseTypeID, t.ExpenseTypeName, 'EXPENSE', 5300 FROM ExpenseTypes AS t " & _
        "WHERE 530000 + t.ExpenseTypeID NOT IN (SELECT AccountCode FROM Accounts)", dbFailOnError
End Sub

Public Function SyncJournal(Optional ByRef Added As Long, Optional ByRef Updated As Long, _
                            Optional ByRef Removed As Long) As String
    ' "" on success, else an Arabic message. Runs in one transaction.
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean
    Dim rs As DAO.Recordset, e As DAO.Recordset, existing As Object, seen As Object, kinds As Object
    Dim sources() As String, i As Long, key As String, info As Variant, k As Variant, skipped As String

    On Error GoTo EH
    Added = 0: Updated = 0: Removed = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    EnsureAccounts
    Set kinds = CreateObject("Scripting.Dictionary")
    Set rs = db.OpenRecordset("SELECT SourceType, TypeName FROM JournalSourceTypes", dbOpenSnapshot)
    Do Until rs.EOF
        kinds(CStr(rs!SourceType)) = CStr(rs!TypeName)
        rs.MoveNext
    Loop
    rs.Close
    Set existing = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set rs = db.OpenRecordset("SELECT EntryID, SourceType, SourceID, Signature, EntryDate, TotalDebit " & _
                              "FROM JournalEntries", dbOpenSnapshot)
    Do Until rs.EOF
        existing(rs!SourceType & "|" & rs!SourceID) = Array(rs!EntryID, rs!Signature, rs!EntryDate, rs!TotalDebit)
        rs.MoveNext
    Loop
    rs.Close

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    Set e = db.OpenRecordset("JournalEntries", dbOpenDynaset)
    sources = Split(SOURCE_QUERIES, ",")
    For i = 0 To UBound(sources)
        Set rs = db.OpenRecordset("SELECT SourceType, SourceID, Max(SourceNumber) AS DocNumber, " & _
            "Max(SourceDate) AS DocDate, Sum(Debit) AS SumDebit, Sum(Credit) AS SumCredit, " & _
            "Sum(AccountCode * (Debit + Debit + Credit)) AS Sig, Count(*) AS LineTotal, " & _
            "Max(Party) AS FirstText FROM " & sources(i) & _
            " GROUP BY SourceType, SourceID", dbOpenSnapshot)
        Do Until rs.EOF
            key = rs!SourceType & "|" & rs!SourceID
            seen(key) = True
            If CCur(Nz(rs!SumDebit, 0)) <> CCur(Nz(rs!SumCredit, 0)) Then
                skipped = skipped & "  " & kinds(CStr(rs!SourceType)) & " " & Nz(rs!DocNumber, "") & vbCrLf
            ElseIf Not existing.Exists(key) Then
                e.AddNew
                e!EntryNumber = "~" & Format$(Added + 1, "000000")      ' numbered in date order below
                FillHeader e, rs, kinds
                e.Update
                Added = Added + 1
            Else
                info = existing(key)
                If info(1) <> CCur(Nz(rs!Sig, 0)) Or info(2) <> SourceDateOf(rs) Or info(3) <> CCur(rs!SumDebit) Then
                    db.Execute "DELETE FROM JournalLines WHERE EntryID = " & info(0), dbFailOnError
                    e.FindFirst "EntryID = " & info(0)
                    e.Edit
                    FillHeader e, rs, kinds
                    e!UpdatedAt = Now
                    e.Update
                    Updated = Updated + 1
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        ' lines of the new and changed entries (LineCount < 0 = waiting for its lines)
        db.Execute "INSERT INTO JournalLines (EntryID, LineNumber, AccountCode, Debit, Credit, LineText) " & _
            "SELECT e.EntryID, q.LineOrder, q.AccountCode, q.Debit, q.Credit, q.LineText FROM " & sources(i) & _
            " AS q INNER JOIN JournalEntries AS e ON (q.SourceType = e.SourceType AND q.SourceID = e.SourceID) " & _
            "WHERE e.LineCount < 0", dbFailOnError
    Next
    e.Close
    db.Execute "UPDATE JournalEntries SET LineCount = -LineCount WHERE LineCount < 0", dbFailOnError

    ' operations that no longer exist (a deleted expense, ...)
    For Each k In existing.Keys
        If Not seen.Exists(k) Then
            info = existing(k)
            db.Execute "DELETE FROM JournalLines WHERE EntryID = " & info(0), dbFailOnError
            db.Execute "DELETE FROM JournalEntries WHERE EntryID = " & info(0), dbFailOnError
            Removed = Removed + 1
        End If
    Next

    ' numbers of the new entries, in date order
    Set rs = db.OpenRecordset("SELECT EntryNumber FROM JournalEntries WHERE EntryNumber Like '~*' " & _
                              "ORDER BY EntryDate, EntryID", dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit
        rs!EntryNumber = NextNumber("JOURNAL")
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    ws.CommitTrans
    inTrans = False
    If Added + Updated + Removed > 0 Then
        LogAction "JOURNAL_SYNC", "JournalEntries", "", "Added=" & Added & " Updated=" & Updated & " Removed=" & Removed
    End If
    If Len(skipped) > 0 Then
        SyncJournal = "⁄„·Ì«  €Ì— „ Ê«“‰… ·„ Ìı‰‘√ ·Â« ﬁÌœ (—«Ã⁄Â«):" & vbCrLf & Left$(skipped, 700)
    End If
    Exit Function

EH:
    SyncJournal = " ⁄–—  ÕœÌÀ «·ﬁÌÊœ: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
    Added = 0: Updated = 0: Removed = 0
End Function

Private Function SourceDateOf(ByVal rs As DAO.Recordset) As Date
    If IsNull(rs!DocDate) Then SourceDateOf = DateSerial(2000, 1, 1) Else SourceDateOf = rs!DocDate
End Function

Private Sub FillHeader(ByVal e As DAO.Recordset, ByVal rs As DAO.Recordset, ByVal kinds As Object)
    Dim docNo As String, text As String
    docNo = Left$(Nz(rs!DocNumber, "-"), 20)
    text = kinds(CStr(rs!SourceType)) & " " & docNo
    If Len(Nz(rs!FirstText, "")) > 0 Then text = text & " - " & rs!FirstText
    e!EntryDate = SourceDateOf(rs)
    e!SourceType = rs!SourceType
    e!SourceID = rs!SourceID
    e!SourceNumber = docNo
    e!Description = Left$(text, 255)
    e!TotalDebit = CCur(rs!SumDebit)
    e!TotalCredit = CCur(rs!SumCredit)
    e!Signature = CCur(Nz(rs!Sig, 0))
    e!LineCount = -rs!LineTotal
End Sub

'==============================================================================
' Opening the origin of an entry
'==============================================================================
Public Sub OpenJournalSource(ByVal EntryID As Variant)
    Dim st As String, id As Long
    If IsNull(EntryID) Then
        ShowWarning "«Œ — ﬁÌœ« √Ê·«."
        Exit Sub
    End If
    st = Nz(DbValue("SELECT SourceType FROM JournalEntries WHERE EntryID = " & CLng(EntryID)), "")
    id = Nz(DbValue("SELECT SourceID FROM JournalEntries WHERE EntryID = " & CLng(EntryID)), 0)
    Select Case st
        Case "SALE":             OpenScreen "frmSalesInvoice", 0, id
        Case "SALES_RETURN":     PrintSalesDocument "RETURN", id
        Case "PURCHASE":         OpenScreen "frmPurchaseView", 0, id
        Case "PURCHASE_RETURN":  PrintPurchaseDocument "RETURN", id
        Case "CUSTOMER_PAYMENT": PrintVoucher "RECEIPT", id
        Case "SUPPLIER_PAYMENT": PrintVoucher "PAYMENT", id
        Case "EXPENSE":          OpenScreen "frmExpenses", 0, id
        Case "CASH_VOUCHER":     PrintCashVoucher id
        Case "STOCK_MOVE":       OpenScreen "frmInventory"
        Case "STOCK_COUNT":      PrintStockCount id
        Case "BOX_OPENING":      OpenScreen "frmCashBoxes", 0, id
        Case "CUSTOMER_OPENING": OpenScreen "frmCustomers", 0, id
        Case "SUPPLIER_OPENING": OpenScreen "frmSuppliers", 0, id
        Case Else:               ShowWarning "«·ﬁÌœ €Ì— „ÊÃÊœ."
    End Select
End Sub

'==============================================================================
' frmJournal: entries of a period, the lines of the selected entry
'==============================================================================
Public Sub JournalLoad(ByVal frm As Access.Form)
    Dim rows As String, rs As DAO.Recordset
    Calendar = vbCalGreg
    rows = """ALL"";""ﬂ· «·⁄„·Ì« """
    Set rs = CurrentDb.OpenRecordset("SELECT SourceType, TypeName FROM JournalSourceTypes ORDER BY SortOrder", dbOpenSnapshot)
    Do Until rs.EOF
        rows = rows & ";""" & rs!SourceType & """;""" & rs!TypeName & """"
        rs.MoveNext
    Loop
    rs.Close
    frm!cboSourceType.RowSource = rows
    frm!cboSourceType.Value = "ALL"
    frm!txtFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    frm!txtTo.Value = Date
    JournalSync frm
End Sub

Public Sub JournalSync(ByVal frm As Access.Form)
    Dim msg As String, added As Long, updated As Long, removed As Long
    DoCmd.Hourglass True
    msg = SyncJournal(added, updated, removed)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then ShowWarning msg
    frm!lblSync.Caption = " ÕœÌÀ «·ﬁÌÊœ „‰ «·⁄„·Ì« : " & added & " ﬁÌœ ÃœÌœ° Ê" & updated & " ﬁÌœ „ıÕœÛ¯À° Ê" & _
                          removed & " ﬁÌœ „Õ–Ê›  (" & GDate(Now, True) & ")"
    JournalRefresh frm
End Sub

Public Sub JournalQuickPeriod(ByVal frm As Access.Form, ByVal Which As String)
    SetQuickPeriod frm, Which
    JournalRefresh frm
End Sub

Private Function JournalWhere(ByVal frm As Access.Form) As String
    ' Filter of the shown entries ("" when the dates are not valid).
    Dim w As String, s As String
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then Exit Function
    w = "e.EntryDate >= " & SqlDate(DateValue(frm!txtFrom.Value)) & " AND e.EntryDate < " & _
        SqlDate(DateAdd("d", 1, DateValue(frm!txtTo.Value)))
    If Nz(frm!cboSourceType.Value, "ALL") <> "ALL" Then w = w & " AND e.SourceType = " & SqlText(frm!cboSourceType.Value)
    s = Trim$(Nz(frm!txtSearch.Value, ""))
    If Len(s) > 0 Then
        w = w & " AND (e.EntryNumber Like " & LikePattern(s) & " OR e.SourceNumber Like " & LikePattern(s) & _
            " OR e.Description Like " & LikePattern(s) & ")"
    End If
    JournalWhere = w
End Function

Public Sub JournalRefresh(ByVal frm As Access.Form)
    Dim w As String, rs As DAO.Recordset
    w = JournalWhere(frm)
    If Len(w) = 0 Then
        ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
        Exit Sub
    End If
    frm!lstEntries.RowSource = "SELECT e.EntryID, e.EntryNumber AS [—ﬁ„ «·ﬁÌœ], GDate(e.EntryDate) AS [«· «—ÌŒ], " & _
        "t.TypeName AS [«·⁄„·Ì…], e.SourceNumber AS [«·„” ‰œ], e.Description AS [«·»Ì«‰], " & _
        "Format(e.TotalDebit, '#,##0.00') AS [«·„»·€] FROM JournalEntries AS e INNER JOIN JournalSourceTypes AS t " & _
        "ON e.SourceType = t.SourceType WHERE " & w & " ORDER BY e.EntryDate, e.EntryNumber"
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS N, Sum(e.TotalDebit) AS D, Sum(e.TotalCredit) AS C " & _
                                     "FROM JournalEntries AS e WHERE " & w, dbOpenSnapshot)
    frm!lblTotals.Caption = "⁄œœ «·ﬁÌÊœ: " & rs!N & "    ≈Ã„«·Ì «·„œÌ‰: " & Format$(Nz(rs!D, 0), "#,##0.00") & _
                            "    ≈Ã„«·Ì «·œ«∆‰: " & Format$(Nz(rs!C, 0), "#,##0.00") & _
                            IIf(Nz(rs!D, 0) = Nz(rs!C, 0), "    („ Ê«“‰)", "    (€Ì— „ Ê«“‰!)")
    rs.Close
    If frm!lstEntries.ListCount > 1 Then frm!lstEntries.Value = frm!lstEntries.ItemData(1) Else frm!lstEntries.Value = Null
    JournalEntryPicked frm
End Sub

Public Sub JournalEntryPicked(ByVal frm As Access.Form)
    If IsNull(frm!lstEntries.Value) Then
        frm!lstLines.RowSource = ""
        Exit Sub
    End If
    frm!lstLines.RowSource = EntryLinesSql(frm!lstEntries.Value)
End Sub

Private Function EntryLinesSql(ByVal EntryID As Variant) As String
    EntryLinesSql = "SELECT l.AccountCode AS [«·Õ”«»], a.AccountName AS [«”„ «·Õ”«»], l.LineText AS [«·»Ì«‰], " & _
        "IIf(l.Debit = 0, Null, Format(l.Debit, '#,##0.00')) AS [„œÌ‰], " & _
        "IIf(l.Credit = 0, Null, Format(l.Credit, '#,##0.00')) AS [œ«∆‰] " & _
        "FROM JournalLines AS l INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode " & _
        "WHERE l.EntryID = " & CLng(EntryID) & " ORDER BY l.LineNumber"
End Function

Public Sub JournalOpenEntry(ByVal frm As Access.Form)
    If IsNull(frm!lstEntries.Value) Then
        ShowWarning "«Œ — ﬁÌœ« √Ê·«."
        Exit Sub
    End If
    OpenScreen "frmJournalEntry", 0, frm!lstEntries.Value
End Sub

Public Sub PrintJournal(ByVal frm As Access.Form, ByVal Which As String)
    ' Which: JOURNAL (the entries shown) or TRIAL (trial balance of the period)
    Dim fromDate As Date, toDate As Date, where As String
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
        Exit Sub
    End If
    fromDate = DateValue(frm!txtFrom.Value)
    toDate = DateValue(frm!txtTo.Value)
    SetPeriod fromDate, toDate
    If Which = "TRIAL" Then
        OpenReportOrQuery "rptTrialBalance", "TrialBalanceQuery", "", PeriodText(fromDate, toDate)
    Else
        If Nz(frm!cboSourceType.Value, "ALL") <> "ALL" Then where = "[SourceType] = " & SqlText(frm!cboSourceType.Value)
        OpenReportOrQuery "rptJournal", "JournalLinesQuery", where, PeriodText(fromDate, toDate) & _
                          IIf(Len(where) > 0, "    " & frm!cboSourceType.Column(1), "")
    End If
End Sub

'==============================================================================
' frmJournalEntry: OpenArgs = EntryID
'==============================================================================
Public Sub JournalEntryLoad(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset
    Calendar = vbCalGreg
    If IsNull(frm.OpenArgs) Then Exit Sub
    frm!txtEntryID.Value = CLng(frm.OpenArgs)
    Set rs = CurrentDb.OpenRecordset("SELECT e.*, t.TypeName FROM JournalEntries AS e INNER JOIN JournalSourceTypes AS t " & _
                                     "ON e.SourceType = t.SourceType WHERE e.EntryID = " & CLng(frm.OpenArgs), dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        frm!lblHeader.Caption = "«·ﬁÌœ €Ì— „ÊÃÊœ (—»„« Õı–›  «·⁄„·Ì…). Õœ¯À «·ﬁÌÊœ."
        Exit Sub
    End If
    frm!lblTitle.Caption = "ﬁÌœ ÌÊ„Ì… —ﬁ„ " & rs!EntryNumber
    frm!lblHeader.Caption = "«· «—ÌŒ: " & GDate(rs!EntryDate, True) & "    «·⁄„·Ì…: " & rs!TypeName & " " & _
                            Nz(rs!SourceNumber, "") & IIf(IsNull(rs!UpdatedAt), "", "    (Õıœˆ¯À ›Ì " & GDate(rs!UpdatedAt, True) & ")")
    frm!lblDescription.Caption = "«·»Ì«‰: " & Nz(rs!Description, "-")
    frm!lblTotals.Caption = "«·≈Ã„«·Ì: „œÌ‰ " & Format$(rs!TotalDebit, "#,##0.00") & " = œ«∆‰ " & _
                            Format$(rs!TotalCredit, "#,##0.00")
    rs.Close
    frm!lstLines.RowSource = EntryLinesSql(frm.OpenArgs)
End Sub

Public Sub PrintJournalEntry(ByVal EntryID As Variant)
    If IsNull(EntryID) Then Exit Sub
    If Not ReportExists("rptJournalEntry") Then
        ShowWarning " ﬁ—Ì— «·ÿ»«⁄… €Ì— „ÊÃÊœ: rptJournalEntry" & vbCrLf & "‘€¯· BuildReports."
        Exit Sub
    End If
    On Error GoTo EH
    TempVars.Add "ReportCriteria", ""
    DoCmd.OpenReport "rptJournalEntry", POS_PRINT_VIEW, , "[EntryID] = " & CLng(EntryID)
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError " ⁄–— › Õ «· ﬁ—Ì—: " & Err.Description
End Sub

'==============================================================================
' In-Access test (RunAllTests)
'==============================================================================
Public Function TestJournal() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String
    Dim added As Long, updated As Long, removed As Long, ws As DAO.Workspace, inTrans As Boolean, id As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestJournal  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    msg = SyncJournal(added, updated, removed)
    CheckJournal Len(msg) = 0, " ÕœÌÀ «·ﬁÌÊœ „‰ ﬂ· «·⁄„·Ì«  " & msg, passed, failed, report
    msg = SyncJournal(added, updated, removed)
    CheckJournal Len(msg) = 0 And added + updated + removed = 0, "«· ÕœÌÀ «·À«‰Ì ·« Ì€Ì¯— ‘Ì∆«", passed, failed, report
    CheckJournal Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE TotalDebit <> TotalCredit"), 0) = 0, _
                 "ﬂ· «·ﬁÌÊœ „ Ê«“‰…", passed, failed, report
    CheckJournal Nz(DbValue("SELECT COUNT(*) FROM JournalEntries AS e WHERE e.TotalDebit <> " & _
                            "(SELECT Sum(Debit) FROM JournalLines AS l WHERE l.EntryID = e.EntryID)"), 0) = 0, _
                 "√”ÿ— ﬂ· ﬁÌœ  ”«ÊÌ —√”Â", passed, failed, report
    CheckJournal Nz(DbValue("SELECT COUNT(*) FROM SalesInvoices WHERE TotalAmount <> 0 AND SalesInvoiceID NOT IN " & _
                            "(SELECT SourceID FROM JournalEntries WHERE SourceType = 'SALE')"), 0) = 0, _
                 "·ﬂ· ›« Ê—… »Ì⁄ ﬁÌœ", passed, failed, report
    CheckJournal AccountBalance(1300) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Customers"), 0), _
                 "Õ”«» –„„ «·⁄„·«¡ = √—’œ… «·⁄„·«¡", passed, failed, report
    CheckJournal -AccountBalance(2100) = Nz(DbValue("SELECT Sum(CurrentBalance) FROM Suppliers"), 0), _
                 "Õ”«» –„„ «·„Ê—œÌ‰ = √—’œ… «·„Ê—œÌ‰", passed, failed, report

    ' a new, changed and deleted operation, rolled back
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO Expenses (ExpenseNumber, ExpenseDate, ExpenseTypeID, Amount, Tax, TotalAmount, " & _
        "PaymentMethodID, Description, EmployeeID) VALUES ('TEST-JV', Date(), 1, 100, 15, 115, 3, 'TEST-JV', " & _
        CurrentUserID() & ")", dbFailOnError
    msg = SyncJournal(added, updated, removed)
    id = Nz(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'EXPENSE' AND SourceNumber = 'TEST-JV'"), 0)
    CheckJournal added = 1 And id > 0, "„’—Ê› ÃœÌœ = ﬁÌœ ÃœÌœ", passed, failed, report
    CheckJournal Nz(DbValue("SELECT Debit FROM JournalLines WHERE EntryID = " & id & " AND AccountCode = 1500"), 0) = 15 And _
                 Nz(DbValue("SELECT Credit FROM JournalLines WHERE EntryID = " & id & " AND AccountCode = 1200"), 0) = 115, _
                 "ﬁÌœ «·„’—Ê›: «·÷—Ì»… „œÌ‰… Ê«·»‰ﬂ œ«∆‰", passed, failed, report
    CurrentDb.Execute "UPDATE Expenses SET Amount = 200, Tax = 0, TotalAmount = 200 WHERE ExpenseNumber = 'TEST-JV'", dbFailOnError
    msg = SyncJournal(added, updated, removed)
    CheckJournal updated = 1 And Nz(DbValue("SELECT TotalDebit FROM JournalEntries WHERE EntryID = " & id), 0) = 200, _
                 " ⁄œÌ· «·„’—Ê› ÌÕœ¯À ﬁÌœÂ »‰›” «·—ﬁ„", passed, failed, report
    CurrentDb.Execute "DELETE FROM Expenses WHERE ExpenseNumber = 'TEST-JV'", dbFailOnError
    msg = SyncJournal(added, updated, removed)
    CheckJournal removed = 1 And Nz(DbValue("SELECT COUNT(*) FROM JournalEntries WHERE EntryID = " & id), 0) = 0, _
                 "Õ–› «·„’—Ê› ÌÕ–› ﬁÌœÂ", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckJournal False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·ﬁÌÊœ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestJournal"
        TestJournal = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestJournal"
    End If
End Function

Public Function AccountBalance(ByVal AccountCode As Long) As Currency
    ' Debit minus credit of all the entries of an account.
    AccountBalance = Nz(DbValue("SELECT Sum(Debit) - Sum(Credit) FROM JournalLines WHERE AccountCode = " & AccountCode), 0)
End Function

Private Sub CheckJournal(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
