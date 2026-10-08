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
    "qryJournalOpening,qryJournalManual,qryJournalYearClose,qryJournalVatReturn,qryJournalBankTx,qryJournalCheque,qryJournalAsset,qryJournalDepreciation,qryJournalPayroll,qryJournalCommission"

'==============================================================================
' Accounts and synchronisation
'==============================================================================
Public Sub EnsureAccounts()
    ' a sub-account for each cash box, each bank and each expense type, then the levels of the tree (modAccounts)
    CurrentDb.Execute "INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode, IsPosting, IsSystem) " & _
        "SELECT 110000 + b.CashBoxID, b.BoxName, 'ASSET', 1100, True, True FROM CashBoxes AS b " & _
        "WHERE 110000 + b.CashBoxID NOT IN (SELECT AccountCode FROM Accounts)", dbFailOnError
    CurrentDb.Execute "INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode, IsPosting, IsSystem) " & _
        "SELECT 530000 + t.ExpenseTypeID, t.ExpenseTypeName, 'EXPENSE', 5300, True, True FROM ExpenseTypes AS t " & _
        "WHERE 530000 + t.ExpenseTypeID NOT IN (SELECT AccountCode FROM Accounts)", dbFailOnError
    CurrentDb.Execute "INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode, IsPosting, IsSystem) " & _
        "SELECT 120000 + k.BankID, k.BankName, 'ASSET', 1210, True, True FROM Banks AS k " & _
        "WHERE 120000 + k.BankID NOT IN (SELECT AccountCode FROM Accounts)", dbFailOnError
    RebuildAccountTree
End Sub

Public Function SyncJournal(Optional ByRef Added As Long, Optional ByRef Updated As Long, _
                            Optional ByRef Removed As Long) As String
    ' "" on success, else an Arabic message. Runs in one transaction.
    Dim db As DAO.Database, ws As DAO.Workspace, inTrans As Boolean
    Dim rs As DAO.Recordset, e As DAO.Recordset, existing As Object, seen As Object, kinds As Object
    Dim sources() As String, i As Long, key As String, info As Variant, k As Variant, skipped As String
    Dim closed As Date, locked As String, st As String

    On Error GoTo EH
    Added = 0: Updated = 0: Removed = 0
    Calendar = vbCalGreg
    Set db = CurrentDb
    closed = ClosedThroughDate()            ' entries until this day never change (modClosing); 0 = none
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
        ' .Value: Array(rs!EntryID) would keep the Field objects, invalid once rs is closed (error 3420)
        existing(rs!SourceType & "|" & rs!SourceID) = Array(rs!EntryID.Value, rs!Signature.Value, _
                                                            rs!EntryDate.Value, rs!TotalDebit.Value)
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
            "Sum(AccountCode * (Debit + Debit + Credit)) + Sum(CostCenter * (Debit + Debit + Credit) * 7) AS Sig, " & _
            "Count(*) AS LineTotal, " & _
            "Max(Party) AS FirstText FROM " & sources(i) & _
            " GROUP BY SourceType, SourceID", dbOpenSnapshot)
        Do Until rs.EOF
            key = rs!SourceType & "|" & rs!SourceID
            st = CStr(rs!SourceType)
            seen(key) = True
            If CCur(Nz(rs!SumDebit, 0)) <> CCur(Nz(rs!SumCredit, 0)) Then
                skipped = skipped & "  " & kinds(st) & " " & Nz(rs!DocNumber, "") & vbCrLf
            ElseIf Not existing.Exists(key) And InClosedPeriod(st, SourceDateOf(rs), closed) Then
                locked = locked & "  " & kinds(st) & " " & Nz(rs!DocNumber, "") & vbCrLf
            ElseIf Not existing.Exists(key) Then
                e.AddNew
                e!EntryNumber = "~" & Format$(Added + 1, "000000")      ' numbered in date order below
                FillHeader e, rs, kinds
                e.Update
                Added = Added + 1
            Else
                info = existing(key)
                If info(1) = CCur(Nz(rs!Sig, 0)) And info(2) = SourceDateOf(rs) And info(3) = CCur(rs!SumDebit) Then
                    ' unchanged
                ElseIf InClosedPeriod(st, info(2), closed) Or InClosedPeriod(st, SourceDateOf(rs), closed) Then
                    locked = locked & "  " & kinds(st) & " " & Nz(rs!DocNumber, "") & vbCrLf
                Else
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
        ' the cost centre: 0 in the source queries = not allocated (an entry without centre keeps its old signature)
        db.Execute "INSERT INTO JournalLines (EntryID, LineNumber, AccountCode, Debit, Credit, LineText, CostCenterID) " & _
            "SELECT e.EntryID, q.LineOrder, q.AccountCode, q.Debit, q.Credit, q.LineText, " & _
            "IIf(q.CostCenter = 0, Null, q.CostCenter) FROM " & sources(i) & _
            " AS q INNER JOIN JournalEntries AS e ON (q.SourceType = e.SourceType AND q.SourceID = e.SourceID) " & _
            "WHERE e.LineCount < 0", dbFailOnError
    Next
    e.Close
    db.Execute "UPDATE JournalEntries SET LineCount = -LineCount WHERE LineCount < 0", dbFailOnError

    ' operations that no longer exist (a deleted expense, ...)
    For Each k In existing.Keys
        If Not seen.Exists(k) Then
            info = existing(k)
            st = Split(k, "|")(0)
            If InClosedPeriod(st, info(2), closed) Then
                locked = locked & "  " & kinds(st) & " (" & GDate(info(2)) & ")" & vbCrLf
            Else
                db.Execute "DELETE FROM JournalLines WHERE EntryID = " & info(0), dbFailOnError
                db.Execute "DELETE FROM JournalEntries WHERE EntryID = " & info(0), dbFailOnError
                Removed = Removed + 1
            End If
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
    StampJournalCurrencies                  ' the currency, rate and foreign amount of each document (modCurrency)
    ws.CommitTrans
    inTrans = False
    If Added + Updated + Removed > 0 Then
        LogAction "JOURNAL_SYNC", "JournalEntries", "", "Added=" & Added & " Updated=" & Updated & " Removed=" & Removed
    End If
    If Len(skipped) > 0 Then
        SyncJournal = "⁄„·Ì«  €Ì— „ Ê«“‰… ·„ Ìı‰‘√ ·Â« ﬁÌœ (—«Ã⁄Â«):" & vbCrLf & Left$(skipped, 700)
    End If
    If Len(locked) > 0 Then
        SyncJournal = SyncJournal & IIf(Len(SyncJournal) > 0, vbCrLf, "") & _
            "⁄„·Ì«  ›Ì › —… „ﬁ›·…  €Ì—  Ê·„ Ì €Ì— ﬁÌœÂ« (√⁄œ › Õ «·› —… ≈‰ ﬂ«‰ «· €ÌÌ— „ﬁ’Êœ«):" & vbCrLf & _
            Left$(locked, 700)
    End If
    Exit Function

EH:
    SyncJournal = " ⁄–—  ÕœÌÀ «·ﬁÌÊœ: " & Err.Description & " (" & Err.Number & ")"
    If inTrans Then ws.Rollback
    Added = 0: Updated = 0: Removed = 0
End Function

Private Function InClosedPeriod(ByVal SourceType As String, ByVal When As Variant, ByVal Closed As Date) As Boolean
    ' A journal entry dated in the closed period is never added, changed or removed, except the year
    ' closing entry, which the closing itself creates and removes.
    If Closed = 0 Or SourceType = "YEAR_CLOSE" Or Not IsDate(When) Then Exit Function
    InClosedPeriod = (DateValue(When) <= Closed)
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
        Case "MANUAL":           OpenScreen "frmManualEntry", 0, id
        Case "YEAR_CLOSE":       OpenScreen "frmPeriodClosing"
        Case "VAT_RETURN":       OpenScreen "frmVatReturn", 0, id
        Case "VAT_PAYMENT":      OpenScreen "frmVatReturn", 0, id
        Case "BANK_OPENING":     OpenScreen "frmBanks", 0, id
        Case "ASSET":            OpenScreen "frmAssets", 0, id
        Case "ASSET_DISPOSAL":   OpenScreen "frmAssets", 0, id
        Case "DEPRECIATION":     OpenScreen "frmDepreciation", 0
        Case "PAYROLL":          OpenScreen "frmPayroll"
        Case "PAYROLL_PAYMENT":  OpenScreen "frmPayroll"
        Case "COMMISSION":       OpenScreen "frmCommissions", 0, id
        Case "CHEQUE":           OpenScreen "frmCheques", 0, Nz(DbValue("SELECT Direction FROM Cheques WHERE ChequeID = " & id), "IN")
        Case "CHEQUE_STATUS":    OpenScreen "frmCheques", 0, Nz(DbValue("SELECT Direction FROM Cheques WHERE ChequeID = " & id), "IN")
        Case "BANK_TX":          OpenScreen "frmBankTx", 0, DbValue("SELECT BankID FROM BankTransactions WHERE BankTxID = " & id)
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
    frm!cboSourceType.RowSource = Tr(rows)
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
    frm!lblSync.Caption = Tr(" ÕœÌÀ «·ﬁÌÊœ „‰ «·⁄„·Ì« : " & added & " ﬁÌœ ÃœÌœ° Ê" & updated & " ﬁÌœ „ıÕœÛ¯À° Ê" & _
                          removed & " ﬁÌœ „Õ–Ê›  (" & GDate(Now, True) & ")")
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
    frm!lstEntries.RowSource = Tr("SELECT e.EntryID, e.EntryNumber AS [—ﬁ„ «·ﬁÌœ], GDate(e.EntryDate) AS [«· «—ÌŒ], " & _
        "t.TypeName AS [«·⁄„·Ì…], e.SourceNumber AS [«·„” ‰œ], e.Description AS [«·»Ì«‰], " & _
        "Format(e.TotalDebit, '#,##0.00') AS [«·„»·€] FROM JournalEntries AS e INNER JOIN JournalSourceTypes AS t " & _
        "ON e.SourceType = t.SourceType WHERE " & w & " ORDER BY e.EntryDate, e.EntryNumber")
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS N, Sum(e.TotalDebit) AS D, Sum(e.TotalCredit) AS C " & _
                                     "FROM JournalEntries AS e WHERE " & w, dbOpenSnapshot)
    frm!lblTotals.Caption = Tr("⁄œœ «·ﬁÌÊœ: " & rs!N & "    ≈Ã„«·Ì «·„œÌ‰: " & Format$(Nz(rs!D, 0), "#,##0.00") & _
                            "    ≈Ã„«·Ì «·œ«∆‰: " & Format$(Nz(rs!C, 0), "#,##0.00") & _
                            IIf(Nz(rs!D, 0) = Nz(rs!C, 0), "    („ Ê«“‰)", "    (€Ì— „ Ê«“‰!)"))
    rs.Close
    If frm!lstEntries.ListCount > 1 Then frm!lstEntries.Value = frm!lstEntries.ItemData(1) Else frm!lstEntries.Value = Null
    JournalEntryPicked frm
End Sub

Public Sub JournalEntryPicked(ByVal frm As Access.Form)
    If IsNull(frm!lstEntries.Value) Then
        frm!lstLines.RowSource = Tr("")
        Exit Sub
    End If
    frm!lstLines.RowSource = Tr(EntryLinesSql(frm!lstEntries.Value))
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
        frm!lblHeader.Caption = Tr("«·ﬁÌœ €Ì— „ÊÃÊœ (—»„« Õı–›  «·⁄„·Ì…). Õœ¯À «·ﬁÌÊœ.")
        Exit Sub
    End If
    frm!lblTitle.Caption = Tr("ﬁÌœ ÌÊ„Ì… —ﬁ„ " & rs!EntryNumber)
    frm!lblHeader.Caption = Tr("«· «—ÌŒ: " & GDate(rs!EntryDate, True) & "    «·⁄„·Ì…: " & rs!TypeName & " " & _
                            Nz(rs!SourceNumber, "") & IIf(IsNull(rs!UpdatedAt), "", "    (Õıœˆ¯À ›Ì " & GDate(rs!UpdatedAt, True) & ")"))
    frm!lblDescription.Caption = Tr("«·»Ì«‰: " & Nz(rs!Description, "-"))
    frm!lblTotals.Caption = Tr("«·≈Ã„«·Ì: „œÌ‰ " & Format$(rs!TotalDebit, "#,##0.00") & " = œ«∆‰ " & _
                            Format$(rs!TotalCredit, "#,##0.00"))
    If Not IsBaseCurrency(rs!CurrencyCode) Then             ' posted in SAR from a document in a currency (modCurrency)
        frm!lblTotals.Caption = Tr(frm!lblTotals.Caption & "    |    «·⁄„·… " & rs!CurrencyCode & "  «·„⁄«„· " & _
            Format$(rs!ExchangeRate, "0.0000") & "  „»·€ «·„” ‰œ " & Format$(rs!ForeignAmount, "#,##0.00") & " " & rs!CurrencyCode)
    End If
    rs.Close
    frm!lstLines.RowSource = Tr(EntryLinesSql(frm.OpenArgs))
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
    Dim manualID As Long, jv As Variant, number As String, opening As Currency, debit As Currency, credit As Currency
    Dim fy As Long, profit As Currency, wasClosed As Date
    Dim vatID As Long, vatFrom As Date, vatTo As Date, vatNet As Currency
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

    ' the account tree
    msg = RebuildAccountTree()
    CheckJournal Len(msg) = 0 And Nz(DbValue("SELECT COUNT(*) FROM Accounts WHERE AccountLevel Is Null OR " & _
                 "TreeKey Is Null OR Level1Code Is Null"), 0) = 0, "‘Ã—… «·Õ”«»« : ·ﬂ· Õ”«» „” Ê«Â Ê„ﬂ«‰Â " & msg, _
                 passed, failed, report
    CheckJournal Nz(DbValue("SELECT COUNT(*) FROM Accounts AS a INNER JOIN Accounts AS p ON a.ParentCode = " & _
                            "p.AccountCode WHERE p.IsPosting = True OR p.AccountType <> a.AccountType"), 0) = 0, _
                 "ﬂ· Õ”«» Ì »⁄ Õ”«»« —∆Ì”Ì« „‰ ‰Ê⁄Â", passed, failed, report
    CheckJournal Nz(DbValue("SELECT COUNT(*) FROM JournalLines AS l INNER JOIN Accounts AS a ON l.AccountCode = " & _
                            "a.AccountCode WHERE a.IsPosting = False"), 0) = 0, "·« ﬁÌÊœ ⁄·Ï «·Õ”«»«  «·—∆Ì”Ì…", _
                 passed, failed, report

    ' the account statement (modLedger): customers and all the assets agree with the account balances
    EnsureLocalTables
    Call FillLedger(1300, DateSerial(2000, 1, 1), Date, opening, debit, credit)
    CheckJournal opening + debit - credit = AccountBalance(1300), "ﬂ‘› Õ”«» «·⁄„·«¡ = —’Ìœ «·Õ”«»", passed, failed, report
    Call FillLedger(1, DateSerial(2000, 1, 1), Date, opening, debit, credit)
    CheckJournal opening + debit - credit = Nz(DbValue("SELECT Sum(l.Debit) - Sum(l.Credit) FROM JournalLines AS l " & _
                 "INNER JOIN Accounts AS a ON l.AccountCode = a.AccountCode WHERE a.AccountType = 'ASSET'"), 0), _
                 "ﬂ‘› «·Õ”«» «·—∆Ì”Ì («·√’Ê·) Ì‘„· ﬂ· Õ”«»« Â «· «»⁄…", passed, failed, report
    CurrentDb.Execute "DELETE FROM tmpLedger", dbFailOnError

    ' the financial statements (modFinancials) over all the entries until today
    Call PrepareFinancials("INCOME", DateSerial(2000, 1, 1), Date)
    CheckJournal Nz(DbValue("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 60"), 0) = _
                 Nz(DbValue("SELECT Sum(l.Credit) - Sum(l.Debit) FROM JournalLines AS l INNER JOIN Accounts AS a ON " & _
                 "l.AccountCode = a.AccountCode WHERE a.AccountType IN ('REVENUE', 'EXPENSE')"), 0), _
                 "’«›Ì «·—»Õ ›Ì ﬁ«∆„… «·œŒ· = «·≈Ì—«œ«  - «·„’—Ê›« ", passed, failed, report
    Call PrepareFinancials("BALANCE", DateSerial(2000, 1, 1), Date)
    CheckJournal Nz(DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 1 AND RowKind = 'T'"), 0) = _
                 Nz(DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 4"), 0), _
                 "«·„Ì“«‰Ì… „ Ê«“‰…: «·√’Ê· = «·Œ’Ê„ + ÕﬁÊﬁ «·„·ﬂÌ…", passed, failed, report

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

    ' a manual entry: refused when unbalanced or on a main account; saved, changed and deleted
    EnsureLocalTables
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (5500, 1000, 0)", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpManualLines (AccountCode, Debit, Credit) VALUES (2310, 0, 900)", dbFailOnError
    CheckJournal Len(PostManualEntry(0, Date, "TEST-MJ", "", Null, manualID)) > 0 And manualID = 0, _
                 "ﬁÌœ ÌœÊÌ €Ì— „ Ê«“‰ Ìı—›÷", passed, failed, report
    CurrentDb.Execute "UPDATE tmpManualLines SET Credit = 1000 WHERE AccountCode = 2310", dbFailOnError
    CurrentDb.Execute "UPDATE tmpManualLines SET AccountCode = 52 WHERE AccountCode = 5500", dbFailOnError
    CheckJournal Len(PostManualEntry(0, Date, "TEST-MJ", "", Null, manualID)) > 0 And manualID = 0, _
                 "ﬁÌœ ÌœÊÌ ⁄·Ï Õ”«» —∆Ì”Ì Ìı—›÷", passed, failed, report
    CurrentDb.Execute "UPDATE tmpManualLines SET AccountCode = 5500 WHERE AccountCode = 52", dbFailOnError
    msg = PostManualEntry(0, Date, "TEST-MJ", "TEST-REF", Null, manualID)
    CheckJournal Len(msg) = 0 And manualID > 0 And _
                 Nz(DbValue("SELECT COUNT(*) FROM ManualEntryLines WHERE ManualEntryID = " & manualID), 0) = 2, _
                 "Õ›Ÿ ﬁÌœ ÌœÊÌ „ Ê«“‰ " & msg, passed, failed, report
    msg = SyncJournal(added, updated, removed)
    jv = JournalEntryOfManual(manualID)
    CheckJournal added = 1 And Not IsNull(jv) And Nz(DbValue("SELECT TotalDebit FROM JournalEntries WHERE EntryID = " & _
                 Nz(jv, 0)), 0) = 1000, "«·ﬁÌœ «·ÌœÊÌ Ì’»Õ ﬁÌœ ÌÊ„Ì… „—»Êÿ« »Â", passed, failed, report
    number = Nz(DbValue("SELECT EntryNumber FROM JournalEntries WHERE EntryID = " & Nz(jv, 0)), "")
    CurrentDb.Execute "UPDATE tmpManualLines SET Debit = 1200 WHERE AccountCode = 5500", dbFailOnError
    CurrentDb.Execute "UPDATE tmpManualLines SET Credit = 1200 WHERE AccountCode = 2310", dbFailOnError
    msg = PostManualEntry(manualID, Date, "TEST-MJ", "", Null, id)
    msg = msg & SyncJournal(added, updated, removed)
    CheckJournal Len(msg) = 0 And updated = 1 And id = manualID And _
                 Nz(DbValue("SELECT EntryNumber FROM JournalEntries WHERE EntryID = " & Nz(jv, 0)), "") = number And _
                 Nz(DbValue("SELECT TotalDebit FROM JournalEntries WHERE EntryID = " & Nz(jv, 0)), 0) = 1200, _
                 " ⁄œÌ· «·ﬁÌœ «·ÌœÊÌ ÌÕœ¯À ﬁÌœÂ »‰›” «·—ﬁ„ " & msg, passed, failed, report
    msg = RemoveManualEntry(manualID)
    msg = msg & SyncJournal(added, updated, removed)
    CheckJournal Len(msg) = 0 And removed = 1 And IsNull(JournalEntryOfManual(manualID)), _
                 "Õ–› «·ﬁÌœ «·ÌœÊÌ ÌÕ–› ﬁÌœÂ", passed, failed, report

    ' closing a period (modClosing): nothing dated in it, then reopened as it was
    wasClosed = ClosedThroughDate()
    If wasClosed < Date - 2 Then
        msg = ClosePeriod(Date - 2, "TEST")
        CheckJournal Len(msg) = 0 And ClosedThroughDate() = Date - 2, "≈ﬁ›«· «·› —… Õ Ï " & GDate(Date - 2) & " " & msg, _
                     passed, failed, report
        CurrentDb.Execute "UPDATE tmpManualLines SET Debit = 100 WHERE AccountCode = 5500", dbFailOnError
        CurrentDb.Execute "UPDATE tmpManualLines SET Credit = 100 WHERE AccountCode = 2310", dbFailOnError
        CheckJournal Len(PostManualEntry(0, Date - 3, "TEST-MJ", "", Null, manualID)) > 0 And manualID = 0, _
                     "≈ﬁ›«· «·› —… Ì„‰⁄ «·ﬁÌœ «·ÌœÊÌ » «—ÌŒ „ﬁ›·", passed, failed, report
        msg = PostManualEntry(0, Date, "TEST-MJ", "", Null, manualID)
        CheckJournal Len(msg) = 0 And manualID > 0, "«·ﬁÌœ » «—ÌŒ „› ÊÕ ÌıÕ›Ÿ " & msg, passed, failed, report
        msg = ReopenPeriod(IIf(wasClosed > 0, wasClosed, Null), "TEST")
        CheckJournal Len(msg) = 0 And ClosedThroughDate() = wasClosed, "≈⁄«œ… › Õ «·› —… " & msg, passed, failed, report
    End If
    CurrentDb.Execute "DELETE FROM tmpManualLines", dbFailOnError

    ' the VAT return of last month (modVat): filed with a balanced settlement entry, paid, back to a draft
    vatFrom = DateSerial(Year(Date), Month(Date) - 1, 1)
    vatTo = DateSerial(Year(Date), Month(Date), 0)
    If Len(VatOverlap(vatFrom, vatTo, 0)) = 0 And ClosedThroughDate() < Date Then
        vatID = 0
        msg = SaveVatDraft(vatFrom, vatTo, 0, 0, vatID)
        CheckJournal Len(msg) = 0 And vatID > 0, "„”Êœ… «·≈ﬁ—«— «·÷—Ì»Ì ··‘Â— «·„«÷Ì " & msg, passed, failed, report
        msg = FileVatReturn(vatID, Date, "TEST")
        vatNet = Nz(DbValue("SELECT SalesStdVAT - PurchStdVAT FROM VatReturns WHERE VatReturnID = " & vatID), 0)
        CheckJournal Len(msg) = 0 And DbValue("SELECT Status FROM VatReturns WHERE VatReturnID = " & vatID) = "FILED" And _
                     (vatNet = 0 Or Not IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = 'VAT_RETURN' " & _
                                                       "AND SourceID = " & vatID))), _
                     "«⁄ „«œ «·≈ﬁ—«— Ì‰‘∆ ﬁÌœ «· ”ÊÌ… " & msg, passed, failed, report
        CheckJournal Nz(DbValue("SELECT Sum(l.Debit) - Sum(l.Credit) FROM JournalLines AS l INNER JOIN JournalEntries AS e " & _
                                "ON l.EntryID = e.EntryID WHERE e.SourceType = 'VAT_RETURN' AND e.SourceID = " & vatID), 0) = 0, _
                     "ﬁÌœ «· ”ÊÌ… „ Ê«“‰", passed, failed, report
        If vatNet > 0 Then
            msg = PayVatReturn(vatID, Date, vatNet, 1200)
            CheckJournal Len(msg) = 0 And Not IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType = " & _
                         "'VAT_PAYMENT' AND SourceID = " & vatID)), "”œ«œ «·÷—Ì»… Ì‰‘∆ ﬁÌœ «·”œ«œ " & msg, passed, failed, report
            msg = CancelVatPayment(vatID)
        End If
        msg = msg & UnfileVatReturn(vatID, "TEST")
        CheckJournal Len(msg) = 0 And IsNull(DbValue("SELECT EntryID FROM JournalEntries WHERE SourceType LIKE 'VAT_*' " & _
                     "AND SourceID = " & vatID)), "≈·€«¡ «·«⁄ „«œ Ê«·”œ«œ ÌÕ–› «·ﬁÌœÌ‰ " & msg, passed, failed, report
    Else
        Debug.Print "[--] «·≈ﬁ—«— «·÷—Ì»Ì: ··‘Â— «·„«÷Ì ≈ﬁ—«— √Ê «·› —… „ﬁ›·…"
    End If

    ' closing last year: revenue and expenses to retained earnings, then reopened (when the years before it are closed)
    fy = Year(Date) - 1
    msg = YearCloseProblem(fy)
    If Len(msg) = 0 Then
        profit = YearNetProfit(fy)
        msg = CloseFiscalYear(fy, "TEST")
        CheckJournal Len(msg) = 0 And YearIsClosed(fy) And YearNetProfit(fy) = 0 And _
                     Nz(DbValue("SELECT NetProfit FROM FiscalYearClosings WHERE FiscalYear = " & fy), 0) = profit, _
                     "≈ﬁ›«· «·”‰… " & fy & ": «·≈Ì—«œ«  Ê«·„’—Ê›«  ≈·Ï «·√—»«Õ «·„Õ Ã“… " & msg, passed, failed, report
        msg = ReopenFiscalYear(fy, "TEST")
        CheckJournal Len(msg) = 0 And Not YearIsClosed(fy) And YearNetProfit(fy) = profit, _
                     "≈⁄«œ… › Õ «·”‰… " & fy & "  ⁄Ìœ √—’œ Â« " & msg, passed, failed, report
    Else
        Debug.Print "[--] ≈ﬁ›«· «·”‰… " & fy & ": " & msg
    End If
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
