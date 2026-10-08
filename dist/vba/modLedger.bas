Attribute VB_Name = "modLedger"
'==============================================================================
' modLedger  -  Retail Store Management System
'
' Account statement and general ledger (screen frmLedger):
'   any account of the tree for a period: the opening balance, then every journal
'   line with the balance after it, and the four figures (opening, debit, credit,
'   closing). A main account takes all its sub-accounts (AccountStatementQuery).
'   The lines shown are kept in the local table tmpLedger.
'   Print: rptAccountStatement (this account) and rptGeneralLedger (each
'   sub-account of this account, or of every account when none is chosen).
' Every line opens its journal entry or the operation it came from.
'==============================================================================
Option Compare Database
Option Explicit

Public Function BalanceText(ByVal Amount As Currency) As String
    ' 1,250.00 „œÌ‰ / 300.00 œ«∆‰ / 0.00
    If Amount > 0 Then
        BalanceText = Format$(Amount, "#,##0.00") & " „œÌ‰"
    ElseIf Amount < 0 Then
        BalanceText = Format$(-Amount, "#,##0.00") & " œ«∆‰"
    Else
        BalanceText = "0.00"
    End If
End Function

Private Function MoneyText(ByVal Amount As Variant) As Variant
    If Nz(Amount, 0) = 0 Then MoneyText = Null Else MoneyText = Format$(Amount, "#,##0.00")
End Function

'------------------------------------------------------------------------------
' Statement (OpenArgs: the account to show)
'------------------------------------------------------------------------------
Public Sub LedgerLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    EnsureLocalTables
    msg = SyncJournal()                                ' the statement shows every operation
    If Len(msg) > 0 Then ShowWarning msg
    SetQuickPeriod frm, "YEAR"
    If Nz(frm.OpenArgs, 0) > 0 Then frm!cboAccount.Value = CLng(frm.OpenArgs)
    LedgerRefresh frm
End Sub

Public Sub LedgerQuickPeriod(ByVal frm As Access.Form, ByVal Which As String)
    SetQuickPeriod frm, Which
    LedgerRefresh frm
End Sub

Public Function FillLedger(ByVal AccountCode As Long, ByVal FromDate As Date, ByVal ToDate As Date, _
                           ByRef Opening As Currency, ByRef Debit As Currency, ByRef Credit As Currency) As Long
    ' tmpLedger = the statement of the account; returns the number of journal lines.
    Dim db As DAO.Database, rs As DAO.Recordset, t As DAO.Recordset, balance As Currency, n As Long
    Set db = CurrentDb
    Opening = 0: Debit = 0: Credit = 0
    db.Execute "DELETE FROM tmpLedger", dbFailOnError
    SetPeriod FromDate, ToDate
    SetQueryParam "AccountCode", AccountCode
    Set rs = db.OpenRecordset("SELECT * FROM AccountStatementQuery ORDER BY SortKey, LineDate, EntryNo", dbOpenSnapshot)
    Set t = db.OpenRecordset("tmpLedger", dbOpenDynaset, dbAppendOnly)
    Do Until rs.EOF
        balance = balance + Nz(rs!LineDebit, 0) - Nz(rs!LineCredit, 0)
        If rs!SortKey = 0 Then
            Opening = Nz(rs!LineDebit, 0) - Nz(rs!LineCredit, 0)
        Else
            Debit = Debit + Nz(rs!LineDebit, 0)
            Credit = Credit + Nz(rs!LineCredit, 0)
            n = n + 1
        End If
        t.AddNew
        t!EntryRef = Nz(rs!EntryRef, 0)
        t!DateText = Format$(rs!LineDate, "yyyy/mm/dd")
        t!EntryNo = rs!EntryNo
        t!KindName = rs!KindName
        t!DocNo = rs!DocNo
        If Len(Nz(rs!Details, "")) > 0 Then t!Details = Left$(rs!Details, 255)
        t!SubName = rs!SubName
        t!DebitText = MoneyText(rs!LineDebit)
        t!CreditText = MoneyText(rs!LineCredit)
        t!BalanceText = BalanceText(balance)
        t.Update
        rs.MoveNext
    Loop
    rs.Close
    t.Close
    FillLedger = n
End Function

Public Sub LedgerRefresh(ByVal frm As Access.Form)
    Dim n As Long, opening As Currency, debit As Currency, credit As Currency, isMain As Boolean
    Calendar = vbCalGreg
    frm!lstLines.RowSource = Tr("")
    If IsNull(frm!cboAccount.Value) Then
        LedgerFigures frm, Null, Null, Null, Null
        frm!lblInfo.Caption = Tr("«Œ — «·Õ”«»° À„ «·› —….")
        Exit Sub
    End If
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
        Exit Sub
    End If
    If DateValue(frm!txtTo.Value) < DateValue(frm!txtFrom.Value) Then
        ShowWarning " «—ÌŒ «·‰Â«Ì… ﬁ»·  «—ÌŒ «·»œ«Ì…."
        Exit Sub
    End If
    n = FillLedger(frm!cboAccount.Value, DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value), opening, debit, credit)
    frm!lstLines.RowSource = Tr("SELECT EntryRef, DateText AS [«· «—ÌŒ], EntryNo AS [«·ﬁÌœ], KindName AS [«·⁄„·Ì…], " & _
        "DocNo AS [«·„” ‰œ], Details AS [«·»Ì«‰], SubName AS [«·Õ”«»], DebitText AS [„œÌ‰], CreditText AS [œ«∆‰], " & _
        "BalanceText AS [«·—’Ìœ] FROM tmpLedger ORDER BY LineNo")
    LedgerFigures frm, opening, debit, credit, opening + debit - credit
    isMain = Not Nz(DbValue("SELECT IsPosting FROM Accounts WHERE AccountCode = " & frm!cboAccount.Value), True)
    frm!lblInfo.Caption = Tr(n & " ”ÿ— ﬁÌœ" & IIf(isMain, "   (Õ”«» —∆Ì”Ì: Ì‘„· ﬂ· Õ”«»« Â «· «»⁄…)", "") & _
                          "   ‰ﬁ— „“œÊÃ ⁄·Ï «·”ÿ— Ì› Õ «·ﬁÌœ")
End Sub

Private Sub LedgerFigures(ByVal frm As Access.Form, ByVal Opening As Variant, ByVal Debit As Variant, _
                          ByVal Credit As Variant, ByVal Closing As Variant)
    frm!lblOpening.Caption = Tr(IIf(IsNull(Opening), "-", BalanceText(Nz(Opening, 0))))
    frm!lblDebit.Caption = Tr(IIf(IsNull(Debit), "-", Format$(Nz(Debit, 0), "#,##0.00")))
    frm!lblCredit.Caption = Tr(IIf(IsNull(Credit), "-", Format$(Nz(Credit, 0), "#,##0.00")))
    frm!lblClosing.Caption = Tr(IIf(IsNull(Closing), "-", BalanceText(Nz(Closing, 0))))
End Sub

Private Function PickedEntry(ByVal frm As Access.Form) As Long
    PickedEntry = Nz(frm!lstLines.Value, 0)
    If PickedEntry = 0 Then ShowWarning "«Œ — ”ÿ— ﬁÌœ „‰ «·ﬁ«∆„… (”ÿ— —’Ìœ √Ê· «·„œ… ·Ì” ﬁÌœ«)."
End Function

Public Sub LedgerOpenEntry(ByVal frm As Access.Form)
    Dim id As Long
    id = PickedEntry(frm)
    If id > 0 Then OpenScreen "frmJournalEntry", 0, id
End Sub

Public Sub LedgerOpenSource(ByVal frm As Access.Form)
    Dim id As Long
    id = PickedEntry(frm)
    If id > 0 Then OpenJournalSource id
End Sub

Public Sub PrintLedger(ByVal frm As Access.Form, ByVal Which As String)
    ' STATEMENT: this account; LEDGER: each sub-account of this account (every account when none is chosen)
    Dim fromDate As Date, toDate As Date, criteria As String, account As Long
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
        Exit Sub
    End If
    fromDate = DateValue(frm!txtFrom.Value)
    toDate = DateValue(frm!txtTo.Value)
    account = Nz(frm!cboAccount.Value, 0)
    If Which = "STATEMENT" And account = 0 Then
        ShowWarning "«Œ — «·Õ”«»."
        Exit Sub
    End If
    SetPeriod fromDate, toDate
    SetQueryParam "AccountCode", account
    criteria = PeriodText(fromDate, toDate)
    If Which = "STATEMENT" Then
        LogAction "REPORT", "ACCOUNT_STATEMENT", CStr(account)
        OpenReportOrQuery "rptAccountStatement", "AccountStatementQuery", "", criteria
    Else
        If account > 0 Then criteria = criteria & "    " & frm!cboAccount.Column(1)
        LogAction "REPORT", "GENERAL_LEDGER", CStr(account)
        OpenReportOrQuery "rptGeneralLedger", "GeneralLedgerQuery", "", criteria
    End If
End Sub
