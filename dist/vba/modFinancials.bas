Attribute VB_Name = "modFinancials"
'==============================================================================
' modFinancials  -  Retail Store Management System
'
' Financial statements from the journal (screen frmFinancials, reports
' rptIncomeStatement and rptBalanceSheet, also in the report centre):
'   income statement   revenue, cost of sales, gross profit, operating expenses,
'                      operating profit, other revenue / expenses, net profit
'                      (IncomeStatementQuery, the period From..To);
'   balance sheet      assets = liabilities + equity at the end of the period, the
'                      profit of the periods not closed yet shown in equity
'                      (BalanceSheetQuery).
' Both with a comparison column (modQueryParams.SetComparePeriod): a month ->
' the month before, Jan 1 - today -> the same dates last year, other periods ->
' the same number of days just before. The rows come in their printed order.
'==============================================================================
Option Compare Database
Option Explicit

Public Function PrepareFinancials(ByVal Kind As String, ByVal FromDate As Date, ByVal ToDate As Date) As String
    ' Sets the period and the comparison period; returns the criteria line of the report.
    ' Kind: "INCOME" / "INCOME_STATEMENT" or "BALANCE" / "BALANCE_SHEET".
    SetPeriod FromDate, ToDate
    SetComparePeriod FromDate, ToDate
    If Left$(Kind, 7) = "BALANCE" Then
        PrepareFinancials = "›Ì " & GDate(ToDate) & "    «·„ﬁ«—‰…: ›Ì " & GDate(DateAdd("d", -1, QDate("CompareEnd")))
    Else
        PrepareFinancials = PeriodText(FromDate, ToDate) & "    «·„ﬁ«—‰…: „‰ " & GDate(QDate("CompareStart")) & _
                            " ≈·Ï " & GDate(DateAdd("d", -1, QDate("CompareEnd")))
    End If
End Function

Private Function StatementQuery(ByVal Kind As String) As String
    StatementQuery = IIf(Kind = "BALANCE", "BalanceSheetQuery", "IncomeStatementQuery")
End Function

'------------------------------------------------------------------------------
' Screen frmFinancials (OpenArgs: "INCOME" or "BALANCE")
'------------------------------------------------------------------------------
Public Sub FinancialsLoad(ByVal frm As Access.Form)
    Dim msg As String
    Calendar = vbCalGreg
    msg = SyncJournal()                                ' the statements show every operation
    If Len(msg) > 0 Then ShowWarning msg
    frm!cboStatement.Value = IIf(Nz(frm.OpenArgs, "") = "BALANCE", "BALANCE", "INCOME")
    SetQuickPeriod frm, "YEAR"
    FinancialsRefresh frm
End Sub

Public Sub FinancialsQuickPeriod(ByVal frm As Access.Form, ByVal Which As String)
    SetQuickPeriod frm, Which
    FinancialsRefresh frm
End Sub

Private Function DatesOK(ByVal frm As Access.Form) As Boolean
    If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
    ElseIf DateValue(frm!txtTo.Value) < DateValue(frm!txtFrom.Value) Then
        ShowWarning " «—ÌŒ «·‰Â«Ì… ﬁ»·  «—ÌŒ «·»œ«Ì…."
    Else
        DatesOK = True
    End If
End Function

Public Sub FinancialsRefresh(ByVal frm As Access.Form)
    Dim kind As String, order As String, amounts As String
    Calendar = vbCalGreg
    frm!lstRows.RowSource = ""
    If Not DatesOK(frm) Then Exit Sub
    kind = Nz(frm!cboStatement.Value, "INCOME")
    frm!lblCompare.Caption = PrepareFinancials(kind, DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value))
    amounts = "IIf(RowKind = 'A', Format(CurrentValue, '#,##0.00'), Null) AS [«·› —…: «·Õ”«»], " & _
              "IIf(RowKind IN ('S', 'T', 'R'), Format(CurrentValue, '#,##0.00'), Null) AS [«·› —…: «·„Ã„Ê⁄], " & _
              "IIf(RowKind = 'A', Format(PriorValue, '#,##0.00'), Null) AS [«·„ﬁ«—‰…: «·Õ”«»], " & _
              "IIf(RowKind IN ('S', 'T', 'R'), Format(PriorValue, '#,##0.00'), Null) AS [«·„ﬁ«—‰…: «·„Ã„Ê⁄]"
    If kind = "BALANCE" Then
        order = "ClassNo, GroupKey, Pos, AccountKey"
    Else
        order = "Block, AccountKey"
    End If
    frm!lstRows.RowSource = "SELECT LineAccount, IIf(RowKind = 'A', '      ' & Caption, IIf(RowKind = 'S', " & _
        "'≈Ã„«·Ì ' & Caption, Caption)) AS [«·»‰œ], " & amounts & " FROM " & StatementQuery(kind) & " ORDER BY " & order
    If kind = "BALANCE" Then
        ShowFigure frm, 1, "≈Ã„«·Ì «·√’Ê·", DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 1 AND RowKind = 'T'")
        ShowFigure frm, 2, "≈Ã„«·Ì «·Œ’Ê„ ÊÕﬁÊﬁ «·„·ﬂÌ…", DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 4")
        ShowFigure frm, 3, "«·›—ﬁ (’›— = „ Ê«“‰…)", Nz(DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 1 " & _
                   "AND RowKind = 'T'"), 0) - Nz(DbValue("SELECT CurrentValue FROM BalanceSheetQuery WHERE ClassNo = 4"), 0)
    Else
        ShowFigure frm, 1, "’«›Ì ≈Ì—«œ«  «·‰‘«ÿ", DbValue("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 12")
        ShowFigure frm, 2, "„Ã„· «·—»Õ", DbValue("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 25")
        ShowFigure frm, 3, "’«›Ì «·—»Õ («·Œ”«—…)", DbValue("SELECT CurrentValue FROM IncomeStatementQuery WHERE Block = 60")
    End If
End Sub

Private Sub ShowFigure(ByVal frm As Access.Form, ByVal Index As Long, ByVal Caption As String, ByVal Amount As Variant)
    frm.Controls("lblCap" & Index).Caption = Caption
    frm.Controls("lblVal" & Index).Caption = Format$(Nz(Amount, 0), "#,##0.00")
    frm.Controls("lblVal" & Index).ForeColor = IIf(Nz(Amount, 0) < 0, CLR_DANGER, CLR_PRIMARY)
End Sub

Public Sub FinancialsOpenLedger(ByVal frm As Access.Form)
    ' the account statement of the picked account, for the same period
    Dim account As Variant, ledger As Access.Form
    account = frm!lstRows.Value
    If IsNull(account) Then
        ShowWarning "«Œ — ”ÿ— Õ”«» „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    If Not DatesOK(frm) Then Exit Sub
    If Not OpenScreen("frmLedger", 0, account) Then Exit Sub
    Set ledger = Forms("frmLedger")
    ledger!txtFrom.Value = DateValue(frm!txtFrom.Value)
    ledger!txtTo.Value = DateValue(frm!txtTo.Value)
    LedgerRefresh ledger
End Sub

Public Sub PrintFinancials(ByVal frm As Access.Form)
    Dim kind As String, criteria As String
    If Not DatesOK(frm) Then Exit Sub
    kind = Nz(frm!cboStatement.Value, "INCOME")
    criteria = PrepareFinancials(kind, DateValue(frm!txtFrom.Value), DateValue(frm!txtTo.Value))
    LogAction "REPORT", IIf(kind = "BALANCE", "BALANCE_SHEET", "INCOME_STATEMENT")
    OpenReportOrQuery IIf(kind = "BALANCE", "rptBalanceSheet", "rptIncomeStatement"), StatementQuery(kind), "", criteria
End Sub
