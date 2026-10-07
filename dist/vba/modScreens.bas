Attribute VB_Name = "modScreens"
'==============================================================================
' modScreens  -  Retail Store Management System (Phases 5 and 8)
'
' Behaviour of the unbound screens: main menu (frmMain), advanced search
' (frmSearch) and report centre (frmReportCenter).
' Search SQL templates and the report catalogue are in modAppData (generated).
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' frmMain
'------------------------------------------------------------------------------
Public Sub MainLoad(ByVal frm As Access.Form)
    If g_SilentMode Then
        Calendar = vbCalGreg
    Else
        AppStartup
        DoCmd.Maximize
    End If
    frm!lblStoreName.Caption = Nz(SettingValue("StoreName"), APP_TITLE)
    frm!lblToday.Caption = Format$(Date, "dddd  yyyy/mm/dd")
    frm!lblUser.Caption = "«·„” Œœ„: " & CurrentUserName() & "  (" & _
        Nz(DbValue("SELECT r.RoleName FROM Employees AS e INNER JOIN Roles AS r ON e.RoleID = r.RoleID " & _
                   "WHERE e.EmployeeID = " & CurrentUserID()), "") & ")"
    ApplyNavPermissions frm                   ' modSecurityScreens
    DashboardRefresh frm                      ' tiles, lists and the integrity line (modDashboard)
    If Not g_SilentMode Then LowStockAlert    ' once per session
End Sub

Public Sub RefreshIntegrityStatus(ByVal frm As Access.Form)
    Dim n As Long
    If Not HasPermission("REPORTS") Then
        frm!lblIntegrity.Caption = " "
        Exit Sub
    End If
    On Error Resume Next
    n = DCount("*", "IntegrityCheckQuery")
    If Err.Number <> 0 Then
        frm!lblIntegrity.Caption = " ⁄–— ›Õ’ ”·«„… «·»Ì«‰«  (" & Err.Description & ")"
        frm!lblIntegrity.ForeColor = CLR_WARNING
    ElseIf n = 0 Then
        frm!lblIntegrity.Caption = "›Õ’ ”·«„… «·»Ì«‰« : ·«  ÊÃœ „‘ﬂ·« "
        frm!lblIntegrity.ForeColor = CLR_SUCCESS
    Else
        frm!lblIntegrity.Caption = " ‰»ÌÂ: ÌÊÃœ " & n & " „‘ﬂ·… ›Ì «·»Ì«‰«  - —«Ã⁄  ﬁ—Ì— ›Õ’ «·”·«„…"
        frm!lblIntegrity.ForeColor = CLR_DANGER
    End If
End Sub

'------------------------------------------------------------------------------
' frmSearch
'------------------------------------------------------------------------------
Public Sub SearchLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    If IsNull(frm!cboKind.Value) Then frm!cboKind.Value = "PRODUCT"
    SearchKindChanged frm
End Sub

Public Sub SearchKindChanged(ByVal frm As Access.Form)
    Dim kind As String, dated As Boolean
    kind = Nz(frm!cboKind.Value, "PRODUCT")
    dated = (kind = "SALE" Or kind = "PURCHASE")
    frm!txtFrom.Enabled = dated
    frm!txtTo.Enabled = dated
    frm!lstResults.RowSource = ""
    frm!lstResults.ColumnCount = SearchColumnCount(kind)
    frm!lstResults.ColumnWidths = SearchColumnWidths(kind)
    frm!lblCount.Caption = " "
End Sub

Public Sub RunSearch(ByVal frm As Access.Form)
    Dim kind As String, txt As String, sql As String, num As String, n As Long
    kind = Nz(frm!cboKind.Value, "PRODUCT")
    txt = Trim$(Nz(frm!txtText.Value, ""))

    num = "-1"                                 ' {NUM}: exact record number search
    If Len(txt) > 0 And Len(txt) <= 9 And txt Like String$(Len(txt), "#") Then num = txt

    sql = SearchTemplate(kind)
    sql = Replace(sql, "{LIKE}", LikePattern(txt))
    sql = Replace(sql, "{NUM}", num)
    If IsDate(frm!txtFrom.Value) Then
        sql = Replace(sql, "{FROM}", SqlDate(DateValue(frm!txtFrom.Value)))
    Else
        sql = Replace(sql, "{FROM}", SqlDate(#1/1/1900#))
    End If
    If IsDate(frm!txtTo.Value) Then
        sql = Replace(sql, "{TO}", SqlDate(DateAdd("d", 1, DateValue(frm!txtTo.Value))))
    Else
        sql = Replace(sql, "{TO}", SqlDate(#12/31/9999#))
    End If

    frm!lstResults.RowSource = sql
    n = frm!lstResults.ListCount - 1           ' first row = column headings
    If n < 0 Then n = 0
    frm!lblCount.Caption = n & " ‰ ÌÃ…" & IIf(n > 0, " - «‰ﬁ— „— Ì‰ ⁄·Ï «·‰ ÌÃ… ·› ÕÂ«", "")
End Sub

Public Sub SearchOpen(ByVal frm As Access.Form)
    Dim id As Variant
    id = frm!lstResults.Value
    If IsNull(id) Then Exit Sub
    Select Case Nz(frm!cboKind.Value, "PRODUCT")
        Case "PRODUCT":  OpenScreen "frmProducts", 0, id
        Case "CUSTOMER": OpenScreen "frmCustomers", 0, id
        Case "SUPPLIER": OpenScreen "frmSuppliers", 0, id
        Case "SALE":     OpenScreen "frmSalesInvoice", 6, id
        Case "PURCHASE": OpenScreen "frmPurchaseView", 7, id
    End Select
End Sub

Public Sub SearchClear(ByVal frm As Access.Form)
    frm!txtText.Value = Null
    frm!txtFrom.Value = Null
    frm!txtTo.Value = Null
    SearchKindChanged frm
    SafeFocus frm!txtText
End Sub

'------------------------------------------------------------------------------
' frmReportCenter
' Catalogue row: Array(Key, Title, QueryName, ReportName, Needs, DateColumn)
' Needs letters:  P = period parameters (SetPeriod)
'                 D = filter DateColumn by the period
'                 C / S / R = customer / supplier / product required
'                 c / s / r = optional filter on CustomerID / SupplierID / ProductID
'                 b = optional cash box (query parameter CashBoxID, 0 = all boxes)
'                 e = optional filter on ExpenseTypeID
'                 # = treasury report: needs the CASH_BOX permission
'                 J = accounting report: needs the JOURNAL permission; the entries are
'                     brought up to date first (SyncJournal)
'                 $ = shows cost / profit: needs the REPORTS_PROFIT permission
'------------------------------------------------------------------------------
Public Sub ReportCenterLoad(ByVal frm As Access.Form)
    Dim i As Long, r As Variant, rows As String
    Calendar = vbCalGreg
    For i = 1 To ReportCount()
        r = ReportRow(i)
        If Len(rows) > 0 Then rows = rows & ";"
        rows = rows & """" & r(0) & """;""" & r(1) & """"
    Next
    frm!lstReports.RowSource = rows
    frm!txtFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    frm!txtTo.Value = Date
    ReportSelected frm
End Sub

Public Sub ReportSelected(ByVal frm As Access.Form)
    Dim r As Variant, needs As String, dated As Boolean, hint As String
    r = SelectedReport(frm)
    If IsEmpty(r) Then
        frm!lblReportTitle.Caption = "«Œ —  ﬁ—Ì—« „‰ «·ﬁ«∆„…"
    Else
        frm!lblReportTitle.Caption = r(1)
        needs = r(4)
    End If
    dated = HasNeed(needs, "P") Or HasNeed(needs, "D")
    frm!txtFrom.Enabled = dated
    frm!txtTo.Enabled = dated
    frm!btnToday.Enabled = dated
    frm!btnThisMonth.Enabled = dated
    frm!btnLastMonth.Enabled = dated
    frm!btnThisYear.Enabled = dated
    frm!cboCustomer.Enabled = HasNeed(needs, "C") Or HasNeed(needs, "c")
    frm!cboSupplier.Enabled = HasNeed(needs, "S") Or HasNeed(needs, "s")
    frm!cboProduct.Enabled = HasNeed(needs, "R") Or HasNeed(needs, "r")
    frm!cboCashBox.Enabled = HasNeed(needs, "b")
    frm!cboExpenseType.Enabled = HasNeed(needs, "e")
    frm!btnRun.Enabled = Not IsEmpty(r)

    If dated Then hint = "Õœœ «·› —…"
    If HasNeed(needs, "C") Then hint = hint & IIf(Len(hint) > 0, " Ê", "") & "«Œ — «·⁄„Ì·"
    If HasNeed(needs, "S") Then hint = hint & IIf(Len(hint) > 0, " Ê", "") & "«Œ — «·„Ê—œ"
    If HasNeed(needs, "R") Then hint = hint & IIf(Len(hint) > 0, " Ê", "") & "«Œ — «·„‰ Ã"
    If HasNeed(needs, "c") Or HasNeed(needs, "s") Or HasNeed(needs, "r") Or HasNeed(needs, "b") Or _
       HasNeed(needs, "e") Then
        hint = hint & IIf(Len(hint) > 0, "° ", "") & "ÊÌ„ﬂ‰ﬂ «· ’›Ì… Õ”» «·«Œ Ì«— («Œ Ì«—Ì)"
    End If
    If Len(hint) = 0 And Not IsEmpty(r) Then hint = "·« ÌÕ «Ã Â–« «· ﬁ—Ì— √Ì «Œ Ì«—« "
    frm!lblNeeds.Caption = IIf(Len(hint) = 0, " ", hint)
End Sub

Public Sub SetQuickPeriod(ByVal frm As Access.Form, ByVal Which As String)
    Dim y As Integer, m As Integer
    y = Year(Date)
    m = Month(Date)
    Select Case Which
        Case "TODAY":     frm!txtFrom.Value = Date: frm!txtTo.Value = Date
        Case "MONTH":     frm!txtFrom.Value = DateSerial(y, m, 1): frm!txtTo.Value = Date
        Case "LASTMONTH": frm!txtFrom.Value = DateSerial(y, m - 1, 1): frm!txtTo.Value = DateSerial(y, m, 0)
        Case "YEAR":      frm!txtFrom.Value = DateSerial(y, 1, 1): frm!txtTo.Value = Date
    End Select
End Sub

Public Sub RunReport(ByVal frm As Access.Form)
    Dim r As Variant, where As String, criteria As String
    If Not PrepareReport(frm, r, where, criteria) Then Exit Sub
    LogAction "REPORT", CStr(r(0))
    OpenReportOrQuery CStr(r(3)), CStr(r(2)), where, criteria
End Sub

Public Sub ExportReport(ByVal frm As Access.Form, ByVal FileKind As String)
    ' FileKind: "PDF" (the formatted report) or "XLSX" (the report data for Excel).
    Dim r As Variant, where As String, criteria As String, file As String
    If Not PrepareReport(frm, r, where, criteria) Then Exit Sub
    On Error GoTo EH
    file = ReportsFolder() & "" & r(0) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & LCase$(FileKind)
    If FileKind = "PDF" Then
        If Not ReportExists(CStr(r(3))) Then
            ShowWarning "«· ﬁ—Ì— €Ì— „ÊÃÊœ: " & r(3) & vbCrLf & "‘€¯· BuildReports."
            Exit Sub
        End If
        TempVars.Add "ReportCriteria", criteria
        DoCmd.OpenReport CStr(r(3)), acViewPreview, , where, acHidden
        DoCmd.OutputTo acOutputReport, CStr(r(3)), acFormatPDF, file
        DoCmd.Close acReport, CStr(r(3)), acSaveNo
    Else
        WritePreviewQuery CStr(r(2)), where
        DoCmd.OutputTo acOutputQuery, "qryReportPreview", acFormatXLSX, file
    End If
    LogAction "REPORT_EXPORT", CStr(r(0)), , file
    ShowInfo " „ «·Õ›Ÿ ›Ì:" & vbCrLf & file
    If Not g_SilentMode Then Application.FollowHyperlink file
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError " ⁄–— «· ’œÌ—: " & Err.Description   ' 2501 = no data
End Sub

Private Function PrepareReport(ByVal frm As Access.Form, ByRef r As Variant, ByRef where As String, _
                               ByRef criteria As String) As Boolean
    ' Checks the choices, sets the query parameters and builds the filter and the criteria line.
    Dim needs As String, fromDate As Date, toDate As Date, msg As String
    r = SelectedReport(frm)
    If IsEmpty(r) Then
        ShowWarning "«Œ —  ﬁ—Ì—« „‰ «·ﬁ«∆„… √Ê·«."
        Exit Function
    End If
    needs = r(4)
    where = ""
    criteria = ""
    If HasNeed(needs, "$") And Not HasPermission("REPORTS_PROFIT") Then
        ShowWarning "Â–« «· ﬁ—Ì— Ì⁄—÷ «· ﬂ·›… Ê«·√—»«Õ ÊÌÕ «Ã ’·«ÕÌ… ´ ﬁ«—Ì— «·√—»«Õ Ê«·÷—Ì»…ª."
        Exit Function
    End If
    If HasNeed(needs, "#") And Not HasPermission("CASH_BOX") Then
        ShowWarning " ﬁ«—Ì— «·Œ“Ì‰…  Õ «Ã ’·«ÕÌ… ´«·Œ“Ì‰…ª."
        Exit Function
    End If
    If HasNeed(needs, "J") Then
        If Not HasPermission("JOURNAL") Then
            ShowWarning "«· ﬁ«—Ì— «·„Õ«”»Ì…  Õ «Ã ’·«ÕÌ… ´ﬁÌÊœ «·ÌÊ„Ì…ª."
            Exit Function
        End If
        msg = SyncJournal()
        If Len(msg) > 0 Then ShowWarning msg
        SetQueryParam "AccountCode", 0             ' general ledger from here: every account
    End If

    If HasNeed(needs, "P") Or HasNeed(needs, "D") Then
        If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
            ShowWarning "√œŒ·  «—ÌŒ «·»œ«Ì… Ê «—ÌŒ «·‰Â«Ì…."
            Exit Function
        End If
        fromDate = DateValue(frm!txtFrom.Value)
        toDate = DateValue(frm!txtTo.Value)
        If toDate < fromDate Then
            ShowWarning " «—ÌŒ «·‰Â«Ì… ÌÃ» √‰ ÌﬂÊ‰ »⁄œ  «—ÌŒ «·»œ«Ì…."
            Exit Function
        End If
        SetPeriod fromDate, toDate
        criteria = PeriodText(fromDate, toDate)
    End If
    If Not RequireChoice(frm!cboCustomer, HasNeed(needs, "C"), "CustomerID", "«·⁄„Ì·") Then Exit Function
    If Not RequireChoice(frm!cboSupplier, HasNeed(needs, "S"), "SupplierID", "«·„Ê—œ") Then Exit Function
    If Not RequireChoice(frm!cboProduct, HasNeed(needs, "R"), "ProductID", "«·„‰ Ã") Then Exit Function

    If HasNeed(needs, "D") Then
        where = "[" & r(5) & "] >= " & SqlDate(fromDate) & " AND [" & r(5) & "] < " & _
                SqlDate(DateAdd("d", 1, toDate))
    End If
    If HasNeed(needs, "c") Then AddFilter where, frm!cboCustomer, "CustomerID"
    If HasNeed(needs, "s") Then AddFilter where, frm!cboSupplier, "SupplierID"
    If HasNeed(needs, "r") Then AddFilter where, frm!cboProduct, "ProductID"
    If HasNeed(needs, "e") Then AddFilter where, frm!cboExpenseType, "ExpenseTypeID"
    If HasNeed(needs, "b") Then SetQueryParam "CashBoxID", CLng(Nz(frm!cboCashBox.Value, 0))   ' 0 = all boxes
    AddCriteria criteria, frm!cboCustomer, HasNeed(needs, "C") Or HasNeed(needs, "c"), "«·⁄„Ì·"
    AddCriteria criteria, frm!cboSupplier, HasNeed(needs, "S") Or HasNeed(needs, "s"), "«·„Ê—œ"
    AddCriteria criteria, frm!cboProduct, HasNeed(needs, "R") Or HasNeed(needs, "r"), "«·„‰ Ã"
    AddCriteria criteria, frm!cboCashBox, HasNeed(needs, "b"), "«·’‰œÊﬁ"
    AddCriteria criteria, frm!cboExpenseType, HasNeed(needs, "e"), "‰Ê⁄ «·„’—Ê›"
    If HasNeed(needs, "b") And IsNull(frm!cboCashBox.Value) Then criteria = criteria & "    ﬂ· «·’‰«œÌﬁ"
    PrepareReport = True
End Function

Private Sub AddCriteria(ByRef criteria As String, ByVal cbo As Access.ComboBox, ByVal Used As Boolean, _
                        ByVal Caption As String)
    If Not Used Or IsNull(cbo.Value) Then Exit Sub
    If Len(criteria) > 0 Then criteria = criteria & "    "
    criteria = criteria & Caption & ": " & cbo.Column(1)
End Sub

Public Sub OpenReportOrQuery(ByVal ReportName As String, ByVal QueryName As String, _
                             ByVal WhereCondition As String, Optional ByVal Criteria As String = "")
    ' The formatted report when it exists (Phase 8), otherwise the query as a table.
    ' Criteria is the line printed under the report title (period, customer, ...).
    On Error GoTo EH
    If ReportExists(ReportName) Then
        TempVars.Add "ReportCriteria", Criteria
        DoCmd.OpenReport ReportName, acViewPreview, , WhereCondition
        Exit Sub
    End If
    WritePreviewQuery QueryName, WhereCondition
    DoCmd.OpenQuery "qryReportPreview", acViewNormal, acReadOnly
    Exit Sub
EH:
    If Err.Number <> 2501 Then ShowError " ⁄–— › Õ «· ﬁ—Ì—: " & Err.Description   ' 2501 = no data
End Sub

Private Sub WritePreviewQuery(ByVal QueryName As String, ByVal WhereCondition As String)
    Dim sql As String, db As DAO.Database
    sql = "SELECT * FROM [" & QueryName & "]"
    If Len(WhereCondition) > 0 Then sql = sql & " WHERE " & WhereCondition
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete "qryReportPreview"
    On Error GoTo 0
    db.CreateQueryDef "qryReportPreview", sql
End Sub

Private Function SelectedReport(ByVal frm As Access.Form) As Variant
    Dim i As Long, r As Variant, key As String
    key = Nz(frm!lstReports.Value, "")
    If Len(key) = 0 Then Exit Function
    For i = 1 To ReportCount()
        r = ReportRow(i)
        If r(0) = key Then
            SelectedReport = r
            Exit Function
        End If
    Next
End Function

Private Function HasNeed(ByVal NeedsText As String, ByVal Letter As String) As Boolean
    HasNeed = InStr(1, NeedsText, Letter, vbBinaryCompare) > 0   ' "C" and "c" differ
End Function

Private Function RequireChoice(ByVal cbo As Access.ComboBox, ByVal Required As Boolean, _
                               ByVal ParamName As String, ByVal Caption As String) As Boolean
    RequireChoice = True
    If Not Required Then Exit Function
    If IsNull(cbo.Value) Then
        ShowWarning "«Œ — " & Caption & "."
        SafeFocus cbo
        RequireChoice = False
        Exit Function
    End If
    SetQueryParam ParamName, CLng(cbo.Value)
End Function

Private Sub AddFilter(ByRef where As String, ByVal cbo As Access.ComboBox, ByVal FieldName As String)
    If IsNull(cbo.Value) Then Exit Sub
    If Len(where) > 0 Then where = where & " AND "
    where = where & "[" & FieldName & "] = " & CLng(cbo.Value)
End Sub
