Attribute VB_Name = "modScreens"
'==============================================================================
' modScreens  -  Retail Store Management System (Phase 5)
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
    frm!lblUser.Caption = "المستخدم: " & CurrentUserName()
    RefreshIntegrityStatus frm
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
        frm!lblIntegrity.Caption = "تعذر فحص سلامة البيانات (" & Err.Description & ")"
        frm!lblIntegrity.ForeColor = CLR_WARNING
    ElseIf n = 0 Then
        frm!lblIntegrity.Caption = "فحص سلامة البيانات: لا توجد مشكلات"
        frm!lblIntegrity.ForeColor = CLR_SUCCESS
    Else
        frm!lblIntegrity.Caption = "تنبيه: يوجد " & n & " مشكلة في البيانات - راجع تقرير فحص السلامة"
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
    frm!lblCount.Caption = n & " نتيجة" & IIf(n > 0, " - انقر مرتين على النتيجة لفتحها", "")
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
        frm!lblReportTitle.Caption = "اختر تقريرًا من القائمة"
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
    frm!btnRun.Enabled = Not IsEmpty(r)

    If dated Then hint = "حدد الفترة"
    If HasNeed(needs, "C") Then hint = hint & IIf(Len(hint) > 0, " و", "") & "اختر العميل"
    If HasNeed(needs, "S") Then hint = hint & IIf(Len(hint) > 0, " و", "") & "اختر المورد"
    If HasNeed(needs, "R") Then hint = hint & IIf(Len(hint) > 0, " و", "") & "اختر المنتج"
    If HasNeed(needs, "c") Or HasNeed(needs, "s") Or HasNeed(needs, "r") Then
        hint = hint & IIf(Len(hint) > 0, "، ", "") & "ويمكنك التصفية حسب الاختيار (اختياري)"
    End If
    If Len(hint) = 0 And Not IsEmpty(r) Then hint = "لا يحتاج هذا التقرير أي اختيارات"
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
    Dim r As Variant, needs As String, where As String, fromDate As Date, toDate As Date
    r = SelectedReport(frm)
    If IsEmpty(r) Then
        ShowWarning "اختر تقريرًا من القائمة أولًا."
        Exit Sub
    End If
    needs = r(4)

    If HasNeed(needs, "P") Or HasNeed(needs, "D") Then
        If Not IsDate(frm!txtFrom.Value) Or Not IsDate(frm!txtTo.Value) Then
            ShowWarning "أدخل تاريخ البداية وتاريخ النهاية."
            Exit Sub
        End If
        fromDate = DateValue(frm!txtFrom.Value)
        toDate = DateValue(frm!txtTo.Value)
        If toDate < fromDate Then
            ShowWarning "تاريخ النهاية يجب أن يكون بعد تاريخ البداية."
            Exit Sub
        End If
        SetPeriod fromDate, toDate
    End If
    If Not RequireChoice(frm!cboCustomer, HasNeed(needs, "C"), "CustomerID", "العميل") Then Exit Sub
    If Not RequireChoice(frm!cboSupplier, HasNeed(needs, "S"), "SupplierID", "المورد") Then Exit Sub
    If Not RequireChoice(frm!cboProduct, HasNeed(needs, "R"), "ProductID", "المنتج") Then Exit Sub

    If HasNeed(needs, "D") Then
        where = "[" & r(5) & "] >= " & SqlDate(fromDate) & " AND [" & r(5) & "] < " & _
                SqlDate(DateAdd("d", 1, toDate))
    End If
    If HasNeed(needs, "c") Then AddFilter where, frm!cboCustomer, "CustomerID"
    If HasNeed(needs, "s") Then AddFilter where, frm!cboSupplier, "SupplierID"
    If HasNeed(needs, "r") Then AddFilter where, frm!cboProduct, "ProductID"

    LogAction "REPORT", CStr(r(0))
    OpenReportOrQuery CStr(r(3)), CStr(r(2)), where
End Sub

Public Sub OpenReportOrQuery(ByVal ReportName As String, ByVal QueryName As String, _
                             ByVal WhereCondition As String)
    ' Formatted reports arrive in Phase 8; until then the query is shown as a table.
    Dim sql As String, db As DAO.Database
    If ReportExists(ReportName) Then
        DoCmd.OpenReport ReportName, acViewPreview, , WhereCondition
        Exit Sub
    End If
    sql = "SELECT * FROM [" & QueryName & "]"
    If Len(WhereCondition) > 0 Then sql = sql & " WHERE " & WhereCondition
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete "qryReportPreview"
    On Error GoTo 0
    db.CreateQueryDef "qryReportPreview", sql
    DoCmd.OpenQuery "qryReportPreview", acViewNormal, acReadOnly
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
        ShowWarning "اختر " & Caption & "."
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
