"""Generate modBuildReports.bas (report builder) from tools/reports.py."""

import re

import reports as RP
from gen_forms import lit
from generate_common import vba_str

TEMPLATE = r'''Attribute VB_Name = "modBuildReports"
'==============================================================================
' modBuildReports  -  Retail Store Management System (Phases 6 and 8: reports)
'
' GENERATED FILE - do not edit by hand.
' Source: tools/reports.py, tools/reports_docs.py, tools/reports_catalog.py
'
'   BuildReports   (re)creates every report:
'                  - sales invoice / credit note: rptSalesReceipt (80 mm) and
'                    rptSalesInvoiceA4, with the ZATCA QR code (DrawDocumentQR);
'                  - documents: purchase invoice / return, vouchers, stocktake sheet;
'                  - the report centre catalogue (one report per saved query).
'   TestReports    opens every report in preview (hidden) and checks it.
' Text functions used by the reports are in modReports.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MIRROR_LAYOUT As Boolean = False      ' same switch as modBuildForms
Private Const EP As String = "[Event Procedure]"

Private m_rpt As Access.Report
Private m_tmp As String
Private m_width As Long
Private m_built As Long
Private m_failed As Long
Private m_report As String
Private m_passed As Long
Private Const REPORT_NAMES As String = "@@REPORT_NAMES@@"

Public Function BuildReports() As Boolean
    Dim i As Long
    On Error GoTo EH
    m_built = 0: m_failed = 0: m_report = ""
    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name, acSaveNo
    Next
    DoCmd.Echo False, "جاري بناء التقارير..."
@@BUILD_ALL@@
    DoCmd.Echo True
    Application.RefreshDatabaseWindow
    If m_failed = 0 Then
        MsgBox "تم بناء التقارير بنجاح (" & m_built & ").", vbInformation + MSG_RTL, "BuildReports"
        BuildReports = True
    Else
        MsgBox "فشل بناء " & m_failed & " تقرير:" & vbCrLf & m_report, vbExclamation + MSG_RTL, "BuildReports"
    End If
    Exit Function
EH:
    DoCmd.Echo True
    MsgBox "خطأ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "BuildReports"
End Function

Public Function TestReports() As Boolean
    ' Opens every report hidden in preview (with this year's period and the first customer,
    ' supplier and product as parameters) and checks the amount-in-words function.
    Dim f As Variant
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestReports  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    SetPeriod DateSerial(Year(Date), 1, 1), Date
    SetQueryParam "CustomerID", Nz(DMin("CustomerID", "Customers"), 0)
    SetQueryParam "SupplierID", Nz(DMin("SupplierID", "Suppliers"), 0)
    SetQueryParam "ProductID", Nz(DMin("ProductID", "Products"), 0)
    For Each f In Split(REPORT_NAMES, ",")
        CheckReportOpens CStr(f)
    Next
@@WORDS@@
    TempVars.Add "ReportCriteria", ""
    g_SilentMode = False
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        TestMsg "جميع اختبارات التقارير ناجحة (" & m_passed & " اختبارًا).", vbInformation + MSG_RTL, "TestReports"
        TestReports = True
    Else
        TestMsg "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "TestReports"
    End If
End Function

Private Sub CheckReportOpens(ByVal ReportName As String)
    On Error GoTo EH
    If Not ReportExists(ReportName) Then
        RecordR False, "التقرير غير موجود: " & ReportName & " (شغّل BuildReports)"
        Exit Sub
    End If
    TempVars.Add "ReportCriteria", "TEST"
    DoCmd.OpenReport ReportName, acViewPreview, , , acHidden
    DoCmd.Close acReport, ReportName, acSaveNo
    RecordR True, "التقرير " & ReportName & " يفتح"
    Exit Sub
EH:
    If Err.Number = 2501 Then              ' cancelled by Report_NoData: no data yet, not an error
        RecordR True, "التقرير " & ReportName & " يفتح (لا توجد بيانات بعد)"
    Else
        RecordR False, ReportName & ": خطأ " & Err.Number & " - " & Err.Description
    End If
    On Error Resume Next
    DoCmd.Close acReport, ReportName, acSaveNo
End Sub

Private Sub RecordR(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        m_failed = m_failed + 1
        m_report = m_report & "- " & Label & vbCrLf
        Debug.Print "[X] " & Label
    End If
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Sub StartReport(ByVal FinalName As String, ByVal Caption As String, ByVal RecordSource As String, _
                        ByVal ReportWidth As Long, ByVal GroupField As String, ByVal SortFields As String, _
                        ByVal Landscape As Boolean, ByVal PageSetup As Boolean)
    ' GroupField: document reports group on it (header 5 = secHeader, footer 6 = secTotals).
    ' SortFields: "Field1,-Field2" (a leading "-" sorts descending). Reports ignore ORDER BY.
    Dim f As Variant, level As Long, fieldName As String
    If ReportExists(FinalName) Then DoCmd.DeleteObject acReport, FinalName
    Set m_rpt = CreateReport()
    m_tmp = m_rpt.Name
    m_width = ReportWidth
    SetRptProp "Orientation", 1                     ' right-to-left
    m_rpt.RecordSource = RecordSource
    m_rpt.Caption = Caption
    If Len(GroupField) > 0 Then
        CreateGroupLevel m_tmp, GroupField, True, True
        level = 1
    End If
    If Len(SortFields) > 0 Then
        For Each f In Split(SortFields, ",")
            fieldName = CStr(f)
            If Left$(fieldName, 1) = "-" Then fieldName = Mid$(fieldName, 2)
            CreateGroupLevel m_tmp, fieldName, False, False
            If Left$(CStr(f), 1) = "-" Then m_rpt.GroupLevel(level).SortOrder = True
            level = level + 1
        Next
    End If
    m_rpt.Width = ReportWidth
    If Len(GroupField) > 0 Then
        m_rpt.Section(5).Name = "secHeader"
        m_rpt.Section(6).Name = "secTotals"
        m_rpt.Section(6).KeepTogether = True
    End If
    If PageSetup Then SetPage Landscape
    m_rpt.HasModule = True
End Sub

Private Sub SetSection(ByVal SectionIndex As Integer, ByVal SectionHeight As Long)
    EnsureSection SectionIndex
    m_rpt.Section(SectionIndex).Height = SectionHeight
    m_rpt.Section(SectionIndex).Visible = True
End Sub

Private Sub EnsureSection(ByVal SectionIndex As Integer)
    ' Page and report header/footer sections come in pairs and are switched on with RunCommand.
    Dim h As Long, missing As Boolean
    On Error Resume Next
    h = m_rpt.Section(SectionIndex).Height
    missing = (Err.Number <> 0)
    On Error GoTo 0
    If Not missing Then Exit Sub
    DoCmd.SelectObject acReport, m_tmp
    If SectionIndex = acPageHeader Or SectionIndex = acPageFooter Then
        DoCmd.RunCommand acCmdPageHdrFtr
    ElseIf SectionIndex = acHeader Or SectionIndex = acFooter Then
        DoCmd.RunCommand acCmdReportHdrFtr
    End If
End Sub

Private Sub SetPage(ByVal Landscape As Boolean)
    ' A4 with 0.8 cm side margins. Needs a printer driver; without one the layout still works.
    On Error Resume Next
    m_rpt.Printer.PaperSize = 9                     ' A4
    If Landscape Then m_rpt.Printer.Orientation = 2 Else m_rpt.Printer.Orientation = 1
    m_rpt.Printer.LeftMargin = 454
    m_rpt.Printer.RightMargin = 454
    m_rpt.Printer.TopMargin = 567
    m_rpt.Printer.BottomMargin = 567
End Sub

Private Sub HideSection(ByVal SectionIndex As Integer)
    On Error Resume Next                             ' the section may not exist
    m_rpt.Section(SectionIndex).Height = 0
    m_rpt.Section(SectionIndex).Visible = False
End Sub

Private Sub SetSecProp(ByVal SectionIndex As Integer, ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    m_rpt.Section(SectionIndex).Properties(PropName).Value = Value
End Sub

Private Function NewRptCtl(ByVal CtlType As AcControlType, ByVal SectionIndex As Integer, _
                           ByVal CtlName As String, ByVal L As Long, ByVal T As Long, ByVal W As Long, _
                           ByVal H As Long, Optional ByVal ColumnName As String = "") As Access.Control
    Dim x As Long
    If MIRROR_LAYOUT Then x = m_width - L - W Else x = L
    Set NewRptCtl = CreateReportControl(m_tmp, CtlType, SectionIndex, "", ColumnName, x, T, W, H)
    NewRptCtl.Name = CtlName
End Function

Private Function RText(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal Source As String, _
                       ByVal L As Long, ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                       ByVal FontSize As Integer, ByVal Bold As Boolean, ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    If Left$(Source, 1) = "=" Then
        Set c = NewRptCtl(acTextBox, SectionIndex, CtlName, L, T, W, H)
        c.ControlSource = Source
    Else
        Set c = NewRptCtl(acTextBox, SectionIndex, CtlName, L, T, W, H, Source)
    End If
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.TextAlign = TextAlign
    c.BorderStyle = 0
    c.BackStyle = 0
    c.ForeColor = 0
    Set RText = c
End Function

Private Function RLabel(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal Caption As String, _
                        ByVal L As Long, ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                        ByVal FontSize As Integer, ByVal Bold As Boolean, ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    Set c = NewRptCtl(acLabel, SectionIndex, CtlName, L, T, W, H)
    c.Caption = Caption
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.TextAlign = TextAlign
    c.ForeColor = 0
    Set RLabel = c
End Function

Private Function RLine(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal T As Long, _
                       ByVal W As Long) As Access.Control
    Set RLine = NewRptCtl(acLine, SectionIndex, CtlName, 0, T, W, 0)
End Function

Private Function RBox(ByVal SectionIndex As Integer, ByVal CtlName As String, ByVal L As Long, _
                      ByVal T As Long, ByVal W As Long, ByVal H As Long) As Access.Control
    Dim c As Access.Control
    Set c = NewRptCtl(acRectangle, SectionIndex, CtlName, L, T, W, H)
    c.BorderStyle = 0
    c.BackStyle = 0
    Set RBox = c
End Function

Private Sub SetRptProp(ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    m_rpt.Properties(PropName).Value = Value
End Sub

Private Sub SetCtl(ByVal c As Access.Control, ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    c.Properties(PropName).Value = Value
End Sub

Private Sub FinishReport(ByVal FinalName As String, ByVal Code As String)
    Dim mdl As Access.Module, i As Long, hasExplicit As Boolean
    Set mdl = m_rpt.Module
    For i = 1 To mdl.CountOfDeclarationLines
        If Trim$(mdl.Lines(i, 1)) = "Option Explicit" Then hasExplicit = True
    Next
    If Not hasExplicit Then mdl.InsertLines mdl.CountOfDeclarationLines + 1, "Option Explicit"
    mdl.AddFromString Code
    DoCmd.Close acReport, m_tmp, acSaveYes
    DoCmd.Rename FinalName, acReport, m_tmp
    m_built = m_built + 1
    Debug.Print "  + " & FinalName
End Sub

Private Sub AbortReport(ByVal FinalName As String, ByVal ErrNumber As Long, ByVal ErrText As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & FinalName & ": " & ErrNumber & " " & ErrText & vbCrLf
    Debug.Print "[X] " & FinalName & ": " & ErrNumber & " " & ErrText
    On Error Resume Next
    DoCmd.Close acReport, m_tmp, acSaveNo
End Sub

'------------------------------------------------------------------------------
' Generated: one procedure per report
'------------------------------------------------------------------------------
@@REPORT_SUBS@@
'''


def control_lines(sec, c):
    p = c.props
    out = []
    if c.kind == "text":
        out.append(f"    Set c = RText({sec}, {vba_str(c.name)}, {vba_str(c.source)}, {c.x}, {c.y}, {c.w}, "
                   f"{c.h}, {p.get('FontSize', 9)}, {lit(bool(p.get('FontBold')))}, {p.get('TextAlign', 0)})")
        for key in ("Format", "CanGrow", "Visible", "RunningSum"):
            if key in p:
                out.append(f"    SetCtl c, {vba_str(key)}, {lit(p[key])}")
    elif c.kind == "label":
        out.append(f"    Set c = RLabel({sec}, {vba_str(c.name)}, {vba_str(p['Caption'])}, {c.x}, {c.y}, "
                   f"{c.w}, {c.h}, {p.get('FontSize', 9)}, {lit(bool(p.get('FontBold')))}, "
                   f"{p.get('TextAlign', 0)})")
    elif c.kind == "line":
        out.append(f"    Set c = RLine({sec}, {vba_str(c.name)}, {c.y}, {c.w})")
    elif c.kind == "rect":
        out.append(f"    Set c = RBox({sec}, {vba_str(c.name)}, {c.x}, {c.y}, {c.w}, {c.h})")
        for key in ("BorderStyle", "BackStyle", "BackColor"):
            if key in p:
                out.append(f"    SetCtl c, {vba_str(key)}, {lit(p[key])}")
    else:
        raise ValueError(c.kind)
    return out


def sort_spec(m: RP.ReportModel) -> str:
    return ",".join(("-" if desc else "") + f for f, desc in m.sorts)


SECTION_ORDER = [RP.SEC_PAGE_HEADER, RP.SEC_PAGE_FOOTER, RP.SEC_RPT_HEADER, RP.SEC_RPT_FOOTER, RP.SEC_DETAIL,
                 RP.SEC_HEADER, RP.SEC_FOOTER]


def report_sub(m: RP.ReportModel) -> str:
    lines = [f"Private Sub BuildReport_{m.name}()",
             "    Dim c As Access.Control, s As String",
             "    On Error GoTo EH",
             f"    StartReport {vba_str(m.name)}, {vba_str(m.caption)}, {vba_str(m.record_source)}, {m.width}, "
             f"{vba_str(m.group)}, {vba_str(sort_spec(m))}, {lit(m.landscape)}, {lit(m.page_setup)}"]
    for sec in SECTION_ORDER:
        if sec in m.heights:
            lines.append(f"    SetSection {sec}, {m.heights[sec]}")
    for sec in (RP.SEC_RPT_HEADER, RP.SEC_RPT_FOOTER, RP.SEC_PAGE_HEADER, RP.SEC_PAGE_FOOTER):
        if sec not in m.heights:
            lines.append(f"    HideSection {sec}")
    for sec in SECTION_ORDER:
        for c in m.controls.get(sec, []):
            lines += control_lines(sec, c)
    lines += [f"    {e}" for e in m.events]
    lines.append('    s = ""')
    for code_line in m.code:
        lines.append(f"    s = s & {vba_str(code_line)} & vbCrLf")
    lines += [f"    FinishReport {vba_str(m.name)}, s", "    Exit Sub", "EH:",
              f"    AbortReport {vba_str(m.name)}, Err.Number, Err.Description", "End Sub"]
    return "\n".join(lines)


def words_lines() -> str:
    import tafqeet as T
    out = []
    for amount in T.CASES:
        out.append(f"    RecordR AmountInWords({amount}) = {vba_str(T.amount_in_words(amount))}, "
                   f"{vba_str('المبلغ بالحروف: ' + str(amount))}")
    return "\n".join(out)


def build_reports_vba() -> str:
    models = RP.all_reports()
    text = TEMPLATE.replace("@@BUILD_ALL@@", "\n".join(f"    BuildReport_{m.name}" for m in models))
    text = text.replace("@@REPORT_NAMES@@", ",".join(m.name for m in models))
    text = text.replace("@@WORDS@@", words_lines())
    text = text.replace("@@REPORT_SUBS@@", "\n\n".join(report_sub(m) for m in models))
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text
