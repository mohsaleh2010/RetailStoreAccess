"""Generate modBuildReports.bas (report builder) from tools/reports.py."""

import re

import reports as RP
from gen_forms import lit
from generate_common import vba_str

TEMPLATE = r'''Attribute VB_Name = "modBuildReports"
'==============================================================================
' modBuildReports  -  Retail Store Management System (Phase 6: invoice reports)
'
' GENERATED FILE - do not edit by hand.  Source: tools/reports.py
'
'   BuildReports   (re)creates the invoice / credit-note reports:
'                  rptSalesReceipt (80 mm thermal) and rptSalesInvoiceA4.
' Both read qrySalesDocPrint, group on DocID and draw the ZATCA QR code in the
' totals section (DrawDocumentQR in modPOS -> DrawQR in modQRCode).
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

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Sub StartReport(ByVal FinalName As String, ByVal Caption As String, ByVal RecordSource As String, _
                        ByVal ReportWidth As Long, ByVal HeaderHeight As Long, ByVal DetailHeight As Long, _
                        ByVal FooterHeight As Long)
    If ReportExists(FinalName) Then DoCmd.DeleteObject acReport, FinalName
    Set m_rpt = CreateReport()
    m_tmp = m_rpt.Name
    m_width = ReportWidth
    SetRptProp "Orientation", 1                     ' right-to-left
    m_rpt.RecordSource = RecordSource
    m_rpt.Caption = Caption
    CreateGroupLevel m_tmp, "DocID", True, True     ' sections 5 (header) and 6 (footer)
    CreateGroupLevel m_tmp, "LineNumber", False, False
    m_rpt.Width = ReportWidth
    m_rpt.Section(acDetail).Height = DetailHeight
    m_rpt.Section(5).Height = HeaderHeight
    m_rpt.Section(6).Height = FooterHeight
    m_rpt.Section(5).Name = "secHeader"
    m_rpt.Section(6).Name = "secTotals"
    m_rpt.Section(6).KeepTogether = True
    HideSection acPageHeader
    HideSection acPageFooter
    HideSection acHeader
    HideSection acFooter
    m_rpt.HasModule = True
End Sub

Private Sub HideSection(ByVal SectionIndex As Integer)
    On Error Resume Next                             ' the section may not exist
    m_rpt.Section(SectionIndex).Height = 0
    m_rpt.Section(SectionIndex).Visible = False
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
        for key in ("Format", "CanGrow", "Visible"):
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


def report_sub(m: RP.ReportModel) -> str:
    lines = [f"Private Sub BuildReport_{m.name}()",
             "    Dim c As Access.Control, s As String",
             "    On Error GoTo EH",
             f"    StartReport {vba_str(m.name)}, {vba_str(m.caption)}, {vba_str(m.record_source)}, {m.width}, "
             f"{m.heights[RP.SEC_HEADER]}, {m.heights[RP.SEC_DETAIL]}, {m.heights[RP.SEC_FOOTER]}"]
    for sec in (RP.SEC_HEADER, RP.SEC_DETAIL, RP.SEC_FOOTER):
        for c in m.controls[sec]:
            lines += control_lines(sec, c)
    if m.code:
        lines.append("    m_rpt.Section(6).OnPrint = EP")
    lines.append('    s = ""')
    for code_line in m.code:
        lines.append(f"    s = s & {vba_str(code_line)} & vbCrLf")
    lines += [f"    FinishReport {vba_str(m.name)}, s", "    Exit Sub", "EH:",
              f"    AbortReport {vba_str(m.name)}, Err.Number, Err.Description", "End Sub"]
    return "\n".join(lines)


def build_reports_vba() -> str:
    models = RP.all_reports()
    text = TEMPLATE.replace("@@BUILD_ALL@@", "\n".join(f"    BuildReport_{m.name}" for m in models))
    text = text.replace("@@REPORT_SUBS@@", "\n\n".join(report_sub(m) for m in models))
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text
