"""Generate modBuildForms.bas (form builder + tests) and modAppData.bas (runtime data)."""

import re

import forms as F
from generate_common import vba_str

EVENT_PROPERTY = {
    "Load": "OnLoad", "Open": "OnOpen", "Current": "OnCurrent", "BeforeUpdate": "BeforeUpdate",
    "AfterUpdate": "AfterUpdate", "Error": "OnError", "KeyDown": "OnKeyDown",
    "Unload": "OnUnload", "Click": "OnClick", "DblClick": "OnDblClick", "Change": "OnChange",
    "Activate": "OnActivate", "Timer": "OnTimer", "Resize": "OnResize",
}

HELPER_PROPS = {"Caption", "FontSize", "FontBold", "ForeColor", "BackColor", "TextAlign",
                "RowSource", "ColumnCount", "ColumnWidths", "ColumnHeads", "Style",
                "RowSourceType", "SourceObject"}


def lit(v) -> str:
    if isinstance(v, F.Sym):
        return v.code
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        return repr(v)
    return vba_str(v)


def twips_widths(widths_cm: str) -> str:
    return ";".join(str(F.cm(float(w))) for w in widths_cm.split(";"))


def chunks(text: str, size: int = 150):
    """Split text at spaces into pieces of about `size` characters."""
    out, cur = [], ""
    for word in text.split(" "):
        if cur and len(cur) + len(word) + 1 > size:
            out.append(cur + " ")
            cur = word
        else:
            cur = f"{cur} {word}" if cur else word
    out.append(cur)
    return out


def assign_long_string(var: str, text: str, indent="    "):
    lines = []
    for i, part in enumerate(chunks(text)):
        lines.append(f"{indent}{var} = {'' if i == 0 else var + ' & '}{vba_str(part)}")
    return lines


def control_lines(c: F.Control):
    p = c.props
    color = lit(p.get("ForeColor", F.Sym("CLR_TEXT")))
    out = []
    if c.kind == "rect":
        out.append(f"    Set c = AddRect({vba_str(c.name)}, {c.x}, {c.y}, {c.w}, {c.h}, {lit(p['BackColor'])})")
    elif c.kind in ("label", "icon"):
        func = "AddLabel" if c.kind == "label" else "AddIcon"
        out.append(f"    Set c = {func}({vba_str(c.name)}, {lit(p.get('Caption', ' '))}, {c.x}, {c.y}, "
                   f"{c.w}, {c.h}, {p.get('FontSize', 10)}, {lit(bool(p.get('FontBold', False)))}, "
                   f"{color}, {vba_str(c.parent)}, {p.get('TextAlign', 0)})")
    elif c.kind == "text":
        out.append(f"    Set c = AddText({vba_str(c.name)}, {vba_str(c.source)}, {c.x}, {c.y}, {c.w}, {c.h})")
    elif c.kind == "combo":
        rows = p["RowSource"]
        out.append(f"    Set c = AddCombo({vba_str(c.name)}, {vba_str(c.source)}, {c.x}, {c.y}, {c.w}, "
                   f"{c.h}, {vba_str(rows)}, {p['ColumnCount']}, "
                   f"{vba_str(twips_widths(p['ColumnWidths']))})")
    elif c.kind == "check":
        out.append(f"    Set c = AddCheck({vba_str(c.name)}, {vba_str(c.source)}, {c.x}, {c.y})")
    elif c.kind == "list":
        widths = twips_widths(p["ColumnWidths"]) if "ColumnWidths" in p else ""
        out.append(f"    Set c = AddList({vba_str(c.name)}, {c.x}, {c.y}, {c.w}, {c.h}, "
                   f"{p.get('ColumnCount', 1)}, {vba_str(widths)}, {lit(bool(p.get('ColumnHeads')))})")
        if p.get("RowSourceType") == "Value List":
            out.append('    c.RowSourceType = "Value List"')
        if "FontSize" in p:
            out.append(f"    c.FontSize = {p['FontSize']}")
    elif c.kind == "subform":
        out.append(f"    Set c = AddSubform({vba_str(c.name)}, {vba_str(p['SourceObject'])}, {c.x}, {c.y}, "
                   f"{c.w}, {c.h})")
    elif c.kind == "button":
        out.append(f"    Set c = AddButton({vba_str(c.name)}, {vba_str(p['Caption'])}, {c.x}, {c.y}, "
                   f"{c.w}, {c.h}, {vba_str(p['Style'])})")
    else:
        raise ValueError(c.kind)
    if c.kind in ("text", "combo", "button") and "FontSize" in p:
        out.append(f"    c.FontSize = {p['FontSize']}")
    if c.kind in ("text", "combo") and p.get("FontBold"):
        out.append("    c.FontBold = True")
    for key, value in p.items():
        if key in HELPER_PROPS:
            continue
        out.append(f"    SetCtlProp c, {vba_str(key)}, {lit(value)}")
        if key == "Locked" and value:
            out.append("    c.BackColor = CLR_LOCKED")
    for ev in c.events:
        out.append(f'    c.{EVENT_PROPERTY[ev]} = EP')
    return out


def fit_spec(m: F.FormModel) -> list:
    by_name = {c.name: c for c in m.controls}
    return [f"{n},{by_name[n].x},{by_name[n].y},{by_name[n].w},{by_name[n].h},{mx},{mw},{my},{mh}"
            for n, (mx, mw, my, mh) in m.fit.items()]


def fit_min_dh(m: F.FormModel) -> int:
    """How far below its design height the window may squeeze the screen (<= 0): the stretched
    lists keep 2.5 cm, and controls following the bottom edge must not reach the ones above."""
    limit = None
    moving = {n for n, (mx, mw, my, mh) in m.fit.items() if my and not mh}
    fixed = [c for c in m.controls if c.name not in m.fit or not (m.fit[c.name][2] or m.fit[c.name][3])]
    for c in m.controls:
        f = m.fit.get(c.name)
        if f and f[3]:
            room = c.h - F.cm(2.5)
            limit = room if limit is None else min(limit, room)
        if c.name in moving:
            for t in fixed:
                if t.y + t.h <= c.y:
                    gap = c.y - (t.y + t.h)
                    limit = gap if limit is None else min(limit, gap)
    # controls that stay where they are must still fit in the window
    room = m.height - max(t.y + t.h for t in fixed)
    limit = room if limit is None else min(limit, room)
    return -max(0, limit - F.cm(0.1))


def fit_code_lines(m: F.FormModel) -> list:
    """Form_Resize: lists and grids grow with the window, buttons follow its edges."""
    items = fit_spec(m)
    code = ["Private Sub Form_Resize()", "    Dim spec As String"]
    for i in range(0, len(items), 6):
        chunk = ";".join(items[i:i + 6])
        code.append(f'    spec = "{chunk}"' if i == 0 else f'    spec = spec & ";{chunk}"')
    lines = [f"    s = s & {vba_str(line)} & vbCrLf" for line in code]
    lines.append(f'    s = s & "    FitControls Me, {m.width}, {m.height}, {fit_min_dh(m)}, " & '
                 'IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf')
    lines.append('    s = s & "End Sub" & vbCrLf')
    return lines


def form_sub(m: F.FormModel):
    lines = [f"Private Sub BuildForm_{m.name}()",
             "    Dim c As Access.Control, s As String",
             "    On Error GoTo EH",
             f"    StartForm {vba_str(m.name)}, {vba_str(m.caption)}, {vba_str(m.record_source)}, "
             f"{m.width}, {m.height}, {lit(m.popup)}, {lit(m.allow_add)}, {lit(m.allow_edit)}, _",
             f"              {vba_str(m.tag)}"]
    for key, value in m.form_props.items():
        lines.append(f"    SetFormProp {vba_str(key)}, {lit(value)}")
    for c in m.controls:
        lines += control_lines(c)
    for ev in m.form_events + (["Resize"] if m.fit else []):
        lines.append(f"    m_frm.{EVENT_PROPERTY[ev]} = EP")
    lines.append('    s = ""')
    for code_line in m.code:
        lines.append(f"    s = s & {vba_str(code_line)} & vbCrLf")
    if m.fit:
        lines += fit_code_lines(m)
    lines += [f"    FinishForm {vba_str(m.name)}, s", "    Exit Sub", "EH:",
              f"    AbortForm {vba_str(m.name)}, Err.Number, Err.Description", "End Sub"]
    return "\n".join(lines)


BUILDER_TEMPLATE = r'''Attribute VB_Name = "modBuildForms"
'==============================================================================
' modBuildForms  -  Retail Store Management System (Phase 5: Forms)
'
' GENERATED FILE - do not edit by hand.
' Source of truth: tools/forms.py  ->  python3 tools/generate.py
'
' Public procedures (run from the Immediate window, Ctrl+G):
'   BuildForms     (re)creates every screen. Existing screens with the same
'                  names are replaced - do not edit the generated screens by hand.
'   TestForms      opens every screen and runs the screen tests through the
'                  screens themselves; test records are removed afterwards.
'
' Requires: modCommon, modStartup, modForms, modScreens, modAppData,
'           modQueryParams, and Phases 2-4 (tables, relationships, queries).
'
' If the screens appear mirrored (labels on the wrong side of the inputs),
' set MIRROR_LAYOUT = True below and run BuildForms again.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MIRROR_LAYOUT As Boolean = False
Private Const EP As String = "[Event Procedure]"
Private Const FORM_NAMES As String = "@@FORM_NAMES@@"

Private m_frm As Access.Form
Private m_tmpName As String
Private m_width As Long
Private m_built As Long
Private m_failed As Long
Private m_passed As Long
Private m_report As String
Private m_warnings As String

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildForms() As Boolean
    On Error GoTo EH
    m_built = 0: m_failed = 0: m_report = "": m_warnings = ""
    Debug.Print "=== BuildForms  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CloseAllForms
    EnsureLocalTables                      ' tmp* working tables of the sales and purchase screens (modPOS)
    DoCmd.Echo False, "جاري بناء الشاشات..."
    BuildAllForms
    DoCmd.Echo True
    Application.RefreshDatabaseWindow
    Debug.Print "--- تم بناء: " & m_built & " | فشل: " & m_failed
    If Len(m_warnings) > 0 Then Debug.Print "تنبيهات:" & vbCrLf & m_warnings

    If m_failed = 0 Then
        MsgBox "تم بناء الشاشات بنجاح (" & m_built & " شاشة)." & vbCrLf & vbCrLf & _
               "الخطوة التالية: شغّل TestForms ثم افتح frmMain", vbInformation + MSG_RTL, "BuildForms"
        BuildForms = True
    Else
        MsgBox "فشل بناء " & m_failed & " شاشة:" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "BuildForms"
    End If
    Exit Function
EH:
    DoCmd.Echo True
    MsgBox "خطأ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "BuildForms"
End Function

Public Function TestForms() As Boolean
    Dim f As Variant, lastLog As Long
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestForms  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Calendar = vbCalGreg
    CloseAllForms
    lastLog = Nz(DMax("LogID", "AuditLog"), 0)
    EnsureTestUser                         ' run as the administrator when nobody is logged in
    g_SilentMode = True
    g_AutoAnswer = True

    For Each f In Split(FORM_NAMES, ",")
        CheckFormOpens CStr(f)
    Next
    TestHelpers
    TestProductScreen
    TestCustomerScreen
    TestExpenseScreen
    TestSearch
    TestReportCenter

    CleanUpTestData lastLog
    g_SilentMode = False
    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed
    If m_failed = 0 Then
        TestMsg "جميع اختبارات الشاشات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "تم حذف بيانات الاختبار.", vbInformation + MSG_RTL, "TestForms"
        TestForms = True
    Else
        TestMsg "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestForms"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    CloseAllForms
    CleanUpTestData lastLog
    g_SilentMode = False
    TestMsg errText, vbCritical + MSG_RTL, "TestForms"
End Function

'------------------------------------------------------------------------------
' Form building helpers
'------------------------------------------------------------------------------
Private Sub StartForm(ByVal FinalName As String, ByVal Caption As String, ByVal RecordSource As String, _
                      ByVal FormWidth As Long, ByVal FormHeight As Long, ByVal IsPopup As Boolean, _
                      ByVal AllowAdd As Boolean, ByVal AllowEdit As Boolean, ByVal TagText As String)
    If FormExists(FinalName) Then DoCmd.DeleteObject acForm, FinalName
    Set m_frm = CreateForm()
    m_tmpName = m_frm.Name
    m_width = FormWidth
    SetFormProp "Orientation", 1                 ' right-to-left
    m_frm.Caption = Caption
    m_frm.RecordSource = RecordSource
    m_frm.DefaultView = 0                        ' single form
    m_frm.RecordSelectors = False
    m_frm.NavigationButtons = False
    m_frm.DividingLines = False
    m_frm.ScrollBars = 0
    m_frm.AutoCenter = True
    m_frm.AutoResize = True
    m_frm.PopUp = IsPopup
    m_frm.Modal = IsPopup
    m_frm.BorderStyle = IIf(IsPopup, 3, 2)       ' dialog / sizable
    m_frm.ShortcutMenu = False
    m_frm.KeyPreview = True
    m_frm.Cycle = 1                              ' Tab stays on the current record
    m_frm.AllowAdditions = AllowAdd
    m_frm.AllowEdits = AllowEdit
    m_frm.AllowDeletions = False                 ' deleting goes through the Delete button
    m_frm.Tag = TagText
    SetFormProp "AllowDatasheetView", False
    SetFormProp "AllowLayoutView", False
    m_frm.Width = FormWidth
    m_frm.Section(acDetail).Height = FormHeight
    m_frm.Section(acDetail).BackColor = CLR_BACKGROUND
    m_frm.HasModule = True
End Sub

Private Sub FinishForm(ByVal FinalName As String, ByVal Code As String)
    Dim mdl As Access.Module, i As Long, hasExplicit As Boolean
    Set mdl = m_frm.Module
    For i = 1 To mdl.CountOfDeclarationLines
        If Trim$(mdl.Lines(i, 1)) = "Option Explicit" Then hasExplicit = True
    Next
    If Not hasExplicit Then mdl.InsertLines mdl.CountOfDeclarationLines + 1, "Option Explicit"
    mdl.AddFromString Code
    DoCmd.Close acForm, m_tmpName, acSaveYes
    DoCmd.Rename FinalName, acForm, m_tmpName
    m_built = m_built + 1
    Debug.Print "  + " & FinalName
End Sub

Private Sub AbortForm(ByVal FinalName As String, ByVal ErrNumber As Long, ByVal ErrText As String)
    Fail FinalName & ": خطأ " & ErrNumber & " - " & ErrText
    On Error Resume Next
    DoCmd.Close acForm, m_tmpName, acSaveNo
End Sub

Private Function NewCtl(ByVal CtlType As AcControlType, ByVal CtlName As String, ByVal L As Long, _
                     ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                     Optional ByVal ParentName As String = "", _
                     Optional ByVal ColumnName As String = "") As Access.Control
    Dim x As Long
    If MIRROR_LAYOUT Then x = m_width - L - W Else x = L
    Set NewCtl = CreateControl(m_tmpName, CtlType, acDetail, ParentName, ColumnName, x, T, W, H)
    NewCtl.Name = CtlName
End Function

Private Function AddRect(ByVal CtlName As String, ByVal L As Long, ByVal T As Long, ByVal W As Long, _
                         ByVal H As Long, ByVal Color As Long) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acRectangle, CtlName, L, T, W, H)
    c.BackStyle = 1
    c.BackColor = Color
    c.BorderStyle = 0
    c.SpecialEffect = 0
    Set AddRect = c
End Function

Private Function AddLabel(ByVal CtlName As String, ByVal Caption As String, ByVal L As Long, _
                          ByVal T As Long, ByVal W As Long, ByVal H As Long, ByVal FontSize As Integer, _
                          ByVal Bold As Boolean, ByVal Color As Long, ByVal ParentName As String, _
                          ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acLabel, CtlName, L, T, W, H, ParentName)
    If Len(Caption) = 0 Then Caption = " "          ' empty labels are deleted by Access
    c.Caption = Caption
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.ForeColor = Color
    c.BackStyle = 0
    c.TextAlign = TextAlign
    Set AddLabel = c
End Function

Private Function AddIcon(ByVal CtlName As String, ByVal Glyph As String, ByVal L As Long, _
                         ByVal T As Long, ByVal W As Long, ByVal H As Long, ByVal FontSize As Integer, _
                         ByVal Bold As Boolean, ByVal Color As Long, ByVal ParentName As String, _
                         ByVal TextAlign As Integer) As Access.Control
    Dim c As Access.Control
    Set c = AddLabel(CtlName, Glyph, L, T, W, H, FontSize, Bold, Color, ParentName, 2)
    c.FontName = ICON_FONT
    Set AddIcon = c
End Function

Private Sub StyleInput(ByVal c As Access.Control)
    c.FontName = FONT_NAME
    c.FontSize = 11
    c.ForeColor = CLR_TEXT
    c.BackColor = CLR_SURFACE
    c.BorderStyle = 1
    c.BorderColor = CLR_BORDER
    c.SpecialEffect = 0
End Sub

Private Function AddText(ByVal CtlName As String, ByVal Source As String, ByVal L As Long, _
                         ByVal T As Long, ByVal W As Long, ByVal H As Long) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acTextBox, CtlName, L, T, W, H, "", Source)
    StyleInput c
    Set AddText = c
End Function

Private Function AddCombo(ByVal CtlName As String, ByVal Source As String, ByVal L As Long, _
                          ByVal T As Long, ByVal W As Long, ByVal H As Long, ByVal Rows As String, _
                          ByVal ColumnCount As Integer, ByVal ColumnWidths As String) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acComboBox, CtlName, L, T, W, H, "", Source)
    StyleInput c
    If UCase$(Left$(Rows, 6)) = "SELECT" Then
        c.RowSourceType = "Table/Query"
    Else
        c.RowSourceType = "Value List"
    End If
    c.RowSource = Rows
    c.ColumnCount = ColumnCount
    c.ColumnWidths = ColumnWidths
    c.BoundColumn = 1
    c.LimitToList = True
    c.ListRows = 12
    SetCtlProp c, "AllowValueListEdits", False
    Set AddCombo = c
End Function

Private Function AddCheck(ByVal CtlName As String, ByVal Source As String, ByVal L As Long, _
                          ByVal T As Long) As Access.Control
    Set AddCheck = NewCtl(acCheckBox, CtlName, L, T, 284, 284, "", Source)
End Function

Private Function AddList(ByVal CtlName As String, ByVal L As Long, ByVal T As Long, ByVal W As Long, _
                         ByVal H As Long, ByVal ColumnCount As Integer, ByVal ColumnWidths As String, _
                         ByVal ColumnHeads As Boolean) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acListBox, CtlName, L, T, W, H)
    StyleInput c
    c.FontSize = 10
    c.RowSourceType = "Table/Query"
    c.ColumnCount = ColumnCount
    If Len(ColumnWidths) > 0 Then c.ColumnWidths = ColumnWidths
    c.ColumnHeads = ColumnHeads
    c.BoundColumn = 1
    Set AddList = c
End Function

Private Function AddSubform(ByVal CtlName As String, ByVal SourceObject As String, ByVal L As Long, _
                            ByVal T As Long, ByVal W As Long, ByVal H As Long) As Access.Control
    Dim c As Access.Control
    Set c = NewCtl(acSubform, CtlName, L, T, W, H)
    c.SourceObject = SourceObject
    c.BorderStyle = 1
    c.BorderColor = CLR_BORDER
    c.SpecialEffect = 0
    Set AddSubform = c
End Function

Private Function AddButton(ByVal CtlName As String, ByVal Caption As String, ByVal L As Long, _
                           ByVal T As Long, ByVal W As Long, ByVal H As Long, _
                           ByVal Style As String) As Access.Control
    Dim c As Access.Control, back As Long, hover As Long, fore As Long
    Set c = NewCtl(acCommandButton, CtlName, L, T, W, H)
    c.Caption = Caption
    c.FontName = FONT_NAME
    c.FontSize = IIf(Style = "nav", 12, 11)
    c.FontBold = True
    Select Case Style
        Case "primary": back = CLR_ACCENT: hover = CLR_ACCENT_HOVER: fore = CLR_SURFACE
        Case "danger":  back = CLR_DANGER: hover = CLR_DANGER_HOVER: fore = CLR_SURFACE
        Case "nav":     back = CLR_PRIMARY: hover = CLR_PRIMARY_HOVER: fore = CLR_SIDEBAR_TEXT
        Case Else:      back = CLR_SECONDARY: hover = CLR_SECONDARY_HOVER: fore = CLR_TEXT
    End Select
    ' Modern button colours (Access 2010+); older versions keep the system look.
    SetCtlProp c, "UseTheme", True
    SetCtlProp c, "BackColor", back
    SetCtlProp c, "HoverColor", hover
    SetCtlProp c, "PressedColor", CLR_PRIMARY_PRESSED
    SetCtlProp c, "ForeColor", fore
    SetCtlProp c, "HoverForeColor", IIf(Style = "secondary", CLR_TEXT, CLR_SURFACE)
    SetCtlProp c, "PressedForeColor", CLR_SURFACE
    SetCtlProp c, "BorderStyle", 0
    SetCtlProp c, "CursorOnHover", 1
    Set AddButton = c
End Function

Private Sub SetCtlProp(ByVal c As Access.Control, ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    c.Properties(PropName).Value = Value
    If Err.Number <> 0 Then
        m_warnings = m_warnings & "  " & m_tmpName & "." & c.Name & "." & PropName & _
                     ": " & Err.Description & vbCrLf
    End If
End Sub

Private Sub SetFormProp(ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    m_frm.Properties(PropName).Value = Value
    If Err.Number <> 0 Then m_warnings = m_warnings & "  form." & PropName & ": " & Err.Description & vbCrLf
End Sub

Private Sub CloseAllForms()
    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next
End Sub

'------------------------------------------------------------------------------
' Tests
'------------------------------------------------------------------------------
Private Sub CheckFormOpens(ByVal FormName As String)
    Dim frm As Access.Form, ctl As Access.Control, rs As DAO.Recordset
    On Error GoTo EH
    DoCmd.OpenForm FormName, acNormal, , , , acHidden
    Set frm = Forms(FormName)
    For Each ctl In frm.Controls
        If ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
            If ctl.RowSourceType = "Table/Query" And Len(ctl.RowSource) > 0 Then
                Set rs = CurrentDb.OpenRecordset(ctl.RowSource, dbOpenSnapshot)
                rs.Close
            End If
        End If
    Next
    DoCmd.Close acForm, FormName, acSaveNo
    Record True, "الشاشة " & FormName & " تفتح وكل قوائمها تعمل"
    Exit Sub
EH:
    Record False, "الشاشة " & FormName & ": خطأ " & Err.Number & " - " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, FormName, acSaveNo
End Sub

Private Sub TestHelpers()
    Call Record(RoundMoney(2.345) = 2.35, "تقريب 2.345 = 2.35 (وليس 2.34 كما في Round)")
    Call Record(RoundMoney(2.344) = 2.34, "تقريب 2.344 = 2.34")
    Call Record(RoundMoney(-2.345) = -2.35, "تقريب -2.345 = -2.35")
    Call Record(RoundMoney(149.999, 2) = 150, "تقريب 149.999 = 150")
    Call Record(LikePattern("a'b*") = "'*a''b[*]*'", "تهريب نص البحث")
    Call Record(SqlDate(DateSerial(2026, 3, 15) + TimeSerial(14, 5, 9)) = "#2026-03-15 14:05:09#", _
                "صيغة التاريخ في SQL")
End Sub

Private Sub TestProductScreen()
    Dim frm As Access.Form, code As String, seqBefore As Long
    seqBefore = DLookup("NextValue", "Sequences", "SequenceName = 'PRODUCT_CODE'")
    DoCmd.OpenForm "frmProducts", acNormal, , , , acHidden
    Set frm = Forms("frmProducts")

    FormAction frm, "NEW"
    frm!ProductName.Value = "TEST-UI منتج"
    frm!Barcode.Value = "TESTUI0001"
    frm!PurchasePrice.Value = 7
    frm!SellingPrice.Value = 11.5
    Call Record(SaveRecord(frm), "حفظ منتج جديد من شاشة المنتجات")
    code = Nz(DLookup("ProductCode", "Products", "Barcode = 'TESTUI0001'"), "")
    Call Record(code Like "P#####", "توليد كود المنتج تلقائيًا (" & code & ")")
    Call Record(Nz(DLookup("AverageCost", "Products", "Barcode = 'TESTUI0001'"), -1) = 7, _
                "متوسط التكلفة لمنتج جديد = سعر الشراء")
    Call Record(Nz(DLookup("CurrentQuantity", "Products", "Barcode = 'TESTUI0001'"), -1) = 0, _
                "الكمية الابتدائية = 0")

    FormAction frm, "NEW"
    frm!ProductName.Value = "TEST-UI منتج 2"
    frm!Barcode.Value = "TESTUI0001"
    Call Record(Not SaveRecord(frm), "رفض باركود مكرر برسالة واضحة")
    frm.Undo

    FormAction frm, "NEW"
    frm!Barcode.Value = "TESTUI0002"
    Call Record(Not SaveRecord(frm), "رفض منتج بدون اسم")
    frm.Undo

    frm!txtSearch.Value = "TESTUI0001"
    RefreshList frm
    Call Record(frm!lstItems.ListCount = 2, "البحث في قائمة المنتجات بالباركود")
    DoCmd.Close acForm, "frmProducts", acSaveNo
    CurrentDb.Execute "UPDATE [Sequences] SET [NextValue] = " & seqBefore & _
                      " WHERE [SequenceName] = 'PRODUCT_CODE'", dbFailOnError
End Sub

Private Sub TestCustomerScreen()
    Dim frm As Access.Form, id As Variant
    DoCmd.OpenForm "frmCustomers", acNormal, , , , acHidden
    Set frm = Forms("frmCustomers")

    FormAction frm, "NEW"
    frm!CustomerName.Value = "TEST-UI عميل"
    frm!Mobile.Value = "0500000000"
    frm!Email.Value = "abc"
    Call Record(Not SaveRecord(frm), "رفض بريد إلكتروني غير صحيح")
    frm!Email.Value = "test@example.com"
    frm!OpeningBalance.Value = 75
    Call Record(SaveRecord(frm), "حفظ عميل برصيد افتتاحي")
    id = DLookup("CustomerID", "Customers", "CustomerName = 'TEST-UI عميل'")
    Call Record(Nz(DLookup("CurrentBalance", "Customers", "CustomerID = " & Nz(id, 0)), -1) = 75, _
                "الرصيد الحالي = الرصيد الافتتاحي")
    Call Record(DCount("*", "IntegrityCheckQuery", "IssueCode = 'CUSTOMER_BALANCE'") = 0, _
                "رصيد العميل مطابق لفحص السلامة")

    DoCmd.Close acForm, "frmCustomers", acSaveNo
    DoCmd.OpenForm "frmCustomers", acNormal, , , , acHidden, 1
    Set frm = Forms("frmCustomers")
    Call Record(frm!CustomerID.Value = 1, "فتح الشاشة على سجل محدد (العميل النقدي)")
    FormAction frm, "DELETE"
    Call Record(DCount("*", "Customers", "CustomerID = 1") = 1, "منع حذف العميل النقدي")
    DoCmd.Close acForm, "frmCustomers", acSaveNo
End Sub

Private Sub TestExpenseScreen()
    Dim frm As Access.Form, seqBefore As Long, num As String
    seqBefore = DLookup("NextValue", "Sequences", "SequenceName = 'EXPENSE'")
    DoCmd.OpenForm "frmExpenses", acNormal, , , , acHidden
    Set frm = Forms("frmExpenses")
    FormAction frm, "NEW"
    frm!ExpenseTypeID.Value = 2
    frm!Amount.Value = 100
    frm!Description.Value = "TEST-UI مصروف"
    CalcExpenseVat frm
    Call Record(frm!Tax.Value = 15 And frm!TotalAmount.Value = 115, "حساب ضريبة المصروف 15% والإجمالي")
    Call Record(SaveRecord(frm), "حفظ مصروف")
    num = Nz(DLookup("ExpenseNumber", "Expenses", "Description = 'TEST-UI مصروف'"), "")
    Call Record(num Like "EXP-######", "ترقيم المصروف تلقائيًا (" & num & ")")
    Call Record(Nz(DLookup("EmployeeID", "Expenses", "Description = 'TEST-UI مصروف'"), 0) = CurrentUserID(), _
                "تسجيل الموظف الحالي على المصروف")
    DoCmd.Close acForm, "frmExpenses", acSaveNo
    CurrentDb.Execute "UPDATE [Sequences] SET [NextValue] = " & seqBefore & _
                      " WHERE [SequenceName] = 'EXPENSE'", dbFailOnError
End Sub

Private Sub TestSearch()
    Dim frm As Access.Form, kind As Variant
    DoCmd.OpenForm "frmSearch", acNormal, , , , acHidden
    Set frm = Forms("frmSearch")
    For Each kind In Array("PRODUCT", "CUSTOMER", "SUPPLIER", "SALE", "PURCHASE")
        frm!cboKind.Value = kind
        SearchKindChanged frm
        frm!txtText.Value = Null
        RunSearch frm
        Call Record(Left$(frm!lstResults.RowSource, 6) = "SELECT", "البحث في " & kind & " يعمل")
    Next
    frm!cboKind.Value = "PRODUCT"
    SearchKindChanged frm
    frm!txtText.Value = "TESTUI0001"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "البحث المتقدم بالباركود يجد المنتج")
    frm!txtText.Value = "TEST-UI"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "البحث المتقدم بجزء من الاسم")
    frm!cboKind.Value = "CUSTOMER"
    SearchKindChanged frm
    frm!txtText.Value = "0500000000"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "البحث عن عميل برقم الجوال")
    DoCmd.Close acForm, "frmSearch", acSaveNo
End Sub

Private Sub TestReportCenter()
    Dim frm As Access.Form, i As Long, r As Variant, missing As String
    For i = 1 To ReportCount()
        r = ReportRow(i)
        If Not QueryExists(CStr(r(2))) Then missing = missing & " " & r(2)
    Next
    Call Record(Len(missing) = 0, "كل استعلامات مركز التقارير موجودة" & missing)
    DoCmd.OpenForm "frmReportCenter", acNormal, , , , acHidden
    Set frm = Forms("frmReportCenter")
    Call Record(frm!lstReports.ListCount = ReportCount(), "قائمة التقارير (" & ReportCount() & ")")
    frm!lstReports.Value = "CUSTOMER_STATEMENT"
    ReportSelected frm
    Call Record(frm!cboCustomer.Enabled And Not frm!cboSupplier.Enabled And frm!txtFrom.Enabled, _
                "كشف حساب عميل يطلب الفترة والعميل فقط")
    frm!lstReports.Value = "STOCK"
    ReportSelected frm
    Call Record(Not frm!txtFrom.Enabled And Not frm!cboCustomer.Enabled, _
                "تقرير المخزون لا يطلب فترة")
    DoCmd.Close acForm, "frmReportCenter", acSaveNo
End Sub

Private Sub CleanUpTestData(ByVal LastLogID As Long)
    On Error Resume Next
    CloseAllForms
    CurrentDb.Execute "DELETE FROM [Products] WHERE [Barcode] Like 'TESTUI*'"
    CurrentDb.Execute "DELETE FROM [Customers] WHERE [CustomerName] = 'TEST-UI عميل'"
    CurrentDb.Execute "DELETE FROM [Expenses] WHERE [Description] = 'TEST-UI مصروف'"
    CurrentDb.Execute "DELETE FROM [AuditLog] WHERE [LogID] > " & LastLogID
End Sub

Private Function QueryExists(ByVal QueryName As String) As Boolean
    Dim qdf As DAO.QueryDef
    For Each qdf In CurrentDb.QueryDefs
        If StrComp(qdf.Name, QueryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next
End Function

Private Sub Record(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        Fail Label
    End If
End Sub

Private Sub Fail(ByVal Msg As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & Msg & vbCrLf
    Debug.Print "[X] " & Msg
End Sub

'------------------------------------------------------------------------------
' Generated: one procedure per screen
'------------------------------------------------------------------------------
Private Sub BuildAllForms()
@@BUILD_ALL@@
End Sub

@@FORM_SUBS@@
'''


def build_forms_vba() -> str:
    models = F.all_forms()
    text = BUILDER_TEMPLATE
    for key, value in {
        "@@FORM_NAMES@@": ",".join(m.name for m in models),
        "@@BUILD_ALL@@": "\n".join(f"    BuildForm_{m.name}" for m in models),
        "@@FORM_SUBS@@": "\n\n".join(form_sub(m) for m in models),
    }.items():
        text = text.replace(key, value)
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text


APPDATA_TEMPLATE = r'''Attribute VB_Name = "modAppData"
'==============================================================================
' modAppData  -  Retail Store Management System (Phase 5)
'
' GENERATED FILE - do not edit by hand.  Source: tools/forms.py
' Runtime data used by modScreens: search SQL templates and the report catalogue.
' Search tokens: {LIKE} quoted Like pattern, {NUM} record number or -1,
'                {FROM} / {TO} date literals (TO is exclusive).
'==============================================================================
Option Compare Database
Option Explicit

Public Function SearchTemplate(ByVal Kind As String) As String
    Dim s As String
    Select Case Kind
@@SEARCH_CASES@@
    End Select
    SearchTemplate = s
End Function

Public Function SearchColumnCount(ByVal Kind As String) As Integer
    Select Case Kind
@@COUNT_CASES@@
    End Select
End Function

Public Function SearchColumnWidths(ByVal Kind As String) As String
    Select Case Kind
@@WIDTH_CASES@@
    End Select
End Function

Public Function ReportCount() As Long
    ReportCount = @@REPORT_COUNT@@
End Function

Public Function ScreenPermission(ByVal FormName As String) As String
    ' Permission needed to open a screen ("" = any logged-in user). Source: forms.SCREEN_PERMISSIONS
    Select Case FormName
@@SCREEN_CASES@@
    End Select
End Function

Public Function ReportRow(ByVal Index As Long) As Variant
    ' Array(Key, Title, QueryName, ReportName, Needs, DateColumn)
    Select Case Index
@@REPORT_CASES@@
    End Select
End Function
'''


def build_appdata_vba() -> str:
    search_cases, count_cases, width_cases = [], [], []
    for k in F.SEARCH_KINDS:
        search_cases.append(f"        Case {vba_str(k.key)}")
        search_cases += assign_long_string("s", k.sql, indent="            ")
        count_cases.append(f"        Case {vba_str(k.key)}: SearchColumnCount = {len(k.widths)}")
        widths = ";".join(str(F.cm(w)) for w in k.widths)
        width_cases.append(f"        Case {vba_str(k.key)}: SearchColumnWidths = {vba_str(widths)}")
    report_cases = []
    for i, r in enumerate(F.REPORTS, 1):
        report_cases.append(
            f"        Case {i}: ReportRow = Array({vba_str(r.key)}, {vba_str(r.title)}, "
            f"{vba_str(r.query)}, {vba_str(r.report)}, {vba_str(r.needs)}, {vba_str(r.date_column)})")
    text = APPDATA_TEMPLATE
    for key, value in {
        "@@SEARCH_CASES@@": "\n".join(search_cases),
        "@@COUNT_CASES@@": "\n".join(count_cases),
        "@@WIDTH_CASES@@": "\n".join(width_cases),
        "@@REPORT_COUNT@@": str(len(F.REPORTS)),
        "@@REPORT_CASES@@": "\n".join(report_cases),
        "@@SCREEN_CASES@@": "\n".join(f'        Case {vba_str(form)}: ScreenPermission = {vba_str(key)}'
                                     for form, key in F.SCREEN_PERMISSIONS.items() if key),
    }.items():
        text = text.replace(key, value)
    assert not re.search(r"@@[A-Z_]+@@", text)
    return text
