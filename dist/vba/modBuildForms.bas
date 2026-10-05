Attribute VB_Name = "modBuildForms"
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
Private Const FORM_NAMES As String = "frmMain,frmProducts,frmCustomers,frmSuppliers,frmExpenses,frmUsers,frmCategories,frmUnits,frmExpenseTypes,frmSettings,frmLabelSettings,frmSearch,frmReportCenter,frmPOSLines,frmPOS,frmReturnLines,frmSalesReturn,frmCustomerPayment,frmSalesInvoice,frmPurchaseLines,frmPurchaseInvoice,frmPurchaseReturnLines,frmPurchaseReturn,frmSupplierPayment,frmPurchaseView,frmInventory,frmStockCountLines,frmStockCount,frmLogin,frmChangePassword,frmRolePermLines,frmRoles,frmBackup,frmLabelLines,frmBarcodeLabels"

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
    DoCmd.Echo False, "Ã«—Ì »‰«¡ «·‘«‘« ..."
    BuildAllForms
    DoCmd.Echo True
    Application.RefreshDatabaseWindow
    Debug.Print "---  „ »‰«¡: " & m_built & " | ›‘·: " & m_failed
    If Len(m_warnings) > 0 Then Debug.Print " ‰»ÌÂ« :" & vbCrLf & m_warnings

    If m_failed = 0 Then
        MsgBox " „ »‰«¡ «·‘«‘«  »‰Ã«Õ (" & m_built & " ‘«‘…)." & vbCrLf & vbCrLf & _
               "«·ŒÿÊ… «· «·Ì…: ‘€¯· TestForms À„ «› Õ frmMain", vbInformation + MSG_RTL, "BuildForms"
        BuildForms = True
    Else
        MsgBox "›‘· »‰«¡ " & m_failed & " ‘«‘…:" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "BuildForms"
    End If
    Exit Function
EH:
    DoCmd.Echo True
    MsgBox "Œÿ√ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "BuildForms"
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
    Debug.Print "--- ‰ÃÕ: " & m_passed & " | ›‘·: " & m_failed
    If m_failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·‘«‘«  ‰«ÃÕ… (" & m_passed & " «Œ »«—«)." & vbCrLf & _
               " „ Õ–› »Ì«‰«  «·«Œ »«—.", vbInformation + MSG_RTL, "TestForms"
        TestForms = True
    Else
        TestMsg "‰ÃÕ " & m_passed & " Ê›‘· " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestForms"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "Œÿ√ €Ì— „ Êﬁ⁄ " & Err.Number & ": " & Err.Description
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
    Fail FinalName & ": Œÿ√ " & ErrNumber & " - " & ErrText
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
    Record True, "«·‘«‘… " & FormName & "  › Õ Êﬂ· ﬁÊ«∆„Â«  ⁄„·"
    Exit Sub
EH:
    Record False, "«·‘«‘… " & FormName & ": Œÿ√ " & Err.Number & " - " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, FormName, acSaveNo
End Sub

Private Sub TestHelpers()
    Call Record(RoundMoney(2.345) = 2.35, " ﬁ—Ì» 2.345 = 2.35 (Ê·Ì” 2.34 ﬂ„« ›Ì Round)")
    Call Record(RoundMoney(2.344) = 2.34, " ﬁ—Ì» 2.344 = 2.34")
    Call Record(RoundMoney(-2.345) = -2.35, " ﬁ—Ì» -2.345 = -2.35")
    Call Record(RoundMoney(149.999, 2) = 150, " ﬁ—Ì» 149.999 = 150")
    Call Record(LikePattern("a'b*") = "'*a''b[*]*'", " Â—Ì» ‰’ «·»ÕÀ")
    Call Record(SqlDate(DateSerial(2026, 3, 15) + TimeSerial(14, 5, 9)) = "#2026-03-15 14:05:09#", _
                "’Ì€… «· «—ÌŒ ›Ì SQL")
End Sub

Private Sub TestProductScreen()
    Dim frm As Access.Form, code As String, seqBefore As Long
    seqBefore = DLookup("NextValue", "Sequences", "SequenceName = 'PRODUCT_CODE'")
    DoCmd.OpenForm "frmProducts", acNormal, , , , acHidden
    Set frm = Forms("frmProducts")

    FormAction frm, "NEW"
    frm!ProductName.Value = "TEST-UI „‰ Ã"
    frm!Barcode.Value = "TESTUI0001"
    frm!PurchasePrice.Value = 7
    frm!SellingPrice.Value = 11.5
    Call Record(SaveRecord(frm), "Õ›Ÿ „‰ Ã ÃœÌœ „‰ ‘«‘… «·„‰ Ã« ")
    code = Nz(DLookup("ProductCode", "Products", "Barcode = 'TESTUI0001'"), "")
    Call Record(code Like "P#####", " Ê·Ìœ ﬂÊœ «·„‰ Ã  ·ﬁ«∆Ì« (" & code & ")")
    Call Record(Nz(DLookup("AverageCost", "Products", "Barcode = 'TESTUI0001'"), -1) = 7, _
                "„ Ê”ÿ «· ﬂ·›… ·„‰ Ã ÃœÌœ = ”⁄— «·‘—«¡")
    Call Record(Nz(DLookup("CurrentQuantity", "Products", "Barcode = 'TESTUI0001'"), -1) = 0, _
                "«·ﬂ„Ì… «·«» œ«∆Ì… = 0")

    FormAction frm, "NEW"
    frm!ProductName.Value = "TEST-UI „‰ Ã 2"
    frm!Barcode.Value = "TESTUI0001"
    Call Record(Not SaveRecord(frm), "—›÷ »«—ﬂÊœ „ﬂ—— »—”«·… Ê«÷Õ…")
    frm.Undo

    FormAction frm, "NEW"
    frm!Barcode.Value = "TESTUI0002"
    Call Record(Not SaveRecord(frm), "—›÷ „‰ Ã »œÊ‰ «”„")
    frm.Undo

    frm!txtSearch.Value = "TESTUI0001"
    RefreshList frm
    Call Record(frm!lstItems.ListCount = 2, "«·»ÕÀ ›Ì ﬁ«∆„… «·„‰ Ã«  »«·»«—ﬂÊœ")
    DoCmd.Close acForm, "frmProducts", acSaveNo
    CurrentDb.Execute "UPDATE [Sequences] SET [NextValue] = " & seqBefore & _
                      " WHERE [SequenceName] = 'PRODUCT_CODE'", dbFailOnError
End Sub

Private Sub TestCustomerScreen()
    Dim frm As Access.Form, id As Variant
    DoCmd.OpenForm "frmCustomers", acNormal, , , , acHidden
    Set frm = Forms("frmCustomers")

    FormAction frm, "NEW"
    frm!CustomerName.Value = "TEST-UI ⁄„Ì·"
    frm!Mobile.Value = "0500000000"
    frm!Email.Value = "abc"
    Call Record(Not SaveRecord(frm), "—›÷ »—Ìœ ≈·ﬂ —Ê‰Ì €Ì— ’ÕÌÕ")
    frm!Email.Value = "test@example.com"
    frm!OpeningBalance.Value = 75
    Call Record(SaveRecord(frm), "Õ›Ÿ ⁄„Ì· »—’Ìœ «›  «ÕÌ")
    id = DLookup("CustomerID", "Customers", "CustomerName = 'TEST-UI ⁄„Ì·'")
    Call Record(Nz(DLookup("CurrentBalance", "Customers", "CustomerID = " & Nz(id, 0)), -1) = 75, _
                "«·—’Ìœ «·Õ«·Ì = «·—’Ìœ «·«›  «ÕÌ")
    Call Record(DCount("*", "IntegrityCheckQuery", "IssueCode = 'CUSTOMER_BALANCE'") = 0, _
                "—’Ìœ «·⁄„Ì· „ÿ«»ﬁ ·›Õ’ «·”·«„…")

    DoCmd.Close acForm, "frmCustomers", acSaveNo
    DoCmd.OpenForm "frmCustomers", acNormal, , , , acHidden, 1
    Set frm = Forms("frmCustomers")
    Call Record(frm!CustomerID.Value = 1, "› Õ «·‘«‘… ⁄·Ï ”Ã· „Õœœ («·⁄„Ì· «·‰ﬁœÌ)")
    FormAction frm, "DELETE"
    Call Record(DCount("*", "Customers", "CustomerID = 1") = 1, "„‰⁄ Õ–› «·⁄„Ì· «·‰ﬁœÌ")
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
    frm!Description.Value = "TEST-UI „’—Ê›"
    CalcExpenseVat frm
    Call Record(frm!Tax.Value = 15 And frm!TotalAmount.Value = 115, "Õ”«» ÷—Ì»… «·„’—Ê› 15% Ê«·≈Ã„«·Ì")
    Call Record(SaveRecord(frm), "Õ›Ÿ „’—Ê›")
    num = Nz(DLookup("ExpenseNumber", "Expenses", "Description = 'TEST-UI „’—Ê›'"), "")
    Call Record(num Like "EXP-######", " —ﬁÌ„ «·„’—Ê›  ·ﬁ«∆Ì« (" & num & ")")
    Call Record(Nz(DLookup("EmployeeID", "Expenses", "Description = 'TEST-UI „’—Ê›'"), 0) = CurrentUserID(), _
                " ”ÃÌ· «·„ÊŸ› «·Õ«·Ì ⁄·Ï «·„’—Ê›")
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
        Call Record(Left$(frm!lstResults.RowSource, 6) = "SELECT", "«·»ÕÀ ›Ì " & kind & " Ì⁄„·")
    Next
    frm!cboKind.Value = "PRODUCT"
    SearchKindChanged frm
    frm!txtText.Value = "TESTUI0001"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "«·»ÕÀ «·„ ﬁœ„ »«·»«—ﬂÊœ ÌÃœ «·„‰ Ã")
    frm!txtText.Value = "TEST-UI"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "«·»ÕÀ «·„ ﬁœ„ »Ã“¡ „‰ «·«”„")
    frm!cboKind.Value = "CUSTOMER"
    SearchKindChanged frm
    frm!txtText.Value = "0500000000"
    RunSearch frm
    Call Record(frm!lstResults.ListCount = 2, "«·»ÕÀ ⁄‰ ⁄„Ì· »—ﬁ„ «·ÃÊ«·")
    DoCmd.Close acForm, "frmSearch", acSaveNo
End Sub

Private Sub TestReportCenter()
    Dim frm As Access.Form, i As Long, r As Variant, missing As String
    For i = 1 To ReportCount()
        r = ReportRow(i)
        If Not QueryExists(CStr(r(2))) Then missing = missing & " " & r(2)
    Next
    Call Record(Len(missing) = 0, "ﬂ· «” ⁄·«„«  „—ﬂ“ «· ﬁ«—Ì— „ÊÃÊœ…" & missing)
    DoCmd.OpenForm "frmReportCenter", acNormal, , , , acHidden
    Set frm = Forms("frmReportCenter")
    Call Record(frm!lstReports.ListCount = ReportCount(), "ﬁ«∆„… «· ﬁ«—Ì— (" & ReportCount() & ")")
    frm!lstReports.Value = "CUSTOMER_STATEMENT"
    ReportSelected frm
    Call Record(frm!cboCustomer.Enabled And Not frm!cboSupplier.Enabled And frm!txtFrom.Enabled, _
                "ﬂ‘› Õ”«» ⁄„Ì· Ìÿ·» «·› —… Ê«·⁄„Ì· ›ﬁÿ")
    frm!lstReports.Value = "STOCK"
    ReportSelected frm
    Call Record(Not frm!txtFrom.Enabled And Not frm!cboCustomer.Enabled, _
                " ﬁ—Ì— «·„Œ“Ê‰ ·« Ìÿ·» › —…")
    DoCmd.Close acForm, "frmReportCenter", acSaveNo
End Sub

Private Sub CleanUpTestData(ByVal LastLogID As Long)
    On Error Resume Next
    CloseAllForms
    CurrentDb.Execute "DELETE FROM [Products] WHERE [Barcode] Like 'TESTUI*'"
    CurrentDb.Execute "DELETE FROM [Customers] WHERE [CustomerName] = 'TEST-UI ⁄„Ì·'"
    CurrentDb.Execute "DELETE FROM [Expenses] WHERE [Description] = 'TEST-UI „’—Ê›'"
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
    BuildForm_frmMain
    BuildForm_frmProducts
    BuildForm_frmCustomers
    BuildForm_frmSuppliers
    BuildForm_frmExpenses
    BuildForm_frmUsers
    BuildForm_frmCategories
    BuildForm_frmUnits
    BuildForm_frmExpenseTypes
    BuildForm_frmSettings
    BuildForm_frmLabelSettings
    BuildForm_frmSearch
    BuildForm_frmReportCenter
    BuildForm_frmPOSLines
    BuildForm_frmPOS
    BuildForm_frmReturnLines
    BuildForm_frmSalesReturn
    BuildForm_frmCustomerPayment
    BuildForm_frmSalesInvoice
    BuildForm_frmPurchaseLines
    BuildForm_frmPurchaseInvoice
    BuildForm_frmPurchaseReturnLines
    BuildForm_frmPurchaseReturn
    BuildForm_frmSupplierPayment
    BuildForm_frmPurchaseView
    BuildForm_frmInventory
    BuildForm_frmStockCountLines
    BuildForm_frmStockCount
    BuildForm_frmLogin
    BuildForm_frmChangePassword
    BuildForm_frmRolePermLines
    BuildForm_frmRoles
    BuildForm_frmBackup
    BuildForm_frmLabelLines
    BuildForm_frmBarcodeLabels
End Sub

Private Sub BuildForm_frmMain()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmMain", "‰Ÿ«„ ≈œ«—… «·„Õ·", "", 18994, 9526, False, False, False, _
              ""
    SetFormProp "TimerInterval", 300000
    Set c = AddRect("boxSidebar", 0, 0, 3515, 9526, CLR_PRIMARY)
    Set c = AddIcon("icoApp", ChrW(&HE80F), 227, 255, 567, 567, 22, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblAppTitle", "‰Ÿ«„ ≈œ«—… «·„Õ·", 850, 227, 2551, 425, 15, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblStoreName", " ", 850, 652, 2551, 312, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSales", "«·„»Ì⁄« ", 142, 1304, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmPOS"
    c.OnClick = EP
    Set c = AddIcon("icoSales", ChrW(&HE7BF), 255, 1389, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavPurchases", "«·„‘ —Ì« ", 142, 1871, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmPurchaseInvoice"
    c.OnClick = EP
    Set c = AddIcon("icoPurchases", ChrW(&HE896), 255, 1956, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavInventory", "«·„Œ“Ê‰", 142, 2438, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmInventory"
    c.OnClick = EP
    Set c = AddIcon("icoInventory", ChrW(&HE7B8), 255, 2523, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavProducts", "«·„‰ Ã« ", 142, 3005, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmProducts"
    c.OnClick = EP
    Set c = AddIcon("icoProducts", ChrW(&HE8EC), 255, 3090, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavCustomers", "«·⁄„·«¡", 142, 3572, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmCustomers"
    c.OnClick = EP
    Set c = AddIcon("icoCustomers", ChrW(&HE716), 255, 3657, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSuppliers", "«·„Ê—œÊ‰", 142, 4139, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmSuppliers"
    c.OnClick = EP
    Set c = AddIcon("icoSuppliers", ChrW(&HE77B), 255, 4224, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavExpenses", "«·„’—Ê›« ", 142, 4706, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmExpenses"
    c.OnClick = EP
    Set c = AddIcon("icoExpenses", ChrW(&HE8C7), 255, 4791, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavStockCount", "«·Ã—œ", 142, 5273, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmStockCount"
    c.OnClick = EP
    Set c = AddIcon("icoStockCount", ChrW(&HE8EF), 255, 5358, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavReports", "«· ﬁ«—Ì—", 142, 5840, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmReportCenter"
    c.OnClick = EP
    Set c = AddIcon("icoReports", ChrW(&HE8A5), 255, 5925, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSearch", "«·»ÕÀ", 142, 6407, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmSearch"
    c.OnClick = EP
    Set c = AddIcon("icoSearch", ChrW(&HE721), 255, 6492, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSettings", "«·≈⁄œ«œ« ", 142, 6974, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmSettings"
    c.OnClick = EP
    Set c = AddIcon("icoSettings", ChrW(&HE713), 255, 7059, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavUsers", "«·„” Œœ„Ê‰", 142, 7541, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmUsers"
    c.OnClick = EP
    Set c = AddIcon("icoUsers", ChrW(&HE8D7), 255, 7626, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavBackup", "‰”Œ… «Õ Ì«ÿÌ…", 142, 8108, 3231, 539, "nav")
    SetCtlProp c, "Tag", "frmBackup"
    c.OnClick = EP
    Set c = AddIcon("icoBackup", ChrW(&HE8B7), 255, 8193, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavLogout", " ”ÃÌ· «·Œ—ÊÃ", 142, 8675, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoLogout", ChrW(&HE7E8), 255, 8760, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblWelcome", "·ÊÕ… «· Õﬂ„", 3969, 340, 7938, 539, 20, True, CLR_TEXT, "", 0)
    Set c = AddLabel("lblToday", " ", 3969, 907, 5670, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblUser", " ", 13437, 907, 5103, 340, 11, False, CLR_MUTED, "", 3)
    Set c = AddButton("btnRefresh", " ÕœÌÀ", 17179, 340, 1361, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnChangePassword", "ﬂ·„… «·„—Ê—", 11964, 340, 1531, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblUpdated", " ", 13550, 425, 3515, 312, 9, False, CLR_MUTED, "", 3)
    Set c = AddRect("boxTile1", 3969, 1418, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle1", "„»Ì⁄«  «·ÌÊ„", 4139, 1503, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue1", "-", 4139, 1843, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub1", " ", 4139, 2439, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile2", 7668, 1418, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle2", "„»Ì⁄«  «·‘Â—", 7838, 1503, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue2", "-", 7838, 1843, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub2", " ", 7838, 2439, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile3", 11367, 1418, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle3", "’«›Ì —»Õ «·‘Â— ( ﬁ—Ì»Ì)", 11537, 1503, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue3", "-", 11537, 1843, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub3", " ", 11537, 2439, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile4", 15066, 1418, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle4", "„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰", 15236, 1503, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue4", "-", 15236, 1843, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub4", " ", 15236, 2439, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile5", 3969, 2977, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle5", "œÌÊ‰ «·⁄„·«¡", 4139, 3062, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue5", "-", 4139, 3402, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub5", " ", 4139, 3998, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile6", 7668, 2977, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle6", "„” Õﬁ«  «·„Ê—œÌ‰", 7838, 3062, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue6", "-", 7838, 3402, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub6", " ", 7838, 3998, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile7", 11367, 2977, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle7", "ﬁÌ„… «·„Œ“Ê‰ »«· ﬂ·›…", 11537, 3062, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue7", "-", 11537, 3402, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub7", " ", 11537, 3998, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile8", 15066, 2977, 3472, 1389, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle8", "„’—Ê›«  «·‘Â—", 15236, 3062, 3132, 340, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue8", "-", 15236, 3402, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub8", " ", 15236, 3998, 3132, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCapRecentSales", "›Ê« Ì— «·ÌÊ„ (‰ﬁ— „“œÊÃ ··⁄—÷)", 3969, 4564, 4705, 340, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstRecentSales", 3969, 4933, 4705, 3743, 5, "0;1247;737;1588;1077", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblCapLowStock", "„‰Œ›÷… «·„Œ“Ê‰ (‰ﬁ— „“œÊÃ ··„‰ Ã)", 8901, 4564, 4705, 340, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstLowStock", 8901, 4933, 4705, 3743, 4, "0;2608;1021;1021", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblCapTopProducts", "«·√ﬂÀ— „»Ì⁄« Â–« «·‘Â—", 13833, 4564, 4705, 340, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstTopProducts", 13833, 4933, 4705, 3743, 4, "0;2381;1021;1247", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblIntegrity", " ", 3969, 8846, 14571, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnOpen = EP
    m_frm.OnLoad = EP
    m_frm.OnActivate = EP
    m_frm.OnTimer = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Open(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not MainOpen(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    MainLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Activate()" & vbCrLf
    s = s & "    DashboardActivate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Timer()" & vbCrLf
    s = s & "    DashboardRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSales_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPOS"", 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoSales_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPOS"", 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavPurchases_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseInvoice"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoPurchases_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseInvoice"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavInventory_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmInventory"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoInventory_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmInventory"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavProducts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoProducts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavCustomers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoCustomers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSuppliers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSuppliers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoSuppliers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSuppliers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavExpenses_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenses"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoExpenses_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenses"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavStockCount_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmStockCount"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoStockCount_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmStockCount"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavReports_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmReportCenter"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoReports_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmReportCenter"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSearch_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSearch"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoSearch_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSearch"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSettings"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSettings"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavUsers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUsers"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoUsers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUsers"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavBackup_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBackup"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoBackup_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBackup"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavLogout_Click()" & vbCrLf
    s = s & "    LogoutUser" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoLogout_Click()" & vbCrLf
    s = s & "    LogoutUser" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    DashboardRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnChangePassword_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmChangePassword"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle1_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""TODAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue1_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""TODAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle2_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue2_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle3_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""PROFIT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue3_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""PROFIT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle4_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""LOW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue4_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""LOW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle5_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""DEBT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue5_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""DEBT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle6_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""DUE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue6_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""DUE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle7_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""STOCK""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue7_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""STOCK""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle8_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""EXPENSES""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue8_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""EXPENSES""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstRecentSales_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    OpenScreen ""frmSalesInvoice"", 6, Me!lstRecentSales.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstLowStock_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 0, Me!lstLowStock.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstTopProducts_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 0, Me!lstTopProducts.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxSidebar,0,0,3515,9526,0,0,0,1000;lblUser,13437,907,5103,340,1000,0,0,0;btnRefresh,17179,340,1361,454,1000,0,0,0;btnChangePassword,11964,340,1531,454,1000,0,0,0;lblUpdated,13550,425,3515,312,1000,0,0,0;boxTile1,3969,1418,3472,1389,0,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle1,4139,1503,3132,340,0,250,0,0;lblTileValue1,4139,1843,3132,567,0,250,0,0;lblTileSub1,4139,2439,3132,284,0,250,0,0;boxTile2,7668,1418,3472,1389,250,250,0,0;lblTileTitle2,7838,1503,3132,340,250,250,0,0;lblTileValue2,7838,1843,3132,567,250,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub2,7838,2439,3132,284,250,250,0,0;boxTile3,11367,1418,3472,1389,500,250,0,0;lblTileTitle3,11537,1503,3132,340,500,250,0,0;lblTileValue3,11537,1843,3132,567,500,250,0,0;lblTileSub3,11537,2439,3132,284,500,250,0,0;boxTile4,15066,1418,3472,1389,750,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle4,15236,1503,3132,340,750,250,0,0;lblTileValue4,15236,1843,3132,567,750,250,0,0;lblTileSub4,15236,2439,3132,284,750,250,0,0;boxTile5,3969,2977,3472,1389,0,250,0,0;lblTileTitle5,4139,3062,3132,340,0,250,0,0;lblTileValue5,4139,3402,3132,567,0,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub5,4139,3998,3132,284,0,250,0,0;boxTile6,7668,2977,3472,1389,250,250,0,0;lblTileTitle6,7838,3062,3132,340,250,250,0,0;lblTileValue6,7838,3402,3132,567,250,250,0,0;lblTileSub6,7838,3998,3132,284,250,250,0,0;boxTile7,11367,2977,3472,1389,500,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle7,11537,3062,3132,340,500,250,0,0;lblTileValue7,11537,3402,3132,567,500,250,0,0;lblTileSub7,11537,3998,3132,284,500,250,0,0;boxTile8,15066,2977,3472,1389,750,250,0,0;lblTileTitle8,15236,3062,3132,340,750,250,0,0;lblTileValue8,15236,3402,3132,567,750,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub8,15236,3998,3132,284,750,250,0,0;lblCapRecentSales,3969,4564,4705,340,0,333,0,0;lstRecentSales,3969,4933,4705,3743,0,333,0,1000;lblCapLowStock,8901,4564,4705,340,333,333,0,0;lstLowStock,8901,4933,4705,3743,333,333,0,1000;lblCapTopProducts,13833,4564,4705,340,666,333,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lstTopProducts,13833,4933,4705,3743,666,333,0,1000;lblIntegrity,3969,8846,14571,340,0,1000,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9526, -142, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmMain", s
    Exit Sub
EH:
    AbortForm "frmMain", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmProducts()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmProducts", "«·„‰ Ã« ", "SELECT * FROM Products", 15309, 9213, True, True, True, _
              "KIND=LIST|TABLE=Products|PK=ProductID|LIST=SELECT t.ProductID, t.ProductCode AS [«·ﬂÊœ], t.ProductName AS [«·„‰ Ã], t.CurrentQuantity AS [«·ﬂ„Ì…] FROM Products AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ProductName|SEARCH=t.ProductName,t.ProductCode,t.Barcode,t.ProductNameEn|ACTIVE=t.IsActive|SEQ=PRODUCT_CODE:ProductCode|UNIQUE=ProductCode,Barcode"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8EC), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„‰ Ã« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "≈÷«›… Ê ⁄œÌ· «·√’‰«› Ê«·√”⁄«—", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 5924, 4, "0;1134;2778;850", True)
    c.AfterUpdate = EP
    Set c = AddText("ProductCode", "ProductCode", 7201, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "ÌıÊ·Û¯œ  ·ﬁ«∆Ì« ≈–«  ı—ﬂ ›«—€«"
    SetCtlProp c, "StatusBarText", "ÌıÊ·Û¯œ  ·ﬁ«∆Ì« ≈–«  ı—ﬂ ›«—€«"
    Set c = AddLabel("lblProductCode", "ﬂÊœ «·„‰ Ã", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ProductCode", 0)
    Set c = AddText("Barcode", "Barcode", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblBarcode", "«·»«—ﬂÊœ", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "Barcode", 0)
    Set c = AddText("ProductName", "ProductName", 7201, 2268, 7881, 425)
    Set c = AddLabel("lblProductName", "«”„ «·„‰ Ã *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ProductName", 0)
    Set c = AddText("ProductNameEn", "ProductNameEn", 7201, 2835, 7881, 425)
    Set c = AddLabel("lblProductNameEn", "«·«”„ »«·≈‰Ã·Ì“Ì…", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "ProductNameEn", 0)
    Set c = AddCombo("CategoryID", "CategoryID", 7201, 3402, 2948, 425, "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", 2, "0;3402")
    Set c = AddLabel("lblCategoryID", "«· ’‰Ì›", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "CategoryID", 0)
    Set c = AddCombo("UnitID", "UnitID", 12134, 3402, 2948, 425, "SELECT UnitID, UnitName FROM Units ORDER BY UnitName", 2, "0;3402")
    Set c = AddLabel("lblUnitID", "«·ÊÕœ…", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "UnitID", 0)
    Set c = AddCombo("SupplierID", "SupplierID", 7201, 3969, 2948, 425, "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", 2, "0;3402")
    Set c = AddLabel("lblSupplierID", "«·„Ê—œ «·«› —«÷Ì", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "SupplierID", 0)
    Set c = AddCombo("VATCategory", "VATCategory", 12134, 3969, 2948, 425, "S;Œ«÷⁄ ··÷—Ì»… 15%;Z;‰”»… ’›—Ì…;E;„⁄›Ï „‰ «·÷—Ì»…", 2, "0;2552")
    c.AfterUpdate = EP
    Set c = AddLabel("lblVATCategory", "«·›∆… «·÷—Ì»Ì…", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "VATCategory", 0)
    Set c = AddText("SellingPrice", "SellingPrice", 7201, 4536, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblSellingPrice", "”⁄— «·»Ì⁄", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "SellingPrice", 0)
    Set c = AddText("PurchasePrice", "PurchasePrice", 12134, 4536, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblPurchasePrice", "¬Œ— ”⁄— ‘—«¡", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "PurchasePrice", 0)
    Set c = AddLabel("lblPriceInfo", " ", 5443, 5103, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("AverageCost", "AverageCost", 7201, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblAverageCost", "„ Ê”ÿ «· ﬂ·›…", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "AverageCost", 0)
    Set c = AddText("CurrentQuantity", "CurrentQuantity", 12134, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentQuantity", "«·ﬂ„Ì… «·Õ«·Ì…", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "CurrentQuantity", 0)
    Set c = AddText("MinimumQuantity", "MinimumQuantity", 7201, 6237, 2948, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMinimumQuantity", "Õœ ≈⁄«œ… «·ÿ·»", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "MinimumQuantity", 0)
    Set c = AddText("ProductLocation", "ProductLocation", 12134, 6237, 2948, 425)
    Set c = AddLabel("lblProductLocation", "„ﬂ«‰ «·„‰ Ã", 10376, 6237, 1701, 425, 10, False, CLR_MUTED, "ProductLocation", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 6889)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblStockNote", "«·ﬂ„Ì…   €Ì— ›ﬁÿ „‰ «·„‘ —Ì«  Ê«·„»Ì⁄«  Ê«·Ã—œ", 10376, 6804, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 7371, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 8533, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub VATCategory_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""VATCategory""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub SellingPrice_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""SellingPrice""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmProducts", s
    Exit Sub
EH:
    AbortForm "frmProducts", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCustomers()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCustomers", "«·⁄„·«¡", "SELECT * FROM Customers", 15309, 9213, True, True, True, _
              "KIND=LIST|TABLE=Customers|PK=CustomerID|LIST=SELECT t.CustomerID, t.CustomerName AS [«·⁄„Ì·], t.Mobile AS [«·ÃÊ«·], t.CurrentBalance AS [«·—’Ìœ] FROM Customers AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.CustomerName|SEARCH=t.CustomerName,t.Mobile,t.Phone,t.VATNumber|ACTIVE=t.IsActive"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·⁄„·«¡", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "»Ì«‰«  «·⁄„·«¡ Ê√—’œ Â„", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "”‰œ ﬁ»÷", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStatement", "ﬂ‘› Õ”«»", 7937, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 5924, 4, "0;2495;1361;907", True)
    c.AfterUpdate = EP
    Set c = AddText("CustomerName", "CustomerName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblCustomerName", "«”„ «·⁄„Ì· *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CustomerName", 0)
    Set c = AddText("Mobile", "Mobile", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblMobile", "«·ÃÊ«·", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("Phone", "Phone", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblPhone", "«·Â« ›", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblEmail", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATNumber", "VATNumber", 12134, 2835, 2948, 425)
    SetCtlProp c, "ControlTipText", "··⁄„·«¡ «·„‰‘¬  (›« Ê—… ÷—Ì»Ì…)"
    SetCtlProp c, "StatusBarText", "··⁄„·«¡ «·„‰‘¬  (›« Ê—… ÷—Ì»Ì…)"
    Set c = AddLabel("lblVATNumber", "«·—ﬁ„ «·÷—Ì»Ì", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 7201, 3402, 2948, 425)
    Set c = AddLabel("lblCRNumber", "«·”Ã· «· Ã«—Ì", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("City", "City", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblCity", "«·„œÌ‰…", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddText("District", "District", 7201, 3969, 2948, 425)
    Set c = AddLabel("lblDistrict", "«·ÕÌ", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "District", 0)
    Set c = AddText("StreetName", "StreetName", 12134, 3969, 2948, 425)
    Set c = AddLabel("lblStreetName", "«·‘«—⁄", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "StreetName", 0)
    Set c = AddText("BuildingNo", "BuildingNo", 7201, 4536, 2948, 425)
    Set c = AddLabel("lblBuildingNo", "—ﬁ„ «·„»‰Ï", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "BuildingNo", 0)
    Set c = AddText("PostalCode", "PostalCode", 12134, 4536, 2948, 425)
    Set c = AddLabel("lblPostalCode", "«·—„“ «·»—ÌœÌ", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "PostalCode", 0)
    Set c = AddText("Address", "Address", 7201, 5103, 7881, 425)
    Set c = AddLabel("lblAddress", "«·⁄‰Ê«‰", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "Address", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "Ìıﬁ›· »⁄œ √Ê· ⁄„·Ì…"
    SetCtlProp c, "StatusBarText", "Ìıﬁ›· »⁄œ √Ê· ⁄„·Ì…"
    Set c = AddLabel("lblOpeningBalance", "«·—’Ìœ «·«›  «ÕÌ", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("CurrentBalance", "CurrentBalance", 12134, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentBalance", "«·—’Ìœ «·Õ«·Ì", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "CurrentBalance", 0)
    Set c = AddCheck("AllowCredit", "AllowCredit", 7201, 6322)
    Set c = AddLabel("lblAllowCredit", "Ì”„Õ »«·»Ì⁄ «·¬Ã·", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "AllowCredit", 0)
    Set c = AddText("CreditLimit", "CreditLimit", 12134, 6237, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "0 = »œÊ‰ Õœ"
    SetCtlProp c, "StatusBarText", "0 = »œÊ‰ Õœ"
    Set c = AddLabel("lblCreditLimit", "Õœ «·«∆ „«‰", 10376, 6237, 1701, 425, 10, False, CLR_MUTED, "CreditLimit", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 6889)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "«·—’Ìœ «·„ÊÃ» = „»·€ „” Õﬁ ⁄·Ï «·⁄„Ì·", 10376, 6804, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 7371, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 8533, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayment_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomerPayment"", 6, Me!CustomerID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStatement_Click()" & vbCrLf
    s = s & "    PrintPartyStatement ""C"", Me!CustomerID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCustomers", s
    Exit Sub
EH:
    AbortForm "frmCustomers", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSuppliers()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSuppliers", "«·„Ê—œÊ‰", "SELECT * FROM Suppliers", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Suppliers|PK=SupplierID|LIST=SELECT t.SupplierID, t.SupplierName AS [«·„Ê—œ], t.Mobile AS [«·ÃÊ«·], t.CurrentBalance AS [«·—’Ìœ] FROM Suppliers AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.SupplierName|SEARCH=t.SupplierName,t.ContactPerson,t.Mobile,t.VATNumber|ACTIVE=t.IsActive"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE77B), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„Ê—œÊ‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "»Ì«‰«  «·„Ê—œÌ‰ Ê√—’œ Â„", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "”‰œ ’—›", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStatement", "ﬂ‘› Õ”«»", 7937, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;2495;1361;907", True)
    c.AfterUpdate = EP
    Set c = AddText("SupplierName", "SupplierName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblSupplierName", "«”„ «·„Ê—œ *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "SupplierName", 0)
    Set c = AddText("ContactPerson", "ContactPerson", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblContactPerson", "«·‘Œ’ «·„”ƒÊ·", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ContactPerson", 0)
    Set c = AddText("Mobile", "Mobile", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblMobile", "«·ÃÊ«·", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("Phone", "Phone", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblPhone", "«·Â« ›", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 12134, 2835, 2948, 425)
    Set c = AddLabel("lblEmail", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATNumber", "VATNumber", 7201, 3402, 2948, 425)
    Set c = AddLabel("lblVATNumber", "«·—ﬁ„ «·÷—Ì»Ì", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblCRNumber", "«·”Ã· «· Ã«—Ì", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("City", "City", 7201, 3969, 2948, 425)
    Set c = AddLabel("lblCity", "«·„œÌ‰…", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddLabel("lblSupplierNote", " ", 10376, 3969, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Address", "Address", 7201, 4536, 7881, 425)
    Set c = AddLabel("lblAddress", "«·⁄‰Ê«‰", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "Address", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 5103, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "Ìıﬁ›· »⁄œ √Ê· ⁄„·Ì…"
    SetCtlProp c, "StatusBarText", "Ìıﬁ›· »⁄œ √Ê· ⁄„·Ì…"
    Set c = AddLabel("lblOpeningBalance", "«·—’Ìœ «·«›  «ÕÌ", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("CurrentBalance", "CurrentBalance", 12134, 5103, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentBalance", "«·—’Ìœ «·Õ«·Ì", 10376, 5103, 1701, 425, 10, False, CLR_MUTED, "CurrentBalance", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 5755)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "«·—’Ìœ «·„ÊÃ» = „»·€ „” Õﬁ ··„Ê—œ", 10376, 5670, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 6237, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 7399, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayment_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSupplierPayment"", 7, Me!SupplierID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStatement_Click()" & vbCrLf
    s = s & "    PrintPartyStatement ""S"", Me!SupplierID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSuppliers", s
    Exit Sub
EH:
    AbortForm "frmSuppliers", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmExpenses()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmExpenses", "«·„’—Ê›« ", "SELECT * FROM Expenses", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Expenses|PK=ExpenseID|LIST=SELECT t.ExpenseID, t.ExpenseNumber AS [«·—ﬁ„], t.ExpenseDate AS [«· «—ÌŒ], x.ExpenseTypeName AS [«·‰Ê⁄], t.TotalAmount AS [«·„»·€] FROM Expenses AS t INNER JOIN ExpenseTypes AS x ON t.ExpenseTypeID = x.ExpenseTypeID WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ExpenseDate DESC, t.ExpenseID DESC|SEARCH=t.ExpenseNumber,t.Description,x.ExpenseTypeName,t.SupplierInvoiceRef|SEQ=EXPENSE:ExpenseNumber|UNIQUE=ExpenseNumber"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8C7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„’—Ê›« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " ”ÃÌ· „’—Ê›«  «·„Õ·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddList("lstItems", 227, 2551, 4990, 5387, 5, "0;1134;1191;1474;964", True)
    c.AfterUpdate = EP
    Set c = AddText("ExpenseNumber", "ExpenseNumber", 7201, 1701, 2948, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "ControlTipText", "ÌıÊ·Û¯œ ⁄‰œ «·Õ›Ÿ"
    SetCtlProp c, "StatusBarText", "ÌıÊ·Û¯œ ⁄‰œ «·Õ›Ÿ"
    Set c = AddLabel("lblExpenseNumber", "—ﬁ„ «·„’—Ê›", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseNumber", 0)
    Set c = AddText("ExpenseDate", "ExpenseDate", 12134, 1701, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblExpenseDate", " «—ÌŒ «·„’—Ê›", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseDate", 0)
    Set c = AddCombo("ExpenseTypeID", "ExpenseTypeID", 7201, 2268, 2948, 425, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName", 2, "0;3402")
    Set c = AddLabel("lblExpenseTypeID", "‰Ê⁄ «·„’—Ê› *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeID", 0)
    Set c = AddCombo("PaymentMethodID", "PaymentMethodID", 12134, 2268, 2948, 425, "SELECT PaymentMethodID, MethodName FROM PaymentMethods ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethodID", "ÿ—Ìﬁ… «·œ›⁄", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "PaymentMethodID", 0)
    Set c = AddText("Amount", "Amount", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblAmount", "«·„»·€ ﬁ»· «·÷—Ì»…", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Amount", 0)
    Set c = AddText("Tax", "Tax", 12134, 2835, 1644, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTax", "÷—Ì»… «·„œŒ·« ", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Tax", 0)
    Set c = AddButton("btnCalcVat", "«Õ”» 15%", 13835, 2835, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("TotalAmount", "TotalAmount", 7201, 3402, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblTotalAmount", "«·≈Ã„«·Ì", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "TotalAmount", 0)
    Set c = AddText("SupplierInvoiceRef", "SupplierInvoiceRef", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblSupplierInvoiceRef", "—ﬁ„ ›« Ê—… «·„’—Ê›", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "SupplierInvoiceRef", 0)
    Set c = AddText("Description", "Description", 7201, 3969, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "«·Ê’›", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 5131, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Amount_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""Amount""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Tax_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""Tax""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCalcVat_Click()" & vbCrLf
    s = s & "    CalcExpenseVat Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmExpenses", s
    Exit Sub
EH:
    AbortForm "frmExpenses", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmUsers()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmUsers", "«·„” Œœ„Ê‰", "SELECT * FROM Employees", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Employees|PK=EmployeeID|LIST=SELECT t.EmployeeID, t.Username AS [«·„” Œœ„], t.EmployeeName AS [«·«”„], r.RoleName AS [«·œÊ—] FROM Employees AS t INNER JOIN Roles AS r ON t.RoleID = r.RoleID WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.EmployeeName|SEARCH=t.EmployeeName,t.Username,t.Mobile|ACTIVE=t.IsActive|UNIQUE=Username"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„” Œœ„Ê‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·„ÊŸ›Ê‰ Ê√”„«¡ «·œŒÊ· Ê«·√œÊ«—", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSetPassword", "ﬂ·„… «·„—Ê—", 4649, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnUnlock", "›ﬂ «·ﬁ›·", 6463, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnRoles", "«·’·«ÕÌ« ", 8277, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;1361;2155;1247", True)
    c.AfterUpdate = EP
    Set c = AddText("EmployeeName", "EmployeeName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblEmployeeName", "«”„ «·„ÊŸ› *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "EmployeeName", 0)
    Set c = AddText("Username", "Username", 7201, 2268, 2948, 425)
    SetCtlProp c, "ControlTipText", "»œÊ‰ „”«›« ° 3 √Õ—› ⁄·Ï «·√ﬁ·"
    SetCtlProp c, "StatusBarText", "»œÊ‰ „”«›« ° 3 √Õ—› ⁄·Ï «·√ﬁ·"
    Set c = AddLabel("lblUsername", "«”„ «·„” Œœ„ *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Username", 0)
    Set c = AddCombo("RoleID", "RoleID", 12134, 2268, 2948, 425, "SELECT RoleID, RoleName FROM Roles ORDER BY RoleID", 2, "0;2268")
    Set c = AddLabel("lblRoleID", "«·œÊ— *", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "RoleID", 0)
    Set c = AddText("JobTitle", "JobTitle", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblJobTitle", "«·„”„Ï «·ÊŸÌ›Ì", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "JobTitle", 0)
    Set c = AddText("Mobile", "Mobile", 12134, 2835, 2948, 425)
    Set c = AddLabel("lblMobile", "«·ÃÊ«·", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("MaxDiscountPercent", "MaxDiscountPercent", 7201, 3402, 2948, 425)
    SetCtlProp c, "Format", "0.00%"
    SetCtlProp c, "ControlTipText", "√ﬁ’Ï Œ’„ »œÊ‰ „Ê«›ﬁ… („À«· 5%)"
    SetCtlProp c, "StatusBarText", "√ﬁ’Ï Œ’„ »œÊ‰ „Ê«›ﬁ… („À«· 5%)"
    Set c = AddLabel("lblMaxDiscountPercent", "√ﬁ’Ï ‰”»… Œ’„", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "MaxDiscountPercent", 0)
    Set c = AddCheck("IsActive", "IsActive", 12134, 3487)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddCheck("MustChangePassword", "MustChangePassword", 7201, 4054)
    Set c = AddLabel("lblMustChangePassword", "ÌÃ»  €ÌÌ— ﬂ·„… «·„—Ê—", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "MustChangePassword", 0)
    Set c = AddText("LastLoginAt", "LastLoginAt", 12134, 3969, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblLastLoginAt", "¬Œ— œŒÊ·", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "LastLoginAt", 0)
    Set c = AddText("FailedLoginCount", "FailedLoginCount", 7201, 4536, 2948, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblFailedLoginCount", "„Õ«Ê·«  «·œŒÊ· «·›«‘·…", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "FailedLoginCount", 0)
    Set c = AddText("LockedUntil", "LockedUntil", 12134, 4536, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblLockedUntil", "„ﬁ›· Õ Ï", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "LockedUntil", 0)
    Set c = AddLabel("lblPasswordState", " ", 5443, 5103, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 5670, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 6832, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSetPassword_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmChangePassword"", 10, Me!EmployeeID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnlock_Click()" & vbCrLf
    s = s & "    UnlockUser Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRoles_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmRoles"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmUsers", s
    Exit Sub
EH:
    AbortForm "frmUsers", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCategories()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCategories", "«· ’‰Ì›« ", "SELECT * FROM Categories", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Categories|PK=CategoryID|LIST=SELECT t.CategoryID, t.CategoryName AS [«· ’‰Ì›] FROM Categories AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.CategoryName|SEARCH=t.CategoryName,t.Description|ACTIVE=t.IsActive|UNIQUE=CategoryName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8FD), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«· ’‰Ì›« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " ’‰Ì›«  «·„‰ Ã« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 2, "0;4763", True)
    c.AfterUpdate = EP
    Set c = AddText("CategoryName", "CategoryName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblCategoryName", "«”„ «· ’‰Ì› *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CategoryName", 0)
    Set c = AddText("Description", "Description", 7201, 2268, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "«·Ê’›", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 3402)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 3317, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 3997, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCategories", s
    Exit Sub
EH:
    AbortForm "frmCategories", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmUnits()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmUnits", "ÊÕœ«  «·ﬁÌ«”", "SELECT * FROM Units", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Units|PK=UnitID|LIST=SELECT t.UnitID, t.UnitName AS [«·ÊÕœ…], t.ZatcaUnitCode AS [«·—„“] FROM Units AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.UnitName|SEARCH=t.UnitName,t.ZatcaUnitCode|ACTIVE=t.IsActive|UNIQUE=UnitName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8FD), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ÊÕœ«  «·ﬁÌ«”", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ÊÕœ«  »Ì⁄ «·„‰ Ã« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 3, "0;3062;1701", True)
    c.AfterUpdate = EP
    Set c = AddText("UnitName", "UnitName", 7201, 1701, 2948, 425)
    Set c = AddLabel("lblUnitName", "«”„ «·ÊÕœ… *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "UnitName", 0)
    Set c = AddText("ZatcaUnitCode", "ZatcaUnitCode", 12134, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "„À«·: PCE ··Õ»…° KGM ··ﬂÌ·Ê"
    SetCtlProp c, "StatusBarText", "„À«·: PCE ··Õ»…° KGM ··ﬂÌ·Ê"
    Set c = AddLabel("lblZatcaUnitCode", "—„“ «·ÊÕœ… (UN/ECE)", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ZatcaUnitCode", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 2353)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 2948, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmUnits", s
    Exit Sub
EH:
    AbortForm "frmUnits", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmExpenseTypes()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmExpenseTypes", "√‰Ê«⁄ «·„’—Ê›« ", "SELECT * FROM ExpenseTypes", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=ExpenseTypes|PK=ExpenseTypeID|LIST=SELECT t.ExpenseTypeID, t.ExpenseTypeName AS [«·‰Ê⁄] FROM ExpenseTypes AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ExpenseTypeName|SEARCH=t.ExpenseTypeName|ACTIVE=t.IsActive|UNIQUE=ExpenseTypeName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8C7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "√‰Ê«⁄ «·„’—Ê›« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁ«∆„… √‰Ê«⁄ «·„’—Ê›« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "≈ŸÂ«— €Ì— «·‰‘ÿ", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 2, "0;4763", True)
    c.AfterUpdate = EP
    Set c = AddText("ExpenseTypeName", "ExpenseTypeName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblExpenseTypeName", "‰Ê⁄ «·„’—Ê› *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeName", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 2353)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 2948, 9639, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    FormAction Me, ""NEW""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    FormAction Me, ""DELETE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmExpenseTypes", s
    Exit Sub
EH:
    AbortForm "frmExpenseTypes", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSettings()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSettings", "«·≈⁄œ«œ« ", "SELECT * FROM Settings WHERE SettingID = 1", 15309, 8222, True, False, True, _
              "KIND=SINGLE|TABLE=Settings|PK=SettingID"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE713), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·≈⁄œ«œ« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "»Ì«‰«  «·„Õ· «·÷—Ì»Ì… Ê≈⁄œ«œ«  «· ‘€Ì·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ", 227, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 1701, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCategories", "«· ’‰Ì›« ", 3175, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnUnits", "«·ÊÕœ« ", 4989, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnExpenseTypes", "√‰Ê«⁄ «·„’—Ê›« ", 6803, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLabelSettings", "„·’ﬁ«  «·»«—ﬂÊœ", 8617, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddText("StoreName", "StoreName", 2552, 1701, 4989, 425)
    Set c = AddLabel("lblStoreName", "«”„ «·„Õ· *", 227, 1701, 2268, 425, 10, False, CLR_MUTED, "StoreName", 0)
    Set c = AddText("StoreNameEn", "StoreNameEn", 10093, 1701, 4989, 425)
    Set c = AddLabel("lblStoreNameEn", "«”„ «·„Õ· »«·≈‰Ã·Ì“Ì…", 7768, 1701, 2268, 425, 10, False, CLR_MUTED, "StoreNameEn", 0)
    Set c = AddText("VATNumber", "VATNumber", 2552, 2268, 4989, 425)
    SetCtlProp c, "ControlTipText", "15 —ﬁ„« Ì»œ√ ÊÌ‰ ÂÌ »‹ 3"
    SetCtlProp c, "StatusBarText", "15 —ﬁ„« Ì»œ√ ÊÌ‰ ÂÌ »‹ 3"
    Set c = AddLabel("lblVATNumber", "«·—ﬁ„ «·÷—Ì»Ì", 227, 2268, 2268, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 10093, 2268, 4989, 425)
    Set c = AddLabel("lblCRNumber", "«·”Ã· «· Ã«—Ì", 7768, 2268, 2268, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("BuildingNo", "BuildingNo", 2552, 2835, 4989, 425)
    Set c = AddLabel("lblBuildingNo", "—ﬁ„ «·„»‰Ï", 227, 2835, 2268, 425, 10, False, CLR_MUTED, "BuildingNo", 0)
    Set c = AddText("StreetName", "StreetName", 10093, 2835, 4989, 425)
    Set c = AddLabel("lblStreetName", "«·‘«—⁄", 7768, 2835, 2268, 425, 10, False, CLR_MUTED, "StreetName", 0)
    Set c = AddText("District", "District", 2552, 3402, 4989, 425)
    Set c = AddLabel("lblDistrict", "«·ÕÌ", 227, 3402, 2268, 425, 10, False, CLR_MUTED, "District", 0)
    Set c = AddText("City", "City", 10093, 3402, 4989, 425)
    Set c = AddLabel("lblCity", "«·„œÌ‰…", 7768, 3402, 2268, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddText("PostalCode", "PostalCode", 2552, 3969, 4989, 425)
    Set c = AddLabel("lblPostalCode", "«·—„“ «·»—ÌœÌ", 227, 3969, 2268, 425, 10, False, CLR_MUTED, "PostalCode", 0)
    Set c = AddText("AdditionalNo", "AdditionalNo", 10093, 3969, 4989, 425)
    Set c = AddLabel("lblAdditionalNo", "«·—ﬁ„ «·≈÷«›Ì", 7768, 3969, 2268, 425, 10, False, CLR_MUTED, "AdditionalNo", 0)
    Set c = AddText("Phone", "Phone", 2552, 4536, 4989, 425)
    Set c = AddLabel("lblPhone", "«·Â« ›", 227, 4536, 2268, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 10093, 4536, 4989, 425)
    Set c = AddLabel("lblEmail", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", 7768, 4536, 2268, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATRate", "VATRate", 2552, 5103, 4989, 425)
    SetCtlProp c, "Format", "0.00%"
    Set c = AddLabel("lblVATRate", "‰”»… «·÷—Ì»…", 227, 5103, 2268, 425, 10, False, CLR_MUTED, "VATRate", 0)
    Set c = AddCheck("PricesIncludeVAT", "PricesIncludeVAT", 10093, 5188)
    Set c = AddLabel("lblPricesIncludeVAT", "«·√”⁄«— ‘«„·… «·÷—Ì»…", 7768, 5103, 2268, 425, 10, False, CLR_MUTED, "PricesIncludeVAT", 0)
    Set c = AddCheck("AllowNegativeStock", "AllowNegativeStock", 2552, 5755)
    Set c = AddLabel("lblAllowNegativeStock", "«·”„«Õ »«·»Ì⁄ »«·”«·»", 227, 5670, 2268, 425, 10, False, CLR_MUTED, "AllowNegativeStock", 0)
    Set c = AddText("SlowMovingDays", "SlowMovingDays", 10093, 5670, 4989, 425)
    Set c = AddLabel("lblSlowMovingDays", "√Ì«„ ⁄œ„ «·Õ—ﬂ…", 7768, 5670, 2268, 425, 10, False, CLR_MUTED, "SlowMovingDays", 0)
    Set c = AddText("BackupFolder", "BackupFolder", 2552, 6237, 3685, 425)
    Set c = AddLabel("lblBackupFolder", "„Ã·œ «·‰”Œ «·«Õ Ì«ÿÌ", 227, 6237, 2268, 425, 10, False, CLR_MUTED, "BackupFolder", 0)
    Set c = AddButton("btnBrowseBackup", "«” ⁄—«÷", 6294, 6237, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("BackupKeepCount", "BackupKeepCount", 10093, 6237, 4989, 425)
    Set c = AddLabel("lblBackupKeepCount", "⁄œœ «·‰”Œ «·„Õ ›Ÿ »Â«", 7768, 6237, 2268, 425, 10, False, CLR_MUTED, "BackupKeepCount", 0)
    Set c = AddText("LogoPath", "LogoPath", 2552, 6804, 3685, 425)
    Set c = AddLabel("lblLogoPath", "„”«— «·‘⁄«—", 227, 6804, 2268, 425, 10, False, CLR_MUTED, "LogoPath", 0)
    Set c = AddButton("btnBrowseLogo", "«” ⁄—«÷", 6294, 6804, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("ReceiptFooter", "ReceiptFooter", 10093, 6804, 4989, 425)
    Set c = AddLabel("lblReceiptFooter", " –ÌÌ· «·›« Ê—…", 7768, 6804, 2268, 425, 10, False, CLR_MUTED, "ReceiptFooter", 0)
    Set c = AddLabel("lblStatus", " ", 227, 7484, 14855, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCategories_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCategories""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnits_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUnits""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnExpenseTypes_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenseTypes""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLabelSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLabelSettings""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBrowseBackup_Click()" & vbCrLf
    s = s & "    BrowseFolder Me, ""BackupFolder""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBrowseLogo_Click()" & vbCrLf
    s = s & "    BrowseFile Me, ""LogoPath""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSettings", s
    Exit Sub
EH:
    AbortForm "frmSettings", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmLabelSettings()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmLabelSettings", "≈⁄œ«œ«  „·’ﬁ«  «·»«—ﬂÊœ", "SELECT * FROM LabelSettings WHERE LabelSettingID = 1", 15309, 8731, True, False, True, _
              "KIND=SINGLE|TABLE=LabelSettings|PK=LabelSettingID"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE713), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "≈⁄œ«œ«  „·’ﬁ«  «·»«—ﬂÊœ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„ﬁ«” «·„·’ﬁ Ê«·Ê—ﬁ Ê«·ÂÊ«„‘° ÊÕÃ„ «·»«—ﬂÊœ° Ê«·‰’Ê’ √⁄·«Â Ê√”›·Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ", 227, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 1701, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblInfoPaper", "«·„·’ﬁ Ê«·Ê—ﬁ (»«·„·Ì„ —). „ﬁ«” Ê—ﬁ «·ÿ«»⁄… ‰›”Â Ìı÷»ÿ „‰ ≈⁄œ«œ«  «·ÿ«»⁄… ›Ì Windows", 227, 1701, 14855, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddCombo("PrinterName", "PrinterName", 2552, 2268, 4989, 425, "", 1, "4536")
    SetCtlProp c, "ControlTipText", "« —ﬂÂ ›«—€« ··ÿ»«⁄… ⁄·Ï «·ÿ«»⁄… «·«› —«÷Ì…"
    SetCtlProp c, "StatusBarText", "« —ﬂÂ ›«—€« ··ÿ»«⁄… ⁄·Ï «·ÿ«»⁄… «·«› —«÷Ì…"
    Set c = AddLabel("lblPrinterName", "ÿ«»⁄… «·„·’ﬁ« ", 227, 2268, 2268, 425, 10, False, CLR_MUTED, "PrinterName", 0)
    Set c = AddText("LabelsAcross", "LabelsAcross", 10093, 2268, 4989, 425)
    SetCtlProp c, "ControlTipText", "1 ·ÿ«»⁄… «·„·’ﬁ« ° Ê√ﬂÀ— ·Ê—ﬁ A4 ›ÌÂ √⁄„œ… „·’ﬁ« "
    SetCtlProp c, "StatusBarText", "1 ·ÿ«»⁄… «·„·’ﬁ« ° Ê√ﬂÀ— ·Ê—ﬁ A4 ›ÌÂ √⁄„œ… „·’ﬁ« "
    Set c = AddLabel("lblLabelsAcross", "⁄œœ «·„·’ﬁ«  ›Ì «·’›", 7768, 2268, 2268, 425, 10, False, CLR_MUTED, "LabelsAcross", 0)
    Set c = AddText("LabelWidth", "LabelWidth", 2552, 2835, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "ControlTipText", "„À«·: 38 √Ê 40 √Ê 50"
    SetCtlProp c, "StatusBarText", "„À«·: 38 √Ê 40 √Ê 50"
    Set c = AddLabel("lblLabelWidth", "⁄—÷ «·„·’ﬁ („„)", 227, 2835, 2268, 425, 10, False, CLR_MUTED, "LabelWidth", 0)
    Set c = AddText("LabelHeight", "LabelHeight", 10093, 2835, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "ControlTipText", "„À«·: 25 √Ê 30"
    SetCtlProp c, "StatusBarText", "„À«·: 25 √Ê 30"
    Set c = AddLabel("lblLabelHeight", "«— ›«⁄ «·„·’ﬁ („„)", 7768, 2835, 2268, 425, 10, False, CLR_MUTED, "LabelHeight", 0)
    Set c = AddText("ColumnGap", "ColumnGap", 2552, 3402, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblColumnGap", "«·„”«›… »Ì‰ «·√⁄„œ… („„)", 227, 3402, 2268, 425, 10, False, CLR_MUTED, "ColumnGap", 0)
    Set c = AddText("RowGap", "RowGap", 10093, 3402, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblRowGap", "«·„”«›… »Ì‰ «·’›Ê› („„)", 7768, 3402, 2268, 425, 10, False, CLR_MUTED, "RowGap", 0)
    Set c = AddText("MarginTop", "MarginTop", 2552, 3969, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMarginTop", "«·Â«„‘ «·⁄·ÊÌ („„)", 227, 3969, 2268, 425, 10, False, CLR_MUTED, "MarginTop", 0)
    Set c = AddText("MarginBottom", "MarginBottom", 10093, 3969, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMarginBottom", "«·Â«„‘ «·”›·Ì („„)", 7768, 3969, 2268, 425, 10, False, CLR_MUTED, "MarginBottom", 0)
    Set c = AddText("MarginRight", "MarginRight", 2552, 4536, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMarginRight", "«·Â«„‘ «·√Ì„‰ („„)", 227, 4536, 2268, 425, 10, False, CLR_MUTED, "MarginRight", 0)
    Set c = AddText("MarginLeft", "MarginLeft", 10093, 4536, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMarginLeft", "«·Â«„‘ «·√Ì”— („„)", 7768, 4536, 2268, 425, 10, False, CLR_MUTED, "MarginLeft", 0)
    Set c = AddLabel("lblInfoBar", "«·»«—ﬂÊœ Ê«·‰’Ê’", 227, 5103, 14855, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("BarHeight", "BarHeight", 2552, 5670, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblBarHeight", "«— ›«⁄ «·»«—ﬂÊœ („„)", 227, 5670, 2268, 425, 10, False, CLR_MUTED, "BarHeight", 0)
    Set c = AddText("BarWidth", "BarWidth", 10093, 5670, 4989, 425)
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "ControlTipText", "0.25 „‰«”» ·ÿ«»⁄«  203 ‰ﬁÿ…/»Ê’…° Êﬂ»¯—Â ≈–« ’⁄»  «·ﬁ—«¡…"
    SetCtlProp c, "StatusBarText", "0.25 „‰«”» ·ÿ«»⁄«  203 ‰ﬁÿ…/»Ê’…° Êﬂ»¯—Â ≈–« ’⁄»  «·ﬁ—«¡…"
    Set c = AddLabel("lblBarWidth", "⁄—÷ √—›⁄ Œÿ („„)", 7768, 5670, 2268, 425, 10, False, CLR_MUTED, "BarWidth", 0)
    Set c = AddCombo("TopLine1", "TopLine1", 2552, 6237, 4989, 425, "NONE;»œÊ‰;STORE;«·«”„ «·„Œ ’— ··„Õ·;NAME;«”„ «·„‰ Ã;PRICE;«·”⁄—;CODE;ﬂÊœ «·„‰ Ã;BARCODE;—ﬁ„ «·»«—ﬂÊœ", 2, "0;3402")
    Set c = AddLabel("lblTopLine1", "«·”ÿ— «·√Ê· √⁄·Ï «·»«—ﬂÊœ", 227, 6237, 2268, 425, 10, False, CLR_MUTED, "TopLine1", 0)
    Set c = AddCombo("TopLine2", "TopLine2", 10093, 6237, 4989, 425, "NONE;»œÊ‰;STORE;«·«”„ «·„Œ ’— ··„Õ·;NAME;«”„ «·„‰ Ã;PRICE;«·”⁄—;CODE;ﬂÊœ «·„‰ Ã;BARCODE;—ﬁ„ «·»«—ﬂÊœ", 2, "0;3402")
    Set c = AddLabel("lblTopLine2", "«·”ÿ— «·À«‰Ì √⁄·Ï «·»«—ﬂÊœ", 7768, 6237, 2268, 425, 10, False, CLR_MUTED, "TopLine2", 0)
    Set c = AddCombo("BottomLine1", "BottomLine1", 2552, 6804, 4989, 425, "NONE;»œÊ‰;STORE;«·«”„ «·„Œ ’— ··„Õ·;NAME;«”„ «·„‰ Ã;PRICE;«·”⁄—;CODE;ﬂÊœ «·„‰ Ã;BARCODE;—ﬁ„ «·»«—ﬂÊœ", 2, "0;3402")
    Set c = AddLabel("lblBottomLine1", "«·”ÿ— «·√Ê· √”›· «·»«—ﬂÊœ", 227, 6804, 2268, 425, 10, False, CLR_MUTED, "BottomLine1", 0)
    Set c = AddCombo("BottomLine2", "BottomLine2", 10093, 6804, 4989, 425, "NONE;»œÊ‰;STORE;«·«”„ «·„Œ ’— ··„Õ·;NAME;«”„ «·„‰ Ã;PRICE;«·”⁄—;CODE;ﬂÊœ «·„‰ Ã;BARCODE;—ﬁ„ «·»«—ﬂÊœ", 2, "0;3402")
    Set c = AddLabel("lblBottomLine2", "«·”ÿ— «·À«‰Ì √”›· «·»«—ﬂÊœ", 7768, 6804, 2268, 425, 10, False, CLR_MUTED, "BottomLine2", 0)
    Set c = AddText("ShortName", "ShortName", 2552, 7371, 4989, 425)
    SetCtlProp c, "ControlTipText", "„À«·: «·‰Œ»…. ›«—€ = «”„ «·„Õ· „‰ «·≈⁄œ«œ« "
    SetCtlProp c, "StatusBarText", "„À«·: «·‰Œ»…. ›«—€ = «”„ «·„Õ· „‰ «·≈⁄œ«œ« "
    Set c = AddLabel("lblShortName", "«·«”„ «·„Œ ’— ··„Õ·", 227, 7371, 2268, 425, 10, False, CLR_MUTED, "ShortName", 0)
    Set c = AddText("FontSize", "FontSize", 10093, 7371, 4989, 425)
    Set c = AddLabel("lblFontSize", "ÕÃ„ «·Œÿ", 7768, 7371, 2268, 425, 10, False, CLR_MUTED, "FontSize", 0)
    Set c = AddLabel("lblStatus", " ", 227, 8051, 14855, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnCurrent = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    m_frm.OnError = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FormLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Current()" & vbCrLf
    s = s & "    FormCurrent Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormBeforeUpdate(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    FormAfterUpdate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Error(DataErr As Integer, Response As Integer)" & vbCrLf
    s = s & "    Response = FormError(Me, DataErr)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    FormKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not FormUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    FormAction Me, ""SAVE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    FormAction Me, ""UNDO""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    FormAction Me, ""CLOSE""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmLabelSettings", s
    Exit Sub
EH:
    AbortForm "frmLabelSettings", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSearch()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSearch", "«·»ÕÀ «·„ ﬁœ„", "", 15309, 8675, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE721), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·»ÕÀ «·„ ﬁœ„", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«»ÕÀ »«·«”„ √Ê «·ﬂÊœ √Ê «·»«—ﬂÊœ √Ê —ﬁ„ «·›« Ê—… √Ê «·ÃÊ«·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboKind", "", 227, 1389, 2381, 454, "PRODUCT;«·„‰ Ã« ;CUSTOMER;«·⁄„·«¡;SUPPLIER;«·„Ê—œÊ‰;SALE;›Ê« Ì— «·»Ì⁄;PURCHASE;›Ê« Ì— «·‘—«¡", 2, "0;2268")
    SetCtlProp c, "DefaultValue", """PRODUCT"""
    c.AfterUpdate = EP
    Set c = AddLabel("lblKind", "«»ÕÀ ›Ì", 227, 1077, 2381, 284, 9, False, CLR_MUTED, "cboKind", 0)
    Set c = AddText("txtText", "", 2778, 1389, 4876, 454)
    c.AfterUpdate = EP
    Set c = AddLabel("lblText", "ﬂ·„… «·»ÕÀ («”„° ﬂÊœ° »«—ﬂÊœ° —ﬁ„° ÃÊ«·)", 2778, 1077, 4876, 284, 9, False, CLR_MUTED, "txtText", 0)
    Set c = AddText("txtFrom", "", 7825, 1389, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰  «—ÌŒ", 7825, 1077, 1814, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 9809, 1389, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï  «—ÌŒ", 9809, 1077, 1814, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnSearch", "»ÕÀ", 11794, 1372, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClear", "„”Õ", 13268, 1372, 907, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 14175, 1372, 907, 482, "secondary")
    c.OnClick = EP
    Set c = AddList("lstResults", 227, 2098, 14855, 5897, 6, "", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblCount", " ", 227, 8136, 14855, 340, 10, False, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    SearchLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboKind_AfterUpdate()" & vbCrLf
    s = s & "    SearchKindChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtText_AfterUpdate()" & vbCrLf
    s = s & "    RunSearch Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstResults_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    SearchOpen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSearch_Click()" & vbCrLf
    s = s & "    RunSearch Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    SearchClear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;btnSearch,11794,1372,1361,482,1000,0,0,0;btnClear,13268,1372,907,482,1000,0,0,0;btnClose,14175,1372,907,482,1000,0,0,0;lstResults,227,2098,14855,5897,0,1000,0,1000;lblCount,227,8136,14855,340,0,1000,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 8675, -4422, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSearch", s
    Exit Sub
EH:
    AbortForm "frmSearch", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmReportCenter()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmReportCenter", "«· ﬁ«—Ì—", "", 15309, 8675, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "„—ﬂ“ «· ﬁ«—Ì—", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «· ﬁ—Ì— À„ Õœœ «·› —… √Ê «·⁄„Ì· √Ê «·„‰ Ã", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddList("lstReports", 227, 1077, 5103, 7314, 2, "0;4876", False)
    c.RowSourceType = "Value List"
    c.FontSize = 11
    c.AfterUpdate = EP
    c.OnDblClick = EP
    Set c = AddLabel("lblReportTitle", " ", 5670, 1077, 9412, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblNeeds", " ", 5670, 1588, 9412, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddText("txtFrom", "", 5670, 2325, 2268, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰  «—ÌŒ", 5670, 2013, 2268, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 8165, 2325, 2268, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï  «—ÌŒ", 8165, 2013, 2268, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnToday", "«·ÌÊ„", 5670, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisMonth", "Â–« «·‘Â—", 7428, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastMonth", "«·‘Â— «·„«÷Ì", 9186, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisYear", "Â–Â «·”‰…", 10944, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboCustomer", "", 5670, 3912, 5103, 454, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5103")
    Set c = AddLabel("lblCustomer", "«·⁄„Ì·", 5670, 3600, 5103, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddCombo("cboSupplier", "", 5670, 4734, 5103, 454, "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", 2, "0;5103")
    Set c = AddLabel("lblSupplier", "«·„Ê—œ", 5670, 4422, 5103, 284, 9, False, CLR_MUTED, "cboSupplier", 0)
    Set c = AddCombo("cboProduct", "", 5670, 5556, 5103, 454, "SELECT ProductID, ProductName & ' (' & ProductCode & ')' AS Item FROM Products ORDER BY ProductName", 2, "0;5103")
    Set c = AddLabel("lblProduct", "«·„‰ Ã", 5670, 5244, 5103, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddButton("btnRun", "⁄—÷ «· ﬁ—Ì—", 5670, 6548, 2835, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPdf", "Õ›Ÿ PDF", 8618, 6548, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnExcel", " ’œÌ— Excel", 10433, 6548, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13948, 6548, 1134, 567, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblPhaseNote", "Ìı⁄—÷ «· ﬁ—Ì— ··„⁄«Ì‰… Ê„‰Â« «·ÿ»«⁄…. ´Õ›Ÿ PDFª Ê´ ’œÌ— Excelª ÌÕ›Ÿ«‰ «·„·› ›Ì „Ã·œ Reports »Ã«‰» „·› «·»—‰«„Ã.", 5670, 7342, 9412, 567, 9, False, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ReportCenterLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstReports_AfterUpdate()" & vbCrLf
    s = s & "    ReportSelected Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstReports_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    RunReport Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnToday_Click()" & vbCrLf
    s = s & "    SetQuickPeriod Me, ""TODAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisMonth_Click()" & vbCrLf
    s = s & "    SetQuickPeriod Me, ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastMonth_Click()" & vbCrLf
    s = s & "    SetQuickPeriod Me, ""LASTMONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisYear_Click()" & vbCrLf
    s = s & "    SetQuickPeriod Me, ""YEAR""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRun_Click()" & vbCrLf
    s = s & "    RunReport Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPdf_Click()" & vbCrLf
    s = s & "    ExportReport Me, ""PDF""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnExcel_Click()" & vbCrLf
    s = s & "    ExportReport Me, ""XLSX""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;lstReports,227,1077,5103,7314,0,0,0,1000;lblReportTitle,5670,1077,9412,482,0,1000,0,0;lblNeeds,5670,1588,9412,340,0,1000,0,0;btnClose,13948,6548,1134,567,1000,0,0,0;lblPhaseNote,5670,7342,9412,567,0,1000,0,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 8675, -709, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmReportCenter", s
    Exit Sub
EH:
    AbortForm "frmReportCenter", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPOSLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPOSLines", "√”ÿ— «·›« Ê—…", "SELECT * FROM tmpPOSLines ORDER BY LineNo", 12020, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("LineNo", "LineNo", 28, 0, 567, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("ProductName", "ProductName", 623, 0, 4252, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("Quantity", "Quantity", 4903, 0, 1247, 425)
    SetCtlProp c, "Format", "#,##0.###"
    c.AfterUpdate = EP
    Set c = AddText("UnitPrice", "UnitPrice", 6178, 0, 1474, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("LineDiscount", "LineDiscount", 7680, 0, 1247, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("LineTotal", "LineTotal", 8955, 0, 1588, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Available", "Available", 10571, 0, 907, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddButton("btnRemove", "Sym(code='ChrW(&HE74D)')", 11506, 17, 454, 391, "danger")
    SetCtlProp c, "FontName", ICON_FONT
    c.OnClick = EP
    s = ""
    s = s & "Private Sub Quantity_AfterUpdate()" & vbCrLf
    s = s & "    LineChanged Me, ""Quantity""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub UnitPrice_AfterUpdate()" & vbCrLf
    s = s & "    LineChanged Me, ""UnitPrice""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub LineDiscount_AfterUpdate()" & vbCrLf
    s = s & "    LineChanged Me, ""LineDiscount""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    RemoveCurrentLine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPOSLines", s
    Exit Sub
EH:
    AbortForm "frmPOSLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPOS()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPOS", "‰ﬁÿ… «·»Ì⁄", "", 18994, 8675, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "‰ﬁÿ… «·»Ì⁄", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "F9 Õ›Ÿ  |  F12 Õ›Ÿ Êÿ»«⁄…  |  F5 ›« Ê—… ÃœÌœ…  |  F4 »ÕÀ »«·«”„  |  F8 «·„»·€ «·„œ›Ê⁄  |  F2 «·»«—ﬂÊœ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtBarcode", "", 227, 1304, 4536, 567)
    c.FontSize = 16
    c.OnKeyDown = EP
    Set c = AddLabel("lblBarcode", "«·»«—ﬂÊœ √Ê ﬂÊœ «·„‰ Ã (Enter)", 227, 992, 4536, 284, 9, False, CLR_MUTED, "txtBarcode", 0)
    Set c = AddText("txtQty", "", 4933, 1304, 1134, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "DefaultValue", "1"
    Set c = AddLabel("lblQty", "«·ﬂ„Ì…", 4933, 992, 1134, 284, 9, False, CLR_MUTED, "txtQty", 0)
    Set c = AddCombo("cboProduct", "", 6237, 1304, 6010, 567, "SELECT ProductID, ProductName & ' - ' & ProductCode AS Item, SellingPrice FROM Products WHERE IsActive = True ORDER BY ProductName", 3, "0;4536;1134")
    c.FontSize = 13
    c.AfterUpdate = EP
    Set c = AddLabel("lblProduct", "√Ê «»ÕÀ »«”„ «·„‰ Ã (F4)", 6237, 992, 6010, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddLabel("lblCol1", "#", 255, 2041, 567, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·’‰›", 850, 2041, 4252, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·ﬂ„Ì…", 5130, 2041, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·”⁄—", 6405, 2041, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "Œ’„ «·”ÿ—", 7907, 2041, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·≈Ã„«·Ì", 9182, 2041, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "«·„ Ê›—", 10798, 2041, 907, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmPOSLines", 227, 2410, 12020, 4706)
    Set c = AddLabel("lblStatus", " ", 227, 7229, 12020, 397, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ (F9)", 227, 7881, 1928, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "Õ›Ÿ Êÿ»«⁄… (F12)", 2268, 7881, 2381, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnNewSale", "›« Ê—… ÃœÌœ… (F5)", 4762, 7881, 2155, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "„— Ã⁄", 7030, 7881, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "”‰œ ﬁ»÷", 8617, 7881, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReprint", "≈⁄«œ… ÿ»«⁄…", 10204, 7881, 1588, 624, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboCustomer", "", 12474, 1304, 6294, 454, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustomer", "«·⁄„Ì·", 12474, 992, 6294, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddLabel("lblCustomerInfo", " ", 12474, 1786, 6294, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddCombo("cboPaymentType", "", 12474, 2438, 3033, 454, "CASH;‰ﬁœÌ;CREDIT;¬Ã·", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentType", "‰Ê⁄ «·»Ì⁄", 12474, 2126, 3033, 284, 9, False, CLR_MUTED, "cboPaymentType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 15735, 2438, 3033, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 15735, 2126, 3033, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtInvoiceDiscount", "", 12474, 3260, 3033, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceDiscount", "Œ’„ ⁄·Ï «·›« Ê—…", 12474, 2948, 3033, 284, 9, False, CLR_MUTED, "txtInvoiceDiscount", 0)
    Set c = AddText("txtNotes", "", 15735, 3260, 3033, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 15735, 2948, 3033, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddRect("boxTotals", 12474, 3884, 6294, 2665, CLR_SURFACE)
    Set c = AddLabel("lblCapSubTotal", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„ Ê«·÷—Ì»…", 12644, 3941, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblSubTotal", "0.00", 16386, 3941, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapDiscount", "«·Œ’„", 12644, 4281, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblDiscount", "0.00", 16386, 4281, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 12644, 4621, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTax", "0.00", 16386, 4621, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTotal", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 12644, 4990, 5954, 340, 12, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTotal", "0.00", 12644, 5330, 5954, 822, 34, True, CLR_ACCENT, "", 2)
    Set c = AddLabel("lblItems", " ", 12644, 6180, 5954, 312, 10, False, CLR_MUTED, "", 2)
    Set c = AddText("txtTendered", "", 12474, 6889, 6294, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTendered", "«·„»·€ «·„œ›Ê⁄ (F8) - « —ﬂÂ ›«—€« ≈–« œ›⁄ «·„»·€ »«·÷»ÿ", 12474, 6577, 6294, 284, 9, False, CLR_MUTED, "txtTendered", 0)
    Set c = AddLabel("lblChange", " ", 12474, 7484, 6294, 369, 14, True, CLR_SUCCESS, "", 2)
    Set c = AddLabel("lblLastInvoiceCap", "¬Œ— ›« Ê—…:", 12474, 7995, 1701, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblLastInvoice", " ", 14232, 7995, 2778, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17067, 7881, 1701, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    POSLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    POSKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not POSUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    BarcodeKeyDown Me, KeyCode" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboProduct_AfterUpdate()" & vbCrLf
    s = s & "    ProductPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCustomer_AfterUpdate()" & vbCrLf
    s = s & "    CustomerChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboPaymentType_AfterUpdate()" & vbCrLf
    s = s & "    PaymentTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboPaymentMethod_AfterUpdate()" & vbCrLf
    s = s & "    PaymentMethodChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtInvoiceDiscount_AfterUpdate()" & vbCrLf
    s = s & "    RecalcPOS Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtTendered_AfterUpdate()" & vbCrLf
    s = s & "    RecalcPOS Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SavePOS Me, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSavePrint_Click()" & vbCrLf
    s = s & "    SavePOS Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewSale_Click()" & vbCrLf
    s = s & "    NewSale Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturn_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSalesReturn"", 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayment_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomerPayment"", 6, Me!cboCustomer.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReprint_Click()" & vbCrLf
    s = s & "    ReprintLast Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;cboProduct,6237,1304,6010,567,0,1000,0,0;subLines,227,2410,12020,4706,0,1000,0,1000;lblStatus,227,7229,12020,397,0,1000,1000,0;btnSave,227,7881,1928,624,0,0,1000,0;btnSavePrint,2268,7881,2381,624,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnNewSale,4762,7881,2155,624,0,0,1000,0;btnReturn,7030,7881,1474,624,0,0,1000,0;btnPayment,8617,7881,1474,624,0,0,1000,0;btnReprint,10204,7881,1588,624,0,0,1000,0;cboCustomer,12474,1304,6294,454,1000,0,0,0;lblCustomer,12474,992,6294,284,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCustomerInfo,12474,1786,6294,284,1000,0,0,0;cboPaymentType,12474,2438,3033,454,1000,0,0,0;lblPaymentType,12474,2126,3033,284,1000,0,0,0;cboPaymentMethod,15735,2438,3033,454,1000,0,0,0;lblPaymentMethod,15735,2126,3033,284,1000,0,0,0;txtInvoiceDiscount,12474,3260,3033,454,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblInvoiceDiscount,12474,2948,3033,284,1000,0,0,0;txtNotes,15735,3260,3033,454,1000,0,0,0;lblNotes,15735,2948,3033,284,1000,0,0,0;boxTotals,12474,3884,6294,2665,1000,0,0,0;lblCapSubTotal,12644,3941,3686,340,1000,0,0,0;lblSubTotal,16386,3941,2211,340,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCapDiscount,12644,4281,3686,340,1000,0,0,0;lblDiscount,16386,4281,2211,340,1000,0,0,0;lblCapTax,12644,4621,3686,340,1000,0,0,0;lblTax,16386,4621,2211,340,1000,0,0,0;lblCapTotal,12644,4990,5954,340,1000,0,0,0;lblTotal,12644,5330,5954,822,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblItems,12644,6180,5954,312,1000,0,0,0;txtTendered,12474,6889,6294,567,1000,0,1000,0;lblTendered,12474,6577,6294,284,1000,0,1000,0;lblChange,12474,7484,6294,369,1000,0,1000,0;lblLastInvoiceCap,12474,7995,1701,312,1000,0,1000,0;lblLastInvoice,14232,7995,2778,312,1000,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnClose,17067,7881,1701,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 8675, 0, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPOS", s
    Exit Sub
EH:
    AbortForm "frmPOS", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmReturnLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmReturnLines", "√”ÿ— «·„— Ã⁄", "SELECT * FROM tmpReturnLines ORDER BY SalesDetailID", 14855, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("ProductName", "ProductName", 28, 0, 5103, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("SoldQty", "SoldQty", 5159, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("ReturnedQty", "ReturnedQty", 6661, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("AvailableQty", "AvailableQty", 8163, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("ReturnQty", "ReturnQty", 9665, 0, 1588, 425)
    c.FontBold = True
    SetCtlProp c, "Format", "#,##0.###"
    c.AfterUpdate = EP
    Set c = AddCheck("ReturnToStock", "ReturnToStock", 11876, 68)
    c.AfterUpdate = EP
    Set c = AddText("ReturnAmount", "ReturnAmount", 12783, 0, 1928, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    s = ""
    s = s & "Private Sub ReturnQty_AfterUpdate()" & vbCrLf
    s = s & "    ReturnLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub ReturnToStock_AfterUpdate()" & vbCrLf
    s = s & "    ReturnLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmReturnLines", s
    Exit Sub
EH:
    AbortForm "frmReturnLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSalesReturn()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSalesReturn", "„— Ã⁄ „»Ì⁄« ", "", 15309, 9639, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "„— Ã⁄ „»Ì⁄«  (≈‘⁄«— œ«∆‰)", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «·›« Ê—… «·√’·Ì… À„ Õœœ «·ﬂ„Ì«  «·„— Ã⁄…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtInvoiceNo", "", 227, 1304, 2835, 454)
    c.FontSize = 12
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceNo", "—ﬁ„ «·›« Ê—… «·√’·Ì…", 227, 992, 2835, 284, 9, False, CLR_MUTED, "txtInvoiceNo", 0)
    Set c = AddButton("btnFind", "»ÕÀ", 3175, 1304, 1134, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblInvoiceInfo", " ", 4479, 1361, 10603, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddLabel("lblCol1", "«·’‰›", 255, 1956, 5103, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·ﬂ„Ì… «·„»«⁄…", 5386, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "„— Ã⁄ ”«»ﬁ«", 6888, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·„ «Õ ··≈—Ã«⁄", 8390, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "«·ﬂ„Ì… «·„— Ã⁄…", 9892, 1956, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "Ì⁄Êœ ··„Œ“Ê‰", 11508, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "ﬁÌ„… «·„— Ã⁄", 13010, 1956, 1928, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subReturnLines", "frmReturnLines", 227, 2296, 14855, 4309)
    Set c = AddText("txtReason", "", 227, 7002, 6577, 454)
    Set c = AddLabel("lblReason", "”»» «·≈—Ã«⁄ *", 227, 6690, 6577, 284, 9, False, CLR_MUTED, "txtReason", 0)
    Set c = AddCombo("cboRefundType", "", 6974, 7002, 3742, 454, "CASH;—œ ‰ﬁœÌ;CREDIT;Œ’„ „‰ —’Ìœ «·⁄„Ì·", 2, "0;3402")
    Set c = AddLabel("lblRefundType", "ÿ—Ìﬁ… —œ «·„»·€", 6974, 6690, 3742, 284, 9, False, CLR_MUTED, "cboRefundType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 10886, 7002, 4196, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·—œ", 10886, 6690, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddLabel("lblReturnTotalCap", "ﬁÌ„… «·„— Ã⁄:", 227, 7711, 2268, 454, 13, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblReturnTotal", "0.00", 2552, 7598, 3402, 624, 20, True, CLR_DANGER, "", 0)
    Set c = AddButton("btnReturnAll", "≈—Ã«⁄ «·ﬂ·", 227, 8675, 1814, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturn", "Õ›Ÿ «·„— Ã⁄", 2154, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturnPrint", "Õ›Ÿ Êÿ»«⁄…", 4308, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 8675, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ReturnLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtInvoiceNo_AfterUpdate()" & vbCrLf
    s = s & "    ReturnFindInvoice Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFind_Click()" & vbCrLf
    s = s & "    ReturnFindInvoice Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturnAll_Click()" & vbCrLf
    s = s & "    ReturnAll Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveReturn_Click()" & vbCrLf
    s = s & "    SaveReturn Me, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveReturnPrint_Click()" & vbCrLf
    s = s & "    SaveReturn Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSalesReturn", s
    Exit Sub
EH:
    AbortForm "frmSalesReturn", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCustomerPayment()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCustomerPayment", "”‰œ ﬁ»÷", "", 9072, 5783, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "”‰œ ﬁ»÷", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " ”ÃÌ· œ›⁄… „‰ ⁄„Ì· ⁄·Ï Õ”«»Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboCustomer", "", 227, 1361, 8618, 454, "SELECT CustomerID, CustomerName FROM Customers WHERE IsSystem = False AND IsActive = True ORDER BY CustomerName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustomer", "«·⁄„Ì·", 227, 1049, 8618, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddLabel("lblBalance", " ", 227, 1871, 8618, 340, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtAmount", "", 227, 2608, 4196, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ *", 227, 2296, 4196, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboPaymentMethod", "", 4649, 2608, 4196, 510, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 4649, 2296, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtNotes", "", 227, 3515, 8618, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 227, 3203, 8618, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «·”‰œ", 227, 4763, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "Õ›Ÿ Êÿ»«⁄…", 2268, 4763, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 7371, 4763, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PaymentLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCustomer_AfterUpdate()" & vbCrLf
    s = s & "    PaymentCustomerChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SavePayment Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSavePrint_Click()" & vbCrLf
    s = s & "    SavePayment Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCustomerPayment", s
    Exit Sub
EH:
    AbortForm "frmCustomerPayment", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSalesInvoice()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSalesInvoice", "›« Ê—… »Ì⁄", "", 15309, 8845, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "›« Ê—… »Ì⁄", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "⁄—÷ ›ﬁÿ - «· ’ÕÌÕ ÌﬂÊ‰ »„— Ã⁄", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblHeader", " ", 227, 1021, 14855, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddList("lstLines", 227, 1531, 14855, 3742, 7, "567;5103;1418;1984;1418;1701;2268", True)
    Set c = AddLabel("lblReturnsCap", "«·„— Ã⁄«  ⁄·Ï Â–Â «·›« Ê—…", 227, 5386, 4536, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstReturns", 227, 5727, 14855, 1304, 4, "2268;2835;1984;6804", True)
    Set c = AddLabel("lblTotals", " ", 227, 7144, 14855, 397, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 227, 7881, 1588, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrintA4", "ÿ»«⁄… A4", 1928, 7881, 1588, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "„— Ã⁄", 3629, 7881, 1474, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 7881, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    InvoiceViewLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintSalesDocument ""SALE"", Me!txtInvoiceID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrintA4_Click()" & vbCrLf
    s = s & "    PrintSalesDocument ""SALE"", Me!txtInvoiceID.Value, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturn_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSalesReturn"", 6, Me!txtInvoiceID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSalesInvoice", s
    Exit Sub
EH:
    AbortForm "frmSalesInvoice", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPurchaseLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPurchaseLines", "√”ÿ— ›« Ê—… «·‘—«¡", "SELECT * FROM tmpPurchaseLines ORDER BY LineNo", 12020, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("LineNo", "LineNo", 28, 0, 510, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("ProductName", "ProductName", 566, 0, 3175, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("Quantity", "Quantity", 3769, 0, 1077, 425)
    SetCtlProp c, "Format", "#,##0.###"
    c.AfterUpdate = EP
    Set c = AddText("UnitCost", "UnitCost", 4874, 0, 1304, 425)
    SetCtlProp c, "Format", "#,##0.00##"
    c.AfterUpdate = EP
    Set c = AddText("LineDiscount", "LineDiscount", 6206, 0, 1077, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("LineTotal", "LineTotal", 7311, 0, 1418, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("SellingPrice", "SellingPrice", 8757, 0, 1247, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("NewSellingPrice", "NewSellingPrice", 10032, 0, 1361, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddButton("btnRemove", "Sym(code='ChrW(&HE74D)')", 11421, 17, 454, 391, "danger")
    SetCtlProp c, "FontName", ICON_FONT
    c.OnClick = EP
    s = ""
    s = s & "Private Sub Quantity_AfterUpdate()" & vbCrLf
    s = s & "    PurLineChanged Me, ""Quantity""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub UnitCost_AfterUpdate()" & vbCrLf
    s = s & "    PurLineChanged Me, ""UnitCost""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub LineDiscount_AfterUpdate()" & vbCrLf
    s = s & "    PurLineChanged Me, ""LineDiscount""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub NewSellingPrice_AfterUpdate()" & vbCrLf
    s = s & "    PurLineChanged Me, ""NewSellingPrice""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    PurRemoveLine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPurchaseLines", s
    Exit Sub
EH:
    AbortForm "frmPurchaseLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPurchaseInvoice()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPurchaseInvoice", "›« Ê—… „‘ —Ì« ", "", 18994, 9242, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE896), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "›« Ê—… „‘ —Ì« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«· ﬂ·›… »œÊ‰ ÷—Ì»…  |  F9 Õ›Ÿ  |  F5 ›« Ê—… ÃœÌœ…  |  F4 »ÕÀ »«·«”„  |  F2 «·»«—ﬂÊœ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtBarcode", "", 227, 1304, 4536, 567)
    c.FontSize = 16
    c.OnKeyDown = EP
    Set c = AddLabel("lblBarcode", "«·»«—ﬂÊœ √Ê ﬂÊœ «·„‰ Ã (Enter)", 227, 992, 4536, 284, 9, False, CLR_MUTED, "txtBarcode", 0)
    Set c = AddText("txtQty", "", 4933, 1304, 1134, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "DefaultValue", "1"
    Set c = AddLabel("lblQty", "«·ﬂ„Ì…", 4933, 992, 1134, 284, 9, False, CLR_MUTED, "txtQty", 0)
    Set c = AddCombo("cboProduct", "", 6237, 1304, 6010, 567, "SELECT ProductID, ProductName & ' - ' & ProductCode AS Item, PurchasePrice FROM Products WHERE IsActive = True ORDER BY ProductName", 3, "0;4536;1134")
    c.FontSize = 13
    c.AfterUpdate = EP
    Set c = AddLabel("lblProduct", "√Ê «»ÕÀ »«”„ «·„‰ Ã (F4)", 6237, 992, 6010, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddLabel("lblCol1", "#", 255, 2041, 510, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·’‰›", 793, 2041, 3175, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·ﬂ„Ì…", 3996, 2041, 1077, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", " ﬂ·›… «·ÊÕœ…", 5101, 2041, 1304, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "«·Œ’„", 6433, 2041, 1077, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·≈Ã„«·Ì", 7538, 2041, 1418, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "”⁄— «·»Ì⁄", 8984, 2041, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol8", "”⁄— »Ì⁄ ÃœÌœ", 10259, 2041, 1361, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmPurchaseLines", 227, 2410, 12020, 5273)
    Set c = AddLabel("lblStatus", " ", 227, 7796, 12020, 397, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ (F9)", 227, 8448, 1814, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnNewInvoice", "›« Ê—… ÃœÌœ… (F5)", 2154, 8448, 2155, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "„— Ã⁄ „‘ —Ì« ", 4422, 8448, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "”‰œ ’—›", 6463, 8448, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNewProduct", "„‰ Ã ÃœÌœ", 8050, 8448, 1588, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastInvoice", "¬Œ— ›« Ê—…", 9751, 8448, 1588, 624, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboSupplier", "", 12474, 1304, 6294, 454, "SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive = True ORDER BY SupplierName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblSupplier", "«·„Ê—œ *", 12474, 992, 6294, 284, 9, False, CLR_MUTED, "cboSupplier", 0)
    Set c = AddLabel("lblSupplierInfo", " ", 12474, 1786, 6294, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("txtSupplierInvoiceNo", "", 12474, 2438, 3033, 454)
    Set c = AddLabel("lblSupplierInvoiceNo", "—ﬁ„ ›« Ê—… «·„Ê—œ", 12474, 2126, 3033, 284, 9, False, CLR_MUTED, "txtSupplierInvoiceNo", 0)
    Set c = AddText("txtInvoiceDate", "", 15735, 2438, 3033, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblInvoiceDate", " «—ÌŒ «·›« Ê—…", 15735, 2126, 3033, 284, 9, False, CLR_MUTED, "txtInvoiceDate", 0)
    Set c = AddCombo("cboPaymentType", "", 12474, 3260, 3033, 454, "CASH;‰ﬁœÌ;CREDIT;¬Ã·", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentType", "‰Ê⁄ «·‘—«¡", 12474, 2948, 3033, 284, 9, False, CLR_MUTED, "cboPaymentType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 15735, 3260, 3033, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;2835")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 15735, 2948, 3033, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtInvoiceDiscount", "", 12474, 4082, 3033, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceDiscount", "Œ’„ ⁄·Ï «·›« Ê—… (»œÊ‰ ÷—Ì»…)", 12474, 3770, 3033, 284, 9, False, CLR_MUTED, "txtInvoiceDiscount", 0)
    Set c = AddCheck("chkChargeVAT", "", 15735, 4167)
    c.AfterUpdate = EP
    Set c = AddLabel("lblChargeVAT", "«·„Ê—œ ÌÕ ”» «·÷—Ì»…", 16104, 4082, 2664, 454, 10, False, CLR_TEXT, "chkChargeVAT", 0)
    Set c = AddText("txtNotes", "", 12474, 4905, 6294, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 12474, 4593, 6294, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddRect("boxTotals", 12474, 5443, 6294, 2126, CLR_SURFACE)
    Set c = AddLabel("lblCapSubTotal", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„ Ê«·÷—Ì»…", 12644, 5500, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblSubTotal", "0.00", 16386, 5500, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapDiscount", "«·Œ’„", 12644, 5840, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblDiscount", "0.00", 16386, 5840, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTax", "÷—Ì»… «·„œŒ·« ", 12644, 6180, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTax", "0.00", 16386, 6180, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTotal", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 12644, 6606, 2835, 340, 12, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTotal", "0.00", 15536, 6520, 3062, 567, 24, True, CLR_ACCENT, "", 3)
    Set c = AddLabel("lblItems", " ", 12644, 7173, 5954, 340, 10, False, CLR_MUTED, "", 2)
    Set c = AddText("txtPaid", "", 12474, 7938, 3033, 510)
    c.FontSize = 13
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaid", "«·„œ›Ê⁄ ··„Ê—œ «·¬‰", 12474, 7626, 3033, 284, 9, False, CLR_MUTED, "txtPaid", 0)
    Set c = AddLabel("lblRemaining", " ", 15735, 7966, 3033, 454, 12, True, CLR_WARNING, "", 0)
    Set c = AddButton("btnLabels", "ÿ»«⁄… «·»«—ﬂÊœ", 12474, 8448, 2268, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17067, 8448, 1701, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PurchaseLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    PurchaseKeyDown Me, KeyCode, Shift" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not PurchaseUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    PurBarcodeKeyDown Me, KeyCode" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboProduct_AfterUpdate()" & vbCrLf
    s = s & "    PurProductPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboSupplier_AfterUpdate()" & vbCrLf
    s = s & "    PurSupplierChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboPaymentType_AfterUpdate()" & vbCrLf
    s = s & "    PurPaymentTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtInvoiceDiscount_AfterUpdate()" & vbCrLf
    s = s & "    RecalcPurchase Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkChargeVAT_AfterUpdate()" & vbCrLf
    s = s & "    RecalcPurchase Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtPaid_AfterUpdate()" & vbCrLf
    s = s & "    RecalcPurchase Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SavePurchase Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewInvoice_Click()" & vbCrLf
    s = s & "    NewPurchase Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturn_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseReturn"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayment_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSupplierPayment"", 7, Me!cboSupplier.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewProduct_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmProducts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastInvoice_Click()" & vbCrLf
    s = s & "    PurShowLast Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLabels_Click()" & vbCrLf
    s = s & "    LabelsFromPurchaseScreen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;cboProduct,6237,1304,6010,567,0,1000,0,0;subLines,227,2410,12020,5273,0,1000,0,1000;lblStatus,227,7796,12020,397,0,1000,1000,0;btnSave,227,8448,1814,624,0,0,1000,0;btnNewInvoice,2154,8448,2155,624,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnReturn,4422,8448,1928,624,0,0,1000,0;btnPayment,6463,8448,1474,624,0,0,1000,0;btnNewProduct,8050,8448,1588,624,0,0,1000,0;btnLastInvoice,9751,8448,1588,624,0,0,1000,0;cboSupplier,12474,1304,6294,454,1000,0,0,0;lblSupplier,12474,992,6294,284,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblSupplierInfo,12474,1786,6294,284,1000,0,0,0;txtSupplierInvoiceNo,12474,2438,3033,454,1000,0,0,0;lblSupplierInvoiceNo,12474,2126,3033,284,1000,0,0,0;txtInvoiceDate,15735,2438,3033,454,1000,0,0,0;lblInvoiceDate,15735,2126,3033,284,1000,0,0,0;cboPaymentType,12474,3260,3033,454,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblPaymentType,12474,2948,3033,284,1000,0,0,0;cboPaymentMethod,15735,3260,3033,454,1000,0,0,0;lblPaymentMethod,15735,2948,3033,284,1000,0,0,0;txtInvoiceDiscount,12474,4082,3033,454,1000,0,0,0;lblInvoiceDiscount,12474,3770,3033,284,1000,0,0,0;chkChargeVAT,15735,4167,284,284,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblChargeVAT,16104,4082,2664,454,1000,0,0,0;txtNotes,12474,4905,6294,454,1000,0,0,0;lblNotes,12474,4593,6294,284,1000,0,0,0;boxTotals,12474,5443,6294,2126,1000,0,0,0;lblCapSubTotal,12644,5500,3686,340,1000,0,0,0;lblSubTotal,16386,5500,2211,340,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCapDiscount,12644,5840,3686,340,1000,0,0,0;lblDiscount,16386,5840,2211,340,1000,0,0,0;lblCapTax,12644,6180,3686,340,1000,0,0,0;lblTax,16386,6180,2211,340,1000,0,0,0;lblCapTotal,12644,6606,2835,340,1000,0,0,0;lblTotal,15536,6520,3062,567,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblItems,12644,7173,5954,340,1000,0,0,0;txtPaid,12474,7938,3033,510,1000,0,1000,0;lblPaid,12474,7626,3033,284,1000,0,1000,0;lblRemaining,15735,7966,3033,454,1000,0,1000,0;btnLabels,12474,8448,2268,624,1000,0,1000,0;btnClose,17067,8448,1701,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9242, 0, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPurchaseInvoice", s
    Exit Sub
EH:
    AbortForm "frmPurchaseInvoice", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPurchaseReturnLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPurchaseReturnLines", "√”ÿ— „— Ã⁄ «·„‘ —Ì« ", "SELECT * FROM tmpPurchaseReturnLines ORDER BY PurchaseDetailID", 14855, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("ProductName", "ProductName", 28, 0, 4536, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("BoughtQty", "BoughtQty", 4592, 0, 1418, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("ReturnedQty", "ReturnedQty", 6038, 0, 1418, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("AvailableQty", "AvailableQty", 7484, 0, 1418, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("InStock", "InStock", 8930, 0, 1418, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("ReturnQty", "ReturnQty", 10376, 0, 1588, 425)
    c.FontBold = True
    SetCtlProp c, "Format", "#,##0.###"
    c.AfterUpdate = EP
    Set c = AddText("ReturnAmount", "ReturnAmount", 11992, 0, 1928, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    s = ""
    s = s & "Private Sub ReturnQty_AfterUpdate()" & vbCrLf
    s = s & "    PurReturnLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPurchaseReturnLines", s
    Exit Sub
EH:
    AbortForm "frmPurchaseReturnLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPurchaseReturn()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPurchaseReturn", "„— Ã⁄ „‘ —Ì« ", "", 15309, 9639, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE896), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "„— Ã⁄ „‘ —Ì« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — ›« Ê—… «·‘—«¡ À„ Õœœ «·ﬂ„Ì«  «· Ì  ⁄Êœ ··„Ê—œ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtInvoiceNo", "", 227, 1304, 2835, 454)
    c.FontSize = 12
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceNo", "—ﬁ„‰« (PUR-) √Ê —ﬁ„ ›« Ê—… «·„Ê—œ", 227, 992, 2835, 284, 9, False, CLR_MUTED, "txtInvoiceNo", 0)
    Set c = AddButton("btnFind", "»ÕÀ", 3175, 1304, 1134, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblInvoiceInfo", " ", 4479, 1361, 10603, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddLabel("lblCol1", "«·’‰›", 255, 1956, 4536, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·ﬂ„Ì… «·„‘ —«…", 4819, 1956, 1418, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "„— Ã⁄ ”«»ﬁ«", 6265, 1956, 1418, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·„ «Õ ··≈—Ã«⁄", 7711, 1956, 1418, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "«·—’Ìœ «·Õ«·Ì", 9157, 1956, 1418, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·ﬂ„Ì… «·„— Ã⁄…", 10603, 1956, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "ﬁÌ„… «·„— Ã⁄", 12219, 1956, 1928, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subReturnLines", "frmPurchaseReturnLines", 227, 2296, 14855, 4309)
    Set c = AddText("txtReason", "", 227, 7002, 6577, 454)
    Set c = AddLabel("lblReason", "”»» «·≈—Ã«⁄ *", 227, 6690, 6577, 284, 9, False, CLR_MUTED, "txtReason", 0)
    Set c = AddCombo("cboRefundType", "", 6974, 7002, 3742, 454, "CREDIT;Œ’„ „‰ —’Ìœ «·„Ê—œ;CASH;«” —œ«œ ‰ﬁœÌ „‰ «·„Ê—œ", 2, "0;3402")
    Set c = AddLabel("lblRefundType", "ÿ—Ìﬁ… «·«” —œ«œ", 6974, 6690, 3742, 284, 9, False, CLR_MUTED, "cboRefundType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 10886, 7002, 4196, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·«” ·«„", 10886, 6690, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddLabel("lblReturnTotalCap", "ﬁÌ„… «·„— Ã⁄:", 227, 7711, 2268, 454, 13, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblReturnTotal", "0.00", 2552, 7598, 3402, 624, 20, True, CLR_DANGER, "", 0)
    Set c = AddButton("btnReturnAll", "≈—Ã«⁄ «·ﬂ·", 227, 8675, 1814, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturn", "Õ›Ÿ «·„— Ã⁄", 2154, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturnPrint", "Õ›Ÿ Êÿ»«⁄…", 4308, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 8675, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PurReturnLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtInvoiceNo_AfterUpdate()" & vbCrLf
    s = s & "    PurReturnFind Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFind_Click()" & vbCrLf
    s = s & "    PurReturnFind Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturnAll_Click()" & vbCrLf
    s = s & "    PurReturnAll Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveReturn_Click()" & vbCrLf
    s = s & "    SavePurchaseReturn Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveReturnPrint_Click()" & vbCrLf
    s = s & "    SavePurchaseReturn Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPurchaseReturn", s
    Exit Sub
EH:
    AbortForm "frmPurchaseReturn", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSupplierPayment()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSupplierPayment", "”‰œ ’—›", "", 9072, 5783, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE77B), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "”‰œ ’—›", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " ”ÃÌ· œ›⁄… ·„Ê—œ „‰ Õ”«»Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboSupplier", "", 227, 1361, 8618, 454, "SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive = True ORDER BY SupplierName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblSupplier", "«·„Ê—œ", 227, 1049, 8618, 284, 9, False, CLR_MUTED, "cboSupplier", 0)
    Set c = AddLabel("lblBalance", " ", 227, 1871, 8618, 340, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtAmount", "", 227, 2608, 4196, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ *", 227, 2296, 4196, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboPaymentMethod", "", 4649, 2608, 4196, 510, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 4649, 2296, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtNotes", "", 227, 3515, 8618, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 227, 3203, 8618, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «·”‰œ", 227, 4763, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "Õ›Ÿ Êÿ»«⁄…", 2268, 4763, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 7371, 4763, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    SupplierPaymentLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboSupplier_AfterUpdate()" & vbCrLf
    s = s & "    SupplierPaymentChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveSupplierPayment Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSavePrint_Click()" & vbCrLf
    s = s & "    SaveSupplierPayment Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmSupplierPayment", s
    Exit Sub
EH:
    AbortForm "frmSupplierPayment", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPurchaseView()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPurchaseView", "›« Ê—… ‘—«¡", "", 15309, 8845, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE896), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "›« Ê—… ‘—«¡", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "⁄—÷ ›ﬁÿ - «· ’ÕÌÕ ÌﬂÊ‰ »„— Ã⁄ „‘ —Ì« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtSupplierID", "", 14345, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblHeader", " ", 227, 1021, 14855, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddList("lstLines", 227, 1531, 14855, 3742, 7, "567;5103;1418;1984;1418;1701;2268", True)
    Set c = AddLabel("lblReturnsCap", "«·„— Ã⁄«  ⁄·Ï Â–Â «·›« Ê—… (‰ﬁ— „“œÊÃ ··ÿ»«⁄…)", 227, 5386, 6804, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstReturns", 227, 5727, 14855, 1304, 5, "0;2268;2835;1984;6804", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblTotals", " ", 227, 7144, 14855, 397, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 227, 7881, 1474, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "„— Ã⁄", 1814, 7881, 1474, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "”‰œ ’—›", 3401, 7881, 1588, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLabels", "ÿ»«⁄… «·»«—ﬂÊœ", 5102, 7881, 1928, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 7881, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PurchaseViewLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstReturns_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    PrintPurchaseDocument ""RETURN"", Me!lstReturns.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintPurchaseDocument ""PURCHASE"", Me!txtInvoiceID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReturn_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseReturn"", 7, Me!txtInvoiceID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayment_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSupplierPayment"", 7, Me!txtSupplierID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLabels_Click()" & vbCrLf
    s = s & "    LabelsForPurchase Me!txtInvoiceID.Value, Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPurchaseView", s
    Exit Sub
EH:
    AbortForm "frmPurchaseView", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmInventory()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmInventory", "«·„Œ“Ê‰", "", 18994, 9355, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7B8), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„Œ“Ê‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "√—’œ… «·„‰ Ã«  ÊÕ—ﬂ« Â«° Ê«·—’Ìœ «·«›  «ÕÌ Ê«·≈÷«›… Ê«·Œ’„ «·ÌœÊÌ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtSearch", "", 227, 1304, 3969, 454)
    c.AfterUpdate = EP
    Set c = AddLabel("lblSearch", "»ÕÀ »«·«”„ √Ê «·ﬂÊœ √Ê «·»«—ﬂÊœ", 227, 992, 3969, 284, 9, False, CLR_MUTED, "txtSearch", 0)
    Set c = AddCombo("cboCategory", "", 4366, 1304, 2608, 454, "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", 2, "0;2552")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCategory", "«· ’‰Ì›", 4366, 992, 2608, 284, 9, False, CLR_MUTED, "cboCategory", 0)
    Set c = AddCheck("chkLowOnly", "", 7144, 1389)
    c.AfterUpdate = EP
    Set c = AddLabel("lblLowOnly", "„‰Œ›÷… «·„Œ“Ê‰ ›ﬁÿ", 7513, 1304, 2239, 454, 10, False, CLR_TEXT, "chkLowOnly", 0)
    Set c = AddButton("btnRefresh", " ÕœÌÀ", 9866, 1287, 1134, 482, "secondary")
    c.OnClick = EP
    Set c = AddList("lstProducts", 227, 1984, 11680, 5613, 8, "0;1361;3686;1701;1134;1134;1247;1304", True)
    c.AfterUpdate = EP
    c.OnDblClick = EP
    Set c = AddLabel("lblInvTotals", " ", 227, 7711, 11680, 340, 10, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnStockCount", "«·Ã—œ", 227, 8221, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPurchase", "›« Ê—… „‘ —Ì« ", 1814, 8221, 2041, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLowReport", " ﬁ—Ì— «·‰Ê«ﬁ’", 3968, 8221, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStockReport", " ﬁ—Ì— «·„Œ“Ê‰", 6009, 8221, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLabels", "ÿ»«⁄… »«—ﬂÊœ", 8050, 8221, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblProductName", "«Œ — „‰ Ã« „‰ «·ﬁ«∆„…", 12134, 1304, 6634, 454, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblProductStock", " ", 12134, 1786, 6634, 340, 10, False, CLR_TEXT, "", 0)
    Set c = AddLabel("lblManualCap", "Õ—ﬂ… ÌœÊÌ… ⁄·Ï «·„‰ Ã «·„Õœœ", 12134, 2268, 6634, 340, 11, True, CLR_MUTED, "", 0)
    Set c = AddCombo("cboMoveType", "", 12134, 2977, 3175, 454, "SELECT TransactionTypeID, TypeName FROM TransactionTypes WHERE IsManual = True ORDER BY TransactionTypeID", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblMoveType", "‰Ê⁄ «·Õ—ﬂ…", 12134, 2665, 3175, 284, 9, False, CLR_MUTED, "cboMoveType", 0)
    Set c = AddText("txtMoveQty", "", 15479, 2977, 1531, 454)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMoveQty", "«·ﬂ„Ì…", 15479, 2665, 1531, 284, 9, False, CLR_MUTED, "txtMoveQty", 0)
    Set c = AddText("txtMoveCost", "", 17180, 2977, 1588, 454)
    SetCtlProp c, "Format", "#,##0.00##"
    Set c = AddLabel("lblMoveCost", " ﬂ·›… «·ÊÕœ…", 17180, 2665, 1588, 284, 9, False, CLR_MUTED, "txtMoveCost", 0)
    Set c = AddText("txtMoveNotes", "", 12134, 3799, 4876, 454)
    Set c = AddLabel("lblMoveNotes", "«·”»» / „·«ÕŸ« ", 12134, 3487, 4876, 284, 9, False, CLR_MUTED, "txtMoveNotes", 0)
    Set c = AddButton("btnPostMove", "Õ›Ÿ «·Õ—ﬂ…", 17180, 3782, 1588, 482, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblMovesCap", "¬Œ— Õ—ﬂ«  «·„‰ Ã", 12134, 4451, 6634, 340, 11, True, CLR_MUTED, "", 0)
    Set c = AddList("lstMoves", 12134, 4820, 6634, 3232, 5, "1474;1474;1021;1134;1361", True)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17067, 8221, 1701, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    InventoryLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_AfterUpdate()" & vbCrLf
    s = s & "    InventoryRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCategory_AfterUpdate()" & vbCrLf
    s = s & "    InventoryRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkLowOnly_AfterUpdate()" & vbCrLf
    s = s & "    InventoryRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstProducts_AfterUpdate()" & vbCrLf
    s = s & "    InventoryProductPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstProducts_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 0, Me!lstProducts.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboMoveType_AfterUpdate()" & vbCrLf
    s = s & "    InventoryMoveTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    InventoryRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStockCount_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmStockCount"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPurchase_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseInvoice"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLowReport_Click()" & vbCrLf
    s = s & "    OpenReportOrQuery ""rptLowStock"", ""LowStockQuery"", """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStockReport_Click()" & vbCrLf
    s = s & "    OpenReportOrQuery ""rptStockBalance"", ""StockBalanceQuery"", """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLabels_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBarcodeLabels"", 0, Me!lstProducts.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPostMove_Click()" & vbCrLf
    s = s & "    PostInventoryMove Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;lstProducts,227,1984,11680,5613,0,1000,0,1000;lblInvTotals,227,7711,11680,340,0,1000,1000,0;btnStockCount,227,8221,1474,624,0,0,1000,0;btnPurchase,1814,8221,2041,624,0,0,1000,0;btnLowReport,3968,8221,1928,624,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnStockReport,6009,8221,1928,624,0,0,1000,0;btnLabels,8050,8221,1928,624,0,0,1000,0;lblProductName,12134,1304,6634,454,1000,0,0,0;lblProductStock,12134,1786,6634,340,1000,0,0,0;lblManualCap,12134,2268,6634,340,1000,0,0,0;cboMoveType,12134,2977,3175,454,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblMoveType,12134,2665,3175,284,1000,0,0,0;txtMoveQty,15479,2977,1531,454,1000,0,0,0;lblMoveQty,15479,2665,1531,284,1000,0,0,0;txtMoveCost,17180,2977,1588,454,1000,0,0,0;lblMoveCost,17180,2665,1588,284,1000,0,0,0;txtMoveNotes,12134,3799,4876,454,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblMoveNotes,12134,3487,4876,284,1000,0,0,0;btnPostMove,17180,3782,1588,482,1000,0,0,0;lblMovesCap,12134,4451,6634,340,1000,0,0,0;lstMoves,12134,4820,6634,3232,1000,0,0,1000;btnClose,17067,8221,1701,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9355, -1757, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmInventory", s
    Exit Sub
EH:
    AbortForm "frmInventory", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmStockCountLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmStockCountLines", "√”ÿ— «·Ã—œ", "SELECT d.StockCountDetailID, d.StockCountID, d.ProductID, d.SystemQuantity, d.ActualQuantity, d.Difference, d.UnitCost, d.DifferenceValue, d.Notes, p.ProductCode, p.ProductName, p.Barcode FROM StockCountDetails AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID WHERE d.StockCountID = 0 ORDER BY p.ProductName", 18541, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("ProductCode", "ProductCode", 28, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("ProductName", "ProductName", 1530, 0, 4876, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("SystemQuantity", "SystemQuantity", 6434, 0, 1701, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("ActualQuantity", "ActualQuantity", 8163, 0, 1701, 425)
    c.FontBold = True
    SetCtlProp c, "Format", "#,##0.###"
    c.AfterUpdate = EP
    Set c = AddText("Difference", "Difference", 9892, 0, 1701, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddText("DifferenceValue", "DifferenceValue", 11621, 0, 1814, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Notes", "Notes", 13463, 0, 4196, 425)
    s = ""
    s = s & "Private Sub ActualQuantity_AfterUpdate()" & vbCrLf
    s = s & "    CountLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmStockCountLines", s
    Exit Sub
EH:
    AbortForm "frmStockCountLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmStockCount()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmStockCount", "«·Ã—œ", "", 18994, 9355, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8EF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·Ã—œ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«»œ√ Ã—œ«° √œŒ· «·ﬂ„Ì… «·›⁄·Ì… √Ê «„”Õ «·»«—ﬂÊœ° À„ —Õ¯· «·›—Êﬁ« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboCount", "", 227, 1304, 5103, 454, "SELECT StockCountID, CountNumber, CountDate, IIf(Status = 'OPEN', '„› ÊÕ', IIf(Status = 'POSTED', '„ı—Õ¯·', '„·€Ï')) AS StatusName FROM StockCounts ORDER BY StockCountID DESC", 4, "0;1701;1984;1304")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCount", "Ã·”… «·Ã—œ («·√ÕœÀ √Ê·«)", 227, 992, 5103, 284, 9, False, CLR_MUTED, "cboCount", 0)
    Set c = AddCombo("cboCategory", "", 5500, 1304, 2835, 454, "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", 2, "0;2835")
    Set c = AddLabel("lblCategory", " ’‰Ì› «·Ã—œ «·ÃœÌœ (›«—€ = «·ﬂ·)", 5500, 992, 2835, 284, 9, False, CLR_MUTED, "cboCategory", 0)
    Set c = AddButton("btnNewCount", "Ã—œ ÃœÌœ", 8505, 1287, 1701, 482, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblCountInfo", " ", 10376, 1304, 8391, 454, 10, True, CLR_TEXT, "", 0)
    Set c = AddText("txtCountBarcode", "", 227, 2183, 3969, 454)
    c.FontSize = 12
    c.OnKeyDown = EP
    Set c = AddLabel("lblCountBarcode", "«„”Õ «·»«—ﬂÊœ (Ì÷Ì› 1) √Ê 3*«·ﬂÊœ", 227, 1871, 3969, 284, 9, False, CLR_MUTED, "txtCountBarcode", 0)
    Set c = AddCheck("chkDiffOnly", "", 4366, 2268)
    c.AfterUpdate = EP
    Set c = AddLabel("lblDiffOnly", "«·›—Êﬁ«  ›ﬁÿ", 4735, 2183, 1672, 454, 10, False, CLR_TEXT, "chkDiffOnly", 0)
    Set c = AddButton("btnRefreshSystem", " ÕœÌÀ «·ﬂ„Ì«  «·„”Ã·…", 6577, 2166, 2608, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblCol1", "«·ﬂÊœ", 255, 2778, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·„‰ Ã", 1757, 2778, 4876, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·ﬂ„Ì… «·„”Ã·…", 6661, 2778, 1701, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·ﬂ„Ì… «·›⁄·Ì…", 8390, 2778, 1701, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "«·›—ﬁ", 10119, 2778, 1701, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "ﬁÌ„… «·›—ﬁ", 11848, 2778, 1814, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "„·«ÕŸ« ", 13690, 2778, 4196, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subCountLines", "frmStockCountLines", 227, 3118, 18541, 4394)
    Set c = AddLabel("lblCountSummary", " ", 227, 7626, 18541, 397, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnPostCount", " —ÕÌ· «·Ã—œ", 227, 8221, 1928, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnCancelCount", "≈·€«¡ «·Ã—œ", 2268, 8221, 1701, 624, "danger")
    c.OnClick = EP
    Set c = AddButton("btnCountReport", "ÿ»«⁄… «·Ã—œ", 4082, 8221, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17066, 8221, 1701, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    StockCountLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCount_AfterUpdate()" & vbCrLf
    s = s & "    StockCountPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtCountBarcode_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    CountBarcodeKeyDown Me, KeyCode" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkDiffOnly_AfterUpdate()" & vbCrLf
    s = s & "    CountFilterChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewCount_Click()" & vbCrLf
    s = s & "    NewStockCount Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRefreshSystem_Click()" & vbCrLf
    s = s & "    CountRefreshSystem Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPostCount_Click()" & vbCrLf
    s = s & "    PostCountScreen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCancelCount_Click()" & vbCrLf
    s = s & "    CancelCountScreen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCountReport_Click()" & vbCrLf
    s = s & "    PrintStockCount Me!cboCount.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;lblCountInfo,10376,1304,8391,454,0,1000,0,0;subCountLines,227,3118,18541,4394,0,1000,0,1000;lblCountSummary,227,7626,18541,397,0,1000,1000,0;btnPostCount,227,8221,1928,624,0,0,1000,0;btnCancelCount,2268,8221,1701,624,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnCountReport,4082,8221,1928,624,0,0,1000,0;btnClose,17066,8221,1701,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9355, -2919, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmStockCount", s
    Exit Sub
EH:
    AbortForm "frmStockCount", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmLogin()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmLogin", " ”ÃÌ· «·œŒÊ·", "", 9072, 5443, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", " ”ÃÌ· «·œŒÊ·", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "‰Ÿ«„ ≈œ«—… «·„Õ·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblStoreName", " ", 227, 992, 8618, 425, 13, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtUsername", "", 227, 1758, 8618, 510)
    c.FontSize = 13
    c.OnKeyDown = EP
    Set c = AddLabel("lblUsername", "«”„ «·„” Œœ„", 227, 1446, 8618, 284, 9, False, CLR_MUTED, "txtUsername", 0)
    Set c = AddText("txtPassword", "", 227, 2608, 8618, 510)
    c.FontSize = 13
    SetCtlProp c, "InputMask", "Password"
    c.OnKeyDown = EP
    Set c = AddLabel("lblPassword", "ﬂ·„… «·„—Ê—", 227, 2296, 8618, 284, 9, False, CLR_MUTED, "txtPassword", 0)
    Set c = AddLabel("lblMessage", " ", 227, 3204, 8618, 680, 10, True, CLR_DANGER, "", 0)
    Set c = AddButton("btnLogin", "œŒÊ·", 227, 4196, 3062, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnExit", "Œ—ÊÃ", 5783, 4196, 3062, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    LoginLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    LoginKeyDown Me, KeyCode, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    LoginKeyDown Me, KeyCode, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLogin_Click()" & vbCrLf
    s = s & "    DoLogin Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnExit_Click()" & vbCrLf
    s = s & "    LoginExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmLogin", s
    Exit Sub
EH:
    AbortForm "frmLogin", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmChangePassword()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmChangePassword", "ﬂ·„… «·„—Ê—", "", 9072, 5897, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ﬂ·„… «·„—Ê—", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " €ÌÌ— √Ê  ⁄ÌÌ‰ ﬂ·„… «·„—Ê—", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblFor", " ", 227, 992, 8618, 369, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtOld", "", 227, 1729, 8618, 482)
    c.FontSize = 12
    SetCtlProp c, "InputMask", "Password"
    Set c = AddLabel("lblOld", "ﬂ·„… «·„—Ê— «·Õ«·Ì…", 227, 1417, 8618, 284, 9, False, CLR_MUTED, "txtOld", 0)
    Set c = AddText("txtNew", "", 227, 2551, 8618, 482)
    c.FontSize = 12
    SetCtlProp c, "InputMask", "Password"
    Set c = AddLabel("lblNew", "ﬂ·„… «·„—Ê— «·ÃœÌœ…", 227, 2239, 8618, 284, 9, False, CLR_MUTED, "txtNew", 0)
    Set c = AddText("txtConfirm", "", 227, 3373, 8618, 482)
    c.FontSize = 12
    SetCtlProp c, "InputMask", "Password"
    Set c = AddLabel("lblConfirm", " √ﬂÌœ ﬂ·„… «·„—Ê— «·ÃœÌœ…", 227, 3061, 8618, 284, 9, False, CLR_MUTED, "txtConfirm", 0)
    Set c = AddCheck("chkMustChange", "", 227, 4196)
    Set c = AddLabel("lblMustChange", "Ì€Ì¯—Â« «·„” Œœ„ ⁄‰œ √Ê· œŒÊ· (ﬂ·„… „ƒﬁ …)", 596, 4111, 8249, 454, 10, False, CLR_TEXT, "chkMustChange", 0)
    Set c = AddLabel("lblRules", "6 √Õ—› ⁄·Ï «·√ﬁ·° Ê·«  ”«ÊÌ «”„ «·„” Œœ„", 227, 4593, 8618, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ", 227, 5103, 3062, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnCancel", "≈·€«¡", 5783, 5103, 3062, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ChangePasswordLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveChangedPassword Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    CancelChangePassword Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmChangePassword", s
    Exit Sub
EH:
    AbortForm "frmChangePassword", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmRolePermLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmRolePermLines", "’·«ÕÌ«  «·œÊ—", "SELECT * FROM tmpRolePermissions ORDER BY SortOrder", 8392, 397, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddCheck("Granted", "Granted", 339, 68)
    Set c = AddText("PermissionName", "PermissionName", 963, 0, 5386, 397)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("ModuleName", "ModuleName", 6377, 0, 1928, 397)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    s = ""
    FinishForm "frmRolePermLines", s
    Exit Sub
EH:
    AbortForm "frmRolePermLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmRoles()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmRoles", "«·√œÊ«— Ê«·’·«ÕÌ« ", "", 9072, 9072, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·√œÊ«— Ê«·’·«ÕÌ« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "Õœœ „« Ì” ÿÌ⁄ ﬂ· œÊ— ›⁄·Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboRole", "", 227, 1304, 3969, 454, "SELECT RoleID, RoleName FROM Roles ORDER BY RoleID", 2, "0;3402")
    c.AfterUpdate = EP
    Set c = AddLabel("lblRole", "«·œÊ—", 227, 992, 3969, 284, 9, False, CLR_MUTED, "cboRole", 0)
    Set c = AddLabel("lblRoleInfo", " ", 4366, 1332, 4479, 397, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCol1", "„„‰ÊÕ…", 255, 1956, 907, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·’·«ÕÌ…", 1190, 1956, 5386, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·ﬁ”„", 6604, 1956, 1928, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subPermissions", "frmRolePermLines", 227, 2296, 8392, 5443)
    Set c = AddLabel("lblLockedNote", " ", 227, 7825, 8618, 340, 9, True, CLR_WARNING, "", 0)
    Set c = AddButton("btnSaveRole", "Õ›Ÿ «·’·«ÕÌ« ", 227, 8278, 2155, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnAll", " ÕœÌœ «·ﬂ·", 2495, 8278, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNone", "≈·€«¡ «·ﬂ·", 4309, 8278, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 7371, 8278, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    RolesLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboRole_AfterUpdate()" & vbCrLf
    s = s & "    RolePicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveRole_Click()" & vbCrLf
    s = s & "    SaveRolePermissions Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAll_Click()" & vbCrLf
    s = s & "    RoleSelectAll Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNone_Click()" & vbCrLf
    s = s & "    RoleSelectAll Me, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmRoles", s
    Exit Sub
EH:
    AbortForm "frmRoles", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBackup()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBackup", "«·‰”Œ «·«Õ Ì«ÿÌ", "", 11340, 8392, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 11340, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8B7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·‰”Œ «·«Õ Ì«ÿÌ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "‰”Œ… „‰ „·› «·»Ì«‰«  »«· «—ÌŒ Ê«·Êﬁ ° Ê«·«” ⁄«œ… ⁄‰œ «·Õ«Ã…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblDataFile", " ", 227, 1021, 10886, 340, 9, False, CLR_TEXT, "", 0)
    Set c = AddLabel("lblFolder", " ", 227, 1418, 7654, 340, 9, False, CLR_TEXT, "", 0)
    Set c = AddButton("btnChooseFolder", " €ÌÌ— «·„Ã·œ", 7995, 1372, 1474, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnOpenFolder", "› Õ «·„Ã·œ", 9639, 1372, 1474, 425, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblLast", " ", 227, 1871, 10886, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnBackupNow", "‰”Œ… «Õ Ì«ÿÌ… «·¬‰", 227, 2325, 3062, 624, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblListCap", "«·‰”Œ «·„ÊÃÊœ… («·√ÕœÀ √Ê·«)", 227, 3147, 6804, 340, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstBackups", 227, 3515, 10886, 3686, 3, "5670;2835;1984", True)
    c.RowSourceType = "Value List"
    Set c = AddButton("btnRestore", "«” ⁄«œ… «·‰”Œ… «·„Õœœ…", 227, 7484, 3062, 567, "danger")
    c.OnClick = EP
    Set c = AddButton("btnDevMode", "Ê÷⁄ «·„ÿÊ¯—", 3402, 7484, 2041, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 9639, 7484, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    BackupLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnChooseFolder_Click()" & vbCrLf
    s = s & "    BackupChooseFolder Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenFolder_Click()" & vbCrLf
    s = s & "    BackupOpenFolder" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBackupNow_Click()" & vbCrLf
    s = s & "    BackupRun Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRestore_Click()" & vbCrLf
    s = s & "    BackupRestoreSelected Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDevMode_Click()" & vbCrLf
    s = s & "    DeveloperModeFromApp" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBackup", s
    Exit Sub
EH:
    AbortForm "frmBackup", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmLabelLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmLabelLines", "√”ÿ— «·„·’ﬁ« ", "", 10773, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("ProductName", "ProductName", 28, 0, 4876, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("LabelCode", "LabelCode", 4932, 0, 2495, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("Price", "Price", 7455, 0, 1361, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Copies", "Copies", 8844, 0, 1361, 425)
    c.FontBold = True
    SetCtlProp c, "Format", "0"
    c.AfterUpdate = EP
    Set c = AddButton("btnRemove", "Sym(code='ChrW(&HE74D)')", 10233, 17, 454, 391, "danger")
    SetCtlProp c, "FontName", ICON_FONT
    c.OnClick = EP
    Set c = AddText("LineNo", "LineNo", 10716, 0, 28, 425)
    SetCtlProp c, "Visible", False
    m_frm.OnOpen = EP
    s = ""
    s = s & "Private Sub Form_Open(Cancel As Integer)" & vbCrLf
    s = s & "    LabelLinesOpen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Copies_AfterUpdate()" & vbCrLf
    s = s & "    LabelCopiesChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    RemoveLabelLine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmLabelLines", s
    Exit Sub
EH:
    AbortForm "frmLabelLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBarcodeLabels()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBarcodeLabels", "ÿ»«⁄… „·’ﬁ«  «·»«—ﬂÊœ", "", 11340, 8278, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 11340, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8EC), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ÿ»«⁄… „·’ﬁ«  «·»«—ﬂÊœ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«„”Õ «·»«—ﬂÊœ √Ê «Œ — «·’‰› »«·«”„° ÊÕœœ ⁄œœ «·„·’ﬁ«   |  Enter ··≈÷«›…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtBarcode", "", 227, 1304, 3402, 510)
    c.FontSize = 14
    c.OnKeyDown = EP
    Set c = AddLabel("lblBarcode", "«·»«—ﬂÊœ √Ê ﬂÊœ «·„‰ Ã (Enter)", 227, 992, 3402, 284, 9, False, CLR_MUTED, "txtBarcode", 0)
    Set c = AddText("txtCopies", "", 3799, 1304, 1247, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "0"
    SetCtlProp c, "DefaultValue", "1"
    Set c = AddLabel("lblCopies", "⁄œœ «·„·’ﬁ« ", 3799, 992, 1247, 284, 9, False, CLR_MUTED, "txtCopies", 0)
    Set c = AddCombo("cboProduct", "", 5216, 1304, 5897, 510, "SELECT ProductID, ProductName, SellingPrice FROM Products WHERE IsActive = True ORDER BY ProductName", 3, "0;4536;1134")
    c.FontSize = 12
    c.AfterUpdate = EP
    Set c = AddLabel("lblProduct", "√Ê «Œ — «·„‰ Ã »«·«”„", 5216, 992, 5897, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddLabel("lblCol1", "«·’‰›", 255, 2013, 4876, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·»«—ﬂÊœ «·„ÿ»Ê⁄", 5159, 2013, 2495, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·”⁄—", 7682, 2013, 1361, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "⁄œœ «·„·’ﬁ« ", 9071, 2013, 1361, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmLabelLines", 227, 2381, 10886, 4309)
    Set c = AddLabel("lblStatus", " ", 227, 6804, 10886, 340, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnPreview", "„⁄«Ì‰…", 227, 7428, 1588, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 1928, 7428, 1588, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClear", "„”Õ «·ﬁ«∆„…", 3629, 7428, 1701, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSettings", "≈⁄œ«œ«  «·„·’ﬁ", 5443, 7428, 1928, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 9639, 7428, 1474, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    LabelsLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    LabelBarcodeKeyDown Me, KeyCode" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboProduct_AfterUpdate()" & vbCrLf
    s = s & "    LabelProductPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPreview_Click()" & vbCrLf
    s = s & "    PrintLabels Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintLabels Me, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    ClearLabels Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLabelSettings""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,11340,850,0,1000,0,0;cboProduct,5216,1304,5897,510,0,1000,0,0;subLines,227,2381,10886,4309,0,1000,0,1000;lblStatus,227,6804,10886,340,0,1000,1000,0;btnPreview,227,7428,1588,624,0,0,1000,0;btnPrint,1928,7428,1588,624,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnClear,3629,7428,1701,624,0,0,1000,0;btnSettings,5443,7428,1928,624,0,0,1000,0;btnClose,9639,7428,1474,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 11340, 8278, -2834, " & IIf(MIRROR_LAYOUT, "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBarcodeLabels", s
    Exit Sub
EH:
    AbortForm "frmBarcodeLabels", Err.Number, Err.Description
End Sub
