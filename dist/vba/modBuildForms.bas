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
' The interface language of this file (modLang.UiLanguage) decides the
' direction: an English file is left-to-right and every screen is mirrored;
' captions, lists, record sources and the screen code are translated by Tr.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MIRROR_LAYOUT As Boolean = False
Private Const EP As String = "[Event Procedure]"
Private Const FORM_NAMES As String = "frmMain,frmProducts,frmCustomers,frmSuppliers,frmExpenses,frmCurrencies,frmCurrencyRates,frmSalesReps,frmRepTargets,frmRecurring,frmUsers,frmCostCenters,frmEmployeePay,frmCategories,frmUnits,frmExpenseTypes,frmCashBoxes,frmBanks,frmAccounts,frmSettings,frmLabelSettings,frmSearch,frmReportCenter,frmPOSLines,frmPOS,frmReturnLines,frmSalesReturn,frmCustomerPayment,frmSalesInvoice,frmPurchaseLines," & _
    "frmPurchaseInvoice,frmPurchaseReturnLines,frmPurchaseReturn,frmSupplierPayment,frmPurchaseView,frmInventory,frmStockCountLines,frmStockCount,frmLogin,frmChangePassword,frmRolePermLines,frmRoles,frmUserScreenLines,frmUserScreens,frmActivation,frmBackup,frmLabelLines,frmBarcodeLabels,frmTouchLines,frmTouchPOS,frmTouchPay,frmCafePOS,frmCafeItem,frmTreasury,frmCashVoucher,frmCashClosing,frmJournal," & _
    "frmJournalEntry,frmManualLines,frmManualEntry,frmLedger,frmFinancials,frmPeriodClosing,frmVatReturn,frmAging,frmAllocation,frmBankTx,frmBankRecon,frmCheques,frmAssets,frmDepreciation,frmPayrollLines,frmPayroll,frmBudgetLines,frmBudget,frmAccounting,frmAuditLog,frmCommissionLines,frmCommissions"

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
    SetFormProp "Orientation", IIf(UiEnglish(), 0, 1)     ' right-to-left, or left-to-right in English
    m_frm.Caption = Tr(Caption)
    m_frm.RecordSource = Tr(RecordSource)
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
    m_frm.Tag = Tr(TagText)                      ' the list query of a data screen has captions
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
    mdl.AddFromString Tr(Code)
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
    If MirrorLayout() Then x = m_width - L - W Else x = L
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
    c.Caption = Tr(Caption)
    c.FontName = FONT_NAME
    c.FontSize = FontSize
    c.FontBold = Bold
    c.ForeColor = Color
    c.BackStyle = 0
    c.TextAlign = UiAlign(TextAlign)
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
    c.RowSource = Tr(Rows)
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

Private Function AddImage(ByVal CtlName As String, ByVal L As Long, ByVal T As Long, ByVal W As Long, _
                          ByVal H As Long) As Access.Control
    ' Picture loaded at run time (product / category images on the touch screens), scaled to fit.
    Dim c As Access.Control
    Set c = NewCtl(acImage, CtlName, L, T, W, H)
    c.SizeMode = 3                                   ' acOLESizeZoom
    c.BorderStyle = 0
    c.BackStyle = 0
    c.PictureType = 1                                ' linked
    Set AddImage = c
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
    c.Caption = Tr(Caption)
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
    If PropName = "TextAlign" Then Value = UiAlign(Value)
    c.Properties(PropName).Value = Tr(Value)
    If Err.Number <> 0 Then
        m_warnings = m_warnings & "  " & m_tmpName & "." & c.Name & "." & PropName & _
                     ": " & Err.Description & vbCrLf
    End If
End Sub

Private Function MirrorLayout() As Boolean
    ' The design is right-to-left; an English file mirrors every screen.
    MirrorLayout = (MIRROR_LAYOUT Xor UiEnglish())
End Function

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
    ' frmCafeItem shows the drink chosen on frmCafePOS: give it one
    If FormName = "frmCafeItem" Then TempVars.Add "CafeItemID", Nz(DMin("ProductID", "Products", "IsActive = True"), 0)
    DoCmd.OpenForm FormName, acNormal, , , , acHidden
    If Not CurrentProject.AllForms(FormName).IsLoaded Then
        ' a dialog that closes itself when it has nothing to show (no product yet)
        Record True, "«·‘«‘… " & FormName & "  › Õ Ê ı€·ﬁ ‰›”Â« ·⁄œ„ ÊÃÊœ »Ì«‰«   ⁄—÷Â«"
        Exit Sub
    End If
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
    Dim db As DAO.Database, qdf As DAO.QueryDef
    Set db = CurrentDb
    For Each qdf In db.QueryDefs
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
    BuildForm_frmCurrencies
    BuildForm_frmCurrencyRates
    BuildForm_frmSalesReps
    BuildForm_frmRepTargets
    BuildForm_frmRecurring
    BuildForm_frmUsers
    BuildForm_frmCostCenters
    BuildForm_frmEmployeePay
    BuildForm_frmCategories
    BuildForm_frmUnits
    BuildForm_frmExpenseTypes
    BuildForm_frmCashBoxes
    BuildForm_frmBanks
    BuildForm_frmAccounts
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
    BuildForm_frmUserScreenLines
    BuildForm_frmUserScreens
    BuildForm_frmActivation
    BuildForm_frmBackup
    BuildForm_frmLabelLines
    BuildForm_frmBarcodeLabels
    BuildForm_frmTouchLines
    BuildForm_frmTouchPOS
    BuildForm_frmTouchPay
    BuildForm_frmCafePOS
    BuildForm_frmCafeItem
    BuildForm_frmTreasury
    BuildForm_frmCashVoucher
    BuildForm_frmCashClosing
    BuildForm_frmJournal
    BuildForm_frmJournalEntry
    BuildForm_frmManualLines
    BuildForm_frmManualEntry
    BuildForm_frmLedger
    BuildForm_frmFinancials
    BuildForm_frmPeriodClosing
    BuildForm_frmVatReturn
    BuildForm_frmAging
    BuildForm_frmAllocation
    BuildForm_frmBankTx
    BuildForm_frmBankRecon
    BuildForm_frmCheques
    BuildForm_frmAssets
    BuildForm_frmDepreciation
    BuildForm_frmPayrollLines
    BuildForm_frmPayroll
    BuildForm_frmBudgetLines
    BuildForm_frmBudget
    BuildForm_frmAccounting
    BuildForm_frmAuditLog
    BuildForm_frmCommissionLines
    BuildForm_frmCommissions
End Sub

Private Sub BuildForm_frmMain()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmMain", "‰Ÿ«„ ≈œ«—… «·„Õ·", "", 18994, 9634, False, False, False, _
              ""
    SetFormProp "TimerInterval", 300000
    Set c = AddRect("boxSidebar", 0, 0, 3515, 9634, CLR_PRIMARY)
    Set c = AddIcon("icoApp", ChrW(&HE80F), 227, 255, 567, 567, 22, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblAppTitle", "‰Ÿ«„ ≈œ«—… «·„Õ·", 850, 227, 2551, 425, 15, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblStoreName", " ", 850, 652, 2551, 312, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddIcon("icoSales", ChrW(&HE7BF), 255, 1361, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavSales", "«·„»Ì⁄« ", 850, 1349, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSales", "«·„»Ì⁄« ", 142, 1304, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPOS"
    c.OnClick = EP
    Set c = AddIcon("icoPurchases", ChrW(&HE896), 255, 1871, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavPurchases", "«·„‘ —Ì« ", 850, 1859, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavPurchases", "«·„‘ —Ì« ", 142, 1814, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPurchaseInvoice"
    c.OnClick = EP
    Set c = AddIcon("icoInventory", ChrW(&HE7B8), 255, 2381, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavInventory", "«·„Œ“Ê‰", 850, 2369, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavInventory", "«·„Œ“Ê‰", 142, 2324, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmInventory"
    c.OnClick = EP
    Set c = AddIcon("icoProducts", ChrW(&HE8EC), 255, 2891, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavProducts", "«·„‰ Ã« ", 850, 2879, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavProducts", "«·„‰ Ã« ", 142, 2834, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmProducts"
    c.OnClick = EP
    Set c = AddIcon("icoCustomers", ChrW(&HE716), 255, 3401, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavCustomers", "«·⁄„·«¡", 850, 3389, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavCustomers", "«·⁄„·«¡", 142, 3344, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCustomers"
    c.OnClick = EP
    Set c = AddIcon("icoSuppliers", ChrW(&HE77B), 255, 3911, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavSuppliers", "«·„Ê—œÊ‰", 850, 3899, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSuppliers", "«·„Ê—œÊ‰", 142, 3854, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSuppliers"
    c.OnClick = EP
    Set c = AddIcon("icoExpenses", ChrW(&HE8C7), 255, 4421, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavExpenses", "«·„’—Ê›« ", 850, 4409, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavExpenses", "«·„’—Ê›« ", 142, 4364, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmExpenses"
    c.OnClick = EP
    Set c = AddIcon("icoTreasury", ChrW(&HE825), 255, 4931, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavTreasury", "«·Œ“Ì‰…", 850, 4919, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavTreasury", "«·Œ“Ì‰…", 142, 4874, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmTreasury"
    c.OnClick = EP
    Set c = AddIcon("icoAccounting", ChrW(&HE8F1), 255, 5441, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavAccounting", "«·„Õ«”»… Ê«·„«·Ì…", 850, 5429, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavAccounting", "«·„Õ«”»… Ê«·„«·Ì…", 142, 5384, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAccounting"
    c.OnClick = EP
    Set c = AddIcon("icoStockCount", ChrW(&HE8EF), 255, 5951, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavStockCount", "«·Ã—œ", 850, 5939, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavStockCount", "«·Ã—œ", 142, 5894, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmStockCount"
    c.OnClick = EP
    Set c = AddIcon("icoReports", ChrW(&HE8A5), 255, 6461, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavReports", "«· ﬁ«—Ì—", 850, 6449, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavReports", "«· ﬁ«—Ì—", 142, 6404, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmReportCenter"
    c.OnClick = EP
    Set c = AddIcon("icoSearch", ChrW(&HE721), 255, 6971, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavSearch", "«·»ÕÀ", 850, 6959, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSearch", "«·»ÕÀ", 142, 6914, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSearch"
    c.OnClick = EP
    Set c = AddIcon("icoSettings", ChrW(&HE713), 255, 7481, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavSettings", "«·≈⁄œ«œ« ", 850, 7469, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSettings", "«·≈⁄œ«œ« ", 142, 7424, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSettings"
    c.OnClick = EP
    Set c = AddIcon("icoUsers", ChrW(&HE8D7), 255, 7991, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavUsers", "«·„” Œœ„Ê‰", 850, 7979, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavUsers", "«·„” Œœ„Ê‰", 142, 7934, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmUsers"
    c.OnClick = EP
    Set c = AddIcon("icoBackup", ChrW(&HE8B7), 255, 8501, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavBackup", "‰”Œ… «Õ Ì«ÿÌ…", 850, 8489, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavBackup", "‰”Œ… «Õ Ì«ÿÌ…", 142, 8444, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmBackup"
    c.OnClick = EP
    Set c = AddIcon("icoLogout", ChrW(&HE7E8), 255, 9011, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblNavLogout", " ”ÃÌ· «·Œ—ÊÃ", 850, 8999, 2438, 380, 12, True, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavLogout", " ”ÃÌ· «·Œ—ÊÃ", 142, 8954, 3231, 471, "nav")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddLabel("lblWelcome", "·ÊÕ… «· Õﬂ„", 11736, 284, 6804, 539, 20, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblToday", " ", 11736, 879, 6804, 340, 11, False, CLR_MUTED, "", 3)
    Set c = AddButton("btnRefresh", " ÕœÌÀ", 3969, 340, 1361, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnChangePassword", "ﬂ·„… «·„—Ê—", 5443, 340, 1531, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblUpdated", " ", 7087, 425, 3402, 312, 9, False, CLR_MUTED, "", 1)
    Set c = AddLabel("lblUser", " ", 3969, 879, 6520, 340, 11, False, CLR_MUTED, "", 1)
    Set c = AddRect("boxTile1", 15066, 1418, 3472, 1361, CLR_SURFACE)
    Set c = AddRect("boxKpiIcon1", 17461, 1645, 907, 907, RGB(67, 160, 71))
    Set c = AddIcon("icoKpi1", ChrW(&HE7BF), 17461, 1787, 907, 624, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileTitle1", "„»Ì⁄«  «·ÌÊ„", 15236, 1503, 2111, 312, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue1", "-", 15236, 1815, 2111, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub1", " ", 15236, 2410, 2111, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile2", 11367, 1418, 3472, 1361, CLR_SURFACE)
    Set c = AddRect("boxKpiIcon2", 13762, 1645, 907, 907, RGB(30, 136, 229))
    Set c = AddIcon("icoKpi2", ChrW(&HE8A5), 13762, 1787, 907, 624, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileTitle2", "„»Ì⁄«  «·‘Â—", 11537, 1503, 2111, 312, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue2", "-", 11537, 1815, 2111, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub2", " ", 11537, 2410, 2111, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile3", 7668, 1418, 3472, 1361, CLR_SURFACE)
    Set c = AddRect("boxKpiIcon3", 10063, 1645, 907, 907, RGB(229, 57, 53))
    Set c = AddIcon("icoKpi3", ChrW(&HE8C7), 10063, 1787, 907, 624, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileTitle3", "’«›Ì —»Õ «·‘Â— ( ﬁ—Ì»Ì)", 7838, 1503, 2111, 312, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue3", "-", 7838, 1815, 2111, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub3", " ", 7838, 2410, 2111, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile4", 3969, 1418, 3472, 1361, CLR_SURFACE)
    Set c = AddRect("boxKpiIcon4", 6364, 1645, 907, 907, RGB(251, 140, 0))
    Set c = AddIcon("icoKpi4", ChrW(&HE7B8), 6364, 1787, 907, 624, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileTitle4", "„‰ Ã«  „‰Œ›÷… «·„Œ“Ê‰", 4139, 1503, 2111, 312, 10, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue4", "-", 4139, 1815, 2111, 567, 20, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub4", " ", 4139, 2410, 2111, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxNavSales", 15066, 2948, 3472, 992, RGB(67, 160, 71))
    Set c = AddIcon("icoTileSales", ChrW(&HE7BF), 15066, 3016, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileSales", "«·„»Ì⁄« ", 15066, 3515, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileSales", "«·„»Ì⁄« ", 15066, 2948, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPOS"
    c.OnClick = EP
    Set c = AddRect("boxNavPurchases", 11367, 2948, 3472, 992, RGB(30, 136, 229))
    Set c = AddIcon("icoTilePurchases", ChrW(&HE896), 11367, 3016, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTilePurchases", "«·„‘ —Ì« ", 11367, 3515, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTilePurchases", "«·„‘ —Ì« ", 11367, 2948, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPurchaseInvoice"
    c.OnClick = EP
    Set c = AddRect("boxNavInventory", 7668, 2948, 3472, 992, RGB(251, 140, 0))
    Set c = AddIcon("icoTileInventory", ChrW(&HE7B8), 7668, 3016, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileInventory", "«·„Œ“Ê‰", 7668, 3515, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileInventory", "«·„Œ“Ê‰", 7668, 2948, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmInventory"
    c.OnClick = EP
    Set c = AddRect("boxNavProducts", 3969, 2948, 3472, 992, RGB(142, 36, 170))
    Set c = AddIcon("icoTileProducts", ChrW(&HE8EC), 3969, 3016, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileProducts", "«·„‰ Ã« ", 3969, 3515, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileProducts", "«·„‰ Ã« ", 3969, 2948, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmProducts"
    c.OnClick = EP
    Set c = AddRect("boxNavCustomers", 15066, 4082, 3472, 992, RGB(229, 57, 53))
    Set c = AddIcon("icoTileCustomers", ChrW(&HE716), 15066, 4150, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileCustomers", "«·⁄„·«¡", 15066, 4649, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileCustomers", "«·⁄„·«¡", 15066, 4082, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCustomers"
    c.OnClick = EP
    Set c = AddRect("boxNavSuppliers", 11367, 4082, 3472, 992, RGB(57, 73, 171))
    Set c = AddIcon("icoTileSuppliers", ChrW(&HE77B), 11367, 4150, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileSuppliers", "«·„Ê—œÊ‰", 11367, 4649, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileSuppliers", "«·„Ê—œÊ‰", 11367, 4082, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSuppliers"
    c.OnClick = EP
    Set c = AddRect("boxNavExpenses", 7668, 4082, 3472, 992, RGB(0, 137, 123))
    Set c = AddIcon("icoTileExpenses", ChrW(&HE8C7), 7668, 4150, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileExpenses", "«·„’—Ê›« ", 7668, 4649, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileExpenses", "«·„’—Ê›« ", 7668, 4082, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmExpenses"
    c.OnClick = EP
    Set c = AddRect("boxNavReports", 3969, 4082, 3472, 992, RGB(216, 27, 96))
    Set c = AddIcon("icoTileReports", ChrW(&HE8A5), 3969, 4150, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileReports", "«· ﬁ«—Ì—", 3969, 4649, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileReports", "«· ﬁ«—Ì—", 3969, 4082, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmReportCenter"
    c.OnClick = EP
    Set c = AddRect("boxNavSettings", 15066, 5216, 3472, 992, RGB(232, 236, 243))
    Set c = AddIcon("icoTileSettings", ChrW(&HE713), 15066, 5284, 3472, 482, 22, False, CLR_PRIMARY, "", 2)
    Set c = AddLabel("lblTileSettings", "«·≈⁄œ«œ« ", 15066, 5783, 3472, 352, 12, True, CLR_TEXT, "", 2)
    Set c = AddButton("btnTileSettings", "«·≈⁄œ«œ« ", 15066, 5216, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSettings"
    c.OnClick = EP
    Set c = AddRect("boxNavUsers", 11367, 5216, 3472, 992, RGB(232, 236, 243))
    Set c = AddIcon("icoTileUsers", ChrW(&HE8D7), 11367, 5284, 3472, 482, 22, False, CLR_PRIMARY, "", 2)
    Set c = AddLabel("lblTileUsers", "«·„” Œœ„Ê‰", 11367, 5783, 3472, 352, 12, True, CLR_TEXT, "", 2)
    Set c = AddButton("btnTileUsers", "«·„” Œœ„Ê‰", 11367, 5216, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmUsers"
    c.OnClick = EP
    Set c = AddRect("boxNavTreasury", 7668, 5216, 3472, 992, RGB(0, 121, 107))
    Set c = AddIcon("icoTileTreasury", ChrW(&HE825), 7668, 5284, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileTreasury", "«·Œ“Ì‰…", 7668, 5783, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileTreasury", "«·Œ“Ì‰…", 7668, 5216, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmTreasury"
    c.OnClick = EP
    Set c = AddRect("boxNavAccounting", 3969, 5216, 3472, 992, RGB(31, 58, 95))
    Set c = AddIcon("icoTileAccounting", ChrW(&HE8F1), 3969, 5284, 3472, 482, 22, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileAccounting", "«·„Õ«”»… Ê«·„«·Ì…", 3969, 5783, 3472, 352, 12, True, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileAccounting", "«·„Õ«”»… Ê«·„«·Ì…", 3969, 5216, 3472, 992, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAccounting"
    c.OnClick = EP
    Set c = AddRect("boxTile5", 15066, 6378, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle5", "œÌÊ‰ «·⁄„·«¡", 15236, 6435, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue5", "-", 15236, 6730, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub5", " ", 15236, 7143, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile6", 11367, 6378, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle6", "„” Õﬁ«  «·„Ê—œÌ‰", 11537, 6435, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue6", "-", 11537, 6730, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub6", " ", 11537, 7143, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile7", 7668, 6378, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle7", "ﬁÌ„… «·„Œ“Ê‰ »«· ﬂ·›…", 7838, 6435, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue7", "-", 7838, 6730, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub7", " ", 7838, 7143, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile8", 3969, 6378, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle8", "„’—Ê›«  «·‘Â—", 4139, 6435, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue8", "-", 4139, 6730, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub8", " ", 4139, 7143, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile9", 15066, 7569, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle9", "Â«„‘ «·—»Õ «·≈Ã„«·Ì («·‘Â—)", 15236, 7626, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue9", "-", 15236, 7921, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub9", " ", 15236, 8334, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile10", 11367, 7569, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle10", "œÊ—«‰ «·„Œ“Ê‰ (12 ‘Â—«)", 11537, 7626, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue10", "-", 11537, 7921, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub10", " ", 11537, 8334, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile11", 7668, 7569, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle11", "„ Ê”ÿ › —… «· Õ’Ì·", 7838, 7626, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue11", "-", 7838, 7921, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub11", " ", 7838, 8334, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddRect("boxTile12", 3969, 7569, 3472, 1049, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle12", "‰”»… «·”ÌÊ·… («· œ«Ê·)", 4139, 7626, 3132, 284, 9, False, CLR_MUTED, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileValue12", "-", 4139, 7921, 3132, 408, 15, True, CLR_PRIMARY, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblTileSub12", " ", 4139, 8334, 3132, 255, 8, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblIntegrity", " ", 3969, 9180, 14571, 312, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnNavPurchases_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseInvoice"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavInventory_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmInventory"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavProducts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavCustomers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSuppliers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSuppliers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavExpenses_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenses"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavTreasury_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmTreasury"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavAccounting_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounting"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavStockCount_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmStockCount"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavReports_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmReportCenter"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSearch_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSearch"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSettings"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavUsers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUsers"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavBackup_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBackup"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNavLogout_Click()" & vbCrLf
    s = s & "    LogoutUser" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    DashboardRefresh Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnChangePassword_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmChangePassword"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileSales_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPOS"", 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTilePurchases_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPurchaseInvoice"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileInventory_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmInventory"", 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileProducts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmProducts"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileCustomers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCustomers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileSuppliers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSuppliers"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileExpenses_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenses"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileReports_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmReportCenter"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileSettings_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSettings"", 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileUsers_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUsers"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileTreasury_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmTreasury"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileAccounting_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounting"", 0" & vbCrLf
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
    s = s & "Private Sub lblTileTitle9_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""MARGIN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue9_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""MARGIN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle10_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""TURNOVER""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue10_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""TURNOVER""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle11_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""COLLECTION""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue11_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""COLLECTION""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileTitle12_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""LIQUIDITY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lblTileValue12_Click()" & vbCrLf
    s = s & "    DashboardTileClick ""LIQUIDITY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxSidebar,0,0,3515,9634,0,0,0,1000;lblWelcome,11736,284,6804,539,1000,0,0,0;lblToday,11736,879,6804,340,1000,0,0,0;boxTile1,15066,1418,3472,1361,750,250,0,0;boxKpiIcon1,17461,1645,907,907,750,250,0,0;icoKpi1,17461,1787,907,624,750,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle1,15236,1503,2111,312,750,250,0,0;lblTileValue1,15236,1815,2111,567,750,250,0,0;lblTileSub1,15236,2410,2111,284,750,250,0,0;boxTile2,11367,1418,3472,1361,500,250,0,0;boxKpiIcon2,13762,1645,907,907,500,250,0,0;icoKpi2,13762,1787,907,624,500,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle2,11537,1503,2111,312,500,250,0,0;lblTileValue2,11537,1815,2111,567,500,250,0,0;lblTileSub2,11537,2410,2111,284,500,250,0,0;boxTile3,7668,1418,3472,1361,250,250,0,0;boxKpiIcon3,10063,1645,907,907,250,250,0,0;icoKpi3,10063,1787,907,624,250,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle3,7838,1503,2111,312,250,250,0,0;lblTileValue3,7838,1815,2111,567,250,250,0,0;lblTileSub3,7838,2410,2111,284,250,250,0,0;boxTile4,3969,1418,3472,1361,0,250,0,0;boxKpiIcon4,6364,1645,907,907,0,250,0,0;icoKpi4,6364,1787,907,624,0,250,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle4,4139,1503,2111,312,0,250,0,0;lblTileValue4,4139,1815,2111,567,0,250,0,0;lblTileSub4,4139,2410,2111,284,0,250,0,0;boxNavSales,15066,2948,3472,992,750,250,0,333;icoTileSales,15066,3016,3472,482,750,250,0,333;lblTileSales,15066,3515,3472,352,750,250,0,333""" & vbCrLf
    s = s & "    spec = spec & "";btnTileSales,15066,2948,3472,992,750,250,0,333;boxNavPurchases,11367,2948,3472,992,500,250,0,333;icoTilePurchases,11367,3016,3472,482,500,250,0,333;lblTilePurchases,11367,3515,3472,352,500,250,0,333;btnTilePurchases,11367,2948,3472,992,500,250,0,333;boxNavInventory,7668,2948,3472,992,250,250,0,333""" & vbCrLf
    s = s & "    spec = spec & "";icoTileInventory,7668,3016,3472,482,250,250,0,333;lblTileInventory,7668,3515,3472,352,250,250,0,333;btnTileInventory,7668,2948,3472,992,250,250,0,333;boxNavProducts,3969,2948,3472,992,0,250,0,333;icoTileProducts,3969,3016,3472,482,0,250,0,333;lblTileProducts,3969,3515,3472,352,0,250,0,333""" & vbCrLf
    s = s & "    spec = spec & "";btnTileProducts,3969,2948,3472,992,0,250,0,333;boxNavCustomers,15066,4082,3472,992,750,250,333,333;icoTileCustomers,15066,4150,3472,482,750,250,333,333;lblTileCustomers,15066,4649,3472,352,750,250,333,333;btnTileCustomers,15066,4082,3472,992,750,250,333,333;boxNavSuppliers,11367,4082,3472,992,500,250,333,333""" & vbCrLf
    s = s & "    spec = spec & "";icoTileSuppliers,11367,4150,3472,482,500,250,333,333;lblTileSuppliers,11367,4649,3472,352,500,250,333,333;btnTileSuppliers,11367,4082,3472,992,500,250,333,333;boxNavExpenses,7668,4082,3472,992,250,250,333,333;icoTileExpenses,7668,4150,3472,482,250,250,333,333;lblTileExpenses,7668,4649,3472,352,250,250,333,333""" & vbCrLf
    s = s & "    spec = spec & "";btnTileExpenses,7668,4082,3472,992,250,250,333,333;boxNavReports,3969,4082,3472,992,0,250,333,333;icoTileReports,3969,4150,3472,482,0,250,333,333;lblTileReports,3969,4649,3472,352,0,250,333,333;btnTileReports,3969,4082,3472,992,0,250,333,333;boxNavSettings,15066,5216,3472,992,750,250,666,333""" & vbCrLf
    s = s & "    spec = spec & "";icoTileSettings,15066,5284,3472,482,750,250,666,333;lblTileSettings,15066,5783,3472,352,750,250,666,333;btnTileSettings,15066,5216,3472,992,750,250,666,333;boxNavUsers,11367,5216,3472,992,500,250,666,333;icoTileUsers,11367,5284,3472,482,500,250,666,333;lblTileUsers,11367,5783,3472,352,500,250,666,333""" & vbCrLf
    s = s & "    spec = spec & "";btnTileUsers,11367,5216,3472,992,500,250,666,333;boxNavTreasury,7668,5216,3472,992,250,250,666,333;icoTileTreasury,7668,5284,3472,482,250,250,666,333;lblTileTreasury,7668,5783,3472,352,250,250,666,333;btnTileTreasury,7668,5216,3472,992,250,250,666,333;boxNavAccounting,3969,5216,3472,992,0,250,666,333""" & vbCrLf
    s = s & "    spec = spec & "";icoTileAccounting,3969,5284,3472,482,0,250,666,333;lblTileAccounting,3969,5783,3472,352,0,250,666,333;btnTileAccounting,3969,5216,3472,992,0,250,666,333;boxTile5,15066,6378,3472,1049,750,250,1000,0;lblTileTitle5,15236,6435,3132,284,750,250,1000,0;lblTileValue5,15236,6730,3132,408,750,250,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub5,15236,7143,3132,255,750,250,1000,0;boxTile6,11367,6378,3472,1049,500,250,1000,0;lblTileTitle6,11537,6435,3132,284,500,250,1000,0;lblTileValue6,11537,6730,3132,408,500,250,1000,0;lblTileSub6,11537,7143,3132,255,500,250,1000,0;boxTile7,7668,6378,3472,1049,250,250,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle7,7838,6435,3132,284,250,250,1000,0;lblTileValue7,7838,6730,3132,408,250,250,1000,0;lblTileSub7,7838,7143,3132,255,250,250,1000,0;boxTile8,3969,6378,3472,1049,0,250,1000,0;lblTileTitle8,4139,6435,3132,284,0,250,1000,0;lblTileValue8,4139,6730,3132,408,0,250,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub8,4139,7143,3132,255,0,250,1000,0;boxTile9,15066,7569,3472,1049,750,250,1000,0;lblTileTitle9,15236,7626,3132,284,750,250,1000,0;lblTileValue9,15236,7921,3132,408,750,250,1000,0;lblTileSub9,15236,8334,3132,255,750,250,1000,0;boxTile10,11367,7569,3472,1049,500,250,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileTitle10,11537,7626,3132,284,500,250,1000,0;lblTileValue10,11537,7921,3132,408,500,250,1000,0;lblTileSub10,11537,8334,3132,255,500,250,1000,0;boxTile11,7668,7569,3472,1049,250,250,1000,0;lblTileTitle11,7838,7626,3132,284,250,250,1000,0;lblTileValue11,7838,7921,3132,408,250,250,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblTileSub11,7838,8334,3132,255,250,250,1000,0;boxTile12,3969,7569,3472,1049,0,250,1000,0;lblTileTitle12,4139,7626,3132,284,0,250,1000,0;lblTileValue12,4139,7921,3132,408,0,250,1000,0;lblTileSub12,4139,8334,3132,255,0,250,1000,0;lblIntegrity,3969,9180,14571,312,0,1000,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9634, 0, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmMain", s
    Exit Sub
EH:
    AbortForm "frmMain", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmProducts()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmProducts", "«·„‰ Ã« ", "SELECT * FROM Products", 15309, 10347, True, True, True, _
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
    Set c = AddList("lstItems", 227, 3005, 4990, 7058, 4, "0;1134;2778;850", True)
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
    Set c = AddCheck("TrackStock", "TrackStock", 7201, 7456)
    SetCtlProp c, "ControlTipText", "√·€ˆ «·⁄·«„… ··ÊÃ»«  Ê«·„‘—Ê»«  «· Ì  ıÕ÷Û¯— ⁄‰œ «·ÿ·»:  ı»«⁄ »·« —’Ìœ"
    SetCtlProp c, "StatusBarText", "√·€ˆ «·⁄·«„… ··ÊÃ»«  Ê«·„‘—Ê»«  «· Ì  ıÕ÷Û¯— ⁄‰œ «·ÿ·»:  ı»«⁄ »·« —’Ìœ"
    Set c = AddLabel("lblTrackStock", "Ì «»⁄ «·„Œ“Ê‰", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "TrackStock", 0)
    Set c = AddText("SizePriceM", "SizePriceM", 12134, 7371, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "«·ﬂ«›ÌÂ: ”⁄— «·»Ì⁄ = «·’€Ì—∫ ”⁄— «·Ê”ÿ √Ê «·ﬂ»Ì— ÌÃ⁄· ··„‘—Ê» √ÕÃ«„«"
    SetCtlProp c, "StatusBarText", "«·ﬂ«›ÌÂ: ”⁄— «·»Ì⁄ = «·’€Ì—∫ ”⁄— «·Ê”ÿ √Ê «·ﬂ»Ì— ÌÃ⁄· ··„‘—Ê» √ÕÃ«„«"
    Set c = AddLabel("lblSizePriceM", "”⁄— «·ÕÃ„ «·Ê”ÿ", 10376, 7371, 1701, 425, 10, False, CLR_MUTED, "SizePriceM", 0)
    Set c = AddText("SizePriceL", "SizePriceL", 7201, 7938, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblSizePriceL", "”⁄— «·ÕÃ„ «·ﬂ»Ì—", 5443, 7938, 1701, 425, 10, False, CLR_MUTED, "SizePriceL", 0)
    Set c = AddText("ImagePath", "ImagePath", 12134, 7938, 1644, 425)
    SetCtlProp c, "ControlTipText", "’Ê—… «·“— ›Ì ‘«‘… «··„”: „”«— ﬂ«„· √Ê «”„ „·› ›Ì „Ã·œ «·’Ê—"
    SetCtlProp c, "StatusBarText", "’Ê—… «·“— ›Ì ‘«‘… «··„”: „”«— ﬂ«„· √Ê «”„ „·› ›Ì „Ã·œ «·’Ê—"
    Set c = AddLabel("lblImagePath", "’Ê—… «·„‰ Ã", 10376, 7938, 1701, 425, 10, False, CLR_MUTED, "ImagePath", 0)
    Set c = AddButton("btnBrowseImage", "«” ⁄—«÷", 13835, 7938, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("Notes", "Notes", 7201, 8505, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 8505, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 9667, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnBrowseImage_Click()" & vbCrLf
    s = s & "    BrowseFile Me, ""ImagePath""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmProducts", s
    Exit Sub
EH:
    AbortForm "frmProducts", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCustomers()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCustomers", "«·⁄„·«¡", "SELECT * FROM Customers", 15309, 9780, True, True, True, _
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
    Set c = AddButton("btnAging", "√⁄„«— «·œÌÊ‰", 9751, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAllocate", "—»ÿ «·”œ«œ", 11565, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 6491, 4, "0;2495;1361;907", True)
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
    Set c = AddText("PaymentTermsDays", "PaymentTermsDays", 7201, 6804, 2948, 425)
    SetCtlProp c, "ControlTipText", "«” Õﬁ«ﬁ «·›« Ê—… «·¬Ã·… =  «—ÌŒÂ« + Â–Â «·„œ…"
    SetCtlProp c, "StatusBarText", "«” Õﬁ«ﬁ «·›« Ê—… «·¬Ã·… =  «—ÌŒÂ« + Â–Â «·„œ…"
    Set c = AddLabel("lblPaymentTermsDays", "„œ… «·”œ«œ (ÌÊ„)", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "PaymentTermsDays", 0)
    Set c = AddCombo("SalesRepID", "SalesRepID", 12134, 6804, 2948, 425, "SELECT SalesRepID, RepName FROM SalesReps WHERE IsActive = True ORDER BY RepName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", " ı‰”» ·Â ›Ê« Ì— «·⁄„Ì· Ê Õ’Ì·« Â"
    SetCtlProp c, "StatusBarText", " ı‰”» ·Â ›Ê« Ì— «·⁄„Ì· Ê Õ’Ì·« Â"
    Set c = AddLabel("lblSalesRepID", "«·„‰œÊ»", 10376, 6804, 1701, 425, 10, False, CLR_MUTED, "SalesRepID", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 7456)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "«·—’Ìœ «·„ÊÃ» = „»·€ „” Õﬁ ⁄·Ï «·⁄„Ì·", 10376, 7371, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 7938, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 7938, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 9100, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnAging_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAging"", 0, ""C|"" & Me!CustomerID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAllocate_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAllocation"", 0, ""C|"" & Me!CustomerID" & vbCrLf
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
    StartForm "frmSuppliers", "«·„Ê—œÊ‰", "SELECT * FROM Suppliers", 15309, 8646, True, True, True, _
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
    Set c = AddButton("btnAging", "√⁄„«— «·œÌÊ‰", 9751, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAllocate", "—»ÿ «·”œ«œ", 11565, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 5357, 4, "0;2495;1361;907", True)
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
    Set c = AddText("PaymentTermsDays", "PaymentTermsDays", 7201, 5670, 2948, 425)
    SetCtlProp c, "ControlTipText", "«” Õﬁ«ﬁ ›« Ê—… «·‘—«¡ «·¬Ã·… =  «—ÌŒÂ« + Â–Â «·„œ…"
    SetCtlProp c, "StatusBarText", "«” Õﬁ«ﬁ ›« Ê—… «·‘—«¡ «·¬Ã·… =  «—ÌŒÂ« + Â–Â «·„œ…"
    Set c = AddLabel("lblPaymentTermsDays", "„œ… «·”œ«œ (ÌÊ„)", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "PaymentTermsDays", 0)
    Set c = AddCombo("CurrencyCode", "CurrencyCode", 12134, 5670, 2948, 425, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ControlTipText", " ıﬁ —Õ ›Ì ›Ê« Ì—Â Ê”‰œ« Â"
    SetCtlProp c, "StatusBarText", " ıﬁ —Õ ›Ì ›Ê« Ì—Â Ê”‰œ« Â"
    Set c = AddLabel("lblCurrencyCode", "⁄„·… «· ⁄«„·", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "CurrencyCode", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 6322)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "«·—’Ìœ «·„ÊÃ» = „»·€ „” Õﬁ ··„Ê—œ (»«·—Ì«· œ«∆„«)", 10376, 6237, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 6804, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 7966, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnAging_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAging"", 0, ""S|"" & Me!SupplierID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAllocate_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAllocation"", 0, ""S|"" & Me!SupplierID" & vbCrLf
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
    Set c = AddButton("btnExpenseTypes", "√‰Ê«⁄ «·„’—Ê›« ", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnTreasury", "«·Œ“Ì‰…", 7937, 1021, 1701, 482, "secondary")
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
    c.AfterUpdate = EP
    Set c = AddLabel("lblExpenseDate", " «—ÌŒ «·„’—Ê›", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseDate", 0)
    Set c = AddCombo("CurrencyCode", "CurrencyCode", 7201, 2268, 2948, 425, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ControlTipText", "«·„’—Ê› »⁄„·… √Œ—Ï: «ﬂ » „»·€Â »Â«"
    SetCtlProp c, "StatusBarText", "«·„’—Ê› »⁄„·… √Œ—Ï: «ﬂ » „»·€Â »Â«"
    c.AfterUpdate = EP
    Set c = AddLabel("lblCurrencyCode", "«·⁄„·…", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "CurrencyCode", 0)
    Set c = AddText("ExchangeRate", "ExchangeRate", 12134, 2268, 2948, 425)
    SetCtlProp c, "Format", "0.00%"
    SetCtlProp c, "ControlTipText", "ﬁÌ„… ÊÕœ… Ê«Õœ… »«·—Ì«·"
    SetCtlProp c, "StatusBarText", "ﬁÌ„… ÊÕœ… Ê«Õœ… »«·—Ì«·"
    c.AfterUpdate = EP
    Set c = AddLabel("lblExchangeRate", "„⁄«„· «· ÕÊÌ·", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "ExchangeRate", 0)
    Set c = AddText("ForeignAmount", "ForeignAmount", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblForeignAmount", "«·„»·€ »«·⁄„·… (»œÊ‰ «·÷—Ì»…)", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "ForeignAmount", 0)
    Set c = AddText("ForeignTax", "ForeignTax", 12134, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblForeignTax", "«·÷—Ì»… »«·⁄„·…", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "ForeignTax", 0)
    Set c = AddCombo("ExpenseTypeID", "ExpenseTypeID", 7201, 3402, 1644, 425, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName", 2, "0;3402")
    Set c = AddLabel("lblExpenseTypeID", "‰Ê⁄ «·„’—Ê› *", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeID", 0)
    Set c = AddButton("btnNewType", "‰Ê⁄ ÃœÌœ", 8902, 3402, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddCombo("PaymentMethodID", "PaymentMethodID", 12134, 3402, 2948, 425, "SELECT PaymentMethodID, MethodName FROM PaymentMethods ORDER BY SortOrder", 2, "0;3402")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentMethodID", "ÿ—Ìﬁ… «·œ›⁄", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "PaymentMethodID", 0)
    Set c = AddText("Amount", "Amount", 7201, 3969, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblAmount", "«·„»·€ ﬁ»· «·÷—Ì»…", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "Amount", 0)
    Set c = AddText("Tax", "Tax", 12134, 3969, 1644, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTax", "÷—Ì»… «·„œŒ·« ", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "Tax", 0)
    Set c = AddButton("btnCalcVat", "«Õ”» 15%", 13835, 3969, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("TotalAmount", "TotalAmount", 7201, 4536, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblTotalAmount", "«·≈Ã„«·Ì", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "TotalAmount", 0)
    Set c = AddText("SupplierInvoiceRef", "SupplierInvoiceRef", 12134, 4536, 2948, 425)
    Set c = AddLabel("lblSupplierInvoiceRef", "—ﬁ„ ›« Ê—… «·„’—Ê›", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "SupplierInvoiceRef", 0)
    Set c = AddCombo("CashBoxID", "CashBoxID", 7201, 5103, 2948, 425, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "«·„’—Ê› «·‰ﬁœÌ ÌıŒ’„ „‰ Â–« «·’‰œÊﬁ (ÌıŒ «— ’‰œÊﬁﬂ  ·ﬁ«∆Ì«)"
    SetCtlProp c, "StatusBarText", "«·„’—Ê› «·‰ﬁœÌ ÌıŒ’„ „‰ Â–« «·’‰œÊﬁ (ÌıŒ «— ’‰œÊﬁﬂ  ·ﬁ«∆Ì«)"
    Set c = AddLabel("lblCashBoxID", "’ı—› „‰ ’‰œÊﬁ", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "CashBoxID", 0)
    Set c = AddCombo("BankID", "BankID", 12134, 5103, 2948, 425, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "«· ÕÊÌ· «·»‰ﬂÌ ÌıŒ’„ „‰ Â–« «·»‰ﬂ («·»‰ﬂ «·«› —«÷Ì  ·ﬁ«∆Ì«)"
    SetCtlProp c, "StatusBarText", "«· ÕÊÌ· «·»‰ﬂÌ ÌıŒ’„ „‰ Â–« «·»‰ﬂ («·»‰ﬂ «·«› —«÷Ì  ·ﬁ«∆Ì«)"
    Set c = AddLabel("lblBankID", "«·»‰ﬂ", 10376, 5103, 1701, 425, 10, False, CLR_MUTED, "BankID", 0)
    Set c = AddCombo("CostCenterID", "CostCenterID", 7201, 5670, 2948, 425, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "›«—€ = „—ﬂ“ «·„” Œœ„ √Ê «·„—ﬂ“ «·«› —«÷Ì"
    SetCtlProp c, "StatusBarText", "›«—€ = „—ﬂ“ «·„” Œœ„ √Ê «·„—ﬂ“ «·«› —«÷Ì"
    Set c = AddLabel("lblCostCenterID", "„—ﬂ“ «· ﬂ·›…", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "CostCenterID", 0)
    Set c = AddText("Description", "Description", 7201, 6237, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "«·Ê’›", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
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
    s = s & "Private Sub btnExpenseTypes_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenseTypes""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTreasury_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmTreasury""" & vbCrLf
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
    s = s & "Private Sub ExpenseDate_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""ExpenseDate""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CurrencyCode_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""CurrencyCode""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub ExchangeRate_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""ExchangeRate""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub ForeignAmount_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""ForeignAmount""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub ForeignTax_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""ForeignTax""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewType_Click()" & vbCrLf
    s = s & "    AddExpenseType Me, ""ExpenseTypeID""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub PaymentMethodID_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""PaymentMethodID""" & vbCrLf
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

Private Sub BuildForm_frmCurrencies()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCurrencies", "«·⁄„·« ", "SELECT * FROM Currencies", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Currencies|PK=CurrencyCode|LIST=SELECT t.CurrencyCode, t.CurrencyCode AS [«·—„“], t.CurrencyName AS [«·⁄„·…] FROM Currencies AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.SortOrder, t.CurrencyCode|SEARCH=t.CurrencyCode,t.CurrencyName,t.CurrencyNameEn|ACTIVE=t.IsActive|UNIQUE=CurrencyCode,CurrencyName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·⁄„·« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "⁄„·«  «· ⁄«„·∫ ⁄„·… «·»—‰«„Ã «·—Ì«· Êﬂ· «·„»«·€  ıÕ›Ÿ »Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnRates", "√”⁄«— «·⁄„·« ", 6123, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 3, "0;907;3856", True)
    c.AfterUpdate = EP
    Set c = AddText("CurrencyCode", "CurrencyCode", 7201, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "3 √Õ—› (ISO) „À· USD"
    SetCtlProp c, "StatusBarText", "3 √Õ—› (ISO) „À· USD"
    Set c = AddLabel("lblCurrencyCode", "—„“ «·⁄„·… *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CurrencyCode", 0)
    Set c = AddText("CurrencyName", "CurrencyName", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblCurrencyName", "«”„ «·⁄„·… *", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "CurrencyName", 0)
    Set c = AddText("CurrencyNameEn", "CurrencyNameEn", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblCurrencyNameEn", "«”„ «·⁄„·… »«·≈‰Ã·Ì“Ì…", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "CurrencyNameEn", 0)
    Set c = AddText("Symbol", "Symbol", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblSymbol", "«·—„“ «·„Œ ’—", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Symbol", 0)
    Set c = AddText("DecimalPlaces", "DecimalPlaces", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblDecimalPlaces", "⁄œœ «·Œ«‰«  «·⁄‘—Ì…", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "DecimalPlaces", 0)
    Set c = AddText("SortOrder", "SortOrder", 12134, 2835, 2948, 425)
    Set c = AddLabel("lblSortOrder", "«· — Ì»", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "SortOrder", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 3487)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblCurrencyNote", "«·„⁄«„· = ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »«·—Ì«·∫ Ìı”ÃÛ¯· ·ﬂ·  «—ÌŒ ›Ì ´√”⁄«— «·⁄„·« ª", 10376, 3402, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 4082, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnRates_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCurrencyRates"", 0" & vbCrLf
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
    FinishForm "frmCurrencies", s
    Exit Sub
EH:
    AbortForm "frmCurrencies", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCurrencyRates()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCurrencyRates", "√”⁄«— «·⁄„·« ", "SELECT * FROM CurrencyRates", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=CurrencyRates|PK=CurrencyRateID|LIST=SELECT t.CurrencyRateID, t.CurrencyCode AS [«·⁄„·…], t.RateDate AS [«· «—ÌŒ], t.Rate AS [«·„⁄«„·] FROM CurrencyRates AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.RateDate DESC, t.CurrencyCode|SEARCH=t.CurrencyCode,t.Notes"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "√”⁄«— «·⁄„·« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„⁄«„· ﬂ· ⁄„·… ›Ì  «—ÌŒ∫ «·„” ‰œ Ì√Œ– ¬Œ— ”⁄— ›Ì  «—ÌŒÂ √Ê ﬁ»·Â", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
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
    Set c = AddList("lstItems", 227, 2551, 4990, 5387, 4, "0;1021;1474;1247", True)
    c.AfterUpdate = EP
    Set c = AddCombo("CurrencyCode", "CurrencyCode", 7201, 1701, 2948, 425, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    Set c = AddLabel("lblCurrencyCode", "«·⁄„·… *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CurrencyCode", 0)
    Set c = AddText("RateDate", "RateDate", 12134, 1701, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblRateDate", "«· «—ÌŒ", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "RateDate", 0)
    Set c = AddText("Rate", "Rate", 7201, 2268, 2948, 425)
    SetCtlProp c, "Format", "0.00%"
    SetCtlProp c, "ControlTipText", "„À«·: «·œÊ·«— 3.75"
    SetCtlProp c, "StatusBarText", "„À«·: «·œÊ·«— 3.75"
    Set c = AddLabel("lblRate", "«·„⁄«„·", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Rate", 0)
    Set c = AddText("Notes", "Notes", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblRateNote", "·ﬂ· ⁄„·… ”⁄— Ê«Õœ ›Ì «·ÌÊ„∫ ⁄„·… «·»—‰«„Ã ·«  Õ «Ã ”⁄—«", 5443, 2835, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 3515, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    FinishForm "frmCurrencyRates", s
    Exit Sub
EH:
    AbortForm "frmCurrencyRates", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSalesReps()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSalesReps", "«·„‰œÊ»Ì‰", "SELECT * FROM SalesReps", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=SalesReps|PK=SalesRepID|LIST=SELECT t.SalesRepID, t.RepCode AS [«·ﬂÊœ], t.RepName AS [«·„‰œÊ»], t.Region AS [«·„‰ÿﬁ…] FROM SalesReps AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.RepName|SEARCH=t.RepCode,t.RepName,t.RepNameEn,t.Mobile,t.Region|ACTIVE=t.IsActive|SEQ=SALES_REP:RepCode|UNIQUE=RepCode,RepName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„‰œÊ»Ì‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„‰œÊ»Ê «·„»Ì⁄« : ⁄„·«ƒÂ„ Ê‰”»… ⁄„Ê· Â„ Ê„—ﬂ“  ﬂ·› Â„", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnTargets", "«·√Âœ«›", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCommissions", "«·⁄„Ê·« ", 7937, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnRepReport", "√œ«¡ «·„‰œÊ»Ì‰", 9751, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;1021;2381;1361", True)
    c.AfterUpdate = EP
    Set c = AddText("RepCode", "RepCode", 7201, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "ÌıÊ·Û¯œ  ·ﬁ«∆Ì« ≈–«  ı—ﬂ ›«—€«"
    SetCtlProp c, "StatusBarText", "ÌıÊ·Û¯œ  ·ﬁ«∆Ì« ≈–«  ı—ﬂ ›«—€«"
    Set c = AddLabel("lblRepCode", "ﬂÊœ «·„‰œÊ»", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "RepCode", 0)
    Set c = AddText("Mobile", "Mobile", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblMobile", "«·ÃÊ«·", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("RepName", "RepName", 7201, 2268, 7881, 425)
    Set c = AddLabel("lblRepName", "«”„ «·„‰œÊ» *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "RepName", 0)
    Set c = AddText("RepNameEn", "RepNameEn", 7201, 2835, 7881, 425)
    Set c = AddLabel("lblRepNameEn", "«·«”„ »«·≈‰Ã·Ì“Ì…", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "RepNameEn", 0)
    Set c = AddCombo("EmployeeID", "EmployeeID", 7201, 3402, 2948, 425, "SELECT EmployeeID, EmployeeName FROM Employees WHERE IsActive = True ORDER BY EmployeeName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "„»Ì⁄«  Â–« «·„” Œœ„ ·⁄„Ì· »·« „‰œÊ»  ı‰”» ··„‰œÊ»"
    SetCtlProp c, "StatusBarText", "„»Ì⁄«  Â–« «·„” Œœ„ ·⁄„Ì· »·« „‰œÊ»  ı‰”» ··„‰œÊ»"
    Set c = AddLabel("lblEmployeeID", "„” Œœ„ «·»—‰«„Ã", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "EmployeeID", 0)
    Set c = AddText("Region", "Region", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblRegion", "«·„‰ÿﬁ… / Œÿ «·”Ì—", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "Region", 0)
    Set c = AddText("CommissionRate", "CommissionRate", 7201, 3969, 2948, 425)
    SetCtlProp c, "Format", "0.00%"
    SetCtlProp c, "ControlTipText", "„À«·: 2%  ıﬂ » 0.02"
    SetCtlProp c, "StatusBarText", "„À«·: 2%  ıﬂ » 0.02"
    Set c = AddLabel("lblCommissionRate", "‰”»… «·⁄„Ê·…", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "CommissionRate", 0)
    Set c = AddCombo("CommissionBase", "CommissionBase", 12134, 3969, 2948, 425, "SALES;’«›Ì «·„»Ì⁄«  (»œÊ‰ «·÷—Ì»…);COLLECTION;«· Õ’Ì·", 2, "0;2552")
    Set c = AddLabel("lblCommissionBase", "√”«” «·⁄„Ê·…", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "CommissionBase", 0)
    Set c = AddCombo("CostCenterID", "CostCenterID", 7201, 4536, 2948, 425, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "„—ﬂ“ ﬁÌœ ⁄„Ê· Â"
    SetCtlProp c, "StatusBarText", "„—ﬂ“ ﬁÌœ ⁄„Ê· Â"
    Set c = AddLabel("lblCostCenterID", "„—ﬂ“ «· ﬂ·›…", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "CostCenterID", 0)
    Set c = AddCheck("IsActive", "IsActive", 12134, 4621)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblRepInfo", " ", 5443, 5103, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 5670, 7881, 425)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 6350, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnTargets_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmRepTargets"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCommissions_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCommissions"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRepReport_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmReportCenter"", 0, ""REP_PERFORMANCE""" & vbCrLf
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
    FinishForm "frmSalesReps", s
    Exit Sub
EH:
    AbortForm "frmSalesReps", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmRepTargets()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmRepTargets", "√Âœ«› «·„‰œÊ»Ì‰", "SELECT * FROM SalesRepTargets", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=SalesRepTargets|PK=TargetID|LIST=SELECT t.TargetID, s.RepName AS [«·„‰œÊ»], t.TargetYear & '/' & t.TargetMonth AS [«·‘Â—], t.TargetAmount AS [«·Âœ›] FROM SalesRepTargets AS t INNER JOIN SalesReps AS s ON t.SalesRepID = s.SalesRepID WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.TargetYear DESC, t.TargetMonth DESC, s.RepName|SEARCH=s.RepName,s.RepCode"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "√Âœ«› «·„‰œÊ»Ì‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·Âœ› «·‘Â—Ì ·’«›Ì „»Ì⁄«  ﬂ· „‰œÊ» (»œÊ‰ «·÷—Ì»…)", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
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
    Set c = AddList("lstItems", 227, 2551, 4990, 5387, 4, "0;2268;1134;1361", True)
    c.AfterUpdate = EP
    Set c = AddCombo("SalesRepID", "SalesRepID", 7201, 1701, 2948, 425, "SELECT SalesRepID, RepName FROM SalesReps WHERE IsActive = True ORDER BY RepName", 2, "0;3402")
    Set c = AddLabel("lblSalesRepID", "«·„‰œÊ» *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "SalesRepID", 0)
    Set c = AddText("TargetYear", "TargetYear", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblTargetYear", "«·”‰… *", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "TargetYear", 0)
    Set c = AddText("TargetMonth", "TargetMonth", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblTargetMonth", "«·‘Â—", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "TargetMonth", 0)
    Set c = AddText("TargetAmount", "TargetAmount", 12134, 2268, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblTargetAmount", "«·Âœ›", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "TargetAmount", 0)
    Set c = AddLabel("lblTargetNote", "·ﬂ· „‰œÊ» Âœ› Ê«Õœ ›Ì «·‘Â—∫  ﬁ—Ì— ´√œ«¡ «·„‰œÊ»Ì‰ª Ìﬁ«—‰ «·›⁄·Ì »«·Âœ›", 5443, 2835, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 3515, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    FinishForm "frmRepTargets", s
    Exit Sub
EH:
    AbortForm "frmRepTargets", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmRecurring()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmRecurring", "«·„’—Ê›«  «·„ ﬂ——…", "SELECT * FROM RecurringExpenses", 15309, 8646, True, True, True, _
              "KIND=LIST|TABLE=RecurringExpenses|PK=RecurringID|LIST=SELECT t.RecurringID, t.RecurringName AS [«·„’—Ê›], IIf(t.Frequency = 'MONTHLY', '‘Â—Ì', IIf(t.Frequency = 'QUARTERLY', '—»⁄ ”‰ÊÌ', '”‰ÊÌ')) AS [«· ﬂ—«—], t.Amount + t.Tax AS [«·„»·€], t.NextDueDate AS [«·„” Õﬁ] FROM RecurringExpenses AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.NextDueDate, t.RecurringName|SEARCH=t.RecurringName,t.Description|ACTIVE=t.IsActive|UNIQUE=RecurringName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8C7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„’—Ê›«  «·„ ﬂ——…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·≈ÌÃ«— Ê«·ﬂÂ—»«¡ Ê«·«‘ —«ﬂ« : Ìı‰‘√ «·„’—Ê›  ·ﬁ«∆Ì« ›Ì  «—ÌŒ «” Õﬁ«ﬁÂ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnCreateDue", "≈‰‘«¡ «·„” Õﬁ «·¬‰", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnExpenses", "«·„’—Ê›« ", 7937, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 5357, 5, "0;1928;907;907;1134", True)
    c.AfterUpdate = EP
    Set c = AddText("RecurringName", "RecurringName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblRecurringName", "«”„ «·„’—Ê› *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "RecurringName", 0)
    Set c = AddCombo("ExpenseTypeID", "ExpenseTypeID", 7201, 2268, 2948, 425, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName", 2, "0;3402")
    Set c = AddLabel("lblExpenseTypeID", "‰Ê⁄ «·„’—Ê› *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeID", 0)
    Set c = AddCombo("PaymentMethodID", "PaymentMethodID", 12134, 2268, 2948, 425, "SELECT PaymentMethodID, MethodName FROM PaymentMethods ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethodID", "ÿ—Ìﬁ… «·œ›⁄ *", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "PaymentMethodID", 0)
    Set c = AddText("Amount", "Amount", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ ﬁ»· «·÷—Ì»…", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Amount", 0)
    Set c = AddText("Tax", "Tax", 12134, 2835, 1644, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblTax", "÷—Ì»… «·„œŒ·« ", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Tax", 0)
    Set c = AddButton("btnCalcVat", "«Õ”» 15%", 13835, 2835, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddCombo("Frequency", "Frequency", 7201, 3402, 2948, 425, "MONTHLY;‘Â—Ì;QUARTERLY;ﬂ· 3 √‘Â—;YEARLY;”‰ÊÌ", 2, "0;2268")
    Set c = AddLabel("lblFrequency", "«· ﬂ—«—", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "Frequency", 0)
    Set c = AddText("DueDay", "DueDay", 12134, 3402, 2948, 425)
    SetCtlProp c, "ControlTipText", "„‰ 1 ≈·Ï 28 (ÌÊ„ «·≈ÌÃ«— „À·«)"
    SetCtlProp c, "StatusBarText", "„‰ 1 ≈·Ï 28 (ÌÊ„ «·≈ÌÃ«— „À·«)"
    Set c = AddLabel("lblDueDay", "ÌÊ„ «·«” Õﬁ«ﬁ", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "DueDay", 0)
    Set c = AddText("StartDate", "StartDate", 7201, 3969, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblStartDate", "Ì»œ√ „‰", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "StartDate", 0)
    Set c = AddText("EndDate", "EndDate", 12134, 3969, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "ControlTipText", "›«—€ = »·« ‰Â«Ì…"
    SetCtlProp c, "StatusBarText", "›«—€ = »·« ‰Â«Ì…"
    Set c = AddLabel("lblEndDate", "Ì‰ ÂÌ ›Ì", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "EndDate", 0)
    Set c = AddText("NextDueDate", "NextDueDate", 7201, 4536, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "ControlTipText", "ÌÕ”»Â «·»—‰«„Ã"
    SetCtlProp c, "StatusBarText", "ÌÕ”»Â «·»—‰«„Ã"
    Set c = AddLabel("lblNextDueDate", "«·«” Õﬁ«ﬁ «· «·Ì", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "NextDueDate", 0)
    Set c = AddText("LastCreatedDate", "LastCreatedDate", 12134, 4536, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblLastCreatedDate", "¬Œ— „’—Ê› √ı‰‘∆", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "LastCreatedDate", 0)
    Set c = AddCombo("CashBoxID", "CashBoxID", 7201, 5103, 2948, 425, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "··œ›⁄ «·‰ﬁœÌ∫ ›«—€ = ’‰œÊﬁ „‰ Ìı‰‘∆ «·„’—Ê›"
    SetCtlProp c, "StatusBarText", "··œ›⁄ «·‰ﬁœÌ∫ ›«—€ = ’‰œÊﬁ „‰ Ìı‰‘∆ «·„’—Ê›"
    Set c = AddLabel("lblCashBoxID", "„‰ ’‰œÊﬁ", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "CashBoxID", 0)
    Set c = AddCombo("BankID", "BankID", 12134, 5103, 2948, 425, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "·· ÕÊÌ· «·»‰ﬂÌ∫ ›«—€ = «·»‰ﬂ «·«› —«÷Ì"
    SetCtlProp c, "StatusBarText", "·· ÕÊÌ· «·»‰ﬂÌ∫ ›«—€ = «·»‰ﬂ «·«› —«÷Ì"
    Set c = AddLabel("lblBankID", "„‰ »‰ﬂ", 10376, 5103, 1701, 425, 10, False, CLR_MUTED, "BankID", 0)
    Set c = AddCombo("CostCenterID", "CostCenterID", 7201, 5670, 2948, 425, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "›«—€ = „—ﬂ“ „‰ Ìı‰‘∆ «·„’—Ê›"
    SetCtlProp c, "StatusBarText", "›«—€ = „—ﬂ“ „‰ Ìı‰‘∆ «·„’—Ê›"
    Set c = AddLabel("lblCostCenterID", "„—ﬂ“ «· ﬂ·›…", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "CostCenterID", 0)
    Set c = AddCheck("IsActive", "IsActive", 12134, 5755)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblRecurringInfo", " ", 5443, 6237, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("Description", "Description", 7201, 6804, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "«·Ê’›", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 7966, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnCreateDue_Click()" & vbCrLf
    s = s & "    RecurringCreateNow Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnExpenses_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmExpenses""" & vbCrLf
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
    s = s & "Private Sub btnCalcVat_Click()" & vbCrLf
    s = s & "    CalcExpenseVat Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmRecurring", s
    Exit Sub
EH:
    AbortForm "frmRecurring", Err.Number, Err.Description
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
    Set c = AddButton("btnRoles", "’·«ÕÌ«  «·√œÊ«—", 8277, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnUserScreens", "’·«ÕÌ«  «·‘«‘« ", 10091, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAuditLog", "”Ã· «· œﬁÌﬁ", 11905, 1021, 1701, 482, "secondary")
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
    Set c = AddCombo("CashBoxID", "CashBoxID", 7201, 3969, 2948, 425, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "‰ﬁœÌ… „»Ì⁄«  «·„” Œœ„ Ê”‰œ« Â  œŒ· Â–« «·’‰œÊﬁ∫ ›«—€ = √Ê· ’‰œÊﬁ ﬂ«‘Ì—"
    SetCtlProp c, "StatusBarText", "‰ﬁœÌ… „»Ì⁄«  «·„” Œœ„ Ê”‰œ« Â  œŒ· Â–« «·’‰œÊﬁ∫ ›«—€ = √Ê· ’‰œÊﬁ ﬂ«‘Ì—"
    Set c = AddLabel("lblCashBoxID", "’‰œÊﬁ «·‰ﬁœÌ…", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "CashBoxID", 0)
    Set c = AddLabel("lblCashBoxNote", "«·„œÌ— «·„”ƒÊ· ⁄‰ «·Œ“Ì‰…: «Œ — ·Â «·Œ“Ì‰… «·—∆Ì”Ì…", 10376, 3969, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddCheck("MustChangePassword", "MustChangePassword", 7201, 4621)
    Set c = AddLabel("lblMustChangePassword", "ÌÃ»  €ÌÌ— ﬂ·„… «·„—Ê—", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "MustChangePassword", 0)
    Set c = AddText("LastLoginAt", "LastLoginAt", 12134, 4536, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblLastLoginAt", "¬Œ— œŒÊ·", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "LastLoginAt", 0)
    Set c = AddText("FailedLoginCount", "FailedLoginCount", 7201, 5103, 2948, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblFailedLoginCount", "„Õ«Ê·«  «·œŒÊ· «·›«‘·…", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "FailedLoginCount", 0)
    Set c = AddText("LockedUntil", "LockedUntil", 12134, 5103, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblLockedUntil", "„ﬁ›· Õ Ï", 10376, 5103, 1701, 425, 10, False, CLR_MUTED, "LockedUntil", 0)
    Set c = AddLabel("lblPasswordState", " ", 5443, 5670, 9639, 425, 10, True, CLR_ACCENT, "", 0)
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
    s = s & "Private Sub btnSetPassword_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmChangePassword"", 10, Me!EmployeeID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnlock_Click()" & vbCrLf
    s = s & "    UnlockUser Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRoles_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmRoles"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUserScreens_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmUserScreens"", 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAuditLog_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAuditLog"", 0" & vbCrLf
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

Private Sub BuildForm_frmCostCenters()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCostCenters", "„—«ﬂ“ «· ﬂ·›…", "SELECT * FROM CostCenters", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=CostCenters|PK=CostCenterID|LIST=SELECT t.CostCenterID, t.CenterCode AS [«·—„“], t.CenterName AS [«·„—ﬂ“], IIf(t.IsDefault, '«› —«÷Ì', '') AS [«·Õ«·…] FROM CostCenters AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.CenterCode|SEARCH=t.CenterCode,t.CenterName|ACTIVE=t.IsActive|UNIQUE=CenterCode,CenterName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "„—«ﬂ“ «· ﬂ·›…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·›—Ê⁄ Ê«·√ﬁ”«„:  ıÊ“Û¯⁄ ⁄·ÌÂ« «·≈Ì—«œ«  Ê«·„’—Ê›« ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;1021;2835;907", True)
    c.AfterUpdate = EP
    Set c = AddText("CenterCode", "CenterCode", 7201, 1701, 2948, 425)
    Set c = AddLabel("lblCenterCode", "—„“ «·„—ﬂ“ *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CenterCode", 0)
    Set c = AddText("CenterName", "CenterName", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblCenterName", "«”„ «·„—ﬂ“ *", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "CenterName", 0)
    Set c = AddCheck("IsDefault", "IsDefault", 7201, 2353)
    SetCtlProp c, "ControlTipText", "·„‰ ·« „—ﬂ“ ·Â „‰ «·„” Œœ„Ì‰"
    SetCtlProp c, "StatusBarText", "·„‰ ·« „—ﬂ“ ·Â „‰ «·„” Œœ„Ì‰"
    Set c = AddLabel("lblIsDefault", "«·„—ﬂ“ «·«› —«÷Ì", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "IsDefault", 0)
    Set c = AddCheck("IsActive", "IsActive", 12134, 2353)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblCenterNote", "„—ﬂ“ «·„ÊŸ› „‰ ‘«‘… —Ê« » «·„ÊŸ›Ì‰∫ «·„” ‰œ«  «·ﬁœÌ„… ´€Ì— „Ê“⁄…ª", 5443, 2835, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 3402, 7881, 425)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 4082, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    FinishForm "frmCostCenters", s
    Exit Sub
EH:
    AbortForm "frmCostCenters", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmEmployeePay()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmEmployeePay", "—Ê« » «·„ÊŸ›Ì‰", "SELECT * FROM Employees", 15309, 8222, True, False, True, _
              "KIND=LIST|TABLE=Employees|PK=EmployeeID|LIST=SELECT t.EmployeeID, t.EmployeeName AS [«·„ÊŸ›], IIf(t.OnPayroll, '‰⁄„', '') AS [›Ì «·„”Ì—], t.BasicSalary AS [«·√”«”Ì] FROM Employees AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.EmployeeName|SEARCH=t.EmployeeName,t.NationalID|ACTIVE=t.IsActive"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "—Ê« » «·„ÊŸ›Ì‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·—« » Ê«·»œ·«  Ê«· √„Ì‰«  Êﬁ”ÿ «·”·›… ·ﬂ· „ÊŸ›", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ", 227, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 1701, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayroll", "„”Ì— «·—Ê« »", 3175, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;2608;1021;1134", True)
    c.AfterUpdate = EP
    Set c = AddText("EmployeeName", "EmployeeName", 7201, 1701, 7881, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblEmployeeName", "«”„ «·„ÊŸ›", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "EmployeeName", 0)
    Set c = AddCheck("OnPayroll", "OnPayroll", 7201, 2353)
    Set c = AddLabel("lblOnPayroll", "›Ì „”Ì— «·—Ê« »", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "OnPayroll", 0)
    Set c = AddCheck("IsSaudi", "IsSaudi", 12134, 2353)
    Set c = AddLabel("lblIsSaudi", "”⁄ÊœÌ («· √„Ì‰«  »Õ’ Ì «·„ÊŸ› Ê«·„‰‘√…)", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "IsSaudi", 0)
    Set c = AddText("BasicSalary", "BasicSalary", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblBasicSalary", "«·—« » «·√”«”Ì", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "BasicSalary", 0)
    Set c = AddText("HousingAllowance", "HousingAllowance", 12134, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblHousingAllowance", "»œ· «·”ﬂ‰", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "HousingAllowance", 0)
    Set c = AddText("TransportAllowance", "TransportAllowance", 7201, 3402, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblTransportAllowance", "»œ· «·‰ﬁ·", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "TransportAllowance", 0)
    Set c = AddText("OtherAllowance", "OtherAllowance", 12134, 3402, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblOtherAllowance", "»œ·«  √Œ—Ï", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "OtherAllowance", 0)
    Set c = AddText("AdvanceInstallment", "AdvanceInstallment", 7201, 3969, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "0 = ÌıŒ’„ ﬂ· —’Ìœ «·”·› ›Ì √Ê· „”Ì—"
    SetCtlProp c, "StatusBarText", "0 = ÌıŒ’„ ﬂ· —’Ìœ «·”·› ›Ì √Ê· „”Ì—"
    Set c = AddLabel("lblAdvanceInstallment", "ﬁ”ÿ «·”·›… «·‘Â—Ì", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "AdvanceInstallment", 0)
    Set c = AddText("HireDate", "HireDate", 12134, 3969, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblHireDate", " «—ÌŒ «· ⁄ÌÌ‰", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "HireDate", 0)
    Set c = AddText("NationalID", "NationalID", 7201, 4536, 2948, 425)
    Set c = AddLabel("lblNationalID", "—ﬁ„ «·ÂÊÌ… / «·≈ﬁ«„…", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "NationalID", 0)
    Set c = AddText("IBAN", "IBAN", 12134, 4536, 2948, 425)
    Set c = AddLabel("lblIBAN", "¬Ì»«‰ «·„ÊŸ›", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "IBAN", 0)
    Set c = AddCombo("CostCenterID", "CostCenterID", 7201, 5103, 2948, 425, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "„»Ì⁄« Â Ê„”Ì— —« »Â ⁄·Ï Â–« «·„—ﬂ“"
    SetCtlProp c, "StatusBarText", "„»Ì⁄« Â Ê„”Ì— —« »Â ⁄·Ï Â–« «·„—ﬂ“"
    Set c = AddLabel("lblCostCenterID", "„—ﬂ“ «· ﬂ·›…", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "CostCenterID", 0)
    Set c = AddLabel("lblPayNote", "«· √„Ì‰«  ⁄·Ï «·√”«”Ì + «·”ﬂ‰: «·”⁄ÊœÌ »Õ’ Ì «·„ÊŸ› Ê«·„‰‘√…° Ê€Ì—Â »Õ’… «·„‰‘√…", 10376, 5103, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 5783, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnPayroll_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPayroll"", 0" & vbCrLf
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
    FinishForm "frmEmployeePay", s
    Exit Sub
EH:
    AbortForm "frmEmployeePay", Err.Number, Err.Description
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
    Set c = AddText("SortOrder", "SortOrder", 12134, 3317, 2948, 425)
    SetCtlProp c, "ControlTipText", " — Ì» «·“— ›Ì ‘«‘… «··„” («·√’€— √Ê·«)"
    SetCtlProp c, "StatusBarText", " — Ì» «·“— ›Ì ‘«‘… «··„” («·√’€— √Ê·«)"
    Set c = AddLabel("lblSortOrder", " — Ì» «·⁄—÷", 10376, 3317, 1701, 425, 10, False, CLR_MUTED, "SortOrder", 0)
    Set c = AddCheck("IsAddOn", "IsAddOn", 7201, 3969)
    SetCtlProp c, "ControlTipText", "«·ﬂ«›ÌÂ: √’‰«› Â–Â «·›∆…  ŸÂ— ﬂ≈÷«›«  ··„‘—Ê» (Õ·Ì»° ‘Ê  ≈÷«›Ì...)"
    SetCtlProp c, "StatusBarText", "«·ﬂ«›ÌÂ: √’‰«› Â–Â «·›∆…  ŸÂ— ﬂ≈÷«›«  ··„‘—Ê» (Õ·Ì»° ‘Ê  ≈÷«›Ì...)"
    Set c = AddLabel("lblIsAddOn", "›∆… ≈÷«›« ", 5443, 3884, 1701, 425, 10, False, CLR_MUTED, "IsAddOn", 0)
    Set c = AddCombo("TileColor", "TileColor", 12134, 3884, 2948, 425, "BLUE;√“—ﬁ;GREEN;√Œ÷—;ORANGE;»— ﬁ«·Ì;PURPLE;»‰›”ÃÌ;RED;√Õ„—;INDIGO;‰Ì·Ì;TEAL;›Ì—Ê“Ì;PINK;Ê—œÌ;BROWN;»‰Ì;GREY;—„«œÌ", 2, "0;3402")
    Set c = AddLabel("lblTileColor", "·Ê‰ «·“—", 10376, 3884, 1701, 425, 10, False, CLR_MUTED, "TileColor", 0)
    Set c = AddText("ImagePath", "ImagePath", 7201, 4451, 1644, 425)
    SetCtlProp c, "ControlTipText", "’Ê—… «·“— ›Ì ‘«‘… «··„”"
    SetCtlProp c, "StatusBarText", "’Ê—… «·“— ›Ì ‘«‘… «··„”"
    Set c = AddLabel("lblImagePath", "’Ê—… «· ’‰Ì›", 5443, 4451, 1701, 425, 10, False, CLR_MUTED, "ImagePath", 0)
    Set c = AddButton("btnBrowseImage", "«” ⁄—«÷", 8902, 4451, 1247, 425, "secondary")
    c.OnClick = EP
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
    s = s & "Private Sub chkShowInactive_AfterUpdate()" & vbCrLf
    s = s & "    RefreshList Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstItems_AfterUpdate()" & vbCrLf
    s = s & "    ListPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBrowseImage_Click()" & vbCrLf
    s = s & "    BrowseFile Me, ""ImagePath""" & vbCrLf
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

Private Sub BuildForm_frmCashBoxes()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCashBoxes", "«·’‰«œÌﬁ", "SELECT * FROM CashBoxes", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=CashBoxes|PK=CashBoxID|LIST=SELECT t.CashBoxID, t.BoxName AS [«·’‰œÊﬁ], IIf(t.BoxType = 'MAIN', 'Œ“Ì‰…', 'ﬂ«‘Ì—') AS [«·‰Ê⁄] FROM CashBoxes AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.BoxType DESC, t.BoxName|SEARCH=t.BoxName,t.Notes|ACTIVE=t.IsActive|UNIQUE=BoxName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·’‰«œÌﬁ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·Œ“Ì‰… «·—∆Ì”Ì… Ê’‰«œÌﬁ «·ﬂ«‘Ì—", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnTreasury", "«·Œ“Ì‰…", 6123, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 3, "0;3175;1588", True)
    c.AfterUpdate = EP
    Set c = AddText("BoxName", "BoxName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblBoxName", "«”„ «·’‰œÊﬁ *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "BoxName", 0)
    Set c = AddCombo("BoxType", "BoxType", 7201, 2268, 2948, 425, "MAIN;Œ“Ì‰… —∆Ì”Ì…;CASHIER;’‰œÊﬁ ﬂ«‘Ì—", 2, "0;2835")
    Set c = AddLabel("lblBoxType", "«·‰Ê⁄", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "BoxType", 0)
    Set c = AddCheck("IsActive", "IsActive", 12134, 2353)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "«·‰ﬁœÌ… «·„ÊÃÊœ… ›Ì «·’‰œÊﬁ ⁄‰œ »œ¡ «” Œœ«„ «·»—‰«„Ã"
    SetCtlProp c, "StatusBarText", "«·‰ﬁœÌ… «·„ÊÃÊœ… ›Ì «·’‰œÊﬁ ⁄‰œ »œ¡ «” Œœ«„ «·»—‰«„Ã"
    Set c = AddLabel("lblOpeningBalance", "«·—’Ìœ «·«›  «ÕÌ", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("OpeningDate", "OpeningDate", 12134, 2835, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblOpeningDate", " «—ÌŒ «·—’Ìœ «·«›  «ÕÌ", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "OpeningDate", 0)
    Set c = AddLabel("lblBoxNote", "«·—’Ìœ ·« Ìıﬂ » ÌœÊÌ«: ÌıÕ”» „‰ «·„»Ì⁄«  Ê«·”‰œ«  Ê«·„’—Ê›« ", 5443, 3402, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 3969, 7881, 425)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 4649, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnTreasury_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmTreasury""" & vbCrLf
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
    FinishForm "frmCashBoxes", s
    Exit Sub
EH:
    AbortForm "frmCashBoxes", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBanks()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBanks", "«·»‰Êﬂ", "SELECT * FROM Banks", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Banks|PK=BankID|LIST=SELECT t.BankID, t.BankName AS [«·»‰ﬂ], t.AccountNo AS [—ﬁ„ «·Õ”«»] FROM Banks AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.BankName|SEARCH=t.BankName,t.AccountNo,t.IBAN|ACTIVE=t.IsActive|UNIQUE=BankName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·»‰Êﬂ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·Õ”«»«  «·»‰ﬂÌ… ··„Õ· Ê√—’œ Â«", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnBankTx", "«·Õ—ﬂ«  «·»‰ﬂÌ…", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnRecon", "«· ”ÊÌ… «·»‰ﬂÌ…", 7937, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStatement", "ﬂ‘› Õ”«»", 9751, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCheques", "«·‘Ìﬂ« ", 11565, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 3, "0;2948;1814", True)
    c.AfterUpdate = EP
    Set c = AddText("BankName", "BankName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblBankName", "«”„ «·»‰ﬂ / «·Õ”«» *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "BankName", 0)
    Set c = AddText("AccountNo", "AccountNo", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblAccountNo", "—ﬁ„ «·Õ”«»", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "AccountNo", 0)
    Set c = AddText("IBAN", "IBAN", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblIBAN", "«·¬Ì»«‰", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "IBAN", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "—’Ìœ «·Õ”«» ›Ì «·»‰ﬂ ⁄‰œ »œ¡ «” Œœ«„ «·»—‰«„Ã"
    SetCtlProp c, "StatusBarText", "—’Ìœ «·Õ”«» ›Ì «·»‰ﬂ ⁄‰œ »œ¡ «” Œœ«„ «·»—‰«„Ã"
    Set c = AddLabel("lblOpeningBalance", "«·—’Ìœ «·«›  «ÕÌ", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("OpeningDate", "OpeningDate", 12134, 2835, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblOpeningDate", " «—ÌŒ «·—’Ìœ «·«›  «ÕÌ", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "OpeningDate", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 3487)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBankNote", "Õ”«» «·»‰ﬂ ›Ì «·œ·Ì· = 120000 + —ﬁ„Â° Ê«·—’Ìœ „‰ «·ﬁÌÊœ", 10376, 3402, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 3969, 7881, 425)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 4649, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnBankTx_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBankTx"", 0, Me!BankID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRecon_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBankRecon"", 0, Me!BankID" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStatement_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLedger"", 0, 120000 + Nz(Me!BankID, 0)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCheques_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCheques"", 0, ""IN""" & vbCrLf
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
    FinishForm "frmBanks", s
    Exit Sub
EH:
    AbortForm "frmBanks", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAccounts()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAccounts", "œ·Ì· «·Õ”«»« ", "SELECT * FROM Accounts", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Accounts|PK=AccountCode|LIST=SELECT t.AccountCode, t.AccountCode AS [«·—ﬁ„], Space((t.AccountLevel - 1) * 3) & t.AccountName AS [«·Õ”«»], IIf(t.IsPosting, '›—⁄Ì', '—∆Ì”Ì') AS [«·‰Ê⁄] FROM Accounts AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.TreeKey|SEARCH=t.AccountName|ACTIVE=t.IsActive|UNIQUE=AccountCode"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "œ·Ì· «·Õ”«»« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "‘Ã—… «·Õ”«»« : «·Õ”«»«  «·—∆Ì”Ì… Ê«·›—⁄Ì…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "ÃœÌœ", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", " —«Ã⁄", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnJournal", "ﬁÌÊœ «·ÌÊ„Ì…", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStatement", "ﬂ‘› Õ”«»", 7937, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnManual", "ﬁÌœ ÌœÊÌ", 9751, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCenters", "„—«ﬂ“ «· ﬂ·›…", 11565, 1021, 1701, 482, "secondary")
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
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;1021;3062;680", True)
    c.AfterUpdate = EP
    Set c = AddText("AccountCode", "AccountCode", 7201, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "—ﬁ„ ÃœÌœ ·« Ì ﬂ——∫ ·« Ì €Ì— »⁄œ «·Õ›Ÿ"
    SetCtlProp c, "StatusBarText", "—ﬁ„ ÃœÌœ ·« Ì ﬂ——∫ ·« Ì €Ì— »⁄œ «·Õ›Ÿ"
    Set c = AddLabel("lblAccountCode", "—ﬁ„ «·Õ”«» *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "AccountCode", 0)
    Set c = AddCombo("ParentCode", "ParentCode", 12134, 1701, 2948, 425, "SELECT AccountCode, Space((AccountLevel - 1) * 3) & AccountName AS Account FROM Accounts WHERE IsPosting = False ORDER BY TreeKey", 2, "0;3969")
    SetCtlProp c, "ControlTipText", "«·Õ”«» «·—∆Ì”Ì «·–Ì Ì »⁄Â (‰Ê⁄ «·Õ”«» Ì »⁄Â  ·ﬁ«∆Ì«)"
    SetCtlProp c, "StatusBarText", "«·Õ”«» «·—∆Ì”Ì «·–Ì Ì »⁄Â (‰Ê⁄ «·Õ”«» Ì »⁄Â  ·ﬁ«∆Ì«)"
    c.AfterUpdate = EP
    Set c = AddLabel("lblParentCode", "«·Õ”«» «·—∆Ì”Ì", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ParentCode", 0)
    Set c = AddText("AccountName", "AccountName", 7201, 2268, 7881, 425)
    Set c = AddLabel("lblAccountName", "«”„ «·Õ”«» *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "AccountName", 0)
    Set c = AddCombo("AccountType", "AccountType", 7201, 2835, 2948, 425, "ASSET;√’Ê·;LIABILITY;Œ’Ê„;EQUITY;ÕﬁÊﬁ „·ﬂÌ…;REVENUE;≈Ì—«œ« ;EXPENSE;„’—Ê›« ", 2, "0;2835")
    Set c = AddLabel("lblAccountType", "‰Ê⁄ «·Õ”«» *", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "AccountType", 0)
    Set c = AddCheck("IsPosting", "IsPosting", 12134, 2920)
    SetCtlProp c, "ControlTipText", "›—⁄Ì =  ıﬂ » ⁄·ÌÂ «·ﬁÌÊœ∫ —∆Ì”Ì = ÌÃ„⁄ Õ”«»« Â «· «»⁄… ›ﬁÿ"
    SetCtlProp c, "StatusBarText", "›—⁄Ì =  ıﬂ » ⁄·ÌÂ «·ﬁÌÊœ∫ —∆Ì”Ì = ÌÃ„⁄ Õ”«»« Â «· «»⁄… ›ﬁÿ"
    Set c = AddLabel("lblIsPosting", "Õ”«» ›—⁄Ì (Ìﬁ»· «·ﬁÌÊœ)", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "IsPosting", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 3487)
    Set c = AddLabel("lblIsActive", "‰‘ÿ", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddCheck("IsSystem", "IsSystem", 12134, 3487)
    SetCtlProp c, "Locked", True
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblIsSystem", "Õ”«» √”«”Ì ›Ì «·‰Ÿ«„", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "IsSystem", 0)
    Set c = AddText("AccountLevel", "AccountLevel", 7201, 3969, 2948, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblAccountLevel", "«·„” ÊÏ", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "AccountLevel", 0)
    Set c = AddLabel("lblAccountNote", "«·Õ”«»«  «·√”«”Ì… («·„⁄·Û¯„…)  ” Œœ„Â« «·ﬁÌÊœ «·¬·Ì…: ·«  ıÕ–› Ê·« Ì €Ì— ‰Ê⁄Â«. «·ﬁÌÊœ «·ÌœÊÌ… „‰ “— ´ﬁÌœ ÌœÊÌª.", 10376, 3969, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblStatus", " ", 5443, 4649, 9639, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnJournal_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmJournal""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStatement_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLedger"", 0, Me!AccountCode" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnManual_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmManualEntry""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCenters_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCostCenters""" & vbCrLf
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
    s = s & "Private Sub ParentCode_AfterUpdate()" & vbCrLf
    s = s & "    FieldChanged Me, ""ParentCode""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAccounts", s
    Exit Sub
EH:
    AbortForm "frmAccounts", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmSettings()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSettings", "«·≈⁄œ«œ« ", "SELECT * FROM Settings WHERE SettingID = 1", 15309, 10432, True, False, True, _
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
    Set c = AddButton("btnActivation", " ›⁄Ì· «·»—‰«„Ã", 10431, 1021, 1701, 482, "secondary")
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
    Set c = AddCombo("POSMode", "POSMode", 2552, 7371, 4989, 425, "RETAIL;«·„Õ·«  (»«—ﬂÊœ);RESTAURANT;«·„ÿ«⁄„ (‘«‘… ·„”);CAFE;«·ﬂ«›ÌÂ«  (‘«‘… ·„”)", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "«·‘«‘… «· Ì Ì› ÕÂ« “— «·„»Ì⁄« "
    SetCtlProp c, "StatusBarText", "«·‘«‘… «· Ì Ì› ÕÂ« “— «·„»Ì⁄« "
    Set c = AddLabel("lblPOSMode", "‘«‘… «·»Ì⁄", 227, 7371, 2268, 425, 10, False, CLR_MUTED, "POSMode", 0)
    Set c = AddText("ImagesFolder", "ImagesFolder", 10093, 7371, 3685, 425)
    SetCtlProp c, "ControlTipText", "›«—€ = „Ã·œ Images »Ã«‰» „·› «·»Ì«‰« "
    SetCtlProp c, "StatusBarText", "›«—€ = „Ã·œ Images »Ã«‰» „·› «·»Ì«‰« "
    Set c = AddLabel("lblImagesFolder", "„Ã·œ ’Ê— «·„‰ Ã« ", 7768, 7371, 2268, 425, 10, False, CLR_MUTED, "ImagesFolder", 0)
    Set c = AddButton("btnBrowseImages", "«” ⁄—«÷", 13835, 7371, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddCombo("InvoicePrintMode", "InvoicePrintMode", 2552, 7938, 4989, 425, "DIRECT;ÿ»«⁄… „»«‘—… »œÊ‰ „⁄«Ì‰…;PREVIEW;⁄—÷ „⁄«Ì‰… «·ÿ»«⁄…;NONE;»œÊ‰ ÿ»«⁄…", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "⁄‰œ Õ›Ÿ ›« Ê—… «·»Ì⁄ √Ê «·„— Ã⁄"
    SetCtlProp c, "StatusBarText", "⁄‰œ Õ›Ÿ ›« Ê—… «·»Ì⁄ √Ê «·„— Ã⁄"
    Set c = AddLabel("lblInvoicePrintMode", "«·ÿ»«⁄… ⁄‰œ Õ›Ÿ «·›« Ê—…", 227, 7938, 2268, 425, 10, False, CLR_MUTED, "InvoicePrintMode", 0)
    Set c = AddText("CreditBlockDays", "CreditBlockDays", 10093, 7938, 4989, 425)
    SetCtlProp c, "ControlTipText", "0 = ·« Ì Êﬁ› «·»Ì⁄ «·¬Ã· »”»» «· √ŒÌ—"
    SetCtlProp c, "StatusBarText", "0 = ·« Ì Êﬁ› «·»Ì⁄ «·¬Ã· »”»» «· √ŒÌ—"
    Set c = AddLabel("lblCreditBlockDays", "≈Ìﬁ«› «·»Ì⁄ «·¬Ã· ·⁄„Ì· „ √Œ— √ﬂÀ— „‰ (ÌÊ„)", 7768, 7938, 2268, 425, 10, False, CLR_MUTED, "CreditBlockDays", 0)
    Set c = AddCombo("DefaultBankID", "DefaultBankID", 2552, 8505, 4989, 425, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3402")
    SetCtlProp c, "ControlTipText", "«· ÕÊÌ·«  «·»‰ﬂÌ… ›Ì «·›Ê« Ì— Ê«·”‰œ«   ıﬁÌÛ¯œ ›ÌÂ"
    SetCtlProp c, "StatusBarText", "«· ÕÊÌ·«  «·»‰ﬂÌ… ›Ì «·›Ê« Ì— Ê«·”‰œ«   ıﬁÌÛ¯œ ›ÌÂ"
    Set c = AddLabel("lblDefaultBankID", "«·»‰ﬂ «·«› —«÷Ì ·· ÕÊÌ·«  «·»‰ﬂÌ…", 227, 8505, 2268, 425, 10, False, CLR_MUTED, "DefaultBankID", 0)
    Set c = AddLabel("lblStoreNameNote", " ", 7768, 8505, 7314, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddCheck("AllowAdminCompanyName", "AllowAdminCompanyName", 2552, 9157)
    SetCtlProp c, "ControlTipText", "ÌŸÂ— ··„»—„Ã ›ﬁÿ"
    SetCtlProp c, "StatusBarText", "ÌŸÂ— ··„»—„Ã ›ﬁÿ"
    Set c = AddLabel("lblAllowAdminCompanyName", "«·”„«Õ ·„œÌ— «·‰Ÿ«„ » €ÌÌ— «”„ «·„Õ·", 227, 9072, 2268, 425, 10, False, CLR_MUTED, "AllowAdminCompanyName", 0)
    Set c = AddLabel("lblStatus", " ", 227, 9752, 14855, 340, 10, True, CLR_MUTED, "", 0)
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
    s = s & "Private Sub btnActivation_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmActivation"", 10" & vbCrLf
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
    s = s & "Private Sub btnBrowseImages_Click()" & vbCrLf
    s = s & "    BrowseFolder Me, ""ImagesFolder""" & vbCrLf
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
    s = s & "    FitControls Me, 15309, 8675, -4422, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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
    Set c = AddCombo("cboCustomer", "", 5670, 3912, 4593, 454, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;4536")
    Set c = AddLabel("lblCustomer", "«·⁄„Ì·", 5670, 3600, 4593, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddCombo("cboSupplier", "", 10490, 3912, 4593, 454, "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", 2, "0;4536")
    Set c = AddLabel("lblSupplier", "«·„Ê—œ", 10490, 3600, 4593, 284, 9, False, CLR_MUTED, "cboSupplier", 0)
    Set c = AddCombo("cboProduct", "", 5670, 4734, 4593, 454, "SELECT ProductID, ProductName & ' (' & ProductCode & ')' AS Item FROM Products ORDER BY ProductName", 2, "0;4536")
    Set c = AddLabel("lblProduct", "«·„‰ Ã", 5670, 4422, 4593, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddCombo("cboCashBox", "", 10490, 4734, 4593, 454, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;4536")
    Set c = AddLabel("lblCashBox", "«·Œ“Ì‰… / «·’‰œÊﬁ", 10490, 4422, 4593, 284, 9, False, CLR_MUTED, "cboCashBox", 0)
    Set c = AddCombo("cboExpenseType", "", 5670, 5556, 4593, 454, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName", 2, "0;4536")
    Set c = AddLabel("lblExpenseType", "‰Ê⁄ «·„’—Ê›", 5670, 5244, 4593, 284, 9, False, CLR_MUTED, "cboExpenseType", 0)
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
    s = s & "    FitControls Me, 15309, 8675, -709, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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
    s = s & "    FitControls Me, 18994, 8675, 0, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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
    Set c = AddText("txtAmount", "", 227, 2608, 2495, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ *", 227, 2296, 2495, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboCurrency", "", 2835, 2608, 1474, 510, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ListWidth", "5.5cm"
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblCurrency", "«·⁄„·… *", 2835, 2296, 1474, 284, 9, False, CLR_MUTED, "cboCurrency", 0)
    Set c = AddText("txtRate", "", 4422, 2608, 1361, 510)
    SetCtlProp c, "Format", "0.0000"
    Set c = AddLabel("lblRate", "«·„⁄«„·", 4422, 2296, 1361, 284, 9, False, CLR_MUTED, "txtRate", 0)
    Set c = AddCombo("cboPaymentMethod", "", 5897, 2608, 2948, 510, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 5897, 2296, 2948, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
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
    s = s & "Private Sub cboCurrency_AfterUpdate()" & vbCrLf
    s = s & "    CurrencyPicked Me" & vbCrLf
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
    Set c = AddText("txtNotes", "", 12474, 4905, 3033, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 12474, 4593, 3033, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddCombo("cboCurrency", "", 15735, 4905, 1474, 454, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ListWidth", "5.5cm"
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblCurrency", "«·⁄„·… *", 15735, 4593, 1474, 284, 9, False, CLR_MUTED, "cboCurrency", 0)
    Set c = AddText("txtRate", "", 17322, 4905, 1445, 454)
    SetCtlProp c, "Format", "0.0000"
    Set c = AddLabel("lblRate", "«·„⁄«„·", 17322, 4593, 1445, 284, 9, False, CLR_MUTED, "txtRate", 0)
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
    s = s & "Private Sub cboCurrency_AfterUpdate()" & vbCrLf
    s = s & "    PurCurrencyPicked Me" & vbCrLf
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
    s = s & "    spec = spec & "";lblChargeVAT,16104,4082,2664,454,1000,0,0,0;txtNotes,12474,4905,3033,454,1000,0,0,0;lblNotes,12474,4593,3033,284,1000,0,0,0;cboCurrency,15735,4905,1474,454,1000,0,0,0;lblCurrency,15735,4593,1474,284,1000,0,0,0;txtRate,17322,4905,1445,454,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblRate,17322,4593,1445,284,1000,0,0,0;boxTotals,12474,5443,6294,2126,1000,0,0,0;lblCapSubTotal,12644,5500,3686,340,1000,0,0,0;lblSubTotal,16386,5500,2211,340,1000,0,0,0;lblCapDiscount,12644,5840,3686,340,1000,0,0,0;lblDiscount,16386,5840,2211,340,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCapTax,12644,6180,3686,340,1000,0,0,0;lblTax,16386,6180,2211,340,1000,0,0,0;lblCapTotal,12644,6606,2835,340,1000,0,0,0;lblTotal,15536,6520,3062,567,1000,0,0,0;lblItems,12644,7173,5954,340,1000,0,0,0;txtPaid,12474,7938,3033,510,1000,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblPaid,12474,7626,3033,284,1000,0,1000,0;lblRemaining,15735,7966,3033,454,1000,0,1000,0;btnLabels,12474,8448,2268,624,1000,0,1000,0;btnClose,17067,8448,1701,624,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9242, 0, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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
    Set c = AddText("txtAmount", "", 227, 2608, 2495, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ *", 227, 2296, 2495, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboCurrency", "", 2835, 2608, 1474, 510, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ListWidth", "5.5cm"
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblCurrency", "«·⁄„·… *", 2835, 2296, 1474, 284, 9, False, CLR_MUTED, "cboCurrency", 0)
    Set c = AddText("txtRate", "", 4422, 2608, 1361, 510)
    SetCtlProp c, "Format", "0.0000"
    Set c = AddLabel("lblRate", "«·„⁄«„·", 4422, 2296, 1361, 284, 9, False, CLR_MUTED, "txtRate", 0)
    Set c = AddCombo("cboPaymentMethod", "", 5897, 2608, 2948, 510, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "ÿ—Ìﬁ… «·œ›⁄", 5897, 2296, 2948, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
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
    s = s & "Private Sub cboCurrency_AfterUpdate()" & vbCrLf
    s = s & "    CurrencyPicked Me" & vbCrLf
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
    s = s & "    FitControls Me, 18994, 9355, -1757, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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
    s = s & "    FitControls Me, 18994, 9355, -2919, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
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

Private Sub BuildForm_frmUserScreenLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmUserScreenLines", "‘«‘«  «·„” Œœ„", "SELECT * FROM tmpUserScreens ORDER BY SortOrder", 11113, 397, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddCheck("CanOpen", "CanOpen", 254, 68)
    c.AfterUpdate = EP
    Set c = AddText("ScreenTitle", "ScreenTitle", 793, 0, 3175, 397)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("ModuleName", "ModuleName", 3996, 0, 1361, 397)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddCheck("CanAdd", "CanAdd", 5781, 68)
    c.AfterUpdate = EP
    Set c = AddCheck("CanEdit", "CanEdit", 6745, 68)
    c.AfterUpdate = EP
    Set c = AddCheck("CanDelete", "CanDelete", 7567, 68)
    c.AfterUpdate = EP
    Set c = AddText("ActionsNote", "ActionsNote", 8134, 0, 2977, 397)
    c.FontSize = 9
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    s = ""
    s = s & "Private Sub CanOpen_AfterUpdate()" & vbCrLf
    s = s & "    UserScreenLineChanged Me, ""CanOpen""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CanAdd_AfterUpdate()" & vbCrLf
    s = s & "    UserScreenLineChanged Me, ""CanAdd""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CanEdit_AfterUpdate()" & vbCrLf
    s = s & "    UserScreenLineChanged Me, ""CanEdit""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CanDelete_AfterUpdate()" & vbCrLf
    s = s & "    UserScreenLineChanged Me, ""CanDelete""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmUserScreenLines", s
    Exit Sub
EH:
    AbortForm "frmUserScreenLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmUserScreens()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmUserScreens", "’·«ÕÌ«  «·‘«‘« ", "", 11567, 9866, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 11567, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "’·«ÕÌ«  «·‘«‘« ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·‘«‘«  «· Ì Ì› ÕÂ« ﬂ· „” Œœ„° Ê«·≈÷«›… Ê«· ⁄œÌ· Ê«·Õ–› ›Ì ﬂ· ‘«‘…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboUser", "", 227, 1304, 4536, 454, "SELECT e.EmployeeID, e.EmployeeName & '  (' & e.Username & ')', r.RoleName FROM Employees AS e INNER JOIN Roles AS r ON e.RoleID = r.RoleID WHERE e.IsDeveloper = False ORDER BY e.EmployeeName", 3, "0;3402;1418")
    c.AfterUpdate = EP
    Set c = AddLabel("lblUser", "«·„” Œœ„", 227, 992, 4536, 284, 9, False, CLR_MUTED, "cboUser", 0)
    Set c = AddLabel("lblUserInfo", " ", 4933, 1332, 6407, 397, 9, False, CLR_MUTED, "", 0)
    Set c = AddCheck("chkCustom", "", 227, 1956)
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustom", "’·«ÕÌ«  ‘«‘«  Œ«’… »Â–« «·„” Œœ„ (»œ· ’·«ÕÌ«  œÊ—Â)", 595, 1871, 10745, 454, 10, True, CLR_TEXT, "chkCustom", 0)
    Set c = AddLabel("lblCol1", "› Õ", 255, 2438, 737, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·‘«‘…", 1020, 2438, 3175, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·ﬁ”„", 4223, 2438, 1361, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "≈÷«›… / Õ›Ÿ", 5612, 2438, 1077, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", " ⁄œÌ·", 6717, 2438, 794, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "Õ–›", 7539, 2438, 794, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "„« Ì‰ÿ»ﬁ ⁄·ÌÂ«", 8361, 2438, 2977, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subScreens", "frmUserScreenLines", 227, 2778, 11113, 5443)
    Set c = AddLabel("lblNote", " ", 227, 8278, 11113, 567, 9, True, CLR_WARNING, "", 0)
    Set c = AddButton("btnSaveScreens", "Õ›Ÿ", 227, 9015, 1588, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnFromRole", "„‰ ’·«ÕÌ«  «·œÊ—", 1928, 9015, 2155, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAll", "ﬂ· «·‘«‘« ", 4196, 9015, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNone", "≈·€«¡ «·ﬂ·", 6010, 9015, 1588, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 9866, 9015, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    UserScreensLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboUser_AfterUpdate()" & vbCrLf
    s = s & "    UserScreensPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub chkCustom_AfterUpdate()" & vbCrLf
    s = s & "    UserScreensCustomChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveScreens_Click()" & vbCrLf
    s = s & "    SaveUserScreens Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFromRole_Click()" & vbCrLf
    s = s & "    UserScreensFromRole Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAll_Click()" & vbCrLf
    s = s & "    UserScreensAll Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNone_Click()" & vbCrLf
    s = s & "    UserScreensAll Me, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmUserScreens", s
    Exit Sub
EH:
    AbortForm "frmUserScreens", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmActivation()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmActivation", " ›⁄Ì· «·»—‰«„Ã", "", 10773, 9979, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 10773, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE713), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", " ›⁄Ì· «·»—‰«„Ã", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "Ì⁄„· «·»—‰«„Ã ⁄·Ï «·√ÃÂ“… «·„›⁄¯·… ›ﬁÿ »ﬂÊœ „‰ «·„»—„Ã", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblState", " ", 227, 1049, 10319, 454, 13, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtMachineID", "", 227, 1928, 5103, 510)
    c.FontSize = 14
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    Set c = AddLabel("lblMachineID", "—ﬁ„ Â–« «·ÃÂ«“ (√—”·Â ··„»—„Ã)", 227, 1616, 5103, 284, 9, False, CLR_MUTED, "txtMachineID", 0)
    Set c = AddButton("btnCopyID", "‰”Œ «·—ﬁ„", 5443, 1928, 1701, 510, "secondary")
    c.OnClick = EP
    Set c = AddText("txtCode", "", 227, 2835, 5103, 510)
    c.FontSize = 14
    Set c = AddLabel("lblCode", "ﬂÊœ «· ›⁄Ì·", 227, 2523, 5103, 284, 9, False, CLR_MUTED, "txtCode", 0)
    Set c = AddButton("btnActivate", " ›⁄Ì· Â–« «·ÃÂ«“", 5443, 2835, 2495, 510, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblListCap", "«·√ÃÂ“… «·„›⁄¯·…", 227, 3572, 6804, 340, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstMachines", 227, 3941, 10319, 2041, 4, "0;2835;3118;2268", True)
    c.RowSource = Tr("SELECT ActivationID, ComputerName AS [«·ÃÂ«“], MachineID AS [—ﬁ„ «·ÃÂ«“], ActivatedAt AS [ «—ÌŒ «· ›⁄Ì·] FROM Activations ORDER BY ActivatedAt")
    Set c = AddButton("btnRemove", "≈·€«¡  ›⁄Ì· «·ÃÂ«“ «·„Õœœ", 227, 6095, 3175, 510, "danger")
    c.OnClick = EP
    Set c = AddRect("boxDeveloper", 227, 6776, 10319, 2126, CLR_SURFACE)
    SetCtlProp c, "BorderColor", CLR_BORDER
    Set c = AddLabel("lblDevCap", "··„»—„Ã:  Ê·Ìœ ﬂÊœ  ›⁄Ì· ·ÃÂ«“ ⁄„Ì·", 397, 6861, 9979, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtForMachine", "", 397, 7569, 4196, 482)
    c.FontSize = 12
    Set c = AddLabel("lblForMachine", "—ﬁ„ ÃÂ«“ «·⁄„Ì·", 397, 7257, 4196, 284, 9, False, CLR_MUTED, "txtForMachine", 0)
    Set c = AddButton("btnGenerate", " Ê·Ìœ «·ﬂÊœ", 4706, 7569, 1701, 482, "primary")
    c.OnClick = EP
    Set c = AddText("txtGenerated", "", 6520, 7569, 3856, 482)
    c.FontSize = 12
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    Set c = AddLabel("lblGenerated", "ﬂÊœ «· ›⁄Ì·", 6520, 7257, 3856, 284, 9, False, CLR_MUTED, "txtGenerated", 0)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 9072, 9185, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ActivationLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCopyID_Click()" & vbCrLf
    s = s & "    CopyMachineID Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnActivate_Click()" & vbCrLf
    s = s & "    ActivateThisMachine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    RemoveSelectedActivation Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnGenerate_Click()" & vbCrLf
    s = s & "    GenerateActivationCode Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmActivation", s
    Exit Sub
EH:
    AbortForm "frmActivation", Err.Number, Err.Description
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
    s = s & "    FitControls Me, 11340, 8278, -2834, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBarcodeLabels", s
    Exit Sub
EH:
    AbortForm "frmBarcodeLabels", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmTouchLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmTouchLines", "√”ÿ— «·ÿ·»", "SELECT * FROM tmpPOSLines ORDER BY LineNo", 6237, 652, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddButton("btnRemove", "Sym(code='ChrW(&HE74D)')", 28, 17, 510, 618, "danger")
    SetCtlProp c, "FontName", ICON_FONT
    c.OnClick = EP
    Set c = AddText("LineTotal", "LineTotal", 566, 0, 1191, 652)
    c.FontSize = 12
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddButton("btnPlus", "+", 1785, 17, 567, 618, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddText("Quantity", "Quantity", 2380, 0, 624, 652)
    c.FontSize = 16
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddButton("btnMinus", "-", 3032, 17, 567, 618, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddText("ProductName", "ProductName", 3627, 0, 2495, 652)
    c.FontSize = 12
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    s = ""
    s = s & "Private Sub btnPlus_Click()" & vbCrLf
    s = s & "    TouchQtyStep Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnMinus_Click()" & vbCrLf
    s = s & "    TouchQtyStep Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    RemoveCurrentLine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmTouchLines", s
    Exit Sub
EH:
    AbortForm "frmTouchLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmTouchPOS()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmTouchPOS", "‰ﬁÿ… »Ì⁄ «·„ÿ⁄„", "", 18994, 9299, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "‰ﬁÿ… »Ì⁄ «·„ÿ⁄„", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «·›∆… À„ «·’‰›  |  + Ê - · ⁄œÌ· «·ﬂ„Ì…  |  ‰ﬁœÌ √Ê »ÿ«ﬁ… ··œ›⁄", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17463, 170, 1361, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReprint", "≈⁄«œ… ÿ»«⁄… ¬Œ— ›« Ê—…", 14685, 170, 2665, 510, "secondary")
    c.OnClick = EP
    Set c = AddText("txtBarcode", "", 9072, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtQty", "", 9270, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtInvoiceDiscount", "", 9468, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtTendered", "", 9666, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtNotes", "", 9864, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtOrderType", "", 10062, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddCombo("cboPaymentType", "", 10260, 57, 170, 170, "CASH;‰ﬁœÌ;CREDIT;¬Ã·", 2, "0;2835")
    SetCtlProp c, "Visible", False
    Set c = AddCombo("cboPaymentMethod", "", 10458, 57, 170, 170, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;2835")
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblChange", " ", 10656, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblLastInvoice", " ", 10854, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblCustomerInfo", " ", 11052, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblOrderTitle", " ", 170, 936, 6237, 340, 13, True, CLR_PRIMARY, "", 3)
    Set c = AddButton("btnTypeDelivery", " Ê’Ì·", 170, 1304, 2041, 737, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnTypeTakeaway", "”›—Ì", 2268, 1304, 2041, 737, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnTypeDineIn", "œ«Œ·Ì", 4366, 1304, 2041, 737, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddCombo("cboTable", "", 3572, 2410, 2835, 567, "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40", 1, "1701")
    c.FontSize = 16
    SetCtlProp c, "LimitToList", False
    c.AfterUpdate = EP
    Set c = AddLabel("lblTable", "—ﬁ„ «·ÿ«Ê·…", 3572, 2098, 2835, 284, 9, False, CLR_MUTED, "cboTable", 0)
    Set c = AddLabel("lblTakeaway", "ÿ·» ”›—Ì: Ìı€·Û¯› ÊÌı”·Û¯„ ··⁄„Ì·", 170, 2410, 6237, 567, 14, True, CLR_MUTED, "", 2)
    Set c = AddCombo("cboCustomer", "", 3345, 2410, 3062, 482, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5670")
    c.FontSize = 12
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustomer", "«·⁄„Ì· («Œ Ì«—Ì)", 3345, 2098, 3062, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddText("txtDeliveryPhone", "", 170, 2410, 3062, 482)
    c.FontSize = 14
    Set c = AddLabel("lblDeliveryPhone", "ÃÊ«· «· Ê’Ì· *", 170, 2098, 3062, 284, 9, False, CLR_MUTED, "txtDeliveryPhone", 0)
    Set c = AddText("txtDeliveryAddress", "", 170, 3204, 6237, 454)
    c.FontSize = 12
    Set c = AddLabel("lblDeliveryAddress", "⁄‰Ê«‰ «· Ê’Ì· *", 170, 2949, 6237, 227, 9, False, CLR_MUTED, "txtDeliveryAddress", 0)
    Set c = AddLabel("lblCol2", "«·≈Ã„«·Ì", 736, 3742, 1191, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·ﬂ„Ì…", 2550, 3742, 624, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·’‰›", 3797, 3742, 2495, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmTouchLines", 170, 4054, 6237, 2552)
    Set c = AddRect("boxTotals", 170, 6691, 6237, 1134, CLR_SURFACE)
    Set c = AddLabel("lblCapSubTotal", "ﬁ»· «·Œ’„ Ê«·÷—Ì»…", 4309, 6748, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblSubTotal", "0.00", 2835, 6748, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapDiscount", "«·Œ’„", 4309, 7032, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblDiscount", "0.00", 2835, 7032, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 4309, 7316, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblTax", "0.00", 2835, 7316, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblItems", " ", 2835, 7598, 3515, 215, 9, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblCapTotal", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 283, 6719, 2438, 255, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTotal", "0.00", 283, 7003, 2438, 737, 26, True, CLR_ACCENT, "", 2)
    Set c = AddLabel("lblStatus", " ", 170, 7881, 6237, 284, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnPayCash", "‰ﬁœÌ", 3969, 8222, 2438, 907, "primary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnPayCard", "„œÏ / »ÿ«ﬁ…", 1531, 8222, 2381, 907, "nav")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnCancelOrder", "≈·€«¡", 170, 8222, 1304, 907, "danger")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddLabel("lblCategoryTitle", " ", 6577, 936, 9356, 340, 14, True, CLR_TEXT, "", 3)
    Set c = AddRect("boxProd1", 13699, 1304, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip1", 13699, 1304, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd1", 13869, 1474, 1892, 851)
    Set c = AddLabel("lblProd1", " ", 13756, 2353, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice1", " ", 13756, 2636, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd1", " ", 13699, 1304, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd2", 11325, 1304, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip2", 11325, 1304, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd2", 11495, 1474, 1892, 851)
    Set c = AddLabel("lblProd2", " ", 11382, 2353, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice2", " ", 11382, 2636, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd2", " ", 11325, 1304, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd3", 8951, 1304, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip3", 8951, 1304, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd3", 9121, 1474, 1892, 851)
    Set c = AddLabel("lblProd3", " ", 9008, 2353, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice3", " ", 9008, 2636, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd3", " ", 8951, 1304, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd4", 6577, 1304, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip4", 6577, 1304, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd4", 6747, 1474, 1892, 851)
    Set c = AddLabel("lblProd4", " ", 6634, 2353, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice4", " ", 6634, 2636, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd4", " ", 6577, 1304, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd5", 13699, 3033, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip5", 13699, 3033, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd5", 13869, 3203, 1892, 851)
    Set c = AddLabel("lblProd5", " ", 13756, 4082, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice5", " ", 13756, 4365, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd5", " ", 13699, 3033, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd6", 11325, 3033, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip6", 11325, 3033, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd6", 11495, 3203, 1892, 851)
    Set c = AddLabel("lblProd6", " ", 11382, 4082, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice6", " ", 11382, 4365, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd6", " ", 11325, 3033, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd7", 8951, 3033, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip7", 8951, 3033, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd7", 9121, 3203, 1892, 851)
    Set c = AddLabel("lblProd7", " ", 9008, 4082, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice7", " ", 9008, 4365, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd7", " ", 8951, 3033, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd8", 6577, 3033, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip8", 6577, 3033, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd8", 6747, 3203, 1892, 851)
    Set c = AddLabel("lblProd8", " ", 6634, 4082, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice8", " ", 6634, 4365, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd8", " ", 6577, 3033, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd9", 13699, 4762, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip9", 13699, 4762, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd9", 13869, 4932, 1892, 851)
    Set c = AddLabel("lblProd9", " ", 13756, 5811, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice9", " ", 13756, 6094, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd9", " ", 13699, 4762, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd10", 11325, 4762, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip10", 11325, 4762, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd10", 11495, 4932, 1892, 851)
    Set c = AddLabel("lblProd10", " ", 11382, 5811, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice10", " ", 11382, 6094, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd10", " ", 11325, 4762, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd11", 8951, 4762, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip11", 8951, 4762, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd11", 9121, 4932, 1892, 851)
    Set c = AddLabel("lblProd11", " ", 9008, 5811, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice11", " ", 9008, 6094, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd11", " ", 8951, 4762, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd12", 6577, 4762, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip12", 6577, 4762, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd12", 6747, 4932, 1892, 851)
    Set c = AddLabel("lblProd12", " ", 6634, 5811, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice12", " ", 6634, 6094, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd12", " ", 6577, 4762, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd13", 13699, 6491, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip13", 13699, 6491, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd13", 13869, 6661, 1892, 851)
    Set c = AddLabel("lblProd13", " ", 13756, 7540, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice13", " ", 13756, 7823, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd13", " ", 13699, 6491, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd14", 11325, 6491, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip14", 11325, 6491, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd14", 11495, 6661, 1892, 851)
    Set c = AddLabel("lblProd14", " ", 11382, 7540, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice14", " ", 11382, 7823, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd14", " ", 11325, 6491, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd15", 8951, 6491, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip15", 8951, 6491, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd15", 9121, 6661, 1892, 851)
    Set c = AddLabel("lblProd15", " ", 9008, 7540, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice15", " ", 9008, 7823, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd15", " ", 8951, 6491, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd16", 6577, 6491, 2232, 1616, CLR_SURFACE)
    Set c = AddRect("boxProdStrip16", 6577, 6491, 2232, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd16", 6747, 6661, 1892, 851)
    Set c = AddLabel("lblProd16", " ", 6634, 7540, 2119, 284, 11, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice16", " ", 6634, 7823, 2119, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd16", " ", 6577, 6491, 2232, 1616, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddButton("btnProdNext", "«· «·Ì", 6577, 8222, 1701, 907, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblProdPage", " ", 8391, 8477, 5727, 397, 12, False, CLR_MUTED, "", 2)
    Set c = AddButton("btnProdPrev", "«·”«»ﬁ", 14232, 8222, 1701, 907, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblCatHeader", "«·›∆« ", 16103, 936, 2722, 340, 14, True, CLR_TEXT, "", 3)
    Set c = AddRect("boxCat1", 16103, 1304, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat1", 18003, 1372, 737, 743)
    Set c = AddLabel("lblCat1", " ", 16188, 1531, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat1", " ", 16103, 1304, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat2", 16103, 2268, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat2", 18003, 2336, 737, 743)
    Set c = AddLabel("lblCat2", " ", 16188, 2495, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat2", " ", 16103, 2268, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat3", 16103, 3232, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat3", 18003, 3300, 737, 743)
    Set c = AddLabel("lblCat3", " ", 16188, 3459, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat3", " ", 16103, 3232, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat4", 16103, 4196, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat4", 18003, 4264, 737, 743)
    Set c = AddLabel("lblCat4", " ", 16188, 4423, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat4", " ", 16103, 4196, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat5", 16103, 5160, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat5", 18003, 5228, 737, 743)
    Set c = AddLabel("lblCat5", " ", 16188, 5387, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat5", " ", 16103, 5160, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat6", 16103, 6124, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat6", 18003, 6192, 737, 743)
    Set c = AddLabel("lblCat6", " ", 16188, 6351, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat6", " ", 16103, 6124, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat7", 16103, 7088, 2722, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat7", 18003, 7156, 737, 743)
    Set c = AddLabel("lblCat7", " ", 16188, 7315, 1730, 425, 13, False, CLR_SURFACE, "", 3)
    Set c = AddButton("btnCat7", " ", 16103, 7088, 2722, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddButton("btnCatDown", "«· «·Ì…", 16103, 8222, 1304, 907, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCatUp", "«·”«»ﬁ…", 17520, 8222, 1304, 907, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnUnload = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    TouchLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not POSUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboTable_AfterUpdate()" & vbCrLf
    s = s & "    TableChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCustomer_AfterUpdate()" & vbCrLf
    s = s & "    DeliveryCustomerChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReprint_Click()" & vbCrLf
    s = s & "    ReprintLast Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTypeDelivery_Click()" & vbCrLf
    s = s & "    SetOrderType Me, ""DELIVERY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTypeTakeaway_Click()" & vbCrLf
    s = s & "    SetOrderType Me, ""TAKEAWAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTypeDineIn_Click()" & vbCrLf
    s = s & "    SetOrderType Me, ""DINE_IN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayCash_Click()" & vbCrLf
    s = s & "    TouchPayCash Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayCard_Click()" & vbCrLf
    s = s & "    TouchPayCard Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCancelOrder_Click()" & vbCrLf
    s = s & "    NewSale Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd1_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd2_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd3_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd4_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd5_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd6_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd7_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd8_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 8" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd9_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 9" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd10_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd11_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 11" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd12_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 12" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd13_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 13" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd14_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 14" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd15_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 15" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd16_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 16" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProdNext_Click()" & vbCrLf
    s = s & "    ProductPage Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProdPrev_Click()" & vbCrLf
    s = s & "    ProductPage Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat1_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat2_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat3_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat4_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat5_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat6_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat7_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCatDown_Click()" & vbCrLf
    s = s & "    CategoryPage Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCatUp_Click()" & vbCrLf
    s = s & "    CategoryPage Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;btnClose,17463,170,1361,510,1000,0,0,0;btnReprint,14685,170,2665,510,1000,0,0,0;subLines,170,4054,6237,2552,0,0,0,1000;boxTotals,170,6691,6237,1134,0,0,1000,0;lblCapSubTotal,4309,6748,2041,255,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblSubTotal,2835,6748,1418,255,0,0,1000,0;lblCapDiscount,4309,7032,2041,255,0,0,1000,0;lblDiscount,2835,7032,1418,255,0,0,1000,0;lblCapTax,4309,7316,2041,255,0,0,1000,0;lblTax,2835,7316,1418,255,0,0,1000,0;lblItems,2835,7598,3515,215,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCapTotal,283,6719,2438,255,0,0,1000,0;lblTotal,283,7003,2438,737,0,0,1000,0;lblStatus,170,7881,6237,284,0,0,1000,0;btnPayCash,3969,8222,2438,907,0,0,1000,0;btnPayCard,1531,8222,2381,907,0,0,1000,0;btnCancelOrder,170,8222,1304,907,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCategoryTitle,6577,936,9356,340,0,1000,0,0;boxProd1,13699,1304,2232,1616,750,250,0,250;boxProdStrip1,13699,1304,2232,102,750,250,0,250;imgProd1,13869,1474,1892,851,750,250,0,250;lblProd1,13756,2353,2119,284,750,250,0,250;lblPrice1,13756,2636,2119,255,750,250,0,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd1,13699,1304,2232,1616,750,250,0,250;boxProd2,11325,1304,2232,1616,500,250,0,250;boxProdStrip2,11325,1304,2232,102,500,250,0,250;imgProd2,11495,1474,1892,851,500,250,0,250;lblProd2,11382,2353,2119,284,500,250,0,250;lblPrice2,11382,2636,2119,255,500,250,0,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd2,11325,1304,2232,1616,500,250,0,250;boxProd3,8951,1304,2232,1616,250,250,0,250;boxProdStrip3,8951,1304,2232,102,250,250,0,250;imgProd3,9121,1474,1892,851,250,250,0,250;lblProd3,9008,2353,2119,284,250,250,0,250;lblPrice3,9008,2636,2119,255,250,250,0,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd3,8951,1304,2232,1616,250,250,0,250;boxProd4,6577,1304,2232,1616,0,250,0,250;boxProdStrip4,6577,1304,2232,102,0,250,0,250;imgProd4,6747,1474,1892,851,0,250,0,250;lblProd4,6634,2353,2119,284,0,250,0,250;lblPrice4,6634,2636,2119,255,0,250,0,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd4,6577,1304,2232,1616,0,250,0,250;boxProd5,13699,3033,2232,1616,750,250,250,250;boxProdStrip5,13699,3033,2232,102,750,250,250,250;imgProd5,13869,3203,1892,851,750,250,250,250;lblProd5,13756,4082,2119,284,750,250,250,250;lblPrice5,13756,4365,2119,255,750,250,250,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd5,13699,3033,2232,1616,750,250,250,250;boxProd6,11325,3033,2232,1616,500,250,250,250;boxProdStrip6,11325,3033,2232,102,500,250,250,250;imgProd6,11495,3203,1892,851,500,250,250,250;lblProd6,11382,4082,2119,284,500,250,250,250;lblPrice6,11382,4365,2119,255,500,250,250,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd6,11325,3033,2232,1616,500,250,250,250;boxProd7,8951,3033,2232,1616,250,250,250,250;boxProdStrip7,8951,3033,2232,102,250,250,250,250;imgProd7,9121,3203,1892,851,250,250,250,250;lblProd7,9008,4082,2119,284,250,250,250,250;lblPrice7,9008,4365,2119,255,250,250,250,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd7,8951,3033,2232,1616,250,250,250,250;boxProd8,6577,3033,2232,1616,0,250,250,250;boxProdStrip8,6577,3033,2232,102,0,250,250,250;imgProd8,6747,3203,1892,851,0,250,250,250;lblProd8,6634,4082,2119,284,0,250,250,250;lblPrice8,6634,4365,2119,255,0,250,250,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd8,6577,3033,2232,1616,0,250,250,250;boxProd9,13699,4762,2232,1616,750,250,500,250;boxProdStrip9,13699,4762,2232,102,750,250,500,250;imgProd9,13869,4932,1892,851,750,250,500,250;lblProd9,13756,5811,2119,284,750,250,500,250;lblPrice9,13756,6094,2119,255,750,250,500,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd9,13699,4762,2232,1616,750,250,500,250;boxProd10,11325,4762,2232,1616,500,250,500,250;boxProdStrip10,11325,4762,2232,102,500,250,500,250;imgProd10,11495,4932,1892,851,500,250,500,250;lblProd10,11382,5811,2119,284,500,250,500,250;lblPrice10,11382,6094,2119,255,500,250,500,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd10,11325,4762,2232,1616,500,250,500,250;boxProd11,8951,4762,2232,1616,250,250,500,250;boxProdStrip11,8951,4762,2232,102,250,250,500,250;imgProd11,9121,4932,1892,851,250,250,500,250;lblProd11,9008,5811,2119,284,250,250,500,250;lblPrice11,9008,6094,2119,255,250,250,500,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd11,8951,4762,2232,1616,250,250,500,250;boxProd12,6577,4762,2232,1616,0,250,500,250;boxProdStrip12,6577,4762,2232,102,0,250,500,250;imgProd12,6747,4932,1892,851,0,250,500,250;lblProd12,6634,5811,2119,284,0,250,500,250;lblPrice12,6634,6094,2119,255,0,250,500,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd12,6577,4762,2232,1616,0,250,500,250;boxProd13,13699,6491,2232,1616,750,250,750,250;boxProdStrip13,13699,6491,2232,102,750,250,750,250;imgProd13,13869,6661,1892,851,750,250,750,250;lblProd13,13756,7540,2119,284,750,250,750,250;lblPrice13,13756,7823,2119,255,750,250,750,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd13,13699,6491,2232,1616,750,250,750,250;boxProd14,11325,6491,2232,1616,500,250,750,250;boxProdStrip14,11325,6491,2232,102,500,250,750,250;imgProd14,11495,6661,1892,851,500,250,750,250;lblProd14,11382,7540,2119,284,500,250,750,250;lblPrice14,11382,7823,2119,255,500,250,750,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd14,11325,6491,2232,1616,500,250,750,250;boxProd15,8951,6491,2232,1616,250,250,750,250;boxProdStrip15,8951,6491,2232,102,250,250,750,250;imgProd15,9121,6661,1892,851,250,250,750,250;lblProd15,9008,7540,2119,284,250,250,750,250;lblPrice15,9008,7823,2119,255,250,250,750,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd15,8951,6491,2232,1616,250,250,750,250;boxProd16,6577,6491,2232,1616,0,250,750,250;boxProdStrip16,6577,6491,2232,102,0,250,750,250;imgProd16,6747,6661,1892,851,0,250,750,250;lblProd16,6634,7540,2119,284,0,250,750,250;lblPrice16,6634,7823,2119,255,0,250,750,250""" & vbCrLf
    s = s & "    spec = spec & "";btnProd16,6577,6491,2232,1616,0,250,750,250;btnProdNext,6577,8222,1701,907,0,0,1000,0;lblProdPage,8391,8477,5727,397,0,1000,1000,0;btnProdPrev,14232,8222,1701,907,1000,0,1000,0;lblCatHeader,16103,936,2722,340,1000,0,0,0;boxCat1,16103,1304,2722,879,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";imgCat1,18003,1372,737,743,1000,0,0,0;lblCat1,16188,1531,1730,425,1000,0,0,0;btnCat1,16103,1304,2722,879,1000,0,0,0;boxCat2,16103,2268,2722,879,1000,0,0,0;imgCat2,18003,2336,737,743,1000,0,0,0;lblCat2,16188,2495,1730,425,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";btnCat2,16103,2268,2722,879,1000,0,0,0;boxCat3,16103,3232,2722,879,1000,0,0,0;imgCat3,18003,3300,737,743,1000,0,0,0;lblCat3,16188,3459,1730,425,1000,0,0,0;btnCat3,16103,3232,2722,879,1000,0,0,0;boxCat4,16103,4196,2722,879,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";imgCat4,18003,4264,737,743,1000,0,0,0;lblCat4,16188,4423,1730,425,1000,0,0,0;btnCat4,16103,4196,2722,879,1000,0,0,0;boxCat5,16103,5160,2722,879,1000,0,0,0;imgCat5,18003,5228,737,743,1000,0,0,0;lblCat5,16188,5387,1730,425,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";btnCat5,16103,5160,2722,879,1000,0,0,0;boxCat6,16103,6124,2722,879,1000,0,0,0;imgCat6,18003,6192,737,743,1000,0,0,0;lblCat6,16188,6351,1730,425,1000,0,0,0;btnCat6,16103,6124,2722,879,1000,0,0,0;boxCat7,16103,7088,2722,879,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";imgCat7,18003,7156,737,743,1000,0,0,0;lblCat7,16188,7315,1730,425,1000,0,0,0;btnCat7,16103,7088,2722,879,1000,0,0,0;btnCatDown,16103,8222,1304,907,1000,0,1000,0;btnCatUp,17520,8222,1304,907,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9299, 0, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmTouchPOS", s
    Exit Sub
EH:
    AbortForm "frmTouchPOS", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmTouchPay()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmTouchPay", "«·œ›⁄ ‰ﬁœ«", "", 8845, 8165, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 8845, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·œ›⁄ ‰ﬁœ«", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«ﬂ » «·„»·€ «·„” ·„ √Ê «Œ — „»·€« Ã«Â“«", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtPayInput", "", 8505, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblCapPayTotal", "«·≈Ã„«·Ì", 4423, 1134, 4195, 340, 13, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblPayTotal", "0.00", 227, 1021, 4082, 567, 20, True, CLR_TEXT, "", 1)
    Set c = AddLabel("lblCapPayAmount", "«·„»·€ «·„” ·„", 4423, 1757, 4195, 340, 13, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblPayAmount", "0.00", 227, 1644, 4082, 567, 24, True, CLR_ACCENT, "", 1)
    Set c = AddLabel("lblPayChange", " ", 227, 2325, 8391, 340, 13, True, CLR_SUCCESS, "", 2)
    Set c = AddButton("btnQuick1", "500", 227, 2778, 1610, 680, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnQuick2", "200", 1922, 2778, 1610, 680, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnQuick3", "100", 3617, 2778, 1610, 680, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnQuick4", "50", 5312, 2778, 1610, 680, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnQuick5", "»«·÷»ÿ", 7007, 2778, 1610, 680, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnKey9", "9", 227, 3600, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey8", "8", 3052, 3600, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey7", "7", 5877, 3600, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey6", "6", 227, 4450, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey5", "5", 3052, 4450, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey4", "4", 5877, 4450, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey3", "3", 227, 5300, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey2", "2", 3052, 5300, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey1", "1", 5877, 5300, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKeyBack", "Õ–›", 227, 6150, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKey0", "0", 3052, 6150, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnKeyDot", ".", 5877, 6150, 2740, 765, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddButton("btnPayConfirm", " √ﬂÌœ «·œ›⁄", 5273, 7059, 3345, 850, "primary")
    c.FontSize = 15
    c.OnClick = EP
    Set c = AddButton("btnPayCancel", "≈·€«¡", 1701, 7059, 1474, 850, "secondary")
    c.FontSize = 15
    c.OnClick = EP
    Set c = AddButton("btnPayClear", "„”Õ", 227, 7059, 1361, 850, "danger")
    c.FontSize = 15
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PayLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQuick1_Click()" & vbCrLf
    s = s & "    PayQuick Me, 500" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQuick2_Click()" & vbCrLf
    s = s & "    PayQuick Me, 200" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQuick3_Click()" & vbCrLf
    s = s & "    PayQuick Me, 100" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQuick4_Click()" & vbCrLf
    s = s & "    PayQuick Me, 50" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQuick5_Click()" & vbCrLf
    s = s & "    PayQuick Me, 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey9_Click()" & vbCrLf
    s = s & "    PayKey Me, ""9""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey8_Click()" & vbCrLf
    s = s & "    PayKey Me, ""8""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey7_Click()" & vbCrLf
    s = s & "    PayKey Me, ""7""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey6_Click()" & vbCrLf
    s = s & "    PayKey Me, ""6""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey5_Click()" & vbCrLf
    s = s & "    PayKey Me, ""5""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey4_Click()" & vbCrLf
    s = s & "    PayKey Me, ""4""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey3_Click()" & vbCrLf
    s = s & "    PayKey Me, ""3""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey2_Click()" & vbCrLf
    s = s & "    PayKey Me, ""2""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey1_Click()" & vbCrLf
    s = s & "    PayKey Me, ""1""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKeyBack_Click()" & vbCrLf
    s = s & "    PayKey Me, ""<""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKey0_Click()" & vbCrLf
    s = s & "    PayKey Me, ""0""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnKeyDot_Click()" & vbCrLf
    s = s & "    PayKey Me, "".""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayConfirm_Click()" & vbCrLf
    s = s & "    PayConfirm Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayCancel_Click()" & vbCrLf
    s = s & "    PayCancel Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayClear_Click()" & vbCrLf
    s = s & "    PayKey Me, ""C""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmTouchPay", s
    Exit Sub
EH:
    AbortForm "frmTouchPay", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCafePOS()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCafePOS", "‰ﬁÿ… »Ì⁄ «·ﬂ«›ÌÂ", "", 18994, 9299, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "‰ﬁÿ… »Ì⁄ «·ﬂ«›ÌÂ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «·„‘—Ê»: «·ÕÃ„ Ê«·≈÷«›«   ŸÂ—  ·ﬁ«∆Ì«  |  ‰ﬁœÌ √Ê »ÿ«ﬁ… ··œ›⁄", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 17463, 170, 1361, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReprint", "≈⁄«œ… ÿ»«⁄… ¬Œ— ›« Ê—…", 14685, 170, 2665, 510, "secondary")
    c.OnClick = EP
    Set c = AddText("txtBarcode", "", 9072, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtQty", "", 9270, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtInvoiceDiscount", "", 9468, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtTendered", "", 9666, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtNotes", "", 9864, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtOrderType", "", 10062, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddCombo("cboPaymentType", "", 10260, 57, 170, 170, "CASH;‰ﬁœÌ;CREDIT;¬Ã·", 2, "0;2835")
    SetCtlProp c, "Visible", False
    Set c = AddCombo("cboPaymentMethod", "", 10458, 57, 170, 170, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;2835")
    SetCtlProp c, "Visible", False
    Set c = AddCombo("cboCustomer", "", 10656, 57, 170, 170, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5670")
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblChange", " ", 10854, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblLastInvoice", " ", 11052, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblCustomerInfo", " ", 11250, 57, 170, 170, 10, False, CLR_TEXT, "", 0)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblOrderTitle", " ", 170, 936, 6237, 340, 13, True, CLR_PRIMARY, "", 3)
    Set c = AddButton("btnTypeTakeaway", "”›—Ì", 170, 1304, 3090, 737, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnTypeDineIn", "„Õ·Ì", 3317, 1304, 3090, 737, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddText("txtOrderName", "", 170, 2410, 6237, 567)
    c.FontSize = 16
    c.AfterUpdate = EP
    Set c = AddLabel("lblOrderName", "«”„ «·⁄„Ì· ⁄·Ï «·ﬂÊ» («Œ Ì«—Ì)", 170, 2098, 6237, 284, 9, False, CLR_MUTED, "txtOrderName", 0)
    Set c = AddLabel("lblCol2", "«·≈Ã„«·Ì", 736, 3090, 1191, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·ﬂ„Ì…", 2550, 3090, 624, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·’‰›", 3797, 3090, 2495, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmTouchLines", 170, 3402, 6237, 3204)
    Set c = AddRect("boxTotals", 170, 6691, 6237, 1134, CLR_SURFACE)
    Set c = AddLabel("lblCapSubTotal", "ﬁ»· «·Œ’„ Ê«·÷—Ì»…", 4309, 6748, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblSubTotal", "0.00", 2835, 6748, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapDiscount", "«·Œ’„", 4309, 7032, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblDiscount", "0.00", 2835, 7032, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 4309, 7316, 2041, 255, 10, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblTax", "0.00", 2835, 7316, 1418, 255, 11, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblItems", " ", 2835, 7598, 3515, 215, 9, False, CLR_MUTED, "", 3)
    Set c = AddLabel("lblCapTotal", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 283, 6719, 2438, 255, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTotal", "0.00", 283, 7003, 2438, 737, 26, True, CLR_ACCENT, "", 2)
    Set c = AddLabel("lblStatus", " ", 170, 7882, 6237, 284, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnPayCash", "‰ﬁœÌ", 3969, 8222, 2438, 907, "primary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnPayCard", "„œÏ / »ÿ«ﬁ…", 1531, 8222, 2381, 907, "nav")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnCancelOrder", "≈·€«¡", 170, 8222, 1304, 907, "danger")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnCatDown", "«· «·Ì…", 6577, 936, 794, 879, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCatUp", "«·”«»ﬁ…", 18030, 936, 794, 879, "secondary")
    c.OnClick = EP
    Set c = AddRect("boxCat1", 15920, 936, 1996, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat1", 17321, 1049, 510, 652)
    Set c = AddLabel("lblCat1", " ", 15977, 1191, 1316, 397, 12, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnCat1", " ", 15920, 936, 1996, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat2", 13811, 936, 1996, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat2", 15212, 1049, 510, 652)
    Set c = AddLabel("lblCat2", " ", 13868, 1191, 1316, 397, 12, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnCat2", " ", 13811, 936, 1996, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat3", 11702, 936, 1996, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat3", 13103, 1049, 510, 652)
    Set c = AddLabel("lblCat3", " ", 11759, 1191, 1316, 397, 12, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnCat3", " ", 11702, 936, 1996, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat4", 9593, 936, 1996, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat4", 10994, 1049, 510, 652)
    Set c = AddLabel("lblCat4", " ", 9650, 1191, 1316, 397, 12, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnCat4", " ", 9593, 936, 1996, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxCat5", 7484, 936, 1996, 879, CLR_PRIMARY)
    Set c = AddImage("imgCat5", 8885, 1049, 510, 652)
    Set c = AddLabel("lblCat5", " ", 7541, 1191, 1316, 397, 12, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnCat5", " ", 7484, 936, 1996, 879, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddLabel("lblCategoryTitle", " ", 6577, 1899, 12247, 340, 14, True, CLR_TEXT, "", 3)
    Set c = AddRect("boxProd1", 16485, 2296, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip1", 16485, 2296, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd1", 16655, 2466, 1995, 1097)
    Set c = AddLabel("lblProd1", " ", 16542, 3591, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice1", " ", 16542, 3874, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd1", " ", 16485, 2296, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd2", 14008, 2296, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip2", 14008, 2296, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd2", 14178, 2466, 1995, 1097)
    Set c = AddLabel("lblProd2", " ", 14065, 3591, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice2", " ", 14065, 3874, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd2", " ", 14008, 2296, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd3", 11531, 2296, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip3", 11531, 2296, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd3", 11701, 2466, 1995, 1097)
    Set c = AddLabel("lblProd3", " ", 11588, 3591, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice3", " ", 11588, 3874, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd3", " ", 11531, 2296, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd4", 9054, 2296, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip4", 9054, 2296, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd4", 9224, 2466, 1995, 1097)
    Set c = AddLabel("lblProd4", " ", 9111, 3591, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice4", " ", 9111, 3874, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd4", " ", 9054, 2296, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd5", 6577, 2296, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip5", 6577, 2296, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd5", 6747, 2466, 1995, 1097)
    Set c = AddLabel("lblProd5", " ", 6634, 3591, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice5", " ", 6634, 3874, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd5", " ", 6577, 2296, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd6", 16485, 4271, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip6", 16485, 4271, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd6", 16655, 4441, 1995, 1097)
    Set c = AddLabel("lblProd6", " ", 16542, 5566, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice6", " ", 16542, 5849, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd6", " ", 16485, 4271, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd7", 14008, 4271, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip7", 14008, 4271, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd7", 14178, 4441, 1995, 1097)
    Set c = AddLabel("lblProd7", " ", 14065, 5566, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice7", " ", 14065, 5849, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd7", " ", 14008, 4271, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd8", 11531, 4271, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip8", 11531, 4271, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd8", 11701, 4441, 1995, 1097)
    Set c = AddLabel("lblProd8", " ", 11588, 5566, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice8", " ", 11588, 5849, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd8", " ", 11531, 4271, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd9", 9054, 4271, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip9", 9054, 4271, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd9", 9224, 4441, 1995, 1097)
    Set c = AddLabel("lblProd9", " ", 9111, 5566, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice9", " ", 9111, 5849, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd9", " ", 9054, 4271, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd10", 6577, 4271, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip10", 6577, 4271, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd10", 6747, 4441, 1995, 1097)
    Set c = AddLabel("lblProd10", " ", 6634, 5566, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice10", " ", 6634, 5849, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd10", " ", 6577, 4271, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd11", 16485, 6246, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip11", 16485, 6246, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd11", 16655, 6416, 1995, 1097)
    Set c = AddLabel("lblProd11", " ", 16542, 7541, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice11", " ", 16542, 7824, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd11", " ", 16485, 6246, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd12", 14008, 6246, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip12", 14008, 6246, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd12", 14178, 6416, 1995, 1097)
    Set c = AddLabel("lblProd12", " ", 14065, 7541, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice12", " ", 14065, 7824, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd12", " ", 14008, 6246, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd13", 11531, 6246, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip13", 11531, 6246, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd13", 11701, 6416, 1995, 1097)
    Set c = AddLabel("lblProd13", " ", 11588, 7541, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice13", " ", 11588, 7824, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd13", " ", 11531, 6246, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd14", 9054, 6246, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip14", 9054, 6246, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd14", 9224, 6416, 1995, 1097)
    Set c = AddLabel("lblProd14", " ", 9111, 7541, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice14", " ", 9111, 7824, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd14", " ", 9054, 6246, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddRect("boxProd15", 6577, 6246, 2335, 1862, CLR_SURFACE)
    Set c = AddRect("boxProdStrip15", 6577, 6246, 2335, 102, CLR_PRIMARY)
    Set c = AddImage("imgProd15", 6747, 6416, 1995, 1097)
    Set c = AddLabel("lblProd15", " ", 6634, 7541, 2222, 284, 12, True, CLR_TEXT, "", 2)
    Set c = AddLabel("lblPrice15", " ", 6634, 7824, 2222, 255, 11, True, CLR_ACCENT, "", 2)
    Set c = AddButton("btnProd15", " ", 6577, 6246, 2335, 1862, "secondary")
    SetCtlProp c, "Transparent", True
    c.OnClick = EP
    Set c = AddButton("btnProdNext", "«· «·Ì", 6577, 8222, 1701, 907, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblProdPage", " ", 8391, 8477, 8618, 397, 12, False, CLR_MUTED, "", 2)
    Set c = AddButton("btnProdPrev", "«·”«»ﬁ", 17123, 8222, 1701, 907, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnUnload = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    TouchLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf
    s = s & "    Cancel = Not POSUnload(Me)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtOrderName_AfterUpdate()" & vbCrLf
    s = s & "    OrderNameChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReprint_Click()" & vbCrLf
    s = s & "    ReprintLast Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTypeTakeaway_Click()" & vbCrLf
    s = s & "    SetOrderType Me, ""TAKEAWAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTypeDineIn_Click()" & vbCrLf
    s = s & "    SetOrderType Me, ""DINE_IN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayCash_Click()" & vbCrLf
    s = s & "    TouchPayCash Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayCard_Click()" & vbCrLf
    s = s & "    TouchPayCard Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCancelOrder_Click()" & vbCrLf
    s = s & "    NewSale Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCatDown_Click()" & vbCrLf
    s = s & "    CategoryPage Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCatUp_Click()" & vbCrLf
    s = s & "    CategoryPage Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat1_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat2_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat3_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat4_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCat5_Click()" & vbCrLf
    s = s & "    CategoryTileClick Me, 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd1_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd2_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd3_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd4_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd5_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd6_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd7_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd8_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 8" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd9_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 9" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd10_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 10" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd11_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 11" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd12_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 12" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd13_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 13" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd14_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 14" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProd15_Click()" & vbCrLf
    s = s & "    ProductTileClick Me, 15" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProdNext_Click()" & vbCrLf
    s = s & "    ProductPage Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnProdPrev_Click()" & vbCrLf
    s = s & "    ProductPage Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,18994,850,0,1000,0,0;btnClose,17463,170,1361,510,1000,0,0,0;btnReprint,14685,170,2665,510,1000,0,0,0;subLines,170,3402,6237,3204,0,0,0,1000;boxTotals,170,6691,6237,1134,0,0,1000,0;lblCapSubTotal,4309,6748,2041,255,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblSubTotal,2835,6748,1418,255,0,0,1000,0;lblCapDiscount,4309,7032,2041,255,0,0,1000,0;lblDiscount,2835,7032,1418,255,0,0,1000,0;lblCapTax,4309,7316,2041,255,0,0,1000,0;lblTax,2835,7316,1418,255,0,0,1000,0;lblItems,2835,7598,3515,215,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCapTotal,283,6719,2438,255,0,0,1000,0;lblTotal,283,7003,2438,737,0,0,1000,0;lblStatus,170,7882,6237,284,0,0,1000,0;btnPayCash,3969,8222,2438,907,0,0,1000,0;btnPayCard,1531,8222,2381,907,0,0,1000,0;btnCancelOrder,170,8222,1304,907,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnCatUp,18030,936,794,879,1000,0,0,0;boxCat1,15920,936,1996,879,800,200,0,0;imgCat1,17321,1049,510,652,800,200,0,0;lblCat1,15977,1191,1316,397,800,200,0,0;btnCat1,15920,936,1996,879,800,200,0,0;boxCat2,13811,936,1996,879,600,200,0,0""" & vbCrLf
    s = s & "    spec = spec & "";imgCat2,15212,1049,510,652,600,200,0,0;lblCat2,13868,1191,1316,397,600,200,0,0;btnCat2,13811,936,1996,879,600,200,0,0;boxCat3,11702,936,1996,879,400,200,0,0;imgCat3,13103,1049,510,652,400,200,0,0;lblCat3,11759,1191,1316,397,400,200,0,0""" & vbCrLf
    s = s & "    spec = spec & "";btnCat3,11702,936,1996,879,400,200,0,0;boxCat4,9593,936,1996,879,200,200,0,0;imgCat4,10994,1049,510,652,200,200,0,0;lblCat4,9650,1191,1316,397,200,200,0,0;btnCat4,9593,936,1996,879,200,200,0,0;boxCat5,7484,936,1996,879,0,200,0,0""" & vbCrLf
    s = s & "    spec = spec & "";imgCat5,8885,1049,510,652,0,200,0,0;lblCat5,7541,1191,1316,397,0,200,0,0;btnCat5,7484,936,1996,879,0,200,0,0;lblCategoryTitle,6577,1899,12247,340,0,1000,0,0;boxProd1,16485,2296,2335,1862,800,200,0,333;boxProdStrip1,16485,2296,2335,102,800,200,0,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd1,16655,2466,1995,1097,800,200,0,333;lblProd1,16542,3591,2222,284,800,200,0,333;lblPrice1,16542,3874,2222,255,800,200,0,333;btnProd1,16485,2296,2335,1862,800,200,0,333;boxProd2,14008,2296,2335,1862,600,200,0,333;boxProdStrip2,14008,2296,2335,102,600,200,0,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd2,14178,2466,1995,1097,600,200,0,333;lblProd2,14065,3591,2222,284,600,200,0,333;lblPrice2,14065,3874,2222,255,600,200,0,333;btnProd2,14008,2296,2335,1862,600,200,0,333;boxProd3,11531,2296,2335,1862,400,200,0,333;boxProdStrip3,11531,2296,2335,102,400,200,0,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd3,11701,2466,1995,1097,400,200,0,333;lblProd3,11588,3591,2222,284,400,200,0,333;lblPrice3,11588,3874,2222,255,400,200,0,333;btnProd3,11531,2296,2335,1862,400,200,0,333;boxProd4,9054,2296,2335,1862,200,200,0,333;boxProdStrip4,9054,2296,2335,102,200,200,0,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd4,9224,2466,1995,1097,200,200,0,333;lblProd4,9111,3591,2222,284,200,200,0,333;lblPrice4,9111,3874,2222,255,200,200,0,333;btnProd4,9054,2296,2335,1862,200,200,0,333;boxProd5,6577,2296,2335,1862,0,200,0,333;boxProdStrip5,6577,2296,2335,102,0,200,0,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd5,6747,2466,1995,1097,0,200,0,333;lblProd5,6634,3591,2222,284,0,200,0,333;lblPrice5,6634,3874,2222,255,0,200,0,333;btnProd5,6577,2296,2335,1862,0,200,0,333;boxProd6,16485,4271,2335,1862,800,200,333,333;boxProdStrip6,16485,4271,2335,102,800,200,333,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd6,16655,4441,1995,1097,800,200,333,333;lblProd6,16542,5566,2222,284,800,200,333,333;lblPrice6,16542,5849,2222,255,800,200,333,333;btnProd6,16485,4271,2335,1862,800,200,333,333;boxProd7,14008,4271,2335,1862,600,200,333,333;boxProdStrip7,14008,4271,2335,102,600,200,333,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd7,14178,4441,1995,1097,600,200,333,333;lblProd7,14065,5566,2222,284,600,200,333,333;lblPrice7,14065,5849,2222,255,600,200,333,333;btnProd7,14008,4271,2335,1862,600,200,333,333;boxProd8,11531,4271,2335,1862,400,200,333,333;boxProdStrip8,11531,4271,2335,102,400,200,333,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd8,11701,4441,1995,1097,400,200,333,333;lblProd8,11588,5566,2222,284,400,200,333,333;lblPrice8,11588,5849,2222,255,400,200,333,333;btnProd8,11531,4271,2335,1862,400,200,333,333;boxProd9,9054,4271,2335,1862,200,200,333,333;boxProdStrip9,9054,4271,2335,102,200,200,333,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd9,9224,4441,1995,1097,200,200,333,333;lblProd9,9111,5566,2222,284,200,200,333,333;lblPrice9,9111,5849,2222,255,200,200,333,333;btnProd9,9054,4271,2335,1862,200,200,333,333;boxProd10,6577,4271,2335,1862,0,200,333,333;boxProdStrip10,6577,4271,2335,102,0,200,333,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd10,6747,4441,1995,1097,0,200,333,333;lblProd10,6634,5566,2222,284,0,200,333,333;lblPrice10,6634,5849,2222,255,0,200,333,333;btnProd10,6577,4271,2335,1862,0,200,333,333;boxProd11,16485,6246,2335,1862,800,200,666,333;boxProdStrip11,16485,6246,2335,102,800,200,666,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd11,16655,6416,1995,1097,800,200,666,333;lblProd11,16542,7541,2222,284,800,200,666,333;lblPrice11,16542,7824,2222,255,800,200,666,333;btnProd11,16485,6246,2335,1862,800,200,666,333;boxProd12,14008,6246,2335,1862,600,200,666,333;boxProdStrip12,14008,6246,2335,102,600,200,666,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd12,14178,6416,1995,1097,600,200,666,333;lblProd12,14065,7541,2222,284,600,200,666,333;lblPrice12,14065,7824,2222,255,600,200,666,333;btnProd12,14008,6246,2335,1862,600,200,666,333;boxProd13,11531,6246,2335,1862,400,200,666,333;boxProdStrip13,11531,6246,2335,102,400,200,666,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd13,11701,6416,1995,1097,400,200,666,333;lblProd13,11588,7541,2222,284,400,200,666,333;lblPrice13,11588,7824,2222,255,400,200,666,333;btnProd13,11531,6246,2335,1862,400,200,666,333;boxProd14,9054,6246,2335,1862,200,200,666,333;boxProdStrip14,9054,6246,2335,102,200,200,666,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd14,9224,6416,1995,1097,200,200,666,333;lblProd14,9111,7541,2222,284,200,200,666,333;lblPrice14,9111,7824,2222,255,200,200,666,333;btnProd14,9054,6246,2335,1862,200,200,666,333;boxProd15,6577,6246,2335,1862,0,200,666,333;boxProdStrip15,6577,6246,2335,102,0,200,666,333""" & vbCrLf
    s = s & "    spec = spec & "";imgProd15,6747,6416,1995,1097,0,200,666,333;lblProd15,6634,7541,2222,284,0,200,666,333;lblPrice15,6634,7824,2222,255,0,200,666,333;btnProd15,6577,6246,2335,1862,0,200,666,333;btnProdNext,6577,8222,1701,907,0,0,1000,0;lblProdPage,8391,8477,8618,397,0,1000,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnProdPrev,17123,8222,1701,907,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 18994, 9299, 0, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCafePOS", s
    Exit Sub
EH:
    AbortForm "frmCafePOS", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCafeItem()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCafeItem", "ŒÌ«—«  «·„‘—Ê»", "", 8845, 7711, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 8845, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„‘—Ê»", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «·ÕÃ„ Ê«·≈÷«›« ° À„ ´≈÷«›… ··ÿ·»ª", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtItemProduct", "", 7484, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtSize", "", 7682, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtAddOns", "", 7880, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtNoteFlags", "", 8078, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtItemQty", "", 8276, 57, 170, 170)
    SetCtlProp c, "Visible", False
    Set c = AddImage("imgItem", 6464, 992, 2155, 1389)
    Set c = AddLabel("lblCapSize", "«·ÕÃ„", 227, 992, 6124, 312, 12, True, CLR_TEXT, "", 3)
    Set c = AddButton("btnSizeL", "ﬂ»Ì—", 227, 1361, 1984, 1021, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnSizeM", "Ê”ÿ", 2296, 1361, 1984, 1021, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddButton("btnSizeS", "’€Ì—", 4365, 1361, 1984, 1021, "secondary")
    c.FontSize = 14
    c.OnClick = EP
    Set c = AddLabel("lblCapAddOns", "«·≈÷«›« ", 227, 2495, 8391, 312, 12, True, CLR_TEXT, "", 3)
    Set c = AddButton("btnAdd1", " ", 6605, 2835, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd2", " ", 4479, 2835, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd3", " ", 2353, 2835, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd4", " ", 227, 2835, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd5", " ", 6605, 3629, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd6", " ", 4479, 3629, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd7", " ", 2353, 3629, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAdd8", " ", 227, 3629, 2013, 709, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblCapNotes", "„·«ÕŸ« ", 227, 4451, 8391, 312, 12, True, CLR_TEXT, "", 3)
    Set c = AddButton("btnNote1", " ", 6605, 4791, 2013, 680, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNote2", " ", 4479, 4791, 2013, 680, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNote3", " ", 2353, 4791, 2013, 680, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNote4", " ", 227, 4791, 2013, 680, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnQtyPlus", "+", 227, 5642, 907, 794, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddLabel("lblItemQty", "1", 1191, 5783, 907, 510, 20, True, CLR_TEXT, "", 2)
    Set c = AddButton("btnQtyMinus", "-", 2155, 5642, 907, 794, "secondary")
    c.FontSize = 20
    c.OnClick = EP
    Set c = AddLabel("lblItemTotal", " ", 3175, 5783, 5443, 510, 18, True, CLR_ACCENT, "", 3)
    Set c = AddButton("btnItemAdd", "≈÷«›… ··ÿ·»", 3459, 6634, 5159, 907, "primary")
    c.FontSize = 16
    c.OnClick = EP
    Set c = AddButton("btnItemCancel", "≈·€«¡", 227, 6634, 3118, 907, "secondary")
    c.FontSize = 16
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ItemLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSizeL_Click()" & vbCrLf
    s = s & "    ItemSize Me, ""L""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSizeM_Click()" & vbCrLf
    s = s & "    ItemSize Me, ""M""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSizeS_Click()" & vbCrLf
    s = s & "    ItemSize Me, ""S""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd1_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd2_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd3_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd4_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd5_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 5" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd6_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 6" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd7_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 7" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAdd8_Click()" & vbCrLf
    s = s & "    ItemToggleAddOn Me, 8" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNote1_Click()" & vbCrLf
    s = s & "    ItemToggleNote Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNote2_Click()" & vbCrLf
    s = s & "    ItemToggleNote Me, 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNote3_Click()" & vbCrLf
    s = s & "    ItemToggleNote Me, 3" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNote4_Click()" & vbCrLf
    s = s & "    ItemToggleNote Me, 4" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQtyPlus_Click()" & vbCrLf
    s = s & "    ItemQty Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnQtyMinus_Click()" & vbCrLf
    s = s & "    ItemQty Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnItemAdd_Click()" & vbCrLf
    s = s & "    ItemConfirm Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnItemCancel_Click()" & vbCrLf
    s = s & "    ItemCancel Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCafeItem", s
    Exit Sub
EH:
    AbortForm "frmCafeItem", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmTreasury()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmTreasury", "«·Œ“Ì‰…", "", 15309, 9015, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·Œ“Ì‰… Ê«·’‰«œÌﬁ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "√—’œ… «·Œ“Ì‰… «·—∆Ì”Ì… Ê’‰«œÌﬁ «·ﬂ«‘Ì— ÊÕ—ﬂ… ﬂ· ÌÊ„", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnCashIn", "”‰œ ﬁ»÷ ‰ﬁœÌ…", 227, 1021, 1814, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCashOut", "”‰œ ’—› ‰ﬁœÌ…", 2154, 1021, 1814, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnTransfer", " ÕÊÌ· »Ì‰ «·’‰«œÌﬁ", 4081, 1021, 2041, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClosing", " ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—", 6235, 1021, 2155, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnBoxes", "«·’‰«œÌﬁ", 8503, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnBanks", "«·»‰Êﬂ", 9977, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCheques", "«·‘Ìﬂ« ", 11451, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblBoxesCap", "«·Œ“Ì‰… Ê«·’‰«œÌﬁ («·—’Ìœ «·¬‰)", 227, 1701, 5103, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstBoxes", 227, 2041, 5103, 6690, 4, "0;2381;1304;1304", True)
    c.FontSize = 11
    c.AfterUpdate = EP
    Set c = AddLabel("lblBoxName", " ", 5670, 1701, 9412, 454, 15, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtDay", "", 5670, 2552, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    c.AfterUpdate = EP
    Set c = AddLabel("lblDay", "«·ÌÊ„", 5670, 2240, 1814, 284, 9, False, CLR_MUTED, "txtDay", 0)
    Set c = AddButton("btnPrevDay", "«·ÌÊ„ «·”«»ﬁ", 7598, 2552, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnNextDay", "«·ÌÊ„ «· «·Ì", 9185, 2552, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnToday", "«·ÌÊ„", 10773, 2552, 1021, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrintDay", "ÿ»«⁄… Õ—ﬂ… «·ÌÊ„", 11907, 2552, 2041, 454, "primary")
    c.OnClick = EP
    Set c = AddRect("boxOpening", 5670, 3175, 2225, 964, CLR_SURFACE)
    Set c = AddLabel("lblOpeningCap", "—’Ìœ √Ê· «·ÌÊ„", 5812, 3260, 1941, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblOpening", "-", 5812, 3572, 1941, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxIn", 8065, 3175, 2225, 964, CLR_SURFACE)
    Set c = AddLabel("lblInCap", "«·„ﬁ»Ê÷« ", 8207, 3260, 1941, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblIn", "-", 8207, 3572, 1941, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxOut", 10460, 3175, 2225, 964, CLR_SURFACE)
    Set c = AddLabel("lblOutCap", "«·„œ›Ê⁄«  Ê«·„’—Ê›« ", 10602, 3260, 1941, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblOut", "-", 10602, 3572, 1941, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxClosing", 12855, 3175, 2225, 964, CLR_SURFACE)
    Set c = AddLabel("lblClosingCap", "—’Ìœ ¬Œ— «·ÌÊ„", 12997, 3260, 1941, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblClosing", "-", 12997, 3572, 1941, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCurrent", " ", 5670, 4252, 9412, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddList("lstMoves", 5670, 4649, 9412, 4082, 6, "850;2495;1644;2495;964;964", True)
    m_frm.OnLoad = EP
    m_frm.OnActivate = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    TreasuryLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Activate()" & vbCrLf
    s = s & "    TreasuryActivate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstBoxes_AfterUpdate()" & vbCrLf
    s = s & "    TreasuryShow Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtDay_AfterUpdate()" & vbCrLf
    s = s & "    TreasuryShow Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCashIn_Click()" & vbCrLf
    s = s & "    TreasuryOpen Me, ""IN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCashOut_Click()" & vbCrLf
    s = s & "    TreasuryOpen Me, ""OUT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTransfer_Click()" & vbCrLf
    s = s & "    TreasuryOpen Me, ""TRANSFER""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClosing_Click()" & vbCrLf
    s = s & "    TreasuryOpen Me, ""CLOSING""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBoxes_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCashBoxes""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBanks_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBanks""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCheques_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCheques"", 0, ""IN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrevDay_Click()" & vbCrLf
    s = s & "    TreasuryDayStep Me, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNextDay_Click()" & vbCrLf
    s = s & "    TreasuryDayStep Me, 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnToday_Click()" & vbCrLf
    s = s & "    TreasuryToday Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrintDay_Click()" & vbCrLf
    s = s & "    PrintCashDay Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;btnClose,13721,1021,1361,482,1000,0,0,0;lstBoxes,227,2041,5103,6690,0,0,0,1000;lblBoxName,5670,1701,9412,454,0,1000,0,0;lblCurrent,5670,4252,9412,340,0,1000,0,0;lstMoves,5670,4649,9412,4082,0,1000,0,1000""" & vbCrLf
    s = s & "    FitControls Me, 15309, 9015, -2607, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmTreasury", s
    Exit Sub
EH:
    AbortForm "frmTreasury", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCashVoucher()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCashVoucher", "”‰œ ‰ﬁœÌ…", "", 9072, 7598, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "”‰œ ‰ﬁœÌ…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁ»÷ ‰ﬁœÌ… ›Ì ’‰œÊﬁ° √Ê ’—› „‰Â° √Ê  ÕÊÌ· »Ì‰ ’‰œÊﬁÌ‰", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboVoucherType", "", 227, 1361, 4196, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblVoucherType", "‰Ê⁄ «·”‰œ", 227, 1049, 4196, 284, 9, False, CLR_MUTED, "cboVoucherType", 0)
    Set c = AddCombo("cboCategory", "", 4649, 1361, 4196, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblCategory", "«·»‰œ", 4649, 1049, 4196, 284, 9, False, CLR_MUTED, "cboCategory", 0)
    Set c = AddCombo("cboBox", "", 227, 2211, 4196, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBox", "«·’‰œÊﬁ", 227, 1899, 4196, 284, 9, False, CLR_MUTED, "cboBox", 0)
    Set c = AddCombo("cboToBox", "", 4649, 2211, 4196, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblToBox", "≈·Ï ’‰œÊﬁ", 4649, 1899, 4196, 284, 9, False, CLR_MUTED, "cboToBox", 0)
    Set c = AddLabel("lblBoxBalance", " ", 227, 2722, 8618, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddCombo("cboExpenseType", "", 227, 3402, 2948, 454, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes WHERE IsActive = True ORDER BY ExpenseTypeName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblExpenseType", "‰Ê⁄ «·„’—Ê› *", 227, 3090, 2948, 284, 9, False, CLR_MUTED, "cboExpenseType", 0)
    Set c = AddButton("btnNewExpenseType", "‰Ê⁄ ÃœÌœ", 3289, 3402, 1134, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtAmount", "", 4649, 3402, 4196, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ *", 4649, 3090, 4196, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboEmployee", "", 227, 4309, 4196, 454, "SELECT EmployeeID, EmployeeName FROM Employees WHERE IsActive = True ORDER BY EmployeeName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblEmployee", "«·„ÊŸ› ’«Õ» «·”·›… *", 227, 3997, 4196, 284, 9, False, CLR_MUTED, "cboEmployee", 0)
    Set c = AddText("txtParty", "", 227, 5103, 8618, 454)
    Set c = AddLabel("lblParty", "«·„” ·„", 227, 4791, 8618, 284, 9, False, CLR_MUTED, "txtParty", 0)
    Set c = AddText("txtDescription", "", 227, 5897, 8618, 454)
    Set c = AddLabel("lblDescription", "«·»Ì«‰", 227, 5585, 8618, 284, 9, False, CLR_MUTED, "txtDescription", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «·”‰œ", 227, 6691, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "Õ›Ÿ Êÿ»«⁄…", 2268, 6691, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 7371, 6691, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    VoucherLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboVoucherType_AfterUpdate()" & vbCrLf
    s = s & "    VoucherTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCategory_AfterUpdate()" & vbCrLf
    s = s & "    VoucherCategoryChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBox_AfterUpdate()" & vbCrLf
    s = s & "    VoucherBoxChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNewExpenseType_Click()" & vbCrLf
    s = s & "    AddExpenseType Me, ""cboExpenseType""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveCashVoucher Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSavePrint_Click()" & vbCrLf
    s = s & "    SaveCashVoucher Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCashVoucher", s
    Exit Sub
EH:
    AbortForm "frmCashVoucher", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCashClosing()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCashClosing", " ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—", "", 10206, 9299, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 10206, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", " ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "⁄œ¯ «·‰ﬁœÌ… Ê —ÕÌ·Â« ··Œ“Ì‰… «·—∆Ì”Ì… √Ê  ”ÊÌ Â« „⁄ «·„«·ﬂ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboBox", "", 227, 1361, 4763, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBox", "«·’‰œÊﬁ", 227, 1049, 4763, 284, 9, False, CLR_MUTED, "cboBox", 0)
    Set c = AddLabel("lblPeriod", " ", 5216, 1418, 4763, 397, 10, False, CLR_MUTED, "", 0)
    Set c = AddText("txtExpected", "", 9639, 57, 340, 227)
    SetCtlProp c, "Visible", False
    SetCtlProp c, "Format", "0.00"
    Set c = AddList("lstBreakdown", 227, 1984, 9752, 2041, 4, "4649;907;2041;2041", True)
    Set c = AddRect("boxOpening", 227, 4139, 2310, 964, CLR_SURFACE)
    Set c = AddLabel("lblOpeningCap", "—’Ìœ «·»œ«Ì…", 369, 4224, 2026, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblOpening", "-", 369, 4536, 2026, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxCashIn", 2707, 4139, 2310, 964, CLR_SURFACE)
    Set c = AddLabel("lblCashInCap", "«·„ﬁ»Ê÷« ", 2849, 4224, 2026, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCashIn", "-", 2849, 4536, 2026, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxCashOut", 5187, 4139, 2310, 964, CLR_SURFACE)
    Set c = AddLabel("lblCashOutCap", "«·„œ›Ê⁄« ", 5329, 4224, 2026, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCashOut", "-", 5329, 4536, 2026, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxExpected", 7667, 4139, 2310, 964, CLR_SURFACE)
    Set c = AddLabel("lblExpectedCap", "«·—’Ìœ «·œ› —Ì («·„›—Ê÷)", 7809, 4224, 2026, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblExpected", "-", 7809, 4536, 2026, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtCounted", "", 227, 5500, 3118, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblCounted", "«·‰ﬁœÌ… «·›⁄·Ì… »«·⁄œ¯ *", 227, 5188, 3118, 284, 9, False, CLR_MUTED, "txtCounted", 0)
    Set c = AddLabel("lblDifference", " ", 3515, 5557, 6464, 454, 13, True, CLR_MUTED, "", 0)
    Set c = AddCombo("cboDestination", "", 227, 6520, 3118, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblDestination", "«· —ÕÌ· ≈·Ï", 227, 6208, 3118, 284, 9, False, CLR_MUTED, "cboDestination", 0)
    Set c = AddCombo("cboToBox", "", 3515, 6520, 3118, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblToBox", "«·Œ“Ì‰… «·„” ·„…", 3515, 6208, 3118, 284, 9, False, CLR_MUTED, "cboToBox", 0)
    Set c = AddText("txtTransfer", "", 6804, 6520, 3175, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTransfer", "«·„»·€ «·„—ÕÛ¯·", 6804, 6208, 3175, 284, 9, False, CLR_MUTED, "txtTransfer", 0)
    Set c = AddLabel("lblKept", " ", 227, 7059, 9752, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtNotes", "", 227, 7768, 9752, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 227, 7456, 9752, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «· ’›Ì…", 227, 8448, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "Õ›Ÿ Êÿ»«⁄…", 2381, 8448, 1928, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 8505, 8448, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ClosingLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBox_AfterUpdate()" & vbCrLf
    s = s & "    ClosingBoxChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtCounted_AfterUpdate()" & vbCrLf
    s = s & "    ClosingCountChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboDestination_AfterUpdate()" & vbCrLf
    s = s & "    ClosingDestinationChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtTransfer_AfterUpdate()" & vbCrLf
    s = s & "    ClosingRecalc Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveCashClosing Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSavePrint_Click()" & vbCrLf
    s = s & "    SaveCashClosing Me, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCashClosing", s
    Exit Sub
EH:
    AbortForm "frmCashClosing", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmJournal()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmJournal", "ﬁÌÊœ «·ÌÊ„Ì…", "", 15309, 9015, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ﬁÌÊœ «·ÌÊ„Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁÌœ ·ﬂ· ⁄„·Ì… „—»Êÿ »√’·Â«: «› Õ «·ﬁÌœ √Ê √’· «·⁄„·Ì…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtFrom", "", 227, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰  «—ÌŒ", 227, 992, 1701, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 2041, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï  «—ÌŒ", 2041, 992, 1701, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnToday", "«·ÌÊ„", 3856, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisMonth", "Â–« «·‘Â—", 5132, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastMonth", "«·‘Â— «·„«÷Ì", 6408, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisYear", "Â–Â «·”‰…", 7684, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboSourceType", "", 9015, 1304, 2381, 454, "", 2, "0;2268")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblSourceType", "‰Ê⁄ «·⁄„·Ì…", 9015, 992, 2381, 284, 9, False, CLR_MUTED, "cboSourceType", 0)
    Set c = AddText("txtSearch", "", 11510, 1304, 2041, 454)
    c.AfterUpdate = EP
    Set c = AddLabel("lblSearch", "»ÕÀ (—ﬁ„ / »Ì«‰)", 11510, 992, 2041, 284, 9, False, CLR_MUTED, "txtSearch", 0)
    Set c = AddButton("btnShow", "⁄—÷", 13665, 1304, 1418, 454, "primary")
    c.OnClick = EP
    Set c = AddList("lstEntries", 227, 1928, 14855, 2665, 7, "0;1474;1361;1928;1588;6577;1701", True)
    c.AfterUpdate = EP
    c.OnDblClick = EP
    Set c = AddLabel("lblTotals", " ", 227, 4649, 14855, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblLinesCap", "√”ÿ— «·ﬁÌœ «·„Œ «—", 227, 5046, 6804, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstLines", 227, 5386, 14855, 2325, 5, "1247;3402;6010;1814;1814", True)
    Set c = AddButton("btnOpenEntry", "› Õ «·ﬁÌœ", 227, 7881, 1588, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnOpenSource", "› Õ √’· «·⁄„·Ì…", 1928, 7881, 2041, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSync", " ÕœÌÀ «·ﬁÌÊœ", 4082, 7881, 1474, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·ÌÊ„Ì…", 5669, 7881, 1474, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnTrial", "„Ì“«‰ «·„—«Ã⁄…", 7256, 7881, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAccounts", "œ·Ì· «·Õ”«»« ", 8957, 7881, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLedger", "ﬂ‘› Õ”«»", 10658, 7881, 1361, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnManual", "ﬁÌœ ÌœÊÌ", 12132, 7881, 1247, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13721, 7881, 1361, 510, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSync", " ", 227, 8505, 14855, 312, 9, False, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    JournalLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstEntries_AfterUpdate()" & vbCrLf
    s = s & "    JournalEntryPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstEntries_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    JournalOpenEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboSourceType_AfterUpdate()" & vbCrLf
    s = s & "    JournalRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_AfterUpdate()" & vbCrLf
    s = s & "    JournalRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnToday_Click()" & vbCrLf
    s = s & "    JournalQuickPeriod Me, ""TODAY""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisMonth_Click()" & vbCrLf
    s = s & "    JournalQuickPeriod Me, ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastMonth_Click()" & vbCrLf
    s = s & "    JournalQuickPeriod Me, ""LASTMONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisYear_Click()" & vbCrLf
    s = s & "    JournalQuickPeriod Me, ""YEAR""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnShow_Click()" & vbCrLf
    s = s & "    JournalRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenEntry_Click()" & vbCrLf
    s = s & "    JournalOpenEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenSource_Click()" & vbCrLf
    s = s & "    OpenJournalSource Me!lstEntries.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSync_Click()" & vbCrLf
    s = s & "    JournalSync Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintJournal Me, ""JOURNAL""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTrial_Click()" & vbCrLf
    s = s & "    PrintJournal Me, ""TRIAL""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAccounts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLedger_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLedger""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnManual_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmManualEntry""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;btnShow,13665,1304,1418,454,1000,0,0,0;lstEntries,227,1928,14855,2665,0,1000,0,1000;lblTotals,227,4649,14855,340,0,1000,1000,0;lblLinesCap,227,5046,6804,312,0,0,1000,0;lstLines,227,5386,14855,2325,0,1000,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnOpenEntry,227,7881,1588,510,0,0,1000,0;btnOpenSource,1928,7881,2041,510,0,0,1000,0;btnSync,4082,7881,1474,510,0,0,1000,0;btnPrint,5669,7881,1474,510,0,0,1000,0;btnTrial,7256,7881,1588,510,0,0,1000,0;btnAccounts,8957,7881,1588,510,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnLedger,10658,7881,1361,510,0,0,1000,0;btnManual,12132,7881,1247,510,0,0,1000,0;btnClose,13721,7881,1361,510,1000,0,1000,0;lblSync,227,8505,14855,312,0,1000,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 9015, -1190, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmJournal", s
    Exit Sub
EH:
    AbortForm "frmJournal", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmJournalEntry()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmJournalEntry", "ﬁÌœ ÌÊ„Ì…", "", 11907, 6464, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 11907, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ﬁÌœ ÌÊ„Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·ﬁÌœ «·¬·Ì ··⁄„·Ì…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtEntryID", "", 11340, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblHeader", " ", 227, 1021, 11453, 369, 11, True, CLR_TEXT, "", 0)
    Set c = AddLabel("lblDescription", " ", 227, 1418, 11453, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddList("lstLines", 227, 1871, 11453, 2835, 5, "1134;2835;3629;1701;1701", True)
    Set c = AddLabel("lblTotals", " ", 227, 4791, 11453, 369, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnOpenSource", "› Õ √’· «·⁄„·Ì…", 227, 5613, 2155, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·ﬁÌœ", 2495, 5613, 1701, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 10206, 5613, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    JournalEntryLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenSource_Click()" & vbCrLf
    s = s & "    OpenJournalSource Me!txtEntryID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintJournalEntry Me!txtEntryID.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmJournalEntry", s
    Exit Sub
EH:
    AbortForm "frmJournalEntry", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmManualLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmManualLines", "√”ÿ— «·ﬁÌœ «·ÌœÊÌ", "SELECT * FROM tmpManualLines ORDER BY LineNo", 14855, 425, False, True, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("LineNo", "LineNo", 28, 0, 510, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddCombo("AccountCode", "AccountCode", 566, 0, 5273, 425, "SELECT AccountCode, AccountCode & '  ' & AccountName AS Account FROM Accounts WHERE IsPosting = True AND IsActive = True ORDER BY TreeKey", 2, "0;5103")
    SetCtlProp c, "BoundColumn", 1
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddText("Debit", "Debit", 5867, 0, 1701, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("Credit", "Credit", 7596, 0, 1701, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("LineText", "LineText", 9325, 0, 3175, 425)
    c.AfterUpdate = EP
    Set c = AddCombo("LineCenter", "LineCenter", 12528, 0, 1758, 425, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;1701")
    SetCtlProp c, "BoundColumn", 1
    SetCtlProp c, "LimitToList", True
    Set c = AddButton("btnRemove", "Sym(code='ChrW(&HE74D)')", 14314, 17, 454, 391, "danger")
    SetCtlProp c, "FontName", ICON_FONT
    c.OnClick = EP
    s = ""
    s = s & "Private Sub AccountCode_AfterUpdate()" & vbCrLf
    s = s & "    ManualLineChanged Me, ""AccountCode""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Debit_AfterUpdate()" & vbCrLf
    s = s & "    ManualLineChanged Me, ""Debit""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Credit_AfterUpdate()" & vbCrLf
    s = s & "    ManualLineChanged Me, ""Credit""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub LineText_AfterUpdate()" & vbCrLf
    s = s & "    ManualLineChanged Me, ""LineText""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    ManualRemoveLine Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmManualLines", s
    Exit Sub
EH:
    AbortForm "frmManualLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmManualEntry()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmManualEntry", "«·ﬁÌÊœ «·ÌœÊÌ…", "", 15309, 9639, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·ﬁÌÊœ «·ÌœÊÌ…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁÌœ Ìﬂ »Â «·„Õ«”»: «·„œÌ‰ = «·œ«∆‰° ÊÌı—ÕÛ¯· ·ﬁÌÊœ «·ÌÊ„Ì… ⁄‰œ «·Õ›Ÿ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboFind", "", 227, 1304, 6804, 454, "SELECT ManualEntryID, EntryNumber, EntryDate, Description FROM ManualEntries ORDER BY EntryDate DESC, ManualEntryID DESC", 4, "0;1474;1474;4536")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblFind", "› Õ ﬁÌœ „Õ›ÊŸ", 227, 992, 6804, 284, 9, False, CLR_MUTED, "cboFind", 0)
    Set c = AddText("txtEntryID", "", 7144, 1304, 284, 454)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtReversalOf", "", 7541, 1304, 284, 454)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtNumber", "", 227, 2098, 1814, 454)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblNumber", "—ﬁ„ «·ﬁÌœ", 227, 1786, 1814, 284, 9, False, CLR_MUTED, "txtNumber", 0)
    Set c = AddText("txtDate", "", 2155, 2098, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblDate", "«· «—ÌŒ", 2155, 1786, 1814, 284, 9, False, CLR_MUTED, "txtDate", 0)
    Set c = AddText("txtReference", "", 4082, 2098, 2381, 454)
    Set c = AddLabel("lblReference", "«·„—Ã⁄ («Œ Ì«—Ì)", 4082, 1786, 2381, 284, 9, False, CLR_MUTED, "txtReference", 0)
    Set c = AddText("txtDescription", "", 6577, 2098, 5443, 454)
    Set c = AddLabel("lblDescription", "«·»Ì«‰", 6577, 1786, 5443, 284, 9, False, CLR_MUTED, "txtDescription", 0)
    Set c = AddCombo("cboCurrency", "", 12134, 2098, 1474, 454, "SELECT CurrencyCode, CurrencyCode & '  ' & CurrencyName AS Currency FROM Currencies WHERE IsActive = True ORDER BY SortOrder, CurrencyCode", 2, "680;2268")
    SetCtlProp c, "ListWidth", "5.5cm"
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblCurrency", "«·⁄„·… *", 12134, 1786, 1474, 284, 9, False, CLR_MUTED, "cboCurrency", 0)
    Set c = AddText("txtRate", "", 13721, 2098, 1361, 454)
    SetCtlProp c, "Format", "0.0000"
    Set c = AddLabel("lblRate", "«·„⁄«„·", 13721, 1786, 1361, 284, 9, False, CLR_MUTED, "txtRate", 0)
    Set c = AddLabel("lblCol1", "#", 255, 2750, 510, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·Õ”«»", 793, 2750, 5273, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "„œÌ‰", 6094, 2750, 1701, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "œ«∆‰", 7823, 2750, 1701, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "»Ì«‰ «·”ÿ—", 9552, 2750, 3175, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "„—ﬂ“ «· ﬂ·›…", 12755, 2750, 1758, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmManualLines", 227, 3090, 14855, 4309)
    Set c = AddLabel("lblTotals", " ", 227, 7513, 14855, 369, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblStatus", " ", 227, 7910, 14855, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnNew", "ﬁÌœ ÃœÌœ", 227, 8732, 1474, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "Õ›Ÿ «·ﬁÌœ", 1814, 8732, 1588, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnReverse", "ﬁÌœ ⁄ﬂ”Ì", 3515, 8732, 1474, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 5102, 8732, 1247, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnInJournal", "⁄—÷ ›Ì «·ÌÊ„Ì…", 6462, 8732, 1814, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAccounts", "œ·Ì· «·Õ”«»« ", 8389, 8732, 1701, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–› «·ﬁÌœ", 10203, 8732, 1474, 510, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13721, 8732, 1361, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ManualEntryLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboFind_AfterUpdate()" & vbCrLf
    s = s & "    ManualFindPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboCurrency_AfterUpdate()" & vbCrLf
    s = s & "    ManualCurrencyPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    ManualNew Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveManualEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReverse_Click()" & vbCrLf
    s = s & "    ReverseManualEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintManualEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnInJournal_Click()" & vbCrLf
    s = s & "    ManualOpenInJournal Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAccounts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    DeleteManualEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;txtDescription,6577,2098,5443,454,0,1000,0,0;cboCurrency,12134,2098,1474,454,1000,0,0,0;lblCurrency,12134,1786,1474,284,1000,0,0,0;txtRate,13721,2098,1361,454,1000,0,0,0;lblRate,13721,1786,1361,284,1000,0,0,0""" & vbCrLf
    s = s & "    spec = spec & "";lblCol6,12755,2750,1758,312,1000,0,0,0;subLines,227,3090,14855,4309,0,1000,0,1000;lblTotals,227,7513,14855,369,0,1000,1000,0;lblStatus,227,7910,14855,340,0,1000,1000,0;btnNew,227,8732,1474,510,0,0,1000,0;btnSave,1814,8732,1588,510,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnReverse,3515,8732,1474,510,0,0,1000,0;btnPrint,5102,8732,1247,510,0,0,1000,0;btnInJournal,6462,8732,1814,510,0,0,1000,0;btnAccounts,8389,8732,1701,510,0,0,1000,0;btnDelete,10203,8732,1474,510,0,0,1000,0;btnClose,13721,8732,1361,510,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 9639, -2834, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmManualEntry", s
    Exit Sub
EH:
    AbortForm "frmManualEntry", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmLedger()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmLedger", "ﬂ‘› Õ”«»", "", 15309, 8732, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "ﬂ‘› Õ”«» Êœ› — «·√” «–", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "Õ—ﬂ… √Ì Õ”«» »—’Ìœ √Ê· «·„œ… Ê«·—’Ìœ »⁄œ ﬂ· ﬁÌœ∫ «·Õ”«» «·—∆Ì”Ì Ì‘„· Õ”«»« Â «· «»⁄…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboAccount", "", 227, 1304, 5103, 454, "SELECT AccountCode, AccountCode & '  ' & Space((AccountLevel - 1) * 2) & AccountName AS Account FROM Accounts ORDER BY TreeKey", 2, "0;5103")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblAccount", "«·Õ”«»", 227, 992, 5103, 284, 9, False, CLR_MUTED, "cboAccount", 0)
    Set c = AddText("txtFrom", "", 5443, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰  «—ÌŒ", 5443, 992, 1701, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 7258, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï  «—ÌŒ", 7258, 992, 1701, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnThisMonth", "Â–« «·‘Â—", 9072, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastMonth", "«·‘Â— «·„«÷Ì", 10348, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisYear", "Â–Â «·”‰…", 11624, 1304, 1191, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnShow", "⁄—÷", 13041, 1304, 1418, 454, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblCapOpening", "—’Ìœ √Ê· «·„œ…", 227, 1956, 3600, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblOpening", "-", 227, 2240, 3600, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCapDebit", "„œÌ‰ «·› —…", 3940, 1956, 3600, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblDebit", "-", 3940, 2240, 3600, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCapCredit", "œ«∆‰ «·› —…", 7653, 1956, 3600, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCredit", "-", 7653, 2240, 3600, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCapClosing", "«·—’Ìœ «·Œ «„Ì", 11366, 1956, 3600, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblClosing", "-", 11366, 2240, 3600, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstLines", 227, 2807, 14855, 3969, 10, "0;1247;1361;1701;1474;3175;1701;1361;1361;1474", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblInfo", " ", 227, 6861, 14855, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnOpenEntry", "› Õ «·ﬁÌœ", 227, 7371, 1474, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnOpenSource", "› Õ √’· «·⁄„·Ì…", 1814, 7371, 2041, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrintStatement", "ÿ»«⁄… ﬂ‘› «·Õ”«»", 3968, 7371, 2155, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrintLedger", "œ› — «·√” «–", 6236, 7371, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnManual", "ﬁÌœ ÌœÊÌ", 7937, 7371, 1361, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAccounts", "œ·Ì· «·Õ”«»« ", 9411, 7371, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnFinancials", "«·ﬁÊ«∆„ «·„«·Ì…", 11112, 7371, 1814, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13721, 7371, 1361, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    LedgerLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboAccount_AfterUpdate()" & vbCrLf
    s = s & "    LedgerRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstLines_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    LedgerOpenEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisMonth_Click()" & vbCrLf
    s = s & "    LedgerQuickPeriod Me, ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastMonth_Click()" & vbCrLf
    s = s & "    LedgerQuickPeriod Me, ""LASTMONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisYear_Click()" & vbCrLf
    s = s & "    LedgerQuickPeriod Me, ""YEAR""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnShow_Click()" & vbCrLf
    s = s & "    LedgerRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenEntry_Click()" & vbCrLf
    s = s & "    LedgerOpenEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOpenSource_Click()" & vbCrLf
    s = s & "    LedgerOpenSource Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrintStatement_Click()" & vbCrLf
    s = s & "    PrintLedger Me, ""STATEMENT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrintLedger_Click()" & vbCrLf
    s = s & "    PrintLedger Me, ""LEDGER""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnManual_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmManualEntry""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAccounts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFinancials_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmFinancials""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;btnShow,13041,1304,1418,454,1000,0,0,0;lstLines,227,2807,14855,3969,0,1000,0,1000;lblInfo,227,6861,14855,340,0,1000,1000,0;btnOpenEntry,227,7371,1474,510,0,0,1000,0;btnOpenSource,1814,7371,2041,510,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnPrintStatement,3968,7371,2155,510,0,0,1000,0;btnPrintLedger,6236,7371,1588,510,0,0,1000,0;btnManual,7937,7371,1361,510,0,0,1000,0;btnAccounts,9411,7371,1588,510,0,0,1000,0;btnFinancials,11112,7371,1814,510,0,0,1000,0;btnClose,13721,7371,1361,510,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 8732, -2494, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmLedger", s
    Exit Sub
EH:
    AbortForm "frmLedger", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmFinancials()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmFinancials", "«·ﬁÊ«∆„ «·„«·Ì…", "", 15309, 8732, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·ﬁÊ«∆„ «·„«·Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁ«∆„… «·œŒ· Ê«·„Ì“«‰Ì… «·⁄„Ê„Ì… „‰ «·ﬁÌÊœ° „⁄ › —… «·„ﬁ«—‰…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboStatement", "", 227, 1304, 2608, 454, "INCOME;ﬁ«∆„… «·œŒ·;BALANCE;«·„Ì“«‰Ì… «·⁄„Ê„Ì…", 2, "0;2552")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblStatement", "«·ﬁ«∆„…", 227, 992, 2608, 284, 9, False, CLR_MUTED, "cboStatement", 0)
    Set c = AddText("txtFrom", "", 2948, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰  «—ÌŒ", 2948, 992, 1701, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 4763, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï  «—ÌŒ («·„Ì“«‰Ì… ›Ì Â–« «·ÌÊ„)", 4763, 992, 1701, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnThisMonth", "Â–« «·‘Â—", 6577, 1304, 1389, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastMonth", "«·‘Â— «·„«÷Ì", 8051, 1304, 1389, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisYear", "Â–Â «·”‰…", 9525, 1304, 1389, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastYear", "«·”‰… «·„«÷Ì…", 10999, 1304, 1389, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnShow", "⁄—÷", 12587, 1304, 1418, 454, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblCompare", " ", 227, 1871, 14855, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCap1", " ", 227, 2240, 4838, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblVal1", "-", 227, 2552, 4838, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCap2", " ", 5178, 2240, 4838, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblVal2", "-", 5178, 2552, 4838, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCap3", " ", 10129, 2240, 4838, 284, 9, True, CLR_MUTED, "", 0)
    Set c = AddLabel("lblVal3", "-", 10129, 2552, 4838, 425, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstRows", 227, 3090, 14855, 3856, 6, "0;6010;2155;2155;2155;2381", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblInfo", "‰ﬁ— „“œÊÃ ⁄·Ï Õ”«» Ì› Õ ﬂ‘› Õ”«»Â ··› —… ‰›”Â«", 227, 7031, 14855, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·ﬁ«∆„…", 227, 7485, 1814, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnLedger", "ﬂ‘› Õ”«»", 2154, 7485, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnTrial", "„Ì“«‰ «·„—«Ã⁄…", 3855, 7485, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClosing", "≈ﬁ›«· «·› —« ", 5556, 7485, 1531, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnVat", "«·≈ﬁ—«— «·÷—Ì»Ì", 7200, 7485, 1644, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnAssets", "«·√’Ê· «·À«» …", 8957, 7485, 1701, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayroll", "«·—Ê« »", 10771, 7485, 1247, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnBudget", "«·„Ê«“‰…", 12131, 7485, 1247, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13721, 7485, 1361, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    FinancialsLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboStatement_AfterUpdate()" & vbCrLf
    s = s & "    FinancialsRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstRows_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    FinancialsOpenLedger Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisMonth_Click()" & vbCrLf
    s = s & "    FinancialsQuickPeriod Me, ""MONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastMonth_Click()" & vbCrLf
    s = s & "    FinancialsQuickPeriod Me, ""LASTMONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnThisYear_Click()" & vbCrLf
    s = s & "    FinancialsQuickPeriod Me, ""YEAR""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastYear_Click()" & vbCrLf
    s = s & "    FinancialsQuickPeriod Me, ""LASTYEAR""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnShow_Click()" & vbCrLf
    s = s & "    FinancialsRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintFinancials Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLedger_Click()" & vbCrLf
    s = s & "    FinancialsOpenLedger Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTrial_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmJournal""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClosing_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPeriodClosing"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnVat_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmVatReturn"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAssets_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAssets"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayroll_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPayroll""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBudget_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBudget""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,15309,850,0,1000,0,0;btnShow,12587,1304,1418,454,1000,0,0,0;lblCompare,227,1871,14855,312,0,1000,0,0;lstRows,227,3090,14855,3856,0,1000,0,1000;lblInfo,227,7031,14855,312,0,1000,1000,0;btnPrint,227,7485,1814,510,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnLedger,2154,7485,1588,510,0,0,1000,0;btnTrial,3855,7485,1588,510,0,0,1000,0;btnClosing,5556,7485,1531,510,0,0,1000,0;btnVat,7200,7485,1644,510,0,0,1000,0;btnAssets,8957,7485,1701,510,0,0,1000,0;btnPayroll,10771,7485,1247,510,0,0,1000,0""" & vbCrLf
    s = s & "    spec = spec & "";btnBudget,12131,7485,1247,510,0,0,1000,0;btnClose,13721,7485,1361,510,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 15309, 8732, -2381, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmFinancials", s
    Exit Sub
EH:
    AbortForm "frmFinancials", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPeriodClosing()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPeriodClosing", "≈ﬁ›«· «·› —«  Ê«·”‰… «·„«·Ì…", "", 12474, 9639, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 12474, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "≈ﬁ›«· «·› —«  Ê«·”‰… «·„«·Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "»⁄œ «·≈ﬁ›«· ·« Ìı÷«› Ê·« Ìı⁄œÛ¯· Ê·« ÌıÕ–› √Ì „” ‰œ » «—ÌŒ „ﬁ›·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddLabel("lblState", " ", 227, 1049, 12020, 454, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblPeriodCap", "≈ﬁ›«· › —… (‘Â— √Ê √ﬂÀ—)", 227, 1616, 6804, 340, 11, True, CLR_TEXT, "", 0)
    Set c = AddText("txtThrough", "", 227, 2268, 1928, 482)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblThrough", "„ﬁ›·… Õ Ï ÌÊ„", 227, 1956, 1928, 284, 9, False, CLR_MUTED, "txtThrough", 0)
    Set c = AddText("txtNotes", "", 2268, 2268, 9979, 482)
    Set c = AddLabel("lblNotes", "«·”»» / „·«ÕŸ«  („ÿ·Ê» ·≈⁄«œ… «·› Õ)", 2268, 1956, 9979, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnClosePeriod", "≈ﬁ›«· Õ Ï Â–« «·ÌÊ„", 227, 2892, 2608, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnReopenPeriod", "≈⁄«œ… «·› Õ ≈·Ï Â–« «·ÌÊ„", 2948, 2892, 2948, 510, "danger")
    c.OnClick = EP
    Set c = AddLabel("lblYearCap", "≈ﬁ›«· «·”‰… «·„«·Ì… («·≈Ì—«œ«  Ê«·„’—Ê›«  ≈·Ï «·√—»«Õ «·„Õ Ã“…)", 227, 3629, 6804, 340, 11, True, CLR_TEXT, "", 0)
    Set c = AddCombo("cboYear", "", 227, 4281, 1928, 482, "", 1, "1701")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblYear", "«·”‰…", 227, 3969, 1928, 284, 9, False, CLR_MUTED, "cboYear", 0)
    Set c = AddLabel("lblYearInfo", " ", 2268, 4309, 9979, 425, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnCloseYear", "≈ﬁ›«· «·”‰…", 227, 4905, 2608, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnReopenYear", "≈⁄«œ… › Õ «·”‰…", 2948, 4905, 2948, 510, "danger")
    c.OnClick = EP
    Set c = AddLabel("lblHistoryCap", "”Ã· «·≈ﬁ›«· Ê≈⁄«œ… «·› Õ", 227, 5642, 6804, 340, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstHistory", 227, 6010, 12020, 2608, 7, "0;1588;1474;794;1928;2041;3402", True)
    c.RowSource = Tr("SELECT p.PeriodClosingID, IIf(p.ActionType = 'CLOSE', '≈ﬁ›«· › —…', IIf(p.ActionType = 'REOPEN', '≈⁄«œ… › Õ', IIf(p.ActionType = 'YEAR_CLOSE', '≈ﬁ›«· ”‰…', '≈⁄«œ… › Õ ”‰…'))) AS [«·⁄„·Ì…], p.ClosedThrough AS [„ﬁ›·… Õ Ï], p.FiscalYear AS [«·”‰…], e.EmployeeName AS [»Ê«”ÿ…], p.CreatedAt AS [›Ì], p.Notes AS [«·”»»] FROM PeriodClosings AS p INNER JOIN Employees AS e ON p.EmployeeID = e.EmployeeID ORDER BY p.PeriodClosingID DESC")
    Set c = AddButton("btnClose", "≈€·«ﬁ", 10773, 8845, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PeriodClosingLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboYear_AfterUpdate()" & vbCrLf
    s = s & "    PeriodClosingRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClosePeriod_Click()" & vbCrLf
    s = s & "    DoClosePeriod Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReopenPeriod_Click()" & vbCrLf
    s = s & "    DoReopenPeriod Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCloseYear_Click()" & vbCrLf
    s = s & "    DoCloseYear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReopenYear_Click()" & vbCrLf
    s = s & "    DoReopenYear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPeriodClosing", s
    Exit Sub
EH:
    AbortForm "frmPeriodClosing", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmVatReturn()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmVatReturn", "≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…", "", 14742, 11000, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 14742, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "Œ«‰«  ‰„Ê–Ã ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ „‰ «·„” ‰œ« ° À„ «·«⁄ „«œ ÊﬁÌœ «· ”ÊÌ… Ê«·”œ«œ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtFrom", "", 227, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "»œ«Ì… «·› —… «·÷—Ì»Ì…", 227, 992, 1701, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 2041, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "‰Â«Ì… «·› —…", 2041, 992, 1701, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnLastMonth", "«·‘Â— «·„«÷Ì", 3856, 1304, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastQuarter", "«·—»⁄ «·„«÷Ì", 5443, 1304, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCalc", "«Õ”»", 7031, 1304, 1361, 454, "primary")
    c.OnClick = EP
    Set c = AddText("txtReturnID", "", 8505, 1304, 567, 454)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblState", " ", 227, 1843, 14288, 567, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstBoxes", 227, 2466, 14288, 3912, 5, "680;7598;1928;1928;1928", True)
    c.RowSourceType = "Value List"
    Set c = AddText("txtCorrections", "", 227, 6804, 1928, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblCorrections", "14-  ’ÕÌÕ«  ”«»ﬁ… (+/-)", 227, 6492, 1928, 284, 9, False, CLR_MUTED, "txtCorrections", 0)
    Set c = AddText("txtCarried", "", 2268, 6804, 1928, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblCarried", "15- —’Ìœ œ«∆‰ „—ÕÛ¯·", 2268, 6492, 1928, 284, 9, False, CLR_MUTED, "txtCarried", 0)
    Set c = AddLabel("lblNetDue", " ", 4309, 6804, 5670, 454, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnSaveDraft", "Õ›Ÿ „”Êœ…", 10093, 6804, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 11680, 6804, 1247, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDeleteDraft", "Õ–› «·„”Êœ…", 13041, 6804, 1474, 454, "danger")
    c.OnClick = EP
    Set c = AddText("txtFiledDate", "", 227, 7683, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFiledDate", " «—ÌŒ «·«⁄ „«œ", 227, 7371, 1701, 284, 9, False, CLR_MUTED, "txtFiledDate", 0)
    Set c = AddText("txtFilingRef", "", 2041, 7683, 2041, 454)
    Set c = AddLabel("lblFilingRef", "—ﬁ„ «·≈ﬁ—«— ·œÏ «·ÂÌ∆…", 2041, 7371, 2041, 284, 9, False, CLR_MUTED, "txtFilingRef", 0)
    Set c = AddText("txtNotes", "", 4196, 7683, 2948, 454)
    Set c = AddLabel("lblNotes", "”»» ≈·€«¡ «·«⁄ „«œ", 4196, 7371, 2948, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnFile", "«⁄ „«œ «·≈ﬁ—«—", 7258, 7683, 1701, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUnfile", "≈·€«¡ «·«⁄ „«œ", 9072, 7683, 1588, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnEntry", "ﬁÌœ «· ”ÊÌ…", 10773, 7683, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtPaidDate", "", 227, 8562, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblPaidDate", " «—ÌŒ «·”œ«œ", 227, 8250, 1701, 284, 9, False, CLR_MUTED, "txtPaidDate", 0)
    Set c = AddText("txtPaidAmount", "", 2041, 8562, 1701, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblPaidAmount", "«·„»·€ «·„”œœ", 2041, 8250, 1701, 284, 9, False, CLR_MUTED, "txtPaidAmount", 0)
    Set c = AddCombo("cboPayAccount", "", 3856, 8562, 3289, 454, "", 2, "0;3175")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblPayAccount", "”ıœˆ¯œ  „‰ Õ”«»", 3856, 8250, 3289, 284, 9, False, CLR_MUTED, "cboPayAccount", 0)
    Set c = AddButton("btnPay", " ”ÃÌ· «·”œ«œ", 7258, 8562, 1701, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUnpay", "≈·€«¡ «·”œ«œ", 9072, 8562, 1588, 454, "danger")
    c.OnClick = EP
    Set c = AddLabel("lblHistoryCap", "«·≈ﬁ—«—«  «·„Õ›ÊŸ… («Œ — ≈ﬁ—«—« ·⁄—÷Â)", 227, 9129, 6804, 312, 9, True, CLR_MUTED, "", 0)
    Set c = AddList("lstReturns", 227, 9469, 12814, 1361, 8, "0;1814;1474;1474;1134;1701;1701;1474", True)
    c.RowSource = Tr("SELECT VatReturnID, ReturnNumber AS [«·≈ﬁ—«—], Format(PeriodFrom, 'yyyy/mm/dd') AS [„‰], Format(PeriodTo, 'yyyy/mm/dd') AS [≈·Ï], IIf(Status = 'FILED', '„⁄ „œ', '„”Êœ…') AS [«·Õ«·…], Format(NetDue, '#,##0.00') AS [«·’«›Ì], Format(PaidAmount, '#,##0.00') AS [«·„”œœ], Format(FiledDate, 'yyyy/mm/dd') AS [«⁄ ı„œ ›Ì] FROM VatReturns ORDER BY PeriodFrom DESC")
    c.AfterUpdate = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13041, 10263, 1474, 567, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    VatReturnLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtCorrections_AfterUpdate()" & vbCrLf
    s = s & "    VatShowNet Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtCarried_AfterUpdate()" & vbCrLf
    s = s & "    VatShowNet Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstReturns_AfterUpdate()" & vbCrLf
    s = s & "    VatPickReturn Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastMonth_Click()" & vbCrLf
    s = s & "    VatQuickPeriod Me, ""LASTMONTH""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnLastQuarter_Click()" & vbCrLf
    s = s & "    VatQuickPeriod Me, ""LASTQUARTER""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCalc_Click()" & vbCrLf
    s = s & "    VatCalculate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveDraft_Click()" & vbCrLf
    s = s & "    VatSaveDraft Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintVatReturn Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDeleteDraft_Click()" & vbCrLf
    s = s & "    VatDeleteDraft Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFile_Click()" & vbCrLf
    s = s & "    VatFile Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnfile_Click()" & vbCrLf
    s = s & "    VatUnfile Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnEntry_Click()" & vbCrLf
    s = s & "    VatOpenEntry Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPay_Click()" & vbCrLf
    s = s & "    VatPay Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnpay_Click()" & vbCrLf
    s = s & "    VatUnpay Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmVatReturn", s
    Exit Sub
EH:
    AbortForm "frmVatReturn", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAging()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAging", "√⁄„«— «·œÌÊ‰", "", 15309, 10433, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "√⁄„«— «·œÌÊ‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«·„ »ﬁÌ „‰ ﬂ· ›« Ê—… ¬Ã·… Õ”»  √ŒÌ—Â« ⁄‰  «—ÌŒ «·«” Õﬁ«ﬁ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboKind", "", 227, 1304, 1814, 454, "C;«·⁄„·«¡;S;«·„Ê—œÊ‰", 2, "0;1701")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblKind", "«·‰Ê⁄", 227, 992, 1814, 284, 9, False, CLR_MUTED, "cboKind", 0)
    Set c = AddText("txtAsOf", "", 2155, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblAsOf", "›Ì ÌÊ„", 2155, 992, 1701, 284, 9, False, CLR_MUTED, "txtAsOf", 0)
    Set c = AddButton("btnShow", "⁄—÷", 3969, 1304, 1361, 454, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblTotals", " ", 227, 1871, 14855, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstParties", 227, 2268, 14855, 3629, 10, "0;3402;1474;1474;1361;1361;1361;1361;1361;1134", True)
    c.AfterUpdate = EP
    c.OnDblClick = EP
    Set c = AddLabel("lblInfo", " ", 227, 5954, 14855, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstDocs", 227, 6350, 14855, 2722, 7, "0;1701;1814;1474;1474;1361;1701", True)
    Set c = AddButton("btnPrint", "ÿ»«⁄…", 227, 9412, 1361, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnAllocate", "—»ÿ «·”œ«œ »«·›Ê« Ì—", 1701, 9412, 2268, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnStatement", "ﬂ‘› Õ”«»", 4082, 9412, 1588, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 9412, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    AgingLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstParties_AfterUpdate()" & vbCrLf
    s = s & "    AgingPartyChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstParties_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    AgingOpenAllocation Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboKind_AfterUpdate()" & vbCrLf
    s = s & "    AgingRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnShow_Click()" & vbCrLf
    s = s & "    AgingRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintAging Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAllocate_Click()" & vbCrLf
    s = s & "    AgingOpenAllocation Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStatement_Click()" & vbCrLf
    s = s & "    AgingStatement Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAging", s
    Exit Sub
EH:
    AbortForm "frmAging", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAllocation()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAllocation", "—»ÿ «·”œ«œ »«·›Ê« Ì—", "", 14742, 10433, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 14742, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "—»ÿ «·”œ«œ »«·›Ê« Ì—", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "«Œ — «·”‰œ À„ «·›« Ê—… «· Ì Ì”œœÂ«. „« ·« Ìı—»ÿ Ì”œœ √ﬁœ„ «·›Ê« Ì— «” Õﬁ«ﬁ«", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboKind", "", 227, 1304, 1814, 454, "C;«·⁄„·«¡;S;«·„Ê—œÊ‰", 2, "0;1701")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblKind", "«·‰Ê⁄", 227, 992, 1814, 284, 9, False, CLR_MUTED, "cboKind", 0)
    Set c = AddCombo("cboParty", "", 2155, 1304, 3969, 454, "", 2, "0;3856")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblParty", "«·⁄„Ì· / «·„Ê—œ", 2155, 992, 3969, 284, 9, False, CLR_MUTED, "cboParty", 0)
    Set c = AddButton("btnAuto", "—»ÿ  ·ﬁ«∆Ì »«·√ﬁœ„", 6237, 1304, 2041, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblPayCap", "«·”‰œ« ", 227, 1928, 7087, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstPayments", 227, 2268, 7087, 3629, 6, "0;1361;1247;1247;1247;1247", True)
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvCap", "«·›Ê« Ì— «·„› ÊÕ…", 7427, 1928, 7087, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstInvoices", 7427, 2268, 7087, 3629, 7, "0;1247;1077;1077;1077;1134;1077", True)
    c.AfterUpdate = EP
    Set c = AddText("txtAmount", "", 227, 6350, 1814, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€ «·„—»Êÿ", 227, 6038, 1814, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddButton("btnAllocate", "—»ÿ »«·›« Ê—…", 2155, 6350, 1814, 454, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblAllocCap", "«·—»ÿ «·„”Ã·", 227, 6974, 6804, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstAllocations", 227, 7314, 14288, 2041, 5, "0;1701;1701;1701;1701", True)
    Set c = AddButton("btnRemove", "≈·€«¡ «·—»ÿ", 227, 9526, 1701, 510, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13041, 9526, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    AllocationLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboParty_AfterUpdate()" & vbCrLf
    s = s & "    AllocationRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstPayments_AfterUpdate()" & vbCrLf
    s = s & "    AllocationPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstInvoices_AfterUpdate()" & vbCrLf
    s = s & "    AllocationPicked Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboKind_AfterUpdate()" & vbCrLf
    s = s & "    AllocationKindChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAuto_Click()" & vbCrLf
    s = s & "    DoAutoAllocate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAllocate_Click()" & vbCrLf
    s = s & "    DoAllocate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRemove_Click()" & vbCrLf
    s = s & "    DoRemoveAllocation Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAllocation", s
    Exit Sub
EH:
    AbortForm "frmAllocation", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBankTx()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBankTx", "«·Õ—ﬂ«  «·»‰ﬂÌ…", "", 14742, 10206, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 14742, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·Õ—ﬂ«  «·»‰ﬂÌ…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "≈Ìœ«⁄ Ê”Õ»°  ”ÊÌ…  Õ’Ì·«  „œÏ »⁄„Ê· Â«° «· ÕÊÌ· »Ì‰ «·»‰Êﬂ° Ê«·Õ—ﬂ«  «·√Œ—Ï", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboType", "", 227, 1304, 3175, 454, "DEPOSIT;≈Ìœ«⁄ ‰ﬁœÌ… „‰ ’‰œÊﬁ;WITHDRAW;”Õ» ‰ﬁœÌ… ≈·Ï ’‰œÊﬁ;SETTLEMENT; ”ÊÌ…  Õ’Ì·«  „œÏ;TRANSFER; ÕÊÌ· ≈·Ï »‰ﬂ ¬Œ—;OTHER_IN;Ê«—œ ¬Œ— (›Ê«∆œ √Ê ﬁ—÷);OTHER_OUT;’«œ— ¬Œ— (—”Ê„ »‰ﬂÌ… Ê€Ì—Â«)", 2, "0;3062")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblType", "‰Ê⁄ «·Õ—ﬂ…", 227, 992, 3175, 284, 9, False, CLR_MUTED, "cboType", 0)
    Set c = AddText("txtDate", "", 3515, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblDate", "«· «—ÌŒ", 3515, 992, 1701, 284, 9, False, CLR_MUTED, "txtDate", 0)
    Set c = AddCombo("cboBank", "", 5330, 1304, 2948, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBank", "«·»‰ﬂ", 5330, 992, 2948, 284, 9, False, CLR_MUTED, "cboBank", 0)
    Set c = AddCombo("cboToBank", "", 8392, 1304, 2948, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblToBank", "≈·Ï »‰ﬂ («· ÕÊÌ·)", 8392, 992, 2948, 284, 9, False, CLR_MUTED, "cboToBank", 0)
    Set c = AddCombo("cboBox", "", 11453, 1304, 3062, 454, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBox", "«·’‰œÊﬁ (≈Ìœ«⁄ / ”Õ»)", 11453, 992, 3062, 284, 9, False, CLR_MUTED, "cboBox", 0)
    Set c = AddText("txtAmount", "", 227, 2183, 1814, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblAmount", "«·„»·€", 227, 1871, 1814, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddText("txtFee", "", 2155, 2183, 1588, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblFee", "⁄„Ê·… „œÏ", 2155, 1871, 1588, 284, 9, False, CLR_MUTED, "txtFee", 0)
    Set c = AddText("txtFeeVAT", "", 3856, 2183, 1588, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblFeeVAT", "÷—Ì»… «·⁄„Ê·… / «·—”Ê„", 3856, 1871, 1588, 284, 9, False, CLR_MUTED, "txtFeeVAT", 0)
    Set c = AddCombo("cboCounter", "", 5557, 2183, 3969, 454, "SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE IsPosting = True AND IsActive = True AND Nz(Level3Code, 0) NOT IN (1100, 1210) AND AccountCode NOT IN (1190, 1200, 1300, 2100) ORDER BY TreeKey", 2, "0;3856")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblCounter", "«·Õ”«» «·„ﬁ«»· («·Õ—ﬂ«  «·√Œ—Ï)", 5557, 1871, 3969, 284, 9, False, CLR_MUTED, "cboCounter", 0)
    Set c = AddText("txtReference", "", 9639, 2183, 1814, 454)
    Set c = AddLabel("lblReference", "„—Ã⁄ «·»‰ﬂ", 9639, 1871, 1814, 284, 9, False, CLR_MUTED, "txtReference", 0)
    Set c = AddText("txtDescription", "", 11567, 2183, 2948, 454)
    Set c = AddLabel("lblDescription", "«·»Ì«‰", 11567, 1871, 2948, 284, 9, False, CLR_MUTED, "txtDescription", 0)
    Set c = AddLabel("lblInfo", " ", 227, 2778, 11113, 567, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «·Õ—ﬂ…", 11567, 2835, 2948, 510, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblListCap", "«·Õ—ﬂ«  «·„”Ã·…", 227, 3515, 6804, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstTx", 227, 3856, 14288, 5443, 7, "0;1474;1361;1361;2268;1701;5670", True)
    c.RowSource = Tr("SELECT t.BankTxID, t.TxNumber AS [«·—ﬁ„], Format(t.TxDate, 'yyyy/mm/dd') AS [«· «—ÌŒ], IIf(t.TxType = 'DEPOSIT', '≈Ìœ«⁄', IIf(t.TxType = 'WITHDRAW', '”Õ»', IIf(t.TxType = 'SETTLEMENT', ' ”ÊÌ… „œÏ', IIf(t.TxType = 'TRANSFER', ' ÕÊÌ·', IIf(t.TxType = 'OTHER_IN', 'Ê«—œ', '’«œ—'))))) AS [«·‰Ê⁄], k.BankName AS [«·»‰ﬂ], Format(t.Amount, '#,##0.00') AS [«·„»·€], t.Description AS [«·»Ì«‰] FROM BankTransactions AS t INNER JOIN Banks AS k ON t.BankID = k.BankID ORDER BY t.TxDate DESC, t.BankTxID DESC")
    Set c = AddButton("btnDelete", "Õ–› «·Õ—ﬂ…", 227, 9469, 1701, 510, "danger")
    c.OnClick = EP
    Set c = AddButton("btnRecon", "«· ”ÊÌ… «·»‰ﬂÌ…", 2041, 9469, 1928, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13041, 9469, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    BankTxLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboType_AfterUpdate()" & vbCrLf
    s = s & "    BankTxTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBank_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboToBank_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBox_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtAmount_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtFee_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtFeeVAT_AfterUpdate()" & vbCrLf
    s = s & "    BankTxRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveBankTx Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    DeleteSelectedBankTx Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRecon_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBankRecon"", 0, Me!cboBank.Value" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBankTx", s
    Exit Sub
EH:
    AbortForm "frmBankTx", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBankRecon()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBankRecon", "«· ”ÊÌ… «·»‰ﬂÌ…", "", 15309, 10886, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«· ”ÊÌ… «·»‰ﬂÌ…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„ÿ«»ﬁ… ﬂ‘› «·»‰ﬂ „⁄ «·œ›« —: ⁄·ˆ¯„ «·⁄„·Ì«  «·Ÿ«Â—… ›Ì «·ﬂ‘›", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboBank", "", 227, 1304, 3062, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBank", "«·»‰ﬂ", 227, 992, 3062, 284, 9, False, CLR_MUTED, "cboBank", 0)
    Set c = AddText("txtStatementDate", "", 3402, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblStatementDate", " «—ÌŒ ﬂ‘› «·»‰ﬂ", 3402, 992, 1701, 284, 9, False, CLR_MUTED, "txtStatementDate", 0)
    Set c = AddText("txtStatementBalance", "", 5216, 1304, 1928, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblStatementBalance", "—’Ìœ «·ﬂ‘› ›Ì Â–« «· «—ÌŒ", 5216, 992, 1928, 284, 9, False, CLR_MUTED, "txtStatementBalance", 0)
    Set c = AddButton("btnStart", "»œ¡ /  ÕœÌÀ «· ”ÊÌ…", 7258, 1304, 2381, 454, "primary")
    c.OnClick = EP
    Set c = AddText("txtReconID", "", 9752, 1304, 340, 454)
    SetCtlProp c, "Visible", False
    Set c = AddRect("boxBook", 227, 1899, 3600, 964, CLR_SURFACE)
    Set c = AddLabel("lblBookCap", "«·—’Ìœ ›Ì «·œ›« —", 369, 1984, 3316, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblBook", "-", 369, 2296, 3316, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxOutstanding", 3940, 1899, 3600, 964, CLR_SURFACE)
    Set c = AddLabel("lblOutstandingCap", "⁄„·Ì«  ·„  ŸÂ— ›Ì «·ﬂ‘›", 4082, 1984, 3316, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblOutstanding", "-", 4082, 2296, 3316, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxAdjusted", 7653, 1899, 3600, 964, CLR_SURFACE)
    Set c = AddLabel("lblAdjustedCap", "«·œ›« — »⁄œ «” »⁄«œÂ«", 7795, 1984, 3316, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblAdjusted", "-", 7795, 2296, 3316, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxDifference", 11366, 1899, 3600, 964, CLR_SURFACE)
    Set c = AddLabel("lblDifferenceCap", "«·›—ﬁ „⁄ «·ﬂ‘›", 11508, 1984, 3316, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblDifference", "-", 11508, 2296, 3316, 482, 14, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblState", " ", 227, 2920, 14855, 539, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblOpenCap", "⁄„·Ì«  «·œ›« — €Ì— «·„ÿ«»ﬁ… (Õ Ï  «—ÌŒ «·ﬂ‘›)", 227, 3515, 7371, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstOpen", 227, 3856, 7371, 4196, 7, "0;1134;1247;1247;1588;1021;1021", True)
    SetCtlProp c, "MultiSelect", 2
    c.OnDblClick = EP
    Set c = AddLabel("lblClearedCap", "«·„ÿ«»ﬁ… ›Ì Â–Â «· ”ÊÌ…", 7711, 3515, 7371, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddList("lstCleared", 7711, 3856, 7371, 4196, 6, "0;1247;1474;1474;1361;1361", True)
    SetCtlProp c, "MultiSelect", 2
    c.OnDblClick = EP
    Set c = AddButton("btnClear", "„ÿ«»ﬁ… «·„Õœœ", 227, 8165, 1928, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnAddTx", "Õ—ﬂ… »‰ﬂÌ… ÃœÌœ…", 2268, 8165, 2155, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnRefresh", " ÕœÌÀ", 4536, 8165, 1247, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnUnclear", "≈·€«¡ «·„ÿ«»ﬁ…", 7711, 8165, 1928, 510, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblReconsCap", " ”ÊÌ«  Â–« «·»‰ﬂ", 227, 8788, 6804, 312, 9, True, CLR_MUTED, "", 0)
    Set c = AddList("lstRecons", 227, 9129, 7371, 1531, 5, "0;1588;1588;1928;1361", True)
    c.AfterUpdate = EP
    Set c = AddButton("btnFinish", "«⁄ „«œ «· ”ÊÌ…", 7711, 9129, 1928, 510, "primary")
    c.OnClick = EP
    Set c = AddButton("btnReopen", "≈⁄«œ… › Õ", 9752, 9129, 1474, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDeleteRecon", "Õ–› «·Ã«—Ì…", 11339, 9129, 1588, 510, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 10149, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    BankReconLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBank_AfterUpdate()" & vbCrLf
    s = s & "    BankReconBankChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstRecons_AfterUpdate()" & vbCrLf
    s = s & "    BankReconPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstOpen_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    BankReconClear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstCleared_DblClick(Cancel As Integer)" & vbCrLf
    s = s & "    BankReconUnclear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnStart_Click()" & vbCrLf
    s = s & "    BankReconStart Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    BankReconClear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAddTx_Click()" & vbCrLf
    s = s & "    BankReconAddTx Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    BankReconRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnclear_Click()" & vbCrLf
    s = s & "    BankReconUnclear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFinish_Click()" & vbCrLf
    s = s & "    BankReconFinish Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReopen_Click()" & vbCrLf
    s = s & "    BankReconReopen Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDeleteRecon_Click()" & vbCrLf
    s = s & "    BankReconDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBankRecon", s
    Exit Sub
EH:
    AbortForm "frmBankRecon", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCheques()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCheques", "«·‘Ìﬂ« ", "", 15309, 10546, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE825), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·‘Ìﬂ«  «·Ê«—œ… Ê«·’«œ—…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", " ”ÃÌ· «·‘Ìﬂ Ì”œœ —’Ìœ «·⁄„Ì· √Ê «·„Ê—œ° À„ ÌıÕ’Û¯· ›Ì «·»‰ﬂ √Ê Ì— œ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboDirection", "", 227, 1304, 3402, 454, "IN;‘Ìﬂ«  Ê«—œ… („‰ «·⁄„·«¡);OUT;‘Ìﬂ«  ’«œ—… (··„Ê—œÌ‰)", 2, "0;3289")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblDirection", "«·‰Ê⁄", 227, 992, 3402, 284, 9, False, CLR_MUTED, "cboDirection", 0)
    Set c = AddCombo("cboShow", "", 3742, 1304, 3629, 454, "PENDING; Õ  «· Õ’Ì·;DUE;„” Õﬁ… Œ·«· 7 √Ì«„ √Ê ›«  «” Õﬁ«ﬁÂ«;COLLECTED;«·„Õ’Û¯·… / «·„’—Ê›…;BOUNCED;«·„— œ…;ALL;«·ﬂ·", 2, "0;3515")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblShow", "⁄—÷", 3742, 992, 3629, 284, 9, False, CLR_MUTED, "cboShow", 0)
    Set c = AddLabel("lblTotals", " ", 7484, 1304, 7598, 454, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblNewCap", " ”ÃÌ· ‘Ìﬂ ÃœÌœ", 227, 1928, 6804, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddCombo("cboParty", "", 227, 2580, 3402, 454, "", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblParty", "«·⁄„Ì·", 227, 2268, 3402, 284, 9, False, CLR_MUTED, "cboParty", 0)
    Set c = AddText("txtChequeNo", "", 3742, 2580, 1701, 454)
    Set c = AddLabel("lblChequeNo", "—ﬁ„ «·‘Ìﬂ", 3742, 2268, 1701, 284, 9, False, CLR_MUTED, "txtChequeNo", 0)
    Set c = AddText("txtDrawerBank", "", 5557, 2580, 2268, 454)
    Set c = AddLabel("lblDrawerBank", "»‰ﬂ «·”«Õ» («·Ê«—œ)", 5557, 2268, 2268, 284, 9, False, CLR_MUTED, "txtDrawerBank", 0)
    Set c = AddCombo("cboBank", "", 7938, 2580, 2495, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblBank", "»‰ﬂ‰« («·’«œ—: «·„”ÕÊ» ⁄·ÌÂ)", 7938, 2268, 2495, 284, 9, False, CLR_MUTED, "cboBank", 0)
    Set c = AddText("txtIssueDate", "", 10546, 2580, 1474, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblIssueDate", " «—ÌŒ «·‘Ìﬂ", 10546, 2268, 1474, 284, 9, False, CLR_MUTED, "txtIssueDate", 0)
    Set c = AddText("txtDueDate", "", 12134, 2580, 1474, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblDueDate", "«·«” Õﬁ«ﬁ", 12134, 2268, 1474, 284, 9, False, CLR_MUTED, "txtDueDate", 0)
    Set c = AddText("txtAmount", "", 13721, 2580, 1361, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "«·„»·€", 13721, 2268, 1361, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddText("txtNotes", "", 227, 3402, 7598, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 227, 3090, 7598, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnSave", " ”ÃÌ· «·‘Ìﬂ", 7938, 3402, 2495, 454, "primary")
    c.OnClick = EP
    Set c = AddList("lstCheques", 227, 4082, 14855, 4649, 9, "0;1361;1588;3062;1361;1474;1361;1361;2608", True)
    Set c = AddText("txtActionDate", "", 227, 9242, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblActionDate", " «—ÌŒ «·⁄„·Ì…", 227, 8930, 1588, 284, 9, False, CLR_MUTED, "txtActionDate", 0)
    Set c = AddCombo("cboActionBank", "", 1928, 9242, 2495, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblActionBank", "«·»‰ﬂ («· Õ’Ì· / «·’—›)", 1928, 8930, 2495, 284, 9, False, CLR_MUTED, "cboActionBank", 0)
    Set c = AddButton("btnCollect", " Õ’Ì· ›Ì «·»‰ﬂ", 4536, 9242, 1928, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnBounce", "«— œ«œ", 6577, 9242, 1247, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "≈·€«¡ «·Õ«·…", 7937, 9242, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 9638, 9242, 1021, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 9866, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ChequesLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboDirection_AfterUpdate()" & vbCrLf
    s = s & "    ChequesDirectionChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboShow_AfterUpdate()" & vbCrLf
    s = s & "    ChequesRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    SaveCheque Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCollect_Click()" & vbCrLf
    s = s & "    CollectCheque Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnBounce_Click()" & vbCrLf
    s = s & "    BounceCheque Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    UndoCheque Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    DeleteSelectedCheque Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCheques", s
    Exit Sub
EH:
    AbortForm "frmCheques", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAssets()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAssets", "«·√’Ê· «·À«» …", "", 15309, 11000, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·√’Ê· «·À«» …", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "”Ã· «·√’Ê· ÊﬁÌœ ‘—«∆Â«° Ê«·≈Â·«ﬂ »«·ﬁ”ÿ «·À«» ° Ê«·»Ì⁄ √Ê «·«” »⁄«œ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtAssetName", "", 227, 1304, 3629, 454)
    Set c = AddLabel("lblAssetName", "«”„ «·√’·", 227, 992, 3629, 284, 9, False, CLR_MUTED, "txtAssetName", 0)
    Set c = AddCombo("cboAssetAccount", "", 3969, 1304, 3062, 454, "SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE IsPosting = True AND IsActive = True AND AccountType = 'ASSET' AND Level2Code = 12 AND AccountCode <> 1790 ORDER BY TreeKey", 2, "0;2948")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblAssetAccount", "Õ”«» «·√’· («·„Ã„Ê⁄…)", 3969, 992, 3062, 284, 9, False, CLR_MUTED, "cboAssetAccount", 0)
    Set c = AddText("txtPurchaseDate", "", 7144, 1304, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblPurchaseDate", " «—ÌŒ «·‘—«¡", 7144, 992, 1588, 284, 9, False, CLR_MUTED, "txtPurchaseDate", 0)
    Set c = AddText("txtCost", "", 8845, 1304, 1701, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblCost", "«· ﬂ·›… »œÊ‰ «·÷—Ì»…", 8845, 992, 1701, 284, 9, False, CLR_MUTED, "txtCost", 0)
    Set c = AddText("txtInputVAT", "", 10660, 1304, 1474, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblInputVAT", "÷—Ì»… «·„œŒ·« ", 10660, 992, 1474, 284, 9, False, CLR_MUTED, "txtInputVAT", 0)
    Set c = AddText("txtSalvage", "", 12247, 1304, 1474, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblSalvage", "«·ﬁÌ„… «·„ »ﬁÌ…", 12247, 992, 1474, 284, 9, False, CLR_MUTED, "txtSalvage", 0)
    Set c = AddText("txtLife", "", 13835, 1304, 1247, 454)
    SetCtlProp c, "Format", "0"
    c.AfterUpdate = EP
    Set c = AddLabel("lblLife", "«·⁄„— (‘Â—)", 13835, 992, 1247, 284, 9, False, CLR_MUTED, "txtLife", 0)
    Set c = AddText("txtAssetID", "", 15139, 907, 113, 227)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtDepStart", "", 227, 2126, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblDepStart", "»œ«Ì… «·≈Â·«ﬂ", 227, 1814, 1588, 284, 9, False, CLR_MUTED, "txtDepStart", 0)
    Set c = AddCombo("cboSource", "", 1928, 2126, 3175, 454, "BANK;„‰ «·»‰ﬂ;CASHBOX;„‰ ’‰œÊﬁ;ACCOUNT;⁄·Ï Õ”«» ¬Œ— („” Õﬁ«  √Ê ﬁ—÷...);OPENING;„ÊÃÊœ ﬁ»· «·»—‰«„Ã (—’Ìœ «›  «ÕÌ)", 2, "0;3062")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblSource", "„’œ— «·‘—«¡", 1928, 1814, 3175, 284, 9, False, CLR_MUTED, "cboSource", 0)
    Set c = AddCombo("cboBank", "", 5216, 2126, 2041, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblBank", "«·»‰ﬂ", 5216, 1814, 2041, 284, 9, False, CLR_MUTED, "cboBank", 0)
    Set c = AddCombo("cboBox", "", 7371, 2126, 1928, 454, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblBox", "«·’‰œÊﬁ", 7371, 1814, 1928, 284, 9, False, CLR_MUTED, "cboBox", 0)
    Set c = AddCombo("cboCounter", "", 9412, 2126, 3175, 454, "SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE IsPosting = True AND IsActive = True AND Nz(Level3Code, 0) NOT IN (1100, 1210) AND AccountCode NOT IN (1190, 1200, 1300, 2100) ORDER BY TreeKey", 2, "0;3062")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblCounter", "«·Õ”«» «·œ«∆‰", 9412, 1814, 3175, 284, 9, False, CLR_MUTED, "cboCounter", 0)
    Set c = AddText("txtOpeningAccum", "", 12701, 2126, 2381, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblOpeningAccum", "≈Â·«ﬂ ”«»ﬁ", 12701, 1814, 2381, 284, 9, False, CLR_MUTED, "txtOpeningAccum", 0)
    Set c = AddText("txtNotes", "", 227, 2948, 4196, 454)
    Set c = AddLabel("lblNotes", "„·«ÕŸ« ", 227, 2636, 4196, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddCombo("cboCenter", "", 4536, 2948, 1928, 454, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblCenter", "„—ﬂ“ «· ﬂ·›…", 4536, 2636, 1928, 284, 9, False, CLR_MUTED, "cboCenter", 0)
    Set c = AddButton("btnSave", "Õ›Ÿ «·√’·", 6577, 2948, 1701, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnNew", "√’· ÃœÌœ", 8391, 2948, 1474, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–›", 9978, 2948, 1021, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnDepreciation", "«·≈Â·«ﬂ «·‘Â—Ì", 11112, 2948, 1814, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·”Ã·", 13039, 2948, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblAssetInfo", " ", 227, 3515, 14855, 567, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstAssets", 227, 4139, 14855, 4536, 9, "0;1247;3402;2495;1361;1588;1588;1701;1021", True)
    c.RowSource = Tr("SELECT AssetID, AssetCode AS [«·—ﬁ„], AssetName AS [«·√’·], AssetGroup AS [«·„Ã„Ê⁄…], Format(PurchaseDate, 'yyyy/mm/dd') AS [«·‘—«¡], Format(q.Cost, '#,##0.00') AS [«· ﬂ·›…], Format(AccumDep, '#,##0.00') AS [„Ã„⁄ «·≈Â·«ﬂ], Format(BookValue, '#,##0.00') AS [«·ﬁÌ„… «·œ› —Ì…], q.StatusName AS [«·Õ«·…] FROM FixedAssetsQuery AS q ORDER BY q.Status, q.AssetCode")
    c.AfterUpdate = EP
    Set c = AddLabel("lblDisposeCap", "»Ì⁄ «·√’· «·„⁄—Ê÷ √Ê «” »⁄«œÂ", 227, 8788, 6804, 312, 10, True, CLR_TEXT, "", 0)
    Set c = AddText("txtDisposalDate", "", 227, 9497, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblDisposalDate", "«· «—ÌŒ", 227, 9185, 1588, 284, 9, False, CLR_MUTED, "txtDisposalDate", 0)
    Set c = AddCombo("cboDisposalTo", "", 1928, 9497, 3175, 454, "BANK;»Ì⁄ - «·À„‰ ›Ì «·»‰ﬂ;CASHBOX;»Ì⁄ - «·À„‰ ›Ì ’‰œÊﬁ;NONE;«” »⁄«œ »œÊ‰ À„‰ ( ·› √Ê ›ﬁœ)", 2, "0;3062")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblDisposalTo", "«·ÿ—Ìﬁ…", 1928, 9185, 3175, 284, 9, False, CLR_MUTED, "cboDisposalTo", 0)
    Set c = AddText("txtProceeds", "", 5216, 9497, 1474, 454)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblProceeds", "À„‰ «·»Ì⁄", 5216, 9185, 1474, 284, 9, False, CLR_MUTED, "txtProceeds", 0)
    Set c = AddCombo("cboDisposalBank", "", 6804, 9497, 1928, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblDisposalBank", "«·»‰ﬂ", 6804, 9185, 1928, 284, 9, False, CLR_MUTED, "cboDisposalBank", 0)
    Set c = AddCombo("cboDisposalBox", "", 8845, 9497, 1814, 454, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblDisposalBox", "«·’‰œÊﬁ", 8845, 9185, 1814, 284, 9, False, CLR_MUTED, "cboDisposalBox", 0)
    Set c = AddButton("btnDispose", "»Ì⁄ / «” »⁄«œ", 10773, 9497, 1701, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnUndoDispose", "≈·€«¡ «·«” »⁄«œ", 12587, 9497, 1814, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 10206, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    AssetsLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboSource_AfterUpdate()" & vbCrLf
    s = s & "    AssetSourceChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstAssets_AfterUpdate()" & vbCrLf
    s = s & "    AssetPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtCost_AfterUpdate()" & vbCrLf
    s = s & "    AssetRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSalvage_AfterUpdate()" & vbCrLf
    s = s & "    AssetRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtLife_AfterUpdate()" & vbCrLf
    s = s & "    AssetRefresh Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    AssetSave Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnNew_Click()" & vbCrLf
    s = s & "    AssetNew Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    AssetDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDepreciation_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmDepreciation"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintAssets Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDispose_Click()" & vbCrLf
    s = s & "    AssetDispose Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndoDispose_Click()" & vbCrLf
    s = s & "    AssetUndoDisposal Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAssets", s
    Exit Sub
EH:
    AbortForm "frmAssets", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmDepreciation()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmDepreciation", "«·≈Â·«ﬂ «·‘Â—Ì", "", 11340, 9299, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 11340, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·≈Â·«ﬂ «·‘Â—Ì", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬁÌœ ≈Â·«ﬂ ﬂ· ‘Â— »«· — Ì»: „’—Ê› «·≈Â·«ﬂ Ê„Ã„⁄ «·≈Â·«ﬂ ·ﬂ· √’·", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtThrough", "", 227, 1304, 1928, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblThrough", "Õ Ï ‘Â— (√Ì ÌÊ„ ›ÌÂ)", 227, 992, 1928, 284, 9, False, CLR_MUTED, "txtThrough", 0)
    Set c = AddButton("btnRun", " ”ÃÌ· «·≈Â·«ﬂ Õ Ï Â–« «·‘Â—", 2268, 1304, 3402, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "Õ–› ¬Œ— ‘Â—", 5783, 1304, 1701, 454, "danger")
    c.OnClick = EP
    Set c = AddButton("btnAssets", "«·√’Ê·", 7598, 1304, 1361, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblNext", " ", 227, 1928, 10886, 397, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstPreview", 227, 2381, 10886, 3062, 2, "7938;2495", True)
    c.RowSourceType = "Value List"
    Set c = AddLabel("lblRunsCap", "ﬁÌÊœ «·≈Â·«ﬂ «·„”Ã·…", 227, 5557, 6804, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstRuns", 227, 5897, 10886, 2495, 5, "0;1928;1701;2268;2268", True)
    c.RowSource = Tr("SELECT RunID, RunNumber AS [«·ﬁÌœ], Format(RunMonth, 'yyyy/mm') AS [«·‘Â—], Format(TotalAmount, '#,##0.00') AS [«·≈Â·«ﬂ], Format(CreatedAt, 'yyyy/mm/dd') AS [”ıÃˆ¯· ›Ì] FROM DepreciationRuns ORDER BY RunMonth DESC")
    Set c = AddButton("btnClose", "≈€·«ﬁ", 9639, 8562, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    DepreciationLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRun_Click()" & vbCrLf
    s = s & "    DepreciationRun Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndo_Click()" & vbCrLf
    s = s & "    DepreciationUndo Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAssets_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAssets"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmDepreciation", s
    Exit Sub
EH:
    AbortForm "frmDepreciation", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPayrollLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPayrollLines", "√”ÿ— „”Ì— «·—Ê« »", "SELECT * FROM PayrollLines WHERE PayrollRunID = 0 ORDER BY EmployeeName", 14855, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("EmployeeName", "EmployeeName", 28, 0, 2495, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("Basic", "Basic", 2551, 0, 1134, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Housing", "Housing", 3713, 0, 1134, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("OtherAllow", "OtherAllow", 4875, 0, 1134, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Overtime", "Overtime", 6037, 0, 1021, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("Additions", "Additions", 7086, 0, 1021, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("AbsenceDeduction", "AbsenceDeduction", 8135, 0, 1021, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("AdvanceDeduction", "AdvanceDeduction", 9184, 0, 1021, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("OtherDeduction", "OtherDeduction", 10233, 0, 1021, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("GosiEmployee", "GosiEmployee", 11282, 0, 1021, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("NetPay", "NetPay", 12331, 0, 1247, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Notes", "Notes", 13606, 0, 1247, 425)
    c.AfterUpdate = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    s = ""
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    AuditFormBefore Me, ""PayrollLines"", ""PayrollLineID""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    AuditFormAfter Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Overtime_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Additions_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub AbsenceDeduction_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub AdvanceDeduction_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub OtherDeduction_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Notes_AfterUpdate()" & vbCrLf
    s = s & "    PayrollLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPayrollLines", s
    Exit Sub
EH:
    AbortForm "frmPayrollLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPayroll()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPayroll", "„”Ì— «·—Ê« »", "", 15309, 10433, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "„”Ì— «·—Ê« »", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„”Ì— ﬂ· ‘Â—: «·—Ê« » Ê«·»œ·«  Ê«·≈÷«›Ì Ê«·Œ’Ê„«  Ê«· √„Ì‰« ° À„ «· —ÕÌ· Ê«·’—›", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboRun", "", 227, 1304, 2835, 454, "SELECT PayrollRunID, Format(PayMonth, 'yyyy/mm') & '  ' & IIf(Status = 'DRAFT', '„”Êœ…', IIf(PaidAmount <> 0, '„—ÕÛ¯· Ê„’—Ê›', '„—ÕÛ¯·')) FROM PayrollRuns ORDER BY PayMonth DESC", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblRun", "«·„”Ì—« ", 227, 992, 2835, 284, 9, False, CLR_MUTED, "cboRun", 0)
    Set c = AddText("txtMonth", "", 3175, 1304, 1701, 454)
    SetCtlProp c, "Format", "yyyy/mm"
    Set c = AddLabel("lblMonth", "‘Â— ÃœÌœ", 3175, 992, 1701, 284, 9, False, CLR_MUTED, "txtMonth", 0)
    Set c = AddButton("btnCreate", "≈‰‘«¡ „”Ì— «·‘Â—", 4990, 1304, 2041, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnRebuild", "≈⁄«œ… «·≈‰‘«¡", 7144, 1304, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPost", " —ÕÌ· «·„”Ì—", 8817, 1304, 1588, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUnpost", "≈·€«¡ «· —ÕÌ·", 10490, 1304, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–› «·„”Êœ…", 12163, 1304, 1474, 454, "danger")
    c.OnClick = EP
    Set c = AddText("txtRunID", "", 15139, 907, 113, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblState", " ", 227, 1899, 14855, 340, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCol1", "«·„ÊŸ›", 255, 2325, 2495, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·√”«”Ì", 2778, 2325, 1134, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«·”ﬂ‰", 3940, 2325, 1134, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "»œ·«  √Œ—Ï", 5102, 2325, 1134, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "«·≈÷«›Ì", 6264, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "„ﬂ«›¬ ", 7313, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "€Ì«»", 8362, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol8", "”·›…", 9411, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol9", "Ã“«¡« ", 10460, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol10", "«· √„Ì‰« ", 11509, 2325, 1021, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol11", "«·’«›Ì", 12558, 2325, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol12", "„·«ÕŸ« ", 13833, 2325, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmPayrollLines", 227, 2665, 14855, 5330)
    Set c = AddLabel("lblTotals", " ", 227, 8108, 14855, 369, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtPaidDate", "", 227, 8845, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblPaidDate", " «—ÌŒ «·’—›", 227, 8533, 1588, 284, 9, False, CLR_MUTED, "txtPaidDate", 0)
    Set c = AddCombo("cboPaidFrom", "", 1928, 8845, 1701, 454, "BANK;„‰ «·»‰ﬂ;CASHBOX;„‰ ’‰œÊﬁ", 2, "0;1588")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblPaidFrom", "«·’—›", 1928, 8533, 1701, 284, 9, False, CLR_MUTED, "cboPaidFrom", 0)
    Set c = AddCombo("cboBank", "", 3742, 8845, 2268, 454, "SELECT BankID, BankName FROM Banks WHERE IsActive = True ORDER BY BankName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblBank", "«·»‰ﬂ", 3742, 8533, 2268, 284, 9, False, CLR_MUTED, "cboBank", 0)
    Set c = AddCombo("cboBox", "", 6124, 8845, 2041, 454, "SELECT CashBoxID, BoxName FROM CashBoxes ORDER BY BoxType DESC, BoxName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblBox", "«·’‰œÊﬁ", 6124, 8533, 2041, 284, 9, False, CLR_MUTED, "cboBox", 0)
    Set c = AddButton("btnPay", " ”ÃÌ· «·’—›", 8278, 8845, 1701, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndoPay", "≈·€«¡ «·’—›", 10093, 8845, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·„”Ì—", 11794, 8845, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnEmployees", "—Ê« » «·„ÊŸ›Ì‰", 227, 9724, 1928, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtGosiEE", "", 2268, 9724, 1588, 454)
    SetCtlProp c, "Format", "0.00%"
    Set c = AddLabel("lblGosiEE", "Õ’… «·„ÊŸ› «·”⁄ÊœÌ", 2268, 9412, 1588, 284, 9, False, CLR_MUTED, "txtGosiEE", 0)
    Set c = AddText("txtGosiER", "", 3969, 9724, 1588, 454)
    SetCtlProp c, "Format", "0.00%"
    Set c = AddLabel("lblGosiER", "Õ’… «·„‰‘√… (”⁄ÊœÌ)", 3969, 9412, 1588, 284, 9, False, CLR_MUTED, "txtGosiER", 0)
    Set c = AddText("txtGosiNonSaudi", "", 5670, 9724, 1588, 454)
    SetCtlProp c, "Format", "0.00%"
    Set c = AddLabel("lblGosiNonSaudi", "«·„‰‘√… (€Ì— ”⁄ÊœÌ)", 5670, 9412, 1588, 284, 9, False, CLR_MUTED, "txtGosiNonSaudi", 0)
    Set c = AddText("txtGosiMax", "", 7371, 9724, 1588, 454)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddLabel("lblGosiMax", "«·Õœ «·√⁄·Ï ··√Ã—", 7371, 9412, 1588, 284, 9, False, CLR_MUTED, "txtGosiMax", 0)
    Set c = AddButton("btnSaveGosi", "Õ›Ÿ ‰”» «· √„Ì‰« ", 9072, 9724, 2041, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13608, 9724, 1474, 454, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    PayrollLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboRun_AfterUpdate()" & vbCrLf
    s = s & "    PayrollPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCreate_Click()" & vbCrLf
    s = s & "    PayrollCreate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRebuild_Click()" & vbCrLf
    s = s & "    PayrollRebuild Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPost_Click()" & vbCrLf
    s = s & "    PayrollPost Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnpost_Click()" & vbCrLf
    s = s & "    PayrollUnpost Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    PayrollDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPay_Click()" & vbCrLf
    s = s & "    PayrollPay Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUndoPay_Click()" & vbCrLf
    s = s & "    PayrollUndoPay Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintPayroll Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnEmployees_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmEmployeePay""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSaveGosi_Click()" & vbCrLf
    s = s & "    SaveGosiRates Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmPayroll", s
    Exit Sub
EH:
    AbortForm "frmPayroll", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBudgetLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBudgetLines", "√”ÿ— «·„Ê«“‰…", "SELECT * FROM BudgetLines WHERE BudgetID = 0 ORDER BY AccountCode", 14855, 397, False, True, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddCombo("AccountCode", "AccountCode", 28, 0, 2211, 397, "SELECT AccountCode, AccountCode & '  ' & AccountName FROM Accounts WHERE AccountType IN ('REVENUE', 'EXPENSE') AND IsActive = True ORDER BY TreeKey", 2, "0;3118")
    SetCtlProp c, "BoundColumn", 1
    SetCtlProp c, "LimitToList", True
    Set c = AddCombo("CostCenterID", "CostCenterID", 2267, 0, 1191, 397, "SELECT CostCenterID, CenterName FROM CostCenters WHERE IsActive = True ORDER BY CenterCode", 2, "0;1701")
    SetCtlProp c, "BoundColumn", 1
    SetCtlProp c, "LimitToList", True
    Set c = AddText("M1", "M1", 3486, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M2", "M2", 4336, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M3", "M3", 5186, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M4", "M4", 6036, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M5", "M5", 6886, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M6", "M6", 7736, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M7", "M7", 8586, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M8", "M8", 9436, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M9", "M9", 10286, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M10", "M10", 11136, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M11", "M11", 11986, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("M12", "M12", 12836, 0, 822, 397)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("txtYearTotal", "=Nz([M1],0)+Nz([M2],0)+Nz([M3],0)+Nz([M4],0)+Nz([M5],0)+Nz([M6],0)+Nz([M7],0)+Nz([M8],0)+Nz([M9],0)+Nz([M10],0)+Nz([M11],0)+Nz([M12],0)", 13686, 0, 964, 397)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0"
    Set c = AddText("BudgetID", "BudgetID", 14678, 0, 28, 397)
    SetCtlProp c, "Visible", False
    m_frm.BeforeInsert = EP
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    s = ""
    s = s & "Private Sub Form_BeforeInsert(Cancel As Integer)" & vbCrLf
    s = s & "    BudgetLineInsert Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    BudgetLineCheck Me, Cancel" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    BudgetLineSaved Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBudgetLines", s
    Exit Sub
EH:
    AbortForm "frmBudgetLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmBudget()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmBudget", "«·„Ê«“‰… «· ﬁœÌ—Ì…", "", 15309, 10886, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„Ê«“‰… «· ﬁœÌ—Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„»·€ ‘Â—Ì ·ﬂ· Õ”«» ≈Ì—«œ«  √Ê „’—Ê›«  (ÊÌ„ﬂ‰ ·ﬂ· „—ﬂ“)° À„ «·„ﬁ«—‰… »«·›⁄·Ì", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboBudget", "", 227, 1304, 2608, 454, "SELECT BudgetID, BudgetYear & '  ' & BudgetName FROM Budgets ORDER BY BudgetYear DESC", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblBudget", "«·„Ê«“‰…", 227, 992, 2608, 284, 9, False, CLR_MUTED, "cboBudget", 0)
    Set c = AddText("txtYear", "", 2948, 1304, 1021, 454)
    SetCtlProp c, "Format", "0"
    Set c = AddLabel("lblYear", "”‰… ÃœÌœ…", 2948, 992, 1021, 284, 9, False, CLR_MUTED, "txtYear", 0)
    Set c = AddButton("btnCreate", "≈‰‘«¡ „Ê«“‰…", 4082, 1304, 1588, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnAddAccounts", "≈÷«›… ﬂ· «·Õ”«»« ", 5783, 1304, 1928, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtFillYear", "", 7825, 1304, 1021, 454)
    SetCtlProp c, "Format", "0"
    Set c = AddLabel("lblFillYear", "›⁄·Ì ”‰…", 7825, 992, 1021, 284, 9, False, CLR_MUTED, "txtFillYear", 0)
    Set c = AddText("txtPercent", "", 8959, 1304, 1021, 454)
    SetCtlProp c, "Format", "0%"
    Set c = AddLabel("lblPercent", "“Ì«œ…", 8959, 992, 1021, 284, 9, False, CLR_MUTED, "txtPercent", 0)
    Set c = AddButton("btnFill", "„·¡ „‰ «·›⁄·Ì", 10093, 1304, 1701, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtAnnual", "", 11907, 1304, 1247, 454)
    SetCtlProp c, "Format", "#,##0"
    Set c = AddLabel("lblAnnual", "„»·€ ”‰ÊÌ", 11907, 992, 1247, 284, 9, False, CLR_MUTED, "txtAnnual", 0)
    Set c = AddButton("btnSpread", " Ê“Ì⁄ ⁄·Ï «·√‘Â—", 13268, 1304, 1814, 454, "secondary")
    c.OnClick = EP
    Set c = AddText("txtBudgetID", "", 15139, 907, 113, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblCol1", "«·Õ”«»", 255, 1956, 2211, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "«·„—ﬂ“", 2494, 1956, 1191, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "Ì‰«Ì—", 3713, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "›»—«Ì—", 4563, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "„«—”", 5413, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "√»—Ì·", 6263, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "„«ÌÊ", 7113, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol8", "ÌÊ‰ÌÊ", 7963, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol9", "ÌÊ·ÌÊ", 8813, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol10", "√€”ÿ”", 9663, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol11", "”» „»—", 10513, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol12", "√ﬂ Ê»—", 11363, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol13", "‰Ê›„»—", 12213, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol14", "œÌ”„»—", 13063, 1956, 822, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol15", "«·”‰…", 13913, 1956, 964, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmBudgetLines", 227, 2296, 14855, 3912)
    Set c = AddLabel("lblTotals", " ", 227, 6294, 11340, 340, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnDelete", "Õ–› «·„Ê«“‰…", 13381, 6237, 1701, 425, "danger")
    c.OnClick = EP
    Set c = AddText("txtFrom", "", 227, 7116, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„ﬁ«—‰… „‰", 227, 6804, 1588, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 1928, 7116, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï", 1928, 6804, 1588, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnCompare", "«·„Ê«“‰… „ﬁ«»· «·›⁄·Ì", 3629, 7116, 2381, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·„ﬁ«—‰…", 6124, 7116, 1701, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblVariance", " ", 7938, 7116, 7144, 454, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstVariance", 227, 7711, 14855, 2381, 9, "0;1134;3402;1928;1701;1701;1701;1021;1361", True)
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 10206, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    BudgetLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboBudget_AfterUpdate()" & vbCrLf
    s = s & "    BudgetPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCreate_Click()" & vbCrLf
    s = s & "    BudgetCreate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAddAccounts_Click()" & vbCrLf
    s = s & "    BudgetAddAccounts Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnFill_Click()" & vbCrLf
    s = s & "    BudgetFill Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnSpread_Click()" & vbCrLf
    s = s & "    BudgetSpread Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    BudgetDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCompare_Click()" & vbCrLf
    s = s & "    BudgetCompare Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintBudgetVariance Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmBudget", s
    Exit Sub
EH:
    AbortForm "frmBudget", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAccounting()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAccounting", "«·„Õ«”»… Ê«·„«·Ì…", "", 14882, 9381, False, False, False, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 14882, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8F1), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "«·„Õ«”»… Ê«·„«·Ì…", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ﬂ· ‘«‘«  «·Õ”«»«  ›Ì „ﬂ«‰ Ê«Õœ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddRect("boxNavJournal", 11140, 1134, 3402, 1304, RGB(31, 58, 95))
    Set c = AddIcon("icoTileJournal", ChrW(&HE8F1), 11140, 1247, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileJournal", "ﬁÌÊœ «·ÌÊ„Ì…", 11140, 1758, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintJournal", "ﬂ· «·ﬁÌÊœ Ê„Ì“«‰ «·„—«Ã⁄…", 11140, 2070, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileJournal", "ﬁÌÊœ «·ÌÊ„Ì…", 11140, 1134, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmJournal"
    c.OnClick = EP
    Set c = AddRect("boxNavAccounts", 7540, 1134, 3402, 1304, RGB(57, 73, 171))
    Set c = AddIcon("icoTileAccounts", ChrW(&HE8FD), 7540, 1247, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileAccounts", "œ·Ì· «·Õ”«»« ", 7540, 1758, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintAccounts", "‘Ã—… «·Õ”«»« ", 7540, 2070, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileAccounts", "œ·Ì· «·Õ”«»« ", 7540, 1134, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAccounts"
    c.OnClick = EP
    Set c = AddRect("boxNavManual", 3940, 1134, 3402, 1304, RGB(94, 53, 177))
    Set c = AddIcon("icoTileManual", ChrW(&HE8F1), 3940, 1247, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileManual", "ﬁÌœ ÌœÊÌ", 3940, 1758, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintManual", "ﬁÌÊœ «· ”ÊÌ… Ê«·«›  «Õ", 3940, 2070, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileManual", "ﬁÌœ ÌœÊÌ", 3940, 1134, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmManualEntry"
    c.OnClick = EP
    Set c = AddRect("boxNavLedger", 340, 1134, 3402, 1304, RGB(30, 136, 229))
    Set c = AddIcon("icoTileLedger", ChrW(&HE8A5), 340, 1247, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileLedger", "ﬂ‘› Õ”«»", 340, 1758, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintLedger", "Õ—ﬂ… Õ”«» Êœ› — «·√” «–", 340, 2070, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileLedger", "ﬂ‘› Õ”«»", 340, 1134, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmLedger"
    c.OnClick = EP
    Set c = AddRect("boxNavFinancials", 11140, 2636, 3402, 1304, RGB(0, 121, 107))
    Set c = AddIcon("icoTileFinancials", ChrW(&HE8A5), 11140, 2749, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileFinancials", "«·Õ”«»«  «·Œ «„Ì…", 11140, 3260, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintFinancials", "ﬁ«∆„… «·œŒ· Ê«·„Ì“«‰Ì…", 11140, 3572, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileFinancials", "«·Õ”«»«  «·Œ «„Ì…", 11140, 2636, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmFinancials"
    c.OnClick = EP
    Set c = AddRect("boxNavClosing", 7540, 2636, 3402, 1304, RGB(84, 110, 122))
    Set c = AddIcon("icoTileClosing", ChrW(&HE713), 7540, 2749, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileClosing", "≈ﬁ›«· «·› —« ", 7540, 3260, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintClosing", "≈ﬁ›«· «·› —… Ê«·”‰… «·„«·Ì…", 7540, 3572, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileClosing", "≈ﬁ›«· «·› —« ", 7540, 2636, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPeriodClosing"
    c.OnClick = EP
    Set c = AddRect("boxNavVat", 3940, 2636, 3402, 1304, RGB(216, 27, 96))
    Set c = AddIcon("icoTileVat", ChrW(&HE8A5), 3940, 2749, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileVat", "«·≈ﬁ—«— «·÷—Ì»Ì", 3940, 3260, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintVat", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 3940, 3572, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileVat", "«·≈ﬁ—«— «·÷—Ì»Ì", 3940, 2636, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmVatReturn"
    c.OnClick = EP
    Set c = AddRect("boxNavAging", 340, 2636, 3402, 1304, RGB(229, 57, 53))
    Set c = AddIcon("icoTileAging", ChrW(&HE716), 340, 2749, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileAging", "√⁄„«— «·œÌÊ‰", 340, 3260, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintAging", "«·⁄„·«¡ Ê«·„Ê—œÊ‰ Õ”» «·«” Õﬁ«ﬁ", 340, 3572, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileAging", "√⁄„«— «·œÌÊ‰", 340, 2636, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAging"
    c.OnClick = EP
    Set c = AddRect("boxNavBanks", 11140, 4138, 3402, 1304, RGB(0, 137, 123))
    Set c = AddIcon("icoTileBanks", ChrW(&HE825), 11140, 4251, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileBanks", "«·»‰Êﬂ", 11140, 4762, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintBanks", "«·Õ”«»«  «·»‰ﬂÌ… Ê«· ”ÊÌ…", 11140, 5074, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileBanks", "«·»‰Êﬂ", 11140, 4138, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmBanks"
    c.OnClick = EP
    Set c = AddRect("boxNavCheques", 7540, 4138, 3402, 1304, RGB(67, 160, 71))
    Set c = AddIcon("icoTileCheques", ChrW(&HE825), 7540, 4251, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileCheques", "«·‘Ìﬂ« ", 7540, 4762, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintCheques", "«·Ê«—œ… Ê«·’«œ—…", 7540, 5074, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileCheques", "«·‘Ìﬂ« ", 7540, 4138, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCheques"
    c.OnClick = EP
    Set c = AddRect("boxNavAssets", 3940, 4138, 3402, 1304, RGB(251, 140, 0))
    Set c = AddIcon("icoTileAssets", ChrW(&HE7B8), 3940, 4251, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileAssets", "«·√’Ê· «·À«» …", 3940, 4762, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintAssets", "«·‘—«¡ Ê«·»Ì⁄ Ê«·«” »⁄«œ", 3940, 5074, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileAssets", "«·√’Ê· «·À«» …", 3940, 4138, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAssets"
    c.OnClick = EP
    Set c = AddRect("boxNavDepreciation", 340, 4138, 3402, 1304, RGB(239, 108, 0))
    Set c = AddIcon("icoTileDepreciation", ChrW(&HE7B8), 340, 4251, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileDepreciation", "«·≈Â·«ﬂ «·‘Â—Ì", 340, 4762, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintDepreciation", " ”ÃÌ· ≈Â·«ﬂ «·√’Ê·", 340, 5074, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileDepreciation", "«·≈Â·«ﬂ «·‘Â—Ì", 340, 4138, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmDepreciation"
    c.OnClick = EP
    Set c = AddRect("boxNavPayroll", 11140, 5640, 3402, 1304, RGB(142, 36, 170))
    Set c = AddIcon("icoTilePayroll", ChrW(&HE8D7), 11140, 5753, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTilePayroll", "«·—Ê« »", 11140, 6264, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintPayroll", "„”Ì— «·—Ê« » Ê«· √„Ì‰« ", 11140, 6576, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTilePayroll", "«·—Ê« »", 11140, 5640, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmPayroll"
    c.OnClick = EP
    Set c = AddRect("boxNavCenters", 7540, 5640, 3402, 1304, RGB(0, 131, 143))
    Set c = AddIcon("icoTileCenters", ChrW(&HE8FD), 7540, 5753, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileCenters", "„—«ﬂ“ «· ﬂ·›…", 7540, 6264, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintCenters", "«·›—Ê⁄ Ê«·√ﬁ”«„", 7540, 6576, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileCenters", "„—«ﬂ“ «· ﬂ·›…", 7540, 5640, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCostCenters"
    c.OnClick = EP
    Set c = AddRect("boxNavBudget", 3940, 5640, 3402, 1304, RGB(121, 85, 72))
    Set c = AddIcon("icoTileBudget", ChrW(&HE8A5), 3940, 5753, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileBudget", "«·„Ê«“‰…", 3940, 6264, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintBudget", "«·„Ê«“‰… „ﬁ«»· «·›⁄·Ì", 3940, 6576, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileBudget", "«·„Ê«“‰…", 3940, 5640, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmBudget"
    c.OnClick = EP
    Set c = AddRect("boxNavRecurring", 340, 5640, 3402, 1304, RGB(198, 40, 40))
    Set c = AddIcon("icoTileRecurring", ChrW(&HE8C7), 340, 5753, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileRecurring", "«·„’—Ê›«  «·„ ﬂ——…", 340, 6264, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintRecurring", "«·≈ÌÃ«— Ê«·ﬂÂ—»«¡ Ê«·«‘ —«ﬂ« ", 340, 6576, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileRecurring", "«·„’—Ê›«  «·„ ﬂ——…", 340, 5640, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmRecurring"
    c.OnClick = EP
    Set c = AddRect("boxNavCurrencies", 11140, 7142, 3402, 1304, RGB(46, 125, 50))
    Set c = AddIcon("icoTileCurrencies", ChrW(&HE825), 11140, 7255, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileCurrencies", "«·⁄„·« ", 11140, 7766, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintCurrencies", "«·⁄„·«  Ê√”⁄«— «· ÕÊÌ·", 11140, 8078, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileCurrencies", "«·⁄„·« ", 11140, 7142, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCurrencies"
    c.OnClick = EP
    Set c = AddRect("boxNavSalesReps", 7540, 7142, 3402, 1304, RGB(0, 105, 92))
    Set c = AddIcon("icoTileSalesReps", ChrW(&HE716), 7540, 7255, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileSalesReps", "«·„‰œÊ»Ì‰", 7540, 7766, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintSalesReps", "«·⁄„·«¡ Ê«·√Âœ«› Ê«·⁄„Ê·« ", 7540, 8078, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileSalesReps", "«·„‰œÊ»Ì‰", 7540, 7142, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmSalesReps"
    c.OnClick = EP
    Set c = AddRect("boxNavCommissions", 3940, 7142, 3402, 1304, RGB(106, 27, 154))
    Set c = AddIcon("icoTileCommissions", ChrW(&HE716), 3940, 7255, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileCommissions", "⁄„Ê·«  «·„‰œÊ»Ì‰", 3940, 7766, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintCommissions", "„”Ì— «·⁄„Ê·«  «·‘Â—Ì", 3940, 8078, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileCommissions", "⁄„Ê·«  «·„‰œÊ»Ì‰", 3940, 7142, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmCommissions"
    c.OnClick = EP
    Set c = AddRect("boxNavAudit", 340, 7142, 3402, 1304, RGB(69, 90, 100))
    Set c = AddIcon("icoTileAudit", ChrW(&HE8D7), 340, 7255, 3402, 510, 20, False, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblTileAudit", "”Ã· «· œﬁÌﬁ", 340, 7766, 3402, 340, 12, True, CLR_SURFACE, "", 2)
    Set c = AddLabel("lblHintAudit", "„‰ √÷«› √Ê ⁄œ¯· √Ê Õ–›", 340, 8078, 3402, 284, 8, False, CLR_SURFACE, "", 2)
    Set c = AddButton("btnTileAudit", "”Ã· «· œﬁÌﬁ", 340, 7142, 3402, 1304, "secondary")
    SetCtlProp c, "Transparent", True
    SetCtlProp c, "Tag", "frmAuditLog"
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 13181, 8757, 1361, 454, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnResize = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    ApplyNavPermissions Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileJournal_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmJournal""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileAccounts_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAccounts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileManual_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmManualEntry""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileLedger_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmLedger""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileFinancials_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmFinancials""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileClosing_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPeriodClosing"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileVat_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmVatReturn"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileAging_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAging"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileBanks_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBanks""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileCheques_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCheques"", 0, ""IN""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileAssets_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAssets"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileDepreciation_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmDepreciation"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTilePayroll_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmPayroll"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileCenters_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCostCenters""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileBudget_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmBudget""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileRecurring_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmRecurring""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileCurrencies_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCurrencies""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileSalesReps_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSalesReps""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileCommissions_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCommissions"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnTileAudit_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmAuditLog"", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_Resize()" & vbCrLf
    s = s & "    Dim spec As String" & vbCrLf
    s = s & "    spec = ""boxTitle,0,0,14882,850,0,1000,0,0;btnClose,13181,8757,1361,454,1000,0,1000,0""" & vbCrLf
    s = s & "    FitControls Me, 14882, 9381, -254, " & IIf(MirrorLayout(), "True", "False") & ", spec" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAccounting", s
    Exit Sub
EH:
    AbortForm "frmAccounting", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmAuditLog()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmAuditLog", "”Ã· «· œﬁÌﬁ", "", 15309, 10433, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8D7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "”Ã· «· œﬁÌﬁ", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„‰ √÷«› √Ê ⁄œ¯· √Ê Õ–›° Ê„ Ï° Ê„‰ √Ì ÃÂ«“° Ê«·ﬁÌ„ ﬁ»· Ê»⁄œ", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtFrom", "", 227, 1304, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "„‰", 227, 992, 1588, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 1928, 1304, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "≈·Ï", 1928, 992, 1588, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddCombo("cboUser", "", 3629, 1304, 2268, 454, "SELECT EmployeeID, EmployeeName FROM Employees ORDER BY EmployeeName", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblUser", "«·„” Œœ„ (›«—€ = «·ﬂ·)", 3629, 992, 2268, 284, 9, False, CLR_MUTED, "cboUser", 0)
    Set c = AddCombo("cboAction", "", 6010, 1304, 2155, 454, "ADD;≈÷«›…;EDIT; ⁄œÌ·;DELETE;Õ–›;DOCS;⁄„·Ì«  «·„” ‰œ« ;LOGIN;«·œŒÊ· Ê«·Œ—ÊÃ", 2, "0;2041")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblAction", "«·⁄„·Ì… (›«—€ = «·ﬂ·)", 6010, 992, 2155, 284, 9, False, CLR_MUTED, "cboAction", 0)
    Set c = AddCombo("cboTable", "", 8278, 1304, 2155, 454, "", 1, "2041")
    SetCtlProp c, "LimitToList", True
    Set c = AddLabel("lblTable", "«·ÃœÊ· (›«—€ = «·ﬂ·)", 8278, 992, 2155, 284, 9, False, CLR_MUTED, "cboTable", 0)
    Set c = AddText("txtSearch", "", 10546, 1304, 2608, 454)
    Set c = AddLabel("lblSearch", "—ﬁ„ √Ê «”„ «·”Ã·", 10546, 992, 2608, 284, 9, False, CLR_MUTED, "txtSearch", 0)
    Set c = AddButton("btnShow", "⁄—÷", 13268, 1304, 1814, 454, "primary")
    c.OnClick = EP
    Set c = AddLabel("lblCount", " ", 227, 1871, 14855, 312, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddList("lstLog", 227, 2268, 14855, 3742, 8, "0;1814;1814;1814;1814;907;3175;1701", True)
    c.AfterUpdate = EP
    Set c = AddLabel("lblDetails", " ", 227, 6067, 14855, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstChanges", 227, 6464, 14855, 2608, 4, "0;2835;5386;5386", True)
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·› —…", 227, 9412, 1814, 510, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "≈€·«ﬁ", 13608, 9412, 1474, 510, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    AuditScreenLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstLog_AfterUpdate()" & vbCrLf
    s = s & "    AuditScreenPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnShow_Click()" & vbCrLf
    s = s & "    AuditScreenShow Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintAuditLog Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmAuditLog", s
    Exit Sub
EH:
    AbortForm "frmAuditLog", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCommissionLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCommissionLines", "√”ÿ— „”Ì— «·⁄„Ê·« ", "SELECT * FROM CommissionLines WHERE CommissionRunID = 0 ORDER BY RepName", 13721, 425, False, False, True, _
              ""
    SetFormProp "DefaultView", 1
    SetFormProp "ScrollBars", 2
    SetFormProp "Cycle", 0
    Set c = AddText("RepName", "RepName", 28, 0, 2608, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddText("NetSales", "NetSales", 2664, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Collections", "Collections", 4166, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddCombo("CommissionBase", "CommissionBase", 5668, 0, 1588, 425, "SALES;’«›Ì «·„»Ì⁄«  (»œÊ‰ «·÷—Ì»…);COLLECTION;«· Õ’Ì·", 2, "0;1474")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddText("BaseAmount", "BaseAmount", 7284, 0, 1474, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("CommissionRate", "CommissionRate", 8786, 0, 907, 425)
    SetCtlProp c, "Format", "0.00%"
    c.AfterUpdate = EP
    Set c = AddText("Adjustment", "Adjustment", 9721, 0, 1247, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddText("Commission", "Commission", 10996, 0, 1361, 425)
    c.FontBold = True
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddText("Notes", "Notes", 12385, 0, 1276, 425)
    m_frm.BeforeUpdate = EP
    m_frm.AfterUpdate = EP
    s = ""
    s = s & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    AuditFormBefore Me, ""CommissionLines"", ""CommissionLineID""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Form_AfterUpdate()" & vbCrLf
    s = s & "    AuditFormAfter Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CommissionBase_AfterUpdate()" & vbCrLf
    s = s & "    CommissionLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub CommissionRate_AfterUpdate()" & vbCrLf
    s = s & "    CommissionLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Adjustment_AfterUpdate()" & vbCrLf
    s = s & "    CommissionLineChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCommissionLines", s
    Exit Sub
EH:
    AbortForm "frmCommissionLines", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmCommissions()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCommissions", "⁄„Ê·«  «·„‰œÊ»Ì‰", "", 14175, 9185, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 14175, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "⁄„Ê·«  «·„‰œÊ»Ì‰", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "„”Ì— ﬂ· ‘Â— „‰ „»Ì⁄«  «·„‰œÊ»Ì‰ √Ê  Õ’Ì·« Â„° À„  —ÕÌ· ﬁÌœÂ∫ «·’—› »”‰œ ’—› ‰ﬁœÌ…", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboRun", "", 227, 1304, 2608, 454, "SELECT CommissionRunID, Format(RunMonth, 'yyyy/mm') & '  ' & IIf(Status = 'DRAFT', '„”Êœ…', '„—ÕÛ¯·') FROM CommissionRuns ORDER BY RunMonth DESC", 2, "0;3969")
    SetCtlProp c, "LimitToList", True
    c.AfterUpdate = EP
    Set c = AddLabel("lblRun", "«·„”Ì—« ", 227, 992, 2608, 284, 9, False, CLR_MUTED, "cboRun", 0)
    Set c = AddText("txtMonth", "", 2948, 1304, 1588, 454)
    SetCtlProp c, "Format", "yyyy/mm"
    Set c = AddLabel("lblMonth", "‘Â— ÃœÌœ", 2948, 992, 1588, 284, 9, False, CLR_MUTED, "txtMonth", 0)
    Set c = AddButton("btnCreate", "≈‰‘«¡ „”Ì— «·‘Â—", 4649, 1304, 2041, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnRebuild", "≈⁄«œ… «·≈‰‘«¡", 6804, 1304, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPost", " —ÕÌ· «·„”Ì—", 8477, 1304, 1588, 454, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUnpost", "≈·€«¡ «· —ÕÌ·", 10150, 1304, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "Õ–› «·„”Êœ…", 11823, 1304, 1474, 454, "danger")
    c.OnClick = EP
    Set c = AddText("txtRunID", "", 14005, 907, 113, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblState", " ", 227, 1899, 13721, 340, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblCol1", "«·„‰œÊ»", 255, 2325, 2608, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "’«›Ì «·„»Ì⁄« ", 2891, 2325, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "«· Õ’Ì·", 4393, 2325, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "«·√”«”", 5895, 2325, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "„»·€ «·√”«”", 7511, 2325, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "«·‰”»…", 9013, 2325, 907, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", " ⁄œÌ· (+/-)", 9948, 2325, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol8", "«·⁄„Ê·…", 11223, 2325, 1361, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol9", "„·«ÕŸ« ", 12612, 2325, 1276, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmCommissionLines", 227, 2665, 13721, 4876)
    Set c = AddLabel("lblTotals", " ", 227, 7654, 13721, 369, 10, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnPrint", "ÿ»«⁄… «·„”Ì—", 227, 8448, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReps", "«·„‰œÊ»Ì‰", 1928, 8448, 1588, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayVoucher", "”‰œ ’—› ⁄„Ê·…", 3629, 8448, 1814, 454, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "—ÃÊ⁄", 12474, 8448, 1474, 454, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    CommissionsLoad Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboRun_AfterUpdate()" & vbCrLf
    s = s & "    CommissionsPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnCreate_Click()" & vbCrLf
    s = s & "    CommissionsCreate Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnRebuild_Click()" & vbCrLf
    s = s & "    CommissionsRebuild Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPost_Click()" & vbCrLf
    s = s & "    CommissionsPost Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnUnpost_Click()" & vbCrLf
    s = s & "    CommissionsUnpost Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    CommissionsDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPrint_Click()" & vbCrLf
    s = s & "    PrintCommissionRun Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnReps_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmSalesReps""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnPayVoucher_Click()" & vbCrLf
    s = s & "    OpenScreen ""frmCashVoucher"", 0, ""OUT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmCommissions", s
    Exit Sub
EH:
    AbortForm "frmCommissions", Err.Number, Err.Description
End Sub
