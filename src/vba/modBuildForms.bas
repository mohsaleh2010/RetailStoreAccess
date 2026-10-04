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
Private Const FORM_NAMES As String = "frmMain,frmProducts,frmCustomers,frmSuppliers,frmExpenses,frmCategories,frmUnits,frmExpenseTypes,frmSettings,frmSearch,frmReportCenter,frmPOSLines,frmPOS,frmReturnLines,frmSalesReturn,frmCustomerPayment,frmSalesInvoice"

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
    EnsureLocalTables                      ' tmpPOSLines / tmpReturnLines (modPOS)
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
        MsgBox "جميع اختبارات الشاشات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "تم حذف بيانات الاختبار.", vbInformation + MSG_RTL, "TestForms"
        TestForms = True
    Else
        MsgBox "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & _
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
    MsgBox errText, vbCritical + MSG_RTL, "TestForms"
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
    BuildForm_frmMain
    BuildForm_frmProducts
    BuildForm_frmCustomers
    BuildForm_frmSuppliers
    BuildForm_frmExpenses
    BuildForm_frmCategories
    BuildForm_frmUnits
    BuildForm_frmExpenseTypes
    BuildForm_frmSettings
    BuildForm_frmSearch
    BuildForm_frmReportCenter
    BuildForm_frmPOSLines
    BuildForm_frmPOS
    BuildForm_frmReturnLines
    BuildForm_frmSalesReturn
    BuildForm_frmCustomerPayment
    BuildForm_frmSalesInvoice
End Sub

Private Sub BuildForm_frmMain()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmMain", "نظام إدارة المحل", "", 18994, 10773, False, False, False, _
              ""
    Set c = AddRect("boxSidebar", 0, 0, 3515, 10773, CLR_PRIMARY)
    Set c = AddIcon("icoApp", ChrW(&HE80F), 227, 255, 567, 567, 22, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblAppTitle", "نظام إدارة المحل", 850, 227, 2551, 425, 15, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblStoreName", " ", 850, 652, 2551, 312, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNavSales", "المبيعات", 142, 1304, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoSales", ChrW(&HE7BF), 255, 1389, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavPurchases", "المشتريات", 142, 1899, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoPurchases", ChrW(&HE896), 255, 1984, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavInventory", "المخزون", 142, 2494, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoInventory", ChrW(&HE7B8), 255, 2579, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavProducts", "المنتجات", 142, 3089, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoProducts", ChrW(&HE8EC), 255, 3174, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavCustomers", "العملاء", 142, 3684, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoCustomers", ChrW(&HE716), 255, 3769, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSuppliers", "الموردون", 142, 4279, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoSuppliers", ChrW(&HE77B), 255, 4364, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavExpenses", "المصروفات", 142, 4874, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoExpenses", ChrW(&HE8C7), 255, 4959, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavStockCount", "الجرد", 142, 5469, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoStockCount", ChrW(&HE8EF), 255, 5554, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavReports", "التقارير", 142, 6064, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoReports", ChrW(&HE8A5), 255, 6149, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSearch", "البحث", 142, 6659, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoSearch", ChrW(&HE721), 255, 6744, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavSettings", "الإعدادات", 142, 7254, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoSettings", ChrW(&HE713), 255, 7339, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavUsers", "المستخدمون", 142, 7849, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoUsers", ChrW(&HE8D7), 255, 7934, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavBackup", "نسخة احتياطية", 142, 8444, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoBackup", ChrW(&HE8B7), 255, 8529, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddButton("btnNavLogout", "تسجيل الخروج", 142, 9039, 3231, 539, "nav")
    c.OnClick = EP
    Set c = AddIcon("icoLogout", ChrW(&HE7E8), 255, 9124, 454, 369, 13, False, CLR_SIDEBAR_TEXT, "", 0)
    c.OnClick = EP
    Set c = AddLabel("lblWelcome", "لوحة التحكم", 3969, 340, 7938, 539, 20, True, CLR_TEXT, "", 0)
    Set c = AddLabel("lblToday", " ", 3969, 907, 5670, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblUser", " ", 13437, 907, 5103, 340, 11, False, CLR_MUTED, "", 3)
    Set c = AddRect("boxTile1", 3969, 1531, 3472, 1418, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle1", "مبيعات اليوم", 4139, 1644, 3132, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTileValue1", "-", 4139, 2041, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxTile2", 7668, 1531, 3472, 1418, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle2", "مبيعات الشهر", 7838, 1644, 3132, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTileValue2", "-", 7838, 2041, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxTile3", 11367, 1531, 3472, 1418, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle3", "صافي الربح التقريبي", 11537, 1644, 3132, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTileValue3", "-", 11537, 2041, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    Set c = AddRect("boxTile4", 15066, 1531, 3472, 1418, CLR_SURFACE)
    Set c = AddLabel("lblTileTitle4", "منتجات منخفضة المخزون", 15236, 1644, 3132, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTileValue4", "-", 15236, 2041, 3132, 567, 20, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTilesNote", "مؤشرات لوحة التحكم تُفعَّل في المرحلة 9", 3969, 3062, 14571, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblIntegrity", " ", 3969, 10093, 14571, 340, 10, True, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
    s = ""
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    MainLoad Me" & vbCrLf
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
    s = s & "    ExitApplication" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub icoLogout_Click()" & vbCrLf
    s = s & "    ExitApplication" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmMain", s
    Exit Sub
EH:
    AbortForm "frmMain", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmProducts()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmProducts", "المنتجات", "SELECT * FROM Products", 15309, 9213, True, True, True, _
              "KIND=LIST|TABLE=Products|PK=ProductID|LIST=SELECT t.ProductID, t.ProductCode AS [الكود], t.ProductName AS [المنتج], t.CurrentQuantity AS [الكمية] FROM Products AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ProductName|SEARCH=t.ProductName,t.ProductCode,t.Barcode,t.ProductNameEn|ACTIVE=t.IsActive|SEQ=PRODUCT_CODE:ProductCode|UNIQUE=ProductCode,Barcode"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8EC), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "المنتجات", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "إضافة وتعديل الأصناف والأسعار", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 5924, 4, "0;1134;2778;850", True)
    c.AfterUpdate = EP
    Set c = AddText("ProductCode", "ProductCode", 7201, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "يُولَّد تلقائيًا إذا تُرك فارغًا"
    SetCtlProp c, "StatusBarText", "يُولَّد تلقائيًا إذا تُرك فارغًا"
    Set c = AddLabel("lblProductCode", "كود المنتج", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ProductCode", 0)
    Set c = AddText("Barcode", "Barcode", 12134, 1701, 2948, 425)
    Set c = AddLabel("lblBarcode", "الباركود", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "Barcode", 0)
    Set c = AddText("ProductName", "ProductName", 7201, 2268, 7881, 425)
    Set c = AddLabel("lblProductName", "اسم المنتج *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ProductName", 0)
    Set c = AddText("ProductNameEn", "ProductNameEn", 7201, 2835, 7881, 425)
    Set c = AddLabel("lblProductNameEn", "الاسم بالإنجليزية", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "ProductNameEn", 0)
    Set c = AddCombo("CategoryID", "CategoryID", 7201, 3402, 2948, 425, "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", 2, "0;3402")
    Set c = AddLabel("lblCategoryID", "التصنيف", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "CategoryID", 0)
    Set c = AddCombo("UnitID", "UnitID", 12134, 3402, 2948, 425, "SELECT UnitID, UnitName FROM Units ORDER BY UnitName", 2, "0;3402")
    Set c = AddLabel("lblUnitID", "الوحدة", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "UnitID", 0)
    Set c = AddCombo("SupplierID", "SupplierID", 7201, 3969, 2948, 425, "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", 2, "0;3402")
    Set c = AddLabel("lblSupplierID", "المورد الافتراضي", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "SupplierID", 0)
    Set c = AddCombo("VATCategory", "VATCategory", 12134, 3969, 2948, 425, "S;خاضع للضريبة 15%;Z;نسبة صفرية;E;معفى من الضريبة", 2, "0;2552")
    c.AfterUpdate = EP
    Set c = AddLabel("lblVATCategory", "الفئة الضريبية", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "VATCategory", 0)
    Set c = AddText("SellingPrice", "SellingPrice", 7201, 4536, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblSellingPrice", "سعر البيع", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "SellingPrice", 0)
    Set c = AddText("PurchasePrice", "PurchasePrice", 12134, 4536, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblPurchasePrice", "آخر سعر شراء", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "PurchasePrice", 0)
    Set c = AddLabel("lblPriceInfo", " ", 5443, 5103, 9639, 425, 10, True, CLR_ACCENT, "", 0)
    Set c = AddText("AverageCost", "AverageCost", 7201, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblAverageCost", "متوسط التكلفة", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "AverageCost", 0)
    Set c = AddText("CurrentQuantity", "CurrentQuantity", 12134, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentQuantity", "الكمية الحالية", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "CurrentQuantity", 0)
    Set c = AddText("MinimumQuantity", "MinimumQuantity", 7201, 6237, 2948, 425)
    SetCtlProp c, "Format", "#,##0.###"
    Set c = AddLabel("lblMinimumQuantity", "حد إعادة الطلب", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "MinimumQuantity", 0)
    Set c = AddText("ProductLocation", "ProductLocation", 12134, 6237, 2948, 425)
    Set c = AddLabel("lblProductLocation", "مكان المنتج", 10376, 6237, 1701, 425, 10, False, CLR_MUTED, "ProductLocation", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 6889)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblStockNote", "الكمية تتغير فقط من المشتريات والمبيعات والجرد", 10376, 6804, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 7371, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "ملاحظات", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
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
    StartForm "frmCustomers", "العملاء", "SELECT * FROM Customers", 15309, 9213, True, True, True, _
              "KIND=LIST|TABLE=Customers|PK=CustomerID|LIST=SELECT t.CustomerID, t.CustomerName AS [العميل], t.Mobile AS [الجوال], t.CurrentBalance AS [الرصيد] FROM Customers AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.CustomerName|SEARCH=t.CustomerName,t.Mobile,t.Phone,t.VATNumber|ACTIVE=t.IsActive"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "العملاء", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "بيانات العملاء وأرصدتهم", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "سند قبض", 6123, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 5924, 4, "0;2495;1361;907", True)
    c.AfterUpdate = EP
    Set c = AddText("CustomerName", "CustomerName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblCustomerName", "اسم العميل *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CustomerName", 0)
    Set c = AddText("Mobile", "Mobile", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblMobile", "الجوال", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("Phone", "Phone", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblPhone", "الهاتف", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblEmail", "البريد الإلكتروني", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATNumber", "VATNumber", 12134, 2835, 2948, 425)
    SetCtlProp c, "ControlTipText", "للعملاء المنشآت (فاتورة ضريبية)"
    SetCtlProp c, "StatusBarText", "للعملاء المنشآت (فاتورة ضريبية)"
    Set c = AddLabel("lblVATNumber", "الرقم الضريبي", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 7201, 3402, 2948, 425)
    Set c = AddLabel("lblCRNumber", "السجل التجاري", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("City", "City", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblCity", "المدينة", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddText("District", "District", 7201, 3969, 2948, 425)
    Set c = AddLabel("lblDistrict", "الحي", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "District", 0)
    Set c = AddText("StreetName", "StreetName", 12134, 3969, 2948, 425)
    Set c = AddLabel("lblStreetName", "الشارع", 10376, 3969, 1701, 425, 10, False, CLR_MUTED, "StreetName", 0)
    Set c = AddText("BuildingNo", "BuildingNo", 7201, 4536, 2948, 425)
    Set c = AddLabel("lblBuildingNo", "رقم المبنى", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "BuildingNo", 0)
    Set c = AddText("PostalCode", "PostalCode", 12134, 4536, 2948, 425)
    Set c = AddLabel("lblPostalCode", "الرمز البريدي", 10376, 4536, 1701, 425, 10, False, CLR_MUTED, "PostalCode", 0)
    Set c = AddText("Address", "Address", 7201, 5103, 7881, 425)
    Set c = AddLabel("lblAddress", "العنوان", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "Address", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "يُقفل بعد أول عملية"
    SetCtlProp c, "StatusBarText", "يُقفل بعد أول عملية"
    Set c = AddLabel("lblOpeningBalance", "الرصيد الافتتاحي", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("CurrentBalance", "CurrentBalance", 12134, 5670, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentBalance", "الرصيد الحالي", 10376, 5670, 1701, 425, 10, False, CLR_MUTED, "CurrentBalance", 0)
    Set c = AddCheck("AllowCredit", "AllowCredit", 7201, 6322)
    Set c = AddLabel("lblAllowCredit", "يسمح بالبيع الآجل", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "AllowCredit", 0)
    Set c = AddText("CreditLimit", "CreditLimit", 12134, 6237, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "0 = بدون حد"
    SetCtlProp c, "StatusBarText", "0 = بدون حد"
    Set c = AddLabel("lblCreditLimit", "حد الائتمان", 10376, 6237, 1701, 425, 10, False, CLR_MUTED, "CreditLimit", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 6889)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 6804, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "الرصيد الموجب = مبلغ مستحق على العميل", 10376, 6804, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 7371, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "ملاحظات", 5443, 7371, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
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
    StartForm "frmSuppliers", "الموردون", "SELECT * FROM Suppliers", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Suppliers|PK=SupplierID|LIST=SELECT t.SupplierID, t.SupplierName AS [المورد], t.Mobile AS [الجوال], t.CurrentBalance AS [الرصيد] FROM Suppliers AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.SupplierName|SEARCH=t.SupplierName,t.ContactPerson,t.Mobile,t.VATNumber|ACTIVE=t.IsActive"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE77B), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "الموردون", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "بيانات الموردين وأرصدتهم", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 4, "0;2495;1361;907", True)
    c.AfterUpdate = EP
    Set c = AddText("SupplierName", "SupplierName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblSupplierName", "اسم المورد *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "SupplierName", 0)
    Set c = AddText("ContactPerson", "ContactPerson", 7201, 2268, 2948, 425)
    Set c = AddLabel("lblContactPerson", "الشخص المسؤول", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ContactPerson", 0)
    Set c = AddText("Mobile", "Mobile", 12134, 2268, 2948, 425)
    Set c = AddLabel("lblMobile", "الجوال", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "Mobile", 0)
    Set c = AddText("Phone", "Phone", 7201, 2835, 2948, 425)
    Set c = AddLabel("lblPhone", "الهاتف", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 12134, 2835, 2948, 425)
    Set c = AddLabel("lblEmail", "البريد الإلكتروني", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATNumber", "VATNumber", 7201, 3402, 2948, 425)
    Set c = AddLabel("lblVATNumber", "الرقم الضريبي", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblCRNumber", "السجل التجاري", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("City", "City", 7201, 3969, 2948, 425)
    Set c = AddLabel("lblCity", "المدينة", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddLabel("lblSupplierNote", " ", 10376, 3969, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Address", "Address", 7201, 4536, 7881, 425)
    Set c = AddLabel("lblAddress", "العنوان", 5443, 4536, 1701, 425, 10, False, CLR_MUTED, "Address", 0)
    Set c = AddText("OpeningBalance", "OpeningBalance", 7201, 5103, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "ControlTipText", "يُقفل بعد أول عملية"
    SetCtlProp c, "StatusBarText", "يُقفل بعد أول عملية"
    Set c = AddLabel("lblOpeningBalance", "الرصيد الافتتاحي", 5443, 5103, 1701, 425, 10, False, CLR_MUTED, "OpeningBalance", 0)
    Set c = AddText("CurrentBalance", "CurrentBalance", 12134, 5103, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblCurrentBalance", "الرصيد الحالي", 10376, 5103, 1701, 425, 10, False, CLR_MUTED, "CurrentBalance", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 5755)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 5670, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
    Set c = AddLabel("lblBalanceNote", "الرصيد الموجب = مبلغ مستحق للمورد", 10376, 5670, 4706, 425, 9, False, CLR_MUTED, "", 0)
    Set c = AddText("Notes", "Notes", 7201, 6237, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblNotes", "ملاحظات", 5443, 6237, 1701, 425, 10, False, CLR_MUTED, "Notes", 0)
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
    StartForm "frmExpenses", "المصروفات", "SELECT * FROM Expenses", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Expenses|PK=ExpenseID|LIST=SELECT t.ExpenseID, t.ExpenseNumber AS [الرقم], t.ExpenseDate AS [التاريخ], x.ExpenseTypeName AS [النوع], t.TotalAmount AS [المبلغ] FROM Expenses AS t INNER JOIN ExpenseTypes AS x ON t.ExpenseTypeID = x.ExpenseTypeID WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ExpenseDate DESC, t.ExpenseID DESC|SEARCH=t.ExpenseNumber,t.Description,x.ExpenseTypeName,t.SupplierInvoiceRef|SEQ=EXPENSE:ExpenseNumber|UNIQUE=ExpenseNumber"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8C7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "المصروفات", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "تسجيل مصروفات المحل", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddList("lstItems", 227, 2551, 4990, 5387, 5, "0;1134;1191;1474;964", True)
    c.AfterUpdate = EP
    Set c = AddText("ExpenseNumber", "ExpenseNumber", 7201, 1701, 2948, 425)
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    SetCtlProp c, "ControlTipText", "يُولَّد عند الحفظ"
    SetCtlProp c, "StatusBarText", "يُولَّد عند الحفظ"
    Set c = AddLabel("lblExpenseNumber", "رقم المصروف", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseNumber", 0)
    Set c = AddText("ExpenseDate", "ExpenseDate", 12134, 1701, 2948, 425)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblExpenseDate", "تاريخ المصروف", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseDate", 0)
    Set c = AddCombo("ExpenseTypeID", "ExpenseTypeID", 7201, 2268, 2948, 425, "SELECT ExpenseTypeID, ExpenseTypeName FROM ExpenseTypes ORDER BY ExpenseTypeName", 2, "0;3402")
    Set c = AddLabel("lblExpenseTypeID", "نوع المصروف *", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeID", 0)
    Set c = AddCombo("PaymentMethodID", "PaymentMethodID", 12134, 2268, 2948, 425, "SELECT PaymentMethodID, MethodName FROM PaymentMethods ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethodID", "طريقة الدفع", 10376, 2268, 1701, 425, 10, False, CLR_MUTED, "PaymentMethodID", 0)
    Set c = AddText("Amount", "Amount", 7201, 2835, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblAmount", "المبلغ قبل الضريبة", 5443, 2835, 1701, 425, 10, False, CLR_MUTED, "Amount", 0)
    Set c = AddText("Tax", "Tax", 12134, 2835, 1644, 425)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTax", "ضريبة المدخلات", 10376, 2835, 1701, 425, 10, False, CLR_MUTED, "Tax", 0)
    Set c = AddButton("btnCalcVat", "احسب 15%", 13835, 2835, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("TotalAmount", "TotalAmount", 7201, 3402, 2948, 425)
    SetCtlProp c, "Format", "#,##0.00"
    SetCtlProp c, "Locked", True
    c.BackColor = CLR_LOCKED
    SetCtlProp c, "TabStop", False
    Set c = AddLabel("lblTotalAmount", "الإجمالي", 5443, 3402, 1701, 425, 10, False, CLR_MUTED, "TotalAmount", 0)
    Set c = AddText("SupplierInvoiceRef", "SupplierInvoiceRef", 12134, 3402, 2948, 425)
    Set c = AddLabel("lblSupplierInvoiceRef", "رقم فاتورة المصروف", 10376, 3402, 1701, 425, 10, False, CLR_MUTED, "SupplierInvoiceRef", 0)
    Set c = AddText("Description", "Description", 7201, 3969, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "الوصف", 5443, 3969, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
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

Private Sub BuildForm_frmCategories()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmCategories", "التصنيفات", "SELECT * FROM Categories", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Categories|PK=CategoryID|LIST=SELECT t.CategoryID, t.CategoryName AS [التصنيف] FROM Categories AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.CategoryName|SEARCH=t.CategoryName,t.Description|ACTIVE=t.IsActive|UNIQUE=CategoryName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8FD), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "التصنيفات", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "تصنيفات المنتجات", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 2, "0;4763", True)
    c.AfterUpdate = EP
    Set c = AddText("CategoryName", "CategoryName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblCategoryName", "اسم التصنيف *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "CategoryName", 0)
    Set c = AddText("Description", "Description", 7201, 2268, 7881, 907)
    SetCtlProp c, "EnterKeyBehavior", True
    SetCtlProp c, "ScrollBars", 2
    Set c = AddLabel("lblDescription", "الوصف", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "Description", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 3402)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 3317, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
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
    StartForm "frmUnits", "وحدات القياس", "SELECT * FROM Units", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=Units|PK=UnitID|LIST=SELECT t.UnitID, t.UnitName AS [الوحدة], t.ZatcaUnitCode AS [الرمز] FROM Units AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.UnitName|SEARCH=t.UnitName,t.ZatcaUnitCode|ACTIVE=t.IsActive|UNIQUE=UnitName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8FD), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "وحدات القياس", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "وحدات بيع المنتجات", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 3, "0;3062;1701", True)
    c.AfterUpdate = EP
    Set c = AddText("UnitName", "UnitName", 7201, 1701, 2948, 425)
    Set c = AddLabel("lblUnitName", "اسم الوحدة *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "UnitName", 0)
    Set c = AddText("ZatcaUnitCode", "ZatcaUnitCode", 12134, 1701, 2948, 425)
    SetCtlProp c, "ControlTipText", "مثال: PCE للحبة، KGM للكيلو"
    SetCtlProp c, "StatusBarText", "مثال: PCE للحبة، KGM للكيلو"
    Set c = AddLabel("lblZatcaUnitCode", "رمز الوحدة (UN/ECE)", 10376, 1701, 1701, 425, 10, False, CLR_MUTED, "ZatcaUnitCode", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 2353)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
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
    StartForm "frmExpenseTypes", "أنواع المصروفات", "SELECT * FROM ExpenseTypes", 15309, 8222, True, True, True, _
              "KIND=LIST|TABLE=ExpenseTypes|PK=ExpenseTypeID|LIST=SELECT t.ExpenseTypeID, t.ExpenseTypeName AS [النوع] FROM ExpenseTypes AS t WHERE ({ACTIVE}) AND ({SEARCH}) ORDER BY t.ExpenseTypeName|SEARCH=t.ExpenseTypeName|ACTIVE=t.IsActive|UNIQUE=ExpenseTypeName"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8C7), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "أنواع المصروفات", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "قائمة أنواع المصروفات", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnNew", "جديد", 227, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSave", "حفظ", 1701, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 3175, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnDelete", "حذف", 4649, 1021, 1361, 482, "danger")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblSearch", "بحث (F3)", 227, 1701, 3118, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblCount", " ", 3402, 1701, 1815, 284, 9, False, CLR_MUTED, "", 3)
    Set c = AddText("txtSearch", "", 227, 1985, 4990, 454)
    c.OnChange = EP
    Set c = AddCheck("chkShowInactive", "", 227, 2579)
    SetCtlProp c, "DefaultValue", "False"
    c.AfterUpdate = EP
    Set c = AddLabel("lblShowInactive", "إظهار غير النشط", 567, 2551, 2835, 340, 9, False, CLR_MUTED, "", 0)
    Set c = AddList("lstItems", 227, 3005, 4990, 4933, 2, "0;4763", True)
    c.AfterUpdate = EP
    Set c = AddText("ExpenseTypeName", "ExpenseTypeName", 7201, 1701, 7881, 425)
    Set c = AddLabel("lblExpenseTypeName", "نوع المصروف *", 5443, 1701, 1701, 425, 10, False, CLR_MUTED, "ExpenseTypeName", 0)
    Set c = AddCheck("IsActive", "IsActive", 7201, 2353)
    Set c = AddLabel("lblIsActive", "نشط", 5443, 2268, 1701, 425, 10, False, CLR_MUTED, "IsActive", 0)
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
    StartForm "frmSettings", "الإعدادات", "SELECT * FROM Settings WHERE SettingID = 1", 15309, 8222, True, False, True, _
              "KIND=SINGLE|TABLE=Settings|PK=SettingID"
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE713), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "الإعدادات", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "بيانات المحل الضريبية وإعدادات التشغيل", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddButton("btnSave", "حفظ", 227, 1021, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnUndo", "تراجع", 1701, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnCategories", "التصنيفات", 3175, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnUnits", "الوحدات", 4989, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnExpenseTypes", "أنواع المصروفات", 6803, 1021, 1701, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13721, 1021, 1361, 482, "secondary")
    c.OnClick = EP
    Set c = AddText("StoreName", "StoreName", 2552, 1701, 4989, 425)
    Set c = AddLabel("lblStoreName", "اسم المحل *", 227, 1701, 2268, 425, 10, False, CLR_MUTED, "StoreName", 0)
    Set c = AddText("StoreNameEn", "StoreNameEn", 10093, 1701, 4989, 425)
    Set c = AddLabel("lblStoreNameEn", "اسم المحل بالإنجليزية", 7768, 1701, 2268, 425, 10, False, CLR_MUTED, "StoreNameEn", 0)
    Set c = AddText("VATNumber", "VATNumber", 2552, 2268, 4989, 425)
    SetCtlProp c, "ControlTipText", "15 رقمًا يبدأ وينتهي بـ 3"
    SetCtlProp c, "StatusBarText", "15 رقمًا يبدأ وينتهي بـ 3"
    Set c = AddLabel("lblVATNumber", "الرقم الضريبي", 227, 2268, 2268, 425, 10, False, CLR_MUTED, "VATNumber", 0)
    Set c = AddText("CRNumber", "CRNumber", 10093, 2268, 4989, 425)
    Set c = AddLabel("lblCRNumber", "السجل التجاري", 7768, 2268, 2268, 425, 10, False, CLR_MUTED, "CRNumber", 0)
    Set c = AddText("BuildingNo", "BuildingNo", 2552, 2835, 4989, 425)
    Set c = AddLabel("lblBuildingNo", "رقم المبنى", 227, 2835, 2268, 425, 10, False, CLR_MUTED, "BuildingNo", 0)
    Set c = AddText("StreetName", "StreetName", 10093, 2835, 4989, 425)
    Set c = AddLabel("lblStreetName", "الشارع", 7768, 2835, 2268, 425, 10, False, CLR_MUTED, "StreetName", 0)
    Set c = AddText("District", "District", 2552, 3402, 4989, 425)
    Set c = AddLabel("lblDistrict", "الحي", 227, 3402, 2268, 425, 10, False, CLR_MUTED, "District", 0)
    Set c = AddText("City", "City", 10093, 3402, 4989, 425)
    Set c = AddLabel("lblCity", "المدينة", 7768, 3402, 2268, 425, 10, False, CLR_MUTED, "City", 0)
    Set c = AddText("PostalCode", "PostalCode", 2552, 3969, 4989, 425)
    Set c = AddLabel("lblPostalCode", "الرمز البريدي", 227, 3969, 2268, 425, 10, False, CLR_MUTED, "PostalCode", 0)
    Set c = AddText("AdditionalNo", "AdditionalNo", 10093, 3969, 4989, 425)
    Set c = AddLabel("lblAdditionalNo", "الرقم الإضافي", 7768, 3969, 2268, 425, 10, False, CLR_MUTED, "AdditionalNo", 0)
    Set c = AddText("Phone", "Phone", 2552, 4536, 4989, 425)
    Set c = AddLabel("lblPhone", "الهاتف", 227, 4536, 2268, 425, 10, False, CLR_MUTED, "Phone", 0)
    Set c = AddText("Email", "Email", 10093, 4536, 4989, 425)
    Set c = AddLabel("lblEmail", "البريد الإلكتروني", 7768, 4536, 2268, 425, 10, False, CLR_MUTED, "Email", 0)
    Set c = AddText("VATRate", "VATRate", 2552, 5103, 4989, 425)
    SetCtlProp c, "Format", "0.00%"
    Set c = AddLabel("lblVATRate", "نسبة الضريبة", 227, 5103, 2268, 425, 10, False, CLR_MUTED, "VATRate", 0)
    Set c = AddCheck("PricesIncludeVAT", "PricesIncludeVAT", 10093, 5188)
    Set c = AddLabel("lblPricesIncludeVAT", "الأسعار شاملة الضريبة", 7768, 5103, 2268, 425, 10, False, CLR_MUTED, "PricesIncludeVAT", 0)
    Set c = AddCheck("AllowNegativeStock", "AllowNegativeStock", 2552, 5755)
    Set c = AddLabel("lblAllowNegativeStock", "السماح بالبيع بالسالب", 227, 5670, 2268, 425, 10, False, CLR_MUTED, "AllowNegativeStock", 0)
    Set c = AddText("SlowMovingDays", "SlowMovingDays", 10093, 5670, 4989, 425)
    Set c = AddLabel("lblSlowMovingDays", "أيام عدم الحركة", 7768, 5670, 2268, 425, 10, False, CLR_MUTED, "SlowMovingDays", 0)
    Set c = AddText("BackupFolder", "BackupFolder", 2552, 6237, 3685, 425)
    Set c = AddLabel("lblBackupFolder", "مجلد النسخ الاحتياطي", 227, 6237, 2268, 425, 10, False, CLR_MUTED, "BackupFolder", 0)
    Set c = AddButton("btnBrowseBackup", "استعراض", 6294, 6237, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("BackupKeepCount", "BackupKeepCount", 10093, 6237, 4989, 425)
    Set c = AddLabel("lblBackupKeepCount", "عدد النسخ المحتفظ بها", 7768, 6237, 2268, 425, 10, False, CLR_MUTED, "BackupKeepCount", 0)
    Set c = AddText("LogoPath", "LogoPath", 2552, 6804, 3685, 425)
    Set c = AddLabel("lblLogoPath", "مسار الشعار", 227, 6804, 2268, 425, 10, False, CLR_MUTED, "LogoPath", 0)
    Set c = AddButton("btnBrowseLogo", "استعراض", 6294, 6804, 1247, 425, "secondary")
    c.OnClick = EP
    Set c = AddText("ReceiptFooter", "ReceiptFooter", 10093, 6804, 4989, 425)
    Set c = AddLabel("lblReceiptFooter", "تذييل الفاتورة", 7768, 6804, 2268, 425, 10, False, CLR_MUTED, "ReceiptFooter", 0)
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

Private Sub BuildForm_frmSearch()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmSearch", "البحث المتقدم", "", 15309, 9639, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE721), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "البحث المتقدم", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "ابحث بالاسم أو الكود أو الباركود أو رقم الفاتورة أو الجوال", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboKind", "", 227, 1389, 2381, 454, "PRODUCT;المنتجات;CUSTOMER;العملاء;SUPPLIER;الموردون;SALE;فواتير البيع;PURCHASE;فواتير الشراء", 2, "0;2268")
    SetCtlProp c, "DefaultValue", """PRODUCT"""
    c.AfterUpdate = EP
    Set c = AddLabel("lblKind", "ابحث في", 227, 1077, 2381, 284, 9, False, CLR_MUTED, "cboKind", 0)
    Set c = AddText("txtText", "", 2778, 1389, 4876, 454)
    c.AfterUpdate = EP
    Set c = AddLabel("lblText", "كلمة البحث (اسم، كود، باركود، رقم، جوال)", 2778, 1077, 4876, 284, 9, False, CLR_MUTED, "txtText", 0)
    Set c = AddText("txtFrom", "", 7825, 1389, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "من تاريخ", 7825, 1077, 1814, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 9809, 1389, 1814, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "إلى تاريخ", 9809, 1077, 1814, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnSearch", "بحث", 11794, 1372, 1361, 482, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClear", "مسح", 13268, 1372, 907, 482, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "رجوع", 14175, 1372, 907, 482, "secondary")
    c.OnClick = EP
    Set c = AddList("lstResults", 227, 2098, 14855, 6861, 6, "", True)
    c.OnDblClick = EP
    Set c = AddLabel("lblCount", " ", 227, 9100, 14855, 340, 10, False, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
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
    FinishForm "frmSearch", s
    Exit Sub
EH:
    AbortForm "frmSearch", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmReportCenter()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmReportCenter", "التقارير", "", 15309, 9639, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE8A5), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "مركز التقارير", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "اختر التقرير ثم حدد الفترة أو العميل أو المنتج", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddList("lstReports", 227, 1077, 5103, 8278, 2, "0;4876", False)
    c.RowSourceType = "Value List"
    c.FontSize = 11
    c.AfterUpdate = EP
    c.OnDblClick = EP
    Set c = AddLabel("lblReportTitle", " ", 5670, 1077, 9412, 482, 16, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblNeeds", " ", 5670, 1588, 9412, 340, 10, False, CLR_MUTED, "", 0)
    Set c = AddText("txtFrom", "", 5670, 2325, 2268, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblFrom", "من تاريخ", 5670, 2013, 2268, 284, 9, False, CLR_MUTED, "txtFrom", 0)
    Set c = AddText("txtTo", "", 8165, 2325, 2268, 454)
    SetCtlProp c, "Format", "yyyy/mm/dd"
    Set c = AddLabel("lblTo", "إلى تاريخ", 8165, 2013, 2268, 284, 9, False, CLR_MUTED, "txtTo", 0)
    Set c = AddButton("btnToday", "اليوم", 5670, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisMonth", "هذا الشهر", 7428, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnLastMonth", "الشهر الماضي", 9186, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnThisYear", "هذه السنة", 10944, 2948, 1644, 425, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboCustomer", "", 5670, 3912, 5103, 454, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5103")
    Set c = AddLabel("lblCustomer", "العميل", 5670, 3600, 5103, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddCombo("cboSupplier", "", 5670, 4734, 5103, 454, "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", 2, "0;5103")
    Set c = AddLabel("lblSupplier", "المورد", 5670, 4422, 5103, 284, 9, False, CLR_MUTED, "cboSupplier", 0)
    Set c = AddCombo("cboProduct", "", 5670, 5556, 5103, 454, "SELECT ProductID, ProductName & ' (' & ProductCode & ')' AS Item FROM Products ORDER BY ProductName", 2, "0;5103")
    Set c = AddLabel("lblProduct", "المنتج", 5670, 5244, 5103, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddButton("btnRun", "عرض التقرير", 5670, 6548, 2835, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "رجوع", 13948, 6548, 1134, 567, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblPhaseNote", "التقارير المنسقة للطباعة تُضاف في المرحلة 8، وحتى ذلك الحين تُعرض النتائج كجدول.", 5670, 7342, 9412, 567, 9, False, CLR_MUTED, "", 0)
    m_frm.OnLoad = EP
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
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishForm "frmReportCenter", s
    Exit Sub
EH:
    AbortForm "frmReportCenter", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmPOSLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmPOSLines", "أسطر الفاتورة", "SELECT * FROM tmpPOSLines ORDER BY LineNo", 12020, 425, False, False, True, _
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
    StartForm "frmPOS", "نقطة البيع", "", 18994, 10546, False, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 18994, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "نقطة البيع", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "F9 حفظ  |  F12 حفظ وطباعة  |  F5 فاتورة جديدة  |  F4 بحث بالاسم  |  F8 المبلغ المدفوع  |  F2 الباركود", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtBarcode", "", 227, 1304, 4536, 567)
    c.FontSize = 16
    c.OnKeyDown = EP
    Set c = AddLabel("lblBarcode", "الباركود أو كود المنتج (Enter)", 227, 992, 4536, 284, 9, False, CLR_MUTED, "txtBarcode", 0)
    Set c = AddText("txtQty", "", 4933, 1304, 1134, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.###"
    SetCtlProp c, "DefaultValue", "1"
    Set c = AddLabel("lblQty", "الكمية", 4933, 992, 1134, 284, 9, False, CLR_MUTED, "txtQty", 0)
    Set c = AddCombo("cboProduct", "", 6237, 1304, 6010, 567, "SELECT ProductID, ProductName & ' - ' & ProductCode AS Item, SellingPrice FROM Products WHERE IsActive = True ORDER BY ProductName", 3, "0;4536;1134")
    c.FontSize = 13
    c.AfterUpdate = EP
    Set c = AddLabel("lblProduct", "أو ابحث باسم المنتج (F4)", 6237, 992, 6010, 284, 9, False, CLR_MUTED, "cboProduct", 0)
    Set c = AddLabel("lblCol1", "#", 255, 2041, 567, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "الصنف", 850, 2041, 4252, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "الكمية", 5130, 2041, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "السعر", 6405, 2041, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "خصم السطر", 7907, 2041, 1247, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "الإجمالي", 9182, 2041, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "المتوفر", 10798, 2041, 907, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subLines", "frmPOSLines", 227, 2410, 12020, 6237)
    Set c = AddLabel("lblStatus", " ", 227, 8760, 12020, 397, 11, True, CLR_MUTED, "", 0)
    Set c = AddButton("btnSave", "حفظ (F9)", 227, 9412, 1928, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSavePrint", "حفظ وطباعة (F12)", 2268, 9412, 2381, 624, "primary")
    c.OnClick = EP
    Set c = AddButton("btnNewSale", "فاتورة جديدة (F5)", 4762, 9412, 2155, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "مرتجع", 7030, 9412, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnPayment", "سند قبض", 8617, 9412, 1474, 624, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReprint", "إعادة طباعة", 10204, 9412, 1588, 624, "secondary")
    c.OnClick = EP
    Set c = AddCombo("cboCustomer", "", 12474, 1304, 6294, 454, "SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustomer", "العميل", 12474, 992, 6294, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddLabel("lblCustomerInfo", " ", 12474, 1786, 6294, 284, 9, False, CLR_MUTED, "", 0)
    Set c = AddCombo("cboPaymentType", "", 12474, 2438, 3033, 454, "CASH;نقدي;CREDIT;آجل", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentType", "نوع البيع", 12474, 2126, 3033, 284, 9, False, CLR_MUTED, "cboPaymentType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 15735, 2438, 3033, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;2835")
    c.AfterUpdate = EP
    Set c = AddLabel("lblPaymentMethod", "طريقة الدفع", 15735, 2126, 3033, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtInvoiceDiscount", "", 12474, 3260, 3033, 454)
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceDiscount", "خصم على الفاتورة", 12474, 2948, 3033, 284, 9, False, CLR_MUTED, "txtInvoiceDiscount", 0)
    Set c = AddText("txtNotes", "", 15735, 3260, 3033, 454)
    Set c = AddLabel("lblNotes", "ملاحظات", 15735, 2948, 3033, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddRect("boxTotals", 12474, 3884, 6294, 3289, CLR_SURFACE)
    Set c = AddLabel("lblCapSubTotal", "المجموع قبل الخصم والضريبة", 12644, 3969, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblSubTotal", "0.00", 16386, 3969, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapDiscount", "الخصم", 12644, 4366, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblDiscount", "0.00", 16386, 4366, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTax", "ضريبة القيمة المضافة", 12644, 4763, 3686, 340, 11, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblTax", "0.00", 16386, 4763, 2211, 340, 12, True, CLR_TEXT, "", 3)
    Set c = AddLabel("lblCapTotal", "الإجمالي شامل الضريبة", 12644, 5245, 5954, 340, 12, True, CLR_PRIMARY, "", 0)
    Set c = AddLabel("lblTotal", "0.00", 12644, 5613, 5954, 964, 34, True, CLR_ACCENT, "", 2)
    Set c = AddLabel("lblItems", " ", 12644, 6662, 5954, 340, 10, False, CLR_MUTED, "", 2)
    Set c = AddText("txtTendered", "", 12474, 7598, 6294, 567)
    c.FontSize = 16
    SetCtlProp c, "Format", "#,##0.00"
    c.AfterUpdate = EP
    Set c = AddLabel("lblTendered", "المبلغ المدفوع (F8) - اتركه فارغًا إذا دفع المبلغ بالضبط", 12474, 7286, 6294, 284, 9, False, CLR_MUTED, "txtTendered", 0)
    Set c = AddLabel("lblChange", " ", 12474, 8250, 6294, 454, 14, True, CLR_SUCCESS, "", 2)
    Set c = AddLabel("lblLastInvoiceCap", "آخر فاتورة:", 12474, 8817, 1701, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblLastInvoice", " ", 14232, 8817, 2835, 312, 9, False, CLR_MUTED, "", 0)
    Set c = AddButton("btnClose", "إغلاق", 17067, 9412, 1701, 624, "secondary")
    c.OnClick = EP
    m_frm.OnLoad = EP
    m_frm.OnKeyDown = EP
    m_frm.OnUnload = EP
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
    FinishForm "frmPOS", s
    Exit Sub
EH:
    AbortForm "frmPOS", Err.Number, Err.Description
End Sub

Private Sub BuildForm_frmReturnLines()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartForm "frmReturnLines", "أسطر المرتجع", "SELECT * FROM tmpReturnLines ORDER BY SalesDetailID", 14855, 425, False, False, True, _
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
    StartForm "frmSalesReturn", "مرتجع مبيعات", "", 15309, 9639, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "مرتجع مبيعات (إشعار دائن)", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "اختر الفاتورة الأصلية ثم حدد الكميات المرتجعة", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddText("txtInvoiceNo", "", 227, 1304, 2835, 454)
    c.FontSize = 12
    c.AfterUpdate = EP
    Set c = AddLabel("lblInvoiceNo", "رقم الفاتورة الأصلية", 227, 992, 2835, 284, 9, False, CLR_MUTED, "txtInvoiceNo", 0)
    Set c = AddButton("btnFind", "بحث", 3175, 1304, 1134, 454, "secondary")
    c.OnClick = EP
    Set c = AddLabel("lblInvoiceInfo", " ", 4479, 1361, 10603, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddLabel("lblCol1", "الصنف", 255, 1956, 5103, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol2", "الكمية المباعة", 5386, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol3", "مرتجع سابقًا", 6888, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol4", "المتاح للإرجاع", 8390, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol5", "الكمية المرتجعة", 9892, 1956, 1588, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol6", "يعود للمخزون", 11508, 1956, 1474, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddLabel("lblCol7", "قيمة المرتجع", 13010, 1956, 1928, 312, 9, True, CLR_MUTED, "", 2)
    Set c = AddSubform("subReturnLines", "frmReturnLines", 227, 2296, 14855, 4309)
    Set c = AddText("txtReason", "", 227, 7002, 6577, 454)
    Set c = AddLabel("lblReason", "سبب الإرجاع *", 227, 6690, 6577, 284, 9, False, CLR_MUTED, "txtReason", 0)
    Set c = AddCombo("cboRefundType", "", 6974, 7002, 3742, 454, "CASH;رد نقدي;CREDIT;خصم من رصيد العميل", 2, "0;3402")
    Set c = AddLabel("lblRefundType", "طريقة رد المبلغ", 6974, 6690, 3742, 284, 9, False, CLR_MUTED, "cboRefundType", 0)
    Set c = AddCombo("cboPaymentMethod", "", 10886, 7002, 4196, 454, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "طريقة الرد", 10886, 6690, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddLabel("lblReturnTotalCap", "قيمة المرتجع:", 227, 7711, 2268, 454, 13, False, CLR_MUTED, "", 0)
    Set c = AddLabel("lblReturnTotal", "0.00", 2552, 7598, 3402, 624, 20, True, CLR_DANGER, "", 0)
    Set c = AddButton("btnReturnAll", "إرجاع الكل", 227, 8675, 1814, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturn", "حفظ المرتجع", 2154, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnSaveReturnPrint", "حفظ وطباعة", 4308, 8675, 2041, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13608, 8675, 1474, 567, "secondary")
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
    StartForm "frmCustomerPayment", "سند قبض", "", 9072, 5783, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 9072, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE716), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "سند قبض", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "تسجيل دفعة من عميل على حسابه", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddCombo("cboCustomer", "", 227, 1361, 8618, 454, "SELECT CustomerID, CustomerName FROM Customers WHERE IsSystem = False AND IsActive = True ORDER BY CustomerName", 2, "0;5670")
    c.AfterUpdate = EP
    Set c = AddLabel("lblCustomer", "العميل", 227, 1049, 8618, 284, 9, False, CLR_MUTED, "cboCustomer", 0)
    Set c = AddLabel("lblBalance", " ", 227, 1871, 8618, 340, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddText("txtAmount", "", 227, 2608, 4196, 510)
    c.FontSize = 14
    SetCtlProp c, "Format", "#,##0.00"
    Set c = AddLabel("lblAmount", "المبلغ *", 227, 2296, 4196, 284, 9, False, CLR_MUTED, "txtAmount", 0)
    Set c = AddCombo("cboPaymentMethod", "", 4649, 2608, 4196, 510, "SELECT PaymentMethodID, MethodName FROM PaymentMethods WHERE IsActive = True ORDER BY SortOrder", 2, "0;3402")
    Set c = AddLabel("lblPaymentMethod", "طريقة الدفع", 4649, 2296, 4196, 284, 9, False, CLR_MUTED, "cboPaymentMethod", 0)
    Set c = AddText("txtNotes", "", 227, 3515, 8618, 454)
    Set c = AddLabel("lblNotes", "ملاحظات", 227, 3203, 8618, 284, 9, False, CLR_MUTED, "txtNotes", 0)
    Set c = AddButton("btnSave", "حفظ السند", 227, 4763, 2268, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 7371, 4763, 1474, 567, "secondary")
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
    StartForm "frmSalesInvoice", "فاتورة بيع", "", 15309, 8845, True, False, True, _
              ""
    Set c = AddRect("boxTitle", 0, 0, 15309, 850, CLR_PRIMARY)
    Set c = AddIcon("icoTitle", ChrW(&HE7BF), 227, 170, 510, 510, 20, False, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblTitle", "فاتورة بيع", 850, 102, 7938, 425, 16, True, CLR_SURFACE, "", 0)
    Set c = AddLabel("lblSubtitle", "عرض فقط - التصحيح يكون بمرتجع", 850, 510, 7938, 284, 9, False, CLR_SIDEBAR_TEXT, "", 0)
    Set c = AddText("txtInvoiceID", "", 14742, 57, 340, 227)
    SetCtlProp c, "Visible", False
    Set c = AddLabel("lblHeader", " ", 227, 1021, 14855, 397, 10, False, CLR_TEXT, "", 0)
    Set c = AddList("lstLines", 227, 1531, 14855, 3742, 7, "567;5103;1418;1984;1418;1701;2268", True)
    Set c = AddLabel("lblReturnsCap", "المرتجعات على هذه الفاتورة", 227, 5386, 4536, 312, 10, True, CLR_MUTED, "", 0)
    Set c = AddList("lstReturns", 227, 5727, 14855, 1304, 4, "2268;2835;1984;6804", True)
    Set c = AddLabel("lblTotals", " ", 227, 7144, 14855, 397, 11, True, CLR_PRIMARY, "", 0)
    Set c = AddButton("btnPrint", "طباعة", 227, 7881, 1588, 567, "primary")
    c.OnClick = EP
    Set c = AddButton("btnPrintA4", "طباعة A4", 1928, 7881, 1588, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnReturn", "مرتجع", 3629, 7881, 1474, 567, "secondary")
    c.OnClick = EP
    Set c = AddButton("btnClose", "إغلاق", 13608, 7881, 1474, 567, "secondary")
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
