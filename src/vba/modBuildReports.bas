Attribute VB_Name = "modBuildReports"
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
Private Const REPORT_NAMES As String = "rptSalesReceipt,rptSalesInvoiceA4,rptPurchaseDocument,rptVoucher,rptStockCount,rptBarcodeLabels,rptStatistics,rptDailySales,rptMonthlySales,rptSalesByPeriod,rptSalesByProduct,rptBestSelling,rptLeastSelling,rptPurchases,rptStockBalance,rptLowStock,rptProductMovement,rptCustomerStatement,rptSupplierStatement,rptExpenses,rptExpensesByType,rptSlowMoving,rptStockByCategory,rptCustomerBalances,rptSupplierBalances,rptIntegrityCheck,rptProfit,rptVatSummary"

Public Function BuildReports() As Boolean
    Dim i As Long
    On Error GoTo EH
    m_built = 0: m_failed = 0: m_report = ""
    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name, acSaveNo
    Next
    DoCmd.Echo False, "جاري بناء التقارير..."
    BuildReport_rptSalesReceipt
    BuildReport_rptSalesInvoiceA4
    BuildReport_rptPurchaseDocument
    BuildReport_rptVoucher
    BuildReport_rptStockCount
    BuildReport_rptBarcodeLabels
    BuildReport_rptStatistics
    BuildReport_rptDailySales
    BuildReport_rptMonthlySales
    BuildReport_rptSalesByPeriod
    BuildReport_rptSalesByProduct
    BuildReport_rptBestSelling
    BuildReport_rptLeastSelling
    BuildReport_rptPurchases
    BuildReport_rptStockBalance
    BuildReport_rptLowStock
    BuildReport_rptProductMovement
    BuildReport_rptCustomerStatement
    BuildReport_rptSupplierStatement
    BuildReport_rptExpenses
    BuildReport_rptExpensesByType
    BuildReport_rptSlowMoving
    BuildReport_rptStockByCategory
    BuildReport_rptCustomerBalances
    BuildReport_rptSupplierBalances
    BuildReport_rptIntegrityCheck
    BuildReport_rptProfit
    BuildReport_rptVatSummary
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
    RecordR AmountInWords(0.5) = "فقط خمسون هللة لا غير", "المبلغ بالحروف: 0.5"
    RecordR AmountInWords(1) = "فقط واحد ريال سعودي لا غير", "المبلغ بالحروف: 1"
    RecordR AmountInWords(2) = "فقط اثنان ريال سعودي لا غير", "المبلغ بالحروف: 2"
    RecordR AmountInWords(3) = "فقط ثلاثة ريال سعودي لا غير", "المبلغ بالحروف: 3"
    RecordR AmountInWords(10) = "فقط عشرة ريال سعودي لا غير", "المبلغ بالحروف: 10"
    RecordR AmountInWords(11) = "فقط أحد عشر ريال سعودي لا غير", "المبلغ بالحروف: 11"
    RecordR AmountInWords(12) = "فقط اثنا عشر ريال سعودي لا غير", "المبلغ بالحروف: 12"
    RecordR AmountInWords(19) = "فقط تسعة عشر ريال سعودي لا غير", "المبلغ بالحروف: 19"
    RecordR AmountInWords(20) = "فقط عشرون ريال سعودي لا غير", "المبلغ بالحروف: 20"
    RecordR AmountInWords(21) = "فقط واحد وعشرون ريال سعودي لا غير", "المبلغ بالحروف: 21"
    RecordR AmountInWords(99) = "فقط تسعة وتسعون ريال سعودي لا غير", "المبلغ بالحروف: 99"
    RecordR AmountInWords(100) = "فقط مائة ريال سعودي لا غير", "المبلغ بالحروف: 100"
    RecordR AmountInWords(101) = "فقط مائة وواحد ريال سعودي لا غير", "المبلغ بالحروف: 101"
    RecordR AmountInWords(115) = "فقط مائة وخمسة عشر ريال سعودي لا غير", "المبلغ بالحروف: 115"
    RecordR AmountInWords(200) = "فقط مائتان ريال سعودي لا غير", "المبلغ بالحروف: 200"
    RecordR AmountInWords(999) = "فقط تسعمائة وتسعة وتسعون ريال سعودي لا غير", "المبلغ بالحروف: 999"
    RecordR AmountInWords(1000) = "فقط ألف ريال سعودي لا غير", "المبلغ بالحروف: 1000"
    RecordR AmountInWords(1001) = "فقط ألف وواحد ريال سعودي لا غير", "المبلغ بالحروف: 1001"
    RecordR AmountInWords(1250.5) = "فقط ألف ومائتان وخمسون ريال سعودي وخمسون هللة لا غير", "المبلغ بالحروف: 1250.5"
    RecordR AmountInWords(2000) = "فقط ألفان ريال سعودي لا غير", "المبلغ بالحروف: 2000"
    RecordR AmountInWords(2500) = "فقط ألفان وخمسمائة ريال سعودي لا غير", "المبلغ بالحروف: 2500"
    RecordR AmountInWords(3000) = "فقط ثلاثة آلاف ريال سعودي لا غير", "المبلغ بالحروف: 3000"
    RecordR AmountInWords(10000) = "فقط عشرة آلاف ريال سعودي لا غير", "المبلغ بالحروف: 10000"
    RecordR AmountInWords(11000) = "فقط أحد عشر ألف ريال سعودي لا غير", "المبلغ بالحروف: 11000"
    RecordR AmountInWords(12345.67) = "فقط اثنا عشر ألف وثلاثمائة وخمسة وأربعون ريال سعودي وسبعة وستون هللة لا غير", "المبلغ بالحروف: 12345.67"
    RecordR AmountInWords(100000) = "فقط مائة ألف ريال سعودي لا غير", "المبلغ بالحروف: 100000"
    RecordR AmountInWords(101000) = "فقط مائة وواحد ألف ريال سعودي لا غير", "المبلغ بالحروف: 101000"
    RecordR AmountInWords(999999.99) = "فقط تسعمائة وتسعة وتسعون ألف وتسعمائة وتسعة وتسعون ريال سعودي وتسعة وتسعون هللة لا غير", "المبلغ بالحروف: 999999.99"
    RecordR AmountInWords(1000000) = "فقط مليون ريال سعودي لا غير", "المبلغ بالحروف: 1000000"
    RecordR AmountInWords(2000000) = "فقط مليونان ريال سعودي لا غير", "المبلغ بالحروف: 2000000"
    RecordR AmountInWords(3500000) = "فقط ثلاثة ملايين وخمسمائة ألف ريال سعودي لا غير", "المبلغ بالحروف: 3500000"
    RecordR AmountInWords(11000000) = "فقط أحد عشر مليون ريال سعودي لا غير", "المبلغ بالحروف: 11000000"
    RecordR AmountInWords(123456789.01) = "فقط مائة وثلاثة وعشرون مليون وأربعمائة وستة وخمسون ألف وسبعمائة وتسعة وثمانون ريال سعودي وواحد هللة لا غير", "المبلغ بالحروف: 123456789.01"
    RecordR AmountInWords(0.05) = "فقط خمسة هللة لا غير", "المبلغ بالحروف: 0.05"
    RecordR AmountInWords(0.11) = "فقط أحد عشر هللة لا غير", "المبلغ بالحروف: 0.11"
    RecordR AmountInWords(4750) = "فقط أربعة آلاف وسبعمائة وخمسون ريال سعودي لا غير", "المبلغ بالحروف: 4750"
    RecordR AmountInWords(3175) = "فقط ثلاثة آلاف ومائة وخمسة وسبعون ريال سعودي لا غير", "المبلغ بالحروف: 3175"
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
Private Sub BuildReport_rptSalesReceipt()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesReceipt", "فاتورة (حراري 80 مم)", "qrySalesDocPrint", 4196, "DocID", "LineNumber", False, False
    SetSection 0, 539
    SetSection 5, 3827
    SetSection 6, 4734
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtProduct", "ProductName", 0, 0, 4196, 255, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtQtyPrice", "=[Quantity] & "" × "" & Format([LineTotal]/[Quantity],""#,##0.00"")", 113, 266, 2835, 238, 8, False, 0)
    Set c = RText(0, "txtLineTotal", "LineTotal", 2948, 266, 1247, 238, 9, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 4196, 340, 12, True, 2)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 397, 4196, 255, 9, False, 2)
    Set c = RText(5, "txtStoreVat", "=""الرقم الضريبي: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 652, 4196, 227, 9, True, 2)
    Set c = RText(5, "txtStoreCR", "=""س.ت: "" & Nz(SettingValue(""CRNumber""),""-"") & ""   هاتف: "" & Nz(SettingValue(""Phone""),""-"")", 0, 879, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""District""),"""") & "" "" & Nz(SettingValue(""StreetName""),""""))", 0, 1106, 4196, 227, 8, False, 2)
    Set c = RLine(5, "lnHeader1", 1361, 4196)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""إشعار دائن"",IIf([InvoiceSubType]=""STANDARD"",""فاتورة ضريبية"",""فاتورة ضريبية مبسطة""))", 0, 1418, 4196, 312, 12, True, 2)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 0, 1730, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtDocNumber", "=""رقم: "" & [DocNumber]", 0, 1985, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtDocDate", "=""التاريخ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 0, 2212, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""عن الفاتورة: "" & [OriginalNumber],OrderTypeText([OrderType],[TableNo]))", 0, 2439, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCashier", "=""الكاشير: "" & [EmployeeName]", 0, 2666, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomer", "=IIf([CustomerID]=Nz(SettingValue(""DefaultCustomerID""),1),"""",""العميل: "" & [CustomerName])", 0, 2893, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""الرقم الضريبي للعميل: "" & [CustomerVAT],"""")", 0, 3120, 4196, 227, 8, False, 0)
    Set c = RLine(5, "lnHeader2", 3375, 4196)
    Set c = RLabel(5, "lblColItem", "الصنف", 0, 3404, 2211, 255, 8, True, 0)
    Set c = RLabel(5, "lblColQty", "الكمية × السعر", 2211, 3404, 1134, 255, 8, True, 0)
    Set c = RLabel(5, "lblColTotal", "الإجمالي", 3345, 3404, 850, 255, 8, True, 1)
    Set c = RLine(5, "lnHeader3", 3688, 4196)
    Set c = RLine(6, "lnTotals1", 45, 4196)
    Set c = RLabel(6, "lblCapTaxableAmount", "الإجمالي بدون الضريبة", 0, 113, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 2665, 113, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "الخصم", 0, 397, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 2665, 397, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapDocTax", "=""ضريبة القيمة المضافة "" & Format(Nz(SettingValue(""VATRate""),0.15),""0%"")", 0, 681, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 2665, 681, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "الإجمالي شامل الضريبة", 0, 965, 2608, 255, 11, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 2665, 965, 1531, 255, 11, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(6, "lnTotals2", 1277, 4196)
    Set c = RText(6, "txtCapPaid", "=IIf([DocKind]=""RETURN"",""المبلغ المردود"",""المدفوع"")", 0, 1334, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValPaid", "PaidAmount", 2665, 1334, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapChange", "=IIf([ChangeDue]>0,""الباقي للعميل"","""")", 0, 1589, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValChange", "=IIf([ChangeDue]>0,Format([ChangeDue],""#,##0.00""),"""")", 2665, 1589, 1531, 255, 9, False, 1)
    Set c = RText(6, "txtCapRemaining", "=IIf([RemainingAmount]<>0,IIf([DocKind]=""RETURN"",""خصم من رصيد العميل"",""المتبقي على الحساب""),"""")", 0, 1844, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValRemaining", "=IIf([RemainingAmount]<>0,Format([RemainingAmount],""#,##0.00""),"""")", 2665, 1844, 1531, 255, 9, False, 1)
    Set c = RBox(6, "boxQR", 1077, 2184, 2041, 2041)
    Set c = RText(6, "txtDocKind", "DocKind", 0, 2184, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtDocID", "DocID", 0, 2439, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtFooterNote", "=Nz(SettingValue(""ReceiptFooter""),"""")", 0, 4282, 4196, 255, 9, False, 2)
    m_rpt.Section(6).OnPrint = EP
    s = ""
    s = s & "Private Sub secTotals_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub   ' no document: the fields have no value (2427)" & vbCrLf
    s = s & "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _" & vbCrLf
    s = s & "                   Me!boxQR.Top, Me!boxQR.Width" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesReceipt", s
    Exit Sub
EH:
    AbortReport "rptSalesReceipt", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesInvoiceA4()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesInvoiceA4", "فاتورة ضريبية (A4)", "qrySalesDocPrint", 10773, "DocID", "LineNumber", False, False
    SetSection 0, 340
    SetSection 5, 3856
    SetSection 6, 4649
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "LineNumber", 0, 28, 454, 284, 8, False, 2)
    Set c = RText(0, "txtCol2", "ProductName", 454, 28, 3515, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "Quantity", 3969, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol4", "UnitPrice", 4876, 28, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "LineDiscount", 6010, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetAmount", 6917, 28, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "VATRate", 8051, 28, 680, 284, 8, False, 2)
    SetCtl c, "Format", "0%"
    Set c = RText(0, "txtCol8", "LineTax", 8731, 28, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "LineTotal", 9638, 28, 1135, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 397, 14, True, 0)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 454, 5273, 284, 10, False, 0)
    Set c = RText(5, "txtStoreVat", "=""الرقم الضريبي: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 738, 5273, 284, 10, True, 0)
    Set c = RText(5, "txtStoreCR", "=""السجل التجاري: "" & Nz(SettingValue(""CRNumber""),"""")", 0, 1022, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""BuildingNo""),"""") & "" "" & Nz(SettingValue(""StreetName""),"""") & "" - "" & Nz(SettingValue(""District""),"""") & "" - "" & Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""PostalCode""),""""))", 0, 1306, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""هاتف: "" & Nz(SettingValue(""Phone""),"""")", 0, 1590, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""إشعار دائن"",IIf([InvoiceSubType]=""STANDARD"",""فاتورة ضريبية"",""فاتورة ضريبية مبسطة""))", 5500, 57, 5273, 454, 16, True, 1)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 5500, 539, 5273, 284, 10, False, 1)
    Set c = RText(5, "txtDocNumber", "=""رقم الفاتورة: "" & [DocNumber]", 5500, 850, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtDocDate", "=""التاريخ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 5500, 1105, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""عن الفاتورة: "" & [OriginalNumber],OrderTypeText([OrderType],[TableNo]))", 5500, 1360, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtPaymentType", "=""طريقة البيع: "" & IIf([PaymentType]=""CREDIT"",""آجل"",""نقدي"")", 5500, 1615, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtCashier", "=""الموظف: "" & [EmployeeName]", 5500, 1870, 5273, 255, 9, False, 1)
    Set c = RBox(5, "boxBuyer", 0, 2183, 10773, 1106)
    SetCtl c, "BorderStyle", 1
    Set c = RLabel(5, "lblBuyer", "بيانات المشتري", 113, 2211, 2268, 255, 9, True, 0)
    Set c = RText(5, "txtBuyerName", "=""الاسم: "" & [CustomerName]", 113, 2467, 5103, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""الرقم الضريبي: "" & [CustomerVAT],"""")", 5443, 2467, 5216, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerAddress", "=Trim(Nz([CustomerBuilding],"""") & "" "" & Nz([CustomerStreet],"""") & "" "" & Nz([CustomerDistrict],"""") & "" "" & Nz([CustomerCity],"""") & "" "" & Nz([CustomerPostal],""""))", 113, 2750, 10546, 255, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 3430, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 3487, 454, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "الصنف", 454, 3487, 3515, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "الكمية", 3969, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "سعر الوحدة", 4876, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "الخصم", 6010, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "الصافي", 6917, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "نسبة الضريبة", 8051, 3487, 680, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol8", "الضريبة", 8731, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol9", "الإجمالي", 9638, 3487, 1135, 255, 8, True, 2)
    Set c = RLine(6, "lnTotals1", 45, 10773)
    Set c = RLabel(6, "lblCapDocSubTotal", "الإجمالي قبل الخصم", 6464, 113, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocSubTotal", "DocSubTotal", 8959, 113, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "الخصم", 6464, 397, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 8959, 397, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTaxableAmount", "الإجمالي الخاضع للضريبة", 6464, 681, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 8959, 681, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocTax", "ضريبة القيمة المضافة", 6464, 965, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 8959, 965, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "الإجمالي شامل الضريبة", 6464, 1249, 2495, 255, 10, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 8959, 1249, 1814, 255, 10, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapPaidAmount", "المدفوع / المردود", 6464, 1533, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtPaidAmount", "PaidAmount", 8959, 1533, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapRemainingAmount", "المتبقي", 6464, 1817, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtRemainingAmount", "RemainingAmount", 8959, 1817, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RBox(6, "boxQR", 4366, 2211, 2041, 2041)
    Set c = RText(6, "txtDocKind", "DocKind", 0, 2211, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtDocID", "DocID", 0, 2466, 284, 227, 9, False, 0)
    SetCtl c, "Visible", False
    Set c = RText(6, "txtFooterNote", "=Nz(SettingValue(""ReceiptFooter""),"""")", 0, 4309, 10773, 255, 9, False, 2)
    m_rpt.Section(6).OnPrint = EP
    s = ""
    s = s & "Private Sub secTotals_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub   ' no document: the fields have no value (2427)" & vbCrLf
    s = s & "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _" & vbCrLf
    s = s & "                   Me!boxQR.Top, Me!boxQR.Width" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesInvoiceA4", s
    Exit Sub
EH:
    AbortReport "rptSalesInvoiceA4", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptPurchaseDocument()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptPurchaseDocument", "فاتورة / مرتجع مشتريات", "qryPurchaseDocPrint", 10773, "DocID", "LineNumber", False, True
    SetSection 0, 340
    SetSection 5, 3345
    SetSection 6, 2948
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "LineNumber", 0, 23, 454, 284, 8, False, 2)
    Set c = RText(0, "txtCol2", "ProductCode", 454, 23, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ProductName", 1701, 23, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol4", "UnitName", 4536, 23, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "Quantity", 5330, 23, 850, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "UnitCost", 6180, 23, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00##"
    Set c = RText(0, "txtCol7", "LineDiscount", 7314, 23, 850, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "LineTax", 8164, 23, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "LineTotal", 9071, 23, 1702, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""الرقم الضريبي: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""هاتف: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""مرتجع مشتريات"",""فاتورة مشتريات"")", 5500, 57, 5273, 482, 16, True, 1)
    Set c = RText(5, "txtDocNumber", "=""الرقم: "" & [DocNumber]", 5500, 567, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtDocDate", "=""التاريخ: "" & GDate([DocDate],True)", 5500, 822, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtSupplierRef", "=""رقم فاتورة المورد: "" & Nz([SupplierInvoiceNo],""-"")", 5500, 1077, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""عن فاتورة الشراء: "" & [OriginalNumber],"""")", 5500, 1332, 5273, 255, 9, False, 1)
    Set c = RBox(5, "boxSupplier", 0, 1701, 10773, 1106)
    SetCtl c, "BorderStyle", 1
    Set c = RText(5, "txtSupplier", "=""المورد: "" & [SupplierName]", 113, 1758, 5103, 284, 10, True, 0)
    Set c = RText(5, "txtSupplierVat", "=IIf(Len(Nz([SupplierVAT],""""))>0,""الرقم الضريبي: "" & [SupplierVAT],"""")", 5330, 1758, 5330, 284, 9, False, 0)
    Set c = RText(5, "txtPayment", "=IIf([DocKind]=""RETURN"",IIf([PaymentType]=""CASH"",""استرداد نقدي"",""خصم من رصيد المورد""),IIf([PaymentType]=""CREDIT"",""شراء آجل"",""شراء نقدي""))", 113, 2098, 5103, 284, 9, False, 0)
    Set c = RText(5, "txtEmployee", "=""الموظف: "" & [EmployeeName]", 5330, 2098, 5330, 284, 9, False, 0)
    Set c = RText(5, "txtReason", "=IIf(Len(Nz([Reason],""""))>0,""سبب الإرجاع: "" & [Reason],"""")", 113, 2438, 10546, 284, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 2920, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 2965, 454, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "الكود", 454, 2965, 1247, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "الصنف", 1701, 2965, 2835, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "الوحدة", 4536, 2965, 794, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "الكمية", 5330, 2965, 850, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "تكلفة الوحدة", 6180, 2965, 1134, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "الخصم", 7314, 2965, 850, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol8", "الضريبة", 8164, 2965, 907, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol9", "الإجمالي", 9071, 2965, 1702, 284, 8, True, 2)
    Set c = RLine(6, "lnTotals", 45, 10773)
    Set c = RLabel(6, "lblCapDocSubTotal", "الإجمالي قبل الخصم", 6464, 113, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocSubTotal", "DocSubTotal", 8959, 113, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "الخصم", 6464, 397, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 8959, 397, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTaxableAmount", "الخاضع للضريبة", 6464, 681, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 8959, 681, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocTax", "ضريبة القيمة المضافة", 6464, 965, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 8959, 965, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "الإجمالي شامل الضريبة", 6464, 1249, 2495, 255, 10, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 8959, 1249, 1814, 255, 10, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapPaid", "=IIf([DocKind]=""RETURN"",""المسترد نقدًا"",""المدفوع"")", 6464, 1533, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtPaidAmount", "PaidAmount", 8959, 1533, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapRemaining", "=IIf([DocKind]=""RETURN"",""خصم من رصيد المورد"",""المتبقي على الحساب"")", 6464, 1817, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtRemainingAmount", "RemainingAmount", 8959, 1817, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtWords", "=AmountInWords([TotalAmount])", 0, 113, 6237, 567, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RLabel(6, "lblSign1", "المستلم: ....................", 0, 2438, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "أمين المخزن: ....................", 3591, 2438, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "المدير: ....................", 7182, 2438, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""المستند غير موجود.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptPurchaseDocument", s
    Exit Sub
EH:
    AbortReport "rptPurchaseDocument", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptVoucher()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptVoucher", "سند قبض / سند صرف", "qryVoucherPrint", 10773, "DocID", "LineNumber", False, True
    SetSection 0, 0
    SetSection 5, 3969
    SetSection 6, 907
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5386, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""الرقم الضريبي: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""هاتف: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5386, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RECEIPT"",""سند قبض"",""سند صرف"")", 5386, 57, 5387, 539, 18, True, 1)
    Set c = RText(5, "txtDocNumber", "=""رقم السند: "" & [DocNumber]", 5386, 624, 5387, 284, 10, False, 1)
    Set c = RText(5, "txtDocDate", "=""التاريخ: "" & GDate([DocDate],True)", 5386, 936, 5387, 284, 10, False, 1)
    Set c = RLine(5, "lnTop", 1304, 10773)
    Set c = RText(5, "txtParty", "=IIf([DocKind]=""RECEIPT"",""استلمنا من: "",""صرفنا إلى: "") & [PartyName]", 0, 1474, 10773, 397, 13, True, 0)
    Set c = RBox(5, "boxAmount", 0, 1984, 3402, 680)
    SetCtl c, "BorderStyle", 1
    Set c = RText(5, "txtAmount", "=Format([Amount],""#,##0.00"") & "" ريال""", 57, 2041, 3289, 567, 18, True, 2)
    Set c = RText(5, "txtWords", "=AmountInWords([Amount])", 3572, 2070, 7201, 567, 11, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtMethod", "=""طريقة الدفع: "" & [MethodName]", 0, 2835, 5386, 284, 10, False, 0)
    Set c = RText(5, "txtEmployee", "=""الموظف: "" & [EmployeeName]", 5386, 2835, 5387, 284, 10, False, 0)
    Set c = RText(5, "txtNotes", "=""البيان: "" & Nz([Notes],""-"")", 0, 3175, 10773, 312, 10, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(5, "txtBalance", "=IIf([DocKind]=""RECEIPT"",""الرصيد المتبقي على العميل حاليًا: "",""الرصيد المستحق للمورد حاليًا: "") & Format([PartyBalance],""#,##0.00"")", 0, 3572, 10773, 284, 9, False, 0)
    Set c = RLabel(6, "lblSign1", "المستلم: ....................", 0, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "المحاسب: ....................", 3591, 284, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "المدير: ....................", 7182, 284, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""السند غير موجود.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptVoucher", s
    Exit Sub
EH:
    AbortReport "rptVoucher", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockCount()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockCount", "ورقة الجرد", "StockCountQuery", 10773, "StockCountID", "ProductName", False, True
    SetSection 0, 340
    SetSection 5, 2041
    SetSection 6, 1474
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtCol1", "=1", 0, 23, 510, 284, 8, False, 0)
    SetCtl c, "RunningSum", 1
    Set c = RText(0, "txtCol2", "ProductCode", 510, 23, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ProductName", 1871, 23, 3402, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol4", "SystemQuantity", 5273, 23, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "ActualQuantity", 6520, 23, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "=IIf(IsNull([ActualQuantity]),Null,[Difference])", 7824, 23, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "=IIf(IsNull([ActualQuantity]),Null,[DifferenceValue])", 8958, 23, 1815, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 425, 14, True, 0)
    Set c = RText(5, "txtStoreVat", "=""الرقم الضريبي: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 510, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""هاتف: "" & Nz(SettingValue(""Phone""),""-"") & ""   "" & Nz(SettingValue(""City""),"""")", 0, 794, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=""ورقة جرد "" & IIf([Status]=""POSTED"",""(مُرحّل)"",IIf([Status]=""CANCELLED"",""(ملغى)"",""(مفتوح)""))", 5500, 57, 5273, 482, 16, True, 1)
    Set c = RText(5, "txtCountNumber", "=""رقم الجرد: "" & [CountNumber] & ""    التاريخ: "" & GDate([CountDate])", 5500, 567, 5273, 284, 9, False, 1)
    Set c = RText(5, "txtCategory", "=""التصنيف: "" & Nz([CategoryName],""كل المنتجات"")", 5500, 879, 5273, 284, 9, False, 1)
    Set c = RBox(5, "boxColumns", 0, 1644, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 1689, 510, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "الكود", 510, 1689, 1361, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "الصنف", 1871, 1689, 3402, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "المسجل", 5273, 1689, 1247, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "الفعلي", 6520, 1689, 1304, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "الفرق", 7824, 1689, 1134, 284, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "قيمة الفرق", 8958, 1689, 1815, 284, 8, True, 2)
    Set c = RLine(6, "lnTotals", 45, 10773)
    Set c = RText(6, "txtSummary", "=""تم عدّ "" & Count([ActualQuantity]) & "" من "" & Count(*) & "" صنف    العجز: "" & Format(-Sum(IIf([DifferenceValue]<0,[DifferenceValue],0)),""#,##0.00"") & ""    الزيادة: "" & Format(Sum(IIf([DifferenceValue]>0,[DifferenceValue],0)),""#,##0.00"")", 0, 113, 10773, 312, 10, True, 0)
    Set c = RLabel(6, "lblSign1", "القائم بالجرد: ....................", 0, 850, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign2", "المراجع: ....................", 3591, 850, 3591, 312, 10, False, 0)
    Set c = RLabel(6, "lblSign3", "المدير: ....................", 7182, 850, 3591, 312, 10, False, 0)
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""جلسة الجرد غير موجودة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockCount", s
    Exit Sub
EH:
    AbortReport "rptStockCount", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBarcodeLabels()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    Application.Run "EnsureLabelTables"
    StartReport "rptBarcodeLabels", "ملصقات الباركود", "SELECT q.LineNo, n.N, p.ProductName, p.ProductCode, p.Barcode, p.SellingPrice FROM tmpLabelQueue AS q, tmpLabelNumbers AS n, Products AS p WHERE p.ProductID = q.ProductID AND n.N <= q.Copies", 2155, "", "LineNo,N", False, False
    SetSection 0, 1418
    HideSection 1
    HideSection 2
    HideSection 3
    HideSection 4
    Set c = RText(0, "txtTop1", "=""المتجر""", 57, 57, 2041, 227, 7, False, 2)
    Set c = RText(0, "txtTop2", "=[ProductName]", 57, 284, 2041, 227, 7, False, 2)
    Set c = RBox(0, "boxBar", 57, 539, 2041, 510)
    SetCtl c, "BorderStyle", 0
    SetCtl c, "BackStyle", 0
    Set c = RText(0, "txtBottom1", "=LabelCode([Barcode],[ProductCode])", 57, 907, 2041, 227, 7, False, 2)
    Set c = RText(0, "txtBottom2", "=LabelPrice([SellingPrice])", 57, 1134, 2041, 227, 7, True, 2)
    Set c = RText(0, "txtCode", "=LabelCode([Barcode],[ProductCode])", 0, 0, 57, 57, 6, False, 0)
    SetCtl c, "Visible", False
    m_rpt.OnNoData = EP
    m_rpt.Section(0).Name = "secLabel"
    m_rpt.Section(0).OnPrint = EP
    m_rpt.OnOpen = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""قائمة الملصقات فارغة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub Report_Open(Cancel As Integer)" & vbCrLf
    s = s & "    LabelReportOpen" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub secLabel_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    If Me.HasData = 0 Then Exit Sub" & vbCrLf
    s = s & "    DrawBarcode Me, Nz(Me!txtCode.Value, """"), Me!boxBar.Left, Me!boxBar.Top, Me!boxBar.Width, _" & vbCrLf
    s = s & "                Me!boxBar.Height, LabelBarWidth()" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBarcodeLabels", s
    Exit Sub
EH:
    AbortReport "rptBarcodeLabels", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStatistics()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStatistics", "الإحصائيات والرسوم البيانية", "", 10773, "", "", False, True
    SetSection 3, 1134
    SetSection 4, 397
    SetSection 0, 13381
    HideSection 1
    HideSection 2
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 5500, 57, 5273, 397, 13, True, 3)
    Set c = RLabel(3, "lblTitle", "الإحصائيات والرسوم البيانية", 0, 57, 5273, 425, 16, True, 0)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 539, 10773, 284, 9, False, 2)
    Set c = RLine(3, "lnHeader", 964, 10773)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6237, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6237, 57, 4536, 255, 8, False, 3)
    Set c = RBox(0, "boxCharts", 0, 0, 10773, 13381)
    SetCtl c, "BorderStyle", 0
    SetCtl c, "BackStyle", 0
    m_rpt.Section(0).Name = "secCharts"
    m_rpt.Section(0).OnPrint = EP
    s = ""
    s = s & "Private Sub secCharts_Print(Cancel As Integer, PrintCount As Integer)" & vbCrLf
    s = s & "    DrawStatistics Me, Me!boxCharts.Left, Me!boxCharts.Top, Me!boxCharts.Width, Me!boxCharts.Height" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStatistics", s
    Exit Sub
EH:
    AbortReport "rptStatistics", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptDailySales()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptDailySales", "المبيعات اليومية", "DailySalesQuery", 15536, "", "SaleDate", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المبيعات اليومية", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "التاريخ", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الفواتير", 1361, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "المرتجعات", 2268, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "المبيعات بدون ضريبة", 3175, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "المرتجعات بدون ضريبة", 4649, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الصافي بدون ضريبة", 6123, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "الضريبة", 7597, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الصافي شامل الضريبة", 8844, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "نقدي", 10375, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "آجل", 11736, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "المحصّل", 13097, 1292, 2439, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([InvoiceCount])", 1361, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([ReturnCount])", 2268, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal4", "=Sum([SalesExVAT])", 3175, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([ReturnsExVAT])", 4649, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesExVAT])", 6123, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([NetVAT])", 7597, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([NetSalesTotal])", 8844, 85, 1531, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([CashSales])", 10375, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CreditSales])", 11736, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([CollectedAmount])", 13097, 85, 2439, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 1361, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([SaleDate])", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "InvoiceCount", 1361, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "ReturnCount", 2268, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol4", "SalesExVAT", 3175, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "ReturnsExVAT", 4649, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetSalesExVAT", 6123, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "NetVAT", 7597, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "NetSalesTotal", 8844, 17, 1531, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "CashSales", 10375, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CreditSales", 11736, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "CollectedAmount", 13097, 17, 2439, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptDailySales", s
    Exit Sub
EH:
    AbortReport "rptDailySales", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptMonthlySales()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptMonthlySales", "المبيعات الشهرية", "MonthlySalesQuery", 10773, "", "SalesYear,SalesMonth", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المبيعات الشهرية", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الشهر", 0, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الفواتير", 1134, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "المرتجعات", 2098, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "الصافي بدون ضريبة", 3062, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الضريبة", 4536, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "شامل الضريبة", 5783, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "التكلفة", 7257, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "مجمل الربح", 8731, 1292, 2042, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([InvoiceCount])", 1134, 85, 964, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([ReturnCount])", 2098, 85, 964, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal4", "=Sum([NetSalesExVAT])", 3062, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([NetVAT])", 4536, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesTotal])", 5783, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([CostOfSales])", 7257, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([GrossProfit])", 8731, 85, 2042, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 1134, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=[SalesYear] & ""/"" & Format([SalesMonth],""00"")", 0, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "InvoiceCount", 1134, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "ReturnCount", 2098, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol4", "NetSalesExVAT", 3062, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "NetVAT", 4536, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "NetSalesTotal", 5783, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "CostOfSales", 7257, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "GrossProfit", 8731, 17, 2042, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptMonthlySales", s
    Exit Sub
EH:
    AbortReport "rptMonthlySales", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesByPeriod()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesByPeriod", "المبيعات حسب فترة", "SalesByPeriodQuery", 15536, "", "DocDate,DocNumber", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المبيعات حسب فترة", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "النوع", 0, 1292, 850, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الرقم", 850, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التاريخ", 2268, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "العميل", 3912, 1292, 2381, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الموظف", 6293, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الدفع", 7767, 1292, 794, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "بدون ضريبة", 8561, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الضريبة", 9922, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "الإجمالي", 11113, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "المدفوع", 12531, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "على الحساب", 13949, 1292, 1587, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal7", "=Sum([NetAmount])", 8561, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([VATAmount])", 9922, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([GrossAmount])", 11113, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([SettledAmount])", 12531, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([OnAccount])", 13949, 85, 1587, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 8561, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=IIf([DocType]=""RETURN"",""مرتجع"",""فاتورة"")", 0, 17, 850, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "DocNumber", 850, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "=GDate([DocDate],True)", 2268, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CustomerName", 3912, 17, 2381, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "EmployeeName", 6293, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "=IIf([PaymentType]=""CREDIT"",""آجل"",""نقدي"")", 7767, 17, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "NetAmount", 8561, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "VATAmount", 9922, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "GrossAmount", 11113, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "SettledAmount", 12531, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "OnAccount", 13949, 17, 1587, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesByPeriod", s
    Exit Sub
EH:
    AbortReport "rptSalesByPeriod", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSalesByProduct()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesByProduct", "المبيعات حسب المنتج", "SalesByProductQuery", 15536, "", "ProductName", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المبيعات حسب المنتج", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1247, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 4309, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "المباع", 5783, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "المرتجع", 6804, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الصافي", 7825, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "المبيعات بدون ضريبة", 8846, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الضريبة", 10264, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "شامل الضريبة", 11455, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "التكلفة", 12873, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "الربح", 14177, 1292, 1359, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtySold])", 5783, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyReturned])", 6804, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetQty])", 7825, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal7", "=Sum([NetSales])", 8846, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([SalesVAT])", 10264, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([SalesTotal])", 11455, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CostOfSales])", 12873, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([GrossProfit])", 14177, 85, 1359, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4309, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtySold", 5783, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyReturned", 6804, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 7825, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "NetSales", 8846, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SalesVAT", 10264, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "SalesTotal", 11455, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CostOfSales", 12873, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "GrossProfit", 14177, 17, 1359, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesByProduct", s
    Exit Sub
EH:
    AbortReport "rptSalesByProduct", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptBestSelling()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptBestSelling", "أفضل المنتجات مبيعًا", "BestSellingProductsQuery", 15536, "", "-NetQty,-NetSales", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "أفضل المنتجات مبيعًا", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1247, 1292, 3062, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 4309, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "المباع", 5783, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "المرتجع", 6804, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الصافي", 7825, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "المبيعات بدون ضريبة", 8846, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الضريبة", 10264, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "شامل الضريبة", 11455, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "التكلفة", 12873, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "الربح", 14177, 1292, 1359, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtySold])", 5783, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyReturned])", 6804, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetQty])", 7825, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal7", "=Sum([NetSales])", 8846, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([SalesVAT])", 10264, 85, 1191, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([SalesTotal])", 11455, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([CostOfSales])", 12873, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([GrossProfit])", 14177, 85, 1359, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 3062, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4309, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtySold", 5783, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyReturned", 6804, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 7825, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "NetSales", 8846, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SalesVAT", 10264, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "SalesTotal", 11455, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "CostOfSales", 12873, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "GrossProfit", 14177, 17, 1359, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptBestSelling", s
    Exit Sub
EH:
    AbortReport "rptBestSelling", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptLeastSelling()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptLeastSelling", "أقل المنتجات مبيعًا", "LeastSellingProductsQuery", 10773, "", "NetQtySold,ProductName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "أقل المنتجات مبيعًا", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1361, 1292, 3629, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 4990, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "المخزون الحالي", 6691, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "صافي المباع", 7995, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "صافي المبيعات", 9299, 1292, 1474, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([CurrentQuantity])", 6691, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([NetQtySold])", 7995, 85, 1304, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([NetSalesAmount])", 9299, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 6691, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1361, 17, 3629, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4990, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 6691, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "NetQtySold", 7995, 17, 1304, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetSalesAmount", 9299, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptLeastSelling", s
    Exit Sub
EH:
    AbortReport "rptLeastSelling", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptPurchases()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptPurchases", "المشتريات", "PurchasesQuery", 15536, "", "DocDate,DocNumber", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المشتريات", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "النوع", 0, 1292, 850, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الرقم", 850, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "فاتورة المورد", 2268, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "التاريخ", 3629, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "المورد", 4876, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الدفع", 7371, 1292, 794, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "بدون ضريبة", 8165, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الضريبة", 9583, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "الإجمالي", 10830, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "المدفوع", 12304, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "على الحساب", 13778, 1292, 1758, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal7", "=Sum([NetAmount])", 8165, 85, 1418, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([VATAmount])", 9583, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal9", "=Sum([GrossAmount])", 10830, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([SettledAmount])", 12304, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal11", "=Sum([OnAccount])", 13778, 85, 1758, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 8165, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=IIf([DocType]=""RETURN"",""مرتجع"",""فاتورة"")", 0, 17, 850, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "DocNumber", 850, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "SupplierRef", 2268, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "=GDate([DocDate])", 3629, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "SupplierName", 4876, 17, 2495, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol6", "=IIf([PaymentType]=""CREDIT"",""آجل"",""نقدي"")", 7371, 17, 794, 284, 8, False, 0)
    Set c = RText(0, "txtCol7", "NetAmount", 8165, 17, 1418, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "VATAmount", 9583, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "GrossAmount", 10830, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "SettledAmount", 12304, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "OnAccount", 13778, 17, 1758, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptPurchases", s
    Exit Sub
EH:
    AbortReport "rptPurchases", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockBalance()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockBalance", "المخزون الحالي", "StockBalanceQuery", 15536, "", "CategoryName,ProductName", True, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 7768, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 7768, 28, 7768, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المخزون الحالي", 0, 397, 15536, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 15536, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 15536, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1304, 1292, 3175, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 4479, 1292, 1531, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "الوحدة", 6010, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الكمية", 6917, 1292, 1077, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الحد الأدنى", 7994, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "متوسط التكلفة", 9015, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "سعر البيع", 10262, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol9", "القيمة بالتكلفة", 11453, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol10", "القيمة بسعر البيع", 12927, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol11", "الحالة", 14401, 1292, 1135, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 9321, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 9321, 57, 6215, 255, 8, False, 1)
    Set c = RText(2, "txtTotal5", "=Sum([CurrentQuantity])", 6917, 85, 1077, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal9", "=Sum([StockCostValue])", 11453, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal10", "=Sum([StockSalesValue])", 12927, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 15536)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 6917, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1304, 17, 3175, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4479, 17, 1531, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "UnitName", 6010, 17, 907, 284, 8, False, 0)
    Set c = RText(0, "txtCol5", "CurrentQuantity", 6917, 17, 1077, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "MinimumQuantity", 7994, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "AverageCost", 9015, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "SellingPrice", 10262, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol9", "StockCostValue", 11453, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol10", "StockSalesValue", 12927, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol11", "=IIf([IsLowStock],""منخفض"","""") & IIf([QuantityMismatch]<>0,"" / فرق!"","""")", 14401, 17, 1135, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockBalance", s
    Exit Sub
EH:
    AbortReport "rptStockBalance", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptLowStock()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptLowStock", "المنتجات منخفضة المخزون", "LowStockQuery", 10773, "", "-ShortageQty,ProductName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المنتجات منخفضة المخزون", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1191, 1292, 2608, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 3799, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "المتوفر", 5103, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الحد الأدنى", 6067, 1292, 964, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "النقص", 7031, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "المورد", 7938, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "جوال المورد", 9639, 1292, 1134, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal6", "=Sum([ShortageQty])", 7031, 85, 907, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 7031, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1191, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1191, 17, 2608, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 3799, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 5103, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "MinimumQuantity", 6067, 17, 964, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "ShortageQty", 7031, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol7", "SupplierName", 7938, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol8", "SupplierMobile", 9639, 17, 1134, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد منتجات منخفضة المخزون.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptLowStock", s
    Exit Sub
EH:
    AbortReport "rptLowStock", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptProductMovement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptProductMovement", "حركة منتج", "ProductMovementQuery", 10773, "", "SortKey,MovementDate,TransactionID", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "حركة منتج", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "التاريخ", 0, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الحركة", 1644, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "المستند", 3118, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "وارد", 4536, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "صادر", 5557, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الرصيد", 6578, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "التكلفة", 7712, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "ملاحظات", 8846, 1292, 1927, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([QtyIn])", 4536, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal5", "=Sum([QtyOut])", 5557, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 4536, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([MovementDate],True)", 0, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "MovementType", 1644, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ReferenceNumber", 3118, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "QtyIn", 4536, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "QtyOut", 5557, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "NetQty", 6578, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    SetCtl c, "RunningSum", 2
    Set c = RText(0, "txtCol7", "UnitCost", 7712, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "Notes", 8846, 17, 1927, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptProductMovement", s
    Exit Sub
EH:
    AbortReport "rptProductMovement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCustomerStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCustomerStatement", "كشف حساب عميل", "CustomerStatementQuery", 10773, "", "SortKey,EntryDate", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "كشف حساب عميل", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "التاريخ", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "البيان", 1701, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "رقم المستند", 3742, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "مدين", 5443, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "دائن", 7144, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الرصيد", 8845, 1292, 1928, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([Debit])", 5443, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Credit])", 7144, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([EntryDate],True)", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryTypeName", 1701, 17, 2041, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "DocNumber", 3742, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Debit", 5443, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Credit", 7144, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "=[Debit]-[Credit]", 8845, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 2
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCustomerStatement", s
    Exit Sub
EH:
    AbortReport "rptCustomerStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSupplierStatement()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSupplierStatement", "كشف حساب مورد", "SupplierStatementQuery", 10773, "", "SortKey,EntryDate", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "كشف حساب مورد", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "التاريخ", 0, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "البيان", 1701, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "رقم المستند", 3742, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "مدين", 5443, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "دائن", 7144, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الرصيد", 8845, 1292, 1928, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([Debit])", 5443, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([Credit])", 7144, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "=GDate([EntryDate],True)", 0, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "EntryTypeName", 1701, 17, 2041, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "DocNumber", 3742, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Debit", 5443, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "Credit", 7144, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "=[Credit]-[Debit]", 8845, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    SetCtl c, "RunningSum", 2
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSupplierStatement", s
    Exit Sub
EH:
    AbortReport "rptSupplierStatement", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptExpenses()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptExpenses", "المصروفات", "ExpensesQuery", 10773, "", "ExpenseDate,ExpenseNumber", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المصروفات", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الرقم", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "التاريخ", 1247, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "النوع", 2438, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "الوصف", 3912, 1292, 2495, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الدفع", 6407, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "المبلغ", 7428, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "الضريبة", 8562, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "الإجمالي", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal6", "=Sum([Amount])", 7428, 85, 1134, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal7", "=Sum([Tax])", 8562, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal8", "=Sum([TotalAmount])", 9583, 85, 1190, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 7428, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ExpenseNumber", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "=GDate([ExpenseDate])", 1247, 17, 1191, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "ExpenseTypeName", 2438, 17, 1474, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "Description", 3912, 17, 2495, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol5", "MethodName", 6407, 17, 1021, 284, 8, False, 0)
    Set c = RText(0, "txtCol6", "Amount", 7428, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "Tax", 8562, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol8", "TotalAmount", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptExpenses", s
    Exit Sub
EH:
    AbortReport "rptExpenses", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptExpensesByType()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptExpensesByType", "المصروفات حسب النوع", "ExpensesByTypeQuery", 10773, "", "-AmountTotal", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المصروفات حسب النوع", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "نوع المصروف", 0, 1292, 3402, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "العدد", 3402, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "بدون ضريبة", 4763, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "الضريبة", 6691, 1292, 1814, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "الإجمالي", 8505, 1292, 2268, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([ExpenseCount])", 3402, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([AmountExVAT])", 4763, 85, 1928, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal4", "=Sum([InputVAT])", 6691, 85, 1814, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([AmountTotal])", 8505, 85, 2268, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 3402, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ExpenseTypeName", 0, 17, 3402, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ExpenseCount", 3402, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "AmountExVAT", 4763, 17, 1928, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "InputVAT", 6691, 17, 1814, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "AmountTotal", 8505, 17, 2268, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptExpensesByType", s
    Exit Sub
EH:
    AbortReport "rptExpensesByType", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSlowMoving()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSlowMoving", "المنتجات غير المتحركة", "SlowMovingProductsQuery", 10773, "", "-DaysWithoutSale", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المنتجات غير المتحركة", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الكود", 0, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتج", 1247, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "التصنيف", 4082, 1292, 1418, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "الكمية", 5500, 1292, 1021, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "متوسط التكلفة", 6521, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "القيمة", 7655, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "آخر بيع", 8902, 1292, 1134, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol8", "أيام", 10036, 1292, 737, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([CurrentQuantity])", 5500, 85, 1021, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal6", "=Sum([StockCostValue])", 7655, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5500, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "ProductCode", 0, 17, 1247, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductName", 1247, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "CategoryName", 4082, 17, 1418, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "CurrentQuantity", 5500, 17, 1021, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol5", "AverageCost", 6521, 17, 1134, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "StockCostValue", 7655, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastSaleDate])", 8902, 17, 1134, 284, 8, False, 0)
    Set c = RText(0, "txtCol8", "DaysWithoutSale", 10036, 17, 737, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد منتجات راكدة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSlowMoving", s
    Exit Sub
EH:
    AbortReport "rptSlowMoving", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptStockByCategory()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptStockByCategory", "المخزون حسب التصنيف", "StockByCategoryQuery", 10773, "", "CategoryName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "المخزون حسب التصنيف", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "التصنيف", 0, 1292, 3402, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المنتجات", 3402, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "الكمية", 4763, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "القيمة بالتكلفة", 6464, 1292, 2041, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "القيمة بسعر البيع", 8505, 1292, 2268, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal2", "=Sum([ProductCount])", 3402, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(2, "txtTotal3", "=Sum([TotalQuantity])", 4763, 85, 1701, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(2, "txtTotal4", "=Sum([StockCostValue])", 6464, 85, 2041, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([StockSalesValue])", 8505, 85, 2268, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 3402, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CategoryName", 0, 17, 3402, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "ProductCount", 3402, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol3", "TotalQuantity", 4763, 17, 1701, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol4", "StockCostValue", 6464, 17, 2041, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "StockSalesValue", 8505, 17, 2268, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptStockByCategory", s
    Exit Sub
EH:
    AbortReport "rptStockByCategory", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptCustomerBalances()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptCustomerBalances", "أرصدة العملاء", "CustomerBalanceQuery", 10773, "", "CustomerName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "أرصدة العملاء", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "العميل", 0, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "الجوال", 2835, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "حد الائتمان", 4196, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "مدين", 5443, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "دائن", 6804, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الرصيد", 8165, 1292, 1474, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "آخر حركة", 9639, 1292, 1134, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([DebitTotal])", 5443, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([CreditTotal])", 6804, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Balance])", 8165, 85, 1474, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5443, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "CustomerName", 0, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol2", "Mobile", 2835, 17, 1361, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "CreditLimit", 4196, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol4", "DebitTotal", 5443, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "CreditTotal", 6804, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Balance", 8165, 17, 1474, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastEntryDate])", 9639, 17, 1134, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptCustomerBalances", s
    Exit Sub
EH:
    AbortReport "rptCustomerBalances", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptSupplierBalances()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSupplierBalances", "أرصدة الموردين", "SupplierBalanceQuery", 10773, "", "SupplierName", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "أرصدة الموردين", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "المورد", 0, 1292, 2835, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المسؤول", 2835, 1292, 1644, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "الجوال", 4479, 1292, 1304, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "مدين", 5783, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "دائن", 7030, 1292, 1247, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الرصيد", 8277, 1292, 1361, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol7", "آخر حركة", 9638, 1292, 1135, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RText(2, "txtTotal4", "=Sum([DebitTotal])", 5783, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal5", "=Sum([CreditTotal])", 7030, 85, 1247, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(2, "txtTotal6", "=Sum([Balance])", 8277, 85, 1361, 284, 8, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 5783, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "SupplierName", 0, 17, 2835, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol2", "ContactPerson", 2835, 17, 1644, 284, 8, False, 0)
    Set c = RText(0, "txtCol3", "Mobile", 4479, 17, 1304, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "DebitTotal", 5783, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol5", "CreditTotal", 7030, 17, 1247, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol6", "Balance", 8277, 17, 1361, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(0, "txtCol7", "=GDate([LastEntryDate])", 9638, 17, 1135, 284, 8, False, 0)
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSupplierBalances", s
    Exit Sub
EH:
    AbortReport "rptSupplierBalances", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptIntegrityCheck()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptIntegrityCheck", "فحص سلامة البيانات", "IntegrityCheckQuery", 10773, "", "IssueCode,RecordID", False, True
    SetSection 3, 1673
    SetSection 4, 340
    SetSection 2, 454
    SetSection 0, 318
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "فحص سلامة البيانات", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RBox(3, "boxColumns", 0, 1247, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(3, "lblCol1", "الرمز", 0, 1292, 1928, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol2", "المشكلة", 1928, 1292, 3856, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol3", "الجدول", 5784, 1292, 1701, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol4", "السجل", 7485, 1292, 907, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol5", "المتوقع", 8392, 1292, 1191, 284, 8, True, 2)
    Set c = RLabel(3, "lblCol6", "الفعلي", 9583, 1292, 1190, 284, 8, True, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLine(2, "lnTotals", 28, 10773)
    Set c = RText(2, "txtCount", "=""الإجمالي ("" & Count(*) & "" سجل)""", 0, 85, 10773, 284, 8, True, 0)
    Set c = RText(0, "txtCol1", "IssueCode", 0, 17, 1928, 284, 8, False, 0)
    Set c = RText(0, "txtCol2", "IssueText", 1928, 17, 3856, 284, 8, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtCol3", "SourceTable", 5784, 17, 1701, 284, 8, False, 0)
    Set c = RText(0, "txtCol4", "RecordID", 7485, 17, 907, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0"
    Set c = RText(0, "txtCol5", "ExpectedValue", 8392, 17, 1191, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    Set c = RText(0, "txtCol6", "ActualValue", 9583, 17, 1190, 284, 8, False, 2)
    SetCtl c, "Format", "#,##0.###"
    m_rpt.OnNoData = EP
    SetSecProp 0, "AlternateBackColor", 15921906
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""فحص سلامة البيانات: لا توجد أي مشكلات.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptIntegrityCheck", s
    Exit Sub
EH:
    AbortReport "rptIntegrityCheck", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptProfit()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptProfit", "الأرباح", "ProfitQuery", 10773, "", "", False, True
    SetSection 3, 1304
    SetSection 4, 340
    SetSection 2, 624
    SetSection 0, 3088
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "الأرباح", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLabel(2, "lblNote", "صافي الربح = مجمل الربح ± فروقات المخزون - المصروفات. الرصيد الافتتاحي لا يدخل في الربح.", 0, 113, 10773, 454, 9, False, 2)
    Set c = RLabel(0, "lblRow1", "صافي المبيعات (بدون الضريبة)", 1134, 57, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow1", "NetSales", 6577, 57, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow2", "تكلفة البضاعة المباعة", 1134, 482, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow2", "CostOfSales", 6577, 482, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow3", "مجمل الربح", 1134, 907, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow3", "GrossProfit", 6577, 907, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow4", "نسبة مجمل الربح", 1134, 1332, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow4", "=IIf([NetSales]=0,0,[GrossProfit]/[NetSales])", 6577, 1332, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "0.0%"
    Set c = RLabel(0, "lblRow5", "فروقات المخزون (جرد، إضافة، خصم)", 1134, 1757, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow5", "InventoryAdjustments", 6577, 1757, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow6", "المصروفات (بدون الضريبة)", 1134, 2182, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow6", "TotalExpenses", 6577, 2182, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow7", "صافي الربح", 1134, 2607, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow7", "NetProfit", 6577, 2607, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptProfit", s
    Exit Sub
EH:
    AbortReport "rptProfit", Err.Number, Err.Description
End Sub

Private Sub BuildReport_rptVatSummary()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptVatSummary", "ملخص ضريبة القيمة المضافة", "VatSummaryQuery", 10773, "", "", False, True
    SetSection 3, 1304
    SetSection 4, 340
    SetSection 2, 624
    SetSection 0, 3088
    HideSection 1
    Set c = RText(3, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 28, 5386, 340, 11, True, 0)
    Set c = RText(3, "txtStoreVat", "=IIf(Len(Nz(SettingValue(""VATNumber""),""""))>0,""الرقم الضريبي: "" & SettingValue(""VATNumber""),"""")", 5386, 28, 5387, 340, 9, False, 1)
    Set c = RLabel(3, "lblTitle", "ملخص ضريبة القيمة المضافة", 0, 397, 10773, 482, 16, True, 2)
    Set c = RText(3, "txtCriteria", "=ReportCriteria()", 0, 907, 10773, 284, 10, False, 2)
    Set c = RText(4, "txtPrinted", "=ReportPrintedAt()", 0, 57, 6463, 255, 8, False, 0)
    Set c = RText(4, "txtPage", "=""صفحة "" & [Page] & "" من "" & [Pages]", 6463, 57, 4310, 255, 8, False, 1)
    Set c = RLabel(2, "lblNote", "الموجب مستحق للهيئة، والسالب رصيد مسترد. راجع الأرقام مع محاسبك قبل تقديم الإقرار.", 0, 113, 10773, 454, 9, False, 2)
    Set c = RLabel(0, "lblRow1", "المبيعات الخاضعة للضريبة (بعد المرتجعات)", 1134, 57, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow1", "TaxableSales", 6577, 57, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow2", "ضريبة المخرجات", 1134, 482, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow2", "OutputVAT", 6577, 482, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow3", "المشتريات الخاضعة للضريبة (بعد المرتجعات)", 1134, 907, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow3", "TaxablePurchases", 6577, 907, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow4", "ضريبة مدخلات المشتريات", 1134, 1332, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow4", "PurchaseVAT", 6577, 1332, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow5", "ضريبة مدخلات المصروفات", 1134, 1757, 5386, 340, 11, False, 0)
    Set c = RText(0, "txtRow5", "ExpenseVAT", 6577, 1757, 3062, 340, 11, False, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow6", "إجمالي ضريبة المدخلات", 1134, 2182, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow6", "InputVAT", 6577, 2182, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(0, "lblRow7", "صافي الضريبة المستحقة", 1134, 2607, 5386, 340, 11, True, 0)
    Set c = RText(0, "txtRow7", "NetVATDue", 6577, 2607, 3062, 340, 12, True, 2)
    SetCtl c, "Format", "#,##0.00"
    m_rpt.OnNoData = EP
    s = ""
    s = s & "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf
    s = s & "    ReportNoData Cancel, ""لا توجد بيانات لعرضها في هذا التقرير للاختيارات المحددة.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptVatSummary", s
    Exit Sub
EH:
    AbortReport "rptVatSummary", Err.Number, Err.Description
End Sub
