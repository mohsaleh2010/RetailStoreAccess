Attribute VB_Name = "modBuildReports"
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
    DoCmd.Echo False, "Ã«—Ì »‰«¡ «· ﬁ«—Ì—..."
    BuildReport_rptSalesReceipt
    BuildReport_rptSalesInvoiceA4
    DoCmd.Echo True
    Application.RefreshDatabaseWindow
    If m_failed = 0 Then
        MsgBox " „ »‰«¡ «· ﬁ«—Ì— »‰Ã«Õ (" & m_built & ").", vbInformation + MSG_RTL, "BuildReports"
        BuildReports = True
    Else
        MsgBox "›‘· »‰«¡ " & m_failed & "  ﬁ—Ì—:" & vbCrLf & m_report, vbExclamation + MSG_RTL, "BuildReports"
    End If
    Exit Function
EH:
    DoCmd.Echo True
    MsgBox "Œÿ√ " & Err.Number & ": " & Err.Description, vbCritical + MSG_RTL, "BuildReports"
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
Private Sub BuildReport_rptSalesReceipt()
    Dim c As Access.Control, s As String
    On Error GoTo EH
    StartReport "rptSalesReceipt", "›« Ê—… (Õ—«—Ì 80 „„)", "qrySalesDocPrint", 4196, 3827, 539, 4734
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 4196, 340, 12, True, 2)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 397, 4196, 255, 9, False, 2)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 652, 4196, 227, 9, True, 2)
    Set c = RText(5, "txtStoreCR", "=""”. : "" & Nz(SettingValue(""CRNumber""),""-"") & ""   Â« ›: "" & Nz(SettingValue(""Phone""),""-"")", 0, 879, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""District""),"""") & "" "" & Nz(SettingValue(""StreetName""),""""))", 0, 1106, 4196, 227, 8, False, 2)
    Set c = RLine(5, "lnHeader1", 1361, 4196)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""≈‘⁄«— œ«∆‰"",IIf([InvoiceSubType]=""STANDARD"",""›« Ê—… ÷—Ì»Ì…"",""›« Ê—… ÷—Ì»Ì… „»”ÿ…""))", 0, 1418, 4196, 312, 12, True, 2)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 0, 1730, 4196, 227, 8, False, 2)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„: "" & [DocNumber]", 0, 1985, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 0, 2212, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""⁄‰ «·›« Ê—…: "" & [OriginalNumber],"""")", 0, 2439, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCashier", "=""«·ﬂ«‘Ì—: "" & [EmployeeName]", 0, 2666, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomer", "=IIf([CustomerID]=Nz(SettingValue(""DefaultCustomerID""),1),"""",""«·⁄„Ì·: "" & [CustomerName])", 0, 2893, 4196, 227, 8, False, 0)
    Set c = RText(5, "txtCustomerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""«·—ﬁ„ «·÷—Ì»Ì ··⁄„Ì·: "" & [CustomerVAT],"""")", 0, 3120, 4196, 227, 8, False, 0)
    Set c = RLine(5, "lnHeader2", 3375, 4196)
    Set c = RLabel(5, "lblColItem", "«·’‰›", 0, 3404, 2211, 255, 8, True, 0)
    Set c = RLabel(5, "lblColQty", "«·ﬂ„Ì… ◊ «·”⁄—", 2211, 3404, 1134, 255, 8, True, 0)
    Set c = RLabel(5, "lblColTotal", "«·≈Ã„«·Ì", 3345, 3404, 850, 255, 8, True, 1)
    Set c = RLine(5, "lnHeader3", 3688, 4196)
    Set c = RText(0, "txtProduct", "ProductName", 0, 0, 4196, 255, 9, False, 0)
    SetCtl c, "CanGrow", True
    Set c = RText(0, "txtQtyPrice", "=[Quantity] & "" ◊ "" & Format([LineTotal]/[Quantity],""#,##0.00"")", 113, 266, 2835, 238, 8, False, 0)
    Set c = RText(0, "txtLineTotal", "LineTotal", 2948, 266, 1247, 238, 9, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(6, "lnTotals1", 45, 4196)
    Set c = RLabel(6, "lblCapTaxableAmount", "«·≈Ã„«·Ì »œÊ‰ «·÷—Ì»…", 0, 113, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 2665, 113, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "«·Œ’„", 0, 397, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 2665, 397, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapDocTax", "=""÷—Ì»… «·ﬁÌ„… «·„÷«›… "" & Format(Nz(SettingValue(""VATRate""),0.15),""0%"")", 0, 681, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 2665, 681, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 0, 965, 2608, 255, 11, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 2665, 965, 1531, 255, 11, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLine(6, "lnTotals2", 1277, 4196)
    Set c = RText(6, "txtCapPaid", "=IIf([DocKind]=""RETURN"",""«·„»·€ «·„—œÊœ"",""«·„œ›Ê⁄"")", 0, 1334, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValPaid", "PaidAmount", 2665, 1334, 1531, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RText(6, "txtCapChange", "=IIf([ChangeDue]>0,""«·»«ﬁÌ ··⁄„Ì·"","""")", 0, 1589, 2608, 255, 9, False, 0)
    Set c = RText(6, "txtValChange", "=IIf([ChangeDue]>0,Format([ChangeDue],""#,##0.00""),"""")", 2665, 1589, 1531, 255, 9, False, 1)
    Set c = RText(6, "txtCapRemaining", "=IIf([RemainingAmount]<>0,IIf([DocKind]=""RETURN"",""Œ’„ „‰ —’Ìœ «·⁄„Ì·"",""«·„ »ﬁÌ ⁄·Ï «·Õ”«»""),"""")", 0, 1844, 2608, 255, 9, False, 0)
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
    StartReport "rptSalesInvoiceA4", "›« Ê—… ÷—Ì»Ì… (A4)", "qrySalesDocPrint", 10773, 3856, 340, 4649
    Set c = RText(5, "txtStoreName", "=Nz(SettingValue(""StoreName""),"""")", 0, 57, 5273, 397, 14, True, 0)
    Set c = RText(5, "txtStoreNameEn", "=Nz(SettingValue(""StoreNameEn""),"""")", 0, 454, 5273, 284, 10, False, 0)
    Set c = RText(5, "txtStoreVat", "=""«·—ﬁ„ «·÷—Ì»Ì: "" & Nz(SettingValue(""VATNumber""),"""")", 0, 738, 5273, 284, 10, True, 0)
    Set c = RText(5, "txtStoreCR", "=""«·”Ã· «· Ã«—Ì: "" & Nz(SettingValue(""CRNumber""),"""")", 0, 1022, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStoreAddress", "=Trim(Nz(SettingValue(""BuildingNo""),"""") & "" "" & Nz(SettingValue(""StreetName""),"""") & "" - "" & Nz(SettingValue(""District""),"""") & "" - "" & Nz(SettingValue(""City""),"""") & "" "" & Nz(SettingValue(""PostalCode""),""""))", 0, 1306, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtStorePhone", "=""Â« ›: "" & Nz(SettingValue(""Phone""),"""")", 0, 1590, 5273, 284, 9, False, 0)
    Set c = RText(5, "txtTitle", "=IIf([DocKind]=""RETURN"",""≈‘⁄«— œ«∆‰"",IIf([InvoiceSubType]=""STANDARD"",""›« Ê—… ÷—Ì»Ì…"",""›« Ê—… ÷—Ì»Ì… „»”ÿ…""))", 5500, 57, 5273, 454, 16, True, 1)
    Set c = RText(5, "txtTitleEn", "=IIf([DocKind]=""RETURN"",""Credit Note"",IIf([InvoiceSubType]=""STANDARD"",""Tax Invoice"",""Simplified Tax Invoice""))", 5500, 539, 5273, 284, 10, False, 1)
    Set c = RText(5, "txtDocNumber", "=""—ﬁ„ «·›« Ê—…: "" & [DocNumber]", 5500, 850, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtDocDate", "=""«· «—ÌŒ: "" & Format([DocDate],""yyyy/mm/dd hh:nn"")", 5500, 1105, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtOriginal", "=IIf(Len(Nz([OriginalNumber],""""))>0,""⁄‰ «·›« Ê—…: "" & [OriginalNumber],"""")", 5500, 1360, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtPaymentType", "=""ÿ—Ìﬁ… «·»Ì⁄: "" & IIf([PaymentType]=""CREDIT"",""¬Ã·"",""‰ﬁœÌ"")", 5500, 1615, 5273, 255, 9, False, 1)
    Set c = RText(5, "txtCashier", "=""«·„ÊŸ›: "" & [EmployeeName]", 5500, 1870, 5273, 255, 9, False, 1)
    Set c = RBox(5, "boxBuyer", 0, 2183, 10773, 1106)
    SetCtl c, "BorderStyle", 1
    Set c = RLabel(5, "lblBuyer", "»Ì«‰«  «·„‘ —Ì", 113, 2211, 2268, 255, 9, True, 0)
    Set c = RText(5, "txtBuyerName", "=""«·«”„: "" & [CustomerName]", 113, 2467, 5103, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerVat", "=IIf(Len(Nz([CustomerVAT],""""))>0,""«·—ﬁ„ «·÷—Ì»Ì: "" & [CustomerVAT],"""")", 5443, 2467, 5216, 255, 9, False, 0)
    Set c = RText(5, "txtBuyerAddress", "=Trim(Nz([CustomerBuilding],"""") & "" "" & Nz([CustomerStreet],"""") & "" "" & Nz([CustomerDistrict],"""") & "" "" & Nz([CustomerCity],"""") & "" "" & Nz([CustomerPostal],""""))", 113, 2750, 10546, 255, 9, False, 0)
    Set c = RBox(5, "boxColumns", 0, 3430, 10773, 369)
    SetCtl c, "BackStyle", 1
    SetCtl c, "BackColor", CLR_SECONDARY
    Set c = RLabel(5, "lblCol1", "#", 0, 3487, 454, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol2", "«·’‰›", 454, 3487, 3515, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol3", "«·ﬂ„Ì…", 3969, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol4", "”⁄— «·ÊÕœ…", 4876, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol5", "«·Œ’„", 6010, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol6", "«·’«›Ì", 6917, 3487, 1134, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol7", "‰”»… «·÷—Ì»…", 8051, 3487, 680, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol8", "«·÷—Ì»…", 8731, 3487, 907, 255, 8, True, 2)
    Set c = RLabel(5, "lblCol9", "«·≈Ã„«·Ì", 9638, 3487, 1135, 255, 8, True, 2)
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
    Set c = RLine(6, "lnTotals1", 45, 10773)
    Set c = RLabel(6, "lblCapDocSubTotal", "«·≈Ã„«·Ì ﬁ»· «·Œ’„", 6464, 113, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocSubTotal", "DocSubTotal", 8959, 113, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocDiscount", "«·Œ’„", 6464, 397, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocDiscount", "DocDiscount", 8959, 397, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTaxableAmount", "«·≈Ã„«·Ì «·Œ«÷⁄ ··÷—Ì»…", 6464, 681, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtTaxableAmount", "TaxableAmount", 8959, 681, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapDocTax", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", 6464, 965, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtDocTax", "DocTax", 8959, 965, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapTotalAmount", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", 6464, 1249, 2495, 255, 10, True, 0)
    Set c = RText(6, "txtTotalAmount", "TotalAmount", 8959, 1249, 1814, 255, 10, True, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapPaidAmount", "«·„œ›Ê⁄ / «·„—œÊœ", 6464, 1533, 2495, 255, 9, False, 0)
    Set c = RText(6, "txtPaidAmount", "PaidAmount", 8959, 1533, 1814, 255, 9, False, 1)
    SetCtl c, "Format", "#,##0.00"
    Set c = RLabel(6, "lblCapRemainingAmount", "«·„ »ﬁÌ", 6464, 1817, 2495, 255, 9, False, 0)
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
    s = s & "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _" & vbCrLf
    s = s & "                   Me!boxQR.Top, Me!boxQR.Width" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FinishReport "rptSalesInvoiceA4", s
    Exit Sub
EH:
    AbortReport "rptSalesInvoiceA4", Err.Number, Err.Description
End Sub
