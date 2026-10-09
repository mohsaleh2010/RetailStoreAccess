Attribute VB_Name = "modZatcaApi"
'==============================================================================
' modZatcaApi  -  Retail Store Management System (ZATCA Fatoora API)
'
' docs/47-ZATCA-Onboarding-Sending.md. The Saudi platform module of modEInvoice:
'   Onboarding of the device (frmZatcaSetup), in three steps:
'     1 ZatcaRequestCompliance  key + certificate request (CSR, OpenSSL) + OTP -> compliance CSID
'     2 ZatcaComplianceChecks   six signed samples (invoice, credit and debit note; standard, simplified)
'     3 ZatcaRequestProduction  production CSID: the certificate that signs the real documents
'   Sending (called by modEInvoice through Application.Run):
'     ZatcaReadyProblem()                 "" when the device is onboarded in the current environment
'     ZatcaSendDocument(DocKind, DocID)   signs (modZatcaXml) then reports a simplified document or clears a
'                                         standard one; returns "RESULT|message"
' tools/zatca_api_reference.py is the Python mirror of the CSR configuration and of the reading of replies.
'==============================================================================
Option Compare Database
Option Explicit

Private Const SOLUTION_NAME As String = "RetailStoreAccess"
Private Const INVOICE_TYPES As String = "1100"          ' standard (B2B) and simplified (B2C)

'==============================================================================
' Environment, CSR configuration, replies (no data access)
'==============================================================================
Public Function ZatcaBaseUrl(ByVal Env As String) As String
    Select Case Env
        Case "PRODUCTION": ZatcaBaseUrl = "https://gw-fatoora.zatca.gov.sa/e-invoicing/core"
        Case "SIMULATION": ZatcaBaseUrl = "https://gw-fatoora.zatca.gov.sa/e-invoicing/simulation"
        Case Else: ZatcaBaseUrl = "https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal"
    End Select
End Function

Public Function ZatcaTemplateName(ByVal Env As String) As String
    Select Case Env
        Case "PRODUCTION": ZatcaTemplateName = "ZATCA-Code-Signing"
        Case "SIMULATION": ZatcaTemplateName = "PREZATCA-Code-Signing"
        Case Else: ZatcaTemplateName = "TSTZATCA-Code-Signing"
    End Select
End Function

Public Function ZatcaConfigValue(ByVal s As String) As String
    ' One line of the OpenSSL configuration: no comment (#) nor variable ($), single spaces.
    s = Replace(Replace(Replace(Replace(Replace(s, "#", " "), "$", " "), vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ZatcaConfigValue = Trim$(s)
End Function

Public Function ZatcaCsrConfig(ByVal Env As String, ByVal Serial As String, ByVal Vat As String, _
                               ByVal Location As String, ByVal Industry As String, ByVal CommonName As String, _
                               ByVal Branch As String, ByVal Organization As String) As String
    ' The configuration of "openssl req" for the CSR ZATCA expects: subject C, OU, O, CN; the template name
    ' extension of the environment; the device data in subjectAltName (dirName).
    ZatcaCsrConfig = "[req]" & vbLf & "prompt = no" & vbLf & "utf8 = yes" & vbLf & "string_mask = utf8only" & vbLf & _
        "distinguished_name = dn" & vbLf & "req_extensions = v3_req" & vbLf & vbLf & _
        "[dn]" & vbLf & "C = SA" & vbLf & "OU = " & ZatcaConfigValue(Branch) & vbLf & "O = " & ZatcaConfigValue(Organization) & _
        vbLf & "CN = " & ZatcaConfigValue(CommonName) & vbLf & vbLf & _
        "[v3_req]" & vbLf & "1.3.6.1.4.1.311.20.2 = ASN1:UTF8String:" & ZatcaTemplateName(Env) & vbLf & _
        "subjectAltName = dirName:dir_sect" & vbLf & vbLf & _
        "[dir_sect]" & vbLf & "SN = 1-" & SOLUTION_NAME & "|2-1.0|3-" & ZatcaConfigValue(Serial) & vbLf & _
        "UID = " & ZatcaConfigValue(Vat) & vbLf & "title = " & INVOICE_TYPES & vbLf & _
        "registeredAddress = " & ZatcaConfigValue(Location) & vbLf & "businessCategory = " & ZatcaConfigValue(Industry) & vbLf
End Function

Public Function ZatcaMessages(ByVal Body As String, ByVal Kind As String) As String
    ' "code: message; ..." (at most 3) of validationResults.errorMessages / warningMessages.
    Dim i As Long, path As String, out As String
    For i = 0 To 2
        path = "validationResults." & Kind & "[" & i & "]"
        If IsNull(JsonGet(Body, path)) Then Exit For
        If Len(out) > 0 Then out = out & "; "
        out = out & Nz(JsonGet(Body, path & ".code"), "") & ": " & Nz(JsonGet(Body, path & ".message"), "")
    Next
    ZatcaMessages = out
End Function

Public Function ZatcaReplyResult(ByVal HttpStatus As Long, ByVal Body As String, ByVal IsStandard As Boolean, _
                                 ByRef DocStatus As String, ByRef Message As String) As String
    ' The result of a reporting / clearance reply: OK, WARNING, REJECTED, ERROR or NETWORK; the new status of
    ' the document ("" = it stays PENDING) and the message.
    DocStatus = ""
    Message = ""
    If HttpStatus = 0 Then
        ZatcaReplyResult = "NETWORK"
        Message = "تعذّر الاتصال بالمنظومة."
    ElseIf HttpStatus = 200 Then
        ZatcaReplyResult = "OK"
        DocStatus = IIf(IsStandard, "CLEARED", "REPORTED")
    ElseIf HttpStatus = 202 Then
        ZatcaReplyResult = "WARNING"
        DocStatus = "WARNING"
        Message = ZatcaMessages(Body, "warningMessages")
    ElseIf HttpStatus = 400 Then
        ZatcaReplyResult = "REJECTED"
        DocStatus = "REJECTED"
        Message = ZatcaMessages(Body, "errorMessages")
        If Len(Message) = 0 Then Message = Left$(Body, 200)
    ElseIf HttpStatus = 401 Or HttpStatus = 403 Then
        ZatcaReplyResult = "ERROR"
        Message = "رفضت الهيئة بيانات الدخول (شهادة الجهاز أو الكلمة السرية)."
    ElseIf HttpStatus = 429 Or HttpStatus >= 500 Then
        ZatcaReplyResult = "NETWORK"
        Message = "المنظومة غير متاحة الآن (رمز " & HttpStatus & ")."
    Else
        ZatcaReplyResult = "ERROR"
        Message = "رد غير متوقع من الهيئة (رمز " & HttpStatus & "): " & Left$(Body, 200)
    End If
End Function

'==============================================================================
' Requests
'==============================================================================
Private Function Txt(ByVal v As Variant) As String
    Txt = Trim$(Nz(v, ""))
End Function

Private Function CurrentEnv() As String
    CurrentEnv = Nz(SettingValue("EInvoiceEnvironment"), "TEST")
End Function

Public Function ZatcaPost(ByVal Path As String, ByVal Body As String, ByVal Token As String, ByVal Secret As String, _
                          ByVal ExtraHeaders As String, ByRef Status As Long, ByRef Reply As String, _
                          ByRef DurationMs As Long) As String
    ' POST to the API of the current environment: "" when ZATCA answered, else the network problem.
    Dim headers As String, started As Single
    headers = "Accept-Version: V2" & vbLf & "Accept-Language: ar" & vbLf & "Content-Type: application/json" & vbLf & _
              "Accept: application/json"
    If Len(Token) > 0 Then headers = headers & vbLf & "Authorization: Basic " & Base64Text(Token & ":" & Secret)
    If Len(ExtraHeaders) > 0 Then headers = headers & vbLf & ExtraHeaders
    started = Timer
    ZatcaPost = HttpSend("POST", ZatcaBaseUrl(CurrentEnv()) & Path, headers, Body, Status, Reply, 30)
    DurationMs = CLng((Timer - started) * 1000)
End Function

Public Function MaskSecret(ByVal Reply As String) As String
    ' A reply that carries the secret of a certificate, as it is kept in the log: the secret is hidden.
    Dim secretText As Variant
    secretText = JsonGet(Reply, "secret")
    MaskSecret = Reply
    If Not IsNull(secretText) Then
        If Len(secretText) > 0 Then MaskSecret = Replace(Reply, CStr(secretText), "***")
    End If
End Function

Private Function TokenCertificate(ByVal Token As String) As String
    ' The certificate in a binarySecurityToken: the token is the base64 of the certificate's base64 text.
    Dim b() As Byte, n As Long
    n = Base64Decode(Token, b)
    If n > 0 Then TokenCertificate = CertificateBody(Utf8Decode(b, 0, n))
End Function

'==============================================================================
' Onboarding (frmZatcaSetup)
'==============================================================================
Public Function ZatcaRequestCompliance(ByVal Otp As String, ByVal Branch As String, ByVal Industry As String) As String
    ' Step 1: a new key and certificate request (OpenSSL), sent with the OTP of the Fatoora portal (any value
    ' in the test environment); ZATCA answers with the compliance CSID. "" or the problem.
    Dim folder As String, keyFile As String, csrFile As String, cnf As String, errText As String, f As Integer
    Dim ln As String, csrPem As String, serial As String, s As DAO.Recordset, status As Long, reply As String
    Dim ms As Long, msg As String, location As String
    If AppCountry() <> "SA" Then
        ZatcaRequestCompliance = "منصة فاتورة للمنشآت في السعودية فقط (دولة التشغيل في الإعدادات)."
        Exit Function
    End If
    msg = ZatcaSellerProblem()
    If Len(msg) > 0 Then
        ZatcaRequestCompliance = "بيانات المحل ناقصة:" & vbCrLf & msg
        Exit Function
    End If
    If Len(Trim$(Otp)) = 0 Then
        ZatcaRequestCompliance = "اكتب رمز التحقق (OTP) من بوابة فاتورة."
        Exit Function
    End If
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    serial = Txt(s!ZatcaDeviceSerial)
    If Len(serial) = 0 Then serial = NewUUID()
    location = Txt(s!BuildingNo) & " " & Txt(s!StreetName) & ", " & Txt(s!District) & ", " & Txt(s!City)
    folder = CurrentProject.Path & "\ZATCA"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    ' a new file each time: the key of the registered device stays until ZATCA accepts the new request
    keyFile = folder & "\zatca-key-" & LCase$(CurrentEnv()) & "-" & Format$(Now, "yyyymmddhhnnss") & ".pem"
    csrFile = ZatcaWorkFolder() & "\zatca.csr"
    cnf = ZatcaWorkFolder() & "\zatca-csr.cnf"
    WriteUtf8File cnf, ZatcaCsrConfig(CurrentEnv(), serial, Txt(s!VATNumber), location, Industry, _
                                      "EGS-" & Left$(serial, 8), Branch, Txt(s!StoreName))
    s.Close
    If RunOpenSsl("ecparam -name secp256k1 -genkey -noout -out """ & keyFile & """", errText) <> 0 Then
        ZatcaRequestCompliance = "تعذّر إنشاء المفتاح ببرنامج OpenSSL: " & errText
        Exit Function
    End If
    If Len(Dir$(csrFile)) > 0 Then Kill csrFile
    If RunOpenSsl("req -new -sha256 -key """ & keyFile & """ -config """ & cnf & """ -out """ & csrFile & """", _
                  errText) <> 0 Or Len(Dir$(csrFile)) = 0 Then
        ZatcaRequestCompliance = "تعذّر إنشاء طلب الشهادة: " & errText
        Exit Function
    End If
    f = FreeFile
    Open csrFile For Input As #f
    Do Until EOF(f)
        Line Input #f, ln
        csrPem = csrPem & ln & vbLf
    Loop
    Close #f
    msg = ZatcaPost("/compliance", "{""csr"": """ & Base64Text(csrPem) & """}", "", "", "OTP: " & Trim$(Otp), _
                    status, reply, ms)
    LogEInvoice "", 0, "COMPLIANCE_CSID", "/compliance", status, IIf(status = 200, "OK", "ERROR"), _
                Left$(msg & MaskSecret(reply), 255), "", MaskSecret(reply), ms
    If status <> 200 Or IsNull(JsonGet(reply, "binarySecurityToken")) Then
        ZatcaRequestCompliance = "لم تُصدر الهيئة شهادة الامتثال: " & msg & " " & Left$(reply, 300)
        Exit Function
    End If
    SaveSettings Array("ZatcaDeviceSerial", serial, "ZatcaBranchName", Branch, "ZatcaIndustry", Industry, _
                 "ZatcaKeyFile", keyFile, "ZatcaCsr", csrPem, "ZatcaComplianceToken", JsonGet(reply, "binarySecurityToken"), _
                 "ZatcaComplianceSecret", JsonGet(reply, "secret"), "ZatcaRequestId", Nz(JsonGet(reply, "requestID"), ""), _
                 "ZatcaProductionToken", Null, "ZatcaProductionSecret", Null, "ZatcaOnboardEnv", CurrentEnv(), _
                 "ZatcaOnboardStage", "COMPLIANCE")
End Function

Private Sub SaveSettings(ByVal Pairs As Variant)
    ' Array(field, value, field, value ...) of the settings, with the audit trail.
    Dim before As Collection, rs As DAO.Recordset, i As Long
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenDynaset)
    rs.Edit
    For i = LBound(Pairs) To UBound(Pairs) - 1 Step 2
        If IsNull(Pairs(i + 1)) Then
            rs(Pairs(i)) = Null
        ElseIf Len(CStr(Pairs(i + 1))) = 0 Then
            rs(Pairs(i)) = Null
        Else
            rs(Pairs(i)) = Pairs(i + 1)
        End If
    Next
    rs.Update
    rs.Close
    AuditEdited "EDIT", "Settings", "SettingID", 1, before
End Sub

Public Function ZatcaSampleXml(ByVal TypeCode As String, ByVal SubType As String, ByVal Icv As Long, _
                               ByVal Pih As String, ByVal Uuid As String, ByRef Totals As Variant) As String
    ' A compliance sample: one line of 100.00 + 15% VAT, the seller of the settings, a fixed buyer for the
    ' standard documents, the original invoice and the reason for the notes.
    Dim s As DAO.Recordset, seller As String, buyer As String, standard As Boolean, note As Boolean, d As String, t As String
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    standard = (SubType = "0100000")
    note = (TypeCode <> "388")
    d = Format$(Date, "yyyy-mm-dd")
    t = Format$(Now, "hh:nn:ss")
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    seller = ZxParty("Supplier", Txt(s!StoreName), Txt(s!VATNumber), Txt(s!CRNumber), Txt(s!StreetName), _
                     Txt(s!BuildingNo), Txt(s!AdditionalNo), Txt(s!District), Txt(s!City), Txt(s!PostalCode), "SA")
    If standard Then
        buyer = ZxParty("Customer", "Compliance Test Buyer", "399999999800003", "", "Prince Sultan", "2322", "", _
                        "Al-Murabba", "Riyadh", "23333", "SA")
    Else
        buyer = ZxParty("Customer", "", "", "", "", "", "", "", "", "", "")
    End If
    ZatcaSampleXml = ZxInvoice("SME" & Format$(Icv, "00000"), Uuid, d, t, TypeCode, SubType, "SAR", _
        IIf(note, "SME00001", ""), Icv, Pih, seller, buyer, IIf(standard, d, ""), "10", IIf(note, "Compliance check", ""), _
        15, ZxTaxSubtotal(100, 15, "S", 0.15, "SAR"), 100, 100, 115, _
        ZxLine(1, 1, "PCE", 100, 0, 15, 115, "Compliance item", "S", 0.15, 100, "SAR"))
    Totals = Array(d & "T" & t, CCur(115), CCur(15), Txt(s!StoreName), Txt(s!VATNumber))
    s.Close
    Calendar = saved
End Function

Public Function ZatcaComplianceChecks(ByRef Report As String) As String
    ' Step 2: the six document types of "1100" signed with the compliance certificate and checked by ZATCA.
    ' "" when none is rejected; Report has one line per sample.
    Dim token As String, secret As String, cert As String, types As Variant, i As Long, pih As String, uuid As String
    Dim xml As String, totals As Variant, signed As String, hash As String, qr As String, msg As String
    Dim status As Long, reply As String, ms As Long, verdict As String, failed As Long
    token = Txt(SettingValue("ZatcaComplianceToken"))
    secret = Txt(SettingValue("ZatcaComplianceSecret"))
    If Len(token) = 0 Then
        ZatcaComplianceChecks = "اطلب شهادة الامتثال أولًا (الخطوة 1)."
        Exit Function
    End If
    cert = TokenCertificate(token)
    types = Array("388", "0100000", "فاتورة ضريبية", "381", "0100000", "إشعار دائن ضريبي", "383", "0100000", _
                  "إشعار مدين ضريبي", "388", "0200000", "فاتورة مبسطة", "381", "0200000", "إشعار دائن مبسط", _
                  "383", "0200000", "إشعار مدين مبسط")
    pih = ZATCA_INITIAL_PIH
    For i = 0 To 5
        uuid = NewUUID()
        xml = ZatcaSampleXml(types(i * 3), types(i * 3 + 1), i + 1, pih, uuid, totals)
        msg = ZatcaSignXml(xml, cert, totals, signed, hash, qr)
        If Len(msg) > 0 Then
            ZatcaComplianceChecks = msg
            Exit Function
        End If
        msg = ZatcaPost("/compliance/invoices", "{""invoiceHash"": """ & hash & """, ""uuid"": """ & uuid & _
                        """, ""invoice"": """ & Base64Text(signed) & """}", token, secret, "", status, reply, ms)
        verdict = Nz(JsonGet(reply, "validationResults.status"), "")
        If status = 0 Or Len(verdict) = 0 Then verdict = "ERROR"
        LogEInvoice "", 0, "COMPLIANCE_CHECK", "/compliance/invoices", status, IIf(verdict = "ERROR", "REJECTED", "OK"), _
                    Left$(types(i * 3 + 2) & ": " & verdict & " " & msg, 255), "", reply, ms
        Report = Report & types(i * 3 + 2) & ": " & verdict
        If verdict = "ERROR" Then
            failed = failed + 1
            Report = Report & " - " & Left$(ZatcaMessages(reply, "errorMessages") & msg, 200)
        End If
        Report = Report & vbCrLf
        pih = hash
    Next
    If failed > 0 Then
        ZatcaComplianceChecks = "لم تجتز " & failed & " من العينات فحص الامتثال."
    Else
        SaveSettings Array("ZatcaOnboardStage", "CHECKED")
    End If
End Function

Public Function ZatcaRequestProduction() As String
    ' Step 3: the production CSID that signs the real documents. "" or the problem.
    Dim token As String, secret As String, status As Long, reply As String, ms As Long, msg As String
    token = Txt(SettingValue("ZatcaComplianceToken"))
    secret = Txt(SettingValue("ZatcaComplianceSecret"))
    If Nz(SettingValue("ZatcaOnboardStage"), "") <> "CHECKED" Then
        ZatcaRequestProduction = "أكمل فحوص الامتثال أولًا (الخطوة 2)."
        Exit Function
    End If
    msg = ZatcaPost("/production/csids", "{""compliance_request_id"": """ & Txt(SettingValue("ZatcaRequestId")) & """}", _
                    token, secret, "", status, reply, ms)
    LogEInvoice "", 0, "PRODUCTION_CSID", "/production/csids", status, IIf(status = 200, "OK", "ERROR"), _
                Left$(msg & MaskSecret(reply), 255), "", MaskSecret(reply), ms
    If status <> 200 Or IsNull(JsonGet(reply, "binarySecurityToken")) Then
        ZatcaRequestProduction = "لم تُصدر الهيئة الشهادة الفعلية: " & msg & " " & Left$(reply, 300)
        Exit Function
    End If
    SaveSettings Array("ZatcaProductionToken", JsonGet(reply, "binarySecurityToken"), "ZatcaProductionSecret", _
                 JsonGet(reply, "secret"), "ZatcaCertificate", TokenCertificate(JsonGet(reply, "binarySecurityToken")), _
                 "ZatcaOnboardStage", "PRODUCTION")
End Function

'==============================================================================
' Sending (modEInvoice)
'==============================================================================
Public Function ZatcaReadyProblem() As String
    ' "" when the device is onboarded in the current environment and its key file is here.
    Dim s As DAO.Recordset
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    If Txt(s!ZatcaOnboardStage) <> "PRODUCTION" Or Len(Txt(s!ZatcaProductionToken)) = 0 Then
        ZatcaReadyProblem = "الجهاز لم يُسجَّل بعد لدى الهيئة (إعداد الربط: الخطوات 1 و2 و3)."
    ElseIf Txt(s!ZatcaOnboardEnv) <> Txt(s!EInvoiceEnvironment) Then
        ZatcaReadyProblem = "سُجّل الجهاز في بيئة أخرى. سجّله في البيئة الحالية من إعداد الربط."
    ElseIf Len(Txt(s!ZatcaKeyFile)) = 0 Then
        ZatcaReadyProblem = "ملف المفتاح الخاص غير محدد في إعداد الربط."
    ElseIf Len(Dir$(Txt(s!ZatcaKeyFile))) = 0 Then
        ZatcaReadyProblem = "ملف المفتاح الخاص غير موجود على هذا الجهاز: " & Txt(s!ZatcaKeyFile)
    End If
    s.Close
End Function

Public Function ZatcaSendDocument(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Signs the document if it is not yet (modZatcaXml), then reports it (simplified) or clears it (standard).
    ' Returns "RESULT|message" for modEInvoice.
    Dim msg As String, rs As DAO.Recordset, standard As Boolean, body As String, status As Long, reply As String
    Dim ms As Long, result As String, docStatus As String, cleared As Variant, b() As Byte, n As Long, xml As String
    Dim path As String
    msg = ZatcaPrepareDocument(DocKind, DocID)
    If Len(msg) > 0 Then
        ZatcaSendDocument = "ERROR|" & msg
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT InvoiceSubType, InvoiceUUID, InvoiceHash, EInvoiceXml FROM " & _
                                     DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID, dbOpenSnapshot)
    standard = (rs!InvoiceSubType = "STANDARD")
    body = "{""invoiceHash"": """ & rs!InvoiceHash & """, ""uuid"": """ & rs!InvoiceUUID & """, ""invoice"": """ & _
           Base64Text(rs!EInvoiceXml) & """}"
    rs.Close
    If standard Then
        path = "/invoices/clearance/single"
        msg = ZatcaPost(path, body, Txt(SettingValue("ZatcaProductionToken")), Txt(SettingValue("ZatcaProductionSecret")), _
                        "Clearance-Status: 1", status, reply, ms)
    Else
        path = "/invoices/reporting/single"
        msg = ZatcaPost(path, body, Txt(SettingValue("ZatcaProductionToken")), Txt(SettingValue("ZatcaProductionSecret")), _
                        "", status, reply, ms)
    End If
    result = ZatcaReplyResult(status, reply, standard, docStatus, msg)
    LogEInvoice DocKind, DocID, IIf(standard, "CLEAR", "REPORT"), path, status, result, msg, body, reply, ms
    If Len(docStatus) > 0 Then SetEInvoiceStatus DocKind, DocID, docStatus, reply
    cleared = JsonGet(reply, "clearedInvoice")
    If standard And (result = "OK" Or result = "WARNING") And Not IsNull(cleared) Then
        n = Base64Decode(CStr(cleared), b)                   ' the document with ZATCA's stamp: it is the invoice now
        xml = Utf8Decode(b, 0, n)
        Set rs = CurrentDb.OpenRecordset("SELECT EInvoiceXml, QRCodeData FROM " & DocTableOf(DocKind) & " WHERE " & _
                                         DocKeyOf(DocKind) & " = " & DocID, dbOpenDynaset)
        rs.Edit
        rs!EInvoiceXml = xml
        If Len(ClearedQR(xml)) > 0 Then rs!QRCodeData = ClearedQR(xml)
        rs.Update
        rs.Close
    End If
    ZatcaSendDocument = result & "|" & msg
End Function

Public Function ClearedQR(ByVal Xml As String) As String
    ' The QR code inside a cleared document (cac:AdditionalDocumentReference with cbc:ID QR).
    Dim i As Long, j As Long, k As Long
    i = InStr(1, Xml, ">QR</cbc:ID>", vbBinaryCompare)
    If i = 0 Then Exit Function
    j = InStr(i, Xml, "EmbeddedDocumentBinaryObject", vbBinaryCompare)
    If j = 0 Then Exit Function
    j = InStr(j, Xml, ">", vbBinaryCompare) + 1
    k = InStr(j, Xml, "<", vbBinaryCompare)
    If j > 1 And k > j Then ClearedQR = Mid$(Xml, j, k - j)
End Function

'==============================================================================
' frmZatcaSetup: the three steps
'==============================================================================
Public Sub ZatcaOnboardShow(ByVal frm As Access.Form)
    Dim stage As String, info As String
    stage = Nz(SettingValue("ZatcaOnboardStage"), "")
    Select Case stage
        Case "COMPLIANCE": info = "الخطوة 1 تمت: شهادة الامتثال. التالي: فحوص الامتثال."
        Case "CHECKED": info = "الخطوة 2 تمت: اجتازت العينات الفحص. التالي: الشهادة الفعلية."
        Case "PRODUCTION": info = "الجهاز مسجّل لدى الهيئة (" & EnvironmentName() & "). فعّل الإرسال من شاشة الفاتورة الإلكترونية."
        Case Else: info = "الجهاز غير مسجّل. اكتب رمز التحقق (OTP) من بوابة فاتورة ثم الخطوة 1."
    End Select
    frm!lblStage.Caption = Tr(info)
    If IsNull(frm!txtBranch.Value) Then frm!txtBranch.Value = Nz(SettingValue("ZatcaBranchName"), Tr("الفرع الرئيسي"))
    If IsNull(frm!txtIndustry.Value) Then frm!txtIndustry.Value = Nz(SettingValue("ZatcaIndustry"), "Retail")
End Sub

Public Sub ZatcaOnboardStep(ByVal frm As Access.Form, ByVal StepNo As Integer)
    Dim msg As String, report As String
    If Not HasPermission("SETTINGS") Then
        ShowWarning "إعداد الربط يحتاج صلاحية إعدادات المحل."
        Exit Sub
    End If
    If StepNo = 1 And Len(Nz(SettingValue("ZatcaOnboardStage"), "")) > 0 Then
        If Not AskYesNo("سيُنشأ مفتاح جديد ويُعاد تسجيل الجهاز من البداية. هل تريد المتابعة؟") Then Exit Sub
    End If
    DoCmd.Hourglass True
    Select Case StepNo
        Case 1: msg = ZatcaRequestCompliance(Nz(frm!txtOtp.Value, ""), Nz(frm!txtBranch.Value, ""), Nz(frm!txtIndustry.Value, ""))
        Case 2: msg = ZatcaComplianceChecks(report)
        Case 3: msg = ZatcaRequestProduction()
    End Select
    DoCmd.Hourglass False
    ZatcaSetupLoad frm
    ZatcaOnboardShow frm
    If Len(msg) > 0 Then
        ShowWarning msg & IIf(Len(report) > 0, vbCrLf & vbCrLf & report, "")
    ElseIf Len(report) > 0 Then
        ShowInfo "اجتازت العينات فحص الامتثال:" & vbCrLf & vbCrLf & report
    Else
        ShowInfo "تمت الخطوة " & StepNo & "."
    End If
End Sub

'==============================================================================
' In-Access test (nothing is sent)
'==============================================================================
Public Function TestZatcaApi() As Boolean
    Dim passed As Long, failed As Long, report As String, st As String, msg As String, ws As DAO.Workspace
    Dim inTrans As Boolean, totals As Variant, xml As String
    g_SilentMode = True
    Debug.Print "=== TestZatcaApi  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record InStr(ZatcaBaseUrl("PRODUCTION"), "/core") > 0 And InStr(ZatcaBaseUrl("TEST"), "/developer-portal") > 0, _
           "عنوان كل بيئة", passed, failed, report
    Record InStr(ZatcaCsrConfig("SIMULATION", "x", "399999999900003", "a", "b", "c", "d", "e"), _
                 "ASN1:UTF8String:PREZATCA-Code-Signing") > 0, "طلب الشهادة لبيئة المحاكاة", passed, failed, report
    Record ZatcaReplyResult(200, "{}", True, st, msg) = "OK" And st = "CLEARED" And _
           ZatcaReplyResult(200, "{}", False, st, msg) = "OK" And st = "REPORTED", "قبول الاعتماد والتبليغ", passed, failed, report
    Record ZatcaReplyResult(400, "{""validationResults"": {""errorMessages"": [{""code"": ""BR-01"", ""message"": ""x""}]}}", _
                            False, st, msg) = "REJECTED" And st = "REJECTED" And msg = "BR-01: x", "قراءة سبب الرفض", _
           passed, failed, report
    Record ZatcaReplyResult(503, "", False, st, msg) = "NETWORK" And st = "", "المنظومة غير متاحة: يبقى بانتظار الإرسال", _
           passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "UPDATE Settings SET ZatcaOnboardStage = 'CHECKED' WHERE SettingID = 1", dbFailOnError
    Record Len(ZatcaReadyProblem()) > 0, "لا إرسال قبل الشهادة الفعلية", passed, failed, report
    xml = ZatcaSampleXml("383", "0100000", 3, ZATCA_INITIAL_PIH, "00000000-0000-0000-0000-000000000003", totals)
    Record InStr(xml, ">383</cbc:InvoiceTypeCode>") > 0 And InStr(xml, "<cac:BillingReference>") > 0 And _
           InStr(xml, "<cbc:InstructionNote>") > 0, "عينة إشعار مدين ضريبي", passed, failed, report
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الربط مع فاتورة ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestZatcaApi"
        TestZatcaApi = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestZatcaApi"
    End If
End Function

Private Sub Record(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                   ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "[X]  " & Title
    End If
End Sub
