Attribute VB_Name = "modEtaInvoice"
'==============================================================================
' modEtaInvoice  -  Retail Store Management System (Egypt: e-invoice to businesses, B2B)
'
' docs/49-ETA-EInvoice.md. A sales invoice to a customer with a tax number (InvoiceSubType STANDARD) and its
' returns are e-invoices of ETA (document version 1.0), not e-receipts. modEtaReceipt hands them here:
'   EtaInvoiceSend(DocKind, DocID)       builds and signs the document, submits it; "RESULT|message"
'   EtaInvoiceRefresh(DocKind, DocID)    VALID / INVALID / CANCELLED from ETA
'   EtaCancelDocument(DocKind, DocID, Reason)   cancels a valid document (frmEInvoices)
' The signature (CAdES-BES) is made by an external signing program that reads the USB token of the taxpayer
' (Settings.EtaSignerPath + EtaSignerArgs): it reads the serialized document from {IN} and writes the base64
' signature to {OUT}. Without a signing program the pre-production system takes the unsigned version 0.9.
' tools/eta_invoice_reference.py is the Python mirror.
'==============================================================================
Option Compare Database
Option Explicit

Public Const ETA_SIGNER_ARGS As String = """{IN}"" ""{OUT}"" ""{PIN}"""
Private Const TEST_DOCUMENT_HASH As String = "c3b1cbaa6579f6c06be88e74e27fa00f3fbaa775e0749239a9986d113d148abb"   ' Python mirror

Private m_erpToken As String                            ' the access token of the program (ERP), per environment
Private m_erpEnv As String
Private m_erpExpires As Date

'==============================================================================
' The document (no data access)
'==============================================================================
Public Function EtaPartyJson(ByVal Role As String, ByVal PartyType As String, ByVal PartyId As String, _
                             ByVal PartyName As String, ByVal BranchId As String, ByVal Governate As String, _
                             ByVal City As String, ByVal Street As String, ByVal Building As String, _
                             ByVal Postal As String) As String
    ' The issuer (with its branch) or the receiver: B business, P person, F foreigner.
    Dim addr As String, s As String
    If Role = "issuer" Then addr = EtaJsonText("branchID", BranchId) & ","
    addr = addr & EtaJsonText("country", "EG") & "," & EtaJsonText("governate", Governate) & "," & _
           EtaJsonText("regionCity", City) & "," & EtaJsonText("street", Street) & "," & _
           EtaJsonText("buildingNumber", Building)
    If Len(Postal) > 0 Then addr = addr & "," & EtaJsonText("postalCode", Postal)
    s = """address"":{" & addr & "}," & EtaJsonText("type", PartyType)
    If Len(PartyId) > 0 Then s = s & "," & EtaJsonText("id", PartyId)
    EtaPartyJson = """" & Role & """:{" & s & "," & EtaJsonText("name", PartyName) & "}"
End Function

Public Function EtaLineJson(ByVal Description As String, ByVal ItemType As String, ByVal ItemCode As String, _
                            ByVal UnitType As String, ByVal Qty As Variant, ByVal InternalCode As String, _
                            ByVal UnitValue As Variant, ByVal SalesTotal As Variant, ByVal Discount As Variant, _
                            ByVal NetTotal As Variant, ByVal Tax As Variant, ByVal Rate As Variant, _
                            ByVal Total As Variant) As String
    EtaLineJson = "{" & EtaJsonText("description", Description) & "," & EtaJsonText("itemType", ItemType) & "," & _
        EtaJsonText("itemCode", ItemCode) & "," & EtaJsonText("unitType", UnitType) & "," & EtaJsonNum("quantity", Qty) & _
        "," & EtaJsonText("internalCode", InternalCode) & "," & EtaJsonNum("salesTotal", SalesTotal) & "," & _
        EtaJsonNum("total", Total) & "," & EtaJsonNum("valueDifference", 0) & "," & EtaJsonNum("totalTaxableFees", 0) & _
        "," & EtaJsonNum("netTotal", NetTotal) & "," & EtaJsonNum("itemsDiscount", 0) & ",""unitValue"":{" & _
        EtaJsonText("currencySold", "EGP") & "," & EtaJsonNum("amountEGP", UnitValue) & "},""discount"":{" & _
        EtaJsonNum("rate", 0) & "," & EtaJsonNum("amount", Discount) & "},""taxableItems"":[{" & _
        EtaJsonText("taxType", "T1") & "," & EtaJsonNum("amount", Tax) & "," & EtaJsonText("subType", "V009") & "," & _
        EtaJsonNum("rate", Rate) & "}]}"
End Function

Public Function EtaInvoiceJson(ByVal DocType As String, ByVal Version As String, ByVal DateTimeIssued As String, _
                               ByVal ActivityCode As String, ByVal InternalId As String, ByVal IssuerJson As String, _
                               ByVal ReceiverJson As String, ByVal LinesJson As String, ByVal ReferenceUuid As String, _
                               ByVal TotalDiscount As Variant, ByVal TotalSales As Variant, ByVal NetAmount As Variant, _
                               ByVal TaxTotal As Variant, ByVal TotalAmount As Variant) As String
    ' The document without its signature: I invoice, C credit note (ReferenceUuid = the invoice).
    Dim s As String
    s = "{" & IssuerJson & "," & ReceiverJson & "," & EtaJsonText("documentType", DocType) & "," & _
        EtaJsonText("documentTypeVersion", Version) & "," & EtaJsonText("dateTimeIssued", DateTimeIssued) & "," & _
        EtaJsonText("taxpayerActivityCode", ActivityCode) & "," & EtaJsonText("internalID", InternalId)
    If Len(ReferenceUuid) > 0 Then s = s & ",""references"":[""" & JsonEscape(ReferenceUuid) & """]"
    EtaInvoiceJson = s & ",""invoiceLines"":[" & LinesJson & "]," & EtaJsonNum("totalDiscountAmount", TotalDiscount) & _
        "," & EtaJsonNum("totalSalesAmount", TotalSales) & "," & EtaJsonNum("netAmount", NetAmount) & _
        ",""taxTotals"":[{" & EtaJsonText("taxType", "T1") & "," & EtaJsonNum("amount", TaxTotal) & "}]," & _
        EtaJsonNum("totalAmount", TotalAmount) & "," & EtaJsonNum("extraDiscountAmount", 0) & "," & _
        EtaJsonNum("totalItemsDiscountAmount", 0) & "}"
End Function

Public Function EtaWithSignature(ByVal Json As String, ByVal Signature As String) As String
    EtaWithSignature = Left$(Json, Len(Json) - 1) & ",""signatures"":[{" & EtaJsonText("signatureType", "I") & "," & _
                       EtaJsonText("value", Signature) & "}]}"
End Function

Public Function EtaSignerCommand(ByVal Template As String, ByVal Exe As String, ByVal InFile As String, _
                                 ByVal OutFile As String, ByVal Pin As String) As String
    ' The command line of the signing program: the quoted program, then the template with {IN} {OUT} {PIN}.
    If Len(Template) = 0 Then Template = ETA_SIGNER_ARGS
    EtaSignerCommand = """" & Exe & """ " & Replace(Replace(Replace(Template, "{IN}", InFile), "{OUT}", OutFile), _
                                                   "{PIN}", Pin)
End Function

Public Function EtaInvoiceQrLink(ByVal Env As String, ByVal Uuid As String, ByVal LongId As String) As String
    ' The public link of the document on the ETA portal.
    EtaInvoiceQrLink = EtaPortalUrl(Env) & "/documents/" & Uuid & "/share/" & LongId
End Function

'==============================================================================
' ETA replies
'==============================================================================
Public Function EtaDocSubmitResult(ByVal HttpStatus As Long, ByVal Body As String, ByRef DocStatus As String, _
                                   ByRef Message As String, ByRef SubmissionId As String, ByRef Uuid As String, _
                                   ByRef LongId As String) As String
    ' The result of a document submission: OK (SUBMITTED, with ETA's UUID and long ID), REJECTED, ERROR or NETWORK.
    DocStatus = ""
    Message = ""
    SubmissionId = ""
    Uuid = ""
    LongId = ""
    If HttpStatus = 0 Then
        EtaDocSubmitResult = "NETWORK"
        Message = "تعذّر الاتصال بالمنظومة."
    ElseIf HttpStatus = 200 Or HttpStatus = 202 Then
        If Not IsNull(JsonGet(Body, "rejectedDocuments[0]")) Then
            EtaDocSubmitResult = "REJECTED"
            DocStatus = "REJECTED"
            Message = EtaErrors(Body, "rejectedDocuments[0].error")
            If Len(Message) = 0 Then Message = "رُفض المستند."
        ElseIf Not IsNull(JsonGet(Body, "acceptedDocuments[0]")) Then
            EtaDocSubmitResult = "OK"
            DocStatus = "SUBMITTED"
            SubmissionId = EtaFirstText(Body, "submissionId", "", "")
            Uuid = EtaFirstText(Body, "acceptedDocuments[0].uuid", "", "")
            LongId = EtaFirstText(Body, "acceptedDocuments[0].longId", "", "")
        Else
            EtaDocSubmitResult = "ERROR"
            Message = "رد غير متوقع من المصلحة: " & Left$(Body, 200)
        End If
    ElseIf HttpStatus = 400 Then
        EtaDocSubmitResult = "REJECTED"
        DocStatus = "REJECTED"
        Message = EtaErrors(Body, "error")
        If Len(Message) = 0 Then Message = Left$(Body, 200)
    ElseIf HttpStatus = 401 Or HttpStatus = 403 Then
        EtaDocSubmitResult = "ERROR"
        Message = "رفضت المصلحة بيانات دخول البرنامج (Client ID و Client Secret)."
    ElseIf HttpStatus = 429 Or HttpStatus >= 500 Then
        EtaDocSubmitResult = "NETWORK"
        Message = "المنظومة غير متاحة الآن (رمز " & HttpStatus & ")."
    Else
        EtaDocSubmitResult = "ERROR"
        Message = "رد غير متوقع من المصلحة (رمز " & HttpStatus & "): " & Left$(Body, 200)
    End If
End Function

Public Function EtaDocStatusResult(ByVal Body As String, ByRef Message As String) As String
    ' VALID, INVALID (Message: the failed validation steps), CANCELLED or "" (in progress).
    Dim st As String, i As Long, p As String, n As Long
    Message = ""
    st = LCase$(EtaFirstText(Body, "status", "", ""))
    If st = "valid" Then
        EtaDocStatusResult = "VALID"
    ElseIf st = "cancelled" Or st = "canceled" Then
        EtaDocStatusResult = "CANCELLED"
    ElseIf st = "invalid" Then
        EtaDocStatusResult = "INVALID"
        For i = 0 To 9
            p = "validationResults.validationSteps[" & i & "]"
            If IsNull(JsonGet(Body, p)) Or n = 3 Then Exit For
            If LCase$(EtaFirstText(Body, p & ".status", "", "")) = "invalid" Then
                If n > 0 Then Message = Message & "; "
                Message = Message & EtaFirstText(Body, p & ".name", "", "") & ": " & _
                          EtaFirstText(Body, p & ".error.innerError[0].error", p & ".error.error", p & ".error.errorCode")
                n = n + 1
            End If
        Next
    End If
End Function

'==============================================================================
' The document of a sale or return
'==============================================================================
Private Function Txt(ByVal v As Variant) As String
    Txt = Trim$(Nz(v, ""))
End Function

Private Function CurrentEnv() As String
    CurrentEnv = EtaEnv(Nz(SettingValue("EInvoiceEnvironment"), "TEST"))
End Function

Private Function LinesOf(ByVal DocKind As String) As String
    If DocKind = "RETURN" Then
        LinesOf = "SalesReturnDetails"
    Else
        LinesOf = "SalesInvoiceDetails"
    End If
End Function

Public Function EtaIsInvoice(ByVal DocKind As String, ByVal DocID As Long) As Boolean
    ' A tax invoice (customer with a tax number) and its returns: e-invoice, not e-receipt.
    EtaIsInvoice = (Nz(DbValue("SELECT InvoiceSubType FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & _
                               " = " & DocID), "") = "STANDARD")
End Function

Public Function EtaSignedVersion() As Boolean
    ' Version 1.0 (signed) in production or when a signing program is set; else the unsigned 0.9 (pre-production).
    EtaSignedVersion = (CurrentEnv() = "PROD" Or Len(Txt(SettingValue("EtaSignerPath"))) > 0)
End Function

Public Function EtaInvoiceDataProblem(ByVal DocKind As String, ByVal DocID As Long) As String
    ' "" when the document can be built: the issuer, the receiver (tax number and address), the ETA code of every
    ' product, the invoice of a credit note sent first, the credentials of the program and the signing program.
    Dim msg As String, c As DAO.Recordset, rs As DAO.Recordset, names As String, n As Long
    msg = EtaIssuerProblem()
    If Len(Txt(SettingValue("EtaErpClientId"))) = 0 Or Len(Txt(SettingValue("EtaErpClientSecret"))) = 0 Then
        msg = msg & "- بيانات دخول البرنامج للفاتورة الإلكترونية (Client ID و Client Secret)" & vbCrLf
    End If
    If CurrentEnv() = "PROD" And Len(Txt(SettingValue("EtaSignerPath"))) = 0 Then
        msg = msg & "- برنامج التوقيع (فلاشة التوقيع) في إعداد الربط" & vbCrLf
    End If
    Set c = CurrentDb.OpenRecordset("SELECT c.* FROM " & DocTableOf(DocKind) & " AS h INNER JOIN Customers AS c ON " & _
                                    "h.CustomerID = c.CustomerID WHERE h." & DocKeyOf(DocKind) & " = " & DocID, dbOpenSnapshot)
    If Not c.EOF Then
        If Len(Txt(c!VATNumber)) = 0 Or Len(TaxNumberProblem(c!VATNumber, "EG")) > 0 Then
            msg = msg & "- رقم التسجيل الضريبي للعميل (9 أرقام)" & vbCrLf
        End If
        If Len(Txt(c!City)) = 0 Then msg = msg & "- مدينة العميل" & vbCrLf
        If Len(Txt(c!StreetName)) = 0 Then msg = msg & "- شارع العميل" & vbCrLf
        If Len(Txt(c!BuildingNo)) = 0 Then msg = msg & "- رقم مبنى العميل" & vbCrLf
    End If
    c.Close
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT p.ProductName FROM " & LinesOf(DocKind) & " AS d INNER JOIN " & _
                                     "Products AS p ON d.ProductID = p.ProductID WHERE d." & DocKeyOf(DocKind) & " = " & _
                                     DocID & " AND (p.EtaItemCode IS NULL OR p.EtaItemCode = '')", dbOpenSnapshot)
    Do Until rs.EOF
        n = n + 1
        If n <= 5 Then names = names & IIf(n > 1, "، ", "") & Txt(rs!ProductName)
        rs.MoveNext
    Loop
    rs.Close
    If n > 0 Then msg = msg & "- كود المصلحة (EGS أو GS1) للأصناف: " & names & IIf(n > 5, " ...", "") & vbCrLf
    If DocKind = "RETURN" Then
        If Len(Nz(DbValue("SELECT o.EtaUUID FROM SalesReturns AS r INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = " & _
                          "o.SalesInvoiceID WHERE r.SalesReturnID = " & DocID), "")) = 0 Then
            msg = msg & "- أرسل الفاتورة الأصلية أولًا" & vbCrLf
        End If
    End If
    If Len(msg) > 0 Then EtaInvoiceDataProblem = "بيانات ناقصة للفاتورة الإلكترونية:" & vbCrLf & msg
End Function

Public Function EtaInvoiceDocumentJson(ByVal DocKind As String, ByVal DocID As Long, ByVal Version As String) As String
    ' The unsigned document of a sales invoice (I) or of its return (C, referencing the invoice).
    Dim db As DAO.Database, h As DAO.Recordset, s As DAO.Recordset, c As DAO.Recordset, rs As DAO.Recordset
    Dim isReturn As Boolean, issuer As String, receiver As String, lines As String, refUuid As String
    Dim totalSales As Currency, totalDisc As Currency, branch As String, region As String
    Set db = CurrentDb
    isReturn = (DocKind = "RETURN")
    If isReturn Then
        Set h = db.OpenRecordset("SELECT r.*, r.ReturnNumber AS DocNumber, r.ReturnDate AS DocDate, o.EtaUUID AS OriginalUUID " & _
                                 "FROM SalesReturns AS r INNER JOIN SalesInvoices AS o ON r.SalesInvoiceID = o.SalesInvoiceID " & _
                                 "WHERE r.SalesReturnID = " & DocID, dbOpenSnapshot)
        refUuid = Txt(h!OriginalUUID)
    Else
        Set h = db.OpenRecordset("SELECT h.*, h.InvoiceNumber AS DocNumber, h.InvoiceDate AS DocDate FROM SalesInvoices AS h " & _
                                 "WHERE h.SalesInvoiceID = " & DocID, dbOpenSnapshot)
    End If
    Set s = db.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    branch = Txt(s!EtaBranchCode)
    If Len(branch) = 0 Then branch = "0"
    issuer = EtaPartyJson("issuer", "B", Txt(s!VATNumber), Txt(s!StoreName), branch, Txt(s!EtaGovernate), Txt(s!City), _
                          Txt(s!StreetName), Txt(s!BuildingNo), Txt(s!PostalCode))
    Set c = db.OpenRecordset("SELECT * FROM Customers WHERE CustomerID = " & h!CustomerID, dbOpenSnapshot)
    region = Txt(c!District)                              ' the customer has a city and a district: governate, city
    If Len(region) = 0 Then region = Txt(c!City)
    receiver = EtaPartyJson("receiver", "B", Txt(c!VATNumber), Txt(c!CustomerName), "", Txt(c!City), region, _
                            Txt(c!StreetName), Txt(c!BuildingNo), Txt(c!PostalCode))
    c.Close
    Set rs = db.OpenRecordset("SELECT d.Quantity, d.UnitPrice, d.Discount, d.NetAmount, d.VATRate, d.Tax, d.LineTotal, " & _
                              "p.ProductCode, p.ProductName, p.EtaItemType, p.EtaItemCode, u.EtaUnitCode FROM (" & LinesOf(DocKind) & _
                              " AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID) INNER JOIN Units AS u ON " & _
                              "p.UnitID = u.UnitID WHERE d." & DocKeyOf(DocKind) & " = " & DocID & " ORDER BY d." & _
                              IIf(isReturn, "ReturnDetailID", "SalesDetailID"), dbOpenSnapshot)
    Do Until rs.EOF
        If Len(lines) > 0 Then lines = lines & ","
        lines = lines & EtaLineJson(Txt(rs!ProductName), Nz(rs!EtaItemType, "EGS"), Txt(rs!EtaItemCode), Nz(rs!EtaUnitCode, "EA"), _
                                    rs!Quantity, Txt(rs!ProductCode), rs!UnitPrice, rs!NetAmount + Nz(rs!Discount, 0), _
                                    Nz(rs!Discount, 0), rs!NetAmount, rs!Tax, CDec(rs!VATRate) * 100, rs!LineTotal)
        totalSales = totalSales + rs!NetAmount + Nz(rs!Discount, 0)
        totalDisc = totalDisc + Nz(rs!Discount, 0)
        rs.MoveNext
    Loop
    rs.Close
    EtaInvoiceDocumentJson = EtaInvoiceJson(IIf(isReturn, "C", "I"), Version, EtaUtcTimestamp(h!DocDate), _
        Txt(s!EtaActivityCode), Txt(h!DocNumber), issuer, receiver, lines, refUuid, totalDisc, totalSales, _
        h!TaxableAmount, h!Tax, h!TotalAmount)
    s.Close
    h.Close
End Function

Public Function EtaSignDocument(ByVal Serialized As String, ByRef Signature As String) As String
    ' The CAdES-BES signature (base64) of the serialized document by the signing program. "" or the problem.
    Dim inFile As String, outFile As String, cmd As String, rc As Long, f As Integer, ln As String
    inFile = ZatcaWorkFolder() & "\eta-document.txt"
    outFile = ZatcaWorkFolder() & "\eta-signature.txt"
    If Len(Dir$(outFile)) > 0 Then Kill outFile
    WriteUtf8File inFile, Serialized
    cmd = EtaSignerCommand(Txt(SettingValue("EtaSignerArgs")), Txt(SettingValue("EtaSignerPath")), inFile, outFile, _
                           Txt(SettingValue("EtaTokenPin")))
    On Error GoTo EH
    rc = CreateObject("WScript.Shell").Run(cmd, 0, True)
    On Error GoTo 0
    If Len(Dir$(outFile)) = 0 Then
        EtaSignDocument = "لم يُنشئ برنامج التوقيع ملف التوقيع (رمز " & rc & "). تأكد من فلاشة التوقيع ورقمها السري."
        Exit Function
    End If
    f = FreeFile
    Open outFile For Input As #f
    Do Until EOF(f)
        Line Input #f, ln
        Signature = Signature & Trim$(ln)
    Loop
    Close #f
    Kill outFile
    Kill inFile
    If Len(Signature) = 0 Then EtaSignDocument = "ملف التوقيع فارغ (رمز " & rc & ")."
    Exit Function
EH:
    EtaSignDocument = "تعذّر تشغيل برنامج التوقيع: " & Err.Description
End Function

Public Function EtaInvoicePrepare(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Builds (and signs) the document once; a rejected or invalid document is built again. "" or the problem.
    Dim json As String, signature As String, msg As String, st As String, rs As DAO.Recordset
    st = Nz(DbValue("SELECT ZatcaStatus FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "")
    If Len(Nz(DbValue("SELECT EInvoiceXml FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), _
              "")) > 0 And st <> "REJECTED" And st <> "INVALID" Then Exit Function
    If EtaSignedVersion() Then
        json = EtaInvoiceDocumentJson(DocKind, DocID, "1.0")
        msg = EtaSignDocument(EtaSerialize(json), signature)
        If Len(msg) > 0 Then
            EtaInvoicePrepare = msg
            Exit Function
        End If
        json = EtaWithSignature(json, signature)
    Else
        json = EtaInvoiceDocumentJson(DocKind, DocID, "0.9")
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT EInvoiceXml FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & _
                                     " = " & DocID, dbOpenDynaset)
    rs.Edit
    rs!EInvoiceXml = json                                 ' a memo: the document as it is sent
    rs.Update
    rs.Close
End Function

'==============================================================================
' Requests
'==============================================================================
Public Function EtaErpToken(ByRef Token As String) As String
    ' The access token of the program (client credentials of the ERP). "" or "RESULT|message".
    Dim body As String, status As Long, reply As String, msg As String, started As Single, ms As Long
    Dim result As String, url As String, life As Long
    If Len(m_erpToken) > 0 And m_erpEnv = CurrentEnv() And Now < m_erpExpires Then
        Token = m_erpToken
        Exit Function
    End If
    body = "grant_type=client_credentials&client_id=" & EtaUrlEncode(Txt(SettingValue("EtaErpClientId"))) & _
           "&client_secret=" & EtaUrlEncode(Txt(SettingValue("EtaErpClientSecret")))
    url = EtaTokenUrl(CurrentEnv())
    started = Timer
    msg = HttpSend("POST", url, "Content-Type: application/x-www-form-urlencoded" & vbLf & "Accept: application/json", _
                   body, status, reply, 30)
    ms = CLng((Timer - started) * 1000)
    If status = 200 And Not IsNull(JsonGet(reply, "access_token")) Then
        result = "OK"
        m_erpToken = JsonGet(reply, "access_token")
        m_erpEnv = CurrentEnv()
        life = Val(Nz(JsonGet(reply, "expires_in"), "3600"))
        m_erpExpires = DateAdd("s", life - 60, Now)
        Token = m_erpToken
    ElseIf status = 0 Or status = 429 Or status >= 500 Then
        result = "NETWORK"
        If Len(msg) = 0 Then msg = "المنظومة غير متاحة الآن (رمز " & status & ")."
    Else
        result = "ERROR"
        msg = "رفضت المصلحة بيانات دخول البرنامج (رمز " & status & "): " & EtaFirstText(reply, "error_description", "error", "")
    End If
    ' the request body carries the client secret: it is not logged
    LogEInvoice "", 0, "TOKEN_ERP", url, status, result, Left$(msg, 255), "", EtaMaskToken(reply), ms
    If result <> "OK" Then EtaErpToken = result & "|" & msg
End Function

Public Sub EtaErpTokenReset()
    m_erpToken = ""
End Sub

Public Function EtaInvoiceSend(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Builds and signs the document if needed, then submits it. "RESULT|message" for modEInvoice.
    Dim msg As String, token As String, body As String, status As Long, reply As String, ms As Long, started As Single
    Dim result As String, docStatus As String, subId As String, uuid As String, longId As String, url As String
    Dim rs As DAO.Recordset
    msg = EtaInvoiceDataProblem(DocKind, DocID)
    If Len(msg) = 0 Then msg = EtaInvoicePrepare(DocKind, DocID)
    If Len(msg) > 0 Then
        EtaInvoiceSend = "ERROR|" & msg
        Exit Function
    End If
    msg = EtaErpToken(token)
    If Len(msg) > 0 Then
        EtaInvoiceSend = msg
        Exit Function
    End If
    body = "{""documents"":[" & Nz(DbValue("SELECT EInvoiceXml FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & _
                                           " = " & DocID), "") & "]}"
    url = EtaApiUrl(CurrentEnv()) & "/api/v1.0/documentsubmissions"
    started = Timer
    msg = HttpSend("POST", url, "Content-Type: application/json" & vbLf & "Accept: application/json" & vbLf & _
                   "Authorization: Bearer " & token, body, status, reply, 60)
    ms = CLng((Timer - started) * 1000)
    result = EtaDocSubmitResult(status, reply, docStatus, msg, subId, uuid, longId)
    If status = 401 Then m_erpToken = ""
    LogEInvoice DocKind, DocID, "SUBMIT_DOC", url, status, result, Left$(msg, 255), body, reply, ms
    If Len(docStatus) > 0 Then SetEInvoiceStatus DocKind, DocID, docStatus, reply
    If Len(uuid) > 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT EtaUUID, EtaLongId, EtaSubmissionId, QRCodeData FROM " & DocTableOf(DocKind) & _
                                         " WHERE " & DocKeyOf(DocKind) & " = " & DocID, dbOpenDynaset)
        rs.Edit
        rs!EtaUUID = uuid
        rs!EtaLongId = IIf(Len(longId) = 0, Null, longId)
        rs!EtaSubmissionId = IIf(Len(subId) = 0, Null, subId)
        rs!QRCodeData = EtaInvoiceQrLink(CurrentEnv(), uuid, longId)
        rs.Update
        rs.Close
    End If
    EtaInvoiceSend = result & "|" & msg
End Function

Public Function EtaInvoiceRefresh(ByVal DocKind As String, ByVal DocID As Long) As String
    ' VALID / INVALID / CANCELLED of a submitted document. "RESULT|message"; still in progress = OK.
    Dim uuid As String, token As String, msg As String, url As String, status As Long, reply As String
    Dim ms As Long, started As Single, st As String, result As String
    uuid = Nz(DbValue("SELECT EtaUUID FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "")
    If Len(uuid) = 0 Then
        EtaInvoiceRefresh = "ERROR|" & "لم يُرسل المستند بعد."
        Exit Function
    End If
    msg = EtaErpToken(token)
    If Len(msg) > 0 Then
        EtaInvoiceRefresh = msg
        Exit Function
    End If
    url = EtaApiUrl(CurrentEnv()) & "/api/v1.0/documents/" & uuid & "/details"
    started = Timer
    msg = HttpSend("GET", url, "Accept: application/json" & vbLf & "Authorization: Bearer " & token, "", status, reply, 30)
    ms = CLng((Timer - started) * 1000)
    If status = 200 Then
        st = EtaDocStatusResult(reply, msg)
        If st = "INVALID" Then
            result = "REJECTED"
        Else
            result = "OK"
        End If
        If Len(st) > 0 Then SetEInvoiceStatus DocKind, DocID, st, Left$(reply, 60000)
    ElseIf status = 0 Or status = 429 Or status >= 500 Then
        result = "NETWORK"
        If Len(msg) = 0 Then msg = "المنظومة غير متاحة الآن (رمز " & status & ")."
    Else
        If status = 401 Then m_erpToken = ""
        result = "ERROR"
        msg = "رد غير متوقع من المصلحة (رمز " & status & "): " & Left$(reply, 200)
    End If
    LogEInvoice DocKind, DocID, "STATUS_DOC", url, status, result, Left$(msg, 255), "", Left$(reply, 60000), ms
    EtaInvoiceRefresh = result & "|" & msg
End Function

Public Function EtaCancelDocument(ByVal DocKind As String, ByVal DocID As Long, ByVal Reason As String) As String
    ' Cancels a valid e-invoice (ETA allows it for a limited time after it became valid). "RESULT|message".
    Dim uuid As String, token As String, msg As String, url As String, status As Long, reply As String, body As String
    Dim ms As Long, started As Single, result As String
    If Not EtaIsInvoice(DocKind, DocID) Then
        EtaCancelDocument = "ERROR|" & "الإلغاء للفاتورة الإلكترونية فقط. الإيصال يُصحَّح بمرتجع."
        Exit Function
    End If
    If Nz(DbValue("SELECT ZatcaStatus FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "") <> "VALID" Then
        EtaCancelDocument = "ERROR|" & "يُلغى المستند الصالح فقط."
        Exit Function
    End If
    uuid = Nz(DbValue("SELECT EtaUUID FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "")
    msg = EtaErpToken(token)
    If Len(msg) > 0 Then
        EtaCancelDocument = msg
        Exit Function
    End If
    url = EtaApiUrl(CurrentEnv()) & "/api/v1.0/documents/state/" & uuid & "/state"
    body = "{" & EtaJsonText("status", "cancelled") & "," & EtaJsonText("reason", Reason) & "}"
    started = Timer
    msg = HttpSend("PUT", url, "Content-Type: application/json" & vbLf & "Accept: application/json" & vbLf & _
                   "Authorization: Bearer " & token, body, status, reply, 30)
    ms = CLng((Timer - started) * 1000)
    If status = 200 Or status = 202 Or status = 204 Then
        result = "OK"
        SetEInvoiceStatus DocKind, DocID, "CANCELLED", reply
    ElseIf status = 0 Or status = 429 Or status >= 500 Then
        result = "NETWORK"
        If Len(msg) = 0 Then msg = "المنظومة غير متاحة الآن (رمز " & status & ")."
    Else
        If status = 401 Then m_erpToken = ""
        result = "ERROR"
        msg = EtaErrors(reply, "error")
        If Len(msg) = 0 Then msg = "رفضت المصلحة الإلغاء (رمز " & status & "): " & Left$(reply, 200)
    End If
    LogEInvoice DocKind, DocID, "CANCEL", url, status, result, Left$(msg, 255), body, reply, ms
    EtaCancelDocument = result & "|" & msg
End Function

'==============================================================================
' In-Access test (nothing is sent nor signed)
'==============================================================================
Public Function TestEtaInvoice() As Boolean
    Dim passed As Long, failed As Long, report As String, st As String, msg As String, subId As String
    Dim uuid As String, longId As String, json As String
    g_SilentMode = True
    Debug.Print "=== TestEtaInvoice  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    json = EtaInvoiceJson("C", "1.0", "2026-10-09T10:00:00Z", "4711", "RET-1", _
                          EtaPartyJson("issuer", "B", "123456789", "Store", "0", "Cairo", "Nasr City", "Abbas", "12", ""), _
                          EtaPartyJson("receiver", "B", "987654321", "Buyer", "", "Giza", "Dokki", "Tahrir", "5", ""), _
                          EtaLineJson("Tea", "EGS", "EG-123456789-1", "EA", 2, "P1", 10, 20, 0, 20, 2.8, 14, 22.8), _
                          "u-1", 0, 20, 20, 2.8, 22.8)
    Record InStr(json, """documentType"":""C""") > 0 And InStr(json, """references"":[""u-1""]") > 0 And _
           InStr(json, """branchID"":""0""") > 0, "إشعار دائن يشير إلى الفاتورة", passed, failed, report
    Record EtaReceiptUuid(json) = TEST_DOCUMENT_HASH, "تسلسل المستند يساوي النسخة المكتوبة بـ Python", passed, failed, report
    Record EtaWithSignature(json, "SIG") = Left$(json, Len(json) - 1) & _
           ",""signatures"":[{""signatureType"":""I"",""value"":""SIG""}]}", "إضافة التوقيع", passed, failed, report
    Record EtaDocSubmitResult(202, "{""submissionId"": ""S"", ""acceptedDocuments"": [{""uuid"": ""U"", ""longId"": ""L""}]}", _
                              st, msg, subId, uuid, longId) = "OK" And uuid = "U" And longId = "L", "قبول المستند", _
           passed, failed, report
    Record EtaDocStatusResult("{""status"": ""Cancelled""}", msg) = "CANCELLED", "قراءة الإلغاء", passed, failed, report
    Record EtaSignerCommand("", "C:\s.exe", "a", "b", "1") = """C:\s.exe"" ""a"" ""b"" ""1""", "أمر برنامج التوقيع", _
           passed, failed, report
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات الفاتورة الإلكترونية المصرية ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestEtaInvoice"
        TestEtaInvoice = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestEtaInvoice"
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
