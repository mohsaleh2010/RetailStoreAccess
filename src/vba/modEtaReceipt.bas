Attribute VB_Name = "modEtaReceipt"
'==============================================================================
' modEtaReceipt  -  Retail Store Management System (Egypt: e-receipt of the point of sale)
'
' docs/48-ETA-EReceipt.md. The Egyptian platform module of modEInvoice (Egyptian Tax Authority, ETA, receipt
' version 1.2):
'   EtaReadyProblem()                    "" when the device data and the credentials of the POS are set
'   EtaSendDocument(DocKind, DocID)      builds the receipt once (JSON, UUID, chain, QR link) and submits it
'   EtaRefreshStatus(DocKind, DocID)     reads the result of a submitted receipt: VALID / INVALID
' The receipt has no signature: its UUID is the SHA-256 of its ETA serialization (EtaSerialize) and it carries
' the UUID of the previous receipt of the device (Settings.EtaLastUUID). The QR code of the receipt is its link on
' the ETA portal. tools/eta_receipt_reference.py is the Python mirror.
'==============================================================================
Option Compare Database
Option Explicit

Private Const TYPE_VERSION As String = "1.2"
Private Const BUYER_ID_LIMIT As Currency = 150000      ' a person buying for this total or more gives the ID
Private Const EG_UTC_OFFSET_HOURS As Integer = 2        ' used only when Windows cannot convert to UTC
Private Const TEST_RECEIPT_UUID As String = "79e73daedbb5176cbbbca5ab2ef71af609e169860cbaad3f6cc7b77faef836ec"   ' Python mirror

Private m_token As String                               ' the access token of the POS, kept for its lifetime
Private m_tokenEnv As String
Private m_tokenExpires As Date

'==============================================================================
' Environment and values (no data access)
'==============================================================================
Public Function EtaEnv(ByVal Environment As String) As String
    ' TEST and SIMULATION use the pre-production system of ETA.
    If Environment = "PRODUCTION" Then
        EtaEnv = "PROD"
    Else
        EtaEnv = "PREPROD"
    End If
End Function

Public Function EtaTokenUrl(ByVal Env As String) As String
    If Env = "PROD" Then
        EtaTokenUrl = "https://id.eta.gov.eg/connect/token"
    Else
        EtaTokenUrl = "https://id.preprod.eta.gov.eg/connect/token"
    End If
End Function

Public Function EtaApiUrl(ByVal Env As String) As String
    If Env = "PROD" Then
        EtaApiUrl = "https://api.invoicing.eta.gov.eg"
    Else
        EtaApiUrl = "https://api.preprod.invoicing.eta.gov.eg"
    End If
End Function

Public Function EtaPortalUrl(ByVal Env As String) As String
    If Env = "PROD" Then
        EtaPortalUrl = "https://invoicing.eta.gov.eg"
    Else
        EtaPortalUrl = "https://preprod.invoicing.eta.gov.eg"
    End If
End Function

Public Function EtaNum(ByVal Value As Variant) As String
    ' A JSON number: at most 5 decimals, no trailing zero ("12.5", "100", "0.00125").
    Dim d As Variant, whole As Variant, frac As String, sign As String
    d = Round(CDec(Value), 5)
    If d = 0 Then
        EtaNum = "0"
        Exit Function
    End If
    If d < 0 Then
        sign = "-"
        d = -d
    End If
    whole = Fix(d)
    frac = Right$("00000" & CStr(CLng((d - whole) * 100000)), 5)
    Do While Len(frac) > 0 And Right$(frac, 1) = "0"
        frac = Left$(frac, Len(frac) - 1)
    Loop
    EtaNum = sign & CStr(whole)
    If Len(frac) > 0 Then EtaNum = EtaNum & "." & frac
End Function

Public Function EtaJsonText(ByVal PropName As String, ByVal Value As String) As String
    EtaJsonText = """" & PropName & """:""" & JsonEscape(Value) & """"
End Function

Public Function EtaJsonNum(ByVal PropName As String, ByVal Value As Variant) As String
    EtaJsonNum = """" & PropName & """:" & EtaNum(Value)
End Function

Public Function EtaItemJson(ByVal InternalCode As String, ByVal Description As String, ByVal ItemType As String, _
                            ByVal ItemCode As String, ByVal UnitType As String, ByVal Qty As Variant, _
                            ByVal UnitPrice As Variant, ByVal TotalSale As Variant, ByVal Discount As Variant, _
                            ByVal NetSale As Variant, ByVal Tax As Variant, ByVal Rate As Variant, _
                            ByVal Total As Variant) As String
    ' One line of the receipt: VAT is tax type T1, sub type V009 (general goods).
    Dim s As String
    s = "{" & EtaJsonText("internalCode", InternalCode) & "," & EtaJsonText("description", Description) & "," & _
        EtaJsonText("itemType", ItemType) & "," & EtaJsonText("itemCode", ItemCode) & "," & _
        EtaJsonText("unitType", UnitType) & "," & EtaJsonNum("quantity", Qty) & "," & _
        EtaJsonNum("unitPrice", UnitPrice) & "," & EtaJsonNum("netSale", NetSale) & "," & _
        EtaJsonNum("totalSale", TotalSale) & "," & EtaJsonNum("total", Total)
    If CDec(Discount) <> 0 Then
        s = s & ",""commercialDiscountData"":[{" & EtaJsonNum("amount", Discount) & "," & _
            EtaJsonText("description", "Discount") & "}]"
    End If
    EtaItemJson = s & ",""taxableItems"":[{" & EtaJsonText("taxType", "T1") & "," & EtaJsonNum("amount", Tax) & "," & _
                  EtaJsonText("subType", "V009") & "," & EtaJsonNum("rate", Rate) & "}]}"
End Function

Public Function EtaSellerJson(ByVal Rin As String, ByVal TradeName As String, ByVal BranchCode As String, _
                              ByVal Governate As String, ByVal City As String, ByVal Street As String, _
                              ByVal Building As String, ByVal Postal As String, ByVal DeviceSerial As String, _
                              ByVal ActivityCode As String) As String
    Dim addr As String
    addr = EtaJsonText("country", "EG") & "," & EtaJsonText("governate", Governate) & "," & _
           EtaJsonText("regionCity", City) & "," & EtaJsonText("street", Street) & "," & _
           EtaJsonText("buildingNumber", Building)
    If Len(Postal) > 0 Then addr = addr & "," & EtaJsonText("postalCode", Postal)
    EtaSellerJson = """seller"":{" & EtaJsonText("rin", Rin) & "," & EtaJsonText("companyTradeName", TradeName) & "," & _
                    EtaJsonText("branchCode", BranchCode) & ",""branchAddress"":{" & addr & "}," & _
                    EtaJsonText("deviceSerialNumber", DeviceSerial) & "," & EtaJsonText("activityCode", ActivityCode) & "}"
End Function

Public Function EtaBuyerJson(ByVal BuyerType As String, ByVal BuyerId As String, ByVal BuyerName As String, _
                             ByVal Mobile As String) As String
    ' P person, B business (its tax number), F foreigner.
    Dim s As String
    s = EtaJsonText("type", BuyerType)
    If Len(BuyerId) > 0 Then s = s & "," & EtaJsonText("id", BuyerId)
    If Len(BuyerName) > 0 Then s = s & "," & EtaJsonText("name", BuyerName)
    If Len(Mobile) > 0 Then s = s & "," & EtaJsonText("mobileNumber", Mobile)
    EtaBuyerJson = """buyer"":{" & s & "}"
End Function

Public Function EtaReceiptJson(ByVal DateTimeIssued As String, ByVal Number As String, ByVal Uuid As String, _
                               ByVal PreviousUuid As String, ByVal ReferenceUuid As String, ByVal SellerJson As String, _
                               ByVal BuyerJson As String, ByVal ItemsJson As String, ByVal TotalSales As Variant, _
                               ByVal TotalDiscount As Variant, ByVal NetAmount As Variant, ByVal TaxTotal As Variant, _
                               ByVal TotalAmount As Variant, ByVal PaymentMethod As String) As String
    ' The receipt (a return when ReferenceUuid, the UUID of the sale, is given). ItemsJson: the lines, comma-separated.
    Dim header As String, kind As String
    header = EtaJsonText("dateTimeIssued", DateTimeIssued) & "," & EtaJsonText("receiptNumber", Number) & "," & _
             EtaJsonText("uuid", Uuid) & "," & EtaJsonText("previousUUID", PreviousUuid)
    If Len(ReferenceUuid) > 0 Then
        header = header & "," & EtaJsonText("referenceUUID", ReferenceUuid)
        kind = "R"
    Else
        kind = "S"
    End If
    header = header & "," & EtaJsonText("currency", "EGP") & "," & EtaJsonNum("exchangeRate", 0)
    EtaReceiptJson = "{""header"":{" & header & "},""documentType"":{" & EtaJsonText("receiptType", kind) & "," & _
        EtaJsonText("typeVersion", TYPE_VERSION) & "}," & SellerJson & "," & BuyerJson & ",""itemData"":[" & ItemsJson & _
        "]," & EtaJsonNum("totalSales", TotalSales) & "," & EtaJsonNum("totalCommercialDiscount", TotalDiscount) & "," & _
        EtaJsonNum("totalItemsDiscount", 0) & "," & EtaJsonNum("netAmount", NetAmount) & "," & _
        EtaJsonNum("feesAmount", 0) & "," & EtaJsonNum("totalAmount", TotalAmount) & ",""taxTotals"":[{" & _
        EtaJsonText("taxType", "T1") & "," & EtaJsonNum("amount", TaxTotal) & "}]," & _
        EtaJsonText("paymentMethod", PaymentMethod) & "}"
End Function

Public Function EtaWithUuid(ByVal Json As String, ByVal Uuid As String) As String
    ' The receipt with its UUID in header.uuid (built with "").
    EtaWithUuid = Replace(Json, """uuid"":""""", """uuid"":""" & Uuid & """", 1, 1, vbBinaryCompare)
End Function

Public Function EtaPaymentCode(ByVal PaymentType As String, ByVal MethodID As Variant) As String
    ' C cash, V visa (bank card), O others (bank transfer, sale on account).
    If PaymentType = "CREDIT" Then
        EtaPaymentCode = "O"
    Else
        Select Case Nz(MethodID, 1)
            Case 2: EtaPaymentCode = "V"
            Case 3: EtaPaymentCode = "O"
            Case Else: EtaPaymentCode = "C"
        End Select
    End If
End Function

Public Function EtaQrLink(ByVal Env As String, ByVal Uuid As String, ByVal DateTimeIssued As String, _
                          ByVal TotalAmount As Variant, ByVal Rin As String) As String
    ' The QR code of the receipt: its page on the ETA portal.
    EtaQrLink = EtaPortalUrl(Env) & "/receipts/search/" & Uuid & "/share/" & DateTimeIssued & "#Total:" & _
                EtaNum(TotalAmount) & ",IssuerRIN:" & Rin
End Function

Public Function EtaUtcTimestamp(ByVal LocalTime As Date) As String
    ' "2026-10-09T12:30:00Z": the local time of the PC in UTC (Windows knows the daylight saving time of Egypt).
    Dim saved As Integer, t As Date, wmi As Object
    saved = Calendar
    Calendar = vbCalGreg
    On Error Resume Next
    Set wmi = CreateObject("WbemScripting.SWbemDateTime")
    wmi.SetVarDate LocalTime, True
    t = wmi.GetVarDate(False)
    If Err.Number <> 0 Then t = DateAdd("h", -EG_UTC_OFFSET_HOURS, LocalTime)
    On Error GoTo 0
    EtaUtcTimestamp = Format$(t, "yyyy-mm-dd") & "T" & Format$(t, "hh:nn:ss") & "Z"
    Calendar = saved
End Function

'==============================================================================
' The UUID: SHA-256 of the ETA serialization
'==============================================================================
Public Function EtaSerialize(ByVal Json As String) As String
    ' Every property is "NAME" (upper case) then its value; a simple value is "value"; an array is "NAME" then,
    ' per element, "NAME" and the element.
    Dim pos As Long
    pos = 1
    EtaSerialize = SerValue(Json, pos)
End Function

Public Function EtaReceiptUuid(ByVal Json As String) As String
    ' The UUID of a receipt whose header.uuid is "": hex SHA-256 of its serialization (UTF-8).
    EtaReceiptUuid = Sha256Text(EtaSerialize(Json))
End Function

Private Sub SkipBlank(ByRef s As String, ByRef pos As Long)
    Do While pos <= Len(s)
        Select Case Mid$(s, pos, 1)
            Case " ", vbTab, vbCr, vbLf: pos = pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Function SerValue(ByRef s As String, ByRef pos As Long) As String
    SkipBlank s, pos
    Select Case Mid$(s, pos, 1)
        Case "{": SerValue = SerObject(s, pos)
        Case """": SerValue = """" & ReadJsonString(s, pos) & """"
        Case Else: SerValue = """" & ReadScalar(s, pos) & """"
    End Select
End Function

Private Function SerObject(ByRef s As String, ByRef pos As Long) As String
    Dim out As String, key As String, ch As String
    pos = pos + 1                                        ' {
    SkipBlank s, pos
    If Mid$(s, pos, 1) = "}" Then
        pos = pos + 1
        Exit Function
    End If
    Do
        SkipBlank s, pos
        key = """" & UCase$(ReadJsonString(s, pos)) & """"
        SkipBlank s, pos
        pos = pos + 1                                    ' :
        SkipBlank s, pos
        If Mid$(s, pos, 1) = "[" Then
            out = out & key
            pos = pos + 1
            SkipBlank s, pos
            If Mid$(s, pos, 1) = "]" Then
                pos = pos + 1
            Else
                Do
                    out = out & key & SerValue(s, pos)
                    SkipBlank s, pos
                    ch = Mid$(s, pos, 1)
                    pos = pos + 1
                    If ch <> "," Then Exit Do
                Loop
            End If
        Else
            out = out & key & SerValue(s, pos)
        End If
        SkipBlank s, pos
        ch = Mid$(s, pos, 1)
        pos = pos + 1
        If ch <> "," Then Exit Do
    Loop
    SerObject = out
End Function

Private Function ReadJsonString(ByRef s As String, ByRef pos As Long) As String
    ' The string that starts at pos (its opening quote), unescaped; pos moves after its closing quote.
    Dim out As String, ch As String
    pos = pos + 1
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = """" Then
            pos = pos + 1
            Exit Do
        ElseIf ch = "\" Then
            ch = Mid$(s, pos + 1, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u"
                    out = out & ChrW(CLng("&H" & Mid$(s, pos + 2, 4)))
                    pos = pos + 4
                Case Else: out = out & ch
            End Select
            pos = pos + 2
        Else
            out = out & ch
            pos = pos + 1
        End If
    Loop
    ReadJsonString = out
End Function

Private Function ReadScalar(ByRef s As String, ByRef pos As Long) As String
    ' A number, true, false or null as written.
    Dim start As Long
    start = pos
    Do While pos <= Len(s)
        Select Case Mid$(s, pos, 1)
            Case ",", "}", "]", " ", vbTab, vbCr, vbLf: Exit Do
        End Select
        pos = pos + 1
    Loop
    ReadScalar = Mid$(s, start, pos - start)
End Function

'==============================================================================
' ETA replies
'==============================================================================
Public Function EtaFirstText(ByVal Json As String, ByVal Path1 As String, ByVal Path2 As String, _
                           ByVal Path3 As String) As String
    ' The first of these values that is not missing nor empty.
    Dim v As Variant, paths As Variant, i As Long
    paths = Array(Path1, Path2, Path3)
    For i = 0 To 2
        If Len(paths(i)) > 0 Then
            v = JsonGet(Json, paths(i))
            If Not IsNull(v) Then
                If Len(CStr(v)) > 0 Then
                    EtaFirstText = CStr(v)
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function EtaErrors(ByVal Json As String, ByVal Path As String) As String
    ' "message (path: message; ...)" of the ETA error object at Path (at most 3 details).
    Dim msg As String, d As String, i As Long, p As String
    msg = EtaFirstText(Json, Path & ".message", Path & ".code", "")
    For i = 0 To 2
        p = Path & ".details[" & i & "]"
        If IsNull(JsonGet(Json, p)) Then Exit For
        If Len(d) > 0 Then d = d & "; "
        d = d & EtaFirstText(Json, p & ".propertyPath", p & ".target", p & ".code") & ": " & EtaFirstText(Json, p & ".message", "", "")
    Next
    If Len(msg) > 0 And Len(d) > 0 Then
        EtaErrors = msg & " (" & d & ")"
    ElseIf Len(msg) > 0 Then
        EtaErrors = msg
    Else
        EtaErrors = d
    End If
End Function

Public Function EtaSubmitResult(ByVal HttpStatus As Long, ByVal Body As String, ByRef DocStatus As String, _
                                ByRef Message As String, ByRef SubmissionId As String) As String
    ' The result of a submission: OK (SUBMITTED), REJECTED, ERROR or NETWORK; "" status = it stays PENDING.
    DocStatus = ""
    Message = ""
    SubmissionId = ""
    If HttpStatus = 0 Then
        EtaSubmitResult = "NETWORK"
        Message = "تعذّر الاتصال بالمنظومة."
    ElseIf HttpStatus = 200 Or HttpStatus = 202 Then
        If Not IsNull(JsonGet(Body, "rejectedDocuments[0]")) Then
            EtaSubmitResult = "REJECTED"
            DocStatus = "REJECTED"
            Message = EtaErrors(Body, "rejectedDocuments[0].error")
            If Len(Message) = 0 Then Message = "رُفض الإيصال."
        ElseIf Not IsNull(JsonGet(Body, "acceptedDocuments[0]")) Then
            EtaSubmitResult = "OK"
            DocStatus = "SUBMITTED"
            SubmissionId = Nz(JsonGet(Body, "submissionId"), "")
        Else
            EtaSubmitResult = "ERROR"
            Message = "رد غير متوقع من المصلحة: " & Left$(Body, 200)
        End If
    ElseIf HttpStatus = 400 Then
        EtaSubmitResult = "REJECTED"
        DocStatus = "REJECTED"
        Message = EtaErrors(Body, "error")
        If Len(Message) = 0 Then Message = Left$(Body, 200)
    ElseIf HttpStatus = 401 Or HttpStatus = 403 Then
        EtaSubmitResult = "ERROR"
        Message = "رفضت المصلحة بيانات الدخول (بيانات جهاز نقطة البيع)."
    ElseIf HttpStatus = 429 Or HttpStatus >= 500 Then
        EtaSubmitResult = "NETWORK"
        Message = "المنظومة غير متاحة الآن (رمز " & HttpStatus & ")."
    Else
        EtaSubmitResult = "ERROR"
        Message = "رد غير متوقع من المصلحة (رمز " & HttpStatus & "): " & Left$(Body, 200)
    End If
End Function

Public Function EtaStatusResult(ByVal Body As String, ByRef Message As String) As String
    ' VALID, INVALID (Message: the errors) or "" (still in progress) from the details of a submission.
    Dim st As String, d As String, i As Long, p As String
    Message = ""
    st = EtaFirstText(Body, "receipts[0].status", "status", "")
    If LCase$(st) = "valid" Then
        EtaStatusResult = "VALID"
    ElseIf LCase$(st) = "invalid" Then
        EtaStatusResult = "INVALID"
        For i = 0 To 2
            p = "receipts[0].errors[" & i & "]"
            If IsNull(JsonGet(Body, p)) Then Exit For
            If Len(d) > 0 Then d = d & "; "
            d = d & EtaFirstText(Body, p & ".propertyPath", p & ".target", p & ".code") & ": " & EtaFirstText(Body, p & ".message", "", "")
        Next
        Message = d
    End If
End Function

'==============================================================================
' The receipt of a document
'==============================================================================
Private Function Txt(ByVal v As Variant) As String
    Txt = Trim$(Nz(v, ""))
End Function

Private Function CurrentEtaEnv() As String
    CurrentEtaEnv = EtaEnv(Nz(SettingValue("EInvoiceEnvironment"), "TEST"))
End Function

Private Function LinesOf(ByVal DocKind As String) As String
    If DocKind = "RETURN" Then
        LinesOf = "SalesReturnDetails"
    Else
        LinesOf = "SalesInvoiceDetails"
    End If
End Function

Public Function EtaSellerProblem() As String
    ' The lines of the missing data of the seller and of the device (POS), "" when complete.
    Dim msg As String
    msg = EtaIssuerProblem()
    If Len(Txt(SettingValue("EtaPosSerial"))) = 0 Then msg = msg & "- الرقم التسلسلي لجهاز نقطة البيع" & vbCrLf
    EtaSellerProblem = msg
End Function

Public Function EtaIssuerProblem() As String
    ' The lines of the missing data of the seller (receipts and e-invoices), "" when complete.
    Dim s As DAO.Recordset, msg As String
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    If Len(Txt(s!VATNumber)) = 0 Or Len(TaxNumberProblem(s!VATNumber, "EG")) > 0 Then
        msg = msg & "- رقم التسجيل الضريبي للمحل (9 أرقام)" & vbCrLf
    End If
    If Len(Txt(s!StoreName)) = 0 Then msg = msg & "- اسم المحل" & vbCrLf
    If Len(Txt(s!EtaGovernate)) = 0 Then msg = msg & "- المحافظة" & vbCrLf
    If Len(Txt(s!City)) = 0 Then msg = msg & "- مدينة المحل" & vbCrLf
    If Len(Txt(s!StreetName)) = 0 Then msg = msg & "- شارع المحل" & vbCrLf
    If Len(Txt(s!BuildingNo)) = 0 Then msg = msg & "- رقم مبنى المحل" & vbCrLf
    If Not (Txt(s!EtaActivityCode) Like "####") Then msg = msg & "- كود النشاط (4 أرقام)" & vbCrLf
    s.Close
    EtaIssuerProblem = msg
End Function

Public Function EtaDataProblem(ByVal DocKind As String, ByVal DocID As Long) As String
    ' "" when the receipt can be built: seller data, the ETA code of every product, the sale of a return sent
    ' first, the ID of a person who buys for 150,000 or more.
    Dim msg As String, rs As DAO.Recordset, names As String, n As Long
    msg = EtaSellerProblem()
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
            msg = msg & "- أرسل إيصال البيع الأصلي أولًا" & vbCrLf
        End If
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT h.TotalAmount, c.VATNumber, c.NationalID, c.CustomerID FROM " & DocTableOf(DocKind) & _
                                     " AS h LEFT JOIN Customers AS c ON h.CustomerID = c.CustomerID WHERE h." & _
                                     DocKeyOf(DocKind) & " = " & DocID, dbOpenSnapshot)
    If Not rs.EOF Then
        If Nz(rs!TotalAmount, 0) >= BUYER_ID_LIMIT And Len(Txt(rs!VATNumber)) = 0 And Len(Txt(rs!NationalID)) = 0 Then
            msg = msg & "- إيصال 150 ألف جنيه أو أكثر: الرقم القومي أو الضريبي للمشتري" & vbCrLf
        End If
    End If
    rs.Close
    If Len(msg) > 0 Then EtaDataProblem = "بيانات ناقصة للإيصال الإلكتروني:" & vbCrLf & msg
End Function

Public Function EtaDocumentJson(ByVal DocKind As String, ByVal DocID As Long, ByVal PreviousUuid As String, _
                                ByRef DateTimeIssued As String, ByRef TotalAmount As Currency) As String
    ' The receipt of a sales invoice or return, header.uuid "" (EtaWithUuid puts it in).
    Dim db As DAO.Database, h As DAO.Recordset, s As DAO.Recordset, c As DAO.Recordset, rs As DAO.Recordset
    Dim isReturn As Boolean, seller As String, buyer As String, items As String, refUuid As String
    Dim totalSales As Currency, totalDisc As Currency, buyerType As String, buyerId As String, buyerName As String
    Dim mobile As String, defaultCustomer As Long, branch As String
    Set db = CurrentDb
    isReturn = (DocKind = "RETURN")
    If isReturn Then
        Set h = db.OpenRecordset("SELECT r.*, r.ReturnNumber AS DocNumber, r.ReturnDate AS DocDate, r.RefundType AS PayType, " & _
                                 "o.EtaUUID AS OriginalUUID FROM SalesReturns AS r INNER JOIN SalesInvoices AS o " & _
                                 "ON r.SalesInvoiceID = o.SalesInvoiceID WHERE r.SalesReturnID = " & DocID, dbOpenSnapshot)
        refUuid = Txt(h!OriginalUUID)
    Else
        Set h = db.OpenRecordset("SELECT h.*, h.InvoiceNumber AS DocNumber, h.InvoiceDate AS DocDate, h.PaymentType AS PayType " & _
                                 "FROM SalesInvoices AS h WHERE h.SalesInvoiceID = " & DocID, dbOpenSnapshot)
    End If
    DateTimeIssued = EtaUtcTimestamp(h!DocDate)
    TotalAmount = h!TotalAmount
    Set s = db.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    branch = Txt(s!EtaBranchCode)
    If Len(branch) = 0 Then branch = "0"                 ' the main branch
    seller = EtaSellerJson(Txt(s!VATNumber), Txt(s!StoreName), branch, Txt(s!EtaGovernate), Txt(s!City), _
                           Txt(s!StreetName), Txt(s!BuildingNo), Txt(s!PostalCode), Txt(s!EtaPosSerial), _
                           Txt(s!EtaActivityCode))
    defaultCustomer = Nz(s!DefaultCustomerID, 0)
    buyerType = "P"
    Set c = db.OpenRecordset("SELECT * FROM Customers WHERE CustomerID = " & Nz(h!CustomerID, 0), dbOpenSnapshot)
    If Not c.EOF Then
        If c!CustomerID <> defaultCustomer And Not Nz(c!IsSystem, False) Then
            If Len(Txt(c!VATNumber)) > 0 Then
                buyerType = "B"
                buyerId = Txt(c!VATNumber)
            Else
                buyerId = Txt(c!NationalID)
            End If
            buyerName = Txt(c!CustomerName)
            mobile = Txt(c!Mobile)
        End If
    End If
    c.Close
    buyer = EtaBuyerJson(buyerType, buyerId, buyerName, mobile)
    Set rs = db.OpenRecordset("SELECT d.Quantity, d.UnitPrice, d.Discount, d.NetAmount, d.VATRate, d.Tax, d.LineTotal, " & _
                              "p.ProductCode, p.ProductName, p.EtaItemType, p.EtaItemCode, u.EtaUnitCode FROM (" & LinesOf(DocKind) & _
                              " AS d INNER JOIN Products AS p ON d.ProductID = p.ProductID) INNER JOIN Units AS u ON " & _
                              "p.UnitID = u.UnitID WHERE d." & DocKeyOf(DocKind) & " = " & DocID & " ORDER BY d." & _
                              IIf(isReturn, "ReturnDetailID", "SalesDetailID"), dbOpenSnapshot)
    Do Until rs.EOF
        If Len(items) > 0 Then items = items & ","
        items = items & EtaItemJson(Txt(rs!ProductCode), Txt(rs!ProductName), Nz(rs!EtaItemType, "EGS"), Txt(rs!EtaItemCode), _
                                    Nz(rs!EtaUnitCode, "EA"), rs!Quantity, rs!UnitPrice, rs!NetAmount + Nz(rs!Discount, 0), _
                                    Nz(rs!Discount, 0), rs!NetAmount, rs!Tax, CDec(rs!VATRate) * 100, rs!LineTotal)
        totalSales = totalSales + rs!NetAmount + Nz(rs!Discount, 0)
        totalDisc = totalDisc + Nz(rs!Discount, 0)
        rs.MoveNext
    Loop
    rs.Close
    EtaDocumentJson = EtaReceiptJson(DateTimeIssued, Txt(h!DocNumber), "", PreviousUuid, refUuid, seller, buyer, items, _
                                     totalSales, totalDisc, h!TaxableAmount, h!Tax, h!TotalAmount, _
                                     EtaPaymentCode(Txt(h!PayType), h!PaymentMethodID))
    s.Close
    h.Close
End Function

Public Function EtaPrepareDocument(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Builds the receipt once, in the order of the chain: its previousUUID is the last UUID of the device, then
    ' its own UUID becomes the last one. A rejected receipt is built again (corrected data) on the same previous
    ' UUID. "" or the problem.
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, prev As String, json As String, uuid As String
    Dim rs As DAO.Recordset, dt As String, total As Currency, oldUuid As String, rebuild As Boolean
    On Error GoTo EH
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT EtaUUID, EtaPreviousUUID, ZatcaStatus FROM " & DocTableOf(DocKind) & " WHERE " & _
                              DocKeyOf(DocKind) & " = " & DocID, dbOpenSnapshot)
    oldUuid = Txt(rs!EtaUUID)
    rebuild = (Len(oldUuid) > 0 And Txt(rs!ZatcaStatus) = "REJECTED")
    prev = Txt(rs!EtaPreviousUUID)
    rs.Close
    If Len(oldUuid) > 0 And Not rebuild Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "UPDATE Settings SET EtaLastUUID = EtaLastUUID WHERE SettingID = 1", dbFailOnError    ' lock the chain
    If Not rebuild Then prev = Nz(DbValue("SELECT EtaLastUUID FROM Settings WHERE SettingID = 1"), "")
    json = EtaDocumentJson(DocKind, DocID, prev, dt, total)
    uuid = EtaReceiptUuid(json)
    json = EtaWithUuid(json, uuid)
    Set rs = db.OpenRecordset("SELECT EtaUUID, EtaPreviousUUID, QRCodeData, EInvoiceXml FROM " & DocTableOf(DocKind) & _
                              " WHERE " & DocKeyOf(DocKind) & " = " & DocID, dbOpenDynaset)
    rs.Edit
    rs!EtaUUID = uuid
    rs!EtaPreviousUUID = IIf(Len(prev) = 0, Null, prev)
    rs!QRCodeData = EtaQrLink(CurrentEtaEnv(), uuid, dt, total, Nz(SettingValue("VATNumber"), ""))
    rs!EInvoiceXml = json                                 ' a memo: the receipt as it is sent
    rs.Update
    rs.Close
    If rebuild Then
        db.Execute "UPDATE Settings SET EtaLastUUID = '" & uuid & "' WHERE SettingID = 1 AND EtaLastUUID = '" & _
                   oldUuid & "'", dbFailOnError
    Else
        db.Execute "UPDATE Settings SET EtaLastUUID = '" & uuid & "' WHERE SettingID = 1", dbFailOnError
    End If
    ws.CommitTrans
    Exit Function
EH:
    EtaPrepareDocument = "تعذّر بناء الإيصال: " & Err.Description
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' Requests
'==============================================================================
Public Function EtaUrlEncode(ByVal Value As String) As String
    ' application/x-www-form-urlencoded: UTF-8, unreserved characters kept.
    Dim b() As Byte, n As Long, i As Long, out As String, c As Long
    n = Utf8Bytes(Value, b)
    For i = 0 To n - 1
        c = b(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            out = out & Chr$(c)
        Else
            out = out & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next
    EtaUrlEncode = out
End Function

Public Function EtaMaskToken(ByVal Reply As String) As String
    ' A token reply as it is kept in the log: the access token is hidden.
    Dim tok As Variant
    tok = JsonGet(Reply, "access_token")
    EtaMaskToken = Reply
    If Not IsNull(tok) Then
        If Len(tok) > 0 Then EtaMaskToken = Replace(Reply, CStr(tok), "***")
    End If
End Function

Public Function EtaToken(ByRef Token As String) As String
    ' The access token of the POS (client credentials + the device headers). "" or "RESULT|message".
    Dim s As DAO.Recordset, headers As String, body As String, status As Long, reply As String, msg As String
    Dim started As Single, ms As Long, result As String, url As String, life As Long
    If Len(m_token) > 0 And m_tokenEnv = CurrentEtaEnv() And Now < m_tokenExpires Then
        Token = m_token
        Exit Function
    End If
    Set s = CurrentDb.OpenRecordset("SELECT EtaClientId, EtaClientSecret, EtaPosSerial, EtaPosOsVersion, " & _
                                    "EtaPreSharedKey FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    headers = "Content-Type: application/x-www-form-urlencoded" & vbLf & "Accept: application/json" & vbLf & _
              "posserial: " & Txt(s!EtaPosSerial) & vbLf & "pososversion: " & Nz(Txt(s!EtaPosOsVersion), "Windows")
    If Len(Txt(s!EtaPreSharedKey)) > 0 Then headers = headers & vbLf & "presharedkey: " & Txt(s!EtaPreSharedKey)
    body = "grant_type=client_credentials&client_id=" & EtaUrlEncode(Txt(s!EtaClientId)) & "&client_secret=" & _
           EtaUrlEncode(Txt(s!EtaClientSecret))
    s.Close
    url = EtaTokenUrl(CurrentEtaEnv())
    started = Timer
    msg = HttpSend("POST", url, headers, body, status, reply, 30)
    ms = CLng((Timer - started) * 1000)
    If status = 200 And Not IsNull(JsonGet(reply, "access_token")) Then
        result = "OK"
        m_token = JsonGet(reply, "access_token")
        m_tokenEnv = CurrentEtaEnv()
        life = Val(Nz(JsonGet(reply, "expires_in"), "3600"))
        m_tokenExpires = DateAdd("s", life - 60, Now)
        Token = m_token
    ElseIf status = 0 Or status = 429 Or status >= 500 Then
        result = "NETWORK"
        If Len(msg) = 0 Then msg = "المنظومة غير متاحة الآن (رمز " & status & ")."
    Else
        result = "ERROR"
        msg = "رفضت المصلحة بيانات جهاز نقطة البيع (رمز " & status & "): " & _
              EtaFirstText(reply, "error_description", "error", "")
    End If
    ' the request body carries the client secret: it is not logged
    LogEInvoice "", 0, "TOKEN", url, status, result, Left$(msg, 255), "", EtaMaskToken(reply), ms
    If result <> "OK" Then EtaToken = result & "|" & msg
End Function

Public Function EtaReadyProblem() As String
    ' "" when Egypt can send: the data of the seller and the credentials of the POS (receipts) or of the program
    ' (e-invoices, modEtaInvoice); each sending checks its own credentials.
    Dim msg As String, pos As Boolean, erp As Boolean
    If AppCountry() <> "EG" Then
        EtaReadyProblem = "الإيصال الإلكتروني المصري لدولة التشغيل مصر فقط."
        Exit Function
    End If
    pos = Len(Txt(SettingValue("EtaClientId"))) > 0 And Len(Txt(SettingValue("EtaClientSecret"))) > 0
    erp = Len(Txt(SettingValue("EtaErpClientId"))) > 0 And Len(Txt(SettingValue("EtaErpClientSecret"))) > 0
    If Not pos And Not erp Then
        msg = "- بيانات دخول جهاز نقطة البيع (Client ID و Client Secret)" & vbCrLf
    End If
    msg = msg & EtaIssuerProblem()
    If Len(msg) > 0 Then EtaReadyProblem = "إعداد الإيصال الإلكتروني ناقص:" & vbCrLf & msg
End Function

Public Function EtaSendDocument(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Builds the receipt if it is not yet, then submits it. Returns "RESULT|message" for modEInvoice. A tax
    ' invoice (customer with a tax number) and its returns are e-invoices: modEtaInvoice.
    Dim msg As String, json As String, token As String, body As String, status As Long, reply As String
    Dim ms As Long, started As Single, result As String, docStatus As String, subId As String, url As String
    If EtaIsInvoice(DocKind, DocID) Then
        EtaSendDocument = EtaInvoiceSend(DocKind, DocID)
        Exit Function
    End If
    msg = EtaDataProblem(DocKind, DocID)
    If Len(msg) = 0 Then msg = EtaPrepareDocument(DocKind, DocID)
    If Len(msg) > 0 Then
        EtaSendDocument = "ERROR|" & msg
        Exit Function
    End If
    msg = EtaToken(token)
    If Len(msg) > 0 Then
        EtaSendDocument = msg
        Exit Function
    End If
    json = Nz(DbValue("SELECT EInvoiceXml FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "")
    body = "{""receipts"":[" & json & "]}"
    url = EtaApiUrl(CurrentEtaEnv()) & "/api/v1/receiptsubmissions"
    started = Timer
    msg = HttpSend("POST", url, "Content-Type: application/json" & vbLf & "Accept: application/json" & vbLf & _
                   "Authorization: Bearer " & token, body, status, reply, 60)
    ms = CLng((Timer - started) * 1000)
    result = EtaSubmitResult(status, reply, docStatus, msg, subId)
    If status = 401 Then m_token = ""                    ' expired: a new token next time
    LogEInvoice DocKind, DocID, "SUBMIT", url, status, result, Left$(msg, 255), body, reply, ms
    If Len(docStatus) > 0 Then SetEInvoiceStatus DocKind, DocID, docStatus, reply
    If Len(subId) > 0 Then
        CurrentDb.Execute "UPDATE " & DocTableOf(DocKind) & " SET EtaSubmissionId = " & SqlText(subId) & " WHERE " & _
                          DocKeyOf(DocKind) & " = " & DocID, dbFailOnError
    End If
    EtaSendDocument = result & "|" & msg
End Function

Public Function EtaRefreshStatus(ByVal DocKind As String, ByVal DocID As Long) As String
    ' The result of a submitted receipt (VALID / INVALID). Returns "RESULT|message"; still in progress = OK.
    Dim subId As String, token As String, msg As String, url As String, status As Long, reply As String
    Dim ms As Long, started As Single, st As String, result As String
    If EtaIsInvoice(DocKind, DocID) Then
        EtaRefreshStatus = EtaInvoiceRefresh(DocKind, DocID)
        Exit Function
    End If
    subId = Nz(DbValue("SELECT EtaSubmissionId FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), "")
    If Len(subId) = 0 Then
        EtaRefreshStatus = "ERROR|" & "لم يُرسل الإيصال بعد."
        Exit Function
    End If
    msg = EtaToken(token)
    If Len(msg) > 0 Then
        EtaRefreshStatus = msg
        Exit Function
    End If
    url = EtaApiUrl(CurrentEtaEnv()) & "/api/v1/receiptsubmissions/" & subId & "/details?PageNo=1&PageSize=10"
    started = Timer
    msg = HttpSend("GET", url, "Accept: application/json" & vbLf & "Authorization: Bearer " & token, "", status, reply, 30)
    ms = CLng((Timer - started) * 1000)
    If status = 200 Then
        st = EtaStatusResult(reply, msg)
        If st = "INVALID" Then
            result = "REJECTED"
        Else
            result = "OK"
        End If
        If Len(st) > 0 Then SetEInvoiceStatus DocKind, DocID, st, reply
    ElseIf status = 0 Or status = 429 Or status >= 500 Then
        result = "NETWORK"
        If Len(msg) = 0 Then msg = "المنظومة غير متاحة الآن (رمز " & status & ")."
    Else
        If status = 401 Then m_token = ""
        result = "ERROR"
        msg = "رد غير متوقع من المصلحة (رمز " & status & "): " & Left$(reply, 200)
    End If
    LogEInvoice DocKind, DocID, "STATUS", url, status, result, Left$(msg, 255), "", reply, ms
    EtaRefreshStatus = result & "|" & msg
End Function

'==============================================================================
' frmEtaSetup
'==============================================================================
Private Function SetupControls() As Variant
    SetupControls = Array("txtClientId", "txtClientSecret", "txtPosSerial", "txtPosOs", "txtPreSharedKey", "txtBranchCode", _
                          "txtActivityCode", "txtGovernate", "txtErpClientId", "txtErpClientSecret", "txtSignerPath", _
                          "txtTokenPin", "txtSignerArgs")
End Function

Private Function SetupFields() As Variant
    ' the Settings field of each control of SetupControls
    SetupFields = Array("EtaClientId", "EtaClientSecret", "EtaPosSerial", "EtaPosOsVersion", "EtaPreSharedKey", "EtaBranchCode", _
                        "EtaActivityCode", "EtaGovernate", "EtaErpClientId", "EtaErpClientSecret", "EtaSignerPath", _
                        "EtaTokenPin", "EtaSignerArgs")
End Function

Public Sub EtaSetupLoad(ByVal frm As Access.Form)
    Dim s As DAO.Recordset, ctl As Variant, fld As Variant, i As Long
    ctl = SetupControls()
    fld = SetupFields()
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    For i = 0 To UBound(ctl)
        frm(ctl(i)).Value = s(fld(i)).Value
    Next
    s.Close
    EtaSetupShow frm
End Sub

Public Sub EtaSetupShow(ByVal frm As Access.Form)
    Dim msg As String
    msg = EtaReadyProblem()
    If Len(msg) = 0 Then
        frm!lblReady.Caption = Tr("الإعداد مكتمل. فعّل الإرسال من شاشة الفاتورة الإلكترونية.")
        frm!lblReady.ForeColor = CLR_PRIMARY
    Else
        frm!lblReady.Caption = Tr(msg)
        frm!lblReady.ForeColor = CLR_DANGER
    End If
End Sub

Public Sub EtaSetupSave(ByVal frm As Access.Form)
    Dim before As Collection, rs As DAO.Recordset, ctl As Variant, fld As Variant, i As Long
    If Not HasPermission("SETTINGS") Then
        ShowWarning "إعداد الربط يحتاج صلاحية إعدادات المحل."
        Exit Sub
    End If
    ctl = SetupControls()
    fld = SetupFields()
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenDynaset)
    rs.Edit
    For i = 0 To UBound(ctl)
        If Len(Trim$(Nz(frm(ctl(i)).Value, ""))) = 0 Then
            rs(fld(i)) = Null
        Else
            rs(fld(i)) = Trim$(Nz(frm(ctl(i)).Value, ""))
        End If
    Next
    rs.Update
    rs.Close
    AuditEdited "EDIT", "Settings", "SettingID", 1, before
    m_token = ""                                          ' the next requests log in with the new data
    EtaErpTokenReset
    EtaSetupShow frm
    ShowInfo "تم حفظ إعداد الربط."
End Sub

Public Sub EtaSetupTestLogin(ByVal frm As Access.Form)
    ' Logs in with the saved credentials of the POS and of the program (those that are set): nothing is sent.
    Dim msg As String, token As String, reason As String, report As String, tried As Boolean
    DoCmd.Hourglass True
    If Len(Txt(SettingValue("EtaClientId"))) > 0 Then
        tried = True
        m_token = ""
        msg = EtaToken(token)
        If Len(msg) = 0 Then
            report = report & Tr("جهاز نقطة البيع (الإيصال): تم الدخول.") & vbCrLf
        Else
            SplitResult msg, reason
            report = report & Tr("جهاز نقطة البيع (الإيصال): ") & reason & vbCrLf
        End If
    End If
    If Len(Txt(SettingValue("EtaErpClientId"))) > 0 Then
        tried = True
        EtaErpTokenReset
        msg = EtaErpToken(token)
        If Len(msg) = 0 Then
            report = report & Tr("البرنامج (الفاتورة الإلكترونية): تم الدخول.") & vbCrLf
        Else
            SplitResult msg, reason
            report = report & Tr("البرنامج (الفاتورة الإلكترونية): ") & reason & vbCrLf
        End If
    End If
    DoCmd.Hourglass False
    If Not tried Then
        ShowWarning "اكتب بيانات الدخول واحفظها أولًا."
    Else
        ShowInfo EnvironmentName() & vbCrLf & report
    End If
End Sub

'==============================================================================
' In-Access test (nothing is sent; changes rolled back)
'==============================================================================
Public Function TestEtaReceipt() As Boolean
    Dim passed As Long, failed As Long, report As String, st As String, msg As String, subId As String
    Dim ws As DAO.Workspace, inTrans As Boolean, json As String, item As String
    g_SilentMode = True
    Debug.Print "=== TestEtaReceipt  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record EtaNum(12.5) = "12.5" And EtaNum(100) = "100" And EtaNum(CCur(0.0125)) = "0.0125" And EtaNum(0) = "0", _
           "صيغة الأرقام", passed, failed, report
    item = EtaItemJson("P1", "Tea", "EGS", "EG-123456789-1", "EA", 2, 10, 20, 0, 20, 2.8, 14, 22.8)
    json = EtaReceiptJson("2026-10-09T10:00:00Z", "INV-1", "", "", "", _
                          EtaSellerJson("123456789", "Store", "0", "Cairo", "Nasr City", "Abbas", "12", "", "POS-1", "4711"), _
                          EtaBuyerJson("P", "", "", ""), item, 20, 0, 20, 2.8, 22.8, "C")
    Record EtaReceiptUuid(json) = TEST_RECEIPT_UUID, "معرّف الإيصال (UUID) يساوي النسخة المكتوبة بـ Python", _
           passed, failed, report
    Record InStr(EtaWithUuid(json, "abc"), """uuid"":""abc"",""previousUUID"":""""") > 0, "وضع المعرّف في الإيصال", _
           passed, failed, report
    Record EtaSubmitResult(202, "{""submissionId"": ""S1"", ""acceptedDocuments"": [{""uuid"": ""u""}], " & _
                           """rejectedDocuments"": []}", st, msg, subId) = "OK" And st = "SUBMITTED" And subId = "S1", _
           "قبول الإيصال", passed, failed, report
    Record EtaSubmitResult(503, "", st, msg, subId) = "NETWORK" And st = "", "المنظومة غير متاحة: يبقى بانتظار الإرسال", _
           passed, failed, report
    Record EtaStatusResult("{""receipts"": [{""status"": ""Valid""}]}", msg) = "VALID", "قراءة حالة الإيصال", _
           passed, failed, report
    Record Right$(EtaUtcTimestamp(Now), 1) = "Z" And Len(EtaUtcTimestamp(Now)) = 20, "الوقت بتوقيت UTC", _
           passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "UPDATE Settings SET EtaClientId = Null WHERE SettingID = 1", dbFailOnError
    Record Len(EtaReadyProblem()) > 0, "لا إرسال بدون بيانات الدخول", passed, failed, report
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
        TestMsg "جميع اختبارات الإيصال الإلكتروني ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestEtaReceipt"
        TestEtaReceipt = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestEtaReceipt"
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
