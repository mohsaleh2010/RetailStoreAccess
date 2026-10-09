Attribute VB_Name = "modZatcaXml"
'==============================================================================
' modZatcaXml  -  Retail Store Management System (ZATCA phase 2 document)
'
' docs/46-ZATCA-Invoice-XML.md. The UBL 2.1 file of a sales invoice (388) or credit note (381):
'   ZatcaDocumentXml    the document from the data, written directly in canonical form (C14N)
'   ZatcaInvoiceHash    SHA-256 of the document without the declaration, ext:UBLExtensions,
'                       cac:Signature and the QR reference (base64)
'   ZatcaCertInfo       issuer, serial number, public key and signature of the certificate (DER)
'   ZatcaExtensions     the XAdES signature block (texts of modZatcaData, spaces included)
'   ZatcaQRCode         the QR code with nine tags
'   ZatcaSignHash       ECDSA secp256k1 signature of the hash bytes by OpenSSL (dgst -sha256 -sign)
'   ZatcaPrepareDocument  signs a document and moves the hash chain (PIH) forward (used by phase C2)
' tools/zatca_xml_reference.py is the Python mirror (checked against lxml C14N, OpenSSL and cryptography).
'==============================================================================
Option Compare Database
Option Explicit

Private Const B64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const SAMPLE_HASH As String = "psYdv0X6uMep5fDv8r4wYBR9M0gOsKOxjyiVw/CL0o0="   ' of SampleXml (Python mirror)

'==============================================================================
' Values, written as C14N writes them
'==============================================================================
Public Function ZxText(ByVal s As String) As String
    ZxText = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

Public Function ZxAttr(ByVal s As String) As String
    ZxAttr = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), """", "&quot;")
End Function

Public Function ZxPrice(ByVal Value As Currency) As String
    ' 2 to 4 decimals: the stored net unit price has 4.
    Dim whole As Currency, frac As String
    whole = Fix(Value)
    frac = Right$("0000" & CStr(CLng((Value - whole) * 10000)), 4)
    Do While Len(frac) > 2 And Right$(frac, 1) = "0"
        frac = Left$(frac, Len(frac) - 1)
    Loop
    ZxPrice = Format$(whole, "0") & "." & frac
End Function

Public Function ZxQuantity(ByVal Value As Currency) As String
    Dim whole As Currency
    whole = Fix(Value)
    ZxQuantity = Format$(whole, "0") & "." & Right$("0000" & CStr(CLng((Value - whole) * 10000)), 4) & "00"
End Function

Public Function ZxPercent(ByVal Rate As Double) As String
    ' 0.15 -> "15.00"
    ZxPercent = ZatcaAmount(CCur(Rate * 100))
End Function

'==============================================================================
' The document
'==============================================================================
Public Function ZxParty(ByVal Role As String, ByVal Name As String, ByVal Vat As String, ByVal Crn As String, _
                        ByVal Street As String, ByVal Building As String, ByVal Plot As String, _
                        ByVal District As String, ByVal City As String, ByVal Postal As String, _
                        ByVal Country As String) As String
    ' Seller (Role "Supplier") or buyer ("Customer"). No name = an empty element (simplified invoice).
    Dim s As String
    If Len(Name) = 0 Then
        ZxParty = "  <cac:Accounting" & Role & "Party></cac:Accounting" & Role & "Party>"
        Exit Function
    End If
    s = "  <cac:Accounting" & Role & "Party>" & vbLf & "    <cac:Party>"
    If Len(Crn) > 0 Then
        s = s & vbLf & "      <cac:PartyIdentification>" & vbLf & "        <cbc:ID schemeID=""CRN"">" & ZxText(Crn) & _
            "</cbc:ID>" & vbLf & "      </cac:PartyIdentification>"
    End If
    s = s & vbLf & "      <cac:PostalAddress>" & vbLf & "        <cbc:StreetName>" & ZxText(Street) & "</cbc:StreetName>" & _
        vbLf & "        <cbc:BuildingNumber>" & ZxText(Building) & "</cbc:BuildingNumber>"
    If Len(Plot) > 0 Then s = s & vbLf & "        <cbc:PlotIdentification>" & ZxText(Plot) & "</cbc:PlotIdentification>"
    If Len(District) > 0 Then
        s = s & vbLf & "        <cbc:CitySubdivisionName>" & ZxText(District) & "</cbc:CitySubdivisionName>"
    End If
    If Len(Country) = 0 Then Country = "SA"
    s = s & vbLf & "        <cbc:CityName>" & ZxText(City) & "</cbc:CityName>" & vbLf & "        <cbc:PostalZone>" & _
        ZxText(Postal) & "</cbc:PostalZone>" & vbLf & "        <cac:Country>" & vbLf & "          <cbc:IdentificationCode>" & _
        ZxText(Country) & "</cbc:IdentificationCode>" & vbLf & "        </cac:Country>" & vbLf & "      </cac:PostalAddress>"
    If Len(Vat) > 0 Then
        s = s & vbLf & "      <cac:PartyTaxScheme>" & vbLf & "        <cbc:CompanyID>" & ZxText(Vat) & "</cbc:CompanyID>" & _
            vbLf & "        <cac:TaxScheme>" & vbLf & "          <cbc:ID>VAT</cbc:ID>" & vbLf & "        </cac:TaxScheme>" & _
            vbLf & "      </cac:PartyTaxScheme>"
    End If
    s = s & vbLf & "      <cac:PartyLegalEntity>" & vbLf & "        <cbc:RegistrationName>" & ZxText(Name) & _
        "</cbc:RegistrationName>" & vbLf & "      </cac:PartyLegalEntity>" & vbLf & "    </cac:Party>" & vbLf & _
        "  </cac:Accounting" & Role & "Party>"
    ZxParty = s
End Function

Public Function ZxLine(ByVal Id As Long, ByVal Qty As Currency, ByVal UnitCode As String, ByVal Net As Currency, _
                       ByVal Discount As Currency, ByVal Tax As Currency, ByVal Total As Currency, _
                       ByVal ItemName As String, ByVal Category As String, ByVal Rate As Double, _
                       ByVal UnitPrice As Currency, ByVal Cur As String) As String
    Dim s As String, c As String
    c = ZxAttr(Cur)
    s = "  <cac:InvoiceLine>" & vbLf & "    <cbc:ID>" & Id & "</cbc:ID>" & vbLf & _
        "    <cbc:InvoicedQuantity unitCode=""" & ZxAttr(UnitCode) & """>" & ZxQuantity(Qty) & "</cbc:InvoicedQuantity>" & _
        vbLf & "    <cbc:LineExtensionAmount currencyID=""" & c & """>" & ZatcaAmount(Net) & "</cbc:LineExtensionAmount>"
    If Discount > 0 Then
        s = s & vbLf & "    <cac:AllowanceCharge>" & vbLf & "      <cbc:ChargeIndicator>false</cbc:ChargeIndicator>" & vbLf & _
            "      <cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>" & vbLf & _
            "      <cbc:Amount currencyID=""" & c & """>" & ZatcaAmount(Discount) & "</cbc:Amount>" & vbLf & _
            "    </cac:AllowanceCharge>"
    End If
    s = s & vbLf & "    <cac:TaxTotal>" & vbLf & "      <cbc:TaxAmount currencyID=""" & c & """>" & ZatcaAmount(Tax) & _
        "</cbc:TaxAmount>" & vbLf & "      <cbc:RoundingAmount currencyID=""" & c & """>" & ZatcaAmount(Total) & _
        "</cbc:RoundingAmount>" & vbLf & "    </cac:TaxTotal>" & vbLf & "    <cac:Item>" & vbLf & "      <cbc:Name>" & _
        ZxText(ItemName) & "</cbc:Name>" & vbLf & "      <cac:ClassifiedTaxCategory>" & vbLf & "        <cbc:ID>" & _
        ZxText(Category) & "</cbc:ID>" & vbLf & "        <cbc:Percent>" & ZxPercent(Rate) & "</cbc:Percent>" & vbLf & _
        "        <cac:TaxScheme>" & vbLf & "          <cbc:ID>VAT</cbc:ID>" & vbLf & "        </cac:TaxScheme>" & vbLf & _
        "      </cac:ClassifiedTaxCategory>" & vbLf & "    </cac:Item>" & vbLf & "    <cac:Price>" & vbLf & _
        "      <cbc:PriceAmount currencyID=""" & c & """>" & ZxPrice(UnitPrice) & "</cbc:PriceAmount>" & vbLf & _
        "    </cac:Price>" & vbLf & "  </cac:InvoiceLine>"
    ZxLine = s
End Function

Public Function ZxTaxSubtotal(ByVal Taxable As Currency, ByVal Tax As Currency, ByVal Category As String, _
                              ByVal Rate As Double, ByVal Cur As String) As String
    Dim c As String
    c = ZxAttr(Cur)
    ZxTaxSubtotal = "    <cac:TaxSubtotal>" & vbLf & "      <cbc:TaxableAmount currencyID=""" & c & """>" & _
        ZatcaAmount(Taxable) & "</cbc:TaxableAmount>" & vbLf & "      <cbc:TaxAmount currencyID=""" & c & """>" & _
        ZatcaAmount(Tax) & "</cbc:TaxAmount>" & vbLf & "      <cac:TaxCategory>" & vbLf & _
        "        <cbc:ID schemeAgencyID=""6"" schemeID=""UN/ECE 5305"">" & ZxText(Category) & "</cbc:ID>" & vbLf & _
        "        <cbc:Percent>" & ZxPercent(Rate) & "</cbc:Percent>" & vbLf & "        <cac:TaxScheme>" & vbLf & _
        "          <cbc:ID schemeAgencyID=""6"" schemeID=""UN/ECE 5153"">VAT</cbc:ID>" & vbLf & _
        "        </cac:TaxScheme>" & vbLf & "      </cac:TaxCategory>" & vbLf & "    </cac:TaxSubtotal>"
End Function

Public Function ZxInvoice(ByVal Number As String, ByVal Uuid As String, ByVal IssueDate As String, _
                          ByVal IssueTime As String, ByVal TypeCode As String, ByVal SubType As String, _
                          ByVal Cur As String, ByVal BillingRef As String, ByVal Icv As Long, ByVal Pih As String, _
                          ByVal SellerXml As String, ByVal BuyerXml As String, ByVal DeliveryDate As String, _
                          ByVal PaymentCode As String, ByVal Reason As String, ByVal TaxTotal As Currency, _
                          ByVal SubtotalsXml As String, ByVal LineExt As Currency, ByVal TaxExcl As Currency, _
                          ByVal TaxIncl As Currency, ByVal LinesXml As String) As String
    ' The document without the declaration and the signature block, with the QR placeholder.
    Dim s As String, c As String
    c = ZxAttr(Cur)
    s = "<Invoice " & ZATCA_NAMESPACES & ">" & vbLf & "  <cbc:ProfileID>reporting:1.0</cbc:ProfileID>" & vbLf & _
        "  <cbc:ID>" & ZxText(Number) & "</cbc:ID>" & vbLf & "  <cbc:UUID>" & ZxText(Uuid) & "</cbc:UUID>" & vbLf & _
        "  <cbc:IssueDate>" & IssueDate & "</cbc:IssueDate>" & vbLf & "  <cbc:IssueTime>" & IssueTime & "</cbc:IssueTime>" & _
        vbLf & "  <cbc:InvoiceTypeCode name=""" & ZxAttr(SubType) & """>" & ZxText(TypeCode) & "</cbc:InvoiceTypeCode>" & _
        vbLf & "  <cbc:DocumentCurrencyCode>" & c & "</cbc:DocumentCurrencyCode>" & vbLf & "  <cbc:TaxCurrencyCode>" & c & _
        "</cbc:TaxCurrencyCode>"
    If Len(BillingRef) > 0 Then
        s = s & vbLf & "  <cac:BillingReference>" & vbLf & "    <cac:InvoiceDocumentReference>" & vbLf & "      <cbc:ID>" & _
            ZxText(BillingRef) & "</cbc:ID>" & vbLf & "    </cac:InvoiceDocumentReference>" & vbLf & "  </cac:BillingReference>"
    End If
    s = s & vbLf & "  <cac:AdditionalDocumentReference>" & vbLf & "    <cbc:ID>ICV</cbc:ID>" & vbLf & "    <cbc:UUID>" & _
        Icv & "</cbc:UUID>" & vbLf & "  </cac:AdditionalDocumentReference>" & vbLf & "  <cac:AdditionalDocumentReference>" & _
        vbLf & "    <cbc:ID>PIH</cbc:ID>" & vbLf & "    <cac:Attachment>" & vbLf & _
        "      <cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">" & ZxText(Pih) & _
        "</cbc:EmbeddedDocumentBinaryObject>" & vbLf & "    </cac:Attachment>" & vbLf & "  </cac:AdditionalDocumentReference>" & _
        vbLf & "  <cac:AdditionalDocumentReference>" & vbLf & "    <cbc:ID>QR</cbc:ID>" & vbLf & "    <cac:Attachment>" & vbLf & _
        "      <cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">" & ZATCA_QR_PLACEHOLDER & _
        "</cbc:EmbeddedDocumentBinaryObject>" & vbLf & "    </cac:Attachment>" & vbLf & "  </cac:AdditionalDocumentReference>" & _
        vbLf & "  <cac:Signature>" & vbLf & "    <cbc:ID>urn:oasis:names:specification:ubl:signature:Invoice</cbc:ID>" & _
        vbLf & "    <cbc:SignatureMethod>urn:oasis:names:specification:ubl:dsig:enveloped:xades</cbc:SignatureMethod>" & _
        vbLf & "  </cac:Signature>" & vbLf & SellerXml & vbLf & BuyerXml
    If Len(DeliveryDate) > 0 Then
        s = s & vbLf & "  <cac:Delivery>" & vbLf & "    <cbc:ActualDeliveryDate>" & DeliveryDate & "</cbc:ActualDeliveryDate>" & _
            vbLf & "  </cac:Delivery>"
    End If
    s = s & vbLf & "  <cac:PaymentMeans>" & vbLf & "    <cbc:PaymentMeansCode>" & ZxText(PaymentCode) & "</cbc:PaymentMeansCode>"
    If Len(Reason) > 0 Then s = s & vbLf & "    <cbc:InstructionNote>" & ZxText(Reason) & "</cbc:InstructionNote>"
    s = s & vbLf & "  </cac:PaymentMeans>" & vbLf & "  <cac:TaxTotal>" & vbLf & "    <cbc:TaxAmount currencyID=""" & c & """>" & _
        ZatcaAmount(TaxTotal) & "</cbc:TaxAmount>" & vbLf & "  </cac:TaxTotal>" & vbLf & "  <cac:TaxTotal>" & vbLf & _
        "    <cbc:TaxAmount currencyID=""" & c & """>" & ZatcaAmount(TaxTotal) & "</cbc:TaxAmount>" & vbLf & SubtotalsXml & _
        vbLf & "  </cac:TaxTotal>" & vbLf & "  <cac:LegalMonetaryTotal>" & vbLf & _
        "    <cbc:LineExtensionAmount currencyID=""" & c & """>" & ZatcaAmount(LineExt) & "</cbc:LineExtensionAmount>" & vbLf & _
        "    <cbc:TaxExclusiveAmount currencyID=""" & c & """>" & ZatcaAmount(TaxExcl) & "</cbc:TaxExclusiveAmount>" & vbLf & _
        "    <cbc:TaxInclusiveAmount currencyID=""" & c & """>" & ZatcaAmount(TaxIncl) & "</cbc:TaxInclusiveAmount>" & vbLf & _
        "    <cbc:AllowanceTotalAmount currencyID=""" & c & """>0.00</cbc:AllowanceTotalAmount>" & vbLf & _
        "    <cbc:PrepaidAmount currencyID=""" & c & """>0.00</cbc:PrepaidAmount>" & vbLf & _
        "    <cbc:PayableAmount currencyID=""" & c & """>" & ZatcaAmount(TaxIncl) & "</cbc:PayableAmount>" & vbLf & _
        "  </cac:LegalMonetaryTotal>" & vbLf & LinesXml & vbLf & "</Invoice>"
    ZxInvoice = s
End Function

Public Function ZatcaSignedDocument(ByVal Xml As String, ByVal Extensions As String, ByVal QR As String) As String
    ' The file: declaration, the signature block right after the root tag, the QR code in place.
    Dim p As Long
    p = InStr(Xml, ">")
    ZatcaSignedDocument = Replace("<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & Left$(Xml, p) & Extensions & _
                                  Mid$(Xml, p + 1), ZATCA_QR_PLACEHOLDER, QR)
End Function

'==============================================================================
' Hash
'==============================================================================
Private Function RemoveElement(ByVal Xml As String, ByVal StartTag As String, ByVal EndTag As String, _
                               Optional ByVal After As Long = 1) As String
    Dim i As Long, j As Long
    i = InStr(After, Xml, StartTag, vbBinaryCompare)
    If i = 0 Then
        RemoveElement = Xml
        Exit Function
    End If
    j = InStr(i, Xml, EndTag, vbBinaryCompare) + Len(EndTag)
    RemoveElement = Left$(Xml, i - 1) & Mid$(Xml, j)
End Function

Public Function ZatcaHashInput(ByVal Xml As String) As String
    ' What ZATCA hashes: without the declaration, ext:UBLExtensions, cac:Signature and the QR reference
    ' (their surrounding white space stays). The document is canonical, so cutting the text is enough.
    Dim s As String, qr As Long, start As Long
    s = Mid$(Xml, InStr(Xml, "<Invoice"))
    s = RemoveElement(s, "<ext:UBLExtensions>", "</ext:UBLExtensions>")
    s = RemoveElement(s, "<cac:Signature>", "</cac:Signature>")
    qr = InStr(1, s, "<cbc:ID>QR</cbc:ID>", vbBinaryCompare)
    start = InStrRev(s, "<cac:AdditionalDocumentReference>", qr, vbBinaryCompare)
    ZatcaHashInput = RemoveElement(s, "<cac:AdditionalDocumentReference>", "</cac:AdditionalDocumentReference>", start)
End Function

Public Function ZatcaInvoiceHash(ByVal Xml As String) As String
    ' base64 of the SHA-256 of the hashed part (UTF-8)
    Dim b() As Byte, n As Long, h() As Byte
    n = Utf8Bytes(ZatcaHashInput(Xml), b)
    HexToBytes Sha256Hex(b, n), h
    ZatcaInvoiceHash = Base64Encode(h, 32)
End Function

Public Function HexDigestB64(ByVal Text As String) As String
    ' ZATCA's digest of the certificate and of the signed properties: base64 of the hex SHA-256 text.
    Dim b() As Byte, n As Long, hx As String
    n = Utf8Bytes(Text, b)
    hx = Sha256Hex(b, n)
    n = Utf8Bytes(hx, b)
    HexDigestB64 = Base64Encode(b, n)
End Function

Private Sub HexToBytes(ByVal Hx As String, ByRef b() As Byte)
    Dim i As Long
    ReDim b(0 To Len(Hx) \ 2 - 1)
    For i = 0 To Len(Hx) \ 2 - 1
        b(i) = CLng("&H" & Mid$(Hx, i * 2 + 1, 2))
    Next
End Sub

'==============================================================================
' Base64 and UTF-8 (decoding)
'==============================================================================
Public Function Base64Decode(ByVal Text As String, ByRef b() As Byte) As Long
    ' The bytes of a base64 text (white space ignored); returns their count.
    Dim i As Long, v As Long, bits As Long, n As Long, ch As String, k As Long
    ReDim b(0 To (Len(Text) \ 4) * 3 + 3)
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        k = InStr(1, B64_CHARS, ch, vbBinaryCompare) - 1
        If k >= 0 Then
            v = (v * 64 + k) And &HFFFFFF
            bits = bits + 6
            If bits >= 8 Then
                bits = bits - 8
                b(n) = (v \ (2 ^ bits)) And &HFF
                n = n + 1
            End If
        End If
    Next
    If n > 0 Then
        ReDim Preserve b(0 To n - 1)
    End If
    Base64Decode = n
End Function

Public Function Utf8Decode(ByRef b() As Byte, ByVal First As Long, ByVal Count As Long) As String
    Dim i As Long, cp As Long, out As String, last As Long
    i = First
    last = First + Count - 1
    Do While i <= last
        If b(i) < &H80 Then
            cp = b(i)
            i = i + 1
        ElseIf b(i) < &HE0 Then
            cp = (b(i) And &H1F) * 64 + (b(i + 1) And &H3F)
            i = i + 2
        ElseIf b(i) < &HF0 Then
            cp = ((b(i) And &HF) * 64 + (b(i + 1) And &H3F)) * 64 + (b(i + 2) And &H3F)
            i = i + 3
        Else
            cp = (((b(i) And &H7) * 64 + (b(i + 1) And &H3F)) * 64 + (b(i + 2) And &H3F)) * 64 + (b(i + 3) And &H3F)
            i = i + 4
        End If
        If cp >= &H10000 Then
            cp = cp - &H10000
            out = out & ChrW(&HD800 + (cp \ 1024)) & ChrW(&HDC00 + (cp And &H3FF))
        Else
            out = out & ChrW(cp)
        End If
    Loop
    Utf8Decode = out
End Function

'==============================================================================
' Certificate (DER)
'==============================================================================
Private Sub DerRead(ByRef b() As Byte, ByVal Pos As Long, ByRef Tag As Long, ByRef ContentStart As Long, _
                    ByRef EndPos As Long)
    Dim size As Long, n As Long, i As Long
    Tag = b(Pos)
    size = b(Pos + 1)
    ContentStart = Pos + 2
    If size And &H80 Then
        n = size And &H7F
        size = 0
        For i = 0 To n - 1
            size = size * 256 + b(ContentStart + i)
        Next
        ContentStart = ContentStart + n
    End If
    EndPos = ContentStart + size
End Sub

Private Function DerChild(ByRef b() As Byte, ByVal ContentStart As Long, ByVal Index As Long) As Long
    ' The position of child Index (from 0) of the element whose content starts at ContentStart.
    Dim pos As Long, i As Long, tag As Long, cs As Long, ep As Long
    pos = ContentStart
    For i = 1 To Index
        DerRead b, pos, tag, cs, ep
        pos = ep
    Next
    DerChild = pos
End Function

Private Function DerOid(ByRef b() As Byte, ByVal First As Long, ByVal EndPos As Long) As String
    Dim i As Long, v As Double, arcs As String, firstArc As Long
    For i = First To EndPos - 1
        v = v * 128 + (b(i) And &H7F)
        If (b(i) And &H80) = 0 Then
            If Len(arcs) = 0 Then
                firstArc = IIf(v >= 80, 2, Int(v / 40))
                arcs = firstArc & "." & CStr(v - firstArc * 40)
            Else
                arcs = arcs & "." & CStr(v)
            End If
            v = 0
        End If
    Next
    DerOid = arcs
End Function

Private Function DnLabel(ByVal Oid As String) As String
    Select Case Oid
        Case "2.5.4.3": DnLabel = "CN"
        Case "2.5.4.4": DnLabel = "SN"
        Case "2.5.4.5": DnLabel = "SERIALNUMBER"
        Case "2.5.4.6": DnLabel = "C"
        Case "2.5.4.7": DnLabel = "L"
        Case "2.5.4.8": DnLabel = "ST"
        Case "2.5.4.10": DnLabel = "O"
        Case "2.5.4.11": DnLabel = "OU"
        Case "2.5.4.12": DnLabel = "T"
        Case "0.9.2342.19200300.100.1.1": DnLabel = "UID"
        Case "0.9.2342.19200300.100.1.25": DnLabel = "DC"
        Case "1.2.840.113549.1.9.1": DnLabel = "E"
        Case Else: DnLabel = Oid
    End Select
End Function

Private Function DerName(ByRef b() As Byte, ByVal ContentStart As Long, ByVal EndPos As Long) As String
    ' "CN=..., O=..., C=SA": the parts of a Name, last first (as ZATCA writes the issuer).
    Dim pos As Long, tag As Long, cs As Long, ep As Long, acs As Long, aep As Long, ocs As Long, oep As Long
    Dim vcs As Long, vep As Long, label As String, out As String
    pos = ContentStart
    Do While pos < EndPos
        DerRead b, pos, tag, cs, ep                  ' SET
        DerRead b, cs, tag, acs, aep                 ' SEQUENCE: OID, value
        DerRead b, acs, tag, ocs, oep
        DerRead b, oep, tag, vcs, vep
        label = DnLabel(DerOid(b, ocs, oep)) & "=" & Utf8Decode(b, vcs, vep - vcs)
        If Len(out) = 0 Then
            out = label
        Else
            out = label & ", " & out
        End If
        pos = ep
    Loop
    DerName = out
End Function

Private Function BigDecimal(ByRef b() As Byte, ByVal First As Long, ByVal EndPos As Long) As String
    ' A big-endian unsigned integer as decimal text (the serial number of a certificate).
    Dim digits() As Long, n As Long, i As Long, j As Long, carry As Long, s As String
    ReDim digits(0 To (EndPos - First) * 3 + 1)
    n = 1
    For i = First To EndPos - 1
        carry = b(i)
        For j = 0 To n - 1
            carry = digits(j) * 256 + carry
            digits(j) = carry Mod 10
            carry = carry \ 10
        Next
        Do While carry > 0
            digits(n) = carry Mod 10
            carry = carry \ 10
            n = n + 1
        Loop
    Next
    For j = n - 1 To 0 Step -1
        s = s & digits(j)
    Next
    BigDecimal = s
End Function

Public Function ZatcaCertInfo(ByVal CertB64 As String, ByRef Issuer As String, ByRef Serial As String, _
                              ByRef PublicKeyB64 As String, ByRef CertSignatureB64 As String, _
                              ByRef CertHash As String) As String
    ' "" when the certificate (base64 DER, or PEM) can be read. The public key is the SubjectPublicKeyInfo
    ' DER (QR tag 8), the signature of the certificate is QR tag 9.
    Dim b() As Byte, n As Long, tag As Long, rcs As Long, rep As Long, tcs As Long, tep As Long
    Dim pos As Long, cs As Long, ep As Long, k() As Byte, i As Long, off As Long
    On Error GoTo EH
    CertB64 = CertificateBody(CertB64)
    n = Base64Decode(CertB64, b)
    DerRead b, 0, tag, rcs, rep                           ' Certificate
    DerRead b, rcs, tag, tcs, tep                         ' TBSCertificate
    DerRead b, tcs, tag, cs, ep
    If tag = &HA0 Then off = 1                            ' [0] version
    pos = DerChild(b, tcs, off)                           ' serial number
    DerRead b, pos, tag, cs, ep
    Serial = BigDecimal(b, cs, ep)
    pos = DerChild(b, tcs, off + 2)                       ' issuer
    DerRead b, pos, tag, cs, ep
    Issuer = DerName(b, cs, ep)
    pos = DerChild(b, tcs, off + 5)                       ' subjectPublicKeyInfo
    DerRead b, pos, tag, cs, ep
    ReDim k(0 To ep - pos - 1)
    For i = 0 To ep - pos - 1
        k(i) = b(pos + i)
    Next
    PublicKeyB64 = Base64Encode(k, ep - pos)
    pos = DerChild(b, rcs, 2)                             ' signatureValue BIT STRING
    DerRead b, pos, tag, cs, ep
    ReDim k(0 To ep - cs - 2)
    For i = 0 To ep - cs - 2
        k(i) = b(cs + 1 + i)                              ' after the "unused bits" byte
    Next
    CertSignatureB64 = Base64Encode(k, ep - cs - 1)
    CertHash = HexDigestB64(CertB64)
    Exit Function
EH:
    ZatcaCertInfo = "تعذّرت قراءة الشهادة: " & Err.Description
End Function

Public Function CertificateBody(ByVal Text As String) As String
    ' The base64 body of a certificate: without the PEM lines and the line breaks.
    Text = Replace(Replace(Text, "-----BEGIN CERTIFICATE-----", ""), "-----END CERTIFICATE-----", "")
    CertificateBody = Replace(Replace(Replace(Replace(Text, vbCr, ""), vbLf, ""), " ", ""), vbTab, "")
End Function

'==============================================================================
' Signature block and QR code
'==============================================================================
Public Function ZatcaSignedProperties(ByVal SigningTime As String, ByVal CertHash As String, ByVal Issuer As String, _
                                      ByVal Serial As String, ByVal ForHash As Boolean) As String
    Dim t As String
    If ForHash Then
        t = ZatcaTemplate("SIGNED_PROPERTIES_FOR_HASH")
    Else
        t = ZatcaTemplate("SIGNED_PROPERTIES_FOR_XML")
    End If
    t = Replace(t, "{SIGNING_TIME}", SigningTime)
    t = Replace(t, "{CERT_HASH}", CertHash)
    t = Replace(t, "{CERT_ISSUER}", ZxText(Issuer))
    ZatcaSignedProperties = Replace(t, "{CERT_SERIAL}", Serial)
End Function

Public Function ZatcaExtensions(ByVal InvoiceHash As String, ByVal Signature As String, ByVal CertB64 As String, _
                                ByVal SigningTime As String, ByVal CertHash As String, ByVal Issuer As String, _
                                ByVal Serial As String) As String
    Dim t As String
    t = ZatcaTemplate("UBL_EXTENSIONS")
    t = Replace(t, "{INVOICE_HASH}", InvoiceHash)
    t = Replace(t, "{PROPS_DIGEST}", HexDigestB64(ZatcaSignedProperties(SigningTime, CertHash, Issuer, Serial, True)))
    t = Replace(t, "{SIGNATURE}", Signature)
    t = Replace(t, "{CERTIFICATE}", CertB64)
    ZatcaExtensions = Replace(t, "{SIGNED_PROPERTIES}", ZatcaSignedProperties(SigningTime, CertHash, Issuer, Serial, False))
End Function

Private Sub AddTlv(ByRef data() As Byte, ByRef n As Long, ByVal Tag As Byte, ByRef v() As Byte, ByVal k As Long)
    Dim i As Long, head As Long
    If k < &H80 Then
        head = 2
    ElseIf k < &H100 Then
        head = 3
    Else
        head = 4
    End If
    ReDim Preserve data(0 To n + head + k)
    data(n) = Tag
    If head = 2 Then
        data(n + 1) = k
    ElseIf head = 3 Then
        data(n + 1) = &H81
        data(n + 2) = k
    Else
        data(n + 1) = &H82
        data(n + 2) = k \ 256
        data(n + 3) = k Mod 256
    End If
    For i = 0 To k - 1
        data(n + head + i) = v(i)
    Next
    n = n + head + k
End Sub

Public Function ZatcaQRCode(ByVal Seller As String, ByVal Vat As String, ByVal Timestamp As String, _
                            ByVal Total As Currency, ByVal VatTotal As Currency, ByVal InvoiceHash As String, _
                            ByVal Signature As String, ByVal PublicKeyB64 As String, ByVal CertSignatureB64 As String) As String
    ' Nine tags: 1-5 as phase 1, 6 invoice hash and 7 signature (base64 text), 8 public key and
    ' 9 signature of the certificate (bytes).
    Dim data() As Byte, n As Long, v() As Byte, k As Long, values As Variant, i As Long
    ReDim data(0 To 0)
    values = Array(Seller, Vat, Timestamp, ZatcaAmount(Total), ZatcaAmount(VatTotal), InvoiceHash, Signature)
    For i = 0 To 6
        k = Utf8Bytes(CStr(values(i)), v)
        AddTlv data, n, i + 1, v, k
    Next
    k = Base64Decode(PublicKeyB64, v)
    AddTlv data, n, 8, v, k
    k = Base64Decode(CertSignatureB64, v)
    AddTlv data, n, 9, v, k
    ZatcaQRCode = Base64Encode(data, n)
End Function

'==============================================================================
' The document of a sales invoice or return, from the data
'==============================================================================
Private Function Txt(ByVal v As Variant) As String
    Txt = Trim$(Nz(v, ""))
End Function

Public Function ZatcaSellerProblem() As String
    ' The lines of the seller's missing data (tax number and national address), "" when complete.
    Dim s As DAO.Recordset, msg As String
    Set s = CurrentDb.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    If Len(TaxNumberProblem(s!VATNumber, "SA")) > 0 Or Len(Txt(s!VATNumber)) = 0 Then
        msg = msg & "- الرقم الضريبي للمحل" & vbCrLf
    End If
    If Len(Txt(s!StoreName)) = 0 Then msg = msg & "- اسم المحل" & vbCrLf
    If Len(Txt(s!StreetName)) = 0 Then msg = msg & "- شارع المحل" & vbCrLf
    If Not (Txt(s!BuildingNo) Like "####") Then msg = msg & "- رقم مبنى المحل (4 أرقام)" & vbCrLf
    If Not (Txt(s!PostalCode) Like "#####") Then msg = msg & "- الرمز البريدي للمحل (5 أرقام)" & vbCrLf
    If Len(Txt(s!City)) = 0 Then msg = msg & "- مدينة المحل" & vbCrLf
    If Len(Txt(s!District)) = 0 Then msg = msg & "- حي المحل" & vbCrLf
    s.Close
    ZatcaSellerProblem = msg
End Function

Public Function ZatcaDataProblem(ByVal DocKind As String, ByVal DocID As Long) As String
    ' "" when the data ZATCA requires is complete: the seller's national address and tax number, the buyer's for
    ' a tax invoice (B2B), standard-rated lines.
    Dim c As DAO.Recordset, msg As String, subType As String
    msg = ZatcaSellerProblem()
    subType = Nz(DbValue("SELECT InvoiceSubType FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & _
                         DocID), "")
    If subType = "STANDARD" Then
        Set c = CurrentDb.OpenRecordset("SELECT c.* FROM " & DocTableOf(DocKind) & " AS h INNER JOIN Customers AS c ON " & _
                                        "h.CustomerID = c.CustomerID WHERE h." & DocKeyOf(DocKind) & " = " & DocID, dbOpenSnapshot)
        If Not c.EOF Then
            If Len(Txt(c!VATNumber)) = 0 Or Len(TaxNumberProblem(c!VATNumber, "SA")) > 0 Then
                msg = msg & "- الرقم الضريبي للعميل" & vbCrLf
            End If
            If Len(Txt(c!StreetName)) = 0 Then msg = msg & "- شارع العميل" & vbCrLf
            If Not (Txt(c!BuildingNo) Like "####") Then msg = msg & "- رقم مبنى العميل (4 أرقام)" & vbCrLf
            If Not (Txt(c!PostalCode) Like "#####") Then msg = msg & "- الرمز البريدي للعميل (5 أرقام)" & vbCrLf
            If Len(Txt(c!City)) = 0 Then msg = msg & "- مدينة العميل" & vbCrLf
            If Len(Txt(c!District)) = 0 Then msg = msg & "- حي العميل" & vbCrLf
        End If
        c.Close
    End If
    If Nz(DbValue("SELECT COUNT(*) FROM " & LinesTable(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID & _
                  " AND VATCategory <> 'S'"), 0) > 0 Then
        msg = msg & "- أصناف معفاة أو بنسبة صفرية: تحتاج سبب الإعفاء، ولم يُدعم بعد" & vbCrLf
    End If
    If Len(msg) > 0 Then ZatcaDataProblem = "بيانات ناقصة للفاتورة الإلكترونية:" & vbCrLf & msg
End Function

Private Function LinesTable(ByVal DocKind As String) As String
    If DocKind = "RETURN" Then
        LinesTable = "SalesReturnDetails"
    Else
        LinesTable = "SalesInvoiceDetails"
    End If
End Function

Private Function ZxPaymentMeans(ByVal PaymentType As String, ByVal MethodID As Variant) As String
    ' UN/ECE 4461: 10 cash, 30 credit transfer (sale on account), 42 bank account, 48 bank card
    If PaymentType = "CREDIT" Then
        ZxPaymentMeans = "30"
    Else
        Select Case Nz(MethodID, 1)
            Case 2: ZxPaymentMeans = "48"
            Case 3: ZxPaymentMeans = "42"
            Case Else: ZxPaymentMeans = "10"
        End Select
    End If
End Function

Public Function ZatcaDocumentXml(ByVal DocKind As String, ByVal DocID As Long, ByVal Pih As String, _
                                 ByRef Totals As Variant) As String
    ' The unsigned document (QR placeholder) of a sales invoice or return. Totals = Array(issue timestamp,
    ' total with VAT, VAT) for the QR code.
    Dim db As DAO.Database, h As DAO.Recordset, s As DAO.Recordset, c As DAO.Recordset, rs As DAO.Recordset
    Dim isReturn As Boolean, seller As String, buyer As String, linesXml As String, subtotals As String
    Dim docDate As Date, issueDate As String, issueTime As String, n As Long, billing As String, reason As String
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    Set db = CurrentDb
    isReturn = (DocKind = "RETURN")
    If isReturn Then
        Set h = db.OpenRecordset("SELECT r.*, r.ReturnNumber AS DocNumber, r.ReturnDate AS DocDate, r.RefundType AS PayType, " & _
                                 "o.InvoiceNumber AS OriginalNumber FROM SalesReturns AS r INNER JOIN SalesInvoices AS o " & _
                                 "ON r.SalesInvoiceID = o.SalesInvoiceID WHERE r.SalesReturnID = " & DocID, dbOpenSnapshot)
        billing = Txt(h!OriginalNumber)
        reason = Txt(h!Reason)
        If Len(reason) = 0 Then reason = "Return"
    Else
        Set h = db.OpenRecordset("SELECT h.*, h.InvoiceNumber AS DocNumber, h.InvoiceDate AS DocDate, h.PaymentType AS PayType " & _
                                 "FROM SalesInvoices AS h WHERE h.SalesInvoiceID = " & DocID, dbOpenSnapshot)
    End If
    docDate = h!DocDate
    issueDate = Format$(docDate, "yyyy-mm-dd")
    issueTime = Format$(docDate, "hh:nn:ss")
    Set s = db.OpenRecordset("SELECT * FROM Settings WHERE SettingID = 1", dbOpenSnapshot)
    seller = ZxParty("Supplier", Txt(s!StoreName), Txt(s!VATNumber), Txt(s!CRNumber), Txt(s!StreetName), _
                     Txt(s!BuildingNo), Txt(s!AdditionalNo), Txt(s!District), Txt(s!City), Txt(s!PostalCode), "SA")
    If h!InvoiceSubType = "STANDARD" Then
        Set c = db.OpenRecordset("SELECT * FROM Customers WHERE CustomerID = " & h!CustomerID, dbOpenSnapshot)
        buyer = ZxParty("Customer", Txt(c!CustomerName), Txt(c!VATNumber), "", Txt(c!StreetName), Txt(c!BuildingNo), _
                        "", Txt(c!District), Txt(c!City), Txt(c!PostalCode), "SA")
        c.Close
    Else
        buyer = ZxParty("Customer", "", "", "", "", "", "", "", "", "", "")
    End If
    Set rs = db.OpenRecordset("SELECT d.Quantity, d.UnitPrice, d.Discount, d.NetAmount, d.VATCategory, d.VATRate, d.Tax, " & _
                              "d.LineTotal, p.ProductName, u.ZatcaUnitCode FROM (" & LinesTable(DocKind) & " AS d INNER JOIN " & _
                              "Products AS p ON d.ProductID = p.ProductID) INNER JOIN Units AS u ON p.UnitID = u.UnitID " & _
                              "WHERE d." & DocKeyOf(DocKind) & " = " & DocID & " ORDER BY d." & _
                              IIf(isReturn, "ReturnDetailID", "SalesDetailID"), dbOpenSnapshot)
    Do Until rs.EOF
        n = n + 1
        If n > 1 Then linesXml = linesXml & vbLf
        linesXml = linesXml & ZxLine(n, rs!Quantity, Nz(rs!ZatcaUnitCode, "PCE"), rs!NetAmount, Nz(rs!Discount, 0), _
                                     rs!Tax, rs!LineTotal, Txt(rs!ProductName), rs!VATCategory, rs!VATRate, rs!UnitPrice, "SAR")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT VATCategory, VATRate, Sum(NetAmount) AS SumNet, Sum(Tax) AS SumTax FROM " & _
                              LinesTable(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID & _
                              " GROUP BY VATCategory, VATRate ORDER BY VATCategory, VATRate", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(subtotals) > 0 Then subtotals = subtotals & vbLf
        subtotals = subtotals & ZxTaxSubtotal(rs!SumNet, rs!SumTax, rs!VATCategory, rs!VATRate, "SAR")
        rs.MoveNext
    Loop
    rs.Close
    ZatcaDocumentXml = ZxInvoice(Txt(h!DocNumber), Txt(h!InvoiceUUID), issueDate, issueTime, Txt(h!InvoiceTypeCode), _
        IIf(h!InvoiceSubType = "STANDARD", "0100000", "0200000"), "SAR", billing, Nz(h!ICV, 0), Pih, seller, buyer, _
        IIf(h!InvoiceSubType = "STANDARD", issueDate, ""), ZxPaymentMeans(Txt(h!PayType), h!PaymentMethodID), reason, _
        h!Tax, subtotals, h!TaxableAmount, h!TaxableAmount, h!TotalAmount, linesXml)
    Totals = Array(issueDate & "T" & issueTime, CCur(h!TotalAmount), CCur(h!Tax), Txt(s!StoreName), Txt(s!VATNumber))
    s.Close
    h.Close
    Calendar = saved
End Function

'==============================================================================
' Signing with OpenSSL
'==============================================================================
Private Function Q(ByVal Path As String) As String
    Q = """" & Path & """"
End Function

Public Function ZatcaWorkFolder() As String
    ZatcaWorkFolder = Environ$("TEMP") & "\RetailStoreZatca"
    If Len(Dir$(ZatcaWorkFolder, vbDirectory)) = 0 Then MkDir ZatcaWorkFolder
End Function

Public Function OpenSslExe() As String
    OpenSslExe = Txt(SettingValue("OpenSslPath"))
    If Len(OpenSslExe) = 0 Then OpenSslExe = "openssl"
End Function

Public Function RunOpenSsl(ByVal Arguments As String, ByRef ErrorText As String) As Long
    ' Runs openssl (Settings.OpenSslPath, else on the PATH), waits, returns its exit code; its messages in ErrorText.
    Dim errFile As String, f As Integer, ln As String
    errFile = ZatcaWorkFolder() & "\openssl.err"
    If Len(Dir$(errFile)) > 0 Then Kill errFile
    RunOpenSsl = CreateObject("WScript.Shell").Run("cmd /s /c """ & Q(OpenSslExe()) & " " & Arguments & " 2> " & _
                                                   Q(errFile) & """", 0, True)
    ErrorText = ""
    If Len(Dir$(errFile)) > 0 Then
        f = FreeFile
        Open errFile For Input As #f
        Do Until EOF(f)
            Line Input #f, ln
            ErrorText = ErrorText & ln & " "
        Loop
        Close #f
    End If
End Function

Private Sub WriteBytes(ByVal Path As String, ByRef b() As Byte, ByVal n As Long)
    Dim f As Integer, part() As Byte, i As Long
    If Len(Dir$(Path)) > 0 Then Kill Path
    ReDim part(0 To n - 1)
    For i = 0 To n - 1
        part(i) = b(i)
    Next
    f = FreeFile
    Open Path For Binary Access Write As #f
    Put #f, , part
    Close #f
End Sub

Private Function ReadBytes(ByVal Path As String, ByRef b() As Byte) As Long
    Dim f As Integer
    f = FreeFile
    Open Path For Binary Access Read As #f
    ReadBytes = LOF(f)
    If ReadBytes > 0 Then
        ReDim b(0 To ReadBytes - 1)
        Get #f, , b
    End If
    Close #f
End Function

Public Sub WriteUtf8File(ByVal Path As String, ByVal Text As String)
    ' UTF-8 without a byte order mark.
    Dim b() As Byte, n As Long
    n = Utf8Bytes(Text, b)
    WriteBytes Path, b, n
End Sub

Public Function ZatcaSignHash(ByVal HashB64 As String, ByRef SignatureB64 As String) As String
    ' ECDSA (secp256k1, SHA-256) signature of the 32 bytes of the invoice hash, with the key file of the
    ' settings: "" or the problem.
    Dim keyFile As String, hashFile As String, sigFile As String, b() As Byte, n As Long, errText As String
    keyFile = Txt(SettingValue("ZatcaKeyFile"))
    If Len(keyFile) = 0 Or Len(Dir$(keyFile)) = 0 Then
        ZatcaSignHash = "ملف المفتاح الخاص غير موجود: " & keyFile
        Exit Function
    End If
    hashFile = ZatcaWorkFolder() & "\hash.bin"
    sigFile = ZatcaWorkFolder() & "\sig.der"
    n = Base64Decode(HashB64, b)
    WriteBytes hashFile, b, n
    If Len(Dir$(sigFile)) > 0 Then Kill sigFile
    If RunOpenSsl("dgst -sha256 -sign " & Q(keyFile) & " -out " & Q(sigFile) & " " & Q(hashFile), errText) <> 0 Or _
       Len(Dir$(sigFile)) = 0 Then
        ZatcaSignHash = "تعذّر التوقيع ببرنامج OpenSSL: " & errText
        Exit Function
    End If
    n = ReadBytes(sigFile, b)
    SignatureB64 = Base64Encode(b, n)
    Kill sigFile
    Kill hashFile
End Function

Public Function ZatcaSignNow() As String
    ' The signing time of the signed properties: "2026-10-08T14:05:10" (UTC).
    Dim saved As Integer, t As Date
    saved = Calendar
    Calendar = vbCalGreg
    t = DateAdd("h", -KSA_UTC_OFFSET_HOURS, Now)
    ZatcaSignNow = Format$(t, "yyyy-mm-dd") & "T" & Format$(t, "hh:nn:ss")
    Calendar = saved
End Function

Public Function ZatcaBuildSigned(ByVal DocKind As String, ByVal DocID As Long, ByVal Pih As String, _
                                 ByRef SignedXml As String, ByRef InvoiceHash As String, ByRef QR As String) As String
    ' The signed document of a sales invoice or return with this previous hash: "" or the problem.
    Dim xml As String, totals As Variant, cert As String, msg As String
    msg = ZatcaDataProblem(DocKind, DocID)
    If Len(msg) > 0 Then
        ZatcaBuildSigned = msg
        Exit Function
    End If
    cert = CertificateBody(Nz(SettingValue("ZatcaCertificate"), ""))
    If Len(cert) = 0 Then
        ZatcaBuildSigned = "لا توجد شهادة الجهاز (CSID) في إعداد الربط."
        Exit Function
    End If
    xml = ZatcaDocumentXml(DocKind, DocID, Pih, totals)
    ZatcaBuildSigned = ZatcaSignXml(xml, cert, totals, SignedXml, InvoiceHash, QR)
End Function

Public Function ZatcaSignXml(ByVal Xml As String, ByVal CertB64 As String, ByVal Totals As Variant, _
                             ByRef SignedXml As String, ByRef InvoiceHash As String, ByRef QR As String) As String
    ' Signs an unsigned document with this certificate and the key file of the settings. Totals = Array(issue
    ' timestamp, total with VAT, VAT, seller name, seller VAT number). "" or the problem.
    Dim issuer As String, serial As String, pub As String, certSig As String, certHash As String, signature As String
    Dim msg As String
    msg = ZatcaCertInfo(CertB64, issuer, serial, pub, certSig, certHash)
    If Len(msg) > 0 Then
        ZatcaSignXml = msg
        Exit Function
    End If
    InvoiceHash = ZatcaInvoiceHash(Xml)
    msg = ZatcaSignHash(InvoiceHash, signature)
    If Len(msg) > 0 Then
        ZatcaSignXml = msg
        Exit Function
    End If
    QR = ZatcaQRCode(Totals(3), Totals(4), Totals(0), Totals(1), Totals(2), InvoiceHash, signature, pub, certSig)
    SignedXml = ZatcaSignedDocument(Xml, ZatcaExtensions(InvoiceHash, signature, CertificateBody(CertB64), ZatcaSignNow(), _
                                                         certHash, issuer, serial), QR)
End Function

Public Function ZatcaPrepareDocument(ByVal DocKind As String, ByVal DocID As Long) As String
    ' Signs a document once, in the order of the chain: its previous hash is the last hash of the settings,
    ' then its own hash becomes the last one. "" or the problem; a signed document is left as it is.
    Dim ws As DAO.Workspace, db As DAO.Database, inTrans As Boolean, pih As String, xml As String, hash As String
    Dim qr As String, msg As String, rs As DAO.Recordset
    On Error GoTo EH
    Set db = CurrentDb
    If Len(Nz(DbValue("SELECT EInvoiceXml FROM " & DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID), _
              "")) > 0 Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    db.Execute "UPDATE Settings SET LastInvoiceHash = LastInvoiceHash WHERE SettingID = 1", dbFailOnError    ' lock the chain
    pih = Nz(DbValue("SELECT LastInvoiceHash FROM Settings WHERE SettingID = 1"), ZATCA_INITIAL_PIH)
    msg = ZatcaBuildSigned(DocKind, DocID, pih, xml, hash, qr)
    If Len(msg) > 0 Then
        ws.Rollback
        ZatcaPrepareDocument = msg
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT InvoiceHash, PreviousInvoiceHash, QRCodeData, EInvoiceXml FROM " & _
                              DocTableOf(DocKind) & " WHERE " & DocKeyOf(DocKind) & " = " & DocID, dbOpenDynaset)
    rs.Edit
    rs!InvoiceHash = hash
    rs!PreviousInvoiceHash = pih
    rs!QRCodeData = qr
    rs!EInvoiceXml = xml                                  ' a memo: longer than a SQL statement may be
    rs.Update
    rs.Close
    db.Execute "UPDATE Settings SET LastInvoiceHash = '" & hash & "' WHERE SettingID = 1", dbFailOnError
    ws.CommitTrans
    Exit Function
EH:
    ZatcaPrepareDocument = "تعذّر توقيع المستند: " & Err.Description
    If inTrans Then ws.Rollback
End Function

'==============================================================================
' Test key and certificate (test environment only)
'==============================================================================
Public Function ZatcaMakeTestCertificate(ByVal Folder As String) As String
    ' A secp256k1 key and a self-signed certificate made by OpenSSL, to try the signing before the
    ' device has its certificate from ZATCA (phase C2). Saved in the settings. "" or the problem.
    Dim keyFile As String, certFile As String, cnf As String, errText As String, f As Integer, ln As String
    Dim body As String
    keyFile = Folder & "\zatca-test-key.pem"
    certFile = Folder & "\zatca-test-cert.pem"
    cnf = ZatcaWorkFolder() & "\test-cert.cnf"
    WriteUtf8File cnf, "[req]" & vbLf & "distinguished_name = dn" & vbLf & "prompt = no" & vbLf & "[dn]" & vbLf & _
                       "C = SA" & vbLf & "O = Test" & vbLf & "CN = TST-EGS-RetailStore" & vbLf
    If RunOpenSsl("ecparam -name secp256k1 -genkey -noout -out " & Q(keyFile), errText) <> 0 Then
        ZatcaMakeTestCertificate = "تعذّر تشغيل OpenSSL: " & errText
        Exit Function
    End If
    If RunOpenSsl("req -new -x509 -sha256 -days 365 -config " & Q(cnf) & " -key " & Q(keyFile) & " -out " & _
                  Q(certFile), errText) <> 0 Then
        ZatcaMakeTestCertificate = "تعذّر إنشاء الشهادة: " & errText
        Exit Function
    End If
    f = FreeFile
    Open certFile For Input As #f
    Do Until EOF(f)
        Line Input #f, ln
        body = body & ln
    Loop
    Close #f
    CurrentDb.Execute "UPDATE Settings SET ZatcaKeyFile = " & SqlText(keyFile) & ", ZatcaCertificate = " & _
                      SqlText(CertificateBody(body)) & " WHERE SettingID = 1", dbFailOnError
End Function

'==============================================================================
' frmZatcaSetup: OpenSSL, the key file and the certificate of the device
'==============================================================================
Public Sub ZatcaSetupLoad(ByVal frm As Access.Form)
    frm!txtOpenSsl.Value = SettingValue("OpenSslPath")
    frm!txtKeyFile.Value = SettingValue("ZatcaKeyFile")
    frm!txtCertificate.Value = SettingValue("ZatcaCertificate")
    ZatcaSetupShowCert frm
End Sub

Public Sub ZatcaSetupShowCert(ByVal frm As Access.Form)
    Dim issuer As String, serial As String, pub As String, sig As String, h As String, msg As String
    If Len(CertificateBody(Nz(frm!txtCertificate.Value, ""))) = 0 Then
        frm!lblCertInfo.Caption = Tr("لا توجد شهادة بعد.")
        Exit Sub
    End If
    msg = ZatcaCertInfo(Nz(frm!txtCertificate.Value, ""), issuer, serial, pub, sig, h)
    If Len(msg) > 0 Then
        frm!lblCertInfo.Caption = Tr(msg)
    Else
        frm!lblCertInfo.Caption = Tr("الجهة المصدرة: ") & issuer & "    " & Tr("الرقم التسلسلي: ") & serial
    End If
End Sub

Public Sub ZatcaSetupBrowseKey(ByVal frm As Access.Form)
    With Application.FileDialog(3)          ' msoFileDialogFilePicker
        .Title = Tr("اختر ملف المفتاح الخاص")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add Tr("المفتاح الخاص"), "*.pem;*.key"
        If .Show Then frm!txtKeyFile.Value = .SelectedItems(1)
    End With
End Sub

Public Sub ZatcaSetupSave(ByVal frm As Access.Form)
    Dim before As Collection, rs As DAO.Recordset
    If Not HasPermission("SETTINGS") Then
        ShowWarning "إعداد الربط يحتاج صلاحية إعدادات المحل."
        Exit Sub
    End If
    Set before = AuditSnapshot("Settings", "SettingID", 1)
    Set rs = CurrentDb.OpenRecordset("SELECT OpenSslPath, ZatcaKeyFile, ZatcaCertificate FROM Settings WHERE SettingID = 1", _
                                     dbOpenDynaset)
    rs.Edit
    rs!OpenSslPath = IIf(Len(Trim$(Nz(frm!txtOpenSsl.Value, ""))) = 0, Null, Trim$(Nz(frm!txtOpenSsl.Value, "")))
    rs!ZatcaKeyFile = IIf(Len(Trim$(Nz(frm!txtKeyFile.Value, ""))) = 0, Null, Trim$(Nz(frm!txtKeyFile.Value, "")))
    rs!ZatcaCertificate = IIf(Len(CertificateBody(Nz(frm!txtCertificate.Value, ""))) = 0, Null, _
                              CertificateBody(Nz(frm!txtCertificate.Value, "")))
    rs.Update
    rs.Close
    AuditEdited "EDIT", "Settings", "SettingID", 1, before
    ZatcaSetupLoad frm
    ShowInfo "تم حفظ إعداد الربط."
End Sub

Public Sub ZatcaSetupTestCertificate(ByVal frm As Access.Form)
    Dim msg As String, folder As String
    If Not HasPermission("SETTINGS") Then
        ShowWarning "إعداد الربط يحتاج صلاحية إعدادات المحل."
        Exit Sub
    End If
    If Nz(SettingValue("EInvoiceEnvironment"), "TEST") <> "TEST" Then
        ShowWarning "الشهادة التجريبية للبيئة التجريبية فقط."
        Exit Sub
    End If
    If Len(Nz(SettingValue("ZatcaCertificate"), "")) > 0 Then
        If Not AskYesNo("توجد شهادة محفوظة. هل تريد استبدالها بشهادة تجريبية؟") Then Exit Sub
    End If
    folder = CurrentProject.Path & "\ZATCA"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    msg = ZatcaMakeTestCertificate(folder)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    ZatcaSetupLoad frm
    ShowInfo "أُنشئ مفتاح وشهادة تجريبية في المجلد:" & vbCrLf & folder
End Sub

Public Sub ZatcaSetupExportXml(ByVal frm As Access.Form)
    ' The signed file of one invoice, to check it with ZATCA's SDK (fatoora -validate). The chain of hashes
    ' and the invoice are not changed.
    Dim number As String, id As Variant, xml As String, hash As String, qr As String, msg As String, path As String
    number = Trim$(Nz(InputBox(Tr("رقم فاتورة البيع:"), Tr("ملف XML")), ""))
    If Len(number) = 0 Then Exit Sub
    id = DbValue("SELECT SalesInvoiceID FROM SalesInvoices WHERE InvoiceNumber = " & SqlText(number))
    If IsNull(id) Then
        ShowWarning "الفاتورة غير موجودة: " & number
        Exit Sub
    End If
    DoCmd.Hourglass True
    msg = ZatcaBuildSigned("SALE", CLng(id), Nz(SettingValue("LastInvoiceHash"), ZATCA_INITIAL_PIH), xml, hash, qr)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    path = ReportsFolder() & "\ZATCA-" & number & ".xml"
    WriteUtf8File path, xml
    ShowInfo "تم إنشاء الملف الموقّع:" & vbCrLf & path & vbCrLf & "بصمة الفاتورة: " & hash
End Sub

'==============================================================================
' In-Access test
'==============================================================================
Public Function TestZatcaXml() As Boolean
    Dim passed As Long, failed As Long, report As String, xml As String, msg As String, sig As String
    Dim id As Variant, totals As Variant
    Calendar = vbCalGreg
    g_SilentMode = True
    Debug.Print "=== TestZatcaXml  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record ZxPrice(8.6957) = "8.6957" And ZxPrice(10) = "10.00" And ZxPrice(2.5) = "2.50" And _
           ZxQuantity(2) = "2.000000", "صيغة السعر والكمية", passed, failed, report
    xml = TestSampleXml()
    Record ZatcaInvoiceHash(xml) = SAMPLE_HASH, "بصمة فاتورة نموذجية = النسخة المكتوبة بـ Python", passed, failed, report
    Record ZatcaInvoiceHash(ZatcaSignedDocument(xml, "<ext:UBLExtensions>x</ext:UBLExtensions>", "QRDATA")) = SAMPLE_HASH, _
           "البصمة لا تتغير بإضافة التوقيع ورمز QR", passed, failed, report
    id = DbValue("SELECT Max(SalesInvoiceID) FROM SalesInvoices")
    If Not IsNull(id) Then
        xml = ZatcaDocumentXml("SALE", CLng(id), ZATCA_INITIAL_PIH, totals)
        Record InStr(xml, "<cbc:ID>" & ZxText(Nz(DbValue("SELECT InvoiceNumber FROM SalesInvoices WHERE SalesInvoiceID = " & _
               id), "")) & "</cbc:ID>") > 0 And InStr(xml, ZATCA_QR_PLACEHOLDER) > 0, "ملف آخر فاتورة بيع", _
               passed, failed, report
    End If
    If Len(Nz(SettingValue("ZatcaKeyFile"), "")) > 0 Then
        msg = ZatcaSignHash(SAMPLE_HASH, sig)
        Record Len(msg) = 0 And Len(sig) > 60, "التوقيع ببرنامج OpenSSL " & msg, passed, failed, report
    Else
        Debug.Print "[--] لا يوجد مفتاح خاص في إعداد الربط: لم يُختبر التوقيع"
    End If
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات ملف الفاتورة السعودية ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestZatcaXml"
        TestZatcaXml = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestZatcaXml"
    End If
End Function

Public Function TestSampleXml() As String
    ' A fixed simplified invoice (tests/test_zatca_xml.py builds the same one with the Python mirror).
    TestSampleXml = ZxInvoice("INV-0001", "3cf5ee18-ee25-44ea-a444-2c37ba7f28be", "2026-10-08", "14:05:09", "388", "0200000", _
        "SAR", "", 7, ZATCA_INITIAL_PIH, _
        ZxParty("Supplier", "مؤسسة النور & أولاده", "399999999900003", "1010010000", "الملك فهد", "1234", "5678", "العليا", _
                "الرياض", "12345", "SA"), ZxParty("Customer", "", "", "", "", "", "", "", "", "", ""), "", "10", "", _
        15, ZxTaxSubtotal(100, 15, "S", 0.15, "SAR"), 100, 100, 115, _
        ZxLine(1, 2, "PCE", 100, 0.87, 15, 115, "منتج ""أ"" <1>", "S", 0.15, 50.4348, "SAR"))
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
