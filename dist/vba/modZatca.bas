Attribute VB_Name = "modZatca"
'==============================================================================
' modZatca  -  Retail Store Management System (Phase 6)
'
' ZATCA (Fatoora) phase 1: the QR code payload printed on every invoice.
'   payload = Base64( TLV(1 seller) TLV(2 VAT no) TLV(3 timestamp)
'                     TLV(4 total incl. VAT) TLV(5 VAT total) )
'   TLV = tag byte + length byte + UTF-8 bytes.
' Pure VBA: no MSXML/ADODB dependency, works on 32- and 64-bit Office.
' Reference implementation and test vectors: tools/zatca_reference.py
'==============================================================================
Option Compare Database
Option Explicit

Public Const KSA_UTC_OFFSET_HOURS As Integer = 3    ' Saudi Arabia: UTC+3, no DST

Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID_T) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID_T) As Long
#End If

Private Const B64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

'------------------------------------------------------------------------------
' QR payload
'------------------------------------------------------------------------------
Public Function BuildZatcaQR(ByVal SellerName As String, ByVal VatNumber As String, _
                             ByVal InvoiceDate As Date, ByVal TotalWithVat As Currency, _
                             ByVal VatTotal As Currency) As String
    Dim data() As Byte, n As Long
    ReDim data(0 To 0)
    n = 0
    AppendTLV data, n, 1, FitTo255Bytes(SellerName)
    AppendTLV data, n, 2, VatNumber
    AppendTLV data, n, 3, ZatcaTimestamp(InvoiceDate)
    AppendTLV data, n, 4, ZatcaAmount(TotalWithVat)
    AppendTLV data, n, 5, ZatcaAmount(VatTotal)
    BuildZatcaQR = Base64Encode(data, n)
End Function

Public Function ZatcaAmount(ByVal Value As Currency) As String
    ' "1150.00": two decimals, "." separator and ASCII digits whatever the Windows locale.
    Dim v As Currency, whole As Currency, cents As Long, sign As String
    v = RoundMoney(Value)
    If v < 0 Then
        sign = "-"
        v = -v
    End If
    whole = Fix(v)
    cents = CLng((v - whole) * 100)
    ZatcaAmount = sign & Format$(whole, "0") & "." & Right$("0" & CStr(cents), 2)
End Function

Public Function ZatcaTimestamp(ByVal LocalTime As Date) As String
    ' Invoice time in Saudi local time -> ISO 8601 UTC, e.g. 2026-10-04T09:15:00Z
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    ZatcaTimestamp = Format$(DateAdd("h", -KSA_UTC_OFFSET_HOURS, LocalTime), "yyyy-mm-dd") & "T" & _
                     Format$(DateAdd("h", -KSA_UTC_OFFSET_HOURS, LocalTime), "hh:nn:ss") & "Z"
    Calendar = saved
End Function

Private Function FitTo255Bytes(ByVal s As String) As String
    ' A TLV length is one byte: shorten very long names on a character boundary.
    Do While Utf8Length(s) > 255
        s = Left$(s, Len(s) - 1)
    Loop
    FitTo255Bytes = s
End Function

Private Sub AppendTLV(ByRef data() As Byte, ByRef n As Long, ByVal Tag As Byte, ByVal Value As String)
    Dim b() As Byte, k As Long, i As Long
    k = Utf8Bytes(Value, b)
    If k > 255 Then Err.Raise vbObjectError + 900, "AppendTLV", "TLV value too long (tag " & Tag & ")"
    ReDim Preserve data(0 To n + k + 1)
    data(n) = Tag
    data(n + 1) = k
    For i = 0 To k - 1
        data(n + 2 + i) = b(i)
    Next
    n = n + k + 2
End Sub

'------------------------------------------------------------------------------
' UTF-8 and Base64
'------------------------------------------------------------------------------
Public Function Utf8Bytes(ByVal s As String, ByRef b() As Byte) As Long
    ' Encodes s as UTF-8 into b(0..n-1) and returns n (surrogate pairs supported).
    Dim i As Long, n As Long, cp As Long, lo As Long
    ReDim b(0 To Len(s) * 4 + 1)
    i = 1
    Do While i <= Len(s)
        cp = AscW(Mid$(s, i, 1)) And &HFFFF&
        If cp >= &HD800& And cp <= &HDBFF& And i < Len(s) Then
            lo = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                cp = &H10000 + (cp - &HD800&) * &H400& + (lo - &HDC00&)
                i = i + 1
            End If
        End If
        If cp < &H80& Then
            b(n) = cp
            n = n + 1
        ElseIf cp < &H800& Then
            b(n) = &HC0& Or (cp \ &H40&)
            b(n + 1) = &H80& Or (cp And &H3F&)
            n = n + 2
        ElseIf cp < &H10000 Then
            b(n) = &HE0& Or (cp \ &H1000&)
            b(n + 1) = &H80& Or ((cp \ &H40&) And &H3F&)
            b(n + 2) = &H80& Or (cp And &H3F&)
            n = n + 3
        Else
            b(n) = &HF0& Or (cp \ &H40000)
            b(n + 1) = &H80& Or ((cp \ &H1000&) And &H3F&)
            b(n + 2) = &H80& Or ((cp \ &H40&) And &H3F&)
            b(n + 3) = &H80& Or (cp And &H3F&)
            n = n + 4
        End If
        i = i + 1
    Loop
    Utf8Bytes = n
End Function

Public Function Utf8Length(ByVal s As String) As Long
    Dim b() As Byte
    Utf8Length = Utf8Bytes(s, b)
End Function

Public Function Base64Encode(ByRef data() As Byte, ByVal n As Long) As String
    Dim i As Long, b0 As Long, b1 As Long, b2 As Long, out As String
    For i = 0 To n - 1 Step 3
        b0 = data(i)
        b1 = -1
        b2 = -1
        If i + 1 < n Then b1 = data(i + 1)
        If i + 2 < n Then b2 = data(i + 2)
        out = out & Mid$(B64, (b0 \ 4) + 1, 1)
        If b1 < 0 Then
            out = out & Mid$(B64, ((b0 And 3) * 16) + 1, 1) & "=="
        ElseIf b2 < 0 Then
            out = out & Mid$(B64, ((b0 And 3) * 16 + (b1 \ 16)) + 1, 1) & _
                        Mid$(B64, ((b1 And 15) * 4) + 1, 1) & "="
        Else
            out = out & Mid$(B64, ((b0 And 3) * 16 + (b1 \ 16)) + 1, 1) & _
                        Mid$(B64, ((b1 And 15) * 4 + (b2 \ 64)) + 1, 1) & _
                        Mid$(B64, (b2 And 63) + 1, 1)
        End If
    Next
    Base64Encode = out
End Function

Public Function Base64Text(ByVal s As String) As String
    ' Base64 of the UTF-8 bytes of s (used by the tests).
    Dim b() As Byte, n As Long
    n = Utf8Bytes(s, b)
    Base64Text = Base64Encode(b, n)
End Function

'------------------------------------------------------------------------------
' Identifiers
'------------------------------------------------------------------------------
Public Function NewUUID() As String
    ' Random UUID, lower case: 8-4-4-4-12 hex digits (needed by ZATCA phase 2).
    Dim g As GUID_T, s As String, i As Integer
    If CoCreateGuid(g) <> 0 Then Err.Raise vbObjectError + 901, "NewUUID", "CoCreateGuid failed"
    s = Right$("00000000" & Hex$(g.Data1), 8) & "-" & _
        Right$("0000" & Hex$(g.Data2 And &HFFFF&), 4) & "-" & _
        Right$("0000" & Hex$(g.Data3 And &HFFFF&), 4) & "-"
    For i = 0 To 7
        s = s & Right$("0" & Hex$(g.Data4(i)), 2)
        If i = 1 Then s = s & "-"
    Next
    NewUUID = LCase$(s)
End Function

Public Function InvoiceSubType(ByVal CustomerID As Long) As String
    ' B2B (customer has a VAT number) -> standard tax invoice, otherwise simplified.
    If Len(Nz(DbValue("SELECT VATNumber FROM Customers WHERE CustomerID = " & CustomerID), "")) > 0 Then
        InvoiceSubType = "STANDARD"
    Else
        InvoiceSubType = "SIMPLIFIED"
    End If
End Function

Public Function ZatcaInitialStatus() As String
    ' Phase 2 integration picks up PENDING documents; phase 1 needs no submission.
    If Nz(SettingValue("ZatcaPhase"), 1) = 2 Then
        ZatcaInitialStatus = "PENDING"
    Else
        ZatcaInitialStatus = "NOT_SENT"
    End If
End Function
