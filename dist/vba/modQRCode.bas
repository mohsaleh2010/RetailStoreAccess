Attribute VB_Name = "modQRCode"
'==============================================================================
' modQRCode  -  Retail Store Management System (Phase 6)
'
' GENERATED FILE - do not edit by hand.  Source: tools/gen_qr.py
' QR Code encoder (ISO/IEC 18004): byte mode, error correction level M,
' versions 1-20. VBA port of tools/qr_reference.py, which is verified
' module-by-module against the "qrcode" library and decoded with OpenCV.
'
'   QRMatrix(Text, M)        builds the code; M(row, col) = 1 for a dark module
'   DrawQR Report, Text, ...  draws it on a report section (Print event)
'   QRChecksum(Text)          fingerprint used by the tests
'==============================================================================
Option Compare Database
Option Explicit

Private Const MAX_VERSION As Integer = 20

Private m_exp(0 To 511) As Integer
Private m_log(0 To 255) As Integer
Private m_ready As Boolean
Private m_bits() As Byte
Private m_bitCount As Long
Private m_mod() As Byte
Private m_res() As Boolean
Private m_size As Integer
Private m_lastVersion As Integer
Private m_lastMask As Integer

'------------------------------------------------------------------------------
' Public API
'------------------------------------------------------------------------------
Public Function QRMatrix(ByVal Text As String, ByRef M() As Byte, _
                         Optional ByVal ForceMask As Integer = -1) As Integer
    Dim data() As Byte, n As Long, version As Integer, cw() As Integer
    Dim candidate() As Byte, msk As Integer, score As Long, best As Long
    InitGF
    n = Utf8Bytes(Text, data)
    version = ChooseVersion(n)
    cw = MakeCodewords(data, n, version)
    BaseMatrix version
    PlaceData cw
    best = -1
    For msk = 0 To 7
        If ForceMask < 0 Or ForceMask = msk Then
            CopyMatrix m_mod, candidate
            ApplyMask candidate, msk
            WriteFormat candidate, msk, version
            score = Penalty(candidate)
            If best < 0 Or score < best Then
                best = score
                CopyMatrix candidate, M
                m_lastMask = msk
            End If
        End If
    Next
    m_lastVersion = version
    QRMatrix = m_size
End Function

Public Function QRChecksum(ByVal Text As String, Optional ByVal ForceMask As Integer = -1) As String
    ' "size:" + hex digits of all modules row by row (same as qr_reference.checksum)
    Dim M() As Byte, size As Integer, r As Integer, c As Integer, k As Integer, nib As Integer
    Dim out As String
    size = QRMatrix(Text, M, ForceMask)
    For r = 0 To size - 1
        For c = 0 To size - 1
            nib = nib * 2 + M(r, c)
            k = k + 1
            If k = 4 Then
                out = out & Hex$(nib)
                k = 0
                nib = 0
            End If
        Next
    Next
    If k > 0 Then out = out & Hex$(nib * (2 ^ (4 - k)))
    QRChecksum = size & ":" & out
End Function

Public Function QRLastInfo() As String
    QRLastInfo = "V" & m_lastVersion & "-M" & m_lastMask
End Function

Public Sub DrawQR(ByVal rpt As Access.Report, ByVal Text As String, ByVal LeftTwips As Long, _
                  ByVal TopTwips As Long, ByVal SizeTwips As Long)
    ' Call from the Print event of the report section that holds the code.
    Dim M() As Byte, size As Integer, cell As Long, r As Integer, c As Integer, runStart As Integer
    Dim x0 As Long, y0 As Long
    If Len(Text) = 0 Then Exit Sub
    size = QRMatrix(Text, M)
    cell = SizeTwips \ (size + 8)                ' 4-module quiet zone on each side
    If cell < 1 Then cell = 1
    x0 = LeftTwips + 4 * cell
    y0 = TopTwips + 4 * cell
    For r = 0 To size - 1
        c = 0
        Do While c < size
            If M(r, c) = 1 Then
                runStart = c
                Do While c < size
                    If M(r, c) = 0 Then Exit Do
                    c = c + 1
                Loop
                rpt.Line (x0 + runStart * cell, y0 + r * cell)-(x0 + c * cell - 1, y0 + (r + 1) * cell - 1), 0, BF
            Else
                c = c + 1
            End If
        Loop
    Next
End Sub

'------------------------------------------------------------------------------
' Galois field GF(256) and Reed-Solomon
'------------------------------------------------------------------------------
Private Sub InitGF()
    Dim i As Integer, x As Integer
    If m_ready Then Exit Sub
    x = 1
    For i = 0 To 254
        m_exp(i) = x
        m_log(x) = i
        x = x * 2
        If (x And &H100) <> 0 Then x = x Xor &H11D
    Next
    For i = 255 To 511
        m_exp(i) = m_exp(i - 255)
    Next
    m_ready = True
End Sub

Private Function GfMul(ByVal a As Integer, ByVal b As Integer) As Integer
    If a = 0 Or b = 0 Then Exit Function
    GfMul = m_exp(m_log(a) + m_log(b))
End Function

Private Function RSGenerator(ByVal n As Integer) As Variant
    ' Multiplies (x - a^i) in place, highest degree first. No array assignment,
    ' so the result does not depend on copy-vs-reference array semantics.
    Dim g() As Integer, i As Integer, j As Integer
    ReDim g(0 To 0)
    g(0) = 1
    For i = 0 To n - 1
        ReDim Preserve g(0 To i + 1)
        g(i + 1) = 0
        For j = i + 1 To 1 Step -1
            g(j) = g(j) Xor GfMul(g(j - 1), m_exp(i))
        Next
    Next
    RSGenerator = g
End Function

'------------------------------------------------------------------------------
' Capacity tables (level M)
'------------------------------------------------------------------------------
Private Sub BlockInfo(ByVal v As Integer, ByRef ec As Integer, ByRef b1 As Integer, _
                      ByRef d1 As Integer, ByRef b2 As Integer, ByRef d2 As Integer)
    Dim t As Variant
    Select Case v
        Case 1: t = Array(10, 1, 16, 0, 0)
        Case 2: t = Array(16, 1, 28, 0, 0)
        Case 3: t = Array(26, 1, 44, 0, 0)
        Case 4: t = Array(18, 2, 32, 0, 0)
        Case 5: t = Array(24, 2, 43, 0, 0)
        Case 6: t = Array(16, 4, 27, 0, 0)
        Case 7: t = Array(18, 4, 31, 0, 0)
        Case 8: t = Array(22, 2, 38, 2, 39)
        Case 9: t = Array(22, 3, 36, 2, 37)
        Case 10: t = Array(26, 4, 43, 1, 44)
        Case 11: t = Array(30, 1, 50, 4, 51)
        Case 12: t = Array(22, 6, 36, 2, 37)
        Case 13: t = Array(22, 8, 37, 1, 38)
        Case 14: t = Array(24, 4, 40, 5, 41)
        Case 15: t = Array(24, 5, 41, 5, 42)
        Case 16: t = Array(28, 7, 45, 3, 46)
        Case 17: t = Array(28, 10, 46, 1, 47)
        Case 18: t = Array(26, 9, 43, 4, 44)
        Case 19: t = Array(26, 3, 44, 11, 45)
        Case 20: t = Array(26, 3, 41, 13, 42)
    End Select
    ec = t(0): b1 = t(1): d1 = t(2): b2 = t(3): d2 = t(4)
End Sub

Private Function AlignPositions(ByVal v As Integer) As Variant
    Select Case v
        Case 2: AlignPositions = Array(6, 18)
        Case 3: AlignPositions = Array(6, 22)
        Case 4: AlignPositions = Array(6, 26)
        Case 5: AlignPositions = Array(6, 30)
        Case 6: AlignPositions = Array(6, 34)
        Case 7: AlignPositions = Array(6, 22, 38)
        Case 8: AlignPositions = Array(6, 24, 42)
        Case 9: AlignPositions = Array(6, 26, 46)
        Case 10: AlignPositions = Array(6, 28, 50)
        Case 11: AlignPositions = Array(6, 30, 54)
        Case 12: AlignPositions = Array(6, 32, 58)
        Case 13: AlignPositions = Array(6, 34, 62)
        Case 14: AlignPositions = Array(6, 26, 46, 66)
        Case 15: AlignPositions = Array(6, 26, 48, 70)
        Case 16: AlignPositions = Array(6, 26, 50, 74)
        Case 17: AlignPositions = Array(6, 30, 54, 78)
        Case 18: AlignPositions = Array(6, 30, 56, 82)
        Case 19: AlignPositions = Array(6, 30, 58, 86)
        Case 20: AlignPositions = Array(6, 34, 62, 90)
    End Select
End Function

Private Function DataCodewords(ByVal v As Integer) As Integer
    Dim ec As Integer, b1 As Integer, d1 As Integer, b2 As Integer, d2 As Integer
    BlockInfo v, ec, b1, d1, b2, d2
    DataCodewords = b1 * d1 + b2 * d2
End Function

Private Function CountBits(ByVal v As Integer) As Integer
    If v <= 9 Then CountBits = 8 Else CountBits = 16
End Function

Private Function ChooseVersion(ByVal n As Long) As Integer
    Dim v As Integer
    For v = 1 To MAX_VERSION
        If 4 + CountBits(v) + 8 * n <= 8 * CLng(DataCodewords(v)) Then
            ChooseVersion = v
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 800, "QRMatrix", "ÇáäÕ ÃØæá ãä ÓÚÉ ÑãÒ QR"
End Function

'------------------------------------------------------------------------------
' Codewords
'------------------------------------------------------------------------------
Private Sub PutBits(ByVal Value As Long, ByVal Length As Integer)
    Dim i As Integer
    For i = Length - 1 To 0 Step -1
        m_bits(m_bitCount) = (Value \ CLng(2 ^ i)) And 1
        m_bitCount = m_bitCount + 1
    Next
End Sub

Private Function MakeCodewords(ByRef data() As Byte, ByVal n As Long, ByVal v As Integer) As Variant
    Dim capacity As Long, nData As Integer, i As Long, j As Long, k As Long, value As Long
    Dim ec As Integer, b1 As Integer, d1 As Integer, b2 As Integer, d2 As Integer
    Dim nBlocks As Integer, maxD As Integer, bi As Integer, blockLen As Integer, pos As Long
    Dim dataCw() As Integer, blocks() As Integer, ecs() As Integer, blen() As Integer
    Dim gen() As Integer, remd() As Integer, factor As Integer, out() As Integer

    nData = DataCodewords(v)
    capacity = 8 * CLng(nData)
    ReDim m_bits(0 To capacity + 32)
    m_bitCount = 0
    PutBits 4, 4                                  ' byte mode
    PutBits n, CountBits(v)
    For i = 0 To n - 1
        PutBits data(i), 8
    Next
    k = capacity - m_bitCount                      ' terminator
    If k > 4 Then k = 4
    PutBits 0, CInt(k)
    Do While m_bitCount Mod 8 <> 0
        PutBits 0, 1
    Loop

    ReDim dataCw(0 To nData - 1)
    For i = 0 To m_bitCount \ 8 - 1
        value = 0
        For j = 0 To 7
            value = value * 2 + m_bits(i * 8 + j)
        Next
        dataCw(i) = value
    Next
    k = 0
    For i = m_bitCount \ 8 To nData - 1          ' pad bytes EC 11 EC 11 ...
        If k Mod 2 = 0 Then dataCw(i) = &HEC Else dataCw(i) = &H11
        k = k + 1
    Next

    BlockInfo v, ec, b1, d1, b2, d2
    nBlocks = b1 + b2
    maxD = d1
    If d2 > maxD Then maxD = d2
    ReDim blocks(0 To nBlocks - 1, 0 To maxD - 1)
    ReDim blen(0 To nBlocks - 1)
    ReDim ecs(0 To nBlocks - 1, 0 To ec - 1)
    gen = RSGenerator(ec)
    pos = 0
    For bi = 0 To nBlocks - 1
        If bi < b1 Then blockLen = d1 Else blockLen = d2
        blen(bi) = blockLen
        For j = 0 To blockLen - 1
            blocks(bi, j) = dataCw(pos + j)
        Next
        pos = pos + blockLen
        ReDim remd(0 To ec - 1)
        For j = 0 To blockLen - 1
            factor = blocks(bi, j) Xor remd(0)
            For k = 0 To ec - 2
                remd(k) = remd(k + 1)
            Next
            remd(ec - 1) = 0
            For k = 0 To ec - 1
                remd(k) = remd(k) Xor GfMul(gen(k + 1), factor)
            Next
        Next
        For k = 0 To ec - 1
            ecs(bi, k) = remd(k)
        Next
    Next

    ReDim out(0 To nData + nBlocks * ec - 1)
    pos = 0
    For i = 0 To maxD - 1
        For bi = 0 To nBlocks - 1
            If i < blen(bi) Then
                out(pos) = blocks(bi, i)
                pos = pos + 1
            End If
        Next
    Next
    For i = 0 To ec - 1
        For bi = 0 To nBlocks - 1
            out(pos) = ecs(bi, i)
            pos = pos + 1
        Next
    Next
    MakeCodewords = out
End Function

'------------------------------------------------------------------------------
' Matrix
'------------------------------------------------------------------------------
Private Sub SetFn(ByVal r As Integer, ByVal c As Integer, ByVal v As Byte)
    m_mod(r, c) = v
    m_res(r, c) = True
End Sub

Private Sub BaseMatrix(ByVal version As Integer)
    Dim r As Integer, c As Integer, rr As Integer, cc As Integer, i As Integer, j As Integer
    Dim corner As Integer, r0 As Integer, c0 As Integer, pos As Variant, last As Integer
    Dim isOn As Boolean
    m_size = 17 + 4 * version
    ReDim m_mod(0 To m_size - 1, 0 To m_size - 1)
    ReDim m_res(0 To m_size - 1, 0 To m_size - 1)

    For corner = 0 To 2                             ' finder patterns + separators
        r0 = 0: c0 = 0
        If corner = 1 Then c0 = m_size - 7
        If corner = 2 Then r0 = m_size - 7
        For r = -1 To 7
            For c = -1 To 7
                rr = r0 + r
                cc = c0 + c
                If rr >= 0 And rr < m_size And cc >= 0 And cc < m_size Then
                    isOn = False
                    If r >= 0 And r <= 6 And c >= 0 And c <= 6 Then
                        If r = 0 Or r = 6 Or c = 0 Or c = 6 Then isOn = True
                        If r >= 2 And r <= 4 And c >= 2 And c <= 4 Then isOn = True
                    End If
                    SetFn rr, cc, IIf(isOn, 1, 0)
                End If
            Next
        Next
    Next
    For i = 8 To m_size - 9                         ' timing patterns
        SetFn 6, i, IIf(i Mod 2 = 0, 1, 0)
        SetFn i, 6, IIf(i Mod 2 = 0, 1, 0)
    Next
    pos = AlignPositions(version)                   ' alignment patterns
    If IsArray(pos) Then
        last = UBound(pos)
        For i = 0 To last
            For j = 0 To last
                If Not ((i = 0 And j = 0) Or (i = 0 And j = last) Or (i = last And j = 0)) Then
                    For r = -2 To 2
                        For c = -2 To 2
                            isOn = (Abs(r) = 2 Or Abs(c) = 2 Or (r = 0 And c = 0))
                            SetFn pos(i) + r, pos(j) + c, IIf(isOn, 1, 0)
                        Next
                    Next
                End If
            Next
        Next
    End If
    SetFn m_size - 8, 8, 1                          ' dark module
    For i = 0 To 8                                  ' format information areas
        m_res(8, i) = True
        m_res(i, 8) = True
    Next
    For i = 0 To 7
        m_res(8, m_size - 1 - i) = True
        m_res(m_size - 1 - i, 8) = True
    Next
    If version >= 7 Then                            ' version information areas
        For i = 0 To 5
            For j = 0 To 2
                m_res(i, m_size - 11 + j) = True
                m_res(m_size - 11 + j, i) = True
            Next
        Next
    End If
End Sub

Private Sub PlaceData(ByRef cw() As Integer)
    Dim nBits As Long, idx As Long, upward As Boolean, col As Integer, k As Integer
    Dim r As Integer, c As Integer, stepDir As Integer, rowStart As Integer, rowEnd As Integer, bit As Byte
    nBits = (UBound(cw) + 1) * 8
    upward = True
    col = m_size - 1
    Do While col > 0
        If col = 6 Then col = col - 1
        If upward Then
            rowStart = m_size - 1: rowEnd = 0: stepDir = -1
        Else
            rowStart = 0: rowEnd = m_size - 1: stepDir = 1
        End If
        For r = rowStart To rowEnd Step stepDir
            For k = 0 To 1
                c = col - k
                If Not m_res(r, c) Then
                    bit = 0
                    If idx < nBits Then bit = (cw(idx \ 8) \ CInt(2 ^ (7 - (idx Mod 8)))) And 1
                    m_mod(r, c) = bit
                    idx = idx + 1
                End If
            Next
        Next
        upward = Not upward
        col = col - 2
    Loop
End Sub

Private Sub CopyMatrix(ByRef src() As Byte, ByRef dst() As Byte)
    Dim r As Integer, c As Integer
    ReDim dst(0 To m_size - 1, 0 To m_size - 1)
    For r = 0 To m_size - 1
        For c = 0 To m_size - 1
            dst(r, c) = src(r, c)
        Next
    Next
End Sub

Private Function MaskBit(ByVal msk As Integer, ByVal r As Integer, ByVal c As Integer) As Boolean
    Select Case msk
        Case 0: MaskBit = ((r + c) Mod 2 = 0)
        Case 1: MaskBit = (r Mod 2 = 0)
        Case 2: MaskBit = (c Mod 3 = 0)
        Case 3: MaskBit = ((r + c) Mod 3 = 0)
        Case 4: MaskBit = (((r \ 2) + (c \ 3)) Mod 2 = 0)
        Case 5: MaskBit = (((r * c) Mod 2) + ((r * c) Mod 3) = 0)
        Case 6: MaskBit = ((((r * c) Mod 2) + ((r * c) Mod 3)) Mod 2 = 0)
        Case Else: MaskBit = ((((r + c) Mod 2) + ((r * c) Mod 3)) Mod 2 = 0)
    End Select
End Function

Private Sub ApplyMask(ByRef M() As Byte, ByVal msk As Integer)
    Dim r As Integer, c As Integer
    For r = 0 To m_size - 1
        For c = 0 To m_size - 1
            If Not m_res(r, c) Then
                If MaskBit(msk, r, c) Then M(r, c) = M(r, c) Xor 1
            End If
        Next
    Next
End Sub

Private Function BchFormat(ByVal msk As Integer) As Long
    Dim data As Long, v As Long, i As Integer
    data = msk                                      ' level M = 00
    v = data * 1024&
    For i = 14 To 10 Step -1
        If (v And CLng(2 ^ i)) <> 0 Then v = v Xor (&H537& * CLng(2 ^ (i - 10)))
    Next
    BchFormat = ((data * 1024&) Or v) Xor &H5412&
End Function

Private Function BchVersion(ByVal version As Integer) As Long
    Dim v As Long, i As Integer
    v = version * 4096&
    For i = 17 To 12 Step -1
        If (v And CLng(2 ^ i)) <> 0 Then v = v Xor (&H1F25& * CLng(2 ^ (i - 12)))
    Next
    BchVersion = (version * 4096&) Or v
End Function

Private Sub WriteFormat(ByRef M() As Byte, ByVal msk As Integer, ByVal version As Integer)
    Dim f As Long, i As Integer, bit As Byte, v As Long
    f = BchFormat(msk)
    For i = 0 To 14
        bit = (f \ CLng(2 ^ i)) And 1
        If i < 6 Then
            M(i, 8) = bit
        ElseIf i < 8 Then
            M(i + 1, 8) = bit
        Else
            M(m_size - 15 + i, 8) = bit
        End If
        If i < 8 Then
            M(8, m_size - 1 - i) = bit
        ElseIf i < 9 Then
            M(8, 7) = bit
        Else
            M(8, 14 - i) = bit
        End If
    Next
    M(m_size - 8, 8) = 1
    If version >= 7 Then
        v = BchVersion(version)
        For i = 0 To 17
            bit = (v \ CLng(2 ^ i)) And 1
            M(i \ 3, m_size - 11 + (i Mod 3)) = bit
            M(m_size - 11 + (i Mod 3), i \ 3) = bit
        Next
    End If
End Sub

'------------------------------------------------------------------------------
' Mask penalty (same rules and order as qr_reference.penalty)
'------------------------------------------------------------------------------
Private Function Penalty(ByRef M() As Byte) As Long
    Dim score As Long, r As Integer, c As Integer, i As Integer, k As Integer, ln As Integer
    Dim runLen As Integer, prev As Integer, v As Integer, dark As Long, total As Long
    Dim p1 As Boolean, p2 As Boolean, cell As Integer
    Dim pat1 As Variant, pat2 As Variant
    pat1 = Array(1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0)
    pat2 = Array(0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1)

    For ln = 0 To 2 * m_size - 1                  ' rule 1: rows then columns
        runLen = 0
        prev = -1
        For i = 0 To m_size - 1
            If ln < m_size Then v = M(ln, i) Else v = M(i, ln - m_size)
            If v = prev Then
                runLen = runLen + 1
            Else
                If runLen >= 5 Then score = score + runLen - 2
                runLen = 1
                prev = v
            End If
        Next
        If runLen >= 5 Then score = score + runLen - 2
    Next
    For r = 0 To m_size - 2                         ' rule 2: 2x2 blocks
        For c = 0 To m_size - 2
            v = M(r, c)
            If v = M(r, c + 1) And v = M(r + 1, c) And v = M(r + 1, c + 1) Then score = score + 3
        Next
    Next
    For ln = 0 To 2 * m_size - 1                  ' rule 3: finder-like patterns
        For i = 0 To m_size - 11
            p1 = True
            p2 = True
            For k = 0 To 10
                If ln < m_size Then cell = M(ln, i + k) Else cell = M(i + k, ln - m_size)
                If cell <> pat1(k) Then p1 = False
                If cell <> pat2(k) Then p2 = False
            Next
            If p1 Or p2 Then score = score + 40
        Next
    Next
    For r = 0 To m_size - 1                         ' rule 4: dark proportion
        For c = 0 To m_size - 1
            dark = dark + M(r, c)
        Next
    Next
    total = CLng(m_size) * m_size
    score = score + 10 * (Abs(dark * 20 - total * 10) \ total)
    Penalty = score
End Function
