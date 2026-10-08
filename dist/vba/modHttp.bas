Attribute VB_Name = "modHttp"
'==============================================================================
' modHttp  -  Retail Store Management System (HTTPS and JSON for e-invoicing)
'
' HttpSend      one HTTPS request (MSXML2.ServerXMLHTTP.6.0): the status and the reply, or the network
'               problem. Used by the e-invoicing platforms (docs/45-EInvoice-Foundation.md).
' JsonGet       a value of a JSON reply by its path ("validationResults.status", "items[0].code").
' JsonEscape    a text inside a JSON string.
' tools/json_reference.py is the Python mirror of JsonGet and JsonEscape.
'==============================================================================
Option Compare Database
Option Explicit

'==============================================================================
' HTTPS
'==============================================================================
Public Function HttpSend(ByVal Method As String, ByVal Url As String, ByVal Headers As String, _
                         ByVal Body As String, ByRef Status As Long, ByRef ResponseText As String, _
                         Optional ByVal TimeoutSeconds As Long = 30) As String
    ' "" when the server answered (whatever its HTTP status), else the network problem (Status = 0).
    ' Headers: "Name: value" lines separated by vbLf. A text body is sent in UTF-8.
    Dim http As Object, item As Variant, colon As Long
    Status = 0
    ResponseText = ""
    On Error GoTo EH
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 15000, 15000, TimeoutSeconds * 1000, TimeoutSeconds * 1000
    http.Open Method, Url, False
    For Each item In Split(Headers, vbLf)
        colon = InStr(item, ":")
        If colon > 1 Then http.setRequestHeader Trim$(Left$(item, colon - 1)), Trim$(Mid$(item, colon + 1))
    Next
    If Len(Body) > 0 Then
        http.send Body
    Else
        http.send
    End If
    Status = http.Status
    ResponseText = http.responseText
    Exit Function
EH:
    Status = 0
    HttpSend = "ÊÚÐøÑ ÇáÇÊÕÇá ÈÇáãäÙæãÉ: " & Err.Description
End Function

'==============================================================================
' JSON
'==============================================================================
Public Function JsonEscape(ByVal Text As String) As String
    ' The text as the inside of a JSON string: \" \\ and the control characters escaped.
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        code = AscW(ch)
        If code = 34 Then
            out = out & "\"""
        ElseIf code = 92 Then
            out = out & "\\"
        ElseIf code = 10 Then
            out = out & "\n"
        ElseIf code = 13 Then
            out = out & "\r"
        ElseIf code = 9 Then
            out = out & "\t"
        ElseIf code >= 0 And code < 32 Then
            out = out & "\u" & Right$("0000" & LCase$(Hex$(code)), 4)
        Else
            out = out & ch
        End If
    Next
    JsonEscape = out
End Function

Public Function JsonGet(ByVal Json As String, ByVal Path As String) As Variant
    ' The value at Path: a string unescaped, a number / true / false as written, an object or array as its
    ' JSON text; Null when the path is missing or the value is null. Path "" = the whole value.
    Dim pos As Long, seg As Variant, segs As String
    JsonGet = Null
    pos = 1
    JsonSkip Json, pos
    If pos > Len(Json) Then Exit Function
    segs = Replace(Replace(Replace(Path, "[", ".#"), "]", ""), "..", ".")
    If Left$(segs, 1) = "." Then segs = Mid$(segs, 2)
    If Len(segs) > 0 Then
        For Each seg In Split(segs, ".")
            If Left$(seg, 1) = "#" Then
                pos = JsonItemAt(Json, pos, CLng(Mid$(seg, 2)))
            Else
                pos = JsonMemberAt(Json, pos, CStr(seg))
            End If
            If pos = 0 Then Exit Function
        Next
    End If
    JsonGet = JsonValueAt(Json, pos)
End Function

Private Sub JsonSkip(ByVal Json As String, ByRef pos As Long)
    Do While pos <= Len(Json)
        Select Case AscW(Mid$(Json, pos, 1))
            Case 32, 9, 10, 13
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function JsonEnd(ByVal Json As String, ByVal Start As Long) As Long
    ' The position just after the value that starts at Start.
    Dim pos As Long, depth As Long, ch As String, inString As Boolean
    pos = Start
    ch = Mid$(Json, pos, 1)
    If ch = """" Then
        pos = pos + 1
        Do While pos <= Len(Json)
            ch = Mid$(Json, pos, 1)
            If ch = "\" Then
                pos = pos + 2
            ElseIf ch = """" Then
                JsonEnd = pos + 1
                Exit Function
            Else
                pos = pos + 1
            End If
        Loop
        JsonEnd = pos
    ElseIf ch = "{" Or ch = "[" Then
        Do While pos <= Len(Json)
            ch = Mid$(Json, pos, 1)
            If inString Then
                If ch = "\" Then
                    pos = pos + 1
                ElseIf ch = """" Then
                    inString = False
                End If
            ElseIf ch = """" Then
                inString = True
            ElseIf ch = "{" Or ch = "[" Then
                depth = depth + 1
            ElseIf ch = "}" Or ch = "]" Then
                depth = depth - 1
                If depth = 0 Then
                    JsonEnd = pos + 1
                    Exit Function
                End If
            End If
            pos = pos + 1
        Loop
        JsonEnd = pos
    Else
        Do While pos <= Len(Json)
            ch = Mid$(Json, pos, 1)
            If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
            pos = pos + 1
        Loop
        JsonEnd = pos
    End If
End Function

Private Function JsonMemberAt(ByVal Json As String, ByVal pos As Long, ByVal Name As String) As Long
    ' The position of the value of member Name of the object at pos, 0 when there is none.
    Dim key As String
    If Mid$(Json, pos, 1) <> "{" Then Exit Function
    pos = pos + 1
    Do
        JsonSkip Json, pos
        If Mid$(Json, pos, 1) <> """" Then Exit Function
        key = JsonStringAt(Json, pos)
        pos = JsonEnd(Json, pos)
        JsonSkip Json, pos
        If Mid$(Json, pos, 1) <> ":" Then Exit Function
        pos = pos + 1
        JsonSkip Json, pos
        If StrComp(key, Name, vbBinaryCompare) = 0 Then
            JsonMemberAt = pos
            Exit Function
        End If
        pos = JsonEnd(Json, pos)
        JsonSkip Json, pos
        If Mid$(Json, pos, 1) <> "," Then Exit Function
        pos = pos + 1
    Loop
End Function

Private Function JsonItemAt(ByVal Json As String, ByVal pos As Long, ByVal Index As Long) As Long
    ' The position of item Index (from 0) of the array at pos, 0 when there is none.
    Dim i As Long
    If Mid$(Json, pos, 1) <> "[" Then Exit Function
    pos = pos + 1
    JsonSkip Json, pos
    If Mid$(Json, pos, 1) = "]" Then Exit Function
    Do
        If i = Index Then
            JsonItemAt = pos
            Exit Function
        End If
        pos = JsonEnd(Json, pos)
        JsonSkip Json, pos
        If Mid$(Json, pos, 1) <> "," Then Exit Function
        pos = pos + 1
        JsonSkip Json, pos
        i = i + 1
    Loop
End Function

Private Function JsonValueAt(ByVal Json As String, ByVal pos As Long) As Variant
    Dim ch As String
    ch = Mid$(Json, pos, 1)
    If ch = """" Then
        JsonValueAt = JsonStringAt(Json, pos)
    ElseIf Mid$(Json, pos, 4) = "null" Then
        JsonValueAt = Null
    Else
        JsonValueAt = Mid$(Json, pos, JsonEnd(Json, pos) - pos)
    End If
End Function

Private Function JsonStringAt(ByVal Json As String, ByVal pos As Long) As String
    ' The text of the JSON string that starts (with its quote) at pos.
    Dim ch As String, out As String
    pos = pos + 1
    Do While pos <= Len(Json)
        ch = Mid$(Json, pos, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            pos = pos + 1
            ch = Mid$(Json, pos, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & ChrW(8)
                Case "f": out = out & ChrW(12)
                Case "u"
                    out = out & ChrW(CLng("&H" & Mid$(Json, pos + 1, 4)))
                    pos = pos + 4
                Case Else: out = out & ch
            End Select
        Else
            out = out & ch
        End If
        pos = pos + 1
    Loop
    JsonStringAt = out
End Function
