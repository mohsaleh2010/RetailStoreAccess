Attribute VB_Name = "modLang"
'==============================================================================
' modLang  -  Retail Store Management System (the interface language)
'
' GENERATED FILE - do not edit by hand.
' Source: tools/gen_lang.py and the dictionary tools/i18n_en.py
'
' Each front-end file has one interface language, kept in its database property
' UILanguage: "AR" (default, right-to-left) or "EN" (left-to-right, mirrored).
'   SetInterfaceLanguage "EN"   then BuildQueries, BuildForms, BuildReports
'                               (BuildFrontEnd.vbs asks for the language).
'   Tr(text)    the text in the interface language: in an English file every
'               Arabic phrase of the dictionary is replaced (the longest first,
'               never inside a longer Arabic word). The builders translate the
'               screens and reports with it; ShowMessage, AskYesNo and the
'               captions set by code translate at run time.
'   UiAlign     left / right text alignment for the interface direction.
'   LangSql     [@Accounts] in SQL -> Accounts, or in English qryLocAccounts: the
'               names of the master data in English (Tr calls it in both languages).
'   MSG_RTL     the right-to-left flags of MsgBox (0 in English).
' The tax invoice and the credit note stay bilingual in both languages (ZATCA).
'==============================================================================
Option Compare Database
Option Explicit

Private Const ENTRY_COUNT As Long = 2419

Private m_lang As String              ' "" = not read yet
Private m_loaded As Boolean
Private m_count As Long
Private m_ar() As String
Private m_en() As String

'------------------------------------------------------------------------------
' The language of this front-end file
'------------------------------------------------------------------------------
Public Function UiLanguage() As String
    If Len(m_lang) = 0 Then
        m_lang = "AR"
        On Error Resume Next
        m_lang = UCase$(CStr(CurrentDb.Properties("UILanguage").Value))
        On Error GoTo 0
        If m_lang <> "EN" Then m_lang = "AR"
    End If
    UiLanguage = m_lang
End Function

Public Function UiEnglish() As Boolean
    UiEnglish = (UiLanguage() = "EN")
End Function

Public Sub SetInterfaceLanguage(ByVal LangCode As String)
    ' "AR" or "EN" for this front-end file. Then run BuildQueries, BuildForms and BuildReports.
    Dim db As DAO.Database, p As DAO.Property
    LangCode = UCase$(Trim$(LangCode))
    If LangCode <> "EN" Then LangCode = "AR"
    Set db = CurrentDb
    On Error Resume Next
    db.Properties("UILanguage").Value = LangCode
    If Err.Number <> 0 Then
        Err.Clear
        Set p = db.CreateProperty("UILanguage", dbText, LangCode)
        db.Properties.Append p
    End If
    On Error GoTo 0
    m_lang = LangCode
    Debug.Print "Interface language: " & LangCode & " - run BuildQueries, BuildForms and BuildReports."
End Sub

Public Sub UseLanguage(ByVal LangCode As String)
    ' This session only (the tests): does not change the file.
    m_lang = UCase$(LangCode)
End Sub

Public Function MSG_RTL() As Long
    ' vbMsgBoxRight + vbMsgBoxRtlReading in the Arabic interface
    If UiEnglish() Then
        MSG_RTL = 0
    Else
        MSG_RTL = &H180000
    End If
End Function

Public Function UiAlign(ByVal TextAlign As Integer) As Integer
    ' Left (1) and right (3) change places in the left-to-right interface; general and centre stay.
    UiAlign = TextAlign
    If UiEnglish() Then
        If TextAlign = 1 Then
            UiAlign = 3
        ElseIf TextAlign = 3 Then
            UiAlign = 1
        End If
    End If
End Function

'------------------------------------------------------------------------------
' Translation
'------------------------------------------------------------------------------
Public Function Tr(ByVal Text As Variant) As Variant
    Dim s As String, i As Long
    Tr = Text
    If VarType(Text) <> vbString Then Exit Function
    s = Text
    If InStr(1, s, "[@", vbBinaryCompare) > 0 Then         ' a master table with English names (both languages)
        s = LangSql(s)
        Tr = s
    End If
    If Not UiEnglish() Then Exit Function
    If Not HasArabic(s) Then Exit Function
    If Not m_loaded Then LoadDictionary
    For i = 1 To m_count
        If InStr(1, s, m_ar(i), vbBinaryCompare) > 0 Then
            s = ReplaceWord(s, m_ar(i), m_en(i))
            If Not HasArabic(s) Then Exit For
        End If
    Next
    s = Replace(s, ChrW(&H60C), ",", 1, -1, vbBinaryCompare)
    s = Replace(s, ChrW(&H61B), ",", 1, -1, vbBinaryCompare)
    s = Replace(s, ChrW(&H61F), "?", 1, -1, vbBinaryCompare)
    s = Replace(s, ChrW(&HAB), "", 1, -1, vbBinaryCompare)          ' the Arabic quotes
    s = Replace(s, ChrW(&HBB), "", 1, -1, vbBinaryCompare)
    Tr = s
End Function

Public Function LangSql(ByVal Text As String) As String
    ' [@Accounts] in SQL: the table Accounts, or in English the saved query qryLocAccounts (the same
    ' columns, with the English name of each row in the name column). tools/i18n.py resolve_names.
    Dim out As String, pos As Long, i As Long, j As Long
    pos = 1
    Do
        i = InStr(pos, Text, "[@", vbBinaryCompare)
        If i = 0 Then Exit Do
        j = InStr(i, Text, "]", vbBinaryCompare)
        If j = 0 Then Exit Do
        out = out & Mid$(Text, pos, i - pos)
        If UiEnglish() Then out = out & "qryLoc"
        out = out & Mid$(Text, i + 2, j - i - 2)
        pos = j + 1
    Loop
    LangSql = out & Mid$(Text, pos)
End Function

Public Function IsArabicCode(ByVal Code As Long) As Boolean
    ' Arabic letters and marks (tools/i18n.py LETTERS and MARKS)
    Select Case Code
        Case &H621 To &H65F, &H66E To &H66F, &H670, &H671 To &H6D3, &H6FA To &H6FC
            IsArabicCode = True
    End Select
End Function

Public Function HasArabic(ByVal Text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(Text)
        If IsArabicCode(AscW(Mid$(Text, i, 1))) Then
            HasArabic = True
            Exit Function
        End If
    Next
End Function

Private Function GluedAt(ByVal Text As String, ByVal Position As Long) As Boolean
    If Position >= 1 And Position <= Len(Text) Then GluedAt = IsArabicCode(AscW(Mid$(Text, Position, 1)))
End Function

Public Function ReplaceWord(ByVal Text As String, ByVal Key As String, ByVal Value As String) As String
    ' Replace Key where it is not glued to another Arabic letter (part of a longer word).
    Dim out As String, pos As Long, i As Long, n As Long
    pos = 1
    n = Len(Key)
    Do
        i = InStr(pos, Text, Key, vbBinaryCompare)
        If i = 0 Then Exit Do
        If GluedAt(Text, i - 1) Or GluedAt(Text, i + n) Then
            out = out & Mid$(Text, pos, i - pos + 1)
            pos = i + 1
        Else
            out = out & Mid$(Text, pos, i - pos) & Value
            pos = i + n
        End If
    Loop
    ReplaceWord = out & Mid$(Text, pos)
End Function

'------------------------------------------------------------------------------
' The dictionary: the longest Arabic phrases first
'------------------------------------------------------------------------------
Public Sub LangAdd(ByVal Arabic As String, ByVal English As String)
    ' Called by the data modules modLangData1, 2 ... (the order is kept: the longest phrases first)
    m_count = m_count + 1
    m_ar(m_count) = Arabic
    m_en(m_count) = English
End Sub

Private Sub LoadDictionary()
    ReDim m_ar(1 To ENTRY_COUNT)
    ReDim m_en(1 To ENTRY_COUNT)
    m_count = 0
    LangData1
    LangData2
    LangData3
    LangData4
    m_loaded = True
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): the expected texts come from tools/i18n.py
'------------------------------------------------------------------------------
Private Sub CheckLang(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestLang() As Boolean
    Dim passed As Long, failed As Long, report As String, saved As String
    saved = UiLanguage()
    Debug.Print "=== TestLang  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    UseLanguage "EN"
    CheckLang MSG_RTL = 0 And UiAlign(1) = 3 And UiAlign(3) = 1 And UiAlign(2) = 2, "English: left to right", _
              passed, failed, report
    CheckLang Tr(12) = 12 And Tr("ABC") = "ABC", "Tr leaves numbers and Latin text", passed, failed, report
    CheckLang Tr("=""æÑÞÉ ÌÑÏ "" & IIf([Status]=""POSTED"",""(ãõÑÍøá)"",IIf([Status]=""CANCELLED"",""(ãáÛì)"",""(ãÝÊæÍ)""))") = "=""Count sheet "" & IIf([Status]=""POSTED"",""(Posted)"",IIf([Status]=""CANCELLED"",""(Cancelled)"",""(Open)""))", "Tr sample 1", passed, failed, report
    CheckLang Tr("ÃõÖíÝ ") = "Added ", "Tr sample 2", passed, failed, report
    CheckLang Tr("ÇáÇÓã ÇáÅäÌáíÒí ÃØæá ãä ") = "The English name is longer than ", "Tr sample 3", passed, failed, report
    CheckLang Tr("ÇáßæÏ") = "Code", "Tr sample 4", passed, failed, report
    CheckLang Tr("ÊÚÐÑ ÇáÍÓÇÈ: ") = "Could not calculate: ", "Tr sample 5", passed, failed, report
    CheckLang Tr("ÓÌá ÇáÃÕæá æÞíÏ ÔÑÇÆåÇ¡ æÇáÅåáÇß ÈÇáÞÓØ ÇáËÇÈÊ¡ æÇáÈíÚ Ãæ ÇáÇÓÊÈÚÇÏ") = "Asset register and purchase entry, straight-line depreciation, and sale or disposal", "Tr sample 6", passed, failed, report
    CheckLang Tr("áÇ ÊæÌÏ ÝÊÑÉ ãÞÝáÉ") = "There is no closed period", "Tr sample 7", passed, failed, report
    CheckLang Tr("TEST æ 12.50: æíÉ") = "TEST and 12.50: æíÉ", "Tr sample 8", passed, failed, report
    UseLanguage "AR"
    CheckLang MSG_RTL = &H180000 And UiAlign(1) = 1 And Tr("ÇáÔåÇÏÉ ÇáÝÚáíÉ ÊõØáÈ ãä ÇáåíÆÉ Ýí ÇáãÑÍáÉ ÇáÊÇáíÉ (ÊÓÌíá ÇáÌåÇÒ). ááÊÌÑÈÉ ÇáÂä: «ÔåÇÏÉ ÊÌÑíÈíÉ» Ëã «ãáÝ XML áÝÇÊæÑÉ»¡ æÇÝÍÕ ÇáãáÝ ÈÃÏÇÉ ÇáåíÆÉ") = "ÇáÔåÇÏÉ ÇáÝÚáíÉ ÊõØáÈ ãä ÇáåíÆÉ Ýí ÇáãÑÍáÉ ÇáÊÇáíÉ (ÊÓÌíá ÇáÌåÇÒ). ááÊÌÑÈÉ ÇáÂä: «ÔåÇÏÉ ÊÌÑíÈíÉ» Ëã «ãáÝ XML áÝÇÊæÑÉ»¡ æÇÝÍÕ ÇáãáÝ ÈÃÏÇÉ ÇáåíÆÉ", "ÇáÚÑÈíÉ: ãä Çáíãíä æáÇ ÊÑÌãÉ", _
              passed, failed, report
    UseLanguage saved
    Debug.Print "--- passed: " & passed & " | failed: " & failed
    If failed = 0 Then
        TestMsg "ÌãíÚ ÇÎÊÈÇÑÇÊ ÇááÛÉ äÇÌÍÉ (" & passed & " ÇÎÊÈÇÑðÇ).", vbInformation + MSG_RTL, "TestLang"
        TestLang = True
    Else
        TestMsg "äÌÍ " & passed & " æÝÔá " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestLang"
    End If
End Function
