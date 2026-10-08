"""Generate modLang.bas: the interface language of a front-end file and the Arabic -> English
dictionary (tools/i18n_en.py). The algorithm is mirrored in tools/i18n.py (translate)."""

from generate_common import vba_str
import i18n

CHUNK = 120          # dictionary entries per procedure (VBA limits the size of one procedure)
MODULE_CHARS = 45000 # dictionary text per data module (modLangData1, 2 ...): small modules import and compile safely

TEMPLATE = r'''Attribute VB_Name = "modLang"
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
'   MSG_RTL     the right-to-left flags of MsgBox (0 in English).
' The tax invoice and the credit note stay bilingual in both languages (ZATCA).
'==============================================================================
Option Compare Database
Option Explicit

Private Const ENTRY_COUNT As Long = @@COUNT@@

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
    If Not UiEnglish() Then Exit Function
    s = Text
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
@@CALLS@@
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
@@SAMPLES@@
    UseLanguage "AR"
    CheckLang MSG_RTL = &H180000 And UiAlign(1) = 1 And Tr(@@AR1@@) = @@AR1@@, "العربية: من اليمين ولا ترجمة", _
              passed, failed, report
    UseLanguage saved
    Debug.Print "--- passed: " & passed & " | failed: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات اللغة ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestLang"
        TestLang = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestLang"
    End If
End Function
@@CHUNKS@@'''


def samples(dictionary) -> list:
    """A few interface strings and a message with data in it, for TestLang."""
    texts = sorted(set(i18n.strings()))
    picked = [texts[i * len(texts) // 8] for i in range(8)] if texts else []
    picked = [t for t in picked if len(t) < 300 and "\n" not in t]
    some = sorted(dictionary, key=len)[:1]
    if some:
        picked.append("TEST " + some[0] + " 12.50: " + some[0] + "ية")     # glued: not translated
    return picked


def build_lang_vba() -> str:
    from i18n_en import EN
    order = i18n.ordered(EN)
    lines = []
    for n, text in enumerate(samples(EN), 1):
        lines.append(f"    CheckLang Tr({vba_str(text)}) = {vba_str(i18n.translate(text, EN, order))}, "
                     f'"Tr sample {n}", passed, failed, report')
    first = vba_str(order[0]) if order else '"-"'
    modules = data_modules(EN, order)
    calls = [f"    LangData{n}" for n in range(1, len(modules) + 1)]
    return (TEMPLATE.replace("@@COUNT@@", str(max(1, len(order))))
            .replace("@@SAMPLES@@", "\n".join(lines)).replace("@@AR1@@", first)
            .replace("@@CALLS@@", "\n".join(calls) if calls else "")
            .replace("@@CHUNKS@@", ""))


DATA_TEMPLATE = """Attribute VB_Name = "modLangData{n}"
'==============================================================================
' modLangData{n}  -  Retail Store Management System: part {n} of the Arabic ->
' English dictionary of the interface (modLang). GENERATED FILE - do not edit.
' Source: tools/i18n_en.py  ->  python3 tools/generate.py
'==============================================================================
Option Compare Database
Option Explicit

Public Sub LangData{n}()
{calls}
End Sub
{procs}"""


def data_modules(dictionary=None, order=None) -> list:
    """The dictionary in modules of about MODULE_CHARS characters, the longest phrases first."""
    if dictionary is None:
        from i18n_en import EN as dictionary
    order = order or i18n.ordered(dictionary)
    groups, cur, size = [], [], 0
    for k in order:
        line = f"    LangAdd {vba_str(k)}, {vba_str(dictionary[k])}"
        if cur and size + len(line) > MODULE_CHARS:
            groups.append(cur)
            cur, size = [], 0
        cur.append(line)
        size += len(line) + 1
    if cur:
        groups.append(cur)
    out = []
    for n, lines in enumerate(groups, 1):
        procs, calls = [], []
        for p, start in enumerate(range(0, len(lines), CHUNK), 1):
            procs.append(f"\nPrivate Sub D{n}_{p}()\n" + "\n".join(lines[start:start + CHUNK]) + "\nEnd Sub\n")
            calls.append(f"    D{n}_{p}")
        out.append(DATA_TEMPLATE.format(n=n, calls="\n".join(calls), procs="".join(procs)))
    return out
