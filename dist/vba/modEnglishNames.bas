Attribute VB_Name = "modEnglishNames"
'==============================================================================
' modEnglishNames  -  Retail Store Management System (the English names screen)
'
' frmEnglishNames (docs/42-English-Names-Screen.md): the names of customers, suppliers, products, boxes,
' banks, accounts ... with their English name, by default only those that have none. "Suggest" fills the
' empty English names with Transliterate (common words, first names and cities in their usual spelling,
' any other word letter by letter); the user reviews them and saves. tools/translit.py is the Python mirror
' of Transliterate; tools/master_en.names_screen_const gives NAME_TABLES.
'==============================================================================
Option Compare Database
Option Explicit

' table,key field,Arabic name field,English name field,English field size,caption
Private Const NAME_TABLES As String = "Customers,CustomerID,CustomerName,CustomerNameEn,150,«·⁄„·«¡;Suppliers,SupplierID,SupplierName,SupplierNameEn,150,«·„Ê—œÊ‰;Products,ProductID,ProductName,ProductNameEn,150,«·„‰ Ã« ;Categories,CategoryID,CategoryName,CategoryNameEn,100,«· ’‰Ì›« ;Units,UnitID,UnitName,UnitNameEn,30,ÊÕœ«  «·ﬁÌ«”;" & _
    "CashBoxes,CashBoxID,BoxName,BoxNameEn,50,«·Œ“Ì‰… Ê«·’‰«œÌﬁ;Banks,BankID,BankName,BankNameEn,100,«·»‰Êﬂ;CostCenters,CostCenterID,CenterName,CenterNameEn,100,„—«ﬂ“ «· ﬂ·›… Ê«·›—Ê⁄;SalesReps,SalesRepID,RepName,RepNameEn,100,«·„‰œÊ»Ì‰;" & _
    "ExpenseTypes,ExpenseTypeID,ExpenseTypeName,ExpenseTypeNameEn,50,√‰Ê«⁄ «·„’—Ê›« ;Accounts,AccountCode,AccountName,AccountNameEn,100,œ·Ì· «·Õ”«»«  (‘Ã—… «·Õ”«»« );PaymentMethods,PaymentMethodID,MethodName,MethodNameEn,50,ÿ—ﬁ «·œ›⁄;Roles,RoleID,RoleName,RoleNameEn,50,«·√œÊ«—"

' Arabic letter = Latin letters, and whole words = their usual English (tools/translit.py)
Private Const TRANSLIT_LETTERS As String = "¡=|¬=a|√=a|ƒ=o|≈=i|∆=e|«=a|»=b|…=a| =t|À=th|Ã=j|Õ=h|Œ=kh|œ=d|–=th|—=r|“=z|”=s|‘=sh|’=s|÷=d|ÿ=t|Ÿ=z|⁄=a|€=gh|›=f|ﬁ=q|ﬂ=k|·=l|„=m|‰=n|Â=h|Ê=o|Ï=a|Ì=i"
Private Const TRANSLIT_WORDS As String = "‘—ﬂ…=Company|‘—ﬂÂ=Company|„ƒ””…=Establishment|„ƒ””Â=Establishment|„ﬂ »=Office|„ÿ⁄„=Restaurant|„ﬁÂÏ=Cafe|ﬂ«›ÌÂ=Cafe|„Õ·=Shop|„Õ·« =Shops|„ Ã—=Store|”Êﬁ=Market|√”Ê«ﬁ=Markets|„’‰⁄=Factory|„Ã„Ê⁄…=Group| Ã«—…=Trading|·· Ã«—…=Trading|«· Ã«—Ì…=Trading|„ﬁ«Ê·« =Contracting|··„ﬁ«Ê·« =Contracting|" & _
    "Ê‘—ﬂ«Â=and Partners|Ê=and|›—⁄=Branch|«·—∆Ì”Ì=Main|«·—∆Ì”Ì…=Main|«·—∆Ì”ÌÂ=Main|’‰œÊﬁ=Box|Œ“Ì‰…=Treasury|«·Œ“Ì‰…=Treasury|»‰ﬂ=Bank|„’—›=Bank|Õ”«»=Account|Ã«—Ì=Current| Ê›Ì—=Savings|ﬁ”„=Department|≈œ«—…=Administration|«·„»Ì⁄« =Sales|«·„‘ —Ì« =Purchases|«·„” Êœ⁄=Warehouse|„” Êœ⁄=Warehouse|⁄„Ì·=Customer|" & _
    "„Ê—œ=Supplier|‰ﬁœÌ=Cash|ﬂ«‘Ì—=Cashier|»‰=bin|«»‰=bin|»‰ =bint|√»Ê=Abu|«»Ê=Abu|¬·=Al|„Õ„œ=Mohammed|√Õ„œ=Ahmed|«Õ„œ=Ahmed|„Õ„Êœ=Mahmoud|⁄·Ì=Ali|⁄„—=Omar|Œ«·œ=Khalid|”⁄œ=Saad|”⁄Ìœ=Saeed|›Âœ=Fahad|”·ÿ«‰=Sultan|‰«’—=Nasser|”⁄Êœ=Saud|›Ì’·=Faisal|ÌÊ”›=Yousef|≈»—«ÂÌ„=Ibrahim|«»—«ÂÌ„=Ibrahim|’«·Õ=Saleh|" & _
    "Õ”‰=Hassan|Õ”Ì‰=Hussein|”·Ì„«‰=Sulaiman|„‰’Ê—=Mansour|„«Ãœ=Majed| —ﬂÌ=Turki|»‰œ—=Bandar|‰«Ì›=Naif|„‘⁄·=Mishal|⁄«œ·=Adel|ÿ«—ﬁ=Tariq|Ì«”—=Yasser|Â‘«„=Hisham|„’ÿ›Ï=Mustafa|⁄»œ«··Â=Abdullah|⁄»œ«·⁄“Ì“=Abdulaziz|⁄»œ«·—Õ„‰=Abdulrahman|”«—…=Sarah|”«—Â=Sarah|›«ÿ„…=Fatimah|‰Ê—…=Noura|‰Ê—Â=Noura|„—Ì„=Maryam|" & _
    "⁄«∆‘…=Aisha|Â‰œ=Hind|«·—Ì«÷=Riyadh|Ãœ…=Jeddah|ÃœÂ=Jeddah|„ﬂ…=Makkah|«·„œÌ‰…=Madinah|«·œ„«„=Dammam|«·Œ»—=Khobar|«·ÿ«∆›=Taif| »Êﬂ=Tabuk|√»Â«=Abha|«·ﬁ’Ì„=Qassim"

Private m_letters() As String, m_lettersEn() As String, m_words() As String, m_wordsEn() As String
Private m_loaded As Boolean

Private Function ArabicAl() As String
    ArabicAl = ChrW(&H627) & ChrW(&H644)                       ' the article al-
End Function

'==============================================================================
' Transliteration
'==============================================================================
Public Function Transliterate(ByVal Text As String) As String
    ' A suggested English name for an Arabic name: "„ƒ””… «·‰Ê— ·· Ã«—…" -> "Establishment Al-Nor Trading".
    Dim i As Long, ch As String, code As Long, word As String, out As String
    If Not m_loaded Then LoadTranslit
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        code = AscW(ch)
        If (code >= &H64B And code <= &H652) Or code = &H640 Then
            ch = ""                                         ' diacritics and tatweel
        ElseIf code >= &H660 And code <= &H669 Then
            ch = ChrW(code - &H660 + 48)                    ' Arabic digits
        ElseIf code = &H60C Or code = &H61B Then
            ch = ","
        ElseIf code = &H61F Then
            ch = "?"
        End If
        If Len(ch) > 0 Then
            If LetterIndex(ch) >= 0 Then
                word = word & ch
            Else
                If Len(word) > 0 Then out = out & WordEn(word)
                word = ""
                out = out & ch
            End If
        End If
    Next
    If Len(word) > 0 Then out = out & WordEn(word)
    Do While InStr(1, out, "  ", vbBinaryCompare) > 0
        out = Replace(out, "  ", " ")
    Loop
    Transliterate = Trim$(out)
End Function

Private Function WordEn(ByVal Word As String) As String
    Dim i As Long, rest As String
    i = WordIndex(Word)
    If i >= 0 Then
        WordEn = m_wordsEn(i)
    ElseIf Len(Word) > 3 And StrComp(Left$(Word, 2), ArabicAl(), vbBinaryCompare) = 0 Then
        rest = Mid$(Word, 3)
        i = WordIndex(rest)
        If i >= 0 Then
            WordEn = "Al-" & m_wordsEn(i)
        Else
            WordEn = "Al-" & Capital(LettersEn(rest))
        End If
    ElseIf Len(Word) > 3 And StrComp(Left$(Word, 3), ChrW(&H639) & ChrW(&H628) & ChrW(&H62F), vbBinaryCompare) = 0 Then
        rest = Mid$(Word, 4)
        If Len(rest) > 2 And StrComp(Left$(rest, 2), ArabicAl(), vbBinaryCompare) = 0 Then rest = Mid$(rest, 3)
        WordEn = "Abdul" & LettersEn(rest)
    Else
        WordEn = Capital(LettersEn(Word))
    End If
End Function

Private Function LettersEn(ByVal Word As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(Word)
        ch = Mid$(Word, i, 1)
        If i = 1 And StrComp(ch, ChrW(&H648), vbBinaryCompare) = 0 Then          ' waw and ya at the start
            out = out & "w"
        ElseIf i = 1 And StrComp(ch, ChrW(&H64A), vbBinaryCompare) = 0 Then
            out = out & "y"
        Else
            out = out & m_lettersEn(LetterIndex(ch))
        End If
    Next
    Do While InStr(1, out, "aa", vbBinaryCompare) > 0
        out = Replace(out, "aa", "a")
    Loop
    LettersEn = out
End Function

Private Function Capital(ByVal Word As String) As String
    Capital = UCase$(Left$(Word, 1)) & Mid$(Word, 2)
End Function

Private Function LetterIndex(ByVal ch As String) As Long
    Dim i As Long
    LetterIndex = -1
    For i = 0 To UBound(m_letters)
        If StrComp(m_letters(i), ch, vbBinaryCompare) = 0 Then
            LetterIndex = i
            Exit Function
        End If
    Next
End Function

Private Function WordIndex(ByVal Word As String) As Long
    Dim i As Long
    WordIndex = -1
    For i = 0 To UBound(m_words)
        If StrComp(m_words(i), Word, vbBinaryCompare) = 0 Then
            WordIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub LoadTranslit()
    Dim items() As String, pair() As String, i As Long
    items = Split(TRANSLIT_LETTERS, "|")
    ReDim m_letters(UBound(items)): ReDim m_lettersEn(UBound(items))
    For i = 0 To UBound(items)
        pair = Split(items(i), "=")
        m_letters(i) = pair(0)
        If UBound(pair) > 0 Then m_lettersEn(i) = pair(1)
    Next
    items = Split(TRANSLIT_WORDS, "|")
    ReDim m_words(UBound(items)): ReDim m_wordsEn(UBound(items))
    For i = 0 To UBound(items)
        pair = Split(items(i), "=")
        m_words(i) = pair(0)
        m_wordsEn(i) = pair(1)
    Next
    m_loaded = True
End Sub

'==============================================================================
' frmEnglishNames
'==============================================================================
Public Sub EnglishNamesLoad(ByVal frm As Access.Form)
    Dim spec As Variant, items As String, p() As String
    EnsureLocalTables
    items = "ALL;ﬂ· «·Ãœ«Ê·"
    For Each spec In Split(NAME_TABLES, ";")
        p = Split(spec, ",")
        items = items & ";" & p(0) & ";" & p(5)
    Next
    frm!cboTable.RowSource = Tr(items)
    frm!cboTable.Value = "ALL"
    frm!chkMissing.Value = True
    EnglishNamesShow frm
End Sub

Public Sub EnglishNamesShow(ByVal frm As Access.Form)
    ' The names of the chosen table (or of all), only those without an English name when chkMissing is ticked.
    Dim spec As Variant, p() As String, sql As String, chosen As String, n As Long
    If frm!subNames.Form.Dirty Then frm!subNames.Form.Dirty = False
    chosen = Nz(frm!cboTable.Value, "ALL")
    CurrentDb.Execute "DELETE FROM tmpEnglishNames", dbFailOnError
    For Each spec In Split(NAME_TABLES, ";")
        p = Split(spec, ",")
        If chosen = "ALL" Or chosen = p(0) Then
            sql = "INSERT INTO tmpEnglishNames (TableName, TableTitle, KeyValue, ArabicName, EnglishName, OldEnglish, " & _
                  "MaxLen) SELECT '" & p(0) & "', '" & p(5) & "', t." & p(1) & ", t." & p(2) & ", t." & p(3) & ", t." & _
                  p(3) & ", " & p(4) & " FROM " & p(0) & " AS t"
            If Nz(frm!chkMissing.Value, False) Then sql = sql & " WHERE t." & p(3) & " Is Null"
            CurrentDb.Execute Tr(sql & " ORDER BY t." & p(2)), dbFailOnError     ' the caption in the interface language
        End If
    Next
    frm!subNames.Form.Requery
    n = DCount("*", "tmpEnglishNames")
    frm!lblCount.Caption = Tr(n & " ”Ã·")
End Sub

Public Sub EnglishNamesSuggest(ByVal frm As Access.Form)
    ' Every empty English name of the list gets its transliteration; nothing is saved yet.
    Dim rs As DAO.Recordset, s As String, n As Long
    If frm!subNames.Form.Dirty Then frm!subNames.Form.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT ArabicName, EnglishName, MaxLen FROM tmpEnglishNames " & _
                                     "WHERE EnglishName Is Null", dbOpenDynaset)
    Do Until rs.EOF
        s = Trim$(Left$(Transliterate(Nz(rs!ArabicName, "")), rs!MaxLen))
        If Len(s) > 0 Then
            rs.Edit
            rs!EnglishName = s
            rs.Update
            n = n + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    frm!subNames.Form.Requery
    ShowInfo "«ﬁ —«Õ«  ÃœÌœ…: " & n & vbCrLf & "—«Ã⁄Â« Ê⁄œ¯· „« Ì·“„ À„ «÷€ÿ Õ›Ÿ."
End Sub

Public Function SaveEnglishNames(ByVal frm As Access.Form) As Boolean
    Dim msg As String, n As Long
    If Not CanScreenAction(frm.Name, "EDIT") Then Exit Function
    If frm!subNames.Form.Dirty Then frm!subNames.Form.Dirty = False
    msg = SaveEnglishNameRows(n)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Function
    End If
    ShowInfo " „ Õ›Ÿ «·√”„«¡ «·≈‰Ã·Ì“Ì…: " & n
    EnglishNamesShow frm
    SaveEnglishNames = True
End Function

Public Function SaveEnglishNameRows(ByRef Saved As Long) As String
    ' Writes the English names changed in tmpEnglishNames ("" = success). Nothing is written when a name is
    ' longer than its field. Each change is in the audit trail.
    Dim db As DAO.Database, rs As DAO.Recordset, p As Variant, newName As String, before As Collection
    Saved = 0
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT EnglishName, MaxLen FROM tmpEnglishNames", dbOpenSnapshot)
    Do Until rs.EOF
        newName = Trim$(Nz(rs!EnglishName, ""))
        If Len(newName) > rs!MaxLen Then
            SaveEnglishNameRows = "«·«”„ «·≈‰Ã·Ì“Ì √ÿÊ· „‰ " & rs!MaxLen & " Õ—›«: " & newName
            rs.Close
            Exit Function
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT * FROM tmpEnglishNames ORDER BY LineNo", dbOpenDynaset)
    Do Until rs.EOF
        newName = Trim$(Nz(rs!EnglishName, ""))
        If StrComp(newName, Nz(rs!OldEnglish, ""), vbBinaryCompare) <> 0 Then
            p = TableSpec(rs!TableName)
            Set before = AuditSnapshot(p(0), p(1), rs!KeyValue)
            db.Execute "UPDATE [" & p(0) & "] SET [" & p(3) & "] = " & SqlText(newName) & " WHERE [" & p(1) & "] = " & _
                       rs!KeyValue, dbFailOnError
            AuditEdited "EDIT", p(0), p(1), rs!KeyValue, before
            rs.Edit
            If Len(newName) > 0 Then
                rs!EnglishName = newName
                rs!OldEnglish = newName
            Else
                rs!EnglishName = Null
                rs!OldEnglish = Null
            End If
            rs.Update
            Saved = Saved + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

Private Function TableSpec(ByVal TableName As String) As Variant
    Dim spec As Variant, p() As String
    For Each spec In Split(NAME_TABLES, ";")
        p = Split(spec, ",")
        If p(0) = TableName Then
            TableSpec = p
            Exit Function
        End If
    Next
End Function

'==============================================================================
' In-Access test (changes rolled back)
'==============================================================================
Public Function TestEnglishNames() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, passed As Long, failed As Long, report As String, msg As String
    Dim id As Long, n As Long
    Calendar = vbCalGreg
    EnsureTestUser
    EnsureLocalTables
    g_SilentMode = True
    Debug.Print "=== TestEnglishNames  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    Record Transliterate("„Õ„œ ⁄»œ«··Â «·€«„œÌ") = "Mohammed Abdullah Al-Ghamdi", "«ﬁ —«Õ «”„ ‘Œ’", passed, failed, report
    Record Transliterate("„ƒ””… «·‰Ê— ·· Ã«—… " & ChrW(&H662)) = "Establishment Al-Nor Trading 2", "«ﬁ —«Õ «”„ „‰‘√…", _
           passed, failed, report
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    CurrentDb.Execute "INSERT INTO Customers (CustomerName) VALUES ('TEST ⁄„Ì· «·√”„«¡')", dbFailOnError
    id = DbValue("SELECT CustomerID FROM Customers WHERE CustomerName = 'TEST ⁄„Ì· «·√”„«¡'")
    CurrentDb.Execute "DELETE FROM tmpEnglishNames", dbFailOnError
    CurrentDb.Execute "INSERT INTO tmpEnglishNames (TableName, TableTitle, KeyValue, ArabicName, EnglishName, MaxLen) " & _
                      "VALUES ('Customers', 'TEST', " & id & ", 'TEST ⁄„Ì· «·√”„«¡', 'TEST names customer', 150)", dbFailOnError
    msg = SaveEnglishNameRows(n)
    Record Len(msg) = 0 And n = 1 And Nz(DbValue("SELECT CustomerNameEn FROM Customers WHERE CustomerID = " & id), "") = _
           "TEST names customer", "Õ›Ÿ «·«”„ «·≈‰Ã·Ì“Ì " & msg, passed, failed, report
    msg = SaveEnglishNameRows(n)
    Record Len(msg) = 0 And n = 0, "«·Õ›Ÿ «·À«‰Ì ·« Ì€Ì¯— ‘Ì∆«", passed, failed, report
    CurrentDb.Execute "UPDATE tmpEnglishNames SET EnglishName = '" & String$(151, "a") & "'", dbFailOnError
    msg = SaveEnglishNameRows(n)
    Record Len(msg) > 0 And n = 0 And Nz(DbValue("SELECT CustomerNameEn FROM Customers WHERE CustomerID = " & id), "") = _
           "TEST names customer", "«·«”„ «·√ÿÊ· „‰ «·Õﬁ· ·« ÌıÕ›Ÿ", passed, failed, report
    ws.Rollback
    inTrans = False
    CurrentDb.Execute "DELETE FROM tmpEnglishNames", dbFailOnError
    GoTo Done
EH:
    Record False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «·√”„«¡ «·≈‰Ã·Ì“Ì… ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestEnglishNames"
        TestEnglishNames = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestEnglishNames"
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
