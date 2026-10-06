Attribute VB_Name = "modAccounts"
'==============================================================================
' modAccounts  -  Retail Store Management System
'
' The chart of accounts as a tree (table Accounts, screen frmAccounts):
'   level 1  the five classes: assets, liabilities, equity, revenue, expenses
'   level 2  groups (current assets, current liabilities, operating expenses...)
'   level 3  accounts; level 4-5 sub-accounts (each cash box, each expense type)
' A main account (IsPosting = False) only totals its children; entries go to
' the sub-accounts (IsPosting = True). System accounts (IsSystem) are used by the
' automatic entries and keep their code, type and kind.
'   RebuildAccountTree   AccountLevel, TreeKey (order) and Level1Code..Level5Code
'                        (the account of each level above, for the totals by level)
'   ValidateAccount      rules of the accounts screen (called by modForms)
'   AccountDeleteProblem why an account cannot be deleted ("" = it can)
' Same steps as tools/sim.py (Store.rebuild_account_tree).
'==============================================================================
Option Compare Database
Option Explicit

Public Const MAX_ACCOUNT_LEVELS As Long = 5

'------------------------------------------------------------------------------
' Tree
'------------------------------------------------------------------------------
Public Function RebuildAccountTree() As String
    ' "" on success, else the accounts that cannot be placed (a loop, or more than 5 levels).
    Dim db As DAO.Database, rs As DAO.Recordset, parents As Object
    Dim chain(1 To 6) As Long, code As Long, n As Long, k As Long, key As String, problem As String
    Set db = CurrentDb
    Set parents = CreateObject("Scripting.Dictionary")
    Set rs = db.OpenRecordset("SELECT AccountCode, ParentCode FROM Accounts", dbOpenSnapshot)
    Do Until rs.EOF
        parents(CLng(rs!AccountCode)) = CLng(Nz(rs!ParentCode, 0))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT AccountCode, AccountLevel, TreeKey, Level1Code, Level2Code, Level3Code, " & _
                              "Level4Code, Level5Code FROM Accounts", dbOpenDynaset)
    Do Until rs.EOF
        ' chain(1) = the account, chain(n) = its level-1 account (a parent that does not exist ends it)
        code = rs!AccountCode
        n = 0
        Do While code <> 0 And n < MAX_ACCOUNT_LEVELS + 1
            n = n + 1
            chain(n) = code
            code = parents(code)
            If code <> 0 And Not parents.Exists(code) Then code = 0
        Loop
        If n > MAX_ACCOUNT_LEVELS Then
            problem = problem & "  " & rs!AccountCode & vbCrLf
        Else
            key = ""
            For k = 1 To n
                key = key & Format$(chain(n - k + 1), "0000000000")
            Next
            If Nz(rs!TreeKey, "") <> key Or Nz(rs!AccountLevel, 0) <> n Or Nz(rs!Level1Code, 0) <> chain(n) Then
                rs.Edit
                rs!AccountLevel = n
                rs!TreeKey = key
                For k = 1 To MAX_ACCOUNT_LEVELS
                    If k <= n Then rs("Level" & k & "Code") = chain(n - k + 1) Else rs("Level" & k & "Code") = Null
                Next
                rs.Update
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(problem) > 0 Then
        RebuildAccountTree = "حسابات لا يمكن وضعها في الشجرة (حلقة أو أكثر من " & MAX_ACCOUNT_LEVELS & _
                             " مستويات):" & vbCrLf & problem
    End If
End Function

Public Function AccountHasEntries(ByVal AccountCode As Long) As Boolean
    AccountHasEntries = Nz(DbValue("SELECT COUNT(*) FROM JournalLines WHERE AccountCode = " & AccountCode), 0) + _
                        Nz(DbValue("SELECT COUNT(*) FROM ManualEntryLines WHERE AccountCode = " & AccountCode), 0) > 0
End Function

Public Function AccountHasChildren(ByVal AccountCode As Long) As Boolean
    AccountHasChildren = Nz(DbValue("SELECT COUNT(*) FROM Accounts WHERE ParentCode = " & AccountCode), 0) > 0
End Function

'------------------------------------------------------------------------------
' Accounts screen (frmAccounts, a data screen on Accounts)
'------------------------------------------------------------------------------
Public Sub AccountParentChanged(ByVal frm As Access.Form)
    ' A sub-account has the type of its main account.
    Dim kind As Variant
    If IsNull(frm!ParentCode.Value) Then Exit Sub
    kind = DbValue("SELECT AccountType FROM Accounts WHERE AccountCode = " & frm!ParentCode.Value)
    If Not IsNull(kind) Then frm!AccountType.Value = kind
End Sub

Public Function ValidateAccount(ByVal frm As Access.Form) As Boolean
    Dim code As Long, parent As Variant, ownKey As String, parentKey As String, depth As Long
    code = Nz(frm!AccountCode.Value, 0)
    parent = frm!ParentCode.Value
    If Not frm.NewRecord And code <> Nz(frm!AccountCode.OldValue, code) Then
        ShowWarning "لا يتغير رقم حساب محفوظ. أنشئ حسابًا جديدًا وعطّل القديم."
        Exit Function
    End If
    If Not frm.NewRecord Then ownKey = Nz(DbValue("SELECT TreeKey FROM Accounts WHERE AccountCode = " & code), "")

    If IsNull(parent) Then
        If code >= 10 Then
            ShowWarning "اختر الحساب الرئيسي الذي يتبعه هذا الحساب."
            SafeFocus frm!ParentCode
            Exit Function
        End If
    Else
        If parent = code Then
            ShowWarning "لا يتبع الحساب نفسه."
            Exit Function
        End If
        If IsNull(DbValue("SELECT AccountCode FROM Accounts WHERE AccountCode = " & parent)) Then
            ShowWarning "الحساب الرئيسي غير موجود."
            Exit Function
        End If
        If DbValue("SELECT IsPosting FROM Accounts WHERE AccountCode = " & parent) Then
            ShowWarning "الحساب الرئيسي يجب أن يكون حسابًا تجميعيًا (لا يقبل القيود)." & vbCrLf & _
                        "إذا لم تكن عليه قيود ألغِ «حساب فرعي» فيه أولًا."
            SafeFocus frm!ParentCode
            Exit Function
        End If
        parentKey = Nz(DbValue("SELECT TreeKey FROM Accounts WHERE AccountCode = " & parent), "")
        If Len(ownKey) > 0 And Left$(parentKey, Len(ownKey)) = ownKey Then
            ShowWarning "لا يتبع الحساب حسابًا من الحسابات التابعة له."
            Exit Function
        End If
        ' the account and everything under it must stay within the levels of the tree
        depth = 1
        If Len(ownKey) > 0 Then
            depth = Nz(DbValue("SELECT Max(AccountLevel) FROM Accounts WHERE Left(TreeKey, " & Len(ownKey) & ") = " & _
                               SqlText(ownKey)), 0) - Len(ownKey) \ 10 + 1
            If depth < 1 Then depth = 1
        End If
        If Nz(DbValue("SELECT AccountLevel FROM Accounts WHERE AccountCode = " & parent), 0) + depth > MAX_ACCOUNT_LEVELS Then
            ShowWarning "شجرة الحسابات " & MAX_ACCOUNT_LEVELS & " مستويات على الأكثر."
            Exit Function
        End If
        AccountParentChanged frm
    End If

    If Nz(frm!IsSystem.Value, False) And Not frm.NewRecord Then
        If frm!IsPosting.Value <> frm!IsPosting.OldValue Or frm!AccountType.Value <> frm!AccountType.OldValue Then
            ShowWarning "هذا حساب أساسي تستخدمه القيود الآلية: لا يتغير نوعه، ولا يتحول بين رئيسي وفرعي."
            Exit Function
        End If
    End If
    If Not frm.NewRecord Then
        If Not frm!IsPosting.Value And AccountHasEntries(code) Then
            ShowWarning "على هذا الحساب قيود، فيبقى حسابًا فرعيًا."
            Exit Function
        End If
        If frm!IsPosting.Value And AccountHasChildren(code) Then
            ShowWarning "لهذا الحساب حسابات تابعة، فيبقى حسابًا رئيسيًا."
            Exit Function
        End If
    End If
    ValidateAccount = True
End Function

Public Function AccountDeleteProblem(ByVal AccountCode As Long) As String
    If AccountHasChildren(AccountCode) Then
        AccountDeleteProblem = "لهذا الحساب حسابات تابعة. احذفها أو انقلها أولًا."
    End If
End Function
