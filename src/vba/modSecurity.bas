Attribute VB_Name = "modSecurity"
'==============================================================================
' modSecurity  -  Retail Store Management System (Phase 10)
'
' Users, passwords and permissions.
'   Sha256Hex / PasswordHash   pure VBA SHA-256 (no Windows library needed);
'                              same results as tools/security_reference.py.
'   LoginUser                  checks a user name + password, counts failures and
'                              locks the account for LOCK_MINUTES after MAX_FAILED.
'   SetUserPassword            stores a new salt + hash (never the password itself).
'   ChangeOwnPassword          for the logged-in user.
'   CanOpenScreen              may the user open a screen: the programmer always; a user with his
'                              own screen permissions (frmUserScreens): table UserScreens;
'                              otherwise the role permission of the screen (map in modAppData).
'   CanScreenAction            add / edit / delete in a screen (only limited for a user with his
'                              own screen permissions).
'   IsDeveloper                the programmer: above the administrator, never limited, hidden
'                              from the users screen, decides if the administrator may rename the shop.
'   ValidateEmployee           rules of the users screen (called by modForms).
' The logged-in user is TempVars("UserID"); 0 / missing = nobody.
'==============================================================================
Option Compare Database
Option Explicit

Public Const PASSWORD_ITERATIONS As Long = 100
Public Const MIN_PASSWORD_LENGTH As Long = 6
Public Const MAX_FAILED_LOGINS As Long = 5
Public Const LOCK_MINUTES As Long = 15
Public Const ADMIN_ROLE_ID As Long = 1

Private Const K_HEX As String = "428a2f98 71374491 b5c0fbcf e9b5dba5 3956c25b 59f111f1 923f82a4 ab1c5ed5 " & _
    "d807aa98 12835b01 243185be 550c7dc3 72be5d74 80deb1fe 9bdc06a7 c19bf174 e49b69c1 efbe4786 0fc19dc6 " & _
    "240ca1cc 2de92c6f 4a7484aa 5cb0a9dc 76f988da 983e5152 a831c66d b00327c8 bf597fc7 c6e00bf3 d5a79147 " & _
    "06ca6351 14292967 27b70a85 2e1b2138 4d2c6dfc 53380d13 650a7354 766a0abb 81c2c92e 92722c85 a2bfe8a1 " & _
    "a81a664b c24b8b70 c76c51a3 d192e819 d6990624 f40e3585 106aa070 19a4c116 1e376c08 2748774c 34b0bcb5 " & _
    "391c0cb3 4ed8aa4a 5b9cca4f 682e6ff3 748f82ee 78a5636f 84c87814 8cc70208 90befffa a4506ceb bef9a3f7 " & _
    "c67178f2"
Private Const H0_HEX As String = "6a09e667 bb67ae85 3c6ef372 a54ff53a 510e527f 9b05688c 1f83d9ab 5be0cd19"

Private m_pow2(0 To 30) As Long
Private m_k(0 To 63) As Long
Private m_h0(0 To 7) As Long
Private m_ready As Boolean

'==============================================================================
' SHA-256 (FIPS 180-4) on 32-bit words held in signed Longs
'==============================================================================
Public Function Sha256Hex(ByRef Data() As Byte, ByVal n As Long) As String
    Dim h(0 To 7) As Long, i As Long, s As String
    Sha256Words Data, n, h
    For i = 0 To 7
        s = s & Right$("0000000" & LCase$(Hex$(h(i))), 8)
    Next
    Sha256Hex = s
End Function

Public Function Sha256Text(ByVal Text As String) As String
    Dim b() As Byte, n As Long
    n = Utf8Bytes(Text, b)
    Sha256Text = Sha256Hex(b, n)
End Function

Public Function PasswordHash(ByVal Password As String, ByVal Salt As String) As String
    ' hex( SHA-256 applied PASSWORD_ITERATIONS times to UTF-8(Salt & Password) )
    Dim b() As Byte, n As Long, h(0 To 7) As Long, i As Long, j As Long, s As String
    n = Utf8Bytes(Salt & Password, b)
    For i = 1 To PASSWORD_ITERATIONS
        Sha256Words b, n, h
        ReDim b(0 To 31)
        For j = 0 To 7
            b(j * 4) = HighWord(h(j)) \ 256
            b(j * 4 + 1) = HighWord(h(j)) And 255
            b(j * 4 + 2) = (h(j) And &HFFFF&) \ 256
            b(j * 4 + 3) = h(j) And 255
        Next
        n = 32
    Next
    For j = 0 To 7
        s = s & Right$("0000000" & LCase$(Hex$(h(j))), 8)
    Next
    PasswordHash = s
End Function

Private Sub Sha256Words(ByRef Data() As Byte, ByVal n As Long, ByRef h() As Long)
    Dim total As Long, msg() As Byte, i As Long, blk As Long, t As Long
    Dim w(0 To 63) As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long
    Dim g As Long, hh As Long, t1 As Long, t2 As Long, s0 As Long, s1 As Long, bitLen As Long
    InitSha
    total = ((n + 8) \ 64 + 1) * 64
    ReDim msg(0 To total - 1)
    For i = 0 To n - 1
        msg(i) = Data(i)
    Next
    msg(n) = &H80
    bitLen = n * 8
    msg(total - 4) = HighWord(bitLen) \ 256
    msg(total - 3) = HighWord(bitLen) And 255
    msg(total - 2) = (bitLen And &HFFFF&) \ 256
    msg(total - 1) = bitLen And 255
    For i = 0 To 7
        h(i) = m_h0(i)
    Next
    For blk = 0 To total - 1 Step 64
        For t = 0 To 15
            i = blk + t * 4
            w(t) = Combine16(CLng(msg(i)) * 256 + msg(i + 1), CLng(msg(i + 2)) * 256 + msg(i + 3))
        Next
        For t = 16 To 63
            s0 = RotR(w(t - 15), 7) Xor RotR(w(t - 15), 18) Xor ShiftRight(w(t - 15), 3)
            s1 = RotR(w(t - 2), 17) Xor RotR(w(t - 2), 19) Xor ShiftRight(w(t - 2), 10)
            w(t) = U32Add(U32Add(w(t - 16), s0), U32Add(w(t - 7), s1))
        Next
        a = h(0): b = h(1): c = h(2): d = h(3): e = h(4): f = h(5): g = h(6): hh = h(7)
        For t = 0 To 63
            s1 = RotR(e, 6) Xor RotR(e, 11) Xor RotR(e, 25)
            t1 = U32Add(U32Add(hh, s1), U32Add(U32Add((e And f) Xor ((Not e) And g), m_k(t)), w(t)))
            s0 = RotR(a, 2) Xor RotR(a, 13) Xor RotR(a, 22)
            t2 = U32Add(s0, (a And b) Xor (a And c) Xor (b And c))
            hh = g: g = f: f = e
            e = U32Add(d, t1)
            d = c: c = b: b = a
            a = U32Add(t1, t2)
        Next
        h(0) = U32Add(h(0), a): h(1) = U32Add(h(1), b): h(2) = U32Add(h(2), c): h(3) = U32Add(h(3), d)
        h(4) = U32Add(h(4), e): h(5) = U32Add(h(5), f): h(6) = U32Add(h(6), g): h(7) = U32Add(h(7), hh)
    Next
End Sub

Private Sub InitSha()
    Dim i As Long, parts() As String
    If m_ready Then Exit Sub
    m_pow2(0) = 1
    For i = 1 To 30
        m_pow2(i) = m_pow2(i - 1) * 2
    Next
    parts = Split(K_HEX, " ")
    For i = 0 To 63
        m_k(i) = HexToLong(parts(i))
    Next
    parts = Split(H0_HEX, " ")
    For i = 0 To 7
        m_h0(i) = HexToLong(parts(i))
    Next
    m_ready = True
End Sub

Private Function HexToLong(ByVal s As String) As Long
    Dim i As Long, hi As Long, lo As Long, v As Long
    For i = 1 To 8
        v = InStr(1, "0123456789abcdef", Mid$(LCase$(s), i, 1)) - 1
        If i <= 4 Then hi = hi * 16 + v Else lo = lo * 16 + v
    Next
    HexToLong = Combine16(hi, lo)
End Function

Private Function HighWord(ByVal x As Long) As Long
    HighWord = ((x And &HFFFF0000) \ &H10000) And &HFFFF&
End Function

Private Function Combine16(ByVal hi As Long, ByVal lo As Long) As Long
    If hi >= &H8000& Then
        Combine16 = ((hi - &H10000) * &H10000) Or lo
    Else
        Combine16 = (hi * &H10000) Or lo
    End If
End Function

Private Function U32Add(ByVal a As Long, ByVal b As Long) As Long
    Dim lo As Long, hi As Long
    lo = (a And &HFFFF&) + (b And &HFFFF&)
    hi = (HighWord(a) + HighWord(b) + (lo \ &H10000)) And &HFFFF&
    U32Add = Combine16(hi, lo And &HFFFF&)
End Function

Private Function ShiftRight(ByVal x As Long, ByVal n As Long) As Long
    If x >= 0 Then
        ShiftRight = x \ m_pow2(n)
    Else
        ShiftRight = ((x And &H7FFFFFFF) \ m_pow2(n)) Or m_pow2(31 - n)
    End If
End Function

Private Function ShiftLeft(ByVal x As Long, ByVal n As Long) As Long
    Dim r As Long
    r = (x And (m_pow2(31 - n) - 1)) * m_pow2(n)
    If (x And m_pow2(31 - n)) <> 0 Then r = r Or &H80000000
    ShiftLeft = r
End Function

Private Function RotR(ByVal x As Long, ByVal n As Long) As Long
    RotR = ShiftRight(x, n) Or ShiftLeft(x, 32 - n)
End Function

'==============================================================================
' Passwords and login
'==============================================================================
Public Function NewSalt() As String
    NewSalt = UCase$(Replace(NewUUID(), "-", ""))
End Function

Public Function PasswordProblem(ByVal Password As String, ByVal Username As String) As String
    ' "" when the password is acceptable, otherwise the reason in Arabic.
    If Len(Password) < MIN_PASSWORD_LENGTH Then
        PasswordProblem = "كلمة المرور يجب ألا تقل عن " & MIN_PASSWORD_LENGTH & " أحرف."
    ElseIf StrComp(Trim$(Password), Trim$(Username), vbTextCompare) = 0 Then
        PasswordProblem = "كلمة المرور لا يجوز أن تكون نفس اسم المستخدم."
    ElseIf Password <> Trim$(Password) Then
        PasswordProblem = "كلمة المرور لا تبدأ ولا تنتهي بمسافة."
    End If
End Function

Public Function LoginUser(ByVal Username As String, ByVal Password As String, ByRef EmployeeID As Long) As String
    ' "" on success (TempVars UserID is set), otherwise the message to show.
    Dim rs As DAO.Recordset, left1 As Long
    EmployeeID = 0
    Calendar = vbCalGreg
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Employees WHERE Username = " & SqlText(Trim$(Username)), _
                                     dbOpenDynaset)
    If rs.EOF Then
        LoginUser = "اسم المستخدم أو كلمة المرور غير صحيحة."
        rs.Close
        LogAction "LOGIN_FAILED", "Employees", "", "Unknown user: " & Left$(Username, 30)
        Exit Function
    End If
    If Not rs!IsActive Then
        LoginUser = "هذا الحساب معطّل. راجع مدير النظام."
    ElseIf Not IsNull(rs!LockedUntil) And Nz(rs!LockedUntil, 0) > Now Then
        LoginUser = "الحساب مقفل مؤقتًا بسبب محاولات دخول خاطئة حتى الساعة " & _
                    Format$(rs!LockedUntil, "hh:nn") & "."
    ElseIf IsNull(rs!PasswordHash) Then
        ' only the built-in administrator may enter without a password, and only until one is set:
        ' DoLogin then forces the password screen (closing it logs out, so this stays possible)
        If rs!EmployeeID = 1 And Len(Password) = 0 Then
            EmployeeID = 1
        Else
            LoginUser = "لم تُعيَّن كلمة مرور لهذا المستخدم بعد. اطلب من مدير النظام تعيينها."
        End If
    ElseIf PasswordHash(Password, Nz(rs!PasswordSalt, "")) <> rs!PasswordHash Then
        rs.Edit
        rs!FailedLoginCount = rs!FailedLoginCount + 1
        left1 = MAX_FAILED_LOGINS - rs!FailedLoginCount
        If left1 <= 0 Then
            rs!LockedUntil = DateAdd("n", LOCK_MINUTES, Now)
            rs!FailedLoginCount = 0
            LoginUser = "كلمة المرور غير صحيحة. تم قفل الحساب " & LOCK_MINUTES & " دقيقة."
        Else
            ' same text as an unknown user name: the screen must not reveal which names exist
            LoginUser = "اسم المستخدم أو كلمة المرور غير صحيحة."
        End If
        rs.Update
        LogAction "LOGIN_FAILED", "Employees", CStr(rs!EmployeeID)
    Else
        EmployeeID = rs!EmployeeID
    End If
    If EmployeeID > 0 Then
        rs.Edit
        rs!FailedLoginCount = 0
        rs!LockedUntil = Null
        rs!LastLoginAt = Now
        rs.Update
        TempVars.Add "UserID", EmployeeID
        LogAction "LOGIN", "Employees", CStr(EmployeeID)
    End If
    rs.Close
End Function

Public Function NeedsPasswordChange(ByVal EmployeeID As Long) As Boolean
    NeedsPasswordChange = Nz(DbValue("SELECT MustChangePassword FROM Employees WHERE EmployeeID = " & EmployeeID), False) _
        Or IsNull(DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & EmployeeID))
End Function

Public Function SetUserPassword(ByVal EmployeeID As Long, ByVal NewPassword As String, _
                                ByVal MustChangeNext As Boolean) As String
    ' Stores a new salt and hash, clears the lock. "" on success.
    Dim rs As DAO.Recordset, msg As String, salt As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Employees WHERE EmployeeID = " & EmployeeID, dbOpenDynaset)
    If rs.EOF Then
        SetUserPassword = "المستخدم غير موجود."
        rs.Close
        Exit Function
    End If
    msg = PasswordProblem(NewPassword, rs!Username)
    If Len(msg) > 0 Then
        SetUserPassword = msg
        rs.Close
        Exit Function
    End If
    salt = NewSalt()
    rs.Edit
    rs!PasswordSalt = salt
    rs!PasswordHash = PasswordHash(NewPassword, salt)
    rs!MustChangePassword = MustChangeNext
    rs!FailedLoginCount = 0
    rs!LockedUntil = Null
    rs.Update
    rs.Close
    LogAction "PASSWORD_SET", "Employees", CStr(EmployeeID), IIf(MustChangeNext, "temporary", "by user")
End Function

Public Function ChangeOwnPassword(ByVal OldPassword As String, ByVal NewPassword As String, _
                                  ByVal Confirm As String, ByVal RequireOld As Boolean) As String
    Dim id As Long, hashNow As Variant, salt As String
    id = CurrentUserID()
    If id = 0 Then
        ChangeOwnPassword = "سجّل الدخول أولًا."
        Exit Function
    End If
    hashNow = DbValue("SELECT PasswordHash FROM Employees WHERE EmployeeID = " & id)
    If RequireOld And Not IsNull(hashNow) Then
        salt = Nz(DbValue("SELECT PasswordSalt FROM Employees WHERE EmployeeID = " & id), "")
        If PasswordHash(OldPassword, salt) <> hashNow Then
            ChangeOwnPassword = "كلمة المرور الحالية غير صحيحة."
            Exit Function
        End If
    End If
    If NewPassword <> Confirm Then
        ChangeOwnPassword = "تأكيد كلمة المرور لا يطابق كلمة المرور الجديدة."
        Exit Function
    End If
    If Not IsNull(hashNow) Then
        salt = Nz(DbValue("SELECT PasswordSalt FROM Employees WHERE EmployeeID = " & id), "")
        If PasswordHash(NewPassword, salt) = hashNow Then
            ChangeOwnPassword = "اختر كلمة مرور مختلفة عن الحالية."
            Exit Function
        End If
    End If
    ChangeOwnPassword = SetUserPassword(id, NewPassword, False)
End Function

Public Sub ResetDeveloperPassword()
    ' The programmer forgot the password of "developer": run it from the Immediate window
    ' (Ctrl+G) of the programmer's own .accdb copy. Refused in an ACCDE (the client's file).
    Dim id As Variant, pwd As String, again As String, msg As String
    If IsCompiledFile() Then
        MsgBox Tr("غير متاح في ملف ACCDE. غيّر كلمة المرور من نسختك ACCDB."), vbExclamation + MSG_RTL, Tr(APP_TITLE)
        Exit Sub
    End If
    id = DbValue("SELECT EmployeeID FROM Employees WHERE IsDeveloper = True")
    If IsNull(id) Then
        MsgBox Tr("لا يوجد حساب مبرمج بعد. شغّل BuildSchema: يطلب كلمة مروره وينشئه."), vbExclamation + MSG_RTL, Tr(APP_TITLE)
        Exit Sub
    End If
    pwd = InputBox(Tr("كلمة مرور جديدة لحساب المبرمج developer (" & MIN_PASSWORD_LENGTH & " أحرف على الأقل):"), Tr(APP_TITLE))
    If Len(pwd) = 0 Then Exit Sub
    again = InputBox(Tr("اكتب كلمة المرور مرة أخرى للتأكيد:"), Tr(APP_TITLE))
    If again <> pwd Then
        MsgBox Tr("كلمتا المرور غير متطابقتين. لم يتغير شيء."), vbExclamation + MSG_RTL, Tr(APP_TITLE)
        Exit Sub
    End If
    msg = SetUserPassword(CLng(id), pwd, False)
    If Len(msg) = 0 Then
        CurrentDb.Execute "UPDATE Employees SET IsActive = True WHERE EmployeeID = " & CLng(id), dbFailOnError
        MsgBox Tr("تم تغيير كلمة مرور المبرمج. ادخل باسم developer وكلمة المرور الجديدة."), vbInformation + MSG_RTL, Tr(APP_TITLE)
    Else
        MsgBox Tr(msg & vbCrLf & "لم يتغير شيء."), vbExclamation + MSG_RTL, Tr(APP_TITLE)
    End If
End Sub

Private Function IsCompiledFile() As Boolean
    ' An ACCDE carries the database property MDE = "T".
    On Error Resume Next
    IsCompiledFile = (CurrentDb.Properties("MDE") = "T")
End Function

Public Sub EnsureTestUser()
    ' The in-Access tests run from the Immediate window as the built-in administrator.
    If CurrentUserID() = 0 Then TempVars.Add "UserID", 1
End Sub

Public Function IsAdministrator() As Boolean
    IsAdministrator = (Nz(DbValue("SELECT RoleID FROM Employees WHERE EmployeeID = " & CurrentUserID() & _
                                  " AND IsActive = True"), 0) = ADMIN_ROLE_ID)
End Function

Public Function IsDeveloper() As Boolean
    IsDeveloper = Nz(DbValue("SELECT IsDeveloper FROM Employees WHERE EmployeeID = " & CurrentUserID() & _
                             " AND IsActive = True"), False)
End Function

'==============================================================================
' Screen permissions
'==============================================================================
Public Function UsesCustomScreens(Optional ByVal EmployeeID As Long = -1) As Boolean
    ' A user whose screens are set one by one in frmUserScreens (never an administrator or the programmer).
    If EmployeeID = -1 Then EmployeeID = CurrentUserID()
    UsesCustomScreens = Nz(DbValue("SELECT CustomScreens FROM Employees WHERE EmployeeID = " & EmployeeID & _
        " AND IsActive = True AND IsDeveloper = False AND RoleID <> " & ADMIN_ROLE_ID), False)
End Function

Private Function ScreenTitle(ByVal FormName As String) As Variant
    ScreenTitle = DbValue("SELECT ScreenTitle FROM Screens WHERE ScreenName = " & SqlText(FormName))
End Function

Public Function CanOpenScreen(ByVal FormName As String, Optional ByVal Quiet As Boolean = False) As Boolean
    Dim key As String, title As Variant
    If FormName = "frmActivation" Then                  ' the activation of this computer: administrators
        CanOpenScreen = IsAdministrator()
        If Not CanOpenScreen And Not Quiet Then ShowWarning "تفعيل البرنامج لمدير النظام فقط."
        Exit Function
    End If
    If IsDeveloper() Then
        CanOpenScreen = True
        Exit Function
    End If
    title = ScreenTitle(FormName)
    If Not IsNull(title) And UsesCustomScreens() Then
        CanOpenScreen = Nz(DbValue("SELECT CanOpen FROM UserScreens WHERE EmployeeID = " & CurrentUserID() & _
                                   " AND ScreenName = " & SqlText(FormName)), False)
        If Not CanOpenScreen And Not Quiet Then ShowWarning "لا تملك صلاحية فتح شاشة «" & title & "». راجع مدير النظام."
        Exit Function
    End If
    key = ScreenPermission(FormName)
    If Len(key) = 0 Then
        CanOpenScreen = True
    ElseIf HasPermission(key) Then
        CanOpenScreen = True
    ElseIf Not Quiet Then
        ShowWarning "لا تملك صلاحية «" & Nz(DLookup("PermissionName", "Permissions", "PermissionKey = " & _
                    SqlText(key)), key) & "». راجع مدير النظام."
    End If
End Function

Public Function CanScreenAction(ByVal FormName As String, ByVal Action As String, _
                                Optional ByVal Quiet As Boolean = False) As Boolean
    ' Action: "ADD" (a new record, or saving a new document), "EDIT" or "DELETE".
    ' Only a user with his own screen permissions is limited, and only in a screen that has the action.
    Dim word As String, verb As String, title As Variant
    CanScreenAction = True
    If Not UsesCustomScreens() Then Exit Function
    Select Case UCase$(Action)
        Case "ADD": word = "Add": verb = "الإضافة والحفظ"
        Case "EDIT": word = "Edit": verb = "التعديل"
        Case "DELETE": word = "Delete": verb = "الحذف"
        Case Else: Exit Function
    End Select
    title = ScreenTitle(FormName)
    If IsNull(title) Then Exit Function
    If Not Nz(DbValue("SELECT Has" & word & " FROM Screens WHERE ScreenName = " & SqlText(FormName)), False) Then Exit Function
    CanScreenAction = Nz(DbValue("SELECT CanOpen AND Can" & word & " FROM UserScreens WHERE EmployeeID = " & _
                                 CurrentUserID() & " AND ScreenName = " & SqlText(FormName)), False)
    If Not CanScreenAction And Not Quiet Then ShowWarning "لا تملك صلاحية " & verb & " في شاشة «" & title & "». راجع مدير النظام."
End Function

Public Function CanChangeStoreName() As Boolean
    ' The shop name (invoices, QR code): the programmer, or an administrator when the programmer allows it.
    CanChangeStoreName = IsDeveloper()
    If Not CanChangeStoreName And IsAdministrator() Then CanChangeStoreName = Nz(SettingValue("AllowAdminCompanyName"), False)
End Function

'==============================================================================
' Users screen rules (called by modForms.FormBeforeUpdate for table Employees)
'==============================================================================
Public Function ValidateEmployee(ByVal frm As Access.Form) As Boolean
    Dim user As String, id As Variant, wasAdmin As Boolean, staysAdmin As Boolean
    user = Trim$(Nz(frm!Username.Value, ""))
    If Len(user) < 3 Or InStr(user, " ") > 0 Then
        ShowWarning "اسم المستخدم 3 أحرف على الأقل وبدون مسافات."
        SafeFocus frm!Username
        Exit Function
    End If
    frm!Username.Value = user
    id = frm!EmployeeID.Value
    If Not frm.NewRecord And Not IsNull(id) Then
        If id = CurrentUserID() Then
            If Not frm!IsActive.Value Then
                ShowWarning "لا يمكنك تعطيل حسابك وأنت مسجل الدخول به."
                Exit Function
            End If
            If frm!RoleID.Value <> Nz(frm!RoleID.OldValue, frm!RoleID.Value) Then
                ShowWarning "لا يمكنك تغيير دورك بنفسك."
                Exit Function
            End If
        End If
        wasAdmin = (Nz(frm!RoleID.OldValue, 0) = ADMIN_ROLE_ID And Nz(frm!IsActive.OldValue, False))
        staysAdmin = (frm!RoleID.Value = ADMIN_ROLE_ID And frm!IsActive.Value)
        If wasAdmin And Not staysAdmin Then
            If Nz(DCount("*", "Employees", "RoleID = " & ADMIN_ROLE_ID & " AND IsActive = True AND IsDeveloper = False " & _
                         "AND EmployeeID <> " & id), 0) = 0 Then
                ShowWarning "يجب أن يبقى مدير نظام نشط واحد على الأقل."
                Exit Function
            End If
        End If
    End If
    ValidateEmployee = True
End Function
