Attribute VB_Name = "modActivation"
'==============================================================================
' modActivation  -  Retail Store Management System
'
' Copy protection: the program runs only on computers activated with a code
' from the programmer.
'   MachineID           "XXXX-XXXX-XXXX-XXXX": SHA-256 of the motherboard serial,
'                       the processor id and the serial of the Windows drive.
'                       A copy of the files on another computer has another id.
'   ActivationCodeFor   "XXXXX-XXXXX-XXXXX-XXXXX": SHA-256 of LICENSE_SECRET and
'                       the machine id. Only who knows LICENSE_SECRET can make a code.
'   IsActivated         this computer has a valid code in table Activations (back-end).
'   ActivateMachine     checks and saves the code of this computer (administrators).
' Login: on a computer that is not activated only an administrator may log in, and
' only to activate it (frmActivation); the programmer is never stopped.
' The programmer also makes the codes for the computers of customers in frmActivation.
'
' BEFORE DELIVERING THE PROGRAM:
'   1. change LICENSE_SECRET below to your own long secret text (and keep it);
'   2. deliver an .accde (File > Save As > Make ACCDE): the code, and so the
'      secret, cannot be read or changed in it.
' Same results as tools/activation_reference.py.
'==============================================================================
Option Compare Database
Option Explicit

Private Const LICENSE_SECRET As String = "RetailStore|change-this-secret-before-delivery|2026"
Private m_machineID As String
Private m_activated As Boolean

'------------------------------------------------------------------------------
' Machine id and codes
'------------------------------------------------------------------------------
Public Function MachineID() As String
    Dim raw As String
    If Len(m_machineID) = 0 Then
        raw = WmiValue("Win32_BaseBoard", "SerialNumber") & "|" & WmiValue("Win32_Processor", "ProcessorId") & _
              "|" & SystemDriveSerial()
        m_machineID = MachineIDFrom(raw)
    End If
    MachineID = m_machineID
End Function

Public Function MachineIDFrom(ByVal HardwareText As String) As String
    Dim h As String
    h = UCase$(Sha256Text("RS-MACHINE|" & UCase$(HardwareText)))
    MachineIDFrom = Mid$(h, 1, 4) & "-" & Mid$(h, 5, 4) & "-" & Mid$(h, 9, 4) & "-" & Mid$(h, 13, 4)
End Function

Public Function ActivationCodeFor(ByVal ForMachineID As String) As String
    Dim h As String
    h = UCase$(Sha256Text(LICENSE_SECRET & "|" & NormalizeKey(ForMachineID)))
    ActivationCodeFor = Mid$(h, 1, 5) & "-" & Mid$(h, 6, 5) & "-" & Mid$(h, 11, 5) & "-" & Mid$(h, 16, 5)
End Function

Public Function NormalizeKey(ByVal Text As String) As String
    ' Upper case without dashes and spaces: codes may be typed in any form.
    NormalizeKey = UCase$(Replace(Replace(Trim$(Text), "-", ""), " ", ""))
End Function

Public Function IsValidCode(ByVal ForMachineID As String, ByVal Code As String) As Boolean
    IsValidCode = (Len(NormalizeKey(ForMachineID)) = 16 And Len(NormalizeKey(Code)) = 20 And _
                   NormalizeKey(Code) = NormalizeKey(ActivationCodeFor(ForMachineID)))
End Function

Private Function WmiValue(ByVal ClassName As String, ByVal PropName As String) As String
    Dim wmi As Object, items As Object, item As Object
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set items = wmi.ExecQuery("SELECT " & PropName & " FROM " & ClassName)
    For Each item In items
        WmiValue = Trim$(Nz(item.Properties_(PropName).Value, ""))
        Exit For
    Next
End Function

Private Function SystemDriveSerial() As String
    On Error Resume Next
    SystemDriveSerial = Hex$(CreateObject("Scripting.FileSystemObject").GetDrive(Environ$("SystemDrive")).SerialNumber)
End Function

'------------------------------------------------------------------------------
' Activation of this computer
'------------------------------------------------------------------------------
Public Function IsActivated() As Boolean
    If Not m_activated Then
        m_activated = IsValidCode(MachineID(), Nz(DbValue("SELECT ActivationCode FROM Activations WHERE MachineID = " & _
                                                          SqlText(MachineID())), ""))
    End If
    IsActivated = m_activated
End Function

Public Function ActivateMachine(ByVal Code As String) As String
    ' "" on success.
    ActivateMachine = SaveActivation(MachineID(), Code, Environ$("COMPUTERNAME"))
    If Len(ActivateMachine) = 0 Then m_activated = True
End Function

Public Function SaveActivation(ByVal ForMachineID As String, ByVal Code As String, _
                               ByVal ComputerName As String) As String
    Dim id As String
    If Not IsAdministrator() Then
        SaveActivation = "تفعيل البرنامج لمدير النظام فقط."
        Exit Function
    End If
    If Not IsValidCode(ForMachineID, Code) Then
        SaveActivation = "كود التفعيل غير صحيح لهذا الجهاز." & vbCrLf & _
                         "تأكد أن المبرمج أصدره لرقم الجهاز الظاهر في الشاشة."
        Exit Function
    End If
    id = UCase$(Trim$(ForMachineID))
    CurrentDb.Execute "DELETE FROM Activations WHERE MachineID = " & SqlText(id), dbFailOnError
    CurrentDb.Execute "INSERT INTO Activations (MachineID, ActivationCode, ComputerName, ActivatedAt, EmployeeID) " & _
        "VALUES (" & SqlText(id) & ", " & SqlText(UCase$(Trim$(Code))) & ", " & SqlText(Left$(ComputerName, 64)) & _
        ", Now(), " & CurrentUserID() & ")", dbFailOnError
    LogAction "ACTIVATE", "Activations", id, ComputerName
End Function

Public Function RemoveActivation(ByVal ActivationID As Long) As String
    Dim id As Variant
    If Not IsAdministrator() Then
        RemoveActivation = "إلغاء التفعيل لمدير النظام فقط."
        Exit Function
    End If
    id = DbValue("SELECT MachineID FROM Activations WHERE ActivationID = " & ActivationID)
    If IsNull(id) Then
        RemoveActivation = "اختر الجهاز من القائمة."
        Exit Function
    End If
    CurrentDb.Execute "DELETE FROM Activations WHERE ActivationID = " & ActivationID, dbFailOnError
    If id = MachineID() Then m_activated = False
    LogAction "DEACTIVATE", "Activations", CStr(id)
End Function

'------------------------------------------------------------------------------
' Screen frmActivation (OpenArgs "LOGIN": opened by the login of an administrator
' on a computer that is not activated yet)
'------------------------------------------------------------------------------
Public Sub ActivationLoad(ByVal frm As Access.Form)
    Dim dev As Boolean
    Calendar = vbCalGreg
    frm!txtMachineID.Value = MachineID()
    dev = IsDeveloper()
    frm!boxDeveloper.Visible = dev
    frm!lblDevCap.Visible = dev
    frm!txtForMachine.Visible = dev
    frm!btnGenerate.Visible = dev
    frm!txtGenerated.Visible = dev
    ShowActivationState frm
End Sub

Private Sub ShowActivationState(ByVal frm As Access.Form)
    frm!lstMachines.Requery
    If IsActivated() Then
        frm!lblState.Caption = "البرنامج مفعّل على هذا الجهاز."
        frm!lblState.ForeColor = CLR_SUCCESS
    Else
        frm!lblState.Caption = "البرنامج غير مفعّل على هذا الجهاز: أرسل رقم الجهاز للمبرمج واكتب الكود الذي يرسله."
        frm!lblState.ForeColor = CLR_DANGER
    End If
End Sub

Public Sub ActivateThisMachine(ByVal frm As Access.Form)
    Dim msg As String
    msg = ActivateMachine(Nz(frm!txtCode.Value, ""))
    If Len(msg) > 0 Then
        ShowWarning msg
        SafeFocus frm!txtCode
        Exit Sub
    End If
    frm!txtCode.Value = Null
    ShowActivationState frm
    ShowInfo "تم تفعيل البرنامج على هذا الجهاز."
    If Nz(frm.OpenArgs, "") = "LOGIN" Then DoCmd.Close acForm, frm.Name
End Sub

Public Sub RemoveSelectedActivation(ByVal frm As Access.Form)
    Dim msg As String
    If IsNull(frm!lstMachines.Value) Then
        ShowWarning "اختر الجهاز من القائمة."
        Exit Sub
    End If
    If Not AskYesNo("إلغاء تفعيل الجهاز المحدد؟ لن يعمل عليه البرنامج حتى يُفعَّل بكود جديد.") Then Exit Sub
    msg = RemoveActivation(frm!lstMachines.Value)
    If Len(msg) > 0 Then ShowWarning msg
    ShowActivationState frm
End Sub

Public Sub CopyMachineID(ByVal frm As Access.Form)
    On Error Resume Next
    CreateObject("htmlfile").parentWindow.clipboardData.SetData "text", CStr(frm!txtMachineID.Value)
    If Err.Number = 0 Then
        ShowInfo "تم نسخ رقم الجهاز. أرسله للمبرمج."
    Else
        SafeFocus frm!txtMachineID                             ' select it, then Ctrl+C
        frm!txtMachineID.SelStart = 0
        frm!txtMachineID.SelLength = Len(frm!txtMachineID.Text)
    End If
End Sub

Public Sub GenerateActivationCode(ByVal frm As Access.Form)
    ' The programmer: the code for the machine id a customer sent.
    Dim id As String
    If Not IsDeveloper() Then
        ShowWarning "توليد أكواد التفعيل للمبرمج فقط."
        Exit Sub
    End If
    id = NormalizeKey(Nz(frm!txtForMachine.Value, ""))
    If Len(id) <> 16 Then
        ShowWarning "رقم الجهاز 16 حرفًا بالشكل XXXX-XXXX-XXXX-XXXX."
        SafeFocus frm!txtForMachine
        Exit Sub
    End If
    frm!txtGenerated.Value = ActivationCodeFor(id)
    LogAction "LICENSE_CODE", "Activations", id
End Sub
