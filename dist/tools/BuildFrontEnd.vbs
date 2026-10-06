' BuildFrontEnd.vbs - RetailStore: builds a ready-to-run RetailStore_FE.accdb in one step.
'
' Double-click this file (on a Windows PC with Microsoft Access 2010 or later). It:
'   1. creates a new, empty RetailStore_FE.accdb (default: the dist folder),
'   2. imports every module from dist\vba,
'   3. compiles and saves the VBA (Debug > Compile),
'   4. runs BuildSchema, BuildRelationships, BuildQueries, BuildForms, BuildReports
'      (RetailStore_BE.accdb is created next to the front-end, or upgraded if it exists),
'   5. compiles again (the screens carry their own code),
'   6. optionally loads the demo data and runs RunAllTests,
'   7. optionally switches to end-user mode (opens on the login screen).
' Access shows a message after each build step: read it and press OK.
' Every step is written to BuildFrontEnd.log next to the new file.
'
' Command line: wscript BuildFrontEnd.vbs "D:\Shop\RetailStore_FE.accdb"
' Arabic text: Windows "Language for non-Unicode programs" must be Arabic.
' First login: admin with no password; the program then asks for a new one.
Option Explicit

Const TITLE = "RetailStore - build front-end"
Const acModule = 5
Const acCmdCompileAndSaveAllModules = 126

Dim fso, here, vbaDir, fePath, logPath, logFile, app, failures

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
vbaDir = fso.BuildPath(fso.GetParentFolderName(here), "vba")
If Not fso.FolderExists(vbaDir) Then
    MsgBox "Folder not found: " & vbaDir & vbCrLf & "Keep this script in dist\tools next to dist\vba.", 16, TITLE
    WScript.Quit 1
End If

If WScript.Arguments.Count > 0 Then
    fePath = WScript.Arguments(0)
Else
    fePath = InputBox("Full path of the new front-end file:" & vbCrLf & vbCrLf & _
                      "RetailStore_BE.accdb (the data) is created next to it, or upgraded if it is already there.", _
                      TITLE, fso.BuildPath(fso.GetParentFolderName(here), "RetailStore_FE.accdb"))
End If
If fePath = "" Then WScript.Quit 0
fePath = fso.GetAbsolutePathName(fePath)
If LCase(fso.GetExtensionName(fePath)) <> "accdb" Then fePath = fePath & ".accdb"
If Not fso.FolderExists(fso.GetParentFolderName(fePath)) Then
    MsgBox "Folder not found: " & fso.GetParentFolderName(fePath), 16, TITLE
    WScript.Quit 1
End If

If fso.FileExists(fePath) Then
    If MsgBox(fePath & vbCrLf & "already exists. Replace it?" & vbCrLf & vbCrLf & _
              "Only the front-end is replaced; the data file RetailStore_BE.accdb is kept.", _
              vbYesNo + vbExclamation + vbDefaultButton2, TITLE) <> vbYes Then WScript.Quit 0
    On Error Resume Next
    fso.DeleteFile fePath, True
    If Err.Number <> 0 Then
        MsgBox "Cannot delete " & fePath & vbCrLf & "Close it in Access first." & vbCrLf & Err.Description, 16, TITLE
        WScript.Quit 1
    End If
    On Error GoTo 0
End If

logPath = fso.BuildPath(fso.GetParentFolderName(fePath), "BuildFrontEnd.log")
Set logFile = fso.CreateTextFile(logPath, True, True)
LogLine "=== BuildFrontEnd " & Now & " ==="
LogLine "Front-end: " & fePath
LogLine "Modules:   " & vbaDir
failures = 0

' --- 1. new database -----------------------------------------------------------
On Error Resume Next
Set app = CreateObject("Access.Application")
If Err.Number <> 0 Then Fail "Microsoft Access is not installed (" & Err.Description & ")"
app.AutomationSecurity = 1          ' msoAutomationSecurityLow: the new file's code may run
Err.Clear
app.Visible = True
app.NewCurrentDatabase fePath
If Err.Number <> 0 Then Fail "Cannot create the database: " & Err.Description
On Error GoTo 0
LogLine "[OK] new database"

' --- 2. modules -------------------------------------------------------------
Dim f, count
count = 0
For Each f In fso.GetFolder(vbaDir).Files
    If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
        ImportModule f.Path, fso.GetBaseName(f.Name)
        count = count + 1
    End If
Next
LogLine "[OK] " & count & " modules imported"

' --- 3. compile ---------------------------------------------------------------
Compile "modules"

' --- 4. build -----------------------------------------------------------------
RunStep "BuildSchema", True
RunStep "BuildRelationships", True
RunStep "BuildQueries", True
RunStep "BuildForms", True
RunStep "BuildReports", True

' --- 5. compile the screens' code ---------------------------------------------
Compile "modules, screens and reports"

' --- 6. demo data and tests ------------------------------------------------------
If failures = 0 Then
    ' LoadDemoData asks first (default No): skip it for a real shop.
    RunStep "LoadDemoData", False
    If MsgBox("Run all the tests now (RunAllTests)?" & vbCrLf & "Each test leaves no data behind.", _
              vbYesNo + vbQuestion, TITLE) = vbYes Then RunStep "RunAllTests", False
End If

' --- 7. start-up --------------------------------------------------------------
If failures = 0 Then
    If MsgBox("Switch to end-user mode?" & vbCrLf & vbCrLf & _
              "Yes: the file opens on the login screen without the Access tools (SHIFT is disabled;" & vbCrLf & _
              "      to get back use the backup screen > developer mode, or EnableShiftKey.vbs)." & vbCrLf & _
              "No:  the file stays in developer mode (run InstallUserMode later).", _
              vbYesNo + vbQuestion, TITLE) = vbYes Then RunStep "InstallUserMode", False
End If

On Error Resume Next
app.CloseCurrentDatabase
app.Quit
Set app = Nothing
On Error GoTo 0

LogLine "=== finished: " & failures & " problem(s) ==="
logFile.Close
If failures = 0 Then
    MsgBox "Done." & vbCrLf & vbCrLf & fePath & vbCrLf & vbCrLf & _
           "Open it, log in as admin (no password) and set a new password." & vbCrLf & _
           "Details: " & logPath, 64, TITLE
Else
    MsgBox failures & " step(s) failed. Details: " & logPath & vbCrLf & vbCrLf & _
           "Send the log (and any message text) so it can be fixed.", 48, TITLE
End If
WScript.Quit failures

'==============================================================================
Sub LogLine(text)
    logFile.WriteLine text
End Sub

Sub Fail(text)
    LogLine "[X]  " & text
    logFile.Close
    On Error Resume Next
    If IsObject(app) Then
        If Not app Is Nothing Then app.Quit 2      ' acQuitSaveNone
    End If
    MsgBox text & vbCrLf & vbCrLf & "Details: " & logPath, 16, TITLE
    WScript.Quit 1
End Sub

Sub ImportModule(path, name)
    On Error Resume Next
    app.VBE.ActiveVBProject.VBComponents.Import path
    If Err.Number <> 0 Then
        Err.Clear
        app.LoadFromText acModule, name, path
    End If
    If Err.Number <> 0 Then Fail "Cannot import " & name & ": " & Err.Description
    app.DoCmd.Save acModule, name
    Err.Clear
End Sub

Sub Compile(what)
    On Error Resume Next
    app.DoCmd.RunCommand acCmdCompileAndSaveAllModules
    Err.Clear
    If app.IsCompiled Then
        LogLine "[OK] compiled: " & what
    Else
        Fail "The VBA code does not compile (" & what & ")." & vbCrLf & _
             "Open the file, press Alt+F11, then Debug > Compile: the line with the error is shown." & vbCrLf & _
             "Send that line and the message."
    End If
End Sub

Sub RunStep(procName, required)
    ' required: a build step; it must succeed. Otherwise (demo data, tests, user mode)
    ' a False result is only logged: the user may have answered No to its question.
    Dim result, errText
    On Error Resume Next
    Err.Clear
    result = app.Run(procName)
    errText = Err.Description
    If Err.Number <> 0 Then
        Err.Clear
        If required Then Fail procName & " failed: " & errText
        LogLine "[X]  " & procName & ": " & errText
        failures = failures + 1
    ElseIf IsEmpty(result) Or result = True Then
        LogLine "[OK] " & procName
    ElseIf required Then
        Fail procName & " reported a problem (see its message). Fix it, then run this script again."
    Else
        LogLine "[--] " & procName & ": not done or not passed (see its message)"
    End If
End Sub
