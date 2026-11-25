' MakeAccde.vbs - RetailStore: makes the ACCDE that is delivered to the customer (docs/88, phase 7D).
'
' The ACCDE is the front-end with its VBA compiled and its source removed: the code (and so LICENSE_SECRET)
' cannot be read or changed, and the screens and reports cannot be opened in design view.
'
' Before it:
'   1. python tools/release.py secret, then python tools/release.py package (release\RetailStore);
'   2. release\RetailStore\tools\BuildFrontEnd.vbs builds RetailStore_FE.accdb from the package modules
'      (they carry the release secret), with "end-user mode" = Yes;
'   3. close that file in every Access window.
' Double-click this file, or: wscript MakeAccde.vbs "D:\Shop\RetailStore_FE.accdb" ["D:\Shop\RetailStore_FE.accde"]
' Make the ACCDE with the same Access as the customer (32-bit or 64-bit, the same version or older):
' an ACCDE made with 64-bit Access does not open in 32-bit Access.
' Keep the ACCDB: every later change is made in it, then a new ACCDE is made.
Option Explicit

Const TITLE = "RetailStore - make ACCDE"
Const acSysCmdMakeAccde = 603           ' SysCmd 603 = File > Save As > Make ACCDE
Const msoAutomationSecurityLow = 1

Dim fso, src, dst, app, lockFile, answer

Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count > 0 Then
    src = WScript.Arguments(0)
Else
    src = InputBox("Full path of the built front-end (.accdb):", TITLE, _
                   fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName)), "RetailStore_FE.accdb"))
End If
If src = "" Then WScript.Quit 0
src = fso.GetAbsolutePathName(src)
If LCase(fso.GetExtensionName(src)) <> "accdb" Or Not fso.FileExists(src) Then
    MsgBox "Not an .accdb file: " & src, 16, TITLE
    WScript.Quit 1
End If
If WScript.Arguments.Count > 1 Then
    dst = fso.GetAbsolutePathName(WScript.Arguments(1))
Else
    dst = fso.BuildPath(fso.GetParentFolderName(src), fso.GetBaseName(src) & ".accde")
End If
If LCase(fso.GetExtensionName(dst)) <> "accde" Then dst = dst & ".accde"

lockFile = fso.BuildPath(fso.GetParentFolderName(src), fso.GetBaseName(src) & ".laccdb")
If fso.FileExists(lockFile) Then
    MsgBox "The file is open (" & fso.GetFileName(lockFile) & " exists)." & vbCrLf & _
           "Close it in every Access window, then run this again.", 48, TITLE
    WScript.Quit 1
End If
If fso.FileExists(dst) Then
    answer = MsgBox("Replace the existing file?" & vbCrLf & vbCrLf & dst, vbYesNo + vbQuestion + vbDefaultButton2, TITLE)
    If answer <> vbYes Then WScript.Quit 0
    On Error Resume Next
    fso.DeleteFile dst, True
    If Err.Number <> 0 Then
        MsgBox "Cannot replace " & dst & ": " & Err.Description, 16, TITLE
        WScript.Quit 1
    End If
    On Error GoTo 0
End If

Set app = CreateObject("Access.Application")
app.Visible = False
app.AutomationSecurity = msoAutomationSecurityLow   ' the VBA project must be enabled to compile; the file is not opened, so its start-up does not run
On Error Resume Next
app.SysCmd acSysCmdMakeAccde, src, dst
If Err.Number <> 0 Then
    answer = Err.Description
    Err.Clear
    app.Quit 2
    Set app = Nothing
    MsgBox "Access could not make the ACCDE:" & vbCrLf & answer & vbCrLf & vbCrLf & _
           "Open the ACCDB, run Debug > Compile in the VBA editor (it must compile with no error), close it and try again.", _
           16, TITLE
    WScript.Quit 1
End If
app.Quit 2
Set app = Nothing
On Error GoTo 0

If fso.FileExists(dst) Then
    MsgBox "Done:" & vbCrLf & dst & vbCrLf & vbCrLf & _
           "Deliver this file with RetailStore_BE.accdb (or the SQL Server link). Keep the ACCDB and release\license-secret.txt.", _
           64, TITLE
    WScript.Quit 0
End If
MsgBox "The ACCDE was not made. Run Debug > Compile in the ACCDB, then try again.", 16, TITLE
WScript.Quit 1
