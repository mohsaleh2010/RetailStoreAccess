' EnableShiftKey.vbs - RetailStore: re-enable the SHIFT bypass key of the front-end file.
'
' In user mode the program disables SHIFT (AllowBypassKey = False), so the Access
' design tools cannot be reached. The normal way back is: frmBackup > Developer mode
' (administrator). If no administrator can log in, use this script:
'   drag RetailStore_FE.accdb onto this file (or double-click and type the path),
'   then open the .accdb while holding SHIFT and run InstallDeveloperMode.
' Needs the Access database engine (installed with Office). On 64-bit Office with a
' 32-bit script host, run: %windir%\System32\cscript.exe EnableShiftKey.vbs <file>
Option Explicit
Dim path, eng, db, prp
If WScript.Arguments.Count > 0 Then
    path = WScript.Arguments(0)
Else
    path = InputBox("Full path of the front-end file (.accdb):", "RetailStore - enable SHIFT key")
End If
If path = "" Then WScript.Quit
Set eng = CreateObject("DAO.DBEngine.120")
Set db = eng.OpenDatabase(path)
On Error Resume Next
db.Properties("AllowBypassKey") = True
If Err.Number <> 0 Then
    Err.Clear
    Set prp = db.CreateProperty("AllowBypassKey", 1, True)
    db.Properties.Append prp
End If
On Error GoTo 0
db.Close
MsgBox "The SHIFT key is enabled again." & vbCrLf & "Open the file while holding SHIFT.", 64, "RetailStore"
