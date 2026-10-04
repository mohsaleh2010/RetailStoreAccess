Attribute VB_Name = "modBackup"
'==============================================================================
' modBackup  -  Retail Store Management System (Phase 10)
'
' Backup and restore of the back-end file (the data):
'   BackupNow       copies the back end to <folder>\<name>_yyyy-mm-dd_hhnnss.accdb,
'                   opens the copy to verify it, keeps the newest BackupKeepCount files
'   RestoreBackup   puts a backup back in place (after a safety copy of the current data)
'   OfferBackupOnExit  asks for a backup at logout/exit if none was made today
' Folder: Settings.BackupFolder, or "Backups" next to the front end.
' Pure helpers (BackupFileName, OldBackups) have the same results as
' tools/security_reference.py.
'==============================================================================
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------
' Pure helpers
'------------------------------------------------------------------------------
Public Function BackupFileName(ByVal BaseName As String, ByVal When As Date) As String
    Dim saved As Integer
    saved = Calendar
    Calendar = vbCalGreg
    ' built from parts: the meaning of "_" and "n" in format strings differs between environments
    BackupFileName = BaseName & "_" & Format$(When, "yyyy-mm-dd") & "_" & Format$(Hour(When), "00") & _
                     Format$(Minute(When), "00") & Format$(Second(When), "00") & ".accdb"
    Calendar = saved
End Function

Public Function OldBackups(ByVal NameList As String, ByVal KeepCount As Long) As String
    ' NameList: file names separated by "|". Returns the ones to delete ("|"-separated),
    ' i.e. all but the newest KeepCount (the time stamp in the name sorts them).
    Dim names() As String, i As Long, j As Long, tmp As String, n As Long, out As String
    If Len(NameList) = 0 Then Exit Function
    If KeepCount < 1 Then KeepCount = 1
    names = Split(NameList, "|")
    n = UBound(names) + 1
    For i = 1 To n - 1                       ' insertion sort, ascending
        tmp = names(i)
        j = i - 1
        Do While j >= 0
            If StrComp(names(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = tmp
    Next
    For i = 0 To n - KeepCount - 1
        If Len(out) > 0 Then out = out & "|"
        out = out & names(i)
    Next
    OldBackups = out
End Function

'------------------------------------------------------------------------------
' Paths
'------------------------------------------------------------------------------
Public Function BackendFilePath() As String
    ' The back-end file of the linked tables, or this file when the tables are local.
    Dim cn As String, p As Long
    On Error Resume Next
    cn = CurrentDb.TableDefs("Settings").Connect
    On Error GoTo 0
    p = InStr(1, cn, ";DATABASE=", vbTextCompare)
    If p > 0 Then
        BackendFilePath = Mid$(cn, p + 10)
        If InStr(BackendFilePath, ";") > 0 Then BackendFilePath = Left$(BackendFilePath, InStr(BackendFilePath, ";") - 1)
    Else
        BackendFilePath = CurrentProject.FullName
    End If
End Function

Public Function BackupFolderPath() As String
    Dim folder As String
    folder = Trim$(Nz(SettingValue("BackupFolder"), ""))
    If Len(folder) = 0 Then folder = CurrentProject.Path & "\Backups"
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    BackupFolderPath = folder
End Function

Private Function BaseNameOf(ByVal FullPath As String) As String
    Dim f As String
    f = Mid$(FullPath, InStrRev(FullPath, "\") + 1)
    If InStrRev(f, ".") > 0 Then f = Left$(f, InStrRev(f, ".") - 1)
    BaseNameOf = f
End Function

'------------------------------------------------------------------------------
' Backup
'------------------------------------------------------------------------------
Public Function BackupNow(ByRef BackupFile As String, Optional ByVal Folder As String = "", _
                          Optional ByVal Suffix As String = "") As String
    ' "" on success; BackupFile receives the full path of the copy.
    Dim fso As Object, source As String, baseName As String, deleted As String, v As Variant
    On Error GoTo EH
    BackupFile = ""
    If Not HasPermission("BACKUP") Then
        BackupNow = "لا تملك صلاحية النسخ الاحتياطي."
        Exit Function
    End If
    source = BackendFilePath()
    If Len(Dir$(source)) = 0 Then
        BackupNow = "ملف البيانات غير موجود: " & source
        Exit Function
    End If
    If Len(Folder) = 0 Then Folder = BackupFolderPath()
    baseName = BaseNameOf(source) & Suffix
    BackupFile = Folder & "\" & BackupFileName(baseName, Now)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile source, BackupFile, False
    If Not BackupIsReadable(BackupFile) Then
        BackupNow = "تم النسخ لكن تعذر فتح النسخة للتحقق منها. أعد المحاولة بعد إغلاق البرنامج على الأجهزة الأخرى."
        Exit Function
    End If
    If Len(Suffix) = 0 Then
        deleted = OldBackups(BackupList(Folder, BaseNameOf(source)), Nz(SettingValue("BackupKeepCount"), 30))
        If Len(deleted) > 0 Then
            For Each v In Split(deleted, "|")
                Kill Folder & "\" & v
            Next
        End If
    End If
    LogAction "BACKUP", "Backup", "", BackupFile
    Exit Function
EH:
    BackupNow = "تعذر عمل النسخة الاحتياطية: " & Err.Description & " (" & Err.Number & ")"
End Function

Public Function BackupIsReadable(ByVal FilePath As String) As Boolean
    Dim db As DAO.Database, tdf As DAO.TableDef
    On Error GoTo EH
    Set db = DBEngine.OpenDatabase(FilePath, False, True)
    For Each tdf In db.TableDefs
        If tdf.Name = "Settings" Then BackupIsReadable = True
    Next
    db.Close
    Exit Function
EH:
    BackupIsReadable = False
End Function

Public Function BackupList(ByVal Folder As String, ByVal BaseName As String) As String
    ' "|"-separated names of the regular backups of BaseName in Folder.
    Dim f As String, out As String
    f = Dir$(Folder & "\" & BaseName & "_????-??-??_??????.accdb")
    Do While Len(f) > 0
        If Len(out) > 0 Then out = out & "|"
        out = out & f
        f = Dir$()
    Loop
    BackupList = out
End Function

Public Function LastBackupDate() As Variant
    LastBackupDate = DMax("LogDate", "AuditLog", "ActionType = 'BACKUP'")
End Function

Public Sub OfferBackupOnExit()
    ' At logout / exit: when the user may back up and no backup was made today.
    Dim lastDate As Variant, file As String, msg As String
    If Not HasPermission("BACKUP") Then Exit Sub
    lastDate = LastBackupDate()
    If Not IsNull(lastDate) Then
        If DateValue(lastDate) = Date Then Exit Sub
    End If
    If Not AskYesNo("لم تُعمل نسخة احتياطية اليوم. هل تريد عمل نسخة الآن؟") Then Exit Sub
    msg = BackupNow(file)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo "تمت النسخة الاحتياطية:" & vbCrLf & file
    End If
End Sub

'------------------------------------------------------------------------------
' Restore
'------------------------------------------------------------------------------
Public Function RestoreBackup(ByVal BackupFile As String) As String
    ' Replaces the back end with BackupFile after a safety copy ("_BeforeRestore").
    ' Every other user must have closed the program. The program closes afterwards.
    Dim fso As Object, target As String, safety As String, msg As String, i As Long
    On Error GoTo EH
    If Not HasPermission("BACKUP") Then
        RestoreBackup = "لا تملك صلاحية النسخ الاحتياطي."
        Exit Function
    End If
    If Not BackupIsReadable(BackupFile) Then
        RestoreBackup = "الملف المختار ليس نسخة صالحة من بيانات البرنامج."
        Exit Function
    End If
    target = BackendFilePath()
    If StrComp(target, CurrentProject.FullName, vbTextCompare) = 0 Then
        RestoreBackup = "البيانات داخل هذا الملف نفسه (غير مقسمة)، فلا يمكن استعادتها وهو مفتوح." & vbCrLf & _
                        "أغلق البرنامج وانسخ ملف النسخة الاحتياطية مكانه يدويًا."
        Exit Function
    End If
    msg = BackupNow(safety, , "_BeforeRestore")
    If Len(msg) > 0 Then
        RestoreBackup = "تعذرت النسخة الوقائية قبل الاستعادة، فلم تتم الاستعادة: " & msg
        Exit Function
    End If
    LogAction "RESTORE", "Backup", "", BackupFile
    For i = Forms.Count - 1 To 0 Step -1           ' release every connection to the back end
        DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile BackupFile, target, True
    Exit Function
EH:
    RestoreBackup = "تعذرت الاستعادة: " & Err.Description & vbCrLf & _
                    "تأكد أن البرنامج مغلق على كل الأجهزة الأخرى ثم أعد المحاولة." & vbCrLf & _
                    "بياناتك الحالية لم تتغير، ونسختها الوقائية: " & safety
End Function
