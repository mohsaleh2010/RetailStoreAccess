Attribute VB_Name = "modSqlServer"
'==============================================================================
' modSqlServer  -  Retail Store Management System
'
' Plan 5B (docs/68): the move of the data to SQL Server, from the program itself (screen frmSqlServer,
' permission SQL_SERVER). The connection is kept in the front-end (each PC): database properties
' SqlServer / SqlDatabase / SqlUser / SqlDriver, and SqlAccessPath = the .accdb file to go back to.
'   ServerConnect      the ODBC connection string (Windows login, or a SQL login), without a DSN
'   RunSchemaScript    dist\sql\RetailStore_Schema.sql (docs/67), batch by batch (GO), from Access
'   TransferAllData    every table of the back-end -> the server: the server tables are emptied, their rules
'                      switched off, the rows copied with their numbers (IDENTITY_INSERT) by pass-through
'                      INSERT batches, then every rule and relationship checked again
'   CompareAllData     each table: the number of rows and the total of each amount field, Access vs server
'   LinkToServer       the program's tables linked to the server (after a clean comparison);
'   LinkToAccess       linked back to the .accdb file
'   (plan 5D, docs/71: the saved queries follow, modServerQueries; "Prepare the database" also runs
'   RetailStore_Queries.sql, the views of the saved queries)
'   CheckServerAtStartup  the login screen says so when the server does not answer
' The pure parts (SqlLiteral, ServerConnect, SplitSqlBatches, InsertBatchSql) have a Python mirror in
' tools/sql_transfer_reference.py and run in LibreOffice in the tests.
'==============================================================================
Option Compare Database
Option Explicit

Private Const BATCH_ROWS As Long = 200          ' rows in one INSERT ... VALUES (SQL Server allows 1000)
Private Const BATCH_CHARS As Long = 400000      ' and at most this many characters
Private m_lastError As String

'------------------------------------------------------------------------------
' Pure helpers (no database)
'------------------------------------------------------------------------------
Public Function SqlLiteral(ByVal v As Variant, ByVal FieldType As Integer) As String
    ' A value of an Access field as a T-SQL literal. FieldType = DAO type (dbBoolean 1, dbDate 8, dbText 10,
    ' dbMemo 12, dbCurrency 5, dbLong 4, dbInteger 3, dbByte 2, dbDouble 7, dbSingle 6).
    If IsNull(v) Or IsEmpty(v) Then
        SqlLiteral = "NULL"
        Exit Function
    End If
    Select Case FieldType
        Case 1                                                    ' Yes/No: Access True = -1, BIT 1
            SqlLiteral = IIf(CBool(v), "1", "0")
        Case 8                                                    ' yyyymmdd hh:nn:ss, any server language
            SqlLiteral = "'" & Format$(v, "yyyymmdd hh:nn:ss") & "'"
        Case 10, 12
            SqlLiteral = "N'" & Replace(CStr(v), "'", "''") & "'"
        Case Else                                                 ' numbers with a dot, whatever the PC locale
            SqlLiteral = Trim$(Str$(v))
    End Select
End Function

Public Function ServerConnect(ByVal Server As String, ByVal Database As String, ByVal UserName As String, _
                              ByVal Password As String, ByVal Driver As String) As String
    ' ODBC connection string without a DSN. No user = Windows login (Trusted_Connection).
    Dim s As String
    s = "ODBC;DRIVER={" & Driver & "};SERVER=" & Server & ";DATABASE=" & Database & ";"
    If Len(Trim$(UserName)) = 0 Then
        s = s & "Trusted_Connection=Yes;"
    Else
        s = s & "UID=" & UserName & ";PWD=" & Password & ";"
    End If
    ' driver 18 encrypts by default: a shop server has no trusted certificate
    If InStr(1, Driver, "ODBC Driver 18", vbTextCompare) > 0 Then s = s & "TrustServerCertificate=Yes;"
    ServerConnect = s & "APP=RetailStore;"
End Function

Public Function SplitSqlBatches(ByVal ScriptText As String) As Variant
    ' The batches of a script: the text between lines that hold only GO (as SSMS and sqlcmd do).
    Dim lines As Variant, i As Long, cur As String, out() As String, n As Long
    lines = Split(Replace(ScriptText, vbCrLf, vbLf), vbLf)
    ReDim out(0)
    For i = 0 To UBound(lines)
        If UCase$(Trim$(lines(i))) = "GO" Then
            If Len(Trim$(Replace(cur, vbLf, ""))) > 0 Then
                ReDim Preserve out(n)
                out(n) = cur
                n = n + 1
            End If
            cur = ""
        Else
            cur = cur & lines(i) & vbLf
        End If
    Next
    If Len(Trim$(Replace(cur, vbLf, ""))) > 0 Then
        ReDim Preserve out(n)
        out(n) = cur
        n = n + 1
    End If
    If n = 0 Then
        SplitSqlBatches = Array()
    Else
        SplitSqlBatches = out
    End If
End Function

Public Function SqlBatchCount(ByVal ScriptText As String) As Long
    Dim b As Variant
    b = SplitSqlBatches(ScriptText)
    SqlBatchCount = UBound(b) + 1
End Function

Public Function InsertBatchSql(ByVal TableName As String, ByVal Columns As String, ByVal ValueRows As String, _
                               ByVal HasIdentity As Boolean) As String
    ' One INSERT of several rows ("(...),(...)"), with IDENTITY_INSERT around it to keep the numbers.
    Dim s As String
    If HasIdentity Then s = "SET IDENTITY_INSERT [dbo].[" & TableName & "] ON;" & vbLf
    s = s & "INSERT INTO [dbo].[" & TableName & "] (" & Columns & ") VALUES " & ValueRows & ";" & vbLf
    If HasIdentity Then s = s & "SET IDENTITY_INSERT [dbo].[" & TableName & "] OFF;" & vbLf
    InsertBatchSql = s
End Function

Public Function SameAmount(ByVal a As Variant, ByVal b As Variant) As Boolean
    SameAmount = Abs(CDbl(Nz(a, 0)) - CDbl(Nz(b, 0))) < 0.00005
End Function

'------------------------------------------------------------------------------
' The connection of this front-end
'------------------------------------------------------------------------------
Public Function SqlSetting(ByVal Name As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb                          ' held in a variable: a chained CurrentDb is released at once (3420)
    SqlSetting = Nz(db.Properties(Name).Value, "")
End Function

Public Sub SaveSqlSetting(ByVal Name As String, ByVal Value As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(Name).Value = Value
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(Name, dbText, IIf(Len(Value) = 0, " ", Value))
        db.Properties(Name).Value = Value
    End If
End Sub

Public Function DriverInstalled(ByVal Driver As String) As Boolean
    Dim sh As Object, v As Variant
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    v = sh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\" & Driver)
    DriverInstalled = (Err.Number = 0 And CStr(v) = "Installed")
End Function

Public Function SqlDriverList() As String
    ' The SQL Server ODBC drivers of this PC, the newest first ("a;b;c" for a value list).
    Dim names As Variant, i As Long, s As String
    names = Array("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", _
                  "SQL Server Native Client 11.0", "SQL Server")
    For i = 0 To UBound(names)
        If DriverInstalled(names(i)) Then s = s & IIf(Len(s) > 0, ";", "") & names(i)
    Next
    SqlDriverList = s
End Function

Public Function LinkConnect() As String
    ' The connection of the program's tables (Settings stands for all of them): ";DATABASE=..." or "ODBC;...".
    Dim db As DAO.Database
    On Error Resume Next
    Set db = CurrentDb
    LinkConnect = db.TableDefs("Settings").Connect
End Function

Public Function IsLinkedToServer() As Boolean
    IsLinkedToServer = (Left$(LinkConnect(), 5) = "ODBC;")
End Function

Public Function AccessBackEndPath() As String
    ' The .accdb the tables are linked to (or were, before the server).
    Dim c As String
    On Error Resume Next
    c = LinkConnect()
    If Left$(c, 10) = ";DATABASE=" Then
        AccessBackEndPath = Mid$(c, 11)
    Else
        AccessBackEndPath = SqlSetting("SqlAccessPath")
    End If
End Function

'------------------------------------------------------------------------------
' Running SQL on the server (pass-through)
'------------------------------------------------------------------------------
Private Function OdbcError() As String
    Dim e As DAO.Error, s As String
    For Each e In DBEngine.Errors
        If InStr(1, s, e.Description, vbBinaryCompare) = 0 Then s = s & e.Description & vbCrLf
    Next
    If Len(s) = 0 Then s = Err.Description
    OdbcError = Trim$(s)
End Function

Public Function ServerExec(ByVal Connect As String, ByVal Sql As String) As String
    ' Runs SQL on the server. "" = done, else the server's message.
    Dim qd As DAO.QueryDef
    On Error GoTo Failed
    Set qd = CurrentDb.CreateQueryDef("")
    qd.Connect = Connect
    qd.ODBCTimeout = 0                                  ' no time limit (a long copy)
    qd.ReturnsRecords = False
    qd.Sql = Sql
    qd.Execute dbFailOnError
    Exit Function
Failed:
    ServerExec = OdbcError()
End Function

Public Function ServerValue(ByVal Connect As String, ByVal Sql As String, Optional ByVal Timeout As Long = 0) As Variant
    ' The first value a query returns on the server (Null when it fails: m_lastError says why).
    Dim qd As DAO.QueryDef, rs As DAO.Recordset
    ServerValue = Null
    m_lastError = ""
    On Error GoTo Failed
    Set qd = CurrentDb.CreateQueryDef("")
    qd.Connect = Connect
    qd.ODBCTimeout = Timeout
    qd.ReturnsRecords = True
    qd.Sql = Sql
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then ServerValue = rs.Fields(0).Value
    rs.Close
    Exit Function
Failed:
    m_lastError = OdbcError()
End Function

Public Function ServerQueryRows(ByVal Connect As String, ByVal Sql As String) As String
    ' Every row a query returns on the server: rows separated by vbLf, values by vbTab ("" when none or it fails).
    Dim qd As DAO.QueryDef, rs As DAO.Recordset, f As DAO.Field, row As String, out As String
    m_lastError = ""
    On Error GoTo Failed
    Set qd = CurrentDb.CreateQueryDef("")
    qd.Connect = Connect
    qd.ODBCTimeout = 60
    qd.ReturnsRecords = True
    qd.Sql = Sql
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Do Until rs.EOF
        row = ""
        For Each f In rs.Fields
            row = row & IIf(Len(row) > 0, vbTab, "") & Nz(f.Value, "")
        Next
        out = out & IIf(Len(out) > 0, vbLf, "") & row
        rs.MoveNext
    Loop
    rs.Close
    ServerQueryRows = out
    Exit Function
Failed:
    m_lastError = OdbcError()
End Function

Public Function ServerLastError() As String
    ServerLastError = m_lastError
End Function

'------------------------------------------------------------------------------
' Backup and restore of the server database (plan 5C, docs/70); used by modBackup
'------------------------------------------------------------------------------
Public Function SqlName(ByVal Name As String) As String
    SqlName = "[" & Replace(Name, "]", "]]") & "]"
End Function

Public Function ConnectDatabase(ByVal Connect As String) As String
    ' The DATABASE= of an ODBC connection.
    Dim p As Long, q As Long
    p = InStr(1, Connect, ";DATABASE=", vbTextCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 10, Connect, ";")
    If q = 0 Then q = Len(Connect) + 1
    ConnectDatabase = Mid$(Connect, p + 10, q - p - 10)
End Function

Public Function ConnectToDatabase(ByVal Connect As String, ByVal Database As String) As String
    ' The same connection to another database of the server (master for a restore).
    Dim p As Long, q As Long
    p = InStr(1, Connect, ";DATABASE=", vbTextCompare)
    If p = 0 Then
        ConnectToDatabase = Connect
        Exit Function
    End If
    q = InStr(p + 10, Connect, ";")
    If q = 0 Then q = Len(Connect) + 1
    ConnectToDatabase = Left$(Connect, p + 9) & Database & Mid$(Connect, q)
End Function

Public Function BackupDatabaseSql(ByVal Database As String, ByVal BackupFile As String) As String
    ' COPY_ONLY: a backup of the program never disturbs the backup chain an administrator may keep.
    BackupDatabaseSql = "BACKUP DATABASE " & SqlName(Database) & " TO DISK = " & SqlLiteral(BackupFile, 10) & _
                        " WITH COPY_ONLY, INIT, CHECKSUM, NAME = N'RetailStore';"
End Function

Public Function VerifyBackupSql(ByVal BackupFile As String) As String
    VerifyBackupSql = "RESTORE VERIFYONLY FROM DISK = " & SqlLiteral(BackupFile, 10) & " WITH CHECKSUM;"
End Function

Public Function RestoreDatabaseSql(ByVal Database As String, ByVal BackupFile As String) As String
    ' Run on master: every other connection is closed, the database replaced, then opened to all again
    ' (also when the restore fails).
    Dim d As String
    d = SqlName(Database)
    RestoreDatabaseSql = "ALTER DATABASE " & d & " SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" & vbLf & _
        "BEGIN TRY" & vbLf & _
        "    RESTORE DATABASE " & d & " FROM DISK = " & SqlLiteral(BackupFile, 10) & " WITH REPLACE;" & vbLf & _
        "END TRY" & vbLf & _
        "BEGIN CATCH" & vbLf & _
        "    ALTER DATABASE " & d & " SET MULTI_USER;" & vbLf & _
        "    THROW;" & vbLf & _
        "END CATCH;" & vbLf & _
        "ALTER DATABASE " & d & " SET MULTI_USER;"
End Function

Public Function ServerBackupHistorySql() As String
    ' The program's backups of this database the server remembers (msdb), newest first.
    ServerBackupHistorySql = "SELECT TOP 100 f.physical_device_name, CONVERT(VARCHAR(19), s.backup_finish_date, 120), " & _
        "CAST(s.backup_size / 1048576.0 AS DECIMAL(12, 1)) FROM msdb.dbo.backupset AS s " & _
        "INNER JOIN msdb.dbo.backupmediafamily AS f ON f.media_set_id = s.media_set_id " & _
        "WHERE s.database_name = DB_NAME() AND s.type = 'D' AND s.name = N'RetailStore' ORDER BY s.backup_finish_date DESC;"
End Function

Public Function ServerLiveSince(ByVal Connect As String) As Variant
    ' When the program started to work on this server database (Settings.ServerLiveSince), Null = not yet:
    ' from then on the server holds the real data (docs/70).
    ServerLiveSince = ServerValue(Connect, "SELECT ServerLiveSince FROM [dbo].[Settings] WHERE SettingID = 1")
End Function

Public Function TestServerConnection(ByVal Connect As String) As String
    ' "" when the server answers and the database exists, else why not.
    Dim v As Variant
    v = ServerValue(Connect, "SELECT DB_NAME()", 15)
    If IsNull(v) Then
        TestServerConnection = " ⁄–¯— «·« ’«· »«·Œ«œ„:" & vbCrLf & m_lastError
    End If
End Function

Public Function ServerTablesMissing(ByVal Connect As String) As Long
    ' How many tables of the program the server database does not have (the script was not run).
    Dim names As Variant, i As Long, n As Long
    names = Split(SchemaTableList(), ",")                       ' modBuildSchema
    For i = 0 To UBound(names)
        If IsNull(ServerValue(Connect, "SELECT OBJECT_ID(N'dbo." & names(i) & "', N'U')")) Then n = n + 1
    Next
    ServerTablesMissing = n
End Function

Private Function ReadUtf8File(ByVal Path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                                ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile Path
    ReadUtf8File = st.ReadText
    st.Close
    If Left$(ReadUtf8File, 1) = ChrW(&HFEFF) Then ReadUtf8File = Mid$(ReadUtf8File, 2)
End Function

Public Function DefaultScriptPath() As String
    Dim p As String
    p = CurrentProject.Path & "\sql\RetailStore_Schema.sql"
    If Len(Dir$(p)) = 0 Then p = CurrentProject.Path & "\RetailStore_Schema.sql"
    DefaultScriptPath = p
End Function

Public Function QueriesScriptPath(ByVal SchemaScriptPath As String) As String
    ' The views of the saved queries (docs/71): RetailStore_Queries.sql in the folder of the database script.
    QueriesScriptPath = Left$(SchemaScriptPath, InStrRev(SchemaScriptPath, "\")) & "RetailStore_Queries.sql"
End Function

Public Function RunSchemaScript(ByVal Connect As String, ByVal ScriptPath As String, ByRef Done As Long) As String
    ' The database script (docs/67), batch by batch: a new database is created, an existing one upgraded.
    Dim batches As Variant, i As Long, msg As String
    Done = 0
    If Not HasPermission("SQL_SERVER") Then
        RunSchemaScript = "·«  „·ﬂ ’·«ÕÌ… ﬁ«⁄œ… SQL Server."
        Exit Function
    End If
    If Len(Dir$(ScriptPath)) = 0 Then
        RunSchemaScript = "„·› «·”ﬂ—»  €Ì— „ÊÃÊœ: " & ScriptPath
        Exit Function
    End If
    batches = SplitSqlBatches(ReadUtf8File(ScriptPath))
    For i = 0 To UBound(batches)
        msg = ServerExec(Connect, batches(i))
        If Len(msg) > 0 Then
            RunSchemaScript = " Êﬁ› «·”ﬂ—»  ›Ì «·Ã“¡ " & (i + 1) & " „‰ " & (UBound(batches) + 1) & ":" & vbCrLf & msg
            Exit Function
        End If
        Done = Done + 1
        SysCmd acSysCmdSetStatus, Tr(" ÃÂÌ“ «·ﬁ«⁄œ…: ") & Done & " / " & (UBound(batches) + 1)
    Next
    SysCmd acSysCmdClearStatus
    LogAction "SQL_SCHEMA", "SqlServer", SqlSetting("SqlDatabase"), CStr(Done)
End Function

'------------------------------------------------------------------------------
' Moving the data
'------------------------------------------------------------------------------
Public Function CopyTable(ByVal Connect As String, ByVal TableName As String, ByRef Copied As Long) As String
    ' Every row of the Access table to the (empty) server table, with the same numbers.
    Dim rs As DAO.Recordset, f As DAO.Field, cols As String, row As String, rows As String, n As Long
    Dim hasIdentity As Boolean, msg As String
    Copied = 0
    hasIdentity = (Nz(ServerValue(Connect, "SELECT OBJECTPROPERTY(OBJECT_ID(N'dbo." & TableName & "'), 'TableHasIdentity')"), 0) = 1)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & TableName & "]", dbOpenSnapshot)
    For Each f In rs.Fields
        cols = cols & IIf(Len(cols) > 0, ", ", "") & "[" & f.Name & "]"
    Next
    Do Until rs.EOF
        row = ""
        For Each f In rs.Fields
            row = row & IIf(Len(row) > 0, ", ", "") & SqlLiteral(f.Value, f.Type)
        Next
        rows = rows & IIf(Len(rows) > 0, "," & vbLf, "") & "(" & row & ")"
        n = n + 1
        rs.MoveNext
        If n = BATCH_ROWS Or Len(rows) > BATCH_CHARS Or rs.EOF Then
            msg = ServerExec(Connect, InsertBatchSql(TableName, cols, rows, hasIdentity))
            If Len(msg) > 0 Then
                CopyTable = TableName & ": " & msg
                rs.Close
                Exit Function
            End If
            Copied = Copied + n
            n = 0
            rows = ""
        End If
    Loop
    rs.Close
End Function

Public Function TransferAllData(ByVal Connect As String, ByRef Report As String) As String
    ' The whole back-end to the server. The server tables are emptied first: run it once, before working
    ' on the server (or again to start over from the Access file).
    Dim names As Variant, i As Long, msg As String, copied As Long, total As Long
    Report = ""
    Calendar = vbCalGreg                                       ' dates written yyyymmdd in the Gregorian calendar
    If Not HasPermission("SQL_SERVER") Then
        TransferAllData = "·«  „·ﬂ ’·«ÕÌ… ﬁ«⁄œ… SQL Server."
        Exit Function
    End If
    If IsLinkedToServer() Then
        TransferAllData = "«·»—‰«„Ã Ì⁄„· «·¬‰ ⁄·Ï «·Œ«œ„: «·‰ﬁ· ÌﬂÊ‰ „‰ „·› Access. «—Ã⁄ ≈·Ï „·› Access √Ê·«."
        Exit Function
    End If
    msg = TestServerConnection(Connect)
    If Len(msg) = 0 Then
        If ServerTablesMissing(Connect) > 0 Then msg = "ﬁ«⁄œ… «·Œ«œ„ ·Ì” ›ÌÂ« ﬂ· «·Ãœ«Ê·: ‘€¯· ´ ÃÂÌ“ «·ﬁ«⁄œ…ª √Ê·«."
    End If
    If Len(msg) = 0 Then                                        ' never over the real data (docs/70)
        If Not IsNull(ServerLiveSince(Connect)) Then
            msg = "ﬁ«⁄œ… «·Œ«œ„ ›ÌÂ« «·»Ì«‰«  «·›⁄·Ì… ··»—‰«„Ã „‰– " & GDate(ServerLiveSince(Connect), True) & _
                  "° Ê«·‰ﬁ· ÌÕ–›Â«." & vbCrLf & "«—»ÿ Â–« «·ÃÂ«“ »«·Œ«œ„ »“— ´«· ‘€Ì· ⁄·Ï «·Œ«œ„ª »œ· «·‰ﬁ·." & _
                  vbCrLf & "Ê··»œ¡ „‰ „·› Access „‰ ÃœÌœ: ´«·—ÃÊ⁄ ≈·Ï Accessª „⁄ Ã⁄· „·› Access ÂÊ «·»Ì«‰«  «·›⁄·Ì…."
        End If
    End If
    If Len(msg) > 0 Then
        TransferAllData = msg
        Exit Function
    End If
    names = Split(SchemaTableList(), ",")
    For i = 0 To UBound(names)                                  ' rules off, then empty (any order)
        msg = ServerExec(Connect, "ALTER TABLE [dbo].[" & names(i) & "] NOCHECK CONSTRAINT ALL;")
        If Len(msg) > 0 Then GoTo Failed
    Next
    For i = 0 To UBound(names)
        msg = ServerExec(Connect, "DELETE FROM [dbo].[" & names(i) & "];")
        If Len(msg) > 0 Then GoTo Failed
    Next
    For i = 0 To UBound(names)
        SysCmd acSysCmdSetStatus, Tr("‰ﬁ· «·»Ì«‰« : ") & names(i) & " (" & (i + 1) & " / " & (UBound(names) + 1) & ")"
        msg = CopyTable(Connect, names(i), copied)
        If Len(msg) > 0 Then GoTo Failed
        total = total + copied
        Report = Report & names(i) & ": " & copied & vbCrLf
    Next
    For i = 0 To UBound(names)                                  ' every row checked against every rule again
        msg = ServerExec(Connect, "ALTER TABLE [dbo].[" & names(i) & "] WITH CHECK CHECK CONSTRAINT ALL;")
        If Len(msg) > 0 Then
            msg = names(i) & ": »Ì«‰«  ·«  Õﬁﬁ ‘—Êÿ «·Œ«œ„ √Ê ⁄·«ﬁ« Â:" & vbCrLf & msg
            GoTo Failed
        End If
    Next
    SysCmd acSysCmdClearStatus
    LogAction "SQL_TRANSFER", "SqlServer", SqlSetting("SqlDatabase"), CStr(total)
    Exit Function
Failed:
    SysCmd acSysCmdClearStatus
    TransferAllData = msg
End Function

Public Function CompareTable(ByVal Connect As String, ByVal TableName As String, ByRef AccessRows As Long, _
                             ByRef ServerRows As Long, ByRef AccessTotal As Double, ByRef ServerTotal As Double) As Boolean
    ' Same number of rows, the same total of every amount (Currency) field, and the same sum of the first
    ' character code of every text field (Arabic that reached the server as "?" is a difference).
    Dim rs As DAO.Recordset, f As DAO.Field, accSql As String, srvSql As String, accText As String, srvText As String
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM [" & TableName & "]", dbOpenSnapshot)
    For Each f In rs.Fields
        If f.Type = dbCurrency Then
            accSql = accSql & " + Nz(Sum([" & f.Name & "]), 0)"
            srvSql = srvSql & " + ISNULL(SUM(CAST([" & f.Name & "] AS FLOAT)), 0)"
        ElseIf f.Type = dbText Or f.Type = dbMemo Then
            accText = accText & " + Nz(Sum(AscW(Nz([" & f.Name & "], '') & ' ')), 0)"
            srvText = srvText & " + ISNULL(SUM(CAST(UNICODE(SUBSTRING(ISNULL([" & f.Name & "], N'') + N' ', 1, 1)) AS FLOAT)), 0)"
        End If
    Next
    rs.Close
    AccessRows = Nz(DbValue("SELECT COUNT(*) FROM [" & TableName & "]"), 0)
    AccessTotal = CDbl(Nz(DbValue("SELECT 0" & accSql & " FROM [" & TableName & "]"), 0))
    ServerRows = Nz(ServerValue(Connect, "SELECT COUNT(*) FROM [dbo].[" & TableName & "]"), -1)
    ServerTotal = CDbl(Nz(ServerValue(Connect, "SELECT 0" & srvSql & " FROM [dbo].[" & TableName & "]"), 0))
    CompareTable = (AccessRows = ServerRows) And SameAmount(AccessTotal, ServerTotal)
    If CompareTable And Len(accText) > 0 Then
        CompareTable = SameAmount(Nz(DbValue("SELECT 0" & accText & " FROM [" & TableName & "]"), 0), _
                                  Nz(ServerValue(Connect, "SELECT 0" & srvText & " FROM [dbo].[" & TableName & "]"), -1))
    End If
End Function

Public Function CompareAllData(ByVal Connect As String, ByRef Rows As String, ByRef Differences As Long) As String
    ' Every table compared; Rows = a value list for the screen (table; Access rows; server rows; Access total;
    ' server total; result).
    Dim names As Variant, i As Long, aRows As Long, sRows As Long, aTot As Double, sTot As Double, ok As Boolean
    Differences = 0
    Rows = ""
    If IsLinkedToServer() Then
        CompareAllData = "«·»—‰«„Ã Ì⁄„· «·¬‰ ⁄·Ï «·Œ«œ„: «·„ÿ«»ﬁ… »Ì‰ „·› Access Ê«·Œ«œ„  ﬂÊ‰ ﬁ»· «·—»ÿ."
        Exit Function
    End If
    CompareAllData = TestServerConnection(Connect)
    If Len(CompareAllData) > 0 Then Exit Function
    names = Split(SchemaTableList(), ",")
    For i = 0 To UBound(names)
        SysCmd acSysCmdSetStatus, Tr("«·„ÿ«»ﬁ…: ") & names(i)
        ok = CompareTable(Connect, names(i), aRows, sRows, aTot, sTot)
        If Not ok Then Differences = Differences + 1
        Rows = Rows & ";""" & names(i) & """;""" & aRows & """;""" & sRows & """;""" & Format$(aTot, "#,##0.00") & _
               """;""" & Format$(sTot, "#,##0.00") & """;""" & IIf(ok, "„ÿ«»ﬁ", "„Œ ·›") & """"
    Next
    SysCmd acSysCmdClearStatus
End Function

'------------------------------------------------------------------------------
' Linking the program to the server, and back
'------------------------------------------------------------------------------
Public Sub RelinkTable(ByVal db As DAO.Database, ByVal TableName As String, ByVal Connect As String, _
                        ByVal Source As String, ByVal SavePassword As Boolean)
    Dim tdf As DAO.TableDef
    On Error Resume Next
    db.TableDefs.Delete TableName
    On Error GoTo 0
    Set tdf = db.CreateTableDef(TableName)
    If SavePassword Then tdf.Attributes = dbAttachSavePWD
    tdf.Connect = Connect
    tdf.SourceTableName = Source
    db.TableDefs.Append tdf
End Sub

Public Function LinkToServer(ByVal Connect As String) As String
    ' The program's tables linked to the server: the first time when its data matches the Access file (then the
    ' server is marked as holding the real data); a server already live is linked as it is (another PC, a new
    ' front-end).
    Dim names As Variant, i As Long, rows As String, diffs As Long, db As DAO.Database, msg As String
    Dim live As Boolean
    If Not HasPermission("SQL_SERVER") Then
        LinkToServer = "·«  „·ﬂ ’·«ÕÌ… ﬁ«⁄œ… SQL Server."
        Exit Function
    End If
    If IsLinkedToServer() Then
        LinkToServer = "«·»—‰«„Ã Ì⁄„· ⁄·Ï «·Œ«œ„ »«·›⁄·."
        Exit Function
    End If
    msg = TestServerConnection(Connect)
    If Len(msg) = 0 Then
        If ServerTablesMissing(Connect) > 0 Then msg = "ﬁ«⁄œ… «·Œ«œ„ ·Ì” ›ÌÂ« ﬂ· «·Ãœ«Ê·: ‘€¯· ´ ÃÂÌ“ «·ﬁ«⁄œ…ª √Ê·«."
    End If
    If Len(msg) = 0 Then msg = ServerQueriesProblem(Connect)       ' the views of the saved queries (docs/71)
    If Len(msg) = 0 Then
        live = Not IsNull(ServerLiveSince(Connect))
        If Not live Then
            msg = CompareAllData(Connect, rows, diffs)
            If Len(msg) = 0 And diffs > 0 Then msg = "»Ì«‰«  «·Œ«œ„ ·«  ÿ«»ﬁ „·› Access ›Ì " & diffs & " ÃœÊ·: «‰ﬁ· «·»Ì«‰«  À„ ÿ«»ﬁÂ« √Ê·«."
            If Len(msg) = 0 Then msg = ServerExec(Connect, "UPDATE [dbo].[Settings] SET ServerLiveSince = GETDATE() WHERE SettingID = 1;")
            If Len(msg) = 0 Then                                ' the Access file knows it is no longer the real data
                Set db = CurrentDb                              ' (a file BuildSchema has not upgraded yet: no warning)
                On Error Resume Next
                db.Execute "UPDATE Settings SET ServerLiveSince = Now() WHERE SettingID = 1", dbFailOnError + dbSeeChanges
                On Error GoTo 0
            End If
        End If
    End If
    If Len(msg) > 0 Then
        LinkToServer = msg
        Exit Function
    End If
    SaveSqlSetting "SqlAccessPath", AccessBackEndPath()
    Set db = CurrentDb
    names = Split(SchemaTableList(), ",")
    For i = 0 To UBound(names)
        RelinkTable db, names(i), Connect, "dbo." & names(i), InStr(1, Connect, "UID=", vbTextCompare) > 0
    Next
    db.TableDefs.Refresh
    LogAction "SQL_LINK", "SqlServer", SqlSetting("SqlDatabase"), SqlSetting("SqlServer")
    LinkToServer = UseServerQueries(Connect)                    ' the saved queries on the server views
End Function

Public Function LinkToAccess(Optional ByVal BackEndPath As String = "", Optional ByVal ReleaseServer As Boolean = False) As String
    ' Back to the .accdb file the program used before the server (its data is as it was then). ReleaseServer:
    ' the Access file is the real data again, so the server may receive a new transfer.
    Dim names As Variant, i As Long, db As DAO.Database, msg As String
    If Len(BackEndPath) = 0 Then BackEndPath = SqlSetting("SqlAccessPath")
    If Len(BackEndPath) = 0 Or Len(Dir$(BackEndPath)) = 0 Then
        LinkToAccess = "„·› »Ì«‰«  Access €Ì— „ÊÃÊœ: " & BackEndPath
        Exit Function
    End If
    If ReleaseServer And IsLinkedToServer() Then
        msg = ServerExec(LinkConnect(), "UPDATE [dbo].[Settings] SET ServerLiveSince = NULL WHERE SettingID = 1;")
        If Len(msg) > 0 Then
            LinkToAccess = msg
            Exit Function
        End If
    End If
    msg = UseAccessQueries()                                    ' the saved queries in Access again
    If Len(msg) > 0 Then
        LinkToAccess = msg
        Exit Function
    End If
    Set db = CurrentDb
    names = Split(SchemaTableList(), ",")
    For i = 0 To UBound(names)
        RelinkTable db, names(i), ";DATABASE=" & BackEndPath, names(i), False
    Next
    db.TableDefs.Refresh
    If ReleaseServer Then
        On Error Resume Next
        db.Execute "UPDATE Settings SET ServerLiveSince = Null WHERE SettingID = 1", dbFailOnError + dbSeeChanges
        On Error GoTo 0
    End If
    LogAction "SQL_UNLINK", "SqlServer", BackEndPath
End Function

Public Function StaleAccessWarning() As String
    ' A front-end on the Access file after the real data moved to SQL Server (a rebuilt front-end, a temporary
    ' return): "" when not the case, else the warning of the login screen.
    Dim v As Variant
    If IsLinkedToServer() Then Exit Function
    On Error Resume Next
    v = DbValue("SELECT ServerLiveSince FROM Settings WHERE SettingID = 1")
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    StaleAccessWarning = "«·»Ì«‰«  «·›⁄·Ì… ··»—‰«„Ã ⁄·Ï SQL Server „‰– " & GDate(v, True) & "° ÊÂ–« «·ÃÂ«“ Ì⁄„· ⁄·Ï „·› Access «·ﬁœÌ„." & _
                         vbCrLf & "”Ã¯· «·œŒÊ· À„: ´‰”Œ… «Õ Ì«ÿÌ…ª À„ ´SQL Serverª À„ ´«· ‘€Ì· ⁄·Ï «·Œ«œ„ª."
End Function

Public Function CheckServerAtStartup() As String
    ' "" when the program works on its Access file or the server answers; else the message to show.
    Dim c As String
    If Not IsLinkedToServer() Then Exit Function
    c = LinkConnect()
    If IsNull(ServerValue(c, "SELECT 1", 10)) Then
        CheckServerAtStartup = "·« Ì—œ Œ«œ„ «·»Ì«‰«  " & SqlSetting("SqlServer") & "." & vbCrLf & _
                               " √ﬂœ „‰  ‘€Ì· «·Œ«œ„ Ê«·‘»ﬂ…° À„ «› Õ «·»—‰«„Ã „‰ ÃœÌœ." & vbCrLf & vbCrLf & m_lastError
    End If
End Function

'------------------------------------------------------------------------------
' Screen frmSqlServer
'------------------------------------------------------------------------------
Private Function ScreenConnect(ByVal frm As Access.Form) As String
    Dim user As String
    If Nz(frm!cboAuth.Value, "WINDOWS") = "SQL" Then user = Nz(frm!txtUser.Value, "")
    ScreenConnect = ServerConnect(Nz(frm!txtServer.Value, ""), Nz(frm!txtDatabase.Value, ""), user, _
                                  Nz(frm!txtPassword.Value, ""), Nz(frm!cboDriver.Value, "SQL Server"))
End Function

Private Function ScreenProblem(ByVal frm As Access.Form) As String
    If Len(Nz(frm!txtServer.Value, "")) = 0 Or Len(Nz(frm!txtDatabase.Value, "")) = 0 Then
        ScreenProblem = "«ﬂ » «”„ «·Œ«œ„ Êﬁ«⁄œ… «·»Ì«‰« ."
    ElseIf IsNull(frm!cboDriver.Value) Then
        ScreenProblem = "·« ÌÊÃœ »—‰«„Ã  ‘€Ì· SQL Server (ODBC) ⁄·Ï Â–« «·ÃÂ«“: À»¯  ODBC Driver 18 for SQL Server."
    ElseIf Nz(frm!cboAuth.Value, "WINDOWS") = "SQL" And Len(Nz(frm!txtUser.Value, "")) = 0 Then
        ScreenProblem = "«ﬂ » «”„ „” Œœ„ SQL Server."
    End If
End Function

Private Sub SaveScreenSettings(ByVal frm As Access.Form)
    SaveSqlSetting "SqlServer", Nz(frm!txtServer.Value, "")
    SaveSqlSetting "SqlDatabase", Nz(frm!txtDatabase.Value, "")
    SaveSqlSetting "SqlUser", IIf(Nz(frm!cboAuth.Value, "WINDOWS") = "SQL", Nz(frm!txtUser.Value, ""), "")
    SaveSqlSetting "SqlDriver", Nz(frm!cboDriver.Value, "")
End Sub

Public Sub SqlServerLoad(ByVal frm As Access.Form)
    Dim drivers As String
    drivers = SqlDriverList()
    frm!cboDriver.RowSource = drivers
    frm!txtServer.Value = IIf(Len(SqlSetting("SqlServer")) > 0, SqlSetting("SqlServer"), ".\SQLEXPRESS")
    frm!txtDatabase.Value = IIf(Len(SqlSetting("SqlDatabase")) > 0, SqlSetting("SqlDatabase"), "RetailStore")
    frm!cboAuth.Value = IIf(Len(SqlSetting("SqlUser")) > 0, "SQL", "WINDOWS")
    frm!txtUser.Value = SqlSetting("SqlUser")
    If Len(SqlSetting("SqlDriver")) > 0 Then
        frm!cboDriver.Value = SqlSetting("SqlDriver")
    ElseIf Len(drivers) > 0 Then
        frm!cboDriver.Value = Split(drivers, ";")(0)
    End If
    frm!txtScript.Value = DefaultScriptPath()
    SqlServerAuthChanged frm
    SqlServerShowState frm
End Sub

Public Sub SqlServerAuthChanged(ByVal frm As Access.Form)
    frm!txtUser.Enabled = (Nz(frm!cboAuth.Value, "WINDOWS") = "SQL")
    frm!txtPassword.Enabled = frm!txtUser.Enabled
End Sub

Public Sub SqlServerShowState(ByVal frm As Access.Form)
    If IsLinkedToServer() Then
        frm!lblState.Caption = Tr("«·»—‰«„Ã Ì⁄„· «·¬‰ ⁄·Ï «·Œ«œ„: " & SqlSetting("SqlServer") & " / " & SqlSetting("SqlDatabase"))
        frm!lblState.ForeColor = CLR_PRIMARY
    Else
        frm!lblState.Caption = Tr("«·»—‰«„Ã Ì⁄„· «·¬‰ ⁄·Ï „·› Access: " & AccessBackEndPath())
        frm!lblState.ForeColor = CLR_MUTED
    End If
End Sub

Public Sub SqlServerTest(ByVal frm As Access.Form)
    Dim msg As String
    msg = ScreenProblem(frm)
    If Len(msg) = 0 Then msg = TestServerConnection(ScreenConnect(frm))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    SaveScreenSettings frm
    ShowInfo " „ «·« ’«· »«·Œ«œ„." & vbCrLf & "Ãœ«Ê· «·»—‰«„Ã «·‰«ﬁ’… ›Ì «·ﬁ«⁄œ…: " & ServerTablesMissing(ScreenConnect(frm))
End Sub

Public Sub SqlServerBrowseScript(ByVal frm As Access.Form)
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)                          ' msoFileDialogFilePicker
    fd.Title = Tr("”ﬂ—»  ﬁ«⁄œ… «·»Ì«‰« ")
    fd.Filters.Clear
    fd.Filters.Add "SQL", "*.sql"
    If fd.Show Then frm!txtScript.Value = fd.SelectedItems(1)
End Sub

Public Sub SqlServerSchema(ByVal frm As Access.Form)
    Dim msg As String, n As Long, views As Long
    msg = ScreenProblem(frm)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Not AskYesNo(" ÃÂÌ“ ﬁ«⁄œ… «·Œ«œ„ „‰ «·”ﬂ—» ø" & vbCrLf & "ﬁ«⁄œ… ÃœÌœ…  ı‰‘√ Ãœ«Ê·Â«° Êﬁ«⁄œ… „ÊÃÊœ…  ı—ﬁ¯Ï Ê·«  ıÕ–› »Ì«‰« Â«.") Then Exit Sub
    msg = RunSchemaScript(ScreenConnect(frm), Nz(frm!txtScript.Value, ""), n)
    If Len(msg) = 0 Then                                        ' then the views of the saved queries (docs/71)
        msg = RunSchemaScript(ScreenConnect(frm), QueriesScriptPath(Nz(frm!txtScript.Value, "")), views)
        n = n + views
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        SaveScreenSettings frm
        ShowInfo " „  ÃÂÌ“ «·ﬁ«⁄œ… (" & n & " Ã“¡« „‰ «·”ﬂ—» )."
    End If
End Sub

Public Sub SqlServerTransfer(ByVal frm As Access.Form)
    Dim msg As String, rep As String
    msg = ScreenProblem(frm)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Not AskYesNo("‰ﬁ· ﬂ· «·»Ì«‰«  „‰ „·› Access ≈·Ï «·Œ«œ„ø" & vbCrLf & vbCrLf & _
                    " ıÕ–› ﬂ· «·»Ì«‰«  «·„ÊÃÊœ… ›Ì ﬁ«⁄œ… «·Œ«œ„ √Ê·«. ·« Ì⁄„· √Õœ ⁄·Ï «·»—‰«„Ã √À‰«¡ «·‰ﬁ·.") Then Exit Sub
    DoCmd.Hourglass True
    msg = TransferAllData(ScreenConnect(frm), rep)
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    SaveScreenSettings frm
    ShowInfo " „ ‰ﬁ· «·»Ì«‰« . «·¬‰ ´„ÿ«»ﬁ… «·»Ì«‰« ª."
    SqlServerCompare frm
End Sub

Public Sub SqlServerCompare(ByVal frm As Access.Form)
    Dim msg As String, rows As String, diffs As Long
    msg = ScreenProblem(frm)
    If Len(msg) = 0 Then
        DoCmd.Hourglass True
        msg = CompareAllData(ScreenConnect(frm), rows, diffs)
        DoCmd.Hourglass False
    End If
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    frm!lstResult.RowSource = Tr("""«·ÃœÊ·"";""”Ã·«  Access"";""”Ã·«  «·Œ«œ„"";""„Ã„Ê⁄ Access"";""„Ã„Ê⁄ «·Œ«œ„"";""«·‰ ÌÃ…""" & rows)
    frm!lblResult.Caption = Tr(IIf(diffs = 0, "ﬂ· «·Ãœ«Ê· „ÿ«»ﬁ….", "Ãœ«Ê· „Œ ·›…: " & diffs))
    frm!lblResult.ForeColor = IIf(diffs = 0, CLR_PRIMARY, CLR_DANGER)
End Sub

Public Sub SqlServerLink(ByVal frm As Access.Form)
    Dim msg As String
    msg = ScreenProblem(frm)
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    If Not AskYesNo("—»ÿ «·»—‰«„Ã ⁄·Ï Â–« «·ÃÂ«“ »«·Œ«œ„ø" & vbCrLf & vbCrLf & _
                    "√Ê· ÃÂ«“:  ıÿ«»Ûﬁ «·»Ì«‰«  √Ê·«° À„ Ì’Ì— «·Œ«œ„ ÂÊ «·»Ì«‰«  «·›⁄·Ì…. «·√ÃÂ“… «· «·Ì…  ı—»ÿ „»«‘—…." & vbCrLf & _
                    "Ì»ﬁÏ „·› Access ﬂ„« ÂÊ ··—ÃÊ⁄ ≈·ÌÂ.") Then Exit Sub
    DoCmd.Hourglass True
    msg = LinkToServer(ScreenConnect(frm))
    DoCmd.Hourglass False
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        SaveScreenSettings frm
        ShowInfo "«·»—‰«„Ã „—»Êÿ »«·Œ«œ„ «·¬‰. √€·ﬁÂ Ê«› ÕÂ „‰ ÃœÌœ."
    End If
    SqlServerShowState frm
End Sub

Public Sub SqlServerBackToAccess(ByVal frm As Access.Form)
    Dim msg As String, release As Boolean
    If Not IsLinkedToServer() Then
        ShowInfo "«·»—‰«„Ã Ì⁄„· ⁄·Ï „·› Access »«·›⁄·."
        Exit Sub
    End If
    If Not AskYesNo("«·—ÃÊ⁄ ≈·Ï „·› Access:" & vbCrLf & SqlSetting("SqlAccessPath") & "ø" & vbCrLf & vbCrLf & _
                    "„« ”ıÃ¯· ⁄·Ï «·Œ«œ„ »⁄œ «·—»ÿ ·« ÌŸÂ— ›Ì „·› Access.") Then Exit Sub
    release = AskYesNo("Â· Ì’Ì— „·› Access ÂÊ «·»Ì«‰«  «·›⁄·Ì… „‰ ÃœÌœø" & vbCrLf & vbCrLf & _
                       "‰⁄„: Ìı”„Õ »‰ﬁ· «·»Ì«‰«  ≈·Ï «·Œ«œ„ „‰ ÃœÌœ° Êﬂ· «·√ÃÂ“… ÌÃ» √‰  —Ã⁄ ≈·Ï Access." & vbCrLf & _
                       "·«: —ÃÊ⁄ „ƒﬁ  ·Â–« «·ÃÂ«“ ›ﬁÿ° Ê«·Œ«œ„ Ì»ﬁÏ «·»Ì«‰«  «·›⁄·Ì….")
    msg = LinkToAccess(, release)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo "«·»—‰«„Ã Ì⁄„· ⁄·Ï „·› Access «·¬‰. √€·ﬁÂ Ê«› ÕÂ „‰ ÃœÌœ."
    End If
    SqlServerShowState frm
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): only the parts that need no server
'------------------------------------------------------------------------------
Private Sub CheckSql(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                     ByRef Report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        Report = Report & "- " & Title & vbCrLf
        Debug.Print "[X]  " & Title
    End If
End Sub

Public Function TestSqlServer() As Boolean
    Dim passed As Long, failed As Long, Report As String, b As Variant, c As String
    Calendar = vbCalGreg
    Debug.Print "=== TestSqlServer  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckSql SqlLiteral(Null, 10) = "NULL" And SqlLiteral(True, 1) = "1" And SqlLiteral(False, 1) = "0", _
             "«·ﬁÌ„ «·›«—€… Ê‰⁄„/·«", passed, failed, Report
    CheckSql SqlLiteral("O'Brien", 10) = "N'O''Brien'" And SqlLiteral(12.5, 5) = "12.5" And SqlLiteral(-3, 4) = "-3", _
             "«·‰’Ê’ Ê«·√—ﬁ«„ »‰ﬁÿ… ⁄‘—Ì…", passed, failed, Report
    CheckSql SqlLiteral(DateSerial(2026, 3, 9) + TimeSerial(14, 5, 7), 8) = "'20260309 14:05:07'", _
             "«· «—ÌŒ yyyymmdd", passed, failed, Report
    c = ServerConnect(".\SQLEXPRESS", "RetailStore", "", "", "ODBC Driver 18 for SQL Server")
    CheckSql InStr(c, "Trusted_Connection=Yes;") > 0 And InStr(c, "TrustServerCertificate=Yes;") > 0, _
             "«·« ’«· »Õ”«» Windows", passed, failed, Report
    b = SplitSqlBatches("A" & vbCrLf & "GO" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "go" & vbCrLf)
    CheckSql UBound(b) = 1, " ﬁ”Ì„ «·”ﬂ—»  ⁄‰œ GO", passed, failed, Report
    CheckSql UBound(Split(SchemaTableList(), ",")) > 50, "ﬁ«∆„… Ãœ«Ê· «·»—‰«„Ã", passed, failed, Report
    CheckSql Len(SqlDriverList()) >= 0 And (IsLinkedToServer() Or Len(AccessBackEndPath()) > 0), _
             "„·› «·»Ì«‰«  «·Õ«·Ì", passed, failed, Report
    ' plan 5C (docs/70)
    CheckSql SqlName("Retail]Store") = "[Retail]]Store]", "«”„ «·ﬁ«⁄œ… »Ì‰ √ﬁÊ«”", passed, failed, Report
    c = ServerConnect("srv", "RetailStore", "", "", "ODBC Driver 18 for SQL Server")
    CheckSql ConnectDatabase(c) = "RetailStore" And ConnectDatabase(ConnectToDatabase(c, "master")) = "master" And _
             InStr(ConnectToDatabase(c, "master"), "TrustServerCertificate=Yes;") > 0, _
             "«·« ’«· »ﬁ«⁄œ… master ··«” ⁄«œ…", passed, failed, Report
    CheckSql BackupDatabaseSql("RetailStore", "D:\B\x.bak") = "BACKUP DATABASE [RetailStore] TO DISK = N'D:\B\x.bak' " & _
             "WITH COPY_ONLY, INIT, CHECKSUM, NAME = N'RetailStore';", "√„— «·‰”Œ «·«Õ Ì«ÿÌ ⁄·Ï «·Œ«œ„", passed, failed, Report
    CheckSql InStr(RestoreDatabaseSql("RetailStore", "x.bak"), "SET SINGLE_USER WITH ROLLBACK IMMEDIATE") > 0 And _
             InStrRev(RestoreDatabaseSql("RetailStore", "x.bak"), "END CATCH;" & vbLf & "ALTER DATABASE [RetailStore] SET MULTI_USER;") > 0, _
             "√„— «·«” ⁄«œ… Ì⁄Ìœ «·ﬁ«⁄œ… ·ﬂ· «·„” Œœ„Ì‰", passed, failed, Report
    CheckSql QueriesScriptPath("D:\Shop\sql\RetailStore_Schema.sql") = "D:\Shop\sql\RetailStore_Queries.sql", _
             "”ﬂ—»  «” ⁄·«„«  «·Œ«œ„ »Ã«‰» ”ﬂ—»  «·ﬁ«⁄œ…", passed, failed, Report
    If IsLinkedToServer() Then                                  ' on the server: the real data is marked
        CheckSql Not IsNull(ServerLiveSince(LinkConnect())), "«·Œ«œ„ „⁄·Û¯„ »√‰Â «·»Ì«‰«  «·›⁄·Ì…", passed, failed, Report
        CheckSql Len(StaleAccessWarning()) = 0, "·«  Õ–Ì— „·› Access «·ﬁœÌ„ ⁄·Ï «·Œ«œ„", passed, failed, Report
    End If
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  SQL Server ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestSqlServer"
        TestSqlServer = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, "TestSqlServer"
    End If
End Function
