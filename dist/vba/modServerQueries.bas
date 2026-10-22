Attribute VB_Name = "modServerQueries"
'==============================================================================
' modServerQueries  -  Retail Store Management System
'
' Plan 5D (docs/71): on SQL Server the saved queries run on the server. dist\sql\RetailStore_Queries.sql
' (tools/gen_sql_views.py) makes each saved query a view of the same name: [dbo] for an Arabic front-end,
' [en] for an English one. Linked to the server, each saved query of the front-end becomes
'     SELECT * FROM [sv_Name] ORDER BY ...
' over the linked view sv_Name, so the server joins and totals and Access receives only the result (the
' journal entries of SyncJournal, the trial balance, the statements, every report). Its Access SQL is kept
' in the query property AccessSql and comes back with the link to the Access file.
'
' The parameters of the queries (QDate / QLong, modQueryParams) are also sent to the server: table
' dbo.QueryParams, one row per Access process (HOST_NAME() + HOST_ID(), the process id of the client);
' the views read them from dbo.QueryParamsNow. Every parameter is set through modQueryParams.
'   UseServerQueries   the saved queries on the server views (LinkToServer)
'   UseAccessQueries   back to their Access SQL (LinkToAccess)
'   SendQueryParams    the parameters just set, to the server (when linked to it)
'   ServerQueriesProblem  "" when the server has the views of this front-end (the same version)
' Generated data: modServerQueryData (the list of the queries, their ORDER BY, the parameters, the version).
'==============================================================================
Option Compare Database
Option Explicit

Private Const VIEW_PREFIX As String = "sv_"
Private Const SERVER_MARK As String = "SELECT * FROM [sv_"
Private Const ACCESS_SQL As String = "AccessSql"

'------------------------------------------------------------------------------
' Pure helpers (no database): mirrored in tools/gen_sql_views.py
'------------------------------------------------------------------------------
Public Function ServerQuerySql(ByVal QueryName As String, ByVal OrderBy As String) As String
    ' The saved query over its linked view.
    ServerQuerySql = SERVER_MARK & QueryName & "]" & OrderBy & ";"
End Function

Public Function ServerParamKind(ByVal ParamName As String) As String
    ' "D" (a date, QDate), "L" (a number, QLong), or "" when no saved query reads it.
    Dim items As Variant, i As Long
    items = Split(ServerParamList(), ",")
    For i = 0 To UBound(items)
        If StrComp(Split(items(i), ":")(0), ParamName, vbTextCompare) = 0 Then
            ServerParamKind = Split(items(i), ":")(1)
            Exit Function
        End If
    Next
End Function

Public Function QueryParamsSql(ByVal Names As Variant, ByVal Values As Variant) As String
    ' MERGE of the parameters into this process's row of dbo.QueryParams ("" when none of them is read by a
    ' saved query). A value Null / Empty = not set (the view gives 1899-12-30 or 0, like QDate / QLong).
    Dim i As Long, kind As String, v As String, sets As String, cols As String, vals As String
    For i = LBound(Names) To UBound(Names)
        kind = ServerParamKind(CStr(Names(i)))
        If Len(kind) > 0 Then
            If IsNull(Values(i)) Or IsEmpty(Values(i)) Then
                v = "NULL"
            ElseIf kind = "D" Then
                v = SqlLiteral(CDate(Values(i)), 8)
            Else
                v = SqlLiteral(CLng(Values(i)), 4)
            End If
            sets = sets & "[" & Names(i) & "] = " & v & ", "
            cols = cols & ", [" & Names(i) & "]"
            vals = vals & ", " & v
        End If
    Next
    If Len(sets) = 0 Then Exit Function
    QueryParamsSql = "MERGE [dbo].[QueryParams] AS t" & vbLf & _
        "USING (SELECT HOST_NAME() AS HostName, HOST_ID() AS HostProcess) AS s" & vbLf & _
        "ON t.HostName = s.HostName AND t.HostProcess = s.HostProcess" & vbLf & _
        "WHEN MATCHED THEN UPDATE SET " & sets & "UpdatedAt = GETDATE()" & vbLf & _
        "WHEN NOT MATCHED THEN INSERT (HostName, HostProcess" & cols & ") VALUES (s.HostName, s.HostProcess" & vals & ");" & vbLf & _
        "DELETE FROM [dbo].[QueryParams] WHERE UpdatedAt < DATEADD(day, -30, GETDATE());"
End Function

'------------------------------------------------------------------------------
' The parameters
'------------------------------------------------------------------------------
Public Sub SendQueryParams(ByVal Names As Variant, ByVal Values As Variant)
    ' Called by modQueryParams after it set TempVars: the server views read the same values.
    Dim sql As String, msg As String
    If Not IsLinkedToServer() Then Exit Sub
    sql = QueryParamsSql(Names, Values)
    If Len(sql) = 0 Then Exit Sub
    msg = ServerExec(LinkConnect(), sql)
    If Len(msg) > 0 Then ShowWarning " ⁄–— ≈—”«· › —… «· ﬁ—Ì— ≈·Ï «·Œ«œ„:" & vbCrLf & msg
End Sub

Public Sub SendAllQueryParams()
    ' Every parameter as it is now in TempVars (after the link to the server).
    Dim items As Variant, i As Long, n() As Variant, v() As Variant
    items = Split(ServerParamList(), ",")
    ReDim n(UBound(items))
    ReDim v(UBound(items))
    On Error Resume Next
    For i = 0 To UBound(items)
        n(i) = Split(items(i), ":")(0)
        v(i) = Null
        v(i) = TempVars(CStr(n(i))).Value
    Next
    On Error GoTo 0
    SendQueryParams n, v
End Sub

'------------------------------------------------------------------------------
' The saved queries
'------------------------------------------------------------------------------
Public Function ServerQueriesProblem(ByVal Connect As String) As String
    ' "" when the server has the views of this front-end, else what to do.
    Dim v As Variant
    v = ServerValue(Connect, "SELECT Version FROM [dbo].[QueriesVersion]", 30)
    If IsNull(v) Then
        ServerQueriesProblem = "ﬁ«⁄œ… «·Œ«œ„ ·Ì” ›ÌÂ« «” ⁄·«„«  «·»—‰«„Ã: ‘€¯· ´ ÃÂÌ“ «·ﬁ«⁄œ…ª √Ê·« („·› RetailStore_Queries.sql)."
    ElseIf CStr(v) <> SERVER_QUERIES_VERSION Then
        ServerQueriesProblem = "«” ⁄·«„«  ﬁ«⁄œ… «·Œ«œ„ „‰ ‰”Œ… √Œ—Ï „‰ «·»—‰«„Ã: ‘€¯· ´ ÃÂÌ“ «·ﬁ«⁄œ…ª »”ﬂ—»  Â–Â «·‰”Œ…."
    End If
End Function

Public Function ServerQueriesWarning() As String
    ' The login screen (after CheckServerAtStartup): a front-end on the server whose views are of another version
    ' (a rebuilt front-end, the script not run again) - its reports may fail until "Prepare the database".
    If Not IsLinkedToServer() Then Exit Function
    ServerQueriesWarning = ServerQueriesProblem(LinkConnect())
End Function

Private Function AccessSqlOf(ByVal qdf As DAO.QueryDef) As String
    On Error Resume Next
    AccessSqlOf = qdf.Properties(ACCESS_SQL).Value
End Function

Private Sub KeepAccessSql(ByVal qdf As DAO.QueryDef)
    ' The Access SQL of a query not on the server yet (a new front-end, or BuildQueries rebuilt it).
    Dim sql As String
    sql = qdf.SQL
    If Left$(sql, Len(SERVER_MARK)) = SERVER_MARK Then Exit Sub
    On Error Resume Next
    qdf.Properties(ACCESS_SQL).Value = sql
    If Err.Number <> 0 Then
        Err.Clear
        qdf.Properties.Append qdf.CreateProperty(ACCESS_SQL, dbMemo, sql)
    End If
End Sub

Private Function QueryDefOf(ByVal db As DAO.Database, ByVal QueryName As String) As DAO.QueryDef
    On Error Resume Next
    Set QueryDefOf = db.QueryDefs(QueryName)
End Function

Public Function UseServerQueries(ByVal Connect As String) As String
    ' Each saved query on its server view ("" = done). Called after the tables were linked to the server.
    Dim db As DAO.Database, items As Variant, i As Long, parts As Variant, qdf As DAO.QueryDef
    Dim schema As String, savePwd As Boolean, done As Long
    On Error GoTo Failed
    If UiEnglish() Then schema = "en" Else schema = "dbo"
    savePwd = InStr(1, Connect, "UID=", vbTextCompare) > 0
    Set db = CurrentDb
    items = Split(ServerQueryList(), ";")
    For i = 0 To UBound(items)
        parts = Split(items(i), "|")
        Set qdf = QueryDefOf(db, CStr(parts(0)))
        If Not qdf Is Nothing And parts(1) <> "~" Then          ' "~" = its ORDER BY stays in Access
            SysCmd acSysCmdSetStatus, Tr("«” ⁄·«„«  «·Œ«œ„: ") & (i + 1) & " / " & (UBound(items) + 1)
            KeepAccessSql qdf
            RelinkTable db, VIEW_PREFIX & parts(0), Connect, schema & "." & parts(0), savePwd
            qdf.SQL = ServerQuerySql(CStr(parts(0)), CStr(parts(1)))
            done = done + 1
        End If
    Next
    db.TableDefs.Refresh
    SysCmd acSysCmdClearStatus
    SendAllQueryParams
    LogAction "SQL_QUERIES", "SqlServer", schema, CStr(done)
    Exit Function
Failed:
    SysCmd acSysCmdClearStatus
    UseServerQueries = " ⁄–— —»ÿ «·«” ⁄·«„ " & parts(0) & " »«·Œ«œ„: " & Err.Description
End Function

Public Function UseAccessQueries() As String
    ' Every saved query back to its Access SQL, and the linked views removed ("" = done).
    Dim db As DAO.Database, qdf As DAO.QueryDef, sql As String, i As Long, current As String
    On Error GoTo Failed
    Set db = CurrentDb
    For Each qdf In db.QueryDefs
        sql = AccessSqlOf(qdf)
        If Len(sql) > 0 Then
            current = qdf.Name
            qdf.SQL = sql
            qdf.Properties.Delete ACCESS_SQL
        End If
    Next
    For i = db.TableDefs.Count - 1 To 0 Step -1
        If Left$(db.TableDefs(i).Name, Len(VIEW_PREFIX)) = VIEW_PREFIX And Len(db.TableDefs(i).Connect) > 0 Then
            db.TableDefs.Delete db.TableDefs(i).Name
        End If
    Next
    db.TableDefs.Refresh
    Exit Function
Failed:
    UseAccessQueries = " ⁄–— ≈—Ã«⁄ «·«” ⁄·«„ " & current & " ≈·Ï Access: " & Err.Description
End Function

Public Sub ServerQueriesAfterBuild()
    ' BuildQueries wrote the Access SQL again: a front-end on the server puts its queries back on the views.
    Dim msg As String
    If Not IsLinkedToServer() Then Exit Sub
    msg = UseServerQueries(LinkConnect())
    If Len(msg) > 0 Then ShowWarning msg
End Sub

Public Function ServerQueryCount() As Long
    ' The saved queries of this front-end that read a server view.
    Dim db As DAO.Database, qdf As DAO.QueryDef
    Set db = CurrentDb
    For Each qdf In db.QueryDefs
        If Left$(qdf.SQL, Len(SERVER_MARK)) = SERVER_MARK Then ServerQueryCount = ServerQueryCount + 1
    Next
End Function

'------------------------------------------------------------------------------
' In-Access test (Immediate window: TestServerQueries). Changes nothing.
'------------------------------------------------------------------------------
Private Sub CheckSq(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                    ByRef Report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "  OK   " & Title
    Else
        failed = failed + 1
        Debug.Print "  FAIL " & Title
        Report = Report & "- " & Title & vbCrLf
    End If
End Sub

Private Function Totals(ByVal Sql As String) As Variant
    ' Array(rows, total of the number columns) of a query, to compare Access and the server.
    Dim rs As DAO.Recordset, f As DAO.Field, n As Long, t As Double
    Set rs = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
    Do Until rs.EOF
        n = n + 1
        For Each f In rs.Fields
            Select Case f.Type
                Case dbByte, dbInteger, dbLong, dbCurrency, dbSingle, dbDouble, dbDecimal
                    t = t + Nz(f.Value, 0)
            End Select
        Next
        rs.MoveNext
    Loop
    rs.Close
    Totals = Array(n, t)
End Function

Public Function TestServerQueries() As Boolean
    Dim passed As Long, failed As Long, Report As String, s As String, items As Variant, i As Long
    Dim db As DAO.Database, qdf As DAO.QueryDef, onServer As Variant, inAccess As Variant
    Dim differ As String, savedStart As Variant, savedEnd As Variant
    Calendar = vbCalGreg
    Debug.Print "=== TestServerQueries  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckSq ServerQuerySql("TrialBalanceQuery", " ORDER BY [AccountCode]") = _
            "SELECT * FROM [sv_TrialBalanceQuery] ORDER BY [AccountCode];", "«·«” ⁄·«„ ⁄·Ï «·⁄—÷ «·„—»Êÿ", passed, failed, Report
    CheckSq ServerParamKind("PeriodStart") = "D" And ServerParamKind("AccountCode") = "L" And ServerParamKind("ReportCriteria") = "", _
            "√‰Ê«⁄ «·„⁄«„·« ", passed, failed, Report
    s = QueryParamsSql(Array("PeriodStart", "AccountCode", "ReportCriteria"), Array(DateSerial(2026, 3, 1), 1300, "x"))
    CheckSq InStr(s, "[PeriodStart] = '20260301 00:00:00', [AccountCode] = 1300, UpdatedAt = GETDATE()") > 0 And _
            InStr(s, "ReportCriteria") = 0 And InStr(s, "HOST_ID()") > 0, "≈—”«· «·„⁄«„·«  ≈·Ï «·Œ«œ„", passed, failed, Report
    CheckSq QueryParamsSql(Array("ReportCriteria"), Array("x")) = "", "„⁄«„· ·« Ìﬁ—ƒÂ «” ⁄·«„ ·« Ìı—”·", passed, failed, Report
    CheckSq InStr(QueryParamsSql(Array("CustomerID"), Array(Null)), "[CustomerID] = NULL") > 0, "„⁄«„· „Õ–Ê›", passed, failed, Report
    items = Split(ServerQueryList(), ";")
    CheckSq UBound(items) > 200 And Len(SERVER_QUERIES_VERSION) = 12, "ﬁ«∆„… «” ⁄·«„«  «·Œ«œ„", passed, failed, Report
    If IsLinkedToServer() Then
        CheckSq Len(ServerQueriesProblem(LinkConnect())) = 0, "«·Œ«œ„ ›ÌÂ «” ⁄·«„«  Â–Â «·‰”Œ…", passed, failed, Report
        CheckSq ServerQueryCount() > 200, "«·«” ⁄·«„«   ⁄„· ⁄·Ï «·Œ«œ„ (" & ServerQueryCount() & ")", passed, failed, Report
        savedStart = TempVars("PeriodStart")
        savedEnd = TempVars("PeriodEnd")
        SetPeriod DateSerial(Year(Date), 1, 1), Date
        Set db = CurrentDb
        For i = 0 To UBound(items)                      ' each query: the server view = its Access SQL
            Set qdf = QueryDefOf(db, CStr(Split(items(i), "|")(0)))
            If Not qdf Is Nothing Then
                If Len(AccessSqlOf(qdf)) > 0 Then
                    onServer = Array(-1, 0)
                    inAccess = Array(-2, 0)
                    On Error Resume Next
                    onServer = Totals("SELECT * FROM [" & qdf.Name & "]")
                    inAccess = Totals(AccessSqlOf(qdf))
                    Err.Clear
                    On Error GoTo 0
                    If onServer(0) <> inAccess(0) Or Abs(onServer(1) - inAccess(1)) > 0.01 Then
                        differ = differ & qdf.Name & " (" & onServer(0) & " <> " & inAccess(0) & ") "
                    End If
                End If
            End If
        Next
        CheckSq Len(differ) = 0, "ﬂ· «” ⁄·«„ ⁄·Ï «·Œ«œ„ = ‰›”Â ›Ì Access " & Left$(differ, 400), passed, failed, Report
        If IsNull(savedStart) Then
            RemoveQueryParams Array("PeriodStart", "PeriodEnd")
        Else
            SetQueryParams "PeriodStart", savedStart, "PeriodEnd", savedEnd
        End If
    End If
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  «” ⁄·«„«  «·Œ«œ„ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestServerQueries"
        TestServerQueries = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & Report, vbExclamation + MSG_RTL, "TestServerQueries"
    End If
End Function
