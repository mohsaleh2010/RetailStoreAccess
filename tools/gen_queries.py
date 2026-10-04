"""Generate modBuildQueries.bas and the queries reference doc."""

import re

from generate_common import vba_str
import queries as Q
from schema import table

_PLACEHOLDER = re.compile(r"\{(ref|day|dayh):([\w:]+)\}")


def vba_expr(text: str) -> str:
    """VBA string expression for an SQL text with placeholders:
    {ref:KEY} -> R("KEY"), {day:N} -> D(N, 0), {dayh:N:H} -> D(N, H)."""
    def sub(m):
        kind, arg = m.group(1), m.group(2)
        if kind == "ref":
            call = f'R("{arg}")'
        elif kind == "day":
            call = f"D({int(arg)}, 0)"
        else:
            n, h = arg.split(":")
            call = f"D({int(n)}, {int(h)})"
        return f'" & {call} & "'
    out = _PLACEHOLDER.sub(sub, vba_str(text))
    if out.endswith(' & ""'):
        out = out[:-5]
    if out.startswith('"" & '):
        out = out[5:]
    return out


def sql_literal(v) -> str:
    if v is None:
        return "Null"
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        return repr(v)
    if isinstance(v, Q.Ref):
        return "{ref:" + v.key + "}"
    if isinstance(v, Q.Day):
        return "{dayh:" + f"{v.days_ago}:{v.hour}" + "}"
    return "'" + str(v).replace("'", "''") + "'"


def fixture_line(row: Q.Row) -> str:
    cols = ", ".join(f"[{c}]" for c in row.values)
    vals = ", ".join(sql_literal(v) for v in row.values.values())
    sql = f"INSERT INTO [{row.table}] ({cols}) VALUES ({vals})"
    pk = table(row.table).pk[0]
    return (f"    Ins {vba_str(row.key)}, {vba_str(row.table)}, {vba_str(pk)}, _\n"
            f"        {vba_expr(sql)}")


def expected_expr(check: Q.Check) -> str:
    e = check.expected
    if isinstance(e, str) and e.startswith("ref:"):
        return f'CDbl(R("{e[4:]}"))'
    return repr(e)


def check_lines(check: Q.Check) -> str:
    out = []
    for name, key in (check.params or {}).items():
        out.append(f'    SetQueryParam {vba_str(name)}, CLng(R("{key}"))')
    out.append(f"    Chk {vba_str(check.label)}, _\n        {vba_expr(check.sql)}, {expected_expr(check)}")
    return "\n".join(out)


def query_sub(q: Q.Query) -> str:
    lines = [l.rstrip() for l in q.sql.strip().splitlines()]
    out = [f"Private Sub Q_{q.name}()", "    Dim s As String"]
    for i, l in enumerate(lines):
        prefix = "s = " if i == 0 else "s = s & "
        out.append(f"    {prefix}{vba_str(l)} & vbCrLf")
    out.append(f"    SaveQuery {vba_str(q.name)}, {vba_str(q.caption)}, s")
    out.append("End Sub")
    return "\n".join(out)


VBA_TEMPLATE = r'''Attribute VB_Name = "modBuildQueries"
'==============================================================================
' modBuildQueries  -  Retail Store Management System (Phase 4: Queries)
'
' GENERATED FILE - do not edit by hand.
' Source of truth: tools/queries.py  ->  python3 tools/generate.py
'
' Public procedures (run from the Immediate window, Ctrl+G):
'   BuildQueries     creates or updates every saved query in this front-end.
'   TestQueries      loads a small known business history inside a transaction,
'                    checks the numbers returned by the queries, then rolls back.
'                    Runs only on a database without transactions.
'   DropQueries      DEVELOPMENT ONLY: deletes the system queries.
'
' Requires: modQueryParams (QDate, QLong, SetPeriod), BuildSchema, BuildRelationships.
'==============================================================================
Option Compare Database
Option Explicit

Private Const MSG_RTL As Long = &H180000           ' vbMsgBoxRight + vbMsgBoxRtlReading
Private Const PERIOD_START_DAYS_AGO As Long = @@PERIOD@@
Private Const TEST_SLOW_MOVING_DAYS As Long = 90
Private Const QUERY_NAMES As String = "@@NAMES@@"

Private m_db As DAO.Database
Private m_created As Long
Private m_updated As Long
Private m_failed As Long
Private m_passed As Long
Private m_report As String
Private m_ids As Collection
Private m_keys As Collection

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildQueries() As Boolean
    m_created = 0: m_updated = 0: m_failed = 0: m_report = ""
    Debug.Print "=== BuildQueries  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Set m_db = CurrentDb

    CreateAllQueries

    m_db.QueryDefs.Refresh
    Application.RefreshDatabaseWindow
    Debug.Print "--- جديدة: " & m_created & " | محدّثة: " & m_updated & " | فشلت: " & m_failed

    If m_failed = 0 Then
        MsgBox "تم إنشاء الاستعلامات بنجاح." & vbCrLf & vbCrLf & _
               "استعلامات جديدة: " & m_created & vbCrLf & _
               "استعلامات محدّثة: " & m_updated & vbCrLf & vbCrLf & _
               "الخطوة التالية: شغّل TestQueries", vbInformation + MSG_RTL, "BuildQueries"
        BuildQueries = True
    Else
        MsgBox "فشل إنشاء " & m_failed & " استعلام:" & vbCrLf & vbCrLf & Left$(m_report, 900), _
               vbExclamation + MSG_RTL, "BuildQueries"
    End If
End Function

Public Function TestQueries() As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, blocker As String, slowDays As Long

    On Error GoTo EH
    Calendar = vbCalGreg
    m_passed = 0: m_failed = 0: m_report = ""
    Set m_db = CurrentDb

    blocker = ExistingDataTable()
    If Len(blocker) > 0 Then
        MsgBox "يعمل هذا الاختبار على قاعدة بدون حركات فقط، لأن نتائجه أرقام محددة مسبقًا." & _
               vbCrLf & "يوجد بيانات في الجدول: " & blocker, vbExclamation + MSG_RTL, "TestQueries"
        Exit Function
    End If

    Debug.Print "=== TestQueries  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    slowDays = ScalarLong("SELECT SlowMovingDays FROM Settings")
    Set m_ids = New Collection
    Set m_keys = New Collection
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    m_db.Execute "UPDATE [Settings] SET [SlowMovingDays] = " & TEST_SLOW_MOVING_DAYS, dbFailOnError
    LoadFixture
    SetPeriod DateAdd("d", -PERIOD_START_DAYS_AGO, Date), Date

    CheckQueriesOpen
    RunChecks
    RunCorruptionChecks

    ws.Rollback
    inTrans = False
    CleanUpAfterTest slowDays
    ClearQueryParams

    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed & " (تم التراجع عن بيانات الاختبار)"
    If m_failed = 0 Then
        MsgBox "جميع اختبارات الاستعلامات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "الأرقام مطابقة للحسابات اليدوية، ولم تُترك أي بيانات اختبار.", _
               vbInformation + MSG_RTL, "TestQueries"
        TestQueries = True
    Else
        MsgBox "نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestQueries"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    If inTrans Then ws.Rollback
    CleanUpAfterTest slowDays
    ClearQueryParams
    MsgBox errText & vbCrLf & "تم التراجع عن بيانات الاختبار.", vbCritical + MSG_RTL, "TestQueries"
End Function

Public Sub DropQueries()
    ' DEVELOPMENT ONLY - deletes the queries created by BuildQueries.
    Dim db As DAO.Database, qName As Variant, removed As Long
    If InputBox("سيتم حذف استعلامات النظام (البيانات لا تُحذف)." & vbCrLf & _
                "للتأكيد اكتب DELETE", "DropQueries") <> "DELETE" Then Exit Sub
    Set db = CurrentDb
    For Each qName In Split(QUERY_NAMES, ",")
        If QueryExists(db, CStr(qName)) Then
            db.QueryDefs.Delete CStr(qName)
            removed = removed + 1
        End If
    Next
    Application.RefreshDatabaseWindow
    MsgBox "تم حذف " & removed & " استعلام.", vbInformation + MSG_RTL
End Sub

'------------------------------------------------------------------------------
' Building
'------------------------------------------------------------------------------
Private Sub SaveQuery(ByVal QueryName As String, ByVal Description As String, ByVal Sql As String)
    Dim qdf As DAO.QueryDef
    On Error GoTo EH
    If QueryExists(m_db, QueryName) Then
        Set qdf = m_db.QueryDefs(QueryName)
        qdf.SQL = Sql
        m_updated = m_updated + 1
        Debug.Print "  ~ " & QueryName
    Else
        Set qdf = m_db.CreateQueryDef(QueryName, Sql)
        m_created = m_created + 1
        Debug.Print "  + " & QueryName
    End If
    SetProp qdf, "Description", dbText, Description
    Exit Sub
EH:
    Fail QueryName & ": خطأ " & Err.Number & " - " & Err.Description
End Sub

Private Sub SetProp(ByVal obj As Object, ByVal PropName As String, _
                    ByVal PropType As Integer, ByVal PropValue As Variant)
    Dim prp As Object
    On Error Resume Next
    Set prp = obj.Properties(PropName)
    On Error GoTo 0
    If prp Is Nothing Then
        obj.Properties.Append obj.CreateProperty(PropName, PropType, PropValue)
    Else
        prp.Value = PropValue
    End If
End Sub

'------------------------------------------------------------------------------
' Testing
'------------------------------------------------------------------------------
Private Function ExistingDataTable() As String
    Dim t As Variant
    For Each t In Array("SalesInvoices", "SalesReturns", "PurchaseInvoices", "PurchaseReturns", _
                        "CustomerPayments", "SupplierPayments", "Expenses", _
                        "InventoryTransactions", "StockCounts", "Products", "Suppliers")
        If ScalarLong("SELECT COUNT(*) FROM [" & t & "]") > 0 Then
            ExistingDataTable = CStr(t)
            Exit Function
        End If
    Next
    If ScalarLong("SELECT COUNT(*) FROM [Customers]") > 1 Then ExistingDataTable = "Customers"
End Function

Private Sub CheckQueriesOpen()
    Dim qName As Variant, rs As DAO.Recordset, opened As Long
    SetQueryParam "CustomerID", CLng(R("C2"))
    SetQueryParam "SupplierID", CLng(R("S1"))
    SetQueryParam "ProductID", CLng(R("P2"))
    For Each qName In Split(QUERY_NAMES, ",")
        On Error Resume Next
        Set rs = m_db.OpenRecordset("SELECT * FROM [" & qName & "]", dbOpenSnapshot)
        If Err.Number <> 0 Then
            Fail "الاستعلام " & qName & " لا يعمل: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            rs.Close
            opened = opened + 1
        End If
        On Error GoTo 0
    Next
    Call Record(opened = UBound(Split(QUERY_NAMES, ",")) + 1, _
                "كل الاستعلامات (" & opened & ") تعمل بدون أخطاء")
End Sub

Private Sub Chk(ByVal Label As String, ByVal Sql As String, ByVal Expected As Double)
    Dim rs As DAO.Recordset, actual As Variant
    On Error GoTo EH
    Set rs = m_db.OpenRecordset(Sql, dbOpenSnapshot)
    If rs.EOF Then
        actual = Null
    Else
        actual = rs(0).Value
    End If
    rs.Close
    If IsNull(actual) Then
        Fail Label & ": لا توجد نتيجة (المتوقع " & Expected & ")"
    ElseIf Abs(CDbl(actual) - Expected) < 0.005 Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label & " = " & actual
    Else
        Fail Label & ": النتيجة " & actual & " والمتوقع " & Expected
    End If
    Exit Sub
EH:
    Fail Label & ": خطأ " & Err.Number & " - " & Err.Description
End Sub

Private Sub Record(ByVal Passed As Boolean, ByVal Label As String)
    If Passed Then
        m_passed = m_passed + 1
        Debug.Print "[OK] " & Label
    Else
        Fail Label
    End If
End Sub

Private Sub Fail(ByVal Msg As String)
    m_failed = m_failed + 1
    m_report = m_report & "- " & Msg & vbCrLf
    Debug.Print "[X] " & Msg
End Sub

Private Sub Ins(ByVal Key As String, ByVal TableName As String, ByVal PkField As String, _
                ByVal Sql As String)
    m_db.Execute Sql, dbFailOnError
    m_ids.Add ScalarLong("SELECT @@IDENTITY"), Key
    m_keys.Add TableName & "|" & PkField & "|" & Key
End Sub

Private Function R(ByVal Key As String) As String
    R = CStr(m_ids(Key))
End Function

Private Function D(ByVal DaysAgo As Long, ByVal HourOfDay As Long) As String
    ' Access SQL date literal, always Gregorian (Saudi PCs may default to Hijri).
    D = "#" & Format$(DateAdd("d", -DaysAgo, Date) + TimeSerial(HourOfDay, 0, 0), _
                      "yyyy-mm-dd hh:nn:ss") & "#"
End Function

Private Sub CleanUpAfterTest(ByVal SlowMovingDays As Long)
    ' Safety net: normally the rollback already removed everything.
    Dim i As Long, parts() As String
    On Error Resume Next
    If m_keys Is Nothing Then Exit Sub
    If ScalarLong("SELECT COUNT(*) FROM [Products] WHERE Left([ProductCode], 5) = 'TEST-'") > 0 Then
        Debug.Print "تنبيه: لم يشمل التراجع كل الجداول، يتم حذف بيانات الاختبار يدويًا."
        For i = m_keys.Count To 1 Step -1
            parts = Split(m_keys(i), "|")
            m_db.Execute "DELETE FROM [" & parts(0) & "] WHERE [" & parts(1) & "] = " & _
                         m_ids(parts(2)), dbFailOnError
        Next
    End If
    If SlowMovingDays > 0 Then
        m_db.Execute "UPDATE [Settings] SET [SlowMovingDays] = " & SlowMovingDays, dbFailOnError
    End If
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function QueryExists(ByVal db As DAO.Database, ByVal QueryName As String) As Boolean
    Dim qdf As DAO.QueryDef
    For Each qdf In db.QueryDefs
        If StrComp(qdf.Name, QueryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next
End Function

Private Function ScalarLong(ByVal Sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = m_db.OpenRecordset(Sql, dbOpenSnapshot)
    If Not rs.EOF Then ScalarLong = Nz(rs(0), 0)
    rs.Close
End Function

'------------------------------------------------------------------------------
' Generated: test fixture, checks and corruption checks
'------------------------------------------------------------------------------
Private Sub LoadFixture()
@@FIXTURE@@
End Sub

Private Sub RunChecks()
@@CHECKS@@
End Sub

Private Sub RunCorruptionChecks()
@@CORRUPTIONS@@
End Sub

'------------------------------------------------------------------------------
' Generated: query definitions (in dependency order)
'------------------------------------------------------------------------------
Private Sub CreateAllQueries()
@@CREATE_ALL@@
End Sub

@@QUERY_SUBS@@
'''


def build_queries_vba() -> str:
    corruption_lines = []
    for update, check in Q.CORRUPTIONS:
        corruption_lines.append(f"    m_db.Execute {vba_expr(update)}, dbFailOnError")
        corruption_lines.append(check_lines(check))
    text = VBA_TEMPLATE
    for key, value in {
        "@@PERIOD@@": str(Q.PERIOD_START_DAYS_AGO),
        "@@NAMES@@": ",".join(q.name for q in Q.QUERIES),
        "@@FIXTURE@@": "\n".join(fixture_line(r) for r in Q.FIXTURE),
        "@@CHECKS@@": "\n".join(check_lines(c) for c in Q.CHECKS),
        "@@CORRUPTIONS@@": "\n".join(corruption_lines),
        "@@CREATE_ALL@@": "\n".join(f"    Q_{q.name}" for q in Q.QUERIES),
        "@@QUERY_SUBS@@": "\n\n".join(query_sub(q) for q in Q.QUERIES),
    }.items():
        text = text.replace(key, value)
    assert not re.search(r"@@[A-Z_]+@@", text), "unfilled template placeholder"
    assert "{ref:" not in text and "{day" not in text
    return text


def build_queries_md() -> str:
    spec = {"DailySalesQuery", "MonthlySalesQuery", "StockBalanceQuery", "LowStockQuery",
            "CustomerBalanceQuery", "SupplierBalanceQuery", "ProfitQuery",
            "BestSellingProductsQuery", "SlowMovingProductsQuery"}
    out = ["# مرجع الاستعلامات (Queries Reference)", "",
           "> ملف مُولَّد تلقائيًا من `tools/queries.py` – لا تعدّله يدويًا.", "",
           f"عدد الاستعلامات: **{len(Q.QUERIES)}**. الاستعلامات التي تبدأ بـ `qry` مساعدة "
           "تستخدمها الاستعلامات الأخرى؛ البقية تُستخدم مباشرة في التقارير والنماذج. ⭐ = مطلوب بالاسم في البرومبت.", "",
           "| # | الاستعلام | الوصف | المعاملات |", "|---|---|---|---|"]
    for i, q in enumerate(Q.QUERIES, 1):
        star = " ⭐" if q.name in spec else ""
        params = ", ".join(f"`{p}`" for p in q.params)
        out.append(f"| {i} | [`{q.name}`](#{q.name.lower()}){star} | {q.caption} | {params} |")
    out += ["", "## بيانات الاختبار والنتائج المتوقعة", "",
            f"الفترة: آخر {Q.PERIOD_START_DAYS_AGO} يومًا حتى اليوم. التواريخ نسبية إلى يوم التشغيل.", "",
            "| # | الاختبار | الاستعلام | المتوقع |", "|---|---|---|---|"]
    for i, c in enumerate(Q.CHECKS, 1):
        exp = c.expected if not isinstance(c.expected, str) else c.expected.replace("ref:", "رقم ")
        out.append(f"| {i} | {c.label} | `{c.sql.replace('|', chr(92) + '|')}` | {exp} |")
    for q in Q.QUERIES:
        out += ["", f"## {q.name}", "", q.caption, ""]
        if q.params:
            out += ["المعاملات: " + ", ".join(f"`{p}`" for p in q.params), ""]
        out += ["```sql", q.sql.strip(), "```"]
    return "\n".join(out)
