"""Generate the Access build module and reference docs from tools/schema.py.

Usage:  python3 tools/generate.py
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from schema import TABLES, Table, Field, table  # noqa: E402
from generate_common import vba_str  # noqa: E402
import gen_relations  # noqa: E402
import gen_queries  # noqa: E402
import gen_forms  # noqa: E402
import gen_qr  # noqa: E402
import gen_lang  # noqa: E402
import gen_zatca  # noqa: E402
import gen_reports  # noqa: E402
import gen_test_sales  # noqa: E402
import gen_test_purchases  # noqa: E402
import gen_test_security  # noqa: E402
import gen_demo  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC_MODULES = ["modQueryParams", "modCommon", "modStartup", "modForms", "modScreens",
                  "modZatca", "modSales", "modPOS", "modPurchases",
                  "modPurchaseScreens", "modReports", "modDashboard",
                  "modSecurity", "modSecurityScreens", "modBackup", "modLabels", "modCharts", "modTouchPOS", "modCash", "modJournal", "modAccounts", "modManualEntry", "modLedger", "modFinancials", "modClosing", "modVat", "modAging", "modBank", "modCheque", "modAssets", "modPayroll", "modCostCenters", "modBudget", "modRecurring", "modAudit", "modIndicators", "modCurrency", "modSalesReps", "modEnglishNames", "modCountry", "modHttp", "modEInvoice", "modZatcaXml", "modZatcaApi", "modActivation", "modTestAll"]   # hand-written (not generated) VBA modules

KIND_LABEL = {
    "AUTO": "AutoNumber", "LONG": "Number (Long)", "INT": "Number (Integer)",
    "BYTE": "Number (Byte)", "MONEY": "Currency", "QTY": "Currency (كمية)",
    "RATE": "Currency (نسبة)", "DATE": "Date/Time (تاريخ)", "DATETIME": "Date/Time",
    "BOOL": "Yes/No", "TEXT": "Short Text", "MEMO": "Long Text",
}


# --------------------------------------------------------------------------
# VBA helpers
# --------------------------------------------------------------------------
def sql_value(v) -> str:
    if v is None:
        return "Null"
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        return repr(v)
    return "'" + str(v).replace("'", "''") + "'"


def vba_table_sub(t: Table) -> str:
    out = [f"Private Sub CreateTable_{t.name}()",
           "    Dim tdf As DAO.TableDef",
           f"    If Not BeginTable(tdf, {vba_str(t.name)}) Then Exit Sub"]
    for f in t.fields:
        out.append(
            "    AddField tdf, {name}, {kind}, {size}, {req}, {default}, _\n"
            "             {rule}, {rule_text}, {caption}, {note}".format(
                name=vba_str(f.name), kind=vba_str(f.kind), size=f.size,
                req="True" if f.required else "False",
                default=vba_str(f.default), rule=vba_str(f.rule),
                rule_text=vba_str(f.rule_text), caption=vba_str(f.caption),
                note=vba_str(f.note)))
    out.append(f"    AddIndex tdf, \"PrimaryKey\", {vba_str(','.join(t.pk))}, True, True, False")
    for ix in t.indexes:
        out.append("    AddIndex tdf, {n}, {f}, False, {u}, {ign}".format(
            n=vba_str(ix.name), f=vba_str(",".join(ix.fields)),
            u="True" if ix.unique else "False", ign="True" if ix.ignore_nulls else "False"))
    out.append("    EndTable tdf, {d}, {r}, {rt}".format(
        d=vba_str(f"{t.caption}: {t.purpose}"), r=vba_str(t.rule), rt=vba_str(t.rule_text)))
    out.append("End Sub")
    return "\n".join(out)


def vba_seed_sub(t: Table) -> str:
    cols = ", ".join(f"[{c}]" for c in t.seed_columns)
    upgrade = "True" if t.seed_missing else "False"
    out = [f"Private Sub Seed_{t.name}()",
           f"    If Not BeginSeed({vba_str(t.name)}, {upgrade}) Then Exit Sub"]
    grants = {}
    if t.name == "Permissions":
        for role_id, key in table("RolePermissions").seed_rows:
            grants.setdefault(key, []).append(str(role_id))
    for row in t.seed_rows:
        vals = ", ".join(sql_value(v) for v in row)
        sql = f"INSERT INTO [{t.name}] ({cols}) VALUES ({vals})"
        if not t.seed_missing:
            out.append(f"    ExecSeed {vba_str(sql)}")
            continue
        where = f"[{t.seed_columns[0]}] = {sql_value(row[0])}"
        if t.name == "Permissions":
            # a permission added by an upgrade is granted to its default roles
            out.append(f"    If SeedRow({vba_str(where)}, {vba_str(sql)}) Then "
                       f"GrantNewPermission {vba_str(row[0])}, {vba_str(','.join(grants.get(row[0], [])))}")
        else:
            out.append(f"    SeedRow {vba_str(where)}, {vba_str(sql)}")
    out.append(f"    EndSeed {vba_str(t.name)}, {len(t.seed_rows)}")
    out.append("End Sub")
    return "\n".join(out)


VBA_HEADER = r'''Attribute VB_Name = "modBuildSchema"
'==============================================================================
' modBuildSchema  -  Retail Store Management System (Phase 2: Tables)
'
' GENERATED FILE - do not edit by hand.
' Source of truth: tools/schema.py  ->  python3 tools/generate.py
'
' Public procedures (run from the Immediate window, Ctrl+G):
'   BuildSchema            creates RetailStore_BE.accdb next to this file,
'                          creates every missing table, seeds lookup data,
'                          then links the tables into this front-end.
'   BuildSchema "D:\Shop\RetailStore_BE.accdb"   same, custom back-end path.
'   VerifySchema           checks tables, field counts, seed data, Arabic text.
'   LinkBackEnd            (re)links all back-end tables into this front-end.
'   DropSchema             DEVELOPMENT ONLY: deletes all system tables.
'
' Requires: Access 2010 or later (DAO 12+ is referenced by default).
' Arabic text: Windows "Language for non-Unicode programs" must be Arabic.
'==============================================================================
Option Compare Database
Option Explicit

Private Const SCHEMA_VERSION As String = "2.0"
Private Const BE_FILE_NAME As String = "RetailStore_BE.accdb"
Private Const DB_VERSION_120 As Long = 128      ' dbVersion120 (.accdb format)
Private Const DISPLAY_CHECKBOX As Integer = 106 ' acCheckBox
Private Const MSG_RTL As Long = &H180000        ' vbMsgBoxRight + vbMsgBoxRtlReading

'@@SCHEMA_CONSTANTS@@

Private m_db As DAO.Database
Private m_pending As Collection
Private m_log As String
Private m_created As Long
Private m_skipped As Long
Private m_upgrade As Boolean        ' the table exists: only its missing fields are added
Private m_addedFields As Long
Private m_seeded As Long
Private m_seedTable As String
Private m_seedAdded As Long
Private m_seedOnlyMissing As Boolean
Private m_currentStep As String
Private m_inTrans As Boolean

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildSchema(Optional ByVal BackEndPath As String = "") As Boolean
    On Error GoTo EH
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    If StrComp(BackEndPath, CurrentProject.FullName, vbTextCompare) = 0 Then
        MsgBox "مسار ملف البيانات يجب أن يختلف عن ملف الواجهة الحالي.", vbExclamation + MSG_RTL
        Exit Function
    End If

    m_log = "": m_created = 0: m_skipped = 0: m_seeded = 0: m_addedFields = 0
    LogLine "=== BuildSchema " & SCHEMA_VERSION & "  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="

    m_currentStep = "open back-end"
    If Len(Dir$(BackEndPath)) = 0 Then
        Set m_db = DBEngine.CreateDatabase(BackEndPath, dbLangArabic, DB_VERSION_120)
        LogLine "تم إنشاء ملف البيانات: " & BackEndPath
    Else
        Set m_db = DBEngine.OpenDatabase(BackEndPath)
        LogLine "ملف البيانات موجود: " & BackEndPath
    End If

    CreateAllTables
    SeedAll
    m_currentStep = "developer user"
    EnsureDeveloperUser
    m_currentStep = "account tree"
    UpgradeAccountTree
    m_currentStep = "english names"
    SeedEnglishNames
    m_currentStep = "field rules"
    UpgradeFieldRules

    m_db.Close
    Set m_db = Nothing

    m_currentStep = "link tables"
    LinkBackEnd BackEndPath
    On Error Resume Next                     ' modAccounts may not be imported yet (manual installs)
    Application.Run "RebuildAccountTree"     ' levels and order of the account tree
    Err.Clear
    On Error GoTo EH

    LogLine "--- جداول جديدة: " & m_created & " | موجودة مسبقًا: " & m_skipped & _
            " | جداول تمت تعبئتها: " & m_seeded
    MsgBox "تم بناء الجداول بنجاح." & vbCrLf & vbCrLf & _
           "جداول جديدة: " & m_created & vbCrLf & _
           "جداول موجودة مسبقًا: " & m_skipped & vbCrLf & _
           "حقول جديدة أُضيفت لجداول موجودة: " & m_addedFields & vbCrLf & _
           "جداول تمت تعبئة بياناتها الأساسية: " & m_seeded & vbCrLf & vbCrLf & _
           "التفاصيل في نافذة Immediate (Ctrl+G)." & vbCrLf & _
           "الخطوة التالية: شغّل VerifySchema", vbInformation + MSG_RTL, "BuildSchema"
    BuildSchema = True
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ " & Err.Number & " أثناء [" & m_currentStep & "]: " & Err.Description
    If m_inTrans Then
        DBEngine.Workspaces(0).Rollback
        m_inTrans = False
    End If
    LogLine errText
    On Error Resume Next
    If Not m_db Is Nothing Then m_db.Close
    Set m_db = Nothing
    MsgBox errText & vbCrLf & vbCrLf & "يمكن إعادة تشغيل BuildSchema بعد الإصلاح؛ " & _
           "الجداول التي أُنشئت لن تتكرر.", vbCritical + MSG_RTL, "BuildSchema"
End Function

Public Sub LinkBackEnd(Optional ByVal BackEndPath As String = "")
    Dim dbFE As DAO.Database, dbBE As DAO.Database
    Dim tdfBE As DAO.TableDef, tdfFE As DAO.TableDef
    Dim linked As Long, refreshed As Long

    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    Set dbFE = CurrentDb
    Set dbBE = DBEngine.OpenDatabase(BackEndPath, False, True)

    For Each tdfBE In dbBE.TableDefs
        If IsUserTable(tdfBE) Then
            If TableExistsIn(dbFE, tdfBE.Name) Then
                Set tdfFE = dbFE.TableDefs(tdfBE.Name)
                If Len(tdfFE.Connect) > 0 Then
                    tdfFE.Connect = ";DATABASE=" & BackEndPath
                    tdfFE.RefreshLink
                    refreshed = refreshed + 1
                Else
                    LogLine "تنبيه: يوجد جدول محلي بنفس الاسم في الواجهة ولم يتم ربطه: " & tdfBE.Name
                End If
            Else
                Set tdfFE = dbFE.CreateTableDef(tdfBE.Name)
                tdfFE.Connect = ";DATABASE=" & BackEndPath
                tdfFE.SourceTableName = tdfBE.Name
                dbFE.TableDefs.Append tdfFE
                linked = linked + 1
            End If
        End If
    Next

    dbBE.Close
    dbFE.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    LogLine "الربط: " & linked & " جدول جديد، " & refreshed & " جدول تم تحديث ربطه."
End Sub

Public Function VerifySchema(Optional ByVal BackEndPath As String = "") As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim items() As String, parts() As String, i As Long
    Dim problems As Long, report As String, s As String

    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    If Len(Dir$(BackEndPath)) = 0 Then
        Call ResultBox("ملف البيانات غير موجود: " & BackEndPath, vbCritical + MSG_RTL)
        Exit Function
    End If
    Set db = DBEngine.OpenDatabase(BackEndPath, False, True)

    ' 1) every table exists with the expected number of fields
    items = Split(EXPECTED_FIELD_COUNTS, ";")
    For i = 0 To UBound(items)
        parts = Split(items(i), "=")
        If Not TableExistsIn(db, parts(0)) Then
            s = "[X] الجدول غير موجود: " & parts(0)
            problems = problems + 1
        ElseIf db.TableDefs(parts(0)).Fields.Count <> CLng(parts(1)) Then
            s = "[X] " & parts(0) & ": عدد الحقول " & db.TableDefs(parts(0)).Fields.Count & _
                " والمتوقع " & parts(1)
            problems = problems + 1
        Else
            s = "[OK] " & parts(0) & " (" & parts(1) & " حقل)"
        End If
        Debug.Print s
        If Left$(s, 3) = "[X]" Then report = report & s & vbCrLf
    Next

    ' 2) lookup data exists
    items = Split(EXPECTED_SEED_COUNTS, ";")
    For i = 0 To UBound(items)
        parts = Split(items(i), "=")
        If TableExistsIn(db, parts(0)) Then
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM [" & parts(0) & "]", dbOpenSnapshot)
            If rs(0) < CLng(parts(1)) Then
                s = "[X] " & parts(0) & ": " & rs(0) & " سجل والمتوقع " & parts(1) & " على الأقل"
                problems = problems + 1
                report = report & s & vbCrLf
            Else
                s = "[OK] بيانات " & parts(0) & ": " & rs(0) & " سجل"
            End If
            rs.Close
            Debug.Print s
        End If
    Next

    ' 3) Arabic text survived the VBA code page
    If TableExistsIn(db, "Roles") Then
        Set rs = db.OpenRecordset("SELECT RoleName FROM Roles WHERE RoleID=1", dbOpenSnapshot)
        If Not rs.EOF Then
            If AscW(Left$(rs!RoleName & " ", 1)) < &H600 Then
                s = "[X] النص العربي محفوظ بشكل خاطئ (" & rs!RoleName & ")." & vbCrLf & _
                    "    اضبط Windows > Region > Administrative > Language for non-Unicode programs = Arabic" & _
                    " ثم أعد الاستيراد والبناء."
                problems = problems + 1
                report = report & s & vbCrLf
            Else
                s = "[OK] النص العربي سليم: " & rs!RoleName
            End If
            Debug.Print s
        End If
        rs.Close
    End If

    db.Close
    If problems = 0 Then
        Call ResultBox("الفحص ناجح: جميع الجداول (" & (UBound(Split(SCHEMA_TABLES, ",")) + 1) & _
               ") والبيانات الأساسية سليمة.", vbInformation + MSG_RTL, "VerifySchema")
        VerifySchema = True
    Else
        Call ResultBox("عدد المشكلات: " & problems & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
               "VerifySchema")
    End If
End Function

Private Sub ResultBox(ByVal Text As String, ByVal Style As Long, Optional ByVal Title As String = "")
    ' Through modCommon.TestMsg when it is installed (RunAllTests collects the results),
    ' otherwise a plain message box: this module is installed before modCommon.
    On Error GoTo Plain
    Application.Run "TestMsg", Text, Style, Title
    Exit Sub
Plain:
    MsgBox Text, Style, Title
End Sub

Public Sub DropSchema(Optional ByVal BackEndPath As String = "")
    ' DEVELOPMENT ONLY - deletes every table of this system and all of its data.
    Dim db As DAO.Database, dbFE As DAO.Database, i As Long, names() As String

    If InputBox("سيتم حذف جميع جداول النظام وبياناتها نهائيًا." & vbCrLf & _
                "للتأكيد اكتب DELETE", "DropSchema") <> "DELETE" Then Exit Sub
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()

    names = Split(SCHEMA_TABLES, ",")
    Set db = DBEngine.OpenDatabase(BackEndPath)
    For i = db.Relations.Count - 1 To 0 Step -1
        If InSchema(db.Relations(i).Table) Or InSchema(db.Relations(i).ForeignTable) Then
            db.Relations.Delete db.Relations(i).Name
        End If
    Next
    For i = UBound(names) To 0 Step -1
        If TableExistsIn(db, names(i)) Then db.TableDefs.Delete names(i)
    Next
    db.Close

    Set dbFE = CurrentDb
    For i = UBound(names) To 0 Step -1
        If TableExistsIn(dbFE, names(i)) Then
            If Len(dbFE.TableDefs(names(i)).Connect) > 0 Then dbFE.TableDefs.Delete names(i)
        End If
    Next
    Application.RefreshDatabaseWindow
    MsgBox "تم حذف جداول النظام.", vbInformation + MSG_RTL
End Sub

'------------------------------------------------------------------------------
' Table building helpers
'------------------------------------------------------------------------------
Private Function BeginTable(ByRef tdf As DAO.TableDef, ByVal TableName As String) As Boolean
    ' A table that exists already is upgraded: its missing fields are added, nothing is removed.
    m_currentStep = "create table " & TableName
    Set m_pending = New Collection
    BeginTable = True
    m_upgrade = TableExistsIn(m_db, TableName)
    If m_upgrade Then
        LogLine "  = موجود مسبقًا: " & TableName
        m_skipped = m_skipped + 1
        Set tdf = m_db.TableDefs(TableName)
    Else
        Set tdf = m_db.CreateTableDef(TableName)
    End If
End Function

Private Function FieldExistsIn(ByVal tdf As DAO.TableDef, ByVal FieldName As String) As Boolean
    Dim fld As DAO.Field
    For Each fld In tdf.Fields
        If StrComp(fld.Name, FieldName, vbTextCompare) = 0 Then
            FieldExistsIn = True
            Exit Function
        End If
    Next
End Function

Private Sub AddField(ByVal tdf As DAO.TableDef, ByVal FieldName As String, ByVal Kind As String, _
                     ByVal Size As Long, ByVal IsRequired As Boolean, ByVal DefaultValue As String, _
                     ByVal ValidationRule As String, ByVal ValidationText As String, _
                     ByVal Caption As String, ByVal Description As String)
    Dim fld As DAO.Field, requiredLater As Boolean
    m_currentStep = "field " & tdf.Name & "." & FieldName
    If m_upgrade Then
        If FieldExistsIn(tdf, FieldName) Then
            ' a rule the schema dropped (PayrollLines.NetPay may be negative in a draft): removing
            ' a rule never fails on the existing rows
            If Len(ValidationRule) = 0 And Len(tdf.Fields(FieldName).ValidationRule) > 0 Then
                tdf.Fields(FieldName).ValidationRule = ""
                tdf.Fields(FieldName).ValidationText = ""
                LogLine "  ~ أُزيل شرط الحقل: " & tdf.Name & "." & FieldName
            End If
            Exit Sub
        End If
        requiredLater = IsRequired                 ' existing rows get the default value first
        IsRequired = False
    End If

    Select Case Kind
        Case "AUTO"
            Set fld = tdf.CreateField(FieldName, dbLong)
            fld.Attributes = fld.Attributes Or dbAutoIncrField
        Case "LONG":                     Set fld = tdf.CreateField(FieldName, dbLong)
        Case "INT":                      Set fld = tdf.CreateField(FieldName, dbInteger)
        Case "BYTE":                     Set fld = tdf.CreateField(FieldName, dbByte)
        Case "MONEY", "QTY", "RATE":     Set fld = tdf.CreateField(FieldName, dbCurrency)
        Case "DATE", "DATETIME":         Set fld = tdf.CreateField(FieldName, dbDate)
        Case "BOOL":                     Set fld = tdf.CreateField(FieldName, dbBoolean)
        Case "MEMO":                     Set fld = tdf.CreateField(FieldName, dbMemo)
        Case "TEXT"
            Set fld = tdf.CreateField(FieldName, dbText, Size)
            fld.AllowZeroLength = False
        Case Else
            Err.Raise vbObjectError + 513, "AddField", "Unknown field kind: " & Kind
    End Select

    If Kind <> "AUTO" And Kind <> "BOOL" Then fld.Required = IsRequired
    If Len(DefaultValue) > 0 Then fld.DefaultValue = DefaultValue
    If Len(ValidationRule) > 0 Then
        fld.ValidationRule = ValidationRule
        fld.ValidationText = ValidationText
    End If
    tdf.Fields.Append fld
    If m_upgrade Then
        If Len(DefaultValue) > 0 Then
            m_db.Execute "UPDATE [" & tdf.Name & "] SET [" & FieldName & "] = " & DefaultValue, dbFailOnError
        End If
        If requiredLater Then tdf.Fields(FieldName).Required = True
        m_addedFields = m_addedFields + 1
        LogLine "  + حقل جديد: " & tdf.Name & "." & FieldName
    End If

    ' Properties that can only be set after the table is saved
    If Len(Caption) > 0 Then AddPending FieldName, "Caption", dbText, Caption
    If Len(Description) > 0 Then AddPending FieldName, "Description", dbText, Description
    Select Case Kind
        Case "MONEY":    AddPending FieldName, "Format", dbText, "#,##0.00"
        Case "QTY":      AddPending FieldName, "Format", dbText, "#,##0.###"
        Case "RATE":     AddPending FieldName, "Format", dbText, "0.00%"
        Case "DATE":     AddPending FieldName, "Format", dbText, "yyyy/mm/dd"
        Case "DATETIME": AddPending FieldName, "Format", dbText, "yyyy/mm/dd hh:nn"
        Case "BOOL":     AddPending FieldName, "DisplayControl", dbInteger, DISPLAY_CHECKBOX
    End Select
End Sub

Private Sub AddIndex(ByVal tdf As DAO.TableDef, ByVal IndexName As String, ByVal FieldList As String, _
                     ByVal IsPrimary As Boolean, ByVal IsUnique As Boolean, ByVal IgnoreNulls As Boolean)
    Dim idx As DAO.Index, fieldName As Variant
    m_currentStep = "index " & tdf.Name & "." & IndexName
    If m_upgrade Then
        For Each idx In tdf.Indexes
            If StrComp(idx.Name, IndexName, vbTextCompare) = 0 Then Exit Sub
        Next
    End If
    Set idx = tdf.CreateIndex(IndexName)
    For Each fieldName In Split(FieldList, ",")
        idx.Fields.Append idx.CreateField(CStr(fieldName))
    Next
    idx.Primary = IsPrimary
    idx.Unique = IsUnique Or IsPrimary
    idx.IgnoreNulls = IgnoreNulls
    tdf.Indexes.Append idx
End Sub

Private Sub EndTable(ByVal tdf As DAO.TableDef, ByVal Description As String, _
                     ByVal TableRule As String, ByVal TableRuleText As String)
    Dim item As Variant, saved As DAO.TableDef
    m_currentStep = "save table " & tdf.Name
    If m_upgrade Then                              ' existing table: properties of the new fields only
        For Each item In m_pending
            m_currentStep = "property " & tdf.Name & "." & item(0) & "." & item(1)
            SetProp tdf.Fields(item(0)), item(1), item(2), item(3)
        Next
        Set m_pending = Nothing
        m_upgrade = False
        Exit Sub
    End If
    If Len(TableRule) > 0 Then
        tdf.ValidationRule = TableRule
        tdf.ValidationText = TableRuleText
    End If
    m_db.TableDefs.Append tdf
    m_db.TableDefs.Refresh

    Set saved = m_db.TableDefs(tdf.Name)
    SetProp saved, "Description", dbText, Description
    For Each item In m_pending
        m_currentStep = "property " & tdf.Name & "." & item(0) & "." & item(1)
        SetProp saved.Fields(item(0)), item(1), item(2), item(3)
    Next
    Set m_pending = Nothing
    m_created = m_created + 1
    LogLine "  + تم إنشاء الجدول: " & tdf.Name
End Sub

Private Sub AddPending(ByVal FieldName As String, ByVal PropName As String, _
                       ByVal PropType As Integer, ByVal PropValue As Variant)
    m_pending.Add Array(FieldName, PropName, PropType, PropValue)
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
' Seed helpers (lookup data is inserted only into empty tables)
'------------------------------------------------------------------------------
Private Function BeginSeed(ByVal TableName As String, Optional ByVal AddMissing As Boolean = False) As Boolean
    ' Empty table: all seed rows. Table with data: nothing, or (AddMissing) only the rows
    ' it does not have yet - new sequences / permissions of a later version.
    Dim rs As DAO.Recordset
    m_currentStep = "seed " & TableName
    m_seedTable = TableName
    m_seedAdded = 0
    Set rs = m_db.OpenRecordset("SELECT COUNT(*) FROM [" & TableName & "]", dbOpenSnapshot)
    m_seedOnlyMissing = (rs(0) > 0)
    rs.Close
    If m_seedOnlyMissing And Not AddMissing Then Exit Function
    DBEngine.Workspaces(0).BeginTrans
    m_inTrans = True
    BeginSeed = True
End Function

Private Sub ExecSeed(ByVal Sql As String)
    m_db.Execute Sql, dbFailOnError
    m_seedAdded = m_seedAdded + 1
End Sub

Private Function SeedRow(ByVal Where As String, ByVal Sql As String) As Boolean
    ' Inserts the row unless the table already has it. True = inserted.
    Dim rs As DAO.Recordset
    If m_seedOnlyMissing Then
        Set rs = m_db.OpenRecordset("SELECT COUNT(*) FROM [" & m_seedTable & "] WHERE " & Where, dbOpenSnapshot)
        If rs(0) > 0 Then
            rs.Close
            Exit Function
        End If
        rs.Close
    End If
    ExecSeed Sql
    SeedRow = True
End Function

Private Sub GrantNewPermission(ByVal PermissionKey As String, ByVal RoleIDs As String)
    ' Only for a permission added to an existing back-end (a new one is granted by Seed_RolePermissions).
    Dim ids() As String, i As Long
    If Not m_seedOnlyMissing Or Len(RoleIDs) = 0 Then Exit Sub
    ids = Split(RoleIDs, ",")
    For i = 0 To UBound(ids)
        If DCountIn("RolePermissions", "[RoleID] = " & ids(i) & " AND [PermissionKey] = '" & PermissionKey & "'") = 0 And _
           DCountIn("Roles", "[RoleID] = " & ids(i)) > 0 Then
            m_db.Execute "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (" & ids(i) & _
                         ", '" & PermissionKey & "')", dbFailOnError
        End If
    Next
End Sub

Private Sub EnsureDeveloperUser()
    ' The programmer: above the administrator, every permission, hidden from the users screen.
    ' Created only with the password typed now: there is never a programmer account without one.
    ' (PasswordHash / NewSalt of modSecurity through Application.Run: no compile-time dependency.)
    Dim pwd As String, salt As String
    If DCountIn("Employees", "[IsDeveloper] = True") > 0 Then Exit Sub
    If DCountIn("Employees", "[Username] = 'developer'") > 0 Then
        LogLine "تنبيه: يوجد مستخدم باسم developer وليس هو المبرمج؛ لم يُنشأ حساب المبرمج."
        Exit Sub
    End If
    Do
        pwd = InputBox("إنشاء حساب المبرمج (اسم المستخدم: developer)." & vbCrLf & vbCrLf & _
                       "اكتب كلمة مرور له (6 أحرف على الأقل) واحتفظ بها. لا تعطها لأحد." & vbCrLf & vbCrLf & _
                       "إلغاء = لا يُنشأ الآن، ويُطلب في التشغيل التالي لـ BuildSchema.", "حساب المبرمج")
        If Len(pwd) = 0 Then
            LogLine "لم يُنشأ حساب المبرمج: لم تُكتب كلمة مرور."
            Exit Sub
        End If
        If Len(pwd) >= 6 And pwd = Trim$(pwd) And pwd <> "developer" Then Exit Do
        MsgBox "كلمة المرور 6 أحرف على الأقل، بدون مسافة في أولها أو آخرها، ولا تساوي اسم المستخدم.", _
               vbExclamation + MSG_RTL, "حساب المبرمج"
    Loop
    salt = Application.Run("NewSalt")
    m_db.Execute "INSERT INTO [Employees] ([EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], " & _
                 "[MustChangePassword], [IsActive], [IsDeveloper], [PasswordSalt], [PasswordHash]) VALUES " & _
                 "('المبرمج', 'المبرمج', 'developer', 1, 1, False, True, True, '" & salt & "', '" & _
                 Application.Run("PasswordHash", pwd, salt) & "')", dbFailOnError
    LogLine "تم إنشاء حساب المبرمج: developer"
End Sub

Private Function DCountIn(ByVal TableName As String, ByVal Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = m_db.OpenRecordset("SELECT COUNT(*) FROM [" & TableName & "] WHERE " & Where, dbOpenSnapshot)
    DCountIn = rs(0)
    rs.Close
End Function

Private Sub EndSeed(ByVal TableName As String, ByVal RowCount As Long)
    DBEngine.Workspaces(0).CommitTrans
    m_inTrans = False
    If m_seedOnlyMissing Then
        If m_seedAdded > 0 Then LogLine "  * سجلات جديدة: " & TableName & " (" & m_seedAdded & " سجل)"
        Exit Sub
    End If
    m_seeded = m_seeded + 1
    LogLine "  * بيانات أساسية: " & TableName & " (" & RowCount & " سجل)"
End Sub

'------------------------------------------------------------------------------
' General helpers
'------------------------------------------------------------------------------
Private Function DefaultBackEndPath() As String
    DefaultBackEndPath = CurrentProject.Path & "\" & BE_FILE_NAME
End Function

Private Function TableExistsIn(ByVal db As DAO.Database, ByVal TableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, TableName, vbTextCompare) = 0 Then
            TableExistsIn = True
            Exit Function
        End If
    Next
End Function

Private Function IsUserTable(ByVal tdf As DAO.TableDef) As Boolean
    If (tdf.Attributes And dbSystemObject) <> 0 Then Exit Function
    If Left$(tdf.Name, 4) = "MSys" Or Left$(tdf.Name, 1) = "~" Then Exit Function
    IsUserTable = True
End Function

Private Function InSchema(ByVal TableName As String) As Boolean
    InSchema = InStr(1, "," & SCHEMA_TABLES & ",", "," & TableName & ",", vbTextCompare) > 0
End Function

Private Sub SetFieldRule(ByVal TableName As String, ByVal FieldName As String, ByVal Rule As String, _
                         ByVal RuleText As String)
    ' A rule the schema changed on a field that already exists (schema.RULE_UPGRADES, looser rules only).
    Dim fld As DAO.Field
    On Error GoTo Failed
    Set fld = m_db.TableDefs(TableName).Fields(FieldName)
    If fld.ValidationRule <> Rule Then
        fld.ValidationRule = Rule
        fld.ValidationText = RuleText
        LogLine "  ~ تغيّر شرط الحقل: " & TableName & "." & FieldName
    End If
    Exit Sub
Failed:
    LogLine "  ! تعذّر تغيير شرط الحقل " & TableName & "." & FieldName & ": " & Err.Description
End Sub

Private Sub LogLine(ByVal Msg As String)
    m_log = m_log & Msg & vbCrLf
    Debug.Print Msg
End Sub

'------------------------------------------------------------------------------
' Generated: table definitions
'------------------------------------------------------------------------------
'''


def account_upgrade_sub() -> str:
    """An older back-end has the accounts without parents: place them in the tree, mark the main
    accounts (no entries) and the accounts the automatic entries use. Safe to run every time."""
    from schema import ACCOUNT_TREE
    out = ["Private Sub UpgradeAccountTree()",
           "    ' accounts of an older back-end without a parent go to their place in the tree"]
    for code, name, kind, parent, posting, system in ACCOUNT_TREE:
        if parent is not None:
            out.append(f'    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = {parent} WHERE [AccountCode] = {code} '
                       f'AND [ParentCode] Is Null", dbFailOnError')
    headers = ", ".join(str(r[0]) for r in ACCOUNT_TREE if not r[4])
    system = ", ".join(str(r[0]) for r in ACCOUNT_TREE if r[5])
    out.append(f'    m_db.Execute "UPDATE [Accounts] SET [IsPosting] = False WHERE [AccountCode] IN ({headers})", '
               f'dbFailOnError')
    out.append(f'    m_db.Execute "UPDATE [Accounts] SET [IsSystem] = True WHERE [AccountCode] IN ({system}) OR '
               f'[AccountCode] BETWEEN 110001 AND 119999 OR [AccountCode] BETWEEN 120001 AND 129999 OR '
               f'[AccountCode] BETWEEN 530001 AND 539999", dbFailOnError')
    out.append("End Sub")
    return "\n".join(out)


def rule_upgrade_sub() -> str:
    """The rules of schema.RULE_UPGRADES, set again on an existing back-end (BuildSchema keeps the
    fields that exist, so a changed rule would otherwise stay as it was)."""
    from schema import RULE_UPGRADES
    out = ["Private Sub UpgradeFieldRules()"]
    for tbl, name in RULE_UPGRADES:
        f = next(x for x in table(tbl).fields if x.name == name)
        out.append(f"    SetFieldRule {vba_str(tbl)}, {vba_str(name)}, {vba_str(f.rule)}, {vba_str(f.rule_text)}")
    out.append("End Sub")
    return "\n".join(out)


def english_names_sub() -> str:
    """The English names of the system rows (tools/master_en.py), for a new or an older back-end:
    only the names still empty are filled, so a name the user changed is kept. Safe to run every time."""
    from master_en import ENGLISH_NAMES, EXTRA_NAMES
    out = ["Private Sub SeedEnglishNames()"]
    for tbl, (en_field, _, key, names) in ENGLISH_NAMES.items():
        for k, en in names.items():
            sql = (f"UPDATE [{tbl}] SET [{en_field}] = {sql_value(en)} WHERE [{key}] = {sql_value(k)} "
                   f"AND [{en_field}] Is Null")
            out.append(f"    m_db.Execute {vba_str(sql)}, dbFailOnError")
        for en_field, ar_field, by_arabic in EXTRA_NAMES.get(tbl, []):
            for ar, en in by_arabic.items():
                sql = (f"UPDATE [{tbl}] SET [{en_field}] = {sql_value(en)} WHERE [{ar_field}] = {sql_value(ar)} "
                       f"AND [{en_field}] Is Null")
                out.append(f"    m_db.Execute {vba_str(sql)}, dbFailOnError")
    out.append("End Sub")
    return "\n".join(out)


def long_const(name: str, text: str, sep: str) -> str:
    """A string constant split at sep into pieces joined with & (a VBA line stays under 1000 characters)."""
    pieces, cur = [], ""
    for item in text.split(sep):
        if cur and len(cur) + len(item) > 400:
            pieces.append(cur + sep)
            cur = item
        else:
            cur = f"{cur}{sep}{item}" if cur else item
    pieces.append(cur)
    return f"Private Const {name} As String = " + " & _\n    ".join(vba_str(p) for p in pieces)


def build_vba() -> str:
    names = ",".join(t.name for t in TABLES)
    field_counts = ";".join(f"{t.name}={len(t.fields)}" for t in TABLES)
    seed_counts = ";".join(f"{t.name}={len(t.seed_rows)}" for t in TABLES if t.seed_rows)
    consts = "\n".join([
        long_const("SCHEMA_TABLES", names, ","),
        long_const("EXPECTED_FIELD_COUNTS", field_counts, ";"),
        long_const("EXPECTED_SEED_COUNTS", seed_counts, ";"),
    ])
    parts = [VBA_HEADER.replace("'@@SCHEMA_CONSTANTS@@", consts)]

    parts.append("Private Sub CreateAllTables()\n" +
                 "\n".join(f"    CreateTable_{t.name}" for t in TABLES) + "\nEnd Sub\n")
    for t in TABLES:
        parts.append(vba_table_sub(t) + "\n")

    seeded = [t for t in TABLES if t.seed_rows]
    parts.append("'" + "-" * 78 + "\n' Generated: lookup / initial data\n'" + "-" * 78)
    parts.append("Private Sub SeedAll()\n" +
                 "\n".join(f"    Seed_{t.name}" for t in seeded) + "\nEnd Sub\n")
    for t in seeded:
        parts.append(vba_seed_sub(t) + "\n")
    parts.append(account_upgrade_sub() + "\n")
    parts.append(english_names_sub() + "\n")
    parts.append(rule_upgrade_sub() + "\n")
    return "\n".join(parts)


# --------------------------------------------------------------------------
# Markdown reference
# --------------------------------------------------------------------------
def md_escape(s) -> str:
    if s is None or s == "":
        return ""
    return str(s).replace("|", "\\|")


def build_reference_md() -> str:
    out = ["# مرجع الجداول (Tables Reference)",
           "",
           "> ملف مُولَّد تلقائيًا من `tools/schema.py` بواسطة `tools/generate.py` – لا تعدّله يدويًا.",
           "",
           f"عدد الجداول: **{len(TABLES)}** | عدد الحقول: **{sum(len(t.fields) for t in TABLES)}**",
           "",
           "## الفهرس",
           ""]
    for i, t in enumerate(TABLES, 1):
        out.append(f"{i}. [`{t.name}`](#{t.name.lower()}) – {t.caption}")
    out.append("")
    for t in TABLES:
        out += [f"## {t.name}", "", f"**{t.caption}** – {t.purpose}", "",
                "| # | الحقل | النوع | الحجم | إلزامي | افتراضي | قاعدة التحقق | يرتبط بـ | الوصف |",
                "|---|---|---|---|---|---|---|---|---|"]
        for i, f in enumerate(t.fields, 1):
            name = f"**{f.name}** 🔑" if f.name in t.pk else f.name
            desc = f.caption + (f" – {f.note}" if f.note else "")
            out.append("| {i} | {n} | {k} | {s} | {r} | {d} | {v} | {fk} | {desc} |".format(
                i=i, n=name, k=KIND_LABEL[f.kind], s=f.size or "",
                r="✔" if (f.required and f.kind not in ("AUTO", "BOOL")) else "",
                d=f"`{md_escape(f.default)}`" if f.default else "",
                v=f"`{md_escape(f.rule)}`" if f.rule else "",
                fk=f"`{f.fk}`" if f.fk else "", desc=md_escape(desc)))
        idx_lines = [f"- المفتاح الأساسي: `{', '.join(t.pk)}`"]
        for ix in t.indexes:
            kind = "فريد" if ix.unique else "عادي"
            extra = " (يتجاهل الفارغ)" if ix.ignore_nulls else ""
            idx_lines.append(f"- فهرس {kind}{extra}: `{', '.join(ix.fields)}`")
        if t.rule:
            idx_lines.append(f"- قاعدة تحقق على مستوى الجدول: `{md_escape(t.rule)}` – {t.rule_text}")
        if t.seed_rows:
            idx_lines.append(f"- بيانات أساسية: {len(t.seed_rows)} سجل")
        out += ["", *idx_lines, ""]
    return "\n".join(out)


# --------------------------------------------------------------------------
def write(path, text, encoding="utf-8", newline="\n"):
    full = os.path.join(ROOT, path)
    os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full, "w", encoding=encoding, newline=newline) as fh:
        fh.write(text if text.endswith("\n") else text + "\n")
    print("wrote", path)


def main():
    vba = build_vba()
    write("src/vba/modBuildSchema.bas", vba)
    # Access imports .bas files in the system ANSI code page -> Windows-1256 + CRLF
    write("dist/vba/modBuildSchema.bas", vba, encoding="cp1256", newline="\r\n")
    write("docs/02-Tables-Reference.md", build_reference_md())
    write("docs/dev/Screens-Index.md", gen_forms.build_screens_md())

    rel = gen_relations.build_relations_vba()
    write("src/vba/modBuildRelations.bas", rel)
    write("dist/vba/modBuildRelations.bas", rel, encoding="cp1256", newline="\r\n")
    write("docs/03-Relationships-Reference.md", gen_relations.build_relations_md())

    qry = gen_queries.build_queries_vba()
    write("src/vba/modBuildQueries.bas", qry)
    write("dist/vba/modBuildQueries.bas", qry, encoding="cp1256", newline="\r\n")
    write("docs/04-Queries-Reference.md", gen_queries.build_queries_md())

    for name, text in (("modBuildForms", gen_forms.build_forms_vba()),
                       ("modAppData", gen_forms.build_appdata_vba()),
                       ("modQRCode", gen_qr.build_qr_vba()),
                       ("modBuildReports", gen_reports.build_reports_vba()),
                       ("modTestSales", gen_test_sales.build_test_sales_vba()),
                       ("modTestPurchases", gen_test_purchases.build_test_purchases_vba()),
                       ("modTestSecurity", gen_test_security.build_test_security_vba()),
                       ("modDemoData", gen_demo.build_demo_vba()),
                       ("modLang", gen_lang.build_lang_vba()),
                       ("modZatcaData", gen_zatca.build_zatca_data_vba())):
        write(f"src/vba/{name}.bas", text)
        write(f"dist/vba/{name}.bas", text, encoding="cp1256", newline="\r\n")
    for n, text in enumerate(gen_lang.data_modules(), 1):      # the dictionary of modLang
        write(f"src/vba/modLangData{n}.bas", text)
        write(f"dist/vba/modLangData{n}.bas", text, encoding="cp1256", newline="\r\n")
    for folder in ("src/vba", "dist/vba"):                      # parts left over from a longer dictionary
        for name in os.listdir(os.path.join(ROOT, folder)):
            m = re.fullmatch(r"modLangData(\d+)\.bas", name)
            if m and int(m.group(1)) > n:
                os.remove(os.path.join(ROOT, folder, name))

    # hand-written modules: readable copy in src/, import copy in dist/
    for name in STATIC_MODULES:
        with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
            write(f"dist/vba/{name}.bas", fh.read(), encoding="cp1256", newline="\r\n")


if __name__ == "__main__":
    main()
