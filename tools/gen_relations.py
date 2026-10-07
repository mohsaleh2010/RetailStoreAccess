"""Generate modBuildRelations.bas and the relationships reference doc."""

import re

from generate_common import vba_str
from relations import (relations, retired_relations, UNENFORCED, REJECTED_INSERTS, REJECTED_DELETES, CASCADE_PRODUCT_SQL,
                       CASCADE_INVOICE_SQL, CASCADE_LINE_SQL, CASCADE_DELETE_PRODUCT_SQL,
                       CASCADE_DELETE_INVOICE_SQL, CASCADE_COUNT_LINES_SQL, RENAME_KEY_SQL,
                       COUNT_RENAMED_SQL, COUNT_OLD_KEY_SQL)
from schema import TABLES


def vba_fmt(sql: str) -> str:
    """Turn a python format string with {invoice}/{product}/{line} into a VBA expression."""
    expr = vba_str(sql)
    for name, var in (("invoice", "invoiceID"), ("product", "productID"), ("line", "lineNo")):
        expr = expr.replace("{" + name + "}", f'" & {var} & "')
    return expr.replace(' & ""', "")


VBA_TEMPLATE = r'''Attribute VB_Name = "modBuildRelations"
'==============================================================================
' modBuildRelations  -  Retail Store Management System (Phase 3: Relationships)
'
' GENERATED FILE - do not edit by hand.
' Source of truth: tools/schema.py (fk) + tools/relations.py
'                  ->  python3 tools/generate.py
'
' Public procedures (run from the Immediate window, Ctrl+G):
'   BuildRelationships     creates every missing relationship in the back-end
'                          with Enforce Referential Integrity.
'   TestRelationships      checks every relationship and proves that orphan
'                          records are rejected (all test data is rolled back).
'   DropRelationships      DEVELOPMENT ONLY: removes the system relationships.
'
' Run BuildSchema (modBuildSchema) first. Close every table/form before running.
'==============================================================================
Option Compare Database
Option Explicit

Private Const BE_FILE_NAME As String = "RetailStore_BE.accdb"
Private Const MSG_RTL As Long = &H180000           ' vbMsgBoxRight + vbMsgBoxRtlReading
Private Const REL_CASCADE_UPDATE As Long = 256     ' dbRelationUpdateCascade
Private Const REL_CASCADE_DELETE As Long = 4096    ' dbRelationDeleteCascade
Private Const REL_DONT_ENFORCE As Long = 2         ' dbRelationDontEnforce
Private Const ERR_HAS_RELATED_RECORDS As Long = 3200
Private Const ERR_RELATED_RECORD_REQUIRED As Long = 3201
Private Const EXPECTED_RELATION_COUNT As Long = @@COUNT@@

Private m_db As DAO.Database
Private m_created As Long
Private m_skipped As Long
Private m_failed As Long
Private m_passed As Long
Private m_report As String

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildRelationships(Optional ByVal BackEndPath As String = "") As Boolean
    Dim spec As Variant
    On Error GoTo EH
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    If Len(Dir$(BackEndPath)) = 0 Then
        MsgBox "ملف البيانات غير موجود: " & BackEndPath & vbCrLf & "شغّل BuildSchema أولًا.", _
               vbCritical + MSG_RTL, "BuildRelationships"
        Exit Function
    End If

    m_created = 0: m_skipped = 0: m_failed = 0: m_report = ""
    Debug.Print "=== BuildRelationships  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Set m_db = DBEngine.OpenDatabase(BackEndPath)

    ' relationships an older version created and this one no longer enforces: they would
    ' hold index slots (Access allows 32 per table, a relationship takes one on each side)
    For Each spec In RetiredRelations()
        If RelationExistsIn(m_db, spec) Then
            m_db.Relations.Delete spec
            Debug.Print "  - أُزيلت: " & spec
        End If
    Next

    For Each spec In RelationSpecs()
        AddRelation spec(0), spec(1), spec(2), spec(3), spec(4), spec(5)
    Next

    m_db.Close
    Set m_db = Nothing
    Debug.Print "--- جديدة: " & m_created & " | موجودة: " & m_skipped & " | فشلت: " & m_failed

    If m_failed = 0 Then
        MsgBox "تم إنشاء العلاقات بنجاح." & vbCrLf & vbCrLf & _
               "علاقات جديدة: " & m_created & vbCrLf & _
               "علاقات موجودة مسبقًا: " & m_skipped & vbCrLf & vbCrLf & _
               "الخطوة التالية: شغّل TestRelationships", vbInformation + MSG_RTL, "BuildRelationships"
        BuildRelationships = True
    Else
        MsgBox "لم يتم إنشاء " & m_failed & " علاقة:" & vbCrLf & vbCrLf & Left$(m_report, 900) & _
               vbCrLf & "التفاصيل في نافذة Immediate (Ctrl+G).", vbExclamation + MSG_RTL, _
               "BuildRelationships"
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    If Not m_db Is Nothing Then m_db.Close
    Set m_db = Nothing
    MsgBox errText, vbCritical + MSG_RTL, "BuildRelationships"
End Function

Public Function TestRelationships(Optional ByVal BackEndPath As String = "") As Boolean
    Dim ws As DAO.Workspace, db As DAO.Database
    Dim inTrans As Boolean, productID As Long, invoiceID As Long, lineNo As Long

    On Error GoTo EH
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    m_passed = 0: m_failed = 0: m_report = ""
    Debug.Print "=== TestRelationships  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(BackEndPath)

    ' 1) every relationship exists, is enforced and has the right cascade options
    CheckRelationsExist db

    ws.BeginTrans
    inTrans = True

    ' 2) records pointing to something that does not exist are rejected
@@REJECTED_INSERTS@@

    ' 3) cascade scenario: a product and an invoice with two lines
    db.Execute @@CASCADE_PRODUCT_SQL@@, dbFailOnError
    productID = LastID(db)
    db.Execute @@CASCADE_INVOICE_SQL@@, dbFailOnError
    invoiceID = LastID(db)
    For lineNo = 1 To 2
        db.Execute @@CASCADE_LINE_SQL@@, dbFailOnError
    Next

    ' 4) master records that are in use cannot be deleted
@@REJECTED_DELETES@@
    ExpectRejected db, @@CASCADE_DELETE_PRODUCT_SQL@@, "حذف منتج له مبيعات"

    ' 5) deleting an invoice removes its lines automatically (cascade delete)
    db.Execute @@CASCADE_DELETE_INVOICE_SQL@@, dbFailOnError
    Call Record(ScalarLong(db, @@CASCADE_COUNT_LINES_SQL@@) = 0, _
                "حذف الفاتورة يحذف أسطرها تلقائيًا")

    ' 6) renaming a permission key updates the roles that use it (cascade update)
    db.Execute @@RENAME_KEY_SQL@@, dbFailOnError
    Call Record(ScalarLong(db, @@COUNT_RENAMED_SQL@@) > 0 And _
                ScalarLong(db, @@COUNT_OLD_KEY_SQL@@) = 0, _
                "تغيير رمز صلاحية ينعكس على صلاحيات الأدوار")

    ws.Rollback
    inTrans = False
    db.Close

    Debug.Print "--- نجح: " & m_passed & " | فشل: " & m_failed & " (تم التراجع عن كل بيانات الاختبار)"
    If m_failed = 0 Then
        Call ResultBox("جميع اختبارات العلاقات ناجحة (" & m_passed & " اختبارًا)." & vbCrLf & _
               "لا يمكن إدخال سجلات يتيمة، ولا حذف بيانات مستخدمة." & vbCrLf & _
               "لم تُترك أي بيانات اختبار في القاعدة.", vbInformation + MSG_RTL, "TestRelationships")
        TestRelationships = True
    Else
        Call ResultBox("نجح " & m_passed & " وفشل " & m_failed & ":" & vbCrLf & vbCrLf & _
               Left$(m_report, 900), vbExclamation + MSG_RTL, "TestRelationships")
    End If
    Exit Function

EH:
    Dim errText As String
    errText = "خطأ غير متوقع " & Err.Number & ": " & Err.Description
    Debug.Print errText
    On Error Resume Next
    If inTrans Then ws.Rollback
    db.Close
    Call ResultBox(errText & vbCrLf & "تم التراجع عن بيانات الاختبار.", vbCritical + MSG_RTL, _
           "TestRelationships")
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

Public Sub DropRelationships(Optional ByVal BackEndPath As String = "")
    ' DEVELOPMENT ONLY - removes the relationships created by BuildRelationships.
    Dim db As DAO.Database, spec As Variant, removed As Long

    If InputBox("سيتم حذف علاقات النظام (البيانات لا تُحذف)." & vbCrLf & _
                "للتأكيد اكتب DELETE", "DropRelationships") <> "DELETE" Then Exit Sub
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    Set db = DBEngine.OpenDatabase(BackEndPath)
    For Each spec In RelationSpecs()
        If RelationExistsIn(db, spec(0)) Then
            db.Relations.Delete spec(0)
            removed = removed + 1
        End If
    Next
    db.Close
    MsgBox "تم حذف " & removed & " علاقة.", vbInformation + MSG_RTL
End Sub

'------------------------------------------------------------------------------
' Building
'------------------------------------------------------------------------------
Private Sub AddRelation(ByVal RelName As String, ByVal ParentTable As String, _
                        ByVal ParentField As String, ByVal ChildTable As String, _
                        ByVal ChildField As String, ByVal Attributes As Long)
    Dim rel As DAO.Relation, existing As String, orphans As Long
    On Error GoTo EH

    If RelationExistsIn(m_db, RelName) Then
        m_skipped = m_skipped + 1
        Debug.Print "  = موجودة مسبقًا: " & RelName
        Exit Sub
    End If
    existing = SameLinkName(m_db, ParentTable, ParentField, ChildTable, ChildField)
    If Len(existing) > 0 Then
        m_skipped = m_skipped + 1
        Debug.Print "  = موجودة باسم آخر (" & existing & "): " & RelName
        Exit Sub
    End If

    orphans = CountOrphans(m_db, ParentTable, ParentField, ChildTable, ChildField)
    If orphans > 0 Then
        Fail RelName & ": يوجد " & orphans & " سجل في " & ChildTable & "." & ChildField & _
             " يشير إلى سجل غير موجود في " & ParentTable & ". أصلحها ثم أعد التشغيل."
        Exit Sub
    End If

    Set rel = m_db.CreateRelation(RelName, ParentTable, ChildTable, Attributes)
    rel.Fields.Append rel.CreateField(ParentField)
    rel.Fields(ParentField).ForeignName = ChildField
    m_db.Relations.Append rel
    m_created = m_created + 1
    Debug.Print "  + " & RelName
    Exit Sub

EH:
    Fail RelName & ": خطأ " & Err.Number & " - " & Err.Description
End Sub

Private Function CountOrphans(ByVal db As DAO.Database, ByVal ParentTable As String, _
                              ByVal ParentField As String, ByVal ChildTable As String, _
                              ByVal ChildField As String) As Long
    CountOrphans = ScalarLong(db, _
        "SELECT COUNT(*) FROM [" & ChildTable & "] AS c LEFT JOIN [" & ParentTable & "] AS p " & _
        "ON c.[" & ChildField & "] = p.[" & ParentField & "] " & _
        "WHERE c.[" & ChildField & "] Is Not Null AND p.[" & ParentField & "] Is Null")
End Function

'------------------------------------------------------------------------------
' Testing
'------------------------------------------------------------------------------
Private Sub CheckRelationsExist(ByVal db As DAO.Database)
    Dim spec As Variant, rel As DAO.Relation, checked As Long, problem As String

    For Each spec In RelationSpecs()
        checked = checked + 1
        problem = ""
        If Not RelationExistsIn(db, spec(0)) Then
            problem = "غير موجودة"
        Else
            Set rel = db.Relations(spec(0))
            If StrComp(rel.Table, spec(1), vbTextCompare) <> 0 _
               Or StrComp(rel.ForeignTable, spec(3), vbTextCompare) <> 0 Then
                problem = "تربط جداول مختلفة"
            ElseIf StrComp(rel.Fields(0).Name, spec(2), vbTextCompare) <> 0 _
                   Or StrComp(rel.Fields(0).ForeignName, spec(4), vbTextCompare) <> 0 Then
                problem = "تربط حقولًا مختلفة"
            ElseIf (rel.Attributes And REL_DONT_ENFORCE) <> 0 Then
                problem = "لا تفرض التكامل المرجعي"
            ElseIf (rel.Attributes And (REL_CASCADE_DELETE Or REL_CASCADE_UPDATE)) <> spec(5) Then
                problem = "خيارات الحذف/التحديث المتتالي غير صحيحة"
            End If
        End If
        If Len(problem) = 0 Then
            m_passed = m_passed + 1
        Else
            Fail "العلاقة " & spec(0) & ": " & problem
        End If
    Next
    Call Record(checked = EXPECTED_RELATION_COUNT, "عدد العلاقات المعرّفة = " & EXPECTED_RELATION_COUNT)
    Debug.Print "فُحصت " & checked & " علاقة"
End Sub

Private Sub ExpectRejected(ByVal db As DAO.Database, ByVal Sql As String, ByVal Label As String)
    Dim errNo As Long, errText As String
    On Error Resume Next
    db.Execute Sql, dbFailOnError
    errNo = Err.Number
    errText = Err.Description
    On Error GoTo 0

    If errNo = ERR_RELATED_RECORD_REQUIRED Or errNo = ERR_HAS_RELATED_RECORDS Then
        Record True, "مرفوض كما هو متوقع: " & Label
    ElseIf errNo = 0 Then
        Record False, "تم قبول عملية يجب رفضها: " & Label
    Else
        Record False, Label & ": رُفضت لسبب آخر (" & errNo & " " & errText & ")"
    End If
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

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function DefaultBackEndPath() As String
    DefaultBackEndPath = CurrentProject.Path & "\" & BE_FILE_NAME
End Function

Private Function RelationExistsIn(ByVal db As DAO.Database, ByVal RelName As String) As Boolean
    Dim rel As DAO.Relation
    For Each rel In db.Relations
        If StrComp(rel.Name, RelName, vbTextCompare) = 0 Then
            RelationExistsIn = True
            Exit Function
        End If
    Next
End Function

Private Function SameLinkName(ByVal db As DAO.Database, ByVal ParentTable As String, _
                              ByVal ParentField As String, ByVal ChildTable As String, _
                              ByVal ChildField As String) As String
    ' Name of a relationship (e.g. drawn by hand) that already links the same fields.
    Dim rel As DAO.Relation
    For Each rel In db.Relations
        If StrComp(rel.Table, ParentTable, vbTextCompare) = 0 _
           And StrComp(rel.ForeignTable, ChildTable, vbTextCompare) = 0 Then
            If rel.Fields.Count = 1 Then
                If StrComp(rel.Fields(0).Name, ParentField, vbTextCompare) = 0 _
                   And StrComp(rel.Fields(0).ForeignName, ChildField, vbTextCompare) = 0 Then
                    SameLinkName = rel.Name
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function ScalarLong(ByVal db As DAO.Database, ByVal Sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(Sql, dbOpenSnapshot)
    ScalarLong = Nz(rs(0), 0)
    rs.Close
End Function

Private Function LastID(ByVal db As DAO.Database) As Long
    LastID = ScalarLong(db, "SELECT @@IDENTITY")
End Function

'------------------------------------------------------------------------------
' Generated: relationship list
' Array(Name, ParentTable, ParentField, ChildTable, ChildField, Attributes)
'------------------------------------------------------------------------------
Private Function RelationSpecs() As Collection
    Dim c As New Collection
@@SPECS@@
    Set RelationSpecs = c
End Function

Private Function RetiredRelations() As Variant
    RetiredRelations = Array(@@RETIRED@@)
End Function
'''


def build_relations_vba() -> str:
    rels = relations()
    specs = "\n".join(
        "    c.Add Array({}, {}, {}, {}, {}, {}&)".format(
            vba_str(r.name), vba_str(r.parent), vba_str(r.parent_field), vba_str(r.child),
            vba_str(r.child_field), r.attributes)
        for r in rels)
    ins = "\n".join(f"    ExpectRejected db, {vba_str(sql)}, _\n                   {vba_str(label)}"
                    for label, sql in REJECTED_INSERTS)
    dels = "\n".join(f"    ExpectRejected db, {vba_str(sql)}, {vba_str(label)}"
                     for label, sql in REJECTED_DELETES)
    text = VBA_TEMPLATE
    for key, value in {
        "@@COUNT@@": str(len(rels)),
        "@@SPECS@@": specs,
        "@@RETIRED@@": ", _\n                             ".join(vba_str(n) for n in retired_relations()),
        "@@REJECTED_INSERTS@@": ins,
        "@@REJECTED_DELETES@@": dels,
        "@@CASCADE_PRODUCT_SQL@@": vba_str(CASCADE_PRODUCT_SQL),
        "@@CASCADE_INVOICE_SQL@@": vba_str(CASCADE_INVOICE_SQL),
        "@@CASCADE_LINE_SQL@@": vba_fmt(CASCADE_LINE_SQL),
        "@@CASCADE_DELETE_PRODUCT_SQL@@": vba_fmt(CASCADE_DELETE_PRODUCT_SQL),
        "@@CASCADE_DELETE_INVOICE_SQL@@": vba_fmt(CASCADE_DELETE_INVOICE_SQL),
        "@@CASCADE_COUNT_LINES_SQL@@": vba_fmt(CASCADE_COUNT_LINES_SQL),
        "@@RENAME_KEY_SQL@@": vba_str(RENAME_KEY_SQL),
        "@@COUNT_RENAMED_SQL@@": vba_str(COUNT_RENAMED_SQL),
        "@@COUNT_OLD_KEY_SQL@@": vba_str(COUNT_OLD_KEY_SQL),
    }.items():
        text = text.replace(key, value)
    assert not re.search(r"@@[A-Z_]+@@", text), "unfilled template placeholder"
    return text


def build_relations_md() -> str:
    rels = relations()
    captions = {t.name: t.caption for t in TABLES}
    out = ["# مرجع العلاقات (Relationships Reference)", "",
           "> ملف مُولَّد تلقائيًا من `tools/schema.py` و`tools/relations.py` – لا تعدّله يدويًا.", "",
           f"عدد العلاقات: **{len(rels)}** – جميعها مع **Enforce Referential Integrity**. "
           f"الحذف المتتالي: **{sum(r.cascade_delete for r in rels)}**، "
           f"التحديث المتتالي: **{sum(r.cascade_update for r in rels)}**.", "",
           "حقول «سجّلها الموظف» التالية تشير إلى `Employees` بلا علاقة مفروضة: Access يسمح بـ32 فهرسًا "
           "لكل جدول، وكل علاقة تُحسب فهرسًا على طرفيها. البرنامج يكتب فيها المستخدم الحالي، "
           "والمستخدمون يُعطَّلون ولا يُحذفون: "
           + "، ".join(f"`{c}.{f}`" for c, f in sorted(UNENFORCED)) + ".", "",
           "## مخطط الكيانات والعلاقات", "", "```mermaid", "erDiagram"]
    for r in rels:
        card = "||--o{" if r.required else "|o--o{"
        out.append(f'    {r.parent} {card} {r.child} : "{r.child_field}"')
    out += ["```", "",
            "`||--o{` = إلزامي (كل سجل في الجدول الفرعي يجب أن يرتبط بسجل في الأصلي)، "
            "`|o--o{` = اختياري (الحقل يمكن أن يكون فارغًا).", "",
            "## قائمة العلاقات", "",
            "| # | اسم العلاقة | الجدول الأصلي (1) | المفتاح | الجدول الفرعي (∞) | الحقل المرتبط | إلزامي | القاعدة |",
            "|---|---|---|---|---|---|---|---|"]
    for i, r in enumerate(rels, 1):
        out.append(f"| {i} | `{r.name}` | {r.parent} ({captions[r.parent]}) | `{r.parent_field}` | "
                   f"{r.child} ({captions[r.child]}) | `{r.child_field}` | "
                   f"{'✔' if r.required else ''} | {r.rule_ar} |")
    return "\n".join(out)
