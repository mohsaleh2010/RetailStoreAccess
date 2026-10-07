Attribute VB_Name = "modBuildRelations"
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
Private Const EXPECTED_RELATION_COUNT As Long = 128

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
    ExpectRejected db, "INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) VALUES ('TEST-RI-1', 999999, 1)", _
                   "فاتورة بيع لعميل غير موجود"
    ExpectRejected db, "INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) VALUES ('TEST-RI-2', 1, 999999)", _
                   "فاتورة بيع بموظف غير موجود"
    ExpectRejected db, "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice]) VALUES (999999, 1, 999999, 1, 1)", _
                   "سطر بيع لفاتورة غير موجودة"
    ExpectRejected db, "INSERT INTO [PurchaseInvoices] ([InvoiceNumber], [SupplierID], [EmployeeID]) VALUES ('TEST-RI-3', 999999, 1)", _
                   "فاتورة شراء لمورد غير موجود"
    ExpectRejected db, "INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID]) VALUES ('TEST-RI-X', 'x', 999999)", _
                   "منتج بتصنيف غير موجود"
    ExpectRejected db, "INSERT INTO [Products] ([ProductCode], [ProductName], [SupplierID]) VALUES ('TEST-RI-Y', 'y', 999999)", _
                   "منتج بمورد غير موجود"
    ExpectRejected db, "INSERT INTO [InventoryTransactions] ([ProductID], [TransactionTypeID], [Quantity]) VALUES (999999, 1, 1)", _
                   "حركة مخزون لمنتج غير موجود"
    ExpectRejected db, "INSERT INTO [CustomerPayments] ([PaymentNumber], [CustomerID], [Amount], [EmployeeID]) VALUES ('TEST-RI-4', 999999, 10, 1)", _
                   "سند قبض لعميل غير موجود"
    ExpectRejected db, "INSERT INTO [SupplierPayments] ([PaymentNumber], [SupplierID], [Amount], [EmployeeID]) VALUES ('TEST-RI-5', 999999, 10, 1)", _
                   "سند صرف لمورد غير موجود"
    ExpectRejected db, "INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseTypeID], [Amount], [TotalAmount], [EmployeeID]) VALUES ('TEST-RI-6', 999999, 10, 10, 1)", _
                   "مصروف بنوع غير موجود"
    ExpectRejected db, "INSERT INTO [Employees] ([EmployeeName], [Username], [RoleID]) VALUES ('x', 'test_ri_user', 999999)", _
                   "مستخدم بدور غير موجود"
    ExpectRejected db, "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'NO_SUCH_KEY')", _
                   "صلاحية غير معرّفة لدور"

    ' 3) cascade scenario: a product and an invoice with two lines
    db.Execute "INSERT INTO [Products] ([ProductCode], [ProductName]) VALUES ('TEST-RI-P', 'TEST')", dbFailOnError
    productID = LastID(db)
    db.Execute "INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) VALUES ('TEST-RI-INV', 1, 1)", dbFailOnError
    invoiceID = LastID(db)
    For lineNo = 1 To 2
        db.Execute "INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], [Quantity], [UnitPrice]) VALUES (" & invoiceID & ", " & lineNo & ", " & productID & ", 2, 10)", dbFailOnError
    Next

    ' 4) master records that are in use cannot be deleted
    ExpectRejected db, "DELETE FROM [Roles] WHERE [RoleID] = 1", "حذف دور مرتبط بمستخدمين"
    ExpectRejected db, "DELETE FROM [Customers] WHERE [CustomerID] = 1", "حذف العميل النقدي (مرتبط بالإعدادات)"
    ExpectRejected db, "DELETE FROM [Categories] WHERE [CategoryID] = 1", "حذف تصنيف تستخدمه منتجات"
    ExpectRejected db, "DELETE FROM [Employees] WHERE [EmployeeID] = 1", "حذف المستخدم admin وله عمليات"
    ExpectRejected db, "DELETE FROM [Products] WHERE [ProductID] = " & productID, "حذف منتج له مبيعات"

    ' 5) deleting an invoice removes its lines automatically (cascade delete)
    db.Execute "DELETE FROM [SalesInvoices] WHERE [SalesInvoiceID] = " & invoiceID, dbFailOnError
    Call Record(ScalarLong(db, "SELECT COUNT(*) FROM [SalesInvoiceDetails] WHERE [SalesInvoiceID] = " & invoiceID) = 0, _
                "حذف الفاتورة يحذف أسطرها تلقائيًا")

    ' 6) renaming a permission key updates the roles that use it (cascade update)
    db.Execute "UPDATE [Permissions] SET [PermissionKey] = 'SALES_POS_TEST' WHERE [PermissionKey] = 'SALES_POS'", dbFailOnError
    Call Record(ScalarLong(db, "SELECT COUNT(*) FROM [RolePermissions] WHERE [PermissionKey] = 'SALES_POS_TEST'") > 0 And _
                ScalarLong(db, "SELECT COUNT(*) FROM [RolePermissions] WHERE [PermissionKey] = 'SALES_POS'") = 0, _
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
    c.Add Array("FK_Settings_DefaultCustomerID", "Customers", "CustomerID", "Settings", "DefaultCustomerID", 0&)
    c.Add Array("FK_Settings_DefaultBankID", "Banks", "BankID", "Settings", "DefaultBankID", 0&)
    c.Add Array("FK_RolePermissions_RoleID", "Roles", "RoleID", "RolePermissions", "RoleID", 4096&)
    c.Add Array("FK_RolePermissions_PermissionKey", "Permissions", "PermissionKey", "RolePermissions", "PermissionKey", 4352&)
    c.Add Array("FK_Employees_RoleID", "Roles", "RoleID", "Employees", "RoleID", 0&)
    c.Add Array("FK_Employees_CashBoxID", "CashBoxes", "CashBoxID", "Employees", "CashBoxID", 0&)
    c.Add Array("FK_Employees_CostCenterID", "CostCenters", "CostCenterID", "Employees", "CostCenterID", 0&)
    c.Add Array("FK_UserScreens_EmployeeID", "Employees", "EmployeeID", "UserScreens", "EmployeeID", 0&)
    c.Add Array("FK_UserScreens_ScreenName", "Screens", "ScreenName", "UserScreens", "ScreenName", 256&)
    c.Add Array("FK_Activations_EmployeeID", "Employees", "EmployeeID", "Activations", "EmployeeID", 0&)
    c.Add Array("FK_Products_CategoryID", "Categories", "CategoryID", "Products", "CategoryID", 0&)
    c.Add Array("FK_Products_UnitID", "Units", "UnitID", "Products", "UnitID", 0&)
    c.Add Array("FK_Products_SupplierID", "Suppliers", "SupplierID", "Products", "SupplierID", 0&)
    c.Add Array("FK_SalesInvoices_CustomerID", "Customers", "CustomerID", "SalesInvoices", "CustomerID", 0&)
    c.Add Array("FK_SalesInvoices_EmployeeID", "Employees", "EmployeeID", "SalesInvoices", "EmployeeID", 0&)
    c.Add Array("FK_SalesInvoices_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "SalesInvoices", "PaymentMethodID", 0&)
    c.Add Array("FK_SalesInvoices_CashBoxID", "CashBoxes", "CashBoxID", "SalesInvoices", "CashBoxID", 0&)
    c.Add Array("FK_SalesInvoices_BankID", "Banks", "BankID", "SalesInvoices", "BankID", 0&)
    c.Add Array("FK_SalesInvoices_CostCenterID", "CostCenters", "CostCenterID", "SalesInvoices", "CostCenterID", 0&)
    c.Add Array("FK_SalesInvoiceDetails_SalesInvoiceID", "SalesInvoices", "SalesInvoiceID", "SalesInvoiceDetails", "SalesInvoiceID", 4096&)
    c.Add Array("FK_SalesInvoiceDetails_ProductID", "Products", "ProductID", "SalesInvoiceDetails", "ProductID", 0&)
    c.Add Array("FK_SalesReturns_SalesInvoiceID", "SalesInvoices", "SalesInvoiceID", "SalesReturns", "SalesInvoiceID", 0&)
    c.Add Array("FK_SalesReturns_CustomerID", "Customers", "CustomerID", "SalesReturns", "CustomerID", 0&)
    c.Add Array("FK_SalesReturns_EmployeeID", "Employees", "EmployeeID", "SalesReturns", "EmployeeID", 0&)
    c.Add Array("FK_SalesReturns_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "SalesReturns", "PaymentMethodID", 0&)
    c.Add Array("FK_SalesReturns_CashBoxID", "CashBoxes", "CashBoxID", "SalesReturns", "CashBoxID", 0&)
    c.Add Array("FK_SalesReturns_BankID", "Banks", "BankID", "SalesReturns", "BankID", 0&)
    c.Add Array("FK_SalesReturns_CostCenterID", "CostCenters", "CostCenterID", "SalesReturns", "CostCenterID", 0&)
    c.Add Array("FK_SalesReturnDetails_SalesReturnID", "SalesReturns", "SalesReturnID", "SalesReturnDetails", "SalesReturnID", 4096&)
    c.Add Array("FK_SalesReturnDetails_SalesDetailID", "SalesInvoiceDetails", "SalesDetailID", "SalesReturnDetails", "SalesDetailID", 0&)
    c.Add Array("FK_SalesReturnDetails_ProductID", "Products", "ProductID", "SalesReturnDetails", "ProductID", 0&)
    c.Add Array("FK_PurchaseInvoices_SupplierID", "Suppliers", "SupplierID", "PurchaseInvoices", "SupplierID", 0&)
    c.Add Array("FK_PurchaseInvoices_EmployeeID", "Employees", "EmployeeID", "PurchaseInvoices", "EmployeeID", 0&)
    c.Add Array("FK_PurchaseInvoices_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "PurchaseInvoices", "PaymentMethodID", 0&)
    c.Add Array("FK_PurchaseInvoices_CashBoxID", "CashBoxes", "CashBoxID", "PurchaseInvoices", "CashBoxID", 0&)
    c.Add Array("FK_PurchaseInvoices_BankID", "Banks", "BankID", "PurchaseInvoices", "BankID", 0&)
    c.Add Array("FK_PurchaseInvoiceDetails_PurchaseInvoiceID", "PurchaseInvoices", "PurchaseInvoiceID", "PurchaseInvoiceDetails", "PurchaseInvoiceID", 4096&)
    c.Add Array("FK_PurchaseInvoiceDetails_ProductID", "Products", "ProductID", "PurchaseInvoiceDetails", "ProductID", 0&)
    c.Add Array("FK_PurchaseReturns_PurchaseInvoiceID", "PurchaseInvoices", "PurchaseInvoiceID", "PurchaseReturns", "PurchaseInvoiceID", 0&)
    c.Add Array("FK_PurchaseReturns_SupplierID", "Suppliers", "SupplierID", "PurchaseReturns", "SupplierID", 0&)
    c.Add Array("FK_PurchaseReturns_EmployeeID", "Employees", "EmployeeID", "PurchaseReturns", "EmployeeID", 0&)
    c.Add Array("FK_PurchaseReturns_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "PurchaseReturns", "PaymentMethodID", 0&)
    c.Add Array("FK_PurchaseReturns_CashBoxID", "CashBoxes", "CashBoxID", "PurchaseReturns", "CashBoxID", 0&)
    c.Add Array("FK_PurchaseReturns_BankID", "Banks", "BankID", "PurchaseReturns", "BankID", 0&)
    c.Add Array("FK_PurchaseReturnDetails_PurchaseReturnID", "PurchaseReturns", "PurchaseReturnID", "PurchaseReturnDetails", "PurchaseReturnID", 4096&)
    c.Add Array("FK_PurchaseReturnDetails_PurchaseDetailID", "PurchaseInvoiceDetails", "PurchaseDetailID", "PurchaseReturnDetails", "PurchaseDetailID", 0&)
    c.Add Array("FK_PurchaseReturnDetails_ProductID", "Products", "ProductID", "PurchaseReturnDetails", "ProductID", 0&)
    c.Add Array("FK_CustomerPayments_CustomerID", "Customers", "CustomerID", "CustomerPayments", "CustomerID", 0&)
    c.Add Array("FK_CustomerPayments_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "CustomerPayments", "PaymentMethodID", 0&)
    c.Add Array("FK_CustomerPayments_SalesInvoiceID", "SalesInvoices", "SalesInvoiceID", "CustomerPayments", "SalesInvoiceID", 0&)
    c.Add Array("FK_CustomerPayments_EmployeeID", "Employees", "EmployeeID", "CustomerPayments", "EmployeeID", 0&)
    c.Add Array("FK_CustomerPayments_CashBoxID", "CashBoxes", "CashBoxID", "CustomerPayments", "CashBoxID", 0&)
    c.Add Array("FK_CustomerPayments_BankID", "Banks", "BankID", "CustomerPayments", "BankID", 0&)
    c.Add Array("FK_SupplierPayments_SupplierID", "Suppliers", "SupplierID", "SupplierPayments", "SupplierID", 0&)
    c.Add Array("FK_SupplierPayments_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "SupplierPayments", "PaymentMethodID", 0&)
    c.Add Array("FK_SupplierPayments_PurchaseInvoiceID", "PurchaseInvoices", "PurchaseInvoiceID", "SupplierPayments", "PurchaseInvoiceID", 0&)
    c.Add Array("FK_SupplierPayments_EmployeeID", "Employees", "EmployeeID", "SupplierPayments", "EmployeeID", 0&)
    c.Add Array("FK_SupplierPayments_CashBoxID", "CashBoxes", "CashBoxID", "SupplierPayments", "CashBoxID", 0&)
    c.Add Array("FK_SupplierPayments_BankID", "Banks", "BankID", "SupplierPayments", "BankID", 0&)
    c.Add Array("FK_BankTransactions_BankID", "Banks", "BankID", "BankTransactions", "BankID", 0&)
    c.Add Array("FK_BankTransactions_ToBankID", "Banks", "BankID", "BankTransactions", "ToBankID", 0&)
    c.Add Array("FK_BankTransactions_CashBoxID", "CashBoxes", "CashBoxID", "BankTransactions", "CashBoxID", 0&)
    c.Add Array("FK_BankTransactions_CounterAccount", "Accounts", "AccountCode", "BankTransactions", "CounterAccount", 0&)
    c.Add Array("FK_BankTransactions_EmployeeID", "Employees", "EmployeeID", "BankTransactions", "EmployeeID", 0&)
    c.Add Array("FK_Cheques_CustomerID", "Customers", "CustomerID", "Cheques", "CustomerID", 0&)
    c.Add Array("FK_Cheques_SupplierID", "Suppliers", "SupplierID", "Cheques", "SupplierID", 0&)
    c.Add Array("FK_Cheques_BankID", "Banks", "BankID", "Cheques", "BankID", 0&)
    c.Add Array("FK_Cheques_EmployeeID", "Employees", "EmployeeID", "Cheques", "EmployeeID", 0&)
    c.Add Array("FK_FixedAssets_AssetAccount", "Accounts", "AccountCode", "FixedAssets", "AssetAccount", 0&)
    c.Add Array("FK_FixedAssets_CostCenterID", "CostCenters", "CostCenterID", "FixedAssets", "CostCenterID", 0&)
    c.Add Array("FK_FixedAssets_BankID", "Banks", "BankID", "FixedAssets", "BankID", 0&)
    c.Add Array("FK_FixedAssets_CashBoxID", "CashBoxes", "CashBoxID", "FixedAssets", "CashBoxID", 0&)
    c.Add Array("FK_FixedAssets_CounterAccount", "Accounts", "AccountCode", "FixedAssets", "CounterAccount", 0&)
    c.Add Array("FK_FixedAssets_DisposalBankID", "Banks", "BankID", "FixedAssets", "DisposalBankID", 0&)
    c.Add Array("FK_FixedAssets_DisposalCashBoxID", "CashBoxes", "CashBoxID", "FixedAssets", "DisposalCashBoxID", 0&)
    c.Add Array("FK_FixedAssets_EmployeeID", "Employees", "EmployeeID", "FixedAssets", "EmployeeID", 0&)
    c.Add Array("FK_AssetDepreciations_RunID", "DepreciationRuns", "RunID", "AssetDepreciations", "RunID", 4096&)
    c.Add Array("FK_AssetDepreciations_AssetID", "FixedAssets", "AssetID", "AssetDepreciations", "AssetID", 0&)
    c.Add Array("FK_BudgetLines_BudgetID", "Budgets", "BudgetID", "BudgetLines", "BudgetID", 4096&)
    c.Add Array("FK_BudgetLines_AccountCode", "Accounts", "AccountCode", "BudgetLines", "AccountCode", 0&)
    c.Add Array("FK_BudgetLines_CostCenterID", "CostCenters", "CostCenterID", "BudgetLines", "CostCenterID", 0&)
    c.Add Array("FK_PayrollRuns_BankID", "Banks", "BankID", "PayrollRuns", "BankID", 0&)
    c.Add Array("FK_PayrollRuns_CashBoxID", "CashBoxes", "CashBoxID", "PayrollRuns", "CashBoxID", 0&)
    c.Add Array("FK_PayrollRuns_EmployeeID", "Employees", "EmployeeID", "PayrollRuns", "EmployeeID", 0&)
    c.Add Array("FK_PayrollLines_PayrollRunID", "PayrollRuns", "PayrollRunID", "PayrollLines", "PayrollRunID", 4096&)
    c.Add Array("FK_PayrollLines_EmployeeID", "Employees", "EmployeeID", "PayrollLines", "EmployeeID", 0&)
    c.Add Array("FK_PayrollLines_CostCenterID", "CostCenters", "CostCenterID", "PayrollLines", "CostCenterID", 0&)
    c.Add Array("FK_BankReconciliations_BankID", "Banks", "BankID", "BankReconciliations", "BankID", 0&)
    c.Add Array("FK_BankClearings_ReconciliationID", "BankReconciliations", "ReconciliationID", "BankClearings", "ReconciliationID", 4096&)
    c.Add Array("FK_BankClearings_BankID", "Banks", "BankID", "BankClearings", "BankID", 0&)
    c.Add Array("FK_CustomerAllocations_PaymentID", "CustomerPayments", "PaymentID", "CustomerAllocations", "PaymentID", 4096&)
    c.Add Array("FK_CustomerAllocations_SalesInvoiceID", "SalesInvoices", "SalesInvoiceID", "CustomerAllocations", "SalesInvoiceID", 0&)
    c.Add Array("FK_SupplierAllocations_PaymentID", "SupplierPayments", "PaymentID", "SupplierAllocations", "PaymentID", 4096&)
    c.Add Array("FK_SupplierAllocations_PurchaseInvoiceID", "PurchaseInvoices", "PurchaseInvoiceID", "SupplierAllocations", "PurchaseInvoiceID", 0&)
    c.Add Array("FK_Expenses_ExpenseTypeID", "ExpenseTypes", "ExpenseTypeID", "Expenses", "ExpenseTypeID", 0&)
    c.Add Array("FK_Expenses_PaymentMethodID", "PaymentMethods", "PaymentMethodID", "Expenses", "PaymentMethodID", 0&)
    c.Add Array("FK_Expenses_EmployeeID", "Employees", "EmployeeID", "Expenses", "EmployeeID", 0&)
    c.Add Array("FK_Expenses_CashBoxID", "CashBoxes", "CashBoxID", "Expenses", "CashBoxID", 0&)
    c.Add Array("FK_Expenses_BankID", "Banks", "BankID", "Expenses", "BankID", 0&)
    c.Add Array("FK_Expenses_CostCenterID", "CostCenters", "CostCenterID", "Expenses", "CostCenterID", 0&)
    c.Add Array("FK_CashVouchers_CashBoxID", "CashBoxes", "CashBoxID", "CashVouchers", "CashBoxID", 0&)
    c.Add Array("FK_CashVouchers_ToCashBoxID", "CashBoxes", "CashBoxID", "CashVouchers", "ToCashBoxID", 0&)
    c.Add Array("FK_CashVouchers_ExpenseID", "Expenses", "ExpenseID", "CashVouchers", "ExpenseID", 0&)
    c.Add Array("FK_CashVouchers_ClosingID", "CashClosings", "ClosingID", "CashVouchers", "ClosingID", 0&)
    c.Add Array("FK_CashVouchers_AdvanceEmployeeID", "Employees", "EmployeeID", "CashVouchers", "AdvanceEmployeeID", 0&)
    c.Add Array("FK_CashVouchers_CostCenterID", "CostCenters", "CostCenterID", "CashVouchers", "CostCenterID", 0&)
    c.Add Array("FK_CashVouchers_EmployeeID", "Employees", "EmployeeID", "CashVouchers", "EmployeeID", 0&)
    c.Add Array("FK_CashClosings_CashBoxID", "CashBoxes", "CashBoxID", "CashClosings", "CashBoxID", 0&)
    c.Add Array("FK_CashClosings_EmployeeID", "Employees", "EmployeeID", "CashClosings", "EmployeeID", 0&)
    c.Add Array("FK_CashClosings_ToCashBoxID", "CashBoxes", "CashBoxID", "CashClosings", "ToCashBoxID", 0&)
    c.Add Array("FK_JournalEntries_SourceType", "JournalSourceTypes", "SourceType", "JournalEntries", "SourceType", 256&)
    c.Add Array("FK_JournalLines_EntryID", "JournalEntries", "EntryID", "JournalLines", "EntryID", 4096&)
    c.Add Array("FK_JournalLines_AccountCode", "Accounts", "AccountCode", "JournalLines", "AccountCode", 0&)
    c.Add Array("FK_JournalLines_CostCenterID", "CostCenters", "CostCenterID", "JournalLines", "CostCenterID", 0&)
    c.Add Array("FK_FiscalYearClosingLines_YearClosingID", "FiscalYearClosings", "YearClosingID", "FiscalYearClosingLines", "YearClosingID", 4096&)
    c.Add Array("FK_FiscalYearClosingLines_AccountCode", "Accounts", "AccountCode", "FiscalYearClosingLines", "AccountCode", 0&)
    c.Add Array("FK_VatReturns_PaidAccount", "Accounts", "AccountCode", "VatReturns", "PaidAccount", 0&)
    c.Add Array("FK_ManualEntries_EmployeeID", "Employees", "EmployeeID", "ManualEntries", "EmployeeID", 0&)
    c.Add Array("FK_ManualEntryLines_ManualEntryID", "ManualEntries", "ManualEntryID", "ManualEntryLines", "ManualEntryID", 4096&)
    c.Add Array("FK_ManualEntryLines_AccountCode", "Accounts", "AccountCode", "ManualEntryLines", "AccountCode", 0&)
    c.Add Array("FK_ManualEntryLines_CostCenterID", "CostCenters", "CostCenterID", "ManualEntryLines", "CostCenterID", 0&)
    c.Add Array("FK_InventoryTransactions_ProductID", "Products", "ProductID", "InventoryTransactions", "ProductID", 0&)
    c.Add Array("FK_InventoryTransactions_TransactionTypeID", "TransactionTypes", "TransactionTypeID", "InventoryTransactions", "TransactionTypeID", 0&)
    c.Add Array("FK_InventoryTransactions_EmployeeID", "Employees", "EmployeeID", "InventoryTransactions", "EmployeeID", 0&)
    c.Add Array("FK_StockCounts_CategoryID", "Categories", "CategoryID", "StockCounts", "CategoryID", 0&)
    c.Add Array("FK_StockCounts_EmployeeID", "Employees", "EmployeeID", "StockCounts", "EmployeeID", 0&)
    c.Add Array("FK_StockCountDetails_StockCountID", "StockCounts", "StockCountID", "StockCountDetails", "StockCountID", 4096&)
    c.Add Array("FK_StockCountDetails_ProductID", "Products", "ProductID", "StockCountDetails", "ProductID", 0&)
    Set RelationSpecs = c
End Function

Private Function RetiredRelations() As Variant
    RetiredRelations = Array("FK_AuditLog_EmployeeID", _
                             "FK_BankReconciliations_EmployeeID", _
                             "FK_Budgets_EmployeeID", _
                             "FK_CustomerAllocations_EmployeeID", _
                             "FK_DepreciationRuns_EmployeeID", _
                             "FK_FiscalYearClosings_EmployeeID", _
                             "FK_PeriodClosings_EmployeeID", _
                             "FK_StockCounts_PostedByID", _
                             "FK_SupplierAllocations_EmployeeID", _
                             "FK_VatReturns_EmployeeID")
End Function
