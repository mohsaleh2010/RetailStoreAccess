Attribute VB_Name = "modBuildSchema"
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

Private Const SCHEMA_TABLES As String = "Settings,Sequences,Roles,Permissions,RolePermissions,Employees,Categories,Units,PaymentMethods,CashBoxes,Suppliers,Customers,Products,SalesInvoices,SalesInvoiceDetails,SalesReturns,SalesReturnDetails,PurchaseInvoices,PurchaseInvoiceDetails,PurchaseReturns,PurchaseReturnDetails,CustomerPayments,SupplierPayments,ExpenseTypes,Expenses,CashVouchers,CashClosings,Accounts,JournalSourceTypes,JournalEntries,JournalLines,TransactionTypes,InventoryTransactions,StockCounts,StockCountDetails,AuditLog,LabelSettings"
Private Const EXPECTED_FIELD_COUNTS As String = "Settings=31;Sequences=5;Roles=4;Permissions=4;RolePermissions=2;Employees=17;Categories=8;Units=4;PaymentMethods=5;CashBoxes=8;Suppliers=15;Customers=21;Products=23;SalesInvoices=35;SalesInvoiceDetails=14;SalesReturns=29;SalesReturnDetails=14;PurchaseInvoices=18;PurchaseInvoiceDetails=11;PurchaseReturns=18;PurchaseReturnDetails=11;CustomerPayments=11;SupplierPayments=11;ExpenseTypes=3;Expenses=13;CashVouchers=14;CashClosings=18;Accounts=5;JournalSourceTypes=3;JournalEntries=13;JournalLines=7;TransactionTypes=5;InventoryTransactions=13;StockCounts=9;StockCountDetails=9;AuditLog=8;LabelSettings=19"
Private Const EXPECTED_SEED_COUNTS As String = "Settings=1;Sequences=16;Roles=3;Permissions=25;RolePermissions=51;Employees=1;Categories=1;Units=8;PaymentMethods=4;CashBoxes=2;Customers=1;ExpenseTypes=9;Accounts=22;JournalSourceTypes=13;TransactionTypes=8;LabelSettings=1"

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

    m_db.Close
    Set m_db = Nothing

    m_currentStep = "link tables"
    LinkBackEnd BackEndPath

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
        If FieldExistsIn(tdf, FieldName) Then Exit Sub
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

Private Sub LogLine(ByVal Msg As String)
    m_log = m_log & Msg & vbCrLf
    Debug.Print Msg
End Sub

'------------------------------------------------------------------------------
' Generated: table definitions
'------------------------------------------------------------------------------

Private Sub CreateAllTables()
    CreateTable_Settings
    CreateTable_Sequences
    CreateTable_Roles
    CreateTable_Permissions
    CreateTable_RolePermissions
    CreateTable_Employees
    CreateTable_Categories
    CreateTable_Units
    CreateTable_PaymentMethods
    CreateTable_CashBoxes
    CreateTable_Suppliers
    CreateTable_Customers
    CreateTable_Products
    CreateTable_SalesInvoices
    CreateTable_SalesInvoiceDetails
    CreateTable_SalesReturns
    CreateTable_SalesReturnDetails
    CreateTable_PurchaseInvoices
    CreateTable_PurchaseInvoiceDetails
    CreateTable_PurchaseReturns
    CreateTable_PurchaseReturnDetails
    CreateTable_CustomerPayments
    CreateTable_SupplierPayments
    CreateTable_ExpenseTypes
    CreateTable_Expenses
    CreateTable_CashVouchers
    CreateTable_CashClosings
    CreateTable_Accounts
    CreateTable_JournalSourceTypes
    CreateTable_JournalEntries
    CreateTable_JournalLines
    CreateTable_TransactionTypes
    CreateTable_InventoryTransactions
    CreateTable_StockCounts
    CreateTable_StockCountDetails
    CreateTable_AuditLog
    CreateTable_LabelSettings
End Sub

Private Sub CreateTable_Settings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Settings") Then Exit Sub
    AddField tdf, "SettingID", "LONG", 0, True, "1", _
             "=1", "يسمح بسجل إعدادات واحد فقط", "رقم الإعداد", ""
    AddField tdf, "StoreName", "TEXT", 150, True, "", _
             "", "", "اسم المحل", ""
    AddField tdf, "StoreNameEn", "TEXT", 150, False, "", _
             "", "", "اسم المحل بالإنجليزية", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "الرقم الضريبي 15 رقمًا ويبدأ وينتهي بالرقم 3", "الرقم الضريبي", ""
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "السجل التجاري", ""
    AddField tdf, "BuildingNo", "TEXT", 10, False, "", _
             "", "", "رقم المبنى", ""
    AddField tdf, "StreetName", "TEXT", 100, False, "", _
             "", "", "الشارع", ""
    AddField tdf, "District", "TEXT", 100, False, "", _
             "", "", "الحي", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "المدينة", ""
    AddField tdf, "PostalCode", "TEXT", 10, False, "", _
             "", "", "الرمز البريدي", ""
    AddField tdf, "AdditionalNo", "TEXT", 10, False, "", _
             "", "", "الرقم الإضافي", ""
    AddField tdf, "CountryCode", "TEXT", 2, True, """SA""", _
             "", "", "رمز الدولة", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "الهاتف", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "البريد الإلكتروني", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "النسبة يجب أن تكون بين 0% و 100%", "نسبة الضريبة", ""
    AddField tdf, "PricesIncludeVAT", "BOOL", 0, False, "True", _
             "", "", "الأسعار شاملة الضريبة", ""
    AddField tdf, "AllowNegativeStock", "BOOL", 0, False, "False", _
             "", "", "السماح بالبيع بالسالب", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, True, """SAR""", _
             "", "", "العملة", ""
    AddField tdf, "DefaultCustomerID", "LONG", 0, True, "1", _
             "", "", "العميل الافتراضي", ""
    AddField tdf, "ZatcaPhase", "BYTE", 0, True, "1", _
             "In (1,2)", "المرحلة 1 أو 2", "مرحلة فاتورة", ""
    AddField tdf, "ZatcaEnvironment", "TEXT", 20, False, "", _
             "", "", "بيئة الربط مع الهيئة", ""
    AddField tdf, "LastInvoiceHash", "TEXT", 255, False, "", _
             "", "", "بصمة آخر مستند", "تبدأ بالقيمة الافتراضية التي تحددها الهيئة لأول فاتورة"
    AddField tdf, "BackupFolder", "TEXT", 255, False, "", _
             "", "", "مجلد النسخ الاحتياطي", ""
    AddField tdf, "BackupKeepCount", "INT", 0, True, "30", _
             ">=1", "احتفظ بنسخة واحدة على الأقل", "عدد النسخ المحتفظ بها", ""
    AddField tdf, "SlowMovingDays", "INT", 0, True, "90", _
             ">=1", "عدد الأيام يجب أن يكون 1 أو أكثر", "أيام عدم الحركة", ""
    AddField tdf, "ReceiptFooter", "TEXT", 255, False, "", _
             "", "", "تذييل الفاتورة", ""
    AddField tdf, "LogoPath", "TEXT", 255, False, "", _
             "", "", "مسار الشعار", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "آخر تعديل", ""
    AddField tdf, "POSMode", "TEXT", 10, True, """RETAIL""", _
             "In (""RETAIL"",""RESTAURANT"",""CAFE"")", "اختر شاشة البيع من القائمة", "شاشة البيع", "RETAIL = المحلات (باركود)، RESTAURANT = المطاعم (لمس)، CAFE = الكافيهات (لمس)"
    AddField tdf, "ImagesFolder", "TEXT", 255, False, "", _
             "", "", "مجلد صور المنتجات", "المسارات النسبية للصور تُقرأ منه؛ فارغ = مجلد Images بجانب ملف البيانات"
    AddField tdf, "InvoicePrintMode", "TEXT", 10, True, """PREVIEW""", _
             "In (""DIRECT"",""PREVIEW"",""NONE"")", "اختر طريقة الطباعة من القائمة", "الطباعة عند حفظ الفاتورة", "DIRECT = طباعة مباشرة بدون معاينة، PREVIEW = عرض المعاينة، NONE = بدون طباعة"
    AddIndex tdf, "PrimaryKey", "SettingID", True, True, False
    EndTable tdf, "إعدادات المحل: سجل واحد فقط يحتوي بيانات المحل الضريبية وإعدادات التشغيل.", "", ""
End Sub

Private Sub CreateTable_Sequences()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Sequences") Then Exit Sub
    AddField tdf, "SequenceName", "TEXT", 30, True, "", _
             "", "", "اسم العدّاد", ""
    AddField tdf, "Prefix", "TEXT", 10, False, "", _
             "", "", "البادئة", ""
    AddField tdf, "NextValue", "LONG", 0, True, "1", _
             ">=1", "الرقم التالي يجب أن يكون 1 أو أكثر", "الرقم التالي", ""
    AddField tdf, "PadLength", "BYTE", 0, True, "6", _
             "Between 0 And 12", "من 0 إلى 12 خانة", "عدد الخانات", ""
    AddField tdf, "Description", "TEXT", 100, False, "", _
             "", "", "الوصف", ""
    AddIndex tdf, "PrimaryKey", "SequenceName", True, True, False
    EndTable tdf, "عدّادات الترقيم: يولّد أرقامًا تسلسلية بدون فجوات للفواتير والسندات (يُزاد داخل نفس معاملة الحفظ).", "", ""
End Sub

Private Sub CreateTable_Roles()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Roles") Then Exit Sub
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "رقم الدور", ""
    AddField tdf, "RoleCode", "TEXT", 20, True, "", _
             "", "", "رمز الدور", ""
    AddField tdf, "RoleName", "TEXT", 50, True, "", _
             "", "", "اسم الدور", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "الوصف", ""
    AddIndex tdf, "PrimaryKey", "RoleID", True, True, False
    AddIndex tdf, "UX_RoleCode", "RoleCode", False, True, False
    EndTable tdf, "الأدوار: مستويات الصلاحيات: مدير النظام، مدير، كاشير.", "", ""
End Sub

Private Sub CreateTable_Permissions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Permissions") Then Exit Sub
    AddField tdf, "PermissionKey", "TEXT", 50, True, "", _
             "", "", "رمز الصلاحية", ""
    AddField tdf, "PermissionName", "TEXT", 100, True, "", _
             "", "", "اسم الصلاحية", ""
    AddField tdf, "ModuleName", "TEXT", 50, False, "", _
             "", "", "القسم", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "الترتيب", ""
    AddIndex tdf, "PrimaryKey", "PermissionKey", True, True, False
    EndTable tdf, "الصلاحيات: قائمة الصلاحيات التي يمكن منحها للأدوار.", "", ""
End Sub

Private Sub CreateTable_RolePermissions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "RolePermissions") Then Exit Sub
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "الدور", ""
    AddField tdf, "PermissionKey", "TEXT", 50, True, "", _
             "", "", "الصلاحية", ""
    AddIndex tdf, "PrimaryKey", "RoleID,PermissionKey", True, True, False
    EndTable tdf, "صلاحيات الأدوار: ربط كل دور بالصلاحيات الممنوحة له (علاقة متعدد لمتعدد).", "", ""
End Sub

Private Sub CreateTable_Employees()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Employees") Then Exit Sub
    AddField tdf, "EmployeeID", "AUTO", 0, False, "", _
             "", "", "رقم الموظف", ""
    AddField tdf, "EmployeeName", "TEXT", 100, True, "", _
             "", "", "اسم الموظف", ""
    AddField tdf, "JobTitle", "TEXT", 50, False, "", _
             "", "", "المسمى الوظيفي", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "الجوال", ""
    AddField tdf, "Username", "TEXT", 30, True, "", _
             "", "", "اسم المستخدم", ""
    AddField tdf, "PasswordHash", "TEXT", 64, False, "", _
             "", "", "بصمة كلمة المرور", "SHA-256 بصيغة hex؛ كلمة المرور نفسها لا تُحفظ أبدًا"
    AddField tdf, "PasswordSalt", "TEXT", 32, False, "", _
             "", "", "ملح التشفير", ""
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "الدور", ""
    AddField tdf, "MaxDiscountPercent", "RATE", 0, True, "0", _
             ">=0 And <=1", "النسبة بين 0% و 100%", "أقصى نسبة خصم", ""
    AddField tdf, "MustChangePassword", "BOOL", 0, False, "True", _
             "", "", "يجب تغيير كلمة المرور", ""
    AddField tdf, "FailedLoginCount", "INT", 0, True, "0", _
             "", "", "محاولات الدخول الفاشلة", ""
    AddField tdf, "LockedUntil", "DATETIME", 0, False, "", _
             "", "", "مقفل حتى", ""
    AddField tdf, "LastLoginAt", "DATETIME", 0, False, "", _
             "", "", "آخر دخول", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "تدخل فيه نقدية مبيعاته وسنداته؛ فارغ = أول صندوق كاشير نشط"
    AddIndex tdf, "PrimaryKey", "EmployeeID", True, True, False
    AddIndex tdf, "UX_Username", "Username", False, True, False
    EndTable tdf, "الموظفون والمستخدمون: كل موظف هو مستخدم للنظام؛ لا يُحذف بل يُعطَّل للحفاظ على سجل عملياته.", "", ""
End Sub

Private Sub CreateTable_Categories()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Categories") Then Exit Sub
    AddField tdf, "CategoryID", "AUTO", 0, False, "", _
             "", "", "رقم التصنيف", ""
    AddField tdf, "CategoryName", "TEXT", 100, True, "", _
             "", "", "اسم التصنيف", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "الوصف", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "ImagePath", "TEXT", 255, False, "", _
             "", "", "صورة التصنيف", ""
    AddField tdf, "TileColor", "TEXT", 10, True, """BLUE""", _
             "In (""BLUE"",""GREEN"",""ORANGE"",""PURPLE"",""RED"",""INDIGO"",""TEAL"",""PINK"",""BROWN"",""GREY"")", "اختر اللون من القائمة", "لون الزر", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "ترتيب العرض", ""
    AddField tdf, "IsAddOn", "BOOL", 0, False, "False", _
             "", "", "فئة إضافات", ""
    AddIndex tdf, "PrimaryKey", "CategoryID", True, True, False
    AddIndex tdf, "UX_CategoryName", "CategoryName", False, True, False
    EndTable tdf, "التصنيفات: تصنيفات المنتجات (إلكترونيات، مواد غذائية، ...).", "", ""
End Sub

Private Sub CreateTable_Units()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Units") Then Exit Sub
    AddField tdf, "UnitID", "AUTO", 0, False, "", _
             "", "", "رقم الوحدة", ""
    AddField tdf, "UnitName", "TEXT", 30, True, "", _
             "", "", "اسم الوحدة", ""
    AddField tdf, "ZatcaUnitCode", "TEXT", 10, False, "", _
             "", "", "رمز الوحدة (UN/ECE)", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddIndex tdf, "PrimaryKey", "UnitID", True, True, False
    AddIndex tdf, "UX_UnitName", "UnitName", False, True, False
    EndTable tdf, "وحدات القياس: وحدات البيع (حبة، كرتون، كيلو، ...) مع رمزها في فاتورة الهيئة.", "", ""
End Sub

Private Sub CreateTable_PaymentMethods()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PaymentMethods") Then Exit Sub
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "", _
             "", "", "رقم الطريقة", ""
    AddField tdf, "MethodName", "TEXT", 50, True, "", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "ZatcaCode", "TEXT", 5, False, "", _
             "", "", "رمز الهيئة", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "الترتيب", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddIndex tdf, "PrimaryKey", "PaymentMethodID", True, True, False
    AddIndex tdf, "UX_MethodName", "MethodName", False, True, False
    EndTable tdf, "طرق الدفع: طرق دفع المبالغ المسددة مع رمزها في فاتورة الهيئة (UNTDID 4461).", "", ""
End Sub

Private Sub CreateTable_CashBoxes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashBoxes") Then Exit Sub
    AddField tdf, "CashBoxID", "AUTO", 0, False, "", _
             "", "", "رقم الصندوق", ""
    AddField tdf, "BoxName", "TEXT", 50, True, "", _
             "", "", "اسم الصندوق", ""
    AddField tdf, "BoxType", "TEXT", 10, True, """CASHIER""", _
             "In (""MAIN"",""CASHIER"")", "MAIN = خزينة رئيسية، CASHIER = صندوق كاشير", "النوع", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الرصيد الافتتاحي", ""
    AddField tdf, "OpeningDate", "DATE", 0, True, "Date()", _
             "", "", "تاريخ الرصيد الافتتاحي", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "CashBoxID", True, True, False
    AddIndex tdf, "UX_BoxName", "BoxName", False, True, False
    EndTable tdf, "الخزينة والصناديق: الخزينة الرئيسية وصناديق الكاشير. الرصيد لا يُخزَّن: يُحسب من الحركات (qryCashMovements).", "", ""
End Sub

Private Sub CreateTable_Suppliers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Suppliers") Then Exit Sub
    AddField tdf, "SupplierID", "AUTO", 0, False, "", _
             "", "", "رقم المورد", ""
    AddField tdf, "SupplierName", "TEXT", 150, True, "", _
             "", "", "اسم المورد", ""
    AddField tdf, "ContactPerson", "TEXT", 100, False, "", _
             "", "", "الشخص المسؤول", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "الجوال", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "الهاتف", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "البريد الإلكتروني", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "الرقم الضريبي 15 رقمًا ويبدأ وينتهي بالرقم 3", "الرقم الضريبي", ""
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "السجل التجاري", ""
    AddField tdf, "Address", "TEXT", 255, False, "", _
             "", "", "العنوان", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "المدينة", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "الرصيد الافتتاحي", "موجب = المحل مدين للمورد"
    AddField tdf, "CurrentBalance", "MONEY", 0, True, "0", _
             "", "", "الرصيد الحالي", "قيمة مساعدة؛ المرجع هو SupplierBalanceQuery"
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "SupplierID", True, True, False
    AddIndex tdf, "IX_SupplierName", "SupplierName", False, False, False
    AddIndex tdf, "IX_Mobile", "Mobile", False, False, False
    EndTable tdf, "الموردون: بيانات الموردين. الرصيد الحالي قيمة مساعدة تُحدَّث بالكود ويمكن إعادة احتسابها من الحركات.", "", ""
End Sub

Private Sub CreateTable_Customers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Customers") Then Exit Sub
    AddField tdf, "CustomerID", "AUTO", 0, False, "", _
             "", "", "رقم العميل", ""
    AddField tdf, "CustomerName", "TEXT", 150, True, "", _
             "", "", "اسم العميل", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "الجوال", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "الهاتف", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "البريد الإلكتروني", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "الرقم الضريبي 15 رقمًا ويبدأ وينتهي بالرقم 3", "الرقم الضريبي", "إذا وُجد تصدر للعميل فاتورة ضريبية B2B"
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "السجل التجاري", ""
    AddField tdf, "BuildingNo", "TEXT", 10, False, "", _
             "", "", "رقم المبنى", ""
    AddField tdf, "StreetName", "TEXT", 100, False, "", _
             "", "", "الشارع", ""
    AddField tdf, "District", "TEXT", 100, False, "", _
             "", "", "الحي", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "المدينة", ""
    AddField tdf, "PostalCode", "TEXT", 10, False, "", _
             "", "", "الرمز البريدي", ""
    AddField tdf, "Address", "TEXT", 255, False, "", _
             "", "", "العنوان", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "الرصيد الافتتاحي", "موجب = العميل مدين للمحل"
    AddField tdf, "CurrentBalance", "MONEY", 0, True, "0", _
             "", "", "الرصيد الحالي", "قيمة مساعدة؛ المرجع هو CustomerBalanceQuery"
    AddField tdf, "AllowCredit", "BOOL", 0, False, "True", _
             "", "", "يسمح بالبيع الآجل", ""
    AddField tdf, "CreditLimit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "حد الائتمان", "0 = بدون حد"
    AddField tdf, "IsSystem", "BOOL", 0, False, "False", _
             "", "", "سجل نظام", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "CustomerID", True, True, False
    AddIndex tdf, "IX_CustomerName", "CustomerName", False, False, False
    AddIndex tdf, "IX_Mobile", "Mobile", False, False, False
    EndTable tdf, "العملاء: بيانات العملاء. السجل رقم 1 هو ""عميل نقدي"" الافتراضي ولا يُحذف.", "", ""
End Sub

Private Sub CreateTable_Products()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Products") Then Exit Sub
    AddField tdf, "ProductID", "AUTO", 0, False, "", _
             "", "", "رقم المنتج", ""
    AddField tdf, "ProductCode", "TEXT", 30, True, "", _
             "", "", "كود المنتج", ""
    AddField tdf, "Barcode", "TEXT", 50, False, "", _
             "", "", "الباركود", ""
    AddField tdf, "ProductName", "TEXT", 150, True, "", _
             "", "", "اسم المنتج", ""
    AddField tdf, "ProductNameEn", "TEXT", 150, False, "", _
             "", "", "الاسم بالإنجليزية", ""
    AddField tdf, "CategoryID", "LONG", 0, True, "1", _
             "", "", "التصنيف", ""
    AddField tdf, "UnitID", "LONG", 0, True, "1", _
             "", "", "الوحدة", ""
    AddField tdf, "PurchasePrice", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "آخر سعر شراء", "بدون ضريبة"
    AddField tdf, "AverageCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "متوسط التكلفة", "المتوسط المرجّح، يُحدَّث مع كل شراء"
    AddField tdf, "SellingPrice", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "سعر البيع", "شامل الضريبة إذا كان الإعداد PricesIncludeVAT مفعّلًا"
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = خاضع 15%، Z = نسبة صفرية، E = معفى", "الفئة الضريبية", ""
    AddField tdf, "CurrentQuantity", "QTY", 0, True, "0", _
             "", "", "الكمية الحالية", ""
    AddField tdf, "MinimumQuantity", "QTY", 0, True, "0", _
             ">=0", "الحد الأدنى لا يمكن أن يكون سالبًا", "حد إعادة الطلب", ""
    AddField tdf, "SupplierID", "LONG", 0, False, "", _
             "", "", "المورد الافتراضي", ""
    AddField tdf, "ProductLocation", "TEXT", 50, False, "", _
             "", "", "مكان المنتج", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "آخر تعديل", ""
    AddField tdf, "ImagePath", "TEXT", 255, False, "", _
             "", "", "صورة المنتج", "لشاشات اللمس؛ مسار كامل أو اسم ملف في مجلد الصور"
    AddField tdf, "TrackStock", "BOOL", 0, False, "True", _
             "", "", "يتابع المخزون", ""
    AddField tdf, "SizePriceM", "MONEY", 0, False, "", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "سعر الحجم الوسط", ""
    AddField tdf, "SizePriceL", "MONEY", 0, False, "", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "سعر الحجم الكبير", ""
    AddIndex tdf, "PrimaryKey", "ProductID", True, True, False
    AddIndex tdf, "UX_ProductCode", "ProductCode", False, True, False
    AddIndex tdf, "UX_Barcode", "Barcode", False, True, True
    AddIndex tdf, "IX_ProductName", "ProductName", False, False, False
    EndTable tdf, "المنتجات: الأصناف وأسعارها. CurrentQuantity قيمة مساعدة؛ المرجع هو مجموع InventoryTransactions.", "", ""
End Sub

Private Sub CreateTable_SalesInvoices()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesInvoices") Then Exit Sub
    AddField tdf, "SalesInvoiceID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "InvoiceNumber", "TEXT", 20, True, "", _
             "", "", "رقم الفاتورة", ""
    AddField tdf, "InvoiceDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ ووقت الفاتورة", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "1", _
             "", "", "العميل", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الكاشير", ""
    AddField tdf, "PaymentType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = نقدي، CREDIT = آجل", "نوع البيع", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المجموع قبل الخصم", "مجموع (الكمية × سعر الوحدة) بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", "مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخاضع للضريبة", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة القيمة المضافة", "مجموع ضريبة الأسطر"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", "TaxableAmount + Tax"
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المدفوع", "المبلغ المحتسب من الفاتورة (لا يتجاوز الإجمالي)"
    AddField tdf, "RemainingAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المتبقي", "يُضاف إلى رصيد العميل في البيع الآجل"
    AddField tdf, "AmountTendered", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المبلغ المستلم", "ما سلّمه العميل نقدًا"
    AddField tdf, "ChangeDue", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الباقي للعميل", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "InvoiceSubType", "TEXT", 10, True, """SIMPLIFIED""", _
             "In (""SIMPLIFIED"",""STANDARD"")", "SIMPLIFIED = فاتورة مبسطة، STANDARD = فاتورة ضريبية", "نوع الفاتورة الضريبية", "مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي)"
    AddField tdf, "InvoiceTypeCode", "TEXT", 3, True, """388""", _
             "In (""388"",""381"",""383"")", "388 فاتورة، 381 إشعار دائن، 383 إشعار مدين", "رمز نوع المستند", ""
    AddField tdf, "InvoiceUUID", "TEXT", 36, False, "", _
             "", "", "المعرّف الفريد UUID", ""
    AddField tdf, "ICV", "LONG", 0, False, "", _
             "", "", "عدّاد الفواتير ICV", "تسلسل مشترك لكل المستندات المرسلة للهيئة"
    AddField tdf, "InvoiceHash", "TEXT", 255, False, "", _
             "", "", "بصمة المستند", ""
    AddField tdf, "PreviousInvoiceHash", "TEXT", 255, False, "", _
             "", "", "بصمة المستند السابق", ""
    AddField tdf, "QRCodeData", "MEMO", 0, False, "", _
             "", "", "بيانات رمز QR", ""
    AddField tdf, "ZatcaStatus", "TEXT", 20, True, """NOT_SENT""", _
             "In (""NOT_SENT"",""PENDING"",""REPORTED"",""CLEARED"",""WARNING"",""REJECTED"")", "حالة غير معروفة", "حالة الإرسال للهيئة", ""
    AddField tdf, "ZatcaSubmittedAt", "DATETIME", 0, False, "", _
             "", "", "تاريخ الإرسال للهيئة", ""
    AddField tdf, "ZatcaResponse", "MEMO", 0, False, "", _
             "", "", "رد الهيئة", ""
    AddField tdf, "SignedXmlPath", "TEXT", 255, False, "", _
             "", "", "مسار ملف XML الموقّع", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "OrderType", "TEXT", 10, False, "", _
             "Is Null Or In (""DINE_IN"",""TAKEAWAY"",""DELIVERY"")", "نوع الطلب: داخلي أو سفري أو توصيل", "نوع الطلب", ""
    AddField tdf, "TableNo", "TEXT", 10, False, "", _
             "", "", "رقم الطاولة", ""
    AddField tdf, "DeliveryPhone", "TEXT", 20, False, "", _
             "", "", "جوال التوصيل", ""
    AddField tdf, "DeliveryAddress", "TEXT", 255, False, "", _
             "", "", "عنوان التوصيل", ""
    AddField tdf, "OrderName", "TEXT", 50, False, "", _
             "", "", "اسم العميل على الطلب", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "يُملأ عند الدفع النقدي: المبلغ المدفوع يدخل هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "SalesInvoiceID", True, True, False
    AddIndex tdf, "UX_InvoiceNumber", "InvoiceNumber", False, True, False
    AddIndex tdf, "IX_InvoiceDate", "InvoiceDate", False, False, False
    AddIndex tdf, "UX_InvoiceUUID", "InvoiceUUID", False, True, True
    AddIndex tdf, "IX_ZatcaStatus", "ZatcaStatus", False, False, False
    EndTable tdf, "فواتير المبيعات: رأس فاتورة البيع. لا تُحذف ولا تُعدَّل بعد الحفظ؛ التصحيح بمرتجع.", "[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]", "المدفوع لا يتجاوز الإجمالي، والمتبقي = الإجمالي - المدفوع"
End Sub

Private Sub CreateTable_SalesInvoiceDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesInvoiceDetails") Then Exit Sub
    AddField tdf, "SalesDetailID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, True, "", _
             "", "", "الفاتورة", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "رقم السطر", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "الكمية يجب أن تكون أكبر من صفر", "الكمية", ""
    AddField tdf, "UnitPrice", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "سعر الوحدة", "بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", "قيمة الخصم على السطر بدون ضريبة"
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الصافي قبل الضريبة", "Quantity × UnitPrice - Discount"
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = خاضع 15%، Z = نسبة صفرية، E = معفى", "الفئة الضريبية", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "النسبة يجب أن تكون بين 0% و 100%", "نسبة الضريبة", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الضريبة", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", "NetAmount + Tax"
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", "AverageCost لحظة البيع"
    AddField tdf, "LineNote", "TEXT", 100, False, "", _
             "", "", "ملاحظة السطر", "الحجم والخيارات (الكافيه)"
    AddIndex tdf, "PrimaryKey", "SalesDetailID", True, True, False
    AddIndex tdf, "UX_SalesInvoiceID_LineNumber", "SalesInvoiceID,LineNumber", False, True, False
    EndTable tdf, "تفاصيل فواتير المبيعات: أسطر فاتورة البيع، مع حفظ تكلفة الصنف لحظة البيع لحساب الربح بدقة.", "", ""
End Sub

Private Sub CreateTable_SalesReturns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesReturns") Then Exit Sub
    AddField tdf, "SalesReturnID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "ReturnNumber", "TEXT", 20, True, "", _
             "", "", "رقم المرتجع", ""
    AddField tdf, "ReturnDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ المرتجع", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, True, "", _
             "", "", "الفاتورة الأصلية", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "", _
             "", "", "العميل", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "Reason", "TEXT", 255, True, "", _
             "", "", "سبب الإرجاع", "إلزامي في الإشعار الدائن حسب متطلبات الهيئة"
    AddField tdf, "RefundType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = رد نقدي، CREDIT = خصم من رصيد العميل", "طريقة رد المبلغ", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "طريقة الرد", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المجموع قبل الخصم", "مجموع (الكمية × سعر الوحدة) بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", "مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخاضع للضريبة", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة القيمة المضافة", "مجموع ضريبة الأسطر"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", "TaxableAmount + Tax"
    AddField tdf, "RefundedAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المبلغ المردود نقدًا", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "InvoiceSubType", "TEXT", 10, True, """SIMPLIFIED""", _
             "In (""SIMPLIFIED"",""STANDARD"")", "SIMPLIFIED = فاتورة مبسطة، STANDARD = فاتورة ضريبية", "نوع الفاتورة الضريبية", "مبسطة للأفراد B2C، ضريبية للمنشآت B2B (للعميل رقم ضريبي)"
    AddField tdf, "InvoiceTypeCode", "TEXT", 3, True, """381""", _
             "In (""388"",""381"",""383"")", "388 فاتورة، 381 إشعار دائن، 383 إشعار مدين", "رمز نوع المستند", ""
    AddField tdf, "InvoiceUUID", "TEXT", 36, False, "", _
             "", "", "المعرّف الفريد UUID", ""
    AddField tdf, "ICV", "LONG", 0, False, "", _
             "", "", "عدّاد الفواتير ICV", "تسلسل مشترك لكل المستندات المرسلة للهيئة"
    AddField tdf, "InvoiceHash", "TEXT", 255, False, "", _
             "", "", "بصمة المستند", ""
    AddField tdf, "PreviousInvoiceHash", "TEXT", 255, False, "", _
             "", "", "بصمة المستند السابق", ""
    AddField tdf, "QRCodeData", "MEMO", 0, False, "", _
             "", "", "بيانات رمز QR", ""
    AddField tdf, "ZatcaStatus", "TEXT", 20, True, """NOT_SENT""", _
             "In (""NOT_SENT"",""PENDING"",""REPORTED"",""CLEARED"",""WARNING"",""REJECTED"")", "حالة غير معروفة", "حالة الإرسال للهيئة", ""
    AddField tdf, "ZatcaSubmittedAt", "DATETIME", 0, False, "", _
             "", "", "تاريخ الإرسال للهيئة", ""
    AddField tdf, "ZatcaResponse", "MEMO", 0, False, "", _
             "", "", "رد الهيئة", ""
    AddField tdf, "SignedXmlPath", "TEXT", 255, False, "", _
             "", "", "مسار ملف XML الموقّع", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "الرد النقدي يخرج من هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "SalesReturnID", True, True, False
    AddIndex tdf, "UX_ReturnNumber", "ReturnNumber", False, True, False
    AddIndex tdf, "IX_ReturnDate", "ReturnDate", False, False, False
    AddIndex tdf, "UX_InvoiceUUID", "InvoiceUUID", False, True, True
    EndTable tdf, "مرتجعات المبيعات: إشعار دائن مرتبط بالفاتورة الأصلية؛ يعيد الكمية للمخزون ويعكس المبلغ.", "[RefundedAmount]<=[TotalAmount]", "المبلغ المردود لا يتجاوز قيمة المرتجع"
End Sub

Private Sub CreateTable_SalesReturnDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesReturnDetails") Then Exit Sub
    AddField tdf, "ReturnDetailID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "SalesReturnID", "LONG", 0, True, "", _
             "", "", "المرتجع", ""
    AddField tdf, "SalesDetailID", "LONG", 0, True, "", _
             "", "", "سطر الفاتورة الأصلي", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "الكمية يجب أن تكون أكبر من صفر", "الكمية المرتجعة", ""
    AddField tdf, "UnitPrice", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "سعر الوحدة", ""
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الصافي قبل الضريبة", ""
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = خاضع 15%، Z = نسبة صفرية، E = معفى", "الفئة الضريبية", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "النسبة يجب أن تكون بين 0% و 100%", "نسبة الضريبة", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الضريبة", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", "نفس تكلفة سطر البيع الأصلي"
    AddField tdf, "ReturnToStock", "BOOL", 0, False, "True", _
             "", "", "يعاد للمخزون", ""
    AddIndex tdf, "PrimaryKey", "ReturnDetailID", True, True, False
    EndTable tdf, "تفاصيل مرتجعات المبيعات: الأسطر المرتجعة، كل سطر يشير إلى سطر الفاتورة الأصلي لمنع إرجاع أكثر مما بيع.", "", ""
End Sub

Private Sub CreateTable_PurchaseInvoices()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseInvoices") Then Exit Sub
    AddField tdf, "PurchaseInvoiceID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "InvoiceNumber", "TEXT", 20, True, "", _
             "", "", "رقم الفاتورة الداخلي", ""
    AddField tdf, "SupplierInvoiceNo", "TEXT", 30, False, "", _
             "", "", "رقم فاتورة المورد", ""
    AddField tdf, "InvoiceDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الفاتورة", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "المورد", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "PaymentType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = نقدي، CREDIT = آجل", "نوع الشراء", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المجموع قبل الخصم", "مجموع (الكمية × سعر الوحدة) بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", "مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخاضع للضريبة", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة القيمة المضافة", "مجموع ضريبة الأسطر"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", "TaxableAmount + Tax"
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المدفوع", ""
    AddField tdf, "RemainingAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المتبقي", "يُضاف إلى رصيد المورد"
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "المدفوع نقدًا يخرج من هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "PurchaseInvoiceID", True, True, False
    AddIndex tdf, "UX_InvoiceNumber", "InvoiceNumber", False, True, False
    AddIndex tdf, "IX_InvoiceDate", "InvoiceDate", False, False, False
    AddIndex tdf, "IX_SupplierInvoiceNo", "SupplierInvoiceNo", False, False, False
    EndTable tdf, "فواتير المشتريات: رأس فاتورة الشراء من المورد.", "[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]", "المدفوع لا يتجاوز الإجمالي، والمتبقي = الإجمالي - المدفوع"
End Sub

Private Sub CreateTable_PurchaseInvoiceDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseInvoiceDetails") Then Exit Sub
    AddField tdf, "PurchaseDetailID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, True, "", _
             "", "", "الفاتورة", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "رقم السطر", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "الكمية يجب أن تكون أكبر من صفر", "الكمية", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", "بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الصافي قبل الضريبة", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "النسبة يجب أن تكون بين 0% و 100%", "نسبة الضريبة", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة المدخلات", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", ""
    AddIndex tdf, "PrimaryKey", "PurchaseDetailID", True, True, False
    AddIndex tdf, "UX_PurchaseInvoiceID_LineNumber", "PurchaseInvoiceID,LineNumber", False, True, False
    EndTable tdf, "تفاصيل فواتير المشتريات: أسطر فاتورة الشراء.", "", ""
End Sub

Private Sub CreateTable_PurchaseReturns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseReturns") Then Exit Sub
    AddField tdf, "PurchaseReturnID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "ReturnNumber", "TEXT", 20, True, "", _
             "", "", "رقم المرتجع", ""
    AddField tdf, "ReturnDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ المرتجع", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, True, "", _
             "", "", "فاتورة الشراء الأصلية", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "المورد", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "Reason", "TEXT", 255, True, "", _
             "", "", "سبب الإرجاع", ""
    AddField tdf, "RefundType", "TEXT", 10, True, """CREDIT""", _
             "In (""CASH"",""CREDIT"")", "CASH = استرداد نقدي، CREDIT = خصم من رصيد المورد", "طريقة الاسترداد", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "طريقة الاسترداد", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المجموع قبل الخصم", "مجموع (الكمية × سعر الوحدة) بدون ضريبة"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", "مجموع خصومات الأسطر (خصم الفاتورة يوزَّع على الأسطر)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخاضع للضريبة", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة القيمة المضافة", "مجموع ضريبة الأسطر"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", "TaxableAmount + Tax"
    AddField tdf, "RefundedAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المبلغ المسترد نقدًا", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "الاسترداد النقدي يدخل هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "PurchaseReturnID", True, True, False
    AddIndex tdf, "UX_ReturnNumber", "ReturnNumber", False, True, False
    AddIndex tdf, "IX_ReturnDate", "ReturnDate", False, False, False
    EndTable tdf, "مرتجعات المشتريات: إرجاع بضاعة للمورد مرتبط بفاتورة الشراء الأصلية.", "[RefundedAmount]<=[TotalAmount]", "المبلغ المسترد لا يتجاوز قيمة المرتجع"
End Sub

Private Sub CreateTable_PurchaseReturnDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseReturnDetails") Then Exit Sub
    AddField tdf, "PurchaseReturnDetailID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "PurchaseReturnID", "LONG", 0, True, "", _
             "", "", "المرتجع", ""
    AddField tdf, "PurchaseDetailID", "LONG", 0, True, "", _
             "", "", "سطر الفاتورة الأصلي", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "الكمية يجب أن تكون أكبر من صفر", "الكمية المرتجعة", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", ""
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الخصم", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الصافي قبل الضريبة", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "النسبة يجب أن تكون بين 0% و 100%", "نسبة الضريبة", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الضريبة", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي شامل الضريبة", ""
    AddIndex tdf, "PrimaryKey", "PurchaseReturnDetailID", True, True, False
    EndTable tdf, "تفاصيل مرتجعات المشتريات: الأسطر المرتجعة للمورد.", "", ""
End Sub

Private Sub CreateTable_CustomerPayments()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CustomerPayments") Then Exit Sub
    AddField tdf, "PaymentID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "PaymentNumber", "TEXT", 20, True, "", _
             "", "", "رقم السند", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "", _
             "", "", "العميل", ""
    AddField tdf, "PaymentDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الدفعة", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "المبلغ يجب أن يكون أكبر من صفر", "المبلغ", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "1", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, False, "", _
             "", "", "عن فاتورة (اختياري)", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "المبلغ النقدي يدخل هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "PaymentID", True, True, False
    AddIndex tdf, "UX_PaymentNumber", "PaymentNumber", False, True, False
    AddIndex tdf, "IX_PaymentDate", "PaymentDate", False, False, False
    EndTable tdf, "دفعات العملاء (سندات القبض): المبالغ المستلمة من العملاء لسداد أرصدتهم الآجلة.", "", ""
End Sub

Private Sub CreateTable_SupplierPayments()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SupplierPayments") Then Exit Sub
    AddField tdf, "PaymentID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "PaymentNumber", "TEXT", 20, True, "", _
             "", "", "رقم السند", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "المورد", ""
    AddField tdf, "PaymentDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الدفعة", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "المبلغ يجب أن يكون أكبر من صفر", "المبلغ", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "1", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, False, "", _
             "", "", "عن فاتورة (اختياري)", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صندوق النقدية", "المبلغ النقدي يخرج من هذا الصندوق"
    AddIndex tdf, "PrimaryKey", "PaymentID", True, True, False
    AddIndex tdf, "UX_PaymentNumber", "PaymentNumber", False, True, False
    AddIndex tdf, "IX_PaymentDate", "PaymentDate", False, False, False
    EndTable tdf, "دفعات الموردين (سندات الصرف): المبالغ المدفوعة للموردين لسداد أرصدتهم.", "", ""
End Sub

Private Sub CreateTable_ExpenseTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ExpenseTypes") Then Exit Sub
    AddField tdf, "ExpenseTypeID", "AUTO", 0, False, "", _
             "", "", "رقم النوع", ""
    AddField tdf, "ExpenseTypeName", "TEXT", 50, True, "", _
             "", "", "نوع المصروف", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddIndex tdf, "PrimaryKey", "ExpenseTypeID", True, True, False
    AddIndex tdf, "UX_ExpenseTypeName", "ExpenseTypeName", False, True, False
    EndTable tdf, "أنواع المصروفات: قائمة ثابتة لأنواع المصروفات لضمان دقة التقارير المجمّعة.", "", ""
End Sub

Private Sub CreateTable_Expenses()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Expenses") Then Exit Sub
    AddField tdf, "ExpenseID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "ExpenseNumber", "TEXT", 20, True, "", _
             "", "", "رقم المصروف", ""
    AddField tdf, "ExpenseDate", "DATE", 0, True, "Date()", _
             "", "", "تاريخ المصروف", ""
    AddField tdf, "ExpenseTypeID", "LONG", 0, True, "", _
             "", "", "نوع المصروف", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "المبلغ يجب أن يكون أكبر من صفر", "المبلغ قبل الضريبة", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "ضريبة المدخلات", "فقط إذا كانت لدى المحل فاتورة ضريبية بالمصروف"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "الإجمالي", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "طريقة الدفع", ""
    AddField tdf, "SupplierInvoiceRef", "TEXT", 30, False, "", _
             "", "", "رقم فاتورة المصروف", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "الوصف", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "صُرف من صندوق", "المصروف النقدي يخرج من هذا الصندوق؛ فارغ = لم يُدفع من صندوق"
    AddIndex tdf, "PrimaryKey", "ExpenseID", True, True, False
    AddIndex tdf, "UX_ExpenseNumber", "ExpenseNumber", False, True, False
    AddIndex tdf, "IX_ExpenseDate", "ExpenseDate", False, False, False
    EndTable tdf, "المصروفات: مصروفات المحل التشغيلية؛ المبلغ بدون ضريبة والضريبة منفصلة (ضريبة مدخلات).", "[TotalAmount]=[Amount]+[Tax]", "الإجمالي = المبلغ + الضريبة"
End Sub

Private Sub CreateTable_CashVouchers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashVouchers") Then Exit Sub
    AddField tdf, "CashVoucherID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "VoucherNumber", "TEXT", 20, True, "", _
             "", "", "رقم السند", ""
    AddField tdf, "VoucherDate", "DATETIME", 0, True, "Now()", _
             "", "", "التاريخ", ""
    AddField tdf, "VoucherType", "TEXT", 10, True, "", _
             "In (""IN"",""OUT"",""TRANSFER"")", "IN = قبض، OUT = صرف، TRANSFER = تحويل بين صندوقين", "نوع السند", ""
    AddField tdf, "CashBoxID", "LONG", 0, True, "", _
             "", "", "الصندوق", "القبض يدخله، والصرف والتحويل يخرجان منه"
    AddField tdf, "ToCashBoxID", "LONG", 0, False, "", _
             "", "", "إلى صندوق", "للتحويل فقط"
    AddField tdf, "Category", "TEXT", 10, True, """OTHER""", _
             "In (""OTHER"",""OWNER"",""EXPENSE"",""ADVANCE"",""SHORTAGE"",""OVERAGE"",""TRANSFER"")", "اختر البند من القائمة", "البند", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "المبلغ يجب أن يكون أكبر من صفر", "المبلغ", ""
    AddField tdf, "PartyName", "TEXT", 100, False, "", _
             "", "", "المستلم / المسلِّم", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "البيان", ""
    AddField tdf, "ExpenseID", "LONG", 0, False, "", _
             "", "", "المصروف المسجَّل", "صرف بند مصروف يسجل مصروفًا بنفس المبلغ في المصروفات"
    AddField tdf, "ClosingID", "LONG", 0, False, "", _
             "", "", "تصفية الكاشير", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "الموظف", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "CashVoucherID", True, True, False
    AddIndex tdf, "UX_VoucherNumber", "VoucherNumber", False, True, False
    AddIndex tdf, "IX_VoucherDate", "VoucherDate", False, False, False
    AddIndex tdf, "IX_CashBoxID", "CashBoxID", False, False, False
    EndTable tdf, "سندات النقدية: قبض نقدية لصندوق، أو صرف منه، أو تحويل بين صندوقين (ومنها ترحيل يومية الكاشير).", "[VoucherType]<>""TRANSFER"" Or ([ToCashBoxID] Is Not Null And [ToCashBoxID]<>[CashBoxID])", "التحويل يحتاج صندوقًا آخر غير صندوق الصرف"
End Sub

Private Sub CreateTable_CashClosings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashClosings") Then Exit Sub
    AddField tdf, "ClosingID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "ClosingNumber", "TEXT", 20, True, "", _
             "", "", "رقم التصفية", ""
    AddField tdf, "ClosingDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ التصفية", ""
    AddField tdf, "CashBoxID", "LONG", 0, True, "", _
             "", "", "الصندوق", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "أجراها", ""
    AddField tdf, "PeriodStart", "DATETIME", 0, False, "", _
             "", "", "من (آخر تصفية)", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "رصيد البداية", ""
    AddField tdf, "CashIn", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المقبوضات", ""
    AddField tdf, "CashOut", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المدفوعات", ""
    AddField tdf, "ExpectedBalance", "MONEY", 0, True, "0", _
             "", "", "الرصيد الدفتري", ""
    AddField tdf, "CountedAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "النقدية الفعلية", ""
    AddField tdf, "Difference", "MONEY", 0, True, "0", _
             "", "", "الفرق", "سالب = عجز، موجب = زيادة"
    AddField tdf, "Destination", "TEXT", 10, True, """MAIN""", _
             "In (""MAIN"",""OWNER"",""KEEP"")", "MAIN = الخزينة الرئيسية، OWNER = تسوية مع المالك، KEEP = يبقى في الصندوق", "الترحيل إلى", ""
    AddField tdf, "ToCashBoxID", "LONG", 0, False, "", _
             "", "", "الخزينة المستلمة", ""
    AddField tdf, "TransferAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المبلغ المرحَّل", ""
    AddField tdf, "KeptAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "المتبقي في الصندوق (عهدة)", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "ClosingID", True, True, False
    AddIndex tdf, "UX_ClosingNumber", "ClosingNumber", False, True, False
    AddIndex tdf, "IX_CashBoxID_ClosingDate", "CashBoxID,ClosingDate", False, False, False
    EndTable tdf, "تصفية يومية الكاشير: جرد نقدية صندوق الكاشير في نهاية الوردية وترحيلها للخزينة الرئيسية أو تسويتها مع المالك.", "[TransferAmount]+[KeptAmount]=[CountedAmount]", "المرحَّل + المتبقي = النقدية الفعلية"
End Sub

Private Sub CreateTable_Accounts()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Accounts") Then Exit Sub
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             ">0", "رقم الحساب أكبر من صفر", "رقم الحساب", ""
    AddField tdf, "AccountName", "TEXT", 100, True, "", _
             "", "", "اسم الحساب", ""
    AddField tdf, "AccountType", "TEXT", 10, True, "", _
             "In (""ASSET"",""LIABILITY"",""EQUITY"",""REVENUE"",""EXPENSE"")", "أصول، خصوم، حقوق ملكية، إيرادات، مصروفات", "نوع الحساب", ""
    AddField tdf, "ParentCode", "LONG", 0, False, "", _
             "", "", "الحساب الرئيسي", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddIndex tdf, "PrimaryKey", "AccountCode", True, True, False
    EndTable tdf, "دليل الحسابات: الحسابات التي تُرحَّل إليها القيود. حسابات الصناديق (110000 + رقم الصندوق) وأنواع المصروفات (530000 + رقم النوع) تُنشأ تلقائيًا.", "", ""
End Sub

Private Sub CreateTable_JournalSourceTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalSourceTypes") Then Exit Sub
    AddField tdf, "SourceType", "TEXT", 20, True, "", _
             "", "", "نوع العملية", ""
    AddField tdf, "TypeName", "TEXT", 50, True, "", _
             "", "", "الاسم", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "الترتيب", ""
    AddIndex tdf, "PrimaryKey", "SourceType", True, True, False
    EndTable tdf, "أنواع مصادر القيود: أنواع العمليات التي يُنشأ عنها قيد آلي، ومنها يُعرف أصل القيد.", "", ""
End Sub

Private Sub CreateTable_JournalEntries()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalEntries") Then Exit Sub
    AddField tdf, "EntryID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "EntryNumber", "TEXT", 20, True, "", _
             "", "", "رقم القيد", ""
    AddField tdf, "EntryDate", "DATETIME", 0, True, "", _
             "", "", "تاريخ القيد", ""
    AddField tdf, "SourceType", "TEXT", 20, True, "", _
             "", "", "نوع العملية", ""
    AddField tdf, "SourceID", "LONG", 0, True, "", _
             "", "", "رقم العملية الداخلي", ""
    AddField tdf, "SourceNumber", "TEXT", 20, False, "", _
             "", "", "رقم مستند العملية", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "البيان", ""
    AddField tdf, "TotalDebit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "إجمالي المدين", ""
    AddField tdf, "TotalCredit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "إجمالي الدائن", ""
    AddField tdf, "LineCount", "INT", 0, True, "0", _
             "", "", "عدد الأسطر", ""
    AddField tdf, "Signature", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "بصمة القيد", "تكشف تغيّر العملية بعد إنشاء القيد"
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "آخر تحديث", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "EntryID", True, True, False
    AddIndex tdf, "UX_EntryNumber", "EntryNumber", False, True, False
    AddIndex tdf, "UX_SourceType_SourceID", "SourceType,SourceID", False, True, False
    AddIndex tdf, "IX_EntryDate", "EntryDate", False, False, False
    EndTable tdf, "قيود اليومية: قيد آلي لكل عملية، مربوط بأصلها (SourceType + SourceID). يُحدَّث إذا تغيرت العملية.", "[TotalDebit]=[TotalCredit]", "القيد غير متوازن: المدين يجب أن يساوي الدائن"
End Sub

Private Sub CreateTable_JournalLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalLines") Then Exit Sub
    AddField tdf, "JournalLineID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "EntryID", "LONG", 0, True, "", _
             "", "", "القيد", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "رقم السطر", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "الحساب", ""
    AddField tdf, "Debit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "مدين", ""
    AddField tdf, "Credit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "دائن", ""
    AddField tdf, "LineText", "TEXT", 255, False, "", _
             "", "", "البيان", ""
    AddIndex tdf, "PrimaryKey", "JournalLineID", True, True, False
    AddIndex tdf, "UX_EntryID_LineNumber", "EntryID,LineNumber", False, True, False
    AddIndex tdf, "IX_AccountCode", "AccountCode", False, False, False
    EndTable tdf, "أسطر القيود: الطرف المدين والطرف الدائن لكل قيد.", "", ""
End Sub

Private Sub CreateTable_TransactionTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "TransactionTypes") Then Exit Sub
    AddField tdf, "TransactionTypeID", "LONG", 0, True, "", _
             "", "", "رقم النوع", ""
    AddField tdf, "TypeCode", "TEXT", 20, True, "", _
             "", "", "رمز النوع", ""
    AddField tdf, "TypeName", "TEXT", 50, True, "", _
             "", "", "نوع الحركة", ""
    AddField tdf, "Direction", "INT", 0, True, "", _
             "In (-1,0,1)", "الاتجاه 1 أو -1 أو 0", "الاتجاه", ""
    AddField tdf, "IsManual", "BOOL", 0, False, "False", _
             "", "", "متاح للإدخال اليدوي", ""
    AddIndex tdf, "PrimaryKey", "TransactionTypeID", True, True, False
    AddIndex tdf, "UX_TypeCode", "TypeCode", False, True, False
    EndTable tdf, "أنواع حركات المخزون: أنواع الحركة وإشارتها: +1 تزيد المخزون، -1 تنقصه، 0 تسوية بالإشارة.", "", ""
End Sub

Private Sub CreateTable_InventoryTransactions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "InventoryTransactions") Then Exit Sub
    AddField tdf, "TransactionID", "AUTO", 0, False, "", _
             "", "", "رقم الحركة", ""
    AddField tdf, "TransactionDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الحركة", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "TransactionTypeID", "LONG", 0, True, "", _
             "", "", "نوع الحركة", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             "<>0", "الكمية لا يمكن أن تكون صفرًا", "الكمية (+/-)", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", ""
    AddField tdf, "QuantityAfter", "QTY", 0, True, "0", _
             "", "", "الرصيد بعد الحركة", ""
    AddField tdf, "ReferenceType", "TEXT", 20, False, "", _
             "", "", "نوع المستند", "SALE, SALES_RETURN, PURCHASE, PURCHASE_RETURN, STOCK_COUNT, MANUAL"
    AddField tdf, "ReferenceID", "LONG", 0, False, "", _
             "", "", "رقم المستند الداخلي", ""
    AddField tdf, "ReferenceNumber", "TEXT", 20, False, "", _
             "", "", "رقم المستند", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "الموظف", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddIndex tdf, "PrimaryKey", "TransactionID", True, True, False
    AddIndex tdf, "IX_TransactionDate", "TransactionDate", False, False, False
    AddIndex tdf, "IX_ReferenceType_ReferenceID", "ReferenceType,ReferenceID", False, False, False
    EndTable tdf, "حركة المخزون: دفتر أستاذ المخزون: الكمية مخزنة بإشارتها، ورصيد أي صنف = مجموع Quantity.", "", ""
End Sub

Private Sub CreateTable_StockCounts()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "StockCounts") Then Exit Sub
    AddField tdf, "StockCountID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "CountNumber", "TEXT", 20, True, "", _
             "", "", "رقم الجرد", ""
    AddField tdf, "CountDate", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الجرد", ""
    AddField tdf, "CategoryID", "LONG", 0, False, "", _
             "", "", "تصنيف محدد (اختياري)", ""
    AddField tdf, "Status", "TEXT", 10, True, """OPEN""", _
             "In (""OPEN"",""POSTED"",""CANCELLED"")", "OPEN مفتوح، POSTED مُرحّل، CANCELLED ملغى", "الحالة", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "أجراه", ""
    AddField tdf, "PostedAt", "DATETIME", 0, False, "", _
             "", "", "تاريخ الترحيل", ""
    AddField tdf, "PostedByID", "LONG", 0, False, "", _
             "", "", "رحّله", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddIndex tdf, "PrimaryKey", "StockCountID", True, True, False
    AddIndex tdf, "UX_CountNumber", "CountNumber", False, True, False
    EndTable tdf, "جلسات الجرد: رأس عملية الجرد؛ تبقى مفتوحة حتى الترحيل الذي ينشئ حركات التسوية.", "", ""
End Sub

Private Sub CreateTable_StockCountDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "StockCountDetails") Then Exit Sub
    AddField tdf, "StockCountDetailID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "StockCountID", "LONG", 0, True, "", _
             "", "", "الجرد", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "المنتج", ""
    AddField tdf, "SystemQuantity", "QTY", 0, True, "0", _
             "", "", "الكمية المسجلة", ""
    AddField tdf, "ActualQuantity", "QTY", 0, False, "", _
             "Is Null Or >=0", "الكمية الفعلية لا يمكن أن تكون سالبة", "الكمية الفعلية", ""
    AddField tdf, "Difference", "QTY", 0, True, "0", _
             "", "", "الفرق", "ActualQuantity - SystemQuantity"
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "تكلفة الوحدة", ""
    AddField tdf, "DifferenceValue", "MONEY", 0, True, "0", _
             "", "", "قيمة الفرق", "سالب = عجز، موجب = زيادة"
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "ملاحظات", ""
    AddIndex tdf, "PrimaryKey", "StockCountDetailID", True, True, False
    AddIndex tdf, "UX_StockCountID_ProductID", "StockCountID,ProductID", False, True, False
    EndTable tdf, "تفاصيل الجرد: الكمية المسجلة والفعلية والفرق لكل صنف في جلسة الجرد.", "", ""
End Sub

Private Sub CreateTable_AuditLog()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "AuditLog") Then Exit Sub
    AddField tdf, "LogID", "AUTO", 0, False, "", _
             "", "", "رقم السجل", ""
    AddField tdf, "LogDate", "DATETIME", 0, True, "Now()", _
             "", "", "التاريخ", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "الموظف", ""
    AddField tdf, "ActionType", "TEXT", 30, True, "", _
             "", "", "نوع العملية", ""
    AddField tdf, "ObjectName", "TEXT", 50, False, "", _
             "", "", "الكائن", ""
    AddField tdf, "RecordID", "TEXT", 30, False, "", _
             "", "", "رقم السجل المتأثر", ""
    AddField tdf, "Details", "MEMO", 0, False, "", _
             "", "", "التفاصيل", ""
    AddField tdf, "ComputerName", "TEXT", 50, False, "", _
             "", "", "اسم الجهاز", ""
    AddIndex tdf, "PrimaryKey", "LogID", True, True, False
    AddIndex tdf, "IX_LogDate", "LogDate", False, False, False
    AddIndex tdf, "IX_ActionType", "ActionType", False, False, False
    EndTable tdf, "سجل العمليات: يسجل الدخول والخروج والعمليات الحساسة (تجاوز المخزون، تعديل الأسعار، النسخ الاحتياطي).", "", ""
End Sub

Private Sub CreateTable_LabelSettings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "LabelSettings") Then Exit Sub
    AddField tdf, "LabelSettingID", "LONG", 0, True, "1", _
             "=1", "يسمح بسجل إعدادات واحد فقط", "رقم الإعداد", ""
    AddField tdf, "PrinterName", "TEXT", 255, False, "", _
             "", "", "طابعة الملصقات", "فارغ = الطابعة الافتراضية"
    AddField tdf, "LabelWidth", "QTY", 0, True, "38", _
             "Between 15 And 210", "عرض الملصق من 15 إلى 210 مم", "عرض الملصق (مم)", ""
    AddField tdf, "LabelHeight", "QTY", 0, True, "25", _
             "Between 10 And 297", "ارتفاع الملصق من 10 إلى 297 مم", "ارتفاع الملصق (مم)", ""
    AddField tdf, "LabelsAcross", "BYTE", 0, True, "1", _
             "Between 1 And 10", "من 1 إلى 10 ملصقات في الصف", "عدد الملصقات في الصف", ""
    AddField tdf, "ColumnGap", "QTY", 0, True, "2", _
             "Between 0 And 50", "من 0 إلى 50 مم", "المسافة بين الأعمدة (مم)", ""
    AddField tdf, "RowGap", "QTY", 0, True, "0", _
             "Between 0 And 50", "من 0 إلى 50 مم", "المسافة بين الصفوف (مم)", ""
    AddField tdf, "MarginTop", "QTY", 0, True, "0", _
             "Between 0 And 50", "من 0 إلى 50 مم", "الهامش العلوي (مم)", ""
    AddField tdf, "MarginBottom", "QTY", 0, True, "0", _
             "Between 0 And 50", "من 0 إلى 50 مم", "الهامش السفلي (مم)", ""
    AddField tdf, "MarginLeft", "QTY", 0, True, "0", _
             "Between 0 And 50", "من 0 إلى 50 مم", "الهامش الأيسر (مم)", ""
    AddField tdf, "MarginRight", "QTY", 0, True, "0", _
             "Between 0 And 50", "من 0 إلى 50 مم", "الهامش الأيمن (مم)", ""
    AddField tdf, "BarHeight", "QTY", 0, True, "10", _
             "Between 3 And 100", "ارتفاع الباركود من 3 إلى 100 مم", "ارتفاع الباركود (مم)", ""
    AddField tdf, "BarWidth", "QTY", 0, True, "0.25", _
             "Between 0.1 And 1", "عرض أرفع خط من 0.1 إلى 1 مم (المعتاد 0.25 - 0.33)", "عرض أرفع خط (مم)", ""
    AddField tdf, "TopLine1", "TEXT", 10, True, """STORE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "اختر من القائمة", "السطر الأول أعلى الباركود", ""
    AddField tdf, "TopLine2", "TEXT", 10, True, """NAME""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "اختر من القائمة", "السطر الثاني أعلى الباركود", ""
    AddField tdf, "BottomLine1", "TEXT", 10, True, """BARCODE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "اختر من القائمة", "السطر الأول أسفل الباركود", ""
    AddField tdf, "BottomLine2", "TEXT", 10, True, """PRICE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "اختر من القائمة", "السطر الثاني أسفل الباركود", ""
    AddField tdf, "ShortName", "TEXT", 30, False, "", _
             "", "", "الاسم المختصر للمحل", "يُطبع إذا اخترت «الاسم المختصر»"
    AddField tdf, "FontSize", "BYTE", 0, True, "7", _
             "Between 5 And 16", "حجم الخط من 5 إلى 16", "حجم الخط", ""
    AddIndex tdf, "PrimaryKey", "LabelSettingID", True, True, False
    EndTable tdf, "إعدادات ملصقات الباركود: سجل واحد: مقاس الملصق والورق والهوامش، وحجم الباركود، والنصوص أعلاه وأسفله.", "", ""
End Sub

'------------------------------------------------------------------------------
' Generated: lookup / initial data
'------------------------------------------------------------------------------
Private Sub SeedAll()
    Seed_Settings
    Seed_Sequences
    Seed_Roles
    Seed_Permissions
    Seed_RolePermissions
    Seed_Employees
    Seed_Categories
    Seed_Units
    Seed_PaymentMethods
    Seed_CashBoxes
    Seed_Customers
    Seed_ExpenseTypes
    Seed_Accounts
    Seed_JournalSourceTypes
    Seed_TransactionTypes
    Seed_LabelSettings
End Sub

Private Sub Seed_Settings()
    If Not BeginSeed("Settings", False) Then Exit Sub
    ExecSeed "INSERT INTO [Settings] ([SettingID], [StoreName], [CountryCode], [VATRate], [PricesIncludeVAT], [AllowNegativeStock], [CurrencyCode], [DefaultCustomerID], [ZatcaPhase], [LastInvoiceHash], [BackupKeepCount], [SlowMovingDays], [ReceiptFooter]) VALUES (1, 'اسم المحل', 'SA', 0.15, True, False, 'SAR', 1, 1, 'NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ==', 30, 90, 'شكرًا لزيارتكم')"
    EndSeed "Settings", 1
End Sub

Private Sub Seed_Sequences()
    If Not BeginSeed("Sequences", True) Then Exit Sub
    SeedRow "[SequenceName] = 'SALES_INVOICE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_INVOICE', 'INV-', 1, 6, 'فواتير المبيعات')"
    SeedRow "[SequenceName] = 'SALES_RETURN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_RETURN', 'CRN-', 1, 6, 'مرتجعات المبيعات (إشعار دائن)')"
    SeedRow "[SequenceName] = 'PURCHASE_INVOICE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_INVOICE', 'PUR-', 1, 6, 'فواتير المشتريات')"
    SeedRow "[SequenceName] = 'PURCHASE_RETURN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_RETURN', 'PRT-', 1, 6, 'مرتجعات المشتريات')"
    SeedRow "[SequenceName] = 'CUSTOMER_PAYMENT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CUSTOMER_PAYMENT', 'RCV-', 1, 6, 'سندات القبض من العملاء')"
    SeedRow "[SequenceName] = 'SUPPLIER_PAYMENT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SUPPLIER_PAYMENT', 'PAY-', 1, 6, 'سندات الصرف للموردين')"
    SeedRow "[SequenceName] = 'EXPENSE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('EXPENSE', 'EXP-', 1, 6, 'المصروفات')"
    SeedRow "[SequenceName] = 'STOCK_COUNT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_COUNT', 'CNT-', 1, 5, 'جلسات الجرد')"
    SeedRow "[SequenceName] = 'STOCK_ADJUST'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_ADJUST', 'ADJ-', 1, 6, 'حركات المخزون اليدوية')"
    SeedRow "[SequenceName] = 'PRODUCT_CODE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PRODUCT_CODE', 'P', 1, 5, 'أكواد المنتجات')"
    SeedRow "[SequenceName] = 'ZATCA_ICV'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('ZATCA_ICV', Null, 1, 0, 'عدّاد ICV لمستندات فاتورة')"
    SeedRow "[SequenceName] = 'CASH_IN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_IN', 'CIN-', 1, 6, 'سندات قبض النقدية (الخزينة)')"
    SeedRow "[SequenceName] = 'CASH_OUT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_OUT', 'COT-', 1, 6, 'سندات صرف النقدية (الخزينة)')"
    SeedRow "[SequenceName] = 'CASH_TRANSFER'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_TRANSFER', 'TRF-', 1, 6, 'التحويل بين الصناديق')"
    SeedRow "[SequenceName] = 'CASH_CLOSING'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_CLOSING', 'CLS-', 1, 6, 'تصفية يومية الكاشير')"
    SeedRow "[SequenceName] = 'JOURNAL'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('JOURNAL', 'JV-', 1, 6, 'قيود اليومية')"
    EndSeed "Sequences", 16
End Sub

Private Sub Seed_Roles()
    If Not BeginSeed("Roles", False) Then Exit Sub
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (1, 'ADMIN', 'مدير النظام', 'جميع الصلاحيات')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (2, 'MANAGER', 'مدير', 'المبيعات والمشتريات والمخزون والتقارير')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (3, 'CASHIER', 'كاشير', 'المبيعات والعملاء فقط')"
    EndSeed "Roles", 3
End Sub

Private Sub Seed_Permissions()
    If Not BeginSeed("Permissions", True) Then Exit Sub
    If SeedRow("[PermissionKey] = 'SALES_POS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_POS', 'نقطة البيع', 'المبيعات', 10)") Then GrantNewPermission "SALES_POS", "1,2,3"
    If SeedRow("[PermissionKey] = 'SALES_VIEW'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_VIEW', 'عرض وإعادة طباعة الفواتير', 'المبيعات', 11)") Then GrantNewPermission "SALES_VIEW", "1,2,3"
    If SeedRow("[PermissionKey] = 'SALES_RETURN'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_RETURN', 'مرتجعات المبيعات', 'المبيعات', 12)") Then GrantNewPermission "SALES_RETURN", "1,2"
    If SeedRow("[PermissionKey] = 'PRICE_OVERRIDE'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRICE_OVERRIDE', 'تعديل سعر البيع في الفاتورة', 'المبيعات', 13)") Then GrantNewPermission "PRICE_OVERRIDE", "1,2"
    If SeedRow("[PermissionKey] = 'DISCOUNT_OVERRIDE'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DISCOUNT_OVERRIDE', 'خصم أعلى من الحد المسموح', 'المبيعات', 14)") Then GrantNewPermission "DISCOUNT_OVERRIDE", "1,2"
    If SeedRow("[PermissionKey] = 'ALLOW_NEGATIVE_STOCK'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('ALLOW_NEGATIVE_STOCK', 'البيع بكمية أكبر من المتوفر', 'المبيعات', 15)") Then GrantNewPermission "ALLOW_NEGATIVE_STOCK", "1"
    If SeedRow("[PermissionKey] = 'CUSTOMERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMERS', 'إدارة العملاء', 'العملاء', 20)") Then GrantNewPermission "CUSTOMERS", "1,2,3"
    If SeedRow("[PermissionKey] = 'CUSTOMER_PAYMENTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMER_PAYMENTS', 'سندات القبض', 'العملاء', 21)") Then GrantNewPermission "CUSTOMER_PAYMENTS", "1,2,3"
    If SeedRow("[PermissionKey] = 'PURCHASES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASES', 'فواتير المشتريات', 'المشتريات', 30)") Then GrantNewPermission "PURCHASES", "1,2"
    If SeedRow("[PermissionKey] = 'PURCHASE_RETURN'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASE_RETURN', 'مرتجعات المشتريات', 'المشتريات', 31)") Then GrantNewPermission "PURCHASE_RETURN", "1,2"
    If SeedRow("[PermissionKey] = 'SUPPLIERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIERS', 'إدارة الموردين', 'الموردون', 40)") Then GrantNewPermission "SUPPLIERS", "1,2"
    If SeedRow("[PermissionKey] = 'SUPPLIER_PAYMENTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIER_PAYMENTS', 'سندات الصرف', 'الموردون', 41)") Then GrantNewPermission "SUPPLIER_PAYMENTS", "1,2"
    If SeedRow("[PermissionKey] = 'PRODUCTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRODUCTS', 'إدارة المنتجات والأسعار', 'المخزون', 50)") Then GrantNewPermission "PRODUCTS", "1,2"
    If SeedRow("[PermissionKey] = 'INVENTORY_ADJUST'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('INVENTORY_ADJUST', 'إضافة وخصم مخزون يدوي', 'المخزون', 51)") Then GrantNewPermission "INVENTORY_ADJUST", "1,2"
    If SeedRow("[PermissionKey] = 'STOCK_COUNT'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('STOCK_COUNT', 'الجرد', 'المخزون', 52)") Then GrantNewPermission "STOCK_COUNT", "1,2"
    If SeedRow("[PermissionKey] = 'EXPENSES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('EXPENSES', 'المصروفات', 'المصروفات', 60)") Then GrantNewPermission "EXPENSES", "1,2"
    If SeedRow("[PermissionKey] = 'CASH_BOX'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CASH_BOX', 'الخزينة: سندات القبض والصرف والتحويل والصناديق', 'الخزينة', 65)") Then GrantNewPermission "CASH_BOX", "1,2"
    If SeedRow("[PermissionKey] = 'CASH_CLOSING'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CASH_CLOSING', 'تصفية يومية الكاشير', 'الخزينة', 66)") Then GrantNewPermission "CASH_CLOSING", "1,2,3"
    If SeedRow("[PermissionKey] = 'JOURNAL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('JOURNAL', 'قيود اليومية ودليل الحسابات وميزان المراجعة', 'الحسابات', 75)") Then GrantNewPermission "JOURNAL", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS', 'التقارير التشغيلية', 'التقارير', 70)") Then GrantNewPermission "REPORTS", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS_PROFIT'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS_PROFIT', 'تقارير الأرباح والضريبة', 'التقارير', 71)") Then GrantNewPermission "REPORTS_PROFIT", "1,2"
    If SeedRow("[PermissionKey] = 'DASHBOARD_FINANCIAL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DASHBOARD_FINANCIAL', 'الأرقام المالية في لوحة التحكم', 'التقارير', 72)") Then GrantNewPermission "DASHBOARD_FINANCIAL", "1,2"
    If SeedRow("[PermissionKey] = 'SETTINGS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SETTINGS', 'إعدادات المحل', 'النظام', 80)") Then GrantNewPermission "SETTINGS", "1"
    If SeedRow("[PermissionKey] = 'USERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('USERS', 'المستخدمون والصلاحيات', 'النظام', 81)") Then GrantNewPermission "USERS", "1"
    If SeedRow("[PermissionKey] = 'BACKUP'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BACKUP', 'النسخ الاحتياطي', 'النظام', 82)") Then GrantNewPermission "BACKUP", "1"
    EndSeed "Permissions", 25
End Sub

Private Sub Seed_RolePermissions()
    If Not BeginSeed("RolePermissions", False) Then Exit Sub
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SALES_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PRICE_OVERRIDE')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'DISCOUNT_OVERRIDE')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'ALLOW_NEGATIVE_STOCK')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CUSTOMER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PURCHASES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PURCHASE_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SUPPLIERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SUPPLIER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PRODUCTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'INVENTORY_ADJUST')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'STOCK_COUNT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'EXPENSES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CASH_BOX')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CASH_CLOSING')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'JOURNAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SETTINGS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'USERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'BACKUP')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SALES_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'PRICE_OVERRIDE')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'DISCOUNT_OVERRIDE')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CUSTOMER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'PURCHASES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'PURCHASE_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SUPPLIERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SUPPLIER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'PRODUCTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'INVENTORY_ADJUST')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'STOCK_COUNT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'EXPENSES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CASH_BOX')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CASH_CLOSING')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'JOURNAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CASH_CLOSING')"
    EndSeed "RolePermissions", 51
End Sub

Private Sub Seed_Employees()
    If Not BeginSeed("Employees", False) Then Exit Sub
    ExecSeed "INSERT INTO [Employees] ([EmployeeID], [EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], [MustChangePassword], [IsActive]) VALUES (1, 'مدير النظام', 'مدير النظام', 'admin', 1, 1, True, True)"
    EndSeed "Employees", 1
End Sub

Private Sub Seed_Categories()
    If Not BeginSeed("Categories", False) Then Exit Sub
    ExecSeed "INSERT INTO [Categories] ([CategoryID], [CategoryName], [Description]) VALUES (1, 'عام', 'تصنيف افتراضي')"
    EndSeed "Categories", 1
End Sub

Private Sub Seed_Units()
    If Not BeginSeed("Units", False) Then Exit Sub
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (1, 'حبة', 'PCE')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (2, 'علبة', 'BX')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (3, 'كرتون', 'CT')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (4, 'باكيت', 'PK')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (5, 'كيلو', 'KGM')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (6, 'لتر', 'LTR')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (7, 'متر', 'MTR')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (8, 'طقم', 'SET')"
    EndSeed "Units", 8
End Sub

Private Sub Seed_PaymentMethods()
    If Not BeginSeed("PaymentMethods", False) Then Exit Sub
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (1, 'نقدي', '10', 1)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (2, 'مدى / بطاقة', '48', 2)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (3, 'تحويل بنكي', '42', 3)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (4, 'محفظة إلكترونية', '1', 4)"
    EndSeed "PaymentMethods", 4
End Sub

Private Sub Seed_CashBoxes()
    If Not BeginSeed("CashBoxes", False) Then Exit Sub
    ExecSeed "INSERT INTO [CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (1, 'الخزينة الرئيسية', 'MAIN')"
    ExecSeed "INSERT INTO [CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (2, 'صندوق الكاشير', 'CASHIER')"
    EndSeed "CashBoxes", 2
End Sub

Private Sub Seed_Customers()
    If Not BeginSeed("Customers", False) Then Exit Sub
    ExecSeed "INSERT INTO [Customers] ([CustomerID], [CustomerName], [AllowCredit], [IsSystem]) VALUES (1, 'عميل نقدي', False, True)"
    EndSeed "Customers", 1
End Sub

Private Sub Seed_ExpenseTypes()
    If Not BeginSeed("ExpenseTypes", False) Then Exit Sub
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (1, 'الإيجار')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (2, 'الكهرباء')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (3, 'المياه')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (4, 'الإنترنت والاتصالات')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (5, 'النقل')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (6, 'الصيانة')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (7, 'الرواتب')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (8, 'المستلزمات')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (9, 'مصروفات أخرى')"
    EndSeed "ExpenseTypes", 9
End Sub

Private Sub Seed_Accounts()
    If Not BeginSeed("Accounts", True) Then Exit Sub
    SeedRow "[AccountCode] = 1100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1100, 'النقدية بالخزينة والصناديق', 'ASSET', Null)"
    SeedRow "[AccountCode] = 110001", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (110001, 'الخزينة الرئيسية', 'ASSET', 1100)"
    SeedRow "[AccountCode] = 110002", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (110002, 'صندوق الكاشير', 'ASSET', 1100)"
    SeedRow "[AccountCode] = 1190", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1190, 'نقدية غير موزعة على صندوق', 'ASSET', Null)"
    SeedRow "[AccountCode] = 1200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1200, 'البنك والشبكة (مدى والتحويلات)', 'ASSET', Null)"
    SeedRow "[AccountCode] = 1300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1300, 'ذمم العملاء', 'ASSET', Null)"
    SeedRow "[AccountCode] = 1400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1400, 'المخزون', 'ASSET', Null)"
    SeedRow "[AccountCode] = 1500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1500, 'ضريبة القيمة المضافة - مدخلات', 'ASSET', Null)"
    SeedRow "[AccountCode] = 1600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (1600, 'سلف الموظفين', 'ASSET', Null)"
    SeedRow "[AccountCode] = 2100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (2100, 'ذمم الموردين', 'LIABILITY', Null)"
    SeedRow "[AccountCode] = 2200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (2200, 'ضريبة القيمة المضافة - مخرجات', 'LIABILITY', Null)"
    SeedRow "[AccountCode] = 3100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (3100, 'جاري المالك', 'EQUITY', Null)"
    SeedRow "[AccountCode] = 3900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (3900, 'أرصدة افتتاحية', 'EQUITY', Null)"
    SeedRow "[AccountCode] = 4100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (4100, 'المبيعات', 'REVENUE', Null)"
    SeedRow "[AccountCode] = 4110", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (4110, 'مردودات المبيعات', 'REVENUE', Null)"
    SeedRow "[AccountCode] = 4200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (4200, 'إيرادات أخرى', 'REVENUE', Null)"
    SeedRow "[AccountCode] = 4300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (4300, 'زيادة الصناديق', 'REVENUE', Null)"
    SeedRow "[AccountCode] = 5100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (5100, 'تكلفة البضاعة المباعة', 'EXPENSE', Null)"
    SeedRow "[AccountCode] = 5200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (5200, 'فروقات وتسويات المخزون', 'EXPENSE', Null)"
    SeedRow "[AccountCode] = 5300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (5300, 'المصروفات التشغيلية', 'EXPENSE', Null)"
    SeedRow "[AccountCode] = 5400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (5400, 'عجز الصناديق', 'EXPENSE', Null)"
    SeedRow "[AccountCode] = 5900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode]) VALUES (5900, 'مصروفات أخرى', 'EXPENSE', Null)"
    EndSeed "Accounts", 22
End Sub

Private Sub Seed_JournalSourceTypes()
    If Not BeginSeed("JournalSourceTypes", True) Then Exit Sub
    SeedRow "[SourceType] = 'SALE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SALE', 'فاتورة بيع', 1)"
    SeedRow "[SourceType] = 'SALES_RETURN'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SALES_RETURN', 'مرتجع بيع', 2)"
    SeedRow "[SourceType] = 'PURCHASE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PURCHASE', 'فاتورة شراء', 3)"
    SeedRow "[SourceType] = 'PURCHASE_RETURN'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PURCHASE_RETURN', 'مرتجع شراء', 4)"
    SeedRow "[SourceType] = 'CUSTOMER_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CUSTOMER_PAYMENT', 'سند قبض من عميل', 5)"
    SeedRow "[SourceType] = 'SUPPLIER_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SUPPLIER_PAYMENT', 'سند صرف لمورد', 6)"
    SeedRow "[SourceType] = 'EXPENSE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('EXPENSE', 'مصروف', 7)"
    SeedRow "[SourceType] = 'CASH_VOUCHER'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CASH_VOUCHER', 'سند نقدية', 8)"
    SeedRow "[SourceType] = 'STOCK_MOVE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('STOCK_MOVE', 'حركة مخزون يدوية', 9)"
    SeedRow "[SourceType] = 'STOCK_COUNT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('STOCK_COUNT', 'تسوية جرد', 10)"
    SeedRow "[SourceType] = 'BOX_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('BOX_OPENING', 'رصيد افتتاحي لصندوق', 11)"
    SeedRow "[SourceType] = 'CUSTOMER_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CUSTOMER_OPENING', 'رصيد افتتاحي لعميل', 12)"
    SeedRow "[SourceType] = 'SUPPLIER_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SUPPLIER_OPENING', 'رصيد افتتاحي لمورد', 13)"
    EndSeed "JournalSourceTypes", 13
End Sub

Private Sub Seed_TransactionTypes()
    If Not BeginSeed("TransactionTypes", False) Then Exit Sub
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (1, 'PURCHASE', 'شراء', 1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (2, 'SALE', 'بيع', -1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (3, 'PURCHASE_RETURN', 'مرتجع شراء', -1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (4, 'SALES_RETURN', 'مرتجع بيع', 1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (5, 'STOCK_IN', 'إضافة مخزون', 1, True)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (6, 'STOCK_OUT', 'خصم مخزون', -1, True)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (7, 'ADJUSTMENT', 'تسوية جرد', 0, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (8, 'OPENING', 'رصيد افتتاحي', 1, True)"
    EndSeed "TransactionTypes", 8
End Sub

Private Sub Seed_LabelSettings()
    If Not BeginSeed("LabelSettings", False) Then Exit Sub
    ExecSeed "INSERT INTO [LabelSettings] ([LabelSettingID]) VALUES (1)"
    EndSeed "LabelSettings", 1
End Sub
