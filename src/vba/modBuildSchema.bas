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

Private Const SCHEMA_TABLES As String = "Settings,Sequences,Roles,Permissions,RolePermissions,Employees,Screens,UserScreens,Activations,Categories,Units,PaymentMethods,CashBoxes,Suppliers,Customers,Products,SalesInvoices,SalesInvoiceDetails,SalesReturns,SalesReturnDetails,PurchaseInvoices,PurchaseInvoiceDetails,PurchaseReturns,PurchaseReturnDetails,CustomerPayments,SupplierPayments,ExpenseTypes,Expenses,CashVouchers,CashClosings,Accounts,JournalSourceTypes,JournalEntries,JournalLines,ManualEntries,ManualEntryLines,TransactionTypes,InventoryTransactions,StockCounts,StockCountDetails,AuditLog,LabelSettings"
Private Const EXPECTED_FIELD_COUNTS As String = "Settings=32;Sequences=5;Roles=4;Permissions=4;RolePermissions=2;Employees=19;Screens=8;UserScreens=6;Activations=6;Categories=8;Units=4;PaymentMethods=5;CashBoxes=8;Suppliers=15;Customers=21;Products=23;SalesInvoices=35;SalesInvoiceDetails=14;SalesReturns=29;SalesReturnDetails=14;PurchaseInvoices=18;PurchaseInvoiceDetails=11;PurchaseReturns=18;PurchaseReturnDetails=11;CustomerPayments=11;SupplierPayments=11;ExpenseTypes=3;Expenses=13;CashVouchers=14;CashClosings=18;Accounts=14;JournalSourceTypes=3;JournalEntries=13;JournalLines=7;ManualEntries=10;ManualEntryLines=7;TransactionTypes=5;InventoryTransactions=13;StockCounts=9;StockCountDetails=9;AuditLog=8;LabelSettings=19"
Private Const EXPECTED_SEED_COUNTS As String = "Settings=1;Sequences=17;Roles=3;Permissions=26;RolePermissions=53;Employees=1;Screens=37;Categories=1;Units=8;PaymentMethods=4;CashBoxes=2;Customers=1;ExpenseTypes=9;Accounts=75;JournalSourceTypes=14;TransactionTypes=8;LabelSettings=1"

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
    CreateTable_Screens
    CreateTable_UserScreens
    CreateTable_Activations
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
    CreateTable_ManualEntries
    CreateTable_ManualEntryLines
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
    AddField tdf, "AllowAdminCompanyName", "BOOL", 0, False, "False", _
             "", "", "السماح لمدير النظام بتغيير اسم المحل", ""
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
    AddField tdf, "IsDeveloper", "BOOL", 0, False, "False", _
             "", "", "المبرمج", ""
    AddField tdf, "CustomScreens", "BOOL", 0, False, "False", _
             "", "", "صلاحيات شاشات خاصة", ""
    AddIndex tdf, "PrimaryKey", "EmployeeID", True, True, False
    AddIndex tdf, "UX_Username", "Username", False, True, False
    EndTable tdf, "الموظفون والمستخدمون: كل موظف هو مستخدم للنظام؛ لا يُحذف بل يُعطَّل للحفاظ على سجل عملياته.", "", ""
End Sub

Private Sub CreateTable_Screens()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Screens") Then Exit Sub
    AddField tdf, "ScreenName", "TEXT", 64, True, "", _
             "", "", "اسم الشاشة في Access", ""
    AddField tdf, "ScreenTitle", "TEXT", 100, True, "", _
             "", "", "الشاشة", ""
    AddField tdf, "ModuleName", "TEXT", 50, False, "", _
             "", "", "القسم", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "الترتيب", ""
    AddField tdf, "PermissionKey", "TEXT", 50, False, "", _
             "", "", "صلاحية الدور", "فارغ = متاحة لكل المستخدمين؛ تُستخدم للمستخدم الذي ليست له صلاحيات شاشات خاصة"
    AddField tdf, "HasAdd", "BOOL", 0, False, "False", _
             "", "", "فيها إضافة / حفظ مستند", ""
    AddField tdf, "HasEdit", "BOOL", 0, False, "False", _
             "", "", "فيها تعديل", ""
    AddField tdf, "HasDelete", "BOOL", 0, False, "False", _
             "", "", "فيها حذف", ""
    AddIndex tdf, "PrimaryKey", "ScreenName", True, True, False
    EndTable tdf, "الشاشات: كل شاشة في البرنامج، وما ينطبق عليها من إضافة وتعديل وحذف، وصلاحية الدور التي تفتحها.", "", ""
End Sub

Private Sub CreateTable_UserScreens()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "UserScreens") Then Exit Sub
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "المستخدم", ""
    AddField tdf, "ScreenName", "TEXT", 64, True, "", _
             "", "", "الشاشة", ""
    AddField tdf, "CanOpen", "BOOL", 0, False, "False", _
             "", "", "فتح", ""
    AddField tdf, "CanAdd", "BOOL", 0, False, "False", _
             "", "", "إضافة", ""
    AddField tdf, "CanEdit", "BOOL", 0, False, "False", _
             "", "", "تعديل", ""
    AddField tdf, "CanDelete", "BOOL", 0, False, "False", _
             "", "", "حذف", ""
    AddIndex tdf, "PrimaryKey", "EmployeeID,ScreenName", True, True, False
    EndTable tdf, "صلاحيات الشاشات للمستخدم: للمستخدم الذي فُعّلت له «صلاحيات شاشات خاصة»: الشاشات التي يفتحها، والإضافة والتعديل والحذف في كل شاشة.", "", ""
End Sub

Private Sub CreateTable_Activations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Activations") Then Exit Sub
    AddField tdf, "ActivationID", "AUTO", 0, False, "", _
             "", "", "رقم التفعيل", ""
    AddField tdf, "MachineID", "TEXT", 24, True, "", _
             "", "", "رقم الجهاز", "بصمة لوحة الأم والمعالج وقرص النظام (modActivation.MachineID)"
    AddField tdf, "ActivationCode", "TEXT", 30, True, "", _
             "", "", "كود التفعيل", ""
    AddField tdf, "ComputerName", "TEXT", 64, False, "", _
             "", "", "اسم الجهاز", ""
    AddField tdf, "ActivatedAt", "DATETIME", 0, False, "Now()", _
             "", "", "تاريخ التفعيل", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "فعّله", ""
    AddIndex tdf, "PrimaryKey", "ActivationID", True, True, False
    AddIndex tdf, "UX_MachineID", "MachineID", False, True, False
    EndTable tdf, "تفعيل البرنامج: الأجهزة المفعّل عليها البرنامج: رقم الجهاز وكود التفعيل الصادر من المبرمج.", "", ""
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
             "", "", "الحساب الرئيسي", "فارغ للحسابات الخمسة في المستوى الأول فقط"
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "نشط", ""
    AddField tdf, "IsPosting", "BOOL", 0, False, "True", _
             "", "", "حساب فرعي (يقبل القيود)", ""
    AddField tdf, "IsSystem", "BOOL", 0, False, "False", _
             "", "", "حساب أساسي في النظام", ""
    AddField tdf, "AccountLevel", "BYTE", 0, True, "1", _
             "", "", "المستوى", ""
    AddField tdf, "TreeKey", "TEXT", 60, False, "", _
             "", "", "مفتاح الترتيب في الشجرة", "يحسبه البرنامج (modAccounts.RebuildAccountTree): رقم كل مستوى بعشر خانات"
    AddField tdf, "Level1Code", "LONG", 0, False, "", _
             "", "", "حساب المستوى 1", ""
    AddField tdf, "Level2Code", "LONG", 0, False, "", _
             "", "", "حساب المستوى 2", ""
    AddField tdf, "Level3Code", "LONG", 0, False, "", _
             "", "", "حساب المستوى 3", ""
    AddField tdf, "Level4Code", "LONG", 0, False, "", _
             "", "", "حساب المستوى 4", ""
    AddField tdf, "Level5Code", "LONG", 0, False, "", _
             "", "", "حساب المستوى 5", ""
    AddIndex tdf, "PrimaryKey", "AccountCode", True, True, False
    AddIndex tdf, "IX_ParentCode", "ParentCode", False, False, False
    AddIndex tdf, "IX_TreeKey", "TreeKey", False, False, False
    EndTable tdf, "دليل الحسابات (شجرة الحسابات): شجرة من خمسة مستويات على الأكثر: الحسابات الرئيسية (تجميعية) والحسابات الفرعية التي تُرحَّل إليها القيود. حسابات الصناديق (110000 + رقم الصندوق) وأنواع المصروفات (530000 + رقم النوع) تُنشأ تلقائيًا.", "", ""
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

Private Sub CreateTable_ManualEntries()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ManualEntries") Then Exit Sub
    AddField tdf, "ManualEntryID", "AUTO", 0, False, "", _
             "", "", "رقم داخلي", ""
    AddField tdf, "EntryNumber", "TEXT", 20, True, "", _
             "", "", "رقم القيد اليدوي", ""
    AddField tdf, "EntryDate", "DATE", 0, True, "Date()", _
             "", "", "تاريخ القيد", ""
    AddField tdf, "Description", "TEXT", 255, True, "", _
             "", "", "البيان", ""
    AddField tdf, "Reference", "TEXT", 50, False, "", _
             "", "", "المرجع", "رقم مستند خارجي: فاتورة، عقد، كشف بنك..."
    AddField tdf, "ReversalOfID", "LONG", 0, False, "", _
             "", "", "عكس القيد", "القيد اليدوي الذي يعكسه هذا القيد"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "إجمالي القيد", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "أدخله", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", "تاريخ الإنشاء", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "آخر تعديل", ""
    AddIndex tdf, "PrimaryKey", "ManualEntryID", True, True, False
    AddIndex tdf, "UX_EntryNumber", "EntryNumber", False, True, False
    AddIndex tdf, "IX_EntryDate", "EntryDate", False, False, False
    EndTable tdf, "القيود اليدوية: قيد يكتبه المحاسب بنفسه (مستحقات، تسويات، رأس المال، أرصدة افتتاحية...). يُرحَّل لليومية كأي عملية أخرى (SourceType = MANUAL)، ويُعدَّل أو يُحذف فيتبعه قيده.", "", ""
End Sub

Private Sub CreateTable_ManualEntryLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ManualEntryLines") Then Exit Sub
    AddField tdf, "ManualLineID", "AUTO", 0, False, "", _
             "", "", "رقم السطر الداخلي", ""
    AddField tdf, "ManualEntryID", "LONG", 0, True, "", _
             "", "", "القيد اليدوي", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "رقم السطر", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "الحساب", ""
    AddField tdf, "Debit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "مدين", ""
    AddField tdf, "Credit", "MONEY", 0, True, "0", _
             ">=0", "المبلغ لا يمكن أن يكون سالبًا", "دائن", ""
    AddField tdf, "LineText", "TEXT", 150, False, "", _
             "", "", "بيان السطر", ""
    AddIndex tdf, "PrimaryKey", "ManualLineID", True, True, False
    AddIndex tdf, "UX_ManualEntryID_LineNumber", "ManualEntryID,LineNumber", False, True, False
    AddIndex tdf, "IX_AccountCode", "AccountCode", False, False, False
    EndTable tdf, "أسطر القيود اليدوية: الطرف المدين والطرف الدائن للقيد اليدوي؛ كل سطر مدين أو دائن فقط.", "([Debit]=0 Or [Credit]=0) And [Debit]+[Credit]>0", "كل سطر مدين أو دائن فقط، وبمبلغ أكبر من صفر"
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
    Seed_Screens
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
    SeedRow "[SequenceName] = 'MANUAL_ENTRY'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('MANUAL_ENTRY', 'MJ-', 1, 6, 'القيود اليدوية')"
    EndSeed "Sequences", 17
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
    If SeedRow("[PermissionKey] = 'MANUAL_ENTRY'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('MANUAL_ENTRY', 'القيود اليدوية: إضافة وتعديل وحذف', 'الحسابات', 76)") Then GrantNewPermission "MANUAL_ENTRY", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS', 'التقارير التشغيلية', 'التقارير', 70)") Then GrantNewPermission "REPORTS", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS_PROFIT'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS_PROFIT', 'تقارير الأرباح والضريبة', 'التقارير', 71)") Then GrantNewPermission "REPORTS_PROFIT", "1,2"
    If SeedRow("[PermissionKey] = 'DASHBOARD_FINANCIAL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DASHBOARD_FINANCIAL', 'الأرقام المالية في لوحة التحكم', 'التقارير', 72)") Then GrantNewPermission "DASHBOARD_FINANCIAL", "1,2"
    If SeedRow("[PermissionKey] = 'SETTINGS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SETTINGS', 'إعدادات المحل', 'النظام', 80)") Then GrantNewPermission "SETTINGS", "1"
    If SeedRow("[PermissionKey] = 'USERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('USERS', 'المستخدمون والصلاحيات', 'النظام', 81)") Then GrantNewPermission "USERS", "1"
    If SeedRow("[PermissionKey] = 'BACKUP'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BACKUP', 'النسخ الاحتياطي', 'النظام', 82)") Then GrantNewPermission "BACKUP", "1"
    EndSeed "Permissions", 26
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
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'MANUAL_ENTRY')"
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
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'MANUAL_ENTRY')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CASH_CLOSING')"
    EndSeed "RolePermissions", 53
End Sub

Private Sub Seed_Employees()
    If Not BeginSeed("Employees", False) Then Exit Sub
    ExecSeed "INSERT INTO [Employees] ([EmployeeID], [EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], [MustChangePassword], [IsActive]) VALUES (1, 'مدير النظام', 'مدير النظام', 'admin', 1, 1, True, True)"
    EndSeed "Employees", 1
End Sub

Private Sub Seed_Screens()
    If Not BeginSeed("Screens", True) Then Exit Sub
    SeedRow "[ScreenName] = 'frmPOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPOS', 'نقطة البيع (المحلات)', 'المبيعات', 10, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmTouchPOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmTouchPOS', 'نقطة البيع (المطاعم)', 'المبيعات', 20, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmCafePOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCafePOS', 'نقطة البيع (الكافيهات)', 'المبيعات', 30, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmSalesInvoice'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSalesInvoice', 'عرض الفواتير وإعادة طباعتها', 'المبيعات', 40, 'SALES_VIEW', False, False, False)"
    SeedRow "[ScreenName] = 'frmSalesReturn'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSalesReturn', 'مرتجعات المبيعات', 'المبيعات', 50, 'SALES_RETURN', True, False, False)"
    SeedRow "[ScreenName] = 'frmCustomers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCustomers', 'العملاء', 'العملاء', 60, 'CUSTOMERS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCustomerPayment'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCustomerPayment', 'سندات القبض من العملاء', 'العملاء', 70, 'CUSTOMER_PAYMENTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseInvoice'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseInvoice', 'فواتير المشتريات', 'المشتريات', 80, 'PURCHASES', True, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseView'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseView', 'عرض فواتير المشتريات', 'المشتريات', 90, 'PURCHASES', False, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseReturn'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseReturn', 'مرتجعات المشتريات', 'المشتريات', 100, 'PURCHASE_RETURN', True, False, False)"
    SeedRow "[ScreenName] = 'frmSuppliers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSuppliers', 'الموردون', 'الموردون', 110, 'SUPPLIERS', True, True, True)"
    SeedRow "[ScreenName] = 'frmSupplierPayment'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSupplierPayment', 'سندات الصرف للموردين', 'الموردون', 120, 'SUPPLIER_PAYMENTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmProducts'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmProducts', 'المنتجات والأسعار', 'المخزون', 130, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCategories'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCategories', 'التصنيفات', 'المخزون', 140, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmUnits'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUnits', 'الوحدات', 'المخزون', 150, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmInventory'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmInventory', 'المخزون والحركات اليدوية', 'المخزون', 160, 'PRODUCTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmStockCount'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmStockCount', 'الجرد', 'المخزون', 170, 'STOCK_COUNT', True, False, False)"
    SeedRow "[ScreenName] = 'frmBarcodeLabels'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBarcodeLabels', 'ملصقات الباركود', 'المخزون', 180, 'PRODUCTS', False, False, False)"
    SeedRow "[ScreenName] = 'frmLabelSettings'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmLabelSettings', 'إعدادات الملصقات', 'المخزون', 190, 'PRODUCTS', False, True, False)"
    SeedRow "[ScreenName] = 'frmExpenses'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmExpenses', 'المصروفات', 'المصروفات', 200, 'EXPENSES', True, True, True)"
    SeedRow "[ScreenName] = 'frmExpenseTypes'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmExpenseTypes', 'أنواع المصروفات', 'المصروفات', 210, 'EXPENSES', True, True, True)"
    SeedRow "[ScreenName] = 'frmTreasury'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmTreasury', 'الخزينة', 'الخزينة', 220, 'CASH_CLOSING', False, False, False)"
    SeedRow "[ScreenName] = 'frmCashVoucher'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashVoucher', 'سندات النقدية والتحويل', 'الخزينة', 230, 'CASH_BOX', True, False, False)"
    SeedRow "[ScreenName] = 'frmCashClosing'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashClosing', 'تصفية يومية الكاشير', 'الخزينة', 240, 'CASH_CLOSING', True, False, False)"
    SeedRow "[ScreenName] = 'frmCashBoxes'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashBoxes', 'الصناديق', 'الخزينة', 250, 'CASH_BOX', True, True, True)"
    SeedRow "[ScreenName] = 'frmJournal'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmJournal', 'قيود اليومية', 'الحسابات', 260, 'JOURNAL', False, False, False)"
    SeedRow "[ScreenName] = 'frmAccounts'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAccounts', 'دليل الحسابات', 'الحسابات', 270, 'JOURNAL', True, True, True)"
    SeedRow "[ScreenName] = 'frmManualEntry'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmManualEntry', 'القيود اليدوية', 'الحسابات', 280, 'MANUAL_ENTRY', True, True, True)"
    SeedRow "[ScreenName] = 'frmLedger'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmLedger', 'كشف حساب ودفتر الأستاذ', 'الحسابات', 290, 'JOURNAL', False, False, False)"
    SeedRow "[ScreenName] = 'frmFinancials'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmFinancials', 'القوائم المالية', 'الحسابات', 300, 'REPORTS_PROFIT', False, False, False)"
    SeedRow "[ScreenName] = 'frmReportCenter'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmReportCenter', 'التقارير', 'التقارير', 310, 'REPORTS', False, False, False)"
    SeedRow "[ScreenName] = 'frmSearch'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSearch', 'البحث', 'النظام', 320, Null, False, False, False)"
    SeedRow "[ScreenName] = 'frmSettings'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSettings', 'إعدادات المحل', 'النظام', 330, 'SETTINGS', False, True, False)"
    SeedRow "[ScreenName] = 'frmUsers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUsers', 'المستخدمون', 'النظام', 340, 'USERS', True, True, False)"
    SeedRow "[ScreenName] = 'frmRoles'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmRoles', 'الأدوار والصلاحيات', 'النظام', 350, 'USERS', False, True, False)"
    SeedRow "[ScreenName] = 'frmUserScreens'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUserScreens', 'صلاحيات الشاشات للمستخدمين', 'النظام', 360, 'USERS', False, True, False)"
    SeedRow "[ScreenName] = 'frmBackup'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBackup', 'النسخ الاحتياطي', 'النظام', 370, 'BACKUP', False, False, False)"
    EndSeed "Screens", 37
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
    SeedRow "[AccountCode] = 1", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1, 'الأصول', 'ASSET', Null, False, True)"
    SeedRow "[AccountCode] = 11", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (11, 'الأصول المتداولة', 'ASSET', 1, False, True)"
    SeedRow "[AccountCode] = 1100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1100, 'النقدية بالخزينة والصناديق', 'ASSET', 11, False, True)"
    SeedRow "[AccountCode] = 110001", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110001, 'الخزينة الرئيسية', 'ASSET', 1100, True, True)"
    SeedRow "[AccountCode] = 110002", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110002, 'صندوق الكاشير', 'ASSET', 1100, True, True)"
    SeedRow "[AccountCode] = 1190", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1190, 'نقدية غير موزعة على صندوق', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1200, 'البنك والشبكة (مدى والتحويلات)', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1250", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1250, 'شيكات تحت التحصيل', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1300, 'ذمم العملاء', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1310", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1310, 'أوراق القبض', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1350", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1350, 'مخصص الديون المشكوك في تحصيلها', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1400, 'المخزون', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1500, 'ضريبة القيمة المضافة - مدخلات', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1600, 'سلف الموظفين', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1610", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1610, 'العهد النقدية', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1650", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1650, 'مصروفات مدفوعة مقدمًا', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1660", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1660, 'إيرادات مستحقة', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1690", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1690, 'تأمينات لدى الغير', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 12", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (12, 'الأصول غير المتداولة', 'ASSET', 1, False, True)"
    SeedRow "[AccountCode] = 1710", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1710, 'الأثاث والتجهيزات', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1720", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1720, 'الأجهزة والحاسبات وأنظمة نقاط البيع', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1730", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1730, 'السيارات ووسائل النقل', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1740", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1740, 'الديكورات وتحسينات المحل', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1790", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1790, 'مجمع إهلاك الأصول الثابتة', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1800, 'البرامج والتراخيص', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 2", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2, 'الخصوم', 'LIABILITY', Null, False, True)"
    SeedRow "[AccountCode] = 21", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (21, 'الخصوم المتداولة', 'LIABILITY', 2, False, True)"
    SeedRow "[AccountCode] = 2100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2100, 'ذمم الموردين', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2110", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2110, 'أوراق الدفع', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2200, 'ضريبة القيمة المضافة - مخرجات', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2250", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2250, 'ضريبة القيمة المضافة - التسوية والسداد', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2300, 'مصروفات مستحقة', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2310", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2310, 'رواتب وأجور مستحقة', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2320", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2320, 'التأمينات الاجتماعية المستحقة', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2400, 'دفعات مقدمة من العملاء', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2500, 'قروض قصيرة الأجل', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2600, 'الزكاة المستحقة', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 22", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (22, 'الخصوم غير المتداولة', 'LIABILITY', 2, False, True)"
    SeedRow "[AccountCode] = 2700", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2700, 'قروض طويلة الأجل', 'LIABILITY', 22, True, False)"
    SeedRow "[AccountCode] = 2800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2800, 'مخصص مكافأة نهاية الخدمة', 'LIABILITY', 22, True, False)"
    SeedRow "[AccountCode] = 3", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3, 'حقوق الملكية', 'EQUITY', Null, False, True)"
    SeedRow "[AccountCode] = 31", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (31, 'رأس المال وجاري المالك', 'EQUITY', 3, False, True)"
    SeedRow "[AccountCode] = 3200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3200, 'رأس المال', 'EQUITY', 31, True, False)"
    SeedRow "[AccountCode] = 3100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3100, 'جاري المالك', 'EQUITY', 31, True, True)"
    SeedRow "[AccountCode] = 3900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3900, 'أرصدة افتتاحية', 'EQUITY', 31, True, True)"
    SeedRow "[AccountCode] = 32", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (32, 'الأرباح المحتجزة ونتائج الأعمال', 'EQUITY', 3, False, True)"
    SeedRow "[AccountCode] = 3300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3300, 'الأرباح المحتجزة', 'EQUITY', 32, True, False)"
    SeedRow "[AccountCode] = 3400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3400, 'صافي ربح (خسارة) العام', 'EQUITY', 32, True, False)"
    SeedRow "[AccountCode] = 4", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4, 'الإيرادات', 'REVENUE', Null, False, True)"
    SeedRow "[AccountCode] = 41", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (41, 'إيرادات النشاط', 'REVENUE', 4, False, True)"
    SeedRow "[AccountCode] = 4100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4100, 'المبيعات', 'REVENUE', 41, True, True)"
    SeedRow "[AccountCode] = 4110", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4110, 'مردودات المبيعات', 'REVENUE', 41, True, True)"
    SeedRow "[AccountCode] = 4120", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4120, 'الخصم المسموح به', 'REVENUE', 41, True, False)"
    SeedRow "[AccountCode] = 42", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (42, 'إيرادات أخرى', 'REVENUE', 4, False, True)"
    SeedRow "[AccountCode] = 4200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4200, 'إيرادات متنوعة', 'REVENUE', 42, True, True)"
    SeedRow "[AccountCode] = 4300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4300, 'زيادة الصناديق', 'REVENUE', 42, True, True)"
    SeedRow "[AccountCode] = 4400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4400, 'الخصم المكتسب', 'REVENUE', 42, True, False)"
    SeedRow "[AccountCode] = 5", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5, 'المصروفات', 'EXPENSE', Null, False, True)"
    SeedRow "[AccountCode] = 51", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (51, 'تكلفة المبيعات', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5100, 'تكلفة البضاعة المباعة', 'EXPENSE', 51, True, True)"
    SeedRow "[AccountCode] = 5200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5200, 'فروقات وتسويات المخزون', 'EXPENSE', 51, True, True)"
    SeedRow "[AccountCode] = 52", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (52, 'المصروفات التشغيلية والإدارية', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5300, 'المصروفات حسب النوع', 'EXPENSE', 52, False, True)"
    SeedRow "[AccountCode] = 5500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5500, 'الرواتب والأجور', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5510", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5510, 'البدلات والحوافز', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5520", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5520, 'التأمينات الاجتماعية', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5600, 'إهلاك الأصول الثابتة', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5610", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5610, 'عمولات البنوك ونقاط البيع', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5620", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5620, 'الرسوم الحكومية والتراخيص', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5630", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5630, 'الدعاية والإعلان', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 53", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (53, 'مصروفات أخرى', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5400, 'عجز الصناديق', 'EXPENSE', 53, True, True)"
    SeedRow "[AccountCode] = 5700", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5700, 'الديون المعدومة', 'EXPENSE', 53, True, False)"
    SeedRow "[AccountCode] = 5800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5800, 'الزكاة', 'EXPENSE', 53, True, False)"
    SeedRow "[AccountCode] = 5900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5900, 'مصروفات متنوعة', 'EXPENSE', 53, True, True)"
    EndSeed "Accounts", 75
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
    SeedRow "[SourceType] = 'MANUAL'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('MANUAL', 'قيد يدوي', 14)"
    EndSeed "JournalSourceTypes", 14
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

Private Sub UpgradeAccountTree()
    ' accounts of an older back-end without a parent go to their place in the tree
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 1 WHERE [AccountCode] = 11 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 1100 WHERE [AccountCode] = 110001 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 1100 WHERE [AccountCode] = 110002 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1190 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1250 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1310 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1350 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1500 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1600 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1610 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1650 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1660 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1690 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 1 WHERE [AccountCode] = 12 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1710 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1720 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1730 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1740 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1790 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 12 WHERE [AccountCode] = 1800 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 2 WHERE [AccountCode] = 21 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2110 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2250 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2310 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2320 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2500 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2600 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 2 WHERE [AccountCode] = 22 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 22 WHERE [AccountCode] = 2700 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 22 WHERE [AccountCode] = 2800 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 3 WHERE [AccountCode] = 31 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 31 WHERE [AccountCode] = 3200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 31 WHERE [AccountCode] = 3100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 31 WHERE [AccountCode] = 3900 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 3 WHERE [AccountCode] = 32 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 32 WHERE [AccountCode] = 3300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 32 WHERE [AccountCode] = 3400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 4 WHERE [AccountCode] = 41 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 41 WHERE [AccountCode] = 4100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 41 WHERE [AccountCode] = 4110 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 41 WHERE [AccountCode] = 4120 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 4 WHERE [AccountCode] = 42 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 42 WHERE [AccountCode] = 4200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 42 WHERE [AccountCode] = 4300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 42 WHERE [AccountCode] = 4400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 51 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 51 WHERE [AccountCode] = 5100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 51 WHERE [AccountCode] = 5200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 52 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5500 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5510 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5520 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5600 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5610 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5620 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5630 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 53 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5700 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5800 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5900 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [IsPosting] = False WHERE [AccountCode] IN (1, 11, 1100, 12, 2, 21, 22, 3, 31, 32, 4, 41, 42, 5, 51, 52, 5300, 53)", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [IsSystem] = True WHERE [AccountCode] IN (1, 11, 1100, 110001, 110002, 1190, 1200, 1300, 1400, 1500, 1600, 12, 2, 21, 2100, 2200, 22, 3, 31, 3100, 3900, 32, 4, 41, 4100, 4110, 42, 4200, 4300, 5, 51, 5100, 5200, 52, 5300, 53, 5400, 5900) OR [AccountCode] BETWEEN 110001 AND 119999 OR [AccountCode] BETWEEN 530001 AND 539999", dbFailOnError
End Sub
