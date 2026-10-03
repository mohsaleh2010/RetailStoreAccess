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

Private Const SCHEMA_TABLES As String = "Settings,Sequences,Roles,Permissions,RolePermissions,Employees,Categories,Units,PaymentMethods,Suppliers,Customers,Products,SalesInvoices,SalesInvoiceDetails,SalesReturns,SalesReturnDetails,PurchaseInvoices,PurchaseInvoiceDetails,PurchaseReturns,PurchaseReturnDetails,CustomerPayments,SupplierPayments,ExpenseTypes,Expenses,TransactionTypes,InventoryTransactions,StockCounts,StockCountDetails,AuditLog"
Private Const EXPECTED_FIELD_COUNTS As String = "Settings=28;Sequences=5;Roles=4;Permissions=4;RolePermissions=2;Employees=16;Categories=4;Units=4;PaymentMethods=5;Suppliers=15;Customers=21;Products=19;SalesInvoices=29;SalesInvoiceDetails=13;SalesReturns=28;SalesReturnDetails=14;PurchaseInvoices=17;PurchaseInvoiceDetails=11;PurchaseReturns=17;PurchaseReturnDetails=11;CustomerPayments=10;SupplierPayments=10;ExpenseTypes=3;Expenses=12;TransactionTypes=5;InventoryTransactions=13;StockCounts=9;StockCountDetails=9;AuditLog=8"
Private Const EXPECTED_SEED_COUNTS As String = "Settings=1;Sequences=11;Roles=3;Permissions=22;RolePermissions=44;Employees=1;Categories=1;Units=8;PaymentMethods=4;Customers=1;ExpenseTypes=9;TransactionTypes=8"

Private m_db As DAO.Database
Private m_pending As Collection
Private m_log As String
Private m_created As Long
Private m_skipped As Long
Private m_seeded As Long
Private m_currentStep As String
Private m_inTrans As Boolean

'------------------------------------------------------------------------------
' Public entry points
'------------------------------------------------------------------------------
Public Function BuildSchema(Optional ByVal BackEndPath As String = "") As Boolean
    On Error GoTo EH
    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    If StrComp(BackEndPath, CurrentProject.FullName, vbTextCompare) = 0 Then
        MsgBox "„”«— „·› «·»Ì«‰«  ÌÃ» √‰ ÌŒ ·› ⁄‰ „·› «·Ê«ÃÂ… «·Õ«·Ì.", vbExclamation + MSG_RTL
        Exit Function
    End If

    m_log = "": m_created = 0: m_skipped = 0: m_seeded = 0
    LogLine "=== BuildSchema " & SCHEMA_VERSION & "  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="

    m_currentStep = "open back-end"
    If Len(Dir$(BackEndPath)) = 0 Then
        Set m_db = DBEngine.CreateDatabase(BackEndPath, dbLangArabic, DB_VERSION_120)
        LogLine " „ ≈‰‘«¡ „·› «·»Ì«‰« : " & BackEndPath
    Else
        Set m_db = DBEngine.OpenDatabase(BackEndPath)
        LogLine "„·› «·»Ì«‰«  „ÊÃÊœ: " & BackEndPath
    End If

    CreateAllTables
    SeedAll

    m_db.Close
    Set m_db = Nothing

    m_currentStep = "link tables"
    LinkBackEnd BackEndPath

    LogLine "--- Ãœ«Ê· ÃœÌœ…: " & m_created & " | „ÊÃÊœ… „”»ﬁ«: " & m_skipped & _
            " | Ãœ«Ê·  „   ⁄»∆ Â«: " & m_seeded
    MsgBox " „ »‰«¡ «·Ãœ«Ê· »‰Ã«Õ." & vbCrLf & vbCrLf & _
           "Ãœ«Ê· ÃœÌœ…: " & m_created & vbCrLf & _
           "Ãœ«Ê· „ÊÃÊœ… „”»ﬁ«: " & m_skipped & vbCrLf & _
           "Ãœ«Ê·  „   ⁄»∆… »Ì«‰« Â« «·√”«”Ì…: " & m_seeded & vbCrLf & vbCrLf & _
           "«· ›«’Ì· ›Ì ‰«›–… Immediate (Ctrl+G)." & vbCrLf & _
           "«·ŒÿÊ… «· «·Ì…: ‘€¯· VerifySchema", vbInformation + MSG_RTL, "BuildSchema"
    BuildSchema = True
    Exit Function

EH:
    Dim errText As String
    errText = "Œÿ√ " & Err.Number & " √À‰«¡ [" & m_currentStep & "]: " & Err.Description
    If m_inTrans Then
        DBEngine.Workspaces(0).Rollback
        m_inTrans = False
    End If
    LogLine errText
    On Error Resume Next
    If Not m_db Is Nothing Then m_db.Close
    Set m_db = Nothing
    MsgBox errText & vbCrLf & vbCrLf & "Ì„ﬂ‰ ≈⁄«œ…  ‘€Ì· BuildSchema »⁄œ «·≈’·«Õ∫ " & _
           "«·Ãœ«Ê· «· Ì √ı‰‘∆  ·‰   ﬂ——.", vbCritical + MSG_RTL, "BuildSchema"
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
                    LogLine " ‰»ÌÂ: ÌÊÃœ ÃœÊ· „Õ·Ì »‰›” «·«”„ ›Ì «·Ê«ÃÂ… Ê·„ Ì „ —»ÿÂ: " & tdfBE.Name
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
    LogLine "«·—»ÿ: " & linked & " ÃœÊ· ÃœÌœ° " & refreshed & " ÃœÊ·  „  ÕœÌÀ —»ÿÂ."
End Sub

Public Function VerifySchema(Optional ByVal BackEndPath As String = "") As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim items() As String, parts() As String, i As Long
    Dim problems As Long, report As String, s As String

    If Len(BackEndPath) = 0 Then BackEndPath = DefaultBackEndPath()
    If Len(Dir$(BackEndPath)) = 0 Then
        MsgBox "„·› «·»Ì«‰«  €Ì— „ÊÃÊœ: " & BackEndPath, vbCritical + MSG_RTL
        Exit Function
    End If
    Set db = DBEngine.OpenDatabase(BackEndPath, False, True)

    ' 1) every table exists with the expected number of fields
    items = Split(EXPECTED_FIELD_COUNTS, ";")
    For i = 0 To UBound(items)
        parts = Split(items(i), "=")
        If Not TableExistsIn(db, parts(0)) Then
            s = "[X] «·ÃœÊ· €Ì— „ÊÃÊœ: " & parts(0)
            problems = problems + 1
        ElseIf db.TableDefs(parts(0)).Fields.Count <> CLng(parts(1)) Then
            s = "[X] " & parts(0) & ": ⁄œœ «·ÕﬁÊ· " & db.TableDefs(parts(0)).Fields.Count & _
                " Ê«·„ Êﬁ⁄ " & parts(1)
            problems = problems + 1
        Else
            s = "[OK] " & parts(0) & " (" & parts(1) & " Õﬁ·)"
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
                s = "[X] " & parts(0) & ": " & rs(0) & " ”Ã· Ê«·„ Êﬁ⁄ " & parts(1) & " ⁄·Ï «·√ﬁ·"
                problems = problems + 1
                report = report & s & vbCrLf
            Else
                s = "[OK] »Ì«‰«  " & parts(0) & ": " & rs(0) & " ”Ã·"
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
                s = "[X] «·‰’ «·⁄—»Ì „Õ›ÊŸ »‘ﬂ· Œ«ÿ∆ (" & rs!RoleName & ")." & vbCrLf & _
                    "    «÷»ÿ Windows > Region > Administrative > Language for non-Unicode programs = Arabic" & _
                    " À„ √⁄œ «·«” Ì—«œ Ê«·»‰«¡."
                problems = problems + 1
                report = report & s & vbCrLf
            Else
                s = "[OK] «·‰’ «·⁄—»Ì ”·Ì„: " & rs!RoleName
            End If
            Debug.Print s
        End If
        rs.Close
    End If

    db.Close
    If problems = 0 Then
        MsgBox "«·›Õ’ ‰«ÃÕ: Ã„Ì⁄ «·Ãœ«Ê· (" & (UBound(Split(SCHEMA_TABLES, ",")) + 1) & _
               ") Ê«·»Ì«‰«  «·√”«”Ì… ”·Ì„….", vbInformation + MSG_RTL, "VerifySchema"
        VerifySchema = True
    Else
        MsgBox "⁄œœ «·„‘ﬂ·« : " & problems & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
               "VerifySchema"
    End If
End Function

Public Sub DropSchema(Optional ByVal BackEndPath As String = "")
    ' DEVELOPMENT ONLY - deletes every table of this system and all of its data.
    Dim db As DAO.Database, dbFE As DAO.Database, i As Long, names() As String

    If InputBox("”Ì „ Õ–› Ã„Ì⁄ Ãœ«Ê· «·‰Ÿ«„ Ê»Ì«‰« Â« ‰Â«∆Ì«." & vbCrLf & _
                "·· √ﬂÌœ «ﬂ » DELETE", "DropSchema") <> "DELETE" Then Exit Sub
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
    MsgBox " „ Õ–› Ãœ«Ê· «·‰Ÿ«„.", vbInformation + MSG_RTL
End Sub

'------------------------------------------------------------------------------
' Table building helpers
'------------------------------------------------------------------------------
Private Function BeginTable(ByRef tdf As DAO.TableDef, ByVal TableName As String) As Boolean
    m_currentStep = "create table " & TableName
    If TableExistsIn(m_db, TableName) Then
        LogLine "  = „ÊÃÊœ „”»ﬁ«: " & TableName
        m_skipped = m_skipped + 1
        Exit Function
    End If
    Set tdf = m_db.CreateTableDef(TableName)
    Set m_pending = New Collection
    BeginTable = True
End Function

Private Sub AddField(ByVal tdf As DAO.TableDef, ByVal FieldName As String, ByVal Kind As String, _
                     ByVal Size As Long, ByVal IsRequired As Boolean, ByVal DefaultValue As String, _
                     ByVal ValidationRule As String, ByVal ValidationText As String, _
                     ByVal Caption As String, ByVal Description As String)
    Dim fld As DAO.Field
    m_currentStep = "field " & tdf.Name & "." & FieldName

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
    LogLine "  +  „ ≈‰‘«¡ «·ÃœÊ·: " & tdf.Name
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
Private Function BeginSeed(ByVal TableName As String) As Boolean
    Dim rs As DAO.Recordset
    m_currentStep = "seed " & TableName
    Set rs = m_db.OpenRecordset("SELECT COUNT(*) FROM [" & TableName & "]", dbOpenSnapshot)
    If rs(0) > 0 Then
        rs.Close
        Exit Function
    End If
    rs.Close
    DBEngine.Workspaces(0).BeginTrans
    m_inTrans = True
    BeginSeed = True
End Function

Private Sub ExecSeed(ByVal Sql As String)
    m_db.Execute Sql, dbFailOnError
End Sub

Private Sub EndSeed(ByVal TableName As String, ByVal RowCount As Long)
    DBEngine.Workspaces(0).CommitTrans
    m_inTrans = False
    m_seeded = m_seeded + 1
    LogLine "  * »Ì«‰«  √”«”Ì…: " & TableName & " (" & RowCount & " ”Ã·)"
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
    CreateTable_TransactionTypes
    CreateTable_InventoryTransactions
    CreateTable_StockCounts
    CreateTable_StockCountDetails
    CreateTable_AuditLog
End Sub

Private Sub CreateTable_Settings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Settings") Then Exit Sub
    AddField tdf, "SettingID", "LONG", 0, True, "1", _
             "=1", "Ì”„Õ »”Ã· ≈⁄œ«œ«  Ê«Õœ ›ﬁÿ", "—ﬁ„ «·≈⁄œ«œ", ""
    AddField tdf, "StoreName", "TEXT", 150, True, "", _
             "", "", "«”„ «·„Õ·", ""
    AddField tdf, "StoreNameEn", "TEXT", 150, False, "", _
             "", "", "«”„ «·„Õ· »«·≈‰Ã·Ì“Ì…", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "«·—ﬁ„ «·÷—Ì»Ì 15 —ﬁ„« ÊÌ»œ√ ÊÌ‰ ÂÌ »«·—ﬁ„ 3", "«·—ﬁ„ «·÷—Ì»Ì", ""
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "«·”Ã· «· Ã«—Ì", ""
    AddField tdf, "BuildingNo", "TEXT", 10, False, "", _
             "", "", "—ﬁ„ «·„»‰Ï", ""
    AddField tdf, "StreetName", "TEXT", 100, False, "", _
             "", "", "«·‘«—⁄", ""
    AddField tdf, "District", "TEXT", 100, False, "", _
             "", "", "«·ÕÌ", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "«·„œÌ‰…", ""
    AddField tdf, "PostalCode", "TEXT", 10, False, "", _
             "", "", "«·—„“ «·»—ÌœÌ", ""
    AddField tdf, "AdditionalNo", "TEXT", 10, False, "", _
             "", "", "«·—ﬁ„ «·≈÷«›Ì", ""
    AddField tdf, "CountryCode", "TEXT", 2, True, """SA""", _
             "", "", "—„“ «·œÊ·…", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "«·Â« ›", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·÷—Ì»…", ""
    AddField tdf, "PricesIncludeVAT", "BOOL", 0, False, "True", _
             "", "", "«·√”⁄«— ‘«„·… «·÷—Ì»…", ""
    AddField tdf, "AllowNegativeStock", "BOOL", 0, False, "False", _
             "", "", "«·”„«Õ »«·»Ì⁄ »«·”«·»", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, True, """SAR""", _
             "", "", "«·⁄„·…", ""
    AddField tdf, "DefaultCustomerID", "LONG", 0, True, "1", _
             "", "", "«·⁄„Ì· «·«› —«÷Ì", ""
    AddField tdf, "ZatcaPhase", "BYTE", 0, True, "1", _
             "In (1,2)", "«·„—Õ·… 1 √Ê 2", "„—Õ·… ›« Ê—…", ""
    AddField tdf, "ZatcaEnvironment", "TEXT", 20, False, "", _
             "", "", "»Ì∆… «·—»ÿ „⁄ «·ÂÌ∆…", ""
    AddField tdf, "LastInvoiceHash", "TEXT", 255, False, "", _
             "", "", "»’„… ¬Œ— „” ‰œ", " »œ√ »«·ﬁÌ„… «·«› —«÷Ì… «· Ì  ÕœœÂ« «·ÂÌ∆… ·√Ê· ›« Ê—…"
    AddField tdf, "BackupFolder", "TEXT", 255, False, "", _
             "", "", "„Ã·œ «·‰”Œ «·«Õ Ì«ÿÌ", ""
    AddField tdf, "BackupKeepCount", "INT", 0, True, "30", _
             ">=1", "«Õ ›Ÿ »‰”Œ… Ê«Õœ… ⁄·Ï «·√ﬁ·", "⁄œœ «·‰”Œ «·„Õ ›Ÿ »Â«", ""
    AddField tdf, "SlowMovingDays", "INT", 0, True, "90", _
             ">=1", "⁄œœ «·√Ì«„ ÌÃ» √‰ ÌﬂÊ‰ 1 √Ê √ﬂÀ—", "√Ì«„ ⁄œ„ «·Õ—ﬂ…", ""
    AddField tdf, "ReceiptFooter", "TEXT", 255, False, "", _
             "", "", " –ÌÌ· «·›« Ê—…", ""
    AddField tdf, "LogoPath", "TEXT", 255, False, "", _
             "", "", "„”«— «·‘⁄«—", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "¬Œ—  ⁄œÌ·", ""
    AddIndex tdf, "PrimaryKey", "SettingID", True, True, False
    EndTable tdf, "≈⁄œ«œ«  «·„Õ·: ”Ã· Ê«Õœ ›ﬁÿ ÌÕ ÊÌ »Ì«‰«  «·„Õ· «·÷—Ì»Ì… Ê≈⁄œ«œ«  «· ‘€Ì·.", "", ""
End Sub

Private Sub CreateTable_Sequences()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Sequences") Then Exit Sub
    AddField tdf, "SequenceName", "TEXT", 30, True, "", _
             "", "", "«”„ «·⁄œ¯«œ", ""
    AddField tdf, "Prefix", "TEXT", 10, False, "", _
             "", "", "«·»«œ∆…", ""
    AddField tdf, "NextValue", "LONG", 0, True, "1", _
             ">=1", "«·—ﬁ„ «· «·Ì ÌÃ» √‰ ÌﬂÊ‰ 1 √Ê √ﬂÀ—", "«·—ﬁ„ «· «·Ì", ""
    AddField tdf, "PadLength", "BYTE", 0, True, "6", _
             "Between 0 And 12", "„‰ 0 ≈·Ï 12 Œ«‰…", "⁄œœ «·Œ«‰« ", ""
    AddField tdf, "Description", "TEXT", 100, False, "", _
             "", "", "«·Ê’›", ""
    AddIndex tdf, "PrimaryKey", "SequenceName", True, True, False
    EndTable tdf, "⁄œ¯«œ«  «· —ﬁÌ„: ÌÊ·¯œ √—ﬁ«„«  ”·”·Ì… »œÊ‰ ›ÃÊ«  ··›Ê« Ì— Ê«·”‰œ«  (Ìı“«œ œ«Œ· ‰›” „⁄«„·… «·Õ›Ÿ).", "", ""
End Sub

Private Sub CreateTable_Roles()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Roles") Then Exit Sub
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "—ﬁ„ «·œÊ—", ""
    AddField tdf, "RoleCode", "TEXT", 20, True, "", _
             "", "", "—„“ «·œÊ—", ""
    AddField tdf, "RoleName", "TEXT", 50, True, "", _
             "", "", "«”„ «·œÊ—", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·Ê’›", ""
    AddIndex tdf, "PrimaryKey", "RoleID", True, True, False
    AddIndex tdf, "UX_RoleCode", "RoleCode", False, True, False
    EndTable tdf, "«·√œÊ«—: „” ÊÌ«  «·’·«ÕÌ« : „œÌ— «·‰Ÿ«„° „œÌ—° ﬂ«‘Ì—.", "", ""
End Sub

Private Sub CreateTable_Permissions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Permissions") Then Exit Sub
    AddField tdf, "PermissionKey", "TEXT", 50, True, "", _
             "", "", "—„“ «·’·«ÕÌ…", ""
    AddField tdf, "PermissionName", "TEXT", 100, True, "", _
             "", "", "«”„ «·’·«ÕÌ…", ""
    AddField tdf, "ModuleName", "TEXT", 50, False, "", _
             "", "", "«·ﬁ”„", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddIndex tdf, "PrimaryKey", "PermissionKey", True, True, False
    EndTable tdf, "«·’·«ÕÌ« : ﬁ«∆„… «·’·«ÕÌ«  «· Ì Ì„ﬂ‰ „‰ÕÂ« ··√œÊ«—.", "", ""
End Sub

Private Sub CreateTable_RolePermissions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "RolePermissions") Then Exit Sub
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "«·œÊ—", ""
    AddField tdf, "PermissionKey", "TEXT", 50, True, "", _
             "", "", "«·’·«ÕÌ…", ""
    AddIndex tdf, "PrimaryKey", "RoleID,PermissionKey", True, True, False
    EndTable tdf, "’·«ÕÌ«  «·√œÊ«—: —»ÿ ﬂ· œÊ— »«·’·«ÕÌ«  «·„„‰ÊÕ… ·Â (⁄·«ﬁ… „ ⁄œœ ·„ ⁄œœ).", "", ""
End Sub

Private Sub CreateTable_Employees()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Employees") Then Exit Sub
    AddField tdf, "EmployeeID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·„ÊŸ›", ""
    AddField tdf, "EmployeeName", "TEXT", 100, True, "", _
             "", "", "«”„ «·„ÊŸ›", ""
    AddField tdf, "JobTitle", "TEXT", 50, False, "", _
             "", "", "«·„”„Ï «·ÊŸÌ›Ì", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "«·ÃÊ«·", ""
    AddField tdf, "Username", "TEXT", 30, True, "", _
             "", "", "«”„ «·„” Œœ„", ""
    AddField tdf, "PasswordHash", "TEXT", 64, False, "", _
             "", "", "»’„… ﬂ·„… «·„—Ê—", "SHA-256 »’Ì€… hex∫ ﬂ·„… «·„—Ê— ‰›”Â« ·«  ıÕ›Ÿ √»œ«"
    AddField tdf, "PasswordSalt", "TEXT", 32, False, "", _
             "", "", "„·Õ «· ‘›Ì—", ""
    AddField tdf, "RoleID", "LONG", 0, True, "", _
             "", "", "«·œÊ—", ""
    AddField tdf, "MaxDiscountPercent", "RATE", 0, True, "0", _
             ">=0 And <=1", "«·‰”»… »Ì‰ 0% Ê 100%", "√ﬁ’Ï ‰”»… Œ’„", ""
    AddField tdf, "MustChangePassword", "BOOL", 0, False, "True", _
             "", "", "ÌÃ»  €ÌÌ— ﬂ·„… «·„—Ê—", ""
    AddField tdf, "FailedLoginCount", "INT", 0, True, "0", _
             "", "", "„Õ«Ê·«  «·œŒÊ· «·›«‘·…", ""
    AddField tdf, "LockedUntil", "DATETIME", 0, False, "", _
             "", "", "„ﬁ›· Õ Ï", ""
    AddField tdf, "LastLoginAt", "DATETIME", 0, False, "", _
             "", "", "¬Œ— œŒÊ·", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "EmployeeID", True, True, False
    AddIndex tdf, "UX_Username", "Username", False, True, False
    EndTable tdf, "«·„ÊŸ›Ê‰ Ê«·„” Œœ„Ê‰: ﬂ· „ÊŸ› ÂÊ „” Œœ„ ··‰Ÿ«„∫ ·« ÌıÕ–› »· Ìı⁄ÿÛ¯· ··Õ›«Ÿ ⁄·Ï ”Ã· ⁄„·Ì« Â.", "", ""
End Sub

Private Sub CreateTable_Categories()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Categories") Then Exit Sub
    AddField tdf, "CategoryID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «· ’‰Ì›", ""
    AddField tdf, "CategoryName", "TEXT", 100, True, "", _
             "", "", "«”„ «· ’‰Ì›", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·Ê’›", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddIndex tdf, "PrimaryKey", "CategoryID", True, True, False
    AddIndex tdf, "UX_CategoryName", "CategoryName", False, True, False
    EndTable tdf, "«· ’‰Ì›« :  ’‰Ì›«  «·„‰ Ã«  (≈·ﬂ —Ê‰Ì« ° „Ê«œ €–«∆Ì…° ...).", "", ""
End Sub

Private Sub CreateTable_Units()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Units") Then Exit Sub
    AddField tdf, "UnitID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·ÊÕœ…", ""
    AddField tdf, "UnitName", "TEXT", 30, True, "", _
             "", "", "«”„ «·ÊÕœ…", ""
    AddField tdf, "ZatcaUnitCode", "TEXT", 10, False, "", _
             "", "", "—„“ «·ÊÕœ… (UN/ECE)", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddIndex tdf, "PrimaryKey", "UnitID", True, True, False
    AddIndex tdf, "UX_UnitName", "UnitName", False, True, False
    EndTable tdf, "ÊÕœ«  «·ﬁÌ«”: ÊÕœ«  «·»Ì⁄ (Õ»…° ﬂ— Ê‰° ﬂÌ·Ê° ...) „⁄ —„“Â« ›Ì ›« Ê—… «·ÂÌ∆….", "", ""
End Sub

Private Sub CreateTable_PaymentMethods()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PaymentMethods") Then Exit Sub
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "", _
             "", "", "—ﬁ„ «·ÿ—Ìﬁ…", ""
    AddField tdf, "MethodName", "TEXT", 50, True, "", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "ZatcaCode", "TEXT", 5, False, "", _
             "", "", "—„“ «·ÂÌ∆…", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddIndex tdf, "PrimaryKey", "PaymentMethodID", True, True, False
    AddIndex tdf, "UX_MethodName", "MethodName", False, True, False
    EndTable tdf, "ÿ—ﬁ «·œ›⁄: ÿ—ﬁ œ›⁄ «·„»«·€ «·„”œœ… „⁄ —„“Â« ›Ì ›« Ê—… «·ÂÌ∆… (UNTDID 4461).", "", ""
End Sub

Private Sub CreateTable_Suppliers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Suppliers") Then Exit Sub
    AddField tdf, "SupplierID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·„Ê—œ", ""
    AddField tdf, "SupplierName", "TEXT", 150, True, "", _
             "", "", "«”„ «·„Ê—œ", ""
    AddField tdf, "ContactPerson", "TEXT", 100, False, "", _
             "", "", "«·‘Œ’ «·„”ƒÊ·", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "«·ÃÊ«·", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "«·Â« ›", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "«·—ﬁ„ «·÷—Ì»Ì 15 —ﬁ„« ÊÌ»œ√ ÊÌ‰ ÂÌ »«·—ﬁ„ 3", "«·—ﬁ„ «·÷—Ì»Ì", ""
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "«·”Ã· «· Ã«—Ì", ""
    AddField tdf, "Address", "TEXT", 255, False, "", _
             "", "", "«·⁄‰Ê«‰", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "«·„œÌ‰…", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·«›  «ÕÌ", "„ÊÃ» = «·„Õ· „œÌ‰ ··„Ê—œ"
    AddField tdf, "CurrentBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·Õ«·Ì", "ﬁÌ„… „”«⁄œ…∫ «·„—Ã⁄ ÂÊ SupplierBalanceQuery"
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "SupplierID", True, True, False
    AddIndex tdf, "IX_SupplierName", "SupplierName", False, False, False
    AddIndex tdf, "IX_Mobile", "Mobile", False, False, False
    EndTable tdf, "«·„Ê—œÊ‰: »Ì«‰«  «·„Ê—œÌ‰. «·—’Ìœ «·Õ«·Ì ﬁÌ„… „”«⁄œ…  ıÕœÛ¯À »«·ﬂÊœ ÊÌ„ﬂ‰ ≈⁄«œ… «Õ ”«»Â« „‰ «·Õ—ﬂ« .", "", ""
End Sub

Private Sub CreateTable_Customers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Customers") Then Exit Sub
    AddField tdf, "CustomerID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·⁄„Ì·", ""
    AddField tdf, "CustomerName", "TEXT", 150, True, "", _
             "", "", "«”„ «·⁄„Ì·", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "«·ÃÊ«·", ""
    AddField tdf, "Phone", "TEXT", 20, False, "", _
             "", "", "«·Â« ›", ""
    AddField tdf, "Email", "TEXT", 100, False, "", _
             "", "", "«·»—Ìœ «·≈·ﬂ —Ê‰Ì", ""
    AddField tdf, "VATNumber", "TEXT", 15, False, "", _
             "Is Null Or Like ""3#############3""", "«·—ﬁ„ «·÷—Ì»Ì 15 —ﬁ„« ÊÌ»œ√ ÊÌ‰ ÂÌ »«·—ﬁ„ 3", "«·—ﬁ„ «·÷—Ì»Ì", "≈–« ÊıÃœ  ’œ— ··⁄„Ì· ›« Ê—… ÷—Ì»Ì… B2B"
    AddField tdf, "CRNumber", "TEXT", 20, False, "", _
             "", "", "«·”Ã· «· Ã«—Ì", ""
    AddField tdf, "BuildingNo", "TEXT", 10, False, "", _
             "", "", "—ﬁ„ «·„»‰Ï", ""
    AddField tdf, "StreetName", "TEXT", 100, False, "", _
             "", "", "«·‘«—⁄", ""
    AddField tdf, "District", "TEXT", 100, False, "", _
             "", "", "«·ÕÌ", ""
    AddField tdf, "City", "TEXT", 50, False, "", _
             "", "", "«·„œÌ‰…", ""
    AddField tdf, "PostalCode", "TEXT", 10, False, "", _
             "", "", "«·—„“ «·»—ÌœÌ", ""
    AddField tdf, "Address", "TEXT", 255, False, "", _
             "", "", "«·⁄‰Ê«‰", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·«›  «ÕÌ", "„ÊÃ» = «·⁄„Ì· „œÌ‰ ··„Õ·"
    AddField tdf, "CurrentBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·Õ«·Ì", "ﬁÌ„… „”«⁄œ…∫ «·„—Ã⁄ ÂÊ CustomerBalanceQuery"
    AddField tdf, "AllowCredit", "BOOL", 0, False, "True", _
             "", "", "Ì”„Õ »«·»Ì⁄ «·¬Ã·", ""
    AddField tdf, "CreditLimit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "Õœ «·«∆ „«‰", "0 = »œÊ‰ Õœ"
    AddField tdf, "IsSystem", "BOOL", 0, False, "False", _
             "", "", "”Ã· ‰Ÿ«„", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "CustomerID", True, True, False
    AddIndex tdf, "IX_CustomerName", "CustomerName", False, False, False
    AddIndex tdf, "IX_Mobile", "Mobile", False, False, False
    EndTable tdf, "«·⁄„·«¡: »Ì«‰«  «·⁄„·«¡. «·”Ã· —ﬁ„ 1 ÂÊ ""⁄„Ì· ‰ﬁœÌ"" «·«› —«÷Ì Ê·« ÌıÕ–›.", "", ""
End Sub

Private Sub CreateTable_Products()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Products") Then Exit Sub
    AddField tdf, "ProductID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·„‰ Ã", ""
    AddField tdf, "ProductCode", "TEXT", 30, True, "", _
             "", "", "ﬂÊœ «·„‰ Ã", ""
    AddField tdf, "Barcode", "TEXT", 50, False, "", _
             "", "", "«·»«—ﬂÊœ", ""
    AddField tdf, "ProductName", "TEXT", 150, True, "", _
             "", "", "«”„ «·„‰ Ã", ""
    AddField tdf, "ProductNameEn", "TEXT", 150, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", ""
    AddField tdf, "CategoryID", "LONG", 0, True, "1", _
             "", "", "«· ’‰Ì›", ""
    AddField tdf, "UnitID", "LONG", 0, True, "1", _
             "", "", "«·ÊÕœ…", ""
    AddField tdf, "PurchasePrice", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "¬Œ— ”⁄— ‘—«¡", "»œÊ‰ ÷—Ì»…"
    AddField tdf, "AverageCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„ Ê”ÿ «· ﬂ·›…", "«·„ Ê”ÿ «·„—Ã¯Õ° ÌıÕœÛ¯À „⁄ ﬂ· ‘—«¡"
    AddField tdf, "SellingPrice", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "”⁄— «·»Ì⁄", "‘«„· «·÷—Ì»… ≈–« ﬂ«‰ «·≈⁄œ«œ PricesIncludeVAT „›⁄¯·«"
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = Œ«÷⁄ 15%° Z = ‰”»… ’›—Ì…° E = „⁄›Ï", "«·›∆… «·÷—Ì»Ì…", ""
    AddField tdf, "CurrentQuantity", "QTY", 0, True, "0", _
             "", "", "«·ﬂ„Ì… «·Õ«·Ì…", ""
    AddField tdf, "MinimumQuantity", "QTY", 0, True, "0", _
             ">=0", "«·Õœ «·√œ‰Ï ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "Õœ ≈⁄«œ… «·ÿ·»", ""
    AddField tdf, "SupplierID", "LONG", 0, False, "", _
             "", "", "«·„Ê—œ «·«› —«÷Ì", ""
    AddField tdf, "ProductLocation", "TEXT", 50, False, "", _
             "", "", "„ﬂ«‰ «·„‰ Ã", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "¬Œ—  ⁄œÌ·", ""
    AddIndex tdf, "PrimaryKey", "ProductID", True, True, False
    AddIndex tdf, "UX_ProductCode", "ProductCode", False, True, False
    AddIndex tdf, "UX_Barcode", "Barcode", False, True, True
    AddIndex tdf, "IX_ProductName", "ProductName", False, False, False
    EndTable tdf, "«·„‰ Ã« : «·√’‰«› Ê√”⁄«—Â«. CurrentQuantity ﬁÌ„… „”«⁄œ…∫ «·„—Ã⁄ ÂÊ „Ã„Ê⁄ InventoryTransactions.", "", ""
End Sub

Private Sub CreateTable_SalesInvoices()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesInvoices") Then Exit Sub
    AddField tdf, "SalesInvoiceID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "InvoiceNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·›« Ê—…", ""
    AddField tdf, "InvoiceDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ ÊÊﬁ  «·›« Ê—…", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "1", _
             "", "", "«·⁄„Ì·", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·ﬂ«‘Ì—", ""
    AddField tdf, "PaymentType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = ‰ﬁœÌ° CREDIT = ¬Ã·", "‰Ê⁄ «·»Ì⁄", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„", "„Ã„Ê⁄ («·ﬂ„Ì… ◊ ”⁄— «·ÊÕœ…) »œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", "„Ã„Ê⁄ Œ’Ê„«  «·√”ÿ— (Œ’„ «·›« Ê—… ÌÊ“Û¯⁄ ⁄·Ï «·√”ÿ—)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ«÷⁄ ··÷—Ì»…", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", "„Ã„Ê⁄ ÷—Ì»… «·√”ÿ—"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", "TaxableAmount + Tax"
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„œ›Ê⁄", "«·„»·€ «·„Õ ”» „‰ «·›« Ê—… (·« Ì Ã«Ê“ «·≈Ã„«·Ì)"
    AddField tdf, "RemainingAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„ »ﬁÌ", "Ìı÷«› ≈·Ï —’Ìœ «·⁄„Ì· ›Ì «·»Ì⁄ «·¬Ã·"
    AddField tdf, "AmountTendered", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„” ·„", "„« ”·¯„Â «·⁄„Ì· ‰ﬁœ«"
    AddField tdf, "ChangeDue", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·»«ﬁÌ ··⁄„Ì·", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "InvoiceSubType", "TEXT", 10, True, """SIMPLIFIED""", _
             "In (""SIMPLIFIED"",""STANDARD"")", "SIMPLIFIED = ›« Ê—… „»”ÿ…° STANDARD = ›« Ê—… ÷—Ì»Ì…", "‰Ê⁄ «·›« Ê—… «·÷—Ì»Ì…", "„»”ÿ… ··√›—«œ B2C° ÷—Ì»Ì… ··„‰‘¬  B2B (··⁄„Ì· —ﬁ„ ÷—Ì»Ì)"
    AddField tdf, "InvoiceTypeCode", "TEXT", 3, True, """388""", _
             "In (""388"",""381"",""383"")", "388 ›« Ê—…° 381 ≈‘⁄«— œ«∆‰° 383 ≈‘⁄«— „œÌ‰", "—„“ ‰Ê⁄ «·„” ‰œ", ""
    AddField tdf, "InvoiceUUID", "TEXT", 36, False, "", _
             "", "", "«·„⁄—¯› «·›—Ìœ UUID", ""
    AddField tdf, "ICV", "LONG", 0, False, "", _
             "", "", "⁄œ¯«œ «·›Ê« Ì— ICV", " ”·”· „‘ —ﬂ ·ﬂ· «·„” ‰œ«  «·„—”·… ··ÂÌ∆…"
    AddField tdf, "InvoiceHash", "TEXT", 255, False, "", _
             "", "", "»’„… «·„” ‰œ", ""
    AddField tdf, "PreviousInvoiceHash", "TEXT", 255, False, "", _
             "", "", "»’„… «·„” ‰œ «·”«»ﬁ", ""
    AddField tdf, "QRCodeData", "MEMO", 0, False, "", _
             "", "", "»Ì«‰«  —„“ QR", ""
    AddField tdf, "ZatcaStatus", "TEXT", 20, True, """NOT_SENT""", _
             "In (""NOT_SENT"",""PENDING"",""REPORTED"",""CLEARED"",""WARNING"",""REJECTED"")", "Õ«·… €Ì— „⁄—Ê›…", "Õ«·… «·≈—”«· ··ÂÌ∆…", ""
    AddField tdf, "ZatcaSubmittedAt", "DATETIME", 0, False, "", _
             "", "", " «—ÌŒ «·≈—”«· ··ÂÌ∆…", ""
    AddField tdf, "ZatcaResponse", "MEMO", 0, False, "", _
             "", "", "—œ «·ÂÌ∆…", ""
    AddField tdf, "SignedXmlPath", "TEXT", 255, False, "", _
             "", "", "„”«— „·› XML «·„Êﬁ¯⁄", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "SalesInvoiceID", True, True, False
    AddIndex tdf, "UX_InvoiceNumber", "InvoiceNumber", False, True, False
    AddIndex tdf, "IX_InvoiceDate", "InvoiceDate", False, False, False
    AddIndex tdf, "UX_InvoiceUUID", "InvoiceUUID", False, True, True
    AddIndex tdf, "IX_ZatcaStatus", "ZatcaStatus", False, False, False
    EndTable tdf, "›Ê« Ì— «·„»Ì⁄« : —√” ›« Ê—… «·»Ì⁄. ·«  ıÕ–› Ê·«  ı⁄œÛ¯· »⁄œ «·Õ›Ÿ∫ «· ’ÕÌÕ »„— Ã⁄.", "[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]", "«·„œ›Ê⁄ ·« Ì Ã«Ê“ «·≈Ã„«·Ì° Ê«·„ »ﬁÌ = «·≈Ã„«·Ì - «·„œ›Ê⁄"
End Sub

Private Sub CreateTable_SalesInvoiceDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesInvoiceDetails") Then Exit Sub
    AddField tdf, "SalesDetailID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, True, "", _
             "", "", "«·›« Ê—…", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ—", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—", "«·ﬂ„Ì…", ""
    AddField tdf, "UnitPrice", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "”⁄— «·ÊÕœ…", "»œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", "ﬁÌ„… «·Œ’„ ⁄·Ï «·”ÿ— »œÊ‰ ÷—Ì»…"
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·’«›Ì ﬁ»· «·÷—Ì»…", "Quantity ◊ UnitPrice - Discount"
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = Œ«÷⁄ 15%° Z = ‰”»… ’›—Ì…° E = „⁄›Ï", "«·›∆… «·÷—Ì»Ì…", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·÷—Ì»…", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", "NetAmount + Tax"
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", "AverageCost ·ÕŸ… «·»Ì⁄"
    AddIndex tdf, "PrimaryKey", "SalesDetailID", True, True, False
    AddIndex tdf, "UX_SalesInvoiceID_LineNumber", "SalesInvoiceID,LineNumber", False, True, False
    EndTable tdf, " ›«’Ì· ›Ê« Ì— «·„»Ì⁄« : √”ÿ— ›« Ê—… «·»Ì⁄° „⁄ Õ›Ÿ  ﬂ·›… «·’‰› ·ÕŸ… «·»Ì⁄ ·Õ”«» «·—»Õ »œﬁ….", "", ""
End Sub

Private Sub CreateTable_SalesReturns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesReturns") Then Exit Sub
    AddField tdf, "SalesReturnID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ReturnNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·„— Ã⁄", ""
    AddField tdf, "ReturnDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·„— Ã⁄", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, True, "", _
             "", "", "«·›« Ê—… «·√’·Ì…", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "", _
             "", "", "«·⁄„Ì·", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "Reason", "TEXT", 255, True, "", _
             "", "", "”»» «·≈—Ã«⁄", "≈·“«„Ì ›Ì «·≈‘⁄«— «·œ«∆‰ Õ”» „ ÿ·»«  «·ÂÌ∆…"
    AddField tdf, "RefundType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = —œ ‰ﬁœÌ° CREDIT = Œ’„ „‰ —’Ìœ «·⁄„Ì·", "ÿ—Ìﬁ… —œ «·„»·€", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "ÿ—Ìﬁ… «·—œ", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„", "„Ã„Ê⁄ («·ﬂ„Ì… ◊ ”⁄— «·ÊÕœ…) »œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", "„Ã„Ê⁄ Œ’Ê„«  «·√”ÿ— (Œ’„ «·›« Ê—… ÌÊ“Û¯⁄ ⁄·Ï «·√”ÿ—)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ«÷⁄ ··÷—Ì»…", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", "„Ã„Ê⁄ ÷—Ì»… «·√”ÿ—"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", "TaxableAmount + Tax"
    AddField tdf, "RefundedAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„—œÊœ ‰ﬁœ«", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "InvoiceSubType", "TEXT", 10, True, """SIMPLIFIED""", _
             "In (""SIMPLIFIED"",""STANDARD"")", "SIMPLIFIED = ›« Ê—… „»”ÿ…° STANDARD = ›« Ê—… ÷—Ì»Ì…", "‰Ê⁄ «·›« Ê—… «·÷—Ì»Ì…", "„»”ÿ… ··√›—«œ B2C° ÷—Ì»Ì… ··„‰‘¬  B2B (··⁄„Ì· —ﬁ„ ÷—Ì»Ì)"
    AddField tdf, "InvoiceTypeCode", "TEXT", 3, True, """381""", _
             "In (""388"",""381"",""383"")", "388 ›« Ê—…° 381 ≈‘⁄«— œ«∆‰° 383 ≈‘⁄«— „œÌ‰", "—„“ ‰Ê⁄ «·„” ‰œ", ""
    AddField tdf, "InvoiceUUID", "TEXT", 36, False, "", _
             "", "", "«·„⁄—¯› «·›—Ìœ UUID", ""
    AddField tdf, "ICV", "LONG", 0, False, "", _
             "", "", "⁄œ¯«œ «·›Ê« Ì— ICV", " ”·”· „‘ —ﬂ ·ﬂ· «·„” ‰œ«  «·„—”·… ··ÂÌ∆…"
    AddField tdf, "InvoiceHash", "TEXT", 255, False, "", _
             "", "", "»’„… «·„” ‰œ", ""
    AddField tdf, "PreviousInvoiceHash", "TEXT", 255, False, "", _
             "", "", "»’„… «·„” ‰œ «·”«»ﬁ", ""
    AddField tdf, "QRCodeData", "MEMO", 0, False, "", _
             "", "", "»Ì«‰«  —„“ QR", ""
    AddField tdf, "ZatcaStatus", "TEXT", 20, True, """NOT_SENT""", _
             "In (""NOT_SENT"",""PENDING"",""REPORTED"",""CLEARED"",""WARNING"",""REJECTED"")", "Õ«·… €Ì— „⁄—Ê›…", "Õ«·… «·≈—”«· ··ÂÌ∆…", ""
    AddField tdf, "ZatcaSubmittedAt", "DATETIME", 0, False, "", _
             "", "", " «—ÌŒ «·≈—”«· ··ÂÌ∆…", ""
    AddField tdf, "ZatcaResponse", "MEMO", 0, False, "", _
             "", "", "—œ «·ÂÌ∆…", ""
    AddField tdf, "SignedXmlPath", "TEXT", 255, False, "", _
             "", "", "„”«— „·› XML «·„Êﬁ¯⁄", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "SalesReturnID", True, True, False
    AddIndex tdf, "UX_ReturnNumber", "ReturnNumber", False, True, False
    AddIndex tdf, "IX_ReturnDate", "ReturnDate", False, False, False
    AddIndex tdf, "UX_InvoiceUUID", "InvoiceUUID", False, True, True
    EndTable tdf, "„— Ã⁄«  «·„»Ì⁄« : ≈‘⁄«— œ«∆‰ „— »ÿ »«·›« Ê—… «·√’·Ì…∫ Ì⁄Ìœ «·ﬂ„Ì… ··„Œ“Ê‰ ÊÌ⁄ﬂ” «·„»·€.", "[RefundedAmount]<=[TotalAmount]", "«·„»·€ «·„—œÊœ ·« Ì Ã«Ê“ ﬁÌ„… «·„— Ã⁄"
End Sub

Private Sub CreateTable_SalesReturnDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesReturnDetails") Then Exit Sub
    AddField tdf, "ReturnDetailID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "SalesReturnID", "LONG", 0, True, "", _
             "", "", "«·„— Ã⁄", ""
    AddField tdf, "SalesDetailID", "LONG", 0, True, "", _
             "", "", "”ÿ— «·›« Ê—… «·√’·Ì", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—", "«·ﬂ„Ì… «·„— Ã⁄…", ""
    AddField tdf, "UnitPrice", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "”⁄— «·ÊÕœ…", ""
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·’«›Ì ﬁ»· «·÷—Ì»…", ""
    AddField tdf, "VATCategory", "TEXT", 1, True, """S""", _
             "In (""S"",""Z"",""E"")", "S = Œ«÷⁄ 15%° Z = ‰”»… ’›—Ì…° E = „⁄›Ï", "«·›∆… «·÷—Ì»Ì…", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·÷—Ì»…", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", "‰›”  ﬂ·›… ”ÿ— «·»Ì⁄ «·√’·Ì"
    AddField tdf, "ReturnToStock", "BOOL", 0, False, "True", _
             "", "", "Ì⁄«œ ··„Œ“Ê‰", ""
    AddIndex tdf, "PrimaryKey", "ReturnDetailID", True, True, False
    EndTable tdf, " ›«’Ì· „— Ã⁄«  «·„»Ì⁄« : «·√”ÿ— «·„— Ã⁄…° ﬂ· ”ÿ— Ì‘Ì— ≈·Ï ”ÿ— «·›« Ê—… «·√’·Ì ·„‰⁄ ≈—Ã«⁄ √ﬂÀ— „„« »Ì⁄.", "", ""
End Sub

Private Sub CreateTable_PurchaseInvoices()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseInvoices") Then Exit Sub
    AddField tdf, "PurchaseInvoiceID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "InvoiceNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·›« Ê—… «·œ«Œ·Ì", ""
    AddField tdf, "SupplierInvoiceNo", "TEXT", 30, False, "", _
             "", "", "—ﬁ„ ›« Ê—… «·„Ê—œ", ""
    AddField tdf, "InvoiceDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·›« Ê—…", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "«·„Ê—œ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "PaymentType", "TEXT", 10, True, """CASH""", _
             "In (""CASH"",""CREDIT"")", "CASH = ‰ﬁœÌ° CREDIT = ¬Ã·", "‰Ê⁄ «·‘—«¡", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„", "„Ã„Ê⁄ («·ﬂ„Ì… ◊ ”⁄— «·ÊÕœ…) »œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", "„Ã„Ê⁄ Œ’Ê„«  «·√”ÿ— (Œ’„ «·›« Ê—… ÌÊ“Û¯⁄ ⁄·Ï «·√”ÿ—)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ«÷⁄ ··÷—Ì»…", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", "„Ã„Ê⁄ ÷—Ì»… «·√”ÿ—"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", "TaxableAmount + Tax"
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„œ›Ê⁄", ""
    AddField tdf, "RemainingAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„ »ﬁÌ", "Ìı÷«› ≈·Ï —’Ìœ «·„Ê—œ"
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PurchaseInvoiceID", True, True, False
    AddIndex tdf, "UX_InvoiceNumber", "InvoiceNumber", False, True, False
    AddIndex tdf, "IX_InvoiceDate", "InvoiceDate", False, False, False
    AddIndex tdf, "IX_SupplierInvoiceNo", "SupplierInvoiceNo", False, False, False
    EndTable tdf, "›Ê« Ì— «·„‘ —Ì« : —√” ›« Ê—… «·‘—«¡ „‰ «·„Ê—œ.", "[PaidAmount]<=[TotalAmount] And [RemainingAmount]=[TotalAmount]-[PaidAmount]", "«·„œ›Ê⁄ ·« Ì Ã«Ê“ «·≈Ã„«·Ì° Ê«·„ »ﬁÌ = «·≈Ã„«·Ì - «·„œ›Ê⁄"
End Sub

Private Sub CreateTable_PurchaseInvoiceDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseInvoiceDetails") Then Exit Sub
    AddField tdf, "PurchaseDetailID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, True, "", _
             "", "", "«·›« Ê—…", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ—", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—", "«·ﬂ„Ì…", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", "»œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·’«›Ì ﬁ»· «·÷—Ì»…", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·„œŒ·« ", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", ""
    AddIndex tdf, "PrimaryKey", "PurchaseDetailID", True, True, False
    AddIndex tdf, "UX_PurchaseInvoiceID_LineNumber", "PurchaseInvoiceID,LineNumber", False, True, False
    EndTable tdf, " ›«’Ì· ›Ê« Ì— «·„‘ —Ì« : √”ÿ— ›« Ê—… «·‘—«¡.", "", ""
End Sub

Private Sub CreateTable_PurchaseReturns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseReturns") Then Exit Sub
    AddField tdf, "PurchaseReturnID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ReturnNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·„— Ã⁄", ""
    AddField tdf, "ReturnDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·„— Ã⁄", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, True, "", _
             "", "", "›« Ê—… «·‘—«¡ «·√’·Ì…", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "«·„Ê—œ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "Reason", "TEXT", 255, True, "", _
             "", "", "”»» «·≈—Ã«⁄", ""
    AddField tdf, "RefundType", "TEXT", 10, True, """CREDIT""", _
             "In (""CASH"",""CREDIT"")", "CASH = «” —œ«œ ‰ﬁœÌ° CREDIT = Œ’„ „‰ —’Ìœ «·„Ê—œ", "ÿ—Ìﬁ… «·«” —œ«œ", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "ÿ—Ìﬁ… «·«” —œ«œ", ""
    AddField tdf, "SubTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„Ã„Ê⁄ ﬁ»· «·Œ’„", "„Ã„Ê⁄ («·ﬂ„Ì… ◊ ”⁄— «·ÊÕœ…) »œÊ‰ ÷—Ì»…"
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", "„Ã„Ê⁄ Œ’Ê„«  «·√”ÿ— (Œ’„ «·›« Ê—… ÌÊ“Û¯⁄ ⁄·Ï «·√”ÿ—)"
    AddField tdf, "TaxableAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ«÷⁄ ··÷—Ì»…", "SubTotal - Discount"
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·ﬁÌ„… «·„÷«›…", "„Ã„Ê⁄ ÷—Ì»… «·√”ÿ—"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", "TaxableAmount + Tax"
    AddField tdf, "RefundedAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„” —œ ‰ﬁœ«", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PurchaseReturnID", True, True, False
    AddIndex tdf, "UX_ReturnNumber", "ReturnNumber", False, True, False
    AddIndex tdf, "IX_ReturnDate", "ReturnDate", False, False, False
    EndTable tdf, "„— Ã⁄«  «·„‘ —Ì« : ≈—Ã«⁄ »÷«⁄… ··„Ê—œ „— »ÿ »›« Ê—… «·‘—«¡ «·√’·Ì….", "[RefundedAmount]<=[TotalAmount]", "«·„»·€ «·„” —œ ·« Ì Ã«Ê“ ﬁÌ„… «·„— Ã⁄"
End Sub

Private Sub CreateTable_PurchaseReturnDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PurchaseReturnDetails") Then Exit Sub
    AddField tdf, "PurchaseReturnDetailID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "PurchaseReturnID", "LONG", 0, True, "", _
             "", "", "«·„— Ã⁄", ""
    AddField tdf, "PurchaseDetailID", "LONG", 0, True, "", _
             "", "", "”ÿ— «·›« Ê—… «·√’·Ì", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             ">0", "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ √ﬂ»— „‰ ’›—", "«·ﬂ„Ì… «·„— Ã⁄…", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", ""
    AddField tdf, "Discount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Œ’„", ""
    AddField tdf, "NetAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·’«›Ì ﬁ»· «·÷—Ì»…", ""
    AddField tdf, "VATRate", "RATE", 0, True, "0.15", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·÷—Ì»…", ""
    AddField tdf, "LineTotal", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì ‘«„· «·÷—Ì»…", ""
    AddIndex tdf, "PrimaryKey", "PurchaseReturnDetailID", True, True, False
    EndTable tdf, " ›«’Ì· „— Ã⁄«  «·„‘ —Ì« : «·√”ÿ— «·„— Ã⁄… ··„Ê—œ.", "", ""
End Sub

Private Sub CreateTable_CustomerPayments()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CustomerPayments") Then Exit Sub
    AddField tdf, "PaymentID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "PaymentNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·”‰œ", ""
    AddField tdf, "CustomerID", "LONG", 0, True, "", _
             "", "", "«·⁄„Ì·", ""
    AddField tdf, "PaymentDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·œ›⁄…", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "1", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, False, "", _
             "", "", "⁄‰ ›« Ê—… («Œ Ì«—Ì)", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PaymentID", True, True, False
    AddIndex tdf, "UX_PaymentNumber", "PaymentNumber", False, True, False
    AddIndex tdf, "IX_PaymentDate", "PaymentDate", False, False, False
    EndTable tdf, "œ›⁄«  «·⁄„·«¡ (”‰œ«  «·ﬁ»÷): «·„»«·€ «·„” ·„… „‰ «·⁄„·«¡ ·”œ«œ √—’œ Â„ «·¬Ã·….", "", ""
End Sub

Private Sub CreateTable_SupplierPayments()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SupplierPayments") Then Exit Sub
    AddField tdf, "PaymentID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "PaymentNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·”‰œ", ""
    AddField tdf, "SupplierID", "LONG", 0, True, "", _
             "", "", "«·„Ê—œ", ""
    AddField tdf, "PaymentDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·œ›⁄…", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "1", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, False, "", _
             "", "", "⁄‰ ›« Ê—… («Œ Ì«—Ì)", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PaymentID", True, True, False
    AddIndex tdf, "UX_PaymentNumber", "PaymentNumber", False, True, False
    AddIndex tdf, "IX_PaymentDate", "PaymentDate", False, False, False
    EndTable tdf, "œ›⁄«  «·„Ê—œÌ‰ (”‰œ«  «·’—›): «·„»«·€ «·„œ›Ê⁄… ··„Ê—œÌ‰ ·”œ«œ √—’œ Â„.", "", ""
End Sub

Private Sub CreateTable_ExpenseTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ExpenseTypes") Then Exit Sub
    AddField tdf, "ExpenseTypeID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·‰Ê⁄", ""
    AddField tdf, "ExpenseTypeName", "TEXT", 50, True, "", _
             "", "", "‰Ê⁄ «·„’—Ê›", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddIndex tdf, "PrimaryKey", "ExpenseTypeID", True, True, False
    AddIndex tdf, "UX_ExpenseTypeName", "ExpenseTypeName", False, True, False
    EndTable tdf, "√‰Ê«⁄ «·„’—Ê›« : ﬁ«∆„… À«» … ·√‰Ê«⁄ «·„’—Ê›«  ·÷„«‰ œﬁ… «· ﬁ«—Ì— «·„Ã„¯⁄….", "", ""
End Sub

Private Sub CreateTable_Expenses()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Expenses") Then Exit Sub
    AddField tdf, "ExpenseID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ExpenseNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·„’—Ê›", ""
    AddField tdf, "ExpenseDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·„’—Ê›", ""
    AddField tdf, "ExpenseTypeID", "LONG", 0, True, "", _
             "", "", "‰Ê⁄ «·„’—Ê›", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€ ﬁ»· «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·„œŒ·« ", "›ﬁÿ ≈–« ﬂ«‰  ·œÏ «·„Õ· ›« Ê—… ÷—Ì»Ì… »«·„’—Ê›"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·≈Ã„«·Ì", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, False, "", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "SupplierInvoiceRef", "TEXT", 30, False, "", _
             "", "", "—ﬁ„ ›« Ê—… «·„’—Ê›", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·Ê’›", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "ExpenseID", True, True, False
    AddIndex tdf, "UX_ExpenseNumber", "ExpenseNumber", False, True, False
    AddIndex tdf, "IX_ExpenseDate", "ExpenseDate", False, False, False
    EndTable tdf, "«·„’—Ê›« : „’—Ê›«  «·„Õ· «· ‘€Ì·Ì…∫ «·„»·€ »œÊ‰ ÷—Ì»… Ê«·÷—Ì»… „‰›’·… (÷—Ì»… „œŒ·« ).", "[TotalAmount]=[Amount]+[Tax]", "«·≈Ã„«·Ì = «·„»·€ + «·÷—Ì»…"
End Sub

Private Sub CreateTable_TransactionTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "TransactionTypes") Then Exit Sub
    AddField tdf, "TransactionTypeID", "LONG", 0, True, "", _
             "", "", "—ﬁ„ «·‰Ê⁄", ""
    AddField tdf, "TypeCode", "TEXT", 20, True, "", _
             "", "", "—„“ «·‰Ê⁄", ""
    AddField tdf, "TypeName", "TEXT", 50, True, "", _
             "", "", "‰Ê⁄ «·Õ—ﬂ…", ""
    AddField tdf, "Direction", "INT", 0, True, "", _
             "In (-1,0,1)", "«·« Ã«Â 1 √Ê -1 √Ê 0", "«·« Ã«Â", ""
    AddField tdf, "IsManual", "BOOL", 0, False, "False", _
             "", "", "„ «Õ ··≈œŒ«· «·ÌœÊÌ", ""
    AddIndex tdf, "PrimaryKey", "TransactionTypeID", True, True, False
    AddIndex tdf, "UX_TypeCode", "TypeCode", False, True, False
    EndTable tdf, "√‰Ê«⁄ Õ—ﬂ«  «·„Œ“Ê‰: √‰Ê«⁄ «·Õ—ﬂ… Ê≈‘«— Â«: +1  “Ìœ «·„Œ“Ê‰° -1  ‰ﬁ’Â° 0  ”ÊÌ… »«·≈‘«—….", "", ""
End Sub

Private Sub CreateTable_InventoryTransactions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "InventoryTransactions") Then Exit Sub
    AddField tdf, "TransactionID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·Õ—ﬂ…", ""
    AddField tdf, "TransactionDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·Õ—ﬂ…", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "TransactionTypeID", "LONG", 0, True, "", _
             "", "", "‰Ê⁄ «·Õ—ﬂ…", ""
    AddField tdf, "Quantity", "QTY", 0, True, "", _
             "<>0", "«·ﬂ„Ì… ·« Ì„ﬂ‰ √‰  ﬂÊ‰ ’›—«", "«·ﬂ„Ì… (+/-)", ""
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", ""
    AddField tdf, "QuantityAfter", "QTY", 0, True, "0", _
             "", "", "«·—’Ìœ »⁄œ «·Õ—ﬂ…", ""
    AddField tdf, "ReferenceType", "TEXT", 20, False, "", _
             "", "", "‰Ê⁄ «·„” ‰œ", "SALE, SALES_RETURN, PURCHASE, PURCHASE_RETURN, STOCK_COUNT, MANUAL"
    AddField tdf, "ReferenceID", "LONG", 0, False, "", _
             "", "", "—ﬁ„ «·„” ‰œ «·œ«Œ·Ì", ""
    AddField tdf, "ReferenceNumber", "TEXT", 20, False, "", _
             "", "", "—ﬁ„ «·„” ‰œ", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "TransactionID", True, True, False
    AddIndex tdf, "IX_TransactionDate", "TransactionDate", False, False, False
    AddIndex tdf, "IX_ReferenceType_ReferenceID", "ReferenceType,ReferenceID", False, False, False
    EndTable tdf, "Õ—ﬂ… «·„Œ“Ê‰: œ› — √” «– «·„Œ“Ê‰: «·ﬂ„Ì… „Œ“‰… »≈‘«— Â«° Ê—’Ìœ √Ì ’‰› = „Ã„Ê⁄ Quantity.", "", ""
End Sub

Private Sub CreateTable_StockCounts()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "StockCounts") Then Exit Sub
    AddField tdf, "StockCountID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "CountNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·Ã—œ", ""
    AddField tdf, "CountDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·Ã—œ", ""
    AddField tdf, "CategoryID", "LONG", 0, False, "", _
             "", "", " ’‰Ì› „Õœœ («Œ Ì«—Ì)", ""
    AddField tdf, "Status", "TEXT", 10, True, """OPEN""", _
             "In (""OPEN"",""POSTED"",""CANCELLED"")", "OPEN „› ÊÕ° POSTED „ı—Õ¯·° CANCELLED „·€Ï", "«·Õ«·…", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√Ã—«Â", ""
    AddField tdf, "PostedAt", "DATETIME", 0, False, "", _
             "", "", " «—ÌŒ «· —ÕÌ·", ""
    AddField tdf, "PostedByID", "LONG", 0, False, "", _
             "", "", "—Õ¯·Â", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddIndex tdf, "PrimaryKey", "StockCountID", True, True, False
    AddIndex tdf, "UX_CountNumber", "CountNumber", False, True, False
    EndTable tdf, "Ã·”«  «·Ã—œ: —√” ⁄„·Ì… «·Ã—œ∫  »ﬁÏ „› ÊÕ… Õ Ï «· —ÕÌ· «·–Ì Ì‰‘∆ Õ—ﬂ«  «· ”ÊÌ….", "", ""
End Sub

Private Sub CreateTable_StockCountDetails()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "StockCountDetails") Then Exit Sub
    AddField tdf, "StockCountDetailID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "StockCountID", "LONG", 0, True, "", _
             "", "", "«·Ã—œ", ""
    AddField tdf, "ProductID", "LONG", 0, True, "", _
             "", "", "«·„‰ Ã", ""
    AddField tdf, "SystemQuantity", "QTY", 0, True, "0", _
             "", "", "«·ﬂ„Ì… «·„”Ã·…", ""
    AddField tdf, "ActualQuantity", "QTY", 0, False, "", _
             "Is Null Or >=0", "«·ﬂ„Ì… «·›⁄·Ì… ·« Ì„ﬂ‰ √‰  ﬂÊ‰ ”«·»…", "«·ﬂ„Ì… «·›⁄·Ì…", ""
    AddField tdf, "Difference", "QTY", 0, True, "0", _
             "", "", "«·›—ﬁ", "ActualQuantity - SystemQuantity"
    AddField tdf, "UnitCost", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", " ﬂ·›… «·ÊÕœ…", ""
    AddField tdf, "DifferenceValue", "MONEY", 0, True, "0", _
             "", "", "ﬁÌ„… «·›—ﬁ", "”«·» = ⁄Ã“° „ÊÃ» = “Ì«œ…"
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddIndex tdf, "PrimaryKey", "StockCountDetailID", True, True, False
    AddIndex tdf, "UX_StockCountID_ProductID", "StockCountID,ProductID", False, True, False
    EndTable tdf, " ›«’Ì· «·Ã—œ: «·ﬂ„Ì… «·„”Ã·… Ê«·›⁄·Ì… Ê«·›—ﬁ ·ﬂ· ’‰› ›Ì Ã·”… «·Ã—œ.", "", ""
End Sub

Private Sub CreateTable_AuditLog()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "AuditLog") Then Exit Sub
    AddField tdf, "LogID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”Ã·", ""
    AddField tdf, "LogDate", "DATETIME", 0, True, "Now()", _
             "", "", "«· «—ÌŒ", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "ActionType", "TEXT", 30, True, "", _
             "", "", "‰Ê⁄ «·⁄„·Ì…", ""
    AddField tdf, "ObjectName", "TEXT", 50, False, "", _
             "", "", "«·ﬂ«∆‰", ""
    AddField tdf, "RecordID", "TEXT", 30, False, "", _
             "", "", "—ﬁ„ «·”Ã· «·„ √À—", ""
    AddField tdf, "Details", "MEMO", 0, False, "", _
             "", "", "«· ›«’Ì·", ""
    AddField tdf, "ComputerName", "TEXT", 50, False, "", _
             "", "", "«”„ «·ÃÂ«“", ""
    AddIndex tdf, "PrimaryKey", "LogID", True, True, False
    AddIndex tdf, "IX_LogDate", "LogDate", False, False, False
    AddIndex tdf, "IX_ActionType", "ActionType", False, False, False
    EndTable tdf, "”Ã· «·⁄„·Ì« : Ì”Ã· «·œŒÊ· Ê«·Œ—ÊÃ Ê«·⁄„·Ì«  «·Õ”«”… ( Ã«Ê“ «·„Œ“Ê‰°  ⁄œÌ· «·√”⁄«—° «·‰”Œ «·«Õ Ì«ÿÌ).", "", ""
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
    Seed_Customers
    Seed_ExpenseTypes
    Seed_TransactionTypes
End Sub

Private Sub Seed_Settings()
    If Not BeginSeed("Settings") Then Exit Sub
    ExecSeed "INSERT INTO [Settings] ([SettingID], [StoreName], [CountryCode], [VATRate], [PricesIncludeVAT], [AllowNegativeStock], [CurrencyCode], [DefaultCustomerID], [ZatcaPhase], [LastInvoiceHash], [BackupKeepCount], [SlowMovingDays], [ReceiptFooter]) VALUES (1, '«”„ «·„Õ·', 'SA', 0.15, True, False, 'SAR', 1, 1, 'NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ==', 30, 90, '‘ﬂ—« ·“Ì«— ﬂ„')"
    EndSeed "Settings", 1
End Sub

Private Sub Seed_Sequences()
    If Not BeginSeed("Sequences") Then Exit Sub
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_INVOICE', 'INV-', 1, 6, '›Ê« Ì— «·„»Ì⁄« ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_RETURN', 'CRN-', 1, 6, '„— Ã⁄«  «·„»Ì⁄«  (≈‘⁄«— œ«∆‰)')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_INVOICE', 'PUR-', 1, 6, '›Ê« Ì— «·„‘ —Ì« ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_RETURN', 'PRT-', 1, 6, '„— Ã⁄«  «·„‘ —Ì« ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CUSTOMER_PAYMENT', 'RCV-', 1, 6, '”‰œ«  «·ﬁ»÷ „‰ «·⁄„·«¡')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SUPPLIER_PAYMENT', 'PAY-', 1, 6, '”‰œ«  «·’—› ··„Ê—œÌ‰')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('EXPENSE', 'EXP-', 1, 6, '«·„’—Ê›« ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_COUNT', 'CNT-', 1, 5, 'Ã·”«  «·Ã—œ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_ADJUST', 'ADJ-', 1, 6, 'Õ—ﬂ«  «·„Œ“Ê‰ «·ÌœÊÌ…')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PRODUCT_CODE', 'P', 1, 5, '√ﬂÊ«œ «·„‰ Ã« ')"
    ExecSeed "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('ZATCA_ICV', '', 1, 0, '⁄œ¯«œ ICV ·„” ‰œ«  ›« Ê—…')"
    EndSeed "Sequences", 11
End Sub

Private Sub Seed_Roles()
    If Not BeginSeed("Roles") Then Exit Sub
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (1, 'ADMIN', '„œÌ— «·‰Ÿ«„', 'Ã„Ì⁄ «·’·«ÕÌ« ')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (2, 'MANAGER', '„œÌ—', '«·„»Ì⁄«  Ê«·„‘ —Ì«  Ê«·„Œ“Ê‰ Ê«· ﬁ«—Ì—')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (3, 'CASHIER', 'ﬂ«‘Ì—', '«·„»Ì⁄«  Ê«·⁄„·«¡ ›ﬁÿ')"
    EndSeed "Roles", 3
End Sub

Private Sub Seed_Permissions()
    If Not BeginSeed("Permissions") Then Exit Sub
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_POS', '‰ﬁÿ… «·»Ì⁄', '«·„»Ì⁄« ', 10)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_VIEW', '⁄—÷ Ê≈⁄«œ… ÿ»«⁄… «·›Ê« Ì—', '«·„»Ì⁄« ', 11)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_RETURN', '„— Ã⁄«  «·„»Ì⁄« ', '«·„»Ì⁄« ', 12)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRICE_OVERRIDE', ' ⁄œÌ· ”⁄— «·»Ì⁄ ›Ì «·›« Ê—…', '«·„»Ì⁄« ', 13)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DISCOUNT_OVERRIDE', 'Œ’„ √⁄·Ï „‰ «·Õœ «·„”„ÊÕ', '«·„»Ì⁄« ', 14)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('ALLOW_NEGATIVE_STOCK', '«·»Ì⁄ »ﬂ„Ì… √ﬂ»— „‰ «·„ Ê›—', '«·„»Ì⁄« ', 15)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMERS', '≈œ«—… «·⁄„·«¡', '«·⁄„·«¡', 20)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMER_PAYMENTS', '”‰œ«  «·ﬁ»÷', '«·⁄„·«¡', 21)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASES', '›Ê« Ì— «·„‘ —Ì« ', '«·„‘ —Ì« ', 30)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASE_RETURN', '„— Ã⁄«  «·„‘ —Ì« ', '«·„‘ —Ì« ', 31)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIERS', '≈œ«—… «·„Ê—œÌ‰', '«·„Ê—œÊ‰', 40)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIER_PAYMENTS', '”‰œ«  «·’—›', '«·„Ê—œÊ‰', 41)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRODUCTS', '≈œ«—… «·„‰ Ã«  Ê«·√”⁄«—', '«·„Œ“Ê‰', 50)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('INVENTORY_ADJUST', '≈÷«›… ÊŒ’„ „Œ“Ê‰ ÌœÊÌ', '«·„Œ“Ê‰', 51)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('STOCK_COUNT', '«·Ã—œ', '«·„Œ“Ê‰', 52)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('EXPENSES', '«·„’—Ê›« ', '«·„’—Ê›« ', 60)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS', '«· ﬁ«—Ì— «· ‘€Ì·Ì…', '«· ﬁ«—Ì—', 70)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS_PROFIT', ' ﬁ«—Ì— «·√—»«Õ Ê«·÷—Ì»…', '«· ﬁ«—Ì—', 71)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DASHBOARD_FINANCIAL', '«·√—ﬁ«„ «·„«·Ì… ›Ì ·ÊÕ… «· Õﬂ„', '«· ﬁ«—Ì—', 72)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SETTINGS', '≈⁄œ«œ«  «·„Õ·', '«·‰Ÿ«„', 80)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('USERS', '«·„” Œœ„Ê‰ Ê«·’·«ÕÌ« ', '«·‰Ÿ«„', 81)"
    ExecSeed "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BACKUP', '«·‰”Œ «·«Õ Ì«ÿÌ', '«·‰Ÿ«„', 82)"
    EndSeed "Permissions", 22
End Sub

Private Sub Seed_RolePermissions()
    If Not BeginSeed("RolePermissions") Then Exit Sub
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
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMER_PAYMENTS')"
    EndSeed "RolePermissions", 44
End Sub

Private Sub Seed_Employees()
    If Not BeginSeed("Employees") Then Exit Sub
    ExecSeed "INSERT INTO [Employees] ([EmployeeID], [EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], [MustChangePassword], [IsActive]) VALUES (1, '„œÌ— «·‰Ÿ«„', '„œÌ— «·‰Ÿ«„', 'admin', 1, 1, True, True)"
    EndSeed "Employees", 1
End Sub

Private Sub Seed_Categories()
    If Not BeginSeed("Categories") Then Exit Sub
    ExecSeed "INSERT INTO [Categories] ([CategoryID], [CategoryName], [Description]) VALUES (1, '⁄«„', ' ’‰Ì› «› —«÷Ì')"
    EndSeed "Categories", 1
End Sub

Private Sub Seed_Units()
    If Not BeginSeed("Units") Then Exit Sub
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (1, 'Õ»…', 'PCE')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (2, '⁄·»…', 'BX')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (3, 'ﬂ— Ê‰', 'CT')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (4, '»«ﬂÌ ', 'PK')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (5, 'ﬂÌ·Ê', 'KGM')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (6, '· —', 'LTR')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (7, '„ —', 'MTR')"
    ExecSeed "INSERT INTO [Units] ([UnitID], [UnitName], [ZatcaUnitCode]) VALUES (8, 'ÿﬁ„', 'SET')"
    EndSeed "Units", 8
End Sub

Private Sub Seed_PaymentMethods()
    If Not BeginSeed("PaymentMethods") Then Exit Sub
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (1, '‰ﬁœÌ', '10', 1)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (2, '„œÏ / »ÿ«ﬁ…', '48', 2)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (3, ' ÕÊÌ· »‰ﬂÌ', '42', 3)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (4, '„Õ›Ÿ… ≈·ﬂ —Ê‰Ì…', '1', 4)"
    EndSeed "PaymentMethods", 4
End Sub

Private Sub Seed_Customers()
    If Not BeginSeed("Customers") Then Exit Sub
    ExecSeed "INSERT INTO [Customers] ([CustomerID], [CustomerName], [AllowCredit], [IsSystem]) VALUES (1, '⁄„Ì· ‰ﬁœÌ', False, True)"
    EndSeed "Customers", 1
End Sub

Private Sub Seed_ExpenseTypes()
    If Not BeginSeed("ExpenseTypes") Then Exit Sub
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (1, '«·≈ÌÃ«—')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (2, '«·ﬂÂ—»«¡')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (3, '«·„Ì«Â')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (4, '«·≈‰ —‰  Ê«·« ’«·« ')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (5, '«·‰ﬁ·')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (6, '«·’Ì«‰…')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (7, '«·—Ê« »')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (8, '«·„” ·“„« ')"
    ExecSeed "INSERT INTO [ExpenseTypes] ([ExpenseTypeID], [ExpenseTypeName]) VALUES (9, '„’—Ê›«  √Œ—Ï')"
    EndSeed "ExpenseTypes", 9
End Sub

Private Sub Seed_TransactionTypes()
    If Not BeginSeed("TransactionTypes") Then Exit Sub
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (1, 'PURCHASE', '‘—«¡', 1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (2, 'SALE', '»Ì⁄', -1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (3, 'PURCHASE_RETURN', '„— Ã⁄ ‘—«¡', -1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (4, 'SALES_RETURN', '„— Ã⁄ »Ì⁄', 1, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (5, 'STOCK_IN', '≈÷«›… „Œ“Ê‰', 1, True)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (6, 'STOCK_OUT', 'Œ’„ „Œ“Ê‰', -1, True)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (7, 'ADJUSTMENT', ' ”ÊÌ… Ã—œ', 0, False)"
    ExecSeed "INSERT INTO [TransactionTypes] ([TransactionTypeID], [TypeCode], [TypeName], [Direction], [IsManual]) VALUES (8, 'OPENING', '—’Ìœ «›  «ÕÌ', 1, True)"
    EndSeed "TransactionTypes", 8
End Sub
