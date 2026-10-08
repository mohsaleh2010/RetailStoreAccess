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

Private Const SCHEMA_TABLES As String = "Settings,Sequences,Roles,Permissions,RolePermissions,Employees,Screens,UserScreens,Activations,Categories,Units,PaymentMethods,Currencies,CurrencyRates,CashBoxes,Suppliers,Customers,Products,SalesInvoices,SalesInvoiceDetails,SalesReturns,SalesReturnDetails,PurchaseInvoices,PurchaseInvoiceDetails,PurchaseReturns,PurchaseReturnDetails,CustomerPayments,SupplierPayments,Banks,BankTransactions,Cheques," & _
    "FixedAssets,DepreciationRuns,AssetDepreciations,CostCenters,SalesReps,SalesRepTargets,CommissionRuns,CommissionLines,Budgets,BudgetLines,PayrollRuns,PayrollLines,BankReconciliations,BankClearings,CustomerAllocations,SupplierAllocations,ExpenseTypes,Expenses,RecurringExpenses,CashVouchers,CashClosings,Accounts,JournalSourceTypes,JournalEntries,JournalLines,PeriodClosings,FiscalYearClosings," & _
    "FiscalYearClosingLines,VatReturns,ManualEntries,ManualEntryLines,TransactionTypes,InventoryTransactions,StockCounts,StockCountDetails,AuditLog,AuditChanges,LabelSettings"
Private Const EXPECTED_FIELD_COUNTS As String = "Settings=39;Sequences=5;Roles=4;Permissions=4;RolePermissions=2;Employees=30;Screens=8;UserScreens=6;Activations=6;Categories=8;Units=4;PaymentMethods=6;Currencies=7;CurrencyRates=6;CashBoxes=8;Suppliers=17;Customers=23;Products=23;SalesInvoices=39;SalesInvoiceDetails=14;SalesReturns=32;SalesReturnDetails=14;PurchaseInvoices=23;PurchaseInvoiceDetails=11;PurchaseReturns=22;PurchaseReturnDetails=11;" & _
    "CustomerPayments=16;SupplierPayments=15;Banks=9;BankTransactions=15;Cheques=16;FixedAssets=26;DepreciationRuns=6;AssetDepreciations=5;CostCenters=7;SalesReps=13;SalesRepTargets=5;CommissionRuns=9;CommissionLines=13;Budgets=6;BudgetLines=17;PayrollRuns=12;PayrollLines=19;BankReconciliations=12;BankClearings=7;CustomerAllocations=6;SupplierAllocations=6;ExpenseTypes=3;Expenses=20;RecurringExpenses=18;" & _
    "CashVouchers=17;CashClosings=18;Accounts=15;JournalSourceTypes=4;JournalEntries=16;JournalLines=8;PeriodClosings=8;FiscalYearClosings=8;FiscalYearClosingLines=7;VatReturns=28;ManualEntries=13;ManualEntryLines=10;TransactionTypes=6;InventoryTransactions=13;StockCounts=9;StockCountDetails=9;AuditLog=9;AuditChanges=7;LabelSettings=19"
Private Const EXPECTED_SEED_COUNTS As String = "Settings=1;Sequences=25;Roles=3;Permissions=36;RolePermissions=71;Employees=1;Screens=58;Categories=1;Units=8;PaymentMethods=4;Currencies=11;CurrencyRates=5;CashBoxes=2;Customers=1;ExpenseTypes=9;Accounts=80;JournalSourceTypes=27;TransactionTypes=8;LabelSettings=1"

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
        MsgBox "„”«— „·› «·»Ì«‰«  ÌÃ» √‰ ÌŒ ·› ⁄‰ „·› «·Ê«ÃÂ… «·Õ«·Ì.", vbExclamation + MSG_RTL
        Exit Function
    End If

    m_log = "": m_created = 0: m_skipped = 0: m_seeded = 0: m_addedFields = 0
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
    m_currentStep = "developer user"
    EnsureDeveloperUser
    m_currentStep = "account tree"
    UpgradeAccountTree
    m_currentStep = "english names"
    SeedEnglishNames

    m_db.Close
    Set m_db = Nothing

    m_currentStep = "link tables"
    LinkBackEnd BackEndPath
    On Error Resume Next                     ' modAccounts may not be imported yet (manual installs)
    Application.Run "RebuildAccountTree"     ' levels and order of the account tree
    Err.Clear
    On Error GoTo EH

    LogLine "--- Ãœ«Ê· ÃœÌœ…: " & m_created & " | „ÊÃÊœ… „”»ﬁ«: " & m_skipped & _
            " | Ãœ«Ê·  „   ⁄»∆ Â«: " & m_seeded
    MsgBox " „ »‰«¡ «·Ãœ«Ê· »‰Ã«Õ." & vbCrLf & vbCrLf & _
           "Ãœ«Ê· ÃœÌœ…: " & m_created & vbCrLf & _
           "Ãœ«Ê· „ÊÃÊœ… „”»ﬁ«: " & m_skipped & vbCrLf & _
           "ÕﬁÊ· ÃœÌœ… √ı÷Ì›  ·Ãœ«Ê· „ÊÃÊœ…: " & m_addedFields & vbCrLf & _
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
        Call ResultBox("„·› «·»Ì«‰«  €Ì— „ÊÃÊœ: " & BackEndPath, vbCritical + MSG_RTL)
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
        Call ResultBox("«·›Õ’ ‰«ÃÕ: Ã„Ì⁄ «·Ãœ«Ê· (" & (UBound(Split(SCHEMA_TABLES, ",")) + 1) & _
               ") Ê«·»Ì«‰«  «·√”«”Ì… ”·Ì„….", vbInformation + MSG_RTL, "VerifySchema")
        VerifySchema = True
    Else
        Call ResultBox("⁄œœ «·„‘ﬂ·« : " & problems & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
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
    ' A table that exists already is upgraded: its missing fields are added, nothing is removed.
    m_currentStep = "create table " & TableName
    Set m_pending = New Collection
    BeginTable = True
    m_upgrade = TableExistsIn(m_db, TableName)
    If m_upgrade Then
        LogLine "  = „ÊÃÊœ „”»ﬁ«: " & TableName
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
                LogLine "  ~ √ı“Ì· ‘—ÿ «·Õﬁ·: " & tdf.Name & "." & FieldName
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
        LogLine "  + Õﬁ· ÃœÌœ: " & tdf.Name & "." & FieldName
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
        LogLine " ‰»ÌÂ: ÌÊÃœ „” Œœ„ »«”„ developer Ê·Ì” ÂÊ «·„»—„Ã∫ ·„ Ìı‰‘√ Õ”«» «·„»—„Ã."
        Exit Sub
    End If
    Do
        pwd = InputBox("≈‰‘«¡ Õ”«» «·„»—„Ã («”„ «·„” Œœ„: developer)." & vbCrLf & vbCrLf & _
                       "«ﬂ » ﬂ·„… „—Ê— ·Â (6 √Õ—› ⁄·Ï «·√ﬁ·) Ê«Õ ›Ÿ »Â«. ·«  ⁄ÿÂ« ·√Õœ." & vbCrLf & vbCrLf & _
                       "≈·€«¡ = ·« Ìı‰‘√ «·¬‰° ÊÌıÿ·» ›Ì «· ‘€Ì· «· «·Ì ·‹ BuildSchema.", "Õ”«» «·„»—„Ã")
        If Len(pwd) = 0 Then
            LogLine "·„ Ìı‰‘√ Õ”«» «·„»—„Ã: ·„  ıﬂ » ﬂ·„… „—Ê—."
            Exit Sub
        End If
        If Len(pwd) >= 6 And pwd = Trim$(pwd) And pwd <> "developer" Then Exit Do
        MsgBox "ﬂ·„… «·„—Ê— 6 √Õ—› ⁄·Ï «·√ﬁ·° »œÊ‰ „”«›… ›Ì √Ê·Â« √Ê ¬Œ—Â«° Ê·«  ”«ÊÌ «”„ «·„” Œœ„.", _
               vbExclamation + MSG_RTL, "Õ”«» «·„»—„Ã"
    Loop
    salt = Application.Run("NewSalt")
    m_db.Execute "INSERT INTO [Employees] ([EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], " & _
                 "[MustChangePassword], [IsActive], [IsDeveloper], [PasswordSalt], [PasswordHash]) VALUES " & _
                 "('«·„»—„Ã', '«·„»—„Ã', 'developer', 1, 1, False, True, True, '" & salt & "', '" & _
                 Application.Run("PasswordHash", pwd, salt) & "')", dbFailOnError
    LogLine " „ ≈‰‘«¡ Õ”«» «·„»—„Ã: developer"
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
        If m_seedAdded > 0 Then LogLine "  * ”Ã·«  ÃœÌœ…: " & TableName & " (" & m_seedAdded & " ”Ã·)"
        Exit Sub
    End If
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
    CreateTable_Screens
    CreateTable_UserScreens
    CreateTable_Activations
    CreateTable_Categories
    CreateTable_Units
    CreateTable_PaymentMethods
    CreateTable_Currencies
    CreateTable_CurrencyRates
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
    CreateTable_Banks
    CreateTable_BankTransactions
    CreateTable_Cheques
    CreateTable_FixedAssets
    CreateTable_DepreciationRuns
    CreateTable_AssetDepreciations
    CreateTable_CostCenters
    CreateTable_SalesReps
    CreateTable_SalesRepTargets
    CreateTable_CommissionRuns
    CreateTable_CommissionLines
    CreateTable_Budgets
    CreateTable_BudgetLines
    CreateTable_PayrollRuns
    CreateTable_PayrollLines
    CreateTable_BankReconciliations
    CreateTable_BankClearings
    CreateTable_CustomerAllocations
    CreateTable_SupplierAllocations
    CreateTable_ExpenseTypes
    CreateTable_Expenses
    CreateTable_RecurringExpenses
    CreateTable_CashVouchers
    CreateTable_CashClosings
    CreateTable_Accounts
    CreateTable_JournalSourceTypes
    CreateTable_JournalEntries
    CreateTable_JournalLines
    CreateTable_PeriodClosings
    CreateTable_FiscalYearClosings
    CreateTable_FiscalYearClosingLines
    CreateTable_VatReturns
    CreateTable_ManualEntries
    CreateTable_ManualEntryLines
    CreateTable_TransactionTypes
    CreateTable_InventoryTransactions
    CreateTable_StockCounts
    CreateTable_StockCountDetails
    CreateTable_AuditLog
    CreateTable_AuditChanges
    CreateTable_LabelSettings
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
    AddField tdf, "POSMode", "TEXT", 10, True, """RETAIL""", _
             "In (""RETAIL"",""RESTAURANT"",""CAFE"")", "«Œ — ‘«‘… «·»Ì⁄ „‰ «·ﬁ«∆„…", "‘«‘… «·»Ì⁄", "RETAIL = «·„Õ·«  (»«—ﬂÊœ)° RESTAURANT = «·„ÿ«⁄„ (·„”)° CAFE = «·ﬂ«›ÌÂ«  (·„”)"
    AddField tdf, "ImagesFolder", "TEXT", 255, False, "", _
             "", "", "„Ã·œ ’Ê— «·„‰ Ã« ", "«·„”«—«  «·‰”»Ì… ··’Ê—  ıﬁ—√ „‰Â∫ ›«—€ = „Ã·œ Images »Ã«‰» „·› «·»Ì«‰« "
    AddField tdf, "InvoicePrintMode", "TEXT", 10, True, """PREVIEW""", _
             "In (""DIRECT"",""PREVIEW"",""NONE"")", "«Œ — ÿ—Ìﬁ… «·ÿ»«⁄… „‰ «·ﬁ«∆„…", "«·ÿ»«⁄… ⁄‰œ Õ›Ÿ «·›« Ê—…", "DIRECT = ÿ»«⁄… „»«‘—… »œÊ‰ „⁄«Ì‰…° PREVIEW = ⁄—÷ «·„⁄«Ì‰…° NONE = »œÊ‰ ÿ»«⁄…"
    AddField tdf, "AllowAdminCompanyName", "BOOL", 0, False, "False", _
             "", "", "«·”„«Õ ·„œÌ— «·‰Ÿ«„ » €ÌÌ— «”„ «·„Õ·", ""
    AddField tdf, "ClosedThrough", "DATETIME", 0, False, "", _
             "", "", "«·› —… „ﬁ›·… Õ Ï (·« Ìı÷«› Ê·« Ìı⁄œÛ¯· „” ‰œ » «—ÌŒ Õ Ï Â–« «·ÌÊ„)", ""
    AddField tdf, "DefaultBankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ «·«› —«÷Ì ·· ÕÊÌ·«  «·»‰ﬂÌ…", "«·„»«·€ «·„œ›Ê⁄… √Ê «·„” ·„… »ÿ—Ìﬁ… ´ ÕÊÌ· »‰ﬂÌª  ıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "GosiEmployeeRate", "RATE", 0, True, "0.0975", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "«· √„Ì‰« : Õ’… «·„ÊŸ› «·”⁄ÊœÌ", ""
    AddField tdf, "GosiEmployerRate", "RATE", 0, True, "0.1175", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "«· √„Ì‰« : Õ’… «·„‰‘√… ⁄‰ «·”⁄ÊœÌ", ""
    AddField tdf, "GosiNonSaudiRate", "RATE", 0, True, "0.02", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "«· √„Ì‰« : Õ’… «·„‰‘√… ⁄‰ €Ì— «·”⁄ÊœÌ («·√Œÿ«— «·„Â‰Ì…)", ""
    AddField tdf, "GosiMaxWage", "MONEY", 0, True, "45000", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«· √„Ì‰« : «·Õœ «·√⁄·Ï ··√Ã— «·Œ«÷⁄", ""
    AddField tdf, "CreditBlockDays", "INT", 0, False, "0", _
             ">=0", "⁄œœ «·√Ì«„ ·« ÌﬂÊ‰ ”«·»«", "≈Ìﬁ«› «·»Ì⁄ «·¬Ã· ·⁄„Ì· „ √Œ— √ﬂÀ— „‰ (ÌÊ„)", ""
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", " œŒ· ›ÌÂ ‰ﬁœÌ… „»Ì⁄« Â Ê”‰œ« Â∫ ›«—€ = √Ê· ’‰œÊﬁ ﬂ«‘Ì— ‰‘ÿ"
    AddField tdf, "IsDeveloper", "BOOL", 0, False, "False", _
             "", "", "«·„»—„Ã", ""
    AddField tdf, "CustomScreens", "BOOL", 0, False, "False", _
             "", "", "’·«ÕÌ«  ‘«‘«  Œ«’…", ""
    AddField tdf, "OnPayroll", "BOOL", 0, False, "False", _
             "", "", "›Ì „”Ì— «·—Ê« »", ""
    AddField tdf, "BasicSalary", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·—« » «·√”«”Ì", ""
    AddField tdf, "HousingAllowance", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»œ· «·”ﬂ‰", ""
    AddField tdf, "TransportAllowance", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»œ· «·‰ﬁ·", ""
    AddField tdf, "OtherAllowance", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»œ·«  √Œ—Ï", ""
    AddField tdf, "IsSaudi", "BOOL", 0, False, "False", _
             "", "", "”⁄ÊœÌ («· √„Ì‰«  »Õ’ Ì «·„ÊŸ› Ê«·„‰‘√…)", ""
    AddField tdf, "AdvanceInstallment", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "ﬁ”ÿ «·”·›… «·‘Â—Ì", "0 = ÌıŒ’„ ﬂ· «·—’Ìœ ›Ì √Ê· „”Ì—"
    AddField tdf, "NationalID", "TEXT", 15, False, "", _
             "", "", "—ﬁ„ «·ÂÊÌ… / «·≈ﬁ«„…", ""
    AddField tdf, "IBAN", "TEXT", 34, False, "", _
             "", "", "¬Ì»«‰ «·„ÊŸ›", ""
    AddField tdf, "HireDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «· ⁄ÌÌ‰", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„»Ì⁄« Â Ê„”Ì— —« »Â ⁄·Ï Â–« «·„—ﬂ“"
    AddIndex tdf, "PrimaryKey", "EmployeeID", True, True, False
    AddIndex tdf, "UX_Username", "Username", False, True, False
    EndTable tdf, "«·„ÊŸ›Ê‰ Ê«·„” Œœ„Ê‰: ﬂ· „ÊŸ› ÂÊ „” Œœ„ ··‰Ÿ«„∫ ·« ÌıÕ–› »· Ìı⁄ÿÛ¯· ··Õ›«Ÿ ⁄·Ï ”Ã· ⁄„·Ì« Â.", "", ""
End Sub

Private Sub CreateTable_Screens()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Screens") Then Exit Sub
    AddField tdf, "ScreenName", "TEXT", 64, True, "", _
             "", "", "«”„ «·‘«‘… ›Ì Access", ""
    AddField tdf, "ScreenTitle", "TEXT", 100, True, "", _
             "", "", "«·‘«‘…", ""
    AddField tdf, "ModuleName", "TEXT", 50, False, "", _
             "", "", "«·ﬁ”„", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddField tdf, "PermissionKey", "TEXT", 50, False, "", _
             "", "", "’·«ÕÌ… «·œÊ—", "›«—€ = „ «Õ… ·ﬂ· «·„” Œœ„Ì‰∫  ı” Œœ„ ··„” Œœ„ «·–Ì ·Ì”  ·Â ’·«ÕÌ«  ‘«‘«  Œ«’…"
    AddField tdf, "HasAdd", "BOOL", 0, False, "False", _
             "", "", "›ÌÂ« ≈÷«›… / Õ›Ÿ „” ‰œ", ""
    AddField tdf, "HasEdit", "BOOL", 0, False, "False", _
             "", "", "›ÌÂ«  ⁄œÌ·", ""
    AddField tdf, "HasDelete", "BOOL", 0, False, "False", _
             "", "", "›ÌÂ« Õ–›", ""
    AddIndex tdf, "PrimaryKey", "ScreenName", True, True, False
    EndTable tdf, "«·‘«‘« : ﬂ· ‘«‘… ›Ì «·»—‰«„Ã° Ê„« Ì‰ÿ»ﬁ ⁄·ÌÂ« „‰ ≈÷«›… Ê ⁄œÌ· ÊÕ–›° Ê’·«ÕÌ… «·œÊ— «· Ì  › ÕÂ«.", "", ""
End Sub

Private Sub CreateTable_UserScreens()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "UserScreens") Then Exit Sub
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„” Œœ„", ""
    AddField tdf, "ScreenName", "TEXT", 64, True, "", _
             "", "", "«·‘«‘…", ""
    AddField tdf, "CanOpen", "BOOL", 0, False, "False", _
             "", "", "› Õ", ""
    AddField tdf, "CanAdd", "BOOL", 0, False, "False", _
             "", "", "≈÷«›…", ""
    AddField tdf, "CanEdit", "BOOL", 0, False, "False", _
             "", "", " ⁄œÌ·", ""
    AddField tdf, "CanDelete", "BOOL", 0, False, "False", _
             "", "", "Õ–›", ""
    AddIndex tdf, "PrimaryKey", "EmployeeID,ScreenName", True, True, False
    EndTable tdf, "’·«ÕÌ«  «·‘«‘«  ··„” Œœ„: ··„” Œœ„ «·–Ì ›ı⁄¯·  ·Â ´’·«ÕÌ«  ‘«‘«  Œ«’…ª: «·‘«‘«  «· Ì Ì› ÕÂ«° Ê«·≈÷«›… Ê«· ⁄œÌ· Ê«·Õ–› ›Ì ﬂ· ‘«‘….", "", ""
End Sub

Private Sub CreateTable_Activations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Activations") Then Exit Sub
    AddField tdf, "ActivationID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «· ›⁄Ì·", ""
    AddField tdf, "MachineID", "TEXT", 24, True, "", _
             "", "", "—ﬁ„ «·ÃÂ«“", "»’„… ·ÊÕ… «·√„ Ê«·„⁄«·Ã Êﬁ—’ «·‰Ÿ«„ (modActivation.MachineID)"
    AddField tdf, "ActivationCode", "TEXT", 30, True, "", _
             "", "", "ﬂÊœ «· ›⁄Ì·", ""
    AddField tdf, "ComputerName", "TEXT", 64, False, "", _
             "", "", "«”„ «·ÃÂ«“", ""
    AddField tdf, "ActivatedAt", "DATETIME", 0, False, "Now()", _
             "", "", " «—ÌŒ «· ›⁄Ì·", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "›⁄¯·Â", ""
    AddIndex tdf, "PrimaryKey", "ActivationID", True, True, False
    AddIndex tdf, "UX_MachineID", "MachineID", False, True, False
    EndTable tdf, " ›⁄Ì· «·»—‰«„Ã: «·√ÃÂ“… «·„›⁄¯· ⁄·ÌÂ« «·»—‰«„Ã: —ﬁ„ «·ÃÂ«“ ÊﬂÊœ «· ›⁄Ì· «·’«œ— „‰ «·„»—„Ã.", "", ""
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
    AddField tdf, "ImagePath", "TEXT", 255, False, "", _
             "", "", "’Ê—… «· ’‰Ì›", ""
    AddField tdf, "TileColor", "TEXT", 10, True, """BLUE""", _
             "In (""BLUE"",""GREEN"",""ORANGE"",""PURPLE"",""RED"",""INDIGO"",""TEAL"",""PINK"",""BROWN"",""GREY"")", "«Œ — «··Ê‰ „‰ «·ﬁ«∆„…", "·Ê‰ «·“—", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", " — Ì» «·⁄—÷", ""
    AddField tdf, "IsAddOn", "BOOL", 0, False, "False", _
             "", "", "›∆… ≈÷«›« ", ""
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
    AddField tdf, "MethodNameEn", "TEXT", 50, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", "ÌŸÂ— ›Ì «·Ê«ÃÂ… «·≈‰Ã·Ì“Ì…"
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

Private Sub CreateTable_Currencies()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Currencies") Then Exit Sub
    AddField tdf, "CurrencyCode", "TEXT", 3, True, "", _
             "Len([CurrencyCode])=3", "—„“ «·⁄„·… 3 √Õ—› (ISO 4217) „À· USD", "—„“ «·⁄„·…", ""
    AddField tdf, "CurrencyName", "TEXT", 50, True, "", _
             "", "", "«”„ «·⁄„·…", ""
    AddField tdf, "CurrencyNameEn", "TEXT", 50, False, "", _
             "", "", "«”„ «·⁄„·… »«·≈‰Ã·Ì“Ì…", ""
    AddField tdf, "Symbol", "TEXT", 10, False, "", _
             "", "", "«·—„“ «·„Œ ’—", ""
    AddField tdf, "DecimalPlaces", "BYTE", 0, True, "2", _
             "Between 0 And 3", "„‰ 0 ≈·Ï 3 Œ«‰« ", "⁄œœ «·Œ«‰«  «·⁄‘—Ì…", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddIndex tdf, "PrimaryKey", "CurrencyCode", True, True, False
    AddIndex tdf, "UX_CurrencyName", "CurrencyName", False, True, False
    EndTable tdf, "«·⁄„·« : ⁄„·«  «· ⁄«„·. ⁄„·… «·»—‰«„Ã ÂÌ Settings.CurrencyCode («·—Ì«·)° Êﬂ· «·„»«·€  ıÕ›Ÿ »Â«∫ «·„” ‰œ »⁄„·… √Œ—Ï ÌÕ›Ÿ ⁄„· Â Ê„⁄«„·Â Ê„»·€Â »Â«.", "", ""
End Sub

Private Sub CreateTable_CurrencyRates()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CurrencyRates") Then Exit Sub
    AddField tdf, "CurrencyRateID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, True, "", _
             "", "", "«·⁄„·…", ""
    AddField tdf, "RateDate", "DATE", 0, True, "Date()", _
             "", "", "«· «—ÌŒ", ""
    AddField tdf, "Rate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„⁄«„·", ""
    AddField tdf, "Notes", "TEXT", 100, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "CurrencyRateID", True, True, False
    AddIndex tdf, "UX_CurrencyCode_RateDate", "CurrencyCode,RateDate", False, True, False
    EndTable tdf, "√”⁄«— «·⁄„·« : „⁄«„· ﬂ· ⁄„·… ›Ì  «—ÌŒ: ﬁÌ„… ÊÕœ… Ê«Õœ… „‰Â« »⁄„·… «·»—‰«„Ã. «·„” ‰œ Ì√Œ– ¬Œ— ”⁄— ›Ì  «—ÌŒÂ √Ê ﬁ»·Â.", "", ""
End Sub

Private Sub CreateTable_CashBoxes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashBoxes") Then Exit Sub
    AddField tdf, "CashBoxID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·’‰œÊﬁ", ""
    AddField tdf, "BoxName", "TEXT", 50, True, "", _
             "", "", "«”„ «·’‰œÊﬁ", ""
    AddField tdf, "BoxType", "TEXT", 10, True, """CASHIER""", _
             "In (""MAIN"",""CASHIER"")", "MAIN = Œ“Ì‰… —∆Ì”Ì…° CASHIER = ’‰œÊﬁ ﬂ«‘Ì—", "«·‰Ê⁄", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·—’Ìœ «·«›  «ÕÌ", ""
    AddField tdf, "OpeningDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·—’Ìœ «·«›  «ÕÌ", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "CashBoxID", True, True, False
    AddIndex tdf, "UX_BoxName", "BoxName", False, True, False
    EndTable tdf, "«·Œ“Ì‰… Ê«·’‰«œÌﬁ: «·Œ“Ì‰… «·—∆Ì”Ì… Ê’‰«œÌﬁ «·ﬂ«‘Ì—. «·—’Ìœ ·« ÌıŒ“Û¯‰: ÌıÕ”» „‰ «·Õ—ﬂ«  (qryCashMovements).", "", ""
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
    AddField tdf, "PaymentTermsDays", "INT", 0, False, "30", _
             ">=0", "⁄œœ «·√Ì«„ ·« ÌﬂÊ‰ ”«·»«", "„œ… «·”œ«œ (ÌÊ„)", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "⁄„·… «· ⁄«„·", " ıﬁ —Õ ›Ì ›Ê« Ì—Â Ê”‰œ« Â"
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
    AddField tdf, "PaymentTermsDays", "INT", 0, False, "30", _
             ">=0", "⁄œœ «·√Ì«„ ·« ÌﬂÊ‰ ”«·»«", "„œ… «·”œ«œ (ÌÊ„)", ""
    AddField tdf, "IsSystem", "BOOL", 0, False, "False", _
             "", "", "”Ã· ‰Ÿ«„", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "MEMO", 0, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "SalesRepID", "LONG", 0, False, "", _
             "", "", "«·„‰œÊ»", "«·„‰œÊ» «·„”ƒÊ· ⁄‰ «·⁄„Ì·:  ı‰”» ·Â ›Ê« Ì—Â Ê Õ’Ì·« Â"
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
    AddField tdf, "ImagePath", "TEXT", 255, False, "", _
             "", "", "’Ê—… «·„‰ Ã", "·‘«‘«  «··„”∫ „”«— ﬂ«„· √Ê «”„ „·› ›Ì „Ã·œ «·’Ê—"
    AddField tdf, "TrackStock", "BOOL", 0, False, "True", _
             "", "", "Ì «»⁄ «·„Œ“Ê‰", ""
    AddField tdf, "SizePriceM", "MONEY", 0, False, "", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "”⁄— «·ÕÃ„ «·Ê”ÿ", ""
    AddField tdf, "SizePriceL", "MONEY", 0, False, "", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "”⁄— «·ÕÃ„ «·ﬂ»Ì—", ""
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
    AddField tdf, "DueDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·«” Õﬁ«ﬁ", ""
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
    AddField tdf, "OrderType", "TEXT", 10, False, "", _
             "Is Null Or In (""DINE_IN"",""TAKEAWAY"",""DELIVERY"")", "‰Ê⁄ «·ÿ·»: œ«Œ·Ì √Ê ”›—Ì √Ê  Ê’Ì·", "‰Ê⁄ «·ÿ·»", ""
    AddField tdf, "TableNo", "TEXT", 10, False, "", _
             "", "", "—ﬁ„ «·ÿ«Ê·…", ""
    AddField tdf, "DeliveryPhone", "TEXT", 20, False, "", _
             "", "", "ÃÊ«· «· Ê’Ì·", ""
    AddField tdf, "DeliveryAddress", "TEXT", 255, False, "", _
             "", "", "⁄‰Ê«‰ «· Ê’Ì·", ""
    AddField tdf, "OrderName", "TEXT", 50, False, "", _
             "", "", "«”„ «·⁄„Ì· ⁄·Ï «·ÿ·»", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "Ìı„·√ ⁄‰œ «·œ›⁄ «·‰ﬁœÌ: «·„»·€ «·„œ›Ê⁄ ÌœŒ· Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„‰ „—ﬂ“ «·ﬂ«‘Ì—° Ê≈·« «·„—ﬂ“ «·«› —«÷Ì"
    AddField tdf, "SalesRepID", "LONG", 0, False, "", _
             "", "", "«·„‰œÊ»", "„‰ „‰œÊ» «·⁄„Ì·° Ê≈·« „‰œÊ» «·„” Œœ„"
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
    AddField tdf, "LineNote", "TEXT", 100, False, "", _
             "", "", "„·«ÕŸ… «·”ÿ—", "«·ÕÃ„ Ê«·ŒÌ«—«  («·ﬂ«›ÌÂ)"
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "«·—œ «·‰ﬁœÌ ÌŒ—Ã „‰ Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„—ﬂ“ «·›« Ê—… «·√’·Ì…"
    AddField tdf, "SalesRepID", "LONG", 0, False, "", _
             "", "", "«·„‰œÊ»", "„‰œÊ» «·›« Ê—… «·√’·Ì…"
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
    AddField tdf, "DueDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·«” Õﬁ«ﬁ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "«·„œ›Ê⁄ ‰ﬁœ« ÌŒ—Ã „‰ Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ »«·⁄„·…", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "«·«” —œ«œ «·‰ﬁœÌ ÌœŒ· Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ »«·⁄„·…", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "«·„»·€ «·‰ﬁœÌ ÌœŒ· Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ »«·⁄„·…", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
    AddField tdf, "SalesRepID", "LONG", 0, False, "", _
             "", "", "«·„‰œÊ»", "«·„Õ’ˆ¯·: „‰œÊ» «·⁄„Ì·° Ê≈·« „‰œÊ» «·„” Œœ„"
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «·‰ﬁœÌ…", "«·„»·€ «·‰ﬁœÌ ÌŒ—Ã „‰ Â–« «·’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ »«·⁄„·…", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
    AddIndex tdf, "PrimaryKey", "PaymentID", True, True, False
    AddIndex tdf, "UX_PaymentNumber", "PaymentNumber", False, True, False
    AddIndex tdf, "IX_PaymentDate", "PaymentDate", False, False, False
    EndTable tdf, "œ›⁄«  «·„Ê—œÌ‰ (”‰œ«  «·’—›): «·„»«·€ «·„œ›Ê⁄… ··„Ê—œÌ‰ ·”œ«œ √—’œ Â„.", "", ""
End Sub

Private Sub CreateTable_Banks()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Banks") Then Exit Sub
    AddField tdf, "BankID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·»‰ﬂ", ""
    AddField tdf, "BankName", "TEXT", 100, True, "", _
             "", "", "«”„ «·»‰ﬂ / «·Õ”«»", ""
    AddField tdf, "AccountNo", "TEXT", 30, False, "", _
             "", "", "—ﬁ„ «·Õ”«»", ""
    AddField tdf, "IBAN", "TEXT", 34, False, "", _
             "", "", "«·¬Ì»«‰", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·«›  «ÕÌ", "—’Ìœ «·Õ”«» ›Ì «·»‰ﬂ ⁄‰œ »œ¡ «” Œœ«„ «·»—‰«„Ã"
    AddField tdf, "OpeningDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·—’Ìœ «·«›  «ÕÌ", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "BankID", True, True, False
    AddIndex tdf, "UX_BankName", "BankName", False, True, False
    EndTable tdf, "«·»‰Êﬂ: ﬂ· Õ”«» »‰ﬂÌ ··„Õ·. Õ”«»Â ›Ì «·œ·Ì· 120000 + —ﬁ„Â  Õ  ´«·Õ”«»«  «·»‰ﬂÌ…ª (1210).", "", ""
End Sub

Private Sub CreateTable_BankTransactions()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "BankTransactions") Then Exit Sub
    AddField tdf, "BankTxID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "TxNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·Õ—ﬂ…", ""
    AddField tdf, "TxDate", "DATE", 0, True, "Date()", _
             "", "", "«· «—ÌŒ", ""
    AddField tdf, "TxType", "TEXT", 12, True, "", _
             "In (""DEPOSIT"",""WITHDRAW"",""SETTLEMENT"",""TRANSFER"",""OTHER_IN"",""OTHER_OUT"")", "≈Ìœ«⁄° ”Õ»°  ”ÊÌ… „œÏ°  ÕÊÌ·° Ê«—œ ¬Œ—° ’«œ— ¬Œ—", "«·‰Ê⁄", ""
    AddField tdf, "BankID", "LONG", 0, True, "", _
             "", "", "«·»‰ﬂ", ""
    AddField tdf, "ToBankID", "LONG", 0, False, "", _
             "", "", "≈·Ï »‰ﬂ", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "«·’‰œÊﬁ", ""
    AddField tdf, "CounterAccount", "LONG", 0, False, "", _
             "", "", "«·Õ”«» «·„ﬁ«»·", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", "›Ì  ”ÊÌ… „œÏ: ≈Ã„«·Ì «· Õ’Ì·«  ﬁ»· «·⁄„Ê·…"
    AddField tdf, "FeeAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·⁄„Ê·…", " ”ÊÌ… „œÏ: ⁄„Ê·… «·»‰ﬂ »œÊ‰ ÷—Ì»…"
    AddField tdf, "FeeVAT", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·⁄„Ê·…", "÷—Ì»… „œŒ·«  ⁄·Ï «·⁄„Ê·…"
    AddField tdf, "Reference", "TEXT", 40, False, "", _
             "", "", "„—Ã⁄ «·»‰ﬂ", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·»Ì«‰", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "BankTxID", True, True, False
    AddIndex tdf, "UX_TxNumber", "TxNumber", False, True, False
    AddIndex tdf, "IX_TxDate", "TxDate", False, False, False
    EndTable tdf, "«·Õ—ﬂ«  «·»‰ﬂÌ…: ≈Ìœ«⁄ ‰ﬁœÌ… „‰ ’‰œÊﬁ° ”Õ» ≈·Ï ’‰œÊﬁ°  ”ÊÌ…  Õ’Ì·«  „œÏ (»⁄„Ê· Â«)°  ÕÊÌ· »Ì‰ »‰ﬂÌ‰° ÊÕ—ﬂ«  √Œ—Ï (⁄„Ê·« ° ›Ê«∆œ° ﬁ—Ê÷...) »Õ”«» „ﬁ«»·.", "[FeeAmount]+[FeeVAT]<[Amount] Or [TxType]<>""SETTLEMENT""", "«·⁄„Ê·… Ê÷—Ì» Â« √ﬁ· „‰ „»·€ «· ”ÊÌ…"
End Sub

Private Sub CreateTable_Cheques()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Cheques") Then Exit Sub
    AddField tdf, "ChequeID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ChequeRef", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·ﬁÌœ «·œ«Œ·Ì", ""
    AddField tdf, "Direction", "TEXT", 3, True, "", _
             "In (""IN"",""OUT"")", "IN = Ê«—œ° OUT = ’«œ—", "«·« Ã«Â", ""
    AddField tdf, "CustomerID", "LONG", 0, False, "", _
             "", "", "«·⁄„Ì· («·Ê«—œ)", ""
    AddField tdf, "SupplierID", "LONG", 0, False, "", _
             "", "", "«·„Ê—œ («·’«œ—)", ""
    AddField tdf, "ChequeNo", "TEXT", 30, True, "", _
             "", "", "—ﬁ„ «·‘Ìﬂ", ""
    AddField tdf, "DrawerBank", "TEXT", 100, False, "", _
             "", "", "»‰ﬂ «·”«Õ» («·Ê«—œ)", ""
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "»‰ﬂ‰«", "«·Ê«—œ: «·»‰ﬂ «·–Ì Õı’ˆ¯· ›ÌÂ∫ «·’«œ—: «·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ"
    AddField tdf, "IssueDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·«” ·«„ / «·≈’œ«—", ""
    AddField tdf, "DueDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·«” Õﬁ«ﬁ", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "Status", "TEXT", 10, True, """PENDING""", _
             "In (""PENDING"",""COLLECTED"",""BOUNCED"")", "PENDING =  Õ  «· Õ’Ì·° COLLECTED = „Õ’Û¯· / „’—Ê›° BOUNCED = „— œ", "«·Õ«·…", ""
    AddField tdf, "StatusDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «· Õ’Ì· / «·«— œ«œ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "ChequeID", True, True, False
    AddIndex tdf, "UX_ChequeRef", "ChequeRef", False, True, False
    AddIndex tdf, "IX_DueDate", "DueDate", False, False, False
    AddIndex tdf, "IX_ChequeNo", "ChequeNo", False, False, False
    EndTable tdf, "«·‘Ìﬂ«  «·Ê«—œ… Ê«·’«œ—…: ‘Ìﬂ Ê«—œ „‰ ⁄„Ì· (Ì”œœ —’ÌœÂ ÊÌ»ﬁÏ ›Ì ´‘Ìﬂ«   Õ  «· Õ’Ì·ª Õ Ï ÌıÕ’Û¯· ›Ì «·»‰ﬂ √Ê Ì— œ) √Ê ’«œ— ·„Ê—œ (Ì”œœ —’ÌœÂ ÊÌ»ﬁÏ ›Ì ´√Ê—«ﬁ «·œ›⁄ª Õ Ï Ì’—›Â «·»‰ﬂ √Ê Ì— œ).", "([Direction]=""IN"" And [CustomerID] Is Not Null And [SupplierID] Is Null) Or ([Direction]=""OUT"" And [SupplierID] Is Not Null And [CustomerID] Is Null)", "«·‘Ìﬂ «·Ê«—œ ·⁄„Ì·° Ê«·’«œ— ·„Ê—œ"
End Sub

Private Sub CreateTable_FixedAssets()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "FixedAssets") Then Exit Sub
    AddField tdf, "AssetID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "AssetCode", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·√’·", ""
    AddField tdf, "AssetName", "TEXT", 150, True, "", _
             "", "", "«”„ «·√’·", ""
    AddField tdf, "AssetAccount", "LONG", 0, True, "", _
             "", "", "Õ”«» «·√’·", "Õ”«» ›—⁄Ì „‰ «·√’Ê· €Ì— «·„ œ«Ê·… (12)° „À· «·√À«À √Ê «·√ÃÂ“…"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "Ì–Â» ≈·ÌÂ ≈Â·«ﬂ «·√’· Ê—»Õ √Ê Œ”«—… »Ì⁄Â"
    AddField tdf, "PurchaseDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·‘—«¡", ""
    AddField tdf, "Cost", "MONEY", 0, True, "0", _
             ">0", "«· ﬂ·›… √ﬂ»— „‰ ’›—", "«· ﬂ·›… »œÊ‰ «·÷—Ì»…", ""
    AddField tdf, "InputVAT", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·„œŒ·« ", ""
    AddField tdf, "SalvageValue", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·ﬁÌ„… «·„ »ﬁÌ… ›Ì ¬Œ— «·⁄„—", ""
    AddField tdf, "UsefulLifeMonths", "INT", 0, True, "", _
             ">0", "«·⁄„— «·≈‰ «ÃÌ ‘Â— ⁄·Ï «·√ﬁ·", "«·⁄„— «·≈‰ «ÃÌ (‘Â—)", ""
    AddField tdf, "DepStartDate", "DATE", 0, True, "Date()", _
             "", "", "»œ«Ì… «·≈Â·«ﬂ (ÌıÂ·ﬂ „‰ ‘Â— Â–« «· «—ÌŒ)", ""
    AddField tdf, "SourceType", "TEXT", 10, True, """BANK""", _
             "In (""OPENING"",""BANK"",""CASHBOX"",""ACCOUNT"")", "—’Ìœ «›  «ÕÌ° »‰ﬂ° ’‰œÊﬁ° √Ê Õ”«» ¬Œ—", "„’œ— «·‘—«¡", ""
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "«·’‰œÊﬁ", ""
    AddField tdf, "CounterAccount", "LONG", 0, False, "", _
             "", "", "«·Õ”«» «·œ«∆‰ (Õ”«» ¬Œ—)", ""
    AddField tdf, "OpeningAccumDep", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Â·«ﬂ ”«»ﬁ (··√’Ê· «·„ÊÃÊœ… ﬁ»· «·»—‰«„Ã)", ""
    AddField tdf, "Status", "TEXT", 10, True, """ACTIVE""", _
             "In (""ACTIVE"",""DISPOSED"")", "ACTIVE = ﬁ«∆„° DISPOSED = „” »⁄œ", "«·Õ«·…", ""
    AddField tdf, "DisposalDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·»Ì⁄ / «·«” »⁄«œ", ""
    AddField tdf, "DisposalProceeds", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "À„‰ «·»Ì⁄", ""
    AddField tdf, "DisposalTo", "TEXT", 10, False, "", _
             "Is Null Or In (""NONE"",""BANK"",""CASHBOX"")", "»œÊ‰ À„‰° »‰ﬂ° ’‰œÊﬁ", "«” ·«„ «·À„‰", ""
    AddField tdf, "DisposalBankID", "LONG", 0, False, "", _
             "", "", "»‰ﬂ «” ·«„ «·À„‰", ""
    AddField tdf, "DisposalCashBoxID", "LONG", 0, False, "", _
             "", "", "’‰œÊﬁ «” ·«„ «·À„‰", ""
    AddField tdf, "DisposalAccumDep", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„Ã„⁄ «·≈Â·«ﬂ ÌÊ„ «·«” »⁄«œ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "AssetID", True, True, False
    AddIndex tdf, "UX_AssetCode", "AssetCode", False, True, False
    AddIndex tdf, "IX_AssetAccount", "AssetAccount", False, False, False
    EndTable tdf, "«·√’Ê· «·À«» …: ”Ã· «·√’Ê·: «· ﬂ·›… ÊÕ”«» «·√’·° Ê«·⁄„— «·≈‰ «ÃÌ »«·√‘Â—° Ê«·ﬁÌ„… «·„ »ﬁÌ…° Ê„’œ— «·‘—«¡. ÌıÂ·ﬂ »«·ﬁ”ÿ «·À«»  ‘Â—Ì«° ÊÌı” »⁄œ ⁄‰œ »Ì⁄Â √Ê  ·›Â.", "[SalvageValue]+[OpeningAccumDep]<=[Cost]", "«·ﬁÌ„… «·„ »ﬁÌ… Ê«·≈Â·«ﬂ «·”«»ﬁ ·« Ì Ã«Ê“«‰ «· ﬂ·›…"
End Sub

Private Sub CreateTable_DepreciationRuns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "DepreciationRuns") Then Exit Sub
    AddField tdf, "RunID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RunNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·ﬁÌœ", ""
    AddField tdf, "RunMonth", "DATE", 0, True, "Date()", _
             "", "", "«·‘Â— (¬Œ— ÌÊ„ ›ÌÂ)", ""
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·≈Â·«ﬂ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "RunID", True, True, False
    AddIndex tdf, "UX_RunNumber", "RunNumber", False, True, False
    AddIndex tdf, "UX_RunMonth", "RunMonth", False, True, False
    EndTable tdf, "ﬁÌÊœ «·≈Â·«ﬂ «·‘Â—Ì…: ﬁÌœ ≈Â·«ﬂ ﬂ· ‘Â—: „’—Ê› «·≈Â·«ﬂ (5600) „œÌ‰ Ê„Ã„⁄ «·≈Â·«ﬂ (1790) œ«∆‰° »”ÿ— ·ﬂ· √’·.", "", ""
End Sub

Private Sub CreateTable_AssetDepreciations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "AssetDepreciations") Then Exit Sub
    AddField tdf, "DepreciationID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RunID", "LONG", 0, True, "", _
             "", "", "ﬁÌœ «·≈Â·«ﬂ", ""
    AddField tdf, "LineNo", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ— ›Ì «·ﬁÌœ", ""
    AddField tdf, "AssetID", "LONG", 0, True, "", _
             "", "", "«·√’·", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·≈Â·«ﬂ √ﬂ»— „‰ ’›—", "«·≈Â·«ﬂ", ""
    AddIndex tdf, "PrimaryKey", "DepreciationID", True, True, False
    AddIndex tdf, "UX_RunID_AssetID", "RunID,AssetID", False, True, False
    AddIndex tdf, "IX_AssetID", "AssetID", False, False, False
    EndTable tdf, "≈Â·«ﬂ ﬂ· √’· ›Ì ﬂ· ‘Â—: √”ÿ— ﬁÌœ «·≈Â·«ﬂ «·‘Â—Ì.", "", ""
End Sub

Private Sub CreateTable_CostCenters()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CostCenters") Then Exit Sub
    AddField tdf, "CostCenterID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·„—ﬂ“", ""
    AddField tdf, "CenterCode", "TEXT", 10, True, "", _
             "", "", "—„“ «·„—ﬂ“", ""
    AddField tdf, "CenterName", "TEXT", 100, True, "", _
             "", "", "«”„ «·„—ﬂ“", ""
    AddField tdf, "IsDefault", "BOOL", 0, False, "False", _
             "", "", "«·„—ﬂ“ «·«› —«÷Ì", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "CostCenterID", True, True, False
    AddIndex tdf, "UX_CenterCode", "CenterCode", False, True, False
    AddIndex tdf, "UX_CenterName", "CenterName", False, True, False
    EndTable tdf, "„—«ﬂ“ «· ﬂ·›… Ê«·›—Ê⁄: «·›—Ê⁄ √Ê «·√ﬁ”«„ («· Ã“∆…° «·„ÿ⁄„° «·„ﬁÂÏ...).  ıÊ“Û¯⁄ ⁄·ÌÂ« «·≈Ì—«œ«  Ê«·„’—Ê›«  ›Ì «·ﬁÌÊœ° Ê„‰Â« ﬁ«∆„… œŒ· ·ﬂ· „—ﬂ“.", "", ""
End Sub

Private Sub CreateTable_SalesReps()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesReps") Then Exit Sub
    AddField tdf, "SalesRepID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RepCode", "TEXT", 20, True, "", _
             "", "", "ﬂÊœ «·„‰œÊ»", ""
    AddField tdf, "RepName", "TEXT", 100, True, "", _
             "", "", "«”„ «·„‰œÊ»", ""
    AddField tdf, "RepNameEn", "TEXT", 100, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", ""
    AddField tdf, "Mobile", "TEXT", 20, False, "", _
             "", "", "«·ÃÊ«·", ""
    AddField tdf, "EmployeeID", "LONG", 0, False, "", _
             "", "", "„” Œœ„ «·»—‰«„Ã", "„»Ì⁄«  Â–« «·„” Œœ„ ·⁄„Ì· »·« „‰œÊ»  ı‰”» ·Â–« «·„‰œÊ»"
    AddField tdf, "Region", "TEXT", 50, False, "", _
             "", "", "«·„‰ÿﬁ… / Œÿ «·”Ì—", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„—ﬂ“ ﬁÌœ ⁄„Ê· Â"
    AddField tdf, "CommissionRate", "RATE", 0, True, "0", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "‰”»… «·⁄„Ê·…", ""
    AddField tdf, "CommissionBase", "TEXT", 10, True, """SALES""", _
             "In (""SALES"",""COLLECTION"")", "SALES = ’«›Ì «·„»Ì⁄« ° COLLECTION = «· Õ’Ì·", "√”«” «·⁄„Ê·…", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "SalesRepID", True, True, False
    AddIndex tdf, "UX_RepCode", "RepCode", False, True, False
    AddIndex tdf, "UX_RepName", "RepName", False, True, False
    EndTable tdf, "«·„‰œÊ»Ì‰: „‰œÊ»Ê «·„»Ì⁄« : «·⁄„·«¡ «·„”‰œÊ‰ ·ﬂ· „‰œÊ»° Ê„»Ì⁄« Â Ê Õ’Ì·« Â ÊÂœ›Â «·‘Â—Ì Ê⁄„Ê· Â.", "", ""
End Sub

Private Sub CreateTable_SalesRepTargets()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SalesRepTargets") Then Exit Sub
    AddField tdf, "TargetID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "SalesRepID", "LONG", 0, True, "", _
             "", "", "«·„‰œÊ»", ""
    AddField tdf, "TargetYear", "INT", 0, True, "", _
             "Between 2000 And 2100", "”‰… €Ì— ’ÕÌÕ…", "«·”‰…", ""
    AddField tdf, "TargetMonth", "BYTE", 0, True, "1", _
             "Between 1 And 12", "«·‘Â— „‰ 1 ≈·Ï 12", "«·‘Â—", ""
    AddField tdf, "TargetAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·Âœ›", ""
    AddIndex tdf, "PrimaryKey", "TargetID", True, True, False
    AddIndex tdf, "UX_SalesRepID_TargetYear_TargetMonth", "SalesRepID,TargetYear,TargetMonth", False, True, False
    EndTable tdf, "√Âœ«› «·„‰œÊ»Ì‰: «·Âœ› «·‘Â—Ì ·„»Ì⁄«  ﬂ· „‰œÊ» (’«›Ì «·„»Ì⁄«  »œÊ‰ «·÷—Ì»…).", "", ""
End Sub

Private Sub CreateTable_CommissionRuns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CommissionRuns") Then Exit Sub
    AddField tdf, "CommissionRunID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RunNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·„”Ì—", ""
    AddField tdf, "RunMonth", "DATE", 0, True, "Date()", _
             "", "", "«·‘Â— (¬Œ— ÌÊ„)", ""
    AddField tdf, "Status", "TEXT", 10, True, """DRAFT""", _
             "In (""DRAFT"",""POSTED"")", "DRAFT = „”Êœ…° POSTED = „—ÕÛ¯·", "«·Õ«·…", ""
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·⁄„Ê·« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "PostedAt", "DATETIME", 0, False, "", _
             "", "", " «—ÌŒ «· —ÕÌ·", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "CommissionRunID", True, True, False
    AddIndex tdf, "UX_RunNumber", "RunNumber", False, True, False
    AddIndex tdf, "UX_RunMonth", "RunMonth", False, True, False
    EndTable tdf, "„”Ì—«  «·⁄„Ê·« : ⁄„Ê·«  «·„‰œÊ»Ì‰ ·‘Â—: „”Êœ…  ı⁄œÛ¯·° À„ Ìı—ÕÛ¯· ﬁÌœÂ« („’—Ê› «·⁄„Ê·«  5530 ⁄·Ï ⁄„Ê·«  „” Õﬁ… 2330)° Ê ı’—› »”‰œ ’—› ‰ﬁœÌ… „‰ »‰œ ´’—› ⁄„Ê·… „‰œÊ»ª.", "", ""
End Sub

Private Sub CreateTable_CommissionLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CommissionLines") Then Exit Sub
    AddField tdf, "CommissionLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "CommissionRunID", "LONG", 0, True, "", _
             "", "", "«·„”Ì—", ""
    AddField tdf, "SalesRepID", "LONG", 0, True, "", _
             "", "", "«·„‰œÊ»", ""
    AddField tdf, "RepName", "TEXT", 100, False, "", _
             "", "", "«”„ «·„‰œÊ»", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", ""
    AddField tdf, "NetSales", "MONEY", 0, True, "0", _
             "", "", "’«›Ì «·„»Ì⁄« ", ""
    AddField tdf, "Collections", "MONEY", 0, True, "0", _
             "", "", "«· Õ’Ì·", ""
    AddField tdf, "CommissionBase", "TEXT", 10, True, """SALES""", _
             "", "", "«·√”«”", ""
    AddField tdf, "BaseAmount", "MONEY", 0, True, "0", _
             "", "", "„»·€ «·√”«”", ""
    AddField tdf, "CommissionRate", "RATE", 0, True, "0", _
             ">=0 And <1", "«·‰”»… ÌÃ» √‰  ﬂÊ‰ »Ì‰ 0% Ê 100%", "«·‰”»…", ""
    AddField tdf, "Adjustment", "MONEY", 0, True, "0", _
             "", "", " ⁄œÌ· (+/-)", ""
    AddField tdf, "Commission", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·⁄„Ê·…", "«·√”«” ◊ «·‰”»… + «· ⁄œÌ·° Ê·«  ﬁ· ⁄‰ ’›—"
    AddField tdf, "Notes", "TEXT", 150, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddIndex tdf, "PrimaryKey", "CommissionLineID", True, True, False
    AddIndex tdf, "UX_CommissionRunID_SalesRepID", "CommissionRunID,SalesRepID", False, True, False
    EndTable tdf, "√”ÿ— „”Ì— «·⁄„Ê·« : ⁄„Ê·… ﬂ· „‰œÊ» ›Ì «·‘Â—: ’«›Ì „»Ì⁄« Â Ê Õ’Ì·« Â° Ê«·√”«” Ê«·‰”»…° Ê«· ⁄œÌ· «·ÌœÊÌ.", "", ""
End Sub

Private Sub CreateTable_Budgets()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Budgets") Then Exit Sub
    AddField tdf, "BudgetID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "BudgetYear", "INT", 0, True, "", _
             "", "", "«·”‰…", ""
    AddField tdf, "BudgetName", "TEXT", 100, True, "", _
             "", "", "«”„ «·„Ê«“‰…", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√⁄œ¯Â«", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "BudgetID", True, True, False
    AddIndex tdf, "UX_BudgetYear", "BudgetYear", False, True, False
    EndTable tdf, "«·„Ê«“‰«  «· ﬁœÌ—Ì…: „Ê«“‰… ﬂ· ”‰…: „»·€ ‘Â—Ì ·ﬂ· Õ”«» ≈Ì—«œ«  √Ê „’—Ê›« ° ÊÌ„ﬂ‰ ·ﬂ· „—ﬂ“  ﬂ·›….", "", ""
End Sub

Private Sub CreateTable_BudgetLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "BudgetLines") Then Exit Sub
    AddField tdf, "BudgetLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "BudgetID", "LONG", 0, True, "", _
             "", "", "«·„Ê«“‰…", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "«·Õ”«»", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "›«—€ = ﬂ· «·„—«ﬂ“"
    AddField tdf, "M1", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 1", ""
    AddField tdf, "M2", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 2", ""
    AddField tdf, "M3", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 3", ""
    AddField tdf, "M4", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 4", ""
    AddField tdf, "M5", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 5", ""
    AddField tdf, "M6", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 6", ""
    AddField tdf, "M7", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 7", ""
    AddField tdf, "M8", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 8", ""
    AddField tdf, "M9", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 9", ""
    AddField tdf, "M10", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 10", ""
    AddField tdf, "M11", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 11", ""
    AddField tdf, "M12", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‘Â— 12", ""
    AddField tdf, "Notes", "TEXT", 150, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddIndex tdf, "PrimaryKey", "BudgetLineID", True, True, False
    AddIndex tdf, "IX_BudgetID", "BudgetID", False, False, False
    AddIndex tdf, "IX_AccountCode", "AccountCode", False, False, False
    EndTable tdf, "√”ÿ— «·„Ê«“‰…: Õ”«» (Ê„—ﬂ“  ﬂ·›… «Œ Ì«—Ì) Ê„»·€Â ›Ì ﬂ· ‘Â—. »·« „—ﬂ“ = «·Õ”«» ›Ì ﬂ· «·„—«ﬂ“.", "", ""
End Sub

Private Sub CreateTable_PayrollRuns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PayrollRuns") Then Exit Sub
    AddField tdf, "PayrollRunID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RunNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·„”Ì—", ""
    AddField tdf, "PayMonth", "DATE", 0, True, "Date()", _
             "", "", "«·‘Â— (¬Œ— ÌÊ„ ›ÌÂ)", ""
    AddField tdf, "Status", "TEXT", 10, True, """DRAFT""", _
             "In (""DRAFT"",""POSTED"")", "DRAFT = „”Êœ…° POSTED = „—ÕÛ¯·", "«·Õ«·…", ""
    AddField tdf, "PaidDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·’—›", ""
    AddField tdf, "PaidFrom", "TEXT", 10, False, "", _
             "Is Null Or In (""BANK"",""CASHBOX"")", "»‰ﬂ √Ê ’‰œÊﬁ", "«·’—› „‰", ""
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "«·’‰œÊﬁ", ""
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„’—Ê›", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√⁄œ¯Â", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PayrollRunID", True, True, False
    AddIndex tdf, "UX_RunNumber", "RunNumber", False, True, False
    AddIndex tdf, "UX_PayMonth", "PayMonth", False, True, False
    EndTable tdf, "„”Ì—«  «·—Ê« »: „”Ì— ﬂ· ‘Â—: „”Êœ…  ı⁄œÛ¯·° À„ Ìı—ÕÛ¯· ﬁÌœÂ ›Ì ¬Œ— ÌÊ„ „‰ «·‘Â—° À„ Ìı”ÃÛ¯· ’—›Â „‰ »‰ﬂ √Ê ’‰œÊﬁ.", "", ""
End Sub

Private Sub CreateTable_PayrollLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PayrollLines") Then Exit Sub
    AddField tdf, "PayrollLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "PayrollRunID", "LONG", 0, True, "", _
             "", "", "«·„”Ì—", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "EmployeeName", "TEXT", 100, False, "", _
             "", "", "«”„ «·„ÊŸ›", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„—ﬂ“ «·„ÊŸ› ÌÊ„ ≈‰‘«¡ «·„”Ì—"
    AddField tdf, "IsSaudi", "BOOL", 0, False, "False", _
             "", "", "”⁄ÊœÌ", ""
    AddField tdf, "Basic", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·√”«”Ì", ""
    AddField tdf, "Housing", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»œ· «·”ﬂ‰", ""
    AddField tdf, "OtherAllow", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»œ·«  √Œ—Ï («·‰ﬁ· Ê€Ì—Â)", ""
    AddField tdf, "Overtime", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·⁄„· «·≈÷«›Ì", ""
    AddField tdf, "Additions", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„ﬂ«›¬  Ê≈÷«›« ", ""
    AddField tdf, "AbsenceDeduction", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "Œ’„ «·€Ì«»", ""
    AddField tdf, "AdvanceDeduction", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "Œ’„ «·”·›…", ""
    AddField tdf, "OtherDeduction", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "Œ’Ê„«  √Œ—Ï (Ã“«¡« )", ""
    AddField tdf, "GosiWage", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·√Ã— «·Œ«÷⁄ ·· √„Ì‰« ", ""
    AddField tdf, "GosiEmployee", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«· √„Ì‰«  - Õ’… «·„ÊŸ›", ""
    AddField tdf, "GosiEmployer", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«· √„Ì‰«  - Õ’… «·„‰‘√…", ""
    AddField tdf, "NetPay", "MONEY", 0, True, "0", _
             "", "", "’«›Ì «·—« »", "ﬁœ ÌﬂÊ‰ ”«·»« ›Ì «·„”Êœ…∫ «· —ÕÌ· Ì—›÷Â («·Œ’Ê„«  √ﬂ»— „‰ «·—« »)"
    AddField tdf, "Notes", "TEXT", 150, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddIndex tdf, "PrimaryKey", "PayrollLineID", True, True, False
    AddIndex tdf, "UX_PayrollRunID_EmployeeID", "PayrollRunID,EmployeeID", False, True, False
    EndTable tdf, "√”ÿ— „”Ì— «·—Ê« »: ”ÿ— ﬂ· „ÊŸ›: «·—« » Ê«·»œ·«  Ê«·≈÷«›Ì Ê«·Œ’Ê„«  Ê«· √„Ì‰«  Ê’«›Ì «·—« ».", "", ""
End Sub

Private Sub CreateTable_BankReconciliations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "BankReconciliations") Then Exit Sub
    AddField tdf, "ReconciliationID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ReconNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «· ”ÊÌ…", ""
    AddField tdf, "BankID", "LONG", 0, True, "", _
             "", "", "«·»‰ﬂ", ""
    AddField tdf, "StatementDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ ﬂ‘› «·»‰ﬂ", ""
    AddField tdf, "StatementBalance", "MONEY", 0, True, "0", _
             "", "", "—’Ìœ ﬂ‘› «·»‰ﬂ", ""
    AddField tdf, "BookBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ ›Ì «·œ›« — ›Ì «· «—ÌŒ", ""
    AddField tdf, "Outstanding", "MONEY", 0, True, "0", _
             "", "", "Õ—ﬂ«  ·„  ŸÂ— ›Ì «·ﬂ‘› (’«›Ì)", ""
    AddField tdf, "Difference", "MONEY", 0, True, "0", _
             "", "", "«·›—ﬁ", ""
    AddField tdf, "Status", "TEXT", 10, True, """OPEN""", _
             "In (""OPEN"",""DONE"")", "OPEN = Ã«—Ì…° DONE = „⁄ „œ…", "«·Õ«·…", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "ReconciliationID", True, True, False
    AddIndex tdf, "UX_ReconNumber", "ReconNumber", False, True, False
    AddIndex tdf, "IX_BankID", "BankID", False, False, False
    EndTable tdf, "«· ”ÊÌ«  «·»‰ﬂÌ…: „ÿ«»ﬁ… ﬂ‘› «·»‰ﬂ ›Ì  «—ÌŒ „⁄ «·œ›« —: —’Ìœ «·ﬂ‘›° Ê«·—’Ìœ ›Ì «·œ›« —° Ê«·Õ—ﬂ«  €Ì— «·Ÿ«Â—… ›Ì «·ﬂ‘›.", "", ""
End Sub

Private Sub CreateTable_BankClearings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "BankClearings") Then Exit Sub
    AddField tdf, "ClearingID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ReconciliationID", "LONG", 0, True, "", _
             "", "", "«· ”ÊÌ…", ""
    AddField tdf, "BankID", "LONG", 0, True, "", _
             "", "", "«·»‰ﬂ", ""
    AddField tdf, "SourceType", "TEXT", 20, True, "", _
             "", "", "‰Ê⁄ «·⁄„·Ì…", ""
    AddField tdf, "SourceID", "LONG", 0, True, "", _
             "", "", "—ﬁ„ «·⁄„·Ì…", ""
    AddField tdf, "ClearedAmount", "MONEY", 0, True, "0", _
             "", "", "«·„»·€ ÌÊ„ «·„ÿ«»ﬁ… („œÌ‰ „ÊÃ»)", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "ClearingID", True, True, False
    AddIndex tdf, "UX_BankID_SourceType_SourceID", "BankID,SourceType,SourceID", False, True, False
    EndTable tdf, "Õ—ﬂ«  «·œ›« — «·„ÿ«»ﬁ… ·ﬂ‘› «·»‰ﬂ: ﬂ· ⁄„·Ì… ﬁÌœÂ« ⁄·Ï Õ”«» «·»‰ﬂ ŸÂ—  ›Ì ﬂ‘› «·»‰ﬂ: ‰Ê⁄ «·⁄„·Ì… Ê—ﬁ„Â« Ê„»·€Â« ÌÊ„ «·„ÿ«»ﬁ….", "", ""
End Sub

Private Sub CreateTable_CustomerAllocations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CustomerAllocations") Then Exit Sub
    AddField tdf, "AllocationID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "PaymentID", "LONG", 0, True, "", _
             "", "", "”‰œ «·ﬁ»÷", ""
    AddField tdf, "SalesInvoiceID", "LONG", 0, True, "", _
             "", "", "«·›« Ê—…", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "AllocationID", True, True, False
    AddIndex tdf, "UX_PaymentID_SalesInvoiceID", "PaymentID,SalesInvoiceID", False, True, False
    AddIndex tdf, "IX_SalesInvoiceID", "SalesInvoiceID", False, False, False
    EndTable tdf, "—»ÿ ”‰œ«  «·ﬁ»÷ »«·›Ê« Ì—: ﬂ„ „‰ ”‰œ «·ﬁ»÷ ”œ¯œ ﬂ· ›« Ê—… ¬Ã·…. „« ·« Ìı—»ÿ »›« Ê—… Ì”œœ √ﬁœ„ «·›Ê« Ì— «” Õﬁ«ﬁ«.", "", ""
End Sub

Private Sub CreateTable_SupplierAllocations()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "SupplierAllocations") Then Exit Sub
    AddField tdf, "AllocationID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "PaymentID", "LONG", 0, True, "", _
             "", "", "”‰œ «·’—›", ""
    AddField tdf, "PurchaseInvoiceID", "LONG", 0, True, "", _
             "", "", "«·›« Ê—…", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "AllocationID", True, True, False
    AddIndex tdf, "UX_PaymentID_PurchaseInvoiceID", "PaymentID,PurchaseInvoiceID", False, True, False
    AddIndex tdf, "IX_PurchaseInvoiceID", "PurchaseInvoiceID", False, False, False
    EndTable tdf, "—»ÿ ”‰œ«  «·’—› »›Ê« Ì— «·‘—«¡: ﬂ„ „‰ ”‰œ «·’—› ”œ¯œ ﬂ· ›« Ê—… ‘—«¡ ¬Ã·…. „« ·« Ìı—»ÿ »›« Ê—… Ì”œœ √ﬁœ„ «·›Ê« Ì— «” Õﬁ«ﬁ«.", "", ""
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
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "’ı—› „‰ ’‰œÊﬁ", "«·„’—Ê› «·‰ﬁœÌ ÌŒ—Ã „‰ Â–« «·’‰œÊﬁ∫ ›«—€ = ·„ Ìıœ›⁄ „‰ ’‰œÊﬁ"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "«·»‰ﬂ", "«·„»·€ «·„ÕÊÛ¯· »‰ﬂÌ« ÌıﬁÌÛ¯œ ›Ì Õ”«» Â–« «·»‰ﬂ"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", ""
    AddField tdf, "RecurringID", "LONG", 0, False, "", _
             "", "", "„‰ „’—Ê› „ ﬂ——", "√‰‘√Â «·»—‰«„Ã „‰ «·„’—Ê› «·„ ﬂ—— » «—ÌŒ «” Õﬁ«ﬁÂ"
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ »«·⁄„·… (»œÊ‰ «·÷—Ì»…)", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
    AddField tdf, "ForeignTax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·÷—Ì»… »«·⁄„·…", ""
    AddIndex tdf, "PrimaryKey", "ExpenseID", True, True, False
    AddIndex tdf, "UX_ExpenseNumber", "ExpenseNumber", False, True, False
    AddIndex tdf, "IX_ExpenseDate", "ExpenseDate", False, False, False
    EndTable tdf, "«·„’—Ê›« : „’—Ê›«  «·„Õ· «· ‘€Ì·Ì…∫ «·„»·€ »œÊ‰ ÷—Ì»… Ê«·÷—Ì»… „‰›’·… (÷—Ì»… „œŒ·« ).", "[TotalAmount]=[Amount]+[Tax]", "«·≈Ã„«·Ì = «·„»·€ + «·÷—Ì»…"
End Sub

Private Sub CreateTable_RecurringExpenses()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "RecurringExpenses") Then Exit Sub
    AddField tdf, "RecurringID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "RecurringName", "TEXT", 100, True, "", _
             "", "", "«”„ «·„’—Ê›", ""
    AddField tdf, "ExpenseTypeID", "LONG", 0, True, "", _
             "", "", "‰Ê⁄ «·„’—Ê›", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€ ﬁ»· «·÷—Ì»…", ""
    AddField tdf, "Tax", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "÷—Ì»… «·„œŒ·« ", ""
    AddField tdf, "PaymentMethodID", "LONG", 0, True, "", _
             "", "", "ÿ—Ìﬁ… «·œ›⁄", ""
    AddField tdf, "CashBoxID", "LONG", 0, False, "", _
             "", "", "„‰ ’‰œÊﬁ", "··œ›⁄ «·‰ﬁœÌ∫ ›«—€ = ’‰œÊﬁ „‰ Ìı‰‘∆ «·„’—Ê›"
    AddField tdf, "BankID", "LONG", 0, False, "", _
             "", "", "„‰ »‰ﬂ", "·· ÕÊÌ· «·»‰ﬂÌ∫ ›«—€ = «·»‰ﬂ «·«› —«÷Ì"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", ""
    AddField tdf, "Frequency", "TEXT", 10, True, """MONTHLY""", _
             "In (""MONTHLY"",""QUARTERLY"",""YEARLY"")", "MONTHLY = ‘Â—Ì° QUARTERLY = ﬂ· 3 √‘Â—° YEARLY = ”‰ÊÌ", "«· ﬂ—«—", ""
    AddField tdf, "DueDay", "INT", 0, True, "1", _
             "Between 1 And 28", "ÌÊ„ «·«” Õﬁ«ﬁ „‰ 1 ≈·Ï 28", "ÌÊ„ «·«” Õﬁ«ﬁ", ""
    AddField tdf, "StartDate", "DATE", 0, True, "Date()", _
             "", "", "Ì»œ√ „‰", ""
    AddField tdf, "EndDate", "DATE", 0, False, "", _
             "", "", "Ì‰ ÂÌ ›Ì", ""
    AddField tdf, "NextDueDate", "DATE", 0, False, "", _
             "", "", "«·«” Õﬁ«ﬁ «· «·Ì", ""
    AddField tdf, "LastCreatedDate", "DATE", 0, False, "", _
             "", "", "¬Œ— „’—Ê› √ı‰‘∆", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·Ê’›", ""
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "RecurringID", True, True, False
    AddIndex tdf, "UX_RecurringName", "RecurringName", False, True, False
    AddIndex tdf, "IX_NextDueDate", "NextDueDate", False, False, False
    EndTable tdf, "«·„’—Ê›«  «·„ ﬂ——…: „’—Ê› À«»  Ì ﬂ—— («·≈ÌÃ«—° «·ﬂÂ—»«¡° «·«‘ —«ﬂ« ): Ìı‰‘∆ «·»—‰«„Ã «·„’—Ê› ›Ì  «—ÌŒ «” Õﬁ«ﬁÂ° „—… Ê«Õœ… ·ﬂ·  «—ÌŒ.", "", ""
End Sub

Private Sub CreateTable_CashVouchers()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashVouchers") Then Exit Sub
    AddField tdf, "CashVoucherID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "VoucherNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·”‰œ", ""
    AddField tdf, "VoucherDate", "DATETIME", 0, True, "Now()", _
             "", "", "«· «—ÌŒ", ""
    AddField tdf, "VoucherType", "TEXT", 10, True, "", _
             "In (""IN"",""OUT"",""TRANSFER"")", "IN = ﬁ»÷° OUT = ’—›° TRANSFER =  ÕÊÌ· »Ì‰ ’‰œÊﬁÌ‰", "‰Ê⁄ «·”‰œ", ""
    AddField tdf, "CashBoxID", "LONG", 0, True, "", _
             "", "", "«·’‰œÊﬁ", "«·ﬁ»÷ ÌœŒ·Â° Ê«·’—› Ê«· ÕÊÌ· ÌŒ—Ã«‰ „‰Â"
    AddField tdf, "ToCashBoxID", "LONG", 0, False, "", _
             "", "", "≈·Ï ’‰œÊﬁ", "·· ÕÊÌ· ›ﬁÿ"
    AddField tdf, "Category", "TEXT", 10, True, """OTHER""", _
             "In (""OTHER"",""OWNER"",""EXPENSE"",""ADVANCE"",""SHORTAGE"",""OVERAGE"",""TRANSFER"",""COMMISSION"")", "«Œ — «·»‰œ „‰ «·ﬁ«∆„…", "«·»‰œ", ""
    AddField tdf, "Amount", "MONEY", 0, True, "0", _
             ">0", "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "«·„»·€", ""
    AddField tdf, "PartyName", "TEXT", 100, False, "", _
             "", "", "«·„” ·„ / «·„”·ˆ¯„", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·»Ì«‰", ""
    AddField tdf, "ExpenseID", "LONG", 0, False, "", _
             "", "", "«·„’—Ê› «·„”ÃÛ¯·", "’—› »‰œ „’—Ê› Ì”Ã· „’—Ê›« »‰›” «·„»·€ ›Ì «·„’—Ê›« "
    AddField tdf, "ClosingID", "LONG", 0, False, "", _
             "", "", " ’›Ì… «·ﬂ«‘Ì—", ""
    AddField tdf, "AdvanceEmployeeID", "LONG", 0, False, "", _
             "", "", "«·„ÊŸ› ’«Õ» «·”·›…", "”·›… „ÊŸ› (’—›) √Ê ”œ«œÂ« ‰ﬁœ« (ﬁ»÷)"
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "«·„ÊŸ›", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "SalesRepID", "LONG", 0, False, "", _
             "", "", "«·„‰œÊ»", "’—› ⁄„Ê·… «·„‰œÊ»"
    AddIndex tdf, "PrimaryKey", "CashVoucherID", True, True, False
    AddIndex tdf, "UX_VoucherNumber", "VoucherNumber", False, True, False
    AddIndex tdf, "IX_VoucherDate", "VoucherDate", False, False, False
    AddIndex tdf, "IX_CashBoxID", "CashBoxID", False, False, False
    EndTable tdf, "”‰œ«  «·‰ﬁœÌ…: ﬁ»÷ ‰ﬁœÌ… ·’‰œÊﬁ° √Ê ’—› „‰Â° √Ê  ÕÊÌ· »Ì‰ ’‰œÊﬁÌ‰ (Ê„‰Â«  —ÕÌ· ÌÊ„Ì… «·ﬂ«‘Ì—).", "[VoucherType]<>""TRANSFER"" Or ([ToCashBoxID] Is Not Null And [ToCashBoxID]<>[CashBoxID])", "«· ÕÊÌ· ÌÕ «Ã ’‰œÊﬁ« ¬Œ— €Ì— ’‰œÊﬁ «·’—›"
End Sub

Private Sub CreateTable_CashClosings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "CashClosings") Then Exit Sub
    AddField tdf, "ClosingID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ClosingNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «· ’›Ì…", ""
    AddField tdf, "ClosingDate", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «· ’›Ì…", ""
    AddField tdf, "CashBoxID", "LONG", 0, True, "", _
             "", "", "«·’‰œÊﬁ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√Ã—«Â«", ""
    AddField tdf, "PeriodStart", "DATETIME", 0, False, "", _
             "", "", "„‰ (¬Œ—  ’›Ì…)", ""
    AddField tdf, "OpeningBalance", "MONEY", 0, True, "0", _
             "", "", "—’Ìœ «·»œ«Ì…", ""
    AddField tdf, "CashIn", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„ﬁ»Ê÷« ", ""
    AddField tdf, "CashOut", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„œ›Ê⁄« ", ""
    AddField tdf, "ExpectedBalance", "MONEY", 0, True, "0", _
             "", "", "«·—’Ìœ «·œ› —Ì", ""
    AddField tdf, "CountedAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·‰ﬁœÌ… «·›⁄·Ì…", ""
    AddField tdf, "Difference", "MONEY", 0, True, "0", _
             "", "", "«·›—ﬁ", "”«·» = ⁄Ã“° „ÊÃ» = “Ì«œ…"
    AddField tdf, "Destination", "TEXT", 10, True, """MAIN""", _
             "In (""MAIN"",""OWNER"",""KEEP"")", "MAIN = «·Œ“Ì‰… «·—∆Ì”Ì…° OWNER =  ”ÊÌ… „⁄ «·„«·ﬂ° KEEP = Ì»ﬁÏ ›Ì «·’‰œÊﬁ", "«· —ÕÌ· ≈·Ï", ""
    AddField tdf, "ToCashBoxID", "LONG", 0, False, "", _
             "", "", "«·Œ“Ì‰… «·„” ·„…", ""
    AddField tdf, "TransferAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„—ÕÛ¯·", ""
    AddField tdf, "KeptAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„ »ﬁÌ ›Ì «·’‰œÊﬁ (⁄Âœ…)", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "ClosingID", True, True, False
    AddIndex tdf, "UX_ClosingNumber", "ClosingNumber", False, True, False
    AddIndex tdf, "IX_CashBoxID_ClosingDate", "CashBoxID,ClosingDate", False, False, False
    EndTable tdf, " ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—: Ã—œ ‰ﬁœÌ… ’‰œÊﬁ «·ﬂ«‘Ì— ›Ì ‰Â«Ì… «·Ê—œÌ… Ê —ÕÌ·Â« ··Œ“Ì‰… «·—∆Ì”Ì… √Ê  ”ÊÌ Â« „⁄ «·„«·ﬂ.", "[TransferAmount]+[KeptAmount]=[CountedAmount]", "«·„—ÕÛ¯· + «·„ »ﬁÌ = «·‰ﬁœÌ… «·›⁄·Ì…"
End Sub

Private Sub CreateTable_Accounts()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "Accounts") Then Exit Sub
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             ">0", "—ﬁ„ «·Õ”«» √ﬂ»— „‰ ’›—", "—ﬁ„ «·Õ”«»", ""
    AddField tdf, "AccountName", "TEXT", 100, True, "", _
             "", "", "«”„ «·Õ”«»", ""
    AddField tdf, "AccountNameEn", "TEXT", 100, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", "ÌŸÂ— ›Ì «·Ê«ÃÂ… «·≈‰Ã·Ì“Ì… Ê«·ﬁÊ«∆„ «·„«·Ì… »Â«"
    AddField tdf, "AccountType", "TEXT", 10, True, "", _
             "In (""ASSET"",""LIABILITY"",""EQUITY"",""REVENUE"",""EXPENSE"")", "√’Ê·° Œ’Ê„° ÕﬁÊﬁ „·ﬂÌ…° ≈Ì—«œ« ° „’—Ê›« ", "‰Ê⁄ «·Õ”«»", ""
    AddField tdf, "ParentCode", "LONG", 0, False, "", _
             "", "", "«·Õ”«» «·—∆Ì”Ì", "›«—€ ··Õ”«»«  «·Œ„”… ›Ì «·„” ÊÏ «·√Ê· ›ﬁÿ"
    AddField tdf, "IsActive", "BOOL", 0, False, "True", _
             "", "", "‰‘ÿ", ""
    AddField tdf, "IsPosting", "BOOL", 0, False, "True", _
             "", "", "Õ”«» ›—⁄Ì (Ìﬁ»· «·ﬁÌÊœ)", ""
    AddField tdf, "IsSystem", "BOOL", 0, False, "False", _
             "", "", "Õ”«» √”«”Ì ›Ì «·‰Ÿ«„", ""
    AddField tdf, "AccountLevel", "BYTE", 0, True, "1", _
             "", "", "«·„” ÊÏ", ""
    AddField tdf, "TreeKey", "TEXT", 60, False, "", _
             "", "", "„› «Õ «· — Ì» ›Ì «·‘Ã—…", "ÌÕ”»Â «·»—‰«„Ã (modAccounts.RebuildAccountTree): —ﬁ„ ﬂ· „” ÊÏ »⁄‘— Œ«‰« "
    AddField tdf, "Level1Code", "LONG", 0, False, "", _
             "", "", "Õ”«» «·„” ÊÏ 1", ""
    AddField tdf, "Level2Code", "LONG", 0, False, "", _
             "", "", "Õ”«» «·„” ÊÏ 2", ""
    AddField tdf, "Level3Code", "LONG", 0, False, "", _
             "", "", "Õ”«» «·„” ÊÏ 3", ""
    AddField tdf, "Level4Code", "LONG", 0, False, "", _
             "", "", "Õ”«» «·„” ÊÏ 4", ""
    AddField tdf, "Level5Code", "LONG", 0, False, "", _
             "", "", "Õ”«» «·„” ÊÏ 5", ""
    AddIndex tdf, "PrimaryKey", "AccountCode", True, True, False
    AddIndex tdf, "IX_ParentCode", "ParentCode", False, False, False
    AddIndex tdf, "IX_TreeKey", "TreeKey", False, False, False
    EndTable tdf, "œ·Ì· «·Õ”«»«  (‘Ã—… «·Õ”«»« ): ‘Ã—… „‰ Œ„”… „” ÊÌ«  ⁄·Ï «·√ﬂÀ—: «·Õ”«»«  «·—∆Ì”Ì… ( Ã„Ì⁄Ì…) Ê«·Õ”«»«  «·›—⁄Ì… «· Ì  ı—ÕÛ¯· ≈·ÌÂ« «·ﬁÌÊœ. Õ”«»«  «·’‰«œÌﬁ (110000 + —ﬁ„ «·’‰œÊﬁ) Ê√‰Ê«⁄ «·„’—Ê›«  (530000 + —ﬁ„ «·‰Ê⁄)  ı‰‘√  ·ﬁ«∆Ì«.", "", ""
End Sub

Private Sub CreateTable_JournalSourceTypes()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalSourceTypes") Then Exit Sub
    AddField tdf, "SourceType", "TEXT", 20, True, "", _
             "", "", "‰Ê⁄ «·⁄„·Ì…", ""
    AddField tdf, "TypeName", "TEXT", 50, True, "", _
             "", "", "«·«”„", ""
    AddField tdf, "TypeNameEn", "TEXT", 50, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", ""
    AddField tdf, "SortOrder", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddIndex tdf, "PrimaryKey", "SourceType", True, True, False
    EndTable tdf, "√‰Ê«⁄ „’«œ— «·ﬁÌÊœ: √‰Ê«⁄ «·⁄„·Ì«  «· Ì Ìı‰‘√ ⁄‰Â« ﬁÌœ ¬·Ì° Ê„‰Â« Ìı⁄—› √’· «·ﬁÌœ.", "", ""
End Sub

Private Sub CreateTable_JournalEntries()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalEntries") Then Exit Sub
    AddField tdf, "EntryID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "EntryNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·ﬁÌœ", ""
    AddField tdf, "EntryDate", "DATETIME", 0, True, "", _
             "", "", " «—ÌŒ «·ﬁÌœ", ""
    AddField tdf, "SourceType", "TEXT", 20, True, "", _
             "", "", "‰Ê⁄ «·⁄„·Ì…", ""
    AddField tdf, "SourceID", "LONG", 0, True, "", _
             "", "", "—ﬁ„ «·⁄„·Ì… «·œ«Œ·Ì", ""
    AddField tdf, "SourceNumber", "TEXT", 20, False, "", _
             "", "", "—ﬁ„ „” ‰œ «·⁄„·Ì…", ""
    AddField tdf, "Description", "TEXT", 255, False, "", _
             "", "", "«·»Ì«‰", ""
    AddField tdf, "TotalDebit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·„œÌ‰", ""
    AddField tdf, "TotalCredit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·œ«∆‰", ""
    AddField tdf, "LineCount", "INT", 0, True, "0", _
             "", "", "⁄œœ «·√”ÿ—", ""
    AddField tdf, "Signature", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "»’„… «·ﬁÌœ", " ﬂ‘›  €Ì¯— «·⁄„·Ì… »⁄œ ≈‰‘«¡ «·ﬁÌœ"
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "¬Œ—  ÕœÌÀ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„»·€ «·„” ‰œ »⁄„· Â", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
    AddIndex tdf, "PrimaryKey", "EntryID", True, True, False
    AddIndex tdf, "UX_EntryNumber", "EntryNumber", False, True, False
    AddIndex tdf, "UX_SourceType_SourceID", "SourceType,SourceID", False, True, False
    AddIndex tdf, "IX_EntryDate", "EntryDate", False, False, False
    EndTable tdf, "ﬁÌÊœ «·ÌÊ„Ì…: ﬁÌœ ¬·Ì ·ﬂ· ⁄„·Ì…° „—»Êÿ »√’·Â« (SourceType + SourceID). ÌıÕœÛ¯À ≈–«  €Ì—  «·⁄„·Ì….", "[TotalDebit]=[TotalCredit]", "«·ﬁÌœ €Ì— „ Ê«“‰: «·„œÌ‰ ÌÃ» √‰ Ì”«ÊÌ «·œ«∆‰"
End Sub

Private Sub CreateTable_JournalLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "JournalLines") Then Exit Sub
    AddField tdf, "JournalLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "EntryID", "LONG", 0, True, "", _
             "", "", "«·ﬁÌœ", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ—", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "«·Õ”«»", ""
    AddField tdf, "Debit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„œÌ‰", ""
    AddField tdf, "Credit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "œ«∆‰", ""
    AddField tdf, "LineText", "TEXT", 255, False, "", _
             "", "", "«·»Ì«‰", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", "„‰ „” ‰œ «·⁄„·Ì…∫ ›«—€ = €Ì— „Ê“⁄"
    AddIndex tdf, "PrimaryKey", "JournalLineID", True, True, False
    AddIndex tdf, "UX_EntryID_LineNumber", "EntryID,LineNumber", False, True, False
    AddIndex tdf, "IX_AccountCode", "AccountCode", False, False, False
    EndTable tdf, "√”ÿ— «·ﬁÌÊœ: «·ÿ—› «·„œÌ‰ Ê«·ÿ—› «·œ«∆‰ ·ﬂ· ﬁÌœ.", "", ""
End Sub

Private Sub CreateTable_PeriodClosings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "PeriodClosings") Then Exit Sub
    AddField tdf, "PeriodClosingID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ActionType", "TEXT", 12, True, "", _
             "In (""CLOSE"",""REOPEN"",""YEAR_CLOSE"",""YEAR_OPEN"")", "≈ﬁ›«·° ≈⁄«œ… › Õ° ≈ﬁ›«· ”‰…° ≈⁄«œ… › Õ ”‰…", "«·⁄„·Ì…", ""
    AddField tdf, "ClosedThrough", "DATETIME", 0, False, "", _
             "", "", "„ﬁ›·… Õ Ï (»⁄œ «·⁄„·Ì…)", ""
    AddField tdf, "PreviousThrough", "DATETIME", 0, False, "", _
             "", "", "„ﬁ›·… Õ Ï (ﬁ»· «·⁄„·Ì…)", ""
    AddField tdf, "FiscalYear", "INT", 0, False, "", _
             "", "", "«·”‰… «·„«·Ì…", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "«·”»» / „·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "ﬁ«„ »Â", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "PeriodClosingID", True, True, False
    EndTable tdf, "”Ã· ≈ﬁ›«· «·› —« : ﬂ· ≈ﬁ›«· √Ê ≈⁄«œ… › Õ ·› —… √Ê ”‰… „«·Ì…: «· «—ÌŒ «·ÃœÌœ ··≈ﬁ›«· Ê«·”«»ﬁ° Ê„‰ ﬁ«„ »Â Ê«·”»».", "", ""
End Sub

Private Sub CreateTable_FiscalYearClosings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "FiscalYearClosings") Then Exit Sub
    AddField tdf, "YearClosingID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "FiscalYear", "INT", 0, True, "", _
             "", "", "«·”‰… «·„«·Ì…", ""
    AddField tdf, "ClosingNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·≈ﬁ›«·", ""
    AddField tdf, "ClosingDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·≈ﬁ›«·", ""
    AddField tdf, "NetProfit", "MONEY", 0, True, "0", _
             "", "", "’«›Ì —»Õ (Œ”«—…) «·”‰…", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√ﬁ›·Â«", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "YearClosingID", True, True, False
    AddIndex tdf, "UX_FiscalYear", "FiscalYear", False, True, False
    AddIndex tdf, "UX_ClosingNumber", "ClosingNumber", False, True, False
    EndTable tdf, "≈ﬁ›«· «·”‰Ê«  «·„«·Ì…: ﬁÌœ ≈ﬁ›«· ﬂ· ”‰…: √—’œ… «·≈Ì—«œ«  Ê«·„’—Ê›«  ›Ì 31 œÌ”„»—  ıﬁ›· ›Ì «·√—»«Õ «·„Õ Ã“… (3300). √”ÿ—Â „Õ›ÊŸ… ﬂ„« ﬂ«‰  ÌÊ„ «·≈ﬁ›«·° ÊÌı—ÕÛ¯· ··ÌÊ„Ì… ﬂ⁄„·Ì… ´ﬁÌœ ≈ﬁ›«· «·”‰…ª.", "", ""
End Sub

Private Sub CreateTable_FiscalYearClosingLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "FiscalYearClosingLines") Then Exit Sub
    AddField tdf, "YearClosingLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "YearClosingID", "LONG", 0, True, "", _
             "", "", "≈ﬁ›«· «·”‰…", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ—", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "«·Õ”«»", ""
    AddField tdf, "Debit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„œÌ‰", ""
    AddField tdf, "Credit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "œ«∆‰", ""
    AddField tdf, "LineText", "TEXT", 150, False, "", _
             "", "", "«·»Ì«‰", ""
    AddIndex tdf, "PrimaryKey", "YearClosingLineID", True, True, False
    AddIndex tdf, "UX_YearClosingID_LineNumber", "YearClosingID,LineNumber", False, True, False
    EndTable tdf, "√”ÿ— ﬁÌÊœ ≈ﬁ›«· «·”‰Ê« : ·ﬂ· Õ”«» ≈Ì—«œ«  √Ê „’—Ê›«  —’ÌœÂ „⁄ﬂÊ”«° À„ ’«›Ì «·—»Õ ›Ì «·√—»«Õ «·„Õ Ã“….", "", ""
End Sub

Private Sub CreateTable_VatReturns()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "VatReturns") Then Exit Sub
    AddField tdf, "VatReturnID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "ReturnNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·≈ﬁ—«—", "VAT-yyyymmdd (¬Œ— ÌÊ„ ›Ì «·› —…)"
    AddField tdf, "PeriodFrom", "DATE", 0, True, "", _
             "", "", "»œ«Ì… «·› —…", ""
    AddField tdf, "PeriodTo", "DATE", 0, True, "", _
             "", "", "‰Â«Ì… «·› —…", ""
    AddField tdf, "Status", "TEXT", 10, True, """DRAFT""", _
             "In (""DRAFT"",""FILED"")", "DRAFT = „”Êœ…° FILED = „⁄ „œ", "«·Õ«·…", ""
    AddField tdf, "SalesStdAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "1- «·„»Ì⁄«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì…", ""
    AddField tdf, "SalesStdAdjust", "MONEY", 0, True, "0", _
             "", "", "1-  ⁄œÌ·«  („— Ã⁄« ) «·„»Ì⁄«  «·Œ«÷⁄…", ""
    AddField tdf, "SalesStdVAT", "MONEY", 0, True, "0", _
             "", "", "1- ÷—Ì»… «·„»Ì⁄«  «·Œ«÷⁄…", ""
    AddField tdf, "SalesZeroAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "3- «·„»Ì⁄«  »‰”»… ’›—Ì…", ""
    AddField tdf, "SalesZeroAdjust", "MONEY", 0, True, "0", _
             "", "", "3-  ⁄œÌ·«  «·„»Ì⁄«  »‰”»… ’›—Ì…", ""
    AddField tdf, "SalesExemptAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "5- «·„»Ì⁄«  «·„⁄›«…", ""
    AddField tdf, "SalesExemptAdjust", "MONEY", 0, True, "0", _
             "", "", "5-  ⁄œÌ·«  «·„»Ì⁄«  «·„⁄›«…", ""
    AddField tdf, "PurchStdAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "7- «·„‘ —Ì«  Ê«·„’—Ê›«  «·Œ«÷⁄… ··‰”»… «·√”«”Ì…", ""
    AddField tdf, "PurchStdAdjust", "MONEY", 0, True, "0", _
             "", "", "7-  ⁄œÌ·«  („— Ã⁄« ) «·„‘ —Ì«  «·Œ«÷⁄…", ""
    AddField tdf, "PurchStdVAT", "MONEY", 0, True, "0", _
             "", "", "7- ÷—Ì»… «·„‘ —Ì«  «·Œ«÷⁄…", ""
    AddField tdf, "PurchZeroAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "10- «·„‘ —Ì«  »‰”»… ’›—Ì…", ""
    AddField tdf, "PurchZeroAdjust", "MONEY", 0, True, "0", _
             "", "", "10-  ⁄œÌ·«  «·„‘ —Ì«  »‰”»… ’›—Ì…", ""
    AddField tdf, "Corrections", "MONEY", 0, True, "0", _
             "", "", "14-  ’ÕÌÕ«  „‰ «·› —«  «·”«»ﬁ…", "„ÊÃ»…  “Ìœ «·÷—Ì»… «·„” Õﬁ…° ”«·»…  ‰ﬁ’Â«"
    AddField tdf, "CarriedCredit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "15- «·—’Ìœ «·œ«∆‰ «·„—ÕÛ¯· „‰ «·› —«  «·”«»ﬁ…", ""
    AddField tdf, "NetDue", "MONEY", 0, True, "0", _
             "", "", "16- ’«›Ì «·÷—Ì»… «·„” Õﬁ… (”«·» = „” —œ…)", "SalesStdVAT - PurchStdVAT + Corrections - CarriedCredit"
    AddField tdf, "FiledDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·«⁄ „«œ ( «—ÌŒ ﬁÌœ «· ”ÊÌ…)", ""
    AddField tdf, "FilingRef", "TEXT", 30, False, "", _
             "", "", "—ﬁ„ «·≈ﬁ—«— ·œÏ «·ÂÌ∆…", ""
    AddField tdf, "PaidDate", "DATE", 0, False, "", _
             "", "", " «—ÌŒ «·”œ«œ", ""
    AddField tdf, "PaidAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "«·„»·€ «·„”œœ", ""
    AddField tdf, "PaidAccount", "LONG", 0, False, "", _
             "", "", "Õ”«» «·”œ«œ («·»‰ﬂ)", ""
    AddField tdf, "Notes", "TEXT", 255, False, "", _
             "", "", "„·«ÕŸ« ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√⁄œ¯Â / «⁄ „œÂ", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddIndex tdf, "PrimaryKey", "VatReturnID", True, True, False
    AddIndex tdf, "UX_ReturnNumber", "ReturnNumber", False, True, False
    AddIndex tdf, "IX_PeriodFrom", "PeriodFrom", False, False, False
    EndTable tdf, "≈ﬁ—«—«  ÷—Ì»… «·ﬁÌ„… «·„÷«›…: ≈ﬁ—«— ﬂ· › —… ÷—Ì»Ì… (‘Â— √Ê —»⁄ ”‰…) »Œ«‰«  ‰„Ê–Ã ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ. «·„”Êœ…  ıÕ”» „‰ «·„” ‰œ« ° Ê⁄‰œ «·«⁄ „«œ  ıÕ›Ÿ ﬁÌ„Â« ﬂ„« ÂÌ ÊÌı‰‘√ ﬁÌœ «· ”ÊÌ… (÷—Ì»… «·„Œ—Ã«  Ê«·„œŒ·«  ≈·Ï Õ”«» «· ”ÊÌ… 2250)° À„ ﬁÌœ «·”œ«œ ⁄‰œ  ”ÃÌ·Â.", "[PeriodTo]>=[PeriodFrom] And ([PaidAmount]=0 Or [PaidAmount]<=[NetDue])", "‰Â«Ì… «·› —… ﬁ»· »œ«Ì Â«° √Ê «·„”œœ √ﬂ»— „‰ «·÷—Ì»… «·„” Õﬁ…"
End Sub

Private Sub CreateTable_ManualEntries()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ManualEntries") Then Exit Sub
    AddField tdf, "ManualEntryID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "EntryNumber", "TEXT", 20, True, "", _
             "", "", "—ﬁ„ «·ﬁÌœ «·ÌœÊÌ", ""
    AddField tdf, "EntryDate", "DATE", 0, True, "Date()", _
             "", "", " «—ÌŒ «·ﬁÌœ", ""
    AddField tdf, "Description", "TEXT", 255, True, "", _
             "", "", "«·»Ì«‰", ""
    AddField tdf, "Reference", "TEXT", 50, False, "", _
             "", "", "«·„—Ã⁄", "—ﬁ„ „” ‰œ Œ«—ÃÌ: ›« Ê—…° ⁄ﬁœ° ﬂ‘› »‰ﬂ..."
    AddField tdf, "ReversalOfID", "LONG", 0, False, "", _
             "", "", "⁄ﬂ” «·ﬁÌœ", "«·ﬁÌœ «·ÌœÊÌ «·–Ì Ì⁄ﬂ”Â Â–« «·ﬁÌœ"
    AddField tdf, "TotalAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·ﬁÌœ", ""
    AddField tdf, "EmployeeID", "LONG", 0, True, "", _
             "", "", "√œŒ·Â", ""
    AddField tdf, "CreatedAt", "DATETIME", 0, True, "Now()", _
             "", "", " «—ÌŒ «·≈‰‘«¡", ""
    AddField tdf, "UpdatedAt", "DATETIME", 0, False, "", _
             "", "", "¬Œ—  ⁄œÌ·", ""
    AddField tdf, "CurrencyCode", "TEXT", 3, False, """SAR""", _
             "", "", "«·⁄„·…", "›«—€ = ⁄„·… «·»—‰«„Ã (Settings.CurrencyCode)"
    AddField tdf, "ExchangeRate", "RATE", 0, True, "1", _
             ">0", "«·„⁄«„· ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—", "„⁄«„· «· ÕÊÌ·", "ﬁÌ„… ÊÕœ… Ê«Õœ… „‰ «·⁄„·… »⁄„·… «·»—‰«„Ã"
    AddField tdf, "ForeignAmount", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "≈Ã„«·Ì «·ﬁÌœ »«·⁄„·…", "«·≈Ã„«·Ì »⁄„·… «·„” ‰œ (0 ··„” ‰œ«  «·ﬁœÌ„…)"
    AddIndex tdf, "PrimaryKey", "ManualEntryID", True, True, False
    AddIndex tdf, "UX_EntryNumber", "EntryNumber", False, True, False
    AddIndex tdf, "IX_EntryDate", "EntryDate", False, False, False
    EndTable tdf, "«·ﬁÌÊœ «·ÌœÊÌ…: ﬁÌœ Ìﬂ »Â «·„Õ«”» »‰›”Â („” Õﬁ« °  ”ÊÌ« ° —√” «·„«·° √—’œ… «›  «ÕÌ…...). Ìı—ÕÛ¯· ··ÌÊ„Ì… ﬂ√Ì ⁄„·Ì… √Œ—Ï (SourceType = MANUAL)° ÊÌı⁄œÛ¯· √Ê ÌıÕ–› ›Ì »⁄Â ﬁÌœÂ.", "", ""
End Sub

Private Sub CreateTable_ManualEntryLines()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "ManualEntryLines") Then Exit Sub
    AddField tdf, "ManualLineID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ «·”ÿ— «·œ«Œ·Ì", ""
    AddField tdf, "ManualEntryID", "LONG", 0, True, "", _
             "", "", "«·ﬁÌœ «·ÌœÊÌ", ""
    AddField tdf, "LineNumber", "INT", 0, True, "", _
             "", "", "—ﬁ„ «·”ÿ—", ""
    AddField tdf, "AccountCode", "LONG", 0, True, "", _
             "", "", "«·Õ”«»", ""
    AddField tdf, "Debit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„œÌ‰", ""
    AddField tdf, "Credit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "œ«∆‰", ""
    AddField tdf, "LineText", "TEXT", 150, False, "", _
             "", "", "»Ì«‰ «·”ÿ—", ""
    AddField tdf, "CostCenterID", "LONG", 0, False, "", _
             "", "", "„—ﬂ“ «· ﬂ·›…", ""
    AddField tdf, "ForeignDebit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "„œÌ‰ »«·⁄„·…", ""
    AddField tdf, "ForeignCredit", "MONEY", 0, True, "0", _
             ">=0", "«·„»·€ ·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ”«·»«", "œ«∆‰ »«·⁄„·…", ""
    AddIndex tdf, "PrimaryKey", "ManualLineID", True, True, False
    AddIndex tdf, "UX_ManualEntryID_LineNumber", "ManualEntryID,LineNumber", False, True, False
    AddIndex tdf, "IX_AccountCode", "AccountCode", False, False, False
    EndTable tdf, "√”ÿ— «·ﬁÌÊœ «·ÌœÊÌ…: «·ÿ—› «·„œÌ‰ Ê«·ÿ—› «·œ«∆‰ ··ﬁÌœ «·ÌœÊÌ∫ ﬂ· ”ÿ— „œÌ‰ √Ê œ«∆‰ ›ﬁÿ.", "([Debit]=0 Or [Credit]=0) And [Debit]+[Credit]>0", "ﬂ· ”ÿ— „œÌ‰ √Ê œ«∆‰ ›ﬁÿ° Ê»„»·€ √ﬂ»— „‰ ’›—"
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
    AddField tdf, "TypeNameEn", "TEXT", 50, False, "", _
             "", "", "«·«”„ »«·≈‰Ã·Ì“Ì…", ""
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
    AddField tdf, "RecordLabel", "TEXT", 100, False, "", _
             "", "", "«·”Ã·", "«”„ «·”Ã· √Ê —ﬁ„Â «·„ﬁ—Ê¡ («·„‰ Ã° —ﬁ„ «·„’—Ê›...)"
    AddIndex tdf, "PrimaryKey", "LogID", True, True, False
    AddIndex tdf, "IX_LogDate", "LogDate", False, False, False
    AddIndex tdf, "IX_ActionType", "ActionType", False, False, False
    EndTable tdf, "”Ã· «·⁄„·Ì« : ”Ã· «· œﬁÌﬁ: «·œŒÊ· Ê«·Œ—ÊÃ° Ê≈÷«›… √Ì ”Ã· √Ê  ⁄œÌ·Â √Ê Õ–›Â (Ê«·ÕﬁÊ· ›Ì AuditChanges)° Ê⁄„·Ì«  «·„” ‰œ«  Ê«·⁄„·Ì«  «·Õ”«”… ( Ã«Ê“ «·„Œ“Ê‰°  ⁄œÌ· «·√”⁄«—° «·‰”Œ «·«Õ Ì«ÿÌ).", "", ""
End Sub

Private Sub CreateTable_AuditChanges()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "AuditChanges") Then Exit Sub
    AddField tdf, "ChangeID", "AUTO", 0, False, "", _
             "", "", "—ﬁ„ œ«Œ·Ì", ""
    AddField tdf, "LogID", "LONG", 0, True, "", _
             "", "", "«·⁄„·Ì…", ""
    AddField tdf, "LineNo", "INT", 0, True, "0", _
             "", "", "«· — Ì»", ""
    AddField tdf, "FieldName", "TEXT", 64, True, "", _
             "", "", "«·Õﬁ·", ""
    AddField tdf, "FieldCaption", "TEXT", 100, False, "", _
             "", "", "«”„ «·Õﬁ·", ""
    AddField tdf, "OldValue", "TEXT", 255, False, "", _
             "", "", "«·ﬁÌ„… ﬁ»·", ""
    AddField tdf, "NewValue", "TEXT", 255, False, "", _
             "", "", "«·ﬁÌ„… »⁄œ", ""
    AddIndex tdf, "PrimaryKey", "ChangeID", True, True, False
    AddIndex tdf, "IX_LogID", "LogID", False, False, False
    EndTable tdf, " ›«’Ì· ”Ã· «· œﬁÌﬁ: ÕﬁÊ· ﬂ· ⁄„·Ì… ›Ì ”Ã· «· œﬁÌﬁ: «·ﬁÌ„… ﬁ»· Ê»⁄œ («·≈÷«›…: »⁄œ ›ﬁÿ° «·Õ–›: ﬁ»· ›ﬁÿ).", "", ""
End Sub

Private Sub CreateTable_LabelSettings()
    Dim tdf As DAO.TableDef
    If Not BeginTable(tdf, "LabelSettings") Then Exit Sub
    AddField tdf, "LabelSettingID", "LONG", 0, True, "1", _
             "=1", "Ì”„Õ »”Ã· ≈⁄œ«œ«  Ê«Õœ ›ﬁÿ", "—ﬁ„ «·≈⁄œ«œ", ""
    AddField tdf, "PrinterName", "TEXT", 255, False, "", _
             "", "", "ÿ«»⁄… «·„·’ﬁ« ", "›«—€ = «·ÿ«»⁄… «·«› —«÷Ì…"
    AddField tdf, "LabelWidth", "QTY", 0, True, "38", _
             "Between 15 And 210", "⁄—÷ «·„·’ﬁ „‰ 15 ≈·Ï 210 „„", "⁄—÷ «·„·’ﬁ („„)", ""
    AddField tdf, "LabelHeight", "QTY", 0, True, "25", _
             "Between 10 And 297", "«— ›«⁄ «·„·’ﬁ „‰ 10 ≈·Ï 297 „„", "«— ›«⁄ «·„·’ﬁ („„)", ""
    AddField tdf, "LabelsAcross", "BYTE", 0, True, "1", _
             "Between 1 And 10", "„‰ 1 ≈·Ï 10 „·’ﬁ«  ›Ì «·’›", "⁄œœ «·„·’ﬁ«  ›Ì «·’›", ""
    AddField tdf, "ColumnGap", "QTY", 0, True, "2", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·„”«›… »Ì‰ «·√⁄„œ… („„)", ""
    AddField tdf, "RowGap", "QTY", 0, True, "0", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·„”«›… »Ì‰ «·’›Ê› („„)", ""
    AddField tdf, "MarginTop", "QTY", 0, True, "0", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·Â«„‘ «·⁄·ÊÌ („„)", ""
    AddField tdf, "MarginBottom", "QTY", 0, True, "0", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·Â«„‘ «·”›·Ì („„)", ""
    AddField tdf, "MarginLeft", "QTY", 0, True, "0", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·Â«„‘ «·√Ì”— („„)", ""
    AddField tdf, "MarginRight", "QTY", 0, True, "0", _
             "Between 0 And 50", "„‰ 0 ≈·Ï 50 „„", "«·Â«„‘ «·√Ì„‰ („„)", ""
    AddField tdf, "BarHeight", "QTY", 0, True, "10", _
             "Between 3 And 100", "«— ›«⁄ «·»«—ﬂÊœ „‰ 3 ≈·Ï 100 „„", "«— ›«⁄ «·»«—ﬂÊœ („„)", ""
    AddField tdf, "BarWidth", "QTY", 0, True, "0.25", _
             "Between 0.1 And 1", "⁄—÷ √—›⁄ Œÿ „‰ 0.1 ≈·Ï 1 „„ («·„⁄ «œ 0.25 - 0.33)", "⁄—÷ √—›⁄ Œÿ („„)", ""
    AddField tdf, "TopLine1", "TEXT", 10, True, """STORE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "«Œ — „‰ «·ﬁ«∆„…", "«·”ÿ— «·√Ê· √⁄·Ï «·»«—ﬂÊœ", ""
    AddField tdf, "TopLine2", "TEXT", 10, True, """NAME""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "«Œ — „‰ «·ﬁ«∆„…", "«·”ÿ— «·À«‰Ì √⁄·Ï «·»«—ﬂÊœ", ""
    AddField tdf, "BottomLine1", "TEXT", 10, True, """BARCODE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "«Œ — „‰ «·ﬁ«∆„…", "«·”ÿ— «·√Ê· √”›· «·»«—ﬂÊœ", ""
    AddField tdf, "BottomLine2", "TEXT", 10, True, """PRICE""", _
             "In (""NONE"",""STORE"",""NAME"",""PRICE"",""CODE"",""BARCODE"")", "«Œ — „‰ «·ﬁ«∆„…", "«·”ÿ— «·À«‰Ì √”›· «·»«—ﬂÊœ", ""
    AddField tdf, "ShortName", "TEXT", 30, False, "", _
             "", "", "«·«”„ «·„Œ ’— ··„Õ·", "Ìıÿ»⁄ ≈–« «Œ —  ´«·«”„ «·„Œ ’—ª"
    AddField tdf, "FontSize", "BYTE", 0, True, "7", _
             "Between 5 And 16", "ÕÃ„ «·Œÿ „‰ 5 ≈·Ï 16", "ÕÃ„ «·Œÿ", ""
    AddIndex tdf, "PrimaryKey", "LabelSettingID", True, True, False
    EndTable tdf, "≈⁄œ«œ«  „·’ﬁ«  «·»«—ﬂÊœ: ”Ã· Ê«Õœ: „ﬁ«” «·„·’ﬁ Ê«·Ê—ﬁ Ê«·ÂÊ«„‘° ÊÕÃ„ «·»«—ﬂÊœ° Ê«·‰’Ê’ √⁄·«Â Ê√”›·Â.", "", ""
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
    Seed_Currencies
    Seed_CurrencyRates
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
    ExecSeed "INSERT INTO [Settings] ([SettingID], [StoreName], [CountryCode], [VATRate], [PricesIncludeVAT], [AllowNegativeStock], [CurrencyCode], [DefaultCustomerID], [ZatcaPhase], [LastInvoiceHash], [BackupKeepCount], [SlowMovingDays], [ReceiptFooter]) VALUES (1, '«”„ «·„Õ·', 'SA', 0.15, True, False, 'SAR', 1, 1, 'NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ==', 30, 90, '‘ﬂ—« ·“Ì«— ﬂ„')"
    EndSeed "Settings", 1
End Sub

Private Sub Seed_Sequences()
    If Not BeginSeed("Sequences", True) Then Exit Sub
    SeedRow "[SequenceName] = 'SALES_INVOICE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_INVOICE', 'INV-', 1, 6, '›Ê« Ì— «·„»Ì⁄« ')"
    SeedRow "[SequenceName] = 'SALES_RETURN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_RETURN', 'CRN-', 1, 6, '„— Ã⁄«  «·„»Ì⁄«  (≈‘⁄«— œ«∆‰)')"
    SeedRow "[SequenceName] = 'PURCHASE_INVOICE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_INVOICE', 'PUR-', 1, 6, '›Ê« Ì— «·„‘ —Ì« ')"
    SeedRow "[SequenceName] = 'PURCHASE_RETURN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PURCHASE_RETURN', 'PRT-', 1, 6, '„— Ã⁄«  «·„‘ —Ì« ')"
    SeedRow "[SequenceName] = 'CUSTOMER_PAYMENT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CUSTOMER_PAYMENT', 'RCV-', 1, 6, '”‰œ«  «·ﬁ»÷ „‰ «·⁄„·«¡')"
    SeedRow "[SequenceName] = 'SUPPLIER_PAYMENT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SUPPLIER_PAYMENT', 'PAY-', 1, 6, '”‰œ«  «·’—› ··„Ê—œÌ‰')"
    SeedRow "[SequenceName] = 'EXPENSE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('EXPENSE', 'EXP-', 1, 6, '«·„’—Ê›« ')"
    SeedRow "[SequenceName] = 'SALES_REP'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('SALES_REP', 'REP-', 1, 3, '√ﬂÊ«œ «·„‰œÊ»Ì‰')"
    SeedRow "[SequenceName] = 'COMMISSION_RUN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('COMMISSION_RUN', 'COM-', 1, 5, '„”Ì—«  «·⁄„Ê·« ')"
    SeedRow "[SequenceName] = 'STOCK_COUNT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_COUNT', 'CNT-', 1, 5, 'Ã·”«  «·Ã—œ')"
    SeedRow "[SequenceName] = 'STOCK_ADJUST'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('STOCK_ADJUST', 'ADJ-', 1, 6, 'Õ—ﬂ«  «·„Œ“Ê‰ «·ÌœÊÌ…')"
    SeedRow "[SequenceName] = 'PRODUCT_CODE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PRODUCT_CODE', 'P', 1, 5, '√ﬂÊ«œ «·„‰ Ã« ')"
    SeedRow "[SequenceName] = 'ZATCA_ICV'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('ZATCA_ICV', Null, 1, 0, '⁄œ¯«œ ICV ·„” ‰œ«  ›« Ê—…')"
    SeedRow "[SequenceName] = 'CASH_IN'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_IN', 'CIN-', 1, 6, '”‰œ«  ﬁ»÷ «·‰ﬁœÌ… («·Œ“Ì‰…)')"
    SeedRow "[SequenceName] = 'CASH_OUT'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_OUT', 'COT-', 1, 6, '”‰œ«  ’—› «·‰ﬁœÌ… («·Œ“Ì‰…)')"
    SeedRow "[SequenceName] = 'CASH_TRANSFER'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_TRANSFER', 'TRF-', 1, 6, '«· ÕÊÌ· »Ì‰ «·’‰«œÌﬁ')"
    SeedRow "[SequenceName] = 'CASH_CLOSING'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CASH_CLOSING', 'CLS-', 1, 6, ' ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—')"
    SeedRow "[SequenceName] = 'JOURNAL'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('JOURNAL', 'JV-', 1, 6, 'ﬁÌÊœ «·ÌÊ„Ì…')"
    SeedRow "[SequenceName] = 'MANUAL_ENTRY'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('MANUAL_ENTRY', 'MJ-', 1, 6, '«·ﬁÌÊœ «·ÌœÊÌ…')"
    SeedRow "[SequenceName] = 'BANK_TX'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('BANK_TX', 'BNK-', 1, 6, '«·Õ—ﬂ«  «·»‰ﬂÌ…')"
    SeedRow "[SequenceName] = 'BANK_RECON'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('BANK_RECON', 'REC-', 1, 6, '«· ”ÊÌ«  «·»‰ﬂÌ…')"
    SeedRow "[SequenceName] = 'CHEQUE'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('CHEQUE', 'CHQ-', 1, 6, '«·‘Ìﬂ«  «·Ê«—œ… Ê«·’«œ—…')"
    SeedRow "[SequenceName] = 'FIXED_ASSET'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('FIXED_ASSET', 'FA-', 1, 5, '«·√’Ê· «·À«» …')"
    SeedRow "[SequenceName] = 'PAYROLL'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('PAYROLL', 'PAY-', 1, 5, '„”Ì—«  «·—Ê« »')"
    SeedRow "[SequenceName] = 'DEPRECIATION'", "INSERT INTO [Sequences] ([SequenceName], [Prefix], [NextValue], [PadLength], [Description]) VALUES ('DEPRECIATION', 'DEP-', 1, 6, 'ﬁÌÊœ «·≈Â·«ﬂ «·‘Â—Ì…')"
    EndSeed "Sequences", 25
End Sub

Private Sub Seed_Roles()
    If Not BeginSeed("Roles", False) Then Exit Sub
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (1, 'ADMIN', '„œÌ— «·‰Ÿ«„', 'Ã„Ì⁄ «·’·«ÕÌ« ')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (2, 'MANAGER', '„œÌ—', '«·„»Ì⁄«  Ê«·„‘ —Ì«  Ê«·„Œ“Ê‰ Ê«· ﬁ«—Ì—')"
    ExecSeed "INSERT INTO [Roles] ([RoleID], [RoleCode], [RoleName], [Description]) VALUES (3, 'CASHIER', 'ﬂ«‘Ì—', '«·„»Ì⁄«  Ê«·⁄„·«¡ ›ﬁÿ')"
    EndSeed "Roles", 3
End Sub

Private Sub Seed_Permissions()
    If Not BeginSeed("Permissions", True) Then Exit Sub
    If SeedRow("[PermissionKey] = 'SALES_POS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_POS', '‰ﬁÿ… «·»Ì⁄', '«·„»Ì⁄« ', 10)") Then GrantNewPermission "SALES_POS", "1,2,3"
    If SeedRow("[PermissionKey] = 'SALES_VIEW'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_VIEW', '⁄—÷ Ê≈⁄«œ… ÿ»«⁄… «·›Ê« Ì—', '«·„»Ì⁄« ', 11)") Then GrantNewPermission "SALES_VIEW", "1,2,3"
    If SeedRow("[PermissionKey] = 'SALES_RETURN'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_RETURN', '„— Ã⁄«  «·„»Ì⁄« ', '«·„»Ì⁄« ', 12)") Then GrantNewPermission "SALES_RETURN", "1,2"
    If SeedRow("[PermissionKey] = 'PRICE_OVERRIDE'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRICE_OVERRIDE', ' ⁄œÌ· ”⁄— «·»Ì⁄ ›Ì «·›« Ê—…', '«·„»Ì⁄« ', 13)") Then GrantNewPermission "PRICE_OVERRIDE", "1,2"
    If SeedRow("[PermissionKey] = 'DISCOUNT_OVERRIDE'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DISCOUNT_OVERRIDE', 'Œ’„ √⁄·Ï „‰ «·Õœ «·„”„ÊÕ', '«·„»Ì⁄« ', 14)") Then GrantNewPermission "DISCOUNT_OVERRIDE", "1,2"
    If SeedRow("[PermissionKey] = 'ALLOW_NEGATIVE_STOCK'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('ALLOW_NEGATIVE_STOCK', '«·»Ì⁄ »ﬂ„Ì… √ﬂ»— „‰ «·„ Ê›—', '«·„»Ì⁄« ', 15)") Then GrantNewPermission "ALLOW_NEGATIVE_STOCK", "1"
    If SeedRow("[PermissionKey] = 'CUSTOMERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMERS', '≈œ«—… «·⁄„·«¡', '«·⁄„·«¡', 20)") Then GrantNewPermission "CUSTOMERS", "1,2,3"
    If SeedRow("[PermissionKey] = 'CUSTOMER_PAYMENTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CUSTOMER_PAYMENTS', '”‰œ«  «·ﬁ»÷', '«·⁄„·«¡', 21)") Then GrantNewPermission "CUSTOMER_PAYMENTS", "1,2,3"
    If SeedRow("[PermissionKey] = 'PURCHASES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASES', '›Ê« Ì— «·„‘ —Ì« ', '«·„‘ —Ì« ', 30)") Then GrantNewPermission "PURCHASES", "1,2"
    If SeedRow("[PermissionKey] = 'PURCHASE_RETURN'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PURCHASE_RETURN', '„— Ã⁄«  «·„‘ —Ì« ', '«·„‘ —Ì« ', 31)") Then GrantNewPermission "PURCHASE_RETURN", "1,2"
    If SeedRow("[PermissionKey] = 'SUPPLIERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIERS', '≈œ«—… «·„Ê—œÌ‰', '«·„Ê—œÊ‰', 40)") Then GrantNewPermission "SUPPLIERS", "1,2"
    If SeedRow("[PermissionKey] = 'SUPPLIER_PAYMENTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SUPPLIER_PAYMENTS', '”‰œ«  «·’—›', '«·„Ê—œÊ‰', 41)") Then GrantNewPermission "SUPPLIER_PAYMENTS", "1,2"
    If SeedRow("[PermissionKey] = 'PRODUCTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PRODUCTS', '≈œ«—… «·„‰ Ã«  Ê«·√”⁄«—', '«·„Œ“Ê‰', 50)") Then GrantNewPermission "PRODUCTS", "1,2"
    If SeedRow("[PermissionKey] = 'INVENTORY_ADJUST'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('INVENTORY_ADJUST', '≈÷«›… ÊŒ’„ „Œ“Ê‰ ÌœÊÌ', '«·„Œ“Ê‰', 51)") Then GrantNewPermission "INVENTORY_ADJUST", "1,2"
    If SeedRow("[PermissionKey] = 'STOCK_COUNT'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('STOCK_COUNT', '«·Ã—œ', '«·„Œ“Ê‰', 52)") Then GrantNewPermission "STOCK_COUNT", "1,2"
    If SeedRow("[PermissionKey] = 'EXPENSES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('EXPENSES', '«·„’—Ê›« ', '«·„’—Ê›« ', 60)") Then GrantNewPermission "EXPENSES", "1,2"
    If SeedRow("[PermissionKey] = 'CASH_BOX'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CASH_BOX', '«·Œ“Ì‰…: ”‰œ«  «·ﬁ»÷ Ê«·’—› Ê«· ÕÊÌ· Ê«·’‰«œÌﬁ', '«·Œ“Ì‰…', 65)") Then GrantNewPermission "CASH_BOX", "1,2"
    If SeedRow("[PermissionKey] = 'CASH_CLOSING'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CASH_CLOSING', ' ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—', '«·Œ“Ì‰…', 66)") Then GrantNewPermission "CASH_CLOSING", "1,2,3"
    If SeedRow("[PermissionKey] = 'JOURNAL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('JOURNAL', 'ﬁÌÊœ «·ÌÊ„Ì… Êœ·Ì· «·Õ”«»«  Ê„Ì“«‰ «·„—«Ã⁄…', '«·Õ”«»« ', 75)") Then GrantNewPermission "JOURNAL", "1,2"
    If SeedRow("[PermissionKey] = 'MANUAL_ENTRY'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('MANUAL_ENTRY', '«·ﬁÌÊœ «·ÌœÊÌ…: ≈÷«›… Ê ⁄œÌ· ÊÕ–›', '«·Õ”«»« ', 76)") Then GrantNewPermission "MANUAL_ENTRY", "1,2"
    If SeedRow("[PermissionKey] = 'PERIOD_CLOSE'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PERIOD_CLOSE', '≈ﬁ›«· «·› —«  Ê«·”‰… «·„«·Ì… Ê≈⁄«œ… › ÕÂ«', '«·Õ”«»« ', 77)") Then GrantNewPermission "PERIOD_CLOSE", "1"
    If SeedRow("[PermissionKey] = 'VAT_RETURN'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('VAT_RETURN', '≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…: «·«⁄ „«œ Ê«·”œ«œ', '«·Õ”«»« ', 78)") Then GrantNewPermission "VAT_RETURN", "1,2"
    If SeedRow("[PermissionKey] = 'BANKS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BANKS', '«·»‰Êﬂ: «·Õ”«»«  Ê«·Õ—ﬂ«  «·»‰ﬂÌ… Ê«· ”ÊÌ… «·»‰ﬂÌ…', '«·Œ“Ì‰…', 67)") Then GrantNewPermission "BANKS", "1,2"
    If SeedRow("[PermissionKey] = 'CHEQUES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CHEQUES', '«·‘Ìﬂ«  «·Ê«—œ… Ê«·’«œ—…: «· ”ÃÌ· Ê«· Õ’Ì· Ê«·«— œ«œ', '«·Œ“Ì‰…', 68)") Then GrantNewPermission "CHEQUES", "1,2"
    If SeedRow("[PermissionKey] = 'FIXED_ASSETS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('FIXED_ASSETS', '«·√’Ê· «·À«» … Ê«·≈Â·«ﬂ', '«·Õ”«»« ', 79)") Then GrantNewPermission "FIXED_ASSETS", "1,2"
    If SeedRow("[PermissionKey] = 'PAYROLL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('PAYROLL', '„”Ì— «·—Ê« »: «·≈⁄œ«œ Ê«· —ÕÌ· Ê«·’—›', '«·Õ”«»« ', 80)") Then GrantNewPermission "PAYROLL", "1,2"
    If SeedRow("[PermissionKey] = 'BUDGET'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BUDGET', '«·„Ê«“‰… «· ﬁœÌ—Ì…: «·≈⁄œ«œ Ê«·„ﬁ«—‰… »«·›⁄·Ì', '«·Õ”«»« ', 81)") Then GrantNewPermission "BUDGET", "1,2"
    If SeedRow("[PermissionKey] = 'CURRENCIES'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('CURRENCIES', '«·⁄„·«  Ê√”⁄«—Â«', '«·Õ”«»« ', 84)") Then GrantNewPermission "CURRENCIES", "1,2"
    If SeedRow("[PermissionKey] = 'SALES_REPS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SALES_REPS', '«·„‰œÊ»Ì‰: «·»Ì«‰«  Ê«·√Âœ«› Ê«·⁄„Ê·«  Ê ﬁ«—Ì—Â„', '«·„»Ì⁄« ', 15)") Then GrantNewPermission "SALES_REPS", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS', '«· ﬁ«—Ì— «· ‘€Ì·Ì…', '«· ﬁ«—Ì—', 70)") Then GrantNewPermission "REPORTS", "1,2"
    If SeedRow("[PermissionKey] = 'REPORTS_PROFIT'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('REPORTS_PROFIT', ' ﬁ«—Ì— «·√—»«Õ Ê«·÷—Ì»…', '«· ﬁ«—Ì—', 71)") Then GrantNewPermission "REPORTS_PROFIT", "1,2"
    If SeedRow("[PermissionKey] = 'DASHBOARD_FINANCIAL'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('DASHBOARD_FINANCIAL', '«·√—ﬁ«„ «·„«·Ì… ›Ì ·ÊÕ… «· Õﬂ„', '«· ﬁ«—Ì—', 72)") Then GrantNewPermission "DASHBOARD_FINANCIAL", "1,2"
    If SeedRow("[PermissionKey] = 'SETTINGS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('SETTINGS', '≈⁄œ«œ«  «·„Õ·', '«·‰Ÿ«„', 80)") Then GrantNewPermission "SETTINGS", "1"
    If SeedRow("[PermissionKey] = 'USERS'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('USERS', '«·„” Œœ„Ê‰ Ê«·’·«ÕÌ« ', '«·‰Ÿ«„', 81)") Then GrantNewPermission "USERS", "1"
    If SeedRow("[PermissionKey] = 'BACKUP'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('BACKUP', '«·‰”Œ «·«Õ Ì«ÿÌ', '«·‰Ÿ«„', 82)") Then GrantNewPermission "BACKUP", "1"
    If SeedRow("[PermissionKey] = 'AUDIT_LOG'", "INSERT INTO [Permissions] ([PermissionKey], [PermissionName], [ModuleName], [SortOrder]) VALUES ('AUDIT_LOG', '”Ã· «· œﬁÌﬁ: „‰ √÷«› √Ê ⁄œ¯· √Ê Õ–›° Ê«·ﬁÌ„ ﬁ»· Ê»⁄œ', '«·‰Ÿ«„', 83)") Then GrantNewPermission "AUDIT_LOG", "1"
    EndSeed "Permissions", 36
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
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PERIOD_CLOSE')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'VAT_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'BANKS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CHEQUES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'FIXED_ASSETS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'PAYROLL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'BUDGET')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'CURRENCIES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SALES_REPS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'SETTINGS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'USERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'BACKUP')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (1, 'AUDIT_LOG')"
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
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'VAT_RETURN')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'BANKS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CHEQUES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'FIXED_ASSETS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'PAYROLL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'BUDGET')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'CURRENCIES')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'SALES_REPS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'REPORTS_PROFIT')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (2, 'DASHBOARD_FINANCIAL')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_POS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'SALES_VIEW')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMERS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CUSTOMER_PAYMENTS')"
    ExecSeed "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'CASH_CLOSING')"
    EndSeed "RolePermissions", 71
End Sub

Private Sub Seed_Employees()
    If Not BeginSeed("Employees", False) Then Exit Sub
    ExecSeed "INSERT INTO [Employees] ([EmployeeID], [EmployeeName], [JobTitle], [Username], [RoleID], [MaxDiscountPercent], [MustChangePassword], [IsActive]) VALUES (1, '„œÌ— «·‰Ÿ«„', '„œÌ— «·‰Ÿ«„', 'admin', 1, 1, True, True)"
    EndSeed "Employees", 1
End Sub

Private Sub Seed_Screens()
    If Not BeginSeed("Screens", True) Then Exit Sub
    SeedRow "[ScreenName] = 'frmPOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPOS', '‰ﬁÿ… «·»Ì⁄ («·„Õ·« )', '«·„»Ì⁄« ', 10, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmTouchPOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmTouchPOS', '‰ﬁÿ… «·»Ì⁄ («·„ÿ«⁄„)', '«·„»Ì⁄« ', 20, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmCafePOS'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCafePOS', '‰ﬁÿ… «·»Ì⁄ («·ﬂ«›ÌÂ« )', '«·„»Ì⁄« ', 30, 'SALES_POS', True, False, False)"
    SeedRow "[ScreenName] = 'frmSalesInvoice'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSalesInvoice', '⁄—÷ «·›Ê« Ì— Ê≈⁄«œ… ÿ»«⁄ Â«', '«·„»Ì⁄« ', 40, 'SALES_VIEW', False, False, False)"
    SeedRow "[ScreenName] = 'frmSalesReturn'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSalesReturn', '„— Ã⁄«  «·„»Ì⁄« ', '«·„»Ì⁄« ', 50, 'SALES_RETURN', True, False, False)"
    SeedRow "[ScreenName] = 'frmCustomers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCustomers', '«·⁄„·«¡', '«·⁄„·«¡', 60, 'CUSTOMERS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCustomerPayment'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCustomerPayment', '”‰œ«  «·ﬁ»÷ „‰ «·⁄„·«¡', '«·⁄„·«¡', 70, 'CUSTOMER_PAYMENTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseInvoice'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseInvoice', '›Ê« Ì— «·„‘ —Ì« ', '«·„‘ —Ì« ', 80, 'PURCHASES', True, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseView'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseView', '⁄—÷ ›Ê« Ì— «·„‘ —Ì« ', '«·„‘ —Ì« ', 90, 'PURCHASES', False, False, False)"
    SeedRow "[ScreenName] = 'frmPurchaseReturn'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPurchaseReturn', '„— Ã⁄«  «·„‘ —Ì« ', '«·„‘ —Ì« ', 100, 'PURCHASE_RETURN', True, False, False)"
    SeedRow "[ScreenName] = 'frmSuppliers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSuppliers', '«·„Ê—œÊ‰', '«·„Ê—œÊ‰', 110, 'SUPPLIERS', True, True, True)"
    SeedRow "[ScreenName] = 'frmSupplierPayment'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSupplierPayment', '”‰œ«  «·’—› ··„Ê—œÌ‰', '«·„Ê—œÊ‰', 120, 'SUPPLIER_PAYMENTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmProducts'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmProducts', '«·„‰ Ã«  Ê«·√”⁄«—', '«·„Œ“Ê‰', 130, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCategories'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCategories', '«· ’‰Ì›« ', '«·„Œ“Ê‰', 140, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmUnits'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUnits', '«·ÊÕœ« ', '«·„Œ“Ê‰', 150, 'PRODUCTS', True, True, True)"
    SeedRow "[ScreenName] = 'frmInventory'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmInventory', '«·„Œ“Ê‰ Ê«·Õ—ﬂ«  «·ÌœÊÌ…', '«·„Œ“Ê‰', 160, 'PRODUCTS', True, False, False)"
    SeedRow "[ScreenName] = 'frmStockCount'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmStockCount', '«·Ã—œ', '«·„Œ“Ê‰', 170, 'STOCK_COUNT', True, False, False)"
    SeedRow "[ScreenName] = 'frmBarcodeLabels'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBarcodeLabels', '„·’ﬁ«  «·»«—ﬂÊœ', '«·„Œ“Ê‰', 180, 'PRODUCTS', False, False, False)"
    SeedRow "[ScreenName] = 'frmLabelSettings'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmLabelSettings', '≈⁄œ«œ«  «·„·’ﬁ« ', '«·„Œ“Ê‰', 190, 'PRODUCTS', False, True, False)"
    SeedRow "[ScreenName] = 'frmExpenses'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmExpenses', '«·„’—Ê›« ', '«·„’—Ê›« ', 200, 'EXPENSES', True, True, True)"
    SeedRow "[ScreenName] = 'frmExpenseTypes'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmExpenseTypes', '√‰Ê«⁄ «·„’—Ê›« ', '«·„’—Ê›« ', 210, 'EXPENSES', True, True, True)"
    SeedRow "[ScreenName] = 'frmRecurring'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmRecurring', '«·„’—Ê›«  «·„ ﬂ——…', '«·„’—Ê›« ', 220, 'EXPENSES', True, True, True)"
    SeedRow "[ScreenName] = 'frmTreasury'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmTreasury', '«·Œ“Ì‰…', '«·Œ“Ì‰…', 230, 'CASH_CLOSING', False, False, False)"
    SeedRow "[ScreenName] = 'frmCashVoucher'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashVoucher', '”‰œ«  «·‰ﬁœÌ… Ê«· ÕÊÌ·', '«·Œ“Ì‰…', 240, 'CASH_BOX', True, False, False)"
    SeedRow "[ScreenName] = 'frmCashClosing'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashClosing', ' ’›Ì… ÌÊ„Ì… «·ﬂ«‘Ì—', '«·Œ“Ì‰…', 250, 'CASH_CLOSING', True, False, False)"
    SeedRow "[ScreenName] = 'frmCashBoxes'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCashBoxes', '«·’‰«œÌﬁ', '«·Œ“Ì‰…', 260, 'CASH_BOX', True, True, True)"
    SeedRow "[ScreenName] = 'frmAccounting'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAccounting', '«·„Õ«”»… Ê«·„«·Ì…', '«·Õ”«»« ', 270, Null, False, False, False)"
    SeedRow "[ScreenName] = 'frmJournal'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmJournal', 'ﬁÌÊœ «·ÌÊ„Ì…', '«·Õ”«»« ', 280, 'JOURNAL', False, False, False)"
    SeedRow "[ScreenName] = 'frmAccounts'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAccounts', 'œ·Ì· «·Õ”«»« ', '«·Õ”«»« ', 290, 'JOURNAL', True, True, True)"
    SeedRow "[ScreenName] = 'frmManualEntry'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmManualEntry', '«·ﬁÌÊœ «·ÌœÊÌ…', '«·Õ”«»« ', 300, 'MANUAL_ENTRY', True, True, True)"
    SeedRow "[ScreenName] = 'frmLedger'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmLedger', 'ﬂ‘› Õ”«» Êœ› — «·√” «–', '«·Õ”«»« ', 310, 'JOURNAL', False, False, False)"
    SeedRow "[ScreenName] = 'frmFinancials'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmFinancials', '«·ﬁÊ«∆„ «·„«·Ì…', '«·Õ”«»« ', 320, 'REPORTS_PROFIT', False, False, False)"
    SeedRow "[ScreenName] = 'frmPeriodClosing'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPeriodClosing', '≈ﬁ›«· «·› —«  Ê«·”‰… «·„«·Ì…', '«·Õ”«»« ', 330, 'PERIOD_CLOSE', False, False, False)"
    SeedRow "[ScreenName] = 'frmVatReturn'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmVatReturn', '≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…', '«·Õ”«»« ', 340, 'VAT_RETURN', True, True, True)"
    SeedRow "[ScreenName] = 'frmAging'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAging', '√⁄„«— «·œÌÊ‰ («·⁄„·«¡ Ê«·„Ê—œÊ‰)', '«· ﬁ«—Ì—', 350, 'REPORTS', False, False, False)"
    SeedRow "[ScreenName] = 'frmBanks'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBanks', '«·»‰Êﬂ', '«·Œ“Ì‰…', 360, 'BANKS', True, True, True)"
    SeedRow "[ScreenName] = 'frmBankTx'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBankTx', '«·Õ—ﬂ«  «·»‰ﬂÌ…', '«·Œ“Ì‰…', 370, 'BANKS', True, False, True)"
    SeedRow "[ScreenName] = 'frmBankRecon'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBankRecon', '«· ”ÊÌ… «·»‰ﬂÌ…', '«·Œ“Ì‰…', 380, 'BANKS', True, True, False)"
    SeedRow "[ScreenName] = 'frmCheques'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCheques', '«·‘Ìﬂ«  «·Ê«—œ… Ê«·’«œ—…', '«·Œ“Ì‰…', 390, 'CHEQUES', True, True, True)"
    SeedRow "[ScreenName] = 'frmAssets'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAssets', '«·√’Ê· «·À«» …', '«·Õ”«»« ', 400, 'FIXED_ASSETS', True, True, True)"
    SeedRow "[ScreenName] = 'frmDepreciation'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmDepreciation', '«·≈Â·«ﬂ «·‘Â—Ì', '«·Õ”«»« ', 410, 'FIXED_ASSETS', True, False, True)"
    SeedRow "[ScreenName] = 'frmPayroll'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmPayroll', '„”Ì— «·—Ê« »', '«·Õ”«»« ', 420, 'PAYROLL', True, True, True)"
    SeedRow "[ScreenName] = 'frmCostCenters'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCostCenters', '„—«ﬂ“ «· ﬂ·›… Ê«·›—Ê⁄', '«·Õ”«»« ', 430, 'JOURNAL', True, True, True)"
    SeedRow "[ScreenName] = 'frmBudget'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBudget', '«·„Ê«“‰… «· ﬁœÌ—Ì…', '«·Õ”«»« ', 440, 'BUDGET', True, True, True)"
    SeedRow "[ScreenName] = 'frmCurrencies'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCurrencies', '«·⁄„·« ', '«·Õ”«»« ', 450, 'CURRENCIES', True, True, True)"
    SeedRow "[ScreenName] = 'frmSalesReps'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSalesReps', '«·„‰œÊ»Ì‰', '«·„»Ì⁄« ', 460, 'SALES_REPS', True, True, True)"
    SeedRow "[ScreenName] = 'frmRepTargets'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmRepTargets', '√Âœ«› «·„‰œÊ»Ì‰', '«·„»Ì⁄« ', 470, 'SALES_REPS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCommissions'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCommissions', '⁄„Ê·«  «·„‰œÊ»Ì‰', '«·„»Ì⁄« ', 480, 'SALES_REPS', True, True, True)"
    SeedRow "[ScreenName] = 'frmCurrencyRates'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmCurrencyRates', '√”⁄«— «·⁄„·« ', '«·Õ”«»« ', 490, 'CURRENCIES', True, True, True)"
    SeedRow "[ScreenName] = 'frmAllocation'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAllocation', '—»ÿ «·”œ«œ »«·›Ê« Ì—', '«·⁄„·«¡', 500, 'CUSTOMER_PAYMENTS', True, False, True)"
    SeedRow "[ScreenName] = 'frmReportCenter'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmReportCenter', '«· ﬁ«—Ì—', '«· ﬁ«—Ì—', 510, 'REPORTS', False, False, False)"
    SeedRow "[ScreenName] = 'frmSearch'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSearch', '«·»ÕÀ', '«·‰Ÿ«„', 520, Null, False, False, False)"
    SeedRow "[ScreenName] = 'frmSettings'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmSettings', '≈⁄œ«œ«  «·„Õ·', '«·‰Ÿ«„', 530, 'SETTINGS', False, True, False)"
    SeedRow "[ScreenName] = 'frmUsers'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUsers', '«·„” Œœ„Ê‰', '«·‰Ÿ«„', 540, 'USERS', True, True, False)"
    SeedRow "[ScreenName] = 'frmRoles'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmRoles', '«·√œÊ«— Ê«·’·«ÕÌ« ', '«·‰Ÿ«„', 550, 'USERS', False, True, False)"
    SeedRow "[ScreenName] = 'frmUserScreens'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmUserScreens', '’·«ÕÌ«  «·‘«‘«  ··„” Œœ„Ì‰', '«·‰Ÿ«„', 560, 'USERS', False, True, False)"
    SeedRow "[ScreenName] = 'frmAuditLog'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmAuditLog', '”Ã· «· œﬁÌﬁ', '«·‰Ÿ«„', 570, 'AUDIT_LOG', False, False, False)"
    SeedRow "[ScreenName] = 'frmBackup'", "INSERT INTO [Screens] ([ScreenName], [ScreenTitle], [ModuleName], [SortOrder], [PermissionKey], [HasAdd], [HasEdit], [HasDelete]) VALUES ('frmBackup', '«·‰”Œ «·«Õ Ì«ÿÌ', '«·‰Ÿ«„', 580, 'BACKUP', False, False, False)"
    EndSeed "Screens", 58
End Sub

Private Sub Seed_Categories()
    If Not BeginSeed("Categories", False) Then Exit Sub
    ExecSeed "INSERT INTO [Categories] ([CategoryID], [CategoryName], [Description]) VALUES (1, '⁄«„', ' ’‰Ì› «› —«÷Ì')"
    EndSeed "Categories", 1
End Sub

Private Sub Seed_Units()
    If Not BeginSeed("Units", False) Then Exit Sub
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
    If Not BeginSeed("PaymentMethods", False) Then Exit Sub
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (1, '‰ﬁœÌ', '10', 1)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (2, '„œÏ / »ÿ«ﬁ…', '48', 2)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (3, ' ÕÊÌ· »‰ﬂÌ', '42', 3)"
    ExecSeed "INSERT INTO [PaymentMethods] ([PaymentMethodID], [MethodName], [ZatcaCode], [SortOrder]) VALUES (4, '„Õ›Ÿ… ≈·ﬂ —Ê‰Ì…', '1', 4)"
    EndSeed "PaymentMethods", 4
End Sub

Private Sub Seed_Currencies()
    If Not BeginSeed("Currencies", True) Then Exit Sub
    SeedRow "[CurrencyCode] = 'SAR'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('SAR', '—Ì«· ”⁄ÊœÌ', 'Saudi Riyal', '—.”', 2, 1)"
    SeedRow "[CurrencyCode] = 'USD'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('USD', 'œÊ·«— √„—ÌﬂÌ', 'US Dollar', '$', 2, 2)"
    SeedRow "[CurrencyCode] = 'EUR'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('EUR', 'ÌÊ—Ê', 'Euro', 'Ä', 2, 3)"
    SeedRow "[CurrencyCode] = 'AED'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('AED', 'œ—Â„ ≈„«—« Ì', 'UAE Dirham', 'œ.≈', 2, 4)"
    SeedRow "[CurrencyCode] = 'KWD'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('KWD', 'œÌ‰«— ﬂÊÌ Ì', 'Kuwaiti Dinar', 'œ.ﬂ', 3, 5)"
    SeedRow "[CurrencyCode] = 'BHD'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('BHD', 'œÌ‰«— »Õ—Ì‰Ì', 'Bahraini Dinar', 'œ.»', 3, 6)"
    SeedRow "[CurrencyCode] = 'QAR'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('QAR', '—Ì«· ﬁÿ—Ì', 'Qatari Riyal', '—.ﬁ', 2, 7)"
    SeedRow "[CurrencyCode] = 'OMR'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('OMR', '—Ì«· ⁄„«‰Ì', 'Omani Rial', '—.⁄', 3, 8)"
    SeedRow "[CurrencyCode] = 'EGP'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('EGP', 'Ã‰ÌÂ „’—Ì', 'Egyptian Pound', 'Ã.„', 2, 9)"
    SeedRow "[CurrencyCode] = 'GBP'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('GBP', 'Ã‰ÌÂ ≈” —·Ì‰Ì', 'Pound Sterling', '£', 2, 10)"
    SeedRow "[CurrencyCode] = 'CNY'", "INSERT INTO [Currencies] ([CurrencyCode], [CurrencyName], [CurrencyNameEn], [Symbol], [DecimalPlaces], [SortOrder]) VALUES ('CNY', 'ÌÊ«‰ ’Ì‰Ì', 'Chinese Yuan', '•', 2, 11)"
    EndSeed "Currencies", 11
End Sub

Private Sub Seed_CurrencyRates()
    If Not BeginSeed("CurrencyRates", True) Then Exit Sub
    SeedRow "[CurrencyRateID] = 1", "INSERT INTO [CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (1, 'USD', '2000-01-01', 3.75, '”⁄— «·—»ÿ «·—”„Ì')"
    SeedRow "[CurrencyRateID] = 2", "INSERT INTO [CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (2, 'AED', '2000-01-01', 1.0211, '„— »ÿ »«·œÊ·«—')"
    SeedRow "[CurrencyRateID] = 3", "INSERT INTO [CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (3, 'BHD', '2000-01-01', 9.9734, '„— »ÿ »«·œÊ·«—')"
    SeedRow "[CurrencyRateID] = 4", "INSERT INTO [CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (4, 'QAR', '2000-01-01', 1.0302, '„— »ÿ »«·œÊ·«—')"
    SeedRow "[CurrencyRateID] = 5", "INSERT INTO [CurrencyRates] ([CurrencyRateID], [CurrencyCode], [RateDate], [Rate], [Notes]) VALUES (5, 'OMR', '2000-01-01', 9.7529, '„— »ÿ »«·œÊ·«—')"
    EndSeed "CurrencyRates", 5
End Sub

Private Sub Seed_CashBoxes()
    If Not BeginSeed("CashBoxes", False) Then Exit Sub
    ExecSeed "INSERT INTO [CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (1, '«·Œ“Ì‰… «·—∆Ì”Ì…', 'MAIN')"
    ExecSeed "INSERT INTO [CashBoxes] ([CashBoxID], [BoxName], [BoxType]) VALUES (2, '’‰œÊﬁ «·ﬂ«‘Ì—', 'CASHIER')"
    EndSeed "CashBoxes", 2
End Sub

Private Sub Seed_Customers()
    If Not BeginSeed("Customers", False) Then Exit Sub
    ExecSeed "INSERT INTO [Customers] ([CustomerID], [CustomerName], [AllowCredit], [IsSystem]) VALUES (1, '⁄„Ì· ‰ﬁœÌ', False, True)"
    EndSeed "Customers", 1
End Sub

Private Sub Seed_ExpenseTypes()
    If Not BeginSeed("ExpenseTypes", False) Then Exit Sub
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

Private Sub Seed_Accounts()
    If Not BeginSeed("Accounts", True) Then Exit Sub
    SeedRow "[AccountCode] = 1", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1, '«·√’Ê·', 'ASSET', Null, False, True)"
    SeedRow "[AccountCode] = 11", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (11, '«·√’Ê· «·„ œ«Ê·…', 'ASSET', 1, False, True)"
    SeedRow "[AccountCode] = 1100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1100, '«·‰ﬁœÌ… »«·Œ“Ì‰… Ê«·’‰«œÌﬁ', 'ASSET', 11, False, True)"
    SeedRow "[AccountCode] = 110001", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110001, '«·Œ“Ì‰… «·—∆Ì”Ì…', 'ASSET', 1100, True, True)"
    SeedRow "[AccountCode] = 110002", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (110002, '’‰œÊﬁ «·ﬂ«‘Ì—', 'ASSET', 1100, True, True)"
    SeedRow "[AccountCode] = 1190", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1190, '‰ﬁœÌ… €Ì— „Ê“⁄… ⁄·Ï ’‰œÊﬁ', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1200, '„œÏ Ê«·„Õ«›Ÿ  Õ  «· ”ÊÌ… («·»‰ﬂ Ê«·‘»ﬂ…)', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1210", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1210, '«·Õ”«»«  «·»‰ﬂÌ…', 'ASSET', 11, False, True)"
    SeedRow "[AccountCode] = 1250", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1250, '‘Ìﬂ«   Õ  «· Õ’Ì·', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1300, '–„„ «·⁄„·«¡', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1310", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1310, '√Ê—«ﬁ «·ﬁ»÷', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1350", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1350, '„Œ’’ «·œÌÊ‰ «·„‘ﬂÊﬂ ›Ì  Õ’Ì·Â«', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1400, '«·„Œ“Ê‰', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1500, '÷—Ì»… «·ﬁÌ„… «·„÷«›… - „œŒ·« ', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1600, '”·› «·„ÊŸ›Ì‰', 'ASSET', 11, True, True)"
    SeedRow "[AccountCode] = 1610", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1610, '«·⁄Âœ «·‰ﬁœÌ…', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1650", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1650, '„’—Ê›«  „œ›Ê⁄… „ﬁœ„«', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1660", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1660, '≈Ì—«œ«  „” Õﬁ…', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 1690", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1690, ' √„Ì‰«  ·œÏ «·€Ì—', 'ASSET', 11, True, False)"
    SeedRow "[AccountCode] = 12", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (12, '«·√’Ê· €Ì— «·„ œ«Ê·…', 'ASSET', 1, False, True)"
    SeedRow "[AccountCode] = 1710", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1710, '«·√À«À Ê«· ÃÂÌ“« ', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1720", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1720, '«·√ÃÂ“… Ê«·Õ«”»«  Ê√‰Ÿ„… ‰ﬁ«ÿ «·»Ì⁄', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1730", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1730, '«·”Ì«—«  ÊÊ”«∆· «·‰ﬁ·', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1740", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1740, '«·œÌﬂÊ—«  Ê Õ”Ì‰«  «·„Õ·', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 1790", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1790, '„Ã„⁄ ≈Â·«ﬂ «·√’Ê· «·À«» …', 'ASSET', 12, True, True)"
    SeedRow "[AccountCode] = 1800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (1800, '«·»—«„Ã Ê«· —«ŒÌ’', 'ASSET', 12, True, False)"
    SeedRow "[AccountCode] = 2", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2, '«·Œ’Ê„', 'LIABILITY', Null, False, True)"
    SeedRow "[AccountCode] = 21", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (21, '«·Œ’Ê„ «·„ œ«Ê·…', 'LIABILITY', 2, False, True)"
    SeedRow "[AccountCode] = 2100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2100, '–„„ «·„Ê—œÌ‰', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2110", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2110, '√Ê—«ﬁ «·œ›⁄ (‘Ìﬂ«  ’«œ—…)', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2200, '÷—Ì»… «·ﬁÌ„… «·„÷«›… - „Œ—Ã« ', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2250", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2250, '÷—Ì»… «·ﬁÌ„… «·„÷«›… - «· ”ÊÌ… Ê«·”œ«œ', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2300, '„’—Ê›«  „” Õﬁ…', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2310", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2310, '—Ê« » Ê√ÃÊ— „” Õﬁ…', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2320", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2320, '«· √„Ì‰«  «·«Ã „«⁄Ì… «·„” Õﬁ…', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2330", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2330, '⁄„Ê·«  „” Õﬁ… ··„‰œÊ»Ì‰', 'LIABILITY', 21, True, True)"
    SeedRow "[AccountCode] = 2400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2400, 'œ›⁄«  „ﬁœ„… „‰ «·⁄„·«¡', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2500, 'ﬁ—Ê÷ ﬁ’Ì—… «·√Ã·', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 2600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2600, '«·“ﬂ«… «·„” Õﬁ…', 'LIABILITY', 21, True, False)"
    SeedRow "[AccountCode] = 22", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (22, '«·Œ’Ê„ €Ì— «·„ œ«Ê·…', 'LIABILITY', 2, False, True)"
    SeedRow "[AccountCode] = 2700", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2700, 'ﬁ—Ê÷ ÿÊÌ·… «·√Ã·', 'LIABILITY', 22, True, False)"
    SeedRow "[AccountCode] = 2800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (2800, '„Œ’’ „ﬂ«›√… ‰Â«Ì… «·Œœ„…', 'LIABILITY', 22, True, False)"
    SeedRow "[AccountCode] = 3", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3, 'ÕﬁÊﬁ «·„·ﬂÌ…', 'EQUITY', Null, False, True)"
    SeedRow "[AccountCode] = 31", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (31, '—√” «·„«· ÊÃ«—Ì «·„«·ﬂ', 'EQUITY', 3, False, True)"
    SeedRow "[AccountCode] = 3200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3200, '—√” «·„«·', 'EQUITY', 31, True, False)"
    SeedRow "[AccountCode] = 3100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3100, 'Ã«—Ì «·„«·ﬂ', 'EQUITY', 31, True, True)"
    SeedRow "[AccountCode] = 3900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3900, '√—’œ… «›  «ÕÌ…', 'EQUITY', 31, True, True)"
    SeedRow "[AccountCode] = 32", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (32, '«·√—»«Õ «·„Õ Ã“… Ê‰ «∆Ã «·√⁄„«·', 'EQUITY', 3, False, True)"
    SeedRow "[AccountCode] = 3300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3300, '«·√—»«Õ «·„Õ Ã“…', 'EQUITY', 32, True, True)"
    SeedRow "[AccountCode] = 3400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (3400, '’«›Ì —»Õ (Œ”«—…) «·⁄«„', 'EQUITY', 32, True, False)"
    SeedRow "[AccountCode] = 4", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4, '«·≈Ì—«œ« ', 'REVENUE', Null, False, True)"
    SeedRow "[AccountCode] = 41", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (41, '≈Ì—«œ«  «·‰‘«ÿ', 'REVENUE', 4, False, True)"
    SeedRow "[AccountCode] = 4100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4100, '«·„»Ì⁄« ', 'REVENUE', 41, True, True)"
    SeedRow "[AccountCode] = 4110", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4110, '„—œÊœ«  «·„»Ì⁄« ', 'REVENUE', 41, True, True)"
    SeedRow "[AccountCode] = 4120", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4120, '«·Œ’„ «·„”„ÊÕ »Â', 'REVENUE', 41, True, False)"
    SeedRow "[AccountCode] = 42", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (42, '≈Ì—«œ«  √Œ—Ï', 'REVENUE', 4, False, True)"
    SeedRow "[AccountCode] = 4200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4200, '≈Ì—«œ«  „ ‰Ê⁄…', 'REVENUE', 42, True, True)"
    SeedRow "[AccountCode] = 4300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4300, '“Ì«œ… «·’‰«œÌﬁ', 'REVENUE', 42, True, True)"
    SeedRow "[AccountCode] = 4400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4400, '«·Œ’„ «·„ﬂ ”»', 'REVENUE', 42, True, False)"
    SeedRow "[AccountCode] = 4500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (4500, '√—»«Õ »Ì⁄ «·√’Ê· «·À«» …', 'REVENUE', 42, True, True)"
    SeedRow "[AccountCode] = 5", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5, '«·„’—Ê›« ', 'EXPENSE', Null, False, True)"
    SeedRow "[AccountCode] = 51", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (51, ' ﬂ·›… «·„»Ì⁄« ', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5100", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5100, ' ﬂ·›… «·»÷«⁄… «·„»«⁄…', 'EXPENSE', 51, True, True)"
    SeedRow "[AccountCode] = 5200", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5200, '›—Êﬁ«  Ê ”ÊÌ«  «·„Œ“Ê‰', 'EXPENSE', 51, True, True)"
    SeedRow "[AccountCode] = 52", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (52, '«·„’—Ê›«  «· ‘€Ì·Ì… Ê«·≈œ«—Ì…', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5300", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5300, '«·„’—Ê›«  Õ”» «·‰Ê⁄', 'EXPENSE', 52, False, True)"
    SeedRow "[AccountCode] = 5500", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5500, '«·—Ê« » Ê«·√ÃÊ—', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5510", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5510, '«·»œ·«  Ê«·ÕÊ«›“', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5520", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5520, '«· √„Ì‰«  «·«Ã „«⁄Ì…', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5530", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5530, '⁄„Ê·«  «·„‰œÊ»Ì‰', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5600", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5600, '≈Â·«ﬂ «·√’Ê· «·À«» …', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5610", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5610, '⁄„Ê·«  «·»‰Êﬂ Ê‰ﬁ«ÿ «·»Ì⁄', 'EXPENSE', 52, True, True)"
    SeedRow "[AccountCode] = 5620", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5620, '«·—”Ê„ «·ÕﬂÊ„Ì… Ê«· —«ŒÌ’', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 5630", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5630, '«·œ⁄«Ì… Ê«·≈⁄·«‰', 'EXPENSE', 52, True, False)"
    SeedRow "[AccountCode] = 53", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (53, '„’—Ê›«  √Œ—Ï', 'EXPENSE', 5, False, True)"
    SeedRow "[AccountCode] = 5400", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5400, '⁄Ã“ «·’‰«œÌﬁ', 'EXPENSE', 53, True, True)"
    SeedRow "[AccountCode] = 5650", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5650, 'Œ”«∆— »Ì⁄ Ê«” »⁄«œ «·√’Ê· «·À«» …', 'EXPENSE', 53, True, True)"
    SeedRow "[AccountCode] = 5700", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5700, '«·œÌÊ‰ «·„⁄œÊ„…', 'EXPENSE', 53, True, False)"
    SeedRow "[AccountCode] = 5800", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5800, '«·“ﬂ«…', 'EXPENSE', 53, True, False)"
    SeedRow "[AccountCode] = 5900", "INSERT INTO [Accounts] ([AccountCode], [AccountName], [AccountType], [ParentCode], [IsPosting], [IsSystem]) VALUES (5900, '„’—Ê›«  „ ‰Ê⁄…', 'EXPENSE', 53, True, True)"
    EndSeed "Accounts", 80
End Sub

Private Sub Seed_JournalSourceTypes()
    If Not BeginSeed("JournalSourceTypes", True) Then Exit Sub
    SeedRow "[SourceType] = 'SALE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SALE', '›« Ê—… »Ì⁄', 1)"
    SeedRow "[SourceType] = 'SALES_RETURN'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SALES_RETURN', '„— Ã⁄ »Ì⁄', 2)"
    SeedRow "[SourceType] = 'PURCHASE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PURCHASE', '›« Ê—… ‘—«¡', 3)"
    SeedRow "[SourceType] = 'PURCHASE_RETURN'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PURCHASE_RETURN', '„— Ã⁄ ‘—«¡', 4)"
    SeedRow "[SourceType] = 'CUSTOMER_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CUSTOMER_PAYMENT', '”‰œ ﬁ»÷ „‰ ⁄„Ì·', 5)"
    SeedRow "[SourceType] = 'SUPPLIER_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SUPPLIER_PAYMENT', '”‰œ ’—› ·„Ê—œ', 6)"
    SeedRow "[SourceType] = 'EXPENSE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('EXPENSE', '„’—Ê›', 7)"
    SeedRow "[SourceType] = 'CASH_VOUCHER'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CASH_VOUCHER', '”‰œ ‰ﬁœÌ…', 8)"
    SeedRow "[SourceType] = 'STOCK_MOVE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('STOCK_MOVE', 'Õ—ﬂ… „Œ“Ê‰ ÌœÊÌ…', 9)"
    SeedRow "[SourceType] = 'STOCK_COUNT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('STOCK_COUNT', ' ”ÊÌ… Ã—œ', 10)"
    SeedRow "[SourceType] = 'BOX_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('BOX_OPENING', '—’Ìœ «›  «ÕÌ ·’‰œÊﬁ', 11)"
    SeedRow "[SourceType] = 'CUSTOMER_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CUSTOMER_OPENING', '—’Ìœ «›  «ÕÌ ·⁄„Ì·', 12)"
    SeedRow "[SourceType] = 'SUPPLIER_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('SUPPLIER_OPENING', '—’Ìœ «›  «ÕÌ ·„Ê—œ', 13)"
    SeedRow "[SourceType] = 'MANUAL'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('MANUAL', 'ﬁÌœ ÌœÊÌ', 14)"
    SeedRow "[SourceType] = 'YEAR_CLOSE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('YEAR_CLOSE', 'ﬁÌœ ≈ﬁ›«· «·”‰…', 15)"
    SeedRow "[SourceType] = 'VAT_RETURN'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('VAT_RETURN', ' ”ÊÌ… ≈ﬁ—«— ÷—Ì»… «·ﬁÌ„… «·„÷«›…', 16)"
    SeedRow "[SourceType] = 'VAT_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('VAT_PAYMENT', '”œ«œ ÷—Ì»… «·ﬁÌ„… «·„÷«›…', 17)"
    SeedRow "[SourceType] = 'BANK_OPENING'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('BANK_OPENING', '—’Ìœ «›  «ÕÌ ·»‰ﬂ', 18)"
    SeedRow "[SourceType] = 'BANK_TX'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('BANK_TX', 'Õ—ﬂ… »‰ﬂÌ…', 19)"
    SeedRow "[SourceType] = 'CHEQUE'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CHEQUE', '‘Ìﬂ Ê«—œ √Ê ’«œ—', 20)"
    SeedRow "[SourceType] = 'CHEQUE_STATUS'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('CHEQUE_STATUS', ' Õ’Ì· √Ê «— œ«œ ‘Ìﬂ', 21)"
    SeedRow "[SourceType] = 'ASSET'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('ASSET', '«ﬁ ‰«¡ √’· À«» ', 22)"
    SeedRow "[SourceType] = 'ASSET_DISPOSAL'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('ASSET_DISPOSAL', '»Ì⁄ √Ê «” »⁄«œ √’· À«» ', 23)"
    SeedRow "[SourceType] = 'DEPRECIATION'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('DEPRECIATION', 'ﬁÌœ «·≈Â·«ﬂ «·‘Â—Ì', 24)"
    SeedRow "[SourceType] = 'PAYROLL'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PAYROLL', 'ﬁÌœ „”Ì— «·—Ê« »', 25)"
    SeedRow "[SourceType] = 'PAYROLL_PAYMENT'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('PAYROLL_PAYMENT', '’—› «·—Ê« »', 26)"
    SeedRow "[SourceType] = 'COMMISSION'", "INSERT INTO [JournalSourceTypes] ([SourceType], [TypeName], [SortOrder]) VALUES ('COMMISSION', 'ﬁÌœ ⁄„Ê·«  «·„‰œÊ»Ì‰', 27)"
    EndSeed "JournalSourceTypes", 27
End Sub

Private Sub Seed_TransactionTypes()
    If Not BeginSeed("TransactionTypes", False) Then Exit Sub
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
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 11 WHERE [AccountCode] = 1210 AND [ParentCode] Is Null", dbFailOnError
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
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 21 WHERE [AccountCode] = 2330 AND [ParentCode] Is Null", dbFailOnError
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
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 42 WHERE [AccountCode] = 4500 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 51 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 51 WHERE [AccountCode] = 5100 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 51 WHERE [AccountCode] = 5200 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 52 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5300 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5500 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5510 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5520 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5530 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5600 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5610 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5620 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 52 WHERE [AccountCode] = 5630 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 5 WHERE [AccountCode] = 53 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5400 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5650 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5700 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5800 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [ParentCode] = 53 WHERE [AccountCode] = 5900 AND [ParentCode] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [IsPosting] = False WHERE [AccountCode] IN (1, 11, 1100, 1210, 12, 2, 21, 22, 3, 31, 32, 4, 41, 42, 5, 51, 52, 5300, 53)", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [IsSystem] = True WHERE [AccountCode] IN (1, 11, 1100, 110001, 110002, 1190, 1200, 1210, 1250, 1300, 1400, 1500, 1600, 12, 1790, 2, 21, 2100, 2110, 2200, 2250, 2310, 2320, 2330, 22, 3, 31, 3100, 3900, 32, 3300, 4, 41, 4100, 4110, 42, 4200, 4300, 4500, 5, 51, 5100, 5200, 52, 5300, 5500, 5510, 5520, 5530, 5600, 5610, 53, 5400, 5650, 5900) OR [AccountCode] BETWEEN 110001 AND 119999 OR [AccountCode] BETWEEN 120001 AND 129999 OR [AccountCode] BETWEEN 530001 AND 539999", dbFailOnError
End Sub

Private Sub SeedEnglishNames()
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Assets' WHERE [AccountCode] = 1 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Current assets' WHERE [AccountCode] = 11 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cash in treasury and boxes' WHERE [AccountCode] = 1100 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Main treasury' WHERE [AccountCode] = 110001 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cashier box' WHERE [AccountCode] = 110002 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cash not allocated to a box' WHERE [AccountCode] = 1190 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Mada and wallets under settlement (bank and network)' WHERE [AccountCode] = 1200 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Bank accounts' WHERE [AccountCode] = 1210 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cheques under collection' WHERE [AccountCode] = 1250 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Accounts receivable' WHERE [AccountCode] = 1300 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Notes receivable' WHERE [AccountCode] = 1310 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Allowance for doubtful debts' WHERE [AccountCode] = 1350 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Inventory' WHERE [AccountCode] = 1400 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'VAT - input' WHERE [AccountCode] = 1500 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Employee advances' WHERE [AccountCode] = 1600 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cash floats' WHERE [AccountCode] = 1610 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Prepaid expenses' WHERE [AccountCode] = 1650 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Accrued revenue' WHERE [AccountCode] = 1660 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Deposits with others' WHERE [AccountCode] = 1690 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Non-current assets' WHERE [AccountCode] = 12 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Furniture and fixtures' WHERE [AccountCode] = 1710 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Computers, devices and POS systems' WHERE [AccountCode] = 1720 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Vehicles' WHERE [AccountCode] = 1730 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Decorations and store improvements' WHERE [AccountCode] = 1740 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Accumulated depreciation of fixed assets' WHERE [AccountCode] = 1790 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Software and licenses' WHERE [AccountCode] = 1800 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Liabilities' WHERE [AccountCode] = 2 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Current liabilities' WHERE [AccountCode] = 21 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Accounts payable' WHERE [AccountCode] = 2100 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Notes payable (issued cheques)' WHERE [AccountCode] = 2110 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'VAT - output' WHERE [AccountCode] = 2200 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'VAT - settlement and payment' WHERE [AccountCode] = 2250 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Accrued expenses' WHERE [AccountCode] = 2300 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Salaries and wages payable' WHERE [AccountCode] = 2310 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Social insurance payable' WHERE [AccountCode] = 2320 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Sales rep commissions payable' WHERE [AccountCode] = 2330 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Customer advances' WHERE [AccountCode] = 2400 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Short-term loans' WHERE [AccountCode] = 2500 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Zakat payable' WHERE [AccountCode] = 2600 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Non-current liabilities' WHERE [AccountCode] = 22 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Long-term loans' WHERE [AccountCode] = 2700 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'End-of-service benefits provision' WHERE [AccountCode] = 2800 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Equity' WHERE [AccountCode] = 3 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Capital and owner current account' WHERE [AccountCode] = 31 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Capital' WHERE [AccountCode] = 3200 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Owner current account' WHERE [AccountCode] = 3100 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Opening balances' WHERE [AccountCode] = 3900 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Retained earnings and results' WHERE [AccountCode] = 32 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Retained earnings' WHERE [AccountCode] = 3300 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Net profit (loss) of the year' WHERE [AccountCode] = 3400 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Revenue' WHERE [AccountCode] = 4 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Operating revenue' WHERE [AccountCode] = 41 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Sales' WHERE [AccountCode] = 4100 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Sales returns' WHERE [AccountCode] = 4110 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Discounts allowed' WHERE [AccountCode] = 4120 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Other revenue' WHERE [AccountCode] = 42 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Miscellaneous revenue' WHERE [AccountCode] = 4200 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cash overages' WHERE [AccountCode] = 4300 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Discounts received' WHERE [AccountCode] = 4400 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Gain on sale of fixed assets' WHERE [AccountCode] = 4500 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Expenses' WHERE [AccountCode] = 5 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cost of sales' WHERE [AccountCode] = 51 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cost of goods sold' WHERE [AccountCode] = 5100 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Stock differences and adjustments' WHERE [AccountCode] = 5200 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Operating and administrative expenses' WHERE [AccountCode] = 52 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Expenses by type' WHERE [AccountCode] = 5300 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Salaries and wages' WHERE [AccountCode] = 5500 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Allowances and incentives' WHERE [AccountCode] = 5510 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Social insurance' WHERE [AccountCode] = 5520 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Sales rep commissions' WHERE [AccountCode] = 5530 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Depreciation of fixed assets' WHERE [AccountCode] = 5600 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Bank and POS fees' WHERE [AccountCode] = 5610 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Government fees and licenses' WHERE [AccountCode] = 5620 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Advertising' WHERE [AccountCode] = 5630 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Other expenses' WHERE [AccountCode] = 53 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Cash shortages' WHERE [AccountCode] = 5400 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Loss on sale and disposal of fixed assets' WHERE [AccountCode] = 5650 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Bad debts' WHERE [AccountCode] = 5700 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Zakat' WHERE [AccountCode] = 5800 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [Accounts] SET [AccountNameEn] = 'Miscellaneous expenses' WHERE [AccountCode] = 5900 AND [AccountNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [PaymentMethods] SET [MethodNameEn] = 'Cash' WHERE [PaymentMethodID] = 1 AND [MethodNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [PaymentMethods] SET [MethodNameEn] = 'Mada / card' WHERE [PaymentMethodID] = 2 AND [MethodNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [PaymentMethods] SET [MethodNameEn] = 'Bank transfer' WHERE [PaymentMethodID] = 3 AND [MethodNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [PaymentMethods] SET [MethodNameEn] = 'E-wallet' WHERE [PaymentMethodID] = 4 AND [MethodNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Sales invoice' WHERE [SourceType] = 'SALE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Sales return' WHERE [SourceType] = 'SALES_RETURN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Purchase invoice' WHERE [SourceType] = 'PURCHASE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Purchase return' WHERE [SourceType] = 'PURCHASE_RETURN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Customer receipt' WHERE [SourceType] = 'CUSTOMER_PAYMENT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Supplier payment' WHERE [SourceType] = 'SUPPLIER_PAYMENT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Expense' WHERE [SourceType] = 'EXPENSE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Cash voucher' WHERE [SourceType] = 'CASH_VOUCHER' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Manual stock move' WHERE [SourceType] = 'STOCK_MOVE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Stock count adjustment' WHERE [SourceType] = 'STOCK_COUNT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Box opening balance' WHERE [SourceType] = 'BOX_OPENING' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Customer opening balance' WHERE [SourceType] = 'CUSTOMER_OPENING' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Supplier opening balance' WHERE [SourceType] = 'SUPPLIER_OPENING' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Manual entry' WHERE [SourceType] = 'MANUAL' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Year closing entry' WHERE [SourceType] = 'YEAR_CLOSE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'VAT return settlement' WHERE [SourceType] = 'VAT_RETURN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'VAT payment' WHERE [SourceType] = 'VAT_PAYMENT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Bank opening balance' WHERE [SourceType] = 'BANK_OPENING' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Bank transaction' WHERE [SourceType] = 'BANK_TX' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Received or issued cheque' WHERE [SourceType] = 'CHEQUE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Cheque collection or bounce' WHERE [SourceType] = 'CHEQUE_STATUS' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Fixed asset purchase' WHERE [SourceType] = 'ASSET' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Fixed asset sale or disposal' WHERE [SourceType] = 'ASSET_DISPOSAL' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Monthly depreciation entry' WHERE [SourceType] = 'DEPRECIATION' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Payroll entry' WHERE [SourceType] = 'PAYROLL' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Salary payment' WHERE [SourceType] = 'PAYROLL_PAYMENT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [JournalSourceTypes] SET [TypeNameEn] = 'Sales rep commissions entry' WHERE [SourceType] = 'COMMISSION' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Purchase' WHERE [TypeCode] = 'PURCHASE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Sale' WHERE [TypeCode] = 'SALE' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Purchase return' WHERE [TypeCode] = 'PURCHASE_RETURN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Sales return' WHERE [TypeCode] = 'SALES_RETURN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Stock addition' WHERE [TypeCode] = 'STOCK_IN' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Stock deduction' WHERE [TypeCode] = 'STOCK_OUT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Count adjustment' WHERE [TypeCode] = 'ADJUSTMENT' AND [TypeNameEn] Is Null", dbFailOnError
    m_db.Execute "UPDATE [TransactionTypes] SET [TypeNameEn] = 'Opening balance' WHERE [TypeCode] = 'OPENING' AND [TypeNameEn] Is Null", dbFailOnError
End Sub
