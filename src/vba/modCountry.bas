Attribute VB_Name = "modCountry"
'==============================================================================
' modCountry  -  Retail Store Management System (the operating country)
'
' Settings.CountryCode (docs/44-Operating-Country.md): SA = Saudi Arabia (ZATCA), EG = Egypt (ETA).
' It sets the program currency (SAR / EGP), the VAT rate (15% / 14%), the form of the tax number,
' the amount in words, the title and QR code of the sales documents and, in the next phases
' (docs/43-Plan-EInvoicing-SA-EG.md), the e-invoicing platform. The country can change only while
' the data file has no operation. tools/country_reference.py is the Python mirror.
'==============================================================================
Option Compare Database
Option Explicit

' a row in one of these tables means the data file is in use: the country cannot change any more
Private Const OPERATION_TABLES As String = "SalesInvoices,SalesReturns,PurchaseInvoices,PurchaseReturns," & _
    "CustomerPayments,SupplierPayments,Expenses,CashVouchers,ManualEntries,JournalEntries,BankTransactions," & _
    "Cheques,FixedAssets,PayrollRuns,InventoryTransactions,StockCounts,CashClosings,CommissionRuns"

Private m_newCountry As String          ' set by ValidateSettings when the country changes, applied after the save

'==============================================================================
' The country and what it decides
'==============================================================================
Public Function AppCountry() As String
    AppCountry = UCase$(Nz(SettingValue("CountryCode"), "SA"))
    If AppCountry <> "EG" Then AppCountry = "SA"
End Function

Public Function IsEgypt() As Boolean
    IsEgypt = (AppCountry() = "EG")
End Function

Public Function CountryVatRate(ByVal Country As String) As Double
    If Country = "EG" Then
        CountryVatRate = 0.14
    Else
        CountryVatRate = 0.15
    End If
End Function

Public Function CountryCurrency(ByVal Country As String) As String
    If Country = "EG" Then
        CountryCurrency = "EGP"
    Else
        CountryCurrency = "SAR"
    End If
End Function

Public Function CountryName(ByVal Country As String) As String
    If Country = "EG" Then
        CountryName = "مصر"
    Else
        CountryName = "السعودية"
    End If
End Function

Public Function TaxNumberProblem(ByVal Value As Variant, Optional ByVal Country As String = "") As String
    ' "" when the tax number fits the country (an empty number is allowed: not VAT-registered).
    Dim s As String
    s = Trim$(Nz(Value, ""))
    If Len(s) = 0 Then Exit Function
    If Len(Country) = 0 Then Country = AppCountry()
    If Country = "EG" Then
        If Not (s Like "#########") Then TaxNumberProblem = "رقم التسجيل الضريبي في مصر 9 أرقام."
    Else
        If Not (s Like "3#############3") Then
            TaxNumberProblem = "الرقم الضريبي في السعودية 15 رقمًا ويبدأ وينتهي بالرقم 3."
        End If
    End If
End Function

Public Function MobileFits(ByVal Mobile As String, Optional ByVal Country As String = "") As Boolean
    ' A mobile number of the country: 05xxxxxxxx (Saudi Arabia), 01xxxxxxxxx (Egypt), or with the country code.
    If Len(Country) = 0 Then Country = AppCountry()
    If Country = "EG" Then
        MobileFits = Mobile Like "01#########" Or Mobile Like "+201#########" Or Mobile Like "201#########"
    Else
        MobileFits = Mobile Like "05########" Or Mobile Like "+9665########" Or Mobile Like "9665########"
    End If
End Function

Public Function MobileExample(Optional ByVal Country As String = "") As String
    If Len(Country) = 0 Then Country = AppCountry()
    If Country = "EG" Then
        MobileExample = "01xxxxxxxxx"
    Else
        MobileExample = "05xxxxxxxx"
    End If
End Function

Public Function CurrencyWord() As String
    ' The program currency after an amount: "ريال" / "جنيه", "SAR" / "EGP" in English (modLang).
    If UiEnglish() Then
        CurrencyWord = BaseCurrency()
    ElseIf BaseCurrency() = "EGP" Then
        CurrencyWord = "جنيه"
    Else
        CurrencyWord = "ريال"
    End If
End Function

'==============================================================================
' Title of the sales documents (rptSalesReceipt, rptSalesInvoiceA4)
'==============================================================================
Public Function DocTitleFor(ByVal Country As String, ByVal DocKind As Variant, ByVal SubType As Variant, _
                            ByVal English As Boolean) As String
    ' Saudi Arabia: simplified tax invoice / tax invoice / credit note (ZATCA).
    ' Egypt: sales receipt (e-receipt) / tax invoice (e-invoice) / credit note.
    Dim i As Long
    If Nz(DocKind, "") = "RETURN" Then
        i = 0
    ElseIf Nz(SubType, "") = "STANDARD" Then
        i = 1
    ElseIf Country = "EG" Then
        i = 3
    Else
        i = 2
    End If
    If English Then
        DocTitleFor = Split("Credit Note|Tax Invoice|Simplified Tax Invoice|Sales Receipt", "|")(i)
    Else
        DocTitleFor = Split("إشعار دائن|فاتورة ضريبية|فاتورة ضريبية مبسطة|إيصال بيع", "|")(i)
    End If
End Function

Public Function DocTitleAr(ByVal DocKind As Variant, ByVal SubType As Variant) As String
    DocTitleAr = DocTitleFor(AppCountry(), DocKind, SubType, False)
End Function

Public Function DocTitleEn(ByVal DocKind As Variant, ByVal SubType As Variant) As String
    DocTitleEn = DocTitleFor(AppCountry(), DocKind, SubType, True)
End Function

'==============================================================================
' Changing the country (frmSettings)
'==============================================================================
Public Function CountryChangeProblem() As String
    ' "" while the data file has no operation; afterwards the country stays (a new data file is needed).
    Dim t As Variant
    For Each t In Split(OPERATION_TABLES, ",")
        If Nz(DbValue("SELECT COUNT(*) FROM [" & t & "]"), 0) > 0 Then
            CountryChangeProblem = "لا يمكن تغيير دولة التشغيل بعد تسجيل عمليات في البرنامج." & vbCrLf & _
                                   "للعمل في الدولة الأخرى ابدأ ملف بيانات جديدًا."
            Exit Function
        End If
    Next
End Function

Public Function SettingsCountryCheck(ByVal frm As Access.Form) As Boolean
    ' From ValidateSettings (modForms): the country changed -> allowed now? the rate and currency follow it.
    Dim newCountry As String, msg As String
    m_newCountry = ""
    newCountry = Nz(frm!CountryCode.Value, "SA")
    If newCountry = Nz(frm!CountryCode.OldValue, "SA") Then
        SettingsCountryCheck = True
        Exit Function
    End If
    msg = CountryChangeProblem()
    If Len(msg) > 0 Then
        ShowWarning msg
        SafeFocus frm!CountryCode
        Exit Function
    End If
    If Not AskYesNo("تغيير دولة التشغيل إلى " & CountryName(newCountry) & ":" & vbCrLf & _
                    "عملة البرنامج " & CountryCurrency(newCountry) & "، ونسبة الضريبة " & _
                    Format$(CountryVatRate(newCountry) * 100, "0") & "%، والفاتورة الإلكترونية لهذه الدولة." & vbCrLf & _
                    "هل تريد المتابعة؟") Then
        SafeFocus frm!CountryCode
        Exit Function
    End If
    frm!VATRate.Value = CountryVatRate(newCountry)
    frm!CurrencyCode.Value = CountryCurrency(newCountry)
    m_newCountry = newCountry
    SettingsCountryCheck = True
End Function

Public Sub SettingsAfterSave()
    ' From FormAfterUpdate (modForms): the data file follows a changed country.
    Dim msg As String
    If Len(m_newCountry) = 0 Then Exit Sub
    msg = ApplyCountryData(m_newCountry)
    m_newCountry = ""
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo "تم ضبط البرنامج على " & CountryName(AppCountry()) & ". أعد فتح الشاشات المفتوحة."
    End If
End Sub

Public Function ApplyCountryData(ByVal Country As String) As String
    ' The currency of the country becomes the default of every currency field of the data file, the
    ' currency of the suppliers and an active currency. "" = done, else the problem.
    Dim db As DAO.Database, be As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
    Dim cur As String, prev As String, cnn As String
    On Error GoTo EH
    cur = CountryCurrency(Country)
    prev = CountryCurrency(IIf(Country = "EG", "SA", "EG"))
    Set db = CurrentDb
    db.Execute "UPDATE Currencies SET IsActive = True WHERE CurrencyCode = '" & cur & "'", dbFailOnError
    db.Execute "UPDATE Suppliers SET CurrencyCode = '" & cur & "' WHERE CurrencyCode = '" & prev & "'", dbFailOnError
    Set tdf = db.TableDefs("Settings")
    cnn = tdf.Connect
    If InStr(cnn, "DATABASE=") > 0 Then                    ' the data file of the linked tables
        Set be = DBEngine.OpenDatabase(Mid$(cnn, InStr(cnn, "DATABASE=") + 9), False, False)
    Else
        Set be = db                                        ' the tables in this file (development)
    End If
    For Each tdf In be.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" Then
            For Each fld In tdf.Fields
                If fld.Name = "CurrencyCode" And fld.DefaultValue = """" & prev & """" Then
                    fld.DefaultValue = """" & cur & """"
                End If
            Next
        End If
    Next
    If InStr(cnn, "DATABASE=") > 0 Then be.Close
    Exit Function
EH:
    ApplyCountryData = "تعذّر ضبط عملة ملف البيانات (" & Err.Description & ")." & vbCrLf & _
                       "أغلق البرنامج على الأجهزة الأخرى ثم احفظ الإعدادات مرة أخرى."
End Function

'==============================================================================
' In-Access test
'==============================================================================
Public Function TestCountry() As Boolean
    Dim passed As Long, failed As Long, report As String, country As String
    g_SilentMode = True
    Debug.Print "=== TestCountry  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    country = AppCountry()
    Record country = "SA" Or country = "EG", "دولة التشغيل: " & country, passed, failed, report
    Record Len(TaxNumberProblem("300000000000003", "SA")) = 0 And Len(TaxNumberProblem("123456789", "SA")) > 0, _
           "الرقم الضريبي السعودي", passed, failed, report
    Record Len(TaxNumberProblem("123456789", "EG")) = 0 And Len(TaxNumberProblem("300000000000003", "EG")) > 0, _
           "رقم التسجيل الضريبي المصري", passed, failed, report
    Record Len(TaxNumberProblem(Null)) = 0, "الرقم الضريبي الفارغ مقبول", passed, failed, report
    Record BaseCurrency() = CountryCurrency(country), "عملة البرنامج = عملة الدولة " & BaseCurrency(), passed, failed, report
    Record DocTitleFor("EG", "SALE", "SIMPLIFIED", False) = "إيصال بيع" And _
           DocTitleFor("SA", "SALE", "SIMPLIFIED", False) = "فاتورة ضريبية مبسطة", "عنوان المستند حسب الدولة", _
           passed, failed, report
    Record AmountInWordsAr(1250.5, "EGP") = "فقط ألف ومائتان وخمسون جنيه مصري وخمسون قرش لا غير", _
           "المبلغ بالحروف بالجنيه", passed, failed, report
    If Nz(DbValue("SELECT COUNT(*) FROM SalesInvoices"), 0) > 0 Then
        Record Len(CountryChangeProblem()) > 0, "لا تغيير للدولة بعد تسجيل فواتير", passed, failed, report
    End If
    GoTo Done
EH:
    Record False, "خطأ: " & Err.Description, passed, failed, report
Done:
    g_SilentMode = False
    Debug.Print "--- نجح: " & passed & " | فشل: " & failed
    If failed = 0 Then
        TestMsg "جميع اختبارات دولة التشغيل ناجحة (" & passed & " اختبارًا).", vbInformation + MSG_RTL, "TestCountry"
        TestCountry = True
    Else
        TestMsg "نجح " & passed & " وفشل " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCountry"
    End If
End Function

Private Sub Record(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
                   ByRef report As String)
    If ok Then
        passed = passed + 1
        Debug.Print "[OK] " & Title
    Else
        failed = failed + 1
        report = report & "- " & Title & vbCrLf
        Debug.Print "[X]  " & Title
    End If
End Sub
