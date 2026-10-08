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
        CountryName = "„’—"
    Else
        CountryName = "«·”⁄ÊœÌ…"
    End If
End Function

Public Function TaxNumberProblem(ByVal Value As Variant, Optional ByVal Country As String = "") As String
    ' "" when the tax number fits the country (an empty number is allowed: not VAT-registered).
    Dim s As String
    s = Trim$(Nz(Value, ""))
    If Len(s) = 0 Then Exit Function
    If Len(Country) = 0 Then Country = AppCountry()
    If Country = "EG" Then
        If Not (s Like "#########") Then TaxNumberProblem = "—ﬁ„ «· ”ÃÌ· «·÷—Ì»Ì ›Ì „’— 9 √—ﬁ«„."
    Else
        If Not (s Like "3#############3") Then
            TaxNumberProblem = "«·—ﬁ„ «·÷—Ì»Ì ›Ì «·”⁄ÊœÌ… 15 —ﬁ„« ÊÌ»œ√ ÊÌ‰ ÂÌ »«·—ﬁ„ 3."
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
    ' The program currency after an amount: "—Ì«·" / "Ã‰ÌÂ", "SAR" / "EGP" in English (modLang).
    If UiEnglish() Then
        CurrencyWord = BaseCurrency()
    ElseIf BaseCurrency() = "EGP" Then
        CurrencyWord = "Ã‰ÌÂ"
    Else
        CurrencyWord = "—Ì«·"
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
        DocTitleFor = Split("≈‘⁄«— œ«∆‰|›« Ê—… ÷—Ì»Ì…|›« Ê—… ÷—Ì»Ì… „»”ÿ…|≈Ì’«· »Ì⁄", "|")(i)
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
            CountryChangeProblem = "·« Ì„ﬂ‰  €ÌÌ— œÊ·… «· ‘€Ì· »⁄œ  ”ÃÌ· ⁄„·Ì«  ›Ì «·»—‰«„Ã." & vbCrLf & _
                                   "··⁄„· ›Ì «·œÊ·… «·√Œ—Ï «»œ√ „·› »Ì«‰«  ÃœÌœ«."
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
    If Not AskYesNo(" €ÌÌ— œÊ·… «· ‘€Ì· ≈·Ï " & CountryName(newCountry) & ":" & vbCrLf & _
                    "⁄„·… «·»—‰«„Ã " & CountryCurrency(newCountry) & "° Ê‰”»… «·÷—Ì»… " & _
                    Format$(CountryVatRate(newCountry) * 100, "0") & "%° Ê«·›« Ê—… «·≈·ﬂ —Ê‰Ì… ·Â–Â «·œÊ·…." & vbCrLf & _
                    "Â·  —Ìœ «·„ «»⁄…ø") Then
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
        ShowInfo " „ ÷»ÿ «·»—‰«„Ã ⁄·Ï " & CountryName(AppCountry()) & ". √⁄œ › Õ «·‘«‘«  «·„› ÊÕ…."
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
    ApplyCountryData = " ⁄–¯— ÷»ÿ ⁄„·… „·› «·»Ì«‰«  (" & Err.Description & ")." & vbCrLf & _
                       "√€·ﬁ «·»—‰«„Ã ⁄·Ï «·√ÃÂ“… «·√Œ—Ï À„ «Õ›Ÿ «·≈⁄œ«œ«  „—… √Œ—Ï."
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
    Record country = "SA" Or country = "EG", "œÊ·… «· ‘€Ì·: " & country, passed, failed, report
    Record Len(TaxNumberProblem("300000000000003", "SA")) = 0 And Len(TaxNumberProblem("123456789", "SA")) > 0, _
           "«·—ﬁ„ «·÷—Ì»Ì «·”⁄ÊœÌ", passed, failed, report
    Record Len(TaxNumberProblem("123456789", "EG")) = 0 And Len(TaxNumberProblem("300000000000003", "EG")) > 0, _
           "—ﬁ„ «· ”ÃÌ· «·÷—Ì»Ì «·„’—Ì", passed, failed, report
    Record Len(TaxNumberProblem(Null)) = 0, "«·—ﬁ„ «·÷—Ì»Ì «·›«—€ „ﬁ»Ê·", passed, failed, report
    Record BaseCurrency() = CountryCurrency(country), "⁄„·… «·»—‰«„Ã = ⁄„·… «·œÊ·… " & BaseCurrency(), passed, failed, report
    Record DocTitleFor("EG", "SALE", "SIMPLIFIED", False) = "≈Ì’«· »Ì⁄" And _
           DocTitleFor("SA", "SALE", "SIMPLIFIED", False) = "›« Ê—… ÷—Ì»Ì… „»”ÿ…", "⁄‰Ê«‰ «·„” ‰œ Õ”» «·œÊ·…", _
           passed, failed, report
    Record AmountInWordsAr(1250.5, "EGP") = "›ﬁÿ √·› Ê„«∆ «‰ ÊŒ„”Ê‰ Ã‰ÌÂ „’—Ì ÊŒ„”Ê‰ ﬁ—‘ ·« €Ì—", _
           "«·„»·€ »«·Õ—Ê› »«·Ã‰ÌÂ", passed, failed, report
    If Nz(DbValue("SELECT COUNT(*) FROM SalesInvoices"), 0) > 0 Then
        Record Len(CountryChangeProblem()) > 0, "·«  €ÌÌ— ··œÊ·… »⁄œ  ”ÃÌ· ›Ê« Ì—", passed, failed, report
    End If
    GoTo Done
EH:
    Record False, "Œÿ√: " & Err.Description, passed, failed, report
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  œÊ·… «· ‘€Ì· ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestCountry"
        TestCountry = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestCountry"
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
