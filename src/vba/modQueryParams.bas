Attribute VB_Name = "modQueryParams"
'==============================================================================
' modQueryParams  -  Retail Store Management System (Phase 4)
'
' Runtime parameters for saved queries and reports.
' Queries read their parameters through QDate()/QLong(), which work the same
' when a query is opened from the navigation pane, by a report, or from VBA
' through DAO (unlike [Forms]! references, which fail in DAO recordsets).
'
' Usage from VBA:
'     SetPeriod #2026-09-01#, #2026-09-30#     ' both days included
'     SetQueryParam "CustomerID", 5
'     DoCmd.OpenQuery "CustomerStatementQuery"
'
' Parameters used by the queries:
'     PeriodStart  first day of the period (inclusive)
'     PeriodEnd    day AFTER the last day of the period (exclusive)
'     CustomerID   customer for CustomerStatementQuery
'     SupplierID   supplier for SupplierStatementQuery
'     ProductID    product for ProductMovementQuery
'     CashBoxID    cash box for the treasury queries (0 = all boxes)
'     AccountCode  account of the statement / ledger: a main account takes all its
'                  sub-accounts (0 = every account, general ledger only)
'     CompareStart / CompareEnd   the period the financial statements compare with
'                  (SetComparePeriod: the same months one step back, else the same length before)
'
' On SQL Server (docs/71) the saved queries are server views: SetQueryParams also sends the values to the
' server (modServerQueries.SendQueryParams). Set a parameter only through this module.
'==============================================================================
Option Compare Database
Option Explicit

Public Sub SetPeriod(ByVal FromDate As Date, ByVal ToDate As Date)
    ' Both FromDate and ToDate are included in the period.
    If ToDate < FromDate Then Err.Raise vbObjectError + 600, "SetPeriod", _
        "تاريخ النهاية يجب أن يكون بعد تاريخ البداية"
    SetQueryParams "PeriodStart", DateValue(FromDate), "PeriodEnd", DateAdd("d", 1, DateValue(ToDate))
End Sub

Public Sub SetComparePeriod(ByVal FromDate As Date, ByVal ToDate As Date)
    ' The period before FromDate..ToDate: from the first of a month, the same months one step back
    ' (a month -> the month before, Jan 1 - today -> the same dates last year); otherwise the same
    ' number of days just before. CompareEnd is exclusive like PeriodEnd.
    Dim k As Long, days As Long, cFrom As Date, cTo As Date
    FromDate = DateValue(FromDate)
    ToDate = DateValue(ToDate)
    If Day(FromDate) = 1 Then
        k = (Year(ToDate) - Year(FromDate)) * 12 + Month(ToDate) - Month(FromDate) + 1
        cFrom = DateAdd("m", -k, FromDate)
        cTo = DateAdd("m", -k, ToDate)
        If cTo >= FromDate Then cTo = DateAdd("d", -1, FromDate)
    Else
        days = ToDate - FromDate + 1
        cTo = DateAdd("d", -1, FromDate)
        cFrom = DateAdd("d", -days + 1, cTo)
    End If
    SetQueryParams "CompareStart", cFrom, "CompareEnd", DateAdd("d", 1, cTo)
End Sub

Public Sub SetQueryParam(ByVal ParamName As String, ByVal ParamValue As Variant)
    SetQueryParams ParamName, ParamValue
End Sub

Public Sub SetQueryParams(ParamArray NameValue() As Variant)
    ' Name, value, name, value ... in one step: on SQL Server the views read them too (modServerQueries, docs/71),
    ' so every parameter of a saved query is set here, never with TempVars.Add elsewhere.
    Dim i As Long, n As Long, names() As Variant, values() As Variant
    n = (UBound(NameValue) - LBound(NameValue) + 1) \ 2
    If n = 0 Then Exit Sub
    ReDim names(n - 1)
    ReDim values(n - 1)
    For i = 0 To n - 1
        names(i) = NameValue(LBound(NameValue) + 2 * i)
        values(i) = NameValue(LBound(NameValue) + 2 * i + 1)
        TempVars.Add CStr(names(i)), values(i)
    Next
    SendQueryParams names, values
End Sub

Public Sub RemoveQueryParams(ByVal Names As Variant)
    ' The parameters back to "not set" (QDate = 1899-12-30, QLong = 0), here and on the server.
    Dim i As Long, values() As Variant
    ReDim values(LBound(Names) To UBound(Names))
    On Error Resume Next
    For i = LBound(Names) To UBound(Names)
        TempVars.Remove CStr(Names(i))
        values(i) = Null
    Next
    On Error GoTo 0
    SendQueryParams Names, values
End Sub

Public Sub ClearQueryParams()
    RemoveQueryParams Array("PeriodStart", "PeriodEnd", "CustomerID", "SupplierID", "ProductID", "CashBoxID", _
                            "AccountCode", "CompareStart", "CompareEnd")
End Sub

Public Function QDate(ByVal ParamName As String) As Date
    ' Returns the parameter as a date, or 1899-12-30 when it is not set.
    Dim v As Variant
    On Error Resume Next
    v = TempVars(ParamName)
    If IsDate(v) Then QDate = CDate(v)
End Function

Public Function QLong(ByVal ParamName As String) As Long
    ' Returns the parameter as a whole number, or 0 when it is not set.
    Dim v As Variant
    On Error Resume Next
    v = TempVars(ParamName)
    If IsNumeric(v) Then QLong = CLng(v)
End Function
