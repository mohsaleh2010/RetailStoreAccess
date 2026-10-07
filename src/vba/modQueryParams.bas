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
'==============================================================================
Option Compare Database
Option Explicit

Public Sub SetPeriod(ByVal FromDate As Date, ByVal ToDate As Date)
    ' Both FromDate and ToDate are included in the period.
    If ToDate < FromDate Then Err.Raise vbObjectError + 600, "SetPeriod", _
        "تاريخ النهاية يجب أن يكون بعد تاريخ البداية"
    TempVars.Add "PeriodStart", DateValue(FromDate)
    TempVars.Add "PeriodEnd", DateAdd("d", 1, DateValue(ToDate))
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
    TempVars.Add "CompareStart", cFrom
    TempVars.Add "CompareEnd", DateAdd("d", 1, cTo)
End Sub

Public Sub SetQueryParam(ByVal ParamName As String, ByVal ParamValue As Variant)
    TempVars.Add ParamName, ParamValue
End Sub

Public Sub ClearQueryParams()
    Dim names As Variant, i As Long
    names = Array("PeriodStart", "PeriodEnd", "CustomerID", "SupplierID", "ProductID", "CashBoxID", "AccountCode", _
                  "CompareStart", "CompareEnd")
    On Error Resume Next
    For i = LBound(names) To UBound(names)
        TempVars.Remove names(i)
    Next
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
