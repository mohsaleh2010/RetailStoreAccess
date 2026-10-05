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

Public Sub SetQueryParam(ByVal ParamName As String, ByVal ParamValue As Variant)
    TempVars.Add ParamName, ParamValue
End Sub

Public Sub ClearQueryParams()
    Dim names As Variant, i As Long
    names = Array("PeriodStart", "PeriodEnd", "CustomerID", "SupplierID", "ProductID", "CashBoxID")
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
