Attribute VB_Name = "modAging"
'==============================================================================
' modAging  -  Retail Store Management System
'
' Receivables and payables by document:
'   ComputeAging     what each customer ("C") or supplier ("S") still owes on a day,
'                    invoice by invoice (and the opening balance):
'                      1. a return pays its own invoice, a payment the invoices it is
'                         linked to (CustomerAllocations / SupplierAllocations);
'                      2. everything else (unlinked payments, the rest of a return,
'                         a credit opening balance) pays the oldest due documents first.
'                    The due date is DueDate, else the document date + the terms of the
'                    party (PaymentTermsDays). What is left is a credit of the party.
'   FillAging        the result in the local table tmpAging with the age columns
'                    (not due, 1-30, 31-60, 61-90, over 90 days late) for frmAging and
'                    rptAging.
'   MaxDaysLate      the oldest late invoice of a party (credit sales stop after
'                    Settings.CreditBlockDays, modSales.CheckCustomer).
'   AllocatePayment  links part of a payment to an invoice (frmAllocation);
'   AutoAllocate     links every free payment amount to the oldest due invoices.
' Same steps as tools/sim.py (Store.aging).
'==============================================================================
Option Compare Database
Option Explicit

Private Const DUE_KEY As String = "IIf(DueDate Is Null, DateAdd('d', Nz(TermsDays, 0), DateValue(DocDate)), DueDate)"

'------------------------------------------------------------------------------
' Engine
'------------------------------------------------------------------------------
Public Function ComputeAging(ByVal Kind As String, ByVal AsOf As Date, Optional ByVal PartyID As Long = 0) As Collection
    ' Items: Array(PartyID, DocType, DocID, DocNo, DocDate, DueDate, OpenAmount); a credit rest to a party is
    ' an item with DocType "CREDIT" and a negative amount.
    Dim db As DAO.Database, rs As DAO.Recordset, where As String, nextDay As String
    Dim docs As Collection, openAmt As Object, pool As Object, links As Object, result As Collection
    Dim d As Variant, key As String, rest As Currency, part As Currency, k As Variant, l As Variant
    Set db = CurrentDb
    Set docs = New Collection
    Set openAmt = CreateObject("Scripting.Dictionary")
    Set pool = CreateObject("Scripting.Dictionary")
    Set links = CreateObject("Scripting.Dictionary")
    Set result = New Collection
    nextDay = SqlDate(DateValue(AsOf) + 1)
    where = "PartyKind = " & SqlText(Kind) & IIf(PartyID > 0, " AND PartyID = " & PartyID, "")

    Set rs = db.OpenRecordset("SELECT PartyID, DocType, DocID, DocNo, DocDate, " & DUE_KEY & " AS DueKey, Amount " & _
        "FROM qryAgingDebits WHERE " & where & " AND DocDate < " & nextDay & " ORDER BY PartyID, " & DUE_KEY & _
        ", DocDate, IIf(DocType = 'OPENING', 0, 1), DocID", dbOpenSnapshot)
    Do Until rs.EOF
        docs.Add Array(rs!PartyID.Value, rs!DocType.Value, rs!DocID.Value, rs!DocNo.Value, rs!DocDate.Value, _
                       rs!DueKey.Value)
        openAmt(rs!PartyID & "|" & rs!DocType & "|" & rs!DocID) = CCur(rs!Amount)
        rs.MoveNext
    Loop
    rs.Close

    Set rs = db.OpenRecordset("SELECT PaymentID, InvoiceID, Amount FROM qryAgingAllocations WHERE " & where & _
                              " ORDER BY PaymentID, InvoiceID", dbOpenSnapshot)
    Do Until rs.EOF
        key = CStr(rs!PaymentID)
        If Not links.Exists(key) Then links.Add key, New Collection
        links(key).Add Array(rs!InvoiceID.Value, CCur(rs!Amount))
        rs.MoveNext
    Loop
    rs.Close

    ' 1. returns on their invoice, payments on their linked invoices
    Set rs = db.OpenRecordset("SELECT PartyID, CreditType, CreditID, Amount, TargetID FROM qryAgingCredits WHERE " & _
        where & " AND CreditDate < " & nextDay & " ORDER BY CreditDate, CreditType, CreditID", dbOpenSnapshot)
    Do Until rs.EOF
        rest = CCur(rs!Amount)
        If rs!CreditType = "RETURN" Then
            rest = ApplyTo(openAmt, rs!PartyID, Nz(rs!TargetID, 0), rest)
        ElseIf rs!CreditType = "PAYMENT" And links.Exists(CStr(rs!CreditID)) Then
            For Each l In links(CStr(rs!CreditID))
                part = l(1)
                If part > rest Then part = rest
                rest = rest - part + ApplyTo(openAmt, rs!PartyID, l(0), part)
            Next
        End If
        key = CStr(rs!PartyID)
        pool(key) = CCur(Nz(pool(key), 0)) + rest
        rs.MoveNext
    Loop
    rs.Close

    ' 2. the rest pays the oldest due documents first
    For Each d In docs
        key = d(0) & "|" & d(1) & "|" & d(2)
        part = CCur(Nz(pool(CStr(d(0))), 0))
        If part > openAmt(key) Then part = openAmt(key)
        If part > 0 Then
            openAmt(key) = openAmt(key) - part
            pool(CStr(d(0))) = pool(CStr(d(0))) - part
        End If
        If openAmt(key) <> 0 Then result.Add Array(d(0), d(1), d(2), d(3), d(4), d(5), openAmt(key))
    Next
    For Each k In pool.Keys
        If pool(k) <> 0 Then result.Add Array(CLng(k), "CREDIT", 0, "—’Ìœ œ«∆‰ €Ì— „” Œœ„", Null, Null, -pool(k))
    Next
    Set ComputeAging = result
End Function

Private Function ApplyTo(ByVal OpenAmt As Object, ByVal PartyID As Long, ByVal InvoiceID As Long, _
                         ByVal Amount As Currency) As Currency
    ' Pays the invoice up to what is open; returns the rest.
    Dim key As String, part As Currency
    key = PartyID & "|INVOICE|" & InvoiceID
    If OpenAmt.Exists(key) Then
        part = OpenAmt(key)
        If part > Amount Then part = Amount
        OpenAmt(key) = OpenAmt(key) - part
        Amount = Amount - part
    End If
    ApplyTo = Amount
End Function

Public Function AgeColumn(ByVal DaysLate As Long) As Long
    ' 0 not due, 1 = 1-30 days late, 2 = 31-60, 3 = 61-90, 4 = over 90.
    Select Case DaysLate
        Case Is <= 0: AgeColumn = 0
        Case Is <= 30: AgeColumn = 1
        Case Is <= 60: AgeColumn = 2
        Case Is <= 90: AgeColumn = 3
        Case Else: AgeColumn = 4
    End Select
End Function

Public Function FillAging(ByVal Kind As String, ByVal AsOf As Date, Optional ByVal PartyID As Long = 0) As Long
    ' tmpAging = the open documents of every party (or of one); returns their number.
    Dim db As DAO.Database, t As DAO.Recordset, item As Variant, names As Object, rs As DAO.Recordset
    Dim late As Long, cols As Variant, n As Long
    EnsureLocalTables
    Set db = CurrentDb
    db.Execute "DELETE FROM tmpAging", dbFailOnError
    Set names = CreateObject("Scripting.Dictionary")
    If Kind = "C" Then
        Set rs = db.OpenRecordset("SELECT CustomerID AS ID, CustomerName AS PartyName FROM Customers", dbOpenSnapshot)
    Else
        Set rs = db.OpenRecordset("SELECT SupplierID AS ID, SupplierName AS PartyName FROM Suppliers", dbOpenSnapshot)
    End If
    Do Until rs.EOF
        names(CStr(rs!ID)) = rs!PartyName.Value
        rs.MoveNext
    Loop
    rs.Close
    cols = Array("NotDue", "Days30", "Days60", "Days90", "Over90")
    Set t = db.OpenRecordset("tmpAging", dbOpenDynaset, dbAppendOnly)
    For Each item In ComputeAging(Kind, AsOf, PartyID)
        t.AddNew
        t!PartyKind = Kind
        t!PartyID = item(0)
        t!PartyName = Left$(Nz(names(CStr(item(0))), ""), 150)
        t!DocType = item(1)
        t!DocTypeName = IIf(item(1) = "INVOICE", "›« Ê—…", IIf(item(1) = "OPENING", "—’Ìœ «›  «ÕÌ", "—’Ìœ œ«∆‰"))
        t!DocID = item(2)
        t!DocNo = Left$(Nz(item(3), ""), 30)
        t!OpenAmount = item(6)
        For late = 0 To 4                       ' the table has no defaults: empty columns would be Null
            t.Fields(cols(late)).Value = 0
        Next
        t!Credit = 0
        If item(1) = "CREDIT" Then
            t!Credit = item(6)
        Else
            t!DocDate = DateValue(item(4))
            t!DueDate = item(5)
            late = DateValue(AsOf) - DateValue(item(5))
            t!DaysLate = IIf(late > 0, late, 0)
            t.Fields(cols(AgeColumn(late))).Value = item(6)
        End If
        t!AsOfDate = DateValue(AsOf)
        t.Update
        n = n + 1
    Next
    t.Close
    FillAging = n
End Function

Public Function MaxDaysLate(ByVal Kind As String, ByVal PartyID As Long, ByVal AsOf As Date) As Long
    Dim item As Variant, late As Long
    For Each item In ComputeAging(Kind, AsOf, PartyID)
        If item(1) <> "CREDIT" And item(6) > 0 Then
            late = DateValue(AsOf) - DateValue(item(5))
            If late > MaxDaysLate Then MaxDaysLate = late
        End If
    Next
End Function

Public Function DueDateFor(ByVal Kind As String, ByVal PartyID As Long, ByVal DocDate As Date) As Date
    ' The due date of a new credit invoice: the document date + the terms of the party.
    DueDateFor = DateValue(DocDate) + Nz(DbValue("SELECT PaymentTermsDays FROM " & IIf(Kind = "C", "Customers WHERE CustomerID", _
                                                 "Suppliers WHERE SupplierID") & " = " & PartyID), 0)
End Function

'------------------------------------------------------------------------------
' Links between payments and invoices
'------------------------------------------------------------------------------
Private Function KindParts(ByVal Kind As String, ByRef AllocTable As String, ByRef InvoiceCol As String, _
                           ByRef Prefix As String) As String
    ' The tables of a kind; the message when the user may not link them.
    If Kind = "C" Then
        AllocTable = "CustomerAllocations": InvoiceCol = "SalesInvoiceID": Prefix = "qryCustomer"
        If Not HasPermission("CUSTOMER_PAYMENTS") Then KindParts = "·«  „·ﬂ ’·«ÕÌ… ”‰œ«  «·ﬁ»÷."
    Else
        AllocTable = "SupplierAllocations": InvoiceCol = "PurchaseInvoiceID": Prefix = "qrySupplier"
        If Not HasPermission("SUPPLIER_PAYMENTS") Then KindParts = "·«  „·ﬂ ’·«ÕÌ… ”‰œ«  «·’—›."
    End If
End Function

Public Function AllocatePayment(ByVal Kind As String, ByVal PaymentID As Long, ByVal InvoiceID As Long, _
                                ByVal Amount As Currency) As String
    Dim allocTable As String, invoiceCol As String, prefix As String, payParty As Variant, invParty As Variant
    Dim payFree As Currency, invFree As Currency, existing As Variant
    AllocatePayment = KindParts(Kind, allocTable, invoiceCol, prefix)
    If Len(AllocatePayment) > 0 Then Exit Function
    If Not CanScreenAction("frmAllocation", "ADD", True) Then
        AllocatePayment = "·«  „·ﬂ ’·«ÕÌ… «·≈÷«›… ›Ì ‘«‘… —»ÿ «·”œ«œ »«·›Ê« Ì—."
        Exit Function
    End If
    payParty = DbValue("SELECT PartyID FROM " & prefix & "PaymentFree WHERE PaymentID = " & PaymentID)
    invParty = DbValue("SELECT PartyID FROM " & prefix & "InvoiceFree WHERE InvoiceID = " & InvoiceID)
    payFree = Nz(DbValue("SELECT Free FROM " & prefix & "PaymentFree WHERE PaymentID = " & PaymentID), 0)
    invFree = Nz(DbValue("SELECT Free FROM " & prefix & "InvoiceFree WHERE InvoiceID = " & InvoiceID), 0)
    If IsNull(payParty) Then
        AllocatePayment = "«Œ — «·”‰œ."
    ElseIf IsNull(invParty) Then
        AllocatePayment = "«Œ — ›« Ê—… ¬Ã·… ⁄·ÌÂ« „ »ﬁÚ."
    ElseIf payParty <> invParty Then
        AllocatePayment = "«·”‰œ Ê«·›« Ê—… ·ÿ—›Ì‰ „Œ ·›Ì‰."
    ElseIf Amount <= 0 Then
        AllocatePayment = "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—."
    ElseIf Amount > payFree Then
        AllocatePayment = "«·»«ﬁÌ €Ì— «·„—»Êÿ „‰ «·”‰œ " & Format$(payFree, "#,##0.00") & " ›ﬁÿ."
    ElseIf Amount > invFree Then
        AllocatePayment = "«·»«ﬁÌ „‰ «·›« Ê—… »⁄œ „« —ı»ÿ »Â« Ê„— Ã⁄« Â« " & Format$(invFree, "#,##0.00") & " ›ﬁÿ."
    End If
    If Len(AllocatePayment) > 0 Then Exit Function
    existing = DbValue("SELECT AllocationID FROM " & allocTable & " WHERE PaymentID = " & PaymentID & " AND " & _
                       invoiceCol & " = " & InvoiceID)
    If IsNull(existing) Then
        CurrentDb.Execute "INSERT INTO " & allocTable & " (PaymentID, " & invoiceCol & ", Amount, EmployeeID) VALUES (" & _
                          PaymentID & ", " & InvoiceID & ", " & Str$(Amount) & ", " & CurrentUserID() & ")", dbFailOnError
    Else
        CurrentDb.Execute "UPDATE " & allocTable & " SET Amount = Amount + " & Str$(Amount) & " WHERE AllocationID = " & _
                          existing, dbFailOnError
    End If
    LogAction "ALLOCATE", allocTable, PaymentID & ">" & InvoiceID, Format$(Amount, "0.00")
End Function

Public Function RemoveAllocation(ByVal Kind As String, ByVal AllocationID As Long) As String
    Dim allocTable As String, invoiceCol As String, prefix As String
    RemoveAllocation = KindParts(Kind, allocTable, invoiceCol, prefix)
    If Len(RemoveAllocation) > 0 Then Exit Function
    If Not CanScreenAction("frmAllocation", "DELETE", True) Then
        RemoveAllocation = "·«  „·ﬂ ’·«ÕÌ… «·Õ–› ›Ì ‘«‘… —»ÿ «·”œ«œ »«·›Ê« Ì—."
        Exit Function
    End If
    CurrentDb.Execute "DELETE FROM " & allocTable & " WHERE AllocationID = " & AllocationID, dbFailOnError
    LogAction "UNALLOCATE", allocTable, CStr(AllocationID)
End Function

Public Function AutoAllocate(ByVal Kind As String, ByVal PartyID As Long, ByRef Count As Long) As String
    ' Every free payment amount (oldest payment first) to the free invoices (oldest due first).
    Dim db As DAO.Database, pays As DAO.Recordset, invs As DAO.Recordset, prefix As String
    Dim allocTable As String, invoiceCol As String, free As Currency, part As Currency, invLeft As Object, key As String
    Count = 0
    AutoAllocate = KindParts(Kind, allocTable, invoiceCol, prefix)
    If Len(AutoAllocate) > 0 Then Exit Function
    Set db = CurrentDb
    Set invLeft = CreateObject("Scripting.Dictionary")
    Set pays = db.OpenRecordset("SELECT PaymentID, Free FROM " & prefix & "PaymentFree WHERE PartyID = " & PartyID & _
                                " AND Free > 0 ORDER BY PaymentDate, PaymentID", dbOpenSnapshot)
    Do Until pays.EOF
        free = pays!Free
        Set invs = db.OpenRecordset("SELECT InvoiceID, Free FROM " & prefix & "InvoiceFree WHERE PartyID = " & PartyID & _
            " AND Free > 0 ORDER BY IIf(DueDate Is Null, InvoiceDate, DueDate), InvoiceDate, InvoiceID", dbOpenSnapshot)
        Do Until invs.EOF Or free <= 0
            key = CStr(invs!InvoiceID)
            If Not invLeft.Exists(key) Then invLeft(key) = CCur(invs!Free)
            part = invLeft(key)
            If part > free Then part = free
            If part > 0 Then
                AutoAllocate = AllocatePayment(Kind, pays!PaymentID, invs!InvoiceID, part)
                If Len(AutoAllocate) > 0 Then Exit Do
                invLeft(key) = invLeft(key) - part
                free = free - part
                Count = Count + 1
            End If
            invs.MoveNext
        Loop
        invs.Close
        If Len(AutoAllocate) > 0 Then Exit Do
        pays.MoveNext
    Loop
    pays.Close
End Function

'------------------------------------------------------------------------------
' Screen frmAging (OpenArgs "C|12" / "S|3": the kind and the party to show)
'------------------------------------------------------------------------------
Private Function Money(ByVal Value As Variant) As String
    Money = Format$(Nz(Value, 0), "#,##0.00")
End Function

Public Sub AgingLoad(ByVal frm As Access.Form)
    Dim args As String
    Calendar = vbCalGreg
    args = Nz(frm.OpenArgs, "")
    frm!cboKind.Value = IIf(Left$(args, 1) = "S", "S", "C")
    frm!txtAsOf.Value = Date
    AgingRefresh frm
    If InStr(args, "|") > 0 Then
        frm!lstParties.Value = CLng(Mid$(args, InStr(args, "|") + 1))
        AgingPartyChanged frm
    End If
End Sub

Public Sub AgingRefresh(ByVal frm As Access.Form)
    Dim n As Long, total As Variant
    If Not IsDate(frm!txtAsOf.Value) Then
        ShowWarning "√œŒ·  «—ÌŒ √⁄„«— «·œÌÊ‰."
        Exit Sub
    End If
    n = FillAging(frm!cboKind.Value, DateValue(frm!txtAsOf.Value))
    frm!lstParties.RowSource = Tr("SELECT PartyID, PartyName AS [" & IIf(frm!cboKind.Value = "C", "«·⁄„Ì·", "«·„Ê—œ") & "], " & _
        "Format(Sum(OpenAmount), '#,##0.00') AS [«·—’Ìœ], Format(Sum(NotDue), '#,##0.00') AS [€Ì— „” Õﬁ], " & _
        "Format(Sum(Days30), '#,##0.00') AS [1-30], Format(Sum(Days60), '#,##0.00') AS [31-60], " & _
        "Format(Sum(Days90), '#,##0.00') AS [61-90], Format(Sum(Over90), '#,##0.00') AS [+90], " & _
        "Format(Sum(a.Credit), '#,##0.00') AS [—’Ìœ œ«∆‰], Max(DaysLate) AS [√ﬁ’Ï  √ŒÌ—] FROM tmpAging AS a " & _
        "GROUP BY PartyID, PartyName ORDER BY Sum(Over90) DESC, Sum(Days90) DESC, Sum(OpenAmount) DESC")
    frm!lstDocs.RowSource = Tr("")
    total = DbValue("SELECT Sum(OpenAmount) FROM tmpAging")
    frm!lblTotals.Caption = Tr("«·≈Ã„«·Ì " & Money(total) & "   €Ì— „” Õﬁ " & Money(DbValue("SELECT Sum(NotDue) FROM tmpAging")) & _
        "   1-30 " & Money(DbValue("SELECT Sum(Days30) FROM tmpAging")) & "   31-60 " & _
        Money(DbValue("SELECT Sum(Days60) FROM tmpAging")) & "   61-90 " & Money(DbValue("SELECT Sum(Days90) FROM tmpAging")) & _
        "   √ﬂÀ— „‰ 90 " & Money(DbValue("SELECT Sum(Over90) FROM tmpAging")))
    frm!lblInfo.Caption = Tr(n & " „” ‰œ „› ÊÕ. «Œ — " & IIf(frm!cboKind.Value = "C", "⁄„Ì·«", "„Ê—œ«") & " ·⁄—÷ ›Ê« Ì—Â.")
End Sub

Public Sub AgingPartyChanged(ByVal frm As Access.Form)
    If IsNull(frm!lstParties.Value) Then
        frm!lstDocs.RowSource = Tr("")
        Exit Sub
    End If
    frm!lstDocs.RowSource = Tr("SELECT DocID, DocTypeName AS [«·„” ‰œ], DocNo AS [«·—ﬁ„], Format(DocDate, 'yyyy/mm/dd') AS " & _
        "[«· «—ÌŒ], Format(DueDate, 'yyyy/mm/dd') AS [«·«” Õﬁ«ﬁ], DaysLate AS [√Ì«„ «· √ŒÌ—], " & _
        "Format(OpenAmount, '#,##0.00') AS [«·„ »ﬁÌ] FROM tmpAging WHERE PartyID = " & frm!lstParties.Value & _
        " ORDER BY DueDate, LineNo")
End Sub

Private Function PickedParty(ByVal frm As Access.Form) As Long
    PickedParty = Nz(frm!lstParties.Value, 0)
    If PickedParty = 0 Then ShowWarning "«Œ — " & IIf(frm!cboKind.Value = "C", "«·⁄„Ì·", "«·„Ê—œ") & " „‰ «·ﬁ«∆„…."
End Function

Public Sub AgingOpenAllocation(ByVal frm As Access.Form)
    Dim id As Long
    id = PickedParty(frm)
    If id > 0 Then OpenScreen "frmAllocation", 0, frm!cboKind.Value & "|" & id
End Sub

Public Sub AgingStatement(ByVal frm As Access.Form)
    Dim id As Long
    id = PickedParty(frm)
    If id > 0 Then PrintPartyStatement frm!cboKind.Value, id
End Sub

Public Sub PrintAging(ByVal frm As Access.Form)
    If Not IsDate(frm!txtAsOf.Value) Then Exit Sub
    FillAging frm!cboKind.Value, DateValue(frm!txtAsOf.Value)
    LogAction "REPORT", "AGING", frm!cboKind.Value
    OpenReportOrQuery "rptAging", "tmpAging", "", IIf(frm!cboKind.Value = "C", "«·⁄„·«¡", "«·„Ê—œÊ‰") & _
                      "    ›Ì " & GDate(frm!txtAsOf.Value)
End Sub

'------------------------------------------------------------------------------
' Screen frmAllocation (OpenArgs "C|12" / "S|3")
'------------------------------------------------------------------------------
Public Sub AllocationLoad(ByVal frm As Access.Form)
    Dim args As String
    Calendar = vbCalGreg
    args = Nz(frm.OpenArgs, "")
    frm!cboKind.Value = IIf(Left$(args, 1) = "S", "S", "C")
    AllocationKindChanged frm
    If InStr(args, "|") > 0 Then
        frm!cboParty.Value = CLng(Mid$(args, InStr(args, "|") + 1))
        AllocationRefresh frm
    End If
End Sub

Public Sub AllocationKindChanged(ByVal frm As Access.Form)
    If frm!cboKind.Value = "C" Then
        frm!cboParty.RowSource = Tr("SELECT CustomerID, CustomerName FROM Customers WHERE AllowCredit = True Or " & _
                                 "CurrentBalance <> 0 ORDER BY CustomerName")
    Else
        frm!cboParty.RowSource = Tr("SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName")
    End If
    frm!cboParty.Value = Null
    AllocationRefresh frm
End Sub

Public Sub AllocationRefresh(ByVal frm As Access.Form)
    Dim prefix As String, party As Long, allocTable As String, invoiceCol As String, payTable As String
    Dim invTable As String
    party = Nz(frm!cboParty.Value, 0)
    If frm!cboKind.Value = "C" Then
        prefix = "qryCustomer": allocTable = "CustomerAllocations": invoiceCol = "SalesInvoiceID"
        payTable = "CustomerPayments": invTable = "SalesInvoices"
    Else
        prefix = "qrySupplier": allocTable = "SupplierAllocations": invoiceCol = "PurchaseInvoiceID"
        payTable = "SupplierPayments": invTable = "PurchaseInvoices"
    End If
    frm!lstPayments.RowSource = Tr("SELECT PaymentID, PaymentNumber AS [«·”‰œ], Format(PaymentDate, 'yyyy/mm/dd') AS [«· «—ÌŒ], " & _
        "Format(f.Amount, '#,##0.00') AS [„»·€ «·”‰œ], Format(Allocated, '#,##0.00') AS [„—»Êÿ], Format(Free, '#,##0.00') AS " & _
        "[€Ì— „—»Êÿ] FROM " & prefix & "PaymentFree AS f WHERE PartyID = " & party & " ORDER BY PaymentDate DESC")
    frm!lstInvoices.RowSource = Tr("SELECT InvoiceID, InvoiceNumber AS [«·›« Ê—…], Format(InvoiceDate, 'yyyy/mm/dd') AS " & _
        "[«· «—ÌŒ], Format(DueDate, 'yyyy/mm/dd') AS [«·«” Õﬁ«ﬁ], Format(RemainingAmount, '#,##0.00') AS [«·„ »ﬁÌ], " & _
        "Format(Allocated + Returned, '#,##0.00') AS [„—»Êÿ Ê„— Ã⁄], Format(Free, '#,##0.00') AS [Ì„ﬂ‰ —»ÿÂ] FROM " & _
        prefix & "InvoiceFree WHERE PartyID = " & party & " AND Free > 0 ORDER BY IIf(DueDate Is Null, InvoiceDate, DueDate)")
    frm!lstAllocations.RowSource = Tr("SELECT a.AllocationID, p.PaymentNumber AS [«·”‰œ], h.InvoiceNumber AS [«·›« Ê—…], " & _
        "Format(a.Amount, '#,##0.00') AS [«·„»·€ «·„—»Êÿ], Format(a.CreatedAt, 'yyyy/mm/dd') AS [›Ì] FROM (" & allocTable & _
        " AS a INNER JOIN " & payTable & " AS p ON a.PaymentID = p.PaymentID) INNER JOIN " & invTable & " AS h ON a." & _
        invoiceCol & " = h." & invoiceCol & " WHERE p." & IIf(frm!cboKind.Value = "C", "CustomerID", "SupplierID") & _
        " = " & party & " ORDER BY a.AllocationID DESC")
    frm!txtAmount.Value = Null
End Sub

Public Sub AllocationPicked(ByVal frm As Access.Form)
    ' The amount proposed: the smaller of what is free on the payment and on the invoice.
    Dim payFree As Currency, invFree As Currency, prefix As String
    If IsNull(frm!lstPayments.Value) Or IsNull(frm!lstInvoices.Value) Then Exit Sub
    prefix = IIf(frm!cboKind.Value = "C", "qryCustomer", "qrySupplier")
    payFree = Nz(DbValue("SELECT Free FROM " & prefix & "PaymentFree WHERE PaymentID = " & frm!lstPayments.Value), 0)
    invFree = Nz(DbValue("SELECT Free FROM " & prefix & "InvoiceFree WHERE InvoiceID = " & frm!lstInvoices.Value), 0)
    frm!txtAmount.Value = IIf(payFree < invFree, payFree, invFree)
End Sub

Public Sub DoAllocate(ByVal frm As Access.Form)
    Dim msg As String
    If IsNull(frm!lstPayments.Value) Or IsNull(frm!lstInvoices.Value) Then
        ShowWarning "«Œ — «·”‰œ Ê«·›« Ê—…."
        Exit Sub
    End If
    msg = AllocatePayment(frm!cboKind.Value, frm!lstPayments.Value, frm!lstInvoices.Value, CCur(Nz(frm!txtAmount.Value, 0)))
    If Len(msg) > 0 Then
        ShowWarning msg
        Exit Sub
    End If
    AllocationRefresh frm
End Sub

Public Sub DoRemoveAllocation(ByVal frm As Access.Form)
    Dim msg As String
    If IsNull(frm!lstAllocations.Value) Then
        ShowWarning "«Œ — «·—»ÿ „‰ «·ﬁ«∆„…."
        Exit Sub
    End If
    If Not AskYesNo("≈·€«¡ —»ÿ Â–« «·„»·€ »«·›« Ê—…ø") Then Exit Sub
    msg = RemoveAllocation(frm!cboKind.Value, frm!lstAllocations.Value)
    If Len(msg) > 0 Then ShowWarning msg
    AllocationRefresh frm
End Sub

Public Sub DoAutoAllocate(ByVal frm As Access.Form)
    Dim msg As String, n As Long
    If IsNull(frm!cboParty.Value) Then
        ShowWarning "«Œ — " & IIf(frm!cboKind.Value = "C", "«·⁄„Ì·", "«·„Ê—œ") & "."
        Exit Sub
    End If
    If Not AskYesNo("—»ÿ ﬂ· «·„»«·€ €Ì— «·„—»Êÿ… »√ﬁœ„ «·›Ê« Ì— «” Õﬁ«ﬁ«ø") Then Exit Sub
    msg = AutoAllocate(frm!cboKind.Value, frm!cboParty.Value, n)
    If Len(msg) > 0 Then
        ShowWarning msg
    Else
        ShowInfo IIf(n = 0, "·«  ÊÃœ „»«·€ €Ì— „—»Êÿ… √Ê ›Ê« Ì— „› ÊÕ….", " „ «·—»ÿ (" & n & ").")
    End If
    AllocationRefresh frm
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests): inside a transaction that is rolled back
'------------------------------------------------------------------------------
Private Function AgingTotal(ByVal Kind As String, Optional ByVal PartyID As Long = 0) As Currency
    Dim item As Variant
    For Each item In ComputeAging(Kind, Date, PartyID)
        AgingTotal = AgingTotal + item(6)
    Next
End Function

Private Function InvoiceOpen(ByVal PartyID As Long, ByVal InvoiceID As Long) As Currency
    Dim item As Variant
    For Each item In ComputeAging("C", Date, PartyID)
        If item(1) = "INVOICE" And item(2) = InvoiceID Then InvoiceOpen = item(6)
    Next
End Function

Private Sub CheckAging(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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

Public Function TestAging() As Boolean
    Dim passed As Long, failed As Long, report As String, msg As String, ws As DAO.Workspace, inTrans As Boolean
    Dim inv As Variant, party As Long, payID As Long, before As Currency, n As Long
    Calendar = vbCalGreg
    EnsureTestUser
    g_SilentMode = True
    Debug.Print "=== TestAging  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    On Error GoTo EH
    CheckAging AgeColumn(0) = 0 And AgeColumn(1) = 1 And AgeColumn(30) = 1 And AgeColumn(31) = 2 And _
               AgeColumn(61) = 3 And AgeColumn(91) = 4, "√⁄„œ… «·√⁄„«—: €Ì— „” Õﬁ° 1-30° 31-60° 61-90° +90", _
               passed, failed, report
    CheckAging AgingTotal("C") = Nz(DbValue("SELECT Sum(Balance) FROM CustomerBalanceQuery"), 0), _
               "„Ã„Ê⁄ √⁄„«— œÌÊ‰ «·⁄„·«¡ = √—’œ… «·⁄„·«¡", passed, failed, report
    CheckAging AgingTotal("S") = Nz(DbValue("SELECT Sum(Balance) FROM SupplierBalanceQuery"), 0), _
               "„Ã„Ê⁄ √⁄„«— œÌÊ‰ «·„Ê—œÌ‰ = √—’œ… «·„Ê—œÌ‰", passed, failed, report
    n = FillAging("C", Date)
    CheckAging Nz(DbValue("SELECT Sum(OpenAmount) FROM tmpAging"), 0) = Nz(DbValue("SELECT Sum(NotDue + Days30 + " & _
               "Days60 + Days90 + Over90 + Credit) FROM tmpAging"), 0), "ﬂ· „»·€ ›Ì ⁄„Êœ ⁄„— Ê«Õœ (" & n & ")", _
               passed, failed, report

    ' a payment linked to an invoice pays that invoice
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    inv = DbValue("SELECT TOP 1 InvoiceID FROM qryCustomerInvoiceFree WHERE Free >= 1 ORDER BY InvoiceDate DESC")
    If IsNull(inv) Then
        Debug.Print "[--] —»ÿ «·”œ«œ: ·«  ÊÃœ ›« Ê—… ¬Ã·… „› ÊÕ…"
    Else
        party = DbValue("SELECT PartyID FROM qryCustomerInvoiceFree WHERE InvoiceID = " & inv)
        before = AgingTotal("C", party)
        msg = PostCustomerPayment(party, 1, 1, "TEST", payID)
        msg = msg & AllocatePayment("C", payID, inv, 1)
        CheckAging Len(msg) = 0 And Nz(DbValue("SELECT Free FROM qryCustomerPaymentFree WHERE PaymentID = " & payID), -1) = 0, _
                   "—»ÿ ”‰œ ﬁ»÷ »›« Ê—… " & msg, passed, failed, report
        CheckAging AgingTotal("C", party) = before - 1, "«·”‰œ Ì‰ﬁ’ √⁄„«— «·⁄„Ì· »„»·€Â", passed, failed, report
        CheckAging Len(AllocatePayment("C", payID, inv, 1)) > 0, "·« Ìı—»ÿ √ﬂÀ— „‰ „»·€ «·”‰œ", passed, failed, report
        msg = RemoveAllocation("C", Nz(DbValue("SELECT AllocationID FROM CustomerAllocations WHERE PaymentID = " & payID), 0))
        msg = msg & AutoAllocate("C", party, n)
        CheckAging Len(msg) = 0 And n >= 1 And Nz(DbValue("SELECT Free FROM qryCustomerPaymentFree WHERE PaymentID = " & _
                   payID), -1) = 0, "«·—»ÿ «· ·ﬁ«∆Ì Ì—»ÿ «·”‰œ »√ﬁœ„ ›« Ê—…", passed, failed, report
    End If
    ws.Rollback
    inTrans = False
    GoTo Done
EH:
    CheckAging False, "Œÿ√: " & Err.Description, passed, failed, report
    If inTrans Then ws.Rollback
Done:
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  √⁄„«— «·œÌÊ‰ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestAging"
        TestAging = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, "TestAging"
    End If
End Function
