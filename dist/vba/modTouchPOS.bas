Attribute VB_Name = "modTouchPOS"
'==============================================================================
' modTouchPOS  -  Retail Store Management System: restaurant touch screen
'
'   frmTouchPOS    order type (dine-in / takeaway / delivery), category tiles,
'                  product tiles with pictures, large cart rows with + / -,
'                  cash (frmTouchPay number pad) or card payment.
'   It reuses the shop point of sale: the cart table tmpPOSLines, AddLine,
'   RecalcPOS, SavePOS and ResetSaleHeader of modPOS (the same control names),
'   so prices, VAT, discounts, posting and the ZATCA QR code are identical.
'   Which sales screen opens is chosen in Settings.POSMode (SalesScreenName).
'==============================================================================
Option Compare Database
Option Explicit

Public Const CAT_TILES As Long = 7
Public Const PRODUCT_TILES As Long = 16
Private Const DEFAULT_ORDER_TYPE As String = "DINE_IN"
Private Const CARD_METHOD_ID As Long = 2              ' PaymentMethods: „œÏ / »ÿ«ﬁ…

Private m_catPage As Long
Private m_prodPage As Long
Private m_catID As Long

'------------------------------------------------------------------------------
' Pure helpers
'------------------------------------------------------------------------------
Public Function OrderTypeText(ByVal OrderType As Variant, ByVal TableNo As Variant) As String
    Select Case Nz(OrderType, "")
        Case "DINE_IN"
            OrderTypeText = "ÿ·» œ«Œ·Ì" & IIf(Len(Nz(TableNo, "")) > 0, " - ÿ«Ê·… " & Nz(TableNo, ""), "")
        Case "TAKEAWAY"
            OrderTypeText = "ÿ·» ”›—Ì"
        Case "DELIVERY"
            OrderTypeText = "ÿ·»  Ê’Ì·"
        Case Else
            OrderTypeText = ""
    End Select
End Function

Public Function TileColorValue(ByVal ColorCode As Variant) As Long
    Select Case Nz(ColorCode, "")
        Case "GREEN": TileColorValue = RGB(67, 160, 71)
        Case "ORANGE": TileColorValue = RGB(251, 140, 0)
        Case "PURPLE": TileColorValue = RGB(142, 36, 170)
        Case "RED": TileColorValue = RGB(229, 57, 53)
        Case "INDIGO": TileColorValue = RGB(57, 73, 171)
        Case "TEAL": TileColorValue = RGB(0, 137, 123)
        Case "PINK": TileColorValue = RGB(216, 27, 96)
        Case "BROWN": TileColorValue = RGB(121, 85, 72)
        Case "GREY": TileColorValue = RGB(96, 125, 139)
        Case Else: TileColorValue = RGB(30, 136, 229)
    End Select
End Function

Public Function PageCount(ByVal Items As Long, ByVal PerPage As Long) As Long
    PageCount = (Items + PerPage - 1) \ PerPage
    If PageCount < 1 Then PageCount = 1
End Function

Public Function IsAbsolutePath(ByVal Path As String) As Boolean
    IsAbsolutePath = (Mid$(Path, 2, 1) = ":") Or (Left$(Path, 2) = "\\")
End Function

Public Function ImageCandidate(ByVal Path As String, ByVal Folder As String) As String
    ' Full path of a picture: absolute paths as they are, names relative to the images folder.
    Path = Trim$(Path)
    If Len(Path) = 0 Then Exit Function
    If IsAbsolutePath(Path) Or Len(Folder) = 0 Then
        ImageCandidate = Path
    ElseIf Right$(Folder, 1) = "\" Then
        ImageCandidate = Folder & Path
    Else
        ImageCandidate = Folder & "\" & Path
    End If
End Function

Public Function ImagesFolder() As String
    ' Settings.ImagesFolder, otherwise the Images folder next to the back-end (shared by every PC).
    Dim be As String
    ImagesFolder = Trim$(Nz(SettingValue("ImagesFolder"), ""))
    If Len(ImagesFolder) > 0 Then Exit Function
    be = BackendFilePath()
    If InStrRev(be, "\") > 0 Then ImagesFolder = Left$(be, InStrRev(be, "\")) & "Images"
End Function

Public Function ImageFile(ByVal Path As Variant) As String
    ' "" when there is no picture file.
    Dim p As String
    On Error Resume Next
    p = ImageCandidate(Nz(Path, ""), ImagesFolder())
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then ImageFile = p
    End If
End Function

Public Function SalesScreenName() As String
    ' The sales button opens the screen chosen in Settings.POSMode.
    SalesScreenName = "frmPOS"
    Select Case Nz(SettingValue("POSMode"), "RETAIL")
        Case "RESTAURANT", "CAFE"
            If FormExists("frmTouchPOS") Then SalesScreenName = "frmTouchPOS"
    End Select
End Function

'------------------------------------------------------------------------------
' Screen
'------------------------------------------------------------------------------
Public Sub TouchLoad(ByVal frm As Access.Form)
    Calendar = vbCalGreg
    EnsureLocalTables
    If DCount("*", "tmpPOSLines") > 0 And Not g_SilentMode Then
        If Not AskYesNo("ÌÊÃœ ÿ·» €Ì— „ﬂ „· „‰ Ã·”… ”«»ﬁ…. Â·  —Ìœ «” ﬂ„«·Âø") Then
            CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
        End If
    End If
    ResetSaleHeader frm                                   ' modPOS: also resets the order (ResetTouchOrder)
    m_catPage = 0
    m_prodPage = 0
    m_catID = Nz(DbValue("SELECT TOP 1 c.CategoryID FROM Categories AS c WHERE c.IsActive = True AND " & _
                         "EXISTS (SELECT 1 FROM Products AS p WHERE p.CategoryID = c.CategoryID AND p.IsActive = True) " & _
                         "ORDER BY c.SortOrder, c.CategoryName"), 0)
    LoadCategories frm
    LoadProducts frm
    RecalcPOS frm
End Sub

Public Sub ResetTouchOrder(ByVal frm As Access.Form)
    frm!txtOrderType.Value = DEFAULT_ORDER_TYPE
    frm!cboTable.Value = Null
    frm!txtDeliveryPhone.Value = Null
    frm!txtDeliveryAddress.Value = Null
    ShowOrderType frm
End Sub

Public Sub SetOrderType(ByVal frm As Access.Form, ByVal OrderType As String)
    frm!txtOrderType.Value = OrderType
    If OrderType <> "DELIVERY" Then
        frm!cboCustomer.Value = Nz(SettingValue("DefaultCustomerID"), 1)
        frm!txtDeliveryPhone.Value = Null
        frm!txtDeliveryAddress.Value = Null
        CustomerChanged frm
    End If
    If OrderType <> "DINE_IN" Then frm!cboTable.Value = Null
    ShowOrderType frm
End Sub

Private Sub ShowOrderType(ByVal frm As Access.Form)
    Dim t As String, keys As Variant, btns As Variant, i As Long, selected As Boolean
    On Error Resume Next                                   ' a focused control cannot be hidden
    t = Nz(frm!txtOrderType.Value, DEFAULT_ORDER_TYPE)
    keys = Array("DINE_IN", "TAKEAWAY", "DELIVERY")
    btns = Array("btnTypeDineIn", "btnTypeTakeaway", "btnTypeDelivery")
    frm!btnPayCash.SetFocus
    For i = 0 To 2
        selected = (keys(i) = t)
        frm.Controls(btns(i)).BackColor = IIf(selected, CLR_PRIMARY, CLR_SECONDARY)
        frm.Controls(btns(i)).ForeColor = IIf(selected, CLR_SURFACE, CLR_TEXT)
        frm.Controls(btns(i)).FontBold = selected
    Next
    frm!lblTable.Visible = (t = "DINE_IN")
    frm!cboTable.Visible = (t = "DINE_IN")
    frm!lblTakeaway.Visible = (t = "TAKEAWAY")
    frm!lblCustomer.Visible = (t = "DELIVERY")
    frm!cboCustomer.Visible = (t = "DELIVERY")
    frm!lblDeliveryPhone.Visible = (t = "DELIVERY")
    frm!txtDeliveryPhone.Visible = (t = "DELIVERY")
    frm!lblDeliveryAddress.Visible = (t = "DELIVERY")
    frm!txtDeliveryAddress.Visible = (t = "DELIVERY")
    frm!lblOrderTitle.Caption = OrderTypeText(t, frm!cboTable.Value)
End Sub

Public Sub TableChanged(ByVal frm As Access.Form)
    frm!lblOrderTitle.Caption = OrderTypeText(frm!txtOrderType.Value, frm!cboTable.Value)
End Sub

Public Sub DeliveryCustomerChanged(ByVal frm As Access.Form)
    ' A registered customer brings the mobile number and the address.
    Dim rs As DAO.Recordset
    CustomerChanged frm
    If Nz(frm!cboCustomer.Value, 1) = Nz(SettingValue("DefaultCustomerID"), 1) Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT Mobile, Phone, Address, District, City FROM Customers " & _
                                     "WHERE CustomerID = " & frm!cboCustomer.Value, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!txtDeliveryPhone.Value = Nz(rs!Mobile, rs!Phone)
        frm!txtDeliveryAddress.Value = Nz(rs!Address, Trim$(Nz(rs!District, "") & " " & Nz(rs!City, "")))
    End If
    rs.Close
End Sub

Public Function TouchOrderProblem(ByVal frm As Access.Form) As String
    ' Checked before saving (called by modPOS.SavePOS).
    If Nz(frm!txtOrderType.Value, "") = "DELIVERY" Then
        If Len(Trim$(Nz(frm!txtDeliveryPhone.Value, ""))) = 0 Then
            TouchOrderProblem = "ÿ·» «· Ê’Ì· ÌÕ «Ã ÃÊ«· «·⁄„Ì·."
        ElseIf Len(Trim$(Nz(frm!txtDeliveryAddress.Value, ""))) = 0 Then
            TouchOrderProblem = "ÿ·» «· Ê’Ì· ÌÕ «Ã «·⁄‰Ê«‰."
        End If
    End If
End Function

'------------------------------------------------------------------------------
' Category and product tiles
'------------------------------------------------------------------------------
Private Sub ShowPicture(ByVal ctl As Access.Control, ByVal Path As Variant)
    Dim f As String
    On Error Resume Next
    f = ImageFile(Path)
    If Len(f) = 0 Then
        ctl.Visible = False
    Else
        ctl.Picture = f
        ctl.Visible = (Err.Number = 0)
    End If
End Sub

Public Sub LoadCategories(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, n As Long, i As Long, k As String, sel As Boolean
    Set rs = CurrentDb.OpenRecordset("SELECT c.CategoryID, c.CategoryName, c.ImagePath, c.TileColor " & _
        "FROM Categories AS c WHERE c.IsActive = True AND EXISTS (SELECT 1 FROM Products AS p " & _
        "WHERE p.CategoryID = c.CategoryID AND p.IsActive = True) ORDER BY c.SortOrder, c.CategoryName", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        n = rs.RecordCount
        rs.MoveFirst
    End If
    If m_catPage > PageCount(n, CAT_TILES) - 1 Then m_catPage = PageCount(n, CAT_TILES) - 1
    If n > 0 Then rs.Move m_catPage * CAT_TILES
    For i = 1 To CAT_TILES
        k = CStr(i)
        If rs.EOF Then
            frm.Controls("boxCat" & k).Visible = False
            frm.Controls("imgCat" & k).Visible = False
            frm.Controls("lblCat" & k).Visible = False
            frm.Controls("btnCat" & k).Visible = False
        Else
            sel = (rs!CategoryID = m_catID)
            With frm.Controls("boxCat" & k)
                .Visible = True
                .BackColor = TileColorValue(rs!TileColor)
                .BorderStyle = IIf(sel, 1, 0)
                .BorderWidth = 3
                .BorderColor = CLR_TEXT
            End With
            ShowPicture frm.Controls("imgCat" & k), rs!ImagePath
            frm.Controls("lblCat" & k).Caption = rs!CategoryName
            frm.Controls("lblCat" & k).FontBold = sel
            frm.Controls("lblCat" & k).Visible = True
            frm.Controls("btnCat" & k).Tag = CStr(rs!CategoryID)
            frm.Controls("btnCat" & k).Visible = True
            rs.MoveNext
        End If
    Next
    rs.Close
    frm!btnCatUp.Enabled = (m_catPage > 0)
    frm!btnCatDown.Enabled = (m_catPage < PageCount(n, CAT_TILES) - 1)
End Sub

Public Sub LoadProducts(ByVal frm As Access.Form)
    Dim rs As DAO.Recordset, n As Long, i As Long, k As String, color As Long
    color = TileColorValue(DbValue("SELECT TileColor FROM Categories WHERE CategoryID = " & m_catID))
    frm!lblCategoryTitle.Caption = Nz(DbValue("SELECT CategoryName FROM Categories WHERE CategoryID = " & m_catID), " ")
    Set rs = CurrentDb.OpenRecordset("SELECT ProductID, ProductName, SellingPrice, ImagePath FROM Products " & _
        "WHERE IsActive = True AND CategoryID = " & m_catID & " ORDER BY ProductName", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        n = rs.RecordCount
        rs.MoveFirst
    End If
    If m_prodPage > PageCount(n, PRODUCT_TILES) - 1 Then m_prodPage = PageCount(n, PRODUCT_TILES) - 1
    If n > 0 Then rs.Move m_prodPage * PRODUCT_TILES
    For i = 1 To PRODUCT_TILES
        k = CStr(i)
        If rs.EOF Then
            frm.Controls("boxProd" & k).Visible = False
            frm.Controls("boxProdStrip" & k).Visible = False
            frm.Controls("imgProd" & k).Visible = False
            frm.Controls("lblProd" & k).Visible = False
            frm.Controls("lblPrice" & k).Visible = False
            frm.Controls("btnProd" & k).Visible = False
        Else
            frm.Controls("boxProd" & k).Visible = True
            frm.Controls("boxProdStrip" & k).BackColor = color
            frm.Controls("boxProdStrip" & k).Visible = True
            ShowPicture frm.Controls("imgProd" & k), rs!ImagePath
            frm.Controls("lblProd" & k).Caption = rs!ProductName
            frm.Controls("lblProd" & k).Visible = True
            frm.Controls("lblPrice" & k).Caption = Format$(Nz(rs!SellingPrice, 0), "#,##0.00") & " —.”"
            frm.Controls("lblPrice" & k).Visible = True
            frm.Controls("btnProd" & k).Tag = CStr(rs!ProductID)
            frm.Controls("btnProd" & k).Visible = True
            rs.MoveNext
        End If
    Next
    rs.Close
    frm!lblProdPage.Caption = "’›Õ… " & (m_prodPage + 1) & " „‰ " & PageCount(n, PRODUCT_TILES) & _
                              "   (" & n & " ’‰›)"
    frm!btnProdPrev.Enabled = (m_prodPage > 0)
    frm!btnProdNext.Enabled = (m_prodPage < PageCount(n, PRODUCT_TILES) - 1)
End Sub

Public Sub CategoryTileClick(ByVal frm As Access.Form, ByVal Index As Long)
    Dim id As String
    id = Nz(frm.Controls("btnCat" & Index).Tag, "")
    If Len(id) = 0 Then Exit Sub
    m_catID = CLng(id)
    m_prodPage = 0
    LoadCategories frm
    LoadProducts frm
End Sub

Public Sub CategoryPage(ByVal frm As Access.Form, ByVal Delta As Long)
    m_catPage = m_catPage + Delta
    If m_catPage < 0 Then m_catPage = 0
    LoadCategories frm
End Sub

Public Sub ProductPage(ByVal frm As Access.Form, ByVal Delta As Long)
    m_prodPage = m_prodPage + Delta
    If m_prodPage < 0 Then m_prodPage = 0
    LoadProducts frm
End Sub

Public Sub ProductTileClick(ByVal frm As Access.Form, ByVal Index As Long)
    Dim id As String
    id = Nz(frm.Controls("btnProd" & Index).Tag, "")
    If Len(id) = 0 Then Exit Sub
    AddLine frm, CLng(id), 1                               ' modPOS: same cart as the shop
End Sub

'------------------------------------------------------------------------------
' Cart rows (frmTouchLines)
'------------------------------------------------------------------------------
Public Sub TouchQtyStep(ByVal sf As Access.Form, ByVal Delta As Currency)
    Dim q As Currency
    If sf.NewRecord Or IsNull(sf!LineNo.Value) Then Exit Sub
    q = Nz(sf!Quantity.Value, 0) + Delta
    If q <= 0 Then
        CurrentDb.Execute "DELETE FROM tmpPOSLines WHERE LineNo = " & sf!LineNo.Value, dbFailOnError
    Else
        CurrentDb.Execute "UPDATE tmpPOSLines SET Quantity = " & Str$(q) & " WHERE LineNo = " & sf!LineNo.Value, _
                          dbFailOnError
    End If
    RecalcPOS sf.Parent
End Sub

'------------------------------------------------------------------------------
' Payment
'------------------------------------------------------------------------------
Public Sub TouchPayCash(ByVal frm As Access.Form)
    If DCount("*", "tmpPOSLines") = 0 Then
        frm!lblStatus.Caption = "√÷› ’‰›« Ê«Õœ« ⁄·Ï «·√ﬁ·."
        Exit Sub
    End If
    RecalcPOS frm
    TempVars.Add "TouchPayTotal", CDbl(CalcTotal("TOTAL"))
    TempVars.Add "TouchPayOK", False
    DoCmd.OpenForm "frmTouchPay", acNormal, , , , acDialog
    If Not Nz(TempVars("TouchPayOK"), False) Then Exit Sub
    frm!cboPaymentType.Value = "CASH"
    frm!cboPaymentMethod.Value = 1
    frm!txtTendered.Value = CCur(TempVars("TouchPayAmount"))
    RecalcPOS frm
    SavePOS frm, True
End Sub

Public Sub TouchPayCard(ByVal frm As Access.Form)
    If DCount("*", "tmpPOSLines") = 0 Then
        frm!lblStatus.Caption = "√÷› ’‰›« Ê«Õœ« ⁄·Ï «·√ﬁ·."
        Exit Sub
    End If
    frm!cboPaymentType.Value = "CASH"
    frm!cboPaymentMethod.Value = CARD_METHOD_ID
    RecalcPOS frm
    frm!txtTendered.Value = CalcTotal("TOTAL")
    SavePOS frm, True
End Sub

' --- frmTouchPay: number pad --------------------------------------------------
Public Sub PayLoad(ByVal frm As Access.Form)
    frm!lblPayTotal.Caption = Format$(Nz(TempVars("TouchPayTotal"), 0), "#,##0.00")
    frm!txtPayInput.Value = ""
    PayShow frm
End Sub

Public Function PayAmountOf(ByVal Typed As String, ByVal Total As Double) As Double
    ' Nothing typed = the exact total.
    If Len(Typed) = 0 Or Typed = "." Then
        PayAmountOf = Total
    Else
        PayAmountOf = Val(Typed)
    End If
End Function

Public Function PayAppend(ByVal Typed As String, ByVal Key As String) As String
    ' Number pad: digits, one decimal point, two decimals at most, "<" deletes, "C" clears.
    Select Case Key
        Case "C"
            PayAppend = ""
        Case "<"
            If Len(Typed) > 0 Then PayAppend = Left$(Typed, Len(Typed) - 1)
        Case "."
            If InStr(Typed, ".") > 0 Then
                PayAppend = Typed
            ElseIf Len(Typed) = 0 Then
                PayAppend = "0."
            Else
                PayAppend = Typed & "."
            End If
        Case Else
            If InStr(Typed, ".") > 0 And Len(Typed) - InStr(Typed, ".") >= 2 Then
                PayAppend = Typed
            ElseIf Len(Typed) >= 9 Then
                PayAppend = Typed
            ElseIf Typed = "0" Then
                PayAppend = Key
            Else
                PayAppend = Typed & Key
            End If
    End Select
End Function

Public Sub PayKey(ByVal frm As Access.Form, ByVal Key As String)
    frm!txtPayInput.Value = PayAppend(Nz(frm!txtPayInput.Value, ""), Key)
    PayShow frm
End Sub

Public Sub PayQuick(ByVal frm As Access.Form, ByVal Amount As Double)
    If Amount <= 0 Then
        frm!txtPayInput.Value = ""
    Else
        frm!txtPayInput.Value = CStr(Amount)
    End If
    PayShow frm
End Sub

Private Sub PayShow(ByVal frm As Access.Form)
    Dim total As Double, amount As Double
    total = Nz(TempVars("TouchPayTotal"), 0)
    amount = PayAmountOf(Nz(frm!txtPayInput.Value, ""), total)
    frm!lblPayAmount.Caption = Format$(amount, "#,##0.00")
    If amount >= total Then
        frm!lblPayChange.Caption = "«·»«ﬁÌ ··⁄„Ì·: " & Format$(amount - total, "#,##0.00")
        frm!lblPayChange.ForeColor = CLR_SUCCESS
    Else
        frm!lblPayChange.Caption = "«·„»·€ √ﬁ· „‰ «·≈Ã„«·Ì »‹ " & Format$(total - amount, "#,##0.00")
        frm!lblPayChange.ForeColor = CLR_DANGER
    End If
End Sub

Public Sub PayConfirm(ByVal frm As Access.Form)
    Dim total As Double, amount As Double
    total = Nz(TempVars("TouchPayTotal"), 0)
    amount = PayAmountOf(Nz(frm!txtPayInput.Value, ""), total)
    If amount < total Then
        Beep
        Exit Sub
    End If
    TempVars.Add "TouchPayAmount", amount
    TempVars.Add "TouchPayOK", True
    DoCmd.Close acForm, frm.Name
End Sub

Public Sub PayCancel(ByVal frm As Access.Form)
    TempVars.Add "TouchPayOK", False
    DoCmd.Close acForm, frm.Name
End Sub

'------------------------------------------------------------------------------
' In-Access test (RunAllTests). Leaves no data: nothing is saved, the cart is emptied.
'------------------------------------------------------------------------------
Public Function TestTouchPOS() As Boolean
    Dim passed As Long, failed As Long, report As String, frm As Access.Form, pid As Variant
    Calendar = vbCalGreg
    EnsureTestUser
    EnsureLocalTables
    g_SilentMode = True
    Debug.Print "=== TestTouchPOS  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    CheckTouch OrderTypeText("DINE_IN", "5") = "ÿ·» œ«Œ·Ì - ÿ«Ê·… 5", "‰’ «·ÿ·» «·œ«Œ·Ì", passed, failed, report
    CheckTouch OrderTypeText("DELIVERY", Null) = "ÿ·»  Ê’Ì·", "‰’ ÿ·» «· Ê’Ì·", passed, failed, report
    CheckTouch PageCount(0, 16) = 1 And PageCount(16, 16) = 1 And PageCount(17, 16) = 2, "⁄œœ «·’›Õ« ", _
               passed, failed, report
    CheckTouch PayAppend(PayAppend(PayAppend("", "5"), "."), "5") = "5.5" And PayAppend("12.34", "9") = "12.34", _
               "·ÊÕ… «·√—ﬁ«„", passed, failed, report
    CheckTouch ImageCandidate("burger.png", "D:\Shop\Images") = "D:\Shop\Images\burger.png", "„”«— «·’Ê—…", _
               passed, failed, report
    CheckTouch Not IsNull(DbValue("SELECT POSMode FROM Settings WHERE SettingID = 1")), _
               "ÕﬁÊ· ‘«‘… «··„” „ÊÃÊœ… (BuildSchema)", passed, failed, report
    pid = DMin("ProductID", "Products", "IsActive = True")
    If IsNull(pid) Then
        Debug.Print "[i] ·«  ÊÃœ „‰ Ã« :  ŒÿÌ «Œ »«— «·‘«‘…"
    ElseIf DCount("*", "tmpPOSLines") > 0 Then
        Debug.Print "[i] ÌÊÃœ ÿ·» €Ì— „ﬂ „·:  ŒÿÌ «Œ »«— «·‘«‘…"
    Else
        DoCmd.OpenForm "frmTouchPOS", acNormal, , , , acHidden
        Set frm = Forms("frmTouchPOS")
        CheckTouch Len(Trim$(Nz(frm!lblCategoryTitle.Caption, ""))) > 0 And frm!btnProd1.Visible, _
                   "«·›∆«  Ê«·√’‰«›  ŸÂ—", passed, failed, report
        ProductTileClick frm, 1
        CheckTouch DCount("*", "tmpPOSLines") = 1, "«·÷€ÿ ⁄·Ï «·’‰› Ì÷Ì›Â ··ÿ·»", passed, failed, report
        ProductTileClick frm, 1
        CheckTouch Nz(DbValue("SELECT Quantity FROM tmpPOSLines"), 0) = 2, "«·÷€ÿ „—… À«‰Ì… Ì“Ìœ «·ﬂ„Ì…", _
                   passed, failed, report
        CheckTouch Nz(frm!lblTotal.Caption, "0.00") <> "0.00" Or Nz(DbValue("SELECT UnitPrice FROM tmpPOSLines"), 0) = 0, _
                   "«·≈Ã„«·Ì ÌıÕ”»", passed, failed, report
        SetOrderType frm, "DELIVERY"
        CheckTouch Len(TouchOrderProblem(frm)) > 0, "«· Ê’Ì· Ìÿ·» «·ÃÊ«· Ê«·⁄‰Ê«‰", passed, failed, report
        frm!txtDeliveryPhone.Value = "0500000000"
        frm!txtDeliveryAddress.Value = "TEST"
        CheckTouch Len(TouchOrderProblem(frm)) = 0, "«· Ê’Ì· „ﬂ „· «·»Ì«‰« ", passed, failed, report
        SetOrderType frm, "DINE_IN"
        CheckTouch frm!cboTable.Visible And Not frm!txtDeliveryPhone.Visible, "«·ÿ·» «·œ«Œ·Ì Ì⁄—÷ —ﬁ„ «·ÿ«Ê·…", _
                   passed, failed, report
        CurrentDb.Execute "DELETE FROM tmpPOSLines", dbFailOnError
        DoCmd.Close acForm, "frmTouchPOS", acSaveNo
        CheckTouch DCount("*", "tmpPOSLines") = 0, "·„ ÌıÕ›Ÿ ‘Ì¡", passed, failed, report
    End If
    g_SilentMode = False
    Debug.Print "--- ‰ÃÕ: " & passed & " | ›‘·: " & failed
    If failed = 0 Then
        TestMsg "Ã„Ì⁄ «Œ »«—«  ‘«‘… «·„ÿ«⁄„ ‰«ÃÕ… (" & passed & " «Œ »«—«).", vbInformation + MSG_RTL, "TestTouchPOS"
        TestTouchPOS = True
    Else
        TestMsg "‰ÃÕ " & passed & " Ê›‘· " & failed & ":" & vbCrLf & vbCrLf & report, vbExclamation + MSG_RTL, _
                "TestTouchPOS"
    End If
End Function

Private Sub CheckTouch(ByVal ok As Boolean, ByVal Title As String, ByRef passed As Long, ByRef failed As Long, _
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
