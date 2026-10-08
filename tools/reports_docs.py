"""Phase 8 document reports (A4, grouped on the document):
  rptPurchaseDocument   purchase invoice / purchase return     (qryPurchaseDocPrint)
  rptVoucher            receipt voucher / payment voucher      (qryVoucherPrint)
  rptStockCount         stocktake sheet: blank to count on paper while the
                        count is open, with the differences once counted
"""

from typing import List

from forms import Control, Sym, cm
from reports import (ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, SEC_PAGE_HEADER, SEC_PAGE_FOOTER,
                     setting, txt, lbl, hline)

QTY = "#,##0.###"
W = cm(19.0)


def store_block(m: ReportModel, half: int):
    H = SEC_HEADER
    txt(m, H, "txtStoreName", setting("StoreName"), 0, cm(0.1), half, cm(0.75), 14, True)
    txt(m, H, "txtStoreVat", setting("VATNumber", "الرقم الضريبي: "), 0, cm(0.9), half, cm(0.5), 9)
    txt(m, H, "txtStorePhone", '="هاتف: " & Nz(SettingValue("Phone"),"-") & "   " & '
                               'Nz(SettingValue("City"),"")', 0, cm(1.4), half, cm(0.5), 9)


def signatures(m: ReportModel, sec: int, y: int, captions: List[str]):
    part = W // len(captions)
    for i, caption in enumerate(captions):
        lbl(m, sec, f"lblSign{i + 1}", caption + ": ....................", i * part, y, part, cm(0.55), 10)


def columns(m: ReportModel, y: int, cols):
    m.add(SEC_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.65),
                              {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    x = 0
    for i, (title, source, width, fmt, props) in enumerate(cols):
        cw = cm(width) if i < len(cols) - 1 else W - x
        lbl(m, SEC_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        t = txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.04), cw, cm(0.5), 8,
                align=0 if fmt is None and source != "LineNumber" else 2, fmt=fmt,
                grow=source == "ProductName")
        t.props.update(props)
        x += cw
    assert x == W


def purchase_document() -> ReportModel:
    m = ReportModel("rptPurchaseDocument", "فاتورة / مرتجع مشتريات", W,
                    {SEC_HEADER: cm(5.9), SEC_DETAIL: cm(0.6), SEC_FOOTER: cm(5.2)},
                    record_source="qryPurchaseDocPrint", page_setup=True,
                    no_data="المستند غير موجود.")
    H = SEC_HEADER
    half = cm(9.3)
    store_block(m, half)
    x2 = W - half
    txt(m, H, "txtTitle", '=IIf([DocKind]="RETURN","مرتجع مشتريات","فاتورة مشتريات")', x2, cm(0.1), half,
        cm(0.85), 16, True, align=1)
    y = cm(1.0)
    for name, src in [
            ("txtDocNumber", '="الرقم: " & [DocNumber]'),
            ("txtDocDate", '="التاريخ: " & GDate([DocDate],True)'),
            ("txtSupplierRef", '="رقم فاتورة المورد: " & Nz([SupplierInvoiceNo],"-")'),
            ("txtOriginal", '=IIf(Len(Nz([OriginalNumber],""))>0,"عن فاتورة الشراء: " & [OriginalNumber],"")')]:
        txt(m, H, name, src, x2, y, half, cm(0.45), 9, align=1)
        y += cm(0.45)
    by = cm(3.0)
    m.add(H, Control("rect", "boxSupplier", 0, by, W, cm(1.95), {"BorderStyle": 1}, decorative=True))
    txt(m, H, "txtSupplier", '="المورد: " & [SupplierName]', cm(0.2), by + cm(0.1), cm(9.0), cm(0.5), 10, True)
    txt(m, H, "txtSupplierVat", '=IIf(Len(Nz([SupplierVAT],""))>0,"الرقم الضريبي: " & [SupplierVAT],"")',
        cm(9.4), by + cm(0.1), cm(9.4), cm(0.5), 9)
    txt(m, H, "txtPayment", '=IIf([DocKind]="RETURN",IIf([PaymentType]="CASH","استرداد نقدي",'
                            '"خصم من رصيد المورد"),IIf([PaymentType]="CREDIT","شراء آجل","شراء نقدي"))',
        cm(0.2), by + cm(0.7), cm(9.0), cm(0.5), 9)
    txt(m, H, "txtEmployee", '="الموظف: " & [EmployeeName]', cm(9.4), by + cm(0.7), cm(9.4), cm(0.5), 9)
    txt(m, H, "txtReason", '=IIf(Len(Nz([Reason],""))>0,"سبب الإرجاع: " & [Reason],"")', cm(0.2),
        by + cm(1.3), W - cm(0.4), cm(0.5), 9)
    columns(m, cm(5.15), [
        ("#", "LineNumber", 0.8, None, {}), ("الكود", "ProductCode", 2.2, None, {}),
        ("الصنف", "ProductName", 5.0, None, {}), ("الوحدة", "UnitName", 1.4, None, {}),
        ("الكمية", "Quantity", 1.5, QTY, {}), ("تكلفة الوحدة", "UnitCost", 2.0, "#,##0.00##", {}),
        ("الخصم", "LineDiscount", 1.5, MONEY, {}), ("الضريبة", "LineTax", 1.6, MONEY, {}),
        ("الإجمالي", "LineTotal", 3.0, MONEY, {})])

    F = SEC_FOOTER
    hline(m, F, "lnTotals", cm(0.08), W)
    tx, lw, vw = cm(11.4), cm(4.4), cm(3.2)
    y = cm(0.2)
    for cap, src, bold in [("الإجمالي قبل الخصم", "DocSubTotal", False), ("الخصم", "DocDiscount", False),
                           ("الخاضع للضريبة", "TaxableAmount", False), ("ضريبة القيمة المضافة", "DocTax", False),
                           ("الإجمالي شامل الضريبة", "TotalAmount", True)]:
        lbl(m, F, f"lblCap{src}", cap, tx, y, lw, cm(0.45), 10 if bold else 9, bold)
        txt(m, F, f"txt{src}", src, tx + lw, y, vw, cm(0.45), 10 if bold else 9, bold, align=1, fmt=MONEY)
        y += cm(0.5)
    txt(m, F, "txtCapPaid", '=IIf([DocKind]="RETURN","المسترد نقدًا","المدفوع")', tx, y, lw, cm(0.45), 9)
    txt(m, F, "txtPaidAmount", "PaidAmount", tx + lw, y, vw, cm(0.45), 9, align=1, fmt=MONEY)
    y += cm(0.5)
    txt(m, F, "txtCapRemaining", '=IIf([DocKind]="RETURN","خصم من رصيد المورد","المتبقي على الحساب")', tx, y,
        lw, cm(0.45), 9)
    txt(m, F, "txtRemainingAmount", "RemainingAmount", tx + lw, y, vw, cm(0.45), 9, align=1, fmt=MONEY)
    txt(m, F, "txtWords", "=AmountInWords([TotalAmount])", 0, cm(0.2), cm(11.0), cm(1.0), 10, grow=True)
    signatures(m, F, cm(4.3), ["المستلم", "أمين المخزن", "المدير"])
    return m


def voucher() -> ReportModel:
    m = ReportModel("rptVoucher", "سند قبض / سند صرف", W,
                    {SEC_HEADER: cm(7.0), SEC_DETAIL: 0, SEC_FOOTER: cm(1.6)},
                    record_source="qryVoucherPrint", page_setup=True, no_data="السند غير موجود.")
    H = SEC_HEADER
    half = cm(9.5)
    store_block(m, half)
    txt(m, H, "txtTitle", '=IIf([DocKind]="RECEIPT","سند قبض","سند صرف")', half, cm(0.1), W - half, cm(0.95),
        18, True, align=1)
    txt(m, H, "txtDocNumber", '="رقم السند: " & [DocNumber]', half, cm(1.1), W - half, cm(0.5), 10, align=1)
    txt(m, H, "txtDocDate", '="التاريخ: " & GDate([DocDate],True)', half, cm(1.65), W - half, cm(0.5), 10,
        align=1)
    hline(m, H, "lnTop", cm(2.3), W)
    txt(m, H, "txtParty", '=IIf([DocKind]="RECEIPT","استلمنا من: ","صرفنا إلى: ") & [PartyName]', 0, cm(2.6),
        W, cm(0.7), 13, True)
    m.add(H, Control("rect", "boxAmount", 0, cm(3.5), cm(6.0), cm(1.2), {"BorderStyle": 1}, decorative=True))
    txt(m, H, "txtAmount", '=Format([Amount],"#,##0.00") & " " & CurrencyWord()', cm(0.1), cm(3.6), cm(5.8), cm(1.0), 18,
        True, align=2)
    txt(m, H, "txtWords", "=AmountInWords([Amount])", cm(6.3), cm(3.65), W - cm(6.3), cm(1.0), 11, grow=True)
    txt(m, H, "txtMethod", '="طريقة الدفع: " & [MethodName]', 0, cm(5.0), half, cm(0.5), 10)
    txt(m, H, "txtEmployee", '="الموظف: " & [EmployeeName]', half, cm(5.0), W - half, cm(0.5), 10)
    txt(m, H, "txtNotes", '="البيان: " & Nz([Notes],"-")', 0, cm(5.6), W, cm(0.55), 10, grow=True)
    txt(m, H, "txtBalance", '=IIf([DocKind]="RECEIPT","الرصيد المتبقي على العميل حاليًا: ",'
                            '"الرصيد المستحق للمورد حاليًا: ") & Format([PartyBalance],"#,##0.00")',
        0, cm(6.3), W, cm(0.5), 9)
    signatures(m, SEC_FOOTER, cm(0.5), ["المستلم", "المحاسب", "المدير"])
    return m


def stock_count_sheet() -> ReportModel:
    m = ReportModel("rptStockCount", "ورقة الجرد", W,
                    {SEC_HEADER: cm(3.6), SEC_DETAIL: cm(0.6), SEC_FOOTER: cm(2.6)},
                    record_source="StockCountQuery", group="StockCountID", sorts=[("ProductName", False)],
                    page_setup=True, no_data="جلسة الجرد غير موجودة.")
    H = SEC_HEADER
    half = cm(9.3)
    store_block(m, half)
    x2 = W - half
    txt(m, H, "txtTitle", '="ورقة جرد " & IIf([Status]="POSTED","(مُرحّل)",IIf([Status]="CANCELLED",'
                          '"(ملغى)","(مفتوح)"))', x2, cm(0.1), half, cm(0.85), 16, True, align=1)
    txt(m, H, "txtCountNumber", '="رقم الجرد: " & [CountNumber] & "    التاريخ: " & GDate([CountDate])',
        x2, cm(1.0), half, cm(0.5), 9, align=1)
    txt(m, H, "txtCategory", '="التصنيف: " & Nz([CategoryName],"كل المنتجات")', x2, cm(1.55), half, cm(0.5), 9,
        align=1)
    columns(m, cm(2.9), [
        ("#", "=1", 0.9, None, {"RunningSum": 1}), ("الكود", "ProductCode", 2.4, None, {}),
        ("الصنف", "ProductName", 6.0, None, {}), ("المسجل", "SystemQuantity", 2.2, QTY, {}),
        ("الفعلي", "ActualQuantity", 2.3, QTY, {}),
        ("الفرق", "=IIf(IsNull([ActualQuantity]),Null,[Difference])", 2.0, QTY, {}),
        ("قيمة الفرق", "=IIf(IsNull([ActualQuantity]),Null,[DifferenceValue])", 3.2, MONEY, {})])
    F = SEC_FOOTER
    hline(m, F, "lnTotals", cm(0.08), W)
    txt(m, F, "txtSummary", '="تم عدّ " & Count([ActualQuantity]) & " من " & Count(*) & " صنف    العجز: " & '
                            'Format(-Sum(IIf([DifferenceValue]<0,[DifferenceValue],0)),"#,##0.00") & '
                            '"    الزيادة: " & Format(Sum(IIf([DifferenceValue]>0,[DifferenceValue],0)),"#,##0.00")',
        0, cm(0.2), W, cm(0.55), 10, True)
    signatures(m, F, cm(1.5), ["القائم بالجرد", "المراجع", "المدير"])
    return m


LABEL_SQL = ("SELECT q.LineNo, n.N, p.ProductName, p.ProductCode, p.Barcode, p.SellingPrice "
             "FROM tmpLabelQueue AS q, tmpLabelNumbers AS n, Products AS p "
             "WHERE p.ProductID = q.ProductID AND n.N <= q.Copies")


def barcode_labels() -> ReportModel:
    """One label per row (copies via tmpLabelNumbers). The size, columns, margins, printer and text
    lines come from LabelSettings: modLabels.ApplyLabelLayout rewrites this design before printing."""
    w, h = cm(3.8), cm(2.5)
    m = ReportModel("rptBarcodeLabels", "ملصقات الباركود", w, {SEC_DETAIL: h},
                    record_source=LABEL_SQL, group="", sorts=[("LineNo", False), ("N", False)],
                    no_data="قائمة الملصقات فارغة.", prepare=["EnsureLabelTables"])
    D = SEC_DETAIL
    pad, line = cm(0.1), cm(0.4)
    txt(m, D, "txtTop1", '="المتجر"', pad, pad, w - 2 * pad, line, 7, align=2)
    txt(m, D, "txtTop2", "=[ProductName]", pad, pad + line, w - 2 * pad, line, 7, align=2)
    m.add(D, Control("rect", "boxBar", pad, pad + 2 * line + cm(0.05), w - 2 * pad, cm(0.9),
                     {"BorderStyle": 0, "BackStyle": 0}, decorative=True))
    txt(m, D, "txtBottom1", "=LabelCode([Barcode],[ProductCode])", pad, h - pad - 2 * line, w - 2 * pad, line,
        7, align=2)
    txt(m, D, "txtBottom2", "=LabelPrice([SellingPrice])", pad, h - pad - line, w - 2 * pad, line, 7, True,
        align=2)
    txt(m, D, "txtCode", "=LabelCode([Barcode],[ProductCode])", 0, 0, cm(0.1), cm(0.1), 6, visible=False)
    m.events += ['m_rpt.Section(0).Name = "secLabel"', "m_rpt.Section(0).OnPrint = EP", "m_rpt.OnOpen = EP"]
    m.code += ["Private Sub Report_Open(Cancel As Integer)", "    LabelReportOpen", "End Sub",
               "Private Sub secLabel_Print(Cancel As Integer, PrintCount As Integer)",
               "    If Me.HasData = 0 Then Exit Sub",
               '    DrawBarcode Me, Nz(Me!txtCode.Value, ""), Me!boxBar.Left, Me!boxBar.Top, Me!boxBar.Width, _',
               "                Me!boxBar.Height, LabelBarWidth()",
               "End Sub"]
    return m


def statistics() -> ReportModel:
    """Charts drawn by modCharts.DrawStatistics in the Print event of the (unbound) detail section."""
    m = ReportModel("rptStatistics", "الإحصائيات والرسوم البيانية", W,
                    {SEC_PAGE_HEADER: cm(2.0), SEC_DETAIL: cm(23.6), SEC_PAGE_FOOTER: cm(0.7)},
                    record_source="", group="", sorts=[], page_setup=True)
    half = cm(9.3)
    txt(m, SEC_PAGE_HEADER, "txtStoreName", setting("StoreName"), W - half, cm(0.1), half, cm(0.7), 13, True,
        align=3)
    lbl(m, SEC_PAGE_HEADER, "lblTitle", "الإحصائيات والرسوم البيانية", 0, cm(0.1), half, cm(0.75), 16, True)
    txt(m, SEC_PAGE_HEADER, "txtCriteria", "=ReportCriteria()", 0, cm(0.95), W, cm(0.5), 9, align=2)
    hline(m, SEC_PAGE_HEADER, "lnHeader", cm(1.7), W)
    m.add(SEC_DETAIL, Control("rect", "boxCharts", 0, 0, W, cm(23.6), {"BorderStyle": 0, "BackStyle": 0},
                              decorative=True))
    txt(m, SEC_PAGE_FOOTER, "txtPrinted", "=ReportPrintedAt()", 0, cm(0.1), cm(11), cm(0.45), 8)
    txt(m, SEC_PAGE_FOOTER, "txtPage", '="صفحة " & [Page] & " من " & [Pages]', cm(11), cm(0.1), W - cm(11),
        cm(0.45), 8, align=3)
    m.events += ['m_rpt.Section(0).Name = "secCharts"', "m_rpt.Section(0).OnPrint = EP"]
    m.code += ["Private Sub secCharts_Print(Cancel As Integer, PrintCount As Integer)",
               "    DrawStatistics Me, Me!boxCharts.Left, Me!boxCharts.Top, Me!boxCharts.Width, Me!boxCharts.Height",
               "End Sub"]
    return m


def document_reports() -> List[ReportModel]:
    return [purchase_document(), voucher(), stock_count_sheet(), barcode_labels(), statistics()]
