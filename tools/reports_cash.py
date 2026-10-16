"""Treasury documents (A4, grouped on the document):
  rptCashVoucher   cash in / cash out / transfer voucher       (qryCashVoucherPrint)
  rptCashClosing   cashier closing slip                        (qryCashClosingPrint)
  rptCashDayBook   cash day book: a page per box and day        (CashDayBookQuery, docs/55)
The treasury list reports (statement, daily, balances, closings) are in reports_catalog.py.
"""

from typing import List

from forms import Control, cm
from reports import (ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, SEC_PAGE_HEADER, SEC_PAGE_FOOTER, txt,
                     lbl, hline)
from reports_docs import W, signatures, store_block


def cash_voucher() -> ReportModel:
    m = ReportModel("rptCashVoucher", "سند نقدية", W,
                    {SEC_HEADER: cm(7.6), SEC_DETAIL: 0, SEC_FOOTER: cm(1.6)},
                    record_source="qryCashVoucherPrint", sorts=[], page_setup=True, no_data="السند غير موجود.")
    H = SEC_HEADER
    half = cm(9.5)
    store_block(m, half)
    txt(m, H, "txtTitle", "VoucherTitle", half, cm(0.1), W - half, cm(0.95), 18, True, align=1)
    txt(m, H, "txtDocNumber", '="رقم السند: " & [VoucherNumber]', half, cm(1.1), W - half, cm(0.5), 10, align=1)
    txt(m, H, "txtDocDate", '="التاريخ: " & GDate([VoucherDate],True)', half, cm(1.65), W - half, cm(0.5), 10,
        align=1)
    hline(m, H, "lnTop", cm(2.3), W)
    txt(m, H, "txtParty", '=IIf([VoucherType]="IN","استلمنا من: ",IIf([VoucherType]="OUT","صرفنا إلى: ",'
                          '"حُوِّل إلى: ")) & IIf([VoucherType]="TRANSFER",[ToBoxName],Nz([PartyName],"-"))',
        0, cm(2.6), W, cm(0.7), 13, True)
    m.add(H, Control("rect", "boxAmount", 0, cm(3.5), cm(6.0), cm(1.2), {"BorderStyle": 1}, decorative=True))
    txt(m, H, "txtAmount", '=Format([Amount],"#,##0.00") & " " & CurrencyWord()', cm(0.1), cm(3.6), cm(5.8), cm(1.0), 18,
        True, align=2)
    txt(m, H, "txtWords", "=AmountInWords([Amount])", cm(6.3), cm(3.65), W - cm(6.3), cm(1.0), 11, grow=True)
    txt(m, H, "txtCategory", '="البند: " & [CategoryName] & IIf(IsNull([ExpenseTypeName]),""," - " & '
                             '[ExpenseTypeName])', 0, cm(5.0), half, cm(0.5), 10)
    txt(m, H, "txtBox", '=IIf([VoucherType]="IN","الصندوق: ","من صندوق: ") & [BoxName]', half, cm(5.0),
        W - half, cm(0.5), 10)
    txt(m, H, "txtNotes", '="البيان: " & Nz([Description],"-")', 0, cm(5.6), W, cm(0.55), 10, grow=True)
    txt(m, H, "txtEmployee", '="الموظف: " & [EmployeeName]', 0, cm(6.3), W, cm(0.5), 9)
    signatures(m, SEC_FOOTER, cm(0.5), ["المستلم", "أمين الصندوق", "المدير"])
    return m


def cash_closing() -> ReportModel:
    m = ReportModel("rptCashClosing", "تصفية يومية الكاشير", W,
                    {SEC_HEADER: cm(12.4), SEC_DETAIL: 0, SEC_FOOTER: cm(1.6)},
                    record_source="qryCashClosingPrint", group="ClosingID", sorts=[], page_setup=True,
                    no_data="التصفية غير موجودة.")
    H = SEC_HEADER
    half = cm(9.5)
    store_block(m, half)
    txt(m, H, "txtTitle", '="تصفية يومية الكاشير"', half, cm(0.1), W - half, cm(0.95), 18, True, align=1)
    txt(m, H, "txtDocNumber", '="رقم التصفية: " & [ClosingNumber]', half, cm(1.1), W - half, cm(0.5), 10, align=1)
    txt(m, H, "txtDocDate", '="التاريخ: " & GDate([ClosingDate],True)', half, cm(1.65), W - half, cm(0.5), 10,
        align=1)
    hline(m, H, "lnTop", cm(2.3), W)
    txt(m, H, "txtBox", '="الصندوق: " & [BoxName] & "    الكاشير: " & [EmployeeName]', 0, cm(2.6), W, cm(0.7),
        13, True)
    txt(m, H, "txtPeriod", '="الفترة: " & IIf(IsNull([PeriodStart]),"من بداية الصندوق","من " & '
                           'GDate([PeriodStart],True)) & " إلى " & GDate([ClosingDate],True)', 0, cm(3.4), W,
        cm(0.5), 10)
    rows = [("رصيد البداية", "OpeningBalance", False), ("المقبوضات (+)", "CashIn", False),
            ("المدفوعات (-)", "CashOut", False), ("الرصيد الدفتري (المفروض)", "ExpectedBalance", True),
            ("النقدية الفعلية بالعدّ", "CountedAmount", True),
            ('=IIf([Difference]<0,"العجز",IIf([Difference]>0,"الزيادة","الفرق"))', "=IIf([Difference]<0,-[Difference],[Difference])", True),
            ('="المرحَّل إلى: " & IIf(IsNull([ToBoxName]),[DestinationName],[ToBoxName])', "TransferAmount", False),
            ("المتبقي في الصندوق (عهدة)", "KeptAmount", False)]
    y = cm(4.2)
    for i, (caption, source, bold) in enumerate(rows):
        if caption.startswith("="):
            txt(m, H, f"txtCap{i + 1}", caption, cm(2.0), y, cm(9.5), cm(0.6), 11, bold)
        else:
            lbl(m, H, f"lblCap{i + 1}", caption, cm(2.0), y, cm(9.5), cm(0.6), 11, bold)
        txt(m, H, f"txtVal{i + 1}", source if source.startswith("=") else source, cm(11.6), y, cm(5.4), cm(0.6),
            12 if bold else 11, bold, align=2, fmt=MONEY)
        y += cm(0.8)
    txt(m, H, "txtNotes", '="ملاحظات: " & Nz([Notes],"-")', 0, y + cm(0.2), W, cm(0.55), 10, grow=True)
    signatures(m, SEC_FOOTER, cm(0.5), ["الكاشير", "المستلم", "المدير"])
    return m


# the day book columns: caption, source, width (cm), format
DAYBOOK_COLS = [("الوقت", '=Format([MoveDate],"hh:nn")', 1.3, None),
                ("الحركة", "MoveTypeName", 3.4, None),
                ("المستند", "DocNumber", 2.3, None),
                ("الجهة", "PartyName", 3.4, None),
                ("البيان", "Details", 3.9, None),
                ("مقبوض", "AmountIn", 2.2, MONEY),
                ("مدفوع", "AmountOut", 2.2, MONEY)]


def cash_day_book() -> ReportModel:
    """A page per box and day: the opening of the day, the receipts then the payments in detail, the totals,
    the closing of the day and the cashier closings of that day."""
    m = ReportModel("rptCashDayBook", "يومية الصندوق التفصيلية", W,
                    {SEC_PAGE_HEADER: cm(1.45), SEC_HEADER: cm(2.35), SEC_DETAIL: cm(0.55), SEC_FOOTER: cm(3.3),
                     SEC_PAGE_FOOTER: cm(0.6)},
                    record_source="CashDayBookQuery", group="DayKey", sorts=[("SideNo", False), ("MoveDate", False)],
                    page_setup=True, no_data="لا توجد حركة نقدية للاختيارات المحددة.")
    P, H, D, F = SEC_PAGE_HEADER, SEC_HEADER, SEC_DETAIL, SEC_FOOTER
    half = W // 2
    txt(m, P, "txtStoreName", '=Nz(SettingValue("StoreName"),"")', 0, cm(0.05), half, cm(0.55), 10, True)
    lbl(m, P, "lblTitle", "يومية الصندوق", half, cm(0.0), W - half, cm(0.7), 14, True, align=1)
    txt(m, P, "txtCriteria", "=ReportCriteria()", 0, cm(0.8), W, cm(0.5), 9, align=2)
    # the box and the day
    txt(m, H, "txtBoxDay", '="الصندوق: " & [BoxName] & "    اليوم: " & GDate([CashDay])', 0, cm(0.1), W, cm(0.75),
        13, True)
    lbl(m, H, "lblOpening", "رصيد أول اليوم", 0, cm(0.95), cm(4.0), cm(0.6), 11, True)
    txt(m, H, "txtOpening", "DayOpening", cm(4.1), cm(0.95), cm(3.6), cm(0.6), 11, True, align=2, fmt=MONEY)
    m.add(H, Control("rect", "boxColumns", 0, cm(1.7), W, cm(0.6), {}, decorative=True))
    x = 0
    for i, (caption, _, width, _) in enumerate(DAYBOOK_COLS):
        lbl(m, H, f"lblCol{i + 1}", caption, x, cm(1.75), cm(width), cm(0.5), 8, True, align=2)
        x += cm(width) + cm(0.05)
    # the moves: the receipts first, then the payments (sorted on SideNo)
    x = 0
    for i, (_, source, width, fmt) in enumerate(DAYBOOK_COLS):
        txt(m, D, f"txtCol{i + 1}", source, x, cm(0.02), cm(width), cm(0.5), 8, align=2 if fmt else 0, fmt=fmt,
            grow=source == "Details")
        x += cm(width) + cm(0.05)
    # the totals of the day
    hline(m, F, "lnTotals", cm(0.05), W)
    rows = [("إجمالي المقبوضات", "=Sum([AmountIn])", False),
            ("إجمالي المصروفات", "=Sum([AmountOut])", False),
            ("رصيد آخر اليوم", "=[DayOpening]+Sum([AmountIn])-Sum([AmountOut])", True)]
    y = cm(0.15)
    for i, (caption, source, bold) in enumerate(rows):
        lbl(m, F, f"lblTotal{i + 1}", caption, 0, y, cm(5.0), cm(0.55), 10, bold)
        txt(m, F, f"txtTotal{i + 1}", source, cm(5.1), y, cm(3.6), cm(0.55), 10, bold, align=2, fmt=MONEY)
        y += cm(0.65)
    # the cashier closings of the day: how many, the cash counted and the difference
    x = 0
    for i, (caption, source, fmt) in enumerate([("تصفيات الكاشير", "ClosingCount", None),
                                                 ("النقدية بالعدّ", "CountedTotal", MONEY),
                                                 ("الفرق", "DifferenceTotal", MONEY)]):
        lbl(m, F, f"lblClosing{i + 1}", caption, x, y + cm(0.05), cm(3.0), cm(0.55), 9)
        txt(m, F, f"txtClosing{i + 1}", source, x + cm(3.05), y + cm(0.05), cm(2.9), cm(0.55), 9, align=2, fmt=fmt)
        x += cm(6.3)
    txt(m, SEC_PAGE_FOOTER, "txtPage", '="صفحة " & [Page] & " من " & [Pages]', 0, cm(0.1), W, cm(0.45), 8, align=2)
    m.events.append("m_rpt.Section(6).ForceNewPage = 2      ' a page per box and day (after the day totals)")
    return m


def cash_reports() -> List[ReportModel]:
    return [cash_voucher(), cash_closing(), cash_day_book()]
