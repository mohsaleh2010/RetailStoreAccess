"""Treasury documents (A4, grouped on the document):
  rptCashVoucher   cash in / cash out / transfer voucher       (qryCashVoucherPrint)
  rptCashClosing   cashier closing slip                        (qryCashClosingPrint)
The treasury list reports (statement, daily, balances, closings) are in reports_catalog.py.
"""

from typing import List

from forms import Control, cm
from reports import ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, txt, lbl, hline
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
    txt(m, H, "txtAmount", '=Format([Amount],"#,##0.00") & " ريال"', cm(0.1), cm(3.6), cm(5.8), cm(1.0), 18,
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


def cash_reports() -> List[ReportModel]:
    return [cash_voucher(), cash_closing()]
