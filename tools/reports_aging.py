"""Receivables / payables aging (rptAging, landscape A4, grouped on the party): the open documents
of each customer or supplier with their age columns, the totals of each party and of all.
Source: the local table tmpAging filled by modAging.FillAging."""

from typing import List

from forms import Control, Sym, cm
from reports import (ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, SEC_PAGE_HEADER, SEC_PAGE_FOOTER,
                     SEC_RPT_FOOTER, txt, lbl, hline)

W = cm(27.4)
AGE_COLS = [("غير مستحق", "NotDue"), ("1-30 يومًا", "Days30"), ("31-60", "Days60"), ("61-90", "Days90"),
            ("أكثر من 90", "Over90"), ("رصيد دائن", "Credit"), ("المتبقي", "OpenAmount")]


def aging() -> ReportModel:
    from reports_catalog import title_block, page_footer
    m = ReportModel("rptAging", "أعمار الديون", W,
                    {SEC_PAGE_HEADER: cm(3.0), SEC_HEADER: cm(0.75), SEC_DETAIL: cm(0.55), SEC_FOOTER: cm(0.8),
                     SEC_RPT_FOOTER: cm(1.0), SEC_PAGE_FOOTER: cm(0.6)},
                    record_source="tmpAging", group="PartyID",
                    sorts=[("PartyName", False), ("DueDate", False), ("LineNo", False)], landscape=True,
                    page_setup=True, no_data="لا توجد أرصدة مفتوحة.")
    y = title_block(m, "أعمار الديون", W, True)
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.65),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    cols = [("المستند", "DocTypeName", 2.4, None), ("الرقم", "DocNo", 2.8, None),
            ("التاريخ", "=GDate([DocDate])", 2.0, None), ("الاستحقاق", "=GDate([DueDate])", 2.0, None),
            ("أيام التأخير", "=IIf([DaysLate]>0,[DaysLate],Null)", 1.6, None)]
    cols += [(t, f"=IIf([{c}]=0,Null,[{c}])", 2.37, MONEY) for t, c in AGE_COLS]
    x = 0
    for i, (title, source, width, fmt) in enumerate(cols):
        cw = cm(width) if i < len(cols) - 1 else W - x
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.03), cw, cm(0.5), 8, align=2 if fmt else 0, fmt=fmt)
        cols[i] = (title, source, cw, fmt)
        x += cw
    assert x == W
    txt(m, SEC_HEADER, "txtParty", "PartyName", 0, cm(0.12), W, cm(0.55), 11, True)
    start = sum(c[2] for c in cols[:5])
    for sec, prefix, caption in [(SEC_FOOTER, "txtSum", '="إجمالي " & [PartyName]'),
                                 (SEC_RPT_FOOTER, "txtTotal", '="الإجمالي"')]:
        hline(m, sec, "ln" + prefix, cm(0.04), W)
        txt(m, sec, prefix + "Caption", caption, 0, cm(0.15), start, cm(0.55), 9, True)
        x = start
        for i, (title, column) in enumerate(AGE_COLS):
            cw = cols[5 + i][2]
            txt(m, sec, f"{prefix}{i + 1}", f"=Sum([{column}])", x, cm(0.15), cw, cm(0.55), 9, True, align=2, fmt=MONEY)
            x += cw
    page_footer(m, W)
    return m


def aging_reports() -> List[ReportModel]:
    return [aging(), payroll_sheet(), commission_sheet()]


# ------------------------------------------------------------------ payroll sheet
PAY_COLS = [("الموظف", "EmployeeName", 4.2, None), ("الأساسي", "Basic", 0, "M"), ("السكن", "Housing", 0, "M"),
            ("بدلات أخرى", "OtherAllow", 0, "M"), ("الإضافي", "Overtime", 0, "M"), ("مكافآت", "Additions", 0, "M"),
            ("الإجمالي", "Gross", 0, "M"), ("غياب", "AbsenceDeduction", 0, "M"), ("سلفة", "AdvanceDeduction", 0, "M"),
            ("جزاءات", "OtherDeduction", 0, "M"), ("التأمينات", "GosiEmployee", 0, "M"), ("الصافي", "NetPay", 0, "M")]


def payroll_sheet() -> ReportModel:
    from reports_catalog import title_block, page_footer
    from reports_docs import signatures
    m = ReportModel("rptPayroll", "مسير الرواتب", W,
                    {SEC_PAGE_HEADER: cm(3.0), SEC_DETAIL: cm(0.6), SEC_RPT_FOOTER: cm(2.4), SEC_PAGE_FOOTER: cm(0.6)},
                    record_source="PayrollSheetQuery", group="", sorts=[("EmployeeName", False)], landscape=True,
                    page_setup=True, no_data="المسير بلا أسطر.")
    y = title_block(m, "مسير الرواتب", W, True)
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.65),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    money_w = (W - cm(4.2)) // 11
    x = 0
    for i, (title, source, width, fmt) in enumerate(PAY_COLS):
        cw = cm(width) if width else (money_w if i < len(PAY_COLS) - 1 else W - x)
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.03), cw, cm(0.5), 8, bold=source == "NetPay",
            align=2 if fmt else 0, fmt=MONEY if fmt else None)
        if fmt:
            txt(m, SEC_RPT_FOOTER, f"txtSum{i + 1}", f"=Sum([{source}])", x, cm(0.15), cw, cm(0.55), 8, True, align=2,
                fmt=MONEY)
        x += cw
    assert x == W
    hline(m, SEC_RPT_FOOTER, "lnTotals", cm(0.05), W)
    txt(m, SEC_RPT_FOOTER, "txtSumCaption", '="الإجمالي"', 0, cm(0.15), cm(4.2), cm(0.55), 9, True)
    signatures(m, SEC_RPT_FOOTER, cm(1.3), ["أعدّه", "راجعه", "اعتمده"])
    page_footer(m, W)
    return m


# ------------------------------------------------------------------ commission run sheet
COM_COLS = [("المندوب", "RepName", 5.0, None), ("صافي المبيعات", "NetSales", 3.0, "M"), ("التحصيل", "Collections", 3.0, "M"),
            ("الأساس", "BaseName", 2.0, None), ("مبلغ الأساس", "BaseAmount", 3.0, "M"),
            ("النسبة", "CommissionRate", 1.8, "%"), ("التعديل", "Adjustment", 2.6, "M"),
            ("العمولة", "Commission", 3.0, "M"), ("ملاحظات", "Notes", 0, None)]


def commission_sheet() -> ReportModel:
    from reports_catalog import title_block, page_footer, PCT
    from reports_docs import signatures
    m = ReportModel("rptCommissionRun", "مسير عمولات المندوبين", W,
                    {SEC_PAGE_HEADER: cm(3.0), SEC_DETAIL: cm(0.6), SEC_RPT_FOOTER: cm(2.4), SEC_PAGE_FOOTER: cm(0.6)},
                    record_source="CommissionSheetQuery", group="", sorts=[("RepName", False)], landscape=True,
                    page_setup=True, no_data="المسير بلا أسطر.")
    y = title_block(m, "مسير عمولات المندوبين", W, True)
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.65),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    x = 0
    for i, (title, source, width, fmt) in enumerate(COM_COLS):
        cw = cm(width) if width else W - x
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.03), cw, cm(0.5), 8, bold=source == "Commission",
            align=2 if fmt else 0, fmt={"M": MONEY, "%": PCT}.get(fmt))
        if fmt == "M" and source != "BaseAmount":
            txt(m, SEC_RPT_FOOTER, f"txtSum{i + 1}", f"=Sum([{source}])", x, cm(0.15), cw, cm(0.55), 8, True, align=2,
                fmt=MONEY)
        x += cw
    assert x == W
    hline(m, SEC_RPT_FOOTER, "lnTotals", cm(0.05), W)
    txt(m, SEC_RPT_FOOTER, "txtSumCaption", '="الإجمالي"', 0, cm(0.15), cm(5.0), cm(0.55), 9, True)
    signatures(m, SEC_RPT_FOOTER, cm(1.3), ["أعدّه", "راجعه", "اعتمده"])
    page_footer(m, W)
    return m
