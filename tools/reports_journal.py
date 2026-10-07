"""Journal entry print (A4, grouped on the entry): rptJournalEntry (qryJournalEntryPrint).
Account statement (rptAccountStatement, AccountStatementQuery) and general ledger
(rptGeneralLedger, GeneralLedgerQuery): A4, the opening balance first, a running balance per
account (RunningSum over the group) and the totals of each account.
The journal and the trial balance are list reports of reports_catalog.py."""

from typing import List

from reports import (ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, SEC_PAGE_HEADER, SEC_PAGE_FOOTER,
                     SEC_RPT_FOOTER,
                     lbl, txt, hline)
from reports_docs import W, columns, signatures, store_block
from forms import Control, Sym, cm


def journal_entry() -> ReportModel:
    m = ReportModel("rptJournalEntry", "قيد يومية", W,
                    {SEC_HEADER: cm(4.4), SEC_DETAIL: cm(0.6), SEC_FOOTER: cm(2.6)},
                    record_source="qryJournalEntryPrint", group="EntryID", page_setup=True,
                    no_data="القيد غير موجود.")
    H = SEC_HEADER
    half = cm(9.5)
    store_block(m, half)
    txt(m, H, "txtTitle", '="قيد يومية"', half, cm(0.1), W - half, cm(0.95), 18, True, align=1)
    txt(m, H, "txtDocNumber", '="رقم القيد: " & [EntryNumber]', half, cm(1.1), W - half, cm(0.5), 10, align=1)
    txt(m, H, "txtDocDate", '="التاريخ: " & GDate([EntryDate],True)', half, cm(1.65), W - half, cm(0.5), 10,
        align=1)
    txt(m, H, "txtSource", '="العملية: " & [TypeName] & " " & Nz([SourceNumber],"")', 0, cm(2.3), W, cm(0.55),
        11, True)
    txt(m, H, "txtDescription", '="البيان: " & Nz([Description],"-")', 0, cm(2.9), W, cm(0.5), 9)
    columns(m, cm(3.6), [
        ("الحساب", "AccountCode", 2.2, None, {}), ("اسم الحساب", "AccountName", 5.2, None, {}),
        ("البيان", "LineText", 6.0, None, {}), ("مدين", "=IIf([Debit]=0,Null,[Debit])", 2.8, MONEY, {}),
        ("دائن", "=IIf([Credit]=0,Null,[Credit])", 2.8, MONEY, {})])
    F = SEC_FOOTER
    hline(m, F, "lnTotals", cm(0.08), W)
    txt(m, F, "txtTotalCaption", '="الإجمالي"', 0, cm(0.2), cm(13.4), cm(0.55), 10, True)
    txt(m, F, "txtTotalDebit", "=Sum([Debit])", cm(13.4), cm(0.2), cm(2.8), cm(0.55), 10, True, align=2, fmt=MONEY)
    txt(m, F, "txtTotalCredit", "=Sum([Credit])", cm(13.4) + cm(2.8), cm(0.2), W - cm(13.4) - cm(2.8), cm(0.55), 10, True, align=2,
        fmt=MONEY)
    signatures(m, F, cm(1.4), ["المحاسب", "المراجع", "المدير"])
    return m


LEDGER_COLS = [  # (title, source, width cm)
    ("التاريخ", "=GDate([LineDate])", 2.0), ("القيد", "EntryNo", 2.0), ("العملية", "KindName", 2.5),
    ("المستند", "DocNo", 2.0), ("البيان", "Details", 4.4), ("مدين", "=IIf([LineDebit]=0,Null,[LineDebit])", 2.0),
    ("دائن", "=IIf([LineCredit]=0,Null,[LineCredit])", 2.0), ("الرصيد", "=[LineDebit]-[LineCredit]", 2.1)]


def _ledger(name, caption, query, group, account_expr, sub_col):
    from reports_catalog import title_block, page_footer
    cols = list(LEDGER_COLS)
    if sub_col:                                  # a main account: the sub-account of each line
        cols[4] = ("البيان", "Details", 2.6)
        cols.insert(5, ("الحساب", "SubName", 1.8))
    m = ReportModel(name, caption, W,
                    {SEC_PAGE_HEADER: cm(2.95), SEC_HEADER: cm(0.8), SEC_DETAIL: cm(0.55),
                     SEC_FOOTER: cm(2.0), SEC_PAGE_FOOTER: cm(0.6)},
                    record_source=query, group=group,
                    sorts=[("SortKey", False), ("LineDate", False), ("EntryNo", False)], page_setup=True,
                    no_data="لا توجد قيود للاختيارات المحددة.")
    y = title_block(m, caption, W, True)
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.65),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    txt(m, SEC_HEADER, "txtAccount", account_expr, 0, cm(0.12), W, cm(0.6), 11, True)
    x = 0
    for i, (title, source, width) in enumerate(cols):
        cw = cm(width) if i < len(cols) - 1 else W - x
        money = title in ("مدين", "دائن", "الرصيد")
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.08), cw, cm(0.5), 8, True, align=2)
        t = txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.03), cw, cm(0.5), 8, align=2 if money else 0,
                fmt=MONEY if money else None, grow=title == "البيان")
        if title == "الرصيد":
            t.props["RunningSum"] = 1            # over the account (the group)
        x += cw
    assert x == W
    F = SEC_FOOTER
    hline(m, F, "lnTotals", cm(0.05), W)
    part = W // 4
    for i, (caption, expr) in enumerate([
            ("رصيد أول المدة", "=Sum(IIf([SortKey]=0,[LineDebit]-[LineCredit],0))"),
            ("مدين الفترة", "=Sum(IIf([SortKey]=1,[LineDebit],0))"),
            ("دائن الفترة", "=Sum(IIf([SortKey]=1,[LineCredit],0))"),
            ("الرصيد الختامي", "=Sum([LineDebit]-[LineCredit])")]):
        lbl(m, F, f"lblSum{i + 1}", caption, i * part, cm(0.15), part, cm(0.5), 9, True, align=2)
        txt(m, F, f"txtSum{i + 1}", expr, i * part, cm(0.7), part, cm(0.55), 11, True, align=2, fmt=MONEY)
    txt(m, F, "txtNature", '=IIf(Sum([LineDebit]-[LineCredit])>=0,"الرصيد مدين","الرصيد دائن")',
        3 * part, cm(1.3), W - 3 * part, cm(0.5), 9, align=2)
    page_footer(m, W)
    return m


def account_statement() -> ReportModel:
    return _ledger("rptAccountStatement", "كشف حساب", "AccountStatementQuery", "StatementAccount",
                   '="الحساب: " & [StatementAccount] & "  " & [StatementName]', True)


def general_ledger() -> ReportModel:
    return _ledger("rptGeneralLedger", "دفتر الأستاذ", "GeneralLedgerQuery", "AccountKey",
                   '=[LedgerCode] & "  " & [LedgerName]', False)


# --------------------------------------------------------------- financial statements
# account amounts in the first column of each period, totals and results in the second
CAPTION_EXPR = ('=IIf([RowKind]="A","      " & [Caption],IIf([RowKind]="S","إجمالي " & [Caption],[Caption]))')
TOTAL_KINDS = '[RowKind]="S" Or [RowKind]="T" Or [RowKind]="R"'


def _financial(name, caption, query, sorts):
    from reports_catalog import title_block, page_footer
    m = ReportModel(name, caption, W,
                    {SEC_PAGE_HEADER: cm(3.6), SEC_DETAIL: cm(0.6), SEC_PAGE_FOOTER: cm(0.6)},
                    record_source=query, group="", sorts=sorts, page_setup=True,
                    no_data="لا توجد قيود للفترة المحددة.")
    y = title_block(m, caption, W, True)
    name_w, col_w = cm(8.2), (W - cm(8.2)) // 4
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(1.25),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    lbl(m, SEC_PAGE_HEADER, "lblItem", "البند", 0, y + cm(0.35), name_w, cm(0.5), 9, True, align=2)
    lbl(m, SEC_PAGE_HEADER, "lblCurrent", "الفترة الحالية", name_w, y + cm(0.08), 2 * col_w, cm(0.5), 9, True,
        align=2)
    lbl(m, SEC_PAGE_HEADER, "lblPrior", "فترة المقارنة", name_w + 2 * col_w, y + cm(0.08), W - name_w - 2 * col_w,
        cm(0.5), 9, True, align=2)
    txt(m, SEC_DETAIL, "txtItem", CAPTION_EXPR, 0, cm(0.05), name_w, cm(0.5), 9, grow=True)
    for i, (head, col, kinds) in enumerate([("الحساب", "CurrentValue", '[RowKind]="A"'),
                                            ("المجموع", "CurrentValue", TOTAL_KINDS),
                                            ("الحساب", "PriorValue", '[RowKind]="A"'),
                                            ("المجموع", "PriorValue", TOTAL_KINDS)]):
        x = name_w + i * col_w
        w = col_w if i < 3 else W - x
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", head, x, y + cm(0.65), w, cm(0.5), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", f"=IIf({kinds},[{col}],Null)", x, cm(0.05), w, cm(0.5), 9,
            bold=i % 2 == 1, align=2, fmt=MONEY)
    page_footer(m, W)
    return m


def income_statement() -> ReportModel:
    return _financial("rptIncomeStatement", "قائمة الدخل", "IncomeStatementQuery",
                      [("Block", False), ("AccountKey", False)])


def balance_sheet() -> ReportModel:
    return _financial("rptBalanceSheet", "الميزانية العمومية (قائمة المركز المالي)", "BalanceSheetQuery",
                      [("ClassNo", False), ("GroupKey", False), ("Pos", False), ("AccountKey", False)])


# ------------------------------------------------------------------ VAT return
def vat_return() -> ReportModel:
    from reports_catalog import title_block, page_footer
    m = ReportModel("rptVatReturn", "إقرار ضريبة القيمة المضافة", W,
                    {SEC_PAGE_HEADER: cm(3.7), SEC_DETAIL: cm(0.7), SEC_RPT_FOOTER: cm(2.2),
                     SEC_PAGE_FOOTER: cm(0.6)},
                    record_source="VatReturnQuery", group="", sorts=[("BoxNo", False)], page_setup=True,
                    no_data="الإقرار غير موجود.")
    y = title_block(m, "إقرار ضريبة القيمة المضافة", W, True)
    txt(m, SEC_PAGE_HEADER, "txtReturn",
        '="الإقرار: " & [ReturnNumber] & "    الحالة: " & IIf([Status]="FILED","معتمد في " & GDate([FiledDate]),'
        '"مسودة (غير معتمد)") & IIf(IsNull([FilingRef]),"","    رقم الإقرار لدى الهيئة: " & [FilingRef])',
        0, y, W, cm(0.55), 10, True)
    y += cm(0.7)
    num_w, amt_w = cm(1.2), cm(3.2)
    text_w = W - num_w - 3 * amt_w
    m.add(SEC_PAGE_HEADER, Control("rect", "boxColumns", 0, y, W, cm(0.7),
                                   {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")}, decorative=True))
    cols = [("البند", "BoxNo", num_w, None), ("الوصف", "BoxText", text_w, None),
            ("المبلغ (بدون الضريبة)", "Amount", amt_w, MONEY), ("التعديلات (المرتجعات)", "Adjust", amt_w, MONEY),
            ("ضريبة القيمة المضافة", "VAT", W - num_w - text_w - 2 * amt_w, MONEY)]
    x = 0
    for i, (title, source, w, fmt) in enumerate(cols):
        lbl(m, SEC_PAGE_HEADER, f"lblCol{i + 1}", title, x, y + cm(0.1), w, cm(0.5), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", source, x, cm(0.08), w, cm(0.55), 9, align=2 if fmt else 0, fmt=fmt,
            grow=source == "BoxText")
        x += w
    assert x == W
    hline(m, SEC_DETAIL, "lnRow", cm(0.66), W)
    F = SEC_RPT_FOOTER
    hline(m, F, "lnNet", cm(0.1), W)
    txt(m, F, "txtNet", '=IIf(Sum(IIf([BoxNo]=16,[VAT],0))>=0,"صافي الضريبة المستحقة للسداد: ",'
        '"ضريبة مستردة تُرحَّل للإقرار التالي: ") & Format(IIf(Sum(IIf([BoxNo]=16,[VAT],0))>=0,1,-1)*Sum(IIf([BoxNo]=16,[VAT],0)),"#,##0.00")',
        0, cm(0.3), W, cm(0.7), 13, True)
    signatures(m, F, cm(1.4), ["المحاسب", "المدير"])
    page_footer(m, W)
    return m


def journal_reports() -> List[ReportModel]:
    return [journal_entry(), account_statement(), general_ledger(), income_statement(), balance_sheet(), vat_return()]
