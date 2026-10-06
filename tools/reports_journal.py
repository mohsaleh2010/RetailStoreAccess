"""Journal entry print (A4, grouped on the entry): rptJournalEntry (qryJournalEntryPrint).
The journal and the trial balance are list reports of reports_catalog.py."""

from typing import List

from reports import ReportModel, MONEY, SEC_DETAIL, SEC_HEADER, SEC_FOOTER, txt, hline
from reports_docs import W, columns, signatures, store_block
from forms import cm


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


def journal_reports() -> List[ReportModel]:
    return [journal_entry()]
