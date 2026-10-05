"""Phase 6 reports: sales invoice / credit note, 80 mm receipt and A4.
Phase 8 adds the other documents (reports_docs.py) and the catalogue (reports_catalog.py).

Reports are grouped on DocID (header = seller/buyer/document info,
footer = totals + QR code) and sorted on LineNumber. Positions are in twips,
x measured from the start (right) edge like the forms.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from forms import Control, Sym, cm

# Access section numbers. Documents use the DocID group header/footer (5/6);
# list reports use the page header (3), report footer (2) and page footer (4).
SEC_DETAIL, SEC_RPT_HEADER, SEC_RPT_FOOTER, SEC_PAGE_HEADER, SEC_PAGE_FOOTER = 0, 1, 2, 3, 4
SEC_HEADER, SEC_FOOTER = 5, 6
SECTION_NAMES = {SEC_DETAIL: "Detail", SEC_HEADER: "secHeader", SEC_FOOTER: "secTotals"}

TITLE_AR = ('=IIf([DocKind]="RETURN","إشعار دائن",IIf([InvoiceSubType]="STANDARD",'
            '"فاتورة ضريبية","فاتورة ضريبية مبسطة"))')
TITLE_EN = ('=IIf([DocKind]="RETURN","Credit Note",IIf([InvoiceSubType]="STANDARD",'
            '"Tax Invoice","Simplified Tax Invoice"))')
MONEY = "#,##0.00"


@dataclass
class ReportModel:
    name: str
    caption: str
    width: int
    heights: Dict[int, int]
    record_source: str = "qrySalesDocPrint"
    controls: Dict[int, List[Control]] = field(default_factory=dict)
    code: List[str] = field(default_factory=list)       # report module lines
    events: List[str] = field(default_factory=list)     # e.g. "m_rpt.OnNoData = EP"
    group: str = "DocID"                                  # "" = no group header/footer
    sorts: List[Tuple[str, bool]] = field(default_factory=lambda: [("LineNumber", False)])
    landscape: bool = False
    page_setup: bool = False                              # A4 margins (list reports)
    no_data: str = ""                                     # message instead of an empty report
    prepare: List[str] = field(default_factory=list)     # procedures run first (Application.Run)

    def __post_init__(self):
        for sec in self.heights:
            self.controls.setdefault(sec, [])
        if self.no_data:
            self.events.append("m_rpt.OnNoData = EP")
            self.code += ["Private Sub Report_NoData(Cancel As Integer)",
                          f"    ReportNoData Cancel, {vba_literal(self.no_data)}", "End Sub"]

    def add(self, section, c: Control):
        self.controls[section].append(c)
        return c


def vba_literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def txt(m, sec, name, source, x, y, w, h, size=9, bold=False, align=0, fmt=None, grow=False,
        visible=True):
    props = {"FontSize": size, "FontBold": bold, "TextAlign": align}
    if fmt:
        props["Format"] = fmt
    if grow:
        props["CanGrow"] = True
    if not visible:
        props["Visible"] = False
    return m.add(sec, Control("text", name, x, y, w, h, props, source=source))


def lbl(m, sec, name, caption, x, y, w, h, size=9, bold=False, align=0):
    return m.add(sec, Control("label", name, x, y, w, h,
                              {"Caption": caption, "FontSize": size, "FontBold": bold,
                               "TextAlign": align}))


def hline(m, sec, name, y, width):
    return m.add(sec, Control("line", name, 0, y, width, 1, {}, decorative=True))


def qr_box(m: ReportModel, y: int, size: int):
    x = (m.width - size) // 2            # centred: independent of the RTL mirroring
    m.add(SEC_FOOTER, Control("rect", "boxQR", x, y, size, size, {}, decorative=True))
    txt(m, SEC_FOOTER, "txtDocKind", "DocKind", 0, y, cm(0.5), cm(0.4), visible=False)
    txt(m, SEC_FOOTER, "txtDocID", "DocID", 0, y + cm(0.45), cm(0.5), cm(0.4), visible=False)
    m.events.append("m_rpt.Section(6).OnPrint = EP")
    m.code += ["Private Sub secTotals_Print(Cancel As Integer, PrintCount As Integer)",
               "    If Me.HasData = 0 Then Exit Sub   ' no document: the fields have no value (2427)",
               "    DrawDocumentQR Me, Me!txtDocKind.Value, Me!txtDocID.Value, Me!boxQR.Left, _",
               "                   Me!boxQR.Top, Me!boxQR.Width",
               "End Sub"]


def setting(name, prefix=""):
    if prefix:
        return f'="{prefix}" & Nz(SettingValue("{name}"),"")'
    return f'=Nz(SettingValue("{name}"),"")'


def receipt() -> ReportModel:
    w = cm(7.4)
    m = ReportModel("rptSalesReceipt", "فاتورة (حراري 80 مم)", w,
                    {SEC_HEADER: cm(6.75), SEC_DETAIL: cm(0.95), SEC_FOOTER: cm(8.35)})
    H = SEC_HEADER
    rows = [
        ("txtStoreName", setting("StoreName"), 0.6, 12, True),
        ("txtStoreNameEn", setting("StoreNameEn"), 0.45, 9, False),
        ("txtStoreVat", setting("VATNumber", "الرقم الضريبي: "), 0.4, 9, True),
        ("txtStoreCR", '="س.ت: " & Nz(SettingValue("CRNumber"),"-") & "   هاتف: " & '
                       'Nz(SettingValue("Phone"),"-")', 0.4, 8, False),
        ("txtStoreAddress", '=Trim(Nz(SettingValue("City"),"") & " " & Nz(SettingValue("District"),"") '
                            '& " " & Nz(SettingValue("StreetName"),""))', 0.4, 8, False),
    ]
    y = cm(0.1)
    for name, src, h, size, bold in rows:
        txt(m, H, name, src, 0, y, w, cm(h), size, bold, align=2)
        y += cm(h)
    hline(m, H, "lnHeader1", y + cm(0.05), w)
    y += cm(0.15)
    txt(m, H, "txtTitle", TITLE_AR, 0, y, w, cm(0.55), 12, True, align=2)
    txt(m, H, "txtTitleEn", TITLE_EN, 0, y + cm(0.55), w, cm(0.4), 8, align=2)
    y += cm(1.0)
    for name, src in [("txtDocNumber", '="رقم: " & [DocNumber]'),
                      ("txtDocDate", '="التاريخ: " & Format([DocDate],"yyyy/mm/dd hh:nn")'),
                      ("txtOriginal", '=IIf(Len(Nz([OriginalNumber],""))>0,"عن الفاتورة: " & [OriginalNumber],'
                                     'OrderTypeText([OrderType],[TableNo]))'),
                      ("txtCashier", '="الكاشير: " & [EmployeeName]'),
                      ("txtCustomer", '=IIf([CustomerID]=Nz(SettingValue("DefaultCustomerID"),1),"",'
                                      '"العميل: " & [CustomerName])'),
                      ("txtCustomerVat", '=IIf(Len(Nz([CustomerVAT],""))>0,"الرقم الضريبي للعميل: " & '
                                         '[CustomerVAT],"")')]:
        txt(m, H, name, src, 0, y, w, cm(0.4), 8)
        y += cm(0.4)
    hline(m, H, "lnHeader2", y + cm(0.05), w)
    y += cm(0.1)
    lbl(m, H, "lblColItem", "الصنف", 0, y, cm(3.9), cm(0.45), 8, True)
    lbl(m, H, "lblColQty", "الكمية × السعر", cm(3.9), y, cm(2.0), cm(0.45), 8, True)
    lbl(m, H, "lblColTotal", "الإجمالي", cm(5.9), y, cm(1.5), cm(0.45), 8, True, align=1)
    hline(m, H, "lnHeader3", y + cm(0.5), w)
    assert y + cm(0.55) <= m.heights[H]

    D = SEC_DETAIL
    txt(m, D, "txtProduct", "ProductName", 0, 0, w, cm(0.45), 9, grow=True)
    txt(m, D, "txtQtyPrice", '=[Quantity] & " × " & Format([LineTotal]/[Quantity],"#,##0.00")',
        cm(0.2), cm(0.47), cm(5.0), cm(0.42), 8)
    txt(m, D, "txtLineTotal", "LineTotal", cm(5.2), cm(0.47), cm(2.2), cm(0.42), 9, True, align=1,
        fmt=MONEY)

    F = SEC_FOOTER
    hline(m, F, "lnTotals1", cm(0.08), w)
    y = cm(0.2)
    for cap, src, bold, size in [
            ("الإجمالي بدون الضريبة", "TaxableAmount", False, 9),
            ("الخصم", "DocDiscount", False, 9),
            ('="ضريبة القيمة المضافة " & Format(Nz(SettingValue("VATRate"),0.15),"0%")', "DocTax", False, 9),
            ("الإجمالي شامل الضريبة", "TotalAmount", True, 11)]:
        n = src
        if cap.startswith("="):
            txt(m, F, f"txtCap{n}", cap, 0, y, cm(4.6), cm(0.45), size, bold)
        else:
            lbl(m, F, f"lblCap{n}", cap, 0, y, cm(4.6), cm(0.45), size, bold)
        txt(m, F, f"txt{n}", src, cm(4.7), y, cm(2.7), cm(0.45), size, bold, align=1, fmt=MONEY)
        y += cm(0.5)
    hline(m, F, "lnTotals2", y + cm(0.05), w)
    y += cm(0.15)
    for name, cap, val in [
            ("Paid", '=IIf([DocKind]="RETURN","المبلغ المردود","المدفوع")', "PaidAmount"),
            ("Change", '=IIf([ChangeDue]>0,"الباقي للعميل","")', '=IIf([ChangeDue]>0,Format([ChangeDue],"#,##0.00"),"")'),
            ("Remaining", '=IIf([RemainingAmount]<>0,IIf([DocKind]="RETURN","خصم من رصيد العميل",'
                          '"المتبقي على الحساب"),"")',
             '=IIf([RemainingAmount]<>0,Format([RemainingAmount],"#,##0.00"),"")')]:
        txt(m, F, f"txtCap{name}", cap, 0, y, cm(4.6), cm(0.45), 9)
        txt(m, F, f"txtVal{name}", val, cm(4.7), y, cm(2.7), cm(0.45), 9, align=1,
            fmt=MONEY if not val.startswith("=") else None)
        y += cm(0.45)
    qr_box(m, y + cm(0.15), cm(3.6))
    txt(m, F, "txtFooterNote", setting("ReceiptFooter"), 0, y + cm(3.85), w, cm(0.45), 9, align=2)
    assert y + cm(4.3) <= m.heights[F]
    return m


A4_COLUMNS = [("#", "LineNumber", 0.8, None), ("الصنف", "ProductName", 6.2, None),
              ("الكمية", "Quantity", 1.6, "#,##0.###"), ("سعر الوحدة", "UnitPrice", 2.0, MONEY),
              ("الخصم", "LineDiscount", 1.6, MONEY), ("الصافي", "NetAmount", 2.0, MONEY),
              ("نسبة الضريبة", "VATRate", 1.2, "0%"), ("الضريبة", "LineTax", 1.6, MONEY),
              ("الإجمالي", "LineTotal", 2.0, MONEY)]


def a4() -> ReportModel:
    w = cm(19.0)
    m = ReportModel("rptSalesInvoiceA4", "فاتورة ضريبية (A4)", w,
                    {SEC_HEADER: cm(6.8), SEC_DETAIL: cm(0.6), SEC_FOOTER: cm(8.2)})
    H = SEC_HEADER
    half = cm(9.3)
    y = cm(0.1)
    for name, src, size, bold in [("txtStoreName", setting("StoreName"), 14, True),
                                  ("txtStoreNameEn", setting("StoreNameEn"), 10, False),
                                  ("txtStoreVat", setting("VATNumber", "الرقم الضريبي: "), 10, True),
                                  ("txtStoreCR", setting("CRNumber", "السجل التجاري: "), 9, False),
                                  ("txtStoreAddress", '=Trim(Nz(SettingValue("BuildingNo"),"") & " " & '
                                   'Nz(SettingValue("StreetName"),"") & " - " & Nz(SettingValue("District"),"") '
                                   '& " - " & Nz(SettingValue("City"),"") & " " & Nz(SettingValue("PostalCode"),""))',
                                   9, False),
                                  ("txtStorePhone", setting("Phone", "هاتف: "), 9, False)]:
        h = cm(0.7) if size >= 14 else cm(0.5)
        txt(m, H, name, src, 0, y, half, h, size, bold)
        y += h
    x2 = w - half
    txt(m, H, "txtTitle", TITLE_AR, x2, cm(0.1), half, cm(0.8), 16, True, align=1)
    txt(m, H, "txtTitleEn", TITLE_EN, x2, cm(0.95), half, cm(0.5), 10, align=1)
    yy = cm(1.5)
    for name, src in [("txtDocNumber", '="رقم الفاتورة: " & [DocNumber]'),
                      ("txtDocDate", '="التاريخ: " & Format([DocDate],"yyyy/mm/dd hh:nn")'),
                      ("txtOriginal", '=IIf(Len(Nz([OriginalNumber],""))>0,"عن الفاتورة: " & [OriginalNumber],'
                                     'OrderTypeText([OrderType],[TableNo]))'),
                      ("txtPaymentType", '="طريقة البيع: " & IIf([PaymentType]="CREDIT","آجل","نقدي")'),
                      ("txtCashier", '="الموظف: " & [EmployeeName]')]:
        txt(m, H, name, src, x2, yy, half, cm(0.45), 9, align=1)
        yy += cm(0.45)
    by = cm(3.85)
    m.add(H, Control("rect", "boxBuyer", 0, by, w, cm(1.95), {"BorderStyle": 1}, decorative=True))
    lbl(m, H, "lblBuyer", "بيانات المشتري", cm(0.2), by + cm(0.05), cm(4), cm(0.45), 9, True)
    txt(m, H, "txtBuyerName", '="الاسم: " & [CustomerName]', cm(0.2), by + cm(0.5), cm(9), cm(0.45), 9)
    txt(m, H, "txtBuyerVat", '=IIf(Len(Nz([CustomerVAT],""))>0,"الرقم الضريبي: " & [CustomerVAT],"")',
        cm(9.6), by + cm(0.5), cm(9.2), cm(0.45), 9)
    txt(m, H, "txtBuyerAddress", '=Trim(Nz([CustomerBuilding],"") & " " & Nz([CustomerStreet],"") & " " & '
                                 'Nz([CustomerDistrict],"") & " " & Nz([CustomerCity],"") & " " & '
                                 'Nz([CustomerPostal],""))', cm(0.2), by + cm(1.0), w - cm(0.4), cm(0.45), 9)
    hy = cm(6.05)
    m.add(H, Control("rect", "boxColumns", 0, hy, w, cm(0.65), {"BackStyle": 1, "BackColor": Sym("CLR_SECONDARY")},
                     decorative=True))
    x = 0
    for i, (title, fld, width, fmt) in enumerate(A4_COLUMNS):
        cw = cm(width) if i < len(A4_COLUMNS) - 1 else w - x      # last column absorbs rounding
        lbl(m, H, f"lblCol{i + 1}", title, x, hy + cm(0.1), cw, cm(0.45), 8, True, align=2)
        txt(m, SEC_DETAIL, f"txtCol{i + 1}", fld, x, cm(0.05), cw, cm(0.5), 8,
            align=0 if fld == "ProductName" else 2, fmt=fmt, grow=fld == "ProductName")
        x += cw
    assert x == w

    F = SEC_FOOTER
    hline(m, F, "lnTotals1", cm(0.08), w)
    tx, lw, vw = cm(11.4), cm(4.4), cm(3.2)
    y = cm(0.2)
    for cap, src, bold in [("الإجمالي قبل الخصم", "DocSubTotal", False), ("الخصم", "DocDiscount", False),
                           ("الإجمالي الخاضع للضريبة", "TaxableAmount", False),
                           ("ضريبة القيمة المضافة", "DocTax", False),
                           ("الإجمالي شامل الضريبة", "TotalAmount", True),
                           ("المدفوع / المردود", "PaidAmount", False),
                           ("المتبقي", "RemainingAmount", False)]:
        lbl(m, F, f"lblCap{src}", cap, tx, y, lw, cm(0.45), 10 if bold else 9, bold)
        txt(m, F, f"txt{src}", src, tx + lw, y, vw, cm(0.45), 10 if bold else 9, bold, align=1, fmt=MONEY)
        y += cm(0.5)
    qr_box(m, cm(3.9), cm(3.6))
    txt(m, F, "txtFooterNote", setting("ReceiptFooter"), 0, cm(7.6), w, cm(0.45), 9, align=2)
    return m


def sales_reports() -> List[ReportModel]:
    return [receipt(), a4()]


def all_reports() -> List[ReportModel]:
    from reports_docs import document_reports
    from reports_catalog import catalog_reports
    return sales_reports() + document_reports() + catalog_reports()
