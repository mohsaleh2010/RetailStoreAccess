"""The start-up opening balances (docs/52-Opening-Balances.md): frmOpeningBalances with the start-up date
(Settings.OpeningDate), the balance of the opening balances account 3900 by source, and one grid per part:
the general accounts (OpeningBalances), the customers, the suppliers, the cash boxes and the banks
(their OpeningBalance field). The logic is in modOpening; layout conventions are the same as forms.py."""

from typing import List, Tuple

from forms import Control, FormModel, Sym, button, cm, fit_window, labelled, title_band
from forms_sales import LOCKED, grid_row

ROW_H = cm(0.75)
GRID_W = cm(25.4)
MONEY = {"Format": "#,##0.00"}
# the general accounts: posting accounts that have no opening balance of their own elsewhere
# (customers 1300, suppliers 2100, stock 1400, the opening balances account 3900, cash boxes, banks)
OPENING_ACCOUNTS = ("SELECT a.AccountCode, a.AccountCode & '  ' & a.AccountName AS Account FROM [@Accounts] AS a "
                    "WHERE IsPosting = True AND IsActive = True AND AccountCode NOT IN (1300, 2100, 1400, 3900) "
                    "AND AccountCode NOT BETWEEN 110001 AND 119999 AND AccountCode NOT BETWEEN 120001 AND 129999 "
                    "AND AccountCode NOT BETWEEN 140001 AND 149999 "
                    "ORDER BY TreeKey")
# part key, grid form, table, key field, the names shown (a locked combo on the key, so an English front-end
# shows the English name), last column, record source (the table itself: the grid edits it), sign text
PARTY_GRIDS = [
    ("CUSTOMERS", "frmOpeningCustomers", "Customers", "CustomerID",
     "SELECT c.CustomerID, c.CustomerName FROM [@Customers] AS c", "CurrentBalance",
     "SELECT CustomerID, OpeningBalance, CurrentBalance, CreatedAt FROM Customers WHERE IsSystem = False "
     "ORDER BY CustomerID",
     '=IIf([OpeningBalance]>0,"عليه (مدين)",IIf([OpeningBalance]<0,"له (دائن)",""))'),
    ("SUPPLIERS", "frmOpeningSuppliers", "Suppliers", "SupplierID",
     "SELECT s.SupplierID, s.SupplierName FROM [@Suppliers] AS s", "CurrentBalance",
     "SELECT SupplierID, OpeningBalance, CurrentBalance, CreatedAt FROM Suppliers ORDER BY SupplierID",
     '=IIf([OpeningBalance]>0,"له (دائن)",IIf([OpeningBalance]<0,"عليه (مدين)",""))'),
    ("BOXES", "frmOpeningBoxes", "CashBoxes", "CashBoxID",
     "SELECT b.CashBoxID, b.BoxName FROM [@CashBoxes] AS b", "OpeningDate",
     "SELECT CashBoxID, OpeningBalance, OpeningDate FROM CashBoxes ORDER BY CashBoxID",
     '=IIf([OpeningBalance]>0,"مدين",IIf([OpeningBalance]<0,"دائن",""))'),
    ("BANKS", "frmOpeningBanks", "Banks", "BankID",
     "SELECT k.BankID, k.BankName FROM [@Banks] AS k", "OpeningDate",
     "SELECT BankID, OpeningBalance, OpeningDate FROM Banks ORDER BY BankID",
     '=IIf([OpeningBalance]>0,"مدين (لنا)",IIf([OpeningBalance]<0,"دائن (سحب على المكشوف)",""))'),
]
PARTS = [("ACCOUNTS", "الحسابات العامة"), ("CUSTOMERS", "العملاء"), ("SUPPLIERS", "الموردون"),
         ("BOXES", "الصناديق"), ("BANKS", "البنوك"), ("STOCK", "المخزون")]
# the stock part (docs/53): the products that track stock, by code and name
STOCK_PRODUCTS = ("SELECT p.ProductID, p.ProductCode & '  ' & p.ProductName AS Item FROM Products AS p "
                  "WHERE p.TrackStock = True ORDER BY p.ProductName")
# the warehouse of each line (docs/57): empty = the main warehouse
STOCK_WAREHOUSES = ("SELECT w.WarehouseID, w.WarehouseName FROM [@Warehouses] AS w WHERE w.IsActive = True "
                    "ORDER BY w.WarehouseID")
# its buttons, shown with the stock part only: name, caption, call
STOCK_BUTTONS = [("btnStockFill", "كل الأصناف", "OpeningStockFill Me"),
                 ("btnStockImport", "استيراد من Excel", "OpeningStockImport Me"),
                 ("btnStockTemplate", "قالب Excel", "OpeningStockTemplate Me"),
                 ("btnStockPost", "ترحيل الرصيد", "OpeningStockPost Me")]


def layout_account_lines() -> Tuple[FormModel, list]:
    m = FormModel("frmOpeningAccountLines", "الأرصدة الافتتاحية للحسابات", GRID_W, ROW_H, popup=False,
                  record_source="SELECT * FROM OpeningBalances ORDER BY OpeningLineID", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    heads = grid_row(m, [
        ("AccountCode", "AccountCode", 10.0, {"_kind": "combo", "RowSource": OPENING_ACCOUNTS, "ColumnCount": 2,
                                              "ColumnWidths": "0;9.8", "BoundColumn": 1, "LimitToList": True}, []),
        ("Debit", "Debit", 3.2, dict(MONEY), []),
        ("Credit", "Credit", 3.2, dict(MONEY), []),
        ("LineText", "LineText", 7.4, {}, []),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], ROW_H)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)", "    OpeningLineBeforeUpdate Me, Cancel", "End Sub",
              "Private Sub Form_AfterUpdate()", "    OpeningGridAfterUpdate Me", "End Sub",
              "Private Sub btnRemove_Click()", "    OpeningRemoveLine Me", "End Sub"]
    return m, heads


def layout_stock_grid() -> FormModel:
    m = FormModel("frmOpeningStock", "الرصيد الافتتاحي للمخزون", GRID_W, ROW_H, popup=False,
                  record_source="SELECT * FROM OpeningStock ORDER BY OpeningStockID", allow_add=True)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    grid_row(m, [
        ("ProductID", "ProductID", 6.9, {"_kind": "combo", "RowSource": STOCK_PRODUCTS, "ColumnCount": 2,
                                         "ColumnWidths": "0;6.7", "BoundColumn": 1, "LimitToList": True}, []),
        ("WarehouseID", "WarehouseID", 3.05, {"_kind": "combo", "RowSource": STOCK_WAREHOUSES, "ColumnCount": 2,
                                              "ColumnWidths": "0;2.9", "BoundColumn": 1, "LimitToList": True,
                                              "DefaultValue": "1"}, []),
        ("Quantity", "Quantity", 3.2, {"Format": "#,##0.###"}, []),
        ("UnitCost", "UnitCost", 3.2, dict(MONEY), []),
        ("txtValue", "=Nz([Quantity],0)*Nz([UnitCost],0)", 3.6, {**LOCKED, **MONEY}, []),
        ("txtStatus", '=IIf(IsNull([TransactionID]),"لم يُرحَّل","مرحّل")', 3.75, dict(LOCKED), []),
        ("btnRemove", "", 0.8, {"_kind": "button", "Caption": Sym(f"ChrW(&H{0xE74D:X})"),
                               "Style": "danger", "FontName": Sym("ICON_FONT")}, ["Click"]),
    ], ROW_H)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)", "    OpeningStockBeforeUpdate Me, Cancel", "End Sub",
              "Private Sub Form_AfterUpdate()", "    OpeningGridAfterUpdate Me", "End Sub",
              "Private Sub btnRemove_Click()", "    OpeningStockRemoveLine Me", "End Sub"]
    return m


def layout_party_grid(part, name, tbl, key, names, last, sql, side) -> FormModel:
    m = FormModel(name, "الأرصدة الافتتاحية", GRID_W, ROW_H, popup=False, record_source=sql, allow_add=False)
    m.form_props = {"DefaultView": 1, "ScrollBars": 2, "Cycle": 0, "AllowDeletions": False}
    last_props = {**LOCKED, **(MONEY if last == "CurrentBalance" else {"Format": "yyyy/mm/dd"})}
    grid_row(m, [
        ("cboName", key, 10.0, {**LOCKED, "_kind": "combo", "RowSource": names, "ColumnCount": 2,
                                "ColumnWidths": "0;9.8", "BoundColumn": 1}, []),
        ("OpeningBalance", "OpeningBalance", 3.2, dict(MONEY), []),
        ("txtSide", side, 3.2, dict(LOCKED), []),
        (last, last, 7.4, last_props, []),
    ], ROW_H)
    m.form_events = ["BeforeUpdate", "AfterUpdate"]
    m.code = ["Private Sub Form_BeforeUpdate(Cancel As Integer)",
              f'    OpeningPartyBeforeUpdate Me, Cancel, "{tbl}", "{key}"', "End Sub",
              "Private Sub Form_AfterUpdate()", "    OpeningGridAfterUpdate Me", "End Sub"]
    return m


def layout_opening(heads) -> FormModel:
    width, height = cm(26.2), cm(17.0)
    m = FormModel("frmOpeningBalances", "أرصدة بداية التشغيل", width, height, popup=False, allow_add=False)
    title_band(m, "أرصدة بداية التشغيل",
               "رصيد كل حساب وعميل ومورد وصندوق وبنك وصنف يوم بدء التشغيل، ومقابلها حساب إجمالي الأرصدة الافتتاحية",
               "journal")
    y = cm(2.3)
    c = m.add(Control("text", "txtOpeningDate", cm(0.4), y, cm(3.2), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtOpeningDate", "تاريخ بدء التشغيل", c)
    button(m, "btnSaveDate", "حفظ التاريخ", cm(3.8), y, "primary", w=cm(2.8), h=cm(0.8), call="OpeningSaveDate Me")
    m.add(Control("label", "lblDateHint", cm(6.8), y, cm(18.2), cm(0.8),
                  {"Caption": "كل الأرصدة الافتتاحية بهذا التاريخ. لا يكون بعد أول مستند، ولا في فترة مقفلة.",
                   "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("text", "txtPart", cm(25.2), y, cm(0.5), cm(0.8), {"Visible": False}))
    # the opening balances account 3900 and its sources
    m.add(Control("label", "lblTotal", cm(0.4), cm(3.4), cm(14.2), cm(1.6),
                  {"Caption": " ", "FontSize": 12, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    button(m, "btnRefresh", "تحديث الإجمالي", cm(0.4), cm(5.1), "secondary", w=cm(3.4), h=cm(0.8),
           call="OpeningRefresh Me")
    m.add(Control("list", "lstSources", cm(14.8), cm(3.4), cm(11.0), cm(2.5),
                  {"RowSourceType": "Table/Query", "RowSource": "", "ColumnCount": 3,
                   "ColumnWidths": "5.4;2.6;2.6", "ColumnHeads": True}))
    # the parts
    y, x = cm(6.3), cm(0.4)
    for key, caption in PARTS:
        button(m, f"btnPart{key.title()}", caption, x, y, "secondary", w=cm(3.4), h=cm(0.8),
               call=f'OpeningShowPart Me, "{key}"')
        x += cm(3.6)
    x = cm(0.4)
    for name, caption, call in STOCK_BUTTONS:          # the stock part (OpeningShowPart shows them)
        button(m, name, caption, x, cm(15.6), "secondary", w=cm(3.4), h=cm(0.9), call=call)
        m.controls[-1].props["Visible"] = False
        x += cm(3.6)
    m.add(Control("label", "lblHint", cm(0.4), cm(7.25), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    for i, (hx, hw) in enumerate(heads[:4]):
        m.add(Control("label", f"lblCol{i + 1}", cm(0.4) + hx, cm(7.95), hw, cm(0.55),
                      {"Caption": " ", "FontSize": 9, "FontBold": True, "ForeColor": Sym("CLR_MUTED"),
                       "TextAlign": 2}))
    m.add(Control("subform", "subGrid", cm(0.4), cm(8.55), GRID_W, cm(6.6),
                  {"SourceObject": "frmOpeningAccountLines"}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.6), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    OpeningLoad Me", "End Sub"] + m.code
    fit_window(m, bottom_y=cm(15.5), stretch_h=("subGrid",))     # a taller window shows more rows
    return m


def opening_forms() -> List[FormModel]:
    lines, heads = layout_account_lines()
    return [lines] + [layout_party_grid(*g) for g in PARTY_GRIDS] + [layout_stock_grid(), layout_opening(heads)]
