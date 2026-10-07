"""The accounting hub (frmAccounting): one tile per accounting screen, so none of them is hidden
three screens deep. Tiles the user may not open are greyed by ApplyNavPermissions (modSecurityScreens),
like the dashboard tiles. Layout conventions are the same as forms.py."""

from typing import List

from forms import ICONS, Control, FormModel, Sym, button, cm, fit_window, title_band

# key, caption, hint, screen, OpenScreen arguments after the screen, icon, colour
HUB_TILES = [
    ("Journal", "قيود اليومية", "كل القيود وميزان المراجعة", "frmJournal", "", "journal", (31, 58, 95)),
    ("Accounts", "دليل الحسابات", "شجرة الحسابات", "frmAccounts", "", "category", (57, 73, 171)),
    ("Manual", "قيد يدوي", "قيود التسوية والافتتاح", "frmManualEntry", "", "journal", (94, 53, 177)),
    ("Ledger", "كشف حساب", "حركة حساب ودفتر الأستاذ", "frmLedger", "", "reports", (30, 136, 229)),
    ("Financials", "الحسابات الختامية", "قائمة الدخل والميزانية", "frmFinancials", "", "reports", (0, 121, 107)),
    ("Closing", "إقفال الفترات", "إقفال الفترة والسنة المالية", "frmPeriodClosing", ", 0", "settings", (84, 110, 122)),
    ("Vat", "الإقرار الضريبي", "ضريبة القيمة المضافة", "frmVatReturn", ", 0", "reports", (216, 27, 96)),
    ("Aging", "أعمار الديون", "العملاء والموردون حسب الاستحقاق", "frmAging", ", 0", "customers", (229, 57, 53)),
    ("Banks", "البنوك", "الحسابات البنكية والتسوية", "frmBanks", "", "treasury", (0, 137, 123)),
    ("Cheques", "الشيكات", "الواردة والصادرة", "frmCheques", ', 0, "IN"', "treasury", (67, 160, 71)),
    ("Assets", "الأصول الثابتة", "الشراء والبيع والاستبعاد", "frmAssets", ", 0", "inventory", (251, 140, 0)),
    ("Depreciation", "الإهلاك الشهري", "تسجيل إهلاك الأصول", "frmDepreciation", ", 0", "inventory", (239, 108, 0)),
    ("Payroll", "الرواتب", "مسير الرواتب والتأمينات", "frmPayroll", ", 0", "users", (142, 36, 170)),
    ("Centers", "مراكز التكلفة", "الفروع والأقسام", "frmCostCenters", "", "category", (0, 131, 143)),
    ("Budget", "الموازنة", "الموازنة مقابل الفعلي", "frmBudget", "", "reports", (121, 85, 72)),
    ("Recurring", "المصروفات المتكررة", "الإيجار والكهرباء والاشتراكات", "frmRecurring", "", "expenses", (198, 40, 40)),
]


def layout_accounting() -> FormModel:
    cols, tile_w, tile_h, gap = 4, cm(6.0), cm(2.3), cm(0.35)
    width = cm(0.6) * 2 + cols * tile_w + (cols - 1) * gap
    top = cm(2.0)
    rows = (len(HUB_TILES) + cols - 1) // cols
    height = top + rows * (tile_h + gap) + cm(1.3)
    m = FormModel("frmAccounting", "المحاسبة والمالية", width, height, popup=False, allow_add=False,
                  allow_edit=False)
    title_band(m, "المحاسبة والمالية", "كل شاشات الحسابات في مكان واحد", "journal")
    for i, (key, caption, hint, screen, args, icon, rgb) in enumerate(HUB_TILES):
        # first tile on the right (Arabic reading order), as on the dashboard
        x = cm(0.6) + (cols - 1 - i % cols) * (tile_w + gap)
        y = top + (i // cols) * (tile_h + gap)
        m.add(Control("rect", f"boxNav{key}", x, y, tile_w, tile_h,
                      {"BackColor": Sym(f"RGB({rgb[0]}, {rgb[1]}, {rgb[2]})")}, decorative=True))
        m.add(Control("icon", f"icoTile{key}", x, y + cm(0.2), tile_w, cm(0.9),
                      {"Caption": Sym(f"ChrW(&H{ICONS[icon]:X})"), "FontSize": 20, "TextAlign": 2,
                       "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
        m.add(Control("label", f"lblTile{key}", x, y + cm(1.1), tile_w, cm(0.6),
                      {"Caption": caption, "FontSize": 12, "FontBold": True, "TextAlign": 2,
                       "ForeColor": Sym("CLR_SURFACE")}, decorative=True))
        m.add(Control("label", f"lblHint{key}", x, y + cm(1.65), tile_w, cm(0.5),
                      {"Caption": hint, "FontSize": 8, "TextAlign": 2, "ForeColor": Sym("CLR_SURFACE")},
                      decorative=True))
        button(m, f"btnTile{key}", caption, x, y, "secondary", w=tile_w, h=tile_h,
               call=f'OpenScreen "{screen}"{args}')
        m.controls[-1].props["Transparent"] = True       # the coloured tile under it shows through
        m.controls[-1].props["Tag"] = screen             # ApplyNavPermissions greys what the user may not open
    button(m, "btnClose", "رجوع", width - cm(0.6) - cm(2.4), height - cm(1.1), "secondary", w=cm(2.4), h=cm(0.8),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    ApplyNavPermissions Me", "End Sub"] + m.code
    fit_window(m, split_x=width - cm(3.1), bottom_y=height - cm(1.2))     # the title band and the close button
    return m


def accounting_forms() -> List[FormModel]:
    return [layout_accounting()]
