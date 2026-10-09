"""The e-invoicing screen (frmEInvoices, docs/45-EInvoice-Foundation.md): the sales invoices and returns with
their e-invoicing status, the requests of the chosen document, and sending again. The logic is in modEInvoice."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_security import PASSWORD, check_with_label

ENVIRONMENTS = "TEST;تجريبية;SIMULATION;محاكاة (السعودية);PRODUCTION;فعلية"
STATUS_FILTER = ("ATTENTION;تحتاج متابعة (بانتظار الإرسال، مرفوضة، تحذير);SENT;أُرسلت;"
                 "NOT_SENT;قبل التفعيل;ALL;الكل")


def layout_einvoices() -> FormModel:
    width, height = cm(27.0), cm(18.4)
    m = FormModel("frmEInvoices", "الفاتورة الإلكترونية", width, height, popup=True, allow_add=False)
    title_band(m, "الفاتورة الإلكترونية", "حالة إرسال فواتير البيع والمرتجعات للمنظومة، وإعادة الإرسال", "sales")
    y = cm(2.3)
    c = m.add(Control("text", "txtFrom", cm(0.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtFrom", "من", c)
    c = m.add(Control("text", "txtTo", cm(3.4), y, cm(2.8), cm(0.8), {"Format": "yyyy/mm/dd"}))
    labelled(m, "txtTo", "إلى", c)
    c = m.add(Control("combo", "cboStatus", cm(6.4), y, cm(8.0), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": STATUS_FILTER, "ColumnCount": 2,
                       "ColumnWidths": "0;7.8", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboStatus", "المستندات", c)
    button(m, "btnShow", "عرض", cm(14.6), y, "primary", w=cm(3.0), h=cm(0.8), call="EInvoicesShow Me")
    # the e-invoicing settings (Settings.EInvoiceEnabled / EInvoiceEnvironment), saved at once (modEInvoice)
    c = m.add(Control("combo", "cboEnvironment", cm(18.0), y, cm(4.4), cm(0.8),
                      {"RowSourceType": "Value List", "RowSource": ENVIRONMENTS, "ColumnCount": 2,
                       "ColumnWidths": "0;4.2", "LimitToList": True}, events=["AfterUpdate"]))
    labelled(m, "cboEnvironment", "البيئة", c)
    check_with_label(m, "chkEnabled", "تفعيل الإرسال", cm(22.8), y, cm(3.8), events=["AfterUpdate"])
    m.add(Control("label", "lblSummary", cm(0.4), cm(3.3), width - cm(0.8), cm(0.55),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstDocs", cm(0.4), cm(4.0), width - cm(0.8), cm(7.4),
                  {"ColumnCount": 9, "ColumnWidths": "0;2.8;3.2;2.6;4.4;2.6;3.2;1.8;5.6", "ColumnHeads": True},
                  events=["AfterUpdate"]))
    m.add(Control("label", "lblLogTitle", cm(0.4), cm(11.5), width - cm(0.8), cm(0.55),
                  {"Caption": "طلبات المستند المختار وردود المنظومة", "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    m.add(Control("list", "lstLog", cm(0.4), cm(12.1), width - cm(0.8), cm(4.0),
                  {"ColumnCount": 6, "ColumnWidths": "0;3.6;2.4;2.4;1.8;15.4", "ColumnHeads": True}))
    y = cm(16.6)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSendPicked", "إرسال المختار", "primary", 3.6, "EInvoicesSendPicked Me"),
            ("btnSendAll", "إرسال كل المعلّق", "secondary", 3.8, "EInvoicesSendAll Me"),
            ("btnRefresh", "تحديث الحالة", "secondary", 3.2, "EInvoicesRefresh Me"),
            ("btnCancelDoc", "إلغاء المستند", "secondary", 3.2, "EInvoicesCancelPicked Me"),
            ("btnSetup", "إعداد الربط", "secondary", 3.2, "EInvoiceSetupOpen")]:
        button(m, name, caption, bx, y, style, w=cm(w), h=cm(0.9), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    EInvoicesLoad Me", "End Sub",
               "Private Sub cboStatus_AfterUpdate()", "    EInvoicesShow Me", "End Sub",
               "Private Sub lstDocs_AfterUpdate()", "    EInvoicesPick Me", "End Sub",
               "Private Sub cboEnvironment_AfterUpdate()", "    EInvoiceSettingChanged Me", "End Sub",
               "Private Sub chkEnabled_AfterUpdate()", "    EInvoiceSettingChanged Me", "End Sub"] + m.code)
    return m


def layout_zatca_setup() -> FormModel:
    """frmZatcaSetup (docs/46, docs/47): OpenSSL, the private key file and the certificate of the device (modZatcaXml),
    and the three steps of the onboarding of the device (modZatcaApi)."""
    width, height = cm(22.0), cm(16.8)
    m = FormModel("frmZatcaSetup", "إعداد ربط منصة فاتورة", width, height, popup=True, allow_add=False)
    title_band(m, "إعداد ربط منصة فاتورة", "مفتاح الجهاز وشهادته وبرنامج التوقيع (السعودية)", "settings")
    y = cm(2.4)
    c = m.add(Control("text", "txtOpenSsl", cm(0.4), y, width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtOpenSsl", "مسار برنامج OpenSSL (فارغ = openssl من مسار النظام)", c)
    y = cm(3.9)
    c = m.add(Control("text", "txtKeyFile", cm(0.4), y, width - cm(4.2), cm(0.8), {}))
    labelled(m, "txtKeyFile", "ملف المفتاح الخاص للجهاز (يبقى على جهاز آمن)", c)
    button(m, "btnBrowseKey", "استعراض", width - cm(3.6), y, "secondary", w=cm(3.2), h=cm(0.8),
           call="ZatcaSetupBrowseKey Me")
    y = cm(5.4)
    c = m.add(Control("text", "txtCertificate", cm(0.4), y, width - cm(0.8), cm(2.2),
                      {"EnterKeyBehavior": True, "ScrollBars": 2, "FontSize": 8, "TextAlign": 1},
                      events=["AfterUpdate"]))
    labelled(m, "txtCertificate", "شهادة الجهاز (CSID): تُملأ من الخطوة 3، أو تُلصق كما تصدرها الهيئة", c)
    m.add(Control("label", "lblCertInfo", cm(0.4), cm(7.7), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_PRIMARY")}))
    # device onboarding (docs/47): branch and activity in the certificate, the OTP of the Fatoora portal, 3 steps
    m.add(Control("label", "lblOnboard", cm(0.4), cm(8.5), width - cm(0.8), cm(0.6),
                  {"Caption": "تسجيل الجهاز لدى الهيئة", "FontSize": 11, "FontWeight": 700,
                   "ForeColor": Sym("CLR_PRIMARY")}))
    half = (width - cm(1.0)) / 2
    c = m.add(Control("text", "txtBranch", cm(0.4), cm(9.8), half, cm(0.8), {}))
    labelled(m, "txtBranch", "اسم الفرع", c)
    c = m.add(Control("text", "txtIndustry", cm(0.6) + half, cm(9.8), half, cm(0.8), {}))
    labelled(m, "txtIndustry", "نشاط المنشأة (مثل Retail)", c)
    y = cm(11.3)
    c = m.add(Control("text", "txtOtp", cm(0.4), y, cm(4.0), cm(0.9), {}))
    labelled(m, "txtOtp", "رمز التحقق OTP", c)
    bx = cm(4.6)
    for name, caption, w, step in [("btnStep1", "1. شهادة الامتثال", 4.6, 1),
                                   ("btnStep2", "2. فحوص الامتثال", 4.6, 2),
                                   ("btnStep3", "3. الشهادة الفعلية", 4.6, 3)]:
        button(m, name, caption, bx, y, "secondary", w=cm(w), h=cm(0.9), call=f"ZatcaOnboardStep Me, {step}")
        bx += cm(w) + cm(0.2)
    m.add(Control("label", "lblStage", cm(0.4), cm(12.4), width - cm(0.8), cm(0.9),
                  {"Caption": " ", "FontSize": 10, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblSetupNote", cm(0.4), cm(13.4), width - cm(0.8), cm(1.6),
                  {"Caption": "رمز التحقق من بوابة فاتورة (في البيئة التجريبية أي رقم مثل 123345). الخطوة 1 تنشئ "
                              "مفتاحًا جديدًا في مجلد ZATCA بجانب البرنامج. بعد الخطوة 3 فعّل الإرسال من شاشة الفاتورة "
                              "الإلكترونية. للتجربة بلا هيئة: «شهادة تجريبية» ثم «ملف XML لفاتورة».",
                   "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    y = cm(15.2)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSave", "حفظ", "primary", 3.0, "ZatcaSetupSave Me"),
            ("btnTestCert", "شهادة تجريبية", "secondary", 3.4, "ZatcaSetupTestCertificate Me"),
            ("btnExportXml", "ملف XML لفاتورة", "secondary", 3.6, "ZatcaSetupExportXml Me")]:
        button(m, name, caption, bx, y, style, w=cm(w), h=cm(0.9), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    ZatcaSetupLoad Me", "    ZatcaOnboardShow Me", "End Sub",
               "Private Sub txtCertificate_AfterUpdate()", "    ZatcaSetupShowCert Me", "End Sub"] + m.code)
    return m


def layout_eta_setup() -> FormModel:
    """frmEtaSetup (docs/48): the POS of the Egyptian e-receipt: its credentials and the data of the seller
    (modEtaReceipt)."""
    width, height = cm(20.0), cm(18.6)
    m = FormModel("frmEtaSetup", "إعداد ربط منظومة الإيصال الإلكتروني", width, height, popup=True, allow_add=False)
    title_band(m, "إعداد ربط منظومة الإيصال الإلكتروني", "جهاز نقطة البيع المسجّل في بوابة مصلحة الضرائب (مصر)",
               "settings")
    half = (width - cm(1.0)) / 2
    rows = [("txtClientId", "Client ID لجهاز نقطة البيع", {}), ("txtClientSecret", "Client Secret", PASSWORD),
            ("txtPosSerial", "الرقم التسلسلي للجهاز (POS Serial)", {}), ("txtPosOs", "نظام تشغيل الجهاز (فارغ = Windows)", {}),
            ("txtPreSharedKey", "المفتاح المشترك (إن أعطته المصلحة)", PASSWORD), ("txtBranchCode", "كود الفرع (فارغ = 0)", {}),
            ("txtActivityCode", "كود النشاط (4 أرقام)", {}), ("txtGovernate", "المحافظة", {}),
            # the e-invoice to businesses (docs/49): the program (ERP) and the signing program of the USB token
            ("txtErpClientId", "الفاتورة الإلكترونية: Client ID للبرنامج", {}),
            ("txtErpClientSecret", "الفاتورة الإلكترونية: Client Secret", PASSWORD),
            ("txtSignerPath", "برنامج التوقيع (فارغ = بلا توقيع، للتجربة فقط)", {}),
            ("txtTokenPin", "الرقم السري لفلاشة التوقيع", PASSWORD)]
    for i, (name, caption, props) in enumerate(rows):
        x = cm(0.4) if i % 2 == 0 else cm(0.6) + half
        y = cm(2.6) + (i // 2) * cm(1.5) + (cm(0.3) if i >= 8 else 0)
        c = m.add(Control("text", name, x, y, half, cm(0.8), dict(props)))
        labelled(m, name, caption, c)
    c = m.add(Control("text", "txtSignerArgs", cm(0.4), cm(12.0), width - cm(0.8), cm(0.8), {}))
    labelled(m, "txtSignerArgs", 'معاملات برنامج التوقيع: {IN} ملف المستند، {OUT} ملف التوقيع، {PIN} الرقم السري (فارغ = "{IN}" "{OUT}" "{PIN}")', c)
    m.add(Control("label", "lblReady", cm(0.4), cm(13.2), width - cm(0.8), cm(1.6),
                  {"Caption": " ", "FontSize": 9, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("label", "lblSetupNote", cm(0.4), cm(14.9), width - cm(0.8), cm(1.4),
                  {"Caption": "سجّل الجهاز في بوابة المصلحة ثم انسخ بياناته هنا. الرقم الضريبي والعنوان من الإعدادات، "
                              "وكود المصلحة لكل صنف من شاشة المنتجات، ورمز الوحدة من وحدات القياس.",
                   "FontSize": 9, "ForeColor": Sym("CLR_MUTED")}))
    y = cm(17.2)
    bx = cm(0.4)
    for name, caption, style, w, call in [
            ("btnSave", "حفظ", "primary", 3.0, "EtaSetupSave Me"),
            ("btnTestLogin", "تجربة الدخول", "secondary", 3.4, "EtaSetupTestLogin Me")]:
        button(m, name, caption, bx, y, style, w=cm(w), h=cm(0.9), call=call)
        bx += cm(w) + cm(0.2)
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), y, "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = ["Private Sub Form_Load()", "    EtaSetupLoad Me", "End Sub"] + m.code
    return m


def einvoice_forms() -> List[FormModel]:
    return [layout_einvoices()]
