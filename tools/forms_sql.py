"""The SQL Server screen (frmSqlServer, plan 5B, docs/68): the connection of this front-end, the database
script, the move of the data, its comparison and the link of the program to the server. Logic: modSqlServer."""

from typing import List

from forms import Control, FormModel, Sym, button, cm, labelled, title_band
from forms_cash import value_list_combo

AUTH = "WINDOWS;حساب Windows;SQL;مستخدم SQL Server"


def _text(m, name, caption, x, y, w, props=None, events=()):
    c = m.add(Control("text", name, x, y, w, cm(0.8), props or {}, events=list(events)))
    labelled(m, name, caption, c)
    return c


def layout_sql_server() -> FormModel:
    width, height = cm(22.0), cm(16.2)
    m = FormModel("frmSqlServer", "قاعدة البيانات على SQL Server", width, height, popup=True, allow_add=False)
    title_band(m, "قاعدة البيانات على SQL Server", "الاتصال بالخادم، وتجهيز القاعدة، ونقل البيانات ومطابقتها، "
               "وتشغيل البرنامج على الخادم", "backup")
    y = cm(2.4)
    _text(m, "txtServer", "الخادم (مثل .\\SQLEXPRESS)", cm(0.4), y, cm(5.6))
    _text(m, "txtDatabase", "قاعدة البيانات", cm(6.2), y, cm(3.6))
    c = value_list_combo(m, "cboAuth", cm(10.0), y, cm(3.4))
    c.props.update({"RowSource": AUTH, "ColumnWidths": "0;3.2"})
    labelled(m, "cboAuth", "الدخول", c)
    _text(m, "txtUser", "المستخدم", cm(13.6), y, cm(3.6))
    _text(m, "txtPassword", "كلمة المرور", cm(17.4), y, cm(4.2), {"InputMask": "Password"})
    y = cm(3.9)
    c = value_list_combo(m, "cboDriver", cm(0.4), y, cm(6.2), events=())
    c.props.update({"ColumnCount": 1, "ColumnWidths": "6"})
    labelled(m, "cboDriver", "برنامج التشغيل (ODBC)", c)
    _text(m, "txtScript", "ملف سكربت القاعدة", cm(6.8), y, cm(11.0))
    button(m, "btnBrowse", "استعراض", cm(18.0), y, "secondary", w=cm(3.6), h=cm(0.8), call="SqlServerBrowseScript Me")
    m.add(Control("label", "lblState", cm(0.4), cm(4.95), width - cm(0.8), cm(0.65),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_MUTED")}))
    x = cm(0.4)
    for name, caption, style, w, call in [("btnTest", "اختبار الاتصال", "secondary", 2.9, "SqlServerTest Me"),
                                          ("btnSchema", "تجهيز القاعدة", "secondary", 3.3, "SqlServerSchema Me"),
                                          ("btnTransfer", "نقل البيانات", "primary", 3.0, "SqlServerTransfer Me"),
                                          ("btnCompare", "مطابقة البيانات", "secondary", 3.2, "SqlServerCompare Me"),
                                          ("btnLink", "التشغيل على الخادم", "primary", 3.4, "SqlServerLink Me"),
                                          ("btnAccess", "الرجوع إلى Access", "danger", 3.6, "SqlServerBackToAccess Me")]:
        button(m, name, caption, x, cm(5.75), style, w=cm(w), h=cm(0.9), call=call)
        x += cm(w) + cm(0.2)
    m.add(Control("label", "lblResult", cm(0.4), cm(6.9), width - cm(0.8), cm(0.6),
                  {"Caption": " ", "FontSize": 10, "FontBold": True, "ForeColor": Sym("CLR_PRIMARY")}))
    m.add(Control("list", "lstResult", cm(0.4), cm(7.6), width - cm(0.8), cm(7.2),
                  {"RowSourceType": "Value List", "RowSource": "", "ColumnCount": 6,
                   "ColumnWidths": "5.2;2.6;2.6;3.6;3.6;2.4", "ColumnHeads": True}))
    button(m, "btnClose", "إغلاق", width - cm(0.4) - cm(2.6), cm(15.05), "secondary", w=cm(2.6), h=cm(0.9),
           call="DoCmd.Close acForm, Me.Name")
    m.form_events = ["Load"]
    m.code = (["Private Sub Form_Load()", "    SqlServerLoad Me", "End Sub",
               "Private Sub cboAuth_AfterUpdate()", "    SqlServerAuthChanged Me", "End Sub"] + m.code)
    return m


def sql_forms() -> List[FormModel]:
    return [layout_sql_server()]
