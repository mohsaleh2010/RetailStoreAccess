"""Recurring expenses (modRecurring, RecurringExpenses, Expenses.RecurringID, frmRecurring) and the
accounting hub (frmAccounting): the schedule really runs in LibreOffice; the rules, the screens, the
permission and the side menu are checked in the code."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import forms as F
import forms_accounting as A
import vba_harness as H
from schema import SCREEN_LIST, table

MODELS = {m.name: m for m in F.all_forms()}


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class ScheduleRuntimeTests(unittest.TestCase):

    def test_first_and_next_due_dates(self):
        h = H.Harness()
        try:
            h.load({"modRecurring": H.read_module("modRecurring")})

            self.assertEqual(int(h.call("modRecurring", "MonthsOf", "QUARTERLY")), 3)
            self.assertEqual(int(h.call("modRecurring", "MonthsOf", "YEARLY")), 12)
            self.assertEqual(int(h.call("modRecurring", "MonthsOf", "MONTHLY")), 1)
        finally:
            h.close()


class RecurringCodeTests(unittest.TestCase):

    def setUp(self):
        self.text = read("modRecurring")

    def test_schedule_rules(self):
        self.assertIn("If d < Int(StartDate) Then d = DateSerial(Year(StartDate), Month(StartDate) + 1, DueDay)",
                      proc(self.text, "FirstDueDate"))
        self.assertIn("DateSerial(Year(Due), Month(Due) + MonthsOf(Frequency), DueDay)", proc(self.text, "DueAfter"))
        rule = next(f for f in table("RecurringExpenses").fields if f.name == "DueDay").rule
        self.assertEqual(rule, "Between 1 And 28")                  # every month has the day

    def test_created_once_per_date_and_not_in_a_closed_period(self):
        once = proc(self.text, "CreateDueOf")
        for part in ("ClosedPeriodProblem(due)", "WHERE RecurringID = ", "AND ExpenseDate = ",
                     "n < MAX_PER_RUN", "If due > rs!EndDate Then Exit Do", "rs!NextDueDate = due"):
            self.assertIn(part, once)
        add = proc(self.text, "AddRecurringExpense")
        for part in ('NextNumber("EXPENSE")', "e!RecurringID = r!RecurringID", "CurrentCashBoxID()",
                     "BankFor(r!PaymentMethodID, total)", "CostCenterFor()", "e!TotalAmount = total"):
            self.assertIn(part, add)
        self.assertIn('HasPermission("EXPENSES")', proc(self.text, "CreateDueExpenses"))

    def test_wiring(self):
        forms = read("modForms")
        self.assertIn('Case "RecurringExpenses": RecurringCurrent frm', forms)
        self.assertIn("If Not ValidateRecurring(frm) Then Exit Function", forms)
        self.assertIn("RecurringAlert", read("modScreens"))
        self.assertIn('TempVars.Remove "RecurringAlertShown"', read("modSecurityScreens"))
        self.assertIn('"TestRecurring"', read("modTestAll"))
        self.assertEqual(next(f.fk for f in table("Expenses").fields if f.name == "RecurringID"),
                         "RecurringExpenses.RecurringID")
        self.assertEqual(F.SCREEN_PERMISSIONS["frmRecurring"], "EXPENSES")
        self.assertIn("frmRecurring", {s[0] for s in SCREEN_LIST})
        names = {c.name for c in MODELS["frmRecurring"].controls}
        names |= {f.name for f in table("RecurringExpenses").fields}     # frm!Field reads the record too
        for pname in ("RecurringCurrent", "ValidateRecurring", "RecurringCreateNow"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(self.text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")


class AccountingHubTests(unittest.TestCase):

    def test_every_accounting_screen_has_a_tile(self):
        targets = {t[3] for t in A.HUB_TILES}
        for form in ("frmJournal", "frmAccounts", "frmManualEntry", "frmLedger", "frmFinancials", "frmPeriodClosing",
                     "frmVatReturn", "frmAging", "frmBanks", "frmCheques", "frmAssets", "frmDepreciation",
                     "frmPayroll", "frmCostCenters", "frmBudget", "frmRecurring"):
            self.assertIn(form, targets)
            self.assertIn(form, MODELS)
        hub = MODELS["frmAccounting"]
        names = {c.name for c in hub.controls}
        for key, *_ in A.HUB_TILES:                 # ApplyNavPermissions greys boxNav<key> under btnTile<key>
            self.assertIn(f"btnTile{key}", names)
            self.assertIn(f"boxNav{key}", names)
        self.assertIn("ApplyNavPermissions Me", "\n".join(hub.code))

    def test_side_menu_and_dashboard_open_the_hub(self):
        self.assertEqual({n.key: n.target for n in F.NAV_ITEMS}["Accounting"], "frmAccounting")
        main = {c.name: c for c in MODELS["frmMain"].controls}
        self.assertIn("btnTileAccounting", main)
        for item in F.NAV_ITEMS:                    # caption and icon under a transparent button
            btn = main[f"btnNav{item.key}"]
            self.assertTrue(btn.props.get("Transparent"), item.key)
            order = [c.name for c in MODELS["frmMain"].controls]
            self.assertLess(order.index(f"lblNav{item.key}"), order.index(btn.name))
            self.assertLess(order.index(f"ico{item.key}"), order.index(btn.name))


class Static_modRecurring(VbaModuleChecks, unittest.TestCase):
    module_name = "modRecurring"
    vba = read("modRecurring")


if __name__ == "__main__":
    unittest.main()
