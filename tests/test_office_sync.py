"""Plan 6H (docs/84): Access kept for the office next to the web, the Windows POS and the rep app.

  * the settings and the log of the office tasks (Settings.OfficeSync*, OfficeSyncLog), permission OFFICE_SYNC
    (administrators only), the screen frmOfficeSync reached from the backup screen;
  * modOfficeSync: the automatic run only on the named computer, from the dashboard timer, never twice at once;
    a run = SyncJournal, then the pending e-invoices sent and followed; each run logged;
  * the rules of the run (due / stale) are the same in VBA and in C# (OfficeSyncRules, xUnit);
  * the web warns its administrators (Home.razor, /api/v1/dashboard/office-sync).
"""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import forms as F
import generate as GEN
from schema import SCREEN_LIST, table, _role_permissions


def read(*parts):
    with open(os.path.join(ROOT, *parts), encoding="utf-8") as fh:
        return fh.read()


VBA = read("src", "vba", "modOfficeSync.bas")


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


def due(last_minutes_ago, minutes):
    """Python mirror of OfficeSyncDue (None = never ran)."""
    if minutes <= 0:
        return False
    return last_minutes_ago is None or last_minutes_ago >= minutes


def stale(last_minutes_ago, minutes):
    """Python mirror of OfficeSyncStale."""
    if minutes <= 0:
        return False
    return last_minutes_ago is None or last_minutes_ago > 2 * minutes + 10


class SchemaTests(unittest.TestCase):
    def test_settings_fields_are_optional_or_defaulted(self):
        fields = {f.name: f for f in table("Settings").fields}
        for name in ("OfficeSyncComputer", "OfficeSyncMinutes", "OfficeSyncLastAt", "OfficeSyncLastResult"):
            self.assertIn(name, fields)
            self.assertFalse(fields[name].required, name)
        self.assertEqual(fields["OfficeSyncMinutes"].default, "15")

    def test_log_table(self):
        t = table("OfficeSyncLog")
        names = [f.name for f in t.fields]
        for name in ("RunAt", "ComputerName", "RunKind", "JournalAdded", "EInvoicesSent", "EInvoicesFailed", "Problem"):
            self.assertIn(name, names)

    def test_permission_is_for_administrators_only(self):
        rows = _role_permissions()
        self.assertIn((1, "OFFICE_SYNC"), rows)
        self.assertNotIn((2, "OFFICE_SYNC"), rows)
        self.assertIn(("frmOfficeSync", "مهام المكتب", "النظام", "OFFICE_SYNC", False, False, False), SCREEN_LIST)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmOfficeSync"], "OFFICE_SYNC")

    def test_screen_is_built_and_reached_from_backup(self):
        names = [m.name for m in F.all_forms()]
        self.assertIn("frmOfficeSync", names)
        self.assertIn('OpenScreen "frmOfficeSync", 0', read("tools", "forms_security.py"))
        self.assertIn("modOfficeSync", GEN.STATIC_MODULES)
        self.assertIn('"TestOfficeSync"', read("src", "vba", "modTestAll.bas"))


class RuleTests(unittest.TestCase):
    def test_due_and_stale(self):
        self.assertTrue(due(None, 15))
        self.assertFalse(due(None, 0))
        self.assertFalse(due(14, 15))
        self.assertTrue(due(15, 15))
        self.assertFalse(stale(40, 15))
        self.assertTrue(stale(41, 15))

    def test_vba_has_the_same_numbers(self):
        self.assertIn("OfficeSyncDue = (WholeMinutes(CDate(LastAt), NowAt) >= Minutes)", proc(VBA, "OfficeSyncDue"))
        self.assertIn("OfficeSyncStale = (WholeMinutes(CDate(LastAt), NowAt) > 2 * Minutes + 10)", proc(VBA, "OfficeSyncStale"))

    def test_a_time_given_as_a_number(self):
        """In VBA Date - Date is a Double and IsDate(Double) is False: OfficeSyncDue(t0 - TimeSerial(0, 14, 0), ...)
        took it as "never ran" and said due (TestOfficeSync "التشغيل التالي بعد عدد الدقائق" failed three times)."""
        body = proc(VBA, "OfficeSyncDue")
        self.assertIn("If Not IsDate(LastAt) And Not IsNumeric(LastAt) Then", body)
        self.assertNotIn("If Not IsDate(LastAt) Then", body)
        self.assertIn('OfficeSyncDue(DateAdd("n", -14, t0), 15, t0)', proc(VBA, "TestOfficeSync"))

    def test_whole_minutes_without_date_doubles(self):
        """Minutes are counted from whole seconds as integers, as OfficeSyncRules.WholeMinutes does."""
        self.assertNotIn('DateDiff("n", CDate(LastAt)', VBA)
        self.assertIn("secs = Round(CDbl(At) * 86400#, 0)", proc(VBA, "MinuteIndex"))
        self.assertIn("MinuteIndex = Int(secs / 60#)", proc(VBA, "MinuteIndex"))
        # the same arithmetic in Python, on the Doubles VBA has: 14 and 15 minutes before 10:00
        import datetime, math
        base = (datetime.datetime(2026, 11, 16, 10, 0) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
        index = lambda d: math.floor(round(d * 86400) / 60)
        for m in range(0, 61):
            self.assertEqual(index(base) - index(base - m * 60 / 86400), m)
        self.assertIn("> 2 * Minutes + 10", proc(VBA, "OfficeSyncStale"))
        cs = read("web", "src", "RetailStore.Domain", "Common", "OfficeSyncRules.cs")
        self.assertIn(">= minutes", cs)
        self.assertIn("> 2 * minutes + 10", cs)


class CodeTests(unittest.TestCase):
    def test_tick_runs_only_on_the_office_computer(self):
        tick = proc(VBA, "OfficeSyncTick")
        order = [tick.index(s) for s in ("m_running", "CurrentUserID() = 0", "IsOfficeSyncComputer()",
                                          "OfficeSyncDue(", 'RunOfficeSync "AUTO"')]
        self.assertEqual(order, sorted(order))
        self.assertNotIn("Show", tick)                     # never a message from the timer

    def test_dashboard_timer_calls_the_tick(self):
        self.assertIn("OfficeSyncTick", proc(read("src", "vba", "modDashboard.bas"), "DashboardRefresh"))

    def test_run_order_and_log(self):
        run = proc(VBA, "RunOfficeSync")
        order = [run.index(s) for s in ("SyncJournal(", "EInvoiceReadyProblem()", "SendPendingEInvoices(",
                                        "RefreshSubmittedEInvoices(", "WriteOfficeSyncLog")]
        self.assertEqual(order, sorted(order))
        self.assertIn("m_running = False", run)
        log = proc(VBA, "WriteOfficeSyncLog")
        self.assertIn("INSERT INTO OfficeSyncLog", log)
        self.assertIn("OfficeSyncLastAt", log)

    def test_settings_change_is_audited(self):
        save = proc(VBA, "OfficeSyncSave")
        self.assertIn('AuditSnapshot("Settings"', save)
        self.assertIn("AuditEdited", save)

    def test_test_never_sends_to_a_platform(self):
        self.assertIn("EInvoiceEnabled = 0", proc(VBA, "TestOfficeSync"))


class WebTests(unittest.TestCase):
    def test_web_warns_who_may_run_the_tasks(self):
        home = read("web", "src", "RetailStore.UI", "Pages", "Home.razor")
        self.assertIn('HasPermissionAsync(UserInfo.EmployeeId(user), "OFFICE_SYNC")', home)
        self.assertIn("NeedsAttention", home)
        api = read("web", "src", "RetailStore.Web", "Api", "TreasuryApi.cs")
        self.assertIn('"/dashboard/office-sync"', api)
        self.assertIn('.Needs("OFFICE_SYNC", "frmOfficeSync")', api)


class Static_modOfficeSync(VbaModuleChecks, unittest.TestCase):
    module_name = "modOfficeSync"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
