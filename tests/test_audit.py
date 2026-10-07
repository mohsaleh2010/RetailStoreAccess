"""Audit trail (modAudit, AuditLog + AuditChanges, frmAuditLog, report AUDIT_TRAIL): who added, changed or
deleted what, with the values before and after. The comparison and the action names really run in
LibreOffice; the hooks in the screens and in the code that deletes or changes documents are checked here."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import forms as F
import forms_accounting as A
import vba_harness as H
from relations import relations
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
class AuditRuntimeTests(unittest.TestCase):

    def test_compare_and_names(self):
        h = H.Harness()
        try:
            h.load({"modAudit": H.read_module("modAudit")})
            same = lambda a, b: bool(h.call("modAudit", "SameValue", a, b))     # noqa: E731
            self.assertTrue(same("", ""))
            self.assertTrue(same("5", 5))
            self.assertFalse(same(1, 2))
            self.assertFalse(same("a", ""))
            name = lambda a: h.call("modAudit", "ActionName", a)                 # noqa: E731
            self.assertEqual(name("EDIT"), "تعديل")
            self.assertEqual(name("DELETE"), "حذف")
            self.assertEqual(name("CHEQUE_DELETE"), "حذف (CHEQUE_DELETE)")
            self.assertEqual(name("VAT_EDIT"), "تعديل (VAT_EDIT)")
            self.assertEqual(name("SALE"), "SALE")
        finally:
            h.close()


class AuditSchemaTests(unittest.TestCase):

    def test_tables_and_permission(self):
        self.assertIn("RecordLabel", [f.name for f in table("AuditLog").fields])
        fields = {f.name: f for f in table("AuditChanges").fields}
        for name in ("LogID", "LineNo", "FieldName", "FieldCaption", "OldValue", "NewValue"):
            self.assertIn(name, fields)
        rel = {(r.parent, r.child): r for r in relations()}[("AuditLog", "AuditChanges")]
        self.assertTrue(rel.cascade_delete)
        self.assertIn("AUDIT_LOG", [r[0] for r in table("Permissions").seed_rows])
        roles = {(r, p) for r, p in table("RolePermissions").seed_rows}
        self.assertIn((1, "AUDIT_LOG"), roles)                  # administrator only
        self.assertNotIn((2, "AUDIT_LOG"), roles)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmAuditLog"], "AUDIT_LOG")
        self.assertIn("frmAuditLog", {s[0] for s in SCREEN_LIST})


class AuditHookTests(unittest.TestCase):

    def test_data_screens(self):
        forms = read("modForms")
        before = proc(forms, "FormBeforeUpdate")
        self.assertLess(before.index("AuditFormBefore frm"), before.index("FormBeforeUpdate = True"))
        self.assertGreater(before.index("AuditFormBefore frm"), before.index("AssignSequence(frm)"))
        self.assertIn("AuditFormAfter frm", proc(forms, "FormAfterUpdate"))
        delete = proc(forms, "DeleteRecord")
        self.assertLess(delete.index("AuditSnapshot(table, pk, id)"), delete.index('CurrentDb.Execute "DELETE FROM ['))
        self.assertIn('AuditDeleted "DELETE", table, id, before', delete)
        self.assertNotIn('LogAction "SAVE"', forms)

    def test_grids(self):
        budget = read("modBudget")
        self.assertIn('AuditFormBefore lines, "BudgetLines", "BudgetLineID"', proc(budget, "BudgetLineCheck"))
        self.assertIn("AuditFormAfter lines", proc(budget, "BudgetLineSaved"))
        lines = "\n".join(MODELS["frmPayrollLines"].code)
        self.assertIn('AuditFormBefore Me, "PayrollLines", "PayrollLineID"', lines)
        self.assertIn("AuditFormAfter Me", lines)

    def test_every_document_delete_keeps_the_record(self):
        found = 0
        for module in ("modAssets", "modBank", "modBudget", "modCheque", "modManualEntry", "modPayroll", "modVat"):
            text = read(module)
            self.assertNotRegex(text, r'LogAction "\w+_DELETE"', module)
            for m in re.finditer(r'CurrentDb\.Execute "DELETE FROM (\w+) WHERE (\w+) = " & (\w+)', text):
                t, k, v = m.groups()
                if t in ("FixedAssets", "DepreciationRuns", "BankTransactions", "BankReconciliations", "Budgets",
                         "Cheques", "ManualEntries", "PayrollRuns", "VatReturns"):
                    head = text[:m.start()]
                    self.assertIn(f'Set auditBefore = AuditSnapshot("{t}", "{k}", {v})', head[-300:], t)
                    self.assertRegex(text[m.end():m.end() + 200], rf'AuditDeleted "\w+_DELETE", "{t}", {v}, auditBefore')
                    found += 1
        self.assertEqual(found, 9)

    def test_document_edits(self):
        for module, call in [("modAssets", 'AuditEdited "FIXED_ASSET_EDIT", "FixedAssets", "AssetID"'),
                             ("modManualEntry", 'AuditEdited "MANUAL_ENTRY_EDIT", "ManualEntries", "ManualEntryID"'),
                             ("modVat", 'AuditEdited "VAT_EDIT", "VatReturns", "VatReturnID"')]:
            self.assertIn(call, read(module))

    def test_secrets_never_written(self):
        text = read("modAudit")
        self.assertIn('IsSecretField = (FieldName = "PasswordHash" Or FieldName = "PasswordSalt")',
                      proc(text, "IsSecretField"))
        self.assertIn('changes.Add Array(FieldName, caption, "", "(تغيّرت)")', proc(text, "AddChange"))
        self.assertIn("Not IsSecretField(f.Name)", proc(text, "AuditSnapshot"))

    def test_screen_report_and_hub(self):
        text = read("modAudit")
        names = {c.name for c in MODELS["frmAuditLog"].controls}
        for pname in ("AuditScreenLoad", "AuditScreenShow", "AuditScreenPick", "PrintAuditLog"):
            for ctl in set(re.findall(r"frm!(\w+)", proc(text, pname))):
                self.assertIn(ctl, names, f"{pname}: {ctl}")
        entry = {r.key: r for r in F.REPORTS}["AUDIT_TRAIL"]
        self.assertIn("A", entry.needs)
        self.assertIn('If HasNeed(needs, "A") And Not HasPermission("AUDIT_LOG") Then', read("modScreens"))
        self.assertIn("frmAuditLog", {t[3] for t in A.HUB_TILES})
        self.assertIn("btnAuditLog", {c.name for c in MODELS["frmUsers"].controls})
        self.assertIn('"TestAudit"', read("modTestAll"))


class Static_modAudit(VbaModuleChecks, unittest.TestCase):
    module_name = "modAudit"
    vba = read("modAudit")


if __name__ == "__main__":
    unittest.main()
