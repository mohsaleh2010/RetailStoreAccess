"""Phase 3 checks: relationship definitions, the generated VBA module, and the
same referential-integrity scenario that TestRelationships runs inside Access,
executed here against an SQLite mirror of the back-end.

Run:  python3 -m unittest discover -s tests -v
"""

import sqlite3
import unittest

from helpers import VbaModuleChecks, build_sqlite
import gen_relations
import relations as R
from schema import TABLES, table


class RelationDefinitionTests(unittest.TestCase):

    def setUp(self):
        self.rels = R.relations()

    def test_count_and_names(self):
        self.assertEqual(len(self.rels), 105)
        names = [r.name for r in self.rels]
        self.assertEqual(len(names), len(set(names)))
        for n in names:
            self.assertLessEqual(len(n), 64, n)

    def test_cascade_delete_only_from_header_to_its_lines(self):
        cascading = {(r.parent, r.child) for r in self.rels if r.cascade_delete}
        self.assertEqual(cascading, {
            ("SalesInvoices", "SalesInvoiceDetails"),
            ("SalesReturns", "SalesReturnDetails"),
            ("PurchaseInvoices", "PurchaseInvoiceDetails"),
            ("PurchaseReturns", "PurchaseReturnDetails"),
            ("StockCounts", "StockCountDetails"),
            ("Roles", "RolePermissions"),
            ("Permissions", "RolePermissions"),
            ("JournalEntries", "JournalLines"),
            ("ManualEntries", "ManualEntryLines"),
            ("CustomerPayments", "CustomerAllocations"),
            ("SupplierPayments", "SupplierAllocations"),
            ("BankReconciliations", "BankClearings"),
            ("FiscalYearClosings", "FiscalYearClosingLines"),
        })

    def test_cascade_update_only_for_text_keys(self):
        self.assertEqual([r.name for r in self.rels if r.cascade_update],
                         ["FK_RolePermissions_PermissionKey", "FK_UserScreens_ScreenName",
                          "FK_JournalEntries_SourceType"])

    def test_documents_are_never_cascade_deleted_from_master_data(self):
        masters = {"Customers", "Suppliers", "Products", "Employees", "Categories", "Units",
                   "PaymentMethods", "ExpenseTypes", "TransactionTypes"}
        for r in self.rels:
            if r.parent in masters:
                self.assertFalse(r.cascade_delete, r.name)

    def test_every_table_except_roots_is_connected(self):
        connected = {r.parent for r in self.rels} | {r.child for r in self.rels}
        self.assertEqual({t.name for t in TABLES} - connected, {"Sequences", "LabelSettings"})

    def test_attribute_values(self):
        by_name = {r.name: r for r in self.rels}
        self.assertEqual(by_name["FK_SalesInvoiceDetails_SalesInvoiceID"].attributes, 4096)
        self.assertEqual(by_name["FK_RolePermissions_PermissionKey"].attributes, 4352)
        self.assertEqual(by_name["FK_SalesInvoices_CustomerID"].attributes, 0)


class IntegrityScenarioTests(unittest.TestCase):
    """Mirror of TestRelationships (modBuildRelations) on SQLite."""

    def setUp(self):
        self.con = build_sqlite(with_relationship_rules=True)

    def scalar(self, sql):
        return self.con.execute(sql).fetchone()[0]

    def assertRejected(self, sql, label):
        with self.assertRaises(sqlite3.IntegrityError, msg=label) as ctx:
            self.con.execute(sql)
        self.assertIn("FOREIGN KEY", str(ctx.exception), f"{label}: {ctx.exception}")

    def test_seed_data_satisfies_all_relationships(self):
        self.assertEqual(self.con.execute("PRAGMA foreign_key_check").fetchall(), [])

    def test_orphan_inserts_rejected(self):
        for label, sql in R.REJECTED_INSERTS:
            self.assertRejected(sql, label)

    def test_full_scenario(self):
        c = self.con
        c.execute(R.CASCADE_PRODUCT_SQL)
        product = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute(R.CASCADE_INVOICE_SQL)
        invoice = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        for line in (1, 2):
            c.execute(R.CASCADE_LINE_SQL.format(invoice=invoice, line=line, product=product))
        self.assertEqual(self.scalar(R.CASCADE_COUNT_LINES_SQL.format(invoice=invoice)), 2)

        for label, sql in R.REJECTED_DELETES:
            self.assertRejected(sql, label)
        self.assertRejected(R.CASCADE_DELETE_PRODUCT_SQL.format(product=product), "product in use")

        c.execute(R.CASCADE_DELETE_INVOICE_SQL.format(invoice=invoice))
        self.assertEqual(self.scalar(R.CASCADE_COUNT_LINES_SQL.format(invoice=invoice)), 0)

        before = self.scalar(R.COUNT_OLD_KEY_SQL)
        self.assertGreater(before, 0)
        c.execute(R.RENAME_KEY_SQL)
        self.assertEqual(self.scalar(R.COUNT_RENAMED_SQL), before)
        self.assertEqual(self.scalar(R.COUNT_OLD_KEY_SQL), 0)

    def test_return_lines_block_deleting_the_original_invoice(self):
        """An invoice that has returns cannot be removed, even via cascade."""
        c = self.con
        c.execute(R.CASCADE_PRODUCT_SQL)
        product = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute(R.CASCADE_INVOICE_SQL)
        invoice = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute(R.CASCADE_LINE_SQL.format(invoice=invoice, line=1, product=product))
        line_id = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute("INSERT INTO SalesReturns (ReturnNumber, SalesInvoiceID, CustomerID, "
                  "EmployeeID, Reason) VALUES ('CRN-T', ?, 1, 1, 'test')", (invoice,))
        ret = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute("INSERT INTO SalesReturnDetails (SalesReturnID, SalesDetailID, ProductID, "
                  "Quantity) VALUES (?, ?, ?, 1)", (ret, line_id, product))
        self.assertRejected(R.CASCADE_DELETE_INVOICE_SQL.format(invoice=invoice),
                            "invoice with returns")


class GeneratedRelationsModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modBuildRelations"
    vba = gen_relations.build_relations_vba()

    def test_every_relation_is_in_the_module(self):
        for r in R.relations():
            self.assertIn(f'c.Add Array("{r.name}", "{r.parent}", "{r.parent_field}", '
                          f'"{r.child}", "{r.child_field}", {r.attributes}&)', self.vba)

    def test_expected_count_constant(self):
        self.assertIn(f"EXPECTED_RELATION_COUNT As Long = {len(R.relations())}", self.vba)

    def test_all_scenario_statements_present(self):
        for _, sql in R.REJECTED_INSERTS + R.REJECTED_DELETES:
            self.assertIn(sql, self.vba)

    def test_runtime_sql_is_concatenated_correctly(self):
        self.assertIn('VALUES (" & invoiceID & ", " & lineNo & ", " & productID & ", 2, 10)"',
                      self.vba)
        self.assertIn('"DELETE FROM [Products] WHERE [ProductID] = " & productID,', self.vba)

    def test_public_api(self):
        for name in ("Public Function BuildRelationships", "Public Function TestRelationships",
                     "Public Sub DropRelationships"):
            self.assertIn(name, self.vba)


if __name__ == "__main__":
    unittest.main()
