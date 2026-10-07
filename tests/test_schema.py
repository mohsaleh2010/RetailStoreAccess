"""Automated checks for the Phase 2 schema and the generated VBA module.

Run:  python3 -m unittest discover -s tests -v
"""

import re
import unittest

from helpers import VbaModuleChecks, build_sqlite
from schema import TABLES, table
import generate

KINDS = {"AUTO", "LONG", "INT", "BYTE", "MONEY", "QTY", "RATE", "DATE", "DATETIME",
         "BOOL", "TEXT", "MEMO"}

# Access / Jet reserved words that must never be used as object names
RESERVED = {
    "date", "time", "name", "value", "text", "level", "section", "position", "password",
    "user", "year", "month", "day", "order", "group", "select", "table", "index", "key",
    "number", "type", "note", "memo", "field", "count", "sum", "avg", "min", "max",
    "currency", "money", "integer", "long", "single", "double", "byte", "boolean",
    "general", "percent", "column", "databases", "property", "section", "size", "status_",
    "from", "where", "by", "and", "or", "not", "in", "is", "like", "between", "null",
    "true", "false", "yes", "no", "on", "off", "full", "left", "right", "join", "union",
    "add", "alter", "drop", "create", "delete", "insert", "update", "values", "set",
    "top", "percent", "distinct", "having", "as", "asc", "desc", "all", "any", "exists",
    "owner", "level", "option", "references", "constraint", "unique", "primary",
}

def field(table_name, field_name):
    for f in table(table_name).fields:
        if f.name == field_name:
            return f
    raise KeyError(f"{table_name}.{field_name}")


class SchemaStructureTests(unittest.TestCase):

    def test_table_names_unique(self):
        names = [t.name for t in TABLES]
        self.assertEqual(len(names), len(set(names)))

    def test_required_tables_from_specification_exist(self):
        required = ["Products", "Categories", "Customers", "Suppliers", "SalesInvoices",
                    "SalesInvoiceDetails", "PurchaseInvoices", "PurchaseInvoiceDetails",
                    "Expenses", "Employees", "CustomerPayments", "SupplierPayments",
                    "InventoryTransactions"]
        names = {t.name for t in TABLES}
        for r in required:
            self.assertIn(r, names)

    def test_fields_from_specification_exist(self):
        spec = {
            "Products": "ProductID ProductCode Barcode ProductName CategoryID PurchasePrice "
                        "SellingPrice CurrentQuantity MinimumQuantity SupplierID ProductLocation "
                        "Notes IsActive",
            "Categories": "CategoryID CategoryName Description",
            "Customers": "CustomerID CustomerName Mobile Phone Address Email OpeningBalance "
                         "CurrentBalance Notes",
            "Suppliers": "SupplierID SupplierName ContactPerson Mobile Phone Address Email "
                         "OpeningBalance CurrentBalance Notes",
            "SalesInvoices": "SalesInvoiceID InvoiceNumber InvoiceDate CustomerID EmployeeID "
                             "PaymentType SubTotal Discount Tax TotalAmount PaidAmount "
                             "RemainingAmount Notes",
            "SalesInvoiceDetails": "SalesDetailID SalesInvoiceID ProductID Quantity UnitPrice "
                                   "Discount Tax LineTotal",
            "PurchaseInvoices": "PurchaseInvoiceID InvoiceNumber InvoiceDate SupplierID "
                                "EmployeeID SubTotal Discount Tax TotalAmount PaidAmount "
                                "RemainingAmount Notes",
            "PurchaseInvoiceDetails": "PurchaseDetailID PurchaseInvoiceID ProductID Quantity "
                                      "UnitCost Discount Tax LineTotal",
            "Expenses": "ExpenseID ExpenseDate Amount Description EmployeeID",
            "Employees": "EmployeeID EmployeeName JobTitle Mobile Username IsActive",
            "CustomerPayments": "PaymentID CustomerID PaymentDate Amount PaymentMethodID Notes",
            "SupplierPayments": "PaymentID SupplierID PaymentDate Amount PaymentMethodID Notes",
            "InventoryTransactions": "TransactionID TransactionDate ProductID Quantity "
                                     "ReferenceNumber ReferenceType Notes",
        }
        for tname, fields in spec.items():
            names = {f.name for f in table(tname).fields}
            for fname in fields.split():
                self.assertIn(fname, names, f"{tname}.{fname}")

    def test_fields_valid(self):
        for t in TABLES:
            names = [f.name for f in t.fields]
            self.assertEqual(len(names), len(set(names)), f"duplicate field in {t.name}")
            self.assertLess(len(t.fields), 255)
            for f in t.fields:
                where = f"{t.name}.{f.name}"
                self.assertIn(f.kind, KINDS, where)
                self.assertTrue(f.caption, f"missing caption {where}")
                self.assertLessEqual(len(f.name), 64, where)
                self.assertNotIn(f.name.lower(), RESERVED, f"reserved word {where}")
                self.assertTrue(re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", f.name), where)
                if f.kind == "TEXT":
                    self.assertTrue(1 <= f.size <= 255, f"text size {where}")
                else:
                    self.assertEqual(f.size, 0, where)
                if f.default and f.kind == "TEXT":
                    self.assertRegex(f.default, r'^".*"$', f"text default must be quoted {where}")
                if f.rule:
                    self.assertTrue(f.rule_text, f"validation text missing {where}")
            self.assertNotIn(t.name.lower(), RESERVED)

    def test_exactly_one_autonumber_and_it_is_the_pk(self):
        for t in TABLES:
            autos = [f.name for f in t.fields if f.kind == "AUTO"]
            self.assertLessEqual(len(autos), 1, t.name)
            if autos:
                self.assertEqual(t.pk, autos, t.name)

    def test_primary_and_index_fields_exist(self):
        for t in TABLES:
            names = {f.name for f in t.fields}
            self.assertTrue(t.pk, f"{t.name} has no primary key")
            for k in t.pk:
                self.assertIn(k, names, t.name)
            index_names = ["PrimaryKey"]
            for ix in t.indexes:
                index_names.append(ix.name)
                for k in ix.fields:
                    self.assertIn(k, names, f"{t.name}.{ix.name}")
            self.assertEqual(len(index_names), len(set(index_names)), t.name)

    def test_table_rules_reference_existing_fields(self):
        for t in TABLES:
            if t.rule:
                names = {f.name for f in t.fields}
                for ref in re.findall(r"\[(\w+)\]", t.rule):
                    self.assertIn(ref, names, t.name)

    def test_foreign_keys_point_to_keys_with_matching_type(self):
        for t in TABLES:
            for f in t.fields:
                if not f.fk:
                    continue
                tt, tf = f.fk.split(".")
                target_table = table(tt)
                self.assertEqual(target_table.pk, [tf], f"{t.name}.{f.name} -> {f.fk}")
                target = field(tt, tf)
                if target.kind in ("AUTO", "LONG"):
                    self.assertEqual(f.kind, "LONG", f"{t.name}.{f.name}")
                else:
                    self.assertEqual((f.kind, f.size), (target.kind, target.size),
                                     f"{t.name}.{f.name}")

    def test_details_cascade_from_headers(self):
        for tname, fname in [("SalesInvoiceDetails", "SalesInvoiceID"),
                             ("PurchaseInvoiceDetails", "PurchaseInvoiceID"),
                             ("SalesReturnDetails", "SalesReturnID"),
                             ("PurchaseReturnDetails", "PurchaseReturnID"),
                             ("StockCountDetails", "StockCountID")]:
            self.assertTrue(field(tname, fname).on_delete_cascade, f"{tname}.{fname}")

    def test_money_amounts_cannot_be_negative_except_balances(self):
        allowed_negative = {"OpeningBalance", "CurrentBalance", "DifferenceValue", "ExpectedBalance", "Difference",
                            "NetProfit",                     # a year closing may be a loss
                            # the VAT return: returns are negative adjustments, the VAT of a period
                            # may be a credit, corrections go both ways
                            "SalesStdAdjust", "SalesStdVAT", "SalesZeroAdjust", "SalesExemptAdjust",
                            "PurchStdAdjust", "PurchStdVAT", "PurchZeroAdjust", "Corrections", "NetDue",
                            # a bank account may be overdrawn; the operations in a reconciliation go both ways
                            "StatementBalance", "BookBalance", "Outstanding", "ClearedAmount"}
        for t in TABLES:
            for f in t.fields:
                if f.kind == "MONEY" and f.name not in allowed_negative:
                    self.assertIn(f.rule, (">=0", ">0"), f"{t.name}.{f.name}")

    def test_vat_number_pattern(self):
        rule = field("Settings", "VATNumber").rule
        pattern = re.search(r'"(.*)"', rule).group(1)
        self.assertEqual(len(pattern), 15)
        self.assertTrue(pattern.startswith("3") and pattern.endswith("3"))


class SeedDataTests(unittest.TestCase):

    def test_seed_rows_match_columns_and_required_fields(self):
        for t in TABLES:
            if not t.seed_rows:
                continue
            names = {f.name for f in t.fields}
            for c in t.seed_columns:
                self.assertIn(c, names, f"{t.name}.{c}")
            needed = {f.name for f in t.fields
                      if f.required and f.default is None and f.kind not in ("AUTO", "BOOL")}
            self.assertTrue(needed <= set(t.seed_columns), f"{t.name} misses {needed}")
            for row in t.seed_rows:
                self.assertEqual(len(row), len(t.seed_columns), f"{t.name} {row}")

    def test_no_zero_length_text_in_seed(self):
        # Access text fields refuse '' (AllowZeroLength = False, error 3315): use None
        for t in TABLES:
            kinds = {f.name: f.kind for f in t.fields}
            for row in t.seed_rows:
                for c, v in zip(t.seed_columns, row):
                    if kinds[c] in ("TEXT", "MEMO"):
                        self.assertNotEqual(v, "", f"{t.name}.{c} in {row}")

    def test_cash_customer_and_admin(self):
        cust = table("Customers")
        row = dict(zip(cust.seed_columns, cust.seed_rows[0]))
        self.assertEqual(row["CustomerID"], 1)
        self.assertFalse(row["AllowCredit"])
        emp = table("Employees")
        row = dict(zip(emp.seed_columns, emp.seed_rows[0]))
        self.assertEqual((row["Username"], row["RoleID"], row["MustChangePassword"]),
                         ("admin", 1, True))

    def test_role_permissions(self):
        rp = table("RolePermissions").seed_rows
        all_perms = {r[0] for r in table("Permissions").seed_rows}
        self.assertEqual({p for r, p in rp if r == 1}, all_perms)
        cashier = {p for r, p in rp if r == 3}
        self.assertEqual(cashier, {"SALES_POS", "SALES_VIEW", "CUSTOMERS", "CUSTOMER_PAYMENTS", "CASH_CLOSING"})
        manager = {p for r, p in rp if r == 2}
        for p in ("SALES_POS", "PURCHASES", "REPORTS", "STOCK_COUNT"):
            self.assertIn(p, manager)
        for p in ("USERS", "SETTINGS", "BACKUP"):
            self.assertNotIn(p, manager)

    def test_transaction_type_directions(self):
        dirs = {r[1]: r[3] for r in table("TransactionTypes").seed_rows}
        self.assertEqual(dirs, {"PURCHASE": 1, "SALE": -1, "PURCHASE_RETURN": -1,
                                "SALES_RETURN": 1, "STOCK_IN": 1, "STOCK_OUT": -1,
                                "ADJUSTMENT": 0, "OPENING": 1})

    def test_seed_loads_into_relational_db_with_foreign_keys(self):
        con = build_sqlite()
        self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])
        self.assertEqual(con.execute("SELECT COUNT(*) FROM RolePermissions").fetchone()[0],
                         len(table("RolePermissions").seed_rows))


class GeneratedSchemaModuleTests(VbaModuleChecks, unittest.TestCase):
    module_name = "modBuildSchema"
    vba = generate.build_vba()

    def test_every_called_procedure_is_defined(self):
        defined = set(re.findall(r"(?:Sub|Function) (\w+)\(", self.vba))
        create_all = self.vba.split("Private Sub CreateAllTables")[1].split("End Sub")[0]
        for t in TABLES:
            self.assertIn(f"CreateTable_{t.name}", defined)
            self.assertIn(f"CreateTable_{t.name}", create_all)
        for name in ("BuildSchema", "VerifySchema", "LinkBackEnd", "DropSchema", "AddField",
                     "AddIndex", "BeginTable", "EndTable", "SetProp", "BeginSeed", "ExecSeed",
                     "EndSeed", "SeedAll", "CreateAllTables"):
            self.assertIn(name, defined)


if __name__ == "__main__":
    unittest.main()
