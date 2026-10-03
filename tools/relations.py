"""Phase 3: relationships derived from the `fk` attributes in schema.py,
plus the referential-integrity test scenarios shared by the VBA self-test
(TestRelationships) and the Python test-suite.
"""

from dataclasses import dataclass
from typing import List

from schema import TABLES, table

REL_CASCADE_UPDATE = 256    # dbRelationUpdateCascade
REL_CASCADE_DELETE = 4096   # dbRelationDeleteCascade


@dataclass
class Relation:
    name: str
    parent: str
    parent_field: str
    child: str
    child_field: str
    cascade_delete: bool
    cascade_update: bool
    required: bool        # child field is mandatory (1-to-many) or optional (0..1-to-many)

    @property
    def attributes(self) -> int:
        a = 0
        if self.cascade_update:
            a |= REL_CASCADE_UPDATE
        if self.cascade_delete:
            a |= REL_CASCADE_DELETE
        return a

    @property
    def rule_ar(self) -> str:
        parts = ["فرض التكامل"]
        if self.cascade_delete:
            parts.append("حذف متتالٍ")
        if self.cascade_update:
            parts.append("تحديث متتالٍ")
        return " + ".join(parts)


def relations() -> List[Relation]:
    out = []
    for t in TABLES:
        for f in t.fields:
            if not f.fk:
                continue
            parent, parent_field = f.fk.split(".")
            parent_kind = next(p.kind for p in table(parent).fields if p.name == parent_field)
            out.append(Relation(
                name=f"FK_{t.name}_{f.name}",
                parent=parent, parent_field=parent_field,
                child=t.name, child_field=f.name,
                cascade_delete=f.on_delete_cascade,
                # only natural (text) keys can change; AutoNumber keys never do
                cascade_update=(parent_kind == "TEXT"),
                required=f.required,
            ))
    return out


# --------------------------------------------------------------------------
# Self-test scenarios. Every statement must be rejected by the database
# because of referential integrity. They run inside a transaction that is
# always rolled back, so no test data is left behind.
# --------------------------------------------------------------------------
MISSING = 999999

REJECTED_INSERTS = [
    ("فاتورة بيع لعميل غير موجود",
     f"INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) "
     f"VALUES ('TEST-RI-1', {MISSING}, 1)"),
    ("فاتورة بيع بموظف غير موجود",
     f"INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) "
     f"VALUES ('TEST-RI-2', 1, {MISSING})"),
    ("سطر بيع لفاتورة غير موجودة",
     f"INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], [ProductID], "
     f"[Quantity], [UnitPrice]) VALUES ({MISSING}, 1, {MISSING}, 1, 1)"),
    ("فاتورة شراء لمورد غير موجود",
     f"INSERT INTO [PurchaseInvoices] ([InvoiceNumber], [SupplierID], [EmployeeID]) "
     f"VALUES ('TEST-RI-3', {MISSING}, 1)"),
    ("منتج بتصنيف غير موجود",
     f"INSERT INTO [Products] ([ProductCode], [ProductName], [CategoryID]) "
     f"VALUES ('TEST-RI-X', 'x', {MISSING})"),
    ("منتج بمورد غير موجود",
     f"INSERT INTO [Products] ([ProductCode], [ProductName], [SupplierID]) "
     f"VALUES ('TEST-RI-Y', 'y', {MISSING})"),
    ("حركة مخزون لمنتج غير موجود",
     f"INSERT INTO [InventoryTransactions] ([ProductID], [TransactionTypeID], [Quantity]) "
     f"VALUES ({MISSING}, 1, 1)"),
    ("سند قبض لعميل غير موجود",
     f"INSERT INTO [CustomerPayments] ([PaymentNumber], [CustomerID], [Amount], [EmployeeID]) "
     f"VALUES ('TEST-RI-4', {MISSING}, 10, 1)"),
    ("سند صرف لمورد غير موجود",
     f"INSERT INTO [SupplierPayments] ([PaymentNumber], [SupplierID], [Amount], [EmployeeID]) "
     f"VALUES ('TEST-RI-5', {MISSING}, 10, 1)"),
    ("مصروف بنوع غير موجود",
     f"INSERT INTO [Expenses] ([ExpenseNumber], [ExpenseTypeID], [Amount], [TotalAmount], "
     f"[EmployeeID]) VALUES ('TEST-RI-6', {MISSING}, 10, 10, 1)"),
    ("مستخدم بدور غير موجود",
     f"INSERT INTO [Employees] ([EmployeeName], [Username], [RoleID]) "
     f"VALUES ('x', 'test_ri_user', {MISSING})"),
    ("صلاحية غير معرّفة لدور",
     "INSERT INTO [RolePermissions] ([RoleID], [PermissionKey]) VALUES (3, 'NO_SUCH_KEY')"),
]

REJECTED_DELETES = [
    ("حذف دور مرتبط بمستخدمين", "DELETE FROM [Roles] WHERE [RoleID] = 1"),
    ("حذف العميل النقدي (مرتبط بالإعدادات)", "DELETE FROM [Customers] WHERE [CustomerID] = 1"),
    ("حذف تصنيف تستخدمه منتجات", "DELETE FROM [Categories] WHERE [CategoryID] = 1"),
    ("حذف المستخدم admin وله عمليات", "DELETE FROM [Employees] WHERE [EmployeeID] = 1"),
]

# Cascade scenario (ids are filled in at run time)
CASCADE_PRODUCT_SQL = ("INSERT INTO [Products] ([ProductCode], [ProductName]) "
                       "VALUES ('TEST-RI-P', 'TEST')")
CASCADE_INVOICE_SQL = ("INSERT INTO [SalesInvoices] ([InvoiceNumber], [CustomerID], [EmployeeID]) "
                       "VALUES ('TEST-RI-INV', 1, 1)")
CASCADE_LINE_SQL = ("INSERT INTO [SalesInvoiceDetails] ([SalesInvoiceID], [LineNumber], "
                    "[ProductID], [Quantity], [UnitPrice]) VALUES ({invoice}, {line}, {product}, 2, 10)")
CASCADE_DELETE_PRODUCT_SQL = "DELETE FROM [Products] WHERE [ProductID] = {product}"
CASCADE_DELETE_INVOICE_SQL = "DELETE FROM [SalesInvoices] WHERE [SalesInvoiceID] = {invoice}"
CASCADE_COUNT_LINES_SQL = ("SELECT COUNT(*) FROM [SalesInvoiceDetails] "
                           "WHERE [SalesInvoiceID] = {invoice}")

# Cascade update of a text key (Permissions -> RolePermissions)
RENAME_KEY_SQL = ("UPDATE [Permissions] SET [PermissionKey] = 'SALES_POS_TEST' "
                  "WHERE [PermissionKey] = 'SALES_POS'")
COUNT_RENAMED_SQL = ("SELECT COUNT(*) FROM [RolePermissions] "
                     "WHERE [PermissionKey] = 'SALES_POS_TEST'")
COUNT_OLD_KEY_SQL = ("SELECT COUNT(*) FROM [RolePermissions] "
                     "WHERE [PermissionKey] = 'SALES_POS'")
