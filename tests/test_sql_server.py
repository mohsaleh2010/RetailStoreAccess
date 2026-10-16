"""Plan 5A (docs/67): the SQL Server script of the back-end (tools/gen_sql.py -> dist/sql/RetailStore_Schema.sql).

SQL Server is not available here, so:
  * the script is checked as text: every table, column, rule, index, relationship and seed row is there,
    and every statement checks first, so the script can run again on an existing database;
  * the plain statements are parsed as T-SQL (sqlglot, optional);
  * the tables, rules, defaults, seed data and relationships are created in SQLite from the same
    generated pieces: the seed rows satisfy every rule and every relationship.
"""
import os
import re
import sqlite3
import unittest

from helpers import ROOT
import gen_sql as G
from schema import DROPPED_INDEXES, RULE_UPGRADES, TABLES, table
from relations import UNENFORCED, relations

try:
    import sqlglot
except ImportError:                                  # optional (requirements-dev.txt)
    sqlglot = None

PATH = os.path.join(ROOT, "dist", "sql", "RetailStore_Schema.sql")


def read_script():
    with open(PATH, "rb") as fh:
        return fh.read()


class GeneratedFileTests(unittest.TestCase):

    def test_up_to_date_utf8_bom_crlf(self):
        data = read_script()
        self.assertTrue(data.startswith(b"\xef\xbb\xbf"), "UTF-8 with BOM, so SSMS and sqlcmd read the Arabic")
        text = data[3:].decode("utf-8")
        self.assertNotIn("\n", text.replace("\r\n", ""))
        self.assertEqual(text.replace("\r\n", "\n").rstrip("\n"), G.build_sql().rstrip("\n"),
                         "run tools/generate.py")


class ScriptTextTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.sql = G.build_sql()

    def test_every_table_column_and_key(self):
        for t in TABLES:
            create = G.create_table_sql(t)
            self.assertIn(f"IF OBJECT_ID(N'dbo.{t.name}', N'U') IS NULL", create)
            self.assertIn(f"CONSTRAINT [PK_{t.name}] PRIMARY KEY ({', '.join('[' + c + ']' for c in t.pk)})", create)
            self.assertIn("[RowVer] ROWVERSION NOT NULL", create)
            for f in t.fields:
                with self.subTest(table=t.name, field=f.name):
                    col = G.column_sql(t, f, inline_check=False)
                    self.assertIn(col, create)
                    self.assertIn(f"IF COL_LENGTH(N'dbo.{t.name}', N'{f.name}') IS NULL", self.sql)
                    if f.kind == "AUTO":
                        self.assertIn("INT IDENTITY(1,1) NOT NULL", col)
                    elif f.required or f.kind == "BOOL":
                        self.assertIn(" NOT NULL", col)
                    else:
                        self.assertIn(" NULL", col)
                        self.assertNotIn("NOT NULL", col)
                    if f.kind == "TEXT":
                        self.assertIn(f"NVARCHAR({f.size})", col)

    def test_types(self):
        self.assertEqual(G.SQL_TYPES["MONEY"], "MONEY", "Access Currency: 4 decimals, linked back as Currency")
        self.assertEqual(G.SQL_TYPES["DATE"], "DATETIME")
        self.assertEqual(G.SQL_TYPES["MEMO"], "NVARCHAR(MAX)")
        self.assertEqual(G.SQL_TYPES["BOOL"], "BIT")

    def test_defaults(self):
        t = table("FixedAssets")
        f = {x.name: x for x in t.fields}
        self.assertEqual(G.default_sql(f["DepMethod"]), "N'SL'")
        self.assertEqual(G.default_sql(f["CreatedAt"]), "GETDATE()")
        self.assertEqual(G.default_sql(f["PurchaseDate"]), "CONVERT(DATETIME, CONVERT(DATE, GETDATE()))")
        self.assertEqual(G.default_sql(f["Cost"]), "0")
        self.assertEqual(G.default_sql(next(x for x in table("Products").fields if x.name == "IsActive")), "1")
        for t in TABLES:
            for x in t.fields:
                G.default_sql(x)                    # every default is translated (raises otherwise)

    def test_rules(self):
        def rule(tbl, name):
            return G.field_rule_sql(next(x for x in table(tbl).fields if x.name == name))
        self.assertEqual(rule("FixedAssets", "DepMethod"), "[DepMethod] IN (N'SL',N'DB')")
        self.assertEqual(rule("FixedAssets", "DepRate"), "([DepRate] >=0 AND [DepRate] <=1)")
        self.assertEqual(rule("RouteStops", "VisitWeekday"), "[VisitWeekday] BETWEEN 1 AND 7")
        self.assertEqual(rule("FixedAssets", "DisposalTo"), "([DisposalTo] IS NULL) OR ([DisposalTo] IN (N'NONE',N'BANK',N'CASHBOX'))")
        self.assertIn("[VATNumber] LIKE N'3[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]3'",
                      rule("Settings", "VATNumber"))
        self.assertEqual(G.table_rule_sql(table("FixedAssets")), "[SalvageValue]+[OpeningAccumDep]<=[Cost]")
        for t in TABLES:
            for x in [x for x in t.fields if x.rule]:
                with self.subTest(table=t.name, field=x.name):
                    r = G.field_rule_sql(x)
                    self.assertIn(f"[{x.name}]", r)
                    self.assertNotIn('"', r)
                    self.assertNotRegex(r, r"\b(And|Or|Like|Between|Is Null)\b")
                    self.assertIn(f"CONSTRAINT [CK_{t.name}_{x.name}] CHECK ({r})", self.sql)
            if t.rule:
                self.assertNotIn('"', G.table_rule_sql(t))
                self.assertIn(f"CONSTRAINT [CK_{t.name}] CHECK", self.sql)

    def test_loosened_rules_are_created_again(self):
        for tbl, name in RULE_UPGRADES:
            self.assertIn(f"IF OBJECT_ID(N'dbo.CK_{tbl}_{name}', N'C') IS NOT NULL\n"
                          f"    ALTER TABLE [dbo].[{tbl}] DROP CONSTRAINT [CK_{tbl}_{name}];", self.sql)

    def test_indexes(self):
        for t in TABLES:
            for ix in t.indexes:
                self.assertIn(f"WHERE name = N'{ix.name}' AND object_id = OBJECT_ID(N'dbo.{t.name}'))", self.sql)
        # Access lets a unique index hold many Nulls: only the filled values are unique
        self.assertIn("CREATE UNIQUE INDEX [UX_Barcode] ON [dbo].[Products] ([Barcode]) WHERE [Barcode] IS NOT NULL;",
                      self.sql)
        self.assertIn("CREATE UNIQUE INDEX [UX_RevalNumber] ON [dbo].[AssetRevaluations] ([RevalNumber]);", self.sql)
        for tbl, name in DROPPED_INDEXES:
            self.assertIn(f"DROP INDEX [{name}] ON [dbo].[{tbl}];", self.sql)
        for t in TABLES:
            kinds = {f.name: f.kind for f in t.fields}
            for ix in t.indexes:
                self.assertTrue(all(kinds[c] != "MEMO" for c in ix.fields), "NVARCHAR(MAX) cannot be indexed")

    def test_relationships(self):
        self.assertEqual(G.cascade_conflicts(), [], "SQL Server refuses two cascading paths (error 1785)")
        for r in relations():
            with self.subTest(fk=r.name):
                parent = table(r.parent)
                keys = [parent.pk] + [ix.fields for ix in parent.indexes if ix.unique]
                self.assertIn([r.parent_field], keys, "a foreign key references a primary key or a unique index")
                self.assertIn(f"IF OBJECT_ID(N'dbo.{r.name}', N'F') IS NULL", self.sql)
        self.assertIn("ALTER TABLE [dbo].[JournalLines] WITH CHECK ADD CONSTRAINT [FK_JournalLines_EntryID] FOREIGN KEY "
                      "([EntryID]) REFERENCES [dbo].[JournalEntries] ([EntryID]) ON DELETE CASCADE;", self.sql)
        for child, field in UNENFORCED:
            self.assertIn(f"ALTER TABLE [dbo].[{child}] DROP CONSTRAINT [FK_{child}_{field}];", self.sql)
        self.assertLess(self.sql.index("7. Relationships"), self.sql.rindex("FOREIGN KEY"))
        self.assertLess(self.sql.index("5. Seed data"), self.sql.index("7. Relationships"),
                        "the relationships come after the data")

    def test_seed_rows(self):
        for t in TABLES:
            if not t.seed_rows:
                continue
            with self.subTest(table=t.name):
                part = G.seed_sql(t)
                self.assertEqual(part.count(f"INSERT INTO [dbo].[{t.name}] "), len(t.seed_rows))
                identity = any(f.kind == "AUTO" and f.name in t.seed_columns for f in t.fields)
                self.assertEqual("SET IDENTITY_INSERT" in part, identity)
                if t.seed_missing:
                    self.assertIn(f"IF NOT EXISTS (SELECT 1 FROM [dbo].[{t.name}] WHERE [{t.seed_columns[0]}] = ", part)
                else:
                    self.assertIn(f"IF NOT EXISTS (SELECT 1 FROM [dbo].[{t.name}])\nBEGIN", part)
        perm = G.seed_sql(table("Permissions"))
        self.assertIn("IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])", perm, "a new permission goes to its roles")
        self.assertEqual(G.literal("2000-01-01", "DATE"), "'20000101'", "read the same in every language")
        self.assertEqual(G.literal("O'Brien"), "N'O''Brien'")

    def test_every_statement_checks_first(self):
        body = self.sql.split("SET XACT_ABORT ON;", 1)[1]
        for m in re.finditer(r"^(CREATE TABLE|ALTER TABLE .* ADD (?!CONSTRAINT \[CK_)|CREATE (UNIQUE )?INDEX)", body, re.M):
            self.fail(f"not guarded: {m.group(0)}")
        self.assertEqual(len(re.findall(r"\bBEGIN\b", body)), len(re.findall(r"\bEND\b", body)))
        self.assertNotIn("\t", self.sql)


@unittest.skipIf(sqlglot is None, "pip install sqlglot")
class TsqlParseTests(unittest.TestCase):
    """The plain statements of the script parse as T-SQL (the IF / BEGIN around them is not parsed)."""

    def parse(self, stmt):
        try:
            sqlglot.parse_one(stmt, read="tsql")
        except Exception as e:                       # noqa: BLE001
            self.fail(f"{e}\n{stmt[:300]}")

    def test_tables(self):
        for t in TABLES:
            self.parse(G.create_table_sql(t).split("\n", 1)[1])

    def test_rules_indexes_keys_seed_and_names(self):
        for t in TABLES:
            for st in G.checks_sql(t):
                self.parse(st.split("\n")[-1])
            for st in G.indexes_sql(t):
                self.parse(st.split("\n", 1)[1])
            for row in G.seed_sql(t).split("\n") if t.seed_rows else []:
                if row.strip().startswith("INSERT"):
                    self.parse(row.strip())
        for st in G.foreign_keys_sql():
            self.parse(st.split("\n", 1)[1])
        for st in G.english_names_sql()[:50]:
            self.parse(st)


class SqliteRunTests(unittest.TestCase):
    """The generated columns, defaults, rules and seed rows, run in SQLite: every seed row satisfies the
    rules, the required fields and the relationships of the script."""

    @staticmethod
    def lite(sql):
        sql = sql.replace(" IDENTITY(1,1)", "").replace("NVARCHAR(MAX)", "TEXT")
        sql = sql.replace("CONVERT(DATETIME, CONVERT(DATE, GETDATE()))", "CURRENT_DATE").replace("GETDATE()",
                                                                                              "CURRENT_TIMESTAMP")
        sql = re.sub(r"\bN'", "'", sql)
        return sql.replace("LEN(", "LENGTH(")

    def test_seed_data_satisfies_the_script(self):
        con = sqlite3.connect(":memory:")
        for t in TABLES:
            cols = [self.lite(G.column_sql(t, f, inline_check=True)) for f in t.fields]
            cols.append(f"CONSTRAINT [PK_{t.name}] PRIMARY KEY ({', '.join('[' + c + ']' for c in t.pk)})")
            if t.rule:
                cols.append(f"CONSTRAINT [CK_{t.name}] CHECK ({self.lite(G.table_rule_sql(t))})")
            for r in relations():
                if r.child == t.name:
                    cols.append(f"FOREIGN KEY ([{r.child_field}]) REFERENCES [{r.parent}] ([{r.parent_field}])")
            con.execute(f"CREATE TABLE [{t.name}] ({', '.join(cols)})")
        for t in TABLES:
            for line in (G.seed_sql(t).split("\n") if t.seed_rows else []):
                stmt = line.strip()
                if stmt.startswith(f"INSERT INTO [dbo].[{t.name}] "):     # not the grants of a new permission
                    with self.subTest(row=stmt[:80]):
                        con.execute(self.lite(stmt).replace("[dbo].", ""))
        self.assertEqual(con.execute("PRAGMA foreign_key_check").fetchall(), [])
        self.assertEqual(con.execute("SELECT COUNT(*) FROM Permissions").fetchone()[0],
                         len(table("Permissions").seed_rows))
        self.assertEqual(con.execute("SELECT DepMethod FROM FixedAssets").fetchall(), [])
        con.execute("INSERT INTO Employees (EmployeeID, EmployeeName, Username, RoleID) VALUES (9, 'x', 'x9', 1)")
        self.assertEqual(con.execute("SELECT IsActive FROM Employees WHERE EmployeeID = 9").fetchone()[0], 1,
                         "defaults")
        with self.assertRaises(sqlite3.IntegrityError):
            con.execute("INSERT INTO RouteStops (SalesRepID, CustomerID, VisitWeekday) VALUES (1, 1, 9)")


if __name__ == "__main__":
    unittest.main()
