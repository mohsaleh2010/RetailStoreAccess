"""English names of the master data (tools/master_en.py, Accounts.AccountNameEn, PaymentMethods.MethodNameEn,
JournalSourceTypes.TypeNameEn, TransactionTypes.TypeNameEn, qryLoc* and the [@Table] marker): BuildSchema fills
the empty English names, an English front-end reads qryLoc<Table> (same columns, English name) and an Arabic
one the table itself. The saved queries run on the SQLite mirror in both languages."""
import os
import re
import unittest

from helpers import ROOT
from access_sqlite import AccessOnSqlite
import forms as F
import generate as G
import i18n
import queries as Q
from master_en import ENGLISH_NAMES
from schema import table

MARK = re.compile(r"\[@(\w+)\]")


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def fill_english(con):
    """What modBuildSchema.SeedEnglishNames does."""
    for tbl, (en_field, _, key, names) in ENGLISH_NAMES.items():
        for k, en in names.items():
            con.execute(f'UPDATE "{tbl}" SET {en_field} = ? WHERE {key} = ? AND {en_field} IS NULL', (en, k))


class SchemaTests(unittest.TestCase):

    def test_every_seeded_row_has_an_english_name(self):
        for tbl, (en_field, ar_field, key, names) in ENGLISH_NAMES.items():
            t = table(tbl)
            fld = next(f for f in t.fields if f.name == en_field)
            self.assertFalse(fld.required, tbl)                       # BuildSchema upgrades an existing back-end
            keys = {row[t.seed_columns.index(key)] for row in t.seed_rows}
            self.assertEqual(set(names), keys, tbl)
            for k, en in names.items():
                with self.subTest(table=tbl, key=k):
                    self.assertLessEqual(len(en), fld.size)
                    en.encode("ascii")
                    self.assertNotIn("'", en)

    def test_build_schema_fills_only_empty_names(self):
        vba = read("modBuildSchema")
        body = re.search(r"Private Sub SeedEnglishNames\(\).*?End Sub", vba, re.S).group(0)
        self.assertEqual(body.count("m_db.Execute"), sum(len(v[3]) for v in ENGLISH_NAMES.values()))
        self.assertEqual(body.count("Is Null"), body.count("m_db.Execute"))
        self.assertLess(vba.index("    UpgradeAccountTree\n"), vba.index("    SeedEnglishNames\n"))

    def test_accounts_screen_has_the_english_name(self):
        model = next(m for m in F.all_forms() if m.name == "frmAccounts")
        self.assertIn("AccountNameEn", {c.name for c in model.controls})


class MarkerTests(unittest.TestCase):

    def all_sql(self):
        out = [q.sql for q in Q.QUERIES]
        for m in F.all_forms():
            out += [m.tag, m.record_source] + [v for c in m.controls for v in c.props.values() if isinstance(v, str)]
        for name in G.STATIC_MODULES:
            out += ["".join(i18n.vba_literals(st.replace(" _\n", " ")))
                    for st in re.split(r"(?<! _)\n", read(name))]
        return [s for s in out if s]

    def test_markers_name_a_table_with_english_names(self):
        for s in self.all_sql():
            for name in MARK.findall(s):
                self.assertIn(name, ENGLISH_NAMES, s[:100])

    def test_every_name_read_goes_through_the_marker(self):
        """A list or query that shows the name of an account, payment method, journal source or stock move
        type reads [@Table], so an English front-end shows the English name."""
        for s in self.all_sql():
            if s.lstrip().upper().startswith(("INSERT", "UPDATE", "DELETE")) or "AS LocArabicName" in s:
                continue                                 # writing, or the first step of qryLoc<Table>
            if "'SALE' AS DocKind" in s:
                continue                                 # qrySalesDocPrint: the tax invoice stays Arabic (ZATCA)
            if re.search(r"Name =\s*$|Name = \w", s):
                continue                                 # a look-up by the name the user typed (the table itself)
            for tbl, (_, ar_field, _, _) in ENGLISH_NAMES.items():
                for m in re.finditer(rf"(?<![@\w\[]){tbl}\s+AS\s+(\w+)", s):
                    with self.subTest(sql=s[:90]):
                        self.assertIsNone(re.search(rf"(?<![\w.]){m.group(1)}\.{ar_field}\b", s))
                if re.search(rf"\b(?:FROM|JOIN)\s+{tbl}\b(?!\s+AS)", s) and re.search(rf"(?<![\w.]){ar_field}\b", s):
                    with self.subTest(sql=s[:90]):
                        self.fail(f"reads {tbl}.{ar_field} without [@{tbl}] AS x: {s[:120]}")

    def test_resolution(self):
        sql = "SELECT a.AccountName FROM [@Accounts] AS a"
        self.assertEqual(i18n.resolve_names(sql), "SELECT a.AccountName FROM Accounts AS a")
        self.assertEqual(i18n.resolve_names(sql, True), "SELECT a.AccountName FROM qryLocAccounts AS a")
        names = [q.name for q in Q.QUERIES]
        for tbl in ENGLISH_NAMES:                     # created before the queries that read them
            self.assertLess(names.index(f"qryLoc{tbl}"), names.index("JournalLinesQuery"))
        lang = read("modLang")
        self.assertIn('If UiEnglish() Then out = out & "qryLoc"', lang)
        self.assertLess(lang.index("s = LangSql(s)"), lang.index("If Not UiEnglish() Then Exit Function"))


class EnglishDataTests(unittest.TestCase):

    def db(self, english):
        db = AccessOnSqlite(english=english)
        db.load_fixture()
        fill_english(db.con)
        db.set_period(400)
        return db

    def test_names_follow_the_language(self):
        for english, cash, sales in ((False, "النقدية بالخزينة والصناديق", "المبيعات"),
                                     (True, "Cash in treasury and boxes", "Sales")):
            c = self.db(english).con
            self.assertEqual(c.execute("SELECT AccountName FROM AccountTreeQuery WHERE AccountCode = 1100")
                             .fetchone()[0], cash)
            self.assertEqual(c.execute("SELECT AccountName FROM AccountTreeQuery WHERE AccountCode = 4100")
                             .fetchone()[0], sales)
            self.assertEqual(c.execute("SELECT TypeName FROM [@JournalSourceTypes] WHERE SourceType = 'SALE'"
                                       .replace("[@JournalSourceTypes]", "qryLocJournalSourceTypes" if english
                                                else "JournalSourceTypes")).fetchone()[0],
                             "Sales invoice" if english else "فاتورة بيع")

    def test_a_row_without_an_english_name_keeps_its_arabic_name(self):
        db = self.db(True)
        db.con.execute("INSERT INTO Accounts (AccountCode, AccountName, AccountType, ParentCode, IsActive, IsPosting, "
                       "IsSystem) VALUES (5990, 'مصروف خاص', 'EXPENSE', 53, 1, 1, 0)")
        self.assertEqual(db.con.execute("SELECT AccountName FROM qryLocAccounts WHERE AccountCode = 5990").fetchone()[0],
                         "مصروف خاص")
        db.con.execute("UPDATE Accounts SET AccountNameEn = 'Special' WHERE AccountCode = 5990")
        self.assertEqual(db.con.execute("SELECT AccountName FROM qryLocAccounts WHERE AccountCode = 5990").fetchone()[0],
                         "Special")

    def test_every_saved_query_runs_in_english(self):
        db = self.db(True)
        db.params.update({"CustomerID": 1, "SupplierID": 1, "ProductID": 1, "PayrollRunID": 1, "CommissionRunID": 1,
                          "DashDay": db.params["PeriodStart"], "DashEnd": db.params["PeriodEnd"],
                          "DashMonth": db.params["PeriodStart"]})
        for q in Q.QUERIES:
            for p in q.params:
                db.params.setdefault(p, db.params.get("PeriodStart") if "Date" in p or "Start" in p or "End" in p
                                     else 1)
            with self.subTest(q.name):
                db.con.execute(f'SELECT * FROM "{q.name}" LIMIT 1').fetchall()


if __name__ == "__main__":
    unittest.main()
