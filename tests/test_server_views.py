"""Plan 5D (docs/71): the saved queries as SQL Server views (tools/gen_sql_views.py, modServerQueries).

  * the translation Access SQL -> T-SQL, one rule at a time;
  * every view of dist/sql/RetailStore_Queries.sql gives the same rows as its Access query on the demo data:
    in SQLite (the T-SQL functions given their SQL Server meaning) always, and on a real SQL Server when the
    tests have one (tests/sql_server_live.py: RETAIL_SQLSERVER), with the parameters sent by the MERGE of
    modServerQueries, in both sets of views ([dbo] Arabic, [en] English);
  * the code: the saved queries switch to the linked views with the link to the server and back with the link
    to Access, the parameters reach the server, the script runs with "Prepare the database".
"""
import datetime as dt
import glob
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
from access_sqlite import TODAY, day
import demo_data as DD
import gen_sql_views as V
import queries as Q
import sql_server_live as L
import vba_harness as H


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


SCRIPT = os.path.join(ROOT, "dist", "sql", "RetailStore_Queries.sql")
VBA = read("modServerQueries")
SQL = read("modSqlServer")


def body(access_sql, schema="dbo"):
    return V.view_body(Q.Query("qryX", "", access_sql), schema)


def proc(text, name):
    m = re.search(rf"^(?:Public |Private )?(?:Function|Sub) {name}\(.*?^End (?:Function|Sub)", text, re.M | re.S)
    assert m, name
    return m.group(0)


# --------------------------------------------------------------------------
# The translation
# --------------------------------------------------------------------------
class TranslationTests(unittest.TestCase):

    def test_functions(self):
        s = body("SELECT CCur(Nz(Sum(l.Debit), 0)) AS A, IIf(l.Debit > 0, 1, 2) AS B, CStr(l.AccountCode) AS C, "
                 "Round(l.Debit, 2) AS D, Year(e.EntryDate) AS E, Month(e.EntryDate) AS F "
                 "FROM JournalLines AS l INNER JOIN JournalEntries AS e ON l.EntryID = e.EntryID GROUP BY l.AccountCode")
        self.assertIn("CAST(COALESCE(Sum(l.Debit), 0) AS MONEY) AS A", s)
        self.assertIn("IIF(l.Debit > 0, 1, 2) AS B", s)
        self.assertIn("CAST(l.AccountCode AS NVARCHAR(255)) AS C", s)
        self.assertIn("ROUND(l.Debit, 2) AS D, YEAR(e.EntryDate) AS E, MONTH(e.EntryDate) AS F", s)
        self.assertIn("FROM [dbo].[JournalLines] AS l INNER JOIN [dbo].[JournalEntries] AS e", s)

    def test_dates(self):
        s = body("SELECT DateValue(h.InvoiceDate) AS D, DateDiff('d', h.InvoiceDate, Date()) AS N, "
                 "DateAdd('d', 1, h.InvoiceDate) AS T FROM SalesInvoices AS h")
        self.assertIn("DATEADD(day, DATEDIFF(day, 0, h.InvoiceDate), 0) AS D", s)
        self.assertIn("DATEDIFF(day, h.InvoiceDate, DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0)) AS N", s)
        self.assertIn("DATEADD(day, 1, h.InvoiceDate) AS T", s)

    def test_texts_numbers_and_true(self):
        s = body("SELECT 'المبيعات' AS T, h.Tax / h.TaxableAmount AS R, IIf(h.Tax > 0, True, False) AS B "
                 "FROM SalesInvoices AS h WHERE h.Notes = 'O''Brien'")
        self.assertIn("N'المبيعات' AS T", s)
        self.assertIn("h.Tax * 1e0 / h.TaxableAmount AS R", s, "Access divides in floating point")
        self.assertIn("IIF(h.Tax > 0, CAST(1 AS BIT), CAST(0 AS BIT)) AS B", s)
        self.assertIn("h.Notes = N'O''Brien'", s)

    def test_names(self):
        s = body("SELECT q.SaleCost AS Expenses, e.Amount FROM qrySaleCost AS q, Expenses AS e")
        self.assertIn("q.SaleCost AS Expenses,", s, "an alias that is a table name stays")
        self.assertIn("FROM [dbo].[qrySaleCost] AS q, [dbo].[Expenses] AS e", s)
        self.assertIn("FROM [en].[qrySaleCost]", body("SELECT * FROM qrySaleCost", "en"))
        self.assertIn("l.LineNumber AS [LineNo]", body("SELECT l.LineNumber AS LineNo FROM JournalLines AS l"))

    def test_english_names_and_texts(self):
        s = body("SELECT c.CustomerName, 'المبيعات' AS T FROM [@Customers] AS c", "en")
        self.assertIn("FROM [en].[qryLocCustomers] AS c", s)
        self.assertNotIn("المبيعات", s)
        self.assertIn("FROM [dbo].[Customers] AS c", body("SELECT c.CustomerName FROM [@Customers] AS c"))

    def test_yes_no_conditions(self):
        s = body("SELECT IIf(d.ReturnToStock, 1, 0) AS A, p.IsActive FROM SalesReturnDetails AS d, Products AS p "
                 "WHERE p.IsActive AND Not p.IsActive AND (p.IsActive) AND p.IsActive <> 0")
        self.assertIn("IIF(d.ReturnToStock <> 0, 1, 0) AS A, p.IsActive FROM", s, "a value stays a value")
        self.assertIn("WHERE p.IsActive <> 0 AND Not p.IsActive <> 0 AND (p.IsActive <> 0) AND p.IsActive <> 0", s)

    def test_parameters(self):
        s = body("SELECT l.AccountCode, Sum(IIf(e.EntryDate < QDate('PeriodStart'), l.Debit, 0)) AS Before, "
                 "QDate('PeriodStart') AS FromDay "
                 "FROM JournalEntries AS e INNER JOIN JournalLines AS l ON e.EntryID = l.EntryID "
                 "WHERE e.EntryDate < QDate('PeriodEnd') AND l.AccountCode = QLong('AccountCode') "
                 "GROUP BY l.AccountCode HAVING Sum(l.Debit) > QLong('AccountCode')")
        self.assertIn("ON e.EntryID = l.EntryID CROSS JOIN [dbo].[QueryParamsNow] AS qp", s)
        self.assertIn("Sum(IIF(e.EntryDate < qp.[qp_PeriodStart], l.Debit, 0)) AS Before", s, "in an aggregate")
        self.assertIn("(SELECT p.[qp_PeriodStart] FROM [dbo].[QueryParamsNow] AS p) AS FromDay", s,
                      "a column of a grouped query cannot be the joined row")
        self.assertIn("WHERE e.EntryDate < qp.[qp_PeriodEnd] AND l.AccountCode = qp.[qp_AccountCode]", s)
        self.assertIn("HAVING Sum(l.Debit) > (SELECT p.[qp_AccountCode]", s)

    def test_parameter_in_on_and_select_star(self):
        s = body("SELECT * FROM SalesInvoices AS h LEFT JOIN Customers AS c ON (c.CustomerID = h.CustomerID "
                 "AND h.InvoiceDate >= QDate('PeriodStart')) WHERE h.InvoiceDate < QDate('PeriodEnd')")
        self.assertIn("h.InvoiceDate >= (SELECT p.[qp_PeriodStart] FROM [dbo].[QueryParamsNow] AS p))", s,
                      "qp is not joined yet inside ON")
        self.assertIn("SELECT h.*, c.* FROM", s, "SELECT * would add the parameter columns")
        s = body("SELECT * FROM qrySalesLineItems WHERE DocDate >= QDate('PeriodStart')")
        self.assertIn("SELECT [dbo].[qrySalesLineItems].* FROM", s)

    def test_sub_queries_have_their_own_row(self):
        s = V.view_body(Q.query("DashboardQuery"))
        self.assertEqual(s.count("CROSS JOIN [dbo].[QueryParamsNow] AS qp"), s.count("qp.[qp_DashEnd]"))

    def test_unknown_function_is_refused(self):
        with self.assertRaises(ValueError):
            body("SELECT Format(h.InvoiceDate, 'yyyy') AS Y FROM SalesInvoices AS h")

    def test_order_by(self):
        self.assertEqual(V.server_order(Q.query("TrialBalanceQuery").sql), " ORDER BY [AccountCode]")
        self.assertEqual(V.server_order(Q.query("DailySalesQuery").sql), " ORDER BY [SaleDate] DESC",
                         "an expression of the column list: its alias")
        self.assertEqual(V.server_order(Q.query("LowStockQuery").sql), " ORDER BY [ShortageQty] DESC, [ProductName]")
        self.assertEqual(V.server_order("SELECT p.ProductName FROM Products AS p"), "")
        self.assertIsNone(V.server_order("SELECT p.ProductName FROM Products AS p ORDER BY p.CreatedAt"),
                          "not one of its columns: stays in Access")
        self.assertNotIn("ORDER BY", V.view_body(Q.query("TrialBalanceQuery")))

    def test_every_query_goes_to_the_server(self):
        self.assertEqual([n for n, o in V.server_query_list() if o is None], [])


# --------------------------------------------------------------------------
# The script and the generated module
# --------------------------------------------------------------------------
class ScriptTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        with open(SCRIPT, encoding="utf-8-sig") as fh:
            cls.text = fh.read()

    def test_dist_is_generated(self):
        self.assertEqual(self.text.replace("\r\n", "\n"), V.build_views_sql())

    def test_every_query_in_both_sets(self):
        for q in Q.QUERIES:
            for schema in V.SCHEMAS:
                self.assertIn(f"CREATE OR ALTER VIEW [{schema}].[{q.name}] AS", self.text)
        self.assertEqual(len(L.batches(self.text)), 2 * len(Q.QUERIES) + 5)

    def test_order_follows_the_queries(self):
        pos = [self.text.index(f"VIEW [dbo].[{q.name}] AS") for q in Q.QUERIES]
        self.assertEqual(pos, sorted(pos))
        self.assertLess(self.text.index("VIEW [dbo].[QueryParamsNow]"), pos[0])
        self.assertLess(self.text.index("CREATE SCHEMA [en]"), self.text.index("VIEW [en]."))

    def test_parameters_table(self):
        params = V.query_params()
        self.assertEqual(params["PeriodStart"], "D")
        self.assertEqual(params["AccountCode"], "L")
        for n, k in params.items():
            self.assertIn(f"ADD [{n}] {'DATETIME' if k == 'D' else 'INT'} NULL;", self.text)
            self.assertIn(f"AS [qp_{n}]", self.text)
        self.assertIn("p.HostName = HOST_NAME() AND p.HostProcess = HOST_ID()", self.text)
        self.assertIn("FROM (SELECT 1 AS One) AS o LEFT JOIN", self.text, "always one row")

    def test_version(self):
        data = read("modServerQueryData")
        self.assertIn(f'SERVER_QUERIES_VERSION As String = "{V.views_version()}"', data)
        self.assertIn(f"SELECT N'{V.views_version()}' AS Version;", self.text)
        self.assertIn(f'ServerParamList = "{",".join(f"{n}:{k}" for n, k in V.query_params().items())}"', data)

    def test_no_access_left(self):
        for word in ("CCur(", "Nz(", "IIf(", "QDate(", "QLong(", "DateValue(", "True", "False", "[@"):
            self.assertNotIn(word, self.text, word)
        self.assertNotIn("ORDER BY", self.text)


class Generated_modServerQueryData(VbaModuleChecks, unittest.TestCase):
    module_name = "modServerQueryData"
    vba = V.build_data_vba()


class Static_modServerQueries(VbaModuleChecks, unittest.TestCase):
    module_name = "modServerQueries"
    vba = VBA


# --------------------------------------------------------------------------
# The same rows as Access: SQLite (always) and a real SQL Server (when there is one)
# --------------------------------------------------------------------------
def _sqlite_date(v):
    if v == 0 or v is None:
        return dt.datetime(1900, 1, 1) if v == 0 else None
    return dt.datetime.fromisoformat(str(v)[:19]) if len(str(v)) > 10 else dt.datetime.fromisoformat(str(v)[:10])


def _dateadd(unit, n, v):
    # SQLite names are not case sensitive: these also serve the Access mirror's DateAdd('d', ...)
    d = _sqlite_date(v)
    assert unit in ("day", "d"), unit
    return None if d is None else (d + dt.timedelta(days=n)).strftime("%Y-%m-%d %H:%M:%S")


def _datediff(unit, a, b):
    assert unit in ("day", "d"), unit
    a, b = _sqlite_date(a), _sqlite_date(b)
    return None if a is None or b is None else (b.date() - a.date()).days


def to_sqlite(view_sql: str, prefix: str) -> str:
    """A view of the script for SQLite, with the SQL Server meaning of its functions (registered below)."""
    m = re.match(r"CREATE OR ALTER VIEW \[(\w+)\]\.\[(\w+)\] AS\n", view_sql)
    sql = view_sql[m.end():].rstrip().rstrip(";")
    sql = re.sub(r"\[(dbo|en)\]\.\[(\w+)\]", lambda x: f'"{x.group(2)}"' if x.group(2) in V.TABLE_NAMES
                 else f'"srv_params"' if x.group(2) == V.PARAMS_VIEW else f'"{prefix}{x.group(1)}_{x.group(2)}"', sql)
    sql = re.sub(r"\bN'", "'", sql)
    sql = re.sub(r"\b(DATEADD|DATEDIFF)\((\w+),", r"\1('\2',", sql)
    return f'CREATE VIEW "{prefix}{m.group(1)}_{m.group(2)}" AS {sql}'


def demo():
    """The demo data in the SQLite mirror of Access, with the parameters of every report set."""
    r = DD.simulate()
    db = r.db
    db.set_period(31)
    db.params.update({"CustomerID": r.ids["cust"][1], "SupplierID": r.ids["sup"][1], "ProductID": r.ids["prod"][1],
                      "DashDay": day(0), "DashEnd": day(-1), "DashMonth": day(30), "AccountCode": 0, "CashBoxID": 0,
                      "CompareStart": day(62), "CompareEnd": day(31), "IndEnd": day(-1), "IndMonth": day(30),
                      "IndPrevMonth": day(60), "IndYear": day(364), "Ind90": day(89)})
    return r


def param_value(db, name, kind):
    v = db.params.get(name)
    if kind == "D":
        return dt.datetime.fromisoformat(v) if v else None
    return int(v or 0)


def comparable(rows, names):
    """Rows as sorted tuples of plain values; RowVer (SQL Server only) left out; a date always with its time."""
    keep = [i for i, n in enumerate(names) if n != "RowVer"]
    out = []
    for row in rows:
        vals = []
        for i in keep:
            v = L.plain(row[i])
            if isinstance(v, str) and re.fullmatch(r"\d{4}-\d\d-\d\d", v):
                v += " 00:00:00"
            vals.append(v)
        out.append(tuple(vals))
    return sorted(out, key=repr)


TODAY_QUERIES = {q.name for q in Q.QUERIES if "Date()" in q.sql}       # SQL Server's today is the real one
MIRROR_ONLY = {"IntegrityCheckQuery"}           # rows of the mirror's floating point (exact money on the server)


class SameRowsInSqliteTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = demo()
        c = cls.c = cls.r.db.con
        c.create_function("DATEADD", 3, _dateadd)
        c.create_function("DATEDIFF", 3, _datediff)
        c.create_function("GETDATE", 0, lambda: dt.datetime.combine(TODAY, dt.time(12)).strftime("%Y-%m-%d %H:%M:%S"))
        c.create_function("YEAR", 1, lambda s: None if s is None else int(str(s)[:4]))
        c.create_function("MONTH", 1, lambda s: None if s is None else int(str(s)[5:7]))
        cols = ", ".join(f"? AS [qp_{n}]" for n in V.query_params())
        vals = []
        for n, k in V.query_params().items():
            v = cls.r.db.params.get(n)
            vals.append((v or "1899-12-30 00:00:00") if k == "D" else int(v or 0))
        c.execute(f'CREATE TABLE "srv_params" AS SELECT {cols}', vals)
        for schema in V.SCHEMAS:
            for q in Q.QUERIES:
                c.execute(to_sqlite(V.view_sql(q, schema), "srv_"))

    def rows(self, sql):
        cur = self.c.execute(sql)
        return comparable(cur.fetchall(), [d[0] for d in cur.description])

    def test_every_view_gives_the_access_rows(self):
        for q in Q.QUERIES:
            if q.name in MIRROR_ONLY:
                continue
            with self.subTest(q.name):
                self.assertEqual(self.rows(f'SELECT * FROM "srv_dbo_{q.name}"'), self.rows(f'SELECT * FROM "{q.name}"'))

    def test_the_demo_has_data_to_compare(self):
        self.assertGreater(len(self.rows('SELECT * FROM "srv_dbo_qryJournalSale"')), 40)
        self.assertGreater(len(self.rows('SELECT * FROM "srv_dbo_TrialBalanceQuery"')), 10)

    def test_english_views_same_numbers(self):
        for q in Q.QUERIES:
            with self.subTest(q.name):
                a = self.c.execute(f'SELECT COUNT(*) FROM "srv_dbo_{q.name}"').fetchone()[0]
                b = self.c.execute(f'SELECT COUNT(*) FROM "srv_en_{q.name}"').fetchone()[0]
                self.assertEqual(a, b)

    def test_a_wrong_translation_is_found(self):
        wrong = V.view_sql(Q.query("qryTrialPeriod"), "dbo").replace("< qp.[qp_PeriodEnd]", "< qp.[qp_PeriodStart]")
        self.c.execute(to_sqlite(wrong, "bad_"))
        self.assertNotEqual(self.rows('SELECT * FROM "bad_dbo_qryTrialPeriod"'),
                            self.rows('SELECT * FROM "qryTrialPeriod"'))


@unittest.skipIf(L.why_skipped(), L.why_skipped())
class SameRowsOnSqlServerTests(unittest.TestCase):
    """The script and the views on a real SQL Server, against the Access queries of the SQLite mirror."""

    DB = "RetailStore_Views_Test"

    @classmethod
    def setUpClass(cls):
        cls.r = demo()
        cls.con = L.fresh_database(cls.DB)
        cls.parts = L.run_script(cls.con, "dist/sql/RetailStore_Schema.sql")
        cls.views = L.run_script(cls.con, "dist/sql/RetailStore_Queries.sql")
        L.load(cls.con, cls.r.db.con)
        params = V.query_params()
        names = list(params)
        cls.con.cursor().execute(V.query_params_sql(names, [param_value(cls.r.db, n, params[n]) for n in names]))

    @classmethod
    def tearDownClass(cls):
        cls.con.close()
        L.drop_database(cls.DB)

    def server_rows(self, sql):
        cur = self.con.cursor()
        cur.execute(sql)
        return comparable(cur.fetchall(), [d[0] for d in cur.description])

    def access_rows(self, name):
        cur = self.r.db.con.execute(f'SELECT * FROM "{name}"')
        return comparable(cur.fetchall(), [d[0] for d in cur.description])

    def test_every_view_gives_the_access_rows(self):
        for q in Q.QUERIES:
            if q.name in TODAY_QUERIES | MIRROR_ONLY:
                continue
            with self.subTest(q.name):
                self.assertEqual(self.server_rows(f"SELECT * FROM [dbo].[{q.name}]"), self.access_rows(q.name))

    def test_views_that_read_today_run(self):
        for name in sorted(TODAY_QUERIES):
            with self.subTest(name):
                self.server_rows(f"SELECT * FROM [dbo].[{name}]")

    def test_english_views_same_numbers(self):
        cur = self.con.cursor()
        for q in Q.QUERIES:
            with self.subTest(q.name):
                cur.execute(f"SELECT COUNT(*) FROM [dbo].[{q.name}]")
                a = cur.fetchone()[0]
                cur.execute(f"SELECT COUNT(*) FROM [en].[{q.name}]")
                self.assertEqual(cur.fetchone()[0], a)

    def test_scripts_run_again(self):
        self.assertEqual(L.run_script(self.con, "dist/sql/RetailStore_Schema.sql"), self.parts)
        self.assertEqual(L.run_script(self.con, "dist/sql/RetailStore_Queries.sql"), self.views)
        cur = self.con.cursor()
        cur.execute("SELECT Version FROM [dbo].[QueriesVersion]")
        self.assertEqual(cur.fetchone()[0], V.views_version())

    def test_parameters_are_per_process(self):
        cur = self.con.cursor()
        cur.execute("SELECT COUNT(*) FROM [dbo].[QueryParams] WHERE HostName = HOST_NAME() AND HostProcess = HOST_ID()")
        self.assertEqual(cur.fetchone()[0], 1)
        cur.execute(V.query_params_sql(["AccountCode"], [1300]))            # the MERGE again: an update
        cur.execute("SELECT qp_AccountCode, qp_PeriodStart FROM [dbo].[QueryParamsNow]")
        code, start = cur.fetchone()
        self.assertEqual(code, 1300)
        self.assertEqual(start, param_value(self.r.db, "PeriodStart", "D"))
        cur.execute("UPDATE [dbo].[QueryParams] SET HostProcess = N'other'")  # another Access: not this one's
        cur.execute("SELECT qp_AccountCode, qp_PeriodStart FROM [dbo].[QueryParamsNow]")
        self.assertEqual(cur.fetchone(), (0, dt.datetime(1899, 12, 30)), "not set: like QDate / QLong")
        cur.execute("UPDATE [dbo].[QueryParams] SET HostProcess = HOST_ID()")
        cur.execute(V.query_params_sql(["AccountCode"], [0]))


# --------------------------------------------------------------------------
# The code
# --------------------------------------------------------------------------
QUERY_PARAMS = set(V.query_params())


class CodeTests(unittest.TestCase):

    def test_link_moves_the_queries(self):
        link = proc(SQL, "LinkToServer")
        self.assertIn("msg = ServerQueriesProblem(Connect)", link, "refused without the views of this version")
        self.assertGreater(link.index("UseServerQueries(Connect)"), link.index("RelinkTable db, names(i), Connect"))
        back = proc(SQL, "LinkToAccess")
        self.assertLess(back.index("UseAccessQueries()"), back.index("RelinkTable db, names(i)"))

    def test_queries_keep_their_access_sql(self):
        use = proc(VBA, "UseServerQueries")
        self.assertLess(use.index("KeepAccessSql qdf"), use.index("qdf.SQL = ServerQuerySql("))
        self.assertIn('If UiEnglish() Then schema = "en" Else schema = "dbo"', use)
        self.assertIn("SendAllQueryParams", use, "the period already chosen reaches the server")
        keep = proc(VBA, "KeepAccessSql")
        self.assertIn("If Left$(sql, Len(SERVER_MARK)) = SERVER_MARK Then Exit Sub", keep,
                      "never keeps its own server SQL")
        self.assertIn("dbMemo", keep)
        back = proc(VBA, "UseAccessQueries")
        self.assertIn("qdf.Properties.Delete ACCESS_SQL", back)
        self.assertIn("db.TableDefs.Delete db.TableDefs(i).Name", back)

    def test_build_queries_puts_them_back_on_the_server(self):
        b = read("modBuildQueries")
        self.assertIn('Application.Run "ServerQueriesAfterBuild"', proc(b, "BuildQueries"))
        self.assertIn("If Not IsLinkedToServer() Then Exit Sub", proc(VBA, "ServerQueriesAfterBuild"))

    def test_prepare_runs_both_scripts(self):
        s = proc(SQL, "SqlServerSchema")
        self.assertIn('QueriesScriptPath(Nz(frm!txtScript.Value, ""))', s)
        self.assertIn("RetailStore_Queries.sql", proc(SQL, "QueriesScriptPath"))

    def test_login_warns_of_another_version(self):
        self.assertIn("ServerQueriesWarning()", read("modSecurityScreens"))

    def test_parameters_reach_the_server(self):
        p = read("modQueryParams")
        self.assertIn("SendQueryParams names, values", proc(p, "SetQueryParams"))
        self.assertIn("SendQueryParams Names, values", proc(p, "RemoveQueryParams"))
        for name in ("SetPeriod", "SetComparePeriod", "SetQueryParam", "ClearQueryParams"):
            self.assertNotIn("TempVars.Add", proc(p, name), name)
        self.assertIn("If Not IsLinkedToServer() Then Exit Sub", proc(VBA, "SendQueryParams"))

    def test_no_parameter_set_outside_mod_query_params(self):
        """A parameter set with TempVars.Add elsewhere would not reach the server views."""
        bad = []
        paths = glob.glob(os.path.join(ROOT, "src", "vba", "*.bas")) + glob.glob(os.path.join(ROOT, "tools", "*.py"))
        for path in paths:
            if path.endswith("modQueryParams.bas"):
                continue
            with open(path, encoding="utf-8") as fh:
                for no, line in enumerate(fh, 1):
                    for m in re.finditer(r'TempVars\.(?:Add|Remove)\s+"(\w+)"', line):
                        if m.group(1) in QUERY_PARAMS:
                            bad.append(f"{os.path.basename(path)}:{no}: {line.strip()}")
        self.assertEqual(bad, [])

    def test_the_check_finds_one(self):
        line = '    TempVars.Add "PeriodStart", x'
        self.assertIn(re.search(r'TempVars\.(?:Add|Remove)\s+"(\w+)"', line).group(1), QUERY_PARAMS)

    def test_in_access_test_is_listed(self):
        self.assertIn('"TestServerQueries"', read("modTestAll"))


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class RuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modSqlServer": SQL, "modServerQueryData": read("modServerQueryData"), "modServerQueries": VBA,
                    "modProbe": 'Function ProbeParams() As String\n'
                                '    ProbeParams = QueryParamsSql(Array("PeriodStart", "AccountCode", "ReportCriteria", '
                                '"CustomerID"), Array(DateSerial(2026, 3, 1), 1300, "x", Null))\n'
                                'End Function\n'})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_mirror(self):
        self.assertEqual(self.h.call("modProbe", "ProbeParams"),
                         V.query_params_sql(["PeriodStart", "AccountCode", "ReportCriteria", "CustomerID"],
                                            [dt.datetime(2026, 3, 1), 1300, "x", None]))
        self.assertEqual(self.h.call("modServerQueries", "ServerQuerySql", "TrialBalanceQuery", " ORDER BY [AccountCode]"),
                         V.server_query_sql("TrialBalanceQuery", " ORDER BY [AccountCode]"))
        self.assertEqual(self.h.call("modServerQueries", "ServerParamKind", "IndYear"), "D")
        self.assertEqual(self.h.call("modServerQueries", "ServerParamKind", "Nothing"), "")


if __name__ == "__main__":
    unittest.main()
