"""Run Access-dialect queries on SQLite: registers Access/VBA functions used by
the saved queries (Nz, CCur, IIf, DateValue, ...) and the modQueryParams
functions (QDate, QLong), loads the fixture and evaluates checks."""

import datetime as dt
import re

from helpers import build_sqlite
import queries as Q
from schema import table

TODAY = dt.date(2026, 10, 3)
ACCESS_ZERO_DATE = "1899-12-30 00:00:00"


def fmt(d: dt.datetime) -> str:
    return d.strftime("%Y-%m-%d %H:%M:%S")


def day(days_ago, hour=0) -> str:
    return fmt(dt.datetime.combine(TODAY, dt.time(hour)) - dt.timedelta(days=days_ago))


def _date(s):
    return dt.date.fromisoformat(str(s)[:10])


class AccessOnSqlite:

    def __init__(self):
        self.con = build_sqlite(with_relationship_rules=True)
        self.params = {}
        self.ids = {}
        self._register_functions()
        for q in Q.QUERIES:
            self.con.execute(f'CREATE VIEW "{q.name}" AS {q.sql}')
        self.set_period(Q.PERIOD_START_DAYS_AGO)

    # ---------------------------------------------------------------- setup
    def _register_functions(self):
        c = self.con
        c.create_function("Nz", 2, lambda v, d: d if v is None else v)
        c.create_function("CCur", 1, lambda v: None if v is None else round(float(v), 4))
        c.create_function("CLng", 1, lambda v: None if v is None else int(round(float(v))))
        c.create_function("DateValue", 1, lambda s: None if s is None else str(s)[:10] + " 00:00:00")
        c.create_function("Year", 1, lambda s: None if s is None else int(str(s)[:4]))
        c.create_function("Month", 1, lambda s: None if s is None else int(str(s)[5:7]))
        c.create_function("Date", 0, lambda: fmt(dt.datetime.combine(TODAY, dt.time())))

        def date_add(interval, n, s):
            assert interval == "d"
            return fmt(dt.datetime.fromisoformat(str(s)) + dt.timedelta(days=n))

        def date_diff(interval, a, b):
            assert interval == "d"
            return (_date(b) - _date(a)).days

        c.create_function("DateAdd", 3, date_add)
        c.create_function("DateDiff", 3, date_diff)
        c.create_function("QDate", 1, lambda n: self.params.get(n, ACCESS_ZERO_DATE))
        c.create_function("QLong", 1, lambda n: int(self.params.get(n, 0) or 0))

    def set_period(self, days_ago):
        # same as SetPeriod(Date - days_ago, Date): end is exclusive (tomorrow 00:00)
        self.params["PeriodStart"] = day(days_ago)
        self.params["PeriodEnd"] = day(-1)

    def load_fixture(self):
        for row in Q.FIXTURE:
            cols = list(row.values)
            vals = [self.value(v) for v in row.values.values()]
            cur = self.con.execute(
                f'INSERT INTO "{row.table}" ({", ".join(cols)}) '
                f'VALUES ({", ".join("?" for _ in vals)})', vals)
            self.ids[row.key] = cur.lastrowid

    def value(self, v):
        if isinstance(v, Q.Ref):
            return self.ids[v.key]
        if isinstance(v, Q.Day):
            return day(v.days_ago, v.hour)
        if isinstance(v, bool):
            return 1 if v else 0
        return v

    # ---------------------------------------------------------------- run
    def sql(self, access_sql):
        def sub(m):
            kind, arg = m.group(1), m.group(2)
            if kind == "ref":
                return str(self.ids[arg])
            return "'" + day(int(arg)) + "'"
        s = re.sub(r"\{(ref|day):(\w+)\}", sub, access_sql)
        m = re.match(r"^SELECT TOP (\d+) (.*)$", s, re.S)
        if m:
            s = f"SELECT {m.group(2)} LIMIT {m.group(1)}"
        return s

    def scalar(self, access_sql):
        return self.con.execute(self.sql(access_sql)).fetchone()[0]

    def expected(self, check):
        e = check.expected
        if isinstance(e, str) and e.startswith("ref:"):
            return self.ids[e[4:]]
        return e

    def run_check(self, check):
        for name, key in (check.params or {}).items():
            self.params[name] = self.ids[key]
        return self.scalar(check.sql), self.expected(check)


def pk_of(table_name):
    return table(table_name).pk[0]
