"""Python mirror of the pure parts of modSqlServer (plan 5B, docs/68), and the move of a back-end into a
database built from the SQL Server script pieces (tools/gen_sql.py), run on SQLite by the tests.

  sql_literal       = SqlLiteral       (a value of an Access field as a T-SQL literal)
  server_connect    = ServerConnect    (the ODBC connection string without a DSN)
  split_sql_batches = SplitSqlBatches  (the batches of a script, between the GO lines)
  insert_batch_sql  = InsertBatchSql   (one INSERT of several rows, with IDENTITY_INSERT)
"""

import datetime as dt
import re
import sqlite3
from typing import Dict, List, Tuple

import gen_sql as G
from relations import relations
from schema import TABLES

# schema kind -> the DAO type of the field in Access (what SqlLiteral receives)
DAO_TYPE = {"BOOL": 1, "BYTE": 2, "INT": 3, "LONG": 4, "AUTO": 4, "MONEY": 5, "QTY": 5, "RATE": 5,
            "DATE": 8, "DATETIME": 8, "TEXT": 10, "MEMO": 12}
BATCH_ROWS = 200


def _str_number(v) -> str:
    """VBA Trim$(Str$(v)): a dot, no exponent for ordinary amounts, no trailing .0."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return repr(v) if isinstance(v, float) else str(v)


def sql_literal(v, field_type: int) -> str:
    if v is None:
        return "NULL"
    if field_type == 1:
        return "1" if v else "0"
    if field_type == 8:
        if isinstance(v, str):
            v = dt.datetime.fromisoformat(v[:19])
        return "'" + v.strftime("%Y%m%d %H:%M:%S") + "'"
    if field_type in (10, 12):
        return "N'" + str(v).replace("'", "''") + "'"
    return _str_number(v)


def server_connect(server: str, database: str, user: str, password: str, driver: str) -> str:
    s = f"ODBC;DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
    s += "Trusted_Connection=Yes;" if not user.strip() else f"UID={user};PWD={password};"
    if "odbc driver 18" in driver.lower():
        s += "TrustServerCertificate=Yes;"
    return s + "APP=RetailStore;"


def split_sql_batches(text: str) -> List[str]:
    out, cur = [], ""
    for line in text.replace("\r\n", "\n").split("\n"):
        if line.strip().upper() == "GO":
            if cur.replace("\n", "").strip():
                out.append(cur)
            cur = ""
        else:
            cur += line + "\n"
    if cur.replace("\n", "").strip():
        out.append(cur)
    return out


def insert_batch_sql(table_name: str, columns: str, value_rows: str, has_identity: bool) -> str:
    s = f"SET IDENTITY_INSERT [dbo].[{table_name}] ON;\n" if has_identity else ""
    s += f"INSERT INTO [dbo].[{table_name}] ({columns}) VALUES {value_rows};\n"
    if has_identity:
        s += f"SET IDENTITY_INSERT [dbo].[{table_name}] OFF;\n"
    return s


# --------------------------------------------------------------------------
# The "server" database on SQLite, from the same generated pieces as the script
# --------------------------------------------------------------------------
def to_sqlite(sql: str) -> str:
    sql = sql.replace(" IDENTITY(1,1)", "").replace("NVARCHAR(MAX)", "TEXT")
    sql = sql.replace("CONVERT(DATETIME, CONVERT(DATE, GETDATE()))", "CURRENT_DATE").replace("GETDATE()", "CURRENT_TIMESTAMP")
    sql = re.sub(r"\bN'", "'", sql).replace("[dbo].", "")
    return sql.replace("LEN(", "LENGTH(")


def tsql_like(pattern, value) -> bool:
    """T-SQL LIKE (% any text, _ one character, [0-9] a class), case-insensitive like Arabic_100_CI_AS."""
    if pattern is None or value is None:
        return None
    rx, i = "", 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "[":
            j = pattern.index("]", i + 1)
            rx += "[" + pattern[i + 1:j] + "]"
            i = j + 1
            continue
        rx += ".*" if ch == "%" else "." if ch == "_" else re.escape(ch)
        i += 1
    return re.fullmatch(rx, str(value), re.I | re.S) is not None


_SUM = r"\[\w+\](?:[+\-]\[\w+\])*"


def exact_money(rule: str) -> str:
    """SQLite has no exact MONEY: a comparison of two field sums is made on values rounded to 4 decimals,
    as SQL Server (MONEY) and Access (Currency) compare them."""
    return re.sub(rf"({_SUM})(<=|>=|<>|=|<|>)({_SUM})", r"ROUND(\1, 4)\2ROUND(\3, 4)", rule)


def server_database() -> sqlite3.Connection:
    """Tables, defaults, CHECK rules, primary keys and relationships of the SQL Server script."""
    con = sqlite3.connect(":memory:")
    con.create_function("like", 2, tsql_like)            # x LIKE y calls like(y, x)
    for t in TABLES:
        cols = [to_sqlite(G.column_sql(t, f, inline_check=True)) for f in t.fields]
        cols.append(f"CONSTRAINT [PK_{t.name}] PRIMARY KEY ({', '.join('[' + c + ']' for c in t.pk)})")
        if t.rule:
            cols.append(f"CONSTRAINT [CK_{t.name}] CHECK ({exact_money(to_sqlite(G.table_rule_sql(t)))})")
        for r in relations():
            if r.child == t.name:
                cols.append(f"FOREIGN KEY ([{r.child_field}]) REFERENCES [{r.parent}] ([{r.parent_field}])")
        con.execute(f"CREATE TABLE [{t.name}] ({', '.join(cols)})")
    return con


def transfer(source: sqlite3.Connection, server: sqlite3.Connection) -> Dict[str, int]:
    """TransferAllData: every table emptied, then copied by INSERT batches built like CopyTable."""
    copied = {}
    for t in TABLES:
        server.execute(f"DELETE FROM [{t.name}]")
    for t in TABLES:
        kinds = {f.name: f.kind for f in t.fields}
        cur = source.execute(f'SELECT * FROM "{t.name}"')
        names = [d[0] for d in cur.description]
        identity = any(kinds.get(n) == "AUTO" for n in names)
        cols = ", ".join(f"[{n}]" for n in names)
        rows, n = [], 0
        for rec in cur.fetchall():
            rows.append("(" + ", ".join(sql_literal(v, DAO_TYPE[kinds[c]]) for c, v in zip(names, rec)) + ")")
            n += 1
        for i in range(0, len(rows), BATCH_ROWS):
            sql = insert_batch_sql(t.name, cols, ",\n".join(rows[i:i + BATCH_ROWS]), identity)
            body = "\n".join(line for line in sql.split("\n") if not line.startswith("SET IDENTITY_INSERT"))
            server.execute(to_sqlite(body))
        copied[t.name] = n
    return copied


def compare(source: sqlite3.Connection, server: sqlite3.Connection) -> List[Tuple[str, bool]]:
    """CompareTable: rows, the total of the amount fields, and the first character codes of the text fields."""
    out = []
    for t in TABLES:
        money = " + ".join(f'COALESCE(SUM("{f.name}"), 0)' for f in t.fields if DAO_TYPE[f.kind] == 5) or "0"
        texts = " + ".join(f"COALESCE(SUM(UNICODE(COALESCE(\"{f.name}\", '') || ' ')), 0)"
                           for f in t.fields if DAO_TYPE[f.kind] in (10, 12)) or "0"
        q = f'SELECT COUNT(*), {money}, {texts} FROM "{t.name}"'
        a, b = source.execute(q).fetchone(), server.execute(q).fetchone()
        out.append((t.name, a[0] == b[0] and abs((a[1] or 0) - (b[1] or 0)) < 0.00005 and a[2] == b[2]))
    return out
