"""A real SQL Server for the tests that need one (plan 5D, docs/71).

The tests that use it are skipped unless the environment variable RETAIL_SQLSERVER gives a server the tests may
create and drop databases on, and the Python package pymssql is installed:

    RETAIL_SQLSERVER="127.0.0.1;sa;password"         # server;user;password (SQL login)

A developer gets one with Docker (SQL Server Express on Linux, free):

    docker run -d --name mssql -e ACCEPT_EULA=Y -e MSSQL_SA_PASSWORD='Dev_Pass_2026!' -e MSSQL_PID=Express \\
        -p 1433:1433 mcr.microsoft.com/mssql/server:2022-latest
    pip install pymssql
    export RETAIL_SQLSERVER="127.0.0.1;sa;Dev_Pass_2026!"
"""
import datetime as dt
import decimal
import os
import re
import sqlite3
from typing import List, Optional

from helpers import ROOT
import sql_transfer_reference as R
from schema import TABLES

ENV = "RETAIL_SQLSERVER"


def settings() -> Optional[List[str]]:
    v = os.environ.get(ENV, "")
    parts = v.split(";")
    return parts if len(parts) == 3 and parts[0] else None


def why_skipped() -> str:
    """"" when a server is there, else why the tests are skipped."""
    if not settings():
        return f"no SQL Server ({ENV} not set, see tests/sql_server_live.py)"
    try:
        import pymssql  # noqa: F401
    except ImportError:
        return "pymssql is not installed"
    try:
        connect("master").close()
    except Exception as e:                               # noqa: BLE001 - any connection failure skips
        return f"SQL Server {settings()[0]} does not answer: {e}"
    return ""


def connect(database: str):
    import pymssql
    server, user, password = settings()
    return pymssql.connect(server=server, user=user, password=password, database=database, autocommit=True,
                           login_timeout=10)


def fresh_database(name: str):
    """An empty database of that name (dropped first if it is there); returns a connection to it."""
    m = connect("master")
    cur = m.cursor()
    cur.execute(f"IF DB_ID(N'{name}') IS NOT NULL BEGIN ALTER DATABASE [{name}] SET SINGLE_USER WITH ROLLBACK "
                f"IMMEDIATE; DROP DATABASE [{name}]; END")
    cur.execute(f"CREATE DATABASE [{name}]")
    m.close()
    return connect(name)


def drop_database(name: str):
    m = connect("master")
    m.cursor().execute(f"IF DB_ID(N'{name}') IS NOT NULL BEGIN ALTER DATABASE [{name}] SET SINGLE_USER WITH "
                       f"ROLLBACK IMMEDIATE; DROP DATABASE [{name}]; END")
    m.close()


def batches(script_text: str) -> List[str]:
    """The script split at GO, as modSqlServer.SplitSqlBatches does."""
    return [b for b in re.split(r"^\s*GO\s*$", script_text.lstrip("﻿"), flags=re.M | re.I) if b.strip()]


def run_script(con, path: str) -> int:
    with open(os.path.join(ROOT, path), encoding="utf-8-sig") as fh:
        parts = batches(fh.read())
    cur = con.cursor()
    for i, b in enumerate(parts, 1):
        try:
            cur.execute(b)
        except Exception as e:                           # noqa: BLE001 - name the part that failed
            raise AssertionError(f"{path} part {i} of {len(parts)}: {e}\n{b[:400]}") from e
    return len(parts)


def load(con, source: sqlite3.Connection):
    """TransferAllData: the rows of every table, with their numbers (IDENTITY_INSERT), constraints off meanwhile."""
    cur = con.cursor()
    for t in TABLES:
        cur.execute(f"ALTER TABLE [dbo].[{t.name}] NOCHECK CONSTRAINT ALL;")
    for t in TABLES:
        cur.execute(f"DELETE FROM [dbo].[{t.name}];")
    for t in TABLES:
        kinds = {f.name: f.kind for f in t.fields}
        rs = source.execute(f'SELECT * FROM "{t.name}"')
        names = [d[0] for d in rs.description]
        identity = any(kinds.get(n) == "AUTO" for n in names)
        cols = ", ".join(f"[{n}]" for n in names)
        rows = ["(" + ", ".join(R.sql_literal(v, R.DAO_TYPE[kinds[c]]) for c, v in zip(names, rec)) + ")"
                for rec in rs.fetchall()]
        for i in range(0, len(rows), R.BATCH_ROWS):
            cur.execute(R.insert_batch_sql(t.name, cols, ",\n".join(rows[i:i + R.BATCH_ROWS]), identity))
    for t in TABLES:
        cur.execute(f"ALTER TABLE [dbo].[{t.name}] WITH CHECK CHECK CONSTRAINT ALL;")


def plain(v):
    """A value as the SQLite mirror gives it: numbers rounded to the halala, dates as text."""
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (decimal.Decimal, float)):
        v = round(float(v), 2)
        return 0.0 if v == 0 else v
    if isinstance(v, int):
        return float(v)
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, str) and re.fullmatch(r"-?\d+(\.\d+)?", v):
        return round(float(v), 2)
    return v
