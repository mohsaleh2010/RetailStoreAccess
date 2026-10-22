"""Plan 5A (docs/67): the SQL Server (T-SQL) script of the back-end, from the same definition as BuildSchema.

One script both creates a new database and upgrades an existing one, like BuildSchema does in Access:
every statement checks first (a table, a column, an index, a constraint, a seed row), so it can run
any number of times. Steps, in order:

  1. the tables that are missing (columns, primary key, defaults, CHECK rules);
  2. the columns an existing table is missing (with their default and rule);
  3. the CHECK rules, dropped and created again (a rule the schema loosened takes effect);
  4. the indexes (a unique index skips Nulls, as Access does), and the ones the schema dropped;
  5. the seed data: an empty table gets all its rows, a seed_missing table the rows it lacks;
  6. the English names of the system rows (only the empty ones);
  7. the relationships (FOREIGN KEY), after the data; the retired ones are dropped.

Types: AUTO INT IDENTITY, LONG INT, INT SMALLINT, BYTE TINYINT, MONEY / QTY / RATE MONEY (the Access
Currency type: 4 decimals, linked back as Currency), DATE / DATETIME DATETIME (linked as Date/Time by every
ODBC driver), BOOL BIT NOT NULL, TEXT NVARCHAR(size), MEMO NVARCHAR(MAX). Every table gets a ROWVERSION
column (RowVer): Access uses it to see whether a linked row changed, without it a row with a BIT or a
MONEY column can give "write conflict" errors.
"""

import re
from typing import List

from schema import DROPPED_INDEXES, RULE_UPGRADES, TABLES, Field, Table, table
from relations import UNENFORCED, relations

SQL_TYPES = {"AUTO": "INT", "LONG": "INT", "INT": "SMALLINT", "BYTE": "TINYINT", "MONEY": "MONEY",
             "QTY": "MONEY", "RATE": "MONEY", "DATE": "DATETIME", "DATETIME": "DATETIME", "BOOL": "BIT",
             "TEXT": "NVARCHAR", "MEMO": "NVARCHAR(MAX)"}
ROWVERSION = "RowVer"
GO = "GO"


# --------------------------------------------------------------------------
# Values, defaults and rules
# --------------------------------------------------------------------------
def q(name: str) -> str:
    return f"[{name}]"


def obj(name: str) -> str:
    return f"[dbo].[{name}]"


def literal(v, kind: str = "") -> str:
    """A seed value as T-SQL. A date written yyyy-mm-dd becomes yyyymmdd, which SQL Server reads the same
    whatever the language of the login."""
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):
        return repr(v)
    s = str(v)
    if kind in ("DATE", "DATETIME") and re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return "'" + s.replace("-", "") + "'"
    return "N'" + s.replace("'", "''") + "'"


def sql_type(f: Field) -> str:
    t = SQL_TYPES[f.kind]
    if f.kind == "TEXT":
        t = f"NVARCHAR({f.size})"
    return t


def default_sql(f: Field):
    """The Access default of a field as a T-SQL expression (None = no default)."""
    d = f.default
    if f.kind == "BOOL":
        return "1" if (d or "").strip().lower() == "true" else "0"
    if d is None or f.kind == "AUTO":
        return None
    d = d.strip()
    if d == "Now()":
        return "GETDATE()"
    if d == "Date()":
        return "CONVERT(DATETIME, CONVERT(DATE, GETDATE()))"
    if re.fullmatch(r'"[^"]*"', d):
        return literal(d[1:-1])
    if re.fullmatch(r"-?\d+(\.\d+)?", d):
        return d
    raise ValueError(f"default not translated: {f.name} = {d}")


def _strings(expr: str) -> str:
    """Access "text" literals -> N'text'."""
    return re.sub(r'"([^"]*)"', lambda m: "N'" + m.group(1).replace("'", "''") + "'", expr)


def _like(pattern: str) -> str:
    """An Access Like pattern: # one digit, ? one character, * any text."""
    out = pattern.replace("[", "[[]").replace("%", "[%]").replace("_", "[_]")
    return out.replace("#", "[0-9]").replace("?", "_").replace("*", "%")


def _split_top(expr: str, word: str) -> List[str]:
    """Split on a word (And / Or) outside parentheses and quotes."""
    parts, depth, cur, i, inq = [], 0, "", 0, False
    token = f" {word} "
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            inq = not inq
        elif not inq and ch == "(":
            depth += 1
        elif not inq and ch == ")":
            depth -= 1
        if not inq and depth == 0 and expr[i:i + len(token)].lower() == token.lower():
            parts.append(cur)
            cur, i = "", i + len(token)
            continue
        cur += ch
        i += 1
    parts.append(cur)
    return [p.strip() for p in parts]


def _words(expr: str) -> str:
    """Access keywords and functions -> T-SQL."""
    expr = re.sub(r"\bIs Not Null\b", "IS NOT NULL", expr, flags=re.I)
    expr = re.sub(r"\bIs Null\b", "IS NULL", expr, flags=re.I)
    expr = re.sub(r"\bAnd\b", "AND", expr)
    expr = re.sub(r"\bOr\b", "OR", expr)
    expr = re.sub(r"\bIn \(", "IN (", expr)
    expr = re.sub(r"\bBetween\b", "BETWEEN", expr)
    expr = re.sub(r"\bLen\(", "LEN(", expr)
    return expr


def field_rule_sql(f: Field) -> str:
    """An Access field rule (written without the field: ">=0 And <1") as a CHECK on the field."""
    rule = re.sub(r"\bBetween (\S+) And (\S+)", lambda m: f"Between {m.group(1)} \x00 {m.group(2)}", f.rule)      # its And is not a clause
    clauses = []
    for part_or in _split_top(rule, "Or"):
        ands = []
        for part in _split_top(part_or, "And"):
            p = part.replace("\x00", "AND")
            if re.match(r"(<>|>=|<=|=|>|<)", p) or re.match(r"(Between|In \(|Is )", p, re.I):
                p = f"{q(f.name)} {p}"
            elif p.lower().startswith("like "):
                pattern = re.fullmatch(r'Like "([^"]*)"', p, re.I).group(1)
                p = f"{q(f.name)} LIKE N'{_like(pattern)}'"
            ands.append(_words(_strings(p)))
        clauses.append(" AND ".join(ands) if len(ands) == 1 else "(" + " AND ".join(ands) + ")")
    return clauses[0] if len(clauses) == 1 else " OR ".join(c if c.startswith("(") else f"({c})" for c in clauses)


def table_rule_sql(t: Table) -> str:
    return _words(_strings(t.rule))


# --------------------------------------------------------------------------
# Statements
# --------------------------------------------------------------------------
def column_sql(t: Table, f: Field, inline_check: bool) -> str:
    out = f"{q(f.name)} {sql_type(f)}"
    if f.kind == "AUTO":
        return out + " IDENTITY(1,1) NOT NULL"
    out += " NOT NULL" if (f.required or f.kind == "BOOL") else " NULL"
    d = default_sql(f)
    if d is not None:
        out += f" CONSTRAINT {q(f'DF_{t.name}_{f.name}')} DEFAULT ({d})"
    if inline_check and f.rule:
        out += f" CONSTRAINT {q(f'CK_{t.name}_{f.name}')} CHECK ({field_rule_sql(f)})"
    return out


def create_table_sql(t: Table) -> str:
    cols = [column_sql(t, f, inline_check=False) for f in t.fields]
    cols.append(f"{q(ROWVERSION)} ROWVERSION NOT NULL")
    cols.append(f"CONSTRAINT {q('PK_' + t.name)} PRIMARY KEY ({', '.join(q(c) for c in t.pk)})")
    body = ",\n        ".join(cols)
    return (f"IF OBJECT_ID(N'dbo.{t.name}', N'U') IS NULL\n"
            f"    CREATE TABLE {obj(t.name)} (\n        {body}\n    );")


def add_columns_sql(t: Table) -> List[str]:
    out = []
    for f in t.fields + [None]:
        if f is None:
            name, col = ROWVERSION, f"{q(ROWVERSION)} ROWVERSION NOT NULL"
        else:
            name, col = f.name, column_sql(t, f, inline_check=False)
        out.append(f"IF COL_LENGTH(N'dbo.{t.name}', N'{name}') IS NULL\n"
                   f"    ALTER TABLE {obj(t.name)} ADD {col};")
    return out


def checks_sql(t: Table) -> List[str]:
    """Each rule dropped and created again, so a loosened rule (schema.RULE_UPGRADES) is the one in force."""
    out = []
    rules = [(f"CK_{t.name}_{f.name}", field_rule_sql(f)) for f in t.fields if f.rule]
    if t.rule:
        rules.append((f"CK_{t.name}", table_rule_sql(t)))
    for name, expr in rules:
        out.append(f"IF OBJECT_ID(N'dbo.{name}', N'C') IS NOT NULL\n"
                   f"    ALTER TABLE {obj(t.name)} DROP CONSTRAINT {q(name)};\n"
                   f"ALTER TABLE {obj(t.name)} WITH CHECK ADD CONSTRAINT {q(name)} CHECK ({expr});")
    return out


def indexes_sql(t: Table) -> List[str]:
    required = {f.name for f in t.fields if f.required or f.kind in ("AUTO", "BOOL")}
    out = []
    for ix in t.indexes:
        cols = ", ".join(q(c) for c in ix.fields)
        where = ""
        if ix.unique:
            # Access lets a unique index hold many Nulls; SQL Server would allow only one
            nullable = [c for c in ix.fields if c not in required]
            if nullable:
                where = " WHERE " + " AND ".join(f"{q(c)} IS NOT NULL" for c in nullable)
        kind = "UNIQUE INDEX" if ix.unique else "INDEX"
        out.append(f"IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'{ix.name}' "
                   f"AND object_id = OBJECT_ID(N'dbo.{t.name}'))\n"
                   f"    CREATE {kind} {q(ix.name)} ON {obj(t.name)} ({cols}){where};")
    return out


def dropped_indexes_sql() -> List[str]:
    return [f"IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'{name}' AND object_id = OBJECT_ID(N'dbo.{tbl}'))\n"
            f"    DROP INDEX {q(name)} ON {obj(tbl)};" for tbl, name in DROPPED_INDEXES]


def seed_sql(t: Table) -> str:
    kinds = {f.name: f.kind for f in t.fields}
    cols = ", ".join(q(c) for c in t.seed_columns)
    identity = any(kinds[c] == "AUTO" for c in t.seed_columns)
    grants = {}
    if t.name == "Permissions":
        for role, key in table("RolePermissions").seed_rows:
            grants.setdefault(key, []).append(role)
    rows = []
    for row in t.seed_rows:
        vals = ", ".join(literal(v, kinds[c]) for c, v in zip(t.seed_columns, row))
        ins = f"INSERT INTO {obj(t.name)} ({cols}) VALUES ({vals});"
        if t.seed_missing:
            key = t.seed_columns[0]
            if grants.get(row[0]):
                # a permission an upgrade adds goes to its default roles (BuildSchema: GrantNewPermission);
                # a new database gets all the grants from the RolePermissions seed
                give = "\n".join(f"            INSERT INTO [dbo].[RolePermissions] ([RoleID], [PermissionKey]) "
                                 f"VALUES ({role}, {literal(row[0])});" for role in grants[row[0]])
                ins = (f"{ins}\n        IF EXISTS (SELECT 1 FROM [dbo].[RolePermissions])\n        BEGIN\n{give}\n"
                       f"        END")
                ins = f"BEGIN\n        {ins}\n    END"
            ins = (f"IF NOT EXISTS (SELECT 1 FROM {obj(t.name)} WHERE {q(key)} = {literal(row[0], kinds[key])})\n"
                   + ("    " if ins.startswith("BEGIN") else "        ") + ins)
        rows.append("    " + ins)
    body = "\n".join(rows)
    head = "" if t.seed_missing else f"IF NOT EXISTS (SELECT 1 FROM {obj(t.name)})\n"
    out = [f"-- {t.name}: {len(t.seed_rows)} rows" + (" (the missing ones)" if t.seed_missing else " (an empty table)")]
    if identity:
        out.append(f"SET IDENTITY_INSERT {obj(t.name)} ON;")
    out.append(f"{head}BEGIN\n{body}\nEND;")
    if identity:
        out.append(f"SET IDENTITY_INSERT {obj(t.name)} OFF;")
    return "\n".join(out)


def english_names_sql() -> List[str]:
    from master_en import ENGLISH_NAMES, EXTRA_NAMES
    kinds = {t.name: {f.name: f.kind for f in t.fields} for t in TABLES}
    out = []
    for tbl, (en_field, _, key, names) in ENGLISH_NAMES.items():
        for k, en in names.items():
            out.append(f"UPDATE {obj(tbl)} SET {q(en_field)} = {literal(en)} WHERE {q(key)} = "
                       f"{literal(k, kinds[tbl][key])} AND {q(en_field)} IS NULL;")
        for en_field2, ar_field, by_arabic in EXTRA_NAMES.get(tbl, []):
            for ar, en in by_arabic.items():
                out.append(f"UPDATE {obj(tbl)} SET {q(en_field2)} = {literal(en)} WHERE {q(ar_field)} = "
                           f"{literal(ar)} AND {q(en_field2)} IS NULL;")
    return out


def foreign_keys_sql() -> List[str]:
    out = []
    for r in relations():
        actions = ""
        if r.cascade_delete:
            actions += " ON DELETE CASCADE"
        if r.cascade_update:
            actions += " ON UPDATE CASCADE"
        out.append(f"IF OBJECT_ID(N'dbo.{r.name}', N'F') IS NULL\n"
                   f"    ALTER TABLE {obj(r.child)} WITH CHECK ADD CONSTRAINT {q(r.name)} FOREIGN KEY ({q(r.child_field)}) "
                   f"REFERENCES {obj(r.parent)} ({q(r.parent_field)}){actions};")
    for child, field in sorted(UNENFORCED):
        name = f"FK_{child}_{field}"
        out.append(f"IF OBJECT_ID(N'dbo.{name}', N'F') IS NOT NULL\n"
                   f"    ALTER TABLE {obj(child)} DROP CONSTRAINT {q(name)};")
    return out


def cascade_conflicts() -> List[str]:
    """SQL Server refuses a table reached by two cascading paths from one table, or a cascade cycle
    (error 1785); Access accepts both. Returns the problems (none expected)."""
    problems = []
    for action in ("cascade_delete", "cascade_update"):
        edges = {}
        for r in relations():
            if getattr(r, action) and r.parent != r.child:
                edges.setdefault(r.parent, []).append(r.child)
        for root in edges:
            seen, stack = {}, [(root, (root,))]
            while stack:
                node, path = stack.pop()
                for child in edges.get(node, []):
                    if child in path:
                        problems.append(f"{action}: cycle {' > '.join(path + (child,))}")
                        continue
                    if child in seen:
                        problems.append(f"{action}: {child} reached twice from {root}")
                        continue
                    seen[child] = path
                    stack.append((child, path + (child,)))
    return problems


# --------------------------------------------------------------------------
# The script
# --------------------------------------------------------------------------
HEADER = """/*==============================================================================
  RetailStore - the back-end on SQL Server (plan 5A, docs/67-SQL-Server-Script.md)

  GENERATED FILE - do not edit by hand.
  Source of truth: tools/schema.py  ->  python3 tools/generate.py

  Creates a new database or upgrades an existing one (like BuildSchema in Access):
  every statement checks first, so the script can run again after each update.

  First time (SQL Server Management Studio, or sqlcmd):
      CREATE DATABASE RetailStore COLLATE Arabic_100_CI_AS;
  Then run this file in that database:
      sqlcmd -S .\\SQLEXPRESS -d RetailStore -E -f 65001 -i RetailStore_Schema.sql

  @@COUNTS@@
==============================================================================*/
SET NOCOUNT ON;
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET XACT_ABORT ON;
GO"""


# Readers see the last committed rows instead of waiting for another user's posting transaction (the program
# keeps a document and its stock moves and journal in one transaction). Set once: WITH ROLLBACK IMMEDIATE ends
# the other connections, which Access opens again.
DATABASE_OPTIONS = """IF EXISTS (SELECT 1 FROM sys.databases WHERE database_id = DB_ID() AND is_read_committed_snapshot_on = 0)
    ALTER DATABASE CURRENT SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;"""


def section(title: str) -> str:
    return f"\n/*------------------------------------------------------------------------------\n  {title}\n" \
           f"------------------------------------------------------------------------------*/"


def build_sql() -> str:
    problems = cascade_conflicts()
    if problems:
        raise ValueError("SQL Server would refuse these cascades: " + "; ".join(problems))
    rels = relations()
    counts = (f"{len(TABLES)} tables, {sum(len(t.fields) for t in TABLES)} fields, "
              f"{sum(len(t.indexes) for t in TABLES)} indexes, {len(rels)} relationships, "
              f"{sum(len(t.seed_rows) for t in TABLES)} seed rows")
    parts = [HEADER.replace("@@COUNTS@@", counts)]
    parts.append(section("1. Tables"))
    for t in TABLES:
        parts.append(create_table_sql(t))
    parts.append(GO)
    parts.append(section("2. Columns an existing table is missing"))
    for t in TABLES:
        parts.extend(add_columns_sql(t))
    parts.append(GO)
    parts.append(section("3. Rules (CHECK), created again on every run"))
    for t in TABLES:
        parts.extend(checks_sql(t))
    parts.append(GO)
    parts.append(section("4. Indexes"))
    for t in TABLES:
        parts.extend(indexes_sql(t))
    parts.extend(dropped_indexes_sql())
    parts.append(GO)
    parts.append(section("5. Seed data"))
    for t in TABLES:
        if t.seed_rows:
            parts.append(seed_sql(t))
    parts.append(GO)
    parts.append(section("6. English names of the system rows (only the empty ones)"))
    parts.extend(english_names_sql())
    parts.append(GO)
    parts.append(section("7. Relationships"))
    parts.extend(foreign_keys_sql())
    parts.append(GO)
    parts.append(section("8. Database options (plan 5C, docs/70)"))
    parts.append(DATABASE_OPTIONS)
    parts.append(GO)
    parts.append("PRINT N'RetailStore schema: done (" + counts + ").';\nGO\n")
    return "\n".join(parts)
