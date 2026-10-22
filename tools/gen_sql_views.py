"""Plan 5D (docs/71): the saved queries of the program as SQL Server views.

Every saved query of tools/queries.py becomes a view of the same name, so the server does the work (joins,
totals, UNIONs) and sends Access only the result. Two sets: [dbo] with the Arabic texts of an Arabic front-end
and [en] with the texts and the English master data names of an English one (what modLang.Tr makes of the
saved query in Access). The tables stay in [dbo].

Access SQL -> T-SQL:
  * CCur(x) -> CAST(x AS MONEY); CStr(x) -> CAST(x AS NVARCHAR(255)); Nz(x, d) -> COALESCE(x, d);
    IIf(c, a, b) -> IIF(c, a, b); Year / Month / Round -> YEAR / MONTH / ROUND;
    DateValue(x) -> the day of x; Date() -> today; DateAdd('d', n, x) / DateDiff('d', a, b) -> DATEADD / DATEDIFF;
  * True / False (a value) -> CAST(1 AS BIT) / CAST(0 AS BIT); 'text' -> N'text' (Arabic);
  * a / b -> a * 1e0 / b (Access divides in floating point, SQL Server two integers as integers);
  * a table -> [dbo].[T], a saved query -> [dbo].[Q] (or [en].[Q]);
  * ORDER BY is not allowed in a view: it stays in Access (see server_order);
  * the parameters QDate('X') / QLong('X') (modQueryParams) -> qp.[X] of the view [dbo].[QueryParamsNow]: one row
    with the parameters this Access process gave the server (table QueryParams, keyed by HOST_NAME() and
    HOST_ID(), the process id of the client). The SELECT that uses one gets "CROSS JOIN [dbo].[QueryParamsNow]
    AS qp" at the end of its FROM; inside an ON clause (where qp is not joined yet) a sub-query reads it.

In Access (modServerQueries): each saved query becomes "SELECT * FROM [sv_Q] ORDER BY ..." over the linked view
sv_Q; a query whose ORDER BY does not name its own columns stays an Access query (on the views below it).
"""

import hashlib
import re
from typing import Dict, List, Optional, Tuple

import i18n
import i18n_en
import queries as Q
from schema import TABLES

GO = "GO"
SCHEMAS = ("dbo", "en")
PARAMS_TABLE = "QueryParams"
PARAMS_VIEW = "QueryParamsNow"
PARAM_PREFIX = "qp_"                      # its columns: never the name of a column of the query (an unqualified name)
DATE_ZERO = "'18991230'"                  # QDate of a parameter not set (Access day 0)
INTERVALS = {"d": "day", "m": "month", "yyyy": "year", "ww": "week", "q": "quarter", "h": "hour", "n": "minute"}

# T-SQL reserved words that Access SQL does not use as keywords: as a column or alias they need brackets
TSQL_RESERVED = set("""ADD ALTER ANY AUTHORIZATION BACKUP BEGIN BREAK BROWSE BULK CASCADE CASE CHECK CHECKPOINT CLOSE
    CLUSTERED COALESCE COLLATE COLUMN COMMIT COMPUTE CONSTRAINT CONTAINS CONTAINSTABLE CONTINUE CONVERT CREATE CROSS
    CURRENT CURRENT_DATE CURRENT_TIME CURRENT_TIMESTAMP CURRENT_USER CURSOR DATABASE DBCC DEALLOCATE DECLARE DEFAULT
    DELETE DENY DISK DISTRIBUTED DOUBLE DROP DUMP ELSE END ERRLVL ESCAPE EXCEPT EXEC EXECUTE EXIT EXTERNAL FETCH
    FILE FILLFACTOR FOR FOREIGN FREETEXT FREETEXTTABLE FULL FUNCTION GOTO GRANT HOLDLOCK IDENTITY IDENTITY_INSERT
    IDENTITYCOL IF INDEX INSERT INTERSECT INTO KEY KILL LINENO LOAD MERGE NATIONAL NOCHECK NONCLUSTERED NULLIF OF
    OFF OFFSETS OPEN OPENDATASOURCE OPENQUERY OPENROWSET OPENXML OPTION OUTER OVER PERCENT PIVOT PLAN PRECISION
    PRIMARY PRINT PROC PROCEDURE PUBLIC RAISERROR READ READTEXT RECONFIGURE REFERENCES REPLICATION RESTORE RESTRICT
    RETURN REVERT REVOKE ROLLBACK ROWCOUNT ROWGUIDCOL RULE SAVE SCHEMA SECURITYAUDIT SESSION_USER SET SETUSER
    SHUTDOWN SOME STATISTICS SYSTEM_USER TABLE TABLESAMPLE TEXTSIZE THEN TO TRAN TRANSACTION TRIGGER TRUNCATE
    TRY_CONVERT TSEQUAL UNIQUE UNPIVOT UPDATE UPDATETEXT USE USER VALUES VARYING VIEW WAITFOR WHEN WHILE WITH
    WITHIN WRITETEXT""".split())
TABLE_NAMES = {t.name for t in TABLES}
QUERY_NAMES = [q.name for q in Q.QUERIES]


# --------------------------------------------------------------------------
# Tokens
# --------------------------------------------------------------------------
_TOKEN = re.compile(r"""
    (?P<str>'(?:[^']|'')*')
  | (?P<br>\[[^\]]*\])
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<word>[A-Za-z_@][A-Za-z0-9_]*)
  | (?P<ws>\s+)
  | (?P<op><>|<=|>=|[-+*/=<>(),.;&\\])
""", re.X)


PARAM_MARK = "QP__"                       # QDate('X') / QLong('X') once translated: placed by _place_params


def tokenize(sql: str) -> List[str]:
    out, pos = [], 0
    while pos < len(sql):
        m = _TOKEN.match(sql, pos)
        if not m:
            raise ValueError(f"cannot read SQL at: {sql[pos:pos + 40]!r}")
        out.append(m.group(0))
        pos = m.end()
    return out


def _is_ws(t) -> bool:
    return isinstance(t, str) and t.isspace()


def _word(t) -> str:
    return t.upper() if isinstance(t, str) and re.match(r"[A-Za-z_]", t) else ""


def _split_args(toks: List[str]) -> List[List[str]]:
    """The arguments of a call (toks: between the parentheses) at depth 0."""
    args, cur, depth = [], [], 0
    for t in toks:
        if t == "(":
            depth += 1
        elif t == ")":
            depth -= 1
        if t == "," and depth == 0:
            args.append(cur)
            cur = []
        else:
            cur.append(t)
    args.append(cur)
    return [_strip(a) for a in args]


def _strip(toks):
    while toks and _is_ws(toks[0]):
        toks = toks[1:]
    while toks and _is_ws(toks[-1]):
        toks = toks[:-1]
    return toks


def _close(toks: List[str], i: int) -> int:
    """Index of the ")" matching the "(" at i."""
    depth = 0
    for j in range(i, len(toks)):
        if toks[j] == "(":
            depth += 1
        elif toks[j] == ")":
            depth -= 1
            if depth == 0:
                return j
    raise ValueError("unbalanced parentheses")


def _text(toks) -> str:
    return "".join(t if isinstance(t, str) else repr(t) for t in toks)


def _literal(toks) -> str:
    t = _strip(toks)
    assert len(t) == 1 and re.match(r"N?'", t[0]), _text(toks)
    return t[0].lstrip("N")[1:-1]


# --------------------------------------------------------------------------
# Translation
# --------------------------------------------------------------------------
def _day_of(x: List) -> List:
    """The day of a date (time 00:00), as a DATETIME: DATEADD(day, DATEDIFF(day, 0, x), 0)."""
    return ["DATEADD(day, DATEDIFF(day, 0, "] + x + ["), 0)"]


def _call(name: str, args: List[List]) -> List:
    n = name.lower()
    if n == "ccur":
        return ["CAST("] + args[0] + [" AS MONEY)"]
    if n == "cdbl":
        return ["CAST("] + args[0] + [" AS FLOAT)"]
    if n == "clng":
        return ["CAST(ROUND("] + args[0] + [", 0) AS INT)"]
    if n == "cstr":
        return ["CAST("] + args[0] + [" AS NVARCHAR(255))"]
    if n == "nz":
        return ["COALESCE("] + args[0] + [", "] + (args[1] if len(args) > 1 else ["0"]) + [")"]
    if n == "iif":
        return ["IIF("] + args[0] + [", "] + args[1] + [", "] + args[2] + [")"]
    if n in ("year", "month", "day", "abs"):
        return [n.upper() + "("] + args[0] + [")"]
    if n == "round":
        return ["ROUND("] + args[0] + [", "] + (args[1] if len(args) > 1 else ["0"]) + [")"]
    if n == "datevalue":
        return _day_of(args[0])
    if n == "date":
        return _day_of(["GETDATE()"])
    if n in ("dateadd", "datediff"):
        unit = INTERVALS[_literal(args[0]).lower()]
        return [n.upper() + "(" + unit + ", "] + args[1] + [", "] + args[2] + [")"]
    if n in ("qdate", "qlong"):
        return [PARAM_MARK + _literal(args[0])]
    if n in ("sum", "count", "max", "min", "avg"):
        return [name.capitalize() + "("] + (args[0] if args != [[]] else []) + [")"]
    raise ValueError(f"no T-SQL for the Access function {name}()")


SQL_BEFORE_PARENTHESIS = {"IN", "EXISTS", "ON", "AND", "OR", "NOT", "FROM", "JOIN", "WHERE", "SELECT", "AS", "ALL",
                          "HAVING", "BY", "WHEN", "THEN", "ELSE", "UNION"}
FUNCTIONS = {"ccur", "cdbl", "clng", "cstr", "nz", "iif", "year", "month", "day", "abs", "round", "datevalue",
             "date", "dateadd", "datediff", "qdate", "qlong", "sum", "count", "max", "min", "avg"}


def _source_position(prev, toks, i) -> bool:
    """A name that can be a table or query: not a column (x.Name), not an alias (AS Name), not qualified."""
    return prev != "." and _word(prev) != "AS" and (i + 1 >= len(toks) or toks[i + 1] != ".")


def _translate(toks: List, schema: str) -> List:
    out, i = [], 0
    while i < len(toks):
        t = toks[i]
        nxt = i + 1
        while nxt < len(toks) and _is_ws(toks[nxt]):
            nxt += 1
        if isinstance(t, str) and t.lower() in FUNCTIONS and nxt < len(toks) and toks[nxt] == "(" and \
                (not out or out[-1] != "."):
            end = _close(toks, nxt)
            args = [_translate(a, schema) for a in _split_args(toks[nxt + 1:end])]
            out += _call(t, args)
            i = end + 1
            continue
        if isinstance(t, str) and re.match(r"[A-Za-z_]\w*$", t) and nxt < len(toks) and toks[nxt] == "(" and \
                t.upper() not in SQL_BEFORE_PARENTHESIS and (not out or out[-1] != "."):
            raise ValueError(f"no T-SQL for the Access function {t}()")
        prev = next((x for x in reversed(out) if not _is_ws(x)), "")
        name = t[1:-1] if isinstance(t, str) and t.startswith("[") else t
        if isinstance(t, str) and t.startswith("'"):
            out.append("N" + t)
        elif t in ("True", "False") and prev != ".":
            out.append("CAST(1 AS BIT)" if t == "True" else "CAST(0 AS BIT)")
        elif t == "/":
            out.append("* 1e0 /")
        elif isinstance(t, str) and t.upper() in TSQL_RESERVED:
            out.append(f"[{t}]")
        elif _source_position(prev, toks, i) and name in TABLE_NAMES:
            out.append(f"[dbo].[{name}]")
        elif _source_position(prev, toks, i) and name in QUERY_SET:
            out.append(f"[{schema}].[{name}]")
        else:
            out.append(t)
        i += 1
    return out


QUERY_SET = set(QUERY_NAMES)


# --------------------------------------------------------------------------
# Yes/No fields used as a condition (IIf(d.ReturnToStock, ...), WHERE p.IsActive): SQL Server needs "<> 0"
# --------------------------------------------------------------------------
BOOL_FIELDS = {f.name for t in TABLES for f in t.fields if f.kind == "BOOL"}
_COND_BEFORE = {"AND", "OR", "NOT", "WHERE", "HAVING", "ON", "WHEN"}
_COMPARE = {"=", "<>", "<", ">", "<=", ">=", "IS", "IN", "LIKE", "BETWEEN", "+", "-", "*", "/"}


def _bool_conditions(toks: List[str]) -> List[str]:
    def prev_of(k):
        k -= 1
        while k >= 0 and _is_ws(toks[k]):
            k -= 1
        return k

    def next_of(k):
        k += 1
        while k < len(toks) and _is_ws(toks[k]):
            k += 1
        return k

    after = set()
    for k, t in enumerate(toks):
        if t.strip("[]") not in BOOL_FIELDS:
            continue
        start = k - 2 if k >= 2 and toks[k - 1] == "." else k
        if k + 1 < len(toks) and toks[k + 1] == ".":
            continue
        p = prev_of(start)
        before = toks[p] if p >= 0 else ""
        cond = _word(before) in _COND_BEFORE
        if before == "(":
            pp = prev_of(p)
            cond = pp < 0 or _word(toks[pp]) in ("IIF", "") or _word(toks[pp]) in _COND_BEFORE
        n = next_of(k)
        nxt = toks[n].upper() if n < len(toks) else ""
        if cond and nxt not in _COMPARE:
            after.add(k)
    return [x for k, t in enumerate(toks) for x in ([t, " <> 0"] if k in after else [t])]


# --------------------------------------------------------------------------
# SELECT scopes: where the parameter row is joined
# --------------------------------------------------------------------------
_CLAUSE_END = {"WHERE", "GROUP", "HAVING", "ORDER", "UNION"}


def _selects(toks: List) -> List[Tuple[int, int, int, int]]:
    """(select index, from index or -1, end of the FROM clause, end of the SELECT) for every SELECT."""
    out = []
    for i, t in enumerate(toks):
        if _word(t) != "SELECT":
            continue
        depth, frm, from_end, j = 0, -1, None, i + 1
        while j < len(toks):
            u = toks[j]
            if u == "(":
                depth += 1
            elif u == ")":
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0:
                w = _word(u)
                if w == "FROM" and frm < 0:
                    frm = j
                elif w in _CLAUSE_END:
                    if frm >= 0 and from_end is None:
                        from_end = j
                    if w == "UNION":
                        break
            j += 1
        if from_end is None:
            from_end = j
        out.append((i, frm, from_end, j))
    return out


def _sources(toks: List[str], start: int, end: int) -> List[str]:
    """The tables and views of a FROM clause, by alias when they have one."""
    out, k = [], start
    while k < end:
        if toks[k] == "[dbo]" or toks[k] == "[en]":
            name = "".join(toks[k:k + 3])
            j = k + 3
            while j < end and _is_ws(toks[j]):
                j += 1
            if j < end and _word(toks[j]) == "AS":
                j += 1
                while _is_ws(toks[j]):
                    j += 1
                name = toks[j]
            out.append(name)
            k = j
        k += 1
    return out


def _in_aggregate(toks: List[str], k: int) -> bool:
    """True when the token k is inside Sum(...) / Count(...) / Max(...) / Min(...) / Avg(...)."""
    depth = 0
    for j in range(k - 1, -1, -1):
        if toks[j] == ")":
            depth += 1
        elif toks[j] == "(":
            if depth == 0:
                w = j - 1
                while w >= 0 and _is_ws(toks[w]):
                    w -= 1
                if w >= 0 and _word(toks[w]) in ("SUM", "COUNT", "MAX", "MIN", "AVG"):
                    return True
            else:
                depth -= 1
    return False


def _clause_of(toks: List[str], sel, k: int) -> str:
    """WHERE, GROUP or HAVING: the clause of its SELECT the token k is in (after the FROM clause)."""
    depth, clause = 0, ""
    for j in range(sel[2], k):
        if toks[j] == "(":
            depth += 1
        elif toks[j] == ")":
            depth -= 1
        elif depth == 0 and _word(toks[j]) in ("WHERE", "GROUP", "HAVING"):
            clause = _word(toks[j])
    return clause


def _place_params(toks: List) -> List:
    """A parameter in WHERE / GROUP BY / an aggregate reads the joined row qp; in the column list (a grouped
    query cannot show a joined column that is not grouped), HAVING or an ON clause a sub-query reads it."""
    sels = _selects(toks)
    joins = set()
    repl = {}
    for k, t in enumerate(toks):
        if not t.startswith(PARAM_MARK):
            continue
        name = t[len(PARAM_MARK):]
        inner = max((s for s in sels if s[0] < k < s[3]), key=lambda s: s[0])
        in_from = inner[1] >= 0 and inner[1] < k < inner[2]
        in_list = k < inner[1] or _clause_of(toks, inner, k) == "HAVING"
        if in_from or (in_list and not _in_aggregate(toks, k)):
            repl[k] = f"(SELECT p.[{PARAM_PREFIX}{name}] FROM [dbo].[{PARAMS_VIEW}] AS p)"
        else:
            repl[k] = f"qp.[{PARAM_PREFIX}{name}]"
            joins.add(inner)
    for s in joins:                                  # "SELECT *" would also give the parameter columns
        stars = [k for k in range(s[0] + 1, s[1]) if not _is_ws(toks[k])]
        if len(stars) == 1 and toks[stars[0]] == "*":
            repl[stars[0]] = ", ".join(f"{src}.*" for src in _sources(toks, s[1], s[2]))
    out = []
    ends = {}
    for s in joins:
        end = s[2]
        while end > 0 and _is_ws(toks[end - 1]):
            end -= 1
        ends[end] = True
    for k, t in enumerate(toks):
        if k in ends:
            out.append(f" CROSS JOIN [dbo].[{PARAMS_VIEW}] AS qp")
        out.append(repl.get(k, t))
    if len(toks) in ends:
        out.append(f" CROSS JOIN [dbo].[{PARAMS_VIEW}] AS qp")
    return out


# --------------------------------------------------------------------------
# ORDER BY
# --------------------------------------------------------------------------
def split_order(toks: List[str]) -> Tuple[List[str], List[List[str]]]:
    """The query without its final ORDER BY, and the ORDER BY terms."""
    depth, at = 0, None
    for i, t in enumerate(toks):
        if t == "(":
            depth += 1
        elif t == ")":
            depth -= 1
        elif depth == 0 and _word(t) == "ORDER":
            at = i
    if at is None:
        return toks, []
    j = at + 1
    while _word(toks[j]) != "BY":
        j += 1
    return _strip(toks[:at]), _split_args(toks[j + 1:])


def _norm(toks) -> str:
    return re.sub(r"\s+", "", _text(toks)).replace("[", "").replace("]", "").lower()


def select_items(toks: List[str]) -> List[Tuple[str, str]]:
    """(normalised expression, output name) of the columns of the first SELECT."""
    i = next(k for k, t in enumerate(toks) if _word(t) == "SELECT")
    depth, j = 0, i + 1
    while j < len(toks):
        if toks[j] == "(":
            depth += 1
        elif toks[j] == ")":
            depth -= 1
        elif depth == 0 and _word(toks[j]) == "FROM":
            break
        j += 1
    items = []
    for item in _split_args(toks[i + 1:j]):
        words = [k for k, t in enumerate(item) if _word(t) == "AS"]
        depth, last_as = 0, None
        for k, t in enumerate(item):
            if t == "(":
                depth += 1
            elif t == ")":
                depth -= 1
            elif depth == 0 and k in words:
                last_as = k
        if last_as is not None:
            expr, alias = _strip(item[:last_as]), _strip(item[last_as + 1:])[0].strip("[]")
        else:
            expr = item
            alias = _strip(item)[-1].strip("[]") if _strip(item)[-1] != "*" else "*"
        items.append((_norm(expr), alias))
    return items


def server_order(sql: str) -> Optional[str]:
    """The ORDER BY of the Access query over its linked view ("" = none), or None when a term is not one of
    its own columns (the query then stays in Access)."""
    toks = tokenize(sql)
    body, terms = split_order(toks)
    if not terms:
        return ""
    items = select_items(body)
    star = any(a == "*" or e.endswith(".*") for e, a in items)
    out = []
    for term in terms:
        desc = bool(term) and _word(term[-1]) in ("DESC", "ASC")
        direction = " DESC" if desc and _word(term[-1]) == "DESC" else ""
        expr = _strip(term[:-1]) if desc else term
        key = _norm(expr)
        col = None
        for e, a in items:
            if a != "*" and (key == a.lower() or key == e):
                col = a
                break
        if col is None and re.fullmatch(r"(\w+\.)?\w+", key):
            bare = key.split(".")[-1]
            match = [a for e, a in items if a.lower() == bare and (e == key or e == bare)]
            if match:
                col = match[0]
            elif star:
                col = _strip(expr)[-1].strip("[]")
        if col is None:
            return None
        out.append(f"[{col}]{direction}")
    return " ORDER BY " + ", ".join(out)


# --------------------------------------------------------------------------
# The views
# --------------------------------------------------------------------------
def access_sql(q: Q.Query, english: bool) -> str:
    """The saved query as Access keeps it in an Arabic or English front-end (SaveQuery: Tr)."""
    if english:
        return i18n.translate(q.sql, i18n_en.EN, ENGLISH_ORDER)
    return i18n.resolve_names(q.sql, False)


ENGLISH_ORDER = i18n.ordered(i18n_en.EN)


def view_body(q: Q.Query, schema: str = "dbo") -> str:
    toks = tokenize(access_sql(q, schema == "en"))
    body, _ = split_order(toks)
    return "".join(_place_params(_bool_conditions(tokenize(_text(_translate(body, schema)))))).strip()


def query_params() -> Dict[str, str]:
    """Every parameter of the saved queries: name -> D (date, QDate) or L (number, QLong), in order of use."""
    out = {}
    for q in Q.QUERIES:
        for kind, name in re.findall(r"Q(Date|Long)\('(\w+)'\)", q.sql):
            k = "D" if kind == "Date" else "L"
            assert out.get(name, k) == k, name
            out[name] = k
    return out


def params_sql() -> List[str]:
    p = query_params()
    cols = [f"IF COL_LENGTH(N'dbo.{PARAMS_TABLE}', N'{n}') IS NULL "
            f"ALTER TABLE [dbo].[{PARAMS_TABLE}] ADD [{n}] {'DATETIME' if k == 'D' else 'INT'} NULL;"
            for n, k in p.items()]
    view_cols = ",\n    ".join(f"COALESCE(p.[{n}], 0) AS [{PARAM_PREFIX}{n}]" if k == "L" else
                               f"COALESCE(p.[{n}], CAST({DATE_ZERO} AS DATETIME)) AS [{PARAM_PREFIX}{n}]"
                               for n, k in p.items())
    return [
        "\n".join([
            f"IF OBJECT_ID(N'dbo.{PARAMS_TABLE}', N'U') IS NULL",
            f"CREATE TABLE [dbo].[{PARAMS_TABLE}] (",
            "    [HostName] NVARCHAR(128) NOT NULL,",
            "    [HostProcess] NVARCHAR(10) NOT NULL,",
            "    [UpdatedAt] DATETIME NOT NULL CONSTRAINT [DF_QueryParams_UpdatedAt] DEFAULT GETDATE(),",
            f"    CONSTRAINT [PK_{PARAMS_TABLE}] PRIMARY KEY ([HostName], [HostProcess]));"] + cols),
        "IF SCHEMA_ID(N'en') IS NULL EXEC (N'CREATE SCHEMA [en]');",
        f"CREATE OR ALTER VIEW [dbo].[{PARAMS_VIEW}] AS\n"
        f"-- the parameters of this Access process (modQueryParams), else those of a parameter not set\n"
        f"SELECT {view_cols}\n"
        f"FROM (SELECT 1 AS One) AS o LEFT JOIN [dbo].[{PARAMS_TABLE}] AS p\n"
        f"    ON p.HostName = HOST_NAME() AND p.HostProcess = HOST_ID();",
    ]


def view_sql(q: Q.Query, schema: str) -> str:
    return f"CREATE OR ALTER VIEW [{schema}].[{q.name}] AS\n-- {q.caption}\n{view_body(q, schema)};"


def build_views_sql() -> str:
    body = _views_sql()
    return body + "\n".join([f"CREATE OR ALTER VIEW [dbo].[{VERSION_VIEW}] AS\n"
                             f"-- the version of these views: the front-end links to them only when it is its own\n"
                             f"SELECT N'{views_version()}' AS Version;", GO]) + "\n"


VERSION_VIEW = "QueriesVersion"


def views_version() -> str:
    """12 characters that change with any view (modServerQueryData.SERVER_QUERIES_VERSION)."""
    return hashlib.sha1(_views_sql().encode("utf-8")).hexdigest()[:12]


def _views_sql() -> str:
    parts = ["-- RetailStore: the saved queries of the program as views (docs/71-SQL-Server-Views.md, plan 5D).",
             "-- Generated by tools/gen_sql_views.py from tools/queries.py - do not edit.",
             "-- Run by the program after the database script (SQL Server screen, Prepare the database).",
             "-- [dbo] = an Arabic front-end, [en] = an English one; SQL Server 2016 SP1 or later (CREATE OR ALTER).",
             "SET NOCOUNT ON;", GO]
    for s in params_sql():
        parts += [s, GO]
    for schema in SCHEMAS:
        for q in Q.QUERIES:
            parts += [view_sql(q, schema), GO]
    return "\n".join(parts) + "\n"


def server_query_list() -> List[Tuple[str, Optional[str]]]:
    """(query, ORDER BY in Access over its view; None = stays an Access query), in creation order."""
    return [(q.name, server_order(q.sql)) for q in Q.QUERIES]


def build_data_vba() -> str:
    """modServerQueryData: what modServerQueries needs to know of the views."""
    lines = []
    chunk = ""
    for name, order in server_query_list():
        item = f"{name}|{'~' if order is None else order};"
        if len(chunk) + len(item) > 400:
            lines.append(chunk)
            chunk = ""
        chunk += item
    lines.append(chunk)
    body = "\n".join(f'    s = s & "{c}"' for c in lines)
    params = ",".join(f"{n}:{k}" for n, k in query_params().items())
    return f"""Attribute VB_Name = "modServerQueryData"
'==============================================================================
' modServerQueryData  -  GENERATED FILE - do not edit by hand.
' Source: tools/queries.py -> tools/gen_sql_views.py (plan 5D, docs/71); used by modServerQueries.
'==============================================================================
Option Compare Database
Option Explicit

' the version of dist\\sql\\RetailStore_Queries.sql this front-end needs (view dbo.{VERSION_VIEW})
Public Const SERVER_QUERIES_VERSION As String = "{views_version()}"

Public Function ServerQueryList() As String
    ' Name|ORDER BY of the query over its view ("~" = it stays an Access query), separated by ";"
    Dim s As String
{body}
    ServerQueryList = Left$(s, Len(s) - 1)
End Function

Public Function ServerParamList() As String
    ' the parameters the views read (QDate = D, QLong = L): the columns of dbo.{PARAMS_TABLE}
    ServerParamList = "{params}"
End Function
"""


# --------------------------------------------------------------------------
# Python mirror of the pure parts of modServerQueries (checked in LibreOffice)
# --------------------------------------------------------------------------
def server_query_sql(name: str, order: str) -> str:
    return f"SELECT * FROM [sv_{name}]{order};"


def query_params_sql(names, values) -> str:
    import sql_transfer_reference as R
    kinds = query_params()
    sets, cols, vals = "", "", ""
    for n, v in zip(names, values):
        kind = kinds.get(n)
        if not kind:
            continue
        lit = "NULL" if v is None else R.sql_literal(v, 8 if kind == "D" else 4)
        sets += f"[{n}] = {lit}, "
        cols += f", [{n}]"
        vals += f", {lit}"
    if not sets:
        return ""
    return ("MERGE [dbo].[QueryParams] AS t\n"
            "USING (SELECT HOST_NAME() AS HostName, HOST_ID() AS HostProcess) AS s\n"
            "ON t.HostName = s.HostName AND t.HostProcess = s.HostProcess\n"
            f"WHEN MATCHED THEN UPDATE SET {sets}UpdatedAt = GETDATE()\n"
            f"WHEN NOT MATCHED THEN INSERT (HostName, HostProcess{cols}) VALUES (s.HostName, s.HostProcess{vals});\n"
            "DELETE FROM [dbo].[QueryParams] WHERE UpdatedAt < DATEADD(day, -30, GETDATE());")
