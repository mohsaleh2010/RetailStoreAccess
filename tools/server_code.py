"""Plan 5C (docs/70): the VBA that runs on linked SQL Server tables.

Access refuses (error 3622) a dynaset or an UPDATE / DELETE on a SQL Server table that has an IDENTITY column
unless the call says dbSeeChanges. The option does nothing on Access tables, so every statement and every
updatable recordset of the program carries it:

  db.Execute sql                          -> db.Execute sql, dbSeeChanges
  db.Execute sql, dbFailOnError           -> db.Execute sql, dbFailOnError + dbSeeChanges
  db.OpenRecordset(src)                   -> db.OpenRecordset(src, dbOpenDynaset, dbSeeChanges)
  db.OpenRecordset(src, dbOpenDynaset)    -> db.OpenRecordset(src, dbOpenDynaset, dbSeeChanges)
  db.OpenRecordset(src, t, dbAppendOnly)  -> db.OpenRecordset(src, t, dbAppendOnly + dbSeeChanges)
  (a snapshot or forward-only recordset is read-only: unchanged)

`see_changes` is applied once to the hand-written modules and by tools/generate.py to the generated modules that
run on the program's tables (demo data, in-Access tests). The build modules that open the back-end file itself
(BUILD_ON_FILE) are left as they are, and so are the pass-through querydefs of modSqlServer (the server itself). `missing_see_changes` lists the calls without the option (the test).
"""
import re
from typing import List, Tuple

BUILD_ON_FILE = {"modBuildSchema", "modBuildRelations", "modBuildQueries", "modBuildReports"}
_CALL = re.compile(r"\b(CurrentDb|db|m_db)\.(Execute|OpenRecordset)\b")
OPTION = "dbSeeChanges"


def _code_positions(text: str) -> List[bool]:
    """True for each character that is code (not inside a string or a comment)."""
    out, in_str, in_comment = [], False, False
    for i, ch in enumerate(text):
        if ch == "\n":
            in_comment = False
            in_str = False
            out.append(True)
            continue
        if in_comment:
            out.append(False)
            continue
        if ch == '"':
            in_str = not in_str
            out.append(False)
            continue
        if not in_str and ch == "'":
            in_comment = True
            out.append(False)
            continue
        out.append(not in_str)
    return out


def _arguments(text: str, start: int, closing: str) -> Tuple[List[Tuple[int, int]], int]:
    """Top-level argument spans from start up to the closing ")" (closing=")") or the end of the logical line
    (closing="\n"). Returns the spans and the end position."""
    spans, depth, in_str, i, arg_start = [], 0, False, start, start
    while i < len(text):
        ch = text[i]
        if in_str:
            if ch == '"':
                in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0 and closing == ")":
                spans.append((arg_start, i))
                return spans, i
            depth -= 1
        elif ch == "," and depth == 0:
            spans.append((arg_start, i))
            arg_start = i + 1
        elif ch == "'" and closing == "\n":
            break
        elif ch == "\n" and closing == "\n":
            line = text[:i].rstrip("\r")
            if not line.endswith(" _"):
                break
        i += 1
    end = i
    while end > arg_start and text[end - 1] in " \r\t":
        end -= 1
    spans.append((arg_start, end))
    return spans, end


def _plan(text: str):
    """(position, insertion) for every call without the option."""
    code = _code_positions(text)
    edits = []
    for m in _CALL.finditer(text):
        if not code[m.start()]:
            continue
        method = m.group(2)
        if method == "Execute":
            spans, end = _arguments(text, m.end(), "\n")
            args = [text[a:b].strip() for a, b in spans]
            if args == [""]:
                args = []
            if args and re.match(r'"(CREATE|ALTER|DROP) ', args[0]):
                continue                                    # the local temp tables: DDL takes no dbSeeChanges
            if any(OPTION in a for a in args):
                continue
            if len(args) == 1:
                edits.append((end, ", " + OPTION))
            else:
                edits.append((spans[-1][0] + len(text[spans[-1][0]:spans[-1][1]].rstrip()), " + " + OPTION))
        else:
            if m.end() >= len(text) or text[m.end()] != "(":
                continue
            spans, end = _arguments(text, m.end() + 1, ")")
            args = [text[a:b].strip() for a, b in spans]
            if any(OPTION in a for a in args):
                continue
            kind = args[1] if len(args) > 1 else ""
            if kind in ("dbOpenSnapshot", "dbOpenForwardOnly"):
                continue
            if len(args) == 1:
                edits.append((end, ", dbOpenDynaset, " + OPTION))
            elif len(args) == 2:
                edits.append((end, ", " + OPTION))
            else:
                last = spans[-1]
                edits.append((last[0] + len(text[last[0]:last[1]].rstrip()), " + " + OPTION))
    return edits


def ddl_with_see_changes(text: str) -> List[str]:
    """CREATE / ALTER / DROP statements given dbSeeChanges (they must not have it)."""
    out = []
    for m in _CALL.finditer(text):
        if m.group(2) != "Execute":
            continue
        spans, end = _arguments(text, m.end(), "\n")
        if re.match(r'\s*"(CREATE|ALTER|DROP) ', text[spans[0][0]:spans[0][1]]) and OPTION in text[m.end():end]:
            out.append(text[m.start():end])
    return out


def see_changes(text: str) -> str:
    for pos, ins in sorted(_plan(text), reverse=True):
        text = text[:pos] + ins + text[pos:]
    return text


def missing_see_changes(text: str) -> List[str]:
    """The lines of the calls that do not carry dbSeeChanges."""
    out = []
    for pos, _ in _plan(text):
        line_no = text.count("\n", 0, pos) + 1
        out.append(f"{line_no}: {text.split(chr(10))[line_no - 1].strip()[:100]}")
    return out


def auto_fields() -> set:
    from schema import TABLES
    return {f.name for t in TABLES for f in t.fields if f.kind == "AUTO"}


def ids_read_before_update(text: str, names: set) -> List[str]:
    """Lines that read an AutoNumber field between rs.AddNew and rs.Update: Access numbers the row at AddNew,
    SQL Server only when it is saved (read it after Update with rs.Bookmark = rs.LastModified)."""
    out, lines = [], text.split("\n")
    for i, line in enumerate(lines):
        m = re.search(r"\b(\w+)\.AddNew\b", line)
        if not m or line.lstrip().startswith("'"):
            continue
        rs = m.group(1)
        for j in range(i + 1, len(lines)):
            if re.search(rf"\b{rs}\.(Update|CancelUpdate)\b", lines[j]):
                break
            for f in re.findall(rf"[=(,&+]\s*{rs}!(\w+)", lines[j]):    # read, not written
                if f in names:
                    out.append(f"{j + 1}: {lines[j].strip()}")
    return out


# --------------------------------------------------------------------------
# Yes/No fields in SQL: SQL Server keeps 1 for "yes", Access -1. "= True" sent to the server as "= -1" finds
# nothing; "<> 0" is right on both (and "= 0" for "no"). Assignments (SET f = True) are converted by the server.
# --------------------------------------------------------------------------
def bool_fields() -> set:
    from schema import TABLES
    return {f.name for t in TABLES for f in t.fields if f.kind == "BOOL"}


def _bool_pattern(names):
    return re.compile(r"((?:\b\w+\.)?\[?\b(?:%s)\b\]?)\s*(=|<>)\s*(True|False|-1|Yes|No)\b(?!\s*,\s*\w+\s*=)"
                      % "|".join(sorted(names, key=len, reverse=True)))


_KEYWORD = re.compile(r"(?i)\b(SET|WHERE|AND|OR|ON|HAVING|NOT|WHEN|SELECT|IIf\s*\(|Filter\s*=)|\(")


def _is_assignment(before: str) -> bool:
    """True when the nearest SQL keyword before the field is SET (an UPDATE's SET list)."""
    found = list(_KEYWORD.finditer(before))
    return bool(found) and found[-1].group(0).upper() == "SET"


def _replacement(op: str, value: str) -> str:
    yes = value in ("True", "-1", "Yes")
    return ("<> 0" if yes else "= 0") if op == "=" else ("= 0" if yes else "<> 0")


def bool_comparisons(text: str, names=None) -> List[Tuple[int, int, str]]:
    """(start, end, replacement) of every comparison of a Yes/No field with True / False / -1 / Yes / No
    (assignments excepted). Only text inside string literals of one line is looked at."""
    pat = _bool_pattern(names or bool_fields())
    out, offset = [], 0
    for line in text.split("\n"):
        for lit in re.finditer(r'"((?:[^"\n]|"")*)"', line):
            body = lit.group(1)
            for m in pat.finditer(body):
                if _is_assignment(body[:m.start()]):
                    continue
                s = offset + lit.start(1) + m.start()
                out.append((s, s + len(m.group(0)), m.group(1) + " " + _replacement(m.group(2), m.group(3))))
        offset += len(line) + 1
    return out


def server_booleans(text: str, names=None) -> str:
    for s, e, rep in sorted(bool_comparisons(text, names), reverse=True):
        text = text[:s] + rep + text[e:]
    return text
