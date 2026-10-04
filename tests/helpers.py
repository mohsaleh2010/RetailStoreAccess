"""Shared helpers for the test-suite: an SQLite mirror of the Access back-end
and static checks for generated VBA modules."""

import os
import re
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))

from schema import TABLES  # noqa: E402

SQLITE_TYPE = {
    "AUTO": "INTEGER", "LONG": "INTEGER", "INT": "INTEGER", "BYTE": "INTEGER",
    "MONEY": "NUMERIC", "QTY": "NUMERIC", "RATE": "NUMERIC", "DATE": "TEXT",
    "DATETIME": "TEXT", "BOOL": "INTEGER", "TEXT": "TEXT", "MEMO": "TEXT",
}


def sqlite_default(access_default):
    """Translate an Access DefaultValue expression to SQLite."""
    d = access_default
    if d in ("Now()", "Date()"):
        return "CURRENT_TIMESTAMP"
    if d in ("True", "False"):
        return "1" if d == "True" else "0"
    if d.startswith('"'):
        return "'" + d.strip('"') + "'"
    float(d)  # anything else must be a plain number
    return d


def build_sqlite(with_relationship_rules=False):
    """Create an in-memory mirror of the back-end and load the seed data.

    with_relationship_rules=True mirrors Phase 3: ON DELETE CASCADE where the
    schema says so, RESTRICT everywhere else, ON UPDATE CASCADE for text keys.
    """
    from relations import relations
    rel_by_child = {(r.child, r.child_field): r for r in relations()}

    con = sqlite3.connect(":memory:", isolation_level=None)
    con.execute("PRAGMA foreign_keys = ON")
    for t in TABLES:
        cols = []
        for f in t.fields:
            c = f'"{f.name}" {SQLITE_TYPE[f.kind]}'
            if f.required and f.kind not in ("AUTO", "BOOL"):
                c += " NOT NULL"
            if f.default is not None:
                c += " DEFAULT " + sqlite_default(f.default)
            cols.append(c)
        cols.append("PRIMARY KEY (" + ", ".join(f'"{k}"' for k in t.pk) + ")")
        for f in t.fields:
            if f.fk:
                tt, tf = f.fk.split(".")
                c = f'FOREIGN KEY ("{f.name}") REFERENCES "{tt}" ("{tf}")'
                if with_relationship_rules:
                    r = rel_by_child[(t.name, f.name)]
                    c += " ON DELETE " + ("CASCADE" if r.cascade_delete else "RESTRICT")
                    c += " ON UPDATE " + ("CASCADE" if r.cascade_update else "RESTRICT")
                cols.append(c)
        for ix in t.indexes:
            if ix.unique:
                cols.append("UNIQUE (" + ", ".join(f'"{k}"' for k in ix.fields) + ")")
        con.execute(f'CREATE TABLE "{t.name}" ({", ".join(cols)})')

    # Seed in one transaction with deferred FK checks: Settings points to the
    # cash customer, which is seeded later (Access builds relationships after seeding).
    con.execute("BEGIN")
    con.execute("PRAGMA defer_foreign_keys = ON")
    for t in TABLES:
        for row in t.seed_rows:
            ph = ", ".join("?" for _ in row)
            cols = ", ".join(f'"{c}"' for c in t.seed_columns)
            con.execute(f'INSERT INTO "{t.name}" ({cols}) VALUES ({ph})', row)
    con.execute("COMMIT")
    return con


# --------------------------------------------------------------------------
# Static checks for generated VBA
# --------------------------------------------------------------------------
def logical_lines(vba):
    """Code lines with comments removed and ' _' continuations joined."""
    out, buf = [], ""
    for raw in vba.splitlines():
        line = strip_comment(raw).rstrip()
        if line.endswith(" _"):
            buf += line[:-2] + " "
            continue
        out.append((buf + line).strip())
        buf = ""
    return out


def split_statements(line):
    """Split a logical line on ': ' statement separators outside string literals."""
    out, buf, in_str = [], "", False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
        if not in_str and line.startswith(": ", i):
            out.append(buf)
            buf = ""
            i += 2
            continue
        buf += ch
        i += 1
    out.append(buf)
    return out


def strip_comment(line):
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    return line


class VbaModuleChecks:
    """Mixin: set `module_name` and `vba` (generated text) on the TestCase class."""

    module_name = ""
    vba = ""

    def test_generated_files_are_up_to_date(self):
        src = os.path.join(ROOT, "src", "vba", self.module_name + ".bas")
        dist = os.path.join(ROOT, "dist", "vba", self.module_name + ".bas")
        with open(src, encoding="utf-8") as fh:
            self.assertEqual(fh.read().rstrip("\n"), self.vba.rstrip("\n"),
                             "run: python3 tools/generate.py")
        with open(dist, "rb") as fh:
            data = fh.read()
        self.assertEqual(data.decode("cp1256").replace("\r\n", "\n").rstrip("\n"),
                         self.vba.rstrip("\n"))
        self.assertNotIn(b"\n", data.replace(b"\r\n", b""), "dist file must use CRLF only")

    def test_module_header(self):
        lines = self.vba.splitlines()
        self.assertEqual(lines[0], f'Attribute VB_Name = "{self.module_name}"')
        self.assertIn("Option Explicit", lines)

    def test_windows_1256_encodable(self):
        self.vba.encode("cp1256")

    def test_line_length_and_continuations(self):
        run = 0
        for n, line in enumerate(self.vba.splitlines(), 1):
            self.assertLess(len(line.encode("cp1256")), 1000, f"line {n} too long")
            run = run + 1 if line.endswith(" _") else 0
            self.assertLess(run, 24, f"too many continuations at line {n}")

    def test_blocks_balanced(self):
        code = logical_lines(self.vba)

        def count(pattern):
            return sum(1 for l in code if re.match(pattern, l))

        self.assertEqual(count(r"^(Public |Private )?(Sub|Function) "),
                         count(r"^End (Sub|Function)$"))
        self.assertEqual(count(r"^For "), count(r"^Next\b"))
        self.assertEqual(count(r"^Select Case "), count(r"^End Select$"))
        self.assertEqual(count(r"^If .* Then$"), count(r"^End If$"))

    def test_string_literals_balanced(self):
        for n, line in enumerate(self.vba.splitlines(), 1):
            if line.lstrip().startswith("'"):
                continue
            in_str = False
            for ch in line.replace('""', ""):
                if ch == '"':
                    in_str = not in_str
                elif ch == "'" and not in_str:
                    break
            self.assertFalse(in_str, f"unbalanced quotes at line {n}: {line}")

    def test_variables_declared(self):
        """Option Explicit: every assigned local identifier must be declared somewhere."""
        declared = {"calendar"}   # built-in VBA properties that may be assigned
        for l in logical_lines(self.vba):
            for m in re.finditer(r"\b(?:Dim|Private|Public|Const|ByVal|ByRef|Optional)\s+"
                                 r"(?:Const\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)", l):
                declared.add(m.group(1).lower())
            for m in re.finditer(r",\s*(?:ByVal\s+|ByRef\s+|Optional\s+ByVal\s+)?(\w+)(?:\(\))?\s+As\b", l):
                declared.add(m.group(1).lower())
            m = re.match(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", l)
            if m:
                declared.add(m.group(1).lower())
            m = re.match(r"^For (?:Each )?(\w+)", l)
            if m:
                self.assertIn(m.group(1).lower(), declared, f"loop variable {m.group(1)}")
        for l in logical_lines(self.vba):
            for stmt in split_statements(l):
                m = re.match(r"^(?:Set\s+)?(\w+)\s*=(?!=)", stmt.strip())
                if m and not stmt.strip().startswith(("If ", "ElseIf ", "Case ")):
                    self.assertIn(m.group(1).lower(), declared, f"undeclared: {stmt.strip()}")
