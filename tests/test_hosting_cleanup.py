"""Removing the demo data from a database on shared hosting (docs/88, section 3).

A shared host refuses BACKUP DATABASE, so RemoveDemoData stopped before deleting anything. On a
front-end linked to SQL Server it now offers ServerDataCopy instead: every table of the schema copied
into an Access file on this PC, the rows of each table counted in both, before the delete transaction.
A second run after DEMO_REMOVED deletes nothing. Also: no VBA variable is named after a keyword
(`step` of For ... Step was the first name chosen for ServerDataCopy and does not compile).
"""

import os
import re
import unittest

from test_server_code import modules, proc, read


class DemoRemovalOnHostingTests(unittest.TestCase):

    def test_remove_falls_back_to_a_data_copy_on_the_server(self):
        body = proc(read("modDemoData"), "RemoveDemoData")
        backup = body.index('BackupNow(file, , "_BeforeDemoRemoval")')
        fallback = body.index('ServerDataCopy(file, "_BeforeDemoRemoval")')
        self.assertLess(backup, fallback)
        self.assertLess(fallback, body.index("ws.BeginTrans"))
        self.assertIn("If Len(msg) > 0 And IsLinkedToServer() Then", body[backup:fallback])
        self.assertIn("<> vbYes Then Exit Function", body[backup:fallback], "the owner agrees to the copy")
        self.assertLess(fallback, body.index("If Len(msg) > 0 Then", fallback), "a failed copy still deletes nothing")

    def test_a_second_removal_deletes_nothing(self):
        body = proc(read("modDemoData"), "RemoveDemoData")
        check = body.index("ActionType = 'DEMO_REMOVED' AND LogDate > ")
        self.assertLess(body.index("loaded = DMax("), check)
        self.assertLess(check, body.index("ws.BeginTrans"))
        self.assertIn("Exit Function", body[check:body.index("For Each t", check)])

    def test_the_data_copy_takes_every_table_and_counts_the_rows(self):
        body = proc(read("modBackup"), "ServerDataCopy")
        self.assertIn('HasPermission("BACKUP")', body)
        self.assertIn('CurrentProject.Path & "\\Backups"', body, "a folder of this PC, not the server's")
        self.assertIn("DBEngine.CreateDatabase(CopyFile, dbLangGeneral)", body)
        self.assertEqual(body.count("Split(SchemaTableList(), \",\")"), 1)
        copy = body.index('"SELECT * INTO [" & t & "] IN \'" & quoted & "\' FROM [" & t & "]"')
        count = body.index("If n <> m Then")
        self.assertLess(copy, count)
        self.assertLess(count, body.index('LogAction "BACKUP"'))
        self.assertIn('Replace(CopyFile, "\'", "\'\'")', body)


VBA_KEYWORDS = set("""
step to next loop wend until while each in is mod not and or xor eqv imp then else elseif end sub function
property dim redim preserve as new set let get call exit goto gosub return select case with option private
public static const type enum declare lib alias byval byref optional paramarray nothing empty null true false
me resume error stop debug print open close input output append binary random lock unlock seek line len date
time
""".split())
DECL = re.compile(r"^\s*(?:Dim|Static|Private|Public|ReDim)\s+(?!Function\b|Sub\b|Property\b|Const\b|Type\b|"
                  r"Enum\b|Declare\b|Preserve\b)(.+)$", re.I | re.M)


def code_only(text):
    """The text without comments and string literals (enough for declarations)."""
    out = []
    for raw in text.splitlines():
        line = re.sub(r'"[^"]*"', '""', raw)
        if "'" in line:
            line = line[:line.index("'")]
        out.append(line)
    return "\n".join(out)


class KeywordNameTests(unittest.TestCase):

    def test_no_variable_is_named_after_a_keyword(self):
        found = []
        for name, text in modules():
            for m in DECL.finditer(code_only(text)):
                for part in m.group(1).split(","):
                    var = re.match(r"\s*(?:WithEvents\s+)?(\w+)", part)
                    if var and var.group(1).lower() in VBA_KEYWORDS:
                        found.append(f"{name}: {var.group(1)}")
        self.assertEqual(found, [])

    def test_the_check_finds_step(self):
        text = "    Dim quoted As String, step As String\n"
        names = [re.match(r"\s*(\w+)", p).group(1) for p in DECL.search(code_only(text)).group(1).split(",")]
        self.assertIn("step", [n.lower() for n in names if n.lower() in VBA_KEYWORDS])


if __name__ == "__main__":
    unittest.main()
