"""A ByRef variable given to two calls of one expression and also read there.

In Access, `Record F(a, st) = "OK" And st = "X" And F(b, st) = "OK" And st = "Y"` failed: VBA read `st`
after both calls (TestZatcaApi "قبول الاعتماد والتبليغ", TestEInvoice "قراءة نتيجة المنظومة").
Each call that fills a ByRef variable goes in its own statement.
"""
import glob
import os
import re
import unittest

from helpers import ROOT

PROC = re.compile(r"^(?:Public|Private)\s+(?:Function|Sub)\s+(\w+)\s*\((.*?)\)", re.M)
STRING = re.compile(r'"(?:[^"]|"")*"')


def modules():
    for path in sorted(glob.glob(os.path.join(ROOT, "src", "vba", "*.bas"))):
        with open(path, encoding="utf-8") as fh:
            yield os.path.basename(path)[:-4], re.sub(r" _\r?\n\s*", " ", fh.read())


def byref_positions():
    """{procedure: positions of its ByRef parameters} (written ByRef; VBA's default ByRef is not used here)."""
    out = {}
    for _, text in modules():
        for m in PROC.finditer(text):
            params = [p.strip() for p in m.group(2).split(",")] if m.group(2).strip() else []
            pos = {i for i, p in enumerate(params) if p.startswith("ByRef ")}
            if pos:
                out[m.group(1)] = pos
    return out


def calls(statement):
    """(name, [arguments]) of each call with no nested brackets in its arguments."""
    for m in re.finditer(r"\b(\w+)\(([^()]*)\)", statement):
        yield m.group(1), [a.strip() for a in m.group(2).split(",")]


def problems():
    refs = byref_positions()
    found = []
    for name, text in modules():
        for line in text.split("\n"):
            code = STRING.sub('""', line)
            if code.lstrip().startswith("'"):
                continue
            filled = {}
            for proc, args in calls(code):
                for i in refs.get(proc, ()):
                    if i < len(args) and re.fullmatch(r"[A-Za-z]\w*", args[i]):
                        filled[args[i]] = filled.get(args[i], 0) + 1
            for var, n in filled.items():
                if n >= 2 and re.search(r"\b%s\s*(=|<>)" % var, code):
                    found.append(f"{name}: {line.strip()[:120]}")
    return found


class ByRefTests(unittest.TestCase):

    def test_the_byref_list(self):
        refs = byref_positions()
        self.assertEqual(refs["ZatcaReplyResult"], {3, 4})
        self.assertEqual(refs["SplitResult"], {1})

    def test_one_call_per_byref_variable_in_an_expression(self):
        self.assertEqual(problems(), [])


if __name__ == "__main__":
    unittest.main()
