"""The English interface (modLang, modLangData*, tools/i18n.py, tools/i18n_en.py): every Arabic text that
reaches the interface has an English text, the translation leaves nothing Arabic and breaks no SQL, value
list or code, the builders mirror and translate an English file, the code translates its run-time captions
and messages, and modLang gives the same result as the Python mirror in LibreOffice."""
import glob
import os
import random
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import gen_lang
import generate as G
import i18n
import reports as RP
import vba_harness as H
from i18n_en import EN

ORDER = i18n.ordered(EN)
STRINGS = sorted(set(i18n.strings()))


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def statements(text):
    """Logical VBA statements (continuation lines joined) with the name of their procedure."""
    proc, cur = "", []
    for line in text.splitlines():
        head = re.match(r"\s*(?:Public |Private )?(?:Function|Sub) (\w+)", line)
        if head:
            proc = head.group(1)
        cur.append(line)
        if not line.rstrip().endswith(" _"):
            yield proc, "\n".join(cur)
            cur = []


def code_of(statement):
    """The statement without its comments."""
    out = []
    for line in statement.splitlines():
        in_str = False
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str
            if ch == "'" and not in_str:
                line = line[:i]
                break
        out.append(line)
    return "\n".join(out)


def runtime_sql():
    """The SQL that the hand-written code puts together at run time and translates (the literals of one
    statement passed to Tr)."""
    out = []
    for name in G.STATIC_MODULES + ["modAppData"]:
        for _, st in statements(read(name)):
            if "Tr(" not in st:
                continue
            joined = "".join(i18n.vba_literals(st.replace(" _\n", " ")))
            if "SELECT" in joined.upper() and i18n.has_arabic(joined):
                out.append(joined)
    return out


def select_lists(sql):
    for part in re.split(r"\bSELECT\b", sql, flags=re.I)[1:]:
        cols = re.split(r"\bFROM\b", part, maxsplit=1, flags=re.I)[0]
        depth, buf, out = 0, "", []
        for ch in cols:
            depth += (ch == "(") - (ch == ")")
            if ch == "," and depth == 0:
                out.append(buf)
                buf = ""
            else:
                buf += ch
        out.append(buf)
        yield out


class DictionaryTests(unittest.TestCase):

    def test_every_key_has_an_english_text_and_none_is_unused(self):
        keys = i18n.keys()
        self.assertEqual(sorted(keys - set(EN)), [], "add these to tools/i18n_en.py")
        self.assertEqual(sorted(set(EN) - keys), [], "no longer used: remove them from tools/i18n_en.py")

    def test_english_texts_are_safe(self):
        for ar, en in EN.items():
            with self.subTest(ar):
                self.assertTrue(en.strip(), "empty")
                self.assertEqual(en, en.strip())
                self.assertFalse(set(en) & i18n.FORBIDDEN_IN_ENGLISH, en)
                self.assertFalse(i18n.has_arabic(en), en)
                en.encode("ascii")
                self.assertEqual(i18n.fragments(ar), [ar], "a key is one fragment")

    def test_nothing_arabic_is_left(self):
        left = [s for s in STRINGS if i18n.has_arabic(i18n.translate(s, EN, ORDER))]
        self.assertEqual(left[:5], [])

    def test_a_key_never_changes_part_of_a_longer_word(self):
        self.assertEqual(i18n.replace_word("عميلنا عميل", "عميل", "X"), "عميلنا X")
        self.assertEqual(i18n.translate("حفظ", EN, ORDER), EN["حفظ"])
        self.assertEqual(i18n.translate("حفظها", {"حفظ": "Save"}), "حفظها")
        self.assertEqual(i18n.translate("«حفظ»؟", {"حفظ": "Save"}), "Save?")

    def test_value_lists_keep_their_items(self):
        for s in STRINGS:
            if re.fullmatch(r"([A-Z0-9_]+;[^;]+;)*[A-Z0-9_]+;[^;]+", s):
                with self.subTest(s):
                    self.assertEqual(i18n.translate(s, EN, ORDER).count(";"), s.count(";"))

    def test_translated_sql_has_no_duplicate_or_circular_alias(self):
        """Access refuses two columns with one name, and an alias used in its own expression, even as
        q.Cost (error: circular reference caused by alias): choose a caption whose English differs."""
        for sql in [s for s in STRINGS if re.search(r"\bSELECT\b", s)] + runtime_sql():
            english = i18n.translate(sql, EN, ORDER)
            for cols in select_lists(english):
                names = []
                for col in cols:
                    m = re.search(r"\bAS\s+(\[[^\]]+\]|\w+)\s*$", col.strip(), re.I)
                    if not m:
                        continue
                    alias = m.group(1).strip("[]")
                    names.append(alias.lower())
                    expr = col[:m.start()]
                    with self.subTest(col=col.strip()[:80]):     # also t.Cost (the rule of test_queries)
                        self.assertIsNone(re.search(r"(?<![\w\[])\[?" + re.escape(alias) + r"\]?(?![\w(])",
                                                    expr, re.I))
                self.assertEqual(len(names), len(set(names)), cols)

    def test_no_sql_compares_with_an_arabic_value(self):
        """A translated literal in a comparison would no longer find the data."""
        for s in STRINGS + runtime_sql():
            for m in re.finditer(r"(=|<>|\bLike|\bIn\s*\()\s*'([^']*)'", s, re.I):
                self.assertFalse(i18n.has_arabic(m.group(2)), s[:120])


class CodeTests(unittest.TestCase):

    def test_runtime_captions_and_lists_are_translated(self):
        for name in G.STATIC_MODULES:
            if name == "modTestAll":
                continue
            for proc, st in statements(read(name)):
                code = code_of(st)
                if proc.startswith("Test") or "Debug.Print" in code:
                    continue
                for m in re.finditer(r"\.(Caption|RowSource|ControlTipText) = (.*)", code, re.S):
                    rhs = m.group(2)
                    if any(i18n.has_arabic(x) for x in i18n.vba_literals(rhs)) or re.match(r"[A-Z_]{4,}\b", rhs):
                        with self.subTest(module=name, statement=code.strip()[:100]):
                            self.assertTrue(rhs.lstrip().startswith("Tr("), rhs[:80])

    def test_message_boxes_are_translated(self):
        common = read("modCommon")
        self.assertIn("MsgBox Tr(Text), Icon + MSG_RTL, Tr(Title)", common)
        self.assertIn("MsgBox(Tr(Text), vbQuestion", common)
        self.assertNotIn("Const MSG_RTL", common)
        self.assertIn("Public Function MSG_RTL() As Long", read("modLang"))
        for name in G.STATIC_MODULES:
            if name in ("modCommon", "modTestAll"):
                continue
            for proc, st in statements(read(name)):
                code = code_of(st)
                if proc.startswith("Test"):
                    continue
                for m in re.finditer(r"\b(MsgBox|InputBox)\b\(?\s*(.*)", code, re.S):
                    if any(i18n.has_arabic(x) for x in i18n.vba_literals(m.group(2))):
                        with self.subTest(module=name, statement=code.strip()[:100]):
                            self.assertTrue(m.group(2).startswith("Tr("), code[:120])

    def test_reports_text_functions(self):
        reports = read("modReports")
        self.assertIn('ReportCriteria = Tr(Nz(TempVars("ReportCriteria"), ""))', reports)
        self.assertIn("AmountInWords = AmountInWordsEn(Amount)", reports)

    def test_tests_run_in_arabic(self):
        test_all = read("modTestAll")
        self.assertLess(test_all.index('UseLanguage "AR"'), test_all.index("Application.Run"))
        self.assertIn("UseLanguage uiLang", test_all)
        self.assertIn('"TestLang"', test_all)


class BuilderTests(unittest.TestCase):

    def test_forms(self):
        vba = read("modBuildForms")
        for part in ('SetFormProp "Orientation", IIf(UiEnglish(), 0, 1)', "m_frm.Caption = Tr(Caption)",
                     "m_frm.RecordSource = Tr(RecordSource)", "m_frm.Tag = Tr(TagText)",
                     "mdl.AddFromString Tr(Code)", "c.Caption = Tr(Caption)", "c.TextAlign = UiAlign(TextAlign)",
                     "c.RowSource = Tr(Rows)", "c.Properties(PropName).Value = Tr(Value)",
                     "If MirrorLayout() Then x = m_width - L - W Else x = L",
                     "MirrorLayout = (MIRROR_LAYOUT Xor UiEnglish())"):
            self.assertIn(part, vba)
        self.assertNotIn('IIf(MIRROR_LAYOUT, "True"', vba)
        self.assertEqual(re.findall(r"^    c\.RowSource = (?!Tr\()", vba, re.M), [])

    def test_reports(self):
        vba = read("modBuildReports")
        for part in ('SetRptProp "Orientation", IIf(m_english, 0, 1)', "m_rpt.RecordSource = RT(RecordSource)",
                     "c.Caption = RT(Caption)", "Source = RT(Source)", "mdl.AddFromString RT(Code)",
                     "If MIRROR_LAYOUT Xor m_english Then", "c.TextAlign = RAlign(TextAlign)"):
            self.assertIn(part, vba)
        names = {r.name for r in RP.all_reports()}
        self.assertLessEqual(i18n.ARABIC_REPORTS, names)
        for name in i18n.ARABIC_REPORTS:
            self.assertIn("," + name + ",", vba)       # the tax invoice stays Arabic / bilingual (ZATCA)

    def test_queries(self):
        self.assertIn("Sql = Tr(Sql)", read("modBuildQueries"))

    def test_front_end_build_asks_the_language(self):
        with open(os.path.join(ROOT, "dist", "tools", "BuildFrontEnd.vbs"), encoding="utf-8") as fh:
            vbs = fh.read()
        self.assertIn("SetInterfaceLanguage", vbs)
        self.assertLess(vbs.index("SetInterfaceLanguage"), vbs.index('"BuildQueries"'))


class GeneratedTests(unittest.TestCase):

    def test_dictionary_modules(self):
        modules = gen_lang.data_modules()
        files = sorted(glob.glob(os.path.join(ROOT, "src", "vba", "modLangData*.bas")))
        self.assertEqual(len(files), len(modules))
        lang = read("modLang")
        self.assertIn(f"ENTRY_COUNT As Long = {len(EN)}", lang)
        total = 0
        for n, text in enumerate(modules, 1):
            self.assertLess(len(text), gen_lang.MODULE_CHARS + 20000)
            total += text.count("\n    LangAdd ")
            self.assertIn(f"    LangData{n}\n", lang)
        self.assertEqual(total, len(EN))
        # the longest phrases first, across the modules
        added = re.findall(r'LangAdd "((?:[^"]|"")*)"', "".join(modules))
        self.assertEqual([a.replace('""', '"') for a in added], ORDER)


class Static_modLang(VbaModuleChecks, unittest.TestCase):
    module_name = "modLang"
    vba = gen_lang.build_lang_vba()


def _data_case(n, text):
    return type(f"Static_modLangData{n}", (VbaModuleChecks, unittest.TestCase),
                {"module_name": f"modLangData{n}", "vba": text})


for _n, _text in enumerate(gen_lang.data_modules(), 1):
    globals()[f"Static_modLangData{_n}"] = _data_case(_n, _text)


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class LangRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        modules = {"modLang": H.read_module("modLang")}
        for path in sorted(glob.glob(os.path.join(ROOT, "src", "vba", "modLangData*.bas"))):
            name = os.path.basename(path)[:-4]
            modules[name] = H.read_module(name)
        modules["Driver"] = ("Option VBASupport 1\n"
                             "Public Function RunTr(ByVal LangCode As String, ByVal Text As String) As String\n"
                             "    UseLanguage LangCode\n    RunTr = Tr(Text)\nEnd Function\n"
                             "Public Function RunAlign(ByVal LangCode As String, ByVal a As Integer) As Integer\n"
                             "    UseLanguage LangCode\n    RunAlign = UiAlign(a)\nEnd Function\n"
                             "Public Function RunRtl(ByVal LangCode As String) As Long\n"
                             "    UseLanguage LangCode\n    RunRtl = MSG_RTL\nEnd Function\n")
        cls.h = H.Harness()
        cls.h.load(modules)

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_python_mirror(self):
        rnd = random.Random(3)
        sample = rnd.sample([s for s in STRINGS if len(s) < 600], 150)
        sample += ["TEST عميل 12.50: «حفظ»؟", "عميلنا حفظ", "حفظ" + "ية"]
        for text in sample:
            with self.subTest(text[:60]):
                self.assertEqual(self.h.call("Driver", "RunTr", "EN", text), i18n.translate(text, EN, ORDER))

    def test_arabic_file_is_unchanged(self):
        for text in STRINGS[:20] + [s for s in STRINGS if "[@" in s][:5]:
            self.assertEqual(self.h.call("Driver", "RunTr", "AR", text), i18n.resolve_names(text))

    def test_direction(self):
        self.assertEqual([self.h.call("Driver", "RunAlign", "EN", a) for a in (0, 1, 2, 3)], [0, 3, 2, 1])
        self.assertEqual([self.h.call("Driver", "RunAlign", "AR", a) for a in (0, 1, 2, 3)], [0, 1, 2, 3])
        self.assertEqual(self.h.call("Driver", "RunRtl", "EN"), 0)
        self.assertEqual(self.h.call("Driver", "RunRtl", "AR"), 0x180000)


if __name__ == "__main__":
    unittest.main()
