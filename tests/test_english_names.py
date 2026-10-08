"""The English names screen (frmEnglishNames, modEnglishNames, docs/42-English-Names-Screen.md): the names without
an English name, a suggested transliteration (tools/translit.py is the Python mirror of Transliterate) and one
save, written to each table and to the audit trail."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import forms as F
import generate as G
import i18n
import vba_harness as H
from master_en import ENGLISH_NAMES, NAMES_SCREEN_TABLES, names_screen_const, names_screen_specs
from schema import SCREEN_LIST, table
from translit import LETTERS, WORDS, transliterate, vba_table


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modEnglishNames")


def const(name):
    m = re.search(rf"Private Const {name} As String = (.*?)\n(?!    \")", VBA, re.S)
    return "".join(re.findall(r'"((?:[^"]|"")*)"', m.group(1))).replace('""', '"')


SAMPLES = ["محمد عبدالله الغامدي", "مؤسسة النور للتجارة", "شركة الراجحي وشركاه", "صندوق الكاشير ٢",
           "عبدالكريم", "بنك الرياض - فرع جدة", "Ali محمد", "فاطمة الزهراء", "وليد", "يزيد", "عُمَر",
           "مطعم  البيت،  الدمام؟", "عبد", "ال", "أبو بكر الصديق", "مكتبة ـ الأمل 15", "", "ABC 123",
           "آل سعود", "مؤمن", "عائشة بنت أبي بكر", "عبدالرحمن بن ناصر", "ذهب وفضة", "ظافر الشهري"]


class TransliterationTests(unittest.TestCase):

    def test_examples(self):
        self.assertEqual(transliterate("محمد عبدالله الغامدي"), "Mohammed Abdullah Al-Ghamdi")
        self.assertEqual(transliterate("مؤسسة النور للتجارة"), "Establishment Al-Nor Trading")
        self.assertEqual(transliterate("بنك الرياض - فرع جدة"), "Bank Riyadh - Branch Jeddah")
        self.assertEqual(transliterate("صندوق الكاشير ٢"), "Box Al-Cashier 2")
        self.assertEqual(transliterate("عُمَر"), "Omar")                         # diacritics are dropped
        self.assertEqual(transliterate("عبدالكريم"), "Abdulkrim")
        self.assertEqual(transliterate("وليد"), "Wlid")                         # waw at the start = w
        self.assertEqual(transliterate("Ali محمد"), "Ali Mohammed")             # Latin text stays
        self.assertEqual(transliterate("مطعم  البيت،  الدمام؟"), "Restaurant Al-Bit, Dammam?")

    def test_every_sample_is_latin(self):
        for s in SAMPLES:
            self.assertFalse(i18n.has_arabic(transliterate(s)), s)

    def test_vba_tables_are_the_python_tables(self):
        self.assertEqual(const("TRANSLIT_LETTERS"), vba_table(LETTERS))
        self.assertEqual(const("TRANSLIT_WORDS"), vba_table(WORDS))
        for k, v in list(LETTERS.items()) + list(WORDS.items()):
            self.assertFalse(set(k + v) & set("|="), k)
            k.encode("cp1256")
            v.encode("ascii")

    def test_arabic_digits_are_not_in_the_code(self):
        """Windows-1256 has no Arabic-Indic digits: the code tests their character codes."""
        self.assertIn("ElseIf code >= &H660 And code <= &H669 Then", VBA)
        VBA.encode("cp1256")


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class TransliterationRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modEnglishNames": VBA})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_python_mirror(self):
        for s in SAMPLES:
            with self.subTest(s):
                self.assertEqual(self.h.call("modEnglishNames", "Transliterate", s), transliterate(s))


class ScreenTests(unittest.TestCase):

    def test_tables_of_the_screen(self):
        self.assertIn(f'Private Const NAME_TABLES As String = "', VBA)
        self.assertEqual(const("NAME_TABLES"), names_screen_const())
        for tbl in NAMES_SCREEN_TABLES:
            self.assertTrue(tbl in ENGLISH_NAMES or tbl == "Products", tbl)
        for tbl, key, ar, en, size, caption in names_screen_specs():
            fields = {f.name: f for f in table(tbl).fields}
            self.assertIn(fields[key].kind, ("AUTO", "LONG"), tbl)     # the key is written without quotes
            self.assertEqual(fields[en].size, size)
            self.assertFalse(fields[en].required)

    def test_screen_is_registered(self):
        names = {m.name for m in F.all_forms()}
        self.assertTrue({"frmEnglishNames", "frmEnglishNameLines"} <= names)
        self.assertEqual(F.SCREEN_PERMISSIONS["frmEnglishNames"], "SETTINGS")
        self.assertIn(("frmEnglishNames", "الأسماء الإنجليزية", "النظام", "SETTINGS", False, True, False), SCREEN_LIST)
        settings = next(m for m in F.all_forms() if m.name == "frmSettings")
        self.assertTrue(any(c.name == "btnEnglishNames" for c in settings.controls))

    def test_local_table(self):
        pos = read("modPOS")
        create = re.search(r'CREATE TABLE tmpEnglishNames \((.*?)\)", dbFailOnError', pos, re.S).group(1)
        for f in ("LineNo COUNTER", "TableName", "TableTitle", "KeyValue LONG", "ArabicName", "EnglishName",
                  "OldEnglish", "MaxLen"):
            self.assertIn(f, create)

    def test_suggest_fills_only_empty_names(self):
        body = re.search(r"Public Sub EnglishNamesSuggest.*?End Sub", VBA, re.S).group(0)
        self.assertIn("WHERE EnglishName Is Null", body)
        self.assertIn("Left$(Transliterate(", body)                     # cut to the size of the field

    def test_save_checks_lengths_first_and_audits(self):
        body = re.search(r"Public Function SaveEnglishNameRows.*?End Function", VBA, re.S).group(0)
        self.assertLess(body.index("If Len(newName) > rs!MaxLen Then"), body.index("db.Execute"))
        self.assertLess(body.index("AuditSnapshot("), body.index("db.Execute"))
        self.assertLess(body.index("db.Execute"), body.index('AuditEdited "EDIT"'))
        self.assertIn("SqlText(newName)", body)
        self.assertIn("vbBinaryCompare", body)                    # "Ali" -> "ALI" is a change
        self.assertIn('CanScreenAction(frm.Name, "EDIT")', VBA)

    def test_in_access_test_is_run(self):
        self.assertIn('"TestEnglishNames"', read("modTestAll"))
        self.assertIn("modEnglishNames", G.STATIC_MODULES)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modEnglishNames"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
