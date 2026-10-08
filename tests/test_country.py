"""The operating country (modCountry, docs/44-Operating-Country.md): Settings.CountryCode SA / EG decides the
program currency, the VAT rate, the tax number, the mobile format, the amount in words, the title and QR code
of the sales documents; it changes only before the first operation."""
import os
import re
import unittest

from helpers import ROOT, VbaModuleChecks
import country_reference as C
import forms as F
import generate as G
import reports as RP
import vba_harness as H
from schema import RULE_UPGRADES, TABLES, table


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


VBA = read("modCountry")
TAX_SAMPLES = ["", "300000000000003", "310123456700003", "300000000000004", "200000000000003", "123456789",
               "12345678", "1234567890", "12345678A", "30000000000003", " 123456789 "]
MOBILE_SAMPLES = ["0501234567", "+966501234567", "966501234567", "05012345678", "01012345678", "+201012345678",
                  "201012345678", "0101234567", "1012345678"]


class CountryReferenceTests(unittest.TestCase):

    def test_rules(self):
        self.assertEqual(C.tax_number_problem("300000000000003", "SA"), "")
        self.assertTrue(C.tax_number_problem("123456789", "SA"))
        self.assertEqual(C.tax_number_problem("123456789", "EG"), "")
        self.assertTrue(C.tax_number_problem("300000000000003", "EG"))
        self.assertTrue(C.mobile_fits("01012345678", "EG") and not C.mobile_fits("01012345678", "SA"))
        self.assertEqual(C.doc_title("EG", "SALE", "SIMPLIFIED", False), "إيصال بيع")
        self.assertEqual(C.doc_title("SA", "SALE", "SIMPLIFIED", True), "Simplified Tax Invoice")
        self.assertEqual(C.doc_title("EG", "RETURN", "SIMPLIFIED", False), "إشعار دائن")


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class CountryRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCountry": VBA})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_same_as_the_python_mirror(self):
        for code in ("SA", "EG"):
            for s in TAX_SAMPLES:
                with self.subTest(code=code, tax=s):
                    self.assertEqual(self.h.call("modCountry", "TaxNumberProblem", s, code),
                                     C.tax_number_problem(s, code))
            for m in MOBILE_SAMPLES:
                with self.subTest(code=code, mobile=m):
                    self.assertEqual(self.h.call("modCountry", "MobileFits", m, code), C.mobile_fits(m, code))
            for kind in ("SALE", "RETURN"):
                for sub in ("SIMPLIFIED", "STANDARD"):
                    for en in (False, True):
                        self.assertEqual(self.h.call("modCountry", "DocTitleFor", code, kind, sub, en),
                                         C.doc_title(code, kind, sub, en))
            self.assertAlmostEqual(self.h.call("modCountry", "CountryVatRate", code), C.VAT_RATE[code])
            self.assertEqual(self.h.call("modCountry", "CountryCurrency", code), C.CURRENCY[code])


class SchemaTests(unittest.TestCase):

    def test_country_field(self):
        f = next(x for x in table("Settings").fields if x.name == "CountryCode")
        self.assertEqual(f.rule, 'In ("SA","EG")')
        self.assertEqual(f.default, '"SA"')
        self.assertIn(("EGP", "جنيه مصري"), {r[:2] for r in table("Currencies").seed_rows})

    def test_rules_reach_an_existing_back_end(self):
        """BuildSchema keeps existing fields: the changed rules are set by UpgradeFieldRules."""
        vba = read("modBuildSchema")
        self.assertIn("    UpgradeFieldRules\n", vba)
        body = re.search(r"Private Sub UpgradeFieldRules\(\).*?End Sub", vba, re.S).group(0)
        for tbl, name in RULE_UPGRADES:
            self.assertIn(f'SetFieldRule "{tbl}", "{name}", ', body)
        # looser only: every Saudi number the old rule allowed is still allowed
        for tbl, name in RULE_UPGRADES:
            rule = next(x for x in table(tbl).fields if x.name == name).rule
            self.assertTrue(rule.startswith("In (") or 'Like "3#############3"' in rule, rule)

    def test_operation_tables(self):
        """Every table of documents blocks a change of country, among them every table with a currency."""
        tables = re.search(r'OPERATION_TABLES As String = (.*?)\n\n', VBA, re.S).group(1)
        names = set("".join(re.findall(r'"([^"]*)"', tables)).split(","))
        self.assertTrue(names <= {t.name for t in TABLES}, names - {t.name for t in TABLES})
        with_rate = {t.name for t in TABLES if any(f.name == "ExchangeRate" for f in t.fields)}
        with_rate -= {"JournalEntries", "ManualEntries"} - names
        self.assertTrue(with_rate <= names, with_rate - names)
        self.assertTrue({"SalesInvoices", "SalesReturns", "JournalEntries"} <= names)


class CodeTests(unittest.TestCase):

    def test_settings_screen(self):
        model = next(m for m in F.all_forms() if m.name == "frmSettings")
        combo = next(c for c in model.controls if c.name == "CountryCode")
        self.assertTrue(combo.props["RowSource"].startswith("SA;"))
        forms_ = read("modForms")
        validate = re.search(r"Private Function ValidateSettings.*?End Function", forms_, re.S).group(0)
        self.assertIn("If Not SettingsCountryCheck(frm) Then Exit Function", validate)
        self.assertIn('TaxNumberProblem(frm!VATNumber.Value, Nz(frm!CountryCode.Value, "SA"))', validate)
        self.assertNotIn("0.15", validate)
        self.assertIn("SettingsAfterSave", forms_)
        partner = re.search(r"Private Function ValidatePartner.*?End Function", forms_, re.S).group(0)
        self.assertIn("TaxNumberProblem(frm!VATNumber.Value)", partner)
        self.assertIn("MobileFits(mobile)", partner)

    def test_country_change_sets_rate_currency_and_defaults(self):
        check = re.search(r"Public Function SettingsCountryCheck.*?End Function", VBA, re.S).group(0)
        self.assertLess(check.index("CountryChangeProblem()"), check.index("AskYesNo("))
        self.assertIn("frm!VATRate.Value = CountryVatRate(newCountry)", check)
        self.assertIn("frm!CurrencyCode.Value = CountryCurrency(newCountry)", check)
        apply_ = re.search(r"Public Function ApplyCountryData.*?End Function", VBA, re.S).group(0)
        self.assertIn('fld.Name = "CurrencyCode"', apply_)
        self.assertIn("fld.DefaultValue = ", apply_)
        self.assertIn("UPDATE Suppliers SET CurrencyCode", apply_)

    def test_zatca_qr_only_in_saudi_arabia(self):
        self.assertIn('If Len(vatNo) > 0 And AppCountry() = "SA" Then', read("modSales"))

    def test_documents_follow_the_country(self):
        self.assertEqual(RP.TITLE_AR, "=DocTitleAr([DocKind],[InvoiceSubType])")
        self.assertEqual(RP.TITLE_EN, "=DocTitleEn([DocKind],[InvoiceSubType])")
        reports = read("modReports")
        self.assertIn("If Len(CurrencyCode) = 0 Then CurrencyCode = BaseCurrency()", reports)
        import queries as Q
        moves = next(q for q in Q.QUERIES if q.name == "qrySupplierFxMoves").sql
        self.assertNotIn("'SAR'", moves)

    def test_no_fixed_riyal_in_screens_and_reports(self):
        """The program currency follows the country: no "ريال" left in captions, lists and reports."""
        texts = []
        for m in F.all_forms():
            texts += [m.caption, m.tag] + [v for c in m.controls for v in c.props.values() if isinstance(v, str)]
        for r in RP.all_reports():
            texts += [c.source or "" for cs in r.controls.values() for c in cs]
            texts += [v for cs in r.controls.values() for c in cs for v in c.props.values() if isinstance(v, str)]
        self.assertEqual([t for t in texts if "ريال" in (t or "")], [])

    def test_in_access_test_is_run(self):
        self.assertIn('"TestCountry"', read("modTestAll"))
        self.assertIn("modCountry", G.STATIC_MODULES)


class StaticModule(VbaModuleChecks, unittest.TestCase):
    module_name = "modCountry"
    vba = VBA


if __name__ == "__main__":
    unittest.main()
