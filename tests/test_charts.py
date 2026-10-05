"""Statistics report (modCharts): helpers checked against tools/charts_reference.py through
LibreOffice, report wiring, the category query, and static checks."""
import datetime as dt
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, HERE)

import charts_reference as C
import forms as F
import queries as Q
import reports as RP
import vba_harness as H
from helpers import VbaModuleChecks

DRIVER = r'''
Option VBASupport 1

Public Function Ping() As Long
    Ping = 42
End Function
'''


def read(name):
    with open(os.path.join(HERE, "..", "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


def ole_date(d: dt.date) -> float:
    return float((d - dt.date(1899, 12, 30)).days)


class ReferenceTests(unittest.TestCase):

    def test_nice_max(self):
        for v, exp in [(0, 1), (-5, 1), (1, 1), (1.2, 2), (2, 2), (2.2, 2.5), (3, 5), (7, 10), (10, 10),
                       (11, 20), (950, 1000), (3289.9, 5000), (24000, 25000), (48750, 50000)]:
            self.assertEqual(C.nice_max(v), exp, v)

    def test_chart_months(self):
        self.assertEqual(C.chart_months(dt.date(2026, 10, 5)).split(",")[0], "202511")
        self.assertEqual(C.chart_months(dt.date(2026, 10, 5)).split(",")[-1], "202610")
        self.assertEqual(C.chart_months(dt.date(2026, 3, 31)).split(","),
                         ["2025%02d" % m for m in range(4, 13)] + ["202601", "202602", "202603"])


@unittest.skipUnless(H.available(), "LibreOffice (soffice + python3-uno) is not installed")
class ChartRuntimeTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.h = H.Harness()
        cls.h.load({"modCommon": H.read_module("modCommon"), "modCharts": H.read_module("modCharts"),
                    "Driver": DRIVER})

    @classmethod
    def tearDownClass(cls):
        cls.h.close()

    def test_nice_max_matches(self):
        for v in (0.0, 0.4, 1.0, 1.7, 2.3, 9.9, 99.0, 101.0, 3289.9, 48750.0, 123456.0):
            self.assertAlmostEqual(self.h.call("modCharts", "NiceMax", v), C.nice_max(v), places=6, msg=v)

    def test_months_match(self):
        for d in (dt.date(2026, 10, 5), dt.date(2026, 1, 1), dt.date(2025, 12, 31), dt.date(2024, 2, 29)):
            self.assertEqual(self.h.call("modCharts", "ChartMonths", ole_date(d)), C.chart_months(d), d)

    def test_month_names(self):
        self.assertEqual(self.h.call("modCharts", "MonthShortName", 1), "يناير")
        self.assertEqual(self.h.call("modCharts", "MonthShortName", 12), "ديسمبر")


class WiringTests(unittest.TestCase):

    def test_report_draws_in_its_detail_section(self):
        m = next(r for r in RP.all_reports() if r.name == "rptStatistics")
        self.assertEqual(m.record_source, "")
        self.assertIn('m_rpt.Section(0).Name = "secCharts"', m.events)
        self.assertIn("m_rpt.Section(0).OnPrint = EP", m.events)
        code = "\n".join(m.code)
        self.assertIn("Private Sub secCharts_Print(", code)
        self.assertIn("DrawStatistics Me, Me!boxCharts.Left", code)
        self.assertIn("Public Sub DrawStatistics(", read("modCharts"))

    def test_report_centre_entry(self):
        entry = next(r for r in F.REPORTS if r.key == "STATISTICS")
        self.assertEqual((entry.report, entry.needs), ("rptStatistics", "P"))

    def test_data_comes_from_tested_queries(self):
        body = read("modCharts")
        names = {q.name for q in Q.QUERIES}
        for used in re.findall(r"FROM (\w+Query)", body):
            self.assertIn(used, names)
        for col in ("CategoryName", "CategoryTotal"):
            self.assertIn(col, Q.query("SalesByCategoryQuery").sql)

    def test_errors_never_stop_printing(self):
        body = read("modCharts").split("Public Sub DrawStatistics(")[1].split("\nEnd Sub")[0]
        self.assertIn("On Error GoTo EH", body)


class Static_modCharts(VbaModuleChecks, unittest.TestCase):
    module_name = "modCharts"
    vba = read("modCharts")
