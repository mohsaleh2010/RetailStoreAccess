"""The cash day book (CashDayBookQuery, rptCashDayBook, docs/55-Cash-Day-Book.md), on the demo data replayed in
the SQLite mirror: one group per box and day, the opening of each day = every move of the box before it, the
closing of a day = the opening of its next day, the receipts before the payments; the report and its entries."""
import os
import re
import unittest
from collections import defaultdict

from helpers import ROOT
from access_sqlite import day
import demo_data as DD
import forms as F
import reports as RP


def read(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()


class DayBookTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = DD.simulate()
        cls.c = cls.r.db.con

    def rows(self, box=0, days_ago=60):
        self.r.db.set_period(days_ago)
        self.r.db.params["CashBoxID"] = box
        return self.c.execute('SELECT DayKey, CashBoxID, CashDay, SideNo, AmountIn, AmountOut, DayOpening, '
                              'ClosingCount FROM "CashDayBookQuery"').fetchall()

    def before(self, box, when):
        return self.c.execute("SELECT Coalesce(Sum(AmountIn - AmountOut), 0) FROM qryCashMovements "
                              "WHERE CashBoxID = ? AND MoveDate < ?", (box, when)).fetchone()[0]

    def test_days_chain_and_openings(self):
        rows = self.rows()
        self.assertTrue(rows, "the demo data has cash moves")
        days = defaultdict(list)
        for r in rows:
            days[r[0]].append(r)
        by_box = defaultdict(list)
        for key, items in days.items():
            box, cash_day = items[0][1], items[0][2]
            self.assertEqual({(i[1], i[2]) for i in items}, {(box, cash_day)}, "one box and day per key")
            self.assertEqual({i[6] for i in items}, {items[0][6]}, "one opening per day")
            self.assertAlmostEqual(items[0][6], self.before(box, cash_day), places=2)
            for i in items:
                self.assertEqual(i[3], 1 if i[4] > 0 else 2)
            by_box[box].append((cash_day, items[0][6], sum(i[4] - i[5] for i in items)))
        for box, seq in by_box.items():
            seq.sort()
            for (d1, opening, net), (d2, next_opening, _) in zip(seq, seq[1:]):
                self.assertAlmostEqual(opening + net, next_opening, places=2, msg=f"box {box}: {d1} -> {d2}")

    def test_one_box(self):
        everything = self.rows()
        box = everything[0][1]
        self.assertEqual({r[1] for r in self.rows(box)}, {box})

    def test_period_limits(self):
        self.assertEqual(self.rows(days_ago=-5), [])                  # a period after every move


class DayBookCodeTests(unittest.TestCase):

    def test_report_and_entry(self):
        e = next(x for x in F.REPORTS if x.key == "CASH_DAYBOOK")
        self.assertEqual((e.query, e.report, e.needs), ("CashDayBookQuery", "rptCashDayBook", "Pb#"))
        m = next(x for x in RP.all_reports() if x.name == "rptCashDayBook")
        self.assertEqual(m.group, "DayKey")
        self.assertEqual(m.sorts[0], ("SideNo", False))
        self.assertTrue(any("Section(6).ForceNewPage = 2" in ev for ev in m.events))
        sources = {c.source for sec in m.controls.values() for c in sec}
        for part in ("DayOpening", "=Sum([AmountIn])", "=Sum([AmountOut])",
                     "=[DayOpening]+Sum([AmountIn])-Sum([AmountOut])"):
            self.assertIn(part, sources)

    def test_treasury_prints_the_day_book(self):
        body = re.search(r"Public Sub PrintCashDay\(.*?End Sub", read("modCash"), re.S).group(0)
        self.assertIn('"rptCashDayBook", "CashDayBookQuery"', body)
        self.assertIn('SetQueryParam "CashBoxID"', body)
        self.assertIn("SetPeriod d, d", body)


if __name__ == "__main__":
    unittest.main()
