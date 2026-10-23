"""Python mirror of the rules of modQuotations (docs/75, customer quotations)."""
from datetime import date
from decimal import Decimal
from typing import Iterable, Optional, Tuple

import pricing


def state_for(status: str, valid_until: Optional[date], today: date) -> str:
    """QuotationStateFor: an issued quotation past its validity date is EXPIRED."""
    if status != "ISSUED" or valid_until is None:
        return status
    return "EXPIRED" if valid_until < today else status


def convert_problem(state: str) -> Optional[str]:
    """QuotationConvertProblem: only an issued quotation still valid becomes an order or an invoice."""
    return None if state == "ISSUED" else state


def issue_totals(lines: Iterable[Tuple], vat_rate, prices_include_vat=True):
    """IssueQuotation: (qty, price, standard rated) lines priced by the invoice engine.
    Returns the lines (net, tax, total) and the totals kept on the quotation: SubTotal = the taxable amount,
    so that SubTotal + Tax = TotalAmount."""
    rate = Decimal(str(vat_rate))
    ins = [pricing.LineIn(Decimal(str(q)), Decimal(str(p)), Decimal(0), rate if std else Decimal(0))
           for q, p, std in lines]
    out = pricing.calc(ins, Decimal(0), prices_include_vat)
    return [(x.net, x.tax, x.total) for x in out.lines], (out.taxable, out.tax, out.total)
