"""Reference invoice calculation (sales and purchases). modSales.bas implements
the same algorithm in VBA; the VBA tests replay the cases generated here.

Inputs per line: quantity, unit price (as typed: incl. VAT when prices
include VAT), line discount amount (same basis as the price), VAT rate.
Invoice discount: an amount on the same basis, spread over the lines in
proportion to their value (remainder to the largest line).

Stored per line (all excl. VAT): UnitPrice (4 dp), Discount, NetAmount, Tax,
LineTotal = NetAmount + Tax. When prices include VAT the customer pays
exactly the shelf price: LineTotal is fixed first, then Tax is extracted.
"""

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from typing import List

D = Decimal


def r2(x) -> Decimal:
    return D(x).quantize(D("0.01"), rounding=ROUND_HALF_UP)


def r4(x) -> Decimal:
    return D(x).quantize(D("0.0001"), rounding=ROUND_HALF_UP)


@dataclass
class LineIn:
    qty: Decimal
    price: Decimal
    discount: Decimal = D(0)
    rate: Decimal = D("0.15")


@dataclass
class LineOut:
    unit_price: Decimal      # excl. VAT, 4 dp
    discount: Decimal        # excl. VAT
    net: Decimal
    tax: Decimal
    total: Decimal           # incl. VAT


@dataclass
class InvoiceOut:
    lines: List[LineOut]
    subtotal: Decimal
    discount: Decimal
    taxable: Decimal
    tax: Decimal
    total: Decimal


class PricingError(ValueError):
    pass


def calc(lines: List[LineIn], invoice_discount=D(0), prices_include_vat=True) -> InvoiceOut:
    if not lines:
        raise PricingError("no lines")
    gross = []
    for ln in lines:
        if ln.qty <= 0 or ln.price < 0 or ln.discount < 0:
            raise PricingError("invalid line")
        g = r2(ln.qty * ln.price) - ln.discount
        if g < 0:
            raise PricingError("line discount larger than the line")
        gross.append(g)
    total_gross = sum(gross, D(0))
    if invoice_discount < 0 or invoice_discount > total_gross:
        raise PricingError("invoice discount larger than the invoice")

    # spread the invoice discount, remainder on the largest line
    shares = [D(0)] * len(lines)
    if invoice_discount > 0:
        for i, g in enumerate(gross):
            shares[i] = r2(invoice_discount * g / total_gross)
        biggest = max(range(len(lines)), key=lambda i: (gross[i], -i))
        shares[biggest] += invoice_discount - sum(shares, D(0))

    out = []
    for ln, g, share in zip(lines, gross, shares):
        after = g - share                          # same basis as the price
        if prices_include_vat:
            total = after
            tax = r2(total * ln.rate / (1 + ln.rate))
            net = total - tax
            unit = r4(ln.price / (1 + ln.rate))
        else:
            net = after
            tax = r2(net * ln.rate)
            total = net + tax
            unit = r4(ln.price)
        disc = r2(ln.qty * unit) - net
        if disc < 0:
            # rounding pushed the net above qty x price: lift the unit price
            unit = r4(net / ln.qty)
            disc = max(D(0), r2(ln.qty * unit) - net)
        out.append(LineOut(unit, disc, net, tax, total))

    return InvoiceOut(
        lines=out,
        subtotal=sum((x.net + x.discount for x in out), D(0)),
        discount=sum((x.discount for x in out), D(0)),
        taxable=sum((x.net for x in out), D(0)),
        tax=sum((x.tax for x in out), D(0)),
        total=sum((x.total for x in out), D(0)),
    )


def settle(total: Decimal, tendered: Decimal, credit: bool):
    """Returns (paid, remaining, change). Cash sales must be fully paid."""
    tendered = max(D(0), tendered)
    if not credit and tendered < total:
        raise PricingError("cash sale not fully paid")
    paid = min(tendered, total)
    return paid, total - paid, max(D(0), tendered - total)


def return_amounts(sold_qty, sold_total, sold_tax, prev_qty, prev_total, prev_tax, qty):
    """(net, tax, total) for returning `qty` units of an original sales line.
    Partial returns are proportional; the return that completes the line takes
    exactly what is left, so all returns together equal the original line."""
    if qty <= 0 or prev_qty + qty > sold_qty:
        raise PricingError("return quantity")
    if prev_qty + qty == sold_qty:
        total = sold_total - prev_total
        tax = sold_tax - prev_tax
    else:
        total = r2(sold_total * qty / sold_qty)
        tax = r2(sold_tax * qty / sold_qty)
    return total - tax, tax, total


def weighted_average(cur_qty, cur_avg, qty, unit_cost):
    """Average cost after `qty` units (signed) at `unit_cost` (modSales.WeightedAverage).
    Incoming: weighted average, or the new cost when there was no stock.
    Outgoing at a known cost (purchase return): the same formula while stock
    remains and the result is not negative; otherwise the average is kept."""
    cur_qty, cur_avg, qty, unit_cost = D(cur_qty), D(cur_avg), D(qty), D(unit_cost)
    new_qty = cur_qty + qty
    if qty > 0:
        if cur_qty <= 0:
            return unit_cost
        return r4((cur_qty * cur_avg + qty * unit_cost) / new_qty)
    if qty < 0 and new_qty > 0 and cur_qty > 0:
        v = r4((cur_qty * cur_avg + qty * unit_cost) / new_qty)
        if v >= 0:
            return v
    return cur_avg


def purchase_unit_cost(net, total, qty, vat_registered):
    """Cost of one purchased unit: excl. VAT when the store reclaims input VAT."""
    qty = D(qty)
    if qty <= 0:
        return D(0)
    return r4((D(net) if vat_registered else D(total)) / qty)


AVERAGE_CASES = [
    # (label, current qty, current average, qty signed, unit cost)
    ("أول شراء بدون رصيد", "0", "0", "100", "52.1739"),
    ("شراء ثانٍ بسعر أعلى", "90", "50", "10", "60"),
    ("كسور تحتاج تقريب 4 خانات", "3", "10", "7", "10.3333"),
    ("شراء بعد رصيد سالب", "-5", "40", "20", "45"),
    ("مرتجع شراء بتكلفة الشراء", "100", "55", "-10", "60"),
    ("مرتجع يُفرغ المخزون يبقي المتوسط", "10", "55", "-10", "60"),
    ("مرتجع ينتج متوسطًا سالبًا يبقي المتوسط", "2", "1", "-1", "5"),
    ("حركة صفرية", "10", "55", "0", "99"),
]


RETURN_CASES = [
    # sold (qty, total, tax) and successive return quantities
    ("إرجاع 3 من 3 دفعة واحدة", ("3", "29.97", "3.91"), ["3"]),
    ("إرجاع 1 ثم 1 ثم 1 من 3", ("3", "29.97", "3.91"), ["1", "1", "1"]),
    ("إرجاع 0.125 ثم 0.25 كيلو", ("0.375", "4.87", "0.64"), ["0.125", "0.25"]),
]


# --------------------------------------------------------------------------
# Test cases shared with the VBA tests
# --------------------------------------------------------------------------
CASES = [
    # (label, prices_include_vat, invoice_discount, [(qty, price, discount, rate)])
    ("سعر شامل 11.50 × 1", True, "0", [("1", "11.50", "0", "0.15")]),
    ("سعر شامل 9.99 × 3", True, "0", [("3", "9.99", "0", "0.15")]),
    ("ثلاثة أسطر وخصم فاتورة 10", True, "10", [("2", "25.00", "0", "0.15"), ("1", "7.95", "0", "0.15"),
                                             ("5", "3.45", "1.00", "0.15")]),
    ("كمية كسرية 0.375 كيلو × 12.99", True, "0", [("0.375", "12.99", "0", "0.15")]),
    ("صنف معفى مع صنف خاضع", True, "0", [("2", "10.00", "0", "0"), ("1", "23.00", "0", "0.15")]),
    ("أسعار غير شاملة 100 × 2 خصم 5", False, "0", [("2", "100.00", "5.00", "0.15")]),
    ("غير شاملة مع خصم فاتورة 3.33", False, "3.33", [("1", "19.99", "0", "0.15"), ("3", "4.10", "0", "0.15")]),
    ("فاتورة 1150 للتحقق من QR", True, "0", [("10", "115.00", "0", "0.15")]),
    ("خصم يجعل السطر صفرًا", True, "0", [("1", "5.00", "5.00", "0.15"), ("1", "2.30", "0", "0.15")]),
    ("تقريب حدّي 0.05 × 7", True, "0", [("7", "0.05", "0", "0.15")]),
]


def case_inputs(case):
    label, incl, inv_disc, lines = case
    return ([LineIn(D(q), D(p), D(d), D(r)) for q, p, d, r in lines], D(inv_disc), incl)
