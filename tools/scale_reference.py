"""Python mirror of the scale barcodes (modScale, docs/86): the labels a scale prints for weighed goods.

A scale barcode is an EAN-13 of 13 digits: a prefix of the settings (LabelSettings.ScalePrefixes, e.g. "20,21"),
the item code (ScaleItemDigits digits = the code of the product), the value (the weight in grams or the price in
cents / piasters, ScaleValueKind WEIGHT / PRICE) and the check digit. The product is the active one whose code is
the item code (as printed, else without its leading zeros). A price label gives the quantity: price / the price of a
new line for the customer, rounded to 3 decimals. The same rule is in VBA (modScale.ScaleCodeParts / ScaleLine) and
in C# (RetailStore.Domain.Sales.ScaleBarcode); the vectors below reach xUnit through gen_web."""
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Optional, Tuple


def check_digit(digits12: str) -> int:
    """The EAN-13 check digit of the first 12 digits: weights 1, 3, 1, 3 ... from the left."""
    total = sum(int(c) * (3 if i % 2 else 1) for i, c in enumerate(digits12))
    return (10 - total % 10) % 10


def prefixes(setting: Optional[str]) -> List[str]:
    """The prefixes of the settings: each of 1 to 3 digits; anything else is ignored."""
    out = []
    for p in (setting or "").split(","):
        p = p.strip()
        if 1 <= len(p) <= 3 and p.isdigit():
            out.append(p)
    return out


def code_parts(code: str, setting: Optional[str], item_digits: int) -> Optional[Tuple[str, int]]:
    """(item code, the value as an integer) when the code is a scale barcode of the settings, else None."""
    code = (code or "").strip()
    if len(code) != 13 or not code.isdigit() or not (4 <= item_digits <= 6):
        return None
    for p in prefixes(setting):
        if code.startswith(p) and len(p) + item_digits <= 11:
            if check_digit(code[:12]) != int(code[12]):
                return None
            item = code[len(p):len(p) + item_digits]
            return item, int(code[len(p) + item_digits:12])
    return None


def item_codes(item: str) -> List[str]:
    """The product codes that match an item code: as printed, then without its leading zeros."""
    stripped = item.lstrip("0")
    return [item] if stripped in ("", item) else [item, stripped]


def quantity(value: int, kind: str, price=None) -> Decimal:
    """The quantity of the line: the weight in kilograms, or the price / the unit price (3 decimals)."""
    if kind == "PRICE":
        amount = Decimal(value) / 100
        if price is None or Decimal(str(price)) <= 0:
            return Decimal(0)
        return (amount / Decimal(str(price))).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    return Decimal(value) / 1000


def make(prefix: str, item: str, value: int, value_digits: int) -> str:
    """A scale barcode with its check digit (for tests)."""
    body = prefix + item + str(value).zfill(value_digits)
    return body + str(check_digit(body))


# (code, prefixes, item digits, kind, price) -> (item, value, quantity) or None: shared with xUnit (gen_web)
VECTORS = [
    (make("20", "12345", 1250, 5), "20,21", 5, "WEIGHT", None),
    (make("21", "00042", 750, 5), "20, 21", 5, "WEIGHT", None),
    (make("2", "123456", 99999, 5), "2", 6, "WEIGHT", None),
    (make("20", "12345", 1250, 5), "20", 5, "PRICE", "25.00"),
    (make("20", "12345", 1000, 5), "20", 5, "PRICE", "3.00"),
    (make("20", "12345", 1250, 5), "20", 5, "PRICE", "0"),
    (make("22", "12345", 1250, 5), "20,21", 5, "WEIGHT", None),
    (make("20", "12345", 1250, 5)[:12] + "0", "20", 5, "WEIGHT", None),
    ("6281234567890", "20", 5, "WEIGHT", None),
    (make("20", "12345", 1250, 5), "", 5, "WEIGHT", None),
    (make("20", "1234", 12500, 6), "20,x,1234", 4, "WEIGHT", None),
]


def vector_results():
    out = []
    for code, setting, digits, kind, price in VECTORS:
        parts = code_parts(code, setting, digits)
        if parts is None:
            out.append((code, setting, digits, kind, price, None, None, None))
        else:
            item, value = parts
            out.append((code, setting, digits, kind, price, item, value, quantity(value, kind, price)))
    return out
