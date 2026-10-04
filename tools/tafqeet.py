"""Reference for modReports.AmountInWords: an amount in Saudi riyals written in
Arabic words, in the form commonly used on vouchers and cheques:

    1250.50 -> "فقط ألف ومائتان وخمسون ريال سعودي وخمسون هللة لا غير"

Numbers use the standard masculine forms (واحد، اثنان، ثلاثة ... عشرون).
Thousands and millions follow the usual counting rules:
    1 ألف، 2 ألفان، 3-10 آلاف، 11 فما فوق ألف  (same for مليون / مليونان / ملايين).
"""

from decimal import Decimal, ROUND_HALF_UP

ONES = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة", "عشرة",
        "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر", "خمسة عشر", "ستة عشر", "سبعة عشر",
        "ثمانية عشر", "تسعة عشر"]
TENS = ["", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون", "ثمانون", "تسعون"]
HUNDREDS = ["", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة",
            "ثمانمائة", "تسعمائة"]


def below_1000(n: int) -> str:
    parts = []
    h, r = divmod(n, 100)
    if h:
        parts.append(HUNDREDS[h])
    if r:
        if r < 20:
            parts.append(ONES[r])
        else:
            t, u = divmod(r, 10)
            parts.append((ONES[u] + " و" if u else "") + TENS[t])
    return " و".join(parts)


def scale(n: int, one: str, two: str, few: str) -> str:
    if n == 1:
        return one
    if n == 2:
        return two
    if 3 <= n <= 10:
        return below_1000(n) + " " + few
    return below_1000(n) + " " + one


def number_words(n: int) -> str:
    if n == 0:
        return "صفر"
    millions, rest = divmod(n, 1_000_000)
    thousands, units = divmod(rest, 1000)
    parts = []
    if millions:
        parts.append(scale(millions, "مليون", "مليونان", "ملايين"))
    if thousands:
        parts.append(scale(thousands, "ألف", "ألفان", "آلاف"))
    if units:
        parts.append(below_1000(units))
    return " و".join(parts)


def amount_in_words(amount) -> str:
    a = abs(Decimal(str(amount))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    riyals = int(a)
    halalas = int((a - riyals) * 100)
    if riyals == 0 and halalas == 0:
        return "صفر ريال"
    text = ""
    if riyals:
        text = number_words(riyals) + " ريال سعودي"
    if halalas:
        text += (" و" if text else "") + number_words(halalas) + " هللة"
    return "فقط " + text + " لا غير"


CASES = [0.5, 1, 2, 3, 10, 11, 12, 19, 20, 21, 99, 100, 101, 115, 200, 999, 1000, 1001, 1250.5, 2000,
         2500, 3000, 10000, 11000, 12345.67, 100000, 101000, 999999.99, 1000000, 2000000, 3500000,
         11000000, 123456789.01, 0.05, 0.11, 4750, 3175]
