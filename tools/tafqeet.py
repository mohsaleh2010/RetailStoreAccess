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


# Arabic unit, Arabic sub-unit, Arabic zero, English unit (one, many), English sub-unit (one, many):
# the program currency of the operating country (modCountry, docs/44)
CURRENCY_NAMES = {
    "SAR": ("ريال سعودي", "هللة", "صفر ريال", "Saudi Riyal", "Saudi Riyals", "Halala", "Halalas"),
    "EGP": ("جنيه مصري", "قرش", "صفر جنيه", "Egyptian Pound", "Egyptian Pounds", "Piaster", "Piasters"),
}


def amount_in_words(amount, currency="SAR") -> str:
    names = CURRENCY_NAMES.get(currency, CURRENCY_NAMES["SAR"])
    a = abs(Decimal(str(amount))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if a >= 1000000000:
        return f"{a:,.2f} {names[0]}"
    units = int(a)
    cents = int((a - units) * 100)
    if units == 0 and cents == 0:
        return names[2]
    text = ""
    if units:
        text = number_words(units) + " " + names[0]
    if cents:
        text += (" و" if text else "") + number_words(cents) + " " + names[1]
    return "فقط " + text + " لا غير"


CASES = [0.5, 1, 2, 3, 10, 11, 12, 19, 20, 21, 99, 100, 101, 115, 200, 999, 1000, 1001, 1250.5, 2000,
         2500, 3000, 10000, 11000, 12345.67, 100000, 101000, 999999.99, 1000000, 2000000, 3500000,
         11000000, 123456789.01, 0.05, 0.11, 4750, 3175]


# ---------------------------------------------------------------- English (modReports.AmountInWordsEn)
ONES_EN = ("zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen "
           "sixteen seventeen eighteen nineteen").split()
TENS_EN = "- - twenty thirty forty fifty sixty seventy eighty ninety".split()


def below_1000_en(n: int) -> str:
    parts = ""
    if n >= 100:
        parts = ONES_EN[n // 100] + " hundred"
    r = n % 100
    if r >= 20:
        parts = (parts + " " + TENS_EN[r // 10]).strip()
        if r % 10:
            parts += "-" + ONES_EN[r % 10]
    elif r:
        parts = (parts + " " + ONES_EN[r]).strip()
    return parts


def number_words_en(n: int) -> str:
    parts = ""
    if n >= 1000000:
        parts = below_1000_en(n // 1000000) + " million"
    if (n // 1000) % 1000:
        parts = (parts + " " + below_1000_en((n // 1000) % 1000) + " thousand").strip()
    if n % 1000:
        parts = (parts + " " + below_1000_en(n % 1000)).strip()
    return parts


def amount_in_words_en(amount, currency="SAR") -> str:
    names = CURRENCY_NAMES.get(currency, CURRENCY_NAMES["SAR"])
    a = Decimal(str(abs(amount))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if a >= 1000000000:
        return f"{a:,.2f} {names[4]}"
    units = int(a)
    cents = int((a - units) * 100)
    if units == 0 and cents == 0:
        return "Zero " + names[4]
    words = ""
    if units:
        words = number_words_en(units) + " " + (names[3] if units == 1 else names[4])
    if cents:
        if words:
            words += " and "
        words += number_words_en(cents) + " " + (names[5] if cents == 1 else names[6])
    return "Only " + words
