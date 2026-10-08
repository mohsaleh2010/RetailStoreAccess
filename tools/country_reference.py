"""Python mirror of modCountry (docs/44-Operating-Country.md): what the operating country decides
(Settings.CountryCode: SA = Saudi Arabia, EG = Egypt)."""
import re

VAT_RATE = {"SA": 0.15, "EG": 0.14}
CURRENCY = {"SA": "SAR", "EG": "EGP"}


def country(code) -> str:
    return "EG" if (code or "").upper() == "EG" else "SA"


def tax_number_problem(value, code) -> str:
    s = (value or "").strip(" ")
    if not s:
        return ""
    if country(code) == "EG":
        return "" if re.fullmatch(r"\d{9}", s) else "رقم التسجيل الضريبي في مصر 9 أرقام."
    return "" if re.fullmatch(r"3\d{13}3", s) else "الرقم الضريبي في السعودية 15 رقمًا ويبدأ وينتهي بالرقم 3."


def mobile_fits(mobile, code) -> bool:
    if country(code) == "EG":
        return bool(re.fullmatch(r"01\d{9}|\+201\d{9}|201\d{9}", mobile))
    return bool(re.fullmatch(r"05\d{8}|\+9665\d{8}|9665\d{8}", mobile))


def doc_title(code, kind, sub_type, english) -> str:
    if kind == "RETURN":
        i = 0
    elif sub_type == "STANDARD":
        i = 1
    elif country(code) == "EG":
        i = 3
    else:
        i = 2
    titles = (["Credit Note", "Tax Invoice", "Simplified Tax Invoice", "Sales Receipt"] if english else
              ["إشعار دائن", "فاتورة ضريبية", "فاتورة ضريبية مبسطة", "إيصال بيع"])
    return titles[i]
