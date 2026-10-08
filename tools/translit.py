"""Python mirror of modEnglishNames.Transliterate (docs/42-English-Names-Screen.md): an English name suggested
for an Arabic name, for the user to review in frmEnglishNames. Common words and Saudi first names / cities have
their usual English spelling (WORDS); any other word is written letter by letter (LETTERS).

The VBA module holds the same two tables as the constants TRANSLIT_LETTERS and TRANSLIT_WORDS
("ar=en|ar=en"); tests/test_english_names.py checks that both sides agree."""

LETTERS = {
    "ء": "", "آ": "a", "أ": "a", "ؤ": "o", "إ": "i", "ئ": "e", "ا": "a", "ب": "b", "ة": "a", "ت": "t",
    "ث": "th", "ج": "j", "ح": "h", "خ": "kh", "د": "d", "ذ": "th", "ر": "r", "ز": "z", "س": "s", "ش": "sh",
    "ص": "s", "ض": "d", "ط": "t", "ظ": "z", "ع": "a", "غ": "gh", "ف": "f", "ق": "q", "ك": "k", "ل": "l",
    "م": "m", "ن": "n", "ه": "h", "و": "o", "ى": "a", "ي": "i",
}
FIRST_LETTERS = {"و": "w", "ي": "y"}          # at the start of a word

WORDS = {
    # business words
    "شركة": "Company", "شركه": "Company", "مؤسسة": "Establishment", "مؤسسه": "Establishment", "مكتب": "Office",
    "مطعم": "Restaurant", "مقهى": "Cafe", "كافيه": "Cafe", "محل": "Shop", "محلات": "Shops", "متجر": "Store",
    "سوق": "Market", "أسواق": "Markets", "مصنع": "Factory", "مجموعة": "Group", "تجارة": "Trading",
    "للتجارة": "Trading", "التجارية": "Trading", "مقاولات": "Contracting", "للمقاولات": "Contracting",
    "وشركاه": "and Partners", "و": "and", "فرع": "Branch", "الرئيسي": "Main", "الرئيسية": "Main",
    "الرئيسيه": "Main", "صندوق": "Box", "خزينة": "Treasury", "الخزينة": "Treasury", "بنك": "Bank",
    "مصرف": "Bank", "حساب": "Account", "جاري": "Current", "توفير": "Savings", "قسم": "Department",
    "إدارة": "Administration", "المبيعات": "Sales", "المشتريات": "Purchases", "المستودع": "Warehouse",
    "مستودع": "Warehouse", "عميل": "Customer", "مورد": "Supplier", "نقدي": "Cash", "كاشير": "Cashier",
    "بن": "bin", "ابن": "bin", "بنت": "bint", "أبو": "Abu", "ابو": "Abu", "آل": "Al",
    # first names
    "محمد": "Mohammed", "أحمد": "Ahmed", "احمد": "Ahmed", "محمود": "Mahmoud", "علي": "Ali", "عمر": "Omar",
    "خالد": "Khalid", "سعد": "Saad", "سعيد": "Saeed", "فهد": "Fahad", "سلطان": "Sultan", "ناصر": "Nasser",
    "سعود": "Saud", "فيصل": "Faisal", "يوسف": "Yousef", "إبراهيم": "Ibrahim", "ابراهيم": "Ibrahim",
    "صالح": "Saleh", "حسن": "Hassan", "حسين": "Hussein", "سليمان": "Sulaiman", "منصور": "Mansour",
    "ماجد": "Majed", "تركي": "Turki", "بندر": "Bandar", "نايف": "Naif", "مشعل": "Mishal", "عادل": "Adel",
    "طارق": "Tariq", "ياسر": "Yasser", "هشام": "Hisham", "مصطفى": "Mustafa", "عبدالله": "Abdullah",
    "عبدالعزيز": "Abdulaziz", "عبدالرحمن": "Abdulrahman", "سارة": "Sarah", "ساره": "Sarah",
    "فاطمة": "Fatimah", "نورة": "Noura", "نوره": "Noura", "مريم": "Maryam", "عائشة": "Aisha", "هند": "Hind",
    # cities
    "الرياض": "Riyadh", "جدة": "Jeddah", "جده": "Jeddah", "مكة": "Makkah", "المدينة": "Madinah",
    "الدمام": "Dammam", "الخبر": "Khobar", "الطائف": "Taif", "تبوك": "Tabuk", "أبها": "Abha", "القصيم": "Qassim",
}

TASHKEEL = {chr(c) for c in range(0x064B, 0x0653)} | {"ـ"}
PUNCTUATION = {"،": ",", "؛": ",", "؟": "?"}


def cap(word: str) -> str:
    return word[:1].upper() + word[1:]


def letters(word: str) -> str:
    out = "".join(FIRST_LETTERS.get(ch, LETTERS[ch]) if i == 0 else LETTERS[ch] for i, ch in enumerate(word))
    while "aa" in out:
        out = out.replace("aa", "a")
    return out


def word_en(word: str) -> str:
    if word in WORDS:
        return WORDS[word]
    if len(word) > 3 and word.startswith("ال"):
        rest = word[2:]
        return "Al-" + (WORDS[rest] if rest in WORDS else cap(letters(rest)))
    if len(word) > 3 and word.startswith("عبد"):
        rest = word[3:]
        if len(rest) > 2 and rest.startswith("ال"):
            rest = rest[2:]
        return "Abdul" + letters(rest)
    return cap(letters(word))


def transliterate(text: str) -> str:
    out, word = [], ""
    for ch in text or "":
        if ch in TASHKEEL:
            continue
        if "٠" <= ch <= "٩":
            ch = chr(ord(ch) - 0x0660 + 0x30)
        ch = PUNCTUATION.get(ch, ch)
        if ch in LETTERS:
            word += ch
            continue
        if word:
            out.append(word_en(word))
            word = ""
        out.append(ch)
    if word:
        out.append(word_en(word))
    s = "".join(out)
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip(" ")


def vba_table(table: dict) -> str:
    """The constant of the VBA module: ar=en|ar=en."""
    return "|".join(f"{k}={v}" for k, v in table.items())
