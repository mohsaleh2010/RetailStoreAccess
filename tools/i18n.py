"""The English interface (docs/38-English-Interface.md).

One Arabic -> English dictionary (tools/i18n_en.py) translates the interface:
  * when the screens, reports and queries are built in an English front-end (modBuildForms,
    modBuildReports, modBuildQueries call Tr on every caption, row source, record source, tag,
    expression and screen code), and
  * at run time (modLang.Tr in ShowMessage, AskYesNo, the captions and row sources set by code).

Tr works on fragments: a key is a run of Arabic text from its first to its last Arabic letter,
cut at the characters that never belong to a phrase (quotes, brackets, ; | & = < > and line
breaks). Tr replaces the longest keys first, and only where the key is not glued to another
Arabic letter, so a key never changes part of a longer Arabic word (a customer's name).
Arabic punctuation becomes Latin (، -> ,  ؛ -> ,  ؟ -> ?) and the quotes « » are dropped.

This module collects every Arabic string that reaches Tr (strings()), cuts it into keys
(fragments()) and mirrors modLang.Tr (translate()). tests/test_i18n.py checks that every key has
an English text and that nothing Arabic is left after translation.
"""

import os
import re
from typing import Dict, Iterable, List, Set

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

LETTERS = "ء-يٮ-ٯٱ-ۓۺ-ۼ"     # Arabic letters
MARKS = "ً-ٰٟـ"                                  # harakat, superscript alef, tatweel
_LETTER = re.compile(f"[{LETTERS}]")
_ANY = re.compile(f"[{LETTERS}{MARKS}]")
_CUT = re.compile(r"[\"'\[\]\;\|\&\r\n=<>]")
PUNCTUATION = {"\u060C": ",", "\u061B": ",", "\u061F": "?", "\u00AB": "", "\u00BB": ""}

# English texts must not break the strings they are put in (SQL literals, value lists, tags, code)
FORBIDDEN_IN_ENGLISH = set("\"'[];|&=<>\r\n")

# Reports kept bilingual in every interface: the tax invoice and the credit note must show Arabic (ZATCA)
ARABIC_REPORTS = {"rptSalesReceipt", "rptSalesInvoiceA4"}


def has_arabic(s: str) -> bool:
    return bool(_ANY.search(s or ""))


def fragments(s: str) -> List[str]:
    """The keys of a string: each piece between cut characters, from its first to its last Arabic letter
    (with the marks that follow that letter)."""
    out = []
    for piece in _CUT.split(s or ""):
        letters = [m.start() for m in _LETTER.finditer(piece)]
        if not letters:
            continue
        start, end = letters[0], letters[-1] + 1
        while end < len(piece) and re.match(f"[{MARKS}]", piece[end]):
            end += 1
        out.append(piece[start:end])
    return out


def _glued(s: str, i: int) -> bool:
    return 0 <= i < len(s) and bool(_ANY.match(s[i]))


def replace_word(s: str, key: str, value: str) -> str:
    """Replace key where it is not glued to other Arabic letters (modLang.ReplaceWord)."""
    out, pos = [], 0
    while True:
        i = s.find(key, pos)
        if i < 0:
            break
        if _glued(s, i - 1) or _glued(s, i + len(key)):
            out.append(s[pos:i + 1])
            pos = i + 1
            continue
        out.append(s[pos:i])
        out.append(value)
        pos = i + len(key)
    out.append(s[pos:])
    return "".join(out)


def ordered(dictionary: Dict[str, str]) -> List[str]:
    """The order in which modLang tries the keys: the longest first (then by text, to be stable)."""
    return sorted(dictionary, key=lambda k: (-len(k), k))


_NAME_MARK = re.compile(r"\[@(\w+)\]")


def resolve_names(s: str, english: bool = False) -> str:
    """[@Accounts] -> Accounts, or in English the saved query qryLocAccounts (modLang.LangSql)."""
    return _NAME_MARK.sub(lambda m: ("qryLoc" if english else "") + m.group(1), s)


def translate(s: str, dictionary: Dict[str, str], keys: List[str] = None) -> str:
    """Python mirror of modLang.Tr in an English front-end."""
    s = resolve_names(s, True)
    if not has_arabic(s):
        return s
    out = s
    for key in keys or ordered(dictionary):
        if key in out:
            out = replace_word(out, key, dictionary[key])
            if not has_arabic(out):
                break
    for ar, en in PUNCTUATION.items():
        out = out.replace(ar, en)
    return out


# ------------------------------------------------------------------ the strings that reach Tr
_VBA_LITERAL = re.compile(r'"((?:[^"]|"")*)"')


def vba_literals(line: str) -> List[str]:
    """String literals of one VBA line (the comment after ' is ignored)."""
    out, i = [], 0
    code = []
    in_str = False
    for ch in line:                                   # cut the comment
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        code.append(ch)
    for m in _VBA_LITERAL.finditer("".join(code)):
        out.append(m.group(1).replace('""', '"'))
    return out


# Constants of Arabic data that never reach Tr (the transliteration tables of modEnglishNames)
DATA_CONSTANTS = {"TRANSLIT_LETTERS", "TRANSLIT_WORDS"}


def module_strings(text: str, skip_tests: bool = True) -> List[str]:
    """Arabic literals of a hand-written module that the user can see: not in Debug.Print, not in the
    in-Access test procedures (Test*, Check*) that only the developer runs."""
    out, in_test, in_data = [], False, False
    for line in text.splitlines():
        head = re.match(r"\s*(?:Public |Private )?(?:Function|Sub) (\w+)", line)
        if head:
            in_test = skip_tests and head.group(1).startswith("Test")
        if re.match(r"\s*(?:Public |Private )?Const (\w+)", line):
            in_data = re.match(r"\s*(?:Public |Private )?Const (\w+)", line).group(1) in DATA_CONSTANTS
        skip = in_test or in_data or "Debug.Print" in line
        if in_data and not line.rstrip().endswith(" _"):
            in_data = False                          # the last line of a data constant
        if skip:
            continue
        out += [s for s in vba_literals(line) if has_arabic(s)]
    return out


def model_strings() -> List[str]:
    """The strings the builders pass through Tr: screens, reports, saved queries."""
    import forms as F
    import queries as Q
    import reports as RP
    out = []
    for m in F.all_forms():
        out += [m.caption, m.tag, m.record_source, "\n".join(m.code)]
        for c in m.controls:
            out += [v for v in c.props.values() if isinstance(v, str)]
    for r in RP.all_reports():
        if r.name in ARABIC_REPORTS:
            continue
        out += [r.caption, r.record_source, "\n".join(r.code)]
        for controls in r.controls.values():
            for c in controls:
                out += [v for v in c.props.values() if isinstance(v, str)]
                out += [c.source or ""]
    out += [q.sql for q in Q.QUERIES]
    return [s for s in out if has_arabic(s)]


def static_strings() -> List[str]:
    import generate as G
    out = []
    for name in G.STATIC_MODULES:
        if name in ("modTestAll",):
            continue
        with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
            out += module_strings(fh.read())
    with open(os.path.join(ROOT, "src", "vba", "modAppData.bas"), encoding="utf-8") as fh:
        out += module_strings(fh.read(), skip_tests=False)
    return out


def strings() -> List[str]:
    """Every Arabic string that reaches Tr, as Access sees it (after the Windows-1256 fallback)."""
    from generate_common import _CP1256_FALLBACK
    return [s.translate(_CP1256_FALLBACK) for s in model_strings() + static_strings()]


def keys(texts: Iterable[str] = None) -> Set[str]:
    out = set()
    for s in (strings() if texts is None else texts):
        out.update(fragments(s))
    return out
