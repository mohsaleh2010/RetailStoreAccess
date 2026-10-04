"""Helpers shared by the generators."""

# Characters that are not in Windows-1256 (the code page Access uses to import
# .bas files on Arabic Windows) and their plain replacements.
_CP1256_FALLBACK = str.maketrans({"−": "-", "≤": "<=", "≥": ">=",
                                  "≠": "<>", "→": "->", "…": "..."})


def vba_str(s) -> str:
    """Return a VBA string literal that survives the Windows-1256 import."""
    if s is None:
        s = ""
    s = str(s).translate(_CP1256_FALLBACK)
    s.encode("cp1256")  # raise now (with the text visible) instead of at write time
    return '"' + s.replace('"', '""') + '"'
