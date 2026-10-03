"""Helpers shared by the generators."""


def vba_str(s) -> str:
    """Return a VBA string literal."""
    if s is None:
        s = ""
    s = str(s).replace("\u2212", "-")   # Unicode minus is not in Windows-1256
    return '"' + s.replace('"', '""') + '"'
