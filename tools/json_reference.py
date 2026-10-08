"""Python mirror of modHttp.JsonGet and modHttp.JsonEscape (docs/45-EInvoice-Foundation.md): the same scan of
the JSON text, so that the VBA can be compared with it (and it with the json module) in tests/test_einvoice.py.

json_get(text, path): a string unescaped, a number / true / false as written, an object or array as its JSON
text, None when the path is missing or the value is null. Path "a.b", "items[0].code", "" = the whole value."""

WS = " \t\n\r"


def json_escape(text: str) -> str:
    out = []
    for ch in text:
        code = ord(ch)
        if ch == '"':
            out.append('\\"')
        elif ch == "\\":
            out.append("\\\\")
        elif ch == "\n":
            out.append("\\n")
        elif ch == "\r":
            out.append("\\r")
        elif ch == "\t":
            out.append("\\t")
        elif code < 32:
            out.append(f"\\u{code:04x}")
        else:
            out.append(ch)
    return "".join(out)


def _skip(s, pos):
    while pos < len(s) and s[pos] in WS:
        pos += 1
    return pos


def _end(s, start):
    pos = start
    ch = s[pos]
    if ch == '"':
        pos += 1
        while pos < len(s):
            if s[pos] == "\\":
                pos += 2
            elif s[pos] == '"':
                return pos + 1
            else:
                pos += 1
        return pos
    if ch in "{[":
        depth, in_string = 0, False
        while pos < len(s):
            ch = s[pos]
            if in_string:
                if ch == "\\":
                    pos += 1
                elif ch == '"':
                    in_string = False
            elif ch == '"':
                in_string = True
            elif ch in "{[":
                depth += 1
            elif ch in "}]":
                depth -= 1
                if depth == 0:
                    return pos + 1
            pos += 1
        return pos
    while pos < len(s) and s[pos] not in ",}]" + WS:
        pos += 1
    return pos


def _string_at(s, pos):
    out = []
    pos += 1
    while pos < len(s):
        ch = s[pos]
        if ch == '"':
            break
        if ch == "\\":
            pos += 1
            ch = s[pos]
            if ch == "u":
                out.append(chr(int(s[pos + 1:pos + 5], 16)))
                pos += 4
            else:
                out.append({"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f"}.get(ch, ch))
        else:
            out.append(ch)
        pos += 1
    # VBA strings are UTF-16: an escaped surrogate pair (\ud83d\ude00) is one character there too
    return "".join(out).encode("utf-16-le", "surrogatepass").decode("utf-16-le")


def _member_at(s, pos, name):
    if pos >= len(s) or s[pos] != "{":
        return None
    pos += 1
    while True:
        pos = _skip(s, pos)
        if pos >= len(s) or s[pos] != '"':
            return None
        key = _string_at(s, pos)
        pos = _skip(s, _end(s, pos))
        if pos >= len(s) or s[pos] != ":":
            return None
        pos = _skip(s, pos + 1)
        if key == name:
            return pos
        pos = _skip(s, _end(s, pos))
        if pos >= len(s) or s[pos] != ",":
            return None
        pos += 1


def _item_at(s, pos, index):
    if pos >= len(s) or s[pos] != "[":
        return None
    pos = _skip(s, pos + 1)
    if pos < len(s) and s[pos] == "]":
        return None
    i = 0
    while True:
        if i == index:
            return pos
        pos = _skip(s, _end(s, pos))
        if pos >= len(s) or s[pos] != ",":
            return None
        pos = _skip(s, pos + 1)
        i += 1


def json_get(text: str, path: str):
    pos = _skip(text, 0)
    if pos >= len(text):
        return None
    segs = path.replace("[", ".#").replace("]", "").replace("..", ".")
    if segs.startswith("."):
        segs = segs[1:]
    if segs:
        for seg in segs.split("."):
            pos = _item_at(text, pos, int(seg[1:])) if seg.startswith("#") else _member_at(text, pos, seg)
            if pos is None:
                return None
    if text[pos] == '"':
        return _string_at(text, pos)
    if text[pos:pos + 4] == "null":
        return None
    return text[pos:_end(text, pos)]
