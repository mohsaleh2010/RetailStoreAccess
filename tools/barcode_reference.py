"""Reference barcode encoders (the VBA in modLabels must give the same modules).

barcode_modules(code) -> "1010..." one character per module (1 = bar), no quiet zone:
  - EAN-13 for 13 digits with a valid check digit (the usual retail barcode),
  - otherwise Code 128 (set C for digit pairs, set B for the rest) for printable ASCII,
  - "" when the text cannot be encoded (e.g. Arabic letters).
"""

EAN_L = ["0001101", "0011001", "0010011", "0111101", "0100011",
         "0110001", "0101111", "0111011", "0110111", "0001011"]
EAN_G = ["0100111", "0110011", "0011011", "0100001", "0011101",
         "0111001", "0000101", "0010001", "0001001", "0010111"]
EAN_R = ["1110010", "1100110", "1101100", "1000010", "1011100",
         "1001110", "1010000", "1000100", "1001000", "1110100"]
EAN_PARITY = ["LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG",
              "LGGLLG", "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL"]

# Code 128 bar/space widths for values 0..106 (106 = stop, 7 elements)
C128 = ("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 "
        "221312 231212 112232 122132 122231 113222 123122 123221 223211 221132 "
        "221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 "
        "212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 "
        "231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 "
        "231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 "
        "314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 "
        "112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 "
        "111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 "
        "214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 "
        "114131 311141 411131 211412 211214 211232 2331112").split()
START_B, START_C, CODE_B, STOP = 104, 105, 100, 106


def ean13_check_ok(code: str) -> bool:
    if len(code) != 13 or not code.isdigit():
        return False
    total = sum(int(d) * (3 if i % 2 else 1) for i, d in enumerate(code[:12]))
    return (10 - total % 10) % 10 == int(code[12])


def ean13_modules(code: str) -> str:
    parity = EAN_PARITY[int(code[0])]
    left = "".join((EAN_L if p == "L" else EAN_G)[int(d)] for p, d in zip(parity, code[1:7]))
    right = "".join(EAN_R[int(d)] for d in code[7:13])
    return "101" + left + "01010" + right + "101"


def widths_to_modules(widths: str) -> str:
    out, bar = [], True
    for w in widths:
        out.append(("1" if bar else "0") * int(w))
        bar = not bar
    return "".join(out)


def code128_values(text: str):
    if not text or any(not (32 <= ord(ch) <= 126) for ch in text):
        return []
    if text.isdigit() and len(text) >= 2:
        values = [START_C]
        even = len(text) - len(text) % 2
        values += [int(text[i:i + 2]) for i in range(0, even, 2)]
        if len(text) % 2:
            values += [CODE_B, ord(text[-1]) - 32]
    else:
        values = [START_B] + [ord(ch) - 32 for ch in text]
    check = values[0] + sum(i * v for i, v in enumerate(values[1:], 1))
    return values + [check % 103, STOP]


def code128_modules(text: str) -> str:
    return "".join(widths_to_modules(C128[v]) for v in code128_values(text))


def barcode_modules(code: str) -> str:
    code = (code or "").strip()
    if ean13_check_ok(code):
        return ean13_modules(code)
    return code128_modules(code)


def render_png(modules: str, path: str, scale: int = 3, height: int = 80, quiet: int = 12):
    """Draw the modules as an image (used by the tests to decode it with a real reader)."""
    from PIL import Image
    w = (len(modules) + 2 * quiet) * scale
    img = Image.new("L", (w, height), 255)
    px = img.load()
    for i, m in enumerate(modules):
        if m == "1":
            for x in range((quiet + i) * scale, (quiet + i + 1) * scale):
                for y in range(5, height - 5):
                    px[x, y] = 0
    img.save(path)
    return img
