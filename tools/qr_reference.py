"""Reference QR Code encoder (ISO/IEC 18004), byte mode, error correction M,
versions 1-20. Written in a deliberately simple style: modQRCode.bas is a
line-by-line port of this file, and the tests compare both against the
`qrcode` library and decode the result with OpenCV.
"""

# Level M block structure per version:
# (ec codewords per block, blocks in group 1, data codewords per block g1, blocks g2, data cw g2)
M_BLOCKS = {
    1: (10, 1, 16, 0, 0), 2: (16, 1, 28, 0, 0), 3: (26, 1, 44, 0, 0), 4: (18, 2, 32, 0, 0),
    5: (24, 2, 43, 0, 0), 6: (16, 4, 27, 0, 0), 7: (18, 4, 31, 0, 0), 8: (22, 2, 38, 2, 39),
    9: (22, 3, 36, 2, 37), 10: (26, 4, 43, 1, 44), 11: (30, 1, 50, 4, 51), 12: (22, 6, 36, 2, 37),
    13: (22, 8, 37, 1, 38), 14: (24, 4, 40, 5, 41), 15: (24, 5, 41, 5, 42), 16: (28, 7, 45, 3, 46),
    17: (28, 10, 46, 1, 47), 18: (26, 9, 43, 4, 44), 19: (26, 3, 44, 11, 45), 20: (26, 3, 41, 13, 42),
}

ALIGNMENT = {
    1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30], 6: [6, 34], 7: [6, 22, 38],
    8: [6, 24, 42], 9: [6, 26, 46], 10: [6, 28, 50], 11: [6, 30, 54], 12: [6, 32, 58],
    13: [6, 34, 62], 14: [6, 26, 46, 66], 15: [6, 26, 48, 70], 16: [6, 26, 50, 74],
    17: [6, 30, 54, 78], 18: [6, 30, 56, 82], 19: [6, 30, 58, 86], 20: [6, 34, 62, 90],
}

MAX_VERSION = 20

# ---------------------------------------------------------------- GF(256)
EXP = [0] * 512
LOG = [0] * 256
_x = 1
for _i in range(255):
    EXP[_i] = _x
    LOG[_x] = _i
    _x <<= 1
    if _x & 0x100:
        _x ^= 0x11D
for _i in range(255, 512):
    EXP[_i] = EXP[_i - 255]


def gf_mul(a, b):
    if a == 0 or b == 0:
        return 0
    return EXP[LOG[a] + LOG[b]]


def rs_generator(n):
    """Coefficients of prod_{i<n} (x - a^i), highest degree first, leading 1."""
    g = [1]
    for i in range(n):
        nxt = [0] * (len(g) + 1)
        for j in range(len(g)):
            nxt[j] ^= g[j]
            nxt[j + 1] ^= gf_mul(g[j], EXP[i])
        g = nxt
    return g


def rs_remainder(data, n):
    gen = rs_generator(n)
    rem = [0] * n
    for d in data:
        factor = d ^ rem[0]
        rem = rem[1:] + [0]
        for j in range(n):
            rem[j] ^= gf_mul(gen[j + 1], factor)
    return rem


# ---------------------------------------------------------------- capacity
def data_codewords(version):
    ec, b1, d1, b2, d2 = M_BLOCKS[version]
    return b1 * d1 + b2 * d2


def count_bits(version):
    return 8 if version <= 9 else 16


def choose_version(byte_len):
    for v in range(1, MAX_VERSION + 1):
        if 4 + count_bits(v) + 8 * byte_len <= 8 * data_codewords(v):
            return v
    raise ValueError("text too long for a version 20-M QR code")


# ---------------------------------------------------------------- codewords
def make_codewords(data: bytes, version):
    bits = []

    def put(value, length):
        for i in range(length - 1, -1, -1):
            bits.append((value >> i) & 1)

    put(0b0100, 4)                       # byte mode
    put(len(data), count_bits(version))
    for b in data:
        put(b, 8)
    capacity = 8 * data_codewords(version)
    put(0, min(4, capacity - len(bits)))  # terminator
    while len(bits) % 8:
        bits.append(0)
    cw = [int("".join(map(str, bits[i:i + 8])), 2) for i in range(0, len(bits), 8)]
    pad = [0xEC, 0x11]
    k = 0
    while len(cw) < data_codewords(version):
        cw.append(pad[k % 2])
        k += 1

    ec_n, b1, d1, b2, d2 = M_BLOCKS[version]
    blocks, ecs, pos = [], [], 0
    for count, size in ((b1, d1), (b2, d2)):
        for _ in range(count):
            block = cw[pos:pos + size]
            pos += size
            blocks.append(block)
            ecs.append(rs_remainder(block, ec_n))
    out = []
    for i in range(max(d1, d2)):
        for block in blocks:
            if i < len(block):
                out.append(block[i])
    for i in range(ec_n):
        for e in ecs:
            out.append(e[i])
    return out


# ---------------------------------------------------------------- matrix
def bch_format(mask):
    data = (0b00 << 3) | mask            # level M = 00
    v = data << 10
    for i in range(14, 9, -1):
        if v & (1 << i):
            v ^= 0x537 << (i - 10)
    return ((data << 10) | v) ^ 0x5412


def bch_version(version):
    v = version << 12
    for i in range(17, 11, -1):
        if v & (1 << i):
            v ^= 0x1F25 << (i - 12)
    return (version << 12) | v


def mask_bit(mask, r, c):
    if mask == 0:
        return (r + c) % 2 == 0
    if mask == 1:
        return r % 2 == 0
    if mask == 2:
        return c % 3 == 0
    if mask == 3:
        return (r + c) % 3 == 0
    if mask == 4:
        return (r // 2 + c // 3) % 2 == 0
    if mask == 5:
        return (r * c) % 2 + (r * c) % 3 == 0
    if mask == 6:
        return ((r * c) % 2 + (r * c) % 3) % 2 == 0
    return ((r + c) % 2 + (r * c) % 3) % 2 == 0


def base_matrix(version):
    """Function patterns. Returns (modules, reserved) as size x size lists (0/1, bool)."""
    size = 17 + 4 * version
    mod = [[0] * size for _ in range(size)]
    res = [[False] * size for _ in range(size)]

    def setm(r, c, v):
        mod[r][c] = v
        res[r][c] = True

    for (r0, c0) in ((0, 0), (0, size - 7), (size - 7, 0)):
        for r in range(-1, 8):
            for c in range(-1, 8):
                rr, cc = r0 + r, c0 + c
                if 0 <= rr < size and 0 <= cc < size:
                    on = (0 <= r <= 6 and 0 <= c <= 6 and
                          (r in (0, 6) or c in (0, 6) or (2 <= r <= 4 and 2 <= c <= 4)))
                    setm(rr, cc, 1 if on else 0)
    for i in range(8, size - 8):
        setm(6, i, 1 if i % 2 == 0 else 0)
        setm(i, 6, 1 if i % 2 == 0 else 0)
    pos = ALIGNMENT[version]
    last = len(pos) - 1
    for i, r0 in enumerate(pos):
        for j, c0 in enumerate(pos):
            if (i == 0 and j == 0) or (i == 0 and j == last) or (i == last and j == 0):
                continue                     # the three corners hold finder patterns
            for r in range(-2, 3):
                for c in range(-2, 3):
                    on = max(abs(r), abs(c)) != 1
                    setm(r0 + r, c0 + c, 1 if on else 0)
    setm(size - 8, 8, 1)                     # dark module
    for i in range(9):                       # reserve format areas
        if not res[8][i]:
            res[8][i] = True
        if not res[i][8]:
            res[i][8] = True
    for i in range(8):
        res[8][size - 1 - i] = True
        res[size - 1 - i][8] = True
    if version >= 7:
        for i in range(6):
            for j in range(3):
                res[i][size - 11 + j] = True
                res[size - 11 + j][i] = True
    return mod, res


def place_data(mod, res, codewords):
    size = len(mod)
    bits = []
    for cw in codewords:
        for i in range(7, -1, -1):
            bits.append((cw >> i) & 1)
    idx = 0
    upward = True
    col = size - 1
    while col > 0:
        if col == 6:
            col -= 1
        rows = range(size - 1, -1, -1) if upward else range(size)
        for r in rows:
            for c in (col, col - 1):
                if not res[r][c]:
                    mod[r][c] = bits[idx] if idx < len(bits) else 0
                    idx += 1
        upward = not upward
        col -= 2


def apply_mask(mod, res, mask):
    size = len(mod)
    out = [row[:] for row in mod]
    for r in range(size):
        for c in range(size):
            if not res[r][c] and mask_bit(mask, r, c):
                out[r][c] ^= 1
    return out


def write_format(mod, mask, version):
    size = len(mod)
    f = bch_format(mask)
    for i in range(15):
        bit = (f >> i) & 1
        # vertical, next to the top-left finder / bottom-left finder
        if i < 6:
            mod[i][8] = bit
        elif i < 8:
            mod[i + 1][8] = bit
        else:
            mod[size - 15 + i][8] = bit
        # horizontal
        if i < 8:
            mod[8][size - 1 - i] = bit
        elif i < 9:
            mod[8][15 - i - 1 + 1] = bit
        else:
            mod[8][15 - i - 1] = bit
    mod[size - 8][8] = 1
    if version >= 7:
        v = bch_version(version)
        for i in range(18):
            bit = (v >> i) & 1
            mod[i // 3][size - 11 + i % 3] = bit
            mod[size - 11 + i % 3][i // 3] = bit


def penalty(mod):
    size = len(mod)
    score = 0
    # rule 1: runs of 5+ same colour
    for line in list(mod) + [list(col) for col in zip(*mod)]:
        run, prev = 0, -1
        for v in line:
            if v == prev:
                run += 1
            else:
                if run >= 5:
                    score += run - 2
                run, prev = 1, v
        if run >= 5:
            score += run - 2
    # rule 2: 2x2 blocks
    for r in range(size - 1):
        for c in range(size - 1):
            v = mod[r][c]
            if v == mod[r][c + 1] == mod[r + 1][c] == mod[r + 1][c + 1]:
                score += 3
    # rule 3: finder-like patterns 1011101 with 4 light modules on a side
    pat1 = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0]
    pat2 = [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1]
    for line in list(mod) + [list(col) for col in zip(*mod)]:
        for i in range(size - 10):
            seg = line[i:i + 11]
            if seg == pat1 or seg == pat2:
                score += 40
    # rule 4: dark proportion
    dark = sum(sum(row) for row in mod)
    k = abs(dark * 20 - size * size * 10) // (size * size)
    score += 10 * k
    return score


def encode(text_bytes: bytes, mask=None):
    """Returns (matrix, version, mask). matrix[r][c] = 1 for dark."""
    version = choose_version(len(text_bytes))
    codewords = make_codewords(text_bytes, version)
    base, res = base_matrix(version)
    place_data(base, res, codewords)
    best, best_mask, best_score = None, None, None
    for m in (range(8) if mask is None else [mask]):
        candidate = apply_mask(base, res, m)
        write_format(candidate, m, version)
        s = penalty(candidate)
        if best_score is None or s < best_score:
            best, best_mask, best_score = candidate, m, s
    return best, version, best_mask


def checksum(matrix) -> str:
    """Compact fingerprint used by the VBA tests: size + hex of all modules."""
    bits = "".join(str(v) for row in matrix for v in row)
    bits += "0" * (-len(bits) % 4)
    return f"{len(matrix)}:" + "".join(f"{int(bits[i:i + 4], 2):X}" for i in range(0, len(bits), 4))
