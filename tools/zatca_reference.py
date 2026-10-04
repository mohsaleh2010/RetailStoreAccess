"""Reference implementation of the ZATCA (Fatoora) phase-1 QR payload.

The QR code holds Base64( TLV(1, seller) TLV(2, vat no) TLV(3, timestamp)
TLV(4, total incl. VAT) TLV(5, VAT total) ), where TLV = tag byte, length
byte, UTF-8 value. modZatca.bas implements the same steps in VBA.
"""

import base64
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP

KSA_UTC_OFFSET_HOURS = 3      # Saudi Arabia: UTC+3, no daylight saving


def tlv(tag: int, value: str) -> bytes:
    raw = value.encode("utf-8")
    if len(raw) > 255:
        raise ValueError(f"TLV value too long for tag {tag}")
    return bytes([tag, len(raw)]) + raw


def amount(x) -> str:
    """Two decimals, '.' separator, half-up rounding, ASCII digits."""
    return str(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def timestamp(local: dt.datetime) -> str:
    """Invoice local time (KSA) -> ISO 8601 UTC with 'Z'."""
    utc = local - dt.timedelta(hours=KSA_UTC_OFFSET_HOURS)
    return utc.strftime("%Y-%m-%dT%H:%M:%SZ")


def qr_payload(seller: str, vat_number: str, iso_timestamp: str, total, vat) -> str:
    data = (tlv(1, seller) + tlv(2, vat_number) + tlv(3, iso_timestamp)
            + tlv(4, amount(total)) + tlv(5, amount(vat)))
    return base64.b64encode(data).decode("ascii")


def decode_payload(b64: str):
    raw = base64.b64decode(b64)
    out, i = {}, 0
    while i < len(raw):
        tag, length = raw[i], raw[i + 1]
        out[tag] = raw[i + 2:i + 2 + length].decode("utf-8")
        i += 2 + length
    return out
