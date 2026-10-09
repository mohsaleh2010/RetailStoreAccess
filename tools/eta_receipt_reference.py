"""Python mirror of modEtaReceipt (docs/48-ETA-EReceipt.md): the Egyptian e-receipt (Egyptian Tax Authority, ETA)
of a point of sale, version 1.2.

- The receipt is a compact JSON object (no spaces), its numbers written with at most 5 decimals and no trailing zero.
- Its UUID is the hex SHA-256 of the ETA serialization of the receipt whose header.uuid is "": every property is
  "NAME" (upper case) followed by its value; a simple value is "value"; an array is "NAME" then, per element,
  "NAME" + the element.
- The QR code of the receipt is the link of the receipt on the ETA portal.
"""
import hashlib
import json
from decimal import Decimal, ROUND_HALF_EVEN

TOKEN_URLS = {"PREPROD": "https://id.preprod.eta.gov.eg/connect/token",
              "PROD": "https://id.eta.gov.eg/connect/token"}
API_URLS = {"PREPROD": "https://api.preprod.invoicing.eta.gov.eg",
            "PROD": "https://api.invoicing.eta.gov.eg"}
PORTAL_URLS = {"PREPROD": "https://preprod.invoicing.eta.gov.eg",
               "PROD": "https://invoicing.eta.gov.eg"}
TYPE_VERSION = "1.2"
BUYER_ID_LIMIT = 150000           # a person buying for this total or more gives the national ID


def eta_env(environment: str) -> str:
    """TEST and SIMULATION use the pre-production system."""
    return "PROD" if environment == "PRODUCTION" else "PREPROD"


def num(value) -> str:
    """A JSON number: rounded to 5 decimals (half even, as VBA Round), no trailing zero, no '.' alone."""
    d = Decimal(str(value)).quantize(Decimal("0.00001"), rounding=ROUND_HALF_EVEN)
    if d == 0:
        return "0"
    sign = "-" if d < 0 else ""
    d = abs(d)
    whole = int(d)
    frac = str(int((d - whole) * 100000)).rjust(5, "0").rstrip("0")
    return sign + str(whole) + ("." + frac if frac else "")


def esc(s) -> str:
    """The inside of a JSON string, as modHttp.JsonEscape writes it."""
    out = []
    for ch in str(s):
        c = ord(ch)
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
        elif c < 32:
            out.append("\\u%04x" % c)
        else:
            out.append(ch)
    return "".join(out)


def jstr(name, value) -> str:
    return f'"{name}":"{esc(value)}"'


def jnum(name, value) -> str:
    return f'"{name}":{num(value)}'


def item_json(it) -> str:
    """One line: it = dict(internal_code, description, item_type, item_code, unit_type, quantity, unit_price,
    total_sale, discount, net_sale, tax, rate, total)."""
    parts = [jstr("internalCode", it["internal_code"]), jstr("description", it["description"]),
             jstr("itemType", it["item_type"]), jstr("itemCode", it["item_code"]), jstr("unitType", it["unit_type"]),
             jnum("quantity", it["quantity"]), jnum("unitPrice", it["unit_price"]), jnum("netSale", it["net_sale"]),
             jnum("totalSale", it["total_sale"]), jnum("total", it["total"])]
    if Decimal(str(it["discount"])) != 0:
        parts.append('"commercialDiscountData":[{' + jnum("amount", it["discount"]) + "," +
                     jstr("description", "Discount") + "}]")
    parts.append('"taxableItems":[{' + jstr("taxType", "T1") + "," + jnum("amount", it["tax"]) + "," +
                 jstr("subType", "V009") + "," + jnum("rate", it["rate"]) + "}]")
    return "{" + ",".join(parts) + "}"


def receipt_json(r, uuid="") -> str:
    """r = dict(date_time, number, previous_uuid, reference_uuid ('' for a sale), rin, trade_name, branch_code,
    governate, city, street, building, postal, device_serial, activity_code, buyer_type, buyer_id, buyer_name,
    buyer_mobile, items, total_sales, total_discount, net_amount, tax_total, total_amount, payment_method)."""
    is_return = bool(r["reference_uuid"])
    header = [jstr("dateTimeIssued", r["date_time"]), jstr("receiptNumber", r["number"]), jstr("uuid", uuid),
              jstr("previousUUID", r["previous_uuid"])]
    if is_return:
        header.append(jstr("referenceUUID", r["reference_uuid"]))
    header += [jstr("currency", "EGP"), jnum("exchangeRate", 0)]
    address = [jstr("country", "EG"), jstr("governate", r["governate"]), jstr("regionCity", r["city"]),
               jstr("street", r["street"]), jstr("buildingNumber", r["building"])]
    if r["postal"]:
        address.append(jstr("postalCode", r["postal"]))
    seller = [jstr("rin", r["rin"]), jstr("companyTradeName", r["trade_name"]), jstr("branchCode", r["branch_code"]),
              '"branchAddress":{' + ",".join(address) + "}", jstr("deviceSerialNumber", r["device_serial"]),
              jstr("activityCode", r["activity_code"])]
    buyer = [jstr("type", r["buyer_type"])]
    if r["buyer_id"]:
        buyer.append(jstr("id", r["buyer_id"]))
    if r["buyer_name"]:
        buyer.append(jstr("name", r["buyer_name"]))
    if r["buyer_mobile"]:
        buyer.append(jstr("mobileNumber", r["buyer_mobile"]))
    parts = ['"header":{' + ",".join(header) + "}",
             '"documentType":{' + jstr("receiptType", "R" if is_return else "S") + "," +
             jstr("typeVersion", TYPE_VERSION) + "}",
             '"seller":{' + ",".join(seller) + "}", '"buyer":{' + ",".join(buyer) + "}",
             '"itemData":[' + ",".join(item_json(it) for it in r["items"]) + "]",
             jnum("totalSales", r["total_sales"]), jnum("totalCommercialDiscount", r["total_discount"]),
             jnum("totalItemsDiscount", 0), jnum("netAmount", r["net_amount"]), jnum("feesAmount", 0),
             jnum("totalAmount", r["total_amount"]),
             '"taxTotals":[{' + jstr("taxType", "T1") + "," + jnum("amount", r["tax_total"]) + "}]",
             jstr("paymentMethod", r["payment_method"])]
    return "{" + ",".join(parts) + "}"


def serialize(value) -> str:
    """The ETA serialization of a parsed receipt (objects keep their order; numbers keep their text)."""
    if isinstance(value, dict):
        out = []
        for name, v in value.items():
            key = '"' + name.upper() + '"'
            if isinstance(v, list):
                out.append(key)
                for el in v:
                    out.append(key + serialize(el))
            else:
                out.append(key + serialize(v))
        return "".join(out)
    if isinstance(value, bool):                       # true / false / null as written
        value = "true" if value else "false"
    elif value is None:
        value = "null"
    return '"' + str(value) + '"'


def parse(text: str):
    """JSON with the numbers kept as written (str), the order of the properties kept."""
    return json.loads(text, parse_float=str, parse_int=str)


def receipt_uuid(text: str) -> str:
    """The UUID of a receipt whose header.uuid is ""."""
    return hashlib.sha256(serialize(parse(text)).encode("utf-8")).hexdigest()


def signed_receipt(r):
    """(uuid, JSON with the uuid) of a receipt."""
    uuid = receipt_uuid(receipt_json(r))
    return uuid, receipt_json(r, uuid)


def qr_link(env, uuid, date_time, total_amount, rin) -> str:
    return f"{PORTAL_URLS[env]}/receipts/search/{uuid}/share/{date_time}#Total:{num(total_amount)},IssuerRIN:{rin}"


def payment_code(payment_type, method_id) -> str:
    """C cash, V visa (card), O others (bank transfer, sale on account)."""
    if payment_type == "CREDIT":
        return "O"
    return {2: "V", 3: "O"}.get(method_id or 1, "C")


def _first(*values) -> str:
    for v in values:
        if v not in (None, ""):
            return str(v)
    return ""


def _errors(err, limit=3):
    """'message (path: message; ...)' of an ETA error object {message, code, details: [...]}."""
    if not isinstance(err, dict):
        return ""
    msg = _first(err.get("message"), err.get("code"))
    d = "; ".join(_first(x.get("propertyPath"), x.get("target"), x.get("code")) + ": " + _first(x.get("message"))
                  for x in (err.get("details") or [])[:limit])
    if msg and d:
        return f"{msg} ({d})"
    return msg or d


def submit_result(http_status, body):
    """(result, document status, message, submissionId) of a receipt submission."""
    try:
        reply = json.loads(body) if body else {}
    except ValueError:
        reply = {}
    if not isinstance(reply, dict):
        reply = {}
    if http_status == 0:
        return "NETWORK", "", "تعذّر الاتصال بالمنظومة.", ""
    if http_status in (200, 202):
        rejected = reply.get("rejectedDocuments") or []
        if rejected:
            return "REJECTED", "REJECTED", _errors(rejected[0].get("error")) or "رُفض الإيصال.", ""
        if reply.get("acceptedDocuments"):
            return "OK", "SUBMITTED", "", str(reply.get("submissionId") or "")
        return "ERROR", "", f"رد غير متوقع من المصلحة: {body[:200]}", ""
    if http_status == 400:
        return "REJECTED", "REJECTED", _errors(reply.get("error")) or body[:200], ""
    if http_status in (401, 403):
        return "ERROR", "", "رفضت المصلحة بيانات الدخول (بيانات جهاز نقطة البيع).", ""
    if http_status == 429 or http_status >= 500:
        return "NETWORK", "", f"المنظومة غير متاحة الآن (رمز {http_status}).", ""
    return "ERROR", "", f"رد غير متوقع من المصلحة (رمز {http_status}): {body[:200]}", ""


def status_result(body):
    """(document status, message) of the details of a submission: VALID, INVALID or "" (still in progress)."""
    try:
        reply = json.loads(body) if body else {}
    except ValueError:
        return "", ""
    receipts = reply.get("receipts") or []
    status = (receipts[0].get("status") if receipts else None) or reply.get("status") or ""
    if status.lower() == "valid":
        return "VALID", ""
    if status.lower() == "invalid":
        errors = (receipts[0].get("errors") if receipts else None) or []
        return "INVALID", _errors({"details": errors if isinstance(errors, list) else []})
    return "", ""
