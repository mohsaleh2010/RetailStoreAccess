"""Python mirror of modEtaInvoice (docs/49-ETA-EInvoice.md): the Egyptian e-invoice (B2B) of the Egyptian Tax
Authority (ETA), document version 1.0 (signed) or 0.9 (unsigned, pre-production only).

- The document is compact JSON with the numbers of eta_receipt_reference.num.
- Version 1.0 carries a CAdES-BES signature of the ETA serialization of the document without "signatures"
  (eta_receipt_reference.serialize), made by an external signing program that reads the USB token of the taxpayer.
- ETA gives the UUID and the long ID; the QR code is the public link of the document.
"""
import json

from eta_receipt_reference import PORTAL_URLS, esc, jnum, jstr, num  # noqa: F401  (num re-exported for tests)

DEFAULT_SIGNER_ARGS = '"{IN}" "{OUT}" "{PIN}"'


def party_json(role, party_type, party_id, party_name, branch_id, governate, city, street, building, postal) -> str:
    """The issuer (with its branch) or the receiver."""
    address = []
    if role == "issuer":
        address.append(jstr("branchID", branch_id))
    address += [jstr("country", "EG"), jstr("governate", governate), jstr("regionCity", city), jstr("street", street),
                jstr("buildingNumber", building)]
    if postal:
        address.append(jstr("postalCode", postal))
    parts = ['"address":{' + ",".join(address) + "}", jstr("type", party_type)]
    if party_id:
        parts.append(jstr("id", party_id))
    parts.append(jstr("name", party_name))
    return f'"{role}":{{' + ",".join(parts) + "}"


def line_json(it) -> str:
    """it = dict(description, item_type, item_code, unit_type, quantity, internal_code, unit_value, sales_total,
    discount, net_total, tax, rate, total)."""
    return "{" + ",".join([
        jstr("description", it["description"]), jstr("itemType", it["item_type"]), jstr("itemCode", it["item_code"]),
        jstr("unitType", it["unit_type"]), jnum("quantity", it["quantity"]), jstr("internalCode", it["internal_code"]),
        jnum("salesTotal", it["sales_total"]), jnum("total", it["total"]), jnum("valueDifference", 0),
        jnum("totalTaxableFees", 0), jnum("netTotal", it["net_total"]), jnum("itemsDiscount", 0),
        '"unitValue":{' + jstr("currencySold", "EGP") + "," + jnum("amountEGP", it["unit_value"]) + "}",
        '"discount":{' + jnum("rate", 0) + "," + jnum("amount", it["discount"]) + "}",
        '"taxableItems":[{' + jstr("taxType", "T1") + "," + jnum("amount", it["tax"]) + "," + jstr("subType", "V009") +
        "," + jnum("rate", it["rate"]) + "}]"]) + "}"


def invoice_json(d) -> str:
    """d = dict(doc_type 'I' / 'C', version, date_time, activity_code, internal_id, issuer, receiver (party_json
    texts), lines (line_json texts), reference_uuid ('' for an invoice), total_discount, total_sales, net_amount,
    tax_total, total_amount)."""
    parts = [d["issuer"], d["receiver"], jstr("documentType", d["doc_type"]), jstr("documentTypeVersion", d["version"]),
             jstr("dateTimeIssued", d["date_time"]), jstr("taxpayerActivityCode", d["activity_code"]),
             jstr("internalID", d["internal_id"])]
    if d["reference_uuid"]:
        parts.append('"references":["' + esc(d["reference_uuid"]) + '"]')
    parts += ['"invoiceLines":[' + ",".join(d["lines"]) + "]", jnum("totalDiscountAmount", d["total_discount"]),
              jnum("totalSalesAmount", d["total_sales"]), jnum("netAmount", d["net_amount"]),
              '"taxTotals":[{' + jstr("taxType", "T1") + "," + jnum("amount", d["tax_total"]) + "}]",
              jnum("totalAmount", d["total_amount"]), jnum("extraDiscountAmount", 0), jnum("totalItemsDiscountAmount", 0)]
    return "{" + ",".join(parts) + "}"


def with_signature(doc: str, signature: str) -> str:
    return doc[:-1] + ',"signatures":[{' + jstr("signatureType", "I") + "," + jstr("value", signature) + "}]}"


def signer_command(template, exe, in_file, out_file, pin) -> str:
    """The command line of the signing program: the template with {IN} {OUT} {PIN}, after the quoted program."""
    args = (template or DEFAULT_SIGNER_ARGS).replace("{IN}", in_file).replace("{OUT}", out_file).replace("{PIN}", pin)
    return f'"{exe}" {args}'


def qr_link(env, uuid, long_id) -> str:
    return f"{PORTAL_URLS[env]}/documents/{uuid}/share/{long_id}"


def _get(obj, *path):
    for p in path:
        if isinstance(p, int):
            obj = obj[p] if isinstance(obj, list) and len(obj) > p else None
        else:
            obj = obj.get(p) if isinstance(obj, dict) else None
        if obj is None:
            return None
    return obj


def _text(*values) -> str:
    for v in values:
        if v not in (None, ""):
            return str(v)
    return ""


def _errors(err, limit=3):
    if not isinstance(err, dict):
        return ""
    msg = _text(err.get("message"), err.get("code"))
    d = "; ".join(_text(x.get("propertyPath"), x.get("target"), x.get("code")) + ": " + _text(x.get("message"))
                  for x in (err.get("details") or [])[:limit])
    if msg and d:
        return f"{msg} ({d})"
    return msg or d


def submit_result(http_status, body):
    """(result, document status, message, submissionId, uuid, longId) of a document submission."""
    try:
        reply = json.loads(body) if body else {}
    except ValueError:
        reply = {}
    if not isinstance(reply, dict):
        reply = {}
    if http_status == 0:
        return "NETWORK", "", "تعذّر الاتصال بالمنظومة.", "", "", ""
    if http_status in (200, 202):
        if _get(reply, "rejectedDocuments", 0) is not None:
            return ("REJECTED", "REJECTED", _errors(_get(reply, "rejectedDocuments", 0, "error")) or "رُفض المستند.",
                    "", "", "")
        acc = _get(reply, "acceptedDocuments", 0)
        if acc is not None:
            return ("OK", "SUBMITTED", "", _text(reply.get("submissionId")), _text(_get(acc, "uuid")),
                    _text(_get(acc, "longId")))
        return "ERROR", "", f"رد غير متوقع من المصلحة: {body[:200]}", "", "", ""
    if http_status == 400:
        return "REJECTED", "REJECTED", _errors(reply.get("error")) or body[:200], "", "", ""
    if http_status in (401, 403):
        return "ERROR", "", "رفضت المصلحة بيانات دخول البرنامج (Client ID و Client Secret).", "", "", ""
    if http_status == 429 or http_status >= 500:
        return "NETWORK", "", f"المنظومة غير متاحة الآن (رمز {http_status}).", "", "", ""
    return "ERROR", "", f"رد غير متوقع من المصلحة (رمز {http_status}): {body[:200]}", "", "", ""


def status_result(body):
    """(document status, message) of the details of a document: VALID, INVALID, CANCELLED or "" (in progress)."""
    try:
        reply = json.loads(body) if body else {}
    except ValueError:
        return "", ""
    st = _text(reply.get("status") if isinstance(reply, dict) else None).lower()
    if st == "valid":
        return "VALID", ""
    if st in ("cancelled", "canceled"):
        return "CANCELLED", ""
    if st == "invalid":
        out = []
        for i in range(10):
            step = _get(reply, "validationResults", "validationSteps", i)
            if step is None or len(out) == 3:
                break
            if _text(step.get("status")).lower() == "invalid":
                out.append(_text(step.get("name")) + ": " +
                           _text(_get(step, "error", "innerError", 0, "error"), _get(step, "error", "error"),
                                 _get(step, "error", "errorCode")))
        return "INVALID", "; ".join(out)
    return "", ""
