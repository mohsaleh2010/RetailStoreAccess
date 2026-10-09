"""Python mirror of modZatcaApi (docs/47-ZATCA-Onboarding-Sending.md): the OpenSSL configuration of the
certificate request (CSR) of the device, the samples of the compliance checks, and the reading of ZATCA's
replies. The API itself (https://gw-fatoora.zatca.gov.sa/e-invoicing/...) follows the npm package zatca-sdk
0.1.2 (MIT License, Copyright (c) 2026 aashahin), whose requests pass ZATCA's sandbox."""
import json

BASE_URLS = {"TEST": "https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal",
             "SIMULATION": "https://gw-fatoora.zatca.gov.sa/e-invoicing/simulation",
             "PRODUCTION": "https://gw-fatoora.zatca.gov.sa/e-invoicing/core"}
TEMPLATE_NAMES = {"TEST": "TSTZATCA-Code-Signing", "SIMULATION": "PREZATCA-Code-Signing",
                  "PRODUCTION": "ZATCA-Code-Signing"}
SOLUTION = "RetailStoreAccess"
INVOICE_TYPES = "1100"          # standard (B2B) and simplified (B2C) invoices

# the compliance checks: every document type of "1100" (invoice, credit note, debit note; standard, simplified)
SAMPLES = [("388", "0100000"), ("381", "0100000"), ("383", "0100000"),
           ("388", "0200000"), ("381", "0200000"), ("383", "0200000")]


def config_value(s: str) -> str:
    """A value in the OpenSSL configuration: one line, no comment (#) nor variable ($)."""
    return " ".join(str(s).replace("#", " ").replace("$", " ").split())


def csr_config(env, serial, vat, location, industry, common_name, branch, organization) -> str:
    v = config_value
    return "\n".join([
        "[req]", "prompt = no", "utf8 = yes", "string_mask = utf8only", "distinguished_name = dn",
        "req_extensions = v3_req", "",
        "[dn]", "C = SA", f"OU = {v(branch)}", f"O = {v(organization)}", f"CN = {v(common_name)}", "",
        "[v3_req]", f"1.3.6.1.4.1.311.20.2 = ASN1:UTF8String:{TEMPLATE_NAMES[env]}",
        "subjectAltName = dirName:dir_sect", "",
        "[dir_sect]", f"SN = 1-{SOLUTION}|2-1.0|3-{v(serial)}", f"UID = {v(vat)}", f"title = {INVOICE_TYPES}",
        f"registeredAddress = {v(location)}", f"businessCategory = {v(industry)}", ""])


def messages(body, kind) -> str:
    """'code: message; ...' (at most 3) of validationResults.errorMessages / warningMessages."""
    try:
        items = json.loads(body).get("validationResults", {}).get(kind) or []
    except (ValueError, AttributeError):
        return ""
    return "; ".join(f"{m.get('code', '')}: {m.get('message', '')}" for m in items[:3])


def reply_result(http_status, body, standard):
    """(result, document status, message) of a reporting / clearance reply."""
    if http_status == 0:
        return "NETWORK", "", "تعذّر الاتصال بالمنظومة."
    if http_status == 200:
        return "OK", "CLEARED" if standard else "REPORTED", ""
    if http_status == 202:
        return "WARNING", "WARNING", messages(body, "warningMessages")
    if http_status == 400:
        return "REJECTED", "REJECTED", messages(body, "errorMessages") or body[:200]
    if http_status in (401, 403):
        return "ERROR", "", "رفضت الهيئة بيانات الدخول (شهادة الجهاز أو الكلمة السرية)."
    if http_status == 429 or http_status >= 500:
        return "NETWORK", "", f"المنظومة غير متاحة الآن (رمز {http_status})."
    return "ERROR", "", f"رد غير متوقع من الهيئة (رمز {http_status}): {body[:200]}"
