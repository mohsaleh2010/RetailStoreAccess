"""Python mirror of the ZATCA phase 2 document (modZatcaXml, docs/46-ZATCA-Invoice-XML.md): the UBL 2.1 invoice,
credit and debit note written directly in canonical form (C14N), the invoice hash, the XAdES signature block,
the certificate fields and the QR code with nine tags.

The signature templates below are the exact texts (spaces included) the ZATCA validator recomputes: the signed
properties are hashed as SIGNED_PROPERTIES_FOR_HASH (base64 of the hex SHA-256) while the document carries
SIGNED_PROPERTIES_FOR_XML at the same indentation. tools/gen_zatca.py writes them into modZatcaData.

The signature templates and the order of the UBL elements follow the npm package zatca-sdk 0.1.2, whose output
passes ZATCA's validator and sandbox (MIT License, Copyright (c) 2026 aashahin)."""
import base64
import hashlib
from decimal import Decimal, ROUND_HALF_UP

# the previous invoice hash of the first document (base64 of the hex SHA-256 of "0")
INITIAL_PIH = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="

NAMESPACES = ('xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" '
              'xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" '
              'xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2" '
              'xmlns:ext="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"')
QR_PLACEHOLDER = "SET_QR_CODE_DATA"

SIGNED_PROPERTIES_FOR_HASH = (
    '<xades:SignedProperties xmlns:xades="http://uri.etsi.org/01903/v1.3.2#" Id="xadesSignedProperties">\n'
    '                                    <xades:SignedSignatureProperties>\n'
    '                                        <xades:SigningTime>{SIGNING_TIME}</xades:SigningTime>\n'
    '                                        <xades:SigningCertificate>\n'
    '                                            <xades:Cert>\n'
    '                                                <xades:CertDigest>\n'
    '                                                    <ds:DigestMethod xmlns:ds="http://www.w3.org/2000/09/xmldsig#" Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>\n'
    '                                                    <ds:DigestValue xmlns:ds="http://www.w3.org/2000/09/xmldsig#">{CERT_HASH}</ds:DigestValue>\n'
    '                                                </xades:CertDigest>\n'
    '                                                <xades:IssuerSerial>\n'
    '                                                    <ds:X509IssuerName xmlns:ds="http://www.w3.org/2000/09/xmldsig#">{CERT_ISSUER}</ds:X509IssuerName>\n'
    '                                                    <ds:X509SerialNumber xmlns:ds="http://www.w3.org/2000/09/xmldsig#">{CERT_SERIAL}</ds:X509SerialNumber>\n'
    '                                                </xades:IssuerSerial>\n'
    '                                            </xades:Cert>\n'
    '                                        </xades:SigningCertificate>\n'
    '                                    </xades:SignedSignatureProperties>\n'
    '                                </xades:SignedProperties>'
)

SIGNED_PROPERTIES_FOR_XML = (
    '<xades:SignedProperties Id="xadesSignedProperties">\n'
    '                                    <xades:SignedSignatureProperties>\n'
    '                                        <xades:SigningTime>{SIGNING_TIME}</xades:SigningTime>\n'
    '                                        <xades:SigningCertificate>\n'
    '                                            <xades:Cert>\n'
    '                                                <xades:CertDigest>\n'
    '                                                    <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>\n'
    '                                                    <ds:DigestValue>{CERT_HASH}</ds:DigestValue>\n'
    '                                                </xades:CertDigest>\n'
    '                                                <xades:IssuerSerial>\n'
    '                                                    <ds:X509IssuerName>{CERT_ISSUER}</ds:X509IssuerName>\n'
    '                                                    <ds:X509SerialNumber>{CERT_SERIAL}</ds:X509SerialNumber>\n'
    '                                                </xades:IssuerSerial>\n'
    '                                            </xades:Cert>\n'
    '                                        </xades:SigningCertificate>\n'
    '                                    </xades:SignedSignatureProperties>\n'
    '                                </xades:SignedProperties>'
)

UBL_EXTENSIONS = (
    '<ext:UBLExtensions>\n'
    '    <ext:UBLExtension>\n'
    '        <ext:ExtensionURI>urn:oasis:names:specification:ubl:dsig:enveloped:xades</ext:ExtensionURI>\n'
    '        <ext:ExtensionContent>\n'
    '            <sig:UBLDocumentSignatures xmlns:sig="urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2" xmlns:sac="urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2" xmlns:sbc="urn:oasis:names:specification:ubl:schema:xsd:SignatureBasicComponents-2">\n'
    '                <sac:SignatureInformation>\n'
    '                    <cbc:ID>urn:oasis:names:specification:ubl:signature:1</cbc:ID>\n'
    '                    <sbc:ReferencedSignatureID>urn:oasis:names:specification:ubl:signature:Invoice</sbc:ReferencedSignatureID>\n'
    '                    <ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#" Id="signature">\n'
    '                        <ds:SignedInfo>\n'
    '                            <ds:CanonicalizationMethod Algorithm="http://www.w3.org/2006/12/xml-c14n11"/>\n'
    '                            <ds:SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"/>\n'
    '                            <ds:Reference Id="invoiceSignedData" URI="">\n'
    '                                <ds:Transforms>\n'
    '                                    <ds:Transform Algorithm="http://www.w3.org/TR/1999/REC-xpath-19991116">\n'
    '                                        <ds:XPath>not(//ancestor-or-self::ext:UBLExtensions)</ds:XPath>\n'
    '                                    </ds:Transform>\n'
    '                                    <ds:Transform Algorithm="http://www.w3.org/TR/1999/REC-xpath-19991116">\n'
    '                                        <ds:XPath>not(//ancestor-or-self::cac:Signature)</ds:XPath>\n'
    '                                    </ds:Transform>\n'
    '                                    <ds:Transform Algorithm="http://www.w3.org/TR/1999/REC-xpath-19991116">\n'
    "                                        <ds:XPath>not(//ancestor-or-self::cac:AdditionalDocumentReference[cbc:ID='QR'])</ds:XPath>\n"
    '                                    </ds:Transform>\n'
    '                                    <ds:Transform Algorithm="http://www.w3.org/2006/12/xml-c14n11"/>\n'
    '                                </ds:Transforms>\n'
    '                                <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>\n'
    '                                <ds:DigestValue>{INVOICE_HASH}</ds:DigestValue>\n'
    '                            </ds:Reference>\n'
    '                            <ds:Reference Type="http://www.w3.org/2000/09/xmldsig#SignatureProperties" URI="#xadesSignedProperties">\n'
    '                                <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>\n'
    '                                <ds:DigestValue>{PROPS_DIGEST}</ds:DigestValue>\n'
    '                            </ds:Reference>\n'
    '                        </ds:SignedInfo>\n'
    '                        <ds:SignatureValue>{SIGNATURE}</ds:SignatureValue>\n'
    '                        <ds:KeyInfo>\n'
    '                            <ds:X509Data>\n'
    '                                <ds:X509Certificate>{CERTIFICATE}</ds:X509Certificate>\n'
    '                            </ds:X509Data>\n'
    '                        </ds:KeyInfo>\n'
    '                        <ds:Object>\n'
    '                            <xades:QualifyingProperties xmlns:xades="http://uri.etsi.org/01903/v1.3.2#" Target="signature">\n'
    '                                {SIGNED_PROPERTIES}\n'
    '                            </xades:QualifyingProperties>\n'
    '                        </ds:Object>\n'
    '                    </ds:Signature>\n'
    '                </sac:SignatureInformation>\n'
    '            </sig:UBLDocumentSignatures>\n'
    '        </ext:ExtensionContent>\n'
    '    </ext:UBLExtension>\n'
    '</ext:UBLExtensions>'
)

DN_LABELS = {"2.5.4.3": "CN", "2.5.4.4": "SN", "2.5.4.5": "SERIALNUMBER", "2.5.4.6": "C", "2.5.4.7": "L",
             "2.5.4.8": "ST", "2.5.4.10": "O", "2.5.4.11": "OU", "2.5.4.12": "T", "0.9.2342.19200300.100.1.1": "UID",
             "0.9.2342.19200300.100.1.25": "DC", "1.2.840.113549.1.9.1": "E"}


# ------------------------------------------------------------------ values, written as C14N writes them
def amount(x) -> str:
    return str(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def price(x) -> str:
    """The unit price: 2 to 4 decimals (the stored net price has 4)."""
    s = str(Decimal(str(x)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))
    while s.endswith("0") and len(s.split(".")[1]) > 2:
        s = s[:-1]
    return s


def quantity(x) -> str:
    return str(Decimal(str(x)).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP))


def text(s) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def attr(s) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")


# ------------------------------------------------------------------ the document
def party(role, p, indent="  ") -> str:
    """Seller (role "Supplier") or buyer ("Customer"); no buyer = an empty element (simplified invoice)."""
    i = indent
    if not p:
        return f"{i}<cac:Accounting{role}Party></cac:Accounting{role}Party>"
    out = [f"{i}<cac:Accounting{role}Party>", f"{i}  <cac:Party>"]
    if p.get("crn"):
        out += [f"{i}    <cac:PartyIdentification>", f'{i}      <cbc:ID schemeID="CRN">{text(p["crn"])}</cbc:ID>',
                f"{i}    </cac:PartyIdentification>"]
    out += [f"{i}    <cac:PostalAddress>", f"{i}      <cbc:StreetName>{text(p['street'])}</cbc:StreetName>",
            f"{i}      <cbc:BuildingNumber>{text(p['building'])}</cbc:BuildingNumber>"]
    if p.get("plot"):
        out.append(f"{i}      <cbc:PlotIdentification>{text(p['plot'])}</cbc:PlotIdentification>")
    if p.get("district"):
        out.append(f"{i}      <cbc:CitySubdivisionName>{text(p['district'])}</cbc:CitySubdivisionName>")
    out += [f"{i}      <cbc:CityName>{text(p['city'])}</cbc:CityName>",
            f"{i}      <cbc:PostalZone>{text(p['postal'])}</cbc:PostalZone>", f"{i}      <cac:Country>",
            f"{i}        <cbc:IdentificationCode>{text(p.get('country') or 'SA')}</cbc:IdentificationCode>",
            f"{i}      </cac:Country>", f"{i}    </cac:PostalAddress>"]
    if p.get("vat"):
        out += [f"{i}    <cac:PartyTaxScheme>", f"{i}      <cbc:CompanyID>{text(p['vat'])}</cbc:CompanyID>",
                f"{i}      <cac:TaxScheme>", f"{i}        <cbc:ID>VAT</cbc:ID>", f"{i}      </cac:TaxScheme>",
                f"{i}    </cac:PartyTaxScheme>"]
    out += [f"{i}    <cac:PartyLegalEntity>", f"{i}      <cbc:RegistrationName>{text(p['name'])}</cbc:RegistrationName>",
            f"{i}    </cac:PartyLegalEntity>", f"{i}  </cac:Party>", f"{i}</cac:Accounting{role}Party>"]
    return "\n".join(out)


def line(ln, currency="SAR") -> str:
    c = attr(currency)
    out = ["  <cac:InvoiceLine>", f"    <cbc:ID>{text(ln['id'])}</cbc:ID>",
           f'    <cbc:InvoicedQuantity unitCode="{attr(ln["unit_code"])}">{quantity(ln["qty"])}</cbc:InvoicedQuantity>',
           f'    <cbc:LineExtensionAmount currencyID="{c}">{amount(ln["net"])}</cbc:LineExtensionAmount>']
    if Decimal(str(ln["discount"])) > 0:
        out += ["    <cac:AllowanceCharge>", "      <cbc:ChargeIndicator>false</cbc:ChargeIndicator>",
                "      <cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>",
                f'      <cbc:Amount currencyID="{c}">{amount(ln["discount"])}</cbc:Amount>', "    </cac:AllowanceCharge>"]
    out += ["    <cac:TaxTotal>", f'      <cbc:TaxAmount currencyID="{c}">{amount(ln["tax"])}</cbc:TaxAmount>',
            f'      <cbc:RoundingAmount currencyID="{c}">{amount(ln["total"])}</cbc:RoundingAmount>', "    </cac:TaxTotal>",
            "    <cac:Item>", f"      <cbc:Name>{text(ln['name'])}</cbc:Name>", "      <cac:ClassifiedTaxCategory>",
            f"        <cbc:ID>{text(ln['category'])}</cbc:ID>", f"        <cbc:Percent>{amount(ln['percent'])}</cbc:Percent>",
            "        <cac:TaxScheme>", "          <cbc:ID>VAT</cbc:ID>", "        </cac:TaxScheme>",
            "      </cac:ClassifiedTaxCategory>", "    </cac:Item>", "    <cac:Price>",
            f'      <cbc:PriceAmount currencyID="{c}">{price(ln["price"])}</cbc:PriceAmount>', "    </cac:Price>",
            "  </cac:InvoiceLine>"]
    return "\n".join(out)


def tax_subtotal(taxable, tax, category, percent, currency="SAR") -> str:
    c = attr(currency)
    return "\n".join([
        "    <cac:TaxSubtotal>", f'      <cbc:TaxableAmount currencyID="{c}">{amount(taxable)}</cbc:TaxableAmount>',
        f'      <cbc:TaxAmount currencyID="{c}">{amount(tax)}</cbc:TaxAmount>', "      <cac:TaxCategory>",
        f'        <cbc:ID schemeAgencyID="6" schemeID="UN/ECE 5305">{text(category)}</cbc:ID>',
        f"        <cbc:Percent>{amount(percent)}</cbc:Percent>", "        <cac:TaxScheme>",
        '          <cbc:ID schemeAgencyID="6" schemeID="UN/ECE 5153">VAT</cbc:ID>', "        </cac:TaxScheme>",
        "      </cac:TaxCategory>", "    </cac:TaxSubtotal>"])


def invoice_xml(d) -> str:
    """The document without the XML declaration and the signature block, with the QR placeholder: already in
    canonical form (namespaces and attributes in C14N order, no empty-element tags, text escaped like C14N)."""
    c = attr(d.get("currency", "SAR"))
    out = [f"<Invoice {NAMESPACES}>", "  <cbc:ProfileID>reporting:1.0</cbc:ProfileID>", f"  <cbc:ID>{text(d['id'])}</cbc:ID>",
           f"  <cbc:UUID>{text(d['uuid'])}</cbc:UUID>", f"  <cbc:IssueDate>{d['issue_date']}</cbc:IssueDate>",
           f"  <cbc:IssueTime>{d['issue_time']}</cbc:IssueTime>",
           f'  <cbc:InvoiceTypeCode name="{attr(d["sub_type"])}">{text(d["type_code"])}</cbc:InvoiceTypeCode>',
           f"  <cbc:DocumentCurrencyCode>{c}</cbc:DocumentCurrencyCode>", f"  <cbc:TaxCurrencyCode>{c}</cbc:TaxCurrencyCode>"]
    if d.get("billing_ref"):
        out += ["  <cac:BillingReference>", "    <cac:InvoiceDocumentReference>",
                f"      <cbc:ID>{text(d['billing_ref'])}</cbc:ID>", "    </cac:InvoiceDocumentReference>",
                "  </cac:BillingReference>"]
    out += ["  <cac:AdditionalDocumentReference>", "    <cbc:ID>ICV</cbc:ID>", f"    <cbc:UUID>{d['icv']}</cbc:UUID>",
            "  </cac:AdditionalDocumentReference>", "  <cac:AdditionalDocumentReference>", "    <cbc:ID>PIH</cbc:ID>",
            "    <cac:Attachment>",
            f'      <cbc:EmbeddedDocumentBinaryObject mimeCode="text/plain">{text(d["pih"])}</cbc:EmbeddedDocumentBinaryObject>',
            "    </cac:Attachment>", "  </cac:AdditionalDocumentReference>", "  <cac:AdditionalDocumentReference>",
            "    <cbc:ID>QR</cbc:ID>", "    <cac:Attachment>",
            f'      <cbc:EmbeddedDocumentBinaryObject mimeCode="text/plain">{QR_PLACEHOLDER}</cbc:EmbeddedDocumentBinaryObject>',
            "    </cac:Attachment>", "  </cac:AdditionalDocumentReference>", "  <cac:Signature>",
            "    <cbc:ID>urn:oasis:names:specification:ubl:signature:Invoice</cbc:ID>",
            "    <cbc:SignatureMethod>urn:oasis:names:specification:ubl:dsig:enveloped:xades</cbc:SignatureMethod>",
            "  </cac:Signature>", party("Supplier", d["seller"]), party("Customer", d.get("buyer"))]
    if d.get("delivery_date"):
        out += ["  <cac:Delivery>", f"    <cbc:ActualDeliveryDate>{d['delivery_date']}</cbc:ActualDeliveryDate>",
                "  </cac:Delivery>"]
    out += ["  <cac:PaymentMeans>", f"    <cbc:PaymentMeansCode>{text(d['payment_code'])}</cbc:PaymentMeansCode>"]
    if d.get("reason"):
        out.append(f"    <cbc:InstructionNote>{text(d['reason'])}</cbc:InstructionNote>")
    out += ["  </cac:PaymentMeans>", "  <cac:TaxTotal>",
            f'    <cbc:TaxAmount currencyID="{c}">{amount(d["tax_total"])}</cbc:TaxAmount>', "  </cac:TaxTotal>",
            "  <cac:TaxTotal>", f'    <cbc:TaxAmount currencyID="{c}">{amount(d["tax_total"])}</cbc:TaxAmount>']
    out += [tax_subtotal(*s, currency=d.get("currency", "SAR")) for s in d["subtotals"]]
    out += ["  </cac:TaxTotal>", "  <cac:LegalMonetaryTotal>",
            f'    <cbc:LineExtensionAmount currencyID="{c}">{amount(d["line_ext"])}</cbc:LineExtensionAmount>',
            f'    <cbc:TaxExclusiveAmount currencyID="{c}">{amount(d["tax_excl"])}</cbc:TaxExclusiveAmount>',
            f'    <cbc:TaxInclusiveAmount currencyID="{c}">{amount(d["tax_incl"])}</cbc:TaxInclusiveAmount>',
            f'    <cbc:AllowanceTotalAmount currencyID="{c}">0.00</cbc:AllowanceTotalAmount>',
            f'    <cbc:PrepaidAmount currencyID="{c}">0.00</cbc:PrepaidAmount>',
            f'    <cbc:PayableAmount currencyID="{c}">{amount(d["tax_incl"])}</cbc:PayableAmount>',
            "  </cac:LegalMonetaryTotal>"]
    out += [line(ln, d.get("currency", "SAR")) for ln in d["lines"]]
    out.append("</Invoice>")
    return "\n".join(out)


def signed_document(xml, extensions, qr) -> str:
    """The file: declaration, the signature block right after the root tag, the QR code in place."""
    start = xml.index(">") + 1
    return ('<?xml version="1.0" encoding="UTF-8"?>\n' + xml[:start] + extensions + xml[start:]).replace(QR_PLACEHOLDER, qr)


# ------------------------------------------------------------------ hash
def _remove(xml, start_tag, end_tag, after=0):
    i = xml.index(start_tag, after)
    j = xml.index(end_tag, i) + len(end_tag)
    return xml[:i] + xml[j:]


def hash_input(xml) -> str:
    """What ZATCA hashes: without the declaration, ext:UBLExtensions, cac:Signature and the QR reference (their
    surrounding white space stays), canonical. The document is written canonical, so cutting the text is enough."""
    s = xml[xml.index("<Invoice"):]
    if "<ext:UBLExtensions>" in s:
        s = _remove(s, "<ext:UBLExtensions>", "</ext:UBLExtensions>")
    s = _remove(s, "<cac:Signature>", "</cac:Signature>")
    qr = s.index("<cbc:ID>QR</cbc:ID>")
    s = _remove(s, "<cac:AdditionalDocumentReference>", "</cac:AdditionalDocumentReference>",
                s.rindex("<cac:AdditionalDocumentReference>", 0, qr))
    return s


def invoice_hash(xml) -> str:
    return base64.b64encode(hashlib.sha256(hash_input(xml).encode("utf-8")).digest()).decode()


def hex_digest_b64(data: bytes) -> str:
    """ZATCA's digest of the certificate and of the signed properties: base64 of the hex SHA-256."""
    return base64.b64encode(hashlib.sha256(data).hexdigest().encode()).decode()


# ------------------------------------------------------------------ certificate (DER)
def _read(buf, pos):
    """(tag, start, content start, end) of the element at pos."""
    tag, length, cursor = buf[pos], buf[pos + 1], pos + 2
    if length & 0x80:
        n = length & 0x7F
        length = int.from_bytes(buf[cursor:cursor + n], "big")
        cursor += n
    return tag, pos, cursor, cursor + length


def _children(buf, node):
    out, cursor = [], node[2]
    while cursor < node[3]:
        child = _read(buf, cursor)
        out.append(child)
        cursor = child[3]
    return out


def _content(buf, node) -> bytes:
    return buf[node[2]:node[3]]


def _oid(content: bytes) -> str:
    arcs, value = [], 0
    for b in content:
        value = (value << 7) | (b & 0x7F)
        if not b & 0x80:
            arcs.append(value)
            value = 0
    first = min(2, arcs[0] // 40)
    return ".".join(str(a) for a in [first, arcs[0] - first * 40] + arcs[1:])


def _name(buf, node) -> str:
    parts = []
    for rdn in _children(buf, node):
        oid_node, value_node = _children(buf, _children(buf, rdn)[0])[:2]
        oid = _oid(_content(buf, oid_node))
        parts.append(f"{DN_LABELS.get(oid, oid)}={_content(buf, value_node).decode('utf-8')}")
    return ", ".join(reversed(parts))


def certificate_info(cert_b64: str) -> dict:
    """Issuer (CN=..., DC=... in reverse order), decimal serial number, public key (the SubjectPublicKeyInfo
    DER, QR tag 8), signature of the certificate (QR tag 9) and ZATCA's certificate digest."""
    der = base64.b64decode(cert_b64)
    tbs, _, sig = _children(der, _read(der, 0))
    items = _children(der, tbs)
    off = 1 if items[0][0] == 0xA0 else 0
    serial, issuer, spki = items[off], items[off + 2], items[off + 5]
    return {"issuer": _name(der, issuer),
            "serial": str(int.from_bytes(_content(der, serial), "big")),
            "public_key": der[spki[1]:spki[3]],
            "signature": _content(der, sig)[1:],
            "hash": hex_digest_b64(cert_b64.encode("ascii"))}


# ------------------------------------------------------------------ signature block and QR
def signed_properties(signing_time, cert, for_hash) -> str:
    t = SIGNED_PROPERTIES_FOR_HASH if for_hash else SIGNED_PROPERTIES_FOR_XML
    return (t.replace("{SIGNING_TIME}", signing_time).replace("{CERT_HASH}", cert["hash"])
            .replace("{CERT_ISSUER}", text(cert["issuer"])).replace("{CERT_SERIAL}", cert["serial"]))


def ubl_extensions(inv_hash, signature, cert_b64, signing_time, cert) -> str:
    digest = hex_digest_b64(signed_properties(signing_time, cert, True).encode("utf-8"))
    return (UBL_EXTENSIONS.replace("{INVOICE_HASH}", inv_hash).replace("{PROPS_DIGEST}", digest)
            .replace("{SIGNATURE}", signature).replace("{CERTIFICATE}", cert_b64)
            .replace("{SIGNED_PROPERTIES}", signed_properties(signing_time, cert, False)))


def _tlv(tag, value: bytes) -> bytes:
    n = len(value)
    length = bytes([n]) if n < 0x80 else bytes([0x81, n]) if n < 0x100 else bytes([0x82, n >> 8, n & 0xFF])
    return bytes([tag]) + length + value


def qr_code(seller, vat, timestamp, total, vat_total, inv_hash, signature, public_key: bytes, cert_signature: bytes) -> str:
    """Nine tags: 1-5 as phase 1, 6 invoice hash, 7 signature (both base64 text), 8 public key, 9 signature of
    the certificate (both bytes)."""
    values = [seller.encode(), vat.encode(), timestamp.encode(), amount(total).encode(), amount(vat_total).encode(),
              inv_hash.encode(), signature.encode(), public_key, cert_signature]
    return base64.b64encode(b"".join(_tlv(i + 1, v) for i, v in enumerate(values))).decode()
