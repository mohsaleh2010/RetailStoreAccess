using System.Globalization;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;

namespace RetailStore.Domain.EInvoicing;

/// <summary>A seller ("Supplier") or buyer ("Customer") of a ZATCA document; the buyer of a simplified invoice is null.</summary>
public sealed record ZatcaParty(string Name, string Vat, string Crn, string Street, string Building, string Plot,
                                string District, string City, string Postal, string Country = "SA");

/// <summary>One line: Percent is the rate in percent (15), as the document writes it.</summary>
public sealed record ZatcaLine(int Id, decimal Qty, string UnitCode, decimal Net, decimal Discount, decimal Tax, decimal Total,
                               string Name, string Category, decimal Percent, decimal Price);

public sealed record ZatcaTaxSubtotal(decimal Taxable, decimal Tax, string Category, decimal Percent);

/// <summary>The data of modZatcaXml.ZxInvoice / tools/zatca_xml_reference.invoice_xml.</summary>
public sealed record ZatcaInvoice(string Id, string Uuid, string IssueDate, string IssueTime, string TypeCode, string SubType,
                                  string Currency, string BillingRef, int Icv, string Pih, ZatcaParty Seller, ZatcaParty? Buyer,
                                  string DeliveryDate, string PaymentCode, string Reason, decimal TaxTotal,
                                  IReadOnlyList<ZatcaTaxSubtotal> Subtotals, decimal LineExt, decimal TaxExcl, decimal TaxIncl,
                                  IReadOnlyList<ZatcaLine> Lines);

/// <summary>What the certificate gives the signature and the QR code (modZatcaXml.ZatcaCertInfo).</summary>
public sealed record ZatcaCertificate(string Issuer, string Serial, byte[] PublicKey, byte[] Signature, string Hash);

/// <summary>The ZATCA phase 2 document in C# (docs/46, docs/85): modZatcaXml and tools/zatca_xml_reference.py, written
/// directly in canonical form (C14N), its hash, the certificate fields, the XAdES block and the QR code with nine tags.
/// The texts are those of the Python mirror, spaces included: ZATCA's validator recomputes them.</summary>
public static class ZatcaDocument
{
    public const string InitialPih = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ==";
    public const string QrPlaceholder = "SET_QR_CODE_DATA";
    public const int KsaUtcOffsetHours = 3;

    private const string Namespaces =
        "xmlns=\"urn:oasis:names:specification:ubl:schema:xsd:Invoice-2\" " +
        "xmlns:cac=\"urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2\" " +
        "xmlns:cbc=\"urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2\" " +
        "xmlns:ext=\"urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2\"";

    private static string Props(bool forHash)
    {
        string ds = forHash ? " xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\"" : "";
        string xades = forHash ? " xmlns:xades=\"http://uri.etsi.org/01903/v1.3.2#\"" : "";
        string i(int n) => "\n" + new string(' ', n);
        return "<xades:SignedProperties" + xades + " Id=\"xadesSignedProperties\">" +
               i(36) + "<xades:SignedSignatureProperties>" +
               i(40) + "<xades:SigningTime>{SIGNING_TIME}</xades:SigningTime>" +
               i(40) + "<xades:SigningCertificate>" +
               i(44) + "<xades:Cert>" +
               i(48) + "<xades:CertDigest>" +
               i(52) + "<ds:DigestMethod" + ds + " Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\"/>" +
               i(52) + "<ds:DigestValue" + ds + ">{CERT_HASH}</ds:DigestValue>" +
               i(48) + "</xades:CertDigest>" +
               i(48) + "<xades:IssuerSerial>" +
               i(52) + "<ds:X509IssuerName" + ds + ">{CERT_ISSUER}</ds:X509IssuerName>" +
               i(52) + "<ds:X509SerialNumber" + ds + ">{CERT_SERIAL}</ds:X509SerialNumber>" +
               i(48) + "</xades:IssuerSerial>" +
               i(44) + "</xades:Cert>" +
               i(40) + "</xades:SigningCertificate>" +
               i(36) + "</xades:SignedSignatureProperties>" +
               "\n" + new string(' ', 32) + "</xades:SignedProperties>";
    }

    public static readonly string SignedPropertiesForHash = Props(true);
    public static readonly string SignedPropertiesForXml = Props(false);

    public static readonly string UblExtensionsTemplate = string.Join("\n",
        "<ext:UBLExtensions>",
        "    <ext:UBLExtension>",
        "        <ext:ExtensionURI>urn:oasis:names:specification:ubl:dsig:enveloped:xades</ext:ExtensionURI>",
        "        <ext:ExtensionContent>",
        "            <sig:UBLDocumentSignatures xmlns:sig=\"urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2\" xmlns:sac=\"urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2\" xmlns:sbc=\"urn:oasis:names:specification:ubl:schema:xsd:SignatureBasicComponents-2\">",
        "                <sac:SignatureInformation>",
        "                    <cbc:ID>urn:oasis:names:specification:ubl:signature:1</cbc:ID>",
        "                    <sbc:ReferencedSignatureID>urn:oasis:names:specification:ubl:signature:Invoice</sbc:ReferencedSignatureID>",
        "                    <ds:Signature xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\" Id=\"signature\">",
        "                        <ds:SignedInfo>",
        "                            <ds:CanonicalizationMethod Algorithm=\"http://www.w3.org/2006/12/xml-c14n11\"/>",
        "                            <ds:SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256\"/>",
        "                            <ds:Reference Id=\"invoiceSignedData\" URI=\"\">",
        "                                <ds:Transforms>",
        "                                    <ds:Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\">",
        "                                        <ds:XPath>not(//ancestor-or-self::ext:UBLExtensions)</ds:XPath>",
        "                                    </ds:Transform>",
        "                                    <ds:Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\">",
        "                                        <ds:XPath>not(//ancestor-or-self::cac:Signature)</ds:XPath>",
        "                                    </ds:Transform>",
        "                                    <ds:Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\">",
        "                                        <ds:XPath>not(//ancestor-or-self::cac:AdditionalDocumentReference[cbc:ID='QR'])</ds:XPath>",
        "                                    </ds:Transform>",
        "                                    <ds:Transform Algorithm=\"http://www.w3.org/2006/12/xml-c14n11\"/>",
        "                                </ds:Transforms>",
        "                                <ds:DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\"/>",
        "                                <ds:DigestValue>{INVOICE_HASH}</ds:DigestValue>",
        "                            </ds:Reference>",
        "                            <ds:Reference Type=\"http://www.w3.org/2000/09/xmldsig#SignatureProperties\" URI=\"#xadesSignedProperties\">",
        "                                <ds:DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\"/>",
        "                                <ds:DigestValue>{PROPS_DIGEST}</ds:DigestValue>",
        "                            </ds:Reference>",
        "                        </ds:SignedInfo>",
        "                        <ds:SignatureValue>{SIGNATURE}</ds:SignatureValue>",
        "                        <ds:KeyInfo>",
        "                            <ds:X509Data>",
        "                                <ds:X509Certificate>{CERTIFICATE}</ds:X509Certificate>",
        "                            </ds:X509Data>",
        "                        </ds:KeyInfo>",
        "                        <ds:Object>",
        "                            <xades:QualifyingProperties xmlns:xades=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"signature\">",
        "                                {SIGNED_PROPERTIES}",
        "                            </xades:QualifyingProperties>",
        "                        </ds:Object>",
        "                    </ds:Signature>",
        "                </sac:SignatureInformation>",
        "            </sig:UBLDocumentSignatures>",
        "        </ext:ExtensionContent>",
        "    </ext:UBLExtension>",
        "</ext:UBLExtensions>");

    // ------------------------------------------------------------------ values, written as C14N writes them
    public static string Amount(decimal x) =>
        Math.Round(x, 2, MidpointRounding.AwayFromZero).ToString("0.00", CultureInfo.InvariantCulture);

    /// <summary>The unit price: 2 to 4 decimals (the stored net price has 4).</summary>
    public static string Price(decimal x) =>
        Math.Round(x, 4, MidpointRounding.AwayFromZero).ToString("0.00##", CultureInfo.InvariantCulture);

    public static string Quantity(decimal x) =>
        Math.Round(x, 6, MidpointRounding.AwayFromZero).ToString("0.000000", CultureInfo.InvariantCulture);

    public static string Text(string? s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    public static string Attr(string? s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace("\"", "&quot;");

    // ------------------------------------------------------------------ the document
    public static string Party(string role, ZatcaParty? p, string i = "  ")
    {
        if (p is null || p.Name.Length == 0)
        {
            return $"{i}<cac:Accounting{role}Party></cac:Accounting{role}Party>";
        }
        var o = new List<string> { $"{i}<cac:Accounting{role}Party>", $"{i}  <cac:Party>" };
        if (p.Crn.Length > 0)
        {
            o.Add($"{i}    <cac:PartyIdentification>");
            o.Add($"{i}      <cbc:ID schemeID=\"CRN\">{Text(p.Crn)}</cbc:ID>");
            o.Add($"{i}    </cac:PartyIdentification>");
        }
        o.Add($"{i}    <cac:PostalAddress>");
        o.Add($"{i}      <cbc:StreetName>{Text(p.Street)}</cbc:StreetName>");
        o.Add($"{i}      <cbc:BuildingNumber>{Text(p.Building)}</cbc:BuildingNumber>");
        if (p.Plot.Length > 0)
        {
            o.Add($"{i}      <cbc:PlotIdentification>{Text(p.Plot)}</cbc:PlotIdentification>");
        }
        if (p.District.Length > 0)
        {
            o.Add($"{i}      <cbc:CitySubdivisionName>{Text(p.District)}</cbc:CitySubdivisionName>");
        }
        o.Add($"{i}      <cbc:CityName>{Text(p.City)}</cbc:CityName>");
        o.Add($"{i}      <cbc:PostalZone>{Text(p.Postal)}</cbc:PostalZone>");
        o.Add($"{i}      <cac:Country>");
        o.Add($"{i}        <cbc:IdentificationCode>{Text(p.Country.Length > 0 ? p.Country : "SA")}</cbc:IdentificationCode>");
        o.Add($"{i}      </cac:Country>");
        o.Add($"{i}    </cac:PostalAddress>");
        if (p.Vat.Length > 0)
        {
            o.Add($"{i}    <cac:PartyTaxScheme>");
            o.Add($"{i}      <cbc:CompanyID>{Text(p.Vat)}</cbc:CompanyID>");
            o.Add($"{i}      <cac:TaxScheme>");
            o.Add($"{i}        <cbc:ID>VAT</cbc:ID>");
            o.Add($"{i}      </cac:TaxScheme>");
            o.Add($"{i}    </cac:PartyTaxScheme>");
        }
        o.Add($"{i}    <cac:PartyLegalEntity>");
        o.Add($"{i}      <cbc:RegistrationName>{Text(p.Name)}</cbc:RegistrationName>");
        o.Add($"{i}    </cac:PartyLegalEntity>");
        o.Add($"{i}  </cac:Party>");
        o.Add($"{i}</cac:Accounting{role}Party>");
        return string.Join("\n", o);
    }

    public static string Line(ZatcaLine ln, string currency = "SAR")
    {
        string c = Attr(currency);
        var o = new List<string>
        {
            "  <cac:InvoiceLine>", $"    <cbc:ID>{ln.Id}</cbc:ID>",
            $"    <cbc:InvoicedQuantity unitCode=\"{Attr(ln.UnitCode)}\">{Quantity(ln.Qty)}</cbc:InvoicedQuantity>",
            $"    <cbc:LineExtensionAmount currencyID=\"{c}\">{Amount(ln.Net)}</cbc:LineExtensionAmount>",
        };
        if (ln.Discount > 0)
        {
            o.AddRange([
                "    <cac:AllowanceCharge>", "      <cbc:ChargeIndicator>false</cbc:ChargeIndicator>",
                "      <cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>",
                $"      <cbc:Amount currencyID=\"{c}\">{Amount(ln.Discount)}</cbc:Amount>", "    </cac:AllowanceCharge>"]);
        }
        o.AddRange([
            "    <cac:TaxTotal>", $"      <cbc:TaxAmount currencyID=\"{c}\">{Amount(ln.Tax)}</cbc:TaxAmount>",
            $"      <cbc:RoundingAmount currencyID=\"{c}\">{Amount(ln.Total)}</cbc:RoundingAmount>", "    </cac:TaxTotal>",
            "    <cac:Item>", $"      <cbc:Name>{Text(ln.Name)}</cbc:Name>", "      <cac:ClassifiedTaxCategory>",
            $"        <cbc:ID>{Text(ln.Category)}</cbc:ID>", $"        <cbc:Percent>{Amount(ln.Percent)}</cbc:Percent>",
            "        <cac:TaxScheme>", "          <cbc:ID>VAT</cbc:ID>", "        </cac:TaxScheme>",
            "      </cac:ClassifiedTaxCategory>", "    </cac:Item>", "    <cac:Price>",
            $"      <cbc:PriceAmount currencyID=\"{c}\">{Price(ln.Price)}</cbc:PriceAmount>", "    </cac:Price>",
            "  </cac:InvoiceLine>"]);
        return string.Join("\n", o);
    }

    public static string TaxSubtotal(ZatcaTaxSubtotal s, string currency = "SAR")
    {
        string c = Attr(currency);
        return string.Join("\n",
            "    <cac:TaxSubtotal>", $"      <cbc:TaxableAmount currencyID=\"{c}\">{Amount(s.Taxable)}</cbc:TaxableAmount>",
            $"      <cbc:TaxAmount currencyID=\"{c}\">{Amount(s.Tax)}</cbc:TaxAmount>", "      <cac:TaxCategory>",
            $"        <cbc:ID schemeAgencyID=\"6\" schemeID=\"UN/ECE 5305\">{Text(s.Category)}</cbc:ID>",
            $"        <cbc:Percent>{Amount(s.Percent)}</cbc:Percent>", "        <cac:TaxScheme>",
            "          <cbc:ID schemeAgencyID=\"6\" schemeID=\"UN/ECE 5153\">VAT</cbc:ID>", "        </cac:TaxScheme>",
            "      </cac:TaxCategory>", "    </cac:TaxSubtotal>");
    }

    /// <summary>The document without the XML declaration and the signature block, with the QR placeholder: already canonical.</summary>
    public static string InvoiceXml(ZatcaInvoice d)
    {
        string c = Attr(d.Currency);
        var o = new List<string>
        {
            $"<Invoice {Namespaces}>", "  <cbc:ProfileID>reporting:1.0</cbc:ProfileID>", $"  <cbc:ID>{Text(d.Id)}</cbc:ID>",
            $"  <cbc:UUID>{Text(d.Uuid)}</cbc:UUID>", $"  <cbc:IssueDate>{d.IssueDate}</cbc:IssueDate>",
            $"  <cbc:IssueTime>{d.IssueTime}</cbc:IssueTime>",
            $"  <cbc:InvoiceTypeCode name=\"{Attr(d.SubType)}\">{Text(d.TypeCode)}</cbc:InvoiceTypeCode>",
            $"  <cbc:DocumentCurrencyCode>{c}</cbc:DocumentCurrencyCode>", $"  <cbc:TaxCurrencyCode>{c}</cbc:TaxCurrencyCode>",
        };
        if (d.BillingRef.Length > 0)
        {
            o.AddRange(["  <cac:BillingReference>", "    <cac:InvoiceDocumentReference>",
                        $"      <cbc:ID>{Text(d.BillingRef)}</cbc:ID>", "    </cac:InvoiceDocumentReference>",
                        "  </cac:BillingReference>"]);
        }
        o.AddRange([
            "  <cac:AdditionalDocumentReference>", "    <cbc:ID>ICV</cbc:ID>", $"    <cbc:UUID>{d.Icv}</cbc:UUID>",
            "  </cac:AdditionalDocumentReference>", "  <cac:AdditionalDocumentReference>", "    <cbc:ID>PIH</cbc:ID>",
            "    <cac:Attachment>",
            $"      <cbc:EmbeddedDocumentBinaryObject mimeCode=\"text/plain\">{Text(d.Pih)}</cbc:EmbeddedDocumentBinaryObject>",
            "    </cac:Attachment>", "  </cac:AdditionalDocumentReference>", "  <cac:AdditionalDocumentReference>",
            "    <cbc:ID>QR</cbc:ID>", "    <cac:Attachment>",
            $"      <cbc:EmbeddedDocumentBinaryObject mimeCode=\"text/plain\">{QrPlaceholder}</cbc:EmbeddedDocumentBinaryObject>",
            "    </cac:Attachment>", "  </cac:AdditionalDocumentReference>", "  <cac:Signature>",
            "    <cbc:ID>urn:oasis:names:specification:ubl:signature:Invoice</cbc:ID>",
            "    <cbc:SignatureMethod>urn:oasis:names:specification:ubl:dsig:enveloped:xades</cbc:SignatureMethod>",
            "  </cac:Signature>", Party("Supplier", d.Seller), Party("Customer", d.Buyer)]);
        if (d.DeliveryDate.Length > 0)
        {
            o.AddRange(["  <cac:Delivery>", $"    <cbc:ActualDeliveryDate>{d.DeliveryDate}</cbc:ActualDeliveryDate>", "  </cac:Delivery>"]);
        }
        o.AddRange(["  <cac:PaymentMeans>", $"    <cbc:PaymentMeansCode>{Text(d.PaymentCode)}</cbc:PaymentMeansCode>"]);
        if (d.Reason.Length > 0)
        {
            o.Add($"    <cbc:InstructionNote>{Text(d.Reason)}</cbc:InstructionNote>");
        }
        o.AddRange([
            "  </cac:PaymentMeans>", "  <cac:TaxTotal>",
            $"    <cbc:TaxAmount currencyID=\"{c}\">{Amount(d.TaxTotal)}</cbc:TaxAmount>", "  </cac:TaxTotal>",
            "  <cac:TaxTotal>", $"    <cbc:TaxAmount currencyID=\"{c}\">{Amount(d.TaxTotal)}</cbc:TaxAmount>"]);
        o.AddRange(d.Subtotals.Select(s => TaxSubtotal(s, d.Currency)));
        o.AddRange([
            "  </cac:TaxTotal>", "  <cac:LegalMonetaryTotal>",
            $"    <cbc:LineExtensionAmount currencyID=\"{c}\">{Amount(d.LineExt)}</cbc:LineExtensionAmount>",
            $"    <cbc:TaxExclusiveAmount currencyID=\"{c}\">{Amount(d.TaxExcl)}</cbc:TaxExclusiveAmount>",
            $"    <cbc:TaxInclusiveAmount currencyID=\"{c}\">{Amount(d.TaxIncl)}</cbc:TaxInclusiveAmount>",
            $"    <cbc:AllowanceTotalAmount currencyID=\"{c}\">0.00</cbc:AllowanceTotalAmount>",
            $"    <cbc:PrepaidAmount currencyID=\"{c}\">0.00</cbc:PrepaidAmount>",
            $"    <cbc:PayableAmount currencyID=\"{c}\">{Amount(d.TaxIncl)}</cbc:PayableAmount>",
            "  </cac:LegalMonetaryTotal>"]);
        o.AddRange(d.Lines.Select(l => Line(l, d.Currency)));
        o.Add("</Invoice>");
        return string.Join("\n", o);
    }

    /// <summary>The file: declaration, the signature block right after the root tag, the QR code in place.</summary>
    public static string SignedDocument(string xml, string extensions, string qr)
    {
        int start = xml.IndexOf('>') + 1;
        return ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + xml[..start] + extensions + xml[start..]).Replace(QrPlaceholder, qr);
    }

    // ------------------------------------------------------------------ hash
    private static string Remove(string xml, string startTag, string endTag, int after = 0)
    {
        int i = xml.IndexOf(startTag, after, StringComparison.Ordinal);
        int j = xml.IndexOf(endTag, i, StringComparison.Ordinal) + endTag.Length;
        return xml[..i] + xml[j..];
    }

    /// <summary>What ZATCA hashes: without the declaration, ext:UBLExtensions, cac:Signature and the QR reference.</summary>
    public static string HashInput(string xml)
    {
        string s = xml[xml.IndexOf("<Invoice", StringComparison.Ordinal)..];
        if (s.Contains("<ext:UBLExtensions>", StringComparison.Ordinal))
        {
            s = Remove(s, "<ext:UBLExtensions>", "</ext:UBLExtensions>");
        }
        s = Remove(s, "<cac:Signature>", "</cac:Signature>");
        int qr = s.IndexOf("<cbc:ID>QR</cbc:ID>", StringComparison.Ordinal);
        int from = s.LastIndexOf("<cac:AdditionalDocumentReference>", qr, StringComparison.Ordinal);
        return Remove(s, "<cac:AdditionalDocumentReference>", "</cac:AdditionalDocumentReference>", from);
    }

    public static string InvoiceHash(string xml) => Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(HashInput(xml))));

    /// <summary>ZATCA's digest of the certificate and of the signed properties: base64 of the hex SHA-256.</summary>
    public static string HexDigestB64(byte[] data) =>
        Convert.ToBase64String(Encoding.ASCII.GetBytes(Convert.ToHexStringLower(SHA256.HashData(data))));

    // ------------------------------------------------------------------ certificate (DER)
    private readonly record struct Der(int Tag, int Start, int Content, int End);

    private static Der Read(byte[] b, int pos)
    {
        int tag = b[pos], length = b[pos + 1], cursor = pos + 2;
        if ((length & 0x80) != 0)
        {
            int n = length & 0x7F;
            length = 0;
            for (int k = 0; k < n; k++)
            {
                length = (length << 8) | b[cursor + k];
            }
            cursor += n;
        }
        return new Der(tag, pos, cursor, cursor + length);
    }

    private static List<Der> Children(byte[] b, Der node)
    {
        var o = new List<Der>();
        for (int cursor = node.Content; cursor < node.End;)
        {
            var child = Read(b, cursor);
            o.Add(child);
            cursor = child.End;
        }
        return o;
    }

    private static byte[] Content(byte[] b, Der n) => b[n.Content..n.End];

    private static string Oid(byte[] content)
    {
        var arcs = new List<long>();
        long value = 0;
        foreach (byte x in content)
        {
            value = (value << 7) | (long)(x & 0x7F);
            if ((x & 0x80) == 0)
            {
                arcs.Add(value);
                value = 0;
            }
        }
        long first = Math.Min(2, arcs[0] / 40);
        return string.Join(".", new[] { first, arcs[0] - first * 40 }.Concat(arcs.Skip(1)));
    }

    private static readonly Dictionary<string, string> DnLabels = new()
    {
        ["2.5.4.3"] = "CN", ["2.5.4.4"] = "SN", ["2.5.4.5"] = "SERIALNUMBER", ["2.5.4.6"] = "C", ["2.5.4.7"] = "L",
        ["2.5.4.8"] = "ST", ["2.5.4.10"] = "O", ["2.5.4.11"] = "OU", ["2.5.4.12"] = "T", ["0.9.2342.19200300.100.1.1"] = "UID",
        ["0.9.2342.19200300.100.1.25"] = "DC", ["1.2.840.113549.1.9.1"] = "E",
    };

    private static string Name(byte[] b, Der node)
    {
        var parts = new List<string>();
        foreach (var rdn in Children(b, node))
        {
            var pair = Children(b, Children(b, rdn)[0]);
            string oid = Oid(Content(b, pair[0]));
            parts.Add((DnLabels.TryGetValue(oid, out var label) ? label : oid) + "=" + Encoding.UTF8.GetString(Content(b, pair[1])));
        }
        parts.Reverse();
        return string.Join(", ", parts);
    }

    /// <summary>The body of a certificate: its base64 text without the PEM lines and white space (modZatcaXml.CertificateBody).</summary>
    public static string CertificateBody(string? text)
    {
        var s = (text ?? "").Replace("-----BEGIN CERTIFICATE-----", "").Replace("-----END CERTIFICATE-----", "");
        return new string(s.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
    }

    /// <summary>Issuer (CN=..., DC=... reversed), decimal serial number, public key (QR tag 8), signature of the
    /// certificate (QR tag 9) and ZATCA's certificate digest.</summary>
    public static ZatcaCertificate CertificateInfo(string certB64)
    {
        byte[] der = Convert.FromBase64String(certB64);
        var top = Children(der, Read(der, 0));
        var items = Children(der, top[0]);
        int off = items[0].Tag == 0xA0 ? 1 : 0;
        var serial = items[off];
        var issuer = items[off + 2];
        var spki = items[off + 5];
        var serialBytes = Content(der, serial);
        return new ZatcaCertificate(
            Name(der, issuer),
            new BigInteger(serialBytes, isUnsigned: true, isBigEndian: true).ToString(CultureInfo.InvariantCulture),
            der[spki.Start..spki.End],
            Content(der, top[2])[1..],
            HexDigestB64(Encoding.ASCII.GetBytes(certB64)));
    }

    // ------------------------------------------------------------------ signature block and QR
    public static string SignedProperties(string signingTime, ZatcaCertificate cert, bool forHash) =>
        (forHash ? SignedPropertiesForHash : SignedPropertiesForXml).Replace("{SIGNING_TIME}", signingTime)
            .Replace("{CERT_HASH}", cert.Hash).Replace("{CERT_ISSUER}", Text(cert.Issuer)).Replace("{CERT_SERIAL}", cert.Serial);

    public static string UblExtensions(string invoiceHash, string signature, string certB64, string signingTime, ZatcaCertificate cert)
    {
        string digest = HexDigestB64(Encoding.UTF8.GetBytes(SignedProperties(signingTime, cert, true)));
        return UblExtensionsTemplate.Replace("{INVOICE_HASH}", invoiceHash).Replace("{PROPS_DIGEST}", digest)
            .Replace("{SIGNATURE}", signature).Replace("{CERTIFICATE}", certB64)
            .Replace("{SIGNED_PROPERTIES}", SignedProperties(signingTime, cert, false));
    }

    private static void Tlv(List<byte> o, int tag, byte[] v)
    {
        o.Add((byte)tag);
        int n = v.Length;
        if (n < 0x80)
        {
            o.Add((byte)n);
        }
        else if (n < 0x100)
        {
            o.AddRange([0x81, (byte)n]);
        }
        else
        {
            o.AddRange([0x82, (byte)(n >> 8), (byte)(n & 0xFF)]);
        }
        o.AddRange(v);
    }

    /// <summary>Nine tags: 1-5 as phase 1, 6 invoice hash, 7 signature (base64 text), 8 public key, 9 signature of the certificate.</summary>
    public static string QrCode(string seller, string vat, string timestamp, decimal total, decimal vatTotal, string invoiceHash,
                                string signature, byte[] publicKey, byte[] certSignature)
    {
        var o = new List<byte>();
        byte[][] values =
        [
            Encoding.UTF8.GetBytes(seller), Encoding.UTF8.GetBytes(vat), Encoding.UTF8.GetBytes(timestamp),
            Encoding.UTF8.GetBytes(Amount(total)), Encoding.UTF8.GetBytes(Amount(vatTotal)), Encoding.UTF8.GetBytes(invoiceHash),
            Encoding.UTF8.GetBytes(signature), publicKey, certSignature,
        ];
        for (int k = 0; k < values.Length; k++)
        {
            Tlv(o, k + 1, values[k]);
        }
        return Convert.ToBase64String(o.ToArray());
    }

    /// <summary>The QR code inside a cleared document (modZatcaApi.ClearedQR).</summary>
    public static string ClearedQr(string xml)
    {
        int i = xml.IndexOf(">QR</cbc:ID>", StringComparison.Ordinal);
        if (i < 0)
        {
            return "";
        }
        int j = xml.IndexOf("EmbeddedDocumentBinaryObject", i, StringComparison.Ordinal);
        if (j < 0)
        {
            return "";
        }
        j = xml.IndexOf('>', j) + 1;
        int k = xml.IndexOf('<', j);
        return j > 0 && k > j ? xml[j..k] : "";
    }

    // ------------------------------------------------------------------ signing
    /// <summary>modZatcaXml.ZatcaSignHash: the ECDSA (secp256k1, SHA-256, DER) signature of the 32 bytes of the invoice hash,
    /// as "openssl dgst -sha256 -sign key" makes it, with the private key of the device (PEM).</summary>
    public static string SignHash(string invoiceHashB64, string privateKeyPem)
    {
        using var ec = ECDsa.Create();
        ec.ImportFromPem(privateKeyPem);
        return Convert.ToBase64String(ec.SignData(Convert.FromBase64String(invoiceHashB64), HashAlgorithmName.SHA256,
                                                  DSASignatureFormat.Rfc3279DerSequence));
    }

    /// <summary>modZatcaXml.ZatcaSignNow: the signing time of the signed properties, Saudi time to UTC.</summary>
    public static string SigningTime(DateTime localSaudi) =>
        localSaudi.AddHours(-KsaUtcOffsetHours).ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);

    /// <summary>modZatcaXml.ZatcaSignXml: the signed file, its hash and its QR code. Totals: issue timestamp, total with VAT,
    /// VAT, seller name, seller VAT number.</summary>
    public static (string SignedXml, string Hash, string Qr) Sign(string xml, string certB64, string privateKeyPem, string signingTime,
                                                                  (string Timestamp, decimal Total, decimal Vat, string Seller, string SellerVat) totals)
    {
        string body = CertificateBody(certB64);
        var cert = CertificateInfo(body);
        string hash = InvoiceHash(xml);
        string signature = SignHash(hash, privateKeyPem);
        string qr = QrCode(totals.Seller, totals.SellerVat, totals.Timestamp, totals.Total, totals.Vat, hash, signature,
                           cert.PublicKey, cert.Signature);
        return (SignedDocument(xml, UblExtensions(hash, signature, body, signingTime, cert), qr), hash, qr);
    }
}
