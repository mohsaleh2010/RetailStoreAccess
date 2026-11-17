using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace RetailStore.Domain.EInvoicing;

/// <summary>One line of an e-receipt (modEtaReceipt.EtaItemJson).</summary>
public sealed record EtaReceiptItem(string InternalCode, string Description, string ItemType, string ItemCode, string UnitType,
                                    decimal Quantity, decimal UnitPrice, decimal TotalSale, decimal Discount, decimal NetSale,
                                    decimal Tax, decimal Rate, decimal Total);

/// <summary>An e-receipt (modEtaReceipt.EtaReceiptJson); ReferenceUuid = the sale of a return.</summary>
public sealed record EtaReceipt(string DateTime, string Number, string PreviousUuid, string ReferenceUuid, string Rin,
                                string TradeName, string BranchCode, string Governate, string City, string Street, string Building,
                                string Postal, string DeviceSerial, string ActivityCode, string BuyerType, string BuyerId,
                                string BuyerName, string BuyerMobile, IReadOnlyList<EtaReceiptItem> Items, decimal TotalSales,
                                decimal TotalDiscount, decimal NetAmount, decimal TaxTotal, decimal TotalAmount, string PaymentMethod);

/// <summary>The issuer (with its branch) or the receiver of an e-invoice (modEtaInvoice.EtaPartyJson).</summary>
public sealed record EtaParty(string Role, string Type, string Id, string Name, string BranchId, string Governate, string City,
                              string Street, string Building, string Postal);

/// <summary>One line of an e-invoice (modEtaInvoice.EtaLineJson).</summary>
public sealed record EtaInvoiceLine(string Description, string ItemType, string ItemCode, string UnitType, decimal Quantity,
                                    string InternalCode, decimal UnitValue, decimal SalesTotal, decimal Discount, decimal NetTotal,
                                    decimal Tax, decimal Rate, decimal Total);

/// <summary>An e-invoice: DocType I invoice / C credit note (ReferenceUuid = the invoice); Version 1.0 signed, 0.9 unsigned.</summary>
public sealed record EtaInvoice(string DocType, string Version, string DateTime, string ActivityCode, string InternalId,
                                EtaParty Issuer, EtaParty Receiver, IReadOnlyList<EtaInvoiceLine> Lines, string ReferenceUuid,
                                decimal TotalDiscount, decimal TotalSales, decimal NetAmount, decimal TaxTotal, decimal TotalAmount);

/// <summary>The Egyptian documents in C# (docs/48, docs/49, docs/85): modEtaReceipt, modEtaInvoice and their Python mirrors
/// tools/eta_receipt_reference.py and tools/eta_invoice_reference.py; compact JSON, the ETA serialization, the replies.</summary>
public static class EtaDocument
{
    public const string TypeVersion = "1.2";
    public const decimal BuyerIdLimit = 150000m;
    public const string DefaultSignerArgs = "\"{IN}\" \"{OUT}\" \"{PIN}\"";

    public static string Env(string? environment) => environment == "PRODUCTION" ? "PROD" : "PREPROD";

    public static string TokenUrl(string env) =>
        env == "PROD" ? "https://id.eta.gov.eg/connect/token" : "https://id.preprod.eta.gov.eg/connect/token";

    public static string ApiUrl(string env) =>
        env == "PROD" ? "https://api.invoicing.eta.gov.eg" : "https://api.preprod.invoicing.eta.gov.eg";

    public static string PortalUrl(string env) =>
        env == "PROD" ? "https://invoicing.eta.gov.eg" : "https://preprod.invoicing.eta.gov.eg";

    // ------------------------------------------------------------------ values
    /// <summary>A JSON number: rounded to 5 decimals (half even, as VBA Round), no trailing zero.</summary>
    public static string Num(decimal value)
    {
        decimal d = Math.Round(value, 5, MidpointRounding.ToEven);
        return d == 0 ? "0" : d.ToString("0.#####", CultureInfo.InvariantCulture);
    }

    /// <summary>The inside of a JSON string, as modHttp.JsonEscape writes it.</summary>
    public static string Esc(string? s)
    {
        var o = new StringBuilder();
        foreach (char ch in s ?? "")
        {
            switch (ch)
            {
                case '"': o.Append("\\\""); break;
                case '\\': o.Append("\\\\"); break;
                case '\n': o.Append("\\n"); break;
                case '\r': o.Append("\\r"); break;
                case '\t': o.Append("\\t"); break;
                default:
                    if (ch < 32)
                    {
                        o.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        o.Append(ch);
                    }
                    break;
            }
        }
        return o.ToString();
    }

    public static string JStr(string name, string? value) => "\"" + name + "\":\"" + Esc(value) + "\"";

    public static string JNum(string name, decimal value) => "\"" + name + "\":" + Num(value);

    // ------------------------------------------------------------------ the receipt
    public static string ItemJson(EtaReceiptItem it)
    {
        var p = new List<string>
        {
            JStr("internalCode", it.InternalCode), JStr("description", it.Description), JStr("itemType", it.ItemType),
            JStr("itemCode", it.ItemCode), JStr("unitType", it.UnitType), JNum("quantity", it.Quantity),
            JNum("unitPrice", it.UnitPrice), JNum("netSale", it.NetSale), JNum("totalSale", it.TotalSale), JNum("total", it.Total),
        };
        if (it.Discount != 0)
        {
            p.Add("\"commercialDiscountData\":[{" + JNum("amount", it.Discount) + "," + JStr("description", "Discount") + "}]");
        }
        p.Add("\"taxableItems\":[{" + JStr("taxType", "T1") + "," + JNum("amount", it.Tax) + "," + JStr("subType", "V009") + "," +
              JNum("rate", it.Rate) + "}]");
        return "{" + string.Join(",", p) + "}";
    }

    public static string ReceiptJson(EtaReceipt r, string uuid = "")
    {
        bool isReturn = r.ReferenceUuid.Length > 0;
        var header = new List<string> { JStr("dateTimeIssued", r.DateTime), JStr("receiptNumber", r.Number), JStr("uuid", uuid),
                                        JStr("previousUUID", r.PreviousUuid) };
        if (isReturn)
        {
            header.Add(JStr("referenceUUID", r.ReferenceUuid));
        }
        header.AddRange([JStr("currency", "EGP"), JNum("exchangeRate", 0)]);
        var address = new List<string> { JStr("country", "EG"), JStr("governate", r.Governate), JStr("regionCity", r.City),
                                         JStr("street", r.Street), JStr("buildingNumber", r.Building) };
        if (r.Postal.Length > 0)
        {
            address.Add(JStr("postalCode", r.Postal));
        }
        var seller = new List<string> { JStr("rin", r.Rin), JStr("companyTradeName", r.TradeName), JStr("branchCode", r.BranchCode),
                                        "\"branchAddress\":{" + string.Join(",", address) + "}",
                                        JStr("deviceSerialNumber", r.DeviceSerial), JStr("activityCode", r.ActivityCode) };
        var buyer = new List<string> { JStr("type", r.BuyerType) };
        if (r.BuyerId.Length > 0)
        {
            buyer.Add(JStr("id", r.BuyerId));
        }
        if (r.BuyerName.Length > 0)
        {
            buyer.Add(JStr("name", r.BuyerName));
        }
        if (r.BuyerMobile.Length > 0)
        {
            buyer.Add(JStr("mobileNumber", r.BuyerMobile));
        }
        var parts = new List<string>
        {
            "\"header\":{" + string.Join(",", header) + "}",
            "\"documentType\":{" + JStr("receiptType", isReturn ? "R" : "S") + "," + JStr("typeVersion", TypeVersion) + "}",
            "\"seller\":{" + string.Join(",", seller) + "}", "\"buyer\":{" + string.Join(",", buyer) + "}",
            "\"itemData\":[" + string.Join(",", r.Items.Select(ItemJson)) + "]",
            JNum("totalSales", r.TotalSales), JNum("totalCommercialDiscount", r.TotalDiscount), JNum("totalItemsDiscount", 0),
            JNum("netAmount", r.NetAmount), JNum("feesAmount", 0), JNum("totalAmount", r.TotalAmount),
            "\"taxTotals\":[{" + JStr("taxType", "T1") + "," + JNum("amount", r.TaxTotal) + "}]",
            JStr("paymentMethod", r.PaymentMethod),
        };
        return "{" + string.Join(",", parts) + "}";
    }

    /// <summary>modEtaReceipt.EtaWithUuid: the receipt with its UUID in header.uuid (built with "").</summary>
    public static string WithUuid(string json, string uuid)
    {
        const string empty = "\"uuid\":\"\"";
        int i = json.IndexOf(empty, StringComparison.Ordinal);
        return i < 0 ? json : json[..i] + "\"uuid\":\"" + uuid + "\"" + json[(i + empty.Length)..];
    }

    /// <summary>The ETA serialization: every property is "NAME" (upper case) then its value; a simple value is "value"; an array
    /// is "NAME" then, per element, "NAME" and the element. Numbers keep their text.</summary>
    public static string Serialize(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var o = new StringBuilder();
        Serialize(doc.RootElement, o);
        return o.ToString();
    }

    private static void Serialize(JsonElement e, StringBuilder o)
    {
        switch (e.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var p in e.EnumerateObject())
                {
                    string key = "\"" + p.Name.ToUpperInvariant() + "\"";
                    if (p.Value.ValueKind == JsonValueKind.Array)
                    {
                        o.Append(key);
                        foreach (var el in p.Value.EnumerateArray())
                        {
                            o.Append(key);
                            Serialize(el, o);
                        }
                    }
                    else
                    {
                        o.Append(key);
                        Serialize(p.Value, o);
                    }
                }
                break;
            case JsonValueKind.String:
                o.Append('"').Append(e.GetString()).Append('"');
                break;
            default:
                o.Append('"').Append(e.GetRawText()).Append('"');
                break;
        }
    }

    /// <summary>The UUID of a receipt whose header.uuid is "": hex SHA-256 of its serialization (UTF-8).</summary>
    public static string ReceiptUuid(string json) =>
        Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(Serialize(json))));

    public static (string Uuid, string Json) SignedReceipt(EtaReceipt r)
    {
        string uuid = ReceiptUuid(ReceiptJson(r));
        return (uuid, ReceiptJson(r, uuid));
    }

    public static string ReceiptQrLink(string env, string uuid, string dateTime, decimal totalAmount, string rin) =>
        PortalUrl(env) + "/receipts/search/" + uuid + "/share/" + dateTime + "#Total:" + Num(totalAmount) + ",IssuerRIN:" + rin;

    /// <summary>C cash, V visa (card), O others (bank transfer, sale on account).</summary>
    public static string PaymentCode(string? paymentType, int? methodId)
    {
        if (paymentType == "CREDIT")
        {
            return "O";
        }
        return (methodId ?? 1) switch { 2 => "V", 3 => "O", _ => "C" };
    }

    /// <summary>modEtaReceipt.EtaUtcTimestamp: the local time of Egypt in UTC ("2026-10-09T12:30:00Z").</summary>
    public static string UtcTimestamp(DateTime local)
    {
        DateTime utc;
        try
        {
            var zone = TimeZoneInfo.FindSystemTimeZoneById("Africa/Cairo");
            utc = TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(local, DateTimeKind.Unspecified), zone);
        }
        catch (Exception e) when (e is TimeZoneNotFoundException or InvalidTimeZoneException or ArgumentException)
        {
            utc = local.AddHours(-2);
        }
        return utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
    }

    public static string UrlEncode(string value) => Uri.EscapeDataString(value);

    // ------------------------------------------------------------------ the e-invoice
    public static string PartyJson(EtaParty p)
    {
        var address = new List<string>();
        if (p.Role == "issuer")
        {
            address.Add(JStr("branchID", p.BranchId));
        }
        address.AddRange([JStr("country", "EG"), JStr("governate", p.Governate), JStr("regionCity", p.City), JStr("street", p.Street),
                          JStr("buildingNumber", p.Building)]);
        if (p.Postal.Length > 0)
        {
            address.Add(JStr("postalCode", p.Postal));
        }
        var parts = new List<string> { "\"address\":{" + string.Join(",", address) + "}", JStr("type", p.Type) };
        if (p.Id.Length > 0)
        {
            parts.Add(JStr("id", p.Id));
        }
        parts.Add(JStr("name", p.Name));
        return "\"" + p.Role + "\":{" + string.Join(",", parts) + "}";
    }

    public static string LineJson(EtaInvoiceLine it) => "{" + string.Join(",",
        JStr("description", it.Description), JStr("itemType", it.ItemType), JStr("itemCode", it.ItemCode),
        JStr("unitType", it.UnitType), JNum("quantity", it.Quantity), JStr("internalCode", it.InternalCode),
        JNum("salesTotal", it.SalesTotal), JNum("total", it.Total), JNum("valueDifference", 0), JNum("totalTaxableFees", 0),
        JNum("netTotal", it.NetTotal), JNum("itemsDiscount", 0),
        "\"unitValue\":{" + JStr("currencySold", "EGP") + "," + JNum("amountEGP", it.UnitValue) + "}",
        "\"discount\":{" + JNum("rate", 0) + "," + JNum("amount", it.Discount) + "}",
        "\"taxableItems\":[{" + JStr("taxType", "T1") + "," + JNum("amount", it.Tax) + "," + JStr("subType", "V009") + "," +
        JNum("rate", it.Rate) + "}]") + "}";

    public static string InvoiceJson(EtaInvoice d)
    {
        var parts = new List<string>
        {
            PartyJson(d.Issuer), PartyJson(d.Receiver), JStr("documentType", d.DocType), JStr("documentTypeVersion", d.Version),
            JStr("dateTimeIssued", d.DateTime), JStr("taxpayerActivityCode", d.ActivityCode), JStr("internalID", d.InternalId),
        };
        if (d.ReferenceUuid.Length > 0)
        {
            parts.Add("\"references\":[\"" + Esc(d.ReferenceUuid) + "\"]");
        }
        parts.AddRange([
            "\"invoiceLines\":[" + string.Join(",", d.Lines.Select(LineJson)) + "]", JNum("totalDiscountAmount", d.TotalDiscount),
            JNum("totalSalesAmount", d.TotalSales), JNum("netAmount", d.NetAmount),
            "\"taxTotals\":[{" + JStr("taxType", "T1") + "," + JNum("amount", d.TaxTotal) + "}]",
            JNum("totalAmount", d.TotalAmount), JNum("extraDiscountAmount", 0), JNum("totalItemsDiscountAmount", 0)]);
        return "{" + string.Join(",", parts) + "}";
    }

    public static string WithSignature(string doc, string signature) =>
        doc[..^1] + ",\"signatures\":[{" + JStr("signatureType", "I") + "," + JStr("value", signature) + "}]}";

    /// <summary>The arguments of the signing program: the template with {IN} {OUT} {PIN} (the program itself is started apart).</summary>
    public static string SignerArguments(string? template, string inFile, string outFile, string pin) =>
        (string.IsNullOrEmpty(template) ? DefaultSignerArgs : template).Replace("{IN}", inFile).Replace("{OUT}", outFile)
            .Replace("{PIN}", pin);

    /// <summary>modEtaInvoice.EtaSignerCommand: the quoted program, then the arguments.</summary>
    public static string SignerCommand(string? template, string exe, string inFile, string outFile, string pin) =>
        "\"" + exe + "\" " + SignerArguments(template, inFile, outFile, pin);

    public static string InvoiceQrLink(string env, string uuid, string longId) => PortalUrl(env) + "/documents/" + uuid + "/share/" + longId;

    // ------------------------------------------------------------------ replies
    private static string Left(string s, int n) => s.Length <= n ? s : s[..n];

    /// <summary>modEtaReceipt.EtaErrors: "message (path: message; ...)" of the ETA error object at path (at most 3 details).</summary>
    public static string Errors(JsonNode? root, string path)
    {
        string msg = JsonPath.First(root, path + ".message", path + ".code");
        var d = new List<string>();
        for (int i = 0; i < 3; i++)
        {
            string p = path + ".details[" + i + "]";
            if (JsonPath.Get(root, p) is null)
            {
                break;
            }
            d.Add(JsonPath.First(root, p + ".propertyPath", p + ".target", p + ".code") + ": " + JsonPath.First(root, p + ".message"));
        }
        string details = string.Join("; ", d);
        if (msg.Length > 0 && details.Length > 0)
        {
            return msg + " (" + details + ")";
        }
        return msg.Length > 0 ? msg : details;
    }

    /// <summary>modEtaReceipt.EtaSubmitResult: OK (SUBMITTED), REJECTED, ERROR or NETWORK of a receipt submission.</summary>
    public static EInvoiceReply ReceiptSubmitResult(int httpStatus, string body)
    {
        var root = JsonPath.Parse(body);
        if (httpStatus == 0)
        {
            return new("NETWORK", "", "تعذّر الاتصال بالمنظومة.");
        }
        if (httpStatus is 200 or 202)
        {
            if (JsonPath.Get(root, "rejectedDocuments[0]") is not null)
            {
                string m = Errors(root, "rejectedDocuments[0].error");
                return new("REJECTED", "REJECTED", m.Length > 0 ? m : "رُفض الإيصال.");
            }
            if (JsonPath.Get(root, "acceptedDocuments[0]") is not null)
            {
                return new("OK", "SUBMITTED", "", JsonPath.Text(root, "submissionId") ?? "");
            }
            return new("ERROR", "", "رد غير متوقع من المصلحة: " + Left(body, 200));
        }
        if (httpStatus == 400)
        {
            string m = Errors(root, "error");
            return new("REJECTED", "REJECTED", m.Length > 0 ? m : Left(body, 200));
        }
        if (httpStatus is 401 or 403)
        {
            return new("ERROR", "", "رفضت المصلحة بيانات الدخول (بيانات جهاز نقطة البيع).");
        }
        if (httpStatus == 429 || httpStatus >= 500)
        {
            return new("NETWORK", "", "المنظومة غير متاحة الآن (رمز " + httpStatus + ").");
        }
        return new("ERROR", "", "رد غير متوقع من المصلحة (رمز " + httpStatus + "): " + Left(body, 200));
    }

    /// <summary>modEtaReceipt.EtaStatusResult: VALID, INVALID (with the errors) or "" (still in progress).</summary>
    public static (string Status, string Message) ReceiptStatusResult(string body)
    {
        var root = JsonPath.Parse(body);
        string st = JsonPath.First(root, "receipts[0].status", "status").ToLowerInvariant();
        if (st == "valid")
        {
            return ("VALID", "");
        }
        if (st != "invalid")
        {
            return ("", "");
        }
        var d = new List<string>();
        for (int i = 0; i < 3; i++)
        {
            string p = "receipts[0].errors[" + i + "]";
            if (JsonPath.Get(root, p) is null)
            {
                break;
            }
            d.Add(JsonPath.First(root, p + ".propertyPath", p + ".target", p + ".code") + ": " + JsonPath.First(root, p + ".message"));
        }
        return ("INVALID", string.Join("; ", d));
    }

    /// <summary>modEtaInvoice.EtaDocSubmitResult: OK (SUBMITTED, with ETA's UUID and long ID), REJECTED, ERROR or NETWORK.</summary>
    public static EInvoiceReply InvoiceSubmitResult(int httpStatus, string body)
    {
        var root = JsonPath.Parse(body);
        if (httpStatus == 0)
        {
            return new("NETWORK", "", "تعذّر الاتصال بالمنظومة.");
        }
        if (httpStatus is 200 or 202)
        {
            if (JsonPath.Get(root, "rejectedDocuments[0]") is not null)
            {
                string m = Errors(root, "rejectedDocuments[0].error");
                return new("REJECTED", "REJECTED", m.Length > 0 ? m : "رُفض المستند.");
            }
            if (JsonPath.Get(root, "acceptedDocuments[0]") is not null)
            {
                return new("OK", "SUBMITTED", "", JsonPath.First(root, "submissionId"), JsonPath.First(root, "acceptedDocuments[0].uuid"),
                           JsonPath.First(root, "acceptedDocuments[0].longId"));
            }
            return new("ERROR", "", "رد غير متوقع من المصلحة: " + Left(body, 200));
        }
        if (httpStatus == 400)
        {
            string m = Errors(root, "error");
            return new("REJECTED", "REJECTED", m.Length > 0 ? m : Left(body, 200));
        }
        if (httpStatus is 401 or 403)
        {
            return new("ERROR", "", "رفضت المصلحة بيانات دخول البرنامج (Client ID و Client Secret).");
        }
        if (httpStatus == 429 || httpStatus >= 500)
        {
            return new("NETWORK", "", "المنظومة غير متاحة الآن (رمز " + httpStatus + ").");
        }
        return new("ERROR", "", "رد غير متوقع من المصلحة (رمز " + httpStatus + "): " + Left(body, 200));
    }

    /// <summary>modEtaInvoice.EtaDocStatusResult: VALID, INVALID (the failed validation steps), CANCELLED or "" (in progress).</summary>
    public static (string Status, string Message) InvoiceStatusResult(string body)
    {
        var root = JsonPath.Parse(body);
        string st = JsonPath.First(root, "status").ToLowerInvariant();
        if (st == "valid")
        {
            return ("VALID", "");
        }
        if (st is "cancelled" or "canceled")
        {
            return ("CANCELLED", "");
        }
        if (st != "invalid")
        {
            return ("", "");
        }
        var o = new List<string>();
        for (int i = 0; i < 10 && o.Count < 3; i++)
        {
            string p = "validationResults.validationSteps[" + i + "]";
            if (JsonPath.Get(root, p) is null)
            {
                break;
            }
            if (JsonPath.First(root, p + ".status").ToLowerInvariant() == "invalid")
            {
                o.Add(JsonPath.First(root, p + ".name") + ": " +
                      JsonPath.First(root, p + ".error.innerError[0].error", p + ".error.error", p + ".error.errorCode"));
            }
        }
        return ("INVALID", string.Join("; ", o));
    }
}
