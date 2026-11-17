using System.Text.Json;
using System.Text.Json.Nodes;

namespace RetailStore.Domain.EInvoicing;

/// <summary>The reading of a platform reply: Result = OK, WARNING, REJECTED, ERROR or NETWORK (modEInvoice.SplitResult);
/// DocStatus = the new e-invoicing status of the document ("" = it stays as it is); the ETA ids when given.</summary>
public sealed record EInvoiceReply(string Result, string DocStatus, string Message, string SubmissionId = "", string Uuid = "",
                                   string LongId = "")
{
    public bool Sent => Result is "OK" or "WARNING";
}

/// <summary>JSON paths of the replies, as modHttp.JsonGet reads them: "a.b[0].c"; missing = null.</summary>
public static class JsonPath
{
    public static JsonNode? Parse(string? body)
    {
        if (string.IsNullOrEmpty(body))
        {
            return null;
        }
        try
        {
            return JsonNode.Parse(body);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    public static JsonNode? Get(JsonNode? node, string path)
    {
        foreach (var part in path.Split('.'))
        {
            string name = part;
            int bracket = part.IndexOf('[');
            int? index = null;
            if (bracket >= 0)
            {
                name = part[..bracket];
                index = int.Parse(part[(bracket + 1)..part.IndexOf(']')], System.Globalization.CultureInfo.InvariantCulture);
            }
            if (name.Length > 0)
            {
                node = node is JsonObject o && o.TryGetPropertyValue(name, out var v) ? v : null;
            }
            if (index is int k)
            {
                node = node is JsonArray a && a.Count > k ? a[k] : null;
            }
            if (node is null)
            {
                return null;
            }
        }
        return node;
    }

    /// <summary>The value as text (a string unquoted, anything else as written); null when missing or JSON null.</summary>
    public static string? Text(JsonNode? node, string path)
    {
        var n = Get(node, path);
        if (n is null)
        {
            return null;
        }
        return n is JsonValue v && v.TryGetValue<string>(out var s) ? s : n.ToJsonString();
    }

    /// <summary>modEtaReceipt.EtaFirstText: the first value that is not missing nor empty.</summary>
    public static string First(JsonNode? node, params string[] paths)
    {
        foreach (var p in paths)
        {
            if (p.Length == 0)
            {
                continue;
            }
            var t = Text(node, p);
            if (!string.IsNullOrEmpty(t))
            {
                return t;
            }
        }
        return "";
    }
}

/// <summary>ZATCA replies (modZatcaApi, tools/zatca_api_reference.py).</summary>
public static class ZatcaReplies
{
    public static string BaseUrl(string env) => env switch
    {
        "PRODUCTION" => "https://gw-fatoora.zatca.gov.sa/e-invoicing/core",
        "SIMULATION" => "https://gw-fatoora.zatca.gov.sa/e-invoicing/simulation",
        _ => "https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal",
    };

    /// <summary>"code: message; ..." (at most 3) of validationResults.errorMessages / warningMessages.</summary>
    public static string Messages(string body, string kind)
    {
        var root = JsonPath.Parse(body);
        var parts = new List<string>();
        for (int i = 0; i < 3; i++)
        {
            string p = "validationResults." + kind + "[" + i + "]";
            if (JsonPath.Get(root, p) is null)
            {
                break;
            }
            parts.Add((JsonPath.Text(root, p + ".code") ?? "") + ": " + (JsonPath.Text(root, p + ".message") ?? ""));
        }
        return string.Join("; ", parts);
    }

    private static string Left(string s, int n) => s.Length <= n ? s : s[..n];

    /// <summary>modZatcaApi.ZatcaReplyResult: the result of a reporting / clearance reply.</summary>
    public static EInvoiceReply Result(int httpStatus, string body, bool standard)
    {
        if (httpStatus == 0)
        {
            return new("NETWORK", "", "تعذّر الاتصال بالمنظومة.");
        }
        if (httpStatus == 200)
        {
            return new("OK", standard ? "CLEARED" : "REPORTED", "");
        }
        if (httpStatus == 202)
        {
            return new("WARNING", "WARNING", Messages(body, "warningMessages"));
        }
        if (httpStatus == 400)
        {
            string m = Messages(body, "errorMessages");
            return new("REJECTED", "REJECTED", m.Length > 0 ? m : Left(body, 200));
        }
        if (httpStatus is 401 or 403)
        {
            return new("ERROR", "", "رفضت الهيئة بيانات الدخول (شهادة الجهاز أو الكلمة السرية).");
        }
        if (httpStatus == 429 || httpStatus >= 500)
        {
            return new("NETWORK", "", "المنظومة غير متاحة الآن (رمز " + httpStatus + ").");
        }
        return new("ERROR", "", "رد غير متوقع من الهيئة (رمز " + httpStatus + "): " + Left(body, 200));
    }
}
