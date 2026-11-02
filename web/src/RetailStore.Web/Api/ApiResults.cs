using Microsoft.AspNetCore.Mvc;
using RetailStore.Data.Services;
using RetailStore.Domain.Localization;

namespace RetailStore.Web.Api;

/// <summary>The answers of the API (docs/77). Errors are RFC 9457 problem details: title = the message of the program
/// (Arabic, or English when the request asks for it: Accept-Language en or ?lang=en), code = what happened (refused,
/// confirm, forbidden, not_found ...), field = the field to fix.</summary>
public static class ApiResults
{
    public static bool English(HttpContext http)
    {
        if (http.Request.Query.TryGetValue("lang", out var lang))
        {
            return string.Equals(lang.ToString(), "en", StringComparison.OrdinalIgnoreCase);
        }
        var header = http.Request.GetTypedHeaders().AcceptLanguage;
        var first = header.OrderByDescending(h => h.Quality ?? 1).FirstOrDefault();
        return first is not null && first.Value.Value?.StartsWith("en", StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>A message of the program in the language of the request (modLang.Tr).</summary>
    public static string Text(HttpContext http, string arabic) => English(http) ? Translator.ToEnglish(arabic) : arabic;

    public static IResult Problem(HttpContext http, int status, string code, string message, string field = "")
    {
        var p = new ProblemDetails { Status = status, Title = Text(http, message), Type = "https://retailstore/api/errors/" + code };
        p.Extensions["code"] = code;
        if (field.Length > 0)
        {
            p.Extensions["field"] = field;
        }
        return Results.Problem(p);
    }

    public static IResult Refused(HttpContext http, string message, string field = "") =>
        Problem(http, StatusCodes.Status422UnprocessableEntity, "refused", message, field);

    public static IResult NotFound(HttpContext http) => Problem(http, StatusCodes.Status404NotFound, "not_found", "السجل غير موجود.");

    /// <summary>A save or delete of a data screen: Saved -> ok(), Refused -> 422, NeedsConfirmation -> 409 (the client asks
    /// the question and sends the same request again with ?confirm=true).</summary>
    public static IResult Of(HttpContext http, SaveResult r, Func<IResult> ok) => r.Status switch
    {
        SaveStatus.Saved => ok(),
        SaveStatus.NeedsConfirmation => Problem(http, StatusCodes.Status409Conflict, "confirm", r.Message, r.Field),
        _ => Refused(http, r.Message, r.Field),
    };

    /// <summary>A service that answers "" or the message.</summary>
    public static IResult OfMessage(HttpContext http, string problem, Func<IResult> ok) =>
        problem.Length == 0 ? ok() : Refused(http, problem);
}

/// <summary>The answer of a data screen save: the key, the message and the record as it is now.</summary>
public sealed record SavedRecord(int Key, string Message, Dictionary<string, object?>? Record);

/// <summary>A page of records: how many there are in all, and these.</summary>
public sealed record RecordPage(int Total, int Skip, int Take, List<Dictionary<string, object?>> Items);
