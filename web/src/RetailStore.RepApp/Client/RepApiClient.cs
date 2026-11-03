using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;

namespace RetailStore.RepApp.Client;

/// <summary>The answer of one call: the value, or the HTTP status with the code and the message of the problem details
/// (docs/77 section 3). Status 0 = the server was not reached (no network, timeout).</summary>
public sealed record ApiResult<T>(T? Value, int Status, string Code, string Message)
{
    public bool Ok => Status is >= 200 and < 300;

    /// <summary>Try again later: not reached, or the server failed (5xx), or too many logins (429).</summary>
    public bool Retry => Status == 0 || Status >= 500 || Status == 429;

    /// <summary>The token is missing, expired or ended (a new password, a disabled user): log in again.</summary>
    public bool NeedsLogin => Status == 401;
}

/// <summary>The calls of the rep app to the API (/api/v1): the token, the language of the messages (Accept-Language),
/// and the address of the server (Settings). One HttpClient for the life of the app.</summary>
public sealed class RepApiClient(HttpClient http)
{
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public string ServerUrl { get; set; } = "";
    public string? Token { get; set; }
    public bool English { get; set; }

    public static string NormalizeServer(string? url)
    {
        string u = (url ?? "").Trim().TrimEnd('/');
        if (u.Length > 0 && !u.Contains("://", StringComparison.Ordinal))
        {
            u = "https://" + u;
        }
        return u;
    }

    private Uri Address(string path)
    {
        string root = NormalizeServer(ServerUrl);
        if (root.Length == 0 && http.BaseAddress is not null)
        {
            root = http.BaseAddress.ToString().TrimEnd('/');
        }
        return new Uri(root + "/api/v1/" + path);
    }

    private async Task<ApiResult<T>> SendAsync<T>(HttpMethod method, string path, object? body, CancellationToken ct)
    {
        try
        {
            using var request = new HttpRequestMessage(method, Address(path));
            if (Token is { Length: > 0 })
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Token);
            }
            request.Headers.AcceptLanguage.Add(new StringWithQualityHeaderValue(English ? "en" : "ar"));
            if (body is not null)
            {
                request.Content = JsonContent.Create(body, body.GetType(), options: Json);
            }
            using var response = await http.SendAsync(request, ct);
            int status = (int)response.StatusCode;
            if (response.IsSuccessStatusCode)
            {
                var value = await response.Content.ReadFromJsonAsync<T>(Json, ct);
                return new ApiResult<T>(value, status, "", "");
            }
            string code = "", message = response.ReasonPhrase ?? status.ToString();
            if (response.Content.Headers.ContentType?.MediaType == "application/problem+json")
            {
                var p = await response.Content.ReadFromJsonAsync<JsonElement>(Json, ct);
                code = p.TryGetProperty("code", out var c) ? c.GetString() ?? "" : "";
                message = p.TryGetProperty("title", out var t) ? t.GetString() ?? message : message;
            }
            else if (response.StatusCode == HttpStatusCode.TooManyRequests)
            {
                code = "too_many";
            }
            return new ApiResult<T>(default, status, code, message);
        }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException or IOException or UriFormatException
                                      or JsonException or NotSupportedException)
        {
            return new ApiResult<T>(default, 0, "offline", e.Message);
        }
    }

    public Task<ApiResult<TokenResponse>> LoginAsync(string username, string password, string device, CancellationToken ct = default) =>
        SendAsync<TokenResponse>(HttpMethod.Post, "auth/token", new { username, password, device }, ct);

    public Task<ApiResult<TokenResponse>> ChangePasswordAsync(string oldPassword, string newPassword, string confirmPassword,
                                                              CancellationToken ct = default) =>
        SendAsync<TokenResponse>(HttpMethod.Post, "auth/password", new { oldPassword, newPassword, confirmPassword }, ct);

    public Task<ApiResult<RepSnapshot>> SnapshotAsync(CancellationToken ct = default) =>
        SendAsync<RepSnapshot>(HttpMethod.Get, "rep/sync", null, ct);

    public Task<ApiResult<CreatedId>> SendOrderAsync(OrderRequest order, CancellationToken ct = default) =>
        SendAsync<CreatedId>(HttpMethod.Post, "rep/orders", order, ct);

    public Task<ApiResult<CreatedId>> PlanVisitsAsync(DateTime day, CancellationToken ct = default) =>
        SendAsync<CreatedId>(HttpMethod.Post, "rep/visits/plan", new { date = day.Date }, ct);

    public Task<ApiResult<CreatedId>> AddVisitAsync(int customerId, DateTime day, CancellationToken ct = default) =>
        SendAsync<CreatedId>(HttpMethod.Post, "rep/visits", new { customerID = customerId, date = day.Date }, ct);

    public Task<ApiResult<CreatedId>> RecordVisitAsync(int visitId, string? result, string? notes, CancellationToken ct = default) =>
        SendAsync<CreatedId>(HttpMethod.Put, "rep/visits/" + visitId, new { result, notes }, ct);

    public Task<ApiResult<List<RepVisit>>> VisitsAsync(DateTime day, CancellationToken ct = default) =>
        SendAsync<List<RepVisit>>(HttpMethod.Get, "rep/visits?date=" + day.ToString("yyyy-MM-dd"), null, ct);
}
