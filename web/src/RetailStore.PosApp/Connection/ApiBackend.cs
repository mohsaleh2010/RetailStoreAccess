using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using RetailStore.Data.Services;

namespace RetailStore.PosApp.Connection;

/// <summary>A remote server (docs/83): the API /api/v1 (docs/77) by its address, with the token of each user. The POS
/// never opens a database port over the internet.</summary>
public sealed class ApiBackend(HttpClient http, string serverUrl, bool ownsClient = true) : IPosBackend
{
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    private sealed record TokenAnswer(string AccessToken, int EmployeeID, string EmployeeName, bool MustChangePassword);

    private sealed record Posted(int Id, string Number);

    private Uri Address(string path)
    {
        string root = PosConnection.NormalizeUrl(serverUrl);
        if (root.Length == 0 && http.BaseAddress is not null)
        {
            root = http.BaseAddress.ToString().TrimEnd('/');
        }
        return new Uri(root + "/api/v1/" + path);
    }

    private async Task<BackendResult<T>> SendAsync<T>(HttpMethod method, string path, object? body, string? token, bool english,
                                                      CancellationToken ct)
    {
        try
        {
            using var request = new HttpRequestMessage(method, Address(path));
            if (token is { Length: > 0 })
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }
            request.Headers.AcceptLanguage.Add(new StringWithQualityHeaderValue(english ? "en" : "ar"));
            if (body is not null)
            {
                request.Content = JsonContent.Create(body, body.GetType(), options: Json);
            }
            using var response = await http.SendAsync(request, ct);
            int status = (int)response.StatusCode;
            if (response.IsSuccessStatusCode)
            {
                return new BackendResult<T>(await response.Content.ReadFromJsonAsync<T>(Json, ct), status, "", "");
            }
            string code = "", message = response.ReasonPhrase ?? status.ToString();
            if (response.Content.Headers.ContentType?.MediaType == "application/problem+json")
            {
                var p = await response.Content.ReadFromJsonAsync<JsonElement>(Json, ct);
                code = p.TryGetProperty("code", out var c) ? c.GetString() ?? "" : "";
                message = p.TryGetProperty("title", out var t) ? t.GetString() ?? message : message;
            }
            return new BackendResult<T>(default, status, code, message);
        }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException or IOException or UriFormatException
                                      or JsonException or NotSupportedException)
        {
            return BackendResult<T>.Unreachable(e.Message);
        }
    }

    /// <summary>The API answers a call without a token with its problem details (401): the server is the program's.</summary>
    public async Task<string> TestAsync(CancellationToken ct = default)
    {
        var r = await SendAsync<JsonElement>(HttpMethod.Get, "me", null, null, false, ct);
        if (r.Status == 401 && r.Code == "unauthorized")
        {
            return "";
        }
        return r.Status == 0 ? "تعذر الوصول إلى الخادم. تأكد من الإنترنت ومن عنوان الخادم." + " " + r.Message
             : "العنوان لا يرد بخادم البرنامج.";
    }

    public async Task<BackendResult<BackendLogin>> LoginAsync(string username, string password, string device, CancellationToken ct = default)
    {
        var r = await SendAsync<TokenAnswer>(HttpMethod.Post, "auth/token", new { username, password, device }, null, false, ct);
        return r.Ok && r.Value is { } t
            ? BackendResult<BackendLogin>.Of(new BackendLogin(t.EmployeeID, t.EmployeeName, t.MustChangePassword, t.AccessToken))
            : new BackendResult<BackendLogin>(null, r.Status, r.Code, r.Message);
    }

    public Task<BackendResult<PosSnapshot>> SnapshotAsync(PosCredential who, bool english, CancellationToken ct = default) =>
        SendAsync<PosSnapshot>(HttpMethod.Get, "pos/sync", null, who.Token, english, ct);

    public async Task<BackendResult<DeviceSaleResult>> SendSaleAsync(PosCredential who, DeviceSaleRequest sale, bool english,
                                                                     CancellationToken ct = default)
    {
        var r = await SendAsync<Posted>(HttpMethod.Post, "pos/sales", sale, who.Token, english, ct);
        return r.Ok && r.Value is { } p
            ? new BackendResult<DeviceSaleResult>(new DeviceSaleResult(SaveStatus.Saved, "", p.Id, p.Number, r.Status == 200), r.Status, "", "")
            : new BackendResult<DeviceSaleResult>(null, r.Status, r.Code, r.Message);
    }

    public Task<BackendResult<ReturnableInvoice>> ReturnableAsync(PosCredential who, string invoiceNumber, bool english,
                                                                 CancellationToken ct = default) =>
        SendAsync<ReturnableInvoice>(HttpMethod.Get, "pos/returns/" + Uri.EscapeDataString((invoiceNumber ?? "").Trim()), null, who.Token,
                                     english, ct);

    public Task<BackendResult<PosDocument>> SendReturnAsync(PosCredential who, DeviceReturnRequest request, bool english,
                                                           CancellationToken ct = default) =>
        SendAsync<PosDocument>(HttpMethod.Post, "pos/returns", request, who.Token, english, ct);

    public Task<BackendResult<PosDocument>> SendPaymentAsync(PosCredential who, DevicePaymentRequest request, bool english,
                                                            CancellationToken ct = default) =>
        SendAsync<PosDocument>(HttpMethod.Post, "pos/payments", request, who.Token, english, ct);

    public Task<BackendResult<PosClosingBoxes>> ClosingBoxesAsync(PosCredential who, bool english, CancellationToken ct = default) =>
        SendAsync<PosClosingBoxes>(HttpMethod.Get, "pos/closing", null, who.Token, english, ct);

    public Task<BackendResult<ClosingPreview>> ClosingPreviewAsync(PosCredential who, int boxId, bool english, CancellationToken ct = default) =>
        SendAsync<ClosingPreview>(HttpMethod.Get, "pos/closing/" + boxId, null, who.Token, english, ct);

    public Task<BackendResult<PosDocument>> SendClosingAsync(PosCredential who, DeviceClosingRequest request, bool english,
                                                            CancellationToken ct = default) =>
        SendAsync<PosDocument>(HttpMethod.Post, "pos/closing", request, who.Token, english, ct);

    public void Dispose()
    {
        if (ownsClient)
        {
            http.Dispose();
        }
    }
}
