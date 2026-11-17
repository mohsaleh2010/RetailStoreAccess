using RetailStore.Data.Services;

namespace RetailStore.PosApp.Connection;

/// <summary>The answer of the data: the value, or a status like the API's (docs/77): 0 = not reached (no network, the
/// server or the database is down), 401 = log in again, 409 = a question, 422 / 404 / 403 = refused by the program.</summary>
public sealed record BackendResult<T>(T? Value, int Status, string Code, string Message)
{
    public bool Ok => Status is >= 200 and < 300;

    /// <summary>Try again later: not reached, or the server failed, or too many logins.</summary>
    public bool Retry => Status == 0 || Status >= 500 || Status == 429;

    public static BackendResult<T> Of(T value) => new(value, 200, "", "");

    public static BackendResult<T> Unreachable(string message) => new(default, 0, "offline", message);

    public static BackendResult<T> Refused(string message, int status = 422, string code = "refused") => new(default, status, code, message);
}

/// <summary>A login on the data: Token = the API token (remote) or null (local).</summary>
public sealed record BackendLogin(int EmployeeID, string EmployeeName, bool MustChangePassword, string? Token);

/// <summary>Who sends: the user, and his API token when the data is remote.</summary>
public sealed record PosCredential(int EmployeeID, string? Token);

/// <summary>The way of the POS to the program's data (docs/83): directly (a local database) or through the API (a remote
/// server). Both answer with the same records and the same Arabic messages, because both run PosService.</summary>
public interface IPosBackend : IDisposable
{
    /// <summary>"" when the data answers, else why not.</summary>
    Task<string> TestAsync(CancellationToken ct = default);

    Task<BackendResult<BackendLogin>> LoginAsync(string username, string password, string device, CancellationToken ct = default);

    Task<BackendResult<PosSnapshot>> SnapshotAsync(PosCredential who, bool english, CancellationToken ct = default);

    /// <summary>Send a sale made on the POS: the invoice id and number (Duplicate = it was already there).</summary>
    Task<BackendResult<DeviceSaleResult>> SendSaleAsync(PosCredential who, DeviceSaleRequest sale, bool english, CancellationToken ct = default);

    // ---- docs/86: made on the program at once (they need the data), each sent once by its number on the device

    /// <summary>The invoice of a number with its lines to return from (404 = no such invoice).</summary>
    Task<BackendResult<ReturnableInvoice>> ReturnableAsync(PosCredential who, string invoiceNumber, bool english, CancellationToken ct = default);

    Task<BackendResult<PosDocument>> SendReturnAsync(PosCredential who, DeviceReturnRequest request, bool english, CancellationToken ct = default);

    /// <summary>A customer receipt (409 = the question on an amount over the balance).</summary>
    Task<BackendResult<PosDocument>> SendPaymentAsync(PosCredential who, DevicePaymentRequest request, bool english, CancellationToken ct = default);

    Task<BackendResult<PosClosingBoxes>> ClosingBoxesAsync(PosCredential who, bool english, CancellationToken ct = default);

    Task<BackendResult<ClosingPreview>> ClosingPreviewAsync(PosCredential who, int boxId, bool english, CancellationToken ct = default);

    Task<BackendResult<PosDocument>> SendClosingAsync(PosCredential who, DeviceClosingRequest request, bool english, CancellationToken ct = default);
}
