using RetailStore.PosApp.Connection;

namespace RetailStore.PosApp.Offline;

/// <summary>What one sync did: Reached = the data answered; Sent / Rejected = sales of the PC; Waiting = sales still to
/// send (the data was not reached, or their cashier must log in again); Message = why it stopped.</summary>
public sealed record PosSyncReport(bool Reached, int Sent, int Rejected, int Waiting, string Message, DateTime? SnapshotAt);

/// <summary>The sync of the POS (docs/83): the sales of the PC in the order they were made, each with the credential of
/// the cashier who made it (his API token when the data is remote). A sale already on the server (sent before, its answer
/// lost) answers with its invoice and is added once. A refusal of the program (422, 409, 404, 403) stops only that sale
/// (REJECTED with its message); the data not reached or a server error stop the sync and keep the sale PENDING; a token
/// that ended (401) keeps the sales of that cashier until he logs in again.</summary>
public static class PosSync
{
    public static async Task<(int Sent, int Rejected, int Status, string Message)> SendSalesAsync(
        PosStore store, IPosBackend backend, Func<int, Task<PosCredential?>> credentialOf, Func<int, Task> tokenEnded, bool english,
        Func<DateTime> now, CancellationToken ct = default)
    {
        int sent = 0, rejected = 0;
        var blocked = new HashSet<int>();
        foreach (var sale in store.Sales(SaleStates.Pending))
        {
            if (blocked.Contains(sale.EmployeeID))
            {
                continue;
            }
            var who = await credentialOf(sale.EmployeeID);
            if (who is null)
            {
                blocked.Add(sale.EmployeeID);
                continue;
            }
            var r = await backend.SendSaleAsync(who, sale.Request, english, ct);
            if (r.Ok && r.Value is { } v)
            {
                store.MarkSent(sale.Id, v.InvoiceID, now());
                sent++;
                continue;
            }
            if (r.Retry)
            {
                return (sent, rejected, r.Status, r.Message);
            }
            if (r.Status == 401)
            {
                await tokenEnded(sale.EmployeeID);
                blocked.Add(sale.EmployeeID);
                continue;
            }
            store.MarkRejected(sale.Id, r.Message);
            rejected++;
        }
        return (sent, rejected, 200, "");
    }
}
