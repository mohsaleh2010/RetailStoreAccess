using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Offline;

/// <summary>What one sync did: Reached = the server answered; NeedsLogin = the token ended (log in again; nothing is
/// lost); Sent / Rejected = items of the outbox; Message = why it stopped.</summary>
public sealed record SyncReport(bool Reached, bool NeedsLogin, int Sent, int Rejected, int Waiting, string Message, DateTime? SnapshotAt);

/// <summary>The sync of the rep app (docs/78): first the outbox, in this order: the orders (each with its number on the
/// phone, so an order sent twice after a lost answer is added once), then the visits in the order they were recorded
/// (a plan or a visit outside the route before its result, and the result after the orders so the result read from the
/// customer's documents sees the order); then the snapshot of the server is read again. A refusal of the program
/// (422, 404, 403) stops only that item (REJECTED with its message); no network, a server error or a token that ended
/// stop the sync and keep the item PENDING for the next one.</summary>
public sealed class SyncService(LocalStore store, RepApiClient api)
{
    public async Task<SyncReport> SyncAsync(Func<DateTime> now, CancellationToken ct = default)
    {
        int sent = 0, rejected = 0;
        var pending = store.Outbox(OutboxKinds.Pending);
        var ordered = pending.Where(i => i.Kind == OutboxKinds.Order).Concat(pending.Where(i => i.Kind != OutboxKinds.Order));
        var visitCache = new Dictionary<DateTime, List<RepVisit>>();
        foreach (var item in ordered)
        {
            var (status, message, serverId) = await SendAsync(item, visitCache, ct);
            if (status is >= 200 and < 300)
            {
                store.MarkSent(item.Id, serverId);
                sent++;
                continue;
            }
            if (status == 0 || status >= 500 || status == 429 || status == 401)
            {
                int waiting = store.Outbox(OutboxKinds.Pending).Count;
                return new SyncReport(status != 0, status == 401, sent, rejected, waiting, message, null);
            }
            store.MarkRejected(item.Id, message);
            rejected++;
        }
        var snapshot = await api.SnapshotAsync(ct);
        if (!snapshot.Ok || snapshot.Value is null)
        {
            return new SyncReport(snapshot.Status != 0, snapshot.NeedsLogin, sent, rejected, 0, snapshot.Message, null);
        }
        var at = now();
        store.SaveSnapshot(snapshot.Value, at);
        store.ForgetSent(snapshot.Value.Orders.Where(o => o.ClientRef is not null).Select(o => o.ClientRef!), at.AddDays(-30));
        return new SyncReport(true, false, sent, rejected, 0, "", at);
    }

    private async Task<(int Status, string Message, int? ServerId)> SendAsync(
        OutboxItem item, Dictionary<DateTime, List<RepVisit>> visitCache, CancellationToken ct)
    {
        switch (item.Kind)
        {
            case OutboxKinds.Order:
            {
                var r = await api.SendOrderAsync(item.Read<PendingOrder>().ToRequest(), ct);
                return (r.Status, r.Message, r.Value?.Id);
            }
            case OutboxKinds.VisitPlan:
            {
                var day = item.Read<VisitDay>().Date;
                var r = await api.PlanVisitsAsync(day, ct);
                visitCache.Remove(day.Date);
                // no route stops that day is not an error of the rep: the plan simply adds nothing
                return r.Status == 422 ? (200, "", 0) : (r.Status, r.Message, r.Value?.Id);
            }
            case OutboxKinds.VisitAdd:
            {
                var v = item.Read<VisitAdd>();
                var r = await api.AddVisitAsync(v.CustomerID, v.Date, ct);
                visitCache.Remove(v.Date.Date);
                if (r.Status == 422)
                {
                    // already a visit of that customer that day (sent before, or planned): it is the same visit
                    var found = await FindVisitAsync(v.CustomerID, v.Date, visitCache, ct);
                    if (found.Status != 200)
                    {
                        return (found.Status, found.Message, null);
                    }
                    if (found.Id > 0)
                    {
                        return (200, "", found.Id);
                    }
                }
                return (r.Status, r.Message, r.Value?.Id);
            }
            case OutboxKinds.VisitResult:
            {
                var v = item.Read<VisitResultEntry>();
                int id = v.VisitID;
                if (id == 0)
                {
                    var found = await FindVisitAsync(v.CustomerID, v.Date, visitCache, ct);
                    if (found.Status != 200)
                    {
                        return (found.Status, found.Message, null);
                    }
                    if (found.Id == 0)
                    {
                        return (404, "اختر الزيارة.", null);
                    }
                    id = found.Id;
                }
                var r = await api.RecordVisitAsync(id, v.Result, v.Notes, ct);
                return (r.Status, r.Message, id);
            }
            default:
                return (422, item.Kind, null);
        }
    }

    private async Task<(int Status, string Message, int Id)> FindVisitAsync(int customerId, DateTime day,
                                                                             Dictionary<DateTime, List<RepVisit>> cache, CancellationToken ct)
    {
        if (!cache.TryGetValue(day.Date, out var visits))
        {
            var r = await api.VisitsAsync(day, ct);
            if (!r.Ok || r.Value is null)
            {
                return (r.Status, r.Message, 0);
            }
            cache[day.Date] = visits = r.Value;
        }
        return (200, "", visits.FirstOrDefault(x => x.CustomerID == customerId)?.RepVisitID ?? 0);
    }
}
