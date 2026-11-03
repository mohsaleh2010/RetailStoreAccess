using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Offline;

/// <summary>A visit as the phone shows it: from the server, or made on the phone and not sent yet (Waiting).</summary>
public sealed record BoardVisit(int RepVisitID, int CustomerID, string CustomerName, bool IsPlanned, int StopOrder, string Result,
                                string? Notes, bool Waiting);

/// <summary>The visits of a day on the phone (docs/78): the server's visits of that day, then what the outbox adds to
/// them: the plan of the day from the route (modRepRoutes.CreateVisitPlan: the route stops of the weekday, the customers
/// already in the day kept), visits outside the route, and the results recorded offline.</summary>
public static class VisitBoard
{
    /// <summary>modRepRoutes.VisitWeekdayOf: VBA Weekday(d, vbSunday), Sunday = 1 ... Saturday = 7.</summary>
    public static int WeekdayOf(DateTime d) => (int)d.DayOfWeek + 1;

    public static List<BoardVisit> Build(RepSnapshot? snapshot, IEnumerable<OutboxItem> outbox, DateTime day)
    {
        var d = day.Date;
        var list = new List<BoardVisit>();
        if (snapshot is null)
        {
            return list;
        }
        list.AddRange(snapshot.Visits.Where(v => v.VisitDate.Date == d)
                          .Select(v => new BoardVisit(v.RepVisitID, v.CustomerID, v.CustomerName, v.IsPlanned, v.StopOrder,
                                                      v.Result, v.Notes, false)));
        string Name(int id) => snapshot.Customers.FirstOrDefault(c => c.CustomerID == id)?.CustomerName ?? id.ToString();
        foreach (var item in outbox.Where(i => i.State != OutboxKinds.Rejected))
        {
            bool waiting = item.State == OutboxKinds.Pending;
            switch (item.Kind)
            {
                case OutboxKinds.VisitPlan when item.Read<VisitDay>().Date.Date == d:
                    int order = 0;
                    foreach (var c in snapshot.Customers.Where(c => c.RouteWeekdays.Contains(WeekdayOf(d))))
                    {
                        order++;
                        if (list.All(v => v.CustomerID != c.CustomerID))
                        {
                            list.Add(new BoardVisit(0, c.CustomerID, c.CustomerName, true, order, "PENDING", null, waiting));
                        }
                    }
                    break;
                case OutboxKinds.VisitAdd:
                    var add = item.Read<VisitAdd>();
                    if (add.Date.Date == d && list.All(v => v.CustomerID != add.CustomerID))
                    {
                        int next = list.Count == 0 ? 1 : list.Max(v => v.StopOrder) + 1;
                        list.Add(new BoardVisit(0, add.CustomerID, Name(add.CustomerID), false, next, "PENDING", null, waiting));
                    }
                    break;
                case OutboxKinds.VisitResult:
                    var r = item.Read<VisitResultEntry>();
                    int i = list.FindIndex(v => r.Date.Date == d && (r.VisitID != 0 ? v.RepVisitID == r.VisitID : v.CustomerID == r.CustomerID));
                    if (i >= 0)
                    {
                        list[i] = list[i] with
                        {
                            Result = r.Result.Length == 0 ? "AUTO" : r.Result, Notes = r.Notes ?? list[i].Notes,
                            Waiting = list[i].Waiting || waiting,
                        };
                    }
                    break;
            }
        }
        return list.OrderBy(v => v.StopOrder).ThenBy(v => v.CustomerName).ToList();
    }
}
