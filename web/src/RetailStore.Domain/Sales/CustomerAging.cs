namespace RetailStore.Domain.Sales;

/// <summary>What a party owes, document by document (qryAgingDebits): an invoice or the opening balance.</summary>
public sealed record AgingDebit(int PartyID, string DocType, int DocID, string DocNo, DateTime DocDate, DateTime DueDate, decimal Amount);

/// <summary>What pays it (qryAgingCredits): a return (TargetID = its invoice), a payment, a cheque, a credit opening balance.</summary>
public sealed record AgingCredit(int PartyID, string CreditType, int CreditID, DateTime CreditDate, decimal Amount, int TargetID);

/// <summary>A payment linked to an invoice (qryAgingAllocations).</summary>
public sealed record AgingLink(int PaymentID, int InvoiceID, decimal Amount);

/// <summary>An open document; DocType CREDIT = a credit left to the party (negative amount).</summary>
public sealed record AgingItem(int PartyID, string DocType, int DocID, string DocNo, DateTime? DocDate, DateTime? DueDate, decimal OpenAmount);

/// <summary>modAging.ComputeAging / MaxDaysLate (tools/sim.py Store.aging): a return pays its own invoice, a payment the
/// invoices it is linked to; everything else pays the oldest due documents first.</summary>
public static class CustomerAging
{
    public static List<AgingItem> Compute(IEnumerable<AgingDebit> debits, IEnumerable<AgingCredit> credits, IEnumerable<AgingLink> links,
                                          DateTime asOf)
    {
        var next = asOf.Date.AddDays(1);
        var docs = debits.Where(d => d.DocDate < next)
            .OrderBy(d => d.PartyID).ThenBy(d => d.DueDate.Date).ThenBy(d => d.DocDate)
            .ThenBy(d => d.DocType == "OPENING" ? 0 : 1).ThenBy(d => d.DocID).ToList();
        var open = new Dictionary<(int, string, int), decimal>();
        foreach (var d in docs)
        {
            open[(d.PartyID, d.DocType, d.DocID)] = d.Amount;
        }
        var byPayment = links.OrderBy(l => l.PaymentID).ThenBy(l => l.InvoiceID).GroupBy(l => l.PaymentID)
            .ToDictionary(g => g.Key, g => g.ToList());
        var pool = new Dictionary<int, decimal>();

        decimal ApplyTo(int party, int invoice, decimal amount)
        {
            var key = (party, "INVOICE", invoice);
            if (open.TryGetValue(key, out var o))
            {
                decimal part = Math.Min(o, amount);
                open[key] = o - part;
                amount -= part;
            }
            return amount;
        }

        // 1. returns on their invoice, payments on their linked invoices
        foreach (var c in credits.Where(c => c.CreditDate < next)
                     .OrderBy(c => c.CreditDate).ThenBy(c => c.CreditType, StringComparer.Ordinal).ThenBy(c => c.CreditID))
        {
            decimal rest = c.Amount;
            if (c.CreditType == "RETURN")
            {
                rest = ApplyTo(c.PartyID, c.TargetID, rest);
            }
            else if (c.CreditType == "PAYMENT" && byPayment.TryGetValue(c.CreditID, out var mine))
            {
                foreach (var l in mine)
                {
                    decimal part = Math.Min(l.Amount, rest);
                    rest = rest - part + ApplyTo(c.PartyID, l.InvoiceID, part);
                }
            }
            pool[c.PartyID] = pool.GetValueOrDefault(c.PartyID) + rest;
        }
        // 2. the rest pays the oldest due documents first
        var result = new List<AgingItem>();
        foreach (var d in docs)
        {
            var key = (d.PartyID, d.DocType, d.DocID);
            decimal part = Math.Min(pool.GetValueOrDefault(d.PartyID), open[key]);
            if (part > 0)
            {
                open[key] -= part;
                pool[d.PartyID] -= part;
            }
            if (open[key] != 0)
            {
                result.Add(new AgingItem(d.PartyID, d.DocType, d.DocID, d.DocNo, d.DocDate, d.DueDate.Date, open[key]));
            }
        }
        foreach (var (party, rest) in pool.Where(p => p.Value != 0))
        {
            result.Add(new AgingItem(party, "CREDIT", 0, "رصيد دائن غير مستخدم", null, null, -rest));
        }
        return result;
    }

    /// <summary>modAging.MaxDaysLate: the days of the oldest late document still open.</summary>
    public static int MaxDaysLate(IEnumerable<AgingItem> items, DateTime asOf) =>
        items.Where(i => i.DocType != "CREDIT" && i.OpenAmount > 0 && i.DueDate is not null)
             .Select(i => (int)(asOf.Date - i.DueDate!.Value.Date).TotalDays).DefaultIfEmpty(0).Max() is var m && m > 0 ? m : 0;
}
