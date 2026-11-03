using RetailStore.Domain.Common;

namespace RetailStore.Domain.Sales;

/// <summary>A line as typed: quantity, unit price (VAT included when the prices include VAT), line discount on the same
/// basis, VAT rate.</summary>
public sealed record CalcLineIn(decimal Quantity, decimal Price, decimal Discount, decimal Rate);

/// <summary>A line as stored (without VAT): UnitPrice (4 decimals), Discount, NetAmount, Tax, LineTotal = NetAmount + Tax.</summary>
public sealed record CalcLineOut(decimal Quantity, decimal Rate, decimal UnitPrice, decimal Discount, decimal Net, decimal Tax, decimal Total);

public sealed record CalcResult(string Problem, IReadOnlyList<CalcLineOut> Lines, decimal Gross, decimal SubTotal, decimal Discount,
                                decimal Taxable, decimal Tax, decimal Total)
{
    public bool Ok => Problem.Length == 0;
}

public sealed record Settlement(string Problem, decimal Paid, decimal Remaining, decimal Change);

/// <summary>The invoice engine of modSales (CalcRun, Settle, ReturnAmounts, WeightedAverage), the same algorithm as
/// tools/pricing.py: the tests replay its cases (Vectors.Pricing).</summary>
public static class InvoiceCalculator
{
    private static decimal R2(decimal v) => MoneyMath.RoundMoney(v);
    private static decimal R4(decimal v) => MoneyMath.RoundMoney(v, 4);

    /// <summary>modSales.CalcRun: the problem in Arabic, or the lines and totals.</summary>
    public static CalcResult Calc(IReadOnlyList<CalcLineIn> lines, decimal invoiceDiscount, bool pricesIncludeVat)
    {
        CalcResult Fail(string message) => new(message, [], 0, 0, 0, 0, 0, 0);
        if (lines.Count == 0)
        {
            return Fail("لا توجد أصناف في الفاتورة.");
        }
        var gross = new decimal[lines.Count];
        decimal totalGross = 0;
        for (int i = 0; i < lines.Count; i++)
        {
            var l = lines[i];
            if (l.Quantity <= 0)
            {
                return Fail("الكمية يجب أن تكون أكبر من صفر (السطر " + (i + 1) + ").");
            }
            if (l.Price < 0 || l.Discount < 0)
            {
                return Fail("السعر والخصم لا يمكن أن يكونا سالبين (السطر " + (i + 1) + ").");
            }
            gross[i] = R2(l.Quantity * l.Price) - l.Discount;
            if (gross[i] < 0)
            {
                return Fail("خصم السطر " + (i + 1) + " أكبر من قيمته.");
            }
            totalGross += gross[i];
        }
        if (invoiceDiscount < 0 || invoiceDiscount > totalGross)
        {
            return Fail("خصم الفاتورة أكبر من قيمتها.");
        }
        // the invoice discount in proportion; the remainder goes to the largest line (the first of equals)
        var share = new decimal[lines.Count];
        if (invoiceDiscount > 0)
        {
            int biggest = 0;
            decimal allocated = 0;
            for (int i = 0; i < lines.Count; i++)
            {
                share[i] = R2(invoiceDiscount * gross[i] / totalGross);
                allocated += share[i];
                if (gross[i] > gross[biggest])
                {
                    biggest = i;
                }
            }
            share[biggest] += invoiceDiscount - allocated;
        }
        var outs = new List<CalcLineOut>();
        for (int i = 0; i < lines.Count; i++)
        {
            var l = lines[i];
            decimal after = gross[i] - share[i], total, tax, net, unit;
            if (pricesIncludeVat)
            {
                total = after;
                tax = R2(after * l.Rate / (1 + l.Rate));
                net = total - tax;
                unit = R4(l.Price / (1 + l.Rate));
            }
            else
            {
                net = after;
                tax = R2(after * l.Rate);
                total = net + tax;
                unit = R4(l.Price);
            }
            decimal disc = R2(l.Quantity * unit) - net;
            if (disc < 0)
            {
                // rounding pushed the net above qty x price: lift the unit price instead
                unit = R4(net / l.Quantity);
                disc = Math.Max(0, R2(l.Quantity * unit) - net);
            }
            outs.Add(new CalcLineOut(l.Quantity, l.Rate, unit, disc, net, tax, total));
        }
        return new CalcResult("", outs, totalGross, outs.Sum(x => x.Net + x.Discount), outs.Sum(x => x.Discount),
                              outs.Sum(x => x.Net), outs.Sum(x => x.Tax), outs.Sum(x => x.Total));
    }

    /// <summary>modSales.Settle: a cash sale is paid in full; a credit sale may be partly paid.</summary>
    public static Settlement Settle(decimal total, decimal tendered, bool isCredit)
    {
        if (tendered < 0)
        {
            tendered = 0;
        }
        if (!isCredit && tendered < total)
        {
            return new Settlement("المبلغ المدفوع (" + Money(tendered) + ") أقل من إجمالي الفاتورة (" + Money(total) + ").", 0, 0, 0);
        }
        decimal paid = Math.Min(tendered, total);
        return new Settlement("", paid, total - paid, tendered > total ? tendered - total : 0);
    }

    /// <summary>modSales.ReturnAmounts: (net, tax, total) of a return of qty from a sold line; the return that completes
    /// the line takes what is left.</summary>
    public static (string Problem, decimal Net, decimal Tax, decimal Total) ReturnAmounts(
        decimal soldQty, decimal soldTotal, decimal soldTax, decimal prevQty, decimal prevTotal, decimal prevTax, decimal qty)
    {
        if (qty <= 0 || prevQty + qty > soldQty)
        {
            return ("الكمية المرتجعة (" + Plain(qty) + ") أكبر من المتبقي (" + Plain(soldQty - prevQty) + ").", 0, 0, 0);
        }
        decimal total, tax;
        if (prevQty + qty == soldQty)
        {
            total = soldTotal - prevTotal;
            tax = soldTax - prevTax;
        }
        else
        {
            total = R2(soldTotal * qty / soldQty);
            tax = R2(soldTax * qty / soldQty);
        }
        return ("", total - tax, tax, total);
    }

    /// <summary>modSales.WeightedAverage: the average cost after qty (signed) at unitCost.</summary>
    public static decimal WeightedAverage(decimal curQty, decimal curAvg, decimal qty, decimal unitCost)
    {
        decimal newQty = curQty + qty;
        if (qty > 0)
        {
            return curQty <= 0 ? unitCost : R4((curQty * curAvg + qty * unitCost) / newQty);
        }
        if (qty < 0 && newQty > 0 && curQty > 0)
        {
            decimal v = R4((curQty * curAvg + qty * unitCost) / newQty);
            if (v >= 0)
            {
                return v;
            }
        }
        return curAvg;
    }

    /// <summary>Format$(x, "#,##0.00").</summary>
    public static string Money(decimal v) => v.ToString("#,##0.00", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>A quantity as VBA prints a Currency: no trailing zeros.</summary>
    public static string Plain(decimal v) => v.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
}
