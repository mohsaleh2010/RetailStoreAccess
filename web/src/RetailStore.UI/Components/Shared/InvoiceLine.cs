using RetailStore.Domain.Sales;

namespace RetailStore.UI.Components.Shared;

/// <summary>A line of InvoiceLinesEditor as typed: the product, quantity, unit price (VAT included when the prices of
/// the document include it), line discount and VAT rate (a fraction, 0.15). The amounts come from the invoice engine
/// of the program (InvoiceCalculator = modSales.CalcRun), never from the screen; the posting still recalculates
/// everything in its service.</summary>
public sealed class InvoiceLine<TProduct>
{
    public TProduct? Product { get; set; }
    public decimal Quantity { get; set; } = 1;
    public decimal Price { get; set; }
    public decimal Discount { get; set; }
    public decimal VatRate { get; set; }

    /// <summary>The line as the engine computed it (net, VAT, total), null while the document has a problem.</summary>
    public CalcLineOut? Amounts { get; internal set; }

    public bool IsEmpty => Product is null;
}

public static class InvoiceLines
{
    /// <summary>Calc of the lines that have a product (the empty line at the end is not a line yet); the amounts of
    /// each line are set on it. No line at all = no problem and no totals.</summary>
    public static CalcResult Calculate<TProduct>(IReadOnlyList<InvoiceLine<TProduct>> lines, decimal invoiceDiscount, bool pricesIncludeVat)
    {
        var filled = lines.Where(l => !l.IsEmpty).ToList();
        foreach (var line in lines)
        {
            line.Amounts = null;
        }
        if (filled.Count == 0)
        {
            return new CalcResult("", [], 0, 0, 0, 0, 0, 0);
        }
        var result = InvoiceCalculator.Calc(filled.Select(l => new CalcLineIn(l.Quantity, l.Price, l.Discount, l.VatRate)).ToList(),
                                            invoiceDiscount, pricesIncludeVat);
        if (result.Ok)
        {
            for (int i = 0; i < filled.Count; i++)
            {
                filled[i].Amounts = result.Lines[i];
            }
        }
        return result;
    }
}
