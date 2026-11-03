using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Offline;

/// <summary>The price of an order line on the phone (docs/78): the server sends the price without a customer of every
/// product and, for each price group (the rep of the order + the customer's price list), the prices that differ; so the
/// phone gives the same price as modSalesOrders.OrderDefaultPrice without the internet.</summary>
public static class RepPricing
{
    public static decimal PriceOf(RepSnapshot snapshot, int customerId, RepProduct product)
    {
        if (snapshot.CustomerPriceGroups.TryGetValue(customerId, out var group)
            && snapshot.PriceGroups.TryGetValue(group, out var prices)
            && prices.TryGetValue(product.ProductID, out var price))
        {
            return price;
        }
        return product.Price;
    }
}

public sealed class DraftLine(RepProduct product, decimal quantity, decimal unitPrice)
{
    public RepProduct Product { get; } = product;
    public decimal Quantity { get; set; } = quantity;
    public decimal UnitPrice { get; } = unitPrice;
    public decimal Amount => OrderDraft.RoundMoney(Quantity * UnitPrice);
}

/// <summary>A customer order being taken on the phone: the same checks as a line of frmSalesOrder
/// (modSalesOrders.SalesOrderLineProblem) with the same messages, so the rep hears the refusal at once, not at the sync.</summary>
public sealed class OrderDraft(RepSnapshot snapshot, RepCustomer customer)
{
    public RepCustomer Customer { get; } = customer;
    public List<DraftLine> Lines { get; } = [];
    public string? Notes { get; set; }
    public DateTime? RequiredDate { get; set; }

    /// <summary>modCommon.RoundMoney: half away from zero.</summary>
    public static decimal RoundMoney(decimal value) => Math.Round(value, 2, MidpointRounding.AwayFromZero);

    public decimal Total => RoundMoney(Lines.Sum(l => l.Quantity * l.UnitPrice));

    public decimal PriceOf(RepProduct product) => RepPricing.PriceOf(snapshot, Customer.CustomerID, product);

    /// <summary>modSalesOrders.SalesOrderLineProblem.</summary>
    public static string LineProblem(RepProduct? product, decimal quantity)
    {
        if (product is null)
        {
            return "اختر الصنف.";
        }
        if (!product.TrackStock)
        {
            return "هذا الصنف لا يتابع المخزون، فلا يُحجز ولا يُسلَّم بإذن.";
        }
        return quantity <= 0 ? "الكمية أكبر من صفر." : "";
    }

    /// <summary>Add a product (or more of a product already in the order): "" or the reason it is refused.</summary>
    public string Add(int productId, decimal quantity)
    {
        var product = snapshot.Products.FirstOrDefault(p => p.ProductID == productId);
        string problem = LineProblem(product, quantity);
        if (problem.Length > 0)
        {
            return problem;
        }
        var line = Lines.FirstOrDefault(l => l.Product.ProductID == productId);
        if (line is null)
        {
            Lines.Add(new DraftLine(product!, quantity, PriceOf(product!)));
        }
        else
        {
            line.Quantity += quantity;
        }
        return "";
    }

    public string SetQuantity(int productId, decimal quantity)
    {
        var line = Lines.FirstOrDefault(l => l.Product.ProductID == productId);
        if (line is null)
        {
            return "اختر الصنف.";
        }
        string problem = LineProblem(line.Product, quantity);
        if (problem.Length == 0)
        {
            line.Quantity = quantity;
        }
        return problem;
    }

    public void Remove(int productId) => Lines.RemoveAll(l => l.Product.ProductID == productId);

    /// <summary>The order to keep in the outbox: "" and the order, or the reason it cannot be saved.</summary>
    public (string Problem, PendingOrder? Order) ToPending(DateTime takenAt)
    {
        if (Lines.Count == 0)
        {
            return ("أضف أسطر الطلبية.", null);
        }
        string text = (Notes ?? "").Trim();
        var order = new PendingOrder(Guid.NewGuid().ToString("N"), Customer.CustomerID, Customer.CustomerName, takenAt,
                                     RequiredDate?.Date, text.Length == 0 ? null : (text.Length <= 255 ? text : text[..255]),
                                     Lines.Select(l => new PendingOrderLine(l.Product.ProductID, l.Product.ProductName, l.Quantity,
                                                                            l.UnitPrice)).ToArray());
        return ("", order);
    }
}
