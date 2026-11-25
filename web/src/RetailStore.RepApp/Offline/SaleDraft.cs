using RetailStore.RepApp.Client;

namespace RetailStore.RepApp.Offline;

/// <summary>A sales invoice being made on the phone (docs/87): the lines at the price the program gives the customer
/// (RepPricing, the price of a new line in Access), cash or credit. It is sent to the program at once; ClientRef = its number
/// on the phone, kept until the program answers so a lost answer never makes two invoices, and new after any change.</summary>
public sealed class SaleDraft(RepSnapshot snapshot, RepCustomer customer)
{
    public RepCustomer Customer { get; } = customer;
    public List<DraftLine> Lines { get; } = [];
    public string PaymentType { get; private set; } = "CASH";
    public string? Notes { get; set; }
    public string ClientRef { get; private set; } = Guid.NewGuid().ToString("N");

    public decimal Total => OrderDraft.RoundMoney(Lines.Sum(l => l.Quantity * l.UnitPrice));

    public decimal PriceOf(RepProduct product) => RepPricing.PriceOf(snapshot, Customer.CustomerID, product);

    private void Changed() => ClientRef = Guid.NewGuid().ToString("N");

    public void SetPaymentType(string type)
    {
        PaymentType = type == "CREDIT" ? "CREDIT" : "CASH";
        Changed();
    }

    /// <summary>Add a product (or more of one already in the invoice): "" or the reason.</summary>
    public string Add(int productId, decimal quantity)
    {
        var product = snapshot.Products.FirstOrDefault(p => p.ProductID == productId);
        if (product is null)
        {
            return "اختر الصنف.";
        }
        if (quantity <= 0)
        {
            return "الكمية أكبر من صفر.";
        }
        var line = Lines.FirstOrDefault(l => l.Product.ProductID == productId);
        if (line is null)
        {
            Lines.Add(new DraftLine(product, quantity, PriceOf(product)));
        }
        else
        {
            line.Quantity += quantity;
        }
        Changed();
        return "";
    }

    public string SetQuantity(int productId, decimal quantity)
    {
        var line = Lines.FirstOrDefault(l => l.Product.ProductID == productId);
        if (line is null)
        {
            return "اختر الصنف.";
        }
        if (quantity <= 0)
        {
            return "الكمية أكبر من صفر.";
        }
        line.Quantity = quantity;
        Changed();
        return "";
    }

    public void Remove(int productId)
    {
        Lines.RemoveAll(l => l.Product.ProductID == productId);
        Changed();
    }

    /// <summary>The invoice to send: "" and the request, or the reason.</summary>
    public (string Problem, RepSaleRequest? Request) ToRequest(bool confirmStock)
    {
        if (Lines.Count == 0)
        {
            return ("أضف أصناف الفاتورة.", null);
        }
        string text = (Notes ?? "").Trim();
        return ("", new RepSaleRequest(ClientRef, Customer.CustomerID, PaymentType, PaymentType == "CASH" ? 1 : null, 0,
                                       text.Length == 0 ? null : (text.Length <= 255 ? text : text[..255]),
                                       Lines.Select(l => new CartLine(l.Product.ProductID, l.Quantity, l.UnitPrice)).ToArray(), confirmStock));
    }
}
