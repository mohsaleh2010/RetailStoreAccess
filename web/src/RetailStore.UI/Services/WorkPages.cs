namespace RetailStore.UI.Services;

/// <summary>A screen of the web that is not a data screen: its route, the Access form whose permission opens it, its
/// caption and subtitle.</summary>
public sealed record WorkPage(string Route, string FormName, string Caption, string Subtitle);

/// <summary>The work screens of the web (plan 6C, docs/79; plan 6D, docs/80), in the menu after the data screens.</summary>
public static class WorkPages
{
    public static readonly IReadOnlyList<WorkPage> All =
    [
        new("pos", "frmPOS", "نقطة البيع", "فاتورة بيع: الأصناف بالباركود أو الاسم، والعميل، والدفع"),
        new("sales", "frmSalesInvoice", "فواتير المبيعات", "الفواتير بتاريخها وطباعتها مع رمز الفاتورة الإلكترونية"),
        new("payments", "frmCustomerPayment", "سندات القبض", "تحصيل من العملاء نقدًا أو بالبطاقة أو التحويل"),
        new("purchase", "frmPurchaseInvoice", "فاتورة شراء", "الأصناف بسعر الشراء ووحدتها، والمورد، والدفع، والمخزن"),
        new("purchases", "frmPurchaseView", "فواتير المشتريات", "فواتير الموردين بتاريخها وأسطرها"),
        new("supplier-payments", "frmSupplierPayment", "سندات الصرف", "الدفع للموردين نقدًا أو بالتحويل"),
        new("stock", "frmInventory", "المخزون", "رصيد كل صنف في كل مخزن، وكارت الصنف"),
        new("transfers", "frmStockTransfer", "التحويل بين المخازن", "إرسال الأصناف من مخزن إلى آخر واستلامها"),
    ];

    /// <summary>The name of a ZatcaStatus, as EInvoiceReportQuery shows it (tools/queries.py).</summary>
    public static string StatusName(string status) => status switch
    {
        "PENDING" => "بانتظار الإرسال",
        "REPORTED" => "مُبلَّغ",
        "CLEARED" => "معتمد",
        "WARNING" => "مقبول مع تحذير",
        "REJECTED" => "مرفوض",
        "SUBMITTED" => "مُرسل",
        "VALID" => "صالح",
        "INVALID" => "غير صالح",
        "CANCELLED" => "ملغى",
        _ => "لا يُرسل",
    };

    public static string PaymentTypeName(string type) => type == "CREDIT" ? "آجل" : "نقدي";

    public static string SubTypeName(string subType) => subType == "STANDARD" ? "فاتورة ضريبية" : "فاتورة ضريبية مبسطة";

    public const string NoProduct = "لا يوجد صنف بهذا الرمز أو الاسم.";
    public const string ChooseCustomer = "اختر العميل.";
    public const string InvoiceSaved = "تم حفظ الفاتورة ";
    public const string ChangeDue = " - الباقي للعميل ";
    public const string PaymentSaved = "تم حفظ سند القبض ";
    public const string ChooseSupplier = "اختر المورد.";
    public const string PurchaseSaved = "تم حفظ فاتورة الشراء ";
    public const string SupplierPaymentSaved = "تم حفظ سند الصرف ";
    public const string TransferSaved = "تم حفظ التحويل ";
    public const string TransferDone = "تمت العملية على التحويل ";

    /// <summary>The name of a transfer status (frmStockTransfer).</summary>
    public static string TransferStatusName(string status) => status switch
    {
        "DRAFT" => "مسودة",
        "SENT" => "بالطريق",
        _ => "مستلم",
    };

    /// <summary>Format$(x, "#,##0.00").</summary>
    public static string Money(decimal v) => v.ToString("#,##0.00", System.Globalization.CultureInfo.InvariantCulture);

    public static string Qty(decimal v) => v.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
