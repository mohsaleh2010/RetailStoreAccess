using Microsoft.Extensions.Localization;
using MudBlazor;

namespace RetailStore.UI.Services;

/// <summary>The texts MudBlazor writes itself (the pager, the column menu and filters of MudDataGrid, the date pickers,
/// "Required") in Arabic when the interface is Arabic; in English MudBlazor keeps its own texts.</summary>
public sealed class ErpMudLocalizer(UiLanguage lang) : MudLocalizer
{
    /// <summary>The MudBlazor resource keys (LanguageResource) and their Arabic.</summary>
    public static readonly IReadOnlyDictionary<string, string> Arabic = new Dictionary<string, string>
    {
        ["MudDataGridPager_AllItems"] = "الكل",
        ["MudDataGridPager_FirstPage"] = "الصفحة الأولى",
        ["MudDataGridPager_InfoFormat"] = "{0}-{1} من {2}",
        ["MudDataGridPager_LastPage"] = "الصفحة الأخيرة",
        ["MudDataGridPager_NextPage"] = "الصفحة التالية",
        ["MudDataGridPager_PreviousPage"] = "الصفحة السابقة",
        ["MudDataGridPager_RowsPerPage"] = "صفوف في الصفحة",
        ["MudDataGrid_AddFilter"] = "إضافة فلتر",
        ["MudDataGrid_Apply"] = "تطبيق",
        ["MudDataGrid_Cancel"] = "إلغاء",
        ["MudDataGrid_Clear"] = "مسح",
        ["MudDataGrid_ClearFilter"] = "مسح الفلتر",
        ["MudDataGrid_Column"] = "العمود",
        ["MudDataGrid_Columns"] = "الأعمدة",
        ["MudDataGrid_Contains"] = "يحتوي",
        ["MudDataGrid_EndsWith"] = "ينتهي بـ",
        ["MudDataGrid_Equals"] = "يساوي",
        ["MudDataGrid_False"] = "لا",
        ["MudDataGrid_Filter"] = "فلتر",
        ["MudDataGrid_FilterValue"] = "قيمة الفلتر",
        ["MudDataGrid_Hide"] = "إخفاء",
        ["MudDataGrid_HideAll"] = "إخفاء الكل",
        ["MudDataGrid_Is"] = "مطابق",
        ["MudDataGrid_IsAfter"] = "بعد",
        ["MudDataGrid_IsAnyOf"] = "أي من",
        ["MudDataGrid_IsBefore"] = "قبل",
        ["MudDataGrid_IsEmpty"] = "فارغ",
        ["MudDataGrid_IsNot"] = "غير مطابق",
        ["MudDataGrid_IsNotEmpty"] = "غير فارغ",
        ["MudDataGrid_IsOnOrAfter"] = "في أو بعد",
        ["MudDataGrid_IsOnOrBefore"] = "في أو قبل",
        ["MudDataGrid_Loading"] = "جار التحميل",
        ["MudDataGrid_NotContains"] = "لا يحتوي",
        ["MudDataGrid_NotEquals"] = "لا يساوي",
        ["MudDataGrid_OpenFilters"] = "فتح الفلاتر",
        ["MudDataGrid_Operator"] = "الشرط",
        ["MudDataGrid_RefreshData"] = "تحديث البيانات",
        ["MudDataGrid_RemoveFilter"] = "حذف الفلتر",
        ["MudDataGrid_Save"] = "حفظ",
        ["MudDataGrid_ShowAll"] = "إظهار الكل",
        ["MudDataGrid_ShowColumnOptions"] = "خيارات العمود",
        ["MudDataGrid_Sort"] = "ترتيب",
        ["MudDataGrid_StartsWith"] = "يبدأ بـ",
        ["MudDataGrid_True"] = "نعم",
        ["MudDataGrid_Unsort"] = "إلغاء الترتيب",
        ["MudDataGrid_Value"] = "القيمة",
        ["MudDateRangePicker_End"] = "إلى تاريخ",
        ["MudDateRangePicker_Start"] = "من تاريخ",
        ["MudDialog_Close"] = "إغلاق",
        ["MudAlert_Close"] = "إغلاق",
        ["MudChip_Close"] = "إغلاق",
        ["MudSnackbar_Close"] = "إغلاق",
        ["MudFormComponent_Required"] = "مطلوب",
        ["MudInput_Clear"] = "مسح",
        ["MudInput_Decrement"] = "إنقاص",
        ["MudInput_Increment"] = "زيادة",
        ["MudSelect_SelectAll"] = "تحديد الكل",
        ["MudBaseDatePicker_Open"] = "فتح",
        ["MudBreadcrumbs_Label"] = "مسار الصفحة",
    };

    public override LocalizedString this[string key] =>
        !lang.English && Arabic.TryGetValue(key, out var text)
            ? new LocalizedString(key, text)
            : new LocalizedString(key, key, resourceNotFound: true);

    public override LocalizedString this[string key, params object[] arguments] =>
        !lang.English && Arabic.TryGetValue(key, out var text)
            ? new LocalizedString(key, string.Format(System.Globalization.CultureInfo.InvariantCulture, text, arguments))
            : new LocalizedString(key, key, resourceNotFound: true);
}
