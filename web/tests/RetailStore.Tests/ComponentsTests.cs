using Microsoft.JSInterop;
using MudBlazor;
using RetailStore.Domain.Sales;
using RetailStore.UI.Components.Shared;
using RetailStore.UI.Services;

namespace RetailStore.Tests;

/// <summary>The logic behind the shared components of the redesign (phase 3): formats, the lookup search, the CSV
/// export, the lines through the invoice engine, the status styles, the shortcuts and the MudBlazor texts.</summary>
public class ComponentsTests
{
    [Fact]
    public void Money_and_quantities_are_invariant_with_their_decimals()
    {
        Assert.Equal("1,234,567.50", ErpFormat.Money(1234567.5m));
        Assert.Equal("-250.00", ErpFormat.Money(-250m));
        Assert.Equal("99.999", ErpFormat.Money(99.999m, 3));
        Assert.Equal("1,235", ErpFormat.Money(1234.6m, 0));
        Assert.Equal("-", ErpFormat.Money(null));
        Assert.Equal("12", ErpFormat.Qty(12m));
        Assert.Equal("0.125", ErpFormat.Qty(0.125m));
        Assert.Equal("15%", ErpFormat.Percent(0.15m));
        Assert.Equal("12.4%", ErpFormat.Percent(0.124m));
        Assert.Equal("2026-11-26", ErpFormat.Date(new DateTime(2026, 11, 26, 13, 5, 0)));
    }

    [Fact]
    public void Date_pickers_are_gregorian_in_both_languages()
    {
        Assert.IsType<System.Globalization.GregorianCalendar>(ErpFormat.DateCulture(false).DateTimeFormat.Calendar);
        Assert.IsType<System.Globalization.GregorianCalendar>(ErpFormat.DateCulture(true).DateTimeFormat.Calendar);
    }

    private sealed record Item(string Code, string Name);

    private static readonly Item[] Items =
    [
        new("10", "سكر أبيض"), new("1001", "أرز مصري"), new("2001", "شاي أخضر"), new("3001", "مكرونة بالأرز"),
    ];

    [Fact]
    public void Lookup_puts_the_exact_code_first_then_codes_then_names()
    {
        var found = LookupMatch.Filter(Items, "10", i => i.Code, i => i.Name).Select(i => i.Code).ToList();
        Assert.Equal(["10", "1001"], found);
        // a name that starts with the text before a name that holds it; alef forms folded (ارز = أرز)
        Assert.Equal(["1001", "3001"], LookupMatch.Filter(Items, "ارز", i => i.Code, i => i.Name).Select(i => i.Code));
        Assert.Equal(4, LookupMatch.Filter(Items, "", i => i.Code, i => i.Name).Count());
        Assert.Equal(2, LookupMatch.Filter(Items, " ", i => i.Code, i => i.Name, max: 2).Count());
        Assert.Empty(LookupMatch.Filter(Items, "قهوة", i => i.Code, i => i.Name));
    }

    [Fact]
    public void Csv_opens_in_excel_with_arabic_and_keeps_text_as_text()
    {
        var rows = new[] { new Item("=1+2", "شاي, \"أخضر\""), new Item("-5", "سكر") };
        string csv = GridCsv.Build(rows, [new GridExport<Item>("الكود", i => i.Code), new GridExport<Item>("الاسم", i => i.Name)]);
        Assert.StartsWith("﻿", csv);
        Assert.Equal("﻿الكود,الاسم\r\n'=1+2,\"شاي, \"\"أخضر\"\"\"\r\n'-5,سكر\r\n", csv);
        Assert.Equal("-12.5", GridCsv.Cell(-12.5m));
        Assert.Equal("2026-11-26", GridCsv.Cell(new DateTime(2026, 11, 26)));
        Assert.Equal("1", GridCsv.Cell(true));
        Assert.Equal("a-b.csv", GridCsv.FileName("a/b"));
    }

    [Fact]
    public void Lines_take_their_amounts_from_the_invoice_engine()
    {
        var lines = new List<InvoiceLine<Item>>
        {
            new() { Product = Items[0], Quantity = 2, Price = 6.5m, VatRate = 0.15m },
            new() { Product = Items[1], Quantity = 1, Price = 42.75m, Discount = 2.75m, VatRate = 0.15m },
            new(),   // the empty line at the end is not a line yet
        };
        var result = InvoiceLines.Calculate(lines, 0, pricesIncludeVat: true);
        var engine = InvoiceCalculator.Calc([new(2, 6.5m, 0, 0.15m), new(1, 42.75m, 2.75m, 0.15m)], 0, true);
        Assert.True(result.Ok);
        Assert.Equal(engine.Total, result.Total);
        Assert.Equal(53.00m, result.Total);
        Assert.Equal(engine.Tax, result.Tax);
        Assert.Equal(engine.Lines[1].Total, lines[1].Amounts!.Total);
        Assert.Null(lines[2].Amounts);

        lines[0].Quantity = 0;
        var wrong = InvoiceLines.Calculate(lines, 0, true);
        Assert.Equal("الكمية يجب أن تكون أكبر من صفر (السطر 1).", wrong.Problem);
        Assert.All(lines, l => Assert.Null(l.Amounts));

        Assert.True(InvoiceLines.Calculate(new List<InvoiceLine<Item>> { new() }, 0, true).Ok);
    }

    [Fact]
    public void Statuses_have_a_caption_and_a_color()
    {
        Assert.Equal(Color.Success, StatusStyles.For("posted").Color);
        Assert.Equal(Color.Error, StatusStyles.For("CANCELLED").Color);
        Assert.Equal(Color.Warning, StatusStyles.For("PARTIAL_PAID").Color);
        // a transfer sent and not yet received (StockTransfers.Status SENT) shows the caption of its page
        Assert.Equal(Color.Warning, StatusStyles.For("SENT").Color);
        Assert.Equal(RetailStore.UI.Services.WorkPages.TransferStatusName("SENT"), StatusStyles.For("SENT").Caption);
        Assert.Equal("مسودة", StatusStyles.For("DRAFT").Caption);
        Assert.Equal("XYZ", StatusStyles.For("XYZ").Caption);
        Assert.Equal(Color.Default, StatusStyles.For(null).Color);
    }

    [Fact]
    public void FilterPanel_default_period_is_this_month()
    {
        var r = FilterPanel.ThisMonth(new DateTime(2026, 2, 14));
        Assert.Equal(new DateTime(2026, 2, 1), r.Start);
        Assert.Equal(new DateTime(2026, 2, 28), r.End);
    }

    private sealed class FakeJs : IJSRuntime
    {
        public List<(string Name, object?[]? Args)> Calls { get; } = [];

        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args)
        {
            Calls.Add((identifier, args));
            return ValueTask.FromResult(default(TValue)!);
        }

        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args) =>
            InvokeAsync<TValue>(identifier, args);
    }

    [Fact]
    public async Task The_last_handler_of_a_key_runs_and_the_browser_knows_the_keys()
    {
        var js = new FakeJs();
        await using var shortcuts = new ShortcutService(js);
        var ran = new List<string>();
        var page = await shortcuts.RegisterAsync(ErpKey.Save, () => { ran.Add("page"); return Task.CompletedTask; });
        await shortcuts.RegisterAsync(ErpKey.New, () => { ran.Add("new"); return Task.CompletedTask; });
        var dialog = await shortcuts.RegisterAsync(ErpKey.Save, () => { ran.Add("dialog"); return Task.CompletedTask; });

        await shortcuts.OnKey("F2");
        await dialog.DisposeAsync();
        await shortcuts.OnKey("F2");
        await shortcuts.OnKey("F4");
        await shortcuts.OnKey("F3");   // no handler: the browser focuses the search box
        Assert.Equal(["dialog", "page", "new"], ran);

        var last = js.Calls.Last(c => c.Name == "erp.shortcuts.attach");
        Assert.Equal(["F2", "F4"], (string[])last.Args![1]!);
        await page.DisposeAsync();
        Assert.Equal(["F4"], (string[])js.Calls.Last().Args![1]!);
        Assert.Equal("Ctrl+P", ShortcutService.KeyName(ErpKey.Print));
        Assert.Equal("Escape", ShortcutService.KeyName(ErpKey.Close));
    }

    [Fact]
    public void MudBlazor_texts_are_arabic_only_in_the_arabic_interface()
    {
        var lang = new UiLanguage();
        var localizer = new ErpMudLocalizer(lang);
        Assert.Equal("صفوف في الصفحة", localizer["MudDataGridPager_RowsPerPage"].Value);
        Assert.Equal("1-25 من 80", localizer["MudDataGridPager_InfoFormat", 1, 25, 80].Value);
        Assert.True(localizer["MudColorPicker_Hue"].ResourceNotFound);
        lang.Set("en");
        Assert.True(localizer["MudDataGridPager_RowsPerPage"].ResourceNotFound);
    }
}
