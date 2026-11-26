---
name: mudblazor-erp-ui
description: Build or redesign any Razor page of the RetailStoreAccess web (RetailStore.UI / RetailStore.Web) with MudBlazor, following the approved templates A data screen, B document, C report, D dashboard. Use for every new web page, every page moved to MudBlazor (phase 6 batches of docs/ui-inventory.md) and every review of a web screen.
---

# The MudBlazor design system of the web

The rules of `AGENTS.md` ("Web (`docs/76`)" and the rest of the web rules) apply in full; this skill is how a screen is
built. The decisions, the components and the problems already met are in `docs/ui-inventory.md` sections 6-8, the
original plan in `docs/ERP-Redesign-Prompt.md`. Paths below are under `web/src/RetailStore.UI/`.

## 1) The approved templates: copy the pattern from them
| Type | Approved screen | Use it for |
|---|---|---|
| **A** data (list + add / edit) | `Components/DataScreen.razor` + `Components/DataScreenDialog.razor` | a master table: a row of `gen_web.WEB_SCREENS` makes the page (`<DataScreen T="..." />`, see `Pages/ProductsPage.razor`); never a hand-made list page for an Access `DataScreen` |
| **B** document (head + lines) | `Pages/Sales/PosPage.razor` | invoices, returns, transfers, vouchers, manual entries, closings |
| **C** report | `Pages/Accounting/LedgerPage.razor` | statements, trial balance, financials, stock, report centre, document lists (`/sales`, `/purchases`) |
| **D** dashboard / special | `Pages/Home.razor` (+ `Services/HomeReports.cs`) | the home page and any page of figures and charts |

The order of a page, always:
- **A**: `PageHeader` (primary "إضافة جديد", F4) -> `EntityGrid` (columns of the spec, inactive switch, actions column edit /
  delete, CSV) -> add / edit in `EntityDialog` (two columns, the service message on top and its field marked, a question
  of the service through `ConfirmDialog` then save again, "no" = `EntityDialog.StayOpen`).
- **B**: `PageHeader` (document number + `StatusChip`, save F2, post, print, new F4) -> the head in `MudPaper` + `MudGrid`
  (date, `LookupField` for the party, warehouse, payment, notes) -> `InvoiceLinesEditor` -> the summary card in
  `erp-doc-side` (sticky beside the lines from 1280px, under them below).
- **C**: `PageHeader` (print) -> `FilterPanel` (period + choices) -> a row of `StatCard` (e.g. opening / debit / credit /
  balance) -> `EntityGrid` of the rows (search, export) -> `EmptyState` before the first "عرض".
- **D**: a row of `StatCard` -> `MudChart` (colors `ErpTheme.ChartPalette`) -> short tables -> links to the user's screens.

## 2) The shared components (`Components/Shared`, never copy their markup into a page)
| Component | One line |
|---|---|
| `PageHeader` | `<PageHeader Title="@L.T("الأصناف")" Group="@L.T("المخازن")" Icon="@NavGroups.IconOf("products")"><Actions>...</Actions></PageHeader>` |
| `EntityGrid<T>` | `<EntityGrid T="Row" Items="rows" SearchText="r => r.Code + ' ' + r.Name" Export="cols" ExportName="@L.T("...")"><Columns>...</Columns></EntityGrid>` |
| `EntityDialog` | `<EntityDialog Title="@L.T("صنف جديد")" OnSave="SaveAsync">fields</EntityDialog>`, shown by `Dialogs.ShowAsync<X>(title, p, EntityDialog.Options)` (`WideOptions` for two columns); `OnSave` returns "" or the service message |
| `LookupField<T>` | `<LookupField T="Customer" Label="@L.T("العميل")" Items="customers" @bind-Value="customer" CodeOf="c => c.Code" NameOf="c => c.Name" />` (+ `InfoTemplate`, or `SearchFunc` for a service) |
| `InvoiceLinesEditor<TProduct>` | `<InvoiceLinesEditor TProduct="P" Lines="lines" Products="products" CodeOf=... NameOf=... PriceOf=... VatRate=... OnChange="Recalc" />` (+ `UnitTemplate`, `OnProductPicked`, `ExtraColumn` + `ExtraColumnTitle` for a column of the page after the discount) |
| `FilterPanel` | `<FilterPanel Range="range" RangeChanged="r => range = r" OnShow="Load" OnClear="Reset" Busy="busy"><MudItem ...>choices</MudItem></FilterPanel>`; default `FilterPanel.ThisMonth()` |
| `MoneyText` / `QtyText` | `<MoneyText Value="x.Total" ShowCurrency="true" />`, `<QtyText Value="l.Quantity" Unit="@l.UnitName" />` |
| `StatusChip` | `<StatusChip Status="@doc.Status" />` (codes of `StatusStyles`), or `Text` + `Color` |
| `ConfirmDialog` | `if (await ConfirmDialog.AskAsync(Dialogs, L.T("حذف ..."), L.T("..."), L.T("حذف"))) { ... }` (`Color.Primary` for a posting) |
| `StatCard` | `<StatCard Title="@L.T("...")" Icon="@Icons.Material.Outlined.X" Color="Color.Primary"><ValueContent><MoneyText .../></ValueContent></StatCard>` (`Change`, `DownIsGood`, `Href`) |
| `EmptyState` | `<EmptyState Icon=... Title="@L.T("لا توجد ...")" ActionText="@L.T("...")" OnAction="New" />` |
| `KeyboardShortcuts` | `<KeyboardShortcuts OnSave="Save" OnNew="New" OnPrint="Print" />` |

A change a page needs in a shared component is made once, in the component (with its example comment and the gallery
`/dev/components`), never by a copy in the page. A sub-agent working on one batch never edits `Components/Shared`.

## 3) Design rules
- **Spacing**: the layout gives `pa-4` (`MainLayout`); `MudGrid Spacing="3"`; `mb-4` between sections, `mb-3` inside one.
- **Fields**: `Variant="Variant.Outlined"` and `Margin="Margin.Dense"` on every `MudTextField` / `MudNumericField` /
  `MudSelect` / `MudDatePicker` / `MudAutocomplete`; `Label="@L.T("...")"`; `Required` + `RequiredError="@L.T("مطلوب")"`
  only for what the service also refuses empty. Dates: `Culture="@ErpFormat.DateCulture(L.English)"` (ar-SA is Hijri),
  `DateFormat="yyyy-MM-dd"`. Numbers: `HideSpinButtons`, `Class="erp-cell-num"`, invariant culture.
- **Buttons**: the main action `Variant.Filled` + `Color.Primary` (one action per page; it may be repeated at the foot of the summary card, as the post button of `/pos`), others `Variant.Outlined`, delete
  `Color.Error`; the shortcut shown as `<kbd class="erp-kbd">F2</kbd>`.
- **Tables**: always `EntityGrid` (Dense, Hover, Striped); numbers aligned with numbers, sorted by value, negative red
  (`MoneyText`); a report grid has no column filters (they overflow); never a raw `<table>` except inside a shared component.
- **Colors and fonts** only from `Theme/ErpTheme.cs` (`Color.*`, `var(--mud-palette-*)`); no hex color, no `style=` in a page;
  page CSS goes into `wwwroot/erp.css` with logical properties (`margin-inline-start`, `inset-inline-end`, `text-align: start`),
  never `left` / `right`.
- **Icons**: `Icons.Material.Outlined.*` in pages (Filled only for arrows inside shared components); the icon of a menu
  page from `NavGroups.IconOf(route)`.
- **Texts**: every Arabic text is `L.T("...")` (also `Label=`, `Placeholder=`, `Title=`, `aria-label=`) with its English in
  `tools/i18n_web.py` `WEB_EN` (or `i18n_en.EN`); messages of services are shown with `L.T(message)` as they come.
- **Both directions**: `MudRTLProvider` follows `L.Dir`; a number or code inside Arabic text is `<bdi dir="ltr">` /
  `erp-num`. Check the English (LTR) page as well as the Arabic one.
- **Money**: through `MoneyText` (currency of the settings, `UiCurrency`), never "ريال" / "جنيه" / "SAR" written in a page;
  line amounts and totals from `InvoiceCalculator` (`InvoiceLines.Calculate`), never `Quantity * Price` in a page.
- **States**: loading (`Busy` / `MudProgressLinear` of the grid), empty (`EmptyState`), error
  (`<MudAlert Severity="Severity.Error" Variant="Variant.Outlined" role="alert">`), no permission (the message of
  `Permissions.OpenProblemAsync`).
- **Keyboard**: F2 save, F3 search (`data-erp-search`), F4 new, Esc close, Ctrl+P print, all through `KeyboardShortcuts` /
  `ShortcutService`; the first field takes the focus (`AutoFocus`).
- **Print**: `class="no-print"` on filters and buttons, `erp-print-title` on the printed heading.

## 4) Forbidden
- Changing a service, `DbContext`, a calculation, a check, its order or its message (`RetailStore.Data`, `RetailStore.Domain`):
  the redesign is the UI layer only. A small back-end change a screen needs (paging for `ServerData`, a new list method) is
  asked of the owner first.
- The old markup of `retailstore.css` in a moved page (`class="screen"`, `class="btn`, `class="message`, `<select`,
  `<table`, `dialog-backdrop`), Bootstrap or a second UI library, inline styles.
- Loading every row of a large table at once without the service limit (`MaxListRows`, `take`).
- Editing `Generated/*.g.cs` (run `python3 tools/generate.py`).

## 5) How to work a batch (phase 6)
1. Pick the batch in `docs/ui-inventory.md` (the order: sales, purchases, stock, accounts, reports, settings) and mark
   its rows 🟨; read each page and its service first.
2. Decide its template (A / B / C / D) from the table of section 2 of the inventory, open the approved screen of that
   template and follow its order and its components.
3. Keep every `@inject`, every service call, its arguments and the order of save -> question -> save again; keep the
   `OpenProblemAsync` of the Access form of the page (`WorkPages`).
4. Add the new texts to `tools/i18n_web.py`, then:
   ```bash
   python3 tools/generate.py
   cd web && dotnet build RetailStore.slnx
   cd web && dotnet test RetailStore.slnx
   python3 -m unittest discover -s tests -p "test_web*.py"
   ```
   (`RETAIL_SKIP_DOTNET=1` skips the `dotnet test` inside `test_web` when it was just run.)
5. Run the site (`cd web/src/RetailStore.Web && dotnet run`, http://localhost:5137, `admin` with no password on a new
   development database, then set one) and take screenshots with Playwright (preinstalled; deviceScaleFactor 1 for your
   own check, 2 for the owner): every page light and dark, Arabic and English, 1440px and 768px; fix alignment, RTL,
   overflow (`contain: paint` on a wide table) before showing them.
6. Mark the rows ✅ in `docs/ui-inventory.md`, update `docs/dev/05-Progress.md` and the log of `docs/51`, commit, push,
   report in Arabic and wait for approval.

## 6) Checklist before handing over a screen
- [ ] it follows its approved template and starts with `PageHeader`
- [ ] `dotnet build` with no error and no new warning; `test_web*.py` and xUnit green (`tests/test_web_rules.py` included)
- [ ] no business logic changed (same services, arguments, order, messages)
- [ ] every text through `L.T`, with its English
- [ ] RTL and LTR correct (icons, alignment, arrows, negative numbers)
- [ ] dark mode readable
- [ ] keyboard shortcuts work, the focus is on the first field
- [ ] loading / empty / error / no permission states
- [ ] screenshots at 1440px and 768px, no horizontal scroll of the page
