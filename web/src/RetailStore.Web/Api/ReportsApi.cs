using RetailStore.Data.Reports;
using RetailStore.Data.Services;
using RetailStore.UI.Services;

namespace RetailStore.Web.Api;

/// <summary>A column of a report run: its caption (in the language asked), its format and its total.</summary>
public sealed record ReportColumnJson(string Caption, string Format, string? Total);

/// <summary>A report run as JSON: the cells as the report shows them (running balances done), the totals, the summary rows, or the
/// rows of a card report.</summary>
public sealed record ReportJson(string Key, string Title, string Criteria, string Warning, string Link, List<ReportColumnJson> Columns,
                                List<List<string>> Rows, List<ReportLineJson> Summary, List<ReportLineJson> Card, string Note);

public sealed record ReportLineJson(string Caption, string Value);

/// <summary>A report of the centre: its key, title, what it asks for (Needs, see docs/82) and the page that shows it instead.</summary>
public sealed record ReportInfoJson(string Key, string Title, string Needs, string Hint, string Link);

/// <summary>The report centre and the financial indicators (plan 6F, docs/82): the same services as the pages of the web
/// (modScreens.PrepareReport over the saved queries, modIndicators).</summary>
public static class ReportsApi
{
    public static void Map(RouteGroupBuilder secured)
    {
        var reports = secured.MapGroup("/reports").WithTags("Reports");

        reports.MapGet("", (HttpContext http) =>
            Results.Ok(ReportCatalog.All.Select(r => new ReportInfoJson(r.Key, ApiResults.Text(http, r.Title), r.Needs,
                                                                         ApiResults.Text(http, ReportService.NeedsHint(r)), r.Link)).ToList()))
        .Needs("REPORTS", "frmReportCenter").WithName("ListReports")
        .WithSummary("The reports of the report centre (frmReportCenter) and what each one asks for")
        .WithDescription("Needs: P a period, D a period on the date of the records, C / S / R a customer / supplier / product to " +
                         "choose, c / s / r / e / w an optional filter, W / b an optional warehouse / cash box, $ profit figures " +
                         "(REPORTS_PROFIT), A the audit trail (AUDIT_LOG), # the treasury (CASH_BOX), J accounting (JOURNAL, the " +
                         "journal is updated first). Link = a page of the web that shows the report (the financial statements).")
        .Produces<List<ReportInfoJson>>();

        reports.MapGet("/{key}", async (string key, HttpContext http, ReportService service, DateTime? from, DateTime? to, int? customerId,
                                        int? supplierId, int? productId, int? warehouseId, int? cashBoxId, int? expenseTypeId) =>
        {
            bool english = ApiResults.English(http);
            var choices = new ReportChoices(from, to, customerId, supplierId, productId, warehouseId, cashBoxId, expenseTypeId);
            var r = await service.RunAsync(UserInfo.Acting(http.User), key, choices, english);
            if (r.Report is null)
            {
                return ApiResults.NotFound(http);
            }
            if (!r.Ok)
            {
                return ApiResults.Refused(http, r.Message);
            }
            return Results.Ok(Json(http, r.Report, r));
        })
        .Needs("REPORTS", "frmReportCenter").WithName("RunReport")
        .WithSummary("Runs a report of the centre with its choices (modScreens.PrepareReport) and gives its rows as the report shows them")
        .WithDescription("The period (from, to) and the choices the report asks for (GET /reports). Each cell is the text of the Access " +
                         "report (dates yyyy/mm/dd, amounts #,##0.00); running balances and totals are done. A refused choice is a " +
                         "problem 422 with the message of Access.")
        .Produces<ReportJson>().ProducesProblem(StatusCodes.Status422UnprocessableEntity).ProducesProblem(StatusCodes.Status404NotFound);

        secured.MapGet("/dashboard/indicators", async (HttpContext http, IndicatorsService service, bool? refresh) =>
            Results.Ok(await service.IndicatorsAsync(UserInfo.Acting(http.User), DateTime.Today, refresh == true)))
        .Needs("DASHBOARD_FINANCIAL").WithTags("Dashboard").WithName("FinancialIndicators")
        .WithSummary("The financial indicators of the dashboard (modIndicators): gross margin, stock turnover, collection period, current ratio")
        .WithDescription("From the journal (FinancialIndicatorsQuery), brought up to date first; the figures are kept 10 minutes " +
                         "(refresh=true computes them again). Each card: its value, the line under it and its tone.")
        .Produces<IndicatorsView>();
    }

    private static ReportJson Json(HttpContext http, ReportDef rep, ReportResult r)
    {
        string T(string s) => ApiResults.Text(http, s);
        var running = ReportService.Running(rep, r.Rows);
        var cols = rep.Columns.Select((c, i) => new ReportColumnJson(T(c.Caption), c.Format,
                                                                     c.Total ? ReportService.TotalText(rep, r.Rows, i) : null)).ToList();
        var rows = r.Rows.Select((_, i) => rep.Columns.Select((_, col) => ReportService.CellText(rep, r.Rows, running, i, col)).ToList()).ToList();
        var summary = rep.Summary.Select(s => new ReportLineJson(T(s.Caption), ReportRow.Cell(s.Value(r.Rows), "money"))).ToList();
        var card = r.Rows.Count == 0 ? [] : rep.Card.Select(c => new ReportLineJson(T(c.Caption), ReportRow.Cell(c.Value(r.Rows[0]), c.Format))).ToList();
        return new ReportJson(rep.Key, T(rep.Title), T(r.Criteria), r.Warning.Length > 0 ? T(r.Warning) : "", rep.Link, cols, rows, summary,
                              card, rep.Note.Length > 0 ? T(rep.Note) : "");
    }
}
