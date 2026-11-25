using System.Net;
using System.Net.Http.Json;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Services;
using RetailStore.RepApp.Offline;

namespace RetailStore.Tests;

/// <summary>The rest of the web (plan 7C, docs/87): the purchase return, the bank reconciliation, the manual journal entries,
/// the cashier closing of the API, and the invoice and the receipt of the rep app.</summary>
public class WebCompletionTests
{
    private static readonly DateTime Today = DateTime.Today;

    // ------------------------------------------------------------------ the purchase return (modPurchases.PostPurchaseReturn)

    [Fact]
    public async Task A_purchase_return_takes_the_goods_out_at_their_cost_and_the_checks_are_those_of_access()
    {
        using var t = new TestDb();
        int supplier, rice;
        await using (var db = t.CreateDbContext())
        {
            var s = new Suppliers { SupplierName = "مورد" };
            var p = new Products { ProductCode = "R1", ProductName = "أرز", SellingPrice = 20, PurchasePrice = 10 };
            db.AddRange(s, p);
            db.Settings.First().VATNumber = "300000000000003";
            await db.SaveChangesAsync();
            (supplier, rice) = (s.SupplierID, p.ProductID);
        }
        var purchases = new PurchaseService(t, t.Permissions);
        var buy = await purchases.PostPurchaseAsync(TestDb.Admin, new PurchaseRequest(supplier, "F-9", null, "CREDIT", 1, true, 0, 0, "",
                                                                                    [new PurchaseLine(rice, 10, 10)]));
        Assert.Equal(SaveStatus.Saved, buy.Status);
        var source = (await purchases.ReturnableAsync("F-9", false))!;                // found by the supplier's invoice number too
        Assert.Equal(("CREDIT", 10m, 10m), (source.DefaultRefund, source.Lines[0].AvailableQty, source.Lines[0].InStock));
        int line = source.Lines[0].PurchaseDetailID;
        PurchaseReturnRequest Ret(string? reason, decimal qty) => new(buy.InvoiceID, reason, "CREDIT", null, [new PurchaseReturnLineRequest(line, qty)]);

        Assert.Equal("سبب الإرجاع مطلوب.", (await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret(" ", 1))).Message);
        Assert.Equal("فاتورة الشراء الأصلية غير موجودة.",
                     (await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret("تالف", 1) with { PurchaseInvoiceID = 999 })).Message);
        Assert.Equal("حدد الكمية المرتجعة لصنف واحد على الأقل.", (await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret("تالف", 0))).Message);
        Assert.StartsWith("الكمية المرتجعة (11)", (await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret("تالف", 11))).Message);

        var r = await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret("تالف", 4));
        Assert.Equal((SaveStatus.Saved, "PRT-000001", 46m), (r.Status, r.InvoiceNumber, r.Total));     // 4 x 10 + 15 %
        await using (var db = t.CreateDbContext())
        {
            var p = await db.Products.FindAsync(rice);
            Assert.Equal((6m, 10m), (p!.CurrentQuantity, p.AverageCost));
            var move = await db.InventoryTransactions.SingleAsync(m => m.ReferenceType == "PURCHASE_RETURN");
            Assert.Equal((-4m, 10m, 3, 1), (move.Quantity, move.UnitCost, move.TransactionTypeID, move.WarehouseID));
            Assert.Equal(115m - 46m, (await db.Suppliers.FindAsync(supplier))!.CurrentBalance);       // taken off what is owed
            var row = await db.PurchaseReturns.SingleAsync();
            Assert.Equal((40m, 6m, 0m, "CREDIT"), (row.TaxableAmount, row.Tax, row.RefundedAmount, row.RefundType));
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "PURCHASE_RETURN" && a.RecordID == "PRT-000001"));
        }
        Assert.Equal(6m, (await purchases.ReturnableAsync(source.InvoiceNumber, false))!.Lines[0].AvailableQty);
        Assert.Single(await purchases.PurchaseReturnsAsync(Today, Today, false));
        // the rest of the line takes what is left of its value and VAT
        var last = await purchases.PostPurchaseReturnAsync(TestDb.Admin, Ret("تالف", 6));
        Assert.Equal(69m, last.Total);
    }

    // ------------------------------------------------------------------ the bank reconciliation (modBank)

    private static async Task<int> BankWithItemsAsync(TestDb t)
    {
        await using var db = t.CreateDbContext();
        var bank = new Banks { BankName = "بنك الرياض" };
        db.Banks.Add(bank);
        await db.SaveChangesAsync();
        await RetailStore.Data.Journal.JournalService.EnsureAccountsAsync(db);          // 120000 + bank
        int n = 0;
        async Task Item(DateTime d, decimal amount)
        {
            n++;
            var e = new JournalEntries { EntryNumber = "JV-" + n, EntryDate = d, SourceType = "BANK_TX", SourceID = n, SourceNumber = "BNK-" + n,
                                         Description = "حركة " + n, TotalDebit = Math.Abs(amount), TotalCredit = Math.Abs(amount) };
            db.JournalEntries.Add(e);
            await db.SaveChangesAsync();
            db.JournalLines.AddRange(new JournalLines { EntryID = e.EntryID, LineNumber = 1, AccountCode = 120000 + bank.BankID, Debit = Math.Max(amount, 0), Credit = Math.Max(-amount, 0) },
                                     new JournalLines { EntryID = e.EntryID, LineNumber = 2, AccountCode = 110001, Debit = Math.Max(-amount, 0), Credit = Math.Max(amount, 0) });
            await db.SaveChangesAsync();
        }
        await Item(Today.AddDays(-10), 1000);
        await Item(Today.AddDays(-5), -300);
        await Item(Today.AddDays(-1), 200);
        await Item(Today, 50);                       // after the statement day
        return bank.BankID;
    }

    [Fact]
    public async Task A_bank_reconciliation_ticks_the_statement_and_is_approved_without_a_difference()
    {
        using var t = new TestDb();
        int bank = await BankWithItemsAsync(t);
        var recon = new BankReconService(t, t.Permissions);
        Assert.Equal("اختر البنك.", (await recon.OpenAsync(TestDb.Admin, 999, Today, 0)).Message);
        Assert.Equal("اكتب تاريخ كشف البنك.", (await recon.OpenAsync(TestDb.Admin, bank, null, 0)).Message);
        Assert.Equal("تاريخ الكشف بعد اليوم.", (await recon.OpenAsync(TestDb.Admin, bank, Today.AddDays(1), 0)).Message);

        var open = await recon.OpenAsync(TestDb.Admin, bank, Today.AddDays(-1), 700);
        Assert.Equal((SaveStatus.Saved, "REC-000001"), (open.Status, open.Number));
        int id = open.Id;
        var v = (await recon.ViewAsync(id, false))!;
        Assert.Equal((900m, 900m, 0m, 700m), (v.Figures.Book, v.Figures.Outstanding, v.Figures.Adjusted, v.Figures.Difference));
        Assert.Equal(3, v.Open.Count);                                                     // not the move of today
        Assert.Equal("حركة بنكية", v.Open[0].TypeName);

        Assert.Equal("", await recon.ClearAsync(id, "BANK_TX", 1));
        Assert.Equal("", await recon.ClearAsync(id, "BANK_TX", 2));
        Assert.Equal("العملية غير موجودة أو مطابقة من قبل.", await recon.ClearAsync(id, "BANK_TX", 1));
        v = (await recon.ViewAsync(id, false))!;
        Assert.Equal((200m, 700m, 0m), (v.Figures.Outstanding, v.Figures.Adjusted, v.Figures.Difference));
        // the same balance, so the open one is changed, not doubled
        Assert.Equal(id, (await recon.OpenAsync(TestDb.Admin, bank, Today.AddDays(-1), 800)).Id);
        Assert.StartsWith("يوجد فرق 100.00 بين كشف البنك والدفاتر.", await recon.FinishAsync(TestDb.Admin, id));
        await recon.OpenAsync(TestDb.Admin, bank, Today.AddDays(-1), 700);
        Assert.Equal("", await recon.FinishAsync(TestDb.Admin, id));
        Assert.Equal(BankReconService.NotFoundOrDone, await recon.UnclearAsync(id, "BANK_TX", 1));
        await using (var db = t.CreateDbContext())
        {
            var r = await db.BankReconciliations.SingleAsync();
            Assert.Equal(("DONE", 900m, 200m), (r.Status, r.BookBalance, r.Outstanding));
        }
        Assert.StartsWith("آخر تسوية معتمدة لهذا البنك في", (await recon.OpenAsync(TestDb.Admin, bank, Today.AddDays(-2), 0)).Message);
        Assert.Equal("تُحذف التسوية الجارية فقط (أعد فتح المعتمدة أولًا).", await recon.DeleteAsync(TestDb.Admin, id));
        Assert.Equal("", await recon.ReopenAsync(TestDb.Admin, id));
        Assert.Equal("", await recon.DeleteAsync(TestDb.Admin, id));
        await using (var db = t.CreateDbContext())
        {
            Assert.Empty(await db.BankClearings.ToListAsync());
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "BANK_RECON_DELETE"));
        }
        int cashier;
        await using (var db = t.CreateDbContext())
        {
            var e = new Employees { EmployeeName = "كاشير", Username = "c1", RoleID = 3 };
            db.Employees.Add(e);
            await db.SaveChangesAsync();
            cashier = e.EmployeeID;
        }
        Assert.Equal("لا تملك صلاحية البنوك.", (await recon.OpenAsync(new ActingUser(cashier, "TEST"), bank, Today, 0)).Message);
    }

    // ------------------------------------------------------------------ manual journal entries (modManualEntry)

    [Fact]
    public async Task A_manual_entry_follows_manual_entry_problem_and_can_be_changed_reversed_and_deleted()
    {
        using var t = new TestDb();
        var manual = new ManualEntryService(t, t.Permissions);
        ManualEntryRequest Entry(params ManualLineInput[] lines) => new(0, Today, "إثبات إيجار مستحق", "R-1", null, lines);
        async Task<string> Problem(ManualEntryRequest r) => (await manual.SaveAsync(TestDb.Admin, r)).Message;

        Assert.Equal("اكتب تاريخ القيد.", await Problem(Entry() with { EntryDate = null }));
        Assert.Equal("اكتب بيان القيد.", await Problem(Entry() with { Description = " " }));
        Assert.Equal("اختر الحساب في كل سطر فيه مبلغ.", await Problem(Entry(new ManualLineInput(null, 10, 0))));
        Assert.Equal("المبالغ لا تكون سالبة: اكتب المبلغ في الجانب الآخر.", await Problem(Entry(new ManualLineInput(5610, -10, 0))));
        Assert.Equal("السطر يكون مدينًا أو دائنًا، وليس الاثنين.", await Problem(Entry(new ManualLineInput(5610, 10, 10))));
        Assert.Equal("اكتب المبلغ مدينًا أو دائنًا في كل سطر فيه حساب.", await Problem(Entry(new ManualLineInput(5610, 0, 0))));
        Assert.Equal("الحساب 1100 حساب رئيسي أو معطّل: القيود على الحسابات الفرعية النشطة فقط.",
                     await Problem(Entry(new ManualLineInput(1100, 10, 0))));
        Assert.Equal("القيد سطران على الأقل: طرف مدين وطرف دائن.", await Problem(Entry(new ManualLineInput(5610, 10, 0), new ManualLineInput(null, 0, 0))));
        Assert.Equal("القيد غير متوازن: المدين 100.00 والدائن 90.00 (الفرق 10.00).",
                     await Problem(Entry(new ManualLineInput(5610, 100, 0), new ManualLineInput(2100, 0, 90))));

        var saved = await manual.SaveAsync(TestDb.Admin, Entry(new ManualLineInput(5610, 100, 0, "إيجار"), new ManualLineInput(2100, 0, 100)));
        Assert.Equal((SaveStatus.Saved, "MJ-000001"), (saved.Status, saved.Number));
        var v = (await manual.EntryAsync(saved.Id))!;
        Assert.Equal((100m, 2, "إيجار", false), (v.TotalAmount, v.Lines.Count, v.Lines[0].LineText, v.InClosedPeriod));

        var changed = await manual.SaveAsync(TestDb.Admin, Entry(new ManualLineInput(5610, 120, 0), new ManualLineInput(2100, 0, 120)) with { ManualEntryID = saved.Id });
        Assert.Equal(("MJ-000001", saved.Id), (changed.Number, changed.Id));
        await using (var db = t.CreateDbContext())
        {
            Assert.Equal(120m, (await db.ManualEntries.SingleAsync()).TotalAmount);
            Assert.Equal(2, await db.ManualEntryLines.CountAsync());
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "MANUAL_ENTRY_EDIT" && a.ObjectName == "ManualEntries"));
        }
        var reversal = ManualEntryService.Reversal((await manual.EntryAsync(saved.Id))!, Today);
        Assert.Equal(("عكس القيد MJ-000001: إثبات إيجار مستحق", saved.Id, 0m, 120m),
                     (reversal.Description, reversal.ReversalOfID, reversal.Lines[0].Debit, reversal.Lines[0].Credit));
        Assert.Equal("MJ-000002", (await manual.SaveAsync(TestDb.Admin, reversal)).Number);
        Assert.Equal(2, (await manual.EntriesAsync(Today, Today)).Count);

        // a closed period: no change and no delete
        await using (var db = t.CreateDbContext())
        {
            db.Settings.First().ClosedThrough = Today;
            await db.SaveChangesAsync();
        }
        Assert.True((await manual.EntryAsync(saved.Id))!.InClosedPeriod);
        Assert.Equal(SaveStatus.Refused, (await manual.DeleteAsync(TestDb.Admin, saved.Id)).Status);
        await using (var db = t.CreateDbContext())
        {
            db.Settings.First().ClosedThrough = null;
            await db.SaveChangesAsync();
        }
        Assert.Equal(SaveStatus.Saved, (await manual.DeleteAsync(TestDb.Admin, saved.Id)).Status);
        Assert.Equal(ManualEntryService.NotFound, (await manual.DeleteAsync(TestDb.Admin, saved.Id)).Message);
        await using (var db = t.CreateDbContext())
        {
            Assert.True(await db.AuditLog.AnyAsync(a => a.ActionType == "MANUAL_ENTRY_DELETE"));
            Assert.Equal(2, await db.ManualEntryLines.CountAsync());                      // the lines of the reversal only
        }
        Assert.Equal("المدين: 10.00     الدائن: 10.00     متوازن",
                     ManualEntryService.TotalsText([new ManualLineInput(1, 10, 0), new ManualLineInput(2, 0, 10)]));
    }

    // ------------------------------------------------------------------ the invoice and the receipt of the rep app

    private static async Task GrantSalesAsync(ApiHost host)
    {
        await using var db = host.Db.CreateDbContext();
        var user = await db.Employees.SingleAsync(e => e.Username == "rep1");
        db.UserScreens.Add(new UserScreens { EmployeeID = user.EmployeeID, ScreenName = "frmPOS", CanOpen = true, CanAdd = true });
        db.UserScreens.Add(new UserScreens { EmployeeID = user.EmployeeID, ScreenName = "frmCustomerPayment", CanOpen = true, CanAdd = true });
        await db.SaveChangesAsync();
    }

    private static async Task<int> OpenRepDayAsync(ApiHost host, int rep, int van)
    {
        await using var db = host.Db.CreateDbContext();
        var box = new CashBoxes { BoxName = "عهدة المندوب", BoxType = "CUSTODY", SalesRepID = rep };
        db.CashBoxes.Add(box);
        await db.SaveChangesAsync();
        db.RepDays.Add(new RepDays { DayNumber = "RD-000001", SalesRepID = rep, WarehouseID = van, CashBoxID = box.CashBoxID, OpenedAt = DateTime.Now.AddHours(-2) });
        await db.SaveChangesAsync();
        return box.CashBoxID;
    }

    [Fact]
    public async Task The_rep_sells_from_his_van_and_collects_into_his_custody_box_once()
    {
        using var host = new ApiHost();
        var (rep, customerId, van, riceId) = await ApiTests.MakeRepAsync(host);
        await GrantSalesAsync(host);
        var phone = new RepPhone(host);
        await phone.Session.LoadAsync();
        Assert.Equal("", await phone.Session.LoginAsync("http://localhost", "rep1", "Rep@2026x"));
        await phone.Session.SyncAsync();
        var s = phone.Session;
        Assert.False(s.CanSell);                                                     // no rep day open yet

        var draft = s.NewSale(customerId)!;
        Assert.Equal("", draft.Add(riceId, 2));
        var refused = await s.SellAsync(draft);
        Assert.Equal((DocAnswer.Refused, RepDocService.NoOpenDay), (refused.State, refused.Message));

        int box = await OpenRepDayAsync(host, rep, van);
        await s.SyncAsync();
        Assert.True(s.CanSell);
        draft = s.NewSale(customerId)!;
        Assert.Equal("", draft.Add(riceId, 2));
        Assert.Equal(17m, draft.Lines[0].UnitPrice);                                  // the price of the customer's list
        string reference = draft.ClientRef;
        phone.Platform.HasNetwork = false;
        Assert.Equal(DocAnswer.Offline, (await s.SellAsync(draft)).State);         // never kept on the phone
        phone.Platform.HasNetwork = true;
        var sold = await s.SellAsync(draft);
        Assert.True(sold.State == DocAnswer.Saved, sold.Message);
        Assert.StartsWith("INV-", sold.Document!.Number);
        Assert.Equal(34m, sold.Document.Rows.Single(r => r.Label == "الإجمالي شامل الضريبة").Amount);
        Assert.Equal("اختر الطابعة في الإعدادات.", await s.PrintDocumentAsync(sold.Document));
        s.SetPrinter("Test printer", 58);
        Assert.Equal("", await s.PrintDocumentAsync(sold.Document));                 // the text with the QR under it
        Assert.Single(phone.Platform.Printed);
        // the answer was lost: the same invoice sent again is not added twice
        var again = await s.SellAsync(draft);
        Assert.Equal((DocAnswer.Saved, sold.Document.Number), (again.State, again.Document!.Number));
        await using (var db = host.Db.CreateDbContext())
        {
            var inv = await db.SalesInvoices.SingleAsync();
            Assert.Equal((van, box, rep, reference), (inv.WarehouseID, inv.CashBoxID!.Value, inv.SalesRepID!.Value, inv.ClientRef));
            Assert.Equal(-2m, (await db.InventoryTransactions.SingleAsync(m => m.ReferenceType == "SALE")).Quantity);
            Assert.Equal(van, (await db.InventoryTransactions.SingleAsync(m => m.ReferenceType == "SALE")).WarehouseID);
        }

        string payRef = Guid.NewGuid().ToString("N");
        var over = await s.CollectAsync(payRef, customerId, 500, null);
        Assert.True(over.State == DocAnswer.Confirm, over.Message);                                // over the balance of 300: the question of Access
        var paid = await s.CollectAsync(payRef, customerId, 500, null, confirmOverBalance: true);
        Assert.True(paid.State == DocAnswer.Saved, paid.Message);
        Assert.StartsWith("RCV-", paid.Document!.Number);
        Assert.Equal(DocAnswer.Saved, (await s.CollectAsync(payRef, customerId, 500, null, true)).State);
        await using (var db = host.Db.CreateDbContext())
        {
            var pay = await db.CustomerPayments.SingleAsync();
            Assert.Equal((box, rep, payRef), (pay.CashBoxID!.Value, pay.SalesRepID!.Value, pay.ClientRef));
        }
        phone.Dispose();
    }

    // ------------------------------------------------------------------ the API of the web

    [Fact]
    public async Task The_api_posts_manual_entries_and_reads_the_cashier_closing()
    {
        using var host = new ApiHost();
        var c = await host.AdminAsync();
        var made = await c.PostAsJsonAsync("/api/v1/manual-entries", new
        {
            EntryDate = Today, Description = "قيد افتتاحي", Lines = new[] { new { AccountCode = 110001, Debit = 500m, Credit = 0m }, new { AccountCode = 3200, Debit = 0m, Credit = 500m } },
        });
        Assert.True(made.StatusCode == HttpStatusCode.Created, await made.Content.ReadAsStringAsync());
        var bad = await c.PostAsJsonAsync("/api/v1/manual-entries", new { EntryDate = Today, Description = "x", Lines = Array.Empty<object>() });
        Assert.Equal(HttpStatusCode.UnprocessableEntity, bad.StatusCode);
        Assert.Contains("القيد سطران على الأقل", await bad.Content.ReadAsStringAsync());
        Assert.Equal(HttpStatusCode.OK, (await c.GetAsync("/api/v1/manual-entries")).StatusCode);
        var boxes = await c.GetAsync("/api/v1/cash-closings/boxes");
        Assert.Equal(HttpStatusCode.OK, boxes.StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await c.GetAsync("/api/v1/cash-closings/boxes/1")).StatusCode);
        var closing = await c.PostAsJsonAsync("/api/v1/cash-closings", new { CashBoxID = 1, Counted = 0m, Destination = "KEEP", TransferAmount = 0m });
        Assert.True(closing.StatusCode == HttpStatusCode.Created, await closing.Content.ReadAsStringAsync());
        Assert.Equal(HttpStatusCode.NotFound, (await c.GetAsync("/api/v1/sales-returns/source/NOPE")).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await c.GetAsync("/api/v1/purchase-returns")).StatusCode);
    }
}
