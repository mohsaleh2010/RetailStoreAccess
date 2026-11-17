using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Domain.Common;
using RetailStore.Domain.EInvoicing;

namespace RetailStore.Data.Services;

/// <summary>The signing program of the Egyptian e-invoice (modEtaInvoice.EtaSignDocument): it reads the serialized document and
/// returns the CAdES-BES signature (base64) made with the USB token of the taxpayer. Signature "" = Problem says why.</summary>
public interface IEtaSigner
{
    Task<(string Signature, string Problem)> SignAsync(string serialized, string exe, string? argsTemplate, string pin);
}

/// <summary>Runs the signing program of the settings (EtaSignerPath + EtaSignerArgs) with {IN} / {OUT} files, as Access does.</summary>
public sealed class ProcessEtaSigner : IEtaSigner
{
    public async Task<(string Signature, string Problem)> SignAsync(string serialized, string exe, string? argsTemplate, string pin)
    {
        string folder = Path.Combine(Path.GetTempPath(), "RetailStoreZatca");
        Directory.CreateDirectory(folder);
        string inFile = Path.Combine(folder, "eta-document.txt");
        string outFile = Path.Combine(folder, "eta-signature.txt");
        File.Delete(outFile);
        await File.WriteAllTextAsync(inFile, serialized, new UTF8Encoding(false));
        int rc;
        try
        {
            using var p = Process.Start(new ProcessStartInfo(exe, EtaDocument.SignerArguments(argsTemplate, inFile, outFile, pin))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
            }) ?? throw new InvalidOperationException(exe);
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(2));
            await p.WaitForExitAsync(cts.Token);
            rc = p.ExitCode;
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException or OperationCanceledException)
        {
            return ("", "تعذّر تشغيل برنامج التوقيع: " + e.Message);
        }
        if (!File.Exists(outFile))
        {
            return ("", "لم يُنشئ برنامج التوقيع ملف التوقيع (رمز " + rc + "). تأكد من فلاشة التوقيع ورقمها السري.");
        }
        string signature = string.Concat((await File.ReadAllLinesAsync(outFile)).Select(l => l.Trim()));
        File.Delete(outFile);
        File.Delete(inFile);
        return signature.Length == 0 ? ("", "ملف التوقيع فارغ (رمز " + rc + ").") : (signature, "");
    }
}

/// <summary>E-invoicing in C# (plan 7A, docs/85): modEInvoice with its two platforms, ZATCA phase 2 (modZatcaXml, modZatcaApi:
/// signing with the device key and reporting / clearance) and the Egyptian Tax Authority (modEtaReceipt: e-receipt, chain of
/// UUIDs; modEtaInvoice: B2B e-invoice signed by the signing program). It reads and writes the same fields as Access
/// (ZatcaStatus, EInvoiceXml, InvoiceHash, QRCodeData, EtaUUID ..., Settings.LastInvoiceHash / EtaLastUUID, EInvoiceLog),
/// so Access, the web and the office service can each send a document. The onboarding of the device stays in Access.</summary>
public sealed class EInvoiceService(IDbContextFactory<RetailDbContext> factory, HttpClient http, IEtaSigner? signer = null,
                                    Func<DateTime>? clock = null)
{
    private const string NL = "\r\n";
    private readonly IEtaSigner _signer = signer ?? new ProcessEtaSigner();
    private readonly Func<DateTime> _now = clock ?? (() => DateTime.Now);
    private string _posToken = "", _posTokenEnv = "", _erpToken = "", _erpTokenEnv = "";
    private DateTime _posTokenExpires, _erpTokenExpires;

    /// <summary>Who sends (EInvoiceLog.EmployeeID); the office service sends as nobody.</summary>
    public ActingUser User { get; init; } = ActingUser.Nobody;

    private static string Txt(string? v) => (v ?? "").Trim();

    private static string Left(string s, int n) => s.Length <= n ? s : s[..n];

    private static bool Digits(string s, int n) => s.Length == n && s.All(char.IsAsciiDigit);

    // ================================================================== modEInvoice
    /// <summary>modEInvoice.EInvoiceReadyProblem: "" when the platform of the country is set up.</summary>
    public async Task<string> ReadyProblemAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        return CountryRules.Normalize(s.CountryCode) == "EG" ? EtaReadyProblem(s) : ZatcaReadyProblem(s);
    }

    /// <summary>modEInvoice.SendEInvoice: sends one document to the platform of the country.</summary>
    public async Task<EInvoiceReply> SendAsync(string docKind, int docId)
    {
        Settings s;
        await using (var db = await factory.CreateDbContextAsync())
        {
            s = await SettingsReader.ReadAsync(db);
        }
        if (!s.EInvoiceEnabled)
        {
            return new("ERROR", "", "الفاتورة الإلكترونية غير مفعّلة في الإعدادات.");
        }
        bool egypt = CountryRules.Normalize(s.CountryCode) == "EG";
        string problem = egypt ? EtaReadyProblem(s) : ZatcaReadyProblem(s);
        EInvoiceReply reply;
        if (problem.Length > 0)
        {
            return new("ERROR", "", problem);
        }
        try
        {
            reply = egypt ? await EtaSendAsync(docKind, docId) : await ZatcaSendAsync(docKind, docId);
        }
        catch (Exception e) when (e is not OutOfMemoryException)
        {
            reply = new("ERROR", "", "تعذّر الإرسال: " + e.Message);
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            string? error = reply.Result == "OK" || reply.Message.Length == 0 ? null : Left(reply.Message, 255);
            if (docKind == "RETURN")
            {
                await db.SalesReturns.Where(r => r.SalesReturnID == docId).ExecuteUpdateAsync(u => u
                    .SetProperty(r => r.EInvoiceAttempts, r => (short)((r.EInvoiceAttempts ?? 0) + 1)).SetProperty(r => r.EInvoiceError, error));
            }
            else
            {
                await db.SalesInvoices.Where(h => h.SalesInvoiceID == docId).ExecuteUpdateAsync(u => u
                    .SetProperty(h => h.EInvoiceAttempts, h => (short)((h.EInvoiceAttempts ?? 0) + 1)).SetProperty(h => h.EInvoiceError, error));
            }
        }
        return reply;
    }

    private sealed record DocRef(string Kind, int Id, DateTime Date);

    private static async Task<List<DocRef>> DocsWithStatusAsync(RetailDbContext db, string status)
    {
        var docs = await db.SalesInvoices.Where(h => h.ZatcaStatus == status)
            .Select(h => new DocRef("SALE", h.SalesInvoiceID, h.InvoiceDate)).ToListAsync();
        docs.AddRange(await db.SalesReturns.Where(r => r.ZatcaStatus == status)
            .Select(r => new DocRef("RETURN", r.SalesReturnID, r.ReturnDate)).ToListAsync());
        // qryEInvoiceDocs ORDER BY DocDate, DocKind DESC, DocID
        return docs.OrderBy(d => d.Date).ThenByDescending(d => d.Kind, StringComparer.Ordinal).ThenBy(d => d.Id).ToList();
    }

    /// <summary>modEInvoice.SendPendingEInvoices: every PENDING document, oldest first; stops at the first network problem or
    /// error (the next would fail too). Returns "" or the last problem.</summary>
    public async Task<(int Sent, int Failed, string Problem)> SendPendingAsync()
    {
        List<DocRef> docs;
        await using (var db = await factory.CreateDbContextAsync())
        {
            docs = await DocsWithStatusAsync(db, "PENDING");
        }
        int sent = 0, failed = 0;
        string last = "";
        foreach (var d in docs)
        {
            var r = await SendAsync(d.Kind, d.Id);
            if (r.Sent)
            {
                sent++;
                continue;
            }
            failed++;
            last = r.Message;
            if (r.Result is "NETWORK" or "ERROR")
            {
                break;
            }
        }
        return (sent, failed, last);
    }

    /// <summary>modEInvoice.RefreshSubmittedEInvoices: Egypt reads the result of every SUBMITTED document.</summary>
    public async Task<(int Checked, int Changed, string Problem)> RefreshSubmittedAsync()
    {
        List<DocRef> docs;
        await using (var db = await factory.CreateDbContextAsync())
        {
            if (CountryRules.Normalize((await SettingsReader.ReadAsync(db)).CountryCode) != "EG")
            {
                return (0, 0, "");                    // Saudi Arabia answers at once: nothing to follow
            }
            docs = await DocsWithStatusAsync(db, "SUBMITTED");
        }
        int done = 0, changed = 0;
        foreach (var d in docs)
        {
            EInvoiceReply r;
            try
            {
                r = await EtaRefreshAsync(d.Kind, d.Id);
            }
            catch (Exception e) when (e is not OutOfMemoryException)
            {
                r = new("ERROR", "", e.Message);
            }
            done++;
            if (await StatusOfAsync(d.Kind, d.Id) != "SUBMITTED")
            {
                changed++;
            }
            if (r.Result is "NETWORK" or "ERROR")
            {
                return (done, changed, r.Message);
            }
        }
        return (done, changed, "");
    }

    public async Task<string> StatusOfAsync(string docKind, int docId)
    {
        await using var db = await factory.CreateDbContextAsync();
        return docKind == "RETURN"
            ? await db.SalesReturns.Where(r => r.SalesReturnID == docId).Select(r => r.ZatcaStatus).FirstOrDefaultAsync() ?? ""
            : await db.SalesInvoices.Where(h => h.SalesInvoiceID == docId).Select(h => h.ZatcaStatus).FirstOrDefaultAsync() ?? "";
    }

    /// <summary>modEInvoice.SetEInvoiceStatus.</summary>
    private async Task SetStatusAsync(string docKind, int docId, string status, string response)
    {
        await using var db = await factory.CreateDbContextAsync();
        string resp = Left(response, 60000);
        DateTime at = _now();
        if (docKind == "RETURN")
        {
            await db.SalesReturns.Where(r => r.SalesReturnID == docId).ExecuteUpdateAsync(u => u
                .SetProperty(r => r.ZatcaStatus, status).SetProperty(r => r.ZatcaSubmittedAt, at).SetProperty(r => r.ZatcaResponse, resp));
        }
        else
        {
            await db.SalesInvoices.Where(h => h.SalesInvoiceID == docId).ExecuteUpdateAsync(u => u
                .SetProperty(h => h.ZatcaStatus, status).SetProperty(h => h.ZatcaSubmittedAt, at).SetProperty(h => h.ZatcaResponse, resp));
        }
    }

    /// <summary>modEInvoice.LogEInvoice: one request to the platform and its reply. docKind "" = a request of the device.</summary>
    private async Task LogAsync(string docKind, int docId, string action, string endpoint, int httpStatus, string result, string message,
                                string requestBody, string responseBody, long ms)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        var row = new EInvoiceLog
        {
            LoggedAt = _now(),
            Country = CountryRules.Normalize(s.CountryCode),
            Action = Left(action, 20),
            Endpoint = endpoint.Length > 0 ? Left(endpoint, 255) : null,
            HttpStatus = (short)httpStatus,
            Result = result,
            Message = message.Length > 0 ? Left(message, 255) : null,
            RequestBody = requestBody.Length > 0 ? requestBody : null,
            ResponseBody = responseBody.Length > 0 ? responseBody : null,
            DurationMs = (int)Math.Min(ms, int.MaxValue),
            EmployeeID = User.EmployeeID == 0 ? null : User.EmployeeID,
        };
        if (docKind.Length > 0)
        {
            row.DocKind = docKind;
            row.DocID = docId;
            row.DocNumber = docKind == "RETURN"
                ? await db.SalesReturns.Where(r => r.SalesReturnID == docId).Select(r => r.ReturnNumber).FirstOrDefaultAsync()
                : await db.SalesInvoices.Where(h => h.SalesInvoiceID == docId).Select(h => h.InvoiceNumber).FirstOrDefaultAsync();
        }
        db.EInvoiceLog.Add(row);
        await db.SaveChangesAsync();
    }

    /// <summary>modHttp.HttpSend: the status (0 = no answer), the body and the time it took.</summary>
    private async Task<(int Status, string Body, long Ms, string Problem)> HttpAsync(HttpMethod method, string url,
        IEnumerable<(string Name, string Value)> headers, HttpContent? content, int seconds)
    {
        var watch = Stopwatch.StartNew();
        try
        {
            using var req = new HttpRequestMessage(method, url) { Content = content };
            foreach (var (name, value) in headers)
            {
                req.Headers.TryAddWithoutValidation(name, value);
            }
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(seconds));
            using var resp = await http.SendAsync(req, cts.Token);
            string body = await resp.Content.ReadAsStringAsync(cts.Token);
            return ((int)resp.StatusCode, body, watch.ElapsedMilliseconds, "");
        }
        catch (Exception e) when (e is HttpRequestException or OperationCanceledException or IOException)
        {
            return (0, "", watch.ElapsedMilliseconds, "تعذّر الاتصال بالمنظومة.");
        }
    }

    private static StringContent Json(string body) => new(body, new UTF8Encoding(false), "application/json");

    // ================================================================== the documents
    private sealed record Head(string Kind, int Id, string Number, DateTime Date, int CustomerID, string PayType, int? MethodId,
                               string SubType, string TypeCode, string? Uuid, int? Icv, decimal Taxable, decimal Tax, decimal Total,
                               string Reason, string OriginalNumber, string OriginalEtaUuid, string? Xml, string Status, string? EtaUuid,
                               string? EtaPrevious, string? SubmissionId, string? Hash);

    private sealed record LineRow(decimal Quantity, decimal UnitPrice, decimal Discount, decimal NetAmount, string VATCategory,
                                  decimal VATRate, decimal Tax, decimal LineTotal, string ProductCode, string ProductName,
                                  string? EtaItemType, string? EtaItemCode, string? ZatcaUnitCode, string? EtaUnitCode);

    private static async Task<Head?> HeadAsync(RetailDbContext db, string kind, int id)
    {
        if (kind == "RETURN")
        {
            return await (from r in db.SalesReturns
                          join o in db.SalesInvoices on r.SalesInvoiceID equals o.SalesInvoiceID
                          where r.SalesReturnID == id
                          select new Head("RETURN", r.SalesReturnID, r.ReturnNumber, r.ReturnDate, r.CustomerID, r.RefundType,
                                          r.PaymentMethodID, r.InvoiceSubType, r.InvoiceTypeCode, r.InvoiceUUID, r.ICV, r.TaxableAmount,
                                          r.Tax, r.TotalAmount, r.Reason, o.InvoiceNumber, o.EtaUUID ?? "", r.EInvoiceXml, r.ZatcaStatus,
                                          r.EtaUUID, r.EtaPreviousUUID, r.EtaSubmissionId, r.InvoiceHash)).FirstOrDefaultAsync();
        }
        return await db.SalesInvoices.Where(h => h.SalesInvoiceID == id)
            .Select(h => new Head("SALE", h.SalesInvoiceID, h.InvoiceNumber, h.InvoiceDate, h.CustomerID, h.PaymentType, h.PaymentMethodID,
                                  h.InvoiceSubType, h.InvoiceTypeCode, h.InvoiceUUID, h.ICV, h.TaxableAmount, h.Tax, h.TotalAmount, "", "",
                                  "", h.EInvoiceXml, h.ZatcaStatus, h.EtaUUID, h.EtaPreviousUUID, h.EtaSubmissionId, h.InvoiceHash))
            .FirstOrDefaultAsync();
    }

    private static async Task<List<LineRow>> LinesAsync(RetailDbContext db, string kind, int id)
    {
        if (kind == "RETURN")
        {
            return await (from d in db.SalesReturnDetails
                          join p in db.Products on d.ProductID equals p.ProductID
                          join u in db.Units on p.UnitID equals u.UnitID
                          join v0 in db.Units on d.UnitID equals (int?)v0.UnitID into vs
                          from v in vs.DefaultIfEmpty()
                          where d.SalesReturnID == id
                          orderby d.ReturnDetailID
                          select new LineRow(d.Quantity, d.UnitPrice, d.Discount, d.NetAmount, d.VATCategory, d.VATRate, d.Tax, d.LineTotal,
                                             p.ProductCode, p.ProductName, p.EtaItemType, p.EtaItemCode,
                                             v == null ? u.ZatcaUnitCode : v.ZatcaUnitCode, v == null ? u.EtaUnitCode : v.EtaUnitCode))
                .ToListAsync();
        }
        return await (from d in db.SalesInvoiceDetails
                      join p in db.Products on d.ProductID equals p.ProductID
                      join u in db.Units on p.UnitID equals u.UnitID
                      join v0 in db.Units on d.UnitID equals (int?)v0.UnitID into vs
                      from v in vs.DefaultIfEmpty()
                      where d.SalesInvoiceID == id
                      orderby d.SalesDetailID
                      select new LineRow(d.Quantity, d.UnitPrice, d.Discount, d.NetAmount, d.VATCategory, d.VATRate, d.Tax, d.LineTotal,
                                         p.ProductCode, p.ProductName, p.EtaItemType, p.EtaItemCode,
                                         v == null ? u.ZatcaUnitCode : v.ZatcaUnitCode, v == null ? u.EtaUnitCode : v.EtaUnitCode))
            .ToListAsync();
    }

    private static async Task SaveDocAsync(RetailDbContext db, string kind, int id, Action<SalesInvoices> sale, Action<SalesReturns> ret)
    {
        if (kind == "RETURN")
        {
            ret(await db.SalesReturns.FirstAsync(r => r.SalesReturnID == id));
        }
        else
        {
            sale(await db.SalesInvoices.FirstAsync(h => h.SalesInvoiceID == id));
        }
        await db.SaveChangesAsync();
    }

    /// <summary>"UPDATE Settings SET X = X": takes the lock of the chain inside the transaction (Access does the same).</summary>
    private static Task LockChainAsync(RetailDbContext db, string field) => field == "EtaLastUUID"
        ? db.Database.ExecuteSqlRawAsync("UPDATE Settings SET EtaLastUUID = EtaLastUUID WHERE SettingID = 1")
        : db.Database.ExecuteSqlRawAsync("UPDATE Settings SET LastInvoiceHash = LastInvoiceHash WHERE SettingID = 1");

    // ================================================================== ZATCA (modZatcaXml, modZatcaApi)
    /// <summary>modZatcaApi.ZatcaReadyProblem: the device is onboarded in the current environment and its key file is here.</summary>
    public static string ZatcaReadyProblem(Settings s)
    {
        if (Txt(s.ZatcaOnboardStage) != "PRODUCTION" || Txt(s.ZatcaProductionToken).Length == 0)
        {
            return "الجهاز لم يُسجَّل بعد لدى الهيئة (إعداد الربط: الخطوات 1 و2 و3).";
        }
        if (Txt(s.ZatcaOnboardEnv) != Txt(s.EInvoiceEnvironment))
        {
            return "سُجّل الجهاز في بيئة أخرى. سجّله في البيئة الحالية من إعداد الربط.";
        }
        if (Txt(s.ZatcaKeyFile).Length == 0)
        {
            return "ملف المفتاح الخاص غير محدد في إعداد الربط.";
        }
        if (!File.Exists(Txt(s.ZatcaKeyFile)))
        {
            return "ملف المفتاح الخاص غير موجود على هذا الجهاز: " + Txt(s.ZatcaKeyFile);
        }
        return "";
    }

    /// <summary>For the service's --check: Saudi Arabia signs a test hash with the key file of the settings and checks it with the
    /// certificate of the device. "" = it works (Egypt signs with its signing program: nothing to check here).</summary>
    public async Task<string> CheckSigningAsync()
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        if (CountryRules.Normalize(s.CountryCode) == "EG" || !File.Exists(Txt(s.ZatcaKeyFile)))
        {
            return "";
        }
        try
        {
            string hash = Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData([1, 2, 3]));
            string signature = ZatcaDocument.SignHash(hash, await File.ReadAllTextAsync(Txt(s.ZatcaKeyFile)));
            string cert = ZatcaDocument.CertificateBody(s.ZatcaCertificate);
            if (cert.Length == 0)
            {
                return "لا توجد شهادة الجهاز (CSID) في إعداد الربط.";
            }
            using var x509 = System.Security.Cryptography.X509Certificates.X509CertificateLoader.LoadCertificate(Convert.FromBase64String(cert));
            using var key = x509.GetECDsaPublicKey();
            bool ok = key is not null && key.VerifyData(Convert.FromBase64String(hash), Convert.FromBase64String(signature),
                                                        System.Security.Cryptography.HashAlgorithmName.SHA256,
                                                        System.Security.Cryptography.DSASignatureFormat.Rfc3279DerSequence);
            return ok ? "" : "Signature does not match the certificate.";
        }
        catch (Exception e) when (e is System.Security.Cryptography.CryptographicException or FormatException or ArgumentException)
        {
            return "تعذّر التوقيع ببرنامج OpenSSL: " + e.Message;
        }
    }

    /// <summary>modZatcaXml.ZatcaSellerProblem: the lines of the seller's missing data, "" when complete.</summary>
    public static string ZatcaSellerProblem(Settings s)
    {
        var m = new StringBuilder();
        void Need(string line) => m.Append(line).Append(NL);
        if (CountryRules.TaxNumberProblem(s.VATNumber, "SA").Length > 0 || Txt(s.VATNumber).Length == 0)
        {
            Need("- الرقم الضريبي للمحل");
        }
        if (Txt(s.StoreName).Length == 0)
        {
            Need("- اسم المحل");
        }
        if (Txt(s.StreetName).Length == 0)
        {
            Need("- شارع المحل");
        }
        if (!Digits(Txt(s.BuildingNo), 4))
        {
            Need("- رقم مبنى المحل (4 أرقام)");
        }
        if (!Digits(Txt(s.PostalCode), 5))
        {
            Need("- الرمز البريدي للمحل (5 أرقام)");
        }
        if (Txt(s.City).Length == 0)
        {
            Need("- مدينة المحل");
        }
        if (Txt(s.District).Length == 0)
        {
            Need("- حي المحل");
        }
        return m.ToString();
    }

    /// <summary>modZatcaXml.ZatcaDataProblem: the seller's and, for a tax invoice, the buyer's data; standard-rated lines.</summary>
    public async Task<string> ZatcaDataProblemAsync(string kind, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        var m = new StringBuilder(ZatcaSellerProblem(s));
        void Need(string line) => m.Append(line).Append(NL);
        if (h.SubType == "STANDARD" && await db.Customers.AsNoTracking().FirstOrDefaultAsync(c => c.CustomerID == h.CustomerID) is { } c)
        {
            if (Txt(c.VATNumber).Length == 0 || CountryRules.TaxNumberProblem(c.VATNumber, "SA").Length > 0)
            {
                Need("- الرقم الضريبي للعميل");
            }
            if (Txt(c.StreetName).Length == 0)
            {
                Need("- شارع العميل");
            }
            if (!Digits(Txt(c.BuildingNo), 4))
            {
                Need("- رقم مبنى العميل (4 أرقام)");
            }
            if (!Digits(Txt(c.PostalCode), 5))
            {
                Need("- الرمز البريدي للعميل (5 أرقام)");
            }
            if (Txt(c.City).Length == 0)
            {
                Need("- مدينة العميل");
            }
            if (Txt(c.District).Length == 0)
            {
                Need("- حي العميل");
            }
        }
        if ((await LinesAsync(db, kind, id)).Any(l => l.VATCategory != "S"))
        {
            Need("- أصناف معفاة أو بنسبة صفرية: تحتاج سبب الإعفاء، ولم يُدعم بعد");
        }
        return m.Length > 0 ? "بيانات ناقصة للفاتورة الإلكترونية:" + NL + m : "";
    }

    /// <summary>UN/ECE 4461: 10 cash, 30 credit transfer (sale on account), 42 bank account, 48 bank card.</summary>
    public static string ZatcaPaymentMeans(string payType, int? methodId) =>
        payType == "CREDIT" ? "30" : (methodId ?? 1) switch { 2 => "48", 3 => "42", _ => "10" };

    private static ZatcaParty SellerParty(Settings s) =>
        new(Txt(s.StoreName), Txt(s.VATNumber), Txt(s.CRNumber), Txt(s.StreetName), Txt(s.BuildingNo), Txt(s.AdditionalNo),
            Txt(s.District), Txt(s.City), Txt(s.PostalCode), "SA");

    /// <summary>modZatcaXml.ZatcaDocumentXml: the unsigned document of a sales invoice or return, and the totals of its QR code.</summary>
    public async Task<(string Xml, (string Timestamp, decimal Total, decimal Vat, string Seller, string SellerVat) Totals)> ZatcaDocumentXmlAsync(
        string kind, int id, string pih)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await ZatcaDocumentXmlAsync(db, kind, id, pih);
    }

    private static async Task<(string Xml, (string Timestamp, decimal Total, decimal Vat, string Seller, string SellerVat) Totals)> ZatcaDocumentXmlAsync(
        RetailDbContext db, string kind, int id, string pih)
    {
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        string issueDate = h.Date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        string issueTime = h.Date.ToString("HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
        bool standard = h.SubType == "STANDARD";
        ZatcaParty? buyer = null;
        if (standard)
        {
            var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == h.CustomerID);
            buyer = new ZatcaParty(Txt(c.CustomerName), Txt(c.VATNumber), "", Txt(c.StreetName), Txt(c.BuildingNo), "", Txt(c.District),
                                   Txt(c.City), Txt(c.PostalCode), "SA");
        }
        var rows = await LinesAsync(db, kind, id);
        var lines = rows.Select((r, i) => new ZatcaLine(i + 1, r.Quantity, r.ZatcaUnitCode ?? "PCE", r.NetAmount, r.Discount, r.Tax,
                                                        r.LineTotal, Txt(r.ProductName), r.VATCategory, r.VATRate * 100, r.UnitPrice)).ToList();
        var subtotals = rows.GroupBy(r => (r.VATCategory, r.VATRate)).OrderBy(g => g.Key.VATCategory, StringComparer.Ordinal)
            .ThenBy(g => g.Key.VATRate)
            .Select(g => new ZatcaTaxSubtotal(g.Sum(r => r.NetAmount), g.Sum(r => r.Tax), g.Key.VATCategory, g.Key.VATRate * 100)).ToList();
        string reason = "";
        if (kind == "RETURN")
        {
            reason = Txt(h.Reason).Length > 0 ? Txt(h.Reason) : "Return";
        }
        var doc = new ZatcaInvoice(h.Number, Txt(h.Uuid), issueDate, issueTime, h.TypeCode, standard ? "0100000" : "0200000", "SAR",
                                   kind == "RETURN" ? Txt(h.OriginalNumber) : "", h.Icv ?? 0, pih, SellerParty(s), buyer,
                                   standard ? issueDate : "", ZatcaPaymentMeans(h.PayType, h.MethodId), reason, h.Tax, subtotals,
                                   h.Taxable, h.Taxable, h.Total, lines);
        return (ZatcaDocument.InvoiceXml(doc), (issueDate + "T" + issueTime, h.Total, h.Tax, Txt(s.StoreName), Txt(s.VATNumber)));
    }

    /// <summary>modZatcaXml.ZatcaPrepareDocument: signs a document once, in the order of the chain (its previous hash is the last
    /// hash of the settings, then its own hash becomes the last one). "" or the problem; a signed document is left as it is.</summary>
    public async Task<string> ZatcaPrepareAsync(string kind, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        if (!string.IsNullOrEmpty(h.Xml))
        {
            return "";
        }
        string problem = await ZatcaDataProblemAsync(kind, id);
        if (problem.Length > 0)
        {
            return problem;
        }
        var s = await SettingsReader.ReadAsync(db);
        string cert = ZatcaDocument.CertificateBody(s.ZatcaCertificate);
        if (cert.Length == 0)
        {
            return "لا توجد شهادة الجهاز (CSID) في إعداد الربط.";
        }
        string keyFile = Txt(s.ZatcaKeyFile);
        if (keyFile.Length == 0 || !File.Exists(keyFile))
        {
            return "ملف المفتاح الخاص غير موجود: " + keyFile;
        }
        string key = await File.ReadAllTextAsync(keyFile);
        try
        {
            await using var tx = await db.Database.BeginTransactionAsync();
            await LockChainAsync(db, "LastInvoiceHash");
            string pih = await db.Settings.Where(x => x.SettingID == 1).Select(x => x.LastInvoiceHash).FirstOrDefaultAsync()
                         ?? ZatcaDocument.InitialPih;
            if (pih.Length == 0)
            {
                pih = ZatcaDocument.InitialPih;
            }
            var (xml, totals) = await ZatcaDocumentXmlAsync(db, kind, id, pih);
            string signed, hash, qr;
            try
            {
                (signed, hash, qr) = ZatcaDocument.Sign(xml, cert, key, ZatcaDocument.SigningTime(_now()), totals);
            }
            catch (System.Security.Cryptography.CryptographicException e)
            {
                return "تعذّر التوقيع ببرنامج OpenSSL: " + e.Message;
            }
            await SaveDocAsync(db, kind, id,
                x => { x.InvoiceHash = hash; x.PreviousInvoiceHash = pih; x.QRCodeData = qr; x.EInvoiceXml = signed; },
                x => { x.InvoiceHash = hash; x.PreviousInvoiceHash = pih; x.QRCodeData = qr; x.EInvoiceXml = signed; });
            await db.Settings.Where(x => x.SettingID == 1).ExecuteUpdateAsync(u => u.SetProperty(x => x.LastInvoiceHash, hash));
            await tx.CommitAsync();
            return "";
        }
        catch (Exception e) when (e is DbUpdateException or InvalidOperationException or FormatException)
        {
            return "تعذّر توقيع المستند: " + e.Message;
        }
    }

    /// <summary>modZatcaApi.ZatcaSendDocument: signs the document if it is not yet, then reports it (simplified) or clears it
    /// (standard); a cleared document carries ZATCA's stamp and becomes the invoice.</summary>
    public async Task<EInvoiceReply> ZatcaSendAsync(string kind, int id)
    {
        string problem = await ZatcaPrepareAsync(kind, id);
        if (problem.Length > 0)
        {
            return new("ERROR", "", problem);
        }
        Settings s;
        Head h;
        await using (var db = await factory.CreateDbContextAsync())
        {
            s = await SettingsReader.ReadAsync(db);
            h = (await HeadAsync(db, kind, id))!;
        }
        bool standard = h.SubType == "STANDARD";
        string body = "{\"invoiceHash\": \"" + h.Hash + "\", \"uuid\": \"" + h.Uuid + "\", \"invoice\": \"" +
                      Convert.ToBase64String(Encoding.UTF8.GetBytes(h.Xml ?? "")) + "\"}";
        string path = standard ? "/invoices/clearance/single" : "/invoices/reporting/single";
        var headers = new List<(string, string)>
        {
            ("Accept-Version", "V2"), ("Accept-Language", "ar"), ("Accept", "application/json"),
            ("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(Txt(s.ZatcaProductionToken) + ":" +
                                                                                        Txt(s.ZatcaProductionSecret)))),
        };
        if (standard)
        {
            headers.Add(("Clearance-Status", "1"));
        }
        var (status, reply, ms, _) = await HttpAsync(HttpMethod.Post, ZatcaReplies.BaseUrl(Txt(s.EInvoiceEnvironment)) + path, headers,
                                                     Json(body), 30);
        var result = ZatcaReplies.Result(status, reply, standard);
        await LogAsync(kind, id, standard ? "CLEAR" : "REPORT", path, status, result.Result, result.Message, body, reply, ms);
        if (result.DocStatus.Length > 0)
        {
            await SetStatusAsync(kind, id, result.DocStatus, reply);
        }
        string? cleared = JsonPath.Text(JsonPath.Parse(reply), "clearedInvoice");
        if (standard && result.Sent && !string.IsNullOrEmpty(cleared))
        {
            string xml = Encoding.UTF8.GetString(Convert.FromBase64String(cleared));
            string qr = ZatcaDocument.ClearedQr(xml);
            await using var db = await factory.CreateDbContextAsync();
            await SaveDocAsync(db, kind, id,
                x => { x.EInvoiceXml = xml; if (qr.Length > 0) { x.QRCodeData = qr; } },
                x => { x.EInvoiceXml = xml; if (qr.Length > 0) { x.QRCodeData = qr; } });
        }
        return result;
    }

    // ================================================================== Egypt (modEtaReceipt, modEtaInvoice)
    private static string EtaEnvOf(Settings s) => EtaDocument.Env(Txt(s.EInvoiceEnvironment));

    /// <summary>modEtaReceipt.EtaIssuerProblem: the lines of the missing data of the seller.</summary>
    public static string EtaIssuerProblem(Settings s)
    {
        var m = new StringBuilder();
        void Need(string line) => m.Append(line).Append(NL);
        if (Txt(s.VATNumber).Length == 0 || CountryRules.TaxNumberProblem(s.VATNumber, "EG").Length > 0)
        {
            Need("- رقم التسجيل الضريبي للمحل (9 أرقام)");
        }
        if (Txt(s.StoreName).Length == 0)
        {
            Need("- اسم المحل");
        }
        if (Txt(s.EtaGovernate).Length == 0)
        {
            Need("- المحافظة");
        }
        if (Txt(s.City).Length == 0)
        {
            Need("- مدينة المحل");
        }
        if (Txt(s.StreetName).Length == 0)
        {
            Need("- شارع المحل");
        }
        if (Txt(s.BuildingNo).Length == 0)
        {
            Need("- رقم مبنى المحل");
        }
        if (!Digits(Txt(s.EtaActivityCode), 4))
        {
            Need("- كود النشاط (4 أرقام)");
        }
        return m.ToString();
    }

    /// <summary>modEtaReceipt.EtaReadyProblem: the data of the seller and the credentials of the POS or of the program.</summary>
    public static string EtaReadyProblem(Settings s)
    {
        if (CountryRules.Normalize(s.CountryCode) != "EG")
        {
            return "الإيصال الإلكتروني المصري لدولة التشغيل مصر فقط.";
        }
        bool pos = Txt(s.EtaClientId).Length > 0 && Txt(s.EtaClientSecret).Length > 0;
        bool erp = Txt(s.EtaErpClientId).Length > 0 && Txt(s.EtaErpClientSecret).Length > 0;
        string m = !pos && !erp ? "- بيانات دخول جهاز نقطة البيع (Client ID و Client Secret)" + NL : "";
        m += EtaIssuerProblem(s);
        return m.Length > 0 ? "إعداد الإيصال الإلكتروني ناقص:" + NL + m : "";
    }

    private static async Task<string> MissingItemCodesAsync(RetailDbContext db, string kind, int id)
    {
        var names = (await LinesAsync(db, kind, id)).Where(l => string.IsNullOrEmpty(l.EtaItemCode)).Select(l => Txt(l.ProductName))
            .Distinct().ToList();
        if (names.Count == 0)
        {
            return "";
        }
        return "- كود المصلحة (EGS أو GS1) للأصناف: " + string.Join("، ", names.Take(5)) + (names.Count > 5 ? " ..." : "") + NL;
    }

    /// <summary>modEtaReceipt.EtaDataProblem: seller data, the ETA code of every product, the sale of a return sent first, the ID of a
    /// person who buys for 150,000 or more.</summary>
    public async Task<string> EtaReceiptDataProblemAsync(string kind, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        string m = EtaIssuerProblem(s);
        if (Txt(s.EtaPosSerial).Length == 0)
        {
            m += "- الرقم التسلسلي لجهاز نقطة البيع" + NL;
        }
        m += await MissingItemCodesAsync(db, kind, id);
        if (kind == "RETURN" && h.OriginalEtaUuid.Length == 0)
        {
            m += "- أرسل إيصال البيع الأصلي أولًا" + NL;
        }
        var c = await db.Customers.AsNoTracking().FirstOrDefaultAsync(x => x.CustomerID == h.CustomerID);
        if (h.Total >= EtaDocument.BuyerIdLimit && Txt(c?.VATNumber).Length == 0 && Txt(c?.NationalID).Length == 0)
        {
            m += "- إيصال 150 ألف جنيه أو أكثر: الرقم القومي أو الضريبي للمشتري" + NL;
        }
        return m.Length > 0 ? "بيانات ناقصة للإيصال الإلكتروني:" + NL + m : "";
    }

    /// <summary>modEtaReceipt.EtaDocumentJson: the receipt of a sales invoice or return, header.uuid "".</summary>
    public async Task<(string Json, string DateTime, decimal Total)> EtaReceiptJsonAsync(string kind, int id, string previousUuid)
    {
        await using var db = await factory.CreateDbContextAsync();
        return await EtaReceiptJsonAsync(db, kind, id, previousUuid);
    }

    private static async Task<(string Json, string DateTime, decimal Total)> EtaReceiptJsonAsync(RetailDbContext db, string kind, int id,
                                                                                               string previousUuid)
    {
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        string dt = EtaDocument.UtcTimestamp(h.Date);
        string branch = Txt(s.EtaBranchCode).Length > 0 ? Txt(s.EtaBranchCode) : "0";
        string buyerType = "P", buyerId = "", buyerName = "", mobile = "";
        var c = await db.Customers.AsNoTracking().FirstOrDefaultAsync(x => x.CustomerID == h.CustomerID);
        if (c is not null && c.CustomerID != s.DefaultCustomerID && !c.IsSystem)
        {
            if (Txt(c.VATNumber).Length > 0)
            {
                buyerType = "B";
                buyerId = Txt(c.VATNumber);
            }
            else
            {
                buyerId = Txt(c.NationalID);
            }
            buyerName = Txt(c.CustomerName);
            mobile = Txt(c.Mobile);
        }
        var rows = await LinesAsync(db, kind, id);
        var items = rows.Select(r => new EtaReceiptItem(Txt(r.ProductCode), Txt(r.ProductName), r.EtaItemType ?? "EGS", Txt(r.EtaItemCode),
                                                        r.EtaUnitCode ?? "EA", r.Quantity, r.UnitPrice, r.NetAmount + r.Discount, r.Discount,
                                                        r.NetAmount, r.Tax, r.VATRate * 100, r.LineTotal)).ToList();
        var receipt = new EtaReceipt(dt, h.Number, previousUuid, kind == "RETURN" ? h.OriginalEtaUuid : "", Txt(s.VATNumber),
                                     Txt(s.StoreName), branch, Txt(s.EtaGovernate), Txt(s.City), Txt(s.StreetName), Txt(s.BuildingNo),
                                     Txt(s.PostalCode), Txt(s.EtaPosSerial), Txt(s.EtaActivityCode), buyerType, buyerId, buyerName, mobile,
                                     items, rows.Sum(r => r.NetAmount + r.Discount), rows.Sum(r => r.Discount), h.Taxable, h.Tax, h.Total,
                                     EtaDocument.PaymentCode(h.PayType, h.MethodId));
        return (EtaDocument.ReceiptJson(receipt), dt, h.Total);
    }

    /// <summary>modEtaReceipt.EtaPrepareDocument: builds the receipt once in the order of the chain; a rejected receipt is built
    /// again on the same previous UUID. "" or the problem.</summary>
    public async Task<string> EtaReceiptPrepareAsync(string kind, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        string oldUuid = Txt(h.EtaUuid);
        bool rebuild = oldUuid.Length > 0 && h.Status == "REJECTED";
        if (oldUuid.Length > 0 && !rebuild)
        {
            return "";
        }
        try
        {
            await using var tx = await db.Database.BeginTransactionAsync();
            await LockChainAsync(db, "EtaLastUUID");
            var s = await db.Settings.AsNoTracking().FirstAsync(x => x.SettingID == 1);
            string prev = rebuild ? Txt(h.EtaPrevious) : Txt(s.EtaLastUUID);
            var (json, dt, total) = await EtaReceiptJsonAsync(db, kind, id, prev);
            string uuid = EtaDocument.ReceiptUuid(json);
            json = EtaDocument.WithUuid(json, uuid);
            string qr = EtaDocument.ReceiptQrLink(EtaEnvOf(s), uuid, dt, total, Txt(s.VATNumber));
            await SaveDocAsync(db, kind, id,
                x => { x.EtaUUID = uuid; x.EtaPreviousUUID = prev.Length == 0 ? null : prev; x.QRCodeData = qr; x.EInvoiceXml = json; },
                x => { x.EtaUUID = uuid; x.EtaPreviousUUID = prev.Length == 0 ? null : prev; x.QRCodeData = qr; x.EInvoiceXml = json; });
            if (rebuild)
            {
                await db.Settings.Where(x => x.SettingID == 1 && x.EtaLastUUID == oldUuid)
                    .ExecuteUpdateAsync(u => u.SetProperty(x => x.EtaLastUUID, uuid));
            }
            else
            {
                await db.Settings.Where(x => x.SettingID == 1).ExecuteUpdateAsync(u => u.SetProperty(x => x.EtaLastUUID, uuid));
            }
            await tx.CommitAsync();
            return "";
        }
        catch (Exception e) when (e is DbUpdateException or InvalidOperationException)
        {
            return "تعذّر بناء الإيصال: " + e.Message;
        }
    }

    private async Task<(string Token, EInvoiceReply? Problem)> TokenAsync(Settings s, bool erp)
    {
        string env = EtaEnvOf(s);
        DateTime now = _now();
        if (erp ? _erpToken.Length > 0 && _erpTokenEnv == env && now < _erpTokenExpires
                : _posToken.Length > 0 && _posTokenEnv == env && now < _posTokenExpires)
        {
            return (erp ? _erpToken : _posToken, null);
        }
        var headers = new List<(string, string)> { ("Accept", "application/json") };
        string id = erp ? Txt(s.EtaErpClientId) : Txt(s.EtaClientId);
        string secret = erp ? Txt(s.EtaErpClientSecret) : Txt(s.EtaClientSecret);
        if (!erp)
        {
            headers.Add(("posserial", Txt(s.EtaPosSerial)));
            headers.Add(("pososversion", Txt(s.EtaPosOsVersion).Length > 0 ? Txt(s.EtaPosOsVersion) : "Windows"));
            if (Txt(s.EtaPreSharedKey).Length > 0)
            {
                headers.Add(("presharedkey", Txt(s.EtaPreSharedKey)));
            }
        }
        string form = "grant_type=client_credentials&client_id=" + EtaDocument.UrlEncode(id) + "&client_secret=" +
                      EtaDocument.UrlEncode(secret);
        string url = EtaDocument.TokenUrl(env);
        var (status, reply, ms, msg) = await HttpAsync(HttpMethod.Post, url, headers,
                                                       new StringContent(form, Encoding.UTF8, "application/x-www-form-urlencoded"), 30);
        var root = JsonPath.Parse(reply);
        string? token = JsonPath.Text(root, "access_token");
        string result;
        if (status == 200 && token is not null)
        {
            result = "OK";
            int life = int.TryParse(JsonPath.Text(root, "expires_in"), out var l) ? l : 3600;
            if (erp)
            {
                (_erpToken, _erpTokenEnv, _erpTokenExpires) = (token, env, now.AddSeconds(life - 60));
            }
            else
            {
                (_posToken, _posTokenEnv, _posTokenExpires) = (token, env, now.AddSeconds(life - 60));
            }
        }
        else if (status == 0 || status == 429 || status >= 500)
        {
            result = "NETWORK";
            if (msg.Length == 0)
            {
                msg = "المنظومة غير متاحة الآن (رمز " + status + ").";
            }
        }
        else
        {
            result = "ERROR";
            msg = (erp ? "رفضت المصلحة بيانات دخول البرنامج (رمز " : "رفضت المصلحة بيانات جهاز نقطة البيع (رمز ") + status + "): " +
                  JsonPath.First(root, "error_description", "error");
        }
        // the request body carries the client secret: it is not logged
        string masked = token is { Length: > 0 } ? reply.Replace(token, "***") : reply;
        await LogAsync("", 0, erp ? "TOKEN_ERP" : "TOKEN", url, status, result, msg, "", masked, ms);
        return result == "OK" ? (token!, null) : ("", new EInvoiceReply(result, "", msg));
    }

    private void ForgetToken(bool erp)
    {
        if (erp)
        {
            _erpToken = "";
        }
        else
        {
            _posToken = "";
        }
    }

    /// <summary>modEtaInvoice.EtaIsInvoice: a tax invoice (customer with a tax number) and its returns are e-invoices.</summary>
    private static bool IsInvoice(Head h) => h.SubType == "STANDARD";

    /// <summary>modEtaReceipt.EtaSendDocument: builds the receipt if it is not yet, then submits it; an e-invoice goes to
    /// EtaInvoiceSendAsync.</summary>
    public async Task<EInvoiceReply> EtaSendAsync(string kind, int id)
    {
        Head h;
        Settings s;
        await using (var db = await factory.CreateDbContextAsync())
        {
            h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
            s = await SettingsReader.ReadAsync(db);
        }
        if (IsInvoice(h))
        {
            return await EtaInvoiceSendAsync(kind, id);
        }
        string problem = await EtaReceiptDataProblemAsync(kind, id);
        if (problem.Length == 0)
        {
            problem = await EtaReceiptPrepareAsync(kind, id);
        }
        if (problem.Length > 0)
        {
            return new("ERROR", "", problem);
        }
        var (token, tokenProblem) = await TokenAsync(s, false);
        if (tokenProblem is not null)
        {
            return tokenProblem;
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            h = (await HeadAsync(db, kind, id))!;
        }
        string body = "{\"receipts\":[" + h.Xml + "]}";
        string url = EtaDocument.ApiUrl(EtaEnvOf(s)) + "/api/v1/receiptsubmissions";
        var (status, reply, ms, _) = await HttpAsync(HttpMethod.Post, url,
            [("Accept", "application/json"), ("Authorization", "Bearer " + token)], Json(body), 60);
        var result = EtaDocument.ReceiptSubmitResult(status, reply);
        if (status == 401)
        {
            ForgetToken(false);                     // expired: a new token next time
        }
        await LogAsync(kind, id, "SUBMIT", url, status, result.Result, result.Message, body, reply, ms);
        if (result.DocStatus.Length > 0)
        {
            await SetStatusAsync(kind, id, result.DocStatus, reply);
        }
        if (result.SubmissionId.Length > 0)
        {
            await using var db = await factory.CreateDbContextAsync();
            await SaveDocAsync(db, kind, id, x => x.EtaSubmissionId = result.SubmissionId, x => x.EtaSubmissionId = result.SubmissionId);
        }
        return result;
    }

    /// <summary>modEtaReceipt.EtaRefreshStatus: VALID / INVALID of a submitted receipt (an e-invoice: EtaInvoiceRefreshAsync).</summary>
    public async Task<EInvoiceReply> EtaRefreshAsync(string kind, int id)
    {
        Head h;
        Settings s;
        await using (var db = await factory.CreateDbContextAsync())
        {
            h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
            s = await SettingsReader.ReadAsync(db);
        }
        if (IsInvoice(h))
        {
            return await EtaInvoiceRefreshAsync(kind, id);
        }
        if (Txt(h.SubmissionId).Length == 0)
        {
            return new("ERROR", "", "لم يُرسل الإيصال بعد.");
        }
        var (token, tokenProblem) = await TokenAsync(s, false);
        if (tokenProblem is not null)
        {
            return tokenProblem;
        }
        string url = EtaDocument.ApiUrl(EtaEnvOf(s)) + "/api/v1/receiptsubmissions/" + h.SubmissionId + "/details?PageNo=1&PageSize=10";
        var (status, reply, ms, msg) = await HttpAsync(HttpMethod.Get, url,
            [("Accept", "application/json"), ("Authorization", "Bearer " + token)], null, 30);
        string result;
        if (status == 200)
        {
            var (st, m) = EtaDocument.ReceiptStatusResult(reply);
            msg = m;
            result = st == "INVALID" ? "REJECTED" : "OK";
            if (st.Length > 0)
            {
                await SetStatusAsync(kind, id, st, reply);
            }
        }
        else if (status == 0 || status == 429 || status >= 500)
        {
            result = "NETWORK";
            if (msg.Length == 0)
            {
                msg = "المنظومة غير متاحة الآن (رمز " + status + ").";
            }
        }
        else
        {
            if (status == 401)
            {
                ForgetToken(false);
            }
            result = "ERROR";
            msg = "رد غير متوقع من المصلحة (رمز " + status + "): " + Left(reply, 200);
        }
        await LogAsync(kind, id, "STATUS", url, status, result, msg, "", reply, ms);
        return new(result, "", msg);
    }

    /// <summary>modEtaInvoice.EtaInvoiceDataProblem: the issuer, the receiver, the ETA code of every product, the invoice of a credit
    /// note sent first, the credentials of the program and the signing program.</summary>
    public async Task<string> EtaInvoiceDataProblemAsync(string kind, int id)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        var m = new StringBuilder(EtaIssuerProblem(s));
        void Need(string line) => m.Append(line).Append(NL);
        if (Txt(s.EtaErpClientId).Length == 0 || Txt(s.EtaErpClientSecret).Length == 0)
        {
            Need("- بيانات دخول البرنامج للفاتورة الإلكترونية (Client ID و Client Secret)");
        }
        if (EtaEnvOf(s) == "PROD" && Txt(s.EtaSignerPath).Length == 0)
        {
            Need("- برنامج التوقيع (فلاشة التوقيع) في إعداد الربط");
        }
        if (await db.Customers.AsNoTracking().FirstOrDefaultAsync(x => x.CustomerID == h.CustomerID) is { } c)
        {
            if (Txt(c.VATNumber).Length == 0 || CountryRules.TaxNumberProblem(c.VATNumber, "EG").Length > 0)
            {
                Need("- رقم التسجيل الضريبي للعميل (9 أرقام)");
            }
            if (Txt(c.City).Length == 0)
            {
                Need("- مدينة العميل");
            }
            if (Txt(c.StreetName).Length == 0)
            {
                Need("- شارع العميل");
            }
            if (Txt(c.BuildingNo).Length == 0)
            {
                Need("- رقم مبنى العميل");
            }
        }
        m.Append(await MissingItemCodesAsync(db, kind, id));
        if (kind == "RETURN" && h.OriginalEtaUuid.Length == 0)
        {
            Need("- أرسل الفاتورة الأصلية أولًا");
        }
        return m.Length > 0 ? "بيانات ناقصة للفاتورة الإلكترونية:" + NL + m : "";
    }

    /// <summary>modEtaInvoice.EtaInvoiceDocumentJson: the unsigned document of a sales invoice (I) or of its return (C).</summary>
    public async Task<string> EtaInvoiceJsonAsync(string kind, int id, string version)
    {
        await using var db = await factory.CreateDbContextAsync();
        var s = await SettingsReader.ReadAsync(db);
        var h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        string branch = Txt(s.EtaBranchCode).Length > 0 ? Txt(s.EtaBranchCode) : "0";
        var issuer = new EtaParty("issuer", "B", Txt(s.VATNumber), Txt(s.StoreName), branch, Txt(s.EtaGovernate), Txt(s.City),
                                  Txt(s.StreetName), Txt(s.BuildingNo), Txt(s.PostalCode));
        var c = await db.Customers.AsNoTracking().FirstAsync(x => x.CustomerID == h.CustomerID);
        string region = Txt(c.District).Length > 0 ? Txt(c.District) : Txt(c.City);    // a city and a district: governate, city
        var receiver = new EtaParty("receiver", "B", Txt(c.VATNumber), Txt(c.CustomerName), "", Txt(c.City), region, Txt(c.StreetName),
                                    Txt(c.BuildingNo), Txt(c.PostalCode));
        var rows = await LinesAsync(db, kind, id);
        var lines = rows.Select(r => new EtaInvoiceLine(Txt(r.ProductName), r.EtaItemType ?? "EGS", Txt(r.EtaItemCode), r.EtaUnitCode ?? "EA",
                                                        r.Quantity, Txt(r.ProductCode), r.UnitPrice, r.NetAmount + r.Discount, r.Discount,
                                                        r.NetAmount, r.Tax, r.VATRate * 100, r.LineTotal)).ToList();
        return EtaDocument.InvoiceJson(new EtaInvoice(kind == "RETURN" ? "C" : "I", version, EtaDocument.UtcTimestamp(h.Date),
                                                      Txt(s.EtaActivityCode), h.Number, issuer, receiver, lines,
                                                      kind == "RETURN" ? h.OriginalEtaUuid : "", rows.Sum(r => r.Discount),
                                                      rows.Sum(r => r.NetAmount + r.Discount), h.Taxable, h.Tax, h.Total));
    }

    /// <summary>modEtaInvoice.EtaInvoicePrepare: builds (and signs) the document once; a rejected or invalid one is built again.</summary>
    public async Task<string> EtaInvoicePrepareAsync(string kind, int id)
    {
        Head h;
        Settings s;
        await using (var db = await factory.CreateDbContextAsync())
        {
            h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
            s = await SettingsReader.ReadAsync(db);
        }
        if (!string.IsNullOrEmpty(h.Xml) && h.Status != "REJECTED" && h.Status != "INVALID")
        {
            return "";
        }
        string json;
        if (EtaEnvOf(s) == "PROD" || Txt(s.EtaSignerPath).Length > 0)
        {
            json = await EtaInvoiceJsonAsync(kind, id, "1.0");
            var (signature, problem) = await _signer.SignAsync(EtaDocument.Serialize(json), Txt(s.EtaSignerPath), Txt(s.EtaSignerArgs),
                                                               Txt(s.EtaTokenPin));
            if (problem.Length > 0)
            {
                return problem;
            }
            json = EtaDocument.WithSignature(json, signature);
        }
        else
        {
            json = await EtaInvoiceJsonAsync(kind, id, "0.9");
        }
        await using (var db = await factory.CreateDbContextAsync())
        {
            await SaveDocAsync(db, kind, id, x => x.EInvoiceXml = json, x => x.EInvoiceXml = json);
        }
        return "";
    }

    /// <summary>modEtaInvoice.EtaInvoiceSend: builds and signs the document if needed, then submits it.</summary>
    public async Task<EInvoiceReply> EtaInvoiceSendAsync(string kind, int id)
    {
        string problem = await EtaInvoiceDataProblemAsync(kind, id);
        if (problem.Length == 0)
        {
            problem = await EtaInvoicePrepareAsync(kind, id);
        }
        if (problem.Length > 0)
        {
            return new("ERROR", "", problem);
        }
        Settings s;
        Head h;
        await using (var db = await factory.CreateDbContextAsync())
        {
            s = await SettingsReader.ReadAsync(db);
            h = (await HeadAsync(db, kind, id))!;
        }
        var (token, tokenProblem) = await TokenAsync(s, true);
        if (tokenProblem is not null)
        {
            return tokenProblem;
        }
        string body = "{\"documents\":[" + h.Xml + "]}";
        string url = EtaDocument.ApiUrl(EtaEnvOf(s)) + "/api/v1.0/documentsubmissions";
        var (status, reply, ms, _) = await HttpAsync(HttpMethod.Post, url,
            [("Accept", "application/json"), ("Authorization", "Bearer " + token)], Json(body), 60);
        var result = EtaDocument.InvoiceSubmitResult(status, reply);
        if (status == 401)
        {
            ForgetToken(true);
        }
        await LogAsync(kind, id, "SUBMIT_DOC", url, status, result.Result, result.Message, body, reply, ms);
        if (result.DocStatus.Length > 0)
        {
            await SetStatusAsync(kind, id, result.DocStatus, reply);
        }
        if (result.Uuid.Length > 0)
        {
            string qr = EtaDocument.InvoiceQrLink(EtaEnvOf(s), result.Uuid, result.LongId);
            string? longId = result.LongId.Length == 0 ? null : result.LongId;
            string? subId = result.SubmissionId.Length == 0 ? null : result.SubmissionId;
            await using var db = await factory.CreateDbContextAsync();
            await SaveDocAsync(db, kind, id,
                x => { x.EtaUUID = result.Uuid; x.EtaLongId = longId; x.EtaSubmissionId = subId; x.QRCodeData = qr; },
                x => { x.EtaUUID = result.Uuid; x.EtaLongId = longId; x.EtaSubmissionId = subId; x.QRCodeData = qr; });
        }
        return result;
    }

    /// <summary>modEtaInvoice.EtaInvoiceRefresh: VALID / INVALID / CANCELLED of a submitted document; still in progress = OK.</summary>
    public async Task<EInvoiceReply> EtaInvoiceRefreshAsync(string kind, int id)
    {
        Settings s;
        Head h;
        await using (var db = await factory.CreateDbContextAsync())
        {
            s = await SettingsReader.ReadAsync(db);
            h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        }
        if (Txt(h.EtaUuid).Length == 0)
        {
            return new("ERROR", "", "لم يُرسل المستند بعد.");
        }
        var (token, tokenProblem) = await TokenAsync(s, true);
        if (tokenProblem is not null)
        {
            return tokenProblem;
        }
        string url = EtaDocument.ApiUrl(EtaEnvOf(s)) + "/api/v1.0/documents/" + h.EtaUuid + "/details";
        var (status, reply, ms, msg) = await HttpAsync(HttpMethod.Get, url,
            [("Accept", "application/json"), ("Authorization", "Bearer " + token)], null, 30);
        string result;
        if (status == 200)
        {
            var (st, m) = EtaDocument.InvoiceStatusResult(reply);
            msg = m;
            result = st == "INVALID" ? "REJECTED" : "OK";
            if (st.Length > 0)
            {
                await SetStatusAsync(kind, id, st, reply);
            }
        }
        else if (status == 0 || status == 429 || status >= 500)
        {
            result = "NETWORK";
            if (msg.Length == 0)
            {
                msg = "المنظومة غير متاحة الآن (رمز " + status + ").";
            }
        }
        else
        {
            if (status == 401)
            {
                ForgetToken(true);
            }
            result = "ERROR";
            msg = "رد غير متوقع من المصلحة (رمز " + status + "): " + Left(reply, 200);
        }
        await LogAsync(kind, id, "STATUS_DOC", url, status, result, msg, "", Left(reply, 60000), ms);
        return new(result, "", msg);
    }

    /// <summary>modEtaInvoice.EtaCancelDocument: cancels a valid e-invoice at ETA.</summary>
    public async Task<EInvoiceReply> EtaCancelAsync(string kind, int id, string reason)
    {
        Settings s;
        Head h;
        await using (var db = await factory.CreateDbContextAsync())
        {
            s = await SettingsReader.ReadAsync(db);
            h = await HeadAsync(db, kind, id) ?? throw new InvalidOperationException(kind + " " + id);
        }
        if (!IsInvoice(h))
        {
            return new("ERROR", "", "الإلغاء للفاتورة الإلكترونية فقط. الإيصال يُصحَّح بمرتجع.");
        }
        if (h.Status != "VALID")
        {
            return new("ERROR", "", "يُلغى المستند الصالح فقط.");
        }
        var (token, tokenProblem) = await TokenAsync(s, true);
        if (tokenProblem is not null)
        {
            return tokenProblem;
        }
        string url = EtaDocument.ApiUrl(EtaEnvOf(s)) + "/api/v1.0/documents/state/" + h.EtaUuid + "/state";
        string body = "{" + EtaDocument.JStr("status", "cancelled") + "," + EtaDocument.JStr("reason", reason) + "}";
        var (status, reply, ms, msg) = await HttpAsync(HttpMethod.Put, url,
            [("Accept", "application/json"), ("Authorization", "Bearer " + token)], Json(body), 30);
        string result;
        if (status is 200 or 202 or 204)
        {
            result = "OK";
            await SetStatusAsync(kind, id, "CANCELLED", reply);
        }
        else if (status == 0 || status == 429 || status >= 500)
        {
            result = "NETWORK";
            if (msg.Length == 0)
            {
                msg = "المنظومة غير متاحة الآن (رمز " + status + ").";
            }
        }
        else
        {
            if (status == 401)
            {
                ForgetToken(true);
            }
            result = "ERROR";
            msg = EtaDocument.Errors(JsonPath.Parse(reply), "error");
            if (msg.Length == 0)
            {
                msg = "رفضت المصلحة الإلغاء (رمز " + status + "): " + Left(reply, 200);
            }
        }
        await LogAsync(kind, id, "CANCEL", url, status, result, msg, body, reply, ms);
        return new(result, "", msg);
    }
}
