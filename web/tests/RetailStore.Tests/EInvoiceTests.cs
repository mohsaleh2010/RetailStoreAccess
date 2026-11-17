using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json.Nodes;
using Microsoft.EntityFrameworkCore;
using RetailStore.Data.Entities;
using RetailStore.Data.Journal;
using RetailStore.Data.Services;
using RetailStore.Domain.EInvoicing;

namespace RetailStore.Tests;

/// <summary>E-invoicing in C# (plan 7A, docs/85): the documents, hashes, certificate fields, QR codes and the readings of the
/// replies are those of the Python mirrors of modZatcaXml / modZatcaApi / modEtaReceipt / modEtaInvoice (EInvoiceVectors).</summary>
public class EInvoiceDocumentTests
{
    [Fact]
    public void The_zatca_documents_and_their_hashes_are_those_of_the_mirror()
    {
        foreach (var (doc, xml, hash) in EInvoiceVectors.ZatcaDocs)
        {
            Assert.Equal(xml, ZatcaDocument.InvoiceXml(doc));
            Assert.Equal(hash, ZatcaDocument.InvoiceHash(xml));
        }
        foreach (var (value, amount, price, quantity) in EInvoiceVectors.ZatcaValues)
        {
            Assert.Equal((amount, price, quantity), (ZatcaDocument.Amount(value), ZatcaDocument.Price(value), ZatcaDocument.Quantity(value)));
        }
    }

    [Fact]
    public void The_certificate_the_signature_block_and_the_qr_are_those_of_the_mirror()
    {
        var cert = ZatcaDocument.CertificateInfo(EInvoiceVectors.TestCert);
        Assert.Equal(EInvoiceVectors.CertIssuer, cert.Issuer);
        Assert.Equal(EInvoiceVectors.CertSerial, cert.Serial);
        Assert.Equal(EInvoiceVectors.CertPublicKey, Convert.ToBase64String(cert.PublicKey));
        Assert.Equal(EInvoiceVectors.CertSignature, Convert.ToBase64String(cert.Signature));
        Assert.Equal(EInvoiceVectors.CertHash, cert.Hash);
        Assert.Equal(EInvoiceVectors.SignedPropertiesDigest,
                     ZatcaDocument.HexDigestB64(Encoding.UTF8.GetBytes(ZatcaDocument.SignedProperties(EInvoiceVectors.SigningTime, cert, true))));

        var (doc, xml, hash) = EInvoiceVectors.ZatcaDocs[0];
        string qr = ZatcaDocument.QrCode(doc.Seller.Name, doc.Seller.Vat, "2026-10-08T14:05:09", 115m, 15m, hash, EInvoiceVectors.FakeSignature,
                                         cert.PublicKey, cert.Signature);
        Assert.Equal(EInvoiceVectors.SampleQr, qr);
        string ext = ZatcaDocument.UblExtensions(hash, EInvoiceVectors.FakeSignature, EInvoiceVectors.TestCert, EInvoiceVectors.SigningTime, cert);
        string signed = ZatcaDocument.SignedDocument(xml, ext, qr);
        Assert.Equal(EInvoiceVectors.SampleSigned, signed);
        Assert.Equal(hash, ZatcaDocument.InvoiceHash(signed));              // the signature block and the QR are not hashed
        Assert.Equal(qr, ZatcaDocument.ClearedQr(signed));
        Assert.Equal(EInvoiceVectors.TestCert, ZatcaDocument.CertificateBody(
            "-----BEGIN CERTIFICATE-----\n" + EInvoiceVectors.TestCert[..64] + "\r\n" + EInvoiceVectors.TestCert[64..] + "\n-----END CERTIFICATE-----"));
    }

    [Fact]
    public void The_device_key_signs_the_hash_as_openssl_does()
    {
        var (_, xml, hash) = EInvoiceVectors.ZatcaDocs[1];
        string signature = ZatcaDocument.SignHash(hash, EInvoiceVectors.TestKey);
        using var cert = X509CertificateLoader.LoadCertificate(Convert.FromBase64String(EInvoiceVectors.TestCert));
        using var key = cert.GetECDsaPublicKey()!;
        // openssl dgst -sha256 -sign: SHA-256 of the 32 bytes of the hash, DER signature
        Assert.True(key.VerifyData(Convert.FromBase64String(hash), Convert.FromBase64String(signature), HashAlgorithmName.SHA256,
                                   DSASignatureFormat.Rfc3279DerSequence));

        var (signed, h, qr) = ZatcaDocument.Sign(xml, "-----BEGIN CERTIFICATE-----\n" + EInvoiceVectors.TestCert + "\n-----END CERTIFICATE-----",
                                                 EInvoiceVectors.TestKey, ZatcaDocument.SigningTime(new DateTime(2026, 10, 8, 14, 5, 10)),
                                                 ("2026-10-08T14:05:09", 149.5m, 19.5m, "Seller", "399999999900003"));
        Assert.Equal(hash, h);
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Invoice ", signed);
        Assert.Contains("<xades:SigningTime>2026-10-08T11:05:10</xades:SigningTime>", signed);
        Assert.Contains("<ds:X509Certificate>" + EInvoiceVectors.TestCert + "</ds:X509Certificate>", signed);
        Assert.Contains(qr, signed);
        Assert.DoesNotContain(ZatcaDocument.QrPlaceholder, signed);
        var tags = Convert.FromBase64String(qr);
        Assert.Equal((byte)1, tags[0]);
    }

    [Fact]
    public void The_zatca_replies_are_read_as_access_reads_them()
    {
        foreach (var (status, body, standard, result, docStatus, message) in EInvoiceVectors.ZatcaReplies)
        {
            var r = ZatcaReplies.Result(status, body, standard);
            Assert.Equal((result, docStatus, message), (r.Result, r.DocStatus, r.Message));
        }
    }

    [Fact]
    public void The_egyptian_receipts_are_those_of_the_mirror()
    {
        foreach (var (receipt, json, uuid, signed) in EInvoiceVectors.EtaReceipts)
        {
            Assert.Equal(json, EtaDocument.ReceiptJson(receipt));
            Assert.Equal(uuid, EtaDocument.ReceiptUuid(json));
            Assert.Equal((uuid, signed), EtaDocument.SignedReceipt(receipt));
            Assert.Equal(signed, EtaDocument.WithUuid(json, uuid));
        }
        foreach (var (value, text) in EInvoiceVectors.EtaNumbers)
        {
            Assert.Equal(text, EtaDocument.Num(value));
        }
        Assert.Equal("\"A\"\"x y\"\"B\"\"B\"\"C\"\"2\"\"B\"\"C\"\"3.5\"\"D\"\"E\"\"\"\"F\"",
                     EtaDocument.Serialize("{\"a\":\"x y\",\"b\":[{\"c\":2},{\"c\":3.5}],\"d\":{\"e\":\"\"},\"f\":[]}"));
        Assert.Equal(["C", "O", "V", "O", "C"], new[] { EtaDocument.PaymentCode("CASH", null), EtaDocument.PaymentCode("CREDIT", 2),
                                                        EtaDocument.PaymentCode("CASH", 2), EtaDocument.PaymentCode("CASH", 3),
                                                        EtaDocument.PaymentCode("CASH", 9) });
        Assert.Equal("2026-10-09T09:00:00Z", EtaDocument.UtcTimestamp(new DateTime(2026, 10, 9, 12, 0, 0)));   // Cairo summer time, UTC+3
        Assert.Equal("2026-01-09T10:00:00Z", EtaDocument.UtcTimestamp(new DateTime(2026, 1, 9, 12, 0, 0)));    // winter, UTC+2
    }

    [Fact]
    public void The_egyptian_invoices_are_those_of_the_mirror()
    {
        foreach (var (doc, json, serialized, signed) in EInvoiceVectors.EtaInvoices)
        {
            Assert.Equal(json, EtaDocument.InvoiceJson(doc));
            Assert.Equal(serialized, EtaDocument.Serialize(json));
            Assert.Equal(signed, EtaDocument.WithSignature(json, "SIG=="));
        }
        Assert.Equal("\"C:\\sign.exe\" \"a.txt\" \"b.txt\" \"1234\"", EtaDocument.SignerCommand("", "C:\\sign.exe", "a.txt", "b.txt", "1234"));
        Assert.Equal("https://preprod.invoicing.eta.gov.eg/documents/U/share/L", EtaDocument.InvoiceQrLink("PREPROD", "U", "L"));
    }

    [Fact]
    public void The_egyptian_replies_are_read_as_access_reads_them()
    {
        foreach (var (status, body, result, docStatus, message, sub) in EInvoiceVectors.EtaReceiptReplies)
        {
            var r = EtaDocument.ReceiptSubmitResult(status, body);
            Assert.Equal((result, docStatus, message, sub), (r.Result, r.DocStatus, r.Message, r.SubmissionId));
        }
        foreach (var (body, status, message) in EInvoiceVectors.EtaReceiptStatus)
        {
            Assert.Equal((status, message), EtaDocument.ReceiptStatusResult(body));
        }
        foreach (var (status, body, result, docStatus, message, sub, uuid, longId) in EInvoiceVectors.EtaInvoiceReplies)
        {
            var r = EtaDocument.InvoiceSubmitResult(status, body);
            Assert.Equal((result, docStatus, message, sub, uuid, longId), (r.Result, r.DocStatus, r.Message, r.SubmissionId, r.Uuid, r.LongId));
        }
        foreach (var (body, status, message) in EInvoiceVectors.EtaInvoiceStatus)
        {
            Assert.Equal((status, message), EtaDocument.InvoiceStatusResult(body));
        }
    }
}

/// <summary>A platform that answers from the test: every request is kept with its body.</summary>
public sealed class FakePlatform(Func<HttpRequestMessage, string, (int Status, string Body)> answer) : HttpMessageHandler
{
    public List<(HttpRequestMessage Request, string Body)> Requests { get; } = [];

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        string body = request.Content is null ? "" : await request.Content.ReadAsStringAsync(cancellationToken);
        Requests.Add((request, body));
        var (status, reply) = answer(request, body);
        if (status == 0)
        {
            throw new HttpRequestException("no network");
        }
        return new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(reply, Encoding.UTF8, "application/json") };
    }
}

/// <summary>The sending from C# on a test database: the same fields, chain and log as Access (EInvoiceService, OfficeSyncService).</summary>
public class EInvoiceServiceTests
{
    private static readonly DateTime T0 = new(2026, 11, 17, 10, 0, 0);

    private sealed class Store : IDisposable
    {
        public TestDb Db { get; } = new();
        public string KeyFile { get; } = Path.Combine(Path.GetTempPath(), "rs-test-key-" + Guid.NewGuid().ToString("N") + ".pem");
        public int Cash = 1, Firm, Rice, Water;

        public Store(string country)
        {
            File.WriteAllText(KeyFile, EInvoiceVectors.TestKey);
            using var db = Db.CreateDbContext();
            bool egypt = country == "EG";
            var firm = new Customers
            {
                CustomerName = "شركة الأمل", VATNumber = egypt ? "987654321" : "311111111100003", StreetName = "التحرير", BuildingNo = "4321",
                District = "الدقي", City = "الجيزة", PostalCode = "23456", CreditLimit = 100000,
            };
            var rice = new Products { ProductCode = "R1", ProductName = "أرز", SellingPrice = 23, AverageCost = 12, CurrentQuantity = 100,
                                      EtaItemCode = "EG-123456789-1001" };
            var water = new Products { ProductCode = "W1", ProductName = "مياه", SellingPrice = 2, AverageCost = 1, CurrentQuantity = 100,
                                       VATCategory = "Z" };
            db.AddRange(firm, rice, water);
            db.SaveChanges();
            (Firm, Rice, Water) = (firm.CustomerID, rice.ProductID, water.ProductID);
            var s = db.Settings.First();
            s.EInvoiceEnabled = true;
            s.EInvoiceEnvironment = "TEST";
            s.StoreName = "مؤسسة النور";
            s.StreetName = "الملك فهد";
            s.BuildingNo = "1234";
            s.PostalCode = "12345";
            s.City = "الرياض";
            s.District = "العليا";
            if (egypt)
            {
                s.CountryCode = "EG";
                s.VATRate = 0.14m;
                s.CurrencyCode = "EGP";
                s.VATNumber = "123456789";
                s.EtaGovernate = "القاهرة";
                s.EtaActivityCode = "4711";
                s.EtaPosSerial = "POS-0001";
                s.EtaClientId = "pos-id";
                s.EtaClientSecret = "pos-secret";
                s.EtaErpClientId = "erp-id";
                s.EtaErpClientSecret = "erp-secret";
                s.EtaSignerPath = "C:\\signer.exe";
            }
            else
            {
                s.VATNumber = "399999999900003";
                s.ZatcaOnboardStage = "PRODUCTION";
                s.ZatcaOnboardEnv = "TEST";
                s.ZatcaProductionToken = "TOKEN";
                s.ZatcaProductionSecret = "SECRET";
                s.ZatcaKeyFile = KeyFile;
                s.ZatcaCertificate = EInvoiceVectors.TestCert;
            }
            db.SaveChanges();
        }

        public async Task<int> SellAsync(int customer, params CartLine[] lines)
        {
            var r = await new SalesService(Db, Db.Permissions).PostSaleAsync(TestDb.Admin,
                new SaleRequest(customer, customer == Cash ? "CASH" : "CREDIT", 1, 0, null, "", lines));
            Assert.Equal(SaveStatus.Saved, r.Status);
            return r.InvoiceID;
        }

        public EInvoiceService Service(FakePlatform platform, IEtaSigner? signer = null) =>
            new(Db, new HttpClient(platform), signer, () => T0);

        public void Dispose()
        {
            Db.Dispose();
            File.Delete(KeyFile);
        }
    }

    private sealed class FakeSigner : IEtaSigner
    {
        public List<string> Signed { get; } = [];

        public Task<(string Signature, string Problem)> SignAsync(string serialized, string exe, string? argsTemplate, string pin)
        {
            Signed.Add(serialized);
            return Task.FromResult(("CADES-SIGNATURE==", ""));
        }
    }

    [Fact]
    public async Task Saudi_documents_are_signed_in_a_chain_and_reported()
    {
        using var s = new Store("SA");
        int first = await s.SellAsync(s.Cash, new CartLine(s.Rice, 2, 23));
        int second = await s.SellAsync(s.Cash, new CartLine(s.Rice, 1, 23));
        var platform = new FakePlatform((req, _) => req.RequestUri!.AbsolutePath.EndsWith("/invoices/reporting/single") ? (200, "{\"reportingStatus\":\"REPORTED\"}") : (404, ""));
        var service = s.Service(platform);
        Assert.Equal("", await service.ReadyProblemAsync());
        Assert.Equal("", await service.CheckSigningAsync());               // the --check of the service: the key signs, the certificate verifies

        var (sent, failed, problem) = await service.SendPendingAsync();
        Assert.Equal((2, 0, ""), (sent, failed, problem));
        await using var db = s.Db.CreateDbContext();
        var a = await db.SalesInvoices.SingleAsync(h => h.SalesInvoiceID == first);
        var b = await db.SalesInvoices.SingleAsync(h => h.SalesInvoiceID == second);
        Assert.Equal(("REPORTED", "REPORTED", (short?)1), (a.ZatcaStatus, b.ZatcaStatus, a.EInvoiceAttempts));
        Assert.Equal(ZatcaDocument.InitialPih, a.PreviousInvoiceHash);
        Assert.Equal(a.InvoiceHash, b.PreviousInvoiceHash);                   // the chain
        Assert.Equal(b.InvoiceHash, (await db.Settings.SingleAsync()).LastInvoiceHash);
        Assert.Equal(a.InvoiceHash, ZatcaDocument.InvoiceHash(a.EInvoiceXml!));
        Assert.Contains("<cbc:UUID>" + a.ICV + "</cbc:UUID>", a.EInvoiceXml);
        Assert.Equal(9, Tags(a.QRCodeData!).Count);
        Assert.Equal("مؤسسة النور", Encoding.UTF8.GetString(Tags(a.QRCodeData!)[1]));

        // the signature in the document is the device key's signature of the hash
        string signature = Between(a.EInvoiceXml!, "<ds:SignatureValue>", "</ds:SignatureValue>");
        using var cert = X509CertificateLoader.LoadCertificate(Convert.FromBase64String(EInvoiceVectors.TestCert));
        Assert.True(cert.GetECDsaPublicKey()!.VerifyData(Convert.FromBase64String(a.InvoiceHash!), Convert.FromBase64String(signature),
                                                         HashAlgorithmName.SHA256, DSASignatureFormat.Rfc3279DerSequence));

        // the request: Basic token:secret, the hash, the uuid and the document in base64
        var (request, body) = platform.Requests[0];
        Assert.Equal("https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal/invoices/reporting/single", request.RequestUri!.ToString());
        Assert.Equal("Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes("TOKEN:SECRET")), request.Headers.GetValues("Authorization").Single());
        Assert.Equal("V2", request.Headers.GetValues("Accept-Version").Single());
        var json = JsonNode.Parse(body)!;
        Assert.Equal((a.InvoiceHash, a.InvoiceUUID), ((string?)json["invoiceHash"], (string?)json["uuid"]));
        Assert.Equal(a.EInvoiceXml, Encoding.UTF8.GetString(Convert.FromBase64String((string)json["invoice"]!)));
        var log = await db.EInvoiceLog.OrderBy(l => l.LogID).ToListAsync();
        Assert.Equal(("REPORT", "OK", (short?)200, "SALE", a.InvoiceNumber), (log[0].Action, log[0].Result, log[0].HttpStatus, log[0].DocKind, log[0].DocNumber));
        Assert.Null(log[0].EmployeeID);
    }

    [Fact]
    public async Task A_saudi_tax_invoice_is_cleared_and_takes_the_stamped_document()
    {
        using var s = new Store("SA");
        int id = await s.SellAsync(s.Firm, new CartLine(s.Rice, 1, 23));
        string stamped = "<Invoice><cbc:ID>QR</cbc:ID><cac:Attachment><cbc:EmbeddedDocumentBinaryObject mimeCode=\"text/plain\">STAMPED-QR" +
                         "</cbc:EmbeddedDocumentBinaryObject></cac:Attachment></Invoice>";
        var platform = new FakePlatform((req, _) => (200, "{\"clearanceStatus\":\"CLEARED\",\"clearedInvoice\":\"" +
                                                         Convert.ToBase64String(Encoding.UTF8.GetBytes(stamped)) + "\"}"));
        var r = await s.Service(platform).SendAsync("SALE", id);
        Assert.Equal(("OK", "CLEARED"), (r.Result, r.DocStatus));
        Assert.Equal("1", platform.Requests[0].Request.Headers.GetValues("Clearance-Status").Single());
        Assert.EndsWith("/invoices/clearance/single", platform.Requests[0].Request.RequestUri!.AbsolutePath);
        await using var db = s.Db.CreateDbContext();
        var h = await db.SalesInvoices.SingleAsync(x => x.SalesInvoiceID == id);
        Assert.Equal(("CLEARED", stamped, "STAMPED-QR"), (h.ZatcaStatus, h.EInvoiceXml, h.QRCodeData));
    }

    [Fact]
    public async Task Problems_leave_the_document_pending_as_in_access()
    {
        using var s = new Store("SA");
        int exempt = await s.SellAsync(s.Cash, new CartLine(s.Water, 1, 2));
        int good = await s.SellAsync(s.Cash, new CartLine(s.Rice, 1, 23));
        var platform = new FakePlatform((_, _) => (0, ""));                   // no network
        var service = s.Service(platform);
        var r = await service.SendAsync("SALE", exempt);
        Assert.Equal("ERROR", r.Result);
        Assert.Contains("- أصناف معفاة أو بنسبة صفرية: تحتاج سبب الإعفاء، ولم يُدعم بعد", r.Message);
        Assert.Empty(platform.Requests);
        r = await service.SendAsync("SALE", good);
        Assert.Equal(("NETWORK", "", "تعذّر الاتصال بالمنظومة."), (r.Result, r.DocStatus, r.Message));
        await using (var db = s.Db.CreateDbContext())
        {
            var h = await db.SalesInvoices.SingleAsync(x => x.SalesInvoiceID == good);
            Assert.Equal(("PENDING", (short?)1, "تعذّر الاتصال بالمنظومة."), (h.ZatcaStatus, h.EInvoiceAttempts, h.EInvoiceError));
            Assert.NotNull(h.EInvoiceXml);                                     // signed once: sent again as it is
            Assert.Equal(h.InvoiceHash, (await db.Settings.SingleAsync()).LastInvoiceHash);
        }
        // SendPending stops at the first error (the next would fail too)
        var (sent, failed, problem) = await service.SendPendingAsync();
        Assert.Equal((0, 1), (sent, failed));
        Assert.StartsWith("بيانات ناقصة للفاتورة الإلكترونية:", problem);

        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.Single().ZatcaKeyFile = s.KeyFile + ".missing";
            db.SaveChanges();
        }
        Assert.Equal("ملف المفتاح الخاص غير موجود على هذا الجهاز: " + s.KeyFile + ".missing", await service.ReadyProblemAsync());
        Assert.Equal("ERROR", (await service.SendAsync("SALE", good)).Result);
    }

    [Fact]
    public async Task Egyptian_receipts_are_chained_submitted_and_followed()
    {
        using var s = new Store("EG");
        int first = await s.SellAsync(s.Cash, new CartLine(s.Rice, 2, 23));
        int second = await s.SellAsync(s.Cash, new CartLine(s.Rice, 1, 23));
        int n = 0;
        var platform = new FakePlatform((req, _) =>
        {
            string path = req.RequestUri!.AbsolutePath;
            if (path == "/connect/token")
            {
                return (200, "{\"access_token\":\"POS-TOKEN\",\"expires_in\":3600}");
            }
            if (path == "/api/v1/receiptsubmissions")
            {
                return (202, "{\"submissionId\":\"SUB" + (++n) + "\",\"acceptedDocuments\":[{\"uuid\":\"x\"}],\"rejectedDocuments\":[]}");
            }
            return (200, "{\"receipts\":[{\"status\":\"Valid\"}]}");
        });
        var service = s.Service(platform);
        var (sent, _, problem) = await service.SendPendingAsync();
        Assert.Equal((2, ""), (sent, problem));
        await using (var db = s.Db.CreateDbContext())
        {
            var a = await db.SalesInvoices.SingleAsync(h => h.SalesInvoiceID == first);
            var b = await db.SalesInvoices.SingleAsync(h => h.SalesInvoiceID == second);
            Assert.Equal(("SUBMITTED", "SUB1", "SUB2"), (a.ZatcaStatus, a.EtaSubmissionId, b.EtaSubmissionId));
            Assert.Null(a.EtaPreviousUUID);
            Assert.Equal(a.EtaUUID, b.EtaPreviousUUID);                         // the chain of the device
            Assert.Equal(b.EtaUUID, (await db.Settings.SingleAsync()).EtaLastUUID);
            Assert.Equal(a.EtaUUID, EtaDocument.ReceiptUuid(EtaDocument.WithUuid(a.EInvoiceXml!.Replace(a.EtaUUID!, ""), "")));
            Assert.StartsWith("https://preprod.invoicing.eta.gov.eg/receipts/search/" + a.EtaUUID + "/share/", a.QRCodeData);
            var log = await db.EInvoiceLog.OrderBy(l => l.LogID).ToListAsync();
            Assert.Equal(["TOKEN", "SUBMIT", "SUBMIT"], log.Select(l => l.Action));       // one token for both
            Assert.DoesNotContain("POS-TOKEN", log[0].ResponseBody);
            Assert.DoesNotContain("pos-secret", string.Concat(log.Select(l => l.RequestBody)));
        }
        var tokenRequest = platform.Requests[0];
        Assert.Equal("POS-0001", tokenRequest.Request.Headers.GetValues("posserial").Single());
        Assert.Equal("grant_type=client_credentials&client_id=pos-id&client_secret=pos-secret", tokenRequest.Body);
        Assert.Equal("Bearer POS-TOKEN", platform.Requests[1].Request.Headers.GetValues("Authorization").Single());

        var (done, changed, p2) = await service.RefreshSubmittedAsync();
        Assert.Equal((2, 2, ""), (done, changed, p2));
        Assert.Equal("VALID", await service.StatusOfAsync("SALE", first));
        Assert.EndsWith("/api/v1/receiptsubmissions/SUB1/details", platform.Requests[^2].Request.RequestUri!.AbsolutePath);
    }

    [Fact]
    public async Task An_egyptian_tax_invoice_is_signed_by_the_signing_program_and_submitted()
    {
        using var s = new Store("EG");
        int id = await s.SellAsync(s.Firm, new CartLine(s.Rice, 1, 23));
        var platform = new FakePlatform((req, _) => req.RequestUri!.AbsolutePath switch
        {
            "/connect/token" => (200, "{\"access_token\":\"ERP-TOKEN\",\"expires_in\":3600}"),
            "/api/v1.0/documentsubmissions" => (202, "{\"submissionId\":\"S9\",\"acceptedDocuments\":[{\"uuid\":\"U1\",\"longId\":\"L1\"}]}"),
            _ => (200, "{\"status\":\"Valid\"}"),
        });
        var signer = new FakeSigner();
        var service = s.Service(platform, signer);
        var r = await service.SendAsync("SALE", id);
        Assert.Equal(("OK", "SUBMITTED"), (r.Result, r.DocStatus));
        await using (var db = s.Db.CreateDbContext())
        {
            var h = await db.SalesInvoices.SingleAsync(x => x.SalesInvoiceID == id);
            Assert.Equal(("STANDARD", "U1", "L1", "S9"), (h.InvoiceSubType, h.EtaUUID, h.EtaLongId, h.EtaSubmissionId));
            Assert.Equal("https://preprod.invoicing.eta.gov.eg/documents/U1/share/L1", h.QRCodeData);
            Assert.EndsWith(",\"signatures\":[{\"signatureType\":\"I\",\"value\":\"CADES-SIGNATURE==\"}]}", h.EInvoiceXml);
            Assert.Contains("\"documentTypeVersion\":\"1.0\"", h.EInvoiceXml);
            Assert.Equal(EtaDocument.Serialize(h.EInvoiceXml![..h.EInvoiceXml.IndexOf(",\"signatures\"", StringComparison.Ordinal)] + "}"),
                         signer.Signed.Single());
        }
        Assert.Equal("grant_type=client_credentials&client_id=erp-id&client_secret=erp-secret", platform.Requests[0].Body);
        Assert.Equal(("OK", "VALID"), ((await service.EtaRefreshAsync("SALE", id)).Result, await service.StatusOfAsync("SALE", id)));
        Assert.EndsWith("/api/v1.0/documents/U1/details", platform.Requests[^1].Request.RequestUri!.AbsolutePath);
        var c = await service.EtaCancelAsync("SALE", id, "خطأ في السعر");
        Assert.Equal(("OK", "CANCELLED"), (c.Result, await service.StatusOfAsync("SALE", id)));
        Assert.Equal(HttpMethod.Put, platform.Requests[^1].Request.Method);
        Assert.Equal("{\"status\":\"cancelled\",\"reason\":\"خطأ في السعر\"}", platform.Requests[^1].Body);
    }

    [Fact]
    public async Task The_office_service_brings_the_journal_up_to_date_and_sends()
    {
        using var s = new Store("SA");
        await s.SellAsync(s.Cash, new CartLine(s.Rice, 2, 23));
        var platform = new FakePlatform((_, _) => (200, "{}"));
        var office = new OfficeSyncService(s.Db, new JournalService(s.Db, new ServerViewJournalReader()), s.Service(platform), () => T0);
        string name = OfficeSyncService.ServiceName("", "OFFICE-PC");
        Assert.Equal("SERVICE-OFFICE-PC", name);
        Assert.Equal("SERVICE-X", OfficeSyncService.ServiceName(" SERVICE-X ", "PC"));
        Assert.False(await office.IsDueAsync(name));                         // nobody named yet
        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.Single().OfficeSyncComputer = "service-office-pc";
            db.SaveChanges();
        }
        Assert.True(await office.IsDueAsync(name));                          // named (not case sensitive) and never ran

        var run = await office.RunAsync("AUTO", name, new ActingUser(0, name));
        Assert.Equal("", run.Problem);
        Assert.True(run.JournalAdded > 0);
        Assert.Equal((1, 0, 0), (run.Sent, run.NotSent, run.Checked));
        Assert.StartsWith("القيود: " + run.JournalAdded + " جديد، 0 معدّل، 0 محذوف. الفواتير الإلكترونية: أُرسلت 1،", run.Summary);
        await using (var db = s.Db.CreateDbContext())
        {
            var log = await db.OfficeSyncLog.SingleAsync();
            Assert.Equal(("AUTO", name, 1, (string?)null, (int?)null), (log.RunKind, log.ComputerName, log.EInvoicesSent!.Value, log.Problem, log.EmployeeID));
            var st = await db.Settings.SingleAsync();
            Assert.Equal((T0, (string?)null), (st.OfficeSyncLastAt!.Value, st.OfficeSyncLastResult));
            Assert.True(await db.JournalEntries.AnyAsync(e => e.SourceType == "SALE"));
            Assert.Equal("REPORTED", (await db.SalesInvoices.SingleAsync()).ZatcaStatus);
        }
        Assert.False(await office.IsDueAsync(name));                         // it just ran

        // a platform that is not set up: the pending documents make it a problem, written for the web to warn
        await s.SellAsync(s.Cash, new CartLine(s.Rice, 1, 23));
        await using (var db = s.Db.CreateDbContext())
        {
            db.Settings.Single().ZatcaOnboardStage = null;
            db.SaveChanges();
        }
        run = await office.RunAsync("MANUAL", name, TestDb.Admin);
        Assert.Equal("الجهاز لم يُسجَّل بعد لدى الهيئة (إعداد الربط: الخطوات 1 و2 و3).", run.Problem);
        Assert.Equal(1, await office.PendingCountAsync());
        await using (var db = s.Db.CreateDbContext())
        {
            Assert.Equal(run.Problem, (await db.Settings.SingleAsync()).OfficeSyncLastResult);
            Assert.Equal(1, (await db.OfficeSyncLog.OrderBy(l => l.OfficeSyncLogID).LastAsync()).EmployeeID);
        }
    }

    private static Dictionary<int, byte[]> Tags(string qr)
    {
        var raw = Convert.FromBase64String(qr);
        var o = new Dictionary<int, byte[]>();
        for (int i = 0; i < raw.Length;)
        {
            int tag = raw[i], len = raw[i + 1], start = i + 2;
            if (len == 0x81)
            {
                len = raw[i + 2];
                start = i + 3;
            }
            else if (len == 0x82)
            {
                len = raw[i + 2] * 256 + raw[i + 3];
                start = i + 4;
            }
            o[tag] = raw[start..(start + len)];
            i = start + len;
        }
        return o;
    }

    private static string Between(string s, string a, string b)
    {
        int i = s.IndexOf(a, StringComparison.Ordinal) + a.Length;
        return s[i..s.IndexOf(b, i, StringComparison.Ordinal)];
    }
}
