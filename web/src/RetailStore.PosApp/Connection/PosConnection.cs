using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using RetailStore.Data.Services;

namespace RetailStore.PosApp.Connection;

/// <summary>Where the POS reads and writes the program's data (docs/83), chosen by an administrator on the screen
/// «إعدادات قاعدة البيانات»:
/// LOCAL = a database of the store reached directly: SQL Server on this PC or the store network (the server of plan 5,
/// the same one Access links to), or a SQLite file for a trial;
/// REMOTE = the program's server elsewhere, reached only through the API by its address (HTTPS), never a database port
/// over the internet.
/// DeviceCode = the code of this PC in the numbers of its invoices (POS1-000001), one code per device.
/// The SQL password is not here: it is in the protected storage of Windows (IPosPlatform).</summary>
public sealed record PosConnection(string Mode, string Provider, string Server, string Database, bool WindowsLogin, string User,
                                   bool TrustCertificate, string SqliteFile, string ServerUrl, string DeviceCode)
{
    public const string Local = "LOCAL";
    public const string Remote = "REMOTE";
    public const string SqlServer = "SqlServer";
    public const string Sqlite = "Sqlite";

    public static readonly PosConnection Empty = new(Local, SqlServer, @".\SQLEXPRESS", "RetailStore", true, "", true, "", "", "POS1");

    public bool IsLocal => Mode == Local;

    public bool IsSqlite => IsLocal && Provider == Sqlite;

    /// <summary>The address of the server as the API client uses it: https:// added when no scheme is written.</summary>
    public static string NormalizeUrl(string? url)
    {
        string u = (url ?? "").Trim().TrimEnd('/');
        if (u.Length > 0 && !u.Contains("://", StringComparison.Ordinal))
        {
            u = "https://" + u;
        }
        return u;
    }

    /// <summary>"" or what is missing or wrong, checked before the connection is tried.</summary>
    public string Problem(string? password)
    {
        string code = PosService.DeviceCodeProblem(DeviceCode, "");
        if (code.Length > 0)
        {
            return code;
        }
        if (Mode == Remote)
        {
            string url = NormalizeUrl(ServerUrl);
            if (url.Length == 0)
            {
                return "اكتب عنوان الخادم.";
            }
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || (uri.Scheme != "https" && uri.Scheme != "http"))
            {
                return "عنوان الخادم غير صالح، مثل https://store.example.com";
            }
            return uri.Scheme == "http" && !IsPrivateHost(uri.Host)
                ? "الخادم البعيد يحتاج HTTPS. يُقبل http داخل شبكة المحل فقط."
                : "";
        }
        if (Mode != Local)
        {
            return "اختر نوع قاعدة البيانات: محلية أو بعيدة.";
        }
        if (Provider == Sqlite)
        {
            return SqliteFile.Trim().Length == 0 ? "اختر ملف قاعدة التجربة." : "";
        }
        if (Provider != SqlServer)
        {
            return "اختر نوع قاعدة البيانات: محلية أو بعيدة.";
        }
        if (Server.Trim().Length == 0)
        {
            return "اكتب اسم الخادم، مثل .\\SQLEXPRESS";
        }
        if (Database.Trim().Length == 0)
        {
            return "اكتب اسم قاعدة البيانات.";
        }
        if (!WindowsLogin && User.Trim().Length == 0)
        {
            return "اكتب اسم مستخدم SQL Server، أو اختر الدخول بحساب Windows.";
        }
        return !WindowsLogin && string.IsNullOrEmpty(password) ? "اكتب كلمة مرور مستخدم SQL Server." : "";
    }

    /// <summary>A host of the store network: this PC, a private IPv4 address, or a name without a dot.</summary>
    public static bool IsPrivateHost(string host)
    {
        if (host == "localhost" || !host.Contains('.'))
        {
            return true;
        }
        if (System.Net.IPAddress.TryParse(host, out var ip) && ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
        {
            var b = ip.GetAddressBytes();
            return b[0] == 10 || b[0] == 127 || (b[0] == 192 && b[1] == 168) || (b[0] == 172 && b[1] >= 16 && b[1] <= 31);
        }
        return false;
    }

    /// <summary>The connection string of a local database (docs/68: DSN-less, encrypted, the certificate of the store
    /// server trusted when asked).</summary>
    public string ConnectionString(string? password)
    {
        if (Provider == Sqlite)
        {
            return new SqliteConnectionStringBuilder { DataSource = SqliteFile.Trim(), ForeignKeys = true }.ToString();
        }
        var b = new SqlConnectionStringBuilder
        {
            DataSource = Server.Trim(), InitialCatalog = Database.Trim(), Encrypt = true, TrustServerCertificate = TrustCertificate,
            ConnectTimeout = 8, ApplicationName = "RetailStore POS " + DeviceCode.Trim().ToUpperInvariant(),
        };
        if (WindowsLogin)
        {
            b.IntegratedSecurity = true;
        }
        else
        {
            b.UserID = User.Trim();
            b.Password = password ?? "";
        }
        return b.ToString();
    }

    /// <summary>What the screens show of the connection (never the password).</summary>
    public string Describe() => Mode == Remote ? NormalizeUrl(ServerUrl)
        : Provider == Sqlite ? SqliteFile : Server.Trim() + " / " + Database.Trim();
}
