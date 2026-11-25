"""Publish the web system (docs/76-87) to a shared Windows host by FTP (docs/88, phase 7D).

Run it on the owner's computer (the .NET 10 SDK and Python 3), never in a shared place: the settings file it keeps,
release/web-deploy.json, holds the SQL Server password and the key of the API tokens. release/ is in .gitignore.

    python tools/deploy_web.py init        ask the database, FTP and site settings once (release/web-deploy.json)
    python tools/deploy_web.py publish     dotnet publish (Windows x64) into release/web-publish + appsettings.Production.json
    python tools/deploy_web.py upload      app_offline.htm, every file but App_Data and the development files, then remove it
    python tools/deploy_web.py check       GET <site>/health: {"status":"ok"} = the site runs and reads its database
    python tools/deploy_web.py sql-check   connect to SQL Server only: the user, the password, the database, db_owner
    python tools/deploy_web.py sql         run dist/sql/RetailStore_Schema.sql then RetailStore_Queries.sql (pip install pymssql)
    python tools/deploy_web.py all         publish, upload, check

Options: --framework-dependent (the host has the .NET 10 runtime: about 15 MB instead of 135 MB), --stdout-log (the
start-up errors of the site go to App_Data/logs: a 500.30 page), --yes (no questions). The FTP password is asked each time
(or RETAIL_FTP_PASSWORD), never kept.
"""
import argparse
import ftplib
import getpass
import json
import os
import posixpath
import re
import secrets
import string
import subprocess
import sys
import urllib.request

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
RELEASE = os.path.join(ROOT, "release")
SETTINGS_FILE = os.path.join(RELEASE, "web-deploy.json")
PUBLISH_DIR = os.path.join(RELEASE, "web-publish")
PROJECT = os.path.join(ROOT, "web", "src", "RetailStore.Web", "RetailStore.Web.csproj")
SQL_DIR = os.path.join(ROOT, "dist", "sql")

# Never uploaded: the data folder of the server (keys, logs) and what only development uses.
SKIP_TOP = {"App_Data"}
SKIP_FILES = {"appsettings.Development.json"}
SKIP_SUFFIXES = (".pdb", ".db", ".db-shm", ".db-wal")

APP_OFFLINE = ("<!doctype html><html lang=\"ar\" dir=\"rtl\"><head><meta charset=\"utf-8\"><title>RetailStore</title></head>"
               "<body style=\"font-family:Tahoma,Arial;text-align:center;padding-top:15%\">"
               "<h2>&#1580;&#1575;&#1585;&#1610; &#1578;&#1581;&#1583;&#1610;&#1579; &#1575;&#1604;&#1576;&#1585;&#1606;&#1575;&#1605;&#1580;</h2>"
               "<p>Updating the program, please try again in a few minutes.</p></body></html>")

FIELDS = [  # key, question, default, secret
    ("sql_server", "SQL Server address (e.g. sql1234.host.net or sql1234.host.net,1433)", "", False),
    ("sql_database", "SQL Server database name", "", False),
    ("sql_user", "SQL Server user", "", False),
    ("sql_password", "SQL Server password", "", True),
    ("ftp_host", "FTP host (e.g. ftp.mysite.com)", "", False),
    ("ftp_user", "FTP user", "", False),
    ("ftp_dir", "FTP folder of the site (e.g. /site/wwwroot or /httpdocs)", "/", False),
    ("ftp_tls", "FTP over TLS (FTPS)? yes/no", "yes", False),
    ("site_url", "Address of the site (e.g. https://erp.mysite.com)", "", False),
    ("api_docs", "Show the API documentation page /api/docs? yes/no", "no", False),
]


class DeployError(Exception):
    pass


def shown(path: str) -> str:
    """A path to print: relative to the repository, or whole when it is on another drive (Windows: C: and E:)."""
    try:
        return os.path.relpath(path, ROOT)
    except ValueError:
        return os.path.abspath(path)


YES_WORDS = ("y", "yes", "true", "1", "نعم", "ن")
NO_WORDS = ("n", "no", "false", "0", "لا", "ل")


def yes(text) -> bool:
    return str(text).strip().lower() in YES_WORDS


def yes_no_value(text: str):
    """'yes' / 'no' from an answer, None when it is neither (an Arabic keyboard letter, a typo): asked again."""
    word = str(text).strip().lower()
    if word in YES_WORDS:
        return "yes"
    if word in NO_WORDS:
        return "no"
    return None


def normalize_site_url(text: str):
    """(address, problem): https:// is added when no scheme is typed, the last '/' is removed; http:// is refused."""
    url = str(text).strip().rstrip("/")
    if not url:
        return "", "the address of the site is needed"
    if re.match(r"^http://", url, re.I):
        return "", "the address of the site must be https:// (the rep app and the POS send passwords to it)"
    if not re.match(r"^https://", url, re.I):
        url = "https://" + url
    if not re.match(r"^https://[A-Za-z0-9.-]+(:\d+)?(/[^\s]*)?$", url, re.I):
        return "", "not a site address (e.g. https://erp.mysite.com)"
    return url, ""


YES_NO_FIELDS = ("ftp_tls", "api_docs")


def new_jwt_key() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(64))


def load_settings(path: str = SETTINGS_FILE) -> dict:
    if not os.path.exists(path):
        raise DeployError(f"{shown(path)} not found: run  python tools/deploy_web.py init  first")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def save_settings(settings: dict, path: str = SETTINGS_FILE) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(settings, fh, indent=2, ensure_ascii=False)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def settings_problem(settings: dict) -> str:
    for key, question, _default, _secret in FIELDS:
        if not str(settings.get(key, "")).strip():
            return f"missing: {question}"
    if not re.match(r"^https://", settings["site_url"].strip(), re.I):
        return "the address of the site must start with https:// (the rep app and the POS send passwords to it)"
    if len(settings.get("jwt_key", "")) < 32:
        return "the key of the API tokens is missing: run init again"
    for key in ("sql_server", "sql_database", "sql_user", "sql_password"):
        value = settings[key].strip()
        if ";" in value or value[:1] in ("'", '"'):
            return f"{key}: ';' is not allowed, nor a quote at the start"
    return ""


def connection_string(s: dict) -> str:
    """The SQL Server connection of the site: SQL login, encrypted (the host's certificate is trusted as is)."""
    return (f"Server={s['sql_server']};Database={s['sql_database']};User Id={s['sql_user']};"
            f"Password={s['sql_password']};Encrypt=True;TrustServerCertificate=True;"
            "MultipleActiveResultSets=False;Connect Timeout=30;Application Name=RetailStore.Web")


def production_settings(s: dict) -> dict:
    """appsettings.Production.json of the site: over appsettings.json, read because the host runs it as Production."""
    return {
        "Database": {"Provider": "SqlServer"},
        "ConnectionStrings": {"SqlServer": connection_string(s), "Sqlite": ""},
        "Api": {"JwtKey": s["jwt_key"], "Docs": yes(s.get("api_docs", "no"))},
        "Hosting": {"ForwardedHeaders": yes(s.get("forwarded_headers", "no"))},
        "Logging": {"LogLevel": {"Default": "Warning", "Microsoft.AspNetCore": "Warning"}},
    }


def field_answer(key: str, value: str):
    """(value to keep, problem) of one answer of init: a wrong answer is asked again, never loses the others."""
    value = str(value).strip()
    if key == "site_url":
        return normalize_site_url(value)
    if key in YES_NO_FIELDS:
        answer = yes_no_value(value)
        return (answer, "") if answer else ("", "answer yes or no (y / n)")
    if not value:
        return "", "an answer is needed"
    if key.startswith("sql_") and (";" in value or value[:1] in ("'", '"')):
        return "", "';' is not allowed, nor a quote at the start"
    return value, ""


def ask_settings(old: dict, ask=input, ask_secret=getpass.getpass, log=print) -> dict:
    settings = dict(old)
    for key, question, default, secret in FIELDS:
        current = old.get(key, default)
        while True:
            if secret:
                typed = ask_secret(f"{question}{' [Enter = keep]' if current else ''}: ")
            else:
                typed = ask(f"{question}{f' [{current}]' if current else ''}: ")
            value, problem = field_answer(key, typed.strip() or current)
            if not problem:
                settings[key] = value
                break
            log(f"  {problem}")
    settings.setdefault("jwt_key", new_jwt_key())  # once: a new key would end every token of the phones and PCs
    return settings


def cmd_init(args) -> None:
    old = load_settings() if os.path.exists(SETTINGS_FILE) else {}
    print("RetailStore web: deployment settings (kept in release/web-deploy.json on this computer only)")
    print("Enter keeps the value in [ ]. The password is not shown while you type it.")
    settings = ask_settings(old)
    problem = settings_problem(settings)
    if problem:
        raise DeployError(problem)
    save_settings(settings)
    print("Saved. Next:  python tools/deploy_web.py sql-check")


def publish_command(output: str, framework_dependent: bool) -> list:
    command = ["dotnet", "publish", PROJECT, "-c", "Release", "-r", "win-x64", "-o", output,
               "--self-contained", "false" if framework_dependent else "true"]
    return command


def set_stdout_log(publish_dir: str, enabled: bool) -> None:
    path = os.path.join(publish_dir, "web.config")
    with open(path, encoding="utf-8-sig") as fh:
        text = fh.read()
    if "AspNetCoreModuleV2" not in text or 'hostingModel="inprocess"' not in text:
        raise DeployError("web.config of the publish is not the in-process ASP.NET Core module")
    text = re.sub(r'stdoutLogEnabled="(true|false)"', f'stdoutLogEnabled="{"true" if enabled else "false"}"', text)
    text = re.sub(r'stdoutLogFile="[^"]*"', r'stdoutLogFile=".\\App_Data\\logs\\stdout"', text)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)


def cmd_publish(args, settings: dict) -> None:
    problem = settings_problem(settings)
    if problem:
        raise DeployError(problem)
    if os.path.isdir(PUBLISH_DIR):
        import shutil
        shutil.rmtree(PUBLISH_DIR)
    subprocess.run(publish_command(PUBLISH_DIR, args.framework_dependent), check=True)
    with open(os.path.join(PUBLISH_DIR, "appsettings.Production.json"), "w", encoding="utf-8") as fh:
        json.dump(production_settings(settings), fh, indent=2)
    set_stdout_log(PUBLISH_DIR, args.stdout_log)
    print(f"Published to {shown(PUBLISH_DIR)}")


def upload_plan(publish_dir: str) -> list:
    """(relative path with '/', local path) of every file to send, folders before their files."""
    plan = []
    for folder, dirs, files in os.walk(publish_dir):
        rel_folder = os.path.relpath(folder, publish_dir).replace(os.sep, "/")
        if rel_folder == ".":
            rel_folder = ""
            dirs[:] = [d for d in dirs if d not in SKIP_TOP]
        dirs.sort()
        for name in sorted(files):
            if name in SKIP_FILES or name.lower().endswith(SKIP_SUFFIXES) or name.lower() == "app_offline.htm":
                continue
            plan.append((posixpath.join(rel_folder, name) if rel_folder else name, os.path.join(folder, name)))
    return plan


def ftp_connect(settings: dict, password: str):
    ftp = ftplib.FTP_TLS(timeout=60) if yes(settings.get("ftp_tls", "yes")) else ftplib.FTP(timeout=60)
    ftp.connect(settings["ftp_host"], 21)
    ftp.login(settings["ftp_user"], password)
    if isinstance(ftp, ftplib.FTP_TLS):
        ftp.prot_p()
    return ftp


def ensure_dirs(ftp, base: str, rel_dir: str, made: set) -> None:
    path = base.rstrip("/")
    for part in [p for p in rel_dir.split("/") if p]:
        path = f"{path}/{part}"
        if path in made:
            continue
        try:
            ftp.mkd(path)
        except ftplib.error_perm:
            pass  # already there
        made.add(path)


def upload(ftp, settings: dict, publish_dir: str, log=print) -> int:
    """app_offline.htm first (IIS stops the site and frees its files), every file, then app_offline.htm removed."""
    base = settings.get("ftp_dir", "/").rstrip("/") or ""
    plan = upload_plan(publish_dir)
    if not any(rel == "RetailStore.Web.dll" for rel, _ in plan):
        raise DeployError("RetailStore.Web.dll is not in the publish folder: run publish first")
    if not any(rel == "appsettings.Production.json" for rel, _ in plan):
        raise DeployError("appsettings.Production.json is not in the publish folder: run publish first")
    from io import BytesIO
    ftp.storbinary(f"STOR {base}/app_offline.htm", BytesIO(APP_OFFLINE.encode("utf-8")))
    made = set()
    sent = 0
    try:
        for rel, local in plan:
            ensure_dirs(ftp, base, posixpath.dirname(rel), made)
            with open(local, "rb") as fh:
                ftp.storbinary(f"STOR {base}/{rel}", fh)
            sent += 1
            if sent % 50 == 0:
                log(f"  {sent} / {len(plan)} files")
    finally:
        if sent == len(plan):
            ftp.delete(f"{base}/app_offline.htm")
    log(f"Uploaded {sent} files; the site starts again on its next request.")
    return sent


def cmd_upload(args, settings: dict) -> None:
    password = os.environ.get("RETAIL_FTP_PASSWORD") or getpass.getpass("FTP password: ")
    ftp = ftp_connect(settings, password)
    try:
        upload(ftp, settings, PUBLISH_DIR)
    finally:
        try:
            ftp.quit()
        except ftplib.all_errors:
            pass


def health_url(settings: dict) -> str:
    return settings["site_url"].rstrip("/") + "/health"


def cmd_check(args, settings: dict) -> bool:
    try:
        with urllib.request.urlopen(health_url(settings), timeout=60) as reply:
            body = json.loads(reply.read().decode("utf-8"))
            ok = reply.status == 200 and body.get("status") == "ok"
    except Exception as ex:  # noqa: BLE001 - every failure is reported the same way
        print(f"The site does not answer well: {ex}")
        return False
    print("The site runs and reads its database." if ok else f"The site answers: {body}")
    return ok


def split_batches(script: str) -> list:
    """The batches of a T-SQL script, split at the lines that hold only GO (as modSqlServer.RunSchemaScript)."""
    batches, current = [], []
    for line in script.splitlines():
        if line.strip().upper() == "GO":
            text = "\n".join(current).strip()
            if text:
                batches.append(text)
            current = []
        else:
            current.append(line)
    text = "\n".join(current).strip()
    if text:
        batches.append(text)
    return batches


def run_scripts(connection, files: list, log=print) -> int:
    count = 0
    cursor = connection.cursor()
    for path in files:
        with open(path, encoding="utf-8-sig") as fh:
            batches = split_batches(fh.read())
        for number, batch in enumerate(batches, 1):
            try:
                cursor.execute(batch)
                connection.commit()
            except Exception as ex:  # noqa: BLE001 - the server's message, with where it happened
                try:
                    connection.rollback()
                except Exception:  # noqa: BLE001
                    pass
                first = next((l.strip() for l in batch.splitlines() if l.strip() and not l.strip().startswith("--")), "")
                raise DeployError(f"{os.path.basename(path)}, part {number} of {len(batches)} ({first[:120]}):\n"
                                  f"  {str(ex).splitlines()[0][:400] if str(ex) else type(ex).__name__}\n"
                                  "  The parts before it are saved; the script can be run again (it is safe to repeat).") from None
            count += 1
            if number % 100 == 0:
                log(f"  {os.path.basename(path)}: {number} / {len(batches)}")
        log(f"{os.path.basename(path)}: {len(batches)} batches")
    return count


SQL_LOGIN_HELP = (
    "SQL Server refused the login (error 18456): the server answered, so the address is right.\n"
    "  - the password: run  python tools/deploy_web.py init  again (Enter keeps every other answer) and type it again;\n"
    "  - the user: it must be the database user of the control panel, and it must be given to this database;\n"
    "  - try the same user and password in SQL Server Management Studio (SQL Server Authentication).")
SQL_REACH_HELP = (
    "SQL Server does not answer: check the address and the port (address,port), that remote access to the "
    "database is allowed in the control panel, and that no firewall blocks port 1433.")
SQL_DATABASE_HELP = (
    "SQL Server knows the user but not this database for it (error 4060): check the database name in the control "
    "panel and that the user is given to it.")


def sql_error_help(error: Exception) -> str:
    """The advice for a refused connection, from the SQL Server error number in the message of pymssql."""
    text = str(error)
    numbers = [int(n) for n in re.findall(r"\((\d{3,5}),", text)] + [int(n) for n in re.findall(r"error (\d{3,5})", text, re.I)]
    if 18456 in numbers or "Login failed" in text:
        return SQL_LOGIN_HELP
    if 4060 in numbers or "Cannot open database" in text:
        return SQL_DATABASE_HELP
    return SQL_REACH_HELP + f"\n  ({text.splitlines()[0][:200] if text else type(error).__name__})"


def sql_connect(settings: dict, connect=None):
    if connect is None:
        try:
            import pymssql  # noqa: WPS433 - optional
        except ImportError as ex:
            raise DeployError("pip install pymssql (or run the two scripts from frmSqlServer / "
                              "SQL Server Management Studio)") from ex
        connect = pymssql.connect
    server, _, port = settings["sql_server"].partition(",")
    try:
        return connect(server=server.strip(), port=(port.strip() or "1433"), user=settings["sql_user"],
                       password=settings["sql_password"], database=settings["sql_database"], autocommit=False,
                       login_timeout=30)
    except Exception as ex:  # noqa: BLE001 - pymssql raises its own classes
        raise DeployError(sql_error_help(ex)) from None


def cmd_sql_check(args, settings: dict, connect=None, log=print) -> None:
    connection = sql_connect(settings, connect)
    try:
        cursor = connection.cursor()
        cursor.execute("SELECT DB_NAME(), CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(40)), "
                       "IS_ROLEMEMBER('db_owner')")
        name, version, owner = cursor.fetchone()
        log(f"Connected to {name} (SQL Server {version}).")
        if owner != 1:
            log("Warning: the user is not db_owner of the database: the schema script needs it "
                "(CREATE TABLE, ALTER, CREATE VIEW). Ask the host or give it in the control panel.")
    finally:
        connection.close()


def cmd_sql(args, settings: dict, connect=None) -> None:
    if not args.yes and not yes(input(f"Create / upgrade the schema of {settings['sql_database']}? yes/no: ")):
        return
    connection = sql_connect(settings, connect)
    try:
        run_scripts(connection, [os.path.join(SQL_DIR, "RetailStore_Schema.sql"),
                                 os.path.join(SQL_DIR, "RetailStore_Queries.sql")])
    finally:
        connection.close()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["init", "publish", "upload", "check", "sql-check", "sql", "all"])
    parser.add_argument("--framework-dependent", action="store_true")
    parser.add_argument("--stdout-log", action="store_true")
    parser.add_argument("--yes", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.command == "init":
            cmd_init(args)
            return 0
        settings = load_settings()
        if args.command in ("publish", "all"):
            cmd_publish(args, settings)
        if args.command in ("upload", "all"):
            cmd_upload(args, settings)
        if args.command == "sql-check":
            cmd_sql_check(args, settings)
        if args.command == "sql":
            cmd_sql(args, settings)
        if args.command in ("check", "all"):
            return 0 if cmd_check(args, settings) else 1
        return 0
    except DeployError as ex:
        print(f"Error: {ex}")
        return 2
    except KeyboardInterrupt:
        print("\nCancelled: nothing was saved or sent.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
