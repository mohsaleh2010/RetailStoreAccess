"""The delivery package of the Access program with the owner's own LICENSE_SECRET (docs/88, phase 7D).

The secret in the repository (src/vba/modActivation.bas) is the development one: everyone who reads the repository
knows it. The secret of the delivered program is made here, on the owner's computer, kept in
release/license-secret.txt (release/ is in .gitignore) and written only into the copy of the modules in the package.
It is never printed. Keep a copy of that file in a safe place: without it no new activation code can be made.

    python tools/release.py secret                 make release/license-secret.txt once (refuses to replace it)
    python tools/release.py package                release/RetailStore/ = dist/vba (modActivation with the secret),
                                                   dist/tools, dist/sql and the installation guide docs/88
    python tools/release.py code XXXX-XXXX-XXXX-XXXX   the activation code of a customer computer (the release secret)
    python tools/release.py check                  the package has the release secret and not the development one

Then on Windows: dist\\tools\\BuildFrontEnd.vbs of the package (release\\RetailStore\\tools) builds the ACCDB,
and release\\RetailStore\\tools\\MakeAccde.vbs makes the ACCDE that is delivered.
"""
import os
import re
import secrets
import shutil
import string
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import activation_reference  # noqa: E402

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
RELEASE = os.path.join(ROOT, "release")
SECRET_FILE = os.path.join(RELEASE, "license-secret.txt")
PACKAGE = os.path.join(RELEASE, "RetailStore")
DIST = os.path.join(ROOT, "dist")
GUIDE = os.path.join(ROOT, "docs", "88-Delivery-Installation.md")
MODULE = "modActivation.bas"
SECRET_LINE = re.compile(r'(Private Const LICENSE_SECRET As String = ")([^"]*)(")')
SECRET_LENGTH = 48
MIN_LENGTH = 32


class ReleaseError(Exception):
    pass


def new_secret() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(SECRET_LENGTH))


def secret_problem(secret: str, dev_secret: str) -> str:
    if len(secret) < MIN_LENGTH:
        return f"the secret must have at least {MIN_LENGTH} characters"
    if not re.fullmatch(r"[A-Za-z0-9_\-.+*/!@#$%^()=~]+", secret):
        return "the secret may hold only Latin letters, digits and simple signs (no quote, no space)"
    if secret == dev_secret:
        return "the secret is the development one of the repository"
    return ""


def read_secret(path: str = None) -> str:
    path = path or SECRET_FILE
    if not os.path.exists(path):
        raise ReleaseError("release/license-secret.txt not found: run  python tools/release.py secret  first")
    with open(path, encoding="ascii") as fh:
        return fh.read().strip()


def make_secret(path: str = None) -> bool:
    path = path or SECRET_FILE
    if os.path.exists(path):
        return False
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="ascii") as fh:
        fh.write(new_secret() + "\n")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    return True


def with_secret(module_bytes: bytes, secret: str) -> bytes:
    """modActivation of dist/vba (Windows-1256, CRLF) with the given secret; every other byte unchanged."""
    text = module_bytes.decode("cp1256")
    if len(SECRET_LINE.findall(text)) != 1:
        raise ReleaseError("modActivation: the line of LICENSE_SECRET was not found once")
    text = SECRET_LINE.sub(lambda m: m.group(1) + secret + m.group(3), text)
    return text.encode("cp1256")


def secret_in(module_bytes: bytes) -> str:
    match = SECRET_LINE.search(module_bytes.decode("cp1256"))
    return match.group(2) if match else ""


def build_package(secret: str, dist: str = DIST, package: str = None, guide: str = GUIDE) -> str:
    package = package or PACKAGE
    dev = activation_reference.vba_secret()
    problem = secret_problem(secret, dev)
    if problem:
        raise ReleaseError(problem)
    if os.path.isdir(package):
        shutil.rmtree(package)
    for folder in ("vba", "tools", "sql"):
        shutil.copytree(os.path.join(dist, folder), os.path.join(package, folder))
    path = os.path.join(package, "vba", MODULE)
    with open(path, "rb") as fh:
        data = fh.read()
    with open(path, "wb") as fh:
        fh.write(with_secret(data, secret))
    if os.path.exists(guide):
        shutil.copy2(guide, os.path.join(package, os.path.basename(guide)))
    return package


def check_package(secret: str, package: str = None) -> str:
    package = package or PACKAGE
    path = os.path.join(package, "vba", MODULE)
    if not os.path.exists(path):
        return "the package is not made: run  python tools/release.py package"
    with open(path, "rb") as fh:
        found = secret_in(fh.read())
    if found == activation_reference.vba_secret():
        return "the package has the development secret"
    if found != secret:
        return "the package has another secret than release/license-secret.txt: run package again"
    return ""


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv or argv[0] not in ("secret", "package", "code", "check"):
        print(__doc__)
        return 1
    try:
        if argv[0] == "secret":
            if make_secret():
                print("Made release/license-secret.txt (not shown). Keep a copy of it in a safe place.")
            else:
                print("release/license-secret.txt is already there: kept (a new one would make every code given invalid).")
        elif argv[0] == "package":
            build_package(read_secret())
            print(f"Package: {os.path.relpath(PACKAGE, ROOT)} (modActivation has the release secret).")
        elif argv[0] == "code":
            secret = read_secret()
            for machine in argv[1:]:
                if len(activation_reference.normalize(machine)) != 16:
                    print(f"{machine}: not a machine id (XXXX-XXXX-XXXX-XXXX)")
                    continue
                print(machine, "->", activation_reference.activation_code(machine, secret))
        else:
            problem = check_package(read_secret())
            print(problem or "The package is ready: it has the release secret.")
            return 1 if problem else 0
    except ReleaseError as ex:
        print(f"Error: {ex}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
