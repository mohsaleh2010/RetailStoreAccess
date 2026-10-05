"""Execute pure-logic VBA modules with LibreOffice Basic (VBA compatibility mode).

Access-only statements are stubbed so that the calculation code (pricing,
UTF-8/Base64/TLV, QR encoder) really runs, not just its Python reference.
Used by tests/test_vba_runtime.py; skipped when LibreOffice is unavailable.
"""

import os
import re
import shutil
import socket
import subprocess
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

STUBS = r'''
Option VBASupport 1
Option Explicit
Public Calendar As Integer
Public Const vbCalGreg As Integer = 0

Public Function HarnessDec(v As Variant) As Double
    ' LibreOffice has no Decimal type off Windows: use Double, and round with a
    ' tiny epsilon in HarnessRoundMoney to absorb binary noise (2.345*100 = 234.4999...).
    HarnessDec = CDbl(v)
End Function

Public Function Nz(v As Variant, Optional d As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then
        If IsMissing(d) Then Nz = "" Else Nz = d
    Else
        Nz = v
    End If
End Function
'''


HARNESS_ROUND = """Public Function RoundMoney(ByVal Value As Variant, Optional ByVal Digits As Integer = 2) As Currency
    Dim x As Double, f As Double
    f = 10 ^ Digits
    x = CDbl(Nz(Value, 0)) * f
    If x >= 0 Then x = Fix(x + 0.5 + 0.000001) Else x = -Fix(-x + 0.5 + 0.000001)
    RoundMoney = CCur(x / f)
End Function"""


def available() -> bool:
    if not shutil.which("soffice"):
        return False
    try:
        import uno  # noqa: F401
    except ImportError:
        return False
    return True


def prepare(code: str) -> str:
    """Make an Access VBA module loadable by LibreOffice Basic."""
    out, skip = [], False
    for line in code.splitlines():
        s = line.strip()
        if s.startswith("Attribute VB_Name") or s == "Option Compare Database":
            continue
        if s.startswith("#If"):
            skip = True
            continue
        if s.startswith("#End If"):
            skip = False
            continue
        if skip or s.startswith("#Else"):
            continue
        out.append(line)
    text = "\n".join(out)
    # LibreOffice differences that are not VBA bugs:
    text = text.replace("CDec(", "HarnessDec(")        # Decimal is Windows-only in LO
    text = text.replace("hh:nn:ss", "hh:mm:ss")       # LO reads "nn" as a weekday
    text = re.sub(r"^(\s*)(rpt\.(?:Line|Circle) \(|rpt\.Print )", r"\1' \2", text, flags=re.M)   # report drawing
    text = re.sub(r"\bAs (?:Access|DAO)\.\w+", "As Object", text)          # Access/DAO object types
    # RoundMoney relies on Decimal; replace it with an equivalent Double version.
    text = re.sub(r"Public Function RoundMoney\(.*?\nEnd Function", HARNESS_ROUND, text, flags=re.S)
    if "Option VBASupport" not in text:
        text = "Option VBASupport 1\n" + text
    return text


class Harness:
    def __init__(self):
        import uno
        self.uno = uno
        self.profile = tempfile.mkdtemp(prefix="lo_profile_")
        with socket.socket() as s:
            s.bind(("127.0.0.1", 0))
            self.port = s.getsockname()[1]
        self.proc = subprocess.Popen(
            ["soffice", "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
             f"-env:UserInstallation=file://{self.profile}",
             f"--accept=socket,host=127.0.0.1,port={self.port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        deadline = time.time() + 60
        while True:
            try:
                self.ctx = resolver.resolve(
                    f"uno:socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                if time.time() > deadline:
                    raise
                time.sleep(0.5)
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        hidden = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        hidden.Name, hidden.Value = "Hidden", True
        self.doc = self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (hidden,))
        libs = self.doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        self.lib = libs.getByName("Standard")
        self._modules = 0

    def add_module(self, name: str, code: str):
        if self.lib.hasByName(name):
            self.lib.replaceByName(name, code)
        else:
            self.lib.insertByName(name, code)

    def load(self, modules):
        self.add_module("HarnessStubs", STUBS)
        for name, code in modules.items():
            self.add_module(name, prepare(code))

    def call(self, module: str, func: str, *args):
        provider = self.doc.getScriptProvider()
        script = provider.getScript(
            f"vnd.sun.star.script:Standard.{module}.{func}?language=Basic&location=document")
        result, _, _ = script.invoke(tuple(args), (), ())
        return result

    def close(self):
        try:
            self.doc.close(True)
        except Exception:
            pass
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.profile, ignore_errors=True)


def read_module(name: str) -> str:
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="utf-8") as fh:
        return fh.read()
