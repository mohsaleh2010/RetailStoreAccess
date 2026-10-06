"""dist/tools/BuildFrontEnd.vbs builds the whole front-end in Access in one step:
every procedure it runs exists as a public procedure of the imported modules, the
build steps run in the installation order, and the script is plain ASCII (WSH)."""
import os
import re
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
SCRIPT = os.path.join(ROOT, "dist", "tools", "BuildFrontEnd.vbs")


def public_procs():
    found = set()
    folder = os.path.join(ROOT, "dist", "vba")
    for name in os.listdir(folder):
        with open(os.path.join(folder, name), encoding="cp1256") as fh:
            found |= set(re.findall(r"^Public (?:Function|Sub) (\w+)", fh.read(), re.M))
    return found


class BuildFrontEndTests(unittest.TestCase):

    def setUp(self):
        with open(SCRIPT, "rb") as fh:
            self.raw = fh.read()
        self.text = self.raw.decode("ascii")

    def test_runs_existing_procedures_in_order(self):
        steps = re.findall(r'RunStep "(\w+)", (True|False)', self.text)
        names = [s for s, _ in steps]
        self.assertEqual([s for s, req in steps if req == "True"],
                         ["BuildSchema", "BuildRelationships", "BuildQueries", "BuildForms", "BuildReports"])
        self.assertLessEqual(set(names), public_procs())
        self.assertLess(names.index("BuildReports"), names.index("LoadDemoData"))
        self.assertLess(names.index("LoadDemoData"), names.index("RunAllTests"))

    def test_imports_every_module_then_compiles(self):
        self.assertIn('fso.BuildPath(fso.GetParentFolderName(here), "vba")', self.text)
        self.assertIn('LCase(fso.GetExtensionName(f.Name)) = "bas"', self.text)
        self.assertEqual(self.text.count('Compile "'), 2)
        self.assertIn("app.IsCompiled", self.text)

    def test_ascii_crlf(self):
        self.assertNotIn(b"\n", self.raw.replace(b"\r\n", b""))


if __name__ == "__main__":
    unittest.main()
