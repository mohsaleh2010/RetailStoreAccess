"""Reference for modActivation: machine id and activation code (SHA-256, upper-case hex).

    machine id      = first 16 hex of SHA-256("RS-MACHINE|" + upper(hardware text)), as XXXX-XXXX-XXXX-XXXX
    activation code = first 20 hex of SHA-256(LICENSE_SECRET + "|" + machine id without dashes),
                      as XXXXX-XXXXX-XXXXX-XXXXX

The programmer can also make codes with this file:
    python3 tools/activation_reference.py ABCD-1234-...      (uses LICENSE_SECRET of src/vba/modActivation.bas)
"""
import hashlib
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")


def sha_upper(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest().upper()


def normalize(text: str) -> str:
    return text.strip().replace("-", "").replace(" ", "").upper()


def machine_id_from(hardware_text: str) -> str:
    h = sha_upper("RS-MACHINE|" + hardware_text.upper())
    return "-".join(h[i:i + 4] for i in range(0, 16, 4))


def activation_code(machine_id: str, secret: str) -> str:
    h = sha_upper(secret + "|" + normalize(machine_id))
    return "-".join(h[i:i + 5] for i in range(0, 20, 5))


def is_valid(machine_id: str, code: str, secret: str) -> bool:
    return len(normalize(machine_id)) == 16 and len(normalize(code)) == 20 and \
        normalize(code) == normalize(activation_code(machine_id, secret))


def vba_secret() -> str:
    with open(os.path.join(ROOT, "src", "vba", "modActivation.bas"), encoding="utf-8") as fh:
        return re.search(r'Private Const LICENSE_SECRET As String = "([^"]*)"', fh.read()).group(1)


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        print(arg, "->", activation_code(arg, vba_secret()))
