"""Phase 10 reference for modSecurity and modBackup (pure functions only).

Passwords are never stored. Employees.PasswordHash holds
    hex( SHA-256^ITERATIONS( UTF-8(salt + password) ) )
where SHA-256^n means "hash the previous 32-byte digest again", and the salt is
a random 32-hex-character value per user (Employees.PasswordSalt).
The VBA implementation (modSecurity.Sha256Hex / PasswordHash) must give the
same digests as Python's hashlib; tests/test_security.py runs it in LibreOffice.

Backups are named <base>_yyyy-mm-dd_hhnnss.accdb, so sorting the names sorts
them by time; the newest KeepCount files are kept.
"""

import datetime as dt
import hashlib
from typing import List

ITERATIONS = 100
MIN_PASSWORD_LENGTH = 6
MAX_FAILED_LOGINS = 5
LOCK_MINUTES = 15


def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def password_hash(password: str, salt: str) -> str:
    x = (salt + password).encode("utf-8")
    for _ in range(ITERATIONS):
        x = hashlib.sha256(x).digest()
    return x.hex()


def password_problem(password: str, username: str) -> str:
    """"" when acceptable, otherwise a short reason code (the VBA shows Arabic text)."""
    if len(password) < MIN_PASSWORD_LENGTH:
        return "SHORT"
    if password.strip().lower() == username.strip().lower():
        return "SAME_AS_USER"
    if password != password.strip():
        return "SPACES"
    return ""


def backup_file_name(base: str, when: dt.datetime) -> str:
    return f"{base}_{when:%Y-%m-%d_%H%M%S}.accdb"


def backups_to_delete(names: List[str], keep: int) -> List[str]:
    ordered = sorted(names)
    if keep < 1:
        keep = 1
    return ordered[:max(0, len(ordered) - keep)]


SHA_VECTORS = ["", "abc", "abcdbcdecdefdefgefghfghighijhijkijkljklmmnomnopnopq", "a" * 55, "a" * 56,
               "a" * 64, "a" * 1000, "متجر الاختبار", "كلمة-سر 123 😀"]
PASSWORD_VECTORS = [("Admin@2026", "0F1E2D3C4B5A69788796A5B4C3D2E1F0"),
                    ("كلمة سر عربية", "00000000000000000000000000000000"),
                    ("123456", "ABCDEFABCDEFABCDEFABCDEFABCDEF12")]
