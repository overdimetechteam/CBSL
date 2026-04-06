"""
auth.py — Authentication & user management helpers
====================================================
* Fernet symmetric encryption using FERNET_KEY system env variable
* users.xlsx storage: un, pw, permissions, created_at, updated_at
* permissions field: comma-separated flags — CBSL, IFRS, ADMIN
  e.g.  "CBSL,IFRS,ADMIN"  or  "CBSL"  or  "IFRS"
* Admin = any user whose permissions contains "ADMIN"
* No credentials are hardcoded — all stored encrypted in users.xlsx
"""

import os
import functools
from datetime import datetime
from pathlib import Path

import openpyxl
from cryptography.fernet import Fernet, InvalidToken
from flask import session, redirect, url_for, flash, request

# ── Paths ───────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
USERS_FILE = BASE_DIR / "users.xlsx"

# ── Fernet key ──────────────────────────────────────────────────────────────
_FERNET_KEY_ENV = "FERNET_KEY"

def _get_fernet() -> Fernet:
    key = os.environ.get(_FERNET_KEY_ENV)
    if not key:
        raise RuntimeError(
            f"Environment variable '{_FERNET_KEY_ENV}' is not set. "
            "Generate one with: python -c \"from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())\" "
            "then add it as a system environment variable named FERNET_KEY."
        )
    return Fernet(key.encode() if isinstance(key, str) else key)


def encrypt(value: str) -> str:
    """Encrypt a plaintext string; returns a UTF-8 token string."""
    return _get_fernet().encrypt(value.encode()).decode()


def decrypt(token: str) -> str:
    """Decrypt a token string; raises InvalidToken on failure."""
    return _get_fernet().decrypt(token.encode()).decode()


def check_password(plaintext: str, stored_token: str) -> bool:
    """Return True if plaintext matches the stored encrypted password."""
    try:
        return decrypt(stored_token) == plaintext
    except (InvalidToken, Exception):
        return False


# ── XLSX helpers ─────────────────────────────────────────────────────────────
COLUMNS = ["un", "pw", "permissions", "created_at", "updated_at"]

def _ensure_file():
    """Create users.xlsx with headers if it doesn't exist."""
    if USERS_FILE.exists():
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "users"
    ws.append(COLUMNS)
    wb.save(USERS_FILE)


def _load_users() -> list[dict]:
    """Load all rows from users.xlsx as a list of dicts."""
    _ensure_file()
    wb = openpyxl.load_workbook(USERS_FILE)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() for h in rows[0]]
    users = []
    for row in rows[1:]:
        if not any(row):
            continue
        users.append(dict(zip(headers, row)))
    return users


def _save_users(users: list[dict]):
    """Overwrite users.xlsx with the given list of user dicts."""
    _ensure_file()
    wb = openpyxl.load_workbook(USERS_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None
    for i, user in enumerate(users, start=2):
        for j, col in enumerate(COLUMNS, start=1):
            ws.cell(row=i, column=j, value=user.get(col, ""))
    wb.save(USERS_FILE)


def get_user(username: str) -> dict | None:
    """Return user dict for the given username (decrypts to compare), or None.
    If username contains no '@', also tries appending '@assetlinefinance.lk'
    so users can log in with just the short username (e.g. 'cbslbotuser').
    """
    candidates = [username]
    if "@" not in username:
        candidates.append(username + "@assetlinefinance.lk")
    for user in _load_users():
        try:
            stored = decrypt(str(user["un"]))
            if stored in candidates:
                return user
        except Exception:
            continue
    return None


def authenticate(username: str, password: str) -> dict | None:
    """Return user dict if credentials are valid, else None."""
    user = get_user(username)
    if user and check_password(password, str(user["pw"])):
        return user
    return None


def create_user(username: str, password: str, permissions: str) -> bool:
    """
    Create a new user. permissions is comma-separated flags: CBSL, IFRS, ADMIN.
    Returns False if username already exists.
    """
    for u in _load_users():
        try:
            if decrypt(str(u["un"])) == username:
                return False
        except Exception:
            continue

    users = _load_users()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    users.append({
        "un":          encrypt(username),
        "pw":          encrypt(password),
        "permissions": permissions.upper(),
        "created_at":  now,
        "updated_at":  now,
    })
    _save_users(users)
    return True


def update_password(username: str, new_password: str) -> bool:
    """Update password for the given username. Returns False if not found."""
    users = _load_users()
    found = False
    for user in users:
        try:
            if decrypt(str(user["un"])) == username:
                user["pw"]         = encrypt(new_password)
                user["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                found = True
                break
        except Exception:
            continue
    if found:
        _save_users(users)
    return found


def update_permissions(username: str, permissions: str) -> bool:
    """Update permissions for the given username."""
    users = _load_users()
    found = False
    for user in users:
        try:
            if decrypt(str(user["un"])) == username:
                user["permissions"] = permissions.upper()
                user["updated_at"]  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                found = True
                break
        except Exception:
            continue
    if found:
        _save_users(users)
    return found


def delete_user(username: str) -> bool:
    """Delete a user by username. Returns False if not found."""
    users = _load_users()
    new_users = []
    found = False
    for user in users:
        try:
            if decrypt(str(user["un"])) == username:
                found = True
                continue
        except Exception:
            pass
        new_users.append(user)
    if found:
        _save_users(new_users)
    return found


def list_users_display() -> list[dict]:
    """Return all users with username decrypted (password hidden) for display."""
    result = []
    for user in _load_users():
        try:
            un = decrypt(str(user["un"]))
        except Exception:
            un = "???"
        result.append({
            "un":          un,
            "permissions": user.get("permissions", ""),
            "created_at":  user.get("created_at", ""),
            "updated_at":  user.get("updated_at", ""),
        })
    return result


# ── Admin detection — based on permissions field, no hardcoded username ──────

def is_admin(username: str) -> bool:
    """Admin = any user whose permissions contains 'ADMIN'."""
    user = get_user(username)
    if not user:
        return False
    return "ADMIN" in str(user.get("permissions", "")).upper().split(",")


def current_user() -> str | None:
    """Return the currently logged-in username from session."""
    return session.get("username")


def current_permissions() -> str:
    """Return permissions string for the logged-in user."""
    return session.get("permissions", "")


def has_permission(perm: str) -> bool:
    """Check if the current user has a specific permission (CBSL, IFRS, or ADMIN)."""
    perms = [p.strip() for p in current_permissions().upper().split(",")]
    return perm.upper() in perms


# ── Decorators ───────────────────────────────────────────────────────────────

def login_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user():
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper


def admin_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user():
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login", next=request.path))
        if not has_permission("ADMIN"):
            flash("Admin access required.", "error")
            return redirect(url_for("home"))
        return f(*args, **kwargs)
    return wrapper


def cbsl_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user():
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login", next=request.path))
        if not has_permission("CBSL"):
            flash("You do not have permission to access the CBSL portal.", "error")
            return redirect(url_for("home"))
        return f(*args, **kwargs)
    return wrapper


def ifrs_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user():
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login", next=request.path))
        if not has_permission("IFRS"):
            flash("You do not have permission to access the IFRS portal.", "error")
            return redirect(url_for("home"))
        return f(*args, **kwargs)
    return wrapper


# ── Bootstrap ─────────────────────────────────────────────────────────────────

def bootstrap_admin():
    """Ensure users.xlsx exists with headers. No credentials are hardcoded."""
    _ensure_file()
    admins = [u for u in list_users_display() if "ADMIN" in u.get("permissions", "").upper()]
    if not admins:
        print("[auth] WARNING: No admin user found in users.xlsx.")
        print("       Add a row manually — use 'python auth.py gen-encrypted' to get encrypted values.")


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    def _usage():
        print("Usage:")
        print("  python auth.py gen-encrypted <username> <password>")
        print("      Prints encrypted un and pw to paste into users.xlsx.")
        print()
        print("  Example:")
        print("  python auth.py gen-encrypted cbslbotuser@assetlinefinance.lk \"2SC,3q{>aU21\"")

    if len(sys.argv) >= 4 and sys.argv[1] == "gen-encrypted":
        un_plain = sys.argv[2]
        pw_plain = sys.argv[3]
        now      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        enc_un   = encrypt(un_plain)
        enc_pw   = encrypt(pw_plain)
        print()
        print("Paste these values into users.xlsx (one row):")
        print(f"  un          : {enc_un}")
        print(f"  pw          : {enc_pw}")
        print(f"  permissions : CBSL,IFRS,ADMIN")
        print(f"  created_at  : {now}")
        print(f"  updated_at  : {now}")
        print()
    else:
        _usage()
