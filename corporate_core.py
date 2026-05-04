from __future__ import annotations

import hashlib
import json
import os
import secrets
import sqlite3
import zipfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CORP_DIR = DATA_DIR / "corporativo"
BACKUP_DIR = DATA_DIR / "backups"
DB_FILE = CORP_DIR / "bi_qualidade_corporativo.db"

MODULE_DIRS = [
    DATA_DIR / "qualidade",
    DATA_DIR / "sqdcp",
    DATA_DIR / "cpk",
    CORP_DIR,
    BACKUP_DIR,
]

DEFAULT_PERMISSIONS = {
    "admin": ["indicadores", "sqdcp", "cpk", "admin"],
    "qualidade": ["indicadores", "cpk"],
    "producao": ["sqdcp"],
    "consulta": ["indicadores", "sqdcp", "cpk"],
}


def ensure_structure() -> None:
    for p in MODULE_DIRS:
        p.mkdir(parents=True, exist_ok=True)


def _hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 180000)
    return salt, digest.hex()


@contextmanager
def db_conn():
    ensure_structure()
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    ensure_structure()
    with db_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                display_name TEXT NOT NULL,
                role TEXT NOT NULL,
                salt TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                last_login TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_at TEXT NOT NULL,
                username TEXT,
                module TEXT,
                action TEXT NOT NULL,
                detail TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        count = conn.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
        if count == 0:
            salt, pwd = _hash_password("QualidadeRS2026")
            conn.execute(
                "INSERT INTO users(username, display_name, role, salt, password_hash, active, created_at) VALUES(?,?,?,?,?,?,?)",
                ("admin", "Administrador", "admin", salt, pwd, 1, datetime.now().isoformat(timespec="seconds")),
            )
            conn.execute(
                "INSERT INTO audit_log(event_at, username, module, action, detail) VALUES(?,?,?,?,?)",
                (datetime.now().isoformat(timespec="seconds"), "system", "admin", "create_default_admin", "Usuário admin inicial criado."),
            )


def authenticate(username: str, password: str) -> dict[str, Any] | None:
    init_db()
    username = (username or "").strip()
    with db_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        if not row:
            audit(username, "login", "login_failed", "Usuário inexistente ou inativo")
            return None
        salt, digest = _hash_password(password, row["salt"])
        if digest != row["password_hash"]:
            audit(username, "login", "login_failed", "Senha inválida")
            return None
        conn.execute("UPDATE users SET last_login=? WHERE username=?", (datetime.now().isoformat(timespec="seconds"), username))
    audit(username, "login", "login_success", f"Perfil: {row['role']}")
    return {"username": row["username"], "display_name": row["display_name"], "role": row["role"]}


def permissions_for(role: str) -> list[str]:
    return DEFAULT_PERMISSIONS.get(role, [])


def can_access(user: dict[str, Any] | None, module: str) -> bool:
    if not user:
        return False
    return module in permissions_for(user.get("role", ""))


def audit(username: str | None, module: str | None, action: str, detail: str | None = None) -> None:
    try:
        with db_conn() as conn:
            conn.execute(
                "INSERT INTO audit_log(event_at, username, module, action, detail) VALUES(?,?,?,?,?)",
                (datetime.now().isoformat(timespec="seconds"), username, module, action, detail),
            )
    except Exception:
        pass


def list_users() -> list[dict[str, Any]]:
    init_db()
    with db_conn() as conn:
        rows = conn.execute("SELECT username, display_name, role, active, created_at, last_login FROM users ORDER BY username").fetchall()
    return [dict(r) for r in rows]


def create_or_update_user(username: str, display_name: str, role: str, password: str | None, active: bool = True) -> None:
    init_db()
    username = username.strip()
    display_name = display_name.strip() or username
    role = role.strip()
    with db_conn() as conn:
        existing = conn.execute("SELECT username FROM users WHERE username=?", (username,)).fetchone()
        if existing:
            if password:
                salt, pwd = _hash_password(password)
                conn.execute("UPDATE users SET display_name=?, role=?, salt=?, password_hash=?, active=? WHERE username=?", (display_name, role, salt, pwd, int(active), username))
            else:
                conn.execute("UPDATE users SET display_name=?, role=?, active=? WHERE username=?", (display_name, role, int(active), username))
        else:
            if not password:
                raise ValueError("Senha obrigatória para novo usuário.")
            salt, pwd = _hash_password(password)
            conn.execute(
                "INSERT INTO users(username, display_name, role, salt, password_hash, active, created_at) VALUES(?,?,?,?,?,?,?)",
                (username, display_name, role, salt, pwd, int(active), datetime.now().isoformat(timespec="seconds")),
            )
    audit(None, "admin", "upsert_user", f"Usuário: {username}; perfil: {role}; ativo: {active}")


def recent_audit(limit: int = 100) -> list[dict[str, Any]]:
    init_db()
    with db_conn() as conn:
        rows = conn.execute("SELECT event_at, username, module, action, detail FROM audit_log ORDER BY id DESC LIMIT ?", (int(limit),)).fetchall()
    return [dict(r) for r in rows]


def make_backup_zip() -> Path:
    ensure_structure()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = BACKUP_DIR / f"backup_bi_qualidade_{stamp}.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder in [DATA_DIR, BASE_DIR / "modulos", BASE_DIR / "manual"]:
            if folder.exists():
                for path in folder.rglob("*"):
                    if path.is_file() and path.resolve() != out.resolve():
                        zf.write(path, path.relative_to(BASE_DIR))
        for filename in ["app.py", "corporate_core.py", "requirements.txt"]:
            path = BASE_DIR / filename
            if path.exists():
                zf.write(path, path.relative_to(BASE_DIR))
    audit(None, "admin", "backup_created", out.name)
    return out


def read_setting(key: str, default: Any = None) -> Any:
    init_db()
    with db_conn() as conn:
        row = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row["value"])
    except Exception:
        return row["value"]


def write_setting(key: str, value: Any) -> None:
    init_db()
    data = json.dumps(value, ensure_ascii=False)
    with db_conn() as conn:
        conn.execute(
            "INSERT INTO app_settings(key, value, updated_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
            (key, data, datetime.now().isoformat(timespec="seconds")),
        )
