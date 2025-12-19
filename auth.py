# auth.py
import json
import os
import shutil

SECRETS_DIR = os.environ.get("SECRETS_DIR", "/etc/secrets")
from dataclasses import dataclass
from typing import Dict, Optional

from passlib.hash import pbkdf2_sha256

USERS_PATH_DEFAULT = "users.json"

@dataclass
class User:
    username: str
    name: str
    role: str


def _ensure_default_admin(users_path: str) -> None:
    """
    Cria o arquivo users.json caso não exista.
    Hash: PBKDF2-SHA256 (padrão estável, sem dependência nativa no Windows/Py3.13).
    """
    if os.path.exists(users_path):
        return

    default = {
        "users": {
            "admin": {
                "name": "Administrador",
                "role": "admin",
                "password_hash": pbkdf2_sha256.hash("admin123"),
            }
        }
    }
    with open(users_path, "w", encoding="utf-8") as f:
        json.dump(default, f, ensure_ascii=False, indent=2)


def load_users(users_path: str = USERS_PATH_DEFAULT) -> Dict:
    _ensure_default_admin(users_path)
    with open(users_path, "r", encoding="utf-8") as f:
        return json.load(f)


def authenticate(username: str, password: str, users_path: str = USERS_PATH_DEFAULT) -> Optional[User]:
    data = load_users(users_path)
    rec = (data.get("users") or {}).get(username)
    if not rec:
        return None

    stored_hash = rec.get("password_hash", "") or ""

    # Governança: se já existir users.json legado com bcrypt, não valida aqui.
    # Procedimento padrão: apagar/renomear users.json e recriar.
    if isinstance(stored_hash, str) and stored_hash.startswith("$2"):
        raise RuntimeError(
            "users.json usa hash bcrypt legado. Apague/renomeie users.json para recriar o ADMIN com hash PBKDF2."
        )

    if not pbkdf2_sha256.verify(password, stored_hash):
        return None

    return User(username=username, name=rec.get("name", username), role=rec.get("role", "user"))
