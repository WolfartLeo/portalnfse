# scripts/create_user.py
import json
import sys
from passlib.hash import bcrypt

USERS_PATH = "users.json"

def main():
    if len(sys.argv) < 4:
        print("Uso: python scripts/create_user.py <username> <nome> <senha> [role]")
        sys.exit(1)

    username = sys.argv[1].strip()
    nome = sys.argv[2].strip()
    senha = sys.argv[3]
    role = sys.argv[4].strip() if len(sys.argv) >= 5 else "user"

    try:
        with open(USERS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        data = {"users": {}}

    data.setdefault("users", {})
    data["users"][username] = {
        "name": nome,
        "role": role,
        "password_hash": bcrypt.hash(senha),
    }

    with open(USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"OK: usuário '{username}' criado/atualizado com role='{role}'.")

if __name__ == "__main__":
    main()
