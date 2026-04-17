from __future__ import annotations

import argparse
import json
import secrets
import string
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
BACKUP_DIR = ROOT / "backups"
WEAK_PASSWORDS = {
    "",
    "admin",
    "1234",
    "12345",
    "123456",
    "12345678",
    "123123",
    "password",
    "producao",
    "qualidade",
    "planeamento",
    "orcamentista",
    "operador",
    "test",
}
ALPHABET = string.ascii_letters + string.digits + "!@#$%*+-_?"


def is_weak_password(username: str, password: str) -> bool:
    username_txt = str(username or "").strip().lower()
    password_txt = str(password or "").strip()
    if password_txt.lower() in WEAK_PASSWORDS:
        return True
    if username_txt and password_txt.lower() == username_txt:
        return True
    return len(password_txt) < 8


def generate_password(length: int = 18) -> str:
    while True:
        value = "".join(secrets.choice(ALPHABET) for _ in range(max(16, int(length))))
        if (
            any(ch.islower() for ch in value)
            and any(ch.isupper() for ch in value)
            and any(ch.isdigit() for ch in value)
            and any(ch in "!@#$%*+-_?" for ch in value)
        ):
            return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Rotate weak local user passwords in the LuGEST runtime storage.")
    parser.add_argument("--write", action="store_true", help="Apply changes to the runtime storage.")
    parser.add_argument("--length", type=int, default=18, help="Generated password length.")
    args = parser.parse_args()

    try:
        import main as runtime_main
    except Exception as exc:
        print(f"Nao foi possivel carregar o runtime LuGEST: {exc}")
        return 2

    try:
        data = runtime_main.load_data()
    except Exception as exc:
        print(f"Nao foi possivel carregar os dados atuais: {exc}")
        return 2
    users = list(data.get("users", []) or [])
    weak_users = []
    for user in users:
        if not isinstance(user, dict):
            continue
        username = str(user.get("username", "") or "").strip()
        password = str(user.get("password", "") or "").strip()
        if callable(getattr(runtime_main, "is_password_hash", None)) and runtime_main.is_password_hash(password):
            candidates = {str(value or "").strip() for value in WEAK_PASSWORDS}
            if username:
                candidates.add(username)
                candidates.add(username.lower())
            weak = any(runtime_main.verify_password(candidate, password) for candidate in sorted(value for value in candidates if value))
        elif callable(getattr(runtime_main, "is_weak_password_value", None)):
            weak = bool(runtime_main.is_weak_password_value(username, password))
        else:
            weak = is_weak_password(username, password)
        if weak:
            weak_users.append(user)

    if not weak_users:
        print("Nao foram encontrados utilizadores com password fraca.")
        return 0

    print("Utilizadores com password fraca detetados:")
    for user in weak_users:
        print(f"- {str(user.get('username', '') or '').strip()}")

    if not args.write:
        print("")
        print("Modo simulacao. Usa --write para gerar novas passwords e gravar um relatorio em backups.")
        return 1

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = BACKUP_DIR / f"lugest_runtime_before_password_rotation_{stamp}.json"
    report_file = BACKUP_DIR / f"lugest_password_rotation_{stamp}.txt"
    backup_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    report_lines = [
        "Rotacao de passwords locais LuGEST",
        f"Data: {datetime.now().isoformat(timespec='seconds')}",
        "",
    ]
    for user in weak_users:
        username = str(user.get("username", "") or "").strip()
        new_password = generate_password(args.length)
        user["password"] = runtime_main.hash_password(new_password)
        report_lines.append(f"{username}: {new_password}")

    runtime_main.save_data(data, force=True)
    report_file.write_text("\n".join(report_lines) + "\n", encoding="utf-8")
    print("Passwords atualizadas no armazenamento runtime (MySQL / configuracao ativa).")
    print(f"Backup anterior gravado em: {backup_file}")
    print(f"Relatorio com novas passwords gravado em: {report_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
