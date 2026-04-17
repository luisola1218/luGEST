from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


def _persist(data: dict) -> None:
    main.normalize_notas_encomenda(data)
    main._save_data_now(data, fp=main._save_data_fingerprint(data), token=0, blocking=True)


def main_verify() -> int:
    username = "__verify_admin_setup__"
    password = "Verify#Admin2026!"
    bootstrap_username = "__verify_bootstrap_env__"

    previous = main.load_data()
    original_user = main.find_local_user(previous, username)
    original_bootstrap = main.find_local_user(previous, bootstrap_username)

    previous_env = {
        "LUGEST_BOOTSTRAP_ADMIN_USERNAME": os.environ.get("LUGEST_BOOTSTRAP_ADMIN_USERNAME"),
        "LUGEST_BOOTSTRAP_ADMIN_PASSWORD": os.environ.get("LUGEST_BOOTSTRAP_ADMIN_PASSWORD"),
        "LUGEST_BOOTSTRAP_ADMIN_ROLE": os.environ.get("LUGEST_BOOTSTRAP_ADMIN_ROLE"),
    }

    try:
        proc = subprocess.run(
            [
                sys.executable,
                str(ROOT / "main.py"),
                "--setup-admin",
                "--admin-username",
                username,
                "--admin-password",
                password,
                "--admin-role",
                "Admin",
                "--reset-admin",
            ],
            cwd=ROOT,
            text=True,
            capture_output=True,
        )
        if proc.returncode != 0:
            raise AssertionError(f"CLI de setup admin falhou: {proc.stdout}\n{proc.stderr}")
        if "admin-setup-ok" not in str(proc.stdout or ""):
            raise AssertionError(f"CLI de setup admin nao devolveu confirmacao esperada: {proc.stdout}")

        data = main.load_data()
        row = main.find_local_user(data, username)
        if not isinstance(row, dict):
            raise AssertionError("Utilizador criado nao encontrado na base.")
        if str(row.get("role", "") or "").strip() != "Admin":
            raise AssertionError("Role do admin nao ficou correta.")
        stored_password = str(row.get("password", "") or "").strip()
        if not stored_password or not main.is_password_hash(stored_password):
            raise AssertionError("Password do admin nao ficou guardada em hash.")
        auth_row = main.authenticate_local_user(data, username, password, persist_upgrade=False)
        if not isinstance(auth_row, dict):
            raise AssertionError("Nao foi possivel autenticar o admin criado.")

        os.environ["LUGEST_BOOTSTRAP_ADMIN_USERNAME"] = bootstrap_username
        os.environ["LUGEST_BOOTSTRAP_ADMIN_PASSWORD"] = "Bootstrap#Admin2026!"
        os.environ["LUGEST_BOOTSTRAP_ADMIN_ROLE"] = "Admin"
        data_after_bootstrap = main.load_data()
        if main.find_local_user(data_after_bootstrap, bootstrap_username) is not None:
            raise AssertionError("Bootstrap ENV nao devia criar utilizador quando a base ja tem users.")

        print("verify-admin-setup-ok")
        return 0
    finally:
        for key, value in previous_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value

        restored = main.load_data()
        users = [row for row in list(restored.get("users", []) or []) if str(row.get("username", "") or "").strip().lower() not in {username.lower(), bootstrap_username.lower()}]
        if isinstance(original_user, dict):
            users.append(original_user)
        if isinstance(original_bootstrap, dict):
            users.append(original_bootstrap)
        restored["users"] = users
        _persist(restored)


if __name__ == "__main__":
    raise SystemExit(main_verify())
