from __future__ import annotations

import json
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

sys.dont_write_bytecode = True


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

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

TRASH_PATHS = [
    ROOT / "build",
    ROOT / "__pycache__",
    ROOT / "dist" / "lugest_trial.json",
    ROOT / ".pytest_cache",
    ROOT / ".mypy_cache",
    ROOT / ".ruff_cache",
    ROOT / "impulse_mobile_app" / ".dart_tool",
    ROOT / "impulse_mobile_app" / ".idea",
    ROOT / "impulse_mobile_app" / "android" / ".gradle",
    ROOT / "impulse_mobile_app" / "android" / ".kotlin",
]


@dataclass(frozen=True)
class Finding:
    severity: str
    area: str
    message: str
    path: Path | None = None


def _load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _load_env(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    if not path.exists():
        return data
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = str(raw or "").strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        clean_key = key.strip().lstrip("\ufeff")
        data[clean_key] = value.strip().strip('"').strip("'")
    return data


def _load_runtime_main():
    try:
        import main as runtime_main

        return runtime_main
    except Exception:
        return None


def _runtime_data() -> dict:
    runtime_main = _load_runtime_main()
    if runtime_main is not None and callable(getattr(runtime_main, "load_data", None)):
        try:
            data = runtime_main.load_data()
            if isinstance(data, dict):
                return data
        except Exception:
            return {}
    return {}


def _is_weak_password(username: str, password: str) -> bool:
    runtime_main = _load_runtime_main()
    if (
        runtime_main is not None
        and callable(getattr(runtime_main, "is_password_hash", None))
        and runtime_main.is_password_hash(password)
        and callable(getattr(runtime_main, "verify_password", None))
    ):
        username_txt = str(username or "").strip()
        candidates = {str(value or "").strip() for value in WEAK_PASSWORDS}
        if username_txt:
            candidates.add(username_txt)
            candidates.add(username_txt.lower())
        for candidate in sorted(value for value in candidates if value):
            try:
                if runtime_main.verify_password(candidate, password):
                    return True
            except Exception:
                continue
        return False
    if runtime_main is not None and callable(getattr(runtime_main, "is_weak_password_value", None)):
        try:
            return bool(runtime_main.is_weak_password_value(username, password))
        except Exception:
            pass
    username_txt = str(username or "").strip().lower()
    password_txt = str(password or "").strip()
    if password_txt.lower() in WEAK_PASSWORDS:
        return True
    if username_txt and password_txt.lower() == username_txt:
        return True
    return len(password_txt) < 8


def _is_placeholder_secret(secret: str) -> bool:
    value = str(secret or "").strip().lower()
    if not value:
        return True
    if value in {"trocar-esta-chave", "trocar-password", "trocar-password-forte", "change-me", "changeme"}:
        return True
    return len(value) < 16


def _audit_data_users(findings: list[Finding]) -> None:
    runtime_main = _load_runtime_main()
    data = _runtime_data()
    users = list(data.get("users", []) or [])
    weak_users = []
    for user in users:
        if not isinstance(user, dict):
            continue
        username = str(user.get("username", "") or "").strip()
        password = str(user.get("password", "") or "").strip()
        if _is_weak_password(username, password):
            weak_users.append(username or "<sem-utilizador>")
    if weak_users:
        findings.append(
            Finding(
                "HIGH",
                "Autenticacao",
                "Existem utilizadores locais com password fraca ou previsivel: " + ", ".join(sorted(set(weak_users))),
                ROOT / "main.py",
            )
        )


def _audit_qt_config(findings: list[Finding]) -> None:
    runtime_main = _load_runtime_main()
    cfg = _load_json(ROOT / "lugest_qt_config.json")
    ui_options = dict(cfg.get("ui_options", {}) or {})
    supervisor_password = str(ui_options.get("operator_supervisor_password", "") or "").strip()
    if runtime_main is not None and callable(getattr(runtime_main, "is_password_hash", None)):
        try:
            if runtime_main.is_password_hash(supervisor_password):
                return
        except Exception:
            pass
    if _is_weak_password("supervisor", supervisor_password):
        findings.append(
            Finding(
                "HIGH",
                "Supervisor",
                "A password de supervisor do Operador esta vazia, curta ou num valor fraco.",
                ROOT / "lugest_qt_config.json",
            )
        )


def _audit_env_files(findings: list[Finding]) -> None:
    runtime_main = _load_runtime_main()
    desktop_env = _load_env(ROOT / "lugest.env")
    api_env = _load_env(ROOT / "impulse_mobile_api" / ".env")
    owner_user = str(desktop_env.get("LUGEST_OWNER_USERNAME", "") or "").strip()
    owner_pass = str(desktop_env.get("LUGEST_OWNER_PASSWORD", "") or "").strip()
    owner_pass_is_hash = bool(runtime_main and callable(getattr(runtime_main, "is_password_hash", None)) and runtime_main.is_password_hash(owner_pass))
    if owner_user and not owner_pass_is_hash and _is_weak_password(owner_user, owner_pass):
        findings.append(
            Finding(
                "HIGH",
                "Trial",
                "A conta proprietaria do trial tem password fraca ou previsivel.",
                ROOT / "lugest.env",
            )
        )
    for path, env_data, label in (
        (ROOT / "lugest.env", desktop_env, "Desktop"),
        (ROOT / "impulse_mobile_api" / ".env", api_env, "API"),
    ):
        db_user = str(env_data.get("LUGEST_DB_USER", "") or "").strip().lower()
        db_pass = str(env_data.get("LUGEST_DB_PASS", "") or "").strip()
        db_host = str(env_data.get("LUGEST_DB_HOST", "") or "").strip()
        if not db_user or not db_pass or not db_host:
            findings.append(
                Finding(
                    "MEDIUM",
                    "Base de Dados",
                    f"{label} sem configuracao DB completa no ficheiro .env/lugest.env.",
                    path,
                )
            )
        elif db_user == "root":
            findings.append(
                Finding(
                    "MEDIUM",
                    "Base de Dados",
                    f"{label} ainda usa o utilizador root na base de dados. Convem usar uma conta dedicada.",
                    path,
                )
            )
    api_secret = str(api_env.get("LUGEST_API_SECRET", "") or "").strip()
    if _is_placeholder_secret(api_secret):
        findings.append(
            Finding(
                "HIGH",
                "API",
                "A API mobile esta sem segredo forte configurado em LUGEST_API_SECRET.",
                ROOT / "impulse_mobile_api" / ".env",
            )
        )
    allowed_origins = str(api_env.get("LUGEST_API_ALLOWED_ORIGINS", "") or "").strip()
    if not allowed_origins:
        findings.append(
            Finding(
                "MEDIUM",
                "API",
                "A API nao tem origens explicitamente definidas em LUGEST_API_ALLOWED_ORIGINS.",
                ROOT / "impulse_mobile_api" / ".env",
            )
        )


def _audit_source_defaults(findings: list[Finding]) -> None:
    runtime_main = _load_runtime_main()
    source_defaults = None
    if runtime_main is not None:
        try:
            source_defaults = list(getattr(runtime_main, "DEFAULT_DATA", {}).get("users", []) or [])
        except Exception:
            source_defaults = None
    if source_defaults:
        findings.append(
            Finding(
                "MEDIUM",
                "Bootstrap",
                "O codigo-fonte ainda contem utilizadores seed com credenciais fracas para instalacoes novas.",
                ROOT / "main.py",
            )
        )
    mysql_script = (ROOT / "scripts" / "apply_mysql_update.py").read_text(encoding="utf-8", errors="ignore")
    if "280874" in mysql_script:
        findings.append(
            Finding(
                "MEDIUM",
                "Scripts",
                "O script de atualizacao MySQL ainda contem uma password hardcoded.",
                ROOT / "scripts" / "apply_mysql_update.py",
            )
        )
    api_main = (ROOT / "impulse_mobile_api" / "app" / "main.py").read_text(encoding="utf-8", errors="ignore")
    if 'allow_origins=["*"]' in api_main.replace(" ", ""):
        findings.append(
            Finding(
                "MEDIUM",
                "API",
                "A API esta com CORS totalmente aberto.",
                ROOT / "impulse_mobile_api" / "app" / "main.py",
            )
        )


def _audit_plaintext_files(findings: list[Finding]) -> None:
    for path in (ROOT / "lugest.env", ROOT / "impulse_mobile_api" / ".env"):
        if path.exists():
            findings.append(
                Finding(
                    "LOW",
                    "Segredos",
                    "Ficheiro sensivel presente no workspace. Confirmar protecao de acesso e backups cifrados.",
                    path,
                )
            )


def _audit_trash(findings: list[Finding]) -> None:
    found = [path for path in TRASH_PATHS if path.exists()]
    if found:
        findings.append(
            Finding(
                "LOW",
                "Higiene",
                f"Foram encontrados {len(found)} caminhos de lixo regeneravel. Convem correr a limpeza antes de entregar.",
                ROOT / "scripts" / "cleanup_workspace.ps1",
            )
        )


def _print_report(findings: list[Finding]) -> int:
    buckets = {"HIGH": [], "MEDIUM": [], "LOW": []}
    for finding in findings:
        buckets.setdefault(finding.severity, []).append(finding)

    print("== LuGEST Security Audit ==")
    for severity in ("HIGH", "MEDIUM", "LOW"):
        print(f"{severity}: {len(buckets.get(severity, []))}")
    print("")
    for severity in ("HIGH", "MEDIUM", "LOW"):
        items = buckets.get(severity, [])
        if not items:
            continue
        print(f"[{severity}]")
        for item in items:
            suffix = f" | {item.path}" if item.path else ""
            print(f"- {item.area}: {item.message}{suffix}")
        print("")
    return 1 if buckets.get("HIGH") else 0


def main() -> int:
    findings: list[Finding] = []
    _audit_data_users(findings)
    _audit_qt_config(findings)
    _audit_env_files(findings)
    _audit_source_defaults(findings)
    _audit_plaintext_files(findings)
    _audit_trash(findings)
    return _print_report(findings)


if __name__ == "__main__":
    raise SystemExit(main())
