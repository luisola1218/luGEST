from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().lstrip("\ufeff")
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_dotenv(Path(__file__).resolve().parents[1] / ".env")


def _env_int(name: str, default: int) -> int:
    raw = str(os.environ.get(name, default) or "").strip()
    try:
        return int(raw)
    except Exception:
        return int(default)


def _parse_allowed_origins(raw: str) -> tuple[str, ...]:
    values = []
    for item in str(raw or "").replace(";", ",").split(","):
        text = str(item or "").strip()
        if text:
            values.append(text)
    return tuple(values)


@dataclass(frozen=True)
class Settings:
    api_host: str = os.environ.get("LUGEST_API_HOST", "0.0.0.0")
    api_port: int = _env_int("LUGEST_API_PORT", 8050)
    api_secret: str = os.environ.get("LUGEST_API_SECRET", "")
    api_allowed_origins: tuple[str, ...] = field(default_factory=lambda: _parse_allowed_origins(os.environ.get("LUGEST_API_ALLOWED_ORIGINS", "")))
    db_host: str = os.environ.get("LUGEST_DB_HOST", "127.0.0.1")
    db_port: int = _env_int("LUGEST_DB_PORT", 3306)
    db_user: str = os.environ.get("LUGEST_DB_USER", "")
    db_pass: str = os.environ.get("LUGEST_DB_PASS", "")
    db_name: str = os.environ.get("LUGEST_DB_NAME", "lugest")

    def api_secret_configured(self) -> bool:
        secret = str(self.api_secret or "").strip()
        if not secret:
            return False
        if secret.lower() in {"trocar-esta-chave", "change-me", "changeme"}:
            return False
        return len(secret) >= 16

    def db_config_errors(self) -> list[str]:
        issues: list[str] = []
        if not str(self.db_host or "").strip():
            issues.append("Falta definir LUGEST_DB_HOST.")
        if not isinstance(self.db_port, int) or self.db_port < 1 or self.db_port > 65535:
            issues.append("LUGEST_DB_PORT tem um valor invalido.")
        if not str(self.db_user or "").strip():
            issues.append("Falta definir LUGEST_DB_USER.")
        if not str(self.db_pass or "").strip():
            issues.append("Falta definir LUGEST_DB_PASS.")
        if not str(self.db_name or "").strip():
            issues.append("Falta definir LUGEST_DB_NAME.")
        if str(self.db_user or "").strip().lower() == "root":
            issues.append("Evita usar root; cria um utilizador MySQL dedicado para a API.")
        return issues

    def db_configured(self) -> bool:
        return not self.db_config_errors()

    def api_bind_errors(self) -> list[str]:
        issues: list[str] = []
        if not str(self.api_host or "").strip():
            issues.append("Falta definir LUGEST_API_HOST.")
        if not isinstance(self.api_port, int) or self.api_port < 1 or self.api_port > 65535:
            issues.append("LUGEST_API_PORT tem um valor invalido.")
        return issues

    def runtime_errors(self) -> list[str]:
        issues = []
        issues.extend(self.api_bind_errors())
        issues.extend(self.db_config_errors())
        if not self.api_secret_configured():
            issues.append("LUGEST_API_SECRET nao esta configurado com uma chave forte.")
        return issues


settings = Settings()
