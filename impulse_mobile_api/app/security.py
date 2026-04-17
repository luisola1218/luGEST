from __future__ import annotations

import base64
import hashlib
import hmac
import json
import time
from typing import Any

from .config import settings


def _secret_bytes() -> bytes:
    if not settings.api_secret_configured():
        raise RuntimeError("LUGEST_API_SECRET nao esta configurado com uma chave forte.")
    return str(settings.api_secret or "").encode("utf-8")


def _b64encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _b64decode(raw: str) -> bytes:
    pad = "=" * (-len(raw) % 4)
    return base64.urlsafe_b64decode(raw + pad)


def issue_token(username: str, role: str, ttl_hours: int = 12) -> str:
    payload = {
        "sub": str(username or "").strip(),
        "role": str(role or "").strip(),
        "exp": int(time.time()) + max(1, int(ttl_hours)) * 3600,
    }
    raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    sig = hmac.new(_secret_bytes(), raw, hashlib.sha256).hexdigest()
    return f"{_b64encode(raw)}.{sig}"


def decode_token(token: str) -> dict[str, Any]:
    try:
        data_b64, sig = str(token or "").strip().split(".", 1)
    except ValueError as exc:
        raise ValueError("Token invalido") from exc
    raw = _b64decode(data_b64)
    expected = hmac.new(_secret_bytes(), raw, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expected):
        raise ValueError("Assinatura invalida")
    payload = json.loads(raw.decode("utf-8"))
    if int(payload.get("exp", 0)) < int(time.time()):
        raise ValueError("Token expirado")
    return payload
