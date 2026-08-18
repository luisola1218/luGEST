from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
import threading
import time
import uuid
from collections import defaultdict, deque
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


class DeviceRegistry:
    """Registo local de dispositivos; guarda apenas hashes dos tokens."""

    def __init__(self, path: Path):
        self.path = path
        self._lock = threading.RLock()
        self._last_write: dict[str, float] = {}

    def create(self, name: str) -> tuple[dict[str, Any], str]:
        clean_name = str(name or "").strip()[:120]
        if not clean_name:
            raise ValueError("Indica um nome para o dispositivo.")
        device_id = uuid.uuid4().hex[:16]
        token = f"lgf_{device_id}_{secrets.token_urlsafe(48)}"
        record = {
            "id": device_id,
            "name": clean_name,
            "token_hash": self._hash(token),
            "enabled": True,
            "created_at": utc_now(),
            "last_seen_at": None,
        }
        with self._lock:
            payload = self._load()
            payload["devices"].append(record)
            self._save(payload)
        return self._public(record), token

    def import_token(self, name: str, token: str) -> dict[str, Any]:
        clean_name = str(name or "").strip()[:120]
        candidate = str(token or "").strip()
        if not clean_name or len(candidate) < 32:
            raise ValueError("Nome ou chave inválidos.")
        digest = self._hash(candidate)
        with self._lock:
            payload = self._load()
            for existing in payload["devices"]:
                if hmac.compare_digest(str(existing.get("token_hash") or ""), digest):
                    return self._public(existing)
            record = {
                "id": uuid.uuid4().hex[:16],
                "name": clean_name,
                "token_hash": digest,
                "enabled": True,
                "created_at": utc_now(),
                "imported": True,
                "last_seen_at": None,
            }
            payload["devices"].append(record)
            self._save(payload)
            return self._public(record)

    def list(self) -> list[dict[str, Any]]:
        with self._lock:
            return [self._public(row) for row in self._load()["devices"]]

    def revoke(self, device_id: str) -> dict[str, Any]:
        target = str(device_id or "").strip()
        with self._lock:
            payload = self._load()
            for record in payload["devices"]:
                if str(record.get("id")) == target:
                    record["enabled"] = False
                    record["revoked_at"] = utc_now()
                    self._save(payload)
                    return self._public(record)
        raise ValueError("Dispositivo não encontrado.")

    def rotate(self, device_id: str) -> tuple[dict[str, Any], str]:
        target = str(device_id or "").strip()
        with self._lock:
            payload = self._load()
            for record in payload["devices"]:
                if str(record.get("id")) == target:
                    token = f"lgf_{target}_{secrets.token_urlsafe(48)}"
                    record["token_hash"] = self._hash(token)
                    record["enabled"] = True
                    record["rotated_at"] = utc_now()
                    record.pop("revoked_at", None)
                    self._save(payload)
                    return self._public(record), token
        raise ValueError("Dispositivo não encontrado.")

    def authenticate(self, token: str) -> dict[str, Any] | None:
        candidate = str(token or "").strip()
        if not candidate:
            return None
        digest = self._hash(candidate)
        matched: dict[str, Any] | None = None
        with self._lock:
            payload = self._load()
            for record in payload["devices"]:
                stored = str(record.get("token_hash") or "")
                if stored and hmac.compare_digest(stored, digest):
                    matched = record
            if matched is None or not bool(matched.get("enabled", False)):
                return None
            device_id = str(matched.get("id") or "")
            now = time.monotonic()
            if now - self._last_write.get(device_id, 0.0) >= 60:
                matched["last_seen_at"] = utc_now()
                self._last_write[device_id] = now
                self._save(payload)
            return self._public(matched)

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {"version": 1, "devices": []}
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError("O registo de dispositivos está danificado.") from exc
        devices = payload.get("devices", [])
        if not isinstance(devices, list):
            raise RuntimeError("O registo de dispositivos está danificado.")
        return {"version": 1, "devices": devices}

    def _save(self, payload: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temporary, self.path)

    @staticmethod
    def _hash(token: str) -> str:
        return hashlib.sha256(token.encode("utf-8")).hexdigest()

    @staticmethod
    def _public(record: dict[str, Any]) -> dict[str, Any]:
        return {key: value for key, value in record.items() if key != "token_hash"}


class RateLimiter:
    def __init__(self, requests: int = 180, window_seconds: int = 60):
        self.requests = max(10, int(requests))
        self.window_seconds = max(10, int(window_seconds))
        self._events: dict[str, deque[float]] = defaultdict(deque)
        self._lock = threading.Lock()

    def allow(self, key: str) -> bool:
        now = time.monotonic()
        cutoff = now - self.window_seconds
        with self._lock:
            events = self._events[str(key or "-")]
            while events and events[0] < cutoff:
                events.popleft()
            if len(events) >= self.requests:
                return False
            events.append(now)
            return True
