"""Configuration persistence with explicit local-store and SQL dependencies.

Reads prefer the shared database; local storage is the offline fallback.
Writes succeed if at least one destination succeeds. No connection is opened
until load/save is called, and reading never creates database tables.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
import json
from typing import Any, Callable, Protocol


class ConfigurationStore(Protocol):
    def load(self, default: dict | None = None) -> dict: ...
    def save(self, payload: dict) -> None: ...


@dataclass(frozen=True)
class ConfigurationResult:
    payload: dict[str, Any]
    error: str = ""


class ConfigurationWriteError(RuntimeError):
    pass


class ConfigurationRepository:
    def __init__(self, local_store: ConfigurationStore, connect: Callable[[], Any] | None = None,
                 *, key: str = "qt_desktop_config") -> None:
        self.local_store = local_store
        self.connect = connect
        self.key = key

    @staticmethod
    def _close(connection) -> None:
        if connection is not None:
            try:
                connection.close()
            except Exception:
                pass

    def load(self) -> ConfigurationResult:
        payload = {}
        connection = None
        errors = []
        if self.connect is not None:
            try:
                connection = self.connect()
                with connection.cursor() as cursor:
                    cursor.execute("SELECT cvalue FROM app_config WHERE ckey=%s LIMIT 1", (self.key,))
                    row = cursor.fetchone()
                if row:
                    raw = row.get("cvalue") if isinstance(row, dict) else row[0]
                    if isinstance(raw, (bytes, bytearray)):
                        raw = raw.decode("utf-8", errors="ignore")
                    parsed = json.loads(str(raw or "{}"))
                    if isinstance(parsed, dict):
                        payload = parsed
            except Exception as exc:
                errors.append(f"MySQL: {exc}")
            finally:
                self._close(connection)
        if not payload:
            try:
                parsed = self.local_store.load(default={})
                if isinstance(parsed, dict):
                    payload = parsed
            except Exception as exc:
                errors.append(f"local: {exc}")
        return ConfigurationResult(deepcopy(payload), "; ".join(errors))

    def save(self, payload: dict[str, Any]) -> ConfigurationResult:
        clean = deepcopy(dict(payload or {}))
        saved = False
        errors = []
        try:
            self.local_store.save(deepcopy(clean))
            saved = True
        except Exception as exc:
            errors.append(f"local: {exc}")
        connection = None
        if self.connect is not None:
            try:
                connection = self.connect()
                with connection.cursor() as cursor:
                    cursor.execute("""
                        CREATE TABLE IF NOT EXISTS app_config (
                            ckey VARCHAR(80) PRIMARY KEY,
                            cvalue LONGTEXT NULL,
                            updated_at DATETIME NULL
                        )
                    """)
                    cursor.execute("""
                        INSERT INTO app_config (ckey, cvalue, updated_at)
                        VALUES (%s, %s, NOW())
                        ON DUPLICATE KEY UPDATE cvalue=VALUES(cvalue), updated_at=VALUES(updated_at)
                    """, (self.key, json.dumps(clean, ensure_ascii=False)))
                connection.commit()
                saved = True
            except Exception as exc:
                errors.append(f"MySQL: {exc}")
                if connection is not None:
                    try:
                        connection.rollback()
                    except Exception:
                        pass
            finally:
                self._close(connection)
        detail = "; ".join(errors)
        if not saved:
            raise ConfigurationWriteError(detail or "nenhum destino de configuração disponível")
        return ConfigurationResult(clean, detail)
