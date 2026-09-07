from __future__ import annotations

import json
from copy import deepcopy
from lugest_infra.config import AtomicJsonStore
from lugest_infra.config.repository import ConfigurationRepository, ConfigurationWriteError
from pathlib import Path
from typing import Any


class ConfigurationBackendMixin:
    """Legacy adapter for configuration; see BACKEND_GUIDE.md."""

    def _qt_config_path(self) -> Path:
        target = self.app_paths.config_file("lugest_qt_config.json")
        self.app_paths.migrate_legacy_file(self.base_dir / "lugest_qt_config.json", target)
        return target

    def _qt_config_store(self) -> AtomicJsonStore:
        return AtomicJsonStore(self._qt_config_path())

    def _load_qt_config(self) -> dict[str, Any]:
        if isinstance(self._qt_config_cache, dict):
            return deepcopy(self._qt_config_cache)
        result = self._configuration_repository().load()
        self._qt_config_last_error = result.error
        self._qt_config_cache = deepcopy(result.payload)
        return deepcopy(result.payload)

    def _save_qt_config(self, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            result = self._configuration_repository().save(payload)
        except ConfigurationWriteError as exc:
            self._qt_config_last_error = str(exc)
            raise RuntimeError(f"Não foi possível guardar a configuração do luGEST ({exc}).") from exc
        self._qt_config_last_error = result.error
        self._qt_config_cache = deepcopy(result.payload)
        self._product_taxonomy_nodes_cache = None
        return deepcopy(result.payload)

    def ui_options(self) -> dict[str, Any]:
        defaults = {
            "operator_show_client_name": True,
            "operator_supervisor_password": "",
            "operator_supervisor_password_set": False,
        }
        cfg = self._load_qt_config()
        stored = dict(cfg.get("ui_options", {}) or {})
        supervisor_password = str(stored.get("operator_supervisor_password", "") or "").strip()
        if supervisor_password and not self.desktop_main.is_password_hash(supervisor_password):
            stored["operator_supervisor_password"] = self.desktop_main.normalize_password_for_storage(
                "supervisor",
                supervisor_password,
                require_strong=False,
            )
            cfg["ui_options"] = stored
            self._save_qt_config(cfg)
        safe = {**defaults, **stored}
        safe["operator_supervisor_password_set"] = bool(str(stored.get("operator_supervisor_password", "") or "").strip())
        safe["operator_supervisor_password"] = ""
        return safe

    def set_ui_option(self, key: str, value: Any) -> dict[str, Any]:
        key_txt = str(key or "").strip()
        if not key_txt:
            return self.ui_options()
        cfg = self._load_qt_config()
        options = dict(cfg.get("ui_options", {}) or {})
        if key_txt == "operator_supervisor_password":
            raw_password = str(value or "").strip()
            options[key_txt] = (
                self.desktop_main.normalize_password_for_storage("supervisor", raw_password, require_strong=False)
                if raw_password
                else ""
            )
        else:
            options[key_txt] = value
        cfg["ui_options"] = options
        self._save_qt_config(cfg)
        return self.ui_options()

    def _configuration_repository(self) -> ConfigurationRepository:
        connect = getattr(self.desktop_main, "_mysql_connect", None)
        return ConfigurationRepository(self._qt_config_store(), connect if callable(connect) else None)
