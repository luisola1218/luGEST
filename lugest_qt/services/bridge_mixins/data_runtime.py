from __future__ import annotations

import copy
import json
import time
from datetime import datetime
from lugest_core.snapshots import SnapshotMergePolicy
from typing import Any


class DataRuntimeBackendMixin:
    """Legacy adapter for data runtime; see BACKEND_GUIDE.md."""

    def _clone_data(self, data: dict[str, Any] | None) -> dict[str, Any] | None:
        return SnapshotMergePolicy().clone_mapping(data)

    def _replace_data_cache(self, data: dict[str, Any]) -> dict[str, Any]:
        self.data = data
        self._base_data_snapshot = self._clone_data(data)
        self._data_loaded_at = time.time()
        self._data_cache_generation += 1
        self._operation_catalog_cache = None
        try:
            self.desktop_main._RUNTIME_DATA_REF = self.data
            self.desktop_main._LATEST_RUNTIME_DATA = self.data
        except Exception:
            pass
        return data

    def _bucket_signature(self, value: Any) -> str:
        return SnapshotMergePolicy().signature(value)

    def _changed_data_buckets(self, current: dict[str, Any], base: dict[str, Any] | None) -> list[str]:
        return SnapshotMergePolicy().changed_buckets(current, base)

    def _bucket_identity_field(self, bucket_name: str) -> str:
        return SnapshotMergePolicy().identity_field(bucket_name)

    def _merge_list_bucket_by_identity(self, bucket_name: str, current_value: Any, base_value: Any, latest_value: Any) -> list[Any] | None:
        return SnapshotMergePolicy().merge_list_bucket(bucket_name, current_value, base_value, latest_value)

    def _merge_latest_for_save(self) -> tuple[dict[str, Any], list[str]]:
        current = self.ensure_data()
        changed_keys = self._changed_data_buckets(current, self._base_data_snapshot)
        if not changed_keys:
            return current, []
        latest = self.desktop_main.load_data()
        for key in changed_keys:
            merged_bucket = self._merge_list_bucket_by_identity(
                key,
                current.get(key),
                (self._base_data_snapshot or {}).get(key) if isinstance(self._base_data_snapshot, dict) else None,
                latest.get(key),
            )
            if merged_bucket is not None:
                latest[key] = merged_bucket
            else:
                latest[key] = self._clone_data(current.get(key)) if isinstance(current.get(key), dict) else copy.deepcopy(current.get(key))
        for key, value in list(current.items()):
            if str(key or "").startswith("__"):
                latest[key] = value
        return latest, changed_keys

    def ensure_data(self) -> dict[str, Any]:
        if not isinstance(self.data, dict):
            self._replace_data_cache(self.desktop_main.load_data())
        return self.data

    def reload(self, *, force: bool = False, max_age_sec: float | None = None) -> dict[str, Any]:
        ttl = self._reload_cache_ttl_sec if max_age_sec is None else max(0.0, float(max_age_sec or 0.0))
        if (
            not force
            and isinstance(self.data, dict)
            and ttl > 0
            and self._data_loaded_at > 0
            and (time.time() - self._data_loaded_at) <= ttl
        ):
            return self.data
        return self._replace_data_cache(self.desktop_main.load_data())

    def data_cache_needs_reload(self, *, max_age_sec: float | None = None) -> bool:
        ttl = self._reload_cache_ttl_sec if max_age_sec is None else max(0.0, float(max_age_sec or 0.0))
        return not (
            isinstance(self.data, dict)
            and ttl > 0
            and self._data_loaded_at > 0
            and (time.time() - self._data_loaded_at) <= ttl
        )

    def data_cache_generation(self) -> int:
        return int(self._data_cache_generation)

    def load_data_snapshot(self) -> dict[str, Any]:
        return self.desktop_main.load_data()

    def apply_data_snapshot(self, data: dict[str, Any], *, expected_generation: int) -> bool:
        if int(expected_generation) != self._data_cache_generation:
            return False
        self._replace_data_cache(data)
        return True

    def save_runtime_state(self) -> dict[str, Any]:
        return {
            "async_enabled": bool(getattr(self.desktop_main, "_ASYNC_SAVE_ENABLED", False)),
            "pending": bool(
                getattr(self.desktop_main, "_PENDING_SAVE_DATA", None) is not None
                or getattr(self.desktop_main, "_ASYNC_SAVE_PENDING_DATA", None) is not None
            ),
            "in_progress": bool(getattr(self.desktop_main, "_ASYNC_SAVE_IN_PROGRESS", False)),
            "last_error": str(getattr(self.desktop_main, "_ASYNC_SAVE_LAST_ERROR", "") or ""),
        }

    def flush_pending_save(self, force: bool = False) -> bool:
        flusher = getattr(self.desktop_main, "flush_pending_save", None)
        if not callable(flusher):
            return False
        return bool(flusher(force=force))

    def drain_async_saves(self, timeout_sec: float = 12.0) -> bool:
        drainer = getattr(self.desktop_main, "_drain_async_saves", None)
        if not callable(drainer):
            return True
        return bool(drainer(timeout_sec=timeout_sec))

    def consume_async_save_error(self) -> str:
        getter = getattr(self.desktop_main, "_consume_async_save_error", None)
        if not callable(getter):
            return ""
        return str(getter() or "").strip()

    def stop_async_save_worker(self, timeout_sec: float = 1.0) -> None:
        stop_evt = getattr(self.desktop_main, "_ASYNC_SAVE_STOP", None)
        save_evt = getattr(self.desktop_main, "_ASYNC_SAVE_EVENT", None)
        thread = getattr(self.desktop_main, "_ASYNC_SAVE_THREAD", None)
        try:
            if stop_evt is not None:
                stop_evt.set()
        except Exception:
            pass
        try:
            if save_evt is not None:
                save_evt.set()
        except Exception:
            pass
        try:
            if thread is not None and thread.is_alive():
                thread.join(timeout=max(0.1, float(timeout_sec or 0)))
        except Exception:
            pass

    def _current_user_label(self) -> str:
        user = dict(self.user or {})
        username = str(user.get("username", "") or "").strip()
        role = str(user.get("role", "") or "").strip()
        if username and role:
            return f"{username} | {role}"
        return username or role or "Sistema"

    def _append_audit_event(
        self,
        data: dict[str, Any],
        *,
        action: str,
        entity_type: str = "",
        entity_id: str = "",
        summary: str = "",
        before: Any = None,
        after: Any = None,
    ) -> dict[str, Any]:
        if not isinstance(data, dict):
            return {}
        event = {
            "id": f"AUD-{int(time.time() * 1000)}-{len(list(data.get('audit_log', []) or [])) + 1:04d}",
            "created_at": str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds")),
            "user": self._current_user_label(),
            "action": str(action or "Atualizacao").strip() or "Atualizacao",
            "entity_type": str(entity_type or "").strip(),
            "entity_id": str(entity_id or "").strip(),
            "summary": str(summary or "").strip(),
        }
        if before is not None:
            event["before"] = self._json_safe_clone(before)
        if after is not None:
            event["after"] = self._json_safe_clone(after)
        log = list(data.get("audit_log", []) or [])
        log.append(event)
        data["audit_log"] = log[-3000:]
        return event

    def _save(self, force: bool = False, audit: bool = True, blocking: bool = False) -> None:
        async_enabled = bool(getattr(self.desktop_main, "_ASYNC_SAVE_ENABLED", False))
        current = self.ensure_data()
        initial_changed = self._changed_data_buckets(current, self._base_data_snapshot)
        if not initial_changed and not force:
            return
        if initial_changed:
            self._normalize_storage_paths_for_save(initial_changed)
        if async_enabled:
            payload = current
            _changed = self._changed_data_buckets(payload, self._base_data_snapshot)
        else:
            payload, _changed = self._merge_latest_for_save()
        changed = [str(key) for key in list(_changed or []) if str(key or "") != "audit_log"]
        if audit and changed:
            self._append_audit_event(
                payload,
                action="Guardar dados",
                entity_type="Sistema",
                entity_id=",".join(changed[:8]),
                summary=f"Buckets alterados: {', '.join(changed[:8])}{'...' if len(changed) > 8 else ''}",
            )
        self.desktop_main.save_data(payload, force=bool(force and (blocking or not async_enabled)))
        if isinstance(payload, dict):
            self._replace_data_cache(payload)
        if "materiais" in changed or "produtos" in changed:
            self.ensure_inventory_scan_codes(persist=True)
