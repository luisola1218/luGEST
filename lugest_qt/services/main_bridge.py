from __future__ import annotations

import copy
import csv
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from lugest_infra.storage import files as lugest_storage
from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import fit_font_size as _pdf_fit_font_size
from lugest_infra.pdf.text import mix_hex as _pdf_mix_hex
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text

from lugest_core.laser.quote_engine import estimate_laser_quote, estimate_profile_laser_quote, merge_laser_quote_settings
from .bridge_mixins import (
    BillingBridgeMixin,
    DashboardBridgeMixin,
    PlanningBridgeMixin,
    PurchasingBridgeMixin,
    QuotesBridgeMixin,
    ShippingBridgeMixin,
    TransportBridgeMixin,
)


_STEEL_DENSITY_G_CM3 = 7.85

_PROFILE_STANDARD_KG_M: dict[str, dict[str, float]] = {
    "IPN": {
        "80": 5.94,
        "100": 8.34,
        "120": 11.1,
        "140": 14.3,
        "160": 17.9,
        "180": 21.9,
        "200": 26.2,
        "220": 31.1,
        "240": 36.2,
        "260": 41.9,
        "280": 47.9,
        "300": 54.2,
        "320": 61.0,
    },
    "IPE": {
        "80": 6.0,
        "100": 8.1,
        "120": 10.4,
        "140": 12.9,
        "160": 15.8,
        "180": 18.8,
        "200": 22.4,
        "220": 26.2,
        "240": 30.7,
        "270": 36.1,
        "300": 42.2,
        "330": 49.1,
        "360": 57.1,
        "400": 66.3,
        "450": 77.6,
        "500": 90.7,
        "550": 106.0,
        "600": 122.0,
    },
    "HEA": {
        "100": 16.7,
        "120": 19.9,
        "140": 24.7,
        "160": 30.4,
        "180": 35.5,
        "200": 42.3,
        "220": 50.5,
        "240": 60.3,
        "260": 68.2,
        "280": 76.4,
        "300": 88.3,
        "320": 97.6,
        "340": 105.0,
        "360": 112.0,
        "400": 125.0,
        "450": 140.0,
        "500": 155.0,
        "550": 166.0,
        "600": 178.0,
        "650": 190.0,
        "700": 204.0,
        "800": 224.0,
        "900": 252.0,
        "1000": 272.0,
    },
    "HEB": {
        "100": 20.4,
        "120": 26.7,
        "140": 33.7,
        "160": 42.6,
        "180": 51.2,
        "200": 61.3,
        "220": 71.5,
        "240": 83.2,
        "260": 93.0,
        "280": 103.0,
        "300": 117.0,
        "320": 127.0,
        "340": 134.0,
        "360": 142.0,
        "400": 155.0,
        "450": 171.0,
        "500": 187.0,
        "550": 199.0,
        "600": 212.0,
        "650": 225.0,
        "700": 241.0,
        "800": 262.0,
        "900": 291.0,
        "1000": 314.0,
    },
    "HEM": {
        "100": 41.8,
        "120": 52.1,
        "140": 63.2,
        "160": 76.2,
        "180": 88.9,
        "200": 103.0,
        "220": 117.0,
        "240": 157.0,
        "260": 172.0,
        "280": 189.0,
        "300": 238.0,
        "320": 245.0,
        "340": 248.0,
        "360": 250.0,
        "400": 256.0,
        "450": 263.0,
        "500": 270.0,
        "550": 278.0,
        "600": 285.0,
        "650": 293.0,
        "700": 301.0,
        "800": 317.0,
        "900": 333.0,
        "1000": 349.0,
    },
    "UPN": {
        "80": 8.65,
        "100": 10.6,
        "120": 13.4,
        "140": 16.0,
        "160": 18.8,
        "180": 22.0,
        "200": 25.3,
        "220": 29.4,
        "240": 33.2,
        "260": 37.9,
        "280": 41.8,
        "300": 46.2,
        "320": 59.5,
        "400": 71.8,
    },
}

_TUBE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "quadrado", "label": "Quadrado"},
    {"key": "retangular", "label": "Retangular"},
    {"key": "redondo", "label": "Redondo"},
]

_PROFILE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "IPN", "label": "IPN (tabela ACAIL)", "catalog": True},
    {"key": "IPE", "label": "IPE (tabela ACAIL)", "catalog": True},
    {"key": "HEA", "label": "HEA (tabela ACAIL)", "catalog": True},
    {"key": "HEB", "label": "HEB (tabela ACAIL)", "catalog": True},
    {"key": "HEM", "label": "HEM (tabela ACAIL)", "catalog": True},
    {"key": "UPN", "label": "UPN (tabela ACAIL)", "catalog": True},
    {"key": "L", "label": "L (kg/m manual)", "catalog": False},
    {"key": "T", "label": "T (kg/m manual)", "catalog": False},
    {"key": "U", "label": "U genérico (kg/m manual)", "catalog": False},
    {"key": "I", "label": "I genérico (kg/m manual)", "catalog": False},
    {"key": "H", "label": "H genérico (kg/m manual)", "catalog": False},
    {"key": "OUTRO", "label": "Outro / manual", "catalog": False},
]

_ANGLE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "abas_iguais", "label": "Abas iguais"},
    {"key": "abas_desiguais", "label": "Abas desiguais"},
]

_BAR_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "chata", "label": "Barra chata"},
    {"key": "quadrada", "label": "Barra quadrada"},
    {"key": "retangular", "label": "Barra retangular"},
]


def _profile_catalog_lookup_key(value: Any) -> str:
    token = str(value or "").strip().upper()
    return token if token in _PROFILE_STANDARD_KG_M else ""


def _profile_size_lookup_key(value: Any) -> str:
    text = str(value or "").strip().replace(",", ".")
    if not text:
        return ""
    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if not match:
        return ""
    try:
        number = float(match.group(0))
    except Exception:
        return ""
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    return f"{number:.3f}".rstrip("0").rstrip(".")


def _detect_profile_catalog_from_text(value: Any) -> tuple[str, str]:
    raw = str(value or "").upper()
    match = re.search(r"\b(IPN|IPE|HEA|HEB|HEM|UPN)\s*[- ]?(\d{2,4})\b", raw)
    if not match:
        return "", ""
    return str(match.group(1) or "").strip(), str(match.group(2) or "").strip()


class _ValueHolder:
    def __init__(self, value: str = "") -> None:
        self._value = str(value or "")

    def get(self) -> str:
        return self._value


class LegacyBackend(
    BillingBridgeMixin,
    PurchasingBridgeMixin,
    QuotesBridgeMixin,
    PlanningBridgeMixin,
    ShippingBridgeMixin,
    TransportBridgeMixin,
    DashboardBridgeMixin,
):
    def __init__(self) -> None:
        from lugest_desktop.legacy import app_misc_actions
        from lugest_infra.pdf import billing_invoice as billing_pdf_actions
        import main as desktop_main
        from lugest_desktop.legacy import encomendas_actions
        from lugest_desktop.legacy import materia_actions
        from lugest_desktop.legacy import ne_expedicao_actions
        from lugest_desktop.legacy import operador_ordens_actions
        from lugest_desktop.legacy import orc_actions
        from lugest_desktop.legacy import plan_actions
        from lugest_desktop.legacy import produtos_actions
        try:
            from lugest_core.compliance import tax as tax_compliance
        except Exception:
            tax_compliance = SimpleNamespace()

        app_misc_actions.configure(desktop_main.__dict__)
        encomendas_actions.configure(desktop_main.__dict__)
        materia_actions.configure(desktop_main.__dict__)
        ne_expedicao_actions.configure(desktop_main.__dict__)
        orc_actions.configure(desktop_main.__dict__)
        operador_ordens_actions.configure(desktop_main.__dict__)
        plan_actions.configure(desktop_main.__dict__)
        produtos_actions.configure(desktop_main.__dict__)

        self.app_misc_actions = app_misc_actions
        self.billing_pdf_actions = billing_pdf_actions
        self.encomendas_actions = encomendas_actions
        self.desktop_main = desktop_main
        self.materia_actions = materia_actions
        self.ne_expedicao_actions = ne_expedicao_actions
        self.orc_actions = orc_actions
        self.operador_actions = operador_ordens_actions
        self.plan_actions = plan_actions
        self.produtos_actions = produtos_actions
        self.tax_compliance = tax_compliance
        self.base_dir = Path(getattr(desktop_main, "BASE_DIR", Path.cwd()))
        self.data: dict[str, Any] | None = None
        self._base_data_snapshot: dict[str, Any] | None = None
        self._data_loaded_at = 0.0
        self._reload_cache_ttl_sec = 4.0
        self._trial_status_cache: dict[str, Any] | None = None
        self._trial_status_loaded_at = 0.0
        self._trial_status_cache_ttl_sec = 5.0
        self.user: dict[str, Any] | None = None
        self._qt_config_cache: dict[str, Any] | None = None

    @property
    def branding(self) -> dict[str, Any]:
        cfg = dict(self.desktop_main.get_branding_config() or {})
        cfg["logo_path"] = str(self.logo_path or "")
        return cfg

    @property
    def logo_path(self) -> Path | None:
        candidates: list[str] = []
        seen: set[str] = set()

        def add_candidate(value: Any) -> None:
            text = str(value or "").strip()
            if not text:
                return
            key = text.lower()
            if key in seen:
                return
            seen.add(key)
            candidates.append(text)

        try:
            cfg = self.desktop_main.get_branding_config() or {}
            for value in list(cfg.get("logo_candidates", []) or []):
                add_candidate(value)
        except Exception:
            pass
        try:
            add_candidate(self.desktop_main.get_orc_logo_path() or "")
        except Exception:
            pass
        for fallback in ("Logos/image (1).jpg", "Logos/image.jpg", "Logos/logo.png", "logo.jpg", "logo.png", "Logos/logo(1).jpg"):
            add_candidate(fallback)
        for candidate in candidates:
            if not candidate:
                continue
            path = Path(candidate)
            if not path.is_absolute():
                path = self.base_dir / path
            if path.exists():
                return path
        return None

    def branding_settings(self) -> dict[str, Any]:
        cfg = dict(self.desktop_main.get_branding_config() or {})
        emit = dict(cfg.get("guia_emitente", {}) or {})
        serie_id = ""
        validation_code = ""
        try:
            issue_date = self.desktop_main.now_iso()
            serie_id = str(self.desktop_main._exp_default_serie_id("GT", issue_date) or "").strip()
            find_series_fn = getattr(self.desktop_main, "_find_at_series", None)
            if callable(find_series_fn):
                serie_obj = find_series_fn(self.ensure_data(), doc_type="GT", serie_id=serie_id) or {}
                validation_code = str(serie_obj.get("validation_code", "") or "").strip()
        except Exception:
            serie_id = ""
            validation_code = ""
        return {
            "logo_path": str(self.logo_path or (self.base_dir / "Logos" / "image (1).jpg")),
            "primary_color": str(cfg.get("primary_color", "#000040") or "#000040"),
            "empresa_info_rodape": list(cfg.get("empresa_info_rodape", []) or []),
            "guia_emitente": {
                "nome": str(emit.get("nome", "") or "").strip(),
                "nif": str(emit.get("nif", "") or "").strip(),
                "morada": str(emit.get("morada", "") or "").strip(),
                "local_carga": str(emit.get("local_carga", "") or "").strip(),
            },
            "guia_info_extra": list(cfg.get("guia_info_extra", []) or []),
            "guia_serie_id": serie_id,
            "guia_validation_code": validation_code,
        }

    def ensure_branding_logo(self, preferred_logo: str = "") -> dict[str, Any]:
        preferred = str(preferred_logo or "").strip()
        if not preferred:
            preferred = str(self.base_dir / "Logos" / "image (1).jpg")
        if not Path(preferred).exists():
            return self.branding_settings()
        current = self.branding_settings()
        current_logo = str(current.get("logo_path", "") or "").strip()
        if current_logo and Path(current_logo).exists():
            return current
        payload = dict(current)
        payload["logo_path"] = preferred
        return self.save_branding_settings(payload)

    def save_branding_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = dict(self.desktop_main.get_branding_config() or {})
        logo_path = str(payload.get("logo_path", "") or "").strip()
        guia_serie_id = str(payload.get("guia_serie_id", "") or "").strip()
        guia_validation_code = str(payload.get("guia_validation_code", "") or "").strip()
        rodape = payload.get("empresa_info_rodape", cfg.get("empresa_info_rodape", []))
        if isinstance(rodape, str):
            rodape = [line.strip() for line in rodape.replace("\r", "").split("\n") if line.strip()]
        rodape = [str(line).strip() for line in list(rodape or []) if str(line).strip()]
        guia_extra = payload.get("guia_info_extra", cfg.get("guia_info_extra", []))
        if isinstance(guia_extra, str):
            guia_extra = [line.strip() for line in guia_extra.replace("\r", "").split("\n") if line.strip()]
        guia_extra = [str(line).strip() for line in list(guia_extra or []) if str(line).strip()]
        current_emit = dict(cfg.get("guia_emitente", {}) or {})
        emit_payload = dict(payload.get("guia_emitente", {}) or {})
        emit_cfg = {
            "nome": str(emit_payload.get("nome", current_emit.get("nome", "")) or "").strip(),
            "nif": str(emit_payload.get("nif", current_emit.get("nif", "")) or "").strip(),
            "morada": str(emit_payload.get("morada", current_emit.get("morada", "")) or "").strip(),
            "local_carga": str(emit_payload.get("local_carga", current_emit.get("local_carga", "")) or "").strip(),
        }
        if logo_path:
            cfg["logo"] = logo_path
            existing = [str(value or "").strip() for value in list(cfg.get("logo_candidates", []) or []) if str(value or "").strip()]
            merged = [logo_path] + [value for value in existing if value.lower() != logo_path.lower()]
            cfg["logo_candidates"] = merged
        if rodape:
            cfg["empresa_info_rodape"] = rodape
        cfg["guia_emitente"] = emit_cfg
        cfg["guia_info_extra"] = guia_extra
        if str(payload.get("primary_color", "") or "").strip():
            cfg["primary_color"] = str(payload.get("primary_color", "") or "").strip()

        try:
            branding_file = self.base_dir / str(getattr(self.desktop_main, "BRANDING_FILE", "lugest_branding.json"))
            branding_file.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if callable(connect):
                conn = connect()
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS app_config (
                            ckey VARCHAR(80) PRIMARY KEY,
                            cvalue LONGTEXT NULL,
                            updated_at DATETIME NULL
                        )
                        """
                    )
                    cur.execute(
                        """
                        INSERT INTO app_config (ckey, cvalue, updated_at)
                        VALUES (%s, %s, NOW())
                        ON DUPLICATE KEY UPDATE cvalue=VALUES(cvalue), updated_at=VALUES(updated_at)
                        """,
                        ("branding_config", json.dumps(cfg, ensure_ascii=False)),
                    )
                conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

        try:
            if hasattr(self.desktop_main, "_BRANDING_CACHE"):
                self.desktop_main._BRANDING_CACHE = None
            invalidator = getattr(self.app_misc_actions, "_invalidate_branding_cache", None)
            if callable(invalidator):
                invalidator()
        except Exception:
            pass
        if guia_serie_id or guia_validation_code:
            try:
                serie_id = guia_serie_id or str(self.desktop_main._exp_default_serie_id("GT", self.desktop_main.now_iso()) or "").strip()
                ensure_series_fn = getattr(self.desktop_main, "ensure_at_series_record", None)
                if callable(ensure_series_fn):
                    serie_obj = ensure_series_fn(
                        self.ensure_data(),
                        doc_type="GT",
                        serie_id=serie_id,
                        issue_date=self.desktop_main.now_iso(),
                        validation_code_hint=guia_validation_code,
                    )
                    if guia_validation_code:
                        serie_obj["validation_code"] = guia_validation_code
                        serie_obj["status"] = "REGISTADA"
                        serie_obj["updated_at"] = self.desktop_main.now_iso()
                    self._save(force=True)
            except Exception:
                pass
        return self.branding_settings()

    @property
    def window_icon_path(self) -> Path | None:
        path = self.base_dir / "app.ico"
        return path if path.exists() else None

    def _clone_data(self, data: dict[str, Any] | None) -> dict[str, Any] | None:
        if not isinstance(data, dict):
            return None
        try:
            return copy.deepcopy(data)
        except Exception:
            try:
                return json.loads(json.dumps(data, ensure_ascii=False, default=str))
            except Exception:
                return dict(data)

    def _replace_data_cache(self, data: dict[str, Any]) -> dict[str, Any]:
        self.data = data
        self._base_data_snapshot = self._clone_data(data)
        self._data_loaded_at = time.time()
        try:
            self.desktop_main._RUNTIME_DATA_REF = self.data
        except Exception:
            pass
        return data

    def _bucket_signature(self, value: Any) -> str:
        try:
            return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str, separators=(",", ":"))
        except Exception:
            return repr(value)

    def _changed_data_buckets(self, current: dict[str, Any], base: dict[str, Any] | None) -> list[str]:
        if not isinstance(current, dict):
            return []
        if not isinstance(base, dict):
            return [key for key in current.keys() if not str(key or "").startswith("__")]
        changed: list[str] = []
        keys = {
            str(key)
            for key in set(list(current.keys()) + list(base.keys()))
            if not str(key or "").startswith("__")
        }
        for key in sorted(keys):
            if self._bucket_signature(current.get(key)) != self._bucket_signature(base.get(key)):
                changed.append(key)
        return changed

    def _merge_latest_for_save(self) -> tuple[dict[str, Any], list[str]]:
        current = self.ensure_data()
        changed_keys = self._changed_data_buckets(current, self._base_data_snapshot)
        if not changed_keys:
            return current, []
        latest = self.desktop_main.load_data()
        for key in changed_keys:
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

    def _qt_config_path(self) -> Path:
        return self.base_dir / "lugest_qt_config.json"

    def _load_qt_config(self) -> dict[str, Any]:
        if isinstance(self._qt_config_cache, dict):
            return dict(self._qt_config_cache)
        payload: dict[str, Any] = {}
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if callable(connect):
                conn = connect()
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS app_config (
                            ckey VARCHAR(80) PRIMARY KEY,
                            cvalue LONGTEXT NULL,
                            updated_at DATETIME NULL
                        )
                        """
                    )
                    cur.execute("SELECT cvalue FROM app_config WHERE ckey=%s LIMIT 1", ("qt_desktop_config",))
                    row = cur.fetchone()
                if row:
                    raw = row.get("cvalue") if isinstance(row, dict) else row[0]
                    if isinstance(raw, (bytes, bytearray)):
                        raw = raw.decode("utf-8", errors="ignore")
                    parsed = json.loads(str(raw or "{}"))
                    if isinstance(parsed, dict):
                        payload = parsed
        except Exception:
            payload = {}
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
        if not payload:
            try:
                path = self._qt_config_path()
                if path.exists():
                    parsed = json.loads(path.read_text(encoding="utf-8"))
                    if isinstance(parsed, dict):
                        payload = parsed
            except Exception:
                payload = {}
        self._qt_config_cache = dict(payload)
        return dict(payload)

    def _save_qt_config(self, payload: dict[str, Any]) -> dict[str, Any]:
        clean = dict(payload or {})
        self._qt_config_cache = dict(clean)
        try:
            self._qt_config_path().write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if callable(connect):
                conn = connect()
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS app_config (
                            ckey VARCHAR(80) PRIMARY KEY,
                            cvalue LONGTEXT NULL,
                            updated_at DATETIME NULL
                        )
                        """
                    )
                    cur.execute(
                        """
                        INSERT INTO app_config (ckey, cvalue, updated_at)
                        VALUES (%s, %s, NOW())
                        ON DUPLICATE KEY UPDATE cvalue=VALUES(cvalue), updated_at=VALUES(updated_at)
                        """,
                        ("qt_desktop_config", json.dumps(clean, ensure_ascii=False)),
                    )
                conn.commit()
        except Exception:
            pass
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
        return dict(clean)

    def laser_quote_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return merge_laser_quote_settings(dict(cfg.get("laser_quote", {}) or {}))

    def laser_quote_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        merged = merge_laser_quote_settings(dict(payload or {}))
        cfg["laser_quote"] = merged
        self._save_qt_config(cfg)
        return merge_laser_quote_settings(dict(cfg.get("laser_quote", {}) or {}))

    def laser_quote_analyze(self, payload: dict[str, Any]) -> dict[str, Any]:
        return estimate_laser_quote(dict(payload or {}), self.laser_quote_settings())

    def laser_quote_build_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        analysis = self.laser_quote_analyze(payload)
        return {
            "analysis": analysis,
            "line": dict(analysis.get("line_suggestion", {}) or {}),
        }

    def profile_laser_quote_analyze(self, payload: dict[str, Any]) -> dict[str, Any]:
        return estimate_profile_laser_quote(dict(payload or {}), self.laser_quote_settings())

    def profile_laser_quote_build_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        analysis = self.profile_laser_quote_analyze(payload)
        return {
            "analysis": analysis,
            "line": dict(analysis.get("line_suggestion", {}) or {}),
        }

    def _default_operation_cost_settings(self) -> dict[str, Any]:
        def _profile(
            pricing_mode: str,
            driver_label: str,
            *,
            default_units: float = 1.0,
            setup_min: float = 0.0,
            unit_time_min: float = 0.0,
            hour_rate_eur: float = 0.0,
            fixed_unit_eur: float = 0.0,
            min_unit_eur: float = 0.0,
            extra_unit_eur: float = 0.0,
            requires_driver_input: bool = False,
            note: str = "",
        ) -> dict[str, Any]:
            return {
                "pricing_mode": pricing_mode,
                "driver_label": driver_label,
                "default_units": float(default_units),
                "setup_min": float(setup_min),
                "unit_time_min": float(unit_time_min),
                "hour_rate_eur": float(hour_rate_eur),
                "fixed_unit_eur": float(fixed_unit_eur),
                "min_unit_eur": float(min_unit_eur),
                "extra_unit_eur": float(extra_unit_eur),
                "requires_driver_input": bool(requires_driver_input),
                "note": str(note or "").strip(),
            }

        return {
            "active_profile": "Base",
            "profiles": {
                "Base": {
                    "Corte Laser": _profile(
                        "manual",
                        "Programa",
                        note="Usar o motor de corte laser ou detalhe manual quando esta operacao aparece combinada com outras.",
                    ),
                    "Quinagem": _profile("per_feature", "Dobras/peca", setup_min=6.0, unit_time_min=0.35, hour_rate_eur=42.0, requires_driver_input=True),
                    "Roscagem": _profile("per_feature", "Roscas/peca", setup_min=4.0, unit_time_min=0.2, hour_rate_eur=38.0, requires_driver_input=True),
                    "Serralharia": _profile("per_piece", "Operacoes/peca", setup_min=10.0, unit_time_min=4.0, hour_rate_eur=40.0, requires_driver_input=True),
                    "Lacagem": _profile("per_area_m2", "m2/peca", default_units=1.0, setup_min=6.0, fixed_unit_eur=14.0),
                    "Maquinacao": _profile("per_feature", "Operacoes/peca", setup_min=12.0, unit_time_min=3.0, hour_rate_eur=55.0, requires_driver_input=True),
                    "Soldadura": _profile("per_feature", "Pontos cordoes/peca", setup_min=8.0, unit_time_min=1.5, hour_rate_eur=42.0, requires_driver_input=True),
                    "Montagem": _profile("per_piece", "Operacoes/peca", setup_min=5.0, unit_time_min=2.5, hour_rate_eur=30.0),
                    "Embalamento": _profile("per_piece", "Volumes/peca", setup_min=2.0, unit_time_min=0.8, hour_rate_eur=24.0, min_unit_eur=0.25),
                }
            },
        }

    def _merge_operation_cost_settings(self, stored: dict[str, Any] | None = None) -> dict[str, Any]:
        base = self._default_operation_cost_settings()
        raw = dict(stored or {})
        merged_profiles: dict[str, dict[str, Any]] = {}
        stored_profiles = dict(raw.get("profiles", {}) or {})
        active_profile = str(raw.get("active_profile", base.get("active_profile", "Base")) or "Base").strip() or "Base"

        profile_names = set(stored_profiles.keys()) | set(base.get("profiles", {}).keys()) | {active_profile}
        for profile_name in sorted(profile_names):
            base_profile = dict(dict(base.get("profiles", {}) or {}).get(profile_name, {}) or {})
            current_profile = dict(stored_profiles.get(profile_name, {}) or {})
            merged_profile: dict[str, Any] = {}
            operation_names = set(base_profile.keys()) | set(current_profile.keys()) | set(self.desktop_main.OFF_OPERACOES_DISPONIVEIS)
            for operation_name in sorted(operation_names):
                raw_name = self.desktop_main.normalize_operacao_nome(operation_name) or str(operation_name or "").strip()
                if not raw_name:
                    continue
                template = dict(base_profile.get(raw_name, {}) or {})
                current = dict(current_profile.get(raw_name, current_profile.get(operation_name, {})) or {})
                merged_profile[raw_name] = {
                    "pricing_mode": str(current.get("pricing_mode", template.get("pricing_mode", "manual")) or "manual").strip() or "manual",
                    "driver_label": str(current.get("driver_label", template.get("driver_label", "Qtd./peca")) or "Qtd./peca").strip(),
                    "default_units": round(self._parse_float(current.get("default_units", template.get("default_units", 1)), 1), 4),
                    "setup_min": round(self._parse_float(current.get("setup_min", template.get("setup_min", 0)), 0), 4),
                    "unit_time_min": round(self._parse_float(current.get("unit_time_min", template.get("unit_time_min", 0)), 0), 4),
                    "hour_rate_eur": round(self._parse_float(current.get("hour_rate_eur", template.get("hour_rate_eur", 0)), 0), 4),
                    "fixed_unit_eur": round(self._parse_float(current.get("fixed_unit_eur", template.get("fixed_unit_eur", 0)), 0), 4),
                    "min_unit_eur": round(self._parse_float(current.get("min_unit_eur", template.get("min_unit_eur", 0)), 0), 4),
                    "extra_unit_eur": round(self._parse_float(current.get("extra_unit_eur", template.get("extra_unit_eur", 0)), 0), 4),
                    "requires_driver_input": bool(current.get("requires_driver_input", template.get("requires_driver_input", False))),
                    "note": str(current.get("note", template.get("note", "")) or "").strip(),
                }
            merged_profiles[profile_name] = merged_profile

        if active_profile not in merged_profiles:
            active_profile = next(iter(merged_profiles.keys()), "Base")
        return {"active_profile": active_profile, "profiles": merged_profiles}

    def operation_cost_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return self._merge_operation_cost_settings(dict(cfg.get("operation_costing", {}) or {}))

    def operation_cost_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        merged = self._merge_operation_cost_settings(dict(payload or {}))
        cfg["operation_costing"] = merged
        self._save_qt_config(cfg)
        return self._merge_operation_cost_settings(dict(cfg.get("operation_costing", {}) or {}))

    def operation_cost_estimate(self, payload: dict[str, Any]) -> dict[str, Any]:
        row = dict(payload or {})
        settings = self.operation_cost_settings()
        active_profile = str(settings.get("active_profile", "Base") or "Base").strip() or "Base"
        profile_map = dict(dict(settings.get("profiles", {}) or {}).get(active_profile, {}) or {})
        raw_costing_operations = row.get("costing_operations")
        if isinstance(raw_costing_operations, (list, tuple, set)):
            operations = []
            for raw_name in list(raw_costing_operations):
                normalized = str(self.desktop_main.normalize_operacao_nome(raw_name) or raw_name or "").strip()
                if normalized and normalized not in operations:
                    operations.append(normalized)
        else:
            operations = [
                str(op or "").strip()
                for op in list(self.quote_parse_operacoes_lista(row.get("operacao", "")) or [])
                if str(op or "").strip()
            ]
        qty = max(1.0, self._parse_float(row.get("qtd", 0), 1))
        raw_detail = {
            str(self.desktop_main.normalize_operacao_nome(item.get("nome", "")) or item.get("nome", "") or "").strip(): dict(item or {})
            for item in list(row.get("operacoes_detalhe", []) or [])
            if isinstance(item, dict) and str(item.get("nome", "") or "").strip()
        }
        area_m2 = max(0.0, self._parse_float(row.get("area_m2", row.get("net_area_m2", 0)), 0))
        detail_rows: list[dict[str, Any]] = []
        complete = True
        partial = False
        total_time_unit = 0.0
        total_cost_unit = 0.0
        for index, op_name in enumerate(operations, start=1):
            profile = dict(profile_map.get(op_name, {}) or {})
            existing = dict(raw_detail.get(op_name, {}) or {})
            pricing_mode = str(existing.get("pricing_mode", profile.get("pricing_mode", "manual")) or "manual").strip() or "manual"
            if pricing_mode == "per_area_m2" and area_m2 > 0:
                default_units = area_m2
            else:
                default_units = profile.get("default_units", 1)
            driver_units = round(self._parse_float(existing.get("driver_units", default_units), default_units), 4)
            driver_label = str(existing.get("driver_label", profile.get("driver_label", "Qtd./peca")) or "Qtd./peca").strip()
            setup_min = round(self._parse_float(existing.get("setup_min", profile.get("setup_min", 0)), 0), 4)
            unit_time_base_min = round(self._parse_float(existing.get("unit_time_base_min", existing.get("unit_time_min", profile.get("unit_time_min", 0))), 0), 4)
            hour_rate_eur = round(self._parse_float(existing.get("hour_rate_eur", profile.get("hour_rate_eur", 0)), 0), 4)
            fixed_unit_eur = round(self._parse_float(existing.get("fixed_unit_eur", profile.get("fixed_unit_eur", 0)), 0), 4)
            min_unit_eur = round(self._parse_float(existing.get("min_unit_eur", profile.get("min_unit_eur", 0)), 0), 4)
            extra_unit_eur = round(self._parse_float(existing.get("extra_unit_eur", profile.get("extra_unit_eur", 0)), 0), 4)
            requires_driver_input = bool(existing.get("requires_driver_input", profile.get("requires_driver_input", False)))
            driver_units_confirmed = bool(existing.get("driver_units_confirmed", False))
            if not driver_units_confirmed and pricing_mode != "manual":
                driver_units_confirmed = any(existing.get(field_name) not in (None, "") for field_name in ("tempo_unit_min", "custo_unit_eur"))
            note = str(existing.get("note", profile.get("note", "")) or "").strip()
            manual_time = existing.get("tempo_unit_min")
            manual_cost = existing.get("custo_unit_eur")
            manual_values_confirmed = bool(existing.get("manual_values_confirmed", False))
            if not manual_values_confirmed and (manual_time not in (None, "") or manual_cost not in (None, "")):
                manual_values_confirmed = abs(self._parse_float(manual_time, 0)) > 0.000001 or abs(self._parse_float(manual_cost, 0)) > 0.000001
            computed_time_unit: float | None = None
            computed_cost_unit: float | None = None
            resolved = False
            profile_source = "manual"
            if pricing_mode == "manual":
                if manual_values_confirmed and manual_time not in (None, ""):
                    computed_time_unit = round(self._parse_float(manual_time, 0), 4)
                if manual_values_confirmed and manual_cost not in (None, ""):
                    computed_cost_unit = round(self._parse_float(manual_cost, 0), 4)
                resolved = computed_time_unit is not None and computed_cost_unit is not None
            else:
                profile_source = active_profile
                if requires_driver_input and not driver_units_confirmed:
                    resolved = False
                else:
                    computed_time_unit = round((setup_min / qty) + (driver_units * unit_time_base_min), 4)
                    computed_cost_unit = round(max(min_unit_eur, extra_unit_eur + (driver_units * fixed_unit_eur) + ((hour_rate_eur * computed_time_unit) / 60.0)), 4)
                    resolved = True
            if resolved:
                partial = True
                total_time_unit += float(computed_time_unit or 0)
                total_cost_unit += float(computed_cost_unit or 0)
            else:
                complete = False
                if computed_time_unit is not None or computed_cost_unit is not None:
                    partial = True
            detail_rows.append(
                {
                    "seq": index,
                    "nome": op_name,
                    "pricing_mode": pricing_mode,
                    "profile_name": active_profile,
                    "profile_source": profile_source,
                    "driver_label": driver_label,
                    "driver_units": driver_units,
                    "setup_min": setup_min,
                    "unit_time_base_min": unit_time_base_min,
                    "hour_rate_eur": hour_rate_eur,
                    "fixed_unit_eur": fixed_unit_eur,
                    "min_unit_eur": min_unit_eur,
                    "extra_unit_eur": extra_unit_eur,
                    "requires_driver_input": requires_driver_input,
                    "driver_units_confirmed": driver_units_confirmed,
                    "manual_values_confirmed": manual_values_confirmed,
                    "missing_driver_input": bool(requires_driver_input and pricing_mode != "manual" and not driver_units_confirmed),
                    "tempo_unit_min": computed_time_unit,
                    "custo_unit_eur": computed_cost_unit,
                    "resolved": resolved,
                    "tem_detalhe": resolved,
                    "note": note,
                }
            )
        if detail_rows and complete:
            costing_mode = "detailed"
        elif partial:
            costing_mode = "partial_detail"
        elif len(operations) <= 1:
            costing_mode = "single_operation_total"
        else:
            costing_mode = "aggregate_pending"
        return {
            "active_profile": active_profile,
            "operations": detail_rows,
            "summary": {
                "complete": complete and bool(detail_rows),
                "partial": partial and not (complete and bool(detail_rows)),
                "tempo_unit_total_min": round(total_time_unit, 4),
                "custo_unit_total_eur": round(total_cost_unit, 4),
                "tempo_total_min": round(total_time_unit * qty, 2),
                "custo_total_eur": round(total_cost_unit * qty, 2),
                "costing_mode": costing_mode,
                "qtd": round(qty, 2),
            },
        }

    def _user_profiles(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return dict(cfg.get("user_profiles", {}) or {})

    def _save_user_profiles(self, profiles: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        cfg["user_profiles"] = dict(profiles or {})
        self._save_qt_config(cfg)
        return dict(cfg.get("user_profiles", {}) or {})

    def _user_profile(self, username: str) -> dict[str, Any]:
        return dict(self._user_profiles().get(str(username or "").strip().lower(), {}) or {})

    def _authenticate_error_message(self, data: dict[str, Any], username: str, password: str) -> str:
        username_txt = str(username or "").strip()
        if not username_txt:
            return "Indica um utilizador."

        owner_username = str(self.desktop_main.trial_owner_username() or "").strip()
        local_user = self.desktop_main.find_local_user(data, username_txt)
        local_users = [
            row
            for row in list(data.get("users", []) or [])
            if isinstance(row, dict) and str(row.get("username", "") or "").strip()
        ]

        if owner_username and username_txt.lower() == owner_username.lower():
            if not str(password or "").strip():
                return (
                    f"O login '{owner_username}' e reservado ao proprietario/licenciamento. "
                    "Indica a password real do owner ou entra com um utilizador local da aplicacao."
                )
            return (
                f"O login '{owner_username}' e reservado ao proprietario/licenciamento e esta password nao foi aceite. "
                "Nao uses o hash do lugest.env como password; entra com um utilizador local da aplicacao "
                "ou repoe o administrador local."
            )

        if not local_users:
            return (
                "Nao existem utilizadores locais configurados. "
                "Cria ou repoe o administrador inicial antes de entrar."
            )

        if not isinstance(local_user, dict):
            return "Utilizador nao encontrado."

        return "Password incorreta."

    def authenticate(self, username: str, password: str) -> dict[str, Any]:
        data = self.desktop_main.load_data()
        owner_session = self.desktop_main.ensure_trial_login_session(username, password, allow_owner=True)
        if isinstance(owner_session, dict):
            merged = dict(owner_session)
            self.user = merged
            self.data = data
            try:
                self.desktop_main._RUNTIME_DATA_REF = self.data
            except Exception:
                pass
            return self.user
        user = self.desktop_main.authenticate_local_user(data, username, password)
        if user is None:
            raise ValueError(self._authenticate_error_message(data, username, password))
        self.desktop_main.ensure_trial_login_session(username, password, allow_owner=False)
        merged = self.desktop_main.build_authenticated_user_session(user, password)
        profile = self._user_profile(str(merged.get("username", "") or ""))
        if profile and not bool(profile.get("active", True)):
            raise ValueError("Utilizador desativado.")
        for key in ("posto", "posto_trabalho", "work_center"):
            if str(profile.get("posto", "") or "").strip():
                merged[key] = str(profile.get("posto", "") or "").strip()
        merged["active"] = bool(profile.get("active", True))
        merged["menu_permissions"] = dict(profile.get("menu_permissions", {}) or {})
        self.user = merged
        self.data = data
        try:
            self.desktop_main._RUNTIME_DATA_REF = self.data
        except Exception:
            pass
        try:
            self.desktop_main.touch_trial_success(str(merged.get("username", "") or "").strip(), owner=False)
        except Exception:
            pass
        return self.user

    def trial_status(self, *, force: bool = False) -> dict[str, Any]:
        if (
            not force
            and isinstance(self._trial_status_cache, dict)
            and self._trial_status_loaded_at > 0
            and (time.time() - self._trial_status_loaded_at) <= self._trial_status_cache_ttl_sec
        ):
            return dict(self._trial_status_cache)
        payload = dict(self.desktop_main.get_trial_status() or {})
        payload["management_allowed"] = self.is_owner_session()
        self._trial_status_cache = dict(payload)
        self._trial_status_loaded_at = time.time()
        return payload

    def is_owner_session(self) -> bool:
        return bool(dict(self.user or {}).get("owner_session", False))

    def _ensure_trial_management_access(self) -> None:
        if self.is_owner_session():
            return
        raise ValueError("A gestao de trial/licenca exige login pelo utilizador OWNER.")

    def activate_trial_license(self, company_name: str = "", duration_days: int = 60, notes: str = "") -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        return dict(
            self.desktop_main.activate_trial_license(
                company_name=company_name,
                duration_days=duration_days,
                created_by=actor,
                notes=notes,
                reset_start=True,
            )
            or {}
        )

    def extend_trial_license(self, extra_days: int = 30) -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        return dict(self.desktop_main.extend_trial_license(extra_days=extra_days, updated_by=actor) or {})

    def disable_trial_license(self) -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        return dict(self.desktop_main.disable_trial_license(updated_by=actor) or {})

    def _parse_float(self, value: Any, default: float = 0.0) -> float:
        return float(self.desktop_main.parse_float(value, default))

    def _fmt(self, value: Any) -> str:
        return str(self.desktop_main.fmt_num(value))

    def _write_basic_pdf(self, path: str | Path, lines: list[str]) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        safe_lines = [str(line or "") for line in list(lines or [])]
        width = 595
        height = 842
        content_lines = ["BT", "/F1 11 Tf", "50 800 Td", "14 TL"]
        first = True
        for raw in safe_lines[:52]:
            text = raw.encode("latin-1", errors="replace").decode("latin-1")
            text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            if first:
                content_lines.append(f"({text}) Tj")
                first = False
            else:
                content_lines.append(f"T* ({text}) Tj")
        content_lines.append("ET")
        stream = "\n".join(content_lines).encode("latin-1", errors="replace")
        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>".encode("ascii"),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
        ]
        out = bytearray(b"%PDF-1.4\n")
        offsets = [0]
        for index, obj in enumerate(objects, start=1):
            offsets.append(len(out))
            out.extend(f"{index} 0 obj\n".encode("ascii"))
            out.extend(obj)
            out.extend(b"\nendobj\n")
        xref_offset = len(out)
        out.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
        out.extend(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            out.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        out.extend(
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii")
        )
        target.write_bytes(bytes(out))
        return target

    def _parse_dimension_mm(self, value: Any, default: float = 0.0) -> float:
        if isinstance(value, (int, float)):
            try:
                return float(value)
            except Exception:
                return default
        text = str(value or "").strip().replace(" ", "")
        if not text:
            return default
        if "," in text:
            try:
                return float(text.replace(".", "").replace(",", "."))
            except Exception:
                pass
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
            try:
                return float(text.replace(".", ""))
            except Exception:
                pass
        try:
            return float(text)
        except Exception:
            return default

    def _localizacao(self, record: dict[str, Any]) -> str:
        return str(
            record.get("Localizacao")
            or record.get("Localização")
            or ""
            or ""
        ).strip()

    def _next_material_id(self) -> str:
        highest = 0
        for row in self.ensure_data().get("materiais", []):
            try:
                highest = max(highest, int(str(row.get("id", "")).replace("MAT", "")))
            except Exception:
                continue
        return f"MAT{highest + 1:05d}"

    def _file_reference_name(self, raw: Any, fallback: str = "ficheiro") -> str:
        current = str(raw or "").strip()
        if not current:
            return fallback
        resolved = lugest_storage.resolve_file_reference(current, base_dir=self.base_dir)
        if resolved is not None:
            name = str(resolved.name or "").strip()
            if name:
                return name
        try:
            return Path(current).name or fallback
        except Exception:
            return fallback

    def _storage_output_path(self, category: str, filename: str) -> Path:
        return lugest_storage.allocate_storage_output_path(category, filename, base_dir=self.base_dir)

    def _store_shared_file(self, raw: Any, category: str, preferred_name: str = "") -> str:
        return lugest_storage.import_file_to_storage(
            raw,
            category,
            base_dir=self.base_dir,
            preferred_name=preferred_name,
        )

    def _resolve_file_reference(self, raw: Any) -> Path | None:
        return lugest_storage.resolve_file_reference(raw, base_dir=self.base_dir)

    def open_file_reference(self, raw: str) -> Path:
        target = self._resolve_file_reference(raw)
        if target is None:
            raise ValueError("Ficheiro não indicado.")
        target = target.resolve()
        if not target.exists():
            raise ValueError(f"Ficheiro não encontrado: {target}")
        os.startfile(str(target))
        return target

    def _normalize_storage_paths_for_save(self) -> None:
        data = self.ensure_data()

        def normalize_drawings(node: Any) -> None:
            if isinstance(node, dict):
                for key, value in list(node.items()):
                    if key in {"desenho", "desenho_path"}:
                        node[key] = self._store_shared_file(
                            value,
                            "drawings",
                            preferred_name=self._file_reference_name(value, "desenho"),
                        )
                    else:
                        normalize_drawings(value)
            elif isinstance(node, list):
                for item in node:
                    normalize_drawings(item)

        normalize_drawings(data)

        for note in list(data.get("notas_encomenda", []) or []):
            if not isinstance(note, dict):
                continue
            note["fatura_caminho_ultima"] = self._store_shared_file(
                note.get("fatura_caminho_ultima", ""),
                "notas_encomenda/documentos",
                preferred_name=self._file_reference_name(note.get("fatura_caminho_ultima", ""), "fatura"),
            )
            for key in ("documentos", "entregas"):
                for row in list(note.get(key, []) or []):
                    if not isinstance(row, dict):
                        continue
                    row["caminho"] = self._store_shared_file(
                        row.get("caminho", ""),
                        "notas_encomenda/documentos",
                        preferred_name=self._file_reference_name(row.get("caminho", ""), row.get("titulo", "") or "documento"),
                    )

        for record in list(data.get("faturacao", []) or []):
            if not isinstance(record, dict):
                continue
            for invoice in list(record.get("faturas", []) or []):
                if not isinstance(invoice, dict):
                    continue
                invoice_number = str(invoice.get("numero_fatura", "") or invoice.get("id", "") or "fatura").strip() or "fatura"
                invoice["caminho"] = self._store_shared_file(
                    invoice.get("caminho", ""),
                    "billing/invoices",
                    preferred_name=self._file_reference_name(invoice.get("caminho", ""), f"{invoice_number}.pdf"),
                )
                invoice["communication_filename"] = self._store_shared_file(
                    invoice.get("communication_filename", ""),
                    "billing/compliance",
                    preferred_name=self._file_reference_name(invoice.get("communication_filename", ""), f"{invoice_number}_at.xml"),
                )
            for payment in list(record.get("pagamentos", []) or []):
                if not isinstance(payment, dict):
                    continue
                payment_name = str(payment.get("titulo_comprovativo", "") or payment.get("referencia", "") or payment.get("id", "") or "comprovativo").strip()
                payment["caminho_comprovativo"] = self._store_shared_file(
                    payment.get("caminho_comprovativo", ""),
                    "billing/payments",
                    preferred_name=self._file_reference_name(payment.get("caminho_comprovativo", ""), payment_name or "comprovativo"),
                )

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

    def _save(self, force: bool = False, audit: bool = True) -> None:
        self._normalize_storage_paths_for_save()
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
        self.desktop_main.save_data(payload, force=force)
        if isinstance(payload, dict):
            self._replace_data_cache(payload)

    def _sync_ne_from_materia(self) -> None:
        data = self.ensure_data()
        holder = SimpleNamespace(data=data)
        changed = False
        for ne in data.get("notas_encomenda", []):
            if self.ne_expedicao_actions._sync_ne_linhas_with_materia(holder, ne):
                changed = True
        if changed:
            self._save(force=True)

    def _sync_ne_from_products(self) -> None:
        data = self.ensure_data()
        changed = False
        for note in list(data.get("notas_encomenda", []) or []):
            if self._sync_note_lines_with_products(note):
                self._recalculate_note_totals(note)
                changed = True
        if changed:
            self._save(force=True)

    def _sync_note_lines_with_products(self, note: dict[str, Any]) -> bool:
        changed = False
        product_map = {str(row.get("codigo", "") or "").strip(): row for row in list(self.ensure_data().get("produtos", []) or [])}
        for line in list(note.get("linhas", []) or []):
            if self.desktop_main.origem_is_materia(line.get("origem", "Produto")):
                continue
            product = product_map.get(str(line.get("ref", "") or "").strip())
            if not product:
                continue
            new_price = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 6)
            old_price = self._parse_float(line.get("preco", 0), 0)
            if abs(new_price - old_price) > 1e-9:
                line["preco"] = new_price
                qty = self._parse_float(line.get("qtd", 0), 0)
                discount = max(0.0, min(100.0, self._parse_float(line.get("desconto", 0), 0)))
                iva = max(0.0, min(100.0, self._parse_float(line.get("iva", 23), 23)))
                base = (qty * new_price) * (1.0 - (discount / 100.0))
                line["total"] = round(base + (base * iva / 100.0), 4)
                changed = True
            new_desc = str(product.get("descricao", "") or "").strip()
            if new_desc and str(line.get("descricao", "") or "").strip() != new_desc:
                line["descricao"] = new_desc
                changed = True
        return changed

    def _update_materia_preco_from_unit(self, materia_id: str, preco_unit: Any) -> bool:
        material = self.material_by_id(str(materia_id or "").strip())
        if material is None:
            return False
        price_line = self._parse_float(preco_unit, 0)
        old = self._parse_float(material.get("p_compra", 0), 0)
        new_value = old
        formato = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "").strip()
        if formato == "Tubo":
            metros = self._parse_float(material.get("metros", 0), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        elif formato in ("Chapa", "Perfil"):
            peso = self._parse_float(material.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        material["p_compra"] = new_value
        material["atualizado_em"] = self.desktop_main.now_iso()
        return True

    def _update_produto_preco_from_unit(self, produto_codigo: str, preco_unit: Any) -> bool:
        code = str(produto_codigo or "").strip()
        product = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if product is None:
            return False
        price_line = self._parse_float(preco_unit, 0)
        old = self._parse_float(product.get("p_compra", 0), 0)
        new_value = old
        modo = self.desktop_main.produto_modo_preco(product.get("categoria", ""), product.get("tipo", ""))
        if modo == "peso":
            peso = self._parse_float(product.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        elif modo == "metros":
            metros = self._parse_float(product.get("metros_unidade", product.get("metros", 0)), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        product["p_compra"] = new_value
        product["atualizado_em"] = self.desktop_main.now_iso()
        return True

    def _resolve_supplier(self, raw_value: str) -> tuple[str, str, str]:
        raw = str(raw_value or "").strip()
        if not raw:
            return "", "", ""
        supplier_id = ""
        supplier_name = raw
        if " - " in raw:
            supplier_id, supplier_name = [part.strip() for part in raw.split(" - ", 1)]
        supplier = None
        if supplier_id:
            supplier = next((row for row in list(self.ensure_data().get("fornecedores", []) or []) if str(row.get("id", "") or "").strip() == supplier_id), None)
        if supplier is None and supplier_name:
            supplier = next((row for row in list(self.ensure_data().get("fornecedores", []) or []) if str(row.get("nome", "") or "").strip().lower() == supplier_name.lower()), None)
        if supplier is None:
            return supplier_id, raw, ""
        supplier_id = str(supplier.get("id", "") or "").strip()
        supplier_name = str(supplier.get("nome", "") or "").strip()
        return supplier_id, supplier_name, str(supplier.get("contacto", "") or "").strip()

    def _normalize_supplier_reference(self, supplier_id_value: Any, supplier_text_value: Any) -> tuple[str, str, str]:
        supplier_id_raw = str(supplier_id_value or "").strip()
        supplier_text_raw = str(supplier_text_value or "").strip()
        candidates: list[str] = []
        if supplier_id_raw and supplier_text_raw:
            combined = supplier_text_raw if " - " in supplier_text_raw else f"{supplier_id_raw} - {supplier_text_raw}"
            candidates.append(combined.strip())
        if supplier_id_raw:
            candidates.append(supplier_id_raw)
        if supplier_text_raw:
            candidates.append(supplier_text_raw)
        seen: set[str] = set()
        for candidate in candidates:
            key = candidate.lower()
            if not candidate or key in seen:
                continue
            seen.add(key)
            resolved_id, resolved_text, resolved_contact = self._resolve_supplier(candidate)
            if resolved_id:
                return resolved_id, resolved_text, resolved_contact
        return "", supplier_text_raw, ""

    def _recalculate_note_totals(self, note: dict[str, Any]) -> None:
        note["total"] = round(sum(self._parse_float(line.get("total", 0), 0) for line in list(note.get("linhas", []) or [])), 2)

    def _note_kind(self, note: dict[str, Any]) -> str:
        if str(note.get("origem_cotacao", "") or "").strip():
            return "supplier_order"
        suppliers = {
            str(line.get("fornecedor_linha", "") or "").strip().lower()
            for line in list(note.get("linhas", []) or [])
            if str(line.get("fornecedor_linha", "") or "").strip()
        }
        if len(suppliers) > 1:
            return "rfq"
        if len(suppliers) == 1 and not str(note.get("fornecedor", "") or "").strip():
            return "rfq"
        if list(note.get("ne_geradas", []) or []):
            return "rfq"
        return "purchase_note"

    def _ne_document_type(self, payload: dict[str, Any] | None, fallback: str = "") -> str:
        raw = str((payload or {}).get("tipo", "") or fallback or "").strip().upper()
        raw = raw.replace("+", "_").replace("-", "_").replace(" ", "_")
        normalized = "".join(ch for ch in raw if ch.isalnum() or ch == "_").strip("_")
        if normalized in {"GUIA", "FATURA", "GUIA_FATURA", "ENTREGA", "OUTRO", "DOCUMENTO"}:
            return "DOCUMENTO" if normalized == "OUTRO" else normalized
        guia = str((payload or {}).get("guia", "") or "").strip()
        fatura = str((payload or {}).get("fatura", "") or "").strip()
        if guia and fatura:
            return "GUIA_FATURA"
        if fatura:
            return "FATURA"
        if guia:
            return "GUIA"
        return normalized or "DOCUMENTO"

    def _ne_document_type_label(self, doc_type: str) -> str:
        return {
            "ENTREGA": "Entrega",
            "GUIA": "Guia",
            "FATURA": "Fatura",
            "GUIA_FATURA": "Guia + Fatura",
            "DOCUMENTO": "Documento",
        }.get(str(doc_type or "").strip().upper(), "Documento")

    def _ne_document_title(self, payload: dict[str, Any] | None, doc_type: str = "") -> str:
        raw = payload or {}
        explicit = str(raw.get("titulo", "") or "").strip()
        if explicit:
            return explicit
        guia = str(raw.get("guia", "") or "").strip()
        fatura = str(raw.get("fatura", "") or "").strip()
        caminho = str(raw.get("caminho", "") or "").strip()
        data_documento = str(raw.get("data_documento", "") or "").strip()
        data_entrega = str(raw.get("data_entrega", "") or "").strip()
        resolved_type = self._ne_document_type(raw, fallback=doc_type)
        if guia and fatura:
            return f"Guia {guia} / Fatura {fatura}"
        if fatura:
            return f"Fatura {fatura}"
        if guia:
            return f"Guia {guia}"
        if resolved_type == "ENTREGA":
            date_txt = data_entrega or data_documento
            return f"Entrega {date_txt}".strip()
        if caminho:
            try:
                return Path(caminho).name or self._ne_document_type_label(resolved_type)
            except Exception:
                return self._ne_document_type_label(resolved_type)
        return self._ne_document_type_label(resolved_type)

    def _ne_document_signature(self, payload: dict[str, Any]) -> tuple[str, str, str, str, str, str, str]:
        return (
            str(payload.get("data_registo", "") or "").strip()[:19],
            str(payload.get("tipo", "") or "").strip().upper(),
            str(payload.get("guia", "") or "").strip(),
            str(payload.get("fatura", "") or "").strip(),
            str(payload.get("data_entrega", "") or "").strip()[:10],
            str(payload.get("data_documento", "") or "").strip()[:10],
            str(payload.get("obs", "") or "").strip(),
        )

    def _ne_normalize_document(self, payload: dict[str, Any] | None, default_type: str = "DOCUMENTO") -> dict[str, Any]:
        raw = dict(payload or {})
        doc_type = self._ne_document_type(raw, fallback=default_type)
        data_registo = str(raw.get("data_registo", "") or "").strip() or self.desktop_main.now_iso()
        normalized = {
            "data_registo": data_registo,
            "tipo": doc_type,
            "titulo": self._ne_document_title(raw, doc_type=doc_type),
            "caminho": str(raw.get("caminho", "") or "").strip(),
            "guia": str(raw.get("guia", "") or "").strip(),
            "fatura": str(raw.get("fatura", "") or "").strip(),
            "data_entrega": str(raw.get("data_entrega", "") or "").strip()[:10],
            "data_documento": str(raw.get("data_documento", "") or "").strip()[:10],
            "obs": str(raw.get("obs", "") or "").strip(),
        }
        normalized["tipo_label"] = self._ne_document_type_label(doc_type)
        normalized["has_path"] = bool(normalized["caminho"])
        return normalized

    def _ne_document_rows(self, note: dict[str, Any]) -> list[dict[str, Any]]:
        docs: list[dict[str, Any]] = []
        seen: set[tuple[str, str, str, str, str, str, str]] = set()
        for raw_index, raw_doc in enumerate(list(note.get("documentos", []) or [])):
            if not isinstance(raw_doc, dict):
                continue
            doc = self._ne_normalize_document(raw_doc)
            doc["source"] = "documento"
            doc["source_index"] = raw_index
            signature = self._ne_document_signature(doc)
            seen.add(signature)
            docs.append(doc)
        for raw_index, entrega in enumerate(list(note.get("entregas", []) or [])):
            if not isinstance(entrega, dict):
                continue
            doc = self._ne_normalize_document(
                {
                    "data_registo": entrega.get("data_registo"),
                    "tipo": entrega.get("tipo") or "ENTREGA",
                    "titulo": entrega.get("titulo", ""),
                    "caminho": entrega.get("caminho", ""),
                    "guia": entrega.get("guia", ""),
                    "fatura": entrega.get("fatura", ""),
                    "data_entrega": entrega.get("data_entrega", ""),
                    "data_documento": entrega.get("data_documento", ""),
                    "obs": entrega.get("obs", ""),
                },
                default_type="ENTREGA",
            )
            signature = self._ne_document_signature(doc)
            if signature in seen:
                continue
            doc["source"] = "entrega"
            doc["source_index"] = raw_index
            doc["derived"] = True
            docs.append(doc)
        docs.sort(
            key=lambda item: (
                str(item.get("data_registo", "") or ""),
                str(item.get("data_documento", "") or ""),
                str(item.get("data_entrega", "") or ""),
                str(item.get("titulo", "") or ""),
            ),
            reverse=True,
        )
        for index, doc in enumerate(docs):
            doc["index"] = index
        return docs

    def _fmt_eur(self, value: Any) -> str:
        try:
            number = float(value or 0)
        except Exception:
            number = 0.0
        return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")

    def _planning_norm_esp(self, value: Any) -> str:
        txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
        txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
        if not txt:
            return ""
        try:
            number = float(txt)
            if number.is_integer():
                return str(int(number))
            return f"{number:.6f}".rstrip("0").rstrip(".")
        except Exception:
            return txt

    def _order_reserved_sheet(self, numero: str, material: str = "", espessura: str = "") -> str:
        enc = self.get_encomenda_by_numero(numero)
        if not enc or not list(enc.get("reservas", []) or []):
            return "-"
        mat_norm = str(material or "").strip().lower()
        esp_norm = self._planning_norm_esp(espessura)
        for row in list(enc.get("reservas", []) or []):
            if mat_norm and str(row.get("material", "") or "").strip().lower() != mat_norm:
                continue
            if esp_norm and self._planning_norm_esp(row.get("espessura", "")) != esp_norm:
                continue
            return f"{row.get('material', '')} {row.get('espessura', '')} ({row.get('quantidade', 0)})"
        return "-"

    def material_family_options(self) -> list[dict[str, Any]]:
        helper = getattr(self.materia_actions, "_material_family_options", None)
        if callable(helper):
            try:
                return [dict(row or {}) for row in list(helper(include_auto=True) or [])]
            except Exception:
                pass
        return [
            {"key": "", "label": "Auto", "density": 0.0},
            {"key": "steel", "label": "Aço / Ferro", "density": 7.85},
            {"key": "stainless", "label": "Inox", "density": 7.93},
            {"key": "aluminum", "label": "Alumínio", "density": 2.70},
            {"key": "brass", "label": "Latão", "density": 8.50},
            {"key": "copper", "label": "Cobre", "density": 8.96},
        ]

    def material_family_profile(self, material: Any = "", family: Any = "") -> dict[str, Any]:
        helper = getattr(self.materia_actions, "_resolve_material_family", None)
        if callable(helper):
            try:
                profile = dict(helper(material, family) or {})
                return {
                    "key": str(profile.get("key", "") or "").strip(),
                    "label": str(profile.get("label", "") or "").strip(),
                    "density": round(self._parse_float(profile.get("density", 7.85), 7.85), 3),
                    "explicit": bool(profile.get("explicit")),
                }
            except Exception:
                pass
        family_key = str(family or "").strip()
        options = {str(row.get("key", "") or "").strip(): dict(row or {}) for row in self.material_family_options()}
        if family_key and family_key in options:
            row = options[family_key]
            return {
                "key": family_key,
                "label": str(row.get("label", "") or "").strip(),
                "density": round(self._parse_float(row.get("density", 7.85), 7.85), 3),
                "explicit": True,
            }
        return {
            "key": "steel",
            "label": "Aço / Ferro",
            "density": 7.85,
            "explicit": False,
        }

    def material_presets(self) -> dict[str, list[str]]:
        data = self.ensure_data()
        materiais = list(
            dict.fromkeys(
                list(self.desktop_main.MATERIAIS_PRESET)
                + list(data.get("materiais_hist", []))
                + [str(m.get("material", "")).strip() for m in data.get("materiais", []) if str(m.get("material", "")).strip()]
            )
        )
        espessuras = list(
            dict.fromkeys(
                [self._fmt(v) for v in self.desktop_main.ESPESSURAS_PRESET]
                + [str(v).strip() for v in data.get("espessuras_hist", []) if str(v).strip()]
                + [str(m.get("espessura", "")).strip() for m in data.get("materiais", []) if str(m.get("espessura", "")).strip()]
            )
        )
        locais = list(
            dict.fromkeys(
                list(self.desktop_main.LOCALIZACOES_PRESET)
                + [self._localizacao(m) for m in data.get("materiais", []) if self._localizacao(m)]
                + ["RETALHO"]
            )
        )
        return {
            "formatos": list(self.desktop_main.MATERIA_FORMATOS),
            "materiais": materiais,
            "espessuras": espessuras,
            "locais": locais,
        }

    def material_section_options(self, formato: Any = "") -> list[dict[str, Any]]:
        formato_txt = str(formato or "").strip().title()
        if formato_txt == "Tubo":
            return [dict(row or {}) for row in _TUBE_SECTION_OPTIONS]
        if formato_txt == "Perfil":
            return [dict(row or {}) for row in _PROFILE_SECTION_OPTIONS]
        if formato_txt == "Cantoneira":
            return [dict(row or {}) for row in _ANGLE_SECTION_OPTIONS]
        if formato_txt == "Barra":
            return [dict(row or {}) for row in _BAR_SECTION_OPTIONS]
        return []

    def material_profile_size_options(self, secao_tipo: Any = "") -> list[str]:
        lookup_key = _profile_catalog_lookup_key(secao_tipo)
        if not lookup_key:
            return []
        options = [str(value or "").strip() for value in _PROFILE_STANDARD_KG_M.get(lookup_key, {}).keys() if str(value or "").strip()]
        return sorted(options, key=lambda value: int(_profile_size_lookup_key(value) or "0"))

    def _material_section_type(self, formato: str, row: dict[str, Any] | None = None) -> str:
        payload = dict(row or {})
        raw_value = str(payload.get("secao_tipo", payload.get("tipo_secao", "")) or "").strip()
        material_txt = str(payload.get("material", "") or "").strip()
        formato_txt = str(formato or "").strip().title() or "Chapa"
        if formato_txt == "Tubo":
            for option in _TUBE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if token in {"redondo", "round", "tubo redondo"}:
                return "redondo"
            if token in {"quadrado", "square", "tubo quadrado"}:
                return "quadrado"
            if token in {"retangular", "rectangular", "tubo retangular"}:
                return "retangular"
            mat_token = self._norm_material_token(material_txt)
            if any(key in mat_token for key in ("redond", "diam", "ø")):
                return "redondo"
            if "retang" in mat_token:
                return "retangular"
            if "quadrad" in mat_token:
                return "quadrado"
            diametro = self._parse_dimension_mm(payload.get("diametro", 0), 0)
            comp = self._parse_dimension_mm(payload.get("comprimento", 0), 0)
            larg = self._parse_dimension_mm(payload.get("largura", 0), 0)
            if diametro > 0:
                return "redondo"
            if comp > 0 and larg > 0:
                if abs(comp - larg) <= 1e-6:
                    return "quadrado"
                return "retangular"
            return "quadrado"
        if formato_txt == "Perfil":
            for option in _PROFILE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            lookup_key = _profile_catalog_lookup_key(raw_value)
            if lookup_key:
                return lookup_key
            detected_series, _detected_size = _detect_profile_catalog_from_text(raw_value or material_txt)
            if detected_series:
                return detected_series
            normalized = str(raw_value or "").strip().upper()
            if normalized in {"L", "T", "U", "I", "H", "OUTRO"}:
                return normalized
            mat_token = self._norm_material_token(material_txt)
            if re.search(r"\bperfil\s+l\b", mat_token):
                return "L"
            if re.search(r"\bperfil\s+t\b", mat_token):
                return "T"
            if re.search(r"\bperfil\s+u\b", mat_token):
                return "U"
            if re.search(r"\bperfil\s+i\b", mat_token):
                return "I"
            if re.search(r"\bperfil\s+h\b", mat_token):
                return "H"
            return "OUTRO"
        if formato_txt == "Cantoneira":
            for option in _ANGLE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if "desigu" in token:
                return "abas_desiguais"
            if "igual" in token or token in {"l", "cantoneira"}:
                return "abas_iguais"
            mat_token = self._norm_material_token(material_txt)
            if "desigu" in mat_token:
                return "abas_desiguais"
            return "abas_iguais"
        if formato_txt == "Barra":
            for option in _BAR_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if "quadrad" in token:
                return "quadrada"
            if "retang" in token:
                return "retangular"
            if "chata" in token or "barra" in token:
                return "chata"
            mat_token = self._norm_material_token(material_txt)
            if "quadrad" in mat_token:
                return "quadrada"
            if "retang" in mat_token:
                return "retangular"
            return "chata"
        if "nervurado" in self.desktop_main.norm_text(formato_txt):
            return "nervurado"
        return ""

    def _material_section_label(self, formato: str, secao_tipo: Any = "") -> str:
        key = str(secao_tipo or "").strip()
        if not key:
            return "-"
        formato_txt = str(formato or "").strip().title()
        if formato_txt == "Tubo":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _TUBE_SECTION_OPTIONS}
            return labels.get(key, key.title())
        if formato_txt == "Perfil":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _PROFILE_SECTION_OPTIONS}
            return labels.get(key, key)
        if formato_txt == "Cantoneira":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _ANGLE_SECTION_OPTIONS}
            return labels.get(key, key.replace("_", " ").title())
        if formato_txt == "Barra":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _BAR_SECTION_OPTIONS}
            return labels.get(key, key.replace("_", " ").title())
        if "nervurado" in self.desktop_main.norm_text(formato_txt):
            return "Nervurado"
        return key

    def material_geometry_preview(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(payload or {})
        formato_raw = str(row.get("formato") or self.desktop_main.detect_materia_formato(row) or "Chapa").strip() or "Chapa"
        formato_norm = self.desktop_main.norm_text(formato_raw)
        formato = "Varão nervurado" if "nervurado" in formato_norm else formato_raw.title()
        material = str(row.get("material", "") or "").strip()
        material_familia = str(row.get("material_familia", row.get("familia", "")) or "").strip()
        family_profile = self.material_family_profile(material, material_familia)
        density = round(self._parse_float(family_profile.get("density", _STEEL_DENSITY_G_CM3), _STEEL_DENSITY_G_CM3), 3)
        espessura = str(row.get("espessura", "") or "").strip()
        espessura_mm = self._parse_float(espessura, 0)
        comprimento = round(self._parse_dimension_mm(row.get("comprimento", 0), 0), 3)
        largura = round(self._parse_dimension_mm(row.get("largura", 0), 0), 3)
        altura = round(self._parse_dimension_mm(row.get("altura", 0), 0), 3)
        diametro = round(self._parse_dimension_mm(row.get("diametro", 0), 0), 3)
        metros = round(self._parse_float(row.get("metros", 0), 0), 4)
        kg_m_manual = round(self._parse_float(row.get("kg_m", row.get("peso_metro", 0)), 0), 4)
        peso_existente = round(self._parse_float(row.get("peso_unid", 0), 0), 4)
        secao_tipo = self._material_section_type(formato, row)
        secao_label = self._material_section_label(formato, secao_tipo)
        kg_m = 0.0
        peso_unid = 0.0
        area_mm2 = 0.0
        altura_nominal = altura
        lookup_size_key = ""
        base_lookup_kg_m = 0.0
        uses_catalog = False
        auto_weight = True
        dimension_label = "-"
        dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
        dim_b_text = self._fmt(largura) if largura > 0 else "-"
        calc_hint = ""

        if formato == "Chapa":
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                peso_unid = round((comprimento * largura * espessura_mm * density) / 1000000.0, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm" if comprimento > 0 and largura > 0 else "-"
            calc_hint = "Chapa: comprimento x largura x espessura x densidade."
        elif formato == "Tubo":
            if metros <= 0 and comprimento > 0 and largura <= 0 and diametro <= 0:
                metros = round(comprimento / 1000.0, 4)
            if secao_tipo == "redondo":
                if diametro <= 0 and comprimento > 0 and largura <= 0:
                    diametro = comprimento
                inner_d = max(0.0, diametro - (2.0 * espessura_mm))
                if diametro > 0 and espessura_mm > 0:
                    area_mm2 = max(0.0, math.pi * ((diametro ** 2) - (inner_d ** 2)) / 4.0)
                dim_a_text = f"Ø{self._fmt(diametro)}" if diametro > 0 else "-"
                dim_b_text = "-"
                dimension_label = f"Ø{self._fmt(diametro)} x {self._fmt(espessura_mm)} mm" if diametro > 0 and espessura_mm > 0 else "-"
            else:
                if largura <= 0 and altura > 0:
                    largura = altura
                if altura <= 0 and largura > 0:
                    altura = largura
                inner_w = max(0.0, comprimento - (2.0 * espessura_mm))
                inner_h = max(0.0, largura - (2.0 * espessura_mm))
                if comprimento > 0 and largura > 0 and espessura_mm > 0:
                    area_mm2 = max(0.0, (comprimento * largura) - (inner_w * inner_h))
                dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
                dim_b_text = self._fmt(largura) if largura > 0 else "-"
                if comprimento > 0 and largura > 0 and espessura_mm > 0:
                    dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} x {self._fmt(espessura_mm)} mm"
            if area_mm2 > 0:
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            calc_hint = "Tubo: secção metálica x densidade x comprimento da barra."
        elif formato == "Perfil":
            detected_series, detected_size = _detect_profile_catalog_from_text(material)
            if _profile_catalog_lookup_key(secao_tipo):
                uses_catalog = True
                altura_nominal = altura_nominal or self._parse_dimension_mm(row.get("perfil_tamanho", row.get("size", 0)), 0)
                if altura_nominal <= 0 and detected_series == secao_tipo:
                    altura_nominal = self._parse_dimension_mm(detected_size, 0)
                lookup_size_key = _profile_size_lookup_key(altura_nominal)
                if lookup_size_key:
                    base_lookup_kg_m = float(_PROFILE_STANDARD_KG_M.get(secao_tipo, {}).get(lookup_size_key, 0.0) or 0.0)
                if base_lookup_kg_m > 0:
                    kg_m = round(base_lookup_kg_m * (density / _STEEL_DENSITY_G_CM3), 4)
                elif kg_m_manual > 0:
                    kg_m = kg_m_manual
                    uses_catalog = False
            else:
                if detected_series and not row.get("secao_tipo"):
                    secao_tipo = detected_series
                    secao_label = self._material_section_label(formato, secao_tipo)
                if altura_nominal <= 0:
                    altura_nominal = self._parse_dimension_mm(detected_size, 0)
                kg_m = kg_m_manual
                if kg_m <= 0 and peso_existente > 0 and metros > 0:
                    kg_m = round(peso_existente / metros, 4)
            if kg_m > 0 and metros > 0:
                peso_unid = round(kg_m * metros, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(altura_nominal) if altura_nominal > 0 else "-"
            dim_b_text = secao_tipo or "-"
            if secao_tipo and altura_nominal > 0:
                dimension_label = f"{secao_tipo} {self._fmt(altura_nominal)}"
            elif secao_tipo:
                dimension_label = secao_tipo
            elif altura_nominal > 0:
                dimension_label = f"{self._fmt(altura_nominal)} mm"
            if kg_m > 0:
                calc_hint = "Perfil: kg/m da secção x comprimento da barra."
                if uses_catalog:
                    calc_hint = "Perfil: tabela ACAIL ajustada pela densidade da família selecionada x comprimento da barra."
        elif formato == "Cantoneira":
            if comprimento <= 0 and largura > 0:
                comprimento = largura
            if largura <= 0 and comprimento > 0:
                largura = comprimento
            if secao_tipo == "abas_iguais" and comprimento > 0:
                largura = comprimento
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                area_mm2 = max(0.0, espessura_mm * ((comprimento + largura) - espessura_mm))
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
            dim_b_text = self._fmt(largura) if largura > 0 else "-"
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} x {self._fmt(espessura_mm)} mm"
            calc_hint = "Cantoneira: área aproximada t x (a + b - t) x densidade x comprimento."
        elif formato == "Barra":
            if secao_tipo == "quadrada":
                if comprimento <= 0 and largura > 0:
                    comprimento = largura
                if largura <= 0 and comprimento > 0:
                    largura = comprimento
            elif largura <= 0 and espessura_mm > 0:
                largura = espessura_mm
            if espessura_mm <= 0 and largura > 0:
                espessura_mm = largura
            if comprimento > 0 and largura > 0:
                area_mm2 = max(0.0, comprimento * largura)
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
            dim_b_text = self._fmt(largura) if largura > 0 else "-"
            if comprimento > 0 and largura > 0:
                dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm"
            calc_hint = "Barra maciça: lado A x lado B x densidade x comprimento da barra."
        elif formato == "Varão nervurado":
            if diametro <= 0 and espessura_mm > 0:
                diametro = espessura_mm
            if espessura_mm <= 0 and diametro > 0:
                espessura_mm = diametro
            if diametro > 0:
                area_mm2 = max(0.0, math.pi * (diametro ** 2) / 4.0)
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif kg_m_manual > 0:
                kg_m = kg_m_manual
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = f"Ø{self._fmt(diametro)}" if diametro > 0 else "-"
            dim_b_text = "-"
            if diametro > 0:
                dimension_label = f"Ø{self._fmt(diametro)} mm"
            calc_hint = "Varão nervurado: secção circular maciça x densidade x comprimento da barra."
        else:
            peso_unid = peso_existente
            auto_weight = False

        resolved_espessura = self._fmt(espessura_mm) if espessura_mm > 0 else espessura

        return {
            "formato": formato,
            "secao_tipo": secao_tipo,
            "secao_label": secao_label,
            "comprimento": round(comprimento, 3),
            "largura": round(largura, 3),
            "altura": round(altura_nominal if formato == "Perfil" else altura, 3),
            "diametro": round(diametro, 3),
            "espessura": resolved_espessura,
            "espessura_mm": round(espessura_mm, 3),
            "metros": round(metros, 4),
            "kg_m": round(kg_m, 4),
            "peso_unid": round(peso_unid, 4),
            "area_mm2": round(area_mm2, 3),
            "material_familia": material_familia,
            "material_familia_resolved": str(family_profile.get("key", "") or "").strip(),
            "material_familia_label": str(family_profile.get("label", "") or "").strip(),
            "densidade": density,
            "usa_catalogo": uses_catalog,
            "altura_lookup_key": lookup_size_key,
            "kg_m_catalogo": round(base_lookup_kg_m, 4),
            "dimension_label": dimension_label,
            "dim_a_text": dim_a_text,
            "dim_b_text": dim_b_text,
            "peso_auto": auto_weight,
            "calc_hint": calc_hint,
        }

    def material_price_preview(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(payload or {})
        geometry = self.material_geometry_preview(row)
        formato = str(geometry.get("formato", "Chapa") or "Chapa").strip() or "Chapa"
        p_compra = self._parse_float(row.get("p_compra", 0), 0)
        preco_unid = float(
            self.materia_actions._materia_preco_unid_record(
                {
                    "formato": formato,
                    "metros": geometry.get("metros", 0),
                    "peso_unid": geometry.get("peso_unid", 0),
                    "p_compra": p_compra,
                }
            )
            or 0.0
        )
        return {
            "base_label": "EUR/m" if formato == "Tubo" else "EUR/kg",
            "p_compra": round(p_compra, 4),
            "preco_unid": round(preco_unid, 4),
            "espessura_required": formato in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"},
            **geometry,
        }

    def material_rows(self, filter_text: str = "", in_stock_only: bool = False) -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for index, material in enumerate(data.get("materiais", [])):
            if bool(material.get("is_sobra")) or str(material.get("Localizacao", material.get("Localização", "")) or "").strip().upper() == "RETALHO":
                self.materia_actions._hydrate_retalho_record(data, material)
            preview = self.material_price_preview(material)
            material["preco_unid"] = float(preview.get("preco_unid", 0.0) or 0.0)
            disponivel = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
            if in_stock_only and disponivel <= 0:
                continue
            formato = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "Chapa").strip()
            quality_blocked = self._material_quality_is_blocked(material)
            quality_status = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip()
            try:
                has_contorno = bool(self._parse_material_contour_points(material.get("contorno_points", material.get("shape_points", []))))
            except Exception:
                has_contorno = bool(material.get("contorno_points") or material.get("shape_points"))
            tipo = "Retalho" if material.get("is_sobra") else "Normal"
            if has_contorno:
                tipo = f"{tipo} contorno"
            secao_txt = str(preview.get("secao_label", "") or "").strip()
            if secao_txt and secao_txt != "-":
                tipo = f"{tipo} / {secao_txt}"
            lote_txt = str(material.get("lote_fornecedor", "") or "").strip()
            origem_lotes = list(material.get("origem_lotes_baixa", []) or [])
            if bool(material.get("is_sobra")) and origem_lotes:
                lote_txt = " + ".join(str(item or "").strip() for item in origem_lotes if str(item or "").strip()) or lote_txt
            elif bool(material.get("is_sobra")) and not lote_txt:
                lote_txt = str(material.get("origem_lote", "") or "").strip()
            espessura_raw = str(material.get("espessura", "") or "").strip()
            values = {
                "lote": lote_txt,
                "material": str(material.get("material", "")).strip(),
                "comprimento": str(preview.get("dim_a_text", self._fmt(material.get("comprimento", 0))) or "-"),
                "largura": str(preview.get("dim_b_text", self._fmt(material.get("largura", 0))) or "-"),
                "espessura": self._fmt(espessura_raw) if espessura_raw else "-",
                "quantidade": self._fmt(material.get("quantidade", 0)),
                "reservado": self._fmt(material.get("reservado", 0)),
                "formato": formato,
                "metros": self._fmt(material.get("metros", 0)),
                "peso_unid": self._fmt(preview.get("peso_unid", material.get("peso_unid", 0))),
                "p_compra": self._fmt(material.get("p_compra", 0)),
                "preco_unid": self._fmt(preview.get("preco_unid", material.get("preco_unid", 0))),
                "disponivel": self._fmt(disponivel),
                "tipo": f"{formato} / {tipo}",
                "local": self._localizacao(material),
                "id": str(material.get("id", "")).strip(),
            }
            if query and not any(query in str(value).lower() for value in values.values()):
                continue
            severity = "ok"
            if quality_blocked:
                severity = "critical"
            elif self._parse_float(material.get("quantidade", 0), 0) == 1:
                severity = "one"
            elif disponivel <= float(self.desktop_main.STOCK_VERMELHO):
                severity = "critical"
            elif disponivel <= float(self.desktop_main.STOCK_AMARELO):
                severity = "warning"
            rows.append(
                {
                    "row": values,
                    "severity": severity,
                    "band": "even" if index % 2 == 0 else "odd",
                    "record": material,
                }
            )
        return rows

    def material_by_id(self, material_id: str) -> dict[str, Any] | None:
        material_id = str(material_id or "").strip()
        return next((m for m in self.ensure_data().get("materiais", []) if str(m.get("id", "")).strip() == material_id), None)

    def _material_contour_bbox(self, points: list[list[float]] | list[tuple[float, float]]) -> dict[str, float]:
        raw = [(float(point[0]), float(point[1])) for point in list(points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
        if not raw:
            return {"min_x": 0.0, "min_y": 0.0, "max_x": 0.0, "max_y": 0.0, "width": 0.0, "height": 0.0}
        xs = [point[0] for point in raw]
        ys = [point[1] for point in raw]
        min_x = min(xs)
        min_y = min(ys)
        max_x = max(xs)
        max_y = max(ys)
        return {
            "min_x": round(min_x, 3),
            "min_y": round(min_y, 3),
            "max_x": round(max_x, 3),
            "max_y": round(max_y, 3),
            "width": round(max_x - min_x, 3),
            "height": round(max_y - min_y, 3),
        }

    def _parse_material_contour_points(self, value: Any) -> list[list[float]]:
        if isinstance(value, dict):
            value = value.get("points", value.get("outer", value.get("outer_polygon", [])))
        raw_points: Any = value
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return []
            try:
                parsed = json.loads(text)
            except Exception:
                parsed = None
            if isinstance(parsed, dict):
                raw_points = parsed.get("points", parsed.get("outer", parsed.get("outer_polygon", [])))
            elif parsed is not None:
                raw_points = parsed
            else:
                raw_points = []
                for chunk in text.replace("|", ";").split(";"):
                    piece = str(chunk or "").strip()
                    if not piece:
                        continue
                    xy = [part.strip() for part in piece.split(",")]
                    if len(xy) != 2:
                        raise ValueError("Contorno invalido. Usa o formato x,y; x,y; x,y.")
                    try:
                        raw_points.append([float(xy[0].replace(",", ".")), float(xy[1].replace(",", "."))])
                    except Exception as exc:
                        raise ValueError("Contorno invalido. Usa apenas coordenadas numericas.") from exc
        if not isinstance(raw_points, (list, tuple)):
            raise ValueError("Contorno invalido. Usa uma lista de pontos ou texto x,y; x,y.")
        points: list[list[float]] = []
        for point in list(raw_points or []):
            if not isinstance(point, (list, tuple)) or len(point) < 2:
                raise ValueError("Contorno invalido. Cada ponto deve ter X e Y.")
            x = round(float(point[0]), 3)
            y = round(float(point[1]), 3)
            candidate = [x, y]
            if points and points[-1] == candidate:
                continue
            points.append(candidate)
        if len(points) >= 2 and points[0] == points[-1]:
            points = points[:-1]
        if not points:
            return []
        if len(points) < 3:
            raise ValueError("Contorno invalido. Define pelo menos 3 pontos.")
        bbox = self._material_contour_bbox(points)
        return [
            [round(point[0] - bbox["min_x"], 3), round(point[1] - bbox["min_y"], 3)]
            for point in points
        ]

    def format_material_contour_points(self, value: Any) -> str:
        try:
            points = self._parse_material_contour_points(value)
        except Exception:
            return str(value or "").strip()
        return "; ".join(f"{self._fmt(point[0])},{self._fmt(point[1])}" for point in points)

    def _normalise_material_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        formato_raw = str(payload.get("formato", "Chapa") or "Chapa").strip() or "Chapa"
        formato_norm = self.desktop_main.norm_text(formato_raw)
        formato = "Varão nervurado" if "nervurado" in formato_norm else formato_raw.title()
        material = str(payload.get("material", "")).strip()
        material_familia = str(payload.get("material_familia", payload.get("familia", "")) or "").strip()
        espessura = str(payload.get("espessura", "")).strip()
        comprimento = self._parse_dimension_mm(payload.get("comprimento", 0), 0)
        largura = self._parse_dimension_mm(payload.get("largura", 0), 0)
        altura = self._parse_dimension_mm(payload.get("altura", 0), 0)
        diametro = self._parse_dimension_mm(payload.get("diametro", 0), 0)
        metros = self._parse_float(payload.get("metros", 0), 0)
        quantidade = self._parse_float(payload.get("quantidade", 0), 0)
        reservado = self._parse_float(payload.get("reservado", 0), 0)
        peso_unid = self._parse_float(payload.get("peso_unid", 0), 0)
        kg_m = self._parse_float(payload.get("kg_m", payload.get("peso_metro", 0)), 0)
        p_compra = self._parse_float(payload.get("p_compra", 0), 0)
        local = str(payload.get("local", "")).strip()
        lote = str(payload.get("lote_fornecedor", "")).strip()
        secao_tipo = str(payload.get("secao_tipo", payload.get("tipo_secao", "")) or "").strip()
        contorno_points = self._parse_material_contour_points(payload.get("contorno_points", payload.get("shape_points", [])))
        if contorno_points:
            contour_bbox = self._material_contour_bbox(contorno_points)
            comprimento = max(comprimento, contour_bbox["height"])
            largura = max(largura, contour_bbox["width"])
        if material_familia:
            material_familia = str(self.material_family_profile(material, material_familia).get("key", "") or "").strip()
        else:
            material_familia = ""
        geometry = self.material_geometry_preview(
            {
                "formato": formato,
                "material": material,
                "material_familia": material_familia,
                "espessura": espessura,
                "comprimento": comprimento,
                "largura": largura,
                "altura": altura,
                "diametro": diametro,
                "metros": metros,
                "peso_unid": peso_unid,
                "kg_m": kg_m,
                "secao_tipo": secao_tipo,
            }
        )
        comprimento = float(geometry.get("comprimento", comprimento) or 0)
        largura = float(geometry.get("largura", largura) or 0)
        altura = float(geometry.get("altura", altura) or 0)
        diametro = float(geometry.get("diametro", diametro) or 0)
        espessura = str(geometry.get("espessura", espessura) or "").strip()
        metros = float(geometry.get("metros", metros) or 0)
        peso_unid = float(geometry.get("peso_unid", peso_unid) or 0)
        kg_m = float(geometry.get("kg_m", kg_m) or 0)
        secao_tipo = str(geometry.get("secao_tipo", secao_tipo) or "").strip()
        if not material or quantidade <= 0:
            raise ValueError("Material e quantidade sao obrigatorios.")
        if formato in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"} and not espessura:
            raise ValueError("Para chapa, tubo, cantoneira e varão nervurado, espessura/diâmetro e obrigatoria.")
        if reservado < 0 or reservado > quantidade:
            raise ValueError("Reserva invalida.")
        if formato == "Chapa" and (comprimento <= 0 or largura <= 0):
            raise ValueError("Para chapa, comprimento e largura sao obrigatorios.")
        if formato == "Tubo":
            if metros <= 0:
                raise ValueError("Para tubo, o comprimento da barra e obrigatorio.")
            if secao_tipo == "redondo" and diametro <= 0:
                raise ValueError("Para tubo redondo, o diametro exterior e obrigatorio.")
            if secao_tipo != "redondo" and (comprimento <= 0 or largura <= 0):
                raise ValueError("Para tubo quadrado/retangular, indica lado A e lado B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do tubo com os dados indicados.")
        if formato == "Perfil":
            if metros <= 0:
                raise ValueError("Para perfil, o comprimento da barra e obrigatorio.")
            if not secao_tipo:
                raise ValueError("Para perfil, indica o tipo ou serie.")
            if _profile_catalog_lookup_key(secao_tipo):
                if altura <= 0:
                    raise ValueError("Para perfis de tabela, indica a altura/tamanho nominal.")
                if kg_m <= 0:
                    raise ValueError("Nao existe kg/m tabelado para a serie e altura indicadas.")
            elif kg_m <= 0:
                raise ValueError("Para perfis manuais, indica o peso por metro (kg/m).")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do perfil com os dados indicados.")
        if formato == "Cantoneira":
            if metros <= 0:
                raise ValueError("Para cantoneira, o comprimento da barra e obrigatorio.")
            if comprimento <= 0 or largura <= 0:
                raise ValueError("Para cantoneira, indica aba A e aba B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso da cantoneira com os dados indicados.")
        if formato == "Barra":
            if metros <= 0:
                raise ValueError("Para barra, o comprimento da barra e obrigatorio.")
            if comprimento <= 0 or largura <= 0:
                raise ValueError("Para barra, indica lado A e lado B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso da barra com os dados indicados.")
        if formato == "Varão nervurado":
            if metros <= 0:
                raise ValueError("Para varão nervurado, o comprimento da barra e obrigatorio.")
            if diametro <= 0:
                raise ValueError("Para varão nervurado, indica o diâmetro.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do varão nervurado com os dados indicados.")
        return {
            "formato": formato,
            "material": material,
            "espessura": espessura,
            "comprimento": comprimento,
            "largura": largura,
            "altura": altura,
            "diametro": diametro,
            "metros": metros,
            "kg_m": kg_m,
            "quantidade": quantidade,
            "reservado": reservado,
            "peso_unid": peso_unid,
            "p_compra": p_compra,
            "local": local,
            "lote_fornecedor": lote,
            "secao_tipo": secao_tipo,
            "material_familia": material_familia,
            "contorno_points": contorno_points,
        }

    def add_material(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        values = self._normalise_material_payload(payload)
        record = {
            "id": self._next_material_id(),
            "formato": values["formato"],
            "material": values["material"],
            "espessura": values["espessura"],
            "comprimento": values["comprimento"],
            "largura": values["largura"],
            "altura": values["altura"],
            "diametro": values["diametro"],
            "metros": values["metros"],
            "kg_m": values["kg_m"],
            "quantidade": values["quantidade"],
            "reservado": 0.0,
            "Localização": values["local"],
            "Localizacao": values["local"],
            "lote_fornecedor": values["lote_fornecedor"],
            "secao_tipo": values["secao_tipo"],
            "material_familia": values["material_familia"],
            "peso_unid": values["peso_unid"],
            "p_compra": values["p_compra"],
            "contorno_points": [list(point) for point in list(values.get("contorno_points", []) or [])],
            "preco_unid": float(
                self.materia_actions._materia_preco_unid_record(
                    {
                        "formato": values["formato"],
                        "metros": values["metros"],
                        "peso_unid": values["peso_unid"],
                        "p_compra": values["p_compra"],
                    }
                )
            ),
            "is_sobra": False,
            "atualizado_em": self.desktop_main.now_iso(),
        }
        record = self.materia_actions._hydrate_retalho_record(data, record)
        data.setdefault("materiais", []).append(record)
        self.desktop_main.push_unique(data.setdefault("materiais_hist", []), values["material"])
        if values["espessura"]:
            self.desktop_main.push_unique(data.setdefault("espessuras_hist", []), values["espessura"])
        self.desktop_main.log_stock(
            data,
            "ADICIONAR",
            f"{values['material']} {values['espessura']} qtd={values['quantidade']}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def update_material(self, material_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material n?o encontrado.")
        values = self._normalise_material_payload(payload)
        record.update(
            {
                "formato": values["formato"],
                "material": values["material"],
                "espessura": values["espessura"],
                "comprimento": values["comprimento"],
                "largura": values["largura"],
                "altura": values["altura"],
                "diametro": values["diametro"],
                "metros": values["metros"],
                "kg_m": values["kg_m"],
                "quantidade": values["quantidade"],
                "reservado": values["reservado"],
                "Localização": values["local"],
                "Localizacao": values["local"],
                "lote_fornecedor": values["lote_fornecedor"],
                "secao_tipo": values["secao_tipo"],
                "material_familia": values["material_familia"],
                "peso_unid": values["peso_unid"],
                "p_compra": values["p_compra"],
                "contorno_points": [list(point) for point in list(values.get("contorno_points", []) or [])],
                "atualizado_em": self.desktop_main.now_iso(),
            }
        )
        self.materia_actions._hydrate_retalho_record(data, record)
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        self.desktop_main.log_stock(
            data,
            "EDITAR",
            f"{record.get('id')} qtd={record.get('quantidade', 0)} reservado={record.get('reservado', 0)}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def remove_material(self, material_id: str) -> None:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material n?o encontrado.")
        data["materiais"] = [m for m in data.get("materiais", []) if str(m.get("id", "")).strip() != str(material_id).strip()]
        self.desktop_main.log_stock(
            data,
            "REMOVER",
            f"{record.get('id')} qtd={record.get('quantidade', 0)} reservado={record.get('reservado', 0)}",
            operador=self._current_user_label(),
        )
        self._save(force=True)

    def correct_material_stock(self, material_id: str, quantidade: Any, reservado: Any, metros: Any) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material n?o encontrado.")
        qtd = self._parse_float(quantidade, -1)
        res = self._parse_float(reservado, -1)
        met = self._parse_float(metros, 0)
        if qtd < 0 or res < 0 or res > qtd:
            raise ValueError("Valores inv?lidos.")
        record["quantidade"] = qtd
        record["reservado"] = res
        record["metros"] = met
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record["atualizado_em"] = self.desktop_main.now_iso()
        self.desktop_main.log_stock(
            data,
            "CORRIGIR",
            f"{record.get('id')} qtd={qtd} reservado={res}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def consume_material(self, material_id: str, quantidade: Any, retalho: dict[str, Any] | None = None) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material n?o encontrado.")
        if self._material_quality_is_blocked(record):
            raise ValueError(
                f"Material {material_id} bloqueado pela qualidade: "
                f"{str(record.get('quality_status', record.get('inspection_status', '')) or 'em inspeção')}."
            )
        qtd = self._parse_float(quantidade, 0)
        if qtd <= 0 or qtd > self._parse_float(record.get("quantidade", 0), 0):
            raise ValueError("Quantidade invalida.")
        retalho = dict(retalho or {})
        has_retalho = any(str(retalho.get(key, "")).strip() for key in ("comprimento", "largura", "quantidade", "metros"))
        retalho_row = None
        if has_retalho:
            comp = self._parse_float(retalho.get("comprimento", 0), 0)
            larg = self._parse_float(retalho.get("largura", 0), 0)
            q_retalho = self._parse_float(retalho.get("quantidade", 0), 0)
            metros = self._parse_float(retalho.get("metros", 0), 0)
            contorno_points = self._parse_material_contour_points(retalho.get("contorno_points", retalho.get("shape_points", [])))
            if contorno_points:
                contour_bbox = self._material_contour_bbox(contorno_points)
                comp = max(comp, contour_bbox["height"])
                larg = max(larg, contour_bbox["width"])
            if q_retalho <= 0:
                raise ValueError("Quantidade do retalho invalida.")
            retalho_row = {
                "id": self._next_material_id(),
                "formato": record.get("formato", self.desktop_main.detect_materia_formato(record)),
                "material": record.get("material", ""),
                "espessura": record.get("espessura", ""),
                "comprimento": comp,
                "largura": larg,
                "metros": metros,
                "quantidade": q_retalho,
                "reservado": 0.0,
                "Localizacao": self._localizacao(record),
                "lote_fornecedor": record.get("lote_fornecedor", ""),
                "peso_unid": 0.0,
                "p_compra": record.get("p_compra", 0),
                "preco_unid": 0.0,
                "is_sobra": True,
                "contorno_points": [list(point) for point in list(contorno_points or [])],
                "origem_material_id": str(record.get("id", "") or "").strip(),
                "origem_lote": str(record.get("lote_fornecedor", "") or "").strip(),
                "atualizado_em": self.desktop_main.now_iso(),
            }
            self.materia_actions._hydrate_retalho_record(data, retalho_row, template=record)
        record["quantidade"] = self._parse_float(record.get("quantidade", 0), 0) - qtd
        record["atualizado_em"] = self.desktop_main.now_iso()
        self.desktop_main.log_stock(data, "BAIXA", f"{record.get('id')} qtd={qtd}", operador=self._current_user_label())
        if retalho_row is not None:
            data.setdefault("materiais", []).append(retalho_row)
            self.desktop_main.log_stock(
                data,
                "RETALHO",
                f"{record.get('id')} qtd={retalho_row.get('quantidade', 0)}",
                operador=self._current_user_label(),
            )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def material_candidates(self, material: str, espessura: str, *, include_reserved: bool = False) -> list[dict[str, Any]]:
        material_norm = self.encomendas_actions._norm_material(material)
        esp_norm = self.encomendas_actions._norm_espessura(espessura)
        rows: list[dict[str, Any]] = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
            if self._material_quality_is_blocked(stock):
                continue
            if self.encomendas_actions._norm_material(stock.get("material")) != material_norm:
                continue
            if self.encomendas_actions._norm_espessura(stock.get("espessura")) != esp_norm:
                continue
            try:
                contorno_points = [list(point) for point in list(self._parse_material_contour_points(stock.get("contorno_points", stock.get("shape_points", []))) or [])]
            except Exception:
                contorno_points = []
            total_qty = self._parse_float(stock.get("quantidade", 0), 0)
            reserved = self._parse_float(stock.get("reservado", 0), 0)
            disponivel = total_qty if include_reserved else max(0.0, total_qty - reserved)
            if disponivel <= 0:
                continue
            comprimento = round(self._parse_float(stock.get("comprimento", 0), 0), 2)
            largura = round(self._parse_float(stock.get("largura", 0), 0), 2)
            dimensao = "x".join(
                part
                for part in (self._fmt(comprimento), self._fmt(largura))
                if str(part).strip() and str(part).strip() != "0"
            ) or "-"
            rows.append(
                {
                    "material_id": str(stock.get("id", "") or "").strip(),
                    "material": str(stock.get("material", "") or "").strip(),
                    "espessura": str(stock.get("espessura", "") or "").strip(),
                    "dimensao": dimensao,
                    "comprimento": comprimento,
                    "largura": largura,
                    "disponivel": round(disponivel, 2),
                    "quantidade_total": round(total_qty, 2),
                    "reservado": round(reserved, 2),
                    "local": self._localizacao(stock),
                    "lote": str(stock.get("lote_fornecedor", "") or "").strip(),
                    "origem_lote": str(stock.get("origem_lote", "") or "").strip(),
                    "origem_encomenda": str(stock.get("origem_encomenda", "") or "").strip(),
                    "peso_unid": round(self._parse_float(stock.get("peso_unid", 0), 0), 3),
                    "p_compra": round(self._parse_float(stock.get("p_compra", 0), 0), 6),
                    "is_retalho": bool(stock.get("is_sobra")),
                    "contorno_points": contorno_points,
                }
            )
        rows.sort(
            key=lambda row: (
                str(row.get("lote", "") or ""),
                float(row.get("disponivel", 0) or 0),
                str(row.get("material_id", "") or ""),
            ),
            reverse=True,
        )
        return rows

    def material_price_rows(self, formato_filter: str = "") -> list[dict[str, Any]]:
        filtro = str(formato_filter or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for material in list(self.ensure_data().get("materiais", []) or []):
            if not isinstance(material, dict):
                continue
            preview = dict(self.material_price_preview(material) or {})
            formato = str(preview.get("formato", material.get("formato", "")) or "").strip()
            if filtro and formato.lower() != filtro:
                continue
            price_kg = 0.0
            kg_m = float(preview.get("kg_m", material.get("kg_m", 0)) or 0.0)
            base_price = float(material.get("p_compra", 0) or 0.0)
            if formato == "Tubo":
                price_kg = round(base_price / kg_m, 4) if kg_m > 0 else 0.0
            else:
                price_kg = round(base_price, 4)
            rows.append(
                {
                    "id": str(material.get("id", "") or "").strip(),
                    "formato": formato,
                    "material": str(material.get("material", "") or "").strip(),
                    "secao_tipo": str(preview.get("secao_tipo", material.get("secao_tipo", "")) or "").strip(),
                    "dimension_label": str(preview.get("dimension_label", "") or "").strip(),
                    "espessura": str(preview.get("espessura", material.get("espessura", "")) or "").strip(),
                    "kg_m": round(kg_m, 4),
                    "peso_unid": round(float(preview.get("peso_unid", material.get("peso_unid", 0)) or 0.0), 4),
                    "p_compra": round(base_price, 4),
                    "price_kg": price_kg,
                    "preco_unid": round(float(preview.get("preco_unid", material.get("preco_unid", 0)) or 0.0), 4),
                    "base_label": str(preview.get("base_label", "EUR/kg") or "EUR/kg").strip(),
                    "quantidade": round(float(material.get("quantidade", 0) or 0.0), 2),
                }
            )
        rows.sort(key=lambda item: (item.get("formato", ""), item.get("material", ""), item.get("dimension_label", ""), item.get("id", "")))
        return rows

    def material_default_price_kg(self, formato: str = "", material_name: str = "") -> float:
        formato_txt = str(formato or "").strip().lower()
        material_txt = str(material_name or "").strip().lower()
        rows = list(self.material_price_rows(formato_filter=formato) or [])
        exact = [row for row in rows if material_txt and material_txt in str(row.get("material", "") or "").strip().lower()]
        source = exact or rows
        values = [float(row.get("price_kg", 0) or 0.0) for row in source if float(row.get("price_kg", 0) or 0.0) > 0]
        if not values:
            return 0.0
        return round(sum(values) / len(values), 4)

    def material_update_price_kg(self, material_id: str, price_kg: Any) -> dict[str, Any]:
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        price_kg_value = round(self._parse_float(price_kg, 0), 4)
        if price_kg_value <= 0:
            raise ValueError("Preço/kg inválido.")
        preview = dict(self.material_price_preview(record) or {})
        formato = str(preview.get("formato", record.get("formato", "")) or "").strip() or self.desktop_main.detect_materia_formato(record)
        kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
        if str(formato).strip().lower() == "tubo":
            base_value = round(price_kg_value * kg_m, 4) if kg_m > 0 else 0.0
        else:
            base_value = price_kg_value
        if base_value <= 0:
            raise ValueError("Não foi possível calcular o preço base do material.")
        data = self.ensure_data()
        record["p_compra"] = base_value
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record["atualizado_em"] = self.desktop_main.now_iso()
        self.desktop_main.log_stock(
            data,
            "PRECO",
            f"{record.get('id')} base={base_value} price_kg={price_kg_value}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        updated_preview = dict(self.material_price_preview(record) or {})
        return {
            "id": str(record.get("id", "") or "").strip(),
            "p_compra": round(float(record.get("p_compra", 0) or 0.0), 4),
            "price_kg": round(price_kg_value, 4),
            "preco_unid": round(float(updated_preview.get("preco_unid", record.get("preco_unid", 0)) or 0.0), 4),
            "base_label": str(updated_preview.get("base_label", "EUR/kg") or "EUR/kg").strip(),
        }

    def laser_sheet_stock_candidates(self, material: str, espessura: str, *, include_reserved: bool = False) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for index, row in enumerate(self.material_candidates(material, espessura, include_reserved=include_reserved)):
            width_mm = round(self._parse_float(row.get("largura", 0), 0), 3)
            height_mm = round(self._parse_float(row.get("comprimento", 0), 0), 3)
            quantity_available = max(0, int(self._parse_float(row.get("disponivel", 0), 0)))
            if width_mm <= 0.0 or height_mm <= 0.0 or quantity_available <= 0:
                continue
            is_retalho = bool(row.get("is_retalho"))
            material_id = str(row.get("material_id", "") or "").strip()
            source_kind = "retalho" if is_retalho else "stock"
            source_label = str(row.get("dimensao", "") or "").strip() or f"{height_mm:g} x {width_mm:g}"
            lot_label = str(row.get("lote", "") or "").strip() or str(row.get("origem_lote", "") or "").strip()
            if not lot_label:
                lot_label = material_id or f"stock-{index + 1}"
            rows.append(
                {
                    "name": f"{'Retalho' if is_retalho else 'Stock'} {lot_label} | {source_label}",
                    "source_kind": source_kind,
                    "source_label": f"{'Retalho' if is_retalho else 'Stock'} {lot_label}",
                    "material_id": material_id,
                    "lote": lot_label,
                    "local": str(row.get("local", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "espessura": str(row.get("espessura", "") or "").strip(),
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "area_mm2": round(width_mm * height_mm, 2),
                    "quantity_available": quantity_available,
                    "p_compra": round(self._parse_float(row.get("p_compra", 0), 0), 6),
                    "peso_unid": round(self._parse_float(row.get("peso_unid", 0), 0), 3),
                    "is_retalho": is_retalho,
                    "outer_polygons": [[list(point) for point in list(row.get("contorno_points", []) or [])]] if list(row.get("contorno_points", []) or []) else [],
                }
            )
        rows.sort(
            key=lambda row: (
                0 if str(row.get("source_kind", "") or "") == "retalho" else 1,
                float(row.get("area_mm2", 0) or 0),
                str(row.get("lote", "") or ""),
                str(row.get("material_id", "") or ""),
            )
        )
        return rows

    def consume_material_allocations(
        self,
        allocations: list[dict[str, Any]],
        *,
        retalho: dict[str, Any] | None = None,
        source_material_id: str = "",
        reason: str = "",
    ) -> dict[str, Any]:
        data = self.ensure_data()
        cleaned: list[tuple[dict[str, Any], float]] = []
        for row in list(allocations or []):
            material_id = str((row or {}).get("material_id", "") or "").strip()
            qty = self._parse_float((row or {}).get("quantidade", 0), 0)
            if not material_id or qty <= 0:
                continue
            stock = self.material_by_id(material_id)
            if stock is None:
                raise ValueError(f"Material n?o encontrado: {material_id}")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {material_id} bloqueado pela qualidade.")
            if qty > self._parse_float(stock.get("quantidade", 0), 0):
                raise ValueError(f"Quantidade superior ao stock em {material_id}.")
            cleaned.append((stock, qty))
        if not cleaned:
            raise ValueError("Nenhuma quantidade definida para baixa.")

        retalho_payload = dict(retalho or {})
        has_retalho = any(str(retalho_payload.get(key, "")).strip() for key in ("comprimento", "largura", "quantidade", "metros"))
        chosen_source_id = str(source_material_id or "").strip()
        if has_retalho and not chosen_source_id:
            unique_ids = {str(stock.get("id", "") or "").strip() for stock, _qty in cleaned}
            if len(unique_ids) == 1:
                chosen_source_id = next(iter(unique_ids))
            else:
                raise ValueError("Seleciona o lote de origem do retalho.")
        source_stock = self.material_by_id(chosen_source_id) if chosen_source_id else None
        if chosen_source_id and source_stock is None:
            raise ValueError("Lote de origem do retalho n?o encontrado.")
        if source_stock is not None and not any(str(stock.get("id", "") or "").strip() == chosen_source_id for stock, _qty in cleaned):
            raise ValueError("O lote escolhido para o retalho tem de fazer parte da baixa.")

        consumed_total = 0.0
        used_lots: list[str] = []
        for stock, qty in cleaned:
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - qty)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            consumed_total += qty
            lote = str(stock.get("lote_fornecedor", "") or "").strip()
            if lote:
                used_lots.append(lote)
            obs = f"{stock.get('id', '')} qtd={qty}"
            if reason:
                obs = f"{obs} motivo={reason}"
            self.desktop_main.log_stock(data, "BAIXA", obs, operador=self._current_user_label())

        created_retalho = None
        if has_retalho and source_stock is not None:
            comp = self._parse_float(retalho_payload.get("comprimento", 0), 0)
            larg = self._parse_float(retalho_payload.get("largura", 0), 0)
            q_retalho = self._parse_float(retalho_payload.get("quantidade", 0), 0)
            metros = self._parse_float(retalho_payload.get("metros", 0), 0)
            contorno_points = self._parse_material_contour_points(retalho_payload.get("contorno_points", retalho_payload.get("shape_points", [])))
            if contorno_points:
                contour_bbox = self._material_contour_bbox(contorno_points)
                comp = max(comp, contour_bbox["height"])
                larg = max(larg, contour_bbox["width"])
            if q_retalho <= 0:
                raise ValueError("Quantidade do retalho invalida.")
            created_retalho = {
                "id": self._next_material_id(),
                "formato": source_stock.get("formato", self.desktop_main.detect_materia_formato(source_stock)),
                "material": source_stock.get("material", ""),
                "espessura": source_stock.get("espessura", ""),
                "comprimento": comp,
                "largura": larg,
                "metros": metros,
                "quantidade": q_retalho,
                "reservado": 0.0,
                "Localização": "RETALHO",
                "Localizacao": "RETALHO",
                "lote_fornecedor": source_stock.get("lote_fornecedor", ""),
                "peso_unid": 0.0,
                "p_compra": source_stock.get("p_compra", 0),
                "preco_unid": 0.0,
                "is_sobra": True,
                "contorno_points": [list(point) for point in list(contorno_points or [])],
                "origem_material_id": str(source_stock.get("id", "") or "").strip(),
                "origem_lote": str(source_stock.get("lote_fornecedor", "") or "").strip(),
                "origem_lotes_baixa": list(dict.fromkeys(used_lots)),
                "atualizado_em": self.desktop_main.now_iso(),
            }
            self.materia_actions._hydrate_retalho_record(data, created_retalho, template=source_stock)
            data.setdefault("materiais", []).append(created_retalho)
            log_msg = f"{source_stock.get('id', '')} qtd={created_retalho.get('quantidade', 0)}"
            if reason:
                log_msg = f"{log_msg} motivo={reason}"
            self.desktop_main.log_stock(data, "RETALHO", log_msg, operador=self._current_user_label())
        self._sync_ne_from_materia()
        self._save(force=True)
        return {
            "consumed_total": round(consumed_total, 2),
            "retalho_id": str((created_retalho or {}).get("id", "") or "").strip(),
            "used_lots": list(dict.fromkeys(used_lots)),
        }

    def export_materials_csv(self, path: str | Path, filter_text: str = "") -> Path:
        target = Path(path)
        rows = self.material_rows(filter_text)
        with target.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(
                [
                    "Lote",
                    "Material",
                    "Comprimento",
                    "Largura",
                    "Espessura",
                    "Quantidade",
                    "Reserva",
                    "Formato",
                    "Metros (m)",
                    "Peso/Un. (kg)",
                    "Compra (EUR/kg|EUR/m)",
                    "Preco/Unid (EUR)",
                    "Disponivel",
                    "Tipo",
                    "Localizacao",
                    "ID",
                ]
            )
            for row in rows:
                values = row["row"]
                writer.writerow(
                    [
                        values["lote"],
                        values["material"],
                        values["comprimento"],
                        values["largura"],
                        values["espessura"],
                        values["quantidade"],
                        values["reservado"],
                        values["formato"],
                        values["metros"],
                        values["peso_unid"],
                        values["p_compra"],
                        values["preco_unid"],
                        values["disponivel"],
                        values["tipo"],
                        values["local"],
                        values["id"],
                    ]
                )
        return target

    def stock_log_rows(self, limit: int = 18) -> list[dict[str, str]]:
        rows = []
        for entry in list(reversed(self.ensure_data().get("stock_log", [])[-limit:])):
            rows.append(
                {
                    "data": str(entry.get("data", "")),
                    "acao": str(entry.get("acao", "")),
                    "operador": str(entry.get("operador", "") or "").strip(),
                    "detalhes": str(entry.get("detalhes", "")),
                }
            )
        return rows

    def _material_history_entry_row(self, entry: dict[str, Any], record: dict[str, Any] | None = None) -> dict[str, str]:
        details = str(entry.get("detalhes", "") or "").strip()
        material_id = str((record or {}).get("id", "") or "").strip()
        material_name = str((record or {}).get("material", "") or "").strip()
        espessura = str((record or {}).get("espessura", "") or "").strip()
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        local = self._localizacao(record or {}) if record else ""
        dimensao = ""
        quantity = ""
        reserved = ""

        qty_match = re.search(r"\bqtd=([-+]?\d+(?:[.,]\d+)?)", details, flags=re.IGNORECASE)
        if qty_match:
            quantity = qty_match.group(1)
        reserved_match = re.search(r"\breservado=([-+]?\d+(?:[.,]\d+)?)", details, flags=re.IGNORECASE)
        if reserved_match:
            reserved = reserved_match.group(1)
        if not material_name:
            prefix = re.split(r"\bqtd=|\breservado=", details, maxsplit=1, flags=re.IGNORECASE)[0].strip(" |,-")
            if " Lote:" in prefix:
                base_txt, lote_txt = prefix.split(" Lote:", 1)
                if not lote:
                    lote = lote_txt.strip()
                prefix = base_txt.strip()
            dim_match = re.search(r"(\d+(?:[.,]\d+)?x\d+(?:[.,]\d+)?)", prefix)
            if dim_match:
                dimensao = dim_match.group(1)
                prefix = prefix.replace(dimensao, " ").strip()
            esp_match = re.search(r"(\d+(?:[.,]\d+)?)\s*$", prefix)
            if esp_match:
                if not espessura:
                    espessura = esp_match.group(1)
                prefix = prefix[: esp_match.start()].strip(" |-")
            material_name = prefix or material_name
        else:
            comp = str((record or {}).get("comprimento", "") or "").strip()
            larg = str((record or {}).get("largura", "") or "").strip()
            if comp and larg:
                dimensao = f"{comp}x{larg}"

        return {
            "data": str(entry.get("data", "") or "").replace("T", " ")[:19],
            "acao": str(entry.get("acao", "") or "").strip(),
            "operador": str(entry.get("operador", "") or "").strip(),
            "material_id": material_id,
            "material": material_name,
            "espessura": espessura,
            "dimensao": dimensao,
            "lote": lote,
            "local": local,
            "qtd": quantity,
            "reservado": reserved,
            "detalhes": details,
        }

    def material_history_rows(self, material_id: str = "", limit: int = 240) -> list[dict[str, str]]:
        material_id = str(material_id or "").strip()
        record = self.material_by_id(material_id) if material_id else None
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        material_name = str((record or {}).get("material", "") or "").strip()
        rows: list[dict[str, str]] = []
        for entry in list(reversed(self.ensure_data().get("stock_log", [])[-max(limit, 1) * 4 :])):
            detalhes = str(entry.get("detalhes", "") or "").strip()
            if material_id and material_id not in detalhes and (not lote or lote not in detalhes) and (not material_name or material_name not in detalhes):
                continue
            rows.append(self._material_history_entry_row(entry, record))
            if len(rows) >= limit:
                break
        return rows

    def material_open_stock_pdf(self) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_stock.pdf"
        helper = SimpleNamespace(data=self.ensure_data())
        renderer = getattr(self.materia_actions, "render_stock_a4_pdf", None)
        if callable(renderer):
            renderer(helper, str(target))
            os.startfile(str(target))
            return target
        self.materia_actions.preview_stock_a4(helper)
        return target

    def material_render_stock_pdf(self, path: str | Path) -> Path:
        target = Path(path)
        helper = SimpleNamespace(data=self.ensure_data())
        renderer = getattr(self.materia_actions, "render_stock_a4_pdf", None)
        if callable(renderer):
            renderer(helper, str(target))
            return target
        self.materia_actions.preview_stock_a4(helper)
        return target

    def material_open_history_pdf(self) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_qt_materiais_historico.pdf"
        helper = SimpleNamespace(data=self.ensure_data())
        self.materia_actions.render_stock_log_pdf(helper, str(target))
        os.startfile(str(target))
        return target

    def material_render_history_pdf(self, rows: list[dict[str, Any]], title: str, path: str | Path) -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=landscape(A4))
        page_w, page_h = landscape(A4)
        margin = 18
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        row_h = 18
        header_h = 22
        palette = self._operator_label_palette()
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        columns = [
            ("Data", 98),
            ("Acao", 74),
            ("Operador", 82),
            ("Materia-prima", 164),
            ("Esp.", 44),
            ("Dim.", 78),
            ("Lote", 112),
            ("Qtd", 52),
            ("Reserv.", 56),
            ("Detalhes", page_w - (margin * 2) - 760),
        ]
        source_rows = list(rows or [])
        if not source_rows:
            source_rows = [
                {
                    "data": "-",
                    "acao": "-",
                    "operador": "-",
                    "material": "-",
                    "espessura": "-",
                    "dimensao": "-",
                    "lote": "-",
                    "qtd": "-",
                    "reservado": "-",
                    "detalhes": "Sem registos para imprimir.",
                }
            ]
        rows_per_page = max(1, int((page_h - 112) // row_h))
        total_pages = max(1, math.ceil(len(source_rows) / rows_per_page))

        def draw_page_header(page_no: int) -> float:
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(palette["line_strong"])
            canvas_obj.roundRect(margin, page_h - 74, page_w - (margin * 2), 54, 12, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["ink"])
            title_font = _pdf_fit_font_size(title or "Historico de materia-prima", bold_font, page_w - 260, 17.0, 12.0)
            canvas_obj.setFont(bold_font, title_font)
            canvas_obj.drawString(margin + 14, page_h - 42, self._operator_pdf_text(title or "Historico de materia-prima"))
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 8.6)
            canvas_obj.drawString(margin + 14, page_h - 58, self._operator_pdf_text("Pesquisa consolidada do historico de movimentos para analise e auditoria."))
            canvas_obj.drawRightString(page_w - margin - 14, page_h - 42, self._operator_pdf_text(f"Pagina {page_no}/{total_pages}"))
            canvas_obj.drawRightString(page_w - margin - 14, page_h - 58, self._operator_pdf_text(printed_at))
            table_y = page_h - 96
            canvas_obj.setFillColor(palette["primary"])
            canvas_obj.roundRect(margin, table_y, page_w - (margin * 2), header_h, 8, stroke=0, fill=1)
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setFont(bold_font, 8.0)
            x_cursor = margin
            for label, width in columns:
                canvas_obj.drawString(x_cursor + 6, table_y + 7, self._operator_pdf_text(label))
                x_cursor += width
            return table_y - 4

        for page_index in range(total_pages):
            if page_index:
                canvas_obj.showPage()
            y_cursor = draw_page_header(page_index + 1)
            page_rows = source_rows[page_index * rows_per_page : (page_index + 1) * rows_per_page]
            for row_index, row in enumerate(page_rows):
                draw_y = y_cursor - ((row_index + 1) * row_h)
                fill = palette["surface"] if row_index % 2 == 0 else palette["surface_alt"]
                canvas_obj.setFillColor(fill)
                canvas_obj.setStrokeColor(palette["line"])
                canvas_obj.roundRect(margin, draw_y, page_w - (margin * 2), row_h - 2, 6, stroke=1, fill=1)
                values = [
                    str(row.get("data", "") or "-"),
                    str(row.get("acao", "") or "-"),
                    str(row.get("operador", "") or "-"),
                    str(row.get("material", row.get("material_id", "")) or "-"),
                    str(row.get("espessura", "") or "-"),
                    str(row.get("dimensao", "") or "-"),
                    str(row.get("lote", "") or "-"),
                    str(row.get("qtd", "") or "-"),
                    str(row.get("reservado", "") or "-"),
                    str(row.get("detalhes", "") or "-"),
                ]
                x_cursor = margin
                for col_index, ((_, width), value) in enumerate(zip(columns, values)):
                    align_right = col_index in (7, 8)
                    font_name = bold_font if col_index in (1, 3) else regular_font
                    font_size = 7.2 if col_index != 9 else 7.0
                    clipped = _pdf_clip_text(value, width - 12, font_name, font_size)
                    canvas_obj.setFillColor(palette["ink"] if col_index in (1, 3) else palette["muted"])
                    canvas_obj.setFont(font_name, font_size)
                    if align_right:
                        canvas_obj.drawRightString(x_cursor + width - 6, draw_y + 5.4, self._operator_pdf_text(clipped))
                    else:
                        canvas_obj.drawString(x_cursor + 6, draw_y + 5.4, self._operator_pdf_text(clipped))
                    x_cursor += width
        canvas_obj.save()
        return target

    def _material_is_retalho(self, record: dict[str, Any]) -> bool:
        checker = getattr(self.materia_actions, "_is_retalho_like", None)
        if callable(checker):
            try:
                return bool(checker(record))
            except Exception:
                pass
        return bool((record or {}).get("is_sobra")) or self._localizacao(record).strip().upper() == "RETALHO"

    def _material_label_lot_text(self, record: dict[str, Any]) -> str:
        origem_lotes = [str(item or "").strip() for item in list((record or {}).get("origem_lotes_baixa", []) or []) if str(item or "").strip()]
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        origem_lote = str((record or {}).get("origem_lote", "") or "").strip()
        if origem_lotes:
            return " + ".join(origem_lotes)
        if lote:
            return lote
        if origem_lote:
            return origem_lote
        return "-"

    def _material_label_dimension_text(self, record: dict[str, Any]) -> str:
        preview = self.material_geometry_preview(record)
        dimension_text = str(preview.get("dimension_label", "") or "").strip()
        metros = float(preview.get("metros", 0) or 0)
        if dimension_text and dimension_text != "-":
            if metros > 0:
                return f"{dimension_text} | {self._fmt(metros)} m"
            return dimension_text
        if metros > 0:
            return f"{self._fmt(metros)} m"
        return "-"

    def _draw_code128_fit(
        self,
        canvas_obj,
        value: Any,
        x: float,
        y: float,
        max_width: float,
        bar_height: float,
        min_bar_width: float = 0.38,
        max_bar_width: float = 1.05,
        align: str = "center",
    ) -> float:
        from reportlab.graphics.barcode import code128

        safe_value = str(value or "-").strip() or "-"
        probe = code128.Code128(safe_value, barHeight=bar_height, barWidth=min_bar_width)
        probe_width = float(getattr(probe, "width", 0.0) or 0.0)
        if probe_width > 0:
            unit_width = probe_width / max(min_bar_width, 0.01)
            target_bar_width = max_width / unit_width if unit_width > 0 else min_bar_width
            target_bar_width = max(min_bar_width, min(max_bar_width, target_bar_width))
            barcode = code128.Code128(safe_value, barHeight=bar_height, barWidth=target_bar_width)
        else:
            barcode = probe
        actual_width = float(getattr(barcode, "width", max_width) or max_width)
        draw_x = x
        if align == "center":
            draw_x = x + max(0.0, (max_width - actual_width) / 2.0)
        elif align == "right":
            draw_x = x + max(0.0, max_width - actual_width)
        barcode.drawOn(canvas_obj, draw_x, y)
        return actual_width

    def _draw_material_stock_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        record: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        margin = 16
        outer_x = margin
        outer_y = margin
        outer_w = page_width - (margin * 2)
        outer_h = page_height - (margin * 2)
        header_h = 70
        banner_y = outer_y + outer_h - header_h
        card_w = 112
        card_h = 22
        card_gap = 8
        card_v_gap = 6
        card_group_w = (card_w * 2) + card_gap
        logo_box_w = 88
        logo_box_h = 42
        logo_gap = 14
        logo_x = outer_x + 16
        logo_y = banner_y + 14
        banner_x = logo_x + logo_box_w + logo_gap
        banner_w = outer_w - (banner_x - outer_x)
        group_x = banner_x + banner_w - card_group_w - 18
        title_left = banner_x + 18
        title_right = group_x - 14
        title_w = max(150.0, title_right - title_left)
        material_title = f"{str(record.get('material', '-') or '-').strip()} | {str(record.get('espessura', '-') or '-').strip()} mm"
        dimension_text = self._material_label_dimension_text(record)
        lot_text = self._material_label_lot_text(record)
        formato = str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "Chapa").strip() or "Chapa"
        disponivel = max(0.0, self._parse_float(record.get("quantidade", 0), 0) - self._parse_float(record.get("reservado", 0), 0))
        is_retalho = self._material_is_retalho(record)
        tipo_text = str(record.get("tipo", "") or "").strip() or ("Retalho" if is_retalho else "Chapa / Palete")
        local_text = self._localizacao(record) or "-"
        barcode_value = str(record.get("id", "") or "-").strip() or "-"
        updated_text = str(record.get("atualizado_em", "") or "").replace("T", " ")[:16] or printed_at[:16]

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 14, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(banner_x, banner_y, banner_w, header_h, 14, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            logo_x,
            logo_y,
            logo_box_w,
            logo_box_h,
            radius=12,
            padding_x=6,
            padding_y=5,
            line_width=0.9,
        )

        canvas_obj.setFillColor(palette["ink"])
        title = "Etiqueta de Identificacao"
        subtitle = "Chapa / palete para controlo interno"
        title_font = _pdf_fit_font_size(title, bold_font, title_w, 20.6, 15.2)
        subtitle_font = _pdf_fit_font_size(subtitle, regular_font, title_w, 8.5, 6.7)
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 47, self._operator_pdf_text(title))
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_left + (title_w / 2.0),
            banner_y + 29,
            self._operator_pdf_text(_pdf_clip_text(subtitle, title_w, regular_font, subtitle_font)),
        )

        header_cards = [
            ("ID Stock", barcode_value),
            ("Formato", formato or "-"),
            ("Tipo", tipo_text),
            ("Impresso", printed_at[:16]),
        ]
        for index, (label, value) in enumerate(header_cards):
            row_idx = index // 2
            col_idx = index % 2
            box_x = group_x + (col_idx * (card_w + card_gap))
            box_y = banner_y + header_h - 12 - card_h - (row_idx * (card_h + card_v_gap))
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, box_y, card_w, card_h, 8, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 5.8)
            canvas_obj.drawString(box_x + 8, box_y + card_h - 8, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, card_w - 16, 8.8, 6.0)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(box_x + 8, box_y + 5.4, self._operator_pdf_text(_pdf_clip_text(value, card_w - 16, bold_font, value_font)))

        body_left = outer_x + 14
        body_w = outer_w - 28
        body_top = banner_y - 16
        section_gap = 12
        hero_h = 74
        hero_left_w = 332
        hero_right_w = body_w - hero_left_w - section_gap
        hero_y = body_top - hero_h

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_left, hero_y, hero_left_w, hero_h, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(body_left + 14, hero_y + hero_h - 16, self._operator_pdf_text("Material / espessura"))
        material_font = _pdf_fit_font_size(material_title, bold_font, hero_left_w - 28, 24.0, 16.2)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, material_font)
        canvas_obj.drawString(
            body_left + 14,
            hero_y + hero_h - 39,
            self._operator_pdf_text(_pdf_clip_text(material_title, hero_left_w - 28, bold_font, material_font)),
        )
        material_subtitle = f"Formato {formato} | Tipo {tipo_text}"
        sub_font = _pdf_fit_font_size(material_subtitle, regular_font, hero_left_w - 28, 10.0, 7.4)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, sub_font)
        canvas_obj.drawString(
            body_left + 14,
            hero_y + 16,
            self._operator_pdf_text(_pdf_clip_text(material_subtitle, hero_left_w - 28, regular_font, sub_font)),
        )

        summary_x = body_left + hero_left_w + section_gap
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(summary_x, hero_y, hero_right_w, hero_h, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(summary_x + 14, hero_y + hero_h - 16, self._operator_pdf_text("Stock"))
        id_font = _pdf_fit_font_size(barcode_value, bold_font, hero_right_w - 28, 18.4, 12.6)
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, id_font)
        canvas_obj.drawString(
            summary_x + 14,
            hero_y + hero_h - 38,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, hero_right_w - 28, bold_font, id_font)),
        )
        qty_text = f"Qtd {self._fmt(record.get('quantidade', 0))} | Disp {self._fmt(disponivel)}"
        local_max_w = max(72.0, hero_right_w - 138)
        qty_max_w = max(78.0, hero_right_w - 138)
        qty_font = _pdf_fit_font_size(qty_text, bold_font, qty_max_w, 11.0, 7.8)
        canvas_obj.setFont(regular_font, 7.4)
        canvas_obj.drawString(summary_x + 14, hero_y + 28, self._operator_pdf_text("Local"))
        canvas_obj.drawRightString(summary_x + hero_right_w - 14, hero_y + 28, self._operator_pdf_text("Qtd / Disp"))
        local_font = _pdf_fit_font_size(local_text, bold_font, local_max_w, 11.0, 7.6)
        canvas_obj.setFont(bold_font, local_font)
        canvas_obj.drawString(
            summary_x + 14,
            hero_y + 16,
            self._operator_pdf_text(_pdf_clip_text(local_text, local_max_w, bold_font, local_font)),
        )
        canvas_obj.setFont(bold_font, qty_font)
        canvas_obj.drawRightString(
            summary_x + hero_right_w - 14,
            hero_y + 16,
            self._operator_pdf_text(_pdf_clip_text(qty_text, qty_max_w, bold_font, qty_font)),
        )

        dim_y = hero_y - 56
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_left, dim_y, body_w, 44, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(body_left + 16, dim_y + 28, self._operator_pdf_text("Dimensao identificada"))
        dim_font = _pdf_fit_font_size(dimension_text, bold_font, body_w - 32, 21.5, 13.0)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, dim_font)
        canvas_obj.drawCentredString(
            body_left + (body_w / 2.0),
            dim_y + 12,
            self._operator_pdf_text(_pdf_clip_text(dimension_text, body_w - 32, bold_font, dim_font)),
        )

        info_y = dim_y - 56
        info_gap = 10
        info_w = (body_w - (info_gap * 3)) / 4.0
        info_cards = [
            ("Lote", lot_text),
            ("Peso / un.", f"{self._fmt(record.get('peso_unid', 0))} kg"),
            ("Compra / un.", f"{self._fmt(record.get('preco_unid', 0))} EUR"),
            ("Atualizado", updated_text),
        ]
        for index, (label, value) in enumerate(info_cards):
            box_x = body_left + (index * (info_w + info_gap))
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, info_y, info_w, 42, 10, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 7.0)
            canvas_obj.drawString(box_x + 10, info_y + 27, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, info_w - 20, 10.8, 7.2)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(
                box_x + 10,
                info_y + 12,
                self._operator_pdf_text(_pdf_clip_text(value, info_w - 20, bold_font, value_font)),
            )

        barcode_y = outer_y + 18
        barcode_h = info_y - barcode_y - 14
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_left, barcode_y, body_w, barcode_h, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(body_left + 14, barcode_y + barcode_h - 16, self._operator_pdf_text("Codigo para picagem / identificacao"))
        barcode_area_x = body_left + 18
        barcode_area_w = body_w - 36
        barcode_draw_y = barcode_y + 18
        self._draw_code128_fit(canvas_obj, barcode_value, barcode_area_x, barcode_draw_y, barcode_area_w, 30, min_bar_width=0.52, max_bar_width=1.55)
        canvas_obj.setFillColor(palette["ink"])
        human_font = _pdf_fit_font_size(barcode_value, bold_font, barcode_area_w, 10.8, 8.0)
        canvas_obj.setFont(bold_font, human_font)
        canvas_obj.drawCentredString(
            body_left + (body_w / 2.0),
            barcode_y + 6,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, barcode_area_w, bold_font, human_font)),
        )
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 7.0)
        canvas_obj.drawRightString(outer_x + outer_w - 16, outer_y + 10, self._operator_pdf_text(printed_at))

    def _draw_material_retalho_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        record: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        outer_x = 9
        outer_y = 9
        outer_w = page_width - 18
        outer_h = page_height - 18
        header_h = 36
        banner_y = outer_y + outer_h - header_h
        logo_box_w = 42
        logo_box_h = 20
        logo_gap = 8
        logo_x = outer_x + 8
        logo_y = banner_y + 8
        banner_x = logo_x + logo_box_w + logo_gap
        banner_w = outer_w - (banner_x - outer_x)
        chip_w = 86
        chip_h = 22
        material_title = f"{str(record.get('material', '-') or '-').strip()} {str(record.get('espessura', '-') or '-').strip()} mm"
        dim_text = self._material_label_dimension_text(record)
        lot_text = self._material_label_lot_text(record)
        qty_text = f"Qtd {self._fmt(record.get('quantidade', 0))} | Disp {self._fmt(max(0.0, self._parse_float(record.get('quantidade', 0), 0) - self._parse_float(record.get('reservado', 0), 0)))}"
        barcode_value = str(record.get("id", "") or "-").strip() or "-"
        local_text = self._localizacao(record) or "-"

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(banner_x, banner_y, banner_w, header_h, 10, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            logo_x,
            logo_y,
            logo_box_w,
            logo_box_h,
            radius=7,
            padding_x=4,
            padding_y=3,
            line_width=0.8,
        )

        canvas_obj.setFillColor(palette["ink"])
        title_left = banner_x + 8
        title_right = outer_x + outer_w - chip_w - 14
        title_w = max(52.0, title_right - title_left)
        title_font = _pdf_fit_font_size("Etiqueta Retalho", bold_font, title_w, 12.6, 9.0)
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 20, self._operator_pdf_text("Etiqueta Retalho"))
        subtitle_font = _pdf_fit_font_size(barcode_value, regular_font, title_w, 7.0, 5.7)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_left + (title_w / 2.0),
            banner_y + 9,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, title_w, regular_font, subtitle_font)),
        )

        chip_x = outer_x + outer_w - chip_w - 8
        chip_y = banner_y + 7
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(chip_x, chip_y, chip_w, chip_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.4)
        canvas_obj.drawString(chip_x + 6, chip_y + 13, self._operator_pdf_text("Local"))
        value_font = _pdf_fit_font_size(local_text, bold_font, chip_w - 12, 8.6, 6.0)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, value_font)
        canvas_obj.drawString(chip_x + 6, chip_y + 4.8, self._operator_pdf_text(_pdf_clip_text(local_text, chip_w - 12, bold_font, value_font)))

        body_x = outer_x + 11
        body_w = outer_w - 22
        body_top = banner_y - 11
        ref_font = _pdf_fit_font_size(barcode_value, bold_font, body_w, 14.6, 10.8)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, ref_font)
        canvas_obj.drawString(body_x, body_top - 4, self._operator_pdf_text(_pdf_clip_text(barcode_value, body_w, bold_font, ref_font)))

        material_font = _pdf_fit_font_size(material_title, bold_font, body_w, 11.8, 8.8)
        canvas_obj.setFont(bold_font, material_font)
        canvas_obj.drawString(body_x, body_top - 22, self._operator_pdf_text(_pdf_clip_text(material_title, body_w, bold_font, material_font)))

        dim_y = body_top - 60
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, dim_y, body_w, 28, 10, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 6.4)
        canvas_obj.drawString(body_x + 10, dim_y + 18, self._operator_pdf_text("Dimensao"))
        dim_font = _pdf_fit_font_size(dim_text, bold_font, body_w - 20, 13.4, 9.2)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, dim_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            dim_y + 7,
            self._operator_pdf_text(_pdf_clip_text(dim_text, body_w - 20, bold_font, dim_font)),
        )

        info_y = dim_y - 30
        info_gap = 8
        info_w = (body_w - info_gap) / 2.0
        info_cards = [
            ("Lote", lot_text),
            ("Quantidade", qty_text),
        ]
        for index, (label, value) in enumerate(info_cards):
            box_x = body_x + (index * (info_w + info_gap))
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, info_y, info_w, 23, 8, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 5.4)
            canvas_obj.drawString(box_x + 8, info_y + 13, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, info_w - 16, 7.4, 5.5)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(
                box_x + 8,
                info_y + 5,
                self._operator_pdf_text(_pdf_clip_text(value, info_w - 16, bold_font, value_font)),
            )

        barcode_y = outer_y + 10
        barcode_h = info_y - barcode_y - 10
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, barcode_y, body_w, barcode_h, 10, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.8)
        canvas_obj.drawString(body_x + 9, barcode_y + barcode_h - 11.5, self._operator_pdf_text("Codigo para picagem"))
        barcode_area_x = body_x + 12
        barcode_area_w = body_w - 24
        self._draw_code128_fit(canvas_obj, barcode_value, barcode_area_x, barcode_y + 12.5, barcode_area_w, 18.5, min_bar_width=0.5, max_bar_width=1.2)
        canvas_obj.setFillColor(palette["ink"])
        human_font = _pdf_fit_font_size(barcode_value, bold_font, barcode_area_w, 8.2, 6.2)
        canvas_obj.setFont(bold_font, human_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            barcode_y + 3.8,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, barcode_area_w, bold_font, human_font)),
        )
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.0)
        canvas_obj.drawRightString(outer_x + outer_w - 9, outer_y + 4.8, self._operator_pdf_text(printed_at[:16]))

    def material_identification_label_pdf(self, material_id: str, output_path: str | Path | None = None) -> Path:
        from reportlab.lib.pagesizes import A5, landscape
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        self.materia_actions._hydrate_retalho_record(self.ensure_data(), record)
        is_retalho = self._material_is_retalho(record)
        target = Path(output_path) if output_path else self._operator_label_tmp_path(material_id, "material_identification")
        target.parent.mkdir(parents=True, exist_ok=True)
        page_size = ((100 * mm), (70 * mm)) if is_retalho else landscape(A5)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
        if is_retalho:
            self._draw_material_retalho_label(canvas_obj, page_size[0], page_size[1], record, palette, logo_path, printed_at)
        else:
            self._draw_material_stock_label(canvas_obj, page_size[0], page_size[1], record, palette, logo_path, printed_at)
        canvas_obj.save()
        return target

    def product_presets(self) -> dict[str, list[str]]:
        return {
            "categorias": [str(v) for v in list(getattr(self.desktop_main, "PROD_CATEGORIAS", []) or [])],
            "subcats": [str(v) for v in list(getattr(self.desktop_main, "PROD_SUBCATS", []) or [])],
            "tipos": [str(v) for v in list(getattr(self.desktop_main, "PROD_TIPOS", []) or [])],
            "unidades": [str(v) for v in list(getattr(self.desktop_main, "PROD_UNIDS", []) or [])],
        }

    def product_next_code(self) -> str:
        return str(self.desktop_main.peek_next_produto_numero(self.ensure_data()))

    def _product_dimensoes(self, prod: dict[str, Any]) -> str:
        dim = str(prod.get("dimensoes", "") or "").strip()
        if dim:
            return dim
        comp = self._parse_float(prod.get("comprimento", 0), 0)
        larg = self._parse_float(prod.get("largura", 0), 0)
        esp = self._parse_float(prod.get("espessura", 0), 0)
        if comp > 0 or larg > 0 or esp > 0:
            return f"{self._fmt(comp)}x{self._fmt(larg)}x{self._fmt(esp)}"
        return "-"

    def _product_normalize_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        code = str(payload.get("codigo", "") or "").strip() or self.product_next_code()
        descricao = str(payload.get("descricao", "") or "").strip()
        if not code:
            raise ValueError("Codigo do produto em falta.")
        if not descricao:
            raise ValueError("Descricao do produto em falta.")
        categoria = str(payload.get("categoria", "") or "").strip()
        tipo = str(payload.get("tipo", "") or "").strip()
        metros_unidade = self._parse_float(payload.get("metros_unidade", payload.get("metros", 0)), 0)
        prod = {
            "codigo": code,
            "descricao": descricao,
            "categoria": categoria,
            "subcat": str(payload.get("subcat", "") or "").strip(),
            "tipo": tipo,
            "dimensoes": str(payload.get("dimensoes", "") or "").strip(),
            "comprimento": self._parse_float(payload.get("comprimento", 0), 0),
            "largura": self._parse_float(payload.get("largura", 0), 0),
            "espessura": self._parse_float(payload.get("espessura", 0), 0),
            "metros_unidade": metros_unidade,
            "metros": metros_unidade,
            "peso_unid": self._parse_float(payload.get("peso_unid", 0), 0),
            "fabricante": str(payload.get("fabricante", "") or "").strip(),
            "modelo": str(payload.get("modelo", "") or "").strip(),
            "unid": str(payload.get("unid", "UN") or "UN").strip() or "UN",
            "qty": self._parse_float(payload.get("qty", payload.get("quantidade", 0)), 0),
            "alerta": self._parse_float(payload.get("alerta", 0), 0),
            "p_compra": self._parse_float(payload.get("p_compra", 0), 0),
            "pvp1": self._parse_float(payload.get("pvp1", 0), 0),
            "pvp2": self._parse_float(payload.get("pvp2", 0), 0),
            "obs": str(payload.get("obs", "") or "").strip(),
        }
        if not prod["dimensoes"] and (prod["comprimento"] > 0 or prod["largura"] > 0 or prod["espessura"] > 0):
            prod["dimensoes"] = self._product_dimensoes(prod)
        prod["preco_unid"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
        return prod

    def product_rows(self, filter_text: str = "", in_stock_only: bool = False) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for index, prod in enumerate(list(self.ensure_data().get("produtos", []) or [])):
            price_unit = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
            qty = self._parse_float(prod.get("qty", 0), 0)
            physical_qty = qty + max(0.0, self._parse_float(prod.get("quality_pending_qty", 0), 0))
            if in_stock_only and qty <= 0:
                continue
            quality_status = str(prod.get("quality_status", "") or "").strip()
            quality_display_status = (
                "EM_INSPECAO"
                if self._parse_float(prod.get("quality_pending_qty", 0), 0) > 0
                else quality_status
            )
            quality_blocked = bool(prod.get("quality_blocked")) or (
                bool(quality_display_status) and not self._quality_status_is_available(quality_display_status)
            )
            alerta = self._parse_float(prod.get("alerta", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "descricao": str(prod.get("descricao", "") or "").strip(),
                "categoria": str(prod.get("categoria", "") or "").strip(),
                "subcat": str(prod.get("subcat", "") or "").strip(),
                "tipo": str(prod.get("tipo", "") or "").strip(),
                "dimensoes": self._product_dimensoes(prod),
                "unid": str(prod.get("unid", "UN") or "UN").strip() or "UN",
                "qty": physical_qty,
                "available_qty": qty,
                "alerta": alerta,
                "p_compra": round(self._parse_float(prod.get("p_compra", 0), 0), 4),
                "preco_unid": price_unit,
                "valor_stock": round(physical_qty * price_unit, 2),
                "metros_unidade": round(self._parse_float(prod.get("metros_unidade", prod.get("metros", 0)), 0), 4),
                "peso_unid": round(self._parse_float(prod.get("peso_unid", 0), 0), 4),
                "fabricante": str(prod.get("fabricante", "") or "").strip(),
                "modelo": str(prod.get("modelo", "") or "").strip(),
                "obs": str(prod.get("obs", "") or "").strip(),
                "quality_status": quality_display_status,
                "quality_pending_qty": round(self._parse_float(prod.get("quality_pending_qty", 0), 0), 4),
                "updated_at": str(prod.get("atualizado_em", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            if quality_blocked:
                severity = "warning"
            elif qty <= 0 or (alerta > 0 and qty <= alerta):
                severity = "warning"
            else:
                severity = "ok"
            row["severity"] = severity
            row["band"] = "even" if index % 2 == 0 else "odd"
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("descricao") or ""))
        return rows

    def _product_issue_meta_from_obs(self, obs: str) -> dict[str, Any]:
        text = str(obs or "").strip()
        meta: dict[str, Any] = {}
        if "|meta|" not in text:
            return meta
        for part in text.split("|meta|")[1:]:
            chunk = part.strip()
            if "=" not in chunk:
                continue
            key, value = chunk.split("=", 1)
            meta[key.strip()] = value.strip()
        return meta

    def product_movement_years(self, codigo: str = "") -> list[str]:
        years: set[str] = set()
        for row in self.product_movements(codigo=codigo, limit=5000):
            stamp = str(row.get("data", "") or "")
            if len(stamp) >= 4 and stamp[:4].isdigit():
                years.add(stamp[:4])
        if not years:
            years.add(str(date.today().year))
        return sorted(years, reverse=True)

    def product_movements(
        self,
        codigo: str = "",
        limit: int = 120,
        operator_name: str = "",
        year: str = "",
        issue_only: bool = False,
    ) -> list[dict[str, Any]]:
        code = str(codigo or "").strip()
        operator_name = str(operator_name or "").strip().lower()
        year = str(year or "").strip()
        rows = []
        for row in reversed(list(self.ensure_data().get("produtos_mov", []) or [])):
            mov = dict(row or {})
            mov_code = str(mov.get("codigo", "") or mov.get("produto", "") or "").strip()
            if code and mov_code != code:
                continue
            mov_tipo = str(mov.get("tipo", "") or "").strip()
            if issue_only and mov_tipo != "ENTREGA_OPERADOR":
                continue
            mov_operator = str(mov.get("operador", "") or "").strip()
            if operator_name and mov_operator.lower() != operator_name:
                continue
            mov_data = str(mov.get("data", "") or "").replace("T", " ")[:19]
            if year and year != "Todos" and not mov_data.startswith(year):
                continue
            meta = self._product_issue_meta_from_obs(mov.get("obs", ""))
            rows.append(
                {
                    "data": mov_data,
                    "tipo": mov_tipo,
                    "operador": mov_operator,
                    "codigo": mov_code,
                    "descricao": str(mov.get("descricao", "") or "").strip(),
                    "qtd": round(self._parse_float(mov.get("qtd", 0), 0), 2),
                    "antes": round(self._parse_float(mov.get("antes", 0), 0), 2),
                    "depois": round(self._parse_float(mov.get("depois", 0), 0), 2),
                    "obs": str(mov.get("obs", "") or "").strip(),
                    "origem": str(mov.get("origem", "") or "").strip(),
                    "valor_unit": round(self._parse_float(meta.get("valor_unit", 0), 0), 4),
                    "valor_total": round(self._parse_float(meta.get("valor_total", 0), 0), 2),
                }
            )
            if len(rows) >= limit:
                break
        return rows

    def product_issue_summary(self, operator_name: str = "", year: str = "", codigo: str = "") -> dict[str, Any]:
        rows = self.product_movements(codigo=codigo, limit=5000, operator_name=operator_name, year=year, issue_only=True)
        total_qtd = sum(self._parse_float(row.get("qtd", 0), 0) for row in rows)
        total_valor = sum(self._parse_float(row.get("valor_total", 0), 0) for row in rows)
        return {
            "linhas": len(rows),
            "qtd_total": round(total_qtd, 2),
            "valor_total": round(total_valor, 2),
        }

    def product_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        prod = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if prod is None:
            raise ValueError("Produto n?o encontrado.")
        detail = dict(prod)
        detail["preco_unid"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(detail), 0), 4)
        detail["valor_stock"] = round(self._parse_float(detail.get("qty", 0), 0) * detail["preco_unid"], 2)
        detail["dimensoes"] = self._product_dimensoes(detail)
        detail["movimentos"] = self.product_movements(code, limit=80)
        return detail

    def product_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        prod = self._product_normalize_payload(payload)
        code = str(prod.get("codigo", "") or "").strip()
        rows = data.setdefault("produtos", [])
        existing = next((row for row in rows if str(row.get("codigo", "") or "").strip() == code), None)
        old_qty = self._parse_float((existing or {}).get("qty", 0), 0)
        if existing is None:
            rows.append(prod)
            target = prod
        else:
            existing.update(prod)
            target = existing
        target["atualizado_em"] = self.desktop_main.now_iso()
        new_qty = self._parse_float(target.get("qty", 0), 0)
        operador = str((self.user or {}).get("username", "") or "Sistema")
        if existing is None and new_qty > 1e-9:
            self.desktop_main.add_produto_mov(
                data,
                tipo="ENTRADA_INICIAL",
                operador=operador,
                codigo=code,
                descricao=str(target.get("descricao", "") or "").strip(),
                qtd=new_qty,
                antes=0.0,
                depois=new_qty,
                obs="Stock inicial no registo do produto",
                origem="PRODUTOS",
                ref_doc=code,
            )
        elif existing is not None and abs(new_qty - old_qty) > 1e-9:
            delta = new_qty - old_qty
            self.desktop_main.add_produto_mov(
                data,
                tipo="AJUSTE_STOCK",
                operador=operador,
                codigo=code,
                descricao=str(target.get("descricao", "") or "").strip(),
                qtd=abs(delta),
                antes=old_qty,
                depois=new_qty,
                obs=f"Ajuste manual no cadastro ({self._fmt(delta)})",
                origem="PRODUTOS",
                ref_doc=code,
            )
        self.desktop_main.ensure_produto_seq(data, code)
        self._save(force=True)
        return self.product_detail(code)

    def product_remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        rows = list(self.ensure_data().get("produtos", []) or [])
        before = len(rows)
        self.ensure_data()["produtos"] = [row for row in rows if str(row.get("codigo", "") or "").strip() != code]
        if len(self.ensure_data()["produtos"]) == before:
            raise ValueError("Produto n?o encontrado.")
        self._save(force=True)

    def product_consume(
        self,
        codigo: str,
        quantidade: Any,
        obs: str = "",
        target_operator: str = "",
        issue_mode: str = "stock",
    ) -> dict[str, Any]:
        code = str(codigo or "").strip()
        qty = self._parse_float(quantidade, 0)
        if qty <= 0:
            raise ValueError("Quantidade invalida.")
        prod = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if prod is None:
            raise ValueError("Produto n?o encontrado.")
        before = self._parse_float(prod.get("qty", 0), 0)
        if qty > before + 1e-9:
            raise ValueError("Quantidade superior ao stock disponivel.")
        prod["qty"] = max(0.0, before - qty)
        prod["atualizado_em"] = self.desktop_main.now_iso()
        actor = str((self.user or {}).get("username", "") or "Sistema")
        issue_mode = str(issue_mode or "stock").strip().lower()
        operator_txt = str(target_operator or "").strip()
        movement_type = "BAIXA"
        movement_operator = operator_txt or actor
        detail_obs = str(obs or "").strip() or "Baixa manual no desktop Qt"
        if issue_mode == "operator":
            if not operator_txt:
                raise ValueError("Seleciona o operador que recebe o material.")
            movement_type = "ENTREGA_OPERADOR"
            movement_operator = operator_txt
            valor_unit = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
            valor_total = round(valor_unit * qty, 2)
            note = str(obs or "").strip() or "Entrega a operador"
            detail_obs = (
                f"{note} |meta|actor={actor} |meta|valor_unit={valor_unit:.4f} "
                f"|meta|valor_total={valor_total:.2f}"
            )
        self.desktop_main.add_produto_mov(
            self.ensure_data(),
            tipo=movement_type,
            operador=movement_operator,
            codigo=code,
            descricao=str(prod.get("descricao", "") or "").strip(),
            qtd=qty,
            antes=before,
            depois=self._parse_float(prod.get("qty", 0), 0),
            obs=detail_obs,
            origem="OPERADOR" if movement_type == "ENTREGA_OPERADOR" else "PRODUTOS",
            ref_doc=code,
        )
        self._save(force=True)
        return self.product_detail(code)

    def product_render_stock_pdf(self, path: str | Path) -> Path:
        helper = SimpleNamespace(data=self.ensure_data(), user=self.user or {})
        target = Path(path)
        self.produtos_actions.render_produtos_stock_pdf(helper, str(target))
        return target

    def product_open_stock_pdf(self) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_qt_produtos_stock.pdf"
        self.product_render_stock_pdf(target)
        os.startfile(str(target))
        return target

    def _draw_product_stock_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        product: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        outer_x = 8
        outer_y = 8
        outer_w = page_width - 16
        outer_h = page_height - 16
        header_h = 32
        banner_y = outer_y + outer_h - header_h
        logo_box_w = 46
        logo_box_h = 18
        logo_gap = 7
        logo_x = outer_x + 8
        logo_y = banner_y + 7
        banner_x = logo_x + logo_box_w + logo_gap
        banner_w = outer_w - (banner_x - outer_x)
        chip_w = 80
        chip_h = 18
        chip_x = outer_x + outer_w - chip_w - 8
        chip_y = banner_y + 7
        title_left = banner_x + 10
        title_right = chip_x - 10
        title_w = max(64.0, title_right - title_left)
        code = str(product.get("codigo", "") or "-").strip() or "-"
        desc = str(product.get("descricao", "") or "-").strip() or "-"
        category = str(product.get("categoria", "") or "-").strip() or "-"
        kind = str(product.get("tipo", "") or "-").strip() or "-"
        unit = str(product.get("unid", "") or "UN").strip() or "UN"
        dim_text = self._product_dimensoes(product)
        qty = self._parse_float(product.get("qty", 0), 0)
        price_unit = self._parse_float(product.get("preco_unid", 0), 0)
        updated_text = str(product.get("atualizado_em", "") or "").replace("T", " ")[:16] or printed_at[:16]
        meta_text = f"{category} | {kind} | {dim_text}"

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(banner_x, banner_y, banner_w, header_h, 10, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            logo_x,
            logo_y,
            logo_box_w,
            logo_box_h,
            radius=6,
            padding_x=4,
            padding_y=2,
            line_width=0.8,
        )

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(chip_x, chip_y, chip_w, chip_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.2)
        canvas_obj.drawString(chip_x + 6, chip_y + 10.8, self._operator_pdf_text("Stock"))
        qty_text = f"{self._fmt(qty)} {unit}"
        qty_font = _pdf_fit_font_size(qty_text, bold_font, chip_w - 12, 7.8, 5.8)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, qty_font)
        canvas_obj.drawString(chip_x + 6, chip_y + 3.8, self._operator_pdf_text(_pdf_clip_text(qty_text, chip_w - 12, bold_font, qty_font)))

        canvas_obj.setFillColor(palette["ink"])
        title_font = _pdf_fit_font_size("Etiqueta Produto", bold_font, title_w, 11.2, 8.8)
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 19.2, self._operator_pdf_text("Etiqueta Produto"))
        subtitle_font = _pdf_fit_font_size(code, regular_font, title_w, 6.3, 5.5)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_left + (title_w / 2.0),
            banner_y + 7.0,
            self._operator_pdf_text(_pdf_clip_text(code, title_w, regular_font, subtitle_font)),
        )

        body_x = outer_x + 10
        body_w = outer_w - 20
        body_top = banner_y - 10
        hero_h = 28
        hero_y = body_top - hero_h
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, hero_y, body_w, hero_h, 10, stroke=1, fill=1)
        code_font = _pdf_fit_font_size(code, bold_font, body_w - 18, 12.8, 9.8)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, code_font)
        canvas_obj.drawString(body_x + 8, hero_y + 16.8, self._operator_pdf_text(_pdf_clip_text(code, body_w - 16, bold_font, code_font)))

        desc_lines = _pdf_wrap_text(desc, regular_font, 7.0, body_w - 16, max_lines=2) or ["-"]
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 7.0)
        for line_index, line in enumerate(desc_lines):
            canvas_obj.drawString(body_x + 8, hero_y + 7.3 - (line_index * 7.2), self._operator_pdf_text(line))

        meta_y = hero_y - 20
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, meta_y, body_w, 20, 9, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.6)
        canvas_obj.drawString(body_x + 8, meta_y + 12.0, self._operator_pdf_text("Categoria / dimensao"))
        meta_font = _pdf_fit_font_size(meta_text, bold_font, body_w - 16, 7.4, 5.8)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, meta_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            meta_y + 4.2,
            self._operator_pdf_text(_pdf_clip_text(meta_text, body_w - 16, bold_font, meta_font)),
        )

        info_y = meta_y - 22
        info_gap = 6
        info_w = (body_w - info_gap) / 2.0
        info_cards = [
            ("Preco / unid.", f"{self._fmt(price_unit)} EUR"),
            ("Atualizado", updated_text),
        ]
        for index, (label, value) in enumerate(info_cards):
            box_x = body_x + (index * (info_w + info_gap))
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, info_y, info_w, 20, 8, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 5.1)
            canvas_obj.drawString(box_x + 7, info_y + 11.2, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, info_w - 14, 6.8, 5.3)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(
                box_x + 7,
                info_y + 4.0,
                self._operator_pdf_text(_pdf_clip_text(value, info_w - 14, bold_font, value_font)),
            )

        barcode_y = outer_y + 11
        barcode_h = info_y - barcode_y - 8
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, barcode_y, body_w, barcode_h, 9, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.6)
        canvas_obj.drawString(body_x + 8, barcode_y + barcode_h - 10.5, self._operator_pdf_text("Codigo para picagem"))
        barcode_area_x = body_x + 10
        barcode_area_w = body_w - 20
        barcode_bar_h = max(10.0, min(13.0, barcode_h - 16.0))
        barcode_draw_y = barcode_y + 9.0
        self._draw_code128_fit(canvas_obj, code, barcode_area_x, barcode_draw_y, barcode_area_w, barcode_bar_h, min_bar_width=0.42, max_bar_width=0.92)
        barcode_code_font = _pdf_fit_font_size(code, bold_font, barcode_area_w, 6.4, 5.6)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, barcode_code_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            barcode_y + 3.8,
            self._operator_pdf_text(_pdf_clip_text(code, barcode_area_w, bold_font, barcode_code_font)),
        )
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 4.9)
        canvas_obj.drawRightString(outer_x + outer_w - 8, outer_y + 4.6, self._operator_pdf_text(printed_at[:16]))

    def product_label_pdf(self, codigo: str, output_path: str | Path | None = None) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        product = self.product_detail(codigo)
        code = str(product.get("codigo", "") or "").strip() or "produto"
        target = (
            Path(output_path)
            if output_path
            else self._storage_output_path("products/labels", f"Etiqueta_Produto_{code}.pdf")
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        page_size = ((100 * mm), (70 * mm))
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
        self._draw_product_stock_label(canvas_obj, page_size[0], page_size[1], product, palette, logo_path, printed_at)
        canvas_obj.save()
        return target

    def product_open_label_pdf(self, codigo: str) -> Path:
        target = self.product_label_pdf(codigo)
        os.startfile(str(target))
        return target

    def get_encomenda_by_numero(self, numero: str) -> dict[str, Any] | None:
        numero = str(numero or "").strip()
        enc = next((e for e in self.ensure_data().get("encomendas", []) if str(e.get("numero", "")).strip() == numero), None)
        if enc is not None:
            enc.setdefault("montagem_itens", [])
            self._ensure_unique_order_piece_refs(enc)
        return enc

    def _ensure_unique_order_piece_refs(self, enc: dict[str, Any]) -> bool:
        data = self.ensure_data()
        cliente_codigo = str((enc or {}).get("cliente", "") or "").strip()
        if not cliente_codigo:
            return False
        seen: set[str] = set()
        changed = False
        pieces = list(self.desktop_main.encomenda_pecas(enc) or [])
        for piece in pieces:
            current_ref = str(piece.get("ref_interna", "") or "").strip()
            needs_new = (not current_ref) or (current_ref in seen)
            if needs_new:
                new_ref = str(self.desktop_main.next_ref_interna_unique(data, cliente_codigo, list(seen)))
                piece["ref_interna"] = new_ref
                current_ref = new_ref
                changed = True
            seen.add(current_ref)
            if str(piece.get("ref_externa", "") or "").strip():
                try:
                    self.desktop_main.update_refs(data, current_ref, str(piece.get("ref_externa", "") or "").strip())
                except Exception:
                    pass
        if changed:
            self.desktop_main.update_estado_encomenda_por_espessuras(enc)
            self._save(force=True)
        return changed

    def operator_names(self) -> list[str]:
        names: list[str] = []
        seen: set[str] = set()
        for raw in list(self.ensure_data().get("operadores", []) or []):
            if isinstance(raw, dict):
                values = [raw.get("nome"), raw.get("name"), raw.get("username"), raw.get("user"), raw.get("utilizador"), raw.get("id")]
            else:
                values = [raw]
            for value in values:
                txt = str(value or "").strip()
                key = txt.lower()
                if txt and key not in seen:
                    seen.add(key)
                    names.append(txt)
        for extra in (str((self.user or {}).get("username", "") or "").strip(),):
            key = extra.lower()
            if extra and key not in seen:
                seen.add(key)
                names.append(extra)
        return names

    def operator_avaria_options(self) -> list[str]:
        return list(self.operador_actions._metalurgica_paragem_options())

    def operator_interruption_options(self) -> list[str]:
        func = getattr(self.operador_actions, "_interrupcao_operacional_options", None)
        if callable(func):
            return list(func())
        return ["Mudanca de Turno", "Alteracao de Prioridades", "Aguardar material/documentacao", "Outro"]

    def operator_default_posto(self, operator_name: str = "") -> str:
        operator_txt = str(operator_name or "").strip()
        if not operator_txt:
            operator_txt = str((self.user or {}).get("username", "") or "").strip()
        if not operator_txt:
            return "Geral"
        profile = self._user_profile(operator_txt)
        profile_posto = str(profile.get("posto", "") or "").strip()
        if profile_posto:
            return profile_posto
        data = self.ensure_data()
        try:
            posto_map = dict(data.get("operador_posto_map", {}) or {})
        except Exception:
            posto_map = {}
        mapped = str(posto_map.get(operator_txt, "") or "").strip()
        if mapped:
            return mapped
        for user in list(data.get("users", []) or []):
            if not isinstance(user, dict):
                continue
            names = [
                str(user.get("username", "") or "").strip(),
                str(user.get("nome", "") or "").strip(),
                str(user.get("name", "") or "").strip(),
            ]
            if operator_txt not in names:
                continue
            for key in ("posto", "posto_trabalho", "work_center", "workcenter"):
                posto = str(user.get(key, "") or "").strip()
                if posto:
                    return posto
        return "Geral"

    def operator_has_posto_assignment(self, operator_name: str = "") -> bool:
        operator_txt = str(operator_name or "").strip()
        if not operator_txt:
            operator_txt = str((self.user or {}).get("username", "") or "").strip()
        if not operator_txt:
            return False
        profile = self._user_profile(operator_txt)
        if str(profile.get("posto", "") or "").strip():
            return True
        data = self.ensure_data()
        try:
            posto_map = dict(data.get("operador_posto_map", {}) or {})
        except Exception:
            posto_map = {}
        if str(posto_map.get(operator_txt, "") or "").strip():
            return True
        for user in list(data.get("users", []) or []):
            if not isinstance(user, dict):
                continue
            names = [
                str(user.get("username", "") or "").strip(),
                str(user.get("nome", "") or "").strip(),
                str(user.get("name", "") or "").strip(),
            ]
            if operator_txt not in names:
                continue
            if any(str(user.get(key, "") or "").strip() for key in ("posto", "posto_trabalho", "work_center", "workcenter")):
                return True
        return False

    def _event_ts(self, raw: Any) -> datetime | None:
        txt = str(raw or "").strip()
        if not txt:
            return None
        for candidate in (txt, txt.replace("Z", "+00:00")):
            try:
                return datetime.fromisoformat(candidate)
            except Exception:
                continue
        return None

    def operator_open_operation_elapsed_min(self, enc_num: str, piece_id: str, operation: str = "") -> float:
        enc_txt = str(enc_num or "").strip()
        piece_txt = str(piece_id or "").strip()
        op_norm = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        op_norm = str(op_norm or "").strip().lower()
        if not enc_txt or not piece_txt:
            return 0.0
        open_ts: datetime | None = None
        for row in sorted(list(self.ensure_data().get("op_eventos", []) or []), key=lambda r: str((r or {}).get("created_at", "") or "")):
            if not isinstance(row, dict):
                continue
            if str(row.get("encomenda_numero", "") or row.get("encomenda", "") or "").strip() != enc_txt:
                continue
            if str(row.get("peca_id", "") or "").strip() != piece_txt:
                continue
            event_norm = self.desktop_main.norm_text(row.get("evento", ""))
            row_op = self.desktop_main.normalize_operacao_nome(row.get("operacao", "")) or str(row.get("operacao", "") or "").strip()
            row_op_norm = str(row_op or "").strip().lower()
            if op_norm and row_op_norm and row_op_norm != op_norm:
                continue
            ts = self._event_ts(row.get("created_at", ""))
            if ts is None:
                continue
            if event_norm in ("start_op", "resume_piece"):
                open_ts = ts
            elif event_norm in ("finish_op", "pause_piece", "paragem"):
                open_ts = None
        if open_ts is None:
            return 0.0
        now_dt = self._event_ts(self.desktop_main.now_iso()) or datetime.now()
        return round(max(0.0, (now_dt - open_ts).total_seconds() / 60.0), 1)

    def _operator_ctx(self, operator_name: str = "", posto: str = "Geral") -> SimpleNamespace:
        return SimpleNamespace(
            data=self.ensure_data(),
            user=self.user or {},
            op_user=_ValueHolder(operator_name),
            op_posto=_ValueHolder(posto),
        )

    def _operator_info(self, operator_name: str, posto: str, text: str) -> str:
        ctx = self._operator_ctx(operator_name, posto)
        return str(self.operador_actions._format_event_info_with_posto(ctx, text) or text or "").strip()

    def _save_operator_state(self, enc: dict[str, Any]) -> None:
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)

    def _find_piece(self, enc_num: str, piece_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        for piece in self.desktop_main.encomenda_pecas(enc):
            if str(piece.get("id", "")).strip() == str(piece_id or "").strip():
                return enc, piece
        raise ValueError("Pe?a n?o encontrada.")

    def _operator_esp_obj(self, enc: dict[str, Any], material: str, espessura: str) -> dict[str, Any] | None:
        mat_txt = str(material or "").strip()
        esp_txt = str(espessura or "").strip()
        for mat in list(enc.get("materiais", []) or []):
            if str(mat.get("material", "") or "").strip() != mat_txt:
                continue
            for esp in list(mat.get("espessuras", []) or []):
                if str(esp.get("espessura", "") or "").strip() == esp_txt:
                    return esp
        return None

    def _norm_material_token(self, value: Any) -> str:
        return str(value or "").strip().lower()

    def _norm_esp_token(self, value: Any) -> str:
        txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
        txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
        if not txt:
            return ""
        try:
            num = float(txt)
        except Exception:
            return txt
        return str(int(num)) if num.is_integer() else f"{num:.6f}".rstrip("0").rstrip(".")

    def _is_laser_operation(self, operation: str) -> bool:
        return "laser" in self.desktop_main.norm_text(self.desktop_main.normalize_operacao_nome(operation or ""))

    def planning_operation_options(self) -> list[str]:
        return [
            str(row.get("name", "") or "").strip()
            for row in self.operation_catalog_rows()
            if bool(row.get("active", True)) and bool(row.get("planeavel", False))
        ]

    def _default_operation_catalog(self) -> list[dict[str, Any]]:
        planning_defaults = [
            "Corte Laser",
            "Quinagem",
            "Serralharia",
            "Maquinacao",
            "Roscagem",
            "Lacagem",
            "Montagem",
            "Embalamento",
            "Expedicao",
            "Furo Manual",
        ]
        non_planning_defaults = [
            "Departamento de Desenho",
            "Departamento de Orcamentacao",
            "Outros",
        ]
        names = list(dict.fromkeys(
            list(getattr(self.desktop_main, "OFF_OPERACOES_DISPONIVEIS", []) or [])
            + list(getattr(self.desktop_main, "PLANEAMENTO_OPERACOES_DISPONIVEIS", []) or [])
            + planning_defaults
            + non_planning_defaults
        ))
        planning_keys = {self.desktop_main.norm_text(name) for name in planning_defaults}
        return [
            {
                "name": self._planning_normalize_operation(name, default=str(name).strip()) if self.desktop_main.norm_text(name) in planning_keys else str(name).strip(),
                "active": True,
                "planeavel": self.desktop_main.norm_text(name) in planning_keys,
            }
            for name in names
            if str(name or "").strip()
        ]

    def operation_catalog_rows(self) -> list[dict[str, Any]]:
        data = self.ensure_data()
        raw_rows = list(data.get("operations_catalog", []) or [])
        seed_rows = self._default_operation_catalog() if not raw_rows else []
        seed_rows.extend(dict(row or {}) for row in raw_rows if isinstance(row, dict))
        for row in list(self._workcenter_catalog(sync_legacy=False) or []):
            op_name = self._planning_normalize_operation(row.get("operation", row.get("name", "")), default=str(row.get("operation", row.get("name", "")) or "").strip())
            if op_name:
                seed_rows.append({"name": op_name})
        merged: dict[str, dict[str, Any]] = {}
        for raw in seed_rows:
            name = str(raw.get("name", raw.get("nome", "")) or "").strip()
            if not name:
                continue
            normalized = self._planning_normalize_operation(name, default=name)
            if self.desktop_main.norm_text(name).startswith("departamento") or self.desktop_main.norm_text(name) == "outros":
                normalized = name
            key = normalized.casefold()
            current = merged.get(key)
            if current is None:
                current = {"name": normalized, "active": True, "planeavel": False}
                merged[key] = current
            if "active" in raw or "ativo" in raw:
                current["active"] = bool(raw.get("active", raw.get("ativo")))
            if "planeavel" in raw:
                current["planeavel"] = bool(raw.get("planeavel"))
        ordered_names = [row["name"] for row in self._default_operation_catalog()]
        order_index = {name.casefold(): index for index, name in enumerate(ordered_names)}
        cleaned = sorted(
            merged.values(),
            key=lambda row: (order_index.get(str(row.get("name", "")).casefold(), 999), str(row.get("name", "")).casefold()),
        )
        data["operations_catalog"] = cleaned
        return [dict(row) for row in cleaned]

    def operation_catalog_options(self, *, include_inactive: bool = False, planeavel_only: bool = False) -> list[str]:
        rows = []
        for row in self.operation_catalog_rows():
            if not include_inactive and not bool(row.get("active", True)):
                continue
            if planeavel_only and not bool(row.get("planeavel", False)):
                continue
            name = str(row.get("name", "") or "").strip()
            if name:
                rows.append(name)
        return rows

    def _operation_usage_counts(self, operation: str, *, data: dict[str, Any] | None = None) -> dict[str, int]:
        op_txt = self._planning_normalize_operation(operation, default=str(operation or "").strip())
        if not op_txt:
            return {"workcenters": 0, "quotes": 0, "orders": 0, "planning": 0, "total": 0}
        data = data if isinstance(data, dict) else self.ensure_data()
        op_key = self.desktop_main.norm_text(op_txt)

        def op_matches(value: Any) -> bool:
            normalized = self._planning_normalize_operation(value, default=str(value or "").strip())
            return bool(normalized) and self.desktop_main.norm_text(normalized) == op_key

        workcenters = sum(1 for row in list(data.get("workcenter_catalog", []) or []) if op_matches((row or {}).get("operation", (row or {}).get("name", ""))))
        quotes = 0
        for quote in list(data.get("orcamentos", []) or []):
            for line in list((quote or {}).get("linhas", []) or []):
                if any(op_matches(op_name) for op_name in self._planning_ops_from_ops_value((line or {}).get("operacao", ""))):
                    quotes += 1
                    break
        orders = 0
        for order in list(data.get("encomendas", []) or []):
            order_used = False
            for mat in list((order or {}).get("materiais", []) or []):
                for esp in list((mat or {}).get("espessuras", []) or []):
                    maps = [
                        dict((esp or {}).get("tempos_operacao", {}) or {}),
                        dict((esp or {}).get("maquinas_operacao", (esp or {}).get("recursos_operacao", {})) or {}),
                    ]
                    if any(op_matches(op_name) for values in maps for op_name in values.keys()):
                        order_used = True
                        break
                if order_used:
                    break
            if order_used:
                orders += 1
        planning = sum(
            1
            for bucket_name in ("plano", "plano_hist")
            for row in list(data.get(bucket_name, []) or [])
            if op_matches((row or {}).get("operacao", ""))
        )
        total = workcenters + quotes + orders + planning
        return {"workcenters": workcenters, "quotes": quotes, "orders": orders, "planning": planning, "total": total}

    def save_operation_catalog_row(self, name: str, *, current_name: str = "", active: bool = True, planeavel: bool = False) -> dict[str, Any]:
        data = self.ensure_data()
        new_name = str(name or "").strip()
        current_txt = str(current_name or "").strip()
        if not new_name:
            raise ValueError("Nome da operação obrigatório.")
        rows = self.operation_catalog_rows()
        new_key = new_name.casefold()
        current_key = current_txt.casefold()
        if any(str(row.get("name", "") or "").strip().casefold() == new_key and str(row.get("name", "") or "").strip().casefold() != current_key for row in rows):
            raise ValueError("Já existe uma operação com esse nome.")
        target = next((row for row in rows if str(row.get("name", "") or "").strip().casefold() == current_key), None) if current_txt else None
        if target is None:
            target = {"name": new_name}
            rows.append(target)
        old_name = str(target.get("name", "") or "").strip()
        target["name"] = self._planning_normalize_operation(new_name, default=new_name)
        if self.desktop_main.norm_text(new_name).startswith("departamento") or self.desktop_main.norm_text(new_name) == "outros":
            target["name"] = new_name
        target["active"] = bool(active)
        target["planeavel"] = bool(planeavel)
        if old_name and old_name.casefold() != str(target["name"]).casefold():
            for wc in list(data.get("workcenter_catalog", []) or []):
                if str((wc or {}).get("operation", "") or "").strip().casefold() == old_name.casefold():
                    wc["operation"] = target["name"]
            for bucket_name in ("plano", "plano_hist"):
                for block in list(data.get(bucket_name, []) or []):
                    if str((block or {}).get("operacao", "") or "").strip().casefold() == old_name.casefold():
                        block["operacao"] = target["name"]
        data["operations_catalog"] = rows
        self.operation_catalog_rows()
        self._workcenter_catalog()
        self._save(force=True)
        return next(row for row in self.operation_catalog_rows() if str(row.get("name", "") or "").strip().casefold() == str(target["name"]).casefold())

    def remove_operation_catalog_row(self, name: str) -> None:
        data = self.ensure_data()
        target = str(name or "").strip()
        if not target:
            raise ValueError("Operação inválida.")
        rows = self.operation_catalog_rows()
        current = next((row for row in rows if str(row.get("name", "") or "").strip().casefold() == target.casefold()), None)
        if current is None:
            raise ValueError("Operação não encontrada.")
        usage = self._operation_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover esta operação porque ainda está em uso "
                f"(postos: {usage['workcenters']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        data["operations_catalog"] = [
            row
            for row in rows
            if str(row.get("name", "") or "").strip().casefold() != target.casefold()
        ]
        self.operation_catalog_rows()
        self._save(force=True)

    def _planning_normalize_operation(self, operation: Any, default: str = "Corte Laser") -> str:
        normalize_fn = getattr(self.desktop_main, "normalize_planeamento_operacao", None)
        if callable(normalize_fn):
            normalized = str(normalize_fn(operation or "") or "").strip()
        else:
            normalized = str(self.desktop_main.normalize_operacao_nome(operation or "") or "").strip()
        if normalized:
            return normalized
        return str(default or "Corte Laser").strip() or "Corte Laser"

    def _planning_operation_aliases(self, operation: Any) -> set[str]:
        op_txt = self._planning_normalize_operation(operation)
        alias_map = {
            "Corte Laser": {"Corte Laser", "Laser"},
            "Quinagem": {"Quinagem"},
            "Serralharia": {"Serralharia", "Soldadura"},
            "Maquinacao": {"Maquinacao"},
            "Roscagem": {"Roscagem"},
            "Lacagem": {"Lacagem", "Pintura"},
            "Montagem": {"Montagem"},
            "Embalamento": {"Embalamento"},
            "Expedicao": {"Expedicao", "Expedição"},
            "Furo Manual": {"Furo Manual"},
        }
        return set(alias_map.get(op_txt, {op_txt}))

    def _planning_operation_from_piece_name(self, operation: Any) -> str:
        normalize_fn = getattr(self.desktop_main, "normalize_planeamento_operacao", None)
        if callable(normalize_fn):
            return str(normalize_fn(operation or "") or "").strip()
        return str(self.desktop_main.normalize_operacao_nome(operation or "") or "").strip()

    def _planning_operation_buffer_minutes(self) -> int:
        try:
            return max(0, int(float(self.ensure_data().get("planeamento_buffer_min", 15) or 15)))
        except Exception:
            return 15

    def _planning_apply_operation_sequence_rules(self, operations: list[str]) -> list[str]:
        ordered = [self._planning_normalize_operation(op, default="") for op in list(operations or []) if str(op or "").strip()]
        ordered = [op for op in ordered if op and op in self.planning_operation_options()]
        base = [op for op in ordered if op not in {"Embalamento", "Expedicao"}]
        if "Embalamento" in ordered:
            base.append("Embalamento")
        if "Expedicao" in ordered:
            base.append("Expedicao")
        return list(dict.fromkeys(base))

    def _planning_default_posto_for_operation(self, operation: Any, numero: str = "") -> str:
        op_txt = self._planning_normalize_operation(operation)
        if op_txt == "Corte Laser":
            posto_txt = self.workcenter_default_resource(op_txt, preferred=self._order_workcenter(numero))
            return posto_txt or "Corte Laser"
        return self.workcenter_default_resource(op_txt, preferred=op_txt) or op_txt or "Geral"

    def _planning_row_operation(self, row: dict[str, Any] | None, default: str = "Corte Laser") -> str:
        return self._planning_normalize_operation((row or {}).get("operacao", ""), default=default)

    def _planning_row_resource(self, row: dict[str, Any] | None, default: str = "") -> str:
        raw_row = dict(row or {})
        maquina_txt = self._normalize_workcenter_value(raw_row.get("maquina", ""))
        if maquina_txt:
            return maquina_txt
        op_txt = self._planning_row_operation(raw_row, default="")
        posto_txt = self._normalize_workcenter_value(raw_row.get("posto", ""))
        posto_trabalho_txt = self._normalize_workcenter_value(raw_row.get("posto_trabalho", ""))
        stored_txt = posto_txt or posto_trabalho_txt
        if stored_txt and op_txt:
            group_txt = self.workcenter_group_for_resource(stored_txt, op_txt) or self._legacy_workcenter_group_name(stored_txt)
            if group_txt and (
                stored_txt.lower() == group_txt.lower()
                or self.desktop_main.norm_text(stored_txt) in self._workcenter_group_aliases(group_txt)
            ):
                inferred = self._order_operation_resource(
                    str(raw_row.get("encomenda", "") or "").strip(),
                    str(raw_row.get("material", "") or "").strip(),
                    str(raw_row.get("espessura", "") or "").strip(),
                    op_txt,
                )
                if inferred:
                    return inferred
        if stored_txt:
            return stored_txt
        return str(default or "").strip()

    def _planning_apply_resource_to_row(self, row: dict[str, Any], resource: Any, operation: Any = "") -> dict[str, Any]:
        target = row if isinstance(row, dict) else {}
        op_txt = self._planning_normalize_operation(operation or target.get("operacao", ""), default="")
        resource_txt = self._normalize_workcenter_value(resource)
        if not resource_txt and op_txt:
            resource_txt = self._planning_default_posto_for_operation(op_txt, str(target.get("encomenda", "") or "").strip())
        posto_group = ""
        if resource_txt:
            posto_group = (
                self.workcenter_group_for_resource(resource_txt, op_txt)
                or self._legacy_workcenter_group_name(resource_txt)
                or resource_txt
            )
        if posto_group and resource_txt.lower() != posto_group.lower():
            target["maquina"] = resource_txt
            target["posto"] = posto_group
            target["posto_trabalho"] = posto_group
        else:
            target["maquina"] = ""
            target["posto"] = resource_txt or posto_group
            target["posto_trabalho"] = resource_txt or posto_group
        target["posto_grupo"] = posto_group or target.get("posto_grupo", "")
        return target

    def _planning_row_matches_operation(self, row: dict[str, Any] | None, operation: Any) -> bool:
        return self._planning_row_operation(row) == self._planning_normalize_operation(operation)

    def _planning_operation_times_map(self, esp_obj: dict[str, Any] | None) -> dict[str, str]:
        times: dict[str, str] = {}
        if not isinstance(esp_obj, dict):
            return times
        raw_map = dict(esp_obj.get("tempos_operacao", {}) or {})
        for op_name, raw_value in raw_map.items():
            op_txt = self._planning_normalize_operation(op_name)
            if not op_txt:
                continue
            value_txt = str(raw_value if raw_value is not None else "").strip()
            if value_txt:
                times[op_txt] = value_txt
        laser_txt = str(esp_obj.get("tempo_min", "") or "").strip()
        if laser_txt:
            times.setdefault("Corte Laser", laser_txt)
        return times

    def _planning_ops_from_piece(self, piece: dict[str, Any]) -> list[str]:
        ordered: list[str] = []
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            op_txt = self._planning_operation_from_piece_name(op.get("nome", ""))
            if not op_txt:
                continue
            if op_txt not in self.planning_operation_options():
                continue
            if op_txt not in ordered:
                ordered.append(op_txt)
        return ordered

    def _planning_ops_from_ops_value(self, value: Any) -> list[str]:
        parse_fn = getattr(self.desktop_main, "parse_planeamento_operacoes", None)
        if callable(parse_fn):
            ordered = list(parse_fn(value) or [])
        else:
            ordered = []
            for op in list(self.desktop_main.parse_operacoes_lista(value) or []):
                op_txt = self._planning_operation_from_piece_name(op)
                if op_txt and op_txt not in ordered:
                    ordered.append(op_txt)
        return [op for op in ordered if op in self.planning_operation_options()]

    def _planning_ops_from_esp_obj(self, esp_obj: dict[str, Any] | None) -> list[str]:
        ordered: list[str] = []
        for piece in list((esp_obj or {}).get("pecas", []) or []):
            for op_txt in self._planning_ops_from_piece(piece):
                if op_txt not in ordered:
                    ordered.append(op_txt)
        for op_txt in self._planning_operation_times_map(esp_obj):
            if op_txt not in ordered:
                ordered.append(op_txt)
        return ordered

    def _planning_piece_has_operation(self, piece: dict[str, Any], operation: Any) -> bool:
        aliases = self._planning_operation_aliases(operation)
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            normalized = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
            if normalized in aliases:
                return True
        return False

    def _planning_piece_operation_done(self, piece: dict[str, Any], operation: Any) -> bool:
        aliases = self._planning_operation_aliases(operation)
        matched = False
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            normalized = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
            if normalized not in aliases:
                continue
            matched = True
            if not getattr(self.desktop_main, "operacao_esta_concluida")(piece, op):
                return False
        return matched

    def _planning_item_has_operation(self, numero: str, material: str, espessura: str, operation: Any) -> bool:
        op_txt = self._planning_normalize_operation(operation)
        if op_txt == "Montagem":
            enc = self.get_encomenda_by_numero(str(numero or "").strip())
            return bool(list(self.desktop_main.encomenda_montagem_itens(enc) or []))
        if op_txt == "Corte Laser":
            return self._planning_item_has_laser(numero, material, espessura)
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        if not isinstance(esp_obj, dict):
            return False
        for piece in list(esp_obj.get("pecas", []) or []):
            if self._planning_piece_has_operation(piece, op_txt):
                return True
        time_map = self._planning_operation_times_map(esp_obj)
        return self._parse_float(time_map.get(op_txt, 0), 0) > 0

    def _planning_item_operation_done(self, numero: str, material: str, espessura: str, operation: Any) -> bool:
        op_txt = self._planning_normalize_operation(operation)
        if op_txt == "Montagem":
            enc = self.get_encomenda_by_numero(str(numero or "").strip())
            return str(self.desktop_main.encomenda_montagem_estado(enc) or "") == "Consumida"
        if op_txt == "Corte Laser":
            item = {"encomenda": numero, "material": material, "espessura": espessura}
            return bool(self.plan_actions._laser_done_for_item(self, item))
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        if not isinstance(esp_obj, dict):
            return False
        saw_any = False
        for piece in list(esp_obj.get("pecas", []) or []):
            if not self._planning_piece_has_operation(piece, op_txt):
                continue
            saw_any = True
            if not self._planning_piece_operation_done(piece, op_txt):
                return False
        return saw_any

    def _planning_item_operation_sequence(
        self,
        numero: str,
        material: str,
        espessura: str,
        start_operation: Any = "",
    ) -> list[str]:
        if self._planning_is_montagem_item(material, espessura):
            return ["Montagem"]
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        ordered: list[str] = []
        for op_txt in self._planning_ops_from_esp_obj(esp_obj):
            if op_txt in ordered:
                continue
            if not self._planning_item_has_operation(numero, material, espessura, op_txt):
                continue
            ordered.append(op_txt)
        start_txt = self._planning_normalize_operation(start_operation, default="") if str(start_operation or "").strip() else ""
        if not start_txt:
            return self._planning_apply_operation_sequence_rules(ordered)
        if start_txt in ordered:
            return self._planning_apply_operation_sequence_rules(ordered[ordered.index(start_txt) :])
        if self._planning_item_has_operation(numero, material, espessura, start_txt):
            return self._planning_apply_operation_sequence_rules([start_txt, *ordered])
        return self._planning_apply_operation_sequence_rules(ordered)

    def _planning_slot_datetime(self, day_txt: str, start_min: int) -> datetime | None:
        try:
            base_dt = datetime.fromisoformat(str(day_txt or "").strip())
        except Exception:
            return None
        return base_dt.replace(hour=max(0, int(start_min // 60)), minute=max(0, int(start_min % 60)), second=0, microsecond=0)

    def _planning_cursor_from_datetime(self, dates: list[date], anchor_dt: datetime | None) -> tuple[int, int]:
        start_min, end_min, slot = self._planning_grid_metrics()
        if not dates or anchor_dt is None:
            return 0, start_min
        if anchor_dt.date() < dates[0]:
            return 0, start_min
        if anchor_dt.date() > dates[-1]:
            return len(dates), start_min
        day_idx = max(0, (anchor_dt.date() - dates[0]).days)
        cursor = (anchor_dt.hour * 60) + anchor_dt.minute
        if cursor % slot != 0:
            cursor = int((cursor + slot - 1) // slot) * slot
        if cursor < start_min:
            cursor = start_min
        if cursor >= end_min:
            return day_idx + 1, start_min
        return day_idx, cursor

    def _planning_item_operation_range(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any,
        buckets: tuple[str, ...] = ("plano", "plano_hist"),
    ) -> tuple[datetime | None, datetime | None]:
        target = self._planning_item_op_key(numero, material, espessura, operation)
        op_txt = self._planning_normalize_operation(operation)
        first_start: datetime | None = None
        last_end: datetime | None = None
        for bucket_name in buckets:
            for row in list(self.ensure_data().get(bucket_name, []) or []):
                if not isinstance(row, dict):
                    continue
                row_key = self._planning_item_op_key(
                    row.get("encomenda", ""),
                    row.get("material", ""),
                    row.get("espessura", ""),
                    self._planning_row_operation(row),
                )
                if row_key != target:
                    continue
                start_dt, end_dt = self._planning_block_bounds(row)
                if start_dt is not None and (first_start is None or start_dt < first_start):
                    first_start = start_dt
                if end_dt is not None and (last_end is None or end_dt > last_end):
                    last_end = end_dt
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        if op_txt == "Corte Laser":
            esp_obj = self._planning_find_esp_obj(enc, material, espessura)
            raw_finished = str((esp_obj or {}).get("laser_concluido_em", "") or "").strip()
            if raw_finished:
                try:
                    finished_dt = datetime.fromisoformat(raw_finished)
                    if first_start is None:
                        first_start = finished_dt
                    if last_end is None or finished_dt > last_end:
                        last_end = finished_dt
                except Exception:
                    pass
        elif op_txt == "Montagem" and isinstance(enc, dict):
            consumed_marks: list[datetime] = []
            for item in list(enc.get("montagem_itens", []) or []):
                raw_consumed = str(item.get("consumed_at", "") or "").strip()
                if not raw_consumed:
                    continue
                try:
                    consumed_marks.append(datetime.fromisoformat(raw_consumed))
                except Exception:
                    continue
            if consumed_marks:
                consumed_dt = max(consumed_marks)
                if first_start is None:
                    first_start = consumed_dt
                if last_end is None or consumed_dt > last_end:
                    last_end = consumed_dt
        return first_start, last_end

    def _planning_item_operation_status(self, numero: str, material: str, espessura: str, operation: Any) -> dict[str, Any]:
        total = self._planning_item_total_minutes(numero, material, espessura, operation=operation)
        planned = min(total, self._planning_planned_minutes(numero, material, espessura, operation=operation)) if total > 0 else 0
        first_dt, end_dt = self._planning_item_operation_range(numero, material, espessura, operation)
        resolved = bool(self._planning_item_operation_done(numero, material, espessura, operation))
        if total > 0 and planned >= total:
            resolved = True
        return {
            "total_min": total,
            "planned_min": planned,
            "resolved": resolved,
            "first_dt": first_dt,
            "end_dt": end_dt,
        }

    def _planning_schedule_operation_blocks(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any,
        dates: list[date],
        *,
        anchor_dt: datetime | None = None,
        resource: str = "",
    ) -> dict[str, Any]:
        op_txt = self._planning_normalize_operation(operation)
        start_min, end_min, slot = self._planning_grid_metrics()
        existing_first, existing_last = self._planning_item_operation_range(numero, material, espessura, op_txt)
        remaining = self._planning_remaining_minutes(numero, material, espessura, operation=op_txt)
        resource_txt = (
            self._normalize_workcenter_value(resource)
            or self._order_operation_resource(numero, material, espessura, op_txt)
            or self._planning_default_posto_for_operation(op_txt, numero)
        )
        effective_anchor = anchor_dt
        if existing_last is not None and (effective_anchor is None or existing_last > effective_anchor):
            effective_anchor = existing_last
        cursor_day_idx, cursor_min = self._planning_cursor_from_datetime(dates, effective_anchor)
        if remaining <= 0:
            return {
                "placed": [],
                "exhausted": False,
                "remaining_min": 0,
                "first_dt": existing_first,
                "end_dt": existing_last,
                "cursor_day_idx": cursor_day_idx,
                "cursor_min": cursor_min,
                "resource": resource_txt,
            }
        if cursor_day_idx >= len(dates):
            return {
                "placed": [],
                "exhausted": True,
                "remaining_min": remaining,
                "first_dt": existing_first,
                "end_dt": existing_last,
                "cursor_day_idx": cursor_day_idx,
                "cursor_min": start_min,
                "resource": resource_txt,
            }
        item_color = self._planning_item_color(numero, material, espessura)
        placed: list[dict[str, Any]] = []
        last_end = existing_last
        first_dt = existing_first
        cursor_dt_day_idx = cursor_day_idx
        cursor_dt_min = cursor_min
        while remaining > 0:
            next_day_idx, day_txt, segment_start, segment_end = self._planning_next_free_segment(
                dates,
                cursor_dt_day_idx,
                cursor_dt_min,
                operation=op_txt,
                resource=resource_txt,
            )
            if not day_txt or segment_start is None or segment_end is None:
                return {
                    "placed": placed,
                    "exhausted": True,
                    "remaining_min": remaining,
                    "first_dt": first_dt,
                    "end_dt": last_end,
                    "cursor_day_idx": next_day_idx,
                    "cursor_min": cursor_dt_min,
                    "resource": resource_txt,
                }
            free_minutes = max(0, int(segment_end - segment_start))
            if free_minutes <= 0:
                return {
                    "placed": placed,
                    "exhausted": True,
                    "remaining_min": remaining,
                    "first_dt": first_dt,
                    "end_dt": last_end,
                    "cursor_day_idx": next_day_idx,
                    "cursor_min": cursor_dt_min,
                    "resource": resource_txt,
                }
            chunk = min(remaining, free_minutes)
            if chunk % slot != 0:
                chunk = max(slot, int(chunk // slot) * slot)
            block = self._planning_make_block(
                numero,
                material,
                espessura,
                op_txt,
                day_txt,
                segment_start,
                chunk,
                color=item_color,
                posto=resource_txt,
            )
            self.ensure_data().setdefault("plano", []).append(block)
            placed.append(block)
            start_dt = self._planning_slot_datetime(day_txt, segment_start)
            end_dt = (start_dt + timedelta(minutes=chunk)) if start_dt is not None else None
            if start_dt is not None and (first_dt is None or start_dt < first_dt):
                first_dt = start_dt
            if end_dt is not None and (last_end is None or end_dt > last_end):
                last_end = end_dt
            remaining -= chunk
            cursor_dt_day_idx = next_day_idx
            cursor_dt_min = segment_start + chunk
            if cursor_dt_min >= end_min:
                cursor_dt_day_idx += 1
                cursor_dt_min = start_min
        return {
            "placed": placed,
            "exhausted": False,
            "remaining_min": 0,
            "first_dt": first_dt,
            "end_dt": last_end,
            "cursor_day_idx": cursor_dt_day_idx,
            "cursor_min": cursor_dt_min,
            "resource": resource_txt,
        }

    def _planning_delivery_sort_key(self, value: Any) -> tuple[str, str]:
        raw = str(value or "").strip()
        if not raw:
            return ("9999-99-99", "")
        try:
            parsed = datetime.fromisoformat(raw[:10]).date()
            return (parsed.isoformat(), raw)
        except Exception:
            return ("9999-99-99", raw)

    def _planning_schedule_followup_jobs(
        self,
        flow_jobs: list[dict[str, Any]],
        dates: list[date],
        *,
        initial_cursor_dt: datetime | None = None,
    ) -> dict[str, Any]:
        placed: list[dict[str, Any]] = []
        pending: list[dict[str, Any]] = []
        downstream_cursors: dict[tuple[str, str], datetime | None] = {}
        active_jobs = [dict(job or {}) for job in list(flow_jobs or []) if isinstance(job, dict)]

        def later_dt(left: datetime | None, right: datetime | None) -> datetime | None:
            if left is None:
                return right
            if right is None:
                return left
            return right if right > left else left

        while active_jobs:
            grouped_jobs: dict[tuple[str, str], list[dict[str, Any]]] = {}
            for job in active_jobs:
                sequence = [str(op or "").strip() for op in list(job.get("sequence", []) or []) if str(op or "").strip()]
                index = int(job.get("index", 0) or 0)
                if index < 0 or index >= len(sequence):
                    continue
                op_name = self._planning_normalize_operation(sequence[index])
                if not op_name:
                    continue
                resource_txt = self._normalize_workcenter_value(job.get("resource", ""))
                if not resource_txt:
                    resource_txt = self._order_operation_resource(
                        str(job.get("numero", "") or "").strip(),
                        str(job.get("material", "") or "").strip(),
                        str(job.get("espessura", "") or "").strip(),
                        op_name,
                    )
                key = (op_name, resource_txt.lower())
                payload = dict(job)
                payload["resource"] = resource_txt
                payload["sequence"] = sequence
                grouped_jobs.setdefault(key, []).append(payload)

            next_round: list[dict[str, Any]] = []
            for key, jobs in grouped_jobs.items():
                op_name = str(key[0] or "").strip()
                cursor_dt = later_dt(initial_cursor_dt, downstream_cursors.get(key))
                remaining_jobs = list(jobs)
                while remaining_jobs:
                    ready_jobs = [
                        job
                        for job in remaining_jobs
                        if job.get("anchor_dt") is None or (cursor_dt is not None and job.get("anchor_dt") <= cursor_dt)
                    ]
                    if not ready_jobs:
                        next_anchor = min(
                            (
                                job.get("anchor_dt")
                                for job in remaining_jobs
                                if isinstance(job.get("anchor_dt"), datetime)
                            ),
                            default=None,
                        )
                        cursor_dt = later_dt(cursor_dt, next_anchor)
                        ready_jobs = [
                            job
                            for job in remaining_jobs
                            if job.get("anchor_dt") is None or (cursor_dt is not None and job.get("anchor_dt") <= cursor_dt)
                        ]
                    if not ready_jobs:
                        ready_jobs = list(remaining_jobs)
                    ready_jobs.sort(
                        key=lambda job: (
                            self._planning_delivery_sort_key(job.get("data_entrega", "")),
                            job.get("anchor_dt") or datetime.max,
                            str(job.get("numero", "") or "").strip(),
                            str(job.get("material", "") or "").strip(),
                            self._parse_float(job.get("espessura", 0), 0),
                        )
                    )
                    job = ready_jobs[0]
                    remaining_jobs.remove(job)
                    numero = str(job.get("numero", "") or "").strip()
                    material = str(job.get("material", "") or "").strip()
                    espessura = str(job.get("espessura", "") or "").strip()
                    resource_txt = str(job.get("resource", "") or "").strip()
                    schedule_anchor = later_dt(job.get("anchor_dt"), cursor_dt)
                    result = self._planning_schedule_operation_blocks(
                        numero,
                        material,
                        espessura,
                        op_name,
                        dates,
                        anchor_dt=schedule_anchor,
                        resource=resource_txt,
                    )
                    placed.extend(list(result.get("placed", []) or []))
                    cursor_dt = later_dt(cursor_dt, result.get("end_dt"))
                    downstream_cursors[key] = later_dt(downstream_cursors.get(key), result.get("end_dt"))
                    if bool(result.get("exhausted")) and int(result.get("remaining_min", 0) or 0) > 0:
                        pending.append(
                            {
                                "numero": numero,
                                "material": material,
                                "espessura": espessura,
                                "operacao": op_name,
                                "recurso": str(result.get("resource", "") or resource_txt),
                                "remaining_min": int(result.get("remaining_min", 0) or 0),
                            }
                        )
                        continue
                    next_index = int(job.get("index", 0) or 0) + 1
                    sequence = list(job.get("sequence", []) or [])
                    if next_index >= len(sequence):
                        continue
                    next_op = self._planning_normalize_operation(sequence[next_index])
                    next_resource = self._order_operation_resource(numero, material, espessura, next_op)
                    end_anchor = result.get("end_dt")
                    if isinstance(end_anchor, datetime):
                        end_anchor = end_anchor + timedelta(minutes=self._planning_operation_buffer_minutes())
                    next_round.append(
                        {
                            "numero": numero,
                            "material": material,
                            "espessura": espessura,
                            "sequence": sequence,
                            "index": next_index,
                            "anchor_dt": end_anchor,
                            "resource": next_resource,
                            "data_entrega": str(job.get("data_entrega", "") or "").strip(),
                        }
                    )
            active_jobs = next_round
        return {"placed": placed, "pending": pending}

    def _operator_group_total_output(self, esp_obj: dict[str, Any] | None) -> float:
        total = 0.0
        for piece in list((esp_obj or {}).get("pecas", []) or []):
            total += (
                self._parse_float(piece.get("produzido_ok", 0), 0)
                + self._parse_float(piece.get("produzido_nok", 0), 0)
                + self._parse_float(piece.get("produzido_qualidade", 0), 0)
            )
        return round(total, 1)

    def _operator_esp_laser_concluido(self, esp_obj: dict[str, Any] | None) -> bool:
        if not esp_obj:
            return False
        saw_laser = False
        for piece in list(esp_obj.get("pecas", []) or []):
            ops = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
            piece_has_laser = False
            piece_laser_done = False
            for op in ops:
                op_name = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
                if not self._is_laser_operation(op_name):
                    continue
                piece_has_laser = True
                saw_laser = True
                if "concl" in self.desktop_main.norm_text(op.get("estado", "")):
                    piece_laser_done = True
            if piece_has_laser and not piece_laser_done:
                return False
        return saw_laser

    def _operator_esp_laser_resolved(self, esp_obj: dict[str, Any] | None) -> bool:
        if not esp_obj:
            return False
        return bool(esp_obj.get("baixa_laser_feita")) or bool(esp_obj.get("baixa_laser_confirmada_sem_baixa"))

    def _piece_operation_row(self, piece: dict[str, Any], operation: str) -> dict[str, Any] | None:
        target = self.desktop_main.normalize_operacao_nome(operation or "")
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            if self.desktop_main.normalize_operacao_nome(op.get("nome", "")) == target:
                return op
        return None

    def _piece_operation_limit(self, piece: dict[str, Any], operation: str) -> float:
        return self._parse_float(getattr(self.desktop_main, "operacao_input_qtd")(piece, operation), 0)

    def _piece_operation_total(self, op_row: dict[str, Any] | None, limit: float = 0.0) -> float:
        if not isinstance(op_row, dict):
            return 0.0
        total = (
            self._parse_float(op_row.get("qtd_ok", 0), 0)
            + self._parse_float(op_row.get("qtd_nok", 0), 0)
            + self._parse_float(op_row.get("qtd_qual", 0), 0)
        )
        if total > 0:
            return round(total, 4)
        if "concl" in self.desktop_main.norm_text(op_row.get("estado", "")) and limit > 0:
            return round(limit, 4)
        return 0.0

    def _sync_piece_output_from_flow(self, piece: dict[str, Any]) -> None:
        fluxo = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
        final_ok = 0.0
        final_nok = 0.0
        final_qual = 0.0
        for op in fluxo:
            total = self._piece_operation_total(op, self._piece_operation_limit(piece, op.get("nome", "")))
            if total <= 0:
                continue
            final_ok = self._parse_float(op.get("qtd_ok", 0), 0)
            final_nok = self._parse_float(op.get("qtd_nok", 0), 0)
            final_qual = self._parse_float(op.get("qtd_qual", 0), 0)
        piece["produzido_ok"] = round(final_ok, 4)
        piece["produzido_nok"] = round(final_nok, 4)
        piece["produzido_qualidade"] = round(final_qual, 4)

    def operator_laser_stock_state(self, enc_num: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura n?o encontrado.")
        total_qty = self._operator_group_total_output(esp_obj)
        reserved_qty = 0.0
        reserved_sources: list[dict[str, Any]] = []
        for row in list(enc.get("reservas", []) or []):
            if self._norm_material_token(row.get("material")) != self._norm_material_token(material):
                continue
            if self._norm_esp_token(row.get("espessura")) != self._norm_esp_token(espessura):
                continue
            qty_res = self._parse_float(row.get("quantidade", 0), 0)
            reserved_qty += qty_res
            stock = self.material_by_id(str(row.get("material_id", "") or "").strip()) if str(row.get("material_id", "") or "").strip() else None
            reserved_sources.append(
                {
                    "material_id": str(row.get("material_id", "") or "").strip(),
                    "dimensao": f"{(stock or {}).get('comprimento', '')}x{(stock or {}).get('largura', '')}",
                    "disponivel": round(qty_res, 2),
                    "local": self._localizacao(stock) if stock else "-",
                    "lote": str((stock or {}).get("lote_fornecedor", "") or row.get("lote", "") or "").strip(),
                    "reserved": True,
                }
            )
        manual_stock_required = reserved_qty <= 1e-9
        return {
            "encomenda": enc_num,
            "material": str(material or "").strip(),
            "espessura": str(espessura or "").strip(),
            "total_qty": total_qty,
            "reserved_qty": round(reserved_qty, 1),
            "remaining_qty": 0.0,
            "manual_stock_required": manual_stock_required,
            "laser_complete": self._operator_esp_laser_concluido(esp_obj),
            "resolved": self._operator_esp_laser_resolved(esp_obj),
            "lote_baixa": str(esp_obj.get("lote_baixa", "") or "").strip(),
            "candidates": self.order_stock_candidates(enc_num, material, espessura),
            "reserved_sources": reserved_sources,
        }

    def operator_resolve_laser_stock(
        self,
        enc_num: str,
        material: str,
        espessura: str,
        material_id: str = "",
        quantidade: Any = 0,
        allow_without_stock: bool = False,
        retalho: dict[str, Any] | None = None,
        source_material_id: str = "",
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura n?o encontrado.")
        if not self._operator_esp_laser_concluido(esp_obj):
            return {"resolved": False, "reason": "laser_not_complete"}
        if self._operator_esp_laser_resolved(esp_obj):
            state = self.operator_laser_stock_state(enc_num, material, espessura)
            state["resolved"] = True
            return state
        total_qty = self._operator_group_total_output(esp_obj)
        lote_sel = str(esp_obj.get("lote_baixa", "") or "").strip()
        reserved_consumed = 0.0
        keep_reservas: list[dict[str, Any]] = []
        for row in list(enc.get("reservas", []) or []):
            if self._norm_material_token(row.get("material")) != self._norm_material_token(material):
                keep_reservas.append(row)
                continue
            if self._norm_esp_token(row.get("espessura")) != self._norm_esp_token(espessura):
                keep_reservas.append(row)
                continue
            qty_res = self._parse_float(row.get("quantidade", 0), 0)
            if qty_res <= 0:
                continue
            stock = None
            ref_material_id = str(row.get("material_id", "") or "").strip()
            if ref_material_id:
                stock = self.material_by_id(ref_material_id)
            if stock is None:
                for candidate in list(self.ensure_data().get("materiais", []) or []):
                    if self._norm_material_token(candidate.get("material")) != self._norm_material_token(material):
                        continue
                    if self._norm_esp_token(candidate.get("espessura")) != self._norm_esp_token(espessura):
                        continue
                    stock = candidate
                    break
            if stock is None:
                keep_reservas.append(row)
                continue
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material reservado {stock.get('id', '')} bloqueado pela qualidade.")
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - qty_res)
            stock["reservado"] = max(0.0, self._parse_float(stock.get("reservado", 0), 0) - qty_res)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            if not lote_sel:
                lote_sel = str(stock.get("lote_fornecedor", "") or "").strip()
            reserved_consumed += qty_res
            self.desktop_main.log_stock(
                self.ensure_data(),
                "BAIXA CATIVADA",
                f"{stock.get('id', '')} qtd={qty_res} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )
        enc["reservas"] = keep_reservas
        enc["cativar"] = bool(keep_reservas)

        extra_consumed = 0.0
        stock_id = str(material_id or "").strip()
        extra_qty = self._parse_float(quantidade, 0)
        retalho_payload = dict(retalho or {})
        has_retalho = any(str(retalho_payload.get(key, "")).strip() for key in ("comprimento", "largura", "quantidade", "metros"))
        chosen_source_id = str(source_material_id or "").strip()
        if reserved_consumed > 0 and stock_id and extra_qty > 0:
            raise ValueError("Com material cativado n?o ? permitida baixa manual adicional. Apenas podes registar retalho.")
        if stock_id and extra_qty > 0:
            stock = self.material_by_id(stock_id)
            if stock is None:
                raise ValueError("Material n?o encontrado para baixa.")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {stock_id} bloqueado pela qualidade.")
            if self._norm_material_token(stock.get("material")) != self._norm_material_token(material) or self._norm_esp_token(stock.get("espessura")) != self._norm_esp_token(espessura):
                raise ValueError("O stock selecionado n?o corresponde ao material/espessura.")
            if extra_qty > self._parse_float(stock.get("quantidade", 0), 0):
                raise ValueError("Quantidade superior ao stock disponivel.")
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - extra_qty)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            if not lote_sel:
                lote_sel = str(stock.get("lote_fornecedor", "") or "").strip()
            extra_consumed = extra_qty
            self.desktop_main.log_stock(
                self.ensure_data(),
                "BAIXA",
                f"{stock_id} qtd={extra_qty} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )

        consumed_total = round(reserved_consumed + extra_consumed, 1)
        manual_stock_required = reserved_consumed <= 1e-9
        if manual_stock_required and extra_consumed <= 1e-9 and not allow_without_stock:
            raise ValueError("Falta registar a baixa do material consumido.")
        if manual_stock_required and extra_consumed <= 1e-9 and allow_without_stock:
            self.desktop_main.log_stock(
                self.ensure_data(),
                "SEM_BAIXA",
                f"encomenda={enc.get('numero', '')} mat={material} esp={espessura} motivo=laser_sem_stock_qt",
                operador=self._current_user_label(),
            )
        created_retalho = None
        if has_retalho:
            if not chosen_source_id:
                reserved_ids = [str(row.get("material_id", "") or "").strip() for row in list(enc.get("reservas", []) or []) if self._norm_material_token(row.get("material")) == self._norm_material_token(material) and self._norm_esp_token(row.get("espessura")) == self._norm_esp_token(espessura) and str(row.get("material_id", "") or "").strip()]
                candidate_ids = list(dict.fromkeys([*reserved_ids, stock_id]))
                candidate_ids = [row for row in candidate_ids if row]
                if len(candidate_ids) == 1:
                    chosen_source_id = candidate_ids[0]
                elif len(candidate_ids) > 1:
                    raise ValueError("Seleciona o lote de origem do retalho.")
            source_stock = self.material_by_id(chosen_source_id) if chosen_source_id else None
            if source_stock is None:
                raise ValueError("Lote de origem do retalho n?o encontrado.")
            comp = self._parse_float(retalho_payload.get("comprimento", 0), 0)
            larg = self._parse_float(retalho_payload.get("largura", 0), 0)
            q_retalho = self._parse_float(retalho_payload.get("quantidade", 0), 0)
            metros = self._parse_float(retalho_payload.get("metros", 0), 0)
            if q_retalho <= 0:
                raise ValueError("Quantidade do retalho invalida.")
            created_retalho = {
                "id": self._next_material_id(),
                "formato": source_stock.get("formato", self.desktop_main.detect_materia_formato(source_stock)),
                "material": source_stock.get("material", ""),
                "espessura": source_stock.get("espessura", ""),
                "comprimento": comp,
                "largura": larg,
                "metros": metros,
                "quantidade": q_retalho,
                "reservado": 0.0,
                "Localização": "RETALHO",
                "Localizacao": "RETALHO",
                "lote_fornecedor": source_stock.get("lote_fornecedor", ""),
                "peso_unid": 0.0,
                "p_compra": source_stock.get("p_compra", 0),
                "preco_unid": 0.0,
                "is_sobra": True,
                "origem_material_id": str(source_stock.get("id", "") or "").strip(),
                "origem_lote": str(source_stock.get("lote_fornecedor", "") or "").strip(),
                "origem_lotes_baixa": [lot for lot in list(dict.fromkeys([lote_sel, str(source_stock.get('lote_fornecedor', '') or '').strip()])) if lot],
                "atualizado_em": self.desktop_main.now_iso(),
            }
            self.materia_actions._hydrate_retalho_record(self.ensure_data(), created_retalho, template=source_stock)
            self.ensure_data().setdefault("materiais", []).append(created_retalho)
            self.desktop_main.log_stock(
                self.ensure_data(),
                "RETALHO",
                f"{source_stock.get('id', '')} qtd={created_retalho.get('quantidade', 0)} encomenda={enc.get('numero', '')} motivo=fecho_laser_qt",
                operador=self._current_user_label(),
            )
        if lote_sel:
            esp_obj["lote_baixa"] = lote_sel
            for piece in list(esp_obj.get("pecas", []) or []):
                piece["lote_baixa"] = lote_sel
        if not esp_obj.get("laser_concluido"):
            esp_obj["laser_concluido"] = True
            esp_obj["laser_concluido_em"] = self.desktop_main.now_iso()
        esp_obj["baixa_laser_feita"] = bool((reserved_consumed > 0) or (extra_consumed > 0))
        esp_obj["baixa_laser_confirmada_sem_baixa"] = bool(allow_without_stock and manual_stock_required and extra_consumed <= 0)
        esp_obj["baixa_laser_em"] = self.desktop_main.now_iso()
        self._sync_ne_from_materia()
        self._save_operator_state(enc)
        return {
            "resolved": True,
            "total_qty": total_qty,
            "reserved_consumed": round(reserved_consumed, 1),
            "extra_consumed": round(extra_consumed, 1),
            "consumed_total": consumed_total,
            "remaining_qty": 0.0,
            "allow_without_stock": bool(allow_without_stock and manual_stock_required and extra_consumed <= 0),
            "manual_stock_required": manual_stock_required,
            "lote_baixa": lote_sel,
            "retalho_id": str((created_retalho or {}).get("id", "") or "").strip(),
        }

    def operator_piece_context(self, enc_num: str, piece_id: str) -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        self.desktop_main.ensure_peca_operacoes(piece)
        avaria_index = self.operador_actions._op_open_avaria_index(self.ensure_data(), str(enc.get("numero", "") or ""))
        live_row = self.operador_actions._op_live_avaria_row_for_piece(avaria_index, piece)
        if live_row:
            self.operador_actions._op_sync_piece_live_avaria(piece, live_row)
        qtd_total = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        ok = self._parse_float(piece.get("produzido_ok", 0), 0)
        nok = self._parse_float(piece.get("produzido_nok", 0), 0)
        qual = self._parse_float(piece.get("produzido_qualidade", 0), 0)
        raw_pending_ops = list(self.desktop_main.peca_operacoes_pendentes(piece))
        done_ops = list(self.desktop_main.peca_operacoes_concluidas(piece))
        pending_ops: list[str] = []
        op_limits: dict[str, float] = {}
        op_totals: dict[str, float] = {}
        for op_name in raw_pending_ops:
            limit = self._piece_operation_limit(piece, op_name)
            total_done = self._piece_operation_total(self._piece_operation_row(piece, op_name), limit)
            op_limits[op_name] = limit
            op_totals[op_name] = total_done
            if limit > total_done:
                pending_ops.append(op_name)
        live_status_map: dict[str, dict[str, Any]] = {}
        status_fn = getattr(self.operador_actions, "_mysql_ops_status_for_piece", None)
        if callable(status_fn):
            try:
                for row in list(status_fn(str(enc.get("numero", "") or ""), str(piece.get("id", "") or "")) or []):
                    op_name = self.desktop_main.normalize_operacao_nome((row or {}).get("operacao", "")) or str((row or {}).get("operacao", "") or "").strip()
                    if op_name:
                        live_status_map[op_name] = dict(row or {})
            except Exception:
                live_status_map = {}
        active_pending_ops = []
        for op_name in pending_ops:
            live_state = self.desktop_main.norm_text((live_status_map.get(op_name, {}) or {}).get("estado", ""))
            if "produ" in live_state:
                active_pending_ops.append(op_name)
        current_op = active_pending_ops[0] if active_pending_ops else (self.desktop_main.normalize_operacao_nome(piece.get("operacao_atual", "")) or (pending_ops[0] if pending_ops else ""))
        current_limit = op_limits.get(current_op, self._piece_operation_limit(piece, current_op))
        current_done = op_totals.get(current_op, self._piece_operation_total(self._piece_operation_row(piece, current_op), current_limit))
        avaria_closed_min = self.operador_actions._op_piece_closed_avaria_minutes(self.ensure_data(), enc_num, piece)
        avaria_open_min = self.operador_actions._op_piece_current_avaria_minutes(self.ensure_data(), enc_num, piece, live_row=live_row)
        return {
            "encomenda": enc,
            "piece": piece,
            "pending_ops": pending_ops,
            "active_pending_ops": active_pending_ops,
            "done_ops": done_ops,
            "has_open_avaria": bool(self.operador_actions._op_piece_has_open_avaria(self.ensure_data(), enc_num, piece, avaria_index=avaria_index)),
            "avaria_motivo": str((live_row or {}).get("causa", "") or piece.get("avaria_motivo", "") or piece.get("interrupcao_peca_motivo", "") or "").strip(),
            "quantidade_pedida": qtd_total,
            "produzido_ok": ok,
            "produzido_nok": nok,
            "produzido_qualidade": qual,
            "default_ok": current_done if current_done > 0 else current_limit,
            "current_operation": current_op,
            "current_operation_elapsed_min": self.operator_open_operation_elapsed_min(enc_num, piece_id, current_op),
            "current_operation_limit": current_limit,
            "current_operation_done": current_done,
            "operation_limits": op_limits,
            "operation_done": op_totals,
            "avaria_closed_min": round(avaria_closed_min, 2),
            "avaria_open_min": round(avaria_open_min, 2),
            "avaria_total_min": round(avaria_closed_min + avaria_open_min, 2),
        }

    def operator_start_piece(self, enc_num: str, piece_id: str, operator_name: str, operation: str = "", posto: str = "Geral") -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        ctx = self.operator_piece_context(enc_num, piece_id)
        if ctx["has_open_avaria"]:
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de iniciar a peca.")
        pending = list(ctx["pending_ops"])
        if not pending:
            raise ValueError("Esta pe?a n?o tem opera??es pendentes.")
        selected_op = self.desktop_main.normalize_operacao_nome(operation or pending[0])
        if selected_op not in pending:
            raise ValueError("A opera??o selecionada n?o est? pendente.")
        available_qty = self._piece_operation_limit(piece, selected_op)
        current_qty = self._piece_operation_total(self._piece_operation_row(piece, selected_op), available_qty)
        if available_qty <= current_qty:
            raise ValueError("Esta opera??o n?o tem quantidade dispon?vel do posto anterior.")
        result = self.operador_actions._mysql_ops_acquire(
            str(enc.get("numero", "") or ""),
            str(piece.get("id", "") or ""),
            [selected_op],
            operator_name,
            valid_operators=self.operator_names(),
        )
        acquired = list(result.get("acquired", []) or [])
        blocked = list(result.get("blocked", []) or [])
        if not acquired:
            if blocked:
                owner = str((blocked[0] or {}).get("operador", "") or "").strip() or "outro operador"
                raise ValueError(f"Operacao ocupada por {owner}.")
            raise ValueError("Nao foi possivel iniciar a operacao.")
        if not piece.get("inicio_producao"):
            piece["inicio_producao"] = self.desktop_main.now_iso()
        piece["interrupcao_peca_motivo"] = ""
        piece["interrupcao_peca_ts"] = ""
        piece["avaria_ativa"] = False
        piece["avaria_motivo"] = ""
        piece["avaria_fim_ts"] = ""
        self.operador_actions._mark_piece_ops_in_progress(piece, acquired, operator_name)
        self.desktop_main.atualizar_estado_peca(piece)
        piece["estado"] = "Em producao"
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            for op_name in acquired:
                log_fn(
                    evento="START_OP",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(piece.get("id", "") or ""),
                    ref_interna=piece.get("ref_interna", ""),
                    material=piece.get("material", ""),
                    espessura=piece.get("espessura", ""),
                    operacao=op_name,
                    operador=operator_name,
                    info=self._operator_info(operator_name, posto, "Operacao iniciada no Qt Operador"),
                )
        self._save_operator_state(enc)
        return {"operation": selected_op, "piece": piece, "blocked": blocked}

    def operator_finish_piece(
        self,
        enc_num: str,
        piece_id: str,
        operator_name: str,
        ok: Any,
        nok: Any,
        qual: Any,
        operation: str = "",
        posto: str = "Geral",
    ) -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        ctx = self.operator_piece_context(enc_num, piece_id)
        if ctx["has_open_avaria"]:
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de concluir a peca.")
        ok_val = self._parse_float(ok, 0)
        nok_val = self._parse_float(nok, 0)
        qual_val = self._parse_float(qual, 0)
        if min(ok_val, nok_val, qual_val) < 0:
            raise ValueError("Valores inv?lidos.")
        pending = list(ctx["pending_ops"])
        if not pending:
            raise ValueError("Nao existem operacoes pendentes nesta peca.")
        selected_op = self.desktop_main.normalize_operacao_nome(operation or pending[0])
        if selected_op not in pending:
            raise ValueError("A opera??o selecionada n?o est? pendente.")
        active_pending_ops = list(ctx.get("active_pending_ops", []) or [])
        if selected_op not in active_pending_ops:
            raise ValueError("Inicia primeiro a operacao antes de a concluir.")
        op_row = self._piece_operation_row(piece, selected_op)
        operation_limit = self._piece_operation_limit(piece, selected_op)
        if operation_limit <= 0:
            raise ValueError("Nao existe quantidade produzida no posto anterior para esta operacao.")
        ctx_done = self._parse_float(dict(ctx.get("operation_done", {}) or {}).get(selected_op, 0), 0)
        current_done = max(self._piece_operation_total(op_row, operation_limit), ctx_done)
        remaining_limit = round(max(0.0, operation_limit - current_done), 4)
        delta_val = round(ok_val + nok_val + qual_val, 4)
        if delta_val <= 0:
            raise ValueError("Indica pelo menos uma quantidade para concluir.")
        if delta_val > remaining_limit + 1e-9:
            raise ValueError(f"Quantidade acima do disponivel nesta operacao. Maximo restante: {remaining_limit:.1f}")
        existing_ok = self._parse_float((op_row or {}).get("qtd_ok", 0), 0)
        existing_nok = self._parse_float((op_row or {}).get("qtd_nok", 0), 0)
        existing_qual = self._parse_float((op_row or {}).get("qtd_qual", 0), 0)
        existing_total = round(existing_ok + existing_nok + existing_qual, 4)
        if current_done > existing_total + 1e-9:
            # Registos antigos podem ter o acumulado correto no contexto mas não na linha da operação.
            existing_ok = round(existing_ok + (current_done - existing_total), 4)
        final_ok = round(existing_ok + ok_val, 4)
        final_nok = round(existing_nok + nok_val, 4)
        final_qual = round(existing_qual + qual_val, 4)
        new_total = round(current_done + delta_val, 4)
        operation_complete = new_total >= operation_limit - 1e-9
        result = self.operador_actions._mysql_ops_finish(
            str(enc.get("numero", "") or ""),
            str(piece.get("id", "") or ""),
            [selected_op],
            operator_name,
            ok_val,
            nok_val,
            qual_val,
            valid_operators=self.operator_names(),
            complete=operation_complete,
        )
        blocked = list(result.get("blocked", []) or [])
        if blocked:
            blocked_state = str((blocked[0] or {}).get("estado", "") or "").strip()
            blocked_norm = self.desktop_main.norm_text(blocked_state)
            should_repair_stale_lock = ("concl" in blocked_norm) and (current_done < operation_limit - 1e-9)
            if should_repair_stale_lock:
                reset_fn = getattr(self.operador_actions, "_mysql_ops_reset_piece", None)
                acquire_fn = getattr(self.operador_actions, "_mysql_ops_acquire", None)
                if callable(reset_fn) and callable(acquire_fn):
                    reset_fn(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
                    reacquire = acquire_fn(
                        str(enc.get("numero", "") or ""),
                        str(piece.get("id", "") or ""),
                        [selected_op],
                        operator_name,
                        valid_operators=self.operator_names(),
                    )
                    reacquired = list((reacquire or {}).get("acquired", []) or [])
                    if selected_op in reacquired:
                        result = self.operador_actions._mysql_ops_finish(
                            str(enc.get("numero", "") or ""),
                            str(piece.get("id", "") or ""),
                            [selected_op],
                            operator_name,
                            ok_val,
                            nok_val,
                            qual_val,
                            valid_operators=self.operator_names(),
                            complete=operation_complete,
                        )
                        blocked = list(result.get("blocked", []) or [])
        if blocked:
            owner = str((blocked[0] or {}).get("operador", "") or "").strip() or "outro operador"
            state = str((blocked[0] or {}).get("estado", "") or "").strip()
            raise ValueError(f"Nao foi possivel concluir a operacao. Estado atual: {state or '-'} | Operador: {owner}")
        ts_fim = self.desktop_main.now_iso()
        if op_row is None:
            self.desktop_main.ensure_peca_operacoes(piece)
            op_row = self._piece_operation_row(piece, selected_op)
        if op_row is not None:
            op_row["qtd_ok"] = final_ok
            op_row["qtd_nok"] = final_nok
            op_row["qtd_qual"] = final_qual
            if not op_row.get("inicio"):
                op_row["inicio"] = ts_fim
            op_row["fim"] = ts_fim
            op_row["user"] = operator_name
            op_row["estado"] = "Concluida" if operation_complete else "Incompleta"
        if operation_complete:
            self.desktop_main.concluir_operacoes_peca(piece, [selected_op], user=operator_name)
        else:
            piece["operacoes_fluxo"] = self.desktop_main.ensure_peca_operacoes(piece)
            piece["Operacoes"] = " + ".join([x.get("nome", "") for x in piece.get("operacoes_fluxo", []) if x.get("nome")])
        final_row = self._piece_operation_row(piece, selected_op)
        if final_row is not None:
            final_row["qtd_ok"] = final_ok
            final_row["qtd_nok"] = final_nok
            final_row["qtd_qual"] = final_qual
            if not final_row.get("inicio"):
                final_row["inicio"] = ts_fim
            final_row["fim"] = ts_fim
            final_row["user"] = operator_name
            final_row["estado"] = "Concluida" if operation_complete else "Incompleta"
        piece["operacoes_fluxo"] = self.desktop_main.ensure_peca_operacoes(piece)
        piece["Operacoes"] = " + ".join([x.get("nome", "") for x in piece.get("operacoes_fluxo", []) if x.get("nome")])
        self._sync_piece_output_from_flow(piece)
        self.desktop_main.atualizar_estado_peca(piece)
        self.operador_actions._flush_piece_elapsed_minutes(piece, ts_fim)
        if piece.get("estado") == "Concluida":
            piece["fim_producao"] = ts_fim
        else:
            piece["fim_producao"] = ""
            piece["estado"] = "Em producao/Pausada"
        piece.setdefault("hist", []).append(
            {
                "ts": ts_fim,
                "user": operator_name,
                "acao": "Fim Peca" if piece.get("estado") == "Concluida" else "Registo Operacao",
                "operacoes": [selected_op],
                "ok": ok_val,
                "nok": nok_val,
                "qual": qual_val,
                "tempo_min": piece.get("tempo_producao_min", 0),
                "limite_operacao": operation_limit,
                "registo_acumulado": new_total,
            }
        )
        if nok_val > 0:
            self.ensure_data().setdefault("rejeitadas_hist", []).append(
                {
                    "data": self.desktop_main.now_iso(),
                    "operador": operator_name,
                    "encomenda": enc.get("numero", ""),
                    "material": piece.get("material", ""),
                    "espessura": piece.get("espessura", ""),
                    "ref_interna": piece.get("ref_interna", ""),
                    "ref_externa": piece.get("ref_externa", ""),
                    "nok": nok_val,
                }
            )
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="FINISH_OP",
                encomenda_numero=enc.get("numero", ""),
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operacao=selected_op,
                operador=operator_name,
                qtd_ok=ok_val,
                qtd_nok=nok_val,
                info=self._operator_info(operator_name, posto, "Operacao concluida no Qt Operador"),
            )
            if nok_val > 0:
                log_fn(
                    evento="SCRAP",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(piece.get("id", "") or ""),
                    ref_interna=piece.get("ref_interna", ""),
                    material=piece.get("material", ""),
                    espessura=piece.get("espessura", ""),
                    operador=operator_name,
                    qtd_nok=nok_val,
                    info=self._operator_info(operator_name, posto, "Registo NOK no Qt Operador"),
                )
        self._save_operator_state(enc)
        return {"operation": selected_op, "piece": piece}

    def operator_resume_piece(self, enc_num: str, piece_id: str, operator_name: str, posto: str = "Geral") -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        ctx = self.operator_piece_context(enc_num, piece_id)
        if ctx["has_open_avaria"]:
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de retomar a peca.")
        self.operador_actions._mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        motivo = str(piece.get("interrupcao_peca_motivo", "") or "").strip()
        piece.setdefault("hist", []).append({"ts": self.desktop_main.now_iso(), "user": operator_name, "acao": "Retomar Peca", "motivo": motivo})
        piece["interrupcao_peca_motivo"] = ""
        piece["interrupcao_peca_ts"] = ""
        piece["estado"] = "Em producao/Pausada"
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="RESUME_PIECE",
                encomenda_numero=enc.get("numero", ""),
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operador=operator_name,
                info=self._operator_info(operator_name, posto, f"Retoma de peca. Motivo anterior: {motivo or '-'}"),
            )
        self._save_operator_state(enc)
        return {"piece": piece}

    def operator_pause_piece(self, enc_num: str, piece_id: str, operator_name: str, motivo: str, posto: str = "Geral") -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        motivo = str(motivo or "").strip()
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        if not motivo:
            raise ValueError("Indica o motivo da interrupcao.")
        ctx = self.operator_piece_context(enc_num, piece_id)
        if ctx["has_open_avaria"]:
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de interromper a peca.")
        ts_pause = self.desktop_main.now_iso()
        piece["estado"] = "Interrompida"
        piece["interrupcao_peca_motivo"] = motivo
        piece["interrupcao_peca_ts"] = ts_pause
        self.operador_actions._flush_piece_elapsed_minutes(piece, ts_pause)
        fluxo = self.desktop_main.ensure_peca_operacoes(piece)
        for op in fluxo:
            if "concl" not in self.desktop_main.norm_text(op.get("estado", "")):
                op["estado"] = "Preparacao"
                op["fim"] = ""
        piece["operacoes_fluxo"] = fluxo
        piece.setdefault("hist", []).append({"ts": ts_pause, "user": operator_name, "acao": "Interromper Peca", "motivo": motivo})
        self.operador_actions._mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="PAUSE_PIECE",
                encomenda_numero=enc.get("numero", ""),
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operador=operator_name,
                info=self._operator_info(operator_name, posto, f"Interrupcao da peca. Motivo: {motivo}"),
            )
        self._save_operator_state(enc)
        return {"piece": piece}

    def operator_register_avaria(
        self,
        enc_num: str,
        piece_id: str,
        operator_name: str,
        motivo: str,
        posto: str = "Geral",
        group_id: str = "",
        ts_now: str = "",
    ) -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        motivo = str(motivo or "").strip() or "Avaria n?o especificada"
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        ctx = self.operator_piece_context(enc_num, piece_id)
        if ctx["has_open_avaria"] or bool(piece.get("avaria_ativa")):
            raise ValueError("Ja existe uma avaria aberta nesta peca.")
        ts_now = str(ts_now or self.desktop_main.now_iso()).strip() or self.desktop_main.now_iso()
        group_id = str(
            group_id
            or self.operador_actions._op_make_avaria_group_id(
                str(enc.get("numero", "") or ""),
                motivo,
                operator_name,
                ts_now=ts_now,
            )
        ).strip()
        piece["estado"] = "Avaria"
        piece["avaria_ativa"] = True
        piece["avaria_motivo"] = motivo
        piece["avaria_grupo_id"] = group_id
        piece["avaria_inicio_ts"] = ts_now
        piece["avaria_fim_ts"] = ""
        piece["interrupcao_peca_motivo"] = motivo
        piece["interrupcao_peca_ts"] = ts_now
        self.operador_actions._flush_piece_elapsed_minutes(piece, ts_now)
        fluxo = self.desktop_main.ensure_peca_operacoes(piece)
        for op in fluxo:
            if "concl" not in self.desktop_main.norm_text(op.get("estado", "")):
                op["estado"] = "Preparacao"
                op["fim"] = ""
        piece["operacoes_fluxo"] = fluxo
        piece.setdefault("hist", []).append({"ts": ts_now, "user": operator_name, "acao": "Registar Avaria", "motivo": motivo})
        self.operador_actions._mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="PARAGEM",
                encomenda_numero=enc.get("numero", ""),
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operador=operator_name,
                causa_paragem=motivo,
                info=self._operator_info(operator_name, posto, "Registo de avaria no Qt Operador"),
                created_at=ts_now,
                grupo_id=group_id,
            )
        self._save_operator_state(enc)
        return {"piece": piece, "avaria_group_key": group_id, "avaria_started_at": ts_now}

    def operator_close_avaria(self, enc_num: str, piece_id: str, operator_name: str, posto: str = "Geral") -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        avaria_index = self.operador_actions._op_open_avaria_index(self.ensure_data(), str(enc.get("numero", "") or ""))
        live_row = self.operador_actions._op_live_avaria_row_for_piece(avaria_index, piece)
        if live_row:
            self.operador_actions._op_sync_piece_live_avaria(piece, live_row)
        if not bool(piece.get("avaria_ativa")) and not live_row:
            raise ValueError("Nao existe avaria aberta nesta peca.")
        ts_now = self.desktop_main.now_iso()
        dur_segment = self.operador_actions._op_piece_current_avaria_minutes(
            self.ensure_data(),
            str(enc.get("numero", "") or ""),
            piece,
            live_row=live_row,
            ts_ref=ts_now,
        )
        group_key = self.operador_actions._op_avaria_group_key(live_row or {})
        if not group_key:
            group_key = str(piece.get("avaria_grupo_id", "") or "").strip() or self.operador_actions._op_piece_lookup_key(piece)
        motivo = str(piece.get("avaria_motivo", "") or piece.get("interrupcao_peca_motivo", "") or (live_row or {}).get("causa", "") or "").strip() or "Avaria n?o especificada"
        piece["avaria_ativa"] = False
        piece["avaria_fim_ts"] = ts_now
        piece["interrupcao_peca_motivo"] = ""
        piece["interrupcao_peca_ts"] = ""
        piece["avaria_motivo"] = ""
        piece["avaria_grupo_id"] = ""
        self.desktop_main.atualizar_estado_peca(piece)
        piece.setdefault("hist", []).append(
            {
                "ts": ts_now,
                "user": operator_name,
                "acao": "Fechar Avaria",
                "motivo": motivo,
                "inicio": str(piece.get("avaria_inicio_ts", "") or ""),
                "duracao_min": round(dur_segment, 2),
            }
        )
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="CLOSE_AVARIA",
                encomenda_numero=enc.get("numero", ""),
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operador=operator_name,
                causa_paragem=motivo,
                info=self._operator_info(operator_name, posto, f"Avaria fechada no Qt Operador. Motivo: {motivo}"),
            )
        self._save_operator_state(enc)
        return {
            "piece": piece,
            "duracao_avaria_min": round(dur_segment, 2),
            "avaria_group_key": group_key,
        }

    def operator_alert_chefia(self, enc_num: str, piece_id: str, operator_name: str, posto: str = "Geral") -> dict[str, Any]:
        enc, piece = self._find_piece(enc_num, piece_id)
        operator_name = str(operator_name or "").strip()
        posto = str(posto or "").strip() or "Geral"
        if not operator_name:
            raise ValueError("Seleciona o operador.")
        detalhe = f"Chefia solicitada para deslocacao imediata ao colaborador no posto {posto}."
        log_fn = getattr(self.desktop_main, "mysql_log_production_event", None)
        if callable(log_fn):
            log_fn(
                evento="POKE_CHEFIA",
                encomenda_numero=enc_num,
                peca_id=str(piece.get("id", "") or ""),
                ref_interna=piece.get("ref_interna", ""),
                material=piece.get("material", ""),
                espessura=piece.get("espessura", ""),
                operador=operator_name,
                info=self._operator_info(operator_name, posto, detalhe),
            )
        self.ensure_data().setdefault("chefia_alertas", []).append(
            {
                "created_at": self.desktop_main.now_iso(),
                "tipo": "POKE_CHEFIA",
                "encomenda_numero": enc_num,
                "peca_id": str(piece.get("id", "") or ""),
                "ref_interna": piece.get("ref_interna", ""),
                "material": piece.get("material", ""),
                "espessura": piece.get("espessura", ""),
                "operador": operator_name,
                "posto": posto,
                "mensagem": detalhe,
            }
        )
        self.ensure_data()["chefia_alertas"] = list(self.ensure_data().get("chefia_alertas", []) or [])[-200:]
        self._save(force=True)
        return {"message": detalhe}

    def operator_open_drawing(self, enc_num: str, piece_id: str) -> str:
        _, piece = self._find_piece(enc_num, piece_id)
        refs_db = self.ensure_data().get("orc_refs", {})
        ref_ext = str(piece.get("ref_externa", "") or "").strip()
        ref_info = refs_db.get(ref_ext, {}) if isinstance(refs_db, dict) and ref_ext else {}
        candidates = [
            piece.get("desenho"),
            piece.get("desenho_path"),
            ref_info.get("desenho"),
            ref_info.get("desenho_path"),
        ]
        for raw in candidates:
            path_txt = str(raw or "").strip()
            if not path_txt:
                continue
            path = self._resolve_file_reference(path_txt)
            if path is not None and path.exists():
                os.startfile(str(path))
                return str(path)
        raise ValueError("Esta peça não tem desenho associado ou o ficheiro não existe.")

    def _operator_pdf_text(self, value: Any) -> str:
        formatter = getattr(self.desktop_main, "pdf_normalize_text", None)
        if callable(formatter):
            try:
                return str(formatter(value) or "")
            except Exception:
                return str(value or "")
        return str(value or "")

    def _operator_label_palette(self) -> dict[str, Any]:
        from reportlab.lib import colors

        branding = self.branding_settings()
        primary_hex = str(branding.get("primary_color", "") or "#1F3C88").strip() or "#1F3C88"
        line_hex = _pdf_mix_hex(primary_hex, "#D7DEE8", 0.74)
        return {
            "primary": colors.HexColor(primary_hex),
            "primary_hex": primary_hex,
            "primary_dark": colors.HexColor(_pdf_mix_hex(primary_hex, "#000000", 0.18)),
            "primary_soft": colors.HexColor(_pdf_mix_hex(primary_hex, "#FFFFFF", 0.82)),
            "primary_soft_2": colors.HexColor(_pdf_mix_hex(primary_hex, "#FFFFFF", 0.92)),
            "ink": colors.HexColor(_pdf_mix_hex(primary_hex, "#1A1A1A", 0.72)),
            "muted": colors.HexColor("#667085"),
            "line": colors.HexColor(line_hex),
            "line_strong": colors.HexColor(_pdf_mix_hex(primary_hex, "#7A8699", 0.30)),
            "surface": colors.white,
            "surface_alt": colors.HexColor("#F8FAFC"),
            "success": colors.HexColor("#107569"),
            "danger": colors.HexColor("#B42318"),
            "warning": colors.HexColor("#B54708"),
        }

    def _operator_label_tmp_path(self, enc_num: str, variant: str) -> Path:
        safe_enc = "".join(ch if ch.isalnum() else "_" for ch in str(enc_num or "").strip()) or "operador"
        safe_variant = "".join(ch if ch.isalnum() else "_" for ch in str(variant or "").strip()) or "labels"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        return Path(tempfile.gettempdir()) / f"lugest_{safe_variant}_{safe_enc}_{stamp}.pdf"

    def _operator_operation_from_posto(self, posto: str = "") -> str:
        posto_norm = self.desktop_main.norm_text(self._legacy_workcenter_group_name(posto) or posto or "")
        if "laser" in posto_norm:
            return "Corte Laser"
        if "quin" in posto_norm:
            return "Quinagem"
        if "rosc" in posto_norm:
            return "Roscagem"
        if "sold" in posto_norm:
            return "Soldadura"
        if "embal" in posto_norm:
            return "Embalamento"
        if "mont" in posto_norm:
            return "Montagem"
        if "maquin" in posto_norm:
            return "Maquinacao"
        if "pint" in posto_norm:
            return "Pintura"
        if "furo" in posto_norm:
            return "Furo Manual"
        return ""

    def _operator_posto_for_operation(self, operation: str = "") -> str:
        normalized = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        op_norm = self.desktop_main.norm_text(normalized)
        keyword_map = [
            ("laser", "Laser"),
            ("quin", "Quinagem"),
            ("rosc", "Roscagem"),
            ("sold", "Soldadura"),
            ("embal", "Embalamento"),
            ("mont", "Montagem"),
            ("maquin", "Maquinacao"),
            ("pint", "Pintura"),
            ("furo", "Furo Manual"),
            ("exped", "Expedicao"),
        ]
        for token, fallback in keyword_map:
            if token not in op_norm:
                continue
            for posto in self.workcenter_group_options(operation=normalized):
                if token in self.desktop_main.norm_text(posto):
                    return posto
            return fallback
        return normalized or "-"

    def _operator_next_route(self, piece: dict[str, Any], source_posto: str = "Geral", source_operation: str = "") -> dict[str, str]:
        flow = []
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            name = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
            if name:
                flow.append(name)
        pending_ops = list(self.desktop_main.peca_operacoes_pendentes(piece))
        current_operation = self.desktop_main.normalize_operacao_nome(piece.get("operacao_atual", "")) or ""
        chosen_source_operation = self.desktop_main.normalize_operacao_nome(source_operation or self._operator_operation_from_posto(source_posto)) or current_operation
        available_pending_ops: list[str] = []
        for op_name in flow:
            limit = self._piece_operation_limit(piece, op_name)
            total_done = self._piece_operation_total(self._piece_operation_row(piece, op_name), limit)
            if limit > total_done + 1e-9:
                available_pending_ops.append(op_name)
        next_operation = available_pending_ops[0] if available_pending_ops else (pending_ops[0] if pending_ops else "")
        if not current_operation:
            current_operation = next_operation or (flow[-1] if flow else "")
        origin_posto = str(source_posto or "").strip()
        if not origin_posto or self.desktop_main.norm_text(origin_posto) == "geral":
            origin_posto = self._operator_posto_for_operation(chosen_source_operation or current_operation)
        next_posto = "Expedicao" if (not next_operation and flow) else self._operator_posto_for_operation(next_operation)
        return {
            "flow": " -> ".join(flow),
            "source_operation": chosen_source_operation or current_operation or "-",
            "current_operation": current_operation or "-",
            "source_posto": origin_posto or "-",
            "next_operation": next_operation or "Expedicao",
            "next_posto": next_posto or "Expedicao",
        }

    def _ensure_operator_piece_opp(self, piece: dict[str, Any]) -> bool:
        opp = str(piece.get("opp", "") or "").strip()
        if opp:
            return False
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if any(row is piece for row in list(self.desktop_main.encomenda_pecas(enc) or [])):
                piece["opp"] = self._next_order_opp_codigo(enc)
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self._order_of_code(enc, create=True)
                return bool(piece.get("opp"))
        piece["opp"] = str(self.desktop_main.next_opp_numero(self.ensure_data()) or "").strip()
        return bool(piece.get("opp"))

    def _operator_label_row(self, enc: dict[str, Any], piece: dict[str, Any], source_posto: str = "Geral") -> dict[str, Any]:
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn) and cliente_codigo:
            try:
                cliente_obj = find_cliente_fn(self.ensure_data(), cliente_codigo) or {}
            except Exception:
                cliente_obj = {}
        route = self._operator_next_route(piece, source_posto=source_posto)
        descricao = str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip()
        qtd = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        return {
            "piece_id": str(piece.get("id", "") or "").strip(),
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "cliente_label": f"{cliente_codigo} - {str(cliente_obj.get('nome', '') or '').strip()}".strip(" -"),
            "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
            "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
            "descricao": descricao,
            "material": str(piece.get("material", "") or "").strip(),
            "espessura": str(piece.get("espessura", "") or "").strip(),
            "quantidade": qtd,
            "quantidade_txt": self._fmt(qtd),
            "estado": str(piece.get("estado", "") or "").strip(),
            "operacao_atual": route.get("current_operation", "-"),
            "operacao_origem": route.get("source_operation", "-"),
            "posto_origem": route.get("source_posto", "-"),
            "proxima_operacao": route.get("next_operation", "Expedicao"),
            "proximo_posto": route.get("next_posto", "Expedicao"),
            "fluxo": route.get("flow", ""),
        }

    def _operator_label_rows_for_order(self, enc_num: str, source_posto: str = "Geral") -> tuple[dict[str, Any], list[dict[str, Any]]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        changed = False
        rows: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            changed = self._ensure_operator_piece_opp(piece) or changed
            rows.append(self._operator_label_row(enc, piece, source_posto=source_posto))
        if changed:
            self._save(force=True)
        rows.sort(
            key=lambda row: (
                str(row.get("proximo_posto", "") or ""),
                str(row.get("ref_interna", "") or ""),
                str(row.get("opp", "") or ""),
            )
        )
        return enc, rows

    def _operator_selected_label_rows(self, enc_num: str, piece_ids: list[str] | tuple[str, ...], source_posto: str = "Geral") -> tuple[dict[str, Any], list[dict[str, Any]]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        selected_ids = [str(piece_id or "").strip() for piece_id in list(piece_ids or []) if str(piece_id or "").strip()]
        if not selected_ids:
            raise ValueError("Seleciona pelo menos uma referencia para imprimir etiquetas.")
        selected_set = set(selected_ids)
        changed = False
        rows: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            piece_id = str(piece.get("id", "") or "").strip()
            if piece_id not in selected_set:
                continue
            changed = self._ensure_operator_piece_opp(piece) or changed
            rows.append(self._operator_label_row(enc, piece, source_posto=source_posto))
        if not rows:
            raise ValueError("As referencias selecionadas ja nao existem nesta encomenda.")
        order_map = {piece_id: index for index, piece_id in enumerate(selected_ids)}
        rows.sort(key=lambda row: (order_map.get(str(row.get("piece_id", "") or "").strip(), 999999), str(row.get("ref_interna", "") or "")))
        if changed:
            self._save(force=True)
        return enc, rows

    def _draw_operator_logo(self, canvas_obj, logo_path: Path | None, x: float, y: float, width: float, height: float) -> None:
        from reportlab.lib.utils import ImageReader

        if not logo_path or not logo_path.exists():
            return
        try:
            canvas_obj.drawImage(ImageReader(str(logo_path)), x, y, width=width, height=height, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    def _draw_operator_logo_plate(
        self,
        canvas_obj,
        palette: dict[str, Any],
        logo_path: Path | None,
        x: float,
        y: float,
        width: float,
        height: float,
        *,
        radius: float = 10,
        padding_x: float = 5,
        padding_y: float = 4,
        line_width: float = 0.9,
    ) -> None:
        canvas_obj.saveState()
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(line_width)
        canvas_obj.roundRect(x, y, width, height, radius, stroke=1, fill=1)
        canvas_obj.restoreState()
        self._draw_operator_logo(
            canvas_obj,
            logo_path,
            x + padding_x,
            y + padding_y,
            max(10.0, width - (padding_x * 2)),
            max(8.0, height - (padding_y * 2)),
        )

    def _draw_operator_unit_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        row: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        from reportlab.graphics.barcode import code128

        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        outer_x = 8
        outer_y = 7
        outer_w = page_width - (outer_x * 2)
        outer_h = page_height - (outer_y * 2)
        banner_h = 30
        right_chip_w = 86
        right_meta_w = 104
        body_left_w = outer_w - right_meta_w - 28
        logo_box_w = 42
        logo_gap = 8
        banner_x = outer_x + 9 + logo_box_w + logo_gap
        banner_w = outer_w - (banner_x - outer_x)
        logo_box_x = outer_x + 9
        logo_box_y = outer_y + outer_h - banner_h + 5
        chip_x = outer_x + outer_w - right_chip_w - 8
        chip_y = outer_y + outer_h - banner_h + 5

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 11, stroke=1, fill=1)

        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(banner_x, outer_y + outer_h - banner_h, banner_w, banner_h, 11, stroke=0, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            logo_box_x,
            logo_box_y,
            42,
            20,
            radius=6,
            padding_x=3,
            padding_y=2,
            line_width=0.8,
        )

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(chip_x, chip_y, right_chip_w, 20, 7, stroke=0, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 6.2)
        canvas_obj.drawString(chip_x + 6, chip_y + 12, self._operator_pdf_text("Seguinte"))
        next_font = _pdf_fit_font_size(row.get("proximo_posto", "-"), bold_font, right_chip_w - 12, 8.6, 6.7)
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, next_font)
        canvas_obj.drawString(
            chip_x + 6,
            chip_y + 4.5,
            self._operator_pdf_text(_pdf_clip_text(row.get("proximo_posto", "-"), right_chip_w - 10, bold_font, next_font)),
        )

        title_x_left = banner_x + 10
        title_x_right = chip_x - 10
        title_width = max(50.0, title_x_right - title_x_left)
        title_font = _pdf_fit_font_size("Etiqueta OPP", bold_font, title_width, 11.6, 8.8)
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_x_left + (title_width / 2.0), chip_y + 12.8, self._operator_pdf_text("Etiqueta OPP"))
        subtitle = f"{row.get('encomenda', '-') or '-'} | {row.get('cliente', '-') or '-'}"
        subtitle_font = _pdf_fit_font_size(subtitle, regular_font, title_width, 6.2, 5.4)
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_x_left + (title_width / 2.0),
            chip_y + 4.2,
            self._operator_pdf_text(_pdf_clip_text(subtitle, title_width, regular_font, subtitle_font)),
        )

        meta_x = outer_x + outer_w - right_meta_w - 10
        meta_y = outer_y + 47
        meta_h = 42
        canvas_obj.setFillColor(palette["primary_soft_2"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(meta_x, meta_y, right_meta_w, meta_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 6.2)
        canvas_obj.drawString(meta_x + 8, meta_y + meta_h - 10, self._operator_pdf_text("OPP"))
        canvas_obj.drawString(meta_x + 8, meta_y + meta_h - 25, self._operator_pdf_text("Qtd"))
        canvas_obj.setFillColor(palette["ink"])
        opp_font = _pdf_fit_font_size(row.get("opp", "-"), bold_font, right_meta_w - 16, 10.8, 7.2)
        canvas_obj.setFont(bold_font, opp_font)
        canvas_obj.drawString(meta_x + 8, meta_y + meta_h - 19, self._operator_pdf_text(_pdf_clip_text(row.get("opp", "-"), right_meta_w - 16, bold_font, opp_font)))
        qty_text = row.get("quantidade_txt", "0")
        qty_font = _pdf_fit_font_size(qty_text, bold_font, right_meta_w - 16, 10.6, 7.6)
        canvas_obj.setFont(bold_font, qty_font)
        canvas_obj.drawString(meta_x + 8, meta_y + 8, self._operator_pdf_text(_pdf_clip_text(qty_text, right_meta_w - 16, bold_font, qty_font)))

        text_x = outer_x + 10
        top_y = outer_y + outer_h - 42
        ref_font = _pdf_fit_font_size(row.get("ref_interna", "-"), bold_font, body_left_w, 11.5, 8.4)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, ref_font)
        canvas_obj.drawString(text_x, top_y, self._operator_pdf_text(_pdf_clip_text(row.get("ref_interna", "-"), body_left_w, bold_font, ref_font)))

        description = row.get("descricao", "") or row.get("ref_externa", "") or "-"
        desc_lines = _pdf_wrap_text(description, regular_font, 7.0, body_left_w, max_lines=2) or ["-"]
        desc_line_gap = 7.6
        desc_start_y = top_y - 10
        canvas_obj.setFont(regular_font, 7.0)
        canvas_obj.setFillColor(palette["muted"])
        for line_index, line in enumerate(desc_lines[:2]):
            canvas_obj.drawString(text_x, desc_start_y - (line_index * desc_line_gap), self._operator_pdf_text(line))

        meta_lines = [
            f"OF {row.get('of', '-') or '-'} | Ref. Ext. {row.get('ref_externa', '-') or '-'}",
            f"{row.get('material', '-') or '-'} {row.get('espessura', '-') or '-'} mm | Estado {row.get('estado', '-') or '-'}",
        ]
        meta_line_gap = 6.6
        desc_last_y = desc_start_y - ((max(1, len(desc_lines[:2])) - 1) * desc_line_gap)
        meta_start_y = desc_last_y - 10.5
        for line_index, line in enumerate(meta_lines):
            meta_font = _pdf_fit_font_size(line, regular_font, body_left_w - 2, 5.5 if line_index == 0 else 5.3, 4.8)
            canvas_obj.setFont(regular_font, meta_font)
            canvas_obj.drawString(
                text_x,
                meta_start_y - (line_index * meta_line_gap),
                self._operator_pdf_text(_pdf_clip_text(line, body_left_w - 2, regular_font, meta_font)),
            )

        route_y = outer_y + 24
        route_h = 16
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(outer_x + 8, route_y, outer_w - 16, route_h, 7, stroke=1, fill=1)
        route_left = f"Origem: {row.get('posto_origem', '-') or '-'}"
        route_right = f"Seguinte: {row.get('proximo_posto', '-') or '-'}"
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, 6.7)
        canvas_obj.drawString(outer_x + 14, route_y + 5.1, self._operator_pdf_text(_pdf_clip_text(route_left, outer_w - 170, bold_font, 6.7)))
        canvas_obj.drawRightString(
            outer_x + outer_w - 14,
            route_y + 5.1,
            self._operator_pdf_text(_pdf_clip_text(route_right, 150, bold_font, 6.7)),
        )

        barcode_value = str(row.get("opp", "") or row.get("piece_id", "") or "-").strip() or "-"
        barcode = code128.Code128(barcode_value, barHeight=12, barWidth=0.46)
        barcode.drawOn(canvas_obj, outer_x + 10, outer_y + 8)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.8)
        canvas_obj.drawRightString(outer_x + outer_w - 10, outer_y + 11, self._operator_pdf_text(printed_at))

    def _draw_operator_pallet_page(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        rows: list[dict[str, Any]],
        group_rows: list[dict[str, Any]],
        group_name: str,
        palette: dict[str, Any],
        logo_path: Path | None,
        source_posto: str,
        printed_at: str,
        page_number: int,
        total_pages: int,
    ) -> None:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        margin = 26
        header_h = 80
        card_gap = 10
        page_inner_w = page_width - (margin * 2)
        banner_x = margin
        banner_y = page_height - margin - header_h
        small_card_w = 114
        small_card_h = 28
        card_grid_gap = 8
        card_grid_v_gap = 6
        card_grid_pad = 10
        card_group_w = (small_card_w * 2) + card_grid_gap + (card_grid_pad * 2)
        logo_box_w = 82
        logo_gap = 12
        logo_box_x = banner_x
        banner_x = margin + logo_box_w + logo_gap
        page_inner_w = page_width - margin - banner_x
        group_x = banner_x + page_inner_w - card_group_w - 12
        logo_box_y = banner_y + 17
        title_left = banner_x + 18
        title_right = group_x - 12
        title_w = max(120.0, title_right - title_left)

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(banner_x, banner_y, page_inner_w, header_h, 16, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(banner_x, banner_y, page_inner_w, header_h, 16, stroke=0, fill=1)

        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            logo_box_x,
            logo_box_y,
            82,
            46,
            radius=12,
            padding_x=6,
            padding_y=6,
            line_width=0.9,
        )

        title = "Etiqueta de Palete"
        subtitle = f"Destino {group_name or '-'} | {rows[0].get('encomenda', '-') or '-'}"
        title_font = _pdf_fit_font_size(title, bold_font, title_w, 22.4, 16.6)
        subtitle_font = _pdf_fit_font_size(subtitle, regular_font, title_w, 9.4, 7.0)
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 50, self._operator_pdf_text(title))
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_left + (title_w / 2.0),
            banner_y + 32,
            self._operator_pdf_text(_pdf_clip_text(subtitle, title_w, regular_font, subtitle_font)),
        )

        header_cards = [
            ("Documento", f"PLT-{rows[0].get('encomenda', '-') or '-'}"),
            ("Pagina", f"{page_number}/{total_pages}"),
            ("Origem", source_posto or "-"),
            ("Impresso", printed_at[:16]),
        ]
        for index, (label, value) in enumerate(header_cards):
            row_idx = index // 2
            col_idx = index % 2
            box_x = group_x + card_grid_pad + (col_idx * (small_card_w + card_grid_gap))
            box_y = banner_y + header_h - card_grid_pad - small_card_h - (row_idx * (small_card_h + card_grid_v_gap))
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.roundRect(box_x, box_y, small_card_w, small_card_h, 9, stroke=0, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 6.5)
            canvas_obj.drawString(box_x + 8, box_y + small_card_h - 10, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, small_card_w - 16, 10.5, 7.0)
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(box_x + 8, box_y + 8, self._operator_pdf_text(_pdf_clip_text(value, small_card_w - 16, bold_font, value_font)))

        cards_y = page_height - margin - header_h - 74
        top_cards = [
            ("Cliente", rows[0].get("cliente_label", "-") or "-", 250),
            ("Resumo", f"Refs {len(group_rows)} | OPP {len(group_rows)} | Qtd {self._fmt(sum(float(row.get('quantidade', 0) or 0) for row in group_rows))}", 250),
            (
                "Operacao Seguinte",
                ", ".join(list(dict.fromkeys(str(row.get("proxima_operacao", "") or "").strip() for row in group_rows if str(row.get("proxima_operacao", "") or "").strip()))[:3]) or "-",
                page_inner_w - 520 - (card_gap * 2),
            ),
        ]
        card_x = margin
        for title_txt, body_txt, card_w in top_cards:
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(card_x, cards_y, card_w, 52, 12, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 7.0)
            canvas_obj.drawString(card_x + 12, cards_y + 36, self._operator_pdf_text(title_txt))
            body_font = _pdf_fit_font_size(body_txt, bold_font, card_w - 24, 10.4, 7.2)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(bold_font, body_font)
            body_lines = _pdf_wrap_text(body_txt, bold_font, body_font, card_w - 24, max_lines=2) or ["-"]
            for line_index, line in enumerate(body_lines):
                canvas_obj.drawString(card_x + 12, cards_y + 22 - (line_index * 11), self._operator_pdf_text(line))
            card_x += card_w + card_gap

        destination_y = cards_y - 64
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.roundRect(margin, destination_y, page_inner_w, 46, 14, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(margin + 16, destination_y + 29, self._operator_pdf_text("Destino desta palete"))
        dest_font = _pdf_fit_font_size(group_name or "-", bold_font, page_inner_w - 32, 23.0, 16.0)
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, dest_font)
        canvas_obj.drawString(margin + 16, destination_y + 11, self._operator_pdf_text(_pdf_clip_text(group_name or "-", page_inner_w - 32, bold_font, dest_font)))

        table_y = destination_y - 28
        header_y = table_y
        row_h = 22
        columns = [
            ("Ref. Int.", 112),
            ("Ref. Ext.", 112),
            ("Descricao", 162),
            ("Material", 72),
            ("Esp.", 42),
            ("Qtd", 54),
            ("OPP", 92),
            ("Seguinte", page_inner_w - 646),
        ]
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(margin, header_y, page_inner_w, 20, 9, stroke=0, fill=1)
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setFont(bold_font, 8.0)
        x_cursor = margin
        for label, width in columns:
            canvas_obj.drawString(x_cursor + 8, header_y + 6, self._operator_pdf_text(label))
            x_cursor += width

        body_top = header_y - 4
        for row_index, row in enumerate(rows):
            row_y = body_top - ((row_index + 1) * row_h)
            canvas_obj.setFillColor(palette["surface"] if row_index % 2 == 0 else palette["surface_alt"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(margin, row_y, page_inner_w, row_h - 2, 8, stroke=1, fill=1)
            values = [
                row.get("ref_interna", "-"),
                row.get("ref_externa", "-"),
                row.get("descricao", "") or "-",
                row.get("material", "-"),
                row.get("espessura", "-"),
                row.get("quantidade_txt", "0"),
                row.get("opp", "-"),
                row.get("proximo_posto", "-"),
            ]
            x_cursor = margin
            for column_index, ((_, width), value) in enumerate(zip(columns, values)):
                is_emphasis = column_index in (0, 6, 7)
                font_name = bold_font if is_emphasis else regular_font
                if column_index in (0, 1):
                    font_size = 7.0
                elif column_index == 2:
                    font_size = 6.8
                else:
                    font_size = 7.4
                clipped = _pdf_clip_text(value, width - 16, font_name, font_size)
                canvas_obj.setFillColor(palette["ink"] if is_emphasis else palette["muted"])
                canvas_obj.setFont(font_name, font_size)
                if column_index in (4, 5):
                    canvas_obj.drawRightString(x_cursor + width - 8, row_y + 7, self._operator_pdf_text(clipped))
                else:
                    canvas_obj.drawString(x_cursor + 8, row_y + 7, self._operator_pdf_text(clipped))
                x_cursor += width

        footer_y = 20
        footer_left = f"Origem {source_posto or '-'} | Encomenda {rows[0].get('encomenda', '-') or '-'}"
        footer_right = f"LUGEST | {printed_at}"
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(margin, footer_y, self._operator_pdf_text(footer_left))
        canvas_obj.drawRightString(page_width - margin, footer_y, self._operator_pdf_text(footer_right))

    def operator_label_rows(self, enc_num: str, source_posto: str = "Geral") -> dict[str, Any]:
        enc, rows = self._operator_label_rows_for_order(enc_num, source_posto=source_posto)
        return {
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": str(enc.get("cliente", "") or "").strip(),
            "source_posto": str(source_posto or "").strip() or "Geral",
            "rows": rows,
        }

    def operator_unit_labels_pdf(
        self,
        enc_num: str,
        piece_ids: list[str] | tuple[str, ...],
        source_posto: str = "Geral",
        output_path: str | Path | None = None,
    ) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        _enc, rows = self._operator_selected_label_rows(enc_num, piece_ids, source_posto=source_posto)
        target = Path(output_path) if output_path else self._operator_label_tmp_path(enc_num, "operator_unit_labels")
        target.parent.mkdir(parents=True, exist_ok=True)
        width, height = (110 * mm, 50 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(width, height))
        for index, row in enumerate(rows):
            if index:
                canvas_obj.showPage()
            self._draw_operator_unit_label(canvas_obj, width, height, row, palette, logo_path, printed_at)
        canvas_obj.save()
        return target

    def operator_pallet_labels_pdf(
        self,
        enc_num: str,
        piece_ids: list[str] | tuple[str, ...],
        source_posto: str = "Geral",
        output_path: str | Path | None = None,
    ) -> Path:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        _enc, selected_rows = self._operator_selected_label_rows(enc_num, piece_ids, source_posto=source_posto)
        groups: dict[str, list[dict[str, Any]]] = {}
        for row in selected_rows:
            key = str(row.get("proximo_posto", "") or "Expedicao").strip() or "Expedicao"
            groups.setdefault(key, []).append(row)
        ordered_groups = sorted(groups.items(), key=lambda item: item[0].lower())
        rows_per_page = 13
        pages: list[tuple[str, list[dict[str, Any]], list[dict[str, Any]]]] = []
        for group_name, group_rows in ordered_groups:
            group_rows.sort(key=lambda row: (str(row.get("ref_interna", "") or ""), str(row.get("opp", "") or "")))
            for start in range(0, len(group_rows), rows_per_page):
                pages.append((group_name, group_rows[start : start + rows_per_page], group_rows))
        if not pages:
            raise ValueError("Sem referencias para gerar a etiqueta de palete.")
        target = Path(output_path) if output_path else self._operator_label_tmp_path(enc_num, "operator_pallet_labels")
        target.parent.mkdir(parents=True, exist_ok=True)
        page_width, page_height = landscape(A4)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(page_width, page_height))
        total_pages = len(pages)
        for page_index, (group_name, page_rows, group_rows) in enumerate(pages, start=1):
            if page_index > 1:
                canvas_obj.showPage()
            self._draw_operator_pallet_page(
                canvas_obj,
                page_width,
                page_height,
                page_rows,
                group_rows,
                group_name,
                palette,
                logo_path,
                str(source_posto or "").strip() or "Geral",
                printed_at,
                page_index,
                total_pages,
            )
        canvas_obj.save()
        return target

    def _preferred_supplier_for_product(self, produto_codigo: str) -> dict[str, str]:
        code = str(produto_codigo or "").strip()
        if not code:
            return {"fornecedor_id": "", "fornecedor": "", "contacto": "", "origem": ""}
        for note in reversed(list(self.ensure_data().get("notas_encomenda", []) or [])):
            note_supplier = str(note.get("fornecedor", "") or "").strip()
            note_contact = str(note.get("contacto", "") or "").strip()
            note_number = str(note.get("numero", "") or "").strip()
            for line in list(note.get("linhas", []) or []):
                if self.desktop_main.origem_is_materia(line.get("origem", "")):
                    continue
                if str(line.get("ref", "") or "").strip() != code:
                    continue
                raw_supplier = str(line.get("fornecedor_linha", "") or note_supplier or "").strip()
                if not raw_supplier:
                    continue
                supplier_id, supplier_text, supplier_contact = self._resolve_supplier(raw_supplier)
                return {
                    "fornecedor_id": supplier_id,
                    "fornecedor": supplier_text or raw_supplier,
                    "contacto": supplier_contact or note_contact,
                    "origem": note_number,
                }
        return {"fornecedor_id": "", "fornecedor": "", "contacto": "", "origem": ""}

    def _order_montagem_shortages(self, enc: dict[str, Any]) -> list[dict[str, Any]]:
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(self.ensure_data().get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        shortages: list[dict[str, Any]] = []
        for item in list((enc or {}).get("montagem_itens", []) or []):
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            consumed = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            pending = round(max(0.0, plan - consumed), 2)
            if pending <= 1e-9:
                continue
            if self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                material_record = self.material_by_id(stock_id) if stock_id else None
                if material_record is not None:
                    available = max(
                        0.0,
                        self._parse_float(material_record.get("quantidade", 0), 0)
                        - self._parse_float(material_record.get("reservado", 0), 0),
                    )
                    material_txt = str(material_record.get("material", item.get("material", "")) or "").strip()
                    esp_txt = str(material_record.get("espessura", item.get("espessura", "")) or "").strip()
                    unit_price = self._parse_float(material_record.get("preco_unid", material_record.get("p_compra", item.get("preco_unit", 0))), 0)
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    candidates = self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []
                    available = round(sum(self._parse_float(row.get("disponivel", 0), 0) for row in candidates), 2)
                    unit_price = self._parse_float(item.get("preco_unit", item.get("price_base_value", 0)), 0)
                missing = round(max(0.0, pending - available), 2)
                if missing > 1e-9:
                    item_key = self._montagem_item_key(item)
                    shortages.append(
                        {
                            "kind": "material",
                            "item_key": item_key,
                            "produto_codigo": "",
                            "descricao": str(item.get("descricao", "") or material_txt or "").strip(),
                            "produto_unid": str(item.get("produto_unid", "") or "UN").strip() or "UN",
                            "qtd_pendente": pending,
                            "qtd_disponivel": round(available, 2),
                            "qtd_em_falta": missing,
                            "produto_encontrado": material_record is not None or available > 0,
                            "preco_unit": round(unit_price, 4),
                            "material": material_txt,
                            "espessura": esp_txt,
                            "dimensao": str(item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                            "stock_material_id": stock_id,
                            "fornecedor_id": "",
                            "fornecedor_sugerido": "",
                            "fornecedor_contacto": "",
                            "fornecedor_origem": "",
                        }
                    )
                continue
            if item_type != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            product = product_map.get(code)
            supplier_meta = self._preferred_supplier_for_product(code)
            available = round(self._parse_float((product or {}).get("qty", 0), 0), 2)
            missing = round(max(0.0, pending - available), 2)
            if product is None or missing > 1e-9:
                unit_price = round(
                    self._parse_float(
                        self.desktop_main.produto_preco_unitario(product or {}) if product is not None else item.get("preco_unit", 0),
                        0,
                    ),
                    4,
                )
                shortages.append(
                    {
                        "produto_codigo": code,
                        "kind": "product",
                        "item_key": self._montagem_item_key(item),
                        "descricao": str(item.get("descricao", "") or (product or {}).get("descricao", "") or "").strip(),
                        "produto_unid": str((product or {}).get("unid", "") or item.get("produto_unid", "") or "UN").strip() or "UN",
                        "qtd_pendente": pending,
                        "qtd_disponivel": available,
                        "qtd_em_falta": missing if product is not None else pending,
                        "produto_encontrado": product is not None,
                        "preco_unit": unit_price,
                        "fornecedor_id": str(supplier_meta.get("fornecedor_id", "") or "").strip(),
                        "fornecedor_sugerido": str(supplier_meta.get("fornecedor", "") or "").strip(),
                        "fornecedor_contacto": str(supplier_meta.get("contacto", "") or "").strip(),
                        "fornecedor_origem": str(supplier_meta.get("origem", "") or "").strip(),
                    }
                )
        shortages.sort(key=lambda row: (-self._parse_float(row.get("qtd_em_falta", 0), 0), str(row.get("produto_codigo", "") or row.get("descricao", "") or "")))
        return shortages

    def _montagem_item_is_raw_material(self, item: dict[str, Any] | None) -> bool:
        row = dict(item or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if str(row.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(row.get("stock_material_id", "") or "").strip():
            return True
        return False

    def _montagem_item_key(self, item: dict[str, Any] | None) -> str:
        row = dict(item or {})
        for key in ("linha_ordem", "grupo_uuid", "stock_material_id", "produto_codigo"):
            value = str(row.get(key, "") or "").strip()
            if value:
                return f"{key}:{value}"
        parts = [
            str(row.get("tipo_item", "") or "").strip(),
            str(row.get("descricao", "") or "").strip(),
            str(row.get("material", "") or "").strip(),
            str(row.get("espessura", "") or "").strip(),
        ]
        return "raw:" + "|".join(parts)

    def montagem_purchase_needs(self, order_numbers: list[str] | None = None) -> list[dict[str, Any]]:
        selected = {str(value or "").strip() for value in list(order_numbers or []) if str(value or "").strip()}
        data = self.ensure_data()
        grouped: dict[str, dict[str, Any]] = {}
        for enc in list(data.get("encomendas", []) or []):
            numero = str(enc.get("numero", "") or "").strip()
            if selected and numero not in selected:
                continue
            shortages = self._order_montagem_shortages(enc)
            if not shortages:
                continue
            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = next(
                (
                    str(row.get("nome", "") or "").strip()
                    for row in list(data.get("clientes", []) or [])
                    if str(row.get("codigo", "") or "").strip() == client_code
                ),
                "",
            )
            client_label = " - ".join([value for value in (client_code, client_name) if value]).strip(" -")
            delivery_date = str(enc.get("data_entrega", "") or "").strip()
            for shortage in shortages:
                kind = str(shortage.get("kind", "product") or "product").strip()
                code = str(shortage.get("produto_codigo", "") or "").strip()
                key = f"{kind}:{code or shortage.get('stock_material_id', '') or shortage.get('material', '')}|{shortage.get('espessura', '')}|{shortage.get('descricao', '')}"
                entry = grouped.setdefault(
                    key,
                    {
                        "kind": kind,
                        "produto_codigo": code,
                        "descricao": str(shortage.get("descricao", "") or "").strip(),
                        "produto_unid": str(shortage.get("produto_unid", "") or "UN").strip() or "UN",
                        "preco_unit": round(self._parse_float(shortage.get("preco_unit", 0), 0), 4),
                        "material": str(shortage.get("material", "") or "").strip(),
                        "espessura": str(shortage.get("espessura", "") or "").strip(),
                        "dimensao": str(shortage.get("dimensao", "") or "").strip(),
                        "stock_material_id": str(shortage.get("stock_material_id", "") or "").strip(),
                        "qtd_em_falta": 0.0,
                        "produto_encontrado": bool(shortage.get("produto_encontrado")),
                        "fornecedor_id": str(shortage.get("fornecedor_id", "") or "").strip(),
                        "fornecedor": str(shortage.get("fornecedor_sugerido", "") or "").strip(),
                        "fornecedor_contacto": str(shortage.get("fornecedor_contacto", "") or "").strip(),
                        "fornecedor_origem": str(shortage.get("fornecedor_origem", "") or "").strip(),
                        "encomendas": [],
                        "clientes": [],
                        "_datas_entrega": [],
                    },
                )
                entry["qtd_em_falta"] = round(
                    self._parse_float(entry.get("qtd_em_falta", 0), 0) + self._parse_float(shortage.get("qtd_em_falta", 0), 0),
                    2,
                )
                if not str(entry.get("descricao", "") or "").strip():
                    entry["descricao"] = str(shortage.get("descricao", "") or "").strip()
                if not str(entry.get("fornecedor", "") or "").strip() and str(shortage.get("fornecedor_sugerido", "") or "").strip():
                    entry["fornecedor_id"] = str(shortage.get("fornecedor_id", "") or "").strip()
                    entry["fornecedor"] = str(shortage.get("fornecedor_sugerido", "") or "").strip()
                    entry["fornecedor_contacto"] = str(shortage.get("fornecedor_contacto", "") or "").strip()
                    entry["fornecedor_origem"] = str(shortage.get("fornecedor_origem", "") or "").strip()
                if numero and numero not in entry["encomendas"]:
                    entry["encomendas"].append(numero)
                if client_label and client_label not in entry["clientes"]:
                    entry["clientes"].append(client_label)
                if delivery_date and delivery_date not in entry["_datas_entrega"]:
                    entry["_datas_entrega"].append(delivery_date)
        rows = list(grouped.values())
        for row in rows:
            delivery_dates = sorted(str(value or "").strip() for value in list(row.pop("_datas_entrega", []) or []) if str(value or "").strip())
            row["data_entrega"] = delivery_dates[0] if delivery_dates else ""
        rows.sort(
            key=lambda row: (
                not bool(str(row.get("fornecedor", "") or "").strip()),
                str(row.get("data_entrega", "") or "9999-99-99"),
                -self._parse_float(row.get("qtd_em_falta", 0), 0),
                str(row.get("produto_codigo", "") or row.get("descricao", "") or ""),
            )
        )
        return rows

    def ne_create_from_montagem_shortages(self, order_numbers: list[str] | None = None) -> dict[str, Any]:
        needs = list(self.montagem_purchase_needs(order_numbers))
        if not needs:
            raise ValueError("Nao existem faltas de stock de montagem para gerar nota.")
        all_orders = sorted(
            {
                str(order or "").strip()
                for need in needs
                for order in list(need.get("encomendas", []) or [])
                if str(order or "").strip()
            }
        )
        delivery_dates = sorted(
            {
                str(need.get("data_entrega", "") or "").strip()
                for need in needs
                if str(need.get("data_entrega", "") or "").strip()
            }
        )
        unique_suppliers = []
        missing_supplier = []
        for need in needs:
            supplier_txt = str(need.get("fornecedor", "") or "").strip()
            if supplier_txt and supplier_txt.lower() not in [value.lower() for value in unique_suppliers]:
                unique_suppliers.append(supplier_txt)
            if not supplier_txt:
                missing_supplier.append(str(need.get("produto_codigo", "") or need.get("descricao", "") or "").strip())
        supplier_id = ""
        supplier_text = ""
        supplier_contact = ""
        if len(unique_suppliers) == 1:
            supplier_id, supplier_text, supplier_contact = self._resolve_supplier(unique_suppliers[0])
        obs_parts = ["Reposicao automatica de montagem"]
        if all_orders:
            obs_parts.append("Encomendas: " + ", ".join(all_orders))
        if missing_supplier:
            obs_parts.append("Fornecedor por validar: " + ", ".join(sorted(set(item for item in missing_supplier if item))))
        note = self.ne_save(
            {
                "fornecedor": supplier_text,
                "fornecedor_id": supplier_id,
                "contacto": supplier_contact,
                "data_entrega": delivery_dates[0] if delivery_dates else "",
                "obs": " | ".join(obs_parts),
                "lines": [
                    (
                        {
                            "ref": str(need.get("stock_material_id", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or need.get("material", "") or "").strip(),
                            "fornecedor_linha": str(need.get("fornecedor", "") or "").strip(),
                            "origem": "Materia-prima",
                            "qtd": round(self._parse_float(need.get("qtd_em_falta", 0), 0), 2),
                            "unid": str(need.get("produto_unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco_unit", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "material": str(need.get("material", "") or "").strip(),
                            "espessura": str(need.get("espessura", "") or "").strip(),
                            "dimensao": str(need.get("dimensao", "") or "").strip(),
                            "dimensoes": str(need.get("dimensao", "") or "").strip(),
                        }
                        if str(need.get("kind", "") or "").strip() == "material"
                        else {
                            "ref": str(need.get("produto_codigo", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip(),
                            "fornecedor_linha": str(need.get("fornecedor", "") or "").strip(),
                            "origem": "Produto",
                            "qtd": round(self._parse_float(need.get("qtd_em_falta", 0), 0), 2),
                            "unid": str(need.get("produto_unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco_unit", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                        }
                    )
                    for need in needs
                    if self._parse_float(need.get("qtd_em_falta", 0), 0) > 0
                ],
            }
        )
        note_number = str(note.get("numero", "") or "").strip()
        return {
            "numero": note_number,
            "orders": all_orders,
            "line_count": len(list(note.get("linhas", []) or [])),
            "missing_supplier": sorted(set(item for item in missing_supplier if item)),
            "detail": self.ne_detail(note_number),
        }

    def order_detail(self, numero: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        data = self.ensure_data()
        if self._ensure_order_fabrication_order(enc):
            self._save(force=True)
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn) and cliente_codigo:
            cliente_obj = find_cliente_fn(data, cliente_codigo) or {}
        pieces = []
        for piece in self.desktop_main.encomenda_pecas(enc):
            ops = self.desktop_main.ensure_peca_operacoes(piece)
            qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
            qty_prod = (
                self._parse_float(piece.get("produzido_ok", 0), 0)
                + self._parse_float(piece.get("produzido_nok", 0), 0)
                + self._parse_float(piece.get("produzido_qualidade", 0), 0)
            )
            pieces.append(
                {
                    "id": str(piece.get("id", "")).strip(),
                    "ref_interna": str(piece.get("ref_interna", "")).strip(),
                    "ref_externa": str(piece.get("ref_externa", "")).strip(),
                    "material": str(piece.get("material", "")).strip(),
                    "tipo_material": str(piece.get("tipo_material", "") or "CHAPA").strip(),
                    "subtipo_material": str(piece.get("subtipo_material", "") or piece.get("material", "")).strip(),
                    "espessura": str(piece.get("espessura", "")).strip(),
                    "dimensao": str(piece.get("dimensao", piece.get("dimensoes", "")) or "").strip(),
                    "perfil_tipo": str(piece.get("perfil_tipo", "") or "").strip(),
                    "perfil_tamanho": str(piece.get("perfil_tamanho", "") or "").strip(),
                    "comprimento_mm": self._parse_float(piece.get("comprimento_mm", 0), 0),
                    "tubo_forma": str(piece.get("tubo_forma", "") or "").strip(),
                    "lado_a": self._parse_float(piece.get("lado_a", 0), 0),
                    "lado_b": self._parse_float(piece.get("lado_b", 0), 0),
                    "tubo_espessura": self._parse_float(piece.get("tubo_espessura", 0), 0),
                    "diametro": self._parse_float(piece.get("diametro", 0), 0),
                    "estado": str(piece.get("estado", "")).strip(),
                    "qtd_plan": self._fmt(qty_plan),
                    "qtd_prod": self._fmt(qty_prod),
                    "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                    "of": str(piece.get("of", "") or "").strip(),
                    "opp": str(piece.get("opp", "") or "").strip(),
                    "operacoes": " + ".join(
                        [self.desktop_main.normalize_operacao_nome(op.get("nome", "")) for op in list(ops or []) if str(op.get("nome", "")).strip()]
                    ),
                    "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                    "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                    "ficheiros": [
                        str(item or "").strip()
                        for item in list(piece.get("ficheiros", []) or [])
                        if str(item or "").strip()
                    ],
                }
            )
        materials = []
        materials_tree = []
        for mat in list(enc.get("materiais", []) or []):
            esp_rows = []
            for esp in list(mat.get("espessuras", []) or []):
                op_times = self._planning_operation_times_map(esp)
                machine_map = self._order_esp_machine_map(esp)
                planning_ops = [op for op in self._planning_ops_from_esp_obj(esp) if op != "Montagem"]
                other_ops = [op for op in planning_ops if op != "Corte Laser"]
                ops_summary = []
                for op_name in other_ops:
                    op_value = str(op_times.get(op_name, "") or "").strip()
                    if op_value:
                        ops_summary.append(f"{op_name}: {self._fmt(op_value)} min")
                    else:
                        ops_summary.append(op_name)
                resource_summary = []
                for op_name in planning_ops:
                    resource_txt = str(machine_map.get(op_name, "") or "").strip()
                    if resource_txt:
                        resource_summary.append(f"{op_name}: {resource_txt}")
                esp_row = {
                    "material": str(mat.get("material", "")).strip(),
                    "espessura": str(esp.get("espessura", "")).strip(),
                    "estado": str(esp.get("estado", "")).strip(),
                    "tempo_min": self._fmt(esp.get("tempo_min", 0)),
                    "tempos_operacao": {op: self._fmt(value) for op, value in op_times.items() if str(value or "").strip()},
                    "operacoes_planeamento": planning_ops,
                    "tempo_operacoes_txt": " | ".join(ops_summary) or "-",
                    "maquinas_operacao": machine_map,
                    "recursos_operacao_txt": " | ".join(resource_summary) or "-",
                    "pecas": len(list(esp.get("pecas", []) or [])),
                }
                materials.append(esp_row)
                esp_rows.append(esp_row)
            materials_tree.append(
                {
                    "material": str(mat.get("material", "")).strip(),
                    "estado": str(mat.get("estado", "")).strip(),
                    "espessuras": esp_rows,
                }
            )
        montagem_items = []
        for item in list(enc.get("montagem_itens", []) or []):
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            plan = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            consumed = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            montagem_items.append(
                {
                    "tipo_item": item_type,
                    "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                    "tipo_label": "Matéria-prima" if self._montagem_item_is_raw_material(item) else self.desktop_main.orc_line_type_label(item_type),
                    "item_key": self._montagem_item_key(item),
                    "descricao": str(item.get("descricao", "") or "").strip(),
                    "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                    "material": str(item.get("material", "") or "").strip(),
                    "espessura": str(item.get("espessura", "") or "").strip(),
                    "dimensao": str(item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                    "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                    "qtd_planeada": plan,
                    "qtd_consumida": consumed,
                    "qtd_pendente": round(max(0.0, plan - consumed), 2),
                    "tempo_total_min": round(self._parse_float(item.get("tempo_total_min", item.get("tempo_min", 0)), 0), 2),
                    "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                    "conjunto_codigo": str(item.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(item.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(item.get("grupo_uuid", "") or "").strip(),
                    "estado": str(item.get("estado", "") or "").strip(),
                    "created_at": str(item.get("created_at", "") or "").strip(),
                    "consumed_at": str(item.get("consumed_at", "") or "").strip(),
                    "consumed_by": str(item.get("consumed_by", "") or "").strip(),
                }
            )
        montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "Nao aplicavel")
        montagem_shortages = self._order_montagem_shortages(enc)
        montagem_tempo_min = round(self._parse_float(self.desktop_main.encomenda_montagem_tempo_min(enc), 0), 2)
        montagem_resumo = str(self.desktop_main.encomenda_montagem_resumo(enc) or "").strip()
        return {
            "numero": str(enc.get("numero", "")).strip(),
            "tipo_encomenda": str(enc.get("tipo_encomenda", "") or "Cliente").strip(),
            "of_codigo": str(enc.get("of_codigo", "") or "").strip(),
            "ordem_fabrico": dict(enc.get("ordem_fabrico", {}) or {}),
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "posto_trabalho": self._order_workcenter(enc),
            "estado": str(enc.get("estado", "")).strip(),
            "data_entrega": str(enc.get("data_entrega", "")).strip(),
            "zona_transporte": self._transport_zone_for_order(enc, cliente_obj),
            "local_descarga": str(enc.get("local_descarga", "") or cliente_obj.get("morada", "") or "").strip(),
            "nota_cliente": str(enc.get("nota_cliente", "")).strip(),
            "nota_transporte": str(enc.get("nota_transporte", "") or "").strip(),
            "preco_transporte": round(self._parse_float(enc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(enc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(enc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(enc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(enc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(enc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(enc.get("referencia_transporte", "") or "").strip(),
            "transporte_numero": str(enc.get("transporte_numero", "") or "").strip(),
            "estado_transporte": str(enc.get("estado_transporte", "") or "").strip(),
            "tempo_estimado": self._parse_float(enc.get("tempo_estimado", enc.get("tempo", 0)), 0),
            "observacoes": str(enc.get("Observacoes", "") or enc.get("Observa??es", "") or "").strip(),
            "cativar": bool(enc.get("cativar")),
            "numero_orcamento": str(enc.get("numero_orcamento", "") or "").strip(),
            "can_edit_structure": not bool(str(enc.get("numero_orcamento", "") or "").strip()),
            "montagem_estado": montagem_estado,
            "montagem_tempo_min": montagem_tempo_min,
            "montagem_resumo": montagem_resumo,
            "montagem_stock_ready": not bool(montagem_shortages),
            "montagem_shortages": montagem_shortages,
            "montagem_items": montagem_items,
            "can_consume_montagem": any(
                (
                    self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) in {self.desktop_main.ORC_LINE_TYPE_PRODUCT, self.desktop_main.ORC_LINE_TYPE_SERVICE}
                    or self._montagem_item_is_raw_material(row)
                )
                and self._parse_float(row.get("qtd_planeada", 0), 0) > self._parse_float(row.get("qtd_consumida", 0), 0)
                for row in montagem_items
            ),
            "reservas": [
                {
                    "material_id": str(row.get("material_id", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "espessura": str(row.get("espessura", "") or "").strip(),
                    "quantidade": self._fmt(row.get("quantidade", 0)),
                }
                for row in list(enc.get("reservas", []) or [])
            ],
            "obs_interrupcao": str(enc.get("obs_interrupcao", "")).strip(),
            "pieces": pieces,
            "materials": materials,
            "materials_tree": materials_tree,
        }

    def _material_assistant_status_payload(
        self,
        suggestion_id: str,
        feedback_map: dict[str, dict[str, Any]] | None = None,
    ) -> dict[str, str]:
        feedback = dict(feedback_map or {})
        row = dict(feedback.get(str(suggestion_id or "").strip(), {}) or {})
        decision = str(row.get("decision", "") or "").strip().lower()
        if decision == "accepted":
            return {"key": "accepted", "label": "Validada", "tone": "success"}
        if decision == "ignored":
            return {"key": "ignored", "label": "Ignorada hoje", "tone": "default"}
        return {"key": "new", "label": "Nova", "tone": "warning"}

    def _material_assistant_priority_meta(
        self,
        kind: str,
        *,
        due_days: int | None = None,
        next_action_hours: float | None = None,
    ) -> dict[str, Any]:
        score = 35
        if due_days is not None:
            if due_days <= 0:
                score = max(score, 94)
            elif due_days == 1:
                score = max(score, 86)
            elif due_days <= 3:
                score = max(score, 72)
            elif due_days <= 5:
                score = max(score, 58)
        if next_action_hours is not None:
            if next_action_hours <= 12:
                score = max(score, 96)
            elif next_action_hours <= 24:
                score = max(score, 88)
            elif next_action_hours <= 48:
                score = max(score, 74)
            elif next_action_hours <= 120:
                score = max(score, 56)
        kind_txt = str(kind or "").strip().lower()
        if kind_txt == "shortage":
            score = min(100, score + 8)
        elif kind_txt == "uncativated":
            score = min(100, score + 6)
        elif kind_txt in {"retalho", "keep_ready"}:
            score = min(100, score + 4)
        elif kind_txt in {"fito_lot", "separate_lot"}:
            score = min(100, score + 2)
        if score >= 90:
            return {"score": score, "label": "Critica", "tone": "danger"}
        if score >= 70:
            return {"score": score, "label": "Alta", "tone": "warning"}
        if score >= 50:
            return {"score": score, "label": "Media", "tone": "info"}
        return {"score": score, "label": "Baixa", "tone": "default"}

    def _material_assistant_resource_key(self, material: str, espessura: str) -> tuple[str, str]:
        return (
            self.encomendas_actions._norm_material(material),
            self.encomendas_actions._norm_espessura(espessura),
        )

    def _material_assistant_need_sort_key(self, need: dict[str, Any]) -> tuple[Any, ...]:
        next_action_txt = str(need.get("next_action_at", "") or "").strip()
        try:
            next_action_dt = datetime.fromisoformat(next_action_txt) if next_action_txt else None
        except Exception:
            next_action_dt = None
        delivery_txt = str(need.get("data_entrega", "") or "").strip() or "9999-99-99"
        return (
            next_action_dt or datetime.max,
            delivery_txt,
            str(need.get("numero", "") or "").strip(),
            str(need.get("material", "") or "").strip(),
            str(need.get("espessura", "") or "").strip(),
        )

    def _material_assistant_shift_payload(self, next_action_at: Any, fallback_date: Any = "") -> dict[str, Any]:
        raw_next = str(next_action_at or "").strip()
        raw_fallback = str(fallback_date or "").strip()[:10]
        next_dt: datetime | None = None
        if raw_next:
            try:
                next_dt = datetime.fromisoformat(raw_next)
            except Exception:
                next_dt = None
        if next_dt is None and raw_fallback:
            try:
                next_dt = datetime.combine(date.fromisoformat(raw_fallback), datetime.min.time()) + timedelta(hours=8)
            except Exception:
                next_dt = None

        date_key = ""
        date_label = "-"
        time_label = "-"
        shift_label = "Sem turno"
        shift_order = 9
        if next_dt is not None:
            date_key = next_dt.date().isoformat()
            date_label = next_dt.strftime("%d/%m/%Y")
            time_label = next_dt.strftime("%H:%M")
            hour = next_dt.hour
            if 6 <= hour < 14:
                shift_label = "Manha"
                shift_order = 0
            elif 14 <= hour < 22:
                shift_label = "Tarde"
                shift_order = 1
            else:
                shift_label = "Noite"
                shift_order = 2
        elif raw_fallback:
            date_key = raw_fallback
            try:
                date_label = date.fromisoformat(raw_fallback).strftime("%d/%m/%Y")
            except Exception:
                date_label = raw_fallback
            shift_label = "Sem turno"
            shift_order = 8

        return {
            "date_key": date_key or "9999-99-99",
            "date_label": date_label,
            "time_label": time_label,
            "shift_label": shift_label,
            "shift_order": shift_order,
        }

    def _material_assistant_apply_resource_priority(self, needs: list[dict[str, Any]]) -> None:
        groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for need in list(needs or []):
            resource_key = tuple(need.get("resource_key", ()) or ())
            if len(resource_key) != 2:
                resource_key = self._material_assistant_resource_key(
                    str(need.get("material", "") or ""),
                    str(need.get("espessura", "") or ""),
                )
                need["resource_key"] = resource_key
            groups.setdefault(resource_key, []).append(need)

        for group_rows in groups.values():
            group_rows.sort(key=self._material_assistant_need_sort_key)

            standard_pool: list[dict[str, Any]] = []
            seen_pool: set[tuple[str, str, str]] = set()
            for need in group_rows:
                for raw_candidate in list(need.get("standard_candidates", []) or []):
                    candidate = dict(raw_candidate or {})
                    marker = (
                        str(candidate.get("lote", "") or "").strip().lower(),
                        str(candidate.get("material_id", "") or "").strip(),
                        str(candidate.get("dimensao", "") or "").strip(),
                    )
                    if marker in seen_pool:
                        continue
                    seen_pool.add(marker)
                    standard_pool.append(candidate)

            standard_pool.sort(
                key=lambda row: (
                    str(row.get("lote", "") or "").strip().lower() or "zzzz",
                    -self._parse_float(row.get("disponivel", 0), 0),
                    str(row.get("material_id", "") or "").strip(),
                )
            )

            for index, need in enumerate(group_rows, start=1):
                need["priority_position"] = index
                need["group_size"] = len(group_rows)
                need["quantidade_preparar"] = round(
                    self._parse_float(
                        need.get("reserved_qty", 0) if self._parse_float(need.get("reserved_qty", 0), 0) > 0 else need.get("piece_qty", 0),
                        0,
                    ),
                    2,
                )
                material_cativado = bool(need.get("material_cativado"))
                if not material_cativado:
                    need["preferred_lot"] = ""
                    need["preferred_material_id"] = ""
                    need["preferred_dimensao"] = ""
                    need["preferred_disponivel"] = 0.0
                else:
                    need["preferred_lot"] = str(need.get("preferred_lot", "") or "").strip()
                    need["preferred_material_id"] = str(need.get("preferred_material_id", "") or "").strip()
                    need["preferred_dimensao"] = str(need.get("preferred_dimensao", "") or "").strip()
                    need["preferred_disponivel"] = round(self._parse_float(need.get("preferred_disponivel", 0), 0), 2)
                need["lot_change_required"] = False
                need["lot_change_from_lot"] = ""
                need["lot_change_to_lot"] = str(need.get("preferred_lot", "") or "").strip()
                need["lot_change_conflict_order"] = ""
                need["lot_change_conflict_client"] = ""
                need["lot_change_note"] = ""

            occupied_by_lot: dict[str, list[dict[str, Any]]] = {}
            for need in group_rows:
                current_lot = str(need.get("current_lot", "") or "").strip()
                if not current_lot:
                    continue
                occupied_by_lot.setdefault(current_lot.lower(), []).append(need)
            for rows in occupied_by_lot.values():
                rows.sort(key=lambda row: int(row.get("priority_position", 9999) or 9999))

            for need in group_rows:
                if not bool(need.get("material_cativado")):
                    continue
                suggested_lot = str(need.get("preferred_lot", "") or "").strip()
                current_lot = str(need.get("current_lot", "") or "").strip()
                if not suggested_lot:
                    continue
                if current_lot and current_lot.lower() == suggested_lot.lower():
                    continue
                competing_rows = [
                    row
                    for row in list(occupied_by_lot.get(suggested_lot.lower(), []) or [])
                    if row is not need
                ]
                if not competing_rows:
                    continue
                competing = dict(competing_rows[0] or {})
                if int(competing.get("priority_position", 9999) or 9999) <= int(need.get("priority_position", 9999) or 9999):
                    continue
                need["lot_change_required"] = True
                need["lot_change_from_lot"] = current_lot or "Sem lote definido"
                need["lot_change_to_lot"] = suggested_lot
                need["lot_change_conflict_order"] = str(competing.get("numero", "") or "").strip()
                need["lot_change_conflict_client"] = str(competing.get("cliente", "") or "").strip()
                need["lot_change_note"] = (
                    f"{need.get('numero', '-')} ficou mais urgente do que {need.get('lot_change_conflict_order', '-')}; "
                    f"o lote {suggested_lot} deve seguir para a encomenda mais urgente."
                )

    def _material_assistant_stock_option_label(self, row: dict[str, Any]) -> str:
        lote = str(row.get("lote", "") or "-").strip() or "-"
        dimensao = str(row.get("dimensao", "") or "").strip()
        if dimensao:
            return f"{lote} {dimensao}"
        return lote

    def _material_assistant_stock_options_text(self, standard_rows: list[dict[str, Any]], retalho_rows: list[dict[str, Any]]) -> str:
        def _join_options(rows: list[dict[str, Any]], limit: int = 4) -> str:
            labels = [
                self._material_assistant_stock_option_label(dict(row or {}))
                for row in list(rows or [])[: max(1, int(limit or 1))]
                if self._material_assistant_stock_option_label(dict(row or {}))
            ]
            extra = max(0, len(list(rows or [])) - len(labels))
            if extra > 0:
                labels.append(f"+{extra} opcoes")
            return ", ".join(labels)

        parts: list[str] = []
        standard_txt = _join_options(standard_rows, limit=4)
        retalho_txt = _join_options(retalho_rows, limit=4)
        if standard_txt:
            parts.append(f"Chapas: {standard_txt}")
        if retalho_txt:
            parts.append(f"Retalhos: {retalho_txt}")
        return " | ".join(parts)

    def _business_horizon_end_date(self, anchor_date: date | None = None, business_days: int = 4) -> date:
        current = anchor_date if isinstance(anchor_date, date) else date.today()
        target = max(1, int(business_days or 1))
        counted = 1 if current.weekday() < 5 else 0
        while counted < target:
            current += timedelta(days=1)
            if current.weekday() < 5:
                counted += 1
        if counted <= 0:
            while current.weekday() >= 5:
                current += timedelta(days=1)
        return current

    def material_assistant_snapshot(self, horizon_days: int = 5) -> dict[str, Any]:
        data = self.ensure_data()
        horizon = max(1, int(horizon_days or 5))
        now_dt = datetime.now()
        today_dt = date.today()
        horizon_date = self._business_horizon_end_date(today_dt, horizon)
        feedback_map = self.material_assistant_feedback()

        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if isinstance(row, dict)
        }

        upcoming_plan: dict[tuple[str, str, str], dict[str, Any]] = {}
        for block in list(data.get("plano", []) or []):
            if not isinstance(block, dict):
                continue
            if not self._planning_row_matches_operation(block, "Corte Laser"):
                continue
            start_dt, _end_dt = self._planning_block_bounds(block)
            if start_dt is None:
                continue
            if start_dt < now_dt - timedelta(hours=12):
                continue
            key = self._planning_item_key(block.get("encomenda", ""), block.get("material", ""), block.get("espessura", ""))
            current = upcoming_plan.get(key)
            if current is None or start_dt < current["start_dt"]:
                posto_txt = (
                    str(block.get("posto", "") or "").strip()
                    or str(block.get("posto_trabalho", "") or "").strip()
                    or str(block.get("maquina", "") or "").strip()
                    or "Sem posto"
                )
                upcoming_plan[key] = {
                    "start_dt": start_dt,
                    "data": str(block.get("data", "") or "").strip(),
                    "inicio": str(block.get("inicio", "") or "").strip(),
                    "duracao_min": int(float(block.get("duracao_min", 0) or 0)),
                    "posto": posto_txt,
                }

        needs: list[dict[str, Any]] = []
        for enc in list(data.get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            numero = str(enc.get("numero", "") or "").strip()
            if not numero:
                continue
            order_workcenter = self._order_workcenter(enc)
            enc_state = str(enc.get("estado", "") or "").strip()
            enc_state_norm = self.desktop_main.norm_text(enc_state)
            if "concl" in enc_state_norm or "cancel" in enc_state_norm:
                continue
            cliente_codigo = str(enc.get("cliente", "") or "").strip()
            cliente_nome = clients.get(cliente_codigo, "")
            cliente_label = " - ".join(value for value in (cliente_codigo, cliente_nome) if value).strip(" -")
            delivery_txt = str(enc.get("data_entrega", "") or "").strip()[:10]
            delivery_date: date | None = None
            if len(delivery_txt) == 10:
                try:
                    delivery_date = date.fromisoformat(delivery_txt)
                except Exception:
                    delivery_date = None
            for mat in list(enc.get("materiais", []) or []):
                mat_name = str(mat.get("material", "") or "").strip()
                if not mat_name:
                    continue
                for esp_obj in list(mat.get("espessuras", []) or []):
                    esp = str(esp_obj.get("espessura", "") or "").strip()
                    if not esp:
                        continue
                    esp_state = str(esp_obj.get("estado", "") or enc_state).strip()
                    esp_state_norm = self.desktop_main.norm_text(esp_state)
                    if "concl" in esp_state_norm or "cancel" in esp_state_norm:
                        continue
                    key = self._planning_item_key(numero, mat_name, esp)
                    plan_info = upcoming_plan.get(key)
                    next_action_dt = plan_info.get("start_dt") if plan_info else None
                    if next_action_dt is None and delivery_date is not None:
                        # A folha operacional de separacao nao deve inventar trabalho a partir
                        # da data de entrega. Sem planeamento laser, reserva ou lote definido,
                        # a linha fica fora da separacao para evitar instrucoes falsas.
                        current_lot_probe = str(esp_obj.get("lote_baixa", "") or "").strip()
                        has_reservation_probe = any(
                            self.encomendas_actions._norm_material(reserva.get("material")) == self.encomendas_actions._norm_material(mat_name)
                            and self.encomendas_actions._norm_espessura(reserva.get("espessura")) == self.encomendas_actions._norm_espessura(esp)
                            for reserva in list(enc.get("reservas", []) or [])
                            if isinstance(reserva, dict)
                        )
                        if not current_lot_probe and not has_reservation_probe:
                            continue
                        next_action_dt = datetime.combine(delivery_date, datetime.min.time()) + timedelta(hours=8)
                    if next_action_dt is None:
                        continue
                    if plan_info:
                        if next_action_dt.date() > horizon_date:
                            continue
                    elif delivery_date is not None and delivery_date > horizon_date:
                        continue
                    due_days = (delivery_date - today_dt).days if delivery_date is not None else None
                    next_action_hours = round((next_action_dt - now_dt).total_seconds() / 3600.0, 1)
                    matching_reservas = []
                    reserved_qty = 0.0
                    for reserva in list(enc.get("reservas", []) or []):
                        if self.encomendas_actions._norm_material(reserva.get("material")) != self.encomendas_actions._norm_material(mat_name):
                            continue
                        if self.encomendas_actions._norm_espessura(reserva.get("espessura")) != self.encomendas_actions._norm_espessura(esp):
                            continue
                        matching_reservas.append(dict(reserva))
                        reserved_qty += self._parse_float(reserva.get("quantidade", 0), 0)
                    candidates = list(self.material_candidates(mat_name, esp) or [])
                    current_lot = str(esp_obj.get("lote_baixa", "") or "").strip()
                    retalho_candidates = [
                        dict(row)
                        for row in candidates
                        if bool(row.get("is_retalho"))
                        and str(row.get("origem_encomenda", "") or "").strip() != numero
                        and str(row.get("origem_lote", "") or "").strip().lower() != current_lot.lower()
                    ]
                    standard_candidates = [dict(row) for row in candidates if not bool(row.get("is_retalho"))]
                    standard_candidates.sort(
                        key=lambda row: (
                            str(row.get("lote", "") or "zzzz").lower(),
                            -self._parse_float(row.get("disponivel", 0), 0),
                            str(row.get("material_id", "") or ""),
                        )
                    )
                    retalho_candidates.sort(
                        key=lambda row: (
                            -self._parse_float(row.get("disponivel", 0), 0),
                            str(row.get("dimensao", "") or ""),
                            str(row.get("material_id", "") or ""),
                        )
                    )
                    piece_qty = round(
                        sum(self._parse_float(piece.get("quantidade_pedida", 0), 0) for piece in list(esp_obj.get("pecas", []) or [])),
                        2,
                    )
                    if piece_qty <= 0:
                        piece_qty = float(len(list(esp_obj.get("pecas", []) or [])) or 1)
                    chapa_reservada = self._order_reserved_sheet(numero, mat_name, esp)
                    reserved_lot = next(
                        (
                            str(reserva.get("lote", "") or "").strip()
                            for reserva in matching_reservas
                            if str(reserva.get("lote", "") or "").strip()
                        ),
                        "",
                    )
                    reserved_material_id = next(
                        (
                            str(reserva.get("material_id", "") or "").strip()
                            for reserva in matching_reservas
                            if str(reserva.get("material_id", "") or "").strip()
                        ),
                        "",
                    )
                    material_cativado = bool(current_lot or matching_reservas)
                    preferred_standard = {}
                    if material_cativado:
                        for row in standard_candidates:
                            row_lote = str(row.get("lote", "") or "").strip()
                            row_id = str(row.get("material_id", "") or "").strip()
                            if (
                                (current_lot and row_lote.lower() == current_lot.lower())
                                or (reserved_lot and row_lote.lower() == reserved_lot.lower())
                                or (reserved_material_id and row_id == reserved_material_id)
                            ):
                                preferred_standard = dict(row)
                                break
                    need = {
                        "key": "|".join(key),
                        "resource_key": self._material_assistant_resource_key(mat_name, esp),
                        "numero": numero,
                        "cliente": cliente_label or cliente_codigo or "-",
                        "cliente_codigo": cliente_codigo,
                        "cliente_nome": cliente_nome,
                        "material": mat_name,
                        "espessura": esp,
                        "estado": esp_state or enc_state,
                        "data_entrega": delivery_txt,
                        "due_days": due_days,
                        "next_action_at": next_action_dt.isoformat(timespec="minutes"),
                        "next_action_label": (
                            f"{plan_info.get('data', '')} {plan_info.get('inicio', '')}".strip()
                            if plan_info
                            else (delivery_txt or "-")
                        ),
                        "posto_trabalho": (
                            str((plan_info or {}).get("posto", "") or "").strip()
                            or order_workcenter
                            or "Sem posto"
                        ),
                        "next_action_hours": next_action_hours,
                        "plan_origin": "Planeamento" if plan_info else "Entrega",
                        "reserved_qty": round(reserved_qty, 2),
                        "reserved_count": len(matching_reservas),
                        "material_cativado": material_cativado,
                        "current_lot": current_lot,
                        "preferred_lot": str(current_lot or reserved_lot or preferred_standard.get("lote", "") or "").strip(),
                        "preferred_material_id": str(reserved_material_id or preferred_standard.get("material_id", "") or "").strip(),
                        "preferred_dimensao": str(preferred_standard.get("dimensao", "") or "").strip(),
                        "preferred_disponivel": round(self._parse_float(preferred_standard.get("disponivel", 0), 0), 2),
                        "retalho_count": len(retalho_candidates),
                        "standard_lot_count": len(standard_candidates),
                        "piece_qty": piece_qty,
                        "chapa": chapa_reservada,
                        "retalho_candidates": retalho_candidates[:4],
                        "standard_candidates": standard_candidates[:4],
                        "stock_options_txt": self._material_assistant_stock_options_text(standard_candidates[:4], retalho_candidates[:4]),
                        "reservation_rows": matching_reservas,
                        "stock_state": "Sem stock" if not candidates else (f"{len(retalho_candidates)} retalhos + {len(standard_candidates)} lotes" if retalho_candidates else f"{len(standard_candidates)} lotes disponiveis"),
                        "stock_ready": bool(candidates),
                    }
                    needs.append(need)

        self._material_assistant_apply_resource_priority(needs)

        suggestions: list[dict[str, Any]] = []
        for need in needs:
            numero = str(need.get("numero", "") or "").strip()
            material = str(need.get("material", "") or "").strip()
            esp = str(need.get("espessura", "") or "").strip()
            due_days = need.get("due_days")
            due_days = int(due_days) if isinstance(due_days, int) else (int(due_days) if due_days is not None else None)
            next_action_hours = float(need.get("next_action_hours", 0) or 0) if need.get("next_action_hours") is not None else None
            has_stock = bool(need.get("stock_ready"))
            retalhos = list(need.get("retalho_candidates", []) or [])
            standard = list(need.get("standard_candidates", []) or [])
            reservations = list(need.get("reservation_rows", []) or [])
            current_lot = str(need.get("current_lot", "") or "").strip()
            preferred_lot = str(need.get("preferred_lot", "") or "").strip()
            material_cativado = bool(need.get("material_cativado"))

            def _append_suggestion(kind: str, headline: str, recommendation: str, detail_lines: list[str], *, target_id: str = "") -> None:
                suggestion_id = "|".join(
                    [
                        str(kind or "").strip(),
                        numero,
                        material,
                        esp,
                        str(target_id or preferred_lot or current_lot or "-").strip(),
                    ]
                )
                priority = self._material_assistant_priority_meta(kind, due_days=due_days, next_action_hours=next_action_hours)
                status = self._material_assistant_status_payload(suggestion_id, feedback_map)
                suggestions.append(
                    {
                        "id": suggestion_id,
                        "kind": kind,
                        "headline": headline,
                        "recommendation": recommendation,
                        "detail_lines": list(detail_lines or []),
                        "numero": numero,
                        "cliente": need.get("cliente", "-"),
                        "posto_trabalho": str(need.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto",
                        "material": material,
                        "espessura": esp,
                        "when": str(need.get("next_action_label", "") or "-").strip(),
                        "delivery": str(need.get("data_entrega", "") or "").strip(),
                        "priority_score": int(priority.get("score", 0) or 0),
                        "priority_label": str(priority.get("label", "") or "Baixa"),
                        "priority_tone": str(priority.get("tone", "") or "default"),
                        "next_action_at": str(need.get("next_action_at", "") or "").strip(),
                        "plan_origin": str(need.get("plan_origin", "") or "Entrega").strip(),
                        "status_key": status["key"],
                        "status_label": status["label"],
                        "status_tone": status["tone"],
                        "need_key": str(need.get("key", "") or ""),
                        "preferred_lot": preferred_lot,
                        "current_lot": current_lot,
                        "reserved_qty": float(need.get("reserved_qty", 0) or 0),
                        "retalho_count": int(need.get("retalho_count", 0) or 0),
                        "stock_state": str(need.get("stock_state", "") or ""),
                    }
                )

            if not has_stock:
                _append_suggestion(
                    "shortage",
                    "Sem matéria-prima disponível",
                    f"{numero} precisa de {material} {esp} e não existe lote nem retalho disponível.",
                    [
                        f"Cliente: {need.get('cliente', '-')}",
                        f"Próxima ação: {need.get('next_action_label', '-')}",
                        "Ação sugerida: validar compra/abertura de lote antes de libertar para o corte.",
                    ],
                )
                continue

            if not material_cativado:
                _append_suggestion(
                    "uncativated",
                    "Material sem cativação",
                    f"{numero} precisa de {material} {esp}, mas ainda não existe chapa/lote cativado.",
                    [
                        f"Cliente: {need.get('cliente', '-')}",
                        f"Próxima ação: {need.get('next_action_label', '-')}",
                        f"Opções disponíveis: {need.get('stock_options_txt', '-') or '-'}",
                        "Ação obrigatória: cativar a chapa antes de aparecer uma instrução de separação com lote.",
                    ],
                )
                continue

            if reservations and next_action_hours is not None and next_action_hours <= 24:
                _append_suggestion(
                    "keep_ready",
                    "Evitar arrumação desnecessária",
                    (
                        f"Material já cativado para {numero}; vai ser necessário"
                        f" em {need.get('next_action_label', '-')}. Mantém disponível."
                    ),
                    [
                        f"Quantidade cativada: {self._fmt(need.get('reserved_qty', 0))}",
                        f"Chapa/Lote atual: {current_lot or need.get('chapa', '-')}",
                        "Ação sugerida: não arrumar este material no stock intermédio.",
                    ],
                )

            if preferred_lot:
                if bool(need.get("lot_change_required")):
                    conflicting_order = str(need.get("lot_change_conflict_order", "") or "").strip()
                    conflicting_client = str(need.get("lot_change_conflict_client", "") or "").strip()
                    _append_suggestion(
                        "fito_lot",
                        "Ajuste de lote por urgencia / FIFO",
                        (
                            f"{numero} ficou mais urgente e deve usar o lote {preferred_lot}"
                            f" em vez de {current_lot or 'sem lote definido'}."
                        ),
                        [
                            f"Prioridade atual no recurso: posição {need.get('priority_position', '-')}",
                            (
                                f"Conflito identificado com {conflicting_order}"
                                f"{f' ({conflicting_client})' if conflicting_client else ''}."
                            ),
                            f"Lote atual: {current_lot or 'Sem lote definido'} | lote sugerido: {preferred_lot}",
                            f"Opcoes disponiveis: {need.get('stock_options_txt', '-') or '-'}",
                            "Sugestao: rever a cativacao e trocar a separacao para respeitar a encomenda mais urgente sem abrir chapa desnecessaria.",
                        ],
                    )
        def _suggestion_sort_key(row: dict[str, Any]) -> tuple[Any, ...]:
            raw_next = str(row.get("next_action_at", "") or "").strip()
            try:
                next_dt = datetime.fromisoformat(raw_next) if raw_next else None
            except Exception:
                next_dt = None
            return (
                -int(row.get("priority_score", 0) or 0),
                next_dt or datetime.max,
                0 if str(row.get("plan_origin", "") or "").strip() == "Planeamento" else 1,
                str(row.get("delivery", "") or "9999-99-99"),
                str(row.get("numero", "") or ""),
                str(row.get("material", "") or ""),
            )

        suggestions.sort(key=_suggestion_sort_key)
        needs.sort(
            key=lambda row: (
                0 if bool(row.get("stock_ready")) else 1,
                *self._material_assistant_need_sort_key(row),
            )
        )
        cards = [
            {
                "title": "Linhas separacao",
                "value": str(len(needs)),
                "subtitle": f"Horizonte de {horizon} dias úteis",
                "tone": "info",
            },
            {
                "title": "Trocas por urgencia",
                "value": str(len([row for row in suggestions if str(row.get("kind", "") or "") == "fito_lot" and str(row.get("status_key", "") or "") == "new"])),
                "subtitle": "Mudancas reais de prioridade",
                "tone": "warning",
            },
            {
                "title": "Nao arrumar",
                "value": str(len([row for row in suggestions if str(row.get("kind", "") or "") == "keep_ready" and str(row.get("status_key", "") or "") == "new"])),
                "subtitle": "Material para manter pronto",
                "tone": "success",
            },
            {
                "title": "Sem stock",
                "value": str(len([row for row in suggestions if str(row.get("kind", "") or "") == "shortage" and str(row.get("status_key", "") or "") == "new"])),
                "subtitle": "Necessita validacao de compra",
                "tone": "danger",
            },
        ]
        return {
            "generated_at": str(self.desktop_main.now_iso() or "").strip(),
            "horizon_days": horizon,
            "horizon_label": f"{horizon} dias úteis",
            "cards": cards,
            "suggestions": suggestions[:60],
            "needs": needs[:80],
        }

    def material_assistant_separation_rows(self, horizon_days: int = 5) -> list[dict[str, Any]]:
        snapshot = self.material_assistant_snapshot(horizon_days=horizon_days)
        check_map = self.material_assistant_checks()
        suggestions_by_need: dict[str, list[dict[str, Any]]] = {}
        for row in list(snapshot.get("suggestions", []) or []):
            need_key = str(row.get("need_key", "") or "").strip()
            if not need_key:
                continue
            suggestions_by_need.setdefault(need_key, []).append(dict(row))
        for rows in suggestions_by_need.values():
            rows.sort(key=lambda row: (-int(row.get("priority_score", 0) or 0), str(row.get("kind", "") or "")))

        result: list[dict[str, Any]] = []
        for need in list(snapshot.get("needs", []) or []):
            current_suggestions = list(suggestions_by_need.get(str(need.get("key", "") or "").strip(), []) or [])
            lead = dict(current_suggestions[0]) if current_suggestions else {}
            material_cativado = bool(need.get("material_cativado"))
            preferred_lot = str(need.get("preferred_lot", "") or "-").strip() or "-"
            if not material_cativado:
                preferred_lot = "-"
                recommendation = "Sem material cativado - cativar chapa antes de separar"
            else:
                recommendation = (
                    f"Separar lote {preferred_lot}"
                    if preferred_lot and preferred_lot != "-"
                    else "Validar materia-prima cativada"
                )
            if bool(need.get("lot_change_required")) and preferred_lot and preferred_lot != "-":
                recommendation = f"Separar lote {preferred_lot} para a encomenda urgente"
            reserved_qty = self._parse_float(need.get("reserved_qty", 0), 0)
            need_key = str(need.get("key", "") or "").strip()
            checks = dict(check_map.get(need_key, {}) or {})
            shift = self._material_assistant_shift_payload(need.get("next_action_at"), need.get("data_entrega", ""))
            material_label = str(need.get("material", "") or "").strip()
            espessura_label = str(need.get("espessura", "") or "").strip()
            posto_trabalho = str(need.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto"
            posto_sort = self.desktop_main.norm_text(posto_trabalho)
            material_sort = self.encomendas_actions._norm_material(material_label)
            esp_sort = self._parse_float(espessura_label, 999999)
            base_group_key = "|".join(
                [
                    posto_sort or "sem-posto",
                    str(shift.get("date_key", "9999-99-99") or "9999-99-99"),
                    f"{int(shift.get('shift_order', 9) or 9):02d}",
                    material_sort,
                    self.encomendas_actions._norm_espessura(espessura_label),
                ]
            )
            material_group = " ".join(part for part in (material_label, f"{espessura_label} mm" if espessura_label else "") if part).strip()
            base_group_label = (
                f"{shift.get('date_label', '-') or '-'} | "
                f"{shift.get('shift_label', 'Sem turno') or 'Sem turno'} | "
                f"{material_group or '-'}"
            )
            standard_candidates = list(need.get("standard_candidates", []) or [])
            retalho_candidates = list(need.get("retalho_candidates", []) or [])
            reservation_rows = list(need.get("reservation_rows", []) or [])

            def _resolve_source_candidate(reserva_row: dict[str, Any]) -> dict[str, Any]:
                reserva_id = str(reserva_row.get("material_id", "") or "").strip()
                if reserva_id:
                    direct = self.material_by_id(reserva_id)
                    if isinstance(direct, dict):
                        return {
                            "material_id": reserva_id,
                            "lote": str(direct.get("lote_fornecedor", "") or direct.get("origem_lote", "") or "-").strip() or "-",
                            "dimensao": "x".join(
                                part
                                for part in (
                                    self._fmt(direct.get("comprimento", 0)),
                                    self._fmt(direct.get("largura", 0)),
                                )
                                if part and part != "0"
                            )
                            or "-",
                            "disponivel": self._parse_float(direct.get("quantidade", 0), 0),
                            "is_retalho": bool(direct.get("is_sobra")),
                        }
                    for candidate in list(standard_candidates) + list(retalho_candidates):
                        if str((candidate or {}).get("material_id", "") or "").strip() == reserva_id:
                            return dict(candidate or {})
                reserva_lote = str(reserva_row.get("lote", "") or "").strip()
                if reserva_lote:
                    for candidate in list(standard_candidates) + list(retalho_candidates):
                        if str((candidate or {}).get("lote", "") or "").strip().lower() == reserva_lote.lower():
                            return dict(candidate or {})
                return {}

            def _build_row(
                *,
                source_kind: str,
                source_index: int,
                quantity_value: Any,
                source_candidate: dict[str, Any] | None = None,
                reserva_row: dict[str, Any] | None = None,
            ) -> dict[str, Any]:
                source_candidate = dict(source_candidate or {})
                reserva_row = dict(reserva_row or {})
                source_id = (
                    str(source_candidate.get("material_id", "") or "").strip()
                    or str(reserva_row.get("material_id", "") or "").strip()
                    or str(source_candidate.get("lote", "") or "").strip()
                    or f"{source_kind}-{source_index}"
                )
                row_key = f"{need_key}|{source_id}|{source_kind}"
                row_checks = dict(check_map.get(row_key, {}) or {})
                if not material_cativado and source_kind != "reserva":
                    lote_sugerido = "-"
                    dimensao = "-"
                else:
                    lote_sugerido = (
                        str(source_candidate.get("lote", "") or "").strip()
                        or str(reserva_row.get("lote", "") or "").strip()
                        or preferred_lot
                        or "-"
                    )
                    dimensao = (
                        str(source_candidate.get("dimensao", "") or "").strip()
                        or str(need.get("preferred_dimensao", "") or "").strip()
                        or "-"
                    )
                formato_sort = self.desktop_main.norm_text(dimensao.replace(" ", "")) or "sem-formato"
                is_retalho = bool(source_candidate.get("is_retalho")) or source_kind == "retalho"
                if source_kind == "reserva":
                    reserva_label = "Cativado"
                elif not material_cativado:
                    reserva_label = "Sem material cativado"
                else:
                    reserva_label = "Retalho sugerido" if is_retalho else "Por separar"
                action_text = recommendation
                if source_kind == "reserva":
                    action_text = f"Separar {lote_sugerido} (cativado)"
                elif is_retalho:
                    action_text = f"Avaliar retalho {lote_sugerido}"
                elif material_cativado and lote_sugerido and lote_sugerido != "-":
                    action_text = f"Separar lote {lote_sugerido}"
                parsed_quantity = round(self._parse_float(quantity_value, 0), 2)
                operational_quantity = parsed_quantity if material_cativado else 0.0
                return {
                    "numero": str(need.get("numero", "") or "").strip(),
                    "cliente": str(need.get("cliente", "") or "").strip(),
                    "posto_trabalho": posto_trabalho,
                    "material": material_label,
                    "espessura": espessura_label,
                    "dimensao": dimensao,
                    "quantidade": operational_quantity,
                    "quantidade_necessaria": parsed_quantity,
                    "quantidade_label": self._fmt(operational_quantity) if material_cativado else "-",
                    "necessidade_label": self._fmt(parsed_quantity),
                    "disponivel": round(self._parse_float(source_candidate.get("disponivel", need.get("preferred_disponivel", 0)), 0), 2) if material_cativado else 0.0,
                    "data_entrega": str(need.get("data_entrega", "") or "").strip(),
                    "proxima_acao": str(need.get("next_action_label", "") or "-").strip(),
                    "origem_planeamento": str(need.get("plan_origin", "") or "-").strip(),
                    "reserva_estado": reserva_label,
                    "reserva_qtd": round(self._parse_float(quantity_value if source_kind == "reserva" else reserved_qty, 0), 2),
                    "lote_atual": str(need.get("current_lot", "") or need.get("chapa", "") or "-").strip() or "-",
                    "lote_sugerido": lote_sugerido,
                    "retalhos": int(need.get("retalho_count", 0) or 0),
                    "stock_state": str(need.get("stock_state", "") or "-").strip(),
                    "acao_sugerida": action_text,
                    "alerta_retalho": bool(int(need.get("retalho_count", 0) or 0) > 0),
                    "alerta_texto": "Existe retalho compativel para avaliar" if int(need.get("retalho_count", 0) or 0) > 0 else "",
                    "opcoes_mp": str(need.get("stock_options_txt", "") or "").strip(),
                    "priority_label": str(lead.get("priority_label", "") or ("Alta" if (not bool(need.get("stock_ready")) or not material_cativado) else "Media")).strip(),
                    "priority_tone": str(lead.get("priority_tone", "") or ("danger" if not bool(need.get("stock_ready")) else "warning" if not material_cativado else "info")).strip(),
                    "priority_score": int(lead.get("priority_score", 0) or 0),
                    "status_label": str(lead.get("status_label", "") or "Pendente").strip(),
                    "status_key": str(lead.get("status_key", "") or "new").strip(),
                    "stock_ready": bool(need.get("stock_ready")),
                    "material_cativado": material_cativado,
                    "need_key": need_key,
                    "check_key": row_key,
                    "visto_sep_checked": bool(row_checks.get("sep")),
                    "visto_conf_checked": bool(row_checks.get("conf")),
                    "visto_sep": "[x]" if bool(row_checks.get("sep")) else "[ ]",
                    "visto_conf": "[x]" if bool(row_checks.get("conf")) else "[ ]",
                    "headline": str(lead.get("headline", "") or "").strip(),
                    "planeamento_dia": str(shift.get("date_label", "-") or "-"),
                    "planeamento_dia_iso": str(shift.get("date_key", "9999-99-99") or "9999-99-99"),
                    "planeamento_hora": str(shift.get("time_label", "-") or "-"),
                    "planeamento_turno": str(shift.get("shift_label", "Sem turno") or "Sem turno"),
                    "planeamento_turno_ordem": int(shift.get("shift_order", 9) or 9),
                    "material_group": material_group or "-",
                    "formato_group": dimensao or "-",
                    "formato_sort": formato_sort,
                    "base_group_key": base_group_key,
                    "base_group_label": base_group_label,
                    "posto_sort": posto_sort,
                    "material_sort": material_sort,
                    "esp_sort": esp_sort,
                    "group_key": base_group_key,
                    "group_label": base_group_label,
                    "group_format_label": "",
                    "row_rank": source_index,
                    "source_kind": source_kind,
                }

            if reservation_rows:
                for idx, reserva in enumerate(reservation_rows, start=1):
                    result.append(
                        _build_row(
                            source_kind="reserva",
                            source_index=idx,
                            quantity_value=reserva.get("quantidade", 0),
                            source_candidate=_resolve_source_candidate(dict(reserva or {})),
                            reserva_row=reserva,
                        )
                    )
                continue

            result.append(
                _build_row(
                    source_kind="principal",
                    source_index=1,
                    quantity_value=need.get("quantidade_preparar", 0),
                    source_candidate=dict(standard_candidates[0]) if standard_candidates and material_cativado else {},
                )
            )
        format_map: dict[str, set[str]] = {}
        for row in result:
            base_key = str(row.get("base_group_key", "") or "").strip()
            formato = str(row.get("formato_group", "") or "-").strip() or "-"
            if not base_key:
                continue
            format_map.setdefault(base_key, set()).add(formato)

        for row in result:
            base_key = str(row.get("base_group_key", "") or "").strip()
            base_label = str(row.get("base_group_label", "") or "-").strip() or "-"
            formato = str(row.get("formato_group", "") or "-").strip() or "-"
            grouped_by_format = len(format_map.get(base_key, set())) > 1
            row["format_grouped"] = grouped_by_format
            if grouped_by_format:
                row["group_key"] = f"{base_key}|{str(row.get('formato_sort', '') or 'sem-formato')}"
                row["group_label"] = f"{base_label} | Formato {formato}"
                row["group_format_label"] = f"Formato {formato}"
            else:
                row["group_key"] = base_key
                row["group_label"] = base_label
                row["group_format_label"] = ""

        result.sort(
            key=lambda row: (
                str(row.get("posto_sort", "") or ""),
                str(row.get("planeamento_dia_iso", "") or "9999-99-99"),
                int(row.get("planeamento_turno_ordem", 9) or 9),
                str(row.get("material_sort", "") or ""),
                float(row.get("esp_sort", 999999) or 999999),
                str(row.get("formato_sort", "") or "sem-formato"),
                str(row.get("planeamento_hora", "") or "99:99"),
                int(row.get("row_rank", 999) or 999),
                -int(row.get("priority_score", 0) or 0),
                str(row.get("numero", "") or ""),
            )
        )
        valid_keys = {
            str(row.get("check_key", "") or "").strip()
            for row in result
            if str(row.get("check_key", "") or "").strip()
        }
        self._material_assistant_check_map(valid_keys=valid_keys, persist_pruned=True)
        return result

    def material_assistant_alert_rows(self, horizon_days: int = 5) -> list[dict[str, Any]]:
        snapshot = self.material_assistant_snapshot(horizon_days=horizon_days)
        rows: list[dict[str, Any]] = []
        for row in list(snapshot.get("suggestions", []) or []):
            current = dict(row or {})
            shift = self._material_assistant_shift_payload(current.get("when", ""), current.get("delivery", ""))
            current["planeamento_dia"] = str(shift.get("date_label", "-") or "-")
            current["planeamento_turno"] = str(shift.get("shift_label", "Sem turno") or "Sem turno")
            current["planeamento_hora"] = str(shift.get("time_label", "-") or "-")
            rows.append(current)
        rows.sort(
            key=lambda row: (
                0 if str(row.get("kind", "") or "") == "fito_lot" else 1,
                0 if str(row.get("status_key", "") or "") == "new" else 1,
                -int(row.get("priority_score", 0) or 0),
                str(row.get("next_action_at", "") or "9999-99-99T99:99"),
                0 if str(row.get("plan_origin", "") or "").strip() == "Planeamento" else 1,
                str(row.get("delivery", "") or "9999-99-99"),
                str(row.get("numero", "") or ""),
            )
        )
        return rows

    def material_assistant_render_separation_pdf(self, horizon_days: int = 5, output_path: str | Path | None = None) -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        rows = list(self.material_assistant_separation_rows(horizon_days=horizon_days))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = Path(output_path) if output_path else Path(tempfile.gettempdir()) / f"lugest_separacao_mp_{stamp}.pdf"
        path.parent.mkdir(parents=True, exist_ok=True)
        c = pdf_canvas.Canvas(str(path), pagesize=landscape(A4))
        width, height = landscape(A4)
        margin = 24
        usable_w = width - (margin * 2)
        title = "Separação - Matéria-Prima"
        subtitle = (
            f"Horizonte {int(horizon_days or 5)} dias úteis | "
            f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )

        columns = [
            ("Prio", 34),
            ("Lote", 108),
            ("Dimensao", 86),
            ("Qtd.", 34),
            ("Planeado", 58),
            ("Hora", 36),
            ("Acao sugerida", 190),
            ("V.Sep", 34),
            ("V.Conf", 34),
        ]

        def clip(text: object, col_w: float, bold: bool = False) -> str:
            return _pdf_clip_text(str(text or "-"), max(24.0, col_w - 8), "Helvetica-Bold" if bold else "Helvetica", 7.5)

        posto_groups: list[tuple[str, list[dict[str, Any]]]] = []
        for row in rows:
            posto_label = str(row.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto"
            if posto_groups and posto_groups[-1][0] == posto_label:
                posto_groups[-1][1].append(row)
            else:
                posto_groups.append((posto_label, [row]))

        page_no = 0

        def pdf_group_key(row: dict[str, Any]) -> str:
            return "|".join(
                [
                    str(row.get("planeamento_dia_iso", "") or "9999-99-99"),
                    f"{int(row.get('planeamento_turno_ordem', 9) or 9):02d}",
                    str(row.get("numero", "") or "-").strip() or "-",
                    str(row.get("material_group", "") or "-").strip() or "-",
                ]
            )

        def pdf_group_label(row: dict[str, Any]) -> str:
            numero = str(row.get("numero", "") or "-").strip() or "-"
            material_group = str(row.get("material_group", "") or "-").strip() or "-"
            return f"{numero} | {material_group}"

        def draw_table_header(current_y: float) -> float:
            c.setFillColor(colors.HexColor("#0f172a"))
            c.roundRect(margin, current_y - 16, usable_w, 18, 6, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            x = margin
            for label, col_w in columns:
                c.drawString(x + 4, current_y - 10, label)
                x += col_w
            return current_y - 20

        def draw_header(page_rows: list[dict[str, Any]], posto_label: str) -> float:
            nonlocal page_no
            page_no += 1
            top_y = height - margin
            header_h = 48
            header_y = top_y - header_h
            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#dbe3f0"))
            c.roundRect(margin, header_y, usable_w, header_h, 12, fill=1, stroke=1)

            c.setFillColor(colors.HexColor("#0f172a"))
            c.setFont("Helvetica-Bold", 18)
            c.drawString(margin + 14, header_y + 30, title)
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.HexColor("#475569"))
            c.drawString(margin + 14, header_y + 14, subtitle)

            page_chip_w = 82
            page_chip_h = 24
            page_chip_x = width - margin - page_chip_w
            page_chip_y = header_y + header_h - page_chip_h - 10
            c.setFillColor(colors.HexColor("#f8fafc"))
            c.setStrokeColor(colors.HexColor("#dbe3f0"))
            c.roundRect(page_chip_x, page_chip_y, page_chip_w, page_chip_h, 10, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#0f172a"))
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(page_chip_x + (page_chip_w / 2), page_chip_y + 8, f"Pag. {page_no}")

            posto_card_h = 34
            posto_card_y = header_y - 10 - posto_card_h
            c.setFillColor(colors.HexColor("#0f172a"))
            c.setStrokeColor(colors.HexColor("#0f172a"))
            c.roundRect(margin, posto_card_y, usable_w, posto_card_h, 12, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#cbd5e1"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(margin + 14, posto_card_y + 22, "POSTO DE TRABALHO")
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 15)
            c.drawString(margin + 14, posto_card_y + 9, posto_label)

            summary_y = posto_card_y - 12
            info_box_w = (usable_w - 24) / 4
            group_count = len({pdf_group_key(row) for row in page_rows if pdf_group_key(row)})
            summary_cards = [
                ("Linhas", str(len(page_rows))),
                ("Grupos operacionais", str(group_count)),
                ("Materiais / esp.", str(len({str(row.get('material_group', '') or '').strip() for row in page_rows if str(row.get('material_group', '') or '').strip()}))),
                ("Posto", posto_label),
            ]
            for index, (label, value) in enumerate(summary_cards):
                x0 = margin + (index * (info_box_w + 8))
                c.setFillColor(colors.HexColor("#f8fafc"))
                c.setStrokeColor(colors.HexColor("#dbe3f0"))
                c.roundRect(x0, summary_y - 24, info_box_w, 28, 8, fill=1, stroke=1)
                c.setFillColor(colors.HexColor("#64748b"))
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(x0 + 10, summary_y - 10, label)
                c.setFillColor(colors.HexColor("#0f172a"))
                c.setFont("Helvetica-Bold", 9.5 if index == 3 else 10)
                c.drawRightString(x0 + info_box_w - 10, summary_y - 10, clip(value, info_box_w - 24, bold=True))
            y0 = summary_y - 34
            return draw_table_header(y0)

        def draw_footer_fields(current_y: float, posto_label: str, posto_rows: list[dict[str, Any]]) -> float:
            footer_height = 52
            if current_y < margin + footer_height + 6:
                c.showPage()
                current_y = draw_header(posto_rows, posto_label)
            box_y = current_y - footer_height
            box_w = (usable_w - 12) / 2
            c.setStrokeColor(colors.HexColor("#cbd5e1"))
            c.setFillColor(colors.white)
            c.roundRect(margin, box_y, box_w, footer_height, 8, fill=1, stroke=1)
            c.roundRect(margin + box_w + 12, box_y, box_w, footer_height, 8, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#0f172a"))
            c.setFont("Helvetica-Bold", 9)
            c.drawString(margin + 10, box_y + footer_height - 14, "Visto separacao")
            c.drawString(margin + box_w + 22, box_y + footer_height - 14, "Visto conferencia")
            c.setFont("Helvetica", 8)
            c.drawString(margin + 10, box_y + 14, "Nome / rubrica:")
            c.drawString(margin + box_w + 22, box_y + 14, "Nome / rubrica:")
            c.line(margin + 86, box_y + 13, margin + box_w - 12, box_y + 13)
            c.line(margin + box_w + 98, box_y + 13, margin + (box_w * 2), box_y + 13)
            return box_y - 8

        row_h = 22
        c.setFont("Helvetica", 7.0)

        if not rows:
            y = draw_header([], "Sem posto")
            c.setFillColor(colors.HexColor("#334155"))
            c.drawString(margin, y - 10, "Sem necessidades de separação de matéria-prima no horizonte atual.")
        else:
            for posto_label, posto_rows in posto_groups:
                if page_no > 0:
                    c.showPage()
                y = draw_header(posto_rows, posto_label)

                grouped_rows: list[tuple[str, list[dict[str, Any]]]] = []
                for row in posto_rows:
                    group_key = pdf_group_key(row)
                    if grouped_rows and grouped_rows[-1][0] == group_key:
                        grouped_rows[-1][1].append(row)
                    else:
                        grouped_rows.append((group_key, [row]))

                def draw_group_header(current_y: float, group_rows: list[dict[str, Any]]) -> float:
                    group_ref = dict(group_rows[0] or {})
                    group_qty = sum(self._parse_float(row.get("quantidade", 0), 0) for row in group_rows)
                    group_need_qty = sum(self._parse_float(row.get("quantidade_necessaria", row.get("quantidade", 0)), 0) for row in group_rows)
                    all_uncativated = bool(group_rows) and all(not bool(row.get("material_cativado")) for row in group_rows)
                    lotes = {
                        str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-").strip()
                        for row in group_rows
                        if str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-").strip()
                    }
                    formatos = {
                        str(row.get("dimensao", "") or "-").strip() or "-"
                        for row in group_rows
                    }
                    if current_y < margin + 54:
                        c.showPage()
                        current_y = draw_header(posto_rows, posto_label)
                    c.setFillColor(colors.HexColor("#e8eefc"))
                    c.setStrokeColor(colors.HexColor("#c4d2f3"))
                    c.roundRect(margin, current_y - 18, usable_w, 20, 8, fill=1, stroke=1)
                    c.setFillColor(colors.HexColor("#0f172a"))
                    c.setFont("Helvetica-Bold", 8.2)
                    c.drawString(margin + 8, current_y - 10, pdf_group_label(group_ref))
                    c.setFont("Helvetica", 7.0)
                    qty_summary = (
                        f"sem cativacao | nec. {self._fmt(group_need_qty)} un."
                        if all_uncativated
                        else f"{self._fmt(group_qty)} un."
                    )
                    c.drawRightString(
                        width - margin - 8,
                        current_y - 10,
                        (
                            f"{group_ref.get('planeamento_dia', '-')} | "
                            f"{group_ref.get('planeamento_turno', '-')} | "
                            f"{group_ref.get('cliente', '-')} | "
                            f"{len(formatos)} formatos | {qty_summary} | {len(lotes)} lotes"
                        ),
                    )
                    return current_y - 24

                for _group_key, group_rows in grouped_rows:
                    group_ref = dict(group_rows[0] or {})
                    y = draw_group_header(y, group_rows)
                    group_qty = sum(self._parse_float(row.get("quantidade", 0), 0) for row in group_rows)
                    lotes = {
                        str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-").strip()
                        for row in group_rows
                        if str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-").strip()
                    }

                    for row in group_rows:
                        if y < margin + 34:
                            c.showPage()
                            y = draw_header(posto_rows, posto_label)
                            y = draw_group_header(y, group_rows)

                        tone = str(row.get("priority_tone", "") or "default").strip()
                        fill = {
                            "danger": "#fff1f2",
                            "warning": "#fff8e6",
                            "success": "#ecfdf3",
                            "info": "#eef4ff",
                        }.get(tone, "#ffffff")
                        c.setFillColor(colors.HexColor(fill))
                        c.roundRect(margin, y - row_h + 2, usable_w, row_h - 2, 4, fill=1, stroke=0)
                        c.setFillColor(colors.HexColor("#0f172a"))
                        x = margin
                        values = [
                            str(row.get("priority_label", "") or "-"),
                            clip(row.get("lote_sugerido", row.get("lote_atual", "-")), columns[1][1]),
                            clip(row.get("dimensao", "-"), columns[2][1]),
                            str(row.get("quantidade_label", "") or self._fmt(row.get("quantidade", 0))),
                            clip(row.get("planeamento_dia", "-"), columns[4][1]),
                            clip(row.get("planeamento_hora", row.get("proxima_acao", "-")), columns[5][1]),
                            clip(row.get("acao_sugerida", "-"), columns[6][1]),
                        ]
                        value_map = {label: value for (label, _col_w), value in zip(columns, values)}
                        for label, col_w in columns:
                            font_name = "Helvetica-Bold" if label in {"Prio", "Lote"} else "Helvetica"
                            if label in {"V.Sep", "V.Conf"}:
                                box_size = 9
                                box_x = x + (col_w - box_size) / 2
                                box_y = y - 15
                                c.setStrokeColor(colors.HexColor("#64748b"))
                                c.rect(box_x, box_y, box_size, box_size, fill=0, stroke=1)
                                checked = bool(row.get("visto_sep_checked")) if label == "V.Sep" else bool(row.get("visto_conf_checked"))
                                if checked:
                                    c.setFont("Helvetica-Bold", 8)
                                    c.setFillColor(colors.HexColor("#0f172a"))
                                    c.drawString(box_x + 1.5, box_y + 1.2, "X")
                            elif label == "Acao sugerida":
                                c.setFont("Helvetica-Bold", 6.9)
                                c.setFillColor(colors.HexColor("#0f172a"))
                                c.drawString(x + 4, y - 9, str(value_map.get(label, "-") or "-"))
                                options_txt = clip(row.get("opcoes_mp", ""), col_w, bold=False)
                                if options_txt:
                                    c.setFont("Helvetica", 5.8)
                                    c.setFillColor(colors.HexColor("#475569"))
                                    c.drawString(x + 4, y - 17, options_txt)
                            else:
                                c.setFont(font_name, 7.0)
                                c.setFillColor(colors.HexColor("#0f172a"))
                                if label == "Qtd.":
                                    c.drawRightString(x + col_w - 4, y - 11, str(value_map.get(label, "-") or "-"))
                                else:
                                    c.drawString(x + 4, y - 11, str(value_map.get(label, "-") or "-"))
                            x += col_w
                        y -= row_h
                    y -= 6

                draw_footer_fields(y, posto_label, posto_rows)

        self._material_assistant_append_planning_page(c, width, height, rows)
        self._material_assistant_append_suggestions_page(c, width, height, horizon_days=horizon_days)
        c.save()
        return path

    def material_assistant_open_separation_pdf(self, horizon_days: int = 5) -> Path:
        path = self.material_assistant_render_separation_pdf(horizon_days=horizon_days)
        os.startfile(str(path))
        return path

    def _material_assistant_planning_week_start(self, rows: list[dict[str, Any]] | None = None) -> date:
        candidates: list[date] = []
        for row in list(rows or []) or []:
            raw = str((row or {}).get("planeamento_dia_iso", "") or "").strip()
            if not raw or raw == "9999-99-99":
                continue
            try:
                candidates.append(date.fromisoformat(raw))
            except Exception:
                continue
        anchor = min(candidates) if candidates else date.today()
        return self._planning_week_start(anchor)

    def _material_assistant_append_planning_page(
        self,
        canvas_obj: Any,
        width: float,
        height: float,
        rows: list[dict[str, Any]] | None = None,
    ) -> None:
        from reportlab.lib import colors

        planning_operation = "Corte Laser"
        week_start = self._material_assistant_planning_week_start(rows)
        week_dates = [week_start + timedelta(days=index) for index in range(6)]
        start_min, end_min, slot = self._planning_grid_metrics()
        total_slots = max(1, int((end_min - start_min) / max(1, slot)))
        margin = 20
        top_margin = 60
        time_w = 66
        footer_box_h = 58
        cols = 6
        usable_w = width - (margin * 2)
        grid_w = usable_w - time_w
        col_w = max(60, grid_w / cols)
        grid_h = height - top_margin - margin - footer_box_h - 10
        row_h = max(18, grid_h / (total_slots + 1))
        dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab"]

        def yinv(y: float) -> float:
            return height - y

        def draw_block_text(cx: float, cy: float, lines: list[Any], box_w: float, box_h: float) -> None:
            raw_specs = []
            for line in list(lines or []):
                if isinstance(line, dict):
                    text = str(line.get("text", "") or "").strip()
                    role = str(line.get("role", "body") or "body").strip()
                else:
                    text = str(line or "").strip()
                    role = "body"
                if text:
                    raw_specs.append({"text": text, "role": role})
            if not raw_specs:
                raw_specs = [{"text": "-", "role": "body"}]
            text_w = max(28.0, box_w - 10)
            for font_size in (9.2, 8.2, 7.2, 6.4, 5.8):
                line_h = font_size + 1.2
                max_lines = max(1, int((box_h - 6) // line_h))
                wrapped: list[dict[str, str]] = []
                for spec in raw_specs:
                    role = str(spec.get("role", "body") or "body")
                    font_name = "Helvetica-Bold" if role in ("title", "time") else "Helvetica"
                    for part in _pdf_wrap_text(spec.get("text", ""), font_name, font_size, text_w, max_lines=2):
                        part_txt = str(part or "").strip()
                        if part_txt:
                            wrapped.append({"text": part_txt, "role": role})
                if len(wrapped) <= max_lines:
                    total_h = len(wrapped) * line_h
                    start_y = cy - (total_h / 2.0) + (line_h / 2.0)
                    canvas_obj.setFillColor(colors.black)
                    for idx, line in enumerate(wrapped):
                        role = str(line.get("role", "body") or "body")
                        canvas_obj.setFont("Helvetica-Bold" if role in ("title", "time") else "Helvetica", font_size)
                        canvas_obj.drawCentredString(cx, yinv(start_y + (idx * line_h)), str(line.get("text", "") or "-"))
                    return

            font_size = 5.8
            line_h = font_size + 1.1
            max_lines = max(1, int((box_h - 6) // line_h))
            wrapped: list[dict[str, str]] = []
            for spec in raw_specs:
                role = str(spec.get("role", "body") or "body")
                font_name = "Helvetica-Bold" if role in ("title", "time") else "Helvetica"
                for part in _pdf_wrap_text(spec.get("text", ""), font_name, font_size, text_w, max_lines=2):
                    part_txt = str(part or "").strip()
                    if part_txt:
                        wrapped.append({"text": part_txt, "role": role})
            wrapped = wrapped[:max_lines]
            if wrapped:
                last = dict(wrapped[-1])
                font_name = "Helvetica-Bold" if last.get("role") in ("title", "time") else "Helvetica"
                text = str(last.get("text", "") or "").strip()
                while text and len(_pdf_wrap_text(f"{text}...", font_name, font_size, text_w, max_lines=1)) != 1:
                    text = text[:-1].rstrip()
                last["text"] = f"{text}..." if text else "..."
                wrapped[-1] = last
            total_h = len(wrapped or [{"text": "-", "role": "body"}]) * line_h
            start_y = cy - (total_h / 2.0) + (line_h / 2.0)
            canvas_obj.setFillColor(colors.black)
            for idx, line in enumerate(wrapped or [{"text": "-", "role": "body"}]):
                role = str(line.get("role", "body") or "body")
                canvas_obj.setFont("Helvetica-Bold" if role in ("title", "time") else "Helvetica", font_size)
                canvas_obj.drawCentredString(cx, yinv(start_y + (idx * line_h)), str(line.get("text", "") or "-"))

        canvas_obj.showPage()
        canvas_obj.setStrokeColor(colors.HexColor("#c7ccd6"))
        canvas_obj.rect(margin, yinv(height - margin), width - (margin * 2), height - (margin * 2), stroke=1, fill=0)

        canvas_obj.setFillColor(colors.HexColor("#0f172a"))
        canvas_obj.setFont("Helvetica-Bold", 15)
        canvas_obj.drawString(margin + 80, yinv(margin + 20), "Planeamento associado")
        canvas_obj.setFont("Helvetica", 9)
        canvas_obj.setFillColor(colors.HexColor("#475569"))
        canvas_obj.drawString(margin + 80, yinv(margin + 34), planning_operation)
        canvas_obj.drawRightString(
            width - margin - 4,
            yinv(margin + 20),
            f"Semana: {week_dates[0].strftime('%d/%m/%Y')} - {week_dates[-1].strftime('%d/%m/%Y')}",
        )
        canvas_obj.line(margin, yinv(55), width - margin, yinv(55))

        canvas_obj.setFont("Helvetica-Bold", 8)
        for col_idx, day in enumerate(week_dates):
            x0 = margin + time_w + (col_idx * col_w)
            canvas_obj.setFillColor(colors.HexColor("#e8eefc"))
            canvas_obj.setStrokeColor(colors.HexColor("#c4d2f3"))
            canvas_obj.rect(x0, yinv(top_margin + row_h), col_w, row_h, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.drawCentredString(x0 + (col_w / 2.0), yinv(top_margin + (row_h / 2.0)), f"{dias[col_idx]} {day.strftime('%d/%m')}")

        canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
        canvas_obj.rect(margin, yinv(top_margin + ((total_slots + 1) * row_h)), time_w, total_slots * row_h, stroke=1, fill=1)
        for row_idx in range(total_slots):
            slot_start = start_min + (row_idx * slot)
            slot_end = slot_start + slot
            y0 = top_margin + ((row_idx + 1) * row_h)
            hhmm = str(self.desktop_main.minutes_to_time(slot_start) or "")
            if hhmm.endswith(":00"):
                canvas_obj.setFont("Helvetica-Bold", 7.5)
                canvas_obj.setFillColor(colors.HexColor("#334155"))
            else:
                canvas_obj.setFont("Helvetica", 6.8)
                canvas_obj.setFillColor(colors.HexColor("#64748b"))
            canvas_obj.drawCentredString(margin + (time_w / 2.0), yinv(y0 + (row_h / 2.0)), hhmm)
            for col_idx in range(cols):
                x0 = margin + time_w + (col_idx * col_w)
                canvas_obj.setFillColor(colors.HexColor("#e5e7eb") if self._planning_interval_blocked(slot_start, slot_end) else colors.white)
                canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
                canvas_obj.rect(x0, yinv(y0 + row_h), col_w, row_h, stroke=1, fill=1)

        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(self.ensure_data().get("clientes", []) or [])
            if isinstance(row, dict)
        }
        block_color_fn = getattr(self.plan_actions, "_pdf_block_color_for_item", None)
        planning_items = [
            item
            for item in list(self.ensure_data().get("plano", []) or [])
            if isinstance(item, dict) and self._planning_row_matches_operation(item, planning_operation)
        ]
        for item in planning_items:
            raw_date = str(item.get("data", "") or "").strip()
            raw_start = str(item.get("inicio", "") or "").strip()
            if not raw_date or not raw_start:
                continue
            try:
                block_day = date.fromisoformat(raw_date)
                start_value = self.desktop_main.time_to_minutes(raw_start)
            except Exception:
                continue
            if block_day < week_dates[0] or block_day > week_dates[-1]:
                continue
            duration = self._planning_round_duration(item.get("duracao_min", 0))
            row_start = int((start_value - start_min) // max(1, slot))
            row_span = max(1, int(math.ceil(float(duration) / float(max(1, slot)))))
            if row_start < 0 or row_start >= total_slots:
                continue
            col_idx = (block_day - week_dates[0]).days
            x0 = margin + time_w + (col_idx * col_w)
            y0 = top_margin + ((row_start + 1) * row_h)
            y1 = y0 + (row_span * row_h)
            block_color = str(block_color_fn(item) or "").strip() if callable(block_color_fn) else ""
            if not block_color:
                block_color = str(item.get("color", "") or self._planning_item_color(item.get("encomenda", ""), item.get("material", ""), item.get("espessura", ""))).strip()
            fill_hex = _pdf_mix_hex(block_color or "#c7d2fe", "#ffffff", 0.30)
            edge_hex = _pdf_mix_hex(block_color or "#1f3c88", "#0f172a", 0.20)
            canvas_obj.setFillColor(colors.HexColor(fill_hex))
            canvas_obj.setStrokeColor(colors.HexColor(edge_hex))
            canvas_obj.rect(x0 + 2, yinv(y1 - 2), col_w - 4, (y1 - y0) - 4, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor(block_color or "#1f3c88"))
            canvas_obj.rect(x0 + 2, yinv(y1 - 2), 6, (y1 - y0) - 4, stroke=0, fill=1)

            enc_num = str(item.get("encomenda", "") or "").strip()
            enc = self.get_encomenda_by_numero(enc_num) or {}
            cliente_txt = clients.get(str(enc.get("cliente", "") or "").strip(), "")
            mat = str(item.get("material", "") or "").strip()
            esp = str(item.get("espessura", "") or "").strip()
            fim_txt = self.desktop_main.minutes_to_time(start_value + duration)
            block_h = (y1 - y0) - 6
            mat_esp = " | ".join(part for part in (mat, f"{esp} mm" if esp else "") if part).strip()
            tempo_txt = f"{duration} min"
            if block_h <= (row_h * 1.2):
                lines = [{"text": enc_num or "-", "role": "title"}, {"text": tempo_txt, "role": "time"}]
                if mat_esp and block_h > (row_h * 0.9):
                    lines.append({"text": mat_esp, "role": "body"})
            elif block_h <= (row_h * 1.9):
                lines = [{"text": enc_num or "-", "role": "title"}]
                if mat_esp:
                    lines.append({"text": mat_esp, "role": "body"})
                lines.append({"text": f"{raw_start} - {fim_txt} | {tempo_txt}", "role": "time"})
            else:
                lines = [{"text": enc_num or "-", "role": "title"}]
                if cliente_txt and block_h >= (row_h * 1.8):
                    lines.append({"text": f"Cliente: {cliente_txt}", "role": "body"})
                if mat_esp:
                    lines.append({"text": mat_esp, "role": "body"})
                lines.append({"text": f"{raw_start} - {fim_txt}", "role": "body"})
                lines.append({"text": tempo_txt, "role": "time"})
            if cliente_txt and block_h >= (row_h * 2.4) and lines and all("Cliente:" not in line for line in lines):
                lines.append({"text": f"Cliente: {cliente_txt}", "role": "body"})
            chapa = str(item.get("chapa", "") or "").strip()
            if chapa and chapa != "-" and block_h >= (row_h * 2.6):
                lines.append({"text": f"Chapa: {chapa}", "role": "body"})
            draw_block_text(x0 + (col_w / 2.0), y0 + ((y1 - y0) / 2.0), lines, col_w - 10, (y1 - y0) - 6)

        box_y = height - margin - footer_box_h
        canvas_obj.setStrokeColor(colors.HexColor("#cbd5e1"))
        canvas_obj.rect(margin, yinv(box_y + footer_box_h), width - (margin * 2), footer_box_h, stroke=1, fill=0)
        canvas_obj.setFillColor(colors.HexColor("#0f172a"))
        canvas_obj.setFont("Helvetica-Bold", 8)
        canvas_obj.drawString(margin + 6, yinv(box_y + 14), "Observacoes:")
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.line(margin + 74, yinv(box_y + 15), width - margin - 6, yinv(box_y + 15))
        canvas_obj.setFont("Helvetica-Bold", 8)
        canvas_obj.drawString(margin + 6, yinv(box_y + 36), "Data:")
        canvas_obj.drawString(margin + 180, yinv(box_y + 36), "Operador:")

    def _material_assistant_append_suggestions_page(
        self,
        canvas_obj: Any,
        width: float,
        height: float,
        *,
        horizon_days: int = 5,
    ) -> None:
        from reportlab.lib import colors

        rows = [
            dict(row or {})
            for row in list(self.material_assistant_alert_rows(horizon_days=horizon_days) or [])
            if str((row or {}).get("status_key", "") or "").strip().lower() != "ignored"
        ]

        kind_order = {
            "fito_lot": 0,
            "keep_ready": 1,
            "uncativated": 2,
            "shortage": 3,
        }
        rows.sort(
            key=lambda row: (
                int(kind_order.get(str(row.get("kind", "") or "").strip(), 9)),
                0 if str(row.get("status_key", "") or "").strip() == "new" else 1,
                -int(row.get("priority_score", 0) or 0),
                str(row.get("next_action_at", "") or "9999-99-99T99:99"),
                str(row.get("numero", "") or ""),
            )
        )

        margin = 24
        usable_w = width - (margin * 2)
        page_no = 0

        def _kind_label(row: dict[str, Any]) -> str:
            kind = str(row.get("kind", "") or "").strip()
            mapping = {
                "fito_lot": "Troca por urgencia / FIFO",
                "keep_ready": "Nao arrumar / manter pronto",
                "uncativated": "Sem material cativado",
                "shortage": "Sem stock",
            }
            return mapping.get(kind, "Sugestao operacional")

        def _kind_color(row: dict[str, Any]) -> str:
            kind = str(row.get("kind", "") or "").strip()
            mapping = {
                "fito_lot": "#0f3d91",
                "keep_ready": "#b45309",
                "uncativated": "#b45309",
                "shortage": "#b91c1c",
            }
            return mapping.get(kind, "#334155")

        def _draw_page_header() -> float:
            nonlocal page_no
            page_no += 1
            canvas_obj.showPage()
            top_y = height - margin
            header_h = 56
            header_y = top_y - header_h
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(margin, header_y, usable_w, header_h, 12, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 17)
            canvas_obj.drawString(margin + 14, header_y + 34, "Sugestoes recomendadas")
            canvas_obj.setFont("Helvetica", 9)
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.drawString(
                margin + 14,
                header_y + 16,
                "Baseado no planeamento atual e na prioridade FIFO. Serve de apoio ao operador; nao altera nada automaticamente.",
            )
            page_chip_w = 82
            page_chip_h = 24
            page_chip_x = width - margin - page_chip_w
            page_chip_y = header_y + header_h - page_chip_h - 10
            canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(page_chip_x, page_chip_y, page_chip_w, page_chip_h, 10, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.drawCentredString(page_chip_x + (page_chip_w / 2), page_chip_y + 8, f"Pag. {page_no}")
            info_y = header_y - 8
            info_h = 32
            info_box_w = (usable_w - 16) / 3
            cards = [
                ("Sugestoes", str(len(rows))),
                ("Trocas FIFO", str(len([row for row in rows if str(row.get("kind", "") or "") == "fito_lot"]))),
                ("Horizonte", f"{int(horizon_days or 5)} dias úteis"),
            ]
            for index, (label, value) in enumerate(cards):
                x0 = margin + (index * (info_box_w + 8))
                canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
                canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
                canvas_obj.roundRect(x0, info_y - info_h, info_box_w, info_h, 9, fill=1, stroke=1)
                canvas_obj.setFillColor(colors.HexColor("#64748b"))
                canvas_obj.setFont("Helvetica-Bold", 7.5)
                canvas_obj.drawString(x0 + 10, info_y - 11, label)
                canvas_obj.setFillColor(colors.HexColor("#0f172a"))
                canvas_obj.setFont("Helvetica-Bold", 10)
                canvas_obj.drawRightString(x0 + info_box_w - 10, info_y - 11, str(value or "-"))
            return info_y - info_h - 10

        def _draw_empty_page() -> None:
            y = _draw_page_header()
            canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(margin, y - 74, usable_w, 64, 14, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 13)
            canvas_obj.drawString(margin + 16, y - 32, "Sem sugestoes operacionais neste horizonte")
            canvas_obj.setFont("Helvetica", 9)
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.drawString(
                margin + 16,
                y - 50,
                "Nao existem alertas de troca FIFO, falta de stock ou manutencao de material pronto para os proximos dias.",
            )

        def _draw_row(y: float, row: dict[str, Any]) -> float:
            card_h = 74
            if y < margin + card_h + 8:
                y = _draw_page_header()
            color_hex = _kind_color(row)
            fill_hex = _pdf_mix_hex(color_hex, "#ffffff", 0.90)
            edge_hex = _pdf_mix_hex(color_hex, "#dbeafe", 0.35)
            canvas_obj.setFillColor(colors.HexColor(fill_hex))
            canvas_obj.setStrokeColor(colors.HexColor(edge_hex))
            canvas_obj.roundRect(margin, y - card_h, usable_w, card_h - 2, 12, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor(color_hex))
            canvas_obj.roundRect(margin, y - card_h, 8, card_h - 2, 8, fill=1, stroke=0)

            header_y = y - 16
            left_x = margin + 16
            right_x = width - margin - 16

            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 10)
            canvas_obj.drawString(left_x, header_y, _pdf_clip_text(_kind_label(row), usable_w * 0.42, "Helvetica-Bold", 10))
            canvas_obj.setFont("Helvetica-Bold", 8.5)
            canvas_obj.drawRightString(right_x, header_y, _pdf_clip_text(str(row.get("priority_label", "") or "Media"), usable_w * 0.16, "Helvetica-Bold", 8.5))

            meta_txt = " | ".join(
                part
                for part in (
                    str(row.get("numero", "") or "-").strip() or "-",
                    str(row.get("cliente", "") or "-").strip() or "-",
                    str(row.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto",
                )
                if part
            )
            canvas_obj.setFont("Helvetica", 8.1)
            canvas_obj.setFillColor(colors.HexColor("#334155"))
            canvas_obj.drawString(left_x, header_y - 13, _pdf_clip_text(meta_txt, usable_w - 32, "Helvetica", 8.1))

            recommendation = str(row.get("recommendation", "") or "").strip()
            if not recommendation:
                recommendation = str(row.get("headline", "") or "").strip() or "Rever sugestao operacional."
            detail_parts = [
                " | ".join(
                    part
                    for part in (
                        str(row.get("material", "") or "-").strip(),
                        f"{str(row.get('espessura', '') or '').strip()} mm".strip() if str(row.get("espessura", "") or "").strip() else "",
                    )
                    if part
                ).strip(" |"),
                "Planeado: " + " ".join(
                    part
                    for part in (
                        str(row.get("planeamento_dia", "") or "").strip(),
                        str(row.get("planeamento_turno", "") or "").strip(),
                        str(row.get("planeamento_hora", "") or "").strip(),
                    )
                    if part and part != "-"
                ).strip(),
                f"Estado: {str(row.get('status_label', '') or 'Nova').strip()}",
            ]
            detail_lines = list(row.get("detail_lines", []) or [])
            if detail_lines:
                detail_parts.append(str(detail_lines[0] or "").strip())

            text_y = y - 44
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 8.6)
            for wrapped in _pdf_wrap_text(recommendation, "Helvetica-Bold", 8.6, usable_w - 32, max_lines=2):
                canvas_obj.drawString(left_x, text_y, wrapped)
                text_y -= 10
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.setFont("Helvetica", 7.4)
            detail_text = " | ".join(part for part in detail_parts if str(part or "").strip())
            for wrapped in _pdf_wrap_text(detail_text, "Helvetica", 7.4, usable_w - 32, max_lines=2):
                canvas_obj.drawString(left_x, text_y, wrapped)
                text_y -= 8
            return y - card_h - 8

        if not rows:
            _draw_empty_page()
            return

        y = _draw_page_header()
        for row in rows:
            y = _draw_row(y, row)

    def _order_is_orc_based(self, enc: dict[str, Any]) -> bool:
        return bool(str((enc or {}).get("numero_orcamento", "") or "").strip())

    def _order_find_material(self, enc: dict[str, Any], material: str) -> dict[str, Any] | None:
        material_txt = str(material or "").strip().lower()
        for row in list(enc.get("materiais", []) or []):
            if str(row.get("material", "") or "").strip().lower() == material_txt:
                return row
        return None

    def _order_find_espessura(self, enc: dict[str, Any], material: str, espessura: str) -> dict[str, Any] | None:
        mat = self._order_find_material(enc, material)
        if mat is None:
            return None
        esp_txt = str(espessura or "").strip()
        for row in list(mat.get("espessuras", []) or []):
            if str(row.get("espessura", "") or "").strip() == esp_txt:
                return row
        return None

    def _order_find_piece(self, enc: dict[str, Any], ref_interna: str, ref_externa: str = "") -> tuple[dict[str, Any] | None, dict[str, Any] | None, dict[str, Any] | None]:
        ref_int = str(ref_interna or "").strip()
        ref_ext = str(ref_externa or "").strip()
        for mat in list(enc.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                for piece in list(esp.get("pecas", []) or []):
                    piece_ref_int = str(piece.get("ref_interna", "") or "").strip()
                    piece_ref_ext = str(piece.get("ref_externa", "") or "").strip()
                    if ref_int and piece_ref_int == ref_int:
                        return mat, esp, piece
                    if (not ref_int) and ref_ext and piece_ref_ext == ref_ext:
                        return mat, esp, piece
        return None, None, None

    def order_rows(self, filter_text: str = "", estado: str = "Ativas", ano: str = "Todos", cliente: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Ativas").strip().lower()
        ano_filter = str(ano or "Todos").strip()
        cliente_filter = str(cliente or "Todos").strip()
        clientes_nome = {
            str(c.get("codigo", "") or "").strip(): str(c.get("nome", "") or "").strip()
            for c in list(data.get("clientes", []) or [])
            if isinstance(c, dict)
        }
        rows = []
        for enc in data.get("encomendas", []):
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            planeado = sum(self._parse_float(p.get("quantidade_pedida", 0), 0) for p in pieces)
            produzido = sum(
                self._parse_float(p.get("produzido_ok", 0), 0)
                + self._parse_float(p.get("produzido_nok", 0), 0)
                + self._parse_float(p.get("produzido_qualidade", 0), 0)
                for p in pieces
            )
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "")
            progress = round((produzido / planeado) * 100.0, 1) if planeado > 0 else (100.0 if montagem_estado == "Consumida" else 0.0)
            estado_txt = str(enc.get("estado", "") or "").strip()
            estado_norm = self.desktop_main.norm_text(estado_txt)
            enc_year = ""
            try:
                enc_year = str(
                    self.desktop_main._enc_extract_year(
                        enc.get("data_criacao", ""),
                        enc.get("data_entrega", ""),
                        enc.get("numero", ""),
                        enc.get("ano"),
                    )
                    or ""
                ).strip()
            except Exception:
                enc_year = ""
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_display = f"{cli_code} - {clientes_nome.get(cli_code, '')}".strip(" -")
            if ano_filter.lower() not in ("todos", "todas", "all", "") and enc_year != ano_filter:
                continue
            if cliente_filter.lower() not in ("todos", "todas", "all", "") and cli_code != cliente_filter.split(" - ", 1)[0].strip():
                continue
            if estado_filter not in ("todos", "todas", "all", ""):
                if "ativ" in estado_filter and "concl" in estado_norm:
                    continue
                if "prepar" in estado_filter and "prepar" not in estado_norm:
                    continue
                if "montag" in estado_filter and "montag" not in estado_norm:
                    continue
                if "produ" in estado_filter and "produ" not in estado_norm:
                    continue
                if "concl" in estado_filter and "concl" not in estado_norm:
                    continue
            row = {
                "numero": str(enc.get("numero", "")).strip(),
                "nota_cliente": str(enc.get("nota_cliente", "") or "").strip(),
                "cliente": cli_display or cli_code or "-",
                "cliente_codigo": cli_code,
                "posto_trabalho": self._order_workcenter(enc),
                "data_criacao": str(enc.get("data_criacao", "") or "").strip(),
                "data_entrega": str(enc.get("data_entrega", "")).strip(),
                "tempo": self._fmt(enc.get("tempo_estimado", 0)),
                "estado": estado_txt,
                "cativar": "SIM" if bool(enc.get("cativar")) else "NAO",
                "pecas": len(pieces),
                "materiais": len(enc.get("materiais", []) or []),
                "planeado": self._fmt(planeado),
                "produzido": self._fmt(produzido),
                "progress": progress,
                "ano": enc_year,
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: item["numero"])
        return rows

    def _find_piece_by_opp(self, opp: str) -> tuple[dict[str, Any], dict[str, Any]]:
        target = str(opp or "").strip()
        if not target:
            raise ValueError("OPP obrigatoria.")
        data = self.ensure_data()
        changed = False
        for enc in list(data.get("encomendas", []) or []):
            for piece in self.desktop_main.encomenda_pecas(enc):
                if not str(piece.get("opp", "") or "").strip():
                    piece["opp"] = self._next_order_opp_codigo(enc)
                    changed = True
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self._order_of_code(enc, create=True)
                    changed = True
                if str(piece.get("opp", "") or "").strip() == target:
                    if changed:
                        self._save(force=True)
                    return enc, piece
        if changed:
            self._save(force=True)
        raise ValueError("OPP n?o encontrada.")

    def _opp_rows_base(self) -> list[dict[str, Any]]:
        data = self.ensure_data()
        cliente_nome = {
            str(c.get("codigo", "") or "").strip(): str(c.get("nome", "") or "").strip()
            for c in list(data.get("clientes", []) or [])
            if isinstance(c, dict)
        }
        rows: list[dict[str, Any]] = []
        changed = False
        for enc in list(data.get("encomendas", []) or []):
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_display = f"{cli_code} - {cliente_nome.get(cli_code, '')}".strip(" -")
            enc_year = ""
            try:
                enc_year = str(
                    self.desktop_main._enc_extract_year(
                        enc.get("data_criacao", ""),
                        enc.get("data_entrega", ""),
                        enc.get("numero", ""),
                        enc.get("ano"),
                    )
                    or ""
                ).strip()
            except Exception:
                enc_year = ""
            if not enc_year:
                raw_delivery = str(enc.get("data_entrega", "") or "").strip()
                if len(raw_delivery) >= 4 and raw_delivery[:4].isdigit():
                    enc_year = raw_delivery[:4]
                else:
                    enc_year = str(datetime.now().year)
            for piece in self.desktop_main.encomenda_pecas(enc):
                if not str(piece.get("opp", "") or "").strip():
                    piece["opp"] = self.desktop_main.next_opp_numero(data)
                    changed = True
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self.desktop_main.next_of_numero(data)
                    changed = True
                ops = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
                qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
                qty_ok = self._parse_float(piece.get("produzido_ok", 0), 0)
                qty_nok = self._parse_float(piece.get("produzido_nok", 0), 0)
                qty_qual = self._parse_float(piece.get("produzido_qualidade", 0), 0)
                qty_prod = qty_ok + qty_nok + qty_qual
                qty_exp = self._parse_float(piece.get("qtd_expedida", 0), 0)
                progress = round((qty_prod / qty_plan) * 100.0, 1) if qty_plan > 0 else 0.0
                running_ops = [op for op in ops if "produ" in self.desktop_main.norm_text(op.get("estado", ""))]
                pending_ops = [op for op in ops if "concl" not in self.desktop_main.norm_text(op.get("estado", ""))]
                current_op = ""
                if running_ops:
                    current_op = self.desktop_main.normalize_operacao_nome(running_ops[0].get("nome", ""))
                elif pending_ops:
                    current_op = self.desktop_main.normalize_operacao_nome(pending_ops[0].get("nome", ""))
                elif ops:
                    current_op = self.desktop_main.normalize_operacao_nome(ops[-1].get("nome", ""))
                current_operator = ""
                if running_ops:
                    current_operator = str(running_ops[0].get("user", "") or "").strip()
                if not current_operator:
                    hist_rows = list(piece.get("hist", []) or [])
                    if hist_rows:
                        current_operator = str(hist_rows[-1].get("user", "") or "").strip()
                tempo_real = self._parse_float(piece.get("tempo_producao_min", 0), 0)
                if tempo_real <= 0 and piece.get("inicio_producao") and not piece.get("fim_producao"):
                    tempo_real = self._parse_float(
                        self.desktop_main.iso_diff_minutes(piece.get("inicio_producao"), self.desktop_main.now_iso()),
                        0,
                    )
                ops_total = len([op for op in ops if str(op.get("nome", "") or "").strip()])
                ops_done = len([op for op in ops if self.desktop_main.operacao_esta_concluida(piece, op)])
                expedicoes = [str(num or "").strip() for num in list(piece.get("expedicoes", []) or []) if str(num or "").strip()]
                rows.append(
                    {
                        "opp": str(piece.get("opp", "") or "").strip(),
                        "of": str(piece.get("of", "") or "").strip(),
                        "piece_id": str(piece.get("id", "") or "").strip(),
                        "encomenda": str(enc.get("numero", "") or "").strip(),
                        "cliente": cli_display or cli_code or "-",
                        "cliente_codigo": cli_code,
                        "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                        "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                        "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                        "material": str(piece.get("material", "") or "").strip(),
                        "espessura": str(piece.get("espessura", "") or "").strip(),
                        "estado": str(piece.get("estado", "") or "").strip(),
                        "operacao_atual": current_op or "-",
                        "operador_atual": current_operator or "-",
                        "qtd_plan": qty_plan,
                        "qtd_prod": qty_prod,
                        "qtd_exp": qty_exp,
                        "progress": progress,
                        "tempo_real": round(tempo_real, 2),
                        "ops_total": ops_total,
                        "ops_done": ops_done,
                        "ops_pending": max(0, ops_total - ops_done),
                        "inicio": str(piece.get("inicio_producao", "") or "").strip(),
                        "fim": str(piece.get("fim_producao", "") or "").strip(),
                        "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                        "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                        "ano": enc_year,
                        "expedicoes": expedicoes,
                    }
                )
        if changed:
            self._save(force=True)
        rows.sort(key=lambda item: (item.get("opp", ""), item.get("encomenda", ""), item.get("ref_interna", "")))
        return rows

    def opp_rows(
        self,
        filter_text: str = "",
        estado: str = "Ativas",
        ano: str = "Todos",
        operacao: str = "Todas",
        cliente: str = "Todos",
    ) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Ativas").strip().lower()
        ano_filter = str(ano or "Todos").strip()
        operacao_filter = str(operacao or "Todas").strip().lower()
        cliente_filter = str(cliente or "Todos").strip()
        rows: list[dict[str, Any]] = []
        for row in self._opp_rows_base():
            estado_norm = self.desktop_main.norm_text(row.get("estado", ""))
            if ano_filter.lower() not in ("todos", "todas", "all", "") and str(row.get("ano", "") or "").strip() != ano_filter:
                continue
            if cliente_filter.lower() not in ("todos", "todas", "all", ""):
                cliente_codigo = str(row.get("cliente_codigo", "") or "").strip()
                if cliente_codigo != cliente_filter.split(" - ", 1)[0].strip():
                    continue
            if operacao_filter not in ("todas", "todos", "all", ""):
                if operacao_filter not in self.desktop_main.norm_text(row.get("operacao_atual", "")):
                    continue
            if estado_filter not in ("todos", "todas", "all", ""):
                if "ativ" in estado_filter and "concl" in estado_norm and float(row.get("qtd_exp", 0) or 0) >= float(row.get("qtd_prod", 0) or 0):
                    continue
                elif "ativ" in estado_filter and "concl" in estado_norm:
                    continue
                if "prepar" in estado_filter and "prepar" not in estado_norm:
                    continue
                if ("curso" in estado_filter or "produc" in estado_filter) and ("produ" not in estado_norm and "incomplet" not in estado_norm):
                    continue
                if "concl" in estado_filter and "concl" not in estado_norm:
                    continue
                if "exped" in estado_filter and float(row.get("qtd_exp", 0) or 0) <= 0:
                    continue
                if "avaria" in estado_filter and "avari" not in estado_norm:
                    continue
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def opp_operations(self) -> list[str]:
        ops: set[str] = set()
        for row in self._opp_rows_base():
            op = str(row.get("operacao_atual", "") or "").strip()
            if op and op != "-":
                ops.add(op)
        return sorted(ops)

    def opp_detail(self, opp: str) -> dict[str, Any]:
        enc, piece = self._find_piece_by_opp(opp)
        ops = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
        qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        qty_ok = self._parse_float(piece.get("produzido_ok", 0), 0)
        qty_nok = self._parse_float(piece.get("produzido_nok", 0), 0)
        qty_qual = self._parse_float(piece.get("produzido_qualidade", 0), 0)
        qty_prod = qty_ok + qty_nok + qty_qual
        qty_exp = self._parse_float(piece.get("qtd_expedida", 0), 0)
        progress = round((qty_prod / qty_plan) * 100.0, 1) if qty_plan > 0 else 0.0
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        if cliente_codigo:
            try:
                cliente_obj = self.desktop_main.find_cliente(self.ensure_data(), cliente_codigo) or {}
            except Exception:
                cliente_obj = {}
        op_rows: list[dict[str, Any]] = []
        for op in ops:
            nome = self.desktop_main.normalize_operacao_nome(op.get("nome", ""))
            capacidade = self.desktop_main.operacao_input_qtd(piece, nome) if nome else 0.0
            qtd_total = self.desktop_main.operacao_qtd_total(op, fallback_done=capacidade)
            op_progress = round((qtd_total / capacidade) * 100.0, 1) if capacidade > 0 else 0.0
            op_rows.append(
                {
                    "nome": nome or "-",
                    "estado": str(op.get("estado", "") or "").strip() or "Pendente",
                    "user": str(op.get("user", "") or "").strip(),
                    "inicio": str(op.get("inicio", "") or "").replace("T", " ")[:19],
                    "fim": str(op.get("fim", "") or "").replace("T", " ")[:19],
                    "qtd_ok": self._fmt(op.get("qtd_ok", 0)),
                    "qtd_nok": self._fmt(op.get("qtd_nok", 0)),
                    "qtd_qual": self._fmt(op.get("qtd_qual", 0)),
                    "capacidade": self._fmt(capacidade),
                    "progress": op_progress,
                }
            )
        event_rows: list[dict[str, Any]] = []
        target_piece_id = str(piece.get("id", "") or "").strip()
        target_ref = str(piece.get("ref_interna", "") or "").strip()
        target_enc = str(enc.get("numero", "") or "").strip()
        for ev in list(self.ensure_data().get("op_eventos", []) or []):
            if not isinstance(ev, dict):
                continue
            ev_piece = str(ev.get("peca_id", "") or "").strip()
            ev_ref = str(ev.get("ref_interna", "") or "").strip()
            ev_enc = str(ev.get("encomenda_numero", "") or "").strip()
            if target_piece_id and ev_piece == target_piece_id:
                pass
            elif target_ref and ev_ref == target_ref and ev_enc == target_enc:
                pass
            else:
                continue
            event_rows.append(
                {
                    "data": str(ev.get("created_at", "") or "").replace("T", " ")[:19],
                    "evento": str(ev.get("evento", "") or "").strip(),
                    "operacao": str(ev.get("operacao", "") or "").strip(),
                    "operador": str(ev.get("operador", "") or "").strip(),
                    "qtd_ok": self._fmt(ev.get("qtd_ok", 0)),
                    "qtd_nok": self._fmt(ev.get("qtd_nok", 0)),
                    "info": str(ev.get("info", "") or "").strip(),
                }
            )
        for ev in list(piece.get("hist", []) or []):
            if not isinstance(ev, dict):
                continue
            event_rows.append(
                {
                    "data": str(ev.get("ts", "") or "").replace("T", " ")[:19],
                    "evento": str(ev.get("acao", "") or "").strip(),
                    "operacao": " + ".join(str(item or "").strip() for item in list(ev.get("operacoes", []) or []) if str(item or "").strip()),
                    "operador": str(ev.get("user", "") or "").strip(),
                    "qtd_ok": self._fmt(ev.get("ok", 0)),
                    "qtd_nok": self._fmt(ev.get("nok", 0)),
                    "info": str(ev.get("motivo", "") or ev.get("inicio", "") or "").strip(),
                }
            )
        event_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        exp_rows: list[dict[str, Any]] = []
        for ex in list(self.ensure_data().get("expedicoes", []) or []):
            if not isinstance(ex, dict):
                continue
            for line in list(ex.get("linhas", []) or []):
                if str(line.get("peca_id", "") or "").strip() != target_piece_id:
                    continue
                exp_rows.append(
                    {
                        "guia": str(ex.get("numero", "") or "").strip(),
                        "data": str(ex.get("data_transporte", "") or ex.get("data_emissao", "") or "").replace("T", " ")[:19],
                        "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
                        "destinatario": str(ex.get("destinatario", "") or "").strip(),
                        "qtd": self._fmt(line.get("qtd", 0)),
                        "obs": str(ex.get("observacoes", "") or "").strip(),
                    }
                )
        exp_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        return {
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "piece_id": target_piece_id,
            "encomenda": target_enc,
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "ref_interna": target_ref,
            "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
            "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
            "material": str(piece.get("material", "") or "").strip(),
            "espessura": str(piece.get("espessura", "") or "").strip(),
            "estado": str(piece.get("estado", "") or "").strip(),
            "operacoes": op_rows,
            "events": event_rows,
            "expedicoes": exp_rows,
            "qtd_plan": self._fmt(qty_plan),
            "qtd_prod": self._fmt(qty_prod),
            "qtd_exp": self._fmt(qty_exp),
            "qtd_ok": self._fmt(qty_ok),
            "qtd_nok": self._fmt(qty_nok),
            "qtd_qual": self._fmt(qty_qual),
            "progress": progress,
            "tempo_real": self._fmt(piece.get("tempo_producao_min", 0)),
            "inicio": str(piece.get("inicio_producao", "") or "").replace("T", " ")[:19],
            "fim": str(piece.get("fim_producao", "") or "").replace("T", " ")[:19],
            "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
            "expedida": float(qty_exp) > 0,
        }

    def opp_open_drawing(self, opp: str) -> str:
        enc, piece = self._find_piece_by_opp(opp)
        return self.operator_open_drawing(str(enc.get("numero", "") or "").strip(), str(piece.get("id", "") or "").strip())

    def opp_open_pdf(self, opp: str) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        enc, piece = self._find_piece_by_opp(opp)
        source_posto = self._operator_posto_for_operation(str(piece.get("operacao_atual", "") or "").strip()) or "Geral"
        row = self._operator_label_row(enc, piece, source_posto=source_posto)
        target = self._operator_label_tmp_path(str(enc.get("numero", "") or "").strip(), "opp_label")
        width, height = (110 * mm, 50 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(width, height))
        self._draw_operator_unit_label(canvas_obj, width, height, row, palette, logo_path, printed_at)
        canvas_obj.save()
        os.startfile(str(target))
        return target

    def order_clients(self) -> list[dict[str, str]]:
        rows = []
        for client in list(self.ensure_data().get("clientes", []) or []):
            codigo = str(client.get("codigo", "") or "").strip()
            nome = str(client.get("nome", "") or "").strip()
            if not codigo:
                continue
            rows.append({"codigo": codigo, "nome": nome, "label": f"{codigo} - {nome}".strip(" -")})
        rows.sort(key=lambda item: item["codigo"])
        return rows

    def client_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("clientes", []) or []):
            row = {
                "codigo": str(raw.get("codigo", "") or "").strip(),
                "nome": str(raw.get("nome", "") or "").strip(),
                "nif": str(raw.get("nif", "") or "").strip(),
                "morada": str(raw.get("morada", "") or "").strip(),
                "contacto": str(raw.get("contacto", "") or "").strip(),
                "email": str(raw.get("email", "") or "").strip(),
                "observacoes": str(raw.get("observacoes", "") or "").strip(),
                "prazo_entrega": str(raw.get("prazo_entrega", "") or "").strip(),
                "cond_pagamento": str(raw.get("cond_pagamento", "") or "").strip(),
                "obs_tecnicas": str(raw.get("obs_tecnicas", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("nome") or ""))
        return rows

    def client_next_code(self) -> str:
        return str(self.desktop_main.next_cliente_codigo(self.ensure_data()))

    def client_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        codigo = str(payload.get("codigo", "") or "").strip() or str(self.desktop_main.next_cliente_codigo(data))
        nome = str(payload.get("nome", "") or "").strip()
        if not nome:
            raise ValueError("Nome do cliente obrigatorio.")
        row = {
            "codigo": codigo,
            "nome": nome,
            "nif": str(payload.get("nif", "") or "").strip(),
            "morada": str(payload.get("morada", "") or "").strip(),
            "contacto": str(payload.get("contacto", "") or "").strip(),
            "email": str(payload.get("email", "") or "").strip(),
            "observacoes": str(payload.get("observacoes", "") or "").strip(),
            "prazo_entrega": str(payload.get("prazo_entrega", "") or "").strip(),
            "cond_pagamento": str(payload.get("cond_pagamento", "") or "").strip(),
            "obs_tecnicas": str(payload.get("obs_tecnicas", "") or "").strip(),
        }
        rows = data.setdefault("clientes", [])
        existing = next((item for item in rows if str(item.get("codigo", "") or "").strip() == codigo), None)
        if existing is None:
            rows.append(row)
            target = row
        else:
            existing.update(row)
            target = existing
        self._save(force=True)
        return dict(target)

    def client_remove(self, codigo: str) -> None:
        data = self.ensure_data()
        code = str(codigo or "").strip()
        if not code:
            raise ValueError("Cliente inv?lido.")
        if any(str(enc.get("cliente", "") or "").strip() == code for enc in list(data.get("encomendas", []) or [])):
            raise ValueError("Nao e possivel remover um cliente usado em encomendas.")
        if any(str(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", "") or "").strip() == code for orc in list(data.get("orcamentos", []) or [])):
            raise ValueError("Nao e possivel remover um cliente usado em orcamentos.")
        before = len(list(data.get("clientes", []) or []))
        data["clientes"] = [row for row in list(data.get("clientes", []) or []) if str(row.get("codigo", "") or "").strip() != code]
        if len(data["clientes"]) == before:
            raise ValueError("Cliente n?o encontrado.")
        self._save(force=True)

    def order_presets(self) -> dict[str, Any]:
        data = self.ensure_data()
        materiais = list(
            dict.fromkeys(
                list(self.desktop_main.MATERIAIS_PRESET)
                + list(data.get("materiais_hist", []) or [])
                + [str(row.get("material", "") or "").strip() for row in list(data.get("materiais", []) or []) if str(row.get("material", "") or "").strip()]
            )
        )
        espessuras = list(
            dict.fromkeys(
                [self._fmt(v) for v in list(self.desktop_main.ESPESSURAS_PRESET)]
                + [str(value).strip() for value in list(data.get("espessuras_hist", []) or []) if str(value).strip()]
                + [str(row.get("espessura", "") or "").strip() for row in list(data.get("materiais", []) or []) if str(row.get("espessura", "") or "").strip()]
            )
        )
        return {
            "materiais": materiais,
            "espessuras": espessuras,
            "operacoes": list(dict.fromkeys(list(self.desktop_main.OFF_OPERACOES_DISPONIVEIS) + list(self.planning_operation_options()))),
            "operacao_default": str(self.desktop_main.OFF_OPERACAO_OBRIGATORIA),
        }

    def quote_parse_operacoes_lista(self, value: Any) -> list[str]:
        ops: list[str] = []
        items: list[Any]
        if isinstance(value, list):
            items = list(value)
        elif isinstance(value, dict):
            items = [value]
        else:
            txt = str(value or "").strip()
            items = [token for token in re.split(r"[+,;|/\n]+", txt) if str(token or "").strip()] if txt else []
        for item in items:
            if isinstance(item, dict):
                raw_name = item.get("nome") or item.get("operacao")
            else:
                raw_name = item
            normalized = str(self.desktop_main.normalize_operacao_nome(raw_name) or raw_name or "").strip()
            if normalized and normalized not in ops:
                ops.append(normalized)
        ordered_base = list(dict.fromkeys(list(self.desktop_main.OFF_OPERACOES_DISPONIVEIS) + list(self.planning_operation_options())))
        ordered = [op_name for op_name in ordered_base if op_name in ops]
        for op_name in ops:
            if op_name not in ordered:
                ordered.append(op_name)
        return ordered

    def quote_format_operacoes(self, value: Any) -> str:
        return " + ".join(self.quote_parse_operacoes_lista(value))

    def _quote_collect_non_laser_map(self, *sources: Any, digits: int = 4) -> dict[str, float]:
        collected: dict[str, float] = {}
        for source in sources:
            if not isinstance(source, dict):
                continue
            for op_name, raw_value in dict(source or {}).items():
                normalized = str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if not normalized or normalized == "Corte Laser":
                    continue
                value = round(self._parse_float(raw_value, 0), digits)
                if value <= 0:
                    continue
                current = float(collected.get(normalized, 0) or 0)
                if value > current:
                    collected[normalized] = value
        return collected

    def _normalize_quote_operation_map(
        self,
        value: Any,
        operations: list[str],
        *,
        digits: int = 4,
    ) -> dict[str, float]:
        raw = dict(value or {}) if isinstance(value, dict) else {}
        cleaned: dict[str, float] = {}
        allowed = {str(op or "").strip() for op in list(operations or []) if str(op or "").strip()}
        for raw_name, raw_value in raw.items():
            op_name = self.desktop_main.normalize_operacao_nome(raw_name) or str(raw_name or "").strip()
            if not op_name or op_name not in allowed:
                continue
            if raw_value in (None, ""):
                continue
            cleaned[op_name] = round(self._parse_float(raw_value, 0), digits)
        return cleaned

    def _quote_line_operation_snapshot(
        self,
        payload: dict[str, Any],
        *,
        quote_number: str = "",
        quote_state: str = "",
    ) -> dict[str, Any]:
        row = dict(payload or {})
        raw_detail_map = {
            str(self.desktop_main.normalize_operacao_nome(item.get("nome", "")) or item.get("nome", "") or "").strip(): dict(item or {})
            for item in list(row.get("operacoes_detalhe", []) or [])
            if isinstance(item, dict) and str(item.get("nome", "") or "").strip()
        }
        operations = [
            str(op or "").strip()
            for op in list(self.quote_parse_operacoes_lista(row.get("operacao", "")) or [])
            if str(op or "").strip()
        ]
        ops_txt = " + ".join(operations)
        raw_flow = row.get("operacoes_fluxo")
        flow = self.desktop_main.build_operacoes_fluxo(ops_txt, raw_flow if isinstance(raw_flow, list) else None)
        tempo_total = round(self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 3)
        preco_unit = round(self._parse_float(row.get("preco_unit", 0), 0), 4)
        tempos_operacao = self._normalize_quote_operation_map(row.get("tempos_operacao", {}), operations, digits=3)
        custos_operacao = self._normalize_quote_operation_map(row.get("custos_operacao", {}), operations, digits=4)
        if raw_detail_map and (not tempos_operacao or not custos_operacao):
            estimate = dict(self.operation_cost_estimate(row) or {})
            estimated_rows = [dict(item or {}) for item in list(estimate.get("operations", []) or []) if isinstance(item, dict)]
            for item in estimated_rows:
                op_name = str(item.get("nome", "") or "").strip()
                if not op_name:
                    continue
                if op_name not in tempos_operacao and item.get("tempo_unit_min") not in (None, ""):
                    tempos_operacao[op_name] = round(self._parse_float(item.get("tempo_unit_min", 0), 0), 3)
                if op_name not in custos_operacao and item.get("custo_unit_eur") not in (None, ""):
                    custos_operacao[op_name] = round(self._parse_float(item.get("custo_unit_eur", 0), 0), 4)
        explicit_breakdown = bool(tempos_operacao or custos_operacao)
        if len(operations) == 1:
            single_name = operations[0]
            if single_name == "Corte Laser":
                tempos_operacao = {"Corte Laser": tempo_total} if tempo_total > 0 else {}
                custos_operacao = {"Corte Laser": preco_unit} if preco_unit > 0 else {}
                explicit_breakdown = bool(tempos_operacao or custos_operacao)
            else:
                if single_name not in tempos_operacao and tempo_total > 0:
                    tempos_operacao[single_name] = tempo_total
                if single_name not in custos_operacao and preco_unit > 0:
                    custos_operacao[single_name] = preco_unit
        resolved_count = sum(1 for op_name in operations if op_name in tempos_operacao and op_name in custos_operacao)
        if operations and resolved_count == len(operations):
            costing_mode = "detailed"
        elif explicit_breakdown:
            costing_mode = "partial_detail"
        elif len(operations) <= 1:
            costing_mode = "single_operation_total"
        else:
            costing_mode = "aggregate_pending"
        breakdown: list[dict[str, Any]] = []
        for index, op_name in enumerate(operations, start=1):
            existing = dict(raw_detail_map.get(op_name, {}) or {})
            breakdown.append(
                {
                    **existing,
                    "seq": index,
                    "nome": op_name,
                    "tempo_unit_min": tempos_operacao.get(op_name),
                    "custo_unit_eur": custos_operacao.get(op_name),
                    "tem_detalhe": op_name in tempos_operacao or op_name in custos_operacao,
                }
            )
        snapshot_tempo_total = tempo_total
        snapshot_preco_total = preco_unit
        if operations and resolved_count == len(operations):
            snapshot_tempo_total = round(sum(float(tempos_operacao.get(op_name, 0) or 0) for op_name in operations), 3)
            snapshot_preco_total = round(sum(float(custos_operacao.get(op_name, 0) or 0) for op_name in operations), 4)
        elif explicit_breakdown and snapshot_tempo_total <= 0 and snapshot_preco_total <= 0:
            snapshot_tempo_total = round(sum(float(value or 0) for value in tempos_operacao.values()), 3)
            snapshot_preco_total = round(sum(float(value or 0) for value in custos_operacao.values()), 4)
        return {
            "operacoes": operations,
            "operacoes_fluxo": [dict(item or {}) for item in list(flow or []) if isinstance(item, dict)],
            "operacoes_detalhe": breakdown,
            "tempos_operacao": tempos_operacao,
            "custos_operacao": custos_operacao,
            "quote_cost_snapshot": {
                "costing_mode": costing_mode,
                "tempo_total_peca_min": snapshot_tempo_total,
                "preco_unit_total_eur": snapshot_preco_total,
                "qtd": round(self._parse_float(row.get("qtd", 0), 0), 2),
                "quote_number": str(quote_number or "").strip(),
                "quote_state": str(quote_state or "").strip(),
            },
        }

    def _sync_quote_piece_registry(self, orc: dict[str, Any]) -> None:
        if not isinstance(orc, dict):
            return
        data = self.ensure_data()
        refs_db = data.setdefault("orc_refs", {})
        piece_history = data.setdefault("peca_hist", {})
        numero = str(orc.get("numero", "") or "").strip()
        estado = str(orc.get("estado", "") or "").strip()
        estado_norm = self.desktop_main.norm_text(estado)
        approved = "aprovado" in estado_norm
        client_code = self._ref_client_code(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", ""))
        updated_at = self.desktop_main.now_iso()
        for line in list(orc.get("linhas", []) or []):
            if not self.desktop_main.orc_line_is_piece(line):
                continue
            if str(line.get("stock_item_kind", "") or "").strip() == "raw_material" or str(line.get("stock_material_id", "") or "").strip():
                line["ref_interna"] = ""
                continue
            ref_ext = str(line.get("ref_externa", "") or "").strip()
            if not ref_ext:
                continue
            snapshot = self._quote_line_operation_snapshot(line, quote_number=numero, quote_state=estado)
            existing_ref = dict(refs_db.get(ref_ext, {}) or {})
            approved_at = str(existing_ref.get("approved_at", "") or "").strip()
            if approved and not approved_at:
                approved_at = updated_at
            record = {
                **existing_ref,
                "ref_interna": str(line.get("ref_interna", existing_ref.get("ref_interna", "")) or "").strip(),
                "ref_externa": ref_ext,
                "descricao": str(line.get("descricao", existing_ref.get("descricao", "")) or "").strip(),
                "material": str(line.get("material", existing_ref.get("material", "")) or "").strip(),
                "espessura": str(line.get("espessura", existing_ref.get("espessura", "")) or "").strip(),
                "preco_unit": round(self._parse_float(line.get("preco_unit", existing_ref.get("preco_unit", 0)), 0), 4),
                "operacao": str(line.get("operacao", existing_ref.get("operacao", "")) or "").strip(),
                "tempo_peca_min": round(self._parse_float(line.get("tempo_peca_min", existing_ref.get("tempo_peca_min", 0)), 0), 3),
                "desenho": str(line.get("desenho", existing_ref.get("desenho", existing_ref.get("desenho_path", ""))) or "").strip(),
                "cliente_codigo": client_code,
                "origem_doc": numero,
                "origem_tipo": "Orcamento aprovado" if approved else "Orcamento",
                "estado_origem": estado,
                "approved_at": approved_at,
                "updated_at": updated_at,
                "operacoes_lista": list(snapshot.get("operacoes", []) or []),
                "operacoes_fluxo": [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                "operacoes_detalhe": [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(snapshot.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(snapshot.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(snapshot.get("quote_cost_snapshot", {}) or {}),
            }
            refs_db[ref_ext] = record

            existing_piece = dict(piece_history.get(ref_ext, {}) or {})
            quote_history = [
                dict(item or {})
                for item in list(existing_piece.get("quote_history", []) or [])
                if isinstance(item, dict) and str(item.get("numero", "") or "").strip() != numero
            ]
            quote_history.append(
                {
                    "numero": numero,
                    "estado": estado,
                    "updated_at": updated_at,
                    "approved_at": approved_at,
                    "preco_unit": record.get("preco_unit", 0),
                    "tempo_peca_min": record.get("tempo_peca_min", 0),
                }
            )
            piece_history[ref_ext] = {
                **existing_piece,
                "ref_interna": record.get("ref_interna", ""),
                "ref_externa": ref_ext,
                "descricao": record.get("descricao", ""),
                "material": record.get("material", ""),
                "espessura": record.get("espessura", ""),
                "Operacoes": record.get("operacao", ""),
                "Observacoes": record.get("descricao", ""),
                "desenho": record.get("desenho", ""),
                "operacoes_fluxo": [dict(item or {}) for item in list(record.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(record.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(record.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(record.get("quote_cost_snapshot", {}) or {}),
                "cliente_codigo": client_code,
                "origem_doc": numero,
                "estado_origem": estado,
                "approved_at": approved_at,
                "updated_at": updated_at,
                "quote_history": quote_history,
            }
            self._upsert_orc_ref_history_entry(ref_ext, record)

    def _ref_client_code(self, value: Any) -> str:
        raw = str(value or "").strip().upper()
        if raw.startswith("CL") and len(raw) >= 6 and raw[2:6].isdigit():
            return raw[:6]
        return ""

    def _active_client_ref_usage(self, cliente_codigo: str, exclude_orc_numero: str = "") -> tuple[set[str], set[tuple[str, str]]]:
        code = self._ref_client_code(cliente_codigo)
        exclude_num = str(exclude_orc_numero or "").strip()
        ref_internas: set[str] = set()
        ref_pairs: set[tuple[str, str]] = set()
        if not code:
            return ref_internas, ref_pairs

        for orc in list(self.ensure_data().get("orcamentos", []) or []):
            numero = str(orc.get("numero", "") or "").strip()
            if exclude_num and numero == exclude_num:
                continue
            orc_client = self._ref_client_code(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", ""))
            for line in list(orc.get("linhas", []) or []):
                if str(line.get("stock_item_kind", "") or "").strip() == "raw_material" or str(line.get("stock_material_id", "") or "").strip():
                    continue
                ref_int = str(line.get("ref_interna", "") or "").strip().upper()
                ref_ext = str(line.get("ref_externa", "") or "").strip()
                ref_client = self._ref_client_code(ref_int)
                ext_client = self._ref_client_code(ref_ext)
                if code not in {orc_client, ref_client, ext_client}:
                    continue
                if ref_int:
                    ref_internas.add(ref_int)
                if ref_ext or ref_int:
                    ref_pairs.add((ref_ext, ref_int))

        for enc in list(self.ensure_data().get("encomendas", []) or []):
            enc_client = self._ref_client_code(enc.get("cliente", ""))
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                ref_int = str(piece.get("ref_interna", "") or "").strip().upper()
                ref_ext = str(piece.get("ref_externa", "") or "").strip()
                ref_client = self._ref_client_code(ref_int)
                ext_client = self._ref_client_code(ref_ext)
                if code not in {enc_client, ref_client, ext_client}:
                    continue
                if ref_int:
                    ref_internas.add(ref_int)
                if ref_ext or ref_int:
                    ref_pairs.add((ref_ext, ref_int))

        return ref_internas, ref_pairs

    def _known_client_ref_pairs(self, cliente_codigo: str) -> set[tuple[str, str]]:
        code = self._ref_client_code(cliente_codigo)
        pairs: set[tuple[str, str]] = set()
        if not code:
            return pairs
        refs_db = self.ensure_data().get("orc_refs", {})
        for ref_ext, payload in list((refs_db or {}).items()):
            ref_externa = str(ref_ext or "").strip()
            ref_interna = str((payload or {}).get("ref_interna", "") or "").strip().upper()
            if not ref_interna:
                continue
            if code not in {self._ref_client_code(ref_externa), self._ref_client_code(ref_interna)}:
                continue
            pairs.add((ref_externa, ref_interna))
        _taken, active_pairs = self._active_client_ref_usage(code)
        for ref_externa, ref_interna in list(active_pairs or set()):
            ref_ext_txt = str(ref_externa or "").strip()
            ref_int_txt = str(ref_interna or "").strip().upper()
            if not ref_int_txt:
                continue
            if code not in {self._ref_client_code(ref_ext_txt), self._ref_client_code(ref_int_txt)}:
                continue
            pairs.add((ref_ext_txt, ref_int_txt))
        return pairs

    def _known_client_ref_for_external(self, cliente_codigo: str, ref_externa: str) -> str:
        code = self._ref_client_code(cliente_codigo)
        ref_ext_txt = str(ref_externa or "").strip()
        if not code or not ref_ext_txt:
            return ""
        refs_db = self.ensure_data().get("orc_refs", {})
        payload = (refs_db or {}).get(ref_ext_txt)
        ref_interna = str((payload or {}).get("ref_interna", "") or "").strip().upper()
        if ref_interna and self._ref_client_code(ref_interna) == code:
            return ref_interna

        candidates: list[str] = []
        for orc in list(self.ensure_data().get("orcamentos", []) or []):
            orc_client = self._ref_client_code(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", ""))
            for line in list(orc.get("linhas", []) or []):
                if str(line.get("ref_externa", "") or "").strip() != ref_ext_txt:
                    continue
                ref_int = str(line.get("ref_interna", "") or "").strip().upper()
                if ref_int and code in {orc_client, self._ref_client_code(ref_int), self._ref_client_code(ref_ext_txt)}:
                    candidates.append(ref_int)
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            enc_client = self._ref_client_code(enc.get("cliente", ""))
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                if str(piece.get("ref_externa", "") or "").strip() != ref_ext_txt:
                    continue
                ref_int = str(piece.get("ref_interna", "") or "").strip().upper()
                if ref_int and code in {enc_client, self._ref_client_code(ref_int), self._ref_client_code(ref_ext_txt)}:
                    candidates.append(ref_int)
        if not candidates:
            return ""
        candidates = sorted(set(candidates), key=lambda ref: (self.desktop_main._extract_ref_interna_seq(ref, code) or 999999, ref))
        return candidates[0]

    def _upsert_orc_ref_history_entry(self, ref_ext: str, payload: dict[str, Any]) -> None:
        try:
            self.desktop_main.mysql_upsert_orc_referencia(
                ref_externa=ref_ext,
                ref_interna=str(payload.get("ref_interna", "") or "").strip(),
                descricao=str(payload.get("descricao", "") or "").strip(),
                material=str(payload.get("material", "") or "").strip(),
                espessura=str(payload.get("espessura", "") or "").strip(),
                preco_unit=self._parse_float(payload.get("preco_unit", 0), 0),
                operacao=str(payload.get("operacao", "") or "").strip(),
                desenho_path=str(payload.get("desenho", "") or payload.get("desenho_path", "") or "").strip(),
                tempo_peca_min=self._parse_float(payload.get("tempo_peca_min", 0), 0),
                operacoes_lista=list(payload.get("operacoes_lista", []) or []),
                operacoes_fluxo=[dict(item or {}) for item in list(payload.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                operacoes_detalhe=[dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                tempos_operacao=dict(payload.get("tempos_operacao", {}) or {}),
                custos_operacao=dict(payload.get("custos_operacao", {}) or {}),
                quote_cost_snapshot=dict(payload.get("quote_cost_snapshot", {}) or {}),
                origem_doc=str(payload.get("origem_doc", "") or "").strip(),
                origem_tipo=str(payload.get("origem_tipo", "") or "").strip(),
                estado_origem=str(payload.get("estado_origem", "") or "").strip(),
                approved_at=str(payload.get("approved_at", "") or "").strip(),
            )
        except Exception:
            pass

    def _delete_orc_ref_history_entry(self, ref_ext: str) -> None:
        delete_fn = getattr(self.desktop_main, "mysql_delete_orc_referencia", None)
        if not callable(delete_fn):
            return
        try:
            delete_fn(ref_ext)
        except Exception:
            pass

    def _repair_quote_refs_from_orders(self, cliente_codigo: str) -> bool:
        code = self._ref_client_code(cliente_codigo)
        if not code:
            return False
        data = self.ensure_data()
        changed = False
        for orc in list(data.get("orcamentos", []) or []):
            orc_client = self._ref_client_code(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", ""))
            if orc_client != code:
                continue
            enc_numero = str(orc.get("numero_encomenda", "") or "").strip()
            if not enc_numero:
                continue
            enc = self.get_encomenda_by_numero(enc_numero)
            if enc is None:
                continue
            pieces = list(self.desktop_main.encomenda_pecas(enc) or [])
            used_piece_keys: set[str] = set()
            for index, line in enumerate(list(orc.get("linhas", []) or [])):
                ref_ext = str(line.get("ref_externa", "") or "").strip()
                ref_int = str(line.get("ref_interna", "") or "").strip().upper()
                match = None
                if ref_ext:
                    candidates = [piece for piece in pieces if str(piece.get("ref_externa", "") or "").strip() == ref_ext]
                    if candidates:
                        exact = next((piece for piece in candidates if str(piece.get("ref_interna", "") or "").strip().upper() == ref_int), None)
                        if exact is not None:
                            match = exact
                        else:
                            for piece in candidates:
                                piece_key = str(piece.get("id", "") or "").strip() or str(id(piece))
                                if piece_key not in used_piece_keys:
                                    match = piece
                                    break
                if match is None and index < len(pieces):
                    candidate = pieces[index]
                    candidate_key = str(candidate.get("id", "") or "").strip() or str(id(candidate))
                    if candidate_key not in used_piece_keys:
                        match = candidate
                if match is None:
                    continue
                match_key = str(match.get("id", "") or "").strip() or str(id(match))
                used_piece_keys.add(match_key)
                piece_ref = str(match.get("ref_interna", "") or "").strip().upper()
                if piece_ref and piece_ref != ref_int:
                    line["ref_interna"] = piece_ref
                    changed = True
        if changed:
            self._save(force=True)
        return changed

    def _repair_orc_ref_history(self, cliente_codigo: str) -> bool:
        code = self._ref_client_code(cliente_codigo)
        if not code:
            return False
        data = self.ensure_data()
        refs_db = data.setdefault("orc_refs", {})
        active_ref_internas, active_ref_pairs = self._active_client_ref_usage(code)
        changed = False

        for ref_ext, payload in list((refs_db or {}).items()):
            key_client = self._ref_client_code(ref_ext)
            ref_interna = str((payload or {}).get("ref_interna", "") or "").strip().upper()
            ref_client = self._ref_client_code(ref_interna)
            if key_client != code:
                continue
            if ref_client and ref_client != code and (str(ref_ext or "").strip(), ref_interna) not in active_ref_pairs:
                refs_db.pop(ref_ext, None)
                self._delete_orc_ref_history_entry(str(ref_ext or "").strip())
                changed = True

        client_history: list[tuple[str, dict[str, Any]]] = []
        for ref_ext, payload in sorted(
            (refs_db or {}).items(),
            key=lambda item: (
                self.desktop_main._extract_ref_interna_seq((item[1] or {}).get("ref_interna", ""), code) or 999999,
                str(item[0] or ""),
            ),
        ):
            row = dict(payload or {})
            ref_interna = str(row.get("ref_interna", "") or "").strip().upper()
            ref_client = self._ref_client_code(ref_interna)
            ext_client = self._ref_client_code(ref_ext)
            if ref_client == code or (not ref_interna and ext_client == code):
                client_history.append((str(ref_ext or "").strip(), row))

        seq_map = data.setdefault("seq", {}).setdefault("ref_interna", {})
        if not active_ref_internas and client_history:
            for index, (ref_ext, row) in enumerate(client_history, start=1):
                new_ref = f"{code}-{index:04d}REV00"
                old_ref = str(row.get("ref_interna", "") or "").strip().upper()
                if old_ref != new_ref:
                    row["ref_interna"] = new_ref
                    refs_db[ref_ext] = row
                    self._upsert_orc_ref_history_entry(ref_ext, row)
                    changed = True
            seq_map[code] = len(client_history)
        else:
            max_seq = 0
            for _ref_ext, row in client_history:
                max_seq = max(max_seq, self.desktop_main._extract_ref_interna_seq(row.get("ref_interna", ""), code))
            if max_seq:
                seq_map[code] = max_seq

        if changed:
            self._save(force=True)
        return changed

    def normalize_client_reference_sequence(self, cliente_codigo: str) -> dict[str, Any]:
        code = self._ref_client_code(cliente_codigo)
        if not code:
            raise ValueError("Cliente invalido para normalizar referencias.")
        self._repair_quote_refs_from_orders(code)
        data = self.ensure_data()
        live_refs, _pairs = self._active_client_ref_usage(code)
        ordered_refs = sorted(
            {str(ref or "").strip().upper() for ref in list(live_refs or []) if str(ref or "").strip()},
            key=lambda ref: (self.desktop_main._extract_ref_interna_seq(ref, code) or 999999, ref),
        )
        mapping = {old: f"{code}-{index:04d}REV00" for index, old in enumerate(ordered_refs, start=1)}
        if mapping and any(old != new for old, new in mapping.items()):
            def replace_refs(node: Any) -> Any:
                if isinstance(node, dict):
                    for key, value in list(node.items()):
                        if isinstance(value, str):
                            updated = mapping.get(value.strip().upper())
                            if updated:
                                node[key] = updated
                        else:
                            replace_refs(value)
                elif isinstance(node, list):
                    for index, value in enumerate(list(node)):
                        if isinstance(value, str):
                            updated = mapping.get(value.strip().upper())
                            if updated:
                                node[index] = updated
                        else:
                            replace_refs(value)
                return node

            replace_refs(data)
            for bucket in ("op_eventos", "op_paragens"):
                for row in list(data.get(bucket, []) or []):
                    raw_ref = str((row or {}).get("ref_interna", "") or "").strip().upper()
                    updated = mapping.get(raw_ref)
                    if updated:
                        row["ref_interna"] = updated
            conn = None
            try:
                connect = getattr(self.desktop_main, "_mysql_connect", None)
                if callable(connect):
                    conn = connect()
                    existing_tables_fn = getattr(self.desktop_main, "_mysql_existing_tables", None)
                    tables = set()
                    if callable(existing_tables_fn):
                        try:
                            with conn.cursor() as cur:
                                tables = set(existing_tables_fn(cur, force=True) or [])
                        except Exception:
                            tables = set()
                    with conn.cursor() as cur:
                        for old_ref, new_ref in mapping.items():
                            if "op_eventos" in tables:
                                cur.execute("UPDATE op_eventos SET ref_interna=%s WHERE ref_interna=%s", (new_ref, old_ref))
                            if "op_paragens" in tables:
                                cur.execute("UPDATE op_paragens SET ref_interna=%s WHERE ref_interna=%s", (new_ref, old_ref))
                    conn.commit()
            except Exception:
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
            finally:
                try:
                    if conn:
                        conn.close()
                except Exception:
                    pass
        data.setdefault("seq", {}).setdefault("ref_interna", {})[code] = len(ordered_refs)
        self._repair_orc_ref_history(code)
        self._save(force=True)
        return {
            "cliente": code,
            "total_refs": len(ordered_refs),
            "mapping": mapping,
        }

    def _suggest_ref_interna_for_client(
        self,
        cliente_codigo: str,
        existing_refs: list[str] | tuple[str, ...] | set[str] | None = None,
        exclude_orc_numero: str = "",
    ) -> str:
        code = self._ref_client_code(cliente_codigo)
        if not code:
            return ""
        self._repair_orc_ref_history(code)
        taken_refs, _pairs = self._active_client_ref_usage(code, exclude_orc_numero=exclude_orc_numero)
        reserved = {
            str(ref or "").strip().upper()
            for ref in list(existing_refs or [])
            if str(ref or "").strip()
        }
        return str(self.desktop_main.next_ref_interna_unique(self.ensure_data(), code, list(taken_refs | reserved)))

    def order_reference_rows(self, filter_text: str = "", cliente: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        cliente_codigo = str(cliente or "").strip().upper()
        if cliente_codigo:
            self._repair_quote_refs_from_orders(cliente_codigo)
            self._repair_orc_ref_history(cliente_codigo)
        data = self.ensure_data()
        refs_db = data.get("orc_refs", {})
        rows: list[dict[str, Any]] = []
        row_index_by_pair: dict[tuple[str, str], int] = {}

        def row_client(ref_interna: str, ref_externa: str, explicit_client: str = "") -> str:
            return (
                self._ref_client_code(explicit_client)
                or self._ref_client_code(ref_interna)
                or self._ref_client_code(ref_externa)
            )

        def include_row(ref_interna: str, ref_externa: str, explicit_client: str = "") -> bool:
            if not cliente_codigo:
                return True
            return row_client(ref_interna, ref_externa, explicit_client) == cliente_codigo

        def append_row(payload: dict[str, Any]) -> None:
            raw_operacoes = str(payload.get("operacoes", payload.get("operacao", "")) or "").strip()
            row = {
                "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
                "ref_interna": str(payload.get("ref_interna", "") or "").strip(),
                "descricao": str(payload.get("descricao", "") or "").strip(),
                "material": str(payload.get("material", "") or "").strip(),
                "espessura": str(payload.get("espessura", "") or "").strip(),
                "preco": round(self._parse_float(payload.get("preco", payload.get("preco_unit", 0)), 0), 4),
                "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
                "operacoes": raw_operacoes,
                "operacoes_lista": self.quote_parse_operacoes_lista(payload.get("operacoes_lista", []) or raw_operacoes),
                "operacoes_fluxo": [dict(item or {}) for item in list(payload.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                "operacoes_detalhe": [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(payload.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(payload.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(payload.get("quote_cost_snapshot", {}) or {}),
                "desenho": str(payload.get("desenho", "") or payload.get("desenho_path", "") or "").strip(),
                "laser_base_active": bool(payload.get("laser_base_active", False)),
                "laser_base_tempo_unit": round(self._parse_float(payload.get("laser_base_tempo_unit", 0), 0), 4),
                "laser_base_preco_unit": round(self._parse_float(payload.get("laser_base_preco_unit", 0), 0), 4),
                "origem_doc": str(payload.get("origem_doc", "") or "").strip(),
                "origem_tipo": str(payload.get("origem_tipo", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                return
            pair_key = (row["ref_interna"].upper(), row["ref_externa"])
            existing_index = row_index_by_pair.get(pair_key)
            if existing_index is not None:
                existing = rows[existing_index]

                def score(candidate: dict[str, Any]) -> int:
                    base = {"Orcamento": 30, "Encomenda": 20, "Historico": 10}.get(str(candidate.get("origem_tipo", "") or ""), 0)
                    richness = 0
                    for field in ("descricao", "material", "espessura", "operacoes", "desenho"):
                        if str(candidate.get(field, "") or "").strip():
                            richness += 2
                    if float(candidate.get("preco", 0) or 0) > 0:
                        richness += 1
                    if float(candidate.get("tempo_peca_min", 0) or 0) > 0:
                        richness += 1
                    return base + richness

                for field in ("descricao", "material", "espessura", "operacoes", "desenho"):
                    if not str(existing.get(field, "") or "").strip() and str(row.get(field, "") or "").strip():
                        existing[field] = row[field]
                if float(existing.get("preco", 0) or 0) <= 0 and float(row.get("preco", 0) or 0) > 0:
                    existing["preco"] = row["preco"]
                if float(existing.get("tempo_peca_min", 0) or 0) <= 0 and float(row.get("tempo_peca_min", 0) or 0) > 0:
                    existing["tempo_peca_min"] = row["tempo_peca_min"]
                if score(row) > score(existing):
                    existing["origem_doc"] = row["origem_doc"]
                    existing["origem_tipo"] = row["origem_tipo"]
                return
            row_index_by_pair[pair_key] = len(rows)
            rows.append(row)

        for orc in list(data.get("orcamentos", []) or []):
            explicit_client = self._normalize_orc_client(orc.get("cliente", {})).get("codigo", "")
            for line in list(orc.get("linhas", []) or []):
                ref_interna = str(line.get("ref_interna", "") or "").strip()
                ref_externa = str(line.get("ref_externa", "") or "").strip()
                if not include_row(ref_interna, ref_externa, explicit_client):
                    continue
                append_row(
                    {
                        "ref_interna": ref_interna,
                        "ref_externa": ref_externa,
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "material": str(line.get("material", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "preco_unit": line.get("preco_unit", 0),
                        "tempo_peca_min": line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)),
                        "operacoes": self.quote_format_operacoes(line.get("operacao", "")),
                        "desenho": str(line.get("desenho", "") or "").strip(),
                        "origem_doc": str(orc.get("numero", "") or "").strip(),
                        "origem_tipo": "Orcamento",
                    }
                )

        for enc in list(data.get("encomendas", []) or []):
            explicit_client = str(enc.get("cliente", "") or "").strip()
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                ref_interna = str(piece.get("ref_interna", "") or "").strip()
                ref_externa = str(piece.get("ref_externa", "") or "").strip()
                if not include_row(ref_interna, ref_externa, explicit_client):
                    continue
                append_row(
                    {
                        "ref_interna": ref_interna,
                        "ref_externa": ref_externa,
                        "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                        "material": str(piece.get("material", "") or "").strip(),
                        "espessura": str(piece.get("espessura", "") or "").strip(),
                        "preco_unit": piece.get("preco_unit", 0),
                        "tempo_peca_min": piece.get("tempo_peca_min", piece.get("tempo_pecas_min", 0)),
                        "operacoes": self.quote_format_operacoes(
                            piece.get("operacoes")
                            or " + ".join(
                                self.desktop_main.normalize_operacao_nome(op.get("nome", ""))
                                for op in list(self.desktop_main.ensure_peca_operacoes(piece) or [])
                                if str(op.get("nome", "") or "").strip()
                            )
                        ),
                        "desenho": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                        "origem_doc": str(enc.get("numero", "") or "").strip(),
                        "origem_tipo": "Encomenda",
                    }
                )

        for ref_ext, payload in sorted((refs_db or {}).items(), key=lambda item: str(item[0] or "")):
            ref_interna = str(payload.get("ref_interna", "") or "").strip()
            ref_externa = str(ref_ext or "").strip()
            if not include_row(ref_interna, ref_externa):
                continue
            append_row(
                {
                    "ref_externa": ref_externa,
                    "ref_interna": ref_interna,
                    "descricao": str(payload.get("descricao", "") or "").strip(),
                    "material": str(payload.get("material", "") or "").strip(),
                    "espessura": str(payload.get("espessura", "") or "").strip(),
                    "preco_unit": payload.get("preco_unit", 0),
                    "tempo_peca_min": payload.get("tempo_peca_min", 0),
                    "operacoes": self.quote_format_operacoes(payload.get("operacao", "")),
                    "desenho": str(payload.get("desenho", "") or payload.get("desenho_path", "") or "").strip(),
                    "origem_doc": "Historico",
                    "origem_tipo": "Historico",
                }
            )
        rows.sort(
            key=lambda row: (
                self.desktop_main._extract_ref_interna_seq(str(row.get("ref_interna", "") or ""), self._ref_client_code(str(row.get("ref_interna", "") or "")) or cliente_codigo)
                or 999999,
                str(row.get("ref_interna", "") or ""),
                str(row.get("ref_externa", "") or ""),
                str(row.get("origem_doc", "") or ""),
            )
        )
        return rows

    def _order_sequence_from_numero(self, numero: str) -> str:
        numero_txt = str(numero or "").strip()
        match = re.search(r"(\d+)$", numero_txt)
        if not match:
            return ""
        return match.group(1).zfill(4)[-4:]

    def _order_expected_of_code(self, enc: dict[str, Any]) -> str:
        seq = self._order_sequence_from_numero(str((enc or {}).get("numero", "") or ""))
        if not seq:
            return ""
        year_txt = str((enc or {}).get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:4]
        if not (len(year_txt) == 4 and year_txt.isdigit()):
            year_txt = str(datetime.now().year)
        return f"OF-{year_txt}-{seq}"

    def _order_of_code(self, enc: dict[str, Any], *, create: bool = True) -> str:
        expected_code = self._order_expected_of_code(enc)
        code = str(enc.get("of_codigo", "") or "").strip()
        if not code:
            ordem = enc.get("ordem_fabrico", {})
            if isinstance(ordem, dict):
                code = str(ordem.get("id", "") or ordem.get("codigo", "") or "").strip()
        if not code:
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                code = str(piece.get("of", "") or "").strip()
                if code:
                    break
        if expected_code and (not code or re.fullmatch(r"OF-\d{4}-\d{4,}", code, flags=re.IGNORECASE)):
            code = expected_code
        if not code and create:
            code = expected_code or str(self.desktop_main.next_of_numero(self.ensure_data()) or "").strip()
        if code:
            enc["of_codigo"] = code
            enc["ordem_fabrico"] = {
                "id": code,
                "encomenda_id": str(enc.get("numero", "") or "").strip(),
                "estado": str(enc.get("estado", "") or "Preparacao").strip() or "Preparacao",
                "data": str(enc.get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:10],
            }
        return code

    def _next_order_opp_codigo(self, enc: dict[str, Any]) -> str:
        of_code = self._order_of_code(enc, create=True)
        prefix = ""
        parts = of_code.split("-")
        if len(parts) >= 3 and parts[0].upper() == "OF":
            prefix = f"OPP-{parts[1]}-{parts[2]}"
        max_seq = 0
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            opp = str(piece.get("opp", "") or "").strip()
            if prefix and opp.startswith(prefix + "-"):
                suffix = opp.rsplit("-", 1)[-1]
                if suffix.isdigit():
                    max_seq = max(max_seq, int(suffix))
        if prefix:
            return f"{prefix}-{max_seq + 1:02d}"
        return str(self.desktop_main.next_opp_numero(self.ensure_data()) or "").strip()

    def _ensure_order_fabrication_order(self, enc: dict[str, Any], *, sync_existing: bool = False) -> bool:
        changed = False
        previous_of = str(enc.get("of_codigo", "") or "").strip()
        of_code = self._order_of_code(enc, create=True)
        if of_code and previous_of != of_code:
            changed = True
        if of_code and str(enc.get("of_codigo", "") or "").strip() != of_code:
            enc["of_codigo"] = of_code
            changed = True
        if of_code:
            ordem = {
                "id": of_code,
                "encomenda_id": str(enc.get("numero", "") or "").strip(),
                "estado": str(enc.get("estado", "") or "Preparacao").strip() or "Preparacao",
                "data": str(enc.get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:10],
            }
            if dict(enc.get("ordem_fabrico", {}) or {}) != ordem:
                enc["ordem_fabrico"] = ordem
                changed = True
        previous_opp_prefix = ""
        if previous_of and previous_of != of_code and previous_of.startswith("OF-"):
            previous_opp_prefix = "OPP-" + previous_of.split("-", 1)[1]
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            piece_of = str(piece.get("of", "") or "").strip()
            should_sync_piece_of = sync_existing or not piece_of or (previous_of and piece_of == previous_of)
            if of_code and should_sync_piece_of and piece_of != of_code:
                piece["of"] = of_code
                changed = True
            piece_opp = str(piece.get("opp", "") or "").strip()
            should_regen_opp = not piece_opp or bool(previous_opp_prefix and piece_opp.startswith(previous_opp_prefix + "-"))
            if should_regen_opp:
                piece["opp"] = self._next_order_opp_codigo(enc)
                changed = True
        return changed

    def order_suggest_ref_interna(
        self,
        numero: str = "",
        cliente: str = "",
        existing_refs: list[str] | tuple[str, ...] | set[str] | None = None,
    ) -> str:
        enc = self.get_encomenda_by_numero(numero) if str(numero or "").strip() else None
        cliente_codigo = str(cliente or (enc or {}).get("cliente", "") or "").strip()
        if not cliente_codigo:
            return ""
        existing = list(existing_refs or [])
        if enc is not None:
            existing.extend(str(piece.get("ref_interna", "") or "").strip() for piece in list(self.desktop_main.encomenda_pecas(enc)))
        return self._suggest_ref_interna_for_client(cliente_codigo, existing)

    def orc_suggest_ref_interna(
        self,
        cliente: str = "",
        existing_refs: list[str] | tuple[str, ...] | set[str] | None = None,
        numero: str = "",
    ) -> str:
        return self._suggest_ref_interna_for_client(cliente, existing_refs, exclude_orc_numero=numero)

    def order_create_or_update(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip()
        enc = self.get_encomenda_by_numero(numero) if numero else None
        is_new = enc is None
        cliente = str(payload.get("cliente", "") or "").strip()
        if not cliente:
            raise ValueError("Cliente obrigatorio.")
        try:
            tempo_estimado = self._parse_float(payload.get("tempo_estimado", 0), 0)
        except Exception as exc:
            raise ValueError("Tempo estimado inv?lido.") from exc
        if enc is None:
            numero = self.desktop_main.next_encomenda_numero(data)
            enc = {
                "id": f"ENC{len(list(data.get('encomendas', []) or [])) + 1:05d}",
                "numero": numero,
                "cliente": cliente,
                "posto_trabalho": "",
                "nota_cliente": "",
                "nota_transporte": "",
                "preco_transporte": 0.0,
                "custo_transporte": 0.0,
                "paletes": 0.0,
                "peso_bruto_kg": 0.0,
                "volume_m3": 0.0,
                "transportadora_id": "",
                "transportadora_nome": "",
                "referencia_transporte": "",
                "zona_transporte": "",
                "local_descarga": "",
                "transporte_numero": "",
                "estado_transporte": "",
                "data_criacao": self.desktop_main.now_iso(),
                "data_entrega": "",
                "tempo_estimado": 0.0,
                "tempo": 0.0,
                "cativar": False,
                "Observacoes": "",
                "Observações": "",
                "estado": "Preparacao",
                "materiais": [],
                "reservas": [],
                "montagem_itens": [],
                "espessuras": [],
                "numero_orcamento": "",
                "tipo_encomenda": "Cliente",
                "of_codigo": "",
                "ordem_fabrico": {},
            }
            of_code = self._order_of_code(enc, create=True)
            data.setdefault("encomendas", []).append(enc)

        enc["cliente"] = cliente
        tipo_txt = str(payload.get("tipo_encomenda", enc.get("tipo_encomenda", "Cliente")) or "Cliente").strip()
        enc["tipo_encomenda"] = "Interna (produção)" if "intern" in self.desktop_main.norm_text(tipo_txt) else "Cliente"
        enc["posto_trabalho"] = self._normalize_workcenter_value(payload.get("posto_trabalho", "") or enc.get("posto_trabalho", ""))
        enc["nota_cliente"] = str(payload.get("nota_cliente", "") or "").strip()
        enc["nota_transporte"] = str(payload.get("nota_transporte", "") or enc.get("nota_transporte", "") or "").strip()
        enc["preco_transporte"] = round(self._parse_float(payload.get("preco_transporte", enc.get("preco_transporte", 0)), 0), 2)
        enc["custo_transporte"] = round(self._parse_float(payload.get("custo_transporte", enc.get("custo_transporte", 0)), 0), 2)
        enc["paletes"] = round(self._parse_float(payload.get("paletes", enc.get("paletes", 0)), 0), 2)
        enc["peso_bruto_kg"] = round(self._parse_float(payload.get("peso_bruto_kg", enc.get("peso_bruto_kg", 0)), 0), 2)
        enc["volume_m3"] = round(self._parse_float(payload.get("volume_m3", enc.get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self._normalize_supplier_reference(
            payload.get("transportadora_id", enc.get("transportadora_id", "")),
            payload.get("transportadora_nome", enc.get("transportadora_nome", "")),
        )
        enc["transportadora_id"] = transportadora_id
        enc["transportadora_nome"] = transportadora_nome
        enc["referencia_transporte"] = str(payload.get("referencia_transporte", enc.get("referencia_transporte", "")) or "").strip()
        enc["zona_transporte"] = str(payload.get("zona_transporte", enc.get("zona_transporte", "")) or "").strip()
        enc["local_descarga"] = str(payload.get("local_descarga", "") or enc.get("local_descarga", "") or "").strip()
        enc["data_entrega"] = str(payload.get("data_entrega", "") or "").strip()
        enc["tempo_estimado"] = tempo_estimado
        enc["tempo"] = tempo_estimado
        obs_txt = str(payload.get("observacoes", "") or "").strip()
        enc["Observacoes"] = obs_txt
        enc["Observações"] = obs_txt
        requested_cativar = bool(payload.get("cativar"))
        if not requested_cativar and list(enc.get("reservas", []) or []):
            self.desktop_main.aplicar_reserva_em_stock(data, list(enc.get("reservas", []) or []), -1)
            enc["reservas"] = []
        enc.setdefault("montagem_itens", [])
        enc["cativar"] = requested_cativar and bool(enc.get("reservas"))
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_remove(self, numero: str) -> None:
        numero = str(numero or "").strip()
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if list(enc.get("reservas", []) or []):
            self.desktop_main.aplicar_reserva_em_stock(self.ensure_data(), list(enc.get("reservas", []) or []), -1)
        self.ensure_data()["encomendas"] = [row for row in list(self.ensure_data().get("encomendas", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        delete_order_fn = getattr(self.operador_actions, "_mysql_ops_delete_order", None)
        if callable(delete_order_fn):
            try:
                delete_order_fn(numero)
            except Exception:
                pass
        try:
            cache = getattr(self, "_op_mysql_ops_status_cache", None)
            if isinstance(cache, dict):
                cache.pop(numero, None)
        except Exception:
            pass
        for trip in list(self.ensure_data().get("transportes", []) or []):
            if not isinstance(trip, dict):
                continue
            trip["paragens"] = [
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() != numero
            ]
            self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        self._save(force=True)

    def order_material_add(self, numero: str, material: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: material bloqueado.")
        material_txt = str(material or "").strip()
        if not material_txt:
            raise ValueError("Material obrigatorio.")
        if self._order_find_material(enc, material_txt) is not None:
            raise ValueError("Material ja existe.")
        enc.setdefault("materiais", []).append({"material": material_txt, "estado": "Preparacao", "espessuras": []})
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_material_remove(self, numero: str, material: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: material bloqueado.")
        material_txt = str(material or "").strip().lower()
        enc["materiais"] = [row for row in list(enc.get("materiais", []) or []) if str(row.get("material", "") or "").strip().lower() != material_txt]
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_add(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: espessuras bloqueadas.")
        mat = self._order_find_material(enc, material)
        if mat is None:
            raise ValueError("Material n?o encontrado.")
        esp_txt = str(espessura or "").strip()
        if not esp_txt:
            raise ValueError("Espessura obrigatoria.")
        if self._order_find_espessura(enc, material, esp_txt) is not None:
            raise ValueError("Espessura ja existe.")
        mat.setdefault("espessuras", []).append({"espessura": esp_txt, "tempo_min": "", "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []})
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_remove(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: espessuras bloqueadas.")
        mat = self._order_find_material(enc, material)
        if mat is None:
            raise ValueError("Material n?o encontrado.")
        esp_txt = str(espessura or "").strip()
        mat["espessuras"] = [row for row in list(mat.get("espessuras", []) or []) if str(row.get("espessura", "") or "").strip() != esp_txt]
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_set_time(self, numero: str, material: str, espessura: str, tempo_min: Any) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            raise ValueError("Espessura nao encontrada.")
        raw = str(tempo_min if tempo_min is not None else "").strip()
        if raw:
            try:
                int(raw)
            except Exception as exc:
                raise ValueError("Tempo inv?lido (minutos inteiros).") from exc
        esp["tempo_min"] = raw
        esp.setdefault("tempos_operacao", {})
        if raw:
            esp["tempos_operacao"]["Corte Laser"] = raw
        else:
            esp["tempos_operacao"].pop("Corte Laser", None)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_set_operation_times(
        self,
        numero: str,
        material: str,
        espessura: str,
        tempos_operacao: dict[str, Any] | None = None,
        maquinas_operacao: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            raise ValueError("Espessura não encontrada.")
        cleaned: dict[str, str] = {}
        cleaned_resources: dict[str, str] = {}
        for op_name, raw_value in dict(tempos_operacao or {}).items():
            op_txt = self._planning_normalize_operation(op_name)
            if op_txt not in self.planning_operation_options():
                continue
            value_txt = str(raw_value if raw_value is not None else "").strip()
            if value_txt:
                try:
                    int(float(value_txt))
                except Exception as exc:
                    raise ValueError(f"Tempo inválido em {op_txt} (minutos inteiros).") from exc
                cleaned[op_txt] = value_txt
        for op_name, raw_value in dict(maquinas_operacao or {}).items():
            op_txt = self._planning_normalize_operation(op_name)
            if op_txt not in cleaned:
                continue
            resource_txt = self._sanitize_operation_resource(op_txt, raw_value)
            available_resources = [str(value or "").strip() for value in list(self.workcenter_resource_options(op_txt) or []) if str(value or "").strip()]
            if not resource_txt and len(available_resources) == 1:
                resource_txt = str(available_resources[0] or "").strip()
            if not resource_txt:
                continue
            if available_resources and all(resource_txt.lower() != value.lower() for value in available_resources):
                raise ValueError(f"O recurso '{resource_txt}' não pertence à operação {op_txt}.")
            cleaned_resources[op_txt] = resource_txt
        missing_resource = [op_name for op_name in cleaned if not str(cleaned_resources.get(op_name, "") or "").strip()]
        if missing_resource:
            raise ValueError(f"Seleciona o recurso/máquina para: {', '.join(missing_resource)}.")
        esp["tempo_min"] = cleaned.get("Corte Laser", str(esp.get("tempo_min", "") or "").strip())
        if not cleaned.get("Corte Laser"):
            esp["tempo_min"] = ""
        esp["tempos_operacao"] = cleaned
        esp["maquinas_operacao"] = cleaned_resources
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def _next_order_piece_id(self, enc: dict[str, Any]) -> str:
        highest = 0
        for row in list(self.desktop_main.encomenda_pecas(enc)):
            try:
                highest = max(highest, int(str(row.get("id", "") or "").replace("PEC", "")))
            except Exception:
                continue
        return f"PEC{highest + 1:05d}"

    def order_piece_create_or_update(self, numero: str, payload: dict[str, Any], current_ref_interna: str = "") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: pecas bloqueadas.")

        current_ref = str(current_ref_interna or "").strip()
        ref_int = str(payload.get("ref_interna", "") or "").strip()
        ref_ext = str(payload.get("ref_externa", "") or "").strip()
        tipo_material = str(payload.get("tipo_material", "") or "").strip().upper()
        if tipo_material not in {"CHAPA", "PERFIL", "TUBO", "OUTROS"}:
            tipo_material = "CHAPA"
        material = str(payload.get("material", "") or "").strip()
        subtipo_material = str(payload.get("subtipo_material", "") or material).strip()
        espessura = str(payload.get("espessura", "") or "").strip()
        dimensao = str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip()
        perfil_tipo = str(payload.get("perfil_tipo", "") or "").strip()
        perfil_tamanho = str(payload.get("perfil_tamanho", "") or "").strip()
        comprimento_mm = self._parse_float(payload.get("comprimento_mm", 0), 0)
        tubo_forma = str(payload.get("tubo_forma", "") or "").strip()
        lado_a = self._parse_float(payload.get("lado_a", 0), 0)
        lado_b = self._parse_float(payload.get("lado_b", 0), 0)
        tubo_espessura = self._parse_float(payload.get("tubo_espessura", 0), 0)
        diametro = self._parse_float(payload.get("diametro", 0), 0)
        descricao = str(payload.get("descricao", "") or "").strip()
        desenho = str(payload.get("desenho", "") or "").strip()
        ficheiros = [
            str(item or "").strip()
            for item in list(payload.get("ficheiros", []) or [])
            if str(item or "").strip()
        ]
        operacoes = " + ".join(self.desktop_main.parse_operacoes_lista(payload.get("operacoes", "")))
        if not operacoes:
            operacoes = str(self.desktop_main.OFF_OPERACAO_OBRIGATORIA)
        quantidade = self._parse_float(payload.get("quantidade_pedida", 0), 0)
        preco_unit = self._parse_float(payload.get("preco_unit", 0), 0)
        tempos_operacao = dict(payload.get("tempos_operacao", {}) or {})
        custos_operacao = dict(payload.get("custos_operacao", {}) or {})
        operacoes_detalhe = [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        guardar_ref = bool(payload.get("guardar_ref", True))

        if not material:
            raise ValueError("Material obrigatorio.")
        if tipo_material == "CHAPA" and not espessura:
            raise ValueError("Espessura obrigatoria para chapa.")
        if tipo_material == "TUBO" and (not dimensao or not espessura):
            raise ValueError("Dimensao e espessura obrigatorias para tubo.")
        if tipo_material == "PERFIL" and not dimensao:
            raise ValueError("Dimensao obrigatoria para perfil.")
        if not espessura and tipo_material in {"PERFIL", "OUTROS"}:
            espessura = "-"
        if quantidade <= 0:
            raise ValueError("Quantidade invalida.")

        existing_refs = {
            str(piece.get("ref_interna", "") or "").strip()
            for piece in list(self.desktop_main.encomenda_pecas(enc))
            if str(piece.get("ref_interna", "") or "").strip()
        }
        if current_ref:
            existing_refs.discard(current_ref)
        if not ref_int:
            ref_int = self.desktop_main.next_ref_interna_unique(self.ensure_data(), enc.get("cliente", ""), list(existing_refs))
        if ref_int and ref_int in existing_refs:
            suggested = self.desktop_main.next_ref_interna_unique(self.ensure_data(), enc.get("cliente", ""), list(existing_refs))
            raise ValueError(f"Referencia interna ja existe nesta encomenda. Nova sugerida: {suggested}")

        mat = self._order_find_material(enc, material)
        if mat is None:
            mat = {"material": material, "estado": "Preparacao", "espessuras": []}
            enc.setdefault("materiais", []).append(mat)
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            esp = {"espessura": espessura, "tempo_min": "", "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []}
            mat.setdefault("espessuras", []).append(esp)

        _, old_esp, piece = self._order_find_piece(enc, current_ref, "")
        if piece is None:
            of_code = self._order_of_code(enc, create=True)
            piece = {
                "id": self._next_order_piece_id(enc),
                "of": of_code,
                "opp": self._next_order_opp_codigo(enc),
                "estado": "Preparacao",
                "produzido_ok": 0.0,
                "produzido_nok": 0.0,
                "produzido_qualidade": 0.0,
                "inicio_producao": "",
                "fim_producao": "",
                "hist": [],
                "qtd_expedida": 0.0,
                "expedicoes": [],
            }
            esp.setdefault("pecas", []).append(piece)
        elif old_esp is not esp:
            old_esp["pecas"] = [row for row in list(old_esp.get("pecas", []) or []) if row is not piece]
            esp.setdefault("pecas", []).append(piece)

        piece["ref_interna"] = ref_int
        piece["ref_externa"] = ref_ext
        piece["material"] = material
        piece["tipo_material"] = tipo_material
        piece["subtipo_material"] = subtipo_material
        piece["espessura"] = espessura
        piece["dimensao"] = dimensao
        piece["dimensoes"] = dimensao
        piece["perfil_tipo"] = perfil_tipo
        piece["perfil_tamanho"] = perfil_tamanho
        piece["comprimento_mm"] = comprimento_mm
        piece["tubo_forma"] = tubo_forma
        piece["lado_a"] = lado_a
        piece["lado_b"] = lado_b
        piece["tubo_espessura"] = tubo_espessura
        piece["diametro"] = diametro
        piece["descricao"] = descricao
        piece["quantidade_pedida"] = quantidade
        piece["Operacoes"] = operacoes
        piece["Observacoes"] = descricao
        piece["desenho"] = desenho
        piece["desenho_path"] = desenho
        piece["ficheiros"] = ficheiros
        piece["tempos_operacao"] = dict(tempos_operacao)
        piece["custos_operacao"] = dict(custos_operacao)
        piece["operacoes_detalhe"] = list(operacoes_detalhe)
        if "of" not in piece or not str(piece.get("of", "")).strip():
            piece["of"] = self._order_of_code(enc, create=True)
        if "opp" not in piece or not str(piece.get("opp", "")).strip():
            piece["opp"] = self._next_order_opp_codigo(enc)
        piece["operacoes_fluxo"] = self.desktop_main.build_operacoes_fluxo(operacoes, piece.get("operacoes_fluxo"))
        self.desktop_main.ensure_peca_operacoes(piece)
        self.desktop_main.atualizar_estado_peca(piece)

        self.desktop_main.push_unique(self.ensure_data().setdefault("materiais_hist", []), material)
        self.desktop_main.push_unique(self.ensure_data().setdefault("espessuras_hist", []), espessura)

        if ref_ext:
            self.ensure_data().setdefault("peca_hist", {})[ref_ext] = {
                "ref_interna": ref_int,
                "descricao": descricao,
                "material": material,
                "espessura": espessura,
                "tipo_material": tipo_material,
                "subtipo_material": subtipo_material,
                "dimensao": dimensao,
                "Operacoes": operacoes,
                "Observacoes": descricao,
                "desenho": desenho,
                "ficheiros": list(ficheiros),
            }
            if guardar_ref:
                self.ensure_data().setdefault("orc_refs", {})[ref_ext] = {
                    "ref_interna": ref_int,
                    "ref_externa": ref_ext,
                    "descricao": descricao,
                    "material": material,
                    "espessura": espessura,
                    "preco_unit": preco_unit,
                    "operacao": operacoes,
                    "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", 0), 0), 3),
                    "operacoes_detalhe": list(operacoes_detalhe),
                    "tempos_operacao": dict(tempos_operacao),
                    "custos_operacao": dict(custos_operacao),
                    "desenho": desenho,
                }
                try:
                    self.desktop_main.mysql_upsert_orc_referencia(
                        ref_externa=ref_ext,
                        ref_interna=ref_int,
                        descricao=descricao,
                        material=material,
                        espessura=espessura,
                        preco_unit=preco_unit,
                        operacao=operacoes,
                        desenho_path=desenho,
                        tempo_peca_min=self._parse_float(payload.get("tempo_peca_min", 0), 0),
                        operacoes_lista=list(self.desktop_main.parse_operacoes_lista(operacoes)),
                        operacoes_detalhe=list(operacoes_detalhe),
                        tempos_operacao=dict(tempos_operacao),
                        custos_operacao=dict(custos_operacao),
                    )
                except Exception:
                    pass

        self.desktop_main.update_refs(self.ensure_data(), ref_int, ref_ext)
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_model_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for source, getter in (("modelo", getattr(self, "assembly_model_rows", None)), ("conjunto", getattr(self, "conjunto_rows", None))):
            if not callable(getter):
                continue
            for row in list(getter(filter_text) or []):
                if not bool(row.get("ativo", True)):
                    continue
                item = dict(row)
                item["origem_tipo"] = source
                item["label"] = f"{item.get('codigo', '')} | {item.get('descricao', '')}".strip(" |")
                if query and not any(query in str(value).lower() for value in item.values()):
                    continue
                rows.append(item)
        rows.sort(key=lambda row: (str(row.get("origem_tipo", "")), str(row.get("codigo", ""))))
        return rows

    def order_import_model(self, numero: str, codigo: str, quantity: Any = 1, source: str = "modelo") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orçamento: estrutura bloqueada.")
        code = str(codigo or "").strip()
        if not code:
            raise ValueError("Seleciona um modelo/conjunto.")
        source_norm = str(source or "").strip().lower()
        expand_fn = self.conjunto_expand if source_norm == "conjunto" else self.assembly_model_expand
        rows = list(expand_fn(code, quantity) or [])
        if not rows:
            raise ValueError("O modelo não tem linhas para importar.")
        imported_pieces = 0
        imported_items = 0
        for line in rows:
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            if self.desktop_main.orc_line_is_piece(line):
                self.order_piece_create_or_update(
                    numero,
                    {
                        "ref_interna": "",
                        "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "tipo_material": str(line.get("tipo_material", "") or line.get("material_family", "") or "CHAPA").strip().upper(),
                        "material": str(line.get("material", "") or line.get("material_subtype", "") or "").strip(),
                        "subtipo_material": str(line.get("material_subtype", "") or line.get("material", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or line.get("profile_size", "") or line.get("tube_section", "") or "").strip(),
                        "operacoes": str(line.get("operacao", "") or "Embalamento").strip(),
                        "quantidade_pedida": self._parse_float(line.get("qtd", 0), 0),
                        "preco_unit": self._parse_float(line.get("preco_unit", 0), 0),
                        "tempo_peca_min": self._parse_float(line.get("tempo_peca_min", 0), 0),
                        "desenho": str(line.get("desenho", "") or "").strip(),
                        "guardar_ref": True,
                    },
                )
                imported_pieces += 1
                continue
            enc.setdefault("montagem_itens", []).append(
                {
                    "linha_ordem": len(list(enc.get("montagem_itens", []) or [])) + 1,
                    "tipo_item": line_type,
                    "stock_item_kind": str(line.get("stock_item_kind", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
                    "material": str(line.get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or "").strip(),
                    "stock_material_id": str(line.get("stock_material_id", "") or "").strip(),
                    "produto_codigo": str(line.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(line.get("produto_unid", "") or "").strip() or ("SV" if self.desktop_main.orc_line_is_service(line) else "UN"),
                    "qtd_planeada": round(self._parse_float(line.get("qtd", 0), 0), 2),
                    "qtd_consumida": 0.0,
                    "preco_unit": round(self._parse_float(line.get("preco_unit", 0), 0), 4),
                    "conjunto_codigo": str(line.get("conjunto_codigo", code) or "").strip(),
                    "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                    "estado": "Pendente",
                    "obs": str(line.get("operacao", "") or "").strip(),
                    "created_at": self.desktop_main.now_iso(),
                    "consumed_at": "",
                    "consumed_by": "",
                }
            )
            imported_items += 1
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        detail = self.order_detail(numero)
        detail["imported_pieces"] = imported_pieces
        detail["imported_items"] = imported_items
        return detail

    def order_fabrication_pdf(self, numero: str, output_path: str | Path | None = None) -> Path:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._ensure_order_fabrication_order(enc):
            self._save(force=True)
        detail = self.order_detail(numero)
        pieces = list(detail.get("pieces", []) or [])
        if not pieces:
            raise ValueError("A ordem de fabrico não tem peças.")
        of_code = str(detail.get("of_codigo", "") or self._order_of_code(enc)).strip()
        target = Path(output_path) if output_path else Path(tempfile.gettempdir()) / f"lugest_ordem_fabrico_{of_code}.pdf"
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except ModuleNotFoundError:
            lines = [
                "Ordem de Fabrico",
                f"OF: {of_code}",
                f"Encomenda: {detail.get('numero', '-')}",
                f"Cliente: {detail.get('cliente', '-') or '-'} - {detail.get('cliente_nome', '') or ''}".strip(" -"),
                f"Data: {str((detail.get('ordem_fabrico') or {}).get('data', '') or '')[:10]}",
                "",
            ]
            fallback_groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
            for piece in pieces:
                key = (str(piece.get("material", "") or "-").strip() or "-", str(piece.get("espessura", "") or "-").strip() or "-")
                fallback_groups.setdefault(key, []).append(piece)
            for (material_txt, esp_txt), group_rows in sorted(fallback_groups.items(), key=lambda item: (item[0][0].lower(), item[0][1].lower())):
                lines.append(f"Espessura: {material_txt} | {esp_txt} mm | {len(group_rows)} peca(s)")
                lines.append(f"Codigo grupo: GRP|{of_code}|{material_txt}|{esp_txt}")
                for piece in group_rows:
                    lines.append(
                        f"- {piece.get('ref_interna', '-') or '-'} | {piece.get('ref_externa', '-') or '-'} | "
                        f"Qtd {piece.get('qtd_plan', '0')} | Ops: {piece.get('operacoes', '-') or '-'}"
                    )
                lines.append("")
            self._write_basic_pdf(target, lines)
            return target
        page_w, page_h = A4
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        margin = 12 * mm
        inner_w = page_w - (margin * 2)
        header_h = 38 * mm
        table_header_h = 8 * mm
        row_h = 7.4 * mm
        group_h = 10.5 * mm
        footer_h = 9 * mm
        columns = [
            ("Ref.", 27 * mm),
            ("Ref. externa", 72 * mm),
            ("Material", 30 * mm),
            ("Esp.", 7 * mm),
            ("Qtd", 11 * mm),
            ("Operacoes", inner_w - (27 + 72 + 30 + 7 + 11) * mm),
        ]
        rows_per_page = max(1, int((page_h - (margin * 2) - header_h - table_header_h - footer_h) // row_h))
        grouped_pieces: list[tuple[str, str, list[dict[str, Any]]]] = []
        groups_map: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for piece in pieces:
            key = (str(piece.get("material", "") or "-").strip() or "-", str(piece.get("espessura", "") or "-").strip() or "-")
            groups_map.setdefault(key, []).append(piece)
        for key, rows in sorted(groups_map.items(), key=lambda item: (item[0][0].lower(), item[0][1].lower())):
            grouped_pieces.append((key[0], key[1], rows))
        visual_units = len(pieces) + (len(grouped_pieces) * max(1, math.ceil(group_h / row_h)))
        total_pages = max(1, math.ceil(max(1, visual_units) / rows_per_page))

        def draw_header(page_number: int) -> float:
            top_y = page_h - margin
            canvas_obj.setFillColor(colors.HexColor("#FFFFFF"))
            canvas_obj.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas_obj.roundRect(margin, top_y - header_h, inner_w, header_h, 5, stroke=1, fill=1)
            self._draw_operator_logo_plate(canvas_obj, palette, logo_path, margin + 8, top_y - 23 * mm, 34 * mm, 15 * mm, radius=4, padding_x=3, padding_y=2)
            canvas_obj.setFillColor(colors.HexColor("#020617"))
            canvas_obj.setFont("Helvetica-Bold", 16)
            canvas_obj.drawString(margin + 48 * mm, top_y - 10 * mm, self._operator_pdf_text("Ordem de Fabrico"))
            canvas_obj.setFont("Helvetica", 7.5)
            client_line = f"{detail.get('cliente', '-') or '-'} - {detail.get('cliente_nome', '') or ''}".strip(" -")
            header_meta = [
                f"Encomenda: {detail.get('numero', '-')}",
                f"Cliente: {client_line or '-'}",
                f"Data OF: {str((detail.get('ordem_fabrico') or {}).get('data', '') or '')[:10]}",
                f"Pecas: {len(pieces)}",
            ]
            meta_y = top_y - 16 * mm
            for line in header_meta:
                canvas_obj.drawString(margin + 48 * mm, meta_y, self._operator_pdf_text(_pdf_clip_text(line, 80 * mm, "Helvetica", 7.5)))
                meta_y -= 4 * mm
            barcode_x = page_w - margin - 66 * mm
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.drawCentredString(barcode_x + 33 * mm, top_y - 8 * mm, self._operator_pdf_text(of_code))
            self._draw_code128_fit(canvas_obj, of_code, barcode_x, top_y - 23 * mm, 66 * mm, 13 * mm, min_bar_width=0.36, max_bar_width=0.82)
            canvas_obj.setFont("Helvetica", 7)
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.drawRightString(page_w - margin - 8, top_y - 30 * mm, self._operator_pdf_text(f"Pagina {page_number}/{total_pages}"))
            return top_y - header_h - 4 * mm

        def draw_table_header(y_pos: float) -> float:
            canvas_obj.setFillColor(colors.HexColor("#05004D"))
            canvas_obj.roundRect(margin, y_pos - table_header_h, inner_w, table_header_h, 3, stroke=0, fill=1)
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setFont("Helvetica-Bold", 6.8)
            x = margin
            for label, width in columns:
                canvas_obj.drawString(x + 2.2, y_pos - 5.2 * mm, self._operator_pdf_text(label))
                x += width
            return y_pos - table_header_h

        def draw_footer(page_number: int) -> None:
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.setFont("Helvetica", 6.8)
            canvas_obj.drawString(margin, margin - 2, self._operator_pdf_text(f"Impresso em {printed_at}"))
            canvas_obj.drawRightString(page_w - margin, margin - 2, self._operator_pdf_text(f"LUGEST | OF {of_code} | {page_number}/{total_pages}"))

        page_number = 1
        y = draw_table_header(draw_header(page_number))
        row_counter = 0

        def ensure_space(height: float) -> None:
            nonlocal page_number, y, row_counter
            if y - height >= margin + footer_h:
                return
            draw_footer(page_number)
            canvas_obj.showPage()
            page_number += 1
            y = draw_table_header(draw_header(page_number))
            row_counter = 0

        for material_group, esp_group, group_rows in grouped_pieces:
            ensure_space(group_h + row_h)
            group_code = f"GRP|{of_code}|{material_group}|{esp_group}"
            group_y = y - group_h
            canvas_obj.setFillColor(colors.HexColor("#EEF6FF"))
            canvas_obj.setStrokeColor(colors.HexColor("#B9D7F2"))
            canvas_obj.roundRect(margin, group_y, inner_w, group_h - 1, 3, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor("#0F172A"))
            canvas_obj.setFont("Helvetica-Bold", 8.2)
            canvas_obj.drawString(margin + 3, group_y + 6.2 * mm, self._operator_pdf_text(f"{material_group} | {esp_group} mm"))
            canvas_obj.setFont("Helvetica", 6.4)
            canvas_obj.drawString(margin + 3, group_y + 2.4 * mm, self._operator_pdf_text(f"{len(group_rows)} peça(s) nesta espessura"))
            self._draw_code128_fit(canvas_obj, group_code, page_w - margin - 86 * mm, group_y + 3.0 * mm, 82 * mm, 5.2 * mm, min_bar_width=0.26, max_bar_width=0.58)
            canvas_obj.setFont("Helvetica-Bold", 5.2)
            canvas_obj.drawCentredString(page_w - margin - 45 * mm, group_y + 1.1 * mm, self._operator_pdf_text(_pdf_clip_text(group_code, 82 * mm, "Helvetica-Bold", 5.2)))
            y = group_y
            row_counter += 1
            for piece in group_rows:
                ensure_space(row_h)
                row_y = y - row_h
                canvas_obj.setFillColor(colors.HexColor("#FFFFFF") if row_counter % 2 == 0 else colors.HexColor("#F8FAFC"))
                canvas_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                canvas_obj.rect(margin, row_y, inner_w, row_h, stroke=1, fill=1)
                values = [
                    str(piece.get("ref_interna", "-") or "-"),
                    str(piece.get("ref_externa", "-") or "-"),
                    str(piece.get("material", "-") or "-"),
                    str(piece.get("espessura", "-") or "-"),
                    str(piece.get("qtd_plan", "0") or "0"),
                    str(piece.get("operacoes", "-") or "-"),
                ]
                x = margin
                canvas_obj.setFillColor(colors.HexColor("#0F172A"))
                for col_index, value in enumerate(values):
                    width = columns[col_index][1]
                    font_name = "Helvetica-Bold" if col_index == 0 else "Helvetica"
                    font_size = 6.1 if col_index in {1, 5} else 6.5
                    canvas_obj.setFont(font_name, font_size)
                    clipped = _pdf_clip_text(value, width - 3.5, font_name, font_size)
                    if col_index in {3, 4}:
                        canvas_obj.drawRightString(x + width - 2.2, row_y + 2.3 * mm, self._operator_pdf_text(clipped))
                    else:
                        canvas_obj.drawString(x + 2.2, row_y + 2.3 * mm, self._operator_pdf_text(clipped))
                    x += width
                y = row_y
                row_counter += 1
        draw_footer(page_number)
        canvas_obj.save()
        return target

    def order_open_fabrication_pdf(self, numero: str) -> Path:
        path = self.order_fabrication_pdf(numero)
        try:
            os.startfile(str(path))
        except Exception:
            pass
        return path

    def operator_scan_code(self, code: str, current_posto: str = "Geral") -> dict[str, Any]:
        code_txt = str(code or "").strip()
        if not code_txt:
            raise ValueError("Código vazio.")
        posto_txt = str(current_posto or "").strip() or "Geral"
        if code_txt.upper().startswith("GRP|"):
            parts = code_txt.split("|")
            if len(parts) >= 4:
                of_code = parts[1].strip()
                material = parts[2].strip()
                espessura = parts[3].strip()
                enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == of_code), None)
                if enc is None:
                    raise ValueError("Grupo/espessura não encontrado.")
                return {
                    "tipo": "GRP",
                    "encomenda_numero": str(enc.get("numero", "") or "").strip(),
                    "of": of_code,
                    "material": material,
                    "espessura": espessura,
                }
        if code_txt.upper().startswith("OF-"):
            enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == code_txt), None)
            if enc is None:
                raise ValueError("OF não encontrada.")
            detail = self.order_detail(str(enc.get("numero", "") or ""))
            return {"tipo": "OF", "encomenda": detail, "pieces": list(detail.get("pieces", []) or [])}
        enc, piece = self._find_piece_by_opp(code_txt)
        ctx = self.operator_piece_context(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        pending_ops = list(ctx.get("pending_ops", []) or [])
        selected_op = ""
        posto_norm = self.desktop_main.norm_text(posto_txt)
        group_for_resource = self.workcenter_group_for_resource(posto_txt)
        for op_name in pending_ops:
            op_posto = self._operator_posto_for_operation(op_name)
            candidates = {self.desktop_main.norm_text(op_posto), self.desktop_main.norm_text(self.workcenter_group_for_resource(posto_txt, op_name)), self.desktop_main.norm_text(group_for_resource)}
            if posto_norm in candidates or any(token and token in self.desktop_main.norm_text(op_name) for token in posto_norm.split()):
                selected_op = op_name
                break
        if not selected_op and pending_ops:
            selected_op = pending_ops[0]
        return {
            "tipo": "OPP",
            "encomenda_numero": str(enc.get("numero", "") or "").strip(),
            "piece_id": str(piece.get("id", "") or "").strip(),
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "posto": posto_txt,
            "operacao": selected_op,
            "pending_ops": pending_ops,
            "context": ctx,
        }

    def order_piece_remove(self, numero: str, ref_interna: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: pecas bloqueadas.")
        ref_int = str(ref_interna or "").strip()
        found = False
        for mat in list(enc.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                before = len(list(esp.get("pecas", []) or []))
                esp["pecas"] = [row for row in list(esp.get("pecas", []) or []) if str(row.get("ref_interna", "") or "").strip() != ref_int]
                if len(list(esp.get("pecas", []) or [])) != before:
                    found = True
        if not found:
            raise ValueError("Pe?a n?o encontrada.")
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_stock_candidates(self, numero: str, material: str, espessura: str) -> list[dict[str, Any]]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        material_norm = self.encomendas_actions._norm_material(material)
        esp_norm = self.encomendas_actions._norm_espessura(espessura)
        rows = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
            if self._material_quality_is_blocked(stock):
                continue
            disponivel = self._parse_float(stock.get("quantidade", 0), 0) - self._parse_float(stock.get("reservado", 0), 0)
            if disponivel <= 0:
                continue
            if self.encomendas_actions._norm_material(stock.get("material")) != material_norm:
                continue
            if self.encomendas_actions._norm_espessura(stock.get("espessura")) != esp_norm:
                continue
            rows.append(
                {
                    "material_id": str(stock.get("id", "") or "").strip(),
                    "dimensao": f"{stock.get('comprimento', '')}x{stock.get('largura', '')}",
                    "disponivel": round(disponivel, 2),
                    "local": self._localizacao(stock),
                    "lote": str(stock.get("lote_fornecedor", "") or "").strip(),
                }
            )
        return rows

    def order_reserve_stock(self, numero: str, material: str, espessura: str, allocations: list[dict[str, Any]]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        if not str(material or "").strip() or not str(espessura or "").strip():
            raise ValueError("Selecione material e espessura antes de cativar.")
        any_saved = False
        for row in allocations or []:
            material_id = str((row or {}).get("material_id", "") or "").strip()
            quantidade = self._parse_float((row or {}).get("quantidade", 0), 0)
            if not material_id or quantidade <= 0:
                continue
            stock = next((m for m in list(self.ensure_data().get("materiais", []) or []) if str(m.get("id", "") or "").strip() == material_id), None)
            if stock is None:
                raise ValueError(f"Material n?o encontrado: {material_id}")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {material_id} bloqueado pela qualidade.")
            disponivel = self._parse_float(stock.get("quantidade", 0), 0) - self._parse_float(stock.get("reservado", 0), 0)
            if quantidade > disponivel:
                raise ValueError(f"Quantidade maior que o disponivel para {material_id}")
            stock["reservado"] = self._parse_float(stock.get("reservado", 0), 0) + quantidade
            stock["atualizado_em"] = self.desktop_main.now_iso()
            enc.setdefault("reservas", []).append(
                {
                    "material_id": material_id,
                    "material": stock.get("material"),
                    "espessura": stock.get("espessura"),
                    "quantidade": quantidade,
                }
            )
            self.desktop_main.log_stock(
                self.ensure_data(),
                "CATIVAR",
                f"{material_id} qtd={quantidade} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )
            any_saved = True
        if not any_saved:
            raise ValueError("Nenhuma quantidade definida.")
        enc["cativar"] = True
        self._save(force=True)
        return self.order_detail(numero)

    def order_release_stock(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        target = []
        keep = []
        for row in list(enc.get("reservas", []) or []):
            if self.encomendas_actions._match_material(row.get("material"), material) and self.encomendas_actions._norm_espessura(row.get("espessura")) == self.encomendas_actions._norm_espessura(espessura):
                target.append(row)
            else:
                keep.append(row)
        if not target:
            raise ValueError(f"Sem reservas para {material} esp. {espessura}.")
        for row in target:
            self.desktop_main.log_stock(
                self.ensure_data(),
                "LIBERTAR",
                f"{row.get('material_id', '')} qtd={row.get('quantidade', 0)} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )
        self.desktop_main.aplicar_reserva_em_stock(self.ensure_data(), target, -1)
        enc["reservas"] = keep
        enc["cativar"] = bool(enc.get("reservas"))
        self._save(force=True)
        return self.order_detail(numero)

    def order_consume_montagem(self, numero: str, operador: str = "") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        items = list(enc.get("montagem_itens", []) or [])
        if not items:
            raise ValueError("Esta encomenda nao tem itens de montagem.")
        actor = str(operador or (self.user or {}).get("username", "") or "Sistema").strip() or "Sistema"
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(self.ensure_data().get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        shortages: list[str] = []
        for item in items:
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            if self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                if stock_id:
                    material = self.material_by_id(stock_id)
                    if material is None:
                        shortages.append(f"{stock_id}: matéria-prima nao encontrada")
                        continue
                    if self._material_quality_is_blocked(material):
                        shortages.append(f"{stock_id}: bloqueado pela qualidade")
                        continue
                    available = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    candidates = self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []
                    available = sum(self._parse_float(row.get("disponivel", 0), 0) for row in candidates)
                if pending > available + 1e-9:
                    label = stock_id or str(item.get("descricao", "") or item.get("material", "") or "-").strip()
                    shortages.append(f"{label}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
                continue
            if self.desktop_main.normalize_orc_line_type(item.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            product = product_map.get(code)
            if product is None:
                shortages.append(f"{code or '-'}: produto nao encontrado")
                continue
            if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                shortages.append(f"{code}: bloqueado pela qualidade")
                continue
            available = self._parse_float(product.get("qty", 0), 0)
            if pending > available + 1e-9:
                shortages.append(f"{code}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
        if shortages:
            raise ValueError("Stock insuficiente para concluir a montagem:\n" + "\n".join(shortages))
        changed = False
        now_txt = self.desktop_main.now_iso()
        for item in items:
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            if item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                code = str(item.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                if product is None:
                    continue
                if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                    raise ValueError(f"Produto {code} bloqueado pela qualidade.")
                before = self._parse_float(product.get("qty", 0), 0)
                after = max(0.0, before - pending)
                product["qty"] = after
                product["atualizado_em"] = now_txt
                self.desktop_main.add_produto_mov(
                    self.ensure_data(),
                    tipo="BAIXA_MONTAGEM",
                    operador=actor,
                    codigo=code,
                    descricao=str(item.get("descricao", "") or product.get("descricao", "") or "").strip(),
                    qtd=pending,
                    antes=before,
                    depois=after,
                    obs=f"Montagem da encomenda {numero}",
                    origem="MONTAGEM",
                    ref_doc=numero,
                )
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Consumido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                changed = True
            elif self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                allocations: list[dict[str, Any]] = []
                remaining = pending
                if stock_id:
                    allocations.append({"material_id": stock_id, "quantidade": remaining})
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    for candidate in self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []:
                        if remaining <= 1e-9:
                            break
                        available = self._parse_float(candidate.get("disponivel", 0), 0)
                        qty = min(available, remaining)
                        if qty <= 1e-9:
                            continue
                        allocations.append({"material_id": str(candidate.get("material_id", "") or "").strip(), "quantidade": qty})
                        remaining = round(remaining - qty, 6)
                result = self.consume_material_allocations(
                    allocations,
                    reason=f"montagem_{numero}_{str(item.get('descricao', '') or item.get('material', '') or '').strip()}",
                )
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Consumido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                if not str(item.get("stock_material_id", "") or "").strip() and allocations:
                    item["stock_material_id"] = str(allocations[0].get("material_id", "") or "").strip()
                item["stock_consumption"] = {
                    "consumed_total": round(self._parse_float(result.get("consumed_total", pending), pending), 2),
                    "used_lots": list(result.get("used_lots", []) or []),
                }
                changed = True
            elif item_type == self.desktop_main.ORC_LINE_TYPE_SERVICE:
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Concluido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                changed = True
            else:
                # Componentes de conjunto sem produto nao consomem stock nem fecham aqui.
                continue
        if not changed:
            raise ValueError("Nao existem consumos pendentes de montagem.")
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

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

    def app_version(self) -> str:
        candidates = [
            self.base_dir / "VERSION",
            Path.cwd() / "VERSION",
        ]
        for path in candidates:
            try:
                if path.exists():
                    value = path.read_text(encoding="utf-8").strip()
                    if value:
                        return value
            except Exception:
                continue
        return "0.0.0"

    def update_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        stored = dict(cfg.get("update_settings", {}) or {})
        manifest_env = str(os.environ.get("LUGEST_UPDATE_MANIFEST_URL", "") or "").strip()
        defaults = {
            "current_version": self.app_version(),
            "manifest_url": manifest_env or "..\\Atualizacoes\\latest.json",
            "channel": "stable",
            "github_token": "",
            "auto_check": False,
        }
        return {**defaults, **stored, "current_version": self.app_version()}

    def update_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        current = dict(cfg.get("update_settings", {}) or {})
        for key in ("manifest_url", "channel", "github_token", "auto_check"):
            if key in dict(payload or {}):
                current[key] = payload.get(key)
        current["current_version"] = self.app_version()
        cfg["update_settings"] = current
        self._save_qt_config(cfg)
        return self.update_settings()

    def _update_version_parts(self, value: Any) -> tuple[int, int, int, int]:
        parts = [int(match.group(0)) for match in re.finditer(r"\d+", str(value or ""))]
        while len(parts) < 4:
            parts.append(0)
        return tuple(parts[:4])

    def _update_resolve_ref(self, value: Any, base: Path | None = None) -> str:
        txt = str(value or "").strip()
        if not txt:
            return ""
        if re.match(r"^https?://", txt, flags=re.IGNORECASE):
            return txt
        parsed = urllib.parse.urlparse(txt)
        if parsed.scheme.lower() == "file":
            return urllib.request.url2pathname(parsed.path)
        path = Path(txt)
        if path.is_absolute():
            return str(path)
        return str((base or self.base_dir) / path)

    def _update_github_headers(self, token: str = "", *, binary_asset: bool = False) -> dict[str, str]:
        headers: dict[str, str] = {}
        token_txt = str(token or "").strip()
        if token_txt:
            headers["Authorization"] = f"Bearer {token_txt}"
        headers["User-Agent"] = "LuisGEST-Updater"
        headers["Accept"] = "application/octet-stream" if binary_asset else "application/vnd.github+json"
        return headers

    def _update_resolve_github_release_asset_api_url(self, url: str, token: str = "") -> str:
        token_txt = str(token or "").strip()
        if not token_txt:
            return ""
        txt = str(url or "").strip()
        tag_match = re.match(
            r"^https://github\.com/(?P<owner>[^/]+)/(?P<repo>[^/]+)/releases/download/(?P<tag>[^/]+)/(?P<asset>[^/?#]+)$",
            txt,
            flags=re.IGNORECASE,
        )
        latest_match = re.match(
            r"^https://github\.com/(?P<owner>[^/]+)/(?P<repo>[^/]+)/releases/latest/download/(?P<asset>[^/?#]+)$",
            txt,
            flags=re.IGNORECASE,
        )
        match = tag_match or latest_match
        if match is None:
            return ""
        owner = str(match.group("owner") or "").strip()
        repo = str(match.group("repo") or "").strip()
        asset_name = urllib.parse.unquote(str(match.group("asset") or "").strip())
        if not owner or not repo or not asset_name:
            return ""
        api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        if tag_match is not None:
            tag = str(tag_match.group("tag") or "").strip()
            if not tag:
                return ""
            api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/tags/{tag}"
        request = urllib.request.Request(api_url, headers=self._update_github_headers(token_txt))
        with urllib.request.urlopen(request, timeout=12) as response:
            payload = json.loads(response.read().decode("utf-8", errors="ignore"))
        if isinstance(payload, dict):
            for asset in list(payload.get("assets", []) or []):
                if str(dict(asset).get("name", "") or "") == asset_name:
                    return str(dict(asset).get("url", "") or "").strip()
        return ""

    def _update_read_json_ref(self, ref: str) -> tuple[dict[str, Any], Path | None]:
        resolved = self._update_resolve_ref(ref)
        if not resolved:
            raise ValueError("Configura o URL/caminho do manifest de atualizacao.")
        if re.match(r"^https?://", resolved, flags=re.IGNORECASE):
            settings = self.update_settings()
            headers = {}
            token = str(settings.get("github_token", "") or "").strip()
            request_url = resolved
            if token:
                asset_api_url = self._update_resolve_github_release_asset_api_url(resolved, token)
                if asset_api_url:
                    request_url = asset_api_url
                    headers = self._update_github_headers(token, binary_asset=True)
                else:
                    headers["Authorization"] = f"Bearer {token}"
                    headers["User-Agent"] = "LuisGEST-Updater"
            request = urllib.request.Request(request_url, headers=headers)
            with urllib.request.urlopen(request, timeout=12) as response:
                payload = json.loads(response.read().decode("utf-8-sig", errors="ignore"))
            return (payload if isinstance(payload, dict) else {}, None)
        path = Path(resolved)
        if not path.exists():
            raise ValueError(f"Manifest nao encontrado: {path}")
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
        return (payload if isinstance(payload, dict) else {}, path)

    def _update_resolve_relative_ref(self, ref: str, base_ref: str, manifest_path: Path | None = None) -> str:
        ref_txt = str(ref or "").strip()
        if not ref_txt:
            return ""
        if re.match(r"^https?://", ref_txt, flags=re.IGNORECASE):
            return ref_txt
        if ref_txt.lower().startswith("file:///"):
            return str(Path(urllib.request.url2pathname(urllib.parse.urlparse(ref_txt).path)))
        base_txt = str(base_ref or "").strip()
        if base_txt and re.match(r"^https?://", base_txt, flags=re.IGNORECASE):
            return urllib.parse.urljoin(base_txt, ref_txt)
        if manifest_path is not None:
            return str((manifest_path.parent / ref_txt).resolve())
        return self._update_resolve_ref(ref_txt)

    def _update_download_ref_to_temp(self, ref: str, suffix: str = ".tmp") -> Path:
        resolved = self._update_resolve_ref(ref)
        if not resolved:
            raise ValueError("Referencia de atualizacao vazia.")
        temp_path = Path(tempfile.mkdtemp(prefix="lugest_update_bootstrap_")) / f"asset{suffix}"
        if re.match(r"^https?://", resolved, flags=re.IGNORECASE):
            settings = self.update_settings()
            headers = {}
            token = str(settings.get("github_token", "") or "").strip()
            request_url = resolved
            if token:
                asset_api_url = self._update_resolve_github_release_asset_api_url(resolved, token)
                if asset_api_url:
                    request_url = asset_api_url
                    headers = self._update_github_headers(token, binary_asset=True)
                else:
                    headers["Authorization"] = f"Bearer {token}"
                    headers["User-Agent"] = "LuisGEST-Updater"
            request = urllib.request.Request(request_url, headers=headers)
            with urllib.request.urlopen(request, timeout=20) as response, temp_path.open("wb") as handle:
                shutil.copyfileobj(response, handle)
            return temp_path
        source = Path(resolved)
        if not source.exists():
            raise ValueError(f"Ficheiro de atualizacao nao encontrado: {source}")
        shutil.copy2(source, temp_path)
        return temp_path

    def _update_download_ref_to_path(self, ref: str, target: Path) -> Path:
        downloaded = self._update_download_ref_to_temp(ref, suffix=target.suffix or ".tmp")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(downloaded, target)
        return target

    def update_check(self) -> dict[str, Any]:
        settings = self.update_settings()
        current_version = self.app_version()
        manifest, manifest_path = self._update_read_json_ref(str(settings.get("manifest_url", "") or ""))
        latest_version = str(manifest.get("version", "") or "").strip()
        if not latest_version:
            raise ValueError("Manifest sem campo 'version'.")
        package_url = str(manifest.get("package_url", "") or "").strip()
        if not package_url:
            raise ValueError("Manifest sem campo 'package_url'.")
        available = self._update_version_parts(latest_version) > self._update_version_parts(current_version)
        return {
            "current_version": current_version,
            "latest_version": latest_version,
            "update_available": available,
            "manifest_url": str(settings.get("manifest_url", "") or ""),
            "manifest_path": str(manifest_path or ""),
            "package_url": package_url,
            "bootstrap_url": str(manifest.get("bootstrap_url", "") or ""),
            "sha256": str(manifest.get("sha256", "") or ""),
            "notes": str(manifest.get("notes", "") or ""),
            "channel": str(manifest.get("channel", settings.get("channel", "stable")) or "stable"),
        }

    def update_installer_command(self) -> list[str]:
        settings = dict(self.update_settings() or {})
        powershell_exe = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
        manifest_url = str(settings.get("manifest_url", "") or "").strip()
        if not manifest_url:
            raise ValueError("Configura primeiro o manifest de atualizacao.")
        manifest, manifest_path = self._update_read_json_ref(manifest_url)
        bootstrap_ref = str(manifest.get("bootstrap_url", "") or "").strip() or "Reparar Atualizador Instalado.ps1"
        bootstrap_resolved = self._update_resolve_relative_ref(bootstrap_ref, manifest_url, manifest_path)
        local_repair_script = self.base_dir / "Reparar Atualizador Instalado.ps1"
        self._update_download_ref_to_path(bootstrap_resolved, local_repair_script)
        command = [
            str(powershell_exe),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(local_repair_script),
            "-InstallDir",
            str(self.base_dir),
            "-ManifestUrl",
            manifest_url,
            "-CurrentVersion",
            self.app_version(),
        ]
        token = str(settings.get("github_token", "") or "").strip()
        if token:
            command.extend(["-GitHubToken", token])
        return command

    def _update_sync_installer_config(self) -> Path:
        target = self.base_dir / "update_config.json"
        settings = dict(self.update_settings() or {})
        payload = {
            "current_version": self.app_version(),
            "manifest_url": str(settings.get("manifest_url", "") or "").strip(),
            "channel": str(settings.get("channel", "stable") or "stable").strip() or "stable",
            "github_token": str(settings.get("github_token", "") or "").strip(),
            "auto_check": bool(settings.get("auto_check", False)),
        }
        target.write_text(json.dumps(payload, indent=4, ensure_ascii=False) + "\n", encoding="utf-8")
        return target

    def update_start_installer(self) -> dict[str, Any]:
        config_path = self._update_sync_installer_config()
        command = self.update_installer_command()
        creationflags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        subprocess.Popen(command, cwd=str(self.base_dir), close_fds=True, creationflags=creationflags)
        return {"started": True, "command": command, "config_path": str(config_path)}

    def verify_supervisor_password(self, password: str) -> bool:
        stored = str(dict(self._load_qt_config().get("ui_options", {}) or {}).get("operator_supervisor_password", "") or "").strip()
        if not stored:
            return False
        return bool(self.desktop_main.verify_password(str(password or "").strip(), stored))

    def _material_assistant_feedback_map(self, *, persist_pruned: bool = True) -> dict[str, dict[str, Any]]:
        cfg = self._load_qt_config()
        raw = dict(cfg.get("material_assistant_feedback", {}) or {})
        today_txt = date.today().isoformat()
        cleaned: dict[str, dict[str, Any]] = {}
        changed = False
        for suggestion_id, payload in raw.items():
            if not str(suggestion_id or "").strip():
                changed = True
                continue
            meta = dict(payload or {})
            if str(meta.get("date", "") or "").strip() != today_txt:
                changed = True
                continue
            decision = str(meta.get("decision", "") or "").strip().lower()
            if decision not in {"accepted", "ignored"}:
                changed = True
                continue
            cleaned[str(suggestion_id).strip()] = {
                "decision": decision,
                "date": today_txt,
                "at": str(meta.get("at", "") or "").strip(),
            }
        if changed and persist_pruned:
            cfg["material_assistant_feedback"] = cleaned
            self._save_qt_config(cfg)
        return cleaned

    def material_assistant_feedback(self) -> dict[str, dict[str, Any]]:
        return self._material_assistant_feedback_map()

    def material_assistant_set_feedback(self, suggestion_id: str, decision: str) -> dict[str, dict[str, Any]]:
        suggestion_txt = str(suggestion_id or "").strip()
        if not suggestion_txt:
            raise ValueError("Sugestão inválida.")
        decision_txt = str(decision or "").strip().lower()
        if decision_txt in {"", "clear", "reset", "remove"}:
            cfg = self._load_qt_config()
            stored = self._material_assistant_feedback_map(persist_pruned=False)
            if suggestion_txt in stored:
                stored.pop(suggestion_txt, None)
                cfg["material_assistant_feedback"] = stored
                self._save_qt_config(cfg)
            return self.material_assistant_feedback()
        if decision_txt not in {"accepted", "ignored"}:
            raise ValueError("Decisão inválida.")
        cfg = self._load_qt_config()
        stored = self._material_assistant_feedback_map(persist_pruned=False)
        stored[suggestion_txt] = {
            "decision": decision_txt,
            "date": date.today().isoformat(),
            "at": str(self.desktop_main.now_iso() or "").strip(),
        }
        cfg["material_assistant_feedback"] = stored
        self._save_qt_config(cfg)
        return self.material_assistant_feedback()

    def _material_assistant_check_map(
        self,
        *,
        valid_keys: set[str] | None = None,
        persist_pruned: bool = True,
    ) -> dict[str, dict[str, Any]]:
        cfg = self._load_qt_config()
        raw = dict(cfg.get("material_assistant_checks", {}) or {})
        cleaned: dict[str, dict[str, Any]] = {}
        changed = False
        for need_key, payload in raw.items():
            key_txt = str(need_key or "").strip()
            if not key_txt:
                changed = True
                continue
            if valid_keys is not None and key_txt not in valid_keys:
                changed = True
                continue
            row = dict(payload or {})
            normalized = {
                "sep": bool(row.get("sep")),
                "conf": bool(row.get("conf")),
                "updated_at": str(row.get("updated_at", "") or "").strip(),
            }
            cleaned[key_txt] = normalized
            if row != normalized:
                changed = True
        if changed and persist_pruned:
            cfg["material_assistant_checks"] = cleaned
            self._save_qt_config(cfg)
        return cleaned

    def material_assistant_checks(self) -> dict[str, dict[str, Any]]:
        return self._material_assistant_check_map()

    def material_assistant_set_check(self, need_key: str, field: str, checked: bool) -> dict[str, dict[str, Any]]:
        need_key_txt = str(need_key or "").strip()
        if not need_key_txt:
            raise ValueError("Linha de separação inválida.")
        field_txt = str(field or "").strip().lower()
        if field_txt not in {"sep", "conf"}:
            raise ValueError("Campo de visto inválido.")
        cfg = self._load_qt_config()
        stored = self._material_assistant_check_map(persist_pruned=False)
        row = dict(stored.get(need_key_txt, {}) or {})
        row[field_txt] = bool(checked)
        row["updated_at"] = str(self.desktop_main.now_iso() or "").strip()
        stored[need_key_txt] = row
        cfg["material_assistant_checks"] = stored
        self._save_qt_config(cfg)
        return self.material_assistant_checks()

    def _pulse_plan_delay_reason_map(
        self,
        *,
        valid_keys: set[str] | None = None,
        persist_pruned: bool = True,
    ) -> dict[str, dict[str, Any]]:
        cfg = self._load_qt_config()
        raw = dict(cfg.get("pulse_plan_delay_reasons", {}) or {})
        cleaned: dict[str, dict[str, Any]] = {}
        changed = False
        for item_key, payload in raw.items():
            key_txt = str(item_key or "").strip()
            if not key_txt:
                changed = True
                continue
            if valid_keys is not None and key_txt not in valid_keys:
                changed = True
                continue
            row = dict(payload or {})
            reason_txt = str(row.get("reason", "") or "").strip()
            if not reason_txt:
                changed = True
                continue
            normalized = {
                "reason": reason_txt,
                "at": str(row.get("at", "") or "").strip(),
                "user": str(row.get("user", "") or "").strip(),
            }
            cleaned[key_txt] = normalized
            if row != normalized:
                changed = True
        if changed and persist_pruned:
            cfg["pulse_plan_delay_reasons"] = cleaned
            self._save_qt_config(cfg)
        return cleaned

    def pulse_plan_delay_reason_map(self) -> dict[str, dict[str, Any]]:
        return self._pulse_plan_delay_reason_map()

    def pulse_plan_delay_set_reason(self, item_key: str, reason: str) -> dict[str, dict[str, Any]]:
        item_key_txt = str(item_key or "").strip()
        reason_txt = str(reason or "").strip()
        if not item_key_txt:
            raise ValueError("Linha de atraso inválida.")
        if not reason_txt:
            raise ValueError("Indica o motivo da sinalização.")
        cfg = self._load_qt_config()
        stored = self._pulse_plan_delay_reason_map(persist_pruned=False)
        stored[item_key_txt] = {
            "reason": reason_txt,
            "at": str(self.desktop_main.now_iso() or "").strip(),
            "user": str((self.user or {}).get("username", "") or "").strip(),
        }
        cfg["pulse_plan_delay_reasons"] = stored
        self._save_qt_config(cfg)
        return self.pulse_plan_delay_reason_map()

    def pulse_plan_delay_clear_reason(self, item_key: str) -> dict[str, dict[str, Any]]:
        item_key_txt = str(item_key or "").strip()
        if not item_key_txt:
            return self.pulse_plan_delay_reason_map()
        cfg = self._load_qt_config()
        stored = self._pulse_plan_delay_reason_map(persist_pruned=False)
        if item_key_txt in stored:
            stored.pop(item_key_txt, None)
            cfg["pulse_plan_delay_reasons"] = stored
            self._save_qt_config(cfg)
        return self.pulse_plan_delay_reason_map()

    def pulse_plan_delay_rows(
        self,
        *,
        period_days: int = 7,
        year_filter: str | None = None,
        encomenda: str = "Todas",
    ) -> dict[str, Any]:
        data = self.ensure_data()
        now_dt = datetime.now()
        cutoff_date: date | None = None
        try:
            pd = int(period_days or 0)
        except Exception:
            pd = 0
        if pd > 0:
            cutoff_date = date.today() - timedelta(days=max(0, pd - 1))
        try:
            yf = int(str(year_filter or "").strip()) if str(year_filter or "").strip().isdigit() else None
        except Exception:
            yf = None
        enc_filter = str(encomenda or "Todas").strip()
        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if isinstance(row, dict)
        }
        rows_by_group: dict[tuple[str, str, str], dict[str, Any]] = {}
        for block in list(data.get("plano", []) or []):
            if not isinstance(block, dict):
                continue
            if not self._planning_row_matches_operation(block, "Corte Laser"):
                continue
            start_dt, end_dt = self._planning_block_bounds(block)
            if start_dt is None or end_dt is None:
                continue
            if yf is not None and int(start_dt.year) != int(yf):
                continue
            if cutoff_date is not None and start_dt.date() < cutoff_date:
                continue
            if end_dt > now_dt:
                continue
            numero = str(block.get("encomenda", "") or "").strip()
            material = str(block.get("material", "") or "").strip()
            espessura = str(block.get("espessura", "") or "").strip()
            if not numero or not material or not espessura:
                continue
            if enc_filter and enc_filter.lower() != "todas" and numero != enc_filter:
                continue
            if not self._planning_item_has_laser(numero, material, espessura):
                continue
            enc = self.get_encomenda_by_numero(numero)
            if not isinstance(enc, dict):
                continue
            enc_state = self.desktop_main.norm_text(enc.get("estado", ""))
            if "concl" in enc_state or "cancel" in enc_state:
                continue
            esp_obj = self._planning_find_esp_obj(enc, material, espessura)
            if not isinstance(esp_obj, dict):
                continue
            esp_state = self.desktop_main.norm_text(esp_obj.get("estado", ""))
            if "concl" in esp_state or "cancel" in esp_state:
                continue
            if self._operator_esp_laser_resolved(esp_obj):
                continue
            group_key = self._planning_item_key(numero, material, espessura)
            existing = rows_by_group.get(group_key)
            if existing is not None and start_dt >= existing["planned_end_dt"]:
                continue
            posto_txt = (
                str(block.get("posto", "") or "").strip()
                or str(block.get("posto_trabalho", "") or "").strip()
                or str(block.get("maquina", "") or "").strip()
                or self._order_workcenter(enc)
                or "Sem posto"
            )
            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = clients.get(client_code, "") or str(enc.get("cliente_nome", "") or "").strip()
            item_key = "|".join(
                [
                    numero,
                    material,
                    espessura,
                    start_dt.strftime("%Y-%m-%d"),
                    start_dt.strftime("%H:%M"),
                    str(block.get("id", "") or "").strip() or "sem-id",
                ]
            )
            rows_by_group[group_key] = {
                "item_key": item_key,
                "numero": numero,
                "cliente": " - ".join(part for part in (client_code, client_name) if part).strip(" -") or client_code or "-",
                "material": material,
                "espessura": espessura,
                "posto": posto_txt,
                "planned_start_dt": start_dt,
                "planned_end_dt": end_dt,
                "planned_start_txt": start_dt.strftime("%d/%m/%Y %H:%M"),
                "planned_end_txt": end_dt.strftime("%d/%m/%Y %H:%M"),
                "overdue_min": round(max(0.0, (now_dt - end_dt).total_seconds() / 60.0), 1),
                "baixa_estado": "Por dar baixa no corte laser",
            }
        rows = list(rows_by_group.values())
        valid_keys = {str(row.get("item_key", "") or "").strip() for row in rows if str(row.get("item_key", "") or "").strip()}
        reason_map = self._pulse_plan_delay_reason_map(valid_keys=valid_keys)
        open_count = 0
        acknowledged_count = 0
        for row in rows:
            item_key_txt = str(row.get("item_key", "") or "").strip()
            reason_row = dict(reason_map.get(item_key_txt, {}) or {})
            acknowledged = bool(reason_row)
            row["reason"] = str(reason_row.get("reason", "") or "").strip()
            row["reason_at"] = str(reason_row.get("at", "") or "").strip()
            row["reason_user"] = str(reason_row.get("user", "") or "").strip()
            row["acknowledged"] = acknowledged
            row["status_key"] = "acknowledged" if acknowledged else "open"
            row["status_label"] = "Justificado" if acknowledged else "Pendente"
            if acknowledged:
                acknowledged_count += 1
            else:
                open_count += 1
        rows.sort(
            key=lambda row: (
                0 if not bool(row.get("acknowledged")) else 1,
                row.get("planned_end_dt") or datetime.max,
                -self._parse_float(row.get("overdue_min", 0), 0),
                str(row.get("numero", "") or ""),
            )
        )
        return {
            "open_count": open_count,
            "acknowledged_count": acknowledged_count,
            "items": rows,
            "updated_at": str(self.desktop_main.now_iso() or "").strip(),
        }

    def available_menu_pages(self) -> list[dict[str, str]]:
        return [
            {"key": "stock_dashboard", "label": "Dashboard"},
            {"key": "materials", "label": "Matéria-Prima"},
            {"key": "products", "label": "Produtos"},
            {"key": "clients", "label": "Clientes"},
            {"key": "suppliers", "label": "Fornecedores"},
            {"key": "orders", "label": "Encomendas"},
            {"key": "quotes", "label": "Orçamentos"},
            {"key": "planning", "label": "Planeamento"},
            {"key": "transportes", "label": "Transportes"},
            {"key": "material_assistant", "label": "Assistente MP"},
            {"key": "operator", "label": "Operador"},
            {"key": "opp", "label": "OPP"},
            {"key": "shipping", "label": "Expedição"},
            {"key": "billing", "label": "Faturação"},
            {"key": "purchase_notes", "label": "Notas Encomenda"},
            {"key": "quality", "label": "Qualidade"},
            {"key": "pulse", "label": "Pulse"},
            {"key": "avarias", "label": "Avarias"},
            {"key": "home", "label": "Resumo"},
        ]

    def audit_rows(self, filter_text: str = "", limit: int = 500) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in reversed(list(self.ensure_data().get("audit_log", []) or [])):
            if not isinstance(raw, dict):
                continue
            row = {
                "created_at": str(raw.get("created_at", "") or "").strip(),
                "user": str(raw.get("user", "") or "").strip(),
                "action": str(raw.get("action", "") or "").strip(),
                "entity_type": str(raw.get("entity_type", "") or "").strip(),
                "entity_id": str(raw.get("entity_id", "") or "").strip(),
                "summary": str(raw.get("summary", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
            if len(rows) >= max(1, int(limit or 500)):
                break
        return rows

    def _next_prefixed_id(self, rows: list[Any], prefix: str, key: str = "id") -> str:
        max_seq = 0
        prefix_txt = str(prefix or "ID").strip().upper()
        for row in list(rows or []):
            if not isinstance(row, dict):
                continue
            raw = str(row.get(key, "") or "").strip().upper()
            if raw.startswith(f"{prefix_txt}-"):
                suffix = raw.split("-", 1)[1]
                if suffix.isdigit():
                    max_seq = max(max_seq, int(suffix))
        return f"{prefix_txt}-{max_seq + 1:04d}"

    def quality_summary(self) -> dict[str, Any]:
        data = self.ensure_data()
        ncs = [row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict)]
        docs = [row for row in list(data.get("quality_documents", []) or []) if isinstance(row, dict)]
        health = self.quality_data_health()
        open_ncs = [row for row in ncs if str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}]
        overdue = 0
        today = date.today().isoformat()
        for row in open_ncs:
            due = str(row.get("prazo", "") or "").strip()[:10]
            if due and due < today:
                overdue += 1
        supplier_ncs = [
            row
            for row in open_ncs
            if str(row.get("tipo", "") or "").strip().casefold() == "fornecedor"
            or str(row.get("fornecedor_id", "") or row.get("fornecedor_nome", "") or "").strip()
        ]
        blocked_materials = [
            row
            for row in self._quality_iter_delivery_movements(ensure_ids=False)
            if str(row.get("entity_type", "") or "") == "Material"
            and (
                self._parse_float(row.get("pending_qty", 0), 0) > 0
                or not self._quality_status_is_available(row.get("status", ""))
            )
        ]
        return {
            "open_nc": len(open_ncs),
            "overdue_nc": overdue,
            "supplier_nc": len(supplier_ncs),
            "blocked_materials": len(blocked_materials),
            "documents": len(docs),
            "audit_events": len(list(data.get("audit_log", []) or [])),
            "quality_issues": len(list(health.get("issues", []) or [])),
            "updated_at": str(self.desktop_main.now_iso() or "").strip(),
        }

    def quality_data_health(self) -> dict[str, Any]:
        data = self.ensure_data()
        known: dict[str, set[str]] = {
            "Encomenda": {str(row.get("numero", "") or "").strip() for row in list(data.get("encomendas", []) or []) if isinstance(row, dict)},
            "Material": {str(row.get("id", "") or "").strip() for row in list(data.get("materiais", []) or []) if isinstance(row, dict)},
            "Produto": {str(row.get("codigo", "") or "").strip() for row in list(data.get("produtos", []) or []) if isinstance(row, dict)},
            "Fornecedor": {
                str(row.get("id", "") or row.get("nome", "") or "").strip()
                for row in list(data.get("fornecedores", []) or [])
                if isinstance(row, dict)
            },
            "Cliente": {str(row.get("codigo", "") or row.get("nome", "") or "").strip() for row in list(data.get("clientes", []) or []) if isinstance(row, dict)},
            "Documento": {str(row.get("id", "") or row.get("titulo", "") or "").strip() for row in list(data.get("quality_documents", []) or []) if isinstance(row, dict)},
        }
        reception_keys = self._quality_reception_entity_keys()
        issues: list[dict[str, str]] = []
        for row in list(data.get("quality_nonconformities", []) or []):
            if not isinstance(row, dict):
                continue
            entity_type = str(row.get("entidade_tipo", "") or "").strip()
            entity_id = str(row.get("entidade_id", "") or "").strip()
            if entity_type and entity_type != "Livre" and entity_id:
                if entity_type in {"Material", "Produto"} and (entity_type, entity_id) not in reception_keys and str(row.get("origem", "") or "").strip().casefold() == "receção fornecedor":
                    continue
                known_ids = known.get(entity_type)
                if known_ids is not None and entity_id not in known_ids:
                    issues.append({"tipo": "NC", "id": str(row.get("id", "") or ""), "problema": f"Ligacao inexistente: {entity_type} {entity_id}"})
        for row in list(data.get("quality_documents", []) or []):
            if not isinstance(row, dict):
                continue
            path_txt = str(row.get("caminho", "") or "").strip()
            if path_txt and not Path(path_txt).exists():
                issues.append({"tipo": "Documento", "id": str(row.get("id", "") or ""), "problema": f"Ficheiro nao encontrado: {path_txt}"})
        open_nc_ids = {
            str(row.get("id", "") or "").strip()
            for row in list(data.get("quality_nonconformities", []) or [])
            if isinstance(row, dict) and str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}
        }
        for material in list(data.get("materiais", []) or []):
            if not isinstance(material, dict) or not self._material_quality_is_blocked(material):
                continue
            material_id = str(material.get("id", "") or "").strip()
            if ("Material", material_id) not in reception_keys:
                continue
            status_norm = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip().casefold()
            if "inspe" in status_norm and not any(token in status_norm for token in ("bloque", "reclam", "rejeit")):
                continue
            nc_id = str(material.get("quality_nc_id", "") or material.get("supplier_claim_id", "") or "").strip()
            if not nc_id:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": "Material bloqueado sem NC/reclamacao ligada."})
            elif nc_id not in open_nc_ids:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": f"Material bloqueado com NC inexistente/fechada: {nc_id}"})
        return {"issues": issues, "ok": not issues}

    def quality_link_options(self) -> dict[str, list[dict[str, str]]]:
        data = self.ensure_data()
        options: dict[str, list[dict[str, str]]] = {
            "Livre": [{"id": "", "label": ""}],
            "OPP": [],
            "Encomenda": [],
            "Material": [],
            "Produto": [],
            "Fornecedor": [],
            "Cliente": [],
            "Documento": [],
        }
        for row in list(data.get("plano", []) or []):
            if not isinstance(row, dict):
                continue
            opp = str(row.get("opp", "") or row.get("OPP", "") or "").strip()
            if opp:
                options["OPP"].append({"id": opp, "label": f"{opp} | {str(row.get('encomenda', '') or '').strip()}".strip(" |")})
        for enc in list(data.get("encomendas", []) or []):
            numero = str((enc or {}).get("numero", "") or "").strip()
            if numero:
                options["Encomenda"].append({"id": numero, "label": f"{numero} | {str((enc or {}).get('cliente', '') or '').strip()}"})
        for row in list(data.get("materiais", []) or []):
            if not isinstance(row, dict):
                continue
            material_id = str(row.get("id", "") or "").strip()
            if material_id:
                options["Material"].append(
                    {
                        "id": material_id,
                        "label": " | ".join(
                            part
                            for part in (
                                material_id,
                                str(row.get("material", "") or "").strip(),
                                str(row.get("espessura", "") or "").strip(),
                                str(row.get("formato", "") or "").strip(),
                            )
                            if part
                        ),
                    }
                )
        for row in list(data.get("produtos", []) or []):
            if not isinstance(row, dict):
                continue
            code = str(row.get("codigo", "") or "").strip()
            if code:
                options["Produto"].append({"id": code, "label": f"{code} | {str(row.get('descricao', '') or '').strip()}".strip(" |")})
        for row in list(data.get("fornecedores", []) or []):
            supplier_id = str((row or {}).get("id", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if supplier_id or name:
                options["Fornecedor"].append({"id": supplier_id or name, "label": f"{supplier_id} | {name}".strip(" |")})
        for row in list(data.get("clientes", []) or []):
            code = str((row or {}).get("codigo", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if code or name:
                options["Cliente"].append({"id": code or name, "label": f"{code} | {name}".strip(" |")})
        for row in list(data.get("quality_documents", []) or []):
            doc_id = str((row or {}).get("id", "") or "").strip()
            title = str((row or {}).get("titulo", "") or "").strip()
            if doc_id or title:
                options["Documento"].append({"id": doc_id or title, "label": f"{doc_id} | {title}".strip(" |")})
        for rows in options.values():
            rows.sort(key=lambda item: str(item.get("label", "") or item.get("id", "") or "").lower())
        return options

    def _quality_link_label(self, entity_type: str, entity_id: str) -> str:
        entity_type_txt = str(entity_type or "Livre").strip() or "Livre"
        entity_id_txt = str(entity_id or "").strip()
        if not entity_id_txt:
            return ""
        for row in list(self.quality_link_options().get(entity_type_txt, []) or []):
            if str(row.get("id", "") or "").strip() == entity_id_txt:
                return str(row.get("label", "") or "").strip()
        return entity_id_txt

    def _quality_status_code(self, value: Any) -> str:
        raw = str(value or "").strip().casefold()
        if "devol" in raw:
            return "DEVOLVER_FORNECEDOR"
        if "averig" in raw or "analise" in raw or "análise" in raw:
            return "EM_AVERIGUACAO"
        if "rejeit" in raw:
            return "REJEITADO"
        if "aprov" in raw:
            return "APROVADO"
        return "EM_INSPECAO"

    def _quality_status_is_available(self, value: Any) -> bool:
        return self._quality_status_code(value) == "APROVADO"

    def _quality_quarantine_pending_stock(self, item: dict[str, Any], *, kind: str, max_qty: float | None = None) -> bool:
        if not isinstance(item, dict):
            return False
        status = self._quality_status_code(item.get("quality_status", item.get("inspection_status", "")))
        if status == "APROVADO":
            return False
        qty_key = "qty" if str(kind or "").casefold().startswith("prod") else "quantidade"
        current_qty = self._parse_float(item.get(qty_key, 0), 0)
        pending = self._parse_float(item.get("quality_pending_qty", 0), 0)
        if current_qty <= 0:
            return False
        quarantine_qty = current_qty
        if max_qty is not None:
            quarantine_qty = min(current_qty, max(0.0, self._parse_float(max_qty, 0) - pending))
        if quarantine_qty <= 0:
            return False
        item["quality_pending_qty"] = pending + quarantine_qty
        item["quality_received_qty"] = max(self._parse_float(item.get("quality_received_qty", 0), 0), pending + quarantine_qty)
        item[qty_key] = max(0.0, current_qty - quarantine_qty)
        item["quality_blocked"] = True
        item["logistic_status"] = str(item.get("logistic_status", "") or "RECEBIDO").strip()
        item["atualizado_em"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        return True

    def _quality_movement_pending_qty(self, movement: dict[str, Any]) -> float:
        qty = self._parse_float(movement.get("qtd", 0), 0)
        approved = self._parse_float(movement.get("quality_approved_qty", 0), 0)
        rejected = self._parse_float(movement.get("quality_rejected_qty", 0), 0)
        status = self._quality_status_code(movement.get("quality_status", movement.get("inspection_status", "")))
        if approved <= 0 and rejected <= 0:
            if status == "APROVADO":
                approved = qty
            elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                rejected = qty
        return max(0.0, qty - approved - rejected)

    def _quality_delivery_movement_id(self, note_number: str, line_index: int, movement_index: int, entity_type: str, entity_id: str) -> str:
        return "|".join(
            str(part or "").strip()
            for part in (
                note_number,
                f"L{int(line_index or 0)}",
                f"M{int(movement_index or 0)}",
                entity_type,
                entity_id,
            )
        )

    def _quality_reception_entity_keys(self) -> set[tuple[str, str]]:
        return {
            (str(row.get("entity_type", "") or "").strip(), str(row.get("entity_id", "") or "").strip())
            for row in self._quality_iter_delivery_movements(ensure_ids=False)
            if str(row.get("entity_type", "") or "").strip() and str(row.get("entity_id", "") or "").strip()
        }

    def _quality_iter_delivery_movements(self, *, ensure_ids: bool = False) -> list[dict[str, Any]]:
        data = self.ensure_data()
        rows: list[dict[str, Any]] = []
        changed = False
        for note in list(data.get("notas_encomenda", []) or []):
            if not isinstance(note, dict):
                continue
            note_number = str(note.get("numero", "") or "").strip()
            for line_index, line in enumerate(list(note.get("linhas", []) or [])):
                if not isinstance(line, dict):
                    continue
                entity_type = "Material" if self.desktop_main.origem_is_materia(line.get("origem", "")) else "Produto"
                fallback_ref = str(line.get("ref", "") or "").strip()
                for movement_index, movement in enumerate(list(line.get("entregas_linha", []) or [])):
                    if not isinstance(movement, dict):
                        continue
                    entity_id = str(movement.get("stock_ref", "") or fallback_ref).strip()
                    if not entity_id:
                        continue
                    movement_id = str(movement.get("quality_movement_id", "") or "").strip()
                    if not movement_id:
                        movement_id = self._quality_delivery_movement_id(note_number, line_index, movement_index, entity_type, entity_id)
                    status = self._quality_status_code(
                        movement.get("quality_status", "")
                        or movement.get("inspection_status", "")
                        or line.get("quality_status", "")
                        or line.get("inspection_status", "")
                        or "EM_INSPECAO"
                    )
                    qty = self._parse_float(movement.get("qtd", 0), 0)
                    approved = self._parse_float(movement.get("quality_approved_qty", 0), 0)
                    rejected = self._parse_float(movement.get("quality_rejected_qty", 0), 0)
                    if approved <= 0 and rejected <= 0:
                        if status == "APROVADO":
                            approved = qty
                        elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                            rejected = qty
                        else:
                            open_nc = self._quality_find_open_nc(
                                {
                                    "origem": "Receção fornecedor",
                                    "referencia": note_number,
                                    "entidade_tipo": entity_type,
                                    "entidade_id": entity_id,
                                }
                            )
                            if open_nc is not None:
                                approved = min(qty, self._quality_nc_quantity(open_nc, "qtd_aprovada"))
                                rejected = min(qty - approved, self._quality_nc_quantity(open_nc, "qtd_rejeitada"))
                                if approved > 0 or rejected > 0:
                                    pending_guess = max(0.0, qty - approved - rejected)
                                    status = "EM_INSPECAO" if pending_guess > 0 else ("APROVADO" if approved > 0 else "REJEITADO")
                    pending = max(0.0, qty - approved - rejected)
                    rows.append(
                        {
                            "movement_id": movement_id,
                            "note": note,
                            "line": line,
                            "movement": movement,
                            "note_number": note_number,
                            "line_index": line_index,
                            "movement_index": movement_index,
                            "entity_type": entity_type,
                            "entity_id": entity_id,
                            "qty": qty,
                            "approved_qty": approved,
                            "rejected_qty": rejected,
                            "pending_qty": pending,
                            "status": status,
                        }
                    )
        if changed:
            self._save(force=True, audit=False)
        return rows

    def _quality_sync_pending_from_delivery_movements(self) -> None:
        movement_rows = self._quality_iter_delivery_movements(ensure_ids=True)
        data = self.ensure_data()
        totals: dict[tuple[str, str], dict[str, float]] = {}
        for row in movement_rows:
            key = (str(row.get("entity_type", "") or ""), str(row.get("entity_id", "") or ""))
            bucket = totals.setdefault(key, {"received": 0.0, "pending": 0.0, "approved": 0.0, "rejected": 0.0})
            bucket["received"] += self._parse_float(row.get("qty", 0), 0)
            bucket["pending"] += self._parse_float(row.get("pending_qty", 0), 0)
            bucket["approved"] += self._parse_float(row.get("approved_qty", 0), 0)
            bucket["rejected"] += self._parse_float(row.get("rejected_qty", 0), 0)
        changed = False

        def _apply(item: dict[str, Any], entity_type: str, entity_id: str) -> None:
            nonlocal changed
            total = totals.get((entity_type, entity_id))
            if not total:
                return
            pending = round(total["pending"], 4)
            received = round(total["received"], 4)
            approved = round(total["approved"], 4)
            rejected = round(total["rejected"], 4)
            updates = {
                "quality_pending_qty": pending,
                "quality_received_qty": received,
                "quality_approved_qty": approved,
                "quality_rejected_qty": rejected,
            }
            for key, value in updates.items():
                if abs(self._parse_float(item.get(key, 0), 0) - value) > 1e-6:
                    item[key] = value
                    changed = True
            current_status = self._quality_status_code(item.get("quality_status", item.get("inspection_status", "")))
            if pending > 0:
                target_status = "EM_AVERIGUACAO" if current_status == "EM_AVERIGUACAO" else "EM_INSPECAO"
                target_blocked = True
            elif approved > 0:
                target_status = "APROVADO"
                target_blocked = False
            elif rejected > 0:
                target_status = "REJEITADO"
                target_blocked = True
            else:
                target_status = current_status
                target_blocked = current_status != "APROVADO"
            if str(item.get("quality_status", "") or "").strip() != target_status:
                item["quality_status"] = target_status
                item["inspection_status"] = target_status
                changed = True
            if bool(item.get("quality_blocked")) != target_blocked:
                item["quality_blocked"] = target_blocked
                changed = True

        for material in list(data.get("materiais", []) or []):
            if isinstance(material, dict):
                _apply(material, "Material", str(material.get("id", "") or "").strip())
        for product in list(data.get("produtos", []) or []):
            if isinstance(product, dict):
                _apply(product, "Produto", str(product.get("codigo", "") or "").strip())
        if changed:
            self._save(force=True, audit=False)

    def _quality_reference_key(self, value: Any, fallback: Any = "") -> str:
        raw = str(value or fallback or "").strip()
        match = re.search(r"\bNE-\d{4}-\d{4}\b", raw, flags=re.IGNORECASE)
        if match:
            return match.group(0).upper()
        return re.sub(r"\s+", " ", raw).casefold()

    def _quality_nc_key(self, payload: dict[str, Any]) -> tuple[str, str, str, str]:
        origem_raw = str(payload.get("origem", "") or "").strip()
        origem = re.sub(r"\s+", " ", origem_raw).casefold()
        if "rece" in origem and "fornecedor" in origem:
            origem = "rececao fornecedor"
        referencia = self._quality_reference_key(payload.get("referencia", ""), payload.get("ne_numero", ""))
        entidade_tipo = str(payload.get("entidade_tipo", "") or payload.get("linked_entity_type", "") or "").strip()
        entidade_id = str(payload.get("entidade_id", "") or payload.get("linked_entity_id", "") or "").strip()
        if not entidade_tipo and str(payload.get("material_id", "") or "").strip():
            entidade_tipo = "Material"
            entidade_id = str(payload.get("material_id", "") or "").strip()
        if not entidade_id:
            entidade_id = str(payload.get("material_id", "") or payload.get("produto_codigo", "") or payload.get("fornecedor_id", "") or payload.get("fornecedor_nome", "") or "").strip()
        return (origem, referencia, entidade_tipo.casefold(), entidade_id.casefold())

    def _quality_is_open_nc(self, row: dict[str, Any]) -> bool:
        return str(row.get("estado", "") or "Aberta").strip().casefold() == "aberta"

    def _quality_find_open_nc(self, payload: dict[str, Any], *, exclude_id: str = "") -> dict[str, Any] | None:
        key = self._quality_nc_key(payload)
        exclude = str(exclude_id or "").strip()
        for row in list(self.ensure_data().get("quality_nonconformities", []) or []):
            if not isinstance(row, dict) or not self._quality_is_open_nc(row):
                continue
            if exclude and str(row.get("id", "") or "").strip() == exclude:
                continue
            if self._quality_nc_key(row) == key:
                return row
        return None

    def _quality_nc_quantity(self, row: dict[str, Any] | None, field: str) -> float:
        if not isinstance(row, dict):
            return 0.0
        for key in (field, field.replace("qtd_", "quality_"), field.replace("qtd_", "")):
            if key in row:
                value = self._parse_float(row.get(key, 0), 0)
                if value:
                    return value
        desc = str(row.get("descricao", "") or "")
        label = {
            "qtd_recebida": "recebido",
            "qtd_aprovada": "aprovado",
            "qtd_rejeitada": "rejeitado",
            "qtd_pendente": "pendente",
        }.get(field, field)
        match = re.search(rf"{re.escape(label)}\s*:\s*([0-9]+(?:[.,][0-9]+)?)", desc, flags=re.IGNORECASE)
        if match:
            return self._parse_float(match.group(1), 0)
        return 0.0

    def _quality_normalize_open_nc_duplicates(self) -> None:
        data = self.ensure_data()
        rows = [row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict)]
        first_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        changed = False
        for row in rows:
            if not self._quality_is_open_nc(row):
                continue
            key = self._quality_nc_key(row)
            if not all(key):
                continue
            keeper = first_by_key.get(key)
            if keeper is None:
                first_by_key[key] = row
                continue
            row["estado"] = "Cancelada"
            row["updated_at"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
            row["updated_by"] = self._current_user_label()
            row["acao"] = (
                str(row.get("acao", "") or "").strip()
                + f"\nCancelada automaticamente: NC duplicada de {str(keeper.get('id', '') or '').strip()}."
            ).strip()
            changed = True
        if changed:
            self._save(force=True, audit=False)

    def quality_reception_rows(self, filter_text: str = "", state_filter: str = "Pendentes") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Pendentes").strip().casefold()
        rows: list[dict[str, Any]] = []
        changed_quarantine = False

        def accept_status(status: str) -> bool:
            code = self._quality_status_code(status)
            if "todo" in state:
                return True
            if "aprov" in state:
                return code == "APROVADO"
            if "rejeit" in state:
                return code == "REJEITADO"
            if "devol" in state:
                return code == "DEVOLVER_FORNECEDOR"
            if "averig" in state:
                return code == "EM_AVERIGUACAO"
            return code == "EM_INSPECAO"

        for movement_row in self._quality_iter_delivery_movements(ensure_ids=True):
            entity_type = str(movement_row.get("entity_type", "") or "").strip()
            entity_id = str(movement_row.get("entity_id", "") or "").strip()
            if not entity_type or not entity_id:
                continue
            status = str(movement_row.get("status", "") or "EM_INSPECAO").strip()
            if not accept_status(status):
                continue
            pending_qty = self._parse_float(movement_row.get("pending_qty", 0), 0)
            if "pend" in state and pending_qty <= 0:
                continue
            target = (
                self.material_by_id(entity_id)
                if entity_type == "Material"
                else next((row for row in list(self.ensure_data().get("produtos", []) or []) if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == entity_id), None)
            )
            target = dict(target or {})
            line = dict(movement_row.get("line", {}) or {})
            movement = dict(movement_row.get("movement", {}) or {})
            row = {
                "tipo": entity_type,
                "id": entity_id,
                "ref": entity_id,
                "movement_id": str(movement_row.get("movement_id", "") or "").strip(),
                "referencia": str(movement_row.get("note_number", "") or "").strip(),
                "material": str(target.get("material", "") or target.get("categoria", "") or line.get("material", "") or line.get("categoria", "") or "").strip(),
                "espessura": str(target.get("espessura", "") or line.get("espessura", "") or "").strip(),
                "descricao": str(target.get("descricao", "") or line.get("descricao", "") or "").strip(),
                "lote": str(movement.get("lote_fornecedor", "") or target.get("lote_fornecedor", "") or "").strip(),
                "fornecedor": str(target.get("inspection_supplier_name", "") or target.get("fornecedor", "") or "").strip(),
                "fornecedor_id": str(target.get("inspection_supplier_id", "") or target.get("fornecedor_id", "") or "").strip(),
                "logistic_status": str(movement.get("logistic_status", "") or target.get("logistic_status", "") or "RECEBIDO").strip(),
                "quality_status": self._quality_status_code(status),
                "defeito": str(movement.get("inspection_defect", "") or target.get("inspection_defect", "") or "").strip(),
                "decisao": str(movement.get("inspection_decision", "") or target.get("inspection_decision", "") or "").strip(),
                "qtd": round(pending_qty, 4),
                "qtd_recebida": round(self._parse_float(movement_row.get("qty", 0), 0), 4),
                "qtd_aprovada": round(self._parse_float(movement_row.get("approved_qty", 0), 0), 4),
                "qtd_rejeitada": round(self._parse_float(movement_row.get("rejected_qty", 0), 0), 4),
                "qtd_disponivel": self._parse_float(target.get("qty" if entity_type == "Produto" else "quantidade", 0), 0),
                "nc_id": str(movement.get("quality_nc_id", "") or target.get("quality_nc_id", "") or target.get("supplier_claim_id", "") or "").strip(),
                "guia": str(movement.get("guia", "") or target.get("inspection_guia", "") or "").strip(),
                "fatura": str(movement.get("fatura", "") or target.get("inspection_fatura", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (str(item.get("quality_status", "")) == "APROVADO", str(item.get("referencia", "")), str(item.get("ref", ""))))
        if changed_quarantine:
            self._save(force=True, audit=False)
        return rows

    def quality_reception_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._quality_sync_pending_from_delivery_movements()
        data = self.ensure_data()
        item_type = str(payload.get("tipo", payload.get("kind", "")) or "").strip().casefold()
        item_id = str(payload.get("id", payload.get("ref", "")) or "").strip()
        movement_id = str(payload.get("movement_id", "") or "").strip()
        status = self._quality_status_code(payload.get("quality_status", payload.get("inspection_status", "")))
        defect = str(payload.get("defeito", payload.get("inspection_defect", "")) or "").strip()
        decision = str(payload.get("decisao", payload.get("inspection_decision", "")) or "").strip()
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        if not item_id:
            raise ValueError("Seleciona uma linha de receção para avaliar.")
        if not movement_id:
            raise ValueError("A qualidade só pode avaliar movimentos de receção provenientes de notas de encomenda.")

        movement_ctx = None
        movement_ctx = next(
            (
                row
                for row in self._quality_iter_delivery_movements(ensure_ids=True)
                if str(row.get("movement_id", "") or "").strip() == movement_id
            ),
            None,
        )
        if movement_ctx is None:
            raise ValueError("Movimento de receção não encontrado.")
        item_type = str(movement_ctx.get("entity_type", "") or item_type).casefold()
        item_id = str(movement_ctx.get("entity_id", "") or item_id).strip()

        if "prod" in item_type:
            target = next((row for row in list(data.get("produtos", []) or []) if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == item_id), None)
            entity_type = "Produto"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or "").strip()
            entity_label = " | ".join(part for part in (entity_id, str((target or {}).get("descricao", "") or "").strip()) if part)
        else:
            target = self.material_by_id(item_id)
            entity_type = "Material"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or (target or {}).get("origem_ne", "") or "").strip()
            entity_label = self._quality_link_label(entity_type, entity_id)
        if target is None:
            raise ValueError("Linha de receção não encontrada.")

        self._quality_quarantine_pending_stock(
            target,
            kind=entity_type,
            max_qty=self._parse_float((movement_ctx or {}).get("pending_qty", 0), 0) if movement_ctx is not None else None,
        )
        before = copy.deepcopy(target)
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = (
            self._parse_float((movement_ctx or {}).get("pending_qty", 0), 0)
            if movement_ctx is not None
            else self._parse_float(target.get("quality_pending_qty", 0), 0)
        )
        received_qty = (
            self._parse_float((movement_ctx or {}).get("qty", pending_qty), pending_qty)
            if movement_ctx is not None
            else self._parse_float(target.get("quality_received_qty", pending_qty), pending_qty)
        )
        approved_qty = self._parse_float(payload.get("qtd_aprovada", payload.get("approved_qty", None)), -1)
        rejected_qty = self._parse_float(payload.get("qtd_rejeitada", payload.get("rejected_qty", None)), -1)
        if approved_qty < 0 and rejected_qty < 0:
            approved_qty = pending_qty if status == "APROVADO" else 0.0
            rejected_qty = pending_qty if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} else 0.0
        else:
            approved_qty = max(0.0, approved_qty)
            rejected_qty = max(0.0, rejected_qty)
        if approved_qty + rejected_qty > pending_qty + 1e-9:
            raise ValueError(
                "As quantidades aprovadas/rejeitadas não podem ultrapassar a quantidade pendente "
                f"({self._fmt(pending_qty)})."
            )
        if status == "APROVADO" and rejected_qty <= 0:
            defect = ""
            decision = decision or "Libertar para stock"
        remaining_qty = max(0.0, pending_qty - approved_qty - rejected_qty)
        if status == "APROVADO" and approved_qty <= 0 and pending_qty > 0:
            raise ValueError("Para aprovar, indica a quantidade boa a libertar para stock.")
        if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} and rejected_qty <= 0 and pending_qty > 0:
            raise ValueError("Para rejeitar/devolver, indica a quantidade rejeitada.")
        available_before = self._parse_float(target.get(qty_key, 0), 0)
        target["logistic_status"] = str(target.get("logistic_status", "") or "RECEBIDO").strip()
        effective_status = status
        if remaining_qty > 0:
            effective_status = "EM_AVERIGUACAO" if status == "EM_AVERIGUACAO" else "EM_INSPECAO"
        elif approved_qty > 0:
            effective_status = "APROVADO"
        elif rejected_qty > 0:
            effective_status = status if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} else "REJEITADO"
        target["quality_status"] = effective_status
        target["inspection_status"] = effective_status
        target["inspection_defect"] = defect
        target["inspection_decision"] = decision or ("Libertar para stock" if effective_status == "APROVADO" else "Aguardar decisão da qualidade")
        target["inspection_at"] = now
        target["inspection_by"] = self._current_user_label()
        target["quality_blocked"] = effective_status != "APROVADO"
        target["atualizado_em"] = now
        target["quality_last_received_qty"] = round(received_qty, 4)
        target["quality_last_approved_qty"] = round(approved_qty, 4)
        target["quality_last_rejected_qty"] = round(rejected_qty, 4)

        if movement_ctx is not None:
            movement = movement_ctx["movement"]
            movement["inspection_status"] = effective_status
            movement["quality_status"] = effective_status
            movement["inspection_defect"] = defect
            movement["inspection_decision"] = target["inspection_decision"]
            movement["inspection_at"] = now
            movement["inspection_by"] = self._current_user_label()
            movement["quality_movement_id"] = movement_id
            movement["stock_ref"] = entity_id
            movement["quality_approved_qty"] = self._parse_float(movement_ctx.get("approved_qty", 0), 0) + approved_qty
            movement["quality_rejected_qty"] = self._parse_float(movement_ctx.get("rejected_qty", 0), 0) + rejected_qty
            movement["quality_pending_qty"] = remaining_qty

        nc_payload = {
            "origem": "Receção fornecedor",
            "referencia": reference,
            "entidade_tipo": entity_type,
            "entidade_id": entity_id,
            "entidade_label": entity_label,
            "tipo": "Fornecedor",
            "gravidade": "Alta" if status == "REJEITADO" else "Media",
            "estado": "Aberta",
            "responsavel": "Qualidade",
            "descricao": (
                f"Avaliação de receção marcada como {status}. "
                f"Entidade: {entity_label or entity_id}. "
                f"Recebido: {self._fmt(received_qty)} | aprovado: {self._fmt(approved_qty)} | rejeitado: {self._fmt(rejected_qty)}. "
                f"Defeito/observação: {defect or '-'}."
            ),
            "causa": "A apurar com fornecedor/receção.",
            "acao": target["inspection_decision"],
            "fornecedor_id": str(target.get("inspection_supplier_id", "") or target.get("fornecedor_id", "") or "").strip(),
            "fornecedor_nome": str(target.get("inspection_supplier_name", "") or target.get("fornecedor", "") or "").strip(),
            "material_id": entity_id if entity_type == "Material" else "",
            "produto_codigo": entity_id if entity_type == "Produto" else "",
            "lote_fornecedor": str(target.get("lote_fornecedor", "") or "").strip(),
            "ne_numero": reference,
            "guia": str(target.get("inspection_guia", "") or "").strip(),
            "fatura": str(target.get("inspection_fatura", "") or "").strip(),
            "decisao": target["inspection_decision"],
            "movement_id": movement_id,
            "qtd_recebida": received_qty,
            "qtd_aprovada": approved_qty,
            "qtd_rejeitada": rejected_qty,
            "qtd_pendente": remaining_qty,
        }
        existing_open = self._quality_find_open_nc(nc_payload)
        existing_rejected_total = self._quality_nc_quantity(existing_open, "qtd_rejeitada") if existing_open is not None else 0.0
        if existing_open is not None:
            nc_payload["qtd_recebida"] = max(self._quality_nc_quantity(existing_open, "qtd_recebida"), received_qty)
            nc_payload["qtd_aprovada"] = self._quality_nc_quantity(existing_open, "qtd_aprovada") + approved_qty
            nc_payload["qtd_rejeitada"] = self._quality_nc_quantity(existing_open, "qtd_rejeitada") + rejected_qty
            nc_payload["qtd_pendente"] = remaining_qty
        if approved_qty > 0:
                target[qty_key] = available_before + approved_qty
                target["quality_approved_qty"] = self._parse_float(target.get("quality_approved_qty", 0), 0) + approved_qty
                if entity_type == "Produto":
                    self.desktop_main.add_produto_mov(
                        data,
                        tipo="Entrada",
                        operador=self._current_user_label(),
                        codigo=entity_id,
                        descricao=str(target.get("descricao", "") or "").strip(),
                        qtd=approved_qty,
                        antes=available_before,
                        depois=target[qty_key],
                        obs=f"Aprovado pela qualidade | {reference}",
                        origem="Qualidade",
                        ref_doc=reference,
                    )
                else:
                    self.desktop_main.log_stock(
                        data,
                        "ENTRADA_QUALIDADE",
                        f"{entity_id} qtd={approved_qty} ref={reference}",
                        operador=self._current_user_label(),
                    )
        if effective_status == "APROVADO" and rejected_qty <= 0 and existing_rejected_total <= 0:
            if existing_open is not None:
                self.quality_nc_close(str(existing_open.get("id", "") or ""), target["inspection_decision"])
            if entity_type == "Material":
                target["supplier_claim_id"] = ""
            target["quality_nc_id"] = ""
            self._append_audit_event(data, action="Receção aprovada", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)
        else:
            if rejected_qty > 0 or (existing_rejected_total > 0 and approved_qty > 0) or status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} or defect or bool(payload.get("create_nc")):
                if existing_open is not None:
                    nc_payload["id"] = str(existing_open.get("id", "") or "").strip()
                nc = self.quality_nc_save(nc_payload)
                target["quality_nc_id"] = str(nc.get("id", "") or "").strip()
                if movement_ctx is not None:
                    movement_ctx["movement"]["quality_nc_id"] = target["quality_nc_id"]
                if entity_type == "Material":
                    target["supplier_claim_id"] = target["quality_nc_id"]
            if status == "DEVOLVER_FORNECEDOR" or "devol" in target["inspection_decision"].casefold():
                doc = self._quality_return_document(target, entity_type=entity_type, entity_id=entity_id, reference=reference, nc_id=str(target.get("quality_nc_id", "") or ""))
                if doc:
                    target["quality_return_document_id"] = str(doc.get("id", "") or "").strip()
            self._append_audit_event(data, action="Receção em inspeção", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)

        for note in list(data.get("notas_encomenda", []) or []):
            if not isinstance(note, dict):
                continue
            for line in list(note.get("linhas", []) or []):
                if not isinstance(line, dict):
                    continue
                line_ref = str(line.get("ref", "") or "").strip()
                if line_ref != entity_id:
                    continue
                line["quality_status"] = status
                line["inspection_status"] = status
                line["inspection_defect"] = defect
                line["inspection_decision"] = target["inspection_decision"]
                line["quality_nc_id"] = str(target.get("quality_nc_id", "") or "").strip()
                line["_stock_in"] = effective_status == "APROVADO"
                for movement in list(line.get("entregas_linha", []) or []):
                    if not isinstance(movement, dict):
                        continue
                    movement_matches = (
                        movement_id
                        and str(movement.get("quality_movement_id", "") or "").strip() == movement_id
                    ) or (
                        not movement_id
                        and str(movement.get("stock_ref", "") or line_ref).strip() == entity_id
                    )
                    if movement_matches:
                        movement["quality_status"] = effective_status
                        movement["quality_nc_id"] = line["quality_nc_id"]
                line_movements = [mv for mv in list(line.get("entregas_linha", []) or []) if isinstance(mv, dict)]
                if line_movements:
                    pending_line = sum(self._quality_movement_pending_qty(mv) for mv in line_movements)
                    approved_line = sum(self._parse_float(mv.get("quality_approved_qty", 0), 0) for mv in line_movements)
                    rejected_line = sum(self._parse_float(mv.get("quality_rejected_qty", 0), 0) for mv in line_movements)
                    if pending_line > 0:
                        line_status = "EM_INSPECAO"
                    elif approved_line > 0:
                        line_status = "APROVADO"
                    elif rejected_line > 0:
                        line_status = "REJEITADO"
                    else:
                        line_status = effective_status
                    line["quality_status"] = line_status
                    line["inspection_status"] = line_status
                    line["_stock_in"] = line_status == "APROVADO"
        self._sync_ne_from_materia()
        self._quality_sync_pending_from_delivery_movements()
        self._save(force=True, audit=False)
        return {"tipo": entity_type, "id": entity_id, "quality_status": effective_status, "quality_nc_id": str(target.get("quality_nc_id", "") or "").strip()}

    def _quality_return_document(
        self,
        target: dict[str, Any],
        *,
        entity_type: str,
        entity_id: str,
        reference: str,
        nc_id: str = "",
    ) -> dict[str, Any] | None:
        data = self.ensure_data()
        docs = data.setdefault("quality_documents", [])
        existing_id = str(target.get("quality_return_document_id", "") or "").strip()
        if existing_id:
            existing = next((row for row in docs if isinstance(row, dict) and str(row.get("id", "") or "").strip() == existing_id), None)
            if isinstance(existing, dict):
                return existing
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        doc_id = self._next_prefixed_id(docs, "DEV")
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = self._parse_float(target.get("quality_pending_qty", 0), 0)
        description = " | ".join(
            part
            for part in (
                f"{entity_type} {entity_id}",
                str(target.get("material", "") or target.get("descricao", "") or "").strip(),
                f"Qtd a devolver {self._fmt(pending_qty)}",
                f"Fornecedor {str(target.get('inspection_supplier_name', '') or target.get('fornecedor', '') or '-').strip()}",
                f"NC {nc_id}" if nc_id else "",
            )
            if part
        )
        doc = {
            "id": doc_id,
            "titulo": f"Nota devolução fornecedor {reference or entity_id}",
            "tipo": "Nota devolução fornecedor",
            "entidade": entity_type,
            "entidade_tipo": entity_type,
            "referencia": reference,
            "entidade_id": entity_id,
            "versao": "1",
            "estado": "Rascunho",
            "responsavel": "Qualidade",
            "caminho": "",
            "obs": description,
            "created_at": now,
            "updated_at": now,
            "created_by": self._current_user_label(),
            "nc_id": nc_id,
            "qtd": pending_qty,
            "qtd_stock": self._parse_float(target.get(qty_key, 0), 0),
        }
        docs.append(doc)
        return doc

    def quality_nc_rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        self._quality_normalize_open_nc_duplicates()
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Ativas").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("quality_nonconformities", []) or []):
            if not isinstance(raw, dict):
                continue
            row = dict(raw)
            estado = str(row.get("estado", "") or "Aberta").strip() or "Aberta"
            estado_norm = estado.lower()
            if state not in {"todos", "todas", "all"}:
                if "ativ" in state and estado_norm in {"fechada", "cancelada"}:
                    continue
                if "abert" in state and estado_norm != "aberta":
                    continue
                if "trat" in state and "trat" not in estado_norm:
                    continue
                if "fech" in state and estado_norm != "fechada":
                    continue
            emitted = {
                "id": str(row.get("id", "") or "").strip(),
                "origem": str(row.get("origem", "") or "").strip(),
                "referencia": str(row.get("referencia", "") or "").strip(),
                "entidade_tipo": str(row.get("entidade_tipo", "") or row.get("linked_entity_type", "") or "").strip(),
                "entidade_id": str(row.get("entidade_id", "") or row.get("linked_entity_id", "") or "").strip(),
                "entidade_label": str(row.get("entidade_label", "") or row.get("linked_entity_label", "") or "").strip(),
                "tipo": str(row.get("tipo", "") or "").strip(),
                "gravidade": str(row.get("gravidade", "") or "Media").strip(),
                "estado": estado,
                "responsavel": str(row.get("responsavel", "") or "").strip(),
                "prazo": str(row.get("prazo", "") or "").strip()[:10],
                "descricao": str(row.get("descricao", "") or "").strip(),
                "causa": str(row.get("causa", "") or "").strip(),
                "acao": str(row.get("acao", "") or "").strip(),
                "eficacia": str(row.get("eficacia", "") or "").strip(),
                "fornecedor_id": str(row.get("fornecedor_id", "") or "").strip(),
                "fornecedor_nome": str(row.get("fornecedor_nome", "") or "").strip(),
                "material_id": str(row.get("material_id", "") or "").strip(),
                "lote_fornecedor": str(row.get("lote_fornecedor", "") or "").strip(),
                "ne_numero": str(row.get("ne_numero", "") or "").strip(),
                "decisao": str(row.get("decisao", "") or "").strip(),
                "movement_id": str(row.get("movement_id", "") or "").strip(),
                "qtd_recebida": round(self._quality_nc_quantity(row, "qtd_recebida"), 4),
                "qtd_aprovada": round(self._quality_nc_quantity(row, "qtd_aprovada"), 4),
                "qtd_rejeitada": round(self._quality_nc_quantity(row, "qtd_rejeitada"), 4),
                "qtd_pendente": round(self._quality_nc_quantity(row, "qtd_pendente"), 4),
                "created_at": str(row.get("created_at", "") or "").strip(),
                "closed_at": str(row.get("closed_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in emitted.values()):
                continue
            rows.append(emitted)
        rows.sort(key=lambda item: (str(item.get("estado", "")) == "Fechada", str(item.get("prazo", "") or "9999"), str(item.get("id", ""))), reverse=False)
        return rows

    def quality_nc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("quality_nonconformities", [])
        nc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id), None) if nc_id else None
        before = copy.deepcopy(existing) if isinstance(existing, dict) else None
        if not nc_id:
            nc_id = self._next_prefixed_id(rows, "NC")
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        row = {
            "id": nc_id,
            "origem": str(payload.get("origem", "") or "").strip(),
            "referencia": str(payload.get("referencia", "") or "").strip(),
            "entidade_tipo": str(payload.get("entidade_tipo", payload.get("linked_entity_type", "")) or "").strip(),
            "entidade_id": str(payload.get("entidade_id", payload.get("linked_entity_id", "")) or "").strip(),
            "tipo": str(payload.get("tipo", "") or "Processo").strip() or "Processo",
            "gravidade": str(payload.get("gravidade", "") or "Media").strip() or "Media",
            "estado": str(payload.get("estado", "") or (existing or {}).get("estado", "Aberta") or "Aberta").strip() or "Aberta",
            "responsavel": str(payload.get("responsavel", "") or "").strip(),
            "prazo": str(payload.get("prazo", "") or "").strip()[:10],
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "causa": str(payload.get("causa", "") or "").strip(),
            "acao": str(payload.get("acao", "") or "").strip(),
            "eficacia": str(payload.get("eficacia", "") or "").strip(),
            "fornecedor_id": str(payload.get("fornecedor_id", (existing or {}).get("fornecedor_id", "")) or "").strip(),
            "fornecedor_nome": str(payload.get("fornecedor_nome", (existing or {}).get("fornecedor_nome", "")) or "").strip(),
            "material_id": str(payload.get("material_id", (existing or {}).get("material_id", "")) or "").strip(),
            "lote_fornecedor": str(payload.get("lote_fornecedor", (existing or {}).get("lote_fornecedor", "")) or "").strip(),
            "ne_numero": str(payload.get("ne_numero", (existing or {}).get("ne_numero", "")) or "").strip(),
            "guia": str(payload.get("guia", (existing or {}).get("guia", "")) or "").strip(),
            "fatura": str(payload.get("fatura", (existing or {}).get("fatura", "")) or "").strip(),
            "decisao": str(payload.get("decisao", (existing or {}).get("decisao", "")) or "").strip(),
            "movement_id": str(payload.get("movement_id", (existing or {}).get("movement_id", "")) or "").strip(),
            "qtd_recebida": round(self._parse_float(payload.get("qtd_recebida", (existing or {}).get("qtd_recebida", 0)), 0), 4),
            "qtd_aprovada": round(self._parse_float(payload.get("qtd_aprovada", (existing or {}).get("qtd_aprovada", 0)), 0), 4),
            "qtd_rejeitada": round(self._parse_float(payload.get("qtd_rejeitada", (existing or {}).get("qtd_rejeitada", 0)), 0), 4),
            "qtd_pendente": round(self._parse_float(payload.get("qtd_pendente", (existing or {}).get("qtd_pendente", 0)), 0), 4),
            "created_at": str((existing or {}).get("created_at", "") or now),
            "updated_at": now,
            "created_by": str((existing or {}).get("created_by", "") or self._current_user_label()),
            "updated_by": self._current_user_label(),
            "closed_at": str((existing or {}).get("closed_at", "") or "").strip(),
        }
        row["entidade_label"] = str(payload.get("entidade_label", "") or "").strip() or self._quality_link_label(
            row["entidade_tipo"], row["entidade_id"]
        )
        if not row["referencia"] and row["entidade_id"]:
            row["referencia"] = row["entidade_id"]
        if self._quality_is_open_nc(row):
            duplicate = self._quality_find_open_nc(row, exclude_id=nc_id)
            if duplicate is not None:
                dup_id = str(duplicate.get("id", "") or "").strip()
                raise ValueError(
                    f"Já existe uma NC aberta ({dup_id}) para esta origem, referência e entidade. "
                    "Fecha ou edita essa NC antes de criar outra."
                )
        if existing is None:
            rows.append(row)
        else:
            existing.update(row)
            row = existing
        self._append_audit_event(
            data,
            action="NC guardada",
            entity_type="Nao conformidade",
            entity_id=nc_id,
            summary=f"{row.get('tipo', '')} | {row.get('estado', '')} | {row.get('referencia', '')}",
            before=before,
            after=row,
        )
        self._save(force=True, audit=False)
        return dict(row)

    def quality_nc_close(self, nc_id: str, eficacia: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        target = next((row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict) and str(row.get("id", "") or "").strip() == str(nc_id or "").strip()), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        before = copy.deepcopy(target)
        target["estado"] = "Fechada"
        target["closed_at"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        target["closed_by"] = self._current_user_label()
        if str(eficacia or "").strip():
            target["eficacia"] = str(eficacia or "").strip()
        self._append_audit_event(data, action="NC fechada", entity_type="Nao conformidade", entity_id=str(nc_id), summary=str(target.get("eficacia", "") or ""), before=before, after=target)
        self._save(force=True, audit=False)
        return dict(target)

    def quality_nc_release_material(self, nc_id: str, decision: str = "Aprovado pela qualidade") -> dict[str, Any]:
        data = self.ensure_data()
        nc_id_txt = str(nc_id or "").strip()
        target = next((row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id_txt), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        material_id = str(target.get("material_id", "") or "").strip()
        if not material_id and str(target.get("entidade_tipo", "") or "").strip() == "Material":
            material_id = str(target.get("entidade_id", "") or "").strip()
        if not material_id:
            raise ValueError("Esta NC nao esta ligada a um material.")
        material = self.material_by_id(material_id)
        if material is None:
            raise ValueError("Material ligado a NC nao encontrado.")
        before_material = copy.deepcopy(material)
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        self._quality_quarantine_pending_stock(material, kind="Material")
        pending_qty = self._parse_float(material.get("quality_pending_qty", 0), 0)
        before_qty = self._parse_float(material.get("quantidade", 0), 0)
        if pending_qty > 0:
            material["quantidade"] = before_qty + pending_qty
            material["quality_pending_qty"] = 0.0
            material["quality_approved_qty"] = self._parse_float(material.get("quality_approved_qty", 0), 0) + pending_qty
            self.desktop_main.log_stock(
                data,
                "ENTRADA_QUALIDADE",
                f"{material_id} qtd={pending_qty} NC={nc_id_txt}",
                operador=self._current_user_label(),
            )
        material["quality_status"] = "APROVADO"
        material["inspection_status"] = "APROVADO"
        material["quality_blocked"] = False
        material["inspection_decision"] = str(decision or "Aprovado pela qualidade").strip()
        material["quality_nc_id"] = ""
        material["supplier_claim_id"] = ""
        material["quality_released_at"] = now
        material["quality_released_by"] = self._current_user_label()
        material["atualizado_em"] = now
        target["decisao"] = str(decision or "Aprovado pela qualidade").strip()
        target["acao"] = (str(target.get("acao", "") or "").strip() + f"\nLibertacao de material: {material['inspection_decision']}").strip()
        target["estado"] = "Fechada"
        target["closed_at"] = now
        target["closed_by"] = self._current_user_label()
        target["updated_at"] = now
        target["updated_by"] = self._current_user_label()
        self._append_audit_event(
            data,
            action="Material libertado pela qualidade",
            entity_type="Material",
            entity_id=material_id,
            summary=f"NC {nc_id_txt}: {material['inspection_decision']}",
            before=before_material,
            after=material,
        )
        self._sync_ne_from_materia()
        self._save(force=True, audit=False)
        return {"material_id": material_id, "quality_status": "APROVADO", "nc_id": nc_id_txt}

    def quality_nc_remove(self, nc_id: str) -> None:
        data = self.ensure_data()
        value = str(nc_id or "").strip()
        rows = list(data.get("quality_nonconformities", []) or [])
        before = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        data["quality_nonconformities"] = [row for row in rows if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        if before is None:
            raise ValueError("Nao conformidade nao encontrada.")
        self._append_audit_event(data, action="NC removida", entity_type="Nao conformidade", entity_id=value, summary=str(before.get("descricao", "") or ""), before=before)
        self._save(force=True, audit=False)

    def quality_document_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("quality_documents", []) or []):
            if not isinstance(raw, dict):
                continue
            row = {
                "id": str(raw.get("id", "") or "").strip(),
                "titulo": str(raw.get("titulo", "") or "").strip(),
                "tipo": str(raw.get("tipo", "") or "").strip(),
                "entidade": str(raw.get("entidade", "") or "").strip(),
                "referencia": str(raw.get("referencia", "") or "").strip(),
                "versao": str(raw.get("versao", "") or "").strip(),
                "estado": str(raw.get("estado", "") or "Ativo").strip(),
                "responsavel": str(raw.get("responsavel", "") or "").strip(),
                "caminho": str(raw.get("caminho", "") or "").strip(),
                "obs": str(raw.get("obs", "") or "").strip(),
                "updated_at": str(raw.get("updated_at", "") or raw.get("created_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (str(item.get("tipo", "")), str(item.get("titulo", ""))))
        return rows

    def quality_document_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("quality_documents", [])
        doc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == doc_id), None) if doc_id else None
        before = copy.deepcopy(existing) if isinstance(existing, dict) else None
        if not doc_id:
            doc_id = self._next_prefixed_id(rows, "DOC")
        titulo = str(payload.get("titulo", "") or "").strip()
        if not titulo:
            raise ValueError("Titulo do documento obrigatorio.")
        source_path = str(payload.get("caminho", "") or "").strip()
        stored_path = source_path
        if source_path:
            stored_path = self._store_shared_file(source_path, "quality/documents", preferred_name=self._file_reference_name(source_path, titulo or doc_id))
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        row = {
            "id": doc_id,
            "titulo": titulo,
            "tipo": str(payload.get("tipo", "") or "Evidencia").strip() or "Evidencia",
            "entidade": str(payload.get("entidade", "") or "").strip(),
            "referencia": str(payload.get("referencia", "") or "").strip(),
            "entidade_tipo": str(payload.get("entidade_tipo", payload.get("entidade", "")) or "").strip(),
            "entidade_id": str(payload.get("entidade_id", payload.get("referencia", "")) or "").strip(),
            "versao": str(payload.get("versao", "") or "1").strip() or "1",
            "estado": str(payload.get("estado", "") or "Ativo").strip() or "Ativo",
            "responsavel": str(payload.get("responsavel", "") or "").strip(),
            "caminho": stored_path,
            "obs": str(payload.get("obs", "") or "").strip(),
            "created_at": str((existing or {}).get("created_at", "") or now),
            "updated_at": now,
            "created_by": str((existing or {}).get("created_by", "") or self._current_user_label()),
            "updated_by": self._current_user_label(),
        }
        if existing is None:
            rows.append(row)
        else:
            existing.update(row)
            row = existing
        self._append_audit_event(data, action="Documento qualidade guardado", entity_type="Documento", entity_id=doc_id, summary=titulo, before=before, after=row)
        self._save(force=True, audit=False)
        return dict(row)

    def quality_document_remove(self, doc_id: str) -> None:
        data = self.ensure_data()
        value = str(doc_id or "").strip()
        rows = list(data.get("quality_documents", []) or [])
        before = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        data["quality_documents"] = [row for row in rows if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        if before is None:
            raise ValueError("Documento nao encontrado.")
        self._append_audit_event(data, action="Documento qualidade removido", entity_type="Documento", entity_id=value, summary=str(before.get("titulo", "") or ""), before=before)
        self._save(force=True, audit=False)

    def _quality_pdf_path(self, name: str) -> Path:
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name or "qualidade").strip()).strip("_") or "qualidade"
        return Path(tempfile.gettempdir()) / f"lugest_{safe}.pdf"

    def _quality_pdf_draw_lines(self, canvas_obj: Any, lines: list[str], x: float, y: float, width: float, *, size: float = 9.0) -> float:
        for raw in lines:
            for line in _pdf_wrap_text(raw, "Helvetica", size, width, max_lines=None) or [""]:
                canvas_obj.drawString(x, y, line)
                y -= size + 3
        return y

    def _quality_pdf_branding_assets(self) -> tuple[dict[str, Any], Path | None, list[str], str]:
        branding = self.branding_settings()
        palette = self._operator_label_palette()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        footer_lines = [str(line or "").strip() for line in list(branding.get("empresa_info_rodape", []) or []) if str(line or "").strip()]
        company_name = str(dict(branding.get("guia_emitente", {}) or {}).get("nome", "") or "").strip()
        if not company_name and footer_lines:
            company_name = footer_lines[0]
        if not company_name:
            company_name = "luGEST"
        return palette, logo_path, footer_lines[:3], company_name

    def _quality_pdf_draw_page_frame(
        self,
        canvas_obj: Any,
        page_w: float,
        page_h: float,
        *,
        title: str,
        subtitle: str = "",
        printed_at: str = "",
        page_label: str = "",
    ) -> dict[str, Any]:
        from reportlab.lib import colors
        from reportlab.lib.units import mm

        palette, logo_path, footer_lines, company_name = self._quality_pdf_branding_assets()
        outer_margin = 10 * mm
        header_h = 24 * mm
        footer_h = 18 * mm
        content_left = outer_margin + (8 * mm)
        content_right = page_w - outer_margin - (8 * mm)
        content_top = page_h - outer_margin - header_h - (6 * mm)
        content_bottom = outer_margin + footer_h + (6 * mm)

        canvas_obj.setFillColor(colors.white)
        canvas_obj.rect(0, 0, page_w, page_h, stroke=0, fill=1)
        canvas_obj.setFillColor(colors.white)
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_margin, outer_margin, page_w - (outer_margin * 2), page_h - (outer_margin * 2), 10, stroke=1, fill=1)

        header_y = page_h - outer_margin - header_h
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(outer_margin, header_y, page_w - (outer_margin * 2), header_h, 10, stroke=0, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            page_w - outer_margin - (34 * mm),
            header_y + (4.5 * mm),
            28 * mm,
            12 * mm,
            radius=5,
            padding_x=3,
            padding_y=2,
            line_width=0.7,
        )
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            page_w - outer_margin - (20 * mm),
            outer_margin + 3.5 * mm,
            14 * mm,
            8 * mm,
            radius=3,
            padding_x=2,
            padding_y=1.5,
            line_width=0.6,
        )
        canvas_obj.setFillColor(colors.white)
        canvas_obj.setFont("Helvetica-Bold", 16)
        canvas_obj.drawString(content_left, header_y + (14.5 * mm), self._operator_pdf_text(title))
        if subtitle:
            canvas_obj.setFont("Helvetica", 8.6)
            canvas_obj.drawString(content_left, header_y + (7.8 * mm), self._operator_pdf_text(_pdf_clip_text(subtitle, 120 * mm, "Helvetica", 8.6)))

        footer_y = outer_margin + 3.5 * mm
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.7)
        canvas_obj.line(content_left, outer_margin + footer_h, content_right, outer_margin + footer_h)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 7.2)
        footer_base = footer_y + 8.0
        footer_texts = footer_lines or [company_name]
        for index, line in enumerate(footer_texts[:2]):
            canvas_obj.drawString(content_left, footer_base + (index * 8.0), self._operator_pdf_text(_pdf_clip_text(line, 110 * mm, "Helvetica", 7.2)))
        if printed_at:
            canvas_obj.drawRightString(content_right, footer_base + 8.0, self._operator_pdf_text(printed_at))
        if page_label:
            canvas_obj.drawRightString(content_right, footer_base, self._operator_pdf_text(page_label))

        return {
            "palette": palette,
            "logo_path": logo_path,
            "content_left": content_left,
            "content_right": content_right,
            "content_top": content_top,
            "content_bottom": content_bottom,
        }

    def _quality_supplier_label_status_text(self, row: dict[str, Any], status_override: str = "") -> str:
        override = str(status_override or "").strip()
        if override:
            return override
        decision = str(row.get("decisao", "") or "").strip()
        decision_norm = decision.casefold()
        if "devolver" in decision_norm:
            return "DEVOLVER AO FORNECEDOR"
        if "aguardar" in decision_norm and "fornecedor" in decision_norm:
            return "AGUARDAR DECISAO DO FORNECEDOR"
        if "repor" in decision_norm or "substitu" in decision_norm:
            return "AGUARDAR REPOSICAO DO FORNECEDOR"
        if self._parse_float(row.get("qtd_rejeitada", 0), 0) > 0:
            return "REJEITADO"
        return decision.upper() or "EM ANALISE"

    def _quality_simple_pdf(self, target: Path, title: str, sections: list[tuple[str, list[str]]]) -> Path:
        def clean(value: Any) -> str:
            return str(value or "").replace("\r", " ").replace("\n", " ").strip()

        def esc(value: Any) -> str:
            text = clean(value)
            text = text.encode("latin-1", errors="replace").decode("latin-1")
            return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

        all_lines: list[str] = [clean(title), ""]
        for section_title, lines in sections:
            all_lines.append(clean(section_title))
            all_lines.extend(clean(line) for line in list(lines or []))
            all_lines.append("")
        wrapped: list[str] = []
        for line in all_lines:
            if not line:
                wrapped.append("")
                continue
            text = line
            while len(text) > 96:
                wrapped.append(text[:96])
                text = text[96:]
            wrapped.append(text)
        page_lines: list[list[str]] = []
        current: list[str] = []
        for line in wrapped:
            current.append(line)
            if len(current) >= 48:
                page_lines.append(current)
                current = []
        if current or not page_lines:
            page_lines.append(current)

        objects: list[bytes] = []
        pages_obj_num = 2
        font_obj_num = 3
        page_obj_nums: list[int] = []
        content_obj_nums: list[int] = []
        next_obj = 4
        for _page in page_lines:
            page_obj_nums.append(next_obj)
            content_obj_nums.append(next_obj + 1)
            next_obj += 2
        objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
        kids = " ".join(f"{num} 0 R" for num in page_obj_nums)
        objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {len(page_obj_nums)} >>".encode("ascii"))
        objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        for page_num, content_num, lines in zip(page_obj_nums, content_obj_nums, page_lines):
            objects.append(f"<< /Type /Page /Parent {pages_obj_num} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 {font_obj_num} 0 R >> >> /Contents {content_num} 0 R >>".encode("ascii"))
            stream_lines = ["BT", "/F1 10 Tf", "50 800 Td", "14 TL"]
            for line in lines:
                stream_lines.append(f"({esc(line)}) Tj")
                stream_lines.append("T*")
            stream_lines.append("ET")
            stream = "\n".join(stream_lines).encode("latin-1", errors="replace")
            objects.append(b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream")

        payload = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for index, obj in enumerate(objects, start=1):
            offsets.append(len(payload))
            payload.extend(f"{index} 0 obj\n".encode("ascii"))
            payload.extend(obj)
            payload.extend(b"\nendobj\n")
        xref_at = len(payload)
        payload.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
        for offset in offsets[1:]:
            payload.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        payload.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_at}\n%%EOF\n".encode("ascii"))
        target.write_bytes(bytes(payload))
        return target

    def quality_nc_pdf(self, nc_id: str) -> Path:
        nc_id_txt = str(nc_id or "").strip()
        row = next((item for item in self.quality_nc_rows("", "Todos") if str(item.get("id", "") or "").strip() == nc_id_txt), None)
        if not row:
            raise ValueError("Nao conformidade nao encontrada.")
        target = self._quality_pdf_path(f"NC_{nc_id_txt}")
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except Exception:
            path = self._quality_simple_pdf(
                target,
                f"Nao conformidade {nc_id_txt}",
                [
                    ("Identificacao", [f"{key}: {row.get(key, '')}" for key in ("estado", "gravidade", "tipo", "origem", "referencia", "entidade_label", "fornecedor_nome", "lote_fornecedor", "ne_numero", "responsavel", "prazo")]),
                    ("Descricao", [row.get("descricao", "") or "-"]),
                    ("Causa", [row.get("causa", "") or "-"]),
                    ("Acao corretiva", [row.get("acao", "") or "-"]),
                    ("Eficacia", [row.get("eficacia", "") or "-"]),
                ],
            )
            self._append_audit_event(self.ensure_data(), action="PDF NC gerado", entity_type="Nao conformidade", entity_id=nc_id_txt, summary=str(path))
            self._save(force=True, audit=False)
            return path
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
        page_w, page_h = A4
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        page_number = 1

        def begin_page() -> tuple[float, float, float, float]:
            frame = self._quality_pdf_draw_page_frame(
                canvas_obj,
                page_w,
                page_h,
                title="Nao conformidade",
                subtitle=f"{nc_id_txt} | {str(row.get('estado', '') or '-').strip()} | {str(row.get('fornecedor_nome', '') or row.get('entidade_label', '') or '-').strip()}",
                printed_at=printed_at,
                page_label=f"Pagina {page_number}",
            )
            return (
                float(frame["content_top"]),
                float(frame["content_left"]),
                float(frame["content_right"]),
                float(frame["content_bottom"]),
            )

        y, content_left, content_right, bottom_limit = begin_page()
        field_gap = 4.0

        def next_page() -> None:
            nonlocal page_number, y, content_left, content_right, bottom_limit
            canvas_obj.showPage()
            page_number += 1
            y, content_left, content_right, bottom_limit = begin_page()

        def ensure_space(required_height: float) -> None:
            if y - required_height >= bottom_limit:
                return
            next_page()

        def draw_field(label: str, value: Any) -> None:
            nonlocal y
            text = str(value or "-").strip() or "-"
            label_w = 32 * mm
            value_w = max(60.0, content_right - content_left - label_w - (6 * mm))
            lines = _pdf_wrap_text(text, "Helvetica", 9, value_w - 10, max_lines=None) or ["-"]
            box_h = max(12 * mm, 8 + (len(lines) * 12))
            ensure_space(box_h + field_gap)
            canvas_obj.setFillColor(colors.HexColor("#334155"))
            canvas_obj.setFont("Helvetica-Bold", 8)
            canvas_obj.drawString(content_left, y - 10, self._operator_pdf_text(label))
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas_obj.roundRect(content_left + label_w, y - box_h, value_w, box_h, 4, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.black)
            canvas_obj.setFont("Helvetica", 9)
            line_y = y - 14
            for line in lines:
                canvas_obj.drawString(content_left + label_w + 6, line_y, self._operator_pdf_text(line))
                line_y -= 12
            y -= box_h + field_gap

        def draw_section(label: str, value: Any) -> None:
            nonlocal y
            raw = str(value or "-").strip() or "-"
            lines = _pdf_wrap_text(raw, "Helvetica", 9, content_right - content_left - 12, max_lines=None) or ["-"]
            box_h = 12 + (len(lines) * 12)
            ensure_space(box_h + 16)
            canvas_obj.setFillColor(colors.HexColor("#0F172A"))
            canvas_obj.setFont("Helvetica-Bold", 10.5)
            canvas_obj.drawString(content_left, y - 6, self._operator_pdf_text(label))
            canvas_obj.setFillColor(colors.HexColor("#F8FAFC"))
            canvas_obj.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas_obj.roundRect(content_left, y - box_h - 14, content_right - content_left, box_h + 6, 5, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.black)
            canvas_obj.setFont("Helvetica", 9)
            line_y = y - 20
            for line in lines:
                canvas_obj.drawString(content_left + 6, line_y, self._operator_pdf_text(line))
                line_y -= 12
            y -= box_h + 20

        for label, key in (
            ("ID", "id"),
            ("Estado", "estado"),
            ("Gravidade", "gravidade"),
            ("Tipo", "tipo"),
            ("Origem", "origem"),
            ("Referencia", "referencia"),
            ("Entidade", "entidade_label"),
            ("Fornecedor", "fornecedor_nome"),
            ("Lote", "lote_fornecedor"),
            ("NE", "ne_numero"),
            ("Guia", "guia"),
            ("Decisao", "decisao"),
            ("Qtd recebida", "qtd_recebida"),
            ("Qtd aprovada", "qtd_aprovada"),
            ("Qtd rejeitada", "qtd_rejeitada"),
            ("Qtd pendente", "qtd_pendente"),
            ("Responsavel", "responsavel"),
            ("Prazo", "prazo"),
            ("Criada em", "created_at"),
            ("Fechada em", "closed_at"),
        ):
            draw_field(label, row.get(key, ""))
        for label, key in (("Descricao", "descricao"), ("Causa", "causa"), ("Acao corretiva", "acao"), ("Eficacia", "eficacia")):
            draw_section(label, row.get(key, ""))
        canvas_obj.save()
        self._append_audit_event(self.ensure_data(), action="PDF NC gerado", entity_type="Nao conformidade", entity_id=nc_id_txt, summary=str(target))
        self._save(force=True, audit=False)
        return target

    def quality_supplier_label_pdf(self, nc_id: str, output_path: str | Path | None = None, status_override: str = "") -> Path:
        nc_id_txt = str(nc_id or "").strip()
        row = next((item for item in self.quality_nc_rows("", "Todos") if str(item.get("id", "") or "").strip() == nc_id_txt), None)
        if not row:
            raise ValueError("Nao conformidade nao encontrada.")
        try:
            from reportlab.lib.pagesizes import A5
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except Exception:
            target = Path(output_path) if output_path else self._quality_pdf_path(f"etiqueta_fornecedor_{nc_id_txt}")
            return self._quality_simple_pdf(
                target,
                f"Etiqueta fornecedor {nc_id_txt}",
                [
                    (
                        "Resumo",
                        [
                            f"Estado etiqueta: {self._quality_supplier_label_status_text(row, status_override)}",
                            f"Fornecedor: {row.get('fornecedor_nome', '') or '-'}",
                            f"Referencia: {row.get('referencia', '') or '-'}",
                            f"Lote: {row.get('lote_fornecedor', '') or '-'}",
                            f"Qtd rejeitada: {row.get('qtd_rejeitada', '') or '-'}",
                            f"Decisao: {row.get('decisao', '') or '-'}",
                        ],
                    )
                ],
            )

        target = (
            Path(output_path)
            if output_path
            else self._storage_output_path("quality/labels", f"Etiqueta_Qualidade_{nc_id_txt}.pdf")
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A5
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(page_w, page_h))
        palette, logo_path, _footer_lines, _company_name = self._quality_pdf_branding_assets()
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        status_text = self._quality_supplier_label_status_text(row, status_override)
        status_norm = status_text.casefold()
        status_color = palette["warning"]
        if "rejeitado" in status_norm or "devolver" in status_norm:
            status_color = palette["danger"]
        elif "repos" in status_norm or "substit" in status_norm:
            status_color = palette["primary_dark"]

        trim_margin = 6 * mm
        safe_margin = 11 * mm
        outer_x = safe_margin
        outer_y = safe_margin
        outer_w = page_w - (safe_margin * 2)
        outer_h = page_h - (safe_margin * 2)
        header_h = 31 * mm
        body_x = outer_x + 10
        body_w = outer_w - 20
        status_box_w = 46 * mm
        status_box_h = 22 * mm

        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.8)
        canvas_obj.roundRect(trim_margin, trim_margin, page_w - (trim_margin * 2), page_h - (trim_margin * 2), 7, stroke=1, fill=0)
        canvas_obj.setDash(3, 2)
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(outer_x - 4, outer_y - 4, outer_w + 8, outer_h + 8, 9, stroke=1, fill=0)
        canvas_obj.setDash()

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 11, stroke=1, fill=1)

        header_y = outer_y + outer_h - header_h
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(outer_x + 8, header_y + 6, outer_w - 16, header_h - 12, 9, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            outer_x + 9,
            outer_y + outer_h - (18 * mm),
            24 * mm,
            10 * mm,
            radius=5,
            padding_x=3,
            padding_y=2,
            line_width=0.7,
        )
        title_x = outer_x + 36 * mm
        title_w = outer_w - (title_x - outer_x) - status_box_w - 16
        canvas_obj.setFillColor(palette["ink"])
        title_font = _pdf_fit_font_size("Etiqueta fornecedor", "Helvetica-Bold", title_w, 14.2, 11.4)
        canvas_obj.setFont("Helvetica-Bold", title_font)
        canvas_obj.drawString(title_x, outer_y + outer_h - (10.1 * mm), self._operator_pdf_text("Etiqueta fornecedor"))
        subtitle = f"NC {nc_id_txt} | Segregacao / devolucao | formato A5"
        subtitle_font = _pdf_fit_font_size(subtitle, "Helvetica", title_w, 7.6, 6.1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", subtitle_font)
        canvas_obj.drawString(title_x, outer_y + outer_h - (14.4 * mm), self._operator_pdf_text(_pdf_clip_text(subtitle, title_w, "Helvetica", subtitle_font)))

        status_box_x = outer_x + outer_w - status_box_w - 10
        status_box_y = outer_y + outer_h - header_h + 1
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(status_color)
        canvas_obj.setLineWidth(1.1)
        canvas_obj.roundRect(status_box_x, status_box_y, status_box_w, status_box_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.1)
        canvas_obj.drawString(status_box_x + 7, status_box_y + status_box_h - 8, self._operator_pdf_text("Estado"))
        canvas_obj.setFillColor(status_color)
        status_lines = _pdf_wrap_text(status_text, "Helvetica-Bold", 9.0, status_box_w - 14, max_lines=2) or [status_text]
        status_y = status_box_y + status_box_h - 16
        status_font = _pdf_fit_font_size(max(status_lines, key=len), "Helvetica-Bold", status_box_w - 14, 8.8, 6.8)
        canvas_obj.setFont("Helvetica-Bold", status_font)
        for line in status_lines[:2]:
            canvas_obj.drawCentredString(status_box_x + (status_box_w / 2.0), status_y, self._operator_pdf_text(line))
            status_y -= status_font + 1.0

        body_top = outer_y + outer_h - header_h - 10
        left_w = body_w - status_box_w - 8
        row_y = body_top - 18
        small_gap = 6
        small_w = (left_w - small_gap) / 2.0
        info_cards = [
            ("Fornecedor", str(row.get("fornecedor_nome", "") or row.get("entidade_label", "") or "-").strip() or "-"),
            ("Lote", str(row.get("lote_fornecedor", "") or "-").strip() or "-"),
            ("Referencia", str(row.get("referencia", "") or "-").strip() or "-"),
            ("Qtd rejeitada", self._fmt(row.get("qtd_rejeitada", 0))),
        ]
        for index, (label, value) in enumerate(info_cards):
            col = index % 2
            line = index // 2
            box_x = body_x + (col * (small_w + small_gap))
            box_y = row_y - (line * 27)
            canvas_obj.setFillColor(palette["surface_alt"] if line == 0 else palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, box_y, small_w, 22, 7, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont("Helvetica", 6.0)
            canvas_obj.drawString(box_x + 7, box_y + 12.6, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, "Helvetica-Bold", small_w - 14, 8.0, 6.2)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont("Helvetica-Bold", value_font)
            canvas_obj.drawString(box_x + 7, box_y + 4.2, self._operator_pdf_text(_pdf_clip_text(value, small_w - 14, "Helvetica-Bold", value_font)))

        decision_y = row_y - 60
        decision_text = str(row.get("decisao", "") or "-").strip() or "-"
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, decision_y, body_w, 22, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.1)
        canvas_obj.drawString(body_x + 8, decision_y + 13.0, self._operator_pdf_text("Instrucao / decisao"))
        decision_font = _pdf_fit_font_size(decision_text, "Helvetica-Bold", body_w - 16, 8.4, 6.5)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont("Helvetica-Bold", decision_font)
        canvas_obj.drawString(body_x + 8, decision_y + 4.2, self._operator_pdf_text(_pdf_clip_text(decision_text, body_w - 16, "Helvetica-Bold", decision_font)))

        barcode_y = outer_y + 10
        barcode_h = max(34.0, decision_y - barcode_y - 8)
        barcode_value = nc_id_txt
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, barcode_y, body_w, barcode_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.0)
        canvas_obj.drawString(body_x + 8, barcode_y + barcode_h - 11.0, self._operator_pdf_text("Codigo para rastreabilidade"))
        barcode_area_x = body_x + 10
        barcode_area_w = body_w - 20
        barcode_bar_h = max(17.0, min(24.0, barcode_h - 21.0))
        self._draw_code128_fit(canvas_obj, barcode_value, barcode_area_x, barcode_y + 11.0, barcode_area_w, barcode_bar_h, min_bar_width=0.42, max_bar_width=1.0)
        canvas_obj.setFillColor(palette["ink"])
        barcode_font = _pdf_fit_font_size(barcode_value, "Helvetica-Bold", barcode_area_w, 8.0, 6.0)
        canvas_obj.setFont("Helvetica-Bold", barcode_font)
        canvas_obj.drawCentredString(body_x + (body_w / 2.0), barcode_y + 4.0, self._operator_pdf_text(_pdf_clip_text(barcode_value, barcode_area_w, "Helvetica-Bold", barcode_font)))
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 5.2)
        canvas_obj.drawRightString(outer_x + outer_w - 8, outer_y + 4.8, self._operator_pdf_text(printed_at[:16]))
        canvas_obj.save()
        self._append_audit_event(self.ensure_data(), action="Etiqueta fornecedor gerada", entity_type="Nao conformidade", entity_id=nc_id_txt, summary=str(target))
        self._save(force=True, audit=False)
        return target

    def quality_dossier_pdf(self) -> Path:
        target = self._quality_pdf_path("dossier_qualidade")
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except Exception:
            summary = self.quality_summary()
            path = self._quality_simple_pdf(
                target,
                "Dossier Qualidade ISO 9001",
                [
                    ("Resumo", [f"{key}: {value}" for key, value in summary.items()]),
                    ("Checklist", [f"{row.get('area', '')}: {row.get('estado', '')} - {row.get('evidencia', '')}" for row in self.quality_iso_checklist()]),
                    ("Nao conformidades", [f"{row.get('id','')} | {row.get('estado','')} | {row.get('entidade_label') or row.get('referencia','')} | {row.get('descricao','')}" for row in self.quality_nc_rows("", "Todos")[:120]]),
                    ("Documentos", [f"{row.get('id','')} | {row.get('tipo','')} | {row.get('titulo','')}" for row in self.quality_document_rows("")[:120]]),
                    ("Auditoria", [f"{row.get('created_at','')} | {row.get('action','')} | {row.get('summary','')}" for row in self.audit_rows("", limit=120)]),
                ],
            )
            self._append_audit_event(self.ensure_data(), action="Dossier qualidade gerado", entity_type="Qualidade", entity_id="ISO9001", summary=str(path))
            self._save(force=True, audit=False)
            return path
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
        page_w, page_h = A4
        margin = 14 * mm

        def header(title: str) -> float:
            canvas_obj.setFillColor(colors.HexColor("#000040"))
            canvas_obj.rect(0, page_h - 22 * mm, page_w, 22 * mm, stroke=0, fill=1)
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setFont("Helvetica-Bold", 15)
            canvas_obj.drawString(margin, page_h - 14 * mm, title)
            canvas_obj.setFont("Helvetica", 8)
            canvas_obj.drawRightString(page_w - margin, page_h - 14 * mm, str(self.desktop_main.now_iso() or "").replace("T", " ")[:19])
            canvas_obj.setFillColor(colors.black)
            return page_h - 32 * mm

        y = header("Dossier Qualidade ISO 9001")
        summary = self.quality_summary()
        canvas_obj.setFont("Helvetica-Bold", 11)
        canvas_obj.drawString(margin, y, "Resumo")
        y -= 15
        canvas_obj.setFont("Helvetica", 9)
        for label, key in (("NC abertas", "open_nc"), ("NC fora prazo", "overdue_nc"), ("Documentos", "documents"), ("Eventos auditoria", "audit_events")):
            canvas_obj.drawString(margin, y, f"{label}: {summary.get(key, 0)}")
            y -= 12
        y -= 8
        canvas_obj.setFont("Helvetica-Bold", 11)
        canvas_obj.drawString(margin, y, "Checklist")
        y -= 14
        canvas_obj.setFont("Helvetica", 8.5)
        for row in self.quality_iso_checklist():
            text = f"{row.get('area', '')}: {row.get('estado', '')} - {row.get('evidencia', '')}"
            y = self._quality_pdf_draw_lines(canvas_obj, [text], margin, y, page_w - margin * 2, size=8.5)
            y -= 2
        canvas_obj.showPage()

        y = header("Nao conformidades")
        canvas_obj.setFont("Helvetica", 8)
        for row in self.quality_nc_rows("", "Todos")[:80]:
            text = f"{row.get('id','')} | {row.get('estado','')} | {row.get('gravidade','')} | {row.get('entidade_label') or row.get('referencia','')} | {row.get('descricao','')}"
            y = self._quality_pdf_draw_lines(canvas_obj, [text], margin, y, page_w - margin * 2, size=8)
            y -= 3
            if y < 24 * mm:
                canvas_obj.showPage()
                y = header("Nao conformidades")
                canvas_obj.setFont("Helvetica", 8)
        canvas_obj.showPage()

        y = header("Documentos e auditoria")
        canvas_obj.setFont("Helvetica-Bold", 10)
        canvas_obj.drawString(margin, y, "Documentos")
        y -= 14
        canvas_obj.setFont("Helvetica", 8)
        for row in self.quality_document_rows("")[:80]:
            text = f"{row.get('id','')} | {row.get('tipo','')} | v{row.get('versao','')} | {row.get('titulo','')} | {row.get('entidade') or row.get('entidade_tipo','')} {row.get('referencia') or row.get('entidade_id','')}"
            y = self._quality_pdf_draw_lines(canvas_obj, [text], margin, y, page_w - margin * 2, size=8)
            y -= 2
            if y < 45 * mm:
                canvas_obj.showPage()
                y = header("Documentos")
                canvas_obj.setFont("Helvetica", 8)
        y -= 8
        canvas_obj.setFont("Helvetica-Bold", 10)
        canvas_obj.drawString(margin, y, "Ultimos eventos de auditoria")
        y -= 14
        canvas_obj.setFont("Helvetica", 8)
        for row in self.audit_rows("", limit=100):
            text = f"{row.get('created_at','')} | {row.get('user','')} | {row.get('action','')} | {row.get('entity_type','')} {row.get('entity_id','')} | {row.get('summary','')}"
            y = self._quality_pdf_draw_lines(canvas_obj, [text], margin, y, page_w - margin * 2, size=8)
            y -= 2
            if y < 24 * mm:
                canvas_obj.showPage()
                y = header("Auditoria")
                canvas_obj.setFont("Helvetica", 8)
        canvas_obj.save()
        self._append_audit_event(self.ensure_data(), action="Dossier qualidade gerado", entity_type="Qualidade", entity_id="ISO9001", summary=str(target))
        self._save(force=True, audit=False)
        return target

    def quality_iso_checklist(self) -> list[dict[str, str]]:
        summary = self.quality_summary()
        docs = self.quality_document_rows()
        audit_count = int(summary.get("audit_events", 0) or 0)
        open_nc = int(summary.get("open_nc", 0) or 0)
        overdue_nc = int(summary.get("overdue_nc", 0) or 0)
        has_docs = bool(docs)
        blocked_materials = int(summary.get("blocked_materials", 0) or 0)
        supplier_nc = int(summary.get("supplier_nc", 0) or 0)
        return [
            {"area": "Rastreabilidade", "estado": "OK" if audit_count > 0 else "Pendente", "evidencia": f"{audit_count} eventos de auditoria registados."},
            {"area": "Nao conformidades", "estado": "Atencao" if overdue_nc else "OK", "evidencia": f"{open_nc} abertas; {overdue_nc} fora de prazo."},
            {"area": "Rececao e stock", "estado": "Atencao" if blocked_materials else "OK", "evidencia": f"{blocked_materials} lotes bloqueados/em inspecao com rastreabilidade."},
            {"area": "Reclamacoes a fornecedores", "estado": "OK" if supplier_nc or not blocked_materials else "Pendente", "evidencia": f"{supplier_nc} NC/reclamacoes de fornecedor registadas."},
            {"area": "Informacao documentada", "estado": "OK" if has_docs else "Pendente", "evidencia": f"{len(docs)} documentos/evidencias ligados ao sistema."},
            {"area": "Integridade das ligacoes", "estado": "OK" if int(summary.get("quality_issues", 0) or 0) == 0 else "Atencao", "evidencia": f"{int(summary.get('quality_issues', 0) or 0)} problemas em NC/documentos ligados."},
            {"area": "Alteracoes climaticas ISO 9001:2015/Amd 1:2024", "estado": "Pendente", "evidencia": "Registar no contexto da organizacao se o tema e relevante e que requisitos de partes interessadas existem."},
        ]

    def available_roles(self) -> list[str]:
        return ["Admin", "Producao", "Qualidade", "Planeamento", "Orcamentista", "Operador"]

    def quote_workcenter_options(self) -> list[str]:
        ordered: list[str] = []
        seen: set[str] = set()
        preferred = list(self.workcenter_machine_options("Corte Laser") or [])
        for value in preferred:
            key = str(value or "").strip().lower()
            if not key or key in seen or key == "geral":
                continue
            seen.add(key)
            ordered.append(str(value).strip())
        for group in self.workcenter_group_options():
            group_txt = str(group or "").strip()
            group_key = group_txt.lower()
            if group_txt and group_key not in seen and group_key != "geral":
                seen.add(group_key)
                ordered.append(group_txt)
            for machine in self.workcenter_machine_options(group_txt):
                machine_txt = str(machine or "").strip()
                machine_key = machine_txt.lower()
                if not machine_txt or machine_key in seen or machine_key == "geral":
                    continue
                seen.add(machine_key)
                ordered.append(machine_txt)
        return ordered or ["Laser", "Maquina 3030", "Maquina 5030", "Maquina 5040"]

    def _default_workcenter_catalog(self) -> list[dict[str, Any]]:
        raw_groups = [
            ("Corte Laser", ["Maquina 3030", "Maquina 5030", "Maquina 5040"]),
            ("Quinagem", []),
            ("Serralharia", []),
            ("Maquinacao", []),
            ("Roscagem", []),
            ("Lacagem", []),
            ("Montagem", []),
            ("Soldadura", []),
            ("Embalamento", []),
            ("Furo Manual", []),
            ("Expedicao", []),
        ]
        return [
            {
                "name": str(name),
                "operation": self._planning_normalize_operation(name, default=str(name)),
                "active": True,
                "machines": [{"name": str(machine), "active": True} for machine in list(machines or []) if str(machine).strip()],
            }
            for name, machines in raw_groups
        ]

    def _workcenter_group_aliases(self, name: str) -> set[str]:
        raw = str(name or "").strip()
        norm = self.desktop_main.norm_text(raw)
        aliases = {raw.lower(), norm}
        if "laser" in norm:
            aliases.update({"laser", "corte laser"})
        if "quin" in norm:
            aliases.update({"quinagem", "quin"})
        if "serralh" in norm:
            aliases.update({"serralharia", "serralh"})
        if "maquin" in norm:
            aliases.update({"maquinacao", "maquinação", "maquin"})
        if "rosc" in norm:
            aliases.update({"roscagem", "rosc"})
        if "laca" in norm or "pint" in norm:
            aliases.update({"lacagem", "pintura"})
        if "mont" in norm:
            aliases.update({"montagem", "mont"})
        if "sold" in norm:
            aliases.update({"soldadura", "sold"})
        if "embal" in norm:
            aliases.update({"embalamento", "embal"})
        if "furo" in norm:
            aliases.update({"furo manual", "furo"})
        if "exped" in norm:
            aliases.update({"expedicao", "expedição", "shipping"})
        return {alias for alias in aliases if alias}

    def _legacy_workcenter_group_name(self, value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        catalog = list(self._workcenter_catalog(sync_legacy=False) or [])
        raw_lower = raw.lower()
        raw_norm = self.desktop_main.norm_text(raw)
        for row in catalog:
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            if raw_lower == group_name.lower() or raw_norm in self._workcenter_group_aliases(group_name):
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str(machine or "").strip()
                if machine_txt and machine_txt.lower() == raw_lower:
                    return group_name
        if raw.replace(" ", "").isdigit() or raw_norm.startswith("maquina "):
            return "Corte Laser"
        alias_map = {
            "laser": "Corte Laser",
            "quin": "Quinagem",
            "serralh": "Serralharia",
            "maquin": "Maquinacao",
            "rosc": "Roscagem",
            "laca": "Lacagem",
            "pint": "Lacagem",
            "mont": "Montagem",
            "sold": "Soldadura",
            "embal": "Embalamento",
            "furo": "Furo Manual",
            "exped": "Expedicao",
        }
        for token, canonical in alias_map.items():
            if token in raw_norm:
                return canonical
        return raw

    def _workcenter_catalog(self, *, sync_legacy: bool = True) -> list[dict[str, Any]]:
        data = self.ensure_data()
        raw_catalog = list(data.get("workcenter_catalog", []) or [])
        default_catalog = list(self._default_workcenter_catalog() or [])
        groups_by_name: dict[str, dict[str, Any]] = {}

        def has_meaningful_usage() -> bool:
            for row in list(data.get("users", []) or []):
                if any(str(row.get(key, "") or "").strip() for key in ("posto", "posto_trabalho", "work_center", "workcenter")):
                    return True
            for row in list(data.get("orcamentos", []) or []):
                if str(row.get("posto_trabalho", "") or "").strip():
                    return True
            for row in list(data.get("encomendas", []) or []):
                if any(str(row.get(key, "") or "").strip() for key in ("posto_trabalho", "posto", "maquina")):
                    return True
                for mat in list(row.get("materiais", []) or []):
                    for esp in list(mat.get("espessuras", []) or []):
                        if dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {}):
                            return True
            for bucket_name in ("plano", "plano_hist"):
                for row in list(data.get(bucket_name, []) or []):
                    if any(str(row.get(key, "") or "").strip() for key in ("posto", "posto_trabalho", "maquina")):
                        return True
            return False

        suspicious_catalog = False
        raw_group_names = {str((row or {}).get("name", "") or "").strip().lower() for row in raw_catalog if isinstance(row, dict) and str((row or {}).get("name", "") or "").strip()}
        for row in raw_catalog:
            if not isinstance(row, dict):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if group_name.replace(" ", "").isdigit():
                suspicious_catalog = True
                break
            for machine in list(row.get("machines", []) or []):
                machine_txt = str(machine or "").strip()
                if machine_txt and machine_txt.lower() in raw_group_names and machine_txt.lower() != group_name.lower():
                    suspicious_catalog = True
                    break
            if suspicious_catalog:
                break
        if raw_catalog and suspicious_catalog and not has_meaningful_usage():
            raw_catalog = []
            data["workcenter_catalog"] = []
            data["postos_trabalho"] = []

        def ensure_group(name: str, operation: str = "") -> dict[str, Any]:
            group_name = str(name or "").strip()
            if not group_name:
                group_name = "Sem grupo"
            key = group_name.lower()
            row = groups_by_name.get(key)
            if row is None:
                row = {
                    "name": group_name,
                    "operation": str(operation or "").strip() or self._planning_normalize_operation(group_name, default=group_name),
                    "active": True,
                    "machines": [],
                }
                groups_by_name[key] = row
            elif operation and not str(row.get("operation", "") or "").strip():
                row["operation"] = str(operation).strip()
            return row

        def machine_name_from(raw_machine: Any) -> str:
            if isinstance(raw_machine, dict):
                return str(raw_machine.get("name", raw_machine.get("nome", "")) or "").strip()
            return str(raw_machine or "").strip()

        def machine_active_from(raw_machine: Any) -> bool:
            if isinstance(raw_machine, dict):
                return bool(raw_machine.get("active", raw_machine.get("ativo", True)))
            return True

        def add_machine(group_name: str, machine_name: Any, active: bool = True) -> None:
            machine_txt = str(machine_name or "").strip()
            if not machine_txt:
                return
            group_row = ensure_group(group_name)
            for owner in groups_by_name.values():
                if owner is group_row:
                    continue
                if any(machine_name_from(value).lower() == machine_txt.lower() for value in list(owner.get("machines", []) or [])):
                    return
            existing = {
                machine_name_from(value).lower()
                for value in list(group_row.get("machines", []) or [])
                if machine_name_from(value)
            }
            if machine_txt.lower() not in existing and machine_txt.lower() != str(group_row.get("name", "") or "").strip().lower():
                group_row.setdefault("machines", []).append({"name": machine_txt, "active": bool(active)})

        seed_catalog = default_catalog
        for default_row in seed_catalog:
            group_row = ensure_group(str(default_row.get("name", "") or "").strip(), str(default_row.get("operation", "") or "").strip())
            group_row["active"] = bool(default_row.get("active", True))
            for machine in list(default_row.get("machines", []) or []):
                add_machine(str(group_row.get("name", "") or "").strip(), machine_name_from(machine), machine_active_from(machine))

        for raw_row in raw_catalog:
            if not isinstance(raw_row, dict):
                continue
            group_name = str(raw_row.get("name", "") or "").strip()
            if not group_name:
                continue
            group_row = ensure_group(group_name, str(raw_row.get("operation", "") or "").strip())
            group_row["active"] = bool(raw_row.get("active", raw_row.get("ativo", True)))
            for machine in list(raw_row.get("machines", []) or []):
                add_machine(str(group_row.get("name", "") or "").strip(), machine_name_from(machine), machine_active_from(machine))

        if sync_legacy and not raw_catalog:
            legacy_postos = [str(value or "").strip() for value in list(data.get("postos_trabalho", []) or []) if str(value or "").strip()]
            for legacy_name in legacy_postos:
                if legacy_name.lower() == "geral":
                    continue
                group_name = self._legacy_workcenter_group_name(legacy_name)
                if not group_name:
                    continue
                normalized_group = str(group_name or "").strip()
                ensure_group(normalized_group)
                if legacy_name.lower() != normalized_group.lower():
                    add_machine(normalized_group, legacy_name)

        cleaned: list[dict[str, Any]] = []
        for row in sorted(groups_by_name.values(), key=lambda item: str(item.get("name", "") or "").lower()):
            machines_by_key: dict[str, dict[str, Any]] = {}
            for machine in list(row.get("machines", []) or []):
                machine_txt = machine_name_from(machine)
                if machine_txt:
                    machines_by_key[machine_txt.lower()] = {"name": machine_txt, "active": machine_active_from(machine)}
            machines = sorted(machines_by_key.values(), key=lambda value: str(value.get("name", "")).lower())
            cleaned.append(
                {
                    "name": str(row.get("name", "") or "").strip(),
                    "operation": str(row.get("operation", "") or "").strip()
                    or self._planning_normalize_operation(row.get("name", ""), default=str(row.get("name", "") or "").strip()),
                    "active": bool(row.get("active", True)),
                    "machines": machines,
                }
            )

        flattened: list[str] = []
        seen_flat: set[str] = set()
        for row in cleaned:
            for value in [str(row.get("name", "") or "").strip(), *[machine_name_from(machine) for machine in list(row.get("machines", []) or [])]]:
                key = str(value or "").strip().lower()
                if not key or key == "geral" or key in seen_flat:
                    continue
                seen_flat.add(key)
                flattened.append(str(value).strip())
        data["workcenter_catalog"] = cleaned
        data["postos_trabalho"] = flattened
        return cleaned

    def workcenter_group_options(self, operation: Any = "", include_general: bool = False) -> list[str]:
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        rows = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_operation and row_operation != target_operation:
                continue
            rows.append(group_name)
        rows = sorted(dict.fromkeys(rows), key=lambda value: value.lower())
        if include_general:
            return ["Geral"] + rows
        return rows

    def workcenter_machine_options(self, group: str = "", operation: Any = "") -> list[str]:
        target_group = str(group or "").strip()
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        rows: list[str] = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_group and group_name.lower() != target_group.lower():
                continue
            if target_operation and row_operation != target_operation:
                continue
            for machine in list(row.get("machines", []) or []):
                machine_name = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                machine_active = bool((machine or {}).get("active", True)) if isinstance(machine, dict) else True
                if machine_name and machine_active:
                    rows.append(machine_name)
        return sorted(dict.fromkeys(rows), key=lambda value: value.lower())

    def workcenter_resource_options(self, operation: Any = "", include_all: bool = False) -> list[str]:
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        resources: list[str] = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_operation and row_operation != target_operation:
                continue
            machines = [
                str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                for machine in list(row.get("machines", []) or [])
                if (bool((machine or {}).get("active", True)) if isinstance(machine, dict) else True)
                and (str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip())
            ]
            if machines:
                resources.extend(machines)
            else:
                resources.append(group_name)
        ordered = sorted(dict.fromkeys(resources), key=lambda value: value.lower())
        if include_all:
            return ["Todos"] + ordered
        return ordered

    def workcenter_group_for_resource(self, resource: Any, operation: Any = "") -> str:
        resource_txt = self._normalize_workcenter_value(resource)
        if not resource_txt:
            return ""
        for row in list(self._workcenter_catalog() or []):
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if str(operation or "").strip() and row_operation != self._planning_normalize_operation(operation):
                continue
            if resource_txt.lower() == group_name.lower():
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                if machine_txt and resource_txt.lower() == machine_txt.lower():
                    return group_name
        return self._legacy_workcenter_group_name(resource_txt)

    def workcenter_default_resource(self, operation: Any = "", preferred: Any = "") -> str:
        preferred_txt = self._normalize_workcenter_value(preferred)
        options = list(self.workcenter_resource_options(operation) or [])
        if preferred_txt and any(preferred_txt.lower() == str(value or "").strip().lower() for value in options):
            return preferred_txt
        return str(options[0] if options else "").strip()

    def _normalize_workcenter_value(self, value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        catalog = list(self._workcenter_catalog() or [])
        for row in catalog:
            group_name = str(row.get("name", "") or "").strip()
            if group_name and raw.lower() == group_name.lower():
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                if machine_txt and raw.lower() == machine_txt.lower():
                    return machine_txt
            if group_name and self.desktop_main.norm_text(raw) in self._workcenter_group_aliases(group_name):
                return group_name
        return raw

    def _workcenter_machine_name(self, machine: Any) -> str:
        if isinstance(machine, dict):
            return str(machine.get("name", machine.get("nome", "")) or "").strip()
        return str(machine or "").strip()

    def _workcenter_machine_active(self, machine: Any) -> bool:
        if isinstance(machine, dict):
            return bool(machine.get("active", machine.get("ativo", True)))
        return True

    def _workcenter_machine_entry(self, name: Any, active: bool = True) -> dict[str, Any]:
        return {"name": str(name or "").strip(), "active": bool(active)}

    def available_postos(self) -> list[str]:
        postos = ["Geral"] + self.quote_workcenter_options()
        seen: set[str] = set()
        ordered: list[str] = []
        for posto in postos:
            key = posto.lower()
            if key in seen:
                continue
            seen.add(key)
            ordered.append(posto)
        return ordered

    def operator_posto_options(self) -> list[str]:
        return self.workcenter_group_options(include_general=True)

    def _workcenter_usage_counts(
        self,
        workcenter: str,
        *,
        data: dict[str, Any] | None = None,
        profiles: dict[str, Any] | None = None,
    ) -> dict[str, int]:
        name = str(workcenter or "").strip()
        if not name:
            return {"users": 0, "quotes": 0, "orders": 0, "planning": 0, "operator_map": 0, "total": 0}
        data = data if isinstance(data, dict) else self.ensure_data()
        profiles = profiles if isinstance(profiles, dict) else self._user_profiles()
        norm = self._normalize_workcenter_value(name).lower()
        tracked_names = {norm}
        group_name = self._legacy_workcenter_group_name(name)
        if group_name and group_name.lower() == norm:
            tracked_names.update(str(machine or "").strip().lower() for machine in self.workcenter_machine_options(group_name))
        users = 0
        for user in list(data.get("users", []) or []):
            username = str(user.get("username", "") or "").strip().lower()
            profile = dict(profiles.get(username, {}) or {})
            posto = str(profile.get("posto", "") or user.get("posto", "") or "").strip()
            if self._normalize_workcenter_value(posto).lower() in tracked_names:
                users += 1
        quotes = sum(
            1
            for row in list(data.get("orcamentos", []) or [])
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() in tracked_names
        )
        orders = sum(
            1
            for row in list(data.get("encomendas", []) or [])
            if self._normalize_workcenter_value(
                str(row.get("posto_trabalho", row.get("posto", row.get("maquina", ""))) or "")
            ).lower()
            in tracked_names
        )
        for row in list(data.get("encomendas", []) or []):
            if not isinstance(row, dict):
                continue
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    for raw_resource in dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {}).values():
                        if self._normalize_workcenter_value(str(raw_resource or "")).lower() in tracked_names:
                            orders += 1
                            break
        planning = 0
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                posto = str(row.get("maquina", row.get("posto", row.get("posto_trabalho", ""))) or "").strip()
                if self._normalize_workcenter_value(posto).lower() in tracked_names:
                    planning += 1
        operator_map = 0
        for value in dict(data.get("operador_posto_map", {}) or {}).values():
            if self._normalize_workcenter_value(str(value or "")).lower() in tracked_names:
                operator_map += 1
        total = users + quotes + orders + planning + operator_map
        return {
            "users": users,
            "quotes": quotes,
            "orders": orders,
            "planning": planning,
            "operator_map": operator_map,
            "total": total,
        }

    def workcenter_rows(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        data = self.ensure_data()
        for row in list(self._workcenter_catalog() or []):
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            usage = self._workcenter_usage_counts(group_name, data=data)
            rows.append(
                {
                    "entry_type": "group",
                    "name": group_name,
                    "group": group_name,
                    "operation": self._planning_normalize_operation(row.get("operation", group_name), default=group_name),
                    "kind": "Posto",
                    "protected": group_name == "Geral",
                    "active": bool(row.get("active", True)),
                    "machine_count": len(list(row.get("machines", []) or [])),
                    "users": usage["users"],
                    "quotes": usage["quotes"],
                    "orders": usage["orders"],
                    "planning": usage["planning"],
                    "operator_map": usage["operator_map"],
                    "usage_total": usage["total"],
                }
            )
            for machine_name in list(row.get("machines", []) or []):
                machine_txt = str((machine_name or {}).get("name", "") or "").strip() if isinstance(machine_name, dict) else str(machine_name or "").strip()
                machine_active = bool((machine_name or {}).get("active", True)) if isinstance(machine_name, dict) else True
                if not machine_txt:
                    continue
                machine_usage = self._workcenter_usage_counts(machine_txt, data=data)
                rows.append(
                    {
                        "entry_type": "machine",
                        "name": machine_txt,
                        "group": group_name,
                        "operation": self._planning_normalize_operation(row.get("operation", group_name), default=group_name),
                        "kind": "Maquina",
                        "protected": False,
                        "active": machine_active,
                        "machine_count": 0,
                        "users": machine_usage["users"],
                        "quotes": machine_usage["quotes"],
                        "orders": machine_usage["orders"],
                        "planning": machine_usage["planning"],
                        "operator_map": machine_usage["operator_map"],
                        "usage_total": machine_usage["total"],
                    }
                )
        return rows

    def _replace_removed_resource_references(
        self,
        target_name: str,
        *,
        fallback_resource: str = "",
        operation: Any = "",
        remove_active_planning: bool = False,
    ) -> dict[str, int]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        target_txt = self._normalize_workcenter_value(target_name)
        fallback_txt = self._normalize_workcenter_value(fallback_resource)
        fallback_group = self.workcenter_group_for_resource(fallback_txt, operation) if fallback_txt else ""
        target_key = target_txt.lower()
        stats = {"profiles": 0, "users": 0, "quotes": 0, "orders": 0, "maps": 0, "planning_removed": 0, "planning_updated": 0, "history_updated": 0}
        if not target_key:
            return stats

        def match(value: Any) -> bool:
            return self._normalize_workcenter_value(value).lower() == target_key

        replacement_for_order = fallback_txt or fallback_group
        replacement_for_history = fallback_txt or fallback_group

        for username, profile in list(profiles.items()):
            if match(profile.get("posto", "")):
                profile["posto"] = replacement_for_order
                profiles[username] = profile
                stats["profiles"] += 1
        for user in list(data.get("users", []) or []):
            if match(user.get("posto", "")):
                user["posto"] = replacement_for_order
                stats["users"] += 1
        for row in list(data.get("orcamentos", []) or []):
            if match(row.get("posto_trabalho", "")):
                row["posto_trabalho"] = replacement_for_order
                stats["quotes"] += 1
        for row in list(data.get("encomendas", []) or []):
            changed_order = False
            for key in ("posto_trabalho", "posto", "maquina"):
                if match(row.get(key, "")):
                    row[key] = replacement_for_order
                    changed_order = True
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if not match(raw_value):
                            continue
                        repl = fallback_txt or self.workcenter_default_resource(op_name, preferred=fallback_group)
                        if repl:
                            machine_map[op_name] = repl
                        else:
                            machine_map.pop(op_name, None)
                        changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
                        changed_order = True
                        stats["maps"] += 1
            if changed_order:
                stats["orders"] += 1

        active_rows = []
        for row in list(data.get("plano", []) or []):
            if match(self._planning_row_resource(row)):
                if remove_active_planning:
                    stats["planning_removed"] += 1
                    continue
                self._planning_apply_resource_to_row(row, replacement_for_order, row.get("operacao", operation))
                stats["planning_updated"] += 1
            active_rows.append(row)
        data["plano"] = active_rows

        for row in list(data.get("plano_hist", []) or []):
            if not match(self._planning_row_resource(row)):
                continue
            self._planning_apply_resource_to_row(row, replacement_for_history, row.get("operacao", operation))
            stats["history_updated"] += 1

        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if match(posto):
                posto_map[username] = replacement_for_order
        if posto_map:
            data["operador_posto_map"] = posto_map
        self._save_user_profiles(profiles)
        return stats

    def save_workcenter_group(self, name: str, operation: Any = "", current_name: str = "", active: bool = True) -> dict[str, Any]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        new_name = str(name or "").strip()
        current_txt = str(current_name or "").strip()
        if not new_name:
            raise ValueError("Nome do posto obrigatório.")
        if new_name.lower() == "geral":
            raise ValueError("O posto 'Geral' já existe no sistema e não pode ser redefinido.")
        new_operation = self._planning_normalize_operation(operation or new_name, default=new_name)
        catalog = list(self._workcenter_catalog() or [])
        group_names = {str(row.get("name", "") or "").strip().lower() for row in catalog if str(row.get("name", "") or "").strip()}
        machine_names = {
            self._workcenter_machine_name(machine).lower()
            for row in catalog
            for machine in list(row.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        }
        current_key = current_txt.lower()
        if not current_txt:
            if new_name.lower() in group_names or new_name.lower() in machine_names:
                raise ValueError("Já existe um posto de trabalho com esse nome.")
            catalog.append({"name": new_name, "operation": new_operation, "active": bool(active), "machines": []})
            data["workcenter_catalog"] = catalog
            self._workcenter_catalog()
            self._save(force=True)
            return next(
                (
                    row
                    for row in self.workcenter_rows()
                    if str(row.get("entry_type", "") or "") == "group"
                    and str(row.get("name", "") or "").strip().lower() == new_name.lower()
                ),
                {"name": new_name, "entry_type": "group"},
            )
        current_group = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == current_key), None)
        if current_group is None:
            raise ValueError("Só é possível editar postos existentes.")
        if new_name.lower() != current_key and (new_name.lower() in group_names or new_name.lower() in machine_names):
            raise ValueError("Já existe um posto de trabalho com esse nome.")
        current_group["name"] = new_name
        current_group["operation"] = new_operation
        current_group["active"] = bool(active)
        for username, profile in list(profiles.items()):
            if self._normalize_workcenter_value(str(profile.get("posto", "") or "")).lower() == current_key:
                profile["posto"] = new_name
                profiles[username] = profile
        for user in list(data.get("users", []) or []):
            if self._normalize_workcenter_value(str(user.get("posto", "") or "")).lower() == current_key:
                user["posto"] = new_name
        for row in list(data.get("orcamentos", []) or []):
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() == current_key:
                row["posto_trabalho"] = new_name
        for row in list(data.get("encomendas", []) or []):
            for key in ("posto_trabalho", "posto", "maquina"):
                if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                    row[key] = new_name
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if self._normalize_workcenter_value(str(raw_value or "")).lower() == current_key:
                            machine_map[op_name] = new_name
                            changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                for key in ("posto", "posto_trabalho", "maquina"):
                    if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                        row[key] = new_name
        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if self._normalize_workcenter_value(str(posto or "")).lower() == current_key:
                posto_map[username] = new_name
        if posto_map:
            data["operador_posto_map"] = posto_map
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save_user_profiles(profiles)
        self._save(force=True)
        return next(
            (
                row
                for row in self.workcenter_rows()
                if str(row.get("entry_type", "") or "") == "group"
                and str(row.get("name", "") or "").strip().lower() == new_name.lower()
            ),
            {"name": new_name, "entry_type": "group"},
        )

    def remove_workcenter_group(self, name: str) -> None:
        data = self.ensure_data()
        target = str(name or "").strip()
        if not target:
            raise ValueError("Posto de trabalho inválido.")
        catalog = list(self._workcenter_catalog() or [])
        current_group = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == target.lower()), None)
        if current_group is None:
            raise ValueError("Posto de trabalho não encontrado.")
        if list(current_group.get("machines", []) or []):
            raise ValueError("Remove primeiro as máquinas associadas a este posto.")
        usage = self._workcenter_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover este posto porque ainda está em uso "
                f"(utilizadores: {usage['users']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        data["workcenter_catalog"] = [row for row in catalog if str(row.get("name", "") or "").strip().lower() != target.lower()]
        self._workcenter_catalog()
        self._save(force=True)

    def save_workcenter_machine(self, group_name: str, machine_name: str, current_name: str = "", active: bool = True) -> dict[str, Any]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        parent_group = str(group_name or "").strip()
        new_name = str(machine_name or "").strip()
        current_txt = str(current_name or "").strip()
        if not parent_group:
            raise ValueError("Seleciona o posto de trabalho da máquina.")
        if not new_name:
            raise ValueError("Nome da máquina obrigatório.")
        catalog = list(self._workcenter_catalog() or [])
        group_row = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == parent_group.lower()), None)
        if group_row is None:
            raise ValueError("Posto de trabalho não encontrado.")
        all_group_names = {str(row.get("name", "") or "").strip().lower() for row in catalog if str(row.get("name", "") or "").strip()}
        all_machine_names = {
            self._workcenter_machine_name(machine).lower()
            for row in catalog
            for machine in list(row.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        }
        current_key = current_txt.lower()
        if not current_txt:
            if new_name.lower() in all_group_names or new_name.lower() in all_machine_names:
                raise ValueError("Já existe um posto ou máquina com esse nome.")
            group_row.setdefault("machines", []).append(self._workcenter_machine_entry(new_name, active))
            group_row["machines"] = sorted(
                {self._workcenter_machine_name(value).lower(): self._workcenter_machine_entry(self._workcenter_machine_name(value), self._workcenter_machine_active(value)) for value in list(group_row.get("machines", []) or []) if self._workcenter_machine_name(value)}.values(),
                key=lambda value: str(value.get("name", "")).lower(),
            )
            data["workcenter_catalog"] = catalog
            self._workcenter_catalog()
            self._save(force=True)
            return next(
                (
                    row
                    for row in self.workcenter_rows()
                    if str(row.get("entry_type", "") or "") == "machine"
                    and str(row.get("name", "") or "").strip().lower() == new_name.lower()
                ),
                {"name": new_name, "entry_type": "machine", "group": parent_group},
            )
        machine_owner = next(
            (
                row
                for row in catalog
                if any(self._workcenter_machine_name(machine).lower() == current_key for machine in list(row.get("machines", []) or []))
            ),
            None,
        )
        if machine_owner is None:
            raise ValueError("Máquina não encontrada.")
        if new_name.lower() != current_key and (new_name.lower() in all_group_names or new_name.lower() in all_machine_names):
            raise ValueError("Já existe um posto ou máquina com esse nome.")
        machine_owner["machines"] = [
            self._workcenter_machine_entry(new_name, active)
            if self._workcenter_machine_name(machine).lower() == current_key
            else self._workcenter_machine_entry(self._workcenter_machine_name(machine), self._workcenter_machine_active(machine))
            for machine in list(machine_owner.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        ]
        machine_owner["machines"] = sorted(
            {self._workcenter_machine_name(value).lower(): value for value in machine_owner["machines"]}.values(),
            key=lambda value: str(value.get("name", "")).lower(),
        )
        if machine_owner is not group_row:
            machine_owner["machines"] = [value for value in list(machine_owner.get("machines", []) or []) if self._workcenter_machine_name(value).lower() != new_name.lower()]
            group_row.setdefault("machines", []).append(self._workcenter_machine_entry(new_name, True))
            group_row["machines"] = sorted(
                {self._workcenter_machine_name(value).lower(): self._workcenter_machine_entry(self._workcenter_machine_name(value), self._workcenter_machine_active(value)) for value in list(group_row.get("machines", []) or []) if self._workcenter_machine_name(value)}.values(),
                key=lambda value: str(value.get("name", "")).lower(),
            )
        for username, profile in list(profiles.items()):
            if self._normalize_workcenter_value(str(profile.get("posto", "") or "")).lower() == current_key:
                profile["posto"] = new_name
                profiles[username] = profile
        for user in list(data.get("users", []) or []):
            if self._normalize_workcenter_value(str(user.get("posto", "") or "")).lower() == current_key:
                user["posto"] = new_name
        for row in list(data.get("orcamentos", []) or []):
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() == current_key:
                row["posto_trabalho"] = new_name
        for row in list(data.get("encomendas", []) or []):
            for key in ("posto_trabalho", "posto", "maquina"):
                if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                    row[key] = new_name
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if self._normalize_workcenter_value(str(raw_value or "")).lower() == current_key:
                            machine_map[op_name] = new_name
                            changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                for key in ("posto", "posto_trabalho", "maquina"):
                    if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                        row[key] = new_name
        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if self._normalize_workcenter_value(str(posto or "")).lower() == current_key:
                posto_map[username] = new_name
        if posto_map:
            data["operador_posto_map"] = posto_map
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save_user_profiles(profiles)
        self._save(force=True)
        return next(
            (
                row
                for row in self.workcenter_rows()
                if str(row.get("entry_type", "") or "") == "machine"
                and str(row.get("name", "") or "").strip().lower() == new_name.lower()
            ),
            {"name": new_name, "entry_type": "machine", "group": parent_group},
        )

    def remove_workcenter_machine(self, machine_name: str) -> None:
        data = self.ensure_data()
        target = str(machine_name or "").strip()
        if not target:
            raise ValueError("Máquina inválida.")
        catalog = list(self._workcenter_catalog() or [])
        owner_row = next(
            (
                row
                for row in catalog
                if any(self._workcenter_machine_name(machine).lower() == target.lower() for machine in list(row.get("machines", []) or []))
            ),
            None,
        )
        if owner_row is None:
            raise ValueError("Máquina não encontrada.")
        usage = self._workcenter_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover esta máquina porque ainda está em uso "
                f"(utilizadores: {usage['users']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        owner_row["machines"] = [value for value in list(owner_row.get("machines", []) or []) if self._workcenter_machine_name(value).lower() != target.lower()]
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save(force=True)

    def _order_workcenter(self, enc_or_numero: dict[str, Any] | str | None = None) -> str:
        enc: dict[str, Any] | None
        if isinstance(enc_or_numero, dict):
            enc = enc_or_numero
        else:
            enc = self.get_encomenda_by_numero(str(enc_or_numero or "").strip()) if str(enc_or_numero or "").strip() else None
        if not isinstance(enc, dict):
            return ""
        for key in ("posto_trabalho", "posto", "maquina"):
            value = str(enc.get(key, "") or "").strip()
            if value:
                return self._normalize_workcenter_value(value)
        return ""

    def _order_esp_machine_map(self, esp_obj: dict[str, Any] | None) -> dict[str, str]:
        if not isinstance(esp_obj, dict):
            return {}
        raw_map = dict(esp_obj.get("maquinas_operacao", esp_obj.get("recursos_operacao", {})) or {})
        cleaned: dict[str, str] = {}
        for raw_op, raw_resource in raw_map.items():
            op_txt = self._planning_normalize_operation(raw_op, default="")
            resource_txt = self._sanitize_operation_resource(op_txt, raw_resource)
            if not op_txt or not resource_txt:
                continue
            cleaned[op_txt] = resource_txt
        return cleaned

    def _sanitize_operation_resource(self, operation: Any = "", resource: Any = "") -> str:
        op_txt = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        resource_txt = self._normalize_workcenter_value(resource)
        if not op_txt:
            return resource_txt
        available_resources = [
            str(value or "").strip()
            for value in list(self.workcenter_resource_options(op_txt) or [])
            if str(value or "").strip()
        ]
        if not available_resources:
            return resource_txt
        if resource_txt and any(resource_txt.lower() == value.lower() for value in available_resources):
            return next((value for value in available_resources if value.lower() == resource_txt.lower()), resource_txt)
        if len(available_resources) == 1:
            return str(available_resources[0] or "").strip()
        return ""

    def _order_operation_resource(
        self,
        enc_or_numero: dict[str, Any] | str | None,
        material: str = "",
        espessura: str = "",
        operation: Any = "",
    ) -> str:
        op_txt = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        enc = enc_or_numero if isinstance(enc_or_numero, dict) else self.get_encomenda_by_numero(str(enc_or_numero or "").strip())
        if not isinstance(enc, dict):
            return ""
        esp_obj = self._order_find_espessura(enc, material, espessura) if str(material or "").strip() and str(espessura or "").strip() else None
        machine_map = self._order_esp_machine_map(esp_obj)
        if op_txt and machine_map.get(op_txt):
            return str(machine_map.get(op_txt) or "").strip()
        if op_txt:
            return self.workcenter_default_resource(op_txt, preferred=self._order_workcenter(enc))
        return self._order_workcenter(enc)

    def user_rows(self) -> list[dict[str, Any]]:
        profiles = self._user_profiles()
        rows: list[dict[str, Any]] = []
        for user in list(self.ensure_data().get("users", []) or []):
            username = str(user.get("username", "") or "").strip()
            if not username:
                continue
            profile = dict(profiles.get(username.lower(), {}) or {})
            stored_password = str(user.get("password", "") or "").strip()
            rows.append(
                {
                    "username": username,
                    "password": "",
                    "password_set": bool(stored_password),
                    "role": str(user.get("role", "") or "").strip() or "Operador",
                    "posto": str(profile.get("posto", "") or user.get("posto", "") or "").strip(),
                    "active": bool(profile.get("active", True)),
                    "menu_permissions": dict(profile.get("menu_permissions", {}) or {}),
                }
            )
        rows.sort(key=lambda row: (str(row.get("role", "") or ""), str(row.get("username", "") or "").lower()))
        return rows

    def allowed_pages_for_user(self, user: dict[str, Any] | None = None) -> list[str]:
        current = dict(user or self.user or {})
        if str(current.get("role", "") or "").strip().lower() == "admin":
            return [row["key"] for row in self.available_menu_pages()]
        perms = dict(current.get("menu_permissions", {}) or {})
        if not perms:
            profile = self._user_profile(str(current.get("username", "") or ""))
            perms = dict(profile.get("menu_permissions", {}) or {})
        if not perms:
            return [row["key"] for row in self.available_menu_pages()]
        return [row["key"] for row in self.available_menu_pages() if bool(perms.get(row["key"], False))]

    def save_user(self, payload: dict[str, Any], current_username: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        username = str(payload.get("username", "") or "").strip()
        password = str(payload.get("password", "") or "")
        role = str(payload.get("role", "") or "Operador").strip() or "Operador"
        posto = str(payload.get("posto", "") or "").strip()
        active = bool(payload.get("active", True))
        permissions = {str(key): bool(value) for key, value in dict(payload.get("menu_permissions", {}) or {}).items()}
        if not username:
            raise ValueError("Utilizador obrigatorio.")
        current_txt = str(current_username or "").strip()
        current_key = current_txt.lower()
        logged_key = str((self.user or {}).get("username", "") or "").strip().lower()
        target = None
        for row in list(data.get("users", []) or []):
            row_username = str(row.get("username", "") or "").strip()
            if row_username.lower() == current_key and current_key:
                target = row
                break
        if target is None:
            for row in list(data.get("users", []) or []):
                if str(row.get("username", "") or "").strip().lower() == username.lower():
                    target = row
                    break
        if target is None and any(str(row.get("username", "") or "").strip().lower() == username.lower() for row in list(data.get("users", []) or [])):
            raise ValueError("Ja existe um utilizador com esse username.")
        if target is None and not password.strip():
            raise ValueError("Password obrigatoria para um novo utilizador.")
        stored_password = ""
        if password.strip():
            self.desktop_main.validate_local_password(username, password)
            stored_password = self.desktop_main.normalize_password_for_storage(username, password, require_strong=True)
        if target is not None:
            old_username = str(target.get("username", "") or "").strip()
            if old_username.lower() != username.lower() and any(str(row.get("username", "") or "").strip().lower() == username.lower() for row in list(data.get("users", []) or [])):
                raise ValueError("Ja existe um utilizador com esse username.")
            if not stored_password:
                stored_password = str(target.get("password", "") or "").strip()
            target.update({"username": username, "password": stored_password, "role": role})
        else:
            data.setdefault("users", []).append({"username": username, "password": stored_password, "role": role})
        if logged_key and logged_key in {current_key, username.lower()} and not active:
            raise ValueError("Nao podes desativar o utilizador autenticado.")
        operadores = {str(v).strip() for v in list(data.get("operadores", []) or []) if str(v).strip()}
        operadores.discard(current_txt)
        operadores.discard(username)
        if role.lower() == "operador":
            operadores.add(username)
        data["operadores"] = sorted(operadores)
        orcamentistas = {str(v).strip() for v in list(data.get("orcamentistas", []) or []) if str(v).strip()}
        orcamentistas.discard(current_txt)
        orcamentistas.discard(username)
        if role.lower() in {"orcamentista", "orçamentista"}:
            orcamentistas.add(username)
        data["orcamentistas"] = sorted(orcamentistas)
        profiles = self._user_profiles()
        if current_key and current_key != username.lower() and current_key in profiles:
            profiles.pop(current_key, None)
        profiles[username.lower()] = {
            "posto": posto,
            "active": active,
            "menu_permissions": permissions,
        }
        self._save_user_profiles(profiles)
        self._save(force=True)
        if logged_key and logged_key in {current_key, username.lower()}:
            session_password = str(password or (self.user or {}).get("_session_password", "") or "").strip()
            if session_password:
                self.user = self.authenticate(username, session_password)
        return next((row for row in self.user_rows() if str(row.get("username", "") or "").strip().lower() == username.lower()), {})

    def remove_user(self, username: str) -> None:
        username_txt = str(username or "").strip()
        if not username_txt:
            raise ValueError("Utilizador inv?lido.")
        if self.user and str(self.user.get("username", "") or "").strip().lower() == username_txt.lower():
            raise ValueError("Nao e permitido remover o utilizador atualmente autenticado.")
        before = len(list(self.ensure_data().get("users", []) or []))
        self.ensure_data()["users"] = [row for row in list(self.ensure_data().get("users", []) or []) if str(row.get("username", "") or "").strip().lower() != username_txt.lower()]
        if len(list(self.ensure_data().get("users", []) or [])) == before:
            raise ValueError("Utilizador n?o encontrado.")
        profiles = self._user_profiles()
        profiles.pop(username_txt.lower(), None)
        self._save_user_profiles(profiles)
        self._save(force=True)
