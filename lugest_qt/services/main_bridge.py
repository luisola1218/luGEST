from __future__ import annotations

import copy
import csv
import json
import math
import os
import re
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import lugest_storage

from .laser_quote_engine import estimate_laser_quote, merge_laser_quote_settings


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


def _pdf_hex_to_rgb(value: str) -> tuple[int, int, int]:
    text = str(value or "").strip().lstrip("#")
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    if len(text) != 6:
        text = "1F3C88"
    try:
        return tuple(int(text[index : index + 2], 16) for index in (0, 2, 4))
    except Exception:
        return (31, 60, 136)


def _pdf_rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    r, g, b = [max(0, min(255, int(value))) for value in tuple(rgb or (31, 60, 136))]
    return f"#{r:02X}{g:02X}{b:02X}"


def _pdf_mix_hex(base_hex: str, target_hex: str, ratio: float) -> str:
    ratio = max(0.0, min(1.0, float(ratio)))
    base_rgb = _pdf_hex_to_rgb(base_hex)
    target_rgb = _pdf_hex_to_rgb(target_hex)
    return _pdf_rgb_to_hex(
        tuple(
            round((base_channel * (1.0 - ratio)) + (target_channel * ratio))
            for base_channel, target_channel in zip(base_rgb, target_rgb)
        )
    )


def _pdf_clip_text(value: Any, max_width: float, font_name: str, font_size: float) -> str:
    from reportlab.pdfbase import pdfmetrics

    text = "" if value is None else str(value)
    if pdfmetrics.stringWidth(text, font_name, font_size) <= max_width:
        return text
    ellipsis = "..."
    while text and pdfmetrics.stringWidth(text + ellipsis, font_name, font_size) > max_width:
        text = text[:-1]
    return f"{text}{ellipsis}" if text else ""


def _pdf_wrap_text(value: Any, font_name: str, font_size: float, max_width: float, max_lines: int | None = None) -> list[str]:
    from reportlab.pdfbase import pdfmetrics

    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return []
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip() if current else word
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
        if max_lines and len(lines) >= max_lines:
            break
    if current and (not max_lines or len(lines) < max_lines):
        lines.append(current)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
    return lines


def _pdf_fit_font_size(text: Any, font_name: str, max_width: float, preferred_size: float, min_size: float) -> float:
    from reportlab.pdfbase import pdfmetrics

    size = float(preferred_size)
    raw = str(text or "")
    max_width = max(12.0, float(max_width))
    while size > float(min_size) and pdfmetrics.stringWidth(raw, font_name, size) > max_width:
        size -= 0.3
    return max(float(min_size), round(size, 2))


class _ValueHolder:
    def __init__(self, value: str = "") -> None:
        self._value = str(value or "")

    def get(self) -> str:
        return self._value


class LegacyBackend:
    def __init__(self) -> None:
        import app_misc_actions
        import billing_pdf_actions
        import encomendas_actions
        import main as desktop_main
        import materia_actions
        import ne_expedicao_actions
        import orc_actions
        import operador_ordens_actions
        import plan_actions
        import produtos_actions
        try:
            import tax_compliance
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
        self._reload_cache_ttl_sec = 1.5
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
            raise ValueError("Credenciais invalidas.")
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

    def _save(self, force: bool = False) -> None:
        self._normalize_storage_paths_for_save()
        payload, _changed = self._merge_latest_for_save()
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
        return key

    def material_geometry_preview(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(payload or {})
        formato = str(row.get("formato") or self.desktop_main.detect_materia_formato(row) or "Chapa").strip().title() or "Chapa"
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
        else:
            peso_unid = peso_existente
            auto_weight = False

        return {
            "formato": formato,
            "secao_tipo": secao_tipo,
            "secao_label": secao_label,
            "comprimento": round(comprimento, 3),
            "largura": round(largura, 3),
            "altura": round(altura_nominal if formato == "Perfil" else altura, 3),
            "diametro": round(diametro, 3),
            "espessura": espessura,
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
            "espessura_required": formato in {"Chapa", "Tubo"},
            **geometry,
        }

    def material_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for index, material in enumerate(data.get("materiais", [])):
            if bool(material.get("is_sobra")) or str(material.get("Localizacao", material.get("Localização", "")) or "").strip().upper() == "RETALHO":
                self.materia_actions._hydrate_retalho_record(data, material)
            preview = self.material_price_preview(material)
            material["preco_unid"] = float(preview.get("preco_unid", 0.0) or 0.0)
            disponivel = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
            formato = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "Chapa").strip()
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
            if self._parse_float(material.get("quantidade", 0), 0) == 1:
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
        formato = str(payload.get("formato", "Chapa") or "Chapa").strip().title() or "Chapa"
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
        metros = float(geometry.get("metros", metros) or 0)
        peso_unid = float(geometry.get("peso_unid", peso_unid) or 0)
        kg_m = float(geometry.get("kg_m", kg_m) or 0)
        secao_tipo = str(geometry.get("secao_tipo", secao_tipo) or "").strip()
        if not material or quantidade <= 0:
            raise ValueError("Material e quantidade sao obrigatorios.")
        if formato in {"Chapa", "Tubo"} and not espessura:
            raise ValueError("Para chapa e tubo, espessura e obrigatoria.")
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
        self.desktop_main.log_stock(data, "ADICIONAR", f"{values['material']} {values['espessura']} qtd={values['quantidade']}")
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
        self.desktop_main.log_stock(data, "CORRIGIR", f"{record.get('id')} qtd={qtd} reservado={res}")
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def consume_material(self, material_id: str, quantidade: Any, retalho: dict[str, Any] | None = None) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material n?o encontrado.")
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
        self.desktop_main.log_stock(data, "BAIXA", f"{record.get('id')} qtd={qtd}")
        if retalho_row is not None:
            data.setdefault("materiais", []).append(retalho_row)
            self.desktop_main.log_stock(data, "RETALHO", f"{record.get('id')} qtd={retalho_row.get('quantidade', 0)}")
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def material_candidates(self, material: str, espessura: str, *, include_reserved: bool = False) -> list[dict[str, Any]]:
        material_norm = self.encomendas_actions._norm_material(material)
        esp_norm = self.encomendas_actions._norm_espessura(espessura)
        rows: list[dict[str, Any]] = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
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
            self.desktop_main.log_stock(data, "BAIXA", obs)

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
            self.desktop_main.log_stock(data, "RETALHO", log_msg)
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
                    "detalhes": str(entry.get("detalhes", "")),
                }
            )
        return rows

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
            rows.append(
                {
                    "data": str(entry.get("data", "") or "").replace("T", " ")[:19],
                    "acao": str(entry.get("acao", "") or "").strip(),
                    "detalhes": detalhes,
                }
            )
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
        logo_x = outer_x + 18
        logo_y = banner_y + 14
        group_x = outer_x + outer_w - card_group_w - 18
        title_left = logo_x + 96
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
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(outer_x, banner_y, outer_w, header_h, 14, stroke=0, fill=1)

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(logo_x, logo_y, 82, 42, 12, stroke=0, fill=1)
        self._draw_operator_logo(canvas_obj, logo_path, logo_x + 6, logo_y + 5, 70, 30)

        canvas_obj.setFillColor(palette["surface"])
        title = "Etiqueta de Identificacao"
        subtitle = "Chapa / palete para controlo interno"
        title_font = _pdf_fit_font_size(title, bold_font, title_w, 20.6, 15.2)
        subtitle_font = _pdf_fit_font_size(subtitle, regular_font, title_w, 8.5, 6.7)
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 47, self._operator_pdf_text(title))
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
            canvas_obj.roundRect(box_x, box_y, card_w, card_h, 8, stroke=0, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, 5.8)
            canvas_obj.drawString(box_x + 8, box_y + card_h - 8, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, bold_font, card_w - 16, 8.8, 6.0)
            canvas_obj.setFillColor(palette["primary_dark"])
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
        canvas_obj.setFillColor(palette["primary_soft_2"])
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
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_left, dim_y, body_w, 44, 12, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 8.0)
        canvas_obj.drawString(body_left + 16, dim_y + 28, self._operator_pdf_text("Dimensao identificada"))
        dim_font = _pdf_fit_font_size(dimension_text, bold_font, body_w - 32, 21.5, 13.0)
        canvas_obj.setFillColor(palette["primary_dark"])
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
        self._draw_code128_fit(canvas_obj, barcode_value, barcode_area_x, barcode_draw_y, barcode_area_w, 30, min_bar_width=0.52, max_bar_width=1.18)
        canvas_obj.setFillColor(palette["primary_dark"])
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
        outer_x = 8
        outer_y = 8
        outer_w = page_width - 16
        outer_h = page_height - 16
        header_h = 34
        banner_y = outer_y + outer_h - header_h
        logo_x = outer_x + 8
        logo_y = banner_y + 7
        chip_w = 86
        chip_h = 21
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
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(outer_x, banner_y, outer_w, header_h, 12, stroke=0, fill=1)

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(logo_x, logo_y, 42, 20, 7, stroke=0, fill=1)
        self._draw_operator_logo(canvas_obj, logo_path, logo_x + 4, logo_y + 3, 34, 14)

        canvas_obj.setFillColor(palette["surface"])
        title_left = logo_x + 50
        title_right = outer_x + outer_w - chip_w - 14
        title_w = max(52.0, title_right - title_left)
        title_font = _pdf_fit_font_size("Etiqueta Retalho", bold_font, title_w, 12.6, 9.0)
        canvas_obj.setFont(bold_font, title_font)
        canvas_obj.drawCentredString(title_left + (title_w / 2.0), banner_y + 20, self._operator_pdf_text("Etiqueta Retalho"))
        subtitle_font = _pdf_fit_font_size(barcode_value, regular_font, title_w, 7.0, 5.7)
        canvas_obj.setFont(regular_font, subtitle_font)
        canvas_obj.drawCentredString(
            title_left + (title_w / 2.0),
            banner_y + 9,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, title_w, regular_font, subtitle_font)),
        )

        chip_x = outer_x + outer_w - chip_w - 8
        chip_y = banner_y + 6
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(chip_x, chip_y, chip_w, chip_h, 7, stroke=0, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.4)
        canvas_obj.drawString(chip_x + 6, chip_y + 13, self._operator_pdf_text("Local"))
        value_font = _pdf_fit_font_size(local_text, bold_font, chip_w - 12, 8.6, 6.0)
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, value_font)
        canvas_obj.drawString(chip_x + 6, chip_y + 4.8, self._operator_pdf_text(_pdf_clip_text(local_text, chip_w - 12, bold_font, value_font)))

        body_x = outer_x + 10
        body_w = outer_w - 20
        body_top = banner_y - 10
        ref_font = _pdf_fit_font_size(barcode_value, bold_font, body_w, 14.6, 10.8)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont(bold_font, ref_font)
        canvas_obj.drawString(body_x, body_top - 4, self._operator_pdf_text(_pdf_clip_text(barcode_value, body_w, bold_font, ref_font)))

        material_font = _pdf_fit_font_size(material_title, bold_font, body_w, 11.8, 8.8)
        canvas_obj.setFont(bold_font, material_font)
        canvas_obj.drawString(body_x, body_top - 22, self._operator_pdf_text(_pdf_clip_text(material_title, body_w, bold_font, material_font)))

        dim_y = body_top - 58
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, dim_y, body_w, 28, 10, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 6.4)
        canvas_obj.drawString(body_x + 10, dim_y + 18, self._operator_pdf_text("Dimensao"))
        dim_font = _pdf_fit_font_size(dim_text, bold_font, body_w - 20, 13.4, 9.2)
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(bold_font, dim_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            dim_y + 7,
            self._operator_pdf_text(_pdf_clip_text(dim_text, body_w - 20, bold_font, dim_font)),
        )

        info_y = dim_y - 28
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
            canvas_obj.roundRect(box_x, info_y, info_w, 22, 8, stroke=1, fill=1)
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
        barcode_h = info_y - barcode_y - 8
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, barcode_y, body_w, barcode_h, 10, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.8)
        canvas_obj.drawString(body_x + 8, barcode_y + barcode_h - 11, self._operator_pdf_text("Codigo para picagem"))
        self._draw_code128_fit(canvas_obj, barcode_value, body_x + 8, barcode_y + 13, body_w - 16, 18, min_bar_width=0.5, max_bar_width=1.18)
        canvas_obj.setFillColor(palette["primary_dark"])
        human_font = _pdf_fit_font_size(barcode_value, bold_font, body_w - 16, 8.2, 6.4)
        canvas_obj.setFont(bold_font, human_font)
        canvas_obj.drawCentredString(
            body_x + (body_w / 2.0),
            barcode_y + 3.5,
            self._operator_pdf_text(_pdf_clip_text(barcode_value, body_w - 16, bold_font, human_font)),
        )
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(regular_font, 5.0)
        canvas_obj.drawRightString(outer_x + outer_w - 8, outer_y + 4.5, self._operator_pdf_text(printed_at[:16]))

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

    def product_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for index, prod in enumerate(list(self.ensure_data().get("produtos", []) or [])):
            price_unit = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
            qty = self._parse_float(prod.get("qty", 0), 0)
            alerta = self._parse_float(prod.get("alerta", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "descricao": str(prod.get("descricao", "") or "").strip(),
                "categoria": str(prod.get("categoria", "") or "").strip(),
                "subcat": str(prod.get("subcat", "") or "").strip(),
                "tipo": str(prod.get("tipo", "") or "").strip(),
                "dimensoes": self._product_dimensoes(prod),
                "unid": str(prod.get("unid", "UN") or "UN").strip() or "UN",
                "qty": qty,
                "alerta": alerta,
                "p_compra": round(self._parse_float(prod.get("p_compra", 0), 0), 4),
                "preco_unid": price_unit,
                "valor_stock": round(qty * price_unit, 2),
                "metros_unidade": round(self._parse_float(prod.get("metros_unidade", prod.get("metros", 0)), 0), 4),
                "peso_unid": round(self._parse_float(prod.get("peso_unid", 0), 0), 4),
                "fabricante": str(prod.get("fabricante", "") or "").strip(),
                "modelo": str(prod.get("modelo", "") or "").strip(),
                "obs": str(prod.get("obs", "") or "").strip(),
                "updated_at": str(prod.get("atualizado_em", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            if qty <= 0 or (alerta > 0 and qty <= alerta):
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
        values = list(getattr(self.desktop_main, "PLANEAMENTO_OPERACOES_DISPONIVEIS", []) or [])
        if not values:
            values = ["Corte Laser", "Quinagem", "Serralharia", "Maquinacao", "Roscagem", "Lacagem", "Montagem"]
        ordered: list[str] = []
        for value in values:
            normalized = self._planning_normalize_operation(value)
            if normalized and normalized not in ordered:
                ordered.append(normalized)
        return ordered

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
            "Corte Laser": {"Corte Laser"},
            "Quinagem": {"Quinagem"},
            "Serralharia": {"Serralharia", "Soldadura"},
            "Maquinacao": {"Maquinacao"},
            "Roscagem": {"Roscagem"},
            "Lacagem": {"Lacagem", "Pintura"},
            "Montagem": {"Montagem"},
        }
        return set(alias_map.get(op_txt, {op_txt}))

    def _planning_operation_from_piece_name(self, operation: Any) -> str:
        normalize_fn = getattr(self.desktop_main, "normalize_planeamento_operacao", None)
        if callable(normalize_fn):
            return str(normalize_fn(operation or "") or "").strip()
        return str(self.desktop_main.normalize_operacao_nome(operation or "") or "").strip()

    def _planning_default_posto_for_operation(self, operation: Any, numero: str = "") -> str:
        op_txt = self._planning_normalize_operation(operation)
        if op_txt == "Corte Laser":
            posto_txt = self._normalize_workcenter_value(self._order_workcenter(numero))
            return posto_txt or "Corte Laser"
        return op_txt or "Geral"

    def _planning_row_operation(self, row: dict[str, Any] | None, default: str = "Corte Laser") -> str:
        return self._planning_normalize_operation((row or {}).get("operacao", ""), default=default)

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
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - qty_res)
            stock["reservado"] = max(0.0, self._parse_float(stock.get("reservado", 0), 0) - qty_res)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            if not lote_sel:
                lote_sel = str(stock.get("lote_fornecedor", "") or "").strip()
            reserved_consumed += qty_res
            self.desktop_main.log_stock(self.ensure_data(), "BAIXA CATIVADA", f"{stock.get('id', '')} qtd={qty_res} encomenda={enc.get('numero', '')}")
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
            if self._norm_material_token(stock.get("material")) != self._norm_material_token(material) or self._norm_esp_token(stock.get("espessura")) != self._norm_esp_token(espessura):
                raise ValueError("O stock selecionado n?o corresponde ao material/espessura.")
            if extra_qty > self._parse_float(stock.get("quantidade", 0), 0):
                raise ValueError("Quantidade superior ao stock disponivel.")
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - extra_qty)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            if not lote_sel:
                lote_sel = str(stock.get("lote_fornecedor", "") or "").strip()
            extra_consumed = extra_qty
            self.desktop_main.log_stock(self.ensure_data(), "BAIXA", f"{stock_id} qtd={extra_qty} encomenda={enc.get('numero', '')}")

        consumed_total = round(reserved_consumed + extra_consumed, 1)
        manual_stock_required = reserved_consumed <= 1e-9
        if manual_stock_required and extra_consumed <= 1e-9 and not allow_without_stock:
            raise ValueError("Falta registar a baixa do material consumido.")
        if manual_stock_required and extra_consumed <= 1e-9 and allow_without_stock:
            self.desktop_main.log_stock(
                self.ensure_data(),
                "SEM_BAIXA",
                f"encomenda={enc.get('numero', '')} mat={material} esp={espessura} motivo=laser_sem_stock_qt",
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
        posto_norm = self.desktop_main.norm_text(posto or "")
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
            for posto in self.available_postos():
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
        generator = getattr(self.desktop_main, "next_opp_numero", None)
        if not callable(generator):
            return False
        try:
            piece["opp"] = str(generator(self.ensure_data()) or "").strip()
        except Exception:
            return False
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
        logo_box_x = outer_x + 9
        logo_box_y = outer_y + outer_h - banner_h + 5
        chip_x = outer_x + outer_w - right_chip_w - 8
        chip_y = outer_y + outer_h - banner_h + 5

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 11, stroke=1, fill=1)

        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(outer_x, outer_y + outer_h - banner_h, outer_w, banner_h, 11, stroke=0, fill=1)
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(logo_box_x, logo_box_y, 42, 20, 6, stroke=0, fill=1)
        self._draw_operator_logo(canvas_obj, logo_path, logo_box_x + 3, logo_box_y + 2, 36, 16)

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

        title_x_left = logo_box_x + 42 + 10
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
        group_x = banner_x + page_inner_w - card_group_w - 12
        logo_box_x = banner_x + 18
        logo_box_y = banner_y + 17
        title_left = logo_box_x + 94
        title_right = group_x - 12
        title_w = max(120.0, title_right - title_left)

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(banner_x, banner_y, page_inner_w, header_h, 16, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["primary"])
        canvas_obj.roundRect(banner_x, banner_y, page_inner_w, header_h, 16, stroke=0, fill=1)

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.roundRect(logo_box_x, logo_box_y, 82, 46, 12, stroke=0, fill=1)
        self._draw_operator_logo(canvas_obj, logo_path, logo_box_x + 6, logo_box_y + 6, 70, 32)

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
            if item_type != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            consumed = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            pending = round(max(0.0, plan - consumed), 2)
            if pending <= 1e-9:
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
        shortages.sort(key=lambda row: (-self._parse_float(row.get("qtd_em_falta", 0), 0), str(row.get("produto_codigo", "") or "")))
        return shortages

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
                code = str(shortage.get("produto_codigo", "") or "").strip()
                key = code or str(shortage.get("descricao", "") or "").strip()
                entry = grouped.setdefault(
                    key,
                    {
                        "produto_codigo": code,
                        "descricao": str(shortage.get("descricao", "") or "").strip(),
                        "produto_unid": str(shortage.get("produto_unid", "") or "UN").strip() or "UN",
                        "preco_unit": round(self._parse_float(shortage.get("preco_unit", 0), 0), 4),
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
                    {
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
                    "espessura": str(piece.get("espessura", "")).strip(),
                    "estado": str(piece.get("estado", "")).strip(),
                    "qtd_plan": self._fmt(qty_plan),
                    "qtd_prod": self._fmt(qty_prod),
                    "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                    "operacoes": " + ".join(
                        [self.desktop_main.normalize_operacao_nome(op.get("nome", "")) for op in list(ops or []) if str(op.get("nome", "")).strip()]
                    ),
                    "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                    "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                }
            )
        materials = []
        materials_tree = []
        for mat in list(enc.get("materiais", []) or []):
            esp_rows = []
            for esp in list(mat.get("espessuras", []) or []):
                op_times = self._planning_operation_times_map(esp)
                planning_ops = [op for op in self._planning_ops_from_esp_obj(esp) if op != "Montagem"]
                other_ops = [op for op in planning_ops if op != "Corte Laser"]
                ops_summary = []
                for op_name in other_ops:
                    op_value = str(op_times.get(op_name, "") or "").strip()
                    if op_value:
                        ops_summary.append(f"{op_name}: {self._fmt(op_value)} min")
                    else:
                        ops_summary.append(op_name)
                esp_row = {
                    "material": str(mat.get("material", "")).strip(),
                    "espessura": str(esp.get("espessura", "")).strip(),
                    "estado": str(esp.get("estado", "")).strip(),
                    "tempo_min": self._fmt(esp.get("tempo_min", 0)),
                    "tempos_operacao": {op: self._fmt(value) for op, value in op_times.items() if str(value or "").strip()},
                    "operacoes_planeamento": planning_ops,
                    "tempo_operacoes_txt": " | ".join(ops_summary) or "-",
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
                    "tipo_label": self.desktop_main.orc_line_type_label(item_type),
                    "descricao": str(item.get("descricao", "") or "").strip(),
                    "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(item.get("produto_unid", "") or "").strip(),
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
                self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) == self.desktop_main.ORC_LINE_TYPE_PRODUCT
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
                suggested_candidate = dict(standard_pool[min(index - 1, len(standard_pool) - 1)]) if standard_pool else {}
                if suggested_candidate:
                    need["preferred_lot"] = str(suggested_candidate.get("lote", "") or "").strip()
                    need["preferred_material_id"] = str(suggested_candidate.get("material_id", "") or "").strip()
                    need["preferred_dimensao"] = str(suggested_candidate.get("dimensao", "") or "").strip()
                    need["preferred_disponivel"] = round(self._parse_float(suggested_candidate.get("disponivel", 0), 0), 2)
                else:
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
                    preferred_standard = dict(standard_candidates[0]) if standard_candidates else {}
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
                        "current_lot": current_lot,
                        "preferred_lot": str(preferred_standard.get("lote", "") or "").strip(),
                        "preferred_material_id": str(preferred_standard.get("material_id", "") or "").strip(),
                        "preferred_dimensao": str(preferred_standard.get("dimensao", "") or "").strip(),
                        "preferred_disponivel": round(self._parse_float(preferred_standard.get("disponivel", 0), 0), 2),
                        "retalho_count": len(retalho_candidates),
                        "standard_lot_count": len(standard_candidates),
                        "piece_qty": piece_qty,
                        "chapa": self._order_reserved_sheet(numero, mat_name, esp),
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
            preferred_lot = str(need.get("preferred_lot", "") or "-").strip() or "-"
            recommendation = (
                f"Separar lote {preferred_lot}"
                if preferred_lot and preferred_lot != "-"
                else "Validar materia-prima"
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
                reserva_label = "Cativado" if source_kind == "reserva" else ("Retalho sugerido" if is_retalho else "Por separar")
                action_text = recommendation
                if source_kind == "reserva":
                    action_text = f"Separar {lote_sugerido} (cativado)"
                elif is_retalho:
                    action_text = f"Avaliar retalho {lote_sugerido}"
                elif lote_sugerido and lote_sugerido != "-":
                    action_text = f"Separar lote {lote_sugerido}"
                return {
                    "numero": str(need.get("numero", "") or "").strip(),
                    "cliente": str(need.get("cliente", "") or "").strip(),
                    "posto_trabalho": posto_trabalho,
                    "material": material_label,
                    "espessura": espessura_label,
                    "dimensao": dimensao,
                    "quantidade": round(self._parse_float(quantity_value, 0), 2),
                    "disponivel": round(self._parse_float(source_candidate.get("disponivel", need.get("preferred_disponivel", 0)), 0), 2),
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
                    "priority_label": str(lead.get("priority_label", "") or ("Alta" if not bool(need.get("stock_ready")) else "Media")).strip(),
                    "priority_tone": str(lead.get("priority_tone", "") or ("danger" if not bool(need.get("stock_ready")) else "info")).strip(),
                    "priority_score": int(lead.get("priority_score", 0) or 0),
                    "status_label": str(lead.get("status_label", "") or "Pendente").strip(),
                    "status_key": str(lead.get("status_key", "") or "new").strip(),
                    "stock_ready": bool(need.get("stock_ready")),
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
                    source_candidate=dict(standard_candidates[0]) if standard_candidates else {},
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
                    c.drawRightString(
                        width - margin - 8,
                        current_y - 10,
                        (
                            f"{group_ref.get('planeamento_dia', '-')} | "
                            f"{group_ref.get('planeamento_turno', '-')} | "
                            f"{group_ref.get('cliente', '-')} | "
                            f"{len(formatos)} formatos | {self._fmt(group_qty)} un. | {len(lotes)} lotes"
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
                            self._fmt(row.get("quantidade", 0)),
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
            "shortage": 2,
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
                "shortage": "Sem stock",
            }
            return mapping.get(kind, "Sugestao operacional")

        def _kind_color(row: dict[str, Any]) -> str:
            kind = str(row.get("kind", "") or "").strip()
            mapping = {
                "fito_lot": "#0f3d91",
                "keep_ready": "#b45309",
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
                    piece["opp"] = self.desktop_main.next_opp_numero(data)
                    changed = True
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self.desktop_main.next_of_numero(data)
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
            "operacoes": list(self.desktop_main.OFF_OPERACOES_DISPONIVEIS),
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
        ordered = [op_name for op_name in list(self.desktop_main.OFF_OPERACOES_DISPONIVEIS) if op_name in ops]
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
            }
            data.setdefault("encomendas", []).append(enc)

        enc["cliente"] = cliente
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
        mat.setdefault("espessuras", []).append({"espessura": esp_txt, "tempo_min": "", "tempos_operacao": {}, "estado": "Preparacao", "pecas": []})
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
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            raise ValueError("Espessura não encontrada.")
        cleaned: dict[str, str] = {}
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
        esp["tempo_min"] = cleaned.get("Corte Laser", str(esp.get("tempo_min", "") or "").strip())
        if not cleaned.get("Corte Laser"):
            esp["tempo_min"] = ""
        esp["tempos_operacao"] = cleaned
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
        material = str(payload.get("material", "") or "").strip()
        espessura = str(payload.get("espessura", "") or "").strip()
        descricao = str(payload.get("descricao", "") or "").strip()
        desenho = str(payload.get("desenho", "") or "").strip()
        operacoes = " + ".join(self.desktop_main.parse_operacoes_lista(payload.get("operacoes", "")))
        if not operacoes:
            operacoes = str(self.desktop_main.OFF_OPERACAO_OBRIGATORIA)
        quantidade = self._parse_float(payload.get("quantidade_pedida", 0), 0)
        preco_unit = self._parse_float(payload.get("preco_unit", 0), 0)
        tempos_operacao = dict(payload.get("tempos_operacao", {}) or {})
        custos_operacao = dict(payload.get("custos_operacao", {}) or {})
        operacoes_detalhe = [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        guardar_ref = bool(payload.get("guardar_ref", True))

        if not material or not espessura:
            raise ValueError("Material e espessura sao obrigatorios.")
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
            esp = {"espessura": espessura, "tempo_min": "", "tempos_operacao": {}, "estado": "Preparacao", "pecas": []}
            mat.setdefault("espessuras", []).append(esp)

        _, old_esp, piece = self._order_find_piece(enc, current_ref, "")
        if piece is None:
            piece = {
                "id": self._next_order_piece_id(enc),
                "of": self.desktop_main.next_of_numero(self.ensure_data()),
                "opp": self.desktop_main.next_opp_numero(self.ensure_data()),
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
        piece["espessura"] = espessura
        piece["descricao"] = descricao
        piece["quantidade_pedida"] = quantidade
        piece["Operacoes"] = operacoes
        piece["Observacoes"] = descricao
        piece["desenho"] = desenho
        piece["desenho_path"] = desenho
        piece["tempos_operacao"] = dict(tempos_operacao)
        piece["custos_operacao"] = dict(custos_operacao)
        piece["operacoes_detalhe"] = list(operacoes_detalhe)
        if "of" not in piece or not str(piece.get("of", "")).strip():
            piece["of"] = self.desktop_main.next_of_numero(self.ensure_data())
        if "opp" not in piece or not str(piece.get("opp", "")).strip():
            piece["opp"] = self.desktop_main.next_opp_numero(self.ensure_data())
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
                "Operacoes": operacoes,
                "Observacoes": descricao,
                "desenho": desenho,
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
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

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
            self.desktop_main.log_stock(self.ensure_data(), "CATIVAR", f"{material_id} qtd={quantidade} encomenda={enc.get('numero', '')}")
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
            self.desktop_main.log_stock(self.ensure_data(), "LIBERTAR", f"{row.get('material_id', '')} qtd={row.get('quantidade', 0)} encomenda={enc.get('numero', '')}")
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
            if self.desktop_main.normalize_orc_line_type(item.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            product = product_map.get(code)
            if product is None:
                shortages.append(f"{code or '-'}: produto nao encontrado")
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
            else:
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Concluido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                changed = True
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
            {"key": "pulse", "label": "Pulse"},
            {"key": "avarias", "label": "Avarias"},
            {"key": "home", "label": "Resumo"},
        ]

    def available_roles(self) -> list[str]:
        return ["Admin", "Producao", "Qualidade", "Planeamento", "Orcamentista", "Operador"]

    def quote_workcenter_options(self) -> list[str]:
        preferred = ["Maquina 3030", "Maquina 5030", "Maquina 5040"]
        existing = [
            str(value or "").strip()
            for value in list(self.ensure_data().get("postos_trabalho", []) or [])
            if str(value or "").strip()
        ]
        ordered: list[str] = []
        seen: set[str] = set()
        for posto in preferred + existing:
            key = str(posto or "").strip().lower()
            if not key or key in seen or key == "geral":
                continue
            seen.add(key)
            ordered.append(str(posto).strip())
        return ordered or preferred

    def _normalize_workcenter_value(self, value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        for posto in self.quote_workcenter_options():
            if raw.lower() == str(posto or "").strip().lower():
                return str(posto or "").strip()
        return raw

    def available_postos(self) -> list[str]:
        data = self.ensure_data()
        postos = [str(value or "").strip() for value in list(data.get("postos_trabalho", []) or []) if str(value or "").strip()]
        if not postos:
            postos = ["Laser 1", "Quinagem 1", "Roscagem 1", "Soldadura 1", "Embalamento 1"]
        postos = ["Geral"] + self.quote_workcenter_options() + postos
        seen: set[str] = set()
        ordered: list[str] = []
        for posto in postos:
            key = posto.lower()
            if key in seen:
                continue
            seen.add(key)
            ordered.append(posto)
        return ordered

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

    def expedicao_pending_orders(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Todas").strip()
        rows = []
        for enc in data.get("encomendas", []):
            self.desktop_main.update_estado_expedicao_encomenda(enc)
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            disponivel = sum(max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(p), 0)) for p in pieces)
            if disponivel <= 0:
                continue
            estado_exp = str(enc.get("estado_expedicao", "Nao expedida") or "Nao expedida").strip()
            if estado_filter != "Todas" and estado_exp != estado_filter:
                continue
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn):
                cli_obj = find_cliente_fn(data, cli_code) or {}
            cliente_txt = " - ".join([part for part in [cli_code, str(cli_obj.get("nome", "") or "").strip()] if part]).strip()
            row = {
                "numero": str(enc.get("numero", "") or "").strip(),
                "cliente": cliente_txt or cli_code or "-",
                "estado": str(enc.get("estado", "") or "").strip(),
                "estado_expedicao": estado_exp,
                "disponivel": round(disponivel, 1),
                "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: ((item.get("data_entrega") or "9999-99-99"), item.get("numero") or ""))
        return rows

    def expedicao_available_pieces(self, enc_num: str) -> list[dict[str, Any]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            return []
        rows = []
        for piece in self.desktop_main.encomenda_pecas(enc):
            self.desktop_main.ensure_peca_operacoes(piece)
            pronta = max(0.0, self._parse_float(getattr(self.desktop_main, "peca_qtd_pronta_expedicao")(piece), 0))
            disponivel = max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(piece), 0))
            if disponivel <= 0:
                continue
            rows.append(
                {
                    "id": str(piece.get("id", "") or "").strip(),
                    "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                    "descricao": str(
                        piece.get("Observacoes")
                        or piece.get("Observações")
                        or piece.get("descricao")
                        or piece.get("ref_externa")
                        or piece.get("ref_interna")
                        or ""
                    ).strip(),
                    "estado": str(piece.get("estado", "") or "").strip(),
                    "pronta_expedicao": self._fmt(pronta),
                    "qtd_expedida": self._fmt(piece.get("qtd_expedida", 0)),
                    "disponivel": self._fmt(disponivel),
                    "pronta_expedicao_num": pronta,
                    "qtd_expedida_num": self._parse_float(piece.get("qtd_expedida", 0), 0),
                    "disponivel_num": disponivel,
                    "material": str(piece.get("material", "") or "").strip(),
                    "espessura": str(piece.get("espessura", "") or "").strip(),
                    "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                }
            )
        rows.sort(key=lambda item: (item.get("ref_interna") or "", item.get("ref_externa") or ""))
        return rows

    def expedicao_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for ex in sorted(list(self.ensure_data().get("expedicoes", []) or []), key=lambda row: str(row.get("data_emissao", "") or ""), reverse=True):
            row = {
                "numero": str(ex.get("numero", "") or "").strip(),
                "tipo": str(ex.get("tipo", "") or "").strip(),
                "encomenda": str(ex.get("encomenda", "") or "").strip(),
                "cliente": str(ex.get("cliente_nome", "") or ex.get("cliente", "") or "").strip(),
                "data_emissao": str(ex.get("data_emissao", "") or "").replace("T", " ")[:19],
                "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
                "linhas": len(list(ex.get("linhas", []) or [])),
                "anulada": bool(ex.get("anulada")),
                "transportador": str(ex.get("transportador", "") or "").strip(),
                "matricula": str(ex.get("matricula", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def expedicao_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        lines = []
        for line in list(ex.get("linhas", []) or []):
            lines.append(
                {
                    "peca_id": str(line.get("peca_id", "") or "").strip(),
                    "ref_interna": str(line.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": self._fmt(line.get("qtd", 0)),
                    "peso": self._fmt(line.get("peso", 0)),
                    "manual": bool(line.get("manual")),
                    "encomenda": str(line.get("encomenda", "") or ex.get("encomenda", "") or "").strip(),
                }
            )
        return {
            "numero": str(ex.get("numero", "") or "").strip(),
            "tipo": str(ex.get("tipo", "") or "").strip(),
            "encomenda": str(ex.get("encomenda", "") or "").strip(),
            "cliente": str(ex.get("cliente_nome", "") or ex.get("cliente", "") or "").strip(),
            "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
            "data_emissao": str(ex.get("data_emissao", "") or "").replace("T", " ")[:19],
            "data_transporte": str(ex.get("data_transporte", "") or "").replace("T", " ")[:19],
            "transportador": str(ex.get("transportador", "") or "").strip(),
            "matricula": str(ex.get("matricula", "") or "").strip(),
            "destinatario": str(ex.get("destinatario", "") or "").strip(),
            "local_descarga": str(ex.get("local_descarga", "") or "").strip(),
            "observacoes": str(ex.get("observacoes", "") or "").strip(),
            "anulada_motivo": str(ex.get("anulada_motivo", "") or "").strip(),
            "lines": lines,
        }

    def expedicao_open_pdf(self, numero: str) -> Path:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        path = Path(tempfile.gettempdir()) / f"lugest_guia_{numero}.pdf"
        self.ne_expedicao_actions.render_expedicao_pdf(self, str(path), ex)
        os.startfile(str(path))
        return path

    def expedicao_render_pdf(self, numero: str, path: str | Path, include_all_vias: bool = False) -> Path:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        out_path = Path(path)
        self.ne_expedicao_actions.render_expedicao_pdf(self, str(out_path), ex, include_all_vias=include_all_vias)
        return out_path

    def _exp_validation_code(self, issue_date: str | None = None) -> str:
        issue_date = str(issue_date or self.desktop_main.now_iso())
        serie_guess = ""
        default_serie_fn = getattr(self.desktop_main, "_exp_default_serie_id", None)
        if callable(default_serie_fn):
            try:
                serie_guess = str(default_serie_fn("GT", issue_date) or "").strip()
            except Exception:
                serie_guess = ""
        find_series_fn = getattr(self.desktop_main, "_find_at_series", None)
        if not callable(find_series_fn):
            return ""
        try:
            serie_obj = find_series_fn(self.ensure_data(), doc_type="GT", serie_id=serie_guess) or {}
        except Exception:
            serie_obj = {}
        return str(serie_obj.get("validation_code", "") or "").strip()

    def expedicao_defaults_for_order(self, enc_num: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        cli_code = str(enc.get("cliente", "") or "").strip()
        cli = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn):
            cli = find_cliente_fn(self.ensure_data(), cli_code) or {}
        cli_nome = str(cli.get("nome", "") or cli_code or "").strip()
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        local_carga = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        return {
            "codigo_at": self._exp_validation_code(self.desktop_main.now_iso()),
            "tipo_via": "Original",
            "emitente_nome": str(emit_cfg.get("nome", "") or "").strip(),
            "emitente_nif": str(emit_cfg.get("nif", "") or "").strip(),
            "emitente_morada": str(emit_cfg.get("morada", "") or "").strip(),
            "destinatario": cli_nome,
            "dest_nif": str(cli.get("nif", "") or "").strip(),
            "dest_morada": str(cli.get("morada", "") or "").strip(),
            "local_carga": local_carga,
            "local_descarga": str(cli.get("morada", "") or "").strip(),
            "data_transporte": str(self.desktop_main.now_iso()),
            "transportador": "",
            "matricula": "",
            "observacoes": f"Expedicao da encomenda {enc_num}",
        }

    def expedicao_manual_defaults(self) -> dict[str, Any]:
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        local_carga = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        return {
            "codigo_at": self._exp_validation_code(self.desktop_main.now_iso()),
            "tipo_via": "Original",
            "emitente_nome": str(emit_cfg.get("nome", "") or "").strip(),
            "emitente_nif": str(emit_cfg.get("nif", "") or "").strip(),
            "emitente_morada": str(emit_cfg.get("morada", "") or "").strip(),
            "destinatario": "",
            "dest_nif": "",
            "dest_morada": "",
            "local_carga": local_carga,
            "local_descarga": "",
            "data_transporte": str(self.desktop_main.now_iso()),
            "transportador": "",
            "matricula": "",
            "observacoes": "",
        }

    def expedicao_product_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for prod in list(self.ensure_data().get("produtos", []) or []):
            qty = self._parse_float(prod.get("qty", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "descricao": str(prod.get("descricao", "") or "").strip(),
                "qty": qty,
                "unid": str(prod.get("unid", "UN") or "UN").strip() or "UN",
                "qty_fmt": self._fmt(qty),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("descricao") or ""))
        return rows

    def expedicao_emit_off(self, enc_num: str, draft_lines: list[dict[str, Any]], guide_data: dict[str, Any]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        lines = [dict(line) for line in list(draft_lines or []) if isinstance(line, dict)]
        if not lines:
            raise ValueError("Sem linhas na guia.")
        requested_by_piece: dict[str, float] = {}
        pieces = {str(piece.get("id", "") or ""): piece for piece in self.desktop_main.encomenda_pecas(enc)}
        for line in lines:
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if not piece_id or qty <= 0:
                raise ValueError("Linha de guia invalida.")
            piece = pieces.get(piece_id)
            if piece is None:
                raise ValueError(f"Pe?a n?o encontrada para expedi??o: {piece_id}")
            requested_by_piece[piece_id] = requested_by_piece.get(piece_id, 0.0) + qty
        for piece_id, qty in requested_by_piece.items():
            available = self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(pieces[piece_id]), 0)
            if qty > available + 1e-9:
                raise ValueError(f"Quantidade superior ao disponivel na peca {pieces[piece_id].get('ref_interna', piece_id)}.")
        cli_code = str(enc.get("cliente", "") or "").strip()
        cli = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn):
            cli = find_cliente_fn(self.ensure_data(), cli_code) or {}
        cli_nome = str(guide_data.get("destinatario", "") or cli.get("nome", "") or cli_code).strip()
        exp_ids, exp_err = self.desktop_main.next_expedicao_identifiers(
            self.ensure_data(),
            issue_date=str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            doc_type="GT",
            validation_code_hint=str(guide_data.get("codigo_at", "") or "").strip(),
        )
        if not exp_ids:
            raise ValueError(exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
        ex_num = str(exp_ids.get("numero", "") or "").strip()
        ex = {
            "numero": ex_num,
            "tipo": "OFF",
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": cli_code,
            "cliente_nome": cli_nome,
            "codigo_at": exp_ids.get("validation_code", ""),
            "serie_id": exp_ids.get("serie_id", ""),
            "seq_num": exp_ids.get("seq_num", 0),
            "at_validation_code": exp_ids.get("validation_code", ""),
            "atcud": exp_ids.get("atcud", ""),
            "tipo_via": str(guide_data.get("tipo_via", "Original") or "Original"),
            "emitente_nome": str(guide_data.get("emitente_nome", "") or ""),
            "emitente_nif": str(guide_data.get("emitente_nif", "") or ""),
            "emitente_morada": str(guide_data.get("emitente_morada", "") or ""),
            "destinatario": cli_nome,
            "dest_nif": str(guide_data.get("dest_nif", "") or cli.get("nif", "") or ""),
            "dest_morada": str(guide_data.get("dest_morada", "") or cli.get("morada", "") or ""),
            "local_carga": str(guide_data.get("local_carga", "") or ""),
            "local_descarga": str(guide_data.get("local_descarga", "") or ""),
            "data_emissao": self.desktop_main.now_iso(),
            "data_transporte": str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            "matricula": str(guide_data.get("matricula", "") or ""),
            "transportador": str(guide_data.get("transportador", "") or ""),
            "estado": "Emitida",
            "observacoes": str(guide_data.get("observacoes", "") or ""),
            "created_by": str((self.user or {}).get("username", "") or ""),
            "anulada": False,
            "anulada_motivo": "",
            "linhas": [],
        }
        for line in lines:
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            piece = pieces[piece_id]
            piece["qtd_expedida"] = self._parse_float(piece.get("qtd_expedida", 0), 0) + qty
            piece.setdefault("expedicoes", []).append(ex_num)
            ex["linhas"].append(
                {
                    "encomenda": str(enc.get("numero", "") or "").strip(),
                    "peca_id": piece_id,
                    "ref_interna": str(line.get("ref_interna", "") or piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(line.get("ref_externa", "") or piece.get("ref_externa", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": qty,
                    "unid": str(line.get("unid", "UN") or "UN").strip() or "UN",
                    "peso": self._parse_float(line.get("peso", 0), 0),
                    "manual": False,
                }
            )
            try:
                self.desktop_main.atualizar_estado_peca(piece)
            except Exception:
                pass
        self.ensure_data().setdefault("expedicoes", []).append(ex)
        self.desktop_main.update_estado_expedicao_encomenda(enc)
        self._save(force=True)
        return self.expedicao_detail(ex_num)

    def expedicao_emit_manual(self, guide_data: dict[str, Any], lines: list[dict[str, Any]]) -> dict[str, Any]:
        clean_lines = [dict(line) for line in list(lines or []) if isinstance(line, dict)]
        if not clean_lines:
            raise ValueError("Sem linhas na guia.")
        products = {str(prod.get("codigo", "") or "").strip(): prod for prod in self.ensure_data().get("produtos", [])}
        for line in clean_lines:
            code = str(line.get("produto_codigo", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if qty <= 0:
                raise ValueError("Quantidade invalida numa linha da guia.")
            if code:
                prod = products.get(code)
                if prod is None:
                    raise ValueError(f"Produto nao encontrado: {code}")
                if qty > self._parse_float(prod.get("qty", 0), 0) + 1e-9:
                    raise ValueError(f"Stock insuficiente para {code}.")
        exp_ids, exp_err = self.desktop_main.next_expedicao_identifiers(
            self.ensure_data(),
            issue_date=str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            doc_type="GT",
            validation_code_hint=str(guide_data.get("codigo_at", "") or "").strip(),
        )
        if not exp_ids:
            raise ValueError(exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
        ex_num = str(exp_ids.get("numero", "") or "").strip()
        ex = {
            "numero": ex_num,
            "tipo": "Manual",
            "encomenda": "",
            "cliente": "",
            "cliente_nome": str(guide_data.get("destinatario", "") or "").strip(),
            "codigo_at": exp_ids.get("validation_code", ""),
            "serie_id": exp_ids.get("serie_id", ""),
            "seq_num": exp_ids.get("seq_num", 0),
            "at_validation_code": exp_ids.get("validation_code", ""),
            "atcud": exp_ids.get("atcud", ""),
            "tipo_via": str(guide_data.get("tipo_via", "Original") or "Original"),
            "emitente_nome": str(guide_data.get("emitente_nome", "") or ""),
            "emitente_nif": str(guide_data.get("emitente_nif", "") or ""),
            "emitente_morada": str(guide_data.get("emitente_morada", "") or ""),
            "destinatario": str(guide_data.get("destinatario", "") or "").strip(),
            "dest_nif": str(guide_data.get("dest_nif", "") or "").strip(),
            "dest_morada": str(guide_data.get("dest_morada", "") or "").strip(),
            "local_carga": str(guide_data.get("local_carga", "") or ""),
            "local_descarga": str(guide_data.get("local_descarga", "") or ""),
            "data_emissao": self.desktop_main.now_iso(),
            "data_transporte": str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            "matricula": str(guide_data.get("matricula", "") or ""),
            "transportador": str(guide_data.get("transportador", "") or ""),
            "estado": "Emitida",
            "observacoes": str(guide_data.get("observacoes", "") or ""),
            "created_by": str((self.user or {}).get("username", "") or ""),
            "anulada": False,
            "anulada_motivo": "",
            "linhas": [],
        }
        for line in clean_lines:
            code = str(line.get("produto_codigo", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if code:
                prod = products.get(code)
                if prod is not None:
                    prod["qty"] = max(0.0, self._parse_float(prod.get("qty", 0), 0) - qty)
                    prod["atualizado_em"] = self.desktop_main.now_iso()
            ex["linhas"].append(
                {
                    "encomenda": "",
                    "peca_id": "",
                    "ref_interna": code,
                    "ref_externa": "",
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": qty,
                    "unid": str(line.get("unid", "UN") or "UN").strip() or "UN",
                    "peso": 0.0,
                    "manual": True,
                }
            )
        self.ensure_data().setdefault("expedicoes", []).append(ex)
        self._save(force=True)
        return self.expedicao_detail(ex_num)

    def expedicao_update(self, numero: str, guide_data: dict[str, Any]) -> dict[str, Any]:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        if bool(ex.get("anulada")):
            raise ValueError("Nao e possivel editar uma guia anulada.")
        ex["codigo_at"] = str(guide_data.get("codigo_at", "") or "").strip()
        ex["at_validation_code"] = str(guide_data.get("codigo_at", "") or "").strip()
        seq_num = int(self._parse_float(ex.get("seq_num", 0), 0) or 0)
        if ex.get("at_validation_code") and seq_num > 0:
            ex["atcud"] = f"{str(ex.get('at_validation_code', '')).strip()}-{seq_num}"
        ex["tipo_via"] = "Original"
        ex["emitente_nome"] = str(guide_data.get("emitente_nome", "") or "").strip()
        ex["emitente_nif"] = str(guide_data.get("emitente_nif", "") or "").strip()
        ex["emitente_morada"] = str(guide_data.get("emitente_morada", "") or "").strip()
        ex["destinatario"] = str(guide_data.get("destinatario", "") or "").strip()
        ex["dest_nif"] = str(guide_data.get("dest_nif", "") or "").strip()
        ex["dest_morada"] = str(guide_data.get("dest_morada", "") or "").strip()
        ex["local_carga"] = str(guide_data.get("local_carga", "") or "").strip()
        ex["local_descarga"] = str(guide_data.get("local_descarga", "") or "").strip()
        ex["data_transporte"] = str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso())
        ex["transportador"] = str(guide_data.get("transportador", "") or "").strip()
        ex["matricula"] = str(guide_data.get("matricula", "") or "").strip()
        ex["observacoes"] = str(guide_data.get("observacoes", "") or "").strip()
        self._save(force=True)
        return self.expedicao_detail(numero)

    def expedicao_cancel(self, numero: str, reason: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        motivo = str(reason or "").strip()
        if not motivo:
            raise ValueError("E obrigatorio indicar justificacao.")
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        if bool(ex.get("anulada")):
            return self.expedicao_detail(numero)
        for line in list(ex.get("linhas", []) or []):
            if bool(line.get("manual")):
                code = str(line.get("ref_interna", "") or "").strip()
                qty = self._parse_float(line.get("qtd", 0), 0)
                if code:
                    prod = next((row for row in self.ensure_data().get("produtos", []) if str(row.get("codigo", "") or "").strip() == code), None)
                    if prod is not None:
                        prod["qty"] = self._parse_float(prod.get("qty", 0), 0) + qty
                        prod["atualizado_em"] = self.desktop_main.now_iso()
                continue
            enc = self.get_encomenda_by_numero(str(line.get("encomenda", "") or "").strip())
            if enc is None:
                continue
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            for piece in self.desktop_main.encomenda_pecas(enc):
                if str(piece.get("id", "") or "").strip() == piece_id:
                    piece["qtd_expedida"] = max(0.0, self._parse_float(piece.get("qtd_expedida", 0), 0) - qty)
                    try:
                        self.desktop_main.atualizar_estado_peca(piece)
                    except Exception:
                        pass
                    break
            self.desktop_main.update_estado_expedicao_encomenda(enc)
        ex["anulada"] = True
        ex["estado"] = "Anulada"
        ex["anulada_motivo"] = motivo
        self._save(force=True)
        return self.expedicao_detail(numero)

    def _transport_defaults(self) -> dict[str, Any]:
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        origem = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        now_dt = datetime.now()
        return {
            "numero": "",
            "tipo_responsavel": "Nosso Cargo",
            "estado": "Planeado",
            "data_planeada": now_dt.date().isoformat(),
            "hora_saida": "08:00",
            "viatura": "",
            "matricula": "",
            "motorista": "",
            "telefone_motorista": "",
            "origem": origem,
            "paletes_total_manual": 0.0,
            "peso_total_manual_kg": 0.0,
            "volume_total_manual_m3": 0.0,
            "pedido_transporte_estado": "Nao pedido",
            "pedido_transporte_ref": "",
            "pedido_transporte_at": "",
            "pedido_transporte_by": "",
            "pedido_transporte_obs": "",
            "pedido_resposta_obs": "",
            "pedido_confirmado_at": "",
            "pedido_confirmado_by": "",
            "pedido_recusado_at": "",
            "pedido_recusado_by": "",
            "observacoes": "",
        }

    def _transport_note_for_order(self, enc: dict[str, Any]) -> str:
        note = str((enc or {}).get("nota_transporte", "") or "").strip()
        if note:
            return note
        quote_num = str((enc or {}).get("numero_orcamento", "") or "").strip()
        if not quote_num:
            return ""
        quote = self._billing_quote_by_number(quote_num)
        if not isinstance(quote, dict):
            return ""
        return str(quote.get("nota_transporte", "") or "").strip()

    def _transport_mode_for_order(self, enc: dict[str, Any]) -> str:
        note = str(self._transport_note_for_order(enc) or "").strip()
        note_norm = self.desktop_main.norm_text(note)
        if "subcontrat" in note_norm:
            return "Subcontratado"
        if "nosso cargo" in note_norm or ("transporte" in note_norm and "nosso" in note_norm):
            return "Transporte a Nosso Cargo"
        if "cliente" in note_norm or "vosso cargo" in note_norm:
            return "Transporte a Cargo do Cliente"
        return note

    def _transport_zone_for_order(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> str:
        enc = dict(enc or {})
        zone = str(enc.get("zona_transporte", "") or "").strip()
        if zone:
            return zone
        cliente_obj = dict(cliente_obj or {})
        if not cliente_obj:
            cli_code = str(enc.get("cliente", "") or "").strip()
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn) and cli_code:
                cliente_obj = find_cliente_fn(self.ensure_data(), cli_code) or {}
        for value in (
            cliente_obj.get("localidade", ""),
            cliente_obj.get("codigo_postal", ""),
        ):
            txt = str(value or "").strip()
            if txt:
                return txt
        return ""

    def transport_zone_options(self) -> list[str]:
        values: list[str] = []
        for row in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            txt = str((row or {}).get("zona", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("encomendas", []) or []):
            txt = self._transport_zone_for_order(row)
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("orcamentos", []) or []):
            txt = str((row or {}).get("zona_transporte", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("clientes", []) or []):
            txt = str((row or {}).get("localidade", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        values.sort(key=lambda item: self.desktop_main.norm_text(item))
        return values

    def transport_tariff_defaults(self) -> dict[str, Any]:
        return {
            "id": "",
            "transportadora_id": "",
            "transportadora_nome": "",
            "zona": "",
            "valor_base": 0.0,
            "valor_por_palete": 0.0,
            "valor_por_kg": 0.0,
            "valor_por_m3": 0.0,
            "custo_minimo": 0.0,
            "ativo": True,
            "observacoes": "",
        }

    def _transport_tariff_signature(self, row: dict[str, Any] | None) -> str:
        row = dict(row or {})
        carrier_parts = [
            part
            for part in [
                str(row.get("transportadora_id", "") or "").strip(),
                str(row.get("transportadora_nome", "") or "").strip(),
            ]
            if part
        ]
        carrier = " - ".join(carrier_parts).strip(" -") or "Sem transportadora"
        zone = str(row.get("zona", "") or "").strip() or "Sem zona"
        return f"{carrier} | {zone}"

    def _transport_tariff_next_id(self) -> int:
        highest = 0
        for row in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            highest = max(highest, int(self._parse_float((row or {}).get("id", 0), 0) or 0))
        return highest + 1

    def _transport_tariff_match(self, transportadora_id: Any = "", transportadora_nome: Any = "", zona: Any = "") -> dict[str, Any] | None:
        zone_norm = self.desktop_main.norm_text(zona)
        if not zone_norm:
            return None
        supplier_id = str(transportadora_id or "").strip()
        supplier_name_norm = self.desktop_main.norm_text(transportadora_nome)
        best: tuple[int, int, dict[str, Any]] | None = None
        for raw in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            if not isinstance(raw, dict) or not bool(raw.get("ativo", True)):
                continue
            if self.desktop_main.norm_text(raw.get("zona", "")) != zone_norm:
                continue
            row_supplier_id = str(raw.get("transportadora_id", "") or "").strip()
            row_supplier_name_norm = self.desktop_main.norm_text(raw.get("transportadora_nome", ""))
            score = 0
            if supplier_id and row_supplier_id and row_supplier_id == supplier_id:
                score = 3
            elif supplier_name_norm and row_supplier_name_norm and row_supplier_name_norm == supplier_name_norm:
                score = 2
            elif not row_supplier_id and not row_supplier_name_norm:
                score = 1
            if score <= 0:
                continue
            row_id = int(self._parse_float(raw.get("id", 0), 0) or 0)
            candidate = (score, -row_id, raw)
            if best is None or candidate > best:
                best = candidate
        return dict(best[2]) if best else None

    def _transport_tariff_cost_from_row(self, row: dict[str, Any] | None, paletes: Any = 0, peso_bruto_kg: Any = 0, volume_m3: Any = 0) -> float:
        row = dict(row or {})
        base = round(self._parse_float(row.get("valor_base", 0), 0), 2)
        per_pal = round(self._parse_float(row.get("valor_por_palete", 0), 0), 2)
        per_kg = round(self._parse_float(row.get("valor_por_kg", 0), 0), 4)
        per_m3 = round(self._parse_float(row.get("valor_por_m3", 0), 0), 2)
        minimum = round(self._parse_float(row.get("custo_minimo", 0), 0), 2)
        pal = max(0.0, round(self._parse_float(paletes, 0), 2))
        peso = max(0.0, round(self._parse_float(peso_bruto_kg, 0), 2))
        volume = max(0.0, round(self._parse_float(volume_m3, 0), 3))
        total = round(base + (pal * per_pal) + (peso * per_kg) + (volume * per_m3), 2)
        if minimum > 0:
            total = max(total, minimum)
        return round(total, 2)

    def _transport_tariff_suggestion(
        self,
        transportadora_id: Any = "",
        transportadora_nome: Any = "",
        zona: Any = "",
        paletes: Any = 0,
        peso_bruto_kg: Any = 0,
        volume_m3: Any = 0,
    ) -> dict[str, Any]:
        tariff = self._transport_tariff_match(transportadora_id, transportadora_nome, zona)
        if tariff is None:
            return {
                "tarifario_id": "",
                "tarifario_label": "",
                "custo_sugerido": 0.0,
            }
        return {
            "tarifario_id": tariff.get("id", ""),
            "tarifario_label": self._transport_tariff_signature(tariff),
            "custo_sugerido": self._transport_tariff_cost_from_row(tariff, paletes, peso_bruto_kg, volume_m3),
        }

    def _transport_metrics_for_order(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> dict[str, Any]:
        enc = dict(enc or {})
        supplier_id, supplier_text, supplier_contact = self._normalize_supplier_reference(
            enc.get("transportadora_id", ""),
            enc.get("transportadora_nome", ""),
        )
        return {
            "modo": self._transport_mode_for_order(enc),
            "paletes": round(self._parse_float(enc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(enc.get("volume_m3", 0), 0), 3),
            "preco_transporte": round(self._parse_float(enc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(enc.get("custo_transporte", 0), 0), 2),
            "transportadora_id": supplier_id,
            "transportadora_nome": supplier_text,
            "transportadora_contacto": supplier_contact,
            "referencia_transporte": str(enc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": self._transport_zone_for_order(enc, cliente_obj),
        }

    def _transport_stop_summary(self, stops: list[dict[str, Any]], trip: dict[str, Any] | None = None) -> dict[str, float]:
        paletes_calc = round(sum(self._parse_float(row.get("paletes", 0), 0) for row in list(stops or [])), 2)
        peso_calc = round(sum(self._parse_float(row.get("peso_bruto_kg", 0), 0) for row in list(stops or [])), 2)
        volume_calc = round(sum(self._parse_float(row.get("volume_m3", 0), 0) for row in list(stops or [])), 3)
        preco_total = round(sum(self._parse_float(row.get("preco_transporte", 0), 0) for row in list(stops or [])), 2)
        custo_total = round(sum(self._parse_float(row.get("custo_transporte", 0), 0) for row in list(stops or [])), 2)
        paletes = paletes_calc
        peso = peso_calc
        volume = volume_calc
        carga_manual = False
        if isinstance(trip, dict):
            custo_previsto = round(self._parse_float(trip.get("custo_previsto", 0), 0), 2)
            if custo_previsto > 0:
                custo_total = custo_previsto
            paletes_manual = round(self._parse_float(trip.get("paletes_total_manual", 0), 0), 2)
            peso_manual = round(self._parse_float(trip.get("peso_total_manual_kg", 0), 0), 2)
            volume_manual = round(self._parse_float(trip.get("volume_total_manual_m3", 0), 0), 3)
            if paletes_manual > 0:
                paletes = paletes_manual
                carga_manual = True
            if peso_manual > 0:
                peso = peso_manual
                carga_manual = True
            if volume_manual > 0:
                volume = volume_manual
                carga_manual = True
        return {
            "paletes": paletes,
            "peso_bruto_kg": peso,
            "volume_m3": volume,
            "paletes_calculadas": paletes_calc,
            "peso_bruto_kg_calculado": peso_calc,
            "volume_m3_calculado": volume_calc,
            "carga_manual": carga_manual,
            "preco_total": preco_total,
            "custo_total": custo_total,
            "margem_prevista": round(preco_total - custo_total, 2),
        }

    def _transport_is_own_cargo(self, enc: dict[str, Any]) -> bool:
        note_norm = self.desktop_main.norm_text(self._transport_note_for_order(enc))
        return (
            "nosso cargo" in note_norm
            or ("transporte" in note_norm and "nosso" in note_norm)
            or "subcontrat" in note_norm
        )

    def _transport_vehicle_options(self) -> list[str]:
        options: list[str] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            for value in (tr.get("viatura"), tr.get("matricula")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options

    def _transport_driver_options(self) -> list[str]:
        options: list[str] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            for value in (tr.get("motorista"), tr.get("telefone_motorista")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options

    def _transport_latest_guide_for_order(self, order_num: str) -> dict[str, Any] | None:
        target = str(order_num or "").strip()
        if not target:
            return None
        matches = [
            dict(ex)
            for ex in list(self.ensure_data().get("expedicoes", []) or [])
            if str((ex or {}).get("encomenda", "") or "").strip() == target and not bool((ex or {}).get("anulada"))
        ]
        if not matches:
            return None
        matches.sort(
            key=lambda row: (
                str(row.get("data_transporte", "") or row.get("data_emissao", "") or ""),
                str(row.get("numero", "") or ""),
            ),
            reverse=True,
        )
        return matches[0]

    def transport_guide_options(self, order_num: str) -> list[dict[str, str]]:
        target = str(order_num or "").strip()
        if not target:
            return []
        rows = [
            {
                "numero": str(ex.get("numero", "") or "").strip(),
                "data_emissao": str(ex.get("data_emissao", "") or "").strip(),
                "data_transporte": str(ex.get("data_transporte", "") or "").strip(),
                "estado": str(ex.get("estado", "") or "").strip(),
                "local_descarga": str(ex.get("local_descarga", "") or "").strip(),
                "label": " | ".join(
                    [
                        part
                        for part in [
                            str(ex.get("numero", "") or "").strip(),
                            str(ex.get("data_transporte", "") or ex.get("data_emissao", "") or "").strip(),
                            str(ex.get("estado", "") or "").strip(),
                        ]
                        if part
                    ]
                ),
            }
            for ex in list(self.ensure_data().get("expedicoes", []) or [])
            if str((ex or {}).get("encomenda", "") or "").strip() == target and not bool((ex or {}).get("anulada"))
        ]
        rows.sort(key=lambda row: (row.get("data_transporte") or row.get("data_emissao") or "", row.get("numero") or ""), reverse=True)
        return rows

    def _transport_find(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("transportes", []) or [])
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _transport_stop_state(self, stop: dict[str, Any], trip_state: str = "") -> str:
        state = str((stop or {}).get("estado", "") or "").strip()
        if state:
            return state
        trip_txt = str(trip_state or "").strip()
        return trip_txt or "Planeada"

    def _transport_stop_checklist_state(self, stop: dict[str, Any]) -> str:
        checks = [
            bool((stop or {}).get("check_carga_ok")),
            bool((stop or {}).get("check_docs_ok")),
            bool((stop or {}).get("check_paletes_ok")),
        ]
        if checks and all(checks):
            return "OK"
        if any(checks):
            return "Parcial"
        return "Pendente"

    def _transport_reindex_stops(self, trip: dict[str, Any]) -> None:
        stops = list(trip.get("paragens", []) or [])
        stops.sort(
            key=lambda row: (
                int(self._parse_float((row or {}).get("ordem", 0), 0) or 0),
                str((row or {}).get("data_planeada", "") or ""),
                str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or ""),
            )
        )
        for index, stop in enumerate(stops, start=1):
            stop["ordem"] = index
        trip["paragens"] = stops

    def _transport_sync_order_links(self) -> None:
        assigned: dict[str, tuple[str, str]] = {}
        for tr in list(self.ensure_data().get("transportes", []) or []):
            if not isinstance(tr, dict):
                continue
            trip_num = str(tr.get("numero", "") or "").strip()
            trip_state = str(tr.get("estado", "") or "").strip() or "Planeado"
            if not trip_num or "anulad" in self.desktop_main.norm_text(trip_state):
                continue
            for stop in list(tr.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
                if not enc_num:
                    continue
                assigned[enc_num] = (trip_num, self._transport_stop_state(stop, trip_state))
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            num = str(enc.get("numero", "") or "").strip()
            if not num:
                continue
            trip_info = assigned.get(num)
            if trip_info is None:
                enc["transporte_numero"] = ""
                enc["estado_transporte"] = ""
                continue
            enc["transporte_numero"] = trip_info[0]
            enc["estado_transporte"] = trip_info[1]

    def transport_defaults(self) -> dict[str, Any]:
        payload = dict(self._transport_defaults())
        payload["vehicle_options"] = self._transport_vehicle_options()
        payload["driver_options"] = self._transport_driver_options()
        payload["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.ne_suppliers() or [])]
        payload["zone_options"] = self.transport_zone_options()
        return payload

    def transport_tariff_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            if not isinstance(raw, dict):
                continue
            row = {
                "id": int(self._parse_float(raw.get("id", 0), 0) or 0),
                "transportadora_id": str(raw.get("transportadora_id", "") or "").strip(),
                "transportadora_nome": str(raw.get("transportadora_nome", "") or "").strip(),
                "zona": str(raw.get("zona", "") or "").strip(),
                "valor_base": round(self._parse_float(raw.get("valor_base", 0), 0), 2),
                "valor_por_palete": round(self._parse_float(raw.get("valor_por_palete", 0), 0), 2),
                "valor_por_kg": round(self._parse_float(raw.get("valor_por_kg", 0), 0), 4),
                "valor_por_m3": round(self._parse_float(raw.get("valor_por_m3", 0), 0), 2),
                "custo_minimo": round(self._parse_float(raw.get("custo_minimo", 0), 0), 2),
                "ativo": bool(raw.get("ativo", True)),
                "observacoes": str(raw.get("observacoes", "") or "").strip(),
                "label": self._transport_tariff_signature(raw),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (self.desktop_main.norm_text(item.get("transportadora_nome", "")), self.desktop_main.norm_text(item.get("zona", "")), item.get("id", 0)))
        return rows

    def transport_tariff_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("transportes_tarifarios", [])
        tariff_id = int(self._parse_float(payload.get("id", 0), 0) or 0)
        zone = str(payload.get("zona", "") or "").strip()
        if not zone:
            raise ValueError("Zona obrigatoria no tarifario.")
        transportadora_id, transportadora_nome, _contact = self._normalize_supplier_reference(
            payload.get("transportadora_id", ""),
            payload.get("transportadora_nome", ""),
        )
        zone_norm = self.desktop_main.norm_text(zone)
        for row in rows:
            if not isinstance(row, dict):
                continue
            if int(self._parse_float(row.get("id", 0), 0) or 0) == tariff_id:
                continue
            same_zone = self.desktop_main.norm_text(row.get("zona", "")) == zone_norm
            same_supplier = (
                str(row.get("transportadora_id", "") or "").strip() == transportadora_id
                and self.desktop_main.norm_text(row.get("transportadora_nome", "")) == self.desktop_main.norm_text(transportadora_nome)
            )
            if same_zone and same_supplier:
                raise ValueError("Ja existe um tarifario para essa transportadora e zona.")
        target = next((row for row in rows if int(self._parse_float((row or {}).get("id", 0), 0) or 0) == tariff_id), None) if tariff_id > 0 else None
        if target is None:
            target = self.transport_tariff_defaults()
            target["id"] = self._transport_tariff_next_id()
            rows.append(target)
        target["transportadora_id"] = transportadora_id
        target["transportadora_nome"] = transportadora_nome
        target["zona"] = zone
        target["valor_base"] = round(self._parse_float(payload.get("valor_base", target.get("valor_base", 0)), 0), 2)
        target["valor_por_palete"] = round(self._parse_float(payload.get("valor_por_palete", target.get("valor_por_palete", 0)), 0), 2)
        target["valor_por_kg"] = round(self._parse_float(payload.get("valor_por_kg", target.get("valor_por_kg", 0)), 0), 4)
        target["valor_por_m3"] = round(self._parse_float(payload.get("valor_por_m3", target.get("valor_por_m3", 0)), 0), 2)
        target["custo_minimo"] = round(self._parse_float(payload.get("custo_minimo", target.get("custo_minimo", 0)), 0), 2)
        target["ativo"] = bool(payload.get("ativo", target.get("ativo", True)))
        target["observacoes"] = str(payload.get("observacoes", target.get("observacoes", "")) or "").strip()
        self._save(force=True)
        return dict(target)

    def transport_tariff_remove(self, tariff_id: Any) -> None:
        target_id = int(self._parse_float(tariff_id, 0), 0)
        rows = list(self.ensure_data().get("transportes_tarifarios", []) or [])
        filtered = [row for row in rows if int(self._parse_float((row or {}).get("id", 0), 0) or 0) != target_id]
        if len(filtered) == len(rows):
            raise ValueError("Tarifario nao encontrado.")
        self.ensure_data()["transportes_tarifarios"] = filtered
        self._save(force=True)

    def transport_pending_orders(self, filter_text: str = "") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in list(data.get("transportes", []) or [])
            if isinstance(tr, dict) and "anulad" not in self.desktop_main.norm_text(str(tr.get("estado", "") or ""))
            for stop in list(tr.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        rows: list[dict[str, Any]] = []
        for enc in list(data.get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            enc_num = str(enc.get("numero", "") or "").strip()
            if not enc_num or not self._transport_is_own_cargo(enc):
                continue
            self.desktop_main.update_estado_expedicao_encomenda(enc)
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            disponivel = sum(max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(piece), 0)) for piece in pieces)
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            if disponivel <= 0 and not latest_guide:
                continue
            if active_assignments.get(enc_num):
                continue
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn):
                cli_obj = find_cliente_fn(data, cli_code) or {}
            cliente_txt = " - ".join([part for part in [cli_code, str(cli_obj.get("nome", "") or "").strip()] if part]).strip()
            metrics = self._transport_metrics_for_order(enc, cli_obj)
            suggestion = self._transport_tariff_suggestion(
                metrics.get("transportadora_id", ""),
                metrics.get("transportadora_nome", ""),
                metrics.get("zona_transporte", ""),
                metrics.get("paletes", 0.0),
                metrics.get("peso_bruto_kg", 0.0),
                metrics.get("volume_m3", 0.0),
            )
            row = {
                "numero": enc_num,
                "cliente": cliente_txt or cli_code or "-",
                "cliente_codigo": cli_code,
                "estado": str(enc.get("estado", "") or "").strip(),
                "estado_expedicao": str(enc.get("estado_expedicao", "Nao expedida") or "Nao expedida").strip(),
                "estado_transporte": str(enc.get("estado_transporte", "") or "").strip(),
                "nota_transporte": metrics.get("modo", "") or self._transport_note_for_order(enc),
                "preco_transporte": metrics.get("preco_transporte", 0.0),
                "custo_transporte": metrics.get("custo_transporte", 0.0),
                "paletes": metrics.get("paletes", 0.0),
                "peso_bruto_kg": metrics.get("peso_bruto_kg", 0.0),
                "volume_m3": metrics.get("volume_m3", 0.0),
                "transportadora_id": metrics.get("transportadora_id", ""),
                "transportadora_nome": metrics.get("transportadora_nome", ""),
                "referencia_transporte": metrics.get("referencia_transporte", ""),
                "zona_transporte": metrics.get("zona_transporte", ""),
                "local_descarga": str(enc.get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                "contacto": str(cli_obj.get("contacto", "") or "").strip(),
                "telefone": str(cli_obj.get("contacto", "") or "").strip(),
                "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                "guia_numero": str(latest_guide.get("numero", "") or "").strip(),
                "disponivel": round(disponivel, 1),
                "custo_sugerido": round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("data_entrega") or "9999-99-99", item.get("numero") or ""))
        return rows

    def transport_rows(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_filter = str(estado or "Todas").strip().lower()
        rows: list[dict[str, Any]] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            if not isinstance(tr, dict):
                continue
            trip_state = str(tr.get("estado", "") or "Planeado").strip()
            if state_filter not in ("todas", "todos", "all", "") and trip_state.lower() != state_filter:
                continue
            stops = list(tr.get("paragens", []) or [])
            delivered = sum(1 for stop in stops if "entreg" in self.desktop_main.norm_text(self._transport_stop_state(stop, trip_state)))
            summary = self._transport_stop_summary(stops, tr)
            row = {
                "numero": str(tr.get("numero", "") or "").strip(),
                "data_planeada": str(tr.get("data_planeada", "") or "").strip(),
                "hora_saida": str(tr.get("hora_saida", "") or "").strip(),
                "tipo_responsavel": str(tr.get("tipo_responsavel", "") or "Nosso Cargo").strip(),
                "estado": trip_state,
                "pedido_transporte_estado": str(tr.get("pedido_transporte_estado", "") or "Nao pedido").strip() or "Nao pedido",
                "transportadora_nome": str(tr.get("transportadora_nome", "") or "").strip(),
                "viatura": str(tr.get("viatura", "") or tr.get("matricula", "") or "").strip(),
                "motorista": str(tr.get("motorista", "") or "").strip(),
                "matricula": str(tr.get("matricula", "") or "").strip(),
                "paragens": len(stops),
                "entregues": delivered,
                "pendentes": max(0, len(stops) - delivered),
                "paletes": summary.get("paletes", 0.0),
                "peso_bruto_kg": summary.get("peso_bruto_kg", 0.0),
                "preco_total": summary.get("preco_total", 0.0),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("data_planeada") or "", item.get("hora_saida") or "", item.get("numero") or ""), reverse=True)
        return rows

    def transport_detail(self, numero: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        detail = {
            "numero": str(trip.get("numero", "") or "").strip(),
            "tipo_responsavel": str(trip.get("tipo_responsavel", "") or "Nosso Cargo").strip() or "Nosso Cargo",
            "estado": str(trip.get("estado", "") or "Planeado").strip() or "Planeado",
            "data_planeada": str(trip.get("data_planeada", "") or "").strip(),
            "hora_saida": str(trip.get("hora_saida", "") or "").strip(),
            "viatura": str(trip.get("viatura", "") or "").strip(),
            "matricula": str(trip.get("matricula", "") or "").strip(),
            "motorista": str(trip.get("motorista", "") or "").strip(),
            "telefone_motorista": str(trip.get("telefone_motorista", "") or "").strip(),
            "origem": str(trip.get("origem", "") or "").strip(),
            "transportadora_id": str(trip.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(trip.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(trip.get("referencia_transporte", "") or "").strip(),
            "custo_previsto": round(self._parse_float(trip.get("custo_previsto", 0), 0), 2),
            "paletes_total_manual": round(self._parse_float(trip.get("paletes_total_manual", 0), 0), 2),
            "peso_total_manual_kg": round(self._parse_float(trip.get("peso_total_manual_kg", 0), 0), 2),
            "volume_total_manual_m3": round(self._parse_float(trip.get("volume_total_manual_m3", 0), 0), 3),
            "pedido_transporte_estado": str(trip.get("pedido_transporte_estado", "") or "").strip() or "Nao pedido",
            "pedido_transporte_ref": str(trip.get("pedido_transporte_ref", "") or "").strip(),
            "pedido_transporte_at": str(trip.get("pedido_transporte_at", "") or "").strip(),
            "pedido_transporte_by": str(trip.get("pedido_transporte_by", "") or "").strip(),
            "pedido_transporte_obs": str(trip.get("pedido_transporte_obs", "") or "").strip(),
            "pedido_resposta_obs": str(trip.get("pedido_resposta_obs", "") or "").strip(),
            "pedido_confirmado_at": str(trip.get("pedido_confirmado_at", "") or "").strip(),
            "pedido_confirmado_by": str(trip.get("pedido_confirmado_by", "") or "").strip(),
            "pedido_recusado_at": str(trip.get("pedido_recusado_at", "") or "").strip(),
            "pedido_recusado_by": str(trip.get("pedido_recusado_by", "") or "").strip(),
            "observacoes": str(trip.get("observacoes", "") or "").strip(),
            "created_by": str(trip.get("created_by", "") or "").strip(),
            "created_at": str(trip.get("created_at", "") or "").strip(),
            "updated_at": str(trip.get("updated_at", "") or "").strip(),
            "paragens": [],
        }
        for stop in sorted(list(trip.get("paragens", []) or []), key=lambda row: int(self._parse_float((row or {}).get("ordem", 0), 0) or 0)):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            enc = self.get_encomenda_by_numero(enc_num) if enc_num else None
            cli_code = str(stop.get("cliente_codigo", "") or (enc or {}).get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn) and cli_code:
                cli_obj = find_cliente_fn(self.ensure_data(), cli_code) or {}
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            metrics = self._transport_metrics_for_order(enc or {}, cli_obj)
            supplier_id, supplier_text, supplier_contact = self._normalize_supplier_reference(
                stop.get("transportadora_id", "") or detail.get("transportadora_id", "") or metrics.get("transportadora_id", ""),
                stop.get("transportadora_nome", "") or detail.get("transportadora_nome", "") or metrics.get("transportadora_nome", ""),
            )
            zone_txt = str(stop.get("zona_transporte", "") or metrics.get("zona_transporte", "") or "").strip()
            paletes_value = round(self._parse_float(stop.get("paletes", metrics.get("paletes", 0)), 0), 2)
            peso_value = round(self._parse_float(stop.get("peso_bruto_kg", metrics.get("peso_bruto_kg", 0)), 0), 2)
            volume_value = round(self._parse_float(stop.get("volume_m3", metrics.get("volume_m3", 0)), 0), 3)
            suggestion = self._transport_tariff_suggestion(
                supplier_id,
                supplier_text,
                zone_txt,
                paletes_value,
                peso_value,
                volume_value,
            )
            detail["paragens"].append(
                {
                    "ordem": int(self._parse_float(stop.get("ordem", 0), 0) or 0),
                    "encomenda_numero": enc_num,
                    "cliente_codigo": cli_code,
                    "cliente_nome": str(stop.get("cliente_nome", "") or cli_obj.get("nome", "") or "").strip(),
                    "zona_transporte": zone_txt,
                    "local_descarga": str(stop.get("local_descarga", "") or (enc or {}).get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                    "contacto": str(stop.get("contacto", "") or cli_obj.get("contacto", "") or "").strip(),
                    "telefone": str(stop.get("telefone", "") or cli_obj.get("contacto", "") or "").strip(),
                    "data_planeada": str(stop.get("data_planeada", "") or "").replace("T", " ")[:19],
                    "paletes": paletes_value,
                    "peso_bruto_kg": peso_value,
                    "volume_m3": volume_value,
                    "preco_transporte": round(self._parse_float(stop.get("preco_transporte", metrics.get("preco_transporte", 0)), 0), 2),
                    "custo_transporte": round(self._parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_manual": round(self._parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_sugerido": round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
                    "tarifario_id": suggestion.get("tarifario_id", ""),
                    "tarifario_label": str(suggestion.get("tarifario_label", "") or "").strip(),
                    "transportadora_id": supplier_id,
                    "transportadora_nome": supplier_text,
                    "transportadora_contacto": supplier_contact,
                    "referencia_transporte": str(stop.get("referencia_transporte", "") or detail.get("referencia_transporte", "") or metrics.get("referencia_transporte", "") or "").strip(),
                    "nota_transporte": metrics.get("modo", "") or self._transport_note_for_order(enc or {}),
                    "estado": self._transport_stop_state(stop, detail["estado"]),
                    "check_carga_ok": bool(stop.get("check_carga_ok")),
                    "check_docs_ok": bool(stop.get("check_docs_ok")),
                    "check_paletes_ok": bool(stop.get("check_paletes_ok")),
                    "checklist_estado": self._transport_stop_checklist_state(stop),
                    "pod_estado": str(stop.get("pod_estado", "") or "").strip(),
                    "pod_recebido_nome": str(stop.get("pod_recebido_nome", "") or "").strip(),
                    "pod_recebido_at": str(stop.get("pod_recebido_at", "") or "").replace("T", " ")[:19],
                    "pod_obs": str(stop.get("pod_obs", "") or "").strip(),
                    "observacoes": str(stop.get("observacoes", "") or "").strip(),
                    "guia_numero": str(stop.get("expedicao_numero", "") or latest_guide.get("numero", "") or "").strip(),
                    "estado_expedicao": str((enc or {}).get("estado_expedicao", "") or "").strip(),
                }
            )
        detail.update(self._transport_stop_summary(list(detail.get("paragens", []) or []), detail))
        detail["custo_sugerido_total"] = round(
            sum(self._parse_float(stop.get("custo_sugerido", 0), 0) for stop in list(detail.get("paragens", []) or [])),
            2,
        )
        detail["checklist_ok"] = sum(1 for stop in list(detail.get("paragens", []) or []) if str(stop.get("checklist_estado", "") or "") == "OK")
        detail["pod_recebidos"] = sum(1 for stop in list(detail.get("paragens", []) or []) if "recebid" in self.desktop_main.norm_text(str(stop.get("pod_estado", "") or "")))
        zones = []
        for stop in list(detail.get("paragens", []) or []):
            zone_txt = str(stop.get("zona_transporte", "") or "").strip()
            if zone_txt and zone_txt not in zones:
                zones.append(zone_txt)
        detail["zonas"] = zones
        detail["vehicle_options"] = self._transport_vehicle_options()
        detail["driver_options"] = self._transport_driver_options()
        detail["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.ne_suppliers() or [])]
        detail["zone_options"] = self.transport_zone_options()
        return detail

    def transport_create_or_update(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip()
        trip = self._transport_find(numero) if numero else None
        if trip is None:
            numero = str(numero or self.desktop_main.next_transporte_numero(data)).strip()
            trip = {
                "numero": numero,
                "paragens": [],
                "created_by": str((self.user or {}).get("username", "") or "").strip(),
                "created_at": self.desktop_main.now_iso(),
            }
            data.setdefault("transportes", []).append(trip)
            self.desktop_main.reserve_transporte_numero(data, numero)
        defaults = self._transport_defaults()
        supplier_id, supplier_text, _supplier_contact = self._normalize_supplier_reference(
            payload.get("transportadora_id", trip.get("transportadora_id", "")),
            payload.get("transportadora_nome", trip.get("transportadora_nome", "")),
        )
        trip["tipo_responsavel"] = str(payload.get("tipo_responsavel", trip.get("tipo_responsavel", defaults["tipo_responsavel"])) or defaults["tipo_responsavel"]).strip() or defaults["tipo_responsavel"]
        trip["estado"] = str(payload.get("estado", trip.get("estado", defaults["estado"])) or defaults["estado"]).strip() or defaults["estado"]
        trip["data_planeada"] = str(payload.get("data_planeada", trip.get("data_planeada", defaults["data_planeada"])) or defaults["data_planeada"]).strip()
        trip["hora_saida"] = str(payload.get("hora_saida", trip.get("hora_saida", defaults["hora_saida"])) or defaults["hora_saida"]).strip()
        trip["viatura"] = str(payload.get("viatura", trip.get("viatura", "")) or "").strip()
        trip["matricula"] = str(payload.get("matricula", trip.get("matricula", "")) or "").strip()
        trip["motorista"] = str(payload.get("motorista", trip.get("motorista", "")) or "").strip()
        trip["telefone_motorista"] = str(payload.get("telefone_motorista", trip.get("telefone_motorista", "")) or "").strip()
        trip["origem"] = str(payload.get("origem", trip.get("origem", defaults["origem"])) or defaults["origem"]).strip()
        trip["transportadora_id"] = supplier_id
        trip["transportadora_nome"] = supplier_text
        trip["referencia_transporte"] = str(payload.get("referencia_transporte", trip.get("referencia_transporte", "")) or "").strip()
        trip["custo_previsto"] = round(self._parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        trip["paletes_total_manual"] = round(self._parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self._parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self._parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
        trip["pedido_transporte_estado"] = str(payload.get("pedido_transporte_estado", trip.get("pedido_transporte_estado", "Nao pedido")) or "Nao pedido").strip() or "Nao pedido"
        trip["pedido_transporte_ref"] = str(payload.get("pedido_transporte_ref", trip.get("pedido_transporte_ref", "")) or "").strip()
        trip["pedido_transporte_at"] = str(payload.get("pedido_transporte_at", trip.get("pedido_transporte_at", "")) or "").strip()
        trip["pedido_transporte_by"] = str(payload.get("pedido_transporte_by", trip.get("pedido_transporte_by", "")) or "").strip()
        trip["pedido_transporte_obs"] = str(payload.get("pedido_transporte_obs", trip.get("pedido_transporte_obs", "")) or "").strip()
        trip["pedido_resposta_obs"] = str(payload.get("pedido_resposta_obs", trip.get("pedido_resposta_obs", "")) or "").strip()
        trip["pedido_confirmado_at"] = str(payload.get("pedido_confirmado_at", trip.get("pedido_confirmado_at", "")) or "").strip()
        trip["pedido_confirmado_by"] = str(payload.get("pedido_confirmado_by", trip.get("pedido_confirmado_by", "")) or "").strip()
        trip["pedido_recusado_at"] = str(payload.get("pedido_recusado_at", trip.get("pedido_recusado_at", "")) or "").strip()
        trip["pedido_recusado_by"] = str(payload.get("pedido_recusado_by", trip.get("pedido_recusado_by", "")) or "").strip()
        if "subcontrat" in self.desktop_main.norm_text(trip["tipo_responsavel"]) and not trip["transportadora_nome"]:
            raise ValueError("Seleciona a transportadora externa para viagens subcontratadas.")
        trip["observacoes"] = str(payload.get("observacoes", trip.get("observacoes", "")) or "").strip()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_request_service(self, numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        payload = dict(payload or {})
        supplier_id, supplier_text, _supplier_contact = self._normalize_supplier_reference(
            payload.get("transportadora_id", trip.get("transportadora_id", "")),
            payload.get("transportadora_nome", trip.get("transportadora_nome", "")),
        )
        if "transportadora_id" in payload or "transportadora_nome" in payload:
            trip["transportadora_id"] = supplier_id
            trip["transportadora_nome"] = supplier_text
        trip["paletes_total_manual"] = round(self._parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self._parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self._parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
        trip["custo_previsto"] = round(self._parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        request_state = str(payload.get("pedido_transporte_estado", trip.get("pedido_transporte_estado", "Pedido enviado")) or "Pedido enviado").strip() or "Pedido enviado"
        trip["pedido_transporte_estado"] = request_state
        trip["pedido_transporte_ref"] = str(payload.get("pedido_transporte_ref", trip.get("pedido_transporte_ref", "")) or "").strip()
        trip["pedido_transporte_obs"] = str(payload.get("pedido_transporte_obs", trip.get("pedido_transporte_obs", "")) or "").strip()
        trip["pedido_resposta_obs"] = str(payload.get("pedido_resposta_obs", trip.get("pedido_resposta_obs", "")) or "").strip()
        normalized_state = self.desktop_main.norm_text(request_state)
        if normalized_state in {"nao pedido", "nao-pedido"}:
            trip["pedido_transporte_at"] = ""
            trip["pedido_transporte_by"] = ""
            trip["pedido_confirmado_at"] = ""
            trip["pedido_confirmado_by"] = ""
            trip["pedido_recusado_at"] = ""
            trip["pedido_recusado_by"] = ""
        else:
            trip["pedido_transporte_at"] = self.desktop_main.now_iso()
            trip["pedido_transporte_by"] = str((self.user or {}).get("username", "") or "").strip()
            if "confirm" in normalized_state:
                trip["pedido_confirmado_at"] = self.desktop_main.now_iso()
                trip["pedido_confirmado_by"] = str((self.user or {}).get("username", "") or "").strip()
                trip["pedido_recusado_at"] = ""
                trip["pedido_recusado_by"] = ""
            elif "recus" in normalized_state:
                trip["pedido_recusado_at"] = self.desktop_main.now_iso()
                trip["pedido_recusado_by"] = str((self.user or {}).get("username", "") or "").strip()
                trip["pedido_confirmado_at"] = ""
                trip["pedido_confirmado_by"] = ""
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_remove_trip(self, numero: str) -> None:
        trip_num = str(numero or "").strip()
        if not trip_num:
            raise ValueError("Seleciona uma viagem.")
        data = self.ensure_data()
        trips = [row for row in list(data.get("transportes", []) or []) if isinstance(row, dict)]
        target = next((row for row in trips if str(row.get("numero", "") or "").strip() == trip_num), None)
        if target is None:
            raise ValueError("Transporte nao encontrado.")
        data["transportes"] = [row for row in trips if row is not target]
        self._transport_sync_order_links()
        self._save(force=True)

    def transport_update_stop(self, numero: str, encomenda_numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        enc_num = str(encomenda_numero or "").strip()
        target = next(
            (
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            None,
        )
        if target is None:
            raise ValueError("Paragem nao encontrada.")
        payload = dict(payload or {})
        guide_number = str(payload.get("expedicao_numero", target.get("expedicao_numero", "")) or "").strip()
        if guide_number:
            valid_guides = {row.get("numero", "") for row in self.transport_guide_options(enc_num)}
            if valid_guides and guide_number not in valid_guides:
                raise ValueError("A guia escolhida nao pertence a esta encomenda.")
        target["expedicao_numero"] = guide_number
        target["zona_transporte"] = str(payload.get("zona_transporte", target.get("zona_transporte", "")) or "").strip()
        target["local_descarga"] = str(payload.get("local_descarga", target.get("local_descarga", "")) or "").strip()
        target["contacto"] = str(payload.get("contacto", target.get("contacto", "")) or "").strip()
        target["telefone"] = str(payload.get("telefone", target.get("telefone", "")) or "").strip()
        target["data_planeada"] = str(payload.get("data_planeada", target.get("data_planeada", "")) or "").strip()
        if "check_carga_ok" in payload:
            target["check_carga_ok"] = bool(payload.get("check_carga_ok"))
        if "check_docs_ok" in payload:
            target["check_docs_ok"] = bool(payload.get("check_docs_ok"))
        if "check_paletes_ok" in payload:
            target["check_paletes_ok"] = bool(payload.get("check_paletes_ok"))
        if "pod_estado" in payload:
            target["pod_estado"] = str(payload.get("pod_estado", "") or "").strip()
        if "pod_recebido_nome" in payload:
            target["pod_recebido_nome"] = str(payload.get("pod_recebido_nome", "") or "").strip()
        if "pod_recebido_at" in payload:
            target["pod_recebido_at"] = str(payload.get("pod_recebido_at", "") or "").strip()
        if "pod_obs" in payload:
            target["pod_obs"] = str(payload.get("pod_obs", "") or "").strip()
        if "observacoes" in payload:
            target["observacoes"] = str(payload.get("observacoes", "") or "").strip()
        if (
            "recebid" in self.desktop_main.norm_text(str(target.get("pod_estado", "") or ""))
            and not str(target.get("pod_recebido_at", "") or "").strip()
        ):
            target["pod_recebido_at"] = self.desktop_main.now_iso()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_assign_orders(self, numero: str, order_numbers: list[str]) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        trip_state = str(trip.get("estado", "") or "Planeado").strip()
        if "conclu" in self.desktop_main.norm_text(trip_state) or "anulad" in self.desktop_main.norm_text(trip_state):
            raise ValueError("Nao podes alterar uma viagem concluida ou anulada.")
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in list(self.ensure_data().get("transportes", []) or [])
            if isinstance(tr, dict) and "anulad" not in self.desktop_main.norm_text(str(tr.get("estado", "") or ""))
            for stop in list(tr.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        existing_orders = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            for stop in list(trip.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        order_list = []
        for raw in list(order_numbers or []):
            num = str(raw or "").strip()
            if num and num not in order_list:
                order_list.append(num)
        if not order_list:
            raise ValueError("Seleciona pelo menos uma encomenda.")
        stop_dt = ""
        if str(trip.get("data_planeada", "") or "").strip():
            stop_dt = str(trip.get("data_planeada", "")).strip()
            if str(trip.get("hora_saida", "") or "").strip():
                stop_dt = f"{stop_dt}T{str(trip.get('hora_saida', '')).strip()}:00"
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        for enc_num in order_list:
            if enc_num in existing_orders:
                continue
            assigned_trip = active_assignments.get(enc_num)
            if assigned_trip and assigned_trip != trip.get("numero"):
                raise ValueError(f"A encomenda {enc_num} ja esta afeta ao transporte {assigned_trip}.")
            enc = self.get_encomenda_by_numero(enc_num)
            if enc is None:
                raise ValueError(f"Encomenda nao encontrada: {enc_num}")
            if not self._transport_is_own_cargo(enc):
                raise ValueError(f"A encomenda {enc_num} nao esta definida como transporte a nosso cargo.")
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = find_cliente_fn(self.ensure_data(), cli_code) if callable(find_cliente_fn) and cli_code else {}
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            metrics = self._transport_metrics_for_order(enc, cli_obj)
            carrier_id = str(trip.get("transportadora_id", "") or metrics.get("transportadora_id", "") or "").strip()
            carrier_name = str(trip.get("transportadora_nome", "") or metrics.get("transportadora_nome", "") or "").strip()
            zone_txt = str(metrics.get("zona_transporte", "") or "").strip()
            suggestion = self._transport_tariff_suggestion(
                carrier_id,
                carrier_name,
                zone_txt,
                metrics.get("paletes", 0.0),
                metrics.get("peso_bruto_kg", 0.0),
                metrics.get("volume_m3", 0.0),
            )
            order_cost = round(self._parse_float(metrics.get("custo_transporte", 0), 0), 2)
            suggested_cost = round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2)
            trip.setdefault("paragens", []).append(
                {
                    "ordem": len(list(trip.get("paragens", []) or [])) + 1,
                    "encomenda_numero": enc_num,
                    "expedicao_numero": str(latest_guide.get("numero", "") or "").strip(),
                    "cliente_codigo": cli_code,
                    "cliente_nome": str(cli_obj.get("nome", "") or "").strip(),
                    "zona_transporte": zone_txt,
                    "local_descarga": str(enc.get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                    "contacto": str(cli_obj.get("contacto", "") or "").strip(),
                    "telefone": str(cli_obj.get("contacto", "") or "").strip(),
                    "data_planeada": stop_dt,
                    "paletes": metrics.get("paletes", 0.0),
                    "peso_bruto_kg": metrics.get("peso_bruto_kg", 0.0),
                    "volume_m3": metrics.get("volume_m3", 0.0),
                    "preco_transporte": metrics.get("preco_transporte", 0.0),
                    "custo_transporte": order_cost if order_cost > 0 else suggested_cost,
                    "transportadora_id": carrier_id,
                    "transportadora_nome": carrier_name,
                    "referencia_transporte": str(metrics.get("referencia_transporte", "") or trip.get("referencia_transporte", "") or "").strip(),
                    "check_carga_ok": False,
                    "check_docs_ok": False,
                    "check_paletes_ok": False,
                    "pod_estado": "",
                    "pod_recebido_nome": "",
                    "pod_recebido_at": "",
                    "pod_obs": "",
                    "estado": "Planeada",
                    "observacoes": "",
                }
            )
        self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_apply_suggested_cost(self, numero: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        detail = self.transport_detail(numero)
        suggested_map = {
            str(stop.get("encomenda_numero", "") or "").strip(): round(self._parse_float(stop.get("custo_sugerido", 0), 0), 2)
            for stop in list(detail.get("paragens", []) or [])
        }
        total = 0.0
        applied = 0
        for stop in list(trip.get("paragens", []) or []):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            suggested = round(self._parse_float(suggested_map.get(enc_num, 0), 0), 2)
            if suggested <= 0:
                continue
            stop["custo_transporte"] = suggested
            total += suggested
            applied += 1
        if applied <= 0:
            raise ValueError("Sem custos sugeridos para aplicar nesta viagem.")
        trip["custo_previsto"] = round(total, 2)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_remove_stop(self, numero: str, encomenda_numero: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        enc_num = str(encomenda_numero or "").strip()
        before = len(list(trip.get("paragens", []) or []))
        trip["paragens"] = [
            row
            for row in list(trip.get("paragens", []) or [])
            if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() != enc_num
        ]
        if len(list(trip.get("paragens", []) or [])) == before:
            raise ValueError("Paragem nao encontrada.")
        self._transport_reindex_stops(trip)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_move_stop(self, numero: str, encomenda_numero: str, direction: int) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        enc_num = str(encomenda_numero or "").strip()
        stops = list(trip.get("paragens", []) or [])
        index = next(
            (
                idx
                for idx, row in enumerate(stops)
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            -1,
        )
        if index < 0:
            raise ValueError("Paragem nao encontrada.")
        target = index + (1 if int(direction or 0) > 0 else -1)
        if target < 0 or target >= len(stops):
            return self.transport_detail(numero)
        stops[index], stops[target] = stops[target], stops[index]
        trip["paragens"] = stops
        self._transport_reindex_stops(trip)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_set_status(self, numero: str, estado: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        state_txt = str(estado or "").strip()
        if not state_txt:
            raise ValueError("Estado obrigatorio.")
        trip["estado"] = state_txt
        if "conclu" in self.desktop_main.norm_text(state_txt):
            for stop in list(trip.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                if "inciden" in self.desktop_main.norm_text(str(stop.get("estado", "") or "")):
                    continue
                stop["estado"] = "Entregue"
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_set_stop_status(self, numero: str, encomenda_numero: str, estado: str, observacoes: str = "") -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        enc_num = str(encomenda_numero or "").strip()
        target = next(
            (
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            None,
        )
        if target is None:
            raise ValueError("Paragem nao encontrada.")
        target["estado"] = str(estado or "").strip() or "Planeada"
        if str(observacoes or "").strip():
            target["observacoes"] = str(observacoes or "").strip()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_route_sheet_render(self, numero: str, path: str | Path) -> Path:
        detail = self.transport_detail(numero)
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A4
        margin = 34
        row_h = 22
        c = canvas.Canvas(str(out_path), pagesize=A4)

        def draw_header() -> float:
            c.setTitle(f"Folha de rota {detail.get('numero', '')}")
            c.setFont("Helvetica-Bold", 20)
            c.setFillColor(colors.HexColor("#0f172a"))
            c.drawString(margin, page_h - 44, "Transportes | Folha de rota")
            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin, page_h - 64, f"Viagem {detail.get('numero', '-')}")
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.HexColor("#475569"))
            meta = [
                f"Data {detail.get('data_planeada', '-') or '-'}",
                f"Saida {detail.get('hora_saida', '-') or '-'}",
                f"Tipo {detail.get('tipo_responsavel', '-') or '-'}",
                f"Estado {detail.get('estado', '-') or '-'}",
                f"Viatura {detail.get('viatura', '-') or '-'}",
                f"Motorista {detail.get('motorista', '-') or '-'}",
            ]
            c.drawString(margin, page_h - 80, " | ".join(meta))
            carrier_txt = str(detail.get("transportadora_nome", "") or "-").strip() or "-"
            c.drawString(
                margin,
                page_h - 94,
                f"Origem {detail.get('origem', '-') or '-'} | Transportadora {carrier_txt} | Ref {detail.get('referencia_transporte', '-') or '-'}",
            )
            c.drawString(
                margin,
                page_h - 108,
                f"Totais {detail.get('paletes', 0):.2f} pal | {detail.get('peso_bruto_kg', 0):.1f} kg | "
                f"{detail.get('volume_m3', 0):.3f} m3 | Preco {self._fmt_eur(detail.get('preco_total', 0))} | "
                f"Custo {self._fmt_eur(detail.get('custo_total', 0))} | Sug. {self._fmt_eur(detail.get('custo_sugerido_total', 0))}",
            )
            c.drawString(
                margin,
                page_h - 122,
                f"Pedido transporte {detail.get('pedido_transporte_estado', 'Nao pedido') or 'Nao pedido'} | "
                f"Ref pedido {detail.get('pedido_transporte_ref', '-') or '-'}",
            )
            response_parts = []
            if detail.get("pedido_confirmado_at"):
                response_parts.append(f"Confirmado {detail.get('pedido_confirmado_at', '-')}")
            if detail.get("pedido_recusado_at"):
                response_parts.append(f"Recusado {detail.get('pedido_recusado_at', '-')}")
            if detail.get("pedido_resposta_obs"):
                response_parts.append(f"Resposta {detail.get('pedido_resposta_obs', '-')}")
            if response_parts:
                c.drawString(margin, page_h - 136, " | ".join(response_parts))
                line_y = page_h - 146
            else:
                line_y = page_h - 132
            c.setStrokeColor(colors.HexColor("#cbd5e1"))
            c.line(margin, line_y, page_w - margin, line_y)
            return line_y - 18

        def draw_table_header(y: float) -> float:
            c.setFillColor(colors.HexColor("#0f172a"))
            c.roundRect(margin, y - row_h + 4, page_w - (margin * 2), row_h, 8, fill=1, stroke=0)
            cols = [("Ord", 34), ("Encomenda", 84), ("Cliente", 120), ("Descarga", 168), ("Planeado", 82), ("Guia", 64), ("Estado", 74)]
            x = margin + 8
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            for label, width in cols:
                c.drawString(x, y - 10, label)
                x += width
            return y - row_h - 2

        def new_page() -> float:
            c.showPage()
            return draw_header()

        y = draw_header()
        y = draw_table_header(y)
        widths = [34, 84, 120, 168, 82, 64, 74]
        for stop in list(detail.get("paragens", []) or []):
            metrics_line = (
                f"Pal {self._fmt(stop.get('paletes', 0))} | "
                f"{self._fmt(stop.get('peso_bruto_kg', 0))} kg | "
                f"{self._fmt(stop.get('volume_m3', 0))} m3 | "
                f"Preco {self._fmt_eur(stop.get('preco_transporte', 0))} | "
                f"Custo {self._fmt_eur(stop.get('custo_transporte', 0))} | Sug. {self._fmt_eur(stop.get('custo_sugerido', 0))}"
            )
            carrier_line = ""
            if stop.get("transportadora_nome"):
                carrier_line = f"Transportadora: {stop.get('transportadora_nome', '-')}"
                if stop.get("referencia_transporte"):
                    carrier_line += f" | Ref: {stop.get('referencia_transporte', '-')}"
            zone_line = ""
            if stop.get("zona_transporte"):
                zone_line = f"Zona: {stop.get('zona_transporte', '-')}"
                if stop.get("tarifario_label"):
                    zone_line += f" | Tarifario: {stop.get('tarifario_label', '-')}"
            checklist_line = (
                f"Checklist: carga {'OK' if stop.get('check_carga_ok') else '-'} / "
                f"docs {'OK' if stop.get('check_docs_ok') else '-'} / "
                f"paletes {'OK' if stop.get('check_paletes_ok') else '-'}"
            )
            pod_line = ""
            if stop.get("pod_estado"):
                pod_line = f"POD: {stop.get('pod_estado', '-')}"
                if stop.get("pod_recebido_nome"):
                    pod_line += f" por {stop.get('pod_recebido_nome', '-')}"
            combined_note = " | ".join(
                [
                    part
                    for part in [
                        metrics_line,
                        carrier_line,
                        zone_line,
                        checklist_line,
                        pod_line,
                        str(stop.get("pod_obs", "") or "").strip(),
                        str(stop.get("observacoes", "") or "").strip(),
                    ]
                    if part
                ]
            )
            extra_lines = _pdf_wrap_text(combined_note, "Helvetica", 7.0, page_w - (margin * 2) - 16, max_lines=3)
            needed = row_h + (8 * len(extra_lines)) + 8
            if y < margin + needed:
                y = new_page()
                y = draw_table_header(y)
            c.setFillColor(colors.HexColor("#f8fafc"))
            c.roundRect(margin, y - row_h + 4, page_w - (margin * 2), row_h, 6, fill=1, stroke=0)
            values = [
                str(stop.get("ordem", "-") or "-"),
                _pdf_clip_text(stop.get("encomenda_numero", "-"), widths[1] - 6, "Helvetica-Bold", 7.6),
                _pdf_clip_text(stop.get("cliente_nome", "-"), widths[2] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("local_descarga", "-"), widths[3] - 6, "Helvetica", 7.2),
                _pdf_clip_text(str(stop.get("data_planeada", "") or detail.get("data_planeada", "-")).replace("T", " ")[:16] or "-", widths[4] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("guia_numero", "-"), widths[5] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("estado", "-"), widths[6] - 6, "Helvetica-Bold", 7.4),
            ]
            x = margin + 8
            c.setFillColor(colors.HexColor("#0f172a"))
            for index, value in enumerate(values):
                c.setFont("Helvetica-Bold" if index in (0, 1, 6) else "Helvetica", 7.4)
                c.drawString(x, y - 10, str(value or "-"))
                x += widths[index]
            if extra_lines:
                c.setFillColor(colors.HexColor("#64748b"))
                c.setFont("Helvetica", 7.0)
                text_y = y - 19
                for line in extra_lines:
                    c.drawString(margin + 12, text_y, line)
                    text_y -= 8
                y = text_y - 6
            else:
                y -= row_h + 4
        if y < 110:
            y = new_page()
        c.setStrokeColor(colors.HexColor("#cbd5e1"))
        c.line(margin, 92, page_w - margin, 92)
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#475569"))
        c.drawString(margin, 76, "Observacao: esta folha de rota apoia a distribuicao e nao substitui a guia/documento de transporte.")
        c.drawString(margin, 58, "Motorista: ____________________________")
        c.drawString(margin + 220, 58, "Saida: ____________")
        c.drawString(margin + 360, 58, "Chegada: ____________")
        c.save()
        return out_path

    def transport_route_sheet_open(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_transporte_{str(numero or '').strip()}.pdf"
        self.transport_route_sheet_render(numero, target)
        os.startfile(str(target))
        return target

    def _billing_records(self) -> list[dict[str, Any]]:
        return list(self.ensure_data().get("faturacao", []) or [])

    def _billing_find_record(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in self._billing_records()
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _billing_find_source_record(self, orcamento_numero: str = "", encomenda_numero: str = "") -> dict[str, Any] | None:
        orc_num = str(orcamento_numero or "").strip()
        enc_num = str(encomenda_numero or "").strip()
        for row in self._billing_records():
            if not isinstance(row, dict):
                continue
            row_orc = str(row.get("orcamento_numero", "") or "").strip()
            row_enc = str(row.get("encomenda_numero", "") or "").strip()
            if orc_num and row_orc == orc_num:
                return row
            if enc_num and row_enc == enc_num:
                return row
        return None

    def _billing_next_number(self) -> str:
        year = str(self.desktop_main.datetime.now().year)
        highest = 0
        for row in self._billing_records():
            raw = str((row or {}).get("numero", "") or "").strip().upper()
            if not raw:
                continue
            digits = "".join(ch for ch in raw if ch.isdigit())
            if len(digits) >= 8 and digits.startswith(year):
                try:
                    highest = max(highest, int(digits[-4:]))
                    continue
                except Exception:
                    pass
            if digits:
                try:
                    highest = max(highest, int(digits[-4:]))
                except Exception:
                    continue
        return f"FAT-{year}-{highest + 1:04d}"

    def _billing_quote_by_number(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("orcamentos", []) or [])
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _billing_order_by_number(self, numero: str) -> dict[str, Any] | None:
        return self.get_encomenda_by_numero(str(numero or "").strip())

    def _billing_quote_is_sold(self, quote: dict[str, Any]) -> bool:
        estado = self.desktop_main.norm_text((quote or {}).get("estado", ""))
        return bool(
            "aprov" in estado
            or "convert" in estado
            or str((quote or {}).get("numero_encomenda", "") or "").strip()
        )

    def _billing_client_info(
        self,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, str]:
        if isinstance(quote, dict):
            client = self._normalize_orc_client(quote.get("cliente", {}))
            code = str(client.get("codigo", "") or "").strip()
            name = str(client.get("nome", "") or client.get("empresa", "") or "").strip()
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        if isinstance(order, dict):
            code = str(order.get("cliente", "") or "").strip()
            name = ""
            if code:
                try:
                    name = str((self.desktop_main.find_cliente(self.ensure_data(), code) or {}).get("nome", "") or "").strip()
                except Exception:
                    name = ""
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        if isinstance(record, dict):
            code = str(record.get("cliente_codigo", "") or "").strip()
            name = str(record.get("cliente_nome", "") or "").strip()
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        return {"codigo": "", "nome": "", "label": "-"}

    def _billing_guides_for_order(self, encomenda_numero: str) -> list[dict[str, Any]]:
        enc_num = str(encomenda_numero or "").strip()
        if not enc_num:
            return []
        rows: list[dict[str, Any]] = []
        for row in list(self.ensure_data().get("expedicoes", []) or []):
            if not isinstance(row, dict):
                continue
            if bool(row.get("anulada")):
                continue
            order_number = str(row.get("encomenda", "") or row.get("encomenda_numero", "") or "").strip()
            if order_number != enc_num:
                continue
            rows.append(
                {
                    "numero": str(row.get("numero", "") or "").strip(),
                    "data_emissao": str(row.get("data_emissao", "") or "").replace("T", " ")[:19],
                    "destinatario": str(row.get("destinatario", "") or "").strip(),
                }
            )
        rows.sort(key=lambda item: str(item.get("data_emissao", "") or ""), reverse=True)
        return rows

    def _billing_default_serie_id(self, issue_date: str = "") -> str:
        raw_issue_date = str(issue_date or self.desktop_main.now_iso()).strip() or self.desktop_main.now_iso()
        default_fn = getattr(self.desktop_main, "_exp_default_serie_id", None)
        if callable(default_fn):
            try:
                return str(default_fn("FT", raw_issue_date) or "").strip() or f"FT{str(raw_issue_date)[:4]}"
            except Exception:
                pass
        return f"FT{str(raw_issue_date)[:4]}"

    def _billing_due_days_from_text(self, value: str) -> int:
        digits = "".join(ch if ch.isdigit() else " " for ch in str(value or ""))
        values = [chunk for chunk in digits.split() if chunk.isdigit()]
        if not values:
            return 30
        try:
            return max(0, min(365, int(values[0])))
        except Exception:
            return 30

    def _billing_client_snapshot(
        self,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, str]:
        base = {
            "codigo": "",
            "nome": "",
            "nif": "",
            "morada": "",
            "contacto": "",
            "email": "",
            "cond_pagamento": "",
        }
        if isinstance(quote, dict):
            qclient = dict(self._normalize_orc_client(quote.get("cliente", {})) or {})
            base.update(
                {
                    "codigo": str(qclient.get("codigo", "") or "").strip(),
                    "nome": str(qclient.get("nome", "") or qclient.get("empresa", "") or "").strip(),
                    "nif": str(qclient.get("nif", "") or "").strip(),
                    "morada": str(qclient.get("morada", "") or "").strip(),
                    "contacto": str(qclient.get("contacto", "") or "").strip(),
                    "email": str(qclient.get("email", "") or "").strip(),
                }
            )
        if isinstance(order, dict) and not base.get("codigo"):
            base["codigo"] = str(order.get("cliente", "") or "").strip()
        if isinstance(record, dict):
            if not base.get("codigo"):
                base["codigo"] = str(record.get("cliente_codigo", "") or "").strip()
            if not base.get("nome"):
                base["nome"] = str(record.get("cliente_nome", "") or "").strip()

        client_ref = None
        client_code = str(base.get("codigo", "") or "").strip()
        for row in list(self.ensure_data().get("clientes", []) or []):
            if not isinstance(row, dict):
                continue
            row_code = str(row.get("codigo", "") or "").strip()
            if client_code and row_code == client_code:
                client_ref = row
                break
        if client_ref is None:
            for row in list(self.ensure_data().get("clientes", []) or []):
                if not isinstance(row, dict):
                    continue
                if base.get("nif") and str(row.get("nif", "") or "").strip() == base["nif"]:
                    client_ref = row
                    break
                if base.get("nome") and str(row.get("nome", "") or "").strip() == base["nome"]:
                    client_ref = row
                    break
        if isinstance(client_ref, dict):
            base["codigo"] = str(client_ref.get("codigo", "") or base.get("codigo", "") or "").strip()
            base["nome"] = str(base.get("nome", "") or client_ref.get("nome", "") or "").strip()
            base["nif"] = str(base.get("nif", "") or client_ref.get("nif", "") or "").strip()
            base["morada"] = str(base.get("morada", "") or client_ref.get("morada", "") or "").strip()
            base["contacto"] = str(base.get("contacto", "") or client_ref.get("contacto", "") or "").strip()
            base["email"] = str(base.get("email", "") or client_ref.get("email", "") or "").strip()
            base["cond_pagamento"] = str(client_ref.get("cond_pagamento", "") or "").strip()
        return base

    def _billing_actor(self) -> str:
        return str((self.user or {}).get("username", "") or "Sistema").strip() or "Sistema"

    def _billing_software_cert_number(self) -> str:
        branding_cfg = {}
        try:
            branding_cfg = dict(self.desktop_main.get_branding_config() or {})
        except Exception:
            branding_cfg = {}
        return str(
            os.getenv("LUGEST_SOFTWARE_CERT_NUMBER", "")
            or os.getenv("LUGEST_SOFTWARE_CERT", "")
            or branding_cfg.get("software_cert", "")
            or branding_cfg.get("software_cert_number", "")
            or ""
        ).strip()

    def _billing_software_producer_info(self, issuer: dict[str, Any] | None = None) -> dict[str, str]:
        issuer_row = dict(issuer or getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        branding_cfg = {}
        try:
            branding_cfg = dict(self.desktop_main.get_branding_config() or {})
        except Exception:
            branding_cfg = {}
        producer_name = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCER_NAME", "")
            or branding_cfg.get("software_producer_name", "")
            or issuer_row.get("nome", "")
            or "LuGEST"
        ).strip() or "LuGEST"
        producer_nif = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCER_NIF", "")
            or branding_cfg.get("software_producer_nif", "")
            or issuer_row.get("nif", "")
            or "999999990"
        ).strip() or "999999990"
        product_id = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCT_ID", "")
            or branding_cfg.get("software_product_id", "")
            or self.tax_compliance.DEFAULT_PRODUCT_ID
        ).strip() or self.tax_compliance.DEFAULT_PRODUCT_ID
        product_version = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCT_VERSION", "")
            or branding_cfg.get("software_product_version", "")
            or self.tax_compliance.DEFAULT_PRODUCT_VERSION
        ).strip() or self.tax_compliance.DEFAULT_PRODUCT_VERSION
        hash_control = str(
            os.getenv("LUGEST_HASH_CONTROL", "")
            or branding_cfg.get("hash_control", "")
            or self.tax_compliance.DEFAULT_HASH_CONTROL
        ).strip() or self.tax_compliance.DEFAULT_HASH_CONTROL
        return {
            "producer_name": producer_name,
            "producer_nif": producer_nif,
            "product_id": product_id,
            "product_version": product_version,
            "hash_control": hash_control,
        }

    def _billing_signing_material(self) -> dict[str, str]:
        return self.tax_compliance.load_or_create_signing_material(self.base_dir)

    def _billing_invoice_snapshot(self, invoice: dict[str, Any] | None) -> dict[str, Any]:
        if not isinstance(invoice, dict):
            return {}
        return self.tax_compliance.deserialize_snapshot(invoice.get("document_snapshot_json", ""))

    def _billing_store_invoice_snapshot(self, invoice: dict[str, Any], document: dict[str, Any]) -> None:
        invoice["document_snapshot_json"] = self.tax_compliance.serialize_snapshot(document)

    def _billing_legal_invoice_no(self, invoice: dict[str, Any]) -> str:
        return self.tax_compliance.legal_document_number(
            invoice.get("doc_type", "FT"),
            invoice.get("serie_id") or invoice.get("serie"),
            invoice.get("seq_num", 0),
            invoice.get("numero_fatura", ""),
        )

    def _billing_status_source_id(self, invoice: dict[str, Any]) -> str:
        return str(
            invoice.get("status_source_id", "")
            or invoice.get("source_id", "")
            or self._billing_actor()
        ).strip() or "Sistema"

    def _billing_invoice_core_fields(self, invoice: dict[str, Any]) -> tuple[str, ...]:
        return (
            str(invoice.get("doc_type", "") or "").strip(),
            str(invoice.get("numero_fatura", "") or "").strip(),
            str(invoice.get("serie_id", "") or invoice.get("serie", "") or "").strip(),
            str(int(self._parse_float(invoice.get("seq_num", 0), 0) or 0)),
            str(invoice.get("atcud", "") or "").strip(),
            str(invoice.get("guia_numero", "") or "").strip(),
            str(invoice.get("data_emissao", "") or "").strip()[:10],
            f"{round(self._parse_float(invoice.get('valor_total', 0), 0), 2):.2f}",
        )

    def _billing_invoice_locked(self, invoice: dict[str, Any] | None) -> bool:
        if not isinstance(invoice, dict):
            return False
        return any(
            str(invoice.get(key, "") or "").strip()
            for key in ("system_entry_date", "hash", "legal_invoice_no", "document_snapshot_json")
        )

    def _billing_previous_signed_hash(self, current_invoice: dict[str, Any]) -> str:
        current_id = str(current_invoice.get("id", "") or "").strip()
        current_doc_type = str(current_invoice.get("doc_type", "") or "FT").strip().upper() or "FT"
        current_series = str(current_invoice.get("serie_id", "") or current_invoice.get("serie", "") or "").strip()
        current_seq = int(self._parse_float(current_invoice.get("seq_num", 0), 0) or 0)
        current_date = str(current_invoice.get("data_emissao", "") or "").strip()[:10]
        previous: dict[str, Any] | None = None
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for row in list(record.get("faturas", []) or []):
                if not isinstance(row, dict):
                    continue
                row_id = str(row.get("id", "") or "").strip()
                if current_id and row_id == current_id:
                    continue
                row_doc_type = str(row.get("doc_type", "") or "FT").strip().upper() or "FT"
                row_series = str(row.get("serie_id", "") or row.get("serie", "") or "").strip()
                if row_doc_type != current_doc_type or row_series != current_series:
                    continue
                row_hash = str(row.get("hash", "") or "").strip()
                if not row_hash:
                    continue
                row_seq = int(self._parse_float(row.get("seq_num", 0), 0) or 0)
                if current_seq > 0 and row_seq > 0:
                    if row_seq >= current_seq:
                        continue
                    if previous is None or row_seq > int(self._parse_float(previous.get("seq_num", 0), 0) or 0):
                        previous = row
                    continue
                row_entry = str(row.get("system_entry_date", "") or row.get("created_at", "") or "").strip()
                current_entry = str(current_invoice.get("system_entry_date", "") or current_invoice.get("created_at", "") or "").strip()
                if row_entry and current_entry and row_entry >= current_entry:
                    continue
                if previous is None:
                    previous = row
                    continue
                prev_entry = str(previous.get("system_entry_date", "") or previous.get("created_at", "") or "").strip()
                if row_entry > prev_entry or (row_entry == prev_entry and str(row.get("data_emissao", "") or "") >= str(previous.get("data_emissao", "") or "")):
                    previous = row
        return str((previous or {}).get("hash", "") or "").strip()

    def _billing_saft_hash_value(self, invoice: dict[str, Any]) -> str:
        if not self._billing_software_cert_number():
            return "0"
        if str(invoice.get("source_billing", "") or "").strip().upper() != "P":
            return "0"
        return str(invoice.get("hash", "") or "").strip() or "0"

    def _billing_saft_hash_control(self, invoice: dict[str, Any]) -> str:
        if not self._billing_software_cert_number():
            return "0"
        if str(invoice.get("source_billing", "") or "").strip().upper() != "P":
            return "0"
        return str(invoice.get("hash_control", "") or "").strip() or "0"

    def _billing_ensure_invoice_compliance(
        self,
        record: dict[str, Any],
        invoice: dict[str, Any],
        *,
        actor: str = "",
        force_snapshot: bool = False,
    ) -> dict[str, Any]:
        actor_txt = str(actor or self._billing_actor()).strip() or "Sistema"
        invoice["system_entry_date"] = str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip()
        invoice["source_id"] = str(invoice.get("source_id", "") or actor_txt).strip() or "Sistema"
        invoice["status_source_id"] = str(invoice.get("status_source_id", "") or invoice.get("source_id", "") or actor_txt).strip() or "Sistema"
        fallback_source = "M" if (str(invoice.get("caminho", "") or "").strip() and int(self._parse_float(invoice.get("seq_num", 0), 0) or 0) <= 0) else "P"
        invoice["source_billing"] = self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback=fallback_source)
        invoice["legal_invoice_no"] = str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip()
        producer = self._billing_software_producer_info()
        invoice["hash_control"] = str(invoice.get("hash_control", "") or producer.get("hash_control", self.tax_compliance.DEFAULT_HASH_CONTROL)).strip()
        if invoice["source_billing"] == "P":
            invoice["previous_hash"] = str(invoice.get("previous_hash", "") or self._billing_previous_signed_hash(invoice)).strip()
            if not str(invoice.get("hash", "") or "").strip():
                signing = self._billing_signing_material()
                seed = self.tax_compliance.build_invoice_hash_message(
                    invoice_date=invoice.get("data_emissao", ""),
                    system_entry_date=invoice.get("system_entry_date", ""),
                    invoice_no=invoice.get("legal_invoice_no", "") or invoice.get("numero_fatura", ""),
                    gross_total=invoice.get("valor_total", 0),
                    previous_hash=invoice.get("previous_hash", ""),
                )
                invoice["hash"] = self.tax_compliance.sign_message_pkcs1_sha1(seed, signing["private_key_pem"])
        else:
            invoice["previous_hash"] = ""
        invoice["communication_status"] = str(invoice.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar"
        needs_snapshot = force_snapshot or not self._billing_invoice_snapshot(invoice)
        if needs_snapshot:
            document = self._billing_build_invoice_document(str(record.get("numero", "") or "").strip(), invoice, prefer_snapshot=False)
            self._billing_store_invoice_snapshot(invoice, document)
            return document
        return self._billing_build_invoice_document(str(record.get("numero", "") or "").strip(), invoice)

    def _billing_next_invoice_identifiers(
        self,
        *,
        issue_date: str = "",
        serie_id: str = "",
        validation_code_hint: str = "",
        reserve: bool = False,
    ) -> dict[str, Any]:
        issue_txt = str(issue_date or self.desktop_main.now_iso()).strip() or self.desktop_main.now_iso()
        sid = str(serie_id or "").strip() or self._billing_default_serie_id(issue_txt)
        ensure_series_fn = getattr(self.desktop_main, "ensure_at_series_record", None)
        if callable(ensure_series_fn):
            serie_obj = ensure_series_fn(
                self.ensure_data(),
                doc_type="FT",
                serie_id=sid,
                issue_date=issue_txt,
                validation_code_hint=str(validation_code_hint or "").strip(),
            )
        else:
            serie_obj = {
                "doc_type": "FT",
                "serie_id": sid,
                "inicio_sequencia": 1,
                "next_seq": 1,
                "validation_code": str(validation_code_hint or "").strip(),
            }
        start_seq = max(1, int(self._parse_float(serie_obj.get("inicio_sequencia", 1), 1) or 1))
        seq = max(start_seq, int(self._parse_float(serie_obj.get("next_seq", start_seq), start_seq) or start_seq))
        used_numbers: set[str] = set()
        used_seq: set[tuple[str, int]] = set()
        year = issue_txt[:4] if len(issue_txt) >= 4 and issue_txt[:4].isdigit() else str(self.desktop_main.datetime.now().year)
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for row in list(record.get("faturas", []) or []):
                if not isinstance(row, dict):
                    continue
                number = str(row.get("numero_fatura", "") or "").strip()
                if number:
                    used_numbers.add(number)
                row_sid = str(row.get("serie_id", "") or row.get("serie", "") or "").strip()
                row_seq = int(self._parse_float(row.get("seq_num", 0), 0) or 0)
                if row_sid and row_seq > 0:
                    used_seq.add((row_sid, row_seq))
        while True:
            number = f"FT-{year}-{seq:04d}"
            if number not in used_numbers and (sid, seq) not in used_seq:
                break
            seq += 1
        validation_code = str(serie_obj.get("validation_code", "") or "").strip()
        if reserve:
            serie_obj["next_seq"] = seq + 1
            if validation_code:
                serie_obj["status"] = "REGISTADA"
            serie_obj["updated_at"] = self.desktop_main.now_iso()
        return {
            "doc_type": "FT",
            "numero_fatura": number,
            "serie": sid,
            "serie_id": sid,
            "seq_num": seq,
            "at_validation_code": validation_code,
            "atcud": f"{validation_code}-{seq}" if validation_code else "",
        }

    def _billing_quote_source(self, quote: dict[str, Any]) -> dict[str, Any]:
        detail = self.orc_detail(str(quote.get("numero", "") or "").strip())
        iva_perc = round(self._parse_float(detail.get("iva_perc", 23), 23), 2)
        lines: list[dict[str, Any]] = []
        for row in list(detail.get("linhas", []) or []):
            qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            if qty <= 0:
                continue
            line_type = self.desktop_main.normalize_orc_line_type(row.get("tipo_item"))
            reference = (
                str(row.get("ref_externa", "") or "").strip()
                or str(row.get("ref_interna", "") or "").strip()
                or str(row.get("produto_codigo", "") or "").strip()
                or str(row.get("conjunto_codigo", "") or "").strip()
            )
            description = str(row.get("descricao", "") or "").strip() or reference or "Artigo"
            material = str(row.get("material", "") or "").strip()
            espessura = str(row.get("espessura", "") or "").strip()
            if line_type == self.desktop_main.ORC_LINE_TYPE_PIECE and material and espessura:
                description = f"{description} | {material} {espessura} mm"
            unit_price = round(self._parse_float(row.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": reference or "-",
                    "description": description or "-",
                    "quantity": qty,
                    "unit": str(row.get("produto_unid", "") or "UN").strip() or "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": str(row.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                    "peca_id": "",
                }
            )
        transport = round(self._parse_float(detail.get("preco_transporte", 0), 0), 2)
        if transport > 0:
            transport_tax = round(transport * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": "TRANSP",
                    "description": "Transporte",
                    "quantity": 1.0,
                    "unit": "SV",
                    "unit_price": transport,
                    "iva_perc": iva_perc,
                    "subtotal": transport,
                    "valor_iva": transport_tax,
                    "total": round(transport + transport_tax, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            )
        return {
            "iva_perc": iva_perc,
            "subtotal": round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2),
            "valor_iva": round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2),
            "total": round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2),
            "lines": lines,
        }

    def _billing_order_source(self, order: dict[str, Any]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(str(order.get("numero", "") or "").strip()) or dict(order or {})
        iva_perc = 23.0
        lines: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            qty = round(self._parse_float(piece.get("quantidade_pedida", 0), 0), 2)
            if qty <= 0:
                continue
            unit_price = round(self._parse_float(piece.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            material = str(piece.get("material", "") or "").strip()
            espessura = str(piece.get("espessura", "") or "").strip()
            description = str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip() or "Peca"
            if material and espessura:
                description = f"{description} | {material} {espessura} mm"
            lines.append(
                {
                    "reference": str(piece.get("ref_externa", "") or piece.get("ref_interna", "") or piece.get("id", "")).strip() or "-",
                    "description": description,
                    "quantity": qty,
                    "unit": "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                    "peca_id": str(piece.get("id", "") or "").strip(),
                }
            )
        for item in list(enc.get("montagem_itens", []) or []):
            qty = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            if qty <= 0:
                continue
            unit_price = round(self._parse_float(item.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": str(item.get("produto_codigo", "") or item.get("conjunto_codigo", "") or "").strip() or "ITEM",
                    "description": str(item.get("descricao", "") or item.get("conjunto_nome", "") or "Item de montagem").strip(),
                    "quantity": qty,
                    "unit": str(item.get("produto_unid", "") or "UN").strip() or "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            )
        return {
            "iva_perc": iva_perc,
            "subtotal": round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2),
            "valor_iva": round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2),
            "total": round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2),
            "lines": lines,
        }

    def _billing_apply_guide_filter(self, lines: list[dict[str, Any]], guide_number: str) -> list[dict[str, Any]]:
        guide_num = str(guide_number or "").strip()
        if not guide_num:
            return [dict(row) for row in lines]
        try:
            guide = self.expedicao_detail(guide_num)
        except Exception:
            return [dict(row) for row in lines]
        qty_by_piece: dict[str, float] = {}
        qty_by_ref_int: dict[str, float] = {}
        qty_by_ref_ext: dict[str, float] = {}
        for row in list(guide.get("lines", []) or []):
            qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            if qty <= 0:
                continue
            piece_id = str(row.get("peca_id", "") or "").strip()
            ref_int = str(row.get("ref_interna", "") or "").strip()
            ref_ext = str(row.get("ref_externa", "") or "").strip()
            if piece_id:
                qty_by_piece[piece_id] = round(qty_by_piece.get(piece_id, 0.0) + qty, 2)
            if ref_int:
                qty_by_ref_int[ref_int] = round(qty_by_ref_int.get(ref_int, 0.0) + qty, 2)
            if ref_ext:
                qty_by_ref_ext[ref_ext] = round(qty_by_ref_ext.get(ref_ext, 0.0) + qty, 2)
        filtered: list[dict[str, Any]] = []
        for row in lines:
            qty = 0.0
            if str(row.get("peca_id", "") or "").strip():
                qty = qty_by_piece.get(str(row.get("peca_id", "") or "").strip(), 0.0)
            if qty <= 0 and str(row.get("ref_interna", "") or "").strip():
                qty = qty_by_ref_int.get(str(row.get("ref_interna", "") or "").strip(), 0.0)
            if qty <= 0 and str(row.get("ref_externa", "") or "").strip():
                qty = qty_by_ref_ext.get(str(row.get("ref_externa", "") or "").strip(), 0.0)
            if qty <= 0:
                continue
            new_row = dict(row)
            new_row["quantity"] = min(qty, round(self._parse_float(row.get("quantity", 0), 0), 2))
            new_row["subtotal"] = round(new_row["quantity"] * self._parse_float(new_row.get("unit_price", 0), 0), 2)
            new_row["valor_iva"] = round(new_row["subtotal"] * (self._parse_float(new_row.get("iva_perc", 0), 0) / 100.0), 2)
            new_row["total"] = round(new_row["subtotal"] + new_row["valor_iva"], 2)
            filtered.append(new_row)
        return filtered or [dict(row) for row in lines]

    def _billing_recalculate_lines(self, lines: list[dict[str, Any]]) -> dict[str, Any]:
        subtotal = round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2)
        tax_value = round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2)
        return {
            "subtotal": subtotal,
            "valor_iva": tax_value,
            "total": round(subtotal + tax_value, 2),
            "lines": lines,
        }

    def _billing_source_snapshot(
        self,
        *,
        record: dict[str, Any],
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        guide_number: str = "",
    ) -> dict[str, Any]:
        if isinstance(quote, dict):
            source = self._billing_quote_source(quote)
        elif isinstance(order, dict):
            source = self._billing_order_source(order)
        else:
            source = {"iva_perc": 23.0, "subtotal": 0.0, "valor_iva": 0.0, "total": 0.0, "lines": []}
        source["lines"] = self._billing_apply_guide_filter(list(source.get("lines", []) or []), guide_number)
        return self._billing_recalculate_lines(list(source.get("lines", []) or [])) | {"iva_perc": round(self._parse_float(source.get("iva_perc", 23), 23), 2)}

    def _billing_adjust_document_total(
        self,
        lines: list[dict[str, Any]],
        *,
        target_total: float,
        default_iva: float,
    ) -> list[dict[str, Any]]:
        current_total = round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2)
        diff = round(target_total - current_total, 2)
        if abs(diff) <= 0.02:
            return lines
        iva_perc = round(self._parse_float(default_iva, 23), 2)
        if iva_perc <= -100:
            adj_subtotal = diff
        else:
            adj_subtotal = round(diff / (1.0 + (iva_perc / 100.0)), 2)
        adj_tax = round(diff - adj_subtotal, 2)
        lines = list(lines or [])
        lines.append(
            {
                "reference": "AJUSTE",
                "description": "Ajuste de faturacao",
                "quantity": 1.0,
                "unit": "SV",
                "unit_price": adj_subtotal,
                "iva_perc": iva_perc,
                "subtotal": adj_subtotal,
                "valor_iva": adj_tax,
                "total": round(adj_subtotal + adj_tax, 2),
                "ref_interna": "",
                "ref_externa": "",
                "peca_id": "",
            }
        )
        return lines

    def _billing_build_invoice_document(self, record_number: str, invoice: dict[str, Any], *, prefer_snapshot: bool = True) -> dict[str, Any]:
        reg_num = str(record_number or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        software_cert = self._billing_software_cert_number()
        invoice_id = str(invoice.get("id", "") or "").strip()
        if prefer_snapshot:
            snapshot = self._billing_invoice_snapshot(invoice)
            if snapshot:
                received = round(
                    sum(
                        self._parse_float(row.get("valor", 0), 0)
                        for row in list(record.get("pagamentos", []) or [])
                        if str(row.get("fatura_id", "") or "").strip() == invoice_id
                    ),
                    2,
                )
                total_amount = round(self._parse_float(snapshot.get("valor_total", 0), 0), 2)
                document = dict(snapshot)
                document["software_cert"] = software_cert
                document["valor_recebido"] = received
                document["saldo"] = round(max(0.0, total_amount - received), 2)
                return document
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        quote = self._billing_quote_by_number(quote_num) if quote_num else None
        order = self._billing_order_by_number(order_num) if order_num else None
        client = self._billing_client_snapshot(quote=quote, order=order, record=record)
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        source = self._billing_source_snapshot(
            record=record,
            quote=quote,
            order=order,
            guide_number=str(invoice.get("guia_numero", "") or "").strip(),
        )
        lines = list(source.get("lines", []) or [])
        target_total = round(self._parse_float(invoice.get("valor_total", source.get("total", 0)), 0), 2)
        default_iva = round(self._parse_float(invoice.get("iva_perc", source.get("iva_perc", 23)), 23), 2)
        if not lines:
            manual_subtotal = round(target_total / (1.0 + (default_iva / 100.0)), 2) if default_iva > -100 else target_total
            manual_tax = round(target_total - manual_subtotal, 2)
            lines = [
                {
                    "reference": "SERVICO",
                    "description": f"Venda associada ao registo {reg_num}",
                    "quantity": 1.0,
                    "unit": "SV",
                    "unit_price": manual_subtotal,
                    "iva_perc": default_iva,
                    "subtotal": manual_subtotal,
                    "valor_iva": manual_tax,
                    "total": round(manual_subtotal + manual_tax, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            ]
        elif target_total > 0:
            lines = self._billing_adjust_document_total(lines, target_total=target_total, default_iva=default_iva)
        totals = self._billing_recalculate_lines(lines)
        recebido = round(
            sum(
                self._parse_float(row.get("valor", 0), 0)
                for row in list(record.get("pagamentos", []) or [])
                if str(row.get("fatura_id", "") or "").strip() == invoice_id
            ),
            2,
        )
        saldo = round(max(0.0, totals["total"] - recebido), 2)
        tax_summary: dict[float, dict[str, Any]] = {}
        for row in lines:
            rate = round(self._parse_float(row.get("iva_perc", default_iva), default_iva), 2)
            bucket = tax_summary.setdefault(rate, {"rate": rate, "base": 0.0, "tax": 0.0})
            bucket["base"] = round(bucket["base"] + self._parse_float(row.get("subtotal", 0), 0), 2)
            bucket["tax"] = round(bucket["tax"] + self._parse_float(row.get("valor_iva", 0), 0), 2)
        tax_summary_rows = [
            {
                "rate": rate,
                "rate_label": f"{self._fmt(rate)}%",
                "base": values["base"],
                "tax": values["tax"],
                "label": f"Base {self._fmt(rate)}%",
            }
            for rate, values in sorted(tax_summary.items(), key=lambda item: item[0])
        ]
        return {
            "titulo": "Fatura",
            "subtitulo": "Documento comercial e fiscal",
            "doc_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
            "numero_fatura": str(invoice.get("numero_fatura", "") or "").strip(),
            "serie": str(invoice.get("serie", "") or invoice.get("serie_id", "") or "").strip(),
            "serie_id": str(invoice.get("serie_id", "") or invoice.get("serie", "") or "").strip(),
            "seq_num": int(self._parse_float(invoice.get("seq_num", 0), 0) or 0),
            "at_validation_code": str(invoice.get("at_validation_code", "") or "").strip(),
            "atcud": str(invoice.get("atcud", "") or "").strip(),
            "legal_invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
            "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
            "source_id": str(invoice.get("source_id", "") or self._billing_actor()).strip() or "Sistema",
            "status_source_id": self._billing_status_source_id(invoice),
            "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
            "hash": str(invoice.get("hash", "") or "").strip(),
            "hash_control": str(invoice.get("hash_control", "") or "").strip(),
            "previous_hash": str(invoice.get("previous_hash", "") or "").strip(),
            "data_emissao": str(invoice.get("data_emissao", "") or self.desktop_main.now_iso())[:10],
            "data_vencimento": str(invoice.get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10],
            "moeda": str(invoice.get("moeda", "") or "EUR").strip() or "EUR",
            "issuer": {
                "nome": str(issuer.get("nome", "") or "").strip(),
                "nif": str(issuer.get("nif", "") or "").strip(),
                "morada": str(issuer.get("morada", "") or "").strip(),
                "extra": "Portugal",
            },
            "customer": client,
            "references": {
                "registo": reg_num,
                "orcamento": quote_num,
                "encomenda": order_num,
                "guia": str(invoice.get("guia_numero", "") or "").strip(),
            },
            "subtotal": totals["subtotal"],
            "valor_iva": totals["valor_iva"],
            "valor_total": totals["total"],
            "valor_recebido": recebido,
            "saldo": saldo,
            "tax_summary": tax_summary_rows,
            "lines": lines,
            "obs": str(invoice.get("obs", "") or record.get("obs", "") or "").strip(),
            "software_cert": software_cert,
            "qr_payload": (
                f"ATCUD:{str(invoice.get('atcud', '') or '').strip() or '-'}"
                f"|DOC:{str(invoice.get('numero_fatura', '') or '').strip() or '-'}"
                f"|DT:{str(invoice.get('data_emissao', '') or '').strip()[:10] or '-'}"
                f"|A:{str(issuer.get('nif', '') or '').strip() or '-'}"
                f"|B:{str(client.get('nif', '') or '').strip() or '-'}"
                f"|TOT:{totals['total']:.2f}"
            ),
        }

    def billing_invoice_defaults(self, numero: str, invoice_id: str = "") -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        row_id = str(invoice_id or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        issue_date = str((existing or {}).get("data_emissao", "") or self.desktop_main.now_iso())[:10]
        base_data = dict(existing or {})
        if existing is None:
            base_data.update(self._billing_next_invoice_identifiers(issue_date=issue_date, reserve=False))
        quote = self._billing_quote_by_number(str(record.get("orcamento_numero", "") or "").strip()) if str(record.get("orcamento_numero", "") or "").strip() else None
        order = self._billing_order_by_number(str(record.get("encomenda_numero", "") or "").strip()) if str(record.get("encomenda_numero", "") or "").strip() else None
        quote, order = self._billing_sync_record_source(record, quote=quote, order=order, persist=True)
        client = self._billing_client_snapshot(
            quote=quote,
            order=order,
            record=record,
        )
        due_date = str((existing or {}).get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10]
        if not due_date:
            try:
                base_date = self.desktop_main.date.fromisoformat(issue_date)
                due_date = (base_date + self.desktop_main.timedelta(days=self._billing_due_days_from_text(client.get("cond_pagamento", "")))).isoformat()
            except Exception:
                due_date = issue_date
        guides = self._billing_guides_for_order(str(record.get("encomenda_numero", "") or "").strip())
        guide_default = str((existing or {}).get("guia_numero", "") or "").strip()
        if not guide_default:
            guide_default = str((guides[0] if guides else {}).get("numero", "") or "").strip()
        base_data.setdefault("doc_type", "FT")
        base_data.setdefault("guia_numero", guide_default)
        base_data.setdefault("data_emissao", issue_date)
        base_data.setdefault("data_vencimento", due_date)
        base_data.setdefault("moeda", "EUR")
        estimated_total = round(self._billing_record_sale_total(record, quote), 2)
        if self._parse_float(base_data.get("valor_total", 0), 0) <= 0 and estimated_total > 0:
            base_data["valor_total"] = estimated_total
        base_data["_allow_zero_total"] = True
        provisional = self._billing_normalize_invoice(base_data, existing)
        doc = self._billing_build_invoice_document(reg_num, provisional)
        provisional["iva_perc"] = round(self._parse_float((existing or {}).get("iva_perc", doc.get("tax_summary", [{}])[0].get("rate", 23) if list(doc.get("tax_summary", []) or []) else 23), 23), 2)
        provisional["subtotal"] = round(self._parse_float((existing or {}).get("subtotal", doc.get("subtotal", 0)), doc.get("subtotal", 0)), 2)
        provisional["valor_iva"] = round(self._parse_float((existing or {}).get("valor_iva", doc.get("valor_iva", 0)), doc.get("valor_iva", 0)), 2)
        provisional["valor_total"] = round(self._parse_float((existing or {}).get("valor_total", doc.get("valor_total", 0)), doc.get("valor_total", 0)), 2)
        provisional["guide_options"] = [
            {
                "numero": str(row.get("numero", "") or "").strip(),
                "label": f"{str(row.get('numero', '') or '').strip()} | {str(row.get('data_emissao', '') or '').strip()[:10]}",
            }
            for row in guides
        ]
        return provisional

    def _billing_normalize_invoice(self, payload: dict[str, Any], existing: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(existing or {})
        row_id = str(payload.get("id", "") or row.get("id", "") or self.desktop_main.uuid.uuid4().hex[:12].upper()).strip()
        doc_type = str(payload.get("doc_type", "") or row.get("doc_type", "") or "FT").strip() or "FT"
        numero_fatura = str(payload.get("numero_fatura", "") or row.get("numero_fatura", "") or "").strip()
        serie = str(payload.get("serie", "") or row.get("serie", "") or "").strip()
        serie_id = str(payload.get("serie_id", "") or row.get("serie_id", "") or serie).strip()
        seq_num = int(self._parse_float(payload.get("seq_num", row.get("seq_num", 0)), 0) or 0)
        at_validation_code = str(payload.get("at_validation_code", "") or row.get("at_validation_code", "") or "").strip()
        atcud = str(payload.get("atcud", "") or row.get("atcud", "") or "").strip()
        guia_numero = str(payload.get("guia_numero", "") or row.get("guia_numero", "") or "").strip()
        data_emissao = str(payload.get("data_emissao", "") or row.get("data_emissao", "") or "").strip()[:10]
        data_vencimento = str(payload.get("data_vencimento", "") or row.get("data_vencimento", "") or "").strip()[:10]
        moeda = str(payload.get("moeda", "") or row.get("moeda", "") or "EUR").strip() or "EUR"
        iva_perc = round(self._parse_float(payload.get("iva_perc", row.get("iva_perc", 23)), 23), 2)
        subtotal = round(self._parse_float(payload.get("subtotal", row.get("subtotal", 0)), 0), 2)
        valor_iva = round(self._parse_float(payload.get("valor_iva", row.get("valor_iva", 0)), 0), 2)
        valor_total = round(self._parse_float(payload.get("valor_total", row.get("valor_total", 0)), 0), 2)
        caminho = str(payload.get("caminho", "") or row.get("caminho", "") or "").strip()
        obs = str(payload.get("obs", "") or row.get("obs", "") or "").strip()
        estado = str(payload.get("estado", "") or row.get("estado", "") or "Emitida").strip() or "Emitida"
        anulada = bool(payload.get("anulada", row.get("anulada", False)))
        anulada_motivo = str(payload.get("anulada_motivo", "") or row.get("anulada_motivo", "") or "").strip()
        anulada_at = str(payload.get("anulada_at", "") or row.get("anulada_at", "") or "").strip()
        legal_invoice_no = str(payload.get("legal_invoice_no", "") or row.get("legal_invoice_no", "") or "").strip()
        system_entry_date = str(payload.get("system_entry_date", "") or row.get("system_entry_date", "") or "").strip()
        source_id = str(payload.get("source_id", "") or row.get("source_id", "") or "").strip()
        source_billing = str(payload.get("source_billing", "") or row.get("source_billing", "") or "").strip()
        status_source_id = str(payload.get("status_source_id", "") or row.get("status_source_id", "") or "").strip()
        hash_value = str(payload.get("hash", "") or row.get("hash", "") or "").strip()
        hash_control = str(payload.get("hash_control", "") or row.get("hash_control", "") or "").strip()
        previous_hash = str(payload.get("previous_hash", "") or row.get("previous_hash", "") or "").strip()
        document_snapshot_json = str(payload.get("document_snapshot_json", "") or row.get("document_snapshot_json", "") or "")
        communication_status = str(payload.get("communication_status", "") or row.get("communication_status", "") or "").strip()
        communication_filename = str(payload.get("communication_filename", "") or row.get("communication_filename", "") or "").strip()
        communication_error = str(payload.get("communication_error", "") or row.get("communication_error", "") or "").strip()
        communicated_at = str(payload.get("communicated_at", "") or row.get("communicated_at", "") or "").strip()
        communication_batch_id = str(payload.get("communication_batch_id", "") or row.get("communication_batch_id", "") or "").strip()
        if anulada or "anulad" in self.desktop_main.norm_text(estado):
            anulada = True
            estado = "Anulada"
        allow_zero_total = bool(payload.get("_allow_zero_total"))
        if not numero_fatura and not caminho:
            raise ValueError("Indica o numero da fatura ou associa o ficheiro.")
        if valor_total <= 0 and not allow_zero_total:
            raise ValueError("Valor da fatura invalido.")
        if not data_emissao:
            data_emissao = str(self.desktop_main.now_iso())[:10]
        return {
            "id": row_id,
            "doc_type": doc_type,
            "numero_fatura": numero_fatura,
            "serie": serie,
            "serie_id": serie_id,
            "seq_num": seq_num,
            "at_validation_code": at_validation_code,
            "atcud": atcud,
            "guia_numero": guia_numero,
            "data_emissao": data_emissao,
            "data_vencimento": data_vencimento,
            "moeda": moeda,
            "iva_perc": iva_perc,
            "subtotal": subtotal,
            "valor_iva": valor_iva,
            "valor_total": valor_total,
            "caminho": caminho,
            "obs": obs,
            "estado": estado,
            "anulada": anulada,
            "anulada_motivo": anulada_motivo,
            "anulada_at": anulada_at,
            "legal_invoice_no": legal_invoice_no,
            "system_entry_date": system_entry_date,
            "source_id": source_id,
            "source_billing": source_billing,
            "status_source_id": status_source_id,
            "hash": hash_value,
            "hash_control": hash_control,
            "previous_hash": previous_hash,
            "document_snapshot_json": document_snapshot_json,
            "communication_status": communication_status,
            "communication_filename": communication_filename,
            "communication_error": communication_error,
            "communicated_at": communicated_at,
            "communication_batch_id": communication_batch_id,
            "created_at": str(row.get("created_at", "") or self.desktop_main.now_iso()),
        }

    def _billing_invoice_is_void(self, invoice: dict[str, Any] | None) -> bool:
        if not isinstance(invoice, dict):
            return False
        if bool(invoice.get("anulada")):
            return True
        return "anulad" in self.desktop_main.norm_text(invoice.get("estado", ""))

    def _billing_active_invoices(self, record: dict[str, Any] | None) -> list[dict[str, Any]]:
        if not isinstance(record, dict):
            return []
        return [
            row
            for row in list(record.get("faturas", []) or [])
            if isinstance(row, dict) and not self._billing_invoice_is_void(row)
        ]

    def _billing_effective_payments(self, record: dict[str, Any] | None) -> list[dict[str, Any]]:
        if not isinstance(record, dict):
            return []
        invoices_by_id = {
            str(row.get("id", "") or "").strip(): row
            for row in list(record.get("faturas", []) or [])
            if isinstance(row, dict)
        }
        effective: list[dict[str, Any]] = []
        for row in list(record.get("pagamentos", []) or []):
            if not isinstance(row, dict):
                continue
            invoice_id = str(row.get("fatura_id", "") or "").strip()
            if invoice_id and self._billing_invoice_is_void(invoices_by_id.get(invoice_id)):
                continue
            effective.append(row)
        return effective

    def _billing_normalize_payment(self, payload: dict[str, Any], existing: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(existing or {})
        row_id = str(payload.get("id", "") or row.get("id", "") or self.desktop_main.uuid.uuid4().hex[:12].upper()).strip()
        data_pagamento = str(payload.get("data_pagamento", "") or row.get("data_pagamento", "") or "").strip()[:10]
        valor = round(self._parse_float(payload.get("valor", row.get("valor", 0)), 0), 2)
        metodo = str(payload.get("metodo", "") or row.get("metodo", "") or "").strip()
        referencia = str(payload.get("referencia", "") or row.get("referencia", "") or "").strip()
        titulo = str(payload.get("titulo_comprovativo", "") or row.get("titulo_comprovativo", "") or "").strip()
        caminho = str(payload.get("caminho_comprovativo", "") or row.get("caminho_comprovativo", "") or "").strip()
        fatura_id = str(payload.get("fatura_id", "") or row.get("fatura_id", "") or "").strip()
        obs = str(payload.get("obs", "") or row.get("obs", "") or "").strip()
        if valor <= 0:
            raise ValueError("Valor do pagamento invalido.")
        if not data_pagamento:
            data_pagamento = str(self.desktop_main.now_iso())[:10]
        return {
            "id": row_id,
            "fatura_id": fatura_id,
            "data_pagamento": data_pagamento,
            "valor": valor,
            "metodo": metodo,
            "referencia": referencia,
            "titulo_comprovativo": titulo,
            "caminho_comprovativo": caminho,
            "obs": obs,
            "created_at": str(row.get("created_at", "") or self.desktop_main.now_iso()),
        }

    def _billing_record_sale_total(self, record: dict[str, Any], quote: dict[str, Any] | None = None) -> float:
        manual = round(self._parse_float(record.get("valor_venda_manual", 0), 0), 2)
        if manual > 0:
            return manual
        if isinstance(quote, dict):
            total_quote = round(self._parse_float(quote.get("total", 0), 0), 2)
            if total_quote > 0:
                return total_quote
        invoices_total = round(
            sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)),
            2,
        )
        return invoices_total

    def _billing_unpaid_due_date(self, record: dict[str, Any]) -> str:
        invoices = self._billing_active_invoices(record)
        payments_total = round(
            sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(record)),
            2,
        )
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in invoices), 2)
        if invoices_total <= 0 or payments_total >= (invoices_total - 0.009):
            return ""
        due_candidates = [
            str(row.get("data_vencimento", "") or "").strip()[:10]
            for row in invoices
            if str(row.get("data_vencimento", "") or "").strip()
        ]
        if due_candidates:
            return sorted(due_candidates)[0]
        return str(record.get("data_vencimento", "") or "").strip()[:10]

    def _billing_invoice_status(self, record: dict[str, Any], quote: dict[str, Any] | None = None) -> str:
        sale_total = self._billing_record_sale_total(record, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)), 2)
        if invoices_total <= 0:
            return "Por faturar"
        if sale_total > 0 and invoices_total < (sale_total - 0.009):
            return "Faturada parcial"
        return "Faturada"

    def _billing_payment_status(self, record: dict[str, Any]) -> str:
        manual = str(record.get("estado_pagamento_manual", "") or "").strip()
        if manual and self.desktop_main.norm_text(manual) not in {"auto", "automatico", "automático"}:
            return manual
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)), 2)
        payments_total = round(
            sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(record)),
            2,
        )
        if invoices_total <= 0:
            return "Sem faturação"
        if payments_total >= (invoices_total - 0.009):
            return "Paga"
        if payments_total > 0:
            return "Parcial"
        due_date = self._billing_unpaid_due_date(record)
        today = self.desktop_main.now_iso()[:10]
        if due_date and due_date < today:
            return "Atrasada"
        return "Pendente"

    def _billing_sync_record_source(
        self,
        record: dict[str, Any] | None,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        persist: bool = False,
    ) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        if not isinstance(record, dict):
            return quote, order
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        if quote is None and quote_num:
            quote = self._billing_quote_by_number(quote_num)
        if order is None and order_num:
            order = self._billing_order_by_number(order_num)
        if order is None and isinstance(quote, dict):
            linked_order_num = str(quote.get("numero_encomenda", "") or "").strip()
            if linked_order_num:
                order = self._billing_order_by_number(linked_order_num)
        if quote is None and isinstance(order, dict):
            linked_quote_num = str(order.get("numero_orcamento", "") or "").strip()
            if linked_quote_num:
                quote = self._billing_quote_by_number(linked_quote_num)
        changed = False
        if isinstance(quote, dict):
            synced_quote_num = str(quote.get("numero", "") or "").strip()
            if synced_quote_num and synced_quote_num != quote_num:
                record["orcamento_numero"] = synced_quote_num
                quote_num = synced_quote_num
                changed = True
        if isinstance(order, dict):
            synced_order_num = str(order.get("numero", "") or "").strip()
            if synced_order_num and synced_order_num != order_num:
                record["encomenda_numero"] = synced_order_num
                order_num = synced_order_num
                changed = True
        client = self._billing_client_info(quote=quote, order=order, record=record)
        client_code = str(client.get("codigo", "") or "").strip()
        client_name = str(client.get("nome", "") or "").strip()
        if client_code and client_code != str(record.get("cliente_codigo", "") or "").strip():
            record["cliente_codigo"] = client_code
            changed = True
        if client_name and client_name != str(record.get("cliente_nome", "") or "").strip():
            record["cliente_nome"] = client_name
            changed = True
        if not str(record.get("data_venda", "") or "").strip():
            sale_date = (
                str((quote or {}).get("data", "") or "").strip()[:10]
                or str((order or {}).get("data_criacao", "") or "").strip()[:10]
            )
            if sale_date:
                record["data_venda"] = sale_date
                changed = True
        if persist and changed:
            record["updated_at"] = self.desktop_main.now_iso()
            self._save(force=True)
        return quote, order

    def _billing_build_row(
        self,
        *,
        source_type: str,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        rec = dict(record or {})
        quote_num = str((quote or {}).get("numero", "") or rec.get("orcamento_numero", "") or "").strip()
        order_num = str((order or {}).get("numero", "") or rec.get("encomenda_numero", "") or "").strip()
        if order is None and order_num:
            order = self._billing_order_by_number(order_num)
        if quote is None and quote_num:
            quote = self._billing_quote_by_number(quote_num)
        client = self._billing_client_info(quote=quote, order=order, record=rec)
        if isinstance(order, dict):
            try:
                self.desktop_main.update_estado_expedicao_encomenda(order)
            except Exception:
                pass
        sale_total = self._billing_record_sale_total(rec, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(rec)), 2)
        payments_total = round(sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(rec)), 2)
        balance = round(max(0.0, invoices_total - payments_total), 2)
        guides = self._billing_guides_for_order(order_num)
        payment_status = self._billing_payment_status(rec) if rec else ("Sem faturação" if invoices_total <= 0 else "Pendente")
        invoice_status = self._billing_invoice_status(rec, quote) if rec else "Por faturar"
        sale_date = (
            str(rec.get("data_venda", "") or "").strip()[:10]
            or str((quote or {}).get("data", "") or "").strip()[:10]
            or str((order or {}).get("data_criacao", "") or "").strip()[:10]
        )
        year = sale_date[:4] if len(sale_date) >= 4 and sale_date[:4].isdigit() else str(self.desktop_main.datetime.now().year)
        latest_invoice = ""
        latest_invoice_date = ""
        if list(rec.get("faturas", []) or []):
            ordered_invoices = sorted(
                list(rec.get("faturas", []) or []),
                key=lambda row: (
                    str((row or {}).get("data_emissao", "") or ""),
                    str((row or {}).get("numero_fatura", "") or ""),
                ),
                reverse=True,
            )
            latest_invoice = str(ordered_invoices[0].get("numero_fatura", "") or "").strip()
            latest_invoice_date = str(ordered_invoices[0].get("data_emissao", "") or "").strip()[:10]
        source_label = "Orçamento vendido" if source_type == "quote" else ("Encomenda direta" if source_type == "order" else "Registo manual")
        return {
            "record_number": str(rec.get("numero", "") or "").strip(),
            "source_type": source_type,
            "source_number": quote_num if source_type == "quote" else (order_num or str(rec.get("numero", "") or "").strip()),
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente": client.get("label", "-") or "-",
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "origem": source_label,
            "estado_encomenda": str((order or {}).get("estado", "") or ("Sem encomenda" if quote_num else "Sem encomenda")).strip() or "Sem encomenda",
            "estado_expedicao": str((order or {}).get("estado_expedicao", "") or ("Sem encomenda" if quote_num else "Sem encomenda")).strip() or "Sem encomenda",
            "estado_faturacao": invoice_status,
            "estado_pagamento": payment_status,
            "vendido": sale_total,
            "faturado": invoices_total,
            "recebido": payments_total,
            "saldo": balance,
            "ultima_fatura": latest_invoice,
            "data_ultima_fatura": latest_invoice_date,
            "ultima_guia": str((guides[0] if guides else {}).get("numero", "") or "").strip(),
            "guias": len(guides),
            "data_venda": sale_date,
            "ano": year,
        }

    def billing_available_years(self) -> list[str]:
        years: set[str] = {str(self.desktop_main.datetime.now().year)}
        for row in self.billing_rows("", "Todas", "Todos"):
            year = str(row.get("ano", "") or "").strip()
            if year:
                years.add(year)
        return sorted(years, key=lambda value: int(value) if value.isdigit() else 0, reverse=True)

    def billing_rows(self, filter_text: str = "", state_filter: str = "Ativas", year: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        year_raw = str(year or "Todos").strip()
        rows: list[dict[str, Any]] = []
        seen_keys: set[str] = set()
        records = {str(row.get("numero", "") or "").strip(): dict(row or {}) for row in self._billing_records() if isinstance(row, dict)}

        def append_row(row: dict[str, Any], key: str) -> None:
            if key in seen_keys:
                return
            if year_raw and year_raw.lower() not in {"todos", "todas", "all", ""} and str(row.get("ano", "") or "").strip() != year_raw:
                return
            estado_faturacao = self.desktop_main.norm_text(row.get("estado_faturacao", ""))
            estado_pagamento = self.desktop_main.norm_text(row.get("estado_pagamento", ""))
            if state_raw and state_raw not in {"todos", "todas", "all", ""}:
                if "ativ" in state_raw and ("paga" in estado_pagamento and "atras" not in estado_pagamento):
                    return
                if "faturar" in state_raw and "por faturar" not in estado_faturacao:
                    return
                if "cobrar" in state_raw and not any(token in estado_pagamento for token in ("pendente", "parcial", "atras")):
                    return
                if "paga" in state_raw and "paga" not in estado_pagamento:
                    return
                if "atras" in state_raw and "atras" not in estado_pagamento:
                    return
            if query and not any(query in str(value).lower() for value in row.values()):
                return
            seen_keys.add(key)
            rows.append(row)

        for quote in list(data.get("orcamentos", []) or []):
            if not isinstance(quote, dict) or not self._billing_quote_is_sold(quote):
                continue
            quote_num = str(quote.get("numero", "") or "").strip()
            order_num = str(quote.get("numero_encomenda", "") or "").strip()
            record = self._billing_find_source_record(quote_num, order_num)
            order = self._billing_order_by_number(order_num) if order_num else None
            row = self._billing_build_row(source_type="quote", quote=quote, order=order, record=record)
            append_row(row, f"quote:{quote_num}")

        for order in list(data.get("encomendas", []) or []):
            if not isinstance(order, dict):
                continue
            order_num = str(order.get("numero", "") or "").strip()
            quote_num = str(order.get("numero_orcamento", "") or "").strip()
            if quote_num:
                continue
            record = self._billing_find_source_record("", order_num)
            row = self._billing_build_row(source_type="order", order=order, record=record)
            append_row(row, f"order:{order_num}")

        for record_num, record in records.items():
            quote_num = str(record.get("orcamento_numero", "") or "").strip()
            order_num = str(record.get("encomenda_numero", "") or "").strip()
            key = f"quote:{quote_num}" if quote_num else (f"order:{order_num}" if order_num else f"record:{record_num}")
            if key in seen_keys:
                continue
            row = self._billing_build_row(
                source_type="record" if not quote_num and not order_num else ("quote" if quote_num else "order"),
                quote=self._billing_quote_by_number(quote_num) if quote_num else None,
                order=self._billing_order_by_number(order_num) if order_num else None,
                record=record,
            )
            append_row(row, key)

        rows.sort(
            key=lambda item: (
                str(item.get("data_venda", "") or "0000-00-00"),
                str(item.get("record_number", "") or ""),
                str(item.get("orcamento_numero", "") or item.get("encomenda_numero", "") or ""),
            ),
            reverse=True,
        )
        return rows

    def billing_dashboard(self) -> dict[str, Any]:
        rows = self.billing_rows("", "Todas", "Todos")
        sold_total = round(sum(self._parse_float(row.get("vendido", 0), 0) for row in rows), 2)
        invoiced_total = round(sum(self._parse_float(row.get("faturado", 0), 0) for row in rows), 2)
        received_total = round(sum(self._parse_float(row.get("recebido", 0), 0) for row in rows), 2)
        balance_total = round(sum(self._parse_float(row.get("saldo", 0), 0) for row in rows), 2)
        return {
            "sold_total": sold_total,
            "invoiced_total": invoiced_total,
            "received_total": received_total,
            "balance_total": balance_total,
            "pending_invoice_count": sum(1 for row in rows if "por faturar" in self.desktop_main.norm_text(row.get("estado_faturacao", ""))),
            "open_payment_count": sum(1 for row in rows if any(token in self.desktop_main.norm_text(row.get("estado_pagamento", "")) for token in ("pendente", "parcial", "atras"))),
            "overdue_count": sum(1 for row in rows if "atras" in self.desktop_main.norm_text(row.get("estado_pagamento", ""))),
            "completed_orders": sum(1 for row in rows if "concl" in self.desktop_main.norm_text(row.get("estado_encomenda", ""))),
            "open_orders": sum(1 for row in rows if row.get("estado_encomenda") and "concl" not in self.desktop_main.norm_text(row.get("estado_encomenda", "")) and "sem encomenda" not in self.desktop_main.norm_text(row.get("estado_encomenda", ""))),
            "row_count": len(rows),
        }

    def _billing_create_record_from_source(self, source_type: str, source_number: str) -> dict[str, Any]:
        source_type_txt = str(source_type or "").strip().lower()
        source_number_txt = str(source_number or "").strip()
        quote = self._billing_quote_by_number(source_number_txt) if source_type_txt == "quote" else None
        order = self._billing_order_by_number(source_number_txt) if source_type_txt == "order" else None
        if quote is None and source_type_txt == "quote":
            raise ValueError("Orçamento não encontrado.")
        if order is None and source_type_txt == "order":
            raise ValueError("Encomenda não encontrada.")
        if quote is not None and not self._billing_quote_is_sold(quote):
            raise ValueError("O orçamento ainda não está marcado como vendido/aprovado.")
        order_num = str((quote or {}).get("numero_encomenda", "") or (order or {}).get("numero", "") or "").strip()
        quote_num = str((quote or {}).get("numero", "") or (order or {}).get("numero_orcamento", "") or "").strip()
        existing = self._billing_find_source_record(quote_num, order_num)
        if existing is not None:
            return existing
        client = self._billing_client_info(quote=quote, order=order)
        sale_date = (
            str((quote or {}).get("data", "") or "").strip()[:10]
            or str((order or {}).get("data_criacao", "") or "").strip()[:10]
            or str(self.desktop_main.now_iso())[:10]
        )
        due_date = ""
        try:
            due_date = (datetime.strptime(sale_date, "%Y-%m-%d") + timedelta(days=30)).strftime("%Y-%m-%d")
        except Exception:
            due_date = ""
        record = {
            "numero": self._billing_next_number(),
            "origem": "Orçamento" if quote_num else "Encomenda",
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "data_venda": sale_date,
            "data_vencimento": due_date,
            "valor_venda_manual": 0.0,
            "estado_pagamento_manual": "",
            "obs": "",
            "created_at": self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "faturas": [],
            "pagamentos": [],
        }
        self.ensure_data().setdefault("faturacao", []).append(record)
        self._save(force=True)
        return record

    def billing_open_record(self, *, source_type: str = "", source_number: str = "", record_number: str = "") -> dict[str, Any]:
        reg_num = str(record_number or "").strip()
        if reg_num:
            return self.billing_detail(reg_num)
        record = self._billing_create_record_from_source(source_type, source_number)
        return self.billing_detail(str(record.get("numero", "") or "").strip())

    def billing_detail(self, numero: str) -> dict[str, Any]:
        record = self._billing_find_record(numero)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        quote = self._billing_quote_by_number(quote_num) if quote_num else None
        order = self._billing_order_by_number(order_num) if order_num else None
        quote, order = self._billing_sync_record_source(record, quote=quote, order=order, persist=True)
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        if isinstance(order, dict):
            try:
                self.desktop_main.update_estado_expedicao_encomenda(order)
            except Exception:
                pass
        client = self._billing_client_info(quote=quote, order=order, record=record)
        client_fiscal = self._billing_client_snapshot(quote=quote, order=order, record=record)
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        guides = self._billing_guides_for_order(order_num)
        invoices = [dict(row) for row in list(record.get("faturas", []) or [])]
        payments = [dict(row) for row in list(record.get("pagamentos", []) or [])]
        invoice_map = {str(row.get("id", "") or "").strip(): row for row in invoices}
        payment_sums: dict[str, float] = {}
        effective_payments = self._billing_effective_payments(record)
        for pay in effective_payments:
            invoice_id = str(pay.get("fatura_id", "") or "").strip()
            if invoice_id:
                payment_sums[invoice_id] = round(payment_sums.get(invoice_id, 0.0) + self._parse_float(pay.get("valor", 0), 0), 2)
        for inv in invoices:
            inv_id = str(inv.get("id", "") or "").strip()
            paid_amount = round(payment_sums.get(inv_id, 0.0), 2)
            total_amount = round(self._parse_float(inv.get("valor_total", 0), 0), 2)
            inv["legal_invoice_no"] = str(inv.get("legal_invoice_no", "") or self._billing_legal_invoice_no(inv)).strip()
            inv["system_entry_date"] = str(inv.get("system_entry_date", "") or inv.get("created_at", "") or "").strip()
            inv["source_billing"] = self.tax_compliance.normalize_source_billing(inv.get("source_billing", ""), fallback="P")
            inv["hash_control"] = str(inv.get("hash_control", "") or self.tax_compliance.DEFAULT_HASH_CONTROL).strip()
            inv["communication_status"] = str(inv.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar"
            if self._billing_invoice_is_void(inv):
                inv["anulada"] = True
                inv["estado"] = "Anulada"
                inv["recebido"] = 0.0
                inv["saldo"] = 0.0
                continue
            inv["recebido"] = paid_amount
            inv["saldo"] = round(max(0.0, total_amount - paid_amount), 2)
            if paid_amount >= (total_amount - 0.009):
                inv["estado"] = "Paga"
            elif paid_amount > 0:
                inv["estado"] = "Parcial"
            else:
                due_date = str(inv.get("data_vencimento", "") or "").strip()
                inv["estado"] = "Atrasada" if (due_date and due_date < self.desktop_main.now_iso()[:10]) else "Pendente"
        sold_total = self._billing_record_sale_total(record, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in invoices if not self._billing_invoice_is_void(row)), 2)
        payments_total = round(sum(self._parse_float(row.get("valor", 0), 0) for row in effective_payments), 2)
        balance = round(max(0.0, invoices_total - payments_total), 2)
        last_invoice = ""
        if invoices:
            last_invoice = str(sorted(invoices, key=lambda row: (str(row.get("data_emissao", "") or ""), str(row.get("numero_fatura", "") or "")), reverse=True)[0].get("numero_fatura", "") or "").strip()
        fiscal_invoice = {}
        if invoices:
            fiscal_invoice = dict(
                sorted(
                    invoices,
                    key=lambda row: (
                        str(row.get("data_emissao", "") or ""),
                        str(row.get("numero_fatura", "") or ""),
                    ),
                    reverse=True,
                )[0]
            )
        return {
            "numero": str(record.get("numero", "") or "").strip(),
            "origem": str(record.get("origem", "") or "").strip(),
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "cliente_label": client.get("label", "-") or "-",
            "cliente_nif": str(client_fiscal.get("nif", "") or "").strip(),
            "cliente_morada": str(client_fiscal.get("morada", "") or "").strip(),
            "cliente_contacto": str(client_fiscal.get("contacto", "") or "").strip(),
            "cliente_email": str(client_fiscal.get("email", "") or "").strip(),
            "emitente_nome": str(issuer.get("nome", "") or "").strip(),
            "emitente_nif": str(issuer.get("nif", "") or "").strip(),
            "emitente_morada": str(issuer.get("morada", "") or "").strip(),
            "data_venda": str(record.get("data_venda", "") or "").strip()[:10],
            "data_vencimento": str(record.get("data_vencimento", "") or "").strip()[:10],
            "valor_venda": sold_total,
            "valor_venda_manual": round(self._parse_float(record.get("valor_venda_manual", 0), 0), 2),
            "estado_faturacao": self._billing_invoice_status(record, quote),
            "estado_pagamento": self._billing_payment_status(record),
            "estado_pagamento_manual": str(record.get("estado_pagamento_manual", "") or "").strip(),
            "valor_faturado": invoices_total,
            "valor_recebido": payments_total,
            "saldo": balance,
            "por_faturar": round(max(0.0, sold_total - invoices_total), 2),
            "obs": str(record.get("obs", "") or "").strip(),
            "order_status": str((order or {}).get("estado", "") or "Sem encomenda").strip() or "Sem encomenda",
            "shipping_status": str((order or {}).get("estado_expedicao", "") or "Sem encomenda").strip() or "Sem encomenda",
            "quote_status": str((quote or {}).get("estado", "") or "").strip(),
            "guide_count": len(guides),
            "last_guide": str((guides[0] if guides else {}).get("numero", "") or "").strip(),
            "last_invoice": last_invoice,
            "fiscal_software_cert": self._billing_software_cert_number() or "Nao configurado",
            "fiscal_legal_invoice_no": str(fiscal_invoice.get("legal_invoice_no", "") or fiscal_invoice.get("numero_fatura", "") or "-").strip() or "-",
            "fiscal_source_billing": str(fiscal_invoice.get("source_billing", "") or "-").strip() or "-",
            "fiscal_system_entry_date": str(fiscal_invoice.get("system_entry_date", "") or "-").strip() or "-",
            "fiscal_hash_control": str(fiscal_invoice.get("hash_control", "") or "-").strip() or "-",
            "fiscal_hash": str(fiscal_invoice.get("hash", "") or "").strip(),
            "fiscal_communication_status": str(fiscal_invoice.get("communication_status", "") or "-").strip() or "-",
            "fiscal_communication_file": str(fiscal_invoice.get("communication_filename", "") or "").strip(),
            "guides": guides,
            "guide_options": [
                {
                    "numero": str(row.get("numero", "") or "").strip(),
                    "label": f"{str(row.get('numero', '') or '').strip()} | {str(row.get('data_emissao', '') or '').strip()[:10]}",
                }
                for row in guides
            ],
            "invoices": invoices,
            "invoice_options": [
                {
                    "id": str(row.get("id", "") or "").strip(),
                    "label": f"{str(row.get('numero_fatura', '') or row.get('id', '')).strip()} | {self._fmt(row.get('valor_total', 0))}",
                }
                for row in invoices
                if not self._billing_invoice_is_void(row)
            ],
            "payments": [
                {
                    **row,
                    "fatura_label": str((invoice_map.get(str(row.get('fatura_id', '') or '').strip()) or {}).get("numero_fatura", "") or "").strip(),
                }
                for row in payments
            ],
        }

    def billing_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        numero = str(payload.get("numero", "") or "").strip()
        record = self._billing_find_record(numero)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        record["data_venda"] = str(payload.get("data_venda", "") or record.get("data_venda", "") or "").strip()[:10]
        record["data_vencimento"] = str(payload.get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10]
        record["valor_venda_manual"] = round(self._parse_float(payload.get("valor_venda_manual", record.get("valor_venda_manual", 0)), 0), 2)
        record["estado_pagamento_manual"] = str(payload.get("estado_pagamento_manual", record.get("estado_pagamento_manual", "") or "") or "").strip()
        record["obs"] = str(payload.get("obs", record.get("obs", "") or "") or "").strip()
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(numero)

    def billing_remove(self, numero: str) -> None:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        if list(record.get("faturas", []) or []) or list(record.get("pagamentos", []) or []):
            raise ValueError("Nao e possivel remover um registo de faturacao com faturas ou pagamentos. Use a anulacao dos documentos e mantenha o historico.")
        rows = list(self._billing_records())
        filtered = [row for row in rows if str((row or {}).get("numero", "") or "").strip() != reg_num]
        if len(filtered) == len(rows):
            raise ValueError("Registo de faturação não encontrado.")
        self.ensure_data()["faturacao"] = filtered
        self._save(force=True)

    def billing_add_invoice(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        if existing is not None and self._billing_invoice_is_void(existing):
            raise ValueError("Nao e possivel editar uma fatura anulada.")
        invoice = self._billing_normalize_invoice(payload_dict, existing)
        if existing is not None and self._billing_invoice_locked(existing) and self._billing_invoice_core_fields(existing) != self._billing_invoice_core_fields(invoice):
            raise ValueError("Nao e possivel alterar os dados fiscais de uma fatura ja emitida.")
        document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=existing is None)
        invoice["subtotal"] = round(self._parse_float(document.get("subtotal", invoice.get("subtotal", 0)), 0), 2)
        invoice["valor_iva"] = round(self._parse_float(document.get("valor_iva", invoice.get("valor_iva", 0)), 0), 2)
        invoice["valor_total"] = round(self._parse_float(document.get("valor_total", invoice.get("valor_total", 0)), 0), 2)
        if not invoice.get("iva_perc") and list(document.get("tax_summary", []) or []):
            invoice["iva_perc"] = round(self._parse_float((document.get("tax_summary", [{}])[0] or {}).get("rate", 23), 23), 2)
        if existing is None:
            record.setdefault("faturas", []).append(invoice)
        else:
            existing.update(invoice)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_generate_invoice_pdf(self, numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        self._billing_sync_record_source(record, persist=True)
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        if existing is not None and self._billing_invoice_is_void(existing):
            raise ValueError("Nao e possivel gerar PDF para uma fatura anulada.")
        if existing is None:
            raw_num = str(payload_dict.get("numero_fatura", "") or "").strip()
            raw_seq = int(self._parse_float(payload_dict.get("seq_num", 0), 0) or 0)
            manual_number = bool(raw_num) and raw_seq <= 0
            if not manual_number:
                payload_dict.update(
                    self._billing_next_invoice_identifiers(
                        issue_date=str(payload_dict.get("data_emissao", "") or self.desktop_main.now_iso())[:10],
                        serie_id=str(payload_dict.get("serie_id", "") or payload_dict.get("serie", "") or "").strip(),
                        validation_code_hint=str(payload_dict.get("at_validation_code", "") or "").strip(),
                        reserve=True,
                    )
                )
        invoice = self._billing_normalize_invoice(payload_dict, existing)
        if existing is not None and self._billing_invoice_locked(existing) and self._billing_invoice_core_fields(existing) != self._billing_invoice_core_fields(invoice):
            raise ValueError("Nao e possivel alterar os dados fiscais de uma fatura ja emitida.")
        document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=existing is None)
        invoice["subtotal"] = round(self._parse_float(document.get("subtotal", 0), 0), 2)
        invoice["valor_iva"] = round(self._parse_float(document.get("valor_iva", 0), 0), 2)
        invoice["valor_total"] = round(self._parse_float(document.get("valor_total", invoice.get("valor_total", 0)), 0), 2)
        if not invoice.get("iva_perc") and list(document.get("tax_summary", []) or []):
            invoice["iva_perc"] = round(self._parse_float((document.get("tax_summary", [{}])[0] or {}).get("rate", 23), 23), 2)
        safe_number = "".join(ch if (ch.isalnum() or ch in ("-", "_")) else "_" for ch in str(invoice.get("numero_fatura", "") or "fatura").strip())
        output_hint = str(payload_dict.get("output_path", "") or payload_dict.get("caminho", "") or "").strip()
        output_path = Path(output_hint) if output_hint else self._storage_output_path("billing/invoices", f"{safe_number}.pdf")
        rendered = self.billing_pdf_actions.render_invoice_pdf(self, output_path, document)
        invoice["caminho"] = self._store_shared_file(rendered, "billing/invoices", preferred_name=f"{safe_number}.pdf")
        if existing is None:
            record.setdefault("faturas", []).append(invoice)
        else:
            existing.update(invoice)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_remove_invoice(self, numero: str, invoice_id: str) -> dict[str, Any]:
        return self.billing_cancel_invoice(numero, invoice_id, "Anulada pelo utilizador.")

    def billing_cancel_invoice(self, numero: str, invoice_id: str, reason: str = "") -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        row_id = str(invoice_id or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None)
        if existing is None:
            raise ValueError("Fatura não encontrada.")
        if self._billing_invoice_is_void(existing):
            raise ValueError("A fatura já está anulada.")
        linked_payments = [
            row
            for row in list(record.get("pagamentos", []) or [])
            if str(row.get("fatura_id", "") or "").strip() == row_id and self._parse_float(row.get("valor", 0), 0) > 0
        ]
        if linked_payments:
            raise ValueError("Nao e possivel anular uma fatura com pagamentos associados. Regulariza primeiro os pagamentos.")
        existing["anulada"] = True
        existing["estado"] = "Anulada"
        existing["anulada_motivo"] = str(reason or "").strip() or "Anulada pelo utilizador."
        existing["anulada_at"] = self.desktop_main.now_iso()
        existing["status_source_id"] = self._billing_actor()
        existing["communication_status"] = "Por comunicar"
        existing["communication_error"] = ""
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_add_payment(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        invoice_id = str(payload_dict.get("fatura_id", "") or "").strip()
        if invoice_id and not any(str(row.get("id", "") or "").strip() == invoice_id for row in list(record.get("faturas", []) or [])):
            raise ValueError("A fatura associada ao pagamento não existe neste registo.")
        if invoice_id:
            invoice = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == invoice_id), None)
            if self._billing_invoice_is_void(invoice):
                raise ValueError("Nao e possivel associar pagamentos a uma fatura anulada.")
        existing = next((row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        payment = self._billing_normalize_payment(payload_dict, existing)
        if existing is None:
            record.setdefault("pagamentos", []).append(payment)
        else:
            existing.update(payment)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_remove_payment(self, numero: str, payment_id: str) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        row_id = str(payment_id or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        before = len(list(record.get("pagamentos", []) or []))
        record["pagamentos"] = [row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() != row_id]
        if len(record["pagamentos"]) == before:
            raise ValueError("Pagamento não encontrado.")
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def _billing_export_invoice_rows(self, start_date: str = "", end_date: str = "") -> list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]:
        start_txt = str(start_date or "").strip()[:10]
        end_txt = str(end_date or "").strip()[:10]
        changed = False
        rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for invoice in list(record.get("faturas", []) or []):
                if not isinstance(invoice, dict):
                    continue
                issue_date = str(invoice.get("data_emissao", "") or "").strip()[:10]
                if start_txt and issue_date and issue_date < start_txt:
                    continue
                if end_txt and issue_date and issue_date > end_txt:
                    continue
                before = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
                document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=False)
                after = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
                if before != after:
                    changed = True
                rows.append((record, invoice, document))
        if changed:
            self._save(force=True)
        return rows

    def _billing_export_payload_from_rows(
        self,
        export_rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]],
        *,
        start_date: str = "",
        end_date: str = "",
    ) -> dict[str, Any]:
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        producer = self._billing_software_producer_info(issuer)
        customers_map: dict[str, dict[str, Any]] = {}
        products_map: dict[str, dict[str, Any]] = {}
        tax_map: dict[tuple[str, str, str], dict[str, Any]] = {}
        invoices_payload: list[dict[str, Any]] = []
        issue_dates = [str(invoice.get("data_emissao", "") or "").strip()[:10] for _, invoice, _ in export_rows if str(invoice.get("data_emissao", "") or "").strip()[:10]]
        computed_start = start_date[:10] if start_date else (min(issue_dates) if issue_dates else str(self.desktop_main.now_iso())[:10])
        computed_end = end_date[:10] if end_date else (max(issue_dates) if issue_dates else str(self.desktop_main.now_iso())[:10])
        for _record, invoice, document in export_rows:
            customer = dict(document.get("customer", {}) or {})
            customer_id = str(customer.get("codigo", "") or customer.get("nif", "") or "CONSUMIDOR-FINAL").strip() or "CONSUMIDOR-FINAL"
            customer_tax_id = str(customer.get("nif", "") or "999999990").strip() or "999999990"
            customers_map.setdefault(
                customer_id,
                {
                    "customer_id": customer_id,
                    "account_id": customer_id,
                    "tax_id": customer_tax_id,
                    "name": str(customer.get("nome", "") or "Cliente").strip() or "Cliente",
                    "address_detail": str(customer.get("morada", "") or "-").strip() or "-",
                    "city": "-",
                    "postal_code": "0000-000",
                    "country": "PT",
                },
            )
            line_payloads: list[dict[str, Any]] = []
            for line in list(document.get("lines", []) or []):
                product_code = str(line.get("reference", "") or line.get("ref_externa", "") or line.get("ref_interna", "") or "ITEM").strip() or "ITEM"
                product_type = self.tax_compliance.product_type_from_line(line)
                products_map.setdefault(
                    product_code,
                    {
                        "product_type": product_type,
                        "product_code": product_code,
                        "product_group": "SERVICOS" if product_type == "S" else "ARTIGOS",
                        "product_description": str(line.get("description", "") or product_code).strip() or product_code,
                        "product_number_code": product_code,
                    },
                )
                tax_rate = round(self._parse_float(line.get("iva_perc", 0), 0), 2)
                tax_code = self.tax_compliance.tax_code_from_rate(tax_rate)
                tax_map.setdefault(
                    ("IVA", "PT", tax_code),
                    {
                        "tax_type": "IVA",
                        "tax_country_region": "PT",
                        "tax_code": tax_code,
                        "tax_percentage": tax_rate,
                        "description": "IVA" if tax_rate > 0 else "Nao sujeito",
                    },
                )
                line_payloads.append(
                    {
                        "product_code": product_code,
                        "product_description": str(line.get("description", "") or product_code).strip() or product_code,
                        "quantity": round(self._parse_float(line.get("quantity", 0), 0), 3),
                        "unit_of_measure": str(line.get("unit", "") or "UN").strip() or "UN",
                        "unit_price": round(self._parse_float(line.get("unit_price", 0), 0), 2),
                        "description": str(line.get("description", "") or product_code).strip() or product_code,
                        "credit_amount": round(self._parse_float(line.get("subtotal", 0), 0), 2),
                        "tax_type": "IVA",
                        "tax_country_region": "PT",
                        "tax_code": tax_code,
                        "tax_percentage": tax_rate,
                    }
                )
            invoice_status_date = str(invoice.get("anulada_at", "") or invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip()
            invoices_payload.append(
                {
                    "invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
                    "invoice_status": self.tax_compliance.invoice_status_code(is_void=self._billing_invoice_is_void(invoice)),
                    "invoice_status_date": invoice_status_date,
                    "status_source_id": self._billing_status_source_id(invoice),
                    "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
                    "hash": self._billing_saft_hash_value(invoice),
                    "hash_control": self._billing_saft_hash_control(invoice),
                    "period": int((str(invoice.get("data_emissao", "") or "")[5:7] or "0")) if str(invoice.get("data_emissao", "") or "").strip()[:7] else 0,
                    "invoice_date": str(invoice.get("data_emissao", "") or "").strip()[:10],
                    "invoice_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
                    "source_id": str(invoice.get("source_id", "") or self._billing_actor()).strip() or "Sistema",
                    "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
                    "customer_id": customer_id,
                    "lines": line_payloads,
                    "tax_payable": round(self._parse_float(document.get("valor_iva", 0), 0), 2),
                    "net_total": round(self._parse_float(document.get("subtotal", 0), 0), 2),
                    "gross_total": round(self._parse_float(document.get("valor_total", 0), 0), 2),
                }
            )
        return {
            "header": {
                "audit_file_version": self.tax_compliance.SAFT_PT_AUDIT_FILE_VERSION,
                "company_id": str(issuer.get("nif", "") or producer.get("producer_nif", "999999990")).strip() or "999999990",
                "tax_registration_number": str(issuer.get("nif", "") or "999999990").strip() or "999999990",
                "tax_accounting_basis": "F",
                "company_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "business_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "company_address_detail": str(issuer.get("morada", "") or "-").strip() or "-",
                "company_city": "-",
                "company_postal_code": "0000-000",
                "company_country": "PT",
                "fiscal_year": (computed_start[:4] if len(computed_start) >= 4 else str(self.desktop_main.datetime.now().year)),
                "start_date": computed_start,
                "end_date": computed_end,
                "currency_code": "EUR",
                "date_created": str(self.desktop_main.now_iso())[:10],
                "tax_entity": "Global",
                "product_company_tax_id": producer.get("producer_nif", "999999990"),
                "software_certificate_number": self._billing_software_cert_number() or "0",
                "product_id": producer.get("product_id", self.tax_compliance.DEFAULT_PRODUCT_ID),
                "product_version": producer.get("product_version", self.tax_compliance.DEFAULT_PRODUCT_VERSION),
                "header_comment": "Exportação interna LuGEST SAF-T(PT) - faturação.",
            },
            "customers": list(customers_map.values()),
            "products": list(products_map.values()),
            "tax_table": list(tax_map.values()),
            "invoices": invoices_payload,
        }

    def _billing_prepare_at_payload(
        self,
        export_rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]],
    ) -> tuple[dict[str, Any], list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]], str]:
        pending: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for record, invoice, document in export_rows:
            if self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P") != "P":
                continue
            status = str(invoice.get("communication_status", "") or "").strip().lower()
            if status == "comunicada":
                continue
            pending.append((record, invoice, document))
        if not pending:
            raise ValueError("Nao existem faturas pendentes para preparar comunicacao AT.")
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        producer = self._billing_software_producer_info(issuer)
        batch_id = self.desktop_main.uuid.uuid4().hex[:12].upper()
        payload = {
            "header": {
                "generated_at": self.desktop_main.now_iso(),
                "issuer_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "issuer_tax_id": str(issuer.get("nif", "") or "999999990").strip() or "999999990",
                "software_certificate_number": self._billing_software_cert_number() or "0",
                "product_id": producer.get("product_id", self.tax_compliance.DEFAULT_PRODUCT_ID),
                "product_version": producer.get("product_version", self.tax_compliance.DEFAULT_PRODUCT_VERSION),
                "preparation_mode": "manual",
            },
            "documents": [
                {
                    "document_id": str(invoice.get("id", "") or "").strip(),
                    "invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
                    "invoice_date": str(invoice.get("data_emissao", "") or "").strip()[:10],
                    "invoice_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
                    "atcud": str(invoice.get("atcud", "") or "").strip(),
                    "hash": self._billing_saft_hash_value(invoice),
                    "hash_control": self._billing_saft_hash_control(invoice),
                    "customer_tax_id": str((document.get("customer", {}) or {}).get("nif", "") or "999999990").strip() or "999999990",
                    "gross_total": round(self._parse_float(document.get("valor_total", 0), 0), 2),
                    "status": str(invoice.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar",
                    "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
                    "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
                }
                for _, invoice, document in pending
            ],
        }
        return payload, pending, batch_id

    def _billing_record_export_rows(self, numero: str) -> list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        changed = False
        rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for invoice in list(record.get("faturas", []) or []):
            if not isinstance(invoice, dict):
                continue
            before = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
            document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=False)
            after = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
            if before != after:
                changed = True
            rows.append((record, invoice, document))
        if not rows:
            raise ValueError("Este registo ainda não tem faturas para exportar.")
        if changed:
            self._save(force=True)
        return rows

    def billing_export_saft_pt(self, start_date: str = "", end_date: str = "", output_path: str = "") -> str:
        export_rows = self._billing_export_invoice_rows(start_date, end_date)
        if not export_rows:
            raise ValueError("Nao existem faturas no intervalo indicado para exportar SAF-T(PT).")
        output_target = Path(output_path) if str(output_path or "").strip() else (self.base_dir / "generated" / "compliance" / "saft" / f"saft_pt_{self.desktop_main.datetime.now().strftime('%Y%m%d_%H%M%S')}.xml")
        rendered = self.tax_compliance.render_saft_pt_xml(self._billing_export_payload_from_rows(export_rows, start_date=start_date, end_date=end_date), output_target)
        return str(rendered)

    def billing_prepare_at_communication_batch(self, start_date: str = "", end_date: str = "", output_path: str = "") -> str:
        export_rows = self._billing_export_invoice_rows(start_date, end_date)
        payload, pending, batch_id = self._billing_prepare_at_payload(export_rows)
        output_target = Path(output_path) if str(output_path or "").strip() else (self.base_dir / "generated" / "compliance" / "at" / f"at_preparacao_{batch_id}.xml")
        rendered = self.tax_compliance.render_at_communication_preparation_xml(payload, output_target)
        for record, invoice, _document in pending:
            invoice["communication_status"] = "Preparada"
            invoice["communication_filename"] = str(rendered)
            invoice["communication_batch_id"] = batch_id
            invoice["communication_error"] = ""
            record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return str(rendered)

    def billing_export_record_saft_pt(self, numero: str, output_path: str = "") -> str:
        export_rows = self._billing_record_export_rows(numero)
        issue_dates = [str(invoice.get("data_emissao", "") or "").strip()[:10] for _, invoice, _ in export_rows if str(invoice.get("data_emissao", "") or "").strip()[:10]]
        output_target = Path(output_path) if str(output_path or "").strip() else self._storage_output_path("billing/compliance", f"saft_pt_{str(numero or '').strip() or 'registo'}.xml")
        rendered = self.tax_compliance.render_saft_pt_xml(
            self._billing_export_payload_from_rows(
                export_rows,
                start_date=min(issue_dates) if issue_dates else "",
                end_date=max(issue_dates) if issue_dates else "",
            ),
            output_target,
        )
        return str(self._store_shared_file(rendered, "billing/compliance", preferred_name=Path(str(rendered)).name))

    def billing_prepare_record_at_communication_batch(self, numero: str, output_path: str = "") -> str:
        export_rows = self._billing_record_export_rows(numero)
        payload, pending, batch_id = self._billing_prepare_at_payload(export_rows)
        output_target = Path(output_path) if str(output_path or "").strip() else self._storage_output_path("billing/compliance", f"at_preparacao_{str(numero or '').strip() or batch_id}.xml")
        rendered = self.tax_compliance.render_at_communication_preparation_xml(payload, output_target)
        rendered_ref = self._store_shared_file(rendered, "billing/compliance", preferred_name=Path(str(rendered)).name)
        for record, invoice, _document in pending:
            invoice["communication_status"] = "Preparada"
            invoice["communication_filename"] = rendered_ref
            invoice["communication_batch_id"] = batch_id
            invoice["communication_error"] = ""
            record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return rendered_ref

    def billing_open_path(self, path: str) -> Path:
        return self.open_file_reference(path)

    def ne_suppliers(self) -> list[dict[str, str]]:
        self._maybe_normalize_single_supplier_catalog()
        rows = []
        for raw in list(self.ensure_data().get("fornecedores", []) or []):
            row = {
                "id": str(raw.get("id", "") or "").strip(),
                "nome": str(raw.get("nome", "") or "").strip(),
                "contacto": str(raw.get("contacto", "") or "").strip(),
                "nif": str(raw.get("nif", "") or "").strip(),
                "morada": str(raw.get("morada", "") or "").strip(),
                "email": str(raw.get("email", "") or "").strip(),
                "codigo_postal": str(raw.get("codigo_postal", "") or "").strip(),
                "localidade": str(raw.get("localidade", "") or "").strip(),
                "pais": str(raw.get("pais", "") or "").strip(),
                "cond_pagamento": str(raw.get("cond_pagamento", "") or "").strip(),
                "prazo_entrega_dias": int(self._parse_float(raw.get("prazo_entrega_dias", 0), 0)),
                "website": str(raw.get("website", "") or "").strip(),
                "obs": str(raw.get("obs", "") or "").strip(),
            }
            rows.append(row)
        def _supplier_sort_key(item: dict[str, Any]) -> tuple[int, str]:
            supplier_id = str(item.get("id", "") or "").strip().upper()
            if supplier_id.startswith("FOR-") and supplier_id[4:].isdigit():
                return (int(supplier_id[4:]), supplier_id)
            return (10**9, supplier_id)
        rows.sort(key=_supplier_sort_key)
        return rows

    def supplier_next_id(self) -> str:
        self._maybe_normalize_single_supplier_catalog()
        return str(self.desktop_main.peek_next_fornecedor_numero(self.ensure_data()))

    def _maybe_normalize_single_supplier_catalog(self) -> None:
        if getattr(self, "_supplier_catalog_fixing", False):
            return
        data = self.ensure_data()
        suppliers = list(data.get("fornecedores", []) or [])
        if len(suppliers) != 1:
            return
        supplier = suppliers[0]
        old_id = str(supplier.get("id", "") or "").strip()
        supplier_name = str(supplier.get("nome", "") or "").strip()
        target_id = "FOR-0001"
        current_next = str(self.desktop_main.peek_next_fornecedor_numero(data) or "").strip().upper()
        notes = list(data.get("notas_encomenda", []) or [])
        needs_fix = old_id != target_id or current_next != "FOR-0002"
        if not needs_fix and not any(str(note.get("fornecedor", "") or "").strip() == old_id for note in notes):
            return
        self._supplier_catalog_fixing = True
        changed = False
        try:
            if old_id != target_id:
                supplier["id"] = target_id
                changed = True
            for note in notes:
                note_supplier_id = str(note.get("fornecedor_id", "") or "").strip()
                note_supplier_txt = str(note.get("fornecedor", "") or "").strip()
                if note_supplier_id in {old_id, target_id}:
                    if note_supplier_id != target_id:
                        note["fornecedor_id"] = target_id
                        changed = True
                    if note_supplier_txt != supplier_name:
                        note["fornecedor"] = supplier_name
                        changed = True
                elif note_supplier_txt == old_id:
                    note["fornecedor"] = supplier_name
                    changed = True
                for line in list(note.get("linhas", []) or []):
                    line_supplier = str(line.get("fornecedor_linha", "") or "").strip()
                    if line_supplier == old_id or line_supplier == target_id or line_supplier.startswith(f"{old_id} - "):
                        if line_supplier != supplier_name:
                            line["fornecedor_linha"] = supplier_name
                            changed = True
            self.desktop_main.reserve_fornecedor_numero(data, target_id)
            if str(self.desktop_main.peek_next_fornecedor_numero(data) or "").strip().upper() != "FOR-0002":
                self.desktop_main._store_fornecedor_sequence_next(data, 2)
                changed = True
            if changed:
                self._save(force=True)
        finally:
            self._supplier_catalog_fixing = False

    def supplier_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._maybe_normalize_single_supplier_catalog()
        data = self.ensure_data()
        supplier_id = str(payload.get("id", "") or "").strip()
        nome = str(payload.get("nome", "") or "").strip()
        if not nome:
            raise ValueError("Nome do fornecedor obrigatorio.")
        rows = data.setdefault("fornecedores", [])
        existing = next((item for item in rows if str(item.get("id", "") or "").strip() == supplier_id), None) if supplier_id else None
        if not supplier_id:
            supplier_id = str(self.desktop_main.next_fornecedor_numero(data))
        elif existing is None:
            self.desktop_main.reserve_fornecedor_numero(data, supplier_id)
        row = {
            "id": supplier_id,
            "nome": nome,
            "nif": str(payload.get("nif", "") or "").strip(),
            "morada": str(payload.get("morada", "") or "").strip(),
            "contacto": str(payload.get("contacto", "") or "").strip(),
            "email": str(payload.get("email", "") or "").strip(),
            "codigo_postal": str(payload.get("codigo_postal", "") or "").strip(),
            "localidade": str(payload.get("localidade", "") or "").strip(),
            "pais": str(payload.get("pais", "") or "").strip(),
            "cond_pagamento": str(payload.get("cond_pagamento", "") or "").strip(),
            "prazo_entrega_dias": int(self._parse_float(payload.get("prazo_entrega_dias", 0), 0)),
            "website": str(payload.get("website", "") or "").strip(),
            "obs": str(payload.get("obs", "") or "").strip(),
        }
        if existing is None:
            rows.append(row)
            target = row
        else:
            existing.update(row)
            target = existing
        self._save(force=True)
        return dict(target)

    def supplier_remove(self, supplier_id: str) -> None:
        data = self.ensure_data()
        value = str(supplier_id or "").strip()
        if not value:
            raise ValueError("Fornecedor inv?lido.")
        if any(str(note.get("fornecedor_id", "") or "").strip() == value for note in list(data.get("notas_encomenda", []) or [])):
            raise ValueError("Nao e possivel remover um fornecedor usado em notas de encomenda.")
        before = len(list(data.get("fornecedores", []) or []))
        data["fornecedores"] = [row for row in list(data.get("fornecedores", []) or []) if str(row.get("id", "") or "").strip() != value]
        if len(data["fornecedores"]) == before:
            raise ValueError("Fornecedor n?o encontrado.")
        self._save(force=True)

    def ne_next_number(self) -> str:
        return str(self.desktop_main.peek_next_ne_numero(self.ensure_data()))

    def ne_rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        rows = []
        for note in sorted(list(self.ensure_data().get("notas_encomenda", []) or []), key=lambda item: str(item.get("numero", "") or "")):
            if note.get("oculta") and "convert" not in state_raw and "todas" not in state_raw and "todos" not in state_raw:
                continue
            estado = str(note.get("estado", "Em edicao") or "Em edicao").strip()
            estado_norm = self.desktop_main.norm_text(estado)
            is_partial = "parcial" in estado_norm
            is_delivered = "entreg" in estado_norm and not is_partial
            is_converted = "convert" in estado_norm
            if state_raw and state_raw not in ("todas", "todos", "all"):
                if "ativ" in state_raw and (is_delivered or is_converted):
                    continue
                if "edi" in state_raw and "edi" not in estado_norm:
                    continue
                if "apro" in state_raw and "apro" not in estado_norm:
                    continue
                if "enviad" in state_raw and "enviad" not in estado_norm:
                    continue
                if "parcial" in state_raw and not is_partial:
                    continue
                if "entreg" in state_raw and not is_delivered:
                    continue
                if "convert" in state_raw and not is_converted:
                    continue
            row = {
                "numero": str(note.get("numero", "") or "").strip(),
                "fornecedor": str(note.get("fornecedor", "") or "").strip() or ("Multi-fornecedor" if self._note_kind(note) == "rfq" else "Por adjudicar"),
                "data_entrega": str(note.get("data_entrega", "") or "").strip(),
                "estado": estado,
                "total": round(self._parse_float(note.get("total", 0), 0), 2),
                "linhas": len(list(note.get("linhas", []) or [])),
                "draft": bool(note.get("_draft")),
                "oculta": bool(note.get("oculta")),
                "kind": self._note_kind(note),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def ne_material_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for record in list(self.ensure_data().get("materiais", []) or []):
            material_id = str(record.get("id", "") or "").strip()
            material = str(record.get("material", "") or "").strip()
            esp_raw = str(record.get("espessura", "") or "").strip()
            esp = self._fmt(esp_raw) if esp_raw else ""
            formato = str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "Chapa").strip()
            metrics = self.material_price_preview(record)
            preco_unid = float(metrics.get("preco_unid", 0.0) or 0.0)
            comp = round(self._parse_dimension_mm(metrics.get("comprimento", record.get("comprimento", 0)), 0), 3)
            larg = round(self._parse_dimension_mm(metrics.get("largura", record.get("largura", 0)), 0), 3)
            altura = round(self._parse_dimension_mm(metrics.get("altura", record.get("altura", 0)), 0), 3)
            diametro = round(self._parse_dimension_mm(metrics.get("diametro", record.get("diametro", 0)), 0), 3)
            metros = round(self._parse_float(metrics.get("metros", record.get("metros", 0)), 0), 4)
            dim_txt = str(metrics.get("dimension_label", "") or "").strip() or "-"
            esp_txt = f" | {esp} mm" if esp else ""
            lote_txt = str(record.get("lote_fornecedor", "") or "").strip()
            desc = f"{formato} | {material}{esp_txt} | {dim_txt}"
            if metros > 0:
                desc = f"{desc} | {self._fmt(metros)} m"
            if lote_txt:
                desc = f"{desc} | Lote {lote_txt}"
            row = {
                "id": material_id,
                "descricao": desc,
                "material": material,
                "espessura": esp,
                "formato": formato,
                "preco": round(preco_unid, 4),
                "preco_base": round(self._parse_float(record.get("p_compra", 0), 0), 4),
                "preco_base_label": str(metrics.get("base_label", "EUR/kg")),
                "unid": "UN",
                "lote": str(record.get("lote_fornecedor", "") or "").strip(),
                "localizacao": self._localizacao(record),
                "comprimento": round(comp, 3),
                "largura": round(larg, 3),
                "altura": round(altura, 3),
                "diametro": round(diametro, 3),
                "dimensao": dim_txt,
                "secao_tipo": str(metrics.get("secao_tipo", record.get("secao_tipo", "")) or "").strip(),
                "secao_label": str(metrics.get("secao_label", "") or "").strip(),
                "kg_m": round(self._parse_float(metrics.get("kg_m", record.get("kg_m", 0)), 0), 4),
                "metros": round(metros, 4),
                "peso_unid": round(self._parse_float(metrics.get("peso_unid", record.get("peso_unid", 0)), 0), 4),
                "material_familia": str(record.get("material_familia", "") or "").strip(),
                "material_familia_resolved": str(metrics.get("material_familia_resolved", "") or "").strip(),
                "material_familia_label": str(metrics.get("material_familia_label", "") or "").strip(),
                "densidade": round(self._parse_float(metrics.get("densidade", 0), 0), 3),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("material") or "", self._parse_float(item.get("espessura", 0), 0), item.get("id") or ""))
        return rows

    def ne_product_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        rows = []
        for raw in self.product_rows(filter_text):
            categoria = str(raw.get("categoria", "") or "").strip()
            tipo = str(raw.get("tipo", "") or "").strip()
            rows.append(
                {
                    "codigo": str(raw.get("codigo", "") or "").strip(),
                    "descricao": str(raw.get("descricao", "") or "").strip(),
                    "origem": "Produto",
                    "stock": round(self._parse_float(raw.get("qty", 0), 0), 2),
                    "unid": str(raw.get("unid", "UN") or "UN").strip(),
                    "preco": round(self._parse_float(raw.get("preco_unid", 0), 0), 4),
                    "preco_unid": round(self._parse_float(raw.get("preco_unid", 0), 0), 4),
                    "p_compra": round(self._parse_float(raw.get("p_compra", 0), 0), 4),
                    "categoria": categoria,
                    "tipo": tipo,
                    "dimensoes": str(raw.get("dimensoes", "") or "").strip(),
                    "peso_unid": round(self._parse_float(raw.get("peso_unid", 0), 0), 4),
                    "metros_unidade": round(self._parse_float(raw.get("metros_unidade", 0), 0), 4),
                    "price_mode": str(self.desktop_main.produto_modo_preco(categoria, tipo) or "compra").strip(),
                }
            )
        return rows

    def ne_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        documents = self._ne_document_rows(note)
        lines = []
        for line in list(note.get("linhas", []) or []):
            qtd = self._parse_float(line.get("qtd", 0), 0)
            qtd_ent = self._parse_float(line.get("qtd_entregue", qtd if line.get("entregue") else 0), 0)
            if qtd_ent <= 0:
                entrega = "PENDENTE"
            elif qtd_ent < max(0.0, qtd - 1e-9):
                entrega = f"PARCIAL ({self._fmt(qtd_ent)}/{self._fmt(qtd)})"
            else:
                entrega = "ENTREGUE"
            lines.append(
                {
                    "ref": str(line.get("ref", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "fornecedor_linha": str(line.get("fornecedor_linha", "") or "").strip(),
                    "origem": str(line.get("origem", "Produto") or "Produto").strip(),
                    "qtd": round(qtd, 4),
                    "unid": str(line.get("unid", "") or "").strip(),
                    "preco": round(self._parse_float(line.get("preco", 0), 0), 4),
                    "desconto": round(self._parse_float(line.get("desconto", 0), 0), 2),
                    "iva": round(self._parse_float(line.get("iva", 23), 23), 2),
                    "total": round(self._parse_float(line.get("total", 0), 0), 4),
                    "entrega": entrega,
                    "material": str(line.get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or "").strip(),
                    "comprimento": round(self._parse_dimension_mm(line.get("comprimento", 0), 0), 3),
                    "largura": round(self._parse_dimension_mm(line.get("largura", 0), 0), 3),
                    "altura": round(self._parse_dimension_mm(line.get("altura", 0), 0), 3),
                    "diametro": round(self._parse_dimension_mm(line.get("diametro", 0), 0), 3),
                    "metros": round(self._parse_float(line.get("metros", 0), 0), 4),
                    "kg_m": round(self._parse_float(line.get("kg_m", 0), 0), 4),
                    "localizacao": str(line.get("localizacao", "") or "").strip(),
                    "lote_fornecedor": str(line.get("lote_fornecedor", "") or "").strip(),
                    "peso_unid": round(self._parse_float(line.get("peso_unid", 0), 0), 4),
                    "p_compra": round(self._parse_float(line.get("p_compra", 0), 0), 6),
                    "formato": str(line.get("formato", "") or "").strip(),
                    "secao_tipo": str(line.get("secao_tipo", "") or "").strip(),
                    "material_familia": str(line.get("material_familia", "") or "").strip(),
                    "_material_manual": bool(line.get("_material_manual")),
                    "_material_pending_create": bool(line.get("_material_pending_create")),
                }
            )
        return {
            "numero": str(note.get("numero", "") or "").strip(),
            "fornecedor": str(note.get("fornecedor", "") or "").strip(),
            "fornecedor_id": str(note.get("fornecedor_id", "") or "").strip(),
            "contacto": str(note.get("contacto", "") or "").strip(),
            "data_entrega": str(note.get("data_entrega", "") or "").strip(),
            "obs": str(note.get("obs", "") or "").strip(),
            "local_descarga": str(note.get("local_descarga", "") or "").strip(),
            "meio_transporte": str(note.get("meio_transporte", "") or "").strip(),
            "estado": str(note.get("estado", "Em edicao") or "Em edicao").strip(),
            "total": round(self._parse_float(note.get("total", 0), 0), 2),
            "draft": bool(note.get("_draft")),
            "kind": self._note_kind(note),
            "ne_geradas": list(note.get("ne_geradas", []) or []),
            "origem_cotacao": str(note.get("origem_cotacao", "") or "").strip(),
            "guia_ultima": str(note.get("guia_ultima", "") or "").strip(),
            "fatura_ultima": str(note.get("fatura_ultima", "") or "").strip(),
            "fatura_caminho_ultima": str(note.get("fatura_caminho_ultima", "") or "").strip(),
            "data_doc_ultima": str(note.get("data_doc_ultima", "") or "").strip(),
            "data_ultima_entrega": str(note.get("data_ultima_entrega", "") or "").strip(),
            "documents": documents,
            "document_count": len(documents),
            "lines": lines,
        }

    def _find_existing_material_from_note_line(self, line: dict[str, Any]) -> dict[str, Any] | None:
        material_id = str(line.get("ref", "") or "").strip()
        material_txt = self._norm_material_token(line.get("material", ""))
        esp_txt = self._norm_esp_token(line.get("espessura", ""))
        formato_txt = str(line.get("formato", "") or "Chapa").strip() or "Chapa"
        lote_txt = str(line.get("lote_fornecedor", "") or "").strip().lower()
        local_txt = str(line.get("localizacao", "") or "").strip().lower()
        family_txt = ""
        if str(line.get("material_familia", "") or "").strip():
            family_txt = str(self.material_family_profile(line.get("material", ""), line.get("material_familia", "")).get("key", "") or "").strip()
        probe = self.material_geometry_preview(line)
        if not material_txt:
            return None
        records = [record for record in list(self.ensure_data().get("materiais", []) or []) if isinstance(record, dict)]
        if material_id:
            records.sort(key=lambda record: 0 if str(record.get("id", "") or "").strip() == material_id else 1)
        for record in records:
            if self._norm_material_token(record.get("material", "")) != material_txt:
                continue
            if self._norm_esp_token(record.get("espessura", "")) != esp_txt:
                continue
            if str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "").strip() != formato_txt:
                continue
            if lote_txt and str(record.get("lote_fornecedor", "") or "").strip().lower() != lote_txt:
                continue
            if local_txt and self._localizacao(record).strip().lower() != local_txt:
                continue
            rec_family = ""
            if str(record.get("material_familia", "") or "").strip():
                rec_family = str(self.material_family_profile(record.get("material", ""), record.get("material_familia", "")).get("key", "") or "").strip()
            if family_txt and rec_family and rec_family != family_txt:
                continue
            candidate = self.material_geometry_preview(record)
            if formato_txt == "Tubo":
                if str(candidate.get("secao_tipo", "") or "").strip() != str(probe.get("secao_tipo", "") or "").strip():
                    continue
                if abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
                if str(probe.get("secao_tipo", "") or "").strip() == "redondo":
                    if abs(float(candidate.get("diametro", 0) or 0) - float(probe.get("diametro", 0) or 0)) > 1e-6:
                        continue
                else:
                    if abs(float(candidate.get("comprimento", 0) or 0) - float(probe.get("comprimento", 0) or 0)) > 1e-6:
                        continue
                    if abs(float(candidate.get("largura", 0) or 0) - float(probe.get("largura", 0) or 0)) > 1e-6:
                        continue
            elif formato_txt == "Perfil":
                if str(candidate.get("secao_tipo", "") or "").strip() != str(probe.get("secao_tipo", "") or "").strip():
                    continue
                if abs(float(candidate.get("altura", 0) or 0) - float(probe.get("altura", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("kg_m", 0) or 0) - float(probe.get("kg_m", 0) or 0)) > 1e-4:
                    continue
            else:
                if abs(float(candidate.get("comprimento", 0) or 0) - float(probe.get("comprimento", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("largura", 0) or 0) - float(probe.get("largura", 0) or 0)) > 1e-6:
                    continue
            return record
        return None

    def _create_material_placeholder_from_note_line(
        self,
        line: dict[str, Any],
        note_number: str,
        *,
        quantity: float = 0.0,
        lote_override: str = "",
        localizacao_override: str = "",
    ) -> dict[str, Any] | None:
        data = self.ensure_data()
        material_txt = str(line.get("material", "") or "").strip()
        if not material_txt:
            return None
        formato_txt = str(line.get("formato", "") or "Chapa").strip() or "Chapa"
        lote_txt = str(lote_override or line.get("lote_fornecedor", "") or "").strip()
        local_txt = str(localizacao_override or line.get("localizacao", "") or "").strip()
        geometry = self.material_geometry_preview(line)
        record = {
            "id": self._next_material_id(),
            "formato": formato_txt,
            "material": material_txt,
            "material_familia": str(line.get("material_familia", "") or "").strip(),
            "espessura": str(line.get("espessura", "") or "").strip(),
            "comprimento": self._parse_dimension_mm(geometry.get("comprimento", line.get("comprimento", 0)), 0),
            "largura": self._parse_dimension_mm(geometry.get("largura", line.get("largura", 0)), 0),
            "altura": self._parse_dimension_mm(geometry.get("altura", line.get("altura", 0)), 0),
            "diametro": self._parse_dimension_mm(geometry.get("diametro", line.get("diametro", 0)), 0),
            "metros": self._parse_float(geometry.get("metros", line.get("metros", 0)), 0),
            "kg_m": self._parse_float(geometry.get("kg_m", line.get("kg_m", 0)), 0),
            "quantidade": max(0.0, self._parse_float(quantity, 0)),
            "reservado": 0.0,
            "Localização": local_txt,
            "Localizacao": local_txt,
            "lote_fornecedor": lote_txt,
            "secao_tipo": str(geometry.get("secao_tipo", line.get("secao_tipo", "")) or "").strip(),
            "peso_unid": self._parse_float(geometry.get("peso_unid", line.get("peso_unid", 0)), 0),
            "p_compra": self._parse_float(line.get("p_compra", 0), 0),
            "contorno_points": [],
            "is_sobra": False,
            "atualizado_em": self.desktop_main.now_iso(),
        }
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record = self.materia_actions._hydrate_retalho_record(data, record)
        data.setdefault("materiais", []).append(record)
        self.desktop_main.push_unique(data.setdefault("materiais_hist", []), material_txt)
        if str(record.get("espessura", "") or "").strip():
            self.desktop_main.push_unique(data.setdefault("espessuras_hist", []), str(record.get("espessura", "") or "").strip())
        self.desktop_main.log_stock(data, "CRIAR_NE", f"{record.get('id', '')} via {note_number}")
        return record

    def _ne_normalize_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        origem = str(payload.get("origem", "Produto") or "Produto").strip() or "Produto"
        ref = str(payload.get("ref", "") or "").strip()
        descricao = str(payload.get("descricao", "") or "").strip()
        fornecedor_linha = str(payload.get("fornecedor_linha", "") or "").strip()
        unid = str(payload.get("unid", "UN") or "UN").strip() or "UN"
        qtd = self._parse_float(payload.get("qtd", 0), 0)
        preco = self._parse_float(payload.get("preco", 0), 0)
        desconto = max(0.0, min(100.0, self._parse_float(payload.get("desconto", 0), 0)))
        iva = max(0.0, min(100.0, self._parse_float(payload.get("iva", 23), 23)))
        if not descricao:
            raise ValueError("Descrição da linha obrigatória.")
        if qtd <= 0:
            raise ValueError("Quantidade da linha inválida.")
        base = (qtd * preco) * (1.0 - (desconto / 100.0))
        iva_amt = base * (iva / 100.0)
        total = round(base + iva_amt, 4)
        line = {
            "ref": ref,
            "descricao": descricao,
            "fornecedor_linha": fornecedor_linha,
            "origem": origem,
            "qtd": qtd,
            "unid": unid,
            "preco": preco,
            "total": total,
            "desconto": desconto,
            "iva": iva,
            "entregue": bool(payload.get("entregue")),
            "qtd_entregue": self._parse_float(payload.get("qtd_entregue", qtd if payload.get("entregue") else 0), 0),
        }
        if self.desktop_main.origem_is_materia(origem):
            material = self.material_by_id(ref)
            if material:
                metrics = self.material_geometry_preview(material)
                line.update(
                    {
                        "material": material.get("material", ""),
                        "espessura": material.get("espessura", ""),
                        "comprimento": self._parse_dimension_mm(metrics.get("comprimento", material.get("comprimento", 0)), 0),
                        "largura": self._parse_dimension_mm(metrics.get("largura", material.get("largura", 0)), 0),
                        "altura": self._parse_dimension_mm(metrics.get("altura", material.get("altura", 0)), 0),
                        "diametro": self._parse_dimension_mm(metrics.get("diametro", material.get("diametro", 0)), 0),
                        "metros": self._parse_float(metrics.get("metros", material.get("metros", 0)), 0),
                        "kg_m": self._parse_float(metrics.get("kg_m", material.get("kg_m", 0)), 0),
                        "localizacao": self._localizacao(material),
                        "lote_fornecedor": material.get("lote_fornecedor", ""),
                        "peso_unid": self._parse_float(metrics.get("peso_unid", material.get("peso_unid", 0)), 0),
                        "p_compra": self._parse_float(material.get("p_compra", 0), 0),
                        "formato": material.get("formato", self.desktop_main.detect_materia_formato(material)),
                        "secao_tipo": str(metrics.get("secao_tipo", material.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(material.get("material_familia", "") or "").strip(),
                        "_material_pending_create": False,
                        "_material_manual": False,
                    }
                )
            else:
                formato_txt = str(payload.get("formato", "") or "Chapa").strip() or "Chapa"
                material_txt = str(payload.get("material", "") or "").strip()
                esp_txt = str(payload.get("espessura", "") or "").strip()
                if not material_txt:
                    raise ValueError("Qualidade da matéria-prima obrigatória.")
                if formato_txt in {"Chapa", "Tubo"} and not esp_txt:
                    raise ValueError("Espessura obrigatória para chapa e tubo.")
                metrics = self.material_geometry_preview(payload)
                line.update(
                    {
                        "material": material_txt,
                        "espessura": esp_txt,
                        "comprimento": self._parse_dimension_mm(metrics.get("comprimento", payload.get("comprimento", 0)), 0),
                        "largura": self._parse_dimension_mm(metrics.get("largura", payload.get("largura", 0)), 0),
                        "altura": self._parse_dimension_mm(metrics.get("altura", payload.get("altura", 0)), 0),
                        "diametro": self._parse_dimension_mm(metrics.get("diametro", payload.get("diametro", 0)), 0),
                        "metros": self._parse_float(metrics.get("metros", payload.get("metros", 0)), 0),
                        "kg_m": self._parse_float(metrics.get("kg_m", payload.get("kg_m", 0)), 0),
                        "localizacao": str(payload.get("localizacao", "") or "").strip(),
                        "lote_fornecedor": str(payload.get("lote_fornecedor", "") or "").strip(),
                        "peso_unid": self._parse_float(metrics.get("peso_unid", payload.get("peso_unid", 0)), 0),
                        "p_compra": self._parse_float(payload.get("p_compra", 0), 0),
                        "formato": formato_txt,
                        "secao_tipo": str(metrics.get("secao_tipo", payload.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(payload.get("material_familia", "") or "").strip(),
                        "_material_pending_create": bool(payload.get("_material_pending_create", True)),
                        "_material_manual": bool(payload.get("_material_manual", True)),
                    }
                )
        return line

    def ne_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip() or self.desktop_main.next_ne_numero(data)
        fornecedor_id = str(payload.get("fornecedor_id", "") or "").strip()
        fornecedor = str(payload.get("fornecedor", "") or "").strip()
        contacto = str(payload.get("contacto", "") or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        obs = str(payload.get("obs", "") or "").strip()
        local_descarga = str(payload.get("local_descarga", "") or "").strip()
        meio_transporte = str(payload.get("meio_transporte", "") or "").strip()
        lines_payload = list(payload.get("lines", []) or [])
        normalized_lines = [self._ne_normalize_line(line) for line in lines_payload]
        resolved_supplier_id, resolved_supplier_text, resolved_contact = self._normalize_supplier_reference(fornecedor_id, fornecedor)
        fornecedor_id = resolved_supplier_id
        fornecedor = resolved_supplier_text
        if not contacto and resolved_contact:
            contacto = resolved_contact
        for line in normalized_lines:
            line_supplier = str(line.get("fornecedor_linha", "") or "").strip()
            if not line_supplier:
                continue
            _, resolved_line_supplier, _ = self._normalize_supplier_reference("", line_supplier)
            if resolved_line_supplier:
                line["fornecedor_linha"] = resolved_line_supplier
        lst = data.setdefault("notas_encomenda", [])
        existing = next((row for row in lst if str(row.get("numero", "") or "").strip() == numero), None)
        old_lines = list(existing.get("linhas", []) or []) if isinstance(existing, dict) else []
        if not fornecedor and normalized_lines:
            unique_suppliers = {
                str(line.get("fornecedor_linha", "") or "").strip()
                for line in normalized_lines
                if str(line.get("fornecedor_linha", "") or "").strip()
            }
            if len(unique_suppliers) == 1:
                fornecedor = next(iter(unique_suppliers))
                fornecedor_id, fornecedor, inferred_contact = self._resolve_supplier(fornecedor)
                if not contacto:
                    contacto = inferred_contact
        note = {
            "numero": numero,
            "fornecedor": fornecedor,
            "fornecedor_id": fornecedor_id,
            "contacto": contacto,
            "data_entrega": data_entrega,
            "obs": obs,
            "local_descarga": local_descarga,
            "meio_transporte": meio_transporte,
            "linhas": normalized_lines,
            "estado": str((existing or {}).get("estado", "Em edicao") or "Em edicao").strip() or "Em edicao",
            "oculta": bool((existing or {}).get("oculta", False)),
            "_draft": False,
            "origem_cotacao": str((existing or {}).get("origem_cotacao", "") or "").strip(),
            "ne_geradas": list((existing or {}).get("ne_geradas", []) or []),
            "entregas": list((existing or {}).get("entregas", []) or []),
            "documentos": list((existing or {}).get("documentos", []) or []),
            "guia_ultima": str((existing or {}).get("guia_ultima", "") or "").strip(),
            "fatura_ultima": str((existing or {}).get("fatura_ultima", "") or "").strip(),
            "fatura_caminho_ultima": str((existing or {}).get("fatura_caminho_ultima", "") or "").strip(),
            "data_doc_ultima": str((existing or {}).get("data_doc_ultima", "") or "").strip(),
            "data_ultima_entrega": str((existing or {}).get("data_ultima_entrega", "") or "").strip(),
            "data_entregue": str((existing or {}).get("data_entregue", "") or "").strip(),
            "data_aprovacao": str((existing or {}).get("data_aprovacao", "") or "").strip(),
        }
        for index, line in enumerate(note["linhas"]):
            if index >= len(old_lines):
                if line.get("entregue"):
                    line["qtd_entregue"] = self._parse_float(line.get("qtd", 0), 0)
                continue
            old = old_lines[index]
            qtd_tot = self._parse_float(line.get("qtd", 0), 0)
            qtd_old = self._parse_float(old.get("qtd_entregue", old.get("qtd", 0) if old.get("entregue") else 0), 0)
            qtd_old = max(0.0, min(qtd_tot, qtd_old))
            line["qtd_entregue"] = qtd_old
            if old.get("entregue") or (qtd_tot > 0 and qtd_old >= (qtd_tot - 1e-9)):
                line["entregue"] = True
            if old.get("_stock_in") and qtd_old > 0:
                line["_stock_in"] = True
            for key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                if old.get(key):
                    line[key] = old.get(key)
        note_kind = self._note_kind(note)
        if note_kind == "rfq":
            note["fornecedor"] = ""
            note["fornecedor_id"] = ""
            note["contacto"] = ""
        elif not str(note.get("fornecedor_id", "") or "").strip() and str(note.get("fornecedor", "") or "").strip():
            raise ValueError("Seleciona um fornecedor válido da ficha de fornecedores.")
        product_changed = False
        material_changed = False
        for line in list(note.get("linhas", []) or []):
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                material_changed = self._update_materia_preco_from_unit(line.get("ref", ""), line.get("preco", 0)) or material_changed
            else:
                product_changed = self._update_produto_preco_from_unit(line.get("ref", ""), line.get("preco", 0)) or product_changed
        if material_changed:
            self._sync_ne_from_materia()
        if product_changed:
            self._sync_ne_from_products()
        self._recalculate_note_totals(note)
        if existing:
            existing.update(note)
        else:
            lst.append(note)
        self._save(force=True)
        return note

    def ne_create_draft(self) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(self.desktop_main.next_ne_numero(data))
        note = {
            "numero": numero,
            "fornecedor": "",
            "fornecedor_id": "",
            "contacto": "",
            "data_entrega": "",
            "obs": "",
            "local_descarga": "",
            "meio_transporte": "",
            "linhas": [],
            "total": 0.0,
            "estado": "Em edicao",
            "oculta": False,
            "_draft": True,
            "entregas": [],
            "documentos": [],
            "guia_ultima": "",
            "fatura_ultima": "",
            "fatura_caminho_ultima": "",
            "data_doc_ultima": "",
            "data_ultima_entrega": "",
        }
        data.setdefault("notas_encomenda", []).append(note)
        self._save(force=True)
        return note

    def ne_remove(self, numero: str) -> None:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        before = len(list(data.get("notas_encomenda", []) or []))
        data["notas_encomenda"] = [row for row in list(data.get("notas_encomenda", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        if len(data["notas_encomenda"]) == before:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        self._save(force=True)

    def ne_approve(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        if not list(note.get("linhas", []) or []):
            raise ValueError("A nota n?o tem linhas.")
        note_kind = self._note_kind(note)
        note["estado"] = "Aprovada" if note_kind == "purchase_note" else "Cotacao aprovada"
        note["data_aprovacao"] = self.desktop_main.now_iso()
        note["_draft"] = False
        self._save(force=True)
        return note

    def ne_mark_sent(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        note["estado"] = "Enviada"
        note["data_envio"] = self.desktop_main.now_iso()
        note["_draft"] = False
        self._save(force=True)
        return note

    def ne_generate_supplier_orders(self, numero: str) -> list[dict[str, Any]]:
        data = self.ensure_data()
        number = str(numero or "").strip()
        note = next((row for row in list(data.get("notas_encomenda", []) or []) if str(row.get("numero", "") or "").strip() == number), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        lines = list(note.get("linhas", []) or [])
        if not lines:
            raise ValueError("A nota n?o tem linhas.")
        groups: dict[str, dict[str, Any]] = {}
        missing: list[str] = []
        for line in lines:
            target_supplier = str(line.get("fornecedor_linha", "") or note.get("fornecedor", "") or "").strip()
            supplier_id, supplier_text, supplier_contact = self._resolve_supplier(target_supplier)
            if not supplier_text:
                missing.append(str(line.get("ref", "") or "").strip() or str(line.get("descricao", "") or "").strip())
                continue
            key = supplier_id or supplier_text
            if key not in groups:
                groups[key] = {
                    "fornecedor_id": supplier_id,
                    "fornecedor": supplier_text,
                    "contacto": supplier_contact,
                    "linhas": [],
                }
            new_line = dict(line)
            new_line["fornecedor_linha"] = supplier_text
            new_line["entregue"] = False
            new_line["qtd_entregue"] = 0.0
            new_line["_stock_in"] = False
            for transient_key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                new_line.pop(transient_key, None)
            groups[key]["linhas"].append(new_line)
        if missing:
            raise ValueError("Existem linhas sem fornecedor adjudicado: " + ", ".join(sorted(set(item for item in missing if item))))
        if not groups:
            raise ValueError("Nao existem fornecedores adjudicados para gerar NEs.")
        created: list[dict[str, Any]] = []
        notes = data.setdefault("notas_encomenda", [])
        for group in groups.values():
            new_number = str(self.desktop_main.next_ne_numero(data))
            new_note = {
                "numero": new_number,
                "fornecedor": group.get("fornecedor", ""),
                "fornecedor_id": group.get("fornecedor_id", ""),
                "contacto": group.get("contacto", ""),
                "data_entrega": str(note.get("data_entrega", "") or "").strip(),
                "obs": f"Gerada de {note.get('numero', '')}".strip(),
                "local_descarga": str(note.get("local_descarga", "") or "").strip(),
                "meio_transporte": str(note.get("meio_transporte", "") or "").strip(),
                "linhas": list(group.get("linhas", []) or []),
                "estado": "Aprovada",
                "_draft": False,
                "oculta": False,
                "origem_cotacao": str(note.get("numero", "") or "").strip(),
                "ne_geradas": [],
                "entregas": [],
                "documentos": [],
                "guia_ultima": "",
                "fatura_ultima": "",
                "fatura_caminho_ultima": "",
                "data_doc_ultima": "",
                "data_ultima_entrega": "",
            }
            self._recalculate_note_totals(new_note)
            notes.append(new_note)
            created.append({"numero": new_number, "fornecedor": new_note["fornecedor"], "total": new_note["total"]})
        note["estado"] = "Convertida"
        note["oculta"] = True
        note["_draft"] = False
        note["ne_geradas"] = [row["numero"] for row in created]
        self._save(force=True)
        return created

    def ne_render_pdf(self, numero: str, quote: bool = False, output_path: str | Path = "") -> Path:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        suffix = "_cotacao" if quote else ""
        path = Path(output_path) if str(output_path or "").strip() else Path(tempfile.gettempdir()) / f"lugest_ne_{numero}{suffix}.pdf"
        if quote:
            self.ne_expedicao_actions.render_ne_cotacao_pdf(self, str(path), note)
        else:
            self.ne_expedicao_actions.render_ne_pdf(self, str(path), note)
        return path

    def ne_open_pdf(self, numero: str, quote: bool = False) -> Path:
        path = self.ne_render_pdf(numero, quote=quote)
        os.startfile(str(path))
        return path

    def ne_documents(self, numero: str) -> list[dict[str, Any]]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        return self._ne_document_rows(note)

    def ne_add_document(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        raw = dict(payload or {})
        if not any(str(raw.get(key, "") or "").strip() for key in ("titulo", "guia", "fatura", "caminho", "obs")):
            raise ValueError("Indica pelo menos titulo, guia, fatura, caminho ou observacao.")
        apply_to_lines = bool(raw.get("apply_to_lines"))
        register_history = bool(raw.get("register_history", True))
        doc = self._ne_normalize_document(
            {
                "data_registo": raw.get("data_registo") or self.desktop_main.now_iso(),
                "tipo": raw.get("tipo", ""),
                "titulo": raw.get("titulo", ""),
                "caminho": raw.get("caminho", ""),
                "guia": raw.get("guia", ""),
                "fatura": raw.get("fatura", ""),
                "data_entrega": raw.get("data_entrega", ""),
                "data_documento": raw.get("data_documento", ""),
                "obs": raw.get("obs", ""),
            }
        )
        stored_doc = {
            key: doc.get(key, "")
            for key in ("data_registo", "tipo", "titulo", "caminho", "guia", "fatura", "data_entrega", "data_documento", "obs")
        }
        note.setdefault("documentos", []).append(stored_doc)
        if apply_to_lines:
            for line in list(note.get("linhas", []) or []):
                qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
                qtd_ent = max(0.0, self._parse_float(line.get("qtd_entregue", qtd_total if line.get("entregue") else 0), 0))
                if qtd_ent <= 0 and not bool(line.get("entregue")):
                    continue
                if doc.get("guia"):
                    line["guia_entrega"] = doc["guia"]
                if doc.get("fatura"):
                    line["fatura_entrega"] = doc["fatura"]
                if doc.get("data_documento"):
                    line["data_doc_entrega"] = doc["data_documento"]
                if doc.get("data_entrega"):
                    line["data_entrega_real"] = doc["data_entrega"]
                if doc.get("obs"):
                    line["obs_entrega"] = doc["obs"]
        if register_history:
            note.setdefault("entregas", []).append(
                {
                    "data_registo": doc.get("data_registo", ""),
                    "data_entrega": doc.get("data_entrega", ""),
                    "guia": doc.get("guia", ""),
                    "fatura": doc.get("fatura", ""),
                    "data_documento": doc.get("data_documento", ""),
                    "obs": doc.get("obs", ""),
                    "linhas": [],
                    "quantidade_linhas": 0,
                    "quantidade_total": 0,
                    "tipo": doc.get("tipo", "DOCUMENTO"),
                    "titulo": doc.get("titulo", ""),
                    "caminho": doc.get("caminho", ""),
                }
            )
        if doc.get("guia"):
            note["guia_ultima"] = doc["guia"]
        if doc.get("fatura"):
            note["fatura_ultima"] = doc["fatura"]
        if doc.get("data_documento"):
            note["data_doc_ultima"] = doc["data_documento"]
        if doc.get("data_entrega"):
            note["data_ultima_entrega"] = doc["data_entrega"]
        if doc.get("caminho") and (doc.get("fatura") or "FATURA" in str(doc.get("tipo", "") or "").upper()):
            note["fatura_caminho_ultima"] = doc["caminho"]
        self._save(force=True)
        return self.ne_detail(numero)

    def ne_register_delivery(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == str(numero or "").strip()), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        note_lines = list(note.get("linhas", []) or [])
        line_updates_by_index: dict[int, dict[str, Any]] = {}
        unresolved_legacy_items: list[dict[str, Any]] = []
        for item in list(payload.get("lines", []) or []):
            if not isinstance(item, dict):
                continue
            try:
                line_index = int(item.get("index"))
            except Exception:
                line_index = -1
            if line_index >= 0:
                line_updates_by_index[line_index] = dict(item)
                continue
            unresolved_legacy_items.append(dict(item))
        if unresolved_legacy_items:
            used_indexes = set(line_updates_by_index.keys())
            for item in unresolved_legacy_items:
                ref = str(item.get("ref", "") or "").strip()
                descricao = str(item.get("descricao", "") or "").strip().casefold()
                origem = str(item.get("origem", "") or "").strip().casefold()
                if not ref and not descricao:
                    continue
                matches: list[int] = []
                for idx, line in enumerate(note_lines):
                    if idx in used_indexes:
                        continue
                    line_ref = str(line.get("ref", "") or "").strip()
                    line_desc = str(line.get("descricao", "") or "").strip().casefold()
                    line_origin = str(line.get("origem", "") or "").strip().casefold()
                    if ref and line_ref != ref:
                        continue
                    if descricao and line_desc != descricao:
                        continue
                    if origem and line_origin != origem:
                        continue
                    matches.append(idx)
                if len(matches) == 1:
                    match_index = matches[0]
                    line_updates_by_index[match_index] = dict(item) | {"index": match_index}
                    used_indexes.add(match_index)
        if not line_updates_by_index:
            raise ValueError("Seleciona pelo menos uma linha para entregar.")
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        data_documento = str(payload.get("data_documento", "") or "").strip()
        guia = str(payload.get("guia", "") or "").strip()
        fatura = str(payload.get("fatura", "") or "").strip()
        titulo = str(payload.get("titulo", "") or "").strip()
        caminho = str(payload.get("caminho", "") or "").strip()
        obs = str(payload.get("obs", "") or "").strip()
        registo_ts = self.desktop_main.now_iso()
        any_delivery = False
        delivered_lines: list[str] = []
        total_qtd = 0.0
        for line_index, line in enumerate(note_lines):
            update = line_updates_by_index.get(line_index)
            if update is None:
                continue
            ref = str(line.get("ref", "") or "").strip()
            add_qtd = max(0.0, self._parse_float(update.get("qtd", 0), 0))
            if add_qtd <= 0:
                continue
            lote_override = str(update.get("lote_fornecedor", "") or "").strip()
            local_override = str(update.get("localizacao", "") or "").strip()
            qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            qtd_old = max(0.0, self._parse_float(line.get("qtd_entregue", qtd_total if line.get("entregue") else 0), 0))
            qtd_left = max(0.0, qtd_total - qtd_old)
            qty_apply = min(qtd_left, add_qtd)
            if qty_apply <= 0:
                continue
            working_line = dict(line)
            # Lote e localizacao da entrega sao sempre decididos nesta operacao.
            working_line["lote_fornecedor"] = lote_override
            working_line["localizacao"] = local_override
            qtd_new = qtd_old + qty_apply
            line["qtd_entregue"] = qtd_new
            line["entregue"] = qtd_new >= (qtd_total - 1e-9)
            line["data_entrega_real"] = data_entrega
            line["data_doc_entrega"] = data_documento
            line["guia_entrega"] = guia
            line["fatura_entrega"] = fatura
            line["obs_entrega"] = obs
            line.setdefault("entregas_linha", []).append(
                {
                    "data_registo": registo_ts,
                    "data_entrega": data_entrega,
                    "data_documento": data_documento,
                    "guia": guia,
                    "fatura": fatura,
                    "obs": obs,
                    "qtd": qty_apply,
                    "lote_fornecedor": lote_override,
                    "localizacao": local_override,
                }
            )
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                material = self._find_existing_material_from_note_line(working_line)
                if material is None:
                    material = self._create_material_placeholder_from_note_line(
                        working_line,
                        str(note.get("numero", "") or "").strip(),
                        quantity=qty_apply,
                        lote_override=lote_override,
                        localizacao_override=local_override,
                    )
                    if material is None:
                        raise ValueError(
                            "A linha de matéria-prima não tem dados suficientes para criar stock: "
                            + (str(line.get("descricao", "") or "").strip() or str(line.get("material", "") or "").strip() or "-")
                            + "."
                        )
                else:
                    material["quantidade"] = self._parse_float(material.get("quantidade", 0), 0) + qty_apply
                    if lote_override:
                        material["lote_fornecedor"] = lote_override
                    if local_override:
                        material["Localização"] = local_override
                        material["Localizacao"] = local_override
                    material["atualizado_em"] = registo_ts
                if material is not None:
                    ref = str(material.get("id", "") or "").strip()
                    line["ref"] = ref
                    line["material"] = str(material.get("material", "") or "").strip()
                    line["espessura"] = str(material.get("espessura", "") or "").strip()
                    line["comprimento"] = self._parse_dimension_mm(material.get("comprimento", 0), 0)
                    line["largura"] = self._parse_dimension_mm(material.get("largura", 0), 0)
                    line["metros"] = self._parse_float(material.get("metros", 0), 0)
                    line["localizacao"] = self._localizacao(material)
                    line["lote_fornecedor"] = str(material.get("lote_fornecedor", "") or "").strip()
                    line["peso_unid"] = self._parse_float(material.get("peso_unid", 0), 0)
                    line["p_compra"] = self._parse_float(material.get("p_compra", 0), 0)
                    line["formato"] = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "").strip()
                    line["_material_pending_create"] = False
                    line["_material_manual"] = False
            else:
                product = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == ref), None)
                if product is not None:
                    before = self._parse_float(product.get("qty", 0), 0)
                    product["qty"] = before + qty_apply
                    product["atualizado_em"] = registo_ts
                    self.desktop_main.add_produto_mov(
                        self.ensure_data(),
                        tipo="Entrada",
                        operador=str((self.user or {}).get("username", "") or "Sistema"),
                        codigo=ref,
                        descricao=str(product.get("descricao", "") or "").strip(),
                        qtd=qty_apply,
                        antes=before,
                        depois=product["qty"],
                        obs=f"NE {note.get('numero', '')}",
                        origem="Notas Encomenda",
                        ref_doc=str(note.get("numero", "") or "").strip(),
                    )
            line["_stock_in"] = True
            any_delivery = True
            total_qtd += qty_apply
            delivered_lines.append(f"{(ref or str(line.get('descricao', '') or '-').strip())} ({self._fmt(qty_apply)})")
        if not any_delivery:
            raise ValueError("Nao foi possivel registar entrega para as quantidades indicadas.")
        note.setdefault("entregas", []).append(
            {
                "data_registo": registo_ts,
                "data_entrega": data_entrega,
                "guia": guia,
                "fatura": fatura,
                "data_documento": data_documento,
                "obs": obs,
                "linhas": delivered_lines,
                "quantidade_linhas": len(delivered_lines),
                "quantidade_total": total_qtd,
            }
        )
        for line in list(note.get("linhas", []) or []):
            qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            delivered_qty = 0.0
            for movement in list(line.get("entregas_linha", []) or []):
                delivered_qty += max(0.0, self._parse_float(movement.get("qtd", 0), 0))
            if delivered_qty > qtd_total > 0:
                delivered_qty = qtd_total
            line["qtd_entregue"] = delivered_qty
            line["entregue"] = bool(qtd_total > 0 and delivered_qty >= (qtd_total - 1e-9))
            line["_stock_in"] = bool(delivered_qty > 0)
        note["guia_ultima"] = guia
        note["fatura_ultima"] = fatura
        note["data_doc_ultima"] = data_documento
        note["data_ultima_entrega"] = data_entrega
        if caminho and fatura:
            note["fatura_caminho_ultima"] = caminho
        if any(str(value or "").strip() for value in (titulo, caminho, guia, fatura, obs)):
            note.setdefault("documentos", []).append(
                {
                    "data_registo": registo_ts,
                    "tipo": "ENTREGA",
                    "titulo": self._ne_document_title(
                        {
                            "titulo": titulo,
                            "guia": guia,
                            "fatura": fatura,
                            "caminho": caminho,
                            "data_entrega": data_entrega,
                            "data_documento": data_documento,
                        },
                        doc_type="ENTREGA",
                    ),
                    "caminho": caminho,
                    "guia": guia,
                    "fatura": fatura,
                    "data_entrega": data_entrega,
                    "data_documento": data_documento,
                    "obs": obs,
                }
            )
        all_lines = list(note.get("linhas", []) or [])
        if all(bool(line.get("entregue")) for line in all_lines):
            note["estado"] = "Entregue"
            note["data_entregue"] = data_entrega
        elif any(self._parse_float(line.get("qtd_entregue", 0), 0) > 0 for line in all_lines):
            note["estado"] = "Parcial"
        else:
            note["estado"] = "Aprovada"
        self._save(force=True)
        return self.ne_detail(str(note.get("numero", "") or ""))

    def _peek_next_orc_number(self) -> str:
        data = self.ensure_data()
        try:
            seq = int(data.get("orc_seq", 1) or 1)
        except Exception:
            seq = 1
        year = int(getattr(self.desktop_main.datetime.now(), "year", 0) or 0)
        return f"ORC-{year}-{seq:04d}"

    def _orc_number_sort_key(self, numero: str) -> tuple[int, int, str]:
        raw = str(numero or "").strip()
        parts = raw.split("-")
        year = 0
        seq = 0
        if len(parts) >= 3:
            try:
                year = int(parts[1])
            except Exception:
                year = 0
            try:
                seq = int(parts[2])
            except Exception:
                seq = 0
        return (year, seq, raw)

    def orc_next_number(self) -> str:
        return self._peek_next_orc_number()

    def _normalize_orc_client(self, value: Any) -> dict[str, str]:
        return dict(self.desktop_main._normalize_orc_cliente(value, self.ensure_data()) or {})

    def orc_rows(self, filter_text: str = "", state_filter: str = "Ativas", year: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        year_raw = str(year or "Todos").strip()
        rows: list[dict[str, Any]] = []
        for raw in list(data.get("orcamentos", []) or []):
            if not isinstance(raw, dict):
                continue
            orc = dict(raw)
            client = self._normalize_orc_client(orc.get("cliente", {}))
            estado = str(orc.get("estado", "") or "").strip() or "Em edicao"
            estado_norm = self.desktop_main.norm_text(estado)
            row_year = str(self.orc_actions._orc_extract_year(orc.get("data", ""), orc.get("numero", ""), orc.get("ano")) or "").strip()
            if year_raw and year_raw.lower() not in {"todos", "todas", "all"} and row_year != year_raw:
                continue
            if state_raw and state_raw not in {"todos", "todas", "all"}:
                if "ativ" in state_raw and ("rejeitado" in estado_norm or "convertido" in estado_norm):
                    continue
                if "edi" in state_raw and "edi" not in estado_norm:
                    continue
                if "enviado" in state_raw and "enviado" not in estado_norm:
                    continue
                if "aprovado" in state_raw and "aprovado" not in estado_norm:
                    continue
                if "rejeitado" in state_raw and "rejeitado" not in estado_norm:
                    continue
                if "convertido" in state_raw and "convertido" not in estado_norm:
                    continue
            client_label = f"{client.get('codigo', '')} - {client.get('nome', '')}".strip(" -")
            row = {
                "numero": str(orc.get("numero", "") or "").strip(),
                "cliente": client_label or str(client.get("nome", "") or "").strip() or str(orc.get("cliente", "") or "").strip(),
                "estado": estado,
                "numero_encomenda": str(orc.get("numero_encomenda", "") or "").strip(),
                "total": round(self._parse_float(orc.get("total", 0), 0), 2),
                "data": str(orc.get("data", "") or "").strip()[:10],
                "linhas": len(list(orc.get("linhas", []) or [])),
                "ano": row_year,
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: self._orc_number_sort_key(str(item.get("numero", "") or "")), reverse=True)
        return rows

    def orc_available_years(self) -> list[str]:
        current_year = str(self.desktop_main.datetime.now().year)
        years = {current_year}
        for row in list(self.ensure_data().get("orcamentos", []) or []):
            if not isinstance(row, dict):
                continue
            year = str(self.orc_actions._orc_extract_year(row.get("data", ""), row.get("numero", ""), row.get("ano")) or "").strip()
            if year:
                years.add(year)
        return sorted(years, key=lambda value: int(value) if value.isdigit() else 0, reverse=True)

    def _find_orc_record(self, numero: str) -> dict[str, Any] | None:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("orcamentos", []) or [])
                if str(row.get("numero", "") or "").strip() == numero_txt
            ),
            None,
        )

    def _json_safe_clone(self, payload: Any) -> Any:
        try:
            return json.loads(json.dumps(payload, ensure_ascii=False, default=str))
        except Exception:
            if isinstance(payload, dict):
                return {str(key): self._json_safe_clone(value) for key, value in payload.items()}
            if isinstance(payload, (list, tuple, set)):
                return [self._json_safe_clone(value) for value in payload]
            return payload

    def _ensure_orc_nesting_studies_table(self, conn: Any) -> None:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS orc_nesting_studies (
                    quote_number VARCHAR(80) NOT NULL,
                    group_key VARCHAR(190) NOT NULL,
                    group_label VARCHAR(255) NULL,
                    study_json LONGTEXT NULL,
                    created_at DATETIME NULL,
                    updated_at DATETIME NULL,
                    PRIMARY KEY (quote_number, group_key)
                )
                """
            )

    def _mysql_orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return {}
        conn = None
        studies: dict[str, Any] = {}
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return {}
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT group_key, study_json
                    FROM orc_nesting_studies
                    WHERE quote_number=%s
                    ORDER BY updated_at DESC, group_key ASC
                    """,
                    (numero_txt,),
                )
                rows = list(cur.fetchall() or [])
            for row in rows:
                group_key = str((row.get("group_key") if isinstance(row, dict) else row[0]) or "").strip()
                raw = row.get("study_json") if isinstance(row, dict) else row[1]
                if not group_key:
                    continue
                if isinstance(raw, (bytes, bytearray)):
                    raw = raw.decode("utf-8", errors="ignore")
                try:
                    parsed = json.loads(str(raw or "{}"))
                except Exception:
                    parsed = {}
                if isinstance(parsed, dict):
                    studies[group_key] = parsed
        except Exception:
            studies = {}
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
        return studies

    def _mysql_save_orc_nesting_study(self, numero: str, group_key: str, group_label: str, payload: dict[str, Any]) -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt or not group_key_txt:
            return
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            clean = json.dumps(self._json_safe_clone(payload), ensure_ascii=False)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO orc_nesting_studies (
                        quote_number,
                        group_key,
                        group_label,
                        study_json,
                        created_at,
                        updated_at
                    )
                    VALUES (%s, %s, %s, %s, NOW(), NOW())
                    ON DUPLICATE KEY UPDATE
                        group_label=VALUES(group_label),
                        study_json=VALUES(study_json),
                        updated_at=VALUES(updated_at)
                    """,
                    (numero_txt, group_key_txt, str(group_label or "").strip(), clean),
                )
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def _mysql_delete_orc_nesting_studies(self, numero: str, group_key: str = "") -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt:
            return
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            with conn.cursor() as cur:
                if group_key_txt:
                    cur.execute(
                        "DELETE FROM orc_nesting_studies WHERE quote_number=%s AND group_key=%s",
                        (numero_txt, group_key_txt),
                    )
                else:
                    cur.execute("DELETE FROM orc_nesting_studies WHERE quote_number=%s", (numero_txt,))
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        orc = self._find_orc_record(numero_txt)
        local_studies = {}
        if isinstance(orc, dict):
            local_studies = {
                str(key): self._json_safe_clone(value)
                for key, value in dict(orc.get("nesting_studies", {}) or {}).items()
                if str(key).strip()
            }
        remote_studies = self._mysql_orc_nesting_studies(numero_txt)
        if not remote_studies:
            return local_studies
        merged = dict(local_studies)
        for group_key, remote_value in remote_studies.items():
            local_value = dict(merged.get(group_key, {}) or {})
            remote_updated = str(dict(remote_value or {}).get("updated_at", "") or "").strip()
            local_updated = str(local_value.get("updated_at", "") or "").strip()
            if not local_value or remote_updated >= local_updated:
                merged[group_key] = self._json_safe_clone(remote_value)
        return merged

    def orc_save_nesting_study(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        orc = self._find_orc_record(numero_txt)
        if orc is None:
            raise ValueError("Guarda primeiro o orçamento para associar o estudo de nesting.")
        clean = dict(self._json_safe_clone(payload) or {})
        group_key = str(clean.get("group_key", "") or "").strip()
        if not group_key:
            raise ValueError("Grupo de nesting inválido.")
        previous = dict(dict(orc.get("nesting_studies", {}) or {}).get(group_key, {}) or {})
        clean["quote_number"] = numero_txt
        clean["group_key"] = group_key
        clean["group_label"] = str(clean.get("group_label", previous.get("group_label", "")) or "").strip()
        clean["created_at"] = str(previous.get("created_at", "") or clean.get("created_at", "") or self.desktop_main.now_iso()).strip()
        clean["updated_at"] = self.desktop_main.now_iso()
        orc.setdefault("nesting_studies", {})[group_key] = clean
        orc["latest_nesting_bridge"] = dict(clean.get("quote_bridge", {}) or {})
        orc["latest_nesting_group_key"] = group_key
        orc["latest_nesting_updated_at"] = clean["updated_at"]
        self._save(force=True)
        try:
            self._mysql_save_orc_nesting_study(numero_txt, group_key, clean.get("group_label", ""), clean)
        except Exception:
            pass
        return self._json_safe_clone(clean)

    def orc_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        client = self._normalize_orc_client(orc.get("cliente", {}))
        lines: list[dict[str, Any]] = []
        for row in list(orc.get("linhas", []) or []):
            snapshot = self._quote_line_operation_snapshot(row, quote_number=numero, quote_state=str(orc.get("estado", "") or "").strip())
            raw_operacao = str(row.get("operacao", "") or "").strip()
            current_time = round(self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 4)
            current_price = round(self._parse_float(row.get("preco_unit", 0), 0), 4)
            derived_laser_base = (
                self.desktop_main.orc_line_is_piece(row)
                and bool(str(row.get("desenho", "") or "").strip())
                and "corte laser" in self.desktop_main.norm_text(raw_operacao)
                and (
                    current_time > 0
                    or current_price > 0
                )
            )
            laser_base_active = bool(row.get("laser_base_active", False) or derived_laser_base)
            laser_base_tempo = round(
                self._parse_float(
                    row.get(
                        "laser_base_tempo_unit",
                        current_time if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            laser_base_preco = round(
                self._parse_float(
                    row.get(
                        "laser_base_preco_unit",
                        current_price if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            display_extra_time_map = self._quote_collect_non_laser_map(dict(snapshot.get("tempos_operacao", {}) or {}), digits=4)
            display_extra_price_map = self._quote_collect_non_laser_map(dict(snapshot.get("custos_operacao", {}) or {}), digits=4)
            display_extra_time = round(sum(display_extra_time_map.values()), 4)
            display_extra_price = round(sum(display_extra_price_map.values()), 4)
            repair_extra_time_map = self._quote_collect_non_laser_map(
                dict(row.get("tempos_operacao", {}) or {}),
                dict(snapshot.get("tempos_operacao", {}) or {}),
                digits=4,
            )
            repair_extra_price_map = self._quote_collect_non_laser_map(
                dict(row.get("custos_operacao", {}) or {}),
                dict(snapshot.get("custos_operacao", {}) or {}),
                digits=4,
            )
            repair_extra_time = round(sum(repair_extra_time_map.values()), 4)
            repair_extra_price = round(sum(repair_extra_price_map.values()), 4)
            if laser_base_active:
                max_safe_base_time = round(max(0.0, current_time - repair_extra_time), 4)
                max_safe_base_price = round(max(0.0, current_price - repair_extra_price), 4)
                if laser_base_tempo > max_safe_base_time + 0.0001:
                    laser_base_tempo = max_safe_base_time
                if laser_base_preco > max_safe_base_price + 0.0001:
                    laser_base_preco = max_safe_base_price
            display_time = round(current_time, 2)
            display_price = round(current_price, 4)
            if laser_base_active:
                display_time = round(laser_base_tempo + display_extra_time, 2)
                display_price = round(laser_base_preco + display_extra_price, 4)
            line_qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            material_supplied_by_client = bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False))
            lines.append(
                {
                    "tipo_item": self.desktop_main.normalize_orc_line_type(row.get("tipo_item")),
                    "ref_interna": str(row.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                    "descricao": str(row.get("descricao", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "material_family": str(row.get("material_family", "") or "").strip(),
                    "material_subtype": str(row.get("material_subtype", "") or "").strip(),
                    "material_supplied_by_client": material_supplied_by_client,
                    "material_fornecido_cliente": material_supplied_by_client,
                    "material_cost_included": (False if material_supplied_by_client else bool(row.get("material_cost_included", True))),
                    "espessura": self._fmt(row.get("espessura", "")),
                    "operacao": str(row.get("operacao", "") or "").strip(),
                    "produto_codigo": str(row.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(row.get("produto_unid", "") or "").strip(),
                    "conjunto_codigo": str(row.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(row.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(row.get("grupo_uuid", "") or "").strip(),
                    "qtd_base": round(self._parse_float(row.get("qtd_base", row.get("qtd", 0)), 0), 2),
                    "tempo_peca_min": display_time,
                    "qtd": line_qty,
                    "preco_unit": display_price,
                    "total": round(line_qty * display_price, 2),
                    "desenho": str(row.get("desenho", "") or "").strip(),
                    "laser_base_active": laser_base_active,
                    "laser_base_tempo_unit": laser_base_tempo,
                    "laser_base_preco_unit": laser_base_preco,
                    "operacoes_lista": list(snapshot.get("operacoes", []) or []),
                    "operacoes_fluxo": [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                    "operacoes_detalhe": [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                    "tempos_operacao": dict(snapshot.get("tempos_operacao", {}) or {}),
                    "custos_operacao": dict(snapshot.get("custos_operacao", {}) or {}),
                    "quote_cost_snapshot": dict(snapshot.get("quote_cost_snapshot", {}) or {}),
                }
            )
        return {
            "numero": str(orc.get("numero", "") or "").strip(),
            "data": str(orc.get("data", "") or "").strip()[:10],
            "estado": str(orc.get("estado", "") or "").strip() or "Em edicao",
            "cliente": client,
            "posto_trabalho": self._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "iva_perc": round(self._parse_float(orc.get("iva_perc", 23), 23), 2),
            "desconto_perc": round(self._parse_float(orc.get("desconto_perc", 0), 0), 2),
            "desconto_valor": round(self._parse_float(orc.get("desconto_valor", 0), 0), 2),
            "subtotal_bruto": round(self._parse_float(orc.get("subtotal_bruto", 0), 0), 2),
            "preco_transporte": round(self._parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "subtotal": round(self._parse_float(orc.get("subtotal", 0), 0), 2),
            "total": round(self._parse_float(orc.get("total", 0), 0), 2),
            "numero_encomenda": str(orc.get("numero_encomenda", "") or "").strip(),
            "executado_por": str(orc.get("executado_por", "") or "").strip(),
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "notas_pdf": str(orc.get("notas_pdf", "") or "").strip(),
            "nota_cliente": str(orc.get("nota_cliente", "") or "").strip(),
            "nesting_bridge": dict(orc.get("latest_nesting_bridge", {}) or {}),
            "nesting_group_key": str(orc.get("latest_nesting_group_key", "") or "").strip(),
            "nesting_updated_at": str(orc.get("latest_nesting_updated_at", "") or "").strip(),
            "linhas": lines,
        }

    def orc_clients(self) -> list[dict[str, str]]:
        return list(self.order_clients())

    def _product_lookup(self, codigo: str) -> dict[str, Any] | None:
        code = str(codigo or "").strip()
        if not code:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("produtos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )

    def _next_assembly_model_code(self) -> str:
        highest = 0
        for row in list(self.ensure_data().get("conjuntos_modelo", []) or []):
            codigo = str((row or {}).get("codigo", "") or "").strip().upper()
            digits = "".join(ch for ch in codigo if ch.isdigit())
            if digits:
                try:
                    highest = max(highest, int(digits))
                except Exception:
                    continue
        return f"CJ{highest + 1:04d}"

    def _normalize_assembly_model_item(self, payload: dict[str, Any]) -> dict[str, Any]:
        item_type = self.desktop_main.normalize_orc_line_type(payload.get("tipo_item"))
        quantity = round(self._parse_float(payload.get("qtd", 0), 0), 2)
        if quantity <= 0:
            raise ValueError("Quantidade invalida no conjunto.")
        item = {
            "tipo_item": item_type,
            "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "material": str(payload.get("material", "") or "").strip(),
            "espessura": str(payload.get("espessura", "") or "").strip(),
            "operacao": str(payload.get("operacao", "") or "").strip(),
            "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
            "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
            "qtd": quantity,
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", 0), 0), 4),
            "desenho": str(payload.get("desenho", "") or "").strip(),
        }
        if item_type == self.desktop_main.ORC_LINE_TYPE_PIECE:
            if not item["descricao"]:
                raise ValueError("Descricao obrigatoria na peca do conjunto.")
            if not item["material"] or not item["espessura"]:
                raise ValueError("Material e espessura sao obrigatorios nas pecas fabricadas.")
            item["produto_codigo"] = ""
            item["produto_unid"] = ""
        elif item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            product = self._product_lookup(item["produto_codigo"])
            if product is None:
                raise ValueError("Seleciona um produto de stock valido.")
            item["descricao"] = item["descricao"] or str(product.get("descricao", "") or "").strip()
            item["produto_unid"] = item["produto_unid"] or str(product.get("unid", "") or "UN").strip()
            if item["preco_unit"] <= 0:
                item["preco_unit"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
            if not item["ref_externa"]:
                item["ref_externa"] = item["produto_codigo"]
            item["material"] = ""
            item["espessura"] = ""
            item["desenho"] = ""
            item["operacao"] = item["operacao"] or "Montagem"
        else:
            if not item["descricao"]:
                raise ValueError("Descricao obrigatoria no servico de montagem.")
            item["material"] = ""
            item["espessura"] = ""
            item["produto_codigo"] = ""
            item["produto_unid"] = item["produto_unid"] or "SV"
            item["desenho"] = ""
            item["operacao"] = item["operacao"] or "Montagem"
        return item

    def assembly_model_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for model in list(self.ensure_data().get("conjuntos_modelo", []) or []):
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.desktop_main.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.desktop_main.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.desktop_main.orc_line_is_service(item)),
                "total_base": round(sum(self._parse_float(item.get("qtd", 0), 0) * self._parse_float(item.get("preco_unit", 0), 0) for item in items), 2),
                "notas": str(model.get("notas", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows

    def assembly_model_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        model = next(
            (
                row
                for row in list(self.ensure_data().get("conjuntos_modelo", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [self._normalize_assembly_model_item(dict(item or {})) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "itens": items,
        }

    def assembly_model_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        code = str(payload.get("codigo", "") or "").strip() or self._next_assembly_model_code()
        descricao = str(payload.get("descricao", "") or "").strip()
        if not descricao:
            raise ValueError("Descricao obrigatoria no conjunto.")
        items = [self._normalize_assembly_model_item(dict(row or {})) for row in list(payload.get("itens", []) or [])]
        if not items:
            raise ValueError("O conjunto precisa de pelo menos um item.")
        model = {
            "codigo": code,
            "descricao": descricao,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "created_at": str(payload.get("created_at", "") or "").strip() or self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "itens": [{**item, "linha_ordem": index} for index, item in enumerate(items, start=1)],
        }
        existing = next(
            (
                row
                for row in list(data.get("conjuntos_modelo", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if existing is None:
            data.setdefault("conjuntos_modelo", []).append(model)
        else:
            model["created_at"] = str(existing.get("created_at", "") or "").strip() or model["created_at"]
            existing.update(model)
        self._save(force=True)
        return self.assembly_model_detail(code)

    def assembly_model_remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        rows = list(self.ensure_data().get("conjuntos_modelo", []) or [])
        filtered = [row for row in rows if str(row.get("codigo", "") or "").strip() != code]
        if len(filtered) == len(rows):
            raise ValueError("Conjunto nao encontrado.")
        self.ensure_data()["conjuntos_modelo"] = filtered
        self._save(force=True)

    def assembly_model_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        detail = self.assembly_model_detail(codigo)
        multiplier = round(self._parse_float(quantity, 0), 2)
        if multiplier <= 0:
            raise ValueError("Quantidade do conjunto invalida.")
        group_uuid = self.desktop_main.uuid.uuid4().hex[:12].upper()
        rows: list[dict[str, Any]] = []
        for item in list(detail.get("itens", []) or []):
            line = {
                "tipo_item": self.desktop_main.normalize_orc_line_type(item.get("tipo_item")),
                "ref_interna": "",
                "ref_externa": str(item.get("ref_externa", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(item.get("material", "") or "").strip(),
                "espessura": str(item.get("espessura", "") or "").strip(),
                "operacao": str(item.get("operacao", "") or "").strip(),
                "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                "conjunto_codigo": str(detail.get("codigo", "") or "").strip(),
                "conjunto_nome": str(detail.get("descricao", "") or "").strip(),
                "grupo_uuid": group_uuid,
                "qtd_base": round(self._parse_float(item.get("qtd", 0), 0), 2),
                "tempo_peca_min": round(self._parse_float(item.get("tempo_peca_min", 0), 0), 2),
                "qtd": round(self._parse_float(item.get("qtd", 0), 0) * multiplier, 2),
                "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                "desenho": str(item.get("desenho", "") or "").strip(),
            }
            if self.desktop_main.orc_line_is_product(line) and not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            rows.append(line)
        return rows

    def _normalize_orc_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        line_type = self.desktop_main.normalize_orc_line_type(payload.get("tipo_item"))
        line = {
            "tipo_item": line_type,
            "ref_interna": str(payload.get("ref_interna", "") or "").strip(),
            "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "material": str(payload.get("material", "") or "").strip(),
            "material_family": str(payload.get("material_family", "") or "").strip(),
            "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
            "material_supplied_by_client": bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False)),
            "material_fornecido_cliente": bool(payload.get("material_fornecido_cliente", False) or payload.get("material_supplied_by_client", False)),
            "material_cost_included": (
                bool(payload.get("material_cost_included", True))
                if "material_cost_included" in payload
                else not bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False))
            ),
            "espessura": str(payload.get("espessura", "") or "").strip(),
            "operacao": str(payload.get("operacao", "") or "").strip(),
            "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
            "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
            "conjunto_codigo": str(payload.get("conjunto_codigo", "") or "").strip(),
            "conjunto_nome": str(payload.get("conjunto_nome", "") or "").strip(),
            "grupo_uuid": str(payload.get("grupo_uuid", "") or "").strip(),
            "qtd_base": round(self._parse_float(payload.get("qtd_base", payload.get("qtd", 0)), 0), 2),
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
            "qtd": round(self._parse_float(payload.get("qtd", 0), 0), 2),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", 0), 0), 4),
            "desenho": str(payload.get("desenho", "") or "").strip(),
            "laser_base_active": bool(payload.get("laser_base_active", False)),
            "laser_base_tempo_unit": round(self._parse_float(payload.get("laser_base_tempo_unit", payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0))), 0), 4),
            "laser_base_preco_unit": round(self._parse_float(payload.get("laser_base_preco_unit", payload.get("preco_unit", 0)), 0), 4),
        }
        if line["material_supplied_by_client"] or line["material_fornecido_cliente"]:
            line["material_supplied_by_client"] = True
            line["material_fornecido_cliente"] = True
            line["material_cost_included"] = False
        if line["qtd"] <= 0:
            raise ValueError("Quantidade invalida na linha.")
        if line_type == self.desktop_main.ORC_LINE_TYPE_PIECE:
            if not line["descricao"]:
                raise ValueError("Descricao obrigatoria na linha.")
            if not line["material"] or not line["espessura"]:
                raise ValueError("Material e espessura sao obrigatorios na linha.")
            if not line["material_family"]:
                line["material_family"] = line["material"]
            line["produto_codigo"] = ""
            line["produto_unid"] = ""
        elif line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            product = self._product_lookup(line["produto_codigo"])
            if product is None:
                raise ValueError("Seleciona um produto de stock valido.")
            line["descricao"] = line["descricao"] or str(product.get("descricao", "") or "").strip()
            line["produto_unid"] = line["produto_unid"] or str(product.get("unid", "") or "UN").strip()
            if line["preco_unit"] <= 0:
                line["preco_unit"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
            if not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            line["ref_interna"] = ""
            line["material"] = ""
            line["material_family"] = ""
            line["material_subtype"] = ""
            line["material_supplied_by_client"] = False
            line["material_fornecido_cliente"] = False
            line["material_cost_included"] = False
            line["espessura"] = ""
            line["desenho"] = ""
            line["laser_base_active"] = False
            line["laser_base_tempo_unit"] = 0.0
            line["laser_base_preco_unit"] = 0.0
            line["operacao"] = line["operacao"] or "Montagem"
        else:
            if not line["descricao"]:
                raise ValueError("Descricao obrigatoria na linha de servico.")
            line["ref_interna"] = ""
            line["material"] = ""
            line["material_family"] = ""
            line["material_subtype"] = ""
            line["material_supplied_by_client"] = False
            line["material_fornecido_cliente"] = False
            line["material_cost_included"] = False
            line["espessura"] = ""
            line["produto_codigo"] = ""
            line["produto_unid"] = line["produto_unid"] or "SV"
            line["desenho"] = ""
            line["laser_base_active"] = False
            line["laser_base_tempo_unit"] = 0.0
            line["laser_base_preco_unit"] = 0.0
            line["operacao"] = line["operacao"] or "Montagem"
        line["total"] = round(line["qtd"] * line["preco_unit"], 2)

        def _repair_laser_base(snapshot_payload: dict[str, Any]) -> None:
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            if not bool(line.get("laser_base_active", False)):
                return
            current_time = round(self._parse_float(line.get("tempo_peca_min", 0), 0), 4)
            current_price = round(self._parse_float(line.get("preco_unit", 0), 0), 4)
            repair_extra_time = round(
                sum(
                    self._quote_collect_non_laser_map(
                        dict(payload.get("tempos_operacao", {}) or {}),
                        dict(snapshot_payload.get("tempos_operacao", {}) or {}),
                        digits=4,
                    ).values()
                ),
                4,
            )
            repair_extra_price = round(
                sum(
                    self._quote_collect_non_laser_map(
                        dict(payload.get("custos_operacao", {}) or {}),
                        dict(snapshot_payload.get("custos_operacao", {}) or {}),
                        digits=4,
                    ).values()
                ),
                4,
            )
            max_safe_base_time = round(max(0.0, current_time - repair_extra_time), 4)
            max_safe_base_price = round(max(0.0, current_price - repair_extra_price), 4)
            if round(self._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4) > max_safe_base_time + 0.0001:
                line["laser_base_tempo_unit"] = max_safe_base_time
            if round(self._parse_float(line.get("laser_base_preco_unit", 0), 0), 4) > max_safe_base_price + 0.0001:
                line["laser_base_preco_unit"] = max_safe_base_price

        def _apply_laser_base_blend(snapshot_payload: dict[str, Any]) -> None:
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            if not bool(line.get("laser_base_active", False)):
                return
            base_time = round(self._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4)
            base_price = round(self._parse_float(line.get("laser_base_preco_unit", 0), 0), 4)
            extra_time = 0.0
            extra_price = 0.0
            for op_name, raw_value in dict(snapshot_payload.get("tempos_operacao", {}) or {}).items():
                normalized = str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized and normalized != "Corte Laser":
                    extra_time += self._parse_float(raw_value, 0)
            for op_name, raw_value in dict(snapshot_payload.get("custos_operacao", {}) or {}).items():
                normalized = str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized and normalized != "Corte Laser":
                    extra_price += self._parse_float(raw_value, 0)
            line["tempo_peca_min"] = round(base_time + extra_time, 2)
            line["preco_unit"] = round(base_price + extra_price, 4)
            line["total"] = round(line["qtd"] * line["preco_unit"], 2)

        snapshot_source = {**dict(payload or {}), **line}
        snapshot = self._quote_line_operation_snapshot(snapshot_source)
        _repair_laser_base(snapshot)
        _apply_laser_base_blend(snapshot)
        snapshot_source = {**dict(payload or {}), **line}
        snapshot = self._quote_line_operation_snapshot(snapshot_source)
        line["operacoes_lista"] = list(snapshot.get("operacoes", []) or [])
        line["operacoes_fluxo"] = [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)]
        line["operacoes_detalhe"] = [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        line["tempos_operacao"] = dict(snapshot.get("tempos_operacao", {}) or {})
        line["custos_operacao"] = dict(snapshot.get("custos_operacao", {}) or {})
        line["quote_cost_snapshot"] = dict(snapshot.get("quote_cost_snapshot", {}) or {})
        return line

    def orc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip() or self._peek_next_orc_number()
        existing = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        posto_trabalho = self._normalize_workcenter_value(payload.get("posto_trabalho", "") or (existing or {}).get("posto_trabalho", ""))
        client_payload = dict(payload.get("cliente", {}) or {})
        client = self._normalize_orc_client(client_payload)
        client_code = self._ref_client_code(client.get("codigo", ""))
        if not any(str(client.get(key, "") or "").strip() for key in ("codigo", "nome", "empresa")):
            raise ValueError("Cliente obrigatorio.")
        lines = [self._normalize_orc_line(row) for row in list(payload.get("linhas", []) or [])]
        if client_code:
            self._repair_orc_ref_history(client_code)
            taken_refs, _pairs = self._active_client_ref_usage(client_code, exclude_orc_numero=numero)
            reusable_pairs = self._known_client_ref_pairs(client_code)
            seen_refs: set[str] = set()
            seen_pairs: set[tuple[str, str]] = set()
            reserved_refs = set(taken_refs)
            for line in lines:
                if not self.desktop_main.orc_line_is_piece(line):
                    line["ref_interna"] = ""
                    continue
                ref_externa = str(line.get("ref_externa", "") or "").strip()
                known_ref = self._known_client_ref_for_external(client_code, ref_externa)
                if known_ref:
                    line["ref_interna"] = known_ref
                ref_interna = str(line.get("ref_interna", "") or "").strip().upper()
                pair = (ref_externa, ref_interna)
                can_reuse_known = bool(ref_interna and pair in reusable_pairs)
                can_reuse_current = bool(ref_interna and pair in seen_pairs)
                if not ref_interna or ((ref_interna in seen_refs or ref_interna in reserved_refs) and not can_reuse_known and not can_reuse_current):
                    ref_interna = str(self.desktop_main.next_ref_interna_unique(data, client_code, list(reserved_refs | seen_refs)))
                    line["ref_interna"] = ref_interna
                seen_refs.add(ref_interna)
                seen_pairs.add((ref_externa, ref_interna))
        iva_perc = round(self._parse_float(payload.get("iva_perc", 23), 23), 2)
        desconto_perc = round(max(0.0, min(100.0, self._parse_float(payload.get("desconto_perc", (existing or {}).get("desconto_perc", 0)), 0))), 2)
        preco_transporte = round(self._parse_float(payload.get("preco_transporte", 0), 0), 2)
        custo_transporte = round(self._parse_float(payload.get("custo_transporte", (existing or {}).get("custo_transporte", 0)), 0), 2)
        paletes = round(self._parse_float(payload.get("paletes", (existing or {}).get("paletes", 0)), 0), 2)
        peso_bruto_kg = round(self._parse_float(payload.get("peso_bruto_kg", (existing or {}).get("peso_bruto_kg", 0)), 0), 2)
        volume_m3 = round(self._parse_float(payload.get("volume_m3", (existing or {}).get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self._normalize_supplier_reference(
            payload.get("transportadora_id", (existing or {}).get("transportadora_id", "")),
            payload.get("transportadora_nome", (existing or {}).get("transportadora_nome", "")),
        )
        referencia_transporte = str(payload.get("referencia_transporte", (existing or {}).get("referencia_transporte", "")) or "").strip()
        zona_transporte = str(payload.get("zona_transporte", (existing or {}).get("zona_transporte", "")) or "").strip()
        subtotal_linhas = round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2)
        subtotal_bruto = round(subtotal_linhas + preco_transporte, 2)
        desconto_valor = round(subtotal_bruto * (desconto_perc / 100.0), 2)
        subtotal = round(max(0.0, subtotal_bruto - desconto_valor), 2)
        total = round(subtotal * (1.0 + (iva_perc / 100.0)), 2)
        note = {
            "numero": numero,
            "data": str(payload.get("data", "") or existing.get("data", "") if isinstance(existing, dict) else "") or self.desktop_main.now_iso(),
            "estado": str(payload.get("estado", "") or (existing or {}).get("estado", "") or "Em edição"),
            "cliente": client,
            "posto_trabalho": posto_trabalho,
            "linhas": lines,
            "iva_perc": iva_perc,
            "desconto_perc": desconto_perc,
            "desconto_valor": desconto_valor,
            "preco_transporte": preco_transporte,
            "custo_transporte": custo_transporte,
            "paletes": paletes,
            "peso_bruto_kg": peso_bruto_kg,
            "volume_m3": volume_m3,
            "transportadora_id": transportadora_id,
            "transportadora_nome": transportadora_nome,
            "referencia_transporte": referencia_transporte,
            "zona_transporte": zona_transporte,
            "subtotal_linhas": subtotal_linhas,
            "subtotal_bruto": subtotal_bruto,
            "subtotal": subtotal,
            "total": total,
            "numero_encomenda": str(payload.get("numero_encomenda", "") or (existing or {}).get("numero_encomenda", "") or "").strip(),
            "ano": int(str(payload.get("ano", "") or (existing or {}).get("ano", "") or self.desktop_main.datetime.now().year)),
            "executado_por": str(payload.get("executado_por", "") or (existing or {}).get("executado_por", "") or "").strip(),
            "nota_transporte": str(payload.get("nota_transporte", "") or (existing or {}).get("nota_transporte", "") or "").strip(),
            "notas_pdf": str(payload.get("notas_pdf", "") or (existing or {}).get("notas_pdf", "") or "").strip(),
            "nota_cliente": str(payload.get("nota_cliente", "") or (existing or {}).get("nota_cliente", "") or "").strip(),
        }
        if existing is None:
            data.setdefault("orcamentos", []).append(note)
            if numero == self._peek_next_orc_number():
                try:
                    data["orc_seq"] = max(int(data.get("orc_seq", 1) or 1), int(numero.rsplit("-", 1)[-1]) + 1)
                except Exception:
                    pass
        else:
            existing.update(note)
            note = existing
        self._sync_quote_piece_registry(note)
        self._save(force=True)
        return self.orc_detail(numero)

    def orc_remove(self, numero: str) -> None:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        before = len(list(data.get("orcamentos", []) or []))
        data["orcamentos"] = [row for row in list(data.get("orcamentos", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        if len(data["orcamentos"]) == before:
            raise ValueError("Or?amento n?o encontrado.")
        self._save(force=True)
        try:
            self._mysql_delete_orc_nesting_studies(numero)
        except Exception:
            pass

    def orc_set_state(self, numero: str, estado: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        orc["estado"] = str(estado or "").strip() or "Em edição"
        self._sync_quote_piece_registry(orc)
        self._save(force=True)
        return self.orc_detail(numero)

    def _orc_render_helper(self) -> Any:
        helper = SimpleNamespace(data=self.ensure_data())
        helper._extract_orc_operacoes = lambda orc=None: self.orc_actions._extract_orc_operacoes(helper, orc)
        helper._build_orc_notes_lines = lambda orc: self.orc_actions._build_orc_notes_lines(helper, orc)
        return helper

    def orc_render_pdf(self, numero: str, path: str | Path) -> Path:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        target = Path(path)
        helper = self._orc_render_helper()
        self.orc_actions.render_orc_pdf(helper, str(target), orc)
        return target

    def orc_open_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}.pdf"
        self.orc_render_pdf(numero, target)
        os.startfile(str(target))
        return target

    def orc_render_nesting_study_pdf(self, numero: str, path: str | Path, group_key: str = "") -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        numero_txt = str(numero or "").strip()
        detail = self.orc_detail(numero_txt)
        studies = self.orc_nesting_studies(numero_txt)
        if not studies:
            raise ValueError("Este orçamento ainda não tem estudos de nesting guardados.")
        selected_key = str(group_key or detail.get("nesting_group_key", "") or "").strip()
        if not selected_key or selected_key not in studies:
            ordered_keys = sorted(studies.keys())
            selected_key = ordered_keys[0]
        study = dict(studies.get(selected_key, {}) or {})
        if not study:
            raise ValueError("Estudo de nesting não encontrado.")

        result_data = dict(study.get("result_data", {}) or {})
        summary = dict(result_data.get("summary", study.get("summary", {})) or {})
        bridge = dict(study.get("quote_bridge", {}) or {})
        cost_report = dict(study.get("cost_report", {}) or {})
        options = dict(study.get("options", {}) or {})
        sheets = [dict(row or {}) for row in list(result_data.get("sheets", []) or [])]
        sheet_candidates = [dict(row or {}) for row in list(result_data.get("sheet_candidates", []) or [])]
        unplaced = [dict(row or {}) for row in list(result_data.get("unplaced", []) or [])]
        warnings = [str(row or "").strip() for row in list(result_data.get("warnings", []) or []) if str(row or "").strip()]
        part_rows = [dict(row or {}) for row in list(cost_report.get("part_rows", []) or bridge.get("part_rows", [])) if isinstance(row, dict)]
        decision_lines = [str(row or "").strip() for row in list(cost_report.get("decision_lines", []) or []) if str(row or "").strip()]
        totals = dict(cost_report.get("totals", {}) or {})

        target = Path(path)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        page_width, page_height = landscape(A4)
        margin = 26
        c = pdf_canvas.Canvas(str(target), pagesize=landscape(A4))

        font_regular = "Helvetica"
        font_bold = "Helvetica-Bold"
        for name, file_name in (("SegoeUI", "segoeui.ttf"), ("SegoeUI-Bold", "segoeuib.ttf")):
            font_path = Path(r"C:\Windows\Fonts") / file_name
            if font_path.exists():
                try:
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont

                    pdfmetrics.registerFont(TTFont(name, str(font_path)))
                    if "Bold" in name:
                        font_bold = name
                    else:
                        font_regular = name
                except Exception:
                    pass

        def set_font(bold: bool, size: float) -> None:
            c.setFont(font_bold if bold else font_regular, size)

        def draw_header(title: str, subtitle: str) -> float:
            c.setFillColor(palette["primary"])
            c.roundRect(margin, page_height - margin - 60, page_width - (margin * 2), 60, 18, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 18)
            c.drawString(margin + 16, page_height - margin - 24, title)
            set_font(False, 9)
            c.drawString(margin + 16, page_height - margin - 40, subtitle)
            company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
            generated = datetime.now().strftime("%d/%m/%Y %H:%M")
            c.drawRightString(page_width - margin - 16, page_height - margin - 24, company)
            c.drawRightString(page_width - margin - 16, page_height - margin - 40, generated)
            return page_height - margin - 74

        def draw_footer() -> None:
            c.setFillColor(palette["muted"])
            set_font(False, 8)
            c.drawString(margin, 20, "Estudo de nesting guardado por orçamento, ligado ao Plano de Chapa e ao custo do lote.")
            c.drawRightString(page_width - margin, 20, f"Orçamento {numero_txt}")

        def draw_metric_card(x: float, y_top: float, width: float, title: str, value: str, accent: Any) -> None:
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - 52, width, 46, 12, stroke=1, fill=1)
            c.setFillColor(accent)
            c.rect(x + 10, y_top - 20, 26, 4, stroke=0, fill=1)
            c.setFillColor(palette["muted"])
            set_font(True, 8.2)
            c.drawString(x + 10, y_top - 14, title)
            c.setFillColor(palette["ink"])
            set_font(True, 13)
            c.drawString(x + 10, y_top - 35, value)

        def draw_info_box(x: float, y_top: float, width: float, title: str, lines: list[str], *, tone: str = "default") -> float:
            body_lines = [line for line in lines if str(line or "").strip()]
            wrapped: list[str] = []
            for line in body_lines:
                wrapped.extend(_pdf_wrap_text(line, font_regular, 8.4, width - 22, max_lines=3) or ["-"])
            box_height = max(54.0, 18.0 + (len(wrapped) * 11.0) + 18.0)
            tone_fill = palette["primary_soft_2"] if tone == "info" else colors.HexColor("#FFF8EB") if tone == "warning" else colors.white
            c.setFillColor(tone_fill)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - box_height, width, box_height, 12, stroke=1, fill=1)
            c.setFillColor(palette["ink"])
            set_font(True, 10)
            c.drawString(x + 10, y_top - 16, title)
            set_font(False, 8.4)
            cursor_y = y_top - 30
            for line in wrapped:
                c.drawString(x + 10, cursor_y, line)
                cursor_y -= 11
            return box_height

        def ensure_page(current_y: float, needed: float, title: str, subtitle: str) -> float:
            if current_y - needed >= 46:
                return current_y
            draw_footer()
            c.showPage()
            return draw_header(title, subtitle)

        def draw_table_header(y_top: float, columns: list[tuple[str, float]]) -> tuple[float, list[float], float]:
            total_width = page_width - (margin * 2)
            c.setFillColor(palette["primary_dark"])
            c.roundRect(margin, y_top - 20, total_width, 18, 8, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 8)
            x_positions: list[float] = []
            cursor_x = margin + 7
            for label, ratio in columns:
                x_positions.append(cursor_x)
                c.drawString(cursor_x, y_top - 13, label)
                cursor_x += total_width * ratio
            return y_top - 24, x_positions, total_width

        def draw_sheet_map(x: float, y_top: float, width: float, height: float, sheet: dict[str, Any]) -> None:
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - height, width, height, 14, stroke=1, fill=1)
            title = (
                f"Chapa {int(sheet.get('index', 0) or 0)} | "
                f"{str(sheet.get('source_label', summary.get('selected_sheet_profile', {}).get('name', '-')) or '-').strip()}"
            )
            c.setFillColor(palette["ink"])
            set_font(True, 9.2)
            c.drawString(x + 10, y_top - 16, _pdf_clip_text(title, width - 20, font_bold, 9.2))
            set_font(False, 7.8)
            c.setFillColor(palette["muted"])
            c.drawString(
                x + 10,
                y_top - 28,
                f"{int(sheet.get('part_count', 0) or 0)} peça(s) | real {self._fmt(sheet.get('utilization_net_pct', 0))}%",
            )

            draw_x = x + 12
            draw_y = y_top - height + 12
            draw_w = width - 24
            draw_h = height - 48
            sheet_w = max(1.0, float(sheet.get("sheet_width_mm", 0) or 1.0))
            sheet_h = max(1.0, float(sheet.get("sheet_height_mm", 0) or 1.0))
            scale = min(draw_w / sheet_w, draw_h / sheet_h)
            body_w = sheet_w * scale
            body_h = sheet_h * scale
            offset_x = draw_x + ((draw_w - body_w) / 2.0)
            offset_y = draw_y + ((draw_h - body_h) / 2.0)

            def map_point(px: float, py: float) -> tuple[float, float]:
                return (offset_x + (px * scale), offset_y + (py * scale))

            outer_polygons = [list(points or []) for points in list(sheet.get("sheet_outer_polygons", []) or [])]
            hole_polygons = [list(points or []) for points in list(sheet.get("sheet_hole_polygons", []) or [])]
            c.setStrokeColor(palette["line_strong"])
            c.setFillColor(colors.HexColor("#F8FAFC"))
            if outer_polygons:
                for polygon_points in outer_polygons:
                    mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                    if len(mapped) >= 3:
                        c.lines(
                            [
                                (mapped[index][0], mapped[index][1], mapped[(index + 1) % len(mapped)][0], mapped[(index + 1) % len(mapped)][1])
                                for index in range(len(mapped))
                            ]
                        )
                for polygon_points in hole_polygons:
                    mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                    if len(mapped) >= 3:
                        c.setFillColor(colors.white)
                        c.lines(
                            [
                                (mapped[index][0], mapped[index][1], mapped[(index + 1) % len(mapped)][0], mapped[(index + 1) % len(mapped)][1])
                                for index in range(len(mapped))
                            ]
                        )
            else:
                c.rect(offset_x, offset_y, body_w, body_h, stroke=1, fill=1)

            palette_hexes = ["#dbeafe", "#dcfce7", "#fef3c7", "#ffe4e6", "#ede9fe", "#cffafe", "#e2e8f0"]
            for idx, placement in enumerate(list(sheet.get("placements", []) or [])):
                c.setStrokeColor(colors.HexColor("#274c77"))
                c.setFillColor(colors.HexColor(palette_hexes[idx % len(palette_hexes)]))
                poly_groups = [list(points or []) for points in list(placement.get("shape_outer_polygons", []) or [])]
                if poly_groups:
                    for polygon_points in poly_groups:
                        mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                        if len(mapped) >= 3:
                            c.lines(
                                [
                                    (mapped[index][0], mapped[index][1], mapped[(index + 1) % len(mapped)][0], mapped[(index + 1) % len(mapped)][1])
                                    for index in range(len(mapped))
                                ]
                            )
                else:
                    c.rect(
                        offset_x + (float(placement.get("x_mm", 0) or 0) * scale),
                        offset_y + (float(placement.get("y_mm", 0) or 0) * scale),
                        max(1.0, float(placement.get("width_mm", 0) or 0) * scale),
                        max(1.0, float(placement.get("height_mm", 0) or 0) * scale),
                        stroke=1,
                        fill=1,
                    )

        group_label = str(study.get("group_label", "") or selected_key).strip() or selected_key
        profile_name = str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or bridge.get("selected_profile_name", "") or "Apenas stock").strip() or "Apenas stock"
        subtitle = f"Orçamento {numero_txt} | Grupo {group_label} | Perfil {profile_name}"
        y = draw_header("Estudo de Nesting + Custo", subtitle)

        cards = [
            ("Programadas", f"{int(bridge.get('part_count_placed', summary.get('part_count_placed', 0)) or 0)}/{int(bridge.get('part_count_requested', summary.get('part_count_requested', 0)) or 0)}", palette["primary"]),
            ("Chapas", str(int(bridge.get("sheet_count", summary.get("sheet_count", 0)) or 0)), palette["success"]),
            ("Util. real", f"{self._fmt(summary.get('utilization_net_pct', 0))}%", palette["warning"]),
            ("Matéria", self._fmt_eur(summary.get("material_net_cost_eur", 0)), palette["primary_dark"]),
            ("Compra", self._fmt_eur(summary.get("material_purchase_requirement_eur", 0)), palette["danger"]),
        ]
        card_gap = 8
        card_width = (page_width - (margin * 2) - (card_gap * (len(cards) - 1))) / len(cards)
        card_x = margin
        for title, value, accent in cards:
            draw_metric_card(card_x, y, card_width, title, value, accent)
            card_x += card_width + card_gap
        y -= 64

        study_lines = [
            f"Cliente: {str(dict(detail.get('cliente', {}) or {}).get('nome', '') or '-').strip() or '-'}",
            f"Método: {str(bridge.get('analysis_method', summary.get('selection_mode', '-')) or '-').strip() or '-'}",
            f"Regras: margem peça {float(options.get('part_spacing_mm', 0) or 0):.1f} mm | margem borda {float(options.get('edge_margin_mm', 0) or 0):.1f} mm | rotação auto {'sim' if bool(options.get('allow_rotate')) else 'não'}",
            f"Fluxo: stock primeiro {'sim' if bool(options.get('use_stock_first')) else 'não'} | compra complementar {'sim' if bool(options.get('allow_purchase_fallback', True)) else 'não'} | contorno {'sim' if bool(options.get('shape_aware')) else 'não'}",
        ]
        report_lines = [
            f"Valor comercial colocado: {self._fmt_eur(totals.get('quoted_total_eur', bridge.get('quoted_total_eur', 0)))}",
            f"Tempo máquina: {self._fmt(totals.get('machine_total_min', 0))} min | corte {self._fmt(totals.get('cut_length_m', 0))} m | pierces {int(totals.get('pierce_count', 0) or 0)}",
            f"Stock usado: {int(summary.get('stock_sheet_count', 0) or 0)} | retalhos {int(summary.get('remnant_sheet_count', 0) or 0)} | compra {int(summary.get('purchased_sheet_count', 0) or 0)}",
        ]
        left_box_h = draw_info_box(margin, y, (page_width - (margin * 2) - 10) * 0.52, "Contexto do estudo", study_lines, tone="info")
        right_box_h = draw_info_box(margin + ((page_width - (margin * 2) - 10) * 0.52) + 10, y, (page_width - (margin * 2) - 10) * 0.48, "Resumo económico", report_lines, tone="default")
        y -= max(left_box_h, right_box_h) + 10

        note_lines = (decision_lines[:4] or warnings[:4] or ["Sem observações adicionais registadas."])
        note_height = draw_info_box(margin, y, page_width - (margin * 2), "Decisão e observações", note_lines, tone="warning" if warnings else "default")
        y -= note_height + 12

        candidate_columns = [
            ("Cenário", 0.29),
            ("Método", 0.29),
            ("Chapas", 0.09),
            ("Compact.", 0.11),
            ("Compra m2", 0.11),
            ("Total m2", 0.11),
        ]
        if sheet_candidates:
            y = ensure_page(y, 90, "Estudo de Nesting + Custo", subtitle)
            y, x_positions, total_w = draw_table_header(y, candidate_columns)
            row_h = 18
            for row_index, candidate in enumerate(sheet_candidates[:8]):
                y = ensure_page(y, row_h + 8, "Estudo de Nesting + Custo", subtitle)
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(candidate.get("name", "") or "-").strip() or "-",
                    str(candidate.get("method", "") or bridge.get("analysis_method", "-")).strip() or "-",
                    str(int(candidate.get("sheet_count", 0) or 0)),
                    f"{self._fmt(candidate.get('layout_compactness_pct', 0))}%",
                    f"{self._fmt((candidate.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0)}",
                    f"{self._fmt((candidate.get('total_sheet_area_mm2', 0) or 0) / 1_000_000.0)}",
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    max_w = (total_w * candidate_columns[idx][1]) - 10
                    draw_value = _pdf_clip_text(value, max_w, font_regular, 8)
                    c.drawString(x_positions[idx], y - 10, draw_value)
                y -= row_h
            y -= 8

        part_columns = [
            ("Ref.", 0.16),
            ("Descrição", 0.34),
            ("Qtd", 0.07),
            ("Tempo", 0.10),
            ("Corte", 0.10),
            ("Pierces", 0.09),
            ("Valor", 0.14),
        ]
        if part_rows:
            y = ensure_page(y, 110, "Estudo de Nesting + Custo", subtitle)
            y, x_positions, total_w = draw_table_header(y, part_columns)
            row_h = 18
            for row_index, row in enumerate(part_rows):
                y = ensure_page(y, row_h + 8, "Estudo de Nesting + Custo", subtitle)
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(row.get("ref_externa", "") or "-").strip() or "-",
                    str(row.get("description", "") or "-").strip() or "-",
                    str(int(row.get("qty", 0) or 0)),
                    f"{self._fmt(row.get('machine_total_min', 0))} min",
                    f"{self._fmt(row.get('cut_length_m', 0))} m",
                    str(int(row.get("pierce_count", 0) or 0)),
                    self._fmt_eur(row.get("quoted_total_eur", 0)),
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    max_w = (total_w * part_columns[idx][1]) - 10
                    draw_value = _pdf_clip_text(value, max_w, font_regular, 8)
                    c.drawString(x_positions[idx], y - 10, draw_value)
                y -= row_h
            y -= 10

        if unplaced:
            y = ensure_page(y, 80, "Estudo de Nesting + Custo", subtitle)
            unplaced_preview = [
                f"{str(row.get('ref_externa', '-') or '-').strip() or '-'} | {str(row.get('description', '-') or '-').strip() or '-'}"
                for row in unplaced[:6]
            ]
            box_h = draw_info_box(margin, y, page_width - (margin * 2), "Peças fora do plano", unplaced_preview, tone="warning")
            y -= box_h + 10

        if sheets:
            draw_footer()
            c.showPage()
            y = draw_header("Mapas de Chapa", subtitle)
            box_gap = 12
            box_width = (page_width - (margin * 2) - box_gap) / 2.0
            box_height = 180.0
            x_positions = [margin, margin + box_width + box_gap]
            current_col = 0
            current_y = y
            for sheet in sheets:
                if current_y - box_height < 50:
                    draw_footer()
                    c.showPage()
                    current_y = draw_header("Mapas de Chapa", subtitle)
                    current_col = 0
                draw_sheet_map(x_positions[current_col], current_y, box_width, box_height, sheet)
                if current_col == 1:
                    current_col = 0
                    current_y -= box_height + 14
                else:
                    current_col = 1

        draw_footer()
        c.save()
        return target

    def orc_open_nesting_study_pdf(self, numero: str, group_key: str = "") -> Path:
        safe_group = "".join(ch if ch.isalnum() else "_" for ch in str(group_key or "").strip()) or "grupo"
        target = Path(tempfile.gettempdir()) / f"lugest_nesting_{str(numero or '').strip()}_{safe_group}.pdf"
        self.orc_render_nesting_study_pdf(numero, target, group_key=group_key)
        os.startfile(str(target))
        return target

    def orc_convert_to_order(self, numero: str, nota_cliente: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        note = str(nota_cliente or "").strip()
        orc = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        if str(orc.get("numero_encomenda", "") or "").strip():
            raise ValueError("Orcamento ja convertido.")
        estado_norm = str(orc.get("estado", "") or "").strip().lower()
        if "aprovado" not in estado_norm:
            raise ValueError("Apenas orcamentos aprovados podem ser convertidos.")
        if not list(orc.get("linhas", []) or []):
            raise ValueError("Sem linhas para converter.")
        cli = self._normalize_orc_client(orc.get("cliente", {}))
        codigo = str(cli.get("codigo", "") or "").strip()
        if codigo and self.desktop_main.find_cliente(data, codigo):
            cliente_code = codigo
        else:
            cliente_code = ""
            for row in list(data.get("clientes", []) or []):
                if not isinstance(row, dict):
                    continue
                if cli.get("nif") and str(row.get("nif", "") or "").strip() == str(cli.get("nif", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
                if cli.get("nome") and str(row.get("nome", "") or "").strip() == str(cli.get("nome", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
            if not cliente_code:
                cliente_code = str(self.desktop_main.next_cliente_codigo(data))
                data.setdefault("clientes", []).append(
                    {
                        "codigo": cliente_code,
                        "nome": str(cli.get("nome", "") or "").strip(),
                        "nif": str(cli.get("nif", "") or "").strip(),
                        "morada": str(cli.get("morada", "") or "").strip(),
                        "contacto": str(cli.get("contacto", "") or "").strip(),
                        "email": str(cli.get("email", "") or "").strip(),
                        "observacoes": "",
                    }
                )
        alert_txt = (
            f"ALERTA: Encomenda gerada por conversao do orcamento {orc.get('numero')}. "
            "Confirmar dados de cliente, materiais, espessuras e prazos."
        )
        obs_txt = f"{alert_txt} | Origem: Orcamento {orc.get('numero')}"
        if note:
            obs_txt += f" | Nota cliente: {note}"
        enc = {
            "numero": self.desktop_main.next_encomenda_numero(data),
            "cliente": cliente_code,
            "nota_cliente": note,
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "preco_transporte": round(self._parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "local_descarga": str(cli.get("morada", "") or "").strip(),
            "transporte_numero": "",
            "estado_transporte": "",
            "data_criacao": self.desktop_main.now_iso(),
            "data_entrega": "",
            "tempo": 0.0,
            "tempo_estimado": 0.0,
            "cativar": False,
            "posto_trabalho": self._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "observacoes": obs_txt,
            "alerta_conversao": True,
            "estado": "Preparacao",
            "materiais": [],
            "reservas": [],
            "montagem_itens": [],
            "numero_orcamento": orc.get("numero"),
        }
        mats: dict[str, dict[str, Any]] = {}
        piece_idx = 1
        total_time = 0.0
        used_refs: set[str] = set()
        montagem_items: list[dict[str, Any]] = []
        for line in list(orc.get("linhas", []) or []):
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            qtd_line = float(line.get("qtd", 0) or 0)
            tempo_peca = float(line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)) or 0)
            total_time += tempo_peca * max(qtd_line, 0.0)
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE:
                montagem_items.append(
                    {
                        "linha_ordem": len(montagem_items) + 1,
                        "tipo_item": line_type,
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "produto_codigo": str(line.get("produto_codigo", "") or "").strip(),
                        "produto_unid": str(line.get("produto_unid", "") or "").strip(),
                        "qtd_planeada": round(qtd_line, 2),
                        "qtd_consumida": 0.0,
                        "preco_unit": round(self._parse_float(line.get("preco_unit", 0), 0), 4),
                        "conjunto_codigo": str(line.get("conjunto_codigo", "") or "").strip(),
                        "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                        "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                        "estado": "Pendente",
                        "obs": str(line.get("operacao", "") or "").strip(),
                        "created_at": self.desktop_main.now_iso(),
                        "consumed_at": "",
                        "consumed_by": "",
                    }
                )
                continue
            material = str(line.get("material", "") or "").strip()
            espessura = str(line.get("espessura", "") or "").strip()
            if not material or not espessura:
                raise ValueError("Todas as linhas precisam de material e espessura.")
            mats.setdefault(material, {"material": material, "estado": "Preparacao", "espessuras": {}})
            mats[material]["espessuras"].setdefault(
                espessura,
                {"espessura": espessura, "tempo_min": 0.0, "tempos_operacao": {}, "estado": "Preparacao", "pecas": []},
            )
            planning_ops = [op for op in self._planning_ops_from_ops_value(line.get("operacao", "")) if op != "Montagem"]
            esp_bucket = mats[material]["espessuras"][espessura]
            tempos_operacao = esp_bucket.setdefault("tempos_operacao", {})
            detailed_op_times = {
                str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip(): self._parse_float(raw_value, 0)
                for op_name, raw_value in dict(line.get("tempos_operacao", {}) or {}).items()
                if str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip() and self._parse_float(raw_value, 0) > 0
            }
            if detailed_op_times:
                for op_name, unit_time in detailed_op_times.items():
                    if op_name not in planning_ops:
                        continue
                    total_time = unit_time * max(qtd_line, 0.0)
                    tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + total_time, 2)
                    if op_name == "Corte Laser":
                        esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + total_time, 2)
            elif len(planning_ops) == 1:
                op_name = planning_ops[0]
                tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if op_name == "Corte Laser":
                    esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            elif "Corte Laser" in planning_ops:
                tempos_operacao["Corte Laser"] = round(float(tempos_operacao.get("Corte Laser", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            raw_ref_interna = str(line.get("ref_interna", "") or "").strip()
            if raw_ref_interna and raw_ref_interna not in used_refs:
                ref_interna = raw_ref_interna
            else:
                ref_interna = str(self.desktop_main.next_ref_interna_unique(data, cliente_code, list(used_refs)))
            used_refs.add(ref_interna)
            ops_txt = self.quote_format_operacoes(line.get("operacao", ""))
            peca = {
                "id": f"PEC{piece_idx:05d}",
                "ref_interna": ref_interna,
                "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                "material": material,
                "espessura": espessura,
                "quantidade_pedida": qtd_line,
                "Operacoes": ops_txt,
                "Observacoes": str(line.get("descricao", "") or "").strip(),
                "desenho": str(line.get("desenho", "") or "").strip(),
                "tempo_peca_min": tempo_peca,
                "tempos_operacao": dict(line.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(line.get("custos_operacao", {}) or {}),
                "operacoes_detalhe": [dict(item or {}) for item in list(line.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "of": self.desktop_main.next_of_numero(data),
                "opp": self.desktop_main.next_opp_numero(data),
                "estado": "Preparacao",
                "produzido_ok": 0.0,
                "produzido_nok": 0.0,
                "inicio_producao": "",
                "fim_producao": "",
            }
            peca["operacoes_fluxo"] = self.desktop_main.build_operacoes_fluxo(ops_txt)
            piece_idx += 1
            mats[material]["espessuras"][espessura]["pecas"].append(peca)
            self.desktop_main.update_refs(data, peca["ref_interna"], peca["ref_externa"])
        enc["materiais"] = []
        for row in mats.values():
            row["espessuras"] = list(row["espessuras"].values())
            enc["materiais"].append(row)
        enc["montagem_itens"] = montagem_items
        enc["tempo_estimado"] = round(total_time, 2)
        enc["tempo"] = round(total_time / 60.0, 2) if total_time > 0 else 0.0
        data.setdefault("encomendas", []).append(enc)
        self._ensure_unique_order_piece_refs(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        orc["numero_encomenda"] = enc["numero"]
        if note:
            orc["nota_cliente"] = note
        orc["estado"] = "Convertido em Encomenda"
        self._save(force=True)
        return {
            "orcamento": self.orc_detail(numero),
            "encomenda": self.order_detail(enc["numero"]),
        }

    def orc_suggest_notes(self, payload: dict[str, Any]) -> str:
        helper = self._orc_render_helper()
        lines = self.orc_actions._build_orc_notes_lines(helper, payload)
        return "\n".join([str(line or "").strip() for line in lines if str(line or "").strip()])

    def _planning_week_start(self, week_start: str | date | None = None) -> date:
        if isinstance(week_start, date):
            return week_start - timedelta(days=week_start.weekday())
        raw = str(week_start or "").strip()
        if raw:
            try:
                parsed = datetime.fromisoformat(raw).date()
                return parsed - timedelta(days=parsed.weekday())
            except Exception:
                pass
        today = datetime.now().date()
        return today - timedelta(days=today.weekday())

    def _planning_grid_metrics(self) -> tuple[int, int, int]:
        return 480, 1080, 30

    def _planning_default_blocked_windows(self) -> list[dict[str, Any]]:
        return [
            {
                "id": "LUNCH",
                "label": "Almoco",
                "start_min": 12 * 60 + 30,
                "end_min": 14 * 60,
                "weekdays": [0, 1, 2, 3, 4, 5],
            }
        ]

    def _planning_blocked_windows(self) -> list[dict[str, Any]]:
        data = self.ensure_data()
        rows = list(data.get("plano_bloqueios", []) or [])
        if not rows:
            rows = self._planning_default_blocked_windows()
        normalized: list[dict[str, Any]] = []
        for index, row in enumerate(rows):
            if not isinstance(row, dict):
                continue
            try:
                start_min = int(float(row.get("start_min", 0) or 0))
                end_min = int(float(row.get("end_min", 0) or 0))
            except Exception:
                continue
            if end_min <= start_min:
                continue
            weekdays = [int(v) for v in list(row.get("weekdays", [0, 1, 2, 3, 4, 5])) if str(v).strip().isdigit()]
            if not weekdays:
                weekdays = [0, 1, 2, 3, 4, 5]
            normalized.append(
                {
                    "id": str(row.get("id", "") or f"PB{index+1:03d}").strip(),
                    "label": str(row.get("label", "") or "Bloqueio").strip(),
                    "start_min": start_min,
                    "end_min": end_min,
                    "weekdays": sorted(set(weekdays)),
                }
            )
        return normalized

    def _planning_block_matches_day(self, block: dict[str, Any], day_txt: str = "") -> bool:
        if not day_txt:
            return True
        try:
            weekday = datetime.fromisoformat(str(day_txt or "").strip()).date().weekday()
        except Exception:
            return True
        weekdays = list(block.get("weekdays", []) or [])
        return not weekdays or weekday in weekdays

    def _planning_interval_blocked(self, start_min: int, end_min: int, day_txt: str = "") -> bool:
        for block in self._planning_blocked_windows():
            if not self._planning_block_matches_day(block, day_txt):
                continue
            block_start = int(block.get("start_min", 0) or 0)
            block_end = int(block.get("end_min", 0) or 0)
            if not (end_min <= block_start or start_min >= block_end):
                return True
        return False

    def _planning_is_duplicate(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any = "Corte Laser",
        ignore_id: str = "",
    ) -> bool:
        return (
            self._planning_planned_minutes(numero, material, espessura, operation=operation, ignore_id=ignore_id) > 0
            and self._planning_remaining_minutes(numero, material, espessura, operation=operation, ignore_id=ignore_id) <= 0
        )

    def _planning_is_free(
        self,
        day_txt: str,
        start_min: int,
        end_min: int,
        operation: Any = "Corte Laser",
        ignore_id: str = "",
    ) -> bool:
        if self._planning_interval_blocked(start_min, end_min, day_txt):
            return False
        ignore = str(ignore_id or "").strip()
        op_txt = self._planning_normalize_operation(operation)
        for row in list(self.ensure_data().get("plano", []) or []):
            if str(row.get("data", "") or "").strip() != day_txt:
                continue
            if ignore and str(row.get("id", "") or "").strip() == ignore:
                continue
            if self._planning_row_operation(row) != op_txt:
                continue
            try:
                other_start = self.desktop_main.time_to_minutes(str(row.get("inicio", "") or "").strip())
                other_end = other_start + int(float(row.get("duracao_min", 0) or 0))
            except Exception:
                continue
            if not (end_min <= other_start or start_min >= other_end):
                return False
        return True

    def _planning_item_key(self, numero: str, material: str, espessura: str) -> tuple[str, str, str]:
        return (
            str(numero or "").strip(),
            str(material or "").strip(),
            str(espessura or "").strip(),
        )

    def _planning_item_op_key(self, numero: str, material: str, espessura: str, operation: Any = "Corte Laser") -> tuple[str, str, str, str]:
        return self._planning_item_key(numero, material, espessura) + (self._planning_normalize_operation(operation),)

    def _planning_montagem_material(self) -> str:
        return "Montagem"

    def _planning_montagem_espessura(self) -> str:
        return "Final"

    def _planning_is_montagem_item(self, material: str, espessura: str) -> bool:
        return (
            self.desktop_main.norm_text(material) == self.desktop_main.norm_text(self._planning_montagem_material())
            and self.desktop_main.norm_text(espessura) == self.desktop_main.norm_text(self._planning_montagem_espessura())
        )

    def _planning_round_duration(self, duration: Any) -> int:
        try:
            minutes = int(round(float(duration or 0)))
        except Exception:
            return 0
        if minutes <= 0:
            return 0
        _start_min, _end_min, slot = self._planning_grid_metrics()
        if minutes % slot != 0:
            minutes = int((minutes + slot - 1) // slot) * slot
        return max(slot, minutes)

    def _planning_find_esp_obj(self, enc: dict[str, Any] | None, material: str, espessura: str) -> dict[str, Any] | None:
        if not isinstance(enc, dict):
            return None
        mat_norm = self.desktop_main.norm_text(material or "")
        esp_norm = self._norm_esp_token(espessura)
        for mat in list(enc.get("materiais", []) or []):
            if self.desktop_main.norm_text(mat.get("material", "")) != mat_norm:
                continue
            for esp_obj in list(mat.get("espessuras", []) or []):
                if self._norm_esp_token(esp_obj.get("espessura", "")) == esp_norm:
                    return esp_obj
        return None

    def _planning_item_total_minutes(self, numero: str, material: str, espessura: str, operation: Any = "Corte Laser") -> int:
        op_txt = self._planning_normalize_operation(operation)
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        if op_txt == "Montagem" or self._planning_is_montagem_item(material, espessura):
            if not isinstance(enc, dict):
                return 0
            total = self._parse_float(self.desktop_main.encomenda_montagem_tempo_min(enc), 0)
            if total <= 0 and list(self.desktop_main.encomenda_montagem_itens(enc) or []):
                return self._planning_round_duration(1)
            return self._planning_round_duration(total)
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        time_map = self._planning_operation_times_map(esp_obj)
        return self._planning_round_duration(time_map.get(op_txt, 0))

    def _planning_planned_minutes(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any = "Corte Laser",
        ignore_id: str = "",
    ) -> int:
        target = self._planning_item_op_key(numero, material, espessura, operation)
        ignore = str(ignore_id or "").strip()
        total = 0
        for row in list(self.ensure_data().get("plano", []) or []):
            if ignore and str(row.get("id", "") or "").strip() == ignore:
                continue
            row_key = self._planning_item_op_key(
                row.get("encomenda", ""),
                row.get("material", ""),
                row.get("espessura", ""),
                self._planning_row_operation(row),
            )
            if row_key != target:
                continue
            total += self._planning_round_duration(row.get("duracao_min", 0))
        return total

    def _planning_remaining_minutes(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any = "Corte Laser",
        ignore_id: str = "",
    ) -> int:
        total = self._planning_item_total_minutes(numero, material, espessura, operation=operation)
        planned = self._planning_planned_minutes(numero, material, espessura, operation=operation, ignore_id=ignore_id)
        return max(0, total - planned)

    def _planning_item_has_laser(self, numero: str, material: str, espessura: str) -> bool:
        if self._planning_is_montagem_item(material, espessura):
            return False
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        if not isinstance(esp_obj, dict):
            return False
        if bool(esp_obj.get("laser_concluido")):
            return True
        for piece in list(esp_obj.get("pecas", []) or []):
            for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
                op_name = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
                if self._is_laser_operation(op_name):
                    return True
        return False

    def _planning_item_color(self, numero: str, material: str, espessura: str) -> str:
        target = self._planning_item_key(numero, material, espessura)
        palette = list(getattr(self.desktop_main, "PLANO_CORES", ["#fbecee"])) or ["#fbecee"]
        for row in list(self.ensure_data().get("plano", []) or []):
            row_key = self._planning_item_key(row.get("encomenda", ""), row.get("material", ""), row.get("espessura", ""))
            if row_key != target:
                continue
            color = str(row.get("color", "") or "").strip()
            if color:
                return color
        return str(palette[len(self.ensure_data().get("plano", [])) % len(palette)] or "#fbecee")

    def _planning_montagem_obs(self, enc: dict[str, Any]) -> str:
        resumo = str(self.desktop_main.encomenda_montagem_resumo(enc) or "Montagem final").strip()
        shortages = self._order_montagem_shortages(enc)
        if shortages:
            sample = ", ".join(str(row.get("produto_codigo", "") or "-").strip() for row in shortages[:2])
            suffix = f"Falta stock ({len(shortages)})"
            if sample:
                suffix += f": {sample}"
            return f"{resumo} | {suffix}"
        return f"{resumo} | Stock OK"

    def _planning_next_free_segment(
        self,
        dates: list[date],
        cur_day_idx: int,
        cur_min: int,
    ) -> tuple[int, str | None, int | None, int | None]:
        start_min, end_min, slot = self._planning_grid_metrics()
        day_idx = max(0, int(cur_day_idx or 0))
        cursor = max(start_min, int(cur_min or start_min))
        if cursor % slot != 0:
            cursor = int((cursor + slot - 1) // slot) * slot
        while day_idx < len(dates):
            day_txt = dates[day_idx].isoformat()
            local_cursor = max(start_min, cursor)
            if local_cursor % slot != 0:
                local_cursor = int((local_cursor + slot - 1) // slot) * slot
            while local_cursor + slot <= end_min:
                if not self._planning_is_free(day_txt, local_cursor, local_cursor + slot):
                    local_cursor += slot
                    continue
                segment_start = local_cursor
                segment_end = local_cursor + slot
                while segment_end + slot <= end_min and self._planning_is_free(day_txt, segment_end, segment_end + slot):
                    segment_end += slot
                return day_idx, day_txt, segment_start, segment_end
            day_idx += 1
            cursor = start_min
        return day_idx, None, None, None

    def _planning_now_floor_for_day(self, day_txt: str) -> int | None:
        start_min, end_min, slot = self._planning_grid_metrics()
        try:
            target_day = datetime.fromisoformat(str(day_txt or "").strip()).date()
        except Exception:
            return start_min
        now_dt = datetime.now()
        today = now_dt.date()
        if target_day < today:
            return None
        if target_day > today:
            return start_min
        now_min = now_dt.hour * 60 + now_dt.minute
        floored = max(start_min, int((now_min + slot - 1) // slot) * slot)
        if floored >= end_min:
            return None
        return floored

    def _planning_initial_cursor(self, week_start: date) -> tuple[int, int]:
        start_min, end_min, _slot = self._planning_grid_metrics()
        week_start_dt = self._planning_week_start(week_start)
        now_dt = datetime.now()
        today = now_dt.date()
        week_end_dt = week_start_dt + timedelta(days=5)
        if week_end_dt < today:
            return 6, start_min
        if today < week_start_dt:
            return 0, start_min
        day_idx = max(0, min(5, (today - week_start_dt).days))
        cur_min = self._planning_now_floor_for_day((week_start_dt + timedelta(days=day_idx)).isoformat())
        if cur_min is None:
            return day_idx + 1, start_min
        return day_idx, min(cur_min, end_min)

    def _planning_assert_not_past(self, day_txt: str, start_min: int) -> None:
        floor_min = self._planning_now_floor_for_day(day_txt)
        if floor_min is None:
            try:
                target_day = datetime.fromisoformat(str(day_txt or "").strip()).strftime("%d/%m/%Y")
            except Exception:
                target_day = str(day_txt or "-")
            raise ValueError(f"Nao podes planear para uma data passada ({target_day}).")
        if start_min < floor_min:
            try:
                floor_dt = datetime.fromisoformat(str(day_txt or "").strip()).replace(
                    hour=floor_min // 60,
                    minute=floor_min % 60,
                    second=0,
                    microsecond=0,
                )
                floor_txt = floor_dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                floor_txt = f"{day_txt} {self.desktop_main.minutes_to_time(floor_min)}"
            raise ValueError(f"Na semana atual so podes planear a partir de {floor_txt}.")

    def planning_pending_rows(
        self,
        filter_text: str = "",
        state_filter: str = "Pendentes",
        operation: Any = "Corte Laser",
    ) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_norm = self.desktop_main.norm_text(state_filter or "Pendentes")
        op_txt = self._planning_normalize_operation(operation)
        planned_minutes: dict[tuple[str, str, str, str], int] = {}
        for row in list(self.ensure_data().get("plano", []) or []):
            if not isinstance(row, dict):
                continue
            if self._planning_row_operation(row) != op_txt:
                continue
            key = self._planning_item_op_key(row.get("encomenda", ""), row.get("material", ""), row.get("espessura", ""), op_txt)
            planned_minutes[key] = planned_minutes.get(key, 0) + self._planning_round_duration(row.get("duracao_min", 0))
        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(self.ensure_data().get("clientes", []) or [])
            if isinstance(row, dict)
        }
        rows = []
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            enc_state = str(enc.get("estado", "") or "").strip()
            enc_state_norm = self.desktop_main.norm_text(enc_state)
            if state_norm.startswith("pend") and ("concl" in enc_state_norm or "cancel" in enc_state_norm):
                continue
            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = clients.get(client_code, "")
            if op_txt == "Montagem":
                montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "")
                has_montagem = bool(list(self.desktop_main.encomenda_montagem_itens(enc) or []))
                show_montagem = False
                if state_norm.startswith("pend") and has_montagem and montagem_estado == "Pendente" and "montag" in enc_state_norm:
                    show_montagem = True
                elif state_norm.startswith("concl") and has_montagem and montagem_estado == "Consumida":
                    show_montagem = True
                elif state_norm not in ("concluidas",) and has_montagem and "montag" in enc_state_norm:
                    show_montagem = True
                if not show_montagem:
                    continue
                mat_name = self._planning_montagem_material()
                esp = self._planning_montagem_espessura()
                key = self._planning_item_op_key(enc.get("numero", ""), mat_name, esp, op_txt)
                tempo_total = self._planning_item_total_minutes(enc.get("numero", ""), mat_name, esp, operation=op_txt)
                tempo_planeado = max(0, planned_minutes.get(key, 0))
                tempo_restante = max(0, tempo_total - tempo_planeado)
                if state_norm != "concluidas" and tempo_restante <= 0:
                    continue
                shortages = self._order_montagem_shortages(enc)
                row = {
                    "numero": str(enc.get("numero", "") or "").strip(),
                    "cliente": " - ".join([x for x in [client_code, client_name] if x]).strip(" -"),
                    "material": mat_name,
                    "espessura": esp,
                    "tempo_min": float(tempo_restante if state_norm != "concluidas" else tempo_total),
                    "tempo_total_min": float(tempo_total),
                    "tempo_planeado_min": float(min(tempo_planeado, tempo_total)),
                    "estado": "Montagem pendente",
                    "laser_done": str(montagem_estado) == "Consumida",
                    "has_laser": False,
                    "operacao": op_txt,
                    "operation_done": str(montagem_estado) == "Consumida",
                    "is_montagem": True,
                    "stock_ready": not bool(shortages),
                    "shortage_count": len(shortages),
                    "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                    "chapa": "-",
                    "obs": self._planning_montagem_obs(enc),
                }
                if query and not any(query in str(value).lower() for value in row.values()):
                    continue
                rows.append(row)
                continue
            for mat in list(enc.get("materiais", []) or []):
                mat_name = str(mat.get("material", "") or "").strip()
                for esp_obj in list(mat.get("espessuras", []) or []):
                    esp = str(esp_obj.get("espessura", "") or "").strip()
                    if not self._planning_item_has_operation(str(enc.get("numero", "") or "").strip(), mat_name, esp, op_txt):
                        continue
                    operation_done = bool(self._planning_item_operation_done(str(enc.get("numero", "") or "").strip(), mat_name, esp, op_txt))
                    if state_norm.startswith("pend") and operation_done:
                        continue
                    if state_norm.startswith("concl") and not operation_done:
                        continue
                    key = self._planning_item_op_key(enc.get("numero", ""), mat_name, esp, op_txt)
                    tempo_total = self._planning_item_total_minutes(enc.get("numero", ""), mat_name, esp, operation=op_txt)
                    tempo_planeado = max(0, planned_minutes.get(key, 0))
                    tempo_restante = max(0, tempo_total - tempo_planeado)
                    if state_norm != "concluidas" and tempo_restante <= 0:
                        continue
                    row = {
                        "numero": str(enc.get("numero", "") or "").strip(),
                        "cliente": " - ".join([x for x in [client_code, client_name] if x]).strip(" -"),
                        "material": mat_name,
                        "espessura": esp,
                        "tempo_min": float(tempo_restante if state_norm != "concluidas" else tempo_total),
                        "tempo_total_min": float(tempo_total),
                        "tempo_planeado_min": float(min(tempo_planeado, tempo_total)),
                        "estado": enc_state,
                        "laser_done": operation_done,
                        "has_laser": self._planning_item_has_laser(str(enc.get("numero", "") or "").strip(), mat_name, esp),
                        "operacao": op_txt,
                        "operation_done": operation_done,
                        "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                        "chapa": self._order_reserved_sheet(str(enc.get("numero", "") or "").strip(), mat_name, esp),
                    }
                    if query and not any(query in str(value).lower() for value in row.values()):
                        continue
                    rows.append(row)
        rows.sort(
            key=lambda item: (
                item.get("data_entrega") or "9999-99-99",
                item.get("numero") or "",
                item.get("material") or "",
                self._parse_float(item.get("espessura", 0), 0),
            )
        )
        return rows

    def planning_blocked_windows(self) -> list[dict[str, Any]]:
        rows = []
        day_map = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sab", 6: "Dom"}
        for row in self._planning_blocked_windows():
            weekdays = list(row.get("weekdays", []) or [])
            rows.append(
                {
                    "id": str(row.get("id", "") or "").strip(),
                    "label": str(row.get("label", "") or "").strip(),
                    "start_min": int(row.get("start_min", 0) or 0),
                    "end_min": int(row.get("end_min", 0) or 0),
                    "start": self.desktop_main.minutes_to_time(int(row.get("start_min", 0) or 0)),
                    "end": self.desktop_main.minutes_to_time(int(row.get("end_min", 0) or 0)),
                    "weekdays": weekdays,
                    "dias_txt": ", ".join(day_map.get(day, str(day)) for day in weekdays),
                }
            )
        return rows

    def planning_set_blocked_windows(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        cleaned = []
        for index, row in enumerate(list(rows or [])):
            if not isinstance(row, dict):
                continue
            try:
                start_min = int(float(row.get("start_min", 0) or 0))
                end_min = int(float(row.get("end_min", 0) or 0))
            except Exception:
                continue
            if end_min <= start_min:
                continue
            weekdays = [int(v) for v in list(row.get("weekdays", [0, 1, 2, 3, 4, 5])) if str(v).strip().isdigit()]
            if not weekdays:
                weekdays = [0, 1, 2, 3, 4, 5]
            cleaned.append(
                {
                    "id": str(row.get("id", "") or f"PB{index+1:03d}").strip(),
                    "label": str(row.get("label", "") or "Bloqueio").strip(),
                    "start_min": start_min,
                    "end_min": end_min,
                    "weekdays": sorted(set(weekdays)),
                }
            )
        self.ensure_data()["plano_bloqueios"] = cleaned
        self._save(force=True)
        return self.planning_blocked_windows()

    def _planning_make_block(
        self,
        numero: str,
        material: str,
        espessura: str,
        operation: Any,
        day_txt: str,
        start_min: int,
        duration: int,
        color: str = "",
        posto: str = "",
    ) -> dict[str, Any]:
        op_txt = self._planning_normalize_operation(operation)
        color_txt = str(color or "").strip() or self._planning_item_color(numero, material, espessura)
        posto_txt = str(posto or "").strip() or self._planning_default_posto_for_operation(op_txt, numero)
        if op_txt == "Corte Laser":
            posto_txt = self._normalize_workcenter_value(posto_txt or self._order_workcenter(numero)) or self._planning_default_posto_for_operation(op_txt, numero)
        return {
            "id": f"PL{int(datetime.now().timestamp())}{len(self.ensure_data().get('plano', []))}",
            "encomenda": numero,
            "material": material,
            "espessura": espessura,
            "operacao": op_txt,
            "data": day_txt,
            "inicio": self.desktop_main.minutes_to_time(start_min),
            "duracao_min": duration,
            "color": color_txt,
            "chapa": self._order_reserved_sheet(numero, material, espessura),
            "planeamento_item": "|".join(self._planning_item_op_key(numero, material, espessura, op_txt)),
            "posto": posto_txt,
        }

    def planning_place_block(
        self,
        numero: str,
        material: str,
        espessura: str,
        day_txt: str,
        start_txt: str,
        operation: Any = "Corte Laser",
    ) -> dict[str, Any]:
        numero = str(numero or "").strip()
        material = str(material or "").strip()
        espessura = str(espessura or "").strip()
        day_txt = str(day_txt or "").strip()
        start_txt = str(start_txt or "").strip()
        op_txt = self._planning_normalize_operation(operation)
        pending = next(
            (
                row for row in self.planning_pending_rows(operation=op_txt)
                if str(row.get("numero", "") or "").strip() == numero
                and str(row.get("material", "") or "").strip() == material
                and str(row.get("espessura", "") or "").strip() == espessura
            ),
            None,
        )
        if pending is None:
            raise ValueError("Item do backlog n?o encontrado ou j? planeado.")
        if self._planning_is_duplicate(numero, material, espessura, operation=op_txt):
            raise ValueError("Este item j? est? planeado.")
        try:
            start_min = self.desktop_main.time_to_minutes(start_txt)
        except Exception as exc:
            raise ValueError("Hora inv?lida para planeamento.") from exc
        start_day, end_day, slot = self._planning_grid_metrics()
        duration = self._planning_round_duration(pending.get("tempo_min", 0))
        if duration <= 0:
            raise ValueError("Tempo inv?lido para o bloco.")
        if start_min < start_day or start_min + duration > end_day:
            raise ValueError("Hor?rio fora da grelha di?ria.")
        if start_min % slot != 0:
            raise ValueError("O inicio deve respeitar blocos de 30 minutos.")
        self._planning_assert_not_past(day_txt, start_min)
        if not self._planning_is_free(day_txt, start_min, start_min + duration, operation=op_txt):
            raise ValueError("Posi??o ocupada ou bloqueada no planeamento.")
        block = self._planning_make_block(
            numero,
            material,
            espessura,
            op_txt,
            day_txt,
            start_min,
            duration,
            color=self._planning_item_color(numero, material, espessura),
        )
        self.ensure_data().setdefault("plano", []).append(block)
        self._save(force=True)
        return dict(block)

    def planning_auto_plan(
        self,
        ordered_rows: list[dict[str, Any]],
        week_start: str | date | None = None,
        operation: Any = "Corte Laser",
    ) -> list[dict[str, Any]]:
        week_start_dt = self._planning_week_start(week_start)
        dates = [week_start_dt + timedelta(days=i) for i in range(6)]
        start_min, end_min, slot = self._planning_grid_metrics()
        placed: list[dict[str, Any]] = []
        cur_day_idx, cur_min = self._planning_initial_cursor(week_start_dt)
        exhausted = False
        op_txt = self._planning_normalize_operation(operation)

        if cur_day_idx >= len(dates):
            raise ValueError("A semana selecionada já ficou para trás ou não tem mais tempo útil disponível.")

        for raw in list(ordered_rows or []):
            row = dict(raw or {})
            numero = str(row.get("numero", "") or "").strip()
            material = str(row.get("material", "") or "").strip()
            espessura = str(row.get("espessura", "") or "").strip()
            if self._planning_is_duplicate(numero, material, espessura, operation=op_txt):
                continue
            duration = self._planning_round_duration(row.get("tempo_min", 0))
            if duration <= 0:
                raise ValueError(f"Tempo inv?lido em {numero} / {material} / {espessura}.")
            remaining = duration
            item_color = self._planning_item_color(numero, material, espessura)
            while remaining > 0:
                next_day_idx, day_txt, segment_start, segment_end = self._planning_next_free_segment(dates, cur_day_idx, cur_min)
                if not day_txt or segment_start is None or segment_end is None:
                    if not placed:
                        raise ValueError("Sem espa?o livre na semana para auto planeamento.")
                    exhausted = True
                    break
                free_minutes = max(0, int(segment_end - segment_start))
                if free_minutes <= 0:
                    exhausted = True
                    break
                chunk = min(remaining, free_minutes)
                if chunk % slot != 0:
                    chunk = max(slot, int(chunk // slot) * slot)
                block = self._planning_make_block(numero, material, espessura, op_txt, day_txt, segment_start, chunk, color=item_color)
                self.ensure_data().setdefault("plano", []).append(block)
                placed.append(block)
                remaining -= chunk
                cur_day_idx = next_day_idx
                cur_min = segment_start + chunk
                if cur_min >= end_min:
                    cur_day_idx += 1
                    cur_min = start_min
            if exhausted:
                break
        self._save(force=True)
        return placed

    def _planning_block_bounds(self, row: dict[str, Any]) -> tuple[datetime | None, datetime | None]:
        raw_date = str(row.get("data", "") or "").strip()
        raw_start = str(row.get("inicio", "") or "").strip()
        if not raw_date or not raw_start:
            return None, None
        try:
            start_dt = datetime.combine(datetime.fromisoformat(raw_date).date(), datetime.strptime(raw_start, "%H:%M").time())
        except Exception:
            return None, None
        duration = self._planning_round_duration(row.get("duracao_min", 0))
        end_dt = start_dt + timedelta(minutes=duration)
        return start_dt, end_dt

    def planning_laser_deadline_rows(self) -> list[dict[str, Any]]:
        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(self.ensure_data().get("clientes", []) or [])
            if isinstance(row, dict)
        }
        active_blocks = [dict(row) for row in list(self.ensure_data().get("plano", []) or []) if isinstance(row, dict)]
        active_blocks = [row for row in active_blocks if self._planning_row_matches_operation(row, "Corte Laser")]
        active_blocks.sort(key=lambda row: (str(row.get("data", "") or ""), str(row.get("inicio", "") or ""), str(row.get("encomenda", "") or "")))
        blocks_by_key: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
        for block in active_blocks:
            key = self._planning_item_key(block.get("encomenda", ""), block.get("material", ""), block.get("espessura", ""))
            blocks_by_key.setdefault(key, []).append(block)

        rows: list[dict[str, Any]] = []
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            numero = str(enc.get("numero", "") or "").strip()
            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = clients.get(client_code, "") or str(enc.get("cliente_nome", "") or "").strip()
            laser_groups = 0
            resolved_groups = 0
            partial_groups = 0
            planned_groups = 0
            block_count = 0
            total_minutes = 0
            planned_minutes = 0
            first_start: datetime | None = None
            last_end: datetime | None = None
            completion_marks: list[datetime] = []
            materials: list[str] = []

            for mat in list(enc.get("materiais", []) or []):
                mat_name = str(mat.get("material", "") or "").strip()
                for esp_obj in list(mat.get("espessuras", []) or []):
                    esp = str(esp_obj.get("espessura", "") or "").strip()
                    if not self._planning_item_has_laser(numero, mat_name, esp):
                        continue
                    laser_groups += 1
                    total = self._planning_item_total_minutes(numero, mat_name, esp)
                    total_minutes += total
                    key = self._planning_item_key(numero, mat_name, esp)
                    blocks = list(blocks_by_key.get(key, []) or [])
                    planned = min(total, sum(self._planning_round_duration(block.get("duracao_min", 0)) for block in blocks))
                    planned_minutes += planned
                    if blocks:
                        planned_groups += 1
                        block_count += len(blocks)
                        materials.append(f"{mat_name} {esp}mm")
                    for block in blocks:
                        start_dt, end_dt = self._planning_block_bounds(block)
                        if start_dt is not None and (first_start is None or start_dt < first_start):
                            first_start = start_dt
                        if end_dt is not None and (last_end is None or end_dt > last_end):
                            last_end = end_dt
                    item_ctx = {"encomenda": numero, "material": mat_name, "espessura": esp}
                    laser_done = bool(self.plan_actions._laser_done_for_item(self, item_ctx))
                    if laser_done:
                        resolved_groups += 1
                        raw_finished = str(esp_obj.get("laser_concluido_em", "") or "").strip()
                        if raw_finished:
                            try:
                                completion_marks.append(datetime.fromisoformat(raw_finished))
                            except Exception:
                                pass
                    elif planned >= total and total > 0:
                        resolved_groups += 1
                    elif planned > 0:
                        partial_groups += 1

            if laser_groups <= 0:
                continue
            if completion_marks:
                latest_completion = max(completion_marks)
                if last_end is None or latest_completion > last_end:
                    last_end = latest_completion
            if resolved_groups >= laser_groups and planned_groups <= 0 and completion_marks:
                status = "Laser concluído"
            elif resolved_groups >= laser_groups and laser_groups > 0:
                status = "Planeado completo"
            elif partial_groups > 0 or planned_groups > 0:
                status = "Planeado parcial"
            else:
                status = "Por planear"

            rows.append(
                {
                    "numero": numero,
                    "cliente": " - ".join([part for part in [client_code, client_name] if part]).strip(" -"),
                    "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                    "grupos_total": laser_groups,
                    "grupos_resolvidos": resolved_groups,
                    "grupos_planeados": planned_groups,
                    "grupos_parciais": partial_groups,
                    "planeado_min": planned_minutes,
                    "tempo_total_min": total_minutes,
                    "blocos": block_count,
                    "inicio_dt": first_start,
                    "fim_dt": last_end,
                    "inicio_txt": first_start.strftime("%d/%m/%Y %H:%M") if first_start is not None else "-",
                    "fim_txt": last_end.strftime("%d/%m/%Y %H:%M") if last_end is not None else "-",
                    "grupos_txt": f"{resolved_groups}/{laser_groups}",
                    "planeado_txt": f"{planned_minutes:.0f}/{total_minutes:.0f} min" if total_minutes > 0 else "-",
                    "estado": status,
                    "materiais_txt": ", ".join(materials[:4]) + ("..." if len(materials) > 4 else ""),
                }
            )
        rows.sort(
            key=lambda row: (
                0 if row.get("estado") == "Planeado completo" else (1 if row.get("estado") == "Planeado parcial" else 2),
                row.get("fim_dt") or datetime.max,
                row.get("data_entrega") or "9999-99-99",
                row.get("numero") or "",
            )
        )
        return rows

    def planning_render_laser_deadlines_pdf(self, output_path: str = "") -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        rows = self.planning_laser_deadline_rows()
        if output_path:
            path = Path(output_path)
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            path = Path(tempfile.gettempdir()) / f"lugest_prazos_laser_{stamp}.pdf"
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        width, height = landscape(A4)
        margin = 28
        c = pdf_canvas.Canvas(str(path), pagesize=landscape(A4))

        font_regular = "Helvetica"
        font_bold = "Helvetica-Bold"
        for name, file_name in (("SegoeUI", "segoeui.ttf"), ("SegoeUI-Bold", "segoeuib.ttf")):
            font_path = Path(r"C:\Windows\Fonts") / file_name
            if font_path.exists():
                try:
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont

                    pdfmetrics.registerFont(TTFont(name, str(font_path)))
                    if "Bold" in name:
                        font_bold = name
                    else:
                        font_regular = name
                except Exception:
                    pass

        def set_font(bold: bool, size: float) -> None:
            c.setFont(font_bold if bold else font_regular, size)

        def draw_header() -> float:
            c.setFillColor(palette["primary"])
            c.roundRect(margin, height - margin - 58, width - margin * 2, 58, 18, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 18)
            c.drawString(margin + 16, height - margin - 23, "Prazos Laser")
            set_font(False, 9)
            c.drawString(
                margin + 16,
                height - margin - 39,
                "Previsão de conclusão do corte laser por encomenda, baseada apenas no planeamento atual.",
            )
            generated = datetime.now().strftime("%d/%m/%Y %H:%M")
            c.drawRightString(width - margin - 16, height - margin - 23, generated)
            company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
            c.drawRightString(width - margin - 16, height - margin - 39, company)
            return height - margin - 76

        def draw_summary(y: float) -> float:
            total = len(rows)
            complete = sum(1 for row in rows if str(row.get("estado", "") or "") == "Planeado completo")
            partial = sum(1 for row in rows if str(row.get("estado", "") or "") == "Planeado parcial")
            pending = sum(1 for row in rows if str(row.get("estado", "") or "") == "Por planear")
            boxes = [
                ("Encomendas", str(total), palette["primary"]),
                ("Completas", str(complete), palette["success"]),
                ("Parciais", str(partial), palette["warning"]),
                ("Por planear", str(pending), palette["ink"]),
            ]
            box_w = (width - margin * 2 - 18) / 4
            x = margin
            for label, value, color in boxes:
                c.setFillColor(colors.white)
                c.setStrokeColor(palette["line"])
                c.roundRect(x, y - 46, box_w, 42, 14, stroke=1, fill=1)
                c.setFillColor(color)
                set_font(True, 9)
                c.drawString(x + 10, y - 16, label)
                set_font(True, 15)
                c.drawString(x + 10, y - 34, value)
                x += box_w + 6
            return y - 58

        def draw_table_header(y: float, columns: list[tuple[str, float]]) -> tuple[float, list[float], float]:
            total_w = width - margin * 2 - 18
            c.setFillColor(palette["primary_dark"])
            c.roundRect(margin, y - 22, width - margin * 2, 20, 10, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 8.5)
            x_positions: list[float] = []
            cursor = margin + 8
            for label, ratio in columns:
                x_positions.append(cursor)
                c.drawString(cursor + 3, y - 14, label)
                cursor += total_w * ratio
            return y - 26, x_positions, total_w

        columns = [
            ("Encomenda", 0.14),
            ("Cliente", 0.24),
            ("Entrega", 0.11),
            ("Grupos", 0.08),
            ("Planeado", 0.13),
            ("Fim laser", 0.18),
            ("Estado", 0.12),
        ]

        y = draw_header()
        y = draw_summary(y)
        row_h = 18
        table_min_y = 62
        row_index = 0
        while row_index < len(rows):
            if y < table_min_y + 80:
                c.showPage()
                y = draw_header()
            y, x_positions, total_w = draw_table_header(y, columns)
            while row_index < len(rows) and y - row_h >= table_min_y:
                row = rows[row_index]
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                state = str(row.get("estado", "") or "")
                if state == "Planeado completo":
                    fill_color = palette["primary_soft_2"]
                elif state == "Planeado parcial":
                    fill_color = colors.HexColor("#FFF8EB")
                elif state == "Laser concluído":
                    fill_color = colors.HexColor("#ECFDF3")
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, width - margin * 2, row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(row.get("numero", "") or "-"),
                    _pdf_clip_text(row.get("cliente", "-"), total_w * columns[1][1] - 10, font_regular, 8),
                    str(row.get("data_entrega", "") or "-"),
                    str(row.get("grupos_txt", "") or "-"),
                    str(row.get("planeado_txt", "") or "-"),
                    str(row.get("fim_txt", "") or "-"),
                    state or "-",
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    c.drawString(x_positions[idx] + 3, y - 10, value)
                y -= row_h
                row_index += 1
            y -= 8

        if not rows:
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(margin, y - 60, width - margin * 2, 52, 16, stroke=1, fill=1)
            c.setFillColor(palette["ink"])
            set_font(True, 12)
            c.drawString(margin + 14, y - 26, "Sem encomendas de laser planeadas neste momento.")

        c.setFillColor(palette["muted"])
        set_font(False, 8)
        c.drawString(margin, 28, "Planeamento atual do corte laser. Os postos seguintes serão integrados numa fase posterior.")
        c.drawRightString(width - margin, 28, f"LUGEST | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.save()
        return path

    def planning_open_laser_deadlines_pdf(self) -> Path:
        path = self.planning_render_laser_deadlines_pdf()
        os.startfile(str(path))
        return path

    def planning_shift_block(self, block_id: str, *, day_offset: int = 0, minutes_offset: int = 0) -> dict[str, Any]:
        target = next((row for row in list(self.ensure_data().get("plano", []) or []) if str(row.get("id", "") or "").strip() == str(block_id or "").strip()), None)
        if target is None:
            raise ValueError("Bloco de planeamento n?o encontrado.")
        current_date = datetime.fromisoformat(str(target.get("data", "") or "").strip()).date()
        current_start = self.desktop_main.time_to_minutes(str(target.get("inicio", "") or "").strip())
        duration = int(float(target.get("duracao_min", 0) or 0))
        new_date = current_date + timedelta(days=int(day_offset or 0))
        new_start = current_start + int(minutes_offset or 0)
        start_min, end_min, slot = self._planning_grid_metrics()
        if new_start < start_min or new_start + duration > end_min:
            raise ValueError("Novo hor?rio fora da grelha di?ria.")
        if new_start % slot != 0:
            raise ValueError("Novo hor?rio deve respeitar blocos de 30 minutos.")
        day_txt = new_date.isoformat()
        self._planning_assert_not_past(day_txt, new_start)
        if not self._planning_is_free(
            day_txt,
            new_start,
            new_start + duration,
            operation=self._planning_row_operation(target),
            ignore_id=str(target.get("id", "") or ""),
        ):
            raise ValueError("Posi??o ocupada ou bloqueada no planeamento.")
        target["data"] = day_txt
        target["inicio"] = self.desktop_main.minutes_to_time(new_start)
        target["chapa"] = self._order_reserved_sheet(str(target.get("encomenda", "") or "").strip(), str(target.get("material", "") or "").strip(), str(target.get("espessura", "") or "").strip())
        self._save(force=True)
        return dict(target)

    def planning_move_block_to(self, block_id: str, day_txt: str, start_txt: str) -> dict[str, Any]:
        target = next((row for row in list(self.ensure_data().get("plano", []) or []) if str(row.get("id", "") or "").strip() == str(block_id or "").strip()), None)
        if target is None:
            raise ValueError("Bloco de planeamento não encontrado.")
        day_txt = str(day_txt or "").strip()
        start_txt = str(start_txt or "").strip()
        if not day_txt or not start_txt:
            raise ValueError("Novo destino de planeamento inválido.")
        try:
            new_start = self.desktop_main.time_to_minutes(start_txt)
        except Exception as exc:
            raise ValueError("Hora inválida para planeamento.") from exc
        duration = int(float(target.get("duracao_min", 0) or 0))
        start_min, end_min, slot = self._planning_grid_metrics()
        if new_start < start_min or new_start + duration > end_min:
            raise ValueError("Novo horário fora da grelha diária.")
        if new_start % slot != 0:
            raise ValueError("Novo horário deve respeitar blocos de 30 minutos.")
        self._planning_assert_not_past(day_txt, new_start)
        if not self._planning_is_free(
            day_txt,
            new_start,
            new_start + duration,
            operation=self._planning_row_operation(target),
            ignore_id=str(target.get("id", "") or ""),
        ):
            raise ValueError("Posição ocupada ou bloqueada no planeamento.")
        target["data"] = day_txt
        target["inicio"] = self.desktop_main.minutes_to_time(new_start)
        target["chapa"] = self._order_reserved_sheet(
            str(target.get("encomenda", "") or "").strip(),
            str(target.get("material", "") or "").strip(),
            str(target.get("espessura", "") or "").strip(),
        )
        self._save(force=True)
        return dict(target)

    def planning_remove_block(self, block_id: str) -> None:
        block_txt = str(block_id or "").strip()
        before = len(list(self.ensure_data().get("plano", []) or []))
        self.ensure_data()["plano"] = [row for row in list(self.ensure_data().get("plano", []) or []) if str(row.get("id", "") or "").strip() != block_txt]
        if len(self.ensure_data()["plano"]) == before:
            raise ValueError("Bloco de planeamento n?o encontrado.")
        self._save(force=True)

    def planning_open_pdf(self, week_start: str | date | None = None, operation: Any = "Corte Laser") -> Path:
        week_start_dt = self._planning_week_start(week_start)
        op_txt = self._planning_normalize_operation(operation)
        filtered_data = dict(self.ensure_data())
        filtered_data["plano"] = [
            dict(row)
            for row in list(self.ensure_data().get("plano", []) or [])
            if isinstance(row, dict) and self._planning_row_matches_operation(row, op_txt)
        ]
        ctx = SimpleNamespace(
            data=filtered_data,
            p_week_start=week_start_dt,
            p_inicio=_ValueHolder("08:00"),
            p_fim=_ValueHolder("18:00"),
            planning_operation_label=op_txt,
        )
        ctx.get_plano_grid_metrics = lambda: (480, 1080, 30, 20, 6, 0, 0)
        ctx.plano_intervalo_bloqueado = lambda start_min, end_min: self._planning_interval_blocked(start_min, end_min)
        ctx.get_chapa_reservada = lambda numero, material=None, espessura=None: self._order_reserved_sheet(numero, material or "", espessura or "")
        path = Path(tempfile.gettempdir()) / "lugest_plano.pdf"
        self.plan_actions.preview_plano_a4(ctx)
        return path

    def finance_dashboard(self, ano: str = "Todos") -> dict[str, Any]:
        data = self.ensure_data()
        year_filter = str(ano or "Todos").strip()
        all_years: set[str] = set()
        valor_produtos = 0.0
        compras_produtos_total = 0.0
        compras_materias_total = 0.0
        compras_fornecedor_totais: dict[str, float] = {}
        compras_mes_totais: dict[str, float] = {}
        produtos_rows = []
        for prod in list(data.get("produtos", []) or []):
            qty = self._parse_float(prod.get("qty", 0), 0)
            if qty <= 0:
                continue
            price_unit = self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0)
            total = round(qty * price_unit, 2)
            valor_produtos += total
            produtos_rows.append(
                {
                    "codigo": str(prod.get("codigo", "") or "").strip(),
                    "descricao": str(prod.get("descricao", "") or "").strip(),
                    "qty": round(qty, 2),
                    "preco_unid": round(price_unit, 4),
                    "valor": total,
                }
            )
        valor_materias = 0.0
        materias_rows = []
        for mat in list(data.get("materiais", []) or []):
            qty = self._parse_float(mat.get("quantidade", 0), 0)
            if qty <= 0:
                continue
            price_unit = self._parse_float(self.materia_actions._materia_preco_unid_record(mat), 0)
            total = round(qty * price_unit, 2)
            valor_materias += total
            materias_rows.append(
                {
                    "id": str(mat.get("id", "") or "").strip(),
                    "material": str(mat.get("material", "") or "").strip(),
                    "espessura": str(mat.get("espessura", "") or "").strip(),
                    "qty": round(qty, 2),
                    "preco_unid": round(price_unit, 4),
                    "valor": total,
                }
            )
        compras_rows = []
        compras_materias_rows = []
        compras_produtos_rows = []
        valor_ne_aprovadas = 0.0
        for note in list(data.get("notas_encomenda", []) or []):
            estado_txt = str(note.get("estado", "") or "").strip()
            estado_norm = self.desktop_main.norm_text(estado_txt)
            note_date = str(note.get("data_entrega", "") or note.get("data_documento", "") or "").strip()
            note_year = note_date[:4] if len(note_date) >= 4 and note_date[:4].isdigit() else ""
            if note_year:
                all_years.add(note_year)
            total_note = round(self._parse_float(note.get("total", 0), 0), 2)
            if "aprov" in estado_norm and (year_filter.lower() in ("todos", "todas", "all", "") or note_year == year_filter):
                valor_ne_aprovadas += total_note
            for line in list(note.get("linhas", []) or []):
                qtd_tot = self._parse_float(line.get("qtd", 0), 0)
                qtd_ent = self._parse_float(line.get("qtd_entregue", 0), 0)
                entregue = bool(line.get("entregue") or line.get("_stock_in"))
                qtd_hist = qtd_ent if qtd_ent > 0 else (qtd_tot if entregue else 0.0)
                if qtd_hist <= 0:
                    continue
                line_date = str(line.get("data_doc_entrega", "") or line.get("data_entrega_real", "") or note_date or "").strip()
                line_year = line_date[:4] if len(line_date) >= 4 and line_date[:4].isdigit() else ""
                if line_year:
                    all_years.add(line_year)
                if year_filter.lower() not in ("todos", "todas", "all", "") and line_year != year_filter:
                    continue
                preco = self._parse_float(line.get("preco", 0), 0)
                total_l = round(qtd_hist * preco, 2)
                compras_rows.append(
                    {
                        "data": line_date,
                        "ne": str(note.get("numero", "") or "").strip(),
                        "fornecedor": str(note.get("fornecedor", "") or "").strip(),
                        "artigo": str(line.get("descricao", "") or line.get("ref", "") or "").strip(),
                        "qtd": round(qtd_hist, 2),
                        "preco": round(preco, 4),
                        "total": total_l,
                        "estado": estado_txt,
                        "origem": str(line.get("origem", "") or "").strip(),
                    }
                )
                fornecedor_nome = str(note.get("fornecedor", "") or "").strip() or "Sem fornecedor"
                compras_fornecedor_totais[fornecedor_nome] = round(compras_fornecedor_totais.get(fornecedor_nome, 0.0) + total_l, 2)
                mes_key = "-"
                if len(line_date) >= 7:
                    mes_key = line_date[:7]
                compras_mes_totais[mes_key] = round(compras_mes_totais.get(mes_key, 0.0) + total_l, 2)
                origem_norm = self.desktop_main.norm_text(line.get("origem", ""))
                if "mater" in origem_norm:
                    compras_materias_total += total_l
                    compras_materias_rows.append(dict(compras_rows[-1]))
                else:
                    compras_produtos_total += total_l
                    compras_produtos_rows.append(dict(compras_rows[-1]))
        status_counts = {"Preparacao": 0, "Montagem": 0, "Em producao": 0, "Em pausa": 0, "Avaria": 0, "Concluida": 0}
        montagem_alertas = []
        for enc in list(data.get("encomendas", []) or []):
            estado = str(enc.get("estado", "") or "").strip()
            norm = self.desktop_main.norm_text(estado)
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "").strip()
            shortages = self._order_montagem_shortages(enc)
            if montagem_estado == "Pendente":
                cliente_codigo = str(enc.get("cliente", "") or "").strip()
                cliente_nome = next(
                    (
                        str(row.get("nome", "") or "").strip()
                        for row in list(data.get("clientes", []) or [])
                        if str(row.get("codigo", "") or "").strip() == cliente_codigo
                    ),
                    "",
                )
                stock_txt = "Stock OK"
                shortage_total = round(sum(self._parse_float(row.get("qtd_em_falta", 0), 0) for row in shortages), 2)
                supplier_options: list[str] = []
                for shortage in shortages:
                    supplier_txt = str(shortage.get("fornecedor_sugerido", "") or "").strip()
                    if supplier_txt and supplier_txt.lower() not in [value.lower() for value in supplier_options]:
                        supplier_options.append(supplier_txt)
                supplier_txt = "Por validar"
                if supplier_options:
                    supplier_txt = supplier_options[0]
                    if len(supplier_options) > 1:
                        supplier_txt += f" +{len(supplier_options) - 1}"
                if shortages:
                    first = shortages[0]
                    stock_txt = f"Falta {self._fmt(first.get('qtd_em_falta', 0))} {first.get('produto_codigo', '-')}"
                    if len(shortages) > 1:
                        stock_txt += f" +{len(shortages) - 1}"
                montagem_alertas.append(
                    {
                        "numero": str(enc.get("numero", "") or "").strip(),
                        "cliente": " - ".join([x for x in [cliente_codigo, cliente_nome] if x]).strip(" -"),
                        "montagem": str(self.desktop_main.encomenda_montagem_resumo(enc) or "Montagem final").strip(),
                        "tempo_min": round(self._parse_float(self.desktop_main.encomenda_montagem_tempo_min(enc), 0), 1),
                        "qtd_falta": shortage_total,
                        "fornecedor": supplier_txt if shortages else "-",
                        "stock": stock_txt,
                        "shortage_count": len(shortages),
                        "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                    }
                )
            if "avari" in norm:
                status_counts["Avaria"] += 1
            elif "montag" in norm:
                status_counts["Montagem"] += 1
            elif "produc" in norm or "curso" in norm:
                status_counts["Em producao"] += 1
            elif "paus" in norm or "interromp" in norm:
                status_counts["Em pausa"] += 1
            elif "concl" in norm:
                status_counts["Concluida"] += 1
            else:
                status_counts["Preparacao"] += 1
        montagem_alertas.sort(
            key=lambda item: (
                -int(item.get("shortage_count", 0) or 0),
                str(item.get("data_entrega", "") or "9999-99-99"),
                str(item.get("numero", "") or ""),
            )
        )
        produtos_rows.sort(key=lambda item: item.get("valor", 0), reverse=True)
        materias_rows.sort(key=lambda item: item.get("valor", 0), reverse=True)
        compras_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_materias_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_produtos_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_fornecedor_rows = sorted(
            [{"fornecedor": key, "total": value} for key, value in compras_fornecedor_totais.items()],
            key=lambda item: item.get("total", 0),
            reverse=True,
        )
        compras_mes_rows = sorted(
            [{"mes": key, "total": value} for key, value in compras_mes_totais.items()],
            key=lambda item: str(item.get("mes", "") or ""),
            reverse=True,
        )
        subtitle_suffix = f"Ano {year_filter}" if year_filter.lower() not in ("todos", "todas", "all", "") else "Todos os anos"
        return {
            "cards": [
                {"title": "Stock MP", "value": self._fmt_eur(valor_materias), "subtitle": f"{len(data.get('materiais', []))} referencias", "tone": "warning"},
                {"title": "Stock Produtos", "value": self._fmt_eur(valor_produtos), "subtitle": f"{len(data.get('produtos', []))} refs | montagem {len(montagem_alertas)}", "tone": "success"},
                {"title": "Compras MP", "value": self._fmt_eur(compras_materias_total), "subtitle": subtitle_suffix, "tone": "warning"},
                {"title": "Compras Produtos", "value": self._fmt_eur(compras_produtos_total), "subtitle": subtitle_suffix, "tone": "success"},
                {"title": "Stock Total", "value": self._fmt_eur(valor_produtos + valor_materias), "subtitle": "Matéria-prima + produto acabado", "tone": "info"},
                {"title": "NE Aprovadas", "value": self._fmt_eur(valor_ne_aprovadas), "subtitle": subtitle_suffix, "tone": "default"},
            ],
            "order_status": [{"estado": key, "total": value} for key, value in status_counts.items()],
            "top_materias": materias_rows[:10],
            "top_produtos": produtos_rows[:10],
            "compras": compras_rows[:18],
            "compras_materias": compras_materias_rows[:18],
            "compras_produtos": compras_produtos_rows[:18],
            "compras_por_fornecedor": compras_fornecedor_rows[:12],
            "compras_por_mes": compras_mes_rows[:12],
            "montagem_alertas": montagem_alertas[:12],
            "years": sorted(all_years, reverse=True),
            "selected_year": year_filter or "Todos",
        }

    def operational_dashboard(self, ano: str = "Todos") -> dict[str, Any]:
        data = self.ensure_data()
        year_filter = str(ano or "Todos").strip()
        year_txt = "" if year_filter.lower() in ("todos", "todas", "all", "") else year_filter
        today = date.today()
        today_iso = today.isoformat()

        def _matches_year(enc: dict[str, Any]) -> bool:
            if not year_txt:
                return True
            candidates = [
                str(enc.get("data_entrega", "") or "").strip(),
                str(enc.get("data_criacao", "") or "").strip()[:10],
            ]
            for value in candidates:
                if len(value) >= 4 and value[:4] == year_txt:
                    return True
            return False

        def _safe_date_txt(value: Any) -> date | None:
            txt = str(value or "").strip()
            if not txt:
                return None
            try:
                return datetime.fromisoformat(txt[:10]).date()
            except Exception:
                return None

        def _phase_meta(
            *,
            enc_state: str,
            has_montagem: bool,
            montagem_estado: str,
            laser_status: str,
            shipping_status: str,
            has_guide: bool,
            trip_number: str,
            trip_state: str,
            transport_pending: bool,
            delivered: bool,
        ) -> tuple[str, str]:
            enc_norm = self.desktop_main.norm_text(enc_state)
            laser_norm = self.desktop_main.norm_text(laser_status)
            shipping_norm = self.desktop_main.norm_text(shipping_status)
            trip_norm = self.desktop_main.norm_text(trip_state)
            montagem_norm = self.desktop_main.norm_text(montagem_estado)
            if delivered or "entreg" in trip_norm:
                return "Entregue", "success"
            if trip_number:
                return "Em transporte", "info"
            if transport_pending:
                return "A aguardar transporte", "warning"
            if "totalmente expedida" in shipping_norm or has_guide:
                return "Expedição", "info"
            if has_montagem and "pendente" in montagem_norm and "concluido" in laser_norm:
                return "Montagem", "warning"
            if "concluido" in laser_norm:
                return "Pronta para expedição", "success"
            if "completo" in laser_norm:
                return "Laser planeado", "info"
            if "parcial" in laser_norm:
                return "Laser parcial", "warning"
            if "planear" in laser_norm:
                return "Preparação", "default"
            if "montag" in enc_norm:
                return "Montagem", "warning"
            if "produc" in enc_norm or "curso" in enc_norm:
                return "Em produção", "info"
            if "concl" in enc_norm:
                return "Concluída", "success"
            return "Preparação", "default"

        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if isinstance(row, dict)
        }
        deadline_rows = {
            str(row.get("numero", "") or "").strip(): dict(row)
            for row in list(self.planning_laser_deadline_rows() or [])
            if isinstance(row, dict)
        }
        delay_payload = self.pulse_plan_delay_rows(period_days=60, year_filter=year_txt or None, encomenda="Todas")
        delay_open = {
            str(item.get("numero", "") or "").strip(): dict(item)
            for item in list(delay_payload.get("items", []) or [])
            if isinstance(item, dict) and not bool(item.get("acknowledged"))
        }
        pending_transport_rows = {
            str(row.get("numero", "") or "").strip(): dict(row)
            for row in list(self.transport_pending_orders("") or [])
            if isinstance(row, dict)
        }
        active_trips: dict[str, dict[str, str]] = {}
        for trip in list(data.get("transportes", []) or []):
            if not isinstance(trip, dict):
                continue
            trip_num = str(trip.get("numero", "") or "").strip()
            trip_state = str(trip.get("estado", "") or "Planeado").strip() or "Planeado"
            if not trip_num or "anulad" in self.desktop_main.norm_text(trip_state):
                continue
            for stop in list(trip.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
                if not enc_num:
                    continue
                active_trips[enc_num] = {
                    "numero": trip_num,
                    "estado": self._transport_stop_state(stop, trip_state),
                }

        rows: list[dict[str, Any]] = []
        action_rows: list[dict[str, Any]] = []
        logistics_rows: list[dict[str, Any]] = []
        phase_counts: dict[str, int] = {}

        for enc in list(data.get("encomendas", []) or []):
            if not isinstance(enc, dict) or not _matches_year(enc):
                continue
            numero = str(enc.get("numero", "") or "").strip()
            if not numero:
                continue
            try:
                self.desktop_main.update_estado_expedicao_encomenda(enc)
            except Exception:
                pass

            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = clients.get(client_code, "") or str(enc.get("cliente_nome", "") or "").strip()
            client_display = " - ".join(part for part in [client_code, client_name] if part).strip(" -") or client_code or "-"

            deadline = dict(deadline_rows.get(numero, {}) or {})
            delay = dict(delay_open.get(numero, {}) or {})
            latest_guide = self._transport_latest_guide_for_order(numero) or {}
            active_trip = dict(active_trips.get(numero, {}) or {})
            trip_number = str(active_trip.get("numero", "") or str(enc.get("transporte_numero", "") or "")).strip()
            trip_state = str(active_trip.get("estado", "") or str(enc.get("estado_transporte", "") or "")).strip()
            has_guide = bool(str(latest_guide.get("numero", "") or "").strip())
            shipping_status = str(enc.get("estado_expedicao", "Não expedida") or "Não expedida").strip()
            transport_pending = numero in pending_transport_rows
            delivery_date = _safe_date_txt(enc.get("data_entrega", ""))
            delivery_txt = str(enc.get("data_entrega", "") or "").strip() or "-"
            delivery_overdue = bool(delivery_date and delivery_date < today and "entreg" not in self.desktop_main.norm_text(trip_state))

            montagem_items = list(self.desktop_main.encomenda_montagem_itens(enc) or [])
            has_montagem = bool(montagem_items)
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "Não aplicável").strip()
            montagem_shortages = list(self._order_montagem_shortages(enc) or []) if has_montagem else []
            shortage_count = len(montagem_shortages)

            laser_status = str(deadline.get("estado", "") or ("Sem laser" if not deadline else "-")).strip() or "-"
            laser_plan_txt = str(deadline.get("planeado_txt", "") or "-").strip() or "-"
            laser_end_txt = str(deadline.get("fim_txt", "") or "-").strip() or "-"
            phase_label, phase_tone = _phase_meta(
                enc_state=str(enc.get("estado", "") or "").strip(),
                has_montagem=has_montagem,
                montagem_estado=montagem_estado,
                laser_status=laser_status,
                shipping_status=shipping_status,
                has_guide=has_guide,
                trip_number=trip_number,
                trip_state=trip_state,
                transport_pending=transport_pending,
                delivered=bool("entreg" in self.desktop_main.norm_text(trip_state)),
            )
            phase_counts[phase_label] = phase_counts.get(phase_label, 0) + 1

            signal = "OK"
            signal_tone = "success" if phase_tone == "success" else "default"
            next_action = "-"
            if delivery_overdue:
                signal = "Entrega ultrapassada"
                signal_tone = "danger"
                next_action = "Rever prioridade real e contactar cliente se necessário."
            elif delay:
                signal = "Fora do planeamento"
                signal_tone = "danger"
                next_action = "Rever o plano do laser ou justificar o atraso."
            elif shortage_count > 0:
                signal = f"Falta montagem ({shortage_count})"
                signal_tone = "danger"
                next_action = "Validar stock e gerar reposição de montagem."
            elif transport_pending:
                signal = "Sem transporte"
                signal_tone = "warning"
                next_action = "Requisitar transporte ou subcontrato."
            elif "planear" in self.desktop_main.norm_text(laser_status):
                signal = "Por planear"
                signal_tone = "warning"
                next_action = "Planear o corte laser."
            elif "parcial" in self.desktop_main.norm_text(laser_status):
                signal = "Planeamento parcial"
                signal_tone = "warning"
                next_action = "Completar o planeamento do laser."
            elif has_montagem and "pendente" in self.desktop_main.norm_text(montagem_estado):
                signal = "Montagem pendente"
                signal_tone = "info"
                next_action = "Preparar consumos e fechar montagem."
            elif has_guide and not trip_number:
                signal = "Guia emitida"
                signal_tone = "info"
                next_action = "Confirmar saída / transporte."
            elif trip_number:
                signal = trip_state or "Em transporte"
                signal_tone = "info"
                next_action = "Acompanhar entrega."
            elif phase_label == "Pronta para expedição":
                signal = "Pronta a expedir"
                signal_tone = "success"
                next_action = "Emitir guia ou carregar transporte."

            row = {
                "numero": numero,
                "cliente": client_display,
                "estado": str(enc.get("estado", "") or "").strip() or "-",
                "fase": phase_label,
                "fase_tone": phase_tone,
                "laser": laser_status,
                "laser_planeado": laser_plan_txt,
                "laser_fim": laser_end_txt,
                "montagem": montagem_estado if has_montagem else "Não aplicável",
                "expedicao": shipping_status or "-",
                "guia_numero": str(latest_guide.get("numero", "") or "").strip(),
                "transporte_numero": trip_number,
                "transporte_estado": trip_state or "-",
                "transportadora": str(enc.get("transportadora_nome", "") or pending_transport_rows.get(numero, {}).get("transportadora_nome", "") or "-").strip() or "-",
                "zona": str(enc.get("zona_transporte", "") or pending_transport_rows.get(numero, {}).get("zona_transporte", "") or "-").strip() or "-",
                "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", pending_transport_rows.get(numero, {}).get("peso_bruto_kg", 0)), 0), 2),
                "paletes": round(self._parse_float(enc.get("paletes", pending_transport_rows.get(numero, {}).get("paletes", 0)), 0), 2),
                "entrega": delivery_txt,
                "sinal": signal,
                "signal_tone": signal_tone,
                "next_action": next_action,
                "delay_open": bool(delay),
                "delivery_overdue": delivery_overdue,
            }
            rows.append(row)
            if signal != "OK" or next_action != "-":
                action_rows.append(
                    {
                        "numero": numero,
                        "cliente": client_display,
                        "motivo": signal,
                        "acao": next_action,
                        "entrega": delivery_txt,
                        "tone": signal_tone,
                    }
                )
            if has_guide or trip_number or transport_pending or self._transport_is_own_cargo(enc):
                logistics_rows.append(
                    {
                        "numero": numero,
                        "cliente": client_display,
                        "guia": str(latest_guide.get("numero", "") or "-").strip() or "-",
                        "transporte": trip_number or "-",
                        "transportadora": str(row.get("transportadora", "-") or "-"),
                        "zona": str(row.get("zona", "-") or "-"),
                        "peso": f"{float(row.get('peso_bruto_kg', 0) or 0):.1f} kg",
                        "estado": signal if trip_number or transport_pending or has_guide else shipping_status or "-",
                        "tone": signal_tone if signal_tone != "default" else phase_tone,
                    }
                )

        phase_priority = {
            "Preparação": 0,
            "Laser parcial": 1,
            "Laser planeado": 2,
            "Montagem": 3,
            "Pronta para expedição": 4,
            "Expedição": 5,
            "A aguardar transporte": 6,
            "Em transporte": 7,
            "Entregue": 8,
            "Concluída": 9,
        }
        tone_priority = {"danger": 0, "warning": 1, "info": 2, "success": 3, "default": 4}
        rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("signal_tone", "default")), 4),
                0 if bool(row.get("delivery_overdue")) else 1,
                str(row.get("entrega", "") or "9999-99-99"),
                phase_priority.get(str(row.get("fase", "") or ""), 99),
                str(row.get("numero", "") or ""),
            )
        )
        action_rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("tone", "default")), 4),
                str(row.get("entrega", "") or "9999-99-99"),
                str(row.get("numero", "") or ""),
            )
        )
        logistics_rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("tone", "default")), 4),
                str(row.get("numero", "") or ""),
            )
        )

        open_orders = sum(1 for row in rows if str(row.get("fase", "") or "") not in {"Entregue", "Concluída"})
        risk_count = sum(1 for row in rows if str(row.get("signal_tone", "") or "") == "danger")
        ready_shipping = sum(1 for row in rows if str(row.get("fase", "") or "") in {"Pronta para expedição", "Expedição"})
        waiting_transport = sum(1 for row in rows if str(row.get("fase", "") or "") == "A aguardar transporte")
        in_transport = sum(1 for row in rows if str(row.get("fase", "") or "") == "Em transporte")
        montagem_pending = sum(1 for row in rows if "Montagem" in str(row.get("fase", "") or ""))

        cards = [
            {"title": "Encomendas abertas", "value": str(open_orders), "subtitle": f"{len(rows)} no horizonte atual", "tone": "info"},
            {"title": "Em risco", "value": str(risk_count), "subtitle": "Atraso ao plano, falta ou entrega ultrapassada", "tone": "danger" if risk_count else "success"},
            {"title": "Prontas a expedir", "value": str(ready_shipping), "subtitle": "Laser e/ou montagem já resolvidos", "tone": "success" if ready_shipping else "default"},
            {"title": "A aguardar transporte", "value": str(waiting_transport), "subtitle": "Carga nossa sem viagem ativa", "tone": "warning" if waiting_transport else "default"},
            {"title": "Em transporte", "value": str(in_transport), "subtitle": "Viagens em curso", "tone": "info" if in_transport else "default"},
            {"title": "Montagem pendente", "value": str(montagem_pending), "subtitle": "Itens ainda por fechar em montagem", "tone": "warning" if montagem_pending else "default"},
        ]
        phase_rows = [
            {"fase": key, "total": value}
            for key, value in sorted(
                phase_counts.items(),
                key=lambda item: (phase_priority.get(str(item[0] or ""), 99), str(item[0] or "")),
            )
        ]
        return {
            "cards": cards,
            "phase_rows": phase_rows,
            "order_rows": rows[:24],
            "action_rows": action_rows[:14],
            "logistics_rows": logistics_rows[:14],
            "updated_at": str(self.desktop_main.now_iso() or "").strip(),
            "selected_year": year_filter or "Todos",
        }

    def dashboard_counts(self) -> list[dict[str, str]]:
        data = self.ensure_data()
        encomendas_abertas = sum(1 for enc in data.get("encomendas", []) if "concl" not in str(enc.get("estado", "")).lower())
        encomendas_montagem = sum(1 for enc in data.get("encomendas", []) if "montag" in self.desktop_main.norm_text(enc.get("estado", "")))
        material_disponivel = sum(
            max(0.0, self._parse_float(m.get("quantidade", 0), 0) - self._parse_float(m.get("reservado", 0), 0))
            for m in data.get("materiais", [])
        )
        return [
            {"title": "Materias", "value": str(len(data.get("materiais", []))), "subtitle": f"Disponivel {self._fmt(material_disponivel)}"},
            {"title": "Encomendas", "value": str(len(data.get("encomendas", []))), "subtitle": f"Abertas {encomendas_abertas} | Montagem {encomendas_montagem}"},
            {"title": "Clientes", "value": str(len(data.get("clientes", []))), "subtitle": "Base ativa"},
            {"title": "Fornecedores", "value": str(len(data.get("fornecedores", []))), "subtitle": "Compras e stock"},
        ]
