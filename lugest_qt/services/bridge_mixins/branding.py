from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


class BrandingBackendMixin:
    """Legacy adapter for branding; see BACKEND_GUIDE.md."""

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
        for fallback in ("Logos/lg.png", "lg.png", "Logos/image (1).jpg", "Logos/image.jpg", "Logos/logo.png", "logo.jpg", "logo.png", "Logos/logo(1).jpg"):
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
            "primary_color": self._normalize_pdf_primary_color(cfg.get("primary_color", "#00A6A6")),
            "logo_scale_pct": max(50, min(250, int(self._parse_float(cfg.get("logo_scale_pct", 100), 100)))),
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

    def _branding_logo_scale_factor(self) -> float:
        try:
            scale_pct = max(50.0, min(250.0, float(self.branding_settings().get("logo_scale_pct", 100) or 100)))
        except Exception:
            scale_pct = 100.0
        return scale_pct / 100.0

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
            color_value = str(payload.get("primary_color", "") or "").strip()
            if not re.fullmatch(r"#[0-9A-Fa-f]{6}", color_value):
                raise ValueError("A cor principal deve usar o formato hexadecimal #RRGGBB.")
            cfg["primary_color"] = color_value.upper()
        try:
            cfg["logo_scale_pct"] = max(50, min(250, int(float(payload.get("logo_scale_pct", cfg.get("logo_scale_pct", 100)) or 100))))
        except Exception:
            cfg["logo_scale_pct"] = 100

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

    def ensure_pdf_light_theme(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        current = dict(cfg.get("pdf", {}) or {})
        accent = self._normalize_pdf_primary_color(self.branding_settings().get("primary_color", "#00A6A6"))
        wanted = {
            "theme": "light_v2",
            "accent": accent,
            "ink": "#14212B",
            "large_color_blocks": False,
            "inventory_scan_schema": 1,
        }
        if current != wanted:
            cfg["pdf"] = wanted
            self._save_qt_config(cfg)
        return dict(cfg.get("pdf", wanted) or wanted)
