from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any


class PlanningDelaysBackendMixin:
    """Legacy adapter for planning delays; see BACKEND_GUIDE.md."""

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
