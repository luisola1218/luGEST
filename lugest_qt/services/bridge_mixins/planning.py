from __future__ import annotations

import os
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text


class PlanningBridgeMixin:
    """Planning board, scheduling, and planning PDF operations for the Qt bridge."""

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
        resource: str = "",
        ignore_id: str = "",
    ) -> bool:
        if self._planning_interval_blocked(start_min, end_min, day_txt):
            return False
        ignore = str(ignore_id or "").strip()
        op_txt = self._planning_normalize_operation(operation)
        resource_txt = self._normalize_workcenter_value(resource)
        for row in list(self.ensure_data().get("plano", []) or []):
            if str(row.get("data", "") or "").strip() != day_txt:
                continue
            if ignore and str(row.get("id", "") or "").strip() == ignore:
                continue
            if self._planning_row_operation(row) != op_txt:
                continue
            row_resource = self._planning_row_resource(row)
            if resource_txt and row_resource.lower() != resource_txt.lower():
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

    def _planning_item_piece_qty(self, numero: str, material: str, espessura: str) -> float:
        enc = self.get_encomenda_by_numero(str(numero or "").strip())
        esp_obj = self._planning_find_esp_obj(enc, material, espessura)
        if not isinstance(esp_obj, dict):
            return 0.0
        total = 0.0
        for piece in list(esp_obj.get("pecas", []) or []):
            total += max(
                0.0,
                self._parse_float(piece.get("quantidade_pedida", piece.get("qtd", piece.get("quantidade", 0))), 0),
            )
        return round(total, 4)

    def _planning_operation_plan_candidates(self, operation: Any) -> list[str]:
        op_txt = self._planning_normalize_operation(operation, default="")
        norm = self.desktop_main.norm_text(op_txt)
        candidates: list[str] = []

        def add(value: Any) -> None:
            key = str(value or "").strip().lower()
            if key and key not in candidates:
                candidates.append(key)

        add(norm)
        add(self.desktop_main.norm_text(operation or ""))
        if "laser" in norm:
            add("laser")
        if "quin" in norm:
            add("quinagem")
        if "sold" in norm or "serralh" in norm:
            add("soldadura")
        if "rosc" in norm:
            add("roscagem")
        if "embal" in norm or "exped" in norm:
            add("embalamento")
        if "laca" in norm or "pint" in norm:
            add("lacagem")
        if "maquin" in norm:
            add("maquinacao")
        if "mont" in norm:
            add("montagem")
        return candidates

    def _planning_default_operation_minutes_per_piece(self, operation: Any) -> float:
        raw_map = dict(self.ensure_data().get("tempos_operacao_planeada_min", {}) or {})
        normalized_map = {
            self.desktop_main.norm_text(key): self._parse_float(value, 0)
            for key, value in raw_map.items()
            if str(key or "").strip()
        }
        for candidate in self._planning_operation_plan_candidates(operation):
            if candidate in raw_map:
                value = self._parse_float(raw_map.get(candidate, 0), 0)
                if value > 0:
                    return value
            value = self._parse_float(normalized_map.get(candidate, 0), 0)
            if value > 0:
                return value
        return 0.0

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
        direct_total = self._parse_float(time_map.get(op_txt, 0), 0)
        if direct_total > 0:
            return self._planning_round_duration(direct_total)
        default_per_piece = self._planning_default_operation_minutes_per_piece(op_txt)
        piece_qty = self._planning_item_piece_qty(numero, material, espessura)
        if default_per_piece > 0 and piece_qty > 0:
            return self._planning_round_duration(default_per_piece * piece_qty)
        return 0

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
        legacy_palette = {"#fbecee", "#fde2e4", "#fff0f2", "#fce8ea", "#ffe6e9", "#f7dfe3"}
        palette = [
            "#93c5fd",
            "#86efac",
            "#fcd34d",
            "#fca5a5",
            "#c4b5fd",
            "#67e8f9",
            "#f9a8d4",
            "#fdba74",
            "#bef264",
            "#a5b4fc",
            "#fca5a5",
            "#99f6e4",
        ]
        for row in list(self.ensure_data().get("plano", []) or []):
            row_key = self._planning_item_key(row.get("encomenda", ""), row.get("material", ""), row.get("espessura", ""))
            if row_key != target:
                continue
            color = str(row.get("color", "") or "").strip()
            if color and color.lower() not in legacy_palette:
                return color
        key = "|".join(
            [
                str(numero or "").strip().upper(),
                str(material or "").strip().upper(),
                str(espessura or "").strip().upper(),
            ]
        )
        index = sum(ord(ch) for ch in key) % len(palette)
        return str(palette[index] or "#93c5fd")

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
        *,
        operation: Any = "Corte Laser",
        resource: str = "",
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
                if not self._planning_is_free(day_txt, local_cursor, local_cursor + slot, operation=operation, resource=resource):
                    local_cursor += slot
                    continue
                segment_start = local_cursor
                segment_end = local_cursor + slot
                while segment_end + slot <= end_min and self._planning_is_free(day_txt, segment_end, segment_end + slot, operation=operation, resource=resource):
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
        resource: str = "",
    ) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_norm = self.desktop_main.norm_text(state_filter or "Pendentes")
        op_txt = self._planning_normalize_operation(operation)
        resource_txt = self._normalize_workcenter_value(resource)
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
                assigned_resource = self._order_operation_resource(enc, mat_name, esp, op_txt)
                if resource_txt and assigned_resource.lower() != resource_txt.lower():
                    continue
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
                    "recurso": assigned_resource,
                    "posto_trabalho": assigned_resource,
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
                    assigned_resource = self._order_operation_resource(enc, mat_name, esp, op_txt)
                    if resource_txt and assigned_resource.lower() != resource_txt.lower():
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
                        "recurso": assigned_resource,
                        "posto_trabalho": assigned_resource,
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

    def planning_overview_data(
        self,
        week_start: str | date | None = None,
        operation: Any = "Corte Laser",
        resource: str = "",
    ) -> dict[str, Any]:
        week_start_dt = self._planning_week_start(week_start)
        week_dates = [week_start_dt + timedelta(days=i) for i in range(6)]
        week_end_dt = week_dates[-1]
        month_token = f"{datetime.now().year}-{datetime.now().month:02d}"
        op_txt = self._planning_normalize_operation(operation)
        resource_txt = self._normalize_workcenter_value(resource)
        active: list[dict[str, Any]] = []
        history: list[dict[str, Any]] = []
        total_active_minutes = 0.0
        week_active_minutes = 0.0

        def build_row(row: dict[str, Any], *, history_mode: bool = False) -> dict[str, Any]:
            color_txt = str(row.get("color", "") or "").strip()
            if color_txt.lower() in {"#fbecee", "#fde2e4", "#fff0f2", "#fce8ea", "#ffe6e9", "#f7dfe3"}:
                color_txt = ""
            payload = {
                "id": str(row.get("id", "") or "").strip(),
                "encomenda": str(row.get("encomenda", "") or "").strip(),
                "data": str(row.get("data", "") or "").strip(),
                "inicio": str(row.get("inicio", "") or "").strip(),
                "duracao_min": round(max(0.0, self._parse_float(row.get("duracao_min", 0), 0)), 1),
                "material": str(row.get("material", "") or "-").strip() or "-",
                "espessura": str(row.get("espessura", "") or "-").strip() or "-",
                "operacao": op_txt,
                "chapa": str(row.get("chapa", "") or "-").strip() or "-",
                "posto_trabalho": self._planning_row_resource(row),
                "maquina": self._planning_row_resource(row),
                "color": color_txt or self._planning_item_color(row.get("encomenda", ""), row.get("material", ""), row.get("espessura", "")),
            }
            if history_mode:
                payload["tempo_real_min"] = round(max(0.0, self._parse_float(row.get("tempo_real_min", 0), 0)), 1)
                payload["estado_final"] = str(row.get("estado_final", "") or "-").strip() or "-"
            return payload

        for row in list(self.ensure_data().get("plano", []) or []):
            if not isinstance(row, dict):
                continue
            if self._planning_row_operation(row) != op_txt:
                continue
            row_resource = self._planning_row_resource(row)
            if resource_txt and row_resource.lower() != resource_txt.lower():
                continue
            payload = build_row(row, history_mode=False)
            total_active_minutes += float(payload.get("duracao_min", 0) or 0)
            try:
                row_date = date.fromisoformat(str(payload.get("data", "") or "").strip())
            except Exception:
                row_date = None
            if row_date is not None and week_start_dt <= row_date <= week_end_dt:
                active.append(payload)
                week_active_minutes += float(payload.get("duracao_min", 0) or 0)

        for row in list(self.ensure_data().get("plano_hist", []) or []):
            if not isinstance(row, dict):
                continue
            if self._planning_row_operation(row) != op_txt:
                continue
            row_resource = self._planning_row_resource(row)
            if resource_txt and row_resource.lower() != resource_txt.lower():
                continue
            history.append(build_row(row, history_mode=True))

        active.sort(key=lambda row: (row.get("data", ""), row.get("inicio", ""), row.get("encomenda", "")))
        history.sort(key=lambda row: (row.get("data", ""), row.get("inicio", ""), row.get("encomenda", "")), reverse=True)
        month_hist = [row for row in history if str(row.get("data", "") or "").startswith(month_token)]
        return {
            "summary": {
                "blocos_ativos": len(active),
                "historico_mes": len(month_hist),
                "min_ativos": round(week_active_minutes, 1),
                "min_ativos_total": round(total_active_minutes, 1),
                "min_historico_mes": round(sum(max(0.0, self._parse_float(row.get("duracao_min", 0), 0)) for row in month_hist), 1),
                "week_start": week_start_dt.isoformat(),
                "week_end": week_end_dt.isoformat(),
                "week_label": f"{week_start_dt.strftime('%d/%m')} - {week_end_dt.strftime('%d/%m')}",
                "operacao": op_txt,
                "recurso": resource_txt,
            },
            "week_dates": [day.isoformat() for day in week_dates],
            "active": active[:160],
            "history": history[:220],
        }

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
        posto_txt = (
            self._normalize_workcenter_value(posto)
            or self._order_operation_resource(numero, material, espessura, op_txt)
            or self._planning_default_posto_for_operation(op_txt, numero)
        )
        block = {
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
        }
        return self._planning_apply_resource_to_row(block, posto_txt, op_txt)

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
        resource_txt = self._normalize_workcenter_value(pending.get("recurso", pending.get("posto_trabalho", "")))
        if duration <= 0:
            raise ValueError("Tempo inv?lido para o bloco.")
        if start_min < start_day or start_min + duration > end_day:
            raise ValueError("Hor?rio fora da grelha di?ria.")
        if start_min % slot != 0:
            raise ValueError("O inicio deve respeitar blocos de 30 minutos.")
        self._planning_assert_not_past(day_txt, start_min)
        if not self._planning_is_free(day_txt, start_min, start_min + duration, operation=op_txt, resource=resource_txt):
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
            posto=resource_txt,
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
            resource_txt = self._normalize_workcenter_value(row.get("recurso", row.get("posto_trabalho", "")))
            if self._planning_is_duplicate(numero, material, espessura, operation=op_txt):
                continue
            duration = self._planning_round_duration(row.get("tempo_min", 0))
            if duration <= 0:
                raise ValueError(f"Tempo inv?lido em {numero} / {material} / {espessura}.")
            remaining = duration
            item_color = self._planning_item_color(numero, material, espessura)
            while remaining > 0:
                next_day_idx, day_txt, segment_start, segment_end = self._planning_next_free_segment(
                    dates,
                    cur_day_idx,
                    cur_min,
                    operation=op_txt,
                    resource=resource_txt,
                )
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
                block = self._planning_make_block(numero, material, espessura, op_txt, day_txt, segment_start, chunk, color=item_color, posto=resource_txt)
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

    def planning_auto_plan_full_flow(
        self,
        ordered_rows: list[dict[str, Any]],
        week_start: str | date | None = None,
        operation: Any = "Corte Laser",
    ) -> dict[str, Any]:
        week_start_dt = self._planning_week_start(week_start)
        dates = [week_start_dt + timedelta(days=i) for i in range(6)]
        start_min, _end_min, _slot = self._planning_grid_metrics()
        main_day_idx, main_min = self._planning_initial_cursor(week_start_dt)
        base_day_idx, base_min = main_day_idx, main_min
        if main_day_idx >= len(dates):
            raise ValueError("A semana selecionada já ficou para trás ou não tem mais tempo útil disponível.")
        op_txt = self._planning_normalize_operation(operation)
        placed: list[dict[str, Any]] = []
        pending: list[dict[str, Any]] = []
        downstream_jobs: list[dict[str, Any]] = []

        def later_dt(left: datetime | None, right: datetime | None) -> datetime | None:
            if left is None:
                return right
            if right is None:
                return left
            return right if right > left else left

        for raw in list(ordered_rows or []):
            row = dict(raw or {})
            numero = str(row.get("numero", "") or "").strip()
            material = str(row.get("material", "") or "").strip()
            espessura = str(row.get("espessura", "") or "").strip()
            if not numero or not material or not espessura:
                continue
            if not self._planning_item_has_operation(numero, material, espessura, op_txt):
                continue
            if main_day_idx >= len(dates):
                pending.append(
                    {
                        "numero": numero,
                        "material": material,
                        "espessura": espessura,
                        "operacao": op_txt,
                        "recurso": self._order_operation_resource(numero, material, espessura, op_txt),
                        "remaining_min": self._planning_remaining_minutes(numero, material, espessura, operation=op_txt),
                    }
                )
                break
            main_anchor_dt = None
            if 0 <= main_day_idx < len(dates):
                main_anchor_dt = self._planning_slot_datetime(dates[main_day_idx].isoformat(), main_min)
            current_resource = self._normalize_workcenter_value(row.get("recurso", row.get("posto_trabalho", "")))
            current_result = self._planning_schedule_operation_blocks(
                numero,
                material,
                espessura,
                op_txt,
                dates,
                anchor_dt=main_anchor_dt,
                resource=current_resource,
            )
            placed.extend(list(current_result.get("placed", []) or []))
            main_day_idx = int(current_result.get("cursor_day_idx", main_day_idx) or 0)
            main_min = int(current_result.get("cursor_min", start_min) or start_min)
            if bool(current_result.get("exhausted")) and int(current_result.get("remaining_min", 0) or 0) > 0:
                pending.append(
                    {
                        "numero": numero,
                        "material": material,
                        "espessura": espessura,
                        "operacao": op_txt,
                        "recurso": str(current_result.get("resource", "") or current_resource),
                        "remaining_min": int(current_result.get("remaining_min", 0) or 0),
                    }
                )
                break

            chain_anchor = current_result.get("end_dt")
            sequence = self._planning_item_operation_sequence(numero, material, espessura, start_operation=op_txt)
            if op_txt in sequence:
                next_ops = sequence[sequence.index(op_txt) + 1 :]
            else:
                next_ops = sequence
            filtered_next_ops = [
                next_op
                for next_op in next_ops
                if self._planning_item_has_operation(numero, material, espessura, next_op)
            ]
            if filtered_next_ops:
                first_next_op = filtered_next_ops[0]
                downstream_jobs.append(
                    {
                        "numero": numero,
                        "material": material,
                        "espessura": espessura,
                        "sequence": filtered_next_ops,
                        "index": 0,
                        "anchor_dt": chain_anchor,
                        "resource": self._order_operation_resource(numero, material, espessura, first_next_op),
                        "data_entrega": str(row.get("data_entrega", "") or "").strip(),
                    }
                )

        initial_cursor_dt = None
        if 0 <= base_day_idx < len(dates):
            initial_cursor_dt = self._planning_slot_datetime(dates[base_day_idx].isoformat(), base_min)
        downstream_result = self._planning_schedule_followup_jobs(
            downstream_jobs,
            dates,
            initial_cursor_dt=initial_cursor_dt,
        )
        placed.extend(list(downstream_result.get("placed", []) or []))
        pending.extend(list(downstream_result.get("pending", []) or []))

        if placed:
            self._save(force=True)
        return {"placed": placed, "pending": pending}

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
        blocks_by_item: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
        for block in active_blocks:
            key = self._planning_item_key(block.get("encomenda", ""), block.get("material", ""), block.get("espessura", ""))
            blocks_by_item.setdefault(key, []).append(block)

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
            completed_groups = 0
            block_count = 0
            total_minutes = 0
            planned_minutes = 0
            first_start: datetime | None = None
            last_end: datetime | None = None
            last_item_desc = ""
            materials: list[str] = []

            for mat in list(enc.get("materiais", []) or []):
                mat_name = str(mat.get("material", "") or "").strip()
                for esp_obj in list(mat.get("espessuras", []) or []):
                    esp = str(esp_obj.get("espessura", "") or "").strip()
                    if not self._planning_item_has_laser(numero, mat_name, esp):
                        continue
                    laser_groups += 1
                    sequence = self._planning_item_operation_sequence(numero, mat_name, esp, start_operation="Corte Laser")
                    if not sequence:
                        sequence = ["Corte Laser"]
                    item_total = 0
                    item_planned = 0
                    item_first: datetime | None = None
                    item_end: datetime | None = None
                    item_has_activity = False
                    item_resolved = True
                    item_completed = True
                    key = self._planning_item_key(numero, mat_name, esp)
                    blocks = list(blocks_by_item.get(key, []) or [])
                    block_count += len(blocks)
                    for op_name in sequence:
                        if not self._planning_item_has_operation(numero, mat_name, esp, op_name):
                            continue
                        status = self._planning_item_operation_status(numero, mat_name, esp, op_name)
                        item_total += int(status.get("total_min", 0) or 0)
                        item_planned += int(status.get("planned_min", 0) or 0)
                        op_first = status.get("first_dt")
                        op_end = status.get("end_dt")
                        if op_first is not None and (item_first is None or op_first < item_first):
                            item_first = op_first
                        if op_end is not None and (item_end is None or op_end > item_end):
                            item_end = op_end
                        if bool(status.get("resolved")) or int(status.get("planned_min", 0) or 0) > 0:
                            item_has_activity = True
                        if not bool(status.get("resolved")):
                            item_resolved = False
                        if not self._planning_item_operation_done(numero, mat_name, esp, op_name):
                            item_completed = False
                    total_minutes += item_total
                    planned_minutes += item_planned
                    if item_first is not None and (first_start is None or item_first < first_start):
                        first_start = item_first
                    if item_end is not None and (last_end is None or item_end > last_end):
                        last_end = item_end
                        last_item_desc = f"{mat_name} {esp}mm".strip()
                    if item_has_activity:
                        planned_groups += 1
                        materials.append(f"{mat_name} {esp}mm")
                    if item_completed:
                        completed_groups += 1
                    elif item_resolved:
                        resolved_groups += 1
                    elif item_has_activity:
                        partial_groups += 1

            if laser_groups <= 0:
                continue
            if completed_groups >= laser_groups and laser_groups > 0:
                status = "Fluxo concluído"
            elif (completed_groups + resolved_groups) >= laser_groups and laser_groups > 0:
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
                    "grupos_resolvidos": completed_groups + resolved_groups,
                    "grupos_planeados": planned_groups,
                    "grupos_parciais": partial_groups,
                    "planeado_min": planned_minutes,
                    "tempo_total_min": total_minutes,
                    "blocos": block_count,
                    "inicio_dt": first_start,
                    "fim_dt": last_end,
                    "inicio_txt": first_start.strftime("%d/%m/%Y %H:%M") if first_start is not None else "-",
                    "fim_txt": last_end.strftime("%d/%m/%Y %H:%M") if last_end is not None else "-",
                    "ultimo_item_txt": last_item_desc or "-",
                    "grupos_txt": f"{completed_groups + resolved_groups}/{laser_groups}",
                    "planeado_txt": f"{planned_minutes:.0f}/{total_minutes:.0f} min" if total_minutes > 0 else "-",
                    "estado": status,
                    "materiais_txt": ", ".join(materials[:4]) + ("..." if len(materials) > 4 else ""),
                }
            )
        rows.sort(
            key=lambda row: (
                0
                if row.get("estado") == "Fluxo concluído"
                else (1 if row.get("estado") == "Planeado completo" else (2 if row.get("estado") == "Planeado parcial" else 3)),
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
            path = Path(tempfile.gettempdir()) / f"lugest_prazos_fluxo_{stamp}.pdf"
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
            c.drawString(margin + 16, height - margin - 23, "Prazo Final do Fluxo")
            set_font(False, 9)
            c.drawString(
                margin + 16,
                height - margin - 39,
                "Previsão de conclusão global da encomenda, considerando laser e operações seguintes com base no planeamento atual.",
            )
            generated = datetime.now().strftime("%d/%m/%Y %H:%M")
            c.drawRightString(width - margin - 16, height - margin - 23, generated)
            company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
            c.drawRightString(width - margin - 16, height - margin - 39, company)
            return height - margin - 76

        def draw_summary(y: float) -> float:
            total = len(rows)
            complete = sum(1 for row in rows if str(row.get("estado", "") or "") in {"Planeado completo", "Fluxo concluído"})
            partial = sum(1 for row in rows if str(row.get("estado", "") or "") == "Planeado parcial")
            pending = sum(1 for row in rows if str(row.get("estado", "") or "") == "Por planear")
            boxes = [
                ("Encomendas", str(total), palette["primary"]),
                ("Fluxo fechado", str(complete), palette["success"]),
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
            ("Fim fluxo", 0.18),
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
                elif state == "Fluxo concluído":
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
            c.drawString(margin + 14, y - 26, "Sem encomendas com fluxo planeado neste momento.")

        c.setFillColor(palette["muted"])
        set_font(False, 8)
        c.drawString(margin, 28, "Prazo final calculado a partir do planeamento atual e do histórico concluído por operação.")
        c.drawRightString(width - margin, 28, f"LUGEST | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.save()
        return path

    def planning_open_laser_deadlines_pdf(self) -> Path:
        path = self.planning_render_laser_deadlines_pdf()
        os.startfile(str(path))
        return path

    def _planning_order_flow_rows(self, numero: str) -> list[dict[str, Any]]:
        numero_txt = str(numero or "").strip()
        rows: list[dict[str, Any]] = []
        operation_index = {name: idx for idx, name in enumerate(self.planning_operation_options())}
        for bucket_name, source_label in (("plano", "Ativo"), ("plano_hist", "Histórico")):
            for row in list(self.ensure_data().get(bucket_name, []) or []):
                if not isinstance(row, dict):
                    continue
                if str(row.get("encomenda", "") or "").strip() != numero_txt:
                    continue
                start_dt, end_dt = self._planning_block_bounds(row)
                duration = self._planning_round_duration(row.get("duracao_min", 0))
                rows.append(
                    {
                        "id": str(row.get("id", "") or "").strip(),
                        "operacao": self._planning_row_operation(row),
                        "recurso": self._planning_row_resource(row),
                        "material": str(row.get("material", "") or "").strip() or "-",
                        "espessura": str(row.get("espessura", "") or "").strip() or "-",
                        "data": str(row.get("data", "") or "").strip(),
                        "inicio": str(row.get("inicio", "") or "").strip(),
                        "fim": end_dt.strftime("%H:%M") if end_dt is not None else "-",
                        "duracao_min": duration,
                        "chapa": str(row.get("chapa", "") or "").strip() or "-",
                        "source": source_label,
                        "start_dt": start_dt,
                        "end_dt": end_dt,
                        "color": str(row.get("color", "") or "").strip() or "#dbeafe",
                    }
                )
        rows.sort(
            key=lambda row: (
                row.get("start_dt") or datetime.max,
                operation_index.get(str(row.get("operacao", "") or ""), 999),
                str(row.get("material", "") or ""),
                str(row.get("espessura", "") or ""),
                str(row.get("source", "") or ""),
            )
        )
        return rows

    def planning_render_order_detail_pdf(
        self,
        numero: str,
        output_path: str = "",
        *,
        focus_material: str = "",
        focus_espessura: str = "",
    ) -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        detail = self.order_detail(numero)
        flow_rows = self._planning_order_flow_rows(numero)
        if output_path:
            path = Path(output_path)
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            path = Path(tempfile.gettempdir()) / f"lugest_fluxo_encomenda_{str(numero or '').strip()}_{stamp}.pdf"
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

        focus_mat = str(focus_material or "").strip()
        focus_esp = str(focus_espessura or "").strip()
        focus_label = " | ".join(part for part in [focus_mat, (f"{focus_esp} mm" if focus_esp else "")] if part)
        active_count = sum(1 for row in flow_rows if str(row.get("source", "") or "") == "Ativo")
        history_count = sum(1 for row in flow_rows if str(row.get("source", "") or "") == "Histórico")
        total_minutes = sum(int(row.get("duracao_min", 0) or 0) for row in flow_rows)
        unique_ops = sorted({str(row.get("operacao", "") or "").strip() for row in flow_rows if str(row.get("operacao", "") or "").strip()})
        first_start = next((row.get("start_dt") for row in flow_rows if row.get("start_dt") is not None), None)
        last_end = None
        for row in flow_rows:
            end_dt = row.get("end_dt")
            if end_dt is not None and (last_end is None or end_dt > last_end):
                last_end = end_dt
        materials_rows = list(detail.get("materials", []) or [])

        def set_font(bold: bool, size: float) -> None:
            c.setFont(font_bold if bold else font_regular, size)

        def draw_header() -> float:
            c.setFillColor(palette["primary"])
            c.roundRect(margin, height - margin - 64, width - margin * 2, 64, 18, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 17)
            c.drawString(margin + 16, height - margin - 24, f"Fluxo de Planeamento - Encomenda {detail.get('numero', '-')}")
            set_font(False, 9)
            subtitle = "Resumo completo do planeamento da encomenda, com materiais, operações, recursos e blocos."
            if focus_label:
                subtitle += f" Foco: {focus_label}."
            c.drawString(margin + 16, height - margin - 42, subtitle)
            company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
            c.drawRightString(width - margin - 16, height - margin - 24, company)
            c.drawRightString(width - margin - 16, height - margin - 42, datetime.now().strftime("%d/%m/%Y %H:%M"))
            return height - margin - 82

        def ensure_page(y: float, needed: float) -> float:
            if y - needed >= 46:
                return y
            c.showPage()
            return draw_header()

        def draw_summary(y: float) -> float:
            summary_boxes = [
                ("Cliente", " - ".join([part for part in [detail.get("cliente", ""), detail.get("cliente_nome", "")] if part]).strip(" -") or "-", palette["primary"]),
                ("Entrega", str(detail.get("data_entrega", "") or "-"), palette["ink"]),
                ("Estado", str(detail.get("estado", "") or "-"), palette["success"] if flow_rows else palette["warning"]),
                ("Blocos", f"{active_count} ativos | {history_count} hist.", palette["warning"]),
            ]
            box_w = (width - margin * 2 - 18) / 4
            x = margin
            for label, value, tone in summary_boxes:
                c.setFillColor(colors.white)
                c.setStrokeColor(palette["line"])
                c.roundRect(x, y - 48, box_w, 44, 14, stroke=1, fill=1)
                c.setFillColor(tone)
                set_font(True, 8.5)
                c.drawString(x + 10, y - 17, label)
                c.setFillColor(palette["ink"])
                set_font(False, 9)
                c.drawString(x + 10, y - 33, _pdf_clip_text(value, box_w - 18, font_regular, 9))
                x += box_w + 6
            meta_parts = [
                f"Operações: {', '.join(unique_ops) if unique_ops else '-'}",
                f"Total planeado: {total_minutes:.0f} min",
                f"Início: {first_start.strftime('%d/%m/%Y %H:%M') if first_start is not None else '-'}",
                f"Fim: {last_end.strftime('%d/%m/%Y %H:%M') if last_end is not None else '-'}",
            ]
            c.setFillColor(palette["muted"])
            set_font(False, 8.5)
            c.drawString(margin + 2, y - 60, " | ".join(meta_parts))
            return y - 74

        def draw_table_header(y: float, columns: list[tuple[str, float]]) -> tuple[float, list[float], float]:
            total_w = width - margin * 2 - 18
            c.setFillColor(palette["primary_dark"])
            c.roundRect(margin, y - 22, width - margin * 2, 20, 10, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 8.2)
            x_positions: list[float] = []
            cursor = margin + 8
            for label, ratio in columns:
                x_positions.append(cursor)
                c.drawString(cursor + 3, y - 14, label)
                cursor += total_w * ratio
            return y - 26, x_positions, total_w

        y = draw_header()
        y = draw_summary(y)

        flow_columns = [
            ("Data", 0.12),
            ("Início", 0.08),
            ("Fim", 0.08),
            ("Operação", 0.14),
            ("Recurso", 0.16),
            ("Material", 0.16),
            ("Esp.", 0.07),
            ("Dur.", 0.07),
            ("Origem", 0.10),
        ]
        row_h = 18
        y = ensure_page(y, 52)
        set_font(True, 11)
        c.setFillColor(palette["ink"])
        c.drawString(margin, y - 4, "Blocos de planeamento")
        y -= 8
        y, x_positions, total_w = draw_table_header(y, flow_columns)
        if flow_rows:
            for row_index, row in enumerate(flow_rows):
                if y - (row_h + 12) < 46:
                    y = draw_header()
                    y, x_positions, total_w = draw_table_header(y - 8, flow_columns)
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                if focus_mat and focus_esp:
                    if str(row.get("material", "") or "").strip() == focus_mat and str(row.get("espessura", "") or "").strip() == focus_esp:
                        fill_color = colors.HexColor("#ECFDF3")
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, width - margin * 2, row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(row.get("data", "") or "-"),
                    str(row.get("inicio", "") or "-"),
                    str(row.get("fim", "") or "-"),
                    str(row.get("operacao", "") or "-"),
                    _pdf_clip_text(row.get("recurso", "-"), total_w * flow_columns[4][1] - 10, font_regular, 8),
                    _pdf_clip_text(row.get("material", "-"), total_w * flow_columns[5][1] - 10, font_regular, 8),
                    str(row.get("espessura", "") or "-"),
                    f"{float(row.get('duracao_min', 0) or 0):.0f}",
                    str(row.get("source", "") or "-"),
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    c.drawString(x_positions[idx] + 3, y - 10, value)
                y -= row_h
        else:
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(margin, y - 30, width - margin * 2, 24, 10, stroke=1, fill=1)
            c.setFillColor(palette["ink"])
            set_font(False, 9)
            c.drawString(margin + 10, y - 15, "Ainda não existem blocos de planeamento registados para esta encomenda.")
            y -= 34

        y -= 8
        material_columns = [
            ("Material", 0.17),
            ("Esp.", 0.08),
            ("Operações", 0.24),
            ("Recursos", 0.30),
            ("Tempos", 0.21),
        ]
        y = ensure_page(y, 60)
        set_font(True, 11)
        c.setFillColor(palette["ink"])
        c.drawString(margin, y - 4, "Materiais e operações")
        y -= 8
        y, x_positions, total_w = draw_table_header(y, material_columns)
        for row_index, row in enumerate(materials_rows):
            if y - (row_h + 12) < 46:
                y = draw_header()
                y, x_positions, total_w = draw_table_header(y - 8, material_columns)
            highlight = focus_mat and focus_esp and str(row.get("material", "") or "").strip() == focus_mat and str(row.get("espessura", "") or "").strip() == focus_esp
            fill_color = colors.HexColor("#FEF3C7") if highlight else (colors.white if row_index % 2 == 0 else palette["surface_alt"])
            c.setFillColor(fill_color)
            c.setStrokeColor(palette["line"])
            c.roundRect(margin, y - row_h + 2, width - margin * 2, row_h - 2, 8, stroke=1, fill=1)
            values = [
                _pdf_clip_text(row.get("material", "-"), total_w * material_columns[0][1] - 10, font_regular, 8),
                str(row.get("espessura", "") or "-"),
                _pdf_clip_text(" + ".join(list(row.get("operacoes_planeamento", []) or [])) or "-", total_w * material_columns[2][1] - 10, font_regular, 8),
                _pdf_clip_text(row.get("recursos_operacao_txt", "-"), total_w * material_columns[3][1] - 10, font_regular, 8),
                _pdf_clip_text(row.get("tempo_operacoes_txt", "-"), total_w * material_columns[4][1] - 10, font_regular, 8),
            ]
            c.setFillColor(palette["ink"])
            set_font(False, 8)
            for idx, value in enumerate(values):
                c.drawString(x_positions[idx] + 3, y - 10, value)
            y -= row_h

        c.setFillColor(palette["muted"])
        set_font(False, 8)
        footer_txt = (
            f"Posto principal: {detail.get('posto_trabalho', '-') or '-'} | "
            f"Observações: {detail.get('observacoes', '-') or '-'}"
        )
        c.drawString(margin, 28, _pdf_clip_text(footer_txt, width - (margin * 2) - 110, font_regular, 8))
        c.drawRightString(width - margin, 28, f"LUGEST | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.save()
        return path

    def planning_open_order_detail_pdf(
        self,
        numero: str,
        *,
        focus_material: str = "",
        focus_espessura: str = "",
    ) -> Path:
        path = self.planning_render_order_detail_pdf(
            numero,
            focus_material=focus_material,
            focus_espessura=focus_espessura,
        )
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
            resource=self._planning_row_resource(target),
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
            resource=self._planning_row_resource(target),
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
        target = next(
            (row for row in list(self.ensure_data().get("plano", []) or []) if str(row.get("id", "") or "").strip() == block_txt),
            None,
        )
        if target is None:
            raise ValueError("Bloco de planeamento n?o encontrado.")
        numero = str(target.get("encomenda", "") or "").strip()
        material = str(target.get("material", "") or "").strip()
        espessura = str(target.get("espessura", "") or "").strip()
        current_operation = self._planning_row_operation(target)
        cascade_ops = self._planning_item_operation_sequence(numero, material, espessura, start_operation=current_operation)
        if current_operation not in cascade_ops:
            cascade_ops = [current_operation, *list(cascade_ops or [])]
        cascade_ops = [self._planning_normalize_operation(op_name) for op_name in cascade_ops if str(op_name or "").strip()]
        cascade_set = {op_name for op_name in cascade_ops if op_name}
        target_item_key = self._planning_item_key(numero, material, espessura)
        self.ensure_data()["plano"] = [
            row
            for row in list(self.ensure_data().get("plano", []) or [])
            if (
                self._planning_item_key(row.get("encomenda", ""), row.get("material", ""), row.get("espessura", "")) != target_item_key
                or self._planning_row_operation(row) not in cascade_set
            )
        ]
        self._save(force=True)

    def planning_open_pdf(self, week_start: str | date | None = None, operation: Any = "Corte Laser", resource: str = "") -> Path:
        week_start_dt = self._planning_week_start(week_start)
        op_txt = self._planning_normalize_operation(operation)
        resource_txt = self._normalize_workcenter_value(resource)
        filtered_data = dict(self.ensure_data())
        filtered_data["plano"] = [
            dict(row)
            for row in list(self.ensure_data().get("plano", []) or [])
            if isinstance(row, dict)
            and self._planning_row_matches_operation(row, op_txt)
            and (not resource_txt or self._planning_row_resource(row).lower() == resource_txt.lower())
        ]
        ctx = SimpleNamespace(
            data=filtered_data,
            p_week_start=week_start_dt,
            p_inicio=_ValueHolder("08:00"),
            p_fim=_ValueHolder("18:00"),
            planning_operation_label=op_txt,
            planning_resource_label=resource_txt,
        )
        ctx.get_plano_grid_metrics = lambda: (480, 1080, 30, 20, 6, 0, 0)
        ctx.plano_intervalo_bloqueado = lambda start_min, end_min: self._planning_interval_blocked(start_min, end_min)
        ctx.get_chapa_reservada = lambda numero, material=None, espessura=None: self._order_reserved_sheet(numero, material or "", espessura or "")
        path = Path(tempfile.gettempdir()) / "lugest_plano.pdf"
        self.plan_actions.preview_plano_a4(ctx)
        return path

