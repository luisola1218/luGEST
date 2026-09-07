from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any


class PlanningOperationsBackendMixin:
    """Legacy adapter for planning operations; see BACKEND_GUIDE.md."""

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
