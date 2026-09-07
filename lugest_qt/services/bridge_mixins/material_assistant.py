from __future__ import annotations

import os
import tempfile
from datetime import date, datetime, timedelta
from lugest_infra.pdf.dossier_reports import render_material_separation as _render_dossier_material_separation
from pathlib import Path
from typing import Any


class MaterialAssistantBackendMixin:
    """Legacy adapter for material assistant; see BACKEND_GUIDE.md."""

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
                    assigned_candidate = standard_pool[index - 1] if index <= len(standard_pool) else {}
                    if assigned_candidate:
                        need["preferred_lot"] = str(assigned_candidate.get("lote", "") or "").strip()
                        need["preferred_material_id"] = str(assigned_candidate.get("material_id", "") or "").strip()
                        need["preferred_dimensao"] = str(assigned_candidate.get("dimensao", "") or "").strip()
                        need["preferred_disponivel"] = round(self._parse_float(assigned_candidate.get("disponivel", 0), 0), 2)
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

            if material_cativado and next_action_hours is not None and next_action_hours <= 24:
                keep_ready_qty = self._parse_float(need.get("reserved_qty", 0), 0)
                if keep_ready_qty <= 0:
                    keep_ready_qty = self._parse_float(need.get("quantidade_preparar", need.get("piece_qty", 0)), 0)
                _append_suggestion(
                    "keep_ready",
                    "Evitar arrumação desnecessária",
                    (
                        f"Material já cativado para {numero}; vai ser necessário"
                        f" em {need.get('next_action_label', '-')}. Mantém disponível."
                    ),
                    [
                        f"Quantidade cativada: {self._fmt(keep_ready_qty)}",
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
        return _render_dossier_material_separation(self, path, rows, int(horizon_days or 5))

    def material_assistant_open_separation_pdf(self, horizon_days: int = 5) -> Path:
        path = self.material_assistant_render_separation_pdf(horizon_days=horizon_days)
        os.startfile(str(path))
        return path

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
