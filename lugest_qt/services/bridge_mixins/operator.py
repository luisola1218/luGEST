from __future__ import annotations

import os
import time
from datetime import datetime
from lugest_qt.services.bridge_helpers import _ValueHolder
from types import SimpleNamespace
from typing import Any


class OperatorBackendMixin:
    """Legacy adapter for operator; see BACKEND_GUIDE.md."""

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
        self._save(force=True, blocking=True)

    def _find_piece(self, enc_num: str, piece_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        for piece in self.desktop_main.encomenda_pecas(enc):
            if str(piece.get("id", "")).strip() == str(piece_id or "").strip():
                return enc, piece
        raise ValueError("Peça não encontrada.")

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

    def _piece_operation_limit(self, piece: dict[str, Any], operation: str, enc_num: str = "") -> float:
        target = self.desktop_main.normalize_operacao_nome(operation or "")
        planned_qty = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        if not target:
            return planned_qty
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            op_name = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
            if op_name == target:
                return round(planned_qty, 4)
        return planned_qty

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

    def _piece_operation_history_quantities(self, piece: dict[str, Any], operation: str) -> dict[str, float]:
        op_norm = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        if not op_norm:
            return {"ok": 0.0, "nok": 0.0, "qual": 0.0, "total": 0.0}
        ok_total = 0.0
        nok_total = 0.0
        qual_total = 0.0
        for row in list((piece or {}).get("hist", []) or []):
            if not isinstance(row, dict):
                continue
            action_norm = self.desktop_main.norm_text(row.get("acao", ""))
            if "fim" not in action_norm and "registo" not in action_norm:
                continue
            row_ops = [
                self.desktop_main.normalize_operacao_nome(op) or str(op or "").strip()
                for op in list(row.get("operacoes", []) or [])
                if str(op or "").strip()
            ]
            if row_ops and op_norm not in row_ops:
                continue
            if not row_ops:
                single_op = self.desktop_main.normalize_operacao_nome(row.get("operacao", "")) or str(row.get("operacao", "") or "").strip()
                if single_op != op_norm:
                    continue
            ok_total += self._parse_float(row.get("ok", row.get("qtd_ok", 0)), 0)
            nok_total += self._parse_float(row.get("nok", row.get("qtd_nok", 0)), 0)
            qual_total += self._parse_float(row.get("qual", row.get("qtd_qual", 0)), 0)
        total = round(ok_total + nok_total + qual_total, 4)
        return {
            "ok": round(ok_total, 4),
            "nok": round(nok_total, 4),
            "qual": round(qual_total, 4),
            "total": total,
        }

    def _piece_operation_event_quantities(self, enc_num: str, piece_id: str, operation: str, piece: dict[str, Any] | None = None) -> dict[str, float]:
        enc_txt = str(enc_num or "").strip()
        piece_txt = str(piece_id or "").strip()
        op_norm = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        if not enc_txt or not piece_txt or not op_norm:
            return {"ok": 0.0, "nok": 0.0, "qual": 0.0, "total": 0.0}
        ok_total = 0.0
        nok_total = 0.0
        qual_total = 0.0
        for row in list(self.ensure_data().get("op_eventos", []) or []):
            if not isinstance(row, dict):
                continue
            event_norm = self.desktop_main.norm_text(row.get("evento", ""))
            if event_norm != "finish_op":
                continue
            if str(row.get("encomenda_numero", "") or row.get("encomenda", "") or "").strip() != enc_txt:
                continue
            if str(row.get("peca_id", "") or "").strip() != piece_txt:
                continue
            row_op = self.desktop_main.normalize_operacao_nome(row.get("operacao", "")) or str(row.get("operacao", "") or "").strip()
            if row_op != op_norm:
                continue
            ok_total += self._parse_float(row.get("qtd_ok", row.get("ok", 0)), 0)
            nok_total += self._parse_float(row.get("qtd_nok", row.get("nok", 0)), 0)
            qual_total += self._parse_float(row.get("qtd_qual", row.get("qual", 0)), 0)
        total = round(ok_total + nok_total + qual_total, 4)
        if total <= 0 and isinstance(piece, dict):
            return self._piece_operation_history_quantities(piece, op_norm)
        return {
            "ok": round(ok_total, 4),
            "nok": round(nok_total, 4),
            "qual": round(qual_total, 4),
            "total": total,
        }

    def _piece_operation_mysql_quantities(self, enc_num: str, piece_id: str, operation: str) -> dict[str, float]:
        op_norm = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        if not op_norm:
            return {"ok": 0.0, "nok": 0.0, "qual": 0.0, "total": 0.0}
        rows = self._operator_mysql_ops_status_rows(enc_num, piece_id)
        for row in rows:
            if not isinstance(row, dict):
                continue
            row_op = self.desktop_main.normalize_operacao_nome(row.get("operacao", "")) or str(row.get("operacao", "") or "").strip()
            if row_op != op_norm:
                continue
            ok = self._parse_float(row.get("ok_qty", row.get("qtd_ok", 0)), 0)
            nok = self._parse_float(row.get("nok_qty", row.get("qtd_nok", 0)), 0)
            qual = self._parse_float(row.get("qual_qty", row.get("qtd_qual", 0)), 0)
            return {"ok": round(ok, 4), "nok": round(nok, 4), "qual": round(qual, 4), "total": round(ok + nok + qual, 4)}
        return {"ok": 0.0, "nok": 0.0, "qual": 0.0, "total": 0.0}

    def _operator_mysql_ops_status_rows(self, enc_num: str, piece_id: str) -> list[dict[str, Any]]:
        enc_txt = str(enc_num or "").strip()
        piece_txt = str(piece_id or "").strip()
        if not enc_txt or not piece_txt:
            return []
        key = (enc_txt, piece_txt)
        ttl = float(getattr(self, "_op_mysql_ops_status_ttl_sec", 2.0) or 0.0)
        cache = getattr(self, "_op_mysql_ops_status_cache", None)
        if isinstance(cache, dict) and ttl > 0:
            cached = cache.get(key)
            if cached is not None:
                loaded_at, rows = cached
                if (time.time() - float(loaded_at or 0.0)) <= ttl:
                    return [dict(row or {}) for row in list(rows or [])]
        status_fn = getattr(self.operador_actions, "_mysql_ops_status_for_piece", None)
        if not callable(status_fn):
            return []
        try:
            rows = [dict(row or {}) for row in list(status_fn(enc_txt, piece_txt) or []) if isinstance(row, dict)]
        except Exception:
            rows = []
        if isinstance(cache, dict) and ttl > 0:
            cache[key] = (time.time(), [dict(row or {}) for row in rows])
        return rows

    def _operator_invalidate_ops_status_cache(self, enc_num: str = "", piece_id: str = "") -> None:
        cache = getattr(self, "_op_mysql_ops_status_cache", None)
        if not isinstance(cache, dict) or not cache:
            return
        enc_txt = str(enc_num or "").strip()
        piece_txt = str(piece_id or "").strip()
        if enc_txt and piece_txt:
            cache.pop((enc_txt, piece_txt), None)
            return
        if enc_txt:
            for key in list(cache.keys()):
                if key and key[0] == enc_txt:
                    cache.pop(key, None)
            return
        cache.clear()

    def _piece_operation_recorded_total(self, enc_num: str, piece: dict[str, Any], operation: str, limit: float = 0.0) -> float:
        op_row = self._piece_operation_row(piece, operation)
        row_total = self._piece_operation_total(op_row, limit)
        events = self._piece_operation_event_quantities(enc_num, str(piece.get("id", "") or ""), operation, piece)
        event_total = self._parse_float(events.get("total", 0), 0)
        mysql_quantities = self._piece_operation_mysql_quantities(enc_num, str(piece.get("id", "") or ""), operation)
        mysql_total = self._parse_float(mysql_quantities.get("total", 0), 0)
        if mysql_total > 0 and event_total <= 0:
            event_total = mysql_total
            events = mysql_quantities
        if event_total <= 0:
            return row_total
        row_qual = self._parse_float((op_row or {}).get("qtd_qual", 0), 0)
        event_with_row_qual = round(event_total + row_qual, 4)
        if row_total > event_with_row_qual + 1e-9:
            return event_with_row_qual
        return max(row_total, event_with_row_qual)

    def _repair_operation_quantities_from_events(self, enc_num: str, piece: dict[str, Any], operation: str) -> None:
        op_row = self._piece_operation_row(piece, operation)
        if not isinstance(op_row, dict):
            return
        events = self._piece_operation_event_quantities(enc_num, str(piece.get("id", "") or ""), operation, piece)
        event_total = self._parse_float(events.get("total", 0), 0)
        mysql_quantities = self._piece_operation_mysql_quantities(enc_num, str(piece.get("id", "") or ""), operation)
        mysql_total = self._parse_float(mysql_quantities.get("total", 0), 0)
        if mysql_total > 0 and event_total <= 0:
            event_total = mysql_total
            events = mysql_quantities
        if event_total <= 0:
            return
        row_qual = self._parse_float(op_row.get("qtd_qual", 0), 0)
        row_total = (
            self._parse_float(op_row.get("qtd_ok", 0), 0)
            + self._parse_float(op_row.get("qtd_nok", 0), 0)
            + row_qual
        )
        expected_total = round(event_total + row_qual, 4)
        if row_total <= expected_total + 1e-9:
            return
        op_row["qtd_ok"] = self._parse_float(events.get("ok", 0), 0)
        op_row["qtd_nok"] = self._parse_float(events.get("nok", 0), 0)
        op_row["qtd_qual"] = row_qual

    def _sync_piece_output_from_flow(self, piece: dict[str, Any]) -> None:
        fluxo = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
        produced_totals: list[float] = []
        final_ok_candidates: list[float] = []
        final_nok_candidates: list[float] = []
        final_qual_candidates: list[float] = []
        for op in fluxo:
            total = self._piece_operation_total(op, self._piece_operation_limit(piece, op.get("nome", "")))
            if not str(op.get("nome", "") or "").strip():
                continue
            produced_totals.append(total)
            final_ok_candidates.append(self._parse_float(op.get("qtd_ok", 0), 0))
            final_nok_candidates.append(self._parse_float(op.get("qtd_nok", 0), 0))
            final_qual_candidates.append(self._parse_float(op.get("qtd_qual", 0), 0))
        if not produced_totals:
            piece["produzido_ok"] = 0.0
            piece["produzido_nok"] = 0.0
            piece["produzido_qualidade"] = 0.0
            return
        # Compatibilidade: o produzido global deixa de ser input de validação e passa a ser derivado.
        piece["produzido_ok"] = round(min(final_ok_candidates) if final_ok_candidates else min(produced_totals), 4)
        piece["produzido_nok"] = round(max(final_nok_candidates) if final_nok_candidates else 0.0, 4)
        piece["produzido_qualidade"] = round(max(final_qual_candidates) if final_qual_candidates else 0.0, 4)

    def operator_laser_stock_state(self, enc_num: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura não encontrado.")
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
                    "quantidade": round(qty_res, 4),
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

    def operator_material_session_state(
        self,
        enc_num: str,
        material: str,
        espessura: str,
        operator_name: str,
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura não encontrado.")
        operator_norm = self.desktop_main.norm_text(operator_name)
        status_for_order = getattr(self.operador_actions, "_mysql_ops_status_for_order", None)
        mysql_rows: dict[str, list[dict[str, Any]]] = {}
        if callable(status_for_order):
            try:
                mysql_rows = dict(status_for_order(enc_num, cache_owner=self, force=True) or {})
            except Exception:
                mysql_rows = {}

        active_piece_ids: list[str] = []
        for piece in list(esp_obj.get("pecas", []) or []):
            piece_id = str(piece.get("id", "") or "").strip()
            active = False
            rows = list(mysql_rows.get(piece_id, []) or [])
            if rows:
                active = any(
                    "produc" in self.desktop_main.norm_text(row.get("estado", ""))
                    and self.desktop_main.norm_text(row.get("operador", "")) == operator_norm
                    and self._is_laser_operation(str(row.get("operacao", "") or ""))
                    for row in rows
                )
            else:
                for operation in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
                    state_norm = self.desktop_main.norm_text(operation.get("estado", ""))
                    owner_norm = self.desktop_main.norm_text(operation.get("user", piece.get("operador", "")))
                    if (
                        ("produc" in state_norm or "curso" in state_norm)
                        and owner_norm == operator_norm
                        and self._is_laser_operation(str(operation.get("nome", "") or ""))
                    ):
                        active = True
                        break
            if active and piece_id:
                active_piece_ids.append(piece_id)

        state = self.operator_laser_stock_state(enc_num, material, espessura)
        state.update(
            {
                "operator": str(operator_name or "").strip(),
                "active_piece_ids": active_piece_ids,
                "active_count": len(active_piece_ids),
                "should_prompt": len(active_piece_ids) == 0 and float(state.get("reserved_qty", 0) or 0) > 0,
            }
        )
        return state

    def operator_record_material_session_decision(
        self,
        enc_num: str,
        material: str,
        espessura: str,
        operator_name: str,
        decision: str,
        reason: str = "",
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura não encontrado.")
        row = {
            "ts": self.desktop_main.now_iso(),
            "operador": str(operator_name or "").strip(),
            "decisao": str(decision or "").strip(),
            "motivo": str(reason or "").strip(),
            "material": str(material or "").strip(),
            "espessura": str(espessura or "").strip(),
        }
        esp_obj.setdefault("material_session_history", []).append(row)
        self._save_operator_state(enc)
        return row

    def operator_reserved_materials(self, enc_num: str, material: str = "", espessura: str = "") -> list[dict[str, Any]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        material_norm = self._norm_material_token(material)
        esp_norm = self._norm_esp_token(espessura)
        rows: list[dict[str, Any]] = []
        for index, reserva in enumerate(list(enc.get("reservas", []) or [])):
            if not isinstance(reserva, dict):
                continue
            if material_norm and self._norm_material_token(reserva.get("material")) != material_norm:
                continue
            if esp_norm and self._norm_esp_token(reserva.get("espessura")) != esp_norm:
                continue
            qty = self._parse_float(reserva.get("quantidade", 0), 0)
            if qty <= 0:
                continue
            material_id = str(reserva.get("material_id", "") or "").strip()
            stock = self.material_by_id(material_id) if material_id else None
            mat_txt = str(reserva.get("material", "") or (stock or {}).get("material", "") or "").strip()
            esp_txt = str(reserva.get("espessura", "") or (stock or {}).get("espessura", "") or "").strip()
            rows.append(
                {
                    "index": index,
                    "material_id": material_id,
                    "material": mat_txt,
                    "espessura": esp_txt,
                    "quantidade": round(qty, 4),
                    "lote": str((stock or {}).get("lote_interno", "") or (stock or {}).get("lote_fornecedor", "") or reserva.get("lote", "") or "").strip(),
                    "lote_fornecedor": str((stock or {}).get("lote_fornecedor", "") or reserva.get("lote", "") or "").strip(),
                    "dimensao": f"{(stock or {}).get('comprimento', '-') } x {(stock or {}).get('largura', '-')}",
                    "local": self._localizacao(stock) if stock else "-",
                    "stock_quantidade": round(self._parse_float((stock or {}).get("quantidade", 0), 0), 4) if stock else 0.0,
                    "stock_reservado": round(self._parse_float((stock or {}).get("reservado", 0), 0), 4) if stock else 0.0,
                }
            )
        return rows

    def _create_retalho_from_stock(
        self,
        source_stock: dict[str, Any],
        retalho_payload: dict[str, Any],
        *,
        enc_num: str = "",
        reason: str = "",
        lote_sel: str = "",
    ) -> dict[str, Any] | None:
        retalho_payload = dict(retalho_payload or {})
        has_retalho = any(str(retalho_payload.get(key, "")).strip() for key in ("comprimento", "largura", "quantidade", "metros"))
        if not has_retalho:
            return None
        comp = self._parse_float(retalho_payload.get("comprimento", 0), 0)
        larg = self._parse_float(retalho_payload.get("largura", 0), 0)
        q_retalho = self._parse_float(retalho_payload.get("quantidade", 0), 0)
        metros = self._parse_float(retalho_payload.get("metros", 0), 0)
        if q_retalho <= 0:
            raise ValueError("Quantidade do retalho invalida.")
        created_retalho = {
            "id": self._next_material_id(),
            "lote_interno": self._next_material_internal_lot(),
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
            "origem_lote_interno": str(source_stock.get("lote_interno", "") or "").strip(),
            "origem_lote": str(source_stock.get("lote_fornecedor", "") or "").strip(),
            "origem_lotes_baixa": [
                lot
                for lot in list(
                    dict.fromkeys(
                        [
                            str(lote_sel or "").strip(),
                            str(source_stock.get("lote_interno", "") or "").strip(),
                            str(source_stock.get("lote_fornecedor", "") or "").strip(),
                        ]
                    )
                )
                if lot
            ],
            "atualizado_em": self.desktop_main.now_iso(),
        }
        self.materia_actions._hydrate_retalho_record(self.ensure_data(), created_retalho, template=source_stock)
        self.ensure_data().setdefault("materiais", []).append(created_retalho)
        self.desktop_main.log_stock(
            self.ensure_data(),
            "RETALHO",
            f"{source_stock.get('id', '')} qtd={created_retalho.get('quantidade', 0)} encomenda={enc_num} motivo={reason or 'retalho'}",
            operador=self._current_user_label(),
        )
        return created_retalho

    def operator_partial_reserved_material_consumption(
        self,
        enc_num: str,
        material_id: str,
        quantidade: Any,
        material: str = "",
        espessura: str = "",
        retalho: dict[str, Any] | None = None,
        source_material_id: str = "",
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        target_material_id = str(material_id or "").strip()
        qty = self._parse_float(quantidade, 0)
        if not target_material_id:
            raise ValueError("Seleciona o material cativado.")
        if qty <= 0:
            raise ValueError("Quantidade de baixa invalida.")
        reservas = list(enc.get("reservas", []) or [])
        target_index = None
        target_reserva: dict[str, Any] | None = None
        for index, reserva in enumerate(reservas):
            if not isinstance(reserva, dict):
                continue
            if str(reserva.get("material_id", "") or "").strip() != target_material_id:
                continue
            if material and self._norm_material_token(reserva.get("material")) != self._norm_material_token(material):
                continue
            if espessura and self._norm_esp_token(reserva.get("espessura")) != self._norm_esp_token(espessura):
                continue
            target_index = index
            target_reserva = reserva
            break
        if target_reserva is None or target_index is None:
            raise ValueError("Reserva de material cativado nao encontrada.")
        reserved_qty = self._parse_float(target_reserva.get("quantidade", 0), 0)
        if qty > reserved_qty + 1e-9:
            raise ValueError(f"Quantidade superior ao material cativado. Maximo: {reserved_qty:.2f}")
        stock = self.material_by_id(target_material_id)
        if stock is None:
            raise ValueError("Material cativado nao encontrado no stock.")
        if self._material_quality_is_blocked(stock):
            raise ValueError(f"Material reservado {target_material_id} bloqueado pela qualidade.")
        stock_qty = self._parse_float(stock.get("quantidade", 0), 0)
        stock_reserved = self._parse_float(stock.get("reservado", 0), 0)
        if qty > stock_qty + 1e-9:
            raise ValueError(f"Quantidade superior ao stock do material. Maximo: {stock_qty:.2f}")
        stock["quantidade"] = max(0.0, stock_qty - qty)
        stock["reservado"] = max(0.0, stock_reserved - qty)
        stock["atualizado_em"] = self.desktop_main.now_iso()
        remaining_reserva = round(max(0.0, reserved_qty - qty), 4)
        if remaining_reserva <= 1e-9:
            reservas.pop(target_index)
        else:
            target_reserva["quantidade"] = remaining_reserva
            reservas[target_index] = target_reserva
        enc["reservas"] = reservas
        enc["cativar"] = bool(reservas)
        lote_sel = str(stock.get("lote_interno", "") or stock.get("lote_fornecedor", "") or target_reserva.get("lote", "") or "").strip()
        self.desktop_main.log_stock(
            self.ensure_data(),
            "BAIXA PARCIAL CATIVADA",
            f"{target_material_id} qtd={qty} encomenda={enc.get('numero', '')}",
            operador=self._current_user_label(),
        )
        created_retalho = None
        source_id = str(source_material_id or target_material_id or "").strip()
        if retalho:
            source_stock = self.material_by_id(source_id)
            if source_stock is None:
                raise ValueError("Lote de origem do retalho nao encontrado.")
            created_retalho = self._create_retalho_from_stock(
                source_stock,
                dict(retalho or {}),
                enc_num=str(enc.get("numero", "") or "").strip(),
                reason="baixa_parcial_laser_qt",
                lote_sel=lote_sel,
            )
        self._sync_ne_from_materia()
        self._save_operator_state(enc)
        return {
            "resolved": True,
            "consumed_total": round(qty, 4),
            "remaining_reserved": remaining_reserva,
            "retalho_id": str((created_retalho or {}).get("id", "") or "").strip(),
            "material_id": target_material_id,
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
        session_close: bool = False,
        operator_name: str = "",
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp_obj = self._operator_esp_obj(enc, material, espessura)
        if not esp_obj:
            raise ValueError("Grupo material/espessura não encontrado.")
        laser_complete = self._operator_esp_laser_concluido(esp_obj)
        if not laser_complete and not session_close:
            return {"resolved": False, "reason": "laser_not_complete"}
        if laser_complete and self._operator_esp_laser_resolved(esp_obj):
            state = self.operator_laser_stock_state(enc_num, material, espessura)
            state["resolved"] = True
            return state
        total_qty = self._operator_group_total_output(esp_obj)
        state_before = self.operator_laser_stock_state(enc_num, material, espessura)
        if self._parse_float(state_before.get("reserved_qty", 0), 0) > 1e-9:
            stock_id = str(material_id or "").strip()
            consume_qty = self._parse_float(quantidade, 0)
            if not stock_id or consume_qty <= 0:
                raise ValueError("Seleciona o lote cativado e a quantidade realmente consumida.")
            partial = self.operator_partial_reserved_material_consumption(
                enc_num,
                stock_id,
                consume_qty,
                material=material,
                espessura=espessura,
                retalho=retalho,
                source_material_id=source_material_id,
            )
            stock = self.material_by_id(stock_id) or {}
            lote_sel = str(
                stock.get("lote_interno", "") or stock.get("lote_fornecedor", "") or ""
            ).strip()
            if lote_sel:
                esp_obj["lote_baixa"] = lote_sel
                for piece in list(esp_obj.get("pecas", []) or []):
                    piece["lote_baixa"] = lote_sel
            if laser_complete and not esp_obj.get("laser_concluido"):
                esp_obj["laser_concluido"] = True
                esp_obj["laser_concluido_em"] = self.desktop_main.now_iso()
            if laser_complete:
                esp_obj["baixa_laser_feita"] = True
                esp_obj["baixa_laser_confirmada_sem_baixa"] = False
                esp_obj["baixa_laser_em"] = self.desktop_main.now_iso()
            if session_close:
                esp_obj.setdefault("material_session_history", []).append(
                    {
                        "ts": self.desktop_main.now_iso(),
                        "operador": str(operator_name or "").strip(),
                        "decisao": "dar_baixa",
                        "material": str(material or "").strip(),
                        "espessura": str(espessura or "").strip(),
                        "quantidade": round(consume_qty, 4),
                        "material_id": stock_id,
                        "retalho_id": str(partial.get("retalho_id", "") or "").strip(),
                    }
                )
            state_after = self.operator_laser_stock_state(enc_num, material, espessura)
            self._save_operator_state(enc)
            return {
                "resolved": True,
                "total_qty": total_qty,
                "reserved_consumed": round(consume_qty, 4),
                "extra_consumed": 0.0,
                "consumed_total": round(consume_qty, 4),
                "remaining_qty": 0.0,
                "remaining_reserved": float(state_after.get("reserved_qty", 0) or 0),
                "allow_without_stock": False,
                "manual_stock_required": False,
                "lote_baixa": lote_sel,
                "retalho_id": str(partial.get("retalho_id", "") or "").strip(),
                "material_id": stock_id,
            }
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
            raise ValueError("Com material cativado não é permitida baixa manual adicional. Apenas podes registar retalho.")
        if stock_id and extra_qty > 0:
            stock = self.material_by_id(stock_id)
            if stock is None:
                raise ValueError("Material não encontrado para baixa.")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {stock_id} bloqueado pela qualidade.")
            if self._norm_material_token(stock.get("material")) != self._norm_material_token(material) or self._norm_esp_token(stock.get("espessura")) != self._norm_esp_token(espessura):
                raise ValueError("O stock selecionado não corresponde ao material/espessura.")
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
                raise ValueError("Lote de origem do retalho não encontrado.")
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
        if laser_complete and not esp_obj.get("laser_concluido"):
            esp_obj["laser_concluido"] = True
            esp_obj["laser_concluido_em"] = self.desktop_main.now_iso()
        if laser_complete:
            esp_obj["baixa_laser_feita"] = bool((reserved_consumed > 0) or (extra_consumed > 0))
            esp_obj["baixa_laser_confirmada_sem_baixa"] = bool(allow_without_stock and manual_stock_required and extra_consumed <= 0)
            esp_obj["baixa_laser_em"] = self.desktop_main.now_iso()
        if session_close:
            esp_obj.setdefault("material_session_history", []).append(
                {
                    "ts": self.desktop_main.now_iso(),
                    "operador": str(operator_name or "").strip(),
                    "decisao": "dar_baixa",
                    "material": str(material or "").strip(),
                    "espessura": str(espessura or "").strip(),
                    "quantidade": consumed_total,
                    "retalho_id": str((created_retalho or {}).get("id", "") or "").strip(),
                }
            )
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
        pending_ops: list[str] = []
        done_ops: list[str] = []
        op_limits: dict[str, float] = {}
        op_totals: dict[str, float] = {}
        for op_row in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            op_name = self.desktop_main.normalize_operacao_nome(op_row.get("nome", "")) or str(op_row.get("nome", "") or "").strip()
            if not op_name:
                continue
            limit = self._piece_operation_limit(piece, op_name, str(enc.get("numero", "") or ""))
            total_done = self._piece_operation_recorded_total(str(enc.get("numero", "") or ""), piece, op_name, limit)
            op_limits[op_name] = limit
            op_totals[op_name] = total_done
            if limit > 0 and total_done >= limit - 1e-9:
                done_ops.append(op_name)
            else:
                pending_ops.append(op_name)
        live_status_map: dict[str, dict[str, Any]] = {}
        for row in self._operator_mysql_ops_status_rows(str(enc.get("numero", "") or ""), str(piece.get("id", "") or "")):
            op_name = self.desktop_main.normalize_operacao_nome((row or {}).get("operacao", "")) or str((row or {}).get("operacao", "") or "").strip()
            if op_name:
                live_status_map[op_name] = dict(row or {})
        active_pending_ops = []
        for op_name in pending_ops:
            live_state = self.desktop_main.norm_text((live_status_map.get(op_name, {}) or {}).get("estado", ""))
            if "produ" in live_state:
                active_pending_ops.append(op_name)
        current_op = active_pending_ops[0] if active_pending_ops else (self.desktop_main.normalize_operacao_nome(piece.get("operacao_atual", "")) or (pending_ops[0] if pending_ops else ""))
        current_limit = op_limits.get(current_op, self._piece_operation_limit(piece, current_op, str(enc.get("numero", "") or "")))
        current_done = op_totals.get(current_op, self._piece_operation_recorded_total(str(enc.get("numero", "") or ""), piece, current_op, current_limit))
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
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de iniciar a peça.")
        pending = list(ctx["pending_ops"])
        if not pending:
            raise ValueError("Esta peça não tem operações pendentes.")
        selected_op = self.desktop_main.normalize_operacao_nome(operation or pending[0])
        if selected_op not in pending:
            raise ValueError("A operação selecionada não está pendente.")
        available_qty = self._piece_operation_limit(piece, selected_op, str(enc.get("numero", "") or ""))
        current_qty = self._piece_operation_recorded_total(str(enc.get("numero", "") or ""), piece, selected_op, available_qty)
        if available_qty <= current_qty:
            raise ValueError("Esta operação não tem quantidade disponível do posto anterior.")
        result = self.operador_actions._mysql_ops_acquire(
            str(enc.get("numero", "") or ""),
            str(piece.get("id", "") or ""),
            [selected_op],
            operator_name,
            valid_operators=self.operator_names(),
        )
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        acquired = list(result.get("acquired", []) or [])
        blocked = list(result.get("blocked", []) or [])
        if not acquired:
            if blocked:
                owner = str((blocked[0] or {}).get("operador", "") or "").strip() or "outro operador"
                raise ValueError(f"Operacao ocupada por {owner}.")
            raise ValueError("Não foi possível iniciar a operação.")
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
            raise ValueError("Existe uma avaria aberta. Fecha a avaria antes de concluir a peça.")
        ok_val = self._parse_float(ok, 0)
        nok_val = self._parse_float(nok, 0)
        qual_val = self._parse_float(qual, 0)
        if min(ok_val, nok_val, qual_val) < 0:
            raise ValueError("Valores inválidos.")
        pending = list(ctx["pending_ops"])
        if not pending:
            raise ValueError("Não existem operações pendentes nesta peça.")
        selected_op = self.desktop_main.normalize_operacao_nome(operation or pending[0])
        if selected_op not in pending:
            raise ValueError("A operação selecionada não está pendente.")
        active_pending_ops = list(ctx.get("active_pending_ops", []) or [])
        if selected_op not in active_pending_ops:
            raise ValueError("Inicia primeiro a operação antes de a concluir.")
        op_row = self._piece_operation_row(piece, selected_op)
        self._repair_operation_quantities_from_events(str(enc.get("numero", "") or ""), piece, selected_op)
        op_row = self._piece_operation_row(piece, selected_op)
        operation_limit = self._piece_operation_limit(piece, selected_op, str(enc.get("numero", "") or ""))
        if operation_limit <= 0:
            raise ValueError("Não existe quantidade produzida no posto anterior para esta operação.")
        ctx_done = self._parse_float(dict(ctx.get("operation_done", {}) or {}).get(selected_op, 0), 0)
        current_done = max(
            self._piece_operation_recorded_total(str(enc.get("numero", "") or ""), piece, selected_op, operation_limit),
            ctx_done,
        )
        remaining_limit = round(max(0.0, operation_limit - current_done), 4)
        delta_val = round(ok_val + nok_val + qual_val, 4)
        if delta_val <= 0:
            raise ValueError("Indica pelo menos uma quantidade para concluir.")
        if delta_val > remaining_limit + 1e-9:
            raise ValueError(f"Quantidade acima do disponível nesta operação. Máximo restante: {remaining_limit:.1f}")
        existing_ok = self._parse_float((op_row or {}).get("qtd_ok", 0), 0)
        existing_nok = self._parse_float((op_row or {}).get("qtd_nok", 0), 0)
        existing_qual = self._parse_float((op_row or {}).get("qtd_qual", 0), 0)
        existing_total = round(existing_ok + existing_nok + existing_qual, 4)
        event_qty = self._piece_operation_event_quantities(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""), selected_op, piece)
        event_total = self._parse_float(event_qty.get("total", 0), 0)
        if event_total > 0 and existing_total > event_total + existing_qual + 1e-9:
            existing_ok = self._parse_float(event_qty.get("ok", 0), 0)
            existing_nok = self._parse_float(event_qty.get("nok", 0), 0)
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
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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
                    self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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
        motivo = str(motivo or "").strip() or "Avaria não especificada"
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
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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
        motivo = str(piece.get("avaria_motivo", "") or piece.get("interrupcao_peca_motivo", "") or (live_row or {}).get("causa", "") or "").strip() or "Avaria não especificada"
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
        self._operator_invalidate_ops_status_cache(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
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

        pdf_candidates = self._piece_pdf_references(piece)
        if isinstance(ref_info, dict):
            pdf_candidates.extend(self._piece_pdf_references(ref_info))
        for raw in self._dedupe_document_references(pdf_candidates):
            path_txt = str(raw or "").strip()
            if not path_txt:
                continue
            path = self._resolve_file_reference(path_txt)
            if path is not None and path.exists():
                os.startfile(str(path))
                return str(path)

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
        raise ValueError("Esta peça não tem PDF/desenho associado ou o ficheiro não existe.")
