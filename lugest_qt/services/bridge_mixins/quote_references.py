from __future__ import annotations

import re
from typing import Any


class QuoteReferencesBackendMixin:
    """Legacy adapter for quote references; see BACKEND_GUIDE.md."""

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
        operations_source: list[Any] = []
        raw_operation_text = str(row.get("operacao", "") or "").strip()
        if raw_operation_text:
            operations_source.append(raw_operation_text)
        for key in ("operacoes_lista", "operacoes_fluxo", "operacoes_detalhe"):
            raw_items = row.get(key)
            if isinstance(raw_items, list):
                operations_source.extend(raw_items)
        for key in ("tempos_operacao", "custos_operacao"):
            raw_map = row.get(key)
            if isinstance(raw_map, dict):
                operations_source.extend(str(name or "").strip() for name in raw_map.keys() if str(name or "").strip())
        operations = [str(op or "").strip() for op in list(self.quote_parse_operacoes_lista(operations_source) or []) if str(op or "").strip()]
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
            if ref_ext in dict(data.get("orc_refs_removed", {}) or {}):
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
                "material_subtype": str(line.get("material_subtype", existing_ref.get("material_subtype", "")) or "").strip(),
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
                "material_subtype": record.get("material_subtype", ""),
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
        if ref_ext_txt in dict(self.ensure_data().get("orc_refs_removed", {}) or {}):
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

    def orc_reference_update(self, ref_externa: str, payload: dict[str, Any]) -> dict[str, Any]:
        ref_ext = str(ref_externa or payload.get("ref_externa", "") or "").strip()
        if not ref_ext:
            raise ValueError("Referencia externa obrigatoria.")
        data = self.ensure_data()
        refs_db = data.setdefault("orc_refs", {})
        data.setdefault("orc_refs_removed", {}).pop(ref_ext, None)
        existing = dict(refs_db.get(ref_ext, {}) or {})
        record = {
            **existing,
            "ref_externa": ref_ext,
            "ref_interna": str(payload.get("ref_interna", existing.get("ref_interna", "")) or "").strip(),
            "descricao": str(payload.get("descricao", existing.get("descricao", "")) or "").strip(),
            "material": str(payload.get("material", existing.get("material", "")) or "").strip(),
            "material_subtype": str(payload.get("material_subtype", existing.get("material_subtype", "")) or "").strip(),
            "espessura": str(payload.get("espessura", existing.get("espessura", "")) or "").strip(),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", payload.get("preco", existing.get("preco_unit", 0))), 0), 4),
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", existing.get("tempo_peca_min", 0)), 0), 3),
            "operacao": str(payload.get("operacao", payload.get("operacoes", existing.get("operacao", ""))) or "").strip(),
            "desenho": str(payload.get("desenho", existing.get("desenho", existing.get("desenho_path", ""))) or "").strip(),
            "cliente_codigo": str(payload.get("cliente_codigo", existing.get("cliente_codigo", "")) or "").strip(),
            "origem_doc": str(payload.get("origem_doc", existing.get("origem_doc", "Catalogo")) or "").strip() or "Catalogo",
            "origem_tipo": str(payload.get("origem_tipo", existing.get("origem_tipo", "Historico")) or "").strip() or "Historico",
            "updated_at": self.desktop_main.now_iso(),
            "operacoes_lista": list(payload.get("operacoes_lista", existing.get("operacoes_lista", [])) or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(payload.get("operacoes_fluxo", existing.get("operacoes_fluxo", [])) or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(payload.get("operacoes_detalhe", existing.get("operacoes_detalhe", [])) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(payload.get("tempos_operacao", existing.get("tempos_operacao", {})) or {}),
            "custos_operacao": dict(payload.get("custos_operacao", existing.get("custos_operacao", {})) or {}),
            "quote_cost_snapshot": dict(payload.get("quote_cost_snapshot", existing.get("quote_cost_snapshot", {})) or {}),
        }
        refs_db[ref_ext] = record
        piece_history = data.setdefault("peca_hist", {})
        piece_history[ref_ext] = {
            **dict(piece_history.get(ref_ext, {}) or {}),
            "ref_interna": record.get("ref_interna", ""),
            "ref_externa": ref_ext,
            "descricao": record.get("descricao", ""),
            "material": record.get("material", ""),
            "material_subtype": record.get("material_subtype", ""),
            "espessura": record.get("espessura", ""),
            "Operacoes": record.get("operacao", ""),
            "Observacoes": record.get("descricao", ""),
            "desenho": record.get("desenho", ""),
            "cliente_codigo": record.get("cliente_codigo", ""),
            "origem_doc": record.get("origem_doc", ""),
            "origem_tipo": record.get("origem_tipo", ""),
            "updated_at": record.get("updated_at", ""),
        }
        self._upsert_orc_ref_history_entry(ref_ext, record)
        self._save(force=True)
        return dict(record)

    def orc_reference_remove(self, ref_externa: str) -> None:
        ref_ext = str(ref_externa or "").strip()
        if not ref_ext:
            raise ValueError("Referencia externa obrigatoria.")
        data = self.ensure_data()
        data.setdefault("orc_refs", {}).pop(ref_ext, None)
        data.setdefault("peca_hist", {}).pop(ref_ext, None)
        data.setdefault("orc_refs_removed", {})[ref_ext] = self.desktop_main.now_iso()
        self._delete_orc_ref_history_entry(ref_ext)
        self._save(force=True)

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
            ref_externa_txt = str(payload.get("ref_externa", "") or "").strip()
            if ref_externa_txt and ref_externa_txt in dict(data.get("orc_refs_removed", {}) or {}):
                return
            row = {
                "ref_externa": ref_externa_txt,
                "ref_interna": str(payload.get("ref_interna", "") or "").strip(),
                "descricao": str(payload.get("descricao", "") or "").strip(),
                "material": str(payload.get("material", "") or "").strip(),
                "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
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
                "cliente_codigo": row_client(
                    str(payload.get("ref_interna", "") or "").strip(),
                    str(payload.get("ref_externa", "") or "").strip(),
                    str(payload.get("cliente_codigo", "") or payload.get("cliente", "") or "").strip(),
                ),
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
                    for field in ("descricao", "material", "material_subtype", "espessura", "operacoes", "desenho"):
                        if str(candidate.get(field, "") or "").strip():
                            richness += 2
                    if float(candidate.get("preco", 0) or 0) > 0:
                        richness += 1
                    if float(candidate.get("tempo_peca_min", 0) or 0) > 0:
                        richness += 1
                    return base + richness

                for field in ("descricao", "material", "material_subtype", "espessura", "operacoes", "desenho"):
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
                        "material_subtype": str(line.get("material_subtype", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "preco_unit": line.get("preco_unit", 0),
                        "tempo_peca_min": line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)),
                        "operacoes": self.quote_format_operacoes(line.get("operacao", "")),
                        "desenho": str(line.get("desenho", "") or "").strip(),
                        "origem_doc": str(orc.get("numero", "") or "").strip(),
                        "origem_tipo": "Orcamento",
                        "cliente_codigo": explicit_client,
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
                        "material_subtype": str(piece.get("material_subtype", "") or "").strip(),
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
                        "cliente_codigo": explicit_client,
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
                    "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
                    "espessura": str(payload.get("espessura", "") or "").strip(),
                    "preco_unit": payload.get("preco_unit", 0),
                    "tempo_peca_min": payload.get("tempo_peca_min", 0),
                    "operacoes": self.quote_format_operacoes(payload.get("operacao", "")),
                    "desenho": str(payload.get("desenho", "") or payload.get("desenho_path", "") or "").strip(),
                    "origem_doc": "Historico",
                    "origem_tipo": "Historico",
                    "cliente_codigo": str(payload.get("cliente_codigo", "") or "").strip(),
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
