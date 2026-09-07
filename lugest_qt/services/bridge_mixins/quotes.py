from __future__ import annotations
from lugest_modules.quotes.application.line_normalization import normalize_line
from lugest_qt.services.quote_rules_composition import normalize_line_ports
from lugest_modules.quotes.infrastructure.assembly_report import render_assembly_sheet
from lugest_qt.services.quote_rules_composition import render_assembly_sheet_ports
from lugest_modules.quotes.infrastructure.nesting_report import render_nesting_study
from lugest_qt.services.quote_rules_composition import render_nesting_study_ports
from lugest_qt.services.quote_nesting_composition import nesting_sql_store, nesting_study_service

import uuid

import copy
import json
import math
import os
import tempfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


ORC_STANDARD_IVA_PERC = 23.0
ORC_DEFAULT_DELIVERY_TEXT = "A combinar com o departamento de planeamento."


class QuotesBridgeMixin:
    """Quote, assembly model, and quote-to-order operations for the Qt bridge."""

    def _quote_standard_iva_perc(self) -> float:
        return ORC_STANDARD_IVA_PERC

    def _quote_default_delivery_text(self) -> str:
        return ORC_DEFAULT_DELIVERY_TEXT

    def _normalize_quote_discount_mode(self, value: Any) -> str:
        mode = str(value or "").strip().lower()
        return mode if mode in {"total", "lotes_espessura"} else "total"

    def _normalize_quote_discount_groups(self, value: Any) -> list[str]:
        groups: list[str] = []
        for item in list(value or []):
            clean = str(item or "").strip()
            if clean and clean not in groups:
                groups.append(clean)
        return groups

    def _peek_next_orc_number(self) -> str:
        data = self.ensure_data()
        try:
            return str(self.desktop_main.peek_next_orc_numero(data))
        except Exception:
            try:
                seq = int(data.get("orc_seq", 1) or 1)
            except Exception:
                seq = 1
            year = int(getattr(datetime.now(), "year", 0) or 0)
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
        current_year = str(datetime.now().year)
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
        return nesting_sql_store(self).ensure_table(conn)

    def _mysql_orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        return nesting_sql_store(self).studies(numero)

    def _mysql_save_orc_nesting_study(self, numero: str, group_key: str, group_label: str, payload: dict[str, Any]) -> None:
        return nesting_sql_store(self).save(numero, group_key, group_label, payload)

    def _mysql_delete_orc_nesting_studies(self, numero: str, group_key: str = "") -> None:
        return nesting_sql_store(self).delete(numero, group_key)

    def orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        return nesting_study_service(self).studies(numero)

    def orc_save_nesting_study(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        return nesting_study_service(self).save(numero, payload)

    def orc_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
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
                    "dimensao": str(row.get("dimensao", row.get("dimensoes", "")) or "").strip(),
                    "dimensoes": str(row.get("dimensao", row.get("dimensoes", "")) or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "material_family": str(row.get("material_family", "") or "").strip(),
                    "material_subtype": str(row.get("material_subtype", "") or "").strip(),
                    "stock_item_kind": str(row.get("stock_item_kind", "") or "").strip(),
                    "material_supplied_by_client": material_supplied_by_client,
                    "material_fornecido_cliente": material_supplied_by_client,
                    "material_cost_included": (False if material_supplied_by_client else bool(row.get("material_cost_included", True))),
                    "espessura": self._fmt(row.get("espessura", "")),
                    "operacao": str(row.get("operacao", "") or "").strip(),
                    "produto_codigo": str(row.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(row.get("produto_unid", "") or "").strip(),
                    "_product_pending_create": bool(row.get("_product_pending_create", False)),
                    "conjunto_codigo": str(row.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(row.get("conjunto_nome", "") or "").strip(),
                    "conjunto_param_codigo": str(row.get("conjunto_param_codigo", "") or "").strip(),
                    "grupo_uuid": str(row.get("grupo_uuid", "") or "").strip(),
                    "ficha_tecnica": self._normalize_conjunto_technical_sheet(row.get("ficha_tecnica", {})),
                    "qtd_base": round(self._parse_float(row.get("qtd_base", row.get("qtd", 0)), 0), 2),
                    "tempo_peca_min": display_time,
                    "qtd": line_qty,
                    "preco_unit": display_price,
                    "total": round(line_qty * display_price, 2),
                    "desenho": str(row.get("desenho", "") or "").strip(),
                    "desenho_pdf": str(row.get("desenho_pdf", "") or "").strip(),
                    "desenhos_pdf": [
                        str(item or "").strip()
                        for item in list(row.get("desenhos_pdf", []) or [])
                        if str(item or "").strip()
                    ],
                    "ficheiros": [
                        str(item or "").strip()
                        for item in list(row.get("ficheiros", []) or [])
                        if str(item or "").strip()
                    ],
                    "laser_base_active": laser_base_active,
                    "laser_base_tempo_unit": laser_base_tempo,
                    "laser_base_preco_unit": laser_base_preco,
                    "machine": str(row.get("machine", "") or "").strip(),
                    "laser_machine": str(row.get("laser_machine", row.get("machine", "")) or "").strip(),
                    "commercial_profile": str(row.get("commercial_profile", "") or "").strip(),
                    "gas": str(row.get("gas", "") or "").strip(),
                    "laser_snapshot": dict(row.get("laser_snapshot", {}) or {}),
                    "laser_source_mode": str(row.get("laser_source_mode", "") or "").strip(),
                    "laser_batch_id": str(row.get("laser_batch_id", "") or "").strip(),
                    "operacoes_lista": list(snapshot.get("operacoes", []) or []),
                    "operacoes_fluxo": [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                    "operacoes_detalhe": [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                    "tempos_operacao": dict(snapshot.get("tempos_operacao", {}) or {}),
                    "custos_operacao": dict(snapshot.get("custos_operacao", {}) or {}),
                    "quote_cost_snapshot": dict(snapshot.get("quote_cost_snapshot", {}) or {}),
                    "stock_material_id": str(row.get("stock_material_id", "") or "").strip(),
                    "price_per_kg": round(self._parse_float(row.get("price_per_kg", 0), 0), 4),
                    "price_base_value": round(self._parse_float(row.get("price_base_value", 0), 0), 4),
                    "price_base_label": str(row.get("price_base_label", "") or "").strip(),
                    "price_markup_pct": round(self._parse_float(row.get("price_markup_pct", 0), 0), 2),
                    "stock_metric_value": round(self._parse_float(row.get("stock_metric_value", 0), 0), 4),
                    "meters_per_unit": round(self._parse_float(row.get("meters_per_unit", 0), 0), 3),
                    "kg_per_m": round(self._parse_float(row.get("kg_per_m", 0), 0), 4),
                    "length_mm": round(self._parse_float(row.get("length_mm", 0), 0), 1),
                    "width_mm": round(self._parse_float(row.get("width_mm", 0), 0), 1),
                    "thickness_mm": round(self._parse_float(row.get("thickness_mm", 0), 0), 2),
                    "diameter_mm": round(self._parse_float(row.get("diameter_mm", 0), 0), 1),
                    "profile_section": str(row.get("profile_section", "") or "").strip(),
                    "profile_size": str(row.get("profile_size", "") or "").strip(),
                    "tube_section": str(row.get("tube_section", "") or "").strip(),
                    "quality": str(row.get("quality", "") or "").strip(),
                    "calc_mode": str(row.get("calc_mode", "") or "").strip(),
                    "discount_group_key": str(row.get("discount_group_key", "") or "").strip(),
                }
            )
        return {
            "numero": str(orc.get("numero", "") or "").strip(),
            "data": str(orc.get("data", "") or "").strip()[:10],
            "estado": str(orc.get("estado", "") or "").strip() or "Em edicao",
            "cliente": client,
            "posto_trabalho": self._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "iva_perc": self._quote_standard_iva_perc(),
            "desconto_perc": round(self._parse_float(orc.get("desconto_perc", 0), 0), 2),
            "desconto_modo": self._normalize_quote_discount_mode(orc.get("desconto_modo", "total")),
            "desconto_grupos": self._normalize_quote_discount_groups(orc.get("desconto_grupos", [])),
            "incremento_preco_perc": round(self._parse_float(orc.get("incremento_preco_perc", 0), 0), 2),
            "desconto_valor": round(self._parse_float(orc.get("desconto_valor", 0), 0), 2),
            "subtotal_linhas": round(self._parse_float(orc.get("subtotal_linhas", orc.get("subtotal_bruto", 0)), 0), 2),
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
            "prazo_entrega_texto": str(orc.get("prazo_entrega_texto", "") or self._quote_default_delivery_text()).strip(),
            "prazo_entrega_data": str(orc.get("prazo_entrega_data", "") or "").strip()[:10],
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
        for row in list(self.ensure_data().get("conjuntos_modelo", []) or []) + list(self.ensure_data().get("conjuntos", []) or []):
            codigo = str((row or {}).get("codigo", "") or "").strip().upper()
            digits = "".join(ch for ch in codigo if ch.isdigit())
            if digits:
                try:
                    highest = max(highest, int(digits))
                except Exception:
                    continue
        return f"MOD{highest + 1:04d}"

    def conjunto_next_param_codigo(self) -> str:
        highest = 0
        missing = 0
        for row in list(self.ensure_data().get("conjuntos", []) or []):
            raw = str((row or {}).get("param_codigo", "") or "").strip()
            digits = "".join(ch for ch in raw if ch.isdigit())
            if digits:
                highest = max(highest, int(digits))
            else:
                missing += 1
        return f"{highest + missing + 1:04d}"

    def _ensure_conjunto_param_codes(self) -> bool:
        changed = False
        used: set[str] = set()
        highest = 0
        rows = [row for row in list(self.ensure_data().get("conjuntos", []) or []) if isinstance(row, dict)]
        for row in rows:
            raw = str(row.get("param_codigo", "") or "").strip()
            digits = "".join(ch for ch in raw if ch.isdigit())
            if digits:
                normalized = f"{int(digits):04d}"
                if normalized not in used:
                    used.add(normalized)
                    highest = max(highest, int(normalized))
                    if raw != normalized:
                        row["param_codigo"] = normalized
                        changed = True
                    continue
            row["param_codigo"] = ""
        for row in rows:
            if str(row.get("param_codigo", "") or "").strip():
                continue
            highest += 1
            while f"{highest:04d}" in used:
                highest += 1
            row["param_codigo"] = f"{highest:04d}"
            used.add(row["param_codigo"])
            changed = True
        return changed

    def _conjunto_find_quote_source(self, item: dict[str, Any], conjunto_codigo: str) -> tuple[dict[str, Any] | None, str]:
        ref = str(item.get("source_ref_externa", "") or item.get("ref_externa", "") or "").strip()
        quote_number = str(item.get("source_quote_number", "") or "").strip()
        operation_norm = self.desktop_main.norm_text(str(item.get("operacao", "") or ""))
        if not ref or ("laser" not in operation_norm and not str(item.get("desenho", "") or "").strip()):
            return None, quote_number
        quotes = list(self.ensure_data().get("orcamentos", []) or [])
        if quote_number:
            quotes = sorted(quotes, key=lambda row: str(row.get("numero", "") or "") != quote_number)
        else:
            quotes = list(reversed(quotes))
        fallback: tuple[dict[str, Any] | None, str] = (None, "")
        for quote in quotes:
            number = str(quote.get("numero", "") or "").strip()
            for line in list(quote.get("linhas", []) or []):
                if str(line.get("ref_externa", "") or "").strip() != ref:
                    continue
                if str(line.get("conjunto_codigo", "") or "").strip() == conjunto_codigo:
                    return line, number
                if fallback[0] is None:
                    fallback = (line, number)
            if quote_number and number == quote_number and fallback[0] is not None:
                return fallback
        return fallback

    def _conjunto_live_item(self, raw_item: dict[str, Any], conjunto_codigo: str) -> tuple[dict[str, Any], bool]:
        item = self._normalize_assembly_model_item(dict(raw_item or {}))
        old_price = round(self._parse_float(item.get("preco_unit", 0), 0), 4)
        live_price = old_price
        source_type = "manual"
        source_label = "Valor manual"
        source_ref = ""
        linked = False

        if self.desktop_main.orc_line_is_product(item):
            product = self._product_lookup(item.get("produto_codigo", ""))
            if product is not None:
                live_price = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
                source_type = "product_stock"
                source_label = "Stock produtos"
                source_ref = str(product.get("codigo", "") or "").strip()
                linked = True
        else:
            stock_id = str(item.get("stock_material_id", "") or "").strip()
            material_record = self.material_by_id(stock_id) if stock_id else None
            if material_record is None and item.get("calc_mode") and self._parse_float(item.get("stock_metric_value", 0), 0) > 0:
                wanted_mode = self.desktop_main.norm_text(str(item.get("calc_mode", "") or ""))
                base_value = self._parse_float(item.get("price_base_value", 0), 0)
                candidates = []
                for candidate in list(self.ensure_data().get("materiais", []) or []):
                    candidate_mode = self.desktop_main.norm_text(
                        str(candidate.get("formato", "") or self.desktop_main.detect_materia_formato(candidate) or "")
                    )
                    if wanted_mode and candidate_mode != wanted_mode:
                        continue
                    delta = abs(self._parse_float(candidate.get("p_compra", 0), 0) - base_value)
                    candidates.append((delta, candidate))
                if candidates:
                    candidates.sort(key=lambda pair: pair[0])
                    if candidates[0][0] <= 0.0002:
                        material_record = candidates[0][1]
                        stock_id = str(material_record.get("id", "") or "").strip()
                        item["stock_material_id"] = stock_id
            if material_record is not None:
                preview = self.material_price_preview(material_record)
                metric = self._parse_float(item.get("stock_metric_value", 0), 0)
                base_label = str(item.get("price_base_label", "") or "").strip().lower()
                if metric > 0 and base_label:
                    current_base = self._parse_float(material_record.get("p_compra", 0), 0)
                    live_price = round(current_base * metric, 4)
                    item["price_base_value"] = round(current_base, 4)
                else:
                    live_price = round(self._parse_float(preview.get("preco_unid", old_price), old_price), 4)
                source_type = "material_stock"
                source_label = "Stock materia-prima"
                source_ref = stock_id
                linked = True
            elif self.desktop_main.orc_line_is_piece(item):
                quote_line, quote_number = self._conjunto_find_quote_source(item, conjunto_codigo)
                if quote_line is not None:
                    live_price = round(self._parse_float(quote_line.get("preco_unit", old_price), old_price), 4)
                    source_type = "quote_laser"
                    source_label = f"Orcamento laser {quote_number}".strip()
                    source_ref = str(quote_line.get("ref_externa", "") or "").strip()
                    item["source_quote_number"] = quote_number
                    item["source_ref_externa"] = source_ref
                    linked = True

        changed = abs(live_price - old_price) > 0.00005
        if changed:
            item["preco_anterior"] = old_price
            item["preco_unit"] = live_price
            item["preco_atualizado_em"] = self.desktop_main.now_iso()
        item["pricing_source"] = source_type
        item["pricing_source_label"] = source_label
        item["pricing_source_ref"] = source_ref
        item["pricing_linked"] = linked
        return item, changed or any(item.get(key) != raw_item.get(key) for key in (
            "stock_material_id", "source_quote_number", "source_ref_externa", "pricing_source", "pricing_source_ref"
        ))

    def _conjunto_refresh_model_prices(self, model: dict[str, Any]) -> bool:
        code = str(model.get("codigo", "") or "").strip()
        refreshed: list[dict[str, Any]] = []
        changed = False
        for index, raw_item in enumerate(list(model.get("itens", []) or []), start=1):
            item, item_changed = self._conjunto_live_item(dict(raw_item or {}), code)
            item["linha_ordem"] = index
            refreshed.append(item)
            changed = changed or item_changed
        total_cost = round(sum(self._parse_float(item.get("qtd", 0), 0) * self._parse_float(item.get("preco_unit", 0), 0) for item in refreshed), 2)
        margin = self._parse_float(model.get("margem_perc", 0), 0)
        total_final = round(total_cost * (1.0 + margin / 100.0), 2)
        if abs(total_cost - self._parse_float(model.get("total_custo", 0), 0)) > 0.005:
            changed = True
        if abs(total_final - self._parse_float(model.get("total_final", 0), 0)) > 0.005:
            changed = True
        model["itens"] = refreshed
        model["total_custo"] = total_cost
        model["total_final"] = total_final
        if changed:
            model["precos_atualizados_em"] = self.desktop_main.now_iso()
        return changed

    def conjunto_refresh_prices(self, codigo: str = "") -> dict[str, Any]:
        code = str(codigo or "").strip()
        changed = self._ensure_conjunto_param_codes()
        matched = False
        for model in list(self.ensure_data().get("conjuntos", []) or []):
            if not isinstance(model, dict) or (code and str(model.get("codigo", "") or "").strip() != code):
                continue
            matched = True
            changed = self._conjunto_refresh_model_prices(model) or changed
        if code and not matched:
            raise ValueError("Conjunto nao encontrado.")
        if changed:
            self._save(force=True)
        if code:
            model = next(row for row in self.ensure_data().get("conjuntos", []) if str(row.get("codigo", "") or "").strip() == code)
            return dict(model)
        return {"updated": changed}

    def _normalize_assembly_model_item(self, payload: dict[str, Any]) -> dict[str, Any]:
        item_type = self.desktop_main.normalize_orc_line_type(payload.get("tipo_item"))
        quantity = round(self._parse_float(payload.get("qtd", 0), 0), 2)
        if quantity <= 0:
            raise ValueError("Quantidade invalida no conjunto.")
        stock_item_kind = str(payload.get("stock_item_kind", "") or "").strip()
        if item_type == self.desktop_main.ORC_LINE_TYPE_PIECE and (
            stock_item_kind == "raw_material" or str(payload.get("stock_material_id", "") or "").strip()
        ):
            stock_item_kind = "raw_material"
        elif item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            stock_item_kind = "product"
        else:
            stock_item_kind = ""
        item = {
            "tipo_item": item_type,
            "stock_item_kind": stock_item_kind,
            "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "dimensao": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
            "material": str(payload.get("material", "") or "").strip(),
            "espessura": str(payload.get("espessura", "") or "").strip(),
            "material_supplied_by_client": bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False)),
            "material_fornecido_cliente": bool(payload.get("material_fornecido_cliente", False) or payload.get("material_supplied_by_client", False)),
            "material_cost_included": bool(payload.get("material_cost_included", True)),
            "operacao": str(payload.get("operacao", "") or "").strip(),
            "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
            "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
            "qtd": quantity,
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", 0), 0), 4),
            "desenho": str(payload.get("desenho", "") or "").strip(),
            "desenho_pdf": str(payload.get("desenho_pdf", "") or "").strip(),
            "desenhos_pdf": [
                str(value or "").strip()
                for value in list(payload.get("desenhos_pdf", []) or [])
                if str(value or "").strip()
            ],
            "machine": str(payload.get("machine", "") or "").strip(),
            "laser_machine": str(payload.get("laser_machine", payload.get("machine", "")) or "").strip(),
            "commercial_profile": str(payload.get("commercial_profile", "") or "").strip(),
            "gas": str(payload.get("gas", "") or "").strip(),
            "laser_snapshot": dict(payload.get("laser_snapshot", {}) or {}),
            "laser_source_mode": str(payload.get("laser_source_mode", "") or "").strip(),
            "laser_batch_id": str(payload.get("laser_batch_id", "") or "").strip(),
            "calc_mode": str(payload.get("calc_mode", "") or "").strip(),
            "descricao_base": str(payload.get("descricao_base", "") or "").strip(),
            "weight_total": round(self._parse_float(payload.get("weight_total", 0), 0), 3),
            "total_cost": round(self._parse_float(payload.get("total_cost", 0), 0), 2),
            "quantity_units": round(self._parse_float(payload.get("quantity_units", quantity), quantity), 2),
            "price_per_kg": round(self._parse_float(payload.get("price_per_kg", 0), 0), 4),
            "price_base_value": round(self._parse_float(payload.get("price_base_value", 0), 0), 4),
            "price_markup_pct": round(self._parse_float(payload.get("price_markup_pct", 0), 0), 2),
            "stock_metric_value": round(self._parse_float(payload.get("stock_metric_value", 0), 0), 4),
            "meters_per_unit": round(self._parse_float(payload.get("meters_per_unit", 0), 0), 3),
            "kg_per_m": round(self._parse_float(payload.get("kg_per_m", 0), 0), 4),
            "length_mm": round(self._parse_float(payload.get("length_mm", 0), 0), 1),
            "width_mm": round(self._parse_float(payload.get("width_mm", 0), 0), 1),
            "thickness_mm": round(self._parse_float(payload.get("thickness_mm", 0), 0), 2),
            "density": round(self._parse_float(payload.get("density", 0), 0), 1),
            "diameter_mm": round(self._parse_float(payload.get("diameter_mm", 0), 0), 1),
            "manual_unit_price": round(self._parse_float(payload.get("manual_unit_price", 0), 0), 4),
            "profile_section": str(payload.get("profile_section", "") or "").strip(),
            "profile_size": str(payload.get("profile_size", "") or "").strip(),
            "tube_section": str(payload.get("tube_section", "") or "").strip(),
            "quality": str(payload.get("quality", "") or "").strip(),
            "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
            "hint": str(payload.get("hint", "") or "").strip(),
            "price_base_label": str(payload.get("price_base_label", "") or "").strip(),
            "material_family": str(payload.get("material_family", "") or "").strip(),
            "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
            "operacoes_lista": list(payload.get("operacoes_lista", []) or []),
            "operacoes_fluxo": [dict(row or {}) for row in list(payload.get("operacoes_fluxo", []) or []) if isinstance(row, dict)],
            "operacoes_detalhe": [dict(row or {}) for row in list(payload.get("operacoes_detalhe", []) or []) if isinstance(row, dict)],
            "tempos_operacao": dict(payload.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(payload.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(payload.get("quote_cost_snapshot", {}) or {}),
            "laser_base_active": bool(payload.get("laser_base_active", False)),
            "laser_base_tempo_unit": round(self._parse_float(payload.get("laser_base_tempo_unit", 0), 0), 4),
            "laser_base_preco_unit": round(self._parse_float(payload.get("laser_base_preco_unit", 0), 0), 4),
            "source_quote_number": str(payload.get("source_quote_number", "") or "").strip(),
            "source_ref_externa": str(payload.get("source_ref_externa", "") or "").strip(),
            "pricing_source": str(payload.get("pricing_source", "") or "").strip(),
            "pricing_source_ref": str(payload.get("pricing_source_ref", "") or "").strip(),
            "preco_anterior": round(self._parse_float(payload.get("preco_anterior", 0), 0), 4),
            "preco_atualizado_em": str(payload.get("preco_atualizado_em", "") or "").strip(),
        }
        if item_type == self.desktop_main.ORC_LINE_TYPE_PIECE:
            if not item["descricao"]:
                raise ValueError("Descricao obrigatoria na peca do conjunto.")
            if not item["material"] or not item["espessura"]:
                raise ValueError("Material e espessura sao obrigatorios nas pecas fabricadas.")
            item["produto_codigo"] = ""
            item["produto_unid"] = ""
            if stock_item_kind == "raw_material":
                item["ref_interna"] = ""
                item["operacao"] = ""
                item["desenho"] = ""
                item["tempo_peca_min"] = 0.0
                item["operacoes_lista"] = []
                item["operacoes_fluxo"] = []
                item["operacoes_detalhe"] = []
                item["tempos_operacao"] = {}
                item["custos_operacao"] = {}
        elif item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            product = self._product_lookup(item["produto_codigo"])
            if product is None and not item["descricao"]:
                raise ValueError("Descricao obrigatoria no produto.")
            item["_product_pending_create"] = product is None
            item["descricao"] = item["descricao"] or str((product or {}).get("descricao", "") or "").strip()
            item["produto_unid"] = item["produto_unid"] or str((product or {}).get("unid", "") or "UN").strip()
            if product is not None and item["preco_unit"] <= 0:
                item["preco_unit"] = round(self._parse_float(self.desktop_main.produto_preco_venda(product), 0), 4)
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
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
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
            "param_codigo": str(model.get("param_codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "ficha_tecnica": self._normalize_conjunto_technical_sheet(model.get("ficha_tecnica", {})),
            "itens": items,
        }

    @staticmethod
    def _normalize_conjunto_technical_sheet(raw: Any) -> dict[str, str]:
        source = dict(raw or {}) if isinstance(raw, dict) else {}
        keys = (
            "familia_produto",
            "aplicacao",
            "modelo_versao",
            "configuracao",
            "dimensoes_gerais",
            "materiais_acabamentos",
            "caracteristicas",
            "requisitos_instalacao",
            "normas_conformidade",
            "controlo_qualidade",
        )
        return {key: str(source.get(key, "") or "").strip() for key in keys}

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
            "param_codigo": str(payload.get("param_codigo", "") or "").strip(),
            "descricao": descricao,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "template": bool(payload.get("template", False)),
            "origem": str(payload.get("origem", "") or "").strip(),
            "created_at": str(payload.get("created_at", "") or "").strip() or self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "ficha_tecnica": self._normalize_conjunto_technical_sheet(payload.get("ficha_tecnica", {})),
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
            if "ficha_tecnica" not in payload:
                model["ficha_tecnica"] = self._normalize_conjunto_technical_sheet(existing.get("ficha_tecnica", {}))
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
        group_uuid = uuid.uuid4().hex[:12].upper()
        rows: list[dict[str, Any]] = []
        for item in list(detail.get("itens", []) or []):
            line = {
                "tipo_item": self.desktop_main.normalize_orc_line_type(item.get("tipo_item")),
                "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                "ref_interna": "",
                "ref_externa": str(item.get("ref_externa", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(item.get("material", "") or "").strip(),
                "material_family": str(item.get("material_family", "") or "").strip(),
                "material_subtype": str(item.get("material_subtype", "") or "").strip(),
                "material_supplied_by_client": bool(item.get("material_supplied_by_client", False) or item.get("material_fornecido_cliente", False)),
                "material_fornecido_cliente": bool(item.get("material_fornecido_cliente", False) or item.get("material_supplied_by_client", False)),
                "material_cost_included": bool(item.get("material_cost_included", True)),
                "espessura": str(item.get("espessura", "") or "").strip(),
                "operacao": str(item.get("operacao", "") or "").strip(),
                "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                "conjunto_codigo": str(detail.get("codigo", "") or "").strip(),
                "conjunto_nome": str(detail.get("descricao", "") or "").strip(),
                "conjunto_param_codigo": str(detail.get("param_codigo", "") or "").strip(),
                "grupo_uuid": group_uuid,
                "ficha_tecnica": dict(detail.get("ficha_tecnica", {}) or {}),
                "qtd_base": round(self._parse_float(item.get("qtd", 0), 0), 2),
                "tempo_peca_min": round(self._parse_float(item.get("tempo_peca_min", 0), 0), 2),
                "qtd": round(self._parse_float(item.get("qtd", 0), 0) * multiplier, 2),
                "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                "desenho": str(item.get("desenho", "") or "").strip(),
                "desenho_pdf": str(item.get("desenho_pdf", "") or "").strip(),
                "desenhos_pdf": [str(value or "").strip() for value in list(item.get("desenhos_pdf", []) or []) if str(value or "").strip()],
                "machine": str(item.get("machine", "") or "").strip(),
                "laser_machine": str(item.get("laser_machine", item.get("machine", "")) or "").strip(),
                "commercial_profile": str(item.get("commercial_profile", "") or "").strip(),
                "gas": str(item.get("gas", "") or "").strip(),
                "laser_snapshot": dict(item.get("laser_snapshot", {}) or {}),
                "laser_source_mode": str(item.get("laser_source_mode", "") or "").strip(),
                "laser_batch_id": str(item.get("laser_batch_id", "") or "").strip(),
                "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                "_product_pending_create": bool(item.get("_product_pending_create", False)),
                "price_base_value": round(self._parse_float(item.get("price_base_value", 0), 0), 4),
                "price_base_label": str(item.get("price_base_label", "") or "").strip(),
                "stock_metric_value": round(self._parse_float(item.get("stock_metric_value", 0), 0), 4),
                "meters_per_unit": round(self._parse_float(item.get("meters_per_unit", 0), 0), 3),
                "kg_per_m": round(self._parse_float(item.get("kg_per_m", 0), 0), 4),
                "length_mm": round(self._parse_float(item.get("length_mm", 0), 0), 1),
                "width_mm": round(self._parse_float(item.get("width_mm", 0), 0), 1),
                "thickness_mm": round(self._parse_float(item.get("thickness_mm", 0), 0), 2),
                "diameter_mm": round(self._parse_float(item.get("diameter_mm", 0), 0), 1),
                "profile_section": str(item.get("profile_section", "") or "").strip(),
                "profile_size": str(item.get("profile_size", "") or "").strip(),
                "tube_section": str(item.get("tube_section", "") or "").strip(),
                "quality": str(item.get("quality", "") or "").strip(),
                "calc_mode": str(item.get("calc_mode", "") or "").strip(),
            }
            if self.desktop_main.orc_line_is_product(line) and not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            rows.append(line)
        return rows

    def conjunto_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        self.conjunto_refresh_prices()
        rows: list[dict[str, Any]] = []
        for model in list(self.ensure_data().get("conjuntos", []) or []):
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "param_codigo": str(model.get("param_codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.desktop_main.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.desktop_main.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.desktop_main.orc_line_is_service(item)),
                "total_custo": round(self._parse_float(model.get("total_custo", 0), 0), 2),
                "total_final": round(self._parse_float(model.get("total_final", 0), 0), 2),
                "margem_perc": round(self._parse_float(model.get("margem_perc", 0), 0), 2),
                "notas": str(model.get("notas", "") or "").strip(),
                "created_at": str(model.get("created_at", "") or "").strip(),
                "updated_at": str(model.get("updated_at", "") or "").strip(),
                "precos_atualizados_em": str(model.get("precos_atualizados_em", "") or "").strip(),
                "itens_ligados": sum(1 for item in items if bool(item.get("pricing_linked"))),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows

    def conjunto_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        self.conjunto_refresh_prices(code)
        model = next(
            (
                row
                for row in list(self.ensure_data().get("conjuntos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [dict(item or {}) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "param_codigo": str(model.get("param_codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "margem_perc": round(self._parse_float(model.get("margem_perc", 0), 0), 2),
            "total_custo": round(self._parse_float(model.get("total_custo", 0), 0), 2),
            "total_final": round(self._parse_float(model.get("total_final", 0), 0), 2),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "precos_atualizados_em": str(model.get("precos_atualizados_em", "") or "").strip(),
            "ficha_tecnica": self._normalize_conjunto_technical_sheet(model.get("ficha_tecnica", {})),
            "itens": items,
        }

    def conjunto_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        self._ensure_conjunto_param_codes()
        code = str(payload.get("codigo", "") or "").strip() or self._next_assembly_model_code()
        descricao = str(payload.get("descricao", "") or "").strip()
        if not descricao:
            raise ValueError("Descricao obrigatoria no conjunto.")
        items = [self._normalize_assembly_model_item(dict(row or {})) for row in list(payload.get("itens", []) or [])]
        if not items:
            raise ValueError("O conjunto precisa de pelo menos um item.")
        model = {
            "codigo": code,
            "param_codigo": str(payload.get("param_codigo", "") or "").strip(),
            "descricao": descricao,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "template": bool(payload.get("template", False)),
            "origem": str(payload.get("origem", "") or "").strip(),
            "margem_perc": round(self._parse_float(payload.get("margem_perc", 0), 0), 2),
            "total_custo": round(self._parse_float(payload.get("total_custo", 0), 0), 2),
            "total_final": round(self._parse_float(payload.get("total_final", 0), 0), 2),
            "created_at": str(payload.get("created_at", "") or "").strip() or self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "ficha_tecnica": self._normalize_conjunto_technical_sheet(payload.get("ficha_tecnica", {})),
            "itens": [{**item, "linha_ordem": index} for index, item in enumerate(items, start=1)],
        }
        existing = next(
            (
                row
                for row in list(data.get("conjuntos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if existing is None:
            used_params = {
                str(row.get("param_codigo", "") or "").strip()
                for row in list(data.get("conjuntos", []) or [])
                if isinstance(row, dict)
            }
            if not model["param_codigo"] or model["param_codigo"] in used_params:
                model["param_codigo"] = self.conjunto_next_param_codigo()
            data.setdefault("conjuntos", []).append(model)
        else:
            model["param_codigo"] = str(existing.get("param_codigo", "") or model["param_codigo"] or self.conjunto_next_param_codigo()).strip()
            model["created_at"] = str(existing.get("created_at", "") or "").strip() or model["created_at"]
            if "ficha_tecnica" not in payload:
                model["ficha_tecnica"] = self._normalize_conjunto_technical_sheet(existing.get("ficha_tecnica", {}))
            existing.update(model)
        self._conjunto_refresh_model_prices(model if existing is None else existing)
        self._save(force=True)
        return self.conjunto_detail(code)

    def conjunto_remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        rows = list(self.ensure_data().get("conjuntos", []) or [])
        filtered = [row for row in rows if str(row.get("codigo", "") or "").strip() != code]
        if len(filtered) == len(rows):
            raise ValueError("Conjunto nao encontrado.")
        self.ensure_data()["conjuntos"] = filtered
        self._save(force=True)

    def conjunto_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        detail = self.conjunto_detail(codigo)
        multiplier = round(self._parse_float(quantity, 0), 2)
        if multiplier <= 0:
            raise ValueError("Quantidade do conjunto invalida.")
        group_uuid = uuid.uuid4().hex[:12].upper()
        rows: list[dict[str, Any]] = []
        for item in list(detail.get("itens", []) or []):
            line = {
                "tipo_item": self.desktop_main.normalize_orc_line_type(item.get("tipo_item")),
                "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                "ref_interna": "",
                "ref_externa": str(item.get("ref_externa", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(item.get("material", "") or "").strip(),
                "material_family": str(item.get("material_family", "") or "").strip(),
                "material_subtype": str(item.get("material_subtype", "") or "").strip(),
                "material_supplied_by_client": bool(item.get("material_supplied_by_client", False) or item.get("material_fornecido_cliente", False)),
                "material_fornecido_cliente": bool(item.get("material_fornecido_cliente", False) or item.get("material_supplied_by_client", False)),
                "material_cost_included": bool(item.get("material_cost_included", True)),
                "espessura": str(item.get("espessura", "") or "").strip(),
                "operacao": str(item.get("operacao", "") or "").strip(),
                "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                "conjunto_codigo": str(detail.get("codigo", "") or "").strip(),
                "conjunto_nome": str(detail.get("descricao", "") or "").strip(),
                "conjunto_param_codigo": str(detail.get("param_codigo", "") or "").strip(),
                "grupo_uuid": group_uuid,
                "ficha_tecnica": dict(detail.get("ficha_tecnica", {}) or {}),
                "qtd_base": round(self._parse_float(item.get("qtd", 0), 0), 2),
                "tempo_peca_min": round(self._parse_float(item.get("tempo_peca_min", 0), 0), 2),
                "qtd": round(self._parse_float(item.get("qtd", 0), 0) * multiplier, 2),
                "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                "desenho": str(item.get("desenho", "") or "").strip(),
                "desenho_pdf": str(item.get("desenho_pdf", "") or "").strip(),
                "desenhos_pdf": [str(value or "").strip() for value in list(item.get("desenhos_pdf", []) or []) if str(value or "").strip()],
                "machine": str(item.get("machine", "") or "").strip(),
                "laser_machine": str(item.get("laser_machine", item.get("machine", "")) or "").strip(),
                "commercial_profile": str(item.get("commercial_profile", "") or "").strip(),
                "gas": str(item.get("gas", "") or "").strip(),
                "laser_snapshot": dict(item.get("laser_snapshot", {}) or {}),
                "laser_source_mode": str(item.get("laser_source_mode", "") or "").strip(),
                "laser_batch_id": str(item.get("laser_batch_id", "") or "").strip(),
                "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                "_product_pending_create": bool(item.get("_product_pending_create", False)),
                "price_base_value": round(self._parse_float(item.get("price_base_value", 0), 0), 4),
                "price_base_label": str(item.get("price_base_label", "") or "").strip(),
                "stock_metric_value": round(self._parse_float(item.get("stock_metric_value", 0), 0), 4),
                "meters_per_unit": round(self._parse_float(item.get("meters_per_unit", 0), 0), 3),
                "kg_per_m": round(self._parse_float(item.get("kg_per_m", 0), 0), 4),
                "length_mm": round(self._parse_float(item.get("length_mm", 0), 0), 1),
                "width_mm": round(self._parse_float(item.get("width_mm", 0), 0), 1),
                "thickness_mm": round(self._parse_float(item.get("thickness_mm", 0), 0), 2),
                "diameter_mm": round(self._parse_float(item.get("diameter_mm", 0), 0), 1),
                "profile_section": str(item.get("profile_section", "") or "").strip(),
                "profile_size": str(item.get("profile_size", "") or "").strip(),
                "tube_section": str(item.get("tube_section", "") or "").strip(),
                "quality": str(item.get("quality", "") or "").strip(),
                "calc_mode": str(item.get("calc_mode", "") or "").strip(),
            }
            if self.desktop_main.orc_line_is_product(line) and not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            rows.append(line)
        return rows

    def _normalize_orc_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        return normalize_line(normalize_line_ports(self), payload)

    def orc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip()
        existing = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if existing is None:
            year = int(getattr(datetime.now(), "year", 0) or 0)
            automatic_number = not numero or numero.upper().startswith(f"ORC-{year}-")
            if automatic_number:
                numero = str(self.desktop_main.next_orc_numero(data))
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
            # Repairs can save and replace the active snapshot. Subsequent
            # reference allocation must use that snapshot, not the old mapping.
            data = self.ensure_data()
            taken_refs, _pairs = self._active_client_ref_usage(client_code, exclude_orc_numero=numero)
            reusable_pairs = self._known_client_ref_pairs(client_code)
            seen_refs: set[str] = set()
            seen_pairs: set[tuple[str, str]] = set()
            reserved_refs = set(taken_refs)
            for line in lines:
                if not self.desktop_main.orc_line_is_piece(line):
                    line["ref_interna"] = ""
                    continue
                if self._quote_line_is_raw_material(line):
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
        iva_perc = self._quote_standard_iva_perc()
        desconto_perc = round(max(0.0, min(100.0, self._parse_float(payload.get("desconto_perc", (existing or {}).get("desconto_perc", 0)), 0))), 2)
        desconto_modo = self._normalize_quote_discount_mode(payload.get("desconto_modo", (existing or {}).get("desconto_modo", "total")))
        desconto_grupos = self._normalize_quote_discount_groups(payload.get("desconto_grupos", (existing or {}).get("desconto_grupos", [])))
        incremento_preco_perc = round(max(-100.0, self._parse_float(payload.get("incremento_preco_perc", (existing or {}).get("incremento_preco_perc", 0)), 0)), 2)
        preco_transporte = round(self._parse_float(payload.get("preco_transporte", 0), 0), 2)
        custo_transporte = round(self._parse_float(payload.get("custo_transporte", (existing or {}).get("custo_transporte", 0)), 0), 2)
        paletes = round(self._parse_float(payload.get("paletes", (existing or {}).get("paletes", 0)), 0), 2)
        peso_bruto_kg = round(self._parse_float(payload.get("peso_bruto_kg", (existing or {}).get("peso_bruto_kg", 0)), 0), 2)
        volume_m3 = round(self._parse_float(payload.get("volume_m3", (existing or {}).get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self._normalize_supplier_reference(
            payload.get("transportadora_id", (existing or {}).get("transportadora_id", "")),
            payload.get("transportadora_nome", (existing or {}).get("transportadora_nome", "")),
        )
        if "nota_transporte" in payload and not str(payload.get("nota_transporte", "") or "").strip() and preco_transporte <= 0:
            custo_transporte = 0.0
            transportadora_id = ""
            transportadora_nome = ""
            referencia_transporte = ""
            zona_transporte = ""
        else:
            referencia_transporte = str(payload.get("referencia_transporte", (existing or {}).get("referencia_transporte", "")) or "").strip()
            zona_transporte = str(payload.get("zona_transporte", (existing or {}).get("zona_transporte", "")) or "").strip()
        prazo_entrega_texto = str(
            payload.get(
                "prazo_entrega_texto",
                (existing or {}).get("prazo_entrega_texto", self._quote_default_delivery_text()),
            )
            or self._quote_default_delivery_text()
        ).strip()
        prazo_entrega_data = str(payload.get("prazo_entrega_data", (existing or {}).get("prazo_entrega_data", "")) or "").strip()[:10]
        subtotal_linhas = 0.0
        subtotal_com_desconto = 0.0
        desconto_valor = 0.0
        normalized_discount_groups = {item for item in desconto_grupos if item}
        for line in lines:
            line_total = round(self._parse_float(line.get("total", 0), 0), 2)
            qtd = round(self._parse_float(line.get("qtd", 0), 0), 2)
            preco_unit = round(self._parse_float(line.get("preco_unit", 0), 0), 4)
            preco_unit_incrementado = round(max(0.0, preco_unit * (1.0 + (incremento_preco_perc / 100.0))), 4)
            line_total = round(qtd * preco_unit_incrementado, 2)
            line["preco_unit_incrementado"] = preco_unit_incrementado
            line["total"] = line_total
            subtotal_linhas = round(subtotal_linhas + line_total, 2)
            discount_key = str(line.get("discount_group_key", "") or "").strip()
            apply_discount = desconto_perc > 0 and (
                desconto_modo == "total"
                or not normalized_discount_groups
                or discount_key in normalized_discount_groups
            )
            if apply_discount:
                discounted_unit = round(preco_unit_incrementado * (1.0 - (desconto_perc / 100.0)), 4)
            else:
                discounted_unit = preco_unit_incrementado
            discounted_total = round(qtd * discounted_unit, 2)
            line_discount = round(max(0.0, line_total - discounted_total), 2)
            line["preco_unit_desconto"] = discounted_unit
            line["total_desconto"] = discounted_total
            line["desconto_aplicado"] = line_discount
            subtotal_com_desconto = round(subtotal_com_desconto + discounted_total, 2)
            desconto_valor = round(desconto_valor + line_discount, 2)
        subtotal_bruto = round(subtotal_linhas + preco_transporte, 2)
        subtotal = round(max(0.0, subtotal_com_desconto + preco_transporte), 2)
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
            "desconto_modo": desconto_modo,
            "desconto_grupos": desconto_grupos,
            "incremento_preco_perc": incremento_preco_perc,
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
            "ano": int(str(payload.get("ano", "") or (existing or {}).get("ano", "") or datetime.now().year)),
            "executado_por": str(payload.get("executado_por", "") or (existing or {}).get("executado_por", "") or "").strip(),
            "nota_transporte": (
                str(payload.get("nota_transporte", "") or "").strip()
                if "nota_transporte" in payload
                else str((existing or {}).get("nota_transporte", "") or "").strip()
            ),
            "notas_pdf": str(payload.get("notas_pdf", "") or (existing or {}).get("notas_pdf", "") or "").strip(),
            "prazo_entrega_texto": prazo_entrega_texto,
            "prazo_entrega_data": prazo_entrega_data,
            "nota_cliente": str(payload.get("nota_cliente", "") or (existing or {}).get("nota_cliente", "") or "").strip(),
        }
        # Normalization/repair helpers may have replaced self.data while this
        # quote was being prepared. Publish into the current cache.
        data = self.ensure_data()
        existing = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
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
        upsert = getattr(self.desktop_main, "mysql_upsert_orcamento_com_linhas", None)
        if callable(upsert):
            upsert(data, note)
            if isinstance(self._base_data_snapshot, dict):
                base_rows = self._base_data_snapshot.setdefault("orcamentos", [])
                base_existing = next((row for row in base_rows if str(row.get("numero", "") or "").strip() == numero), None)
                if base_existing is None:
                    base_rows.append(copy.deepcopy(note))
                else:
                    base_existing.update(copy.deepcopy(note))
        else:
            self._save(force=True)
        return self.orc_detail(numero)

    def orc_remove(self, numero: str) -> None:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        before = len(list(data.get("orcamentos", []) or []))
        data["orcamentos"] = [row for row in list(data.get("orcamentos", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        if len(data["orcamentos"]) == before:
            raise ValueError("Orçamento não encontrado.")
        self._save(force=True)
        try:
            self._mysql_delete_orc_nesting_studies(numero)
        except Exception:
            pass

    def orc_set_state(self, numero: str, estado: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
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
            raise ValueError("Orçamento não encontrado.")
        target = Path(path)
        helper = self._orc_render_helper()
        self.orc_actions.render_orc_pdf(helper, str(target), orc)
        return target

    def orc_open_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}.pdf"
        self.orc_render_pdf(numero, target)
        os.startfile(str(target))
        return target

    def orc_print_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}_print.pdf"
        self.orc_render_pdf(numero, target)
        try:
            os.startfile(str(target), "print")
        except Exception:
            os.startfile(str(target))
        return target

    def conjunto_sheet_pdf(self, codigo: str, output_path: str | Path | None = None) -> Path:
        return render_assembly_sheet(render_assembly_sheet_ports(self), codigo, output_path)


    def conjunto_open_sheet_pdf(self, codigo: str) -> Path:
        target = self.conjunto_sheet_pdf(codigo)
        os.startfile(str(target))
        return target

    def orc_render_nesting_study_pdf(self, numero: str, path: str | Path, group_key: str = "") -> Path:
        return render_nesting_study(render_nesting_study_ports(self), numero, path, group_key)

    def orc_open_nesting_study_pdf(self, numero: str, group_key: str = "") -> Path:
        safe_group = "".join(ch if ch.isalnum() else "_" for ch in str(group_key or "").strip()) or "grupo"
        target = Path(tempfile.gettempdir()) / f"lugest_nesting_{str(numero or '').strip()}_{safe_group}.pdf"
        self.orc_render_nesting_study_pdf(numero, target, group_key=group_key)
        os.startfile(str(target))
        return target

    def _quote_line_operations_value(self, line: dict[str, Any] | None = None) -> list[str]:
        row = dict(line or {})
        values: list[Any] = []
        raw_text = str(row.get("operacao", "") or "").strip()
        if raw_text:
            values.append(raw_text)
        for key in ("operacoes_lista", "operacoes_fluxo", "operacoes_detalhe"):
            raw = row.get(key)
            if isinstance(raw, list):
                values.extend(raw)
        for key in ("tempos_operacao", "custos_operacao"):
            raw_map = row.get(key)
            if isinstance(raw_map, dict):
                values.extend(str(name or "").strip() for name in raw_map.keys() if str(name or "").strip())
        return self.quote_parse_operacoes_lista(values)

    def _quote_line_operations_text(self, line: dict[str, Any] | None = None) -> str:
        return self.quote_format_operacoes(self._quote_line_operations_value(line))

    def _quote_line_is_production_ready(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        drawing_path = str(row.get("desenho", "") or "").strip()
        ops = [
            str(self.desktop_main.normalize_operacao_nome(op) or op or "").strip()
            for op in self._quote_line_operations_value(row)
        ]
        ops = [op for op in ops if op and op != "Montagem"]
        material = str(row.get("material", "") or "").strip()
        thickness = str(row.get("espessura", "") or "").strip()
        time_per_piece = self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0)
        detail_ready = bool(
            list(row.get("operacoes_detalhe", []) or [])
            or dict(row.get("tempos_operacao", {}) or {})
            or dict(row.get("custos_operacao", {}) or {})
        )
        has_work = bool(ops or detail_ready or time_per_piece > 0)
        return bool(has_work and (drawing_path or (material and thickness)))

    def _quote_line_is_raw_material(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if str(row.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(row.get("stock_material_id", "") or "").strip():
            return True
        if self._quote_line_looks_stock_material_ref(row.get("ref_externa")):
            if str(row.get("desenho", "") or "").strip():
                return False
            if round(self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 4) > 0:
                return False
            operacao_norm = self.desktop_main.norm_text(str(row.get("operacao", "") or "").strip())
            if operacao_norm in {"", "-", "stockmp", "materia prima", "materia-prima"}:
                return True
        subtype = self.desktop_main.norm_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        if subtype == "stockmp":
            return True
        return False

    def _quote_line_looks_stock_material_ref(self, value: Any) -> bool:
        raw = str(value or "").strip().upper()
        return bool(raw.startswith("MAT") and raw[3:].isdigit())

    def _quote_line_production_route(self, line: dict[str, Any] | None = None) -> str:
        row = dict(line or {})
        line_type = self.desktop_main.normalize_orc_line_type(row.get("tipo_item"))
        if line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return "montagem"
        if line_type == self.desktop_main.ORC_LINE_TYPE_SERVICE:
            return "montagem"
        if not self._quote_line_is_production_ready(row):
            return "conjunto"
        ops = [str(self.desktop_main.normalize_operacao_nome(op) or op or "").strip() for op in self._quote_line_operations_value(row)]
        ops = [op for op in ops if op]
        subtype_norm = self.desktop_main.norm_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        material_norm = self.desktop_main.norm_text(str(row.get("material", "") or "").strip())
        drawing_path = str(row.get("desenho", "") or "").strip()
        if "Corte Laser" in ops and drawing_path:
            return "laser"
        if any(token in subtype_norm for token in ("tubo", "cantoneira", "perfil", "barra", "ferronervurado", "chapa")):
            return "serralharia"
        if any(token in material_norm for token in ("tubo", "cantoneira", "perfil", "barra", "ferro nervurado")):
            return "serralharia"
        if "Serralharia" in ops:
            return "serralharia"
        if "Corte Laser" in ops:
            return "laser"
        if "Montagem" in ops:
            return "montagem"
        return "conjunto"

    def orc_convert_to_order(self, numero: str, nota_cliente: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        note = str(nota_cliente or "").strip()
        orc = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
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
            "data_entrega": str(orc.get("prazo_entrega_data", "") or "").strip()[:10],
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
            "tipo_encomenda": "Cliente",
            "produto_fichas": [],
        }
        sheet_groups: dict[str, dict[str, float]] = {}
        sheet_sources: dict[str, dict[str, Any]] = {}
        for source_line in list(orc.get("linhas", []) or []):
            sheet_code = str(source_line.get("conjunto_codigo", "") or "").strip()
            if not sheet_code:
                continue
            group_key = str(source_line.get("grupo_uuid", "") or "").strip() or sheet_code
            base_qty = self._parse_float(source_line.get("qtd_base", 0), 0)
            line_qty = self._parse_float(source_line.get("qtd", 0), 0)
            group_qty = (line_qty / base_qty) if base_qty > 0 and line_qty > 0 else 1.0
            code_groups = sheet_groups.setdefault(sheet_code, {})
            code_groups[group_key] = max(float(code_groups.get(group_key, 0) or 0), group_qty)
            if sheet_code in sheet_sources:
                continue
            try:
                stored_sheet = dict(self.conjunto_detail(sheet_code) or {})
            except Exception:
                stored_sheet = {}
            sheet_sources[sheet_code] = {
                "codigo": sheet_code,
                "param_codigo": str(stored_sheet.get("param_codigo", "") or source_line.get("conjunto_param_codigo", "") or "").strip(),
                "descricao": str(
                    stored_sheet.get("descricao", "")
                    or source_line.get("conjunto_nome", "")
                    or sheet_code
                ).strip(),
                "notas": str(stored_sheet.get("notas", "") or "").strip(),
                "ficha_tecnica": self._normalize_conjunto_technical_sheet(
                    source_line.get("ficha_tecnica", {}) or stored_sheet.get("ficha_tecnica", {})
                ),
            }
        for sheet_code, snapshot in sheet_sources.items():
            snapshot["quantidade_conjuntos"] = round(
                max(1.0, sum(sheet_groups.get(sheet_code, {}).values())),
                2,
            )
            enc["produto_fichas"].append(snapshot)
        enc_of = str(self._order_of_code(enc, create=True) or "").strip()
        mats: dict[str, dict[str, Any]] = {}
        piece_idx = 1
        total_time = 0.0
        used_refs: set[str] = set()
        montagem_items: list[dict[str, Any]] = []
        for line in list(orc.get("linhas", []) or []):
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            production_route = self._quote_line_production_route(line)
            qtd_line = float(line.get("qtd", 0) or 0)
            tempo_peca = float(line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)) or 0)
            total_time += tempo_peca * max(qtd_line, 0.0)
            if production_route in {"montagem", "conjunto"}:
                montagem_items.append(
                    {
                        "linha_ordem": len(montagem_items) + 1,
                        "tipo_item": line_type,
                        "stock_item_kind": str(line.get("stock_item_kind", "") or "").strip(),
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
                        "material": str(line.get("material", "") or "").strip(),
                        "material_family": str(line.get("material_family", "") or "").strip(),
                        "material_subtype": str(line.get("material_subtype", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "stock_material_id": str(line.get("stock_material_id", "") or "").strip(),
                        "produto_codigo": str(line.get("produto_codigo", "") or "").strip(),
                        "produto_unid": str(line.get("produto_unid", "") or "").strip(),
                        "_product_pending_create": bool(line.get("_product_pending_create", False)),
                        "qtd_planeada": round(qtd_line, 2),
                        "qtd_consumida": 0.0,
                        "preco_unit": round(self._parse_float(line.get("preco_unit", 0), 0), 4),
                        "conjunto_codigo": str(line.get("conjunto_codigo", "") or "").strip(),
                        "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                        "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                        "estado": "Pendente" if production_route == "montagem" else "Componente",
                        "obs": (
                            str(line.get("operacao", "") or production_route).strip()
                            if production_route == "montagem"
                            else "Conjunto sem desenho/operacoes tecnicas. Nao segue para operador."
                        ),
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
                {"espessura": espessura, "tempo_min": 0.0, "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []},
            )
            planning_ops = [op for op in self._quote_line_operations_value(line) if op != "Montagem"]
            if production_route == "serralharia":
                planning_ops = [op for op in planning_ops if op != "Corte Laser"]
                if "Serralharia" not in planning_ops:
                    planning_ops.insert(0, "Serralharia")
            elif production_route == "laser":
                if "Corte Laser" not in planning_ops:
                    planning_ops.insert(0, "Corte Laser")
            esp_bucket = mats[material]["espessuras"][espessura]
            tempos_operacao = esp_bucket.setdefault("tempos_operacao", {})
            maquinas_operacao = esp_bucket.setdefault("maquinas_operacao", {})
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
                        if not str(maquinas_operacao.get(op_name, "") or "").strip():
                            maquinas_operacao[op_name] = self.workcenter_default_resource(op_name, preferred=enc.get("posto_trabalho", ""))
            elif len(planning_ops) == 1:
                op_name = planning_ops[0]
                tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if op_name == "Corte Laser":
                    esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if not str(maquinas_operacao.get(op_name, "") or "").strip():
                    maquinas_operacao[op_name] = self.workcenter_default_resource(op_name, preferred=enc.get("posto_trabalho", ""))
            elif "Corte Laser" in planning_ops:
                tempos_operacao["Corte Laser"] = round(float(tempos_operacao.get("Corte Laser", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if not str(maquinas_operacao.get("Corte Laser", "") or "").strip():
                    maquinas_operacao["Corte Laser"] = self.workcenter_default_resource("Corte Laser", preferred=enc.get("posto_trabalho", ""))
            raw_ref_interna = str(line.get("ref_interna", "") or "").strip()
            if raw_ref_interna and raw_ref_interna not in used_refs:
                ref_interna = raw_ref_interna
            else:
                ref_interna = str(self.desktop_main.next_ref_interna_unique(data, cliente_code, list(used_refs)))
            used_refs.add(ref_interna)
            ops_txt = self._quote_line_operations_text(line)
            peca = {
                "id": f"PEC{piece_idx:05d}",
                "ref_interna": ref_interna,
                "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                "material": material,
                "tipo_material": str(line.get("tipo_material", "") or line.get("material_family", "") or "CHAPA").strip().upper(),
                "subtipo_material": str(line.get("material_subtype", "") or material).strip(),
                "espessura": espessura,
                "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or line.get("profile_size", "") or line.get("tube_section", "") or "").strip(),
                "quantidade_pedida": qtd_line,
                "Operacoes": ops_txt,
                "Observacoes": str(line.get("descricao", "") or "").strip(),
                "conjunto_codigo": str(line.get("conjunto_codigo", "") or "").strip(),
                "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                "desenho": str(line.get("desenho", "") or "").strip(),
                "desenho_pdf": str(line.get("desenho_pdf", "") or "").strip(),
                "desenhos_pdf": [
                    str(item or "").strip()
                    for item in list(line.get("desenhos_pdf", []) or [])
                    if str(item or "").strip()
                ],
                "ficheiros": [
                    str(item or "").strip()
                    for item in [
                        line.get("desenho", ""),
                        line.get("desenho_pdf", ""),
                        *list(line.get("desenhos_pdf", []) or []),
                        *list(line.get("ficheiros", []) or []),
                    ]
                    if str(item or "").strip()
                ],
                "tempo_peca_min": tempo_peca,
                "tempos_operacao": dict(line.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(line.get("custos_operacao", {}) or {}),
                "operacoes_detalhe": [dict(item or {}) for item in list(line.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "of": enc_of,
                "opp": f"OPP-{enc_of.split('-', 1)[1]}-{piece_idx:02d}" if enc_of.startswith("OF-") and "-" in enc_of else self.desktop_main.next_opp_numero(data),
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

    def _quote_purchase_need_key(self, kind: str, line: dict[str, Any]) -> str:
        if kind == "product":
            code = str(line.get("produto_codigo", "") or "").strip()
            if code:
                return f"product:{code}"
            return "product:new:" + self.desktop_main.norm_text(str(line.get("descricao", "") or "").strip())
        ref = str(line.get("stock_material_id", "") or "").strip()
        if ref:
            return f"material:{ref}"
        parts = [
            str(line.get("material", "") or "").strip(),
            str(line.get("espessura", "") or "").strip(),
            str(line.get("material_subtype", "") or line.get("calc_mode", "") or "").strip(),
            str(line.get("descricao", "") or "").strip(),
        ]
        return "material:new:" + "|".join(self.desktop_main.norm_text(part) for part in parts if part)

    def orc_purchase_needs(self, numero: str = "", lines: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        data = self.ensure_data()
        numero_txt = str(numero or "").strip()
        quote_lines = list(lines or [])
        if not quote_lines:
            orc = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero_txt), None)
            if orc is None:
                raise ValueError("Orcamento nao encontrado.")
            quote_lines = list(orc.get("linhas", []) or [])
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(data.get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        grouped: dict[str, dict[str, Any]] = {}
        for raw in quote_lines:
            line = dict(raw or {})
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            if line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                qty = max(0.0, self._parse_float(line.get("qtd", 0), 0))
                if qty <= 1e-9:
                    continue
                code = str(line.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                available = max(0.0, self._parse_float((product or {}).get("qty", 0), 0)) if product is not None else 0.0
                missing = qty if product is None else max(0.0, qty - available)
                if missing <= 1e-9:
                    continue
                key = self._quote_purchase_need_key("product", line)
                entry = grouped.setdefault(
                    key,
                    {
                        "kind": "product",
                        "ref": code,
                        "descricao": str(line.get("descricao", "") or (product or {}).get("descricao", "") or "").strip(),
                        "unid": str(line.get("produto_unid", "") or (product or {}).get("unid", "") or "UN").strip() or "UN",
                        "qtd": 0.0,
                        "qtd_disponivel": available,
                        "preco": round(self._parse_float((product or {}).get("p_compra", line.get("preco_unit", 0)), 0), 4),
                        "_product_pending_create": product is None or bool(line.get("_product_pending_create", False)),
                    },
                )
                entry["qtd"] = round(self._parse_float(entry.get("qtd", 0), 0) + missing, 2)
                continue
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE or not self._quote_line_is_raw_material(line):
                continue
            qty = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            if qty <= 1e-9:
                continue
            stock_id = str(line.get("stock_material_id", "") or "").strip()
            material_record = self.material_by_id(stock_id) if stock_id else None
            available = 0.0
            if isinstance(material_record, dict):
                available = max(
                    0.0,
                    self._parse_float(material_record.get("quantidade", 0), 0)
                    - self._parse_float(material_record.get("reservado", 0), 0),
                )
            missing = qty if material_record is None else max(0.0, qty - available)
            if missing <= 1e-9:
                continue
            formato = str(line.get("material_subtype", "") or line.get("calc_mode", "") or (material_record or {}).get("formato", "") or "Chapa").strip()
            if formato == "Stock MP":
                formato = str((material_record or {}).get("formato", "") or self.desktop_main.detect_materia_formato(material_record or {}) or "Chapa").strip()
            price = self._parse_float(line.get("price_base_value", 0), 0)
            if price <= 0 and isinstance(material_record, dict):
                price = self._parse_float(material_record.get("p_compra", material_record.get("preco_unid", 0)), 0)
            key = self._quote_purchase_need_key("material", line)
            entry = grouped.setdefault(
                key,
                {
                    "kind": "material",
                    "ref": stock_id,
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
                    "unid": "UN",
                    "qtd": 0.0,
                    "qtd_disponivel": available,
                    "preco": round(price, 4),
                    "material": str(line.get("material", "") or (material_record or {}).get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or (material_record or {}).get("espessura", "") or "").strip(),
                    "formato": formato or "Chapa",
                    "comprimento": round(self._parse_float(line.get("length_mm", (material_record or {}).get("comprimento", 0)), 0), 3),
                    "largura": round(self._parse_float(line.get("width_mm", (material_record or {}).get("largura", 0)), 0), 3),
                    "diametro": round(self._parse_float(line.get("diameter_mm", (material_record or {}).get("diametro", 0)), 0), 3),
                    "metros": round(self._parse_float(line.get("meters_per_unit", (material_record or {}).get("metros", 0)), 0), 4),
                    "kg_m": round(self._parse_float(line.get("kg_per_m", (material_record or {}).get("kg_m", 0)), 0), 4),
                    "peso_unid": round(self._parse_float(line.get("stock_metric_value", (material_record or {}).get("peso_unid", 0)), 0), 4),
                    "_material_pending_create": material_record is None,
                    "_material_manual": material_record is None,
                },
            )
            entry["qtd"] = round(self._parse_float(entry.get("qtd", 0), 0) + missing, 2)
        rows = [row for row in grouped.values() if self._parse_float(row.get("qtd", 0), 0) > 0]
        rows.sort(key=lambda row: (str(row.get("kind", "")), str(row.get("ref", "") or row.get("descricao", ""))))
        return rows

    def orc_create_purchase_quote(self, numero: str, lines: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        needs = self.orc_purchase_needs(numero_txt, lines)
        if not needs:
            raise ValueError("Nao existem necessidades de compra nas linhas do orcamento.")
        note = self.ne_save(
            {
                "fornecedor": "",
                "fornecedor_id": "",
                "contacto": "",
                "obs": f"Pedido de cotacao gerado a partir do orcamento {numero_txt}".strip(),
                "lines": [
                    (
                        {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip() or str(need.get("material", "") or "").strip(),
                            "origem": "Materia-prima",
                            "qtd": round(self._parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "material": str(need.get("material", "") or "").strip(),
                            "espessura": str(need.get("espessura", "") or "").strip(),
                            "dimensao": str(need.get("dimensao", "") or "").strip(),
                            "dimensoes": str(need.get("dimensao", "") or "").strip(),
                            "formato": str(need.get("formato", "") or "Chapa").strip() or "Chapa",
                            "comprimento": self._parse_float(need.get("comprimento", 0), 0),
                            "largura": self._parse_float(need.get("largura", 0), 0),
                            "diametro": self._parse_float(need.get("diametro", 0), 0),
                            "metros": self._parse_float(need.get("metros", 0), 0),
                            "kg_m": self._parse_float(need.get("kg_m", 0), 0),
                            "peso_unid": self._parse_float(need.get("peso_unid", 0), 0),
                            "_material_pending_create": bool(need.get("_material_pending_create", False)),
                            "_material_manual": bool(need.get("_material_manual", False)),
                        }
                        if str(need.get("kind", "") or "") == "material"
                        else {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip(),
                            "origem": "Produto",
                            "qtd": round(self._parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "_product_pending_create": bool(need.get("_product_pending_create", False)),
                        }
                    )
                    for need in needs
                ],
            }
        )
        note_number = str(note.get("numero", "") or "").strip()
        return {"numero": note_number, "line_count": len(list(note.get("linhas", []) or [])), "needs": needs, "detail": self.ne_detail(note_number)}

    def orc_suggest_notes(self, payload: dict[str, Any]) -> str:
        helper = self._orc_render_helper()
        lines = self.orc_actions._build_orc_notes_lines(helper, payload)
        return "\n".join([str(line or "").strip() for line in lines if str(line or "").strip()])

