"""Quote read models, independent of the desktop snapshot and widgets."""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class QuoteReadRepository(Protocol):
    def summaries(self) -> list[dict]: ...
    def get(self, number: str) -> dict | None: ...

@dataclass(frozen=True)
class QuoteQueryRules:
    _fmt: Callable[..., Any]
    _normalize_conjunto_technical_sheet: Callable[..., Any]
    _normalize_orc_client: Callable[..., Any]
    _normalize_quote_discount_groups: Callable[..., Any]
    _normalize_quote_discount_mode: Callable[..., Any]
    _normalize_workcenter_value: Callable[..., Any]
    _orc_number_sort_key: Callable[..., Any]
    _parse_float: Callable[..., Any]
    _quote_collect_non_laser_map: Callable[..., Any]
    _quote_default_delivery_text: Callable[..., Any]
    _quote_line_operation_snapshot: Callable[..., Any]
    _quote_standard_iva_perc: Callable[..., Any]
    current_year: Callable[..., Any]
    extract_year: Callable[..., Any]
    norm_text: Callable[..., Any]
    normalize_orc_line_type: Callable[..., Any]
    orc_line_is_piece: Callable[..., Any]

class QuoteQueries:
    def __init__(self, repository: QuoteReadRepository, rules: QuoteQueryRules):
        self.repository = repository
        self.rules = rules

    def rows(self, filter_text: str = "", state_filter: str = "Ativas", year: str = "Todos") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        year_raw = str(year or "Todos").strip()
        rows: list[dict[str, Any]] = []
        for raw in self.repository.summaries():
            if not isinstance(raw, dict):
                continue
            orc = dict(raw)
            client = self.rules._normalize_orc_client(orc.get("cliente", {}))
            estado = str(orc.get("estado", "") or "").strip() or "Em edicao"
            estado_norm = self.rules.norm_text(estado)
            row_year = str(self.rules.extract_year(orc.get("data", ""), orc.get("numero", ""), orc.get("ano")) or "").strip()
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
                "total": round(self.rules._parse_float(orc.get("total", 0), 0), 2),
                "data": str(orc.get("data", "") or "").strip()[:10],
                "linhas": int(orc.get("line_count", 0)),
                "ano": row_year,
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: self.rules._orc_number_sort_key(str(item.get("numero", "") or "")), reverse=True)
        return rows

    def years(self) -> list[str]:
        current_year = str(self.rules.current_year())
        years = {current_year}
        for row in self.repository.summaries():
            if not isinstance(row, dict):
                continue
            year = str(self.rules.extract_year(row.get("data", ""), row.get("numero", ""), row.get("ano")) or "").strip()
            if year:
                years.add(year)
        return sorted(years, key=lambda value: int(value) if value.isdigit() else 0, reverse=True)

    def detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = self.repository.get(numero)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
        client = self.rules._normalize_orc_client(orc.get("cliente", {}))
        lines: list[dict[str, Any]] = []
        for row in list(orc.get("linhas", []) or []):
            snapshot = self.rules._quote_line_operation_snapshot(row, quote_number=numero, quote_state=str(orc.get("estado", "") or "").strip())
            raw_operacao = str(row.get("operacao", "") or "").strip()
            current_time = round(self.rules._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 4)
            current_price = round(self.rules._parse_float(row.get("preco_unit", 0), 0), 4)
            derived_laser_base = (
                self.rules.orc_line_is_piece(row)
                and bool(str(row.get("desenho", "") or "").strip())
                and "corte laser" in self.rules.norm_text(raw_operacao)
                and (
                    current_time > 0
                    or current_price > 0
                )
            )
            laser_base_active = bool(row.get("laser_base_active", False) or derived_laser_base)
            laser_base_tempo = round(
                self.rules._parse_float(
                    row.get(
                        "laser_base_tempo_unit",
                        current_time if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            laser_base_preco = round(
                self.rules._parse_float(
                    row.get(
                        "laser_base_preco_unit",
                        current_price if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            display_extra_time_map = self.rules._quote_collect_non_laser_map(dict(snapshot.get("tempos_operacao", {}) or {}), digits=4)
            display_extra_price_map = self.rules._quote_collect_non_laser_map(dict(snapshot.get("custos_operacao", {}) or {}), digits=4)
            display_extra_time = round(sum(display_extra_time_map.values()), 4)
            display_extra_price = round(sum(display_extra_price_map.values()), 4)
            repair_extra_time_map = self.rules._quote_collect_non_laser_map(
                dict(row.get("tempos_operacao", {}) or {}),
                dict(snapshot.get("tempos_operacao", {}) or {}),
                digits=4,
            )
            repair_extra_price_map = self.rules._quote_collect_non_laser_map(
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
            line_qty = round(self.rules._parse_float(row.get("qtd", 0), 0), 2)
            material_supplied_by_client = bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False))
            lines.append(
                {
                    "tipo_item": self.rules.normalize_orc_line_type(row.get("tipo_item")),
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
                    "espessura": self.rules._fmt(row.get("espessura", "")),
                    "operacao": str(row.get("operacao", "") or "").strip(),
                    "produto_codigo": str(row.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(row.get("produto_unid", "") or "").strip(),
                    "_product_pending_create": bool(row.get("_product_pending_create", False)),
                    "conjunto_codigo": str(row.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(row.get("conjunto_nome", "") or "").strip(),
                    "conjunto_param_codigo": str(row.get("conjunto_param_codigo", "") or "").strip(),
                    "grupo_uuid": str(row.get("grupo_uuid", "") or "").strip(),
                    "ficha_tecnica": self.rules._normalize_conjunto_technical_sheet(row.get("ficha_tecnica", {})),
                    "qtd_base": round(self.rules._parse_float(row.get("qtd_base", row.get("qtd", 0)), 0), 2),
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
                    "price_per_kg": round(self.rules._parse_float(row.get("price_per_kg", 0), 0), 4),
                    "price_base_value": round(self.rules._parse_float(row.get("price_base_value", 0), 0), 4),
                    "price_base_label": str(row.get("price_base_label", "") or "").strip(),
                    "price_markup_pct": round(self.rules._parse_float(row.get("price_markup_pct", 0), 0), 2),
                    "stock_metric_value": round(self.rules._parse_float(row.get("stock_metric_value", 0), 0), 4),
                    "meters_per_unit": round(self.rules._parse_float(row.get("meters_per_unit", 0), 0), 3),
                    "kg_per_m": round(self.rules._parse_float(row.get("kg_per_m", 0), 0), 4),
                    "length_mm": round(self.rules._parse_float(row.get("length_mm", 0), 0), 1),
                    "width_mm": round(self.rules._parse_float(row.get("width_mm", 0), 0), 1),
                    "thickness_mm": round(self.rules._parse_float(row.get("thickness_mm", 0), 0), 2),
                    "diameter_mm": round(self.rules._parse_float(row.get("diameter_mm", 0), 0), 1),
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
            "posto_trabalho": self.rules._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "iva_perc": self.rules._quote_standard_iva_perc(),
            "desconto_perc": round(self.rules._parse_float(orc.get("desconto_perc", 0), 0), 2),
            "desconto_modo": self.rules._normalize_quote_discount_mode(orc.get("desconto_modo", "total")),
            "desconto_grupos": self.rules._normalize_quote_discount_groups(orc.get("desconto_grupos", [])),
            "incremento_preco_perc": round(self.rules._parse_float(orc.get("incremento_preco_perc", 0), 0), 2),
            "desconto_valor": round(self.rules._parse_float(orc.get("desconto_valor", 0), 0), 2),
            "subtotal_linhas": round(self.rules._parse_float(orc.get("subtotal_linhas", orc.get("subtotal_bruto", 0)), 0), 2),
            "subtotal_bruto": round(self.rules._parse_float(orc.get("subtotal_bruto", 0), 0), 2),
            "preco_transporte": round(self.rules._parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self.rules._parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self.rules._parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self.rules._parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self.rules._parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "subtotal": round(self.rules._parse_float(orc.get("subtotal", 0), 0), 2),
            "total": round(self.rules._parse_float(orc.get("total", 0), 0), 2),
            "numero_encomenda": str(orc.get("numero_encomenda", "") or "").strip(),
            "executado_por": str(orc.get("executado_por", "") or "").strip(),
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "notas_pdf": str(orc.get("notas_pdf", "") or "").strip(),
            "prazo_entrega_texto": str(orc.get("prazo_entrega_texto", "") or self.rules._quote_default_delivery_text()).strip(),
            "prazo_entrega_data": str(orc.get("prazo_entrega_data", "") or "").strip()[:10],
            "nota_cliente": str(orc.get("nota_cliente", "") or "").strip(),
            "nesting_bridge": dict(orc.get("latest_nesting_bridge", {}) or {}),
            "nesting_group_key": str(orc.get("latest_nesting_group_key", "") or "").strip(),
            "nesting_updated_at": str(orc.get("latest_nesting_updated_at", "") or "").strip(),
            "linhas": lines,
        }
