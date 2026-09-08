"""Assembly item preparation and live pricing with explicit catalog capabilities.

No shared ERP snapshot is exposed here. Input aggregates are copied before
normalization or repricing so callers retain ownership of their draft.
"""
from __future__ import annotations
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class AssemblyRules:
    ORC_LINE_TYPE_PIECE: str
    ORC_LINE_TYPE_PRODUCT: str
    parse_float: Callable
    normalize_orc_line_type: Callable
    product_lookup: Callable
    produto_preco_venda: Callable
    produto_preco_unitario: Callable
    orc_line_is_product: Callable
    orc_line_is_piece: Callable
    norm_text: Callable
    detect_materia_formato: Callable
    material_by_id: Callable
    material_candidates: Callable
    material_price_preview: Callable
    quote_source: Callable
    now_iso: Callable

def normalize_item(rules: AssemblyRules, payload: dict[str, Any]) -> dict[str, Any]:
    payload = deepcopy(payload)
    item_type = rules.normalize_orc_line_type(payload.get("tipo_item"))
    quantity = round(rules.parse_float(payload.get("qtd", 0), 0), 2)
    if quantity <= 0:
        raise ValueError("Quantidade invalida no conjunto.")
    stock_item_kind = str(payload.get("stock_item_kind", "") or "").strip()
    if item_type == rules.ORC_LINE_TYPE_PIECE and (
        stock_item_kind == "raw_material" or str(payload.get("stock_material_id", "") or "").strip()
    ):
        stock_item_kind = "raw_material"
    elif item_type == rules.ORC_LINE_TYPE_PRODUCT:
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
        "tempo_peca_min": round(rules.parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
        "preco_unit": round(rules.parse_float(payload.get("preco_unit", 0), 0), 4),
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
        "weight_total": round(rules.parse_float(payload.get("weight_total", 0), 0), 3),
        "total_cost": round(rules.parse_float(payload.get("total_cost", 0), 0), 2),
        "quantity_units": round(rules.parse_float(payload.get("quantity_units", quantity), quantity), 2),
        "price_per_kg": round(rules.parse_float(payload.get("price_per_kg", 0), 0), 4),
        "price_base_value": round(rules.parse_float(payload.get("price_base_value", 0), 0), 4),
        "price_markup_pct": round(rules.parse_float(payload.get("price_markup_pct", 0), 0), 2),
        "stock_metric_value": round(rules.parse_float(payload.get("stock_metric_value", 0), 0), 4),
        "meters_per_unit": round(rules.parse_float(payload.get("meters_per_unit", 0), 0), 3),
        "kg_per_m": round(rules.parse_float(payload.get("kg_per_m", 0), 0), 4),
        "length_mm": round(rules.parse_float(payload.get("length_mm", 0), 0), 1),
        "width_mm": round(rules.parse_float(payload.get("width_mm", 0), 0), 1),
        "thickness_mm": round(rules.parse_float(payload.get("thickness_mm", 0), 0), 2),
        "density": round(rules.parse_float(payload.get("density", 0), 0), 1),
        "diameter_mm": round(rules.parse_float(payload.get("diameter_mm", 0), 0), 1),
        "manual_unit_price": round(rules.parse_float(payload.get("manual_unit_price", 0), 0), 4),
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
        "laser_base_tempo_unit": round(rules.parse_float(payload.get("laser_base_tempo_unit", 0), 0), 4),
        "laser_base_preco_unit": round(rules.parse_float(payload.get("laser_base_preco_unit", 0), 0), 4),
        "source_quote_number": str(payload.get("source_quote_number", "") or "").strip(),
        "source_ref_externa": str(payload.get("source_ref_externa", "") or "").strip(),
        "pricing_source": str(payload.get("pricing_source", "") or "").strip(),
        "pricing_source_ref": str(payload.get("pricing_source_ref", "") or "").strip(),
        "preco_anterior": round(rules.parse_float(payload.get("preco_anterior", 0), 0), 4),
        "preco_atualizado_em": str(payload.get("preco_atualizado_em", "") or "").strip(),
    }
    if item_type == rules.ORC_LINE_TYPE_PIECE:
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
    elif item_type == rules.ORC_LINE_TYPE_PRODUCT:
        product = rules.product_lookup(item["produto_codigo"])
        if product is None and not item["descricao"]:
            raise ValueError("Descricao obrigatoria no produto.")
        item["_product_pending_create"] = product is None
        item["descricao"] = item["descricao"] or str((product or {}).get("descricao", "") or "").strip()
        item["produto_unid"] = item["produto_unid"] or str((product or {}).get("unid", "") or "UN").strip()
        if product is not None and item["preco_unit"] <= 0:
            item["preco_unit"] = round(rules.parse_float(rules.produto_preco_venda(product), 0), 4)
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


def price_item(rules: AssemblyRules, raw_item: dict[str, Any], conjunto_codigo: str) -> tuple[dict[str, Any], bool]:
    item = normalize_item(rules, dict(raw_item or {}))
    old_price = round(rules.parse_float(item.get("preco_unit", 0), 0), 4)
    live_price = old_price
    source_type = "manual"
    source_label = "Valor manual"
    source_ref = ""
    linked = False

    if rules.orc_line_is_product(item):
        product = rules.product_lookup(item.get("produto_codigo", ""))
        if product is not None:
            live_price = round(rules.parse_float(rules.produto_preco_unitario(product), 0), 4)
            source_type = "product_stock"
            source_label = "Stock produtos"
            source_ref = str(product.get("codigo", "") or "").strip()
            linked = True
    else:
        stock_id = str(item.get("stock_material_id", "") or "").strip()
        material_record = rules.material_by_id(stock_id) if stock_id else None
        if material_record is None and item.get("calc_mode") and rules.parse_float(item.get("stock_metric_value", 0), 0) > 0:
            wanted_mode = rules.norm_text(str(item.get("calc_mode", "") or ""))
            base_value = rules.parse_float(item.get("price_base_value", 0), 0)
            candidates = []
            for candidate in rules.material_candidates():
                candidate_mode = rules.norm_text(
                    str(candidate.get("formato", "") or rules.detect_materia_formato(candidate) or "")
                )
                if wanted_mode and candidate_mode != wanted_mode:
                    continue
                delta = abs(rules.parse_float(candidate.get("p_compra", 0), 0) - base_value)
                candidates.append((delta, candidate))
            if candidates:
                candidates.sort(key=lambda pair: pair[0])
                if candidates[0][0] <= 0.0002:
                    material_record = candidates[0][1]
                    stock_id = str(material_record.get("id", "") or "").strip()
                    item["stock_material_id"] = stock_id
        if material_record is not None:
            preview = rules.material_price_preview(material_record)
            metric = rules.parse_float(item.get("stock_metric_value", 0), 0)
            base_label = str(item.get("price_base_label", "") or "").strip().lower()
            if metric > 0 and base_label:
                current_base = rules.parse_float(material_record.get("p_compra", 0), 0)
                live_price = round(current_base * metric, 4)
                item["price_base_value"] = round(current_base, 4)
            else:
                live_price = round(rules.parse_float(preview.get("preco_unid", old_price), old_price), 4)
            source_type = "material_stock"
            source_label = "Stock materia-prima"
            source_ref = stock_id
            linked = True
        elif rules.orc_line_is_piece(item):
            quote_line, quote_number = rules.quote_source(item, conjunto_codigo)
            if quote_line is not None:
                live_price = round(rules.parse_float(quote_line.get("preco_unit", old_price), old_price), 4)
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
        item["preco_atualizado_em"] = rules.now_iso()
    item["pricing_source"] = source_type
    item["pricing_source_label"] = source_label
    item["pricing_source_ref"] = source_ref
    item["pricing_linked"] = linked
    return item, changed or any(item.get(key) != raw_item.get(key) for key in (
        "stock_material_id", "source_quote_number", "source_ref_externa", "pricing_source", "pricing_source_ref"
    ))


def refresh_model(rules: AssemblyRules, model: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    model = deepcopy(model)
    code = str(model.get("codigo", "") or "").strip()
    refreshed: list[dict[str, Any]] = []
    changed = False
    for index, raw_item in enumerate(list(model.get("itens", []) or []), start=1):
        item, item_changed = price_item(rules, dict(raw_item or {}), code)
        item["linha_ordem"] = index
        refreshed.append(item)
        changed = changed or item_changed
    total_cost = round(sum(rules.parse_float(item.get("qtd", 0), 0) * rules.parse_float(item.get("preco_unit", 0), 0) for item in refreshed), 2)
    margin = rules.parse_float(model.get("margem_perc", 0), 0)
    total_final = round(total_cost * (1.0 + margin / 100.0), 2)
    if abs(total_cost - rules.parse_float(model.get("total_custo", 0), 0)) > 0.005:
        changed = True
    if abs(total_final - rules.parse_float(model.get("total_final", 0), 0)) > 0.005:
        changed = True
    model["itens"] = refreshed
    model["total_custo"] = total_cost
    model["total_final"] = total_final
    if changed:
        model["precos_atualizados_em"] = rules.now_iso()
    return model, changed


def expand_model(rules: AssemblyRules, detail: dict[str, Any], quantity: Any, group_uuid: str) -> list[dict[str, Any]]:
    detail = deepcopy(detail)
    multiplier = round(rules.parse_float(quantity, 0), 2)
    if multiplier <= 0:
        raise ValueError("Quantidade do conjunto invalida.")
    rows: list[dict[str, Any]] = []
    for item in list(detail.get("itens", []) or []):
        line = {
            "tipo_item": rules.normalize_orc_line_type(item.get("tipo_item")),
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
            "qtd_base": round(rules.parse_float(item.get("qtd", 0), 0), 2),
            "tempo_peca_min": round(rules.parse_float(item.get("tempo_peca_min", 0), 0), 2),
            "qtd": round(rules.parse_float(item.get("qtd", 0), 0) * multiplier, 2),
            "preco_unit": round(rules.parse_float(item.get("preco_unit", 0), 0), 4),
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
            "price_base_value": round(rules.parse_float(item.get("price_base_value", 0), 0), 4),
            "price_base_label": str(item.get("price_base_label", "") or "").strip(),
            "stock_metric_value": round(rules.parse_float(item.get("stock_metric_value", 0), 0), 4),
            "meters_per_unit": round(rules.parse_float(item.get("meters_per_unit", 0), 0), 3),
            "kg_per_m": round(rules.parse_float(item.get("kg_per_m", 0), 0), 4),
            "length_mm": round(rules.parse_float(item.get("length_mm", 0), 0), 1),
            "width_mm": round(rules.parse_float(item.get("width_mm", 0), 0), 1),
            "thickness_mm": round(rules.parse_float(item.get("thickness_mm", 0), 0), 2),
            "diameter_mm": round(rules.parse_float(item.get("diameter_mm", 0), 0), 1),
            "profile_section": str(item.get("profile_section", "") or "").strip(),
            "profile_size": str(item.get("profile_size", "") or "").strip(),
            "tube_section": str(item.get("tube_section", "") or "").strip(),
            "quality": str(item.get("quality", "") or "").strip(),
            "calc_mode": str(item.get("calc_mode", "") or "").strip(),
        }
        if rules.orc_line_is_product(line) and not line["ref_externa"]:
            line["ref_externa"] = line["produto_codigo"]
        rows.append(line)
    return rows


def technical_sheet(raw: Any) -> dict[str, str]:
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

