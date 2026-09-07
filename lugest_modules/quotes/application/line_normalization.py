from __future__ import annotations


from typing import Any



from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class LineNormalizationPorts:
    ORC_LINE_TYPE_PIECE: str
    ORC_LINE_TYPE_PRODUCT: str
    _normalize_conjunto_technical_sheet: Callable[..., Any]
    _parse_float: Callable[..., Any]
    _product_lookup: Callable[..., Any]
    _quote_collect_non_laser_map: Callable[..., Any]
    _quote_line_looks_stock_material_ref: Callable[..., Any]
    _quote_line_operation_snapshot: Callable[..., Any]
    _quote_line_operations_text: Callable[..., Any]
    _quote_line_operations_value: Callable[..., Any]
    norm_text: Callable[..., Any]
    normalize_operacao_nome: Callable[..., Any]
    normalize_orc_line_type: Callable[..., Any]
    produto_preco_venda: Callable[..., Any]

def normalize_line(ports: LineNormalizationPorts, payload: dict[str, Any]) -> dict[str, Any]:
    line_type = ports.normalize_orc_line_type(payload.get("tipo_item"))
    stock_item_kind = str(payload.get("stock_item_kind", "") or "").strip()
    raw_by_stock_ref = ports._quote_line_looks_stock_material_ref(payload.get("ref_externa"))
    if raw_by_stock_ref:
        operacao_norm = ports.norm_text(ports._quote_line_operations_text(payload))
        raw_by_stock_ref = bool(
            line_type == ports.ORC_LINE_TYPE_PIECE
            and not str(payload.get("desenho", "") or "").strip()
            and round(ports._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 4) <= 0
            and operacao_norm in {"", "-", "stockmp", "materia prima", "materia-prima"}
        )
    if line_type == ports.ORC_LINE_TYPE_PIECE and (
        stock_item_kind == "raw_material" or str(payload.get("stock_material_id", "") or "").strip() or raw_by_stock_ref
    ):
        stock_item_kind = "raw_material"
    elif line_type == ports.ORC_LINE_TYPE_PRODUCT:
        stock_item_kind = "product"
    else:
        stock_item_kind = ""
    line = {
        "tipo_item": line_type,
        "stock_item_kind": stock_item_kind,
        "ref_interna": str(payload.get("ref_interna", "") or "").strip(),
        "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
        "descricao": str(payload.get("descricao", "") or "").strip(),
        "dimensao": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
        "material": str(payload.get("material", "") or "").strip(),
        "material_family": str(payload.get("material_family", "") or "").strip(),
        "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
        "material_supplied_by_client": bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False)),
        "material_fornecido_cliente": bool(payload.get("material_fornecido_cliente", False) or payload.get("material_supplied_by_client", False)),
        "material_cost_included": (
            bool(payload.get("material_cost_included", True))
            if "material_cost_included" in payload
            else not bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False))
        ),
        "espessura": str(payload.get("espessura", "") or "").strip(),
        "operacao": str(payload.get("operacao", "") or "").strip(),
        "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
        "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
        "conjunto_codigo": str(payload.get("conjunto_codigo", "") or "").strip(),
        "conjunto_nome": str(payload.get("conjunto_nome", "") or "").strip(),
        "conjunto_param_codigo": str(payload.get("conjunto_param_codigo", "") or "").strip(),
        "grupo_uuid": str(payload.get("grupo_uuid", "") or "").strip(),
        "ficha_tecnica": ports._normalize_conjunto_technical_sheet(payload.get("ficha_tecnica", {})),
        "qtd_base": round(ports._parse_float(payload.get("qtd_base", payload.get("qtd", 0)), 0), 2),
        "tempo_peca_min": round(ports._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
        "qtd": round(ports._parse_float(payload.get("qtd", 0), 0), 2),
        "preco_unit": round(ports._parse_float(payload.get("preco_unit", 0), 0), 4),
        "desenho": str(payload.get("desenho", "") or "").strip(),
        "desenho_pdf": str(payload.get("desenho_pdf", "") or "").strip(),
        "desenhos_pdf": [
            str(item or "").strip()
            for item in list(payload.get("desenhos_pdf", []) or [])
            if str(item or "").strip()
        ],
        "ficheiros": [
            str(item or "").strip()
            for item in list(payload.get("ficheiros", []) or [])
            if str(item or "").strip()
        ],
        "price_per_kg": round(ports._parse_float(payload.get("price_per_kg", 0), 0), 4),
        "price_base_value": round(ports._parse_float(payload.get("price_base_value", 0), 0), 4),
        "price_markup_pct": round(ports._parse_float(payload.get("price_markup_pct", 0), 0), 2),
        "stock_metric_value": round(ports._parse_float(payload.get("stock_metric_value", 0), 0), 4),
        "price_base_label": str(payload.get("price_base_label", "") or "").strip(),
        "kg_per_m": round(ports._parse_float(payload.get("kg_per_m", 0), 0), 4),
        "meters_per_unit": round(ports._parse_float(payload.get("meters_per_unit", 0), 0), 3),
        "length_mm": round(ports._parse_float(payload.get("length_mm", 0), 0), 1),
        "width_mm": round(ports._parse_float(payload.get("width_mm", 0), 0), 1),
        "thickness_mm": round(ports._parse_float(payload.get("thickness_mm", 0), 0), 2),
        "diameter_mm": round(ports._parse_float(payload.get("diameter_mm", 0), 0), 1),
        "profile_section": str(payload.get("profile_section", "") or "").strip(),
        "profile_size": str(payload.get("profile_size", "") or "").strip(),
        "tube_section": str(payload.get("tube_section", "") or "").strip(),
        "quality": str(payload.get("quality", "") or "").strip(),
        "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
        "laser_base_active": bool(payload.get("laser_base_active", False)),
        "laser_base_tempo_unit": round(ports._parse_float(payload.get("laser_base_tempo_unit", payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0))), 0), 4),
        "laser_base_preco_unit": round(ports._parse_float(payload.get("laser_base_preco_unit", payload.get("preco_unit", 0)), 0), 4),
        "machine": str(payload.get("machine", "") or "").strip(),
        "laser_machine": str(payload.get("laser_machine", payload.get("machine", "")) or "").strip(),
        "commercial_profile": str(payload.get("commercial_profile", "") or "").strip(),
        "gas": str(payload.get("gas", "") or "").strip(),
        "laser_snapshot": dict(payload.get("laser_snapshot", {}) or {}),
        "laser_source_mode": str(payload.get("laser_source_mode", "") or "").strip(),
        "laser_batch_id": str(payload.get("laser_batch_id", "") or "").strip(),
        "discount_group_key": str(payload.get("discount_group_key", "") or "").strip(),
    }
    if stock_item_kind == "raw_material" and not line["stock_material_id"] and ports._quote_line_looks_stock_material_ref(line["ref_externa"]):
        line["stock_material_id"] = line["ref_externa"]
    if line["material_supplied_by_client"] or line["material_fornecido_cliente"]:
        line["material_supplied_by_client"] = True
        line["material_fornecido_cliente"] = True
        line["material_cost_included"] = False
    if line["qtd"] <= 0:
        raise ValueError("Quantidade invalida na linha.")
    if line_type == ports.ORC_LINE_TYPE_PIECE:
        if not line["descricao"]:
            raise ValueError("Descricao obrigatoria na linha.")
        if not line["material"] or not line["espessura"]:
            raise ValueError("Material e espessura sao obrigatorios na linha.")
        if not line["material_family"]:
            line["material_family"] = line["material"]
        line["produto_codigo"] = ""
        line["produto_unid"] = ""
        if stock_item_kind == "raw_material":
            has_raw_operations = bool(
                ports._quote_line_operations_value(payload)
                or list(payload.get("operacoes_lista", []) or [])
                or list(payload.get("operacoes_detalhe", []) or [])
                or dict(payload.get("tempos_operacao", {}) or {})
                or dict(payload.get("custos_operacao", {}) or {})
                or ports._parse_float(line.get("tempo_peca_min", 0), 0) > 0
            )
            line["ref_interna"] = ""
            line["produto_codigo"] = ""
            line["produto_unid"] = ""
            if not has_raw_operations:
                line["desenho"] = ""
                line["desenho_pdf"] = ""
                line["desenhos_pdf"] = []
                line["ficheiros"] = []
                line["laser_base_active"] = False
                line["laser_base_tempo_unit"] = 0.0
                line["laser_base_preco_unit"] = 0.0
                line["operacao"] = ""
                line["tempo_peca_min"] = 0.0
                line["operacoes_lista"] = []
                line["operacoes_fluxo"] = []
                line["operacoes_detalhe"] = []
                line["tempos_operacao"] = {}
                line["custos_operacao"] = {}
                line["quote_cost_snapshot"] = {}
    elif line_type == ports.ORC_LINE_TYPE_PRODUCT:
        product = ports._product_lookup(line["produto_codigo"])
        if product is None and not line["descricao"]:
            raise ValueError("Descricao obrigatoria no produto.")
        line["_product_pending_create"] = product is None
        line["descricao"] = line["descricao"] or str((product or {}).get("descricao", "") or "").strip()
        line["produto_unid"] = line["produto_unid"] or str((product or {}).get("unid", "") or "UN").strip()
        if product is not None and line["preco_unit"] <= 0:
            line["preco_unit"] = round(ports._parse_float(ports.produto_preco_venda(product), 0), 4)
        if not line["ref_externa"]:
            line["ref_externa"] = line["produto_codigo"]
        line["ref_interna"] = ""
        line["material"] = ""
        line["material_family"] = ""
        line["material_subtype"] = ""
        line["material_supplied_by_client"] = False
        line["material_fornecido_cliente"] = False
        line["material_cost_included"] = False
        line["espessura"] = ""
        line["desenho"] = ""
        line["desenho_pdf"] = ""
        line["desenhos_pdf"] = []
        line["ficheiros"] = []
        line["laser_base_active"] = False
        line["laser_base_tempo_unit"] = 0.0
        line["laser_base_preco_unit"] = 0.0
        line["operacao"] = line["operacao"] or "Montagem"
    else:
        if not line["descricao"]:
            raise ValueError("Descricao obrigatoria na linha de servico.")
        line["ref_interna"] = ""
        line["material"] = ""
        line["material_family"] = ""
        line["material_subtype"] = ""
        line["material_supplied_by_client"] = False
        line["material_fornecido_cliente"] = False
        line["material_cost_included"] = False
        line["espessura"] = ""
        line["produto_codigo"] = ""
        line["produto_unid"] = line["produto_unid"] or "SV"
        line["desenho"] = ""
        line["desenho_pdf"] = ""
        line["desenhos_pdf"] = []
        line["ficheiros"] = []
        line["laser_base_active"] = False
        line["laser_base_tempo_unit"] = 0.0
        line["laser_base_preco_unit"] = 0.0
        line["operacao"] = line["operacao"] or "Montagem"
    line["total"] = round(line["qtd"] * line["preco_unit"], 2)

    def _repair_laser_base(snapshot_payload: dict[str, Any]) -> None:
        if line_type != ports.ORC_LINE_TYPE_PIECE:
            return
        if not bool(line.get("laser_base_active", False)):
            return
        current_time = round(ports._parse_float(line.get("tempo_peca_min", 0), 0), 4)
        current_price = round(ports._parse_float(line.get("preco_unit", 0), 0), 4)
        repair_extra_time = round(
            sum(
                ports._quote_collect_non_laser_map(
                    dict(payload.get("tempos_operacao", {}) or {}),
                    dict(snapshot_payload.get("tempos_operacao", {}) or {}),
                    digits=4,
                ).values()
            ),
            4,
        )
        repair_extra_price = round(
            sum(
                ports._quote_collect_non_laser_map(
                    dict(payload.get("custos_operacao", {}) or {}),
                    dict(snapshot_payload.get("custos_operacao", {}) or {}),
                    digits=4,
                ).values()
            ),
            4,
        )
        max_safe_base_time = round(max(0.0, current_time - repair_extra_time), 4)
        max_safe_base_price = round(max(0.0, current_price - repair_extra_price), 4)
        if round(ports._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4) > max_safe_base_time + 0.0001:
            line["laser_base_tempo_unit"] = max_safe_base_time
        if round(ports._parse_float(line.get("laser_base_preco_unit", 0), 0), 4) > max_safe_base_price + 0.0001:
            line["laser_base_preco_unit"] = max_safe_base_price

    def _apply_laser_base_blend(snapshot_payload: dict[str, Any]) -> None:
        if line_type != ports.ORC_LINE_TYPE_PIECE:
            return
        if not bool(line.get("laser_base_active", False)):
            return
        base_time = round(ports._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4)
        base_price = round(ports._parse_float(line.get("laser_base_preco_unit", 0), 0), 4)
        extra_time = 0.0
        extra_price = 0.0
        for op_name, raw_value in dict(snapshot_payload.get("tempos_operacao", {}) or {}).items():
            normalized = str(ports.normalize_operacao_nome(op_name) or op_name or "").strip()
            if normalized and normalized != "Corte Laser":
                extra_time += ports._parse_float(raw_value, 0)
        for op_name, raw_value in dict(snapshot_payload.get("custos_operacao", {}) or {}).items():
            normalized = str(ports.normalize_operacao_nome(op_name) or op_name or "").strip()
            if normalized and normalized != "Corte Laser":
                extra_price += ports._parse_float(raw_value, 0)
        line["tempo_peca_min"] = round(base_time + extra_time, 2)
        line["preco_unit"] = round(base_price + extra_price, 4)
        line["total"] = round(line["qtd"] * line["preco_unit"], 2)

    raw_has_operations = bool(
        stock_item_kind == "raw_material"
        and (
            str(line.get("operacao", "") or "").strip()
            or list(payload.get("operacoes_lista", []) or [])
            or list(payload.get("operacoes_detalhe", []) or [])
            or dict(payload.get("tempos_operacao", {}) or {})
            or dict(payload.get("custos_operacao", {}) or {})
            or ports._parse_float(line.get("tempo_peca_min", 0), 0) > 0
        )
    )
    if stock_item_kind == "raw_material" and not raw_has_operations:
        line["operacoes_lista"] = []
        line["operacoes_fluxo"] = []
        line["operacoes_detalhe"] = []
        line["tempos_operacao"] = {}
        line["custos_operacao"] = {}
        line["quote_cost_snapshot"] = {}
    else:
        snapshot_source = {**dict(payload or {}), **line}
        snapshot = ports._quote_line_operation_snapshot(snapshot_source)
        _repair_laser_base(snapshot)
        _apply_laser_base_blend(snapshot)
        snapshot_source = {**dict(payload or {}), **line}
        snapshot = ports._quote_line_operation_snapshot(snapshot_source)
        line["operacoes_lista"] = list(snapshot.get("operacoes", []) or [])
        line["operacoes_fluxo"] = [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)]
        line["operacoes_detalhe"] = [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        line["tempos_operacao"] = dict(snapshot.get("tempos_operacao", {}) or {})
        line["custos_operacao"] = dict(snapshot.get("custos_operacao", {}) or {})
        line["quote_cost_snapshot"] = dict(snapshot.get("quote_cost_snapshot", {}) or {})
    return line
