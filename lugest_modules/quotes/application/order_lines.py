"""Prepare manufacturing and assembly lines for a quote conversion.

Catalog rules, routing, clocks and identifier allocation are explicit. The
result owns its nested data; persistence and reference registration belong to
the conversion coordinator. Allocators can still reserve identifiers externally.
"""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class OrderLinePorts:
    parse_float: Callable
    normalize_orc_line_type: Callable
    normalize_operacao_nome: Callable
    production_route: Callable
    operations: Callable
    operations_text: Callable
    default_resource: Callable
    next_reference: Callable
    next_piece_order: Callable
    build_operacoes_fluxo: Callable
    now_iso: Callable

@dataclass(frozen=True)
class OrderLines:
    materials: list[dict[str, Any]]
    assembly_items: list[dict[str, Any]]
    estimated_minutes: float
    references: list[tuple[str, str]]

def build_order_lines(ports: OrderLinePorts, lines: list[dict[str, Any]], enc_of: str, workcenter: str) -> OrderLines:
    references: list[tuple[str, str]] = []
    mats: dict[str, dict[str, Any]] = {}
    piece_idx = 1
    total_time = 0.0
    used_refs: set[str] = set()
    montagem_items: list[dict[str, Any]] = []
    for line in deepcopy(lines):
        line_type = ports.normalize_orc_line_type(line.get("tipo_item"))
        production_route = ports.production_route(line)
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
                    "preco_unit": round(ports.parse_float(line.get("preco_unit", 0), 0), 4),
                    "conjunto_codigo": str(line.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                    "estado": "Pendente" if production_route == "montagem" else "Componente",
                    "obs": (
                        str(line.get("operacao", "") or production_route).strip()
                        if production_route == "montagem"
                        else "Conjunto sem desenho/operacoes tecnicas. Nao segue para operador."
                    ),
                    "created_at": ports.now_iso(),
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
        planning_ops = [op for op in ports.operations(line) if op != "Montagem"]
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
            str(ports.normalize_operacao_nome(op_name) or op_name or "").strip(): ports.parse_float(raw_value, 0)
            for op_name, raw_value in dict(line.get("tempos_operacao", {}) or {}).items()
            if str(ports.normalize_operacao_nome(op_name) or op_name or "").strip() and ports.parse_float(raw_value, 0) > 0
        }
        if detailed_op_times:
            for op_name, unit_time in detailed_op_times.items():
                if op_name not in planning_ops:
                    continue
                operation_time = unit_time * max(qtd_line, 0.0)
                tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + operation_time, 2)
                if op_name == "Corte Laser":
                    esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + operation_time, 2)
                    if not str(maquinas_operacao.get(op_name, "") or "").strip():
                        maquinas_operacao[op_name] = ports.default_resource(op_name, preferred=workcenter)
        elif len(planning_ops) == 1:
            op_name = planning_ops[0]
            tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            if op_name == "Corte Laser":
                esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            if not str(maquinas_operacao.get(op_name, "") or "").strip():
                maquinas_operacao[op_name] = ports.default_resource(op_name, preferred=workcenter)
        elif "Corte Laser" in planning_ops:
            tempos_operacao["Corte Laser"] = round(float(tempos_operacao.get("Corte Laser", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
            if not str(maquinas_operacao.get("Corte Laser", "") or "").strip():
                maquinas_operacao["Corte Laser"] = ports.default_resource("Corte Laser", preferred=workcenter)
        raw_ref_interna = str(line.get("ref_interna", "") or "").strip()
        if raw_ref_interna and raw_ref_interna not in used_refs:
            ref_interna = raw_ref_interna
        else:
            ref_interna = str(ports.next_reference(list(used_refs)))
        used_refs.add(ref_interna)
        ops_txt = ports.operations_text(line)
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
            "opp": f"OPP-{enc_of.split('-', 1)[1]}-{piece_idx:02d}" if enc_of.startswith("OF-") and "-" in enc_of else ports.next_piece_order(),
            "estado": "Preparacao",
            "produzido_ok": 0.0,
            "produzido_nok": 0.0,
            "inicio_producao": "",
            "fim_producao": "",
        }
        peca["operacoes_fluxo"] = ports.build_operacoes_fluxo(ops_txt)
        piece_idx += 1
        mats[material]["espessuras"][espessura]["pecas"].append(peca)
        references.append((peca["ref_interna"], peca["ref_externa"]))
    materials = []
    for row in mats.values():
        row["espessuras"] = list(row["espessuras"].values())
        materials.append(row)



    return OrderLines(materials, montagem_items, round(total_time, 2), references)
