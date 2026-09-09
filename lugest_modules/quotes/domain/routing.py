"""Quote line classification and production routing without shared state."""
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class RoutingRules:
    piece_type: str
    product_type: str
    service_type: str
    normalize_type: Callable
    normalize_operation: Callable
    normalize_text: Callable
    parse_float: Callable
    parse_operations: Callable
    format_operations: Callable

class LineRouting:
    def __init__(self, rules: RoutingRules):
        self.rules = rules

    def operations(self, line: dict[str, Any] | None = None) -> list[str]:
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
        return self.rules.parse_operations(values)


    def operations_text(self, line: dict[str, Any] | None = None) -> str:
        return self.rules.format_operations(self.operations(line))


    def production_ready(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.rules.normalize_type(row.get("tipo_item")) != self.rules.piece_type:
            return False
        drawing_path = str(row.get("desenho", "") or "").strip()
        ops = [
            str(self.rules.normalize_operation(op) or op or "").strip()
            for op in self.operations(row)
        ]
        ops = [op for op in ops if op and op != "Montagem"]
        material = str(row.get("material", "") or "").strip()
        thickness = str(row.get("espessura", "") or "").strip()
        time_per_piece = self.rules.parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0)
        detail_ready = bool(
            list(row.get("operacoes_detalhe", []) or [])
            or dict(row.get("tempos_operacao", {}) or {})
            or dict(row.get("custos_operacao", {}) or {})
        )
        has_work = bool(ops or detail_ready or time_per_piece > 0)
        return bool(has_work and (drawing_path or (material and thickness)))


    def raw_material(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.rules.normalize_type(row.get("tipo_item")) != self.rules.piece_type:
            return False
        if str(row.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(row.get("stock_material_id", "") or "").strip():
            return True
        if self.stock_material_reference(row.get("ref_externa")):
            if str(row.get("desenho", "") or "").strip():
                return False
            if round(self.rules.parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 4) > 0:
                return False
            operacao_norm = self.rules.normalize_text(str(row.get("operacao", "") or "").strip())
            if operacao_norm in {"", "-", "stockmp", "materia prima", "materia-prima"}:
                return True
        subtype = self.rules.normalize_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        if subtype == "stockmp":
            return True
        return False


    def stock_material_reference(self, value: Any) -> bool:
        raw = str(value or "").strip().upper()
        return bool(raw.startswith("MAT") and raw[3:].isdigit())


    def production_route(self, line: dict[str, Any] | None = None) -> str:
        row = dict(line or {})
        line_type = self.rules.normalize_type(row.get("tipo_item"))
        if line_type == self.rules.product_type:
            return "montagem"
        if line_type == self.rules.service_type:
            return "montagem"
        if not self.production_ready(row):
            return "conjunto"
        ops = [str(self.rules.normalize_operation(op) or op or "").strip() for op in self.operations(row)]
        ops = [op for op in ops if op]
        subtype_norm = self.rules.normalize_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        material_norm = self.rules.normalize_text(str(row.get("material", "") or "").strip())
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
