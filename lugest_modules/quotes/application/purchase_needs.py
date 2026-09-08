"""Calculate quote shortages, sharing each stock balance across repeated lines."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class PurchaseNeedsRepository(Protocol):
    def quote_lines(self, number: str) -> list[dict[str, Any]]: ...
    def products(self) -> list[dict[str, Any]]: ...
    def material(self, code: str) -> dict[str, Any] | None: ...

@dataclass(frozen=True)
class PurchaseNeedsRules:
    ORC_LINE_TYPE_PRODUCT: str
    ORC_LINE_TYPE_PIECE: str
    norm_text: Callable
    normalize_orc_line_type: Callable
    parse_float: Callable
    is_raw_material: Callable
    detect_materia_formato: Callable

class PurchaseNeeds:
    def __init__(self, repository: PurchaseNeedsRepository, rules: PurchaseNeedsRules):
        self.repository = repository
        self.rules = rules

    def key(self, kind: str, line: dict[str, Any]) -> str:
        if kind == "product":
            code = str(line.get("produto_codigo", "") or "").strip()
            if code:
                return f"product:{code}"
            return "product:new:" + self.rules.norm_text(str(line.get("descricao", "") or "").strip())
        ref = str(line.get("stock_material_id", "") or "").strip()
        if ref:
            return f"material:{ref}"
        parts = [
            str(line.get("material", "") or "").strip(),
            str(line.get("espessura", "") or "").strip(),
            str(line.get("material_subtype", "") or line.get("calc_mode", "") or "").strip(),
            str(line.get("descricao", "") or "").strip(),
        ]
        return "material:new:" + "|".join(self.rules.norm_text(part) for part in parts if part)


    def rows(self, numero: str = "", lines: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        numero_txt = str(numero or "").strip()
        quote_lines = deepcopy(list(lines or []))
        if not quote_lines:
            quote_lines = self.repository.quote_lines(numero_txt)
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in self.repository.products()
            if str(prod.get("codigo", "") or "").strip()
        }
        grouped: dict[str, dict[str, Any]] = {}
        remaining: dict[str, float] = {}

        def shortage(key: str, quantity: float, available: float) -> float:
            balance = remaining.setdefault(key, available)
            remaining[key] = max(0.0, balance - quantity)
            return max(0.0, quantity - balance)

        for raw in quote_lines:
            line = dict(raw or {})
            line_type = self.rules.normalize_orc_line_type(line.get("tipo_item"))
            if line_type == self.rules.ORC_LINE_TYPE_PRODUCT:
                qty = max(0.0, self.rules.parse_float(line.get("qtd", 0), 0))
                if qty <= 1e-9:
                    continue
                code = str(line.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                available = max(0.0, self.rules.parse_float((product or {}).get("qty", 0), 0)) if product is not None else 0.0
                missing = shortage(self.key("product", line), qty, available)
                if missing <= 1e-9:
                    continue
                key = self.key("product", line)
                entry = grouped.setdefault(
                    key,
                    {
                        "kind": "product",
                        "ref": code,
                        "descricao": str(line.get("descricao", "") or (product or {}).get("descricao", "") or "").strip(),
                        "unid": str(line.get("produto_unid", "") or (product or {}).get("unid", "") or "UN").strip() or "UN",
                        "qtd": 0.0,
                        "qtd_disponivel": available,
                        "preco": round(self.rules.parse_float((product or {}).get("p_compra", line.get("preco_unit", 0)), 0), 4),
                        "_product_pending_create": product is None or bool(line.get("_product_pending_create", False)),
                    },
                )
                entry["qtd"] = round(self.rules.parse_float(entry.get("qtd", 0), 0) + missing, 2)
                continue
            if line_type != self.rules.ORC_LINE_TYPE_PIECE or not self.rules.is_raw_material(line):
                continue
            qty = max(0.0, self.rules.parse_float(line.get("qtd", 0), 0))
            if qty <= 1e-9:
                continue
            stock_id = str(line.get("stock_material_id", "") or "").strip()
            material_record = self.repository.material(stock_id) if stock_id else None
            available = 0.0
            if isinstance(material_record, dict):
                available = max(
                    0.0,
                    self.rules.parse_float(material_record.get("quantidade", 0), 0)
                    - self.rules.parse_float(material_record.get("reservado", 0), 0),
                )
            missing = shortage(self.key("material", line), qty, available)
            if missing <= 1e-9:
                continue
            formato = str(line.get("material_subtype", "") or line.get("calc_mode", "") or (material_record or {}).get("formato", "") or "Chapa").strip()
            if formato == "Stock MP":
                formato = str((material_record or {}).get("formato", "") or self.rules.detect_materia_formato(material_record or {}) or "Chapa").strip()
            price = self.rules.parse_float(line.get("price_base_value", 0), 0)
            if price <= 0 and isinstance(material_record, dict):
                price = self.rules.parse_float(material_record.get("p_compra", material_record.get("preco_unid", 0)), 0)
            key = self.key("material", line)
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
                    "comprimento": round(self.rules.parse_float(line.get("length_mm", (material_record or {}).get("comprimento", 0)), 0), 3),
                    "largura": round(self.rules.parse_float(line.get("width_mm", (material_record or {}).get("largura", 0)), 0), 3),
                    "diametro": round(self.rules.parse_float(line.get("diameter_mm", (material_record or {}).get("diametro", 0)), 0), 3),
                    "metros": round(self.rules.parse_float(line.get("meters_per_unit", (material_record or {}).get("metros", 0)), 0), 4),
                    "kg_m": round(self.rules.parse_float(line.get("kg_per_m", (material_record or {}).get("kg_m", 0)), 0), 4),
                    "peso_unid": round(self.rules.parse_float(line.get("stock_metric_value", (material_record or {}).get("peso_unid", 0)), 0), 4),
                    "_material_pending_create": material_record is None,
                    "_material_manual": material_record is None,
                },
            )
            entry["qtd"] = round(self.rules.parse_float(entry.get("qtd", 0), 0) + missing, 2)
        rows = [row for row in grouped.values() if self.rules.parse_float(row.get("qtd", 0), 0) > 0]
        rows.sort(key=lambda row: (str(row.get("kind", "")), str(row.get("ref", "") or row.get("descricao", ""))))
        return rows


