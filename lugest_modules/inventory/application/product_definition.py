"""Product definitions and preview calculations with explicit capabilities."""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class ProductDefinitionRules:
    unit_price: Callable[[dict], float]
    _parse_float: Callable[..., Any]
    _product_dimensoes: Callable[..., Any]
    _product_resolve_catalog_fields: Callable[..., Any]
    product_copilot_analysis: Callable[..., Any]
    product_next_code: Callable[..., Any]


def price_preview(rules: ProductDefinitionRules, payload: dict[str, Any]) -> dict[str, Any]:
    """Calcula os indicadores editáveis sem executar classificação inteligente."""
    product = dict(payload or {})
    quantity = rules._parse_float(product.get("qty", product.get("quantidade", 0)), 0)
    unit = str(product.get("unid", "UN") or "UN").strip() or "UN"
    unit_price = rules._parse_float(rules.unit_price(product), 0)
    return {
        "preco_unid": round(unit_price, 4),
        "qty": quantity,
        "unit": unit,
        "valor_stock": round(unit_price * quantity, 2),
    }

def normalize_product(rules: ProductDefinitionRules, payload: dict[str, Any]) -> dict[str, Any]:
    code = str(payload.get("codigo", "") or "").strip() or rules.product_next_code()
    descricao = str(payload.get("descricao", "") or "").strip()
    if not code:
        raise ValueError("Codigo do produto em falta.")
    if not descricao:
        raise ValueError("Descricao do produto em falta.")
    normalized_payload = dict(payload)
    intelligence = rules.product_copilot_analysis(descricao, code)
    suggestion = intelligence
    for field in ("categoria", "subcat", "tipo", "dimensoes"):
        if not str(normalized_payload.get(field, "") or "").strip():
            suggested_value = str(suggestion.get(field, "") or "").strip()
            if suggested_value:
                normalized_payload[field] = suggested_value
    catalog_fields = rules._product_resolve_catalog_fields(normalized_payload)
    categoria = str(catalog_fields.get("categoria", "") or "").strip()
    tipo = str(catalog_fields.get("tipo", "") or "").strip()
    metros_unidade = rules._parse_float(payload.get("metros_unidade", payload.get("metros", 0)), 0)
    prod = {
        "codigo": code,
        "descricao": descricao,
        "categoria": categoria,
        "category_id": str(catalog_fields.get("category_id", "") or "").strip(),
        "subcat": str(catalog_fields.get("subcat", "") or "").strip(),
        "subcategory_id": str(catalog_fields.get("subcategory_id", "") or "").strip(),
        "tipo": tipo,
        "type_id": str(catalog_fields.get("type_id", "") or "").strip(),
        "category_icon": str(catalog_fields.get("category_icon", "") or "").strip(),
        "category_badge": str(catalog_fields.get("category_badge", "") or "").strip(),
        "category_tone": str(catalog_fields.get("category_tone", "") or "").strip(),
        "dimensoes": str(normalized_payload.get("dimensoes", "") or "").strip(),
        "comprimento": rules._parse_float(payload.get("comprimento", 0), 0),
        "largura": rules._parse_float(payload.get("largura", 0), 0),
        "espessura": rules._parse_float(payload.get("espessura", 0), 0),
        "metros_unidade": metros_unidade,
        "metros": metros_unidade,
        "peso_unid": rules._parse_float(payload.get("peso_unid", 0), 0),
        "fabricante": str(payload.get("fabricante", "") or "").strip(),
        "modelo": str(payload.get("modelo", "") or "").strip(),
        "unid": str(payload.get("unid", "UN") or "UN").strip() or "UN",
        "qty": rules._parse_float(payload.get("qty", payload.get("quantidade", 0)), 0),
        "alerta": rules._parse_float(payload.get("alerta", 0), 0),
        "p_compra": rules._parse_float(payload.get("p_compra", 0), 0),
        "pvp1": rules._parse_float(payload.get("pvp1", 0), 0),
        "pvp2": rules._parse_float(payload.get("pvp2", 0), 0),
        "obs": str(payload.get("obs", "") or "").strip(),
        "catalog_intelligence": {
            "engine": str(intelligence.get("engine", "") or ""),
            "confidence": round(float(intelligence.get("confidence", 0) or 0), 4),
            "reason": str(intelligence.get("reason", "") or ""),
            "attributes": dict(intelligence.get("atributos", {}) or {}),
            "normalized_description": str(intelligence.get("descricao_normalizada", "") or ""),
        },
    }
    if not prod["dimensoes"] and (prod["comprimento"] > 0 or prod["largura"] > 0 or prod["espessura"] > 0):
        prod["dimensoes"] = rules._product_dimensoes(prod)
    prod["preco_unid"] = round(rules._parse_float(rules.unit_price(prod), 0), 4)
    return prod
