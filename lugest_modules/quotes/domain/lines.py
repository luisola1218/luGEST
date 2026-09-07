"""Quote line construction, independent of UI and storage."""
from __future__ import annotations

def service_line(descricao: str, qtd: float, unid: str, preco_unit: float, operacao: str, *, line_type: str) -> dict | None:
    quantidade = round(float(qtd or 0), 2)
    preco = round(float(preco_unit or 0), 4)
    if quantidade <= 0 or preco <= 0:
        return None
    return {
        "tipo_item": line_type,
        "descricao": str(descricao or "").strip(),
        "produto_unid": str(unid or "SV").strip() or "SV",
        "operacao": str(operacao or "Montagem").strip() or "Montagem",
        "qtd": quantidade,
        "preco_unit": preco,
    }

def product_line(product: dict | None, qtd: float, *, descricao_extra: str = "", line_type: str) -> dict | None:
    row = dict(product or {})
    code = str(row.get("codigo", "") or "").strip()
    if not code:
        return None
    quantidade = round(float(qtd or 0), 2)
    if quantidade <= 0:
        return None
    preco = round(float(row.get("preco_venda", row.get("pvp1", row.get("preco_unid", row.get("preco", 0)))) or 0), 4)
    if preco <= 0:
        return None
    descricao = str(row.get("descricao", "") or code).strip()
    if descricao_extra:
        descricao = f"{descricao} | {descricao_extra.strip()}"
    return {
        "tipo_item": line_type,
        "stock_item_kind": "product",
        "produto_codigo": code,
        "produto_unid": str(row.get("unid", "") or "UN").strip() or "UN",
        "descricao": descricao,
        "ref_externa": code,
        "operacao": str(row.get("tipo", "") or "Montagem").strip() or "Montagem",
        "qtd": quantidade,
        "preco_unit": preco,
    }
