"""Synchronize technical data and prices on owned purchase note lines."""
from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class MaterialLineRules:
    is_material: Callable
    unit_price: Callable
    parse_float: Callable
    format: Callable
    description: Callable

def sync_material_lines(ne, materials, rules):
    changed = False
    mat_map = {m.get("id"): m for m in materials}
    for l in ne.get("linhas", []):
        if not rules.is_material(l.get("origem", "")):
            continue
        if l.get("_stock_in") or l.get("entregue"):
            continue
        m = mat_map.get(l.get("ref"))
        if not m:
            continue
        new_preco = round(rules.unit_price(m), 6)
        cur_preco = rules.parse_float(l.get("preco", 0), 0)
        if abs(new_preco - cur_preco) > 1e-9:
            l["preco"] = new_preco
            q = rules.parse_float(l.get("qtd", 0), 0)
            desconto = max(0.0, min(100.0, rules.parse_float(l.get("desconto", 0), 0)))
            iva = max(0.0, min(100.0, rules.parse_float(l.get("iva", 23), 23)))
            l["total"] = round(((q * new_preco) * (1.0 - (desconto / 100.0))) * (1.0 + (iva / 100.0)), 6)
            changed = True
        formato = m.get("formato", rules.format(m))
        comp = rules.parse_float(m.get("comprimento", 0), 0)
        larg = rules.parse_float(m.get("largura", 0), 0)
        metros = rules.parse_float(m.get("metros", 0), 0)
        new_desc = rules.description(
            m.get("material", ""),
            m.get("espessura", ""),
            formato,
            comp,
            larg,
            metros,
        )
        if (l.get("descricao", "") or "") != new_desc:
            l["descricao"] = new_desc
            changed = True
        if l.get("unid", "") != "UN":
            l["unid"] = "UN"
            changed = True
        # Mantemos apenas a base tecnica sincronizada. Lote e localizacao passam
        # a ser decididos linha a linha no momento da rececao.
        new_meta = {
            "source_material_id": m.get("id", ""),
            "material": m.get("material", ""),
            "espessura": m.get("espessura", ""),
            "comprimento": rules.parse_float(m.get("comprimento", 0), 0),
            "largura": rules.parse_float(m.get("largura", 0), 0),
            "metros": rules.parse_float(m.get("metros", 0), 0),
            "peso_unid": rules.parse_float(m.get("peso_unid", 0), 0),
            "p_compra": rules.parse_float(m.get("p_compra", 0), 0),
            "formato": formato,
        }
        for k, v in new_meta.items():
            if l.get(k) != v:
                l[k] = v
                changed = True
    if changed:
        ne["total"] = sum(rules.parse_float(x.get("total", 0), 0) for x in ne.get("linhas", []))
    return changed
