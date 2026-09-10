"""Regressions for repeated demand and shared available inventory."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quotes.application.purchase_quote import PurchaseQuote
from lugest_modules.quotes.application.purchase_needs import PurchaseNeeds, PurchaseNeedsRules
from lugest_modules.quotes.infrastructure.legacy_purchase_needs_repository import LegacyPurchaseNeedsRepository


def main():
    state = {"produtos": [{"codigo": "P1", "qty": 10, "p_compra": 2}],
             "orcamentos": [{"numero": "O1", "linhas": [
                 {"tipo_item": "product", "produto_codigo": "P1", "qtd": 6},
                 {"tipo_item": "product", "produto_codigo": "P1", "qtd": 6}]}]}
    material = {"id": "M1", "quantidade": 10, "reservado": 3, "p_compra": 5, "formato": "Chapa"}
    repository = LegacyPurchaseNeedsRepository(lambda: state, lambda code: material if code == "M1" else None)
    rules = PurchaseNeedsRules(
        ORC_LINE_TYPE_PRODUCT="product", ORC_LINE_TYPE_PIECE="piece",
        norm_text=lambda value: value.lower(), normalize_orc_line_type=lambda value: value,
        parse_float=lambda value, default=0: float(value or default),
        is_raw_material=lambda line: bool(line.get("raw")),
        detect_materia_formato=lambda row: "Chapa",
    )
    service = PurchaseNeeds(repository, rules)
    original = deepcopy(state)
    needs = service.rows("O1")
    assert len(needs) == 1 and needs[0]["qtd"] == 2 and needs[0]["qtd_disponivel"] == 10
    assert state == original
    line = {"tipo_item": "piece", "raw": True, "stock_material_id": "M1", "qtd": 6}
    needs = service.rows(lines=[line, line])
    assert needs[0]["qtd"] == 5 and needs[0]["qtd_disponivel"] == 7
    assert material["quantidade"] == 10 and material["reservado"] == 3
    product = state["orcamentos"][0]["linhas"][0]
    assert service.rows(lines=[dict(product, qtd=10)]) == []
    assert service.rows(lines=[dict(product, qtd=12), dict(product, qtd=4)])[0]["qtd"] == 6
    pending = {"tipo_item": "product", "descricao": "Novo", "qtd": 3}
    needs = service.rows(lines=[pending, dict(pending, qtd=4)])
    assert needs[0]["qtd"] == 7 and needs[0]["_product_pending_create"]
    assert service.rows(lines=[{"tipo_item": "service", "qtd": 100}]) == []
    try:
        service.rows("missing")
    except ValueError:
        pass
    else:
        raise AssertionError("Missing quote accepted")
    saved = []
    def create(payload):
        saved.append(deepcopy(payload))
        return {"numero": "NE1", "linhas": payload["lines"]}
    quotation = PurchaseQuote(service, rules.parse_float, create, lambda number: {"numero": number})
    result = quotation.create("O1")
    assert result["numero"] == "NE1" and saved[0]["lines"][0]["qtd"] == 2
    assert result["line_count"] == 1 and state == original
    try:
        quotation.create("O1", [dict(product, qtd=1)])
    except ValueError:
        pass
    else:
        raise AssertionError("Purchase request created without shortage")
    assert len(saved) == 1
    assert "main" not in sys.modules and not any(name.startswith("PySide6") for name in sys.modules)
    print("purchase-needs-ok repeated-products=yes repeated-materials=yes reservations=yes read-only=yes")


if __name__ == "__main__":
    main()
