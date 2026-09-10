"""Purchase line validation and material inference without the desktop runtime."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.purchasing.application.line_normalization import PurchaseLines, PurchaseLineRules
from lugest_modules.purchasing.infrastructure.legacy_line_repository import LegacyLineRepository


def main():
    state = {"produtos": [{"codigo": "P1", "categoria": "Fixacao", "peso_unid": 2}],
             "materiais": [{"id": "M1", "material": "Aco", "espessura": "3", "formato": "Chapa", "peso_unid": 6}]}
    original = deepcopy(state)
    number = lambda value, default=0: float(default if value in (None, "") else value)
    service = PurchaseLines(LegacyLineRepository(lambda: state), PurchaseLineRules(
        str.lower, number, number, lambda value: f"{value:g}", deepcopy,
        lambda value: value == "Material", lambda row: row.get("formato", "Chapa"),
        lambda row: row.get("localizacao", ""), lambda row: {},
    ))
    payload = {"ref": "P1", "descricao": "Produto", "qtd": 2, "preco": 5, "iva": 0, "desconto": 10}
    before = deepcopy(payload)
    row = service.normalize(payload)
    assert row["total"] == 9 and row["categoria"] == "Fixacao" and not row["_product_pending_create"]
    for field in ("qtd", "preco"):
        for value in (float("nan"), float("inf"), -float("inf")):
            try:
                service.normalize(dict(payload, **{field: value}))
            except ValueError:
                pass
            else:
                raise AssertionError("Non-finite line accepted")
    material = service.normalize(dict(payload, ref="M1", origem="Material"))
    assert material["material"] == "Aco" and material["peso_unid"] == 6
    assert not material["_material_pending_create"]
    inferred = service.infer_material({"descricao": "Cantoneira 40 x 40 x 3 mm 2 un x 6 m"})
    assert inferred["formato"] == "Cantoneira" and inferred["espessura"] == "3"
    assert inferred["comprimento"] == 40 and inferred["metros"] == 6
    try:
        service.normalize(dict(payload, ref="UNKNOWN", origem="Material"))
    except ValueError:
        pass
    else:
        raise AssertionError("Missing material specification accepted")
    service.repository.material("M1")["material"] = "Changed"
    assert payload == before and state == original and "main" not in sys.modules
    print("purchase-lines-ok products=yes materials=yes inference=yes finite-values=yes detached=yes")


if __name__ == "__main__":
    main()
