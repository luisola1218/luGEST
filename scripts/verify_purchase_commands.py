"""A purchase save must stage note, catalog prices and assemblies together."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.purchasing.application.note_commands import NoteCommands, PurchaseWriteRules
from lugest_modules.purchasing.application.pricing import PurchasePricing, PurchasePriceRules
from lugest_modules.purchasing.application.material_lines import MaterialLineRules, sync_material_lines
from lugest_modules.purchasing.infrastructure.legacy_purchase_repository import LegacyPurchaseRepository


def main():
    state = {
        "produtos": [{"codigo": "P1", "descricao": "Produto", "categoria": "peso", "p_compra": 2, "peso_unid": 2}],
        "materiais": [{"id": "M1", "material": "Aco", "formato": "Chapa", "peso_unid": 6, "p_compra": 3}],
        "conjuntos": [{"codigo": "C1", "price": 2}],
        "notas_encomenda": [{"numero": "OLD", "linhas": [
            {"ref": "P1", "origem": "Produto", "qtd": 2, "preco": 4, "total": 8, "iva": 0},
            {"ref": "M1", "origem": "Material", "qtd": 1, "preco": 18, "total": 18, "iva": 0},
        ]}],
    }
    saves = []
    allocations = []
    failure = False
    refresh_failure = False
    number = lambda value, default=0: float(default if value in (None, "") else value)
    is_material = lambda value: value == "Material"
    def allocate(candidate):
        allocations.append(True)
        candidate.setdefault("seq", {})["ne"] = 2
        return "NEW"
    def save(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")
    def normalize(line):
        if number(line.get("qtd")) <= 0:
            raise ValueError("Invalid quantity")
        return dict(line, total=number(line["qtd"]) * number(line["preco"]))
    def refresh(assemblies, products, materials):
        if refresh_failure:
            raise RuntimeError("Assembly calculation failed")
        return [dict(row, price=products[0]["p_compra"] + materials[0]["p_compra"]) for row in assemblies]
    material_rules = MaterialLineRules(is_material, lambda material: material["p_compra"] * material["peso_unid"],
                                       number, lambda row: row["formato"], lambda *args: "Material atualizado")
    pricing = PurchasePricing(PurchasePriceRules(
        number, is_material, lambda product: product["p_compra"] * product["peso_unid"],
        lambda material: material["formato"], lambda: "NOW", lambda category, kind: category,
        lambda note, materials: sync_material_lines(note, materials, material_rules), refresh,
    ))
    repository = LegacyPurchaseRepository(lambda: state, save, allocate)
    service = NoteCommands(repository, PurchaseWriteRules(
        normalize, lambda code, name: (code, name, ""), lambda name: (name, name, ""),
        lambda: "NOW", number, lambda note: "rfq",
        lambda note: note.update(total=sum(line["total"] for line in note["linhas"])),
    ), pricing)
    payload = {"lines": [{"ref": "P1", "origem": "Produto", "descricao": "Produto", "qtd": 2, "preco": 8},
                         {"ref": "M1", "origem": "Material", "descricao": "Material", "qtd": 1, "preco": 30}]}
    original = deepcopy(state)
    bad = deepcopy(payload)
    bad["lines"][1]["qtd"] = 0
    try:
        service.save(bad)
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid batch accepted")
    assert state == original and not allocations and not saves
    refresh_failure = True
    try:
        service.save(payload)
    except RuntimeError:
        pass
    else:
        raise AssertionError("Calculation failure hidden")
    assert state == original and not allocations and not saves
    refresh_failure = False
    failure = True
    try:
        service.save(payload)
    except RuntimeError:
        pass
    else:
        raise AssertionError("Persistence failure hidden")
    assert state == original and "seq" not in state
    failure = False
    before = len(saves)
    created = service.save(payload)
    assert len(saves) == before + 1 and saves[-1] == dict(force=True, blocking=True)
    assert created["numero"] == "NEW" and created["total"] == 46
    assert state["produtos"][0]["p_compra"] == 4 and state["materiais"][0]["p_compra"] == 5
    assert state["conjuntos"][0]["price"] == 9 and state["notas_encomenda"][0]["total"] == 46
    created["linhas"].clear()
    assert len(state["notas_encomenda"][-1]["linhas"]) == 2 and payload["lines"][0]["preco"] == 8
    expected = repository.load()
    state["produtos"][0]["p_compra"] = 7
    try:
        repository.save_catalogs(expected, expected=expected)
    except ValueError:
        pass
    else:
        raise AssertionError("Concurrent catalog change overwritten")
    assert state["produtos"][0]["p_compra"] == 7 and "main" not in sys.modules
    print("purchase-commands-ok staged=yes one-save=yes price-sync=yes assembly-prices=yes recovery=yes stale=yes")


if __name__ == "__main__":
    main()
