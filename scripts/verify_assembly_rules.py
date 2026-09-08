"""Exercise assembly pricing without Qt, runtime globals or a database."""
from copy import deepcopy
from dataclasses import replace
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quotes.application.assemblies import AssemblyRules, normalize_item, price_item, refresh_model, expand_model
from lugest_modules.quotes.application.assembly_refresh import AssemblyRefresh
from lugest_modules.quotes.application.assembly_catalog import AssemblyCatalog
from lugest_modules.quotes.application.assembly_queries import AssemblyQueries
from lugest_modules.quotes.infrastructure.legacy_assembly_repository import LegacyAssemblyRepository


def main():
    product = {"codigo": "P1", "descricao": "Parafuso", "unid": "UN", "price": 4}
    material = {"id": "M1", "formato": "chapa", "p_compra": 3}
    rules = AssemblyRules(
        ORC_LINE_TYPE_PIECE="piece", ORC_LINE_TYPE_PRODUCT="product",
        parse_float=lambda value, default=0: float(value or default),
        normalize_orc_line_type=lambda value: value or "service",
        product_lookup=lambda code: product if code == "P1" else None,
        produto_preco_venda=lambda row: row["price"],
        produto_preco_unitario=lambda row: row["price"],
        orc_line_is_product=lambda row: row.get("tipo_item") == "product",
        orc_line_is_piece=lambda row: row.get("tipo_item") == "piece",
        norm_text=lambda value: value.lower(),
        detect_materia_formato=lambda row: "chapa",
        material_by_id=lambda code: material if code == "M1" else None,
        material_candidates=lambda: [material],
        material_price_preview=lambda row: {"preco_unid": 9},
        quote_source=lambda item, code: ({"preco_unit": 7, "ref_externa": "R1"}, "O1"),
        now_iso=lambda: "2026-09-08T10:00:00",
    )
    source = {"tipo_item": "product", "produto_codigo": "P1", "qtd": 2,
              "preco_unit": 2, "laser_snapshot": {"nested": [1]}}
    before = deepcopy(source)
    normalized = normalize_item(rules, source)
    assert normalized["descricao"] == "Parafuso"
    normalized["laser_snapshot"]["nested"].append(2)
    assert source == before
    priced, changed = price_item(rules, source, "C1")
    assert changed and priced["preco_unit"] == 4 and priced["preco_anterior"] == 2
    assert priced["pricing_source"] == "product_stock" and source == before
    stable, changed = price_item(rules, priced, "C1")
    assert not changed and stable["preco_atualizado_em"] == priced["preco_atualizado_em"]

    raw = {"tipo_item": "piece", "descricao": "Chapa", "material": "Aco",
           "espessura": "3", "qtd": 1, "preco_unit": 1,
           "stock_material_id": "M1", "stock_metric_value": 2, "price_base_label": "kg"}
    priced, _ = price_item(rules, raw, "C1")
    assert priced["preco_unit"] == 6 and priced["pricing_source"] == "material_stock"
    inferred = dict(raw, stock_material_id="", calc_mode="chapa", price_base_value=3)
    priced, _ = price_item(rules, inferred, "C1")
    assert priced["stock_material_id"] == "M1" and inferred["stock_material_id"] == ""
    laser = dict(raw, stock_material_id="", stock_metric_value=0)
    priced, _ = price_item(rules, laser, "C1")
    assert priced["preco_unit"] == 7 and priced["source_quote_number"] == "O1"
    manual, _ = price_item(replace(rules, quote_source=lambda *args: (None, "")), laser, "C1")
    assert manual["preco_unit"] == 1 and not manual["pricing_linked"]

    model = {"codigo": "C1", "itens": [source, raw], "margem_perc": 25}
    original = deepcopy(model)
    refreshed, changed = refresh_model(rules, model)
    assert changed and refreshed["total_custo"] == 14 and refreshed["total_final"] == 17.5
    assert model == original
    expanded = expand_model(rules, refreshed, 3, "GROUP1")
    assert [row["qtd"] for row in expanded] == [6, 3]
    assert all(row["grupo_uuid"] == "GROUP1" for row in expanded)
    expanded[0]["laser_snapshot"]["nested"].append(99)
    assert refreshed["itens"][0]["laser_snapshot"]["nested"] == [1]
    invalid = deepcopy(model)
    invalid["itens"].append({"qtd": 0})
    original = deepcopy(invalid)
    try:
        refresh_model(rules, invalid)
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid quantity accepted")
    assert invalid == original
    state = {"conjuntos": [deepcopy(model), dict(deepcopy(model), codigo="C2")]}
    saves = []
    failure = False

    def save_dataset(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)  # Persistence can replace the loaded snapshot.
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyAssemblyRepository(lambda: state, save_dataset)
    service = AssemblyRefresh(repository, rules)
    assert service.refresh() == {"updated": True}
    assert [row["param_codigo"] for row in state["conjuntos"]] == ["0001", "0002"]
    assert len(saves) == 1
    assert service.refresh() == {"updated": False} and len(saves) == 1
    returned = service.refresh("C1")
    returned["itens"].clear()
    assert len(state["conjuntos"][0]["itens"]) == 2
    state["conjuntos"][1]["itens"].append({"qtd": 0})
    product["price"] = 8  # First assembly would change before the invalid second.
    original = deepcopy(state)
    try:
        service.refresh()
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid catalog accepted")
    assert state == original and len(saves) == 1
    state["conjuntos"][1]["itens"].pop()
    failure = True
    original = deepcopy(state)
    try:
        service.refresh()
    except RuntimeError:
        pass
    else:
        raise AssertionError("Save failure hidden")
    assert state == original
    expected = repository.models()
    state["conjuntos"][0]["descricao"] = "Concurrent edit"
    try:
        repository.replace(expected, expected=expected)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale catalog overwrite accepted")
    assert state["conjuntos"][0]["descricao"] == "Concurrent edit"
    failure = False
    state = {}
    catalog = AssemblyCatalog(repository, rules, lambda: "C1", live_prices=True)
    payload = {"descricao": "Conjunto", "itens": [source], "margem_perc": 25,
               "ficha_tecnica": {"familia_produto": "Teste"}}
    assert catalog.save(payload) == "C1"
    saved = state["conjuntos"][0]
    assert saved["total_custo"] == 16 and saved["total_final"] == 20
    assert saved["param_codigo"] == "0001"
    saved["custom_metadata"] = {"preserved": True}
    created_at = saved["created_at"]
    catalog.save({"codigo": "C1", "descricao": "Editado", "itens": [source]})
    saved = state["conjuntos"][0]
    assert saved["custom_metadata"] == {"preserved": True}
    assert saved["ficha_tecnica"]["familia_produto"] == "Teste"
    assert saved["created_at"] == created_at
    failure = True
    for action in (lambda: catalog.save(dict(payload, codigo="C2")),
                   lambda: catalog.save(dict(payload, codigo="C1")),
                   lambda: catalog.remove("C1")):
        original = deepcopy(state)
        try:
            action()
        except RuntimeError:
            pass
        else:
            raise AssertionError("Catalog write failure hidden")
        assert state == original
    failure = False
    templates = AssemblyCatalog(LegacyAssemblyRepository(lambda: state, save_dataset, "conjuntos_modelo"),
                                rules, lambda: "T1", live_prices=False)
    assert templates.save(payload) == "T1"
    assert state["conjuntos_modelo"][0]["itens"][0]["preco_unit"] == 2
    queries = AssemblyQueries(repository, rules, lambda row: row.get("tipo_item") == "service")
    assert queries.rows("Editado")[0]["produtos"] == 1
    assert queries.rows("not present") == []
    detail = queries.detail("C1")
    detail["itens"][0]["laser_snapshot"]["nested"].append(999)
    assert state["conjuntos"][0]["itens"][0]["laser_snapshot"]["nested"] == [1]
    template_queries = AssemblyQueries(templates.repository, rules, queries.is_service)
    assert template_queries.template_rows()[0]["total_base"] == 4
    assert template_queries.template_detail("T1")["descricao"] == "Conjunto"
    templates.remove("T1")
    assert state["conjuntos_modelo"] == [] and len(state["conjuntos"]) == 1
    catalog.remove("C1")
    assert state["conjuntos"] == []
    assert "main" not in sys.modules and not any(name.startswith("PySide6") for name in sys.modules)
    print("assembly-rules-ok catalog-prices=yes fallback=yes detached=yes failure-isolation=yes")


if __name__ == "__main__":
    main()
