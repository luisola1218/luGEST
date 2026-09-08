"""Exercise staged conversion, failures and duplicate protection without runtime."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quotes.application.conversion import QuoteConversion, ConversionRules
from lugest_modules.quotes.application.order_lines import OrderLinePorts
from lugest_modules.quotes.infrastructure.legacy_conversion_repository import LegacyConversionRepository


def main():
    initial = {"orcamentos": [{"numero": "O1", "estado": "Aprovado", "cliente": {"nome": "Novo"},
                              "linhas": [{"qtd": 2, "tempo_peca_min": 3, "material": "Aco", "espessura": "3"}]}],
               "clientes": [], "encomendas": [], "seq": {}, "refs": []}
    state = deepcopy(initial)
    failure = False
    saves = []

    def allocate(data, kind, value):
        data.setdefault("seq", {})[kind] = 2
        return value

    def line_ports(data, client):
        return OrderLinePorts(
            parse_float=lambda value, default=0: float(value or default),
            normalize_orc_line_type=lambda value: value or "piece",
            normalize_operacao_nome=lambda value: value,
            production_route=lambda line: "laser",
            operations=lambda line: ["Corte Laser"],
            operations_text=lambda line: "Corte Laser",
            default_resource=lambda operation, preferred: "Laser1",
            next_reference=lambda used: allocate(data, "reference", "REF1"),
            next_piece_order=lambda: "OPP1",
            build_operacoes_fluxo=lambda value: [{"nome": value}],
            now_iso=lambda: "2026-09-08T12:00:00",
        )

    def save_dataset(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("Persistence failed")

    def repository():
        return LegacyConversionRepository(
            lambda: state, save_dataset,
            next_client=lambda data: allocate(data, "client", "CL1"),
            next_order=lambda data: allocate(data, "order", "E1"),
            order_code=lambda data, order: "OF-1",
            make_line_ports=line_ports,
            register_reference=lambda data, internal, external: data["refs"].extend([internal, external]),
        )

    rules = ConversionRules(
        normalize_client=lambda client: deepcopy(client), now_iso=lambda: "2026-09-08T12:00:00",
        parse_float=lambda value, default=0: float(value or default), normalize_workcenter=lambda value: value,
        assembly_detail=lambda code: {}, update_order_state=lambda order: None,
    )
    pending = repository()
    service = QuoteConversion(pending, rules)
    assert service.convert("O1", "Nota") == "E1"
    assert saves == [{"force": True, "blocking": True}]
    assert state["encomendas"][0]["tempo_estimado"] == 6
    assert state["encomendas"][0]["nota_cliente"] == "Nota"
    assert state["orcamentos"][0]["numero_encomenda"] == "E1"
    assert state["clientes"][0]["nome"] == "Novo"
    original = deepcopy(state)
    for attempt in (lambda: service.convert("O1"), lambda: QuoteConversion(repository(), rules).convert("O1")):
        try:
            attempt()
        except ValueError:
            pass
        else:
            raise AssertionError("Duplicate conversion accepted")
        assert state == original

    state = deepcopy(initial)
    failure = True
    try:
        QuoteConversion(repository(), rules).convert("O1")
    except RuntimeError:
        pass
    else:
        raise AssertionError("Save failure hidden")
    assert state == initial
    failure = False
    state["orcamentos"][0]["linhas"].append({"qtd": 1})
    original = deepcopy(state)
    previous_saves = len(saves)
    try:
        QuoteConversion(repository(), rules).convert("O1")
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid line accepted")
    assert state == original and len(saves) == previous_saves

    state = deepcopy(initial)
    stale = repository()
    state["clientes"].append({"codigo": "OTHER", "nome": "Outra alteracao"})
    original = deepcopy(state)
    try:
        QuoteConversion(stale, rules).convert("O1")
    except ValueError:
        pass
    else:
        raise AssertionError("Stale conversion overwrote local edit")
    assert state == original
    assert "main" not in sys.modules and not any(name.startswith("PySide6") for name in sys.modules)
    print("quote-conversion-ok staged=yes failures-isolated=yes stale-rejected=yes duplicate-rejected=yes blocking-save=yes")


if __name__ == "__main__":
    main()
