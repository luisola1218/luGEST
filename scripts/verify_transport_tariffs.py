"""Tariff CRUD, matching, cost calculation and failure isolation without runtime."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.transport.application.tariffs import TariffService, TariffRules
from lugest_modules.transport.infrastructure.legacy_tariff_repository import LegacyTariffRepository


def main():
    state = {}
    failure = False
    saves = []

    def persist(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyTariffRepository(lambda: state, persist)
    service = TariffService(repository, TariffRules(lambda value, default=0: float(value or default),
                           lambda value: str(value or "").lower().strip(),
                           lambda code, name: (str(code), str(name), "")))
    row = service.save({"zona": "Porto", "valor_base": 10, "valor_por_palete": 3,
                        "valor_por_kg": .1, "valor_por_m3": 2, "custo_minimo": 20})
    assert row["id"] == 1 and service.cost(row, 2, 100, 3) == 32
    assert service.cost(row) == 20
    row["zona"] = "Changed externally"
    assert service.rows()[0]["zona"] == "Porto"
    specific = service.save({"zona": "Porto", "transportadora_id": "S1", "transportadora_nome": "Carrier", "valor_base": 40})
    assert service.match("S1", "Carrier", "porto")["id"] == specific["id"]
    assert service.match("other", "Other", "porto")["id"] == 1
    assert service.suggestion("S1", "Carrier", "porto")["custo_sugerido"] == 40
    original = deepcopy(state)
    try:
        service.save({"zona": "PORTO"})
    except ValueError:
        pass
    else:
        raise AssertionError("Duplicate tariff accepted")
    assert state == original
    failure = True
    for action in (lambda: service.save({"zona": "Lisboa"}),
                   lambda: service.save({"id": 1, "zona": "Porto", "valor_base": 99}),
                   lambda: service.remove(1)):
        original = deepcopy(state)
        try:
            action()
        except RuntimeError:
            pass
        else:
            raise AssertionError("Persistence failure hidden")
        assert state == original
    failure = False
    service.remove(1)
    assert service.match("other", "Other", "porto") is None
    expected = repository.rows()
    state["transportes_tarifarios"][0]["zona"] = "Lisboa"
    try:
        repository.replace([], expected=expected)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale catalog overwrite accepted")
    # Corrupt/duplicate historical IDs must not compare dictionaries on a tie.
    state["transportes_tarifarios"].append(deepcopy(state["transportes_tarifarios"][0]))
    assert service.match("S1", "Carrier", "Lisboa") is not None
    assert all(call == {"force": True, "blocking": True} for call in saves)
    assert "main" not in sys.modules
    print("transport-tariffs-ok crud=yes costs=yes matching=yes failures-restored=yes stale-rejected=yes")


if __name__ == "__main__":
    main()
