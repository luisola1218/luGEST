"""Trip creation validation must precede shared writes and identifier allocation."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.transport.application.trips import TripCommands, TripRules
from lugest_modules.transport.infrastructure.legacy_trip_repository import LegacyTripRepository


def main():
    state = {"encomendas": [{"numero": "E1"}]}
    failure = False
    allocations = []

    def allocate(data, requested):
        allocations.append(requested)
        number = data.setdefault("seq", {}).get("transporte", 1)
        data["seq"]["transporte"] = number + 1
        return requested or f"T{number}"

    def persist(**kwargs):
        nonlocal state
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyTripRepository(lambda: state, persist, str.lower, allocate)
    service = TripCommands(repository, TripRules(
        lambda value, default=0: float(value or default), str.lower, lambda: "2026-09-08T12:00:00",
        lambda: "TEST", lambda code, name: (code, name, ""),
        lambda: {"tipo_responsavel": "Nosso Cargo", "estado": "Planeado", "data_planeada": "2026-09-08", "hora_saida": "08:00", "origem": "Origem"},
        lambda trip: None))
    original = deepcopy(state)
    try:
        service.save({"tipo_responsavel": "Subcontratado"})
    except ValueError:
        pass
    else:
        raise AssertionError("Missing carrier accepted")
    assert state == original and allocations == []
    assert service.save({"motorista": "Motorista"}) == "T1"
    state["transportes"][0]["custom"] = {"preserved": True}
    service.save({"numero": "T1", "motorista": "Novo"})
    assert state["transportes"][0]["custom"] == {"preserved": True}
    assert len(allocations) == 1
    failure = True
    for action in (lambda: service.save({"motorista": "Outro"}),
                   lambda: service.save({"numero": "T1", "motorista": "Falha"}),
                   lambda: service.remove("T1")):
        original = deepcopy(state)
        try:
            action()
        except RuntimeError:
            pass
        else:
            raise AssertionError("Write failure hidden")
        assert state == original
    failure = False
    service.remove("T1")
    assert not state["transportes"]
    assert "main" not in sys.modules
    print("transport-trips-ok validate-before-allocation=yes create=yes edit=yes remove=yes recovery=yes")


if __name__ == "__main__":
    main()
