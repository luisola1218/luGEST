"""Assignment batches must be all-or-nothing in the local snapshot."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.transport.application.assignments import TripAssignments, AssignmentRules
from lugest_modules.transport.infrastructure.legacy_trip_repository import LegacyTripRepository


def main():
    state = {"transportes": [{"numero": "T1", "paragens": []}, {"numero": "T2", "paragens": []}],
             "encomendas": [{"numero": "E1", "cliente": "C1"}, {"numero": "E2", "cliente": "C1"}],
             "clientes": [{"codigo": "C1", "nome": "Cliente"}]}
    failure = False

    def persist(**kwargs):
        nonlocal state
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyTripRepository(lambda: state, persist, str.lower)
    service = TripAssignments(repository, AssignmentRules(
        lambda value, default=0: float(value or default), str.lower, lambda: "2026-09-08T12:00:00",
        lambda order: True, lambda number: {}, lambda order, client: {},
        lambda *args: {"custo_sugerido": 12}, lambda trip: None,
        lambda number: {"paragens": [{"encomenda_numero": "E1", "custo_sugerido": 15}]}))
    original = deepcopy(state)
    try:
        service.assign("T1", ["E1", "MISSING"])
    except ValueError:
        pass
    else:
        raise AssertionError("Missing order accepted")
    assert state == original
    service.assign("T1", ["E1", "E1"])
    assert len(state["transportes"][0]["paragens"]) == 1
    assert state["encomendas"][0]["transporte_numero"] == "T1"
    original = deepcopy(state)
    try:
        service.assign("T2", ["E2", "E1"])
    except ValueError:
        pass
    else:
        raise AssertionError("Order assigned to two trips")
    assert state == original
    failure = True
    try:
        service.assign("T1", ["E2"])
    except RuntimeError:
        pass
    else:
        raise AssertionError("Write failure hidden")
    assert state == original
    failure = False
    service.apply_suggested_cost("T1")
    assert state["transportes"][0]["custo_previsto"] == 15
    assert state["transportes"][0]["paragens"][0]["custo_transporte"] == 15
    assert "main" not in sys.modules
    print("transport-assignments-ok batch-validation=yes duplicates=yes collisions=yes recovery=yes costs=yes")


if __name__ == "__main__":
    main()
