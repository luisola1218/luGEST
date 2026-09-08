"""Trip progression, stop ordering, guide validation and recovery."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.transport.application.stops import TransportStops, StopRules
from lugest_modules.transport.infrastructure.legacy_trip_repository import LegacyTripRepository


def main():
    state = {"transportes": [{"numero": "T1", "estado": "Planeado", "paragens": [
        {"encomenda_numero": "E1", "ordem": 1}, {"encomenda_numero": "E2", "ordem": 2}]}],
        "encomendas": [{"numero": "E1"}, {"numero": "E2"}]}
    failure = False

    def persist(**kwargs):
        nonlocal state
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyTripRepository(lambda: state, persist, str.lower)
    service = TransportStops(repository, StopRules(
        lambda value, default=0: float(value or default), str.lower, lambda: "2026-09-08T12:00:00",
        lambda: "TEST", lambda code, name: (code, name, ""),
        lambda number: [{"numero": "G1"}] if number == "E1" else []))
    service.move("T1", "E1", 1)
    assert [stop["encomenda_numero"] for stop in state["transportes"][0]["paragens"]] == ["E2", "E1"]
    assert [stop["ordem"] for stop in state["transportes"][0]["paragens"]] == [1, 2]
    original = deepcopy(state)
    for action in (lambda: service.set_status("T1", "Em transito"),
                   lambda: service.update("T1", "E2", {"expedicao_numero": "G1"}),
                   lambda: service.set_stop_status("T1", "E1", "Entregue")):
        try:
            action()
        except ValueError:
            pass
        else:
            raise AssertionError("Invalid transition or guide accepted")
        assert state == original
    service.remove("T1", "E2")
    service.update("T1", "E1", {"expedicao_numero": "G1", "check_carga_ok": True,
                                "check_docs_ok": True, "check_paletes_ok": True,
                                "pod_estado": "Recebido"})
    service.set_status("T1", "Em transito")
    assert state["encomendas"][0]["estado_transporte"] == "Em transito"
    service.set_stop_status("T1", "E1", "Entregue")
    service.set_status("T1", "Concluida")
    assert state["transportes"][0]["estado"] == "Concluida"
    assert state["encomendas"][0]["estado_transporte"] == "Entregue"
    failure = True
    original = deepcopy(state)
    try:
        service.remove("T1", "E1")
    except RuntimeError:
        pass
    else:
        raise AssertionError("Write failure hidden")
    assert state == original
    failure = False
    service.request_service("T1", {"pedido_transporte_estado": "Confirmado"})
    assert state["transportes"][0]["pedido_confirmado_by"] == "TEST"
    service.request_service("T1", {"pedido_transporte_estado": "nao pedido"})
    assert state["transportes"][0]["pedido_confirmado_by"] == ""
    stale = repository.get("T1")
    state["transportes"][0]["observacoes"] = "Concurrent"
    try:
        repository.save(stale, expected=stale)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale trip accepted")
    assert "main" not in sys.modules
    print("transport-stops-ok reorder=yes progression=yes guide-ownership=yes order-links=yes recovery=yes")


if __name__ == "__main__":
    main()
