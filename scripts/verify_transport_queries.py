"""Transport planning reads must not repair or mutate shared order records."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.transport.application.queries import TransportQueries, TransportQueryRules
from lugest_modules.transport.infrastructure.legacy_transport_read_repository import LegacyTransportReadRepository


def main():
    state = {
        "encomendas": [
            {"numero": "E1", "cliente": "C1", "numero_orcamento": "Q1", "pecas": [{"available": 3}], "paletes": 2},
            {"numero": "E2", "nota_transporte": "Transporte a Cargo do Cliente"},
            {"numero": "E3", "nota_transporte": "Transporte a Nosso Cargo"},
            {"numero": "E4", "nota_transporte": "Transporte a Nosso Cargo", "pecas": [{"available": 1}]},
        ],
        "clientes": [{"codigo": "C1", "nome": "Cliente", "localidade": "Porto"}],
        "orcamentos": [{}, {"numero": "Q1", "nota_transporte": "Transporte a Nosso Cargo", "zona_transporte": "Lisboa"}],
        "transportes": [{"numero": "T1", "estado": "Planeado", "paragens": [{"encomenda_numero": "E4"}]}],
        "expedicoes": [],
    }
    original = deepcopy(state)
    repository = LegacyTransportReadRepository(lambda: state)
    service = TransportQueries(repository, TransportQueryRules(
        lambda value, default=0: float(value or default), str.lower,
        lambda code, name: (code, name, ""), lambda *args: {"custo_sugerido": 12},
        lambda order: order.update(estado_expedicao="Recalculado"),
        lambda order: order.setdefault("pecas", []), lambda piece: piece.get("available", 0),
        lambda stop: {}, lambda: {},
    ))
    rows = service.pending_orders()
    assert [row["numero"] for row in rows] == ["E1"]
    assert rows[0]["estado_expedicao"] == "Recalculado"
    assert rows[0]["zona_transporte"] == "Porto" and rows[0]["custo_sugerido"] == 12
    assert service.pending_orders("missing") == []
    assert service.pending_overview() == dict(total_orders=4, eligible=1, customer_transport=1, waiting_stock_or_guide=1, already_assigned=1)
    assert service.zones() == ["Lisboa", "Porto"]
    assert service.note(state["encomendas"][0]) == "Transporte a Nosso Cargo"
    repository.order("E1")["pecas"][0]["available"] = 999
    repository.trip("T1")["paragens"].clear()
    assert state == original, "Reading a transport projection changed shared data"
    summary = service.summary([{"paletes": 2, "preco_transporte": 30, "custo_transporte": 10}], {"custo_previsto": 15})
    assert summary["margem_prevista"] == 15 and summary["paletes"] == 2
    print("Transport queries: detached planning reads, eligibility and projections passed")


if __name__ == "__main__":
    main()
