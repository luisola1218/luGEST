"""A reference repair may replace the snapshot during quote preparation."""
from copy import deepcopy
from pathlib import Path
import sys
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_qt.services.main_bridge import LegacyBackend


def main():
    backend = LegacyBackend.__new__(LegacyBackend)
    backend.data = {"orcamentos": []}
    backend._base_data_snapshot = {"orcamentos": []}
    persisted = []
    backend.desktop_main = SimpleNamespace(
        now_iso=lambda: "2026-09-07T12:00:00",
        mysql_upsert_orcamento_com_linhas=lambda data, note: persisted.append((data, deepcopy(note))),
    )
    backend._parse_float = lambda value, default=0: float(value or default)
    backend._normalize_workcenter_value = lambda value: ""
    backend._normalize_orc_client = lambda value: value
    backend._ref_client_code = lambda value: value
    backend._normalize_supplier_reference = lambda *args: ("", "", "")
    backend._active_client_ref_usage = lambda *args, **kwargs: (set(), set())
    backend._known_client_ref_pairs = lambda *args: set()
    backend._peek_next_orc_number = lambda: "unused"
    backend._sync_quote_piece_registry = lambda note: None
    backend.orc_detail = lambda number: next(row for row in backend.data["orcamentos"] if row["numero"] == number)

    def repair(_code):
        backend.data = deepcopy(backend.data)
        backend.data["repaired"] = True

    backend._repair_orc_ref_history = repair
    original = backend.data
    result = backend.orc_save({"numero": "TEST-SNAPSHOT", "cliente": {"codigo": "CL1"}, "linhas": []})
    assert result["numero"] == "TEST-SNAPSHOT"
    assert not original["orcamentos"]
    assert persisted[-1][0] is backend.data and backend.data["repaired"]
    backend.orc_save({"numero": "TEST-SNAPSHOT", "cliente": {"codigo": "CL1"}, "linhas": [], "nota_cliente": "updated"})
    assert len(backend.data["orcamentos"]) == 1
    assert backend.data["orcamentos"][0]["nota_cliente"] == "updated"
    verify_order_import()
    print("snapshot-save-flows-ok quote-create=yes quote-update=yes order-import=yes")


def verify_order_import():
    backend = LegacyBackend.__new__(LegacyBackend)
    backend.data = {"encomendas": [{"numero": "TEST-ORDER"}]}
    backend._parse_float = lambda value, default=0: float(value or default)
    backend._order_is_orc_based = lambda order: False
    backend.conjunto_detail = lambda code: {"descricao": "Model", "total_final": 10}
    backend.conjunto_expand = lambda code, quantity: [
        {"tipo_item": "piece", "qtd": 1},
        {"tipo_item": "service", "qtd": 1},
        {"tipo_item": "piece", "qtd": 1},
    ]
    backend.desktop_main = SimpleNamespace(
        now_iso=lambda: "2026-09-07T12:00:00",
        normalize_orc_line_type=lambda value: value,
        orc_line_is_piece=lambda row: row.get("tipo_item") == "piece",
        orc_line_is_service=lambda row: row.get("tipo_item") == "service",
        update_estado_encomenda_por_espessuras=lambda order: None,
    )

    def save(**kwargs):
        backend.data = deepcopy(backend.data)

    def add_piece(number, payload):
        backend.get_encomenda_by_numero(number).setdefault("pieces", []).append(payload)
        save()

    backend.order_piece_create_or_update = add_piece
    backend._save = save
    backend._ensure_order_fabrication_order = lambda order: None
    backend.order_detail = lambda number: deepcopy(backend.get_encomenda_by_numero(number))
    result = backend.order_import_model("TEST-ORDER", "MODEL", 2, "conjunto")
    assert len(result["pieces"]) == 2
    assert len(result["montagem_itens"]) == 1
    assert result["produto_fichas"][0]["codigo"] == "MODEL"
    assert result["valor_adjudicado"] == 20


if __name__ == "__main__":
    main()
