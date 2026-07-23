from __future__ import annotations

import copy
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QDialog, QDoubleSpinBox, QWidget

from lugest_qt.services.legacy_backend import LegacyBackend
from lugest_qt.ui.pages.runtime_pages import OperatorPage


def _piece(piece_id: str) -> dict:
    return {
        "id": piece_id,
        "material": "TEST SESSION S235",
        "espessura": "7.77",
        "quantidade_pedida": 1,
        "produzido_ok": 0,
        "produzido_nok": 0,
        "produzido_qualidade": 0,
        "operacoes_fluxo": [
            {"nome": "Corte Laser", "estado": "Preparacao", "user": "operador.teste"}
        ],
    }


def main() -> int:
    app = QApplication.instance() or QApplication([])
    host = QWidget()
    dialog_state = {
        "material": "TEST SESSION S235",
        "espessura": "7.77",
        "total_qty": 1.0,
        "reserved_qty": 4.0,
        "manual_stock_required": False,
        "session_close": True,
        "reserved_sources": [
            {
                "material_id": "MAT-SESSION-VERIFY",
                "dimensao": "2000x1000",
                "quantidade": 4.0,
                "disponivel": 4.0,
                "local": "RACK A",
                "lote": "SESSION-LOT",
            }
        ],
        "candidates": [],
    }

    def choose_one_reserved_sheet() -> None:
        dialog = app.activeModalWidget()
        assert isinstance(dialog, QDialog)
        quantity_spins = [
            spin
            for spin in dialog.findChildren(QDoubleSpinBox)
            if "cativadas" in spin.suffix()
        ]
        assert len(quantity_spins) == 1
        quantity_spin = quantity_spins[0]
        assert quantity_spin.maximum() == 4.0
        quantity_spin.setValue(1.0)
        dialog.accept()

    QTimer.singleShot(100, choose_one_reserved_sheet)
    dialog_payload = OperatorPage._prompt_laser_stock_resolution(host, dialog_state)
    assert dialog_payload is not None
    assert dialog_payload["material_id"] == "MAT-SESSION-VERIFY"
    assert dialog_payload["quantity"] == 1.0

    backend = LegacyBackend()
    backend._replace_data_cache(copy.deepcopy(backend.ensure_data()))
    backend._save = lambda *args, **kwargs: None
    data = backend.ensure_data()
    stock = {
        "id": "MAT-SESSION-VERIFY",
        "formato": "Chapa",
        "material": "TEST SESSION S235",
        "espessura": "7.77",
        "quantidade": 5.0,
        "reservado": 4.0,
        "comprimento": 2000,
        "largura": 1000,
        "lote_interno": "SESSION-LOT",
        "lote_fornecedor": "SESSION-SUPPLIER",
    }
    data.setdefault("materiais", []).append(stock)
    pieces = [_piece("SESSION-A"), _piece("SESSION-B"), _piece("SESSION-C")]
    esp_obj = {"espessura": "7.77", "pecas": pieces}
    order = {
        "numero": "SESSION-ORDER",
        "materiais": [{"material": "TEST SESSION S235", "espessuras": [esp_obj]}],
        "reservas": [
            {
                "material_id": stock["id"],
                "material": stock["material"],
                "espessura": stock["espessura"],
                "quantidade": 4.0,
            }
        ],
        "cativar": True,
    }
    data.setdefault("encomendas", []).append(order)

    active_ids = {"SESSION-A", "SESSION-B", "SESSION-C"}

    def status_for_order(*_args, **_kwargs):
        return {
            piece_id: [
                {
                    "operacao": "Corte Laser",
                    "estado": "Em producao",
                    "operador": "operador.teste",
                }
            ]
            for piece_id in active_ids
        }

    original_status = backend.operador_actions._mysql_ops_status_for_order
    backend.operador_actions._mysql_ops_status_for_order = status_for_order
    try:
        for expected in (3, 2, 1):
            state = backend.operator_material_session_state(
                order["numero"], stock["material"], stock["espessura"], "operador.teste"
            )
            assert state["active_count"] == expected
            assert state["should_prompt"] is False
            active_ids.remove(sorted(active_ids)[0])

        state = backend.operator_material_session_state(
            order["numero"], stock["material"], stock["espessura"], "operador.teste"
        )
        assert state["active_count"] == 0
        assert state["should_prompt"] is True

        backend.operator_record_material_session_decision(
            order["numero"], stock["material"], stock["espessura"], "operador.teste", "manter_cativado", "interrupcao"
        )
        assert stock["quantidade"] == 5.0 and stock["reservado"] == 4.0
        assert order["reservas"][0]["quantidade"] == 4.0

        result = backend.operator_resolve_laser_stock(
            order["numero"],
            stock["material"],
            stock["espessura"],
            material_id=stock["id"],
            quantidade=1.0,
            session_close=True,
            operator_name="operador.teste",
        )
        assert result["resolved"] is True
        assert result["consumed_total"] == 1.0
        assert result["remaining_reserved"] == 3.0
        assert stock["quantidade"] == 4.0 and stock["reservado"] == 3.0
        assert order["reservas"][0]["quantidade"] == 3.0
        assert not esp_obj.get("laser_concluido")
        history = list(esp_obj.get("material_session_history", []) or [])
        assert [row.get("decisao") for row in history][-2:] == ["manter_cativado", "dar_baixa"]
    finally:
        backend.operador_actions._mysql_ops_status_for_order = original_status

    print("operator-material-session-flow-ok active=3>2>1>0 dialog=1/4 consume=1 remaining=3 interruption=yes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
