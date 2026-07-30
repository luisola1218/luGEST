from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from lugest_core.intelligence import (
    ERPCopilot,
    MODULE_KNOWLEDGE,
    build_erp_snapshot,
    deterministic_answer,
)
from lugest_core.intelligence.erp_copilot import _valid_chat_answer
from lugest_qt.ui.copilot_dialog import ERPCopilotDialog


DATA = {
    "materiais": [{"id": "MAT-1", "material": "S235JR", "quantidade": 0, "estado": "Crítico"}],
    "produtos": [
        {"codigo": "PRD-1", "descricao": "Parafuso", "qty": 10},
        {"codigo": "PRD-2", "descricao": "Acessório sem stock", "qty": 0},
        {"codigo": "PRD-3", "descricao": "Consumível no mínimo", "qty": 2, "alerta": 2},
    ],
    "encomendas": [
        {"numero": "OF-1", "estado": "Preparação", "data_entrega": "2020-01-01"},
    ],
    "plano": [{"id": "PL-1", "duracao_min": 120, "operacao": "Laser"}],
    "notas_encomenda": [{"numero": "NE-1", "estado": "Em edição"}],
}


class _Backend:
    def __init__(self):
        self.prepared_actions = 0
        self.executed_actions = 0

    def erp_copilot_status(self):
        return {
            "available": False,
            "message": "Modo sem IA",
            "model": "regras-locais",
        }

    def erp_copilot_ask(self, question: str, current_page: str = "", conversation=None):
        snapshot = build_erp_snapshot(DATA)
        snapshot["financeiro"] = {
            "stock_total": 31672.42,
            "stock_materias": 6292.23,
            "stock_produtos": 25380.19,
            "stock_disponivel": 31672.42,
            "stock_reservado": 0,
        }
        return ERPCopilot(ollama_url="http://127.0.0.1:1").ask(
            question,
            snapshot,
            current_page=current_page,
            conversation=conversation,
        )

    def erp_copilot_prepare_action(self, action):
        self.prepared_actions += 1
        return {"candidate": {"formato": "Chapa", "material": "S235JR"}}

    def erp_copilot_resolve_action_followup(self, question, action, conversation=None):
        return ERPCopilot(ollama_url="http://127.0.0.1:1").resolve_action_followup(
            question,
            action,
            conversation=conversation,
        )

    def erp_copilot_execute_action(self, action, confirmed_payload):
        self.executed_actions += 1
        return {"ok": True, "record": {"id": "MAT-TEST"}}


def main() -> int:
    snapshot = build_erp_snapshot(DATA)
    assert snapshot["materia_prima"]["criticos"] == 1
    assert snapshot["produtos"]["sem_stock"] == 1
    assert snapshot["produtos"]["sem_stock_ou_criticos"] == 2
    assert snapshot["encomendas"]["atrasadas"] == 1
    assert snapshot["planeamento"]["minutos_planeados"] == 120
    restricted = build_erp_snapshot(
        DATA,
        permissions={
            "materials": False,
            "products": False,
            "orders": True,
            "planning": False,
            "purchase_notes": False,
        },
    )
    assert "materia_prima" not in restricted
    assert "encomendas" in restricted
    assert "planeamento" not in restricted
    assert deterministic_answer("stock crítico", snapshot)["navigation_target"] == "materials"
    product_answer = deterministic_answer(
        "Que produtos estão sem stock?",
        snapshot,
        current_page="products",
    )
    assert product_answer["navigation_target"] == "products"
    assert "PRD-2" in product_answer["answer"]
    assert "PRD-1" not in product_answer["answer"]
    verified = ERPCopilot(ollama_url="http://127.0.0.1:1").ask(
        "Que produtos estão sem stock?",
        snapshot,
        current_page="products",
    )
    assert verified["engine"] == "dados ERP verificados"
    assert "PRD-2" in verified["answer"]
    assert deterministic_answer("encomendas atrasadas", snapshot)["navigation_target"] == "orders"
    snapshot["financeiro"] = {
        "stock_total": 31672.42,
        "stock_materias": 6292.23,
        "stock_produtos": 25380.19,
        "stock_disponivel": 31672.42,
        "stock_reservado": 0,
    }
    snapshot["modulos_lugest"] = MODULE_KNOWLEDGE
    first_financial = deterministic_answer("Qual é o meu património em stock?", snapshot)
    assert "31.672,42 EUR" in first_financial["answer"]
    follow_up = ERPCopilot(ollama_url="http://127.0.0.1:1").ask(
        "em euros?",
        snapshot,
        current_page="stock_dashboard",
        conversation=[
            {"role": "user", "content": "Qual é o meu património em stock?"},
            {"role": "assistant", "content": first_financial["answer"]},
        ],
    )
    assert follow_up["engine"] == "dados ERP verificados"
    assert "31.672,42 EUR" in follow_up["answer"]
    prepared = deterministic_answer(
        "Cria uma nova viagem para uma encomenda",
        snapshot,
        current_page="stock_dashboard",
    )
    assert prepared["intent"] == "prepare_action"
    assert prepared["requires_confirmation"] is True
    assert prepared["navigation_target"] == "transportes"
    material_action = deterministic_answer(
        "Cria um stock de 10 chapas S235JR 10 mm 3000x1500",
        snapshot,
        current_page="materials",
    )
    assert material_action["proposed_action"]["type"] == "create_material_stock"
    assert material_action["proposed_action"]["mutates_data"] is True
    global_material_action = deterministic_answer(
        "Cria um stock de 10 chapas S235JR 10 mm 3000x1500",
        snapshot,
        current_page="stock_dashboard",
    )
    assert global_material_action["navigation_target"] == "materials"
    assert global_material_action["proposed_action"]["type"] == "create_material_stock"
    assert not _valid_chat_answer("faz lá", "faz lá")
    assert _valid_chat_answer(
        "Vou preparar a ficha do lote para validação.",
        "Cria um lote.",
    )
    natural_confirmation = ERPCopilot(
        ollama_url="http://127.0.0.1:1"
    ).resolve_action_followup(
        "faz o que te pedi",
        global_material_action["proposed_action"],
    )
    assert natural_confirmation["decision"] == "execute"
    structured = ERPCopilot(ollama_url="http://127.0.0.1:1")
    structured.openai_api_key = "test-only"
    structured._openai_response = lambda _body: {
        "output": [{
            "type": "message",
            "content": [{
                "type": "output_text",
                "text": (
                    '{"intent":"weather","module":"","topic":"meteorologia",'
                    '"location":"Apúlia","date_reference":"tomorrow",'
                    '"action_type":"get_weather","command":""}'
                ),
            }],
        }]
    }
    route = structured._route_question("Qual é o tempo em Apúlia amanhã?")
    assert route["intent"] == "weather"
    assert route["location"] == "Apúlia"
    assert route["date_reference"] == "tomorrow"

    app = QApplication.instance() or QApplication([])
    backend = _Backend()
    dialog = ERPCopilotDialog(backend, current_page="orders")
    dialog.show()
    dialog.question.setPlainText("Mostrar encomendas atrasadas")
    from lugest_qt.ui.pages import materials_page

    original_material_dialog = materials_page._MaterialEditorDialog

    class _AcceptedMaterialDialog:
        def __init__(self, *_args, **_kwargs):
            pass

        def exec(self):
            from PySide6.QtWidgets import QDialog
            return QDialog.Accepted

        def payload(self):
            return {"formato": "Chapa", "material": "S235JR", "quantidade": 10}

    materials_page._MaterialEditorDialog = _AcceptedMaterialDialog

    def confirm_action():
        dialog._pending_action = {
            "type": "create_material_stock",
            "command": "Cria 10 chapas S235JR 10mm 3000x1500",
            "target": "materials",
        }
        dialog.question.setPlainText("Então?")
        dialog.ask()

    QTimer.singleShot(200, dialog.ask)
    QTimer.singleShot(2500, confirm_action)
    QTimer.singleShot(6500, dialog.close)
    QTimer.singleShot(6800, app.quit)
    app.exec()
    materials_page._MaterialEditorDialog = original_material_dialog
    assert "OF-1" in dialog.transcript.toPlainText()
    assert backend.prepared_actions == 1
    assert backend.executed_actions == 1
    assert "MAT-TEST" in dialog.transcript.toPlainText()
    print("ERP Copilot verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
