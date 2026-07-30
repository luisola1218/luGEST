from __future__ import annotations

import os
import sys
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QComboBox

from lugest_qt.app import _FullSurfaceComboFilter
from lugest_qt.services.main_bridge import LegacyBackend
from lugest_qt.ui.pages.products_page import ProductsPage
from lugest_core.intelligence import (
    RemoteProductAIClient,
    extract_product_attributes,
    normalize_product_description,
    product_similarity,
    sanitize_gemini_interaction_response,
    sanitize_product_ai_response,
)


def _backend_for_ui() -> LegacyBackend:
    backend = LegacyBackend.__new__(LegacyBackend)
    backend.desktop_main = SimpleNamespace(
        PROD_UNIDS=["UN", "M", "KG", "L"],
        MATERIAIS_PRESET=["S235JR", "S355JR", "AISI 304"],
        ESPESSURAS_PRESET=[1, 1.5, 2, 3, 4, 5, 6, 8, 10, 12, 15, 20],
        LOCALIZACOES_PRESET=["RACK01", "RETALHO"],
        MATERIA_FORMATOS=["Chapa", "Tubo", "Perfil", "Cantoneira", "Barra", "Varão nervurado"],
        peek_next_produto_numero=lambda _data: "PRD-TESTE",
        produto_preco_unitario=lambda _row: 0.0,
        parse_float=lambda value, default=0.0: float(value or default),
        fmt_num=lambda value: f"{float(value or 0):g}",
    )
    backend.materia_actions = SimpleNamespace()
    backend.ensure_data = lambda: {}
    return backend


def main() -> int:
    backend = _backend_for_ui()
    cases = {
        "Parafuso C/ Cilindrica Fenda Inox M6x35": ("Fixacao", "Parafusos", "Fenda", "M6x35"),
        "Parafuso cabeça Torx TX25 M5x20": ("Fixacao", "Parafusos", "Torx (TX)", "M5x20"),
        "Parafuso Umbrako inox M8x30": ("Fixacao", "Parafusos", "Allen / Umbrako", "M8x30"),
        "Porca travante Nyloc M10": ("Fixacao", "Porcas", "Travante / Nyloc", "M10"),
        "Porcas Auto Bloqueio Inox M5": ("Fixacao", "Porcas", "Travante / Nyloc", "M5"),
        "Porca flangeada zincada M12": ("Fixacao", "Porcas", "Flangeada", "M12"),
        "Anilha de pressão inox M8": ("Fixacao", "Anilhas", "Pressao", "M8"),
        "Rebite POP alumínio 4x12": ("Fixacao", "Rebites", "Cego / POP", ""),
        "Bucha química M10": ("Fixacao", "Buchas e ancoragens", "Quimica", "M10"),
        "Esmalte martelado RAL 5003": ("Tintas / Quimicos", "Tintas e revestimentos", "Esmalte", ""),
        "Sensor indutivo M18": ("Eletronica", "Sensores", "Indutivo", "M18"),
        "Bucim Inox PG9": ("Eletronica", "Cablagem", "Bucim / Prensa-cabos", ""),
        "Cilindro pneumático compacto": ("Pneumatica", "Cilindros", "Compacto", ""),
        "Mangueira hidráulica alta pressão": ("Hidraulica", "Mangueiras", "Alta pressao", ""),
        "Luva nitrilo descartável": ("EPIs", "Luvas", "Nitrilo", ""),
        "Broca HSS 10 mm": ("Maquinacao", "Brocas", "HSS", ""),
        "Rolamento de esferas 6204": ("Rolamentos & Transmissao", "Rolamentos", "Esferas", ""),
        "Motor trifásico 1.5 kW": ("Motores & Redutores", "Motores", "Trifasico", ""),
        "O-ring NBR 30x2": ("Vedacao & Borracha", "Juntas", "O-ring", ""),
        "Chapa PVC 10 mm": ("Plasticos Tecnicos", "Chapa", "PVC", ""),
        "Caderno A4 pautado capa preta": (
            "Escritorio & Papelaria",
            "Cadernos e blocos",
            "Caderno",
            "",
        ),
        "Rato USB": (
            "Informatica",
            "Perifericos de computador",
            "Rato",
            "",
        ),
        "IPAD": (
            "Informatica",
            "Tablets e dispositivos moveis",
            "iPad",
            "",
        ),
        "Roda Azul c/Travão 100x34mm Tente 3477 UFR 100 P62": (
            "Movimentacao",
            "Rodizios industriais",
            "Giratorio com travao total",
            "100x34 mm",
        ),
    }
    for description, expected in cases.items():
        suggestion = backend.product_catalog_suggestion(description)
        actual = (
            suggestion["categoria"],
            suggestion["subcat"],
            suggestion["tipo"],
            suggestion["dimensoes"],
        )
        assert actual == expected, (description, actual, expected)

    attributes = extract_product_attributes("Parafuso cabeça cilíndrica Torx Inox A2 M6x35 DIN 912")
    assert attributes["medida"] == "M6x35"
    assert attributes["material"] == "Inox A2 / AISI 304"
    assert attributes["acionamento"] == "Torx"
    assert attributes["cabeca"] == "Cilíndrica"
    assert attributes["norma"] == "DIN 912"
    assert normalize_product_description("  Porcas Auto Bloqueio inox m5  ") == "Porca autoblocante Inox M5"
    assert product_similarity("Porcas Auto Bloqueio Inox M5", "Porca autoblocante inox M5") >= 0.95
    caster_attributes = extract_product_attributes(
        "Roda Azul c/Travão 100x34mm Tente 3477 UFR 100 P62"
    )
    assert caster_attributes["fabricante"] == "TENTE"
    assert caster_attributes["modelo"] == "3477UFR100P62"
    assert caster_attributes["medida"] == "100x34 mm"
    assert caster_attributes["cor"] == "Azul"
    assert caster_attributes["travagem"] == "Bloqueio total"
    remote = sanitize_product_ai_response(
        {
            "candidate": {
                "category": "Movimentacao",
                "subcategory": "Rodizios industriais",
                "type": "Giratorio com travao total",
                "normalized_description": "Rodízio TENTE 3477UFR100P62",
                "manufacturer": "TENTE",
                "model": "3477UFR100P62",
                "dimensions": "100x34 mm",
                "confidence": 94,
            },
            "sources": [
                {"title": "Ficha técnica TENTE", "url": "https://www.tente.com/example"},
                {"title": "Fonte inválida", "url": "file:///tmp/item.pdf"},
            ],
        }
    )
    assert remote["candidate"]["confidence"] == 0.94
    assert len(remote["sources"]) == 1

    class _FakeOllamaResponse:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self, _limit):
            return json.dumps(
                {
                    "response": json.dumps(
                        {
                            "categoria": "Eletronica",
                            "subcat": "Cablagem",
                            "tipo": "Bucim / Prensa-cabos",
                            "descricao_normalizada": "Bucim Inox PG9",
                            "fabricante": "",
                            "modelo": "PG9",
                            "dimensoes": "PG9",
                            "atributos": {"material": "Inox"},
                            "confidence": 0.91,
                        }
                    )
                }
            ).encode("utf-8")

    ollama_client = RemoteProductAIClient(
        ollama_url="http://127.0.0.1:11434",
        ollama_model="qwen3:4b",
    )
    with patch(
        "lugest_core.intelligence.remote_product_ai.urllib.request.urlopen",
        return_value=_FakeOllamaResponse(),
    ) as mocked_ollama:
        ollama = ollama_client.lookup(
            "Bucim Inox PG9",
            local_candidate={"categoria": "Eletronica"},
            taxonomy={"categories": []},
        )
    assert ollama["engine"] == "ollama-qwen3:4b"
    assert ollama["candidate"]["modelo"] == "PG9"
    sent = json.loads(mocked_ollama.call_args.args[0].data.decode("utf-8"))
    assert sent["model"] == "qwen3:4b"
    assert sent["stream"] is False
    assert sent["think"] is False
    assert sent["format"]["required"]

    gemini = sanitize_gemini_interaction_response(
        {
            "id": "gemini-test",
            "steps": [
                {
                    "type": "model_output",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                '{"categoria":"Movimentacao","subcat":"Rodizios industriais",'
                                '"tipo":"Giratorio com travao total",'
                                '"descricao_normalizada":"Rodízio TENTE 3477UFR100P62",'
                                '"fabricante":"TENTE","modelo":"3477UFR100P62",'
                                '"dimensoes":"100x34 mm","atributos":{"cor":"Azul"},'
                                '"confidence":0.96}'
                            ),
                            "annotations": [
                                {
                                    "type": "url_citation",
                                    "title": "TENTE",
                                    "url": "https://www.tente.com/example",
                                }
                            ],
                        }
                    ],
                }
            ],
        }
    )
    assert gemini["candidate"]["modelo"] == "3477UFR100P62"
    assert gemini["sources"][0]["title"] == "TENTE"

    backend.ensure_data = lambda: {
        "produtos": [
            {
                "codigo": "PRD-EXISTE",
                "descricao": "Porca autoblocante Inox M5",
            }
        ]
    }
    duplicate_analysis = backend.product_copilot_analysis("Porcas Auto Bloqueio Inox M5", "PRD-NOVO")
    assert duplicate_analysis["duplicate_warning"]
    assert duplicate_analysis["semelhantes"][0]["codigo"] == "PRD-EXISTE"
    own_analysis = backend.product_copilot_analysis("Porca autoblocante Inox M5", "PRD-EXISTE")
    assert not own_analysis["semelhantes"]
    backend.ensure_data = lambda: {}

    unknown = backend.product_catalog_suggestion("Acoplador especial ZK42")
    assert unknown["categoria"] == "Outros"
    assert unknown["needs_learning"]
    assert unknown["learning_keyword"] == "acoplador"
    learned_config: dict = {}
    backend._load_qt_config = lambda: dict(learned_config)

    def _save_learned(payload: dict) -> dict:
        learned_config.clear()
        learned_config.update(payload)
        return dict(payload)

    backend._save_qt_config = _save_learned
    learned = backend.product_catalog_teach(
        "Acoplador especial ZK42",
        "Rolamentos & Transmissao",
        "Pinhoes e polias",
        "Acoplador",
    )
    assert learned["keyword"] == "acoplador"
    learned_suggestion = backend.product_catalog_suggestion("Acoplador industrial ZK55")
    assert learned_suggestion["learned"]
    assert learned_suggestion["categoria"] == "Rolamentos & Transmissao"
    assert learned_suggestion["tipo"] == "Acoplador"
    learned_options = backend.product_catalog_options("Rolamentos & Transmissao", "Pinhoes e polias")
    assert "Acoplador" in learned_options["tipos"]
    dynamic = backend.product_catalog_teach(
        "Piso vinílico industrial",
        "Pavimentos",
        "Pavimentos vinílicos",
        "Rolo vinílico",
    )
    assert dynamic["created_category"]
    assert dynamic["created_subcategory"]
    assert dynamic["created_type"]
    dynamic_options = backend.product_catalog_options("Pavimentos", "Pavimentos vinílicos")
    assert "Rolo vinílico" in dynamic_options["tipos"]

    material_candidate = backend._local_material_command_candidate(
        "Cria-me um stock de chapa S235JR 15mm formato 3000x1500 "
        "10 unidades lote externo: 9288X20029"
    )
    assert material_candidate["formato"] == "Chapa"
    assert material_candidate["material"] == "S235JR"
    assert material_candidate["espessura"] == "15"
    assert material_candidate["comprimento"] == "3000"
    assert material_candidate["largura"] == "1500"
    assert material_candidate["quantidade"] == "10"
    assert material_candidate["lote_fornecedor"] == "9288X20029"

    normalized = backend._product_normalize_payload(
        {"codigo": "PRD-AUTO", "descricao": "Parafuso C/ Cilindrica Fenda Inox M6x35"}
    )
    assert normalized["categoria"] == "Fixacao"
    assert normalized["subcat"] == "Parafusos"
    assert normalized["tipo"] == "Fenda"
    assert normalized["dimensoes"] == "M6x35"
    manual = backend._product_normalize_payload(
        {
            "codigo": "PRD-MANUAL",
            "descricao": "Parafuso C/ Cilindrica Fenda Inox M6x35",
            "categoria": "Outros",
            "subcat": "Outros",
            "tipo": "Outros",
        }
    )
    assert (manual["categoria"], manual["subcat"], manual["tipo"]) == ("Outros", "Outros", "Outros")

    options = backend.product_catalog_options("Fixacao", "Parafusos")
    expected_drives = {
        "Fenda",
        "Phillips / Cruz (PH)",
        "Pozidriv (PZ)",
        "Torx (TX)",
        "Allen / Umbrako",
        "Sextavado exterior",
    }
    assert expected_drives.issubset(set(options["tipos"]))

    app = QApplication.instance() or QApplication(sys.argv)
    click_filter = _FullSurfaceComboFilter(app)
    app.installEventFilter(click_filter)

    class TrackingCombo(QComboBox):
        def __init__(self) -> None:
            super().__init__()
            self.popup_count = 0

        def showPopup(self) -> None:  # type: ignore[override]
            self.popup_count += 1

    click_combo = TrackingCombo()
    click_combo.setEditable(True)
    click_combo.addItems(["Fenda", "Torx (TX)"])
    click_combo.show()
    app.processEvents()
    assert click_combo.lineEdit() is not None
    QTest.mouseClick(click_combo.lineEdit(), Qt.LeftButton)
    app.processEvents()
    assert click_combo.popup_count == 1
    click_combo.close()

    page = ProductsPage(backend)
    page.desc_edit.setText("Parafuso C/ Cilindrica Fenda Inox M6x35")
    page._apply_catalog_suggestion()
    assert page.category_combo.currentText() == "Fixacao"
    assert page.subcat_combo.currentText() == "Parafusos"
    assert page.type_combo.currentText() == "Fenda"
    assert page.dim_edit.text() == "M6x35"
    assert not page.catalog_suggestion_label.isHidden()
    page._new_product()
    page.desc_edit.setText("Porcas Auto Bloqueio Inox M5")
    page._apply_catalog_suggestion()
    assert page.category_combo.currentText() == "Fixacao"
    assert page.subcat_combo.currentText() == "Porcas"
    assert page.type_combo.currentText() == "Travante / Nyloc"
    assert page.dim_edit.text() == "M5"
    page._new_product()
    page.desc_edit.setText("Roda Azul c/Travão 100x34mm Tente 3477 UFR 100 P62")
    page._apply_catalog_suggestion()
    assert page.category_combo.currentText() == "Movimentacao"
    assert page.subcat_combo.currentText() == "Rodizios industriais"
    assert page.type_combo.currentText() == "Giratorio com travao total"
    assert page.dim_edit.text() == "100x34 mm"
    assert page.maker_edit.text() == "TENTE"
    assert page.model_edit.text() == "3477UFR100P62"
    assert not page.product_ai_button.isHidden()
    page._new_product()
    page.desc_edit.setText("Elemento experimental QX99")
    page._apply_catalog_suggestion()
    assert page.category_combo.currentText() == "Outros"
    assert page.subcat_combo.currentText() == "Outros"
    assert not page.teach_catalog_button.isHidden()
    page.close()
    app.processEvents()
    print(f"product-catalog-intelligence-ok cases={len(cases)} drives={len(options['tipos'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
