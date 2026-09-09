"""Assembly model editing and catalog CRUD without the quote page."""
from copy import deepcopy
from dataclasses import fields
import os
from pathlib import Path
import sys
from unittest.mock import patch
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QInputDialog, QMessageBox, QPushButton, QWidget
from lugest_modules.quotes.presentation.assembly_models import AssemblyModelEditorPorts, AssemblyModelManagerPorts, edit_assembly_model, manage_assembly_models

from lugest_modules.quotes.presentation.saved_assemblies import SavedAssemblyPorts, manage_saved_assemblies

from lugest_modules.quotes.presentation.group_editor import GroupEditorPorts, save_group


def main():
    app = QApplication.instance() or QApplication([])
    owner = QWidget()
    def forbidden(*args, **kwargs):
        raise AssertionError("Unexpected interaction")
    editor = AssemblyModelEditorPorts(
        lambda row: {"kind": "product", "line": deepcopy(row)}, lambda row: "product",
        lambda kind: "Produto", forbidden, forbidden, lambda row: False, forbidden,
    )
    initial = {"codigo": "CJ1", "descricao": "Modelo", "itens": [{"descricao": "Produto", "qtd": 2, "preco_unit": 3}]}
    original = deepcopy(initial)
    with patch.object(QDialog, "exec", lambda dialog: QDialog.Rejected):
        assert edit_assembly_model(owner, editor, initial) is None
    with patch.object(QDialog, "exec", lambda dialog: QDialog.Accepted):
        result = edit_assembly_model(owner, editor, initial)
    assert result["itens"] == initial["itens"] and initial == original
    records = {"CJ1": deepcopy(initial)}
    events = []
    def save(payload):
        events.append("save")
        records[payload["codigo"]] = deepcopy(payload)
    def remove(code):
        events.append("remove")
        records.pop(code)
    ports = AssemblyModelManagerPorts(lambda: list(records.values()), lambda code: deepcopy(records[code]),
                                      save, remove, lambda initial=None: dict(initial or original))
    errors = []
    def interact(dialog):
        buttons = {button.text(): button for button in dialog.findChildren(QPushButton)}
        buttons["Editar"].click()
        buttons["Remover"].click()
        buttons["Novo"].click()
        return QDialog.Rejected
    with patch.object(QDialog, "exec", interact), patch.object(QMessageBox, "question", lambda *args: QMessageBox.Yes), \
         patch.object(QMessageBox, "critical", forbidden), patch.object(sys, "excepthook", lambda *args: errors.append(args)):
        manage_assembly_models(owner, ports)
    assert not errors and events == ["save", "remove", "save"] and "CJ1" in records
    assert "main" not in sys.modules and "lugest_modules.quotes.presentation.page" not in sys.modules
    applied = []
    catalog = SavedAssemblyPorts(lambda: [dict(original, itens=1)], lambda code: deepcopy(original),
                                 lambda code, qty: [{"descricao": "Expanded", "qtd": qty}],
                                 lambda code: events.append("catalog-remove"),
                                 lambda code: "report.pdf", lambda initial=None: None,
                                 lambda rows: applied.extend(rows))
    def apply_from_catalog(dialog):
        next(button for button in dialog.findChildren(QPushButton) if button.text().startswith("Adicionar ao")).click()
        return QDialog.Rejected
    with patch.object(QDialog, "exec", apply_from_catalog), \
         patch.object(QInputDialog, "getDouble", lambda *args: (2.0, True)), \
         patch.object(QMessageBox, "information", lambda *args: None), \
         patch.object(sys, "excepthook", lambda *args: errors.append(args)):
        manage_saved_assemblies(owner, catalog)
    assert not errors and applied == [{"descricao": "Expanded", "qtd": 2.0}]
    source_lines = [{"tipo_item": "service", "descricao": "Montagem", "qtd": 2, "preco_unit": 5}]
    before_lines = deepcopy(source_lines)
    values = {field.name: (lambda *args: None) for field in fields(GroupEditorPorts)}
    values.update(norm_text=str.lower, line_type_label=lambda row: "Servico", save_assembly_pair=lambda *args: events.append("pair"))
    def choose_group(dialog):
        combo = next(widget for widget in dialog.findChildren(QComboBox) if widget.findData("both") >= 0)
        combo.setCurrentIndex(combo.findData("both"))
        return QDialog.Accepted
    with patch.object(QDialog, "exec", lambda dialog: QDialog.Rejected):
        assert save_group(owner, GroupEditorPorts(**values), source_lines, [0], "O1") is None
    with patch.object(QDialog, "exec", choose_group):
        group = save_group(owner, GroupEditorPorts(**values), source_lines, [0], "O1")
    assert group.lines[0]["conjunto_codigo"] == group.code and source_lines == before_lines
    assert events[-1] == "pair"
    def fail_save(*args):
        raise RuntimeError("Failed pair")
    values["save_assembly_pair"] = fail_save
    messages = []
    with patch.object(QDialog, "exec", choose_group), patch.object(QMessageBox, "critical", lambda *args: messages.append(args)):
        assert save_group(owner, GroupEditorPorts(**values), source_lines, [0], "O1") is None
    assert len(messages) == 1 and source_lines == before_lines
    owner.close()
    app.processEvents()
    print("assembly-model-editor-ok independent=yes cancel=yes accept=yes catalog-crud=yes")


if __name__ == "__main__":
    main()
