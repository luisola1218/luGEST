"""Build the calculated assembly dialog using a plain QWidget owner."""
from copy import deepcopy
from dataclasses import fields
import os
from pathlib import Path
import sys
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QWidget
from lugest_modules.quotes.presentation.calculated_assembly_editor import CalculatedAssemblyPorts, edit_calculated_assembly


def main():
    app = QApplication.instance() or QApplication([])
    owner = QWidget()
    saved = []

    def forbidden(*args, **kwargs):
        raise AssertionError("Unexpected editor action")

    values = {field.name: forbidden for field in fields(CalculatedAssemblyPorts)}
    values.update(
        quote_number="O1", workcenter="Montagem", conjunto_next_param_codigo=lambda: "0001",
        save_pair=lambda template, live: saved.append((deepcopy(template), deepcopy(live))),
        norm_text=lambda value: value.lower(), orc_line_is_piece=lambda line: False,
        _quote_pick_workcenter=lambda *args: "Montagem",
        _wrap_assembly_item=lambda line: {"kind": "product", "line": deepcopy(line), "total_cost": 10},
    )
    ports = CalculatedAssemblyPorts(**values)
    initial = {"codigo": "CJ1", "descricao": "Conjunto teste", "margem_perc": 20,
               "itens": [{"descricao": "Produto", "qtd": 2, "preco_unit": 5}]}
    original = deepcopy(initial)
    errors = []
    with patch.object(sys, "excepthook", lambda *args: errors.append(args)), \
         patch.object(QMessageBox, "critical", forbidden), patch.object(QMessageBox, "warning", forbidden):
        with patch.object(QDialog, "exec", lambda dialog: QDialog.Rejected):
            assert edit_calculated_assembly(owner, ports, initial) is None
        assert not saved
        with patch.object(QDialog, "exec", lambda dialog: QDialog.Accepted):
            result = edit_calculated_assembly(owner, ports, initial)
        assert result["assembly_code"] == "CJ1" and result["workcenter"] == "Montagem"
        assert len(saved) == 1
        assert saved[0][1]["total_custo"] == 10 and saved[0][1]["total_final"] == 12
    assert not errors and initial == original
    assert "main" not in sys.modules and "lugest_modules.quotes.presentation.page" not in sys.modules
    owner.close()
    app.processEvents()
    print("calculated-assembly-editor-ok standalone=yes cancel-no-save=yes accept=yes totals=yes no-page=yes")


if __name__ == "__main__":
    main()
