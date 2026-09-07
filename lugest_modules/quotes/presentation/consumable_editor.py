from __future__ import annotations
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from lugest_modules.quotes.domain.lines import service_line as build_service_line
def edit_consumable(owner: QWidget, backend: QuoteEditorPorts, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
    initial = dict(initial or {})
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Item consumivel")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    desc_edit = QLineEdit(str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip())
    op_combo = QComboBox()
    op_combo.addItems(["Pintura", "Serralharia", "Montagem", "Lacagem"])
    op_combo.setCurrentText(str(initial.get("operacao", "Pintura") or "Pintura"))
    qty_spin = QDoubleSpinBox()
    qty_spin.setRange(0.01, 100000.0)
    qty_spin.setDecimals(2)
    qty_spin.setValue(float(initial.get("quantity_units", initial.get("qtd", 1)) or 1))
    unit_edit = QLineEdit(str(initial.get("produto_unid", "un") or "un"))
    unit_price_spin = QDoubleSpinBox()
    unit_price_spin.setRange(0.0, 1000000.0)
    unit_price_spin.setDecimals(4)
    unit_price_spin.setPrefix("EUR ")
    unit_price_spin.setValue(float(initial.get("unit_price", initial.get("preco_unit", 0)) or 0))
    form.addRow("Descricao", desc_edit)
    form.addRow("Operacao", op_combo)
    form.addRow("Quantidade", qty_spin)
    form.addRow("Unidade", unit_edit)
    form.addRow("Preco unitario", unit_price_spin)
    layout.addLayout(form)
    total_label = QLabel("")
    layout.addWidget(total_label)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    def _refresh() -> None:
        total_label.setText(f"Total: {_fmt_eur(float(qty_spin.value() or 0) * float(unit_price_spin.value() or 0))}")

    qty_spin.valueChanged.connect(lambda _v: _refresh())
    unit_price_spin.valueChanged.connect(lambda _v: _refresh())
    _refresh()
    if dialog.exec() != QDialog.Accepted:
        return None
    line = partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(
        desc_edit.text().strip() or "Consumivel",
        float(qty_spin.value() or 0),
        unit_edit.text().strip() or "un",
        float(unit_price_spin.value() or 0),
        op_combo.currentText().strip() or "Pintura",
    )
    if not isinstance(line, dict):
        return None
    return {
        "kind": "consumable",
        "descricao_base": desc_edit.text().strip(),
        "quantity_units": round(float(qty_spin.value() or 0), 2),
        "unit_price": round(float(unit_price_spin.value() or 0), 4),
        "total_cost": round(float(qty_spin.value() or 0) * float(unit_price_spin.value() or 0), 2),
        "line": line,
    }
