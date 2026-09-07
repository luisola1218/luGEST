from __future__ import annotations
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from lugest_modules.quotes.domain.lines import service_line as build_service_line
def edit_labor(owner: QWidget, backend: QuoteEditorPorts, initial: dict | None = None, parent: QWidget | None = None, *, presets: dict) -> dict | None:
    initial = dict(initial or {})
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Item mão de obra")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    desc_edit = QLineEdit(str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip())
    op_combo = QComboBox()
    op_combo.addItems(list(presets.get("operacoes", []) or []) or ["Serralharia", "Pintura", "Retrabalho", "Montagem"])
    op_combo.setCurrentText(str(initial.get("operacao", "Serralharia") or "Serralharia"))
    hours_spin = QDoubleSpinBox()
    hours_spin.setRange(0.01, 100000.0)
    hours_spin.setDecimals(2)
    hours_spin.setValue(float(initial.get("hours", initial.get("qtd", 1)) or 1))
    rate_spin = QDoubleSpinBox()
    rate_spin.setRange(0.0, 1000000.0)
    rate_spin.setDecimals(4)
    rate_spin.setPrefix("EUR ")
    rate_spin.setValue(float(initial.get("hour_rate", initial.get("preco_unit", 20)) or 20))
    form.addRow("Descricao", desc_edit)
    form.addRow("Servico", op_combo)
    form.addRow("Horas", hours_spin)
    form.addRow("Preco / hora", rate_spin)
    layout.addLayout(form)
    total_label = QLabel("")
    layout.addWidget(total_label)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    def _refresh() -> None:
        total_label.setText(f"Total: {_fmt_eur(float(hours_spin.value() or 0) * float(rate_spin.value() or 0))}")

    hours_spin.valueChanged.connect(lambda _v: _refresh())
    rate_spin.valueChanged.connect(lambda _v: _refresh())
    _refresh()
    if dialog.exec() != QDialog.Accepted:
        return None
    line = partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(
        desc_edit.text().strip() or op_combo.currentText().strip() or "Mao de obra",
        float(hours_spin.value() or 0),
        "h",
        float(rate_spin.value() or 0),
        op_combo.currentText().strip() or "Serralharia",
    )
    if not isinstance(line, dict):
        return None
    return {
        "kind": "labor",
        "descricao_base": desc_edit.text().strip(),
        "hours": round(float(hours_spin.value() or 0), 2),
        "hour_rate": round(float(rate_spin.value() or 0), 4),
        "total_cost": round(float(hours_spin.value() or 0) * float(rate_spin.value() or 0), 2),
        "line": line,
    }
