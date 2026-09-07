from __future__ import annotations
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_common import configure_table as _configure_table, fill_table as _fill_table
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from typing import Any
def edit_material_prices(owner: QWidget, backend: QuoteEditorPorts, formato_filter: str = "", preferred_id: str = "", parent: QWidget | None = None) -> dict | None:
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Tabela de preços MP")
    dialog.resize(980, 620)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(8)

    intro = QLabel(
        "Consulta e atualiza os preços da matéria-prima. "
        "Tubos usam EUR/m como preço base; os restantes materiais usam EUR/kg."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(intro)

    filters = QHBoxLayout()
    format_combo = QComboBox()
    format_combo.addItem("Todos", "")
    for value in ("Perfil", "Chapa", "Tubo", "Cantoneira", "Barra", "Ferro nervurado"):
        format_combo.addItem(value, value)
    if str(formato_filter or "").strip():
        for idx in range(format_combo.count()):
            if str(format_combo.itemData(idx) or "").strip().lower() == str(formato_filter or "").strip().lower():
                format_combo.setCurrentIndex(idx)
                break
    search_edit = QLineEdit()
    search_edit.setPlaceholderText("Filtrar por ID, material ou dimensão")
    filters.addWidget(format_combo)
    filters.addWidget(search_edit, 1)
    layout.addLayout(filters)

    table = QTableWidget(0, 9)
    table.setHorizontalHeaderLabels(["ID", "Formato", "Material", "Dimensão", "Esp.", "Kg/m", "Base", "EUR/kg equiv.", "EUR/unid"])
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QTableWidget.NoEditTriggers)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    _configure_table(table, stretch=(2, 3), contents=(0, 1, 4, 5, 6, 7, 8))
    layout.addWidget(table, 1)

    editor_card = CardFrame()
    editor_card.set_tone("info")
    editor_form = QFormLayout(editor_card)
    editor_form.setContentsMargins(10, 8, 10, 8)
    editor_form.setHorizontalSpacing(10)
    editor_form.setVerticalSpacing(6)
    selected_label = QLabel("-")
    base_label = QLabel("Preço base")
    base_spin = QDoubleSpinBox()
    base_spin.setRange(0.0, 1000000.0)
    base_spin.setDecimals(4)
    kg_spin = QDoubleSpinBox()
    kg_spin.setRange(0.0, 1000000.0)
    kg_spin.setDecimals(4)
    unit_label = QLabel("-")
    editor_form.addRow("Selecionado", selected_label)
    editor_form.addRow(base_label, base_spin)
    editor_form.addRow("Preço / kg equiv.", kg_spin)
    editor_form.addRow("Preço / unid.", unit_label)
    layout.addWidget(editor_card)

    actions = QHBoxLayout()
    apply_btn = QPushButton("Atualizar preço")
    apply_btn.setProperty("variant", "primary")
    refresh_btn = QPushButton("Atualizar lista")
    refresh_btn.setProperty("variant", "secondary")
    actions.addWidget(apply_btn)
    actions.addWidget(refresh_btn)
    actions.addStretch(1)
    layout.addLayout(actions)

    buttons = QDialogButtonBox(QDialogButtonBox.Close)
    buttons.rejected.connect(dialog.reject)
    buttons.accepted.connect(dialog.accept)
    layout.addWidget(buttons)

    sync = {"busy": False}
    rows_cache: list[dict[str, Any]] = []

    def current_row() -> dict[str, Any] | None:
        current = table.currentItem()
        if current is None or current.row() >= len(rows_cache):
            return None
        return dict(rows_cache[current.row()] or {})

    def refresh_table() -> None:
        rows = list(backend.material_price_rows(str(format_combo.currentData() or "").strip()) or [])
        query = search_edit.text().strip().lower()
        if query:
            rows = [
                row
                for row in rows
                if query in " ".join(
                    [
                        str(row.get("id", "") or ""),
                        str(row.get("formato", "") or ""),
                        str(row.get("material", "") or ""),
                        str(row.get("dimension_label", "") or ""),
                        str(row.get("espessura", "") or ""),
                    ]
                ).lower()
            ]
        rows_cache[:] = rows
        _fill_table(
            table,
            [
                [
                    row.get("id", "-"),
                    row.get("formato", "-"),
                    row.get("material", "-"),
                    row.get("dimension_label", "-"),
                    row.get("espessura", "-"),
                    f"{float(row.get('kg_m', 0) or 0):.4f}",
                    f"{float(row.get('p_compra', 0) or 0):.4f}",
                    f"{float(row.get('price_kg', 0) or 0):.4f}",
                    _fmt_eur(float(row.get("preco_unid", 0) or 0)),
                ]
                for row in rows
            ],
            align_center_from=4,
        )
        if preferred_id:
            for idx, row in enumerate(rows):
                if str(row.get("id", "") or "").strip() == str(preferred_id or "").strip():
                    table.selectRow(idx)
                    break
        elif rows:
            table.selectRow(0)
        _load_selected_row()

    def _load_selected_row() -> None:
        row = current_row()
        sync["busy"] = True
        try:
            if not row:
                selected_label.setText("-")
                unit_label.setText("-")
                base_label.setText("Preço base")
                base_spin.setValue(0.0)
                kg_spin.setValue(0.0)
                return
            selected_label.setText(
                f"{row.get('id', '-') } | {row.get('formato', '-') } | {row.get('material', '-') } | {row.get('dimension_label', '-') }"
            )
            base_label.setText(f"Preço base ({str(row.get('base_label', 'EUR/kg') or 'EUR/kg')})")
            base_spin.setValue(float(row.get("p_compra", 0) or 0.0))
            kg_spin.setValue(float(row.get("price_kg", 0) or 0.0))
            unit_label.setText(_fmt_eur(float(row.get("preco_unid", 0) or 0.0)))
        finally:
            sync["busy"] = False

    def _sync_from_base(value: float) -> None:
        if sync["busy"]:
            return
        row = current_row()
        if not row:
            return
        sync["busy"] = True
        try:
            if str(row.get("formato", "") or "").strip().lower() == "tubo":
                kg_m = float(row.get("kg_m", 0) or 0.0)
                kg_spin.setValue(round((float(value or 0) / kg_m), 4) if kg_m > 0 else 0.0)
            else:
                kg_spin.setValue(float(value or 0))
        finally:
            sync["busy"] = False

    def _sync_from_kg(value: float) -> None:
        if sync["busy"]:
            return
        row = current_row()
        if not row:
            return
        sync["busy"] = True
        try:
            if str(row.get("formato", "") or "").strip().lower() == "tubo":
                kg_m = float(row.get("kg_m", 0) or 0.0)
                base_spin.setValue(round(float(value or 0) * kg_m, 4) if kg_m > 0 else 0.0)
            else:
                base_spin.setValue(float(value or 0))
        finally:
            sync["busy"] = False

    def apply_price() -> None:
        row = current_row()
        if not row:
            QMessageBox.warning(dialog, "Matéria-prima", "Seleciona primeiro um material.")
            return
        try:
            if str(row.get("formato", "") or "").strip().lower() == "tubo":
                kg_m = float(row.get("kg_m", 0) or 0.0)
                if kg_m <= 0:
                    raise ValueError("Kg/m inválido para converter o preço por metro.")
                backend.material_update_price_kg(str(row.get("id", "") or "").strip(), float(base_spin.value() or 0.0) / kg_m)
            else:
                backend.material_update_price_kg(str(row.get("id", "") or "").strip(), base_spin.value())
        except Exception as exc:
            QMessageBox.critical(dialog, "Matéria-prima", str(exc))
            return
        QMessageBox.information(dialog, "Matéria-prima", "Preço atualizado no stock com sucesso.")
        refresh_table()

    table.itemSelectionChanged.connect(_load_selected_row)
    format_combo.currentTextChanged.connect(lambda _text: refresh_table())
    search_edit.textChanged.connect(lambda _text: refresh_table())
    base_spin.valueChanged.connect(_sync_from_base)
    kg_spin.valueChanged.connect(_sync_from_kg)
    apply_btn.clicked.connect(apply_price)
    refresh_btn.clicked.connect(refresh_table)
    refresh_table()
    dialog.exec()
    return current_row()

def sync_stock_price(owner: QWidget, backend: QuoteEditorPorts, material_id: str, price_kg: float, current_price_kg: float, parent: QWidget | None = None) -> dict[str, Any] | None:
    stock_id = str(material_id or "").strip()
    if not stock_id:
        return None
    new_value = round(float(price_kg or 0.0), 4)
    current_value = round(float(current_price_kg or 0.0), 4)
    if new_value <= 0 or abs(new_value - current_value) < 0.0001:
        return None
    try:
        return dict(backend.material_update_price_kg(stock_id, new_value) or {})
    except Exception as exc:
        QMessageBox.warning(parent if isinstance(parent, QWidget) else owner, "Matéria-prima", str(exc))
        return None
