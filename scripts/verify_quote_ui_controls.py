from __future__ import annotations

import os
import sys
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

    from PySide6.QtCore import Qt, QTimer
    from PySide6.QtWidgets import QApplication, QCheckBox, QDialog, QLineEdit, QMessageBox, QPushButton, QSpinBox, QTableWidget

    from lugest_qt.services.legacy_backend import LegacyBackend
    from lugest_qt.ui.pages.laser_batch_quote_dialog import LaserBatchQuoteDialog
    from lugest_qt.ui.pages.laser_quote_dialogs import LaserQuoteDialog, MaterialSubtypeCatalogDialog
    from lugest_qt.ui.pages.materials_page import _NumericTableWidgetItem
    from lugest_qt.ui.pages.runtime_pages import QuotesPage
    from lugest_qt.ui.theme import apply_theme

    app = QApplication.instance() or QApplication(["verify-quote-ui-controls"])
    apply_theme(app, {})
    theme_css = app.styleSheet()
    assert "border-left: 3px solid #7ed321" not in theme_css
    assert "QTableWidget::item:selected" in theme_css
    assert "background: #f1f8eb" in theme_css
    assert "border-bottom: 1px solid #d8e8ca" in theme_css
    backend = LegacyBackend()
    page = QuotesPage(backend)
    page.line_rows = [
        {
            "tipo_item": backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
            "produto_codigo": "PRD-TEST-1",
            "descricao": "Produto de teste 1",
            "produto_unid": "UN",
            "qtd": 2,
            "preco_unit": 3.5,
        },
        {
            "tipo_item": backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
            "produto_codigo": "PRD-TEST-2",
            "descricao": "Produto de teste 2",
            "produto_unid": "UN",
            "qtd": 5,
            "preco_unit": 4.0,
        },
    ]
    page._render_quote_lines()
    app.processEvents()

    assert page.lines_table.columnCount() == 14
    assert page.lines_table.horizontalHeaderItem(page.LINE_COL_MARK).text() == "Apagar"
    first_mark = page.lines_table.item(0, page.LINE_COL_MARK)
    assert first_mark is not None and bool(first_mark.flags() & Qt.ItemIsUserCheckable)
    first_mark.setCheckState(Qt.Checked)
    app.processEvents()
    assert page._checked_line_indexes() == [0]
    assert page.remove_quote_lines_btn.isEnabled()

    page._handle_quote_lines_sort(page.LINE_COL_QUANTITY)
    app.processEvents()
    assert page._checked_line_indexes() == [0]

    original_question = QMessageBox.question
    QMessageBox.question = staticmethod(lambda *args, **kwargs: QMessageBox.Yes)
    try:
        page._remove_line()
    finally:
        QMessageBox.question = original_question
    assert len(page.line_rows) == 1
    assert page.line_rows[0].get("produto_codigo") == "PRD-TEST-2"
    page.resize(1280, 720)
    page._show_detail()
    page.show()
    app.processEvents()
    footer = page.quote_selected_line_caption.parentWidget()
    footer_bottom = footer.mapTo(page, footer.rect().bottomLeft()).y()
    assert page.quote_lines_card.objectName() == "QuoteReferencesCard"
    assert page.lines_table.verticalScrollBarPolicy() == Qt.ScrollBarAlwaysOn
    assert footer_bottom <= page.height()

    dialog = LaserBatchQuoteDialog(backend)
    dialog.show()
    app.processEvents()
    app.processEvents()
    assert bool(dialog.windowState() & Qt.WindowMaximized)
    dialog.batch_table.insertRow(0)
    quantity = dialog._make_quantity_spin()
    dialog.batch_table.setCellWidget(0, dialog.COL_QUANTITY, quantity)
    quantity.setValue(123456)
    app.processEvents()
    assert isinstance(quantity, QSpinBox)
    assert dialog._row_quantity(0) == 123456
    assert dialog.batch_table.columnWidth(dialog.COL_QUANTITY) >= 76
    assert dialog.batch_table.verticalHeader().defaultSectionSize() >= 38
    assert dialog.batch_table.verticalHeader().defaultSectionSize() >= 46
    assert "Segoe UI Semibold" in page.header_total_label.styleSheet()
    assert "font-weight: 600" in page.header_total_label.styleSheet()

    direct_quantity = dialog._make_quantity_spin(12)
    operation_button = QPushButton("Corte Laser")
    operation_button.setObjectName("BatchOperationButton")
    dialog.batch_table.setCellWidget(0, dialog.COL_QUANTITY, direct_quantity)
    dialog.batch_table.setCellWidget(0, dialog.COL_OPERATIONS, operation_button)
    app.processEvents()
    assert dialog.batch_table.cellWidget(0, dialog.COL_QUANTITY) is direct_quantity
    assert dialog.batch_table.cellWidget(0, dialog.COL_OPERATIONS) is operation_button
    assert direct_quantity.parentWidget() is dialog.batch_table.viewport()
    assert operation_button.parentWidget() is dialog.batch_table.viewport()
    assert "BatchCellControlHost" not in dialog.batch_table.styleSheet()
    assert dialog.batch_table.rowHeight(0) >= direct_quantity.sizeHint().height() + 10
    assert dialog.batch_table.rowHeight(0) >= operation_button.sizeHint().height() + 10
    assert dialog._row_quantity(0) == 12

    operation_layout_checked: list[bool] = []

    def inspect_operation_editor() -> None:
        operation_dialog = next(
            widget
            for widget in QApplication.topLevelWidgets()
            if isinstance(widget, QDialog) and widget.windowTitle() == "Operações adicionais da peça"
        )
        operation_table = operation_dialog.findChild(QTableWidget)
        assert operation_table is not None
        assert operation_table.columnCount() == 8
        assert operation_table.verticalHeader().defaultSectionSize() >= 46
        assert isinstance(operation_table.cellWidget(0, 0), QCheckBox)
        assert "spin-chevron-up.svg" in operation_table.styleSheet()
        assert "spin-chevron-down.svg" in operation_table.styleSheet()
        operation_layout_checked.append(True)
        operation_dialog.reject()

    QTimer.singleShot(0, inspect_operation_editor)
    dialog._edit_row_operations(0)
    assert operation_layout_checked == [True]

    numeric_table = QTableWidget(0, 1)
    numeric_table.setSortingEnabled(False)
    for text_value in ("10", "2", "0,8", "20"):
        row = numeric_table.rowCount()
        numeric_table.insertRow(row)
        numeric_table.setItem(row, 0, _NumericTableWidgetItem(text_value))
    numeric_table.setSortingEnabled(True)
    numeric_table.sortItems(0, Qt.AscendingOrder)
    assert [numeric_table.item(row, 0).text() for row in range(4)] == ["0,8", "2", "10", "20"]

    subtype_editor_dialog = MaterialSubtypeCatalogDialog(
        "Ferro",
        {"S235JR": {"price_per_kg": 1.2, "density_kg_m3": 7850, "scrap_credit_per_kg": 1.1}},
        1.0,
        1.0,
    )
    subtype_editor_dialog.show()
    app.processEvents()
    editable_item = subtype_editor_dialog.table.item(0, 3)
    subtype_editor_dialog.table.setCurrentCell(0, 3)
    subtype_editor_dialog.table.editItem(editable_item)
    app.processEvents()
    active_cell_editor = subtype_editor_dialog.table.findChild(QLineEdit, "SubtypeCellEditor")
    assert active_cell_editor is not None
    assert active_cell_editor.text() == "1,2"
    assert active_cell_editor.height() >= subtype_editor_dialog.table.rowHeight(0) - 6
    assert "color: #101828" in active_cell_editor.styleSheet()
    active_cell_editor.setText("2,35")
    assert active_cell_editor.text() == "2,35"
    subtype_editor_dialog.close()

    dxf_line = {
        "tipo_item": backend.desktop_main.ORC_LINE_TYPE_PIECE,
        "descricao": "Chapa DXF",
        "ref_externa": "DXF-001",
        "material": "S235JR",
        "material_family": "Aco carbono",
        "material_subtype": "S235JR",
        "espessura": "2",
        "qtd": 8,
        "operacao": "Corte Laser",
        "desenho": "C:/desenhos/DXF-001.dxf",
        "laser_base_active": True,
        "laser_snapshot": {
            "material": {"family": "Aco carbono", "subtype": "S235JR"},
            "cutting": {"gas": "Oxigenio", "thickness_mm": 2},
        },
    }
    assert page._quote_line_is_laser_2d(dxf_line)
    assert not page._quote_line_is_laser_2d({**dxf_line, "desenho": "perfil.step"})
    batch_editor = LaserBatchQuoteDialog(backend, initial_lines=[dxf_line])
    assert batch_editor.windowTitle() == "Editar lote DXF/DWG"
    assert batch_editor.batch_table.rowCount() == 1
    assert batch_editor._row_path(0) == dxf_line["desenho"]
    assert batch_editor._row_quantity(0) == 8
    assert batch_editor._row_text(0, batch_editor.COL_REFERENCE) == "DXF-001"
    batch_editor.close()

    routed_batches: list[list[dict]] = []
    original_batch_editor = page._edit_laser_batch_lines
    page._edit_laser_batch_lines = lambda rows, parent=None: (routed_batches.append(rows), rows)[1]

    def exercise_calculated_assembly_edit() -> None:
        assembly_dialog = next(
            widget
            for widget in QApplication.topLevelWidgets()
            if isinstance(widget, QDialog) and widget.windowTitle() == "Conjunto calculado"
        )
        item_table = assembly_dialog.findChild(QTableWidget)
        assert item_table is not None and item_table.rowCount() == 1
        item_table.selectRow(0)
        edit_button = next(
            button
            for button in assembly_dialog.findChildren(QPushButton)
            if button.text() == "Editar item"
        )
        edit_button.click()
        assembly_dialog.reject()

    QTimer.singleShot(0, exercise_calculated_assembly_edit)
    try:
        assert page._calculated_assembly_builder_dialog(
            {"codigo": "CJ-TEST", "descricao": "Teste", "itens": [dxf_line]}
        ) is None
    finally:
        page._edit_laser_batch_lines = original_batch_editor
    assert len(routed_batches) == 1
    assert routed_batches[0][0]["ref_externa"] == "DXF-001"

    laser_editor = LaserQuoteDialog(backend, initial_line=dxf_line)
    assert laser_editor.windowTitle() == "Editar peça DXF/DWG"
    assert laser_editor.file_edit.text() == dxf_line["desenho"]
    assert laser_editor.ref_ext_edit.text() == "DXF-001"
    assert int(laser_editor.quantity_spin.value()) == 8
    assert abs(laser_editor.thickness_spin.value() - 2.0) < 0.001
    laser_editor.close()

    dialog.close()
    page.close()
    print("quote-ui-controls-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
