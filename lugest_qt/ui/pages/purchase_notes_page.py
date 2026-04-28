from __future__ import annotations

import html
import os
import re
import subprocess
import tempfile
from datetime import date
from pathlib import Path
from urllib.parse import quote

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame
from .partners_pages import SuppliersPage
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    cap_width as _cap_width,
    coerce_editor_qdate as _coerce_editor_qdate,
    configure_table as _configure_table,
    elide_middle as _elide_middle,
    fill_table as _fill_table,
    fmt_eur as _fmt_eur,
    paint_table_row as _paint_table_row,
    selected_row_index as _selected_row_index,
    set_panel_tone as _set_panel_tone,
    set_table_columns as _set_table_columns,
    state_tone as _state_tone,
    table_visible_height as _table_visible_height,
)


class PurchaseNotesPage(QWidget):
    page_title = "Notas Encomenda"
    page_subtitle = "Pedido de cotação, adjudicação por linha e notas de compra reais por fornecedor."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.line_rows: list[dict] = []
        self.supplier_rows: list[dict] = []
        self.current_number = ""
        self._rfq_outlook_to = ""
        self._rfq_outlook_cc = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        filters = CardFrame()
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(16, 14, 16, 14)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por número, fornecedor ou estado")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Em edicao", "Aprovada", "Enviada", "Parcial", "Entregue", "Convertidas", "Todas"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.state_combo)
        root.addWidget(filters)

        notes_card = CardFrame()
        notes_card.set_tone("default")
        notes_layout = QVBoxLayout(notes_card)
        notes_layout.setContentsMargins(16, 14, 16, 14)
        notes_title = QLabel("Pedidos / NEs")
        notes_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.notes_table = QTableWidget(0, 6)
        self.notes_table.setHorizontalHeaderLabels(["Número", "Fornecedor", "Entrega", "Estado", "Total", "Linhas"])
        self.notes_table.verticalHeader().setVisible(False)
        self.notes_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.notes_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.notes_table, stretch=(1,), contents=(0, 2, 3, 4, 5))
        self.notes_table.itemSelectionChanged.connect(self._load_selected_note)
        notes_layout.addWidget(notes_title)
        notes_layout.addWidget(self.notes_table)
        root.addWidget(notes_card)

        self.note_form_card = CardFrame()
        form_layout = QVBoxLayout(self.note_form_card)
        form_layout.setContentsMargins(12, 8, 12, 10)
        form_layout.setSpacing(6)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(8)
        self.new_btn = QPushButton("Novo pedido")
        self.new_btn.clicked.connect(self._create_new_note)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_note)
        self.approve_btn = QPushButton("Aprovar")
        self.approve_btn.clicked.connect(self._approve_note)
        self.generate_btn = QPushButton("Gerar NEs")
        self.generate_btn.clicked.connect(self._generate_supplier_orders)
        self.remove_btn = QPushButton("Apagar")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_note)
        self.pdf_btn = QPushButton("Pre-visualizar NE")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(lambda: self._open_pdf(False))
        self.send_order_btn = QPushButton("Enviar encomenda")
        self.send_order_btn.setProperty("variant", "secondary")
        self.send_order_btn.clicked.connect(self._send_order_email)
        self.quote_btn = QPushButton("Pedir orçamento")
        self.quote_btn.clicked.connect(self._request_quote_email)
        self.quote_btn.setStyleSheet(
            "QPushButton {background: #f4c542; color: #0f172a; border: 1px solid #caa12b; "
            "border-radius: 10px; padding: 8px 14px; font-weight: 800;}"
            "QPushButton:hover {background: #ffd65a;}"
            "QPushButton:disabled {background: #f5e9b4; color: #7c6a2b;}"
        )
        self.suppliers_btn = QPushButton("Fornecedores")
        self.suppliers_btn.setProperty("variant", "secondary")
        self.suppliers_btn.clicked.connect(self._manage_suppliers)
        self.delivery_btn = QPushButton("Entregar encomenda")
        self.delivery_btn.setProperty("variant", "secondary")
        self.delivery_btn.clicked.connect(self._register_delivery)
        self.attach_doc_btn = QPushButton("Associar Docs")
        self.attach_doc_btn.setProperty("variant", "secondary")
        self.attach_doc_btn.clicked.connect(self._attach_document)
        self.documents_btn = QPushButton("Registo Docs")
        self.documents_btn.setProperty("variant", "secondary")
        self.documents_btn.clicked.connect(self._show_documents)
        for button in (
            self.new_btn,
            self.save_btn,
            self.approve_btn,
            self.generate_btn,
            self.remove_btn,
            self.pdf_btn,
            self.send_order_btn,
            self.suppliers_btn,
            self.delivery_btn,
            self.attach_doc_btn,
            self.documents_btn,
        ):
            actions_row.addWidget(button)
        actions_row.addStretch(1)
        actions_row.addWidget(self.quote_btn)
        form_layout.addLayout(actions_row)

        header = QHBoxLayout()
        self.number_label = QLabel("Novo pedido")
        self.number_label.setStyleSheet("font-size: 17px; font-weight: 800; color: #0f172a;")
        self.state_chip = QLabel("-")
        _apply_state_chip(self.state_chip, "-")
        header.addWidget(self.number_label, 1)
        header.addWidget(self.state_chip)
        form_layout.addLayout(header)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(12)
        supplier_card = CardFrame()
        supplier_card.set_tone("default")
        supplier_layout = QVBoxLayout(supplier_card)
        supplier_layout.setContentsMargins(12, 10, 12, 10)
        supplier_layout.setSpacing(6)
        self.supplier_combo = QComboBox()
        self._configure_supplier_selector(self.supplier_combo)
        self.supplier_combo.currentTextChanged.connect(self._sync_supplier_contact)
        self.contact_edit = QLineEdit()
        self.delivery_edit = QDateEdit()
        self.delivery_edit.setCalendarPopup(True)
        self.delivery_edit.setDisplayFormat("yyyy-MM-dd")
        self.delivery_edit.setMinimumDate(QDate(2000, 1, 1))
        self.delivery_edit.setDate(QDate.currentDate())
        self.location_edit = QComboBox()
        self.location_edit.setEditable(True)
        self.transport_edit = QComboBox()
        self.transport_edit.setEditable(True)
        self.transport_edit.addItems(["", "Nosso Cargo", "Vosso Cargo"])
        self.location_edit.addItems(["", "Vossas Instalações", "Nossas Instalações"])
        note_field_css = "font-size: 12px; padding: 4px 8px;"
        for field in (self.supplier_combo, self.contact_edit, self.delivery_edit, self.transport_edit, self.location_edit):
            field.setMinimumHeight(38)
            field.setStyleSheet(note_field_css)
        for combo in (self.supplier_combo, self.transport_edit, self.location_edit):
            try:
                combo.view().setStyleSheet("font-size: 12px;")
            except Exception:
                pass
        for widget, width, minimum in (
            (self.supplier_combo, 760, 700),
            (self.contact_edit, 290, 250),
            (self.delivery_edit, 164, 164),
            (self.transport_edit, 240, 230),
            (self.location_edit, 470, 420),
        ):
            _cap_width(widget, width)
            widget.setMinimumWidth(minimum)
        supplier_title = QLabel("Dados do documento")
        supplier_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        supplier_layout.addWidget(supplier_title)

        def _field_block(label_text: str, widget: QWidget) -> QWidget:
            host = QWidget()
            block = QVBoxLayout(host)
            block.setContentsMargins(0, 0, 0, 0)
            block.setSpacing(1)
            label = QLabel(label_text)
            label.setProperty("role", "muted")
            label.setStyleSheet("font-size: 11px; font-weight: 700; color: #415a77;")
            block.addWidget(label)
            block.addWidget(widget)
            return host

        supplier_grid = QGridLayout()
        supplier_grid.setContentsMargins(0, 0, 0, 0)
        supplier_grid.setHorizontalSpacing(8)
        supplier_grid.setVerticalSpacing(4)
        supplier_grid.addWidget(_field_block("Fornecedor base", self.supplier_combo), 0, 0)
        supplier_grid.addWidget(_field_block("Contacto", self.contact_edit), 0, 1)
        supplier_grid.addWidget(_field_block("Entrega", self.delivery_edit), 0, 2)
        supplier_grid.addWidget(_field_block("Transporte", self.transport_edit), 1, 0)
        supplier_grid.addWidget(_field_block("Local Descarga", self.location_edit), 1, 1, 1, 2)
        supplier_grid.setColumnStretch(0, 7)
        supplier_grid.setColumnStretch(1, 3)
        supplier_grid.setColumnStretch(2, 2)
        supplier_layout.addLayout(supplier_grid)
        supplier_card.setFixedHeight(170)
        supplier_card.setMinimumWidth(1320)
        supplier_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        summary_card = CardFrame()
        summary_card.set_tone("info")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(16, 12, 16, 12)
        summary_layout.setSpacing(5)
        summary_title = QLabel("Resumo")
        summary_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.summary_hint = QLabel("Pedido")
        self.summary_hint.setWordWrap(True)
        self.summary_hint.setProperty("role", "muted")
        self.summary_hint.setMaximumHeight(24)
        self.total_label = QLabel("0.00 EUR")
        self.total_label.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        self.total_label.setMinimumHeight(24)
        self.total_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_materials_label = QLabel("0,00 EUR")
        self.total_materials_label.setStyleSheet("font-size: 12px; font-weight: 800; color: #7c4a03;")
        self.total_materials_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_materials_label.setMinimumHeight(18)
        self.total_products_label = QLabel("0,00 EUR")
        self.total_products_label.setStyleSheet("font-size: 12px; font-weight: 800; color: #166534;")
        self.total_products_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_products_label.setMinimumHeight(18)
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(self.summary_hint)
        summary_grid = QGridLayout()
        summary_grid.setContentsMargins(0, 0, 0, 0)
        summary_grid.setHorizontalSpacing(10)
        summary_grid.setVerticalSpacing(4)
        mp_caption = QLabel("Matéria-Prima")
        mp_caption.setProperty("role", "muted")
        products_caption = QLabel("Produtos")
        products_caption.setProperty("role", "muted")
        summary_grid.addWidget(mp_caption, 0, 0)
        summary_grid.addWidget(self.total_materials_label, 0, 1)
        summary_grid.addWidget(products_caption, 1, 0)
        summary_grid.addWidget(self.total_products_label, 1, 1)
        summary_grid.setColumnStretch(0, 1)
        summary_grid.setColumnStretch(1, 1)
        summary_layout.addLayout(summary_grid)
        total_row = QHBoxLayout()
        total_row.setContentsMargins(0, 2, 0, 0)
        total_caption = QLabel("Total")
        total_caption.setProperty("role", "muted")
        total_row.addWidget(total_caption)
        total_row.addStretch(1)
        total_row.addWidget(self.total_label)
        summary_layout.addLayout(total_row)
        summary_card.setMinimumWidth(350)
        summary_card.setMaximumWidth(390)
        summary_card.setFixedHeight(170)
        summary_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        meta_row.addWidget(supplier_card, 10, Qt.AlignTop)
        meta_row.addWidget(summary_card, 3, Qt.AlignTop)
        form_layout.addLayout(meta_row)

        obs_card = CardFrame()
        obs_card.set_tone("default")
        obs_layout = QVBoxLayout(obs_card)
        obs_layout.setContentsMargins(10, 8, 10, 8)
        obs_layout.setSpacing(4)
        obs_title = QLabel("Observações")
        obs_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.obs_edit = QTextEdit()
        self.obs_edit.setPlaceholderText("Observações do documento")
        self.obs_edit.setMinimumHeight(38)
        self.obs_edit.setMaximumHeight(52)
        self.obs_edit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.obs_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        obs_layout.addWidget(obs_title)
        obs_layout.addWidget(self.obs_edit)
        obs_card.setMaximumHeight(82)
        form_layout.addWidget(obs_card)

        lines_card = CardFrame()
        lines_card.set_tone("default")
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(12, 10, 12, 12)
        lines_layout.setSpacing(6)
        line_header = QHBoxLayout()
        lines_title = QLabel("Linhas da nota")
        lines_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        line_header.addWidget(lines_title)
        line_header.addStretch(1)
        lines_layout.addLayout(line_header)

        line_actions = QHBoxLayout()
        self.add_material_btn = QPushButton("Adicionar Stock MP")
        self.add_material_btn.clicked.connect(self._add_material_line)
        self.add_material_btn.setToolTip("Adicionar matéria-prima de stock, incluindo chapa, perfil e tubo.")
        self.add_product_btn = QPushButton("Stock produtos")
        self.add_product_btn.setProperty("variant", "secondary")
        self.add_product_btn.clicked.connect(self._add_product_line)
        self.add_manual_btn = QPushButton("Adicionar Manual")
        self.add_manual_btn.setProperty("variant", "secondary")
        self.add_manual_btn.clicked.connect(self._add_manual_line)
        self.edit_line_btn = QPushButton("Editar Linha")
        self.edit_line_btn.setProperty("variant", "secondary")
        self.edit_line_btn.clicked.connect(self._edit_line)
        self.quick_supplier_btn = QPushButton("Fornecedor rápido")
        self.quick_supplier_btn.setProperty("variant", "secondary")
        self.quick_supplier_btn.clicked.connect(self._quick_assign_supplier)
        self.remove_line_btn = QPushButton("Remover Linha")
        self.remove_line_btn.setProperty("variant", "danger")
        self.remove_line_btn.clicked.connect(self._remove_line)
        for button in (self.add_material_btn, self.add_product_btn, self.add_manual_btn, self.edit_line_btn, self.quick_supplier_btn, self.remove_line_btn):
            line_actions.addWidget(button)
        line_actions.addStretch(1)
        lines_layout.addLayout(line_actions)

        self.lines_table = QTableWidget(0, 15)
        self.lines_table.setHorizontalHeaderLabels(
            [
                "Código",
                "Material",
                "Esp.",
                "Descrição",
                "Origem",
                "Fornecedor",
                "Qtd",
                "Unid.",
                "Peso unit.",
                "Peso tot.",
                "P.Unit.",
                "Desc.%",
                "IVA%",
                "Total",
                "Entrega",
            ]
        )
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.lines_table)
        _set_table_columns(
            self.lines_table,
            [
                (0, "fixed", 130),
                (1, "stretch", 0),
                (2, "fixed", 72),
                (3, "stretch", 0),
                (4, "fixed", 110),
                (5, "stretch", 0),
                (6, "fixed", 72),
                (7, "fixed", 62),
                (8, "fixed", 88),
                (9, "fixed", 88),
                (10, "fixed", 94),
                (11, "fixed", 72),
                (12, "fixed", 66),
                (13, "fixed", 102),
                (14, "fixed", 118),
            ],
        )
        self.lines_table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QHeaderView::section { font-size: 11px; padding: 6px 8px; font-weight: 800; }"
        )
        self.lines_table.verticalHeader().setDefaultSectionSize(28)
        self.lines_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.lines_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.lines_table.setMinimumHeight(320)
        self.lines_table.setMaximumHeight(380)
        lines_layout.addWidget(self.lines_table)
        lines_card.setMinimumHeight(390)
        lines_card.setMaximumHeight(450)
        form_layout.addWidget(lines_card, 1)
        self.note_form_card.set_tone("default")
        root.addWidget(self.note_form_card, 1)
        self.current_documents: list[dict] = []

        self._new_note()

    def refresh(self) -> None:
        previous = self.current_number
        self.supplier_rows = self.backend.ne_suppliers()
        self._set_supplier_items()
        self.rows = self.backend.ne_rows(self.filter_edit.currentText().strip(), self.state_combo.currentText())
        _fill_table(
            self.notes_table,
            [[r.get("numero", "-"), r.get("fornecedor", "-"), r.get("data_entrega", "-"), r.get("estado", "-"), _fmt_eur(r.get("total", 0)), r.get("linhas", 0)] for r in self.rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.notes_table, row_index, str(row.get("estado", "")))
        if self.notes_table.rowCount() == 0:
            self._new_note(reset_number=False)
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "")).strip() == previous:
                    row_index = index
                    break
        self.notes_table.selectRow(row_index)
        self._load_selected_note()

    def _set_supplier_items(self) -> None:
        current = self.supplier_combo.currentText().strip()
        self.supplier_combo.blockSignals(True)
        self.supplier_combo.clear()
        self.supplier_combo.addItem("")
        for row in self.supplier_rows:
            label = f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -")
            self.supplier_combo.addItem(label)
        if current:
            self.supplier_combo.setCurrentText(current)
        self.supplier_combo.blockSignals(False)

    def _configure_supplier_selector(self, combo: QComboBox) -> None:
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMaxVisibleItems(18)
        completer = combo.completer()
        if completer is not None:
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            try:
                completer.setFilterMode(Qt.MatchContains)
            except Exception:
                pass

    def _selected_line_indexes(self) -> list[int]:
        indexes = sorted({idx.row() for idx in self.lines_table.selectionModel().selectedRows()}) if self.lines_table.selectionModel() else []
        return [idx for idx in indexes if 0 <= idx < len(self.line_rows)]

    def _state_text(self) -> str:
        return self.state_chip.text().strip() or "Em edicao"

    def _selected_row(self) -> dict:
        current = self.notes_table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def open_note_numero(self, numero: str) -> None:
        target = str(numero or "").strip()
        if not target:
            return
        self.refresh()
        for row_index, row in enumerate(self.rows):
            if str(row.get("numero", "") or "").strip() != target:
                continue
            self.notes_table.selectRow(row_index)
            self._load_selected_note()
            if hasattr(self, "_show_note_detail"):
                try:
                    self._show_note_detail()
                except Exception:
                    pass
            return

    def _selected_line_index(self) -> int:
        current = self.lines_table.currentItem()
        if current is None:
            return -1
        return current.row() if current.row() < len(self.line_rows) else -1

    def _supplier_lookup(self, text: str) -> dict:
        raw = str(text or "").strip()
        if not raw:
            return {}
        supplier_id = raw.split(" - ", 1)[0].strip()
        for row in self.supplier_rows:
            if supplier_id and supplier_id == str(row.get("id", "")).strip():
                return row
            if raw.lower() == f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -").lower():
                return row
            if raw.lower() == str(row.get("nome", "")).strip().lower():
                return row
        return {}

    def _sync_supplier_contact(self) -> None:
        supplier = self._supplier_lookup(self.supplier_combo.currentText())
        if supplier and not self.contact_edit.text().strip():
            self.contact_edit.setText(str(supplier.get("contacto", "") or "").strip())

    def _line_weight_unit(self, row: dict[str, Any]) -> float:
        return max(0.0, float(row.get("peso_unid", 0) or 0))

    def _line_weight_total(self, row: dict[str, Any]) -> float:
        weight_unit = self._line_weight_unit(row)
        quantity = max(0.0, float(row.get("qtd", 0) or 0))
        if weight_unit <= 0 or quantity <= 0:
            return 0.0
        return round(weight_unit * quantity, 4)

    def _weight_text(self, value: float) -> str:
        return f"{value:.3f}" if value > 0 else ""

    def _render_lines(self) -> None:
        total = 0.0
        materials_total = 0.0
        products_total = 0.0
        def _line_code(row: dict[str, Any]) -> str:
            ref = str(row.get("ref", "") or "").strip()
            if ref:
                return ref
            if self.backend.desktop_main.origem_is_materia(row.get("origem", "")):
                return "ID pendente"
            return "Manual"
        def _line_material(row: dict[str, Any]) -> str:
            material_txt = str(row.get("material", "") or "").strip()
            if material_txt:
                return material_txt
            if self.backend.desktop_main.origem_is_materia(row.get("origem", "")):
                return "-"
            categoria_txt = str(row.get("categoria", "") or "").strip()
            tipo_txt = str(row.get("tipo", "") or "").strip()
            if categoria_txt and tipo_txt:
                return f"{categoria_txt} / {tipo_txt}"
            return categoria_txt or tipo_txt or "-"
        def _line_espessura(row: dict[str, Any]) -> str:
            if not self.backend.desktop_main.origem_is_materia(row.get("origem", "")):
                return ""
            esp_txt = str(row.get("espessura", "") or "").strip()
            return esp_txt or "-"
        _fill_table(
            self.lines_table,
            [
                [
                    _line_code(row),
                    _elide_middle(_line_material(row), 36) or "-",
                    _line_espessura(row),
                    row.get("descricao", "-"),
                    row.get("origem", "-"),
                    row.get("fornecedor_linha", "-"),
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    row.get("unid", "-"),
                    self._weight_text(self._line_weight_unit(row)),
                    self._weight_text(self._line_weight_total(row)),
                    f"{float(row.get('preco', 0) or 0):.4f}",
                    f"{float(row.get('desconto', 0) or 0):.2f}",
                    f"{float(row.get('iva', 0) or 0):.2f}",
                    f"{float(row.get('total', 0) or 0):.2f}",
                    row.get("entrega", "PENDENTE"),
                ]
                for row in self.line_rows
            ],
            align_center_from=6,
        )
        for row_index, row in enumerate(self.line_rows):
            row_total = float(row.get("total", 0) or 0)
            total += row_total
            origem = self.backend.desktop_main.norm_text(str(row.get("origem", "") or "").strip())
            if "mater" in origem:
                materials_total += row_total
            else:
                products_total += row_total
            entrega = str(row.get("entrega", "PENDENTE")).upper()
            tone = "Concluida" if "ENTREGUE" in entrega else "Em pausa" if "PARCIAL" in entrega else "Preparacao"
            _paint_table_row(self.lines_table, row_index, tone)
            code_item = self.lines_table.item(row_index, 0)
            material_item = self.lines_table.item(row_index, 1)
            esp_item = self.lines_table.item(row_index, 2)
            desc_item = self.lines_table.item(row_index, 3)
            supplier_item = self.lines_table.item(row_index, 5)
            weight_unit_item = self.lines_table.item(row_index, 8)
            weight_total_item = self.lines_table.item(row_index, 9)
            if code_item is not None:
                code_item.setToolTip(_line_code(row))
            if material_item is not None:
                material_item.setToolTip(_line_material(row))
            if esp_item is not None:
                esp_item.setToolTip(_line_espessura(row))
            if desc_item is not None:
                desc_item.setToolTip(str(row.get("descricao", "") or "").strip())
            if supplier_item is not None:
                supplier_item.setToolTip(str(row.get("fornecedor_linha", "") or "").strip())
            if weight_unit_item is not None:
                weight_unit_item.setToolTip(
                    f"{self._line_weight_unit(row):.3f} kg" if self._line_weight_unit(row) > 0 else "Sem peso associado"
                )
            if weight_total_item is not None:
                weight_total_item.setToolTip(
                    f"{self._line_weight_total(row):.3f} kg" if self._line_weight_total(row) > 0 else "Sem peso total calculado"
                )
        self.total_label.setText(_fmt_eur(total))
        self.total_materials_label.setText(_fmt_eur(materials_total))
        self.total_products_label.setText(_fmt_eur(products_total))

    def _set_header(self, numero: str, estado: str) -> None:
        self.current_number = str(numero or "").strip()
        self.number_label.setText(self.current_number or "Novo pedido")
        _apply_state_chip(self.state_chip, estado)
        _set_panel_tone(self.note_form_card, _state_tone(estado))

    def _delivery_text(self) -> str:
        selected = self.delivery_edit.date()
        if not selected.isValid() or selected.year() <= 2000:
            selected = QDate.currentDate()
            self.delivery_edit.setDate(selected)
        return selected.toString("yyyy-MM-dd").strip()

    def _set_delivery_text(self, raw: str) -> None:
        self.delivery_edit.setDate(_coerce_editor_qdate(raw, fallback_today=True))

    def _document_dialog(self, title: str, initial: dict | None = None, allow_line_apply: bool = True) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(860, 420)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        type_combo = QComboBox()
        for label, value in (
            ("Guia", "GUIA"),
            ("Fatura", "FATURA"),
            ("Guia + Fatura", "GUIA_FATURA"),
            ("Documento", "DOCUMENTO"),
        ):
            type_combo.addItem(label, value)
        initial_type = str(initial.get("tipo", "") or "").strip().upper() or "FATURA"
        for index in range(type_combo.count()):
            if str(type_combo.itemData(index) or "").strip().upper() == initial_type:
                type_combo.setCurrentIndex(index)
                break
        title_edit = QLineEdit(str(initial.get("titulo", "") or ""))
        guia_edit = QLineEdit(str(initial.get("guia", "") or ""))
        fatura_edit = QLineEdit(str(initial.get("fatura", "") or ""))
        path_edit = QLineEdit(str(initial.get("caminho", "") or ""))
        obs_edit = QLineEdit(str(initial.get("obs", "") or ""))
        entrega_edit = QDateEdit()
        entrega_edit.setCalendarPopup(True)
        entrega_edit.setDisplayFormat("yyyy-MM-dd")
        entrega_edit.setDate(_coerce_editor_qdate(str(initial.get("data_entrega", "") or ""), fallback_today=True))
        doc_edit = QDateEdit()
        doc_edit.setCalendarPopup(True)
        doc_edit.setDisplayFormat("yyyy-MM-dd")
        doc_edit.setDate(_coerce_editor_qdate(str(initial.get("data_documento", "") or ""), fallback_today=True))
        apply_lines_chk = QCheckBox("Aplicar aos registos de linhas já entregues")
        apply_lines_chk.setChecked(bool(initial.get("apply_to_lines", True)))
        apply_lines_chk.setVisible(allow_line_apply)

        def pick_path() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar documento",
                "",
                "Documentos (*.pdf *.png *.jpg *.jpeg *.bmp *.xlsx *.xls *.doc *.docx);;Todos (*.*)",
            )
            if path:
                path_edit.setText(path)

        browse_btn = QPushButton("Selecionar ficheiro")
        browse_btn.setProperty("variant", "secondary")
        browse_btn.clicked.connect(pick_path)

        form.addWidget(QLabel("Tipo"), 0, 0)
        form.addWidget(type_combo, 0, 1)
        form.addWidget(QLabel("Titulo"), 0, 2)
        form.addWidget(title_edit, 0, 3)
        form.addWidget(QLabel("Data entrega"), 1, 0)
        form.addWidget(entrega_edit, 1, 1)
        form.addWidget(QLabel("Data documento"), 1, 2)
        form.addWidget(doc_edit, 1, 3)
        form.addWidget(QLabel("Guia"), 2, 0)
        form.addWidget(guia_edit, 2, 1)
        form.addWidget(QLabel("Fatura"), 2, 2)
        form.addWidget(fatura_edit, 2, 3)
        form.addWidget(QLabel("Caminho"), 3, 0)
        form.addWidget(path_edit, 3, 1, 1, 2)
        form.addWidget(browse_btn, 3, 3)
        form.addWidget(QLabel("Obs."), 4, 0)
        form.addWidget(obs_edit, 4, 1, 1, 3)
        layout.addLayout(form)
        if allow_line_apply:
            layout.addWidget(apply_lines_chk)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = {
            "tipo": str(type_combo.currentData() or "DOCUMENTO").strip(),
            "titulo": title_edit.text().strip(),
            "guia": guia_edit.text().strip(),
            "fatura": fatura_edit.text().strip(),
            "caminho": path_edit.text().strip(),
            "data_entrega": entrega_edit.date().toString("yyyy-MM-dd"),
            "data_documento": doc_edit.date().toString("yyyy-MM-dd"),
            "obs": obs_edit.text().strip(),
            "apply_to_lines": bool(apply_lines_chk.isChecked()) if allow_line_apply else False,
        }
        if not any(str(payload.get(key, "") or "").strip() for key in ("titulo", "guia", "fatura", "caminho", "obs")):
            QMessageBox.warning(dialog, "Notas Encomenda", "Indica pelo menos titulo, guia, fatura, caminho ou observacao.")
            return None
        return payload

    def _attach_document(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            detail = self.backend.ne_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        default_type = "FATURA" if str(detail.get("guia_ultima", "") or "").strip() and not str(detail.get("fatura_ultima", "") or "").strip() else "GUIA"
        payload = self._document_dialog(
            f"Associar documento {self.current_number}",
            {
                "tipo": default_type,
                "guia": str(detail.get("guia_ultima", "") or "").strip() if default_type == "GUIA" else "",
                "fatura": "",
                "caminho": str(detail.get("fatura_caminho_ultima", "") or "").strip() if default_type == "FATURA" else "",
                "data_entrega": str(detail.get("data_ultima_entrega", "") or detail.get("data_entrega", "") or "").strip(),
                "data_documento": str(detail.get("data_doc_ultima", "") or detail.get("data_entrega", "") or "").strip(),
                "apply_to_lines": True,
            },
            allow_line_apply=True,
        )
        if payload is None:
            return
        try:
            self.backend.ne_add_document(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Associar Docs", str(exc))
            return
        self.refresh()
        QMessageBox.information(self, "Associar Docs", f"Documento registado em {self.current_number}.")

    def _show_documents(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona primeiro uma nota.")
            return
        try:
            docs = [dict(row) for row in list(self.backend.ne_documents(self.current_number) or [])]
        except Exception as exc:
            QMessageBox.critical(self, "Registo Docs", str(exc))
            return
        self.current_documents = docs
        if not docs:
            QMessageBox.information(self, "Registo Docs", "Esta nota ainda não tem documentos associados.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Registo documental {self.current_number}")
        dialog.resize(1120, 520)
        layout = QVBoxLayout(dialog)
        info = QLabel(f"{len(docs)} documento(s) associado(s) a esta nota.")
        info.setProperty("role", "muted")
        layout.addWidget(info)
        table = QTableWidget(len(docs), 8)
        table.setHorizontalHeaderLabels(["Tipo", "Titulo", "Guia", "Fatura", "Data Doc", "Data Entrega", "Registo", "Caminho"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        _configure_table(table, stretch=(1, 7), contents=(0, 2, 3, 4, 5, 6))
        _fill_table(
            table,
            [
                [
                    str(doc.get("tipo_label", "") or doc.get("tipo", "") or "-"),
                    str(doc.get("titulo", "") or "-"),
                    str(doc.get("guia", "") or "-"),
                    str(doc.get("fatura", "") or "-"),
                    str(doc.get("data_documento", "") or "-"),
                    str(doc.get("data_entrega", "") or "-"),
                    str(doc.get("data_registo", "") or "-").replace("T", " ")[:19],
                    str(doc.get("caminho", "") or "-"),
                ]
                for doc in docs
            ],
            align_center_from=2,
        )
        layout.addWidget(table, 1)

        def open_selected() -> None:
            current = table.currentItem()
            if current is None:
                QMessageBox.warning(dialog, "Registo Docs", "Seleciona um registo.")
                return
            row_index = current.row()
            if row_index < 0 or row_index >= len(docs):
                return
            raw_path = str(docs[row_index].get("caminho", "") or "").strip()
            if not raw_path:
                QMessageBox.information(dialog, "Registo Docs", "Este registo não tem caminho associado.")
                return
            try:
                self.backend.open_file_reference(raw_path)
            except Exception as exc:
                QMessageBox.warning(dialog, "Registo Docs", str(exc))

        actions = QHBoxLayout()
        open_btn = QPushButton("Abrir ficheiro")
        open_btn.setProperty("variant", "secondary")
        open_btn.clicked.connect(open_selected)
        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept)
        actions.addWidget(open_btn)
        actions.addStretch(1)
        actions.addWidget(close_btn)
        layout.addLayout(actions)
        if table.rowCount() > 0:
            table.selectRow(0)
        dialog.exec()

    def _manage_suppliers(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Fornecedores")
        dialog.resize(1040, 680)
        layout = QVBoxLayout(dialog)
        page = SuppliersPage(self.backend, dialog)
        layout.addWidget(page)
        page.refresh()
        close_btn = QDialogButtonBox(QDialogButtonBox.Close)
        close_btn.rejected.connect(dialog.reject)
        close_btn.button(QDialogButtonBox.Close).clicked.connect(dialog.reject)
        layout.addWidget(close_btn)
        dialog.exec()
        self.refresh()

    def _product_line_dialog(self, title: str, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(920)
        dialog.setStyleSheet(
            "QDialog { font-size: 12px; }"
            " QLabel { font-size: 12px; }"
            " QLineEdit, QComboBox, QDoubleSpinBox { font-size: 12px; min-height: 30px; padding: 0 8px; }"
            " QPushButton { min-height: 34px; font-size: 12px; }"
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        intro = QLabel(
            "Seleciona um produto em stock e ajusta a linha comercial. "
            "A base do preço acompanha o tipo do produto para manter o custo coerente."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        product_rows = [dict(row or {}) for row in list(self.backend.ne_product_options("") or [])]
        current_ref = str(initial.get("ref", "") or "").strip()
        current_payload = next(
            (dict(row or {}) for row in product_rows if str(row.get("codigo", "") or "").strip() == current_ref),
            None,
        )
        if current_payload is None and current_ref:
            current_payload = {
                "codigo": current_ref,
                "descricao": str(initial.get("descricao", "") or "").strip(),
                "origem": "Produto",
                "stock": float(initial.get("stock", 0) or 0),
                "unid": str(initial.get("unid", "UN") or "UN").strip() or "UN",
                "preco": float(initial.get("preco", 0) or 0),
                "preco_unid": float(initial.get("preco", 0) or 0),
                "p_compra": float(initial.get("p_compra", initial.get("preco", 0)) or 0),
                "categoria": str(initial.get("categoria", "") or "").strip(),
                "tipo": str(initial.get("tipo", "") or "").strip(),
                "dimensoes": str(initial.get("dimensoes", "") or "").strip(),
                "peso_unid": float(initial.get("peso_unid", 0) or 0),
                "metros_unidade": float(initial.get("metros_unidade", initial.get("metros", 0)) or 0),
                "price_mode": str(initial.get("price_mode", "") or "").strip(),
            }
            product_rows.insert(0, current_payload)

        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.setInsertPolicy(QComboBox.NoInsert)
        product_combo.lineEdit().setPlaceholderText("Selecionar / pesquisar produto")
        product_combo.addItem("", None)

        def _product_label(row: dict[str, Any]) -> str:
            stock_value = float(row.get("stock", 0) or 0)
            return (
                f"{str(row.get('codigo', '') or '').strip()} | "
                f"{str(row.get('descricao', '') or '').strip()} | "
                f"{stock_value:.2f} {str(row.get('unid', 'UN') or 'UN').strip() or 'UN'}"
            ).strip(" |")

        for row in product_rows:
            product_combo.addItem(_product_label(row), row)

        selector_card = CardFrame()
        selector_layout = QVBoxLayout(selector_card)
        selector_layout.setContentsMargins(12, 10, 12, 10)
        selector_layout.setSpacing(8)
        selector_title = QLabel("Seleção de produto em stock")
        selector_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        selector_hint = QLabel("Pesquisa pelo código ou descrição. O stock disponível aparece no seletor e no cartão técnico.")
        selector_hint.setWordWrap(True)
        selector_hint.setProperty("role", "muted")
        selector_layout.addWidget(selector_title)
        selector_layout.addWidget(selector_hint)
        selector_layout.addWidget(product_combo)
        layout.addWidget(selector_card)

        info_card = CardFrame()
        info_card.set_tone("info")
        info_layout = QGridLayout(info_card)
        info_layout.setContentsMargins(12, 10, 12, 10)
        info_layout.setHorizontalSpacing(12)
        info_layout.setVerticalSpacing(8)
        code_display = QLineEdit()
        category_display = QLineEdit()
        dimensions_display = QLineEdit()
        stock_display = QLineEdit()
        for widget in (code_display, category_display, dimensions_display, stock_display):
            widget.setReadOnly(True)
        price_base_label = QLabel("Preço base (EUR/un)")
        price_base_spin = QDoubleSpinBox()
        price_base_spin.setRange(0.0, 1000000.0)
        price_base_spin.setDecimals(4)
        metric_label = QLabel("Métrica por unid.")
        metric_display = QLineEdit()
        metric_display.setReadOnly(True)
        info_layout.addWidget(QLabel("Código produto"), 0, 0)
        info_layout.addWidget(QLabel("Categoria / Tipo"), 0, 1)
        info_layout.addWidget(code_display, 1, 0)
        info_layout.addWidget(category_display, 1, 1)
        info_layout.addWidget(QLabel("Dimensões"), 2, 0)
        info_layout.addWidget(QLabel("Disponível"), 2, 1)
        info_layout.addWidget(dimensions_display, 3, 0)
        info_layout.addWidget(stock_display, 3, 1)
        info_layout.addWidget(price_base_label, 4, 0)
        info_layout.addWidget(metric_label, 4, 1)
        info_layout.addWidget(price_base_spin, 5, 0)
        info_layout.addWidget(metric_display, 5, 1)
        layout.addWidget(info_card)

        desc_edit = QLineEdit(str(initial.get("descricao", "") or ""))
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
        unid_edit = QLineEdit(str(initial.get("unid", "UN") or "UN"))
        preco_spin = QDoubleSpinBox()
        preco_spin.setRange(0.0, 1000000.0)
        preco_spin.setDecimals(4)
        preco_spin.setValue(float(initial.get("preco", 0) or 0))
        desconto_spin = QDoubleSpinBox()
        desconto_spin.setRange(0.0, 100.0)
        desconto_spin.setDecimals(2)
        desconto_spin.setValue(float(initial.get("desconto", 0) or 0))
        iva_spin = QDoubleSpinBox()
        iva_spin.setRange(0.0, 100.0)
        iva_spin.setDecimals(2)
        iva_spin.setValue(float(initial.get("iva", 23) or 23))

        commercial_card = CardFrame()
        commercial_layout = QFormLayout(commercial_card)
        commercial_layout.setContentsMargins(12, 10, 12, 10)
        commercial_layout.setHorizontalSpacing(10)
        commercial_layout.setVerticalSpacing(8)
        commercial_hint = QLabel("A descrição continua editável e podes ajustar o preço da linha sem perder a referência técnica do produto.")
        commercial_hint.setWordWrap(True)
        commercial_hint.setProperty("role", "muted")
        commercial_layout.addRow(commercial_hint)
        commercial_layout.addRow("Descrição da linha", desc_edit)
        commercial_layout.addRow("Fornecedor linha", supplier_combo)
        commercial_layout.addRow("Quantidade", qtd_spin)
        commercial_layout.addRow("Unid.", unid_edit)
        commercial_layout.addRow("Preço unid. (EUR)", preco_spin)
        commercial_layout.addRow("Desc. %", desconto_spin)
        commercial_layout.addRow("IVA %", iva_spin)
        layout.addWidget(commercial_card)

        sync = {"busy": False}
        desc_state = {"manual": False, "last_auto": ""}
        price_state = {"mode": "compra", "metric": 0.0}

        def _product_mode(payload: dict[str, Any]) -> tuple[str, float]:
            category = str(payload.get("categoria", "") or "").strip()
            prod_type = str(payload.get("tipo", "") or "").strip()
            mode = str(payload.get("price_mode", "") or self.backend.desktop_main.produto_modo_preco(category, prod_type) or "compra").strip()
            if mode == "peso":
                metric_value = float(payload.get("peso_unid", 0) or 0)
                return ("peso", metric_value) if metric_value > 0 else ("compra", 0.0)
            if mode == "metros":
                metric_value = float(payload.get("metros_unidade", payload.get("metros", 0)) or 0)
                return ("metros", metric_value) if metric_value > 0 else ("compra", 0.0)
            return "compra", 0.0

        def _base_label(mode: str) -> str:
            if mode == "peso":
                return "Preço kg (EUR/kg)"
            if mode == "metros":
                return "Preço metro (EUR/m)"
            return "Preço base (EUR/un)"

        def _metric_label_text(mode: str) -> str:
            if mode == "peso":
                return "Peso por unid. (kg)"
            if mode == "metros":
                return "Metros por unid. (m)"
            return "Métrica por unid."

        def _metric_text(mode: str, metric_value: float) -> str:
            if mode == "peso":
                return f"{self.backend._fmt(metric_value)} kg"
            if mode == "metros":
                return f"{self.backend._fmt(metric_value)} m"
            return "-"

        def _compute_unit_price(mode: str, metric_value: float, base_value: float) -> float:
            metric_value = max(0.0, float(metric_value or 0))
            if mode in {"peso", "metros"}:
                return round(base_value * metric_value, 4) if metric_value > 0 else 0.0
            return round(base_value, 4)

        def _compute_base_price(mode: str, metric_value: float, unit_value: float) -> float:
            metric_value = max(0.0, float(metric_value or 0))
            if mode in {"peso", "metros"}:
                return round(unit_value / metric_value, 6) if metric_value > 0 else round(unit_value, 6)
            return round(unit_value, 6)

        def _set_metric_visible(visible: bool) -> None:
            metric_label.setVisible(visible)
            metric_display.setVisible(visible)

        def _set_auto_description(text: str, force: bool = False) -> None:
            current_text = desc_edit.text().strip()
            if not force and desc_state["manual"] and current_text and current_text != desc_state["last_auto"]:
                return
            desc_edit.blockSignals(True)
            desc_edit.setText(text)
            desc_edit.blockSignals(False)
            desc_state["last_auto"] = text
            if force:
                desc_state["manual"] = False

        def _clear_product_state() -> None:
            sync["busy"] = True
            try:
                code_display.setText(current_ref if current_ref and current_payload is None else "")
                category_display.setText("-")
                dimensions_display.setText("-")
                stock_display.setText("-")
                price_state["mode"] = "compra"
                price_state["metric"] = 0.0
                price_base_label.setText(_base_label("compra"))
                if "p_compra" in initial:
                    price_base_spin.setValue(float(initial.get("p_compra", 0) or 0))
                else:
                    price_base_spin.setValue(float(initial.get("preco", 0) or 0))
                metric_label.setText(_metric_label_text("compra"))
                metric_display.setText("-")
                _set_metric_visible(False)
            finally:
                sync["busy"] = False

        def _apply_payload(payload: dict[str, Any], *, prefer_existing_prices: bool = False) -> None:
            if not isinstance(payload, dict):
                return
            mode, metric_value = _product_mode(payload)
            unit_value = (
                float(initial.get("preco", payload.get("preco_unid", payload.get("preco", 0))) or 0)
                if prefer_existing_prices
                else float(payload.get("preco_unid", payload.get("preco", 0)) or 0)
            )
            if prefer_existing_prices and "p_compra" in initial:
                base_value = float(initial.get("p_compra", 0) or 0)
            elif prefer_existing_prices:
                base_value = _compute_base_price(mode, metric_value, unit_value)
            else:
                base_value = float(payload.get("p_compra", payload.get("preco_unid", payload.get("preco", 0))) or 0)
            category = str(payload.get("categoria", "") or "").strip()
            prod_type = str(payload.get("tipo", "") or "").strip()
            category_type = " / ".join(part for part in (category, prod_type) if part) or "-"
            sync["busy"] = True
            try:
                code_display.setText(str(payload.get("codigo", "") or "").strip())
                category_display.setText(category_type)
                dimensions_display.setText(str(payload.get("dimensoes", "") or "-").strip() or "-")
                stock_display.setText(
                    f"{float(payload.get('stock', 0) or 0):.2f} {str(payload.get('unid', 'UN') or 'UN').strip() or 'UN'}"
                )
                unid_edit.setText(str(payload.get("unid", "UN") or "UN").strip() or "UN")
                price_state["mode"] = mode
                price_state["metric"] = metric_value
                price_base_label.setText(_base_label(mode))
                price_base_spin.setValue(base_value)
                metric_label.setText(_metric_label_text(mode))
                metric_display.setText(_metric_text(mode, metric_value))
                _set_metric_visible(mode in {"peso", "metros"})
                preco_spin.setValue(unit_value)
                auto_desc = str(payload.get("descricao", "") or "").strip()
                if prefer_existing_prices and str(initial.get("descricao", "") or "").strip():
                    _set_auto_description(str(initial.get("descricao", "") or "").strip(), force=True)
                else:
                    _set_auto_description(auto_desc, force=True)
            finally:
                sync["busy"] = False

        def _on_product_change() -> None:
            payload = product_combo.currentData()
            if isinstance(payload, dict):
                selected_code = str(payload.get("codigo", "") or "").strip()
                _apply_payload(payload, prefer_existing_prices=bool(current_ref and selected_code == current_ref))
            else:
                _clear_product_state()

        def _on_base_price_changed(value: float) -> None:
            if sync["busy"]:
                return
            sync["busy"] = True
            try:
                preco_spin.setValue(_compute_unit_price(str(price_state["mode"]), float(price_state["metric"]), value))
            finally:
                sync["busy"] = False

        def _on_unit_price_changed(value: float) -> None:
            if sync["busy"]:
                return
            sync["busy"] = True
            try:
                price_base_spin.setValue(_compute_base_price(str(price_state["mode"]), float(price_state["metric"]), value))
            finally:
                sync["busy"] = False

        def _on_description_edited(_text: str) -> None:
            if sync["busy"]:
                return
            desc_state["manual"] = True

        def _accept_dialog() -> None:
            if not isinstance(product_combo.currentData(), dict):
                QMessageBox.warning(dialog, "Notas Encomenda", "Seleciona um produto em stock.")
                return
            if not desc_edit.text().strip():
                QMessageBox.warning(dialog, "Notas Encomenda", "Indica a descrição da linha.")
                return
            dialog.accept()

        product_combo.currentIndexChanged.connect(lambda _i: _on_product_change())
        price_base_spin.valueChanged.connect(_on_base_price_changed)
        preco_spin.valueChanged.connect(_on_unit_price_changed)
        desc_edit.textEdited.connect(_on_description_edited)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(_accept_dialog)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if current_ref:
            for index in range(product_combo.count()):
                payload = product_combo.itemData(index)
                if isinstance(payload, dict) and str(payload.get("codigo", "") or "").strip() == current_ref:
                    product_combo.setCurrentIndex(index)
                    break
        else:
            product_combo.setCurrentIndex(0)
            _clear_product_state()
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = product_combo.currentData()
        if not isinstance(payload, dict):
            return None
        mode = str(price_state.get("mode", "compra") or "compra").strip() or "compra"
        metric_value = float(price_state.get("metric", 0) or 0)
        return {
            "ref": str(payload.get("codigo", "") or "").strip(),
            "descricao": desc_edit.text().strip(),
            "fornecedor_linha": supplier_combo.currentText().strip(),
            "origem": "Produto",
            "qtd": qtd_spin.value(),
            "unid": unid_edit.text().strip() or "UN",
            "preco": preco_spin.value(),
            "desconto": desconto_spin.value(),
            "iva": iva_spin.value(),
            "entrega": str(initial.get("entrega", "PENDENTE") or "PENDENTE"),
            "categoria": str(payload.get("categoria", "") or "").strip(),
            "tipo": str(payload.get("tipo", "") or "").strip(),
            "dimensoes": str(payload.get("dimensoes", "") or "").strip(),
            "p_compra": float(price_base_spin.value() or 0),
            "peso_unid": metric_value if mode == "peso" else 0.0,
            "metros_unidade": metric_value if mode == "metros" else 0.0,
            "price_basis": "kg" if mode == "peso" else ("meter" if mode == "metros" else "unit"),
        }

    def _delivery_dialog(self) -> dict | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Entrega {self.current_number}")
        dialog.setMinimumSize(900, 620)
        dialog.resize(980, 700)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setSizeGripEnabled(True)
        dialog_layout = QVBoxLayout(dialog)
        dialog_layout.setContentsMargins(12, 10, 12, 10)
        dialog_layout.setSpacing(10)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        content = QWidget()
        content.setMaximumWidth(940)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(10)
        scroll.setWidget(content)
        dialog_layout.addWidget(scroll, 1)
        entrega_edit = QDateEdit()
        entrega_edit.setCalendarPopup(True)
        entrega_edit.setDisplayFormat("yyyy-MM-dd")
        entrega_edit.setDate(QDate.currentDate())
        doc_edit = QDateEdit()
        doc_edit.setCalendarPopup(True)
        doc_edit.setDisplayFormat("yyyy-MM-dd")
        doc_edit.setDate(QDate.currentDate())
        title_edit = QLineEdit()
        guia_edit = QLineEdit()
        fatura_edit = QLineEdit()
        path_edit = QLineEdit()
        obs_edit = QLineEdit()
        for field in (entrega_edit, doc_edit, title_edit, guia_edit, fatura_edit, path_edit, obs_edit):
            field.setMinimumHeight(34)
        entrega_edit.setMinimumWidth(170)
        doc_edit.setMinimumWidth(170)
        _cap_width(entrega_edit, 190)
        _cap_width(doc_edit, 190)
        path_edit.setPlaceholderText("Opcional")
        title_edit.setPlaceholderText("Opcional")
        guia_edit.setPlaceholderText("Opcional")
        fatura_edit.setPlaceholderText("Opcional")
        obs_edit.setPlaceholderText("Observações da entrega (opcional)")

        def pick_path() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar documento de entrega",
                "",
                "Documentos (*.pdf *.png *.jpg *.jpeg *.bmp *.xlsx *.xls *.doc *.docx);;Todos (*.*)",
            )
            if path:
                path_edit.setText(path)

        browse_btn = QPushButton("Selecionar ficheiro")
        browse_btn.setProperty("variant", "secondary")
        browse_btn.clicked.connect(pick_path)
        _cap_width(browse_btn, 180)

        meta_card = CardFrame()
        meta_card.set_tone("info")
        meta_layout = QGridLayout(meta_card)
        meta_layout.setContentsMargins(14, 12, 14, 12)
        meta_layout.setHorizontalSpacing(12)
        meta_layout.setVerticalSpacing(8)
        meta_title = QLabel("Documento de entrega")
        meta_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        meta_hint = QLabel(
            "Preenche apenas o essencial. O lote da entrada e a localização são opcionais nas linhas abaixo."
        )
        meta_hint.setProperty("role", "muted")
        meta_hint.setWordWrap(True)
        path_row = QWidget()
        path_row_layout = QHBoxLayout(path_row)
        path_row_layout.setContentsMargins(0, 0, 0, 0)
        path_row_layout.setSpacing(8)
        path_row_layout.addWidget(path_edit, 1)
        path_row_layout.addWidget(browse_btn, 0)
        meta_layout.addWidget(meta_title, 0, 0, 1, 4)
        meta_layout.addWidget(meta_hint, 1, 0, 1, 4)
        label_entrega = QLabel("Data entrega")
        label_doc = QLabel("Data documento")
        label_titulo = QLabel("Título")
        label_guia = QLabel("Guia")
        label_fatura = QLabel("Fatura")
        label_path = QLabel("Documento")
        label_obs = QLabel("Obs.")
        for label in (label_entrega, label_doc, label_titulo, label_guia, label_fatura, label_path, label_obs):
            label.setProperty("role", "muted")
        meta_layout.addWidget(label_entrega, 2, 0)
        meta_layout.addWidget(label_doc, 2, 2)
        meta_layout.addWidget(entrega_edit, 3, 0, 1, 2)
        meta_layout.addWidget(doc_edit, 3, 2, 1, 2)
        meta_layout.addWidget(label_titulo, 4, 0)
        meta_layout.addWidget(label_guia, 4, 2)
        meta_layout.addWidget(title_edit, 5, 0, 1, 2)
        meta_layout.addWidget(guia_edit, 5, 2, 1, 2)
        meta_layout.addWidget(label_fatura, 6, 0)
        meta_layout.addWidget(label_path, 6, 2)
        meta_layout.addWidget(fatura_edit, 7, 0, 1, 2)
        meta_layout.addWidget(path_row, 7, 2, 1, 2)
        meta_layout.addWidget(label_obs, 8, 0)
        meta_layout.addWidget(obs_edit, 9, 0, 1, 4)
        for col in range(4):
            meta_layout.setColumnStretch(col, 1)
        layout.addWidget(meta_card)

        table = QTableWidget(len(self.line_rows), 10)
        table.setHorizontalHeaderLabels(
            ["Ref", "Material", "Esp.", "Descrição", "Peso unit.", "Peso tot.", "Pendente", "Receber", "Lote", "Localização"]
        )
        table.verticalHeader().setVisible(False)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setDefaultSectionSize(34)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.setSectionResizeMode(7, QHeaderView.Fixed)
        header.setSectionResizeMode(8, QHeaderView.Fixed)
        header.setSectionResizeMode(9, QHeaderView.Fixed)
        table.setColumnWidth(0, 102)
        table.setColumnWidth(2, 62)
        table.setColumnWidth(4, 82)
        table.setColumnWidth(5, 82)
        table.setColumnWidth(6, 82)
        table.setColumnWidth(7, 92)
        table.setColumnWidth(8, 130)
        table.setColumnWidth(9, 170)
        table.setMinimumHeight(max(220, min(380, _table_visible_height(table, min(len(self.line_rows), 7), extra=12))))
        qty_inputs: list[QDoubleSpinBox] = []
        lote_inputs: list[QLineEdit] = []
        local_inputs: list[QComboBox] = []
        presets = self.backend.material_presets()
        location_options = [
            str(value or "").strip()
            for value in list(dict.fromkeys(list(presets.get("locais", []) or [])))
            if str(value or "").strip()
        ]
        for row_index, row in enumerate(self.line_rows):
            ref = str(row.get("ref", "") or "").strip()
            desc = str(row.get("descricao", "") or "").strip()
            material_txt = str(row.get("material", "") or "").strip()
            esp_txt = str(row.get("espessura", "") or "").strip()
            entrega_txt = str(row.get("entrega", "PENDENTE") or "PENDENTE").upper()
            qtd_total = float(row.get("qtd", 0) or 0)
            pending = qtd_total
            if entrega_txt.startswith("PARCIAL"):
                try:
                    partial_txt = entrega_txt.split("(", 1)[1].split("/", 1)[0]
                    pending = max(0.0, qtd_total - float(str(partial_txt).replace(",", ".")))
                except Exception:
                    pending = qtd_total
            elif "ENTREGUE" in entrega_txt:
                pending = 0.0
            table.setItem(row_index, 0, QTableWidgetItem(ref))
            material_item = QTableWidgetItem(_elide_middle(material_txt, 34))
            material_item.setToolTip(material_txt)
            table.setItem(row_index, 1, material_item)
            esp_item = QTableWidgetItem(esp_txt if self.backend.desktop_main.origem_is_materia(row.get("origem", "")) else "")
            esp_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            esp_item.setToolTip(esp_txt)
            table.setItem(row_index, 2, esp_item)
            desc_item = QTableWidgetItem(_elide_middle(desc, 52))
            desc_item.setToolTip(desc)
            table.setItem(row_index, 3, desc_item)
            weight_unit = self._line_weight_unit(row)
            weight_total = round(weight_unit * pending, 4) if weight_unit > 0 and pending > 0 else 0.0
            weight_unit_item = QTableWidgetItem(self._weight_text(weight_unit))
            weight_unit_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            weight_unit_item.setToolTip(f"{weight_unit:.3f} kg" if weight_unit > 0 else "Sem peso associado")
            table.setItem(row_index, 4, weight_unit_item)
            weight_total_item = QTableWidgetItem(self._weight_text(weight_total))
            weight_total_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            weight_total_item.setToolTip(f"{weight_total:.3f} kg" if weight_total > 0 else "Sem peso total calculado")
            table.setItem(row_index, 5, weight_total_item)
            pending_item = QTableWidgetItem(f"{pending:.2f}")
            pending_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, 6, pending_item)
            qty_spin = QDoubleSpinBox()
            qty_spin.setRange(0.0, pending)
            qty_spin.setDecimals(2)
            qty_spin.setValue(pending if pending > 0 else 0.0)
            qty_spin.setMinimumWidth(112)
            qty_spin.setMinimumHeight(30)
            qty_spin.setAlignment(Qt.AlignCenter)
            qty_inputs.append(qty_spin)
            table.setCellWidget(row_index, 7, qty_spin)
            is_material_line = self.backend.desktop_main.origem_is_materia(row.get("origem", ""))
            lote_edit = QLineEdit("")
            lote_edit.setMinimumHeight(30)
            lote_edit.setPlaceholderText("Escrever apenas se necessário")
            lote_edit.setEnabled(is_material_line)
            lote_inputs.append(lote_edit)
            table.setCellWidget(row_index, 8, lote_edit)
            local_combo = QComboBox()
            local_combo.setEditable(True)
            local_combo.setInsertPolicy(QComboBox.NoInsert)
            local_combo.addItem("")
            for option in location_options:
                local_combo.addItem(option)
            local_combo.setCurrentText("")
            local_combo.setMinimumHeight(30)
            local_combo.setEnabled(is_material_line)
            local_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            local_combo.setMinimumContentsLength(12)
            if local_combo.lineEdit() is not None:
                local_combo.lineEdit().setPlaceholderText("Selecionar / escrever")
            local_inputs.append(local_combo)
            table.setCellWidget(row_index, 9, local_combo)
        lines_card = CardFrame()
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(14, 12, 14, 12)
        lines_layout.setSpacing(8)
        lines_title = QLabel("Linhas a receber")
        lines_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        lines_hint = QLabel(
            "Nas linhas de matéria-prima, a entrada cria o registo de stock quando ele ainda não existe. "
            "O lote e a localização podem ser definidos agora ou ficar em branco."
        )
        lines_hint.setProperty("role", "muted")
        lines_hint.setWordWrap(True)
        lines_layout.addWidget(lines_title)
        lines_layout.addWidget(lines_hint)
        lines_layout.addWidget(table)
        layout.addWidget(lines_card, 1)
        layout.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn is not None:
            ok_btn.setText("Guardar")
        if cancel_btn is not None:
            cancel_btn.setText("Fechar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        dialog_layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        lines = []
        for row_index, row in enumerate(self.line_rows):
            value = qty_inputs[row_index].value()
            if value > 0:
                lote_txt = lote_inputs[row_index].text().strip()
                lines.append(
                    {
                        "index": row_index,
                        "ref": str(row.get("ref", "") or "").strip(),
                        "qtd": value,
                        "lote_fornecedor": lote_txt,
                        "localizacao": local_inputs[row_index].currentText().strip(),
                    }
                )
        return {
            "data_entrega": entrega_edit.date().toString("yyyy-MM-dd"),
            "data_documento": doc_edit.date().toString("yyyy-MM-dd"),
            "titulo": title_edit.text().strip(),
            "guia": guia_edit.text().strip(),
            "fatura": fatura_edit.text().strip(),
            "caminho": path_edit.text().strip(),
            "obs": obs_edit.text().strip(),
            "lines": lines,
        }

    def _create_new_note(self) -> None:
        try:
            draft = self.backend.ne_create_draft()
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        self.current_number = str(draft.get("numero", "") or "").strip()
        self.refresh()
        if hasattr(self, "view_stack"):
            self._show_note_detail()

    def _new_note(self, reset_number: bool = True) -> None:
        self.notes_table.blockSignals(True)
        self.notes_table.clearSelection()
        self.notes_table.blockSignals(False)
        self.current_number = self.backend.ne_next_number() if reset_number else self.current_number or self.backend.ne_next_number()
        self.current_documents = []
        self.line_rows = []
        self._set_header(self.current_number, "Em edicao")
        self.summary_hint.setText("Pedido")
        self.generate_btn.setEnabled(True)
        self.delivery_btn.setEnabled(False)
        self.attach_doc_btn.setEnabled(False)
        self.documents_btn.setEnabled(False)
        self.pdf_btn.setEnabled(True)
        self.send_order_btn.setEnabled(False)
        self.quote_btn.setEnabled(True)
        self.quote_btn.setToolTip("Gera o PDF de cotação e prepara o email no Outlook com os fornecedores em BCC.")
        self.supplier_combo.setCurrentText("")
        self.contact_edit.clear()
        self._set_delivery_text("")
        self.location_edit.setCurrentText("")
        self.transport_edit.setCurrentText("")
        self.obs_edit.clear()
        self._render_lines()

    def _load_selected_note(self) -> None:
        row = self._selected_row()
        numero = str(row.get("numero", "")).strip()
        if not numero:
            return
        try:
            detail = self.backend.ne_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        self._set_header(detail.get("numero", ""), detail.get("estado", "Em edicao"))
        self.supplier_combo.setCurrentText(str(detail.get("fornecedor", "") or "").strip())
        self.contact_edit.setText(str(detail.get("contacto", "") or "").strip())
        self._set_delivery_text(str(detail.get("data_entrega", "") or "").strip())
        self.location_edit.setCurrentText(str(detail.get("local_descarga", "") or "").strip())
        self.transport_edit.setCurrentText(str(detail.get("meio_transporte", "") or "").strip())
        self.obs_edit.setPlainText(str(detail.get("obs", "") or "").strip())
        self.current_documents = [dict(row) for row in list(detail.get("documents", []) or [])]
        self.line_rows = [dict(line) for line in list(detail.get("lines", []) or [])]
        kind = str(detail.get("kind", "") or "").strip()
        estado_txt = str(detail.get("estado", "") or "").strip().lower()
        can_send_order = kind != "rfq" and ("aprov" in estado_txt or "enviad" in estado_txt)
        generated = list(detail.get("ne_geradas", []) or [])
        doc_count = int(detail.get("document_count", len(self.current_documents)) or 0)
        if kind == "rfq":
            text = "Cotação"
            if generated:
                text = f"{text} | NEs {len(generated)}"
            if doc_count:
                text = f"{text} | Docs {doc_count}"
            self.summary_hint.setText(text)
            self.generate_btn.setEnabled(True)
            self.delivery_btn.setEnabled(False)
            self.attach_doc_btn.setEnabled(False)
            self.documents_btn.setEnabled(doc_count > 0)
            self.pdf_btn.setEnabled(False)
            self.send_order_btn.setEnabled(False)
            self.quote_btn.setEnabled(True)
            self.quote_btn.setToolTip("Gera o PDF de cotação e prepara o email no Outlook com os fornecedores em BCC.")
        elif kind == "supplier_order":
            summary = "NE gerada"
            if doc_count:
                summary = f"{summary} | Docs {doc_count}"
            self.summary_hint.setText(summary)
            self.generate_btn.setEnabled(False)
            self.delivery_btn.setEnabled(True)
            self.attach_doc_btn.setEnabled(True)
            self.documents_btn.setEnabled(True)
            self.pdf_btn.setEnabled(True)
            self.send_order_btn.setEnabled(can_send_order)
            self.quote_btn.setEnabled(False)
            self.quote_btn.setToolTip("As NEs já geradas por fornecedor usam o botão 'Pre-visualizar NE'.")
        else:
            summary = "Compra direta / cotação simples"
            if doc_count:
                summary = f"{summary} | Docs {doc_count}"
            self.summary_hint.setText(summary)
            self.generate_btn.setEnabled(False)
            self.delivery_btn.setEnabled(True)
            self.attach_doc_btn.setEnabled(True)
            self.documents_btn.setEnabled(True)
            self.pdf_btn.setEnabled(True)
            self.send_order_btn.setEnabled(can_send_order)
            self.quote_btn.setEnabled(True)
            self.quote_btn.setToolTip("Também podes preparar um email de pedido de cotação para uma nota com fornecedor único.")
        self._render_lines()

    def _line_dialog(self, title: str, initial: dict | None = None, material_mode: bool = False) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(980 if material_mode else 760)
        dialog.resize(1080 if material_mode else 820, 760 if material_mode else 620)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setSizeGripEnabled(True)
        dialog.setStyleSheet(
            "QDialog { font-size: 12px; }"
            " QLabel { font-size: 12px; }"
            " QLineEdit, QComboBox, QDoubleSpinBox { font-size: 12px; min-height: 30px; padding: 0 8px; }"
            " QPushButton { min-height: 34px; font-size: 12px; }"
        )
        dialog_layout = QVBoxLayout(dialog)
        dialog_layout.setContentsMargins(12, 10, 12, 10)
        dialog_layout.setSpacing(10)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        content = QWidget()
        content.setMaximumWidth(1040 if material_mode else 780)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)
        scroll.setWidget(content)
        dialog_layout.addWidget(scroll, 1)
        ref_edit = QLineEdit(str(initial.get("ref", "") or ""))
        desc_edit = QLineEdit(str(initial.get("descricao", "") or ""))
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
        unid_edit = QLineEdit(str(initial.get("unid", "UN") or "UN"))
        preco_spin = QDoubleSpinBox()
        preco_spin.setRange(0.0, 1000000.0)
        preco_spin.setDecimals(4)
        preco_spin.setValue(float(initial.get("preco", 0) or 0))
        desconto_spin = QDoubleSpinBox()
        desconto_spin.setRange(0.0, 100.0)
        desconto_spin.setDecimals(2)
        desconto_spin.setValue(float(initial.get("desconto", 0) or 0))
        iva_spin = QDoubleSpinBox()
        iva_spin.setRange(0.0, 100.0)
        iva_spin.setDecimals(2)
        iva_spin.setValue(float(initial.get("iva", 23) or 23))
        accept_handler = dialog.accept
        manual_price_controls = None
        if material_mode:
            intro = QLabel(
                "Escolhe stock existente pelos filtros ou cria uma linha manual. "
                "O código é sempre o ID do material; sem stock associado, o ID fica pendente até à entrada."
            )
            intro.setWordWrap(True)
            intro.setProperty("role", "muted")
            layout.addWidget(intro)
        if material_mode:
            material_rows = [dict(row or {}) for row in list(self.backend.ne_material_options("") or [])]
            current_ref = str(initial.get("ref", "") or "").strip()
            initial_has_manual_meta = any(
                str(initial.get(key, "") or "").strip()
                for key in ("material", "espessura", "formato", "lote_fornecedor", "localizacao", "material_familia", "secao_tipo")
            ) or any(float(initial.get(key, 0) or 0) > 0 for key in ("comprimento", "largura", "altura", "diametro", "metros", "kg_m", "peso_unid", "p_compra"))
            current_payload = next(
                (dict(row or {}) for row in material_rows if str(row.get("id", "") or "").strip() == current_ref),
                None,
            )
            if current_payload is None and (current_ref or initial_has_manual_meta):
                current_payload = {
                    "id": current_ref,
                    "descricao": str(initial.get("descricao", "") or "").strip(),
                    "material": str(initial.get("material", "") or "").strip(),
                    "material_familia": str(initial.get("material_familia", "") or "").strip(),
                    "espessura": str(initial.get("espessura", "") or "").strip(),
                    "formato": str(initial.get("formato", "") or "Chapa").strip() or "Chapa",
                    "preco": float(initial.get("preco", 0) or 0),
                    "preco_base": float(initial.get("p_compra", 0) or 0),
                    "preco_base_label": "EUR/m"
                    if str(initial.get("formato", "") or "").strip() == "Tubo"
                    else "EUR/kg",
                    "unid": str(initial.get("unid", "UN") or "UN"),
                    "lote": str(initial.get("lote_fornecedor", "") or "").strip(),
                    "localizacao": str(initial.get("localizacao", "") or "").strip(),
                    "comprimento": float(initial.get("comprimento", 0) or 0),
                    "largura": float(initial.get("largura", 0) or 0),
                    "altura": float(initial.get("altura", 0) or 0),
                    "diametro": float(initial.get("diametro", 0) or 0),
                    "metros": float(initial.get("metros", 0) or 0),
                    "kg_m": float(initial.get("kg_m", 0) or 0),
                    "peso_unid": float(initial.get("peso_unid", 0) or 0),
                    "secao_tipo": str(initial.get("secao_tipo", "") or "").strip(),
                }
                preview = dict(self.backend.material_price_preview(current_payload) or {})
                current_payload["dimensao"] = str(preview.get("dimension_label", "") or "").strip() or "-"
                current_payload["secao_tipo"] = str(preview.get("secao_tipo", current_payload.get("secao_tipo", "")) or "").strip()
                current_payload["kg_m"] = float(preview.get("kg_m", current_payload.get("kg_m", 0)) or 0)
                current_payload["peso_unid"] = float(preview.get("peso_unid", current_payload.get("peso_unid", 0)) or 0)
                material_rows.insert(0, current_payload)

            def _clean(value: Any, fallback: str = "") -> str:
                text = str(value or "").strip()
                return text or fallback

            def _float_text(value: Any, digits: int = 3) -> str:
                try:
                    number = float(value or 0)
                except Exception:
                    number = 0.0
                if digits <= 0:
                    text = f"{int(round(number))}"
                else:
                    text = f"{number:.{digits}f}".rstrip("0").rstrip(".")
                return text.replace(".", ",")

            def _material_family_options() -> list[dict[str, Any]]:
                try:
                    return [dict(row or {}) for row in list(self.backend.material_family_options() or [])]
                except Exception:
                    return [
                        {"key": "", "label": "Auto", "density": 0.0},
                        {"key": "steel", "label": "Aço / Ferro", "density": 7.85},
                        {"key": "stainless", "label": "Inox", "density": 7.93},
                        {"key": "aluminum", "label": "Alumínio", "density": 2.70},
                        {"key": "brass", "label": "Latão", "density": 8.50},
                        {"key": "copper", "label": "Cobre", "density": 8.96},
                    ]

            def _material_family_profile(material_txt: Any = "", family_txt: Any = "") -> dict[str, Any]:
                try:
                    return dict(self.backend.material_family_profile(material_txt, family_txt) or {})
                except Exception:
                    return {"key": "", "label": "Auto", "density": 7.85, "explicit": False}

            def _dimension_text(row: dict[str, Any]) -> str:
                try:
                    preview = dict(self.backend.material_geometry_preview(row) or {})
                except Exception:
                    preview = {}
                return _clean(preview.get("dimension_label", row.get("dimensao", "")), "-")

            def _norm_material_text(value: Any) -> str:
                helper = getattr(self.backend, "_norm_material_token", None)
                if callable(helper):
                    try:
                        return str(helper(value) or "")
                    except Exception:
                        pass
                return re.sub(r"\s+", "", str(value or "").strip().lower())

            def _norm_esp_text(value: Any) -> str:
                helper = getattr(self.backend, "_norm_esp_token", None)
                if callable(helper):
                    try:
                        return str(helper(value) or "")
                    except Exception:
                        pass
                txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
                txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
                return txt

            def _norm_dimension_text(formato_txt: str, dim_txt: Any) -> str:
                raw = str(dim_txt or "").strip().upper()
                if not raw:
                    return ""
                raw = raw.replace("×", "X").replace("MM", "").replace(",", ".")
                raw = re.sub(r"\s+", "", raw)
                raw = raw.replace("Ø", "D")
                tokens = re.findall(r"[A-Z]+|[-+]?\d+(?:\.\d+)?", raw)
                normalized: list[str] = []
                for token in tokens:
                    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", token):
                        try:
                            normalized.append(f"{float(token):.4f}".rstrip("0").rstrip("."))
                        except Exception:
                            normalized.append(token)
                    else:
                        normalized.append(token)
                return "|".join(normalized)

            def _matches(row: dict[str, Any], formato_txt: str, quality_txt: str, esp_txt: str, dim_txt: str) -> bool:
                row_formato = _clean(row.get("formato", ""), "Chapa")
                if formato_txt and row_formato.lower() != formato_txt.lower():
                    return False
                if quality_txt and _norm_material_text(row.get("material", "")) != _norm_material_text(quality_txt):
                    return False
                row_esp = _clean(row.get("espessura", ""), "-")
                if esp_txt and esp_txt != "-" and _norm_esp_text(row_esp) != _norm_esp_text(esp_txt):
                    return False
                if dim_txt and _norm_dimension_text(row_formato, _dimension_text(row)) != _norm_dimension_text(row_formato, dim_txt):
                    return False
                return True

            def _unique(rows: list[dict[str, Any]], key_fn) -> list[str]:
                values: list[str] = []
                for row in rows:
                    value = _clean(key_fn(row))
                    if not value:
                        continue
                    if value not in values:
                        values.append(value)
                return values

            selector_card = CardFrame()
            selector_layout = QGridLayout(selector_card)
            selector_layout.setContentsMargins(12, 10, 12, 10)
            selector_layout.setHorizontalSpacing(10)
            selector_layout.setVerticalSpacing(8)
            selector_title = QLabel("Seleção de matéria-prima")
            selector_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
            selector_hint = QLabel(
                "O stock disponível aparece logo no seletor abaixo. Usa os filtros de tipo, qualidade, espessura e dimensão apenas para afinar a lista. "
                "O stock é identificado pelo ID do material; o lote aparece apenas como referência adicional. "
                "Quando a qualidade é escrita à mão, escolhe a família para acertar a densidade e o peso."
            )
            selector_hint.setProperty("role", "muted")
            selector_hint.setWordWrap(True)
            common_sheet_dimensions = [
                "3000 x 1500 mm",
                "4000 x 2000 mm",
                "2000 x 1000 mm",
                "2500 x 1250 mm",
            ]
            formato_combo = QComboBox()
            qualidade_combo = QComboBox()
            familia_combo = QComboBox()
            espessura_combo = QComboBox()
            dimensao_combo = QComboBox()
            stock_combo = QComboBox()
            for combo, placeholder in (
                (formato_combo, "Selecionar / escrever tipo"),
                (qualidade_combo, "Selecionar / escrever qualidade"),
                (espessura_combo, "Selecionar / escrever espessura"),
                (dimensao_combo, "Selecionar / escrever dimensões"),
            ):
                combo.setEditable(True)
                combo.setInsertPolicy(QComboBox.NoInsert)
                combo.lineEdit().setPlaceholderText(placeholder)
            stock_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            stock_combo.setMinimumContentsLength(24)
            selector_layout.addWidget(selector_title, 0, 0, 1, 4)
            selector_layout.addWidget(selector_hint, 1, 0, 1, 4)
            selector_layout.addWidget(QLabel("Tipo material"), 2, 0)
            selector_layout.addWidget(QLabel("Qualidade"), 2, 1)
            selector_layout.addWidget(QLabel("Espessura"), 2, 2)
            selector_layout.addWidget(QLabel("Dimensões"), 2, 3)
            selector_layout.addWidget(formato_combo, 3, 0)
            selector_layout.addWidget(qualidade_combo, 3, 1)
            selector_layout.addWidget(espessura_combo, 3, 2)
            selector_layout.addWidget(dimensao_combo, 3, 3)
            selector_layout.addWidget(QLabel("Família material"), 4, 0)
            selector_layout.addWidget(QLabel("Densidade (g/cm3)"), 4, 1)
            selector_layout.addWidget(QLabel("Stock disponível (ID material)"), 4, 2, 1, 2)
            densidade_display = QLineEdit()
            densidade_display.setReadOnly(True)
            for option in _material_family_options():
                familia_combo.addItem(str(option.get("label", "") or ""), str(option.get("key", "") or ""))
            selector_layout.addWidget(familia_combo, 5, 0)
            selector_layout.addWidget(densidade_display, 5, 1)
            selector_layout.addWidget(stock_combo, 5, 2, 1, 2)
            for col in range(4):
                selector_layout.setColumnStretch(col, 1)
            layout.addWidget(selector_card)

            info_card = CardFrame()
            info_card.set_tone("info")
            info_layout = QGridLayout(info_card)
            info_layout.setContentsMargins(12, 10, 12, 10)
            info_layout.setHorizontalSpacing(12)
            info_layout.setVerticalSpacing(8)
            code_display = QLineEdit()
            code_display.setReadOnly(True)
            local_display = QLineEdit()
            lote_display = QLineEdit()
            desc_edit.setReadOnly(False)
            price_base_label = QLabel("Preço base")
            price_base_spin = QDoubleSpinBox()
            price_base_spin.setRange(0.0, 1000000.0)
            price_base_spin.setDecimals(4)
            metric_label = QLabel("Peso por unid. (kg)")
            metric_spin = QDoubleSpinBox()
            metric_spin.setRange(0.0, 1000000.0)
            metric_spin.setDecimals(4)
            metric_spin.setReadOnly(True)
            code_display.setPlaceholderText("Gerado na entrada")
            local_display.setPlaceholderText("Opcional")
            lote_display.setPlaceholderText("Opcional")
            info_layout.addWidget(QLabel("ID material"), 0, 0)
            info_layout.addWidget(QLabel("Localização"), 0, 1)
            info_layout.addWidget(QLabel("Lote fornecedor"), 0, 2)
            info_layout.addWidget(price_base_label, 0, 3)
            info_layout.addWidget(code_display, 1, 0)
            info_layout.addWidget(local_display, 1, 1)
            info_layout.addWidget(lote_display, 1, 2)
            info_layout.addWidget(price_base_spin, 1, 3)
            for col in range(4):
                info_layout.setColumnStretch(col, 1)
            layout.addWidget(info_card)

            tech_card = CardFrame()
            tech_layout = QGridLayout(tech_card)
            tech_layout.setContentsMargins(12, 10, 12, 10)
            tech_layout.setHorizontalSpacing(12)
            tech_layout.setVerticalSpacing(8)
            tech_title = QLabel("Dados técnicos")
            tech_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
            tech_hint = QLabel("Os mesmos campos técnicos da matéria-prima são usados aqui para garantir peso e stock coerentes.")
            tech_hint.setProperty("role", "muted")
            tech_hint.setWordWrap(True)
            secao_tipo_combo = QComboBox()
            secao_tipo_combo.setEditable(True)
            secao_tipo_combo.setInsertPolicy(QComboBox.NoInsert)
            secao_tipo_combo.lineEdit().setPlaceholderText("Selecionar tipo / série")
            comprimento_edit = QLineEdit()
            largura_edit = QLineEdit()
            altura_edit = QLineEdit()
            diametro_edit = QLineEdit()
            metros_edit = QLineEdit()
            kg_m_edit = QLineEdit()
            kg_m_edit.setPlaceholderText("Auto por tabela / fórmula")
            tech_layout.addWidget(tech_title, 0, 0, 1, 4)
            tech_layout.addWidget(tech_hint, 1, 0, 1, 4)
            secao_label = QLabel("Tipo secção")
            comp_label = QLabel("Comprimento (mm)")
            larg_label = QLabel("Largura (mm)")
            altura_label = QLabel("Altura / tamanho (mm)")
            diametro_label = QLabel("Diâmetro ext. (mm)")
            metros_label = QLabel("Comprimento barra (m)")
            kgm_label = QLabel("Peso por metro (kg/m)")
            for label in (secao_label, comp_label, larg_label, altura_label, diametro_label, metros_label, kgm_label):
                label.setProperty("role", "muted")
            tech_layout.addWidget(secao_label, 2, 0)
            tech_layout.addWidget(secao_tipo_combo, 2, 1)
            tech_layout.addWidget(comp_label, 2, 2)
            tech_layout.addWidget(comprimento_edit, 2, 3)
            tech_layout.addWidget(larg_label, 3, 0)
            tech_layout.addWidget(largura_edit, 3, 1)
            tech_layout.addWidget(altura_label, 3, 2)
            tech_layout.addWidget(altura_edit, 3, 3)
            tech_layout.addWidget(diametro_label, 4, 0)
            tech_layout.addWidget(diametro_edit, 4, 1)
            tech_layout.addWidget(metros_label, 4, 2)
            tech_layout.addWidget(metros_edit, 4, 3)
            tech_layout.addWidget(kgm_label, 5, 0)
            tech_layout.addWidget(kg_m_edit, 5, 1)
            tech_layout.addWidget(metric_label, 5, 2)
            tech_layout.addWidget(metric_spin, 5, 3)
            for col in range(4):
                tech_layout.setColumnStretch(col, 1)
            layout.addWidget(tech_card)
            commercial_card = CardFrame()
            commercial_layout = QGridLayout(commercial_card)
            commercial_layout.setContentsMargins(12, 10, 12, 10)
            commercial_layout.setHorizontalSpacing(10)
            commercial_layout.setVerticalSpacing(8)
            commercial_hint = QLabel("A linha fica simples: descrição, fornecedor, quantidade e preço. A ponte com o stock e os custos mantém-se por baixo.")
            commercial_hint.setProperty("role", "muted")
            commercial_hint.setWordWrap(True)
            commercial_layout.addWidget(commercial_hint, 0, 0, 1, 4)
            commercial_layout.addWidget(QLabel("Descrição da linha"), 1, 0)
            commercial_layout.addWidget(desc_edit, 1, 1, 1, 3)
            commercial_layout.addWidget(QLabel("Fornecedor linha"), 2, 0)
            commercial_layout.addWidget(supplier_combo, 2, 1)
            commercial_layout.addWidget(QLabel("Quantidade"), 2, 2)
            commercial_layout.addWidget(qtd_spin, 2, 3)
            commercial_layout.addWidget(QLabel("Unid."), 3, 0)
            commercial_layout.addWidget(unid_edit, 3, 1)
            commercial_layout.addWidget(QLabel("Preço unid. (EUR)"), 3, 2)
            commercial_layout.addWidget(preco_spin, 3, 3)
            commercial_layout.addWidget(QLabel("Desc. %"), 4, 0)
            commercial_layout.addWidget(desconto_spin, 4, 1)
            commercial_layout.addWidget(QLabel("IVA %"), 4, 2)
            commercial_layout.addWidget(iva_spin, 4, 3)
            for col in range(4):
                commercial_layout.setColumnStretch(col, 1)
            layout.addWidget(commercial_card)

            sync = {"busy": False}
            desc_state = {"manual": False, "last_auto": desc_edit.text().strip()}
            current_preview = {"data": {}}

            def _combo_text(combo: QComboBox) -> str:
                return combo.currentText().strip()

            def _current_family_key() -> str:
                return str(familia_combo.currentData() or "").strip()

            def _set_family_value(value: Any, *, material_txt: Any = "") -> None:
                if str(value or "").strip():
                    target_key = str(_material_family_profile(material_txt, value).get("key", "") or "").strip()
                else:
                    target_key = ""
                familia_combo.blockSignals(True)
                target_index = 0
                for idx in range(familia_combo.count()):
                    if str(familia_combo.itemData(idx) or "").strip() == target_key:
                        target_index = idx
                        break
                familia_combo.setCurrentIndex(target_index)
                familia_combo.blockSignals(False)

            def _set_section_options(formato_txt: str, current_value: Any = "") -> None:
                secao_tipo_combo.blockSignals(True)
                secao_tipo_combo.clear()
                secao_tipo_combo.addItem("", "")
                target_text = str(current_value or "").strip()
                target_index = 0
                for option in list(self.backend.material_section_options(formato_txt) or []):
                    label_txt = str(option.get("label", "") or "").strip()
                    key_txt = str(option.get("key", "") or "").strip()
                    if not label_txt:
                        continue
                    secao_tipo_combo.addItem(label_txt, key_txt)
                    if target_text and target_text in {label_txt, key_txt}:
                        target_index = secao_tipo_combo.count() - 1
                secao_tipo_combo.setCurrentIndex(target_index)
                if target_text and target_index == 0:
                    secao_tipo_combo.setEditText(target_text)
                secao_tipo_combo.blockSignals(False)

            def _section_value() -> str:
                return str(secao_tipo_combo.currentData() or secao_tipo_combo.currentText() or "").strip()

            def _price_base_label(formato_txt: str) -> str:
                return "Preço metro (EUR/m)" if formato_txt == "Tubo" else "Preço kg (EUR/kg)"

            def _metric_label_text(formato_txt: str) -> str:
                return "Metros por unid. (m)" if formato_txt == "Tubo" else "Peso por unid. (kg)"

            def _set_technical_visibility(formato_txt: str, secao_tipo: str, *, selected_stock: bool, profile_catalog: bool) -> None:
                tube_round = formato_txt == "Tubo" and secao_tipo == "redondo"
                secao_visible = formato_txt in {"Tubo", "Perfil", "Cantoneira", "Barra"}
                comp_visible = formato_txt in {"Chapa", "Cantoneira", "Barra"} or (formato_txt == "Tubo" and not tube_round)
                larg_visible = formato_txt in {"Chapa", "Cantoneira", "Barra"} or (formato_txt == "Tubo" and not tube_round)
                altura_visible = formato_txt == "Perfil"
                diametro_visible = formato_txt == "Tubo" and tube_round
                metros_visible = formato_txt in {"Tubo", "Perfil", "Cantoneira", "Barra"}
                kgm_visible = formato_txt in {"Tubo", "Perfil", "Cantoneira", "Barra"}
                if formato_txt == "Tubo":
                    secao_label.setText("Tipo tubo")
                elif formato_txt == "Perfil":
                    secao_label.setText("Tipo perfil / série")
                elif formato_txt == "Cantoneira":
                    secao_label.setText("Tipo cantoneira")
                elif formato_txt == "Barra":
                    secao_label.setText("Tipo barra")
                else:
                    secao_label.setText("Tipo secção")
                if formato_txt == "Chapa":
                    comp_label.setText("Comprimento (mm)")
                elif formato_txt == "Cantoneira":
                    comp_label.setText("Aba A (mm)")
                else:
                    comp_label.setText("Lado A (mm)")
                if formato_txt == "Chapa":
                    larg_label.setText("Largura (mm)")
                elif formato_txt == "Cantoneira":
                    larg_label.setText("Aba B (mm)")
                else:
                    larg_label.setText("Lado B (mm)")
                altura_label.setText("Altura / tamanho (mm)")
                diametro_label.setText("Diâmetro ext. (mm)")
                metros_label.setText("Comprimento barra (m)")
                kgm_label.setText("Peso por metro (kg/m)")
                for widget in (secao_label, secao_tipo_combo):
                    widget.setVisible(secao_visible)
                for widget in (comp_label, comprimento_edit):
                    widget.setVisible(comp_visible)
                for widget in (larg_label, largura_edit):
                    widget.setVisible(larg_visible)
                for widget in (altura_label, altura_edit):
                    widget.setVisible(altura_visible)
                for widget in (diametro_label, diametro_edit):
                    widget.setVisible(diametro_visible)
                for widget in (metros_label, metros_edit):
                    widget.setVisible(metros_visible)
                for widget in (kgm_label, kg_m_edit):
                    widget.setVisible(kgm_visible)
                metric_label.setVisible(True)
                metric_spin.setVisible(True)
                secao_tipo_combo.setEnabled(not selected_stock and secao_visible)
                comprimento_edit.setReadOnly(selected_stock or not comp_visible)
                largura_edit.setReadOnly(selected_stock or not larg_visible)
                altura_edit.setReadOnly(selected_stock or not altura_visible)
                diametro_edit.setReadOnly(selected_stock or not diametro_visible)
                metros_edit.setReadOnly(selected_stock or not metros_visible)
                kg_m_edit.setReadOnly(formato_txt == "Tubo" or profile_catalog or selected_stock or not kgm_visible)

            def _current_geometry_payload() -> dict[str, Any]:
                return {
                    "formato": _combo_text(formato_combo) or _clean((current_payload or {}).get("formato", ""), "Chapa"),
                    "material": _combo_text(qualidade_combo),
                    "material_familia": _current_family_key(),
                    "espessura": _combo_text(espessura_combo),
                    "secao_tipo": _section_value(),
                    "comprimento": comprimento_edit.text().strip(),
                    "largura": largura_edit.text().strip(),
                    "altura": altura_edit.text().strip(),
                    "diametro": diametro_edit.text().strip(),
                    "metros": metros_edit.text().strip(),
                    "kg_m": kg_m_edit.text().strip(),
                    "peso_unid": float(metric_spin.value() or 0),
                }

            def _apply_preview_controls(preview: dict[str, Any], *, selected_stock: bool, hydrate_dimensions: bool) -> None:
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                secao_tipo = str(preview.get("secao_tipo", "") or "").strip()
                _set_section_options(formato_txt, secao_tipo)
                if secao_tipo:
                    secao_tipo_combo.setCurrentText(secao_tipo)
                profile_catalog = formato_txt == "Perfil" and bool(preview.get("usa_catalogo"))
                _set_technical_visibility(formato_txt, secao_tipo, selected_stock=selected_stock, profile_catalog=profile_catalog)
                if hydrate_dimensions:
                    comprimento_edit.setText(_float_text(preview.get("comprimento", 0), 3))
                    largura_edit.setText(_float_text(preview.get("largura", 0), 3))
                    altura_edit.setText(_float_text(preview.get("altura", 0), 3))
                    diametro_edit.setText(_float_text(preview.get("diametro", 0), 3))
                    metros_edit.setText(_float_text(preview.get("metros", 0), 4))
                kg_m_edit.setText(_float_text(preview.get("kg_m", 0), 4))
                metric_label.setText(_metric_label_text(formato_txt))
                metric_spin.setValue(float(preview.get("peso_unid", 0) or 0))

            def _parse_dimension_values(formato_txt: str, dim_txt: str) -> tuple[float, float]:
                if formato_txt not in {"Chapa", "Perfil"}:
                    return 0.0, 0.0
                raw_values = re.findall(r"[-+]?\d+(?:[.,]\d+)?", str(dim_txt or ""))
                if len(raw_values) < 2:
                    return 0.0, 0.0
                try:
                    comp_val = float(raw_values[0].replace(",", "."))
                    larg_val = float(raw_values[1].replace(",", "."))
                except Exception:
                    return 0.0, 0.0
                return round(comp_val, 3), round(larg_val, 3)

            def _current_rows() -> list[dict[str, Any]]:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                dim_txt = _combo_text(dimensao_combo)
                if not formato_txt:
                    return list(material_rows)
                return [row for row in material_rows if _matches(row, formato_txt, quality_txt, esp_txt, dim_txt)]

            def _set_combo_items(combo: QComboBox, values: list[str], current: str = "", allow_blank: bool = False) -> None:
                current_text = current if current is not None else _combo_text(combo)
                combo.blockSignals(True)
                combo.clear()
                if allow_blank:
                    combo.addItem("")
                for value in values:
                    combo.addItem(value)
                available = [combo.itemText(i) for i in range(combo.count())]
                if current_text and current_text in available:
                    combo.setCurrentText(current_text)
                elif current_text and combo.isEditable():
                    combo.setEditText(current_text)
                elif combo.count() > 0 and not allow_blank:
                    combo.setCurrentIndex(0)
                elif combo.count() > 0 and allow_blank:
                    combo.setCurrentIndex(0)
                combo.blockSignals(False)

            def _stock_label(row: dict[str, Any]) -> str:
                material_txt = str(row.get("material", "") or "").strip()
                esp_txt = str(row.get("espessura", "") or "").strip()
                dim_txt = _dimension_text(row)
                stock_id = str(row.get("id", "") or "").strip()
                parts: list[str] = []
                if material_txt:
                    parts.append(material_txt)
                if esp_txt:
                    parts.append(f"{esp_txt} mm")
                if dim_txt and dim_txt != "-":
                    parts.append(dim_txt)
                if stock_id:
                    parts.append(stock_id)
                lote_txt = str(row.get("lote", "") or "").strip()
                local_txt = str(row.get("localizacao", "") or "").strip()
                if lote_txt:
                    parts.append(f"Lote {lote_txt}")
                if local_txt:
                    parts.append(local_txt)
                return " | ".join(part for part in parts if part) or str(row.get("descricao", "") or "").strip()

            def _compute_unit_price(formato_txt: str, metric_value: float, base_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                return round(base_value * metric_value, 4) if metric_value > 0 else 0.0

            def _compute_base_price(formato_txt: str, metric_value: float, unit_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                return round(unit_value / metric_value, 6) if metric_value > 0 else 0.0

            def _auto_description() -> str:
                preview = dict(current_preview.get("data") or {})
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                dim_txt = str(preview.get("dimension_label", "") or "").strip() or _combo_text(dimensao_combo)
                parts = [part for part in (quality_txt,) if part]
                if esp_txt:
                    parts.append(f"{esp_txt} mm")
                if dim_txt and dim_txt != "-":
                    parts.append(dim_txt)
                if str(preview.get("formato", "") or "").strip() in {"Tubo", "Perfil"} and float(preview.get("metros", 0) or 0) > 0:
                    parts.append(f"{_float_text(preview.get('metros', 0), 3)} m")
                return " | ".join(parts).strip()

            def _set_auto_description(text: str, force: bool = False) -> None:
                current_text = desc_edit.text().strip()
                if not force and desc_state["manual"] and current_text and current_text != desc_state["last_auto"]:
                    return
                desc_edit.blockSignals(True)
                desc_edit.setText(text)
                desc_edit.blockSignals(False)
                desc_state["last_auto"] = text
                if force:
                    desc_state["manual"] = False

            def _selected_metric_value(payload: dict[str, Any] | None = None) -> float:
                if isinstance(payload, dict):
                    if str(payload.get("formato", "") or "").strip() == "Tubo":
                        return float(payload.get("metros", 0) or 0)
                    return float(payload.get("peso_unid", 0) or 0)
                return float(metric_spin.value() or 0)

            def _refresh_manual_state() -> None:
                if sync["busy"]:
                    return
                formato_txt = _combo_text(formato_combo) or _clean((current_payload or {}).get("formato", ""), "Chapa")
                sync["busy"] = True
                try:
                    price_base_label.setText(_price_base_label(formato_txt))
                    selected_stock = stock_combo.currentData()
                    has_selected_stock = isinstance(selected_stock, dict)
                    familia_combo.setEnabled(not has_selected_stock)
                    code_display.setText(
                        str((selected_stock or {}).get("id", "") or "").strip() if has_selected_stock else (current_ref if current_ref else "")
                    )
                    lote_display.setReadOnly(has_selected_stock)
                    local_display.setReadOnly(has_selected_stock)
                    quality_txt = _combo_text(qualidade_combo)
                    selected_family = str((selected_stock or {}).get("material_familia", (selected_stock or {}).get("material_familia_resolved", "")) or "").strip()
                    if has_selected_stock:
                        _set_family_value(selected_family, material_txt=(selected_stock or {}).get("material", ""))
                    family_profile = _material_family_profile(quality_txt or (selected_stock or {}).get("material", ""), _current_family_key() or selected_family)
                    densidade_display.setText(_float_text(family_profile.get("density", 0), 3))
                    manual_payload = dict(selected_stock or {})
                    if not has_selected_stock:
                        manual_payload = _current_geometry_payload()
                    preview = dict(self.backend.material_price_preview(manual_payload) or {})
                    current_preview["data"] = preview
                    _apply_preview_controls(preview, selected_stock=has_selected_stock, hydrate_dimensions=has_selected_stock)
                    if not has_selected_stock:
                        preco_spin.setValue(_compute_unit_price(formato_txt, _selected_metric_value(preview), float(price_base_spin.value() or 0)))
                    _set_auto_description(_auto_description())
                finally:
                    sync["busy"] = False

            def _apply_payload(row: dict[str, Any], *, prefer_existing_prices: bool = False) -> None:
                if not isinstance(row, dict):
                    return
                sync["busy"] = True
                try:
                    ref_edit.setText(str(row.get("id", "") or "").strip())
                    code_display.setText(str(row.get("id", "") or "").strip())
                    lote_display.setText(str(row.get("lote", "") or "").strip())
                    local_display.setText(str(row.get("localizacao", "") or "").strip())
                    unid_edit.setText(str(row.get("unid", "UN") or "UN").strip() or "UN")
                    formato_txt = str(row.get("formato", "") or "").strip() or "Chapa"
                    price_base_label.setText(_price_base_label(formato_txt))
                    _set_family_value(row.get("material_familia", row.get("material_familia_resolved", "")), material_txt=row.get("material", ""))
                    family_profile = _material_family_profile(row.get("material", ""), str(row.get("material_familia", row.get("material_familia_resolved", "")) or "").strip())
                    densidade_display.setText(_float_text(family_profile.get("density", 0), 3))
                    preview = dict(self.backend.material_price_preview(row) or {})
                    current_preview["data"] = preview
                    _apply_preview_controls(preview, selected_stock=True, hydrate_dimensions=True)
                    price_base_spin.setValue(
                        float(initial.get("p_compra", row.get("preco_base", 0)) or 0)
                        if prefer_existing_prices
                        else float(row.get("preco_base", 0) or 0)
                    )
                    preco_spin.setValue(
                        float(initial.get("preco", row.get("preco", 0)) or 0)
                        if prefer_existing_prices
                        else float(row.get("preco", 0) or 0)
                    )
                    lote_display.setReadOnly(True)
                    local_display.setReadOnly(True)
                    _set_auto_description(_auto_description(), force=True)
                finally:
                    sync["busy"] = False

            def _refresh_stock_combo(preserve_stock_id: str = "") -> None:
                rows = _current_rows()
                stock_combo.blockSignals(True)
                stock_combo.clear()
                stock_combo.addItem("", None)
                for row in rows:
                    stock_combo.addItem(_stock_label(row), row)
                stock_combo.blockSignals(False)
                target_id = preserve_stock_id or current_ref
                if target_id:
                    for idx in range(stock_combo.count()):
                        payload = stock_combo.itemData(idx)
                        if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == target_id:
                            stock_combo.setCurrentIndex(idx)
                            break
                elif rows:
                    stock_combo.setCurrentIndex(1)
                payload = stock_combo.currentData()
                if isinstance(payload, dict):
                    _apply_payload(payload, prefer_existing_prices=bool(current_ref and str(payload.get("id", "") or "").strip() == current_ref))
                else:
                    _refresh_manual_state()

            def _refresh_dimensions() -> None:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                if not formato_txt:
                    _set_combo_items(espessura_combo, [], _clean((current_payload or {}).get("espessura", "")), allow_blank=True)
                    _set_combo_items(dimensao_combo, [], _clean((current_payload or {}).get("dimensao", "")), allow_blank=True)
                    _refresh_stock_combo()
                    return
                candidate_rows = [row for row in material_rows if _matches(row, formato_txt, quality_txt, "", "")]
                esp_values = _unique(candidate_rows, lambda row: _clean(row.get("espessura", ""), "-"))
                current_esp = _combo_text(espessura_combo) or _clean((current_payload or {}).get("espessura", ""))
                _set_combo_items(espessura_combo, esp_values, current_esp, allow_blank=True)
                chosen_esp = _combo_text(espessura_combo)
                dim_rows = [row for row in candidate_rows if _matches(row, formato_txt, quality_txt, chosen_esp, "")]
                dim_values = _unique(dim_rows, _dimension_text)
                if formato_txt == "Chapa":
                    merged_dims: list[str] = []
                    for value in [*common_sheet_dimensions, *dim_values]:
                        if value and value not in merged_dims:
                            merged_dims.append(value)
                    dim_values = merged_dims
                current_dim = _combo_text(dimensao_combo) or (_dimension_text(current_payload or {}) if current_payload else "")
                _set_combo_items(dimensao_combo, dim_values, current_dim, allow_blank=True)
                _refresh_stock_combo()

            def _refresh_quality() -> None:
                formato_txt = _combo_text(formato_combo)
                if not formato_txt:
                    _set_combo_items(qualidade_combo, [], _clean((current_payload or {}).get("material", "")), allow_blank=True)
                    _refresh_dimensions()
                    return
                candidate_rows = [row for row in material_rows if _matches(row, formato_txt, "", "", "")]
                qualities = _unique(candidate_rows, lambda row: row.get("material", ""))
                current_quality = _combo_text(qualidade_combo) or _clean((current_payload or {}).get("material", ""))
                _set_combo_items(qualidade_combo, qualities, current_quality, allow_blank=True)
                _refresh_dimensions()

            def _refresh_formats() -> None:
                formats: list[str] = []
                for value in ["Chapa", "Perfil", "Tubo", "Cantoneira", "Barra", *_unique(material_rows, lambda row: row.get("formato", ""))]:
                    if value and value not in formats:
                        formats.append(value)
                current_format = _clean((current_payload or {}).get("formato", ""))
                _set_combo_items(formato_combo, formats, current_format, allow_blank=True)
                _refresh_quality()

            def _on_stock_change() -> None:
                payload = stock_combo.currentData()
                if isinstance(payload, dict):
                    _apply_payload(payload)
                else:
                    _refresh_manual_state()

            def _on_base_price_changed(value: float) -> None:
                if sync["busy"]:
                    return
                preview = dict(current_preview.get("data") or {})
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                metric_value = _selected_metric_value(preview)
                sync["busy"] = True
                try:
                    preco_spin.setValue(_compute_unit_price(formato_txt, metric_value, value))
                finally:
                    sync["busy"] = False

            def _on_unit_price_changed(value: float) -> None:
                if sync["busy"]:
                    return
                preview = dict(current_preview.get("data") or {})
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                metric_value = _selected_metric_value(preview)
                sync["busy"] = True
                try:
                    price_base_spin.setValue(_compute_base_price(formato_txt, metric_value, value))
                finally:
                    sync["busy"] = False

            def _on_description_edited(_text: str) -> None:
                if sync["busy"]:
                    return
                desc_state["manual"] = True

            formato_combo.currentTextChanged.connect(lambda _text: _refresh_quality())
            qualidade_combo.currentTextChanged.connect(lambda _text: _refresh_dimensions())
            familia_combo.currentIndexChanged.connect(lambda _i: _refresh_manual_state())
            espessura_combo.currentTextChanged.connect(lambda _text: _refresh_dimensions())
            dimensao_combo.currentTextChanged.connect(lambda _text: _refresh_stock_combo())
            secao_tipo_combo.currentTextChanged.connect(lambda _text: _refresh_manual_state())
            stock_combo.currentIndexChanged.connect(lambda _i: _on_stock_change())
            price_base_spin.valueChanged.connect(_on_base_price_changed)
            preco_spin.valueChanged.connect(_on_unit_price_changed)
            for edit in (comprimento_edit, largura_edit, altura_edit, diametro_edit, metros_edit, kg_m_edit):
                edit.textChanged.connect(lambda _text: _refresh_manual_state())
            desc_edit.textEdited.connect(_on_description_edited)
            supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
            local_display.setText(str(initial.get("localizacao", "") or "").strip())
            lote_display.setText(str(initial.get("lote_fornecedor", "") or "").strip())
            _set_family_value(str(initial.get("material_familia", (current_payload or {}).get("material_familia", "")) or "").strip(), material_txt=(current_payload or {}).get("material", ""))
            if current_payload is not None and not current_ref:
                preview = dict(self.backend.material_price_preview(current_payload) or {})
                current_preview["data"] = preview
                _apply_preview_controls(preview, selected_stock=False, hydrate_dimensions=True)
                price_base_spin.setValue(float(initial.get("p_compra", current_payload.get("preco_base", 0)) or 0))
                preco_spin.setValue(float(initial.get("preco", current_payload.get("preco", 0)) or 0))
            _refresh_formats()
        else:
            intro = QLabel(
                "Linha manual sem ligação direta ao stock. Define os dados comerciais e escolhe se o preço base é por unidade, kg ou metro."
            )
            intro.setWordWrap(True)
            intro.setProperty("role", "muted")
            layout.addWidget(intro)

            data_card = CardFrame()
            data_layout = QFormLayout(data_card)
            data_layout.setContentsMargins(12, 10, 12, 10)
            data_layout.setHorizontalSpacing(10)
            data_layout.setVerticalSpacing(8)
            data_hint = QLabel("A linha manual fica autónoma, mas com a mesma lógica de preço base da matéria-prima e dos produtos.")
            data_hint.setWordWrap(True)
            data_hint.setProperty("role", "muted")
            data_layout.addRow(data_hint)
            data_layout.addRow("Ref.", ref_edit)
            data_layout.addRow("Descrição", desc_edit)
            data_layout.addRow("Fornecedor linha", supplier_combo)
            data_layout.addRow("Quantidade", qtd_spin)
            data_layout.addRow("Unid.", unid_edit)
            layout.addWidget(data_card)

            pricing_card = CardFrame()
            pricing_card.set_tone("info")
            pricing_layout = QFormLayout(pricing_card)
            pricing_layout.setContentsMargins(12, 10, 12, 10)
            pricing_layout.setHorizontalSpacing(10)
            pricing_layout.setVerticalSpacing(8)
            pricing_hint = QLabel("Se a compra vier por kg ou metro, indica a métrica por unidade para a ponte automática com o preço unitário.")
            pricing_hint.setWordWrap(True)
            pricing_hint.setProperty("role", "muted")
            pricing_layout.addRow(pricing_hint)
            manual_basis_combo = QComboBox()
            manual_basis_combo.addItem("Unidade", "unit")
            manual_basis_combo.addItem("Kg", "kg")
            manual_basis_combo.addItem("Metro", "meter")
            manual_price_base_label = QLabel("Preço base (EUR/un)")
            manual_price_base_spin = QDoubleSpinBox()
            manual_price_base_spin.setRange(0.0, 1000000.0)
            manual_price_base_spin.setDecimals(4)
            manual_metric_label = QLabel("Métrica por unid.")
            manual_metric_spin = QDoubleSpinBox()
            manual_metric_spin.setRange(0.0, 1000000.0)
            manual_metric_spin.setDecimals(4)
            pricing_layout.addRow("Base de preço", manual_basis_combo)
            pricing_layout.addRow(manual_price_base_label, manual_price_base_spin)
            pricing_layout.addRow(manual_metric_label, manual_metric_spin)
            pricing_layout.addRow("Preço unid. (EUR)", preco_spin)
            pricing_layout.addRow("Desc. %", desconto_spin)
            pricing_layout.addRow("IVA %", iva_spin)
            layout.addWidget(pricing_card)

            manual_sync = {"busy": False}

            def _manual_basis() -> str:
                return str(manual_basis_combo.currentData() or "unit").strip() or "unit"

            def _manual_metric_value() -> float:
                return float(manual_metric_spin.value() or 0)

            def _manual_base_label_text(basis: str) -> str:
                if basis == "kg":
                    return "Preço kg (EUR/kg)"
                if basis == "meter":
                    return "Preço metro (EUR/m)"
                return "Preço base (EUR/un)"

            def _manual_metric_label_text(basis: str) -> str:
                if basis == "kg":
                    return "Kg por unid."
                if basis == "meter":
                    return "Metros por unid. (m)"
                return "Métrica por unid."

            def _manual_compute_unit_price(basis: str, metric_value: float, base_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                if basis in {"kg", "meter"}:
                    return round(base_value * metric_value, 4) if metric_value > 0 else 0.0
                return round(base_value, 4)

            def _manual_compute_base_price(basis: str, metric_value: float, unit_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                if basis in {"kg", "meter"}:
                    return round(unit_value / metric_value, 6) if metric_value > 0 else round(unit_value, 6)
                return round(unit_value, 6)

            def _refresh_manual_price_state(*, recalc_from_unit: bool = False) -> None:
                basis = _manual_basis()
                show_metric = basis in {"kg", "meter"}
                manual_price_base_label.setText(_manual_base_label_text(basis))
                manual_metric_label.setText(_manual_metric_label_text(basis))
                manual_metric_label.setVisible(show_metric)
                manual_metric_spin.setVisible(show_metric)
                if manual_sync["busy"]:
                    return
                if basis == "unit":
                    manual_sync["busy"] = True
                    try:
                        manual_price_base_spin.setValue(float(preco_spin.value() or 0))
                    finally:
                        manual_sync["busy"] = False
                elif recalc_from_unit:
                    _on_manual_unit_price_changed(preco_spin.value())

            def _on_manual_base_price_changed(value: float) -> None:
                if manual_sync["busy"]:
                    return
                manual_sync["busy"] = True
                try:
                    preco_spin.setValue(_manual_compute_unit_price(_manual_basis(), _manual_metric_value(), value))
                finally:
                    manual_sync["busy"] = False

            def _on_manual_unit_price_changed(value: float) -> None:
                if manual_sync["busy"]:
                    return
                manual_sync["busy"] = True
                try:
                    manual_price_base_spin.setValue(_manual_compute_base_price(_manual_basis(), _manual_metric_value(), value))
                finally:
                    manual_sync["busy"] = False

            def _on_manual_metric_changed(_value: float) -> None:
                if manual_sync["busy"]:
                    return
                _refresh_manual_price_state(recalc_from_unit=True)

            def _on_manual_basis_changed(_index: int) -> None:
                if manual_sync["busy"]:
                    return
                _refresh_manual_price_state(recalc_from_unit=True)

            def _accept_manual() -> None:
                if not desc_edit.text().strip():
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a descrição da linha.")
                    return
                if _manual_basis() in {"kg", "meter"} and _manual_metric_value() <= 0:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a métrica por unidade para preço por kg ou metro.")
                    return
                dialog.accept()

            initial_basis = str(initial.get("price_basis", "") or "").strip().lower()
            if initial_basis in {"kg", "peso"}:
                initial_basis = "kg"
            elif initial_basis in {"meter", "metro", "metros"}:
                initial_basis = "meter"
            elif float(initial.get("peso_unid", 0) or 0) > 0:
                initial_basis = "kg"
            elif float(initial.get("metros_unidade", initial.get("metros", 0)) or 0) > 0:
                initial_basis = "meter"
            else:
                initial_basis = "unit"

            initial_metric = 0.0
            if initial_basis == "kg":
                initial_metric = float(initial.get("peso_unid", 0) or 0)
            elif initial_basis == "meter":
                initial_metric = float(initial.get("metros_unidade", initial.get("metros", 0)) or 0)

            initial_base_price = (
                float(initial.get("p_compra", 0) or 0)
                if "p_compra" in initial
                else _manual_compute_base_price(initial_basis, initial_metric, float(initial.get("preco", 0) or 0))
            )
            basis_index = {"unit": 0, "kg": 1, "meter": 2}.get(initial_basis, 0)
            manual_basis_combo.setCurrentIndex(basis_index)
            manual_metric_spin.setValue(initial_metric)
            manual_price_base_spin.setValue(initial_base_price)
            _refresh_manual_price_state(recalc_from_unit=False)

            manual_basis_combo.currentIndexChanged.connect(_on_manual_basis_changed)
            manual_price_base_spin.valueChanged.connect(_on_manual_base_price_changed)
            preco_spin.valueChanged.connect(_on_manual_unit_price_changed)
            manual_metric_spin.valueChanged.connect(_on_manual_metric_changed)
            accept_handler = _accept_manual
            manual_price_controls = {
                "basis": manual_basis_combo,
                "base_price": manual_price_base_spin,
                "metric": manual_metric_spin,
            }
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_button = buttons.button(QDialogButtonBox.Ok)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if ok_button is not None:
            ok_button.setText("Guardar")
        if cancel_button is not None:
            cancel_button.setText("Fechar")
        buttons.accepted.connect(accept_handler)
        buttons.rejected.connect(dialog.reject)
        dialog_layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = {
            "ref": ref_edit.text().strip(),
            "descricao": desc_edit.text().strip(),
            "fornecedor_linha": supplier_combo.currentText().strip(),
            "origem": "Matéria-Prima" if material_mode else str(initial.get("origem", "Produto") or "Produto"),
            "qtd": qtd_spin.value(),
            "unid": unid_edit.text().strip() or "UN",
            "preco": preco_spin.value(),
            "desconto": desconto_spin.value(),
            "iva": iva_spin.value(),
            "entrega": str(initial.get("entrega", "PENDENTE") or "PENDENTE"),
        }
        if material_mode:
            selected_material = stock_combo.currentData()
            if isinstance(selected_material, dict):
                preview = dict(self.backend.material_price_preview(selected_material) or {})
                payload.update(
                    {
                        "formato": str(selected_material.get("formato", "") or "").strip(),
                        "material": str(selected_material.get("material", "") or "").strip(),
                        "material_familia": str(selected_material.get("material_familia", selected_material.get("material_familia_resolved", "")) or "").strip(),
                        "espessura": str(selected_material.get("espessura", "") or "").strip(),
                        "comprimento": float(preview.get("comprimento", selected_material.get("comprimento", 0)) or 0),
                        "largura": float(preview.get("largura", selected_material.get("largura", 0)) or 0),
                        "altura": float(preview.get("altura", selected_material.get("altura", 0)) or 0),
                        "diametro": float(preview.get("diametro", selected_material.get("diametro", 0)) or 0),
                        "metros": float(preview.get("metros", selected_material.get("metros", 0)) or 0),
                        "kg_m": float(preview.get("kg_m", selected_material.get("kg_m", 0)) or 0),
                        "peso_unid": float(preview.get("peso_unid", selected_material.get("peso_unid", 0)) or 0),
                        "secao_tipo": str(preview.get("secao_tipo", selected_material.get("secao_tipo", "")) or "").strip(),
                        "p_compra": float(price_base_spin.value() or 0),
                        "localizacao": str(selected_material.get("localizacao", "") or "").strip(),
                        "lote_fornecedor": str(selected_material.get("lote", "") or "").strip(),
                    }
                )
            else:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                manual_raw = {
                    **_current_geometry_payload(),
                    "quantidade": 1,
                    "reservado": 0,
                    "p_compra": float(price_base_spin.value() or 0),
                    "local": local_display.text().strip(),
                    "lote_fornecedor": lote_display.text().strip(),
                }
                manual_desc = desc_edit.text().strip() or _auto_description()
                if not formato_txt:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica o tipo de material.")
                    return None
                if not quality_txt:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a qualidade da matéria-prima.")
                    return None
                try:
                    normalized_material = dict(self.backend._normalise_material_payload(manual_raw) or {})
                except Exception as exc:
                    QMessageBox.warning(dialog, "Notas Encomenda", str(exc))
                    return None
                payload.update(
                    {
                        "ref": "",
                        "descricao": manual_desc,
                        "formato": formato_txt,
                        "material": quality_txt,
                        "material_familia": str(normalized_material.get("material_familia", _current_family_key()) or "").strip(),
                        "espessura": str(normalized_material.get("espessura", esp_txt) or "").strip(),
                        "comprimento": float(normalized_material.get("comprimento", 0) or 0),
                        "largura": float(normalized_material.get("largura", 0) or 0),
                        "altura": float(normalized_material.get("altura", 0) or 0),
                        "diametro": float(normalized_material.get("diametro", 0) or 0),
                        "metros": float(normalized_material.get("metros", 0) or 0),
                        "kg_m": float(normalized_material.get("kg_m", 0) or 0),
                        "peso_unid": float(normalized_material.get("peso_unid", 0) or 0),
                        "secao_tipo": str(normalized_material.get("secao_tipo", "") or "").strip(),
                        "p_compra": float(price_base_spin.value() or 0),
                        "localizacao": local_display.text().strip(),
                        "lote_fornecedor": lote_display.text().strip(),
                        "_material_manual": True,
                        "_material_pending_create": True,
                    }
                )
        elif manual_price_controls is not None:
            basis = str(manual_price_controls["basis"].currentData() or "unit").strip() or "unit"
            metric_value = float(manual_price_controls["metric"].value() or 0)
            payload.update(
                {
                    "p_compra": float(manual_price_controls["base_price"].value() or 0),
                    "price_basis": basis,
                    "peso_unid": metric_value if basis == "kg" else 0.0,
                    "metros": metric_value if basis == "meter" else 0.0,
                    "metros_unidade": metric_value if basis == "meter" else 0.0,
                }
            )
        return payload

    def _quick_assign_supplier(self) -> None:
        if not self.line_rows:
            QMessageBox.warning(self, "Notas Encomenda", "Não existem linhas para adjudicar.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Fornecedor rápido")
        dialog.setMinimumWidth(460)
        layout = QVBoxLayout(dialog)
        info = QLabel("Aplica o fornecedor às linhas selecionadas. Se não houver seleção, aplica a todas as linhas.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(self.supplier_combo.currentText().strip())
        layout.addWidget(supplier_combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        supplier_text = supplier_combo.currentText().strip()
        if not supplier_text:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona um fornecedor.")
            return
        target_indexes = self._selected_line_indexes() or list(range(len(self.line_rows)))
        for idx in target_indexes:
            self.line_rows[idx]["fornecedor_linha"] = supplier_text
        self._render_lines()

    def _add_material_line(self) -> None:
        payload = self._line_dialog("Adicionar linha de Matéria-Prima", {"origem": "Matéria-Prima", "iva": 23, "qtd": 1}, material_mode=True)
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _add_product_line(self) -> None:
        payload = self._product_line_dialog("Adicionar produto em stock", {"origem": "Produto", "iva": 23, "qtd": 1})
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _add_manual_line(self) -> None:
        payload = self._line_dialog("Adicionar linha manual", {"origem": "Produto", "iva": 23, "qtd": 1}, material_mode=False)
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _edit_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma linha.")
            return
        current = dict(self.line_rows[index])
        is_material = self.backend.desktop_main.origem_is_materia(current.get("origem", ""))
        if str(current.get("origem", "") or "").strip().lower() == "produto":
            payload = self._product_line_dialog("Editar linha de produto", current)
        else:
            payload = self._line_dialog("Editar linha", current, material_mode=is_material)
        if payload is None:
            return
        payload["origem"] = current.get("origem", payload.get("origem", "Produto"))
        payload["entrega"] = current.get("entrega", "PENDENTE")
        self.line_rows[index] = payload
        self._render_lines()
        self.lines_table.selectRow(index)

    def _remove_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma linha.")
            return
        del self.line_rows[index]
        self._render_lines()

    def _note_payload(self) -> dict:
        supplier_text = self.supplier_combo.currentText().strip()
        supplier = self._supplier_lookup(supplier_text)
        payload = {
            "numero": self.current_number,
            "fornecedor": supplier_text or str(supplier.get("nome", "") or ""),
            "fornecedor_id": str(supplier.get("id", "") or ""),
            "contacto": self.contact_edit.text().strip(),
            "data_entrega": self._delivery_text(),
            "obs": self.obs_edit.toPlainText().strip(),
            "local_descarga": self.location_edit.currentText().strip(),
            "meio_transporte": self.transport_edit.currentText().strip(),
            "lines": self.line_rows,
        }
        return payload

    def _save_note(self) -> None:
        try:
            note = self.backend.ne_save(self._note_payload())
        except Exception as exc:
            QMessageBox.critical(self, "Guardar NE", str(exc))
            return
        self.current_number = str(note.get("numero", "")).strip()
        self.refresh()
        QMessageBox.information(self, "Guardar NE", f"Nota {self.current_number} guardada.")

    def _approve_note(self) -> None:
        try:
            self.backend.ne_save(self._note_payload())
            note = self.backend.ne_approve(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Aprovar NE", str(exc))
            return
        self.current_number = str(note.get("numero", "")).strip()
        self.refresh()
        QMessageBox.information(self, "Aprovar NE", f"Nota {self.current_number} aprovada.")

    def _remove_note(self) -> None:
        if not self.current_number:
            return
        if QMessageBox.question(self, "Apagar NE", f"Remover nota {self.current_number}?") != QMessageBox.Yes:
            return
        try:
            self.backend.ne_remove(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Apagar NE", str(exc))
            return
        self.current_number = ""
        self.refresh()

    def _open_pdf(self, quote: bool) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            path = self.backend.ne_open_pdf(self.current_number, quote=quote)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        QMessageBox.information(self, "Notas Encomenda", f"PDF aberto:\n{path}")

    def _register_delivery(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Guarda ou seleciona primeiro a nota.")
            return
        payload = self._delivery_dialog()
        if payload is None:
            return
        try:
            self.backend.ne_save(self._note_payload())
            self.backend.ne_register_delivery(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Entrega NE", str(exc))
            return
        self.refresh()
        QMessageBox.information(self, "Entrega NE", f"Entrega registada para {self.current_number}.")

    def _generate_supplier_orders(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Guarda ou seleciona primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            created = self.backend.ne_generate_supplier_orders(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Gerar NEs", str(exc))
            return
        self.refresh()
        summary = ", ".join(f"{row.get('numero')} ({row.get('fornecedor')})" for row in created)
        QMessageBox.information(self, "Gerar NEs", f"Notas geradas: {summary}")

    def _split_email_tokens(self, raw: str) -> list[str]:
        tokens = []
        seen: set[str] = set()
        for part in re.split(r"[;\n,]+", str(raw or "").strip()):
            value = str(part or "").strip()
            if not value:
                continue
            key = value.lower()
            if key in seen:
                continue
            seen.add(key)
            tokens.append(value)
        return tokens

    def _extract_first_email(self, *values: object) -> str:
        for value in values:
            match = re.search(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", str(value or ""), re.IGNORECASE)
            if match:
                return str(match.group(0) or "").strip()
        return ""

    def _default_outbound_email(self) -> str:
        fixed_default = "Engenharia@barcelbal.com"
        if self._rfq_outlook_to:
            return self._rfq_outlook_to
        if fixed_default:
            return fixed_default
        branding = dict(getattr(self.backend, "branding", {}) or {})
        direct = self._extract_first_email(
            branding.get("company_email", ""),
            branding.get("email", ""),
        )
        if direct:
            return direct
        for row in list(branding.get("empresa_info_rodape", []) or []):
            found = self._extract_first_email(row)
            if found:
                return found
        try:
            for row in list(self.backend.desktop_main.get_empresa_rodape_lines() or []):
                found = self._extract_first_email(row)
                if found:
                    return found
        except Exception:
            pass
        return ""

    def _business_company_name(self) -> str:
        branding = dict(getattr(self.backend, "branding_settings", lambda: {})() or {})
        emit = dict(branding.get("guia_emitente", {}) or {})
        branding_public = dict(getattr(self.backend, "branding", {}) or {})
        return (
            str(emit.get("nome", "") or "").strip()
            or str(branding_public.get("company_name", "") or "").strip()
            or "luGEST"
        )

    def _business_primary_color(self) -> str:
        branding = dict(getattr(self.backend, "branding_settings", lambda: {})() or {})
        return str(branding.get("primary_color", "") or "#000040").strip() or "#000040"

    def _quote_recipient_label(self, detail: dict[str, object] | None = None) -> str:
        supplier_context = self._note_quote_supplier_targets(detail)
        suppliers = list(supplier_context.get("suppliers", []) or [])
        if len(suppliers) == 1:
            return str(suppliers[0].get("nome", "") or "").strip() or "Exmos. Senhores"
        return "Exmos. Senhores"

    def _order_recipient_label(self, detail: dict[str, object] | None = None) -> str:
        supplier = self._note_order_supplier(detail)
        return str(supplier.get("nome", "") or "").strip() or "Exmos. Senhores"

    def _build_commercial_email_html(
        self,
        *,
        title: str,
        reference: str,
        greeting: str,
        intro_lines: list[str],
        summary_title: str,
        summary_rows: list[dict[str, str]],
        note_title: str = "",
        note_lines: list[str] | None = None,
        logo_cid: str = "",
    ) -> str:
        company = self._business_company_name()
        primary = self._business_primary_color()
        header_logo = (
            "<table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;\">"
            "<tr><td style=\"height:10px; line-height:10px; font-size:0;\">&nbsp;</td></tr>"
            "<tr><td>"
            f"<img src=\"cid:{html.escape(logo_cid)}\" alt=\"{html.escape(company)}\" width=\"109\" "
            "style=\"display:block; width:109px; height:auto; border:0; outline:none; text-decoration:none;\" />"
            "</td></tr>"
            "</table>"
            if logo_cid
            else f"<div style=\"font-size:24px; font-weight:800; letter-spacing:-0.5px; color:#ffffff;\">{html.escape(company)}</div>"
        )
        intro_html = "".join(
            f"<p style=\"margin:0 0 16px 0; font-size:14px; line-height:1.7; color:#334155;\">{html.escape(row)}</p>"
            for row in [str(value or "").strip() for value in intro_lines if str(value or "").strip()]
        )
        note_block = ""
        clean_notes = [str(value or "").strip() for value in list(note_lines or []) if str(value or "").strip()]
        if clean_notes:
            note_items = "".join(
                f"<li style=\"margin:0 0 6px 0;\">{html.escape(row)}</li>"
                for row in clean_notes
            )
            note_block = (
                "<div style=\"margin:24px 0 0 0; padding:16px 18px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px;\">"
                f"<div style=\"margin:0 0 10px 0; font-size:11px; font-weight:800; color:#94a3b8; letter-spacing:0.55px; text-transform:uppercase;\">{html.escape(note_title or 'Notas')}</div>"
                f"<ul style=\"margin:0; padding-left:18px; font-size:13px; line-height:1.55; color:#334155;\">{note_items}</ul>"
                "</div>"
            )
        summary_html_rows = []
        greeting_txt = str(greeting or "").strip()
        if greeting_txt.lower().startswith("exm"):
            greeting_title = greeting_txt
        else:
            greeting_title = f"Exmo(a). {greeting_txt}"
        for row in list(summary_rows or []):
            label = str(row.get("label", "") or "").strip()
            value = str(row.get("value", "") or "").strip() or "-"
            emphasis = str(row.get("emphasis", "") or "").strip().lower()
            value_style = "font-size:13px; color:#0f172a;"
            if emphasis == "money":
                value_style = "font-size:26px; font-weight:800; color:#16a34a;"
            elif emphasis == "strong":
                value_style = "font-size:14px; font-weight:700; color:#0f172a;"
            summary_html_rows.append(
                "<tr>"
                f"<td style=\"padding:12px 14px; border-top:1px solid #e2e8f0; font-size:13px; color:#475569;\">{html.escape(label)}</td>"
                f"<td style=\"padding:12px 14px; border-top:1px solid #e2e8f0; {value_style}\">{html.escape(value)}</td>"
                "</tr>"
            )
        summary_html = "".join(summary_html_rows)
        return (
            "<html><body style=\"margin:0; padding:0; background:#eef2f7; font-family:Segoe UI, Arial, sans-serif; color:#334155;\">"
            "<div style=\"padding:26px 0;\">"
            "<div style=\"width:660px; margin:0 auto; background:#ffffff; border-radius:20px; overflow:hidden; box-shadow:0 16px 40px rgba(15,23,42,0.10);\">"
            f"<div style=\"background:{html.escape(primary)}; padding:22px 28px;\">"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;\">"
            "<tr>"
            f"<td valign=\"bottom\" style=\"vertical-align:bottom; width:72%; height:112px;\">{header_logo}</td>"
            "<td valign=\"top\" style=\"vertical-align:top; text-align:right; color:#dbe4f0; font-size:10px; letter-spacing:0.55px; text-transform:uppercase;\">"
            f"{html.escape(title)}<br>"
            f"<span style=\"display:inline-block; margin-top:3px; padding-left:16px; border-left:1px solid rgba(255,255,255,0.5); font-size:18px; font-weight:800; color:#ffffff;\">{html.escape(reference or '-')}</span>"
            "</td>"
            "</tr>"
            "</table>"
            "</div>"
            "<div style=\"padding:28px 30px 22px 30px;\">"
            f"<p style=\"margin:0 0 18px 0; font-size:19px; font-weight:800; color:#0f172a;\">{html.escape(greeting_title)},</p>"
            "<p style=\"margin:0 0 16px 0; font-size:15px; line-height:1.7; color:#334155;\">Boa tarde,</p>"
            f"{intro_html}"
            f"{note_block}"
            f"<div style=\"margin:24px 0 10px 0; font-size:11px; font-weight:800; color:#94a3b8; letter-spacing:0.6px; text-transform:uppercase;\">{html.escape(summary_title)}</div>"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden;\">"
            "<tr style=\"background:#1f2933; color:#ffffff; font-size:11px; font-weight:800; text-transform:uppercase;\">"
            "<td style=\"padding:10px 14px;\">Campo</td>"
            "<td style=\"padding:10px 14px;\">Valor</td>"
            "</tr>"
            f"{summary_html}"
            "</table>"
            "</div>"
            "<div style=\"padding:12px 18px; background:#f8fafc; border-top:1px solid #e2e8f0; text-align:center;\">"
            f"<div style=\"font-size:11px; color:#64748b;\">{html.escape(company)}</div>"
            "<div style=\"margin-top:2px; font-size:10px; color:#94a3b8;\">Gerado automaticamente pelo programa luGEST</div>"
            "</div>"
            "</div>"
            "</div>"
            "</body></html>"
        )

    def _note_quote_supplier_targets(self, detail: dict[str, object] | None = None) -> dict[str, object]:
        payload = dict(detail or {})
        unique_rows: dict[str, dict[str, str]] = {}
        missing: list[str] = []
        seen_missing: set[str] = set()
        candidates: list[tuple[str, str]] = []
        supplier_id = str(payload.get("fornecedor_id", "") or "").strip()
        supplier_txt = str(payload.get("fornecedor", "") or "").strip()
        if supplier_id or supplier_txt:
            candidates.append((supplier_id, supplier_txt))
        for line in list(payload.get("lines", []) or []):
            line_supplier = str((line or {}).get("fornecedor_linha", "") or "").strip()
            if line_supplier:
                candidates.append(("", line_supplier))
        for supplier_id_txt, label_txt in candidates:
            supplier = {}
            if supplier_id_txt:
                supplier = next(
                    (row for row in self.supplier_rows if str(row.get("id", "") or "").strip() == supplier_id_txt),
                    {},
                )
            if not supplier and label_txt:
                supplier = self._supplier_lookup(label_txt)
            supplier_name = str((supplier or {}).get("nome", "") or label_txt).strip()
            supplier_key = str((supplier or {}).get("id", "") or supplier_name).strip().lower()
            if not supplier_key:
                continue
            email = str((supplier or {}).get("email", "") or "").strip()
            if email:
                unique_rows[supplier_key] = {
                    "id": str((supplier or {}).get("id", "") or "").strip(),
                    "nome": supplier_name,
                    "email": email,
                }
            elif supplier_name and supplier_key not in seen_missing:
                seen_missing.add(supplier_key)
                missing.append(supplier_name)
        suppliers = list(unique_rows.values())
        suppliers.sort(key=lambda row: str(row.get("nome", "") or "").lower())
        bcc = [str(row.get("email", "") or "").strip() for row in suppliers if str(row.get("email", "") or "").strip()]
        return {"suppliers": suppliers, "bcc": bcc, "missing": missing}

    def _note_quote_line_preview(self, detail: dict[str, object] | None = None, limit: int = 6) -> list[str]:
        payload = dict(detail or {})
        preview: list[str] = []
        for line in list(payload.get("lines", []) or [])[: max(1, limit)]:
            row = dict(line or {})
            ref = str(row.get("ref", "") or "").strip()
            material = str(row.get("material", "") or "").strip()
            desc = str(row.get("descricao", "") or "").strip()
            esp = str(row.get("espessura", "") or "").strip()
            qty = self.backend._fmt(row.get("qtd", 0))
            unit = str(row.get("unid", "") or "").strip() or "UN"
            lead = material or desc or ref or "Linha"
            if material and esp:
                lead = f"{material} {esp} mm"
            elif desc:
                lead = desc
            if ref:
                lead = f"{ref} | {lead}"
            preview.append(f"{lead} | Qtd {qty} {unit}".strip())
        return preview

    def _note_quote_email_subject(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        company = self._business_company_name()
        return f"Pedido de Cotação - {company} [{numero}]" if numero else f"Pedido de Cotação - {company}"

    def _note_quote_email_body(self, detail: dict[str, object] | None = None, *, reply_email: str = "") -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        local_descarga = str(payload.get("local_descarga", "") or "").strip()
        transporte = str(payload.get("meio_transporte", "") or "").strip()
        company = self._business_company_name()
        greeting = self._quote_recipient_label(payload)
        lines = [
            f"{greeting},",
            "",
            "Boa tarde,",
            "",
            f"Segue em anexo o pedido de cotação {numero}." if numero else "Segue em anexo o pedido de cotação.",
            "Agradecemos o envio da vossa melhor proposta para os itens indicados no documento em anexo.",
            "Solicitamos, por favor, indicação de preço, prazo de entrega e condições de pagamento.",
        ]
        if data_entrega:
            lines.append(f"Data pretendida de entrega: {data_entrega}")
        if local_descarga:
            lines.append(f"Local de descarga: {local_descarga}")
        if transporte:
            lines.append(f"Transporte: {transporte}")
        if reply_email:
            lines.extend(["", f"Podem responder diretamente para: {reply_email}"])
        lines.extend(["", "Ficamos ao dispor para qualquer esclarecimento.", "", "Cumprimentos,", company])
        return "\n".join(lines)

    def _note_quote_email_html_body(
        self,
        detail: dict[str, object] | None = None,
        *,
        body_plain: str = "",
        logo_cid: str = "",
    ) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip() or "-"
        local_descarga = str(payload.get("local_descarga", "") or "").strip() or "-"
        transporte = str(payload.get("meio_transporte", "") or "").strip() or "-"
        response_email = self._default_outbound_email()
        preview_lines = self._note_quote_line_preview(payload, limit=4)
        obs_txt = str(payload.get("obs", "") or "").strip()
        note_lines = list(preview_lines)
        if obs_txt:
            note_lines.insert(0, obs_txt)
        return self._build_commercial_email_html(
            title="Pedido de Cotação",
            reference=numero,
            greeting=self._quote_recipient_label(payload),
            intro_lines=[
                f"Segue em anexo o pedido de cotação referente ao documento {numero}." if numero else "Segue em anexo o pedido de cotação solicitado.",
                "Agradecemos o envio da vossa melhor proposta para os itens identificados no documento em anexo.",
                "Solicitamos, por favor, indicação de preço, prazo de entrega e condições de pagamento.",
            ],
            summary_title="Condições do pedido",
            summary_rows=[
                {"label": "Documento", "value": numero or "-", "emphasis": "strong"},
                {"label": "Entrega pretendida", "value": data_entrega},
                {"label": "Local de descarga", "value": local_descarga},
                {"label": "Transporte", "value": transporte},
                {"label": "Responder para", "value": response_email or "-", "emphasis": "strong"},
            ],
            note_title="Notas e itens em destaque",
            note_lines=note_lines,
            logo_cid=logo_cid,
        )

    def _note_quote_email_dialog(self, detail: dict[str, object]) -> dict[str, object] | None:
        supplier_context = self._note_quote_supplier_targets(detail)
        dialog = QDialog(self)
        dialog.setWindowTitle("Pedido de cotação | Outlook")
        dialog.setMinimumWidth(860)
        dialog.resize(920, 700)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        intro = QLabel(
            "O pedido vai ser preparado para Outlook com o PDF de cotação em anexo. "
            "O campo BCC é preenchido automaticamente com os emails encontrados nos fornecedores da nota e continua editável."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        info_card = CardFrame()
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(12, 10, 12, 10)
        info_layout.setSpacing(6)
        supplier_names = ", ".join(
            str(row.get("nome", "") or "").strip()
            for row in list(supplier_context.get("suppliers", []) or [])
            if str(row.get("nome", "") or "").strip()
        ) or "Sem fornecedores com email detetados automaticamente."
        info_main = QLabel(f"Fornecedores em BCC: {supplier_names}")
        info_main.setWordWrap(True)
        info_layout.addWidget(info_main)
        missing_rows = list(supplier_context.get("missing", []) or [])
        if missing_rows:
            missing_label = QLabel("Sem email na ficha: " + ", ".join(missing_rows))
            missing_label.setWordWrap(True)
            missing_label.setStyleSheet("color: #b54708; font-weight: 700;")
            info_layout.addWidget(missing_label)
        layout.addWidget(info_card)

        to_edit = QLineEdit(self._default_outbound_email())
        cc_edit = QLineEdit(self._rfq_outlook_cc)
        bcc_edit = QTextEdit("; ".join(list(supplier_context.get("bcc", []) or [])))
        bcc_edit.setMaximumHeight(84)
        add_supplier_combo = QComboBox()
        self._configure_supplier_selector(add_supplier_combo)
        add_supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_label = f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -")
            add_supplier_combo.addItem(supplier_label)
        add_supplier_btn = QPushButton("Adicionar fornecedor ao BCC")
        add_supplier_btn.setProperty("variant", "secondary")
        bcc_list = QListWidget()
        bcc_list.setMaximumHeight(120)
        bcc_list.setSelectionMode(QAbstractItemView.SingleSelection)
        remove_bcc_btn = QPushButton("Remover selecionado")
        remove_bcc_btn.setProperty("variant", "secondary")
        subject_edit = QLineEdit(self._note_quote_email_subject(detail))
        body_edit = QTextEdit(self._note_quote_email_body(detail, reply_email=to_edit.text().strip()))
        body_edit.setMinimumHeight(280)
        send_now_check = QCheckBox("Enviar diretamente sem abrir o rascunho no Outlook")

        def _bcc_tokens() -> list[str]:
            return self._split_email_tokens(bcc_edit.toPlainText())

        def _set_bcc_tokens(values: list[str]) -> None:
            clean = self._split_email_tokens("; ".join(values))
            bcc_edit.blockSignals(True)
            bcc_edit.setPlainText("; ".join(clean))
            bcc_edit.blockSignals(False)
            bcc_list.clear()
            for value in clean:
                bcc_list.addItem(value)

        def _append_bcc_supplier() -> None:
            supplier = self._supplier_lookup(add_supplier_combo.currentText())
            email = str((supplier or {}).get("email", "") or "").strip()
            if not supplier:
                QMessageBox.warning(dialog, "Pedido de cotação", "Seleciona um fornecedor da base de dados.")
                return
            if not email:
                QMessageBox.warning(dialog, "Pedido de cotação", "O fornecedor selecionado não tem email na ficha.")
                return
            _set_bcc_tokens(_bcc_tokens() + [email])
            add_supplier_combo.setCurrentText("")

        def _remove_selected_bcc() -> None:
            item = bcc_list.currentItem()
            if item is None:
                return
            selected_email = str(item.text() or "").strip().lower()
            _set_bcc_tokens([value for value in _bcc_tokens() if value.strip().lower() != selected_email])

        _set_bcc_tokens(list(supplier_context.get("bcc", []) or []))
        add_supplier_btn.clicked.connect(_append_bcc_supplier)
        remove_bcc_btn.clicked.connect(_remove_selected_bcc)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(10)
        form.addRow("Para", to_edit)
        form.addRow("CC", cc_edit)
        form.addRow("BCC", bcc_edit)
        bcc_tools = QWidget()
        bcc_tools_layout = QVBoxLayout(bcc_tools)
        bcc_tools_layout.setContentsMargins(0, 0, 0, 0)
        bcc_tools_layout.setSpacing(6)
        bcc_add_row = QHBoxLayout()
        bcc_add_row.setContentsMargins(0, 0, 0, 0)
        bcc_add_row.setSpacing(6)
        bcc_add_row.addWidget(add_supplier_combo, 1)
        bcc_add_row.addWidget(add_supplier_btn)
        bcc_tools_layout.addLayout(bcc_add_row)
        bcc_tools_layout.addWidget(bcc_list)
        bcc_tools_layout.addWidget(remove_bcc_btn, 0, Qt.AlignRight)
        form.addRow("Fornecedores BCC", bcc_tools)
        form.addRow("Assunto", subject_edit)
        form.addRow("Mensagem", body_edit)
        layout.addLayout(form)
        layout.addWidget(send_now_check)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn is not None:
            ok_btn.setText("Abrir Outlook")
        if cancel_btn is not None:
            cancel_btn.setText("Cancelar")

        def _sync_ok_text() -> None:
            if ok_btn is not None:
                ok_btn.setText("Enviar" if send_now_check.isChecked() else "Abrir Outlook")

        def _accept() -> None:
            to_tokens = self._split_email_tokens(to_edit.text())
            if not to_tokens:
                QMessageBox.warning(dialog, "Pedido de cotação", "Indica pelo menos o teu email no campo 'Para'.")
                return
            if not subject_edit.text().strip():
                QMessageBox.warning(dialog, "Pedido de cotação", "Indica um assunto para o email.")
                return
            if not body_edit.toPlainText().strip():
                QMessageBox.warning(dialog, "Pedido de cotação", "Indica a mensagem do email.")
                return
            dialog.accept()

        send_now_check.toggled.connect(_sync_ok_text)
        _sync_ok_text()
        buttons.accepted.connect(_accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return None
        to_tokens = self._split_email_tokens(to_edit.text())
        cc_tokens = self._split_email_tokens(cc_edit.text())
        bcc_tokens = _bcc_tokens()
        self._rfq_outlook_to = to_tokens[0] if to_tokens else ""
        self._rfq_outlook_cc = "; ".join(cc_tokens)
        return {
            "to": "; ".join(to_tokens),
            "cc": "; ".join(cc_tokens),
            "bcc": "; ".join(bcc_tokens),
            "subject": subject_edit.text().strip(),
            "body_plain": body_edit.toPlainText().strip(),
            "send_now": bool(send_now_check.isChecked()),
        }

    def _open_note_quote_email_draft(self, detail: dict[str, object], mail_payload: dict[str, object]) -> None:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        if not numero:
            raise ValueError("Guarda primeiro a nota antes de preparar o email.")
        safe_number = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in numero)[:48] or "pedido_cotacao"
        attachment_name = f"Pedido_Cotacao_{safe_number}.pdf"
        attachment_path: Path | None = Path(tempfile.gettempdir()) / attachment_name
        attachment_issue = ""
        try:
            self.backend.ne_render_pdf(numero, quote=True, output_path=attachment_path)
        except Exception as exc:
            attachment_issue = str(exc)
            attachment_path = None

        logo_path = getattr(self.backend, "logo_path", None)
        logo_file = Path(logo_path) if isinstance(logo_path, Path) and logo_path.exists() else None
        logo_cid = "lugest-ne-mail-logo" if logo_file is not None else ""
        subject = str(mail_payload.get("subject", "") or "").strip()
        body_plain = str(mail_payload.get("body_plain", "") or "").strip()
        body_html = self._note_quote_email_html_body(payload, body_plain=body_plain, logo_cid=logo_cid)
        send_now = bool(mail_payload.get("send_now"))

        env = os.environ.copy()
        env["LUGEST_MAIL_TO"] = str(mail_payload.get("to", "") or "").strip()
        env["LUGEST_MAIL_CC"] = str(mail_payload.get("cc", "") or "").strip()
        env["LUGEST_MAIL_BCC"] = str(mail_payload.get("bcc", "") or "").strip()
        env["LUGEST_MAIL_SUBJECT"] = subject
        env["LUGEST_MAIL_BODY"] = body_html
        env["LUGEST_MAIL_ATTACHMENT"] = str(attachment_path) if attachment_path is not None else ""
        env["LUGEST_MAIL_LOGO"] = str(logo_file) if logo_file is not None else ""
        env["LUGEST_MAIL_LOGO_CID"] = logo_cid
        env["LUGEST_MAIL_SEND_NOW"] = "1" if send_now else "0"
        powershell_script = (
            "$ErrorActionPreference='Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            "$mail.To = $env:LUGEST_MAIL_TO; "
            "if ($env:LUGEST_MAIL_CC) { $mail.CC = $env:LUGEST_MAIL_CC }; "
            "if ($env:LUGEST_MAIL_BCC) { $mail.BCC = $env:LUGEST_MAIL_BCC }; "
            "$mail.Subject = $env:LUGEST_MAIL_SUBJECT; "
            "if ($env:LUGEST_MAIL_LOGO -and (Test-Path $env:LUGEST_MAIL_LOGO)) "
            "{ "
            "  $logo = $mail.Attachments.Add($env:LUGEST_MAIL_LOGO); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', $env:LUGEST_MAIL_LOGO_CID); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true) "
            "}; "
            "if ($env:LUGEST_MAIL_ATTACHMENT -and (Test-Path $env:LUGEST_MAIL_ATTACHMENT)) "
            "{ $null = $mail.Attachments.Add($env:LUGEST_MAIL_ATTACHMENT) }; "
            "$mail.HTMLBody = $env:LUGEST_MAIL_BODY; "
            "if ($env:LUGEST_MAIL_SEND_NOW -eq '1') { $mail.Send() } else { $mail.Display() }"
        )
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    powershell_script,
                ],
                check=True,
                env=env,
                timeout=30,
            )
        except Exception:
            mailto = (
                f"mailto:{quote(str(mail_payload.get('to', '') or '').strip())}"
                f"?cc={quote(str(mail_payload.get('cc', '') or '').strip())}"
                f"&bcc={quote(str(mail_payload.get('bcc', '') or '').strip())}"
                f"&subject={quote(subject)}"
                f"&body={quote(body_plain)}"
            )
            os.startfile(mailto)
            fallback_message = "Outlook indisponível. Foi aberto o cliente de email por defeito."
            if attachment_issue:
                fallback_message += f"\n\nTambém não foi possível gerar o PDF em anexo:\n{attachment_issue}"
            else:
                fallback_message += "\n\nNota: o anexo PDF terá de ser adicionado manualmente neste modo."
            QMessageBox.information(self, "Pedido de cotação", fallback_message)
            return

        if attachment_issue:
            QMessageBox.information(
                self,
                "Pedido de cotação",
                f"O email foi preparado no Outlook, mas o PDF não foi anexado automaticamente:\n{attachment_issue}",
            )
        elif send_now:
            QMessageBox.information(self, "Pedido de cotação", "Email enviado com sucesso pelo Outlook.")

    def _note_order_supplier(self, detail: dict[str, object] | None = None) -> dict[str, str]:
        payload = dict(detail or {})
        supplier = self._supplier_lookup(str(payload.get("fornecedor", "") or "").strip())
        if supplier:
            return {
                "id": str(supplier.get("id", "") or "").strip(),
                "nome": str(supplier.get("nome", "") or "").strip(),
                "email": str(supplier.get("email", "") or "").strip(),
                "contacto": str(supplier.get("contacto", "") or "").strip(),
            }
        return {
            "id": str(payload.get("fornecedor_id", "") or "").strip(),
            "nome": str(payload.get("fornecedor", "") or "").strip(),
            "email": "",
            "contacto": str(payload.get("contacto", "") or "").strip(),
        }

    def _note_order_email_subject(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        company = self._business_company_name()
        return f"Encomenda - {company} [{numero}]" if numero else f"Encomenda - {company}"

    def _note_order_email_body(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        local_descarga = str(payload.get("local_descarga", "") or "").strip()
        transporte = str(payload.get("meio_transporte", "") or "").strip()
        company = self._business_company_name()
        greeting = self._order_recipient_label(payload)
        lines = [
            f"{greeting},",
            "",
            "Boa tarde,",
            "",
            f"Segue em anexo a encomenda {numero}." if numero else "Segue em anexo a encomenda.",
            "Agradecemos confirmação de receção e do prazo previsto de entrega.",
        ]
        if data_entrega:
            lines.append(f"Data pretendida de entrega: {data_entrega}")
        if local_descarga:
            lines.append(f"Local de descarga: {local_descarga}")
        if transporte:
            lines.append(f"Transporte: {transporte}")
        lines.extend(["", "Ficamos ao dispor para qualquer esclarecimento.", "", "Cumprimentos,", company])
        return "\n".join(lines)

    def _note_order_email_html_body(
        self,
        detail: dict[str, object] | None = None,
        *,
        body_plain: str = "",
        logo_cid: str = "",
    ) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip() or "-"
        local_descarga = str(payload.get("local_descarga", "") or "").strip() or "-"
        transporte = str(payload.get("meio_transporte", "") or "").strip() or "-"
        total_txt = _fmt_eur(float(payload.get("total", 0) or 0))
        obs_txt = str(payload.get("obs", "") or "").strip()
        return self._build_commercial_email_html(
            title="Encomenda",
            reference=numero,
            greeting=self._order_recipient_label(payload),
            intro_lines=[
                f"Segue em anexo a nossa encomenda {numero}." if numero else "Segue em anexo a nossa encomenda.",
                "Agradecemos confirmação de receção e do prazo previsto de entrega.",
            ],
            summary_title="Condições da encomenda",
            summary_rows=[
                {"label": "Documento", "value": numero or "-", "emphasis": "strong"},
                {"label": "Valor total", "value": total_txt, "emphasis": "money"},
                {"label": "Entrega pretendida", "value": data_entrega},
                {"label": "Local de descarga", "value": local_descarga},
                {"label": "Transporte", "value": transporte},
            ],
            note_title="Observações",
            note_lines=[obs_txt] if obs_txt else [],
            logo_cid=logo_cid,
        )

    def _note_order_email_dialog(self, detail: dict[str, object]) -> dict[str, object] | None:
        supplier = self._note_order_supplier(detail)
        dialog = QDialog(self)
        dialog.setWindowTitle("Enviar encomenda | Outlook")
        dialog.setMinimumWidth(780)
        dialog.resize(860, 620)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        intro = QLabel("A encomenda vai ser preparada para Outlook com o PDF oficial em anexo e o fornecedor desta NE como destinatário principal.")
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        info_card = CardFrame()
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(12, 10, 12, 10)
        info_layout.setSpacing(4)
        info_layout.addWidget(QLabel(f"Fornecedor: {supplier.get('nome', '') or '-'}"))
        info_layout.addWidget(QLabel(f"Email fornecedor: {supplier.get('email', '') or 'Sem email na ficha'}"))
        layout.addWidget(info_card)

        to_edit = QLineEdit(str(supplier.get("email", "") or "").strip())
        cc_edit = QLineEdit(self._default_outbound_email())
        bcc_edit = QLineEdit("")
        subject_edit = QLineEdit(self._note_order_email_subject(detail))
        body_edit = QTextEdit(self._note_order_email_body(detail))
        body_edit.setMinimumHeight(240)
        send_now_check = QCheckBox("Enviar diretamente sem abrir o rascunho no Outlook")

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(10)
        form.addRow("Para", to_edit)
        form.addRow("CC", cc_edit)
        form.addRow("BCC", bcc_edit)
        form.addRow("Assunto", subject_edit)
        form.addRow("Mensagem", body_edit)
        layout.addLayout(form)
        layout.addWidget(send_now_check)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn is not None:
            ok_btn.setText("Abrir Outlook")
        if cancel_btn is not None:
            cancel_btn.setText("Cancelar")

        def _sync_ok_text() -> None:
            if ok_btn is not None:
                ok_btn.setText("Enviar" if send_now_check.isChecked() else "Abrir Outlook")

        def _accept() -> None:
            if not self._split_email_tokens(to_edit.text()):
                QMessageBox.warning(dialog, "Enviar encomenda", "O fornecedor da encomenda precisa de ter um email válido.")
                return
            if not subject_edit.text().strip():
                QMessageBox.warning(dialog, "Enviar encomenda", "Indica um assunto para o email.")
                return
            if not body_edit.toPlainText().strip():
                QMessageBox.warning(dialog, "Enviar encomenda", "Indica a mensagem do email.")
                return
            dialog.accept()

        send_now_check.toggled.connect(_sync_ok_text)
        _sync_ok_text()
        buttons.accepted.connect(_accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "to": "; ".join(self._split_email_tokens(to_edit.text())),
            "cc": "; ".join(self._split_email_tokens(cc_edit.text())),
            "bcc": "; ".join(self._split_email_tokens(bcc_edit.text())),
            "subject": subject_edit.text().strip(),
            "body_plain": body_edit.toPlainText().strip(),
            "send_now": bool(send_now_check.isChecked()),
        }

    def _open_note_order_email_draft(self, detail: dict[str, object], mail_payload: dict[str, object]) -> None:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        if not numero:
            raise ValueError("Guarda primeiro a nota antes de preparar o email.")
        safe_number = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in numero)[:48] or "encomenda"
        attachment_name = f"Encomenda_{safe_number}.pdf"
        attachment_path: Path | None = Path(tempfile.gettempdir()) / attachment_name
        attachment_issue = ""
        try:
            self.backend.ne_render_pdf(numero, quote=False, output_path=attachment_path)
        except Exception as exc:
            attachment_issue = str(exc)
            attachment_path = None

        logo_path = getattr(self.backend, "logo_path", None)
        logo_file = Path(logo_path) if isinstance(logo_path, Path) and logo_path.exists() else None
        logo_cid = "lugest-ne-order-logo" if logo_file is not None else ""
        subject = str(mail_payload.get("subject", "") or "").strip()
        body_plain = str(mail_payload.get("body_plain", "") or "").strip()
        body_html = self._note_order_email_html_body(payload, body_plain=body_plain, logo_cid=logo_cid)
        send_now = bool(mail_payload.get("send_now"))

        env = os.environ.copy()
        env["LUGEST_MAIL_TO"] = str(mail_payload.get("to", "") or "").strip()
        env["LUGEST_MAIL_CC"] = str(mail_payload.get("cc", "") or "").strip()
        env["LUGEST_MAIL_BCC"] = str(mail_payload.get("bcc", "") or "").strip()
        env["LUGEST_MAIL_SUBJECT"] = subject
        env["LUGEST_MAIL_BODY"] = body_html
        env["LUGEST_MAIL_ATTACHMENT"] = str(attachment_path) if attachment_path is not None else ""
        env["LUGEST_MAIL_LOGO"] = str(logo_file) if logo_file is not None else ""
        env["LUGEST_MAIL_LOGO_CID"] = logo_cid
        env["LUGEST_MAIL_SEND_NOW"] = "1" if send_now else "0"
        powershell_script = (
            "$ErrorActionPreference='Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            "$mail.To = $env:LUGEST_MAIL_TO; "
            "if ($env:LUGEST_MAIL_CC) { $mail.CC = $env:LUGEST_MAIL_CC }; "
            "if ($env:LUGEST_MAIL_BCC) { $mail.BCC = $env:LUGEST_MAIL_BCC }; "
            "$mail.Subject = $env:LUGEST_MAIL_SUBJECT; "
            "if ($env:LUGEST_MAIL_LOGO -and (Test-Path $env:LUGEST_MAIL_LOGO)) "
            "{ "
            "  $logo = $mail.Attachments.Add($env:LUGEST_MAIL_LOGO); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', $env:LUGEST_MAIL_LOGO_CID); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true) "
            "}; "
            "if ($env:LUGEST_MAIL_ATTACHMENT -and (Test-Path $env:LUGEST_MAIL_ATTACHMENT)) "
            "{ $null = $mail.Attachments.Add($env:LUGEST_MAIL_ATTACHMENT) }; "
            "$mail.HTMLBody = $env:LUGEST_MAIL_BODY; "
            "if ($env:LUGEST_MAIL_SEND_NOW -eq '1') { $mail.Send() } else { $mail.Display() }"
        )
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    powershell_script,
                ],
                check=True,
                env=env,
                timeout=30,
            )
        except Exception:
            mailto = (
                f"mailto:{quote(str(mail_payload.get('to', '') or '').strip())}"
                f"?cc={quote(str(mail_payload.get('cc', '') or '').strip())}"
                f"&bcc={quote(str(mail_payload.get('bcc', '') or '').strip())}"
                f"&subject={quote(subject)}"
                f"&body={quote(body_plain)}"
            )
            os.startfile(mailto)
            fallback_message = "Outlook indisponível. Foi aberto o cliente de email por defeito."
            if attachment_issue:
                fallback_message += f"\n\nTambém não foi possível gerar o PDF em anexo:\n{attachment_issue}"
            else:
                fallback_message += "\n\nNota: o anexo PDF terá de ser adicionado manualmente neste modo."
            QMessageBox.information(self, "Enviar encomenda", fallback_message)
            return

        if attachment_issue:
            QMessageBox.information(
                self,
                "Enviar encomenda",
                f"O email foi preparado no Outlook, mas o PDF não foi anexado automaticamente:\n{attachment_issue}",
            )
        elif send_now:
            QMessageBox.information(self, "Enviar encomenda", "Encomenda enviada com sucesso pelo Outlook.")

    def _request_quote_email(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            detail = self.backend.ne_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        if not list(detail.get("lines", []) or []):
            QMessageBox.warning(self, "Notas Encomenda", "A nota não tem linhas para pedir cotação.")
            return
        mail_payload = self._note_quote_email_dialog(detail)
        if mail_payload is None:
            return
        try:
            self._open_note_quote_email_draft(detail, mail_payload)
        except Exception as exc:
            QMessageBox.critical(self, "Pedido de cotação", str(exc))

    def _send_order_email(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            detail = self.backend.ne_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Enviar encomenda", str(exc))
            return
        kind = str(detail.get("kind", "") or "").strip()
        estado_txt = str(detail.get("estado", "") or "").strip().lower()
        if kind == "rfq":
            QMessageBox.warning(self, "Enviar encomenda", "Este botão é para NEs aprovadas. Usa 'Pedir orçamento' nas notas de cotação.")
            return
        if "aprov" not in estado_txt and "enviad" not in estado_txt:
            QMessageBox.warning(self, "Enviar encomenda", "Aprova primeiro a encomenda antes de a enviar ao fornecedor.")
            return
        supplier = self._note_order_supplier(detail)
        if not str(supplier.get("email", "") or "").strip():
            QMessageBox.warning(self, "Enviar encomenda", "O fornecedor desta encomenda não tem email na ficha.")
            return
        mail_payload = self._note_order_email_dialog(detail)
        if mail_payload is None:
            return
        try:
            self._open_note_order_email_draft(detail, mail_payload)
            self.backend.ne_mark_sent(self.current_number)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Enviar encomenda", str(exc))
