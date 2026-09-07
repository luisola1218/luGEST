from __future__ import annotations

import re
import unicodedata
from urllib.parse import quote_plus

from PySide6.QtCore import QUrl, Qt, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QScrollArea,
    QSplitter,
    QStyle,
    QTableWidget,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from lugest_qt.ui.widgets import CardFrame
from lugest_qt.ui.pages.runtime_common import (
    configure_table as _configure_table,
    fill_table as _fill_table,
    selected_row_index as _selected_row_index,
    set_table_columns as _set_table_columns,
)


from lugest_qt.ui.partner_components import PAYMENT_TERMS_OPTIONS, _open_google_maps, _section_card, _scrollable_form_area, _prepare_partner_fields, _search_box, _search_terms, _row_matches_terms, _metric_chip, _partner_detail_tabs
from lugest_modules.clients.presentation.page import ClientPage
from lugest_modules.clients.application.actions import ClientActions

class ClientsPage(ClientPage):
    """Compatibility constructor for the desktop page registry."""
    def __init__(self, backend, parent=None):
        super().__init__(ClientActions(
            rows=lambda query: backend.client_rows(query),
            next_code=lambda: backend.client_next_code(),
            save=lambda payload: backend.client_save(payload),
            remove=lambda code: backend.client_remove(code),
        ), parent)

class SuppliersPage(QWidget):
    page_title = "Fornecedores"
    page_subtitle = "Cadastro de fornecedores para notas de encomenda e compras."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_id = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        top = CardFrame()
        top.set_tone("info")
        top_layout = QVBoxLayout(top)
        top_layout.setContentsMargins(12, 9, 12, 9)
        top_layout.setSpacing(7)
        hero_row = QHBoxLayout()
        hero_text = QVBoxLayout()
        hero_text.setSpacing(2)
        hero_title = QLabel("Rede de fornecedores")
        hero_title.setStyleSheet("font-family: 'Segoe UI'; font-size: 15px; font-weight: 800; color: #0f172a;")
        hero_subtitle = QLabel("Controla contactos, prazos e condicoes de compra com leitura rapida.")
        hero_subtitle.setProperty("role", "muted")
        hero_subtitle.setWordWrap(True)
        hero_subtitle.setMinimumWidth(0)
        hero_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        hero_text.addWidget(hero_title)
        hero_text.addWidget(hero_subtitle)
        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.timeout.connect(self.refresh)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar fornecedor, nif, contacto...")
        self.filter_edit.setProperty("compact", "true")
        self.filter_edit.textChanged.connect(lambda _text: self._filter_timer.start(180))
        self.new_btn = QPushButton("Novo fornecedor")
        self.new_btn.clicked.connect(self._new_supplier)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.setProperty("variant", "success")
        self.save_btn.clicked.connect(self._save_supplier)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "destructive")
        self.remove_btn.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.remove_btn.clicked.connect(self._remove_supplier)
        for button in (self.new_btn, self.save_btn, self.remove_btn):
            button.setProperty("compact", "true")
        hero_row.addLayout(hero_text, 1)
        hero_row.addWidget(self.new_btn)
        hero_row.addWidget(self.save_btn)
        hero_row.addWidget(self.remove_btn)
        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(8)
        self.supplier_count_chip = _metric_chip("Fornecedores", "0", "info")
        self.supplier_contact_chip = _metric_chip("Com contacto", "0", "success")
        self.supplier_terms_chip = _metric_chip("Condicoes", "0", "warning")
        metrics_row.addWidget(self.supplier_count_chip)
        metrics_row.addWidget(self.supplier_contact_chip)
        metrics_row.addWidget(self.supplier_terms_chip)
        metrics_row.addStretch(1)
        command_row = QHBoxLayout()
        command_row.setContentsMargins(0, 0, 0, 0)
        command_row.setSpacing(7)
        command_row.addWidget(_search_box(self.filter_edit, "SupplierSearchBox"), 1)
        command_row.addLayout(metrics_row)
        top_layout.addLayout(hero_row)
        top_layout.addLayout(command_row)
        top.setMaximumHeight(108)
        root.addWidget(top)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_title = QLabel("Fornecedores")
        table_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        table_subtitle = QLabel("Base de compras, contactos e condicoes de fornecimento.")
        table_subtitle.setProperty("role", "muted")
        table_subtitle.setWordWrap(True)
        table_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "Nome", "NIF", "Contacto", "Email"])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.table, stretch=(1, 3, 4), contents=(2,))
        _set_table_columns(
            self.table,
            [
                (0, "interactive", 96),
                (1, "stretch", 220),
                (2, "interactive", 132),
                (3, "stretch", 156),
                (4, "stretch", 192),
            ],
        )
        self.table.itemSelectionChanged.connect(self._load_selected_supplier)
        table_layout.addWidget(table_title)
        table_layout.addWidget(table_subtitle)
        table_layout.addWidget(self.table)
        split.addWidget(table_card)

        form_card = CardFrame()
        form_card.set_tone("info")
        form_card.setMinimumWidth(390)
        form_card.setMaximumWidth(520)
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)
        form_title = QLabel("Ficha do fornecedor")
        form_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        form_subtitle = QLabel("Informacao comercial usada nas compras e notas de encomenda.")
        form_subtitle.setProperty("role", "muted")
        form_subtitle.setWordWrap(True)
        form_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.supplier_id_edit = QLineEdit()
        self.supplier_name_edit = QLineEdit()
        self.supplier_nif_edit = QLineEdit()
        self.supplier_contact_edit = QLineEdit()
        self.supplier_email_edit = QLineEdit()
        self.supplier_address_edit = QTextEdit()
        self.supplier_address_edit.setMinimumHeight(62)
        self.supplier_address_edit.setMaximumHeight(84)
        self.supplier_terms_edit = QComboBox()
        self.supplier_latitude_edit = QLineEdit()
        self.supplier_latitude_edit.setPlaceholderText("ex.: 41.1579")
        self.supplier_longitude_edit = QLineEdit()
        self.supplier_longitude_edit.setPlaceholderText("ex.: -8.6291")
        self.supplier_map_btn = QPushButton("Abrir no Google Maps")
        self.supplier_map_btn.setProperty("variant", "secondary")
        self.supplier_map_btn.setProperty("compact", "true")
        self.supplier_map_btn.clicked.connect(
            lambda: _open_google_maps(
                self,
                self.supplier_address_edit.toPlainText(),
                self.supplier_latitude_edit.text(),
                self.supplier_longitude_edit.text(),
            )
        )
        self.supplier_terms_edit.setEditable(True)
        self.supplier_terms_edit.setInsertPolicy(QComboBox.NoInsert)
        self.supplier_terms_edit.addItems(PAYMENT_TERMS_OPTIONS)
        self.supplier_lead_days_edit = QLineEdit()
        self.supplier_website_edit = QLineEdit()
        self.supplier_notes_edit = QTextEdit()
        self.supplier_notes_edit.setMinimumHeight(74)
        self.supplier_notes_edit.setMaximumHeight(110)
        _prepare_partner_fields(
            self.supplier_id_edit,
            self.supplier_name_edit,
            self.supplier_nif_edit,
            self.supplier_contact_edit,
            self.supplier_email_edit,
            self.supplier_address_edit,
            self.supplier_latitude_edit,
            self.supplier_longitude_edit,
            self.supplier_terms_edit,
            self.supplier_lead_days_edit,
            self.supplier_website_edit,
            self.supplier_notes_edit,
        )
        form_grid = QGridLayout()
        form_grid.setContentsMargins(0, 0, 0, 0)
        form_grid.setHorizontalSpacing(10)
        form_grid.setVerticalSpacing(10)
        ident_card, ident_form = _section_card("Identificacao", "Referencia interna e dados fiscais.", "default", 148)
        contact_card, contact_form = _section_card("Contacto e localização", "Morada, coordenadas, email e website.", "default", 330)
        terms_card, terms_form = _section_card("Condicoes de compra", "Prazos, pagamento e observacoes.", "warning", 210)
        for label, widget in (
            ("ID", self.supplier_id_edit),
            ("Nome", self.supplier_name_edit),
            ("NIF", self.supplier_nif_edit),
        ):
            ident_form.addRow(label, widget)
        for label, widget in (
            ("Contacto", self.supplier_contact_edit),
            ("Email", self.supplier_email_edit),
            ("Morada", self.supplier_address_edit),
            ("Latitude", self.supplier_latitude_edit),
            ("Longitude", self.supplier_longitude_edit),
            ("Website", self.supplier_website_edit),
        ):
            contact_form.addRow(label, widget)
        contact_form.addRow("", self.supplier_map_btn)
        for label, widget in (
            ("Cond. pagamento", self.supplier_terms_edit),
            ("Prazo entrega (dias)", self.supplier_lead_days_edit),
            ("Observacoes", self.supplier_notes_edit),
        ):
            terms_form.addRow(label, widget)
        form_layout.addWidget(form_title)
        form_layout.addWidget(form_subtitle)
        self.supplier_detail_tabs = _partner_detail_tabs(
            (("Identificação", ident_card), ("Contacto", contact_card), ("Compras", terms_card))
        )
        form_layout.addWidget(self.supplier_detail_tabs, 1)
        split.addWidget(form_card)
        split.setHandleWidth(7)
        split.setSizes([980, 460])
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 0)
        root.addWidget(split, 1)
        self._new_supplier()

    def can_auto_refresh(self) -> bool:
        return False

    def refresh(self) -> None:
        previous = self.current_id
        self.rows = self.backend.ne_suppliers()
        query = self.filter_edit.text().strip()
        if query:
            self.rows = [row for row in self.rows if _row_matches_terms(row, query)]
        self.supplier_count_chip.setText(f"Fornecedores: {len(self.rows)}")
        self.supplier_contact_chip.setText(f"Com contacto: {sum(1 for r in self.rows if str(r.get('contacto', '') or '').strip())}")
        self.supplier_terms_chip.setText(f"Condicoes: {sum(1 for r in self.rows if str(r.get('cond_pagamento', '') or '').strip())}")
        _fill_table(
            self.table,
            [[r.get("id", "-"), r.get("nome", "-"), r.get("nif", "-"), r.get("contacto", "-"), r.get("email", "-")] for r in self.rows],
            align_center_from=2,
        )
        for idx, row in enumerate(self.rows):
            item = self.table.item(idx, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("id", "") or "").strip())
        if not self.rows:
            self._new_supplier()
            return
        target = 0
        if previous:
            for idx, row in enumerate(self.rows):
                if str(row.get("id", "") or "").strip() == previous:
                    target = idx
                    break
        self.table.selectRow(target)
        self._load_selected_supplier()

    def _selected_supplier(self) -> dict:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        item = self.table.item(row_index, 0)
        supplier_id = str(item.data(Qt.UserRole) or item.text() or "").strip()
        return next((row for row in self.rows if str(row.get("id", "") or "").strip() == supplier_id), {})

    def _new_supplier(self) -> None:
        self.current_id = ""
        self.supplier_id_edit.setText(self.backend.supplier_next_id())
        self.supplier_name_edit.clear()
        self.supplier_nif_edit.clear()
        self.supplier_contact_edit.clear()
        self.supplier_email_edit.clear()
        self.supplier_address_edit.clear()
        self.supplier_latitude_edit.clear()
        self.supplier_longitude_edit.clear()
        self.supplier_terms_edit.setCurrentText("")
        self.supplier_lead_days_edit.clear()
        self.supplier_website_edit.clear()
        self.supplier_notes_edit.clear()

    def _load_selected_supplier(self) -> None:
        row = self._selected_supplier()
        if not row:
            return
        self.current_id = str(row.get("id", "") or "").strip()
        self.supplier_id_edit.setText(self.current_id)
        self.supplier_name_edit.setText(str(row.get("nome", "") or "").strip())
        self.supplier_nif_edit.setText(str(row.get("nif", "") or "").strip())
        self.supplier_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        self.supplier_email_edit.setText(str(row.get("email", "") or "").strip())
        self.supplier_address_edit.setPlainText(str(row.get("morada", "") or "").strip())
        self.supplier_latitude_edit.setText(str(row.get("latitude", "") or "").strip())
        self.supplier_longitude_edit.setText(str(row.get("longitude", "") or "").strip())
        self.supplier_terms_edit.setCurrentText(str(row.get("cond_pagamento", "") or "").strip())
        self.supplier_lead_days_edit.setText(str(row.get("prazo_entrega_dias", "") or "").strip())
        self.supplier_website_edit.setText(str(row.get("website", "") or "").strip())
        self.supplier_notes_edit.setPlainText(str(row.get("obs", "") or "").strip())

    def _save_supplier(self) -> None:
        try:
            saved = self.backend.supplier_save(
                {
                    "id": self.supplier_id_edit.text().strip(),
                    "nome": self.supplier_name_edit.text().strip(),
                    "nif": self.supplier_nif_edit.text().strip(),
                    "contacto": self.supplier_contact_edit.text().strip(),
                    "email": self.supplier_email_edit.text().strip(),
                    "morada": self.supplier_address_edit.toPlainText().strip(),
                    "latitude": self.supplier_latitude_edit.text().strip(),
                    "longitude": self.supplier_longitude_edit.text().strip(),
                    "cond_pagamento": self.supplier_terms_edit.currentText().strip(),
                    "prazo_entrega_dias": self.supplier_lead_days_edit.text().strip(),
                    "website": self.supplier_website_edit.text().strip(),
                    "obs": self.supplier_notes_edit.toPlainText().strip(),
                }
            )
        except Exception as exc:
            QMessageBox.critical(self, "Fornecedores", str(exc))
            return
        self.current_id = str(saved.get("id", "") or "").strip()
        self.refresh()

    def _remove_supplier(self) -> None:
        supplier_id = self.supplier_id_edit.text().strip()
        if not supplier_id:
            return
        if QMessageBox.question(self, "Fornecedores", f"Remover fornecedor {supplier_id}?") != QMessageBox.Yes:
            return
        try:
            self.backend.supplier_remove(supplier_id)
        except Exception as exc:
            QMessageBox.critical(self, "Fornecedores", str(exc))
            return
        self._new_supplier()
        self.refresh()
