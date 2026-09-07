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
from lugest_modules.clients.application.actions import ClientActions

class ClientPage(QWidget):
    page_title = "Clientes"
    page_subtitle = "Cadastro comercial de clientes ligado ao backend atual."
    uses_backend_reload = True

    def __init__(self, actions: ClientActions, parent=None) -> None:
        super().__init__(parent)
        self.actions = actions
        self.rows: list[dict] = []
        self.current_code = ""
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
        hero_title = QLabel("Carteira de clientes")
        hero_title.setStyleSheet("font-family: 'Segoe UI'; font-size: 15px; font-weight: 800; color: #0f172a;")
        hero_subtitle = QLabel("Pesquisa, cria e atualiza contactos comerciais sem sair do painel.")
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
        self.filter_edit.setPlaceholderText("Pesquisar codigo, nome, nif, contacto...")
        self.filter_edit.setProperty("compact", "true")
        self.filter_edit.textChanged.connect(lambda _text: self._filter_timer.start(180))
        self.new_btn = QPushButton("Novo cliente")
        self.new_btn.clicked.connect(self._new_client)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.setProperty("variant", "success")
        self.save_btn.clicked.connect(self._save_client)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "destructive")
        self.remove_btn.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.remove_btn.clicked.connect(self._remove_client)
        for button in (self.new_btn, self.save_btn, self.remove_btn):
            button.setProperty("compact", "true")
        hero_row.addLayout(hero_text, 1)
        hero_row.addWidget(self.new_btn)
        hero_row.addWidget(self.save_btn)
        hero_row.addWidget(self.remove_btn)
        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(8)
        self.client_count_chip = _metric_chip("Clientes", "0", "info")
        self.client_contact_chip = _metric_chip("Com contacto", "0", "success")
        self.client_terms_chip = _metric_chip("Condicoes", "0", "warning")
        metrics_row.addWidget(self.client_count_chip)
        metrics_row.addWidget(self.client_contact_chip)
        metrics_row.addWidget(self.client_terms_chip)
        metrics_row.addStretch(1)
        command_row = QHBoxLayout()
        command_row.setContentsMargins(0, 0, 0, 0)
        command_row.setSpacing(7)
        command_row.addWidget(_search_box(self.filter_edit, "ClientSearchBox"), 1)
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
        table_title = QLabel("Base de clientes")
        table_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        table_subtitle = QLabel("Pesquisa rapida e selecao direta da ficha comercial.")
        table_subtitle.setProperty("role", "muted")
        table_subtitle.setWordWrap(True)
        table_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Codigo", "Nome", "NIF", "Contacto", "Email"])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.table, stretch=(1, 4), contents=(0, 2, 3))
        _set_table_columns(
            self.table,
            [
                (0, "interactive", 100),
                (1, "stretch", 280),
                (2, "interactive", 118),
                (3, "interactive", 130),
                (4, "stretch", 260),
            ],
        )
        self.table.itemSelectionChanged.connect(self._load_selected_client)
        table_layout.addWidget(table_title)
        table_layout.addWidget(table_subtitle)
        table_layout.addWidget(self.table)
        split.addWidget(table_card)

        form_card = CardFrame(object_name="ClientDetailsCard")
        form_card.set_tone("default")
        form_card.setMinimumWidth(410)
        form_card.setMaximumWidth(560)
        form_card.setStyleSheet(
            "QFrame#ClientDetailsCard { background: #f7f9f6; border: 1px solid #cfd8cc; border-radius: 9px; }"
            "QFrame#ClientDetailsCard QLineEdit, QFrame#ClientDetailsCard QComboBox, QFrame#ClientDetailsCard QTextEdit { background: #ffffff; color: #29372e; border: 1px solid #c3cec0; border-radius: 6px; padding: 5px 7px; font-size: 11px; }"
            "QFrame#ClientDetailsCard QLineEdit:focus, QFrame#ClientDetailsCard QComboBox:focus, QFrame#ClientDetailsCard QTextEdit:focus { border-color: #6f9f45; background: #fcfefb; }"
            "QFrame#ClientDetailsCard QLabel { color: #445148; }"
        )
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(15, 14, 15, 15)
        form_layout.setSpacing(10)
        form_title = QLabel("Ficha do cliente")
        form_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #263d2b;")
        form_subtitle = QLabel("Dados comerciais, contactos e condicoes para documentos.")
        form_subtitle.setProperty("role", "muted")
        form_subtitle.setWordWrap(True)
        form_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        form_subtitle.setStyleSheet("font-size: 10.5px; color: #68736b;")
        self.client_code_edit = QLineEdit()
        self.client_name_edit = QLineEdit()
        self.client_nif_edit = QLineEdit()
        self.client_contact_edit = QLineEdit()
        self.client_email_edit = QLineEdit()
        self.client_address_edit = QTextEdit()
        self.client_address_edit.setMinimumHeight(105)
        self.client_address_edit.setMaximumHeight(150)
        self.client_terms_edit = QComboBox()
        self.client_latitude_edit = QLineEdit()
        self.client_latitude_edit.setPlaceholderText("ex.: 41.1579")
        self.client_longitude_edit = QLineEdit()
        self.client_longitude_edit.setPlaceholderText("ex.: -8.6291")
        self.client_map_btn = QPushButton("Abrir no Google Maps")
        self.client_map_btn.setProperty("variant", "secondary")
        self.client_map_btn.setProperty("compact", "true")
        self.client_map_btn.clicked.connect(
            lambda: _open_google_maps(
                self,
                self.client_address_edit.toPlainText(),
                self.client_latitude_edit.text(),
                self.client_longitude_edit.text(),
            )
        )
        self.client_terms_edit.setEditable(True)
        self.client_terms_edit.setInsertPolicy(QComboBox.NoInsert)
        self.client_terms_edit.addItems(PAYMENT_TERMS_OPTIONS)
        self.client_lead_edit = QLineEdit()
        self.client_notes_edit = QTextEdit()
        self.client_notes_edit.setMinimumHeight(125)
        self.client_notes_edit.setMaximumHeight(190)
        _prepare_partner_fields(
            self.client_code_edit,
            self.client_name_edit,
            self.client_nif_edit,
            self.client_contact_edit,
            self.client_email_edit,
            self.client_address_edit,
            self.client_latitude_edit,
            self.client_longitude_edit,
            self.client_terms_edit,
            self.client_lead_edit,
            self.client_notes_edit,
        )
        form_grid = QGridLayout()
        form_grid.setContentsMargins(0, 0, 0, 0)
        form_grid.setHorizontalSpacing(10)
        form_grid.setVerticalSpacing(10)
        ident_card, ident_form = _section_card("Identificação", "Código interno e dados fiscais.", "default", 190)
        contact_card, contact_form = _section_card("Contacto e localização", "Morada, coordenadas e acesso direto ao mapa.", "default", 390)
        terms_card, terms_form = _section_card("Condições comerciais", "Prazos e notas usadas nos documentos.", "default", 300)
        for label, widget in (
            ("Código", self.client_code_edit),
            ("Nome", self.client_name_edit),
            ("NIF", self.client_nif_edit),
        ):
            ident_form.addRow(label, widget)
        for label, widget in (
            ("Contacto", self.client_contact_edit),
            ("Email", self.client_email_edit),
            ("Morada", self.client_address_edit),
            ("Latitude", self.client_latitude_edit),
            ("Longitude", self.client_longitude_edit),
        ):
            contact_form.addRow(label, widget)
        contact_form.addRow("", self.client_map_btn)
        for label, widget in (
            ("Prazo entrega", self.client_lead_edit),
            ("Cond. pagamento", self.client_terms_edit),
            ("Observacoes", self.client_notes_edit),
        ):
            terms_form.addRow(label, widget)
        form_layout.addWidget(form_title)
        form_layout.addWidget(form_subtitle)
        self.client_detail_tabs = _partner_detail_tabs(
            (("Identificação", ident_card), ("Contacto", contact_card), ("Comercial", terms_card))
        )
        self.client_detail_tabs.setMinimumHeight(410)
        form_layout.addWidget(self.client_detail_tabs, 1)
        split.addWidget(form_card)
        split.setHandleWidth(7)
        split.setSizes([960, 500])
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 0)
        root.addWidget(split, 1)
        self._new_client()

    def refresh(self) -> None:
        previous = self.current_code
        self.rows = self.actions.rows(self.filter_edit.text().strip())
        self.client_count_chip.setText(f"Clientes: {len(self.rows)}")
        self.client_contact_chip.setText(f"Com contacto: {sum(1 for r in self.rows if str(r.get('contacto', '') or '').strip())}")
        self.client_terms_chip.setText(f"Condicoes: {sum(1 for r in self.rows if str(r.get('cond_pagamento', '') or '').strip())}")
        _fill_table(
            self.table,
            [[r.get("codigo", "-"), r.get("nome", "-"), r.get("nif", "-"), r.get("contacto", "-"), r.get("email", "-")] for r in self.rows],
            align_center_from=2,
        )
        for idx, row in enumerate(self.rows):
            item = self.table.item(idx, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("codigo", "") or "").strip())
        if not self.rows:
            self._new_client()
            return
        target = 0
        if previous:
            for idx, row in enumerate(self.rows):
                if str(row.get("codigo", "") or "").strip() == previous:
                    target = idx
                    break
        self.table.selectRow(target)
        self._load_selected_client()

    def _selected_client(self) -> dict:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        item = self.table.item(row_index, 0)
        code = str(item.data(Qt.UserRole) or item.text() or "").strip()
        return next((row for row in self.rows if str(row.get("codigo", "") or "").strip() == code), {})

    def _new_client(self) -> None:
        self.current_code = ""
        self.client_code_edit.setText(self.actions.next_code())
        self.client_name_edit.clear()
        self.client_nif_edit.clear()
        self.client_contact_edit.clear()
        self.client_email_edit.clear()
        self.client_address_edit.clear()
        self.client_latitude_edit.clear()
        self.client_longitude_edit.clear()
        self.client_terms_edit.setCurrentText("")
        self.client_lead_edit.clear()
        self.client_notes_edit.clear()

    def _load_selected_client(self) -> None:
        row = self._selected_client()
        if not row:
            return
        self.current_code = str(row.get("codigo", "") or "").strip()
        self.client_code_edit.setText(self.current_code)
        self.client_name_edit.setText(str(row.get("nome", "") or "").strip())
        self.client_nif_edit.setText(str(row.get("nif", "") or "").strip())
        self.client_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        self.client_email_edit.setText(str(row.get("email", "") or "").strip())
        self.client_address_edit.setPlainText(str(row.get("morada", "") or "").strip())
        self.client_latitude_edit.setText(str(row.get("latitude", "") or "").strip())
        self.client_longitude_edit.setText(str(row.get("longitude", "") or "").strip())
        self.client_lead_edit.setText(str(row.get("prazo_entrega", "") or "").strip())
        self.client_terms_edit.setCurrentText(str(row.get("cond_pagamento", "") or "").strip())
        self.client_notes_edit.setPlainText(str(row.get("observacoes", "") or "").strip())

    def _save_client(self) -> None:
        try:
            saved = self.actions.save(
                {
                    "codigo": self.client_code_edit.text().strip(),
                    "nome": self.client_name_edit.text().strip(),
                    "nif": self.client_nif_edit.text().strip(),
                    "contacto": self.client_contact_edit.text().strip(),
                    "email": self.client_email_edit.text().strip(),
                    "morada": self.client_address_edit.toPlainText().strip(),
                    "latitude": self.client_latitude_edit.text().strip(),
                    "longitude": self.client_longitude_edit.text().strip(),
                    "prazo_entrega": self.client_lead_edit.text().strip(),
                    "cond_pagamento": self.client_terms_edit.currentText().strip(),
                    "observacoes": self.client_notes_edit.toPlainText().strip(),
                }
            )
        except Exception as exc:
            QMessageBox.critical(self, "Clientes", str(exc))
            return
        self.current_code = str(saved.get("codigo", "") or "").strip()
        self.refresh()

    def _remove_client(self) -> None:
        code = self.client_code_edit.text().strip()
        if not code:
            return
        if QMessageBox.question(self, "Clientes", f"Remover cliente {code}?") != QMessageBox.Yes:
            return
        try:
            self.actions.remove(code)
        except Exception as exc:
            QMessageBox.critical(self, "Clientes", str(exc))
            return
        self._new_client()
        self.refresh()
