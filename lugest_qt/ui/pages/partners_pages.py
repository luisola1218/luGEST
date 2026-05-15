from __future__ import annotations

from PySide6.QtCore import Qt
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
    QTableWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame
from .runtime_common import (
    configure_table as _configure_table,
    fill_table as _fill_table,
    selected_row_index as _selected_row_index,
    set_table_columns as _set_table_columns,
)


PAYMENT_TERMS_OPTIONS = ["", "Pronto Pagamento", "30 dias", "60 dias", "90 dias", "180 dias"]


def _section_card(title: str, subtitle: str = "", tone: str = "default", minimum_height: int = 0) -> tuple[CardFrame, QFormLayout]:
    card = CardFrame()
    card.set_tone(tone)
    if minimum_height:
        card.setMinimumHeight(minimum_height)
    layout = QVBoxLayout(card)
    layout.setContentsMargins(14, 12, 14, 12)
    layout.setSpacing(8)
    title_label = QLabel(title)
    title_label.setStyleSheet("font-size: 13px; font-weight: 900; color: #0f172a;")
    layout.addWidget(title_label)
    if subtitle:
        subtitle_label = QLabel(subtitle)
        subtitle_label.setProperty("role", "muted")
        subtitle_label.setWordWrap(True)
        layout.addWidget(subtitle_label)
    form = QFormLayout()
    form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    form.setHorizontalSpacing(14)
    form.setVerticalSpacing(8)
    form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
    layout.addLayout(form)
    return card, form


def _scrollable_form_area(content_layout: QGridLayout) -> QScrollArea:
    content = QWidget()
    content.setLayout(content_layout)
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QScrollArea.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    scroll.setWidget(content)
    return scroll


def _prepare_partner_fields(*widgets: QWidget) -> None:
    for widget in widgets:
        if isinstance(widget, (QLineEdit, QComboBox)):
            widget.setMinimumHeight(34)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        elif isinstance(widget, QTextEdit):
            widget.setMinimumHeight(72)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)


def _metric_chip(title: str, value: str = "-", tone: str = "default") -> QLabel:
    colors = {
        "default": ("#eef4fb", "#3b5877"),
        "info": ("#eaf3ff", "#17426b"),
        "success": ("#ecfdf3", "#14532d"),
        "warning": ("#fff7df", "#8a5b00"),
    }
    bg, fg = colors.get(tone, colors["default"])
    label = QLabel(f"{title}: {value}")
    label.setStyleSheet(
        f"background: {bg}; color: {fg}; border: 1px solid rgba(107, 143, 179, 0.35); "
        "border-radius: 12px; padding: 6px 10px; font-weight: 800;"
    )
    return label


class ClientsPage(QWidget):
    page_title = "Clientes"
    page_subtitle = "Cadastro comercial de clientes ligado ao backend atual."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_code = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        top = CardFrame()
        top.set_tone("info")
        top_layout = QVBoxLayout(top)
        top_layout.setContentsMargins(16, 14, 16, 14)
        top_layout.setSpacing(10)
        hero_row = QHBoxLayout()
        hero_text = QVBoxLayout()
        hero_text.setSpacing(2)
        hero_title = QLabel("Carteira de clientes")
        hero_title.setStyleSheet("font-size: 20px; font-weight: 950; color: #0f172a;")
        hero_subtitle = QLabel("Pesquisa, cria e atualiza contactos comerciais sem sair do painel.")
        hero_subtitle.setProperty("role", "muted")
        hero_text.addWidget(hero_title)
        hero_text.addWidget(hero_subtitle)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar codigo, nome, nif, contacto...")
        self.filter_edit.setProperty("compact", "true")
        self.filter_edit.textChanged.connect(self.refresh)
        self.new_btn = QPushButton("Novo cliente")
        self.new_btn.clicked.connect(self._new_client)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_client)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
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
        top_layout.addLayout(hero_row)
        top_layout.addWidget(self.filter_edit)
        top_layout.addLayout(metrics_row)
        root.addWidget(top)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_title = QLabel("Base de clientes")
        table_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        table_subtitle = QLabel("Pesquisa rapida e selecao direta da ficha comercial.")
        table_subtitle.setProperty("role", "muted")
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

        form_card = CardFrame()
        form_card.set_tone("info")
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)
        form_title = QLabel("Ficha do cliente")
        form_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        form_subtitle = QLabel("Dados comerciais, contactos e condicoes para documentos.")
        form_subtitle.setProperty("role", "muted")
        self.client_code_edit = QLineEdit()
        self.client_name_edit = QLineEdit()
        self.client_nif_edit = QLineEdit()
        self.client_contact_edit = QLineEdit()
        self.client_email_edit = QLineEdit()
        self.client_address_edit = QTextEdit()
        self.client_address_edit.setMinimumHeight(70)
        self.client_address_edit.setMaximumHeight(92)
        self.client_terms_edit = QComboBox()
        self.client_terms_edit.setEditable(True)
        self.client_terms_edit.setInsertPolicy(QComboBox.NoInsert)
        self.client_terms_edit.addItems(PAYMENT_TERMS_OPTIONS)
        self.client_lead_edit = QLineEdit()
        self.client_notes_edit = QTextEdit()
        self.client_notes_edit.setMinimumHeight(78)
        self.client_notes_edit.setMaximumHeight(118)
        _prepare_partner_fields(
            self.client_code_edit,
            self.client_name_edit,
            self.client_nif_edit,
            self.client_contact_edit,
            self.client_email_edit,
            self.client_address_edit,
            self.client_terms_edit,
            self.client_lead_edit,
            self.client_notes_edit,
        )
        form_grid = QGridLayout()
        form_grid.setContentsMargins(0, 0, 0, 0)
        form_grid.setHorizontalSpacing(10)
        form_grid.setVerticalSpacing(10)
        ident_card, ident_form = _section_card("Identificacao", "Codigo interno e dados fiscais.", "default", 148)
        contact_card, contact_form = _section_card("Contacto", "Morada e meios de contacto.", "default", 206)
        terms_card, terms_form = _section_card("Condicoes comerciais", "Prazos e notas usadas nos documentos.", "warning", 220)
        for label, widget in (
            ("Codigo", self.client_code_edit),
            ("Nome", self.client_name_edit),
            ("NIF", self.client_nif_edit),
        ):
            ident_form.addRow(label, widget)
        for label, widget in (
            ("Contacto", self.client_contact_edit),
            ("Email", self.client_email_edit),
            ("Morada", self.client_address_edit),
        ):
            contact_form.addRow(label, widget)
        for label, widget in (
            ("Prazo entrega", self.client_lead_edit),
            ("Cond. pagamento", self.client_terms_edit),
            ("Observacoes", self.client_notes_edit),
        ):
            terms_form.addRow(label, widget)
        form_grid.addWidget(ident_card, 0, 0)
        form_grid.addWidget(contact_card, 1, 0)
        form_grid.addWidget(terms_card, 2, 0)
        form_grid.setRowStretch(3, 1)
        form_layout.addWidget(form_title)
        form_layout.addWidget(form_subtitle)
        form_layout.addWidget(_scrollable_form_area(form_grid), 1)
        split.addWidget(form_card)
        split.setSizes([760, 760])
        root.addWidget(split, 1)
        self._new_client()

    def refresh(self) -> None:
        previous = self.current_code
        self.rows = self.backend.client_rows(self.filter_edit.text().strip())
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
        self.client_code_edit.setText(self.backend.client_next_code())
        self.client_name_edit.clear()
        self.client_nif_edit.clear()
        self.client_contact_edit.clear()
        self.client_email_edit.clear()
        self.client_address_edit.clear()
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
        self.client_lead_edit.setText(str(row.get("prazo_entrega", "") or "").strip())
        self.client_terms_edit.setCurrentText(str(row.get("cond_pagamento", "") or "").strip())
        self.client_notes_edit.setPlainText(str(row.get("observacoes", "") or "").strip())

    def _save_client(self) -> None:
        try:
            saved = self.backend.client_save(
                {
                    "codigo": self.client_code_edit.text().strip(),
                    "nome": self.client_name_edit.text().strip(),
                    "nif": self.client_nif_edit.text().strip(),
                    "contacto": self.client_contact_edit.text().strip(),
                    "email": self.client_email_edit.text().strip(),
                    "morada": self.client_address_edit.toPlainText().strip(),
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
            self.backend.client_remove(code)
        except Exception as exc:
            QMessageBox.critical(self, "Clientes", str(exc))
            return
        self._new_client()
        self.refresh()


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
        top_layout.setContentsMargins(16, 14, 16, 14)
        top_layout.setSpacing(10)
        hero_row = QHBoxLayout()
        hero_text = QVBoxLayout()
        hero_text.setSpacing(2)
        hero_title = QLabel("Rede de fornecedores")
        hero_title.setStyleSheet("font-size: 20px; font-weight: 950; color: #0f172a;")
        hero_subtitle = QLabel("Controla contactos, prazos e condicoes de compra com leitura rapida.")
        hero_subtitle.setProperty("role", "muted")
        hero_text.addWidget(hero_title)
        hero_text.addWidget(hero_subtitle)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar fornecedor, nif, contacto...")
        self.filter_edit.setProperty("compact", "true")
        self.filter_edit.textChanged.connect(self.refresh)
        self.new_btn = QPushButton("Novo fornecedor")
        self.new_btn.clicked.connect(self._new_supplier)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_supplier)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
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
        top_layout.addLayout(hero_row)
        top_layout.addWidget(self.filter_edit)
        top_layout.addLayout(metrics_row)
        root.addWidget(top)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_title = QLabel("Fornecedores")
        table_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        table_subtitle = QLabel("Base de compras, contactos e condicoes de fornecimento.")
        table_subtitle.setProperty("role", "muted")
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
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)
        form_title = QLabel("Ficha do fornecedor")
        form_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        form_subtitle = QLabel("Informacao comercial usada nas compras e notas de encomenda.")
        form_subtitle.setProperty("role", "muted")
        self.supplier_id_edit = QLineEdit()
        self.supplier_name_edit = QLineEdit()
        self.supplier_nif_edit = QLineEdit()
        self.supplier_contact_edit = QLineEdit()
        self.supplier_email_edit = QLineEdit()
        self.supplier_address_edit = QTextEdit()
        self.supplier_address_edit.setMinimumHeight(62)
        self.supplier_address_edit.setMaximumHeight(84)
        self.supplier_terms_edit = QComboBox()
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
        contact_card, contact_form = _section_card("Contacto", "Morada, email e website.", "default", 232)
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
            ("Website", self.supplier_website_edit),
        ):
            contact_form.addRow(label, widget)
        for label, widget in (
            ("Cond. pagamento", self.supplier_terms_edit),
            ("Prazo entrega (dias)", self.supplier_lead_days_edit),
            ("Observacoes", self.supplier_notes_edit),
        ):
            terms_form.addRow(label, widget)
        form_grid.addWidget(ident_card, 0, 0)
        form_grid.addWidget(contact_card, 1, 0)
        form_grid.addWidget(terms_card, 2, 0)
        form_grid.setRowStretch(3, 1)
        form_layout.addWidget(form_title)
        form_layout.addWidget(form_subtitle)
        form_layout.addWidget(_scrollable_form_area(form_grid), 1)
        split.addWidget(form_card)
        split.setSizes([760, 760])
        root.addWidget(split, 1)
        self._new_supplier()

    def can_auto_refresh(self) -> bool:
        return False

    def refresh(self) -> None:
        previous = self.current_id
        self.rows = self.backend.ne_suppliers()
        query = self.filter_edit.text().strip().lower()
        if query:
            self.rows = [row for row in self.rows if any(query in str(value).lower() for value in row.values())]
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
