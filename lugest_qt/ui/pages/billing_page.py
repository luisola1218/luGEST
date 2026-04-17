from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard
from .runtime_pages import (
    _apply_state_chip,
    _cap_width,
    _coerce_editor_qdate,
    _configure_table,
    _fill_table,
    _fmt_eur,
    _paint_table_row,
    _set_panel_tone,
    _state_tone,
    _table_visible_height,
)


class BillingPage(QWidget):
    page_title = "Faturação"
    page_subtitle = "Vendidos, encomendas, faturas e comprovativos num único registo comercial."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_number = ""
        self.current_detail: dict = {}
        self.last_saft_export_path = ""
        self.last_saft_export_record = ""
        self.last_at_batch_path = ""
        self.last_at_batch_record = ""

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        self.view_stack = QStackedWidget()
        root.addWidget(self.view_stack, 1)

        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(14)

        filters = CardFrame()
        filters.setProperty("tone", "info")
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setHorizontalSpacing(8)
        filters_layout.setVerticalSpacing(4)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por registo, orçamento, encomenda, cliente ou fatura")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Por faturar", "Por cobrar", "Pagas", "Atrasadas", "Todas"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.open_btn = QPushButton("Abrir registo")
        self.open_btn.clicked.connect(self._open_selected)
        self.remove_btn = QPushButton("Remover registo")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_selected)
        filters_layout.addWidget(QLabel("Pesquisa"), 0, 0)
        filters_layout.addWidget(QLabel("Estado"), 0, 1)
        filters_layout.addWidget(QLabel("Ano"), 0, 2)
        filters_layout.addWidget(QLabel("Ações"), 0, 3)
        filters_layout.addWidget(self.filter_edit, 1, 0)
        filters_layout.addWidget(self.state_combo, 1, 1)
        filters_layout.addWidget(self.year_combo, 1, 2)
        actions_host = QWidget()
        actions_layout = QHBoxLayout(actions_host)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(6)
        for button, width in ((self.open_btn, 136), (self.remove_btn, 150)):
            button.setProperty("compact", "true")
            button.setMinimumWidth(width)
            actions_layout.addWidget(button)
        actions_layout.addStretch(1)
        filters_layout.addWidget(actions_host, 1, 3)
        filters_layout.setColumnStretch(0, 4)
        filters_layout.setColumnStretch(1, 2)
        filters_layout.setColumnStretch(2, 2)
        filters_layout.setColumnStretch(3, 4)
        filters.setMaximumHeight(86)
        list_layout.addWidget(filters)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)
        self.cards = [StatCard(title) for title in ("Vendido", "Faturado", "Recebido", "Saldo", "Atrasos", "Abertos")]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, index // 3, index % 3)
        list_layout.addWidget(cards_host)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(16, 14, 16, 14)
        table_layout.setSpacing(10)
        table_title = QLabel("Registos de Faturação")
        table_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels(
            ["Registo", "Orçamento", "Encomenda", "Cliente", "Produção", "Expedição", "Faturação", "Pagamento", "Vendido", "Faturado", "Saldo"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        _configure_table(self.table, stretch=(3,), contents=(0, 1, 2, 4, 5, 6, 7, 8, 9, 10))
        header = self.table.horizontalHeader()
        for col, width in ((0, 116), (1, 118), (2, 118), (4, 112), (5, 112), (8, 92), (9, 92), (10, 92)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.itemSelectionChanged.connect(self._sync_list_buttons)
        self.table.itemDoubleClicked.connect(lambda *_args: self._open_selected())
        table_layout.addWidget(table_title)
        table_layout.addWidget(self.table)
        list_layout.addWidget(table_card, 1)

        self.detail_page = QWidget()
        detail_outer = QVBoxLayout(self.detail_page)
        detail_outer.setContentsMargins(0, 0, 0, 0)
        detail_outer.setSpacing(0)
        self.detail_scroll = QScrollArea()
        self.detail_scroll.setWidgetResizable(True)
        detail_outer.addWidget(self.detail_scroll)
        self.detail_host = QWidget()
        self.detail_scroll.setWidget(self.detail_host)
        detail_layout = QVBoxLayout(self.detail_host)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)

        actions_card = CardFrame()
        actions_layout = QHBoxLayout(actions_card)
        actions_layout.setContentsMargins(14, 10, 14, 10)
        actions_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a lista")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_list)
        save_btn = QPushButton("Guardar")
        save_btn.clicked.connect(self._save_record)
        delete_btn = QPushButton("Remover")
        delete_btn.setProperty("variant", "danger")
        delete_btn.clicked.connect(self._remove_current)
        add_invoice_btn = QPushButton("Adicionar fatura")
        add_invoice_btn.clicked.connect(self._add_invoice)
        generate_invoice_btn = QPushButton("Gerar PDF fatura")
        generate_invoice_btn.clicked.connect(self._generate_invoice_pdf)
        edit_invoice_btn = QPushButton("Editar fatura")
        edit_invoice_btn.setProperty("variant", "secondary")
        edit_invoice_btn.clicked.connect(self._edit_invoice)
        add_payment_btn = QPushButton("Adicionar pagamento")
        add_payment_btn.clicked.connect(self._add_payment)
        edit_payment_btn = QPushButton("Editar pagamento")
        edit_payment_btn.setProperty("variant", "secondary")
        edit_payment_btn.clicked.connect(self._edit_payment)
        for button, width in (
            (back_btn, 126),
            (save_btn, 100),
            (delete_btn, 112),
            (add_invoice_btn, 144),
            (generate_invoice_btn, 154),
            (edit_invoice_btn, 126),
            (add_payment_btn, 160),
            (edit_payment_btn, 142),
        ):
            button.setMinimumWidth(width)
            actions_layout.addWidget(button)
        actions_layout.addStretch(1)
        detail_layout.addWidget(actions_card)

        self.header_card = CardFrame()
        header_layout = QVBoxLayout(self.header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)
        header_top = QHBoxLayout()
        title_col = QVBoxLayout()
        title_col.setSpacing(4)
        self.number_label = QLabel("Novo registo")
        self.number_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.source_label = QLabel("Sem origem")
        self.source_label.setProperty("role", "muted")
        title_col.addWidget(self.number_label)
        title_col.addWidget(self.source_label)
        header_top.addLayout(title_col, 1)
        self.invoice_status_chip = QLabel("-")
        self.payment_status_chip = QLabel("-")
        header_top.addWidget(self.invoice_status_chip, 0, Qt.AlignTop)
        header_top.addWidget(self.payment_status_chip, 0, Qt.AlignTop)
        header_layout.addLayout(header_top)
        detail_layout.addWidget(self.header_card)

        info_split = QHBoxLayout()
        info_split.setSpacing(10)

        self.source_card = CardFrame()
        source_layout = QVBoxLayout(self.source_card)
        source_layout.setContentsMargins(12, 10, 12, 10)
        source_layout.setSpacing(8)
        source_title = QLabel("Origem da Venda")
        source_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        source_layout.addWidget(source_title)
        source_form = QFormLayout()
        source_form.setContentsMargins(0, 0, 0, 0)
        source_form.setHorizontalSpacing(10)
        source_form.setVerticalSpacing(6)
        self.quote_label = QLabel("-")
        self.order_label = QLabel("-")
        self.client_label = QLabel("-")
        self.guide_label = QLabel("-")
        source_form.addRow("Orçamento", self.quote_label)
        source_form.addRow("Encomenda", self.order_label)
        source_form.addRow("Cliente", self.client_label)
        source_form.addRow("Última guia", self.guide_label)
        source_layout.addLayout(source_form)

        self.order_state_card = CardFrame()
        order_state_layout = QVBoxLayout(self.order_state_card)
        order_state_layout.setContentsMargins(12, 10, 12, 10)
        order_state_layout.setSpacing(8)
        order_state_title = QLabel("Estado Operacional")
        order_state_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        order_state_layout.addWidget(order_state_title)
        order_state_form = QFormLayout()
        order_state_form.setContentsMargins(0, 0, 0, 0)
        order_state_form.setHorizontalSpacing(10)
        order_state_form.setVerticalSpacing(8)
        self.order_status_chip = QLabel("-")
        self.shipping_status_chip = QLabel("-")
        order_state_form.addRow("Produção", self.order_status_chip)
        order_state_form.addRow("Expedição", self.shipping_status_chip)
        order_state_layout.addLayout(order_state_form)

        self.control_card = CardFrame()
        control_layout = QVBoxLayout(self.control_card)
        control_layout.setContentsMargins(12, 10, 12, 10)
        control_layout.setSpacing(8)
        control_title = QLabel("Cobrança")
        control_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        control_layout.addWidget(control_title)
        self.sale_date_edit = QDateEdit()
        self.sale_date_edit.setCalendarPopup(True)
        self.sale_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.due_date_edit = QDateEdit()
        self.due_date_edit.setCalendarPopup(True)
        self.due_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.manual_value_spin = QDoubleSpinBox()
        self.manual_value_spin.setRange(0.0, 100000000.0)
        self.manual_value_spin.setDecimals(2)
        self.manual_value_spin.setPrefix("EUR ")
        self.payment_override_combo = QComboBox()
        self.payment_override_combo.addItems(["Auto", "Pendente", "Parcial", "Paga", "Atrasada", "Incobrável"])
        control_form = QFormLayout()
        control_form.setContentsMargins(0, 0, 0, 0)
        control_form.setHorizontalSpacing(10)
        control_form.setVerticalSpacing(6)
        control_form.addRow("Data venda", self.sale_date_edit)
        control_form.addRow("Vencimento", self.due_date_edit)
        control_form.addRow("Valor manual", self.manual_value_spin)
        control_form.addRow("Estado cobrança", self.payment_override_combo)
        control_layout.addLayout(control_form)
        self.notes_edit = QTextEdit()
        self.notes_edit.setMinimumHeight(80)
        self.notes_edit.setPlaceholderText("Observações internas da faturação, cobrança e documentação.")
        control_layout.addWidget(self.notes_edit)

        for widget in (self.source_card, self.order_state_card, self.control_card):
            info_split.addWidget(widget, 1)
        detail_layout.addLayout(info_split)

        summary_card = CardFrame()
        summary_layout = QGridLayout(summary_card)
        summary_layout.setContentsMargins(12, 10, 12, 10)
        summary_layout.setHorizontalSpacing(12)
        summary_layout.setVerticalSpacing(6)
        self.summary_labels: dict[str, QLabel] = {}
        for row_index, (key, text) in enumerate(
            (
                ("sold", "Vendido"),
                ("invoiced", "Faturado"),
                ("received", "Recebido"),
                ("balance", "Saldo"),
                ("uninvoiced", "Por faturar"),
                ("invoice_count", "Faturas"),
            )
        ):
            label = QLabel(text)
            value = QLabel("-")
            value.setProperty("role", "field_value_strong" if key == "balance" else "field_value")
            summary_layout.addWidget(label, row_index // 3, (row_index % 3) * 2)
            summary_layout.addWidget(value, row_index // 3, (row_index % 3) * 2 + 1)
            self.summary_labels[key] = value
        detail_layout.addWidget(summary_card)

        identity_split = QHBoxLayout()
        identity_split.setSpacing(10)

        self.customer_card = CardFrame()
        customer_layout = QVBoxLayout(self.customer_card)
        customer_layout.setContentsMargins(12, 10, 12, 10)
        customer_layout.setSpacing(8)
        customer_title = QLabel("Cliente / Faturação")
        customer_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        customer_layout.addWidget(customer_title)
        customer_form = QFormLayout()
        customer_form.setContentsMargins(0, 0, 0, 0)
        customer_form.setHorizontalSpacing(10)
        customer_form.setVerticalSpacing(6)
        self.customer_name_label = QLabel("-")
        self.customer_nif_label = QLabel("-")
        self.customer_contact_label = QLabel("-")
        self.customer_email_label = QLabel("-")
        self.customer_address_label = QLabel("-")
        self.customer_address_label.setWordWrap(True)
        customer_form.addRow("Nome", self.customer_name_label)
        customer_form.addRow("NIF", self.customer_nif_label)
        customer_form.addRow("Contacto", self.customer_contact_label)
        customer_form.addRow("Email", self.customer_email_label)
        customer_form.addRow("Morada", self.customer_address_label)
        customer_layout.addLayout(customer_form)

        self.issuer_card = CardFrame()
        issuer_layout = QVBoxLayout(self.issuer_card)
        issuer_layout.setContentsMargins(12, 10, 12, 10)
        issuer_layout.setSpacing(8)
        issuer_title = QLabel("Emitente")
        issuer_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        issuer_layout.addWidget(issuer_title)
        issuer_form = QFormLayout()
        issuer_form.setContentsMargins(0, 0, 0, 0)
        issuer_form.setHorizontalSpacing(10)
        issuer_form.setVerticalSpacing(6)
        self.issuer_name_label = QLabel("-")
        self.issuer_nif_label = QLabel("-")
        self.issuer_address_label = QLabel("-")
        self.issuer_address_label.setWordWrap(True)
        issuer_form.addRow("Empresa", self.issuer_name_label)
        issuer_form.addRow("NIF", self.issuer_nif_label)
        issuer_form.addRow("Morada", self.issuer_address_label)
        issuer_layout.addLayout(issuer_form)

        self.document_card = CardFrame()
        document_layout = QVBoxLayout(self.document_card)
        document_layout.setContentsMargins(12, 10, 12, 10)
        document_layout.setSpacing(8)
        document_title = QLabel("Documentos e Estado")
        document_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        document_layout.addWidget(document_title)
        document_form = QFormLayout()
        document_form.setContentsMargins(0, 0, 0, 0)
        document_form.setHorizontalSpacing(10)
        document_form.setVerticalSpacing(6)
        self.last_invoice_label = QLabel("-")
        self.quote_status_label = QLabel("-")
        self.guide_count_label = QLabel("-")
        self.record_origin_label = QLabel("-")
        self.fiscal_legal_label = QLabel("-")
        self.fiscal_source_label = QLabel("-")
        self.fiscal_entry_label = QLabel("-")
        self.fiscal_hash_control_label = QLabel("-")
        self.fiscal_comm_status_label = QLabel("-")
        self.fiscal_software_cert_label = QLabel("-")
        self.fiscal_hash_label = QLabel("-")
        self.fiscal_hash_label.setWordWrap(True)
        document_form.addRow("Última fatura", self.last_invoice_label)
        document_form.addRow("Estado orçamento", self.quote_status_label)
        document_form.addRow("Guias", self.guide_count_label)
        document_form.addRow("Origem", self.record_origin_label)
        document_form.addRow("N. legal", self.fiscal_legal_label)
        document_form.addRow("SourceBilling", self.fiscal_source_label)
        document_form.addRow("SystemEntry", self.fiscal_entry_label)
        document_form.addRow("HashControl", self.fiscal_hash_control_label)
        document_form.addRow("Comunicacao", self.fiscal_comm_status_label)
        document_form.addRow("Certificado", self.fiscal_software_cert_label)
        document_form.addRow("Hash", self.fiscal_hash_label)
        self.fiscal_comm_file_label = QLabel("-")
        self.fiscal_comm_file_label.setWordWrap(True)
        document_form.addRow("Lote AT", self.fiscal_comm_file_label)
        document_layout.addLayout(document_form)
        self.fiscal_hint_label = QLabel("PDF fiscal, SAF-T(PT) e preparacao AT a partir do registo aberto.")
        self.fiscal_hint_label.setProperty("role", "muted")
        self.fiscal_hint_label.setWordWrap(True)
        document_layout.addWidget(self.fiscal_hint_label)
        self.fiscal_file_label = QLabel("Ultimo lote AT: -")
        self.fiscal_file_label.setProperty("role", "muted")
        self.fiscal_file_label.setWordWrap(True)
        document_layout.addWidget(self.fiscal_file_label)
        fiscal_actions = QHBoxLayout()
        fiscal_actions.setContentsMargins(0, 0, 0, 0)
        fiscal_actions.setSpacing(6)
        self.export_saft_btn = QPushButton("Exportar SAF-T")
        self.export_saft_btn.clicked.connect(self._export_saft)
        self.prepare_at_btn = QPushButton("Preparar lote AT")
        self.prepare_at_btn.clicked.connect(self._prepare_at_batch)
        self.open_saft_btn = QPushButton("Abrir SAF-T")
        self.open_saft_btn.setProperty("variant", "secondary")
        self.open_saft_btn.clicked.connect(self._open_last_saft)
        self.open_at_btn = QPushButton("Abrir lote AT")
        self.open_at_btn.setProperty("variant", "secondary")
        self.open_at_btn.clicked.connect(self._open_last_at_batch)
        for button in (self.export_saft_btn, self.prepare_at_btn, self.open_saft_btn, self.open_at_btn):
            fiscal_actions.addWidget(button)
        fiscal_actions.addStretch(1)
        document_layout.addLayout(fiscal_actions)

        for widget in (self.customer_card, self.issuer_card, self.document_card):
            identity_split.addWidget(widget, 1)
        detail_layout.addLayout(identity_split)

        self.invoice_table = QTableWidget(0, 13)
        self.invoice_table.setHorizontalHeaderLabels(["Número", "Série", "ATCUD", "Guia", "Emissão", "Venc.", "Base", "IVA", "Total", "Saldo", "Estado"])
        self.invoice_table.verticalHeader().setVisible(False)
        self.invoice_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.invoice_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.invoice_table.setHorizontalHeaderLabels(
            ["N. legal", "Serie", "ATCUD", "Guia", "Emissao", "Venc.", "Base", "IVA", "Total", "Saldo", "Estado", "Origem", "Comunicacao"]
        )
        _configure_table(self.invoice_table, contents=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
        _configure_table(self.invoice_table, stretch=(0, 2), contents=(1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12))
        self.invoice_table.setMinimumHeight(_table_visible_height(self.invoice_table, 5, extra=12))
        invoice_header = self.invoice_table.horizontalHeader()
        for col, width in ((0, 176), (1, 88), (3, 108), (4, 96), (5, 96), (6, 88), (7, 78), (8, 96), (9, 96), (10, 98)):
            invoice_header.setSectionResizeMode(col, QHeaderView.Interactive)
            invoice_header.resizeSection(col, width)
        for col, width in ((0, 186), (3, 112), (11, 92), (12, 116)):
            invoice_header.setSectionResizeMode(col, QHeaderView.Interactive)
            invoice_header.resizeSection(col, width)
        invoice_header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.invoice_table.itemSelectionChanged.connect(self._sync_detail_buttons)
        self.invoice_table.itemDoubleClicked.connect(lambda *_args: self._edit_invoice())

        self.payment_table = QTableWidget(0, 6)
        self.payment_table.setHorizontalHeaderLabels(["Data", "Fatura", "Método", "Referência", "Valor", "Comprovativo"])
        self.payment_table.verticalHeader().setVisible(False)
        self.payment_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.payment_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.payment_table, stretch=(1, 3, 5), contents=(0, 2, 4))
        self.payment_table.setMinimumHeight(_table_visible_height(self.payment_table, 4, extra=12))
        payment_header = self.payment_table.horizontalHeader()
        for col, width in ((0, 92), (1, 176), (2, 112), (4, 92)):
            payment_header.setSectionResizeMode(col, QHeaderView.Interactive)
            payment_header.resizeSection(col, width)
        self.payment_table.itemSelectionChanged.connect(self._sync_detail_buttons)
        self.payment_table.itemDoubleClicked.connect(lambda *_args: self._edit_payment())

        self.invoice_generate_btn = QPushButton("Gerar PDF")
        self.invoice_generate_btn.clicked.connect(self._generate_invoice_pdf)
        self.invoice_open_btn = QPushButton("Abrir ficheiro")
        self.invoice_open_btn.setProperty("variant", "secondary")
        self.invoice_open_btn.clicked.connect(self._open_selected_invoice_file)
        self.invoice_remove_btn = QPushButton("Anular")
        self.invoice_remove_btn.setProperty("variant", "danger")
        self.invoice_remove_btn.clicked.connect(self._remove_invoice)
        self.payment_open_btn = QPushButton("Abrir comprovativo")
        self.payment_open_btn.setProperty("variant", "secondary")
        self.payment_open_btn.clicked.connect(self._open_selected_payment_file)
        self.payment_remove_btn = QPushButton("Remover")
        self.payment_remove_btn.setProperty("variant", "danger")
        self.payment_remove_btn.clicked.connect(self._remove_payment)

        detail_layout.addWidget(self._wrap_table("Faturas Associadas", self.invoice_table, self.invoice_generate_btn, self.invoice_open_btn, self.invoice_remove_btn))
        detail_layout.addWidget(self._wrap_table("Pagamentos / Comprovativos", self.payment_table, self.payment_open_btn, self.payment_remove_btn))

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        self._clear_detail()
        self._show_list()

    def _wrap_table(self, title_text: str, table: QTableWidget, *buttons: QPushButton) -> CardFrame:
        card = CardFrame()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)
        actions = QHBoxLayout()
        title = QLabel(title_text)
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        actions.addWidget(title)
        actions.addStretch(1)
        for button in buttons:
            actions.addWidget(button)
        layout.addLayout(actions)
        layout.addWidget(table)
        return card

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def refresh(self) -> None:
        previous_number = self.current_number
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(previous_number)
        dashboard = self.backend.billing_dashboard()
        card_payloads = (
            (_fmt_eur(dashboard.get("sold_total", 0)), "Total vendido", "info"),
            (_fmt_eur(dashboard.get("invoiced_total", 0)), "Total faturado", "success"),
            (_fmt_eur(dashboard.get("received_total", 0)), "Total recebido", "success"),
            (_fmt_eur(dashboard.get("balance_total", 0)), "Saldo em aberto", "warning" if float(dashboard.get("balance_total", 0) or 0) > 0 else "default"),
            (str(int(dashboard.get("overdue_count", 0) or 0)), "Faturas em atraso", "danger" if int(dashboard.get("overdue_count", 0) or 0) > 0 else "default"),
            (str(int(dashboard.get("open_payment_count", 0) or 0)), "Registos a cobrar", "warning" if int(dashboard.get("open_payment_count", 0) or 0) > 0 else "default"),
        )
        for card, (value, subtitle, tone) in zip(self.cards, card_payloads):
            card.set_data(value, subtitle)
            card.set_tone(tone)

        current_filter = self.filter_edit.currentText().strip()
        current_year = self.year_combo.currentText().strip() or "Todos"
        year_values = ["Todos"] + list(self.backend.billing_available_years())
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        self.year_combo.addItems(year_values)
        self.year_combo.setCurrentText(current_year if current_year in year_values else year_values[0])
        self.year_combo.blockSignals(False)
        self.rows = self.backend.billing_rows(current_filter, self.state_combo.currentText(), self.year_combo.currentText() or "Todos")
        self.filter_edit.blockSignals(True)
        if self.filter_edit.count() == 0:
            self.filter_edit.addItem("")
        known = {self.filter_edit.itemText(i) for i in range(self.filter_edit.count())}
        for row in self.rows:
            for value in (row.get("record_number", ""), row.get("orcamento_numero", ""), row.get("encomenda_numero", "")):
                if value and value not in known:
                    self.filter_edit.addItem(str(value))
                    known.add(str(value))
        self.filter_edit.setCurrentText(current_filter)
        self.filter_edit.blockSignals(False)
        _fill_table(
            self.table,
            [
                [
                    row.get("record_number", "-") or "-",
                    row.get("orcamento_numero", "-") or "-",
                    row.get("encomenda_numero", "-") or "-",
                    row.get("cliente", "-") or "-",
                    row.get("estado_encomenda", "-") or "-",
                    row.get("estado_expedicao", "-") or "-",
                    row.get("estado_faturacao", "-") or "-",
                    row.get("estado_pagamento", "-") or "-",
                    _fmt_eur(row.get("vendido", 0)),
                    _fmt_eur(row.get("faturado", 0)),
                    _fmt_eur(row.get("saldo", 0)),
                ]
                for row in self.rows
            ],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.table, row_index, str(row.get("estado_pagamento", "") or row.get("estado_faturacao", "") or ""))
            for col in (8, 9, 10):
                item = self.table.item(row_index, col)
                if item is not None:
                    item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
        if self.table.rowCount() == 0:
            self._clear_detail()
            self._show_list()
            self._sync_list_buttons()
            return
        select_index = 0
        if keep_detail:
            for index, row in enumerate(self.rows):
                if str(row.get("record_number", "") or "").strip() == previous_number:
                    select_index = index
                    break
        self.table.selectRow(select_index)
        if keep_detail and previous_number:
            try:
                self._load_detail(self.backend.billing_detail(previous_number))
                self._show_detail()
            except Exception:
                self._show_list()
        else:
            self._show_list()
        self._sync_list_buttons()

    def _show_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_buttons()

    def _show_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)
        self._sync_detail_buttons()

    def _selected_row(self) -> dict:
        current = self.table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def _selected_invoice(self) -> dict:
        current = self.invoice_table.currentItem()
        invoices = list(self.current_detail.get("invoices", []) or [])
        if current is None or current.row() >= len(invoices):
            return {}
        return invoices[current.row()]

    def _selected_payment(self) -> dict:
        current = self.payment_table.currentItem()
        payments = list(self.current_detail.get("payments", []) or [])
        if current is None or current.row() >= len(payments):
            return {}
        return payments[current.row()]

    def _sync_list_buttons(self) -> None:
        row = self._selected_row()
        has_row = bool(row)
        self.open_btn.setEnabled(has_row)
        self.remove_btn.setEnabled(has_row and bool(str(row.get("record_number", "") or "").strip()))

    def _sync_detail_buttons(self) -> None:
        invoice = self._selected_invoice()
        payment = self._selected_payment()
        selected_invoice_void = bool(invoice) and ("anulad" in str(invoice.get("estado", "") or "").strip().lower() or bool(invoice.get("anulada")))
        self.invoice_generate_btn.setEnabled(bool(self.current_number))
        self.invoice_open_btn.setEnabled(bool(str(invoice.get("caminho", "") or "").strip()))
        self.invoice_remove_btn.setEnabled(bool(invoice) and not selected_invoice_void)
        self.payment_open_btn.setEnabled(bool(str(payment.get("caminho_comprovativo", "") or "").strip()))
        self.payment_remove_btn.setEnabled(bool(payment))
        self.export_saft_btn.setEnabled(bool(self.current_number))
        self.prepare_at_btn.setEnabled(bool(self.current_number))
        self.open_saft_btn.setEnabled(bool(self.last_saft_export_path) and self.last_saft_export_record == self.current_number)
        at_path = str(invoice.get("communication_filename", "") or self.last_at_batch_path or self.current_detail.get("fiscal_communication_file", "") or "").strip()
        self.open_at_btn.setEnabled(bool(at_path))
        self._update_fiscal_panel()

    def _clear_detail(self) -> None:
        self.current_number = ""
        self.current_detail = {}
        self.last_saft_export_path = ""
        self.last_saft_export_record = ""
        self.last_at_batch_path = ""
        self.last_at_batch_record = ""
        self.number_label.setText("Novo registo")
        self.source_label.setText("Sem origem")
        self.quote_label.setText("-")
        self.order_label.setText("-")
        self.client_label.setText("-")
        self.guide_label.setText("-")
        self.customer_name_label.setText("-")
        self.customer_nif_label.setText("-")
        self.customer_contact_label.setText("-")
        self.customer_email_label.setText("-")
        self.customer_address_label.setText("-")
        self.issuer_name_label.setText("-")
        self.issuer_nif_label.setText("-")
        self.issuer_address_label.setText("-")
        self.last_invoice_label.setText("-")
        self.quote_status_label.setText("-")
        self.guide_count_label.setText("-")
        self.record_origin_label.setText("-")
        self.fiscal_legal_label.setText("-")
        self.fiscal_source_label.setText("-")
        self.fiscal_entry_label.setText("-")
        self.fiscal_hash_control_label.setText("-")
        self.fiscal_comm_status_label.setText("-")
        self.fiscal_software_cert_label.setText("-")
        self.fiscal_hash_label.setText("-")
        self.fiscal_hash_label.setToolTip("")
        self.fiscal_comm_file_label.setText("-")
        self.fiscal_comm_file_label.setToolTip("")
        self.fiscal_file_label.setText("Ultimo lote AT: -")
        self.sale_date_edit.setDate(_coerce_editor_qdate("", fallback_today=True))
        self.due_date_edit.setDate(_coerce_editor_qdate("", fallback_today=True))
        self.manual_value_spin.setValue(0.0)
        self.payment_override_combo.setCurrentText("Auto")
        self.notes_edit.clear()
        for key in self.summary_labels:
            self.summary_labels[key].setText("-")
        _fill_table(self.invoice_table, [])
        _fill_table(self.payment_table, [])
        _apply_state_chip(self.invoice_status_chip, "-", "-")
        _apply_state_chip(self.payment_status_chip, "-", "-")
        _apply_state_chip(self.order_status_chip, "-", "-")
        _apply_state_chip(self.shipping_status_chip, "-", "-")
        _set_panel_tone(self.header_card, "default")
        self._sync_detail_buttons()

    def _load_detail(self, detail: dict) -> None:
        self.current_detail = dict(detail or {})
        self.current_number = str(detail.get("numero", "") or "").strip()
        self.number_label.setText(self.current_number or "Registo")
        origem_txt = str(detail.get("origem", "") or "Sem origem").strip()
        self.source_label.setText(f"{origem_txt} | Venda {str(detail.get('data_venda', '') or '-').strip() or '-'}")
        self.quote_label.setText(str(detail.get("orcamento_numero", "") or "-") or "-")
        self.order_label.setText(str(detail.get("encomenda_numero", "") or "-") or "-")
        self.client_label.setText(str(detail.get("cliente_label", "") or "-") or "-")
        self.guide_label.setText(str(detail.get("last_guide", "") or "-") or "-")
        self.customer_name_label.setText(str(detail.get("cliente_nome", "") or "-") or "-")
        self.customer_nif_label.setText(str(detail.get("cliente_nif", "") or "-") or "-")
        self.customer_contact_label.setText(str(detail.get("cliente_contacto", "") or "-") or "-")
        self.customer_email_label.setText(str(detail.get("cliente_email", "") or "-") or "-")
        self.customer_address_label.setText(str(detail.get("cliente_morada", "") or "-") or "-")
        self.issuer_name_label.setText(str(detail.get("emitente_nome", "") or "-") or "-")
        self.issuer_nif_label.setText(str(detail.get("emitente_nif", "") or "-") or "-")
        self.issuer_address_label.setText(str(detail.get("emitente_morada", "") or "-") or "-")
        self.last_invoice_label.setText(str(detail.get("last_invoice", "") or "-") or "-")
        self.quote_status_label.setText(str(detail.get("quote_status", "") or "-") or "-")
        self.guide_count_label.setText(str(detail.get("guide_count", 0) or 0))
        self.record_origin_label.setText(str(detail.get("origem", "") or "-") or "-")
        self.last_at_batch_path = str(detail.get("fiscal_communication_file", "") or "").strip()
        self.last_at_batch_record = self.current_number if self.last_at_batch_path else ""
        self.sale_date_edit.setDate(_coerce_editor_qdate(str(detail.get("data_venda", "") or ""), fallback_today=True))
        self.due_date_edit.setDate(_coerce_editor_qdate(str(detail.get("data_vencimento", "") or ""), fallback_today=True))
        self.manual_value_spin.setValue(float(detail.get("valor_venda_manual", 0) or 0))
        manual_status = str(detail.get("estado_pagamento_manual", "") or "").strip()
        self.payment_override_combo.setCurrentText(manual_status if manual_status else "Auto")
        self.notes_edit.setPlainText(str(detail.get("obs", "") or "").strip())
        _apply_state_chip(self.invoice_status_chip, str(detail.get("estado_faturacao", "") or "-"), f"Faturação: {str(detail.get('estado_faturacao', '') or '-').strip()}")
        _apply_state_chip(self.payment_status_chip, str(detail.get("estado_pagamento", "") or "-"), f"Pagamento: {str(detail.get('estado_pagamento', '') or '-').strip()}")
        _apply_state_chip(self.order_status_chip, str(detail.get("order_status", "") or "-"))
        _apply_state_chip(self.shipping_status_chip, str(detail.get("shipping_status", "") or "-"))
        _set_panel_tone(self.header_card, _state_tone(str(detail.get("estado_pagamento", "") or detail.get("estado_faturacao", "") or "")))
        self.summary_labels["sold"].setText(_fmt_eur(detail.get("valor_venda", 0)))
        self.summary_labels["invoiced"].setText(_fmt_eur(detail.get("valor_faturado", 0)))
        self.summary_labels["received"].setText(_fmt_eur(detail.get("valor_recebido", 0)))
        self.summary_labels["balance"].setText(_fmt_eur(detail.get("saldo", 0)))
        self.summary_labels["uninvoiced"].setText(_fmt_eur(detail.get("por_faturar", 0)))
        self.summary_labels["invoice_count"].setText(str(len(list(detail.get("invoices", []) or []))))
        _fill_table(
            self.invoice_table,
            [
                [
                    row.get("legal_invoice_no", "-") or row.get("numero_fatura", "-"),
                    row.get("serie", "-"),
                    row.get("atcud", "-") or "-",
                    row.get("guia_numero", "-"),
                    row.get("data_emissao", "-"),
                    row.get("data_vencimento", "-"),
                    _fmt_eur(row.get("subtotal", 0)),
                    _fmt_eur(row.get("valor_iva", 0)),
                    _fmt_eur(row.get("valor_total", 0)),
                    _fmt_eur(row.get("saldo", 0)),
                    row.get("estado", "-"),
                    row.get("source_billing", "-"),
                    row.get("communication_status", "-"),
                ]
                for row in list(detail.get("invoices", []) or [])
            ],
            align_center_from=1,
        )
        for row_index, row in enumerate(list(detail.get("invoices", []) or [])):
            _paint_table_row(self.invoice_table, row_index, str(row.get("estado", "") or ""))
            for col in (6, 7, 8, 9):
                item = self.invoice_table.item(row_index, col)
                if item is not None:
                    item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            for col in range(self.invoice_table.columnCount()):
                item = self.invoice_table.item(row_index, col)
                if item is not None:
                    item.setToolTip(str(item.text() or ""))
        _fill_table(
            self.payment_table,
            [
                [
                    row.get("data_pagamento", "-"),
                    row.get("fatura_label", "-") or "-",
                    row.get("metodo", "-"),
                    row.get("referencia", "-"),
                    _fmt_eur(row.get("valor", 0)),
                    row.get("titulo_comprovativo", "-") or ("Anexo" if str(row.get("caminho_comprovativo", "") or "").strip() else "-"),
                ]
                for row in list(detail.get("payments", []) or [])
            ],
            align_center_from=0,
        )
        for row_index, row in enumerate(list(detail.get("payments", []) or [])):
            _paint_table_row(self.payment_table, row_index, str(detail.get("estado_pagamento", "") or ""))
            item = self.payment_table.item(row_index, 4)
            if item is not None:
                item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            for col in range(self.payment_table.columnCount()):
                row_item = self.payment_table.item(row_index, col)
                if row_item is not None:
                    row_item.setToolTip(str(row_item.text() or ""))
        self._update_fiscal_panel()
        self._sync_detail_buttons()

    def _update_fiscal_panel(self) -> None:
        detail = dict(self.current_detail or {})
        invoice = self._selected_invoice()
        legal_no = str(invoice.get("legal_invoice_no", "") or invoice.get("numero_fatura", "") or detail.get("fiscal_legal_invoice_no", "") or "-").strip() or "-"
        source_billing = str(invoice.get("source_billing", "") or detail.get("fiscal_source_billing", "") or "-").strip() or "-"
        system_entry = str(invoice.get("system_entry_date", "") or detail.get("fiscal_system_entry_date", "") or "-").strip() or "-"
        hash_control = str(invoice.get("hash_control", "") or detail.get("fiscal_hash_control", "") or "-").strip() or "-"
        communication = str(invoice.get("communication_status", "") or detail.get("fiscal_communication_status", "") or "-").strip() or "-"
        software_cert = str(detail.get("fiscal_software_cert", "") or "-").strip() or "-"
        hash_value = str(invoice.get("hash", "") or detail.get("fiscal_hash", "") or "").strip()
        communication_file = str(invoice.get("communication_filename", "") or detail.get("fiscal_communication_file", "") or "").strip()
        self.fiscal_legal_label.setText(legal_no)
        self.fiscal_source_label.setText(source_billing)
        self.fiscal_entry_label.setText(system_entry)
        self.fiscal_hash_control_label.setText(hash_control)
        self.fiscal_comm_status_label.setText(communication)
        self.fiscal_software_cert_label.setText(software_cert)
        hash_preview = hash_value if len(hash_value) <= 36 else f"{hash_value[:33]}..."
        self.fiscal_hash_label.setText(hash_preview or "-")
        self.fiscal_hash_label.setToolTip(hash_value or "")
        self.fiscal_comm_file_label.setText(Path(communication_file).name if communication_file else "-")
        self.fiscal_comm_file_label.setToolTip(communication_file or "")
        self.fiscal_file_label.setText(f"Ultimo lote AT: {Path(communication_file).name}" if communication_file else "Ultimo lote AT: -")
        if communication_file:
            self.last_at_batch_path = communication_file
            self.last_at_batch_record = self.current_number

    def _open_selected(self) -> None:
        row = self._selected_row()
        if not row:
            QMessageBox.warning(self, "Faturação", "Seleciona um registo ou origem de venda.")
            return
        try:
            detail = self.backend.billing_open_record(
                source_type=str(row.get("source_type", "") or "").strip(),
                source_number=str(row.get("source_number", "") or "").strip(),
                record_number=str(row.get("record_number", "") or "").strip(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()
        self._show_detail()

    def _remove_selected(self) -> None:
        row = self._selected_row()
        record_number = str(row.get("record_number", "") or "").strip()
        if not record_number:
            QMessageBox.warning(self, "Faturação", "O registo ainda não foi criado.")
            return
        if QMessageBox.question(self, "Faturação", f"Remover o registo {record_number}?") != QMessageBox.Yes:
            return
        try:
            self.backend.billing_remove(record_number)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self.refresh()

    def _save_record(self) -> None:
        if not self.current_number:
            return
        payload = {
            "numero": self.current_number,
            "data_venda": self.sale_date_edit.date().toString("yyyy-MM-dd"),
            "data_vencimento": self.due_date_edit.date().toString("yyyy-MM-dd"),
            "valor_venda_manual": float(self.manual_value_spin.value() or 0),
            "estado_pagamento_manual": "" if self.payment_override_combo.currentText().strip() == "Auto" else self.payment_override_combo.currentText().strip(),
            "obs": self.notes_edit.toPlainText().strip(),
        }
        try:
            detail = self.backend.billing_save(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _remove_current(self) -> None:
        if not self.current_number:
            return
        if QMessageBox.question(self, "Faturação", f"Remover o registo {self.current_number}?") != QMessageBox.Yes:
            return
        try:
            self.backend.billing_remove(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._clear_detail()
        self.refresh()
        self._show_list()

    def _browse_file_into(self, target_edit: QLineEdit, title: str) -> None:
        current = target_edit.text().strip()
        start_dir = str(Path(current).parent) if current else str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir)
        if path:
            target_edit.setText(path)

    def _invoice_dialog(self, initial: dict | None = None) -> dict | None:
        initial_row = dict(initial or {})
        try:
            row = dict(self.backend.billing_invoice_defaults(self.current_number, str(initial_row.get("id", "") or "").strip()))
            row.update(initial_row)
        except Exception:
            row = initial_row
        dialog = QDialog(self)
        dialog.setWindowTitle("Fatura")
        dialog.resize(640, 360)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        number_edit = QLineEdit(str(row.get("numero_fatura", "") or "").strip())
        series_edit = QLineEdit(str(row.get("serie", "") or "").strip())
        guide_combo = QComboBox()
        guide_combo.setEditable(True)
        guide_combo.addItem("", "")
        guide_seen: set[str] = set()
        for option_group in (
            list(row.get("guide_options", []) or []),
            list(self.current_detail.get("guide_options", []) or []),
            list(self.current_detail.get("guides", []) or []),
        ):
            for option in option_group:
                guide_number = str(option.get("numero", "") or "").strip()
                if not guide_number or guide_number in guide_seen:
                    continue
                guide_seen.add(guide_number)
                guide_combo.addItem(str(option.get("label", "") or guide_number).strip(), guide_number)
        guide_combo.setCurrentText(str(row.get("guia_numero", "") or "").strip())
        if guide_combo.currentText().strip() and guide_combo.findData(guide_combo.currentText().strip()) < 0:
            current_guide = guide_combo.currentText().strip()
            guide_combo.addItem(current_guide, current_guide)
        issue_edit = QDateEdit()
        issue_edit.setCalendarPopup(True)
        issue_edit.setDisplayFormat("yyyy-MM-dd")
        issue_edit.setDate(_coerce_editor_qdate(str(row.get("data_emissao", "") or ""), fallback_today=True))
        due_edit = QDateEdit()
        due_edit.setCalendarPopup(True)
        due_edit.setDisplayFormat("yyyy-MM-dd")
        due_edit.setDate(_coerce_editor_qdate(str(row.get("data_vencimento", "") or ""), fallback_today=True))
        amount_spin = QDoubleSpinBox()
        amount_spin.setRange(0.0, 100000000.0)
        amount_spin.setDecimals(2)
        amount_spin.setPrefix("EUR ")
        amount_spin.setValue(float(row.get("valor_total", 0) or 0))
        iva_spin = QDoubleSpinBox()
        iva_spin.setRange(0.0, 100.0)
        iva_spin.setDecimals(2)
        iva_spin.setSuffix(" %")
        iva_spin.setValue(float(row.get("iva_perc", 23) or 23))
        path_edit = QLineEdit(str(row.get("caminho", "") or "").strip())
        browse_btn = QPushButton("...")
        _cap_width(browse_btn, 38)
        browse_btn.clicked.connect(lambda: self._browse_file_into(path_edit, "Associar fatura"))
        path_host = QWidget()
        path_layout = QHBoxLayout(path_host)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.setSpacing(6)
        path_layout.addWidget(path_edit, 1)
        path_layout.addWidget(browse_btn)
        obs_edit = QTextEdit(str(row.get("obs", "") or "").strip())
        obs_edit.setMinimumHeight(72)
        form.addRow("Número", number_edit)
        form.addRow("Série", series_edit)
        form.addRow("Guia", guide_combo)
        form.addRow("Emissão", issue_edit)
        form.addRow("Vencimento", due_edit)
        form.addRow("Taxa IVA", iva_spin)
        form.addRow("Valor", amount_spin)
        form.addRow("Ficheiro", path_host)
        form.addRow("Obs.", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "id": str(row.get("id", "") or "").strip(),
            "doc_type": str(row.get("doc_type", "") or "FT").strip() or "FT",
            "numero_fatura": number_edit.text().strip(),
            "serie": series_edit.text().strip(),
            "serie_id": str(row.get("serie_id", "") or series_edit.text()).strip(),
            "seq_num": int(float(row.get("seq_num", 0) or 0)),
            "at_validation_code": str(row.get("at_validation_code", "") or "").strip(),
            "atcud": str(row.get("atcud", "") or "").strip(),
            "guia_numero": str(guide_combo.currentData() or guide_combo.currentText()).strip(),
            "data_emissao": issue_edit.date().toString("yyyy-MM-dd"),
            "data_vencimento": due_edit.date().toString("yyyy-MM-dd"),
            "moeda": str(row.get("moeda", "") or "EUR").strip() or "EUR",
            "iva_perc": float(iva_spin.value() or 0),
            "subtotal": float(row.get("subtotal", 0) or 0),
            "valor_iva": float(row.get("valor_iva", 0) or 0),
            "valor_total": float(amount_spin.value() or 0),
            "caminho": path_edit.text().strip(),
            "obs": obs_edit.toPlainText().strip(),
        }

    def _payment_dialog(self, initial: dict | None = None) -> dict | None:
        row = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Pagamento")
        dialog.resize(620, 320)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        invoice_combo = QComboBox()
        invoice_combo.addItem("Sem fatura específica", "")
        current_invoice = str(row.get("fatura_id", "") or "").strip()
        for option in list(self.current_detail.get("invoice_options", []) or []):
            invoice_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("id", "") or "").strip())
        index = invoice_combo.findData(current_invoice)
        if index >= 0:
            invoice_combo.setCurrentIndex(index)
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("yyyy-MM-dd")
        date_edit.setDate(_coerce_editor_qdate(str(row.get("data_pagamento", "") or ""), fallback_today=True))
        amount_spin = QDoubleSpinBox()
        amount_spin.setRange(0.0, 100000000.0)
        amount_spin.setDecimals(2)
        amount_spin.setPrefix("EUR ")
        amount_spin.setValue(float(row.get("valor", 0) or 0))
        method_combo = QComboBox()
        method_combo.setEditable(True)
        method_combo.addItems(["", "Transferência", "Dinheiro", "MB", "Cheque", "Outro"])
        method_combo.setCurrentText(str(row.get("metodo", "") or "").strip())
        ref_edit = QLineEdit(str(row.get("referencia", "") or "").strip())
        title_edit = QLineEdit(str(row.get("titulo_comprovativo", "") or "").strip())
        path_edit = QLineEdit(str(row.get("caminho_comprovativo", "") or "").strip())
        browse_btn = QPushButton("...")
        _cap_width(browse_btn, 38)
        browse_btn.clicked.connect(lambda: self._browse_file_into(path_edit, "Associar comprovativo"))
        path_host = QWidget()
        path_layout = QHBoxLayout(path_host)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.setSpacing(6)
        path_layout.addWidget(path_edit, 1)
        path_layout.addWidget(browse_btn)
        obs_edit = QTextEdit(str(row.get("obs", "") or "").strip())
        obs_edit.setMinimumHeight(72)
        form.addRow("Fatura", invoice_combo)
        form.addRow("Data", date_edit)
        form.addRow("Valor", amount_spin)
        form.addRow("Método", method_combo)
        form.addRow("Referência", ref_edit)
        form.addRow("Título", title_edit)
        form.addRow("Comprovativo", path_host)
        form.addRow("Obs.", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "id": str(row.get("id", "") or "").strip(),
            "fatura_id": str(invoice_combo.currentData() or "").strip(),
            "data_pagamento": date_edit.date().toString("yyyy-MM-dd"),
            "valor": float(amount_spin.value() or 0),
            "metodo": method_combo.currentText().strip(),
            "referencia": ref_edit.text().strip(),
            "titulo_comprovativo": title_edit.text().strip(),
            "caminho_comprovativo": path_edit.text().strip(),
            "obs": obs_edit.toPlainText().strip(),
        }

    def _add_invoice(self) -> None:
        payload = self._invoice_dialog()
        if payload is None:
            return
        try:
            detail = self.backend.billing_add_invoice(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _generate_invoice_pdf(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Faturação", "Abre primeiro um registo de faturação.")
            return
        invoice = self._selected_invoice()
        if invoice and (bool(invoice.get("anulada")) or "anulad" in str(invoice.get("estado", "") or "").strip().lower()):
            QMessageBox.information(self, "Faturação", "A fatura selecionada está anulada. Seleciona outra ou limpa a seleção para criar nova fatura.")
            return
        payload = self._invoice_dialog(invoice if invoice else None)
        if payload is None:
            return
        try:
            detail = self.backend.billing_generate_invoice_pdf(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _edit_invoice(self) -> None:
        invoice = self._selected_invoice()
        if not invoice:
            QMessageBox.warning(self, "Faturação", "Seleciona uma fatura.")
            return
        if bool(invoice.get("anulada")) or "anulad" in str(invoice.get("estado", "") or "").strip().lower():
            QMessageBox.information(self, "Faturação", "As faturas anuladas já não podem ser editadas.")
            return
        payload = self._invoice_dialog(invoice)
        if payload is None:
            return
        try:
            detail = self.backend.billing_add_invoice(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _remove_invoice(self) -> None:
        invoice = self._selected_invoice()
        if not invoice:
            QMessageBox.warning(self, "Faturação", "Seleciona uma fatura.")
            return
        if bool(invoice.get("anulada")) or "anulad" in str(invoice.get("estado", "") or "").strip().lower():
            QMessageBox.information(self, "Faturação", "A fatura selecionada já está anulada.")
            return
        number = str(invoice.get("numero_fatura", "") or invoice.get("id", "") or "").strip()
        reason = "Anulada pelo utilizador."
        if QMessageBox.question(self, "Faturação", f"Anular a fatura {number}?") != QMessageBox.Yes:
            return
        reason, accepted = QInputDialog.getText(self, "Faturação", f"Motivo da anulação de {number}:", text=reason)
        if not accepted:
            return
        reason = str(reason or "").strip()
        if not reason:
            QMessageBox.warning(self, "Faturação", "Indica o motivo da anulação.")
            return
        try:
            detail = self.backend.billing_cancel_invoice(self.current_number, str(invoice.get("id", "") or "").strip(), reason)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _add_payment(self) -> None:
        payload = self._payment_dialog()
        if payload is None:
            return
        try:
            detail = self.backend.billing_add_payment(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _edit_payment(self) -> None:
        payment = self._selected_payment()
        if not payment:
            QMessageBox.warning(self, "Faturação", "Seleciona um pagamento.")
            return
        payload = self._payment_dialog(payment)
        if payload is None:
            return
        try:
            detail = self.backend.billing_add_payment(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _remove_payment(self) -> None:
        payment = self._selected_payment()
        if not payment:
            QMessageBox.warning(self, "Faturação", "Seleciona um pagamento.")
            return
        if QMessageBox.question(self, "Faturação", "Remover este pagamento?") != QMessageBox.Yes:
            return
        try:
            detail = self.backend.billing_remove_payment(self.current_number, str(payment.get("id", "") or "").strip())
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
            return
        self._load_detail(detail)
        self.refresh()

    def _export_saft(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Faturacao", "Abre primeiro um registo de faturacao.")
            return
        try:
            output = str(self.backend.billing_export_record_saft_pt(self.current_number))
            detail = self.backend.billing_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Faturacao", str(exc))
            return
        self.last_saft_export_path = output
        self.last_saft_export_record = self.current_number
        self._load_detail(detail)
        self.refresh()
        QMessageBox.information(self, "Faturacao", f"SAF-T exportado com sucesso.\n{output}")

    def _prepare_at_batch(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Faturacao", "Abre primeiro um registo de faturacao.")
            return
        try:
            output = str(self.backend.billing_prepare_record_at_communication_batch(self.current_number))
            detail = self.backend.billing_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Faturacao", str(exc))
            return
        self.last_at_batch_path = output
        self.last_at_batch_record = self.current_number
        self._load_detail(detail)
        self.refresh()
        QMessageBox.information(self, "Faturacao", f"Lote AT preparado com sucesso.\n{output}")

    def _open_last_saft(self) -> None:
        path = str(self.last_saft_export_path or "").strip()
        if not path or self.last_saft_export_record != self.current_number:
            QMessageBox.warning(self, "Faturacao", "Ainda nao existe um ficheiro SAF-T aberto para este registo.")
            return
        try:
            self.backend.billing_open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "Faturacao", str(exc))

    def _open_last_at_batch(self) -> None:
        invoice = self._selected_invoice()
        path = str(invoice.get("communication_filename", "") or self.last_at_batch_path or self.current_detail.get("fiscal_communication_file", "") or "").strip()
        if not path:
            QMessageBox.warning(self, "Faturacao", "Ainda nao existe lote AT preparado para abrir.")
            return
        try:
            self.backend.billing_open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "Faturacao", str(exc))

    def _open_selected_invoice_file(self) -> None:
        invoice = self._selected_invoice()
        path = str(invoice.get("caminho", "") or "").strip()
        if not path:
            QMessageBox.warning(self, "Faturação", "A fatura não tem ficheiro associado.")
            return
        try:
            self.backend.billing_open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))

    def _open_selected_payment_file(self) -> None:
        payment = self._selected_payment()
        path = str(payment.get("caminho_comprovativo", "") or "").strip()
        if not path:
            QMessageBox.warning(self, "Faturação", "O pagamento não tem comprovativo associado.")
            return
        try:
            self.backend.billing_open_path(path)
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))
