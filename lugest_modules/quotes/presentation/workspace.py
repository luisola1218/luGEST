from __future__ import annotations
from lugest_modules.quotes.presentation.workspace_styles import STYLE_OVERVIEW_BAND, STYLE_QUOTE_DETAIL_SCROLL, STYLE_QUOTE_NOTES_CARD, STYLE_NOTES_TABS, STYLE_TRANSPORT_FORM_CARD, STYLE_TRANSPORT_ACTIONS_CARD, STYLE_TOTAL_PANEL, STYLE_SUMMARY_ROWS_HOST, STYLE_CONTROLS_PANEL, STYLE_DISCOUNT_STATUS, STYLE_LINE_TOOLS_TABS, STYLE_LINES_TABLE, STYLE_SELECTED_LINE_FOOTER, STYLE_INSPECTOR_TABS, STYLE_SUMMARY_CONTROL, STYLE_PANEL
from PySide6.QtCore import QDate, QEvent, QTimer, Qt
from PySide6.QtWidgets import QAbstractItemView, QComboBox, QFormLayout, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMenu, QPushButton, QScrollArea, QSizePolicy, QSplitter, QStackedWidget, QTabWidget, QTableWidget, QTextEdit, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_common import apply_state_chip as _apply_state_chip, configure_table as _configure_table, repolish as _repolish, set_table_columns as _set_table_columns
from lugest_qt.ui.pages.runtime_support import LIST_TABLE_FONT_PX, LIST_TABLE_ROW_PX
from lugest_qt.ui.widgets import CardFrame, ClickableDateEdit as QDateEdit, FlexibleDecimalSpinBox as QDoubleSpinBox


from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class QuoteWorkspaceActions:
    _add_laser_batch_lines: Callable[..., object]
    _add_line: Callable[..., object]
    _append_pdf_note: Callable[..., object]
    _apply_default_delivery_deadline: Callable[..., object]
    _apply_transport_calc: Callable[..., object]
    _check_selected_line_weight: Callable[..., object]
    _clear_transport: Callable[..., object]
    _configure_laser_profiles: Callable[..., object]
    _configure_operation_profiles: Callable[..., object]
    _convert_quote: Callable[..., object]
    _create_quote_purchase_note: Callable[..., object]
    _edit_line: Callable[..., object]
    _fill_client_from_combo: Callable[..., object]
    _fill_pdf_notes_from_context: Callable[..., object]
    _handle_quote_line_item_changed: Callable[..., object]
    _handle_quote_lines_sort: Callable[..., object]
    _manage_assembly_models: Callable[..., object]
    _manage_saved_conjuntos: Callable[..., object]
    _new_quote: Callable[..., object]
    _new_structure_quote: Callable[..., object]
    _open_calculated_assembly_builder: Callable[..., object]
    _open_laser_nesting: Callable[..., object]
    _open_line_drawing: Callable[..., object]
    _open_profile_step_igs_quote_builder: Callable[..., object]
    _open_selected_quote: Callable[..., object]
    _open_structure_quote_builder: Callable[..., object]
    _pick_quote_discount_groups: Callable[..., object]
    _prepare_selected_line_for_production: Callable[..., object]
    _preview_quote: Callable[..., object]
    _print_quote_pdf: Callable[..., object]
    _recalc_transport_calc: Callable[..., object]
    _refresh_quote_identity_label: Callable[..., object]
    _remove_line: Callable[..., object]
    _remove_quote: Callable[..., object]
    _render_quote_lines: Callable[..., object]
    _reset_quote_save_button: Callable[..., object]
    _save_quote: Callable[..., object]
    _save_quote_pdf: Callable[..., object]
    _save_selected_lines_as_group: Callable[..., object]
    _set_quote_state: Callable[..., object]
    _show_list: Callable[..., object]
    _sync_list_buttons: Callable[..., object]
    _sync_quote_selected_line_actions: Callable[..., object]
    _toggle_quote_inspector: Callable[..., object]
    refresh: Callable[..., object]


class QuoteWorkspace(QWidget):
    LINE_COL_MARK = 0
    LINE_COL_TYPE = 1
    LINE_COL_REFERENCE = 2
    LINE_COL_EXTERNAL_REFERENCE = 3
    LINE_COL_DESCRIPTION = 4
    LINE_COL_MATERIAL = 5
    LINE_COL_UNIT = 6
    LINE_COL_OPERATION = 7
    LINE_COL_TIME = 8
    LINE_COL_QUANTITY = 9
    LINE_COL_PRICE = 10
    LINE_COL_DISCOUNTED_PRICE = 11
    LINE_COL_TOTAL = 12
    LINE_COL_ASSEMBLY = 13

    def build(self, actions: QuoteWorkspaceActions) -> None:
        self._combo_click_targets: dict[QWidget, QComboBox] = {}
        self._filter_refresh_timer = QTimer(self)
        self._filter_refresh_timer.setSingleShot(True)
        self._filter_refresh_timer.setInterval(180)
        self._filter_refresh_timer.timeout.connect(actions.refresh)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)
        self.view_stack = QStackedWidget()
        root.addWidget(self.view_stack, 1)

        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)

        filters = CardFrame()
        filters.set_tone("default")
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(16, 12, 16, 12)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(6)
        filter_title = QLabel("Carteira comercial")
        filter_title.setStyleSheet("font-size: 14px; font-weight: 900; color: #0f172a;")
        filter_hint = QLabel("Pesquisa, acompanha e abre cada proposta sem perder o contexto.")
        filter_hint.setProperty("role", "muted")
        filter_hint.setStyleSheet("font-size: 10px;")
        filter_heading = QVBoxLayout()
        filter_heading.setContentsMargins(0, 0, 0, 0)
        filter_heading.setSpacing(1)
        filter_heading.addWidget(filter_title)
        filter_heading.addWidget(filter_hint)
        filters_layout.addLayout(filter_heading, 0, 0, 1, 3)
        self.quote_list_count_label = QLabel("0 orçamentos")
        self.quote_list_count_label.setProperty("role", "state_chip")
        self.quote_list_count_label.setAlignment(Qt.AlignCenter)
        self.quote_list_count_label.setMinimumWidth(120)
        filters_layout.addWidget(self.quote_list_count_label, 0, 3, 1, 1, Qt.AlignRight | Qt.AlignVCenter)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Número, cliente ou encomenda")
        self.filter_edit.lineEdit().textChanged.connect(
            lambda _text: self._filter_refresh_timer.start()
        )
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Todos", "Em edicao", "Enviado", "Aprovado", "Rejeitado", "Convertido"])
        self.state_combo.currentTextChanged.connect(actions.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(actions.refresh)
        self.new_quote_btn = QPushButton("Novo orcamento")
        self.new_quote_btn.clicked.connect(actions._new_quote)
        self.structure_quote_btn = QPushButton("Orcamento Estruturas")
        self.structure_quote_btn.setProperty("variant", "secondary")
        self.structure_quote_btn.clicked.connect(actions._new_structure_quote)
        self.open_quote_btn = QPushButton("Abrir orcamento")
        self.open_quote_btn.clicked.connect(actions._open_selected_quote)
        self.remove_quote_btn = QPushButton("Remover")
        self.remove_quote_btn.setProperty("variant", "danger")
        self.remove_quote_btn.clicked.connect(actions._remove_quote)
        filters_layout.addWidget(QLabel("Pesquisa"), 1, 0)
        filters_layout.addWidget(QLabel("Estado"), 1, 1)
        filters_layout.addWidget(QLabel("Ano"), 1, 2)
        filters_layout.addWidget(QLabel("Ações"), 1, 3)
        filters_layout.addWidget(self.filter_edit, 2, 0)
        filters_layout.addWidget(self.state_combo, 2, 1)
        filters_layout.addWidget(self.year_combo, 2, 2)
        action_host = QWidget()
        action_layout = QHBoxLayout(action_host)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(6)
        for button, width in (
            (self.new_quote_btn, 112),
            (self.structure_quote_btn, 142),
            (self.open_quote_btn, 112),
            (self.remove_quote_btn, 92),
        ):
            button.setProperty("compact", "true")
            button.setMinimumWidth(width)
            action_layout.addWidget(button)
        action_layout.addStretch(1)
        filters_layout.addWidget(action_host, 2, 3)
        filters_layout.setColumnStretch(0, 4)
        filters_layout.setColumnStretch(1, 2)
        filters_layout.setColumnStretch(2, 2)
        filters_layout.setColumnStretch(3, 5)
        filters.setMaximumHeight(126)
        list_layout.addWidget(filters)

        overview_band = QFrame()
        overview_band.setObjectName("QuoteOverviewBand")
        overview_band.setStyleSheet(
            STYLE_OVERVIEW_BAND
        )
        overview_layout = QGridLayout(overview_band)
        overview_layout.setContentsMargins(12, 9, 12, 9)
        overview_layout.setHorizontalSpacing(10)
        overview_layout.setVerticalSpacing(0)
        self.quote_metric_labels: dict[str, QLabel] = {}
        for metric_index, (metric_key, metric_title, accent) in enumerate(
            (
                ("value", "Valor visível", "#24746b"),
                ("editing", "Em edição", "#7b91a8"),
                ("sent", "Enviados", "#b48631"),
                ("approved", "Aprovados", "#4f7f5d"),
            )
        ):
            metric = QFrame()
            metric.setObjectName("QuoteMetric")
            metric.setStyleSheet(
                f"QFrame#QuoteMetric {{ background: transparent; border: 0; border-left: 3px solid {accent}; }}"
            )
            metric_layout = QVBoxLayout(metric)
            metric_layout.setContentsMargins(11, 2, 8, 2)
            metric_layout.setSpacing(2)
            title = QLabel(metric_title)
            title.setProperty("role", "muted")
            title.setStyleSheet("font-size: 9px; font-weight: 800;")
            value = QLabel("0")
            value.setStyleSheet("font-size: 16px; font-weight: 900; color: #132238;")
            metric_layout.addWidget(title)
            metric_layout.addWidget(value)
            overview_layout.addWidget(metric, 0, metric_index)
            overview_layout.setColumnStretch(metric_index, 1)
            self.quote_metric_labels[metric_key] = value
        overview_band.setMaximumHeight(72)
        list_layout.addWidget(overview_band)

        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 14)
        table_heading = QHBoxLayout()
        table_heading.setContentsMargins(0, 0, 0, 0)
        table_title = QLabel("Propostas")
        table_title.setStyleSheet("font-size: 15px; font-weight: 900; color: #0f172a;")
        table_help = QLabel("Duplo clique para abrir")
        table_help.setProperty("role", "muted")
        table_help.setStyleSheet("font-size: 9px;")
        table_heading.addWidget(table_title)
        table_heading.addStretch(1)
        table_heading.addWidget(table_help)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Número", "Cliente", "Data", "Estado", "Total", "Encomenda"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setStyleSheet(
            f"QTableWidget {{ font-size: {LIST_TABLE_FONT_PX}px; }}"
            f" QHeaderView::section {{ font-size: {LIST_TABLE_FONT_PX}px; padding: 8px 10px; font-weight: 800; }}"
        )
        self.table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        self.table.horizontalHeader().setFixedHeight(36)
        _configure_table(self.table, stretch=(1,), contents=(2, 3, 4))
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 205),
                (1, "stretch", 0),
                (2, "fixed", 126),
                (3, "fixed", 138),
                (4, "fixed", 150),
                (5, "fixed", 190),
            ],
        )
        self.table.itemSelectionChanged.connect(actions._sync_list_buttons)
        self.table.itemDoubleClicked.connect(lambda *_args: actions._open_selected_quote())
        table_layout.addLayout(table_heading)
        table_layout.addWidget(self.table)
        list_layout.addWidget(table_card, 1)

        self.detail_page = QWidget()
        detail_outer = QVBoxLayout(self.detail_page)
        detail_outer.setContentsMargins(0, 0, 0, 0)
        detail_outer.setSpacing(0)
        self.quote_detail_scroll = QScrollArea()
        self.quote_detail_scroll.setWidgetResizable(True)
        self.quote_detail_scroll.setFrameShape(QFrame.NoFrame)
        self.quote_detail_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.quote_detail_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.quote_detail_scroll.setStyleSheet(
            STYLE_QUOTE_DETAIL_SCROLL
        )
        detail_outer.addWidget(self.quote_detail_scroll)
        self.quote_detail_host = QWidget()
        self.quote_detail_scroll.setWidget(self.quote_detail_host)
        detail_layout = QVBoxLayout(self.quote_detail_host)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)

        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(10, 7, 10, 7)
        detail_actions_layout.setSpacing(5)
        back_btn = QPushButton("Voltar a lista")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(actions._show_list)
        save_btn = QPushButton("Guardar")
        save_btn.setProperty("variant", "success")
        save_btn.clicked.connect(actions._save_quote)
        self.quote_save_btn = save_btn
        edit_btn = QPushButton("Em edicao")
        edit_btn.setProperty("variant", "secondary")
        edit_btn.clicked.connect(lambda: actions._set_quote_state("Em edição"))
        sent_btn = QPushButton("Enviar")
        sent_btn.setProperty("variant", "secondary")
        sent_btn.clicked.connect(lambda: actions._set_quote_state("Enviado"))
        approve_btn = QPushButton("Aprovado")
        approve_btn.setProperty("variant", "success")
        approve_btn.clicked.connect(lambda: actions._set_quote_state("Aprovado"))
        reject_btn = QPushButton("Rejeitado")
        reject_btn.setProperty("variant", "rejected")
        reject_btn.setText("✕  Rejeitado")
        reject_btn.clicked.connect(lambda: actions._set_quote_state("Rejeitado"))
        convert_btn = QPushButton("Criar encomenda")
        convert_btn.clicked.connect(actions._convert_quote)
        purchase_note_btn = QPushButton("Nota encomenda")
        purchase_note_btn.setProperty("variant", "secondary")
        purchase_note_btn.clicked.connect(actions._create_quote_purchase_note)
        pdf_actions_btn = QPushButton("PDF  ▾")
        pdf_actions_btn.setProperty("variant", "secondary")
        pdf_menu = QMenu(pdf_actions_btn)
        pdf_menu.addAction("Pré-visualizar PDF", actions._preview_quote)
        pdf_menu.addAction("Guardar cópia PDF", actions._save_quote_pdf)
        pdf_menu.addAction("Imprimir PDF", actions._print_quote_pdf)
        pdf_actions_btn.setMenu(pdf_menu)
        action_widths = (
            (back_btn, 86),
            (save_btn, 82),
            (edit_btn, 86),
            (sent_btn, 70),
            (convert_btn, 128),
            (pdf_actions_btn, 86),
            (purchase_note_btn, 120),
            (approve_btn, 82),
            (reject_btn, 96),
        )
        for button, width in action_widths:
            button.setProperty("compact", "true")
            button.setFixedWidth(width)
            button.setMinimumHeight(29)
            button.setMaximumHeight(31)
            button.setStyleSheet("font-family: 'Segoe UI'; font-size: 9px; font-weight: 700;")
        for button in (back_btn, save_btn, edit_btn, sent_btn, convert_btn, pdf_actions_btn):
            detail_actions_layout.addWidget(button)
        detail_actions_layout.addStretch(1)
        for button in (purchase_note_btn, approve_btn, reject_btn):
            detail_actions_layout.addWidget(button)
        detail_actions.setMaximumHeight(48)
        detail_layout.addWidget(detail_actions)

        self.client_combo = QComboBox()
        self.client_combo.setEditable(True)
        self.client_combo.currentTextChanged.connect(actions._fill_client_from_combo)
        self.executed_combo = QComboBox()
        self.executed_combo.setEditable(True)
        self.workcenter_combo = QComboBox()
        self.workcenter_combo.setEditable(False)
        for combo in (self.client_combo, self.executed_combo):
            self._combo_click_targets[combo] = combo
            combo.installEventFilter(self)
            combo_line_edit = combo.lineEdit()
            if combo_line_edit is not None:
                combo_line_edit.setReadOnly(True)
                combo_line_edit.installEventFilter(self)
                self._combo_click_targets[combo_line_edit] = combo
        self.client_name_edit = QLineEdit()
        self.client_name_edit.textChanged.connect(lambda _text: actions._refresh_quote_identity_label())
        self.client_company_edit = QLineEdit()
        self.client_nif_edit = QLineEdit()
        self.client_address_edit = QLineEdit()
        self.client_contact_edit = QLineEdit()
        self.client_email_edit = QLineEdit()
        self.note_cliente_edit = QLineEdit()
        self.transport_combo = QComboBox()
        self.transport_combo.setEditable(True)
        self.transport_combo.addItems(["", "Transporte a Cargo do Cliente", "Transporte a Nosso Cargo", "Subcontratado"])
        self.transport_carrier_combo = QComboBox()
        self.transport_carrier_combo.setEditable(True)
        self.transport_zone_combo = QComboBox()
        self.transport_zone_combo.setEditable(True)
        self.transport_price_spin = QDoubleSpinBox()
        self.transport_price_spin.setRange(0.0, 1000000.0)
        self.transport_price_spin.setDecimals(2)
        self.transport_price_spin.setSingleStep(5.0)
        self.transport_price_spin.valueChanged.connect(lambda _value: actions._render_quote_lines())
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(0.0, 100.0)
        self.discount_spin.setDecimals(2)
        self.discount_spin.setSingleStep(1.0)
        self.discount_spin.setSuffix(" %")
        self.discount_spin.setToolTip("Desconto global do orçamento, aplicado antes do IVA sobre linhas + transporte.")
        self.discount_spin.valueChanged.connect(lambda _value: actions._render_quote_lines())
        self.price_increment_spin = QDoubleSpinBox()
        self.price_increment_spin.setRange(-100.0, 500.0)
        self.price_increment_spin.setDecimals(2)
        self.price_increment_spin.setSingleStep(1.0)
        self.price_increment_spin.setSuffix(" %")
        self.price_increment_spin.setToolTip("Incremento global aplicado aos preços unitários antes de descontos e IVA.")
        self.price_increment_spin.valueChanged.connect(lambda _value: actions._render_quote_lines())
        self.discount_mode_combo = QComboBox()
        self.discount_mode_combo.addItem("Total final", "total")
        self.discount_mode_combo.addItem("Por lote / espessura", "lotes_espessura")
        self.discount_mode_combo.setToolTip(
            "Escolhe se o desconto se aplica a todas as linhas ou apenas aos lotes/espessuras selecionados."
        )
        self.discount_mode_combo.currentIndexChanged.connect(lambda _index: actions._render_quote_lines())
        self.discount_groups_btn = QPushButton("Escolher lotes")
        self.discount_groups_btn.setProperty("compact", "true")
        self.discount_groups_btn.setProperty("variant", "secondary")
        self.discount_groups_btn.clicked.connect(actions._pick_quote_discount_groups)
        summary_control_width = 132
        summary_control_height = 30
        for summary_control in (
            self.discount_spin,
            self.price_increment_spin,
            self.discount_mode_combo,
            self.discount_groups_btn,
        ):
            summary_control.setMinimumWidth(summary_control_width)
            summary_control.setMaximumWidth(summary_control_width)
            summary_control.setMinimumHeight(summary_control_height)
            summary_control.setMaximumHeight(summary_control_height)
            summary_control.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.iva_spin = QDoubleSpinBox()
        self.iva_spin.setRange(0.0, 100.0)
        self.iva_spin.setDecimals(2)
        self.iva_spin.setValue(23.0)
        self.iva_spin.setEnabled(False)
        self.iva_spin.setToolTip("IVA standard obrigatório para estes orçamentos em Portugal: 23%.")
        self.iva_spin.valueChanged.connect(lambda _value: actions._render_quote_lines())
        self.notes_edit = QTextEdit()
        self.notes_edit.setMinimumHeight(62)
        self.notes_edit.setMaximumHeight(82)
        self.notes_edit.setPlaceholderText("Notas tecnicas e comerciais para o PDF do orcamento.")
        self.delivery_default_text = "A combinar com o departamento de planeamento."
        self.delivery_date_min = QDate.currentDate()
        self.delivery_date_edit = QDateEdit()
        self.delivery_date_edit.setCalendarPopup(True)
        self.delivery_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.delivery_date_edit.setMinimumDate(self.delivery_date_min)
        self.delivery_date_edit.setSpecialValueText("A combinar")
        self.delivery_date_edit.setDate(self.delivery_date_min)
        self.delivery_date_edit.setToolTip("Seleciona uma data apenas quando assumires um prazo concreto com o cliente.")
        self.delivery_default_btn = QPushButton("Prazo Encomenda")
        self.delivery_default_btn.setProperty("variant", "secondary")
        self.delivery_default_btn.setProperty("compact", "true")
        self.delivery_default_btn.setMinimumHeight(30)
        self.delivery_default_btn.setMaximumHeight(32)
        self.delivery_default_btn.setToolTip("Repõe o prazo predefinido: a combinar com o departamento de planeamento.")
        self.delivery_default_btn.clicked.connect(actions._apply_default_delivery_deadline)
        self.transport_km_spin = QDoubleSpinBox()
        self.transport_km_spin.setRange(0.0, 100000.0)
        self.transport_km_spin.setDecimals(1)
        self.transport_km_spin.setSuffix(" km")
        self.transport_rate_spin = QDoubleSpinBox()
        self.transport_rate_spin.setRange(0.0, 1000.0)
        self.transport_rate_spin.setDecimals(3)
        self.transport_rate_spin.setPrefix("EUR ")
        self.transport_rate_spin.setValue(0.65)
        self.transport_diesel_spin = QDoubleSpinBox()
        self.transport_diesel_spin.setRange(0.0, 100.0)
        self.transport_diesel_spin.setDecimals(3)
        self.transport_diesel_spin.setPrefix("EUR ")
        self.transport_diesel_spin.setValue(1.65)
        self.transport_consumption_spin = QDoubleSpinBox()
        self.transport_consumption_spin.setRange(0.0, 100.0)
        self.transport_consumption_spin.setDecimals(2)
        self.transport_consumption_spin.setValue(8.5)
        self.transport_trip_factor_spin = QDoubleSpinBox()
        self.transport_trip_factor_spin.setRange(1.0, 4.0)
        self.transport_trip_factor_spin.setDecimals(1)
        self.transport_trip_factor_spin.setValue(2.0)
        self.transport_suggest_label = QLabel("Transporte sugerido: 0,00 EUR")
        self.transport_suggest_label.setProperty("role", "field_value_strong")
        for widget in (
            self.client_combo,
            self.executed_combo,
            self.workcenter_combo,
            self.client_name_edit,
            self.client_company_edit,
            self.client_nif_edit,
            self.client_contact_edit,
            self.client_email_edit,
            self.client_address_edit,
            self.note_cliente_edit,
        ):
            widget.setProperty("compact", "true")
            widget.setMaximumWidth(16777215)
        for widget in (
            self.transport_combo,
            self.transport_carrier_combo,
            self.transport_zone_combo,
            self.transport_price_spin,
            self.discount_spin,
            self.discount_mode_combo,
            self.iva_spin,
            self.transport_km_spin,
            self.transport_rate_spin,
            self.transport_diesel_spin,
            self.transport_consumption_spin,
            self.transport_trip_factor_spin,
        ):
            widget.setProperty("compact", "true")
            widget.setMaximumWidth(16777215)
        for widget in (
            self.transport_combo,
            self.transport_carrier_combo,
            self.transport_zone_combo,
        ):
            widget.setMinimumWidth(150)
            widget.setMaximumWidth(164)
            widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        for widget in (
            self.transport_price_spin,
            self.iva_spin,
            self.transport_km_spin,
            self.transport_rate_spin,
            self.transport_diesel_spin,
            self.transport_consumption_spin,
            self.transport_trip_factor_spin,
        ):
            widget.setMinimumWidth(116)
            widget.setMaximumWidth(136)
            widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        for widget in (
            self.transport_km_spin,
            self.transport_rate_spin,
            self.transport_diesel_spin,
            self.transport_consumption_spin,
            self.transport_trip_factor_spin,
        ):
            widget.valueChanged.connect(actions._recalc_transport_calc)

        self.quote_header_card = CardFrame()
        self.quote_header_card.set_tone("info")
        meta_layout = QVBoxLayout(self.quote_header_card)
        meta_layout.setContentsMargins(12, 10, 12, 10)
        meta_layout.setSpacing(6)
        header = QHBoxLayout()
        header.setSpacing(10)
        title_block = QVBoxLayout()
        title_block.setSpacing(4)
        self.number_label = QLabel("Novo orcamento")
        self.number_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.link_order_label = QLabel("Sem encomenda gerada")
        self.link_order_label.setProperty("role", "muted")
        title_block.addWidget(self.number_label)
        title_block.addWidget(self.link_order_label)
        total_block = QVBoxLayout()
        total_block.setContentsMargins(0, 0, 4, 0)
        total_block.setSpacing(1)
        total_caption = QLabel("TOTAL DA PROPOSTA")
        total_caption.setProperty("role", "muted")
        total_caption.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        total_caption.setStyleSheet("font-family: 'Segoe UI Semibold'; font-size: 10px; font-weight: 600; letter-spacing: 0.2px;")
        self.header_total_label = QLabel("0,00 EUR")
        self.header_total_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.header_total_label.setStyleSheet("font-family: 'Segoe UI Semibold'; font-size: 20px; font-weight: 600; color: #155e3b;")
        total_block.addWidget(total_caption)
        total_block.addWidget(self.header_total_label)
        self.state_chip = QLabel("-")
        _apply_state_chip(self.state_chip, "-")
        header.addLayout(title_block, 1)
        header.addLayout(total_block)
        header.addWidget(self.state_chip, 0, Qt.AlignTop)
        meta_layout.addLayout(header)
        detail_layout.addWidget(self.quote_header_card)

        self.quote_client_card = CardFrame()
        self.quote_client_card.set_tone("default")
        client_layout = QVBoxLayout(self.quote_client_card)
        client_layout.setContentsMargins(8, 9, 8, 9)
        client_layout.setSpacing(6)
        client_title = QLabel("Cliente e faturação")
        client_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        client_layout.addWidget(client_title)

        self.quote_exec_card = CardFrame()
        self.quote_exec_card.set_tone("default")
        exec_layout = QVBoxLayout(self.quote_exec_card)
        exec_layout.setContentsMargins(8, 9, 8, 9)
        exec_layout.setSpacing(6)
        exec_title = QLabel("Contacto e responsável")
        exec_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        exec_layout.addWidget(exec_title)

        meta_left_host = QWidget()
        meta_left_form = QFormLayout(meta_left_host)
        meta_left_form.setContentsMargins(0, 0, 0, 0)
        meta_left_form.setHorizontalSpacing(6)
        meta_left_form.setVerticalSpacing(5)
        meta_left_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_left_form.addRow("Cliente", self.client_combo)
        meta_left_form.addRow("Nome", self.client_name_edit)
        meta_left_form.addRow("NIF", self.client_nif_edit)
        meta_left_form.addRow("Email", self.client_email_edit)
        meta_right_host = QWidget()
        meta_right_form = QFormLayout(meta_right_host)
        meta_right_form.setContentsMargins(0, 0, 0, 0)
        meta_right_form.setHorizontalSpacing(6)
        meta_right_form.setVerticalSpacing(5)
        meta_right_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_right_form.addRow("Orcamentista", self.executed_combo)
        meta_right_form.addRow("Empresa", self.client_company_edit)
        meta_right_form.addRow("Contacto", self.client_contact_edit)
        meta_right_form.addRow("Morada", self.client_address_edit)
        client_layout.addWidget(meta_left_host, 1)
        exec_layout.addWidget(meta_right_host, 1)
        meta_note_form = QFormLayout()
        meta_note_form.setContentsMargins(0, 0, 0, 0)
        meta_note_form.setHorizontalSpacing(6)
        meta_note_form.setVerticalSpacing(5)
        meta_note_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_note_form.addRow("Nota cliente", self.note_cliente_edit)
        client_layout.addLayout(meta_note_form)
        for proposal_form in (meta_left_form, meta_right_form, meta_note_form):
            proposal_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            for row_index in range(proposal_form.rowCount()):
                label_item = proposal_form.itemAt(row_index, QFormLayout.LabelRole)
                if label_item is not None and label_item.widget() is not None:
                    label_item.widget().setStyleSheet(
                        "font-family: 'Segoe UI'; font-size: 10.5px; font-weight: 700; color: #34423a;"
                    )
        for proposal_field in (
            self.client_combo,
            self.client_name_edit,
            self.client_nif_edit,
            self.client_email_edit,
            self.executed_combo,
            self.client_company_edit,
            self.client_contact_edit,
            self.client_address_edit,
            self.note_cliente_edit,
        ):
            proposal_field.setMinimumWidth(0)
            proposal_field.setMaximumWidth(16777215)
            proposal_field.setMinimumHeight(29)
            proposal_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            proposal_field.setStyleSheet("font-family: 'Segoe UI'; font-size: 11px; padding: 3px 6px;")

        self.quote_notes_card = CardFrame()
        self.quote_notes_card.set_tone("warning")
        self.quote_notes_card.setStyleSheet(
            STYLE_QUOTE_NOTES_CARD
        )
        notes_layout = QVBoxLayout(self.quote_notes_card)
        notes_layout.setContentsMargins(12, 11, 12, 12)
        notes_layout.setSpacing(8)
        notes_header = QHBoxLayout()
        notes_title = QLabel("Notas do orcamento (PDF)")
        notes_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #263b2d;")
        notes_hint = QLabel("PDF / transporte")
        notes_hint.setProperty("role", "muted")
        notes_hint.setStyleSheet("font-size: 9.5px; color: #737d75;")
        notes_header.addWidget(notes_title)
        notes_header.addStretch(1)
        notes_header.addWidget(notes_hint)
        notes_layout.addLayout(notes_header)

        self.notes_tabs = QTabWidget()
        self.notes_tabs.setDocumentMode(True)
        self.notes_tabs.setStyleSheet(
            STYLE_NOTES_TABS
        )
        self.notes_tabs.tabBar().setExpanding(True)
        self.notes_tabs.tabBar().setUsesScrollButtons(False)
        self.notes_tabs.tabBar().setElideMode(Qt.ElideNone)

        transport_page = QWidget()
        transport_page_layout = QVBoxLayout(transport_page)
        transport_page_layout.setContentsMargins(0, 0, 0, 0)
        transport_page_layout.setSpacing(8)
        transport_intro = QLabel("Define o transporte do orçamento, calcula a sugestão e aplica diretamente ao PDF e ao total final.")
        transport_intro.setProperty("role", "muted")
        transport_intro.setWordWrap(True)
        transport_intro.setStyleSheet("font-size: 9.5px; color: #657068;")
        transport_page_layout.addWidget(transport_intro)

        transport_scroll = QScrollArea()
        transport_scroll.setWidgetResizable(True)
        transport_scroll.setFrameShape(QFrame.NoFrame)
        transport_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        transport_scroll.setMinimumHeight(320)
        transport_scroll.setStyleSheet("QScrollArea { background: transparent; border: 0; }")
        transport_scroll_content = QWidget()
        transport_stack = QVBoxLayout()
        transport_scroll_content.setLayout(transport_stack)
        transport_stack.setContentsMargins(0, 0, 6, 8)
        transport_stack.setSpacing(8)

        transport_form_card = CardFrame()
        transport_form_card.set_tone("default")
        transport_form_card.setStyleSheet(
            STYLE_TRANSPORT_FORM_CARD
        )
        transport_form_layout = QGridLayout(transport_form_card)
        transport_form_layout.setContentsMargins(11, 10, 11, 11)
        transport_form_layout.setHorizontalSpacing(9)
        transport_form_layout.setVerticalSpacing(7)
        transport_form_layout.setColumnStretch(0, 1)
        transport_form_layout.setColumnStretch(1, 1)

        def _transport_field_row(text: str, widget: QWidget) -> QWidget:
            host = QWidget()
            host.setMinimumHeight(43)
            host.setMaximumHeight(46)
            host.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            host_layout = QVBoxLayout(host)
            host_layout.setContentsMargins(0, 0, 0, 0)
            host_layout.setSpacing(2)
            label = QLabel(text)
            label.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
            host_layout.addWidget(label)
            field_box = QFrame()
            field_box.setObjectName("TransportFieldBox")
            field_box.setMinimumHeight(29)
            field_box.setMaximumHeight(31)
            field_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            field_layout = QHBoxLayout(field_box)
            field_layout.setContentsMargins(1, 1, 1, 1)
            field_layout.setSpacing(0)
            widget.setMinimumHeight(25)
            widget.setMaximumHeight(28)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            field_layout.addWidget(widget)
            host_layout.addWidget(field_box)
            return host

        for index, (label_text, widget) in enumerate((
            ("Modo", self.transport_combo),
            ("Transportadora", self.transport_carrier_combo),
            ("Zona", self.transport_zone_combo),
            ("Distância", self.transport_km_spin),
            ("Preço / km", self.transport_rate_spin),
            ("Gasóleo", self.transport_diesel_spin),
            ("Consumo", self.transport_consumption_spin),
            ("Fator viagem", self.transport_trip_factor_spin),
            ("Preço manual", self.transport_price_spin),
            ("IVA %", self.iva_spin),
        )):
            transport_form_layout.addWidget(_transport_field_row(label_text, widget), index // 2, index % 2)

        transport_actions_card = CardFrame()
        transport_actions_card.set_tone("default")
        transport_actions_card.setStyleSheet(
            STYLE_TRANSPORT_ACTIONS_CARD
        )
        transport_actions_layout = QGridLayout(transport_actions_card)
        transport_actions_layout.setContentsMargins(11, 10, 11, 11)
        transport_actions_layout.setHorizontalSpacing(8)
        transport_actions_layout.setVerticalSpacing(7)
        transport_actions_title = QLabel("Cálculo e aplicação")
        transport_actions_title.setStyleSheet("font-size: 10.5px; font-weight: 800; color: #2f4633;")
        transport_actions_hint = QLabel("Usa os parâmetros do transporte para sugerir um valor coerente com a orçamentação.")
        transport_actions_hint.setProperty("role", "muted")
        transport_actions_hint.setWordWrap(True)
        transport_actions_hint.setStyleSheet("font-size: 9px; color: #657068;")
        transport_actions_text = QVBoxLayout()
        transport_actions_text.setContentsMargins(0, 0, 0, 0)
        transport_actions_text.setSpacing(4)
        transport_actions_text.addWidget(transport_actions_title)
        transport_actions_text.addWidget(transport_actions_hint)
        fill_notes_btn = QPushButton("Preencher notas PDF")
        fill_notes_btn.setProperty("variant", "secondary")
        fill_notes_btn.setProperty("compact", "true")
        fill_notes_btn.setProperty("quoteTransportAction", "true")
        fill_notes_btn.setToolTip("Preenche automaticamente as notas PDF com base no contexto do orçamento.")
        fill_notes_btn.clicked.connect(actions._fill_pdf_notes_from_context)
        apply_transport_btn = QPushButton("Aplicar ao orçamento")
        apply_transport_btn.setProperty("variant", "secondary")
        apply_transport_btn.setProperty("compact", "true")
        apply_transport_btn.setProperty("quoteTransportAction", "true")
        apply_transport_btn.setToolTip("Aplica o cálculo do transporte ao orçamento atual.")
        apply_transport_btn.clicked.connect(actions._apply_transport_calc)
        clear_transport_btn = QPushButton("Remover transporte")
        clear_transport_btn.setProperty("variant", "secondary")
        clear_transport_btn.setProperty("compact", "true")
        clear_transport_btn.setProperty("quoteTransportAction", "true")
        clear_transport_btn.setToolTip("Limpa o modo, transportadora, zona e valor de transporte deste orçamento.")
        clear_transport_btn.clicked.connect(actions._clear_transport)
        for button in (fill_notes_btn, apply_transport_btn, clear_transport_btn):
            button.setMinimumHeight(32)
            button.setMaximumHeight(34)
            button.setMinimumWidth(0)
            button.setMaximumWidth(16777215)
            button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            button.setStyleSheet(
                "QPushButton { border: 1px solid #aebba9; border-bottom: 1px solid #7f9279; border-radius: 7px; padding: 0 8px; font-size: 9.5px; font-weight: 700; background: #f7f9f5; color: #2f3c33; }"
                " QPushButton:hover { border-color: #6f9f45; background: #ffffff; color: #315223; }"
            )
        self.transport_suggest_label.setStyleSheet("font-size: 12px; font-weight: 900; color: #315d2a;")
        self.transport_suggest_label.setAlignment(Qt.AlignCenter)
        transport_action_buttons = QGridLayout()
        transport_action_buttons.setContentsMargins(0, 0, 0, 0)
        transport_action_buttons.setHorizontalSpacing(8)
        transport_action_buttons.setSpacing(8)
        transport_action_buttons.addWidget(fill_notes_btn, 0, 0)
        transport_action_buttons.addWidget(apply_transport_btn, 0, 1)
        transport_action_buttons.addWidget(clear_transport_btn, 1, 0, 1, 2)
        transport_action_buttons.setColumnStretch(0, 1)
        transport_action_buttons.setColumnStretch(1, 1)
        transport_actions_layout.addLayout(transport_actions_text, 0, 0, 1, 2)
        transport_actions_layout.addWidget(self.transport_suggest_label, 1, 0, 1, 2)
        transport_actions_layout.addLayout(transport_action_buttons, 2, 0, 1, 2)
        transport_actions_layout.setColumnStretch(0, 1)
        transport_actions_layout.setColumnStretch(1, 1)

        transport_stack.addWidget(transport_form_card)
        transport_stack.addWidget(transport_actions_card)
        transport_stack.addStretch(1)
        transport_scroll.setWidget(transport_scroll_content)
        transport_page_layout.addWidget(transport_scroll, 1)

        operations_page = QWidget()
        operations_layout = QVBoxLayout(operations_page)
        operations_layout.setContentsMargins(0, 0, 0, 0)
        operations_layout.setSpacing(10)
        operations_intro = QLabel("Insere rapidamente notas técnicas e comerciais no PDF conforme o processo considerado.")
        operations_intro.setProperty("role", "muted")
        operations_intro.setWordWrap(True)
        operations_intro.setStyleSheet("font-size: 10px; color: #657068;")
        operations_layout.addWidget(operations_intro)
        op_card = CardFrame()
        op_card.set_tone("default")
        op_card.setStyleSheet(
            "QFrame#Card { background: #f7f9f6; border: 1px solid #d5ddd2; border-radius: 8px; }"
            "QFrame#Card QLabel { background: transparent; border: 0; }"
            "QFrame#Card QPushButton { min-height: 38px; background: #ffffff; color: #314239; border: 1px solid #b9c7b4; border-bottom: 1px solid #83967d; border-radius: 7px; font-size: 10px; font-weight: 750; }"
            "QFrame#Card QPushButton:hover { background: #eef5e9; border-color: #6f9f45; color: #315d2a; }"
            "QFrame#Card QPushButton:pressed { background: #dfecd8; }"
        )
        op_card_layout = QVBoxLayout(op_card)
        op_card_layout.setContentsMargins(12, 12, 12, 12)
        op_card_layout.setSpacing(9)
        op_title = QLabel("Inserções rápidas por operação")
        op_title.setStyleSheet("font-size: 11px; font-weight: 800; color: #2f4633;")
        op_card_layout.addWidget(op_title)
        op_hint = QLabel("Seleciona uma operação para acrescentar uma nota normalizada ao documento.")
        op_hint.setWordWrap(True)
        op_hint.setStyleSheet("font-size: 9.5px; color: #657068;")
        op_card_layout.addWidget(op_hint)
        op_buttons = QGridLayout()
        op_buttons.setHorizontalSpacing(8)
        op_buttons.setVerticalSpacing(8)
        for index, op_text in enumerate(("Corte Laser", "Quinagem", "Roscagem", "Furo Manual", "Soldadura")):
            op_btn = QPushButton(op_text)
            op_btn.setProperty("variant", "secondary")
            op_btn.clicked.connect(lambda _checked=False, text=op_text: actions._append_pdf_note(f"- Foi considerado: {text}."))
            op_btn.setMinimumHeight(38)
            op_buttons.addWidget(op_btn, index // 2, index % 2)
        op_buttons.setColumnStretch(0, 1)
        op_buttons.setColumnStretch(1, 1)
        op_card_layout.addLayout(op_buttons)
        op_card_layout.addStretch(1)
        operations_layout.addWidget(op_card, 1)

        notes_text_page = QWidget()
        notes_text_page_layout = QVBoxLayout(notes_text_page)
        notes_text_page_layout.setContentsMargins(0, 0, 0, 0)
        notes_text_page_layout.setSpacing(8)
        notes_text_intro = QLabel("Texto livre que segue diretamente para o PDF do orçamento.")
        notes_text_intro.setProperty("role", "muted")
        notes_text_intro.setWordWrap(True)
        notes_text_intro.setStyleSheet("font-size: 10px; color: #657068;")
        notes_text_page_layout.addWidget(notes_text_intro)
        notes_text_card = CardFrame()
        notes_text_card.set_tone("default")
        notes_text_card.setStyleSheet(
            "QFrame#Card { background: #f7f9f6; border: 1px solid #d5ddd2; border-radius: 8px; }"
            "QFrame#Card QLabel { background: transparent; border: 0; }"
            "QFrame#Card QTextEdit { background: #ffffff; color: #2c3830; border: 1px solid #bdc9b9; border-radius: 7px; padding: 8px; font-size: 10.5px; }"
            "QFrame#Card QTextEdit:focus { border-color: #6f9f45; }"
        )
        notes_text_layout = QVBoxLayout(notes_text_card)
        notes_text_layout.setContentsMargins(12, 11, 12, 12)
        notes_text_layout.setSpacing(8)
        notes_text_title = QLabel("Texto técnico e comercial")
        notes_text_title.setStyleSheet("font-size: 11px; font-weight: 800; color: #2f4633;")
        notes_text_layout.addWidget(notes_text_title)
        delivery_row = QHBoxLayout()
        delivery_row.setContentsMargins(0, 0, 0, 0)
        delivery_row.setSpacing(8)
        delivery_label = QLabel("Prazo")
        delivery_label.setProperty("role", "field_label")
        delivery_label.setStyleSheet("font-size: 10px; font-weight: 800; color: #4d5b51;")
        delivery_row.addWidget(delivery_label)
        delivery_row.addWidget(self.delivery_default_btn)
        self.delivery_date_edit.setMinimumHeight(30)
        self.delivery_date_edit.setMaximumHeight(32)
        delivery_row.addWidget(self.delivery_date_edit, 1)
        notes_text_layout.addLayout(delivery_row)
        self.notes_edit.setProperty("compact", "false")
        self.notes_edit.setMinimumHeight(220)
        self.notes_edit.setMaximumHeight(16777215)
        notes_text_layout.addWidget(self.notes_edit, 1)
        notes_text_page_layout.addWidget(notes_text_card, 1)

        self.notes_tabs.addTab(transport_page, "Transporte")
        self.notes_tabs.addTab(operations_page, "Operações")
        self.notes_tabs.addTab(notes_text_page, "Texto PDF")
        notes_layout.addWidget(self.notes_tabs, 1)
        transport_widgets = (
            self.transport_combo,
            self.transport_carrier_combo,
            self.transport_zone_combo,
            self.transport_price_spin,
            self.transport_diesel_spin,
            self.iva_spin,
            self.transport_km_spin,
            self.transport_rate_spin,
            self.transport_consumption_spin,
            self.transport_trip_factor_spin,
        )
        for widget in transport_widgets:
            widget.setProperty("compact", "false")
            widget.setMaximumWidth(16777215)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            widget.setMinimumWidth(112)
            widget.setMinimumHeight(22)
            widget.setMaximumHeight(26)
            _repolish(widget)
        for combo in (self.transport_combo, self.transport_carrier_combo, self.transport_zone_combo):
            self._combo_click_targets[combo] = combo
            combo.installEventFilter(self)
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.setReadOnly(True)
                line_edit.installEventFilter(self)
                self._combo_click_targets[line_edit] = combo
        self.notes_tabs.setMinimumHeight(390)
        self.quote_notes_card.setMinimumHeight(450)
        self.quote_notes_card.setMaximumHeight(16777215)
        self.quote_summary_card = CardFrame()
        self.quote_summary_card.set_tone("default")
        self.quote_summary_card.setMaximumWidth(16777215)
        self.quote_summary_card.setMinimumWidth(0)
        summary_layout = QVBoxLayout(self.quote_summary_card)
        summary_layout.setContentsMargins(8, 7, 8, 10)
        summary_layout.setSpacing(9)

        finance_header = QHBoxLayout()
        finance_header.setContentsMargins(2, 0, 2, 2)
        finance_header.setSpacing(8)
        finance_heading = QVBoxLayout()
        finance_heading.setContentsMargins(0, 0, 0, 0)
        finance_heading.setSpacing(2)
        summary_title = QLabel("Resumo financeiro")
        summary_title.setStyleSheet("font-size: 15px; font-weight: 900; color: #23372a;")
        summary_hint = QLabel("Composição comercial atualizada em tempo real.")
        summary_hint.setProperty("role", "muted")
        summary_hint.setStyleSheet("font-family: 'Segoe UI'; font-size: 10px; color: #6c766e;")
        self.quote_finance_status_label = QLabel("CÁLCULO AUTOMÁTICO")
        self.quote_finance_status_label.setAlignment(Qt.AlignCenter)
        self.quote_finance_status_label.setStyleSheet(
            "background: #edf5e8; border: 1px solid #bfd5b4; border-radius: 6px;"
            " color: #426a31; font-size: 9px; font-weight: 900; padding: 4px 8px;"
        )
        finance_heading.addWidget(summary_title)
        finance_heading.addWidget(summary_hint)
        finance_header.addLayout(finance_heading, 1)
        finance_header.addWidget(self.quote_finance_status_label, 0, Qt.AlignTop)
        summary_layout.addLayout(finance_header)

        self.lines_subtotal_label = QLabel("0,00 EUR")
        self.increment_value_label = QLabel("0,00 EUR")
        self.transport_total_label = QLabel("0,00 EUR")
        self.discount_value_label = QLabel("0,00 EUR")
        self.subtotal_without_iva_label = QLabel("0,00 EUR")
        self.iva_total_label = QLabel("0,00 EUR")
        self.total_label = QLabel("0,00 EUR")
        for widget in (
            self.lines_subtotal_label,
            self.increment_value_label,
            self.transport_total_label,
            self.discount_value_label,
            self.subtotal_without_iva_label,
            self.iva_total_label,
        ):
            widget.setMinimumWidth(0)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            widget.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            widget.setStyleSheet("font-family: 'Segoe UI'; font-size: 10.5px; font-weight: 800; color: #34423a;")
        self.total_label.setProperty("role", "field_value_strong")
        self.total_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_label.setMinimumWidth(0)
        self.total_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.total_label.setStyleSheet("font-size: 18px; font-weight: 900; color: #20352a;")

        total_panel = QFrame()
        total_panel.setObjectName("QuoteSummaryTotalPanel")
        total_panel.setStyleSheet(
            STYLE_TOTAL_PANEL
        )
        total_panel_layout = QHBoxLayout(total_panel)
        total_panel_layout.setContentsMargins(12, 10, 12, 10)
        total_panel_layout.setSpacing(8)
        total_caption_host = QVBoxLayout()
        total_caption_host.setContentsMargins(0, 0, 0, 0)
        total_caption_host.setSpacing(1)
        self.total_caption_label = QLabel("Total C/IVA 23%")
        self.total_caption_label.setStyleSheet("font-size: 10.5px; font-weight: 900; color: #3f5f36;")
        self.total_caption_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        total_caption_hint = QLabel("VALOR DA PROPOSTA")
        total_caption_hint.setStyleSheet("font-size: 9px; font-weight: 800; color: #778078;")
        total_caption_host.addWidget(total_caption_hint)
        total_caption_host.addWidget(self.total_caption_label)
        total_panel_layout.addLayout(total_caption_host, 1)
        total_panel_layout.addWidget(self.total_label, 1)
        summary_layout.addWidget(total_panel)

        finance_kpis = QGridLayout()
        finance_kpis.setContentsMargins(0, 0, 0, 0)
        finance_kpis.setHorizontalSpacing(6)
        finance_kpis.setVerticalSpacing(0)

        def _finance_kpi(title: str, value: QLabel, accent: str) -> QFrame:
            card = QFrame()
            card.setObjectName("QuoteFinanceKpi")
            card.setStyleSheet(
                "QFrame#QuoteFinanceKpi {"
                " background: #ffffff;"
                " border: 1px solid #d4ddd1;"
                f" border-top: 3px solid {accent};"
                " border-radius: 6px;"
                "}"
                " QFrame#QuoteFinanceKpi QLabel { background: transparent; border: 0; }"
            )
            layout = QVBoxLayout(card)
            layout.setContentsMargins(9, 7, 9, 8)
            layout.setSpacing(2)
            caption = QLabel(title)
            caption.setStyleSheet("font-size: 9px; font-weight: 800; color: #6b756d;")
            value.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            value.setStyleSheet("font-size: 13px; font-weight: 900; color: #263b2d;")
            layout.addWidget(caption)
            layout.addWidget(value)
            return card

        finance_kpis.addWidget(_finance_kpi("BASE TRIBUTÁVEL", self.subtotal_without_iva_label, "#6f9f45"), 0, 0)
        finance_kpis.addWidget(_finance_kpi("IVA 23%", self.iva_total_label, "#8a9687"), 0, 1)
        finance_kpis.setColumnStretch(0, 1)
        finance_kpis.setColumnStretch(1, 1)
        summary_layout.addLayout(finance_kpis)

        summary_rows_host = QFrame()
        summary_rows_host.setObjectName("QuoteFinancialBreakdown")
        summary_rows_host.setStyleSheet(
            STYLE_SUMMARY_ROWS_HOST
        )
        summary_rows_layout = QGridLayout(summary_rows_host)
        summary_rows_layout.setContentsMargins(10, 7, 10, 7)
        summary_rows_layout.setHorizontalSpacing(12)
        summary_rows_layout.setVerticalSpacing(4)
        self.iva_summary_row_label = None
        for row_index, (label_text, widget) in enumerate(
            (
                ("Valor das linhas", self.lines_subtotal_label),
                ("Incremento", self.increment_value_label),
                ("Transporte", self.transport_total_label),
                ("Desconto aplicado", self.discount_value_label),
            )
        ):
            label = QLabel(label_text)
            label.setProperty("role", "field_label")
            label.setWordWrap(False)
            label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            label.setMinimumHeight(26)
            label.setStyleSheet(
                "font-size: 10px; font-weight: 700; color: #526057; background: transparent; border: 0;"
            )
            widget.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            widget.setMinimumHeight(26)
            value_color = "#b45f06" if widget is self.discount_value_label else "#30343b"
            widget.setStyleSheet(f"font-family: 'Segoe UI'; font-size: 10.5px; font-weight: 800; color: {value_color}; background: transparent; border: 0;")
            summary_rows_layout.addWidget(label, row_index * 2, 0)
            summary_rows_layout.addWidget(widget, row_index * 2, 1)
            if row_index < 3:
                divider = QFrame()
                divider.setObjectName("QuoteSummaryDivider")
                divider.setFrameShape(QFrame.HLine)
                divider.setFixedHeight(1)
                divider.setStyleSheet("background: #edf0eb; border: 0;")
                summary_rows_layout.addWidget(divider, row_index * 2 + 1, 0, 1, 2)
        summary_rows_layout.setColumnStretch(0, 1)
        summary_rows_layout.setColumnStretch(1, 0)
        summary_layout.addWidget(summary_rows_host)

        controls_panel = QFrame()
        controls_panel.setObjectName("QuoteFinancialControls")
        controls_panel.setStyleSheet(
            STYLE_CONTROLS_PANEL
        )
        controls_layout = QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(10, 8, 10, 9)
        controls_layout.setSpacing(6)
        controls_title = QLabel("Ajuste comercial")
        controls_title.setStyleSheet("font-size: 10.5px; font-weight: 900; color: #3f5f36;")
        controls_layout.addWidget(controls_title)

        def _summary_control_label(text: str) -> QLabel:
            label = QLabel(text)
            label.setProperty("role", "field_label")
            label.setMinimumHeight(30)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setStyleSheet("font-family: 'Segoe UI'; font-size: 10px; font-weight: 800; color: #46554b;")
            return label

        def _summary_control_row(label: QLabel, widget: QWidget) -> QHBoxLayout:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(10)
            row.addWidget(label)
            row.addStretch(1)
            row.addWidget(widget)
            return row

        for summary_control in (
            self.discount_spin,
            self.price_increment_spin,
            self.discount_mode_combo,
            self.discount_groups_btn,
        ):
            summary_control.setMinimumHeight(30)
            summary_control.setMaximumHeight(34)
            summary_control.setStyleSheet(
                STYLE_SUMMARY_CONTROL
            )
        discount_row = _summary_control_row(_summary_control_label("Desconto global"), self.discount_spin)
        controls_layout.addLayout(discount_row)
        increment_row = _summary_control_row(_summary_control_label("Incremento preços"), self.price_increment_spin)
        controls_layout.addLayout(increment_row)
        discount_mode_row = _summary_control_row(_summary_control_label("Aplicação"), self.discount_mode_combo)
        controls_layout.addLayout(discount_mode_row)
        discount_group_row = _summary_control_row(_summary_control_label("Lotes elegíveis"), self.discount_groups_btn)
        controls_layout.addLayout(discount_group_row)
        summary_layout.addWidget(controls_panel)

        discount_status = QFrame()
        discount_status.setObjectName("QuoteDiscountStatus")
        discount_status.setStyleSheet(
            STYLE_DISCOUNT_STATUS
        )
        discount_status_layout = QVBoxLayout(discount_status)
        discount_status_layout.setContentsMargins(8, 6, 8, 6)
        discount_status_layout.setSpacing(3)
        self.discount_target_label = QLabel("Desconto aplicado a todas as linhas.")
        self.discount_target_label.setProperty("role", "muted")
        self.discount_target_label.setWordWrap(True)
        self.discount_target_label.setStyleSheet("font-size: 9px; font-weight: 700; color: #526057;")
        discount_status_layout.addWidget(self.discount_target_label)
        self.discount_breakdown_label = QLabel("Sem desconto global aplicado.")
        self.discount_breakdown_label.setProperty("role", "muted")
        self.discount_breakdown_label.setWordWrap(True)
        self.discount_breakdown_label.setStyleSheet("font-size: 9px; color: #6b756d;")
        discount_status_layout.addWidget(self.discount_breakdown_label)
        summary_layout.addWidget(discount_status)
        summary_layout.addStretch(1)
        lines_card = CardFrame()
        lines_card.set_tone("default")
        lines_card.setObjectName("QuoteReferencesCard")
        lines_card.setStyleSheet(
            "QFrame#QuoteReferencesCard { background: #ffffff; border: 1px solid #d5ddd3; border-radius: 8px; }"
        )
        lines_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.quote_lines_card = lines_card
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(14, 12, 14, 12)
        lines_layout.setSpacing(7)
        line_actions = QVBoxLayout()
        line_actions.setSpacing(4)
        line_title_row = QHBoxLayout()
        line_title_row.setSpacing(8)
        lines_title = QLabel("Referências do orçamento")
        lines_title.setStyleSheet("font-family: 'Segoe UI Semibold'; font-size: 14px; font-weight: 600; color: #172b3f;")
        add_line_btn = QPushButton("Adicionar linha")
        add_line_btn.clicked.connect(actions._add_line)
        laser_batch_btn = QPushButton("Lote DXF/DWG")
        laser_batch_btn.setProperty("variant", "secondary")
        laser_batch_btn.clicked.connect(actions._add_laser_batch_lines)
        laser_batch_btn.setToolTip("Permite carregar um único desenho ou um lote completo de DXF/DWG.")
        laser_nesting_btn = QPushButton("Nesting")
        laser_nesting_btn.setProperty("variant", "secondary")
        laser_nesting_btn.clicked.connect(actions._open_laser_nesting)
        profile_step_btn = QPushButton("STEP/IGS Perfil")
        profile_step_btn.setProperty("variant", "secondary")
        profile_step_btn.clicked.connect(actions._open_profile_step_igs_quote_builder)
        add_model_btn = QPushButton("Conjunto/Modelo")
        add_model_btn.setProperty("variant", "secondary")
        add_model_btn.clicked.connect(actions._manage_saved_conjuntos)
        conjunto_builder_btn = QPushButton("Conjunto calculado")
        conjunto_builder_btn.setProperty("variant", "secondary")
        conjunto_builder_btn.clicked.connect(actions._open_calculated_assembly_builder)
        save_selected_group_btn = QPushButton("Guardar conjunto/modelo")
        save_selected_group_btn.setProperty("variant", "secondary")
        save_selected_group_btn.clicked.connect(actions._save_selected_lines_as_group)
        structure_builder_btn = QPushButton("Estrutura metalica")
        structure_builder_btn.setProperty("variant", "secondary")
        structure_builder_btn.clicked.connect(actions._open_structure_quote_builder)
        manage_models_btn = QPushButton("Modelos")
        manage_models_btn.setProperty("variant", "secondary")
        manage_models_btn.clicked.connect(actions._manage_assembly_models)
        laser_cfg_btn = QPushButton("Config. Laser")
        laser_cfg_btn.setProperty("variant", "secondary")
        laser_cfg_btn.clicked.connect(actions._configure_laser_profiles)
        operation_cfg_btn = QPushButton("Config. Operacoes")
        operation_cfg_btn.setProperty("variant", "secondary")
        operation_cfg_btn.clicked.connect(actions._configure_operation_profiles)
        edit_line_btn = QPushButton("Editar linha")
        edit_line_btn.setProperty("variant", "secondary")
        edit_line_btn.clicked.connect(actions._edit_line)
        prepare_line_btn = QPushButton("Preparar produção")
        prepare_line_btn.setProperty("variant", "secondary")
        prepare_line_btn.clicked.connect(actions._prepare_selected_line_for_production)
        check_weight_btn = QPushButton("Verificar peso")
        check_weight_btn.setProperty("variant", "secondary")
        check_weight_btn.clicked.connect(actions._check_selected_line_weight)
        remove_line_btn = QPushButton("Remover linha")
        remove_line_btn.setProperty("variant", "danger")
        remove_line_btn.clicked.connect(actions._remove_line)
        remove_line_btn.setToolTip("Remove as linhas marcadas; se nenhuma estiver marcada, remove a selecao atual.")
        self.remove_quote_lines_btn = remove_line_btn
        open_draw_btn = QPushButton("Ver desenho")
        open_draw_btn.setProperty("variant", "secondary")
        open_draw_btn.clicked.connect(actions._open_line_drawing)
        laser_nesting_btn.setToolTip("Abrir o estudo de nesting das linhas laser do orçamento.")
        line_title_row.addWidget(lines_title)
        line_title_row.addStretch(1)
        remove_line_btn.setMinimumWidth(132)
        remove_line_btn.setMinimumHeight(28)
        line_title_row.addWidget(remove_line_btn)
        self.line_count_label = QLabel("0 linhas")
        self.line_count_label.setProperty("role", "state_chip")
        self.line_count_label.setAlignment(Qt.AlignCenter)
        self.line_count_label.setMinimumWidth(86)
        line_title_row.addWidget(self.line_count_label)
        line_actions.addLayout(line_title_row)
        self.nesting_bridge_label = QLabel(
            "Preco unitario da tabela = orcamentacao por peca. Nesting = validacao global de materia, stock/retalho e pecas realmente programadas."
        )
        self.nesting_bridge_label.setProperty("role", "muted")
        self.nesting_bridge_label.setWordWrap(True)
        self.nesting_bridge_label.setStyleSheet("font-size: 9.5px; line-height: 1.15;")
        self.nesting_bridge_label.setMaximumHeight(26)
        line_actions.addWidget(self.nesting_bridge_label)
        line_tools_tabs = QTabWidget()
        line_tools_tabs.setDocumentMode(True)
        line_tools_tabs.setMaximumHeight(132)
        line_tools_tabs.setStyleSheet(
            STYLE_LINE_TOOLS_TABS
        )

        def _quote_tool_tab(buttons: tuple[QPushButton, ...], columns: int) -> QWidget:
            tab = QWidget()
            tab_layout = QGridLayout(tab)
            tab_layout.setContentsMargins(7, 5, 7, 5)
            tab_layout.setHorizontalSpacing(6)
            tab_layout.setVerticalSpacing(5)
            column_count = max(1, int(columns))
            for index, button in enumerate(buttons):
                button.setProperty("compact", "true")
                button.setMinimumWidth(0)
                button.setMinimumHeight(28)
                button.setMaximumHeight(30)
                button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                tab_layout.addWidget(button, index // column_count, index % column_count)
            for column_index in range(column_count):
                tab_layout.setColumnStretch(column_index, 1)
            return tab

        line_tools_tabs.addTab(
            _quote_tool_tab(
                (
                    add_line_btn,
                    laser_batch_btn,
                    laser_nesting_btn,
                    profile_step_btn,
                    add_model_btn,
                    conjunto_builder_btn,
                    structure_builder_btn,
                ),
                3,
            ),
            "Adicionar e calcular",
        )
        line_tools_tabs.addTab(
            _quote_tool_tab((save_selected_group_btn, manage_models_btn, laser_cfg_btn, operation_cfg_btn), 4),
            "Biblioteca e parâmetros",
        )
        line_actions.addWidget(line_tools_tabs)
        for button in (
            add_line_btn,
            laser_batch_btn,
            laser_nesting_btn,
            profile_step_btn,
            add_model_btn,
            conjunto_builder_btn,
            save_selected_group_btn,
            structure_builder_btn,
            manage_models_btn,
            laser_cfg_btn,
            operation_cfg_btn,
            edit_line_btn,
            prepare_line_btn,
            remove_line_btn,
            open_draw_btn,
            check_weight_btn,
        ):
            button.setProperty("compact", "true")
        self.lines_table = QTableWidget(0, 14)
        self.lines_table.setHorizontalHeaderLabels(["Apagar", "Tipo", "Ref./Cód.", "Ref. ext.", "Descrição", "Material/Produto", "Esp./Unid.", "Operação", "Tempo", "Qtd.", "Preço", "Preço c/ desc.", "Total", "Conjunto"])
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.lines_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.lines_table.setAlternatingRowColors(True)
        self.lines_table.setStyleSheet(
            STYLE_LINES_TABLE
        )
        self.lines_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.lines_table.horizontalHeader().setSectionsClickable(True)
        self.lines_table.horizontalHeader().setSortIndicatorShown(False)
        for column in range(self.lines_table.columnCount()):
            header_item = self.lines_table.horizontalHeaderItem(column)
            if header_item is not None:
                header_item.setToolTip(
                    "Clique para marcar ou desmarcar todas as linhas."
                    if column == self.LINE_COL_MARK
                    else "Clique para ordenar; clique novamente para inverter."
                )
        self.lines_table.horizontalHeader().sectionClicked.connect(actions._handle_quote_lines_sort)
        self.lines_table.verticalHeader().setDefaultSectionSize(30)
        self.lines_table.verticalHeader().setMinimumSectionSize(30)
        self.lines_table.horizontalHeader().setMinimumHeight(34)
        self.lines_table.setMinimumHeight(280)
        self.lines_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lines_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        _set_table_columns(
            self.lines_table,
            [
                (self.LINE_COL_MARK, "fixed", 58),
                (self.LINE_COL_TYPE, "fixed", 92),
                (self.LINE_COL_REFERENCE, "fixed", 104),
                (self.LINE_COL_EXTERNAL_REFERENCE, "fixed", 132),
                (self.LINE_COL_DESCRIPTION, "stretch", 0),
                (self.LINE_COL_MATERIAL, "stretch", 0),
                (self.LINE_COL_UNIT, "fixed", 72),
                (self.LINE_COL_OPERATION, "fixed", 104),
                (self.LINE_COL_TIME, "fixed", 68),
                (self.LINE_COL_QUANTITY, "fixed", 58),
                (self.LINE_COL_PRICE, "fixed", 76),
                (self.LINE_COL_DISCOUNTED_PRICE, "fixed", 88),
                (self.LINE_COL_TOTAL, "fixed", 84),
                (self.LINE_COL_ASSEMBLY, "fixed", 84),
            ],
        )
        line_actions_host = QWidget()
        line_actions_host.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        line_actions_host_layout = QVBoxLayout(line_actions_host)
        line_actions_host_layout.setContentsMargins(0, 0, 0, 0)
        line_actions_host_layout.setSpacing(4)
        line_actions_host_layout.addLayout(line_actions)
        lines_layout.addWidget(line_actions_host, 0, Qt.AlignTop)
        lines_layout.addWidget(self.lines_table, 1)
        selected_line_footer = QFrame()
        selected_line_footer.setObjectName("QuoteSelectedLineFooter")
        selected_line_footer.setStyleSheet(
            STYLE_SELECTED_LINE_FOOTER
        )
        selected_line_footer.setMinimumHeight(43)
        selected_line_footer.setMaximumHeight(47)
        selected_line_footer_layout = QHBoxLayout(selected_line_footer)
        selected_line_footer_layout.setContentsMargins(8, 6, 8, 6)
        selected_line_footer_layout.setSpacing(6)
        selected_line_caption = QLabel("LINHA SELECIONADA")
        selected_line_caption.setObjectName("QuoteSelectedLineCaption")
        self.quote_selected_line_caption = selected_line_caption
        selected_line_footer_layout.addWidget(selected_line_caption)
        selected_line_footer_layout.addStretch(1)
        self.quote_selected_line_buttons = (
            edit_line_btn,
            open_draw_btn,
            prepare_line_btn,
            check_weight_btn,
        )
        for selected_action in self.quote_selected_line_buttons:
            selected_action.setMinimumWidth(118)
            selected_action.setMaximumWidth(178)
            selected_action.setMinimumHeight(29)
            selected_action.setMaximumHeight(31)
            selected_action.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            selected_line_footer_layout.addWidget(selected_action, 1)
        lines_layout.addWidget(selected_line_footer, 0)
        self.lines_table.itemSelectionChanged.connect(actions._sync_quote_selected_line_actions)
        self.lines_table.itemChanged.connect(actions._handle_quote_line_item_changed)
        actions._sync_quote_selected_line_actions()
        inspector_tabs = QTabWidget()
        inspector_tabs.setDocumentMode(True)
        inspector_tabs.setMinimumWidth(350)
        inspector_tabs.setMaximumWidth(470)
        inspector_tabs.setStyleSheet(
            STYLE_INSPECTOR_TABS
        )
        inspector_tabs.tabBar().setExpanding(True)
        inspector_tabs.tabBar().setUsesScrollButtons(False)
        inspector_tabs.tabBar().setElideMode(Qt.ElideNone)
        client_inspector_page = QWidget()
        client_inspector_layout = QVBoxLayout(client_inspector_page)
        client_inspector_layout.setContentsMargins(10, 10, 10, 10)
        client_inspector_layout.setSpacing(6)
        conditions_inspector_page = QWidget()
        conditions_inspector_layout = QVBoxLayout(conditions_inspector_page)
        conditions_inspector_layout.setContentsMargins(8, 8, 8, 8)
        conditions_inspector_layout.setSpacing(0)
        finance_inspector_page = QWidget()
        finance_inspector_layout = QVBoxLayout(finance_inspector_page)
        finance_inspector_layout.setContentsMargins(8, 8, 8, 8)
        finance_inspector_layout.setSpacing(0)

        for panel in (self.quote_client_card, self.quote_exec_card, self.quote_notes_card, self.quote_summary_card):
            panel.setObjectName("QuoteInspectorSection")
            panel.setStyleSheet(
                STYLE_PANEL
            )
            panel.setMaximumWidth(16777215)
            panel.setMinimumWidth(0)
        self.quote_client_card.setMaximumHeight(214)
        self.quote_exec_card.setMaximumHeight(196)
        self.quote_notes_card.setMaximumHeight(16777215)
        self.quote_summary_card.setMaximumHeight(16777215)
        client_inspector_layout.addWidget(self.quote_client_card)
        client_inspector_layout.addWidget(self.quote_exec_card)
        client_inspector_layout.addStretch(1)
        conditions_inspector_layout.addWidget(self.quote_notes_card, 1)
        finance_scroll = QScrollArea()
        finance_scroll.setWidgetResizable(True)
        finance_scroll.setFrameShape(QFrame.NoFrame)
        finance_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        finance_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        finance_scroll.setStyleSheet("QScrollArea { background: transparent; border: 0; }")
        finance_scroll_host = QWidget()
        finance_scroll_layout = QVBoxLayout(finance_scroll_host)
        finance_scroll_layout.setContentsMargins(0, 0, 4, 0)
        finance_scroll_layout.setSpacing(0)
        finance_scroll_layout.addWidget(self.quote_summary_card)
        finance_scroll_layout.addStretch(1)
        finance_scroll.setWidget(finance_scroll_host)
        finance_inspector_layout.addWidget(finance_scroll, 1)
        inspector_tabs.addTab(client_inspector_page, "Cliente")
        inspector_tabs.addTab(conditions_inspector_page, "Condições")
        inspector_tabs.addTab(finance_inspector_page, "Financeiro")

        inspector_host = QWidget()
        inspector_host.setMinimumWidth(350)
        inspector_host.setMaximumWidth(470)
        inspector_host_layout = QVBoxLayout(inspector_host)
        inspector_host_layout.setContentsMargins(0, 0, 0, 0)
        inspector_host_layout.setSpacing(6)
        inspector_command_row = QHBoxLayout()
        inspector_command_row.setContentsMargins(2, 0, 2, 0)
        inspector_command_row.setSpacing(8)
        inspector_caption = QLabel("Dados da proposta")
        inspector_caption.setStyleSheet("font-size: 12px; font-weight: 800; color: #34423a;")
        self.quote_inspector_save_btn = QPushButton("Guardar")
        self.quote_inspector_save_btn.setProperty("variant", "success")
        self.quote_inspector_save_btn.setProperty("compact", "true")
        self.quote_inspector_save_btn.setMinimumWidth(92)
        self.quote_inspector_save_btn.setMaximumWidth(108)
        self.quote_inspector_save_btn.clicked.connect(actions._save_quote)
        inspector_command_row.addWidget(inspector_caption)
        inspector_command_row.addStretch(1)
        inspector_command_row.addWidget(self.quote_inspector_save_btn)
        inspector_host_layout.addLayout(inspector_command_row)
        inspector_host_layout.addWidget(inspector_tabs, 1)

        lines_card.setMinimumHeight(500)
        lines_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        workspace_split = QSplitter(Qt.Horizontal)
        workspace_split.setChildrenCollapsible(False)
        workspace_split.setHandleWidth(7)
        workspace_split.addWidget(lines_card)
        workspace_split.addWidget(inspector_host)
        workspace_split.setSizes([1320, 450])
        workspace_split.setStretchFactor(0, 1)
        workspace_split.setStretchFactor(1, 0)
        workspace_split.setCollapsible(1, True)
        self.quote_workspace_split = workspace_split
        self.quote_inspector_tabs = inspector_tabs
        self.quote_inspector_btn = QPushButton("Ocultar painel")
        self.quote_inspector_btn.setProperty("variant", "secondary")
        self.quote_inspector_btn.setProperty("compact", "true")
        self.quote_inspector_btn.setMinimumHeight(24)
        self.quote_inspector_btn.setMaximumHeight(26)
        self.quote_inspector_btn.clicked.connect(actions._toggle_quote_inspector)
        line_title_row.insertWidget(max(0, line_title_row.count() - 1), self.quote_inspector_btn)
        workspace_split.setMinimumHeight(500)
        detail_layout.addWidget(workspace_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        self._quote_save_feedback_timer = QTimer(self)
        self._quote_save_feedback_timer.setSingleShot(True)
        self._quote_save_feedback_timer.timeout.connect(actions._reset_quote_save_button)

    def eventFilter(self, watched, event):  # type: ignore[override]
        combo = self._combo_click_targets.get(watched)
        if combo is not None and event.type() == QEvent.MouseButtonPress:
            combo.showPopup()
            return True
        return super().eventFilter(watched, event)
