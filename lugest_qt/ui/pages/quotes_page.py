from __future__ import annotations
import html
import math
import os
import re
import tempfile
import unicodedata
from PySide6.QtCore import QDate, QEvent, QTimer, Qt
from PySide6.QtGui import QBrush, QColor, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QAbstractSpinBox,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from datetime import datetime
from lugest_core.cad.profile_analysis import analyze_profile_cut_features, render_step_preview_image
from pathlib import Path
from urllib.parse import quote
from .laser_batch_quote_dialog import LaserBatchQuoteDialog
from .laser_nesting_dialog import LaserNestingDialog
from .laser_quote_dialogs import (
    LaserQuoteDialog,
    LaserSettingsDialog,
    _canonical_material_family as _laser_canonical_material_family,
    _display_material_family as _laser_display_material_family,
    _guess_material_family as _laser_guess_material_family,
    _set_combo_values as _laser_set_combo_values,
    _settings_gas_names as _laser_settings_gas_names,
    _settings_material_names as _laser_settings_material_names,
    _settings_material_subtypes as _laser_settings_material_subtypes,
)
from .materials_page import _MaterialEditorDialog
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    configure_table as _configure_table,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    repolish as _repolish,
    run_process_async as _run_process_async,
    selected_row_index as _selected_row_index,
    set_panel_tone as _set_panel_tone,
    set_table_columns as _set_table_columns,
    smart_sort_key as _smart_sort_key,
    state_tone as _state_tone,
    table_visible_height as _table_visible_height,
)
from .runtime_support import (
    LIST_TABLE_FONT_PX,
    LIST_TABLE_ROW_PX,
    _build_operation_selector,
    _fmt_eur,
    _open_operation_cost_profiles_dialog,
    _open_quote_operation_detail_dialog,
    _operation_tokens,
    _reference_catalog_dialog,
)
from ..widgets import CardFrame, ClickableDateEdit as QDateEdit, FlexibleDecimalSpinBox as QDoubleSpinBox


class QuotesPage(QWidget):
    page_title = "Orçamentos"
    page_subtitle = "Lista de orçamentos primeiro e detalhe apenas quando abres o registo."
    uses_backend_reload = True

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

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self._combo_click_targets: dict[QWidget, QComboBox] = {}
        self.rows: list[dict] = []
        self.client_rows: list[dict] = []
        self.line_rows: list[dict] = []
        self.presets: dict = {}
        self.current_number = ""

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
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.new_quote_btn = QPushButton("Novo orcamento")
        self.new_quote_btn.clicked.connect(self._new_quote)
        self.structure_quote_btn = QPushButton("Orcamento Estruturas")
        self.structure_quote_btn.setProperty("variant", "secondary")
        self.structure_quote_btn.clicked.connect(self._new_structure_quote)
        self.open_quote_btn = QPushButton("Abrir orcamento")
        self.open_quote_btn.clicked.connect(self._open_selected_quote)
        self.remove_quote_btn = QPushButton("Remover")
        self.remove_quote_btn.setProperty("variant", "danger")
        self.remove_quote_btn.clicked.connect(self._remove_quote)
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
            """
            QFrame#QuoteOverviewBand {
                background: #ffffff;
                border: 1px solid #d8e0e8;
                border-radius: 8px;
            }
            QFrame#QuoteMetric {
                background: transparent;
                border: 0;
                border-left: 3px solid #7b91a8;
            }
            QFrame#QuoteMetric QLabel {
                background: transparent;
                border: 0;
            }
            """
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
        self.table.itemSelectionChanged.connect(self._sync_list_buttons)
        self.table.itemDoubleClicked.connect(lambda *_args: self._open_selected_quote())
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
            "QScrollArea { background: transparent; border: 0; }"
            "QScrollBar:vertical { width: 12px; background: #eef1ee; border: 0; margin: 2px; }"
            "QScrollBar::handle:vertical { min-height: 36px; background: #a6aea8; border-radius: 5px; }"
            "QScrollBar::handle:vertical:hover { background: #858e87; }"
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }"
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }"
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
        back_btn.clicked.connect(self._show_list)
        save_btn = QPushButton("Guardar")
        save_btn.setProperty("variant", "success")
        save_btn.clicked.connect(self._save_quote)
        self.quote_save_btn = save_btn
        edit_btn = QPushButton("Em edicao")
        edit_btn.setProperty("variant", "secondary")
        edit_btn.clicked.connect(lambda: self._set_quote_state("Em edição"))
        sent_btn = QPushButton("Enviar")
        sent_btn.setProperty("variant", "secondary")
        sent_btn.clicked.connect(lambda: self._set_quote_state("Enviado"))
        approve_btn = QPushButton("Aprovado")
        approve_btn.setProperty("variant", "success")
        approve_btn.clicked.connect(lambda: self._set_quote_state("Aprovado"))
        reject_btn = QPushButton("Rejeitado")
        reject_btn.setProperty("variant", "rejected")
        reject_btn.setText("✕  Rejeitado")
        reject_btn.clicked.connect(lambda: self._set_quote_state("Rejeitado"))
        convert_btn = QPushButton("Criar encomenda")
        convert_btn.clicked.connect(self._convert_quote)
        purchase_note_btn = QPushButton("Nota encomenda")
        purchase_note_btn.setProperty("variant", "secondary")
        purchase_note_btn.clicked.connect(self._create_quote_purchase_note)
        pdf_actions_btn = QPushButton("PDF  ▾")
        pdf_actions_btn.setProperty("variant", "secondary")
        pdf_menu = QMenu(pdf_actions_btn)
        pdf_menu.addAction("Pré-visualizar PDF", self._preview_quote)
        pdf_menu.addAction("Guardar cópia PDF", self._save_quote_pdf)
        pdf_menu.addAction("Imprimir PDF", self._print_quote_pdf)
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
        self.client_combo.currentTextChanged.connect(self._fill_client_from_combo)
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
        self.client_name_edit.textChanged.connect(lambda _text: self._refresh_quote_identity_label())
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
        self.transport_price_spin.valueChanged.connect(lambda _value: self._render_quote_lines())
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(0.0, 100.0)
        self.discount_spin.setDecimals(2)
        self.discount_spin.setSingleStep(1.0)
        self.discount_spin.setSuffix(" %")
        self.discount_spin.setToolTip("Desconto global do orçamento, aplicado antes do IVA sobre linhas + transporte.")
        self.discount_spin.valueChanged.connect(lambda _value: self._render_quote_lines())
        self.price_increment_spin = QDoubleSpinBox()
        self.price_increment_spin.setRange(-100.0, 500.0)
        self.price_increment_spin.setDecimals(2)
        self.price_increment_spin.setSingleStep(1.0)
        self.price_increment_spin.setSuffix(" %")
        self.price_increment_spin.setToolTip("Incremento global aplicado aos preços unitários antes de descontos e IVA.")
        self.price_increment_spin.valueChanged.connect(lambda _value: self._render_quote_lines())
        self.discount_mode_combo = QComboBox()
        self.discount_mode_combo.addItem("Total final", "total")
        self.discount_mode_combo.addItem("Por lote / espessura", "lotes_espessura")
        self.discount_mode_combo.setToolTip(
            "Escolhe se o desconto se aplica a todas as linhas ou apenas aos lotes/espessuras selecionados."
        )
        self.discount_mode_combo.currentIndexChanged.connect(lambda _index: self._render_quote_lines())
        self.discount_groups_btn = QPushButton("Escolher lotes")
        self.discount_groups_btn.setProperty("compact", "true")
        self.discount_groups_btn.setProperty("variant", "secondary")
        self.discount_groups_btn.clicked.connect(self._pick_quote_discount_groups)
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
        self.iva_spin.valueChanged.connect(lambda _value: self._render_quote_lines())
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
        self.delivery_default_btn.clicked.connect(self._apply_default_delivery_deadline)
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
            widget.valueChanged.connect(self._recalc_transport_calc)

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
            """
            QFrame#Card {
                background: #fffaf0;
                border: 1.5px solid #e4bc6a;
                border-radius: 16px;
            }
            QLabel {
                color: #6f4a12;
            }
            QFrame#Card QLineEdit,
            QFrame#Card QComboBox,
            QFrame#Card QSpinBox,
            QFrame#Card QDoubleSpinBox {
                background: #ffffff;
                border: 1px solid #d4b170;
                border-radius: 8px;
                padding: 5px 8px;
            }
            QFrame#Card QComboBox::drop-down,
            QFrame#Card QSpinBox::down-button,
            QFrame#Card QDoubleSpinBox::down-button,
            QFrame#Card QSpinBox::up-button,
            QFrame#Card QDoubleSpinBox::up-button {
                width: 22px;
                border-left: 1px solid #e2c690;
                background: #fff6e4;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }
            """
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
            """
            QTabWidget::pane {
                border: 1px solid #cfd8cb;
                border-radius: 7px;
                top: -1px;
                background: rgba(255, 255, 255, 0.92);
            }
            QTabBar::tab {
                font-size: 10px;
                min-height: 34px;
                min-width: 0;
                padding: 5px 7px;
                margin-right: 0;
                border: 1px solid #cbd7c5;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                background: #eef4e9;
                color: #42523f;
                font-weight: 700;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                border-color: #81a962;
                color: #2f432d;
                font-weight: 800;
            }
            QTabBar::tab:hover:!selected {
                background: #f5f8f2;
            }
            """
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
            """
            QFrame#Card {
                background: #f7f8f6;
                border: 1px solid #d5dbd3;
                border-radius: 8px;
            }
            QFrame#Card QLabel {
                font-size: 9.5px;
                color: #536057;
                font-weight: 700;
            }
            QFrame#Card QLineEdit,
            QFrame#Card QComboBox,
            QFrame#Card QAbstractSpinBox {
                background: #ffffff;
                border: 1px solid #c5d0c1;
                border-bottom: 1px solid #95aa8d;
                border-radius: 7px;
                font-size: 10px;
                min-height: 27px;
                padding: 1px 6px;
            }
            QFrame#Card QComboBox::drop-down,
            QFrame#Card QAbstractSpinBox::up-button,
            QFrame#Card QAbstractSpinBox::down-button {
                width: 20px;
                border-left: 1px solid #c5d0c1;
                border-bottom: 1px solid #95aa8d;
                background: #eef4ea;
                border-top-right-radius: 7px;
                border-bottom-right-radius: 7px;
            }
            QFrame#Card QComboBox::down-arrow {
                width: 9px;
                height: 9px;
            }
            QFrame#TransportFieldBox {
                background: #ffffff;
                border: 1px solid #bdcabb;
                border-bottom: 1px solid #859d7d;
                border-radius: 6px;
            }
            QFrame#TransportFieldBox QLineEdit,
            QFrame#TransportFieldBox QComboBox,
            QFrame#TransportFieldBox QAbstractSpinBox {
                background: transparent;
                border: 0;
                border-radius: 0;
                padding: 0 5px;
                min-height: 21px;
                font-size: 10px;
            }
            QFrame#TransportFieldBox QComboBox::drop-down,
            QFrame#TransportFieldBox QAbstractSpinBox::up-button,
            QFrame#TransportFieldBox QAbstractSpinBox::down-button {
                width: 18px;
                border-left: 1px solid #c9d4c5;
                border-bottom: 0;
                background: #f0f5ed;
                border-top-right-radius: 5px;
                border-bottom-right-radius: 5px;
            }
            """
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
            """
            QFrame#Card {
                background: #f7f8f6;
                border: 1px solid #d5dbd3;
                border-radius: 8px;
            }
            QFrame#Card QLabel {
                font-size: 9.5px;
                color: #536057;
            }
            """
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
        fill_notes_btn.clicked.connect(self._fill_pdf_notes_from_context)
        apply_transport_btn = QPushButton("Aplicar ao orçamento")
        apply_transport_btn.setProperty("variant", "secondary")
        apply_transport_btn.setProperty("compact", "true")
        apply_transport_btn.setProperty("quoteTransportAction", "true")
        apply_transport_btn.setToolTip("Aplica o cálculo do transporte ao orçamento atual.")
        apply_transport_btn.clicked.connect(self._apply_transport_calc)
        clear_transport_btn = QPushButton("Remover transporte")
        clear_transport_btn.setProperty("variant", "secondary")
        clear_transport_btn.setProperty("compact", "true")
        clear_transport_btn.setProperty("quoteTransportAction", "true")
        clear_transport_btn.setToolTip("Limpa o modo, transportadora, zona e valor de transporte deste orçamento.")
        clear_transport_btn.clicked.connect(self._clear_transport)
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
            op_btn.clicked.connect(lambda _checked=False, text=op_text: self._append_pdf_note(f"- Foi considerado: {text}."))
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
            """
            QFrame#QuoteSummaryTotalPanel {
                background: #f5f8f3;
                border: 1px solid #c5d3c0;
                border-left: 4px solid #6f9f45;
                border-radius: 7px;
            }
            QFrame#QuoteSummaryTotalPanel QLabel {
                background: transparent;
                border: 0;
            }
            """
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
            """
            QFrame#QuoteFinancialBreakdown {
                background: #ffffff;
                border: 1px solid #d7ded5;
                border-radius: 7px;
            }
            QFrame#QuoteFinancialBreakdown QLabel {
                background: transparent;
                border: 0;
                padding: 0;
            }
            QFrame#QuoteSummaryDivider {
                background: #e6ebe4;
                border: 0;
            }
            """
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
            """
            QFrame#QuoteFinancialControls {
                background: #f5f7f4;
                border: 1px solid #d7ded5;
                border-radius: 7px;
            }
            QFrame#QuoteFinancialControls QLabel {
                background: transparent;
                border: 0;
            }
            """
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
                "QDoubleSpinBox, QComboBox, QPushButton {"
                " background: #ffffff;"
                " border: 1px solid #bdc9b9;"
                " border-bottom: 1px solid #8da086;"
                " border-radius: 6px;"
                " padding: 0 8px;"
                " font-size: 10px;"
                " font-weight: 700;"
                " color: #26342b;"
                "}"
                "QDoubleSpinBox::up-button, QDoubleSpinBox::down-button, QComboBox::drop-down {"
                " width: 20px;"
                " border-left: 1px solid #cbd5c8;"
                " background: #f2f5f0;"
                "}"
                "QPushButton:hover, QComboBox:hover, QDoubleSpinBox:hover {"
                " border-color: #6f9f45;"
                " background: #fbfdf9;"
                "}"
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
            """
            QFrame#QuoteDiscountStatus {
                background: #f7f9f6;
                border: 1px solid #d8e0d5;
                border-radius: 6px;
            }
            QFrame#QuoteDiscountStatus QLabel {
                background: transparent;
                border: 0;
            }
            """
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
        add_line_btn.clicked.connect(self._add_line)
        laser_batch_btn = QPushButton("Lote DXF/DWG")
        laser_batch_btn.setProperty("variant", "secondary")
        laser_batch_btn.clicked.connect(self._add_laser_batch_lines)
        laser_batch_btn.setToolTip("Permite carregar um único desenho ou um lote completo de DXF/DWG.")
        laser_nesting_btn = QPushButton("Nesting")
        laser_nesting_btn.setProperty("variant", "secondary")
        laser_nesting_btn.clicked.connect(self._open_laser_nesting)
        profile_step_btn = QPushButton("STEP/IGS Perfil")
        profile_step_btn.setProperty("variant", "secondary")
        profile_step_btn.clicked.connect(self._open_profile_step_igs_quote_builder)
        add_model_btn = QPushButton("Conjunto/Modelo")
        add_model_btn.setProperty("variant", "secondary")
        add_model_btn.clicked.connect(self._manage_saved_conjuntos)
        conjunto_builder_btn = QPushButton("Conjunto calculado")
        conjunto_builder_btn.setProperty("variant", "secondary")
        conjunto_builder_btn.clicked.connect(self._open_calculated_assembly_builder)
        save_selected_group_btn = QPushButton("Guardar conjunto/modelo")
        save_selected_group_btn.setProperty("variant", "secondary")
        save_selected_group_btn.clicked.connect(self._save_selected_lines_as_group)
        structure_builder_btn = QPushButton("Estrutura metalica")
        structure_builder_btn.setProperty("variant", "secondary")
        structure_builder_btn.clicked.connect(self._open_structure_quote_builder)
        manage_models_btn = QPushButton("Modelos")
        manage_models_btn.setProperty("variant", "secondary")
        manage_models_btn.clicked.connect(self._manage_assembly_models)
        laser_cfg_btn = QPushButton("Config. Laser")
        laser_cfg_btn.setProperty("variant", "secondary")
        laser_cfg_btn.clicked.connect(self._configure_laser_profiles)
        operation_cfg_btn = QPushButton("Config. Operacoes")
        operation_cfg_btn.setProperty("variant", "secondary")
        operation_cfg_btn.clicked.connect(self._configure_operation_profiles)
        edit_line_btn = QPushButton("Editar linha")
        edit_line_btn.setProperty("variant", "secondary")
        edit_line_btn.clicked.connect(self._edit_line)
        prepare_line_btn = QPushButton("Preparar produção")
        prepare_line_btn.setProperty("variant", "secondary")
        prepare_line_btn.clicked.connect(self._prepare_selected_line_for_production)
        check_weight_btn = QPushButton("Verificar peso")
        check_weight_btn.setProperty("variant", "secondary")
        check_weight_btn.clicked.connect(self._check_selected_line_weight)
        remove_line_btn = QPushButton("Remover linha")
        remove_line_btn.setProperty("variant", "danger")
        remove_line_btn.clicked.connect(self._remove_line)
        remove_line_btn.setToolTip("Remove as linhas marcadas; se nenhuma estiver marcada, remove a selecao atual.")
        self.remove_quote_lines_btn = remove_line_btn
        open_draw_btn = QPushButton("Ver desenho")
        open_draw_btn.setProperty("variant", "secondary")
        open_draw_btn.clicked.connect(self._open_line_drawing)
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
            """
            QTabWidget::pane {
                border: 1px solid #d8e0e8;
                border-radius: 6px;
                background: #f8fafc;
                top: -1px;
            }
            QTabBar::tab {
                min-height: 22px;
                min-width: 116px;
                padding: 3px 12px;
                margin-right: 3px;
                border: 1px solid #d8e0e8;
                background: #eef3f7;
                color: #475467;
                font-size: 9px;
                font-weight: 800;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #132238;
                border-bottom-color: #ffffff;
            }
            """
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
            "QTableWidget { font-family: 'Segoe UI'; font-size: 10px; }"
            " QTableWidget::indicator { width: 18px; height: 18px; }"
            " QHeaderView::section { font-family: 'Segoe UI Semibold'; font-size: 10px; padding: 5px 5px; font-weight: 600; }"
            " QScrollBar:vertical { background: #eceeeb; width: 14px; margin: 0; border-left: 1px solid #cfd3cf; }"
            " QScrollBar::handle:vertical { background: #8f9691; min-height: 28px; border-radius: 6px; margin: 2px; }"
            " QScrollBar::handle:vertical:hover { background: #747b76; }"
            " QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }"
            " QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }"
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
        self._quote_lines_sort_section = -1
        self._quote_lines_sort_order = Qt.AscendingOrder
        self.lines_table.horizontalHeader().sectionClicked.connect(self._handle_quote_lines_sort)
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
            """
            QFrame#QuoteSelectedLineFooter {
                background: #f1f3f1;
                border: 1px solid #d4d8d4;
                border-radius: 7px;
            }
            QLabel#QuoteSelectedLineCaption {
                color: #3f4943;
                font-size: 9px;
                font-weight: 900;
                background: transparent;
                border: 0;
            }
            """
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
        self.lines_table.itemSelectionChanged.connect(self._sync_quote_selected_line_actions)
        self.lines_table.itemChanged.connect(self._handle_quote_line_item_changed)
        self._sync_quote_selected_line_actions()
        inspector_tabs = QTabWidget()
        inspector_tabs.setDocumentMode(True)
        inspector_tabs.setMinimumWidth(350)
        inspector_tabs.setMaximumWidth(470)
        inspector_tabs.setStyleSheet(
            """
            QTabWidget::pane {
                border: 1px solid #d0d8cd;
                border-radius: 7px;
                background: #f7f9f6;
                top: -1px;
            }
            QTabBar::tab {
                min-height: 30px;
                min-width: 0;
                padding: 5px 4px;
                margin-right: 0;
                border: 1px solid #d0d8cd;
                background: #edf2eb;
                color: #4d5b51;
                font-size: 9px;
                font-weight: 800;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #355d2c;
                border-top: 2px solid #6f9f45;
                border-bottom-color: #ffffff;
            }
            QTabBar::tab:hover:!selected {
                background: #f4f7f2;
                color: #45673b;
            }
            """
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
                """
                QFrame#QuoteInspectorSection {
                    background: transparent;
                    border: 0;
                    border-radius: 0;
                }
                """
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
        self.quote_inspector_save_btn.clicked.connect(self._save_quote)
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
        self.quote_inspector_btn.clicked.connect(self._toggle_quote_inspector)
        line_title_row.insertWidget(max(0, line_title_row.count() - 1), self.quote_inspector_btn)
        workspace_split.setMinimumHeight(500)
        detail_layout.addWidget(workspace_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        self._quote_save_feedback_timer = QTimer(self)
        self._quote_save_feedback_timer.setSingleShot(True)
        self._quote_save_feedback_timer.timeout.connect(self._reset_quote_save_button)
        self._clear_quote_detail()
        self._show_list()

    def eventFilter(self, watched, event):  # type: ignore[override]
        combo = self._combo_click_targets.get(watched)
        if combo is not None and event.type() == QEvent.MouseButtonPress:
            combo.showPopup()
            return True
        return super().eventFilter(watched, event)

    def refresh(self) -> None:
        previous = self.current_number
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(previous)
        current_filter = self.filter_edit.currentText().strip()
        self.client_rows = self.backend.orc_clients()
        self.presets = self.backend.order_presets()
        self._set_client_items()
        current_carrier = self.transport_carrier_combo.currentText().strip()
        self.transport_carrier_combo.blockSignals(True)
        self.transport_carrier_combo.clear()
        self.transport_carrier_combo.addItem("")
        for supplier in list(self.backend.ne_suppliers() or []):
            self.transport_carrier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        self.transport_carrier_combo.setCurrentText(current_carrier)
        self.transport_carrier_combo.blockSignals(False)
        current_zone = self.transport_zone_combo.currentText().strip()
        self.transport_zone_combo.blockSignals(True)
        self.transport_zone_combo.clear()
        self.transport_zone_combo.addItem("")
        for value in list(self.backend.transport_zone_options() or []):
            self.transport_zone_combo.addItem(str(value))
        self.transport_zone_combo.setCurrentText(current_zone)
        self.transport_zone_combo.blockSignals(False)
        current_year = self.year_combo.currentText().strip() or "Todos"
        year_values = ["Todos"] + list(self.backend.orc_available_years())
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        self.year_combo.addItems(year_values)
        self.year_combo.setCurrentText(current_year if current_year in year_values else year_values[0])
        self.year_combo.blockSignals(False)
        self.rows = self.backend.orc_rows(current_filter, self.state_combo.currentText(), self.year_combo.currentText() or "Todos")
        self._refresh_quote_overview()
        self.filter_edit.blockSignals(True)
        if self.filter_edit.count() == 0:
            self.filter_edit.addItem("")
        known_values = {self.filter_edit.itemText(i) for i in range(self.filter_edit.count())}
        for row in self.rows:
            numero = str(row.get("numero", "") or "")
            if numero and numero not in known_values:
                self.filter_edit.addItem(numero)
                known_values.add(numero)
        self.filter_edit.setCurrentText(current_filter)
        self.filter_edit.blockSignals(False)
        _fill_table(
            self.table,
            [
                [
                    r.get("numero", "-"),
                    r.get("cliente", "-"),
                    r.get("data", "-"),
                    r.get("estado", "-"),
                    _fmt_eur(r.get("total", 0)),
                    r.get("numero_encomenda", "-"),
                ]
                for r in self.rows
            ],
            align_center_from=2,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.table, row_index, str(row.get("estado", "")))
            total_item = self.table.item(row_index, 4)
            if total_item is not None:
                total_item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
        if self.table.rowCount() == 0:
            self._clear_quote_detail()
            self._show_list()
            self._sync_list_buttons()
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "") or "").strip() == previous:
                    row_index = index
                    break
        self.table.selectRow(row_index)
        if keep_detail:
            self._load_quote(str(self.rows[row_index].get("numero", "") or "").strip())
            self._show_detail()
        else:
            self._show_list()
        self._sync_list_buttons()

    def _refresh_quote_overview(self) -> None:
        rows = list(self.rows or [])
        state_counts = {"editing": 0, "sent": 0, "approved": 0}
        total_value = 0.0
        for row in rows:
            state = str(row.get("estado", "") or "").strip().casefold()
            total_value += float(row.get("total", 0) or 0)
            if "edi" in state:
                state_counts["editing"] += 1
            elif "enviado" in state:
                state_counts["sent"] += 1
            elif "aprovado" in state:
                state_counts["approved"] += 1
        labels = getattr(self, "quote_metric_labels", {})
        if isinstance(labels, dict):
            if isinstance(labels.get("value"), QLabel):
                labels["value"].setText(_fmt_eur(total_value))
            for key in ("editing", "sent", "approved"):
                if isinstance(labels.get(key), QLabel):
                    labels[key].setText(str(state_counts[key]))
        count = len(rows)
        self.quote_list_count_label.setText(f"{count} orçamento" if count == 1 else f"{count} orçamentos")

    def _show_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_buttons()

    def _show_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def _toggle_quote_inspector(self) -> None:
        inspector = getattr(self, "quote_inspector_tabs", None)
        splitter = getattr(self, "quote_workspace_split", None)
        button = getattr(self, "quote_inspector_btn", None)
        if not isinstance(inspector, QWidget) or not isinstance(splitter, QSplitter):
            return
        show_inspector = not inspector.isVisible()
        inspector.setVisible(show_inspector)
        if show_inspector:
            splitter.setSizes([max(760, splitter.width() - 420), 420])
        if isinstance(button, QPushButton):
            button.setText("Ocultar painel" if show_inspector else "Mostrar painel")

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _selected_quote_row(self) -> dict:
        current = self.table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def _selected_line_index(self) -> int:
        current = self.lines_table.currentItem()
        if current is None:
            return -1
        source_index = current.data(Qt.UserRole)
        if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
            return source_index
        return current.row() if 0 <= current.row() < len(self.line_rows) else -1

    def _selected_line_indexes(self) -> list[int]:
        indexes: set[int] = set()
        for item in self.lines_table.selectedItems():
            if item is None:
                continue
            source_index = item.data(Qt.UserRole)
            if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
                indexes.add(source_index)
            elif 0 <= item.row() < len(self.line_rows):
                indexes.add(item.row())
        indexes = sorted(indexes)
        if indexes:
            return indexes
        single = self._selected_line_index()
        return [single] if single >= 0 else []

    def _checked_line_indexes(self) -> list[int]:
        indexes: set[int] = set()
        for visual_row in range(self.lines_table.rowCount()):
            item = self.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is None or item.checkState() != Qt.Checked:
                continue
            source_index = item.data(Qt.UserRole)
            if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
                indexes.add(source_index)
        return sorted(indexes)

    def _handle_quote_line_item_changed(self, item: QTableWidgetItem) -> None:
        if bool(getattr(self, "_rendering_quote_lines", False)):
            return
        if item is not None and item.column() == self.LINE_COL_MARK:
            self._sync_quote_selected_line_actions()

    def _select_quote_line_source_index(self, source_index: int) -> None:
        for visual_row in range(self.lines_table.rowCount()):
            item = self.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is not None and item.data(Qt.UserRole) == source_index:
                self.lines_table.selectRow(visual_row)
                return

    def _sync_quote_selected_line_actions(self) -> None:
        indexes = self._selected_line_indexes() if hasattr(self, "lines_table") else []
        checked_indexes = self._checked_line_indexes() if hasattr(self, "lines_table") else []
        has_selection = bool(indexes)
        for button in getattr(self, "quote_selected_line_buttons", ()):
            button.setEnabled(has_selection)
        remove_button = getattr(self, "remove_quote_lines_btn", None)
        removal_count = len(checked_indexes or indexes)
        if isinstance(remove_button, QPushButton):
            remove_button.setEnabled(removal_count > 0)
            remove_button.setText("Remover linha" if removal_count <= 1 else f"Remover {removal_count} linhas")
        caption = getattr(self, "quote_selected_line_caption", None)
        if not isinstance(caption, QLabel):
            return
        if checked_indexes:
            caption.setText(
                "1 LINHA MARCADA PARA REMOVER"
                if len(checked_indexes) == 1
                else f"{len(checked_indexes)} LINHAS MARCADAS PARA REMOVER"
            )
            return
        if not indexes:
            caption.setText("SELECIONA UMA LINHA")
            return
        if len(indexes) > 1:
            caption.setText(f"{len(indexes)} LINHAS SELECIONADAS")
            return
        row = dict(self.line_rows[indexes[0]] or {}) if indexes[0] < len(self.line_rows) else {}
        reference = self._quote_line_primary_ref(row)
        caption.setText(f"LINHA · {reference}" if reference and reference != "-" else "LINHA SELECIONADA")

    def _quote_line_sort_value(self, row: dict, section: int) -> tuple[int, object]:
        values: tuple[object, ...] = (
            self._quote_line_type_label(row),
            self._quote_line_primary_ref(row),
            row.get("ref_externa", ""),
            row.get("descricao", ""),
            self._quote_line_material_display(row),
            row.get("espessura", "") or self._quote_line_unit_display(row),
            row.get("operacao", ""),
            row.get("tempo_peca_min", 0),
            row.get("qtd", 0),
            row.get("preco_unit_incrementado", row.get("preco_unit", 0)),
            row.get("preco_unit_desconto", row.get("preco_unit", 0)),
            row.get("total_desconto", row.get("total", 0)),
            row.get("conjunto_nome", ""),
        )
        data_section = max(0, int(section) - 1)
        return _smart_sort_key(values[max(0, min(data_section, len(values) - 1))])

    def _handle_quote_lines_sort(self, section: int) -> None:
        if section == self.LINE_COL_MARK:
            real_items = [
                self.lines_table.item(row, self.LINE_COL_MARK)
                for row in range(self.lines_table.rowCount())
                if isinstance(self.lines_table.item(row, self.LINE_COL_MARK), QTableWidgetItem)
                and isinstance(self.lines_table.item(row, self.LINE_COL_MARK).data(Qt.UserRole), int)
            ]
            mark_all = bool(real_items) and not all(item.checkState() == Qt.Checked for item in real_items)
            self._rendering_quote_lines = True
            for item in real_items:
                item.setCheckState(Qt.Checked if mark_all else Qt.Unchecked)
            self._rendering_quote_lines = False
            self._sync_quote_selected_line_actions()
            return
        if self._quote_lines_sort_section == section:
            self._quote_lines_sort_order = (
                Qt.DescendingOrder
                if self._quote_lines_sort_order == Qt.AscendingOrder
                else Qt.AscendingOrder
            )
        else:
            self._quote_lines_sort_section = section
            self._quote_lines_sort_order = Qt.AscendingOrder
        header = self.lines_table.horizontalHeader()
        header.setSortIndicator(section, self._quote_lines_sort_order)
        header.setSortIndicatorShown(True)
        self._render_quote_lines()

    def _sync_list_buttons(self) -> None:
        has_row = bool(self._selected_quote_row())
        self.open_quote_btn.setEnabled(has_row)
        self.remove_quote_btn.setEnabled(has_row)

    def _set_client_items(self) -> None:
        current = self.client_combo.currentText().strip()
        self.client_combo.blockSignals(True)
        self.client_combo.clear()
        for row in self.client_rows:
            label = str(row.get("label", "") or f"{row.get('codigo', '')} - {row.get('nome', '')}").strip(" -")
            self.client_combo.addItem(label, row)
        self.client_combo.setCurrentText(current)
        self.client_combo.blockSignals(False)
        current_exec = self.executed_combo.currentText().strip()
        self.executed_combo.clear()
        for value in list(self.backend.ensure_data().get("orcamentistas", []) or []):
            self.executed_combo.addItem(str(value))
        self.executed_combo.setCurrentText(current_exec)
        self.workcenter_combo.blockSignals(True)
        self.workcenter_combo.clear()
        self.workcenter_combo.blockSignals(False)

    def _client_lookup(self, text: str) -> dict:
        raw = str(text or "").strip()
        if not raw:
            return {}
        code = raw.split(" - ", 1)[0].strip()
        for row in self.client_rows:
            if code and code == str(row.get("codigo", "") or "").strip():
                return row
            label = str(row.get("label", "") or f"{row.get('codigo', '')} - {row.get('nome', '')}").strip(" -")
            if raw.lower() == label.lower():
                return row
        return {}

    def _client_code_from_text(self, text: str) -> str:
        return str(self._client_lookup(text).get("codigo", "") or text.split(" - ", 1)[0].strip())

    def _fill_client_from_combo(self) -> None:
        row = self._client_lookup(self.client_combo.currentText())
        if not row:
            self._refresh_quote_identity_label()
            return
        if not self.client_name_edit.text().strip():
            self.client_name_edit.setText(str(row.get("nome", "") or "").strip())
        if not self.client_company_edit.text().strip():
            self.client_company_edit.setText(str(row.get("nome", "") or "").strip())
        if not self.client_nif_edit.text().strip():
            self.client_nif_edit.setText(str(row.get("nif", "") or "").strip())
        if not self.client_address_edit.text().strip():
            self.client_address_edit.setText(str(row.get("morada", "") or "").strip())
        if not self.client_contact_edit.text().strip():
            self.client_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        if not self.client_email_edit.text().strip():
            self.client_email_edit.setText(str(row.get("email", "") or "").strip())
        self._refresh_quote_identity_label()

    def _quote_line_type(self, row: dict) -> str:
        return str(self.backend.desktop_main.normalize_orc_line_type((row or {}).get("tipo_item")) or self.backend.desktop_main.ORC_LINE_TYPE_PIECE)

    def _quote_line_is_raw_material_ui(self, row: dict) -> bool:
        data = dict(row or {})
        if self._quote_line_type(data) != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if str(data.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(data.get("stock_material_id", "") or "").strip():
            return True
        ref_ext = str(data.get("ref_externa", "") or "").strip().upper()
        if ref_ext.startswith("MAT") and ref_ext[3:].isdigit():
            if str(data.get("desenho", "") or "").strip():
                return False
            try:
                tempo = float(data.get("tempo_peca_min", data.get("tempo_pecas_min", 0)) or 0)
            except Exception:
                tempo = 0.0
            if tempo > 0:
                return False
            try:
                op_norm = self.backend.desktop_main.norm_text(str(data.get("operacao", "") or "").strip())
            except Exception:
                op_norm = str(data.get("operacao", "") or "").strip().lower()
            return op_norm in {"", "-", "stockmp", "materia prima", "materia-prima"}
        subtype = str(data.get("material_subtype", "") or data.get("calc_mode", "") or "").strip()
        try:
            return self.backend.desktop_main.norm_text(subtype) == "stockmp"
        except Exception:
            return subtype.lower() == "stockmp"

    def _quote_line_should_be_raw_material_ui(self, row: dict) -> bool:
        if self._quote_line_is_raw_material_ui(row):
            return True
        data = dict(row or {})
        if self._quote_line_type(data) != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if str(data.get("desenho", "") or "").strip():
            return False
        try:
            tempo = float(data.get("tempo_peca_min", data.get("tempo_pecas_min", 0)) or 0)
        except Exception:
            tempo = 0.0
        if tempo > 0:
            return False
        operacao = str(data.get("operacao", "") or "").strip()
        if operacao:
            try:
                op_norm = self.backend.desktop_main.norm_text(operacao)
            except Exception:
                op_norm = operacao.lower()
            if op_norm not in {"stockmp", "materia prima", "materia-prima"}:
                return False
        has_raw_origin = any(
            str(data.get(key, "") or "").strip()
            for key in (
                "conjunto_codigo",
                "conjunto_nome",
                "calc_mode",
                "material_subtype",
                "quantity_units",
                "weight_total",
                "stock_material_id",
            )
        )
        if not has_raw_origin:
            return False
        marker_text = " ".join(
            str(data.get(key, "") or "")
            for key in ("ref_externa", "material", "material_family", "material_subtype", "calc_mode", "descricao")
        )
        try:
            marker_norm = self.backend.desktop_main.norm_text(marker_text)
        except Exception:
            marker_norm = marker_text.lower()
        raw_tokens = (
            "perfil",
            "ipe",
            "ipn",
            "hea",
            "heb",
            "upn",
            "rebar",
            "ferro nervurado",
            "barra",
            "chapa",
            "tubo",
            "cantoneira",
            "varao",
            "mat",
            "stock mp",
            "stockmp",
        )
        return any(token in marker_norm for token in raw_tokens)

    def _coerce_quote_line_raw_material_ui(self, row: dict) -> dict:
        data = dict(row or {})
        data["tipo_item"] = self.backend.desktop_main.ORC_LINE_TYPE_PIECE
        data["stock_item_kind"] = "raw_material"
        data["ref_interna"] = ""
        data["operacao"] = ""
        data["tempo_peca_min"] = 0.0
        data["desenho"] = ""
        data["laser_base_active"] = False
        data["laser_base_tempo_unit"] = 0.0
        data["laser_base_preco_unit"] = 0.0
        for key in ("operacoes_lista", "operacoes_fluxo", "operacoes_detalhe"):
            data[key] = []
        for key in ("tempos_operacao", "custos_operacao", "quote_cost_snapshot"):
            data[key] = {}
        if not str(data.get("material_subtype", "") or "").strip():
            data["material_subtype"] = str(data.get("calc_mode", "") or "Stock MP").strip()
        return data

    def _quote_line_type_label(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            return "Matéria-prima"
        return str(self.backend.desktop_main.orc_line_type_label(self._quote_line_type(row)) or "-")

    def _quote_line_primary_ref(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            stock_id = str((row or {}).get("stock_material_id", "") or "").strip()
            return stock_id or str((row or {}).get("ref_externa", "") or "").strip() or "-"
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        return str((row or {}).get("ref_interna", "") or "").strip() or "-"

    def _quote_line_material_display(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            return (
                str((row or {}).get("material_family", "") or "").strip()
                or str((row or {}).get("material", "") or "").strip()
                or "-"
            )
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE:
            return "-"
        return str((row or {}).get("material", "") or "").strip() or "-"

    def _quote_line_unit_display(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            subtype = str((row or {}).get("material_subtype", "") or "").strip()
            esp = str((row or {}).get("espessura", "") or "").strip()
            dimensao = str((row or {}).get("dimensao", (row or {}).get("dimensoes", "")) or "").strip()
            if dimensao:
                return dimensao
            if subtype and esp:
                return f"{subtype} | {esp}"
            return esp or subtype or "-"
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            return str((row or {}).get("espessura", "") or "").strip() or "-"
        return str((row or {}).get("produto_unid", "") or "").strip() or "-"

    def _refresh_quote_identity_label(self) -> None:
        if not hasattr(self, "number_label"):
            return
        number = str(getattr(self, "current_number", "") or "").strip() or "Novo orçamento"
        client_name = ""
        if hasattr(self, "client_name_edit"):
            client_name = str(self.client_name_edit.text() or "").strip()
        if not client_name and hasattr(self, "client_combo"):
            client_row = self._client_lookup(self.client_combo.currentText())
            client_name = str(client_row.get("nome", "") or "").strip()
        self.number_label.setText(f"{number}  ·  {client_name}" if client_name else number)

    def _set_quote_header(self, numero: str, estado: str, encomenda: str = "") -> None:
        self.current_number = str(numero or "").strip()
        self._refresh_quote_identity_label()
        self.link_order_label.setText(f"Encomenda gerada: {encomenda}" if encomenda else "Sem encomenda gerada")
        _apply_state_chip(self.state_chip, estado)
        _set_panel_tone(self.quote_header_card, _state_tone(estado))
        _set_panel_tone(self.quote_summary_card, _state_tone(estado))

    def _clear_quote_detail(self) -> None:
        self.line_rows = []
        self.discount_group_keys = []
        self.nesting_bridge_data = {}
        self._set_quote_header(self.backend.orc_next_number(), "Em edicao", "")
        self.client_combo.setCurrentText("")
        self.client_name_edit.clear()
        self.client_company_edit.clear()
        self.client_nif_edit.clear()
        self.client_address_edit.clear()
        self.client_contact_edit.clear()
        self.client_email_edit.clear()
        self.note_cliente_edit.clear()
        self.transport_combo.setCurrentText("")
        self.transport_carrier_combo.setCurrentText("")
        self.transport_zone_combo.setCurrentText("")
        self.transport_price_spin.setValue(0.0)
        self.discount_spin.setValue(0.0)
        self.price_increment_spin.setValue(0.0)
        self.discount_mode_combo.setCurrentIndex(0)
        self.transport_km_spin.setValue(0.0)
        self.transport_rate_spin.setValue(0.65)
        self.transport_diesel_spin.setValue(1.65)
        self.transport_consumption_spin.setValue(8.5)
        self.transport_trip_factor_spin.setValue(2.0)
        user = dict(getattr(self.backend, "user", {}) or {})
        role = str(user.get("role", "") or "").strip().casefold()
        username = str(user.get("username", "") or "").strip()
        self.executed_combo.setCurrentText(username if username and role in {"orcamentista", "orçamentista"} else "")
        self.workcenter_combo.setCurrentText("")
        self.notes_edit.clear()
        self.delivery_date_min = QDate.currentDate()
        self.delivery_date_edit.setMinimumDate(self.delivery_date_min)
        self.delivery_date_edit.setDate(self.delivery_date_min)
        self.iva_spin.setValue(23.0)
        self._recalc_transport_calc()
        self._refresh_nesting_bridge()
        self._render_quote_lines()

    def _refresh_nesting_bridge(self) -> None:
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        if not bridge:
            self.nesting_bridge_label.setText(
                "Preco unitario da tabela = orcamentacao por peca. Nesting = validacao global de materia, stock/retalho e pecas realmente programadas."
            )
            self.nesting_bridge_label.setToolTip(
                "O preco unitario das linhas continua comercial e unitario. O nesting serve para validar consumo real de materia, stock, compra e pecas efetivamente programadas."
            )
            return
        placed = int(bridge.get("part_count_placed", 0) or 0)
        requested = int(bridge.get("part_count_requested", 0) or 0)
        sheets = int(bridge.get("sheet_count", 0) or 0)
        material_cost = _fmt_eur(float(bridge.get("material_net_cost_eur", 0) or 0))
        purchase_cost = _fmt_eur(float(bridge.get("material_purchase_requirement_eur", 0) or 0))
        quoted_total = _fmt_eur(float(bridge.get("quoted_total_eur", 0) or 0))
        rateio_total = _fmt_eur(float(bridge.get("rateio_adjusted_quote_total_eur", 0) or 0))
        rateio_delta = _fmt_eur(float(bridge.get("rateio_delta_eur", 0) or 0))
        method = str(bridge.get("analysis_method", "-") or "-").strip() or "-"
        profile_name = str(bridge.get("selected_profile_name", "") or "").strip() or "Apenas stock"
        self.nesting_bridge_label.setText(
            f"Nesting atual: {placed}/{requested} programadas | {sheets} chapa(s) | {method} | "
            f"perfil {profile_name} | materia real {material_cost} | compra {purchase_cost} | "
            f"comercial atual {quoted_total} | comercial ajustado ao rateio {rateio_total} | delta {rateio_delta}."
        )
        self.nesting_bridge_label.setToolTip(
            "Tabela do orcamento: preco comercial unitario por referencia.\n"
            "Nesting: custo global real de materia e validacao do lote completo.\n"
            "O rateio usa a area liquida colocada para repartir a materia real e a compra necessaria por referencia.\n"
            "Usa a tabela para vender/orcamentar por referencia e o plano de chapa para validar consumo real, compra e cobertura comercial."
        )

    def _quote_discount_mode(self) -> str:
        mode = str(self.discount_mode_combo.currentData() or "").strip().lower()
        return mode if mode in {"total", "lotes_espessura"} else "total"

    def _quote_discount_groups(self) -> list[dict]:
        groups: list[dict] = []
        by_key: dict[str, dict] = {}
        for row_index, row in enumerate(list(self.line_rows or [])):
            line_total = round(float(row.get("total", 0) or (float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0))), 2)
            if line_total <= 0:
                continue
            key, label = self._quote_discount_group_label(row)
            bucket = by_key.setdefault(key, {"key": key, "label": label, "base": 0.0, "rows": []})
            bucket["base"] = round(float(bucket.get("base", 0) or 0) + line_total, 2)
            bucket["rows"].append(row_index)
        groups.extend(by_key.values())
        return groups

    def _quote_selected_discount_group_keys(self) -> set[str]:
        groups = self._quote_discount_groups()
        available = {str(group.get("key", "") or "").strip() for group in groups if str(group.get("key", "") or "").strip()}
        mode = self._quote_discount_mode()
        selected = {
            str(key or "").strip()
            for key in list(getattr(self, "discount_group_keys", []) or [])
            if str(key or "").strip() in available
        }
        if mode == "lotes_espessura":
            return selected or available
        return available

    def _quote_effective_discount_group_keys(self) -> list[str]:
        return sorted(self._quote_selected_discount_group_keys())

    def _sync_discount_targets_label(self) -> None:
        mode = self._quote_discount_mode()
        self.discount_groups_btn.setEnabled(mode == "lotes_espessura")
        if mode != "lotes_espessura":
            self.discount_target_label.setText("Desconto aplicado a todas as linhas do orçamento. O transporte fica fora do desconto.")
            return
        groups = self._quote_discount_groups()
        selected = self._quote_selected_discount_group_keys()
        if not groups:
            self.discount_target_label.setText("Sem lotes elegíveis para desconto neste orçamento.")
            return
        if not selected:
            self.discount_target_label.setText("Sem lotes selecionados manualmente. A pré-visualização aplica a todos os lotes elegíveis.")
            return
        labels = [str(group.get("label", "") or "").strip() for group in groups if str(group.get("key", "") or "").strip() in selected]
        preview = ", ".join(labels[:3])
        if len(labels) > 3:
            preview = f"{preview}, ..."
        self.discount_target_label.setText(f"Desconto aplicado a: {preview}")

    def _pick_quote_discount_groups(self) -> None:
        groups = self._quote_discount_groups()
        if not groups:
            QMessageBox.information(self, "Desconto por lote", "Este orçamento ainda não tem linhas elegíveis para selecionar lotes.")
            return
        selected = self._quote_selected_discount_group_keys()
        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar lotes / espessuras para desconto")
        dialog.resize(560, 420)
        root = QVBoxLayout(dialog)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)
        helper = QLabel("Marca os lotes/espessuras onde queres aplicar o desconto global.")
        helper.setWordWrap(True)
        root.addWidget(helper)
        list_widget = QListWidget()
        for group in groups:
            label = f"{group.get('label', '-')}: {_fmt_eur(float(group.get('base', 0) or 0))}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, str(group.get("key", "") or "").strip())
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if str(group.get("key", "") or "").strip() in selected else Qt.Unchecked)
            list_widget.addItem(item)
        root.addWidget(list_widget, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Aplicar")
        buttons.button(QDialogButtonBox.Cancel).setText("Fechar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        root.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        chosen: list[str] = []
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item is not None and item.checkState() == Qt.Checked:
                key = str(item.data(Qt.UserRole) or "").strip()
                if key:
                    chosen.append(key)
        self.discount_group_keys = chosen
        self._render_quote_lines()

    def _quote_discount_group_label(self, row: dict) -> tuple[str, str]:
        if self._quote_line_is_raw_material_ui(row) or self._quote_line_type(row) == self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            family = str(row.get("material_family", "") or row.get("material", "") or "Material").strip() or "Material"
            subtype = str(row.get("material_subtype", "") or "").strip()
            thickness = str(row.get("espessura", "") or "").strip()
            label = family
            if subtype:
                label = f"{label} / {subtype}"
            if thickness:
                label = f"{label} / {thickness} mm"
            key = f"piece|{family}|{subtype}|{thickness}"
            return key, label
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return "product", "Produtos / outros artigos"
        return "service", "Servicos / outras operacoes"

    def _quote_discount_breakdown(self, selected_keys: set[str], discount_pct: float) -> list[dict]:
        groups: list[dict] = []
        if discount_pct <= 0:
            return groups
        for group in self._quote_discount_groups():
            key = str(group.get("key", "") or "").strip()
            base_value = round(float(group.get("base", 0) or 0), 2)
            if not key or base_value <= 0:
                continue
            discount_value = round(base_value * (discount_pct / 100.0), 2) if key in selected_keys else 0.0
            groups.append(
                {
                    "key": key,
                    "label": str(group.get("label", "") or "").strip() or "-",
                    "base": base_value,
                    "discount": discount_value,
                    "subtotal_after_discount": round(max(0.0, base_value - discount_value), 2),
                }
            )
        return groups

    def _render_quote_lines(self) -> None:
        checked_source_indexes = (
            set(self._checked_line_indexes())
            if hasattr(self, "lines_table") and not bool(getattr(self, "_rendering_quote_lines", False))
            else set()
        )
        subtotal = 0.0
        subtotal_base = 0.0
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        bridge_rows = {
            str(row.get("ref_externa", "") or "").strip(): dict(row or {})
            for row in list(bridge.get("part_rows", []) or [])
            if str(row.get("ref_externa", "") or "").strip()
        }
        discount_pct = float(self.discount_spin.value() or 0)
        increment_pct = float(self.price_increment_spin.value() or 0)
        selected_discount_keys = self._quote_selected_discount_group_keys()
        normalized_rows = []
        for row in self.line_rows:
            if self._quote_line_should_be_raw_material_ui(row):
                row = self._coerce_quote_line_raw_material_ui(row)
            qtd = float(row.get("qtd", 0) or 0)
            preco_unit = float(row.get("preco_unit", 0) or 0)
            subtotal_base = round(subtotal_base + (qtd * preco_unit), 2)
            preco_unit_incrementado = round(max(0.0, preco_unit * (1.0 + (increment_pct / 100.0))), 4)
            total = round(qtd * preco_unit_incrementado, 2)
            row["preco_unit_incrementado"] = preco_unit_incrementado
            row["total"] = total
            group_key, _group_label = self._quote_discount_group_label(row)
            apply_discount = discount_pct > 0 and group_key in selected_discount_keys
            discounted_unit = round(preco_unit_incrementado * (1.0 - (discount_pct / 100.0)), 4) if apply_discount else round(preco_unit_incrementado, 4)
            discounted_total = round(qtd * discounted_unit, 2)
            row["discount_group_key"] = group_key
            row["preco_unit_desconto"] = discounted_unit
            row["total_desconto"] = discounted_total
            row["desconto_aplicado"] = round(max(0.0, total - discounted_total), 2)
            normalized_rows.append(row)
        self.line_rows = normalized_rows
        display_rows = list(enumerate(self.line_rows))
        sort_section = int(getattr(self, "_quote_lines_sort_section", -1))
        if sort_section >= 0:
            reverse = getattr(self, "_quote_lines_sort_order", Qt.AscendingOrder) == Qt.DescendingOrder
            display_rows.sort(
                key=lambda indexed_row: self._quote_line_sort_value(indexed_row[1], sort_section),
                reverse=reverse,
            )
        self._rendering_quote_lines = True
        previous_signal_state = self.lines_table.blockSignals(True)
        updates_were_enabled = self.lines_table.updatesEnabled()
        self.lines_table.setUpdatesEnabled(False)
        _fill_table(
            self.lines_table,
            [
                [
                    "",
                    self._quote_line_type_label(row),
                    self._quote_line_primary_ref(row),
                    row.get("ref_externa", "-") or "-",
                    row.get("descricao", "-") or "-",
                    self._quote_line_material_display(row),
                    self._quote_line_unit_display(row),
                    "-" if (self._quote_line_is_raw_material_ui(row) and not str(row.get("operacao", "") or "").strip()) or self._quote_line_type(row) == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT else (row.get("operacao", "-") or "-"),
                    "-" if (self._quote_line_is_raw_material_ui(row) and not str(row.get("operacao", "") or "").strip()) or self._quote_line_type(row) == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT else f"{float(row.get('tempo_peca_min', 0) or 0):.2f} min",
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    _fmt_eur(float(row.get("preco_unit_incrementado", row.get("preco_unit", 0)) or 0)),
                    _fmt_eur(float(row.get("preco_unit_desconto", row.get("preco_unit", 0)) or 0)),
                    _fmt_eur(float(row.get("total_desconto", row.get("total", 0)) or 0)),
                    row.get("conjunto_nome", "-") or "-",
                ]
                for _source_index, row in display_rows
            ],
            align_center_from=self.LINE_COL_UNIT,
        )
        for visual_row, (source_index, _row) in enumerate(display_rows):
            for col_index in range(self.lines_table.columnCount()):
                item = self.lines_table.item(visual_row, col_index)
                if item is not None:
                    item.setData(Qt.UserRole, source_index)
            mark_item = self.lines_table.item(visual_row, self.LINE_COL_MARK)
            if mark_item is not None:
                mark_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                mark_item.setCheckState(Qt.Checked if source_index in checked_source_indexes else Qt.Unchecked)
                mark_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                mark_item.setToolTip("Marca esta linha para a remover do orçamento.")
        for _ in range(2):
            spacer_row = self.lines_table.rowCount()
            self.lines_table.insertRow(spacer_row)
            self.lines_table.setRowHeight(spacer_row, 24)
            for col_index in range(self.lines_table.columnCount()):
                spacer_item = QTableWidgetItem("")
                spacer_item.setFlags(Qt.ItemIsEnabled)
                spacer_item.setBackground(QBrush(QColor("#f8fafc")))
                self.lines_table.setItem(spacer_row, col_index, spacer_item)
        transport = float(self.transport_price_spin.value() or 0)
        for row_index, (_source_index, row) in enumerate(display_rows):
            subtotal += float(row.get("total", 0) or 0)
            _paint_table_row(self.lines_table, row_index, "Preparacao")
            for col_index in (self.LINE_COL_PRICE, self.LINE_COL_DISCOUNTED_PRICE, self.LINE_COL_TOTAL):
                item = self.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            line_type = self._quote_line_type(row)
            ref_externa = str(row.get("ref_externa", "") or "").strip()
            bridge_row = bridge_rows.get(ref_externa, {})
            tooltip_lines = [
                f"Ref. externa: {ref_externa or '-'}",
                f"Descricao: {str(row.get('descricao', '') or '-').strip() or '-'}",
                f"Operacao: {str(row.get('operacao', '') or '-').strip() or '-'}",
                f"Quantidade na linha: {float(row.get('qtd', 0) or 0):.2f}",
                f"Preco unitario comercial: {_fmt_eur(float(row.get('preco_unit', 0) or 0))}",
                f"Preco unitario com desconto: {_fmt_eur(float(row.get('preco_unit_desconto', row.get('preco_unit', 0)) or 0))}",
            ]
            pdf_docs = [
                str(item or "").strip()
                for item in [row.get("desenho_pdf", ""), *list(row.get("desenhos_pdf", []) or [])]
                if str(item or "").strip()
            ]
            if pdf_docs:
                tooltip_lines.append(f"Documentacao PDF: {len(dict.fromkeys(pdf_docs))} ficheiro(s) associado(s).")
            else:
                tooltip_lines.append("Documentacao PDF: sem PDF associado.")
            if bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False)):
                tooltip_lines.append("Materia-prima: fornecida pelo cliente. A linha considera apenas transformacao/processo.")
            quote_snapshot = dict(row.get("quote_cost_snapshot", {}) or {})
            if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE and str(quote_snapshot.get("costing_mode", "") or "").strip() == "aggregate_pending":
                tooltip_lines.append("Rota multi-operacao: o preco atual da linha ainda esta agregado e nao repartido por posto.")
            elif line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE and str(quote_snapshot.get("costing_mode", "") or "").strip() in {"detailed", "partial_detail"}:
                for op_row in list(row.get("operacoes_detalhe", []) or []):
                    if not isinstance(op_row, dict):
                        continue
                    op_name = str(op_row.get("nome", "") or "").strip()
                    if not op_name:
                        continue
                    tempo_txt = "-" if op_row.get("tempo_unit_min") in (None, "") else f"{float(op_row.get('tempo_unit_min', 0) or 0):.3f} min/un"
                    custo_txt = "-" if op_row.get("custo_unit_eur") in (None, "") else _fmt_eur(float(op_row.get("custo_unit_eur", 0) or 0))
                    tooltip_lines.append(f"{op_name}: {tempo_txt} | {custo_txt}")
            if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE and str(row.get("desenho", "") or "").strip() and "corte laser" in str(row.get("operacao", "") or "").strip().lower():
                if bridge:
                    programmed = int(bridge_row.get("qty", 0) or 0)
                    requested = int(round(float(row.get("qtd", 0) or 0)))
                    tooltip_lines.append(f"Nesting: {programmed}/{requested} peca(s) programada(s) no ultimo cenario.")
                    if bridge_row:
                        tooltip_lines.append(f"Valor comercial colocado no plano: {_fmt_eur(float(bridge_row.get('quoted_total_eur', 0) or 0))}")
                        tooltip_lines.append(
                            f"Rateio matéria real: {_fmt_eur(float(bridge_row.get('allocated_material_total_eur', 0) or 0))} "
                            f"({_fmt_eur(float(bridge_row.get('allocated_material_unit_eur', 0) or 0))}/un)"
                        )
                        tooltip_lines.append(
                            f"Rateio compra: {_fmt_eur(float(bridge_row.get('allocated_purchase_total_eur', 0) or 0))} "
                            f"({_fmt_eur(float(bridge_row.get('allocated_purchase_unit_eur', 0) or 0))}/un)"
                        )
                        tooltip_lines.append(
                            f"Comercial atual: {_fmt_eur(float(bridge_row.get('current_quote_total_eur', 0) or 0))} | "
                            f"ajustado ao plano: {_fmt_eur(float(bridge_row.get('adjusted_quote_total_eur', 0) or 0))}"
                        )
                        tooltip_lines.append(
                            f"Peso no plano: {float(bridge_row.get('share_pct', 0) or 0):.2f}% | "
                            f"delta comercial: {_fmt_eur(float(bridge_row.get('quote_delta_eur', 0) or 0))}"
                        )
                    else:
                        tooltip_lines.append("Esta referencia nao ficou colocada no ultimo plano analisado.")
                else:
                    tooltip_lines.append("Sem plano de chapa aplicado a este orcamento nesta sessao.")
            tooltip = "\n".join(tooltip_lines)
            for col_index in range(self.lines_table.columnCount()):
                item = self.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(tooltip)
            if bridge_row:
                adjusted_total = float(bridge_row.get("adjusted_quote_total_eur", 0) or 0.0)
                current_total = float(row.get("total_desconto", row.get("total", 0)) or 0.0)
                total_item = self.lines_table.item(row_index, self.LINE_COL_TOTAL)
                if total_item is not None and adjusted_total > 0:
                    if current_total + 0.009 < adjusted_total:
                        total_item.setBackground(QBrush(QColor("#fff8eb")))
                        total_item.setForeground(QBrush(QColor("#b45f06")))
                    elif current_total > adjusted_total + 0.009:
                        total_item.setBackground(QBrush(QColor("#dcfce7")))
                        total_item.setForeground(QBrush(QColor("#166534")))
                    else:
                        total_item.setBackground(QBrush(QColor("#fef3c7")))
                        total_item.setForeground(QBrush(QColor("#92400e")))
        self.lines_table.blockSignals(previous_signal_state)
        self.lines_table.setUpdatesEnabled(updates_were_enabled)
        self._rendering_quote_lines = False
        discount_mode = self._quote_discount_mode()
        discount_value = round(sum(float(row.get("desconto_aplicado", 0) or 0) for row in self.line_rows), 2)
        subtotal_discounted = round(sum(float(row.get("total_desconto", row.get("total", 0)) or 0) for row in self.line_rows), 2)
        subtotal_without_iva = max(0.0, subtotal_discounted + transport)
        iva_rate = 23.0
        if abs(float(self.iva_spin.value() or 0) - iva_rate) > 0.0001:
            self.iva_spin.blockSignals(True)
            self.iva_spin.setValue(iva_rate)
            self.iva_spin.blockSignals(False)
        iva_amount = round(subtotal_without_iva * (iva_rate / 100.0), 2)
        total = round(subtotal_without_iva + iva_amount, 2)
        if hasattr(self, "line_count_label"):
            line_count = len(self.line_rows)
            self.line_count_label.setText(f"{line_count} linha" if line_count == 1 else f"{line_count} linhas")
        if hasattr(self, "quote_finance_status_label"):
            if not self.line_rows:
                finance_status = "AGUARDA LINHAS"
            elif discount_pct > 0 and increment_pct > 0:
                finance_status = "PREÇOS AJUSTADOS"
            elif discount_pct > 0:
                finance_status = f"DESCONTO {discount_pct:.1f}%"
            elif increment_pct != 0:
                finance_status = f"INCREMENTO {increment_pct:.1f}%"
            else:
                finance_status = "SEM AJUSTES"
            self.quote_finance_status_label.setText(finance_status.replace(".", ","))
        if hasattr(self, "total_caption_label"):
            self.total_caption_label.setText("Total C/IVA 23%")
        if getattr(self, "iva_summary_row_label", None) is not None:
            self.iva_summary_row_label.setText("IVA 23%")
        self.lines_subtotal_label.setText(_fmt_eur(subtotal))
        self.increment_value_label.setText(_fmt_eur(max(0.0, round(subtotal - subtotal_base, 2))))
        self.transport_total_label.setText(_fmt_eur(transport))
        self.discount_value_label.setText(_fmt_eur(discount_value))
        self.subtotal_without_iva_label.setText(_fmt_eur(subtotal_without_iva))
        self.iva_total_label.setText(_fmt_eur(iva_amount))
        self.total_label.setText(_fmt_eur(total))
        if hasattr(self, "header_total_label"):
            self.header_total_label.setText(_fmt_eur(total))
        self._sync_discount_targets_label()
        discount_breakdown = self._quote_discount_breakdown(selected_discount_keys, discount_pct)
        if discount_mode == "lotes_espessura" and discount_breakdown:
            breakdown_lines = [
                f"{str(group.get('label', '-') or '-').strip()}: -{_fmt_eur(float(group.get('discount', 0) or 0))}"
                for group in discount_breakdown
                if float(group.get("discount", 0) or 0) > 0
            ]
            if breakdown_lines:
                self.discount_breakdown_label.setText("Desconto aplicado por lote:\n" + "\n".join(breakdown_lines))
            else:
                self.discount_breakdown_label.setText("Nenhum dos lotes selecionados tem desconto aplicado.")
        elif discount_pct > 0:
            self.discount_breakdown_label.setText("Desconto aplicado a todas as linhas do orçamento. O transporte mantém-se fora do desconto.")
        else:
            self.discount_breakdown_label.setText("Sem desconto global aplicado.")
        # A tabela ocupa apenas o espaço disponível no cartão e faz o seu próprio
        # scroll. Assim, o rodapé de ações e o limite inferior do cartão ficam
        # sempre visíveis, independentemente do número de referências.
        self.lines_table.setMinimumHeight(280)
        self.lines_table.setMaximumHeight(16777215)
        if hasattr(self, "quote_lines_card"):
            self.quote_lines_card.setMinimumHeight(500)
        self._sync_quote_selected_line_actions()

    def _load_quote(self, numero: str) -> None:
        detail = self.backend.orc_detail(numero)
        client = dict(detail.get("cliente", {}) or {})
        self._set_quote_header(detail.get("numero", ""), detail.get("estado", "Em edicao"), detail.get("numero_encomenda", ""))
        client_label = f"{client.get('codigo', '')} - {client.get('nome', '')}".strip(" -")
        self.client_combo.setCurrentText(client_label or client.get("nome", ""))
        self.client_name_edit.setText(str(client.get("nome", "") or "").strip())
        self.client_company_edit.setText(str(client.get("empresa", "") or client.get("nome", "") or "").strip())
        self.client_nif_edit.setText(str(client.get("nif", "") or "").strip())
        self.client_address_edit.setText(str(client.get("morada", "") or "").strip())
        self.client_contact_edit.setText(str(client.get("contacto", "") or "").strip())
        self.client_email_edit.setText(str(client.get("email", "") or "").strip())
        self._refresh_quote_identity_label()
        self.executed_combo.setCurrentText(str(detail.get("executado_por", "") or "").strip())
        self.workcenter_combo.setCurrentText("")
        self.note_cliente_edit.setText(str(detail.get("nota_cliente", "") or "").strip())
        self.transport_combo.setCurrentText(str(detail.get("nota_transporte", "") or "").strip())
        carrier_txt = " - ".join(
            [
                part
                for part in [
                    str(detail.get("transportadora_id", "") or "").strip(),
                    str(detail.get("transportadora_nome", "") or "").strip(),
                ]
                if part
            ]
        ).strip(" -")
        self.transport_carrier_combo.setCurrentText(carrier_txt)
        self.transport_zone_combo.setCurrentText(str(detail.get("zona_transporte", "") or "").strip())
        self.transport_price_spin.setValue(float(detail.get("preco_transporte", 0) or 0))
        self.discount_spin.setValue(float(detail.get("desconto_perc", 0) or 0))
        self.price_increment_spin.setValue(float(detail.get("incremento_preco_perc", 0) or 0))
        discount_mode = str(detail.get("desconto_modo", "total") or "total").strip().lower()
        self.discount_mode_combo.setCurrentIndex(1 if discount_mode == "lotes_espessura" else 0)
        self.discount_group_keys = [
            str(key or "").strip()
            for key in list(detail.get("desconto_grupos", []) or [])
            if str(key or "").strip()
        ]
        self.notes_edit.setPlainText(str(detail.get("notas_pdf", "") or "").strip())
        delivery_date = str(detail.get("prazo_entrega_data", "") or "").strip()[:10]
        qdate = QDate.fromString(delivery_date, "yyyy-MM-dd") if delivery_date else QDate()
        self.delivery_date_min = QDate.currentDate()
        self.delivery_date_edit.setMinimumDate(self.delivery_date_min)
        self.delivery_date_edit.setDate(qdate if qdate.isValid() and qdate >= self.delivery_date_min else self.delivery_date_min)
        self.iva_spin.setValue(23.0)
        self.nesting_bridge_data = dict(detail.get("nesting_bridge", {}) or {})
        self._refresh_nesting_bridge()
        self.line_rows = [dict(row) for row in list(detail.get("linhas", []) or [])]
        self._recalc_transport_calc()
        self._render_quote_lines()

    def _open_selected_quote(self) -> None:
        row = self._selected_quote_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Orçamentos", "Seleciona um orcamento.")
            return
        try:
            self._load_quote(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._show_detail()

    def _new_quote(self) -> None:
        self._clear_quote_detail()
        self._show_detail()

    def _quote_pick_workcenter(self, *preferred_values: str) -> str:
        options = [str(self.workcenter_combo.itemText(index) or "").strip() for index in range(self.workcenter_combo.count())]
        normalized = {str(option).strip().lower(): option for option in options if str(option).strip()}
        for value in preferred_values:
            key = str(value or "").strip().lower()
            if key and key in normalized:
                return normalized[key]
        return str(options[0] if options else "").strip()

    def _structure_line(self, descricao: str, qtd: float, unid: str, preco_unit: float, operacao: str) -> dict | None:
        quantidade = round(float(qtd or 0), 2)
        preco = round(float(preco_unit or 0), 4)
        if quantidade <= 0 or preco <= 0:
            return None
        return {
            "tipo_item": self.backend.desktop_main.ORC_LINE_TYPE_SERVICE,
            "descricao": str(descricao or "").strip(),
            "produto_unid": str(unid or "SV").strip() or "SV",
            "operacao": str(operacao or "Montagem").strip() or "Montagem",
            "qtd": quantidade,
            "preco_unit": preco,
        }

    def _structure_product_line(self, product: dict | None, qtd: float, *, descricao_extra: str = "") -> dict | None:
        row = dict(product or {})
        code = str(row.get("codigo", "") or "").strip()
        if not code:
            return None
        quantidade = round(float(qtd or 0), 2)
        if quantidade <= 0:
            return None
        preco = round(float(row.get("preco_venda", row.get("pvp1", row.get("preco_unid", row.get("preco", 0)))) or 0), 4)
        if preco <= 0:
            return None
        descricao = str(row.get("descricao", "") or code).strip()
        if descricao_extra:
            descricao = f"{descricao} | {descricao_extra.strip()}"
        return {
            "tipo_item": self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
            "stock_item_kind": "product",
            "produto_codigo": code,
            "produto_unid": str(row.get("unid", "") or "UN").strip() or "UN",
            "descricao": descricao,
            "ref_externa": code,
            "operacao": str(row.get("tipo", "") or "Montagem").strip() or "Montagem",
            "qtd": quantidade,
            "preco_unit": preco,
        }

    def _structure_quote_dialog(self) -> dict | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Orcamento Estruturas Metalicas")
        dialog.resize(1060, 780)
        dialog.setMinimumSize(920, 700)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        family_presets = {
            "Leve": {"profile_kg_m2": 18.0, "tube_kg_m2": 5.5, "frame_spacing_m": 6.0, "fab_h_m2": 0.16, "assembly_h_m2": 0.11, "eng_h_m2": 0.018},
            "Standard": {"profile_kg_m2": 24.0, "tube_kg_m2": 7.5, "frame_spacing_m": 5.0, "fab_h_m2": 0.24, "assembly_h_m2": 0.16, "eng_h_m2": 0.028},
            "Reforcada": {"profile_kg_m2": 31.0, "tube_kg_m2": 9.0, "frame_spacing_m": 4.5, "fab_h_m2": 0.31, "assembly_h_m2": 0.21, "eng_h_m2": 0.04},
        }
        cladding_presets = {
            "Sem revestimento": {"roof_price": 0.0, "facade_price": 0.0, "roof_factor": 0.0, "facade_factor": 0.0, "name": "Sem revestimento"},
            "Chapa simples": {"roof_price": 19.5, "facade_price": 16.5, "roof_factor": 1.0, "facade_factor": 1.0, "name": "Chapa simples"},
            "Sandwich cobertura": {"roof_price": 32.0, "facade_price": 0.0, "roof_factor": 1.0, "facade_factor": 0.0, "name": "Sandwich cobertura"},
            "Sandwich completo": {"roof_price": 34.0, "facade_price": 29.0, "roof_factor": 1.0, "facade_factor": 1.0, "name": "Sandwich completo"},
        }
        finish_presets = {
            "Sem acabamento": {"price_m2": 0.0, "paint_factor": 0.0, "operation": "Serralharia"},
            "Pintura": {"price_m2": 8.5, "paint_factor": 1.15, "operation": "Pintura"},
            "Galvanizacao": {"price_m2": 11.75, "paint_factor": 1.05, "operation": "Lacagem"},
            "Metalizacao + pintura": {"price_m2": 16.2, "paint_factor": 1.25, "operation": "Lacagem"},
        }
        product_rows = [dict(row) for row in list(self.backend.ne_product_options("") or []) if isinstance(row, dict)]
        product_rows_by_code = {
            str(row.get("codigo", "") or "").strip(): row
            for row in product_rows
            if str(row.get("codigo", "") or "").strip()
        }

        intro = QLabel(
            "Versao 2 do modelo de estruturas metalicas. "
            "Gera uma proposta mais tecnica para pavilhoes, coberturas e estruturas especiais, "
            "com familia estrutural, pórticos, revestimento, acabamento, acessorios e produtos reais de stock."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        intro.setStyleSheet("font-size: 11px;")
        layout.addWidget(intro)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_host = QWidget()
        scroll.setWidget(scroll_host)
        scroll_layout = QVBoxLayout(scroll_host)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        scroll_layout.setSpacing(8)
        layout.addWidget(scroll, 1)

        top_grid = QGridLayout()
        top_grid.setHorizontalSpacing(10)
        top_grid.setVerticalSpacing(8)
        scroll_layout.addLayout(top_grid)

        project_card = CardFrame()
        project_card.set_tone("default")
        project_form = QFormLayout(project_card)
        project_form.setContentsMargins(10, 10, 10, 10)
        project_form.setHorizontalSpacing(10)
        project_form.setVerticalSpacing(6)
        project_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        structure_name_edit = QLineEdit()
        structure_name_edit.setPlaceholderText("Ex.: Pavilhao logistico cliente")
        structure_type_combo = QComboBox()
        structure_type_combo.addItems(["Pavilhao industrial", "Cobertura metalica", "Mezanino", "Estrutura especial"])
        family_combo = QComboBox()
        family_combo.addItems(list(family_presets.keys()))
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(1.0, 100.0)
        qty_spin.setDecimals(0)
        qty_spin.setValue(1.0)
        length_spin = QDoubleSpinBox()
        length_spin.setRange(1.0, 500.0)
        length_spin.setDecimals(2)
        length_spin.setSuffix(" m")
        length_spin.setValue(30.0)
        width_spin = QDoubleSpinBox()
        width_spin.setRange(1.0, 200.0)
        width_spin.setDecimals(2)
        width_spin.setSuffix(" m")
        width_spin.setValue(18.0)
        height_spin = QDoubleSpinBox()
        height_spin.setRange(1.0, 50.0)
        height_spin.setDecimals(2)
        height_spin.setSuffix(" m")
        height_spin.setValue(6.0)
        roof_slope_spin = QDoubleSpinBox()
        roof_slope_spin.setRange(0.0, 100.0)
        roof_slope_spin.setDecimals(1)
        roof_slope_spin.setSuffix(" %")
        roof_slope_spin.setValue(12.0)
        frame_spacing_spin = QDoubleSpinBox()
        frame_spacing_spin.setRange(2.0, 12.0)
        frame_spacing_spin.setDecimals(2)
        frame_spacing_spin.setSuffix(" m")
        frame_spacing_spin.setValue(family_presets["Standard"]["frame_spacing_m"])
        project_form.addRow("Designacao", structure_name_edit)
        project_form.addRow("Tipologia", structure_type_combo)
        project_form.addRow("Familia estrutural", family_combo)
        project_form.addRow("Qtd. estruturas", qty_spin)
        project_form.addRow("Comprimento", length_spin)
        project_form.addRow("Largura", width_spin)
        project_form.addRow("Altura util", height_spin)
        project_form.addRow("Inclinacao cobertura", roof_slope_spin)
        project_form.addRow("Espacamento pórticos", frame_spacing_spin)

        cost_card = CardFrame()
        cost_card.set_tone("default")
        cost_form = QFormLayout(cost_card)
        cost_form.setContentsMargins(10, 10, 10, 10)
        cost_form.setHorizontalSpacing(10)
        cost_form.setVerticalSpacing(6)
        cost_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        profile_kg_spin = QDoubleSpinBox()
        profile_kg_spin.setRange(0.0, 500.0)
        profile_kg_spin.setDecimals(2)
        profile_kg_spin.setSuffix(" kg/m2")
        profile_kg_spin.setValue(24.0)
        profile_price_spin = QDoubleSpinBox()
        profile_price_spin.setRange(0.0, 1000.0)
        profile_price_spin.setDecimals(3)
        profile_price_spin.setPrefix("EUR ")
        profile_price_spin.setValue(3.15)
        tube_kg_spin = QDoubleSpinBox()
        tube_kg_spin.setRange(0.0, 500.0)
        tube_kg_spin.setDecimals(2)
        tube_kg_spin.setSuffix(" kg/m2")
        tube_kg_spin.setValue(7.5)
        tube_price_spin = QDoubleSpinBox()
        tube_price_spin.setRange(0.0, 1000.0)
        tube_price_spin.setDecimals(3)
        tube_price_spin.setPrefix("EUR ")
        tube_price_spin.setValue(3.05)
        cladding_combo = QComboBox()
        cladding_combo.addItems(list(cladding_presets.keys()))
        roof_price_spin = QDoubleSpinBox()
        roof_price_spin.setRange(0.0, 1000.0)
        roof_price_spin.setDecimals(2)
        roof_price_spin.setPrefix("EUR ")
        roof_price_spin.setSuffix("/m2")
        roof_price_spin.setValue(cladding_presets["Sandwich completo"]["roof_price"])
        facade_price_spin = QDoubleSpinBox()
        facade_price_spin.setRange(0.0, 1000.0)
        facade_price_spin.setDecimals(2)
        facade_price_spin.setPrefix("EUR ")
        facade_price_spin.setSuffix("/m2")
        facade_price_spin.setValue(cladding_presets["Sandwich completo"]["facade_price"])
        finish_combo = QComboBox()
        finish_combo.addItems(list(finish_presets.keys()))
        paint_factor_spin = QDoubleSpinBox()
        paint_factor_spin.setRange(0.0, 5.0)
        paint_factor_spin.setDecimals(2)
        paint_factor_spin.setValue(finish_presets["Pintura"]["paint_factor"])
        paint_price_spin = QDoubleSpinBox()
        paint_price_spin.setRange(0.0, 1000.0)
        paint_price_spin.setDecimals(2)
        paint_price_spin.setPrefix("EUR ")
        paint_price_spin.setSuffix("/m2")
        paint_price_spin.setValue(finish_presets["Pintura"]["price_m2"])
        cost_form.addRow("Perfis principais", profile_kg_spin)
        cost_form.addRow("Preco perfis", profile_price_spin)
        cost_form.addRow("Tubos / travamentos", tube_kg_spin)
        cost_form.addRow("Preco tubos", tube_price_spin)
        cost_form.addRow("Revestimento", cladding_combo)
        cost_form.addRow("Cobertura / remates", roof_price_spin)
        cost_form.addRow("Fachadas / fechamentos", facade_price_spin)
        cost_form.addRow("Acabamento", finish_combo)
        cost_form.addRow("Fator acabamento", paint_factor_spin)
        cost_form.addRow("Preco acabamento", paint_price_spin)

        labour_card = CardFrame()
        labour_card.set_tone("default")
        labour_form = QFormLayout(labour_card)
        labour_form.setContentsMargins(10, 10, 10, 10)
        labour_form.setHorizontalSpacing(10)
        labour_form.setVerticalSpacing(6)
        labour_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        fab_hours_spin = QDoubleSpinBox()
        fab_hours_spin.setRange(0.0, 5000.0)
        fab_hours_spin.setDecimals(1)
        fab_hours_spin.setSuffix(" h")
        fab_hours_spin.setValue(140.0)
        fab_rate_spin = QDoubleSpinBox()
        fab_rate_spin.setRange(0.0, 1000.0)
        fab_rate_spin.setDecimals(2)
        fab_rate_spin.setPrefix("EUR ")
        fab_rate_spin.setSuffix("/h")
        fab_rate_spin.setValue(28.0)
        assembly_hours_spin = QDoubleSpinBox()
        assembly_hours_spin.setRange(0.0, 5000.0)
        assembly_hours_spin.setDecimals(1)
        assembly_hours_spin.setSuffix(" h")
        assembly_hours_spin.setValue(90.0)
        assembly_rate_spin = QDoubleSpinBox()
        assembly_rate_spin.setRange(0.0, 1000.0)
        assembly_rate_spin.setDecimals(2)
        assembly_rate_spin.setPrefix("EUR ")
        assembly_rate_spin.setSuffix("/h")
        assembly_rate_spin.setValue(32.0)
        engineering_hours_spin = QDoubleSpinBox()
        engineering_hours_spin.setRange(0.0, 1000.0)
        engineering_hours_spin.setDecimals(1)
        engineering_hours_spin.setSuffix(" h")
        engineering_hours_spin.setValue(24.0)
        engineering_rate_spin = QDoubleSpinBox()
        engineering_rate_spin.setRange(0.0, 1000.0)
        engineering_rate_spin.setDecimals(2)
        engineering_rate_spin.setPrefix("EUR ")
        engineering_rate_spin.setSuffix("/h")
        engineering_rate_spin.setValue(35.0)
        crane_days_spin = QDoubleSpinBox()
        crane_days_spin.setRange(0.0, 365.0)
        crane_days_spin.setDecimals(1)
        crane_days_spin.setSuffix(" dias")
        crane_days_spin.setValue(2.0)
        crane_day_rate_spin = QDoubleSpinBox()
        crane_day_rate_spin.setRange(0.0, 1000000.0)
        crane_day_rate_spin.setDecimals(2)
        crane_day_rate_spin.setPrefix("EUR ")
        crane_day_rate_spin.setSuffix("/dia")
        crane_day_rate_spin.setValue(450.0)
        extras_desc_edit = QLineEdit()
        extras_desc_edit.setPlaceholderText("Ex.: portas, caleiras, platibandas, acessorios")
        extras_value_spin = QDoubleSpinBox()
        extras_value_spin.setRange(0.0, 1000000.0)
        extras_value_spin.setDecimals(2)
        extras_value_spin.setPrefix("EUR ")
        labour_form.addRow("Horas de fabrico", fab_hours_spin)
        labour_form.addRow("Preco fabrico", fab_rate_spin)
        labour_form.addRow("Horas de montagem", assembly_hours_spin)
        labour_form.addRow("Preco montagem", assembly_rate_spin)
        labour_form.addRow("Horas engenharia", engineering_hours_spin)
        labour_form.addRow("Preco engenharia", engineering_rate_spin)
        labour_form.addRow("Grua / elevacao", crane_days_spin)
        labour_form.addRow("Preco grua", crane_day_rate_spin)
        labour_form.addRow("Extras", extras_desc_edit)
        labour_form.addRow("Valor extras", extras_value_spin)

        accessory_card = CardFrame()
        accessory_card.set_tone("default")
        accessory_layout = QVBoxLayout(accessory_card)
        accessory_layout.setContentsMargins(10, 10, 10, 10)
        accessory_layout.setSpacing(6)
        accessory_title = QLabel("Acessorios / produtos stock")
        accessory_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        accessory_hint = QLabel("Seleciona produtos reais de stock para incluir na proposta, como parafusaria, componentes ou apoio de montagem.")
        accessory_hint.setProperty("role", "muted")
        accessory_hint.setWordWrap(True)
        accessory_layout.addWidget(accessory_title)
        accessory_layout.addWidget(accessory_hint)
        accessory_grid = QGridLayout()
        accessory_grid.setHorizontalSpacing(8)
        accessory_grid.setVerticalSpacing(6)
        accessory_layout.addLayout(accessory_grid)
        accessory_controls: list[dict[str, object]] = []
        for idx in range(3):
            product_combo = QComboBox()
            product_combo.setEditable(True)
            product_combo.addItem("")
            for row in product_rows:
                code = str(row.get("codigo", "") or "").strip()
                label = f"{code} - {str(row.get('descricao', '') or '').strip()}".strip(" -")
                product_combo.addItem(label, code)
            qty_combo_spin = QDoubleSpinBox()
            qty_combo_spin.setRange(0.0, 1000000.0)
            qty_combo_spin.setDecimals(2)
            qty_combo_spin.setValue(0.0)
            accessory_grid.addWidget(QLabel(f"Produto {idx + 1}"), idx, 0)
            accessory_grid.addWidget(product_combo, idx, 1)
            accessory_grid.addWidget(QLabel("Qtd."), idx, 2)
            accessory_grid.addWidget(qty_combo_spin, idx, 3)
            accessory_controls.append({"combo": product_combo, "qty": qty_combo_spin})

        top_grid.addWidget(project_card, 0, 0)
        top_grid.addWidget(cost_card, 0, 1)
        top_grid.addWidget(labour_card, 1, 0, 1, 2)
        top_grid.addWidget(accessory_card, 2, 0, 1, 2)
        top_grid.setColumnStretch(0, 1)
        top_grid.setColumnStretch(1, 1)

        summary_card = CardFrame()
        summary_card.set_tone("info")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(10, 8, 10, 8)
        summary_layout.setSpacing(5)
        summary_title = QLabel("Resumo tecnico e comercial")
        summary_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        summary_text = QLabel("")
        summary_text.setWordWrap(True)
        summary_text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(summary_text)
        scroll_layout.addWidget(summary_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
        ok_button = buttons.button(QDialogButtonBox.Ok)
        if ok_button is not None:
            ok_button.setText("Gerar orcamento")
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)

        def _apply_family_preset() -> None:
            preset = dict(family_presets.get(family_combo.currentText().strip(), family_presets["Standard"]))
            profile_kg_spin.setValue(float(preset.get("profile_kg_m2", profile_kg_spin.value()) or 0))
            tube_kg_spin.setValue(float(preset.get("tube_kg_m2", tube_kg_spin.value()) or 0))
            frame_spacing_spin.setValue(float(preset.get("frame_spacing_m", frame_spacing_spin.value()) or 0))

        def _apply_cladding_preset() -> None:
            preset = dict(cladding_presets.get(cladding_combo.currentText().strip(), cladding_presets["Sem revestimento"]))
            roof_price_spin.setValue(float(preset.get("roof_price", roof_price_spin.value()) or 0))
            facade_price_spin.setValue(float(preset.get("facade_price", facade_price_spin.value()) or 0))

        def _apply_finish_preset() -> None:
            preset = dict(finish_presets.get(finish_combo.currentText().strip(), finish_presets["Sem acabamento"]))
            paint_factor_spin.setValue(float(preset.get("paint_factor", paint_factor_spin.value()) or 0))
            paint_price_spin.setValue(float(preset.get("price_m2", paint_price_spin.value()) or 0))

        def _compute_payload() -> dict:
            structure_name = structure_name_edit.text().strip() or structure_type_combo.currentText().strip() or "Estrutura metalica"
            assembly_code = f"EST-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            quantity = float(qty_spin.value() or 1.0)
            length = float(length_spin.value() or 0.0)
            width = float(width_spin.value() or 0.0)
            height = float(height_spin.value() or 0.0)
            slope_pct = float(roof_slope_spin.value() or 0.0)
            frame_spacing = max(2.0, float(frame_spacing_spin.value() or 0.0))
            family_name = family_combo.currentText().strip() or "Standard"
            family_preset = dict(family_presets.get(family_name, family_presets["Standard"]))
            cladding_preset = dict(cladding_presets.get(cladding_combo.currentText().strip(), cladding_presets["Sem revestimento"]))
            finish_name = finish_combo.currentText().strip() or "Sem acabamento"
            finish_preset = dict(finish_presets.get(finish_name, finish_presets["Sem acabamento"]))
            footprint_area = round(length * width * quantity, 2)
            roof_factor = math.sqrt(1.0 + ((slope_pct / 100.0) ** 2))
            roof_area = round(length * width * roof_factor * quantity * float(cladding_preset.get("roof_factor", 1.0) or 0.0), 2)
            perimeter = round((2.0 * (length + width)) * quantity, 2)
            facade_area = round(perimeter * height * float(cladding_preset.get("facade_factor", 1.0) or 0.0), 2)
            frame_count_single = max(2, int(math.ceil(length / frame_spacing)) + 1)
            portal_count = int(frame_count_single * quantity)
            column_count = int(portal_count * 2)
            steel_profiles_kg = round(footprint_area * float(profile_kg_spin.value() or 0.0), 2)
            steel_tubes_kg = round(footprint_area * float(tube_kg_spin.value() or 0.0), 2)
            finish_area = round((roof_area + facade_area) * float(paint_factor_spin.value() or 0.0), 2)
            fab_hours = round(max(float(fab_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("fab_h_m2", 0.0) or 0.0)), 1)
            assembly_hours = round(max(float(assembly_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("assembly_h_m2", 0.0) or 0.0)), 1)
            engineering_hours = round(max(float(engineering_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("eng_h_m2", 0.0) or 0.0)), 1)

            lines: list[dict] = []
            for row in (
                self._structure_line(f"Perfis estruturais {family_name} | {structure_name}", steel_profiles_kg, "kg", float(profile_price_spin.value() or 0.0), "Serralharia"),
                self._structure_line(f"Tubos e travamentos | {structure_name}", steel_tubes_kg, "kg", float(tube_price_spin.value() or 0.0), "Serralharia"),
                self._structure_line(f"Cobertura e remates | {structure_name}", roof_area, "m2", float(roof_price_spin.value() or 0.0), "Montagem"),
                self._structure_line(f"Fachadas e fechamentos | {structure_name}", facade_area, "m2", float(facade_price_spin.value() or 0.0), "Montagem"),
                self._structure_line(f"{finish_name} / protecao estrutural | {structure_name}", finish_area, "m2", float(paint_price_spin.value() or 0.0), str(finish_preset.get("operation", "Pintura") or "Pintura")),
                self._structure_line(f"Fabrico e soldadura | {structure_name}", fab_hours, "h", float(fab_rate_spin.value() or 0.0), "Serralharia"),
                self._structure_line(f"Montagem em obra | {structure_name}", assembly_hours, "h", float(assembly_rate_spin.value() or 0.0), "Montagem"),
                self._structure_line(f"Engenharia e preparacao | {structure_name}", engineering_hours, "h", float(engineering_rate_spin.value() or 0.0), "Serralharia"),
                self._structure_line(f"Grua / elevacao | {structure_name}", float(crane_days_spin.value() or 0.0), "dia", float(crane_day_rate_spin.value() or 0.0), "Montagem"),
            ):
                if isinstance(row, dict):
                    lines.append(row)

            for control in accessory_controls:
                combo = control.get("combo")
                qty_widget = control.get("qty")
                if not isinstance(combo, QComboBox) or not isinstance(qty_widget, QDoubleSpinBox):
                    continue
                code = str(combo.currentData() or "").strip()
                if not code:
                    text = combo.currentText().strip()
                    code = text.split(" - ", 1)[0].strip()
                product_line = self._structure_product_line(
                    product_rows_by_code.get(code),
                    float(qty_widget.value() or 0.0),
                    descricao_extra=structure_name,
                )
                if isinstance(product_line, dict):
                    lines.append(product_line)

            extras_value = float(extras_value_spin.value() or 0.0)
            extras_desc = extras_desc_edit.text().strip() or "Extras de estrutura"
            extra_line = self._structure_line(f"{extras_desc} | {structure_name}", 1.0, "SV", extras_value, "Montagem")
            if isinstance(extra_line, dict):
                lines.append(extra_line)

            estimated_total = round(sum(float(row.get("total", 0) or 0.0) for row in lines), 2)
            note_cliente = structure_name
            note_lines = [
                f"Modelo estrutura: {structure_type_combo.currentText().strip()}",
                f"Familia: {family_name} | acabamento: {finish_name} | revestimento: {cladding_preset.get('name', cladding_combo.currentText().strip())}",
                f"Dimensoes base: {length:.2f} x {width:.2f} x {height:.2f} m | quantidade {quantity:.0f}",
                f"Porticos: {portal_count} | colunas: {column_count} | espacamento medio: {frame_spacing:.2f} m",
                f"Area implantacao: {footprint_area:.2f} m2 | cobertura: {roof_area:.2f} m2 | fachadas: {facade_area:.2f} m2",
                f"Perfis: {steel_profiles_kg:.2f} kg | tubos: {steel_tubes_kg:.2f} kg | acabamento: {finish_area:.2f} m2",
            ]
            return {
                "name": structure_name,
                "type": structure_type_combo.currentText().strip(),
                "assembly_code": assembly_code,
                "assembly_name": structure_name,
                "workcenter": self._quote_pick_workcenter("Serralharia", "Montagem"),
                "note_cliente": note_cliente,
                "notes_pdf": "\n".join(note_lines),
                "summary_html": (
                    f"{structure_name} | {structure_type_combo.currentText().strip()} | familia {family_name} | "
                    f"porticos {portal_count} | implantacao {footprint_area:.2f} m2 | cobertura {roof_area:.2f} m2 | "
                    f"fachadas {facade_area:.2f} m2 | perfis {steel_profiles_kg:.2f} kg | tubos {steel_tubes_kg:.2f} kg | "
                    f"acabamento {finish_name} | total base {_fmt_eur(estimated_total)}"
                ),
                "lines": lines,
            }

        def _refresh_summary() -> None:
            payload = _compute_payload()
            lines = list(payload.get("lines", []) or [])
            if not lines:
                summary_text.setText("Preenche valores comerciais para gerar uma base de orcamento.")
                return
            summary_rows = [
                str(payload.get("summary_html", "") or "").strip(),
                f"Linhas geradas: {len(lines)}",
                "Conjunto gerado como mini-projeto: materiais, mao de obra, consumiveis/produtos stock e extras.",
                "Categorias: perfis, tubos, cobertura, fachadas, acabamento, fabrico, montagem, engenharia, grua, extras e produtos de stock.",
                "Depois de gerar, continuas com acesso total ao orcamento normal para acrescentar pecas, produtos ou linhas manuais.",
            ]
            summary_text.setText("\n".join([row for row in summary_rows if row]))

        family_combo.currentTextChanged.connect(lambda _text: _apply_family_preset())
        cladding_combo.currentTextChanged.connect(lambda _text: _apply_cladding_preset())
        finish_combo.currentTextChanged.connect(lambda _text: _apply_finish_preset())
        for widget in (
            structure_name_edit,
            structure_type_combo,
            family_combo,
            qty_spin,
            length_spin,
            width_spin,
            height_spin,
            roof_slope_spin,
            frame_spacing_spin,
            profile_kg_spin,
            profile_price_spin,
            tube_kg_spin,
            tube_price_spin,
            cladding_combo,
            roof_price_spin,
            facade_price_spin,
            finish_combo,
            paint_factor_spin,
            paint_price_spin,
            fab_hours_spin,
            fab_rate_spin,
            assembly_hours_spin,
            assembly_rate_spin,
            engineering_hours_spin,
            engineering_rate_spin,
            crane_days_spin,
            crane_day_rate_spin,
            extras_desc_edit,
            extras_value_spin,
        ):
            if isinstance(widget, (QLineEdit, QComboBox)):
                signal = widget.textChanged if isinstance(widget, QLineEdit) else widget.currentTextChanged
                signal.connect(_refresh_summary)
            elif isinstance(widget, QDoubleSpinBox):
                widget.valueChanged.connect(lambda _value: _refresh_summary())
        for control in accessory_controls:
            combo = control.get("combo")
            qty_widget = control.get("qty")
            if isinstance(combo, QComboBox):
                combo.currentTextChanged.connect(lambda _text: _refresh_summary())
            if isinstance(qty_widget, QDoubleSpinBox):
                qty_widget.valueChanged.connect(lambda _value: _refresh_summary())
        _apply_family_preset()
        _apply_cladding_preset()
        _apply_finish_preset()
        _refresh_summary()

        if dialog.exec() != QDialog.Accepted:
            return None
        payload = _compute_payload()
        if not list(payload.get("lines", []) or []):
            QMessageBox.warning(self, "Orcamento Estruturas", "Nao foi gerada nenhuma linha para o modelo de estruturas.")
            return None
        return payload

    def _apply_structure_quote_payload(self, payload: dict, *, replace_existing: bool) -> None:
        self._apply_group_payload(payload, replace_existing=replace_existing)

    def _apply_group_payload(self, payload: dict, *, replace_existing: bool) -> None:
        if not isinstance(payload, dict):
            return
        assembly_code = str(payload.get("assembly_code", "") or "").strip()
        assembly_name = str(payload.get("assembly_name", payload.get("name", "")) or "").strip()
        group_uuid = f"{assembly_code}-01" if assembly_code else f"EST-GRP-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        lines = []
        for raw_row in list(payload.get("lines", []) or []):
            if not isinstance(raw_row, dict):
                continue
            row = dict(raw_row)
            if assembly_code:
                row["conjunto_codigo"] = assembly_code
            if assembly_name:
                row["conjunto_nome"] = assembly_name
            row["grupo_uuid"] = group_uuid
            lines.append(row)
        if not lines:
            return
        if replace_existing:
            self.line_rows = lines
        else:
            self.line_rows.extend(lines)
        note_cliente = str(payload.get("note_cliente", "") or "").strip()
        if note_cliente and (replace_existing or not self.note_cliente_edit.text().strip()):
            self.note_cliente_edit.setText(note_cliente)
        notes_pdf = str(payload.get("notes_pdf", "") or "").strip()
        if notes_pdf:
            current_notes = [] if replace_existing else [row.strip() for row in self.notes_edit.toPlainText().splitlines() if row.strip()]
            for row in notes_pdf.splitlines():
                row_txt = str(row or "").strip()
                if row_txt and row_txt not in current_notes:
                    current_notes.append(row_txt)
            self.notes_edit.setPlainText("\n".join(current_notes))
        workcenter = str(payload.get("workcenter", "") or "").strip()
        if workcenter:
            self.workcenter_combo.setCurrentText(workcenter)
        self._render_quote_lines()
        self._show_detail()

    def _material_price_manager_dialog(self, formato_filter: str = "", preferred_id: str = "", parent: QWidget | None = None) -> dict | None:
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
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
            rows = list(self.backend.material_price_rows(str(format_combo.currentData() or "").strip()) or [])
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
                    self.backend.material_update_price_kg(str(row.get("id", "") or "").strip(), float(base_spin.value() or 0.0) / kg_m)
                else:
                    self.backend.material_update_price_kg(str(row.get("id", "") or "").strip(), base_spin.value())
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

    def _sync_stock_price_from_context(self, material_id: str, price_kg: float, current_price_kg: float, parent: QWidget | None = None) -> dict[str, Any] | None:
        stock_id = str(material_id or "").strip()
        if not stock_id:
            return None
        new_value = round(float(price_kg or 0.0), 4)
        current_value = round(float(current_price_kg or 0.0), 4)
        if new_value <= 0 or abs(new_value - current_value) < 0.0001:
            return None
        try:
            return dict(self.backend.material_update_price_kg(stock_id, new_value) or {})
        except Exception as exc:
            QMessageBox.warning(parent if isinstance(parent, QWidget) else self, "Matéria-prima", str(exc))
            return None

    def _material_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        initial = dict(initial or {})
        def _initial_float(value: Any, default: float = 0.0) -> float:
            try:
                return float(str(value).strip().replace(" ", "").replace(",", "."))
            except Exception:
                return float(default)

        if isinstance(initial.get("line"), dict):
            wrapper = dict(initial)
            line_payload = dict(wrapper.get("line") or {})
            merged = dict(line_payload)
            for key, value in wrapper.items():
                if key == "line":
                    continue
                if merged.get(key) in (None, "", [], {}):
                    merged[key] = value
            initial = merged

        allowed_modes = {"Perfil", "Tubo", "Chapa", "Cantoneira", "Barra", "Ferro nervurado", "Manual"}
        initial_desc_text = str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip()
        if not str(initial.get("descricao_base", "") or "").strip() and " | " in initial_desc_text:
            initial["descricao_base"] = initial_desc_text.split(" | ", 1)[0].strip()
        initial_mode = str(initial.get("calc_mode", "") or initial.get("material_subtype", "") or "").strip()
        initial_probe = self.backend.desktop_main.norm_text(
            " ".join(
                str(initial.get(key, "") or "")
                for key in ("descricao_base", "descricao", "material", "material_family", "material_subtype", "calc_mode")
            )
        )
        if initial_mode not in allowed_modes:
            initial_mode = ""
        if not initial_mode or (initial_mode == "Perfil" and any(token in initial_probe for token in ("barra", "chata", "plat", "flat"))):
            if any(token in initial_probe for token in ("barra", "chata", "plat", "flat")):
                initial_mode = "Barra"
            elif "cantoneira" in initial_probe or re.search(r"\bl\s*\d+\s*x\s*\d+", initial_probe):
                initial_mode = "Cantoneira"
            elif "tubo" in initial_probe:
                initial_mode = "Tubo"
            elif "chapa" in initial_probe:
                initial_mode = "Chapa"
            elif "ferro nervurado" in initial_probe or "varao nervurado" in initial_probe:
                initial_mode = "Ferro nervurado"
        initial["calc_mode"] = initial_mode or "Perfil"

        desc_for_metrics = str(initial.get("descricao", initial_desc_text) or "")
        if "quantity_units" not in initial and initial.get("qtd") not in (None, ""):
            initial["quantity_units"] = initial.get("qtd")
        if _initial_float(initial.get("meters_per_unit", 0), 0.0) <= 0:
            match_meters = re.search(r"x\s*(\d+(?:[.,]\d+)?)\s*m\b", desc_for_metrics, flags=re.IGNORECASE)
            if match_meters:
                initial["meters_per_unit"] = float(match_meters.group(1).replace(",", "."))
        if _initial_float(initial.get("kg_per_m", 0), 0.0) <= 0:
            match_kg_m = re.search(r"(\d+(?:[.,]\d+)?)\s*kg\s*/\s*m", desc_for_metrics, flags=re.IGNORECASE)
            if match_kg_m:
                initial["kg_per_m"] = float(match_kg_m.group(1).replace(",", "."))
        if _initial_float(initial.get("price_per_kg", 0), 0.0) <= 0:
            match_price_kg = re.search(r"(\d+(?:[.,]\d+)?)\s*eur\s*/\s*kg", desc_for_metrics, flags=re.IGNORECASE)
            if match_price_kg:
                initial["price_per_kg"] = float(match_price_kg.group(1).replace(",", "."))
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
        dialog.setWindowTitle("Item material")
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)
        dialog.setSizeGripEnabled(True)
        try:
            screen = QApplication.screenAt(dialog.parentWidget().mapToGlobal(dialog.parentWidget().rect().center())) if dialog.parentWidget() is not None else QApplication.primaryScreen()
            available = screen.availableGeometry() if screen is not None else None
            if available is not None:
                dialog.resize(min(980, max(820, available.width() - 100)), min(760, max(560, available.height() - 100)))
                dialog.setMaximumHeight(max(520, available.height() - 60))
            else:
                dialog.resize(940, 720)
        except Exception:
            dialog.resize(940, 720)
        dialog.setMinimumSize(760, 520)
        root_layout = QVBoxLayout(dialog)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(8)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        scroll.setWidget(body)
        root_layout.addWidget(scroll, 1)

        top_form = QFormLayout()
        top_form.setContentsMargins(0, 0, 0, 0)
        top_form.setHorizontalSpacing(10)
        top_form.setVerticalSpacing(8)

        mode_combo = QComboBox()
        mode_combo.addItems(["Perfil", "Tubo", "Chapa", "Cantoneira", "Barra", "Ferro nervurado", "Manual"])
        mode_combo.setCurrentText(str(initial.get("calc_mode", "Perfil") or "Perfil"))
        desc_edit = QLineEdit(str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip())
        ref_edit = QLineEdit(str(initial.get("ref_externa", "") or "").strip())
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(float(initial.get("quantity_units", initial.get("qtd", 1)) or 1))
        top_form.addRow("Modo", mode_combo)
        top_form.addRow("Descricao", desc_edit)
        top_form.addRow("Codigo/Ref.", ref_edit)
        top_form.addRow("Quantidade", qty_spin)
        layout.addLayout(top_form)

        detail_intro = QLabel(
            "Define o material do conjunto com origem em catálogo, stock de matéria-prima ou criação manual. "
            "A descrição e a referência passam a acompanhar a dimensão, a qualidade e o formato escolhido."
        )
        detail_intro.setWordWrap(True)
        detail_intro.setProperty("role", "muted")
        layout.addWidget(detail_intro)

        stack = QStackedWidget()
        stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        stack.setMaximumHeight(360)
        layout.addWidget(stack)

        summary_card = CardFrame()
        summary_card.set_tone("info")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(10, 8, 10, 8)
        summary_layout.setSpacing(4)
        summary_weight = QLabel("")
        summary_cost = QLabel("")
        summary_desc = QLabel("")
        summary_desc.setWordWrap(True)
        summary_layout.addWidget(summary_weight)
        summary_layout.addWidget(summary_cost)
        summary_layout.addWidget(summary_desc)
        layout.addWidget(summary_card)

        tool_row = QHBoxLayout()
        price_table_btn = QPushButton("Tabela precos MP")
        price_table_btn.setProperty("variant", "secondary")
        tool_row.addWidget(price_table_btn)
        tool_row.addStretch(1)
        layout.addLayout(tool_row)

        operation_meta = {
            "operacoes_lista": list(initial.get("operacoes_lista", []) or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(initial.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(initial.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(initial.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(initial.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(initial.get("quote_cost_snapshot", {}) or {}),
        }
        operation_selector, operation_edit, apply_operations = _build_operation_selector(
            list(self.presets.get("operacoes", []) or []),
            " + ".join(list(operation_meta.get("operacoes_lista", []) or [])) or str(initial.get("operacao", "") or ""),
        )
        operation_cost_label = QLabel("")
        operation_cost_label.setWordWrap(True)
        operation_cost_label.setProperty("role", "muted")
        operation_buttons = QHBoxLayout()
        operation_buttons.setContentsMargins(0, 0, 0, 0)
        operation_buttons.setSpacing(8)
        op_detail_btn = QPushButton("Quantificar operações")
        op_detail_btn.setProperty("variant", "secondary")
        op_profiles_btn = QPushButton("Perfis operações")
        op_profiles_btn.setProperty("variant", "secondary")
        operation_buttons.addWidget(op_detail_btn)
        operation_buttons.addWidget(op_profiles_btn)
        operation_buttons.addStretch(1)
        operation_buttons_host = QWidget()
        operation_buttons_host.setLayout(operation_buttons)
        op_form = QFormLayout()
        op_form.setContentsMargins(0, 0, 0, 0)
        op_form.setHorizontalSpacing(10)
        op_form.setVerticalSpacing(6)
        op_form.addRow("Operações seguintes", operation_selector)
        op_form.addRow("Custeio op.", operation_cost_label)
        op_form.addRow("", operation_buttons_host)
        root_layout.addLayout(op_form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        root_layout.addWidget(buttons)

        desc_state = {"manual": bool(desc_edit.text().strip()), "last_auto": ""}
        ref_state = {"manual": bool(ref_edit.text().strip()), "last_auto": ""}
        sync = {"busy": False}

        profile_options = [dict(row or {}) for row in list(self.backend.material_section_options("Perfil") or [])]
        tube_options = [dict(row or {}) for row in list(self.backend.material_section_options("Tubo") or [])]
        stock_rows = [dict(item.get("record") or {}) for item in list(self.backend.material_rows("") or []) if isinstance(item, dict) and isinstance(item.get("record"), dict)]
        family_options = [
            dict(row or {})
            for row in list(self.backend.material_family_options() or [])
            if str((row or {}).get("key", "") or "").strip()
        ]
        if not family_options:
            family_options = [
                {"key": "steel", "label": "Ferro / Aço", "density": 7.85},
                {"key": "stainless", "label": "Inox", "density": 7.93},
                {"key": "aluminum", "label": "Alumínio", "density": 2.7},
                {"key": "brass", "label": "Latão", "density": 8.5},
                {"key": "copper", "label": "Cobre", "density": 8.96},
            ]
        family_quality_presets = {
            "steel": ["S235JR", "S275JR", "S355JR", "DC01", "DD11"],
            "stainless": ["Inox 304L", "Inox 316L", "Inox 430"],
            "aluminum": ["Alumínio 5754", "Alumínio 5083", "Alumínio 6082", "Alumínio 1050"],
            "brass": ["Latão"],
            "copper": ["Cobre"],
        }

        equal_angle_presets: dict[str, list[str]] = {
            "30": ["3", "4"],
            "35": ["3", "4"],
            "40": ["3", "4", "5"],
            "45": ["4", "5"],
            "50": ["4", "5", "6"],
            "60": ["5", "6", "8"],
            "65": ["5", "6", "8"],
            "70": ["6", "7", "8"],
            "75": ["6", "8"],
            "80": ["6", "8", "10"],
            "90": ["8", "9", "10"],
            "100": ["8", "10", "12"],
            "120": ["10", "12"],
        }
        flat_bar_presets = [
            "20x3",
            "25x3",
            "25x5",
            "30x3",
            "30x5",
            "40x5",
            "40x6",
            "50x5",
            "50x6",
            "60x6",
            "60x8",
            "70x8",
            "70x10",
            "70x20",
            "80x8",
            "80x10",
            "100x8",
            "100x10",
            "100x12",
            "120x10",
            "120x12",
        ]
        initial_bar_size = str(initial.get("bar_size", "") or "").strip()
        if not initial_bar_size:
            desc_for_bar = str(initial.get("descricao_base", initial.get("descricao", "")) or "")
            match_bar = re.search(r"(\d+(?:[.,]\d+)?)\s*[xX]\s*(\d+(?:[.,]\d+)?)", desc_for_bar)
            if match_bar:
                initial_bar_size = f"{match_bar.group(1)}x{match_bar.group(2)}".replace(",", ".")
        if initial_bar_size and initial_bar_size not in flat_bar_presets:
            flat_bar_presets.append(initial_bar_size)
        tube_model_presets: dict[str, list[str]] = {
            "redondo": ["Ø33.7x2.6", "Ø42.4x3.2", "Ø48.3x3.2", "Ø60.3x3.2", "Ø76.1x3.2", "Ø88.9x4.0", "Ø114.3x4.0"],
            "quadrado": ["30x30x2.0", "40x40x2.0", "40x40x3.0", "50x50x2.5", "60x60x3.0", "80x80x4.0", "100x100x5.0"],
            "retangular": ["40x20x2.0", "50x30x2.5", "60x40x3.0", "80x40x3.0", "100x50x4.0", "120x60x4.0"],
        }
        sheet_model_presets = [
            "2000x1000x2",
            "Chapa gota 2000x1000x3/5",
            "Chapa gota 2000x1000x4/6",
            "2000x1000x5/7",
            "2000x1000x6/8",
            "2000x1000x8/10",
            "Chapa gota 2500x1250x3/5",
            "Chapa gota 2500x1250x4/6",
            "Chapa gota 2500x1250x5/7",
            "Chapa gota 2500x1250x6/8",
            "Chapa gota 2500x1250x8/10",
            "2500x1250x3",
            "2500x1250x5",
            "3000x1500x3",
            "Chapa gota 3000x1500x3/5",
            "Chapa gota 3000x1500x4/6",
            "Chapa gota 3000x1500x5/7",
            "Chapa gota 3000x1500x6/8",
            "Chapa gota 3000x1500x8/10",
            "Chapa gota 3000x1500x10/12",
            "3000x1500x5",
            "3000x1500x10",
            "3000x1500x15",
            "4000x2000x5",
            "4000x2000x8",
            "4000x2000x10",
        ]

        def _parse_size_value(value: object) -> tuple[str, float]:
            text = str(value or "").strip()
            if not text:
                return "", 0.0
            match = re.search(r"(\d+(?:[.,]\d+)?)", text)
            if not match:
                return text, 0.0
            try:
                return text, float(match.group(1).replace(",", "."))
            except Exception:
                return text, 0.0

        def _parse_pair(value: object) -> tuple[float, float]:
            text = str(value or "").lower().replace(",", ".").replace("mm", "").strip()
            match = re.search(r"(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)", text)
            if not match:
                return 0.0, 0.0
            try:
                return float(match.group(1)), float(match.group(2))
            except Exception:
                return 0.0, 0.0

        def _sheet_model_info(value: object) -> dict[str, Any]:
            raw = str(value or "").strip()
            text = raw.lower().replace(",", ".").replace("mm", "")
            is_checker = str(sheet_kind_combo.currentData() or "").strip() == "checker" or any(
                token in self.backend.desktop_main.norm_text(text) for token in ("gota", "xadrez", "antiderrap")
            )
            slash_match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
            numbers = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", text)]
            length = width = thickness = 0.0
            thickness_label = ""
            if len(numbers) >= 3:
                length = numbers[0]
                width = numbers[1]
                thickness = numbers[2]
            elif is_checker and slash_match:
                length = float(length_mm_spin.value() or 3000.0)
                width = float(width_mm_spin.value() or 1500.0)
                thickness = float(slash_match.group(1))
            if slash_match:
                left_raw = slash_match.group(1)
                right_raw = slash_match.group(2)
                left = left_raw.rstrip("0").rstrip(".") if "." in left_raw else left_raw
                right = right_raw.rstrip("0").rstrip(".") if "." in right_raw else right_raw
                thickness_label = f"{left}/{right}"
                thickness = float(slash_match.group(1))
                is_checker = True
            elif thickness > 0:
                thickness_label = _float_text(thickness, 2).replace(",00", "").replace(",0", "")
            return {
                "raw": raw,
                "length": length,
                "width": width,
                "thickness": thickness,
                "thickness_label": thickness_label,
                "is_checker": is_checker,
            }

        def _sort_dimension_labels(values: list[str]) -> list[str]:
            def _sort_key(value: str) -> tuple[float, float, float, str]:
                text = str(value or "").lower().replace("ø", "").replace("mm", "").replace(",", ".").strip()
                parts = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", text)]
                while len(parts) < 3:
                    parts.append(0.0)
                return (parts[0], parts[1], parts[2], text)

            return sorted({str(value or "").strip() for value in values if str(value or "").strip()}, key=_sort_key)

        def _add_combo_item_if_missing(combo: QComboBox, value: object) -> None:
            text = str(value or "").strip()
            if not text:
                return
            for idx in range(combo.count()):
                if str(combo.itemText(idx) or "").strip().lower() == text.lower():
                    return
            combo.addItem(text)

        def _set_combo_text_preserving_custom(combo: QComboBox, value: object) -> None:
            text = str(value or "").strip()
            if not text:
                return
            _add_combo_item_if_missing(combo, text)
            combo.setCurrentText(text)

        def _float_text(value: float, digits: int = 2) -> str:
            return f"{float(value or 0):.{digits}f}".replace(".", ",")

        def _set_edit_if_auto(edit: QLineEdit, state: dict, value: str) -> None:
            current = edit.text().strip()
            if state["manual"] and current and current != state["last_auto"]:
                return
            edit.blockSignals(True)
            edit.setText(value)
            edit.blockSignals(False)
            state["last_auto"] = value

        def _stock_price_kg(record: dict, preview: dict) -> float:
            base_value = float(record.get("p_compra", 0) or 0.0)
            if base_value <= 0:
                base_value = float(preview.get("p_compra", 0) or 0.0)
            if base_value <= 0:
                total_unit = float(preview.get("preco_unid", record.get("preco_unid", 0)) or 0.0)
                total_weight = float(preview.get("peso_unid", record.get("peso_unid", 0)) or 0.0)
                if total_unit > 0 and total_weight > 0:
                    return round(total_unit / total_weight, 4)
                return 0.0
            if str(preview.get("base_label", "") or "").strip().upper() == "EUR/M":
                kg_m = float(preview.get("kg_m", 0) or 0.0)
                if kg_m > 0:
                    return round(base_value / kg_m, 4)
            return round(base_value, 4)

        def _stock_price_base(record: dict, preview: dict) -> tuple[str, float]:
            base_label = str(preview.get("base_label", "") or "").strip()
            if not base_label:
                base_label = "EUR/m" if str(preview.get("formato", record.get("formato", "")) or "").strip() == "Tubo" else "EUR/kg"
            base_value = float(record.get("p_compra", 0) or preview.get("p_compra", 0) or 0.0)
            if base_value <= 0:
                kg_price = _stock_price_kg(record, preview)
                if base_label.strip().upper() == "EUR/M":
                    kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
                    return base_label, round(kg_price * kg_m, 4) if kg_m > 0 else 0.0
                return base_label, kg_price
            return base_label, round(base_value, 4)

        def _stock_mode(record: dict, preview: dict) -> str:
            formato_txt = str(preview.get("formato", record.get("formato", "")) or "").strip().title()
            secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
            material_txt = self.backend.desktop_main.norm_text(str(record.get("material", "") or ""))
            if formato_txt == "Cantoneira":
                return "Cantoneira"
            if formato_txt == "Barra":
                return "Barra"
            if formato_txt == "Perfil" and secao == "L":
                return "Cantoneira"
            if formato_txt == "Perfil":
                return "Perfil"
            if formato_txt == "Tubo":
                return "Tubo"
            if formato_txt == "Chapa":
                if any(token in material_txt for token in ("barra", "chata", "plat", "flat")):
                    return "Barra"
                return "Chapa"
            return ""

        def _stock_label(record: dict, preview: dict) -> str:
            material_txt = str(record.get("material", "") or "").strip()
            dim_txt = str(preview.get("dimension_label", "") or "-").strip()
            esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
            lote_txt = str(record.get("id", record.get("lote_fornecedor", "")) or "").strip()
            qty_txt = self.backend._fmt(record.get("quantidade", 0))
            price_txt = _fmt_eur(_stock_price_kg(record, preview))
            parts = [part for part in (lote_txt, material_txt, dim_txt) if part and part != "-"]
            if esp_txt:
                parts.append(f"{esp_txt} mm")
            if qty_txt:
                parts.append(f"Qtd {qty_txt}")
            if price_txt:
                base_label, base_value = _stock_price_base(record, preview)
                parts.append(f"Base {_fmt_eur(base_value)}/{base_label.split('/')[-1]}")
            return " | ".join(parts) or material_txt or lote_txt or "Stock"

        stock_cache: dict[str, list[tuple[dict, dict]]] = {}

        def _stock_options(mode: str) -> list[tuple[dict, dict]]:
            key = str(mode or "").strip()
            if key in stock_cache:
                return list(stock_cache[key])
            rows: list[tuple[dict, dict]] = []
            for record in stock_rows:
                preview = dict(self.backend.material_price_preview(record) or {})
                if _stock_mode(record, preview) == key:
                    rows.append((record, preview))
            stock_cache[key] = rows
            return list(rows)

        def _fill_stock_combo(combo: QComboBox, mode: str, preferred_id: str = "") -> None:
            current_id = preferred_id or str(combo.currentData() or {}).strip() if isinstance(combo.currentData(), str) else preferred_id
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("", None)
            for record, preview in _stock_options(mode):
                combo.addItem(_stock_label(record, preview), dict(record))
            if preferred_id:
                for idx in range(combo.count()):
                    payload = combo.itemData(idx)
                    if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == preferred_id:
                        combo.setCurrentIndex(idx)
                        break
            combo.blockSignals(False)

        def _current_stock(combo: QComboBox) -> tuple[dict, dict]:
            payload = combo.currentData()
            if not isinstance(payload, dict):
                return {}, {}
            preview = dict(self.backend.material_price_preview(payload) or {})
            return dict(payload), preview

        def _set_combo_by_data(combo: QComboBox, wanted: str) -> None:
            wanted_txt = str(wanted or "").strip().lower()
            if not wanted_txt:
                return
            for idx in range(combo.count()):
                current_txt = str(combo.itemData(idx) or "").strip().lower()
                if current_txt == wanted_txt:
                    combo.setCurrentIndex(idx)
                    return

        def _profile_series_and_size(record: dict, preview: dict) -> tuple[str, str]:
            secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
            dimension = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
            if secao and dimension.upper().startswith(f"{secao} "):
                dimension = dimension[len(secao) :].strip()
            if not dimension:
                dimension = str(record.get("altura", "") or record.get("espessura", "") or "").strip()
            return secao, dimension

        def _geometry_weight_kg_m(area_mm2: float, density: float) -> float:
            return round((max(0.0, float(area_mm2 or 0.0)) * max(0.0, float(density or 0.0))) / 1000.0, 4)

        def _quality_from_record(record: dict) -> str:
            material_txt = str(record.get("material", "") or "").strip()
            familia_txt = str(record.get("material_familia", "") or "").strip()
            if material_txt:
                return material_txt
            if familia_txt:
                return familia_txt
            return "S235JR"

        def _stock_material_identity(mode: str, record: dict, preview: dict) -> dict[str, str]:
            mode_txt = str(mode or "").strip()
            quality = _quality_from_record(record)
            dimension = str(preview.get("dimension_label", "") or "").strip()
            secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
            if mode_txt == "Perfil":
                normalized_dimension = dimension
                if secao and dimension.upper().startswith(f"{secao} "):
                    normalized_dimension = dimension[len(secao) :].strip()
                ref_txt = str(record.get("id", "") or "").strip() or f"{secao} {normalized_dimension}".strip()
                desc_txt = f"Perfil {secao} {normalized_dimension} {quality}".replace("  ", " ").strip()
                return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
            if mode_txt == "Tubo":
                ref_txt = str(record.get("id", "") or "").strip() or f"TUB-{dimension}".replace(" ", "")
                desc_txt = f"Tubo {quality} {dimension}".replace("  ", " ").strip()
                return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
            if mode_txt == "Chapa":
                esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                size_txt = str(dimension or "").replace(" mm", "").strip()
                ref_txt = str(record.get("id", "") or "").strip() or f"CH-{size_txt}x{esp_txt}".replace(" ", "")
                desc_txt = f"Chapa {quality} {size_txt}x{esp_txt} mm".replace("  ", " ").strip()
                return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
            if mode_txt == "Cantoneira":
                ref_txt = str(record.get("id", "") or "").strip() or f"L {dimension}".strip()
                desc_txt = f"Cantoneira abas iguais {quality} {dimension}".replace("  ", " ").strip()
                return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
            if mode_txt == "Barra":
                ref_txt = str(record.get("id", "") or "").strip() or f"BAR {dimension}".strip()
                desc_txt = f"Barra chata {quality} {dimension}".replace("  ", " ").strip()
                return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
            return {"ref": str(record.get("id", "") or "").strip(), "desc": quality, "quality": quality}

        stock_tube_dimensions: dict[str, list[str]] = {"redondo": [], "quadrado": [], "retangular": []}
        stock_sheet_dimensions: list[str] = []
        stock_bar_dimensions: list[str] = []
        for record in stock_rows:
            preview = dict(self.backend.material_price_preview(record) or {})
            mode_key = _stock_mode(record, preview)
            dimension = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
            if mode_key == "Tubo":
                tube_key = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().lower()
                if tube_key in stock_tube_dimensions and dimension:
                    stock_tube_dimensions[tube_key].append(dimension)
            elif mode_key == "Chapa":
                size_txt = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
                esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                if size_txt and esp_txt:
                    stock_sheet_dimensions.append(f"{size_txt}x{esp_txt}")
            elif mode_key == "Barra":
                if dimension:
                    stock_bar_dimensions.append(dimension)
            elif mode_key == "Cantoneira":
                a_val, b_val = _parse_pair(dimension)
                if a_val > 0:
                    leg_key = str(int(round(a_val)))
                    thickness_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                    if leg_key and thickness_txt:
                        equal_angle_presets.setdefault(leg_key, [])
                        if thickness_txt not in equal_angle_presets[leg_key]:
                            equal_angle_presets[leg_key].append(thickness_txt)

        for key, values in list(stock_tube_dimensions.items()):
            tube_model_presets[key] = _sort_dimension_labels(list(tube_model_presets.get(key, [])) + list(values))
        sheet_model_presets = _sort_dimension_labels(sheet_model_presets + stock_sheet_dimensions)
        flat_bar_presets = _sort_dimension_labels(flat_bar_presets + stock_bar_dimensions)
        equal_angle_presets = {
            leg: _sort_dimension_labels(values)
            for leg, values in sorted(equal_angle_presets.items(), key=lambda item: float(item[0]))
        }

        mode_to_index = {
            "Perfil": 0,
            "Tubo": 1,
            "Chapa": 2,
            "Cantoneira": 3,
            "Barra": 4,
            "Ferro nervurado": 5,
            "Manual": 6,
        }

        profile_source_combo = QComboBox()
        profile_source_combo.addItems(["Catálogo", "Stock MP"])
        profile_stock_combo = QComboBox()
        profile_series_combo = QComboBox()
        for row in profile_options:
            profile_series_combo.addItem(str(row.get("label", row.get("key", "")) or "").strip(), str(row.get("key", "") or "").strip())
        profile_size_combo = QComboBox()
        profile_size_combo.setEditable(True)
        profile_meters_spin = QDoubleSpinBox()
        profile_meters_spin.setRange(0.0, 1000.0)
        profile_meters_spin.setDecimals(3)
        profile_meters_spin.setSuffix(" m")
        profile_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
        profile_kg_per_m_spin = QDoubleSpinBox()
        profile_kg_per_m_spin.setRange(0.0, 1000.0)
        profile_kg_per_m_spin.setDecimals(4)
        profile_kg_per_m_spin.setSuffix(" kg/m")
        profile_price_kg_spin = QDoubleSpinBox()
        profile_price_kg_spin.setRange(0.0, 1000000.0)
        profile_price_kg_spin.setDecimals(4)
        profile_price_kg_spin.setPrefix("EUR ")
        profile_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
        profile_hint = QLabel("")
        profile_hint.setWordWrap(True)
        profile_hint.setProperty("role", "muted")

        page_profile = QWidget()
        profile_form = QFormLayout(page_profile)
        profile_form.setContentsMargins(0, 0, 0, 0)
        profile_form.addRow("Origem", profile_source_combo)
        profile_form.addRow("Stock MP", profile_stock_combo)
        profile_form.addRow("Série", profile_series_combo)
        profile_form.addRow("Tamanho", profile_size_combo)
        profile_form.addRow("Metros / unidade", profile_meters_spin)
        profile_form.addRow("Kg / metro", profile_kg_per_m_spin)
        profile_form.addRow("Preço / kg", profile_price_kg_spin)
        profile_form.addRow("", profile_hint)
        stack.addWidget(page_profile)

        tube_source_combo = QComboBox()
        tube_source_combo.addItems(["Novo modelo", "Stock MP"])
        tube_stock_combo = QComboBox()
        tube_quality_edit = QLineEdit(str(initial.get("quality", initial.get("material", "S235JR")) or "S235JR").strip())
        tube_section_combo = QComboBox()
        for row in tube_options:
            tube_section_combo.addItem(str(row.get("label", row.get("key", "")) or "").strip(), str(row.get("key", "") or "").strip())
        tube_model_combo = QComboBox()
        tube_model_combo.setEditable(True)
        tube_side_a_spin = QDoubleSpinBox()
        tube_side_a_spin.setRange(0.0, 5000.0)
        tube_side_a_spin.setDecimals(1)
        tube_side_a_spin.setSuffix(" mm")
        tube_side_b_spin = QDoubleSpinBox()
        tube_side_b_spin.setRange(0.0, 5000.0)
        tube_side_b_spin.setDecimals(1)
        tube_side_b_spin.setSuffix(" mm")
        tube_diameter_spin = QDoubleSpinBox()
        tube_diameter_spin.setRange(0.0, 5000.0)
        tube_diameter_spin.setDecimals(1)
        tube_diameter_spin.setSuffix(" mm")
        tube_thickness_spin = QDoubleSpinBox()
        tube_thickness_spin.setRange(0.0, 200.0)
        tube_thickness_spin.setDecimals(2)
        tube_thickness_spin.setSuffix(" mm")
        tube_meters_spin = QDoubleSpinBox()
        tube_meters_spin.setRange(0.0, 1000.0)
        tube_meters_spin.setDecimals(3)
        tube_meters_spin.setSuffix(" m")
        tube_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
        tube_kg_per_m_spin = QDoubleSpinBox()
        tube_kg_per_m_spin.setRange(0.0, 1000.0)
        tube_kg_per_m_spin.setDecimals(4)
        tube_kg_per_m_spin.setSuffix(" kg/m")
        tube_price_kg_spin = QDoubleSpinBox()
        tube_price_kg_spin.setRange(0.0, 1000000.0)
        tube_price_kg_spin.setDecimals(4)
        tube_price_kg_spin.setPrefix("EUR ")
        tube_price_kg_spin.setValue(float(initial.get("price_base_value", initial.get("price_per_kg", 3.15)) or 3.15))
        tube_hint = QLabel("")
        tube_hint.setWordWrap(True)
        tube_hint.setProperty("role", "muted")

        page_tube = QWidget()
        tube_form = QFormLayout(page_tube)
        tube_form.setContentsMargins(0, 0, 0, 0)
        tube_form.addRow("Origem", tube_source_combo)
        tube_form.addRow("Stock MP", tube_stock_combo)
        tube_form.addRow("Qualidade", tube_quality_edit)
        tube_form.addRow("Tipo", tube_section_combo)
        tube_form.addRow("Modelo / histórico", tube_model_combo)
        tube_form.addRow("Lado A / lado", tube_side_a_spin)
        tube_form.addRow("Lado B", tube_side_b_spin)
        tube_form.addRow("Diâmetro", tube_diameter_spin)
        tube_form.addRow("Espessura", tube_thickness_spin)
        tube_form.addRow("Metros / unidade", tube_meters_spin)
        tube_form.addRow("Kg / metro", tube_kg_per_m_spin)
        tube_form.addRow("Preço / m", tube_price_kg_spin)
        tube_form.addRow("", tube_hint)
        stack.addWidget(page_tube)

        sheet_source_combo = QComboBox()
        sheet_source_combo.addItems(["Novo modelo", "Stock MP"])
        sheet_stock_combo = QComboBox()
        sheet_family_combo = QComboBox()
        for option in family_options:
            sheet_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
        initial_quality_text = str(initial.get("quality", initial.get("material", "S235JR")) or "S235JR").strip()
        initial_family_key = str(initial.get("material_family_key", "") or initial.get("material_familia", "") or "").strip()
        if not initial_family_key:
            initial_family_key = str(self.backend.material_family_profile(initial_quality_text, "").get("key", "steel") or "steel").strip()
        _set_combo_by_data(sheet_family_combo, initial_family_key)
        sheet_quality_combo = QComboBox()
        sheet_quality_combo.setEditable(True)
        sheet_kind_combo = QComboBox()
        sheet_kind_combo.addItem("Chapa lisa", "plain")
        sheet_kind_combo.addItem("Chapa gota / xadrez", "checker")
        sheet_model_combo = QComboBox()
        sheet_model_combo.setEditable(True)
        length_mm_spin = QDoubleSpinBox()
        length_mm_spin.setRange(0.0, 100000.0)
        length_mm_spin.setDecimals(1)
        length_mm_spin.setSuffix(" mm")
        length_mm_spin.setValue(float(initial.get("length_mm", 3000) or 3000))
        width_mm_spin = QDoubleSpinBox()
        width_mm_spin.setRange(0.0, 100000.0)
        width_mm_spin.setDecimals(1)
        width_mm_spin.setSuffix(" mm")
        width_mm_spin.setValue(float(initial.get("width_mm", 1500) or 1500))
        thickness_mm_spin = QDoubleSpinBox()
        thickness_mm_spin.setRange(0.0, 1000.0)
        thickness_mm_spin.setDecimals(2)
        thickness_mm_spin.setSuffix(" mm")
        thickness_mm_spin.setValue(float(initial.get("thickness_mm", 15) or 15))
        density_spin = QDoubleSpinBox()
        density_spin.setRange(0.0, 50000.0)
        density_spin.setDecimals(1)
        density_spin.setSuffix(" kg/m3")
        density_spin.setValue(float(initial.get("density", 7850) or 7850))
        sheet_price_kg_spin = QDoubleSpinBox()
        sheet_price_kg_spin.setRange(0.0, 1000000.0)
        sheet_price_kg_spin.setDecimals(4)
        sheet_price_kg_spin.setPrefix("EUR ")
        sheet_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
        sheet_hint = QLabel("")
        sheet_hint.setWordWrap(True)
        sheet_hint.setProperty("role", "muted")
        initial_sheet_probe = self.backend.desktop_main.norm_text(
            " ".join(
                str(initial.get(key, "") or "")
                for key in ("sheet_kind", "sheet_model", "descricao_base", "descricao", "material")
            )
        )
        if any(token in initial_sheet_probe for token in ("gota", "xadrez", "antiderrap")):
            sheet_kind_combo.setCurrentIndex(1)

        page_sheet = QWidget()
        sheet_form = QFormLayout(page_sheet)
        sheet_form.setContentsMargins(0, 0, 0, 0)
        sheet_form.addRow("Origem", sheet_source_combo)
        sheet_form.addRow("Stock MP", sheet_stock_combo)
        sheet_form.addRow("Tipo material", sheet_family_combo)
        sheet_form.addRow("Qualidade", sheet_quality_combo)
        sheet_form.addRow("Tipo de chapa", sheet_kind_combo)
        sheet_form.addRow("Modelo / dimensão", sheet_model_combo)
        sheet_form.addRow("Comprimento", length_mm_spin)
        sheet_form.addRow("Largura", width_mm_spin)
        sheet_form.addRow("Espessura", thickness_mm_spin)
        sheet_form.addRow("Densidade", density_spin)
        sheet_form.addRow("Preço / kg", sheet_price_kg_spin)
        sheet_form.addRow("", sheet_hint)
        stack.addWidget(page_sheet)

        def _sheet_family_key() -> str:
            return str(sheet_family_combo.currentData() or "").strip() or "steel"

        def _sheet_family_density_kg_m3() -> float:
            profile = dict(self.backend.material_family_profile("", _sheet_family_key()) or {})
            density = float(profile.get("density", 7.85) or 7.85)
            return round(density * 1000.0, 1)

        def _refresh_sheet_quality_options(*, preserve_current: bool = True) -> None:
            current = sheet_quality_combo.currentText().strip() if preserve_current else ""
            if preserve_current and not current:
                current = initial_quality_text
            key = _sheet_family_key()
            values = list(family_quality_presets.get(key, []) or [])
            if current and current not in values and preserve_current:
                values.insert(0, current)
            if not current or current not in values:
                current = values[0] if values else current
            sheet_quality_combo.blockSignals(True)
            sheet_quality_combo.clear()
            for value in values:
                sheet_quality_combo.addItem(value)
            sheet_quality_combo.setCurrentText(current or (values[0] if values else ""))
            sheet_quality_combo.blockSignals(False)
            density_spin.blockSignals(True)
            density_spin.setValue(_sheet_family_density_kg_m3())
            density_spin.blockSignals(False)

        _refresh_sheet_quality_options()

        angle_source_combo = QComboBox()
        angle_source_combo.addItems(["Catálogo / histórico", "Stock MP"])
        angle_stock_combo = QComboBox()
        angle_family_combo = QComboBox()
        for option in family_options:
            angle_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
        _set_combo_by_data(angle_family_combo, initial_family_key)
        angle_quality_combo = QComboBox()
        angle_quality_combo.setEditable(True)
        angle_leg_combo = QComboBox()
        angle_leg_combo.addItems(list(equal_angle_presets.keys()))
        angle_leg_combo.setCurrentText(str(initial.get("angle_leg", "60") or "60"))
        angle_thickness_combo = QComboBox()
        angle_meters_spin = QDoubleSpinBox()
        angle_meters_spin.setRange(0.0, 1000.0)
        angle_meters_spin.setDecimals(3)
        angle_meters_spin.setSuffix(" m")
        angle_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
        angle_kg_per_m_spin = QDoubleSpinBox()
        angle_kg_per_m_spin.setRange(0.0, 1000.0)
        angle_kg_per_m_spin.setDecimals(4)
        angle_kg_per_m_spin.setSuffix(" kg/m")
        angle_price_kg_spin = QDoubleSpinBox()
        angle_price_kg_spin.setRange(0.0, 1000000.0)
        angle_price_kg_spin.setDecimals(4)
        angle_price_kg_spin.setPrefix("EUR ")
        angle_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
        angle_hint = QLabel("")
        angle_hint.setWordWrap(True)
        angle_hint.setProperty("role", "muted")

        page_angle = QWidget()
        angle_form = QFormLayout(page_angle)
        angle_form.setContentsMargins(0, 0, 0, 0)
        angle_form.addRow("Origem", angle_source_combo)
        angle_form.addRow("Stock MP", angle_stock_combo)
        angle_form.addRow("Tipo material", angle_family_combo)
        angle_form.addRow("Qualidade", angle_quality_combo)
        angle_form.addRow("Abas iguais", angle_leg_combo)
        angle_form.addRow("Espessura", angle_thickness_combo)
        angle_form.addRow("Metros / unidade", angle_meters_spin)
        angle_form.addRow("Kg / metro", angle_kg_per_m_spin)
        angle_form.addRow("Preço / kg", angle_price_kg_spin)
        angle_form.addRow("", angle_hint)
        stack.addWidget(page_angle)

        bar_source_combo = QComboBox()
        bar_source_combo.addItems(["Modelo / histórico", "Stock MP"])
        bar_stock_combo = QComboBox()
        bar_family_combo = QComboBox()
        for option in family_options:
            bar_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
        _set_combo_by_data(bar_family_combo, initial_family_key)
        bar_quality_combo = QComboBox()
        bar_quality_combo.setEditable(True)
        bar_size_combo = QComboBox()
        bar_size_combo.setEditable(True)
        for value in flat_bar_presets:
            bar_size_combo.addItem(value)
        bar_size_combo.setCurrentText(initial_bar_size or "70x20")
        bar_meters_spin = QDoubleSpinBox()
        bar_meters_spin.setRange(0.0, 1000.0)
        bar_meters_spin.setDecimals(3)
        bar_meters_spin.setSuffix(" m")
        bar_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
        bar_kg_per_m_spin = QDoubleSpinBox()
        bar_kg_per_m_spin.setRange(0.0, 1000.0)
        bar_kg_per_m_spin.setDecimals(4)
        bar_kg_per_m_spin.setSuffix(" kg/m")
        bar_price_kg_spin = QDoubleSpinBox()
        bar_price_kg_spin.setRange(0.0, 1000000.0)
        bar_price_kg_spin.setDecimals(4)
        bar_price_kg_spin.setPrefix("EUR ")
        bar_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
        bar_hint = QLabel("")
        bar_hint.setWordWrap(True)
        bar_hint.setProperty("role", "muted")

        page_bar = QWidget()
        bar_form = QFormLayout(page_bar)
        bar_form.setContentsMargins(0, 0, 0, 0)
        bar_form.addRow("Origem", bar_source_combo)
        bar_form.addRow("Stock MP", bar_stock_combo)
        bar_form.addRow("Tipo material", bar_family_combo)
        bar_form.addRow("Qualidade", bar_quality_combo)
        bar_form.addRow("Barra", bar_size_combo)
        bar_form.addRow("Metros / unidade", bar_meters_spin)
        bar_form.addRow("Kg / metro", bar_kg_per_m_spin)
        bar_form.addRow("Preço / kg", bar_price_kg_spin)
        bar_form.addRow("", bar_hint)
        stack.addWidget(page_bar)

        def _family_key(combo: QComboBox) -> str:
            return str(combo.currentData() or "").strip() or "steel"

        def _family_density_g_cm3(combo: QComboBox) -> float:
            profile = dict(self.backend.material_family_profile("", _family_key(combo)) or {})
            return float(profile.get("density", 7.85) or 7.85)

        def _refresh_quality_combo(family_combo: QComboBox, quality_combo: QComboBox, *, preserve_current: bool = True) -> None:
            current = quality_combo.currentText().strip() if preserve_current else ""
            if preserve_current and not current:
                current = initial_quality_text
            values = list(family_quality_presets.get(_family_key(family_combo), []) or [])
            if current and current not in values and preserve_current:
                values.insert(0, current)
            if not current or current not in values:
                current = values[0] if values else current
            quality_combo.blockSignals(True)
            quality_combo.clear()
            for value in values:
                quality_combo.addItem(value)
            quality_combo.setCurrentText(current or (values[0] if values else ""))
            quality_combo.blockSignals(False)

        _refresh_quality_combo(angle_family_combo, angle_quality_combo)
        _refresh_quality_combo(bar_family_combo, bar_quality_combo)

        rebar_meters_spin = QDoubleSpinBox()
        rebar_meters_spin.setRange(0.0, 1000.0)
        rebar_meters_spin.setDecimals(3)
        rebar_meters_spin.setSuffix(" m")
        rebar_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
        diameter_mm_spin = QDoubleSpinBox()
        diameter_mm_spin.setRange(0.0, 500.0)
        diameter_mm_spin.setDecimals(1)
        diameter_mm_spin.setSuffix(" mm")
        diameter_mm_spin.setValue(float(initial.get("diameter_mm", 12) or 12))
        rebar_price_kg_spin = QDoubleSpinBox()
        rebar_price_kg_spin.setRange(0.0, 1000000.0)
        rebar_price_kg_spin.setDecimals(4)
        rebar_price_kg_spin.setPrefix("EUR ")
        rebar_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))

        page_rebar = QWidget()
        rebar_form = QFormLayout(page_rebar)
        rebar_form.setContentsMargins(0, 0, 0, 0)
        rebar_form.addRow("Metros / unidade", rebar_meters_spin)
        rebar_form.addRow("Diâmetro", diameter_mm_spin)
        rebar_form.addRow("Preço / kg", rebar_price_kg_spin)
        stack.addWidget(page_rebar)

        unit_combo = QComboBox()
        unit_combo.addItems(["kg", "m", "un"])
        unit_combo.setCurrentText(str(initial.get("produto_unid", "un") or "un"))
        manual_unit_price_spin = QDoubleSpinBox()
        manual_unit_price_spin.setRange(0.0, 1000000.0)
        manual_unit_price_spin.setDecimals(4)
        manual_unit_price_spin.setPrefix("EUR ")
        manual_unit_price_spin.setValue(float(initial.get("manual_unit_price", initial.get("preco_unit", 0)) or 0))

        page_manual = QWidget()
        manual_form = QFormLayout(page_manual)
        manual_form.setContentsMargins(0, 0, 0, 0)
        manual_form.addRow("Unidade", unit_combo)
        manual_form.addRow("Preço unitário", manual_unit_price_spin)
        stack.addWidget(page_manual)

        def _refresh_profile_size_options() -> None:
            current_key = str(profile_series_combo.currentData() or "").strip()
            wanted = profile_size_combo.currentText().strip() or str(initial.get("profile_size", initial.get("perfil_tamanho", "")) or "").strip()
            options = [str(value or "").strip() for value in list(self.backend.material_profile_size_options(current_key) or []) if str(value or "").strip()]
            profile_size_combo.blockSignals(True)
            profile_size_combo.clear()
            for value in options:
                profile_size_combo.addItem(value)
            _set_combo_text_preserving_custom(profile_size_combo, wanted)
            profile_size_combo.blockSignals(False)

        def _refresh_angle_thickness_options() -> None:
            leg = str(angle_leg_combo.currentText() or "").strip()
            wanted = str(initial.get("angle_thickness", "6") or "6")
            angle_thickness_combo.blockSignals(True)
            angle_thickness_combo.clear()
            for value in equal_angle_presets.get(leg, ["6"]):
                angle_thickness_combo.addItem(value)
            angle_thickness_combo.setCurrentText(wanted)
            angle_thickness_combo.blockSignals(False)

        def _refresh_tube_model_options() -> None:
            secao = str(tube_section_combo.currentData() or "").strip().lower() or "quadrado"
            wanted = str(initial.get("tube_model", "") or "").strip()
            options = list(tube_model_presets.get(secao, []))
            if secao == "redondo" and not wanted and float(tube_diameter_spin.value() or 0) > 0 and float(tube_thickness_spin.value() or 0) > 0:
                wanted = f"Ø{tube_diameter_spin.value():.1f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
            elif not wanted and float(tube_side_a_spin.value() or 0) > 0 and float(tube_thickness_spin.value() or 0) > 0:
                if secao == "retangular" and float(tube_side_b_spin.value() or 0) > 0:
                    wanted = f"{tube_side_a_spin.value():.0f}x{tube_side_b_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
                else:
                    wanted = f"{tube_side_a_spin.value():.0f}x{tube_side_a_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
            tube_model_combo.blockSignals(True)
            tube_model_combo.clear()
            for value in options:
                tube_model_combo.addItem(value)
            _set_combo_text_preserving_custom(tube_model_combo, wanted)
            tube_model_combo.blockSignals(False)

        def _apply_tube_model_selection() -> None:
            if tube_source_combo.currentText().strip() == "Stock MP":
                return
            model_txt = tube_model_combo.currentText().strip().lower().replace(",", ".").replace("mm", "")
            if not model_txt:
                return
            numbers = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", model_txt)]
            secao = str(tube_section_combo.currentData() or "").strip().lower()
            if secao == "redondo":
                if len(numbers) >= 2:
                    tube_diameter_spin.setValue(numbers[0])
                    tube_thickness_spin.setValue(numbers[1])
            elif secao == "retangular":
                if len(numbers) >= 3:
                    tube_side_a_spin.setValue(numbers[0])
                    tube_side_b_spin.setValue(numbers[1])
                    tube_thickness_spin.setValue(numbers[2])
            else:
                if len(numbers) >= 3:
                    tube_side_a_spin.setValue(numbers[0])
                    tube_side_b_spin.setValue(numbers[0])
                    tube_thickness_spin.setValue(numbers[2])
                elif len(numbers) >= 2:
                    tube_side_a_spin.setValue(numbers[0])
                    tube_side_b_spin.setValue(numbers[0])
                    tube_thickness_spin.setValue(numbers[1])

        def _refresh_sheet_model_options() -> None:
            wanted = str(initial.get("sheet_model", "") or "").strip()
            if not wanted and float(length_mm_spin.value() or 0) > 0 and float(width_mm_spin.value() or 0) > 0 and float(thickness_mm_spin.value() or 0) > 0:
                wanted = f"{length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{thickness_mm_spin.value():.1f}".replace(".0", "")
            checker_mode = str(sheet_kind_combo.currentData() or "").strip() == "checker"
            if checker_mode and (not wanted or "/" not in wanted):
                wanted = "Chapa gota 3000x1500x6/8"
            options = []
            for value in sheet_model_presets:
                normalized_value = self.backend.desktop_main.norm_text(value)
                value_is_checker = "/" in str(value or "") or any(token in normalized_value for token in ("gota", "xadrez", "antiderrap"))
                if checker_mode == value_is_checker:
                    options.append(value)
            sheet_model_combo.blockSignals(True)
            sheet_model_combo.clear()
            for value in options:
                sheet_model_combo.addItem(value)
            _set_combo_text_preserving_custom(sheet_model_combo, wanted)
            sheet_model_combo.blockSignals(False)

        def _apply_sheet_model_selection() -> None:
            if sheet_source_combo.currentText().strip() == "Stock MP":
                return
            model = _sheet_model_info(sheet_model_combo.currentText())
            if not str(model.get("raw", "") or "").strip():
                return
            if float(model.get("length", 0) or 0) > 0:
                length_mm_spin.setValue(float(model.get("length", 0) or 0))
            if float(model.get("width", 0) or 0) > 0:
                width_mm_spin.setValue(float(model.get("width", 0) or 0))
            if float(model.get("thickness", 0) or 0) > 0:
                thickness_mm_spin.setValue(float(model.get("thickness", 0) or 0))

        _fill_stock_combo(profile_stock_combo, "Perfil", str(initial.get("stock_material_id", "") or "").strip())
        _fill_stock_combo(tube_stock_combo, "Tubo", str(initial.get("stock_material_id", "") or "").strip())
        _fill_stock_combo(sheet_stock_combo, "Chapa", str(initial.get("stock_material_id", "") or "").strip())
        _fill_stock_combo(angle_stock_combo, "Cantoneira", str(initial.get("stock_material_id", "") or "").strip())
        _fill_stock_combo(bar_stock_combo, "Barra", str(initial.get("stock_material_id", "") or "").strip())
        _refresh_profile_size_options()
        _refresh_angle_thickness_options()
        _refresh_tube_model_options()
        _refresh_sheet_model_options()
        if str(initial.get("stock_material_id", "") or "").strip():
            current_mode = str(initial.get("calc_mode", "Perfil") or "Perfil").strip()
            if current_mode == "Perfil":
                profile_source_combo.setCurrentText("Stock MP")
            elif current_mode == "Tubo":
                tube_source_combo.setCurrentText("Stock MP")
            elif current_mode == "Chapa":
                sheet_source_combo.setCurrentText("Stock MP")
            elif current_mode == "Cantoneira":
                angle_source_combo.setCurrentText("Stock MP")
            elif current_mode == "Barra":
                bar_source_combo.setCurrentText("Stock MP")

        stock_sync_state = {
            "profile": "",
            "tube": "",
            "sheet": "",
            "angle": "",
            "bar": "",
        }

        default_price_map = {
            "Perfil": float(self.backend.material_default_price_kg("Perfil") or 3.15),
            "Tubo": float(self.backend.material_default_price_kg("Tubo") or 3.15),
            "Chapa": float(self.backend.material_default_price_kg("Chapa") or 3.15),
            "Cantoneira": float(self.backend.material_default_price_kg("Cantoneira") or self.backend.material_default_price_kg("Perfil") or 3.15),
            "Barra": float(self.backend.material_default_price_kg("Barra") or self.backend.material_default_price_kg("Chapa") or 3.15),
            "Ferro nervurado": float(self.backend.material_default_price_kg("Barra", "ferro") or self.backend.material_default_price_kg("Barra") or 3.15),
        }
        if float(initial.get("price_per_kg", 0) or 0) <= 0:
            profile_price_kg_spin.setValue(default_price_map["Perfil"])
            tube_price_kg_spin.setValue(default_price_map["Tubo"])
            sheet_price_kg_spin.setValue(default_price_map["Chapa"])
            angle_price_kg_spin.setValue(default_price_map["Cantoneira"])
            bar_price_kg_spin.setValue(default_price_map["Barra"])
            rebar_price_kg_spin.setValue(default_price_map["Ferro nervurado"])

        def _profile_payload() -> dict:
            if profile_source_combo.currentText().strip() == "Stock MP":
                record, preview = _current_stock(profile_stock_combo)
                if record:
                    identity = _stock_material_identity("Perfil", record, preview)
                    series, size_text = _profile_series_and_size(record, preview)
                    meters = float(profile_meters_spin.value() or record.get("metros", 0) or 0.0)
                    kg_m = float(profile_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                    return {
                        "stock_material_id": str(record.get("id", "") or "").strip(),
                        "ref": identity["ref"],
                        "desc": identity["desc"],
                        "meters": meters,
                        "kg_m": kg_m,
                        "price_kg": float(profile_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                        "weight_each": round(meters * kg_m, 4),
                        "hint": str(preview.get("calc_hint", "") or "").strip(),
                        "quality": identity["quality"],
                        "series": series,
                        "size": size_text,
                    }
            series = str(profile_series_combo.currentData() or "").strip()
            size_text, size_mm = _parse_size_value(profile_size_combo.currentText())
            preview = dict(
                self.backend.material_geometry_preview(
                    {
                        "formato": "Perfil",
                        "material": f"{series} {size_text}".strip(),
                        "secao_tipo": series,
                        "altura": size_mm,
                        "perfil_tamanho": size_text,
                        "metros": float(profile_meters_spin.value() or 0.0),
                    }
                )
                or {}
            )
            desc = f"Perfil {series} {size_text}".strip()
            return {
                "stock_material_id": "",
                "ref": f"{series} {size_text}".strip(),
                "desc": desc,
                "meters": float(profile_meters_spin.value() or 0.0),
                "kg_m": float(preview.get("kg_m", 0) or 0.0),
                "price_kg": float(profile_price_kg_spin.value() or 0.0),
                "weight_each": float(preview.get("peso_unid", 0) or 0.0),
                "hint": str(preview.get("calc_hint", "") or "").strip(),
                "series": series,
                "size": size_text,
            }

        def _tube_payload() -> dict:
            if tube_source_combo.currentText().strip() == "Stock MP":
                record, preview = _current_stock(tube_stock_combo)
                if record:
                    identity = _stock_material_identity("Tubo", record, preview)
                    meters = float(tube_meters_spin.value() or record.get("metros", 0) or 0.0)
                    kg_m = float(tube_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                    base_label, base_value = _stock_price_base(record, preview)
                    if base_label.strip().upper() != "EUR/M":
                        base_label = "EUR/m"
                    price_m = float(tube_price_kg_spin.value() or base_value or 0.0)
                    price_kg_equiv = round(price_m / kg_m, 4) if kg_m > 0 else 0.0
                    return {
                        "stock_material_id": str(record.get("id", "") or "").strip(),
                        "ref": identity["ref"],
                        "desc": identity["desc"],
                        "meters": meters,
                        "kg_m": kg_m,
                        "price_kg": price_kg_equiv,
                        "price_base_label": "EUR/m",
                        "price_base_value": price_m,
                        "weight_each": round(meters * kg_m, 4),
                        "hint": str(preview.get("calc_hint", "") or "").strip(),
                        "quality": identity["quality"],
                    }
            secao = str(tube_section_combo.currentData() or "").strip()
            quality = tube_quality_edit.text().strip() or "S235JR"
            payload = {
                "formato": "Tubo",
                "material": quality,
                "secao_tipo": secao,
                "espessura": _float_text(tube_thickness_spin.value(), 2).replace(",", "."),
                "metros": float(tube_meters_spin.value() or 0.0),
                "comprimento": float(tube_side_a_spin.value() or 0.0),
                "largura": float(tube_side_b_spin.value() or 0.0),
                "diametro": float(tube_diameter_spin.value() or 0.0),
            }
            preview = dict(self.backend.material_geometry_preview(payload) or {})
            meters = float(tube_meters_spin.value() or 0.0)
            kg_m = float(preview.get("kg_m", 0) or 0.0)
            price_m = float(tube_price_kg_spin.value() or 0.0)
            price_kg_equiv = round(price_m / kg_m, 4) if kg_m > 0 else 0.0
            if secao == "redondo":
                desc = f"Tubo {quality} Ø{tube_diameter_spin.value():.1f}x{tube_thickness_spin.value():.1f} mm".replace(".0", "")
                ref = f"TUB-RED-{tube_diameter_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
            else:
                a_val = tube_side_a_spin.value()
                b_val = tube_side_b_spin.value() if secao == "retangular" else tube_side_a_spin.value()
                dim_txt = f"{a_val:.0f}x{b_val:.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
                desc = f"Tubo {quality} {dim_txt} mm"
                ref = f"TUB-{dim_txt}"
            return {
                "stock_material_id": "",
                "ref": ref,
                "desc": desc,
                "meters": meters,
                "kg_m": kg_m,
                "price_kg": price_kg_equiv,
                "price_base_label": "EUR/m",
                "price_base_value": price_m,
                "weight_each": float(preview.get("peso_unid", 0) or 0.0),
                "hint": str(preview.get("calc_hint", "") or "").strip(),
                "quality": quality,
            }

        def _sheet_payload() -> dict:
            if sheet_source_combo.currentText().strip() == "Stock MP":
                record, preview = _current_stock(sheet_stock_combo)
                if record:
                    identity = _stock_material_identity("Chapa", record, preview)
                    return {
                        "stock_material_id": str(record.get("id", "") or "").strip(),
                        "ref": identity["ref"],
                        "desc": identity["desc"],
                        "meters": 0.0,
                        "kg_m": 0.0,
                        "price_kg": float(sheet_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                        "weight_each": float(preview.get("peso_unid", 0) or 0.0),
                        "hint": str(preview.get("calc_hint", "") or "").strip(),
                        "quality": identity["quality"],
                    }
            quality = sheet_quality_combo.currentText().strip() or "S235JR"
            family_key = _sheet_family_key()
            model = _sheet_model_info(sheet_model_combo.currentText())
            is_checker = bool(model.get("is_checker"))
            thickness_label = str(model.get("thickness_label", "") or "").strip()
            display_thickness = thickness_label or _float_text(thickness_mm_spin.value(), 2).replace(",00", "").replace(",0", "")
            sheet_kind = "Chapa gota" if is_checker else "Chapa"
            density_g_cm3 = round(float(density_spin.value() or 7850.0) / 1000.0, 4)
            preview = dict(
                self.backend.material_geometry_preview(
                    {
                        "formato": "Chapa",
                        "material": quality,
                        "comprimento": float(length_mm_spin.value() or 0.0),
                        "largura": float(width_mm_spin.value() or 0.0),
                        "espessura": display_thickness.replace(",", "."),
                        "material_familia": family_key,
                        "densidade": density_g_cm3,
                        "descricao": f"{sheet_kind} {quality}",
                    }
                )
                or {}
            )
            desc = (
                f"{sheet_kind} {quality} {length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{display_thickness} mm"
            ).replace(".0", "")
            ref_prefix = "CH-GOTA" if is_checker else "CH"
            ref = f"{ref_prefix}-{length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{display_thickness}".replace(".0", "")
            return {
                "stock_material_id": "",
                "ref": ref,
                "desc": desc,
                "meters": 0.0,
                "kg_m": 0.0,
                "price_kg": float(sheet_price_kg_spin.value() or 0.0),
                "weight_each": float(preview.get("peso_unid", 0) or 0.0),
                "hint": str(preview.get("calc_hint", "") or "").strip(),
                "quality": quality,
                "material_family_key": family_key,
                "esp_label": display_thickness,
                "sheet_kind": sheet_kind,
            }

        def _angle_payload() -> dict:
            if angle_source_combo.currentText().strip() == "Stock MP":
                record, preview = _current_stock(angle_stock_combo)
                if record:
                    identity = _stock_material_identity("Cantoneira", record, preview)
                    meters = float(angle_meters_spin.value() or record.get("metros", 0) or 0.0)
                    kg_m = float(angle_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                    return {
                        "stock_material_id": str(record.get("id", "") or "").strip(),
                        "ref": identity["ref"],
                        "desc": identity["desc"],
                        "meters": meters,
                        "kg_m": kg_m,
                        "price_kg": float(angle_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                        "weight_each": round(meters * kg_m, 4),
                        "hint": str(preview.get("calc_hint", "") or "").strip(),
                        "quality": identity["quality"],
                        "material_family_key": str(record.get("material_familia", preview.get("material_familia_resolved", "")) or "").strip(),
                    }
            leg = float(angle_leg_combo.currentText() or 0.0)
            thickness = float(angle_thickness_combo.currentText() or 0.0)
            density = _family_density_g_cm3(angle_family_combo)
            kg_m = _geometry_weight_kg_m((thickness * ((2.0 * leg) - thickness)), density)
            meters = float(angle_meters_spin.value() or 0.0)
            quality = angle_quality_combo.currentText().strip() or "S235JR"
            desc = f"Cantoneira abas iguais {quality} {leg:.0f}x{leg:.0f}x{thickness:.0f} mm"
            ref = f"L {leg:.0f}x{leg:.0f}x{thickness:.0f}"
            return {
                "stock_material_id": "",
                "ref": ref,
                "desc": desc,
                "meters": meters,
                "kg_m": kg_m,
                "price_kg": float(angle_price_kg_spin.value() or 0.0),
                "weight_each": round(kg_m * meters, 4),
                "hint": f"Cantoneira: área aproximada t x (2a - t) x densidade {density:.3f} g/cm3.",
                "quality": quality,
                "material_family_key": _family_key(angle_family_combo),
            }

        def _bar_payload() -> dict:
            if bar_source_combo.currentText().strip() == "Stock MP":
                record, preview = _current_stock(bar_stock_combo)
                if record:
                    identity = _stock_material_identity("Barra", record, preview)
                    meters = float(bar_meters_spin.value() or record.get("metros", 0) or 0.0)
                    kg_m = float(bar_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                    return {
                        "stock_material_id": str(record.get("id", "") or "").strip(),
                        "ref": identity["ref"],
                        "desc": identity["desc"],
                        "meters": meters,
                        "kg_m": kg_m,
                        "price_kg": float(bar_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                        "weight_each": round(meters * kg_m, 4),
                        "hint": str(preview.get("calc_hint", "") or "").strip(),
                        "quality": identity["quality"],
                        "material_family_key": str(record.get("material_familia", preview.get("material_familia_resolved", "")) or "").strip(),
                    }
            width_mm, thickness_mm = _parse_pair(bar_size_combo.currentText())
            density = _family_density_g_cm3(bar_family_combo)
            kg_m = _geometry_weight_kg_m(width_mm * thickness_mm, density)
            meters = float(bar_meters_spin.value() or 0.0)
            quality = bar_quality_combo.currentText().strip() or "S235JR"
            desc = f"Barra chata {quality} {width_mm:.0f}x{thickness_mm:.0f} mm"
            ref = f"BAR {width_mm:.0f}x{thickness_mm:.0f}"
            return {
                "stock_material_id": "",
                "ref": ref,
                "desc": desc,
                "meters": meters,
                "kg_m": kg_m,
                "price_kg": float(bar_price_kg_spin.value() or 0.0),
                "weight_each": round(kg_m * meters, 4),
                "hint": f"Barra chata: largura x espessura x densidade {density:.3f} g/cm3.",
                "quality": quality,
                "material_family_key": _family_key(bar_family_combo),
            }

        def _rebar_payload() -> dict:
            meters = float(rebar_meters_spin.value() or 0.0)
            diameter = float(diameter_mm_spin.value() or 0.0)
            kg_m = round(0.006165 * (diameter ** 2), 4)
            desc = f"Ferro nervurado Ø{diameter:.0f} mm".replace(".0", "")
            ref = f"REBAR {diameter:.0f}".replace(".0", "")
            return {
                "stock_material_id": "",
                "ref": ref,
                "desc": desc,
                "meters": meters,
                "kg_m": kg_m,
                "price_kg": float(rebar_price_kg_spin.value() or 0.0),
                "weight_each": round(kg_m * meters, 4),
                "hint": "Ferro nervurado: 0.006165 x Ø².",
            }

        def _manual_payload() -> dict:
            unit_txt = unit_combo.currentText().strip() or "un"
            return {
                "stock_material_id": "",
                "ref": ref_edit.text().strip(),
                "desc": desc_edit.text().strip() or "Material manual",
                "meters": 0.0,
                "kg_m": 0.0,
                "price_kg": 0.0,
                "weight_each": 0.0,
                "hint": "",
                "unit": unit_txt,
                "unit_price": float(manual_unit_price_spin.value() or 0.0),
            }

        def _mode_payload() -> dict:
            mode = mode_combo.currentText().strip()
            if mode == "Perfil":
                return _profile_payload()
            if mode == "Tubo":
                return _tube_payload()
            if mode == "Chapa":
                return _sheet_payload()
            if mode == "Cantoneira":
                return _angle_payload()
            if mode == "Barra":
                return _bar_payload()
            if mode == "Ferro nervurado":
                return _rebar_payload()
            return _manual_payload()

        def _selected_operation_names() -> list[str]:
            selected: list[str] = []
            for token in _operation_tokens(operation_edit.text().strip()):
                normalized = str(self.backend.desktop_main.normalize_operacao_nome(token) or token or "").strip()
                if normalized and normalized not in selected:
                    selected.append(normalized)
            return selected

        def _sync_operation_meta_selection() -> None:
            selected_ops = _selected_operation_names()
            selected_keys = {
                str(self.backend.desktop_main.normalize_operacao_nome(op) or op or "").strip()
                for op in selected_ops
                if str(op or "").strip()
            }
            operation_meta["operacoes_lista"] = list(selected_ops)
            operation_meta["operacoes_fluxo"] = self.backend.desktop_main.build_operacoes_fluxo(
                selected_ops,
                operation_meta.get("operacoes_fluxo") if isinstance(operation_meta.get("operacoes_fluxo"), list) else None,
            )
            operation_meta["operacoes_detalhe"] = [
                dict(item or {})
                for item in list(operation_meta.get("operacoes_detalhe", []) or [])
                if str(self.backend.desktop_main.normalize_operacao_nome((item or {}).get("nome", "")) or (item or {}).get("nome", "") or "").strip() in selected_keys
            ]
            operation_meta["tempos_operacao"] = {
                str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip(): float(value or 0)
                for op_name, value in dict(operation_meta.get("tempos_operacao", {}) or {}).items()
                if str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip() in selected_keys
            }
            operation_meta["custos_operacao"] = {
                str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip(): float(value or 0)
                for op_name, value in dict(operation_meta.get("custos_operacao", {}) or {}).items()
                if str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip() in selected_keys
            }

        def _operation_totals() -> tuple[float, float]:
            selected_keys = {
                str(self.backend.desktop_main.normalize_operacao_nome(op) or op or "").strip()
                for op in _selected_operation_names()
                if str(op or "").strip()
            }
            time_total = 0.0
            cost_total = 0.0
            for op_name, value in dict(operation_meta.get("tempos_operacao", {}) or {}).items():
                normalized = str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized in selected_keys:
                    time_total += float(value or 0)
            for op_name, value in dict(operation_meta.get("custos_operacao", {}) or {}).items():
                normalized = str(self.backend.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized in selected_keys:
                    cost_total += float(value or 0)
            return round(time_total, 4), round(cost_total, 4)

        def _operation_cost_payload(base_unit_cost: float = 0.0, qty_value: float | None = None) -> dict:
            return {
                "operacao": operation_edit.text().strip(),
                "operacoes_lista": list(_selected_operation_names()),
                "costing_operations": list(_selected_operation_names()),
                "qtd": float(qty_value if qty_value is not None else qty_spin.value() or 0),
                "tempo_peca_min": float(_operation_totals()[0] or 0),
                "preco_unit": float(base_unit_cost or 0),
                "blend_with_current_line": True,
                "base_tempo_unit_min": 0.0,
                "base_preco_unit_eur": float(base_unit_cost or 0),
                "base_operation_label": "Material base",
                "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
            }

        def _refresh_operation_cost_hint(base_unit_cost: float = 0.0) -> None:
            _sync_operation_meta_selection()
            ops = _selected_operation_names()
            if not ops:
                operation_cost_label.setText("Sem operações seguintes. O item fica apenas como matéria-prima/stock.")
                return
            estimate = dict(self.backend.operation_cost_estimate(_operation_cost_payload(base_unit_cost)) or {})
            summary = dict(estimate.get("summary", {}) or {})
            pending = [
                dict(item)
                for item in list(estimate.get("operations", []) or [])
                if isinstance(item, dict) and bool(item.get("missing_driver_input"))
            ]
            if pending:
                missing_txt = ", ".join(str(item.get("nome", "") or "").strip() for item in pending if str(item.get("nome", "") or "").strip())
                operation_cost_label.setText(f"Falta quantificar: {missing_txt}.")
                return
            extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
            extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
            operation_cost_label.setText(
                f"Operações: {extra_time:.3f} min/un | {_fmt_eur(extra_cost)}/un | "
                f"total com material {_fmt_eur(float(base_unit_cost or 0) + extra_cost)}/un"
            )

        def _edit_operation_costs() -> bool:
            _sync_operation_meta_selection()
            if not _selected_operation_names():
                QMessageBox.information(dialog, "Operações", "Seleciona primeiro as operações seguintes deste material.")
                return False
            current = _compute()
            base_cost = float((current.get("line") or {}).get("_material_base_preco_unit", (current.get("line") or {}).get("preco_unit", 0)) or 0)
            result = _open_quote_operation_detail_dialog(dialog, self.backend, _operation_cost_payload(base_cost))
            if not isinstance(result, dict):
                return False
            operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(result.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
            operation_meta["tempos_operacao"] = dict(result.get("tempos_operacao", {}) or {})
            operation_meta["custos_operacao"] = dict(result.get("custos_operacao", {}) or {})
            operation_meta["quote_cost_snapshot"] = dict(result.get("quote_cost_snapshot", {}) or {})
            _refresh()
            return True

        def _apply_auto_texts(payload: dict) -> None:
            if sync["busy"]:
                return
            auto_desc = str(payload.get("desc", "") or "").strip()
            auto_ref = str(payload.get("ref", "") or "").strip()
            if auto_desc:
                _set_edit_if_auto(desc_edit, desc_state, auto_desc)
            if auto_ref:
                _set_edit_if_auto(ref_edit, ref_state, auto_ref)

        def _toggle_controls() -> None:
            mode = mode_combo.currentText().strip()
            stack.setCurrentIndex(mode_to_index.get(mode, 0))
            profile_is_stock = profile_source_combo.currentText().strip() == "Stock MP"
            tube_is_stock = tube_source_combo.currentText().strip() == "Stock MP"
            sheet_is_stock = sheet_source_combo.currentText().strip() == "Stock MP"
            angle_is_stock = angle_source_combo.currentText().strip() == "Stock MP"
            bar_is_stock = bar_source_combo.currentText().strip() == "Stock MP"

            profile_stock_combo.setVisible(profile_is_stock)
            profile_form.labelForField(profile_stock_combo).setVisible(profile_is_stock)
            for widget in (profile_series_combo, profile_size_combo):
                widget.setEnabled(not profile_is_stock)

            tube_stock_combo.setVisible(tube_is_stock)
            tube_form.labelForField(tube_stock_combo).setVisible(tube_is_stock)
            for widget in (tube_quality_edit, tube_section_combo, tube_model_combo, tube_side_a_spin, tube_side_b_spin, tube_diameter_spin, tube_thickness_spin):
                widget.setEnabled(not tube_is_stock)
            tube_model_combo.setVisible(not tube_is_stock)
            tube_form.labelForField(tube_model_combo).setVisible(not tube_is_stock)
            tube_is_round = str(tube_section_combo.currentData() or "").strip() == "redondo"
            tube_side_a_spin.setVisible(not tube_is_round)
            tube_form.labelForField(tube_side_a_spin).setVisible(not tube_is_round)
            tube_side_b_spin.setVisible(not tube_is_round and str(tube_section_combo.currentData() or "").strip() == "retangular")
            tube_form.labelForField(tube_side_b_spin).setVisible(not tube_is_round and str(tube_section_combo.currentData() or "").strip() == "retangular")
            tube_diameter_spin.setVisible(tube_is_round)
            tube_form.labelForField(tube_diameter_spin).setVisible(tube_is_round)

            sheet_stock_combo.setVisible(sheet_is_stock)
            sheet_form.labelForField(sheet_stock_combo).setVisible(sheet_is_stock)
            for widget in (sheet_family_combo, sheet_quality_combo, sheet_kind_combo, sheet_model_combo, length_mm_spin, width_mm_spin, thickness_mm_spin, density_spin):
                widget.setEnabled(not sheet_is_stock)
            sheet_model_combo.setVisible(not sheet_is_stock)
            sheet_form.labelForField(sheet_model_combo).setVisible(not sheet_is_stock)

            angle_stock_combo.setVisible(angle_is_stock)
            angle_form.labelForField(angle_stock_combo).setVisible(angle_is_stock)
            for widget in (angle_family_combo, angle_quality_combo, angle_leg_combo, angle_thickness_combo):
                widget.setEnabled(not angle_is_stock)

            bar_stock_combo.setVisible(bar_is_stock)
            bar_form.labelForField(bar_stock_combo).setVisible(bar_is_stock)
            for widget in (bar_family_combo, bar_quality_combo, bar_size_combo):
                widget.setEnabled(not bar_is_stock)

        def _compute() -> dict:
            mode = mode_combo.currentText().strip()
            qty_units = float(qty_spin.value() or 0.0)
            payload = dict(_mode_payload() or {})
            description = desc_edit.text().strip() or str(payload.get("desc", "") or mode).strip() or mode
            reference = ref_edit.text().strip() or str(payload.get("ref", "") or "").strip()
            total_weight = 0.0
            unit_cost = 0.0
            total_cost = 0.0
            qty_line = qty_units
            unit_label = "un"
            hint = str(payload.get("hint", "") or "").strip()
            weight_each = float(payload.get("weight_each", 0) or 0.0)
            meters_per_unit = float(payload.get("meters", 0) or 0.0)
            kg_per_m = float(payload.get("kg_m", 0) or 0.0)
            price_base_label = str(payload.get("price_base_label", "EUR/kg" if mode != "Tubo" else "EUR/m") or "").strip() or ("EUR/m" if mode == "Tubo" else "EUR/kg")
            price_base_value = float(payload.get("price_base_value", payload.get("price_kg", 0)) or 0.0)
            if mode == "Manual":
                unit_label = str(payload.get("unit", unit_combo.currentText()) or "un").strip() or "un"
                unit_cost = round(float(payload.get("unit_price", manual_unit_price_spin.value()) or 0.0), 4)
                total_cost = round(qty_units * unit_cost, 2)
                detail = f"{description} | {qty_units:.2f} {unit_label}"
            else:
                if weight_each <= 0 and meters_per_unit > 0 and kg_per_m > 0:
                    weight_each = round(meters_per_unit * kg_per_m, 4)
                total_weight = round(weight_each * qty_units, 3)
                if mode == "Tubo" and price_base_label.upper() == "EUR/M":
                    unit_cost = round(meters_per_unit * price_base_value, 4) if meters_per_unit > 0 else 0.0
                else:
                    unit_cost = round(weight_each * float(payload.get("price_kg", price_base_value) or 0.0), 4) if weight_each > 0 else 0.0
                total_cost = round(unit_cost * qty_units, 2)
                if meters_per_unit > 0:
                    base_txt = f" | {price_base_value:.4f} {price_base_label}" if price_base_value > 0 else ""
                    detail = f"{description} | {qty_units:.2f} un x {meters_per_unit:.2f} m | {kg_per_m:.3f} kg/m | {total_weight:.2f} kg{base_txt}"
                elif total_weight > 0:
                    detail = f"{description} | {qty_units:.2f} un | {total_weight:.2f} kg"
                else:
                    detail = f"{description} | {qty_units:.2f} un"
            material_base_cost = round(unit_cost, 4)
            operation_time_unit, operation_cost_unit = _operation_totals()
            unit_cost = round(unit_cost + operation_cost_unit, 4)
            total_cost = round(unit_cost * qty_units, 2)
            thickness_value = 0.0
            if mode == "Chapa":
                thickness_value = float(thickness_mm_spin.value() or 0.0)
            elif mode == "Tubo":
                thickness_value = float(tube_thickness_spin.value() or 0.0)
            elif mode == "Cantoneira":
                thickness_value = float(angle_thickness_combo.currentText() or 0.0)
            elif mode == "Barra":
                _bar_w, thickness_value = _parse_pair(bar_size_combo.currentText())
            elif mode == "Ferro nervurado":
                thickness_value = float(diameter_mm_spin.value() or 0.0)

            quality_value = str(payload.get("quality", "") or "").strip()
            material_label = quality_value
            if mode != "Manual":
                material_label = f"{mode} {quality_value}".strip() if quality_value else mode
            if mode == "Chapa" and str(payload.get("sheet_kind", "") or "").strip():
                sheet_kind = str(payload.get("sheet_kind", "") or "").strip()
                material_label = f"{sheet_kind} {quality_value}".strip() if quality_value else sheet_kind
            esp_txt = ""
            if mode == "Chapa" and str(payload.get("esp_label", "") or "").strip():
                esp_txt = str(payload.get("esp_label", "") or "").strip()
            elif thickness_value > 0:
                esp_txt = _float_text(thickness_value, 2).replace(",00", "").replace(",0", "")
            elif mode == "Perfil":
                esp_txt = str(payload.get("size", "") or profile_size_combo.currentText() or "").strip()
            elif mode == "Manual":
                esp_txt = str(unit_label).strip()
            line = {
                "tipo_item": self.backend.desktop_main.ORC_LINE_TYPE_PIECE,
                "stock_item_kind": "raw_material",
                "descricao": detail,
                "ref_externa": reference,
                "material": material_label or mode,
                "material_family": quality_value or material_label or mode,
                "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
                "material_subtype": mode,
                "espessura": esp_txt,
                "operacao": " + ".join(_selected_operation_names()),
                "qtd": round(qty_line, 2),
                "preco_unit": round(unit_cost, 4),
                "desenho": "",
                "tempo_peca_min": round(operation_time_unit, 4),
                "calc_mode": mode,
                "descricao_base": description,
                "weight_total": round(total_weight, 3),
                "total_cost": round(total_cost, 2),
                "quantity_units": round(qty_units, 2),
                "price_per_kg": round(float(payload.get("price_kg", 0) or 0.0), 4),
                "price_base_label": price_base_label,
                "price_base_value": round(price_base_value, 4),
                "stock_metric_value": round(meters_per_unit if price_base_label.upper() == "EUR/M" else weight_each, 4),
                "meters_per_unit": round(meters_per_unit, 3),
                "kg_per_m": round(kg_per_m, 4),
                "length_mm": round(float(length_mm_spin.value() or 0.0), 1),
                "width_mm": round(float(width_mm_spin.value() or 0.0), 1),
                "thickness_mm": round(float(thickness_mm_spin.value() or tube_thickness_spin.value() or 0.0), 2),
                "density": round(float(density_spin.value() or 0.0), 1),
                "diameter_mm": round(float(tube_diameter_spin.value() or diameter_mm_spin.value() or 0.0), 1),
                "manual_unit_price": round(float(manual_unit_price_spin.value() or 0.0), 4),
                "profile_section": str(payload.get("series", "") or profile_series_combo.currentData() or "").strip(),
                "profile_size": str(payload.get("size", "") or profile_size_combo.currentText() or "").strip(),
                "tube_section": str(tube_section_combo.currentData() or "").strip(),
                "tube_model": tube_model_combo.currentText().strip(),
                "sheet_model": sheet_model_combo.currentText().strip(),
                "sheet_kind": str(payload.get("sheet_kind", "") or "").strip(),
                "sheet_esp_label": str(payload.get("esp_label", "") or "").strip(),
                "bar_size": bar_size_combo.currentText().strip(),
                "quality": str(payload.get("quality", "") or "").strip(),
                "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
                "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
                "hint": hint,
                "_material_base_preco_unit": material_base_cost,
                "operacoes_lista": list(operation_meta.get("operacoes_lista", []) or []),
                "operacoes_fluxo": [dict(item or {}) for item in list(operation_meta.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
            }
            return {
                "kind": "material",
                "stock_item_kind": "raw_material",
                "calc_mode": mode,
                "descricao_base": description,
                "weight_total": round(total_weight, 3),
                "total_cost": round(total_cost, 2),
                "quantity_units": round(qty_units, 2),
                "produto_unid": unit_label,
                "price_per_kg": round(float(payload.get("price_kg", 0) or 0.0), 4),
                "price_base_label": price_base_label,
                "price_base_value": round(price_base_value, 4),
                "stock_metric_value": round(meters_per_unit if price_base_label.upper() == "EUR/M" else weight_each, 4),
                "meters_per_unit": round(meters_per_unit, 3),
                "kg_per_m": round(kg_per_m, 4),
                "length_mm": round(float(length_mm_spin.value() or 0.0), 1),
                "width_mm": round(float(width_mm_spin.value() or 0.0), 1),
                "thickness_mm": round(float(thickness_mm_spin.value() or tube_thickness_spin.value() or 0.0), 2),
                "density": round(float(density_spin.value() or 0.0), 1),
                "diameter_mm": round(float(tube_diameter_spin.value() or diameter_mm_spin.value() or 0.0), 1),
                "manual_unit_price": round(float(manual_unit_price_spin.value() or 0.0), 4),
                "profile_section": str(profile_series_combo.currentData() or "").strip(),
                "profile_size": profile_size_combo.currentText().strip(),
                "tube_section": str(tube_section_combo.currentData() or "").strip(),
                "tube_model": tube_model_combo.currentText().strip(),
                "sheet_model": sheet_model_combo.currentText().strip(),
                "sheet_kind": str(payload.get("sheet_kind", "") or "").strip(),
                "sheet_esp_label": str(payload.get("esp_label", "") or "").strip(),
                "bar_size": bar_size_combo.currentText().strip(),
                "quality": str(payload.get("quality", "") or "").strip(),
                "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
                "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
                "hint": hint,
                "operation_cost_unit": round(operation_cost_unit, 4),
                "operation_time_unit": round(operation_time_unit, 4),
                "line": line,
            }

        def _refresh() -> None:
            if sync["busy"]:
                return
            sync["busy"] = True
            try:
                _toggle_controls()
                payload = dict(_mode_payload() or {})
                if mode_combo.currentText().strip() == "Perfil":
                    if profile_source_combo.currentText().strip() == "Stock MP":
                        record, preview = _current_stock(profile_stock_combo)
                        if record:
                            current_stock_id = str(record.get("id", "") or "").strip()
                            if stock_sync_state["profile"] != current_stock_id:
                                series_txt, size_txt = _profile_series_and_size(record, preview)
                                stock_sync_state["profile"] = current_stock_id
                                profile_meters_spin.setValue(float(record.get("metros", profile_meters_spin.value()) or profile_meters_spin.value() or 0.0))
                                _set_combo_by_data(profile_series_combo, series_txt)
                                _refresh_profile_size_options()
                                if size_txt:
                                    profile_size_combo.setCurrentText(size_txt)
                                profile_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                                profile_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                            identity = _stock_material_identity("Perfil", record, preview)
                            profile_hint.setText(str(payload.get("hint", "") or "").strip())
                            _set_edit_if_auto(desc_edit, desc_state, identity["desc"])
                            _set_edit_if_auto(ref_edit, ref_state, identity["ref"])
                        else:
                            stock_sync_state["profile"] = ""
                    else:
                        stock_sync_state["profile"] = ""
                    if profile_source_combo.currentText().strip() != "Stock MP":
                        profile_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                    if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0 and profile_source_combo.currentText().strip() != "Stock MP":
                        profile_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                    profile_hint.setText(str(payload.get("hint", "") or "").strip())
                elif mode_combo.currentText().strip() == "Tubo":
                    if tube_source_combo.currentText().strip() == "Stock MP":
                        record, preview = _current_stock(tube_stock_combo)
                        if record:
                            current_stock_id = str(record.get("id", "") or "").strip()
                            if stock_sync_state["tube"] != current_stock_id:
                                stock_sync_state["tube"] = current_stock_id
                                tube_quality_edit.setText(_quality_from_record(record))
                                tube_meters_spin.setValue(float(record.get("metros", tube_meters_spin.value()) or tube_meters_spin.value() or 0.0))
                                secao_txt = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().lower()
                                for idx in range(tube_section_combo.count()):
                                    if str(tube_section_combo.itemData(idx) or "").strip().lower() == secao_txt:
                                        tube_section_combo.setCurrentIndex(idx)
                                        break
                                tube_thickness_spin.setValue(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0))
                                if secao_txt == "redondo":
                                    tube_diameter_spin.setValue(float(preview.get("diametro", record.get("diametro", 0)) or 0.0))
                                else:
                                    tube_side_a_spin.setValue(float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                                    width_value = float(preview.get("largura", record.get("largura", 0)) or 0.0)
                                    tube_side_b_spin.setValue(width_value if width_value > 0 else float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                                tube_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                                tube_price_kg_spin.setValue(float(_stock_price_base(record, preview)[1] or 0.0))
                    else:
                        stock_sync_state["tube"] = ""
                    if tube_source_combo.currentText().strip() != "Stock MP":
                        tube_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                        if float(initial.get("price_base_value", initial.get("price_per_kg", 0)) or 0.0) <= 0 and float(payload.get("price_base_value", 0) or 0.0) > 0:
                            tube_price_kg_spin.setValue(float(payload.get("price_base_value", 0) or 0.0))
                    tube_hint.setText(str(payload.get("hint", "") or "").strip())
                elif mode_combo.currentText().strip() == "Chapa":
                    if sheet_source_combo.currentText().strip() == "Stock MP":
                        record, preview = _current_stock(sheet_stock_combo)
                        if record:
                            current_stock_id = str(record.get("id", "") or "").strip()
                            if stock_sync_state["sheet"] != current_stock_id:
                                stock_sync_state["sheet"] = current_stock_id
                                family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                                if family_key:
                                    _set_combo_by_data(sheet_family_combo, family_key)
                                    _refresh_sheet_quality_options()
                                sheet_quality_combo.setCurrentText(_quality_from_record(record))
                                is_checker_stock = bool(str(preview.get("espessura", record.get("espessura", "")) or "").strip().find("/") >= 0) or any(
                                    token in self.backend.desktop_main.norm_text(
                                        " ".join(str(record.get(key, "") or "") for key in ("material", "descricao", "obs", "formato"))
                                    )
                                    for token in ("gota", "xadrez", "antiderrap")
                                )
                                sheet_kind_combo.setCurrentIndex(1 if is_checker_stock else 0)
                                length_mm_spin.setValue(float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                                width_mm_spin.setValue(float(preview.get("largura", record.get("largura", 0)) or 0.0))
                                thickness_mm_spin.setValue(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0))
                                density_spin.setValue(float(preview.get("densidade", density_spin.value()) or density_spin.value() or 7850.0))
                                sheet_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                    else:
                        stock_sync_state["sheet"] = ""
                    if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0 and sheet_source_combo.currentText().strip() != "Stock MP":
                        sheet_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                    sheet_hint.setText(str(payload.get("hint", "") or "").strip())
                elif mode_combo.currentText().strip() == "Cantoneira":
                    if angle_source_combo.currentText().strip() == "Stock MP":
                        record, preview = _current_stock(angle_stock_combo)
                        if record:
                            current_stock_id = str(record.get("id", "") or "").strip()
                            if stock_sync_state["angle"] != current_stock_id:
                                stock_sync_state["angle"] = current_stock_id
                                family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                                if family_key:
                                    _set_combo_by_data(angle_family_combo, family_key)
                                    _refresh_quality_combo(angle_family_combo, angle_quality_combo)
                                angle_quality_combo.setCurrentText(_quality_from_record(record))
                                angle_meters_spin.setValue(float(record.get("metros", angle_meters_spin.value()) or angle_meters_spin.value() or 0.0))
                                a_val = float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0)
                                if a_val > 0:
                                    angle_leg_combo.setCurrentText(str(int(round(a_val))))
                                thickness_txt = _float_text(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0), 0)
                                if thickness_txt:
                                    angle_thickness_combo.setCurrentText(thickness_txt)
                                angle_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                                angle_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                    else:
                        stock_sync_state["angle"] = ""
                    if angle_source_combo.currentText().strip() != "Stock MP":
                        angle_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                        if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0:
                            angle_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                    angle_hint.setText(str(payload.get("hint", "") or "").strip())
                elif mode_combo.currentText().strip() == "Barra":
                    if bar_source_combo.currentText().strip() == "Stock MP":
                        record, preview = _current_stock(bar_stock_combo)
                        if record:
                            current_stock_id = str(record.get("id", "") or "").strip()
                            if stock_sync_state["bar"] != current_stock_id:
                                stock_sync_state["bar"] = current_stock_id
                                family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                                if family_key:
                                    _set_combo_by_data(bar_family_combo, family_key)
                                    _refresh_quality_combo(bar_family_combo, bar_quality_combo)
                                bar_quality_combo.setCurrentText(_quality_from_record(record))
                                bar_meters_spin.setValue(float(record.get("metros", bar_meters_spin.value()) or bar_meters_spin.value() or 0.0))
                                pair_txt = str(preview.get("dimension_label", "") or "").replace(" mm", "").strip()
                                if pair_txt:
                                    bar_size_combo.setCurrentText(pair_txt)
                                bar_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                                bar_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                    else:
                        stock_sync_state["bar"] = ""
                    if bar_source_combo.currentText().strip() != "Stock MP":
                        bar_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                        if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0:
                            bar_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                    bar_hint.setText(str(payload.get("hint", "") or "").strip())
                _apply_auto_texts(payload)
                result = _compute()
                summary_weight.setText(f"Peso total: {float(result.get('weight_total', 0) or 0):.2f} kg")
                summary_cost.setText(f"Custo total: {_fmt_eur(float(result.get('total_cost', 0) or 0))}")
                summary_desc.setText(str((result.get("line") or {}).get("descricao", "") or "").strip())
                _refresh_operation_cost_hint(float((result.get("line") or {}).get("_material_base_preco_unit", 0) or 0))
            finally:
                sync["busy"] = False

        for widget in (
            mode_combo,
            desc_edit,
            ref_edit,
            qty_spin,
            profile_source_combo,
            profile_stock_combo,
            profile_series_combo,
            profile_size_combo,
            profile_meters_spin,
            profile_kg_per_m_spin,
            profile_price_kg_spin,
            tube_source_combo,
            tube_stock_combo,
            tube_quality_edit,
            tube_section_combo,
            tube_model_combo,
            tube_side_a_spin,
            tube_side_b_spin,
            tube_diameter_spin,
            tube_thickness_spin,
            tube_meters_spin,
            tube_kg_per_m_spin,
            tube_price_kg_spin,
            sheet_source_combo,
            sheet_stock_combo,
            sheet_family_combo,
            sheet_quality_combo,
            sheet_kind_combo,
            sheet_model_combo,
            length_mm_spin,
            width_mm_spin,
            thickness_mm_spin,
            density_spin,
            sheet_price_kg_spin,
            angle_source_combo,
            angle_stock_combo,
            angle_family_combo,
            angle_quality_combo,
            angle_leg_combo,
            angle_thickness_combo,
            angle_meters_spin,
            angle_kg_per_m_spin,
            angle_price_kg_spin,
            bar_source_combo,
            bar_stock_combo,
            bar_family_combo,
            bar_quality_combo,
            bar_size_combo,
            bar_meters_spin,
            bar_kg_per_m_spin,
            bar_price_kg_spin,
            rebar_meters_spin,
            diameter_mm_spin,
            rebar_price_kg_spin,
            unit_combo,
            manual_unit_price_spin,
        ):
            if isinstance(widget, QLineEdit):
                widget.textChanged.connect(_refresh)
            elif isinstance(widget, QComboBox):
                widget.currentTextChanged.connect(_refresh)
            else:
                widget.valueChanged.connect(lambda _value: _refresh())

        desc_edit.textEdited.connect(lambda _text: desc_state.__setitem__("manual", True))
        ref_edit.textEdited.connect(lambda _text: ref_state.__setitem__("manual", True))
        profile_series_combo.currentIndexChanged.connect(lambda _idx: _refresh_profile_size_options())
        tube_section_combo.currentIndexChanged.connect(lambda _idx: _refresh_tube_model_options())
        tube_model_combo.currentTextChanged.connect(lambda _text: _apply_tube_model_selection())
        angle_leg_combo.currentTextChanged.connect(lambda _text: _refresh_angle_thickness_options())
        sheet_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_sheet_quality_options(preserve_current=False), _refresh()))
        sheet_kind_combo.currentIndexChanged.connect(lambda _idx: (_refresh_sheet_model_options(), _apply_sheet_model_selection(), _refresh()))
        sheet_model_combo.currentTextChanged.connect(lambda _text: _apply_sheet_model_selection())
        angle_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_quality_combo(angle_family_combo, angle_quality_combo, preserve_current=False), _refresh()))
        bar_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_quality_combo(bar_family_combo, bar_quality_combo, preserve_current=False), _refresh()))

        def _open_price_table() -> None:
            mode = mode_combo.currentText().strip()
            preferred_id = ""
            if mode == "Perfil":
                preferred_id = str((_current_stock(profile_stock_combo)[0] or {}).get("id", "") or "").strip()
            elif mode == "Tubo":
                preferred_id = str((_current_stock(tube_stock_combo)[0] or {}).get("id", "") or "").strip()
            elif mode == "Chapa":
                preferred_id = str((_current_stock(sheet_stock_combo)[0] or {}).get("id", "") or "").strip()
            elif mode == "Cantoneira":
                preferred_id = str((_current_stock(angle_stock_combo)[0] or {}).get("id", "") or "").strip()
            elif mode == "Barra":
                preferred_id = str((_current_stock(bar_stock_combo)[0] or {}).get("id", "") or "").strip()
            self._material_price_manager_dialog(mode, preferred_id, parent=dialog)
            _fill_stock_combo(profile_stock_combo, "Perfil", preferred_id if mode == "Perfil" else "")
            _fill_stock_combo(tube_stock_combo, "Tubo", preferred_id if mode == "Tubo" else "")
            _fill_stock_combo(sheet_stock_combo, "Chapa", preferred_id if mode == "Chapa" else "")
            _fill_stock_combo(angle_stock_combo, "Cantoneira", preferred_id if mode == "Cantoneira" else "")
            _fill_stock_combo(bar_stock_combo, "Barra", preferred_id if mode == "Barra" else "")
            _refresh()

        price_table_btn.clicked.connect(_open_price_table)
        op_detail_btn.clicked.connect(_edit_operation_costs)
        op_profiles_btn.clicked.connect(lambda: (_open_operation_cost_profiles_dialog(dialog, self.backend), _refresh()))
        operation_edit.textChanged.connect(lambda _text: _refresh())

        _refresh()
        if dialog.exec() != QDialog.Accepted:
            return None
        result = _compute()
        if not isinstance(result.get("line"), dict):
            return None
        mode = mode_combo.currentText().strip()
        if mode == "Perfil":
            stock_record, stock_preview = _current_stock(profile_stock_combo)
            if profile_source_combo.currentText().strip() == "Stock MP" and stock_record:
                updated = self._sync_stock_price_from_context(
                    str(stock_record.get("id", "") or "").strip(),
                    float(result.get("price_per_kg", 0) or 0.0),
                    _stock_price_kg(stock_record, stock_preview),
                    parent=dialog,
                )
                if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                    result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["price_base_value"] = float(updated.get("p_compra", result.get("price_base_value", 0)) or 0.0)
                    result["line"]["price_base_value"] = float(updated.get("p_compra", result["line"].get("price_base_value", 0)) or 0.0)
                    result["price_base_label"] = str(updated.get("base_label", "EUR/kg") or "EUR/kg")
                    result["line"]["price_base_label"] = str(updated.get("base_label", "EUR/kg") or "EUR/kg")
        elif mode == "Tubo":
            stock_record, stock_preview = _current_stock(tube_stock_combo)
            if tube_source_combo.currentText().strip() == "Stock MP" and stock_record:
                updated = self._sync_stock_price_from_context(
                    str(stock_record.get("id", "") or "").strip(),
                    float(result.get("price_per_kg", 0) or 0.0),
                    _stock_price_kg(stock_record, stock_preview),
                    parent=dialog,
                )
                if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                    result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["price_base_value"] = float(updated.get("p_compra", result.get("price_base_value", 0)) or 0.0)
                    result["line"]["price_base_value"] = float(updated.get("p_compra", result["line"].get("price_base_value", 0)) or 0.0)
                    result["price_base_label"] = str(updated.get("base_label", "EUR/m") or "EUR/m")
                    result["line"]["price_base_label"] = str(updated.get("base_label", "EUR/m") or "EUR/m")
        elif mode == "Chapa":
            stock_record, stock_preview = _current_stock(sheet_stock_combo)
            if sheet_source_combo.currentText().strip() == "Stock MP" and stock_record:
                updated = self._sync_stock_price_from_context(
                    str(stock_record.get("id", "") or "").strip(),
                    float(result.get("price_per_kg", 0) or 0.0),
                    _stock_price_kg(stock_record, stock_preview),
                    parent=dialog,
                )
                if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                    result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
        elif mode == "Cantoneira":
            stock_record, stock_preview = _current_stock(angle_stock_combo)
            if angle_source_combo.currentText().strip() == "Stock MP" and stock_record:
                updated = self._sync_stock_price_from_context(
                    str(stock_record.get("id", "") or "").strip(),
                    float(result.get("price_per_kg", 0) or 0.0),
                    _stock_price_kg(stock_record, stock_preview),
                    parent=dialog,
                )
                if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                    result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
        elif mode == "Barra":
            stock_record, stock_preview = _current_stock(bar_stock_combo)
            if bar_source_combo.currentText().strip() == "Stock MP" and stock_record:
                updated = self._sync_stock_price_from_context(
                    str(stock_record.get("id", "") or "").strip(),
                    float(result.get("price_per_kg", 0) or 0.0),
                    _stock_price_kg(stock_record, stock_preview),
                    parent=dialog,
                )
                if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                    result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                    result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
        return result

    def _labor_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
        dialog.setWindowTitle("Item mão de obra")
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        desc_edit = QLineEdit(str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip())
        op_combo = QComboBox()
        op_combo.addItems(list(self.presets.get("operacoes", []) or []) or ["Serralharia", "Pintura", "Retrabalho", "Montagem"])
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
        line = self._structure_line(
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

    def _consumable_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
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
        line = self._structure_line(
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

    def _product_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        wrapped_initial = dict(initial or {})
        line_initial = dict(wrapped_initial.get("line") or {})
        if line_initial:
            for key, value in wrapped_initial.items():
                if key == "line":
                    continue
                if key not in line_initial and value not in (None, ""):
                    line_initial[key] = value
            initial = line_initial
        else:
            initial = wrapped_initial
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
        dialog.setWindowTitle("Produto de stock")
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        search_box = QWidget()
        search_layout = QHBoxLayout(search_box)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)
        search_icon = QLabel("🔍")
        search_icon.setFixedWidth(24)
        search_icon.setAlignment(Qt.AlignCenter)
        search_icon.setStyleSheet("font-size: 16px; color: #10253d;")
        search_edit = QLineEdit()
        search_edit.setPlaceholderText("Pesquisar por codigo, descricao, medida, tipo...")
        search_layout.addWidget(search_icon)
        search_layout.addWidget(search_edit, 1)
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMaxVisibleItems(18)
        combo.addItem("")
        product_rows_all = [dict(row) for row in list(self.backend.ne_product_options("") or []) if isinstance(row, dict)]
        product_rows = list(product_rows_all)
        by_code = {}

        def _product_search_text(value: str) -> str:
            text = unicodedata.normalize("NFKD", str(value or ""))
            text = "".join(ch for ch in text if not unicodedata.combining(ch))
            return text.casefold().strip()

        def _product_row_haystack(row: dict) -> str:
            return _product_search_text(
                " ".join(
                    str(row.get(key, "") or "")
                    for key in (
                        "codigo",
                        "descricao",
                        "categoria",
                        "subcat",
                        "tipo",
                        "marca",
                        "modelo",
                        "dimensoes",
                        "obs",
                    )
                )
            )

        def _filter_product_rows(query: str) -> list[dict]:
            needle = _product_search_text(query)
            if not needle:
                return list(product_rows_all[:120])
            terms = [term for term in needle.split() if term]
            matches = []
            for row in product_rows_all:
                haystack = _product_row_haystack(row)
                if all(term in haystack for term in terms):
                    matches.append(row)
                if len(matches) >= 120:
                    break
            return matches

        def _reload_product_options(select_code: str = "", search_text: str = "") -> None:
            nonlocal product_rows, by_code
            product_rows = [dict(row) for row in _filter_product_rows(search_text)]
            selected_code = str(select_code or "").strip()
            if selected_code and not any(str(row.get("codigo", "") or "").strip() == selected_code for row in product_rows):
                selected_row = next(
                    (dict(row) for row in product_rows_all if str(row.get("codigo", "") or "").strip() == selected_code),
                    None,
                )
                if selected_row:
                    product_rows.insert(0, selected_row)
            by_code = {}
            current_text = combo.currentText().strip()
            keep_text = "" if str(search_text or "").strip() else str(current_text or "").strip()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            for row in product_rows:
                code = str(row.get("codigo", "") or "").strip()
                if not code:
                    continue
                by_code[code] = row
                combo.addItem(f"{code} - {str(row.get('descricao', '') or '').strip()}".strip(" -"), code)
            if select_code:
                for index in range(combo.count()):
                    if str(combo.itemData(index) or "").strip() == select_code:
                        combo.setCurrentIndex(index)
                        break
            elif keep_text:
                combo.setEditText(keep_text)
            else:
                combo.setCurrentIndex(0)
            combo.blockSignals(False)

        product_search_timer = QTimer(dialog)
        product_search_timer.setSingleShot(True)

        def _apply_product_search() -> None:
            _reload_product_options(search_text=search_edit.text())
            search_edit.setFocus()
            _refresh()

        def _on_product_search(_text: str) -> None:
            product_search_timer.start(180)

        def _new_product_dialog() -> dict | None:
            product_dialog = QDialog(dialog)
            product_dialog.setWindowTitle("Novo produto")
            product_layout = QVBoxLayout(product_dialog)
            product_form = QFormLayout()
            presets = dict(self.backend.product_catalog_options() or {})
            code_edit = QLineEdit(str(self.backend.product_next_code() or "").strip())
            desc_edit = QLineEdit()
            category_combo = QComboBox()
            category_combo.setEditable(True)
            category_combo.addItems([str(value or "").strip() for value in list(presets.get("categorias", []) or []) if str(value or "").strip()])
            subcat_combo = QComboBox()
            subcat_combo.setEditable(True)
            subcat_combo.addItems([str(value or "").strip() for value in list(presets.get("subcats", []) or []) if str(value or "").strip()])
            type_combo = QComboBox()
            type_combo.setEditable(True)
            type_combo.addItems([str(value or "").strip() for value in list(presets.get("tipos", []) or []) if str(value or "").strip()])
            unit_combo = QComboBox()
            unit_combo.setEditable(True)
            unit_combo.addItems([str(value or "").strip() for value in list(presets.get("unidades", []) or []) if str(value or "").strip()])
            unit_combo.setCurrentText("UN")
            dim_edit = QLineEdit()
            meters_spin = QDoubleSpinBox()
            meters_spin.setRange(0.0, 1000000.0)
            meters_spin.setDecimals(4)
            weight_spin = QDoubleSpinBox()
            weight_spin.setRange(0.0, 1000000.0)
            weight_spin.setDecimals(4)
            qty_spin_new = QDoubleSpinBox()
            qty_spin_new.setRange(0.0, 1000000.0)
            qty_spin_new.setDecimals(2)
            qty_spin_new.setValue(0.0)
            alert_spin = QDoubleSpinBox()
            alert_spin.setRange(0.0, 1000000.0)
            alert_spin.setDecimals(2)
            buy_spin = QDoubleSpinBox()
            buy_spin.setRange(0.0, 1000000.0)
            buy_spin.setDecimals(4)
            buy_total_spin = QDoubleSpinBox()
            buy_total_spin.setRange(0.0, 1000000000.0)
            buy_total_spin.setDecimals(4)
            pvp1_spin = QDoubleSpinBox()
            pvp1_spin.setRange(0.0, 1000000.0)
            pvp1_spin.setDecimals(4)
            pvp2_spin = QDoubleSpinBox()
            pvp2_spin.setRange(0.0, 1000000.0)
            pvp2_spin.setDecimals(4)
            obs_edit = QLineEdit()

            def _set_combo_items(combo: QComboBox, values: list[str], current_text: str = "") -> None:
                combo.blockSignals(True)
                combo.clear()
                combo.addItems([str(value or "").strip() for value in list(values or []) if str(value or "").strip()])
                combo.setCurrentText(current_text)
                combo.blockSignals(False)

            def _sync_catalog_combos() -> None:
                current_category = category_combo.currentText().strip()
                current_subcat = subcat_combo.currentText().strip()
                current_type = type_combo.currentText().strip()
                subcat_presets = dict(self.backend.product_catalog_options(current_category, current_subcat) or {})
                _set_combo_items(subcat_combo, subcat_presets.get("subcats", []), current_subcat)
                current_subcat = subcat_combo.currentText().strip()
                type_presets = dict(self.backend.product_catalog_options(current_category, current_subcat) or {})
                _set_combo_items(type_combo, type_presets.get("tipos", []), current_type)

            product_form.addRow("Codigo", code_edit)
            product_form.addRow("Descricao", desc_edit)
            product_form.addRow("Categoria", category_combo)
            product_form.addRow("Subcat.", subcat_combo)
            product_form.addRow("Tipo", type_combo)
            product_form.addRow("Unid.", unit_combo)
            product_form.addRow("Dimensoes", dim_edit)
            product_form.addRow("Metros/Unid.", meters_spin)
            product_form.addRow("Peso/Unid.", weight_spin)
            product_form.addRow("Quantidade", qty_spin_new)
            product_form.addRow("Alerta", alert_spin)
            product_form.addRow("Compra/Unid. (EUR)", buy_spin)
            product_form.addRow("Compra total stock", buy_total_spin)
            product_form.addRow("PVP1", pvp1_spin)
            product_form.addRow("PVP2", pvp2_spin)
            product_form.addRow("Observacoes", obs_edit)
            product_layout.addLayout(product_form)
            price_note = QLabel("")
            price_note.setWordWrap(True)
            price_note.setProperty("role", "muted")
            product_layout.addWidget(price_note)
            note = QLabel("Se o produto ainda nao existir em stock, podes cria-lo aqui com stock inicial zero e reutiliza-lo logo no conjunto.")
            note.setWordWrap(True)
            note.setProperty("role", "muted")
            product_layout.addWidget(note)
            product_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            product_buttons.accepted.connect(product_dialog.accept)
            product_buttons.rejected.connect(product_dialog.reject)
            product_layout.addWidget(product_buttons)
            price_sync = {"busy": False, "last_source": "unit"}

            def _refresh_product_price_note() -> None:
                qty_value = float(qty_spin_new.value() or 0.0)
                unit_txt = unit_combo.currentText().strip() or "UN"
                unit_price = float(buy_spin.value() or 0.0)
                total_price = float(buy_total_spin.value() or 0.0)
                if qty_value > 0:
                    price_note.setText(
                        f"Compra: {_fmt_eur(unit_price)}/{unit_txt} | total stock {_fmt_eur(total_price)} "
                        f"| {qty_value:.2f} {unit_txt}."
                    )
                else:
                    price_note.setText(f"Compra: {_fmt_eur(unit_price)}/{unit_txt}. Define quantidade para calcular o total.")

            def _sync_buy_total_from_unit() -> None:
                if price_sync["busy"]:
                    return
                price_sync["busy"] = True
                try:
                    price_sync["last_source"] = "unit"
                    buy_total_spin.setValue(round(float(qty_spin_new.value() or 0.0) * float(buy_spin.value() or 0.0), 4))
                    _refresh_product_price_note()
                finally:
                    price_sync["busy"] = False

            def _sync_buy_unit_from_total() -> None:
                if price_sync["busy"]:
                    return
                price_sync["busy"] = True
                try:
                    price_sync["last_source"] = "total"
                    qty_value = float(qty_spin_new.value() or 0.0)
                    if qty_value > 0:
                        buy_spin.setValue(round(float(buy_total_spin.value() or 0.0) / qty_value, 4))
                    _refresh_product_price_note()
                finally:
                    price_sync["busy"] = False

            def _sync_buy_after_qty_change() -> None:
                if price_sync.get("last_source") == "total":
                    _sync_buy_unit_from_total()
                else:
                    _sync_buy_total_from_unit()

            category_combo.currentTextChanged.connect(lambda _text: _sync_catalog_combos())
            subcat_combo.currentTextChanged.connect(lambda _text: _sync_catalog_combos())
            unit_combo.currentTextChanged.connect(lambda _text: _refresh_product_price_note())
            qty_spin_new.valueChanged.connect(lambda _value: _sync_buy_after_qty_change())
            buy_spin.valueChanged.connect(lambda _value: _sync_buy_total_from_unit())
            buy_total_spin.valueChanged.connect(lambda _value: _sync_buy_unit_from_total())
            _sync_catalog_combos()
            _refresh_product_price_note()
            if product_dialog.exec() != QDialog.Accepted:
                return None
            try:
                return dict(
                    self.backend.product_save(
                        {
                            "codigo": code_edit.text().strip(),
                            "descricao": desc_edit.text().strip(),
                            "categoria": category_combo.currentText().strip(),
                            "subcat": subcat_combo.currentText().strip(),
                            "tipo": type_combo.currentText().strip() or "Montagem",
                            "unid": unit_combo.currentText().strip() or "UN",
                            "dimensoes": dim_edit.text().strip(),
                            "metros_unidade": float(meters_spin.value() or 0.0),
                            "peso_unid": float(weight_spin.value() or 0.0),
                            "qty": float(qty_spin_new.value() or 0.0),
                            "alerta": float(alert_spin.value() or 0.0),
                            "p_compra": float(buy_spin.value() or 0.0),
                            "pvp1": float(pvp1_spin.value() or 0.0),
                            "pvp2": float(pvp2_spin.value() or 0.0),
                            "obs": obs_edit.text().strip(),
                        }
                    )
                    or {}
                )
            except Exception as exc:
                QMessageBox.critical(product_dialog, "Produtos", str(exc))
                return None

        wanted_code = str(initial.get("produto_codigo", "") or initial.get("ref_externa", "") or "").strip()
        wanted_desc = str(initial.get("descricao", "") or "").strip()
        _reload_product_options(wanted_code)
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(float(initial.get("quantity_units", initial.get("qtd", 1)) or 1))
        manual_unit_combo = QComboBox()
        manual_unit_combo.setEditable(True)
        manual_unit_combo.addItems(["UN", "MT", "KG", "L", "CJ"])
        manual_unit_combo.setCurrentText(str(initial.get("produto_unid", "") or "UN").strip() or "UN")
        manual_price_spin = QDoubleSpinBox()
        manual_price_spin.setRange(0.0, 1000000.0)
        manual_price_spin.setDecimals(4)
        manual_price_spin.setValue(float(initial.get("preco_unit", 0) or 0.0))
        manual_state = {"unit_dirty": False, "price_dirty": False}
        form.addRow("Pesquisar", search_box)
        form.addRow("Produto", combo)
        form.addRow("Quantidade", qty_spin)
        form.addRow("Unid. manual", manual_unit_combo)
        form.addRow("Preco manual", manual_price_spin)
        layout.addLayout(form)
        info_label = QLabel("")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        actions = QHBoxLayout()
        new_product_btn = QPushButton("Novo artigo")
        new_product_btn.setProperty("variant", "secondary")
        actions.addWidget(new_product_btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        def _current_product() -> dict | None:
            code = str(combo.currentData() or "").strip()
            if not code:
                text = combo.currentText().strip()
                code = text.split(" - ", 1)[0].strip()
            return by_code.get(code)

        if not _current_product():
            fallback_text = ""
            if wanted_code and wanted_desc:
                fallback_text = f"{wanted_code} - {wanted_desc}"
            elif wanted_code:
                fallback_text = wanted_code
            elif wanted_desc:
                fallback_text = wanted_desc
            if fallback_text:
                combo.blockSignals(True)
                combo.setCurrentText(fallback_text)
                combo.blockSignals(False)

        def _refresh() -> None:
            row = _current_product() or {}
            if row:
                sale_price = float(row.get("preco_venda", row.get("pvp1", row.get("preco_unid", row.get("preco", 0)))) or 0)
                total = float(qty_spin.value() or 0) * sale_price
                info_label.setText(
                    f"Unid.: {str(row.get('unid', '-') or '-')} | PVP1: {_fmt_eur(sale_price)} | "
                    f"Total: {_fmt_eur(total)}"
                )
            else:
                desc = combo.currentText().strip()
                total = float(qty_spin.value() or 0) * float(manual_price_spin.value() or 0)
                if desc:
                    info_label.setText(f"Produto novo para compra/cotacao | Total: {_fmt_eur(total)}")
                else:
                    info_label.setText("Seleciona um produto existente, cria um novo artigo ou escreve uma descricao para cotacao.")

        combo.currentTextChanged.connect(lambda _t: _refresh())
        qty_spin.valueChanged.connect(lambda _v: _refresh())
        manual_unit_combo.currentTextChanged.connect(lambda _t: manual_state.__setitem__("unit_dirty", True))
        manual_price_spin.valueChanged.connect(lambda _v: (manual_state.__setitem__("price_dirty", True), _refresh()))
        def _create_and_select_product() -> None:
            detail = _new_product_dialog()
            if not isinstance(detail, dict) or not str(detail.get("codigo", "") or "").strip():
                return
            product_rows_all[:] = [dict(row) for row in list(self.backend.ne_product_options("") or []) if isinstance(row, dict)]
            _reload_product_options(str(detail.get("codigo", "") or "").strip())
            _refresh()
        new_product_btn.clicked.connect(_create_and_select_product)
        product_search_timer.timeout.connect(_apply_product_search)
        search_edit.textChanged.connect(_on_product_search)
        _refresh()
        if dialog.exec() != QDialog.Accepted:
            return None
        product = _current_product()
        if isinstance(product, dict):
            line = self._structure_product_line(product, float(qty_spin.value() or 0))
            product_code = str(product.get("codigo", "") or "").strip()
            preserve_existing_line = bool(wanted_code and product_code == wanted_code)
            if isinstance(line, dict):
                manual_price = round(float(manual_price_spin.value() or 0.0), 4)
                manual_unit = manual_unit_combo.currentText().strip()
                if manual_price > 0 and (manual_state.get("price_dirty") or preserve_existing_line):
                    line["preco_unit"] = manual_price
                if manual_unit and (manual_state.get("unit_dirty") or preserve_existing_line):
                    line["produto_unid"] = manual_unit
        else:
            manual_desc = combo.currentText().strip()
            if not manual_desc:
                return None
            manual_code = ""
            if " - " in manual_desc:
                manual_code, manual_desc = [chunk.strip() for chunk in manual_desc.split(" - ", 1)]
            line = {
                "tipo_item": self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
                "stock_item_kind": "product",
                "produto_codigo": manual_code,
                "ref_externa": manual_code,
                "descricao": manual_desc,
                "produto_unid": manual_unit_combo.currentText().strip() or "UN",
                "qtd": round(float(qty_spin.value() or 0), 2),
                "tempo_peca_min": 0.0,
                "preco_unit": round(float(manual_price_spin.value() or 0.0), 4),
                "operacao": "Montagem",
                "_product_pending_create": True,
            }
        if not isinstance(line, dict):
            return None
        return {
            "kind": "product",
            "quantity_units": round(float(qty_spin.value() or 0), 2),
            "total_cost": round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2),
            "line": line,
        }

    def _assembly_item_kind_label(self, kind: str) -> str:
        return {
            "material": "Material",
            "labor": "Mao de obra",
            "consumable": "Consumivel",
            "product": "Produto stock",
        }.get(str(kind or "").strip(), "Item")

    def _assembly_item_kind_from_line(self, row: dict | None = None) -> str:
        payload = dict(row or {})
        explicit = str(payload.get("kind", "") or "").strip()
        if explicit in {"material", "labor", "consumable", "product"}:
            return explicit
        line = dict(payload.get("line") or payload)
        if self.backend.desktop_main.orc_line_is_product(line):
            return "product"
        if self.backend.desktop_main.orc_line_is_piece(line):
            return "material"
        unit_txt = str(line.get("produto_unid", "") or "").strip().lower()
        if unit_txt == "h":
            return "labor"
        return "consumable"

    def _wrap_assembly_item(self, row: dict | None = None) -> dict:
        payload = dict(row or {})
        if isinstance(payload.get("line"), dict):
            line = dict(payload.get("line") or {})
            total_cost = float(payload.get("total_cost", line.get("qtd", 0)) or 0) * 1.0
            if float(payload.get("total_cost", 0) or 0) <= 0:
                total_cost = round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2)
            payload["kind"] = self._assembly_item_kind_from_line(payload)
            payload["total_cost"] = round(total_cost, 2)
            return payload
        line = payload
        return {
            "kind": self._assembly_item_kind_from_line(line),
            "descricao_base": str(line.get("descricao", "") or "").strip(),
            "total_cost": round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2),
            "weight_total": 0.0,
            "line": dict(line),
        }

    def _pick_assembly_item_kind(self, parent: QWidget | None = None) -> str:
        labels = ["Material", "Mao de obra", "Consumivel", "Produto stock"]
        selected, ok = QInputDialog.getItem(
            parent if isinstance(parent, QWidget) else self,
            "Adicionar item ao conjunto",
            "Tipo de item",
            labels,
            0,
            False,
        )
        if not ok:
            return ""
        return {
            "Material": "material",
            "Mao de obra": "labor",
            "Consumivel": "consumable",
            "Produto stock": "product",
        }.get(str(selected or "").strip(), "")

    def _open_assembly_item_editor(self, kind: str, current: dict | None = None, parent: QWidget | None = None) -> dict | None:
        if kind == "material":
            return self._material_assembly_item_dialog(current, parent=parent)
        if kind == "labor":
            return self._labor_assembly_item_dialog(current, parent=parent)
        if kind == "consumable":
            return self._consumable_assembly_item_dialog(current, parent=parent)
        if kind == "product":
            return self._product_assembly_item_dialog(current, parent=parent)
        return None

    def _calculated_assembly_builder_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Conjunto calculado")
        dialog.resize(1040, 840)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        intro = QLabel(
            "Cria um conjunto como mini-projeto: materiais, mao de obra, consumiveis e produtos. "
            "Cada conjunto criado aqui fica guardado no backend; se ativares a opcao abaixo, fica tambem marcado como template reutilizavel."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        header_tabs = QTabWidget()
        identity_tab = QWidget()
        header_form = QFormLayout(identity_tab)
        header_form.setContentsMargins(10, 10, 10, 10)
        header_form.setHorizontalSpacing(10)
        header_form.setVerticalSpacing(6)
        code_edit = QLineEdit(str(initial.get("codigo", "") or f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}").strip())
        try:
            suggested_param = str(self.backend.conjunto_next_param_codigo() or "").strip()
        except Exception:
            suggested_param = ""
        param_edit = QLineEdit(str(initial.get("param_codigo", "") or suggested_param or "0001").strip())
        param_edit.setReadOnly(True)
        param_edit.setToolTip("Codigo sequencial e permanente da parametrizacao do conjunto")
        name_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
        name_edit.setPlaceholderText("Ex.: Construcao de Bascula")
        margin_spin = QDoubleSpinBox()
        margin_spin.setRange(0.0, 500.0)
        margin_spin.setDecimals(2)
        margin_spin.setSuffix(" %")
        margin_spin.setValue(float(initial.get("margem_perc", 30.0) or 30.0))
        save_template_check = QCheckBox("Marcar tambem como template de conjunto")
        save_template_check.setChecked(bool(initial.get("template", False)))
        notes_edit = QTextEdit()
        notes_edit.setMaximumHeight(70)
        notes_edit.setPlainText(str(initial.get("notas", "") or "").strip())
        header_form.addRow("Codigo conjunto", code_edit)
        header_form.addRow("Codigo parametrizacao", param_edit)
        header_form.addRow("Descricao", name_edit)
        header_form.addRow("Margem", margin_spin)
        header_form.addRow("Notas", notes_edit)
        header_form.addRow("", save_template_check)
        header_tabs.addTab(identity_tab, "Identificacao e valor")

        technical = dict(initial.get("ficha_tecnica", {}) or {})
        technical_tab = QWidget()
        technical_grid = QGridLayout(technical_tab)
        technical_grid.setContentsMargins(10, 10, 10, 10)
        technical_grid.setHorizontalSpacing(10)
        technical_grid.setVerticalSpacing(6)

        family_combo = QComboBox()
        family_combo.setEditable(True)
        family_combo.addItems([
            "Equipamento de pesagem",
            "Quiosque multimedia",
            "Caixilharia",
            "Estrutura metalica",
            "Maquina / equipamento",
            "Mobiliario tecnico",
            "Outro produto fabricado",
        ])
        family_value = str(technical.get("familia_produto", "") or "").strip()
        if family_value:
            family_combo.setCurrentText(family_value)
        application_edit = QLineEdit(str(technical.get("aplicacao", "") or "").strip())
        application_edit.setPlaceholderText("Ex.: pesagem industrial, atendimento publico, fachada exterior")
        model_edit = QLineEdit(str(technical.get("modelo_versao", "") or "").strip())
        model_edit.setPlaceholderText("Modelo, variante ou revisao")
        configuration_edit = QLineEdit(str(technical.get("configuracao", "") or "").strip())
        configuration_edit.setPlaceholderText("Ex.: 1500 kg / visor remoto / 2 folhas / RAL 7016")
        dimensions_edit = QLineEdit(str(technical.get("dimensoes_gerais", "") or "").strip())
        dimensions_edit.setPlaceholderText("C x L x A, vao, capacidade ou formato relevante")
        finishes_edit = QLineEdit(str(technical.get("materiais_acabamentos", "") or "").strip())
        finishes_edit.setPlaceholderText("Materiais principais, cor e acabamento")

        characteristics_edit = QTextEdit()
        characteristics_edit.setMaximumHeight(66)
        characteristics_edit.setPlaceholderText("Funcoes, desempenho, opcoes e caracteristicas essenciais")
        characteristics_edit.setPlainText(str(technical.get("caracteristicas", "") or "").strip())
        installation_edit = QTextEdit()
        installation_edit.setMaximumHeight(66)
        installation_edit.setPlaceholderText("Alimentacao, fixacao, ligacoes, ambiente e pre-requisitos")
        installation_edit.setPlainText(str(technical.get("requisitos_instalacao", "") or "").strip())
        standards_edit = QTextEdit()
        standards_edit.setMaximumHeight(58)
        standards_edit.setPlaceholderText("Normas, diretivas, classe, IP, CE ou requisitos do cliente")
        standards_edit.setPlainText(str(technical.get("normas_conformidade", "") or "").strip())
        quality_edit = QTextEdit()
        quality_edit.setMaximumHeight(58)
        quality_edit.setPlaceholderText("Inspecoes, ensaios, tolerancias e criterios de aceitacao")
        quality_edit.setPlainText(str(technical.get("controlo_qualidade", "") or "").strip())

        technical_grid.addWidget(QLabel("Familia de produto"), 0, 0)
        technical_grid.addWidget(family_combo, 0, 1)
        technical_grid.addWidget(QLabel("Aplicacao / destino"), 0, 2)
        technical_grid.addWidget(application_edit, 0, 3)
        technical_grid.addWidget(QLabel("Modelo / versao"), 1, 0)
        technical_grid.addWidget(model_edit, 1, 1)
        technical_grid.addWidget(QLabel("Configuracao principal"), 1, 2)
        technical_grid.addWidget(configuration_edit, 1, 3)
        technical_grid.addWidget(QLabel("Dimensoes / capacidade"), 2, 0)
        technical_grid.addWidget(dimensions_edit, 2, 1)
        technical_grid.addWidget(QLabel("Materiais / acabamentos"), 2, 2)
        technical_grid.addWidget(finishes_edit, 2, 3)
        technical_grid.addWidget(QLabel("Caracteristicas"), 3, 0)
        technical_grid.addWidget(characteristics_edit, 3, 1)
        technical_grid.addWidget(QLabel("Instalacao"), 3, 2)
        technical_grid.addWidget(installation_edit, 3, 3)
        technical_grid.addWidget(QLabel("Normas / conformidade"), 4, 0)
        technical_grid.addWidget(standards_edit, 4, 1)
        technical_grid.addWidget(QLabel("Controlo de qualidade"), 4, 2)
        technical_grid.addWidget(quality_edit, 4, 3)
        technical_grid.setColumnStretch(1, 1)
        technical_grid.setColumnStretch(3, 1)
        header_tabs.addTab(technical_tab, "Ficha tecnica do produto")
        layout.addWidget(header_tabs)

        items: list[dict] = [self._wrap_assembly_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["Categoria", "Descricao", "Ref./Cod.", "Qtd", "Unid", "Peso", "Preco", "Total"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(1,), contents=(0, 2, 3, 4, 5, 6, 7))
        layout.addWidget(table, 1)

        actions = QGridLayout()
        actions.setHorizontalSpacing(8)
        actions.setVerticalSpacing(6)
        add_material_btn = QPushButton("Material")
        add_labor_btn = QPushButton("Mao de obra")
        add_consumable_btn = QPushButton("Consumivel")
        add_product_btn = QPushButton("Produto stock")
        add_laser_batch_btn = QPushButton("Lote DXF/DWG")
        add_step_igs_btn = QPushButton("STEP/IGS")
        add_structure_btn = QPushButton("Estrutura metalica")
        edit_btn = QPushButton("Editar item")
        edit_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover item")
        remove_btn.setProperty("variant", "danger")
        for idx, button in enumerate((
            add_material_btn,
            add_labor_btn,
            add_consumable_btn,
            add_product_btn,
            add_laser_batch_btn,
            add_step_igs_btn,
            add_structure_btn,
            edit_btn,
            remove_btn,
        )):
            button.setProperty("compact", "true")
            actions.addWidget(button, idx // 3, idx % 3)
        layout.addLayout(actions)

        summary_card = CardFrame()
        summary_card.set_tone("info")
        summary_layout = QGridLayout(summary_card)
        summary_layout.setContentsMargins(10, 8, 10, 8)
        summary_layout.setHorizontalSpacing(12)
        summary_layout.setVerticalSpacing(6)
        material_total_label = QLabel("0,00 EUR")
        labor_total_label = QLabel("0,00 EUR")
        consumable_total_label = QLabel("0,00 EUR")
        product_total_label = QLabel("0,00 EUR")
        subtotal_label = QLabel("0,00 EUR")
        final_total_label = QLabel("0,00 EUR")
        for text, widget, row, col in (
            ("Materiais", material_total_label, 0, 0),
            ("Mao de obra", labor_total_label, 0, 2),
            ("Consumiveis", consumable_total_label, 1, 0),
            ("Produtos", product_total_label, 1, 2),
            ("Subtotal custo", subtotal_label, 2, 0),
            ("Preco final c/margem", final_total_label, 2, 2),
        ):
            summary_layout.addWidget(QLabel(text), row, col)
            summary_layout.addWidget(widget, row, col + 1)
        layout.addWidget(summary_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        def _selected_index() -> int:
            current = table.currentItem()
            if current is None or current.row() >= len(items):
                return -1
            return current.row()

        def _render_items() -> None:
            _fill_table(
                table,
                [
                    [
                        {"material": "Material", "labor": "Mao de obra", "consumable": "Consumivel", "product": "Produto"}.get(str(item.get("kind", "")), "-"),
                        str(((item.get("line") or {}).get("descricao", "") or "-")).strip() or "-",
                        str(((item.get("line") or {}).get("produto_codigo", "") or (item.get("line") or {}).get("ref_externa", "") or "-")).strip() or "-",
                        f"{float(((item.get('line') or {}).get('qtd', 0) or 0)):.2f}",
                        str(((item.get("line") or {}).get("produto_unid", "") or "-")).strip() or "-",
                        f"{float(item.get('weight_total', 0) or 0):.2f} kg" if float(item.get("weight_total", 0) or 0) > 0 else "-",
                        _fmt_eur(float(((item.get("line") or {}).get("preco_unit", 0) or 0))),
                        _fmt_eur(float(item.get("total_cost", 0) or 0)),
                    ]
                    for item in items
                ],
                align_center_from=3,
            )
            totals = {"material": 0.0, "labor": 0.0, "consumable": 0.0, "product": 0.0}
            for item in items:
                kind = str(item.get("kind", "") or "")
                totals[kind] = totals.get(kind, 0.0) + float(item.get("total_cost", 0) or 0.0)
            subtotal = round(sum(totals.values()), 2)
            final_total = round(subtotal * (1.0 + (float(margin_spin.value() or 0) / 100.0)), 2)
            material_total_label.setText(_fmt_eur(totals.get("material", 0.0)))
            labor_total_label.setText(_fmt_eur(totals.get("labor", 0.0)))
            consumable_total_label.setText(_fmt_eur(totals.get("consumable", 0.0)))
            product_total_label.setText(_fmt_eur(totals.get("product", 0.0)))
            subtotal_label.setText(_fmt_eur(subtotal))
            final_total_label.setText(_fmt_eur(final_total))

        def _open_editor_for(kind: str, current: dict | None = None) -> dict | None:
            if kind == "material":
                return self._material_assembly_item_dialog(current, parent=dialog)
            if kind == "labor":
                return self._labor_assembly_item_dialog(current, parent=dialog)
            if kind == "consumable":
                return self._consumable_assembly_item_dialog(current, parent=dialog)
            if kind == "product":
                return self._product_assembly_item_dialog(current, parent=dialog)
            return None

        def _add_item(kind: str) -> None:
            try:
                payload = _open_editor_for(kind)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjunto", str(exc))
                return
            if payload is None:
                return
            items.append(payload)
            _render_items()

        def _line_kind_for_shortcut(line: dict) -> str:
            operation = str(line.get("operacao", "") or "").strip().lower()
            unit_txt = str(line.get("produto_unid", "") or "").strip().lower()
            if unit_txt == "h" and "corte laser" not in operation:
                return "labor"
            return self._assembly_item_kind_from_line(line)

        def _add_lines_from_shortcut(lines: list[dict], *, laser_batch_id: str = "") -> None:
            added = 0
            for row in list(lines or []):
                if not isinstance(row, dict) or not row:
                    continue
                line = dict(row)
                if laser_batch_id:
                    line["laser_source_mode"] = "batch"
                    line["laser_batch_id"] = laser_batch_id
                kind = _line_kind_for_shortcut(line)
                total_cost = round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2)
                items.append(
                    {
                        "kind": kind,
                        "quantity_units": round(float(line.get("qtd", 0) or 0), 3),
                        "total_cost": total_cost,
                        "weight_total": 0.0,
                        "line": line,
                    }
                )
                added += 1
            if added:
                _render_items()

        def _add_laser_batch_to_assembly() -> None:
            batch_dialog = LaserBatchQuoteDialog(
                self.backend,
                dialog,
                default_machine=self.workcenter_combo.currentText().strip(),
            )
            if batch_dialog.exec() != QDialog.Accepted:
                return
            result = dict(batch_dialog.result_payload() or {})
            _add_lines_from_shortcut(
                [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})],
                laser_batch_id=str(batch_dialog.batch_id or "").strip(),
            )

        def _add_step_igs_to_assembly() -> None:
            lines = self._open_profile_step_igs_quote_builder(return_lines=True, parent=dialog)
            _add_lines_from_shortcut([dict(row or {}) for row in list(lines or []) if dict(row or {})])

        def _add_structure_to_assembly() -> None:
            payload = self._structure_quote_dialog()
            if not payload:
                return
            _add_lines_from_shortcut([dict(row or {}) for row in list(payload.get("lines", []) or []) if dict(row or {})])

        def _edit_item() -> None:
            index = _selected_index()
            if index < 0:
                QMessageBox.warning(dialog, "Conjunto", "Seleciona um item.")
                return
            current = dict(items[index] or {})
            try:
                current_line = dict(current.get("line") or current)
                if self._quote_line_is_laser_2d(current_line):
                    batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
                    batch_indexes = [
                        row_index
                        for row_index, candidate in enumerate(items)
                        if batch_id
                        and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
                    ] or [index]
                    source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
                    edited_lines = self._edit_laser_batch_lines(source_lines, parent=dialog)
                    if edited_lines is None:
                        return
                    insert_at = min(batch_indexes)
                    for row_index in sorted(batch_indexes, reverse=True):
                        del items[row_index]
                    for offset, edited_line in enumerate(edited_lines):
                        items.insert(insert_at + offset, self._wrap_assembly_item(edited_line))
                    _render_items()
                    if edited_lines:
                        table.selectRow(insert_at)
                    return
                payload = _open_editor_for(str(current.get("kind", "") or ""), current)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjunto", str(exc))
                return
            if payload is None:
                return
            items[index] = payload
            _render_items()
            table.selectRow(index)

        def _remove_item() -> None:
            index = _selected_index()
            if index < 0:
                QMessageBox.warning(dialog, "Conjunto", "Seleciona um item.")
                return
            del items[index]
            _render_items()

        add_material_btn.clicked.connect(lambda: _add_item("material"))
        add_labor_btn.clicked.connect(lambda: _add_item("labor"))
        add_consumable_btn.clicked.connect(lambda: _add_item("consumable"))
        add_product_btn.clicked.connect(lambda: _add_item("product"))
        add_laser_batch_btn.clicked.connect(_add_laser_batch_to_assembly)
        add_step_igs_btn.clicked.connect(_add_step_igs_to_assembly)
        add_structure_btn.clicked.connect(_add_structure_to_assembly)
        edit_btn.clicked.connect(_edit_item)
        remove_btn.clicked.connect(_remove_item)
        margin_spin.valueChanged.connect(lambda _v: _render_items())
        _render_items()

        if dialog.exec() != QDialog.Accepted:
            return None
        if not items:
            QMessageBox.warning(self, "Conjunto", "O conjunto precisa de pelo menos um item.")
            return None
        assembly_code = code_edit.text().strip() or f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        assembly_name = name_edit.text().strip() or assembly_code
        lines = [dict(item.get("line") or {}) for item in items if isinstance(item.get("line"), dict)]
        for line in lines:
            operation_norm = self.backend.desktop_main.norm_text(str(line.get("operacao", "") or ""))
            if self.backend.desktop_main.orc_line_is_piece(line) and "laser" in operation_norm:
                line["source_quote_number"] = str(self.current_number or "").strip()
                line["source_ref_externa"] = str(line.get("ref_externa", "") or "").strip()
                line["pricing_source"] = "quote_laser"
        totals = {"material": 0.0, "labor": 0.0, "consumable": 0.0, "product": 0.0}
        for item in items:
            kind = str(item.get("kind", "") or "")
            totals[kind] = totals.get(kind, 0.0) + float(item.get("total_cost", 0) or 0.0)
        subtotal = round(sum(totals.values()), 2)
        final_total = round(subtotal * (1.0 + (float(margin_spin.value() or 0) / 100.0)), 2)
        notes_lines = [
            f"Conjunto: {assembly_name}",
            f"Materiais: {_fmt_eur(totals.get('material', 0.0))}",
            f"Mao de obra: {_fmt_eur(totals.get('labor', 0.0))}",
            f"Consumiveis: {_fmt_eur(totals.get('consumable', 0.0))}",
            f"Produtos: {_fmt_eur(totals.get('product', 0.0))}",
            f"Margem aplicada: {float(margin_spin.value() or 0):.2f}%",
            f"Preco final conjunto: {_fmt_eur(final_total)}",
        ]
        if notes_edit.toPlainText().strip():
            notes_lines.append(notes_edit.toPlainText().strip())
        technical_sheet = {
            "familia_produto": family_combo.currentText().strip(),
            "aplicacao": application_edit.text().strip(),
            "modelo_versao": model_edit.text().strip(),
            "configuracao": configuration_edit.text().strip(),
            "dimensoes_gerais": dimensions_edit.text().strip(),
            "materiais_acabamentos": finishes_edit.text().strip(),
            "caracteristicas": characteristics_edit.toPlainText().strip(),
            "requisitos_instalacao": installation_edit.toPlainText().strip(),
            "normas_conformidade": standards_edit.toPlainText().strip(),
            "controlo_qualidade": quality_edit.toPlainText().strip(),
        }
        try:
            self.backend.assembly_model_save(
                {
                    "codigo": assembly_code,
                    "param_codigo": param_edit.text().strip(),
                    "descricao": assembly_name,
                    "notas": "\n".join(notes_lines),
                    "itens": lines,
                    "template": bool(save_template_check.isChecked()),
                    "origem": "orcamento_conjunto_calculado",
                    "created_at": str(initial.get("created_at", "") or "").strip(),
                    "ficha_tecnica": technical_sheet,
                }
            )
            self.backend.conjunto_save(
                {
                    "codigo": assembly_code,
                    "param_codigo": param_edit.text().strip(),
                    "descricao": assembly_name,
                    "notas": "\n".join(notes_lines),
                    "itens": lines,
                    "template": bool(save_template_check.isChecked()),
                    "origem": "orcamento_conjunto_calculado",
                    "margem_perc": float(margin_spin.value() or 0.0),
                    "total_custo": subtotal,
                    "total_final": final_total,
                    "created_at": str(initial.get("created_at", "") or "").strip(),
                    "ficha_tecnica": technical_sheet,
                }
            )
        except Exception as exc:
            QMessageBox.critical(self, "Conjunto", str(exc))
            return None
        return {
            "assembly_code": assembly_code,
            "assembly_name": assembly_name,
            "name": assembly_name,
            "note_cliente": assembly_name,
            "notes_pdf": "\n".join(notes_lines),
            "summary_html": (
                f"{assembly_name} | materiais {_fmt_eur(totals.get('material', 0.0))} | "
                f"mao de obra {_fmt_eur(totals.get('labor', 0.0))} | "
                f"consumiveis {_fmt_eur(totals.get('consumable', 0.0))} | "
                f"produtos {_fmt_eur(totals.get('product', 0.0))} | final {_fmt_eur(final_total)}"
            ),
            "workcenter": self._quote_pick_workcenter("Serralharia", "Montagem"),
            "lines": lines,
        }

    def _open_calculated_assembly_builder(self) -> None:
        payload = self._calculated_assembly_builder_dialog()
        if not payload:
            return
        self._apply_group_payload(payload, replace_existing=False)

    def _open_structure_quote_builder(self) -> None:
        if self.line_rows:
            if QMessageBox.question(
                self,
                "Orcamento Estruturas",
                "Substituir as linhas atuais pelo modelo de estruturas metalicas?",
            ) != QMessageBox.Yes:
                return
            replace_existing = True
        else:
            replace_existing = True
        payload = self._structure_quote_dialog()
        if not payload:
            return
        self._apply_structure_quote_payload(payload, replace_existing=replace_existing)

    def _new_structure_quote(self) -> None:
        self._clear_quote_detail()
        payload = self._structure_quote_dialog()
        if not payload:
            self._show_list()
            return
        self._apply_structure_quote_payload(payload, replace_existing=True)

    def _transport_calc_value(self) -> float:
        kms = float(self.transport_km_spin.value() or 0)
        factor = float(self.transport_trip_factor_spin.value() or 1)
        rate = float(self.transport_rate_spin.value() or 0)
        diesel_price = float(self.transport_diesel_spin.value() or 0)
        consumption = float(self.transport_consumption_spin.value() or 0)
        route_cost = kms * factor * rate
        diesel_cost = kms * factor * (consumption / 100.0) * diesel_price
        return round(route_cost + diesel_cost, 2)

    def _recalc_transport_calc(self) -> None:
        self.transport_suggest_label.setText(f"Transporte sugerido: {_fmt_eur(self._transport_calc_value())}")

    def _apply_transport_calc(self) -> None:
        self.transport_price_spin.setValue(self._transport_calc_value())
        self._render_quote_lines()

    def _clear_transport(self) -> None:
        self.transport_combo.setCurrentText("")
        self.transport_carrier_combo.setCurrentText("")
        self.transport_zone_combo.setCurrentText("")
        self.transport_price_spin.setValue(0.0)
        self.transport_km_spin.setValue(0.0)
        self.transport_rate_spin.setValue(0.0)
        self.transport_diesel_spin.setValue(0.0)
        self.transport_consumption_spin.setValue(0.0)
        self.transport_trip_factor_spin.setValue(1.0)
        self._recalc_transport_calc()
        self._render_quote_lines()

    def _append_pdf_note(self, line: str) -> None:
        note = str(line or "").strip()
        if not note:
            return
        current = [item.strip() for item in self.notes_edit.toPlainText().splitlines() if item.strip()]
        if note not in current:
            current.append(note)
            self.notes_edit.setPlainText("\n".join(current))

    def _delivery_date_text(self) -> str:
        if self.delivery_date_edit.date() <= self.delivery_date_min:
            return ""
        return self.delivery_date_edit.date().toString("yyyy-MM-dd")

    def _delivery_text(self) -> str:
        return self.delivery_default_text

    def _apply_default_delivery_deadline(self) -> None:
        self.delivery_date_min = QDate.currentDate()
        self.delivery_date_edit.setMinimumDate(self.delivery_date_min)
        self.delivery_date_edit.setDate(self.delivery_date_min)
        self._append_pdf_note(f"- Prazo de entrega: {self.delivery_default_text}")

    def _fill_pdf_notes_from_context(self) -> None:
        try:
            text = self.backend.orc_suggest_notes(self._quote_payload())
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        lines = [item.strip() for item in str(text or "").splitlines() if item.strip()]
        current = [item.strip() for item in self.notes_edit.toPlainText().splitlines() if item.strip()]
        keep = [item for item in current if "foi considerado" not in item.lower()]
        merged = []
        for item in lines + keep:
            if item and item not in merged:
                merged.append(item)
        self.notes_edit.setPlainText("\n".join(merged))

    def _quote_payload(self) -> dict:
        self._render_quote_lines()
        return {
            "numero": self.current_number,
            "estado": self.state_chip.text().strip() or "Em edição",
            "cliente": {
                "codigo": self._client_code_from_text(self.client_combo.currentText()),
                "nome": self.client_name_edit.text().strip(),
                "empresa": self.client_company_edit.text().strip(),
                "nif": self.client_nif_edit.text().strip(),
                "morada": self.client_address_edit.text().strip(),
                "contacto": self.client_contact_edit.text().strip(),
                "email": self.client_email_edit.text().strip(),
            },
            "executado_por": self.executed_combo.currentText().strip(),
            "nota_transporte": self.transport_combo.currentText().strip(),
            "prazo_entrega_texto": self._delivery_text(),
            "prazo_entrega_data": self._delivery_date_text(),
            "transportadora_nome": self.transport_carrier_combo.currentText().strip(),
            "zona_transporte": self.transport_zone_combo.currentText().strip(),
            "preco_transporte": self.transport_price_spin.value(),
            "incremento_preco_perc": self.price_increment_spin.value(),
            "desconto_perc": self.discount_spin.value(),
            "desconto_modo": self._quote_discount_mode(),
            "desconto_grupos": self._quote_effective_discount_group_keys(),
            "notas_pdf": self.notes_edit.toPlainText().strip(),
            "nota_cliente": self.note_cliente_edit.text().strip(),
            "iva_perc": 23.0,
            "linhas": self.line_rows,
        }

    def _save_quote(self) -> None:
        self._set_quote_save_button_state("saving")
        QApplication.processEvents()
        try:
            detail = self.backend.orc_save(self._quote_payload())
        except Exception as exc:
            self._set_quote_save_button_state("error")
            QMessageBox.critical(self, "Guardar Orcamento", str(exc))
            self._quote_save_feedback_timer.start(2200)
            return
        self._load_quote(str(detail.get("numero", "") or "").strip())
        self.refresh()
        self._show_detail()
        self._set_quote_save_button_state("saved")
        self._quote_save_feedback_timer.start(1800)

    def _create_quote_purchase_note(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Nota de encomenda", "Guarda primeiro o orcamento.")
            return
        try:
            detail = self.backend.orc_save(self._quote_payload())
            result = self.backend.orc_create_purchase_quote(str(detail.get("numero", "") or self.current_number).strip())
        except Exception as exc:
            QMessageBox.critical(self, "Nota de encomenda", str(exc))
            return
        note_number = str(result.get("numero", "") or "").strip()
        line_count = int(result.get("line_count", 0) or 0)
        QMessageBox.information(self, "Nota de encomenda", f"Nota {note_number} criada com {line_count} linha(s) para pedido de cotacao.")
        main_window = self.window()
        if hasattr(main_window, "show_page"):
            try:
                main_window.show_page("purchase_notes")
                page = getattr(main_window, "pages", {}).get("purchase_notes")
                if page is not None and hasattr(page, "refresh"):
                    page.refresh()
                if page is not None and hasattr(page, "_load_note"):
                    page._load_note(note_number)
            except Exception:
                pass

    def _set_quote_save_button_state(self, state: str = "idle") -> None:
        buttons = [
            button
            for button in (
                getattr(self, "quote_save_btn", None),
                getattr(self, "quote_inspector_save_btn", None),
            )
            if isinstance(button, QPushButton)
        ]
        if not buttons:
            return
        current_state = str(state or "idle").strip().lower()
        for button in buttons:
            if current_state == "saving":
                button.setText("A guardar...")
                button.setEnabled(False)
                button.setProperty("variant", "secondary")
            elif current_state == "saved":
                button.setText("Guardado")
                button.setEnabled(True)
                button.setProperty("variant", "success")
            elif current_state == "error":
                button.setText("Falhou")
                button.setEnabled(True)
                button.setProperty("variant", "danger")
            else:
                button.setText("Guardar")
                button.setEnabled(True)
                button.setProperty("variant", "warning")
            _repolish(button)

    def _reset_quote_save_button(self) -> None:
        self._set_quote_save_button_state("idle")

    def _quote_email_subject(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.note_cliente_edit.text().strip()).strip()
        client = dict(payload.get("cliente", {}) or {})
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        subject = f"Proposta {client_name} [{numero}]" if numero else "Proposta luGEST"
        if rfq_ref:
            subject += f" | Pedido de Cotação {rfq_ref}"
        return subject

    def _quote_email_body(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.note_cliente_edit.text().strip()).strip()
        lines = [
            f"Exmos. Senhores {client_name},",
            "",
            f"Segue em anexo o orçamento {numero}." if numero else "Segue em anexo o orçamento solicitado.",
        ]
        if rfq_ref:
            lines.append(f"Referência do pedido de cotação: {rfq_ref}.")
        lines.extend(
            [
                "",
                "Ficamos ao dispor para qualquer esclarecimento.",
                "",
                "Cumprimentos,",
                "luGEST",
            ]
        )
        return "\n".join(lines)

    def _quote_email_html_body(self, detail: dict[str, object] | None = None, *, logo_cid: str = "") -> str:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        branding = dict(getattr(self.backend, "branding", {}) or {})
        company_name = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.note_cliente_edit.text().strip()).strip()
        subtotal = _fmt_eur(float(payload.get("subtotal", 0) or 0))
        client_code = str(client.get("codigo", "") or "").strip()
        payment_terms = "Conforme acordado"
        try:
            client_rows = list(self.backend.client_rows("") or [])
            match = next(
                (
                    row
                    for row in client_rows
                    if (
                        client_code
                        and str(row.get("codigo", "") or "").strip() == client_code
                    )
                    or (
                        not client_code
                        and str(row.get("email", "") or "").strip().lower() == str(client.get("email", "") or "").strip().lower()
                    )
                ),
                None,
            )
            if isinstance(match, dict):
                payment_terms = str(match.get("cond_pagamento", "") or "").strip() or payment_terms
        except Exception:
            pass
        header_logo = ""
        if logo_cid:
            header_logo = (
                f"<img src=\"cid:{html.escape(logo_cid)}\" alt=\"{html.escape(company_name)}\" "
                "style=\"max-width:110px; max-height:36px; display:block;\" />"
            )
        else:
            header_logo = (
                f"<div style=\"font-size:34px; font-weight:800; letter-spacing:-1px; color:#ffffff;\">"
                f"{html.escape(company_name)}</div>"
            )
        reference_block = html.escape(rfq_ref or numero or "-")
        return (
            "<html><body style=\"margin:0; padding:0; background:#eef2f7; font-family:Segoe UI, Arial, sans-serif; color:#334155;\">"
            "<div style=\"padding:28px 0;\">"
            "<div style=\"width:620px; margin:0 auto; background:#ffffff; border-radius:18px; overflow:hidden; box-shadow:0 16px 40px rgba(15,23,42,0.10);\">"
            "<div style=\"background:#1f2933; padding:28px 34px;\">"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;\">"
            "<tr>"
            f"<td style=\"vertical-align:middle;\">{header_logo}</td>"
            "<td style=\"vertical-align:top; text-align:right; color:#cbd5e1; font-size:11px; letter-spacing:0.6px;\">"
            "REFERÊNCIA<br>"
            f"<span style=\"display:inline-block; margin-top:8px; font-size:23px; font-weight:800; color:#ffffff;\">{html.escape(reference_block)}</span>"
            "</td>"
            "</tr>"
            "</table>"
            "</div>"
            "<div style=\"padding:34px 36px 26px 36px;\">"
            f"<p style=\"margin:0 0 18px 0; font-size:22px; font-weight:800; color:#0f172a;\">Exmo(a). {html.escape(client_name)},</p>"
            "<p style=\"margin:0 0 16px 0; font-size:16px; line-height:1.7;\">Boa tarde,</p>"
            "<p style=\"margin:0 0 24px 0; font-size:16px; line-height:1.7;\">"
            "Conforme solicitado, segue o orçamento em anexo."
            "</p>"
            "<div style=\"margin:28px 0 12px 0; font-size:13px; font-weight:800; color:#94a3b8; letter-spacing:0.6px; text-transform:uppercase;\">"
            "Condições da proposta"
            "</div>"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse; border:1px solid #e2e8f0;\">"
            "<tr style=\"background:#1f2933; color:#ffffff; font-size:12px; font-weight:800; text-transform:uppercase;\">"
            "<td style=\"padding:12px 16px;\">Campo</td>"
            "<td style=\"padding:12px 16px;\">Valor</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Valor Total (s/ IVA)</td>"
            f"<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:30px; font-weight:800; color:#16a34a;\">{html.escape(subtotal)}</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Validade da proposta</td>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#0f172a;\">5 dias úteis, salvo rutura de stock.</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Condições de pagamento</td>"
            f"<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#0f172a;\">{html.escape(payment_terms)}</td>"
            "</tr>"
            "</table>"
            "<p style=\"margin:28px 0 0 0; font-size:15px; line-height:1.7;\">"
            "Ficamos ao dispor para qualquer esclarecimento."
            "</p>"
            "</div>"
            f"<div style=\"padding:16px 36px; background:#f8fafc; border-top:1px solid #e2e8f0; font-size:12px; color:#94a3b8; text-align:center;\">© {datetime.now().year} {html.escape(company_name)}</div>"
            "</div>"
            "</div>"
            "</body></html>"
        )

    def _open_quote_email_draft(self, detail: dict[str, object] | None = None) -> None:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        recipient = str(client.get("email", "") or self.client_email_edit.text().strip()).strip()
        if not recipient:
            QMessageBox.warning(self, "Orçamentos", "O cliente não tem email definido para preparar o envio.")
            return

        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        if not numero:
            raise ValueError("Guarda primeiro o orçamento antes de preparar o email.")

        rfq_ref = str(payload.get("nota_cliente", "") or self.note_cliente_edit.text().strip()).strip()
        safe_ref = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in (rfq_ref or numero))[:48]
        attachment_name = f"Proposta_{safe_ref}_{numero}.pdf" if numero else f"Proposta_{safe_ref}.pdf"
        attachment_path: Path | None = Path(tempfile.gettempdir()) / attachment_name
        attachment_issue = ""
        try:
            self.backend.orc_render_pdf(numero, attachment_path)
        except Exception as exc:
            attachment_issue = str(exc)
            attachment_path = None

        subject = self._quote_email_subject(payload)
        body_plain = self._quote_email_body(payload)
        logo_path = getattr(self.backend, "logo_path", None)
        logo_file = Path(logo_path) if isinstance(logo_path, Path) and logo_path.exists() else None
        logo_cid = "lugest-mail-logo" if logo_file is not None else ""
        body_html = self._quote_email_html_body(payload, logo_cid=logo_cid)

        env = os.environ.copy()
        env["LUGEST_MAIL_TO"] = recipient
        env["LUGEST_MAIL_SUBJECT"] = subject
        env["LUGEST_MAIL_BODY"] = body_html
        env["LUGEST_MAIL_ATTACHMENT"] = str(attachment_path) if attachment_path is not None else ""
        env["LUGEST_MAIL_LOGO"] = str(logo_file) if logo_file is not None else ""
        env["LUGEST_MAIL_LOGO_CID"] = logo_cid
        powershell_script = (
            "$ErrorActionPreference='Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            "$mail.To = $env:LUGEST_MAIL_TO; "
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
            "$mail.Display()"
        )
        def on_email_ready(ok: bool, _error: str) -> None:
            if not ok:
                mailto = f"mailto:{quote(recipient)}?subject={quote(subject)}&body={quote(body_plain)}"
                try:
                    os.startfile(mailto)
                except Exception as exc:
                    QMessageBox.warning(self, "Orçamentos", f"Não foi possível abrir o cliente de email:\n{exc}")
                    return
                fallback_message = "Outlook indisponível. Foi aberto o cliente de email por defeito."
                if attachment_issue:
                    fallback_message += f"\n\nTambém não foi possível gerar o PDF em anexo:\n{attachment_issue}"
                else:
                    fallback_message += "\n\nNota: o anexo PDF terá de ser adicionado manualmente neste modo."
                QMessageBox.information(self, "Orçamentos", fallback_message)
                return
            if attachment_issue:
                QMessageBox.information(
                    self,
                    "Orçamentos",
                    f"Rascunho aberto no Outlook, mas o PDF não foi anexado automaticamente:\n{attachment_issue}",
                )

        _run_process_async(
            self,
            "powershell",
            ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
            environment=env,
            timeout_ms=30000,
            finished=on_email_ready,
        )

    def _set_quote_state(self, estado: str) -> None:
        try:
            self.backend.orc_save(self._quote_payload())
            detail = self.backend.orc_set_state(self.current_number, estado)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._load_quote(str(detail.get("numero", "") or "").strip())
        self.refresh()
        self._show_detail()
        if str(estado or "").strip().lower() == "enviado":
            try:
                self._open_quote_email_draft(detail)
            except Exception as exc:
                QMessageBox.warning(
                    self,
                    "Orçamentos",
                    "O orçamento foi marcado como Enviado, mas não foi possível abrir o email:\n"
                    f"{exc}",
                )

    def _remove_quote(self) -> None:
        row = self._selected_quote_row()
        numero = str((row or {}).get("numero", "") or self.current_number).strip()
        if not numero:
            QMessageBox.warning(self, "Orçamentos", "Seleciona um orcamento.")
            return
        if QMessageBox.question(self, "Apagar orcamento", f"Remover orcamento {numero}?") != QMessageBox.Yes:
            return
        try:
            self.backend.orc_remove(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._show_list()
        self.refresh()

    def _line_dialog(self, initial: dict | None = None, *, template_mode: bool = False) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Item do conjunto" if template_mode else "Linha de orcamento")
        dialog.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setMinimumSize(860, 560)
        try:
            screen_geo = QApplication.primaryScreen().availableGeometry()
            dialog.resize(min(980, max(860, screen_geo.width() - 120)), min(760, max(560, screen_geo.height() - 110)))
        except Exception:
            dialog.resize(960, 720)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setContentsMargins(0, 0, 0, 0)
        form.setVerticalSpacing(7)
        client_code = self._client_code_from_text(self.client_combo.currentText())
        references = self.backend.order_reference_rows("", client_code)
        refs_by_ext = {str(row.get("ref_externa", "") or "").strip(): row for row in references if str(row.get("ref_externa", "") or "").strip()}
        refs_by_int = {str(row.get("ref_interna", "") or "").strip(): row for row in references if str(row.get("ref_interna", "") or "").strip()}
        presets = self.presets or self.backend.order_presets()
        product_rows = {str(row.get("codigo", "") or "").strip(): row for row in self.backend.ne_product_options("")}
        stock_material_rows: list[tuple[dict[str, Any], dict[str, Any]]] = []

        def reload_stock_material_rows() -> None:
            stock_material_rows.clear()
            for item in list(self.backend.material_rows("") or []):
                if not isinstance(item, dict) or not isinstance(item.get("record"), dict):
                    continue
                record = dict(item.get("record") or {})
                preview = dict(item.get("preview") or self.backend.material_price_preview(record) or {})
                stock_material_rows.append((record, preview))

        reload_stock_material_rows()
        initial_type = str(
            self.backend.desktop_main.normalize_orc_line_type(
                initial.get("tipo_item", self.backend.desktop_main.ORC_LINE_TYPE_PIECE)
            )
        )
        if (
            (
                str(initial.get("stock_material_id", "") or "").strip()
                or str(initial.get("stock_item_kind", "") or "").strip() == "raw_material"
            )
            and initial_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE
        ):
            initial_type = "stock_mp"
        ref_history = QComboBox()
        ref_history.setEditable(True)
        ref_history.addItem("")
        for row in references:
            ref_history.addItem(
                f"{row.get('ref_interna', '')} | {row.get('ref_externa', '')} | {row.get('descricao', '')}",
                row,
            )
        type_combo = QComboBox()
        type_combo.addItem("Componente fabricado", self.backend.desktop_main.ORC_LINE_TYPE_PIECE)
        type_combo.addItem("Matéria Prima - Stock", "stock_mp")
        type_combo.addItem("Produto stock", self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT)
        type_combo.addItem("Servico montagem", self.backend.desktop_main.ORC_LINE_TYPE_SERVICE)
        for index in range(type_combo.count()):
            if str(type_combo.itemData(index) or "") == initial_type:
                type_combo.setCurrentIndex(index)
                break
        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.addItem("")
        for code, row in sorted(product_rows.items()):
            product_combo.addItem(f"{code} - {row.get('descricao', '')}", code)
        stock_material_combo = QComboBox()
        stock_material_combo.setEditable(True)
        stock_kind_combo = QComboBox()
        stock_kind_combo.addItem("Chapa", "Chapa")
        stock_kind_combo.addItem("Viga / Perfil", "Perfil")
        stock_kind_combo.addItem("Tubo", "Tubo")
        stock_kind_combo.addItem("Cantoneira", "Cantoneira")
        stock_kind_combo.addItem("Barra", "Barra")
        stock_kind_combo.addItem("Ferro nervurado", "Ferro nervurado")
        stock_kind_combo.addItem("Todos", "")

        def _stock_kind(record: dict[str, Any], preview: dict[str, Any]) -> str:
            formato = str(preview.get("formato", record.get("formato", "")) or "").strip().title()
            secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
            material_norm = self.backend.desktop_main.norm_text(str(record.get("material", "") or ""))
            if formato == "Perfil" and secao == "L":
                return "Cantoneira"
            if formato == "Chapa" and any(token in material_norm for token in ("barra", "chata", "plat", "flat")):
                return "Barra"
            if formato in {"Chapa", "Tubo", "Perfil", "Cantoneira", "Barra", "Ferro Nervurado"}:
                return "Ferro nervurado" if formato == "Ferro Nervurado" else formato
            detected = str(self.backend.desktop_main.detect_materia_formato(record) or "").strip().title()
            if detected == "Ferro Nervurado":
                return "Ferro nervurado"
            return detected

        def _stock_label(record: dict[str, Any], preview: dict[str, Any]) -> str:
            material_txt = str(record.get("material", "") or "-").strip()
            dim_txt = str(preview.get("dimension_label", "") or "-").strip()
            esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
            qty_txt = self.backend._fmt(record.get("quantidade", 0))
            price_txt = _fmt_eur(float(preview.get("preco_unid", record.get("preco_unid", 0)) or 0))
            parts = [str(record.get("id", "") or "").strip(), material_txt, dim_txt]
            if esp_txt:
                parts.append(f"{esp_txt} mm")
            if qty_txt:
                parts.append(f"Qtd {qty_txt}")
            if price_txt:
                parts.append(f"Base {price_txt}")
            return " | ".join(part for part in parts if part and part != "-")

        def refresh_stock_material_combo(preferred_id: str = "") -> None:
            current_id = str(preferred_id or "").strip()
            if not current_id:
                payload = stock_material_combo.currentData()
                if isinstance(payload, dict):
                    current_id = str(payload.get("id", "") or "").strip()
            wanted_kind = str(stock_kind_combo.currentData() or "").strip()
            stock_material_combo.blockSignals(True)
            stock_material_combo.clear()
            stock_material_combo.addItem("", None)
            for record, preview in stock_material_rows:
                if wanted_kind and _stock_kind(record, preview) != wanted_kind:
                    continue
                stock_material_combo.addItem(_stock_label(record, preview), dict(record))
            if current_id:
                for index in range(stock_material_combo.count()):
                    payload = stock_material_combo.itemData(index)
                    if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == current_id:
                        stock_material_combo.setCurrentIndex(index)
                        break
            stock_material_combo.blockSignals(False)
        initial_product_code = str(initial.get("produto_codigo", "") or "").strip()
        if initial_product_code:
            for index in range(product_combo.count()):
                if str(product_combo.itemData(index) or "").strip() == initial_product_code:
                    product_combo.setCurrentIndex(index)
                    break
            else:
                product_combo.setCurrentText(initial_product_code)
        ref_int_edit = QLineEdit("" if template_mode else str(initial.get("ref_interna", "") or "").strip())
        ref_ext_edit = QLineEdit(str(initial.get("ref_externa", "") or "").strip())
        desc_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
        dimension_edit = QLineEdit(str(initial.get("dimensao", initial.get("dimensoes", "")) or "").strip())
        product_unid_edit = QLineEdit(str(initial.get("produto_unid", "") or "").strip())
        product_unid_edit.setReadOnly(True)
        mp_metric_spin = QDoubleSpinBox()
        mp_metric_spin.setRange(0.0, 1000000.0)
        mp_metric_spin.setDecimals(4)
        mp_metric_spin.setSuffix(" kg")
        mp_metric_spin.setReadOnly(True)
        mp_metric_spin.setButtonSymbols(QAbstractSpinBox.NoButtons)
        mp_price_kg_spin = QDoubleSpinBox()
        mp_price_kg_spin.setRange(0.0, 1000000.0)
        mp_price_kg_spin.setDecimals(4)
        mp_price_kg_spin.setPrefix("EUR ")
        mp_price_kg_spin.setSuffix("/kg")
        mp_margin_spin = QDoubleSpinBox()
        mp_margin_spin.setRange(-100.0, 1000000.0)
        mp_margin_spin.setDecimals(2)
        mp_margin_spin.setSuffix(" %")
        mp_margin_spin.setValue(float(initial.get("price_markup_pct", initial.get("markup_pct", 0)) or 0.0))
        mp_price_info = QLabel("")
        mp_price_info.setWordWrap(True)
        mp_price_info.setProperty("role", "muted")
        mp_price_table_btn = QPushButton("Tabela preços matéria-prima")
        mp_price_table_btn.setProperty("variant", "secondary")
        stock_create_btn = QPushButton("Criar material")
        stock_create_btn.setProperty("variant", "secondary")
        stock_tools_layout = QHBoxLayout()
        stock_tools_layout.setContentsMargins(0, 0, 0, 0)
        stock_tools_layout.setSpacing(8)
        stock_tools_layout.addWidget(stock_create_btn)
        stock_tools_layout.addWidget(mp_price_table_btn)
        stock_tools_layout.addStretch(1)
        stock_tools_host = QWidget()
        stock_tools_host.setLayout(stock_tools_layout)
        stock_price_state = {
            "base_label": "EUR/kg",
            "metric_label": "Kg por unid.",
            "metric_suffix": " kg",
            "metric_value": 0.0,
            "base_value": 0.0,
            "price_kg_equiv": 0.0,
            "busy": False,
            "price_user_touched": False,
        }
        stock_line_meta = {
            "stock_material_id": str(initial.get("stock_material_id", "") or "").strip(),
            "material_family": str(initial.get("material_family", "") or "").strip(),
            "material_subtype": str(initial.get("material_subtype", "") or "").strip(),
        }
        if stock_line_meta["stock_material_id"]:
            for record, preview in stock_material_rows:
                if str(record.get("id", "") or "").strip() == stock_line_meta["stock_material_id"]:
                    kind = _stock_kind(record, preview)
                    for index in range(stock_kind_combo.count()):
                        if str(stock_kind_combo.itemData(index) or "").strip() == kind:
                            stock_kind_combo.setCurrentIndex(index)
                            break
                    break
        refresh_stock_material_combo(stock_line_meta["stock_material_id"])
        material_combo = QComboBox()
        material_combo.setEditable(True)
        for value in list(presets.get("materiais", []) or []):
            material_combo.addItem(str(value))
        material_combo.setCurrentText(str(initial.get("material", "") or "").strip())
        esp_combo = QComboBox()
        esp_combo.setEditable(True)
        for value in list(presets.get("espessuras", []) or []):
            esp_combo.addItem(str(value))
        esp_combo.setCurrentText(str(initial.get("espessura", "") or "").strip())
        def _normalize_ops_from_any(value: Any) -> list[str]:
            items = self.backend.quote_parse_operacoes_lista(value)
            ordered: list[str] = []
            for raw_name in list(items or []):
                normalized = str(self.backend.desktop_main.normalize_operacao_nome(raw_name) or raw_name or "").strip()
                if normalized and normalized not in ordered:
                    ordered.append(normalized)
            return ordered

        def _display_operation_text(value: Any, *, has_laser_base: bool) -> str:
            ops = _normalize_ops_from_any(value)
            if has_laser_base:
                ops = [op_name for op_name in ops if op_name != "Corte Laser"]
            return " + ".join(ops)

        operation_change_state = {"user_initiated": False, "last_prompt_signature": ""}
        operation_selector, operation_edit, apply_operations = _build_operation_selector(
            list(presets.get("operacoes", []) or []),
            "",
            on_change=lambda _text, user_initiated: operation_change_state.__setitem__("user_initiated", bool(user_initiated)),
        )
        tempo_spin = QDoubleSpinBox()
        tempo_spin.setRange(0.0, 1000000.0)
        tempo_spin.setDecimals(2)
        tempo_spin.setValue(float(initial.get("tempo_peca_min", initial.get("tempo_pecas_min", 0)) or 0))
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
        price_spin = QDoubleSpinBox()
        price_spin.setRange(0.0, 1000000.0)
        price_spin.setDecimals(4)
        price_spin.setValue(float(initial.get("preco_unit", 0) or 0))
        drawing_edit = QLineEdit(str(initial.get("desenho", "") or "").strip())
        pdf_docs: list[str] = []
        for raw_doc in [initial.get("desenho_pdf", ""), *list(initial.get("desenhos_pdf", []) or [])]:
            doc_txt = str(raw_doc or "").strip()
            if doc_txt and doc_txt.lower().endswith(".pdf") and doc_txt not in pdf_docs:
                pdf_docs.append(doc_txt)
        for raw_doc in list(initial.get("ficheiros", []) or []):
            doc_txt = str(raw_doc or "").strip()
            if doc_txt and doc_txt.lower().endswith(".pdf") and doc_txt not in pdf_docs:
                pdf_docs.append(doc_txt)
        initial_ref_int = str(initial.get("ref_interna", "") or "").strip()
        operation_meta = {
            "operacoes_lista": list(initial.get("operacoes_lista", []) or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(initial.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(initial.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(initial.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(initial.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(initial.get("quote_cost_snapshot", {}) or {}),
        }
        operation_cost_label = QLabel("")
        operation_cost_label.setWordWrap(True)
        operation_cost_label.setProperty("role", "muted")

        def current_type_token() -> str:
            return str(type_combo.currentData() or self.backend.desktop_main.ORC_LINE_TYPE_PIECE)

        def current_line_type() -> str:
            token = current_type_token()
            if token == "stock_mp":
                return self.backend.desktop_main.ORC_LINE_TYPE_PIECE
            return token

        def _match_stock_material_payload() -> dict[str, Any] | None:
            payload = stock_material_combo.currentData()
            if isinstance(payload, dict):
                return dict(payload)
            probe = stock_material_combo.currentText().strip().lower()
            if not probe:
                return None
            for index in range(1, stock_material_combo.count()):
                candidate = stock_material_combo.itemData(index)
                if not isinstance(candidate, dict):
                    continue
                label = stock_material_combo.itemText(index).strip().lower()
                material_id = str(candidate.get("id", "") or "").strip().lower()
                if probe == label or (material_id and (probe == material_id or probe.startswith(material_id))):
                    stock_material_combo.setCurrentIndex(index)
                    return dict(candidate)
            return None

        def stock_material_detail() -> tuple[dict[str, Any], dict[str, Any]]:
            payload = _match_stock_material_payload()
            if not isinstance(payload, dict):
                return {}, {}
            preview = dict(self.backend.material_price_preview(payload) or {})
            return dict(payload), preview

        def stock_pricing_context(record: dict[str, Any] | None = None, preview: dict[str, Any] | None = None) -> dict[str, Any]:
            record = dict(record or {})
            preview = dict(preview or {})
            formato = str(preview.get("formato", record.get("formato", "")) or "").strip() or "Chapa"
            base_label = str(preview.get("base_label", "EUR/m" if formato == "Tubo" else "EUR/kg") or "EUR/kg").strip()
            metric_label = "Metros por unid." if base_label == "EUR/m" else "Kg por unid."
            metric_suffix = " m" if base_label == "EUR/m" else " kg"
            metric_value = float(
                (
                    preview.get("metros", record.get("metros", 0))
                    if base_label == "EUR/m"
                    else preview.get("peso_unid", record.get("peso_unid", 0))
                )
                or 0.0
            )
            kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
            base_value = float(record.get("p_compra", 0) or 0.0)
            price_kg_equiv = round((base_value / kg_m), 4) if base_label == "EUR/m" and kg_m > 0 else round(base_value, 4)
            return {
                "formato": formato,
                "base_label": base_label,
                "metric_label": metric_label,
                "metric_suffix": metric_suffix,
                "metric_value": round(metric_value, 4),
                "base_value": round(base_value, 4),
                "price_kg_equiv": price_kg_equiv,
                "kg_m": round(kg_m, 4),
            }

        def set_form_row_visible(field: QWidget, visible: bool, label_text: str | None = None) -> None:
            field.setVisible(visible)
            label_widget = form.labelForField(field)
            if label_widget is not None:
                label_widget.setVisible(visible)
                if label_text is not None:
                    label_widget.setText(label_text)

        def current_piece_metric_kg() -> float:
            weight_total = float(initial.get("weight_total", 0) or 0.0)
            quantity_units = float(initial.get("quantity_units", initial.get("qtd", 1)) or 1.0)
            if weight_total > 0 and quantity_units > 0:
                return round(weight_total / quantity_units, 4)
            kg_m = float(initial.get("kg_per_m", 0) or 0.0)
            meters = float(initial.get("meters_per_unit", 0) or 0.0)
            if kg_m > 0 and meters > 0:
                return round(kg_m * meters, 4)
            stock_record = None
            stock_preview = None
            if str(stock_line_meta.get("stock_material_id", "") or "").strip():
                stock_record = self.backend.material_by_id(str(stock_line_meta.get("stock_material_id", "") or "").strip())
                stock_preview = dict(self.backend.material_price_preview(stock_record) or {}) if isinstance(stock_record, dict) else {}
            if isinstance(stock_preview, dict):
                return round(float(stock_preview.get("peso_unid", stock_record.get("peso_unid", 0) if isinstance(stock_record, dict) else 0) or 0.0), 4)
            return 0.0

        def _payload_has_laser_base(payload_row: dict[str, Any] | None) -> bool:
            source = dict(payload_row or {})
            if bool(source.get("laser_base_active", False)):
                return True
            if current_line_type() != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
                return False
            if not str(source.get("desenho", "") or "").strip():
                return False
            if self.backend._parse_float(source.get("tempo_peca_min", source.get("tempo_pecas_min", 0)), 0) <= 0 and self.backend._parse_float(source.get("preco_unit", 0), 0) <= 0:
                return False
            return "Corte Laser" in _normalize_ops_from_any(source.get("operacao", source.get("operacoes", source.get("operacoes_lista", []))))

        base_state = {
            "laser_base_enabled": _payload_has_laser_base(initial),
            "laser_base_time": round(float(initial.get("laser_base_tempo_unit", initial.get("tempo_peca_min", initial.get("tempo_pecas_min", 0))) or 0), 4),
            "laser_base_price": round(float(initial.get("laser_base_preco_unit", initial.get("preco_unit", 0)) or 0), 4),
        }
        apply_operations(_display_operation_text(initial.get("operacao", initial.get("operacoes_lista", [])), has_laser_base=bool(base_state["laser_base_enabled"])))

        def current_line_refs() -> list[str]:
            refs: list[str] = []
            skipped_current = False
            for row in list(self.line_rows):
                ref_txt = str((row or {}).get("ref_interna", "") or "").strip()
                if not ref_txt:
                    continue
                if initial_ref_int and not skipped_current and ref_txt == initial_ref_int:
                    skipped_current = True
                    continue
                refs.append(ref_txt)
            return refs

        def current_product_code() -> str:
            raw_code = str(product_combo.currentData() or "").strip()
            if raw_code:
                return raw_code
            raw_text = product_combo.currentText().strip()
            candidate = raw_text.split(" - ", 1)[0].strip()
            return candidate if candidate in product_rows else ""

        def apply_reference(payload: dict | None) -> None:
            if not isinstance(payload, dict):
                return
            ref_ext_edit.setText(str(payload.get("ref_externa", "") or "").strip())
            if not template_mode:
                ref_int_edit.setText(str(payload.get("ref_interna", "") or "").strip())
            desc_edit.setText(str(payload.get("descricao", "") or "").strip())
            material_combo.setCurrentText(str(payload.get("material", "") or "").strip())
            esp_combo.setCurrentText(str(payload.get("espessura", "") or "").strip())
            base_state["laser_base_enabled"] = _payload_has_laser_base(payload)
            base_state["laser_base_time"] = round(
                float(payload.get("laser_base_tempo_unit", payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0))) or 0),
                4,
            )
            base_state["laser_base_price"] = round(float(payload.get("laser_base_preco_unit", payload.get("preco_unit", payload.get("preco", 0))) or 0), 4)
            operation_meta["operacoes_lista"] = list(payload.get("operacoes_lista", []) or [])
            operation_meta["operacoes_fluxo"] = [dict(item or {}) for item in list(payload.get("operacoes_fluxo", []) or []) if isinstance(item, dict)]
            operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
            operation_meta["tempos_operacao"] = dict(payload.get("tempos_operacao", {}) or {})
            operation_meta["custos_operacao"] = dict(payload.get("custos_operacao", {}) or {})
            operation_meta["quote_cost_snapshot"] = dict(payload.get("quote_cost_snapshot", {}) or {})
            apply_operations(_display_operation_text(payload.get("operacoes", payload.get("operacao", payload.get("operacoes_lista", []))), has_laser_base=bool(base_state["laser_base_enabled"])))
            tempo_spin.setValue(float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)) or 0))
            price_spin.setValue(float(payload.get("preco_unit", payload.get("preco", 0)) or 0))
            if not drawing_edit.text().strip():
                drawing_edit.setText(str(payload.get("desenho", "") or "").strip())

        def sync_product_fields() -> None:
            code = current_product_code()
            row = product_rows.get(code)
            if row is None:
                return
            product_unid_edit.setText(str(row.get("unid", "") or "UN").strip())
            if not desc_edit.text().strip() or current_line_type() == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
                desc_edit.setText(str(row.get("descricao", "") or "").strip())
            sale_price = float(row.get("preco_venda", row.get("pvp1", row.get("preco", 0))) or 0)
            if current_line_type() == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT or float(price_spin.value() or 0) <= 0:
                price_spin.setValue(sale_price)

        def sync_line_price_from_mp() -> None:
            stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
            if not stock_id:
                return
            if stock_price_state["busy"]:
                return
            metric_value = float(mp_metric_spin.value() or stock_price_state.get("metric_value", 0.0) or 0.0)
            base_value = float(mp_price_kg_spin.value() or 0.0)
            if metric_value <= 0 or base_value <= 0:
                return
            sale_unit = round(metric_value * base_value * (1.0 + (float(mp_margin_spin.value() or 0.0) / 100.0)), 4)
            stock_price_state["busy"] = True
            try:
                price_spin.setValue(sale_unit)
            finally:
                stock_price_state["busy"] = False

        def sync_margin_from_line_price() -> None:
            stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
            if not stock_id or stock_price_state["busy"]:
                return
            metric_value = float(mp_metric_spin.value() or stock_price_state.get("metric_value", 0.0) or 0.0)
            base_value = float(mp_price_kg_spin.value() or 0.0)
            cost_unit = metric_value * base_value
            if cost_unit <= 0:
                return
            stock_price_state["busy"] = True
            try:
                mp_margin_spin.setValue(round(((float(price_spin.value() or 0.0) / cost_unit) - 1.0) * 100.0, 2))
            finally:
                stock_price_state["busy"] = False

        def sync_stock_material_fields() -> None:
            record, preview = stock_material_detail()
            if not record:
                stock_line_meta["stock_material_id"] = ""
                mp_metric_spin.setValue(0.0)
                mp_price_kg_spin.setValue(0.0)
                mp_margin_spin.setValue(0.0)
                mp_price_info.setText("")
                return
            previous_stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
            current_stock_id = str(record.get("id", "") or "").strip()
            if current_stock_id and current_stock_id != previous_stock_id:
                stock_price_state["price_user_touched"] = False
            desc_txt = str(record.get("material", "") or "").strip()
            dim_txt = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
            esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
            if not esp_txt:
                esp_txt = dim_txt
            if dim_txt:
                desc_txt = f"{desc_txt} {dim_txt}".strip()
                dimension_edit.setText(dim_txt)
            if esp_txt and esp_txt not in desc_txt:
                desc_txt = f"{desc_txt} {esp_txt} mm".strip()
            material_combo.setCurrentText(str(record.get("material", "") or "").strip())
            if esp_txt:
                esp_combo.setCurrentText(esp_txt)
            if not desc_edit.text().strip() or current_type_token() == "stock_mp":
                desc_edit.setText(desc_txt)
            if not ref_ext_edit.text().strip() or current_type_token() == "stock_mp":
                ref_ext_edit.setText(str(record.get("id", "") or "").strip() or desc_txt)
            stock_line_meta["stock_material_id"] = current_stock_id
            stock_line_meta["material_family"] = str(record.get("material", "") or "").strip()
            stock_line_meta["material_subtype"] = str(preview.get("formato", record.get("formato", "")) or "Stock MP").strip()
            if float(qtd_spin.value() or 0) <= 0:
                qtd_spin.setValue(1.0)
            pricing = stock_pricing_context(record, preview)
            stock_price_state.update(pricing)
            metric_label = form.labelForField(mp_metric_spin)
            if metric_label is not None:
                metric_label.setText(pricing["metric_label"])
            price_label = form.labelForField(mp_price_kg_spin)
            if price_label is not None:
                price_label.setText(f"Preco compra ({pricing['base_label']})")
            mp_metric_spin.setSuffix(pricing["metric_suffix"])
            mp_metric_spin.setValue(float(pricing["metric_value"] or 0.0))
            mp_price_kg_spin.setSuffix("/m" if pricing["base_label"] == "EUR/m" else "/kg")
            if not bool(stock_price_state.get("price_user_touched", False)):
                mp_price_kg_spin.setValue(float(pricing["base_value"] or 0.0))
            if float(initial.get("preco_unit", 0) or 0.0) > 0 and float(qtd_spin.value() or 0.0) > 0:
                cost_unit = float(pricing["metric_value"] or 0.0) * float(pricing["base_value"] or 0.0)
                if cost_unit > 0 and float(initial.get("price_markup_pct", 0) or 0.0) == 0.0:
                    mp_margin_spin.setValue(round(((float(initial.get("preco_unit", 0) or 0.0) / cost_unit) - 1.0) * 100.0, 2))
            sync_line_price_from_mp()
            mp_price_info.setText(
                f"Matéria-prima ligada: {stock_line_meta['stock_material_id']} | "
                f"Base {pricing['base_value']:.4f} {pricing['base_label']} | "
                f"Equiv. {float(pricing['price_kg_equiv'] or 0.0):.4f} EUR/kg"
            )

        def refresh_mp_price_panel() -> None:
            stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
            visible = bool(stock_id) and current_type_token() == "stock_mp"
            tools_visible = current_type_token() == "stock_mp"
            if visible:
                record = self.backend.material_by_id(stock_id)
                preview = dict(self.backend.material_price_preview(record) or {}) if isinstance(record, dict) else {}
                pricing = stock_pricing_context(record, preview)
                stock_price_state.update(pricing)
                metric_label = form.labelForField(mp_metric_spin)
                if metric_label is not None:
                    metric_label.setText(pricing["metric_label"])
                price_label = form.labelForField(mp_price_kg_spin)
                if price_label is not None:
                    price_label.setText(f"Preco compra ({pricing['base_label']})")
                mp_metric_spin.setSuffix(pricing["metric_suffix"])
                mp_price_kg_spin.setSuffix("/m" if pricing["base_label"] == "EUR/m" else "/kg")
                mp_metric_spin.setValue(float(pricing["metric_value"] or 0.0))
                if bool(stock_price_state.get("price_user_touched", False)):
                    pass
                elif float(initial.get("price_base_value", 0) or 0.0) > 0:
                    mp_price_kg_spin.setValue(float(initial.get("price_base_value", 0) or 0.0))
                elif pricing["base_value"] > 0:
                    mp_price_kg_spin.setValue(float(pricing["base_value"] or 0.0))
                mp_price_info.setText(
                    f"Matéria-prima ligada: {stock_id} | {str(preview.get('formato', record.get('formato', '-')) if isinstance(record, dict) else '-')}"
                    f" | Base {float(pricing['base_value'] or 0.0):.4f} {pricing['base_label']} | "
                    f"Equiv. {float(pricing['price_kg_equiv'] or 0.0):.4f} EUR/kg"
                )
                sync_line_price_from_mp()
            else:
                mp_metric_spin.setValue(0.0)
                mp_price_kg_spin.setValue(0.0)
                mp_margin_spin.setValue(0.0)
                mp_price_info.setText("")
            for widget in (mp_metric_spin, mp_price_kg_spin, mp_margin_spin, mp_price_info):
                widget.setVisible(visible)
            stock_tools_host.setVisible(tools_visible)
            for label_widget in (
                form.labelForField(mp_metric_spin),
                form.labelForField(mp_price_kg_spin),
                form.labelForField(mp_margin_spin),
            ):
                if label_widget is not None:
                    label_widget.setVisible(visible)

        def load_ref() -> None:
            selected = ref_history.currentData()
            if isinstance(selected, dict):
                apply_reference(selected)
                return
            key = ref_ext_edit.text().strip() or ref_history.currentText().strip()
            apply_reference(refs_by_ext.get(key) or refs_by_int.get(key))

        def browse_refs() -> None:
            payload = _reference_catalog_dialog(
                self,
                self.backend.order_reference_rows("", ""),
                "Historico de referencias",
                backend=self.backend,
                current_client=client_code,
            )
            apply_reference(payload)

        def generate_ref() -> None:
            if template_mode or current_line_type() != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            ref_int_edit.setText(
                self.backend.orc_suggest_ref_interna(
                    client_code,
                    existing_refs=current_line_refs(),
                    numero=self.current_number,
                )
            )

        def pick_drawing() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar desenho",
                "",
                "Desenhos (*.pdf *.dwg *.dxf *.step *.stp *.iges *.igs *.png *.jpg *.jpeg *.bmp);;Todos (*.*)",
            )
            if path:
                drawing_edit.setText(path)

        def create_stock_material() -> None:
            dialog_editor = _MaterialEditorDialog(self.backend, dialog)
            current_kind = str(stock_kind_combo.currentData() or "").strip()
            if current_kind:
                dialog_editor.formato_combo.setCurrentText("Perfil" if current_kind == "Perfil" else current_kind)
            if dialog_editor.exec() != QDialog.Accepted:
                return
            try:
                record = self.backend.add_material(dialog_editor.payload())
            except Exception as exc:
                QMessageBox.critical(dialog, "Matéria Prima - Stock", str(exc))
                return
            reload_stock_material_rows()
            created_id = str(record.get("id", "") or "").strip()
            if created_id:
                preview = dict(self.backend.material_price_preview(record) or {})
                kind = _stock_kind(dict(record), preview)
                for index in range(stock_kind_combo.count()):
                    if str(stock_kind_combo.itemData(index) or "").strip() == kind:
                        stock_kind_combo.setCurrentIndex(index)
                        break
            refresh_stock_material_combo(created_id)
            sync_stock_material_fields()

        docs_label = QLabel("")
        docs_label.setProperty("role", "muted")
        docs_table = QTableWidget(0, 3)
        docs_table.setHorizontalHeaderLabels(["PDF", "Estado", "Caminho"])
        docs_table.verticalHeader().setVisible(False)
        docs_table.setEditTriggers(QTableWidget.NoEditTriggers)
        docs_table.setSelectionBehavior(QTableWidget.SelectRows)
        docs_table.setSelectionMode(QAbstractItemView.SingleSelection)
        docs_table.setMinimumHeight(118)
        docs_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        docs_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        docs_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)

        def _resolve_pdf_doc_path(raw: str) -> Path:
            resolver = getattr(self.backend, "_resolve_file_reference", None)
            if callable(resolver):
                try:
                    resolved = resolver(raw)
                    if resolved is not None:
                        return Path(resolved)
                except Exception:
                    pass
            return Path(raw)

        def _selected_doc_index() -> int:
            current = docs_table.currentItem()
            if current is None or current.row() >= len(pdf_docs):
                return -1
            return current.row()

        def _render_pdf_docs() -> None:
            docs_table.setRowCount(len(pdf_docs))
            for row_index, pdf_path in enumerate(pdf_docs):
                path_obj = _resolve_pdf_doc_path(pdf_path)
                values = [path_obj.name or pdf_path, "OK" if path_obj.exists() else "Em falta", pdf_path]
                for col_index, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    if col_index == 1:
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                        item.setForeground(QBrush(QColor("#4f7f1f" if path_obj.exists() else "#b45f06")))
                    docs_table.setItem(row_index, col_index, item)
            docs_label.setText(
                f"{len(pdf_docs)} PDF(s) associado(s)." if pdf_docs else "Sem PDFs técnicos associados a esta peça."
            )

        def _add_pdf_docs() -> None:
            paths, _ = QFileDialog.getOpenFileNames(
                dialog,
                "Adicionar PDFs técnicos",
                "",
                "PDF (*.pdf);;Todos (*.*)",
            )
            for path in paths:
                clean = str(path or "").strip()
                if clean and clean.lower().endswith(".pdf") and clean not in pdf_docs:
                    pdf_docs.append(clean)
            _render_pdf_docs()

        def _open_pdf_doc() -> None:
            index = _selected_doc_index()
            if index < 0:
                QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
                return
            path = _resolve_pdf_doc_path(pdf_docs[index])
            if not path.exists():
                QMessageBox.critical(dialog, "PDFs associados", f"PDF não encontrado:\n{path}")
                return
            os.startfile(str(path))

        def _download_pdf_doc() -> None:
            index = _selected_doc_index()
            if index < 0:
                QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
                return
            source = _resolve_pdf_doc_path(pdf_docs[index])
            if not source.exists():
                QMessageBox.critical(dialog, "PDFs associados", f"PDF não encontrado:\n{source}")
                return
            target, _ = QFileDialog.getSaveFileName(dialog, "Guardar cópia do PDF", source.name, "PDF (*.pdf)")
            if not target:
                return
            try:
                import shutil

                shutil.copy2(source, target)
            except Exception as exc:
                QMessageBox.critical(dialog, "PDFs associados", str(exc))

        def _replace_pdf_doc() -> None:
            index = _selected_doc_index()
            if index < 0:
                QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
                return
            path, _ = QFileDialog.getOpenFileName(dialog, "Substituir PDF", "", "PDF (*.pdf);;Todos (*.*)")
            clean = str(path or "").strip()
            if clean:
                pdf_docs[index] = clean
                _render_pdf_docs()
                docs_table.selectRow(index)

        def _remove_pdf_doc() -> None:
            index = _selected_doc_index()
            if index < 0:
                QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
                return
            del pdf_docs[index]
            _render_pdf_docs()

        docs_buttons = QHBoxLayout()
        docs_buttons.setSpacing(8)
        btn_doc_open = QPushButton("Visualizar")
        btn_doc_open.setProperty("variant", "secondary")
        btn_doc_download = QPushButton("Download")
        btn_doc_download.setProperty("variant", "secondary")
        btn_doc_add = QPushButton("Adicionar PDF")
        btn_doc_replace = QPushButton("Substituir")
        btn_doc_replace.setProperty("variant", "secondary")
        btn_doc_remove = QPushButton("Apagar")
        btn_doc_remove.setProperty("variant", "danger")
        for button in (btn_doc_open, btn_doc_download, btn_doc_add, btn_doc_replace, btn_doc_remove):
            docs_buttons.addWidget(button)
        docs_buttons.addStretch(1)
        docs_buttons_host = QWidget()
        docs_buttons_host.setLayout(docs_buttons)
        docs_host = QWidget()
        docs_layout = QVBoxLayout(docs_host)
        docs_layout.setContentsMargins(0, 0, 0, 0)
        docs_layout.setSpacing(6)
        docs_layout.addWidget(docs_label)
        docs_layout.addWidget(docs_table)
        docs_layout.addWidget(docs_buttons_host)
        btn_doc_open.clicked.connect(_open_pdf_doc)
        btn_doc_download.clicked.connect(_download_pdf_doc)
        btn_doc_add.clicked.connect(_add_pdf_docs)
        btn_doc_replace.clicked.connect(_replace_pdf_doc)
        btn_doc_remove.clicked.connect(_remove_pdf_doc)
        _render_pdf_docs()

        def selected_operation_names() -> list[str]:
            selected_ops: list[str] = []
            for token in _operation_tokens(operation_edit.text().strip()):
                normalized = str(self.backend.desktop_main.normalize_operacao_nome(token) or token or "").strip()
                if normalized and normalized not in selected_ops:
                    selected_ops.append(normalized)
            return selected_ops

        def _normalized_op_key(value: str) -> str:
            return str(self.backend.desktop_main.normalize_operacao_nome(value) or value or "").strip()

        def line_has_laser_base() -> bool:
            if current_line_type() != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
                return False
            if not bool(base_state.get("laser_base_enabled", False)):
                return False
            if not drawing_edit.text().strip():
                return False
            return float(tempo_spin.value() or 0) > 0 or float(price_spin.value() or 0) > 0

        def full_route_operation_names() -> list[str]:
            route = list(selected_operation_names())
            if line_has_laser_base() and "Corte Laser" not in route:
                route.insert(0, "Corte Laser")
            return route

        def costing_operation_names() -> list[str]:
            selected_ops = full_route_operation_names()
            if line_has_laser_base():
                return [op_name for op_name in selected_ops if op_name != "Corte Laser"]
            return list(selected_ops)

        def current_extra_totals() -> tuple[float, float]:
            extra_ops = {_normalized_op_key(op_name) for op_name in costing_operation_names() if _normalized_op_key(op_name)}
            tempo_extra = 0.0
            custo_extra = 0.0
            for op_name, raw_value in dict(operation_meta.get("tempos_operacao", {}) or {}).items():
                normalized = _normalized_op_key(str(op_name or ""))
                if normalized in extra_ops:
                    tempo_extra += float(raw_value or 0)
            for op_name, raw_value in dict(operation_meta.get("custos_operacao", {}) or {}).items():
                normalized = _normalized_op_key(str(op_name or ""))
                if normalized in extra_ops:
                    custo_extra += float(raw_value or 0)
            return round(tempo_extra, 4), round(custo_extra, 4)

        def current_laser_base_totals() -> tuple[float, float]:
            if not line_has_laser_base():
                return float(tempo_spin.value() or 0), float(price_spin.value() or 0)
            return (
                round(float(base_state.get("laser_base_time", 0) or 0), 4),
                round(float(base_state.get("laser_base_price", 0) or 0), 4),
            )

        def current_composed_totals() -> tuple[float, float]:
            if not line_has_laser_base():
                return round(float(tempo_spin.value() or 0), 4), round(float(price_spin.value() or 0), 4)
            base_time, base_cost = current_laser_base_totals()
            extra_time, extra_cost = current_extra_totals()
            return round(base_time + extra_time, 4), round(base_cost + extra_cost, 4)

        def sync_line_totals_from_meta() -> None:
            if not line_has_laser_base():
                return
            composed_time, composed_cost = current_composed_totals()
            previous_tempo_state = tempo_spin.blockSignals(True)
            previous_price_state = price_spin.blockSignals(True)
            try:
                tempo_spin.setValue(float(composed_time or 0))
                price_spin.setValue(float(composed_cost or 0))
            finally:
                tempo_spin.blockSignals(previous_tempo_state)
                price_spin.blockSignals(previous_price_state)

        def sync_operation_meta_selection() -> None:
            selected_ops = full_route_operation_names()
            selected_keys = {_normalized_op_key(op_name) for op_name in selected_ops if _normalized_op_key(op_name)}
            previous_keys = {_normalized_op_key(op_name) for op_name in list(operation_meta.get("operacoes_lista", []) or []) if _normalized_op_key(op_name)}
            operation_meta["operacoes_lista"] = list(selected_ops)
            operation_meta["operacoes_fluxo"] = self.backend.desktop_main.build_operacoes_fluxo(
                selected_ops,
                operation_meta.get("operacoes_fluxo") if isinstance(operation_meta.get("operacoes_fluxo"), list) else None,
            )
            operation_meta["operacoes_detalhe"] = [
                dict(item or {})
                for item in list(operation_meta.get("operacoes_detalhe", []) or [])
                if _normalized_op_key(str((item or {}).get("nome", "") or "")) in selected_keys
            ]
            operation_meta["tempos_operacao"] = {
                _normalized_op_key(str(op_name or "")): float(raw_value or 0)
                for op_name, raw_value in dict(operation_meta.get("tempos_operacao", {}) or {}).items()
                if _normalized_op_key(str(op_name or "")) in selected_keys
            }
            operation_meta["custos_operacao"] = {
                _normalized_op_key(str(op_name or "")): float(raw_value or 0)
                for op_name, raw_value in dict(operation_meta.get("custos_operacao", {}) or {}).items()
                if _normalized_op_key(str(op_name or "")) in selected_keys
            }
            if previous_keys != selected_keys:
                operation_meta["quote_cost_snapshot"] = {}
                operation_change_state["last_prompt_signature"] = ""

        def current_operation_cost_payload() -> dict[str, Any]:
            base_blend = line_has_laser_base()
            base_time, base_cost = current_laser_base_totals()
            return {
                "operacao": operation_edit.text().strip(),
                "operacoes_lista": list(selected_operation_names()),
                "costing_operations": list(costing_operation_names()),
                "qtd": qtd_spin.value(),
                "tempo_peca_min": tempo_spin.value(),
                "preco_unit": price_spin.value(),
                "area_m2": float(initial.get("area_m2", initial.get("net_area_m2", 0)) or 0),
                "blend_with_current_line": base_blend,
                "base_tempo_unit_min": base_time if base_blend else 0.0,
                "base_preco_unit_eur": base_cost if base_blend else 0.0,
                "base_operation_label": "Laser base",
                "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
                "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
            }

        def pending_operation_inputs(estimate: dict[str, Any]) -> list[dict[str, Any]]:
            pending_rows: list[dict[str, Any]] = []
            for item in list(estimate.get("operations", []) or []):
                if not isinstance(item, dict):
                    continue
                if bool(item.get("missing_driver_input")):
                    pending_rows.append(dict(item))
            return pending_rows

        def refresh_operation_cost_hint() -> None:
            ops_txt = operation_edit.text().strip()
            if not ops_txt:
                operation_cost_label.setText("Sem postos selecionados nesta linha.")
                return
            estimate = dict(self.backend.operation_cost_estimate(current_operation_cost_payload()) or {})
            summary = dict(estimate.get("summary", {}) or {})
            pending_rows = pending_operation_inputs(estimate)
            base_blend = bool(current_operation_cost_payload().get("blend_with_current_line"))
            extra_ops = list(costing_operation_names())
            state_txt = {
                "detailed": "detalhe completo",
                "partial_detail": "detalhe parcial",
                "aggregate_pending": "preco ainda agregado",
                "single_operation_total": "operacao simples",
            }.get(str(summary.get("costing_mode", "") or ""), "sem detalhe")
            if base_blend and not extra_ops:
                base_time, base_cost = current_laser_base_totals()
                operation_cost_label.setText(
                    f"Base laser carregada na linha: {_fmt_eur(base_cost)}/un | {base_time:.3f} min/un. "
                    "Seleciona as operacoes seguintes para agregar custo."
                )
                return
            if pending_rows:
                missing_txt = ", ".join(
                    f"{str(item.get('nome', '') or '').strip()} ({str(item.get('driver_label', '') or 'Qtd./peca').strip()})"
                    for item in pending_rows
                    if str(item.get("nome", "") or "").strip()
                )
                operation_cost_label.setText(
                    f"Falta quantificar: {missing_txt}. Abre a quantificacao para fechar o custo e somar ao valor do laser."
                )
                return
            if base_blend:
                base_time, base_cost = current_laser_base_totals()
                extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
                extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
                total_time = base_time + extra_time
                total_cost = base_cost + extra_cost
                operation_cost_label.setText(
                    f"Perfis operacao: {str(estimate.get('active_profile', '') or 'Base')} | "
                    f"base atual {_fmt_eur(base_cost)}/un + extras {_fmt_eur(extra_cost)}/un = {_fmt_eur(total_cost)}/un | "
                    f"tempo {base_time:.3f} + {extra_time:.3f} = {total_time:.3f} min/un"
                )
                return
            operation_cost_label.setText(
                f"Perfis operacao: {str(estimate.get('active_profile', '') or 'Base')} | {state_txt} | "
                f"sugestao {float(summary.get('tempo_unit_total_min', 0) or 0):.3f} min/un | "
                f"{_fmt_eur(float(summary.get('custo_unit_total_eur', 0) or 0))}/un"
            )

        def edit_operation_costs(*, auto_prompt: bool = False) -> bool:
            if line_has_laser_base() and not costing_operation_names():
                if not auto_prompt:
                    QMessageBox.information(
                        self,
                        "Operacoes",
                        "Esta linha ja tem o laser como base. Seleciona as operacoes seguintes para quantificar os acrescimos.",
                    )
                return False
            result = _open_quote_operation_detail_dialog(self, self.backend, current_operation_cost_payload())
            if not isinstance(result, dict):
                return False
            operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(result.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
            operation_meta["tempos_operacao"] = dict(result.get("tempos_operacao", {}) or {})
            operation_meta["custos_operacao"] = dict(result.get("custos_operacao", {}) or {})
            operation_meta["quote_cost_snapshot"] = dict(result.get("quote_cost_snapshot", {}) or {})
            if line_has_laser_base():
                sync_line_totals_from_meta()
            elif bool(result.get("apply_totals")):
                tempo_spin.setValue(float(result.get("suggested_tempo_unit_min", tempo_spin.value()) or 0))
                price_spin.setValue(float(result.get("suggested_preco_unit_eur", price_spin.value()) or 0))
            if auto_prompt:
                operation_change_state["last_prompt_signature"] = operation_edit.text().strip()
            refresh_operation_cost_hint()
            return True

        def maybe_prompt_operation_breakdown() -> None:
            if current_line_type() != self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            estimate = dict(self.backend.operation_cost_estimate(current_operation_cost_payload()) or {})
            pending_rows = pending_operation_inputs(estimate)
            if not pending_rows:
                return
            signature = "|".join(
                [" + ".join(costing_operation_names())]
                + [str(item.get("nome", "") or "").strip() for item in pending_rows if str(item.get("nome", "") or "").strip()]
            )
            if signature == str(operation_change_state.get("last_prompt_signature", "") or ""):
                return
            operation_change_state["last_prompt_signature"] = signature
            edit_operation_costs(auto_prompt=True)

        def handle_operation_text_changed() -> None:
            user_initiated = bool(operation_change_state.get("user_initiated", False))
            operation_change_state["user_initiated"] = False
            sync_operation_meta_selection()
            sync_line_totals_from_meta()
            refresh_operation_cost_hint()
            if user_initiated:
                maybe_prompt_operation_breakdown()

        ref_history.currentIndexChanged.connect(lambda _index: apply_reference(ref_history.currentData()))
        product_combo.currentTextChanged.connect(lambda _value: sync_product_fields())
        operation_edit.textChanged.connect(lambda _value: handle_operation_text_changed())
        qtd_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())
        tempo_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())
        price_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())

        ref_buttons = QHBoxLayout()
        btn_generate = QPushButton("Gerar interna")
        btn_generate.setProperty("variant", "secondary")
        btn_generate.clicked.connect(generate_ref)
        btn_history = QPushButton("Referencias criadas")
        btn_history.setProperty("variant", "secondary")
        btn_history.clicked.connect(browse_refs)
        btn_load = QPushButton("Carregar referencia")
        btn_load.setProperty("variant", "secondary")
        btn_load.clicked.connect(load_ref)
        btn_drawing = QPushButton("Selecionar desenho")
        btn_drawing.setProperty("variant", "secondary")
        btn_drawing.clicked.connect(pick_drawing)
        ref_buttons.addWidget(btn_generate)
        ref_buttons.addWidget(btn_history)
        ref_buttons.addWidget(btn_load)
        ref_buttons.addWidget(btn_drawing)
        ref_buttons.addStretch(1)
        ref_buttons_host = QWidget()
        ref_buttons_host.setLayout(ref_buttons)

        operation_buttons = QHBoxLayout()
        operation_buttons.setSpacing(8)
        btn_ops_detail = QPushButton("Quantificar operacoes seguintes")
        btn_ops_detail.setProperty("variant", "secondary")
        btn_ops_detail.clicked.connect(lambda: edit_operation_costs(auto_prompt=False))
        btn_ops_profiles = QPushButton("Perfis operacoes")
        btn_ops_profiles.setProperty("variant", "secondary")
        def configure_operation_profiles_from_line() -> None:
            if _open_operation_cost_profiles_dialog(self, self.backend):
                refresh_operation_cost_hint()
        btn_ops_profiles.clicked.connect(configure_operation_profiles_from_line)
        operation_buttons.addWidget(btn_ops_detail)
        operation_buttons.addWidget(btn_ops_profiles)
        operation_buttons.addStretch(1)
        operation_buttons_host = QWidget()
        operation_buttons_host.setLayout(operation_buttons)

        form.addRow("Tipo", type_combo)
        form.addRow("Historico", ref_history)
        form.addRow("Produto", product_combo)
        form.addRow("Ref. interna", ref_int_edit)
        form.addRow("Ref. externa", ref_ext_edit)
        form.addRow("Descricao", desc_edit)
        form.addRow("Dimensao", dimension_edit)
        form.addRow("Codigo/unid", product_unid_edit)
        form.addRow("Tipo matéria-prima", stock_kind_combo)
        form.addRow("Matéria Prima - Stock", stock_material_combo)
        form.addRow("Kg por unid.", mp_metric_spin)
        form.addRow("Preço matéria-prima", mp_price_kg_spin)
        form.addRow("Margem", mp_margin_spin)
        form.addRow("", stock_tools_host)
        form.addRow("", mp_price_info)
        form.addRow("Material", material_combo)
        form.addRow("Espessura", esp_combo)
        form.addRow("Postos", operation_selector)
        form.addRow("Custeio op.", operation_cost_label)
        form.addRow("", operation_buttons_host)
        form.addRow("Tempo peca (min)", tempo_spin)
        form.addRow("Quantidade", qtd_spin)
        form.addRow("Preco unit.", price_spin)
        form.addRow("Desenho", drawing_edit)
        form.addRow("", ref_buttons_host)
        form.addRow("PDFs associados", docs_host)
        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.NoFrame)
        form_scroll.setWidget(form_host)
        layout.addWidget(form_scroll, 1)
        mode_state = {"initialized": False, "last_token": initial_type}

        def sync_mode() -> None:
            token = current_type_token()
            line_type = current_line_type()
            is_stock_mp = token == "stock_mp"
            entered_stock_mp = bool(mode_state["initialized"] and is_stock_mp and mode_state.get("last_token") != "stock_mp")
            is_piece = line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE
            is_product = line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT
            ref_history.setEnabled(is_piece)
            ref_int_edit.setEnabled(is_piece and not template_mode and not is_stock_mp)
            ref_ext_edit.setEnabled(is_piece and not is_stock_mp)
            material_combo.setEnabled(is_piece)
            esp_combo.setEnabled(is_piece)
            dimension_edit.setEnabled(is_piece)
            drawing_edit.setEnabled(is_piece and not is_stock_mp)
            btn_generate.setEnabled(is_piece and not template_mode and not is_stock_mp)
            btn_history.setEnabled(is_piece and not is_stock_mp)
            btn_load.setEnabled(is_piece and not is_stock_mp)
            btn_drawing.setEnabled(is_piece and not is_stock_mp)
            btn_ops_detail.setEnabled(is_piece and not is_stock_mp)
            btn_ops_profiles.setEnabled(is_piece and not is_stock_mp)
            product_combo.setEnabled(is_product)
            stock_kind_combo.setEnabled(is_stock_mp)
            stock_material_combo.setEnabled(is_stock_mp)
            set_form_row_visible(stock_kind_combo, is_stock_mp, "Tipo matéria-prima")
            set_form_row_visible(stock_material_combo, is_stock_mp, "Matéria Prima - Stock")
            product_unid_edit.setEnabled(is_product or line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE)
            set_form_row_visible(ref_history, is_piece and not is_stock_mp, "Historico")
            set_form_row_visible(ref_int_edit, is_piece and not is_stock_mp and not template_mode, "Ref. interna")
            set_form_row_visible(ref_ext_edit, is_piece and not is_stock_mp, "Ref. externa")
            set_form_row_visible(product_combo, is_product, "Produto")
            set_form_row_visible(product_unid_edit, is_product or line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE, "Codigo/unid")
            set_form_row_visible(material_combo, is_piece, "Material")
            set_form_row_visible(esp_combo, is_piece, "Espessura")
            set_form_row_visible(dimension_edit, is_stock_mp, "Dimensao")
            show_stock_commercial = is_stock_mp
            show_operations = is_piece and not is_stock_mp
            set_form_row_visible(operation_selector, show_operations, "Postos")
            set_form_row_visible(operation_cost_label, show_operations, "Custeio op.")
            operation_buttons_host.setVisible(show_operations)
            set_form_row_visible(tempo_spin, show_operations, "Tempo peca (min)")
            set_form_row_visible(drawing_edit, show_operations, "Desenho")
            set_form_row_visible(docs_host, show_operations, "PDFs associados")
            ref_buttons_host.setVisible(show_operations)
            set_form_row_visible(mp_metric_spin, show_stock_commercial, stock_price_state.get("metric_label", "Kg por unid."))
            set_form_row_visible(mp_price_kg_spin, show_stock_commercial, f"Preco compra ({stock_price_state.get('base_label', 'EUR/kg')})")
            set_form_row_visible(mp_margin_spin, show_stock_commercial, "Margem")
            if is_product:
                base_state["laser_base_enabled"] = False
                base_state["laser_base_time"] = 0.0
                base_state["laser_base_price"] = 0.0
                ref_int_edit.clear()
                ref_ext_edit.clear()
                material_combo.setCurrentText("")
                esp_combo.setCurrentText("")
                dimension_edit.clear()
                drawing_edit.clear()
                apply_operations("")
                tempo_spin.setValue(0.0)
                stock_line_meta["stock_material_id"] = ""
                stock_line_meta["material_family"] = ""
                stock_line_meta["material_subtype"] = ""
                sync_product_fields()
            elif is_stock_mp:
                base_state["laser_base_enabled"] = False
                base_state["laser_base_time"] = 0.0
                base_state["laser_base_price"] = 0.0
                ref_int_edit.clear()
                drawing_edit.clear()
                if entered_stock_mp:
                    desc_edit.clear()
                    ref_ext_edit.clear()
                    dimension_edit.clear()
                    material_combo.setCurrentText("")
                    esp_combo.setCurrentText("")
                    stock_line_meta["stock_material_id"] = ""
                    stock_line_meta["material_family"] = ""
                    stock_line_meta["material_subtype"] = ""
                    stock_material_combo.setCurrentIndex(0)
                    refresh_stock_material_combo("")
                apply_operations("")
                tempo_spin.setValue(0.0)
                sync_stock_material_fields()
            elif line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE:
                base_state["laser_base_enabled"] = False
                base_state["laser_base_time"] = 0.0
                base_state["laser_base_price"] = 0.0
                ref_int_edit.clear()
                ref_ext_edit.clear()
                material_combo.setCurrentText("")
                esp_combo.setCurrentText("")
                dimension_edit.clear()
                drawing_edit.clear()
                product_unid_edit.setText("SV")
                if not operation_edit.text().strip():
                    apply_operations("Montagem")
            else:
                product_unid_edit.clear()
                base_state["laser_base_enabled"] = bool(base_state.get("laser_base_enabled", False) or _payload_has_laser_base(initial))
                if base_state["laser_base_enabled"] and base_state["laser_base_time"] <= 0 and base_state["laser_base_price"] <= 0:
                    base_state["laser_base_time"] = round(float(tempo_spin.value() or 0), 4)
                    base_state["laser_base_price"] = round(float(price_spin.value() or 0), 4)
                if not template_mode and not ref_int_edit.text().strip():
                    generate_ref()
            if not is_stock_mp and line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE and str(initial.get("stock_material_id", "") or "").strip():
                stock_line_meta["stock_material_id"] = str(initial.get("stock_material_id", "") or "").strip()
            sync_operation_meta_selection()
            sync_line_totals_from_meta()
            refresh_operation_cost_hint()
            refresh_mp_price_panel()
            mode_state["initialized"] = True
            mode_state["last_token"] = token

        type_combo.currentTextChanged.connect(lambda _value: sync_mode())
        stock_kind_combo.currentTextChanged.connect(lambda _value: (refresh_stock_material_combo(), sync_stock_material_fields()))
        stock_material_combo.currentTextChanged.connect(lambda _value: sync_stock_material_fields())
        mp_price_kg_spin.valueChanged.connect(lambda _value: sync_line_price_from_mp())
        mp_margin_spin.valueChanged.connect(lambda _value: sync_line_price_from_mp())
        price_spin.valueChanged.connect(lambda _value: sync_margin_from_line_price())
        mp_price_kg_spin.lineEdit().textEdited.connect(lambda _text: stock_price_state.__setitem__("price_user_touched", True))
        mp_price_table_btn.clicked.connect(
            lambda: (
                self._material_price_manager_dialog(
                    str(stock_kind_combo.currentData() or "").strip(),
                    str(stock_line_meta.get("stock_material_id", "") or "").strip(),
                    parent=dialog,
                ),
                reload_stock_material_rows(),
                refresh_stock_material_combo(str(stock_line_meta.get("stock_material_id", "") or "").strip()),
                refresh_mp_price_panel(),
            )
        )
        stock_create_btn.clicked.connect(create_stock_material)
        sync_mode()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        def accept_line_dialog() -> None:
            try:
                if current_type_token() == "stock_mp":
                    record, preview = stock_material_detail()
                    if not record and not material_combo.currentText().strip():
                        QMessageBox.warning(dialog, "Matéria Prima - Stock", "Seleciona stock de matéria-prima ou indica o material manualmente.")
                        return
                    if not desc_edit.text().strip():
                        if record:
                            sync_stock_material_fields()
                        else:
                            desc_edit.setText(" ".join(part for part in (material_combo.currentText().strip(), dimension_edit.text().strip()) if part).strip())
                    if not material_combo.currentText().strip():
                        material_combo.setCurrentText(str(record.get("material", "") or "").strip())
                    if record:
                        esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                        if esp_txt and not esp_combo.currentText().strip():
                            esp_combo.setCurrentText(esp_txt)
                dialog.accept()
            except Exception as exc:
                QMessageBox.critical(dialog, "Linha de orcamento", str(exc))
        buttons.accepted.connect(accept_line_dialog)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        sync_operation_meta_selection()
        line_type = current_line_type()
        token = current_type_token()
        stock_row = None
        if str(stock_line_meta.get("stock_material_id", "") or "").strip():
            stock_row = next(
                (
                    row
                    for row in list(self.backend.material_price_rows() or [])
                    if str(row.get("id", "") or "").strip() == str(stock_line_meta.get("stock_material_id", "") or "").strip()
                ),
                None,
            )
        stock_update = None
        stock_pricing_label = str(stock_price_state.get("base_label", "EUR/kg") or "EUR/kg").strip() or "EUR/kg"
        stock_price_base_value = float(mp_price_kg_spin.value() or 0.0)
        stock_price_kg_value = stock_price_base_value
        if stock_pricing_label == "EUR/m":
            kg_m_value = float((stock_row or {}).get("kg_m", initial.get("kg_per_m", 0)) or 0.0)
            stock_price_kg_value = round((stock_price_base_value / kg_m_value), 4) if kg_m_value > 0 else 0.0
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE and str(stock_line_meta.get("stock_material_id", "") or "").strip():
            if isinstance(stock_row, dict):
                stock_update = self._sync_stock_price_from_context(
                    str(stock_line_meta.get("stock_material_id", "") or "").strip(),
                    float(stock_price_kg_value or 0.0),
                    float(stock_row.get("price_kg", 0) or 0.0),
                    parent=dialog,
                )
        stock_update_payload = dict(stock_update or {}) if isinstance(stock_update, dict) else {}
        final_time_unit, final_price_unit = current_composed_totals()
        if float(stock_update_payload.get("price_kg", 0) or 0.0) > 0:
            if stock_pricing_label == "EUR/m":
                base_unit_value = round(float(stock_update_payload.get("price_kg", 0) or 0.0) * float((stock_row or {}).get("kg_m", initial.get("kg_per_m", 0)) or 0.0), 4)
            else:
                base_unit_value = float(stock_update_payload.get("price_kg", 0) or 0.0)
            final_price_unit = round(
                float(mp_metric_spin.value() or current_piece_metric_kg() or 0.0)
                * base_unit_value
                * (1.0 + (float(mp_margin_spin.value() or 0.0) / 100.0)),
                4,
            )
        elif token == "stock_mp":
            final_time_unit = 0.0
            final_price_unit = round(float(price_spin.value() or 0.0), 4)
        is_raw_stock_line = token == "stock_mp"
        is_product_line = line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT
        route_operations = full_route_operation_names()
        raw_has_technical_route = bool(
            is_raw_stock_line
            and (
                route_operations
                or list(operation_meta.get("operacoes_detalhe", []) or [])
                or dict(operation_meta.get("tempos_operacao", {}) or {})
                or dict(operation_meta.get("custos_operacao", {}) or {})
            )
        )
        allow_technical_docs = bool(
            line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE
            and (not is_raw_stock_line or raw_has_technical_route)
        )
        final_ref_externa = ""
        if is_raw_stock_line:
            record, _preview = stock_material_detail()
            final_ref_externa = str((record or {}).get("id", "") or stock_line_meta.get("stock_material_id", "") or "").strip()
        elif line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            final_ref_externa = ref_ext_edit.text().strip()
        final_operation = ""
        if not is_product_line and (not is_raw_stock_line or raw_has_technical_route):
            final_operation = " + ".join(route_operations)
        keep_operation_meta = not is_product_line and (not is_raw_stock_line or raw_has_technical_route)
        return {
            "tipo_item": line_type,
            "stock_item_kind": "raw_material" if is_raw_stock_line else ("product" if is_product_line else str(initial.get("stock_item_kind", "") or "").strip()),
            "ref_interna": "" if (template_mode or line_type != self.backend.desktop_main.ORC_LINE_TYPE_PIECE or is_raw_stock_line) else ref_int_edit.text().strip(),
            "ref_externa": final_ref_externa,
            "descricao": desc_edit.text().strip(),
            "dimensao": dimension_edit.text().strip(),
            "dimensoes": dimension_edit.text().strip(),
            "material": material_combo.currentText().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "material_family": str(stock_line_meta.get("material_family", "") or initial.get("material_family", "") or (material_combo.currentText().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "")).strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "material_subtype": str(
                stock_line_meta.get("material_subtype", "")
                or initial.get("material_subtype", "")
                or (str(stock_kind_combo.currentData() or "").strip() if token == "stock_mp" else "")
            ).strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "material_supplied_by_client": bool(initial.get("material_supplied_by_client", False) or initial.get("material_fornecido_cliente", False)) if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else False,
            "material_fornecido_cliente": bool(initial.get("material_fornecido_cliente", False) or initial.get("material_supplied_by_client", False)) if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else False,
            "material_cost_included": (
                (
                    bool(initial.get("material_cost_included", True))
                    if "material_cost_included" in initial
                    else not bool(initial.get("material_supplied_by_client", False) or initial.get("material_fornecido_cliente", False))
                )
                if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE
                else False
            ),
            "espessura": esp_combo.currentText().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "operacao": final_operation,
            "tempo_peca_min": final_time_unit,
            "qtd": qtd_spin.value(),
            "qtd_base": float(qtd_spin.value() if template_mode else initial.get("qtd_base", qtd_spin.value())),
            "preco_unit": final_price_unit,
            "desenho": drawing_edit.text().strip() if allow_technical_docs else "",
            "desenho_pdf": pdf_docs[0] if allow_technical_docs and pdf_docs else "",
            "desenhos_pdf": list(pdf_docs) if allow_technical_docs else [],
            "ficheiros": [
                item
                for item in [drawing_edit.text().strip(), *pdf_docs]
                if item
            ] if allow_technical_docs else [],
            "stock_material_id": str(stock_line_meta.get("stock_material_id", "") or "").strip() if (token == "stock_mp" and stock_material_detail()[0]) or str(initial.get("stock_material_id", "") or "").strip() else "",
            "price_per_kg": float(stock_update_payload.get("price_kg", stock_price_kg_value) or 0.0) if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else 0.0,
            "price_base_value": float(stock_price_base_value or 0.0) if token == "stock_mp" else float(initial.get("price_base_value", 0) or 0.0),
            "price_base_label": stock_pricing_label if token == "stock_mp" else str(initial.get("price_base_label", "") or "").strip(),
            "price_markup_pct": float(mp_margin_spin.value() or 0.0) if token == "stock_mp" else float(initial.get("price_markup_pct", 0) or 0.0),
            "stock_metric_value": float(mp_metric_spin.value() or 0.0) if token == "stock_mp" else float(initial.get("stock_metric_value", 0) or 0.0),
            "kg_per_m": float(initial.get("kg_per_m", stock_row.get("kg_m", 0) if isinstance(stock_row, dict) else 0) or 0.0),
            "laser_base_active": bool(line_has_laser_base()),
            "laser_base_tempo_unit": current_laser_base_totals()[0] if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else 0.0,
            "laser_base_preco_unit": current_laser_base_totals()[1] if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else 0.0,
            "produto_codigo": current_product_code() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT else "",
            "produto_unid": product_unid_edit.text().strip() if line_type != self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "_product_pending_create": bool(line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT and not current_product_code()),
            "conjunto_codigo": str(initial.get("conjunto_codigo", "") or "").strip(),
            "conjunto_nome": str(initial.get("conjunto_nome", "") or "").strip(),
            "grupo_uuid": str(initial.get("grupo_uuid", "") or "").strip(),
            "operacoes_lista": list(operation_meta.get("operacoes_lista", []) or []) if keep_operation_meta else [],
            "operacoes_fluxo": [dict(item or {}) for item in list(operation_meta.get("operacoes_fluxo", []) or []) if isinstance(item, dict)] if keep_operation_meta else [],
            "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)] if keep_operation_meta else [],
            "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}) if keep_operation_meta else {},
            "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}) if keep_operation_meta else {},
            "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}) if keep_operation_meta else {},
        }

    def _assembly_model_editor_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Modelo de conjunto")
        dialog.resize(920, 640)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        code_edit = QLineEdit(str(initial.get("codigo", "") or "").strip())
        desc_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
        notes_edit = QTextEdit()
        notes_edit.setMinimumHeight(76)
        notes_edit.setPlainText(str(initial.get("notas", "") or "").strip())
        form.addRow("Codigo", code_edit)
        form.addRow("Descricao", desc_edit)
        form.addRow("Notas", notes_edit)
        layout.addLayout(form)

        intro = QLabel(
            "Modelo/Conjunto tecnico separado das linhas DXF/DWG. "
            "Aqui trabalhas sempre com itens de conjunto: material, mao de obra, consumiveis e produtos."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        items: list[dict] = [self._wrap_assembly_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
        items_table = QTableWidget(0, 7)
        items_table.setHorizontalHeaderLabels(["Tipo", "Descricao", "Codigo/Ref", "Material", "Esp./Unid", "Qtd", "Preco"])
        items_table.verticalHeader().setVisible(False)
        items_table.setEditTriggers(QTableWidget.NoEditTriggers)
        items_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(items_table, stretch=(1, 3), contents=(0, 2, 4, 5, 6))

        def render_items() -> None:
            _fill_table(
                items_table,
                [
                    [
                        self._assembly_item_kind_label(self._assembly_item_kind_from_line(item)),
                        str(((item.get("line") or {}).get("descricao", "") or "-")).strip() or "-",
                        str(((item.get("line") or {}).get("produto_codigo", "") or (item.get("line") or {}).get("ref_externa", "") or "-")).strip() or "-",
                        str(((item.get("line") or {}).get("material", "") or "-")).strip() or "-",
                        str(((item.get("line") or {}).get("espessura", "") or (item.get("line") or {}).get("produto_unid", "") or "-")).strip() or "-",
                        f"{float(((item.get('line') or {}).get('qtd', 0) or 0)):.2f}",
                        _fmt_eur(float(((item.get("line") or {}).get("preco_unit", 0) or 0))),
                    ]
                    for item in items
                ],
                align_center_from=4,
            )

        def selected_item_index() -> int:
            current = items_table.currentItem()
            if current is None or current.row() >= len(items):
                return -1
            return current.row()

        actions = QHBoxLayout()
        add_btn = QPushButton("Adicionar item")
        edit_btn = QPushButton("Editar item")
        edit_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover item")
        remove_btn.setProperty("variant", "danger")
        actions.addWidget(add_btn)
        actions.addWidget(edit_btn)
        actions.addWidget(remove_btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        layout.addWidget(items_table, 1)

        def add_item() -> None:
            kind = self._pick_assembly_item_kind(dialog)
            if not kind:
                return
            payload = self._open_assembly_item_editor(kind, parent=dialog)
            if payload is None:
                return
            items.append(payload)
            render_items()

        def edit_item() -> None:
            index = selected_item_index()
            if index < 0:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um item.")
                return
            current = dict(items[index] or {})
            current_line = dict(current.get("line") or current)
            if self._quote_line_is_laser_2d(current_line):
                batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
                batch_indexes = [
                    row_index
                    for row_index, candidate in enumerate(items)
                    if batch_id
                    and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
                ] or [index]
                source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
                edited_lines = self._edit_laser_batch_lines(source_lines, parent=dialog)
                if edited_lines is None:
                    return
                insert_at = min(batch_indexes)
                for row_index in sorted(batch_indexes, reverse=True):
                    del items[row_index]
                for offset, edited_line in enumerate(edited_lines):
                    items.insert(insert_at + offset, self._wrap_assembly_item(edited_line))
                render_items()
                if edited_lines:
                    items_table.selectRow(insert_at)
                return
            else:
                payload = self._open_assembly_item_editor(self._assembly_item_kind_from_line(current), current, parent=dialog)
            if payload is None:
                return
            items[index] = payload
            render_items()
            items_table.selectRow(index)

        def remove_item() -> None:
            index = selected_item_index()
            if index < 0:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um item.")
                return
            del items[index]
            render_items()

        add_btn.clicked.connect(add_item)
        edit_btn.clicked.connect(edit_item)
        remove_btn.clicked.connect(remove_item)
        render_items()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "codigo": code_edit.text().strip(),
            "descricao": desc_edit.text().strip(),
            "notas": notes_edit.toPlainText().strip(),
            "itens": [dict(item.get("line") or {}) for item in items if isinstance(item.get("line"), dict)],
        }

    def _manage_assembly_models(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Modelos de conjuntos")
        dialog.resize(940, 620)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(0, 6)
        table.setHorizontalHeaderLabels(["Codigo", "Descricao", "Itens", "Pecas", "Produtos", "Total base"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(1,), contents=(0, 2, 3, 4, 5))
        layout.addWidget(table, 1)

        def current_code() -> str:
            current = table.currentItem()
            if current is None:
                return ""
            row_item = table.item(current.row(), 0)
            return str(row_item.text() or "").strip() if row_item is not None else ""

        def refresh_models() -> None:
            rows = list(self.backend.assembly_model_rows() or [])
            _fill_table(
                table,
                [
                    [
                        row.get("codigo", "-"),
                        row.get("descricao", "-"),
                        row.get("itens", 0),
                        row.get("pecas", 0),
                        row.get("produtos", 0),
                        _fmt_eur(float(row.get("total_base", 0) or 0)),
                    ]
                    for row in rows
                ],
                align_center_from=2,
            )
            if table.rowCount() > 0:
                table.selectRow(0)

        actions = QHBoxLayout()
        new_btn = QPushButton("Novo")
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "danger")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        actions.addWidget(new_btn)
        actions.addWidget(edit_btn)
        actions.addWidget(remove_btn)
        actions.addStretch(1)
        actions.addWidget(close_btn)
        layout.addLayout(actions)

        def create_model() -> None:
            payload = self._assembly_model_editor_dialog()
            if payload is None:
                return
            try:
                self.backend.assembly_model_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            refresh_models()

        def edit_model() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um modelo.")
                return
            try:
                detail = self.backend.assembly_model_detail(code)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            payload = self._assembly_model_editor_dialog(detail)
            if payload is None:
                return
            payload["codigo"] = code
            try:
                self.backend.assembly_model_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            refresh_models()

        def remove_model() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um modelo.")
                return
            if QMessageBox.question(dialog, "Conjuntos", f"Remover o modelo {code}?") != QMessageBox.Yes:
                return
            try:
                self.backend.assembly_model_remove(code)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            refresh_models()

        new_btn.clicked.connect(create_model)
        edit_btn.clicked.connect(edit_model)
        remove_btn.clicked.connect(remove_model)
        close_btn.clicked.connect(dialog.reject)
        refresh_models()
        dialog.exec()

    def _manage_saved_conjuntos(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Conjuntos guardados")
        dialog.resize(1080, 640)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Os conjuntos guardados funcionam como produto montado: materiais, mao de obra, consumiveis e produtos. "
            "Daqui podes criar, editar, duplicar, remover e aplicar ao orçamento atual."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        table = QTableWidget(0, 9)
        table.setHorizontalHeaderLabels(["Codigo", "Param.", "Descricao", "Itens", "Ligados", "Template", "Margem", "Custo atual", "Final atual"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6, 7, 8))
        layout.addWidget(table, 1)

        def current_code() -> str:
            current = table.currentItem()
            if current is None:
                return ""
            row_item = table.item(current.row(), 0)
            return str(row_item.text() or "").strip() if row_item is not None else ""

        def refresh_rows(select_code: str = "") -> None:
            rows = list(self.backend.conjunto_rows() or [])
            _fill_table(
                table,
                [
                    [
                        row.get("codigo", "-"),
                        row.get("param_codigo", "-"),
                        row.get("descricao", "-"),
                        row.get("itens", 0),
                        f"{int(row.get('itens_ligados', 0) or 0)}/{int(row.get('itens', 0) or 0)}",
                        "Sim" if bool(row.get("template", False)) else "Nao",
                        f"{float(row.get('margem_perc', 0) or 0):.2f} %",
                        _fmt_eur(float(row.get("total_custo", 0) or 0)),
                        _fmt_eur(float(row.get("total_final", 0) or 0)),
                    ]
                    for row in rows
                ],
                align_center_from=3,
            )
            if table.rowCount() <= 0:
                return
            wanted = str(select_code or "").strip()
            row_index = 0
            if wanted:
                for index, row in enumerate(rows):
                    if str(row.get("codigo", "") or "").strip() == wanted:
                        row_index = index
                        break
            table.selectRow(row_index)

        actions = QHBoxLayout()
        new_btn = QPushButton("Novo conjunto")
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        duplicate_btn = QPushButton("Duplicar")
        duplicate_btn.setProperty("variant", "secondary")
        apply_btn = QPushButton("Adicionar ao orçamento")
        apply_btn.setProperty("variant", "secondary")
        preview_btn = QPushButton("Previsualizar PDF")
        preview_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "danger")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        for button in (new_btn, edit_btn, duplicate_btn, apply_btn, preview_btn, remove_btn):
            actions.addWidget(button)
        actions.addStretch(1)
        actions.addWidget(close_btn)
        layout.addLayout(actions)

        def create_conjunto() -> None:
            payload = self._calculated_assembly_builder_dialog()
            if not payload:
                return
            refresh_rows(str(payload.get("assembly_code", "") or "").strip())

        def edit_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                detail = dict(self.backend.conjunto_detail(code) or {})
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            payload = self._calculated_assembly_builder_dialog(detail)
            if not payload:
                return
            refresh_rows(str(payload.get("assembly_code", code) or code).strip())

        def duplicate_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                detail = dict(self.backend.conjunto_detail(code) or {})
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            detail["codigo"] = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            detail.pop("param_codigo", None)
            detail["descricao"] = f"{str(detail.get('descricao', code) or code).strip()} (Copia)"
            detail.pop("created_at", None)
            detail.pop("updated_at", None)
            payload = self._calculated_assembly_builder_dialog(detail)
            if not payload:
                return
            refresh_rows(str(payload.get("assembly_code", "") or "").strip())

        def apply_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            qty, ok = QInputDialog.getDouble(dialog, "Adicionar conjunto", "Quantidade de conjuntos", 1.0, 0.01, 1000000.0, 2)
            if not ok:
                return
            try:
                self.line_rows.extend(self.backend.conjunto_expand(code, qty))
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            self._render_quote_lines()
            QMessageBox.information(dialog, "Conjuntos", f"O conjunto {code} foi adicionado ao orçamento.")

        def preview_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                path = self.backend.conjunto_open_sheet_pdf(code)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            QMessageBox.information(dialog, "Conjuntos", f"Ficha PDF aberta:\n{path}")

        def remove_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            if QMessageBox.question(dialog, "Conjuntos", f"Remover o conjunto {code}?") != QMessageBox.Yes:
                return
            try:
                self.backend.conjunto_remove(code)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            refresh_rows()

        new_btn.clicked.connect(create_conjunto)
        edit_btn.clicked.connect(edit_conjunto)
        duplicate_btn.clicked.connect(duplicate_conjunto)
        apply_btn.clicked.connect(apply_conjunto)
        preview_btn.clicked.connect(preview_conjunto)
        remove_btn.clicked.connect(remove_conjunto)
        close_btn.clicked.connect(dialog.reject)
        refresh_rows()
        dialog.exec()

    def _save_selected_lines_as_group(self) -> None:
        selected_indexes = self._selected_line_indexes()
        if not selected_indexes:
            QMessageBox.warning(self, "Conjuntos", "Seleciona pelo menos uma linha para guardar no conjunto/modelo.")
            return

        selected_rows = [dict(self.line_rows[index] or {}) for index in selected_indexes]
        selected_codes = {
            str(row.get("conjunto_codigo", "") or "").strip()
            for row in selected_rows
            if str(row.get("conjunto_codigo", "") or "").strip()
        }
        selected_names = {
            str(row.get("conjunto_nome", "") or "").strip()
            for row in selected_rows
            if str(row.get("conjunto_nome", "") or "").strip()
        }
        prefill_code = next(iter(selected_codes), "") if len(selected_codes) == 1 else ""
        prefill_name = next(iter(selected_names), "") if len(selected_names) == 1 else ""
        if not prefill_name:
            prefill_name = str(self.note_cliente_edit.text() or self.current_number or "").strip()
        if not prefill_code:
            prefill_code = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"

        dialog = QDialog(self)
        dialog.setWindowTitle("Guardar linhas como conjunto/modelo")
        dialog.resize(760, 520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QLabel(
            "Seleciona como queres guardar as linhas atuais: num conjunto montado, num modelo reutilizável, "
            "ou nos dois ao mesmo tempo. Se escolheres sobrepor, o sistema junta as linhas selecionadas às linhas "
            "desse conjunto que já estejam no orçamento atual, para não perder o que já montaste."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        form = QFormLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        destination_combo = QComboBox()
        destination_combo.addItem("Conjunto", "conjunto")
        destination_combo.addItem("Modelo", "modelo")
        destination_combo.addItem("Conjunto + Modelo", "both")

        save_mode_combo = QComboBox()
        save_mode_combo.addItem("Novo registo", "new")
        save_mode_combo.addItem("Sobrepor existente", "overwrite")

        existing_combo = QComboBox()
        existing_combo.setEnabled(False)

        code_edit = QLineEdit(prefill_code)
        name_edit = QLineEdit(prefill_name)
        notes_edit = QTextEdit()
        notes_edit.setMaximumHeight(100)
        notes_edit.setPlainText(
            f"Guardado a partir do orçamento {str(self.current_number or '').strip()} com {len(selected_rows)} linha(s) selecionada(s)."
        )
        template_check = QCheckBox("Marcar também como template reutilizável")

        info_label = QLabel("")
        info_label.setWordWrap(True)
        info_label.setProperty("role", "muted")

        form.addRow("Guardar em", destination_combo)
        form.addRow("Modo", save_mode_combo)
        form.addRow("Registo existente", existing_combo)
        form.addRow("Código", code_edit)
        form.addRow("Descrição", name_edit)
        form.addRow("Notas", notes_edit)
        form.addRow("", template_check)
        layout.addLayout(form)
        layout.addWidget(info_label)

        preview_table = QTableWidget(0, 5)
        preview_table.setHorizontalHeaderLabels(["Tipo", "Ref. Ext.", "Descrição", "Qtd", "Total"])
        preview_table.verticalHeader().setVisible(False)
        preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        preview_table.setSelectionMode(QTableWidget.NoSelection)
        preview_table.setAlternatingRowColors(True)
        _configure_table(preview_table, stretch=(2,), contents=(0, 1, 3, 4))
        _fill_table(
            preview_table,
            [
                [
                    self._quote_line_type_label(row),
                    str(row.get("ref_externa", "") or "-").strip() or "-",
                    str(row.get("descricao", "") or "-").strip() or "-",
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    _fmt_eur(float(row.get("total", float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0)) or 0)),
                ]
                for row in selected_rows
            ],
            align_center_from=3,
        )
        preview_table.setMinimumHeight(max(170, min(290, _table_visible_height(preview_table, min(max(len(selected_rows), 3), 6), extra=14))))
        layout.addWidget(preview_table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        ok_btn = buttons.button(QDialogButtonBox.Ok)

        def _list_destination_rows(destination: str) -> list[dict]:
            dest = str(destination or "").strip().lower()
            if dest == "conjunto":
                return [dict(row or {}) for row in list(self.backend.conjunto_rows() or [])]
            if dest == "modelo":
                return [dict(row or {}) for row in list(self.backend.assembly_model_rows() or [])]
            combined: dict[str, dict] = {}
            for source_row in list(self.backend.conjunto_rows() or []):
                row = dict(source_row or {})
                code = str(row.get("codigo", "") or "").strip()
                if code:
                    combined[code] = {
                        "codigo": code,
                        "descricao": str(row.get("descricao", "") or "").strip(),
                        "source": "both",
                    }
            for source_row in list(self.backend.assembly_model_rows() or []):
                row = dict(source_row or {})
                code = str(row.get("codigo", "") or "").strip()
                if not code:
                    continue
                combined.setdefault(
                    code,
                    {
                        "codigo": code,
                        "descricao": str(row.get("descricao", "") or "").strip(),
                        "source": "both",
                    },
                )
            return [combined[key] for key in sorted(combined)]

        def _existing_codes(destination: str) -> set[str]:
            return {
                str(row.get("codigo", "") or "").strip()
                for row in _list_destination_rows(destination)
                if str(row.get("codigo", "") or "").strip()
            }

        def _target_detail_lines(destination: str, code: str) -> list[dict]:
            target_code = str(code or "").strip()
            if not target_code:
                return []
            if destination in {"conjunto", "both"}:
                try:
                    detail = dict(self.backend.conjunto_detail(target_code) or {})
                except Exception:
                    detail = {}
                rows = [dict(item or {}) for item in list(detail.get("itens", []) or []) if isinstance(item, dict)]
                if rows:
                    return rows
            if destination in {"modelo", "both"}:
                try:
                    detail = dict(self.backend.assembly_model_detail(target_code) or {})
                except Exception:
                    detail = {}
                return [dict(item or {}) for item in list(detail.get("itens", []) or []) if isinstance(item, dict)]
            return []

        def _line_signature(row: dict) -> tuple:
            return (
                str(row.get("tipo_item", "") or "").strip(),
                str(row.get("produto_codigo", "") or "").strip(),
                str(row.get("ref_externa", "") or "").strip(),
                str(row.get("descricao", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                str(row.get("operacao", "") or "").strip(),
                round(float(row.get("qtd", 0) or 0), 4),
                round(float(row.get("preco_unit", 0) or 0), 4),
                str(row.get("desenho", "") or "").strip(),
            )

        def _sanitize_group_line(raw_row: dict) -> dict:
            row = dict(raw_row or {})
            row.pop("conjunto_codigo", None)
            row.pop("conjunto_nome", None)
            row.pop("grupo_uuid", None)
            row.pop("total", None)
            if self.backend.desktop_main.orc_line_is_product(row) and not str(row.get("ref_externa", "") or "").strip():
                row["ref_externa"] = str(row.get("produto_codigo", "") or "").strip()
            operation_norm = self.backend.desktop_main.norm_text(str(row.get("operacao", "") or ""))
            if self.backend.desktop_main.orc_line_is_piece(row) and "laser" in operation_norm:
                row["source_quote_number"] = str(self.current_number or "").strip()
                row["source_ref_externa"] = str(row.get("ref_externa", "") or "").strip()
                row["pricing_source"] = "quote_laser"
            return row

        def _compose_rows_for_save(destination: str, target_code: str, overwrite: bool) -> list[dict]:
            rows = []
            if overwrite:
                quote_rows = [
                    dict(row or {})
                    for row in list(self.line_rows)
                    if str((row or {}).get("conjunto_codigo", "") or "").strip() == target_code
                ]
                if quote_rows:
                    rows.extend(quote_rows)
                else:
                    rows.extend(_target_detail_lines(destination, target_code))
            rows.extend(selected_rows)
            merged: list[dict] = []
            seen: set[tuple] = set()
            for raw_row in rows:
                clean_row = _sanitize_group_line(raw_row)
                signature = _line_signature(clean_row)
                if signature in seen:
                    continue
                seen.add(signature)
                merged.append(clean_row)
            return merged

        def _refresh_overwrite_options() -> None:
            destination = str(destination_combo.currentData() or "conjunto").strip()
            overwrite = str(save_mode_combo.currentData() or "new").strip() == "overwrite"
            rows = _list_destination_rows(destination)
            existing_combo.blockSignals(True)
            existing_combo.clear()
            for row in rows:
                code = str(row.get("codigo", "") or "").strip()
                desc = str(row.get("descricao", "") or "").strip()
                existing_combo.addItem(f"{code} - {desc}", code)
            existing_combo.blockSignals(False)
            existing_combo.setVisible(overwrite)
            existing_combo.setEnabled(overwrite and bool(rows))
            code_edit.setReadOnly(overwrite)
            if overwrite and rows:
                wanted_code = str(prefill_code or "").strip()
                wanted_index = 0
                if wanted_code:
                    for idx, row in enumerate(rows):
                        if str(row.get("codigo", "") or "").strip() == wanted_code:
                            wanted_index = idx
                            break
                existing_combo.setCurrentIndex(wanted_index)
                code_edit.setText(str(existing_combo.currentData() or "").strip())
                if not name_edit.text().strip():
                    row = rows[wanted_index]
                    name_edit.setText(str(row.get("descricao", "") or "").strip())
            elif overwrite:
                code_edit.clear()
            elif not code_edit.text().strip():
                code_edit.setText(prefill_code)
            if overwrite and not rows:
                info_label.setText("Nao existem registos desse tipo para sobrepor. Escolhe 'Novo registo' ou cria primeiro um conjunto/modelo.")
                if isinstance(ok_btn, QPushButton):
                    ok_btn.setEnabled(False)
                return
            info_label.setText(
                "As linhas selecionadas vao ficar ligadas ao conjunto/modelo escolhido no orçamento atual."
                if overwrite
                else "Vai ser criado um novo conjunto/modelo a partir das linhas selecionadas."
            )
            if isinstance(ok_btn, QPushButton):
                ok_btn.setEnabled(True)

        def _sync_selected_existing() -> None:
            code = str(existing_combo.currentData() or "").strip()
            if not code:
                return
            destination = str(destination_combo.currentData() or "conjunto").strip()
            rows = _list_destination_rows(destination)
            detail_row = next((row for row in rows if str(row.get("codigo", "") or "").strip() == code), {})
            code_edit.setText(code)
            if detail_row:
                name_edit.setText(str(detail_row.get("descricao", "") or "").strip())

        destination_combo.currentIndexChanged.connect(_refresh_overwrite_options)
        save_mode_combo.currentIndexChanged.connect(_refresh_overwrite_options)
        existing_combo.currentIndexChanged.connect(_sync_selected_existing)
        _refresh_overwrite_options()

        if dialog.exec() != QDialog.Accepted:
            return

        destination = str(destination_combo.currentData() or "conjunto").strip()
        overwrite = str(save_mode_combo.currentData() or "new").strip() == "overwrite"
        code = (
            str(existing_combo.currentData() or "").strip()
            if overwrite
            else str(code_edit.text() or "").strip()
        )
        if not code:
            QMessageBox.warning(self, "Conjuntos", "Indica um código para o conjunto/modelo.")
            return
        description = str(name_edit.text() or "").strip() or code
        notes = str(notes_edit.toPlainText() or "").strip()
        if not overwrite and code in _existing_codes(destination):
            QMessageBox.warning(
                self,
                "Conjuntos",
                f"Ja existe um registo com o codigo {code}. Se queres atualizar esse registo, usa o modo 'Sobrepor existente'.",
            )
            return

        lines_for_save = _compose_rows_for_save(destination, code, overwrite)
        if not lines_for_save:
            QMessageBox.warning(self, "Conjuntos", "Nao ha linhas validas para guardar no conjunto/modelo.")
            return

        total_value = round(
            sum(
                float(row.get("total", float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0)) or 0)
                for row in lines_for_save
            ),
            2,
        )
        payload = {
            "codigo": code,
            "descricao": description,
            "notas": notes,
            "itens": lines_for_save,
            "template": bool(template_check.isChecked()),
            "origem": "orcamento_linhas_guardadas",
            "margem_perc": 0.0,
            "total_custo": total_value,
            "total_final": total_value,
        }

        try:
            if destination in {"conjunto", "both"}:
                self.backend.conjunto_save(payload)
            if destination in {"modelo", "both"}:
                self.backend.assembly_model_save(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return

        existing_group_rows = [
            dict(row or {})
            for row in list(self.line_rows)
            if str((row or {}).get("conjunto_codigo", "") or "").strip() == code
            and str((row or {}).get("grupo_uuid", "") or "").strip()
        ]
        group_uuid = (
            str(existing_group_rows[0].get("grupo_uuid", "") or "").strip()
            if existing_group_rows
            else f"{code}-01"
        )
        affected_indexes = set(selected_indexes)
        if overwrite:
            affected_indexes.update(
                index
                for index, row in enumerate(self.line_rows)
                if str((row or {}).get("conjunto_codigo", "") or "").strip() == code
            )
        for index in sorted(affected_indexes):
            if not (0 <= index < len(self.line_rows)):
                continue
            row = dict(self.line_rows[index] or {})
            row["conjunto_codigo"] = code
            row["conjunto_nome"] = description
            row["grupo_uuid"] = group_uuid
            self.line_rows[index] = row

        self._render_quote_lines()
        if selected_indexes:
            self._select_quote_line_source_index(selected_indexes[0])
        QMessageBox.information(
            self,
            "Conjuntos",
            (
                f"As linhas selecionadas foram guardadas em {description} ({code})."
                if not overwrite
                else f"O registo {code} foi atualizado e as linhas ficaram ligadas ao conjunto/modelo."
            ),
        )

    def _add_assembly_model(self) -> None:
        rows = list(self.backend.conjunto_rows() or [])
        expand_fn = getattr(self.backend, "conjunto_expand", None)
        empty_message = "Ainda nao existem conjuntos guardados. Cria primeiro um conjunto calculado."
        if not rows:
            rows = list(self.backend.assembly_model_rows() or [])
            expand_fn = getattr(self.backend, "assembly_model_expand", None)
            empty_message = "Ainda nao existem modelos. Cria primeiro um modelo de conjunto."
        if not rows:
            QMessageBox.information(self, "Conjuntos", empty_message)
            self._manage_saved_conjuntos()
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Adicionar conjunto guardado")
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Seleciona um conjunto guardado para o expandir no orçamento atual. "
            "Os conjuntos calculados usam uma lógica separada das linhas DXF/DWG."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)
        form = QFormLayout()
        combo = QComboBox()
        for row in rows:
            combo.addItem(f"{row.get('codigo', '')} - {row.get('descricao', '')}", str(row.get("codigo", "") or "").strip())
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(1.0)
        form.addRow("Modelo", combo)
        form.addRow("Quantidade conjuntos", qty_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            if not callable(expand_fn):
                raise ValueError("Funcao de expansao do conjunto indisponivel.")
            self.line_rows.extend(expand_fn(str(combo.currentData() or "").strip(), qty_spin.value()))
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return
        self._render_quote_lines()

    def _configure_laser_profiles(self) -> None:
        dialog = LaserSettingsDialog(self.backend, self)
        dialog.exec()

    def _configure_operation_profiles(self) -> None:
        _open_operation_cost_profiles_dialog(self, self.backend)

    def _resolve_laser_edit_source(self, row: dict) -> dict:
        source = dict(row or {})
        drawing = str(source.get("desenho", "") or "").strip()
        if Path(drawing).suffix.casefold() in {".dxf", ".dwg"}:
            return source
        finder = getattr(self.backend, "_conjunto_find_quote_source", None)
        if not callable(finder):
            return source
        try:
            quote_line, quote_number = finder(source, str(source.get("conjunto_codigo", "") or "").strip())
        except Exception:
            return source
        if not isinstance(quote_line, dict):
            return source
        resolved = dict(source)
        for key, value in dict(quote_line or {}).items():
            if key not in resolved or resolved.get(key) in (None, "", [], {}):
                resolved[key] = value
        if quote_number and not str(resolved.get("source_quote_number", "") or "").strip():
            resolved["source_quote_number"] = str(quote_number).strip()
        return resolved

    def _quote_line_is_laser_2d(self, row: dict | None) -> bool:
        line = dict(row or {})
        if not self.backend.desktop_main.orc_line_is_piece(line):
            return False
        operation = str(line.get("operacao", "") or "").casefold()
        has_laser = bool(line.get("laser_base_active", False) or "laser" in operation)
        if not has_laser:
            return False
        if str(line.get("laser_source_mode", "") or "").strip().casefold() == "batch":
            return True
        if str(line.get("laser_batch_id", "") or "").strip():
            return True
        drawing = str(line.get("desenho", "") or "").strip()
        suffix = Path(drawing).suffix.casefold()
        return suffix in {".dxf", ".dwg"}

    def _edit_laser_batch_lines(
        self,
        rows: list[dict],
        *,
        parent: QWidget | None = None,
    ) -> list[dict] | None:
        source_lines = [self._resolve_laser_edit_source(dict(row or {})) for row in rows if isinstance(row, dict)]
        if not source_lines:
            return None
        batch_dialog = LaserBatchQuoteDialog(
            self.backend,
            parent if isinstance(parent, QWidget) else self,
            default_machine=(
                str(source_lines[0].get("laser_machine", source_lines[0].get("machine", "")) or "").strip()
                or self.workcenter_combo.currentText().strip()
            ),
            initial_lines=source_lines,
        )
        if batch_dialog.exec() != QDialog.Accepted:
            return None
        result = dict(batch_dialog.result_payload() or {})
        edited = [dict(row or {}) for row in list(result.get("lines", []) or []) if isinstance(row, dict) and row]
        if not edited:
            return None

        def identity(line: dict) -> tuple[str, str]:
            return (
                str(line.get("desenho", "") or "").strip().replace("\\", "/").casefold(),
                str(line.get("ref_externa", "") or "").strip().casefold(),
            )

        by_identity = {identity(source): source for source in source_lines}
        batch_id = str(batch_dialog.batch_id or "").strip()
        merged_lines: list[dict] = []
        preserved_keys = (
            "conjunto_codigo",
            "conjunto_nome",
            "conjunto_param_codigo",
            "grupo_uuid",
            "ficha_tecnica",
            "source_quote_number",
            "source_ref_externa",
            "pricing_source",
            "pricing_source_ref",
        )
        for line in edited:
            source = by_identity.get(identity(line), {})
            merged = {**source, **line}
            for key in preserved_keys:
                if key not in line and key in source:
                    merged[key] = source[key]
            merged["laser_source_mode"] = "batch"
            merged["laser_batch_id"] = batch_id
            merged_lines.append(merged)
        return merged_lines

    def _edit_laser_2d_line(self, row: dict, *, parent: QWidget | None = None) -> dict | None:
        source = self._resolve_laser_edit_source(dict(row or {}))
        dialog = LaserQuoteDialog(
            self.backend,
            parent if isinstance(parent, QWidget) else self,
            default_machine=(
                str(source.get("laser_machine", source.get("machine", "")) or "").strip()
                or self.workcenter_combo.currentText().strip()
            ),
            initial_line=source,
        )
        if dialog.exec() != QDialog.Accepted:
            return None
        result = dict(dialog.result_payload() or {})
        laser_line = dict(result.get("line", {}) or {})
        if not laser_line:
            return None

        merged = {**source, **laser_line}
        try:
            source_operations = list(
                self.backend.quote_parse_operacoes_lista(
                    source.get("operacoes_lista", source.get("operacao", ""))
                )
                or []
            )
            laser_operations = list(
                self.backend.quote_parse_operacoes_lista(laser_line.get("operacao", "Corte Laser"))
                or []
            )
        except Exception:
            source_operations = [part.strip() for part in str(source.get("operacao", "") or "").split("+") if part.strip()]
            laser_operations = [part.strip() for part in str(laser_line.get("operacao", "Corte Laser") or "").split("+") if part.strip()]

        def is_laser_component(name: object) -> bool:
            normalized = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode().casefold()
            return any(token in normalized for token in ("laser", "marcacao", "defilm"))

        extra_operations = [name for name in source_operations if not is_laser_component(name)]
        combined_operations: list[str] = []
        for name in [*laser_operations, *extra_operations]:
            clean = str(name or "").strip()
            if clean and clean.casefold() not in {item.casefold() for item in combined_operations}:
                combined_operations.append(clean)

        extra_times = {
            str(key): float(value or 0)
            for key, value in dict(source.get("tempos_operacao", {}) or {}).items()
            if not is_laser_component(key)
        }
        extra_costs = {
            str(key): float(value or 0)
            for key, value in dict(source.get("custos_operacao", {}) or {}).items()
            if not is_laser_component(key)
        }
        laser_time = float(laser_line.get("tempo_peca_min", 0) or 0)
        laser_price = float(laser_line.get("preco_unit", 0) or 0)
        merged["operacao"] = " + ".join(combined_operations or ["Corte Laser"])
        merged["operacoes_lista"] = combined_operations or ["Corte Laser"]
        merged["tempo_peca_min"] = round(laser_time + sum(extra_times.values()), 4)
        merged["preco_unit"] = round(laser_price + sum(extra_costs.values()), 4)
        merged["total"] = round(float(merged.get("qtd", 0) or 0) * float(merged["preco_unit"]), 2)
        merged["laser_base_active"] = True
        merged["laser_base_tempo_unit"] = round(laser_time, 4)
        merged["laser_base_preco_unit"] = round(laser_price, 4)
        merged["tempos_operacao"] = extra_times
        merged["custos_operacao"] = extra_costs
        for key in ("desenho_pdf", "desenhos_pdf", "ficheiros", "conjunto_codigo", "conjunto_nome", "grupo_uuid"):
            if key not in laser_line and key in source:
                merged[key] = source[key]
        return merged

    def _add_laser_line(self) -> None:
        dialog = LaserQuoteDialog(
            self.backend,
            self,
            default_machine=self.workcenter_combo.currentText().strip(),
        )
        if dialog.exec() != QDialog.Accepted:
            return
        result = dict(dialog.result_payload() or {})
        lines = [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})]
        line = dict(result.get("line", {}) or {})
        analysis = dict(result.get("analysis", {}) or {})
        if not lines and line:
            lines = [line]
        if not lines:
            QMessageBox.warning(self, "Peca Unit. DXF/DWG", "Nao foi possivel gerar a linha de orcamento.")
            return
        self.line_rows.extend(lines)
        machine_name = str(dict(analysis.get("machine", {}) or {}).get("name", "") or "").strip()
        if machine_name and not self.workcenter_combo.currentText().strip():
            self.workcenter_combo.setCurrentText(machine_name)
        self._render_quote_lines()

    def _add_laser_batch_lines(self) -> None:
        dialog = LaserBatchQuoteDialog(
            self.backend,
            self,
            default_machine=self.workcenter_combo.currentText().strip(),
        )
        if dialog.exec() != QDialog.Accepted:
            return
        result = dict(dialog.result_payload() or {})
        lines = [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})]
        analysis = dict(result.get("analysis", {}) or {})
        if not lines:
            QMessageBox.warning(self, "Lote DXF/DWG", "Nao foi possivel gerar linhas de orcamento para o lote.")
            return
        self.line_rows.extend(lines)
        machine_name = str(dict(analysis.get("machine", {}) or {}).get("name", "") or "").strip()
        if machine_name and not self.workcenter_combo.currentText().strip():
            self.workcenter_combo.setCurrentText(machine_name)
        self._render_quote_lines()

    def _open_laser_nesting(self) -> None:
        selected_indexes = self._selected_line_indexes()
        candidate_rows = [dict(self.line_rows[index] or {}) for index in selected_indexes] if selected_indexes else [dict(row or {}) for row in self.line_rows]
        laser_rows = [
            row
            for row in candidate_rows
            if str(row.get("desenho", "") or "").strip()
            and "corte laser" in str(row.get("operacao", "") or "").strip().lower()
        ]
        if not laser_rows:
            QMessageBox.information(self, "Nesting Laser", "Seleciona primeiro linhas laser com desenho associado.")
            return
        dialog = LaserNestingDialog(self.backend, laser_rows, self, quote_number=self.current_number)
        dialog.exec()

    def _open_profile_step_igs_quote_builder(
        self,
        _checked: bool = False,
        *,
        return_lines: bool = False,
        parent: QWidget | None = None,
    ) -> list[dict] | None:
        dialog = QDialog(parent if isinstance(parent, QWidget) else self)
        dialog.setWindowTitle("Corte Laser STEP/IGS")
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)
        dialog.setSizeGripEnabled(True)
        dialog.setMinimumSize(860, 520)
        try:
            screen = dialog.screen() or QApplication.primaryScreen()
            available = screen.availableGeometry() if screen is not None else None
            if available is not None:
                width = min(1180, max(860, available.width() - 80))
                height = min(860, max(520, available.height() - 96))
                dialog.resize(width, height)
                dialog.move(
                    available.x() + max(0, (available.width() - width) // 2),
                    available.y() + max(0, (available.height() - height) // 2),
                )
        except Exception:
            dialog.resize(1120, 780)
        outer_layout = QVBoxLayout(dialog)
        outer_layout.setContentsMargins(14, 12, 14, 12)
        outer_layout.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(10)

        title = QLabel("Corte laser de perfis / tubos / cantoneiras")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        intro = QLabel(
            "Fluxo separado do nesting plano. Aqui orçamentas apenas cortes, furos e rasgos a partir de ficheiros STEP/IGS, "
            "sem assumir o custo do material base do perfil."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(intro)
        auto_hint = QLabel(
            "Leitura automatica: Eventos = furos + rasgos + outros cortes internos cobrados. "
            "Cortes terminais do perfil ficam separados e ignorados por defeito; o preco usa os metros reais, "
            "a espessura e a tabela da maquina selecionada."
        )
        auto_hint.setWordWrap(True)
        auto_hint.setProperty("role", "muted")
        layout.addWidget(auto_hint)

        laser_settings = dict(self.backend.laser_quote_settings() or {})

        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)
        add_files_btn = QPushButton("Adicionar STEP/IGS")
        add_files_btn.setProperty("variant", "secondary")
        remove_file_btn = QPushButton("Remover selecionado")
        remove_file_btn.setProperty("variant", "secondary")
        configure_btn = QPushButton("Perfis laser")
        configure_btn.setProperty("variant", "secondary")
        toolbar.addWidget(add_files_btn)
        toolbar.addWidget(remove_file_btn)
        toolbar.addWidget(configure_btn)
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        files_table = QTableWidget(0, 10)
        files_table.setHorizontalHeaderLabels(["Ficheiro", "Tipo", "Secao", "Qtd", "Eventos", "Furos", "Rasgos", "m corte", "Comp. m", "kg/m"])
        files_table.verticalHeader().setVisible(False)
        files_table.verticalHeader().setDefaultSectionSize(42)
        files_table.verticalHeader().setMinimumSectionSize(38)
        files_table.setSelectionBehavior(QTableWidget.SelectRows)
        files_table.setEditTriggers(QTableWidget.NoEditTriggers)
        files_table.setAlternatingRowColors(True)
        files_table.setMinimumHeight(260)
        files_table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QTableWidget::item { padding: 7px 6px; }"
            " QHeaderView::section { padding: 8px 6px; font-weight: 800; }"
        )
        files_table.horizontalHeader().setStretchLastSection(False)
        _set_table_columns(
            files_table,
            [
                (0, "stretch", 0),
                (1, "fixed", 120),
                (2, "fixed", 170),
                (3, "fixed", 72),
                (4, "fixed", 78),
                (5, "fixed", 72),
                (6, "fixed", 78),
                (7, "fixed", 86),
                (8, "fixed", 86),
                (9, "fixed", 86),
            ],
        )
        layout.addWidget(files_table)

        preview_card = CardFrame()
        preview_card.set_tone("default")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(14, 12, 14, 12)
        preview_layout.setSpacing(8)
        preview_title = QLabel("Preview do STEP/IGS")
        preview_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        preview_image_label = QLabel("Seleciona um ficheiro para gerar preview.")
        preview_image_label.setAlignment(Qt.AlignCenter)
        preview_image_label.setMinimumHeight(240)
        preview_image_label.setStyleSheet("border: 1px solid #d0d5dd; border-radius: 10px; background: #f8fafc; color: #475467;")
        preview_info_label = QLabel("O preview usa FreeCAD quando estiver instalado neste posto.")
        preview_info_label.setWordWrap(True)
        preview_info_label.setProperty("role", "muted")
        preview_layout.addWidget(preview_title)
        preview_layout.addWidget(preview_image_label)
        preview_layout.addWidget(preview_info_label)
        layout.addWidget(preview_card)

        pricing_card = CardFrame()
        pricing_card.set_tone("default")
        pricing_layout = QGridLayout(pricing_card)
        pricing_layout.setContentsMargins(14, 12, 14, 12)
        pricing_layout.setHorizontalSpacing(12)
        pricing_layout.setVerticalSpacing(8)

        machine_combo = QComboBox()
        commercial_combo = QComboBox()
        material_combo = QComboBox()
        subtype_combo = QComboBox()
        subtype_combo.setEditable(True)
        gas_combo = QComboBox()
        thickness_spin = QDoubleSpinBox()
        material_price_spin = QDoubleSpinBox()
        material_price_spin.setRange(0.0, 1000000.0)
        material_price_spin.setDecimals(4)
        material_price_spin.setSingleStep(0.1)
        material_price_unit_combo = QComboBox()
        material_price_unit_combo.addItems(["EUR/kg", "EUR/m", "EUR/ton"])
        thickness_spin.setRange(0.1, 200.0)
        thickness_spin.setDecimals(2)
        thickness_spin.setSingleStep(0.5)
        thickness_spin.setValue(3.0)
        customer_material_check = QCheckBox("Perfil/tubo fornecido pelo cliente")
        customer_material_check.setChecked(True)
        customer_material_check.setToolTip("No fluxo STEP/IGS o material base do perfil fica fora do calculo por defeito.")
        total_label = QLabel("Total estimado: 0,00 EUR")
        total_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        status_label = QLabel("Usa as tabelas do laser com contagem de eventos STEP/IGS, sem duplicar furos.")
        status_label.setWordWrap(True)
        status_label.setProperty("role", "muted")
        pricing_layout.addWidget(QLabel("Maquina"), 0, 0)
        pricing_layout.addWidget(machine_combo, 0, 1)
        pricing_layout.addWidget(QLabel("Perfil comercial"), 0, 2)
        pricing_layout.addWidget(commercial_combo, 0, 3)
        pricing_layout.addWidget(QLabel("Familia material"), 1, 0)
        pricing_layout.addWidget(material_combo, 1, 1)
        pricing_layout.addWidget(QLabel("Subtipo / qualidade"), 1, 2)
        pricing_layout.addWidget(subtype_combo, 1, 3)
        pricing_layout.addWidget(QLabel("Gas"), 2, 0)
        pricing_layout.addWidget(gas_combo, 2, 1)
        pricing_layout.addWidget(QLabel("Espessura (mm)"), 2, 2)
        pricing_layout.addWidget(thickness_spin, 2, 3)
        pricing_layout.addWidget(QLabel("Preco material"), 3, 0)
        pricing_layout.addWidget(material_price_spin, 3, 1)
        pricing_layout.addWidget(QLabel("Unid. preco material"), 3, 2)
        pricing_layout.addWidget(material_price_unit_combo, 3, 3)
        pricing_layout.addWidget(customer_material_check, 4, 0, 1, 2)
        pricing_layout.addWidget(total_label, 4, 2, 1, 2)
        pricing_layout.addWidget(status_label, 5, 0, 1, 4)
        layout.addWidget(pricing_card)
        scroll.setWidget(scroll_content)
        outer_layout.addWidget(scroll, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
        buttons.button(QDialogButtonBox.Cancel).setText("Fechar")
        outer_layout.addWidget(buttons)

        result_lines: list[dict] = []
        families = ["Perfil", "Tubo", "Cantoneira", "Barra"]
        material_price_internal_update = False
        material_price_user_touched = False
        material_price_auto_filled = False

        def _infer_family(path_txt: str) -> str:
            probe = Path(path_txt).stem.lower()
            if "tubo" in probe:
                return "Tubo"
            if "cant" in probe:
                return "Cantoneira"
            if "barra" in probe or "chata" in probe:
                return "Barra"
            if any(token in probe for token in ("ipe", "ipn", "hea", "heb", "upn", "rhs", "shs", "perfil")):
                return "Perfil"
            return "Perfil"

        def _family_from_analysis(cad_analysis: dict[str, Any], fallback: str) -> str:
            family = str(cad_analysis.get("family_guess", "") or "").strip()
            if family in {"Tubo", "Cantoneira", "Barra", "Perfil"}:
                return family
            return fallback

        def _infer_section(path_txt: str) -> str:
            stem = Path(path_txt).stem.upper().replace("_", " ").replace("-", " ")
            profile_match = re.search(r"\b(IPE|IPN|HEA|HEB|UPN|RHS|SHS)\s*(\d+)\b", stem)
            if profile_match:
                return f"{profile_match.group(1)} {profile_match.group(2)}"
            size_match = re.search(r"(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:[.,]\d+)?)?)", stem)
            if size_match:
                return re.sub(r"\s*[Xx]\s*", "x", size_match.group(1)).replace(",", ".")
            return ""

        def _infer_profile_length_m(path_txt: str) -> float:
            stem = Path(path_txt).stem.lower().replace("_", " ").replace("-", " ")
            unit_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(mm|m)\b", stem)
            if unit_match:
                value = float(unit_match.group(1).replace(",", "."))
                return round(value / 1000.0 if unit_match.group(2) == "mm" else value, 4)
            numbers = [float(item.replace(",", ".")) for item in re.findall(r"\d+(?:[.,]\d+)?", stem)]
            if len(numbers) == 1 and numbers[0] > 20:
                return round(numbers[0] / 1000.0, 4)
            return 0.0

        def _infer_material_from_profile(path_txt: str, text: str = "") -> tuple[str, str]:
            probe = f"{Path(path_txt).stem} {str(text or '')[:120000]}".upper()
            subtype_patterns = [
                r"\bS235(?:JR)?\b",
                r"\bS275(?:JR)?\b",
                r"\bS355(?:JR|J2\+N|MC|JOW)?\b",
                r"\bS420MC\b",
                r"\bDX5[13]D(?:\+Z)?\b",
                r"\bDD11\b",
                r"\bDC01\b",
                r"\bCORTEN\b",
                r"\bHARDOX(?:\s*4[05]0)?\b",
                r"\bINOX\s*3(?:04|16)L?\b",
                r"\bAISI\s*3(?:04|16)L?\b",
                r"\b1\.4(?:301|307|401|404|016)\b",
            ]
            matched_subtype = ""
            for pattern in subtype_patterns:
                match = re.search(pattern, probe)
                if match:
                    matched_subtype = re.sub(r"\s+", " ", match.group(0).strip())
                    break
            guessed = _laser_guess_material_family(matched_subtype)
            if guessed:
                return _laser_canonical_material_family(guessed) or guessed, matched_subtype
            if re.search(r"\b(INOX|STAINLESS|AISI\s*3(?:04|16)L?|1\.4(?:301|307|401|404|016))\b", probe):
                return "Aco inox", matched_subtype
            if re.search(r"\b(FERRO|ACO|AÇO|S235|S275|S355|S420|CORTEN|HARDOX|DX5[13]D)\b", probe):
                return "Aco carbono", matched_subtype
            return "", matched_subtype

        def _make_int_spin(value: int) -> QDoubleSpinBox:
            spin = QDoubleSpinBox()
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(0)
            spin.setSingleStep(1.0)
            spin.setValue(float(value))
            spin.setMinimumHeight(32)
            spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
            return spin

        def _read_geometry_preview(path_txt: str) -> str:
            try:
                raw = Path(path_txt).read_bytes()
            except Exception:
                return ""
            if not raw:
                return ""
            sample = raw[:1_200_000]
            for encoding in ("utf-8", "latin-1", "cp1252"):
                try:
                    return sample.decode(encoding, errors="ignore")
                except Exception:
                    continue
            return sample.decode("latin-1", errors="ignore")

        def _clear_step_preview(message: str, info: str = "") -> None:
            preview_image_label.clear()
            preview_image_label.setPixmap(QPixmap())
            preview_image_label.setText(message)
            preview_info_label.setText(info or "O preview usa FreeCAD quando estiver instalado neste posto.")

        def _update_step_preview() -> None:
            row_index = _selected_row_index(files_table)
            if row_index < 0:
                _clear_step_preview("Seleciona um ficheiro para gerar preview.")
                return
            file_item = files_table.item(row_index, 0)
            path_txt = str(file_item.data(Qt.UserRole) if isinstance(file_item, QTableWidgetItem) else "").strip()
            if not path_txt:
                _clear_step_preview("Sem ficheiro associado.")
                return
            cached_preview = dict(file_item.data(Qt.UserRole + 2) if isinstance(file_item, QTableWidgetItem) else {} or {})
            if not cached_preview:
                _clear_step_preview("A gerar preview FreeCAD...", Path(path_txt).name)
                QApplication.processEvents()
                try:
                    cached_preview = dict(render_step_preview_image(path_txt) or {})
                except Exception as exc:
                    cached_preview = {"available": False, "note": str(exc)}
                if isinstance(file_item, QTableWidgetItem):
                    file_item.setData(Qt.UserRole + 2, dict(cached_preview))
            image_path = str(cached_preview.get("image_path", "") or "").strip()
            if cached_preview.get("available") and image_path and Path(image_path).exists():
                pixmap = QPixmap(image_path)
                if not pixmap.isNull():
                    preview_image_label.setText("")
                    preview_image_label.setPixmap(
                        pixmap.scaled(760, 260, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    )
                    preview_info_label.setText(
                        f"{Path(path_txt).name} | Preview gerado por {str(cached_preview.get('engine', 'FreeCAD') or 'FreeCAD')}"
                    )
                    return
            note_txt = str(cached_preview.get("note", "") or "").strip()
            _clear_step_preview(
                "Preview indisponivel neste posto.",
                note_txt or f"{Path(path_txt).name} | Instala/configura o FreeCAD para gerar imagem do STEP.",
            )

        def _refresh_machine_and_commercial() -> None:
            _laser_set_combo_values(machine_combo, list(dict(laser_settings.get("machine_profiles", {}) or {}).keys()))
            _laser_set_combo_values(commercial_combo, list(dict(laser_settings.get("commercial_profiles", {}) or {}).keys()))

        def _material_subtype_candidates(material_name: str) -> list[str]:
            extras: list[str] = []
            try:
                presets = dict(self.backend.material_presets() or {})
                family = _laser_guess_material_family(material_name) or str(material_name or "").strip()
                for value in list(presets.get("materiais", []) or []):
                    clean = str(value or "").strip()
                    if not clean or clean == material_name:
                        continue
                    if _laser_guess_material_family(clean) == family and clean not in extras:
                        extras.append(clean)
            except Exception:
                pass
            return _laser_settings_material_subtypes(laser_settings, material_name, extras)

        def _refresh_materials() -> None:
            current_material = _laser_display_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip()
            values = _laser_settings_material_names(laser_settings, machine_combo.currentText().strip())
            _laser_set_combo_values(material_combo, values, current_material if current_material else (values[0] if values else "Ferro"))
            _refresh_subtypes()

        def _refresh_subtypes() -> None:
            current_subtype = subtype_combo.currentText().strip()
            values = _material_subtype_candidates(material_combo.currentText().strip())
            _laser_set_combo_values(subtype_combo, values, current_subtype)
            if current_subtype and current_subtype not in [subtype_combo.itemText(index) for index in range(subtype_combo.count())]:
                subtype_combo.setCurrentText(current_subtype)
            _refresh_gases()

        def _refresh_gases() -> None:
            current_gas = gas_combo.currentText().strip()
            values = _laser_settings_gas_names(laser_settings, machine_combo.currentText().strip(), material_combo.currentText().strip())
            _laser_set_combo_values(gas_combo, values, current_gas if current_gas else (values[0] if values else "Oxigenio"))

        def _configure_profiles() -> None:
            nonlocal laser_settings
            cfg_dialog = LaserSettingsDialog(self.backend, self)
            if cfg_dialog.exec() != QDialog.Accepted:
                return
            current_machine = machine_combo.currentText().strip()
            current_commercial = commercial_combo.currentText().strip()
            current_material = material_combo.currentText().strip()
            current_gas = gas_combo.currentText().strip()
            laser_settings = dict(self.backend.laser_quote_settings() or {})
            _laser_set_combo_values(machine_combo, list(dict(laser_settings.get("machine_profiles", {}) or {}).keys()), current_machine)
            _laser_set_combo_values(commercial_combo, list(dict(laser_settings.get("commercial_profiles", {}) or {}).keys()), current_commercial)
            _refresh_materials()
            if current_material:
                material_combo.setCurrentText(current_material)
                _refresh_subtypes()
            if current_gas:
                gas_combo.setCurrentText(current_gas)
            _recalc_total()

        def _selected_profile_family() -> str:
            row_index = _selected_row_index(files_table)
            if row_index < 0 and files_table.rowCount() > 0:
                row_index = 0
            if row_index < 0:
                return ""
            family_widget = files_table.cellWidget(row_index, 1)
            if isinstance(family_widget, QComboBox):
                return family_widget.currentText().strip()
            return ""

        def _sync_material_unit_for_family() -> None:
            nonlocal material_price_internal_update
            if _selected_profile_family() != "Tubo":
                return
            material_price_internal_update = True
            try:
                material_price_unit_combo.setCurrentText("EUR/m")
            finally:
                material_price_internal_update = False

        def _sync_material_cost_controls() -> None:
            enabled = not bool(customer_material_check.isChecked())
            material_price_spin.setEnabled(enabled)
            material_price_unit_combo.setEnabled(enabled)

        def _refresh_material_price_default() -> None:
            nonlocal material_price_internal_update, material_price_auto_filled
            if material_price_user_touched:
                return
            if _selected_profile_family() == "Tubo":
                _sync_material_unit_for_family()
                return
            if material_price_spin.value() > 0.0:
                return
            try:
                profiles = dict(laser_settings.get("commercial_profiles", {}) or {})
                profile = dict(profiles.get(commercial_combo.currentText().strip(), {}) or {})
                family = _laser_canonical_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip()
                subtype = subtype_combo.currentText().strip()
                material = dict(dict(profile.get("materials", {}) or {}).get(family, {}) or {})
                catalog = dict(dict(profile.get("material_catalog", {}) or {}).get(family, {}) or {})
                if subtype and subtype in catalog:
                    material.update(dict(catalog.get(subtype, {}) or {}))
                price = float(material.get("price_per_kg", 0.0) or 0.0)
            except Exception:
                price = 0.0
            if price > 0.0:
                material_price_internal_update = True
                try:
                    material_price_unit_combo.setCurrentText("EUR/kg")
                    material_price_spin.setValue(price)
                    material_price_auto_filled = True
                finally:
                    material_price_internal_update = False

        def _normalize_counts(cuts_value: float, holes_value: float, slots_value: float) -> tuple[int, int, int, int]:
            holes_count = max(0, int(round(float(holes_value or 0.0))))
            slots_count = max(0, int(round(float(slots_value or 0.0))))
            total_cut_count = max(0, int(round(float(cuts_value or 0.0))))
            if total_cut_count < (holes_count + slots_count):
                total_cut_count = holes_count + slots_count
            outer_cut_count = max(0, total_cut_count - holes_count - slots_count)
            return total_cut_count, holes_count, slots_count, outer_cut_count

        def _estimate_profile_operations(path_txt: str) -> dict[str, Any]:
            family_txt = _infer_family(path_txt)
            section_txt = _infer_section(path_txt)
            section_norm = section_txt.upper().replace(" ", "")
            stem_lower = Path(path_txt).stem.lower()
            suffix = Path(path_txt).suffix.lower()
            try:
                cad_analysis = dict(analyze_profile_cut_features(path_txt) or {})
            except Exception as exc:
                cad_analysis = {"note": str(exc)}
            family_txt = _family_from_analysis(cad_analysis, family_txt)
            if not section_txt:
                section_txt = str(cad_analysis.get("section_label", "") or "").strip()
            text = _read_geometry_preview(path_txt)
            cuts = int(cad_analysis.get("cuts", 0) or 0)
            holes = int(cad_analysis.get("holes", 0) or 0)
            slots = int(cad_analysis.get("slots", 0) or 0)
            outer_cuts = int(cad_analysis.get("outer_cuts", max(0, cuts - holes - slots)) or 0)
            generic_cuts = int(cad_analysis.get("generic_cuts", 0) or 0)
            end_cut_count = int(cad_analysis.get("end_cut_count", 0) or 0)
            cut_length_m = float(cad_analysis.get("cut_length_m", 0.0) or 0.0)
            internal_cut_length_m = float(cad_analysis.get("feature_cut_length_m", 0.0) or 0.0)
            if internal_cut_length_m <= 0.0:
                internal_cut_length_m = (
                    float(cad_analysis.get("hole_cut_length_mm", 0.0) or 0.0)
                    + float(cad_analysis.get("slot_cut_length_mm", 0.0) or 0.0)
                    + float(cad_analysis.get("generic_cut_length_mm", 0.0) or 0.0)
                ) / 1000.0
            if internal_cut_length_m > 0.0:
                cut_length_m = internal_cut_length_m
            if cad_analysis:
                cuts = int(max(0, holes + slots + generic_cuts))
                outer_cuts = int(max(0, generic_cuts))
            notes: list[str] = []
            inferred_material_family, inferred_material_subtype = _infer_material_from_profile(path_txt, text)
            cad_note = str(cad_analysis.get("note", "") or "").strip()
            if cad_note:
                notes.append(cad_note)
            if inferred_material_family:
                material_label = _laser_display_material_family(inferred_material_family) or inferred_material_family
                notes.append(
                    f"material sugerido: {material_label}{f' / {inferred_material_subtype}' if inferred_material_subtype else ''}"
                )
            complex_tokens = ("mitra", "chanfro", "angulo", "bisel", "45")
            if any(token in stem_lower for token in complex_tokens) and not cad_analysis:
                notes.append("nome sugere cortes angulados; confirma cortes internos")
            if text and not cad_analysis:
                plane_count = len(re.findall(r"\bPLANE\b", text, re.IGNORECASE))
                cylindrical_count = len(re.findall(r"\bCYLINDRICAL_SURFACE\b", text, re.IGNORECASE))
                if suffix in {".step", ".stp"}:
                    circle_count = len(re.findall(r"\bCIRCLE\s*\(", text, re.IGNORECASE))
                    ellipse_count = len(re.findall(r"\bELLIPSE\s*\(", text, re.IGNORECASE))
                    spline_count = len(re.findall(r"\bB_SPLINE_CURVE(?:_WITH_KNOTS)?\b", text, re.IGNORECASE))
                    notes.append("leitura STEP textual aplicada")
                else:
                    circle_count = len(re.findall(r"(^|,)\s*100\s*,", text, re.MULTILINE))
                    ellipse_count = len(re.findall(r"(^|,)\s*104\s*,", text, re.MULTILINE))
                    spline_count = len(re.findall(r"(^|,)\s*126\s*,", text, re.MULTILINE))
                    notes.append("leitura IGES textual aplicada")
                if holes <= 0 and 0 < circle_count <= 12:
                    holes = int(circle_count)
                elif any(token in stem_lower for token in ("furo", "furos", "hole", "holes")):
                    holes = 1
                if slots <= 0 and 0 < ellipse_count <= 8:
                    slots = int(ellipse_count)
                elif any(token in stem_lower for token in ("rasgo", "rasgos", "slot", "slots", "oblongo")):
                    slots = 1
                if slots == 0 and 0 < spline_count <= 4 and any(token in stem_lower for token in ("slot", "rasgo", "oblongo")):
                    slots = int(spline_count)
                if any(token in section_norm for token in ("IPE", "IPN", "HEA", "HEB", "HEM", "UPN", "UNP")):
                    notes.append("perfil estrutural identificado; cortes terminais nao sao cobrados neste criterio")
                elif family_txt == "Tubo":
                    if any(token in section_norm for token in ("RHS", "SHS")) or ("X" in section_norm):
                        notes.append("tubo/perfil retangular detetado; cobrados apenas cortes internos")
                    elif re.search(r"(CHS|ROUND|REDONDO|DN\d+|D\d+|Ø)", section_norm) or cylindrical_count >= 4:
                        notes.append("tubo redondo identificado; cobrados apenas cortes internos")
                elif family_txt in {"Cantoneira", "Perfil", "Barra"}:
                    notes.append("cobrados apenas cortes internos; extremidades ignoradas")
            elif not cad_analysis:
                notes.append("sem leitura textual disponivel; sem fallback de cortes exteriores")
            cuts, holes, slots, outer_cuts = _normalize_counts(cuts, holes, slots)
            if generic_cuts > 0 or end_cut_count > 0:
                notes.append(
                    f"leitura atual: {int(cuts)} cortes internos cobrados; {int(end_cut_count)} cortes terminais ignorados"
                )
            else:
                notes.append(
                    f"leitura atual: {int(cuts)} cortes internos = {int(holes)} furos + {int(slots)} rasgos + {int(outer_cuts)} outros internos"
                )
            if cut_length_m > 0:
                notes.append(f"comprimento medido: {cut_length_m:.3f} m")
            else:
                notes.append("comprimento nao medido automaticamente; valor editavel")
            return {
                "family": family_txt,
                "section": section_txt,
                "cuts": int(max(0, cuts)),
                "holes": int(max(0, holes)),
                "slots": int(max(0, slots)),
                "outer_cuts": int(max(0, outer_cuts)),
                "cut_length_m": round(max(0.0, cut_length_m), 4),
                "profile_length_m": float(cad_analysis.get("profile_length_m", 0.0) or 0.0) or _infer_profile_length_m(path_txt),
                "profile_kg_m": float(cad_analysis.get("profile_kg_m", 0.0) or 0.0),
                "thickness_mm": float(cad_analysis.get("thickness_mm_guess", 0.0) or 0.0),
                "material_family": inferred_material_family,
                "material_subtype": inferred_material_subtype,
                "cad_analysis": dict(cad_analysis or {}),
                "note": ". ".join(part for part in notes if part).strip(),
            }

        def _row_payload(row_index: int) -> dict[str, Any]:
            file_item = files_table.item(row_index, 0)
            path_txt = str(file_item.data(Qt.UserRole) if isinstance(file_item, QTableWidgetItem) else "").strip()
            cached_estimate = dict(file_item.data(Qt.UserRole + 1) if isinstance(file_item, QTableWidgetItem) else {} or {})
            family_combo = files_table.cellWidget(row_index, 1)
            section_edit = files_table.cellWidget(row_index, 2)
            qty_spin = files_table.cellWidget(row_index, 3)
            cuts_spin = files_table.cellWidget(row_index, 4)
            holes_spin = files_table.cellWidget(row_index, 5)
            slots_spin = files_table.cellWidget(row_index, 6)
            cut_length_spin = files_table.cellWidget(row_index, 7)
            profile_length_spin = files_table.cellWidget(row_index, 8)
            profile_kg_m_spin = files_table.cellWidget(row_index, 9)
            family_txt = family_combo.currentText().strip() if isinstance(family_combo, QComboBox) else str(cached_estimate.get("family", "Perfil") or "Perfil")
            section_txt = section_edit.text().strip() if isinstance(section_edit, QLineEdit) else str(cached_estimate.get("section", "") or "")
            qty = int(round(float(qty_spin.value() if isinstance(qty_spin, QDoubleSpinBox) else 0.0)))
            cut_length_m = float(cut_length_spin.value() if isinstance(cut_length_spin, QDoubleSpinBox) else float(cached_estimate.get("cut_length_m", 0.0) or 0.0))
            profile_length_m = float(profile_length_spin.value() if isinstance(profile_length_spin, QDoubleSpinBox) else float(cached_estimate.get("profile_length_m", 0.0) or 0.0))
            profile_kg_m = float(profile_kg_m_spin.value() if isinstance(profile_kg_m_spin, QDoubleSpinBox) else float(cached_estimate.get("profile_kg_m", 0.0) or 0.0))
            cuts, holes, slots, outer_cuts = _normalize_counts(
                cuts_spin.value() if isinstance(cuts_spin, QDoubleSpinBox) else 0.0,
                holes_spin.value() if isinstance(holes_spin, QDoubleSpinBox) else 0.0,
                slots_spin.value() if isinstance(slots_spin, QDoubleSpinBox) else 0.0,
            )
            material_family = _laser_canonical_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip() or "Aco carbono"
            subtype_txt = subtype_combo.currentText().strip()
            material_price_unit = material_price_unit_combo.currentText().strip()
            if family_txt == "Tubo":
                material_price_unit = "EUR/m"
            return {
                "path": path_txt,
                "machine_name": machine_combo.currentText().strip(),
                "commercial_name": commercial_combo.currentText().strip(),
                "material": material_family,
                "material_subtype": subtype_txt,
                "gas": gas_combo.currentText().strip(),
                "thickness_mm": float(thickness_spin.value() or 0.0),
                "qtd": max(1, qty),
                "material_supplied_by_client": bool(customer_material_check.isChecked()),
                "material_fornecido_cliente": bool(customer_material_check.isChecked()),
                "profile_material_price": float(material_price_spin.value() or 0.0),
                "profile_material_price_unit": material_price_unit,
                "profile_length_m": max(0.0, profile_length_m),
                "profile_kg_m": max(0.0, profile_kg_m),
                "profile_family": family_txt or "Perfil",
                "section": section_txt,
                "cuts": cuts,
                "holes": holes,
                "slots": slots,
                "outer_cuts": outer_cuts,
                "cut_length_m_override": max(0.0, cut_length_m),
                "include_external_profile_cuts": False,
                "include_profile_event_rates": False,
                "profile_metrics": dict(cached_estimate.get("cad_analysis", {}) or {}),
            }

        def _recalc_total() -> None:
            total = 0.0
            status_parts: list[str] = []
            for row_index in range(files_table.rowCount()):
                try:
                    analysis = dict(self.backend.profile_laser_quote_analyze(_row_payload(row_index)) or {})
                except Exception as exc:
                    total_label.setText("Total estimado: -")
                    status_label.setText(str(exc))
                    return
                pricing = dict(analysis.get("pricing", {}) or {})
                metrics = dict(analysis.get("metrics", {}) or {})
                total += float(pricing.get("total_price", 0.0) or 0.0)
                end_count = int(metrics.get("raw_end_cut_count", metrics.get("end_cut_count", 0)) or 0)
                holes_count = int(metrics.get("hole_count", 0) or 0)
                slots_count = int(metrics.get("slot_count", 0) or 0)
                other_count = int(metrics.get("generic_cut_count", 0) or 0)
                cut_length_m = float(metrics.get("cut_length_m", 0.0) or 0.0)
                thickness_factor = float(metrics.get("thickness_rate_factor", 1.0) or 1.0)
                density_value = float(metrics.get("density_kg_m3", 0.0) or 0.0)
                material_cost = float(pricing.get("material_cost_unit", 0.0) or 0.0)
                material_price_value = float(pricing.get("profile_material_price", 0.0) or 0.0)
                material_price_unit = str(pricing.get("profile_material_price_unit", "") or "").strip().upper()
                cut_length_widget = files_table.cellWidget(row_index, 7)
                if isinstance(cut_length_widget, QDoubleSpinBox) and cut_length_widget.value() <= 0.0 and cut_length_m > 0.0:
                    cut_length_widget.blockSignals(True)
                    cut_length_widget.setValue(cut_length_m)
                    cut_length_widget.blockSignals(False)
                status_parts.append(
                    f"{int(metrics.get('cut_event_count', 0) or 0)} eventos cobrados = "
                    f"{holes_count} furos + {slots_count} rasgos + {other_count} outros; "
                    f"{end_count} terminais ignorados | {cut_length_m:.3f} m | "
                    f"fator esp. x{thickness_factor:.2f} | dens. {density_value:.0f} kg/m3 | "
                    f"MP {_fmt_eur(material_cost)} @ {material_price_value:.4f} {material_price_unit}"
                )
            total_label.setText(f"Total estimado: {_fmt_eur(total)}")
            status_label.setText(status_parts[0] if status_parts else "Usa as tabelas do laser com contagem de eventos STEP/IGS, sem duplicar furos.")

        def _add_file_row(path_txt: str) -> None:
            estimate = _estimate_profile_operations(path_txt)
            row_index = files_table.rowCount()
            files_table.insertRow(row_index)
            estimated_thickness = float(estimate.get("thickness_mm", 0.0) or 0.0)
            if row_index == 0 and estimated_thickness > 0.0:
                thickness_spin.blockSignals(True)
                thickness_spin.setValue(estimated_thickness)
                thickness_spin.blockSignals(False)
            if row_index == 0 and str(estimate.get("material_family", "") or "").strip():
                material_display = _laser_display_material_family(str(estimate.get("material_family", "") or "").strip())
                if material_display:
                    material_combo.setCurrentText(material_display)
                    _refresh_subtypes()
                subtype_txt = str(estimate.get("material_subtype", "") or "").strip()
                if subtype_txt:
                    subtype_combo.setCurrentText(subtype_txt)
                    _refresh_gases()
            file_item = QTableWidgetItem(Path(path_txt).name)
            file_item.setData(Qt.UserRole, str(path_txt))
            file_item.setData(Qt.UserRole + 1, dict(estimate))
            file_item.setData(Qt.UserRole + 2, {})
            file_item.setToolTip(f"{path_txt}\n{str(estimate.get('note', '') or '').strip()}")
            files_table.setItem(row_index, 0, file_item)

            family_combo = QComboBox()
            family_combo.addItems(families)
            family_combo.setCurrentText(str(estimate.get("family", "") or _infer_family(path_txt)))
            section_edit = QLineEdit(str(estimate.get("section", "") or _infer_section(path_txt)))
            qty_spin = _make_int_spin(1)
            cuts_spin = _make_int_spin(int(estimate.get("cuts", 2) or 2))
            holes_spin = _make_int_spin(int(estimate.get("holes", 0) or 0))
            slots_spin = _make_int_spin(int(estimate.get("slots", 0) or 0))
            cut_length_spin = QDoubleSpinBox()
            cut_length_spin.setRange(0.0, 1000000.0)
            cut_length_spin.setDecimals(3)
            cut_length_spin.setSingleStep(0.1)
            cut_length_spin.setValue(float(estimate.get("cut_length_m", 0.0) or 0.0))
            cut_length_spin.setMinimumHeight(32)
            cut_length_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
            profile_length_spin = QDoubleSpinBox()
            profile_length_spin.setRange(0.0, 1000000.0)
            profile_length_spin.setDecimals(3)
            profile_length_spin.setSingleStep(0.1)
            profile_length_spin.setValue(float(estimate.get("profile_length_m", 0.0) or 0.0))
            profile_length_spin.setMinimumHeight(32)
            profile_length_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
            profile_kg_m_spin = QDoubleSpinBox()
            profile_kg_m_spin.setRange(0.0, 1000000.0)
            profile_kg_m_spin.setDecimals(4)
            profile_kg_m_spin.setSingleStep(0.1)
            profile_kg_m_spin.setValue(float(estimate.get("profile_kg_m", 0.0) or 0.0))
            profile_kg_m_spin.setMinimumHeight(32)
            profile_kg_m_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
            family_combo.setMinimumHeight(32)
            section_edit.setMinimumHeight(32)
            files_table.setCellWidget(row_index, 1, family_combo)
            files_table.setCellWidget(row_index, 2, section_edit)
            files_table.setCellWidget(row_index, 3, qty_spin)
            files_table.setCellWidget(row_index, 4, cuts_spin)
            files_table.setCellWidget(row_index, 5, holes_spin)
            files_table.setCellWidget(row_index, 6, slots_spin)
            files_table.setCellWidget(row_index, 7, cut_length_spin)
            files_table.setCellWidget(row_index, 8, profile_length_spin)
            files_table.setCellWidget(row_index, 9, profile_kg_m_spin)
            files_table.setRowHeight(row_index, 42)
            analysis_note = str(estimate.get("note", "") or "").strip()
            for widget in (family_combo, section_edit, qty_spin, cuts_spin, holes_spin, slots_spin, cut_length_spin, profile_length_spin, profile_kg_m_spin):
                widget.setToolTip(analysis_note)
            for widget in (qty_spin, cuts_spin, holes_spin, slots_spin, cut_length_spin, profile_length_spin, profile_kg_m_spin):
                widget.valueChanged.connect(_recalc_total)
            family_combo.currentTextChanged.connect(_sync_material_unit_for_family)
            family_combo.currentTextChanged.connect(_recalc_total)
            section_edit.textChanged.connect(_recalc_total)
            files_table.selectRow(row_index)
            _sync_material_unit_for_family()
            _update_step_preview()

        def _pick_files() -> None:
            paths, _ = QFileDialog.getOpenFileNames(
                dialog,
                "Selecionar ficheiros STEP/IGS",
                "",
                "Modelos 3D (*.step *.stp *.igs *.iges);;Todos (*.*)",
            )
            for path_txt in paths:
                if path_txt:
                    _add_file_row(path_txt)
            _recalc_total()

        def _remove_selected() -> None:
            selected_rows = sorted({item.row() for item in files_table.selectedItems() if item is not None}, reverse=True)
            for row_index in selected_rows:
                files_table.removeRow(row_index)
            _recalc_total()
            _update_step_preview()

        def _on_material_price_changed(_value: float) -> None:
            nonlocal material_price_user_touched, material_price_auto_filled
            if not material_price_internal_update:
                material_price_user_touched = True
                material_price_auto_filled = False
            _recalc_total()

        def _on_material_unit_changed(_text: str) -> None:
            if not material_price_internal_update and _selected_profile_family() == "Tubo":
                _sync_material_unit_for_family()
            _recalc_total()

        def _accept() -> None:
            nonlocal result_lines
            if files_table.rowCount() == 0:
                QMessageBox.warning(dialog, "STEP/IGS", "Adiciona pelo menos um ficheiro STEP/IGS.")
                return
            lines: list[dict[str, Any]] = []
            for row_index in range(files_table.rowCount()):
                payload = _row_payload(row_index)
                if int(payload.get("qtd", 0) or 0) <= 0:
                    continue
                try:
                    result = dict(self.backend.profile_laser_quote_build_line(payload) or {})
                except Exception as exc:
                    QMessageBox.critical(dialog, "STEP/IGS", str(exc))
                    return
                analysis = dict(result.get("analysis", {}) or {})
                line = dict(result.get("line", {}) or {})
                if not line:
                    continue
                metrics = dict(analysis.get("metrics", {}) or {})
                line["tipo_item"] = self.backend.desktop_main.ORC_LINE_TYPE_SERVICE
                line["produto_unid"] = "UN"
                line["line_origin"] = "step_igs_profile_laser"
                line["summary_html"] = (
                    f"{str(payload.get('profile_family', 'Perfil') or 'Perfil')} | "
                    f"cortes internos {int(metrics.get('cut_event_count', 0) or 0)} | "
                    f"m corte {float(metrics.get('cut_length_m', 0.0) or 0.0):.3f} | "
                    f"furos {int(metrics.get('hole_count', 0) or 0)} | "
                    f"rasgos {int(metrics.get('slot_count', 0) or 0)} | "
                    f"terminais ignorados {int(metrics.get('raw_end_cut_count', 0) or 0)}"
                )
                lines.append(line)
            if not lines:
                QMessageBox.warning(dialog, "STEP/IGS", "Nao foi possivel gerar linhas com os dados atuais.")
                return
            if return_lines:
                result_lines = [dict(row or {}) for row in lines]
                dialog.accept()
                return
            self.line_rows.extend(lines)
            if not self.workcenter_combo.currentText().strip():
                self.workcenter_combo.setCurrentText("Laser")
            self._render_quote_lines()
            dialog.accept()

        add_files_btn.clicked.connect(_pick_files)
        remove_file_btn.clicked.connect(_remove_selected)
        configure_btn.clicked.connect(_configure_profiles)
        buttons.accepted.connect(_accept)
        buttons.rejected.connect(dialog.reject)
        files_table.itemSelectionChanged.connect(_update_step_preview)
        files_table.itemSelectionChanged.connect(_sync_material_unit_for_family)
        machine_combo.currentTextChanged.connect(_refresh_materials)
        machine_combo.currentTextChanged.connect(_recalc_total)
        commercial_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
        commercial_combo.currentTextChanged.connect(_recalc_total)
        material_combo.currentTextChanged.connect(_refresh_subtypes)
        material_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
        material_combo.currentTextChanged.connect(_recalc_total)
        subtype_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
        subtype_combo.currentTextChanged.connect(_recalc_total)
        gas_combo.currentTextChanged.connect(_recalc_total)
        thickness_spin.valueChanged.connect(_recalc_total)
        material_price_spin.valueChanged.connect(_on_material_price_changed)
        material_price_unit_combo.currentTextChanged.connect(_on_material_unit_changed)
        customer_material_check.toggled.connect(lambda _checked: _sync_material_cost_controls())
        customer_material_check.toggled.connect(_recalc_total)
        _refresh_machine_and_commercial()
        _refresh_materials()
        _refresh_material_price_default()
        _sync_material_cost_controls()
        _recalc_total()
        _update_step_preview()
        dialog.exec()
        return [dict(row or {}) for row in result_lines] if return_lines else None

    def _add_line(self) -> None:
        try:
            payload = self._line_dialog({"qtd": 1})
            if payload is None:
                return
            self.line_rows.append(payload)
            self._render_quote_lines()
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))

    def _edit_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        try:
            current = dict(self.line_rows[index] or {})
            batch_id = str(current.get("laser_batch_id", "") or "").strip()
            is_saved_batch = (
                str(current.get("laser_source_mode", "") or "").strip().casefold() == "batch"
                or bool(batch_id)
            )
            if self._quote_line_is_laser_2d(current) and is_saved_batch:
                batch_indexes = [
                    row_index
                    for row_index, line in enumerate(self.line_rows)
                    if batch_id
                    and str(dict(line or {}).get("laser_batch_id", "") or "").strip() == batch_id
                ] or [index]
                edited_lines = self._edit_laser_batch_lines(
                    [dict(self.line_rows[row_index] or {}) for row_index in batch_indexes]
                )
                if edited_lines is None:
                    return
                insert_at = min(batch_indexes)
                for row_index in sorted(batch_indexes, reverse=True):
                    del self.line_rows[row_index]
                for offset, edited_line in enumerate(edited_lines):
                    self.line_rows.insert(insert_at + offset, edited_line)
                self._render_quote_lines()
                if edited_lines:
                    self._select_quote_line_source_index(insert_at)
                return
            payload = self._edit_laser_2d_line(current) if self._quote_line_is_laser_2d(current) else self._line_dialog(current)
            if payload is None:
                return
            self.line_rows[index] = payload
            self._render_quote_lines()
            self._select_quote_line_source_index(index)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))

    def _prepare_selected_line_for_production(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        current = dict(self.line_rows[index] or {})
        if not self.backend.desktop_main.orc_line_is_piece(current):
            QMessageBox.information(
                self,
                "Preparar producao",
                "Esta acao so se aplica a pecas fabricadas. Produtos de stock continuam no fluxo de consumo de montagem.",
            )
            return
        payload = self._line_dialog(current)
        if payload is None:
            return
        self.line_rows[index] = payload
        self._render_quote_lines()
        self._select_quote_line_source_index(index)
        drawing_ready = bool(str(payload.get("desenho", "") or "").strip())
        ops_ready = bool(str(payload.get("operacao", "") or "").strip())
        if drawing_ready and ops_ready:
            QMessageBox.information(
                self,
                "Preparar producao",
                "A linha ficou preparada para fluxo de operador. Na conversão para encomenda vai seguir para produção.",
            )
        else:
            QMessageBox.information(
                self,
                "Preparar producao",
                "A linha foi atualizada, mas so entra em producao quando tiver desenho tecnico e operacoes definidas.",
            )

    def _remove_line(self) -> None:
        indexes = self._checked_line_indexes() or self._selected_line_indexes()
        if not indexes:
            QMessageBox.warning(self, "Orçamentos", "Marca ou seleciona pelo menos uma linha.")
            return
        references = [
            self._quote_line_primary_ref(dict(self.line_rows[index] or {}))
            for index in indexes
            if 0 <= index < len(self.line_rows)
        ]
        count = len(indexes)
        detail = ", ".join(reference for reference in references[:3] if reference and reference != "-")
        if len(references) > 3:
            detail += f" e mais {len(references) - 3}"
        answer = QMessageBox.question(
            self,
            "Remover linhas do orçamento",
            (
                f"Remover {count} linha{'s' if count != 1 else ''} do orçamento atual?"
                + (f"\n\n{detail}" if detail else "")
                + "\n\nEsta alteração só fica definitiva quando guardares o orçamento."
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self._rendering_quote_lines = True
        for visual_row in range(self.lines_table.rowCount()):
            item = self.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is not None and item.flags() & Qt.ItemIsUserCheckable:
                item.setCheckState(Qt.Unchecked)
        self._rendering_quote_lines = False
        next_source_index = min(indexes)
        for index in sorted(indexes, reverse=True):
            if 0 <= index < len(self.line_rows):
                del self.line_rows[index]
        self._render_quote_lines()
        if self.line_rows:
            self._select_quote_line_source_index(min(next_source_index, len(self.line_rows) - 1))

    def _open_line_drawing(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        path = str(self.line_rows[index].get("desenho", "") or "").strip()
        if not path:
            QMessageBox.information(self, "Orçamentos", "A linha nao tem desenho associado.")
            return
        if not Path(path).exists():
            QMessageBox.critical(self, "Orçamentos", f"Ficheiro nao encontrado:\n{path}")
            return
        os.startfile(path)

    def _check_selected_line_weight(self) -> None:
        indexes = self._selected_line_indexes()
        if not indexes:
            QMessageBox.warning(self, "Check Weight", "Seleciona pelo menos uma linha.")
            return
        results: list[dict[str, Any]] = []
        issues: list[str] = []
        for index in indexes:
            row = dict(self.line_rows[index] or {})
            desenho = str(row.get("desenho", "") or "").strip()
            if not desenho:
                issues.append(f"Linha {index + 1}: sem desenho associado.")
                continue
            if not Path(desenho).exists():
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: desenho nao encontrado.")
                continue
            esp_raw = row.get("espessura", row.get("espessura_mm", row.get("esp", "")))
            try:
                thickness_mm = float(str(esp_raw).replace(",", ".").strip() or 0)
            except Exception:
                thickness_mm = 0.0
            if thickness_mm <= 0:
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: espessura invalida.")
                continue
            payload = {
                "path": desenho,
                "machine": self.workcenter_combo.currentText().strip(),
                "commercial_profile": "",
                "material": str(row.get("material_family", "") or row.get("material", "") or "").strip(),
                "material_subtype": str(row.get("material_subtype", "") or "").strip(),
                "gas": "",
                "thickness_mm": thickness_mm,
                "quantity": max(1, int(float(row.get("qtd", 1) or 1))),
                "material_supplied_by_client": bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False)),
            }
            try:
                analysis = dict(self.backend.laser_quote_analyze(payload) or {})
            except Exception as exc:
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: {exc}")
                continue
            metrics = dict(analysis.get("metrics", {}) or {})
            mass_unit = float(metrics.get("net_mass_kg", 0.0) or 0.0)
            qty = max(1, int(float(row.get("qtd", 1) or 1)))
            results.append(
                {
                    "row_index": index,
                    "ref": str(row.get("ref_externa", "") or row.get("descricao", "-") or "-").strip(),
                    "material": str((analysis.get("machine", {}) or {}).get("material", row.get("material", "-")) or "-").strip(),
                    "thickness_mm": thickness_mm,
                    "qty": qty,
                    "mass_unit": mass_unit,
                    "mass_total": mass_unit * qty,
                }
            )
        if not results and issues:
            QMessageBox.warning(self, "Check Weight", "\n".join(issues))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Check Weight")
        dialog.resize(760, 360)
        layout = QVBoxLayout(dialog)
        title = QLabel("Peso por unidade das linhas selecionadas")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(title)
        help_lbl = QLabel("Informacao de apoio. O peso e calculado a partir do desenho, espessura e densidade configurada para o material.")
        help_lbl.setProperty("role", "muted")
        help_lbl.setWordWrap(True)
        layout.addWidget(help_lbl)
        table = QTableWidget(len(results), 6)
        table.setHorizontalHeaderLabels(["Ref. Ext.", "Material", "Esp. (mm)", "Qtd", "Peso/un (kg)", "Peso linha (kg)"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setAlternatingRowColors(True)
        table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QHeaderView::section { font-size: 11px; padding: 6px 8px; font-weight: 800; }"
        )
        _set_table_columns(
            table,
            [
                (0, "stretch", 0),
                (1, "fixed", 130),
                (2, "fixed", 80),
                (3, "fixed", 60),
                (4, "fixed", 110),
                (5, "fixed", 118),
            ],
        )
        for row_index, result in enumerate(results):
            values = [
                str(result.get("ref", "-")),
                str(result.get("material", "-")),
                f"{float(result.get('thickness_mm', 0.0) or 0.0):.3f}",
                str(int(result.get("qty", 1) or 1)),
                f"{float(result.get('mass_unit', 0.0) or 0.0):.3f}",
                f"{float(result.get('mass_total', 0.0) or 0.0):.3f}",
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index >= 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                table.setItem(row_index, col_index, item)
        table.setMinimumHeight(max(140, min(320, 42 + (len(results) * 30))))
        layout.addWidget(table)
        if issues:
            issues_lbl = QLabel("Linhas ignoradas:\n" + "\n".join(issues))
            issues_lbl.setWordWrap(True)
            issues_lbl.setStyleSheet("color: #9a3412; font-size: 11px;")
            layout.addWidget(issues_lbl)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _preview_quote(self) -> None:
        try:
            detail = self.backend.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            path = self.backend.orc_open_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF aberto:\n{path}")

    def _save_quote_pdf(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Orçamentos", "Guarda primeiro o orcamento.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", f"orcamento_{self.current_number}.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            detail = self.backend.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            self.backend.orc_render_pdf(numero, path)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF guardado em:\n{path}")

    def _print_quote_pdf(self) -> None:
        try:
            detail = self.backend.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            path = self.backend.orc_print_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF enviado para impressão:\n{path}")

    def _convert_quote(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Orçamentos", "Guarda primeiro o orcamento.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Converter em encomenda")
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        note_edit = QLineEdit(self.note_cliente_edit.text().strip())
        form.addRow("Nota cliente", note_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            self.backend.orc_save(self._quote_payload())
            result = self.backend.orc_convert_to_order(self.current_number, note_edit.text().strip())
        except Exception as exc:
            QMessageBox.critical(self, "Converter em encomenda", str(exc))
            return
        self.refresh()
        self._load_quote(self.current_number)
        self._show_detail()
        enc_num = str(((result or {}).get("encomenda", {}) or {}).get("numero", "") or "").strip()
        QMessageBox.information(self, "Converter em encomenda", f"Encomenda criada: {enc_num or '-'}")
