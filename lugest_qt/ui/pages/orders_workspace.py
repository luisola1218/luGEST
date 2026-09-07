from __future__ import annotations
import unicodedata
from PySide6.QtCore import QDate, QTimer, Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
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
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
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
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    configure_table as _configure_table,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    selected_row_index as _selected_row_index,
    set_panel_tone as _set_panel_tone,
    set_table_columns as _set_table_columns,
    state_tone as _state_tone,
    state_visual as _state_visual,
    table_visible_height as _table_visible_height,
)
from .runtime_support import (
    _adopt_layout_item,
    _build_operation_selector,
    _fmt_eur,
    _format_client_label,
    _make_inline_progress,
    _reference_catalog_dialog,
    _take_layout_items,
)
from ..widgets import CardFrame, ClickableDateEdit as QDateEdit, FlexibleDecimalSpinBox as QDoubleSpinBox


class OrdersPage(QWidget):
    page_title = "Encomendas"
    page_subtitle = "Ordens de fabrico, materiais, peças, aprovisionamento e progresso produtivo."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.all_rows: list[dict] = []
        self.client_rows: list[dict] = []
        self.current_detail: dict = {}
        self.material_rows: list[dict] = []
        self.esp_rows: list[dict] = []
        self.detail_pieces: list[dict] = []
        self.detail_montagem: list[dict] = []
        self.presets: dict = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(16, 14, 16, 14)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Pesquisa")
        self._order_filter_timer = QTimer(self)
        self._order_filter_timer.setSingleShot(True)
        self._order_filter_timer.setInterval(180)
        self._order_filter_timer.timeout.connect(self.refresh)
        self.filter_edit.lineEdit().textChanged.connect(lambda _text: self._order_filter_timer.start())
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Todas", "Preparacao", "Montagem", "Em producao", "Concluida"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.client_combo = QComboBox()
        self.client_combo.currentTextChanged.connect(self.refresh)
        self.new_btn = QPushButton("Criar OF")
        self.new_btn.clicked.connect(self._new_order)
        self.edit_header_btn = QPushButton("Editar")
        self.edit_header_btn.setProperty("variant", "secondary")
        self.edit_header_btn.clicked.connect(self._edit_order_header)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_order)
        filters_layout.addWidget(QLabel("Pesquisa"), 0, 0)
        filters_layout.addWidget(self.filter_edit, 0, 1)
        filters_layout.addWidget(QLabel("Estado"), 0, 2)
        filters_layout.addWidget(self.state_combo, 0, 3)
        filters_layout.addWidget(QLabel("Ano"), 0, 4)
        filters_layout.addWidget(self.year_combo, 0, 5)
        filters_layout.addWidget(QLabel("Cliente"), 1, 0)
        filters_layout.addWidget(self.client_combo, 1, 1, 1, 2)
        filters_layout.addWidget(self.new_btn, 1, 3)
        filters_layout.addWidget(self.edit_header_btn, 1, 4)
        filters_layout.addWidget(self.remove_btn, 1, 5)
        for button, width in (
            (self.new_btn, 152),
            (self.edit_header_btn, 116),
            (self.remove_btn, 116),
        ):
            button.setMinimumWidth(width)
        filters.setMaximumHeight(92)
        root.addWidget(filters)

        self.info_card = CardFrame()
        info_layout = QGridLayout(self.info_card)
        info_layout.setContentsMargins(14, 10, 14, 10)
        info_layout.setHorizontalSpacing(12)
        info_layout.setVerticalSpacing(5)
        self.info_numero = QLabel("-")
        self.info_of = QLabel("-")
        self.info_cliente = QLabel("-")
        self.info_entrega = QLabel("-")
        self.info_estado = QLabel("-")
        self.info_nota = QLabel("-")
        self.info_transporte = QLabel("-")
        self.info_descarga = QLabel("-")
        self.info_viagem = QLabel("-")
        self.info_transportadora = QLabel("-")
        self.info_carga = QLabel("-")
        self.info_custos = QLabel("-")
        self.info_reservas = QLabel("Sem reservas ativas.")
        self.info_reservas.setWordWrap(True)
        self.info_cativar = QLabel("-")
        self.info_chapa = QLabel("-")
        self.info_numero.setProperty("role", "field_value_strong")
        self.info_of.setProperty("role", "field_value_strong")
        self.info_cliente.setProperty("role", "field_value")
        self.info_entrega.setProperty("role", "field_value")
        self.info_nota.setProperty("role", "field_value")
        self.info_transporte.setProperty("role", "field_value")
        self.info_descarga.setProperty("role", "field_value")
        self.info_viagem.setProperty("role", "field_value")
        self.info_transportadora.setProperty("role", "field_value")
        self.info_carga.setProperty("role", "field_value")
        self.info_custos.setProperty("role", "field_value")
        self.info_reservas.setProperty("role", "field_value")
        self.info_chapa.setProperty("role", "field_value")
        self.info_nota.setWordWrap(True)
        self.info_transporte.setWordWrap(True)
        self.info_descarga.setWordWrap(True)
        self.info_viagem.setWordWrap(True)
        self.info_transportadora.setWordWrap(True)
        self.info_carga.setWordWrap(True)
        self.info_custos.setWordWrap(True)
        self.info_chapa.setWordWrap(True)
        _apply_state_chip(self.info_estado, "-")
        _apply_state_chip(self.info_cativar, "-", "-")
        labels = [
            ("Numero", self.info_numero, 0, 0),
            ("Cliente", self.info_cliente, 0, 2),
            ("Estado", self.info_estado, 0, 4),
            ("OF", self.info_of, 1, 0),
            ("Nota cliente", self.info_nota, 1, 2),
            ("Cativar MP", self.info_cativar, 1, 4),
            ("Entrega", self.info_entrega, 2, 0),
            ("Transporte", self.info_transporte, 2, 2),
            ("Descarga", self.info_descarga, 2, 4),
            ("Transportadora", self.info_transportadora, 3, 0),
            ("Carga", self.info_carga, 3, 2),
            ("Custos", self.info_custos, 3, 4),
            ("Chapa cativada", self.info_chapa, 4, 0),
            ("Reservas", self.info_reservas, 4, 2),
        ]
        for title, label, row, col in labels:
            title_widget = QLabel(title)
            title_widget.setProperty("role", "field_label")
            info_layout.addWidget(title_widget, row, col)
            info_layout.addWidget(label, row, col + 1, 1, 1)
        self.reserve_btn = QPushButton("Cativar MP")
        self.reserve_btn.clicked.connect(self._reserve_stock)
        self.release_btn = QPushButton("Descativar MP")
        self.release_btn.setProperty("variant", "secondary")
        self.release_btn.clicked.connect(self._release_stock)
        self.reserve_btn.setProperty("compact", "true")
        self.release_btn.setProperty("compact", "true")
        self.reserve_btn.setMinimumWidth(138)
        self.release_btn.setMinimumWidth(138)
        info_layout.addWidget(self.reserve_btn, 4, 4)
        info_layout.addWidget(self.release_btn, 4, 5)
        self.info_card.set_tone("default")
        self.info_card.setMaximumHeight(182)
        root.addWidget(self.info_card)

        list_card = CardFrame()
        list_card.set_tone("default")
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(16, 14, 16, 14)
        list_header = QHBoxLayout()
        list_title = QLabel("Ordens de fabrico")
        list_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.orders_meta_label = QLabel("0 registos")
        self.orders_meta_label.setProperty("role", "muted")
        list_header.addWidget(list_title)
        list_header.addStretch(1)
        list_header.addWidget(self.orders_meta_label)
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            ["Encomenda", "OF", "Cliente", "Referência cliente", "Entrega", "Estado", "MP", "Peças", "Progresso"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setStyleSheet(
            "QTableWidget { font-size: 12px; alternate-background-color: #f8fafc;"
            " background: #ffffff; border: 1px solid #d6e0eb; border-radius: 6px; }"
            "QTableWidget::item { padding: 5px 8px; border-bottom: 1px solid #edf2f7; }"
            "QTableWidget::item:selected { background: #efe4c8; color: #4a321b;"
            " border: 1px solid #c9b387; }"
            "QHeaderView::section { font-size: 11px; padding: 7px 9px; font-weight: 900;"
            " background: #eef4f8; color: #243b53; border: none; border-bottom: 1px solid #cbd8e6; }"
        )
        self.table.verticalHeader().setDefaultSectionSize(34)
        _configure_table(self.table, stretch=(2,), contents=(4, 5, 6, 7, 8))
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 142),
                (1, "fixed", 154),
                (2, "stretch", 0),
                (3, "fixed", 190),
                (4, "fixed", 108),
                (5, "fixed", 122),
                (6, "fixed", 70),
                (7, "fixed", 68),
                (8, "fixed", 96),
            ],
        )
        self.table.itemSelectionChanged.connect(self._on_order_selected)
        list_layout.addLayout(list_header)
        list_layout.addWidget(self.table)
        root.addWidget(list_card)

        mid = QSplitter(Qt.Horizontal)
        mid.setChildrenCollapsible(False)

        materials_card = CardFrame()
        materials_card.set_tone("info")
        materials_layout = QVBoxLayout(materials_card)
        materials_layout.setContentsMargins(16, 14, 16, 14)
        materials_layout.setSpacing(8)
        materials_title = QLabel("Materiais")
        materials_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.materials_table = QTableWidget(0, 4)
        self.materials_table.setHorizontalHeaderLabels(["Material", "Esp.", "Peças", "Estado"])
        self.materials_table.verticalHeader().setVisible(False)
        self.materials_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.materials_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.materials_table, stretch=(0,), contents=(1, 2, 3))
        _set_table_columns(
            self.materials_table,
            [
                (0, "stretch", 210),
                (1, "fixed", 54),
                (2, "fixed", 58),
                (3, "fixed", 104),
            ],
        )
        self.materials_table.verticalHeader().setDefaultSectionSize(30)
        self.materials_table.setStyleSheet(
            "QTableWidget { font-size: 11px; }"
            "QHeaderView::section { font-size: 11px; font-weight: 800; padding: 5px 8px; }"
        )
        self.materials_table.itemSelectionChanged.connect(self._on_material_selected)
        materials_btns = QHBoxLayout()
        self.add_material_btn = QPushButton("Adicionar material")
        self.add_material_btn.clicked.connect(self._add_material)
        self.remove_material_btn = QPushButton("Remover material")
        self.remove_material_btn.setProperty("variant", "secondary")
        self.remove_material_btn.clicked.connect(self._remove_material)
        self.add_material_btn.setMinimumWidth(118)
        self.remove_material_btn.setMinimumWidth(118)
        materials_btns.addWidget(self.add_material_btn)
        materials_btns.addWidget(self.remove_material_btn)
        materials_layout.addWidget(materials_title)
        materials_layout.addWidget(self.materials_table)
        materials_layout.addLayout(materials_btns)
        materials_card.setMinimumHeight(190)
        materials_card.setMinimumWidth(470)
        mid.addWidget(materials_card)

        esp_card = CardFrame()
        esp_card.set_tone("warning")
        esp_layout = QVBoxLayout(esp_card)
        esp_layout.setContentsMargins(16, 14, 16, 14)
        esp_layout.setSpacing(8)
        esp_title = QLabel("Espessuras")
        esp_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.esp_table = QTableWidget(0, 5)
        self.esp_table.setHorizontalHeaderLabels(["Espessura", "Laser (min)", "Outras operações", "Recursos", "Estado"])
        self.esp_table.verticalHeader().setVisible(False)
        self.esp_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.esp_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.esp_table, stretch=(2,), contents=(0, 1, 3))
        _set_table_columns(
            self.esp_table,
            [
                (0, "fixed", 88),
                (1, "fixed", 76),
                (2, "stretch", 150),
                (3, "stretch", 180),
                (4, "fixed", 126),
            ],
        )
        self.esp_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.esp_table.horizontalHeader().setStretchLastSection(True)
        self.esp_table.verticalHeader().setDefaultSectionSize(30)
        self.esp_table.setStyleSheet(
            "QTableWidget { font-size: 11px; }"
            "QHeaderView::section { font-size: 11px; font-weight: 800; padding: 5px 8px; }"
        )
        self.esp_table.itemSelectionChanged.connect(self._on_esp_selected)
        esp_btns = QGridLayout()
        self.add_esp_btn = QPushButton("Adicionar espessura")
        self.add_esp_btn.clicked.connect(self._add_espessura)
        self.remove_esp_btn = QPushButton("Remover espessura")
        self.remove_esp_btn.setProperty("variant", "secondary")
        self.remove_esp_btn.clicked.connect(self._remove_espessura)
        self.edit_time_btn = QPushButton("Configurar tempos")
        self.edit_time_btn.setProperty("variant", "secondary")
        self.edit_time_btn.clicked.connect(self._edit_esp_time)
        self.add_esp_btn.setMinimumWidth(118)
        self.remove_esp_btn.setMinimumWidth(118)
        self.edit_time_btn.setMinimumWidth(118)
        esp_btns.setHorizontalSpacing(8)
        esp_btns.setVerticalSpacing(8)
        esp_btns.addWidget(self.add_esp_btn, 0, 0)
        esp_btns.addWidget(self.remove_esp_btn, 0, 1)
        esp_btns.addWidget(self.edit_time_btn, 1, 0, 1, 2)
        esp_layout.addWidget(esp_title)
        esp_layout.addWidget(self.esp_table)
        esp_layout.addLayout(esp_btns)
        esp_card.setMinimumHeight(190)
        esp_card.setMinimumWidth(470)
        mid.addWidget(esp_card)

        mid.setStretchFactor(0, 1)
        mid.setStretchFactor(1, 1)
        mid.setSizes([520, 500])
        root.addWidget(mid)

        pieces_card = CardFrame()
        pieces_card.set_tone("default")
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(16, 14, 16, 14)
        pieces_header = QHBoxLayout()
        pieces_title = QLabel("Peças da ordem")
        pieces_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.add_piece_btn = QPushButton("Nova peça")
        self.add_piece_btn.clicked.connect(self._add_piece)
        self.import_model_btn = QPushButton("Adicionar conjunto")
        self.import_model_btn.setProperty("variant", "secondary")
        self.import_model_btn.clicked.connect(self._import_model)
        self.print_of_btn = QPushButton("Pré-visualizar OF")
        self.print_of_btn.setProperty("variant", "secondary")
        self.print_of_btn.clicked.connect(self._print_fabrication_order)
        self.edit_piece_btn = QPushButton("Editar peça")
        self.edit_piece_btn.setProperty("variant", "secondary")
        self.edit_piece_btn.clicked.connect(self._edit_piece)
        self.remove_piece_btn = QPushButton("Remover peça")
        self.remove_piece_btn.setProperty("variant", "secondary")
        self.remove_piece_btn.clicked.connect(self._remove_piece)
        self.open_piece_btn = QPushButton("Abrir desenho")
        self.open_piece_btn.setProperty("variant", "secondary")
        self.open_piece_btn.clicked.connect(self._open_selected_piece_drawing)
        for button, width in (
            (self.add_piece_btn, 92),
            (self.import_model_btn, 132),
            (self.print_of_btn, 128),
            (self.edit_piece_btn, 92),
            (self.remove_piece_btn, 102),
            (self.open_piece_btn, 106),
        ):
            button.setMinimumWidth(width)
            button.setMaximumWidth(width + 18)
            button.setMinimumHeight(30)
        pieces_header.addWidget(pieces_title, 1)
        pieces_header.setSpacing(6)
        for button in (self.add_piece_btn, self.import_model_btn, self.print_of_btn, self.edit_piece_btn, self.remove_piece_btn, self.open_piece_btn):
            pieces_header.addWidget(button)
        self.pieces_table = QTableWidget(0, 9)
        self.pieces_table.setHorizontalHeaderLabels(
            ["Ref. interna", "Ref. externa", "OPP", "Material", "Esp.", "Operações", "Planeada", "Produzida", "Estado"]
        )
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.pieces_table, stretch=(1, 5), contents=(0, 2, 3, 4, 6, 7, 8))
        self.pieces_table.itemSelectionChanged.connect(self._sync_action_buttons)
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.pieces_table)
        pieces_card.setMinimumHeight(520)
        root.addWidget(pieces_card, 2)

        montagem_card = CardFrame()
        montagem_card.set_tone("warning")
        montagem_layout = QVBoxLayout(montagem_card)
        montagem_layout.setContentsMargins(16, 14, 16, 14)
        montagem_layout.setSpacing(8)
        montagem_header = QHBoxLayout()
        montagem_title = QLabel("Montagem e componentes")
        montagem_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.open_operator_montagem_btn = QPushButton("Abrir no Operador")
        self.open_operator_montagem_btn.setProperty("variant", "success")
        self.open_operator_montagem_btn.clicked.connect(self._open_montagem_in_operator)
        self.open_operator_montagem_btn.setMinimumWidth(132)
        self.montagem_note_btn = QPushButton("Criar nota de compra")
        self.montagem_note_btn.setProperty("variant", "secondary")
        self.montagem_note_btn.clicked.connect(self._create_montagem_purchase_note)
        self.montagem_note_btn.setMinimumWidth(154)
        montagem_header.addWidget(montagem_title, 1)
        montagem_header.addWidget(self.open_operator_montagem_btn)
        montagem_header.addWidget(self.montagem_note_btn)
        self.montagem_meta = QLabel("Sem itens de montagem.")
        self.montagem_meta.setWordWrap(True)
        self.montagem_meta.setProperty("role", "muted")
        self.montagem_table = QTableWidget(0, 8)
        self.montagem_table.setHorizontalHeaderLabels(
            ["Tipo", "Código", "Descrição", "Necessário", "Consumido", "Por consumir", "Em falta", "Estado"]
        )
        self.montagem_table.verticalHeader().setVisible(False)
        self.montagem_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.montagem_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.montagem_table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6, 7))
        montagem_layout.addLayout(montagem_header)
        montagem_layout.addWidget(self.montagem_meta)
        montagem_layout.addWidget(self.montagem_table)
        montagem_card.setMinimumHeight(220)
        root.addWidget(montagem_card)

        self._clear_detail()

    def refresh(self) -> None:
        previous_numero = str(self.current_detail.get("numero", "") or "").strip()
        current_filter = self.filter_edit.currentText().strip()
        self.client_rows = self.backend.order_clients()
        self.presets = self.backend.order_presets()
        self.all_rows = self.backend.order_rows("", estado="Todas", ano="Todos", cliente="Todos")
        self._refresh_filter_options()
        self.rows = self.backend.order_rows(
            self.filter_edit.currentText().strip(),
            self.state_combo.currentText(),
            self.year_combo.currentText() or "Todos",
            self.client_combo.currentText() or "Todos",
        )
        self.filter_edit.blockSignals(True)
        if self.filter_edit.count() == 0:
            self.filter_edit.addItem("")
        known_values = {self.filter_edit.itemText(i) for i in range(self.filter_edit.count())}
        for row in self.all_rows:
            if row["numero"] not in known_values:
                self.filter_edit.addItem(row["numero"])
                known_values.add(row["numero"])
        self.filter_edit.setCurrentText(current_filter)
        self.filter_edit.blockSignals(False)
        _fill_table(
            self.table,
            [
                [
                    r.get("numero", "-"),
                    r.get("of", "-"),
                    r.get("cliente", "-"),
                    r.get("nota_cliente", "-"),
                    r.get("data_entrega", "-"),
                    r.get("estado", "-"),
                    r.get("cativar", "NAO"),
                    r.get("pecas", 0),
                    f"{r.get('progress', 0):.1f}%",
                ]
                for r in self.rows
            ],
            align_center_from=4,
        )
        self.table.setSortingEnabled(False)
        self.orders_meta_label.setText(f"{len(self.rows)} OF no filtro atual")
        for row_index, row in enumerate(self.rows):
            item = self.table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("numero", "") or "").strip())
                item.setFont(QFont(item.font().family(), item.font().pointSize(), QFont.Bold))
            of_item = self.table.item(row_index, 1)
            if of_item is not None:
                of_item.setForeground(QBrush(QColor("#1d4ed8")))
                of_item.setFont(QFont(of_item.font().family(), of_item.font().pointSize(), QFont.DemiBold))
            state_item = self.table.item(row_index, 5)
            state_visual = _state_visual(str(row.get("estado", "")))
            if state_item is not None:
                state_item.setBackground(QBrush(QColor(state_visual["bg"])))
                state_item.setForeground(QBrush(QColor(state_visual["fg"])))
                state_item.setFont(QFont(state_item.font().family(), state_item.font().pointSize(), QFont.DemiBold))
            mp_item = self.table.item(row_index, 6)
            if mp_item is not None:
                reserved = str(row.get("cativar", "") or "").strip().upper() == "SIM"
                mp_item.setText("Cativada" if reserved else "Livre")
                mp_item.setForeground(QBrush(QColor("#166534" if reserved else "#64748b")))
            progress_item = self.table.item(row_index, 8)
            if progress_item is not None:
                progress = float(row.get("progress", 0) or 0)
                progress_item.setForeground(QBrush(QColor("#166534" if progress >= 100 else "#1d4ed8")))
                progress_item.setFont(QFont(progress_item.font().family(), progress_item.font().pointSize(), QFont.Bold))
                self.table.setCellWidget(row_index, 8, _make_inline_progress(progress))
        if self.table.rowCount() == 0:
            self._clear_detail()
            return
        target_row = 0
        if previous_numero:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "") or "").strip() == previous_numero:
                    target_row = index
                    break
        self.table.selectRow(target_row)
        self._on_order_selected()

    def _refresh_filter_options(self) -> None:
        current_year = self.year_combo.currentText().strip() or "Todos"
        current_client = self.client_combo.currentText().strip() or "Todos"
        years = sorted({str(row.get("ano", "") or "").strip() for row in self.all_rows if str(row.get("ano", "") or "").strip()}, reverse=True)
        year_values = ["Todos"] + years
        client_values = ["Todos"] + [str(row.get("label", "") or "").strip() for row in self.client_rows]
        self.year_combo.blockSignals(True)
        self.client_combo.blockSignals(True)
        self.year_combo.clear()
        self.client_combo.clear()
        self.year_combo.addItems(year_values or ["Todos"])
        self.client_combo.addItems(client_values or ["Todos"])
        self.year_combo.setCurrentText(current_year if current_year in year_values else "Todos")
        self.client_combo.setCurrentText(current_client if current_client in client_values else "Todos")
        self.year_combo.blockSignals(False)
        self.client_combo.blockSignals(False)

    def _selected_order_row(self) -> dict:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        row_item = self.table.item(row_index, 0)
        numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if numero:
            return next((row for row in self.rows if str(row.get("numero", "") or "").strip() == numero), {})
        return self.rows[row_index] if row_index < len(self.rows) else {}

    def open_order_numero(self, numero: str) -> None:
        target = str(numero or "").strip()
        if not target:
            return
        self.refresh()
        for row_index, row in enumerate(self.rows):
            if str(row.get("numero", "") or "").strip() != target:
                continue
            self.table.selectRow(row_index)
            self._on_order_selected()
            return

    def _selected_material_row(self) -> dict:
        row_index = _selected_row_index(self.materials_table)
        if row_index < 0:
            return {}
        row_item = self.materials_table.item(row_index, 0)
        material = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if material:
            return next((row for row in self.material_rows if str(row.get("material", "") or "").strip() == material), {})
        return self.material_rows[row_index] if row_index < len(self.material_rows) else {}

    def _selected_esp_row(self) -> dict:
        row_index = _selected_row_index(self.esp_table)
        if row_index < 0:
            return {}
        row_item = self.esp_table.item(row_index, 0)
        esp = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if esp:
            return next((row for row in self.esp_rows if str(row.get("espessura", "") or "").strip() == esp), {})
        return self.esp_rows[row_index] if row_index < len(self.esp_rows) else {}

    def _selected_piece_row(self) -> dict:
        row_index = _selected_row_index(self.pieces_table)
        if row_index < 0:
            return {}
        row_item = self.pieces_table.item(row_index, 0)
        ref_int = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if ref_int:
            return next((row for row in self.detail_pieces if str(row.get("ref_interna", "") or "").strip() == ref_int), {})
        return self.detail_pieces[row_index] if row_index < len(self.detail_pieces) else {}

    def _clear_detail(self) -> None:
        self.current_detail = {}
        self.material_rows = []
        self.esp_rows = []
        self.detail_pieces = []
        self.detail_montagem = []
        self.info_numero.setText("-")
        self.info_of.setText("-")
        self.info_cliente.setText("-")
        self.info_entrega.setText("-")
        _apply_state_chip(self.info_estado, "-")
        self.info_nota.setText("-")
        self.info_transporte.setText("-")
        self.info_descarga.setText("-")
        self.info_viagem.setText("-")
        self.info_transportadora.setText("-")
        self.info_carga.setText("-")
        self.info_custos.setText("-")
        _apply_state_chip(self.info_cativar, "-", "-")
        self.info_chapa.setText("-")
        self.info_reservas.setText("Sem reservas ativas.")
        self.montagem_meta.setText("Sem itens de montagem.")
        self.materials_table.setRowCount(0)
        self.esp_table.setRowCount(0)
        self.pieces_table.setRowCount(0)
        self.montagem_table.setRowCount(0)
        _set_panel_tone(self.info_card, "default")
        self._sync_action_buttons()

    def _sync_action_buttons(self) -> None:
        has_order = bool(self.current_detail.get("numero"))
        can_edit = has_order and bool(self.current_detail.get("can_edit_structure", True))
        has_material = bool(self._selected_material_row())
        has_esp = bool(self._selected_esp_row())
        has_piece = bool(self._selected_piece_row())
        self.edit_header_btn.setEnabled(has_order)
        self.remove_btn.setEnabled(has_order)
        self.reserve_btn.setEnabled(has_order and has_material and has_esp)
        self.release_btn.setEnabled(has_order and has_material and has_esp)
        self.reserve_btn.setToolTip("")
        self.release_btn.setToolTip("")
        self.add_material_btn.setEnabled(can_edit)
        self.remove_material_btn.setEnabled(can_edit and has_material)
        self.add_esp_btn.setEnabled(can_edit and has_material)
        self.remove_esp_btn.setEnabled(can_edit and has_esp)
        self.edit_time_btn.setEnabled(has_esp)
        self.add_piece_btn.setEnabled(can_edit and has_order)
        self.import_model_btn.setEnabled(can_edit and has_order)
        self.print_of_btn.setEnabled(has_order)
        self.edit_piece_btn.setEnabled(can_edit and has_piece)
        self.remove_piece_btn.setEnabled(can_edit and has_piece)
        self.open_piece_btn.setEnabled(has_piece)
        self.open_operator_montagem_btn.setEnabled(has_order and bool(self.detail_montagem))
        self.open_operator_montagem_btn.setToolTip(
            "A baixa de componentes e materia-prima fica registada no menu Operador."
            if has_order and bool(self.detail_montagem)
            else ""
        )
        self.montagem_note_btn.setEnabled(has_order and bool(list(self.current_detail.get("montagem_shortages", []) or [])))

    def _on_order_selected(self) -> None:
        row = self._selected_order_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            self._clear_detail()
            return
        try:
            detail = self.backend.order_detail(numero)
        except Exception as exc:
            self._clear_detail()
            self.info_reservas.setText(str(exc))
            return
        self.current_detail = detail
        self.info_numero.setText(str(detail.get("numero", "-")))
        self.info_of.setText(str(detail.get("of_codigo", "-") or "-"))
        self.info_cliente.setText(
            _format_client_label(
                f"{detail.get('cliente', '')} - {detail.get('cliente_nome', '')}".strip(" -"),
                show_name=True,
            )
        )
        self.info_entrega.setText(str(detail.get("data_entrega", "-") or "-"))
        _apply_state_chip(self.info_estado, str(detail.get("estado", "-")))
        self.info_nota.setText(str(detail.get("nota_cliente", "-") or "-"))
        transport_label = str(detail.get("nota_transporte", "") or "").strip() or "Sem transporte definido"
        transport_price = float(detail.get("preco_transporte", 0) or 0)
        if transport_price > 0:
            transport_label = f"{transport_label} | {_fmt_eur(transport_price)}"
        self.info_transporte.setText(transport_label)
        descarga_txt = str(detail.get("local_descarga", "") or "-").strip() or "-"
        zona_txt = str(detail.get("zona_transporte", "") or "").strip()
        if zona_txt:
            descarga_txt = f"{descarga_txt} | Zona {zona_txt}"
        self.info_descarga.setText(descarga_txt)
        carrier_label = str(detail.get("transportadora_nome", "") or "").strip() or "Sem transportadora externa"
        reference_label = str(detail.get("referencia_transporte", "") or "").strip()
        if reference_label:
            carrier_label = f"{carrier_label} | Ref {reference_label}"
        self.info_transportadora.setText(carrier_label)
        paletes = float(detail.get("paletes", 0) or 0)
        peso = float(detail.get("peso_bruto_kg", 0) or 0)
        volume = float(detail.get("volume_m3", 0) or 0)
        if paletes > 0 or peso > 0 or volume > 0:
            self.info_carga.setText(f"{paletes:.2f} pal | {peso:.1f} kg | {volume:.3f} m3")
        else:
            self.info_carga.setText("Carga nao definida")
        transport_cost = float(detail.get("custo_transporte", 0) or 0)
        self.info_custos.setText(f"Venda {_fmt_eur(transport_price)} | Custo {_fmt_eur(transport_cost)}")
        trip_number = str(detail.get("transporte_numero", "") or "").strip()
        trip_state = str(detail.get("estado_transporte", "") or "").strip()
        self.info_viagem.setText(f"{trip_number} | {trip_state}".strip(" |") or "Sem viagem")
        _apply_state_chip(
            self.info_cativar,
            "Concluida" if bool(detail.get("cativar")) else "Preparacao",
            "SIM" if bool(detail.get("cativar")) else "NAO",
        )
        _set_panel_tone(self.info_card, _state_tone(str(detail.get("estado", "-"))))
        self.material_rows = list(detail.get("materials_tree", []) or [])
        self._refresh_materials_table()
        self._refresh_montagem_table()
        self._refresh_reservation_info()
        self._sync_action_buttons()

    def _refresh_materials_table(self) -> None:
        _fill_table(
            self.materials_table,
            [
                [
                    row.get("material", "-"),
                    len(list(row.get("espessuras", []) or [])),
                    sum(int(esp.get("pecas", 0) or 0) for esp in list(row.get("espessuras", []) or [])),
                    row.get("estado", "-"),
                ]
                for row in self.material_rows
            ],
            align_center_from=1,
        )
        self.materials_table.setSortingEnabled(False)
        for row_index, row in enumerate(self.material_rows):
            item = self.materials_table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("material", "") or "").strip())
        for row_index, row in enumerate(self.material_rows):
            _paint_table_row(self.materials_table, row_index, str(row.get("estado", "")))
        if self.materials_table.rowCount() > 0:
            self.materials_table.selectRow(0)
            self._on_material_selected()
        else:
            self.esp_rows = []
            self.detail_pieces = []
            self.esp_table.setRowCount(0)
            self.pieces_table.setRowCount(0)

    def _on_material_selected(self) -> None:
        material_row = self._selected_material_row()
        self.esp_rows = list(material_row.get("espessuras", []) or [])
        _fill_table(
            self.esp_table,
            [
                [
                    row.get("espessura", "-"),
                    row.get("tempo_min", "0"),
                    row.get("tempo_operacoes_txt", "-"),
                    row.get("recursos_operacao_txt", "-"),
                    row.get("estado", "-"),
                ]
                for row in self.esp_rows
            ],
            align_center_from=1,
        )
        self.esp_table.setSortingEnabled(False)
        for row_index, row in enumerate(self.esp_rows):
            item = self.esp_table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("espessura", "") or "").strip())
        for row_index, row in enumerate(self.esp_rows):
            _paint_table_row(self.esp_table, row_index, str(row.get("estado", "")))
        header = self.esp_table.horizontalHeader()
        header.resizeSection(0, 96)
        header.resizeSection(1, 76)
        header.resizeSection(4, 100)
        if self.esp_table.rowCount() > 0:
            self.esp_table.selectRow(0)
            self._on_esp_selected()
        else:
            self.detail_pieces = []
            self.pieces_table.setRowCount(0)
            self._refresh_reservation_info()
            self._sync_action_buttons()

    def _on_esp_selected(self) -> None:
        material_row = self._selected_material_row()
        esp_row = self._selected_esp_row()
        material_name = str(material_row.get("material", "") or "").strip()
        esp_name = str(esp_row.get("espessura", "") or "").strip()
        pieces = []
        for row in list(self.current_detail.get("pieces", []) or []):
            if material_name and str(row.get("material", "") or "").strip() != material_name:
                continue
            if esp_name and str(row.get("espessura", "") or "").strip() != esp_name:
                continue
            pieces.append(row)
        self.detail_pieces = pieces
        _fill_table(
            self.pieces_table,
            [
                [
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("opp", "-"),
                    row.get("material", "-"),
                    row.get("espessura", "-"),
                    row.get("operacoes", "-"),
                    row.get("qtd_plan", "0"),
                    row.get("qtd_prod", "0"),
                    row.get("estado", "-"),
                ]
                for row in self.detail_pieces
            ],
            align_center_from=4,
        )
        self.pieces_table.setSortingEnabled(False)
        for row_index, row in enumerate(self.detail_pieces):
            item = self.pieces_table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("ref_interna", "") or "").strip())
        for row_index, row in enumerate(self.detail_pieces):
            _paint_table_row(self.pieces_table, row_index, str(row.get("estado", "")))
        if self.pieces_table.rowCount() > 0:
            self.pieces_table.selectRow(0)
        self._refresh_reservation_info()
        self._sync_action_buttons()

    def _refresh_reservation_info(self) -> None:
        material_row = self._selected_material_row()
        esp_row = self._selected_esp_row()
        material_name = str(material_row.get("material", "") or "").strip()
        esp_name = str(esp_row.get("espessura", "") or "").strip()
        target_rows = []
        for row in list(self.current_detail.get("reservas", []) or []):
            if material_name and str(row.get("material", "") or "").strip() != material_name:
                continue
            if esp_name and str(row.get("espessura", "") or "").strip() != esp_name:
                continue
            target_rows.append(row)
        if target_rows:
            self.info_chapa.setText(", ".join([str(row.get("material_id", "") or "-") for row in target_rows]))
            self.info_reservas.setText(" | ".join([f"{row.get('material', '-')} {row.get('espessura', '-')} -> {row.get('quantidade', '0')}" for row in target_rows]))
        else:
            self.info_chapa.setText("-")
            self.info_reservas.setText("Sem reservas ativas.")

    def _refresh_montagem_table(self) -> None:
        self.detail_montagem = list(self.current_detail.get("montagem_items", []) or [])
        shortages_map = {
            str(row.get("item_key", "") or row.get("produto_codigo", "") or row.get("stock_material_id", "") or row.get("descricao", "") or "").strip(): row
            for row in list(self.current_detail.get("montagem_shortages", []) or [])
            if str(row.get("item_key", "") or row.get("produto_codigo", "") or row.get("stock_material_id", "") or row.get("descricao", "") or "").strip()
        }
        _fill_table(
            self.montagem_table,
            [
                [
                    row.get("tipo_label", "-"),
                    row.get("produto_codigo", "") or row.get("stock_material_id", "") or row.get("material", "") or "-",
                    row.get("descricao", "-"),
                    f"{float(row.get('qtd_planeada', 0) or 0):.2f}",
                    f"{float(row.get('qtd_consumida', 0) or 0):.2f}",
                    f"{float(row.get('qtd_pendente', 0) or 0):.2f}",
                    f"{float((shortages_map.get(str(row.get('item_key', '') or row.get('produto_codigo', '') or row.get('stock_material_id', '') or row.get('descricao', '') or '').strip(), {}) or {}).get('qtd_em_falta', 0) or 0):.2f}",
                    row.get("estado", "-"),
                ]
                for row in self.detail_montagem
            ],
            align_center_from=3,
        )
        for row_index, row in enumerate(self.detail_montagem):
            _paint_table_row(self.montagem_table, row_index, str(row.get("estado", "")))
            code = str(row.get("item_key", "") or row.get("produto_codigo", "") or row.get("stock_material_id", "") or row.get("descricao", "") or "").strip()
            shortage = shortages_map.get(code, {})
            if not shortage:
                continue
            tip_parts = [
                f"Falta {float(shortage.get('qtd_em_falta', 0) or 0):.2f} {shortage.get('produto_unid', 'UN')}",
                f"Disponivel {float(shortage.get('qtd_disponivel', 0) or 0):.2f}",
            ]
            if str(shortage.get("fornecedor_sugerido", "") or "").strip():
                tip_parts.append(f"Sugestao {shortage.get('fornecedor_sugerido', '-')}")
            if str(shortage.get("ultima_compra", "") or "").strip():
                tip_parts.append(f"Ultima compra {shortage.get('ultima_compra', '-')}")
            if float(shortage.get("preco_medio", 0) or 0) > 0:
                tip_parts.append(f"Preco medio {float(shortage.get('preco_medio', 0) or 0):.4f}")
            if float(shortage.get("prazo_dias", 0) or 0) > 0:
                tip_parts.append(f"Prazo {float(shortage.get('prazo_dias', 0) or 0):g} dias")
            if str(shortage.get("alternativa_stock", "") or "").strip():
                tip_parts.append(f"Alt. stock {shortage.get('alternativa_stock', '-')}")
            for col_index in range(self.montagem_table.columnCount()):
                item = self.montagem_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(" | ".join(tip_parts))
        montagem_estado = str(self.current_detail.get("montagem_estado", "Nao aplicavel") or "Nao aplicavel").strip()
        montagem_tempo = float(self.current_detail.get("montagem_tempo_min", 0) or 0)
        montagem_resumo = str(self.current_detail.get("montagem_resumo", "") or "").strip() or "Montagem final"
        shortages = list(self.current_detail.get("montagem_shortages", []) or [])
        if shortages:
            highlights = []
            for shortage in shortages[:2]:
                chunk = f"{shortage.get('produto_codigo', '-')}: falta {float(shortage.get('qtd_em_falta', 0) or 0):.2f}"
                supplier_txt = str(shortage.get("fornecedor_sugerido", "") or "").strip()
                if supplier_txt:
                    chunk += f" | sug. {supplier_txt}"
                if float(shortage.get("preco_medio", 0) or 0) > 0:
                    chunk += f" | med. {float(shortage.get('preco_medio', 0) or 0):.2f}"
                highlights.append(chunk)
            if len(shortages) > 2:
                highlights.append(f"+{len(shortages) - 2} faltas")
            self.montagem_meta.setText(
                f"Estado {montagem_estado} | Tempo {montagem_tempo:.1f} min | {montagem_resumo} | " + " ; ".join(highlights)
            )
        elif self.detail_montagem:
            self.montagem_meta.setText(f"Estado {montagem_estado} | Tempo {montagem_tempo:.1f} min | {montagem_resumo} | Stock OK")
        else:
            self.montagem_meta.setText("Sem itens de montagem.")

    def _open_montagem_in_operator(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Operador", "Seleciona uma encomenda.")
            return
        main_window = self.window()
        if not hasattr(main_window, "show_page"):
            QMessageBox.information(self, "Operador", "Abre o menu Operador e pica a OF/componentes desta encomenda.")
            return
        try:
            main_window.show_page("operator")
            page = getattr(main_window, "pages", {}).get("operator")
            opener = getattr(page, "open_montagem_stock_group", None)
            if callable(opener) and opener(numero):
                suppress = getattr(main_window, "suppress_next_scheduled_refresh", None)
                if callable(suppress):
                    suppress("operator")
                return
        except Exception as exc:
            QMessageBox.critical(self, "Operador", str(exc))
            return
        QMessageBox.information(
            self,
            "Operador",
            f"Menu Operador aberto. Pica a OF/componentes da encomenda {numero} para fazer a baixa com registo operacional.",
        )

    def _create_montagem_purchase_note(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Montagem", "Seleciona uma encomenda.")
            return
        try:
            result = self.backend.ne_create_from_montagem_shortages([numero])
        except Exception as exc:
            QMessageBox.critical(self, "Montagem", str(exc))
            return
        note_number = str(result.get("numero", "") or "").strip()
        main_window = self.window()
        if hasattr(main_window, "show_page") and note_number:
            try:
                main_window.show_page("purchase_notes")
                page = getattr(main_window, "pages", {}).get("purchase_notes")
                if page is not None and hasattr(page, "open_note_numero"):
                    page.open_note_numero(note_number)
            except Exception:
                pass
        missing = list(result.get("missing_supplier", []) or [])
        summary = f"Nota {note_number or '-'} criada para a montagem da encomenda {numero}."
        if missing:
            summary += "\n\nFornecedor por validar em: " + ", ".join(missing)
        QMessageBox.information(self, "Montagem", summary)

    def _client_code_from_text(self, text: str) -> str:
        raw = str(text or "").strip()
        if not raw:
            return ""
        code = raw.split(" - ", 1)[0].strip()
        for row in self.client_rows:
            if code == str(row.get("codigo", "")).strip():
                return code
            if raw.lower() == str(row.get("label", "")).strip().lower():
                return str(row.get("codigo", "")).strip()
        return code

    def _header_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        is_new = not bool(str(initial.get("numero", "") or "").strip())
        dialog = QDialog(self)
        dialog.setWindowTitle("Criar ordem de fabrico" if is_new else "Editar ordem de fabrico")
        dialog.setWindowFlags(
            dialog.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )
        dialog.setMinimumSize(760, 590)
        dialog.resize(860, 680)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(12)

        heading = QLabel("Nova ordem de fabrico" if is_new else str(initial.get("of_codigo", "") or "Ordem de fabrico"))
        heading.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        subheading = QLabel(
            "Define os dados comerciais e, se pretenderes, inicia a OF diretamente a partir de um conjunto guardado."
            if is_new
            else "Atualiza os dados gerais e logísticos sem alterar a estrutura produtiva existente."
        )
        subheading.setWordWrap(True)
        subheading.setProperty("role", "muted")
        layout.addWidget(heading)
        layout.addWidget(subheading)

        tabs = QTabWidget()
        general_tab = QWidget()
        logistics_tab = QWidget()
        production_tab = QWidget()
        tabs.addTab(general_tab, "Dados gerais")
        tabs.addTab(logistics_tab, "Logística")
        tabs.addTab(production_tab, "Estrutura inicial")
        general_form = QFormLayout(general_tab)
        logistics_form = QFormLayout(logistics_tab)
        for form in (general_form, logistics_form):
            form.setContentsMargins(16, 16, 16, 16)
            form.setHorizontalSpacing(14)
            form.setVerticalSpacing(10)
            form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
            form.setFormAlignment(Qt.AlignTop)
            form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        production_layout = QVBoxLayout(production_tab)
        production_layout.setContentsMargins(16, 16, 16, 16)
        production_layout.setSpacing(12)
        type_combo = QComboBox()
        type_combo.addItems(["Cliente", "Interna (produção)"])
        type_combo.setCurrentText(str(initial.get("tipo_encomenda", "") or "Cliente"))
        client_combo = QComboBox()
        client_combo.setEditable(True)
        for row in self.client_rows:
            client_combo.addItem(str(row.get("label", "")), str(row.get("codigo", "")))
        current_client = str(initial.get("cliente", "") or "").strip()
        if current_client:
            for row_index in range(client_combo.count()):
                if str(client_combo.itemData(row_index) or "").strip() == current_client:
                    client_combo.setCurrentIndex(row_index)
                    break
            else:
                client_combo.setCurrentText(current_client)
        delivery_edit = QDateEdit()
        delivery_edit.setCalendarPopup(True)
        delivery_edit.setDisplayFormat("dd/MM/yyyy")
        delivery_edit.setMinimumHeight(34)
        delivery_edit.setToolTip("Data de entrega da encomenda.")
        delivery_raw = str(initial.get("data_entrega", "") or "").strip()
        delivery_date = QDate.fromString(delivery_raw, "yyyy-MM-dd") if delivery_raw else QDate.currentDate()
        if not delivery_date.isValid():
            delivery_date = QDate.currentDate()
        delivery_edit.setDate(delivery_date)
        calendar = delivery_edit.calendarWidget()
        if calendar is not None:
            calendar.setGridVisible(True)
            calendar.setFirstDayOfWeek(Qt.Monday)
            calendar.setStyleSheet(
                """
                QCalendarWidget QWidget { alternate-background-color: #f8fafc; }
                QCalendarWidget QToolButton {
                    color: #020617;
                    background: #eef4fb;
                    border: 1px solid #bfcee3;
                    border-radius: 6px;
                    padding: 4px 8px;
                    font-weight: 700;
                }
                QCalendarWidget QAbstractItemView {
                    selection-background-color: #eaf7da;
                    selection-color: #26331d;
                    outline: 0;
                }
                """
            )
        transport_combo = QComboBox()
        transport_combo.setEditable(True)
        transport_combo.addItems(["", "Transporte a Cargo do Cliente", "Transporte a Nosso Cargo", "Subcontratado"])
        transport_combo.setCurrentText(str(initial.get("nota_transporte", "") or "").strip())
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for supplier in list(self.backend.ne_suppliers() or []):
            carrier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        carrier_current = " - ".join(
            [
                part
                for part in [
                    str(initial.get("transportadora_id", "") or "").strip(),
                    str(initial.get("transportadora_nome", "") or "").strip(),
                ]
                if part
            ]
        ).strip(" -")
        carrier_combo.setCurrentText(carrier_current)
        zone_combo = QComboBox()
        zone_combo.setEditable(True)
        zone_combo.addItem("")
        for value in list(self.backend.transport_zone_options() or []):
            zone_combo.addItem(str(value))
        zone_combo.setCurrentText(str(initial.get("zona_transporte", "") or "").strip())
        local_descarga_edit = QLineEdit(str(initial.get("local_descarga", "") or "").strip())
        transport_price_spin = QDoubleSpinBox()
        transport_price_spin.setRange(0.0, 1000000.0)
        transport_price_spin.setDecimals(2)
        transport_price_spin.setPrefix("EUR ")
        transport_price_spin.setValue(float(initial.get("preco_transporte", 0) or 0))
        transport_cost_spin = QDoubleSpinBox()
        transport_cost_spin.setRange(0.0, 1000000.0)
        transport_cost_spin.setDecimals(2)
        transport_cost_spin.setPrefix("EUR ")
        transport_cost_spin.setValue(float(initial.get("custo_transporte", 0) or 0))
        paletes_spin = QDoubleSpinBox()
        paletes_spin.setRange(0.0, 9999.0)
        paletes_spin.setDecimals(2)
        paletes_spin.setSuffix(" pal")
        paletes_spin.setValue(float(initial.get("paletes", 0) or 0))
        peso_spin = QDoubleSpinBox()
        peso_spin.setRange(0.0, 100000.0)
        peso_spin.setDecimals(2)
        peso_spin.setSuffix(" kg")
        peso_spin.setValue(float(initial.get("peso_bruto_kg", 0) or 0))
        volume_spin = QDoubleSpinBox()
        volume_spin.setRange(0.0, 10000.0)
        volume_spin.setDecimals(3)
        volume_spin.setSuffix(" m3")
        volume_spin.setValue(float(initial.get("volume_m3", 0) or 0))
        ref_transport_edit = QLineEdit(str(initial.get("referencia_transporte", "") or "").strip())
        tempo_spin = QDoubleSpinBox()
        tempo_spin.setRange(0.0, 100000.0)
        tempo_spin.setDecimals(2)
        tempo_spin.setValue(float(initial.get("tempo_estimado", 0) or 0))
        note_edit = QLineEdit(str(initial.get("nota_cliente", "") or "").strip())
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(96)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        cativar_box = QCheckBox("Cativar MP")
        cativar_box.setChecked(bool(initial.get("cativar")))
        numero_label = QLabel(str(initial.get("numero", "(nova)") or "(nova)"))
        numero_label.setStyleSheet("font-weight: 900; color: #1d4ed8;")
        if not is_new:
            general_form.addRow("Encomenda", numero_label)
        general_form.addRow("Tipo", type_combo)
        general_form.addRow("Cliente", client_combo)
        general_form.addRow("Entrega", delivery_edit)
        general_form.addRow("Tempo estimado", tempo_spin)
        general_form.addRow("Referência cliente", note_edit)
        general_form.addRow("Observações", obs_edit)
        general_form.addRow("", cativar_box)

        logistics_form.addRow("Responsabilidade", transport_combo)
        logistics_form.addRow("Transportadora", carrier_combo)
        logistics_form.addRow("Zona", zone_combo)
        logistics_form.addRow("Local de descarga", local_descarga_edit)
        logistics_form.addRow("Preço ao cliente", transport_price_spin)
        logistics_form.addRow("Custo previsto", transport_cost_spin)
        logistics_form.addRow("Paletes", paletes_spin)
        logistics_form.addRow("Peso bruto", peso_spin)
        logistics_form.addRow("Volume", volume_spin)
        logistics_form.addRow("Referência logística", ref_transport_edit)

        initial_import: dict[str, Any] = {}
        if is_new:
            available_models = list(self.backend.order_model_options() or [])
            source_combo = QComboBox()
            source_combo.addItem("Criar OF vazia", "")
            if any(str(row.get("origem_tipo", "")) == "conjunto" for row in available_models):
                source_combo.addItem("Importar conjunto guardado", "conjunto")
            if any(str(row.get("origem_tipo", "")) == "modelo" for row in available_models):
                source_combo.addItem("Importar modelo reutilizável", "modelo")
            model_combo = QComboBox()
            model_combo.setEditable(True)
            model_combo.setMinimumWidth(480)
            quantity_spin = QDoubleSpinBox()
            quantity_spin.setRange(0.01, 100000.0)
            quantity_spin.setDecimals(2)
            quantity_spin.setValue(1.0)
            quantity_spin.setSuffix(" conj.")
            model_summary = QLabel("A OF será criada sem peças. Poderás adicioná-las depois.")
            model_summary.setWordWrap(True)
            model_summary.setStyleSheet(
                "padding: 12px; background: #f8fafc; border: 1px solid #d6e0eb;"
                "border-radius: 6px; color: #334155;"
            )

            source_form = QFormLayout()
            source_form.setHorizontalSpacing(14)
            source_form.setVerticalSpacing(10)
            source_form.addRow("Origem da estrutura", source_combo)
            source_form.addRow("Conjunto / modelo", model_combo)
            source_form.addRow("Quantidade", quantity_spin)
            production_layout.addLayout(source_form)
            production_layout.addWidget(model_summary)
            production_layout.addStretch(1)

            def refresh_model_options() -> None:
                source = str(source_combo.currentData() or "").strip()
                current_code = str((model_combo.currentData() or {}).get("codigo", "") or "").strip()
                model_combo.blockSignals(True)
                model_combo.clear()
                for row in available_models:
                    if source and str(row.get("origem_tipo", "") or "").strip() == source:
                        model_combo.addItem(str(row.get("label", "") or row.get("codigo", "")), dict(row))
                if current_code:
                    for index in range(model_combo.count()):
                        if str(dict(model_combo.itemData(index) or {}).get("codigo", "") or "").strip() == current_code:
                            model_combo.setCurrentIndex(index)
                            break
                model_combo.setEnabled(bool(source))
                quantity_spin.setEnabled(bool(source))
                model_combo.blockSignals(False)
                refresh_model_summary()

            def refresh_model_summary() -> None:
                source = str(source_combo.currentData() or "").strip()
                selected = dict(model_combo.currentData() or {})
                if not source:
                    model_summary.setText("A OF será criada sem peças. Poderás adicionar referências ou conjuntos depois.")
                    return
                if not selected:
                    model_summary.setText("Seleciona um conjunto válido.")
                    return
                item_count = int(selected.get("itens", 0) or 0)
                pieces = int(selected.get("pecas", 0) or 0)
                final_value = float(selected.get("total_final", 0) or 0) * float(quantity_spin.value() or 0)
                final_txt = f"{final_value:,.2f}".replace(",", " ").replace(".", ",")
                model_summary.setText(
                    f"{selected.get('codigo', '-')} · {selected.get('descricao', '-')}\n"
                    f"{item_count} linhas · {pieces} peças fabricadas · "
                    f"valor previsto EUR {final_txt}"
                )

            source_combo.currentIndexChanged.connect(refresh_model_options)
            model_combo.currentIndexChanged.connect(refresh_model_summary)
            quantity_spin.valueChanged.connect(refresh_model_summary)
            refresh_model_options()
        else:
            current_sheets = list(initial.get("produto_fichas", []) or [])
            production_info = QLabel(
                f"Esta OF já existe e contém {len(current_sheets)} conjunto(s) associado(s).\n"
                "Usa o botão “Importar conjunto” no detalhe da encomenda para acrescentar nova estrutura produtiva."
            )
            production_info.setWordWrap(True)
            production_info.setStyleSheet(
                "padding: 14px; background: #f8fafc; border: 1px solid #d6e0eb;"
                "border-radius: 6px; color: #334155;"
            )
            production_layout.addWidget(production_info)
            production_layout.addStretch(1)

        layout.addWidget(tabs, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Criar OF" if is_new else "Guardar alterações")

        def accept_header() -> None:
            if not self._client_code_from_text(client_combo.currentText()):
                tabs.setCurrentWidget(general_tab)
                QMessageBox.warning(dialog, "Ordem de fabrico", "Seleciona um cliente.")
                return
            if is_new and str(source_combo.currentData() or "").strip() and not dict(model_combo.currentData() or {}):
                tabs.setCurrentWidget(production_tab)
                QMessageBox.warning(dialog, "Ordem de fabrico", "Seleciona o conjunto a importar.")
                return
            dialog.accept()

        buttons.accepted.connect(accept_header)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        if is_new and str(source_combo.currentData() or "").strip():
            selected_model = dict(model_combo.currentData() or {})
            initial_import = {
                "codigo": str(selected_model.get("codigo", "") or "").strip(),
                "source": str(source_combo.currentData() or "").strip(),
                "quantity": quantity_spin.value(),
            }
        return {
            "numero": str(initial.get("numero", "") or "").strip(),
            "tipo_encomenda": type_combo.currentText().strip(),
            "cliente": self._client_code_from_text(client_combo.currentText()),
            "nota_transporte": transport_combo.currentText().strip(),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "zona_transporte": zone_combo.currentText().strip(),
            "local_descarga": local_descarga_edit.text().strip(),
            "preco_transporte": transport_price_spin.value(),
            "custo_transporte": transport_cost_spin.value(),
            "paletes": paletes_spin.value(),
            "peso_bruto_kg": peso_spin.value(),
            "volume_m3": volume_spin.value(),
            "referencia_transporte": ref_transport_edit.text().strip(),
            "data_entrega": delivery_edit.date().toString("yyyy-MM-dd").strip(),
            "tempo_estimado": tempo_spin.value(),
            "nota_cliente": note_edit.text().strip(),
            "observacoes": obs_edit.toPlainText().strip(),
            "cativar": cativar_box.isChecked(),
            "_initial_import": initial_import,
        }

    def _pick_combo_value(self, title: str, label: str, values: list[str], current: str = "") -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        combo = QComboBox()
        combo.setEditable(True)
        for value in values:
            combo.addItem(str(value))
        combo.setCurrentText(str(current or ""))
        form.addRow(label, combo)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return combo.currentText().strip()

    def _piece_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Peca")
        dialog.setMinimumWidth(760)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        current_order = str(self.current_detail.get("numero", "") or "").strip()
        current_client = str(self.current_detail.get("cliente", "") or "").strip()
        references = self.backend.order_reference_rows("", current_client)
        refs_by_ext = {str(row.get("ref_externa", "")).strip(): row for row in references if str(row.get("ref_externa", "")).strip()}
        refs_by_int = {str(row.get("ref_interna", "")).strip(): row for row in references if str(row.get("ref_interna", "")).strip()}
        ref_history = QComboBox()
        ref_history.setEditable(True)
        ref_history.addItem("")
        for row in references:
            label = f"{row.get('ref_interna', '')} | {row.get('ref_externa', '')} | {row.get('descricao', '')}".strip()
            ref_history.addItem(label, row)
        ref_int_edit = QLineEdit(str(initial.get("ref_interna", "") or "").strip())
        ref_ext_edit = QLineEdit(str(initial.get("ref_externa", "") or "").strip())
        desc_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
        tipo_combo = QComboBox()
        tipo_combo.addItems(["CHAPA", "PERFIL", "TUBO", "OUTROS"])
        tipo_combo.setCurrentText(str(initial.get("tipo_material", "") or "CHAPA").strip().upper())
        material_combo = QComboBox()
        material_combo.setEditable(True)
        for value in list(self.presets.get("materiais", []) or []):
            material_combo.addItem(str(value))
        material_combo.setCurrentText(str(initial.get("material", "") or "").strip())
        esp_combo = QComboBox()
        esp_combo.setEditable(True)
        for value in list(self.presets.get("espessuras", []) or []):
            esp_combo.addItem(str(value))
        esp_combo.setCurrentText(str(initial.get("espessura", "") or "").strip())
        dimensao_edit = QLineEdit(str(initial.get("dimensao", initial.get("dimensoes", "")) or "").strip())
        profile_combo = QComboBox()
        profile_combo.addItems([
            "IPE", "IPN", "HEA", "HEB", "HEM",
            "UPN", "UPE", "U", "I", "C", "T", "Z", "Omega",
            "Cantoneira L abas iguais", "Cantoneira L abas desiguais",
            "Barra chata", "Barra quadrada", "Barra redonda", "Barra retangular", "Barra sextavada",
        ])
        profile_combo.setEditable(True)
        profile_combo.setCurrentText(str(initial.get("perfil_tipo", "") or "IPE").strip())
        profile_size_edit = QLineEdit(str(initial.get("perfil_tamanho", "") or "").strip())
        profile_length_spin = QDoubleSpinBox()
        profile_length_spin.setRange(0.0, 100000.0)
        profile_length_spin.setDecimals(1)
        profile_length_spin.setSuffix(" mm")
        profile_length_spin.setValue(float(initial.get("comprimento_mm", 0) or 0))
        tube_shape_combo = QComboBox()
        tube_shape_combo.addItems(["Retangular / Quadrado", "Redondo"])
        tube_shape_combo.setCurrentText(str(initial.get("tubo_forma", "") or "Retangular / Quadrado").strip())
        side_a_spin = QDoubleSpinBox()
        side_a_spin.setRange(0.0, 100000.0)
        side_a_spin.setDecimals(1)
        side_a_spin.setSuffix(" mm")
        side_a_spin.setValue(float(initial.get("lado_a", 0) or 0))
        side_b_spin = QDoubleSpinBox()
        side_b_spin.setRange(0.0, 100000.0)
        side_b_spin.setDecimals(1)
        side_b_spin.setSuffix(" mm")
        side_b_spin.setValue(float(initial.get("lado_b", 0) or 0))
        tube_thick_spin = QDoubleSpinBox()
        tube_thick_spin.setRange(0.0, 10000.0)
        tube_thick_spin.setDecimals(1)
        tube_thick_spin.setSuffix(" mm")
        tube_thick_spin.setValue(float(initial.get("tubo_espessura", 0) or 0))
        diameter_spin = QDoubleSpinBox()
        diameter_spin.setRange(0.0, 100000.0)
        diameter_spin.setDecimals(1)
        diameter_spin.setSuffix(" mm")
        diameter_spin.setValue(float(initial.get("diametro", 0) or 0))
        operation_selector, operacoes_edit, apply_operations = _build_operation_selector(
            list(self.presets.get("operacoes", []) or []),
            str(initial.get("operacoes", self.presets.get("operacao_default", "Embalamento")) or self.presets.get("operacao_default", "Embalamento")).strip(),
        )
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd_plan", initial.get("quantidade_pedida", 1)) or 1))
        preco_spin = QDoubleSpinBox()
        preco_spin.setRange(0.0, 1000000.0)
        preco_spin.setDecimals(4)
        preco_spin.setValue(float(initial.get("preco_unit", 0) or 0))
        drawing_edit = QLineEdit(str(initial.get("desenho_path", "") or "").strip())
        files_edit = QLineEdit("; ".join(str(item or "").strip() for item in list(initial.get("ficheiros", []) or []) if str(item or "").strip()))
        keep_ref_box = QCheckBox("Guardar referencia na base")
        keep_ref_box.setChecked(bool(initial.get("guardar_ref", True)))

        def apply_reference(payload: dict | None) -> None:
            if not isinstance(payload, dict):
                return
            ref_ext_edit.setText(str(payload.get("ref_externa", "") or "").strip())
            ref_int_edit.setText(str(payload.get("ref_interna", "") or "").strip())
            desc_edit.setText(str(payload.get("descricao", "") or "").strip())
            material_combo.setCurrentText(str(payload.get("material", "") or "").strip())
            tipo_combo.setCurrentText(str(payload.get("tipo_material", "") or payload.get("material_family", "") or "CHAPA").strip().upper())
            esp_combo.setCurrentText(str(payload.get("espessura", "") or "").strip())
            dimensao_edit.setText(str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip())
            profile_combo.setCurrentText(str(payload.get("perfil_tipo", "") or "IPE").strip())
            profile_size_edit.setText(str(payload.get("perfil_tamanho", "") or "").strip())
            apply_operations(str(payload.get("operacoes", "") or self.presets.get("operacao_default", "Embalamento")).strip())
            preco_spin.setValue(float(payload.get("preco", 0) or 0))
            if not drawing_edit.text().strip():
                drawing_edit.setText(str(payload.get("desenho", "") or "").strip())

        def load_from_ref_text() -> None:
            selected = ref_history.currentData()
            if isinstance(selected, dict):
                apply_reference(selected)
                return
            raw = ref_ext_edit.text().strip() or ref_history.currentText().strip()
            apply_reference(refs_by_ext.get(raw) or refs_by_int.get(raw))

        def browse_refs() -> None:
            payload = _reference_catalog_dialog(
                self,
                self.backend.order_reference_rows("", ""),
                "Histórico de referencias",
                backend=self.backend,
                current_client=current_client,
            )
            apply_reference(payload)

        def generate_ref() -> None:
            ref_int_edit.setText(self.backend.order_suggest_ref_interna(current_order, current_client))

        def pick_drawing() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar desenho",
                "",
                "Desenhos (*.pdf *.dwg *.dxf *.step *.stp *.iges *.igs *.png *.jpg *.jpeg *.bmp);;Todos (*.*)",
            )
            if path:
                drawing_edit.setText(path)

        def pick_files() -> None:
            paths, _ = QFileDialog.getOpenFileNames(
                self,
                "Selecionar ficheiros associados",
                "",
                "Ficheiros técnicos (*.pdf *.dwg *.dxf *.step *.stp *.iges *.igs *.png *.jpg *.jpeg *.bmp);;Todos (*.*)",
            )
            if paths:
                current = [part.strip() for part in files_edit.text().split(";") if part.strip()]
                for path in paths:
                    if path not in current:
                        current.append(path)
                files_edit.setText("; ".join(current))

        def _spin_text(spin: QDoubleSpinBox) -> str:
            value = float(spin.value() or 0)
            if value <= 0:
                return ""
            return f"{value:g}"

        def update_dimension_from_guided_fields() -> None:
            tipo = tipo_combo.currentText().strip().upper()
            if tipo == "PERFIL":
                parts = [profile_combo.currentText().strip(), profile_size_edit.text().strip()]
                generated = " ".join(part for part in parts if part)
                if profile_length_spin.value() > 0:
                    generated = f"{generated} L={profile_length_spin.value():g}mm".strip()
                if generated:
                    dimensao_edit.setText(generated)
                return
            if tipo == "TUBO":
                shape = tube_shape_combo.currentText().strip().lower()
                thick = _spin_text(tube_thick_spin)
                if "redondo" in shape:
                    dia = _spin_text(diameter_spin)
                    generated = f"Ø{dia}x{thick}" if dia and thick else (f"Ø{dia}" if dia else "")
                else:
                    lado_a = _spin_text(side_a_spin)
                    lado_b = _spin_text(side_b_spin) or lado_a
                    generated = f"{lado_a}x{lado_b}x{thick}" if lado_a and lado_b and thick else ""
                if generated:
                    dimensao_edit.setText(generated)

        def sync_material_fields() -> None:
            tipo = tipo_combo.currentText().strip().upper()
            is_profile = tipo == "PERFIL"
            is_tube = tipo == "TUBO"
            is_round_tube = is_tube and "redondo" in tube_shape_combo.currentText().strip().lower()
            esp_combo.setEnabled(tipo in {"CHAPA", "TUBO", "OUTROS"})
            profile_combo.setVisible(is_profile)
            profile_size_edit.setVisible(is_profile)
            profile_length_spin.setVisible(is_profile)
            tube_shape_combo.setVisible(is_tube)
            side_a_spin.setVisible(is_tube and not is_round_tube)
            side_b_spin.setVisible(is_tube and not is_round_tube)
            tube_thick_spin.setVisible(is_tube)
            diameter_spin.setVisible(is_round_tube)
            for field in (profile_combo, profile_size_edit, profile_length_spin, tube_shape_combo, side_a_spin, side_b_spin, tube_thick_spin, diameter_spin):
                label = form.labelForField(field)
                if label is not None:
                    label.setVisible(field.isVisible())
            if tipo == "PERFIL" and esp_combo.currentText().strip() == "":
                esp_combo.setCurrentText("-")
            update_dimension_from_guided_fields()

        ref_history.currentIndexChanged.connect(lambda _index: apply_reference(ref_history.currentData()))
        ref_buttons = QHBoxLayout()
        btn_generate = QPushButton("Gerar")
        btn_generate.setProperty("variant", "secondary")
        btn_generate.clicked.connect(generate_ref)
        btn_history = QPushButton("Referencias criadas")
        btn_history.setProperty("variant", "secondary")
        btn_history.clicked.connect(browse_refs)
        btn_load = QPushButton("Carregar ref.")
        btn_load.setProperty("variant", "secondary")
        btn_load.clicked.connect(load_from_ref_text)
        ref_buttons.addWidget(btn_generate)
        ref_buttons.addWidget(btn_history)
        ref_buttons.addWidget(btn_load)
        ref_buttons.addStretch(1)
        draw_buttons = QHBoxLayout()
        btn_pick_draw = QPushButton("Selecionar desenho")
        btn_pick_draw.setProperty("variant", "secondary")
        btn_pick_draw.clicked.connect(pick_drawing)
        btn_pick_files = QPushButton("Adicionar ficheiros")
        btn_pick_files.setProperty("variant", "secondary")
        btn_pick_files.clicked.connect(pick_files)
        draw_buttons.addWidget(btn_pick_draw)
        draw_buttons.addWidget(btn_pick_files)
        draw_buttons.addStretch(1)
        defaults_btn = QPushButton("Operacao padrao")
        defaults_btn.setProperty("variant", "secondary")
        defaults_btn.clicked.connect(lambda: apply_operations(str(self.presets.get("operacao_default", "Embalamento"))))

        form.addRow("Histórico", ref_history)
        form.addRow("Ref. interna", ref_int_edit)
        form.addRow("", ref_buttons)
        form.addRow("Ref. externa", ref_ext_edit)
        form.addRow("Descricao", desc_edit)
        form.addRow("Tipo material", tipo_combo)
        form.addRow("Subtipo material", material_combo)
        form.addRow("Espessura", esp_combo)
        form.addRow("Tipo perfil", profile_combo)
        form.addRow("Tamanho perfil", profile_size_edit)
        form.addRow("Comprimento", profile_length_spin)
        form.addRow("Tipo tubo", tube_shape_combo)
        form.addRow("Lado A", side_a_spin)
        form.addRow("Lado B", side_b_spin)
        form.addRow("Esp. tubo", tube_thick_spin)
        form.addRow("Diametro", diameter_spin)
        form.addRow("Dimensao", dimensao_edit)
        form.addRow("Operacoes", operation_selector)
        form.addRow("", defaults_btn)
        form.addRow("Quantidade", qtd_spin)
        form.addRow("Preco unit.", preco_spin)
        form.addRow("Desenho", drawing_edit)
        form.addRow("Ficheiros associados", files_edit)
        form.addRow("", draw_buttons)
        form.addRow("", keep_ref_box)
        layout.addLayout(form)
        if not ref_int_edit.text().strip():
            generate_ref()
        tipo_combo.currentTextChanged.connect(lambda _text: sync_material_fields())
        tube_shape_combo.currentTextChanged.connect(lambda _text: sync_material_fields())
        for widget in (side_a_spin, side_b_spin, tube_thick_spin, diameter_spin, profile_length_spin):
            widget.valueChanged.connect(lambda _value: update_dimension_from_guided_fields())
        profile_combo.currentTextChanged.connect(lambda _text: update_dimension_from_guided_fields())
        profile_size_edit.textChanged.connect(lambda _text: update_dimension_from_guided_fields())
        sync_material_fields()
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "ref_interna": ref_int_edit.text().strip(),
            "ref_externa": ref_ext_edit.text().strip(),
            "descricao": desc_edit.text().strip(),
            "tipo_material": tipo_combo.currentText().strip(),
            "material": material_combo.currentText().strip(),
            "subtipo_material": material_combo.currentText().strip(),
            "espessura": esp_combo.currentText().strip(),
            "dimensao": dimensao_edit.text().strip(),
            "perfil_tipo": profile_combo.currentText().strip(),
            "perfil_tamanho": profile_size_edit.text().strip(),
            "comprimento_mm": profile_length_spin.value(),
            "tubo_forma": tube_shape_combo.currentText().strip(),
            "lado_a": side_a_spin.value(),
            "lado_b": side_b_spin.value(),
            "tubo_espessura": tube_thick_spin.value(),
            "diametro": diameter_spin.value(),
            "operacoes": operacoes_edit.text().strip(),
            "quantidade_pedida": qtd_spin.value(),
            "preco_unit": preco_spin.value(),
            "desenho": drawing_edit.text().strip(),
            "ficheiros": [part.strip() for part in files_edit.text().split(";") if part.strip()],
            "guardar_ref": keep_ref_box.isChecked(),
        }

    def _reserve_dialog(self, candidates: list[dict]) -> list[dict] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Cativar stock")
        dialog.setMinimumSize(1020, 440)
        dialog.setStyleSheet(
            """
            QDialog { font-size: 12px; }
            QTableWidget { font-size: 12px; }
            QHeaderView::section { font-size: 12px; font-weight: 700; }
            QDoubleSpinBox {
                font-size: 12px;
                min-height: 32px;
                padding: 3px 8px;
            }
            QDoubleSpinBox#reserveCellSpin {
                border: none;
                background: transparent;
                min-height: 26px;
                padding: 0 8px;
            }
            QPushButton { min-width: 116px; min-height: 36px; font-size: 12px; }
            """
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)
        table = QTableWidget(len(candidates), 5)
        table.setHorizontalHeaderLabels(["Dimensao", "Disponivel", "Local", "Lote", "Reservar"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.verticalHeader().setDefaultSectionSize(38)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.resizeSection(1, 110)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.resizeSection(2, 120)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.resizeSection(4, 196)
        spinners = []
        for row_index, row in enumerate(candidates):
            table.setItem(row_index, 0, QTableWidgetItem(str(row.get("dimensao", "-"))))
            table.setItem(row_index, 1, QTableWidgetItem(str(row.get("disponivel", "0"))))
            table.setItem(row_index, 2, QTableWidgetItem(str(row.get("local", "-"))))
            table.setItem(row_index, 3, QTableWidgetItem(str(row.get("lote", "-"))))
            spin = QDoubleSpinBox()
            spin.setRange(0.0, float(row.get("disponivel", 0) or 0))
            spin.setDecimals(2)
            spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
            spin.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            spin.setFrame(False)
            spin.setObjectName("reserveCellSpin")
            table.setCellWidget(row_index, 4, spin)
            spinners.append((row, spin))
        table.setMinimumHeight(_table_visible_height(table, max(6, len(candidates)), extra=20))
        layout.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        for button in buttons.buttons():
            button.setMinimumWidth(110)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        out = []
        for row, spin in spinners:
            if spin.value() > 0:
                out.append({"material_id": row.get("material_id"), "quantidade": spin.value()})
        return out

    def _new_order(self) -> None:
        payload = self._header_dialog({})
        if payload is None:
            return
        initial_import = dict(payload.pop("_initial_import", {}) or {})
        try:
            detail = self.backend.order_create_with_models(
                payload,
                [initial_import] if initial_import else [],
            )
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(str(detail.get("numero", "") or "").strip())
        if initial_import:
            QMessageBox.information(
                self,
                "Ordem de fabrico criada",
                (
                    f"{detail.get('of_codigo', 'OF')} criada com o conjunto {initial_import.get('codigo', '')}.\n"
                    f"Peças fabricadas: {int(detail.get('imported_pieces', 0) or 0)} · "
                    f"componentes/serviços: {int(detail.get('imported_items', 0) or 0)}"
                ),
            )

    def _edit_order_header(self) -> None:
        if not self.current_detail.get("numero"):
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        payload = self._header_dialog(self.current_detail)
        if payload is None:
            return
        payload.pop("_initial_import", None)
        try:
            detail = self.backend.order_create_or_update(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(str(detail.get("numero", "") or "").strip())

    def _remove_order(self) -> None:
        row = self._selected_order_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        if QMessageBox.question(self, "Apagar encomenda", f"Remover encomenda {numero}?") != QMessageBox.Yes:
            return
        try:
            self.backend.order_remove(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()

    def _import_model(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        try:
            models = list(self.backend.order_model_options() or [])
        except Exception as exc:
            QMessageBox.critical(self, "Carregar modelo", str(exc))
            return
        if not models:
            QMessageBox.information(self, "Carregar modelo", "Nao existem modelos/conjuntos ativos.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Carregar modelo / conjunto")
        dialog.setMinimumWidth(520)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        model_combo = QComboBox()
        for row in models:
            model_combo.addItem(str(row.get("label", "") or row.get("codigo", "")), row)
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 100000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(1.0)
        form.addRow("Modelo/conjunto", model_combo)
        form.addRow("Quantidade", qty_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dict(model_combo.currentData() or {})
        try:
            self.backend.order_import_model(
                numero,
                str(selected.get("codigo", "") or "").strip(),
                qty_spin.value(),
                str(selected.get("origem_tipo", "modelo") or "modelo"),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Carregar modelo", str(exc))
            return
        self.refresh()
        self._select_order(numero)

    def _fabrication_group_options(self) -> list[dict[str, Any]]:
        groups: dict[tuple[str, str], dict[str, Any]] = {}
        for piece in list(self.current_detail.get("pieces", []) or []):
            material = str(piece.get("material", "") or "").strip()
            espessura = str(piece.get("espessura", "") or "").strip()
            if not material or not espessura:
                continue
            key = (material, espessura)
            row = groups.setdefault(key, {"material": material, "espessura": espessura, "pecas": 0, "qtd": 0.0})
            row["pecas"] = int(row.get("pecas", 0) or 0) + 1
            try:
                row["qtd"] = float(row.get("qtd", 0) or 0) + float(piece.get("qtd_plan", 0) or 0)
            except Exception:
                pass
        def thickness_sort(value: Any) -> float:
            try:
                return float(str(value or "0").replace(",", "."))
            except Exception:
                return 0.0

        return sorted(groups.values(), key=lambda row: (str(row.get("material", "")).lower(), thickness_sort(row.get("espessura", "0"))))

    def _pick_fabrication_groups(self) -> list[dict[str, Any]] | None:
        options = self._fabrication_group_options()
        if not options:
            QMessageBox.warning(self, "Ordem de Fabrico", "Esta encomenda não tem espessuras para imprimir.")
            return None
        dialog = QDialog(self)
        dialog.setWindowTitle("Previsualizar OF por espessura")
        dialog.setMinimumWidth(520)
        layout = QVBoxLayout(dialog)
        title = QLabel("Seleciona as espessuras/material a incluir na Ordem de Fabrico.")
        title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        hint = QLabel("Mantém todas selecionadas para gerar a OF completa, ou deixa apenas as espessuras que vão para esta máquina.")
        hint.setWordWrap(True)
        hint.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(hint)
        checks: list[tuple[QCheckBox, dict[str, Any]]] = []
        for row in options:
            label = (
                f"{row.get('material', '-')} | {row.get('espessura', '-')} mm"
                f"   ·   {int(row.get('pecas', 0) or 0)} referência(s)"
                f"   ·   qtd {float(row.get('qtd', 0) or 0):.0f}"
            )
            check = QCheckBox(label)
            check.setChecked(True)
            check.setStyleSheet("font-size: 12px; font-weight: 700; padding: 4px;")
            checks.append((check, row))
            layout.addWidget(check)
        actions = QHBoxLayout()
        all_btn = QPushButton("Selecionar todas")
        none_btn = QPushButton("Limpar")
        all_btn.setProperty("variant", "secondary")
        none_btn.setProperty("variant", "secondary")
        all_btn.clicked.connect(lambda: [check.setChecked(True) for check, _row in checks])
        none_btn.clicked.connect(lambda: [check.setChecked(False) for check, _row in checks])
        actions.addWidget(all_btn)
        actions.addWidget(none_btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Previsualizar selecionadas")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected = [dict(row) for check, row in checks if check.isChecked()]
        if not selected:
            QMessageBox.warning(self, "Ordem de Fabrico", "Seleciona pelo menos uma espessura.")
            return None
        if len(selected) == len(options):
            return []
        return selected

    def _print_fabrication_order(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Ordem de Fabrico", "Seleciona uma encomenda.")
            return
        selected_groups = self._pick_fabrication_groups()
        if selected_groups is None:
            return
        try:
            print_fn = getattr(self.backend, "order_print_fabrication_pdf", None)
            if callable(print_fn):
                path = print_fn(numero, selected_groups=selected_groups or None)
            else:
                path = self.backend.order_open_fabrication_pdf(numero, selected_groups=selected_groups or None)
        except Exception as exc:
            QMessageBox.critical(self, "Ordem de Fabrico", str(exc))
            return
        scope = "completa" if not selected_groups else f"{len(selected_groups)} espessura(s)"
        QMessageBox.information(
            self,
            "Ordem de Fabrico",
            f"OF {scope} criada para previsualização:\n{path}\n\nPodes guardar ou imprimir a partir do visualizador de PDF.",
        )

    def _add_material(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        material = self._pick_combo_value("Adicionar material", "Material", list(self.presets.get("materiais", []) or []))
        if material is None:
            return
        try:
            self.backend.order_material_add(numero, material)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material)

    def _remove_material(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        if not numero or not material:
            QMessageBox.warning(self, "Encomendas", "Seleciona um material.")
            return
        if QMessageBox.question(self, "Remover material", f"Remover material {material}?") != QMessageBox.Yes:
            return
        try:
            self.backend.order_material_remove(numero, material)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero)

    def _add_espessura(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        if not numero or not material:
            QMessageBox.warning(self, "Encomendas", "Seleciona primeiro um material.")
            return
        esp = self._pick_combo_value("Adicionar espessura", "Espessura (mm)", list(self.presets.get("espessuras", []) or []))
        if esp is None:
            return
        try:
            self.backend.order_espessura_add(numero, material, esp)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material, espessura=esp)

    def _remove_espessura(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        espessura = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if not numero or not material or not espessura:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma espessura.")
            return
        if QMessageBox.question(self, "Remover espessura", f"Remover {material} {espessura} mm?") != QMessageBox.Yes:
            return
        try:
            self.backend.order_espessura_remove(numero, material, espessura)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material)

    def _edit_esp_time(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        esp_row = self._selected_esp_row()
        espessura = str(esp_row.get("espessura", "") or "").strip()
        if not numero or not material or not espessura:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma espessura.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Tempos por operação")
        dialog.setMinimumWidth(430)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Define os minutos e o recurso real de cada operação. A encomenda guarda essa associação e o planeamento usa-a automaticamente."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)
        form = QFormLayout()
        operation_names = [str(op or "").strip() for op in list(esp_row.get("operacoes_planeamento", []) or []) if str(op or "").strip()]
        operation_names = [op for op in operation_names if op != "Montagem"]
        op_times = dict(esp_row.get("tempos_operacao", {}) or {})
        if str(esp_row.get("tempo_min", "") or "").strip() and "Corte Laser" not in operation_names:
            operation_names.insert(0, "Corte Laser")
        if not operation_names:
            operation_names = ["Corte Laser"]
        ordered_ops: list[str] = []
        seen_ops: set[str] = set()
        for op_name in operation_names:
            if op_name in seen_ops:
                continue
            seen_ops.add(op_name)
            ordered_ops.append(op_name)
        editors: dict[str, QDoubleSpinBox] = {}
        resource_editors: dict[str, QComboBox] = {}
        current_resources = dict(esp_row.get("maquinas_operacao", {}) or {})
        for op_name in ordered_ops:
            spin = QDoubleSpinBox()
            spin.setRange(0, 100000)
            spin.setDecimals(0)
            raw_value = str(op_times.get(op_name, "") or "").strip()
            if not raw_value and op_name == "Corte Laser":
                raw_value = str(esp_row.get("tempo_min", "") or "").strip()
            try:
                spin.setValue(float(raw_value or 0))
            except Exception:
                spin.setValue(0)
            spin.setSuffix(" min")
            editors[op_name] = spin
            resource_combo = QComboBox()
            resource_combo.setEditable(False)
            for value in list(self.backend.workcenter_resource_options(op_name) or []):
                resource_combo.addItem(str(value))
            preferred_resource = str(current_resources.get(op_name, "") or "").strip()
            if preferred_resource:
                resource_combo.setCurrentText(preferred_resource)
            elif resource_combo.count() > 0:
                resource_combo.setCurrentText(str(self.backend.workcenter_default_resource(op_name) or resource_combo.itemText(0) or "").strip())
            row_host = QWidget()
            row_layout = QHBoxLayout(row_host)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)
            row_layout.addWidget(spin, 1)
            row_layout.addWidget(resource_combo, 1)
            resource_editors[op_name] = resource_combo
            form.addRow(f"{op_name}", row_host)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            payload = {}
            payload_resources = {}
            for op_name, spin in editors.items():
                minutes = int(spin.value())
                if minutes > 0:
                    payload[op_name] = str(minutes)
                    payload_resources[op_name] = resource_editors[op_name].currentText().strip()
            self.backend.order_espessura_set_operation_times(numero, material, espessura, payload, payload_resources)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material, espessura=espessura)

    def _add_piece(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        espessura = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        payload = self._piece_dialog(
            {
                "material": material,
                "espessura": espessura,
                "operacoes": str(self.presets.get("operacao_default", "Embalamento")),
                "qtd_plan": 1,
                "ref_interna": self.backend.order_suggest_ref_interna(numero, str(self.current_detail.get("cliente", "") or "").strip()),
            }
        )
        if payload is None:
            return
        try:
            self.backend.order_piece_create_or_update(numero, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, piece_ref=str(payload.get("ref_interna", "") or "").strip(), material=str(payload.get("material", "") or "").strip(), espessura=str(payload.get("espessura", "") or "").strip())

    def _edit_piece(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        piece = self._selected_piece_row()
        ref_int = str(piece.get("ref_interna", "") or "").strip()
        if not numero or not ref_int:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma peca.")
            return
        payload = self._piece_dialog(piece)
        if payload is None:
            return
        try:
            self.backend.order_piece_create_or_update(numero, payload, current_ref_interna=ref_int)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero, piece_ref=str(payload.get("ref_interna", "") or ref_int).strip(), material=str(payload.get("material", "") or "").strip(), espessura=str(payload.get("espessura", "") or "").strip())

    def _remove_piece(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        piece = self._selected_piece_row()
        ref_int = str(piece.get("ref_interna", "") or "").strip()
        if not numero or not ref_int:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma peca.")
            return
        if QMessageBox.question(self, "Remover peca", f"Remover peca {ref_int}?") != QMessageBox.Yes:
            return
        try:
            self.backend.order_piece_remove(numero, ref_int)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(numero)

    def _reserve_stock(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        espessura = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if not numero or not material or not espessura:
            QMessageBox.warning(self, "Encomendas", "Seleciona material e espessura.")
            return
        try:
            candidates = self.backend.order_stock_candidates(numero, material, espessura)
        except Exception as exc:
            QMessageBox.critical(self, "Cativar MP", str(exc))
            return
        if not candidates:
            QMessageBox.information(self, "Cativar MP", f"Sem stock disponivel para {material} esp. {espessura}.")
            return
        allocations = self._reserve_dialog(candidates)
        if allocations is None:
            return
        try:
            self.backend.order_reserve_stock(numero, material, espessura, allocations)
        except Exception as exc:
            QMessageBox.critical(self, "Cativar MP", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material, espessura=espessura)

    def _release_stock(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        material = str(self._selected_material_row().get("material", "") or "").strip()
        espessura = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if not numero or not material or not espessura:
            QMessageBox.warning(self, "Encomendas", "Seleciona material e espessura.")
            return
        if QMessageBox.question(self, "Descativar MP", f"Libertar reservas de {material} esp. {espessura}?") != QMessageBox.Yes:
            return
        try:
            self.backend.order_release_stock(numero, material, espessura)
        except Exception as exc:
            QMessageBox.critical(self, "Descativar MP", str(exc))
            return
        self.refresh()
        self._select_order(numero, material=material, espessura=espessura)

    def _select_order(self, numero: str, piece_ref: str = "", material: str = "", espessura: str = "") -> None:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return
        target_row = -1
        for row_index in range(self.table.rowCount()):
            item = self.table.item(row_index, 0)
            if item is not None and str(item.data(Qt.UserRole) or item.text() or "").strip() == numero_txt:
                target_row = row_index
                break
        if target_row < 0:
            self.refresh()
            return
        self.table.selectRow(target_row)
        self._on_order_selected()
        if material:
            for row_index in range(self.materials_table.rowCount()):
                item = self.materials_table.item(row_index, 0)
                if item is not None and str(item.data(Qt.UserRole) or item.text() or "").strip() == material:
                    self.materials_table.selectRow(row_index)
                    self._on_material_selected()
                    break
        if espessura:
            for row_index in range(self.esp_table.rowCount()):
                item = self.esp_table.item(row_index, 0)
                if item is not None and str(item.data(Qt.UserRole) or item.text() or "").strip() == espessura:
                    self.esp_table.selectRow(row_index)
                    self._on_esp_selected()
                    break
        if piece_ref:
            for row_index in range(self.pieces_table.rowCount()):
                item = self.pieces_table.item(row_index, 0)
                if item is not None and str(item.data(Qt.UserRole) or item.text() or "").strip() == piece_ref:
                    self.pieces_table.selectRow(row_index)
                    break
        self._sync_action_buttons()

    def _open_selected_piece_drawing(self) -> None:
        current_order = self._selected_order_row()
        current_piece = self._selected_piece_row()
        if not current_order or not current_piece:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda e uma peca.")
            return
        numero = str(current_order.get("numero", "") or "").strip()
        piece_id = str(current_piece.get("id", "") or "").strip()
        try:
            self.backend.operator_open_drawing(numero, piece_id)
        except Exception as exc:
            QMessageBox.critical(self, "Ver desenho", str(exc))


class LegacyOrdersPage(OrdersPage):
    page_subtitle = "Carteira de ordens de fabrico, materiais, progresso e necessidades de montagem."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        self.new_btn.setText("Nova ordem de fabrico")
        self.new_btn.setProperty("variant", "success")
        self.edit_header_btn.setText("Editar ordem")
        self.remove_btn.setText("Eliminar ordem")
        self.reserve_btn.setText("Cativar material")
        self.release_btn.setText("Libertar material")
        for button, width in (
            (self.add_piece_btn, 80),
            (self.import_model_btn, 116),
            (self.print_of_btn, 116),
            (self.edit_piece_btn, 82),
            (self.remove_piece_btn, 92),
            (self.open_piece_btn, 94),
        ):
            button.setProperty("compact", "true")
            button.setMinimumWidth(width)
            button.setMaximumWidth(width + 10)
            button.setStyleSheet("font-size: 9px; font-weight: 700;")
        self._rebuild_order_overview()
        self.table.verticalHeader().setDefaultSectionSize(38)
        self.materials_table.verticalHeader().setDefaultSectionSize(28)
        self.esp_table.verticalHeader().setDefaultSectionSize(28)
        self.pieces_table.verticalHeader().setDefaultSectionSize(28)
        self.montagem_table.verticalHeader().setDefaultSectionSize(26)
        _set_table_columns(
            self.pieces_table,
            [
                (0, "fixed", 156),
                (1, "stretch", 0),
                (2, "fixed", 154),
                (3, "fixed", 92),
                (4, "fixed", 54),
                (5, "stretch", 0),
                (6, "fixed", 68),
                (7, "fixed", 68),
                (8, "fixed", 96),
            ],
        )
        _set_table_columns(
            self.montagem_table,
            [
                (0, "fixed", 104),
                (1, "fixed", 110),
                (2, "stretch", 0),
                (3, "fixed", 72),
                (4, "fixed", 72),
                (5, "fixed", 78),
                (6, "fixed", 68),
                (7, "fixed", 92),
            ],
        )
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 154),
                (1, "fixed", 164),
                (2, "stretch", 0),
                (3, "fixed", 190),
                (4, "fixed", 102),
                (5, "fixed", 118),
                (6, "fixed", 96),
                (7, "fixed", 68),
                (8, "fixed", 122),
            ],
        )
        self.table.setStyleSheet(
            f"{self.table.styleSheet()}\n"
            "QHeaderView::section {"
            " background: #454945; color: #ffffff; border: 0;"
            " border-right: 1px solid #626862; padding: 7px 9px;"
            " font-size: 10px; font-weight: 800;"
            "}"
            "QTableCornerButton::section { background: #454945; border: 0; }"
        )
        root = self.layout()
        sections = _take_layout_items(root)
        filters_item = sections[0] if len(sections) > 0 else None
        info_item = sections[1] if len(sections) > 1 else None
        list_item = sections[2] if len(sections) > 2 else None
        mid_item = sections[3] if len(sections) > 3 else None
        pieces_item = sections[4] if len(sections) > 4 else None
        montagem_item = sections[5] if len(sections) > 5 else None
        if pieces_item is not None and pieces_item.widget() is not None:
            pieces_item.widget().setMinimumHeight(300)
            self._rebuild_order_section_toolbar(
                pieces_item.widget(),
                "Peças da ordem",
                (
                    self.add_piece_btn,
                    self.import_model_btn,
                    self.print_of_btn,
                    self.edit_piece_btn,
                    self.remove_piece_btn,
                    self.open_piece_btn,
                ),
                "pieces",
            )
        if montagem_item is not None and montagem_item.widget() is not None:
            montagem_item.widget().setMinimumHeight(210)
            self._rebuild_order_section_toolbar(
                montagem_item.widget(),
                "Montagem e componentes",
                (self.open_operator_montagem_btn, self.montagem_note_btn),
                "assembly",
            )
        if filters_item and filters_item.widget() is not None:
            filters_item.widget().hide()

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)

        filters_card = CardFrame()
        filters_card.set_tone("info")
        filters_layout = QVBoxLayout(filters_card)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(8)
        portfolio_header = QHBoxLayout()
        portfolio_identity = QVBoxLayout()
        portfolio_identity.setSpacing(1)
        portfolio_title = QLabel("Carteira de ordens de fabrico")
        portfolio_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #10253d;")
        portfolio_subtitle = QLabel("Acompanhe prazos, material cativado e progresso antes de abrir o detalhe produtivo.")
        portfolio_subtitle.setProperty("role", "muted")
        portfolio_subtitle.setWordWrap(True)
        portfolio_subtitle.setMinimumWidth(0)
        portfolio_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        portfolio_identity.addWidget(portfolio_title)
        portfolio_identity.addWidget(portfolio_subtitle)
        portfolio_header.addLayout(portfolio_identity, 1)

        def portfolio_metric(label_text: str) -> tuple[QFrame, QLabel]:
            frame = QFrame()
            frame.setStyleSheet("background: #ffffff; border: 1px solid #c8d7e6;")
            frame.setFixedSize(132, 46)
            metric_layout = QVBoxLayout(frame)
            metric_layout.setContentsMargins(9, 4, 9, 4)
            metric_layout.setSpacing(0)
            label = QLabel(label_text.upper())
            label.setStyleSheet("font-size: 8px; font-weight: 800; color: #60758d; border: none;")
            value = QLabel("0")
            value.setStyleSheet("font-size: 12px; font-weight: 900; color: #10253d; border: none;")
            metric_layout.addWidget(label)
            metric_layout.addWidget(value)
            return frame, value

        total_metric, self.orders_total_metric = portfolio_metric("Ordens")
        active_metric, self.orders_active_metric = portfolio_metric("Em produção")
        progress_metric, self.orders_average_metric = portfolio_metric("Progresso médio")
        portfolio_header.addWidget(total_metric)
        portfolio_header.addWidget(active_metric)
        portfolio_header.addWidget(progress_metric)
        filters_layout.addLayout(portfolio_header)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(7)
        self.filter_edit.lineEdit().setPlaceholderText("Pesquisar encomenda, OF, cliente ou referência...")
        self.filter_edit.setMinimumWidth(260)
        self.filter_edit.setMaximumWidth(520)
        self.state_combo.setMinimumWidth(130)
        self.state_combo.setMaximumWidth(160)
        self.year_combo.setMinimumWidth(92)
        self.year_combo.setMaximumWidth(112)
        self.client_combo.setMinimumWidth(220)
        self.client_combo.setMaximumWidth(360)
        filter_row.addWidget(self.filter_edit, 2)
        filter_row.addWidget(self.state_combo)
        filter_row.addWidget(self.year_combo)
        filter_row.addWidget(self.client_combo, 1)
        filter_row.addStretch(1)
        filters_layout.addLayout(filter_row)

        command_row = QHBoxLayout()
        command_row.setSpacing(7)
        self.open_order_btn = QPushButton("Abrir ordem")
        self.open_order_btn.clicked.connect(self._open_selected_order)
        refresh_order_btn = QPushButton("Atualizar")
        refresh_order_btn.setProperty("variant", "secondary")
        refresh_order_btn.clicked.connect(self.refresh)
        for button, width in (
            (self.new_btn, 146),
            (self.open_order_btn, 102),
            (self.edit_header_btn, 104),
            (refresh_order_btn, 88),
            (self.remove_btn, 108),
        ):
            button.setProperty("compact", "true")
            button.setMinimumWidth(width)
        command_row.addWidget(self.new_btn)
        command_row.addWidget(self.open_order_btn)
        command_row.addWidget(self.edit_header_btn)
        command_row.addWidget(refresh_order_btn)
        command_row.addStretch(1)
        command_row.addWidget(self.remove_btn)
        filters_layout.addLayout(command_row)
        list_layout.addWidget(filters_card)
        _adopt_layout_item(list_layout, list_item, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)
        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(12, 6, 12, 6)
        detail_actions_layout.setSpacing(8)
        detail_actions.setMaximumHeight(52)
        back_btn = QPushButton("Voltar às ordens")
        back_btn.setProperty("variant", "success")
        back_btn.clicked.connect(self._show_order_list)
        edit_btn = QPushButton("Editar dados da OF")
        edit_btn.clicked.connect(self._edit_order_header)
        remove_btn = QPushButton("Eliminar ordem")
        remove_btn.setProperty("variant", "danger")
        remove_btn.clicked.connect(self._remove_order)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.refresh)
        for button, width in ((back_btn, 128), (edit_btn, 146), (remove_btn, 126), (refresh_btn, 96)):
            button.setMinimumWidth(width)
        detail_actions_layout.addWidget(back_btn)
        context_layout = QVBoxLayout()
        context_layout.setContentsMargins(4, 0, 8, 0)
        context_layout.setSpacing(0)
        self.order_context_title = QLabel("Ordem de fabrico")
        self.order_context_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #10253d;")
        self.order_context_meta = QLabel("Selecione uma ordem para consultar o detalhe.")
        self.order_context_meta.setProperty("role", "muted")
        context_layout.addWidget(self.order_context_title)
        context_layout.addWidget(self.order_context_meta)
        detail_actions_layout.addLayout(context_layout, 1)
        detail_actions_layout.addWidget(edit_btn)
        detail_actions_layout.addWidget(refresh_btn)
        self.import_model_btn.setProperty("variant", "secondary")
        self.import_model_btn.setMinimumWidth(126)
        self.import_model_btn.setMaximumWidth(144)
        self.print_of_btn.setProperty("variant", "success")
        self.print_of_btn.setMinimumWidth(128)
        self.print_of_btn.setMaximumWidth(146)
        detail_actions_layout.addWidget(self.import_model_btn)
        detail_actions_layout.addWidget(self.print_of_btn)
        detail_actions_layout.addStretch(1)
        detail_actions_layout.addWidget(remove_btn)
        detail_layout.addWidget(detail_actions)

        self.order_detail_stack = QStackedWidget()

        overview_page = QWidget()
        overview_layout = QVBoxLayout(overview_page)
        overview_layout.setContentsMargins(0, 0, 0, 0)
        overview_layout.setSpacing(10)
        _adopt_layout_item(overview_layout, info_item)

        selection_strip = CardFrame()
        selection_strip.set_tone("info")
        selection_layout = QHBoxLayout(selection_strip)
        selection_layout.setContentsMargins(14, 7, 14, 7)
        selection_layout.setSpacing(10)
        selection_text = QVBoxLayout()
        selection_text.setSpacing(0)
        selection_title = QLabel("Selecionar material e espessura")
        selection_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #10253d;")
        selection_hint = QLabel("Escolha o grupo de fabrico e avance para trabalhar apenas nas peças correspondentes.")
        selection_hint.setProperty("role", "muted")
        selection_text.addWidget(selection_title)
        selection_text.addWidget(selection_hint)
        self.open_order_group_btn = QPushButton("Abrir peças selecionadas")
        self.open_order_group_btn.clicked.connect(self._show_order_production)
        self.open_order_group_btn.setMinimumWidth(180)
        self.open_order_assembly_btn = QPushButton("Montagem e componentes")
        self.open_order_assembly_btn.setProperty("variant", "secondary")
        self.open_order_assembly_btn.clicked.connect(self._show_order_assembly)
        self.open_order_assembly_btn.setMinimumWidth(190)
        selection_layout.addLayout(selection_text, 1)
        selection_layout.addWidget(self.open_order_assembly_btn)
        selection_layout.addWidget(self.open_order_group_btn)
        selection_strip.setMaximumHeight(62)
        overview_layout.addWidget(selection_strip)

        hierarchy_split = mid_item.widget() if mid_item is not None else None
        if isinstance(hierarchy_split, QSplitter):
            hierarchy_split.setOrientation(Qt.Horizontal)
            hierarchy_split.setHandleWidth(6)
            hierarchy_split.setSizes([760, 1140])
            for index in range(hierarchy_split.count()):
                hierarchy_split.widget(index).setMinimumWidth(320)
        _adopt_layout_item(overview_layout, mid_item, 1)

        production_page = QWidget()
        production_layout = QVBoxLayout(production_page)
        production_layout.setContentsMargins(0, 0, 0, 0)
        production_layout.setSpacing(10)
        production_nav = CardFrame()
        production_nav.set_tone("info")
        production_nav_layout = QHBoxLayout(production_nav)
        production_nav_layout.setContentsMargins(12, 7, 12, 7)
        production_nav_layout.setSpacing(10)
        back_to_selection_btn = QPushButton("Voltar à seleção")
        back_to_selection_btn.setProperty("variant", "secondary")
        back_to_selection_btn.clicked.connect(self._show_order_overview)
        back_to_selection_btn.setMinimumWidth(126)
        production_context = QVBoxLayout()
        production_context.setSpacing(0)
        self.production_context_title = QLabel("Peças da ordem")
        self.production_context_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #10253d;")
        self.production_context_meta = QLabel("Material e espessura selecionados")
        self.production_context_meta.setProperty("role", "muted")
        production_context.addWidget(self.production_context_title)
        production_context.addWidget(self.production_context_meta)
        production_nav_layout.addWidget(back_to_selection_btn)
        production_nav_layout.addLayout(production_context, 1)
        production_nav.setMaximumHeight(58)
        production_layout.addWidget(production_nav)

        pieces_host = QWidget()
        pieces_layout = QVBoxLayout(pieces_host)
        pieces_layout.setContentsMargins(0, 0, 0, 0)
        pieces_layout.setSpacing(0)
        _adopt_layout_item(pieces_layout, pieces_item, 1)
        production_layout.addWidget(pieces_host, 1)

        assembly_page = QWidget()
        assembly_page_layout = QVBoxLayout(assembly_page)
        assembly_page_layout.setContentsMargins(0, 0, 0, 0)
        assembly_page_layout.setSpacing(10)
        assembly_nav = CardFrame()
        assembly_nav.set_tone("warning")
        assembly_nav_layout = QHBoxLayout(assembly_nav)
        assembly_nav_layout.setContentsMargins(12, 7, 12, 7)
        assembly_nav_layout.setSpacing(10)
        back_from_assembly_btn = QPushButton("Voltar à seleção")
        back_from_assembly_btn.setProperty("variant", "secondary")
        back_from_assembly_btn.clicked.connect(self._show_order_overview)
        back_from_assembly_btn.setMinimumWidth(126)
        assembly_context = QVBoxLayout()
        assembly_context.setSpacing(0)
        assembly_title = QLabel("Montagem e componentes da ordem")
        assembly_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #10253d;")
        assembly_hint = QLabel("Necessidades globais da ordem, independentes do material e da espessura selecionados.")
        assembly_hint.setProperty("role", "muted")
        assembly_context.addWidget(assembly_title)
        assembly_context.addWidget(assembly_hint)
        assembly_nav_layout.addWidget(back_from_assembly_btn)
        assembly_nav_layout.addLayout(assembly_context, 1)
        assembly_nav.setMaximumHeight(58)
        assembly_page_layout.addWidget(assembly_nav)
        if montagem_item is not None:
            montagem_host = QWidget()
            montagem_layout = QVBoxLayout(montagem_host)
            montagem_layout.setContentsMargins(0, 0, 0, 0)
            montagem_layout.setSpacing(0)
            _adopt_layout_item(montagem_layout, montagem_item, 1)
            assembly_page_layout.addWidget(montagem_host, 1)

        self.order_detail_stack.addWidget(overview_page)
        self.order_detail_stack.addWidget(production_page)
        self.order_detail_stack.addWidget(assembly_page)
        detail_layout.addWidget(self.order_detail_stack, 1)

        compact_table_style = (
            "QTableWidget { font-size: 10px; }"
            "QTableWidget::item { padding: 2px 5px; }"
            "QHeaderView::section { font-size: 9px; font-weight: 800; padding: 4px 6px; }"
        )
        self.pieces_table.setStyleSheet(compact_table_style)
        self.montagem_table.setStyleSheet(compact_table_style)
        self.pieces_table.verticalHeader().setDefaultSectionSize(26)
        self.montagem_table.verticalHeader().setDefaultSectionSize(25)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.table.itemSelectionChanged.connect(self._sync_list_open_button)
        self.table.itemDoubleClicked.connect(lambda item: self._open_selected_order(item))
        self.table.itemSelectionChanged.connect(self._sync_order_overview)
        self.materials_table.itemSelectionChanged.connect(self._sync_order_section_context)
        self.esp_table.itemSelectionChanged.connect(self._sync_order_section_context)
        self.materials_table.cellDoubleClicked.connect(lambda _row, _column: self._show_order_production())
        self.esp_table.cellDoubleClicked.connect(lambda _row, _column: self._show_order_production())
        self._show_order_overview()
        self._show_order_list()
        self._sync_list_open_button()

    def _rebuild_order_section_toolbar(
        self,
        card: QWidget,
        title_text: str,
        buttons: tuple[QPushButton, ...],
        context_key: str,
    ) -> None:
        card_layout = card.layout()
        if not isinstance(card_layout, QVBoxLayout) or card_layout.count() == 0:
            return
        first_item = card_layout.takeAt(0)
        old_header = first_item.layout()
        if old_header is not None:
            while old_header.count():
                item = old_header.takeAt(0)
                widget = item.widget()
                if widget is not None and widget not in buttons:
                    widget.deleteLater()

        toolbar = QVBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(5)
        heading_row = QHBoxLayout()
        heading_row.setSpacing(8)
        title = QLabel(title_text)
        title.setStyleSheet("font-size: 14px; font-weight: 800; color: #10253d;")
        context = QLabel("-")
        context.setProperty("role", "muted")
        context.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        heading_row.addWidget(title)
        heading_row.addStretch(1)
        heading_row.addWidget(context)
        toolbar.addLayout(heading_row)
        actions = QHBoxLayout()
        actions.setSpacing(6)
        actions.addStretch(1)
        for button in buttons:
            button.setProperty("compact", "true")
            actions.addWidget(button)
        toolbar.addLayout(actions)
        card_layout.insertLayout(0, toolbar)
        if context_key == "pieces":
            self.pieces_context_label = context
        else:
            self.assembly_context_label = context

    def _rebuild_order_overview(self) -> None:
        layout = self.info_card.layout()
        if layout is None:
            return
        preserved = {
            self.info_numero,
            self.info_of,
            self.info_cliente,
            self.info_entrega,
            self.info_estado,
            self.info_nota,
            self.info_transporte,
            self.info_descarga,
            self.info_viagem,
            self.info_transportadora,
            self.info_carga,
            self.info_custos,
            self.info_reservas,
            self.info_cativar,
            self.info_chapa,
            self.reserve_btn,
            self.release_btn,
        }
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None and widget not in preserved:
                widget.deleteLater()

        layout.setContentsMargins(14, 9, 14, 9)
        layout.setHorizontalSpacing(0)
        layout.setVerticalSpacing(0)
        host = QWidget()
        host_layout = QVBoxLayout(host)
        host_layout.setContentsMargins(0, 0, 0, 0)
        host_layout.setSpacing(7)

        identity_row = QHBoxLayout()
        identity_row.setSpacing(16)
        identity = QVBoxLayout()
        identity.setSpacing(2)
        identity_heading = QHBoxLayout()
        identity_heading.setContentsMargins(0, 0, 0, 0)
        identity_heading.setSpacing(8)
        eyebrow = QLabel("RESUMO DA ORDEM")
        eyebrow.setStyleSheet("font-size: 10px; font-weight: 700; color: #52708d; letter-spacing: 0.3px;")
        identity_heading.addWidget(eyebrow)
        identity_heading.addStretch(1)
        identity.addLayout(identity_heading)
        self.info_of.setStyleSheet("font-family: 'Segoe UI Semibold'; font-size: 18px; font-weight: 600; color: #10253d;")
        order_number_row = QHBoxLayout()
        order_number_row.setContentsMargins(0, 0, 0, 0)
        order_number_row.setSpacing(7)
        order_number_label = QLabel("Encomenda")
        order_number_label.setProperty("role", "muted")
        order_number_label.setStyleSheet("font-size: 10px; color: #60758d;")
        self.info_numero.setStyleSheet("font-size: 11px; font-weight: 700; color: #334e68;")
        order_number_row.addWidget(order_number_label)
        order_number_row.addWidget(self.info_numero)
        order_number_row.addStretch(1)
        identity.addWidget(self.info_of)
        identity.addLayout(order_number_row)
        identity_row.addLayout(identity, 2)

        client_block = QVBoxLayout()
        client_block.setSpacing(2)
        client_label = QLabel("CLIENTE E REFERÊNCIA")
        client_label.setStyleSheet("font-size: 10px; font-weight: 700; color: #52708d; letter-spacing: 0.2px;")
        self.info_cliente.setStyleSheet("font-family: 'Segoe UI Semibold'; font-size: 12px; font-weight: 600; color: #10253d;")
        self.info_nota.setStyleSheet("font-size: 10px; color: #516981;")
        client_block.addWidget(client_label)
        client_block.addWidget(self.info_cliente)
        client_block.addWidget(self.info_nota)
        identity_row.addLayout(client_block, 3)
        identity_row.addWidget(self.info_estado, 0, Qt.AlignVCenter)
        host_layout.addLayout(identity_row)

        metrics_frame = QFrame()
        metrics_frame.setObjectName("OrderOverviewMetrics")
        metrics_frame.setStyleSheet(
            "QFrame#OrderOverviewMetrics { background: #f8fafc; border: 1px solid #d5e0ea; border-radius: 6px; }"
            "QLabel#OrderMetricLabel { color: #60758d; font-size: 9px; font-weight: 700; }"
            "QLabel#OrderMetricValue { color: #10253d; font-family: 'Segoe UI Semibold'; font-size: 12px; font-weight: 600; }"
        )
        metrics_layout = QHBoxLayout(metrics_frame)
        metrics_layout.setContentsMargins(12, 5, 12, 5)
        metrics_layout.setSpacing(20)

        def add_metric(label_text: str, value_widget: QLabel) -> None:
            metric = QWidget()
            metric_layout = QVBoxLayout(metric)
            metric_layout.setContentsMargins(0, 0, 0, 0)
            metric_layout.setSpacing(0)
            label = QLabel(label_text.upper())
            label.setObjectName("OrderMetricLabel")
            value_widget.setObjectName("OrderMetricValue")
            metric_layout.addWidget(label)
            metric_layout.addWidget(value_widget)
            metrics_layout.addWidget(metric, 1)

        self.order_piece_metric = QLabel("0")
        self.order_material_metric = QLabel("0")
        self.order_progress_metric = QLabel("0%")
        add_metric("Entrega", self.info_entrega)
        add_metric("Peças", self.order_piece_metric)
        add_metric("Materiais", self.order_material_metric)
        add_metric("Progresso", self.order_progress_metric)
        self.order_progress_bar = QProgressBar()
        self.order_progress_bar.setRange(0, 100)
        self.order_progress_bar.setValue(0)
        self.order_progress_bar.setTextVisible(False)
        self.order_progress_bar.setFixedSize(150, 10)
        self.order_progress_bar.setStyleSheet(
            "QProgressBar { background: #e6edf5; border: 0; border-radius: 5px; }"
            "QProgressBar::chunk { background: #86bc55; border-radius: 5px; }"
        )
        metrics_layout.addWidget(self.order_progress_bar, 0, Qt.AlignVCenter)
        host_layout.addWidget(metrics_frame)

        logistics_frame = QFrame()
        logistics_frame.setObjectName("OrderLogistics")
        logistics_frame.setStyleSheet(
            "QFrame#OrderLogistics { background: #ffffff; border: 1px solid #dce4eb; border-radius: 6px; }"
            "QFrame#OrderLogistics QWidget { background: transparent; border: 0; }"
        )
        logistics_layout = QHBoxLayout(logistics_frame)
        logistics_layout.setContentsMargins(12, 6, 12, 6)
        logistics_layout.setSpacing(24)

        def add_logistics_group(
            title_text: str,
            primary_value: QLabel,
            secondary_value: QLabel,
        ) -> None:
            group = QWidget()
            group_layout = QVBoxLayout(group)
            group_layout.setContentsMargins(0, 0, 0, 0)
            group_layout.setSpacing(1)
            label = QLabel(title_text.upper())
            label.setStyleSheet("font-size: 9px; font-weight: 700; color: #60758d; letter-spacing: 0.2px;")
            primary_value.setStyleSheet("font-size: 10px; font-weight: 600; color: #29445f;")
            secondary_value.setStyleSheet("font-size: 10px; color: #52677d;")
            for value in (primary_value, secondary_value):
                value.setMinimumWidth(0)
                value.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
            group_layout.addWidget(label)
            group_layout.addWidget(primary_value)
            group_layout.addWidget(secondary_value)
            logistics_layout.addWidget(group, 1)

        add_logistics_group("Transporte", self.info_transporte, self.info_transportadora)
        add_logistics_group("Destino e carga", self.info_descarga, self.info_carga)
        add_logistics_group("Viagem e custos", self.info_viagem, self.info_custos)
        host_layout.addWidget(logistics_frame)

        reservation_frame = QFrame()
        reservation_frame.setObjectName("OrderReservationStrip")
        reservation_frame.setStyleSheet(
            "QFrame#OrderReservationStrip { background: #f5f9f1; border: 1px solid #d4e3c7; border-radius: 6px; }"
        )
        reservation_layout = QHBoxLayout(reservation_frame)
        reservation_layout.setContentsMargins(12, 5, 8, 5)
        reservation_layout.setSpacing(10)
        reservation_text = QVBoxLayout()
        reservation_text.setSpacing(1)
        reservation_title = QLabel("MATERIAL CATIVADO")
        reservation_title.setStyleSheet("font-size: 9px; font-weight: 700; color: #527444; letter-spacing: 0.2px;")
        reservation_summary = QHBoxLayout()
        reservation_summary.setSpacing(8)
        self.info_chapa.setStyleSheet("font-size: 10px; font-weight: 700; color: #294126;")
        self.info_reservas.setStyleSheet("font-size: 10px; color: #52677d;")
        reservation_summary.addWidget(self.info_chapa)
        reservation_summary.addWidget(self.info_reservas, 1)
        reservation_text.addWidget(reservation_title)
        reservation_text.addLayout(reservation_summary)
        reservation_layout.addLayout(reservation_text, 1)
        reservation_layout.addWidget(self.info_cativar)
        reservation_layout.addWidget(self.reserve_btn)
        reservation_layout.addWidget(self.release_btn)
        host_layout.addWidget(reservation_frame)

        layout.addWidget(host, 0, 0, 1, 6)
        self.info_card.setMinimumHeight(218)
        self.info_card.setMaximumHeight(236)
        self._sync_order_overview()

    def _sync_order_overview(self) -> None:
        if not hasattr(self, "order_progress_metric"):
            return
        detail = dict(self.current_detail or {})
        pieces = list(detail.get("pieces", []) or [])
        materials = list(detail.get("materials_tree", []) or [])

        def number(value) -> float:
            try:
                return float(value or 0)
            except Exception:
                return 0.0

        planned = sum(max(0.0, number(row.get("qtd_plan", 0))) for row in pieces)
        produced = sum(max(0.0, number(row.get("qtd_prod", 0))) for row in pieces)
        progress = min(100.0, (produced / planned) * 100.0) if planned > 0 else 0.0
        self.order_piece_metric.setText(str(len(pieces)))
        self.order_material_metric.setText(str(len(materials)))
        self.order_progress_metric.setText(f"{progress:.0f}%")
        self.order_progress_bar.setValue(int(round(progress)))
        if hasattr(self, "order_context_title"):
            numero = str(detail.get("numero", "") or "").strip()
            of_code = str(detail.get("of_codigo", "") or "").strip()
            client = str(self.info_cliente.text() or "").strip()
            self.order_context_title.setText(of_code or "Ordem de fabrico")
            context_bits = [value for value in (numero, client) if value and value != "-"]
            self.order_context_meta.setText(" | ".join(context_bits) or "Selecione uma ordem para consultar o detalhe.")
        self._sync_order_section_context()

    def _sync_order_section_context(self) -> None:
        material = str(self._selected_material_row().get("material", "") or "").strip()
        thickness = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if hasattr(self, "pieces_context_label"):
            bits = [f"{len(self.detail_pieces)} peça(s)"]
            if material:
                bits.append(material)
            if thickness:
                bits.append(f"{thickness} mm")
            self.pieces_context_label.setText(" | ".join(bits))
        if hasattr(self, "assembly_context_label"):
            shortages = len(list(self.current_detail.get("montagem_shortages", []) or []))
            if not self.detail_montagem:
                self.assembly_context_label.setText("Não aplicável")
            elif shortages:
                self.assembly_context_label.setText(f"{len(self.detail_montagem)} componentes | {shortages} em falta")
            else:
                self.assembly_context_label.setText(f"{len(self.detail_montagem)} componentes | Stock disponível")
        if hasattr(self, "open_order_group_btn"):
            self.open_order_group_btn.setEnabled(bool(material and thickness))
        if hasattr(self, "production_context_title"):
            group_label = " · ".join(value for value in (material, f"{thickness} mm" if thickness else "") if value)
            self.production_context_title.setText(group_label or "Peças da ordem")
            self.production_context_meta.setText(
                f"{len(self.detail_pieces)} peça(s) neste grupo · consulte e edite as operações abaixo."
                if group_label
                else "Volte à seleção e escolha um material e uma espessura."
            )

    def _sync_order_portfolio_metrics(self) -> None:
        if not hasattr(self, "orders_total_metric"):
            return
        rows = list(self.rows or [])
        active = 0
        progress_values = []
        for row in rows:
            state = unicodedata.normalize("NFKD", str(row.get("estado", "") or "")).encode("ascii", "ignore").decode().casefold()
            if not any(token in state for token in ("conclu", "cancel", "entreg")):
                active += 1
            try:
                progress_values.append(max(0.0, min(100.0, float(row.get("progress", 0) or 0))))
            except Exception:
                progress_values.append(0.0)
        average = sum(progress_values) / len(progress_values) if progress_values else 0.0
        self.orders_total_metric.setText(str(len(rows)))
        self.orders_active_metric.setText(str(active))
        self.orders_average_metric.setText(f"{average:.0f}%")

    def refresh(self) -> None:
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(self.current_detail.get("numero"))
        keep_detail_section = self.order_detail_stack.currentIndex() if keep_detail and hasattr(self, "order_detail_stack") else 0
        OrdersPage.refresh(self)
        self._sync_order_overview()
        self._sync_order_portfolio_metrics()
        if self.table.rowCount() == 0:
            self._show_order_list()
        elif keep_detail:
            self._show_order_detail()
            if keep_detail_section == 1:
                self._show_order_production()
            elif keep_detail_section == 2:
                self._show_order_assembly()
        else:
            self._show_order_list()
        self._sync_list_open_button()

    def _show_order_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._show_order_overview()
        self._sync_list_open_button()

    def _show_order_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def _show_order_overview(self) -> None:
        if hasattr(self, "order_detail_stack"):
            self.order_detail_stack.setCurrentIndex(0)

    def _show_order_production(self) -> None:
        material = str(self._selected_material_row().get("material", "") or "").strip()
        thickness = str(self._selected_esp_row().get("espessura", "") or "").strip()
        if not material or not thickness:
            if self.view_stack.currentWidget() is self.detail_page:
                QMessageBox.information(
                    self,
                    "Peças da ordem",
                    "Seleciona primeiro um material e uma espessura.",
                )
            return
        self._sync_order_section_context()
        self.order_detail_stack.setCurrentIndex(1)

    def _show_order_assembly(self) -> None:
        if hasattr(self, "order_detail_stack"):
            self.order_detail_stack.setCurrentIndex(2)

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _sync_list_open_button(self) -> None:
        self.open_order_btn.setEnabled(bool(self._selected_order_row()))

    def _open_selected_order(self, item: QTableWidgetItem | None = None) -> None:
        row = {}
        if item is not None and hasattr(item, "row"):
            row_item = self.table.item(item.row(), 0) or item
            numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
            if numero:
                row = next((entry for entry in self.rows if str(entry.get("numero", "") or "").strip() == numero), {})
        if not row:
            row = self._selected_order_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        OrdersPage._select_order(self, numero)
        self._show_order_overview()
        self._show_order_detail()

    def _select_order(self, numero: str, piece_ref: str = "", material: str = "", espessura: str = "") -> None:
        OrdersPage._select_order(self, numero, piece_ref=piece_ref, material=material, espessura=espessura)
        self._show_order_detail()
        if piece_ref or (material and espessura):
            self._show_order_production()
        else:
            self._show_order_overview()

    def _remove_order(self) -> None:
        OrdersPage._remove_order(self)
        self._show_order_list()
        self._sync_list_open_button()
