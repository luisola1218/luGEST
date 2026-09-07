from __future__ import annotations
import csv
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)
from pathlib import Path
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    configure_table as _configure_table,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    selected_row_index as _selected_row_index,
    set_table_columns as _set_table_columns,
)
from .runtime_support import _apply_progress_style, _fmt_eur, _format_client_label
from ..widgets import CardFrame, StatCard


class OppPage(QWidget):
    page_title = "OPP"
    page_subtitle = "Ordens de fabrico por peça, com rastreio operacional, histórico e expedição."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.all_rows: list[dict] = []
        self.rows: list[dict] = []
        self.current_detail: dict[str, Any] = {}

        main_root = QVBoxLayout(self)
        main_root.setContentsMargins(0, 0, 0, 10)
        main_root.setSpacing(0)
        self.main_tabs = QTabWidget()
        self.portfolio_page = QWidget()
        self.opp_page = QWidget()
        self.main_tabs.addTab(self.portfolio_page, "Clientes e encomendas")
        self.main_tabs.addTab(self.opp_page, "Ordens OPP")
        self.main_tabs.currentChanged.connect(self._opp_main_tab_changed)
        main_root.addWidget(self.main_tabs, 1)

        root = QVBoxLayout(self.opp_page)
        root.setContentsMargins(0, 0, 0, 10)
        root.setSpacing(14)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(16, 14, 16, 14)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(8)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar OPP, OF, encomenda, cliente, referencia, material ou operacao")
        self.filter_edit.textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Todas", "Preparacao", "Em producao", "Concluida", "Expedidas", "Avaria"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.client_combo = QComboBox()
        self.client_combo.currentTextChanged.connect(self.refresh)
        self.operation_combo = QComboBox()
        self.operation_combo.currentTextChanged.connect(self.refresh)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.export_btn = QPushButton("Exportar CSV")
        self.export_btn.setProperty("variant", "secondary")
        self.export_btn.clicked.connect(self._export_csv)
        self.pdf_btn = QPushButton("Etiqueta OPP")
        self.pdf_btn.clicked.connect(self._open_label_pdf)
        self.order_btn = QPushButton("Abrir encomenda")
        self.order_btn.setProperty("variant", "success")
        self.order_btn.clicked.connect(self._open_order)
        self.drawing_btn = QPushButton("Ver desenho")
        self.drawing_btn.setProperty("variant", "secondary")
        self.drawing_btn.clicked.connect(self._open_drawing)
        filters_layout.addWidget(QLabel("Pesquisa"), 0, 0)
        filters_layout.addWidget(self.filter_edit, 0, 1, 1, 2)
        filters_layout.addWidget(QLabel("Estado"), 0, 3)
        filters_layout.addWidget(self.state_combo, 0, 4)
        filters_layout.addWidget(QLabel("Ano"), 0, 5)
        filters_layout.addWidget(self.year_combo, 0, 6)
        filters_layout.addWidget(QLabel("Cliente"), 1, 0)
        filters_layout.addWidget(self.client_combo, 1, 1, 1, 2)
        filters_layout.addWidget(QLabel("Operacao"), 1, 3)
        filters_layout.addWidget(self.operation_combo, 1, 4)
        action_host = QWidget()
        action_layout = QHBoxLayout(action_host)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(6)
        for button, width in ((self.refresh_btn, 110), (self.export_btn, 120), (self.pdf_btn, 126), (self.drawing_btn, 120), (self.order_btn, 144)):
            button.setMinimumWidth(width)
            action_layout.addWidget(button)
        action_layout.addStretch(1)
        filters_layout.addWidget(action_host, 1, 5, 1, 2)
        root.addWidget(filters)

        stats_host = QWidget()
        stats_layout = QGridLayout(stats_host)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setHorizontalSpacing(12)
        self.stats_cards = [StatCard("OPP ativas"), StatCard("Em curso"), StatCard("Planeado"), StatCard("Expedido")]
        for index, tone in enumerate(("info", "warning", "success", "default")):
            self.stats_cards[index].set_tone(tone)
            stats_layout.addWidget(self.stats_cards[index], 0, index)
        root.addWidget(stats_host)

        list_card = CardFrame()
        list_card.set_tone("default")
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(16, 14, 16, 14)
        list_title = QLabel("Ordens de producao por peca")
        list_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 16)
        self.table.setHorizontalHeaderLabels(
            ["OPP", "OF", "Encomenda", "Cliente", "Ref. Int.", "Ref. Ext.", "Material", "Esp.", "Estado", "Operacao", "Operador", "Plan", "Prod", "Exp", "Tempo", "%"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QHeaderView::section { font-size: 12px; padding: 5px 6px; font-weight: 700; }"
        )
        _configure_table(self.table, stretch=(3, 5, 9, 10), contents=(0, 1, 2, 4, 6, 7, 8, 11, 12, 13, 14, 15))
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 138),
                (1, "fixed", 120),
                (2, "fixed", 162),
                (3, "stretch", 0),
                (4, "fixed", 138),
                (5, "stretch", 0),
                (6, "fixed", 94),
                (7, "fixed", 46),
                (8, "fixed", 116),
                (9, "fixed", 148),
                (10, "fixed", 110),
                (11, "fixed", 54),
                (12, "fixed", 54),
                (13, "fixed", 54),
                (14, "fixed", 78),
                (15, "fixed", 54),
            ],
        )
        self.table.itemSelectionChanged.connect(self._on_selected)
        self.table.itemDoubleClicked.connect(lambda *_args: self._open_order())
        list_layout.addWidget(list_title)
        list_layout.addWidget(self.table)
        root.addWidget(list_card, 2)

        self.detail_card = CardFrame()
        self.detail_card.set_tone("default")
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(16, 14, 16, 14)
        detail_layout.setSpacing(8)
        header_row = QHBoxLayout()
        self.title_label = QLabel("Sem OPP selecionada")
        self.title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.state_chip = QLabel("-")
        _apply_state_chip(self.state_chip, "-")
        header_row.addWidget(self.title_label, 1)
        header_row.addWidget(self.state_chip)
        detail_layout.addLayout(header_row)
        self.meta_label = QLabel("Seleciona uma OPP para ver detalhe de fabrico, tempos, eventos e expedicao.")
        self.meta_label.setWordWrap(True)
        self.meta_label.setProperty("role", "muted")
        detail_layout.addWidget(self.meta_label)
        self.flow_label = QLabel("-")
        self.flow_label.setWordWrap(True)
        detail_layout.addWidget(self.flow_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setFormat("%p%")
        _apply_progress_style(self.progress_bar, compact=True)
        detail_layout.addWidget(self.progress_bar)
        root.addWidget(self.detail_card)

        tabs_card = CardFrame()
        tabs_card.set_tone("info")
        tabs_layout = QVBoxLayout(tabs_card)
        tabs_layout.setContentsMargins(16, 14, 16, 14)
        tabs_layout.setSpacing(8)
        self.tabs = QTabWidget()
        self.ops_table = QTableWidget(0, 9)
        self.ops_table.setHorizontalHeaderLabels(["Operacao", "Estado", "Operador", "Inicio", "Fim", "OK", "NOK", "Qual.", "%"])
        self.ops_table.verticalHeader().setVisible(False)
        self.ops_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.ops_table, stretch=(0, 2), contents=(1, 3, 4, 5, 6, 7, 8))
        self.events_table = QTableWidget(0, 7)
        self.events_table.setHorizontalHeaderLabels(["Data", "Evento", "Operacao", "Operador", "OK", "NOK", "Info"])
        self.events_table.verticalHeader().setVisible(False)
        self.events_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.events_table, stretch=(6,), contents=(0, 1, 2, 3, 4, 5))
        self.exp_table = QTableWidget(0, 6)
        self.exp_table.setHorizontalHeaderLabels(["Guia", "Data", "Estado", "Destinatario", "Qtd", "Obs"])
        self.exp_table.verticalHeader().setVisible(False)
        self.exp_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.exp_table, stretch=(3, 5), contents=(0, 1, 2, 4))
        for title_text, table in (("Operacoes", self.ops_table), ("Histórico", self.events_table), ("Expedição", self.exp_table)):
            host = QWidget()
            host_layout = QVBoxLayout(host)
            host_layout.setContentsMargins(0, 0, 0, 0)
            host_layout.addWidget(table)
            self.tabs.addTab(host, title_text)
        tabs_layout.addWidget(self.tabs)
        root.addWidget(tabs_card, 2)
        self._build_client_portfolio()
        self._sync_buttons()

    def _build_client_portfolio(self) -> None:
        self.portfolio_clients: list[dict[str, Any]] = []
        self.portfolio_orders: list[dict[str, Any]] = []
        layout = QVBoxLayout(self.portfolio_page)
        layout.setContentsMargins(0, 0, 0, 10)
        layout.setSpacing(10)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(8)
        filters_layout.addWidget(QLabel("Cliente"))
        self.portfolio_client_combo = QComboBox()
        self.portfolio_client_combo.setMinimumWidth(290)
        self.portfolio_client_combo.currentIndexChanged.connect(self._refresh_portfolio)
        filters_layout.addWidget(self.portfolio_client_combo)
        filters_layout.addWidget(QLabel("Ano"))
        self.portfolio_year_combo = QComboBox()
        self.portfolio_year_combo.setMinimumWidth(108)
        self.portfolio_year_combo.currentIndexChanged.connect(self._refresh_portfolio)
        filters_layout.addWidget(self.portfolio_year_combo)
        filters_layout.addWidget(QLabel("Pesquisa"))
        self.portfolio_search_edit = QLineEdit()
        self.portfolio_search_edit.setPlaceholderText("Encomenda, OF, orçamento, referência ou cliente")
        self.portfolio_search_edit.textChanged.connect(self._refresh_portfolio)
        filters_layout.addWidget(self.portfolio_search_edit, 1)
        portfolio_refresh_btn = QPushButton("Atualizar")
        portfolio_refresh_btn.setProperty("variant", "secondary")
        portfolio_refresh_btn.clicked.connect(self._refresh_portfolio)
        filters_layout.addWidget(portfolio_refresh_btn)
        layout.addWidget(filters)

        stats_host = QWidget()
        stats_layout = QGridLayout(stats_host)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setHorizontalSpacing(10)
        self.portfolio_stats = [
            StatCard("Carteira"),
            StatCard("Valor adjudicado"),
            StatCard("Faturado"),
            StatCard("Por faturar"),
        ]
        for index, tone in enumerate(("info", "default", "success", "warning")):
            self.portfolio_stats[index].set_tone(tone)
            stats_layout.addWidget(self.portfolio_stats[index], 0, index)
        layout.addWidget(stats_host)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        clients_card = CardFrame()
        clients_layout = QVBoxLayout(clients_card)
        clients_layout.setContentsMargins(12, 10, 12, 10)
        clients_header = QHBoxLayout()
        clients_title = QLabel("Clientes adjudicados")
        clients_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.portfolio_clients_meta = QLabel("0 clientes")
        self.portfolio_clients_meta.setProperty("role", "muted")
        clients_header.addWidget(clients_title)
        clients_header.addStretch(1)
        clients_header.addWidget(self.portfolio_clients_meta)
        clients_layout.addLayout(clients_header)
        self.portfolio_clients_table = QTableWidget(0, 7)
        self.portfolio_clients_table.setHorizontalHeaderLabels(
            ["Cliente", "Enc.", "OPP", "Adjudicado", "Faturado", "Por faturar", "Saldo"]
        )
        self.portfolio_clients_table.verticalHeader().setVisible(False)
        self.portfolio_clients_table.verticalHeader().setDefaultSectionSize(30)
        self.portfolio_clients_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.portfolio_clients_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.portfolio_clients_table.setSelectionMode(QAbstractItemView.SingleSelection)
        _set_table_columns(
            self.portfolio_clients_table,
            [(0, "stretch", 0), (1, "fixed", 50), (2, "fixed", 50), (3, "fixed", 100), (4, "fixed", 100), (5, "fixed", 100), (6, "fixed", 96)],
        )
        self.portfolio_clients_table.itemSelectionChanged.connect(self._portfolio_client_selected)
        clients_layout.addWidget(self.portfolio_clients_table)
        splitter.addWidget(clients_card)

        orders_host = QWidget()
        orders_layout = QVBoxLayout(orders_host)
        orders_layout.setContentsMargins(0, 0, 0, 0)
        orders_layout.setSpacing(8)

        orders_card = CardFrame()
        orders_card_layout = QVBoxLayout(orders_card)
        orders_card_layout.setContentsMargins(12, 10, 12, 10)
        orders_header = QHBoxLayout()
        self.portfolio_orders_title = QLabel("Encomendas adjudicadas")
        self.portfolio_orders_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.portfolio_order_btn = QPushButton("Abrir encomenda")
        self.portfolio_order_btn.setProperty("variant", "success")
        self.portfolio_order_btn.clicked.connect(self._open_portfolio_order)
        self.portfolio_opp_btn = QPushButton("Ver OPP")
        self.portfolio_opp_btn.setProperty("variant", "secondary")
        self.portfolio_opp_btn.clicked.connect(self._show_portfolio_order_opps)
        self.portfolio_billing_btn = QPushButton("Faturação")
        self.portfolio_billing_btn.setProperty("variant", "secondary")
        self.portfolio_billing_btn.clicked.connect(self._open_portfolio_billing)
        orders_header.addWidget(self.portfolio_orders_title, 1)
        orders_header.addWidget(self.portfolio_order_btn)
        orders_header.addWidget(self.portfolio_opp_btn)
        orders_header.addWidget(self.portfolio_billing_btn)
        orders_card_layout.addLayout(orders_header)
        self.portfolio_orders_table = QTableWidget(0, 13)
        self.portfolio_orders_table.setHorizontalHeaderLabels(
            ["Encomenda", "OF", "Estado", "Entrega", "OPP", "Plan.", "Prod.", "Exp.", "%", "Adjudicado", "Faturado", "Por faturar", "Saldo"]
        )
        self.portfolio_orders_table.verticalHeader().setVisible(False)
        self.portfolio_orders_table.verticalHeader().setDefaultSectionSize(30)
        self.portfolio_orders_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.portfolio_orders_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.portfolio_orders_table.setSelectionMode(QAbstractItemView.SingleSelection)
        _set_table_columns(
            self.portfolio_orders_table,
            [
                (0, "fixed", 138), (1, "fixed", 120), (2, "stretch", 0), (3, "fixed", 82),
                (4, "fixed", 48), (5, "fixed", 56), (6, "fixed", 56), (7, "fixed", 56),
                (8, "fixed", 52), (9, "fixed", 96), (10, "fixed", 96), (11, "fixed", 96), (12, "fixed", 92),
            ],
        )
        self.portfolio_orders_table.itemSelectionChanged.connect(self._portfolio_order_selected)
        self.portfolio_orders_table.itemDoubleClicked.connect(lambda *_args: self._open_portfolio_order())
        orders_card_layout.addWidget(self.portfolio_orders_table)
        orders_layout.addWidget(orders_card, 1)

        pieces_card = CardFrame()
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(12, 10, 12, 10)
        pieces_header = QHBoxLayout()
        self.portfolio_pieces_title = QLabel("Produção da encomenda")
        self.portfolio_pieces_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.portfolio_pieces_meta = QLabel("Seleciona uma encomenda")
        self.portfolio_pieces_meta.setProperty("role", "muted")
        pieces_header.addWidget(self.portfolio_pieces_title)
        pieces_header.addStretch(1)
        pieces_header.addWidget(self.portfolio_pieces_meta)
        pieces_layout.addLayout(pieces_header)
        self.portfolio_pieces_table = QTableWidget(0, 12)
        self.portfolio_pieces_table.setHorizontalHeaderLabels(
            ["OPP", "Ref. Int.", "Ref. Ext.", "Descrição", "Material", "Esp.", "Estado", "Operação atual", "Plan.", "Prod.", "Exp.", "%"]
        )
        self.portfolio_pieces_table.verticalHeader().setVisible(False)
        self.portfolio_pieces_table.verticalHeader().setDefaultSectionSize(28)
        self.portfolio_pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.portfolio_pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        _set_table_columns(
            self.portfolio_pieces_table,
            [
                (0, "fixed", 138), (1, "fixed", 130), (2, "fixed", 140), (3, "stretch", 0),
                (4, "fixed", 88), (5, "fixed", 46), (6, "fixed", 104), (7, "stretch", 0),
                (8, "fixed", 52), (9, "fixed", 52), (10, "fixed", 52), (11, "fixed", 50),
            ],
        )
        self.portfolio_pieces_table.itemDoubleClicked.connect(lambda *_args: self._open_portfolio_piece())
        pieces_layout.addWidget(self.portfolio_pieces_table)
        orders_layout.addWidget(pieces_card, 1)
        splitter.addWidget(orders_host)
        splitter.setSizes([560, 1320])
        layout.addWidget(splitter, 1)

    def _portfolio_selected_client_code(self) -> str:
        return str(self.portfolio_client_combo.currentData() or "").strip()

    def _opp_main_tab_changed(self, _index: int) -> None:
        if self.main_tabs.currentWidget() is self.portfolio_page:
            self._refresh_portfolio()

    def _refresh_portfolio(self) -> None:
        if not hasattr(self, "portfolio_client_combo"):
            return
        selected_client = self._portfolio_selected_client_code()
        selected_year = str(self.portfolio_year_combo.currentData() or self.portfolio_year_combo.currentText() or "Todos").strip()
        selected_order = str((self._selected_portfolio_order() or {}).get("encomenda", "") or "").strip()
        payload = dict(
            self.backend.opp_client_portfolio(
                selected_client or "Todos",
                selected_year or "Todos",
                self.portfolio_search_edit.text().strip(),
            )
            or {}
        )
        self.portfolio_clients = list(payload.get("clients", []) or [])
        self.portfolio_orders = list(payload.get("orders", []) or [])

        years = ["Todos"] + [str(value) for value in list(payload.get("years", []) or []) if str(value).strip()]
        self.portfolio_year_combo.blockSignals(True)
        self.portfolio_year_combo.clear()
        for value in years:
            self.portfolio_year_combo.addItem(value, value)
        self.portfolio_year_combo.setCurrentText(selected_year if selected_year in years else "Todos")
        self.portfolio_year_combo.blockSignals(False)

        self.portfolio_client_combo.blockSignals(True)
        self.portfolio_client_combo.clear()
        self.portfolio_client_combo.addItem("Todos os clientes", "")
        for row in self.portfolio_clients:
            self.portfolio_client_combo.addItem(str(row.get("cliente", "") or "Sem cliente"), str(row.get("cliente_codigo", "") or ""))
        client_index = self.portfolio_client_combo.findData(selected_client)
        self.portfolio_client_combo.setCurrentIndex(client_index if client_index >= 0 else 0)
        self.portfolio_client_combo.blockSignals(False)

        self.portfolio_clients_table.blockSignals(True)
        _fill_table(
            self.portfolio_clients_table,
            [
                [
                    row.get("cliente", "-"), row.get("encomendas", 0), row.get("opp", 0),
                    _fmt_eur(row.get("adjudicado", 0)), _fmt_eur(row.get("faturado", 0)),
                    _fmt_eur(row.get("por_faturar", 0)), _fmt_eur(row.get("saldo_receber", 0)),
                ]
                for row in self.portfolio_clients
            ],
            align_center_from=1,
        )
        for index, row in enumerate(self.portfolio_clients):
            item = self.portfolio_clients_table.item(index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("cliente_codigo", "") or ""))
            if selected_client and str(row.get("cliente_codigo", "") or "") == selected_client:
                self.portfolio_clients_table.selectRow(index)
        self.portfolio_clients_table.blockSignals(False)
        self.portfolio_clients_meta.setText(f"{len(self.portfolio_clients)} clientes")

        self.portfolio_orders_table.blockSignals(True)
        _fill_table(
            self.portfolio_orders_table,
            [
                [
                    row.get("encomenda", "-"), row.get("of", "-"), row.get("estado", "-"), row.get("data_entrega", "-"),
                    row.get("opp_count", 0), self.backend._fmt(row.get("qtd_plan", 0)), self.backend._fmt(row.get("qtd_prod", 0)),
                    self.backend._fmt(row.get("qtd_exp", 0)), f"{float(row.get('progress', 0) or 0):.1f}%",
                    _fmt_eur(row.get("adjudicado", 0)), _fmt_eur(row.get("faturado", 0)),
                    _fmt_eur(row.get("por_faturar", 0)), _fmt_eur(row.get("saldo_receber", 0)),
                ]
                for row in self.portfolio_orders
            ],
            align_center_from=3,
        )
        target_index = 0
        for index, row in enumerate(self.portfolio_orders):
            item = self.portfolio_orders_table.item(index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("encomenda", "") or ""))
            _paint_table_row(self.portfolio_orders_table, index, str(row.get("estado", "")))
            if selected_order and str(row.get("encomenda", "") or "") == selected_order:
                target_index = index
        self.portfolio_orders_table.blockSignals(False)

        totals = dict(payload.get("totals", {}) or {})
        self.portfolio_stats[0].set_data(
            str(int(totals.get("encomendas", 0) or 0)),
            f"{int(totals.get('opp', 0) or 0)} OPP | {int(totals.get('clientes', 0) or 0)} clientes",
        )
        self.portfolio_stats[1].set_data(_fmt_eur(totals.get("adjudicado", 0)), "Valor das encomendas")
        self.portfolio_stats[2].set_data(_fmt_eur(totals.get("faturado", 0)), f"Recebido {_fmt_eur(totals.get('recebido', 0))}")
        self.portfolio_stats[3].set_data(_fmt_eur(totals.get("por_faturar", 0)), f"Saldo a receber {_fmt_eur(totals.get('saldo_receber', 0))}")

        selected_label = self.portfolio_client_combo.currentText() or "Todos os clientes"
        self.portfolio_orders_title.setText(f"Encomendas adjudicadas | {selected_label}")
        if self.portfolio_orders:
            self.portfolio_orders_table.selectRow(min(target_index, len(self.portfolio_orders) - 1))
            self._portfolio_order_selected()
        else:
            self.portfolio_pieces_table.setRowCount(0)
            self.portfolio_pieces_meta.setText("Sem encomendas para o filtro atual")
            self._sync_portfolio_buttons()

    def _portfolio_client_selected(self) -> None:
        row_index = _selected_row_index(self.portfolio_clients_table)
        if row_index < 0:
            return
        item = self.portfolio_clients_table.item(row_index, 0)
        client_code = str(item.data(Qt.UserRole) or "").strip() if item is not None else ""
        combo_index = self.portfolio_client_combo.findData(client_code)
        if combo_index >= 0 and combo_index != self.portfolio_client_combo.currentIndex():
            self.portfolio_client_combo.setCurrentIndex(combo_index)

    def _selected_portfolio_order(self) -> dict[str, Any]:
        row_index = _selected_row_index(self.portfolio_orders_table)
        if row_index < 0:
            return {}
        item = self.portfolio_orders_table.item(row_index, 0)
        order_number = str(item.data(Qt.UserRole) or item.text() or "").strip() if item is not None else ""
        return next((row for row in self.portfolio_orders if str(row.get("encomenda", "") or "").strip() == order_number), {})

    def _portfolio_order_selected(self) -> None:
        order = self._selected_portfolio_order()
        pieces = list(order.get("pieces", []) or [])
        _fill_table(
            self.portfolio_pieces_table,
            [
                [
                    row.get("opp", "-"), row.get("ref_interna", "-"), row.get("ref_externa", "-"), row.get("descricao", "-"),
                    row.get("material", "-"), row.get("espessura", "-"), row.get("estado", "-"), row.get("operacao_atual", "-"),
                    self.backend._fmt(row.get("qtd_plan", 0)), self.backend._fmt(row.get("qtd_prod", 0)),
                    self.backend._fmt(row.get("qtd_exp", 0)), f"{float(row.get('progress', 0) or 0):.1f}%",
                ]
                for row in pieces
            ],
            align_center_from=5,
        )
        for index, row in enumerate(pieces):
            item = self.portfolio_pieces_table.item(index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("opp", "") or ""))
            _paint_table_row(self.portfolio_pieces_table, index, str(row.get("estado", "")))
        if order:
            self.portfolio_pieces_meta.setText(
                f"{order.get('encomenda', '-')} | {len(pieces)} OPP | "
                f"Produzido {self.backend._fmt(order.get('qtd_prod', 0))}/{self.backend._fmt(order.get('qtd_plan', 0))} | "
                f"Expedido {self.backend._fmt(order.get('qtd_exp', 0))}"
            )
        else:
            self.portfolio_pieces_meta.setText("Seleciona uma encomenda")
        self._sync_portfolio_buttons()

    def _sync_portfolio_buttons(self) -> None:
        order = self._selected_portfolio_order()
        enabled = bool(order.get("encomenda"))
        self.portfolio_order_btn.setEnabled(enabled)
        self.portfolio_opp_btn.setEnabled(enabled and bool(order.get("pieces")))
        self.portfolio_billing_btn.setEnabled(enabled)

    def _open_portfolio_order(self) -> None:
        order_number = str((self._selected_portfolio_order() or {}).get("encomenda", "") or "").strip()
        if order_number:
            self._open_order_number(order_number)

    def _open_order_number(self, order_number: str) -> None:
        main_window = self.window()
        if not hasattr(main_window, "show_page"):
            return
        try:
            main_window.show_page("orders")
            page = getattr(main_window, "pages", {}).get("orders")
            if page is not None and hasattr(page, "open_order_numero"):
                page.open_order_numero(order_number)
        except Exception as exc:
            QMessageBox.critical(self, "Abrir encomenda", str(exc))

    def _show_portfolio_order_opps(self) -> None:
        order_number = str((self._selected_portfolio_order() or {}).get("encomenda", "") or "").strip()
        if not order_number:
            return
        self.main_tabs.setCurrentWidget(self.opp_page)
        self.filter_edit.setText(order_number)
        self.state_combo.setCurrentText("Todas")
        self.refresh()

    def _open_portfolio_piece(self) -> None:
        row_index = _selected_row_index(self.portfolio_pieces_table)
        if row_index < 0:
            return
        item = self.portfolio_pieces_table.item(row_index, 0)
        opp = str(item.data(Qt.UserRole) or item.text() or "").strip() if item is not None else ""
        if not opp:
            return
        self.main_tabs.setCurrentWidget(self.opp_page)
        self.filter_edit.setText(opp)
        self.state_combo.setCurrentText("Todas")
        self.refresh()

    def _open_portfolio_billing(self) -> None:
        order = self._selected_portfolio_order()
        order_number = str(order.get("encomenda", "") or "").strip()
        client_code = str(order.get("cliente_codigo", "") or self._portfolio_selected_client_code()).strip()
        main_window = self.window()
        if not hasattr(main_window, "show_page"):
            return
        try:
            main_window.show_page("billing")
            page = getattr(main_window, "pages", {}).get("billing")
            filter_widget = getattr(page, "filter_edit", None)
            filter_value = order_number or client_code
            if filter_widget is not None and filter_value:
                if hasattr(filter_widget, "setCurrentText"):
                    filter_widget.setCurrentText(filter_value)
                elif hasattr(filter_widget, "setText"):
                    filter_widget.setText(filter_value)
            if page is not None and hasattr(page, "refresh"):
                page.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Faturação", str(exc))

    def refresh(self) -> None:
        if self.main_tabs.currentWidget() is self.portfolio_page:
            self._refresh_portfolio()
        previous_opp = str(self.current_detail.get("opp", "") or "").strip()
        self.all_rows = list(self.backend.opp_rows("", "Todas", "Todos", "Todas", "Todos"))
        self.rows = list(
            self.backend.opp_rows(
                self.filter_edit.text().strip(),
                self.state_combo.currentText().strip() or "Ativas",
                self.year_combo.currentText().strip() or "Todos",
                self.operation_combo.currentText().strip() or "Todas",
                self.client_combo.currentText().strip() or "Todos",
            )
        )
        self._refresh_filter_options()
        _fill_table(
            self.table,
            [
                [
                    row.get("opp", "-"),
                    row.get("of", "-"),
                    row.get("encomenda", "-"),
                    row.get("cliente", "-"),
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("material", "-"),
                    row.get("espessura", "-"),
                    row.get("estado", "-"),
                    row.get("operacao_atual", "-"),
                    row.get("operador_atual", "-"),
                    self.backend._fmt(row.get("qtd_plan", 0)),
                    self.backend._fmt(row.get("qtd_prod", 0)),
                    self.backend._fmt(row.get("qtd_exp", 0)),
                    self.backend._fmt(row.get("tempo_real", 0)),
                    f"{float(row.get('progress', 0) or 0):.1f}%",
                ]
                for row in self.rows
            ],
            align_center_from=7,
        )
        for row_index, row in enumerate(self.rows):
            item = self.table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("opp", "") or "").strip())
            _paint_table_row(self.table, row_index, str(row.get("estado", "")))
        self._refresh_stats()
        if not self.rows:
            self._clear_detail()
            return
        target_row = 0
        if previous_opp:
            for index, row in enumerate(self.rows):
                if str(row.get("opp", "") or "").strip() == previous_opp:
                    target_row = index
                    break
        self.table.selectRow(target_row)
        self._on_selected()

    def _refresh_filter_options(self) -> None:
        current_year = self.year_combo.currentText().strip() or "Todos"
        current_client = self.client_combo.currentText().strip() or "Todos"
        current_operation = self.operation_combo.currentText().strip() or "Todas"
        years = sorted({str(row.get("ano", "") or "").strip() for row in self.all_rows if str(row.get("ano", "") or "").strip()}, reverse=True)
        clients = ["Todos"] + sorted(
            {
                str(row.get("cliente", "") or "").strip()
                for row in self.all_rows
                if str(row.get("cliente", "") or "").strip()
            }
        )
        operations = ["Todas"] + list(self.backend.opp_operations())
        year_values = ["Todos"] + years
        self.year_combo.blockSignals(True)
        self.client_combo.blockSignals(True)
        self.operation_combo.blockSignals(True)
        self.year_combo.clear()
        self.client_combo.clear()
        self.operation_combo.clear()
        self.year_combo.addItems(year_values)
        self.client_combo.addItems(clients)
        self.operation_combo.addItems(operations)
        self.year_combo.setCurrentText(current_year if current_year in year_values else "Todos")
        self.client_combo.setCurrentText(current_client if current_client in clients else "Todos")
        self.operation_combo.setCurrentText(current_operation if current_operation in operations else "Todas")
        self.year_combo.blockSignals(False)
        self.client_combo.blockSignals(False)
        self.operation_combo.blockSignals(False)

    def _refresh_stats(self) -> None:
        active = len([row for row in self.rows if "concl" not in self.backend.desktop_main.norm_text(row.get("estado", ""))])
        running = len([row for row in self.rows if ("produ" in self.backend.desktop_main.norm_text(row.get("estado", "")) or "incomplet" in self.backend.desktop_main.norm_text(row.get("estado", "")))])
        concluded = len([row for row in self.rows if "concl" in self.backend.desktop_main.norm_text(row.get("estado", ""))])
        shipped = len([row for row in self.rows if float(row.get("qtd_exp", 0) or 0) > 0])
        total_plan = sum(float(row.get("qtd_plan", 0) or 0) for row in self.rows)
        total_prod = sum(float(row.get("qtd_prod", 0) or 0) for row in self.rows)
        total_exp = sum(float(row.get("qtd_exp", 0) or 0) for row in self.rows)
        self.stats_cards[0].set_data(str(active), f"Filtro atual {len(self.rows)}")
        self.stats_cards[1].set_data(str(running), f"Concluidas {concluded}")
        self.stats_cards[2].set_data(self.backend._fmt(total_plan), f"Produzido {self.backend._fmt(total_prod)}")
        self.stats_cards[3].set_data(self.backend._fmt(total_exp), f"OPP expedidas {shipped}")

    def _export_csv(self) -> None:
        target, _selected = QFileDialog.getSaveFileName(
            self,
            "Exportar OPP",
            str(Path.home() / "opp_export.csv"),
            "CSV (*.csv)",
        )
        if not target:
            return
        with open(target, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(["OPP", "OF", "Encomenda", "Cliente", "Ref. Int.", "Ref. Ext.", "Material", "Esp.", "Estado", "Operacao", "Operador", "Qtd Plan", "Qtd Prod", "Qtd Expedida", "Tempo Real", "Ano"])
            for row in self.rows:
                writer.writerow(
                    [
                        row.get("opp", "-"),
                        row.get("of", "-"),
                        row.get("encomenda", "-"),
                        row.get("cliente", "-"),
                        row.get("ref_interna", "-"),
                        row.get("ref_externa", "-"),
                        row.get("material", "-"),
                        row.get("espessura", "-"),
                        row.get("estado", "-"),
                        row.get("operacao_atual", "-"),
                        row.get("operador_atual", "-"),
                        self.backend._fmt(row.get("qtd_plan", 0)),
                        self.backend._fmt(row.get("qtd_prod", 0)),
                        self.backend._fmt(row.get("qtd_exp", 0)),
                        self.backend._fmt(row.get("tempo_real", 0)),
                        row.get("ano", "-"),
                    ]
                )
        QMessageBox.information(self, "Exportar OPP", f"Exportado para:\n{target}")

    def _selected_row(self) -> dict[str, Any]:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        item = self.table.item(row_index, 0)
        opp = str(item.data(Qt.UserRole) or item.text() or "").strip() if item is not None else ""
        if opp:
            return next((row for row in self.rows if str(row.get("opp", "") or "").strip() == opp), {})
        return self.rows[row_index] if row_index < len(self.rows) else {}

    def _on_selected(self) -> None:
        row = self._selected_row()
        opp = str(row.get("opp", "") or "").strip()
        if not opp:
            self._clear_detail()
            return
        try:
            detail = self.backend.opp_detail(opp)
        except Exception as exc:
            self._clear_detail()
            self.meta_label.setText(str(exc))
            return
        self.current_detail = detail
        client_label = _format_client_label(
            f"{detail.get('cliente', '')} - {detail.get('cliente_nome', '')}".strip(" -"),
            show_name=True,
        )
        self.title_label.setText(f"{detail.get('opp', '-')} | {detail.get('ref_interna', '-')} | {detail.get('ref_externa', '-')}".strip(" |"))
        _apply_state_chip(self.state_chip, str(detail.get("estado", "-")))
        self.meta_label.setText(
            f"OF {detail.get('of', '-')} | Enc {detail.get('encomenda', '-')} | "
            f"{client_label} | "
            f"{detail.get('material', '-')} {detail.get('espessura', '-')} mm | "
            f"Plan {detail.get('qtd_plan', '0')} | Produzido {detail.get('qtd_prod', '0')} | Expedida {detail.get('qtd_exp', '0')}"
        )
        ops_done = len([op for op in list(detail.get("operacoes", []) or []) if "concl" in self.backend.desktop_main.norm_text(op.get("estado", ""))])
        ops_total = len(list(detail.get("operacoes", []) or []))
        self.flow_label.setText(
            f"Descricao: {detail.get('descricao', '-') or '-'} | Inicio: {detail.get('inicio', '-') or '-'} | "
            f"Fim: {detail.get('fim', '-') or '-'} | Tempo real: {detail.get('tempo_real', '0')} min | "
            f"Operacoes concluidas {ops_done}/{ops_total}"
        )
        self.progress_bar.setValue(int(round(float(detail.get("progress", 0) or 0))))
        _fill_table(
            self.ops_table,
            [
                [
                    row.get("nome", "-"),
                    row.get("estado", "-"),
                    row.get("user", "-"),
                    row.get("inicio", "-"),
                    row.get("fim", "-"),
                    row.get("qtd_ok", "0"),
                    row.get("qtd_nok", "0"),
                    row.get("qtd_qual", "0"),
                    f"{float(row.get('progress', 0) or 0):.1f}%",
                ]
                for row in list(detail.get("operacoes", []) or [])
            ],
            align_center_from=3,
        )
        for row_index, op_row in enumerate(list(detail.get("operacoes", []) or [])):
            _paint_table_row(self.ops_table, row_index, str(op_row.get("estado", "")))
        _fill_table(
            self.events_table,
            [
                [
                    row.get("data", "-"),
                    row.get("evento", "-"),
                    row.get("operacao", "-"),
                    row.get("operador", "-"),
                    row.get("qtd_ok", "0"),
                    row.get("qtd_nok", "0"),
                    row.get("info", "-"),
                ]
                for row in list(detail.get("events", []) or [])
            ],
            align_center_from=4,
        )
        _fill_table(
            self.exp_table,
            [
                [
                    row.get("guia", "-"),
                    row.get("data", "-"),
                    row.get("estado", "-"),
                    row.get("destinatario", "-"),
                    row.get("qtd", "0"),
                    row.get("obs", "-"),
                ]
                for row in list(detail.get("expedicoes", []) or [])
            ],
            align_center_from=4,
        )
        for row_index, exp_row in enumerate(list(detail.get("expedicoes", []) or [])):
            _paint_table_row(self.exp_table, row_index, str(exp_row.get("estado", "")))
        self._sync_buttons()

    def _sync_buttons(self) -> None:
        has_detail = bool(self.current_detail.get("opp"))
        self.pdf_btn.setEnabled(has_detail)
        self.order_btn.setEnabled(has_detail)
        self.drawing_btn.setEnabled(has_detail and bool(self.current_detail.get("desenho_path")))

    def _clear_detail(self) -> None:
        self.current_detail = {}
        self.title_label.setText("Sem OPP selecionada")
        _apply_state_chip(self.state_chip, "-")
        self.meta_label.setText("Seleciona uma OPP para ver detalhe de fabrico, tempos, eventos e expedicao.")
        self.flow_label.setText("-")
        self.progress_bar.setValue(0)
        self.ops_table.setRowCount(0)
        self.events_table.setRowCount(0)
        self.exp_table.setRowCount(0)
        self._sync_buttons()

    def _open_label_pdf(self) -> None:
        opp = str(self.current_detail.get("opp", "") or "").strip()
        if not opp:
            return
        try:
            self.backend.opp_open_pdf(opp)
        except Exception as exc:
            QMessageBox.critical(self, "Etiqueta OPP", str(exc))

    def _open_drawing(self) -> None:
        opp = str(self.current_detail.get("opp", "") or "").strip()
        if not opp:
            return
        try:
            self.backend.opp_open_drawing(opp)
        except Exception as exc:
            QMessageBox.critical(self, "Ver desenho", str(exc))

    def _open_order(self) -> None:
        numero = str(self.current_detail.get("encomenda", "") or "").strip()
        if not numero:
            return
        self._open_order_number(numero)
