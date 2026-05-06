from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard


def _text(value: Any, fallback: str = "-") -> str:
    raw = str(value or "").strip()
    return raw or fallback


def _tone_colors(tone: str) -> tuple[str, str]:
    tone_txt = str(tone or "default").strip().lower()
    if tone_txt == "danger":
        return "#fff1f2", "#b42318"
    if tone_txt == "warning":
        return "#fff8e6", "#a16207"
    if tone_txt == "success":
        return "#ecfdf3", "#166534"
    if tone_txt == "info":
        return "#eef4ff", "#1d4ed8"
    return "#f8fafc", "#334155"


def _make_item(value: Any, *, align: Qt.AlignmentFlag | None = None, user_data: Any = None) -> QTableWidgetItem:
    item = QTableWidgetItem(_text(value))
    if align is not None:
        item.setTextAlignment(int(align))
    if user_data is not None:
        item.setData(Qt.UserRole, user_data)
    return item


def _paint_row(table: QTableWidget, row: int, tone: str) -> None:
    bg_hex, fg_hex = _tone_colors(tone)
    bg = QBrush(QColor(bg_hex))
    fg = QBrush(QColor(fg_hex))
    for col in range(table.columnCount()):
        item = table.item(row, col)
        if item is None:
            continue
        item.setBackground(bg)
        item.setForeground(fg)


class StockDashboardPage(QWidget):
    page_title = "Dashboard"
    page_subtitle = "Visao operacional clara do estado das encomendas, stock, compras e logistica."
    uses_backend_reload = True
    allow_auto_timer_refresh = False

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.operational_payload: dict[str, Any] = {}
        self.finance_payload: dict[str, Any] = {}
        self.order_rows: list[dict[str, Any]] = []
        self.action_rows: list[dict[str, Any]] = []
        self.logistics_rows: list[dict[str, Any]] = []
        self.montagem_rows: list[dict[str, Any]] = []
        self._op_view_key = "radar"
        self._stock_view_key = "overview"

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        toolbar = CardFrame()
        toolbar.set_tone("info")
        toolbar_layout = QVBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(18, 16, 18, 16)
        toolbar_layout.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        title_col = QVBoxLayout()
        title_col.setContentsMargins(0, 0, 0, 0)
        title_col.setSpacing(3)
        title = QLabel("Painel de comando")
        title.setStyleSheet("font-size: 20px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel(
            "Leitura rapida do que esta em curso, do que precisa de acao e do que impacta entregas e compras."
        )
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        top_row.addLayout(title_col, 1)

        top_row.addWidget(QLabel("Ano"))
        self.year_combo = QComboBox()
        self.year_combo.setMinimumWidth(120)
        self.year_combo.currentTextChanged.connect(self.refresh)
        top_row.addWidget(self.year_combo)

        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.clicked.connect(self.refresh)
        self.open_order_btn = QPushButton("Abrir encomenda")
        self.open_order_btn.setProperty("variant", "secondary")
        self.open_order_btn.clicked.connect(self._open_selected_order)
        self.open_transport_btn = QPushButton("Abrir transporte")
        self.open_transport_btn.setProperty("variant", "secondary")
        self.open_transport_btn.clicked.connect(self._open_selected_transport)
        self.create_note_btn = QPushButton("Gerar nota montagem")
        self.create_note_btn.setProperty("variant", "secondary")
        self.create_note_btn.clicked.connect(self._create_montagem_purchase_note)
        self.export_btn = QPushButton("Exportar CSV")
        self.export_btn.setProperty("variant", "secondary")
        self.export_btn.clicked.connect(self._export_csv)
        for widget in (
            self.refresh_btn,
            self.open_order_btn,
            self.open_transport_btn,
            self.create_note_btn,
            self.export_btn,
        ):
            widget.setProperty("toolbarAction", "true")
            widget.setMinimumHeight(36)
            top_row.addWidget(widget)
        toolbar_layout.addLayout(top_row)

        self.updated_label = QLabel("Sem leitura carregada.")
        self.updated_label.setProperty("role", "muted")
        toolbar_layout.addWidget(self.updated_label)
        root.addWidget(toolbar)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        root.addWidget(self.tabs, 1)

        self._build_operation_tab()
        self._build_stock_tab()

    def _build_operation_tab(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)
        self.operational_cards = [StatCard(title) for title in (
            "Encomendas abertas",
            "Em risco",
            "Prontas a expedir",
            "A aguardar transporte",
            "Em transporte",
            "Montagem pendente",
        )]
        for idx, card in enumerate(self.operational_cards):
            cards_layout.addWidget(card, idx // 3, idx % 3)
        layout.addWidget(cards_host)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)

        left_card = CardFrame()
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(14, 14, 14, 14)
        left_layout.setSpacing(10)
        left_title = QLabel("Fluxo do dia")
        left_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        left_layout.addWidget(left_title)
        left_subtitle = QLabel("Resumo de fases e leitura da encomenda atualmente selecionada.")
        left_subtitle.setProperty("role", "muted")
        left_subtitle.setWordWrap(True)
        left_layout.addWidget(left_subtitle)

        self.phase_table = self._build_table(["Fase", "Total"])
        self.phase_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.phase_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        left_layout.addWidget(self.phase_table, 1)

        detail_card = CardFrame()
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(14, 12, 14, 12)
        detail_layout.setSpacing(6)
        self.op_detail_title = QLabel("Seleciona uma encomenda")
        self.op_detail_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.op_detail_meta = QLabel("O detalhe da leitura operacional aparece aqui.")
        self.op_detail_meta.setProperty("role", "muted")
        self.op_detail_meta.setWordWrap(True)
        self.op_detail_body = QLabel("")
        self.op_detail_body.setWordWrap(True)
        self.op_detail_body.setStyleSheet("color: #0f172a;")
        detail_layout.addWidget(self.op_detail_title)
        detail_layout.addWidget(self.op_detail_meta)
        detail_layout.addWidget(self.op_detail_body)
        left_layout.addWidget(detail_card)

        split.addWidget(left_card)

        right_host = QWidget()
        right_layout = QVBoxLayout(right_host)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        op_menu_card = CardFrame()
        op_menu_layout = QHBoxLayout(op_menu_card)
        op_menu_layout.setContentsMargins(12, 10, 12, 10)
        op_menu_layout.setSpacing(8)
        self.op_buttons: dict[str, QPushButton] = {}
        for key, label in (("radar", "Radar"), ("actions", "Acoes"), ("logistics", "Logistica")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setProperty("dashboardSegment", "true")
            btn.setMinimumHeight(38)
            btn.setMinimumWidth(104)
            btn.clicked.connect(lambda checked=False, view_key=key: self._set_operational_view(view_key))
            self.op_buttons[key] = btn
            op_menu_layout.addWidget(btn)
        op_menu_layout.addStretch(1)
        right_layout.addWidget(op_menu_card)

        self.operational_stack = QStackedWidget()

        radar_page = QWidget()
        radar_layout = QVBoxLayout(radar_page)
        radar_layout.setContentsMargins(0, 0, 0, 0)
        radar_layout.setSpacing(10)
        radar_card = CardFrame()
        radar_card.set_tone("info")
        radar_card_layout = QVBoxLayout(radar_card)
        radar_card_layout.setContentsMargins(14, 12, 14, 12)
        radar_card_layout.setSpacing(8)
        radar_title = QLabel("Radar operacional")
        radar_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        radar_card_layout.addWidget(radar_title)
        radar_subtitle = QLabel("Estado real de cada encomenda no fluxo: laser, montagem, expedicao e transporte.")
        radar_subtitle.setProperty("role", "muted")
        radar_subtitle.setWordWrap(True)
        radar_card_layout.addWidget(radar_subtitle)
        self.order_table = self._build_table(
            ["Encomenda", "Cliente", "Fase", "Laser", "Montagem", "Expedicao", "Transporte", "Sinal", "Planeado"]
        )
        self.order_table.itemSelectionChanged.connect(self._sync_operational_detail)
        radar_widths = [150, 270, 170, 170, 130, 130, 125, 200, 110]
        for column_index, width in enumerate(radar_widths):
            self.order_table.horizontalHeader().setSectionResizeMode(column_index, QHeaderView.Interactive)
            self.order_table.setColumnWidth(column_index, width)
        radar_card_layout.addWidget(self.order_table, 1)
        radar_layout.addWidget(radar_card)
        self.operational_stack.addWidget(radar_page)

        actions_page = QWidget()
        actions_layout = QVBoxLayout(actions_page)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(10)
        actions_card = CardFrame()
        actions_card.set_tone("warning")
        actions_card_layout = QVBoxLayout(actions_card)
        actions_card_layout.setContentsMargins(14, 12, 14, 12)
        actions_card_layout.setSpacing(8)
        actions_title = QLabel("Acoes imediatas")
        actions_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        actions_card_layout.addWidget(actions_title)
        actions_subtitle = QLabel("Mostra apenas o que exige decisao ou desbloqueio no momento.")
        actions_subtitle.setProperty("role", "muted")
        actions_subtitle.setWordWrap(True)
        actions_card_layout.addWidget(actions_subtitle)
        self.action_table = self._build_table(["Encomenda", "Cliente", "Motivo", "Acao sugerida", "Entrega"])
        self.action_table.itemSelectionChanged.connect(self._sync_operational_detail)
        self.action_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.action_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.action_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.action_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.action_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        actions_card_layout.addWidget(self.action_table, 1)
        actions_layout.addWidget(actions_card)
        self.operational_stack.addWidget(actions_page)

        logistics_page = QWidget()
        logistics_layout = QVBoxLayout(logistics_page)
        logistics_layout.setContentsMargins(0, 0, 0, 0)
        logistics_layout.setSpacing(10)
        logistics_card = CardFrame()
        logistics_card.set_tone("success")
        logistics_card_layout = QVBoxLayout(logistics_card)
        logistics_card_layout.setContentsMargins(14, 12, 14, 12)
        logistics_card_layout.setSpacing(8)
        logistics_title = QLabel("Logistica e transporte")
        logistics_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        logistics_card_layout.addWidget(logistics_title)
        logistics_subtitle = QLabel("Encomendas com guia, viagem ativa, subcontrato ou carga a nosso cargo.")
        logistics_subtitle.setProperty("role", "muted")
        logistics_subtitle.setWordWrap(True)
        logistics_card_layout.addWidget(logistics_subtitle)
        self.logistics_table = self._build_table(
            ["Encomenda", "Cliente", "Guia", "Viagem", "Transportadora", "Zona", "Peso", "Estado"]
        )
        self.logistics_table.itemSelectionChanged.connect(self._sync_operational_detail)
        self.logistics_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.logistics_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.logistics_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        logistics_card_layout.addWidget(self.logistics_table, 1)
        logistics_layout.addWidget(logistics_card)
        self.operational_stack.addWidget(logistics_page)

        right_layout.addWidget(self.operational_stack, 1)
        split.addWidget(right_host)
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 3)
        layout.addWidget(split, 1)

        self.tabs.addTab(page, "Operacao")
        self._set_operational_view("radar")

    def _build_stock_tab(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)
        self.finance_cards = [StatCard(title) for title in (
            "Stock MP",
            "Stock Produtos",
            "Compras MP",
            "Compras Produtos",
            "Stock Total",
            "NE Aprovadas",
        )]
        for idx, card in enumerate(self.finance_cards):
            cards_layout.addWidget(card, idx // 3, idx % 3)
        layout.addWidget(cards_host)

        stock_menu_card = CardFrame()
        stock_menu_layout = QHBoxLayout(stock_menu_card)
        stock_menu_layout.setContentsMargins(12, 10, 12, 10)
        stock_menu_layout.setSpacing(8)
        self.stock_buttons: dict[str, QPushButton] = {}
        for key, label in (("overview", "Visao geral"), ("purchases", "Compras"), ("montagem", "Montagem")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setProperty("dashboardSegment", "true")
            btn.setMinimumHeight(38)
            btn.setMinimumWidth(118)
            btn.clicked.connect(lambda checked=False, view_key=key: self._set_stock_view(view_key))
            self.stock_buttons[key] = btn
            stock_menu_layout.addWidget(btn)
        stock_menu_layout.addStretch(1)
        layout.addWidget(stock_menu_card)

        self.stock_stack = QStackedWidget()

        overview_page = QWidget()
        overview_layout = QHBoxLayout(overview_page)
        overview_layout.setContentsMargins(0, 0, 0, 0)
        overview_layout.setSpacing(12)
        materias_card = CardFrame()
        materias_layout = QVBoxLayout(materias_card)
        materias_layout.setContentsMargins(14, 12, 14, 12)
        materias_layout.setSpacing(8)
        materias_title = QLabel("Top materia-prima")
        materias_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        materias_layout.addWidget(materias_title)
        self.top_materias_table = self._build_table(["ID", "Material", "Esp.", "Qtd.", "Valor"])
        self.top_materias_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.top_materias_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.top_materias_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.top_materias_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.top_materias_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        materias_layout.addWidget(self.top_materias_table, 1)
        overview_layout.addWidget(materias_card, 1)

        produtos_card = CardFrame()
        produtos_layout = QVBoxLayout(produtos_card)
        produtos_layout.setContentsMargins(14, 12, 14, 12)
        produtos_layout.setSpacing(8)
        produtos_title = QLabel("Top produto acabado")
        produtos_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        produtos_layout.addWidget(produtos_title)
        self.top_produtos_table = self._build_table(["Codigo", "Descricao", "Qtd.", "Preco", "Valor"])
        self.top_produtos_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.top_produtos_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.top_produtos_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.top_produtos_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.top_produtos_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        produtos_layout.addWidget(self.top_produtos_table, 1)
        overview_layout.addWidget(produtos_card, 1)
        self.stock_stack.addWidget(overview_page)

        purchases_page = QWidget()
        purchases_layout = QVBoxLayout(purchases_page)
        purchases_layout.setContentsMargins(0, 0, 0, 0)
        purchases_layout.setSpacing(12)

        purchase_top = QSplitter(Qt.Horizontal)
        purchase_top.setChildrenCollapsible(False)

        fornecedor_card = CardFrame()
        fornecedor_layout = QVBoxLayout(fornecedor_card)
        fornecedor_layout.setContentsMargins(14, 12, 14, 12)
        fornecedor_layout.setSpacing(8)
        fornecedor_title = QLabel("Compras por fornecedor")
        fornecedor_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        fornecedor_layout.addWidget(fornecedor_title)
        self.compras_fornecedor_table = self._build_table(["Fornecedor", "Total"])
        self.compras_fornecedor_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.compras_fornecedor_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        fornecedor_layout.addWidget(self.compras_fornecedor_table, 1)
        purchase_top.addWidget(fornecedor_card)

        mes_card = CardFrame()
        mes_layout = QVBoxLayout(mes_card)
        mes_layout.setContentsMargins(14, 12, 14, 12)
        mes_layout.setSpacing(8)
        mes_title = QLabel("Compras por mes")
        mes_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        mes_layout.addWidget(mes_title)
        self.compras_mes_table = self._build_table(["Mes", "Total"])
        self.compras_mes_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.compras_mes_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        mes_layout.addWidget(self.compras_mes_table, 1)
        purchase_top.addWidget(mes_card)
        purchase_top.setStretchFactor(0, 1)
        purchase_top.setStretchFactor(1, 1)
        purchases_layout.addWidget(purchase_top, 1)

        compras_card = CardFrame()
        compras_card.set_tone("warning")
        compras_card_layout = QVBoxLayout(compras_card)
        compras_card_layout.setContentsMargins(14, 12, 14, 12)
        compras_card_layout.setSpacing(8)
        compras_title = QLabel("Ultimas compras recebidas")
        compras_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        compras_card_layout.addWidget(compras_title)
        self.compras_table = self._build_table(["Data", "NE", "Fornecedor", "Artigo", "Qtd.", "Preco", "Total", "Origem"])
        self.compras_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.compras_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.compras_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.compras_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.compras_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.compras_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.compras_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.compras_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        compras_card_layout.addWidget(self.compras_table, 1)
        purchases_layout.addWidget(compras_card, 2)
        self.stock_stack.addWidget(purchases_page)

        montagem_page = QWidget()
        montagem_layout = QVBoxLayout(montagem_page)
        montagem_layout.setContentsMargins(0, 0, 0, 0)
        montagem_layout.setSpacing(12)
        montagem_card = CardFrame()
        montagem_card.set_tone("danger")
        montagem_card_layout = QVBoxLayout(montagem_card)
        montagem_card_layout.setContentsMargins(14, 12, 14, 12)
        montagem_card_layout.setSpacing(8)
        montagem_title = QLabel("Montagem por validar")
        montagem_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        montagem_card_layout.addWidget(montagem_title)
        montagem_subtitle = QLabel("Faltas de stock e necessidades para gerar nota de compra automaticamente.")
        montagem_subtitle.setProperty("role", "muted")
        montagem_subtitle.setWordWrap(True)
        montagem_card_layout.addWidget(montagem_subtitle)
        self.montagem_table = self._build_table(
            ["Encomenda", "Cliente", "Montagem", "Tempo", "Qtd. falta", "Fornecedor", "Stock", "Entrega"]
        )
        self.montagem_table.itemSelectionChanged.connect(self._sync_stock_detail)
        self.montagem_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.montagem_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.montagem_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.montagem_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.montagem_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.montagem_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.montagem_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.montagem_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        montagem_card_layout.addWidget(self.montagem_table, 1)
        montagem_layout.addWidget(montagem_card, 1)

        montagem_detail = CardFrame()
        montagem_detail_layout = QVBoxLayout(montagem_detail)
        montagem_detail_layout.setContentsMargins(14, 12, 14, 12)
        montagem_detail_layout.setSpacing(6)
        self.stock_detail_title = QLabel("Seleciona uma encomenda")
        self.stock_detail_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.stock_detail_meta = QLabel("A leitura da montagem aparece aqui.")
        self.stock_detail_meta.setProperty("role", "muted")
        self.stock_detail_meta.setWordWrap(True)
        self.stock_detail_body = QLabel("")
        self.stock_detail_body.setWordWrap(True)
        montagem_detail_layout.addWidget(self.stock_detail_title)
        montagem_detail_layout.addWidget(self.stock_detail_meta)
        montagem_detail_layout.addWidget(self.stock_detail_body)
        montagem_layout.addWidget(montagem_detail)
        self.stock_stack.addWidget(montagem_page)

        layout.addWidget(self.stock_stack, 1)
        self.tabs.addTab(page, "Stock e Compras")
        self._set_stock_view("overview")

    def _build_table(self, headers: list[str]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(32)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)
        table.setWordWrap(False)
        table.setTextElideMode(Qt.ElideRight)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        table.horizontalHeader().setMinimumSectionSize(84)
        table.horizontalHeader().setStretchLastSection(False)
        table.setStyleSheet(
            """
            QTableWidget {
                font-size: 12px;
                gridline-color: #d8e1ec;
            }
            QTableWidget::item {
                padding: 4px 8px;
            }
            QHeaderView::section {
                padding: 8px 10px;
                font-size: 12px;
                font-weight: 800;
            }
            """
        )
        return table

    def _set_operational_view(self, key: str) -> None:
        self._op_view_key = key
        mapping = {"radar": 0, "actions": 1, "logistics": 2}
        self.operational_stack.setCurrentIndex(mapping.get(key, 0))
        for button_key, button in self.op_buttons.items():
            button.setChecked(button_key == key)
        self._sync_actions()

    def _set_stock_view(self, key: str) -> None:
        self._stock_view_key = key
        mapping = {"overview": 0, "purchases": 1, "montagem": 2}
        self.stock_stack.setCurrentIndex(mapping.get(key, 0))
        for button_key, button in self.stock_buttons.items():
            button.setChecked(button_key == key)
        self._sync_actions()

    def refresh(self) -> None:
        year = str(self.year_combo.currentText() or "Todos").strip() or "Todos"
        finance_preview = dict(self.backend.finance_dashboard(year) or {})
        years = ["Todos"]
        years.extend([str(value) for value in list(finance_preview.get("years", []) or []) if str(value).strip()])
        self._sync_year_combo(years, finance_preview.get("selected_year", year))
        selected_year = str(self.year_combo.currentText() or finance_preview.get("selected_year", year) or "Todos").strip() or "Todos"
        if selected_year == year:
            self.finance_payload = finance_preview
        else:
            self.finance_payload = dict(self.backend.finance_dashboard(selected_year) or {})
        self.operational_payload = dict(self.backend.operational_dashboard(selected_year) or {})
        self._fill_operational()
        self._fill_finance()
        updated_txt = _text(self.operational_payload.get("updated_at", ""), "")
        self.updated_label.setText(
            f"Atualizado: {updated_txt[:19].replace('T', ' ')}" if updated_txt else "Atualizado."
        )
        self._sync_actions()

    def _sync_year_combo(self, years: list[str], selected: Any) -> None:
        clean: list[str] = []
        seen: set[str] = set()
        for value in years:
            text = str(value or "").strip()
            if not text or text in seen:
                continue
            clean.append(text)
            seen.add(text)
        current = str(selected or self.year_combo.currentText() or "Todos").strip() or "Todos"
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        for value in clean or ["Todos"]:
            self.year_combo.addItem(value)
        index = self.year_combo.findText(current)
        self.year_combo.setCurrentIndex(index if index >= 0 else 0)
        self.year_combo.blockSignals(False)

    def _fill_operational(self) -> None:
        cards = list(self.operational_payload.get("cards", []) or [])
        for card, payload in zip(self.operational_cards, cards):
            card.set_data(_text(payload.get("value")), _text(payload.get("subtitle"), ""))
            card.set_tone(_text(payload.get("tone"), "default"))
        for card in self.operational_cards[len(cards):]:
            card.set_data("-", "")
            card.set_tone("default")

        phase_rows = list(self.operational_payload.get("phase_rows", []) or [])
        self.phase_table.setRowCount(len(phase_rows))
        for row_index, row in enumerate(phase_rows):
            self.phase_table.setItem(row_index, 0, _make_item(row.get("fase", "-")))
            self.phase_table.setItem(
                row_index,
                1,
                _make_item(row.get("total", "0"), align=Qt.AlignRight | Qt.AlignVCenter),
            )

        self.order_rows = list(self.operational_payload.get("order_rows", []) or [])
        self.order_table.setRowCount(len(self.order_rows))
        for row_index, row in enumerate(self.order_rows):
            self.order_table.setItem(row_index, 0, _make_item(row.get("numero", "-"), user_data=row.get("numero", "")))
            self.order_table.setItem(row_index, 1, _make_item(row.get("cliente", "-")))
            self.order_table.setItem(row_index, 2, _make_item(row.get("fase", "-")))
            self.order_table.setItem(row_index, 3, _make_item(row.get("laser", "-")))
            self.order_table.setItem(row_index, 4, _make_item(row.get("montagem", "-")))
            self.order_table.setItem(row_index, 5, _make_item(row.get("expedicao", "-")))
            self.order_table.setItem(row_index, 6, _make_item(row.get("transporte_numero", "-")))
            self.order_table.setItem(row_index, 7, _make_item(row.get("sinal", "-")))
            self.order_table.setItem(row_index, 8, _make_item(row.get("laser_planeado", "-")))
            for col_index in range(self.order_table.columnCount()):
                item = self.order_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(item.text())
            _paint_row(self.order_table, row_index, str(row.get("signal_tone", "default") or "default"))

        self.action_rows = list(self.operational_payload.get("action_rows", []) or [])
        self.action_table.setRowCount(len(self.action_rows))
        for row_index, row in enumerate(self.action_rows):
            self.action_table.setItem(row_index, 0, _make_item(row.get("numero", "-"), user_data=row.get("numero", "")))
            self.action_table.setItem(row_index, 1, _make_item(row.get("cliente", "-")))
            self.action_table.setItem(row_index, 2, _make_item(row.get("motivo", "-")))
            self.action_table.setItem(row_index, 3, _make_item(row.get("acao", "-")))
            self.action_table.setItem(row_index, 4, _make_item(row.get("entrega", "-")))
            _paint_row(self.action_table, row_index, str(row.get("tone", "default") or "default"))

        self.logistics_rows = list(self.operational_payload.get("logistics_rows", []) or [])
        self.logistics_table.setRowCount(len(self.logistics_rows))
        for row_index, row in enumerate(self.logistics_rows):
            self.logistics_table.setItem(row_index, 0, _make_item(row.get("numero", "-"), user_data=row.get("numero", "")))
            self.logistics_table.setItem(row_index, 1, _make_item(row.get("cliente", "-")))
            self.logistics_table.setItem(row_index, 2, _make_item(row.get("guia", "-")))
            self.logistics_table.setItem(row_index, 3, _make_item(row.get("transporte", "-"), user_data=row.get("transporte", "")))
            self.logistics_table.setItem(row_index, 4, _make_item(row.get("transportadora", "-")))
            self.logistics_table.setItem(row_index, 5, _make_item(row.get("zona", "-")))
            self.logistics_table.setItem(row_index, 6, _make_item(row.get("peso", "-")))
            self.logistics_table.setItem(row_index, 7, _make_item(row.get("estado", "-")))
            _paint_row(self.logistics_table, row_index, str(row.get("tone", "default") or "default"))

        if self.order_rows:
            self.order_table.selectRow(0)
        elif self.action_rows:
            self.action_table.selectRow(0)
        elif self.logistics_rows:
            self.logistics_table.selectRow(0)
        else:
            self._clear_operational_detail()

    def _fill_finance(self) -> None:
        cards = list(self.finance_payload.get("cards", []) or [])
        for card, payload in zip(self.finance_cards, cards):
            card.set_data(_text(payload.get("value")), _text(payload.get("subtitle"), ""))
            card.set_tone(_text(payload.get("tone"), "default"))
        for card in self.finance_cards[len(cards):]:
            card.set_data("-", "")
            card.set_tone("default")

        self._fill_simple_table(
            self.top_materias_table,
            list(self.finance_payload.get("top_materias", []) or []),
            [
                lambda row: row.get("id", "-"),
                lambda row: row.get("material", "-"),
                lambda row: row.get("espessura", "-"),
                lambda row: row.get("qty", "-"),
                lambda row: row.get("valor", "-"),
            ],
        )
        self._fill_simple_table(
            self.top_produtos_table,
            list(self.finance_payload.get("top_produtos", []) or []),
            [
                lambda row: row.get("codigo", "-"),
                lambda row: row.get("descricao", "-"),
                lambda row: row.get("qty", "-"),
                lambda row: row.get("preco_unid", "-"),
                lambda row: row.get("valor", "-"),
            ],
        )
        self._fill_simple_table(
            self.compras_fornecedor_table,
            list(self.finance_payload.get("compras_por_fornecedor", []) or []),
            [
                lambda row: row.get("fornecedor", "-"),
                lambda row: row.get("total", "-"),
            ],
        )
        self._fill_simple_table(
            self.compras_mes_table,
            list(self.finance_payload.get("compras_por_mes", []) or []),
            [
                lambda row: row.get("mes", "-"),
                lambda row: row.get("total", "-"),
            ],
        )
        self._fill_simple_table(
            self.compras_table,
            list(self.finance_payload.get("compras", []) or []),
            [
                lambda row: row.get("data", "-"),
                lambda row: row.get("ne", "-"),
                lambda row: row.get("fornecedor", "-"),
                lambda row: row.get("artigo", "-"),
                lambda row: row.get("qtd", "-"),
                lambda row: row.get("preco", "-"),
                lambda row: row.get("total", "-"),
                lambda row: row.get("origem", "-"),
            ],
        )

        self.montagem_rows = list(self.finance_payload.get("montagem_alertas", []) or [])
        self.montagem_table.setRowCount(len(self.montagem_rows))
        for row_index, row in enumerate(self.montagem_rows):
            self.montagem_table.setItem(row_index, 0, _make_item(row.get("numero", "-"), user_data=row.get("numero", "")))
            self.montagem_table.setItem(row_index, 1, _make_item(row.get("cliente", "-")))
            self.montagem_table.setItem(row_index, 2, _make_item(row.get("montagem", "-")))
            self.montagem_table.setItem(row_index, 3, _make_item(row.get("tempo_min", "-")))
            self.montagem_table.setItem(row_index, 4, _make_item(row.get("qtd_falta", "-")))
            self.montagem_table.setItem(row_index, 5, _make_item(row.get("fornecedor", "-")))
            self.montagem_table.setItem(row_index, 6, _make_item(row.get("stock", "-")))
            self.montagem_table.setItem(row_index, 7, _make_item(row.get("data_entrega", "-")))
            tone = "danger" if float(row.get("qtd_falta", 0) or 0) > 0 else "warning"
            _paint_row(self.montagem_table, row_index, tone)

        if self.montagem_rows:
            self.montagem_table.selectRow(0)
        else:
            self._clear_stock_detail()

    def _fill_simple_table(self, table: QTableWidget, rows: list[dict[str, Any]], columns: list[Any]) -> None:
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, getter in enumerate(columns):
                table.setItem(row_index, col_index, _make_item(getter(row)))

    def _sync_operational_detail(self) -> None:
        row = self._current_operational_row()
        if not row:
            self._clear_operational_detail()
            return
        numero = _text(row.get("numero", ""))
        cliente = _text(row.get("cliente", "-"))
        fase = _text(row.get("fase", row.get("motivo", "-")))
        self.op_detail_title.setText(numero or "Sem encomenda")
        self.op_detail_meta.setText(f"{cliente} | {fase}")
        lines: list[str] = []
        if row in self.order_rows:
            lines.extend(
                [
                    f"Laser: {_text(row.get('laser', '-'))}",
                    f"Planeado: {_text(row.get('laser_planeado', '-'))}",
                    f"Montagem: {_text(row.get('montagem', '-'))}",
                    f"Expedicao: {_text(row.get('expedicao', '-'))}",
                    f"Transporte: {_text(row.get('transporte_numero', '-'))} | {_text(row.get('transporte_estado', '-'))}",
                    f"Sinal: {_text(row.get('sinal', '-'))}",
                    f"Proxima acao: {_text(row.get('next_action', '-'))}",
                ]
            )
        elif row in self.action_rows:
            lines.extend(
                [
                    f"Motivo: {_text(row.get('motivo', '-'))}",
                    f"Acao: {_text(row.get('acao', '-'))}",
                    f"Entrega: {_text(row.get('entrega', '-'))}",
                ]
            )
        else:
            lines.extend(
                [
                    f"Guia: {_text(row.get('guia', '-'))}",
                    f"Viagem: {_text(row.get('transporte', '-'))}",
                    f"Transportadora: {_text(row.get('transportadora', '-'))}",
                    f"Zona: {_text(row.get('zona', '-'))}",
                    f"Peso: {_text(row.get('peso', '-'))}",
                    f"Estado: {_text(row.get('estado', '-'))}",
                ]
            )
        self.op_detail_body.setText("\n".join(lines))

    def _sync_stock_detail(self) -> None:
        row = self._current_montagem_row()
        if not row:
            self._clear_stock_detail()
            return
        self.stock_detail_title.setText(_text(row.get("numero", "Seleciona uma encomenda")))
        self.stock_detail_meta.setText(f"{_text(row.get('cliente', '-'))} | {_text(row.get('montagem', '-'))}")
        self.stock_detail_body.setText(
            "\n".join(
                [
                    f"Tempo previsto: {_text(row.get('tempo_min', '-'))} min",
                    f"Qtd. em falta: {_text(row.get('qtd_falta', '-'))}",
                    f"Fornecedor sugerido: {_text(row.get('fornecedor', '-'))}",
                    f"Stock: {_text(row.get('stock', '-'))}",
                    f"Entrega: {_text(row.get('data_entrega', '-'))}",
                ]
            )
        )

    def _clear_operational_detail(self) -> None:
        self.op_detail_title.setText("Seleciona uma encomenda")
        self.op_detail_meta.setText("O detalhe da leitura operacional aparece aqui.")
        self.op_detail_body.setText("")

    def _clear_stock_detail(self) -> None:
        self.stock_detail_title.setText("Seleciona uma encomenda")
        self.stock_detail_meta.setText("A leitura da montagem aparece aqui.")
        self.stock_detail_body.setText("")

    def _current_operational_row(self) -> dict[str, Any]:
        if not hasattr(self, "order_table"):
            return {}
        if self._op_view_key == "actions":
            return self._current_row_from_table(self.action_table, self.action_rows)
        if self._op_view_key == "logistics":
            return self._current_row_from_table(self.logistics_table, self.logistics_rows)
        return self._current_row_from_table(self.order_table, self.order_rows)

    def _current_montagem_row(self) -> dict[str, Any]:
        if not hasattr(self, "montagem_table"):
            return {}
        return self._current_row_from_table(self.montagem_table, self.montagem_rows)

    def _current_row_from_table(self, table: QTableWidget, rows: list[dict[str, Any]]) -> dict[str, Any]:
        current = table.currentItem()
        if current is None:
            return {}
        row_index = current.row()
        if row_index < 0 or row_index >= len(rows):
            return {}
        return rows[row_index]

    def _selected_order_number(self) -> str:
        for row in (self._current_operational_row(), self._current_montagem_row()):
            numero = str(row.get("numero", "") or "").strip()
            if numero:
                return numero
        return ""

    def _selected_transport_number(self) -> str:
        row = self._current_row_from_table(self.logistics_table, self.logistics_rows)
        return str(row.get("transporte", "") or "").strip()

    def _open_selected_order(self) -> None:
        numero = self._selected_order_number()
        if not numero:
            QMessageBox.information(self, "Dashboard", "Seleciona uma encomenda primeiro.")
            return
        main_window = self.window()
        if hasattr(main_window, "show_page"):
            try:
                main_window.show_page("orders")
                page = getattr(main_window, "pages", {}).get("orders")
                if page is not None and hasattr(page, "open_order_numero"):
                    page.open_order_numero(numero)
                    return
            except Exception as exc:
                QMessageBox.warning(self, "Dashboard", str(exc))
                return
        QMessageBox.information(self, "Dashboard", f"Encomenda selecionada: {numero}")

    def _open_selected_transport(self) -> None:
        numero = self._selected_transport_number()
        if not numero or numero == "-":
            QMessageBox.information(self, "Dashboard", "Seleciona uma linha de logistica com viagem ativa.")
            return
        main_window = self.window()
        if hasattr(main_window, "show_page"):
            try:
                main_window.show_page("transportes")
                page = getattr(main_window, "pages", {}).get("transportes")
                if page is not None and hasattr(page, "open_trip_numero"):
                    page.open_trip_numero(numero)
                    return
            except Exception as exc:
                QMessageBox.warning(self, "Dashboard", str(exc))
                return
        QMessageBox.information(self, "Dashboard", f"Viagem selecionada: {numero}")

    def _create_montagem_purchase_note(self) -> None:
        row = self._current_montagem_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.information(self, "Dashboard", "Seleciona uma encomenda com falta de montagem.")
            return
        try:
            result = self.backend.ne_create_from_montagem_shortages([numero])
        except Exception as exc:
            QMessageBox.critical(self, "Dashboard", str(exc))
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
        message = f"Nota {note_number or '-'} criada para a encomenda {numero}."
        if missing:
            message += "\n\nFornecedor por validar em: " + ", ".join(missing)
        QMessageBox.information(self, "Dashboard", message)

    def _export_csv(self) -> None:
        table, suggested_name = self._current_export_table()
        if table is None:
            QMessageBox.information(self, "Dashboard", "Nao existe uma tabela ativa para exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exportar CSV", suggested_name, "CSV (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow([table.horizontalHeaderItem(col).text() for col in range(table.columnCount())])
                for row in range(table.rowCount()):
                    writer.writerow(
                        [
                            str(table.item(row, col).text()).strip() if table.item(row, col) is not None else ""
                            for col in range(table.columnCount())
                        ]
                    )
        except Exception as exc:
            QMessageBox.critical(self, "Dashboard", str(exc))
            return
        QMessageBox.information(self, "Dashboard", f"CSV guardado em:\n{Path(path)}")

    def _current_export_table(self) -> tuple[QTableWidget | None, str]:
        if self.tabs.currentIndex() == 0:
            mapping = {
                "radar": (self.order_table, "dashboard_operacao_radar.csv"),
                "actions": (self.action_table, "dashboard_operacao_acoes.csv"),
                "logistics": (self.logistics_table, "dashboard_operacao_logistica.csv"),
            }
            return mapping.get(self._op_view_key, (self.order_table, "dashboard_operacao.csv"))
        mapping = {
            "overview": (self.top_materias_table, "dashboard_stock_visao_geral.csv"),
            "purchases": (self.compras_table, "dashboard_compras.csv"),
            "montagem": (self.montagem_table, "dashboard_montagem.csv"),
        }
        return mapping.get(self._stock_view_key, (self.compras_table, "dashboard_stock.csv"))

    def _sync_actions(self) -> None:
        has_order = bool(self._selected_order_number())
        selected_trip = self._selected_transport_number()
        has_transport = bool(selected_trip and selected_trip != "-")
        has_montagem = bool(self._current_montagem_row())
        self.open_order_btn.setEnabled(has_order)
        self.open_transport_btn.setEnabled(has_transport)
        self.create_note_btn.setEnabled(has_montagem and self.tabs.currentIndex() == 1 and self._stock_view_key == "montagem")
