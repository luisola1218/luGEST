from __future__ import annotations

import unicodedata

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard
from .runtime_common import state_visual


def _fmt_eur(value: object) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    return f"{number:,.2f} EUR".replace(",", " ").replace(".", ",")


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).casefold()


class HomePage(QWidget):
    page_title = "Resumo"
    page_subtitle = "Indicadores executivos, vendas e acompanhamento das ordens de fabrico."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.sales_orders: list[dict] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setStyleSheet(
            "QTabBar::tab { min-width: 150px; padding: 9px 16px; font-weight: 800; }"
            "QTabWidget::pane { border: 1px solid #cbd8e6; border-radius: 6px; background: #f8fafc; }"
        )
        root.addWidget(self.tabs)

        self.overview_page = QWidget()
        self.sales_page = QWidget()
        self.tabs.addTab(self.overview_page, "Visão geral")
        self.tabs.addTab(self.sales_page, "Vendas e OF")
        self._build_overview()
        self._build_sales()

    @staticmethod
    def _configure_table(table: QTableWidget, *, stretch_columns: tuple[int, ...] = ()) -> None:
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(32)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.setShowGrid(False)
        table.setStyleSheet(
            "QTableWidget { font-size: 11px; alternate-background-color: #f8fafc;"
            " background: #ffffff; border: 1px solid #d6e0eb; }"
            "QTableWidget::item { padding: 4px 7px; border-bottom: 1px solid #edf2f7; }"
            "QTableWidget::item:selected { background: #dbeafe; color: #102a43; }"
            "QHeaderView::section { background: #eef4f8; color: #243b53; border: none;"
            " border-bottom: 1px solid #cbd8e6; padding: 7px 8px; font-size: 10px; font-weight: 900; }"
        )
        header = table.horizontalHeader()
        for column in range(table.columnCount()):
            header.setSectionResizeMode(
                column,
                QHeaderView.Stretch if column in stretch_columns else QHeaderView.ResizeToContents,
            )

    def _build_overview(self) -> None:
        root = QVBoxLayout(self.overview_page)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(14)

        hero_card = CardFrame()
        hero_card.set_tone("info")
        hero_layout = QHBoxLayout(hero_card)
        hero_layout.setContentsMargins(18, 15, 18, 15)
        hero_layout.setSpacing(16)

        hero_text = QVBoxLayout()
        hero_text.setSpacing(4)
        hero_title = QLabel("Resumo executivo")
        hero_title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        hero_subtitle = QLabel("Leitura imediata de stock, encomendas e atividade recente.")
        hero_subtitle.setProperty("role", "muted")
        self.hero_summary_label = QLabel("A carregar indicadores principais.")
        self.hero_summary_label.setStyleSheet("color: #17324d; font-size: 12px; font-weight: 700;")
        self.hero_summary_label.setWordWrap(True)
        hero_text.addWidget(hero_title)
        hero_text.addWidget(hero_subtitle)
        hero_text.addWidget(self.hero_summary_label)
        hero_layout.addLayout(hero_text, 1)

        status_wrap = QVBoxLayout()
        status_wrap.setSpacing(5)
        status_label = QLabel("Atualização")
        status_label.setProperty("role", "muted")
        self.updated_label = QLabel("Ao abrir a página")
        self.updated_label.setStyleSheet(
            "padding: 8px 12px; border: 1px solid #cbd8e6; border-radius: 6px;"
            "background: #ffffff; color: #16324b; font-size: 11px; font-weight: 700;"
        )
        status_wrap.addWidget(status_label)
        status_wrap.addWidget(self.updated_label)
        hero_layout.addLayout(status_wrap)
        root.addWidget(hero_card)

        cards_host = QWidget()
        self.cards_layout = QGridLayout(cards_host)
        self.cards_layout.setContentsMargins(0, 0, 0, 0)
        self.cards_layout.setHorizontalSpacing(10)
        self.cards_layout.setVerticalSpacing(10)
        self.cards: list[StatCard] = []
        for index in range(4):
            card = StatCard("-")
            card.set_tone(("info", "success", "warning", "default")[index])
            self.cards.append(card)
            self.cards_layout.addWidget(card, 0, index)
        root.addWidget(cards_host)

        log_card = CardFrame()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(16, 13, 16, 13)
        log_layout.setSpacing(8)
        title = QLabel("Atividade recente de stock")
        title.setStyleSheet("font-size: 15px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel("Movimentos de matéria-prima e retalhos com rastreabilidade.")
        subtitle.setProperty("role", "muted")
        log_layout.addWidget(title)
        log_layout.addWidget(subtitle)
        self.log_table = QTableWidget(0, 3)
        self.log_table.setHorizontalHeaderLabels(["Data", "Ação", "Detalhes"])
        self._configure_table(self.log_table, stretch_columns=(2,))
        log_layout.addWidget(self.log_table)
        root.addWidget(log_card, 1)

    def _build_sales(self) -> None:
        root = QVBoxLayout(self.sales_page)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 11, 14, 11)
        filters_layout.setSpacing(9)
        self.sales_client_combo = QComboBox()
        self.sales_client_combo.setMinimumWidth(250)
        self.sales_year_combo = QComboBox()
        self.sales_year_combo.setMinimumWidth(100)
        self.sales_state_combo = QComboBox()
        self.sales_state_combo.addItems(["Todas", "Em produção", "Concluídas", "Por faturar", "Faturadas"])
        self.sales_state_combo.setMinimumWidth(130)
        self.sales_search_edit = QLineEdit()
        self.sales_search_edit.setPlaceholderText("Encomenda, OF, cliente, orçamento ou referência")
        self.sales_refresh_btn = QPushButton("Atualizar")
        self.sales_refresh_btn.setProperty("variant", "secondary")
        filters_layout.addWidget(QLabel("Cliente"))
        filters_layout.addWidget(self.sales_client_combo)
        filters_layout.addWidget(QLabel("Ano"))
        filters_layout.addWidget(self.sales_year_combo)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.sales_state_combo)
        filters_layout.addWidget(self.sales_search_edit, 1)
        filters_layout.addWidget(self.sales_refresh_btn)
        root.addWidget(filters)

        stats_host = QWidget()
        stats_layout = QGridLayout(stats_host)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setHorizontalSpacing(9)
        self.sales_stats: list[StatCard] = []
        for index, (title, tone) in enumerate(
            (
                ("Ordens de fabrico", "info"),
                ("Adjudicado", "success"),
                ("Faturado", "info"),
                ("Por faturar", "warning"),
                ("Recebido", "default"),
            )
        ):
            card = StatCard(title)
            card.set_tone(tone)
            self.sales_stats.append(card)
            stats_layout.addWidget(card, 0, index)
        root.addWidget(stats_host)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        orders_card = CardFrame()
        orders_layout = QVBoxLayout(orders_card)
        orders_layout.setContentsMargins(14, 12, 14, 12)
        orders_layout.setSpacing(8)
        orders_header = QHBoxLayout()
        orders_title = QLabel("Carteira de vendas e produção")
        orders_title.setStyleSheet("font-size: 15px; font-weight: 900; color: #0f172a;")
        self.sales_meta_label = QLabel("0 OF")
        self.sales_meta_label.setProperty("role", "muted")
        self.open_order_btn = QPushButton("Abrir encomenda")
        self.open_order_btn.setProperty("variant", "secondary")
        self.open_billing_btn = QPushButton("Abrir faturação")
        self.open_billing_btn.setProperty("variant", "secondary")
        orders_header.addWidget(orders_title)
        orders_header.addWidget(self.sales_meta_label)
        orders_header.addStretch(1)
        orders_header.addWidget(self.open_order_btn)
        orders_header.addWidget(self.open_billing_btn)
        self.sales_table = QTableWidget(0, 10)
        self.sales_table.setHorizontalHeaderLabels(
            ["Encomenda", "OF", "Cliente", "Estado", "Produção", "Adjudicado", "Faturado", "Por faturar", "Recebido", "Entrega"]
        )
        self._configure_table(self.sales_table, stretch_columns=(2,))
        orders_layout.addLayout(orders_header)
        orders_layout.addWidget(self.sales_table)
        splitter.addWidget(orders_card)

        pieces_card = CardFrame()
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(14, 12, 14, 12)
        pieces_layout.setSpacing(8)
        pieces_header = QHBoxLayout()
        pieces_title = QLabel("Produção da OF selecionada")
        pieces_title.setStyleSheet("font-size: 14px; font-weight: 900; color: #0f172a;")
        self.sales_pieces_meta = QLabel("Seleciona uma ordem de fabrico")
        self.sales_pieces_meta.setProperty("role", "muted")
        pieces_header.addWidget(pieces_title)
        pieces_header.addStretch(1)
        pieces_header.addWidget(self.sales_pieces_meta)
        self.sales_pieces_table = QTableWidget(0, 8)
        self.sales_pieces_table.setHorizontalHeaderLabels(
            ["OPP", "Referência", "Descrição", "Operação atual", "Estado", "Planeado", "Produzido", "Progresso"]
        )
        self._configure_table(self.sales_pieces_table, stretch_columns=(2, 3))
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.sales_pieces_table)
        splitter.addWidget(pieces_card)
        splitter.setSizes([420, 230])
        root.addWidget(splitter, 1)

        self.sales_client_combo.currentIndexChanged.connect(self._refresh_sales)
        self.sales_year_combo.currentIndexChanged.connect(self._refresh_sales)
        self.sales_state_combo.currentIndexChanged.connect(self._refresh_sales)
        self.sales_search_edit.textChanged.connect(self._refresh_sales)
        self.sales_refresh_btn.clicked.connect(self._refresh_sales)
        self.sales_table.itemSelectionChanged.connect(self._refresh_sales_pieces)
        self.sales_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_order())
        self.open_order_btn.clicked.connect(self._open_selected_order)
        self.open_billing_btn.clicked.connect(self._open_selected_billing)

    def _selected_sales_order(self) -> dict:
        row_index = self.sales_table.currentRow()
        if row_index < 0:
            return {}
        item = self.sales_table.item(row_index, 0)
        number = str(item.data(Qt.UserRole) or item.text() or "").strip() if item is not None else ""
        return next(
            (row for row in self.sales_orders if str(row.get("encomenda", "") or "").strip() == number),
            {},
        )

    def _sales_state_matches(self, row: dict) -> bool:
        selected = _norm(self.sales_state_combo.currentText())
        progress = float(row.get("progress", 0) or 0)
        state = _norm(row.get("estado", ""))
        if selected in {"", "todas"}:
            return True
        if "produc" in selected:
            return (0 < progress < 100) or "produc" in state or "curso" in state
        if "conclu" in selected:
            return progress >= 100 or "conclu" in state
        if "por faturar" in selected:
            return float(row.get("por_faturar", 0) or 0) > 0
        if "faturad" in selected:
            return float(row.get("faturado", 0) or 0) > 0
        return True

    def _refresh_sales(self, *_args) -> None:
        if not hasattr(self, "sales_table"):
            return
        selected_client = str(self.sales_client_combo.currentData() or "").strip()
        selected_year = str(self.sales_year_combo.currentData() or self.sales_year_combo.currentText() or "Todos").strip()
        previous_order = str(self._selected_sales_order().get("encomenda", "") or "").strip()
        try:
            payload = dict(
                self.backend.opp_client_portfolio(
                    selected_client or "Todos",
                    selected_year or "Todos",
                    self.sales_search_edit.text().strip(),
                )
                or {}
            )
        except Exception as exc:
            QMessageBox.critical(self, "Vendas e OF", str(exc))
            return

        clients = list(payload.get("clients", []) or [])
        years = ["Todos"] + [str(value) for value in list(payload.get("years", []) or []) if str(value) != "Todos"]
        self.sales_client_combo.blockSignals(True)
        self.sales_client_combo.clear()
        self.sales_client_combo.addItem("Todos os clientes", "")
        for row in clients:
            self.sales_client_combo.addItem(
                str(row.get("cliente", "") or "Sem cliente"),
                str(row.get("cliente_codigo", "") or ""),
            )
        client_index = self.sales_client_combo.findData(selected_client)
        self.sales_client_combo.setCurrentIndex(client_index if client_index >= 0 else 0)
        self.sales_client_combo.blockSignals(False)

        self.sales_year_combo.blockSignals(True)
        self.sales_year_combo.clear()
        for value in years:
            self.sales_year_combo.addItem(value, value)
        self.sales_year_combo.setCurrentText(selected_year if selected_year in years else "Todos")
        self.sales_year_combo.blockSignals(False)

        self.sales_orders = [dict(row) for row in list(payload.get("orders", []) or []) if self._sales_state_matches(dict(row))]
        self.sales_table.setRowCount(len(self.sales_orders))
        target_row = 0
        for row_index, row in enumerate(self.sales_orders):
            if str(row.get("encomenda", "") or "").strip() == previous_order:
                target_row = row_index
            values = [
                row.get("encomenda", "-"),
                row.get("of", "-") or "-",
                row.get("cliente", "-"),
                row.get("estado", "-"),
                f"{float(row.get('progress', 0) or 0):.1f}%",
                _fmt_eur(row.get("adjudicado", 0)),
                _fmt_eur(row.get("faturado", 0)),
                _fmt_eur(row.get("por_faturar", 0)),
                _fmt_eur(row.get("recebido", 0)),
                row.get("data_entrega", "-") or "-",
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, str(row.get("encomenda", "") or "").strip())
                    item.setFont(QFont(item.font().family(), item.font().pointSize(), QFont.Bold))
                if column in (0, 1, 3, 4, 9):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                if column in (5, 6, 7, 8):
                    item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
                self.sales_table.setItem(row_index, column, item)
            visual = state_visual(str(row.get("estado", "") or ""))
            state_item = self.sales_table.item(row_index, 3)
            if state_item is not None:
                state_item.setBackground(QBrush(QColor(visual["bg"])))
                state_item.setForeground(QBrush(QColor(visual["fg"])))
            progress_item = self.sales_table.item(row_index, 4)
            if progress_item is not None:
                progress_item.setForeground(QBrush(QColor("#166534" if float(row.get("progress", 0) or 0) >= 100 else "#1d4ed8")))

        totals = {
            key: round(sum(float(row.get(key, 0) or 0) for row in self.sales_orders), 2)
            for key in ("adjudicado", "faturado", "por_faturar", "recebido")
        }
        self.sales_stats[0].set_data(str(len(self.sales_orders)), "OF no filtro atual")
        self.sales_stats[1].set_data(_fmt_eur(totals["adjudicado"]), "Valor comercial")
        self.sales_stats[2].set_data(_fmt_eur(totals["faturado"]), "Documentos emitidos")
        self.sales_stats[3].set_data(_fmt_eur(totals["por_faturar"]), "Carteira pendente")
        self.sales_stats[4].set_data(_fmt_eur(totals["recebido"]), "Pagamentos registados")
        self.sales_meta_label.setText(f"{len(self.sales_orders)} OF")
        if self.sales_orders:
            self.sales_table.selectRow(min(target_row, len(self.sales_orders) - 1))
        else:
            self.sales_pieces_table.setRowCount(0)
            self.sales_pieces_meta.setText("Sem OF para o filtro atual")
        self._refresh_sales_pieces()

    def _refresh_sales_pieces(self) -> None:
        order = self._selected_sales_order()
        pieces = [dict(row) for row in list(order.get("pieces", []) or [])]
        self.sales_pieces_table.setRowCount(len(pieces))
        for row_index, row in enumerate(pieces):
            values = [
                row.get("opp", "-"),
                row.get("ref_externa", row.get("ref_interna", "-")) or "-",
                row.get("descricao", "-") or "-",
                row.get("operacao_atual", "-") or "-",
                row.get("estado", "-") or "-",
                f"{float(row.get('qtd_plan', 0) or 0):g}",
                f"{float(row.get('qtd_prod', 0) or 0):g}",
                f"{float(row.get('progress', 0) or 0):.1f}%",
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column >= 4:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.sales_pieces_table.setItem(row_index, column, item)
            visual = state_visual(str(row.get("estado", "") or ""))
            state_item = self.sales_pieces_table.item(row_index, 4)
            if state_item is not None:
                state_item.setBackground(QBrush(QColor(visual["bg"])))
                state_item.setForeground(QBrush(QColor(visual["fg"])))
        if order:
            self.sales_pieces_meta.setText(
                f"{order.get('encomenda', '-')} · {len(pieces)} OPP · "
                f"{float(order.get('progress', 0) or 0):.1f}% produzido"
            )
        self.open_order_btn.setEnabled(bool(order))
        self.open_billing_btn.setEnabled(bool(order))

    def _open_selected_order(self) -> None:
        order_number = str(self._selected_sales_order().get("encomenda", "") or "").strip()
        main_window = self.window()
        if not order_number or not hasattr(main_window, "show_page"):
            return
        main_window.show_page("orders")
        page = getattr(main_window, "pages", {}).get("orders")
        if page is not None and hasattr(page, "open_order_numero"):
            page.open_order_numero(order_number)

    def _open_selected_billing(self) -> None:
        order = self._selected_sales_order()
        order_number = str(order.get("encomenda", "") or "").strip()
        main_window = self.window()
        if not order_number or not hasattr(main_window, "show_page"):
            return
        main_window.show_page("billing")
        page = getattr(main_window, "pages", {}).get("billing")
        filter_widget = getattr(page, "filter_edit", None)
        if filter_widget is not None:
            if hasattr(filter_widget, "setCurrentText"):
                filter_widget.setCurrentText(order_number)
            elif hasattr(filter_widget, "setText"):
                filter_widget.setText(order_number)
        if page is not None and hasattr(page, "refresh"):
            page.refresh()

    def refresh(self) -> None:
        counts = self.backend.dashboard_counts()
        for card, payload in zip(self.cards, counts):
            card.title_label.setText(payload["title"])
            card.set_data(payload["value"], payload["subtitle"])
        hero_bits = [f"{payload['title']}: {payload['value']}" for payload in counts[:3]]
        self.hero_summary_label.setText(" | ".join(hero_bits) if hero_bits else "Sem indicadores disponíveis.")

        rows = self.backend.stock_log_rows()
        self.log_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, key in enumerate(("data", "acao", "detalhes")):
                item = QTableWidgetItem(str(row[key]))
                if col_index < 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.log_table.setItem(row_index, col_index, item)
        self.updated_label.setText("Atualizado agora")
        self._refresh_sales()
