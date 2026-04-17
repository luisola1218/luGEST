from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QGridLayout, QHeaderView, QLabel, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

from ..widgets import CardFrame, StatCard


class HomePage(QWidget):
    page_title = "Resumo"
    page_subtitle = "Visão rápida do estado do sistema para iniciar os testes."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        cards_host = QWidget()
        self.cards_layout = QGridLayout(cards_host)
        self.cards_layout.setContentsMargins(0, 0, 0, 0)
        self.cards_layout.setHorizontalSpacing(12)
        self.cards_layout.setVerticalSpacing(12)
        self.cards: list[StatCard] = []
        for index in range(4):
            card = StatCard("-")
            self.cards.append(card)
            self.cards_layout.addWidget(card, 0, index)
        for card, tone in zip(self.cards, ("info", "success", "warning", "default")):
            card.set_tone(tone)
        root.addWidget(cards_host)

        log_card = CardFrame()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(16, 14, 16, 14)
        log_layout.setSpacing(10)
        title = QLabel("Histórico de Stock")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel("Últimos movimentos de matéria-prima e retalhos.")
        subtitle.setProperty("role", "muted")
        log_layout.addWidget(title)
        log_layout.addWidget(subtitle)
        self.log_table = QTableWidget(0, 3)
        self.log_table.setHorizontalHeaderLabels(["Data", "Acao", "Detalhes"])
        self.log_table.verticalHeader().setVisible(False)
        self.log_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.log_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.log_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.log_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        log_layout.addWidget(self.log_table)
        root.addWidget(log_card, 1)

    def refresh(self) -> None:
        counts = self.backend.dashboard_counts()
        for card, payload in zip(self.cards, counts):
            card.title_label.setText(payload["title"])
            card.set_data(payload["value"], payload["subtitle"])

        rows = self.backend.stock_log_rows()
        self.log_table.setSortingEnabled(False)
        self.log_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, key in enumerate(("data", "acao", "detalhes")):
                item = QTableWidgetItem(str(row[key]))
                if col_index < 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.log_table.setItem(row_index, col_index, item)
        self.log_table.setSortingEnabled(True)

