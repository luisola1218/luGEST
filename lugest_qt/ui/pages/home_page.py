from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QGridLayout, QHBoxLayout, QHeaderView, QLabel, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

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
        root.setSpacing(16)

        hero_card = CardFrame()
        hero_card.set_tone("info")
        hero_layout = QHBoxLayout(hero_card)
        hero_layout.setContentsMargins(18, 16, 18, 16)
        hero_layout.setSpacing(16)

        hero_text = QVBoxLayout()
        hero_text.setContentsMargins(0, 0, 0, 0)
        hero_text.setSpacing(4)
        hero_title = QLabel("Resumo executivo")
        hero_title.setStyleSheet("font-size: 20px; color: #0f172a;")
        hero_subtitle = QLabel("Leitura rápida do estado atual do ambiente, do stock e dos movimentos recentes.")
        hero_subtitle.setProperty("role", "muted")
        hero_subtitle.setWordWrap(True)
        self.hero_summary_label = QLabel("A carregar indicadores principais.")
        self.hero_summary_label.setStyleSheet("color: #17324d; font-size: 12px;")
        self.hero_summary_label.setWordWrap(True)
        hero_text.addWidget(hero_title)
        hero_text.addWidget(hero_subtitle)
        hero_text.addWidget(self.hero_summary_label)
        hero_layout.addLayout(hero_text, 1)

        status_wrap = QVBoxLayout()
        status_wrap.setContentsMargins(0, 0, 0, 0)
        status_wrap.setSpacing(6)
        status_label = QLabel("Atualização")
        status_label.setProperty("role", "muted")
        self.updated_label = QLabel("Sem leitura carregada.")
        self.updated_label.setStyleSheet(
            "padding: 8px 12px; border: 1px solid #cfd8e3; border-radius: 10px; "
            "background: #ffffff; color: #16324b; font-size: 12px;"
        )
        self.updated_label.setWordWrap(True)
        status_wrap.addWidget(status_label)
        status_wrap.addWidget(self.updated_label)
        hero_layout.addLayout(status_wrap)
        root.addWidget(hero_card)

        cards_host = QWidget()
        self.cards_layout = QGridLayout(cards_host)
        self.cards_layout.setContentsMargins(0, 0, 0, 0)
        self.cards_layout.setHorizontalSpacing(12)
        self.cards_layout.setVerticalSpacing(12)
        self.cards: list[StatCard] = []
        for index in range(4):
            card = StatCard("-")
            self.cards.append(card)
            self.cards_layout.addWidget(card, index // 2, index % 2)
        for card, tone in zip(self.cards, ("info", "success", "warning", "default")):
            card.set_tone(tone)
        root.addWidget(cards_host)

        log_card = CardFrame()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(16, 14, 16, 14)
        log_layout.setSpacing(10)
        title = QLabel("Atividade recente")
        title.setStyleSheet("font-size: 18px; color: #0f172a;")
        subtitle = QLabel("Últimos movimentos de matéria-prima e retalhos, prontos para validação rápida.")
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
        hero_bits = [f"{payload['title']}: {payload['value']}" for payload in counts[:3]]
        self.hero_summary_label.setText(" | ".join(hero_bits) if hero_bits else "Sem indicadores disponíveis.")

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

