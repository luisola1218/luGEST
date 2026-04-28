from __future__ import annotations

from PySide6.QtWidgets import (
    QHeaderView,
    QLabel,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame
from .runtime_common import (
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    set_panel_tone as _set_panel_tone,
)


class AvariasPage(QWidget):
    page_title = "Avarias"
    page_subtitle = "Avarias abertas e histórico consolidado para não deixar ocorrências esquecidas."
    allow_auto_timer_refresh = True

    def __init__(self, runtime_service, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        self.banner_card = CardFrame()
        banner_layout = QVBoxLayout(self.banner_card)
        banner_layout.setContentsMargins(16, 14, 16, 14)
        title = QLabel("Estado das avarias")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.banner_label = QLabel("-")
        self.banner_label.setWordWrap(True)
        banner_layout.addWidget(title)
        banner_layout.addWidget(self.banner_label)
        self.banner_card.set_tone("default")
        root.addWidget(self.banner_card)

        self.open_table = QTableWidget(0, 6)
        self.open_table.setHorizontalHeaderLabels(["Encomenda", "Posto", "Operador", "Motivo", "Início", "Duração"])
        self.open_table.verticalHeader().setVisible(False)
        self.open_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.open_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table = QTableWidget(0, 7)
        self.history_table.setHorizontalHeaderLabels(["Encomenda", "Ref. Int.", "Motivo", "Estado", "Criada", "Fechada", "Duração"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        for title_text, table in (("Em aberto", self.open_table), ("Histórico", self.history_table)):
            card = CardFrame()
            layout = QVBoxLayout(card)
            layout.setContentsMargins(16, 14, 16, 14)
            title = QLabel(title_text)
            title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
            layout.addWidget(title)
            layout.addWidget(table)
            card.set_tone("danger" if "aberto" in title_text.lower() else "info")
            root.addWidget(card, 1)

    def refresh(self) -> None:
        data = self.runtime_service.avarias()
        summary = data.get("summary", {})
        headline = str(summary.get("headline") or "Sem avarias abertas no momento.")
        self.banner_label.setText(headline)
        _set_panel_tone(self.banner_card, "danger" if summary.get("critical") else "default")
        _fill_table(
            self.open_table,
            [[r.get("encomenda", "-"), r.get("posto", "-"), r.get("operador", "-"), r.get("motivo", "-"), r.get("created_at", "-"), r.get("duracao_txt", "-")] for r in data.get("open", [])],
            align_center_from=4,
        )
        for row_index, _row in enumerate(data.get("open", []) or []):
            _paint_table_row(self.open_table, row_index, "Avaria")
        _fill_table(
            self.history_table,
            [[r.get("encomenda", "-"), r.get("ref_interna", "-"), r.get("motivo", "-"), r.get("estado", "-"), r.get("created_at", "-"), r.get("fechada_at", "-"), r.get("duracao_txt", "-")] for r in data.get("history", [])],
            align_center_from=3,
        )
        for row_index, row in enumerate(data.get("history", []) or []):
            _paint_table_row(self.history_table, row_index, str(row.get("estado", "")))
