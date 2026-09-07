from __future__ import annotations
import html
import json
import os
import tempfile
from PySide6.QtCore import QMimeData, QTime, QTimer, Qt
from PySide6.QtGui import QBrush, QColor, QDrag
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote
from .runtime_common import (
    configure_table as _configure_table,
    elide_middle as _elide_middle,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    run_process_async as _run_process_async,
    selected_row_index as _selected_row_index,
    table_visible_height as _table_visible_height,
)
from .runtime_support import _adopt_layout_item, _clear_layout_widgets, _is_dark, _take_layout_items
from ..widgets import CardFrame, StatCard


class PlanningBacklogTable(QTableWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragOnly)
        self.setDefaultDropAction(Qt.CopyAction)

    def mimeData(self, items):  # type: ignore[override]
        mime = QMimeData()
        if items:
            first = items[0]
            source_item = self.item(first.row(), 0) or first
            payload = source_item.data(Qt.UserRole)
            try:
                raw = json.dumps(payload or {})
            except Exception:
                raw = "{}"
            mime.setData("application/x-lugest-planning-item", raw.encode("utf-8"))
        return mime


class PlanningGridTable(QTableWidget):
    def __init__(self, page_ref, parent=None) -> None:
        super().__init__(parent)
        self.page_ref = page_ref
        self._drag_start_pos = None
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.viewport().setAcceptDrops(True)

    def mousePressEvent(self, event):  # type: ignore[override]
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):  # type: ignore[override]
        if not (event.buttons() & Qt.LeftButton):
            super().mouseMoveEvent(event)
            return
        if self._drag_start_pos is None:
            super().mouseMoveEvent(event)
            return
        if (event.position().toPoint() - self._drag_start_pos).manhattanLength() < QApplication.startDragDistance():
            super().mouseMoveEvent(event)
            return
        item = self.itemAt(self._drag_start_pos)
        payload = dict(item.data(Qt.UserRole) or {}) if item is not None else {}
        if not payload or str(payload.get("drag_type", "") or "").strip() != "planned_block":
            super().mouseMoveEvent(event)
            return
        drag = QDrag(self)
        mime = QMimeData()
        try:
            raw = json.dumps(payload)
        except Exception:
            raw = "{}"
        mime.setData("application/x-lugest-planning-item", raw.encode("utf-8"))
        drag.setMimeData(mime)
        drag.exec(Qt.MoveAction)
        self._drag_start_pos = None

    def mouseReleaseEvent(self, event):  # type: ignore[override]
        self._drag_start_pos = None
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):  # type: ignore[override]
        if event.button() == Qt.LeftButton:
            item = self.itemAt(event.position().toPoint())
            payload = dict(item.data(Qt.UserRole) or {}) if item is not None else {}
            if payload and str(payload.get("drag_type", "") or "").strip() == "planned_block":
                handled = bool(getattr(self.page_ref, "_open_block_flow_pdf_from_payload", lambda *_args, **_kwargs: False)(payload))
                if handled:
                    event.accept()
                    return
        super().mouseDoubleClickEvent(event)

    def dragEnterEvent(self, event):  # type: ignore[override]
        if event.mimeData().hasFormat("application/x-lugest-planning-item"):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):  # type: ignore[override]
        if event.mimeData().hasFormat("application/x-lugest-planning-item"):
            pos = event.position().toPoint()
            idx = self.indexAt(pos)
            if idx.isValid() and idx.column() > 0:
                event.acceptProposedAction()
                return
        super().dragMoveEvent(event)

    def dropEvent(self, event):  # type: ignore[override]
        if event.mimeData().hasFormat("application/x-lugest-planning-item"):
            pos = event.position().toPoint()
            idx = self.indexAt(pos)
            if not idx.isValid() or idx.column() <= 0:
                event.ignore()
                return
            try:
                payload = json.loads(bytes(event.mimeData().data("application/x-lugest-planning-item")).decode("utf-8"))
            except Exception:
                payload = {}
            if str(payload.get("drag_type", "") or "").strip() == "planned_block":
                handled = bool(getattr(self.page_ref, "_drop_planned_block_payload", lambda *_args, **_kwargs: False)(payload, idx.row(), idx.column()))
            else:
                handled = bool(getattr(self.page_ref, "_drop_backlog_payload", lambda *_args, **_kwargs: False)(payload, idx.row(), idx.column()))
            if handled:
                event.acceptProposedAction()
                return
        super().dropEvent(event)


class PlanningPage(QWidget):
    page_title = "Planeamento"
    page_subtitle = "Semana operacional em grelha, com blocos coloridos e backlog real."
    uses_backend_reload = True

    @staticmethod
    def _operation_display_name(operation: str) -> str:
        mapping = {
            "Corte Laser": "Laser",
            "Maquinacao": "Maquinação",
            "Expedicao": "Expedição",
        }
        raw = str(operation or "").strip()
        return mapping.get(raw, raw)

    @staticmethod
    def _operation_accent_color(operation: str) -> str:
        mapping = {
            "Corte Laser": "#2563eb",
            "Quinagem": "#d97706",
            "Serralharia": "#0f766e",
            "Maquinacao": "#7c3aed",
            "Roscagem": "#be123c",
            "Lacagem": "#0f766e",
            "Montagem": "#059669",
            "Embalamento": "#0891b2",
            "Expedicao": "#1d4ed8",
            "Furo Manual": "#475569",
        }
        return mapping.get(str(operation or "").strip(), "#1d4ed8")

    def _operation_button_stylesheet(self, operation: str, *, selected: bool = False) -> str:
        accent = self._operation_accent_color(operation)
        if selected:
            return (
                "QPushButton {"
                f"background:{accent};"
                "color:#ffffff; border:1px solid rgba(15,23,42,0.28);"
                "border-radius:22px; padding:18px 24px; font-size:18px; font-weight:900;}"
                f"QPushButton:hover {{background:{accent}; border-color:#0f172a;}}"
            )
        return (
            "QPushButton {"
            "background:#ffffff;"
            f"border:1px solid {accent};"
            "color:#0f172a; border-radius:22px; padding:18px 24px; font-size:18px; font-weight:850;}"
            f"QPushButton:hover {{background:#f7fafc; border-color:{accent};}}"
        )

    def _planning_operation_options(self) -> list[str]:
        if self.backend is not None and hasattr(self.backend, "planning_operation_options"):
            options = list(self.backend.planning_operation_options() or [])
        else:
            options = []
        if not options:
            options = ["Corte Laser", "Quinagem", "Serralharia", "Maquinacao", "Roscagem", "Lacagem", "Montagem", "Embalamento", "Expedicao", "Furo Manual"]
        return [str(op or "").strip() for op in options if str(op or "").strip()]

    def _operation_grid_position(self, operation: str, fallback_index: int) -> tuple[int, int]:
        layout_order = [
            ["Corte Laser", "Furo Manual"],
            ["Quinagem", "Montagem"],
            ["Roscagem", "Embalamento"],
            ["Maquinacao", "Expedicao"],
            ["Serralharia", "Lacagem"],
        ]
        for row, names in enumerate(layout_order):
            for column, name in enumerate(names):
                if str(operation or "").strip() == name:
                    return row, column
        return 5 + (fallback_index // 2), fallback_index % 2

    def _rebuild_operation_buttons(self) -> None:
        operation_options = self._planning_operation_options()
        if operation_options == list(getattr(self, "_operation_button_options", []) or []):
            return
        _clear_layout_widgets(self.operation_grid)
        self.operation_buttons = {}
        used_positions: set[tuple[int, int]] = set()
        overflow_index = 0
        for op_name in operation_options:
            button = QPushButton(self._operation_display_name(str(op_name)))
            button.setCheckable(True)
            button.setMinimumHeight(72)
            button.setMinimumWidth(340)
            button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            button.setProperty("variant", "secondary")
            button.setStyleSheet(self._operation_button_stylesheet(str(op_name)))
            button.clicked.connect(lambda checked=False, value=op_name: self._set_operation(value))
            self.operation_buttons[str(op_name)] = button
            row, column = self._operation_grid_position(str(op_name), overflow_index)
            while (row, column) in used_positions:
                overflow_index += 1
                row, column = overflow_index, 2
            used_positions.add((row, column))
            self.operation_grid.addWidget(button, row, column)
        for column in range(2):
            self.operation_grid.setColumnStretch(column, 1)
        self._operation_button_options = list(operation_options)

    def __init__(self, runtime_service, backend=None, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        self.backend = backend
        self.week_start = date.today() - timedelta(days=date.today().weekday())
        self.current_active: list[dict] = []
        self.current_history: list[dict] = []
        self.current_pending: list[dict] = []
        self.current_week_dates: list[str] = []
        self.current_time_slots: list[str] = []
        self.current_blocked_windows: list[dict] = []
        self.current_operation = ""
        self.current_resource = ""
        self.operation_selected = False
        self.selected_block_id = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(6)

        top_card = CardFrame()
        self.top_card = top_card
        top_card.set_tone("info")
        top_card.setMinimumHeight(44)
        top_card.setMaximumHeight(48)
        top_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 4, 14, 4)
        top_layout.setSpacing(2)
        week_row = QHBoxLayout()
        week_row.setSpacing(6)
        self.week_label = QLabel("-")
        self.week_label.setStyleSheet("font-size: 15px; font-weight: 900; color: #0f172a;")
        self.period_meta = QLabel("Periodo -")
        self.period_meta.setProperty("role", "muted")
        self.period_meta.setStyleSheet("font-size: 11px;")
        self.period_meta.setMinimumHeight(14)
        week_row.addWidget(self.week_label, 1)
        self.prev_week_btn = QPushButton("Semana -")
        self.prev_week_btn.setProperty("variant", "secondary")
        self.prev_week_btn.clicked.connect(self._prev_week)
        self.next_week_btn = QPushButton("Semana +")
        self.next_week_btn.setProperty("variant", "secondary")
        self.next_week_btn.clicked.connect(self._next_week)
        self.current_week_btn = QPushButton("Semana atual")
        self.current_week_btn.setProperty("variant", "secondary")
        self.current_week_btn.clicked.connect(self._current_week)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.auto_plan_btn = QPushButton("Auto operação")
        self.auto_plan_btn.clicked.connect(self._auto_plan)
        self.auto_flow_btn = QPushButton("Auto fluxo")
        self.auto_flow_btn.clicked.connect(self._auto_plan_full_flow)
        self.clear_week_btn = QPushButton("Limpar semana")
        self.clear_week_btn.setProperty("variant", "secondary")
        self.clear_week_btn.clicked.connect(self._clear_week)
        self.move_earlier_btn = QPushButton("-30m")
        self.move_earlier_btn.setProperty("variant", "secondary")
        self.move_earlier_btn.clicked.connect(lambda: self._move_selected_block(minutes_offset=-30))
        self.move_later_btn = QPushButton("+30m")
        self.move_later_btn.setProperty("variant", "secondary")
        self.move_later_btn.clicked.connect(lambda: self._move_selected_block(minutes_offset=30))
        self.move_prev_day_btn = QPushButton("Dia -")
        self.move_prev_day_btn.setProperty("variant", "secondary")
        self.move_prev_day_btn.clicked.connect(lambda: self._move_selected_block(day_offset=-1))
        self.move_next_day_btn = QPushButton("Dia +")
        self.move_next_day_btn.setProperty("variant", "secondary")
        self.move_next_day_btn.clicked.connect(lambda: self._move_selected_block(day_offset=1))
        self.remove_block_btn = QPushButton("Remover bloco")
        self.remove_block_btn.setProperty("variant", "danger")
        self.remove_block_btn.clicked.connect(self._remove_selected_block)
        self.view_blocks_btn = QPushButton("Blocos")
        self.view_blocks_btn.setProperty("variant", "secondary")
        self.view_blocks_btn.clicked.connect(self._show_active_blocks_dialog)
        self.view_history_btn = QPushButton("Histórico")
        self.view_history_btn.setProperty("variant", "secondary")
        self.view_history_btn.clicked.connect(self._show_history_dialog)
        self.laser_deadline_btn = QPushButton("Prazo final")
        self.laser_deadline_btn.setProperty("variant", "secondary")
        self.laser_deadline_btn.clicked.connect(self._show_laser_deadlines_dialog)
        self.deadline_email_btn = QPushButton("Enviar Prazo")
        self.deadline_email_btn.setProperty("variant", "warning")
        self.deadline_email_btn.setStyleSheet(
            "QPushButton {background: #f4c542; color: #0f172a; border: 1px solid #caa12b; "
            "border-radius: 10px; padding: 6px 10px; font-weight: 800; font-size: 11px;}"
            "QPushButton:hover {background: #ffd65a;}"
            "QPushButton:disabled {background: #f5e9b4; color: #7c6a2b;}"
        )
        self.deadline_email_btn.setMinimumWidth(104)
        self.deadline_email_btn.clicked.connect(self._send_selected_deadline_email)
        self.blocked_btn = QPushButton("Bloqueios")
        self.blocked_btn.setProperty("variant", "secondary")
        self.blocked_btn.clicked.connect(self._show_blocked_windows_dialog)
        self.pdf_btn = QPushButton("Plano PDF")
        self.pdf_btn.clicked.connect(self._open_pdf)
        for button in (
            self.prev_week_btn,
            self.next_week_btn,
            self.current_week_btn,
            self.refresh_btn,
            self.auto_plan_btn,
            self.auto_flow_btn,
            self.clear_week_btn,
            self.move_earlier_btn,
            self.move_later_btn,
            self.move_prev_day_btn,
            self.move_next_day_btn,
            self.remove_block_btn,
            self.view_blocks_btn,
            self.view_history_btn,
            self.laser_deadline_btn,
            self.deadline_email_btn,
            self.blocked_btn,
            self.pdf_btn,
        ):
            button.setProperty("compact", "true")
            button.setMinimumHeight(24)
            week_row.addWidget(button)
        top_layout.addLayout(week_row)
        self.status_meta = QLabel("Blocos 0 | Backlog 0")
        self.status_meta.setProperty("role", "muted")
        self.status_meta.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_meta.setStyleSheet("font-size: 11px;")
        self.status_meta.setMinimumHeight(14)
        meta_host = QWidget()
        meta_host.setMinimumHeight(14)
        meta_row = QHBoxLayout(meta_host)
        meta_row.setContentsMargins(0, 0, 0, 0)
        meta_row.setSpacing(10)
        meta_row.addWidget(self.period_meta, 1)
        meta_row.addWidget(self.status_meta, 0)
        top_layout.addWidget(meta_host)
        root.addWidget(top_card)

        navigation_card = CardFrame()
        self.navigation_card = navigation_card
        navigation_card.set_tone("default")
        navigation_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        navigation_layout = QVBoxLayout(navigation_card)
        navigation_layout.setContentsMargins(30, 24, 30, 26)
        navigation_layout.setSpacing(18)
        selector_header = QHBoxLayout()
        selector_header.setSpacing(8)
        selector_title = QLabel("Seleciona a operação")
        self.selector_title_label = selector_title
        selector_title.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f172a;")
        selector_subtitle = QLabel("Primeiro escolhes a área de trabalho. Depois, dentro da operação, selecionas a máquina/recurso para ver o quadro.")
        self.selector_subtitle_label = selector_subtitle
        selector_subtitle.setProperty("role", "muted")
        selector_subtitle.setWordWrap(True)
        selector_subtitle.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        selector_subtitle.setStyleSheet("font-size: 13px; color: #475569;")
        selector_title_host = QWidget()
        selector_title_layout = QVBoxLayout(selector_title_host)
        selector_title_layout.setContentsMargins(0, 0, 0, 0)
        selector_title_layout.setSpacing(0)
        selector_title_layout.addWidget(selector_title)
        selector_title_layout.addWidget(selector_subtitle)
        self.selector_title_host = selector_title_host
        self.selector_title_layout = selector_title_layout
        selector_header.addWidget(selector_title_host, 1)
        self.operation_summary_label = QLabel("-")
        self.operation_summary_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.operation_summary_label.setStyleSheet("font-size: 12px; font-weight: 700; color: #334155;")
        self.operation_summary_label.setWordWrap(False)
        selector_header.addWidget(self.operation_summary_label, 0)
        self.change_operation_btn = QPushButton("Trocar operação")
        self.change_operation_btn.setProperty("variant", "secondary")
        self.change_operation_btn.setProperty("compact", "true")
        self.change_operation_btn.setMinimumHeight(26)
        self.change_operation_btn.setMinimumWidth(128)
        self.change_operation_btn.clicked.connect(lambda: self._set_operation_entry_state(False))
        selector_header.addWidget(self.change_operation_btn, 0)
        self.selector_header = selector_header
        navigation_layout.addLayout(selector_header)

        self.operation_buttons: dict[str, QPushButton] = {}
        self._operation_button_options: list[str] = []
        operation_grid_host = QWidget()
        self.operation_grid_host = operation_grid_host
        operation_grid_host.setMinimumWidth(780)
        operation_grid_host.setMaximumWidth(1040)
        operation_grid_host.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        operation_grid = QGridLayout(operation_grid_host)
        self.operation_grid = operation_grid
        operation_grid.setContentsMargins(0, 0, 0, 0)
        operation_grid.setHorizontalSpacing(20)
        operation_grid.setVerticalSpacing(14)
        self._rebuild_operation_buttons()
        operation_wrap = QWidget()
        self.operation_wrap = operation_wrap
        operation_wrap.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        operation_wrap_layout = QHBoxLayout(operation_wrap)
        operation_wrap_layout.setContentsMargins(0, 8, 0, 8)
        operation_wrap_layout.setSpacing(0)
        operation_wrap_layout.addStretch(1)
        operation_wrap_layout.addWidget(operation_grid_host, 4, Qt.AlignHCenter | Qt.AlignVCenter)
        operation_wrap_layout.addStretch(1)
        navigation_layout.addWidget(operation_wrap, 1)

        resource_host = QWidget()
        self.resource_host = resource_host
        resource_layout = QHBoxLayout(resource_host)
        resource_layout.setContentsMargins(0, 0, 0, 0)
        resource_layout.setSpacing(6)
        self.resource_panel_title = QLabel("Máquina / recurso")
        self.resource_panel_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        self.resource_panel_hint = QLabel("Seleciona a máquina para analisar o planeamento da operação.")
        self.resource_panel_hint.setProperty("role", "muted")
        self.resource_panel_hint.setWordWrap(True)
        resource_text_host = QWidget()
        resource_text_layout = QVBoxLayout(resource_text_host)
        resource_text_layout.setContentsMargins(0, 0, 0, 0)
        resource_text_layout.setSpacing(0)
        resource_text_layout.addWidget(self.resource_panel_title)
        resource_text_layout.addWidget(self.resource_panel_hint)
        resource_layout.addWidget(resource_text_host, 0)
        resource_layout.addStretch(1)
        self.change_operation_inline_btn = QPushButton("Trocar operação")
        self.change_operation_inline_btn.setProperty("variant", "warning")
        self.change_operation_inline_btn.setMinimumHeight(28)
        self.change_operation_inline_btn.setMinimumWidth(138)
        self.change_operation_inline_btn.setStyleSheet(
            "QPushButton {background: #f4c542; color: #0f172a; border: 1px solid #caa12b; "
            "border-radius: 10px; padding: 6px 12px; font-weight: 800; font-size: 11px;}"
            "QPushButton:hover {background: #ffd65a;}"
            "QPushButton:disabled {background: #f5e9b4; color: #7c6a2b;}"
        )
        self.change_operation_inline_btn.clicked.connect(lambda: self._set_operation_entry_state(False))
        resource_layout.addWidget(self.change_operation_inline_btn, 0, Qt.AlignVCenter)
        self.resource_combo = QComboBox()
        self.resource_combo.setMinimumHeight(28)
        self.resource_combo.setMinimumWidth(220)
        self.resource_combo.currentTextChanged.connect(lambda _value: self._set_resource(self.resource_combo.currentText().strip()))
        resource_layout.addWidget(self.resource_combo, 0, Qt.AlignVCenter)
        navigation_layout.addWidget(resource_host)
        root.addWidget(navigation_card)

        cards_host = QWidget()
        self.cards_host = cards_host
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(10)
        cards_layout.setVerticalSpacing(6)
        self.cards = [StatCard(title) for title in ("Blocos ativos", "Encomendas", "Carga semanal", "Blocos fechados")]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, index)
            card.setMinimumHeight(66)
            card.setMaximumHeight(70)
            card.layout().setContentsMargins(12, 10, 12, 10)
            card.layout().setSpacing(3)
            card.title_label.setWordWrap(True)
            card.subtitle_label.setWordWrap(True)
            card.title_label.setStyleSheet("font-size: 9px;")
            card.value_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
            card.subtitle_label.setStyleSheet("font-size: 9px;")
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("warning")
        self.cards[2].set_tone("success")
        self.cards[3].set_tone("default")
        root.addWidget(cards_host)

        main_split = QSplitter(Qt.Horizontal)
        self.main_split = main_split
        main_split.setChildrenCollapsible(False)

        backlog_card = CardFrame()
        backlog_card.set_tone("warning")
        backlog_layout = QVBoxLayout(backlog_card)
        backlog_layout.setContentsMargins(12, 10, 12, 10)
        self.backlog_title = QLabel("Produção / Montagem")
        self.backlog_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.backlog_title.setWordWrap(True)
        self.backlog_table = PlanningBacklogTable(self)
        self.backlog_table.setColumnCount(6)
        self.backlog_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Material", "Esp.", "Tempo", "Obs."])
        self.backlog_table.verticalHeader().setVisible(False)
        self.backlog_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.backlog_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.backlog_table.verticalHeader().setDefaultSectionSize(18)
        self.backlog_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.backlog_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.backlog_table.setStyleSheet("font-size: 11px;")
        self.backlog_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.backlog_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.backlog_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.backlog_table.itemSelectionChanged.connect(self._sync_planning_actions)
        backlog_layout.addWidget(self.backlog_title)
        backlog_layout.addWidget(self.backlog_table)
        main_split.addWidget(backlog_card)

        grid_card = CardFrame()
        grid_card.set_tone("info")
        grid_layout = QVBoxLayout(grid_card)
        grid_layout.setContentsMargins(12, 10, 12, 10)
        self.grid_title = QLabel("Quadro semanal")
        self.grid_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.grid_title.setWordWrap(True)
        self.grid = PlanningGridTable(self)
        self.grid.setEditTriggers(QTableWidget.NoEditTriggers)
        self.grid.setSelectionMode(QTableWidget.NoSelection)
        self.grid.verticalHeader().setVisible(False)
        self.grid.verticalHeader().setDefaultSectionSize(18)
        self.grid.verticalHeader().setMinimumSectionSize(16)
        self.grid.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.grid.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.grid.setStyleSheet("font-size: 9px;")
        self.grid.setToolTip("Arrasta os blocos para reajustar. Duplo clique num bloco abre o PDF do fluxo da encomenda.")
        self.grid.itemClicked.connect(self._handle_grid_item_clicked)
        grid_layout.addWidget(self.grid_title)
        grid_layout.addWidget(self.grid)
        grid_card.setMaximumHeight(16777215)
        main_split.addWidget(grid_card)
        main_split.setHandleWidth(6)
        main_split.setSizes([350, 1190])
        root.addWidget(main_split, 1)

        self.active_card = CardFrame(self)
        self.active_card.set_tone("info")
        active_layout = QVBoxLayout(self.active_card)
        active_layout.setContentsMargins(16, 14, 16, 14)
        self.active_title = QLabel("Blocos da semana")
        self.active_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.active_table = QTableWidget(0, 8)
        self.active_table.setHorizontalHeaderLabels(["Dia", "Inicio", "Duracao", "Encomenda", "Material", "Esp.", "Recurso", "Chapa"])
        self.active_table.verticalHeader().setVisible(False)
        self.active_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.active_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.active_table.verticalHeader().setDefaultSectionSize(18)
        self.active_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.active_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.active_table.setStyleSheet("font-size: 11px;")
        self.active_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.active_table.setToolTip("Duplo clique abre o PDF do fluxo completo da encomenda.")
        self.active_table.itemSelectionChanged.connect(self._sync_planning_actions)
        self.active_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_block_flow_pdf())
        active_layout.addWidget(self.active_title)
        active_layout.addWidget(self.active_table)
        self.history_card = CardFrame(self)
        self.history_card.set_tone("default")
        history_layout = QVBoxLayout(self.history_card)
        history_layout.setContentsMargins(16, 14, 16, 14)
        self.history_title = QLabel("Histórico de planeamento")
        self.history_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.history_table = QTableWidget(0, 8)
        self.history_table.setHorizontalHeaderLabels(["Data", "Inicio", "Encomenda", "Material", "Esp.", "Recurso", "Planeado", "Real / Estado"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.verticalHeader().setDefaultSectionSize(18)
        self.history_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.history_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.history_table.setStyleSheet("font-size: 11px;")
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        history_layout.addWidget(self.history_title)
        history_layout.addWidget(self.history_table)
        self.active_card.hide()
        self.history_card.hide()
        self._set_operation_entry_state(False)
        self._apply_operation_labels()
        self._fit_planning_grid()
        self._sync_planning_actions()

    def refresh(self) -> None:
        self._rebuild_operation_buttons()
        if self.operation_selected and str(self.current_operation or "").strip() not in self.operation_buttons:
            self.current_operation = ""
            self.current_resource = ""
            self.operation_selected = False
        if not self.operation_selected or not str(self.current_operation or "").strip():
            self._set_operation_entry_state(False)
            self._apply_operation_labels()
            return
        self._set_operation_entry_state(True)
        self._refresh_resource_options()
        self._apply_operation_labels()
        if self.backend is not None and hasattr(self.backend, "planning_overview_data"):
            data = self.backend.planning_overview_data(
                week_start=self.week_start.isoformat(),
                operation=self.current_operation,
                resource=self.current_resource,
            )
        else:
            data = self.runtime_service.planning_overview(
                week_start=self.week_start.isoformat(),
                operation=self.current_operation,
            )
        summary = data.get("summary", {})
        self.current_week_dates = list(data.get("week_dates", []) or [])
        self.current_blocked_windows = list(self.backend.planning_blocked_windows() if self.backend is not None else [])
        resource_suffix = f" | {self.current_resource}" if str(self.current_resource or "").strip() else ""
        self.week_label.setText(f"Semana {summary.get('week_label', '-')} | {self._operation_display_name(self.current_operation)}{resource_suffix}")
        self._fill_grid(data)
        QTimer.singleShot(0, self._fit_planning_grid)
        self.current_active = list(data.get("active", []) or [])
        self.current_history = list(data.get("history", []) or [])
        _fill_table(
            self.active_table,
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    f"{r.get('duracao_min', 0):.1f}",
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
                    r.get("maquina", r.get("posto_trabalho", "-")),
                    r.get("chapa", "-"),
                ]
                for r in self.current_active
            ],
            align_center_from=1,
        )
        for row_index, row in enumerate(self.current_active):
            self._paint_row_with_block_color(self.active_table, row_index, str(row.get("color", "") or "#dbeafe"))
        self._restore_active_block_selection()
        if self.backend is not None:
            self.current_pending = list(self.backend.planning_pending_rows(operation=self.current_operation, resource=self.current_resource))
            _fill_table(
                self.backlog_table,
                [
                    [
                        r.get("numero", "-"),
                        r.get("cliente", "-"),
                        r.get("material", "-"),
                        r.get("espessura", "-"),
                        f"{r.get('tempo_min', 0):.1f}",
                        " | ".join([part for part in [str(r.get("recurso", "") or "").strip(), str(r.get("obs", "") or "").strip()] if part]) or "-",
                    ]
                    for r in self.current_pending
                ],
                align_center_from=4,
            )
            for row_index, row in enumerate(self.current_pending):
                _paint_table_row(self.backlog_table, row_index, "Concluida" if bool(row.get("laser_done")) else str(row.get("estado", "")))
        else:
            self.current_pending = list(data.get("backlog", []) or [])
            _fill_table(
                self.backlog_table,
                [
                    [
                        r.get("numero", "-"),
                        r.get("cliente", "-"),
                        r.get("estado", "-"),
                        r.get("data_entrega", "-"),
                        f"{r.get('tempo_plan_min', 0):.1f}",
                        r.get("obs", "-"),
                    ]
                    for r in self.current_pending
                ],
                align_center_from=4,
            )
            for row_index, row in enumerate(self.current_pending):
                _paint_table_row(self.backlog_table, row_index, str(row.get("estado", "")))
        for row_index, row in enumerate(self.current_pending):
            item = self.backlog_table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, dict(row))
        backlog_count = len(self.current_pending)
        self.period_meta.setText(f"Periodo {summary.get('week_start', '-')} a {summary.get('week_end', '-')}")
        self.status_meta.setText(f"Blocos visiveis {summary.get('blocos_ativos', 0)} | Encomendas por encaixar {backlog_count}")
        self.cards[0].set_data(summary.get("blocos_ativos", 0), f"Ativos total {summary.get('min_ativos_total', 0):.0f} min")
        self.cards[1].set_data(backlog_count, "Pendentes por planear")
        self.cards[2].set_data(f"{summary.get('min_ativos', 0):.0f} min", f"Carga {self._operation_display_name(self.current_operation)}{resource_suffix}")
        self.cards[3].set_data(summary.get("historico_mes", 0), f"{summary.get('min_historico_mes', 0):.0f} min fechados")
        _fill_table(
            self.history_table,
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
                    r.get("maquina", r.get("posto_trabalho", "-")),
                    f"{r.get('duracao_min', 0):.1f}",
                    f"{r.get('tempo_real_min', 0):.1f} / {r.get('estado_final', '-')}",
                ]
                for r in self.current_history[:120]
            ],
            align_center_from=1,
        )
        for row_index, row in enumerate(self.current_history[:120]):
            self._paint_history_row(row_index, row)
        self._sync_planning_actions()

    def _set_operation(self, operation: str) -> None:
        selected = str(operation or "").strip() or "Corte Laser"
        was_selected = bool(self.operation_selected)
        if was_selected and selected == self.current_operation:
            self._apply_operation_labels()
            return
        self.current_operation = selected
        self.current_resource = ""
        self.operation_selected = True
        self._set_operation_entry_state(True)
        self._refresh_resource_options()
        selected_resource = self.resource_combo.currentText().strip()
        self.current_resource = "" if selected_resource.lower() == "todos" else selected_resource
        self.refresh()

    def _set_operation_entry_state(self, selected: bool) -> None:
        self.top_card.setVisible(bool(selected))
        self.resource_host.setVisible(bool(selected))
        self.cards_host.setVisible(bool(selected))
        self.main_split.setVisible(bool(selected))
        self.operation_selected = bool(selected and str(self.current_operation or "").strip())
        self.operation_grid_host.setVisible(not self.operation_selected)
        self.operation_wrap.setVisible(not self.operation_selected)
        self.change_operation_btn.setVisible(self.operation_selected)
        self.change_operation_inline_btn.setVisible(self.operation_selected)
        self.operation_summary_label.setVisible(self.operation_selected)
        if self.operation_selected:
            self.navigation_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            self.navigation_card.setMinimumHeight(58)
            self.navigation_card.setMaximumHeight(62)
            nav_layout = self.navigation_card.layout()
            if nav_layout is not None:
                nav_layout.setContentsMargins(14, 3, 14, 4)
                nav_layout.setSpacing(1)
            self.selector_title_label.setText("Planeamento da operação")
            self.selector_subtitle_label.setText("Aqui escolhes a máquina/recurso da operação atual para consultar o quadro semanal.")
            self.selector_title_label.setStyleSheet("font-size: 12px; font-weight: 900; color: #0f172a;")
            self.selector_subtitle_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.selector_title_layout.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.selector_title_host.setVisible(True)
            self.selector_subtitle_label.setVisible(False)
            self.resource_panel_hint.setVisible(False)
            self.operation_summary_label.setVisible(False)
            self.change_operation_btn.setVisible(False)
            self.change_operation_inline_btn.setVisible(True)
            self.resource_panel_title.setStyleSheet("font-size: 12px; font-weight: 800; color: #0f172a;")
        else:
            self._rebuild_operation_buttons()
            self.navigation_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
            self.navigation_card.setMinimumHeight(520)
            self.navigation_card.setMaximumHeight(16777215)
            nav_layout = self.navigation_card.layout()
            if nav_layout is not None:
                nav_layout.setContentsMargins(30, 18, 30, 22)
                nav_layout.setSpacing(8)
            self.selector_title_label.setText("Seleciona a operação")
            self.selector_subtitle_label.setText("Primeiro escolhes a área de trabalho. Depois, dentro da operação, selecionas a máquina/recurso para ver o quadro.")
            self.selector_title_label.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f172a;")
            self.selector_subtitle_label.setStyleSheet("font-size: 13px; color: #475569;")
            self.selector_subtitle_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.selector_title_layout.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.selector_title_host.setVisible(False)
            self.selector_subtitle_label.setVisible(False)
            self.resource_panel_hint.setVisible(True)
            self.operation_summary_label.setVisible(False)
            self.change_operation_btn.setVisible(False)
            self.change_operation_inline_btn.setVisible(False)
            self.resource_panel_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
            self.current_operation = ""
            self.current_resource = ""
            self._refresh_resource_options()
            self.resource_host.setVisible(False)
        if not selected:
            self.active_card.hide()
            self.history_card.hide()
            self.operation_wrap.setFocus()
        self._apply_operation_labels()

    def _set_resource(self, resource: str) -> None:
        selected = str(resource or "").strip()
        if selected.lower() == "todos":
            selected = ""
        if selected == self.current_resource:
            return
        self.current_resource = selected
        self.refresh()

    def _refresh_resource_options(self) -> None:
        if self.backend is None or not hasattr(self.backend, "workcenter_resource_options"):
            return
        if not self.operation_selected or not str(self.current_operation or "").strip():
            self.resource_combo.blockSignals(True)
            self.resource_combo.clear()
            self.resource_combo.blockSignals(False)
            self.resource_combo.setEnabled(False)
            return
        current_value = str(self.current_resource or self.resource_combo.currentText() or "").strip()
        options = [
            str(value or "").strip()
            for value in list(self.backend.workcenter_resource_options(self.current_operation, include_all=True) or [])
            if str(value or "").strip()
        ]
        self.resource_combo.blockSignals(True)
        self.resource_combo.clear()
        for value in options:
            self.resource_combo.addItem(value)
        if not current_value and any(str(value).strip().lower() == "todos" for value in options):
            self.resource_combo.setCurrentText("Todos")
            self.current_resource = ""
        elif current_value and any(current_value.lower() == value.lower() for value in options):
            self.resource_combo.setCurrentText(current_value)
            self.current_resource = current_value
        elif options:
            self.resource_combo.setCurrentIndex(0)
            selected = self.resource_combo.currentText().strip()
            self.current_resource = "" if selected.lower() == "todos" else selected
        else:
            self.current_resource = ""
        self.resource_combo.setEnabled(bool(options))
        self.resource_combo.blockSignals(False)
        current_label = self._operation_display_name(str(self.current_operation or "Corte Laser").strip() or "Corte Laser")
        selected_resource = str(self.current_resource or "").strip()
        selected_label = selected_resource or "Todos os recursos"
        self.resource_panel_title.setText(f"Máquina / recurso - {current_label}")
        self.resource_panel_hint.setText(
            f"Dentro de {current_label}, seleciona a máquina/recurso que queres analisar. Atual: {selected_label}."
        )

    def _apply_operation_labels(self) -> None:
        current = str(self.current_operation or "").strip()
        resource_suffix = f" | {self.current_resource}" if str(self.current_resource or "").strip() else ""
        current_label = self._operation_display_name(current) if current else "Operação"
        for op_name, button in self.operation_buttons.items():
            is_selected = bool(self.operation_selected and op_name == current)
            button.setChecked(is_selected)
            button.setProperty("variant", "primary" if is_selected else "secondary")
            button.setStyleSheet(self._operation_button_stylesheet(op_name, selected=is_selected))
            button.style().unpolish(button)
            button.style().polish(button)
        if self.operation_selected and current:
            self.operation_summary_label.setText(f"Operação atual: {current_label}")
        else:
            self.operation_summary_label.setText("Escolhe a operação para abrir o planeamento")
        self.backlog_title.setText(f"Pendentes - {current_label}{resource_suffix}")
        self.grid_title.setText(f"Quadro semanal - {current_label}{resource_suffix}")
        self.active_title.setText(f"Blocos - {current_label}{resource_suffix}")
        self.history_title.setText(f"Historico - {current_label}{resource_suffix}")

    def _grid_metrics(self) -> tuple[int, int, int]:
        if self.backend is not None:
            metrics_fn = getattr(self.backend, "_planning_grid_metrics", None)
            if callable(metrics_fn):
                try:
                    start_min, end_min, slot = metrics_fn()
                    start_min = int(start_min)
                    end_min = int(end_min)
                    slot = max(1, int(slot))
                    if end_min > start_min:
                        return start_min, end_min, slot
                except Exception:
                    pass
        return 480, 1080, 30

    def _grid_time_slots(self) -> list[str]:
        start_min, end_min, slot = self._grid_metrics()
        slots: list[str] = []
        for minute in range(start_min, end_min, slot):
            hour = minute // 60
            minute_part = minute % 60
            slots.append(f"{hour:02d}:{minute_part:02d}")
        return slots

    def _fill_grid(self, data: dict) -> None:
        week_dates = list(data.get("week_dates", []))
        active = list(data.get("active", []))
        self.grid.clearSpans()
        times = self._grid_time_slots()
        _start_min, _end_min, slot = self._grid_metrics()
        self.current_time_slots = list(times)
        labels = ["Hora"]
        for date_txt in week_dates:
            try:
                labels.append(datetime.strptime(date_txt, "%Y-%m-%d").strftime("%a %d/%m"))
            except Exception:
                labels.append(date_txt)
        self.grid.setColumnCount(max(2, len(labels)))
        self.grid.setHorizontalHeaderLabels(labels)
        self.grid.setRowCount(len(times))
        self._fit_planning_grid()
        for row in range(len(times)):
            time_item = QTableWidgetItem(times[row])
            time_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            time_item.setBackground(QBrush(QColor("#eef2f7")))
            time_item.setForeground(QBrush(QColor("#0f172a")))
            self.grid.setItem(row, 0, time_item)
        for row in range(len(times)):
            for col in range(1, len(labels)):
                cell = QTableWidgetItem("")
                cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.grid.setItem(row, col, cell)
        for block in self.current_blocked_windows:
            start_min = int(block.get("start_min", 0) or 0)
            end_min = int(block.get("end_min", 0) or 0)
            label = str(block.get("label", "") or "Bloqueio").strip()
            weekdays = set(int(v) for v in list(block.get("weekdays", []) or []))
            for col, day_txt in enumerate(week_dates, start=1):
                try:
                    weekday = datetime.fromisoformat(day_txt).date().weekday()
                except Exception:
                    weekday = -1
                if weekdays and weekday not in weekdays:
                    continue
                for row, slot_txt in enumerate(times):
                    slot_min = datetime.strptime(slot_txt, "%H:%M").hour * 60 + datetime.strptime(slot_txt, "%H:%M").minute
                    slot_end = slot_min + slot
                    if slot_end <= start_min or slot_min >= end_min:
                        continue
                    block_item = QTableWidgetItem(label if slot_min == start_min else "")
                    block_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    block_item.setBackground(QBrush(QColor("#334155")))
                    block_item.setForeground(QBrush(QColor("#f8fafc")))
                    block_item.setToolTip(f"{label}: {block.get('start', '')} - {block.get('end', '')}")
                    self.grid.setItem(row, col, block_item)
        for block in active:
            day = str(block.get("data", ""))
            start = str(block.get("inicio", ""))
            duration = float(block.get("duracao_min", 0) or 0)
            if day not in week_dates or start not in times:
                continue
            row = times.index(start)
            col = week_dates.index(day) + 1
            span = max(1, int(round(duration / float(slot or 30))))
            span = min(span, len(times) - row)
            color_hex = str(block.get("color") or block.get("source_color") or "#cbd5e1")
            material_txt = _elide_middle(str(block.get("material", "-") or "-"), 18)
            resource_txt = str(block.get("maquina", block.get("posto_trabalho", "")) or "").strip()
            text = f"{block.get('encomenda', '-')}\n{material_txt} | {block.get('espessura', '-')}mm"
            if resource_txt:
                text += f"\n{resource_txt}"
            item = QTableWidgetItem(text)
            item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            item.setBackground(QBrush(QColor(color_hex)))
            item.setForeground(QBrush(QColor("#ffffff" if _is_dark(color_hex) else "#0f172a")))
            item.setData(
                Qt.UserRole,
                {
                    "drag_type": "planned_block",
                    "block_id": str(block.get("id", "") or "").strip(),
                    "encomenda": str(block.get("encomenda", "") or "").strip(),
                    "material": str(block.get("material", "") or "").strip(),
                    "espessura": str(block.get("espessura", "") or "").strip(),
                    "data": day,
                    "inicio": start,
                },
            )
            item.setToolTip(
                f"Encomenda: {block.get('encomenda', '-')}\n"
                f"Material: {block.get('material', '-')}\n"
                f"Espessura: {block.get('espessura', '-')} mm\n"
                f"Recurso: {resource_txt or '-'}\n"
                f"Duracao: {duration:.0f} min"
            )
            self.grid.setItem(row, col, item)
            if span > 1:
                self.grid.setSpan(row, col, span, 1)

    def _fit_planning_grid(self) -> None:
        if not hasattr(self, "grid") or self.grid is None:
            return
        header = self.grid.horizontalHeader()
        header.setMinimumHeight(26)
        header.setDefaultAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        if self.grid.columnCount() <= 0:
            return
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.grid.setColumnWidth(0, 64)
        for col in range(1, self.grid.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Stretch)

    def _fmt_day(self, raw: str) -> str:
        try:
            return datetime.strptime(str(raw or ""), "%Y-%m-%d").strftime("%a %d/%m")
        except Exception:
            return str(raw or "-")

    def _paint_row_with_block_color(self, table: QTableWidget, row_index: int, color_hex: str) -> None:
        color = str(color_hex or "#dbeafe").strip() or "#dbeafe"
        bg = QBrush(QColor(color))
        fg = QBrush(QColor("#ffffff" if _is_dark(color) else "#0f172a"))
        for col_index in range(table.columnCount()):
            item = table.item(row_index, col_index)
            if item is None:
                continue
            item.setBackground(bg)
            item.setForeground(fg)

    def _paint_history_row(self, row_index: int, row: dict) -> None:
        status = str(row.get("estado_final", "") or "").strip()
        if status and status != "-":
            _paint_table_row(self.history_table, row_index, status)
            return
        real = float(row.get("tempo_real_min", 0) or 0)
        plan = float(row.get("duracao_min", 0) or 0)
        tone = "Concluida" if real <= plan else "Em pausa"
        _paint_table_row(self.history_table, row_index, tone)

    def _prev_week(self) -> None:
        self.week_start = self.week_start - timedelta(days=7)
        self.refresh()

    def _next_week(self) -> None:
        self.week_start = self.week_start + timedelta(days=7)
        self.refresh()

    def _current_week(self) -> None:
        self.week_start = date.today() - timedelta(days=date.today().weekday())
        self.refresh()

    def _selected_active_row(self) -> dict:
        current = self.active_table.currentItem()
        if current is None or current.row() >= len(self.current_active):
            block_id = str(self.selected_block_id or "").strip()
            if not block_id:
                return {}
            return next((row for row in self.current_active if str(row.get("id", "") or "").strip() == block_id), {})
        return self.current_active[current.row()]

    def _restore_active_block_selection(self) -> None:
        block_id = str(self.selected_block_id or "").strip()
        if not block_id:
            return
        self._select_active_block_by_id(block_id)

    def _select_active_block_by_id(self, block_id: str) -> bool:
        block_txt = str(block_id or "").strip()
        if not block_txt:
            return False
        for row_index, row in enumerate(self.current_active):
            if str(row.get("id", "") or "").strip() != block_txt:
                continue
            self.selected_block_id = block_txt
            self.active_table.selectRow(row_index)
            item = self.active_table.item(row_index, 0)
            if item is not None:
                self.active_table.setCurrentItem(item)
                self.active_table.scrollToItem(item)
            self._sync_planning_actions()
            return True
        return False

    def _handle_grid_item_clicked(self, item: QTableWidgetItem | None) -> None:
        payload = dict(item.data(Qt.UserRole) or {}) if item is not None else {}
        if str(payload.get("drag_type", "") or "").strip() != "planned_block":
            return
        block_id = str(payload.get("block_id", "") or "").strip()
        if not block_id:
            return
        self.selected_block_id = block_id
        self._select_active_block_by_id(block_id)

    def _selected_backlog_row(self) -> dict:
        current = self.backlog_table.currentItem()
        if current is None:
            return {}
        item = self.backlog_table.item(current.row(), 0)
        if item is None:
            return {}
        return dict(item.data(Qt.UserRole) or {})

    def _selected_planning_order_number(self) -> str:
        active = self._selected_active_row()
        if active:
            return str(active.get("encomenda", "") or "").strip()
        backlog = self._selected_backlog_row()
        return str(backlog.get("numero", "") or "").strip()

    def _planning_deadline_email_context(self, numero: str) -> dict[str, object]:
        if self.backend is None:
            raise ValueError("Backend de planeamento indisponível.")
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            raise ValueError("Seleciona uma encomenda no planeamento.")
        data = dict(self.backend.ensure_data() or {})
        enc = next(
            (
                row
                for row in list(data.get("encomendas", []) or [])
                if isinstance(row, dict) and str(row.get("numero", "") or "").strip() == numero_txt
            ),
            None,
        )
        if not isinstance(enc, dict):
            raise ValueError("Encomenda não encontrada.")
        client_code = str(enc.get("cliente", "") or "").strip()
        client = next(
            (
                row
                for row in list(data.get("clientes", []) or [])
                if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == client_code
            ),
            {},
        )
        deadline_row = next(
            (
                row
                for row in list(self.backend.planning_laser_deadline_rows() or [])
                if str(row.get("numero", "") or "").strip() == numero_txt
            ),
            {},
        )
        estimated_dt = deadline_row.get("fim_dt")
        if estimated_dt is None:
            raise ValueError("A encomenda ainda não tem prazo estimado calculado no planeamento.")
        flow_rows_getter = getattr(self.backend, "_planning_order_flow_rows", None)
        flow_rows = list(flow_rows_getter(numero_txt) or []) if callable(flow_rows_getter) else []
        resources = []
        for row in flow_rows:
            resource_txt = str(row.get("recurso", "") or "").strip()
            if resource_txt and resource_txt.lower() not in [value.lower() for value in resources]:
                resources.append(resource_txt)
        branding = dict(self.backend.branding_settings() or {})
        company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
        return {
            "numero": numero_txt,
            "nota_cliente": str(enc.get("nota_cliente", "") or "").strip(),
            "client_code": client_code,
            "client_name": str(client.get("nome", "") or "").strip(),
            "client_email": str(client.get("email", "") or "").strip(),
            "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
            "estimated_dt": estimated_dt,
            "estimated_date": estimated_dt.strftime("%Y-%m-%d"),
            "estimated_moment": estimated_dt.strftime("%d/%m/%Y %H:%M"),
            "grupos_txt": str(deadline_row.get("grupos_txt", "") or "-").strip() or "-",
            "planeado_txt": str(deadline_row.get("planeado_txt", "") or "-").strip() or "-",
            "estado": str(deadline_row.get("estado", "") or "-").strip() or "-",
            "last_item_txt": str(deadline_row.get("ultimo_item_txt", "") or "-").strip() or "-",
            "materiais_txt": str(deadline_row.get("materiais_txt", "") or "-").strip() or "-",
            "resources_txt": " | ".join(resources) if resources else "-",
            "observacoes": str(enc.get("Observacoes", "") or enc.get("Observações", "") or "").strip(),
            "company": company,
        }

    def _planning_deadline_email_subject(self, ctx: dict[str, object]) -> str:
        ref_txt = str(ctx.get("nota_cliente", "") or "").strip() or str(ctx.get("numero", "") or "").strip()
        return f"Prazo entrega {ref_txt}".strip()

    def _planning_deadline_email_body(self, ctx: dict[str, object]) -> str:
        lines = [
            "Boa tarde / Bonsoir / Good afternoon,",
            "",
            "Serve o presente para informar o prazo final previsto da encomenda, calculado pela última espessura/peça a concluir no planeamento.",
            "Par la présente nous vous informons du délai final prévu de la commande, calculé selon le dernier article du flux.",
            "I send this email to inform you of the final estimated due date for the order, based on the last item to finish in planning.",
            "",
            f"Encomenda / Commande / Order: {ctx.get('numero', '-')}",
            f"Num. Cliente / Numéro de Client / Client number: {ctx.get('client_code', '-')}",
            f"Nome cli. / Nom du Client / Client name: {ctx.get('client_name', '-')}",
            f"Prazo / Délais / Deadline: {ctx.get('estimated_date', '-')}",
            f"Enc. Cliente / Commande du Client / Client order: {ctx.get('nota_cliente', '-') or '-'}",
            f"Planeamento / Planification / Planning: {ctx.get('planeado_txt', '-')}",
            f"Estado / Statut / Status: {ctx.get('estado', '-')}",
            f"Último item do fluxo / Dernier article / Final item: {ctx.get('last_item_txt', '-')}",
            f"Máquinas / Machines / Resources: {ctx.get('resources_txt', '-')}",
            f"Materiais / Matières / Materials: {ctx.get('materiais_txt', '-')}",
        ]
        if str(ctx.get("observacoes", "") or "").strip():
            lines.append(f"Obs: {ctx.get('observacoes', '')}")
        lines.extend(
            [
                "",
                "Cumprimentos,",
                str(ctx.get("company", "") or "luGEST"),
            ]
        )
        return "\n".join(lines)

    def _planning_deadline_email_html_body(self, ctx: dict[str, object], *, logo_cid: str = "") -> str:
        company = str(ctx.get("company", "") or "luGEST")
        logo_html = (
            f'<img src="cid:{html.escape(logo_cid)}" alt="{html.escape(company)}" style="height:23px; display:block;">'
            if logo_cid
            else f'<div style="font-size:22px; font-weight:900; color:#0f172a;">{html.escape(company)}</div>'
        )
        rows = [
            ("Encomenda / Commande / Order", str(ctx.get("numero", "") or "-")),
            ("Cliente / Client", str(ctx.get("client_name", "") or "-")),
            ("Num. Cliente / Client number", str(ctx.get("client_code", "") or "-")),
            ("Prazo estimado / Deadline", str(ctx.get("estimated_date", "") or "-")),
            ("Enc. Cliente / Client order", str(ctx.get("nota_cliente", "") or "-") or "-"),
            ("Planeamento / Planning", str(ctx.get("planeado_txt", "") or "-")),
            ("Estado / Status", str(ctx.get("estado", "") or "-")),
            ("Último item / Final item", str(ctx.get("last_item_txt", "") or "-")),
            ("Máquinas / Resources", str(ctx.get("resources_txt", "") or "-")),
            ("Materiais / Materials", str(ctx.get("materiais_txt", "") or "-")),
        ]
        if str(ctx.get("observacoes", "") or "").strip():
            rows.append(("Observações / Notes", str(ctx.get("observacoes", "") or "").strip()))
        table_rows = "".join(
            "<tr>"
            f"<td style=\"padding:12px 16px; border-top:1px solid #e2e8f0; font-size:14px; color:#475569; width:42%;\">{html.escape(label)}</td>"
            f"<td style=\"padding:12px 16px; border-top:1px solid #e2e8f0; font-size:14px; color:#0f172a; font-weight:700;\">{html.escape(value)}</td>"
            "</tr>"
            for label, value in rows
        )
        return (
            "<html><body style=\"margin:0; padding:24px; background:#e2e8f0; font-family:'Segoe UI',Arial,sans-serif; color:#0f172a;\">"
            "<div style=\"max-width:900px; margin:0 auto; background:#ffffff; border-radius:24px; overflow:hidden; box-shadow:0 24px 50px rgba(15,23,42,0.12);\">"
            "<div style=\"padding:28px 34px; background:linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);\">"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;\">"
            "<tr>"
            f"<td>{logo_html}</td>"
            f"<td style=\"text-align:right; color:#e2e8f0; font-size:13px; font-weight:700;\">Prazo estimado de entrega<br>{html.escape(str(ctx.get('estimated_moment', '') or '-'))}</td>"
            "</tr>"
            "</table>"
            "</div>"
            "<div style=\"padding:34px 36px 28px 36px;\">"
            "<p style=\"margin:0 0 16px 0; font-size:22px; font-weight:800; color:#0f172a;\">Boa tarde / Bonsoir / Good afternoon,</p>"
            "<p style=\"margin:0 0 14px 0; font-size:15px; line-height:1.7; color:#334155;\">Serve o presente para informar o prazo final previsto da encomenda, calculado pela última espessura/peça a concluir no planeamento.</p>"
            "<p style=\"margin:0 0 14px 0; font-size:15px; line-height:1.7; color:#334155;\">Par la présente nous vous informons du délai final prévu de la commande, calculé selon le dernier article du flux.</p>"
            "<p style=\"margin:0 0 22px 0; font-size:15px; line-height:1.7; color:#334155;\">I send this email to inform you of the final estimated due date for the order, based on the last item to finish in planning.</p>"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse; border:1px solid #e2e8f0; border-radius:16px; overflow:hidden;\">"
            "<tr style=\"background:#1f2937; color:#ffffff; font-size:12px; font-weight:800; text-transform:uppercase;\">"
            "<td style=\"padding:12px 16px;\">Campo</td>"
            "<td style=\"padding:12px 16px;\">Valor</td>"
            "</tr>"
            f"{table_rows}"
            "</table>"
            "<p style=\"margin:24px 0 0 0; font-size:15px; line-height:1.7; color:#334155;\">Ficamos ao dispor para qualquer esclarecimento.</p>"
            "</div>"
            f"<div style=\"padding:16px 36px; background:#f8fafc; border-top:1px solid #e2e8f0; font-size:12px; color:#94a3b8; text-align:center;\">© {datetime.now().year} {html.escape(company)}</div>"
            "</div>"
            "</body></html>"
        )

    def _open_planning_deadline_email_draft(self, numero: str) -> None:
        if self.backend is None:
            return
        ctx = self._planning_deadline_email_context(numero)
        recipient = str(ctx.get("client_email", "") or "").strip()
        if not recipient:
            raise ValueError("O cliente desta encomenda não tem email definido.")

        safe_number = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in str(numero or "").strip())[:48] or "prazo"
        attachment_path: Path | None = Path(tempfile.gettempdir()) / f"Prazo_Entrega_{safe_number}.pdf"
        attachment_issue = ""
        try:
            self.backend.planning_render_order_detail_pdf(str(numero or "").strip(), output_path=attachment_path)
        except Exception as exc:
            attachment_issue = str(exc)
            attachment_path = None

        logo_path = getattr(self.backend, "logo_path", None)
        logo_file = Path(logo_path) if isinstance(logo_path, Path) and logo_path.exists() else None
        logo_cid = "lugest-planning-deadline-logo" if logo_file is not None else ""
        subject = self._planning_deadline_email_subject(ctx)
        body_plain = self._planning_deadline_email_body(ctx)
        body_html = self._planning_deadline_email_html_body(ctx, logo_cid=logo_cid)

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
                    QMessageBox.warning(self, "Planeamento", f"Não foi possível abrir o cliente de email:\n{exc}")
                    return
                fallback_message = "Outlook indisponível. Foi aberto o cliente de email por defeito."
                if attachment_issue:
                    fallback_message += f"\n\nTambém não foi possível gerar o PDF em anexo:\n{attachment_issue}"
                else:
                    fallback_message += "\n\nNota: o anexo PDF terá de ser adicionado manualmente neste modo."
                QMessageBox.information(self, "Planeamento", fallback_message)
                return
            if attachment_issue:
                QMessageBox.information(
                    self,
                    "Planeamento",
                    f"O email foi preparado no Outlook, mas o PDF não foi anexado automaticamente:\n{attachment_issue}",
                )

        _run_process_async(
            self,
            "powershell",
            ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
            environment=env,
            timeout_ms=30000,
            finished=on_email_ready,
        )

    def _send_selected_deadline_email(self) -> None:
        numero = self._selected_planning_order_number()
        if not numero:
            QMessageBox.warning(self, "Planeamento", "Seleciona primeiro uma encomenda ou um bloco do planeamento.")
            return
        try:
            self._open_planning_deadline_email_draft(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))

    def _open_block_flow_pdf_from_payload(self, payload: dict) -> bool:
        if self.backend is None:
            return False
        numero = str((payload or {}).get("encomenda", "") or "").strip()
        if not numero:
            return False
        try:
            self.backend.planning_open_order_detail_pdf(
                numero,
                focus_material=str((payload or {}).get("material", "") or "").strip(),
                focus_espessura=str((payload or {}).get("espessura", "") or "").strip(),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
        return True

    def _open_selected_block_flow_pdf(self) -> None:
        block = self._selected_active_row()
        if not block:
            QMessageBox.warning(self, "Planeamento", "Seleciona um bloco da semana.")
            return
        self._open_block_flow_pdf_from_payload(block)

    def _sync_planning_actions(self) -> None:
        has_backend = self.backend is not None
        if not self.operation_selected or not str(self.current_operation or "").strip():
            for button in (
                self.auto_plan_btn,
                self.auto_flow_btn,
                self.clear_week_btn,
                self.remove_block_btn,
                self.move_earlier_btn,
                self.move_later_btn,
                self.move_prev_day_btn,
                self.move_next_day_btn,
                self.view_blocks_btn,
                self.view_history_btn,
                self.blocked_btn,
                self.laser_deadline_btn,
                self.deadline_email_btn,
            ):
                button.setEnabled(False)
            self.laser_deadline_btn.setVisible(False)
            self.deadline_email_btn.setVisible(False)
            return
        active_row = self._selected_active_row()
        if active_row:
            self.selected_block_id = str(active_row.get("id", "") or "").strip()
        has_block = bool(active_row)
        has_order = bool(self._selected_planning_order_number())
        self.auto_plan_btn.setEnabled(has_backend and bool(self.current_pending))
        self.auto_flow_btn.setEnabled(has_backend and bool(self.current_pending))
        self.clear_week_btn.setEnabled(has_backend and bool(self.current_active))
        self.remove_block_btn.setEnabled(has_backend and has_block)
        self.move_earlier_btn.setEnabled(has_backend and has_block)
        self.move_later_btn.setEnabled(has_backend and has_block)
        self.move_prev_day_btn.setEnabled(has_backend and has_block)
        self.move_next_day_btn.setEnabled(has_backend and has_block)
        self.view_blocks_btn.setEnabled(bool(self.current_active))
        self.view_history_btn.setEnabled(bool(self.current_history))
        self.blocked_btn.setEnabled(has_backend)
        is_laser = self.current_operation == "Corte Laser"
        self.laser_deadline_btn.setEnabled(has_backend and is_laser)
        self.laser_deadline_btn.setVisible(is_laser)
        self.deadline_email_btn.setEnabled(has_backend and is_laser and has_order)
        self.deadline_email_btn.setVisible(is_laser)

    def _plan_order_dialog(self, rows: list[dict]) -> list[dict] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Ordem de planeamento - {self.current_operation}")
        dialog.resize(980, 600)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(12)
        info_card = CardFrame()
        info_card.set_tone("info")
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(14, 10, 14, 10)
        title = QLabel(f"Ordem de planeamento - {self.current_operation}")
        title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        info = QLabel("Seleciona os pendentes pela ordem desejada. A lista da direita pode ser arrastada para reordenar.")
        info.setProperty("role", "muted")
        info_layout.addWidget(title)
        info_layout.addWidget(info)
        layout.addWidget(info_card)
        body = QHBoxLayout()
        body.setSpacing(12)
        left_card = CardFrame()
        left_card.set_tone("default")
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_title = QLabel("Pendentes disponiveis")
        left_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        left = QListWidget()
        left.setAlternatingRowColors(True)
        left.setStyleSheet("QListWidget { border-radius: 10px; padding: 4px; } QListWidget::item { padding: 7px 8px; }")
        right_card = CardFrame()
        right_card.set_tone("success")
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_title = QLabel("Ordem planeada")
        right_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        right = QListWidget()
        right.setAlternatingRowColors(True)
        right.setStyleSheet("QListWidget { border-radius: 10px; padding: 4px; } QListWidget::item { padding: 7px 8px; }")
        right.setDragDropMode(QAbstractItemView.InternalMove)
        right.setDefaultDropAction(Qt.MoveAction)
        label_map: dict[str, dict] = {}
        for row in rows:
            resource_txt = str(row.get("recurso", "") or "").strip()
            label = f"{row.get('numero', '-')} | {row.get('material', '-')} | {row.get('espessura', '-')} | {row.get('tempo_min', 0):.0f} min"
            if resource_txt:
                label += f" | {resource_txt}"
            label_map[label] = row
            left.addItem(label)
        buttons = QVBoxLayout()
        buttons.setSpacing(8)
        add_btn = QPushButton("Adicionar ->")
        remove_btn = QPushButton("<- Remover")
        add_all_btn = QPushButton("Tudo ->")
        add_btn.clicked.connect(lambda: self._move_plan_items(left, right))
        remove_btn.clicked.connect(lambda: self._move_plan_items(right, left))
        add_all_btn.clicked.connect(lambda: self._move_all_plan_items(left, right))
        for button in (add_btn, remove_btn, add_all_btn):
            button.setMinimumWidth(118)
            button.setMinimumHeight(40)
            buttons.addWidget(button)
        buttons.addStretch(1)
        left_layout.addWidget(left_title)
        left_layout.addWidget(left, 1)
        right_layout.addWidget(right_title)
        right_layout.addWidget(right, 1)
        body.addWidget(left_card, 1)
        body.addLayout(buttons)
        body.addWidget(right_card, 1)
        layout.addLayout(body)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.accepted.connect(dialog.accept)
        box.rejected.connect(dialog.reject)
        layout.addWidget(box)
        if dialog.exec() != QDialog.Accepted:
            return None
        ordered = []
        for index in range(right.count()):
            label = right.item(index).text()
            row = label_map.get(label)
            if row is not None:
                ordered.append(dict(row))
        return ordered

    def _move_plan_items(self, source: QListWidget, target: QListWidget) -> None:
        for item in list(source.selectedItems()):
            target.addItem(item.text())
            source.takeItem(source.row(item))

    def _move_all_plan_items(self, source: QListWidget, target: QListWidget) -> None:
        while source.count():
            target.addItem(source.item(0).text())
            source.takeItem(0)

    def _drop_backlog_payload(self, payload: dict, row_index: int, col_index: int) -> bool:
        if self.backend is None:
            return False
        if col_index <= 0 or row_index < 0 or row_index >= len(self.current_time_slots):
            return False
        day_idx = col_index - 1
        if day_idx >= len(self.current_week_dates):
            return False
        try:
            self.backend.planning_place_block(
                str(payload.get("numero", "") or "").strip(),
                str(payload.get("material", "") or "").strip(),
                str(payload.get("espessura", "") or "").strip(),
                self.current_week_dates[day_idx],
                self.current_time_slots[row_index],
                operation=self.current_operation,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return False
        self.refresh()
        return True

    def _drop_planned_block_payload(self, payload: dict, row_index: int, col_index: int) -> bool:
        if self.backend is None:
            return False
        if col_index <= 0 or row_index < 0 or row_index >= len(self.current_time_slots):
            return False
        day_idx = col_index - 1
        if day_idx >= len(self.current_week_dates):
            return False
        block_id = str(payload.get("block_id", "") or "").strip()
        if not block_id:
            return False
        target_day = self.current_week_dates[day_idx]
        target_start = self.current_time_slots[row_index]
        current_day = str(payload.get("data", "") or "").strip()
        current_start = str(payload.get("inicio", "") or "").strip()
        if current_day == target_day and current_start == target_start:
            return True
        try:
            self.backend.planning_move_block_to(block_id, target_day, target_start)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return False
        self.refresh()
        return True

    def _show_table_dialog(self, title: str, headers: list[str], rows: list[list[str]], tones: list[str] | None = None) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(1100, 520)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(0,), contents=tuple(range(1, len(headers))))
        _fill_table(table, rows, align_center_from=1)
        for row_index, tone in enumerate(list(tones or [])):
            if tone:
                _paint_table_row(table, row_index, tone)
        layout.addWidget(table)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Guardar")
        buttons.button(QDialogButtonBox.Cancel).setText("Fechar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def _show_active_blocks_dialog(self) -> None:
        self._show_table_dialog(
            f"Blocos da semana - {self._operation_display_name(self.current_operation)}",
            ["Dia", "Inicio", "Duracao", "Encomenda", "Material", "Esp.", "Recurso", "Chapa"],
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    f"{r.get('duracao_min', 0):.1f}",
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
                    r.get("maquina", r.get("posto_trabalho", "-")),
                    r.get("chapa", "-"),
                ]
                for r in self.current_active
            ],
            [str(r.get("estado_final", "") or "") for r in self.current_active],
        )

    def _show_history_dialog(self) -> None:
        self._show_table_dialog(
            f"Histórico de planeamento - {self._operation_display_name(self.current_operation)}",
            ["Data", "Inicio", "Encomenda", "Material", "Esp.", "Recurso", "Planeado", "Real / Estado"],
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
                    r.get("maquina", r.get("posto_trabalho", "-")),
                    f"{r.get('duracao_min', 0):.1f}",
                    f"{r.get('tempo_real_min', 0):.1f} / {r.get('estado_final', '-')}",
                ]
                for r in self.current_history[:220]
            ],
            [str(r.get("estado_final", "") or "") for r in self.current_history[:220]],
        )

    def _show_laser_deadlines_dialog(self) -> None:
        if self.backend is None or self.current_operation != "Corte Laser":
            return
        rows = list(self.backend.planning_laser_deadline_rows())
        dialog = QDialog(self)
        dialog.setWindowTitle("Prazo Final")
        dialog.resize(1180, 560)
        layout = QVBoxLayout(dialog)
        info = QLabel(
            "Prazo previsto de conclusão global por encomenda, a contar com o corte laser e os postos seguintes já planeados. "
            "Se a encomenda estiver parcial, o fim apresentado ainda não representa o fluxo completo fechado."
        )
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Entrega", "Grupos", "Planeado", "Fim fluxo", "Estado"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(1,), contents=(0, 2, 3, 4, 5, 6))
        _fill_table(
            table,
            [
                [
                    row.get("numero", "-"),
                    row.get("cliente", "-"),
                    row.get("data_entrega", "-"),
                    row.get("grupos_txt", "-"),
                    row.get("planeado_txt", "-"),
                    row.get("fim_txt", "-"),
                    row.get("estado", "-"),
                ]
                for row in rows
            ],
            align_center_from=2,
        )
        tone_map = {
            "Fluxo concluído": "Concluida",
            "Planeado completo": "Concluida",
            "Planeado parcial": "Incompleta",
            "Por planear": "Pendente",
        }
        for idx, row in enumerate(rows):
            _paint_table_row(table, idx, tone_map.get(str(row.get("estado", "") or ""), "Pendente"))
            item = table.item(idx, 0)
            if item is not None:
                item.setToolTip(str(row.get("materiais_txt", "") or ""))
        layout.addWidget(table, 1)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        send_btn = QPushButton("Enviar Prazo")
        send_btn.setProperty("variant", "warning")
        send_btn.setStyleSheet(
            "QPushButton {background: #f4c542; color: #0f172a; border: 1px solid #caa12b; "
            "border-radius: 10px; padding: 8px 14px; font-weight: 800;}"
            "QPushButton:hover {background: #ffd65a;}"
            "QPushButton:disabled {background: #f5e9b4; color: #7c6a2b;}"
        )
        preview_btn = QPushButton("Abrir PDF")
        preview_btn.setProperty("variant", "secondary")
        save_btn = QPushButton("Guardar PDF")
        save_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.reject)
        btn_row.addWidget(send_btn)
        btn_row.addWidget(preview_btn)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        def selected_deadline_numero() -> str:
            current = table.currentItem()
            if current is None or current.row() >= len(rows):
                return ""
            return str(rows[current.row()].get("numero", "") or "").strip()

        def send_email() -> None:
            numero = selected_deadline_numero()
            if not numero:
                QMessageBox.warning(dialog, "Prazo Final", "Seleciona uma encomenda para enviar o prazo.")
                return
            try:
                self._open_planning_deadline_email_draft(numero)
            except Exception as exc:
                QMessageBox.critical(dialog, "Prazo Final", str(exc))

        def open_pdf() -> None:
            try:
                path = self.backend.planning_open_laser_deadlines_pdf()
            except Exception as exc:
                QMessageBox.critical(dialog, "Prazo Final", str(exc))
                return
            QMessageBox.information(dialog, "Prazo Final", f"PDF aberto:\n{path}")

        def save_pdf() -> None:
            path, _ = QFileDialog.getSaveFileName(dialog, "Guardar PDF", "prazos_fluxo_planeamento.pdf", "PDF (*.pdf)")
            if not path:
                return
            try:
                self.backend.planning_render_laser_deadlines_pdf(path)
            except Exception as exc:
                QMessageBox.critical(dialog, "Prazo Final", str(exc))
                return
            QMessageBox.information(dialog, "Prazo Final", f"PDF guardado em:\n{path}")

        send_btn.clicked.connect(send_email)
        preview_btn.clicked.connect(open_pdf)
        save_btn.clicked.connect(save_pdf)
        dialog.exec()

    def _blocked_window_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Bloqueio de planeamento")
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        label_edit = QLineEdit(str(initial.get("label", "") or "").strip())
        start_edit = QTimeEdit()
        start_edit.setDisplayFormat("HH:mm")
        start_edit.setTime(QTime.fromString(str(initial.get("start", "12:30") or "12:30"), "HH:mm"))
        end_edit = QTimeEdit()
        end_edit.setDisplayFormat("HH:mm")
        end_edit.setTime(QTime.fromString(str(initial.get("end", "14:00") or "14:00"), "HH:mm"))
        days_host = QWidget()
        days_layout = QHBoxLayout(days_host)
        days_layout.setContentsMargins(0, 0, 0, 0)
        days_layout.setSpacing(6)
        days_map = [("Seg", 0), ("Ter", 1), ("Qua", 2), ("Qui", 3), ("Sex", 4), ("Sab", 5)]
        day_boxes = []
        selected_days = set(int(v) for v in list(initial.get("weekdays", [0, 1, 2, 3, 4, 5])) or [])
        for label_txt, day_idx in days_map:
            box = QCheckBox(label_txt)
            box.setChecked(day_idx in selected_days)
            day_boxes.append((day_idx, box))
            days_layout.addWidget(box)
        form.addRow("Nome", label_edit)
        form.addRow("Inicio", start_edit)
        form.addRow("Fim", end_edit)
        form.addRow("Dias", days_host)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        weekdays = [day_idx for day_idx, box in day_boxes if box.isChecked()]
        if not weekdays:
            QMessageBox.warning(self, "Planeamento", "Seleciona pelo menos um dia.")
            return None
        return {
            "id": str(initial.get("id", "") or "").strip(),
            "label": label_edit.text().strip() or "Bloqueio",
            "start": start_edit.time().toString("HH:mm"),
            "end": end_edit.time().toString("HH:mm"),
            "start_min": start_edit.time().hour() * 60 + start_edit.time().minute(),
            "end_min": end_edit.time().hour() * 60 + end_edit.time().minute(),
            "weekdays": weekdays,
        }

    def _show_blocked_windows_dialog(self) -> None:
        if self.backend is None:
            return
        rows = list(self.backend.planning_blocked_windows())
        dialog = QDialog(self)
        dialog.setWindowTitle("Bloqueios de planeamento")
        dialog.resize(760, 420)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(0, 4)
        table.setHorizontalHeaderLabels(["Nome", "Dias", "Inicio", "Fim"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(0, 1), contents=(2, 3))
        def fill_rows() -> None:
            _fill_table(
                table,
                [[r.get("label", "-"), r.get("dias_txt", "-"), r.get("start", "-"), r.get("end", "-")] for r in rows],
                align_center_from=2,
            )
            for idx, row in enumerate(rows):
                item = table.item(idx, 0)
                if item is not None:
                    item.setData(Qt.UserRole, dict(row))
        fill_rows()
        layout.addWidget(table, 1)
        btn_row = QHBoxLayout()
        add_btn = QPushButton("Adicionar")
        edit_btn = QPushButton("Editar")
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "danger")
        btn_row.addWidget(add_btn)
        btn_row.addWidget(edit_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)

        def selected_row() -> tuple[int, dict]:
            idx = _selected_row_index(table)
            if idx < 0 or idx >= len(rows):
                return -1, {}
            return idx, rows[idx]

        def add_row() -> None:
            payload = self._blocked_window_dialog()
            if payload is None:
                return
            rows.append(payload)
            fill_rows()

        def edit_row() -> None:
            idx, row = selected_row()
            if idx < 0:
                QMessageBox.warning(dialog, "Planeamento", "Seleciona um bloqueio.")
                return
            payload = self._blocked_window_dialog(row)
            if payload is None:
                return
            rows[idx] = payload
            fill_rows()

        def remove_row() -> None:
            idx, row = selected_row()
            if idx < 0:
                QMessageBox.warning(dialog, "Planeamento", "Seleciona um bloqueio.")
                return
            if QMessageBox.question(dialog, "Planeamento", f"Remover bloqueio {row.get('label', '-') }?") != QMessageBox.Yes:
                return
            rows.pop(idx)
            fill_rows()

        add_btn.clicked.connect(add_row)
        edit_btn.clicked.connect(edit_row)
        remove_btn.clicked.connect(remove_row)
        if dialog.exec() != QDialog.Accepted:
            return
        self.backend.planning_set_blocked_windows(rows)
        self.refresh()

    def _auto_plan(self) -> None:
        if self.backend is None:
            QMessageBox.information(self, "Planeamento", "Auto planeamento indisponivel sem backend configurado.")
            return
        ordered = self._plan_order_dialog(self.current_pending)
        if ordered is None:
            return
        if not ordered:
            QMessageBox.warning(self, "Planeamento", "Seleciona pelo menos um item para planear.")
            return
        try:
            placed = self.backend.planning_auto_plan(ordered, self.week_start, operation=self.current_operation)
        except Exception as exc:
            QMessageBox.critical(self, "Auto planear", str(exc))
            return
        self.refresh()
        selected_keys = {
            (
                str(row.get("numero", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
            )
            for row in ordered
        }
        remaining = [
            row
            for row in self.current_pending
            if (
                str(row.get("numero", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
            )
            in selected_keys
        ]
        if not placed:
            if not remaining:
                QMessageBox.warning(
                    self,
                    "Planeamento",
                    "Os itens escolhidos já não estavam pendentes neste recurso. A lista foi atualizada para refletir o planeamento real.",
                )
            else:
                QMessageBox.warning(self, "Planeamento", "Não havia espaço livre na semana para encaixar novos blocos.")
            return
        if remaining:
            pending_txt = ", ".join(
                f"{row.get('numero', '-')} {float(row.get('tempo_min', 0) or 0):.0f} min"
                for row in remaining[:3]
            )
            if len(remaining) > 3:
                pending_txt += "..."
            QMessageBox.information(
                self,
                "Planeamento",
                f"{len(placed)} bloco(s) planeado(s). Ficou continuação pendente para: {pending_txt}",
            )
            return
        QMessageBox.information(self, "Planeamento", f"{len(placed)} bloco(s) planeado(s) na semana.")

    def _auto_plan_full_flow(self) -> None:
        if self.backend is None:
            QMessageBox.information(self, "Planeamento", "Auto planeamento indisponivel sem backend configurado.")
            return
        ordered = self._plan_order_dialog(self.current_pending)
        if ordered is None:
            return
        if not ordered:
            QMessageBox.warning(self, "Planeamento", "Seleciona pelo menos um item para planear.")
            return
        try:
            result = self.backend.planning_auto_plan_full_flow(ordered, self.week_start, operation=self.current_operation)
        except Exception as exc:
            QMessageBox.critical(self, "Auto fluxo", str(exc))
            return
        placed = list((result or {}).get("placed", []) or [])
        pending = list((result or {}).get("pending", []) or [])
        self.refresh()
        if not placed:
            still_pending = {
                (
                    str(row.get("numero", "") or "").strip(),
                    str(row.get("material", "") or "").strip(),
                    str(row.get("espessura", "") or "").strip(),
                )
                for row in self.current_pending
            }
            selected_pending = any(
                (
                    str(row.get("numero", "") or "").strip(),
                    str(row.get("material", "") or "").strip(),
                    str(row.get("espessura", "") or "").strip(),
                )
                in still_pending
                for row in ordered
            )
            if not selected_pending:
                QMessageBox.warning(
                    self,
                    "Planeamento",
                    "Os itens escolhidos já não estavam pendentes neste recurso. A lista foi atualizada para refletir o planeamento real.",
                )
            else:
                QMessageBox.warning(self, "Planeamento", "Não havia espaço livre na semana para encaixar novos blocos.")
            return
        if pending:
            pending_txt = ", ".join(
                f"{row.get('numero', '-')} {row.get('operacao', '-')} {float(row.get('remaining_min', 0) or 0):.0f} min"
                for row in pending[:3]
            )
            if len(pending) > 3:
                pending_txt += "..."
            QMessageBox.information(
                self,
                "Planeamento",
                f"{len(placed)} bloco(s) planeado(s) no fluxo. Ficou continuação pendente para: {pending_txt}",
            )
            return
        QMessageBox.information(self, "Planeamento", f"{len(placed)} bloco(s) planeado(s) no fluxo completo.")

    def _clear_week(self) -> None:
        if self.backend is None:
            return
        if QMessageBox.question(self, "Planeamento", "Remover todos os blocos visíveis desta semana?") != QMessageBox.Yes:
            return
        week_days = set((self.week_start + timedelta(days=i)).isoformat() for i in range(6))
        block_ids = [str(row.get("id", "") or "").strip() for row in self.current_active if str(row.get("data", "") or "") in week_days]
        try:
            for block_id in block_ids:
                self.backend.planning_remove_block(block_id)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return
        self.refresh()

    def _move_selected_block(self, *, day_offset: int = 0, minutes_offset: int = 0) -> None:
        if self.backend is None:
            return
        block = self._selected_active_row()
        if not block:
            QMessageBox.warning(self, "Planeamento", "Seleciona um bloco da semana.")
            return
        try:
            self.backend.planning_shift_block(str(block.get("id", "") or "").strip(), day_offset=day_offset, minutes_offset=minutes_offset)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return
        self.refresh()

    def _remove_selected_block(self) -> None:
        if self.backend is None:
            return
        block = self._selected_active_row()
        if not block:
            QMessageBox.warning(self, "Planeamento", "Seleciona um bloco da semana.")
            return
        if QMessageBox.question(self, "Planeamento", f"Remover bloco {block.get('encomenda', '-') } {block.get('material', '-') } {block.get('espessura', '-') }?") != QMessageBox.Yes:
            return
        try:
            self.backend.planning_remove_block(str(block.get("id", "") or "").strip())
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return
        self.selected_block_id = ""
        self.refresh()

    def _open_pdf(self) -> None:
        try:
            if self.backend is not None:
                path = self.backend.planning_open_pdf(self.week_start, operation=self.current_operation, resource=self.current_resource)
            else:
                path = self.runtime_service.planning_pdf(week_start=self.week_start.isoformat(), operation=self.current_operation)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return
        QMessageBox.information(self, "Planeamento", f"PDF aberto:\n{path}")


class LegacyPlanningPage(PlanningPage):
    page_subtitle = "Tabela semanal ao centro, encomendas pendentes à esquerda e bloqueios por semana."

    def __init__(self, runtime_service, backend=None, parent=None) -> None:
        super().__init__(runtime_service, backend, parent)
        root = self.layout()
        root.setSpacing(2)
        sections = _take_layout_items(root)
        top_item = sections[0] if len(sections) > 0 else None
        nav_item = sections[1] if len(sections) > 1 else None
        cards_item = sections[2] if len(sections) > 2 else None
        body_item = sections[3] if len(sections) > 3 else None

        top_widget = top_item.widget() if top_item and top_item.widget() is not None else None
        if top_widget is not None:
            top_widget.setMinimumHeight(44)
            top_widget.setMaximumHeight(48)
            top_layout = top_widget.layout()
            if top_layout is not None:
                top_layout.setContentsMargins(12, 4, 12, 4)
                top_layout.setSpacing(2)

        nav_widget = nav_item.widget() if nav_item and nav_item.widget() is not None else None
        if nav_widget is not None:
            nav_layout = nav_widget.layout()
            if bool(getattr(self, "operation_selected", False)):
                nav_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
                nav_widget.setMinimumHeight(58)
                nav_widget.setMaximumHeight(62)
            else:
                nav_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
                nav_widget.setMinimumHeight(520)
                nav_widget.setMaximumHeight(16777215)
            if nav_layout is not None:
                if bool(getattr(self, "operation_selected", False)):
                    nav_layout.setContentsMargins(12, 6, 12, 6)
                    nav_layout.setSpacing(4)
                else:
                    nav_layout.setContentsMargins(30, 18, 30, 22)
                    nav_layout.setSpacing(8)

        cards_widget = cards_item.widget() if cards_item and cards_item.widget() is not None else None
        if cards_widget is not None:
            cards_widget.setMaximumHeight(70)
            cards_layout = cards_widget.layout()
            if cards_layout is not None:
                if isinstance(cards_layout, QGridLayout):
                    cards_layout.setHorizontalSpacing(8)
                    cards_layout.setVerticalSpacing(4)
                else:
                    cards_layout.setSpacing(4)
        for card in getattr(self, "cards", []):
            card.setMinimumHeight(66)
            card.setMaximumHeight(70)
            card.layout().setContentsMargins(10, 8, 10, 8)
            card.layout().setSpacing(3)
            card.title_label.setWordWrap(True)
            card.subtitle_label.setWordWrap(True)
            card.title_label.setStyleSheet("font-size: 9px;")
            card.value_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
            card.subtitle_label.setStyleSheet("font-size: 9px;")

        self.grid.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid.verticalHeader().setDefaultSectionSize(16)
        self.grid.verticalHeader().setMinimumSectionSize(16)
        self.grid.horizontalHeader().setMinimumSectionSize(88)
        self.grid.horizontalHeader().setFixedHeight(26)
        self.grid.setStyleSheet("font-size: 8px;")
        self.grid.setMinimumHeight(_table_visible_height(self.grid, 20, extra=0))
        self.grid.setMaximumHeight(16777215)
        self.grid.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.backlog_table.verticalHeader().setDefaultSectionSize(17)
        self.backlog_table.setStyleSheet("font-size: 9.5px;")
        self.backlog_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        backlog_header = self.backlog_table.horizontalHeader()
        for col, width in ((0, 150), (1, 126), (2, 96), (3, 50), (4, 54)):
            backlog_header.setSectionResizeMode(col, QHeaderView.Interactive)
            backlog_header.resizeSection(col, width)
        backlog_header.setStretchLastSection(False)

        _adopt_layout_item(root, top_item)
        _adopt_layout_item(root, nav_item)
        _adopt_layout_item(root, cards_item)
        _adopt_layout_item(root, body_item, 1)
        body_widget = body_item.widget() if body_item and body_item.widget() is not None else None
        if body_widget is not None and isinstance(body_widget, QSplitter):
            body_widget.setHandleWidth(6)
            body_widget.setSizes([340, 1220])
        QTimer.singleShot(0, self._fit_planning_grid)
