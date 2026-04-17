from __future__ import annotations

import csv
import html
import json
import os
import re
import subprocess
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

from PySide6.QtCore import Qt, QDate, QMimeData, QTime, QTimer
from PySide6.QtGui import QColor, QBrush, QDrag
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDateEdit,
    QDoubleSpinBox,
    QFrame,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QInputDialog,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
    QStackedWidget,
    QSizePolicy,
)

from .laser_batch_quote_dialog import LaserBatchQuoteDialog
from .laser_nesting_dialog import LaserNestingDialog
from .laser_quote_dialogs import LaserQuoteDialog, LaserSettingsDialog
from ..widgets import CardFrame, StatCard

LIST_TABLE_FONT_PX = 15
LIST_TABLE_ROW_PX = 42


def _is_dark(hex_color: str) -> bool:
    color = QColor(str(hex_color or "#ffffff"))
    return (color.red() * 0.299 + color.green() * 0.587 + color.blue() * 0.114) < 160


def _fill_table(table: QTableWidget, rows: list[list[str]], align_center_from: int = 0) -> None:
    sorting_was_enabled = table.isSortingEnabled()
    table.setSortingEnabled(False)
    table.setRowCount(len(rows))
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            item = QTableWidgetItem(str(value))
            if col_index >= align_center_from:
                item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, col_index, item)
    table.setSortingEnabled(sorting_was_enabled)


def _state_visual(state: str) -> dict[str, str]:
    norm = str(state or "").strip().lower()
    if "rejeit" in norm:
        return {"bg": "#fff1f2", "fg": "#b42318", "border": "#f0c1bc", "tone": "danger"}
    if "avaria" in norm:
        return {"bg": "#fff1f2", "fg": "#b42318", "border": "#f0c1bc", "tone": "danger"}
    if "incomplet" in norm:
        return {"bg": "#fff9db", "fg": "#a16207", "border": "#f2d98b", "tone": "warning"}
    if "prepar" in norm or "edicao" in norm or "edição" in norm or "pendente" in norm:
        return {"bg": "#eef4ff", "fg": "#1d4ed8", "border": "#bfd2ea", "tone": "info"}
    if "produc" in norm or "curso" in norm:
        return {"bg": "#fff4e5", "fg": "#c2410c", "border": "#efcf98", "tone": "warning"}
    if "paus" in norm or "interromp" in norm:
        return {"bg": "#fff9db", "fg": "#a16207", "border": "#f2d98b", "tone": "warning"}
    if "concl" in norm:
        return {"bg": "#ecfdf3", "fg": "#166534", "border": "#b7dcc5", "tone": "success"}
    if "aprov" in norm or "convert" in norm or "entreg" in norm:
        return {"bg": "#f0fdf4", "fg": "#166534", "border": "#b7dcc5", "tone": "success"}
    if "enviado" in norm:
        return {"bg": "#f5f9ff", "fg": "#1d4ed8", "border": "#bfd2ea", "tone": "info"}
    return {"bg": "#f8fafc", "fg": "#334155", "border": "#c6d2e0", "tone": "default"}


def _state_palette(state: str) -> tuple[str, str]:
    visual = _state_visual(state)
    return visual["bg"], visual["fg"]


def _state_tone(state: str) -> str:
    return _state_visual(state)["tone"]


def _repolish(widget: QWidget) -> None:
    style = widget.style()
    if style is not None:
        style.unpolish(widget)
        style.polish(widget)
    widget.update()


def _set_panel_tone(frame: QWidget, tone: str = "default") -> None:
    frame.setProperty("tone", str(tone or "default"))
    _repolish(frame)


def _apply_state_chip(label: QLabel, state: str, text: str | None = None) -> None:
    visual = _state_visual(state)
    label.setProperty("role", "state_chip")
    label.setText(str(text if text is not None else state or "-") or "-")
    label.setStyleSheet(
        "padding: 6px 12px; border-radius: 999px; font-weight: 800;"
        f" background: {visual['bg']}; color: {visual['fg']}; border: 1px solid {visual['border']};"
    )


def _paint_table_row(table: QTableWidget, row_index: int, state: str) -> None:
    bg_hex, fg_hex = _state_palette(state)
    bg = QBrush(QColor(bg_hex))
    fg = QBrush(QColor(fg_hex))
    for col_index in range(table.columnCount()):
        item = table.item(row_index, col_index)
        if item is None:
            continue
        item.setBackground(bg)
        item.setForeground(fg)


def _configure_table(table: QTableWidget, *, stretch: tuple[int, ...] = (), contents: tuple[int, ...] = (), center_from: int | None = None) -> None:
    table.setAlternatingRowColors(True)
    table.setWordWrap(False)
    table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
    table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    header = table.horizontalHeader()
    header.setStretchLastSection(False)
    for col in range(table.columnCount()):
        if col in stretch:
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        elif col in contents:
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        else:
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, 130)
    if center_from is not None:
        for row in range(table.rowCount()):
            for col in range(center_from, table.columnCount()):
                item = table.item(row, col)
                if item is not None:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))


def _set_table_columns(table: QTableWidget, specs: list[tuple[int, str, int]]) -> None:
    header = table.horizontalHeader()
    for column, mode, width in list(specs or []):
        mode_txt = str(mode or "").strip().lower()
        if mode_txt == "stretch":
            header.setSectionResizeMode(column, QHeaderView.Stretch)
            continue
        if mode_txt == "contents":
            header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
            continue
        if mode_txt == "fixed":
            header.setSectionResizeMode(column, QHeaderView.Fixed)
            table.setColumnWidth(column, int(width))
            continue
        header.setSectionResizeMode(column, QHeaderView.Interactive)
        table.setColumnWidth(column, int(width))


def _cap_width(widget: QWidget, width: int) -> None:
    widget.setMaximumWidth(width)


def _coerce_editor_qdate(raw: str | None = None, *, fallback_today: bool = True) -> QDate:
    text = str(raw or "").strip()[:10]
    parsed = QDate.fromString(text, "yyyy-MM-dd") if text else QDate()
    if parsed.isValid() and parsed.year() > 2000:
        return parsed
    return QDate.currentDate() if fallback_today else QDate(2000, 1, 1)


def _table_visible_height(table: QTableWidget, rows: int, extra: int = 10) -> int:
    row_height = table.verticalHeader().defaultSectionSize() or 24
    header_height = table.horizontalHeader().height() or 26
    frame = table.frameWidth() * 2
    return header_height + frame + (max(0, int(rows)) * row_height) + extra


def _elide_middle(text: str, max_chars: int = 48) -> str:
    raw = str(text or "").strip()
    if len(raw) <= max_chars:
        return raw
    keep = max(8, (max_chars - 3) // 2)
    return f"{raw[:keep]}...{raw[-keep:]}"


def _selected_row_index(table: QTableWidget) -> int:
    selection_model = table.selectionModel()
    if selection_model is not None:
        selected_rows = selection_model.selectedRows()
        if selected_rows:
            return selected_rows[0].row()
    current_row = table.currentRow()
    return current_row if current_row >= 0 else -1


def _reference_catalog_dialog(parent: QWidget, references: list[dict], title: str = "Histórico de referencias") -> dict | None:
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.resize(1080, 620)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(10)

    header = QHBoxLayout()
    title_lbl = QLabel(title)
    title_lbl.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
    search_edit = QLineEdit()
    search_edit.setPlaceholderText("Pesquisar por ref. interna, externa, descricao, material ou espessura")
    header.addWidget(title_lbl)
    header.addStretch(1)
    header.addWidget(search_edit, 1)
    layout.addLayout(header)

    card = CardFrame()
    card.set_tone("default")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(12, 12, 12, 12)
    card_layout.setSpacing(8)
    table = QTableWidget(0, 7)
    table.setHorizontalHeaderLabels(["Ref. Interna", "Ref. Externa", "Descricao", "Material", "Esp.", "Tempo", "Preco"])
    table.verticalHeader().setVisible(False)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    table.setEditTriggers(QTableWidget.NoEditTriggers)
    _configure_table(table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6))
    card_layout.addWidget(table)
    layout.addWidget(card, 1)

    result: dict | None = None

    def render() -> None:
        query = search_edit.text().strip().lower()
        filtered = []
        for row in references:
            hay = " | ".join(
                [
                    str(row.get("ref_interna", "") or ""),
                    str(row.get("ref_externa", "") or ""),
                    str(row.get("descricao", "") or ""),
                    str(row.get("material", "") or ""),
                    str(row.get("espessura", "") or ""),
                ]
            ).lower()
            if query and query not in hay:
                continue
            filtered.append(row)
        table.setRowCount(len(filtered))
        for row_index, row in enumerate(filtered):
            values = [
                str(row.get("ref_interna", "") or "").strip(),
                str(row.get("ref_externa", "") or "").strip(),
                str(row.get("descricao", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                f"{float(row.get('tempo_peca_min', row.get('tempo_pecas_min', 0)) or 0):.2f}",
                f"{float(row.get('preco_unit', row.get('preco', 0)) or 0):.4f}",
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index in (4, 5, 6):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                if col_index == 0:
                    item.setData(Qt.UserRole, dict(row))
                table.setItem(row_index, col_index, item)
        if filtered:
            table.selectRow(0)

    def accept_selected() -> None:
        nonlocal result
        row_index = _selected_row_index(table)
        if row_index < 0:
            return
        item = table.item(row_index, 0)
        result = dict(item.data(Qt.UserRole) or {}) if item is not None else None
        if result:
            dialog.accept()

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(accept_selected)
    buttons.rejected.connect(dialog.reject)
    buttons.button(QDialogButtonBox.Ok).setText("Carregar referencia")
    layout.addWidget(buttons)

    search_edit.textChanged.connect(render)
    table.itemDoubleClicked.connect(lambda *_args: accept_selected())
    render()
    if dialog.exec() != QDialog.Accepted:
        return None
    return result


def _split_client_label(text: str) -> tuple[str, str]:
    raw = str(text or "").strip()
    if not raw:
        return "", ""
    if " - " in raw:
        left, right = raw.split(" - ", 1)
        return left.strip(), right.strip()
    parts = raw.split(None, 1)
    if len(parts) == 2 and parts[0].upper().startswith("CL"):
        return parts[0].strip(), parts[1].strip()
    return raw, ""


def _format_client_label(text: str, *, show_name: bool = True) -> str:
    code, name = _split_client_label(text)
    if show_name and name:
        return f"{code} - {name}".strip(" -")
    return code or name or "-"


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


def _piece_ops_progress(piece: dict, current_operation: str = "") -> dict[str, float | int]:
    ops = list(piece.get("ops", []) or [])
    if not ops:
        pending = list(piece.get("pendentes", []) or [])
        total = len(pending)
        return {"total": total, "done": 0, "running": 0, "pending": total, "progress_pct": 0.0}
    current_tokens = {str(token).strip().lower() for token in str(current_operation or "").split("+") if str(token).strip()}
    total = len(ops)
    done = 0
    running = 0
    pending = 0
    weighted_done = 0.0
    planned_qty = float(piece.get("planeado", piece.get("quantidade_pedida", 0)) or 0)
    previous_output = planned_qty
    for index, op in enumerate(ops):
        name = str(op.get("nome", "") or "").strip().lower()
        state = str(op.get("estado", "") or "").strip().lower()
        done_qty = float(op.get("qtd_ok", 0) or 0) + float(op.get("qtd_nok", 0) or 0) + float(op.get("qtd_qual", 0) or 0)
        capacity = max(0.0, planned_qty if index == 0 else previous_output)
        if done_qty <= 0 and "concl" in state and capacity > 0:
            done_qty = capacity
        completion = 1.0 if capacity <= 0 else max(0.0, min(1.0, done_qty / capacity))
        previous_output = done_qty if done_qty > 0 else previous_output
        weighted_done += completion
        if completion >= 1.0 or "concl" in state:
            done += 1
        elif completion > 0 or "curso" in state or "produc" in state or (name and name in current_tokens):
            running += 1
        else:
            pending += 1
    progress = 0.0 if total <= 0 else round((weighted_done / total) * 100.0, 1)
    if done >= total and total > 0:
        progress = 100.0
    return {"total": total, "done": done, "running": running, "pending": pending, "progress_pct": progress}


def _normalize_operation_text(raw: str) -> str:
    txt = str(raw or "").strip().lower()
    if "laser" in txt:
        return "Corte Laser"
    if "quin" in txt:
        return "Quinagem"
    if "rosc" in txt:
        return "Roscagem"
    if "embal" in txt:
        return "Embalamento"
    if "mont" in txt:
        return "Montagem"
    if "sold" in txt:
        return "Soldadura"
    if "furo" in txt:
        return "Furo Manual"
    return str(raw or "").strip()


def _operations_for_posto(posto: str, operations: list[str]) -> list[str]:
    posto_norm = str(posto or "").strip().lower()
    normalized = [_normalize_operation_text(op) for op in list(operations or []) if str(op or "").strip()]
    if not posto_norm or posto_norm == "geral":
        seen: set[str] = set()
        out: list[str] = []
        for op in normalized:
            key = op.lower()
            if key not in seen:
                seen.add(key)
                out.append(op)
        return out
    keyword_map = {
        "laser": "laser",
        "quinagem": "quin",
        "roscagem": "rosc",
        "embalamento": "embal",
        "montagem": "mont",
        "soldadura": "sold",
        "furo manual": "furo",
    }
    keyword = next((token for key, token in keyword_map.items() if key in posto_norm), "")
    if not keyword:
        return normalized
    filtered = [op for op in normalized if keyword in op.lower()]
    return filtered


def _fmt_eur(value: float) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


def _make_inline_progress(value: float) -> QProgressBar:
    pct = int(round(float(value or 0)))
    bar = QProgressBar()
    bar.setRange(0, 100)
    bar.setValue(max(0, min(100, pct)))
    bar.setFormat(f"{float(value or 0):.1f}%")
    bar.setTextVisible(True)
    bar.setMaximumHeight(14)
    bar.setStyleSheet(
        "QProgressBar {"
        " background: #edf2f7;"
        " border: 1px solid #c6d2e0;"
        " border-radius: 7px;"
        " color: #10253d;"
        " font-size: 10px;"
        " text-align: center;"
        "}"
        "QProgressBar::chunk {"
        " background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #f59e0b, stop:1 #ea580c);"
        " border-radius: 6px;"
        "}"
    )
    return bar


def _apply_progress_style(bar: QProgressBar, *, compact: bool = False) -> None:
    height = 14 if compact else 18
    radius = 6 if compact else 7
    bar.setStyleSheet(
        "QProgressBar {"
        " background: #edf2f7;"
        " border: 1px solid #c6d2e0;"
        f" border-radius: {radius}px;"
        " color: #10253d;"
        f" font-size: {'10px' if compact else '11px'};"
        " text-align: center;"
        " font-weight: 700;"
        f" min-height: {height}px;"
        "}"
        "QProgressBar::chunk {"
        " background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #f59e0b, stop:1 #ea580c);"
        f" border-radius: {max(4, radius - 1)}px;"
        "}"
    )


def _operation_tokens(raw: str) -> list[str]:
    return [part.strip() for part in str(raw or "").split("+") if part.strip()]


def _build_operation_selector(
    values: list[str], initial_text: str = "", on_change: callable | None = None
) -> tuple[QWidget, QLineEdit, callable]:
    host = QWidget()
    host_layout = QVBoxLayout(host)
    host_layout.setContentsMargins(0, 0, 0, 0)
    host_layout.setSpacing(8)
    summary = QLineEdit()
    summary.setReadOnly(True)
    summary.setPlaceholderText("Seleciona os postos de trabalho")
    host_layout.addWidget(summary)
    grid = QGridLayout()
    grid.setContentsMargins(0, 0, 0, 0)
    grid.setHorizontalSpacing(10)
    grid.setVerticalSpacing(6)
    host_layout.addLayout(grid)
    checks: list[tuple[str, QCheckBox]] = []
    syncing = {"value": False}

    def sync_summary(*, user_initiated: bool = False) -> None:
        selected = [label for label, checkbox in checks if checkbox.isChecked()]
        summary.setText(" + ".join(selected))
        if callable(on_change) and not bool(syncing["value"]):
            on_change(summary.text().strip(), user_initiated)

    for index, value in enumerate([str(item or "").strip() for item in values if str(item or "").strip()]):
        checkbox = QCheckBox(value)
        checkbox.toggled.connect(lambda _checked, _sync=sync_summary: _sync(user_initiated=True))
        checks.append((value, checkbox))
        grid.addWidget(checkbox, index // 3, index % 3)

    def apply_text(text: str) -> None:
        selected = {token.lower() for token in _operation_tokens(text)}
        changed = False
        syncing["value"] = True
        for label, checkbox in checks:
            state = label.lower() in selected if selected else False
            if checkbox.isChecked() != state:
                checkbox.setChecked(state)
                changed = True
        if not selected and checks and not any(checkbox.isChecked() for _label, checkbox in checks):
            checks[0][1].setChecked(True)
            changed = True
        syncing["value"] = False
        if not changed:
            sync_summary(user_initiated=False)
        else:
            sync_summary(user_initiated=False)

    apply_text(initial_text)
    return host, summary, apply_text


_OPERATION_PRICING_MODE_ITEMS: list[tuple[str, str]] = [
    ("manual", "Manual"),
    ("per_piece", "Por peca"),
    ("per_feature", "Por quantidade"),
    ("per_area_m2", "Por m2"),
]


def _operation_pricing_mode_label(value: str) -> str:
    raw = str(value or "").strip().lower()
    for key, label in _OPERATION_PRICING_MODE_ITEMS:
        if raw == key:
            return label
    return "Manual"


def _open_operation_cost_profiles_dialog(parent: QWidget, backend) -> bool:
    settings = dict(backend.operation_cost_settings() or {})
    active_profile = str(settings.get("active_profile", "Base") or "Base").strip() or "Base"
    profiles = dict(settings.get("profiles", {}) or {})
    active_map = dict(profiles.get(active_profile, {}) or {})
    operations = [str(op or "").strip() for op in list(backend.desktop_main.OFF_OPERACOES_DISPONIVEIS) if str(op or "").strip()]

    dialog = QDialog(parent)
    dialog.setWindowTitle("Perfis de custo por operacao")
    dialog.resize(1400, 540)
    dialog.setStyleSheet(
        "QLabel { font-size: 10px; }"
        "QLineEdit, QComboBox, QDoubleSpinBox { font-size: 10px; min-height: 24px; padding: 0 6px; }"
        "QTableWidget { font-size: 10px; gridline-color: #d3dde9; alternate-background-color: #f8fbff; }"
        "QTableWidget::item { padding: 0 8px; }"
        "QHeaderView::section { font-size: 10px; padding: 6px 8px; }"
    )
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)
    info = QLabel(
        "Define a logica base de custo por posto. O orcamentista pode depois ajustar o detalhe por linha no orcamento."
    )
    info.setWordWrap(True)
    info.setProperty("role", "muted")
    info.setStyleSheet("font-size: 10px;")
    layout.addWidget(info)

    profile_row = QHBoxLayout()
    profile_row.setContentsMargins(2, 0, 2, 0)
    profile_row.setSpacing(8)
    profile_row.addWidget(QLabel("Perfil ativo"))
    profile_name_edit = QLineEdit(active_profile)
    profile_name_edit.setPlaceholderText("Base")
    profile_name_edit.setMaximumHeight(28)
    profile_row.addWidget(profile_name_edit, 1)
    layout.addLayout(profile_row)

    table = QTableWidget(len(operations), 9)
    table.setHorizontalHeaderLabels(
        ["Operacao", "Modo", "Etiqueta", "Qtd. padrao", "Setup min", "Tempo base", "EUR/h", "Fixo/un", "Min/un"]
    )
    table.verticalHeader().setVisible(False)
    table.setAlternatingRowColors(False)
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignCenter)
    _configure_table(table, stretch=(), contents=())
    table.setShowGrid(True)
    header.setSectionResizeMode(0, QHeaderView.Interactive)
    header.setSectionResizeMode(1, QHeaderView.Interactive)
    header.setSectionResizeMode(2, QHeaderView.Interactive)
    for col_index in range(3, 9):
        header.setSectionResizeMode(col_index, QHeaderView.Interactive)
    table.setColumnWidth(0, 230)
    table.setColumnWidth(1, 165)
    table.setColumnWidth(2, 350)
    table.setColumnWidth(3, 132)
    table.setColumnWidth(4, 132)
    table.setColumnWidth(5, 132)
    table.setColumnWidth(6, 132)
    table.setColumnWidth(7, 132)
    table.setColumnWidth(8, 132)
    for row_index in range(len(operations)):
        table.setRowHeight(row_index, 34)
    layout.addWidget(table, 1)

    def _cell_host(widget: QWidget, *, left: int = 6, right: int = 6) -> QWidget:
        host = QWidget()
        host_layout = QHBoxLayout(host)
        host_layout.setContentsMargins(left, 3, right, 3)
        host_layout.setSpacing(0)
        host_layout.addWidget(widget)
        return host

    row_controls: list[dict[str, Any]] = []
    for row_index, op_name in enumerate(operations):
        profile = dict(active_map.get(op_name, {}) or {})
        name_item = QTableWidgetItem(op_name)
        name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
        name_item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
        table.setItem(row_index, 0, name_item)

        mode_combo = QComboBox()
        for key, label in _OPERATION_PRICING_MODE_ITEMS:
            mode_combo.addItem(label, key)
        current_mode = str(profile.get("pricing_mode", "manual") or "manual").strip() or "manual"
        for combo_index in range(mode_combo.count()):
            if str(mode_combo.itemData(combo_index) or "") == current_mode:
                mode_combo.setCurrentIndex(combo_index)
                break
        driver_edit = QLineEdit(str(profile.get("driver_label", "") or "Qtd./peca").strip())
        driver_units_spin = QDoubleSpinBox()
        driver_units_spin.setRange(0.0, 1000000.0)
        driver_units_spin.setDecimals(4)
        driver_units_spin.setValue(float(profile.get("default_units", 1) or 1))
        setup_spin = QDoubleSpinBox()
        setup_spin.setRange(0.0, 1000000.0)
        setup_spin.setDecimals(3)
        setup_spin.setValue(float(profile.get("setup_min", 0) or 0))
        time_spin = QDoubleSpinBox()
        time_spin.setRange(0.0, 1000000.0)
        time_spin.setDecimals(4)
        time_spin.setValue(float(profile.get("unit_time_min", 0) or 0))
        hour_spin = QDoubleSpinBox()
        hour_spin.setRange(0.0, 1000000.0)
        hour_spin.setDecimals(4)
        hour_spin.setValue(float(profile.get("hour_rate_eur", 0) or 0))
        fixed_spin = QDoubleSpinBox()
        fixed_spin.setRange(0.0, 1000000.0)
        fixed_spin.setDecimals(4)
        fixed_spin.setValue(float(profile.get("fixed_unit_eur", 0) or 0))
        min_spin = QDoubleSpinBox()
        min_spin.setRange(0.0, 1000000.0)
        min_spin.setDecimals(4)
        min_spin.setValue(float(profile.get("min_unit_eur", 0) or 0))
        for widget in (mode_combo, driver_edit, driver_units_spin, setup_spin, time_spin, hour_spin, fixed_spin, min_spin):
            widget.setMaximumHeight(28)
            widget.setStyleSheet("font-size: 10px;")

        table.setCellWidget(row_index, 1, _cell_host(mode_combo))
        table.setCellWidget(row_index, 2, _cell_host(driver_edit))
        table.setCellWidget(row_index, 3, _cell_host(driver_units_spin))
        table.setCellWidget(row_index, 4, _cell_host(setup_spin))
        table.setCellWidget(row_index, 5, _cell_host(time_spin))
        table.setCellWidget(row_index, 6, _cell_host(hour_spin))
        table.setCellWidget(row_index, 7, _cell_host(fixed_spin))
        table.setCellWidget(row_index, 8, _cell_host(min_spin))
        row_controls.append(
            {
                "op_name": op_name,
                "mode_combo": mode_combo,
                "driver_edit": driver_edit,
                "driver_units_spin": driver_units_spin,
                "setup_spin": setup_spin,
                "time_spin": time_spin,
                "hour_spin": hour_spin,
                "fixed_spin": fixed_spin,
                "min_spin": min_spin,
            }
        )

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.setStyleSheet("QPushButton { font-size: 10.5px; min-height: 28px; padding: 4px 10px; }")
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return False

    profile_name = profile_name_edit.text().strip() or "Base"
    profile_payload: dict[str, Any] = {}
    for controls in row_controls:
        mode_combo = controls["mode_combo"]
        profile_payload[str(controls["op_name"])] = {
            "pricing_mode": str(mode_combo.currentData() or "manual"),
            "driver_label": controls["driver_edit"].text().strip() or "Qtd./peca",
            "default_units": controls["driver_units_spin"].value(),
            "setup_min": controls["setup_spin"].value(),
            "unit_time_min": controls["time_spin"].value(),
            "hour_rate_eur": controls["hour_spin"].value(),
            "fixed_unit_eur": controls["fixed_spin"].value(),
            "min_unit_eur": controls["min_spin"].value(),
        }
    merged_profiles = dict(profiles)
    merged_profiles[profile_name] = profile_payload
    backend.operation_cost_save_settings({"active_profile": profile_name, "profiles": merged_profiles})
    return True


def _open_quote_operation_detail_dialog(parent: QWidget, backend, payload: dict[str, Any]) -> dict[str, Any] | None:
    row = dict(payload or {})
    if not str(row.get("operacao", "") or "").strip():
        QMessageBox.warning(parent, "Operacoes", "Seleciona primeiro os postos de trabalho desta linha.")
        return None
    dialog = QDialog(parent)
    dialog.setWindowTitle("Detalhe de custo por operacao")
    dialog.resize(1320, 470)
    dialog.setStyleSheet(
        "QLabel { font-size: 10px; }"
        "QLineEdit, QComboBox, QDoubleSpinBox { font-size: 10px; min-height: 24px; padding: 0 6px; }"
        "QTableWidget { font-size: 10px; gridline-color: #d3dde9; alternate-background-color: #f8fbff; }"
        "QTableWidget::item { padding: 0 8px; }"
        "QHeaderView::section { font-size: 10px; padding: 6px 8px; }"
        "QPushButton { font-size: 10.5px; min-height: 28px; padding: 4px 10px; }"
    )
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)
    intro = QLabel(
        "Detalha o custo/tempo por posto. Em 'Qtd/peca' introduzes a quantidade tecnica da operacao: "
        "Quinagem = n. de dobras por peca, Roscagem = n. de roscas por peca, Maquinacao = n. de operacoes por peca."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    intro.setStyleSheet("font-size: 10px;")
    layout.addWidget(intro)

    summary_label = QLabel("")
    summary_label.setWordWrap(True)
    summary_label.setProperty("role", "field_value")
    summary_label.setStyleSheet("font-size: 10px; font-weight: 700;")
    layout.addWidget(summary_label)

    table = QTableWidget(0, 11)
    table.setHorizontalHeaderLabels(
        ["Operacao", "Modo", "Tipo qtd.", "Qtd/peca", "Setup", "Tempo base/manual", "EUR/h", "Fixo/manual", "Min/un", "Tempo final/un", "Custo final/un"]
    )
    table.verticalHeader().setVisible(False)
    _configure_table(table, stretch=(), contents=())
    table.setShowGrid(True)
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignCenter)
    _set_table_columns(
        table,
        [
            (0, "fixed", 170),
            (1, "fixed", 155),
            (2, "fixed", 190),
            (3, "fixed", 118),
            (4, "fixed", 118),
            (5, "fixed", 150),
            (6, "fixed", 112),
            (7, "fixed", 138),
            (8, "fixed", 112),
            (9, "fixed", 130),
            (10, "fixed", 138),
        ],
    )
    layout.addWidget(table)

    controls_by_op: dict[str, dict[str, Any]] = {}
    last_estimate: dict[str, Any] = {}

    def _cell_host(widget: QWidget, *, left: int = 6, right: int = 6) -> QWidget:
        host = QWidget()
        host_layout = QHBoxLayout(host)
        host_layout.setContentsMargins(left, 3, right, 3)
        host_layout.setSpacing(0)
        host_layout.addWidget(widget)
        return host

    def _driver_tooltip(op_name: str, driver_label: str) -> str:
        normalized = str(op_name or "").strip().lower()
        if "quin" in normalized:
            return "Quinagem: indica aqui o numero de dobras por peca."
        if "rosc" in normalized:
            return "Roscagem: indica aqui o numero de roscas por peca."
        if "maquin" in normalized:
            return "Maquinacao: indica aqui o numero de operacoes de maquinação por peca."
        if "sold" in normalized:
            return "Soldadura: indica aqui o numero de pontos/cordoes por peca."
        if "serralh" in normalized:
            return "Serralharia: indica aqui o numero de operacoes por peca."
        if "laca" in normalized:
            return "Lacagem: indica aqui a area em m2 por peca."
        if "montag" in normalized:
            return "Montagem: indica aqui o numero de operacoes por peca."
        if "embal" in normalized:
            return "Embalamento: indica aqui o numero de volumes por peca."
        return f"Indica aqui o valor tecnico de '{driver_label or 'Qtd/peca'}'."

    def _refresh_table_height() -> None:
        visible_rows = min(max(4, table.rowCount()), 8)
        target_height = _table_visible_height(table, visible_rows, extra=24)
        table.setMinimumHeight(target_height)
        table.setMaximumHeight(target_height)

    def _set_combo_value(combo: QComboBox, value: str) -> None:
        for combo_index in range(combo.count()):
            if str(combo.itemData(combo_index) or "") == str(value or ""):
                combo.setCurrentIndex(combo_index)
                return

    def _build_payload_from_widgets() -> dict[str, Any]:
        detail_rows: list[dict[str, Any]] = []
        for op_name, controls in controls_by_op.items():
            mode = str(controls["mode_combo"].currentData() or "manual")
            manual_time_value = controls["time_spin"].value()
            manual_cost_value = controls["fixed_spin"].value()
            manual_confirmed = bool(controls.get("manual_confirmed", False))
            if mode == "manual" and (abs(float(manual_time_value or 0)) > 0.000001 or abs(float(manual_cost_value or 0)) > 0.000001):
                manual_confirmed = True
            detail_rows.append(
                {
                    "nome": op_name,
                    "pricing_mode": mode,
                    "driver_label": controls["driver_edit"].text().strip() or "Qtd./peca",
                    "driver_units": controls["driver_units_spin"].value(),
                    "driver_units_confirmed": mode != "manual",
                    "manual_values_confirmed": manual_confirmed,
                    "setup_min": controls["setup_spin"].value(),
                    "unit_time_base_min": controls["time_spin"].value(),
                    "hour_rate_eur": controls["hour_spin"].value(),
                    "fixed_unit_eur": controls["fixed_spin"].value(),
                    "min_unit_eur": controls["min_spin"].value(),
                    "tempo_unit_min": controls["time_spin"].value() if mode == "manual" and manual_confirmed else None,
                    "custo_unit_eur": controls["fixed_spin"].value() if mode == "manual" and manual_confirmed else None,
                }
            )
        return {**row, "operacoes_detalhe": detail_rows}

    def _apply_estimate(estimate: dict[str, Any], *, overwrite_inputs: bool) -> None:
        operations = list(estimate.get("operations", []) or [])
        summary = dict(estimate.get("summary", {}) or {})
        profile_name = str(estimate.get("active_profile", "") or "").strip() or "Base"
        if overwrite_inputs:
            table.setRowCount(len(operations))
            for current_row in range(len(operations)):
                table.setRowHeight(current_row, 34)
        for row_index, op_row in enumerate(operations):
            op_name = str(op_row.get("nome", "") or "").strip()
            if not op_name:
                continue
            controls = controls_by_op.get(op_name)
            if controls is None:
                op_item = QTableWidgetItem(op_name)
                op_item.setFlags(op_item.flags() & ~Qt.ItemIsEditable)
                op_item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
                table.setItem(row_index, 0, op_item)
                mode_combo = QComboBox()
                for key, label in _OPERATION_PRICING_MODE_ITEMS:
                    mode_combo.addItem(label, key)
                driver_edit = QLineEdit()
                driver_units_spin = QDoubleSpinBox()
                driver_units_spin.setRange(0.0, 1000000.0)
                driver_units_spin.setDecimals(4)
                setup_spin = QDoubleSpinBox()
                setup_spin.setRange(0.0, 1000000.0)
                setup_spin.setDecimals(4)
                time_spin = QDoubleSpinBox()
                time_spin.setRange(0.0, 1000000.0)
                time_spin.setDecimals(4)
                hour_spin = QDoubleSpinBox()
                hour_spin.setRange(0.0, 1000000.0)
                hour_spin.setDecimals(4)
                fixed_spin = QDoubleSpinBox()
                fixed_spin.setRange(0.0, 1000000.0)
                fixed_spin.setDecimals(4)
                min_spin = QDoubleSpinBox()
                min_spin.setRange(0.0, 1000000.0)
                min_spin.setDecimals(4)
                for widget in (mode_combo, driver_edit, driver_units_spin, setup_spin, time_spin, hour_spin, fixed_spin, min_spin):
                    widget.setMaximumHeight(28)
                    widget.setStyleSheet("font-size: 10px;")
                table.setCellWidget(row_index, 1, _cell_host(mode_combo))
                table.setCellWidget(row_index, 2, _cell_host(driver_edit))
                table.setCellWidget(row_index, 3, _cell_host(driver_units_spin))
                table.setCellWidget(row_index, 4, _cell_host(setup_spin))
                table.setCellWidget(row_index, 5, _cell_host(time_spin))
                table.setCellWidget(row_index, 6, _cell_host(hour_spin))
                table.setCellWidget(row_index, 7, _cell_host(fixed_spin))
                table.setCellWidget(row_index, 8, _cell_host(min_spin))
                controls = {
                    "mode_combo": mode_combo,
                    "driver_edit": driver_edit,
                    "driver_units_spin": driver_units_spin,
                    "setup_spin": setup_spin,
                    "time_spin": time_spin,
                    "hour_spin": hour_spin,
                    "fixed_spin": fixed_spin,
                    "min_spin": min_spin,
                    "manual_confirmed": False,
                }
                controls_by_op[op_name] = controls
            if overwrite_inputs:
                _set_combo_value(controls["mode_combo"], str(op_row.get("pricing_mode", "manual") or "manual"))
                driver_label_txt = str(op_row.get("driver_label", "") or "Qtd./peca").strip()
                controls["driver_edit"].setText(driver_label_txt)
                controls["driver_units_spin"].setValue(float(op_row.get("driver_units", 0) or 0))
                controls["setup_spin"].setValue(float(op_row.get("setup_min", 0) or 0))
                controls["time_spin"].setValue(float(op_row.get("unit_time_base_min", op_row.get("tempo_unit_min", 0)) or 0))
                controls["hour_spin"].setValue(float(op_row.get("hour_rate_eur", 0) or 0))
                controls["fixed_spin"].setValue(float(op_row.get("fixed_unit_eur", op_row.get("custo_unit_eur", 0)) or 0))
                controls["min_spin"].setValue(float(op_row.get("min_unit_eur", 0) or 0))
                controls["manual_confirmed"] = bool(op_row.get("manual_values_confirmed", False))
                tip = _driver_tooltip(op_name, driver_label_txt)
                controls["driver_edit"].setToolTip(tip)
                controls["driver_units_spin"].setToolTip(tip)
            time_result = QTableWidgetItem("-" if op_row.get("tempo_unit_min") in (None, "") else f"{float(op_row.get('tempo_unit_min', 0) or 0):.3f}")
            cost_result = QTableWidgetItem("-" if op_row.get("custo_unit_eur") in (None, "") else _fmt_eur(float(op_row.get("custo_unit_eur", 0) or 0)))
            time_result.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            cost_result.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, 9, time_result)
            table.setItem(row_index, 10, cost_result)
        mode_txt = str(summary.get("costing_mode", "") or "").strip() or "aggregate_pending"
        pending_rows = [dict(item) for item in operations if isinstance(item, dict) and bool(item.get("missing_driver_input"))]
        blend_with_current_line = bool(row.get("blend_with_current_line", False))
        base_time = float(row.get("base_tempo_unit_min", row.get("tempo_peca_min", 0)) or 0)
        base_cost = float(row.get("base_preco_unit_eur", row.get("preco_unit", 0)) or 0)
        state_txt = {
            "detailed": "Detalhe completo",
            "partial_detail": "Detalhe parcial",
            "aggregate_pending": "Ainda agregado",
            "single_operation_total": "Linha simples",
        }.get(mode_txt, mode_txt)
        if pending_rows:
            missing_txt = ", ".join(
                f"{str(item.get('nome', '') or '').strip()} ({str(item.get('driver_label', '') or 'Qtd./peca').strip()})"
                for item in pending_rows
                if str(item.get("nome", "") or "").strip()
            )
            summary_label.setText(f"Perfil {profile_name} | falta quantificar: {missing_txt}")
        elif blend_with_current_line and (base_time > 0 or base_cost > 0):
            extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
            extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
            summary_label.setText(
                f"Perfil {profile_name} | base atual {_fmt_eur(base_cost)}/un + extras {_fmt_eur(extra_cost)}/un = "
                f"{_fmt_eur(base_cost + extra_cost)}/un | tempo {base_time:.3f} + {extra_time:.3f} = {base_time + extra_time:.3f} min/un"
            )
        else:
            summary_label.setText(
                f"Perfil {profile_name} | {state_txt} | sugestao {float(summary.get('tempo_unit_total_min', 0) or 0):.3f} min/un "
                f"| {_fmt_eur(float(summary.get('custo_unit_total_eur', 0) or 0))}/un"
            )
        _refresh_table_height()

    def recompute() -> None:
        nonlocal last_estimate
        last_estimate = dict(backend.operation_cost_estimate(_build_payload_from_widgets()) or {})
        _apply_estimate(last_estimate, overwrite_inputs=False)

    initial_estimate = dict(backend.operation_cost_estimate(row) or {})
    _apply_estimate(initial_estimate, overwrite_inputs=True)
    last_estimate = dict(initial_estimate)

    for controls in controls_by_op.values():
        controls["mode_combo"].currentIndexChanged.connect(lambda _idx: recompute())
        controls["driver_edit"].textChanged.connect(lambda _txt: recompute())
        controls["driver_units_spin"].valueChanged.connect(lambda _val: recompute())
        controls["setup_spin"].valueChanged.connect(lambda _val: recompute())
        controls["time_spin"].valueChanged.connect(lambda _val: recompute())
        controls["hour_spin"].valueChanged.connect(lambda _val: recompute())
        controls["fixed_spin"].valueChanged.connect(lambda _val: recompute())
        controls["min_spin"].valueChanged.connect(lambda _val: recompute())

    action_row = QHBoxLayout()
    action_row.setContentsMargins(2, 0, 2, 0)
    action_row.setSpacing(8)
    apply_profiles_btn = QPushButton("Aplicar perfis")
    apply_profiles_btn.setProperty("variant", "secondary")
    config_profiles_btn = QPushButton("Configurar perfis")
    config_profiles_btn.setProperty("variant", "secondary")
    action_row.addWidget(apply_profiles_btn)
    action_row.addWidget(config_profiles_btn)
    action_row.addStretch(1)
    layout.addLayout(action_row)

    def apply_profiles_defaults() -> None:
        nonlocal last_estimate
        clean_payload = dict(row)
        clean_payload["operacoes_detalhe"] = []
        last_estimate = dict(backend.operation_cost_estimate(clean_payload) or {})
        _apply_estimate(last_estimate, overwrite_inputs=True)
        recompute()

    apply_profiles_btn.clicked.connect(apply_profiles_defaults)

    def configure_profiles() -> None:
        if not _open_operation_cost_profiles_dialog(dialog, backend):
            return
        apply_profiles_defaults()

    config_profiles_btn.clicked.connect(configure_profiles)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return None

    final_estimate = dict(backend.operation_cost_estimate(_build_payload_from_widgets()) or {})
    final_rows = [dict(item or {}) for item in list(final_estimate.get("operations", []) or []) if isinstance(item, dict)]
    tempos_operacao = {
        str(item.get("nome", "") or "").strip(): float(item.get("tempo_unit_min", 0) or 0)
        for item in final_rows
        if str(item.get("nome", "") or "").strip() and item.get("tempo_unit_min") not in (None, "")
    }
    custos_operacao = {
        str(item.get("nome", "") or "").strip(): float(item.get("custo_unit_eur", 0) or 0)
        for item in final_rows
        if str(item.get("nome", "") or "").strip() and item.get("custo_unit_eur") not in (None, "")
    }
    summary = dict(final_estimate.get("summary", {}) or {})
    base_time_unit = round(float(row.get("base_tempo_unit_min", row.get("tempo_peca_min", 0)) or 0), 4)
    base_price_unit = round(float(row.get("base_preco_unit_eur", row.get("preco_unit", 0)) or 0), 4)
    blend_with_current_line = bool(row.get("blend_with_current_line", False))
    detailed_extras_time = round(float(summary.get("tempo_unit_total_min", 0) or 0), 4)
    detailed_extras_price = round(float(summary.get("custo_unit_total_eur", 0) or 0), 4)
    suggested_time_unit = round(base_time_unit + detailed_extras_time, 4) if blend_with_current_line else detailed_extras_time
    suggested_price_unit = round(base_price_unit + detailed_extras_price, 4) if blend_with_current_line else detailed_extras_price
    return {
        "operacoes_detalhe": final_rows,
        "tempos_operacao": tempos_operacao,
        "custos_operacao": custos_operacao,
        "quote_cost_snapshot": {
            "costing_mode": str(summary.get("costing_mode", "") or ""),
            "tempo_total_peca_min": round(float(summary.get("tempo_unit_total_min", 0) or 0), 4),
            "preco_unit_total_eur": round(float(summary.get("custo_unit_total_eur", 0) or 0), 4),
            "qtd": round(float(summary.get("qtd", row.get("qtd", 0)) or 0), 2),
            "cost_profile": str(final_estimate.get("active_profile", "") or "").strip(),
        },
        "suggested_tempo_unit_min": suggested_time_unit,
        "suggested_preco_unit_eur": suggested_price_unit,
        "apply_totals": bool(summary.get("complete", False) or blend_with_current_line),
        "blended_with_current_line": blend_with_current_line,
        "summary": summary,
    }


class PulsePage(QWidget):
    page_title = "Pulse"
    page_subtitle = "OEE, desvios, paragens e peças em curso com o mesmo backend do mobile."
    allow_auto_timer_refresh = True

    def __init__(self, runtime_service, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        self.last_pulse_data: dict = {}
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(8)
        self.period_combo = QComboBox()
        self.period_combo.addItems(["Hoje", "7 dias", "30 dias", "Tudo"])
        self.period_combo.setCurrentText("7 dias")
        self.year_combo = QComboBox()
        current_year = str(datetime.now().year)
        self.year_combo.addItems([current_year, "Todos"])
        self.year_combo.setCurrentText(current_year)
        self.origin_combo = QComboBox()
        self.origin_combo.addItems(["Ambos", "Em curso", "Histórico"])
        self.view_combo = QComboBox()
        self.view_combo.addItems(["Todas", "So desvio"])
        self.graphs_btn = QPushButton("Graficos")
        self.graphs_btn.setProperty("variant", "secondary")
        self.graphs_btn.clicked.connect(self._show_graphs_dialog)
        self.plan_delay_btn = QPushButton("Atrasos planeamento")
        self.plan_delay_btn.setProperty("variant", "secondary")
        self.plan_delay_btn.clicked.connect(self._show_plan_delay_dialog)
        for widget in (self.period_combo, self.year_combo, self.origin_combo, self.view_combo):
            widget.currentTextChanged.connect(self.refresh)
        for widget, width in ((self.period_combo, 140), (self.year_combo, 110), (self.origin_combo, 130), (self.view_combo, 130)):
            _cap_width(widget, width)
        filters_layout.addWidget(QLabel("Periodo"))
        filters_layout.addWidget(self.period_combo)
        filters_layout.addWidget(QLabel("Ano"))
        filters_layout.addWidget(self.year_combo)
        filters_layout.addWidget(QLabel("Origem"))
        filters_layout.addWidget(self.origin_combo)
        filters_layout.addWidget(QLabel("Visao"))
        filters_layout.addWidget(self.view_combo)
        filters_layout.addStretch(1)
        filters_layout.addWidget(self.plan_delay_btn)
        filters_layout.addWidget(self.graphs_btn)
        root.addWidget(filters)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(10)
        cards_layout.setVerticalSpacing(10)
        self.cards = [StatCard(title) for title in ("OEE", "Disponibilidade", "Performance", "Paragens", "Desvio max.")]
        for index, card in enumerate(self.cards):
            card.setMaximumHeight(112)
            cards_layout.addWidget(card, 0, index)
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("success")
        self.cards[2].set_tone("warning")
        self.cards[3].set_tone("danger")
        self.cards[4].set_tone("warning")
        root.addWidget(cards_host)

        self.alert_card = CardFrame()
        alert_layout = QVBoxLayout(self.alert_card)
        alert_layout.setContentsMargins(14, 12, 14, 12)
        alert_layout.setSpacing(8)
        alert_title = QLabel("Alertas")
        alert_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.alert_label = QLabel("-")
        self.alert_label.setWordWrap(True)
        alert_layout.addWidget(alert_title)
        alert_layout.addWidget(self.alert_label)
        self.alert_card.setMaximumHeight(120)
        self.alert_card.set_tone("default")

        self.running_table = QTableWidget(0, 6)
        self.running_table.setHorizontalHeaderLabels(["Encomenda", "Peca", "Operacao", "Operador", "Tempo", "Desvio"])
        self.running_table.verticalHeader().setVisible(False)
        self.running_table.verticalHeader().setDefaultSectionSize(24)
        self.running_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.running_table, stretch=(1, 2), contents=(0, 3, 4, 5))

        self.stops_table = QTableWidget(0, 5)
        self.stops_table.setHorizontalHeaderLabels(["Causa", "Encomenda", "Operador", "Ocorrencias", "Minutos"])
        self.stops_table.verticalHeader().setVisible(False)
        self.stops_table.verticalHeader().setDefaultSectionSize(24)
        self.stops_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.stops_table, stretch=(0, 1), contents=(2, 3, 4))

        self.history_table = QTableWidget(0, 5)
        self.history_table.setHorizontalHeaderLabels(["Encomenda", "Ops", "Tempo", "Planeado", "Desvio"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.verticalHeader().setDefaultSectionSize(24)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.history_table, stretch=(0,), contents=(1, 2, 3, 4))

        running_card = CardFrame()
        running_card.set_tone("success")
        running_layout = QVBoxLayout(running_card)
        running_layout.setContentsMargins(14, 12, 14, 12)
        running_layout.setSpacing(8)
        running_title = QLabel("Pecas em curso")
        running_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        running_layout.addWidget(running_title)
        running_layout.addWidget(self.running_table)

        stops_card = CardFrame()
        stops_card.set_tone("danger")
        stops_layout = QVBoxLayout(stops_card)
        stops_layout.setContentsMargins(14, 12, 14, 12)
        stops_layout.setSpacing(8)
        stops_title = QLabel("Top causas de paragem")
        stops_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        stops_layout.addWidget(stops_title)
        stops_layout.addWidget(self.stops_table)

        history_card = CardFrame()
        history_card.set_tone("info")
        history_layout = QVBoxLayout(history_card)
        history_layout.setContentsMargins(14, 12, 14, 12)
        history_layout.setSpacing(8)
        history_title = QLabel("Histórico consolidado")
        history_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        history_layout.addWidget(history_title)
        history_layout.addWidget(self.history_table)

        right_stack = QWidget()
        right_stack_layout = QVBoxLayout(right_stack)
        right_stack_layout.setContentsMargins(0, 0, 0, 0)
        right_stack_layout.setSpacing(10)
        right_stack_layout.addWidget(self.alert_card)
        right_stack_layout.addWidget(stops_card, 1)

        top_split = QSplitter(Qt.Horizontal)
        top_split.setChildrenCollapsible(False)
        top_split.addWidget(running_card)
        top_split.addWidget(right_stack)
        top_split.setSizes([1180, 720])
        root.addWidget(top_split, 1)
        root.addWidget(history_card, 1)

    def refresh(self) -> None:
        year = self.year_combo.currentText()
        data = self.runtime_service.dashboard(
            period=self.period_combo.currentText(),
            year=None if year == "Todos" else year,
            visao=self.view_combo.currentText(),
            origem=self.origin_combo.currentText(),
        )
        self.last_pulse_data = dict(data or {})
        summary = data.get("summary", {})
        self.cards[0].set_data(f"{summary.get('oee', 0):.1f}%", f"Atualizado {data.get('updated_at', '-')}")
        self.cards[1].set_data(f"{summary.get('disponibilidade', 0):.1f}%", f"Qualidade {summary.get('qualidade', 0):.1f}%")
        perf_plan = float(summary.get("perf_plan_total", 0) or 0)
        perf_real = float(summary.get("perf_real_total", 0) or 0)
        self.cards[2].set_data(f"{summary.get('performance', 0):.1f}%", f"Plano {perf_plan:.1f} | Real {perf_real:.1f}")
        self.cards[3].set_data(f"{summary.get('paragens_min', 0):.1f} min", f"Fora do tempo {summary.get('pecas_fora_tempo', 0)}")
        self.cards[4].set_data(f"{summary.get('desvio_max_min', 0):.1f} min", f"Em curso {summary.get('pecas_em_curso', 0)}")
        perf_value = float(summary.get("performance", 0) or 0)
        self.cards[2].set_tone("success" if perf_value >= 100.0 else ("warning" if perf_value >= 80.0 else "danger"))
        self.alert_label.setText(str(summary.get("alerts", "-")))
        _set_panel_tone(self.alert_card, "danger" if str(summary.get("alerts", "") or "-").strip() not in {"", "-"} else "default")
        plan_delay = dict(data.get("plan_delay", {}) or {})
        plan_delay_open = int(plan_delay.get("open_count", 0) or 0)
        plan_delay_ack = int(plan_delay.get("acknowledged_count", 0) or 0)
        self.plan_delay_btn.setText(f"Atrasos planeamento ({plan_delay_open})" if plan_delay_open > 0 else "Atrasos planeamento")
        self.plan_delay_btn.setToolTip(
            f"{plan_delay_open} pendente(s) | {plan_delay_ack} justificado(s)"
            if (plan_delay_open or plan_delay_ack)
            else "Sem grupos fora do horário planeado neste momento."
        )
        self.plan_delay_btn.setProperty(
            "variant",
            "danger" if plan_delay_open > 0 else ("secondary" if plan_delay_ack <= 0 else "warning"),
        )
        _repolish(self.plan_delay_btn)
        _fill_table(
            self.running_table,
            [[r.get("encomenda", "-"), r.get("peca", "-"), r.get("operacao", "-"), r.get("operador", "-"), f"{r.get('elapsed_min', 0):.1f} min", f"{r.get('delta_min', 0):.1f}"] for r in data.get("running", [])],
            align_center_from=4,
        )
        _fill_table(
            self.stops_table,
            [[r.get("causa", "-"), r.get("encomenda", "-"), r.get("operador", "-"), r.get("ocorrencias", 0), f"{r.get('minutos', 0):.1f}"] for r in data.get("top_stops", [])],
            align_center_from=3,
        )
        _fill_table(
            self.history_table,
            [[r.get("encomenda", "-"), r.get("ops", 0), f"{r.get('elapsed_min', 0):.1f}", f"{r.get('plan_min', 0):.1f}", f"{r.get('delta_min', 0):.1f}"] for r in data.get("history", [])],
            align_center_from=1,
        )

    def _prompt_plan_delay_reason(self, current_reason: str = "") -> str:
        options = [
            "Mudança de prioridade / urgência",
            "Matéria-prima ainda não disponível",
            "Aguardar posto / máquina ocupada",
            "Avaria / paragem no posto",
            "Decisão da chefia / replaneado manualmente",
        ]
        default_idx = 0
        current_txt = str(current_reason or "").strip()
        if current_txt:
            for idx, option in enumerate(options):
                if option.lower() == current_txt.lower():
                    default_idx = idx
                    break
        reason_txt, ok = QInputDialog.getItem(
            self,
            "Motivo do atraso ao planeamento",
            "Indica o motivo para retirar este grupo do aviso ativo:",
            options,
            default_idx,
            True,
        )
        if not ok:
            return ""
        return str(reason_txt or "").strip()

    def _show_plan_delay_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Atrasos Face Ao Planeamento")
        dialog.setMinimumWidth(980)
        dialog.setMinimumHeight(520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Grupos de corte laser fora do horário planeado")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        info_label = QLabel("")
        info_label.setWordWrap(True)
        info_label.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(info_label)

        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Material", "Esp.", "Planeado", "Posto", "Estado", "Motivo"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(26)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        _configure_table(table, stretch=(1, 4, 7), contents=(0, 2, 3, 5, 6))
        layout.addWidget(table, 1)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        justify_btn = QPushButton("Sinalizar motivo")
        justify_btn.setProperty("variant", "warning")
        reactivate_btn = QPushButton("Reativar aviso")
        reactivate_btn.setProperty("variant", "secondary")
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        actions.addWidget(justify_btn)
        actions.addWidget(reactivate_btn)
        actions.addStretch(1)
        actions.addWidget(refresh_btn)
        actions.addWidget(close_btn)
        layout.addLayout(actions)

        row_map: dict[int, dict] = {}

        def _reload_dashboard() -> None:
            self.refresh()
            plan_delay = dict((self.last_pulse_data or {}).get("plan_delay", {}) or {})
            items = list(plan_delay.get("items", []) or [])
            info_label.setText(
                f"{int(plan_delay.get('open_count', 0) or 0)} pendente(s) | "
                f"{int(plan_delay.get('acknowledged_count', 0) or 0)} justificado(s). "
                "Este aviso é só face ao planeamento atual; não significa atraso final ao cliente."
            )
            row_map.clear()
            table.setRowCount(len(items))
            for row_idx, item in enumerate(items):
                values = [
                    str(item.get("numero", "") or "-"),
                    str(item.get("cliente", "") or "-"),
                    str(item.get("material", "") or "-"),
                    str(item.get("espessura", "") or "-"),
                    str(item.get("planned_end_txt", "") or item.get("planned_start_txt", "") or "-"),
                    str(item.get("posto", "") or "-"),
                    str(item.get("status_label", "") or "-"),
                    str(item.get("reason", "") or "-"),
                ]
                for col_idx, value in enumerate(values):
                    table_item = QTableWidgetItem(value)
                    if col_idx in {0, 2, 3, 4, 5, 6}:
                        table_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    table.setItem(row_idx, col_idx, table_item)
                status_open = not bool(item.get("acknowledged"))
                bg = QColor("#fff4e5" if status_open else "#ecfdf3")
                fg = QColor("#9a3412" if status_open else "#166534")
                for col_idx in range(table.columnCount()):
                    cell = table.item(row_idx, col_idx)
                    if cell is not None:
                        cell.setBackground(QBrush(bg))
                        if col_idx == 6:
                            cell.setForeground(QBrush(fg))
                row_map[row_idx] = dict(item)
            if items:
                table.selectRow(0)

        def _selected_item() -> dict | None:
            current_row = table.currentRow()
            return dict(row_map.get(current_row, {}) or {}) if current_row >= 0 else None

        def _justify_selected() -> None:
            item = _selected_item()
            if not item:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", "Seleciona uma linha primeiro.")
                return
            reason_txt = self._prompt_plan_delay_reason(str(item.get("reason", "") or "").strip())
            if not reason_txt:
                return
            try:
                self.runtime_service.pulse_plan_delay_set_reason(str(item.get("item_key", "") or "").strip(), reason_txt)
            except Exception as exc:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", str(exc))
                return
            _reload_dashboard()

        def _reactivate_selected() -> None:
            item = _selected_item()
            if not item:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", "Seleciona uma linha primeiro.")
                return
            try:
                self.runtime_service.pulse_plan_delay_clear_reason(str(item.get("item_key", "") or "").strip())
            except Exception as exc:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", str(exc))
                return
            _reload_dashboard()

        justify_btn.clicked.connect(_justify_selected)
        reactivate_btn.clicked.connect(_reactivate_selected)
        refresh_btn.clicked.connect(_reload_dashboard)
        close_btn.clicked.connect(dialog.accept)
        _reload_dashboard()
        dialog.exec()

    def _show_graphs_dialog(self) -> None:
        summary = dict((self.last_pulse_data or {}).get("summary", {}) or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Gráficos Pulse")
        dialog.setMinimumWidth(720)
        dialog.setMinimumHeight(480)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        title = QLabel("Performance e eficiencia operacional")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel(
            f"Periodo {self.period_combo.currentText()} | Ano {self.year_combo.currentText()} | Origem {self.origin_combo.currentText()} | "
            f"Plano {float(summary.get('perf_plan_total', 0) or 0):.1f} min | Real {float(summary.get('perf_real_total', 0) or 0):.1f} min"
        )
        subtitle.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        metrics_card = CardFrame()
        metrics_card.set_tone("info")
        metrics_layout = QVBoxLayout(metrics_card)
        metrics_layout.setContentsMargins(12, 12, 12, 12)
        metrics_layout.setSpacing(10)
        no_data_scope = (
            not list((self.last_pulse_data or {}).get("running", []) or [])
            and not list((self.last_pulse_data or {}).get("history", []) or [])
            and float(summary.get("perf_plan_total", 0) or 0) <= 0
            and float(summary.get("perf_real_total", 0) or 0) <= 0
        )
        for label_text, value, tone in (
            ("OEE", 0.0 if no_data_scope else float(summary.get("oee", 0) or 0), "info"),
            ("Disponibilidade", 0.0 if no_data_scope else float(summary.get("disponibilidade", 0) or 0), "success"),
            ("Performance", 0.0 if no_data_scope else float(summary.get("performance", 0) or 0), "warning"),
            ("Qualidade", 0.0 if no_data_scope else float(summary.get("qualidade", 0) or 0), "success"),
        ):
            row_host = QWidget()
            row_layout = QVBoxLayout(row_host)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(4)
            row_head = QHBoxLayout()
            row_head.setContentsMargins(0, 0, 0, 0)
            name = QLabel(label_text)
            name.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
            value_lbl = QLabel(f"{value:.1f}%")
            value_lbl.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
            row_head.addWidget(name)
            row_head.addStretch(1)
            row_head.addWidget(value_lbl)
            bar = QProgressBar()
            max_range = 250 if label_text == "Performance" else 100
            bar.setRange(0, max_range)
            bar.setValue(max(0, min(max_range, int(round(value)))))
            bar.setTextVisible(False)
            _apply_progress_style(bar, compact=True)
            if tone == "warning":
                bar.setStyleSheet(
                    "QProgressBar {background: #fff7ed; border: 1px solid #f3d6aa; border-radius: 8px; min-height: 14px; text-align: center; color: #9a3412; font-weight: 700;}"
                    "QProgressBar::chunk {background: #ea580c; border-radius: 7px;}"
                )
            row_layout.addLayout(row_head)
            row_layout.addWidget(bar)
            metrics_layout.addWidget(row_host)
        layout.addWidget(metrics_card)

        andon = dict(summary.get("andon", {}) or {})
        andon_card = CardFrame()
        andon_card.set_tone("default")
        andon_layout = QGridLayout(andon_card)
        andon_layout.setContentsMargins(12, 12, 12, 12)
        andon_layout.setHorizontalSpacing(10)
        andon_layout.setVerticalSpacing(8)
        for idx, (label_text, key) in enumerate((("Produção", "prod"), ("Setup", "setup"), ("Espera", "espera"), ("Parado", "stop"))):
            box = CardFrame()
            box.set_tone("warning" if key in {"setup", "espera"} else ("danger" if key == "stop" else "success"))
            box_layout = QVBoxLayout(box)
            box_layout.setContentsMargins(10, 10, 10, 10)
            lbl = QLabel(label_text)
            lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #0f172a;")
            val = QLabel(str(int(andon.get(key, 0) or 0)))
            val.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
            box_layout.addWidget(lbl)
            box_layout.addWidget(val)
            andon_layout.addWidget(box, 0, idx)
        layout.addWidget(andon_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        buttons.button(QDialogButtonBox.Close).setText("Fechar")
        layout.addWidget(buttons)
        dialog.exec()


class OperatorPage(QWidget):
    page_title = "Operador"
    page_subtitle = "Resumo operacional dos grupos ativos, peças e progresso por encomenda."
    uses_backend_reload = True

    def __init__(self, runtime_service, backend, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        self.backend = backend
        self.all_items: list[dict] = []
        self.items: list[dict] = []
        self.current_pieces: list[dict] = []
        self.checked_piece_ids: set[str] = set()
        self._syncing_piece_checks = False
        self.selected_group_key: tuple[str, str, str] | None = None
        self.selected_piece_id = ""
        self.ui_options: dict[str, bool] = {}
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        self.cards = [StatCard(title) for title in ("Encomendas ativas", "Pecas em curso", "Pecas em avaria", "Progresso")]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, index)
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("success")
        self.cards[2].set_tone("danger")
        self.cards[3].set_tone("warning")
        root.addWidget(cards_host)

        self.global_progress = QProgressBar()
        self.global_progress.setRange(0, 100)
        self.global_progress.setFormat("%p%")
        _apply_progress_style(self.global_progress)
        root.addWidget(self.global_progress)

        self.control_card = CardFrame()
        self.control_card.set_tone("info")
        control_layout = QVBoxLayout(self.control_card)
        control_layout.setContentsMargins(16, 14, 16, 14)
        control_layout.setSpacing(10)

        selectors_row = QHBoxLayout()
        selectors_row.setSpacing(10)
        self.operator_combo = QComboBox()
        self.operator_combo.setEditable(True)
        self.operator_combo.currentTextChanged.connect(self._sync_operator_assignment)
        self.posto_combo = QComboBox()
        self.posto_combo.setEditable(True)
        self.posto_combo.addItems(["Geral", "Laser", "Quinagem", "Roscagem", "Embalamento", "Montagem", "Soldadura"])
        self.posto_combo.setCurrentText("Geral")
        self.posto_combo.currentTextChanged.connect(self._update_piece_context)
        self.operation_combo = QComboBox()
        self.operation_combo.setEditable(False)
        selectors_row.addWidget(QLabel("Operador"))
        selectors_row.addWidget(self.operator_combo, 1)
        selectors_row.addWidget(QLabel("Posto"))
        selectors_row.addWidget(self.posto_combo)
        selectors_row.addWidget(QLabel("Operacao"))
        selectors_row.addWidget(self.operation_combo, 1)
        control_layout.addLayout(selectors_row)
        self.options_btn = QPushButton("Opcoes")
        self.options_btn.setProperty("variant", "secondary")
        self.options_btn.clicked.connect(self._open_options_dialog)

        actions_row = QVBoxLayout()
        actions_row.setSpacing(8)
        actions_row_top = QHBoxLayout()
        actions_row_top.setSpacing(8)
        actions_row_bottom = QHBoxLayout()
        actions_row_bottom.setSpacing(8)
        self.start_btn = QPushButton("Iniciar")
        self.start_btn.clicked.connect(self._start_piece)
        self.finish_btn = QPushButton("Finalizar")
        self.finish_btn.clicked.connect(self._finish_piece)
        self.resume_btn = QPushButton("Retomar")
        self.resume_btn.clicked.connect(self._resume_piece)
        self.pause_btn = QPushButton("Interromper")
        self.pause_btn.setProperty("variant", "secondary")
        self.pause_btn.clicked.connect(self._pause_piece)
        self.avaria_btn = QPushButton("Registar Avaria")
        self.avaria_btn.setProperty("variant", "danger")
        self.avaria_btn.clicked.connect(self._register_avaria)
        self.close_avaria_btn = QPushButton("Fim Avaria")
        self.close_avaria_btn.setProperty("variant", "secondary")
        self.close_avaria_btn.clicked.connect(self._close_avaria)
        self.alert_btn = QPushButton("Alertar Chefia")
        self.alert_btn.clicked.connect(self._alert_chefia)
        self.manual_consume_btn = QPushButton("Dar Baixa")
        self.manual_consume_btn.setProperty("variant", "secondary")
        self.manual_consume_btn.clicked.connect(self._manual_consume_material)
        self.drawing_btn = QPushButton("Ver desenho")
        self.drawing_btn.setProperty("variant", "secondary")
        self.drawing_btn.clicked.connect(self._open_drawing)
        self.labels_btn = QPushButton("Etiquetas")
        self.labels_btn.setProperty("variant", "secondary")
        self.labels_btn.clicked.connect(self._open_labels_dialog)
        self.local_refresh_btn = QPushButton("Atualizar")
        self.local_refresh_btn.setProperty("variant", "secondary")
        self.local_refresh_btn.clicked.connect(self.refresh)
        for button in (
            self.start_btn,
            self.finish_btn,
            self.resume_btn,
            self.pause_btn,
            self.avaria_btn,
            self.close_avaria_btn,
        ):
            actions_row_top.addWidget(button)
        for button in (
            self.alert_btn,
            self.manual_consume_btn,
            self.drawing_btn,
            self.labels_btn,
            self.local_refresh_btn,
        ):
            actions_row_bottom.addWidget(button)
        actions_row_top.addStretch(1)
        actions_row_bottom.addStretch(1)
        actions_row.addLayout(actions_row_top, 1)
        actions_row.addLayout(actions_row_bottom, 1)
        control_layout.addLayout(actions_row)

        self.feedback_label = QLabel("Seleciona uma peca para operar.")
        self.feedback_label.setWordWrap(True)
        self.feedback_label.setProperty("role", "muted")
        control_layout.addWidget(self.feedback_label)
        root.addWidget(self.control_card)

        self.context_card = CardFrame()
        context_layout = QVBoxLayout(self.context_card)
        context_layout.setContentsMargins(16, 14, 16, 14)
        context_layout.setSpacing(8)
        context_header = QHBoxLayout()
        self.piece_title_label = QLabel("Sem peca selecionada")
        self.piece_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.piece_state_chip = QLabel("-")
        _apply_state_chip(self.piece_state_chip, "-")
        context_header.addWidget(self.piece_title_label, 1)
        context_header.addWidget(self.options_btn, 0, Qt.AlignRight)
        context_header.addWidget(self.piece_state_chip, 0, Qt.AlignRight)
        context_layout.addLayout(context_header)
        self.piece_meta_label = QLabel("Seleciona um grupo e uma peca para ver contexto, pendencias e bloqueios.")
        self.piece_meta_label.setWordWrap(True)
        self.piece_meta_label.setProperty("role", "muted")
        context_layout.addWidget(self.piece_meta_label)
        self.issue_label = QLabel("-")
        self.issue_label.setWordWrap(True)
        context_layout.addWidget(self.issue_label)
        self.pending_label = QLabel("-")
        self.pending_label.setWordWrap(True)
        self.pending_label.setProperty("role", "muted")
        context_layout.addWidget(self.pending_label)
        self.operation_strip = QWidget()
        self.operation_strip_layout = QHBoxLayout(self.operation_strip)
        self.operation_strip_layout.setContentsMargins(0, 0, 0, 0)
        self.operation_strip_layout.setSpacing(6)
        context_layout.addWidget(self.operation_strip)
        self.piece_progress = QProgressBar()
        self.piece_progress.setRange(0, 100)
        self.piece_progress.setFormat("%p%")
        _apply_progress_style(self.piece_progress, compact=True)
        context_layout.addWidget(self.piece_progress)
        self.context_card.set_tone("default")
        root.addWidget(self.context_card)

        self.groups_table = QTableWidget(0, 9)
        self.groups_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Estado", "Material", "Esp.", "Plan", "Real", "Desvio", "Progress"])
        self.groups_table.verticalHeader().setVisible(False)
        self.groups_table.verticalHeader().setDefaultSectionSize(30)
        self.groups_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.groups_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.groups_table, stretch=(1, 3), contents=(2, 4, 5, 6, 7, 8))
        _set_table_columns(
            self.groups_table,
            [
                (0, "interactive", 190),
                (1, "stretch", 260),
                (2, "interactive", 138),
                (3, "stretch", 190),
                (4, "interactive", 72),
                (5, "interactive", 70),
                (6, "interactive", 70),
                (7, "interactive", 74),
                (8, "interactive", 112),
            ],
        )
        self.groups_table.itemSelectionChanged.connect(self._handle_group_selection)
        self.pieces_table = QTableWidget(0, 11)
        self.pieces_table.setHorizontalHeaderLabels(["Sel.", "Ref. Int.", "Ref. Ext.", "Estado", "Operacao", "Operador", "Produzido", "Tempo", "Plan", "Progress", "Pendentes"])
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.verticalHeader().setDefaultSectionSize(28)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pieces_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        _configure_table(self.pieces_table, stretch=(2, 10), contents=(0, 1, 3, 4, 5, 6, 7, 8))
        _set_table_columns(
            self.pieces_table,
            [
                (0, "fixed", 42),
                (1, "interactive", 170),
                (2, "stretch", 240),
                (3, "interactive", 122),
                (4, "interactive", 132),
                (5, "interactive", 124),
                (6, "interactive", 86),
                (7, "interactive", 76),
                (8, "interactive", 72),
                (9, "interactive", 118),
                (10, "stretch", 200),
            ],
        )
        self.pieces_table.itemSelectionChanged.connect(self._update_piece_context)
        self.pieces_table.itemChanged.connect(self._handle_piece_check_changed)

        self.groups_card = CardFrame()
        self.groups_card.set_tone("info")
        groups_layout = QVBoxLayout(self.groups_card)
        groups_layout.setContentsMargins(16, 14, 16, 14)
        groups_layout.setSpacing(8)
        self.groups_title_label = QLabel("Encomendas ativas no operador")
        self.groups_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        groups_layout.addWidget(self.groups_title_label)
        groups_layout.addWidget(self.groups_table)
        root.addWidget(self.groups_card)

        self.pieces_card = CardFrame()
        self.pieces_card.set_tone("default")
        pieces_layout = QVBoxLayout(self.pieces_card)
        pieces_layout.setContentsMargins(16, 14, 16, 14)
        pieces_layout.setSpacing(8)
        pieces_header = QHBoxLayout()
        pieces_header.setSpacing(8)
        self.pieces_title_label = QLabel("Pecas da encomenda")
        self.pieces_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.select_all_pieces_box = QCheckBox("Selecionar todas")
        self.select_all_pieces_box.toggled.connect(self._toggle_all_piece_checks)
        self.multi_count_chip = QLabel("0 selecionadas")
        _apply_state_chip(self.multi_count_chip, "-", "0 selecionadas")
        pieces_header.addWidget(self.pieces_title_label)
        pieces_header.addStretch(1)
        pieces_header.addWidget(self.multi_count_chip)
        pieces_header.addWidget(self.select_all_pieces_box)
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.pieces_table, 1)
        root.addWidget(self.pieces_card, 1)

    def refresh(self) -> None:
        user = self.backend.user or {}
        getter = getattr(self.backend, "ui_options", None)
        if callable(getter):
            try:
                self.ui_options = dict(getter() or {})
            except Exception:
                self.ui_options = {}
        self._set_combo_items(
            self.operator_combo,
            self.backend.operator_names(),
            preferred=str(user.get("username", "") or "").strip(),
        )
        self._sync_operator_assignment()
        previous_group = self.selected_group_key
        previous_piece = self.selected_piece_id
        data = self.runtime_service.operator_board(username=str(user.get("username", "")), role=str(user.get("role", "")))
        summary = data.get("summary", {})
        all_items = self._hydrate_operator_items(list(data.get("items", [])))
        op_total = 0
        op_done = 0
        op_running = 0
        for board_item in all_items:
            for piece in list(board_item.get("pieces", []) or []):
                stats = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
                op_total += int(stats.get("total", 0) or 0)
                op_done += int(stats.get("done", 0) or 0)
                op_running += int(stats.get("running", 0) or 0)
        global_ops_progress = 0.0 if op_total <= 0 else round(((op_done + (0.5 * op_running)) / op_total) * 100.0, 1)
        self.cards[0].set_data(summary.get("encomendas_ativas", 0), f"Grupos {summary.get('grupos', 0)}")
        self.cards[1].set_data(summary.get("pecas_em_curso", 0), f"Em pausa {summary.get('pecas_em_pausa', 0)}")
        self.cards[2].set_data(summary.get("pecas_em_avaria", 0), f"Concluidas {summary.get('pecas_concluidas', 0)}")
        self.cards[3].set_data(f"{global_ops_progress:.1f}%", f"Ops {op_done}/{op_total} | Em curso {op_running}")
        self.global_progress.setValue(int(round(global_ops_progress)))
        self.all_items = all_items
        self.items = list(self.all_items)
        self._render_groups_table(previous_group, previous_piece)

    def _render_groups_table(self, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        self.groups_table.setSortingEnabled(False)
        self.groups_table.blockSignals(True)
        self.groups_table.setRowCount(len(self.items))
        for row_index, item in enumerate(self.items):
            group_piece_stats = [_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")) for piece in list(item.get("pieces", []) or [])]
            group_progress = round(sum(float(stat.get("progress_pct", 0) or 0) for stat in group_piece_stats) / len(group_piece_stats), 1) if group_piece_stats else float(item.get("progress_pct", 0) or 0)
            row_values = [
                item.get("encomenda", "-"),
                _format_client_label(item.get("cliente", "-"), show_name=self._show_client_name()),
                item.get("estado_espessura", item.get("estado", "-")),
                item.get("material", "-"),
                item.get("espessura", "-"),
                f"{item.get('tempo_plan_min', 0):.1f}",
                f"{item.get('tempo_real_min', 0):.1f}",
                f"{item.get('desvio_min', 0):.1f}",
                f"{group_progress:.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, "|".join(self._group_key(item)))
                if col_index >= 4:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.groups_table.setItem(row_index, col_index, cell)
            _paint_table_row(self.groups_table, row_index, str(item.get("estado_espessura", "")))
        self.groups_table.blockSignals(False)
        self.groups_table.setSortingEnabled(True)
        target_row = 0
        if previous_group:
            target_key = "|".join(previous_group)
            for index in range(self.groups_table.rowCount()):
                cell = self.groups_table.item(index, 0)
                if str((cell.data(Qt.UserRole) if cell is not None else "") or "").strip() == target_key:
                    target_row = index
                    break
        if self.groups_table.rowCount() > 0:
            self.groups_table.selectRow(target_row)
            self.selected_group_key = previous_group if previous_group else self._group_key(self._current_group())
            self.selected_piece_id = previous_piece
            self._handle_group_selection()
        else:
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self.pieces_title_label.setText("Pecas da encomenda")
            self._clear_piece_context()

    def _group_key(self, item: dict) -> tuple[str, str, str]:
        return (
            str(item.get("encomenda", "") or "").strip(),
            str(item.get("material", "") or "").strip(),
            str(item.get("espessura", "") or "").strip(),
        )

    def _set_combo_items(self, combo: QComboBox, values: list[str], preferred: str = "") -> None:
        current = combo.currentText().strip()
        target = current or str(preferred or "").strip()
        combo.blockSignals(True)
        combo.clear()
        seen: set[str] = set()
        ordered: list[str] = []
        for raw in list(values or []):
            text = str(raw or "").strip()
            key = text.lower()
            if text and key not in seen:
                seen.add(key)
                ordered.append(text)
        combo.addItems(ordered)
        if target:
            combo.setCurrentText(target)
        elif ordered:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _client_name_map(self) -> dict[str, str]:
        getter = getattr(self.backend, "ensure_data", None)
        if not callable(getter):
            return {}
        try:
            data = getter() or {}
        except Exception:
            return {}
        return {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list((data or {}).get("clientes", []) or [])
            if isinstance(row, dict) and str(row.get("codigo", "") or "").strip()
        }

    def _hydrate_operator_items(self, rows: list[dict]) -> list[dict]:
        clients = self._client_name_map()
        hydrated: list[dict] = []
        for raw in list(rows or []):
            item = dict(raw or {})
            client_raw = str(item.get("cliente", "") or "").strip()
            client_code, client_name = _split_client_label(client_raw)
            if client_code and not client_name:
                client_name = clients.get(client_code, "")
            client_label = _format_client_label(
                f"{client_code} - {client_name}".strip(" -") if (client_code or client_name) else client_raw,
                show_name=True,
            )
            item["cliente"] = client_label or client_raw or "-"
            item["cliente_codigo"] = client_code or str(item.get("cliente_codigo", "") or "").strip()
            item["cliente_nome"] = client_name or str(item.get("cliente_nome", "") or "").strip()
            item["cliente_label"] = client_label or client_raw or "-"
            hydrated.append(item)
        return hydrated

    def _current_group(self) -> dict:
        row_index = _selected_row_index(self.groups_table)
        if row_index < 0:
            return {}
        row_item = self.groups_table.item(row_index, 0)
        group_key = str(row_item.data(Qt.UserRole) or "").strip()
        if group_key:
            for item in self.items:
                if "|".join(self._group_key(item)) == group_key:
                    return item
        if row_index >= len(self.items):
            return {}
        return self.items[row_index]

    def _current_piece(self) -> dict:
        row_index = _selected_row_index(self.pieces_table)
        if row_index < 0:
            return {}
        row_item = self.pieces_table.item(row_index, 0)
        piece_id = str(row_item.data(Qt.UserRole) or "").strip()
        if piece_id:
            for piece in self.current_pieces:
                if str(piece.get("id", "") or "").strip() == piece_id:
                    return piece
        if row_index >= len(self.current_pieces):
            return {}
        return self.current_pieces[row_index]

    def _show_client_name(self) -> bool:
        return bool(self.ui_options.get("operator_show_client_name", True))

    def _open_options_dialog(self) -> None:
        user = dict(self.backend.user or {})
        if str(user.get("role", "") or "").strip().lower() != "admin":
            QMessageBox.information(self, "Opcoes", "Apenas o admin pode alterar estas opcoes.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Opcoes do Operador")
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        info = QLabel("Opcoes visuais e operacionais do menu Operador.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        show_client_box = QCheckBox("Mostrar nome do cliente no operador")
        show_client_box.setChecked(self._show_client_name())
        layout.addWidget(show_client_box)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        setter = getattr(self.backend, "set_ui_option", None)
        if callable(setter):
            setter("operator_show_client_name", bool(show_client_box.isChecked()))
        self.refresh()

    def _selected_piece_ids(self) -> list[str]:
        ids: list[str] = []
        seen: set[str] = set()
        for piece in self.current_pieces:
            piece_id = str(piece.get("id", "") or "").strip()
            if piece_id and piece_id in self.checked_piece_ids and piece_id not in seen:
                ids.append(piece_id)
                seen.add(piece_id)
        if ids:
            return ids
        selection_model = self.pieces_table.selectionModel()
        if selection_model is not None:
            for model_index in selection_model.selectedRows():
                row = model_index.row()
                item = self.pieces_table.item(row, 0)
                piece_id = str((item.data(Qt.UserRole) if item is not None else "") or "").strip()
                if piece_id and piece_id not in seen:
                    ids.append(piece_id)
                    seen.add(piece_id)
        if ids:
            return ids
        piece = self._current_piece()
        piece_id = str(piece.get("id", "") or "").strip()
        return [piece_id] if piece_id else []

    def _selected_piece_refs(self, *, allow_multiple: bool = True) -> list[tuple[str, str]] | None:
        group = self._current_group()
        enc_num = str(group.get("encomenda", "") or "").strip()
        piece_ids = self._selected_piece_ids()
        if not enc_num or not piece_ids:
            QMessageBox.warning(self, "Operador", "Seleciona pelo menos uma peca.")
            return None
        refs = [(enc_num, piece_id) for piece_id in piece_ids if piece_id]
        if not allow_multiple and len(refs) > 1:
            QMessageBox.warning(self, "Operador", "Esta acao exige apenas uma peca selecionada.")
            return None
        return refs

    def _selected_refs(self) -> tuple[str, str] | None:
        refs = self._selected_piece_refs(allow_multiple=False)
        return refs[0] if refs else None

    def _handle_piece_check_changed(self, item: QTableWidgetItem) -> None:
        if self._syncing_piece_checks or item.column() != 0:
            return
        piece_id = str(item.data(Qt.UserRole) or "").strip()
        if not piece_id:
            return
        if item.checkState() == Qt.Checked:
            self.checked_piece_ids.add(piece_id)
        else:
            self.checked_piece_ids.discard(piece_id)
        self._sync_piece_multi_state()

    def _toggle_all_piece_checks(self, checked: bool) -> None:
        if self._syncing_piece_checks:
            return
        self._syncing_piece_checks = True
        try:
            visible_ids = {str(piece.get("id", "") or "").strip() for piece in self.current_pieces if str(piece.get("id", "") or "").strip()}
            if checked:
                self.checked_piece_ids.update(visible_ids)
            else:
                self.checked_piece_ids.difference_update(visible_ids)
            for row_index in range(self.pieces_table.rowCount()):
                item = self.pieces_table.item(row_index, 0)
                if item is not None:
                    item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        finally:
            self._syncing_piece_checks = False
        self._sync_piece_multi_state()

    def _sync_piece_multi_state(self) -> None:
        visible_ids = [str(piece.get("id", "") or "").strip() for piece in self.current_pieces if str(piece.get("id", "") or "").strip()]
        selected_count = sum(1 for piece_id in visible_ids if piece_id in self.checked_piece_ids)
        total_count = len(visible_ids)
        self._syncing_piece_checks = True
        try:
            self.select_all_pieces_box.setChecked(bool(total_count) and selected_count == total_count)
        finally:
            self._syncing_piece_checks = False
        _apply_state_chip(
            self.multi_count_chip,
            "Em producao" if selected_count else "-",
            f"{selected_count} selecionadas",
        )

    def _current_operator(self) -> str:
        return self.operator_combo.currentText().strip()

    def _current_posto(self) -> str:
        return self.posto_combo.currentText().strip() or "Geral"

    def _current_operation(self) -> str:
        return self.operation_combo.currentText().strip()

    def _default_posto_for_operator(self) -> str:
        resolver = getattr(self.backend, "operator_default_posto", None)
        if callable(resolver):
            try:
                return str(resolver(self._current_operator()) or "").strip() or "Geral"
            except Exception:
                return "Geral"
        return "Geral"

    def _sync_operator_assignment(self) -> None:
        default_posto = self._default_posto_for_operator()
        has_assignment = False
        checker = getattr(self.backend, "operator_has_posto_assignment", None)
        if callable(checker):
            try:
                has_assignment = bool(checker(self._current_operator()))
            except Exception:
                has_assignment = False
        current_posto = self._current_posto()
        if default_posto and default_posto != current_posto:
            self.posto_combo.blockSignals(True)
            self.posto_combo.setCurrentText(default_posto)
            self.posto_combo.blockSignals(False)
        self.posto_combo.setEnabled(not has_assignment)
        self._update_piece_context()

    def _set_feedback(self, text: str, error: bool = False) -> None:
        raw = str(text or "").strip() or "-"
        self.feedback_label.setText(_elide_middle(raw, 110))
        self.feedback_label.setToolTip(raw if len(raw) > 110 else "")
        if error:
            self.feedback_label.setStyleSheet("color: #b42318; font-weight: 700;")
            _set_panel_tone(self.control_card, "danger")
        else:
            self.feedback_label.setStyleSheet("color: #475467;")
            _set_panel_tone(self.control_card, "info")

    def _clear_piece_context(self) -> None:
        self.piece_title_label.setText("Sem peca selecionada")
        self.piece_meta_label.setText("Seleciona um grupo e uma peca para ver contexto, pendencias e bloqueios.")
        self.issue_label.setText("-")
        self.pending_label.setText("-")
        _clear_layout_widgets(self.operation_strip_layout)
        _apply_state_chip(self.piece_state_chip, "-")
        self.piece_progress.setValue(0)
        self.operation_combo.clear()
        _set_panel_tone(self.context_card, "default")
        self._sync_piece_multi_state()
        self._set_button_states(False, False, False, False, False, False)

    def _set_button_states(self, has_piece: bool, has_operator: bool, has_pending: bool, has_open_avaria: bool, can_resume: bool, can_finish: bool | None = None) -> None:
        finish_enabled = has_piece and has_operator and (can_finish if can_finish is not None else has_pending) and not has_open_avaria
        self.start_btn.setEnabled(has_piece and has_operator and has_pending and not has_open_avaria)
        self.finish_btn.setEnabled(finish_enabled)
        self.resume_btn.setEnabled(has_piece and has_operator and can_resume and not has_open_avaria)
        self.pause_btn.setEnabled(has_piece and has_operator and not has_open_avaria)
        self.avaria_btn.setEnabled(has_piece and has_operator and not has_open_avaria)
        self.close_avaria_btn.setEnabled(has_piece and has_operator and has_open_avaria)
        self.alert_btn.setEnabled(has_piece and has_operator)
        self.drawing_btn.setEnabled(has_piece)
        self.manual_consume_btn.setEnabled(has_piece)
        self.labels_btn.setEnabled(has_piece)

    def _handle_group_selection(self) -> None:
        item = self._current_group()
        if not item:
            self.pieces_title_label.setText("Pecas da encomenda")
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self._clear_piece_context()
            return
        self.selected_group_key = self._group_key(item)
        client_label = _format_client_label(item.get("cliente", "-"), show_name=self._show_client_name())
        material = str(item.get("material", "-") or "-").strip() or "-"
        espessura = str(item.get("espessura", "-") or "-").strip() or "-"
        self.pieces_title_label.setText(
            f"Pecas da encomenda {item.get('encomenda', '-')} | Cliente {client_label} | {material} {espessura} mm"
        )
        target_piece_id = self.selected_piece_id
        all_pieces = list(item.get("pieces", []) or [])
        state_filter = self.detail_state_filter_combo.currentText().strip().lower() if hasattr(self, "detail_state_filter_combo") else "todas"
        query = self.detail_search_edit.text().strip().lower() if hasattr(self, "detail_search_edit") else ""
        filtered_pieces = []
        for piece in all_pieces:
            state = str(piece.get("estado", "") or "").strip().lower()
            if state_filter and state_filter != "todas":
                if state_filter == "em producao" and not ("produc" in state or "curso" in state):
                    continue
                if state_filter == "concluida" and "concl" not in state:
                    continue
                if state_filter == "em pausa" and not ("paus" in state or "interromp" in state):
                    continue
                if state_filter == "avaria" and "avaria" not in state:
                    continue
            if query:
                haystack = " ".join(
                    [
                        str(piece.get("ref_interna", "") or ""),
                        str(piece.get("ref_externa", "") or ""),
                        str(piece.get("operacao_atual", "") or ""),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            filtered_pieces.append(piece)
        self.current_pieces = filtered_pieces
        self.checked_piece_ids.intersection_update({str(piece.get("id", "") or "").strip() for piece in self.current_pieces})
        self.pieces_table.setSortingEnabled(False)
        self.pieces_table.blockSignals(True)
        self.pieces_table.setRowCount(len(self.current_pieces))
        for row_index, piece in enumerate(self.current_pieces):
            produced = f"{piece.get('produzido', 0):.1f}"
            piece_id = str(piece.get("id", "") or "").strip()
            ops_progress = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked if piece_id in self.checked_piece_ids else Qt.Unchecked)
            check_item.setData(Qt.UserRole, piece_id)
            check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            self.pieces_table.setItem(row_index, 0, check_item)
            row_values = [
                piece.get("ref_interna", "-"),
                piece.get("ref_externa", "-"),
                piece.get("estado", "-"),
                piece.get("operacao_atual", "-"),
                piece.get("operador", "-"),
                produced,
                f"{piece.get('tempo_min', 0):.1f}",
                f"{piece.get('planeado', 0):.1f}",
                "",
                ", ".join(piece.get("pendentes", []) or []),
            ]
            for offset, value in enumerate(row_values, start=1):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if offset in (6, 7, 8):
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.pieces_table.setItem(row_index, offset, cell)
            self.pieces_table.setCellWidget(row_index, 9, _make_inline_progress(float(ops_progress.get("progress_pct", 0) or 0)))
            _paint_table_row(self.pieces_table, row_index, str(piece.get("estado", "")))
        self.pieces_table.blockSignals(False)
        self.pieces_table.setSortingEnabled(False)
        if self.pieces_table.rowCount() == 0:
            self.selected_piece_id = ""
            self._clear_piece_context()
            return
        target_row = 0
        if target_piece_id:
            for index, piece in enumerate(self.current_pieces):
                if str(piece.get("id", "") or "").strip() == target_piece_id:
                    target_row = index
                    break
        self.pieces_table.selectRow(target_row)
        self._sync_piece_multi_state()
        self._update_piece_context()

    def _update_piece_context(self) -> None:
        group = self._current_group()
        piece = self._current_piece()
        if not group or not piece:
            self.selected_piece_id = ""
            self._clear_piece_context()
            return
        enc_num = str(group.get("encomenda", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        self.selected_piece_id = piece_id
        try:
            ctx = self.backend.operator_piece_context(enc_num, piece_id)
        except Exception as exc:
            self._clear_piece_context()
            self._set_feedback(str(exc), error=True)
            return
        live_piece = dict(ctx.get("piece") or {})
        state = str(live_piece.get("estado", "") or piece.get("estado", "-")).strip() or "-"
        pending = list(ctx.get("pending_ops", []) or [])
        done = list(ctx.get("done_ops", []) or [])
        ops_progress = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
        issue_text = "-"
        if ctx.get("has_open_avaria"):
            issue_text = (
                f"Avaria em aberto: {ctx.get('avaria_motivo', '-') or '-'}"
                f" | Aberta {float(ctx.get('avaria_open_min', 0) or 0):.1f} min"
                f" | Total nesta referencia {float(ctx.get('avaria_total_min', 0) or 0):.1f} min"
            )
        else:
            motivo = str(live_piece.get("interrupcao_peca_motivo", "") or "").strip()
            if motivo:
                issue_text = f"Interrupcao registada: {motivo}"
            else:
                issue_text = (
                    f"Fluxo operacional: {int(ops_progress.get('done', 0))}/{int(ops_progress.get('total', 0))} concluidas"
                    f" | Em curso {int(ops_progress.get('running', 0))} | Pendentes {int(ops_progress.get('pending', 0))}"
                )
        full_piece_title = f"{piece.get('ref_interna', '-')} | {piece.get('ref_externa', '-')}"
        self.piece_title_label.setText(_elide_middle(full_piece_title, 64))
        self.piece_title_label.setToolTip(full_piece_title)
        _apply_state_chip(self.piece_state_chip, state)
        current_op_txt = str(ctx.get("current_operation", "") or piece.get("operacao_atual", "") or "-").strip() or "-"
        self.piece_meta_label.setText(
            f"Encomenda {enc_num} | Cliente {_format_client_label(group.get('cliente', '-'), show_name=self._show_client_name())} | "
            f"{group.get('material', '-')} {group.get('espessura', '-')} mm\n"
            f"Operador {piece.get('operador', '-') or '-'} | Operacao atual {current_op_txt} | Quantidade {float(ctx.get('produzido_ok', 0) or 0):.1f}/{float(ctx.get('quantidade_pedida', 0) or 0):.1f} | "
            f"Fluxo {int(ops_progress.get('done', 0))}/{int(ops_progress.get('total', 0))} | Tempo aberto {float(ctx.get('current_operation_elapsed_min', 0) or 0):.1f} min | "
            f"Avarias acumuladas {float(ctx.get('avaria_total_min', 0) or 0):.1f} min"
        )
        self.issue_label.setText(issue_text)
        self.issue_label.setStyleSheet("color: #b42318; font-weight: 700;" if ctx.get("has_open_avaria") else "color: #475467;")
        pend_txt = ", ".join(pending[:3]) if pending else "-"
        done_txt = ", ".join(done[:3]) if done else "-"
        if len(pending) > 3:
            pend_txt = f"{pend_txt}, +{len(pending) - 3}"
        if len(done) > 3:
            done_txt = f"{done_txt}, +{len(done) - 3}"
        self.pending_label.setText(f"Pendentes: {pend_txt} | Concluidas: {done_txt}")
        _clear_layout_widgets(self.operation_strip_layout)
        for op in list(piece.get("ops", []) or []):
            chip = QLabel()
            op_name = str(op.get("nome", "") or "-")
            op_state = str(op.get("estado", "") or "-")
            _apply_state_chip(chip, op_state, _elide_middle(op_name, 16))
            self.operation_strip_layout.addWidget(chip)
        self.operation_strip_layout.addStretch(1)
        self.piece_progress.setValue(int(round(float(ops_progress.get("progress_pct", 0) or 0))))
        _set_panel_tone(self.context_card, _state_tone(state))
        self._sync_piece_multi_state()
        filtered_pending = _operations_for_posto(self._current_posto(), pending or list(piece.get("pendentes", []) or []))
        active_pending = list(ctx.get("active_pending_ops", []) or [])
        filtered_active_pending = _operations_for_posto(self._current_posto(), active_pending)
        preferred_op = (filtered_active_pending[0] if filtered_active_pending else "") or self._current_operation() or str(ctx.get("current_operation", "") or str(piece.get("operacao_atual", "") or "").split(" + ")[0])
        self._set_combo_items(
            self.operation_combo,
            filtered_pending,
            preferred=preferred_op,
        )
        state_norm = state.lower()
        can_resume = ("paus" in state_norm) or ("interromp" in state_norm)
        self._set_button_states(
            True,
            bool(self._current_operator()),
            self.operation_combo.count() > 0,
            bool(ctx.get("has_open_avaria")),
            can_resume,
            can_finish=bool(filtered_active_pending),
        )

    def _prompt_reason(self, title: str, label: str, options: list[str], default_value: str = "") -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        prompt = QLabel(label)
        prompt.setWordWrap(True)
        combo = QComboBox()
        combo.setEditable(True)
        self._set_combo_items(combo, options, preferred=default_value)
        layout.addWidget(prompt)
        layout.addWidget(combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return combo.currentText().strip() or None

    def _collect_start_operation_targets(self, refs_list: list[tuple[str, str]]) -> dict[str, object]:
        ordered_ops: list[str] = []
        targets: dict[str, list[tuple[str, str]]] = {}
        skipped_without_ops = 0
        errors: list[str] = []
        posto = self._current_posto()
        for enc_num, piece_id in list(refs_list or []):
            try:
                ctx = self.backend.operator_piece_context(enc_num, piece_id)
            except Exception as exc:
                errors.append(str(exc))
                continue
            pending = _operations_for_posto(posto, list(ctx.get("pending_ops", []) or []))
            if not pending:
                skipped_without_ops += 1
                continue
            for op_name in pending:
                op_key = str(op_name or "").strip()
                if not op_key:
                    continue
                if op_key not in targets:
                    targets[op_key] = []
                    ordered_ops.append(op_key)
                if (enc_num, piece_id) not in targets[op_key]:
                    targets[op_key].append((enc_num, piece_id))
        return {
            "ordered_ops": ordered_ops,
            "targets": targets,
            "skipped_without_ops": skipped_without_ops,
            "errors": errors,
        }

    def _prompt_start_operation(
        self,
        refs_list: list[tuple[str, str]],
        ordered_ops: list[str],
        targets: dict[str, list[tuple[str, str]]],
        skipped_without_ops: int = 0,
    ) -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Iniciar operacao")
        dialog.setMinimumWidth(460)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            f"Selecionaste {len(refs_list)} peca(s). Escolhe a operacao que queres iniciar. "
            "So aparecem operacoes realmente pendentes nas pecas selecionadas."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        posto = str(self._current_posto() or "").strip() or "Geral"
        posto_label = QLabel(f"Posto atual: {posto}")
        posto_label.setProperty("role", "muted")
        layout.addWidget(posto_label)
        if skipped_without_ops:
            skipped_label = QLabel(
                f"{skipped_without_ops} peca(s) ficaram fora porque nao têm operacoes pendentes neste posto."
            )
            skipped_label.setWordWrap(True)
            skipped_label.setProperty("role", "muted")
            layout.addWidget(skipped_label)
        list_widget = QListWidget()
        for op_name in list(ordered_ops or []):
            compatible = len(list(targets.get(op_name, []) or []))
            item = QListWidgetItem(f"{op_name} ({compatible} peca{'s' if compatible != 1 else ''})")
            item.setData(Qt.UserRole, op_name)
            list_widget.addItem(item)
        preferred = self._current_operation()
        preferred_row = 0
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            op_name = str((item.data(Qt.UserRole) if item is not None else "") or "").strip()
            if preferred and op_name == preferred:
                preferred_row = index
                break
        if list_widget.count() > 0:
            list_widget.setCurrentRow(preferred_row)
        list_widget.itemDoubleClicked.connect(lambda _item: dialog.accept())
        layout.addWidget(list_widget)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        current_item = list_widget.currentItem()
        if current_item is None:
            QMessageBox.warning(self, "Operador", "Seleciona uma operacao para iniciar.")
            return None
        return str(current_item.data(Qt.UserRole) or "").strip() or None

    def _prompt_finish(self, ctx: dict, *, batch_idx: int = 0, batch_total: int = 0, preferred_operation: str = "") -> dict[str, float | str] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Finalizar operacao ({batch_idx}/{batch_total})" if batch_idx and batch_total else "Finalizar operacao")
        dialog.setMinimumWidth(480)
        layout = QVBoxLayout(dialog)
        piece = dict(ctx.get("piece") or {})
        ref_int = str(piece.get("ref_interna", "") or "-").strip() or "-"
        ref_ext = str(piece.get("ref_externa", "") or "-").strip() or "-"
        title = QLabel(f"{ref_int} | {ref_ext}")
        title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        meta = QLabel(
            f"Qtd {float(ctx.get('produzido_ok', 0) or 0):.1f}/{float(ctx.get('quantidade_pedida', 0) or 0):.1f}"
            f" | Pendentes: {', '.join(list(ctx.get('pending_ops', []) or [])[:3]) or '-'}"
        )
        meta.setProperty("role", "muted")
        meta.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(meta)
        form = QFormLayout()
        op_combo = QComboBox()
        finish_ops = list(ctx.get("active_pending_ops", []) or []) or list(ctx.get("pending_ops", []) or [])
        self._set_combo_items(
            op_combo,
            finish_ops,
            preferred=preferred_operation or self._current_operation(),
        )
        ok_spin = QDoubleSpinBox()
        nok_spin = QDoubleSpinBox()
        qual_spin = QDoubleSpinBox()
        for spin in (ok_spin, nok_spin, qual_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(1)
            spin.setSingleStep(1.0)
        ok_spin.setValue(float(ctx.get("default_ok", 0) or 0))
        form.addRow("Operacao", op_combo)
        form.addRow("OK", ok_spin)
        form.addRow("NOK", nok_spin)
        form.addRow("Qualidade", qual_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "operation": op_combo.currentText().strip(),
            "ok": ok_spin.value(),
            "nok": nok_spin.value(),
            "qual": qual_spin.value(),
        }

    def _prompt_laser_stock_resolution(self, stock_state: dict) -> dict[str, float | str | bool] | None:
        total_qty = float(stock_state.get("total_qty", 0) or 0)
        reserved_qty = float(stock_state.get("reserved_qty", 0) or 0)
        remaining_qty = float(stock_state.get("remaining_qty", 0) or 0)
        manual_stock_required = bool(stock_state.get("manual_stock_required", False))
        reserved_sources = list(stock_state.get("reserved_sources", []) or [])
        has_reserved = reserved_qty > 0
        if not manual_stock_required and not has_reserved:
            return {"material_id": "", "quantity": 0.0, "allow_without_stock": False, "retalho": {}, "source_material_id": ""}
        candidates = list(stock_state.get("candidates", []) or [])
        if not candidates and not has_reserved:
            answer = QMessageBox.question(
                self,
                "Baixa material Laser",
                (
                    f"Laser concluido para {stock_state.get('material', '-')} {stock_state.get('espessura', '-')} mm.\n\n"
                    f"Peças produzidas: {total_qty:.1f}\n"
                    f"Baixa automatica por cativacao: {reserved_qty:.1f}\n"
                    "Baixa manual de stock: obrigatoria\n\n"
                    "Nao existe stock disponivel correspondente.\n"
                    "Pretende concluir sem baixa adicional?"
                ),
                QMessageBox.Yes | QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return None
            return {"material_id": "", "quantity": 0.0, "allow_without_stock": True, "retalho": {}, "source_material_id": ""}

        dialog = QDialog(self)
        dialog.setWindowTitle("Baixa material Laser")
        dialog.setMinimumWidth(720)
        layout = QVBoxLayout(dialog)
        info_lines = [
            f"Laser concluido para {stock_state.get('material', '-')} {stock_state.get('espessura', '-')} mm",
            f"Peças produzidas: {total_qty:.1f} | Baixa automatica por cativacao: {reserved_qty:.1f}",
        ]
        if has_reserved:
            info_lines.append("Existe material cativado. Neste fecho final nao e permitida baixa manual adicional; so podes registar retalho.")
        else:
            info_lines.append("Sem material cativado. Indica a quantidade de stock realmente consumida e o lote utilizado.")
        info = QLabel("\n".join(info_lines))
        info.setWordWrap(True)
        layout.addWidget(info)
        form = QFormLayout()
        material_combo = QComboBox()
        for row in candidates:
            label = (
                f"{row.get('material_id', '-') } | {row.get('dimensao', '-') } | "
                f"Disp. {float(row.get('disponivel', 0) or 0):.1f} | {row.get('local', '-') } | {row.get('lote', '-') }"
            )
            material_combo.addItem(label, row)
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.0, 1000000.0)
        qty_spin.setDecimals(1)
        qty_spin.setSingleStep(1.0)
        qty_spin.setValue(0.0)
        skip_box = QCheckBox("Concluir sem baixa adicional")
        skip_box.toggled.connect(lambda checked: (material_combo.setEnabled(not checked), qty_spin.setEnabled(not checked)))
        form.addRow("Stock", material_combo)
        form.addRow("Qtd stock consumido", qty_spin)
        form.addRow("", skip_box)
        if has_reserved:
            material_combo.setEnabled(False)
            qty_spin.setEnabled(False)
            qty_spin.setValue(0.0)
            skip_box.setChecked(True)
            skip_box.setEnabled(False)
        layout.addLayout(form)

        retalho_card = CardFrame()
        retalho_card.set_tone("warning")
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(12, 10, 12, 10)
        retalho_layout.setHorizontalSpacing(8)
        retalho_layout.setVerticalSpacing(6)
        retalho_layout.addWidget(QLabel("Retalho comprimento"), 0, 0)
        retalho_layout.addWidget(QLabel("Retalho largura"), 0, 1)
        retalho_layout.addWidget(QLabel("Qtd retalho"), 0, 2)
        retalho_layout.addWidget(QLabel("Metros"), 0, 3)
        comp_spin = QDoubleSpinBox()
        larg_spin = QDoubleSpinBox()
        qtd_retalho_spin = QDoubleSpinBox()
        metros_spin = QDoubleSpinBox()
        for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(2)
            spin.setAlignment(Qt.AlignCenter)
        retalho_layout.addWidget(comp_spin, 1, 0)
        retalho_layout.addWidget(larg_spin, 1, 1)
        retalho_layout.addWidget(qtd_retalho_spin, 1, 2)
        retalho_layout.addWidget(metros_spin, 1, 3)
        retalho_layout.addWidget(QLabel("Lote origem do retalho"), 2, 0)
        source_combo = QComboBox()
        source_candidates = reserved_sources if reserved_sources else candidates
        for row in source_candidates:
            source_combo.addItem(
                f"{row.get('material_id', '-') } | {row.get('lote', '-') } | {row.get('dimensao', '-')}",
                str(row.get("material_id", "") or "").strip(),
            )
        retalho_layout.addWidget(source_combo, 2, 1, 1, 3)
        layout.addWidget(retalho_card)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        retalho = {
            "comprimento": comp_spin.value(),
            "largura": larg_spin.value(),
            "quantidade": qtd_retalho_spin.value(),
            "metros": metros_spin.value(),
        }
        has_retalho = any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
        if skip_box.isChecked():
            return {
                "material_id": "",
                "quantity": 0.0,
                "allow_without_stock": True,
                "retalho": retalho if has_retalho else {},
                "source_material_id": str(source_combo.currentData() or "").strip() if has_retalho else "",
            }
        current = dict(material_combo.currentData() or {})
        return {
            "material_id": str(current.get("material_id", "") or "").strip(),
            "quantity": qty_spin.value(),
            "allow_without_stock": False,
            "retalho": retalho if has_retalho else {},
            "source_material_id": str(source_combo.currentData() or "").strip() if has_retalho else "",
        }

    def _prompt_supervisor_password(self) -> bool:
        dialog = QDialog(self)
        dialog.setWindowTitle("Autorizacao superior")
        dialog.setMinimumWidth(360)
        layout = QVBoxLayout(dialog)
        info = QLabel("Introduz a password de supervisor para autorizar a baixa manual.")
        info.setWordWrap(True)
        layout.addWidget(info)
        password_edit = QLineEdit()
        password_edit.setEchoMode(QLineEdit.Password)
        password_edit.setPlaceholderText("Password supervisor")
        layout.addWidget(password_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return False
        verify = getattr(self.backend, "verify_supervisor_password", None)
        if not callable(verify) or not verify(password_edit.text()):
            QMessageBox.critical(self, "Dar Baixa", "Password de supervisor invalida.")
            return False
        return True

    def _prompt_manual_material_consumption(self, material: str, espessura: str) -> dict[str, Any] | None:
        candidates_fn = getattr(self.backend, "material_candidates", None)
        if not callable(candidates_fn):
            return None
        candidates = list(candidates_fn(material, espessura) or [])
        if not candidates:
            QMessageBox.information(self, "Dar Baixa", f"Sem stock disponivel para {material} {espessura} mm.")
            return None
        dialog = QDialog(self)
        dialog.setWindowTitle("Dar Baixa")
        dialog.resize(860, 560)
        layout = QVBoxLayout(dialog)
        info = QLabel(
            f"Baixa manual para {material} {espessura} mm.\n"
            "Seleciona as quantidades por lote. Se criares retalho, associa-o ao lote certo."
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        table = QTableWidget(len(candidates), 6)
        table.setHorizontalHeaderLabels(["Dimensao", "Disponivel", "Local", "Lote", "Peso/Un.", "Baixar"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.resizeSection(5, 122)
        spinners: list[tuple[dict[str, Any], QDoubleSpinBox]] = []
        for row_index, row in enumerate(candidates):
            table.setItem(row_index, 0, QTableWidgetItem(str(row.get("dimensao", "-"))))
            table.setItem(row_index, 1, QTableWidgetItem(f"{float(row.get('disponivel', 0) or 0):.2f}"))
            table.setItem(row_index, 2, QTableWidgetItem(str(row.get("local", "-"))))
            table.setItem(row_index, 3, QTableWidgetItem(str(row.get("lote", "-"))))
            table.setItem(row_index, 4, QTableWidgetItem(f"{float(row.get('peso_unid', 0) or 0):.3f} kg"))
            spin = QDoubleSpinBox()
            spin.setRange(0.0, float(row.get("disponivel", 0) or 0))
            spin.setDecimals(2)
            spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
            spin.setAlignment(Qt.AlignCenter)
            table.setCellWidget(row_index, 5, spin)
            spinners.append((row, spin))
        table.setMinimumHeight(_table_visible_height(table, max(6, len(candidates)), extra=22))
        layout.addWidget(table, 1)

        retalho_card = CardFrame()
        retalho_card.set_tone("warning")
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(12, 10, 12, 10)
        retalho_layout.setHorizontalSpacing(8)
        retalho_layout.setVerticalSpacing(6)
        retalho_layout.addWidget(QLabel("Retalho comprimento"), 0, 0)
        retalho_layout.addWidget(QLabel("Retalho largura"), 0, 1)
        retalho_layout.addWidget(QLabel("Qtd retalho"), 0, 2)
        retalho_layout.addWidget(QLabel("Metros"), 0, 3)
        comp_spin = QDoubleSpinBox()
        larg_spin = QDoubleSpinBox()
        qtd_retalho_spin = QDoubleSpinBox()
        metros_spin = QDoubleSpinBox()
        for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(2)
            spin.setAlignment(Qt.AlignCenter)
        retalho_layout.addWidget(comp_spin, 1, 0)
        retalho_layout.addWidget(larg_spin, 1, 1)
        retalho_layout.addWidget(qtd_retalho_spin, 1, 2)
        retalho_layout.addWidget(metros_spin, 1, 3)
        retalho_layout.addWidget(QLabel("Lote origem do retalho"), 2, 0)
        source_combo = QComboBox()
        for row in candidates:
            source_combo.addItem(
                f"{row.get('material_id', '-') } | {row.get('lote', '-') } | {row.get('dimensao', '-')}",
                str(row.get("material_id", "") or "").strip(),
            )
        retalho_layout.addWidget(source_combo, 2, 1, 1, 3)
        layout.addWidget(retalho_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        allocations: list[dict[str, Any]] = []
        selected_ids: set[str] = set()
        for row, spin in spinners:
            value = spin.value()
            if value <= 0:
                continue
            material_id = str(row.get("material_id", "") or "").strip()
            allocations.append({"material_id": material_id, "quantidade": value})
            selected_ids.add(material_id)
        if not allocations:
            QMessageBox.warning(self, "Dar Baixa", "Define pelo menos uma quantidade para baixa.")
            return None
        retalho = {
            "comprimento": comp_spin.value(),
            "largura": larg_spin.value(),
            "quantidade": qtd_retalho_spin.value(),
            "metros": metros_spin.value(),
        }
        has_retalho = any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
        source_material_id = str(source_combo.currentData() or "").strip()
        if has_retalho and len(selected_ids) > 1 and source_material_id not in selected_ids:
            QMessageBox.warning(self, "Dar Baixa", "Seleciona como origem do retalho um dos lotes efetivamente baixados.")
            return None
        return {
            "allocations": allocations,
            "retalho": retalho if has_retalho else {},
            "source_material_id": source_material_id if has_retalho else "",
        }

    def _manual_consume_material(self) -> None:
        group = self._current_group()
        material = str(group.get("material", "") or "").strip()
        espessura = str(group.get("espessura", "") or "").strip()
        if not material or not espessura:
            QMessageBox.warning(self, "Dar Baixa", "Seleciona primeiro um grupo com material e espessura.")
            return
        if not self._prompt_supervisor_password():
            return
        payload = self._prompt_manual_material_consumption(material, espessura)
        if payload is None:
            return
        consume_fn = getattr(self.backend, "consume_material_allocations", None)
        if not callable(consume_fn):
            QMessageBox.critical(self, "Dar Baixa", "Backend sem suporte para baixa manual.")
            return
        try:
            result = dict(
                consume_fn(
                    payload.get("allocations", []),
                    retalho=payload.get("retalho", {}),
                    source_material_id=str(payload.get("source_material_id", "") or "").strip(),
                    reason=f"operador_{str(group.get('encomenda', '') or '').strip()}_{material}_{espessura}",
                )
                or {}
            )
        except Exception as exc:
            QMessageBox.critical(self, "Dar Baixa", str(exc))
            return
        message = f"Baixa registada: {float(result.get('consumed_total', 0) or 0):.2f}"
        if str(result.get("retalho_id", "") or "").strip():
            message = f"{message} | Retalho {result.get('retalho_id')}"
        self._set_feedback(message, error=False)
        self.refresh()

    def _maybe_resolve_laser_stock(self, enc_num: str, material: str, espessura: str) -> bool:
        status_fn = getattr(self.backend, "operator_laser_stock_state", None)
        resolve_fn = getattr(self.backend, "operator_resolve_laser_stock", None)
        if not callable(status_fn) or not callable(resolve_fn):
            return True
        try:
            stock_state = dict(status_fn(enc_num, material, espessura) or {})
        except Exception as exc:
            self._set_feedback(str(exc), error=True)
            QMessageBox.critical(self, "Baixa material", str(exc))
            return False
        if not stock_state or not stock_state.get("laser_complete") or stock_state.get("resolved"):
            return True
        payload = self._prompt_laser_stock_resolution(stock_state)
        if payload is None:
            self._set_feedback("Baixa material pendente para a ultima peca Laser.", error=True)
            return False
        try:
            result = dict(
                resolve_fn(
                    enc_num,
                    material,
                    espessura,
                    material_id=str(payload.get("material_id", "") or "").strip(),
                    quantidade=float(payload.get("quantity", 0) or 0),
                    allow_without_stock=bool(payload.get("allow_without_stock")),
                    retalho=payload.get("retalho", {}),
                    source_material_id=str(payload.get("source_material_id", "") or "").strip(),
                )
                or {}
            )
        except Exception as exc:
            self._set_feedback(str(exc), error=True)
            QMessageBox.critical(self, "Baixa material", str(exc))
            return False
        consumed = float(result.get("consumed_total", 0) or 0)
        remaining = float(result.get("remaining_qty", 0) or 0)
        if remaining > 0 and bool(result.get("allow_without_stock")):
            self._set_feedback(f"Laser concluido. Baixa pendente assumida: {remaining:.1f}.", error=False)
        elif consumed > 0:
            message = f"Baixa material registada: {consumed:.1f}."
            if str(result.get("retalho_id", "") or "").strip():
                message = f"{message} Retalho {result.get('retalho_id')}."
            self._set_feedback(message, error=False)
        elif str(result.get("retalho_id", "") or "").strip():
            self._set_feedback(f"Retalho registado: {result.get('retalho_id')}.", error=False)
        return True

    def _run_action(self, action_fn, success_text: str, *, allow_multiple: bool = True) -> None:
        refs_list = self._selected_piece_refs(allow_multiple=allow_multiple)
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        for enc_num, piece_id in refs_list:
            try:
                action_fn(enc_num, piece_id)
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if applied > 1:
            self._set_feedback(f"{success_text} ({applied} pecas)", error=False)
        else:
            self._set_feedback(success_text, error=False)

    def _start_piece(self) -> None:
        operator_name = self._current_operator()
        if not operator_name:
            QMessageBox.warning(self, "Operador", "Seleciona o operador antes de iniciar.")
            return
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        collected = self._collect_start_operation_targets(refs_list)
        ordered_ops = list(collected.get("ordered_ops", []) or [])
        targets = dict(collected.get("targets", {}) or {})
        skipped_without_ops = int(collected.get("skipped_without_ops", 0) or 0)
        context_errors = [str(err) for err in list(collected.get("errors", []) or []) if str(err).strip()]
        if not ordered_ops:
            message = context_errors[0] if context_errors else "As pecas selecionadas nao têm operacoes pendentes para iniciar."
            self._set_feedback(message, error=True)
            QMessageBox.warning(self, "Operador", message)
            return
        operation = self._prompt_start_operation(refs_list, ordered_ops, targets, skipped_without_ops=skipped_without_ops)
        if not operation:
            return
        self.operation_combo.setCurrentText(operation)
        compatible_refs = list(targets.get(operation, []) or [])
        skipped_for_choice = max(0, len(refs_list) - len(compatible_refs))
        if not compatible_refs:
            QMessageBox.warning(self, "Operador", "Nenhuma das pecas selecionadas tem essa operacao pendente.")
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = compatible_refs[0][1]
        errors = list(context_errors)
        applied = 0
        for enc_num, piece_id in compatible_refs:
            try:
                self.backend.operator_start_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    operation=operation,
                    posto=self._current_posto(),
                )
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            if applied:
                message = f"Operacao iniciada em {applied} peca(s), mas houve falhas:\n{message}"
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        success = f"Operacao iniciada: {operation}"
        if applied > 1:
            success = f"{success} ({applied} pecas)"
        if skipped_for_choice:
            success = f"{success} | Ignoradas {skipped_for_choice} sem esta operacao"
        self._set_feedback(success, error=False)

    def _finish_piece(self) -> None:
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        operator_name = self._current_operator()
        self.selected_group_key = self._group_key(self._current_group())
        errors: list[str] = []
        completed = 0
        finished_ops: list[str] = []
        batch_total = len(refs_list)
        current_group = self._current_group()
        group_material = str(current_group.get("material", "") or "").strip()
        group_esp = str(current_group.get("espessura", "") or "").strip()
        group_enc = str(current_group.get("encomenda", "") or "").strip()
        for batch_idx, (enc_num, piece_id) in enumerate(refs_list, start=1):
            try:
                ctx = self.backend.operator_piece_context(enc_num, piece_id)
            except Exception as exc:
                errors.append(str(exc))
                continue
            payload = self._prompt_finish(
                ctx,
                batch_idx=batch_idx if batch_total > 1 else 0,
                batch_total=batch_total if batch_total > 1 else 0,
                preferred_operation=self._current_operation(),
            )
            if payload is None:
                break
            self.selected_piece_id = piece_id
            try:
                result = self.backend.operator_finish_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    payload["ok"],
                    payload["nok"],
                    payload["qual"],
                    operation=str(payload["operation"] or ""),
                    posto=self._current_posto(),
                )
                completed += 1
                op_name = str((result or {}).get("operation", "") or payload.get("operation", "") or "").strip()
                if op_name:
                    finished_ops.append(op_name)
            except Exception as exc:
                errors.append(str(exc))
        if completed and group_enc and any("laser" in str(op or "").lower() for op in finished_ops):
            self._maybe_resolve_laser_stock(group_enc, group_material, group_esp)
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if completed > 1:
            self._set_feedback(f"Operacoes concluidas com sucesso ({completed} pecas).", error=False)
        elif completed == 1:
            self._set_feedback("Operacao concluida com sucesso.", error=False)

    def _resume_piece(self) -> None:
        operator_name = self._current_operator()
        self._run_action(
            lambda enc_num, piece_id: self.backend.operator_resume_piece(enc_num, piece_id, operator_name, posto=self._current_posto()),
            "Peca retomada.",
        )

    def _pause_piece(self) -> None:
        reason = self._prompt_reason(
            "Interromper peca",
            "Seleciona ou escreve o motivo da interrupcao.",
            self.backend.operator_interruption_options(),
        )
        if reason is None:
            return
        operator_name = self._current_operator()
        self._run_action(
            lambda enc_num, piece_id: self.backend.operator_pause_piece(enc_num, piece_id, operator_name, reason, posto=self._current_posto()),
            f"Interrupcao registada: {reason}",
        )

    def _register_avaria(self) -> None:
        reason = self._prompt_reason(
            "Registar avaria",
            "Seleciona ou escreve a causa da avaria.",
            self.backend.operator_avaria_options(),
        )
        if reason is None:
            return
        operator_name = self._current_operator()
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        shared_group_id = ""
        shared_started_at = ""
        for enc_num, piece_id in refs_list:
            try:
                result = self.backend.operator_register_avaria(
                    enc_num,
                    piece_id,
                    operator_name,
                    reason,
                    posto=self._current_posto(),
                    group_id=shared_group_id,
                    ts_now=shared_started_at,
                )
                if not shared_group_id:
                    shared_group_id = str((result or {}).get("avaria_group_key", "") or "").strip()
                if not shared_started_at:
                    shared_started_at = str((result or {}).get("avaria_started_at", "") or "").strip()
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if applied > 1:
            self._set_feedback(f"Avaria aberta: {reason} ({applied} pecas)", error=False)
        else:
            self._set_feedback(f"Avaria aberta: {reason}", error=False)

    def _close_avaria(self) -> None:
        operator_name = self._current_operator()
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        group_minutes: dict[str, float] = {}
        for enc_num, piece_id in refs_list:
            try:
                result = self.backend.operator_close_avaria(enc_num, piece_id, operator_name, posto=self._current_posto())
                group_key = str((result or {}).get("avaria_group_key", "") or "").strip() or piece_id
                group_minutes[group_key] = max(
                    group_minutes.get(group_key, 0.0),
                    float((result or {}).get("duracao_avaria_min", 0) or 0),
                )
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        total_minutes = sum(group_minutes.values())
        if applied > 1:
            self._set_feedback(f"Avaria fechada ({applied} pecas) | Tempo de paragem {total_minutes:.1f} min", error=False)
        else:
            self._set_feedback(f"Avaria fechada. Tempo {total_minutes:.1f} min", error=False)

    def _alert_chefia(self) -> None:
        operator_name = self._current_operator()
        self._run_action(
            lambda enc_num, piece_id: self.backend.operator_alert_chefia(enc_num, piece_id, operator_name, posto=self._current_posto()),
            "Poke enviado para a chefia.",
        )

    def _open_drawing(self) -> None:
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        opened = 0
        last_drawing = ""
        for enc_num, piece_id in refs_list:
            try:
                last_drawing = str(self.backend.operator_open_drawing(enc_num, piece_id))
                opened += 1
            except Exception as exc:
                self._set_feedback(str(exc), error=True)
                QMessageBox.critical(self, "Ver desenho", str(exc))
                return
        if opened > 1:
            self._set_feedback(f"Desenhos abertos: {opened}", error=False)
        else:
            drawing_name = Path(str(last_drawing or "")).name or str(last_drawing or "")
            self._set_feedback(f"Desenho aberto: {drawing_name}", error=False)
            self.feedback_label.setToolTip(str(last_drawing or ""))

    def _open_labels_dialog(self) -> None:
        group = self._current_group()
        enc_num = str(group.get("encomenda", "") or "").strip()
        if not enc_num:
            QMessageBox.warning(self, "Etiquetas", "Seleciona primeiro uma encomenda com pecas.")
            return
        dialog = _OperatorLabelsDialog(
            self.backend,
            enc_num,
            current_posto=self._current_posto(),
            preselected_ids=self._selected_piece_ids(),
            parent=self,
        )
        dialog.exec()


class _OperatorLabelsDialog(QDialog):
    def __init__(self, backend, order_number: str, current_posto: str = "Geral", preselected_ids: list[str] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.order_number = str(order_number or "").strip()
        self.selected_ids: set[str] = {str(value or "").strip() for value in list(preselected_ids or []) if str(value or "").strip()}
        self.rows: list[dict] = []
        self.filtered_rows: list[dict] = []
        self._syncing_checks = False

        self.setWindowTitle(f"Etiquetas | {self.order_number or '-'}")
        self.resize(1180, 760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro_card = CardFrame()
        intro_card.set_tone("info")
        intro_layout = QVBoxLayout(intro_card)
        intro_layout.setContentsMargins(14, 12, 14, 12)
        intro_layout.setSpacing(4)
        title = QLabel("Etiquetas do Operador")
        title.setStyleSheet("font-size: 17px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel(
            "Seleciona as referencias da encomenda para imprimir etiqueta por unidade 110x50 ou etiqueta de palete A4. "
            "A etiqueta de palete agrupa automaticamente por proximo posto."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        intro_layout.addWidget(title)
        intro_layout.addWidget(subtitle)
        layout.addWidget(intro_card)

        filters_card = CardFrame()
        filters_card.set_tone("default")
        filters_layout = QGridLayout(filters_card)
        filters_layout.setContentsMargins(12, 10, 12, 10)
        filters_layout.setHorizontalSpacing(8)
        filters_layout.setVerticalSpacing(6)
        filters_layout.addWidget(QLabel("Posto de origem"), 0, 0)
        self.source_posto_combo = QComboBox()
        self.source_posto_combo.setProperty("compact", "true")
        self.source_posto_combo.addItems(list(self.backend.available_postos() or ["Geral"]))
        self.source_posto_combo.setCurrentText(str(current_posto or "").strip() or "Geral")
        self.source_posto_combo.currentTextChanged.connect(self._reload_rows)
        filters_layout.addWidget(self.source_posto_combo, 1, 0)
        filters_layout.addWidget(QLabel("Filtrar"), 0, 1)
        self.search_edit = QLineEdit()
        self.search_edit.setProperty("compact", "true")
        self.search_edit.setPlaceholderText("Pesquisar ref., OPP, descricao ou posto...")
        self.search_edit.textChanged.connect(self._render_rows)
        filters_layout.addWidget(self.search_edit, 1, 1)
        filters_layout.addWidget(QLabel("Proximo posto"), 0, 2)
        self.destination_combo = QComboBox()
        self.destination_combo.setProperty("compact", "true")
        self.destination_combo.currentTextChanged.connect(self._render_rows)
        filters_layout.addWidget(self.destination_combo, 1, 2)
        _cap_width(self.source_posto_combo, 170)
        _cap_width(self.destination_combo, 180)
        layout.addWidget(filters_card)

        selection_row = QHBoxLayout()
        selection_row.setSpacing(8)
        self.select_visible_btn = QPushButton("Selecionar visiveis")
        self.select_visible_btn.setProperty("variant", "secondary")
        self.select_visible_btn.clicked.connect(lambda: self._apply_visible_selection(True))
        self.clear_visible_btn = QPushButton("Limpar visiveis")
        self.clear_visible_btn.setProperty("variant", "secondary")
        self.clear_visible_btn.clicked.connect(lambda: self._apply_visible_selection(False))
        self.selection_chip = QLabel("0 selecionadas")
        _apply_state_chip(self.selection_chip, "-", "0 selecionadas")
        selection_row.addWidget(self.select_visible_btn)
        selection_row.addWidget(self.clear_visible_btn)
        selection_row.addStretch(1)
        selection_row.addWidget(self.selection_chip)
        layout.addLayout(selection_row)

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(["Sel.", "Ref. Int.", "Ref. Ext.", "Descricao", "OPP", "Qtd", "Estado", "Origem", "Proximo posto"])
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        _configure_table(self.table, stretch=(3, 8), contents=(0, 1, 4, 5, 6, 7))
        header = self.table.horizontalHeader()
        for col, width in ((0, 34), (1, 128), (2, 118), (3, 320), (4, 112), (5, 62), (6, 96), (7, 108), (8, 150)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        self.table.itemChanged.connect(self._handle_check_change)
        layout.addWidget(self.table, 1)

        actions_host = QWidget()
        actions_layout = QVBoxLayout(actions_host)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)

        unit_card = CardFrame()
        unit_card.set_tone("info")
        unit_layout = QHBoxLayout(unit_card)
        unit_layout.setContentsMargins(12, 10, 12, 10)
        unit_layout.setSpacing(8)
        unit_info = QLabel("Etiqueta por unidade 110x50")
        unit_info.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
        unit_hint = QLabel("Uma etiqueta por OPP selecionada.")
        unit_hint.setProperty("role", "muted")
        unit_block = QVBoxLayout()
        unit_block.setSpacing(2)
        unit_block.addWidget(unit_info)
        unit_block.addWidget(unit_hint)
        self.preview_unit_btn = QPushButton("Pre-visualizar")
        self.print_unit_btn = QPushButton("Imprimir")
        self.save_unit_btn = QPushButton("Guardar PDF")
        self.preview_unit_btn.setProperty("variant", "secondary")
        self.print_unit_btn.setProperty("variant", "secondary")
        self.save_unit_btn.setProperty("variant", "secondary")
        self.preview_unit_btn.clicked.connect(self._preview_unit_labels)
        self.print_unit_btn.clicked.connect(self._print_unit_labels)
        self.save_unit_btn.clicked.connect(self._save_unit_labels)
        unit_layout.addLayout(unit_block, 1)
        unit_layout.addWidget(self.preview_unit_btn)
        unit_layout.addWidget(self.print_unit_btn)
        unit_layout.addWidget(self.save_unit_btn)
        actions_layout.addWidget(unit_card)

        pallet_card = CardFrame()
        pallet_card.set_tone("default")
        pallet_layout = QHBoxLayout(pallet_card)
        pallet_layout.setContentsMargins(12, 10, 12, 10)
        pallet_layout.setSpacing(8)
        pallet_info = QLabel("Etiqueta de palete A4")
        pallet_info.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
        pallet_hint = QLabel("Agrupa as referencias selecionadas por proximo posto no mesmo PDF.")
        pallet_hint.setProperty("role", "muted")
        pallet_block = QVBoxLayout()
        pallet_block.setSpacing(2)
        pallet_block.addWidget(pallet_info)
        pallet_block.addWidget(pallet_hint)
        self.preview_pallet_btn = QPushButton("Pre-visualizar")
        self.print_pallet_btn = QPushButton("Imprimir")
        self.save_pallet_btn = QPushButton("Guardar PDF")
        self.preview_pallet_btn.setProperty("variant", "secondary")
        self.print_pallet_btn.setProperty("variant", "secondary")
        self.save_pallet_btn.setProperty("variant", "secondary")
        self.preview_pallet_btn.clicked.connect(self._preview_pallet_labels)
        self.print_pallet_btn.clicked.connect(self._print_pallet_labels)
        self.save_pallet_btn.clicked.connect(self._save_pallet_labels)
        pallet_layout.addLayout(pallet_block, 1)
        pallet_layout.addWidget(self.preview_pallet_btn)
        pallet_layout.addWidget(self.print_pallet_btn)
        pallet_layout.addWidget(self.save_pallet_btn)
        actions_layout.addWidget(pallet_card)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        actions_layout.addLayout(close_row)
        layout.addWidget(actions_host)

        self._reload_rows()

    def _selected_piece_ids(self) -> list[str]:
        ordered_ids: list[str] = []
        seen: set[str] = set()
        for row in self.rows:
            piece_id = str(row.get("piece_id", "") or "").strip()
            if piece_id and piece_id in self.selected_ids and piece_id not in seen:
                ordered_ids.append(piece_id)
                seen.add(piece_id)
        return ordered_ids

    def _reload_rows(self) -> None:
        try:
            payload = dict(self.backend.operator_label_rows(self.order_number, source_posto=self.source_posto_combo.currentText().strip()) or {})
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))
            self.rows = []
            self.filtered_rows = []
            self.table.setRowCount(0)
            self._sync_selection_state()
            return
        self.rows = list(payload.get("rows", []) or [])
        valid_ids = {str(row.get("piece_id", "") or "").strip() for row in self.rows}
        self.selected_ids.intersection_update(valid_ids)
        current_filter = self.destination_combo.currentText().strip()
        destinations = ["Todos"] + sorted({str(row.get("proximo_posto", "") or "-").strip() or "-" for row in self.rows})
        self.destination_combo.blockSignals(True)
        self.destination_combo.clear()
        self.destination_combo.addItems(destinations)
        if current_filter and current_filter in destinations:
            self.destination_combo.setCurrentText(current_filter)
        self.destination_combo.blockSignals(False)
        self._render_rows()

    def _render_rows(self) -> None:
        query = self.search_edit.text().strip().lower()
        destination = self.destination_combo.currentText().strip()
        rows = []
        for row in self.rows:
            if destination and destination != "Todos" and str(row.get("proximo_posto", "") or "-").strip() != destination:
                continue
            if query:
                haystack = " ".join(
                    [
                        str(row.get("ref_interna", "") or ""),
                        str(row.get("ref_externa", "") or ""),
                        str(row.get("descricao", "") or ""),
                        str(row.get("opp", "") or ""),
                        str(row.get("posto_origem", "") or ""),
                        str(row.get("proximo_posto", "") or ""),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            rows.append(row)
        self.filtered_rows = rows
        self._syncing_checks = True
        try:
            self.table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                piece_id = str(row.get("piece_id", "") or "").strip()
                check_item = QTableWidgetItem("")
                check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                check_item.setCheckState(Qt.Checked if piece_id in self.selected_ids else Qt.Unchecked)
                check_item.setData(Qt.UserRole, piece_id)
                check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, 0, check_item)
                values = [
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("descricao", "") or "-",
                    row.get("opp", "-"),
                    row.get("quantidade_txt", "0"),
                    row.get("estado", "-"),
                    row.get("posto_origem", "-"),
                    row.get("proximo_posto", "-"),
                ]
                for column, value in enumerate(values, start=1):
                    item = QTableWidgetItem(str(value))
                    if column in (5, 6, 7, 8):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    item.setToolTip(str(value))
                    self.table.setItem(row_index, column, item)
                _paint_table_row(self.table, row_index, str(row.get("estado", "") or ""))
        finally:
            self._syncing_checks = False
        self._sync_selection_state()

    def _handle_check_change(self, item: QTableWidgetItem) -> None:
        if self._syncing_checks or item.column() != 0:
            return
        piece_id = str(item.data(Qt.UserRole) or "").strip()
        if not piece_id:
            return
        if item.checkState() == Qt.Checked:
            self.selected_ids.add(piece_id)
        else:
            self.selected_ids.discard(piece_id)
        self._sync_selection_state()

    def _apply_visible_selection(self, selected: bool) -> None:
        self._syncing_checks = True
        try:
            for row_index, row in enumerate(self.filtered_rows):
                piece_id = str(row.get("piece_id", "") or "").strip()
                if not piece_id:
                    continue
                if selected:
                    self.selected_ids.add(piece_id)
                else:
                    self.selected_ids.discard(piece_id)
                item = self.table.item(row_index, 0)
                if item is not None:
                    item.setCheckState(Qt.Checked if selected else Qt.Unchecked)
        finally:
            self._syncing_checks = False
        self._sync_selection_state()

    def _sync_selection_state(self) -> None:
        selected_count = len(self._selected_piece_ids())
        _apply_state_chip(self.selection_chip, "Em producao" if selected_count else "-", f"{selected_count} selecionadas")
        enabled = selected_count > 0
        for button in (
            self.preview_unit_btn,
            self.print_unit_btn,
            self.save_unit_btn,
            self.preview_pallet_btn,
            self.print_pallet_btn,
            self.save_pallet_btn,
        ):
            button.setEnabled(enabled)

    def _build_pdf(self, kind: str, output_path: str | None = None):
        piece_ids = self._selected_piece_ids()
        if not piece_ids:
            raise ValueError("Seleciona pelo menos uma referencia.")
        source_posto = self.source_posto_combo.currentText().strip() or "Geral"
        if kind == "unit":
            return self.backend.operator_unit_labels_pdf(self.order_number, piece_ids, source_posto=source_posto, output_path=output_path)
        return self.backend.operator_pallet_labels_pdf(self.order_number, piece_ids, source_posto=source_posto, output_path=output_path)

    def _preview_pdf(self, kind: str) -> None:
        try:
            path = self._build_pdf(kind)
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))

    def _print_pdf(self, kind: str) -> None:
        try:
            path = self._build_pdf(kind)
            try:
                os.startfile(str(path), "print")
            except Exception:
                os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))

    def _save_pdf(self, kind: str) -> None:
        default_name = f"etiquetas_{self.order_number}_{'unit' if kind == 'unit' else 'palete'}.pdf"
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", default_name, "PDF (*.pdf)")
        if not path:
            return
        try:
            self._build_pdf(kind, output_path=path)
        except Exception as exc:
            QMessageBox.critical(self, "Guardar PDF", str(exc))
            return
        QMessageBox.information(self, "Guardar PDF", f"PDF guardado em:\n{path}")

    def _preview_unit_labels(self) -> None:
        self._preview_pdf("unit")

    def _print_unit_labels(self) -> None:
        self._print_pdf("unit")

    def _save_unit_labels(self) -> None:
        self._save_pdf("unit")

    def _preview_pallet_labels(self) -> None:
        self._preview_pdf("pallet")

    def _print_pallet_labels(self) -> None:
        self._print_pdf("pallet")

    def _save_pallet_labels(self) -> None:
        self._save_pdf("pallet")


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

        root = QVBoxLayout(self)
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
        self.order_btn.setProperty("variant", "secondary")
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
        self._sync_buttons()

    def refresh(self) -> None:
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
        main_window = self.window()
        if hasattr(main_window, "show_page"):
            try:
                main_window.show_page("orders")
                page = getattr(main_window, "pages", {}).get("orders")
                if page is not None and hasattr(page, "open_order_numero"):
                    page.open_order_numero(numero)
            except Exception as exc:
                QMessageBox.critical(self, "Abrir encomenda", str(exc))


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


class PlanningPage(QWidget):
    page_title = "Planeamento"
    page_subtitle = "Semana operacional em grelha, com blocos coloridos e backlog real."
    uses_backend_reload = True

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
        self.current_operation = "Corte Laser"
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        top_card = CardFrame()
        top_card.set_tone("info")
        top_card.setMinimumHeight(118)
        top_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 12, 14, 12)
        top_layout.setSpacing(8)
        week_row = QHBoxLayout()
        week_row.setSpacing(6)
        self.week_label = QLabel("-")
        self.week_label.setStyleSheet("font-size: 15px; font-weight: 900; color: #0f172a;")
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
        self.auto_plan_btn = QPushButton("Auto planear")
        self.auto_plan_btn.clicked.connect(self._auto_plan)
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
        self.laser_deadline_btn = QPushButton("Prazo laser")
        self.laser_deadline_btn.setProperty("variant", "secondary")
        self.laser_deadline_btn.clicked.connect(self._show_laser_deadlines_dialog)
        self.blocked_btn = QPushButton("Bloqueios")
        self.blocked_btn.setProperty("variant", "secondary")
        self.blocked_btn.clicked.connect(self._show_blocked_windows_dialog)
        self.pdf_btn = QPushButton("PDF")
        self.pdf_btn.clicked.connect(self._open_pdf)
        for button in (
            self.prev_week_btn,
            self.next_week_btn,
            self.current_week_btn,
            self.refresh_btn,
            self.auto_plan_btn,
            self.clear_week_btn,
            self.move_earlier_btn,
            self.move_later_btn,
            self.move_prev_day_btn,
            self.move_next_day_btn,
            self.remove_block_btn,
            self.view_blocks_btn,
            self.view_history_btn,
            self.laser_deadline_btn,
            self.blocked_btn,
            self.pdf_btn,
        ):
            button.setProperty("compact", "true")
            button.setMinimumHeight(24)
            week_row.addWidget(button)
        top_layout.addLayout(week_row)
        operations_host = QWidget()
        operations_host.setMinimumHeight(30)
        operations_row = QHBoxLayout(operations_host)
        operations_row.setContentsMargins(0, 0, 0, 0)
        operations_row.setSpacing(6)
        operations_label = QLabel("Operacao")
        operations_label.setProperty("role", "muted")
        operations_label.setMinimumHeight(18)
        operations_row.addWidget(operations_label)
        self.operation_buttons: dict[str, QPushButton] = {}
        operation_options = list(self.backend.planning_operation_options() if self.backend is not None and hasattr(self.backend, "planning_operation_options") else [])
        if not operation_options:
            operation_options = ["Corte Laser", "Quinagem", "Serralharia", "Maquinacao", "Roscagem", "Lacagem", "Montagem"]
        for op_name in operation_options:
            button = QPushButton(str(op_name))
            button.setCheckable(True)
            button.setProperty("compact", "true")
            button.setMinimumHeight(24)
            button.clicked.connect(lambda checked=False, value=op_name: self._set_operation(value))
            self.operation_buttons[str(op_name)] = button
            operations_row.addWidget(button)
        operations_row.addStretch(1)
        top_layout.addWidget(operations_host)
        meta_host = QWidget()
        meta_host.setMinimumHeight(22)
        meta_row = QHBoxLayout(meta_host)
        meta_row.setContentsMargins(0, 0, 0, 0)
        meta_row.setSpacing(10)
        self.period_meta = QLabel("Periodo -")
        self.period_meta.setProperty("role", "muted")
        self.period_meta.setStyleSheet("font-size: 11px;")
        self.period_meta.setMinimumHeight(18)
        self.status_meta = QLabel("Blocos 0 | Backlog 0")
        self.status_meta.setProperty("role", "muted")
        self.status_meta.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_meta.setStyleSheet("font-size: 11px;")
        self.status_meta.setMinimumHeight(18)
        meta_row.addWidget(self.period_meta, 1)
        meta_row.addWidget(self.status_meta, 0)
        top_layout.addWidget(meta_host)
        root.addWidget(top_card)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        self.cards = [StatCard(title) for title in ("Blocos ativos", "Encomendas", "Carga semanal", "Blocos fechados")]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, index)
            card.setMaximumHeight(84)
            card.title_label.setStyleSheet("font-size: 10px;")
            card.value_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
            card.subtitle_label.setStyleSheet("font-size: 10px;")
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("warning")
        self.cards[2].set_tone("success")
        self.cards[3].set_tone("default")
        root.addWidget(cards_host)

        main_split = QSplitter(Qt.Horizontal)
        main_split.setChildrenCollapsible(False)

        backlog_card = CardFrame()
        backlog_card.set_tone("warning")
        backlog_layout = QVBoxLayout(backlog_card)
        backlog_layout.setContentsMargins(16, 14, 16, 14)
        self.backlog_title = QLabel("Produção / Montagem")
        self.backlog_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.backlog_table = PlanningBacklogTable(self)
        self.backlog_table.setColumnCount(6)
        self.backlog_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Material", "Esp.", "Tempo", "Obs."])
        self.backlog_table.verticalHeader().setVisible(False)
        self.backlog_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.backlog_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.backlog_table.verticalHeader().setDefaultSectionSize(20)
        self.backlog_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.backlog_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.backlog_table.setStyleSheet("font-size: 11px;")
        self.backlog_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.backlog_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.backlog_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.backlog_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        backlog_layout.addWidget(self.backlog_title)
        backlog_layout.addWidget(self.backlog_table)
        main_split.addWidget(backlog_card)

        grid_card = CardFrame()
        grid_card.set_tone("info")
        grid_layout = QVBoxLayout(grid_card)
        grid_layout.setContentsMargins(16, 14, 16, 14)
        self.grid_title = QLabel("Quadro semanal")
        self.grid_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.grid = PlanningGridTable(self)
        self.grid.setEditTriggers(QTableWidget.NoEditTriggers)
        self.grid.setSelectionMode(QTableWidget.NoSelection)
        self.grid.verticalHeader().setVisible(False)
        self.grid.verticalHeader().setDefaultSectionSize(20)
        self.grid.verticalHeader().setMinimumSectionSize(18)
        self.grid.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.grid.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.grid.setStyleSheet("font-size: 9px;")
        grid_layout.addWidget(self.grid_title)
        grid_layout.addWidget(self.grid)
        grid_card.setMaximumHeight(620)
        main_split.addWidget(grid_card)
        main_split.setSizes([430, 1110])
        root.addWidget(main_split, 1)

        self.active_card = CardFrame(self)
        self.active_card.set_tone("info")
        active_layout = QVBoxLayout(self.active_card)
        active_layout.setContentsMargins(16, 14, 16, 14)
        self.active_title = QLabel("Blocos da semana")
        self.active_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.active_table = QTableWidget(0, 7)
        self.active_table.setHorizontalHeaderLabels(["Dia", "Inicio", "Duracao", "Encomenda", "Material", "Esp.", "Chapa"])
        self.active_table.verticalHeader().setVisible(False)
        self.active_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.active_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.active_table.verticalHeader().setDefaultSectionSize(20)
        self.active_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.active_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.active_table.setStyleSheet("font-size: 11px;")
        self.active_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.active_table.itemSelectionChanged.connect(self._sync_planning_actions)
        active_layout.addWidget(self.active_title)
        active_layout.addWidget(self.active_table)
        self.history_card = CardFrame(self)
        self.history_card.set_tone("default")
        history_layout = QVBoxLayout(self.history_card)
        history_layout.setContentsMargins(16, 14, 16, 14)
        self.history_title = QLabel("Histórico de planeamento")
        self.history_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.history_table = QTableWidget(0, 7)
        self.history_table.setHorizontalHeaderLabels(["Data", "Inicio", "Encomenda", "Material", "Esp.", "Planeado", "Real / Estado"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.verticalHeader().setDefaultSectionSize(20)
        self.history_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.history_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.history_table.setStyleSheet("font-size: 11px;")
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        history_layout.addWidget(self.history_title)
        history_layout.addWidget(self.history_table)
        self.active_card.hide()
        self.history_card.hide()
        self._apply_operation_labels()
        self._fit_planning_grid()
        self._sync_planning_actions()

    def refresh(self) -> None:
        self._apply_operation_labels()
        data = self.runtime_service.planning_overview(
            week_start=self.week_start.isoformat(),
            operation=self.current_operation,
        )
        summary = data.get("summary", {})
        self.current_week_dates = list(data.get("week_dates", []) or [])
        self.current_blocked_windows = list(self.backend.planning_blocked_windows() if self.backend is not None else [])
        self.week_label.setText(f"Semana {summary.get('week_label', '-')} | {self.current_operation}")
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
                    r.get("chapa", "-"),
                ]
                for r in self.current_active
            ],
            align_center_from=1,
        )
        for row_index, row in enumerate(self.current_active):
            self._paint_row_with_block_color(self.active_table, row_index, str(row.get("color", "") or "#dbeafe"))
        if self.backend is not None:
            self.current_pending = list(self.backend.planning_pending_rows(operation=self.current_operation))
            _fill_table(
                self.backlog_table,
                [
                    [
                        r.get("numero", "-"),
                        r.get("cliente", "-"),
                        r.get("material", "-"),
                        r.get("espessura", "-"),
                        f"{r.get('tempo_min', 0):.1f}",
                        r.get("obs", "-"),
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
        self.cards[2].set_data(f"{summary.get('min_ativos', 0):.0f} min", f"Carga {self.current_operation}")
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
        if selected == self.current_operation:
            self._apply_operation_labels()
            return
        self.current_operation = selected
        self.refresh()

    def _apply_operation_labels(self) -> None:
        current = str(self.current_operation or "Corte Laser").strip() or "Corte Laser"
        for op_name, button in self.operation_buttons.items():
            button.setChecked(op_name == current)
            button.setProperty("variant", "primary" if op_name == current else "secondary")
            button.style().unpolish(button)
            button.style().polish(button)
        self.backlog_title.setText(f"Pendentes - {current}")
        self.grid_title.setText(f"Quadro semanal - {current}")
        self.active_title.setText(f"Blocos - {current}")
        self.history_title.setText(f"Historico - {current}")

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
            text = f"{block.get('encomenda', '-')}\n{material_txt} | {block.get('espessura', '-')}mm"
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
                f"Duracao: {duration:.0f} min"
            )
            self.grid.setItem(row, col, item)
            if span > 1:
                self.grid.setSpan(row, col, span, 1)

    def _fit_planning_grid(self) -> None:
        if not hasattr(self, "grid") or self.grid is None:
            return
        header = self.grid.horizontalHeader()
        header.setMinimumHeight(28)
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
            return {}
        return self.current_active[current.row()]

    def _sync_planning_actions(self) -> None:
        has_backend = self.backend is not None
        has_block = bool(self._selected_active_row())
        self.auto_plan_btn.setEnabled(has_backend and bool(self.current_pending))
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

    def _plan_order_dialog(self, rows: list[dict]) -> list[dict] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Ordem de planeamento - {self.current_operation}")
        dialog.resize(880, 520)
        layout = QVBoxLayout(dialog)
        info = QLabel("Seleciona os pendentes pela ordem desejada. A lista da direita pode ser arrastada para reordenar.")
        info.setProperty("role", "muted")
        layout.addWidget(info)
        body = QHBoxLayout()
        left = QListWidget()
        right = QListWidget()
        right.setDragDropMode(QAbstractItemView.InternalMove)
        right.setDefaultDropAction(Qt.MoveAction)
        label_map: dict[str, dict] = {}
        for row in rows:
            label = f"{row.get('numero', '-')} | {row.get('material', '-')} | {row.get('espessura', '-')} | {row.get('tempo_min', 0):.0f} min"
            label_map[label] = row
            left.addItem(label)
        buttons = QVBoxLayout()
        add_btn = QPushButton("Adicionar ->")
        remove_btn = QPushButton("<- Remover")
        add_all_btn = QPushButton("Tudo ->")
        add_btn.clicked.connect(lambda: self._move_plan_items(left, right))
        remove_btn.clicked.connect(lambda: self._move_plan_items(right, left))
        add_all_btn.clicked.connect(lambda: self._move_all_plan_items(left, right))
        for button in (add_btn, remove_btn, add_all_btn):
            buttons.addWidget(button)
        buttons.addStretch(1)
        body.addWidget(left, 1)
        body.addLayout(buttons)
        body.addWidget(right, 1)
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
            f"Blocos da semana - {self.current_operation}",
            ["Dia", "Inicio", "Duracao", "Encomenda", "Material", "Esp.", "Chapa"],
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    f"{r.get('duracao_min', 0):.1f}",
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
                    r.get("chapa", "-"),
                ]
                for r in self.current_active
            ],
            [str(r.get("estado_final", "") or "") for r in self.current_active],
        )

    def _show_history_dialog(self) -> None:
        self._show_table_dialog(
            f"Histórico de planeamento - {self.current_operation}",
            ["Data", "Inicio", "Encomenda", "Material", "Esp.", "Planeado", "Real / Estado"],
            [
                [
                    self._fmt_day(r.get("data", "-")),
                    r.get("inicio", "-"),
                    r.get("encomenda", "-"),
                    r.get("material", "-"),
                    r.get("espessura", "-"),
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
        dialog.setWindowTitle("Prazos Laser")
        dialog.resize(1180, 560)
        layout = QVBoxLayout(dialog)
        info = QLabel(
            "Prazo previsto de conclusão do corte laser por encomenda, com base nos blocos atualmente planeados. "
            "Se uma encomenda estiver parcial, o prazo final ainda não está totalmente fechado."
        )
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Entrega", "Grupos", "Planeado", "Fim laser", "Estado"])
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
            "Planeado completo": "Concluida",
            "Planeado parcial": "Incompleta",
            "Por planear": "Pendente",
            "Laser concluído": "Concluida",
        }
        for idx, row in enumerate(rows):
            _paint_table_row(table, idx, tone_map.get(str(row.get("estado", "") or ""), "Pendente"))
            item = table.item(idx, 0)
            if item is not None:
                item.setToolTip(str(row.get("materiais_txt", "") or ""))
        layout.addWidget(table, 1)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        preview_btn = QPushButton("Abrir PDF")
        preview_btn.setProperty("variant", "secondary")
        save_btn = QPushButton("Guardar PDF")
        save_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.reject)
        btn_row.addWidget(preview_btn)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        def open_pdf() -> None:
            try:
                path = self.backend.planning_open_laser_deadlines_pdf()
            except Exception as exc:
                QMessageBox.critical(dialog, "Prazos Laser", str(exc))
                return
            QMessageBox.information(dialog, "Prazos Laser", f"PDF aberto:\n{path}")

        def save_pdf() -> None:
            path, _ = QFileDialog.getSaveFileName(dialog, "Guardar PDF", "prazos_laser_planeamento.pdf", "PDF (*.pdf)")
            if not path:
                return
            try:
                self.backend.planning_render_laser_deadlines_pdf(path)
            except Exception as exc:
                QMessageBox.critical(dialog, "Prazos Laser", str(exc))
                return
            QMessageBox.information(dialog, "Prazos Laser", f"PDF guardado em:\n{path}")

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
        self.refresh()

    def _open_pdf(self) -> None:
        try:
            if self.backend is not None:
                path = self.backend.planning_open_pdf(self.week_start, operation=self.current_operation)
            else:
                path = self.runtime_service.planning_pdf(week_start=self.week_start.isoformat(), operation=self.current_operation)
        except Exception as exc:
            QMessageBox.critical(self, "Planeamento", str(exc))
            return
        QMessageBox.information(self, "Planeamento", f"PDF aberto:\n{path}")


class ExpeditionPage(QWidget):
    page_title = "Expedição"
    page_subtitle = "Expedição operacional com peças disponíveis, guia em preparação e histórico de ações."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.pending_rows: list[dict] = []
        self.piece_rows: list[dict] = []
        self.history_rows: list[dict] = []
        self.current_guide_lines: list[dict] = []
        self.draft_rows: list[dict] = []
        self.current_pending_num = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        actions = CardFrame()
        actions.set_tone("info")
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(14, 10, 14, 10)
        actions_layout.setSpacing(8)
        self.emit_off_btn = QPushButton("Emitir Guia OFF")
        self.emit_off_btn.clicked.connect(self._emit_off_guide)
        self.manual_guide_btn = QPushButton("Criar Guia Manual")
        self.manual_guide_btn.setProperty("variant", "secondary")
        self.manual_guide_btn.clicked.connect(self._create_manual_guide)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.edit_guide_btn = QPushButton("Editar Guia")
        self.edit_guide_btn.setProperty("variant", "secondary")
        self.edit_guide_btn.clicked.connect(self._edit_guide)
        self.cancel_guide_btn = QPushButton("Anular Guia")
        self.cancel_guide_btn.setProperty("variant", "danger")
        self.cancel_guide_btn.clicked.connect(self._cancel_guide)
        self.save_pdf_btn = QPushButton("Guardar PDF")
        self.save_pdf_btn.setProperty("variant", "secondary")
        self.save_pdf_btn.clicked.connect(self._save_guide_pdf)
        self.history_dialog_btn = QPushButton("Histórico Guias")
        self.history_dialog_btn.setProperty("variant", "secondary")
        self.history_dialog_btn.clicked.connect(self._show_history_dialog)
        for button in (self.emit_off_btn, self.manual_guide_btn, self.refresh_btn, self.edit_guide_btn, self.cancel_guide_btn, self.preview_pdf_btn if hasattr(self, "preview_pdf_btn") else None, self.save_pdf_btn, self.history_dialog_btn):
            if button is not None:
                actions_layout.addWidget(button)
        actions_layout.addStretch(1)
        root.addWidget(actions)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por encomenda, cliente, guia ou matrícula")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.estado_combo = QComboBox()
        self.estado_combo.addItems(["Todas", "Não expedida", "Parcialmente expedida", "Totalmente expedida"])
        self.estado_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.estado_combo)
        root.addWidget(filters)

        pending_card = CardFrame()
        pending_card.set_tone("warning")
        pending_layout = QVBoxLayout(pending_card)
        pending_layout.setContentsMargins(14, 12, 14, 12)
        pending_title = QLabel("Encomendas com material para expedir")
        pending_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.pending_table = QTableWidget(0, 6)
        self.pending_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Estado Prod.", "Estado Exp.", "Disponível", "Entrega"])
        self.pending_table.verticalHeader().setVisible(False)
        self.pending_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pending_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pending_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pending_table.verticalHeader().setDefaultSectionSize(24)
        self.pending_table.itemSelectionChanged.connect(self._on_pending_selected)
        pending_layout.addWidget(pending_title)
        pending_layout.addWidget(self.pending_table)
        self.pending_empty_label = QLabel("Sem quantidades prontas para expedição neste momento.")
        self.pending_empty_label.setProperty("role", "muted")
        self.pending_empty_label.setVisible(False)
        pending_layout.addWidget(self.pending_empty_label)
        root.addWidget(pending_card)

        pieces_card = CardFrame()
        pieces_card.set_tone("success")
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(14, 12, 14, 12)
        pieces_header = QHBoxLayout()
        pieces_title = QLabel("Peças disponíveis para expedição")
        pieces_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.open_piece_drawing_btn = QPushButton("Ver desenho da peça")
        self.open_piece_drawing_btn.setProperty("variant", "secondary")
        self.open_piece_drawing_btn.clicked.connect(self._open_piece_drawing)
        pieces_header.addWidget(pieces_title, 1)
        pieces_header.addWidget(self.open_piece_drawing_btn)
        self.pieces_table = QTableWidget(0, 8)
        self.pieces_table.setHorizontalHeaderLabels(["ID", "Ref. Int.", "Ref. Ext.", "Estado", "Pronta", "Expedida", "Disponível", "Desenho"])
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.pieces_table, stretch=(2,), contents=())
        _set_table_columns(
            self.pieces_table,
            [
                (0, "interactive", 120),
                (1, "interactive", 190),
                (2, "stretch", 320),
                (3, "interactive", 136),
                (4, "interactive", 82),
                (5, "interactive", 88),
                (6, "interactive", 96),
                (7, "interactive", 82),
            ],
        )
        self.pieces_table.verticalHeader().setDefaultSectionSize(24)
        self.pieces_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.pieces_table)
        root.addWidget(pieces_card)

        draft_card = CardFrame()
        draft_card.set_tone("info")
        draft_layout = QVBoxLayout(draft_card)
        draft_layout.setContentsMargins(14, 12, 14, 12)
        draft_layout.setSpacing(8)
        draft_header = QHBoxLayout()
        draft_title = QLabel("Guia em preparação")
        draft_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.draft_context = QLabel("Seleciona uma encomenda e uma peça para começar a guia.")
        self.draft_context.setProperty("role", "muted")
        draft_header.addWidget(draft_title)
        draft_header.addStretch(1)
        draft_header.addWidget(self.draft_context)
        draft_layout.addLayout(draft_header)
        draft_actions = QHBoxLayout()
        self.exp_qty_spin = QDoubleSpinBox()
        self.exp_qty_spin.setRange(0.01, 1000000.0)
        self.exp_qty_spin.setDecimals(2)
        self.exp_qty_spin.setValue(1.0)
        self.add_draft_btn = QPushButton("Adicionar linha")
        self.add_draft_btn.clicked.connect(self._add_draft_line)
        self.remove_draft_btn = QPushButton("Remover linha")
        self.remove_draft_btn.setProperty("variant", "secondary")
        self.remove_draft_btn.clicked.connect(self._remove_draft_line)
        self.clear_draft_btn = QPushButton("Limpar guia")
        self.clear_draft_btn.setProperty("variant", "secondary")
        self.clear_draft_btn.clicked.connect(self._clear_draft)
        draft_actions.addWidget(QLabel("Qtd"))
        draft_actions.addWidget(self.exp_qty_spin)
        draft_actions.addWidget(self.add_draft_btn)
        draft_actions.addWidget(self.remove_draft_btn)
        draft_actions.addWidget(self.clear_draft_btn)
        draft_actions.addStretch(1)
        draft_layout.addLayout(draft_actions)
        self.draft_table = QTableWidget(0, 5)
        self.draft_table.setHorizontalHeaderLabels(["Peça", "Ref. Int.", "Ref. Ext.", "Descrição", "Qtd"])
        self.draft_table.verticalHeader().setVisible(False)
        self.draft_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.draft_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.draft_table, stretch=(2, 3), contents=())
        _set_table_columns(
            self.draft_table,
            [
                (0, "interactive", 96),
                (1, "interactive", 160),
                (2, "stretch", 210),
                (3, "stretch", 240),
                (4, "interactive", 82),
            ],
        )
        self.draft_table.verticalHeader().setDefaultSectionSize(24)
        self.draft_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        draft_layout.addWidget(self.draft_table)
        root.addWidget(draft_card)

        history_card = CardFrame()
        history_card.set_tone("default")
        history_layout = QVBoxLayout(history_card)
        history_layout.setContentsMargins(14, 12, 14, 12)
        history_header = QHBoxLayout()
        history_title = QLabel("Histórico de guias")
        history_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.preview_pdf_btn = QPushButton("Abrir PDF da guia")
        self.preview_pdf_btn.clicked.connect(self._open_guide_pdf)
        history_header.addWidget(history_title, 1)
        history_header.addWidget(self.preview_pdf_btn)
        self.history_table = QTableWidget(0, 7)
        self.history_table.setHorizontalHeaderLabels(["Guia", "Tipo", "Encomenda", "Cliente", "Emissão", "Estado", "Linhas"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table.verticalHeader().setDefaultSectionSize(24)
        self.history_table.itemSelectionChanged.connect(self._update_guide_detail)
        self.history_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        history_layout.addLayout(history_header)
        history_layout.addWidget(self.history_table)
        self.history_empty_label = QLabel("Sem guias emitidas neste momento.")
        self.history_empty_label.setProperty("role", "muted")
        self.history_empty_label.setVisible(False)
        history_layout.addWidget(self.history_empty_label)
        root.addWidget(history_card)

        self.detail_card = CardFrame()
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(14, 12, 14, 12)
        detail_layout.setSpacing(8)
        detail_header = QHBoxLayout()
        self.detail_title = QLabel("Seleciona uma guia")
        self.detail_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.detail_state_chip = QLabel("-")
        _apply_state_chip(self.detail_state_chip, "-")
        detail_header.addWidget(self.detail_title, 1)
        detail_header.addWidget(self.detail_state_chip, 0, Qt.AlignRight)
        detail_layout.addLayout(detail_header)
        self.detail_meta = QLabel("-")
        self.detail_meta.setProperty("role", "muted")
        self.detail_meta.setWordWrap(True)
        detail_layout.addWidget(self.detail_meta)
        self.detail_note = QLabel("-")
        self.detail_note.setWordWrap(True)
        detail_layout.addWidget(self.detail_note)
        self.lines_table = QTableWidget(0, 7)
        self.lines_table.setHorizontalHeaderLabels(["Ref. Int.", "Ref. Ext.", "Descrição", "Qtd", "Peso", "Manual", "Encomenda"])
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.lines_table, stretch=(1, 2, 6), contents=())
        _set_table_columns(
            self.lines_table,
            [
                (0, "interactive", 170),
                (1, "stretch", 250),
                (2, "stretch", 280),
                (3, "interactive", 82),
                (4, "interactive", 82),
                (5, "interactive", 82),
                (6, "stretch", 170),
            ],
        )
        self.lines_table.verticalHeader().setDefaultSectionSize(24)
        detail_layout.addWidget(self.lines_table)
        self.detail_card.set_tone("default")
        root.addWidget(self.detail_card, 1)
        self._render_draft_table()
        self._sync_expedicao_actions()

    def refresh(self) -> None:
        query = self.filter_edit.currentText().strip()
        previous_pending = self.current_pending_num or self._current_pending_row().get("numero", "")
        previous_guide = self._current_history_row().get("numero", "")
        self.pending_rows = self.backend.expedicao_pending_orders(query, self.estado_combo.currentText())
        self.history_rows = self.backend.expedicao_rows(query)
        self._refresh_filter_options(query)
        _fill_table(
            self.pending_table,
            [[r.get("numero", "-"), r.get("cliente", "-"), r.get("estado", "-"), r.get("estado_expedicao", "-"), f"{r.get('disponivel', 0):.1f}", r.get("data_entrega", "-")] for r in self.pending_rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.pending_rows):
            _paint_table_row(self.pending_table, row_index, str(row.get("estado_expedicao", "")))
        self.pending_empty_label.setVisible(self.pending_table.rowCount() == 0)
        self.pending_empty_label.setText(
            "Sem quantidades prontas para expedição neste momento."
            if not query
            else "Nenhuma encomenda disponível para expedição com o filtro atual."
        )
        _fill_table(
            self.history_table,
            [[r.get("numero", "-"), r.get("tipo", "-"), r.get("encomenda", "-"), r.get("cliente", "-"), r.get("data_emissao", "-"), r.get("estado", "-"), r.get("linhas", 0)] for r in self.history_rows],
            align_center_from=6,
        )
        self.history_empty_label.setVisible(self.history_table.rowCount() == 0)
        self.history_empty_label.setText(
            "Sem guias emitidas neste momento."
            if not query
            else "Nenhuma guia encontrada com o filtro atual."
        )
        for row_index, row in enumerate(self.history_rows):
            _paint_table_row(self.history_table, row_index, str(row.get("estado", "")))
        self._restore_selection(self.pending_table, self.pending_rows, previous_pending)
        self._restore_selection(self.history_table, self.history_rows, previous_guide)
        if self.pending_table.rowCount() == 0:
            self.current_pending_num = ""
            self.piece_rows = []
            self.pieces_table.setRowCount(0)
            self.draft_context.setText("Seleciona uma encomenda e uma peça para começar a guia.")
        else:
            self._on_pending_selected()
        if self.history_table.rowCount() == 0:
            self._clear_guide_detail()
        else:
            self._update_guide_detail()
        self._sync_expedicao_actions()

    def _refresh_filter_options(self, current_text: str) -> None:
        suggestions: list[str] = []
        for row in list(self.pending_rows) + list(self.history_rows):
            for key in ("numero", "cliente", "encomenda", "matricula"):
                value = str(row.get(key, "") or "").strip()
                if value and value not in suggestions:
                    suggestions.append(value)
        block = self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.addItem("")
        for value in suggestions:
            self.filter_edit.addItem(value)
        self.filter_edit.setCurrentText(current_text)
        self.filter_edit.blockSignals(block)

    def _restore_selection(self, table: QTableWidget, rows: list[dict], target_num: str) -> None:
        if table.rowCount() == 0:
            return
        row_index = 0
        if target_num:
            for index, row in enumerate(rows):
                if str(row.get("numero", "")).strip() == str(target_num).strip():
                    row_index = index
                    break
        table.selectRow(row_index)

    def _current_pending_row(self) -> dict:
        current = self.pending_table.currentItem()
        if current is None or current.row() >= len(self.pending_rows):
            return {}
        return self.pending_rows[current.row()]

    def _current_piece_row(self) -> dict:
        current = self.pieces_table.currentItem()
        if current is None or current.row() >= len(self.piece_rows):
            return {}
        return self.piece_rows[current.row()]

    def _current_history_row(self) -> dict:
        current = self.history_table.currentItem()
        if current is None or current.row() >= len(self.history_rows):
            return {}
        return self.history_rows[current.row()]

    def _current_draft_row(self) -> dict:
        current = self.draft_table.currentItem()
        if current is None or current.row() >= len(self.draft_rows):
            return {}
        return self.draft_rows[current.row()]

    def _sync_expedicao_actions(self) -> None:
        has_pending = bool(self._current_pending_row())
        has_piece = bool(self._current_piece_row())
        has_history = bool(self._current_history_row())
        history_state = str(self._current_history_row().get("estado", "") or "").strip().lower()
        has_draft = bool(self.draft_rows)
        self.emit_off_btn.setEnabled(has_pending and has_draft)
        self.manual_guide_btn.setEnabled(True)
        self.edit_guide_btn.setEnabled(has_history and "anulad" not in history_state)
        self.cancel_guide_btn.setEnabled(has_history and "anulad" not in history_state)
        self.preview_pdf_btn.setEnabled(has_history)
        self.save_pdf_btn.setEnabled(has_history)
        self.open_piece_drawing_btn.setEnabled(has_pending and has_piece)
        self.add_draft_btn.setEnabled(has_pending and has_piece)
        self.remove_draft_btn.setEnabled(bool(self._current_draft_row()))
        self.clear_draft_btn.setEnabled(has_draft)

    def _on_pending_selected(self) -> None:
        current = self._current_pending_row()
        enc_num = str(current.get("numero", "") or "").strip()
        if not enc_num:
            self.current_pending_num = ""
            self.piece_rows = []
            self.pieces_table.setRowCount(0)
            self.draft_context.setText("Seleciona uma encomenda e uma peça para começar a guia.")
            self._sync_expedicao_actions()
            return
        if self.current_pending_num and self.current_pending_num != enc_num and self.draft_rows:
            self.draft_rows = []
            self._render_draft_table()
        self.current_pending_num = enc_num
        self.draft_context.setText(f"Encomenda ativa: {enc_num}")
        self.piece_rows = self.backend.expedicao_available_pieces(enc_num) if enc_num else []
        _fill_table(
            self.pieces_table,
            [[r.get("id", "-"), r.get("ref_interna", "-"), r.get("ref_externa", "-"), r.get("estado", "-"), r.get("pronta_expedicao", "0"), r.get("qtd_expedida", "0"), r.get("disponivel", "0"), "SIM" if r.get("desenho") else "NAO"] for r in self.piece_rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.piece_rows):
            for col_index in range(self.pieces_table.columnCount()):
                item = self.pieces_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
            _paint_table_row(self.pieces_table, row_index, str(row.get("estado", "")))
        if self.pieces_table.rowCount() > 0:
            self.pieces_table.selectRow(0)
        self._sync_expedicao_actions()

    def _clear_guide_detail(self) -> None:
        self.detail_title.setText("Seleciona uma guia")
        _apply_state_chip(self.detail_state_chip, "-")
        self.detail_meta.setText("-")
        self.detail_note.setText("Sem detalhe disponível para o filtro atual.")
        self.current_guide_lines = []
        self.lines_table.setRowCount(0)
        _set_panel_tone(self.detail_card, "default")
        self._sync_expedicao_actions()

    def _update_guide_detail(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            self._clear_guide_detail()
            return
        try:
            detail = self.backend.expedicao_detail(numero)
        except Exception as exc:
            self._clear_guide_detail()
            self.detail_note.setText(str(exc))
            return
        state = str(detail.get("estado", "") or "-").strip() or "-"
        self.detail_title.setText(f"Guia {detail.get('numero', '-')}")
        _apply_state_chip(self.detail_state_chip, state)
        self.detail_meta.setText(
            f"Enc {detail.get('encomenda', '-')} | Cliente {detail.get('cliente', '-')}"
            f" | Emissão {detail.get('data_emissao', '-') or '-'} | Transporte {detail.get('data_transporte', '-') or '-'}"
        )
        self.detail_note.setText(
            f"Destino: {detail.get('destinatario', '-') or '-'} | Descarga: {detail.get('local_descarga', '-') or '-'}"
            f" | Transportador: {detail.get('transportador', '-') or '-'} | Matrícula: {detail.get('matricula', '-') or '-'}"
            f" | Obs: {detail.get('observacoes', '-') or '-'}"
        )
        self.current_guide_lines = list(detail.get("lines", []) or [])
        _fill_table(
            self.lines_table,
            [[r.get("ref_interna", "-"), r.get("ref_externa", "-"), r.get("descricao", "-"), r.get("qtd", "0"), r.get("peso", "0"), "SIM" if r.get("manual") else "NAO", r.get("encomenda", "-")] for r in self.current_guide_lines],
            align_center_from=3,
        )
        for row_index in range(self.lines_table.rowCount()):
            for col_index in range(self.lines_table.columnCount()):
                item = self.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        _set_panel_tone(self.detail_card, _state_tone(state))
        self._sync_expedicao_actions()

    def _render_draft_table(self) -> None:
        _fill_table(
            self.draft_table,
            [
                [
                    row.get("peca_id", "-"),
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("descricao", "-"),
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                ]
                for row in self.draft_rows
            ],
            align_center_from=4,
        )
        for row_index in range(self.draft_table.rowCount()):
            for col_index in range(self.draft_table.columnCount()):
                item = self.draft_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        self._sync_expedicao_actions()

    def _select_history(self, numero: str) -> None:
        numero = str(numero or "").strip()
        if not numero:
            return
        for index, row in enumerate(self.history_rows):
            if str(row.get("numero", "") or "").strip() == numero:
                self.history_table.selectRow(index)
                break

    def _add_draft_line(self) -> None:
        order = self._current_pending_row()
        piece = self._current_piece_row()
        enc_num = str(order.get("numero", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        if not enc_num or not piece_id:
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda e uma peça disponível.")
            return
        qty = float(self.exp_qty_spin.value() or 0)
        if qty <= 0:
            QMessageBox.warning(self, "Expedição", "Quantidade inválida.")
            return
        used = sum(float(row.get("qtd", 0) or 0) for row in self.draft_rows if str(row.get("peca_id", "") or "").strip() == piece_id)
        available = float(piece.get("disponivel_num", 0) or 0)
        if qty > available - used + 1e-9:
            QMessageBox.warning(self, "Expedição", f"Quantidade superior ao disponível ({max(0.0, available - used):.2f}).")
            return
        merged = False
        for row in self.draft_rows:
            if str(row.get("peca_id", "") or "").strip() == piece_id:
                row["qtd"] = float(row.get("qtd", 0) or 0) + qty
                merged = True
                break
        if not merged:
            self.draft_rows.append(
                {
                    "encomenda": enc_num,
                    "peca_id": piece_id,
                    "ref_interna": piece.get("ref_interna", ""),
                    "ref_externa": piece.get("ref_externa", ""),
                    "descricao": piece.get("descricao", "") or piece.get("ref_externa", "") or piece.get("ref_interna", ""),
                    "qtd": qty,
                    "unid": "UN",
                    "peso": 0.0,
                    "manual": False,
                }
            )
        self._render_draft_table()
        self._on_pending_selected()

    def _remove_draft_line(self) -> None:
        current = self.draft_table.currentItem()
        if current is None or current.row() >= len(self.draft_rows):
            return
        del self.draft_rows[current.row()]
        self._render_draft_table()
        self._on_pending_selected()

    def _clear_draft(self) -> None:
        if not self.draft_rows:
            return
        if QMessageBox.question(self, "Expedição", "Limpar todas as linhas da guia em preparação?") != QMessageBox.Yes:
            return
        self.draft_rows = []
        self._render_draft_table()
        self._on_pending_selected()

    def _normalize_datetime_input(self, raw: str) -> str:
        value = str(raw or "").strip()
        if not value:
            return value
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%dT%H:%M:%S")
            except Exception:
                continue
        return value

    def _text_prompt(self, title: str, label: str, initial: str = "") -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        edit = QLineEdit(str(initial or ""))
        form.addRow(label, edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return edit.text().strip()

    def _guide_dialog(self, title: str, initial: dict, lines: list[dict]) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(900)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        fields = {
            "codigo_at": QLineEdit(str(initial.get("codigo_at", "") or "").strip()),
            "data_transporte": QLineEdit(str(initial.get("data_transporte", "") or "").replace("T", " ").strip()),
            "emitente_nome": QLineEdit(str(initial.get("emitente_nome", "") or "").strip()),
            "emitente_nif": QLineEdit(str(initial.get("emitente_nif", "") or "").strip()),
            "emitente_morada": QLineEdit(str(initial.get("emitente_morada", "") or "").strip()),
            "destinatario": QLineEdit(str(initial.get("destinatario", "") or "").strip()),
            "dest_nif": QLineEdit(str(initial.get("dest_nif", "") or "").strip()),
            "dest_morada": QLineEdit(str(initial.get("dest_morada", "") or "").strip()),
            "local_carga": QLineEdit(str(initial.get("local_carga", "") or "").strip()),
            "local_descarga": QLineEdit(str(initial.get("local_descarga", "") or "").strip()),
            "transportador": QLineEdit(str(initial.get("transportador", "") or "").strip()),
            "matricula": QLineEdit(str(initial.get("matricula", "") or "").strip()),
        }
        observations = QTextEdit()
        observations.setPlainText(str(initial.get("observacoes", "") or "").strip())
        observations.setFixedHeight(84)
        labels = [
            ("Cod. validação AT", "codigo_at", 0, 0),
            ("Início transporte", "data_transporte", 0, 2),
            ("Emitente", "emitente_nome", 1, 0),
            ("NIF emitente", "emitente_nif", 1, 2),
            ("Morada emitente", "emitente_morada", 2, 0),
            ("Destinatário", "destinatario", 3, 0),
            ("NIF destino", "dest_nif", 3, 2),
            ("Morada destino", "dest_morada", 4, 0),
            ("Local carga", "local_carga", 5, 0),
            ("Local descarga", "local_descarga", 5, 2),
            ("Transportador", "transportador", 6, 0),
            ("Matrícula", "matricula", 6, 2),
        ]
        for label_text, key, row, col in labels:
            form.addWidget(QLabel(label_text), row, col)
            span = 3 if key in {"emitente_morada", "dest_morada"} else 1
            form.addWidget(fields[key], row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 7, 0)
        form.addWidget(observations, 7, 1, 1, 3)
        layout.addLayout(form)
        preview_title = QLabel("Linhas da guia")
        preview_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        preview_table = QTableWidget(0, 5)
        preview_table.setHorizontalHeaderLabels(["Ref. Int.", "Ref. Ext.", "Descrição", "Qtd", "Unid"])
        preview_table.verticalHeader().setVisible(False)
        preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(preview_table, stretch=(1, 2), contents=())
        _set_table_columns(
            preview_table,
            [
                (0, "interactive", 170),
                (1, "stretch", 240),
                (2, "stretch", 300),
                (3, "interactive", 82),
                (4, "interactive", 76),
            ],
        )
        _fill_table(
            preview_table,
            [
                [row.get("ref_interna", "-"), row.get("ref_externa", "-"), row.get("descricao", "-"), row.get("qtd", "0"), row.get("unid", "UN")]
                for row in list(lines or [])
            ],
            align_center_from=3,
        )
        for row_index in range(preview_table.rowCount()):
            for col_index in range(preview_table.columnCount()):
                item = preview_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        layout.addWidget(preview_title)
        layout.addWidget(preview_table)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        if not fields["destinatario"].text().strip():
            QMessageBox.warning(self, "Expedição", "Indica o destinatário.")
            return None
        return {
            "codigo_at": fields["codigo_at"].text().strip(),
            "tipo_via": "Original",
            "emitente_nome": fields["emitente_nome"].text().strip(),
            "emitente_nif": fields["emitente_nif"].text().strip(),
            "emitente_morada": fields["emitente_morada"].text().strip(),
            "destinatario": fields["destinatario"].text().strip(),
            "dest_nif": fields["dest_nif"].text().strip(),
            "dest_morada": fields["dest_morada"].text().strip(),
            "local_carga": fields["local_carga"].text().strip(),
            "local_descarga": fields["local_descarga"].text().strip(),
            "data_transporte": self._normalize_datetime_input(fields["data_transporte"].text()),
            "transportador": fields["transportador"].text().strip(),
            "matricula": fields["matricula"].text().strip(),
            "observacoes": observations.toPlainText().strip(),
        }

    def _manual_guide_dialog(self) -> dict | None:
        initial = self.backend.expedicao_manual_defaults()
        dialog = QDialog(self)
        dialog.setWindowTitle("Criar Guia Manual")
        dialog.setMinimumWidth(980)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        fields = {
            "codigo_at": QLineEdit(str(initial.get("codigo_at", "") or "").strip()),
            "emitente_nome": QLineEdit(str(initial.get("emitente_nome", "") or "").strip()),
            "destinatario": QLineEdit(""),
            "dest_nif": QLineEdit(""),
            "dest_morada": QLineEdit(""),
            "local_carga": QLineEdit(str(initial.get("local_carga", "") or "").strip()),
            "local_descarga": QLineEdit(""),
            "transportador": QLineEdit(""),
            "matricula": QLineEdit(""),
        }
        observations = QTextEdit()
        observations.setFixedHeight(84)
        meta_labels = [
            ("Cod. validação AT", "codigo_at", 0, 0),
            ("Emitente", "emitente_nome", 0, 2),
            ("Destinatário", "destinatario", 1, 0),
            ("NIF", "dest_nif", 1, 2),
            ("Morada", "dest_morada", 2, 0),
            ("Local carga", "local_carga", 3, 0),
            ("Local descarga", "local_descarga", 3, 2),
            ("Transportador", "transportador", 4, 0),
            ("Matrícula", "matricula", 4, 2),
        ]
        for label_text, key, row, col in meta_labels:
            form.addWidget(QLabel(label_text), row, col)
            span = 3 if key == "dest_morada" else 1
            form.addWidget(fields[key], row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 5, 0)
        form.addWidget(observations, 5, 1, 1, 3)
        layout.addLayout(form)

        product_rows = self.backend.expedicao_product_options("")
        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.addItem("", {})
        for row in product_rows:
            product_combo.addItem(f"{row.get('codigo', '')} - {row.get('descricao', '')}", row)
        desc_edit = QLineEdit()
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(1.0)
        unid_edit = QLineEdit("UN")
        line_items: list[dict] = []
        lines_table = QTableWidget(0, 4)
        lines_table.setHorizontalHeaderLabels(["Código", "Descrição", "Qtd", "Unid"])
        lines_table.verticalHeader().setVisible(False)
        lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        lines_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        def render_manual_lines() -> None:
            _fill_table(
                lines_table,
                [[row.get("produto_codigo", "-"), row.get("descricao", "-"), f"{float(row.get('qtd', 0) or 0):.2f}", row.get("unid", "UN")] for row in line_items],
                align_center_from=2,
            )

        def on_product_change() -> None:
            payload = product_combo.currentData()
            if isinstance(payload, dict):
                desc_edit.setText(str(payload.get("descricao", "") or "").strip())
                unid_edit.setText(str(payload.get("unid", "UN") or "UN").strip() or "UN")

        def add_line() -> None:
            payload = product_combo.currentData()
            product_code = ""
            description = desc_edit.text().strip()
            if isinstance(payload, dict):
                product_code = str(payload.get("codigo", "") or "").strip()
                if not description:
                    description = str(payload.get("descricao", "") or "").strip()
            if not description and not product_code:
                QMessageBox.warning(dialog, "Expedição", "Seleciona um produto ou indica uma descrição.")
                return
            line_items.append(
                {
                    "produto_codigo": product_code,
                    "descricao": description,
                    "qtd": qty_spin.value(),
                    "unid": unid_edit.text().strip() or "UN",
                }
            )
            render_manual_lines()
            qty_spin.setValue(1.0)
            desc_edit.clear()

        def remove_line() -> None:
            current = lines_table.currentItem()
            if current is None or current.row() >= len(line_items):
                return
            del line_items[current.row()]
            render_manual_lines()

        product_combo.currentIndexChanged.connect(lambda _index: on_product_change())
        line_controls = QHBoxLayout()
        line_controls.addWidget(QLabel("Produto"))
        line_controls.addWidget(product_combo, 1)
        line_controls.addWidget(QLabel("Descrição"))
        line_controls.addWidget(desc_edit, 1)
        line_controls.addWidget(QLabel("Qtd"))
        line_controls.addWidget(qty_spin)
        line_controls.addWidget(QLabel("Unid"))
        line_controls.addWidget(unid_edit)
        add_line_btn = QPushButton("Adicionar linha")
        add_line_btn.clicked.connect(add_line)
        remove_line_btn = QPushButton("Remover linha")
        remove_line_btn.setProperty("variant", "secondary")
        remove_line_btn.clicked.connect(remove_line)
        line_controls.addWidget(add_line_btn)
        line_controls.addWidget(remove_line_btn)
        layout.addLayout(line_controls)
        layout.addWidget(lines_table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        if not fields["destinatario"].text().strip():
            QMessageBox.warning(self, "Expedição", "Indica o destinatário.")
            return None
        if not line_items:
            QMessageBox.warning(self, "Expedição", "Adiciona pelo menos uma linha.")
            return None
        return {
            "guide": {
                "codigo_at": fields["codigo_at"].text().strip(),
                "tipo_via": "Original",
                "emitente_nome": fields["emitente_nome"].text().strip(),
                "emitente_nif": str(initial.get("emitente_nif", "") or "").strip(),
                "emitente_morada": str(initial.get("emitente_morada", "") or "").strip(),
                "destinatario": fields["destinatario"].text().strip(),
                "dest_nif": fields["dest_nif"].text().strip(),
                "dest_morada": fields["dest_morada"].text().strip(),
                "local_carga": fields["local_carga"].text().strip(),
                "local_descarga": fields["local_descarga"].text().strip(),
                "data_transporte": self._normalize_datetime_input(str(initial.get("data_transporte", ""))),
                "transportador": fields["transportador"].text().strip(),
                "matricula": fields["matricula"].text().strip(),
                "observacoes": observations.toPlainText().strip(),
            },
            "lines": line_items,
        }

    def _emit_off_guide(self) -> None:
        order = self._current_pending_row()
        enc_num = str(order.get("numero", "") or "").strip()
        if not enc_num:
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda.")
            return
        if not self.draft_rows:
            QMessageBox.warning(self, "Expedição", "Sem linhas na guia.")
            return
        try:
            initial = self.backend.expedicao_defaults_for_order(enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        payload = self._guide_dialog("Confirmar Dados da Guia OFF", initial, self.draft_rows)
        if payload is None:
            return
        try:
            detail = self.backend.expedicao_emit_off(enc_num, self.draft_rows, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Emitir guia", str(exc))
            return
        self.draft_rows = []
        self.refresh()
        self._select_history(detail.get("numero", ""))
        if QMessageBox.question(self, "Expedição", f"Guia emitida: {detail.get('numero', '-')}\n\nAbrir PDF agora?") == QMessageBox.Yes:
            self._open_guide_pdf()

    def _create_manual_guide(self) -> None:
        payload = self._manual_guide_dialog()
        if payload is None:
            return
        try:
            detail = self.backend.expedicao_emit_manual(payload.get("guide", {}), payload.get("lines", []))
        except Exception as exc:
            QMessageBox.critical(self, "Guia manual", str(exc))
            return
        self.refresh()
        self._select_history(detail.get("numero", ""))
        if QMessageBox.question(self, "Expedição", f"Guia manual emitida: {detail.get('numero', '-')}\n\nAbrir PDF agora?") == QMessageBox.Yes:
            self._open_guide_pdf()

    def _edit_guide(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        try:
            detail = self.backend.expedicao_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        payload = self._guide_dialog(f"Editar Guia {numero}", detail, list(detail.get("lines", []) or []))
        if payload is None:
            return
        try:
            self.backend.expedicao_update(numero, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Editar guia", str(exc))
            return
        self.refresh()
        self._select_history(numero)

    def _cancel_guide(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        if QMessageBox.question(self, "Anular Guia", f"Anular guia {numero}?") != QMessageBox.Yes:
            return
        motivo = self._text_prompt("Anular Guia", "Justificação")
        if motivo is None:
            return
        try:
            self.backend.expedicao_cancel(numero, motivo)
        except Exception as exc:
            QMessageBox.critical(self, "Anular guia", str(exc))
            return
        self.refresh()
        self._select_history(numero)

    def _open_guide_pdf(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        try:
            path = self.backend.expedicao_open_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        QMessageBox.information(self, "Expedição", f"PDF aberto:\n{path}")

    def _save_guide_pdf(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", f"guia_{numero}.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            self.backend.expedicao_render_pdf(numero, path)
        except Exception as exc:
            QMessageBox.critical(self, "Guardar PDF", str(exc))
            return
        QMessageBox.information(self, "Guardar PDF", f"PDF guardado em:\n{path}")

    def _show_history_dialog(self) -> None:
        self.refresh()
        dialog = QDialog(self)
        dialog.setWindowTitle("Histórico de guias")
        dialog.resize(1100, 520)
        layout = QVBoxLayout(dialog)
        title = QLabel("Histórico de guias emitidas")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(title)
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["Guia", "Tipo", "Encomenda", "Cliente", "Emissao", "Estado", "Linhas"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(3,), contents=(0, 1, 2, 4, 5, 6))
        _fill_table(
            table,
            [[r.get("numero", "-"), r.get("tipo", "-"), r.get("encomenda", "-"), r.get("cliente", "-"), r.get("data_emissao", "-"), r.get("estado", "-"), r.get("linhas", 0)] for r in self.history_rows],
            align_center_from=6,
        )
        for row_index, row in enumerate(self.history_rows):
            _paint_table_row(table, row_index, str(row.get("estado", "")))
        layout.addWidget(table, 1)
        if not self.history_rows:
            empty = QLabel("Sem guias emitidas para o filtro atual.")
            empty.setProperty("role", "muted")
            layout.addWidget(empty)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _open_piece_drawing(self) -> None:
        order = self._current_pending_row()
        piece = self._current_piece_row()
        enc_num = str(order.get("numero", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        if not enc_num or not piece_id:
            QMessageBox.warning(self, "Expedição", "Seleciona uma peça disponível.")
            return
        try:
            self.backend.operator_open_drawing(enc_num, piece_id)
        except Exception as exc:
            QMessageBox.critical(self, "Ver desenho", str(exc))


class LegacyExpeditionPage(ExpeditionPage):
    page_subtitle = "Expedição por encomenda, com registo da guia apenas no detalhe da encomenda."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        root = self.layout()
        sections = _take_layout_items(root)
        actions_item = sections[0] if len(sections) > 0 else None
        filters_item = sections[1] if len(sections) > 1 else None
        pending_item = sections[2] if len(sections) > 2 else None
        pieces_item = sections[3] if len(sections) > 3 else None
        draft_item = sections[4] if len(sections) > 4 else None
        history_item = sections[5] if len(sections) > 5 else None
        detail_item = sections[6] if len(sections) > 6 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)
        _adopt_layout_item(list_layout, filters_item)
        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(14, 10, 14, 10)
        list_actions_layout.setSpacing(8)
        self.open_pending_btn = QPushButton("Abrir encomenda")
        self.open_pending_btn.clicked.connect(self._open_selected_pending)
        history_btn = QPushButton("Histórico Guias")
        history_btn.setProperty("variant", "secondary")
        history_btn.clicked.connect(self._show_history_dialog)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.refresh)
        list_actions_layout.addWidget(self.open_pending_btn)
        list_actions_layout.addWidget(history_btn)
        list_actions_layout.addWidget(refresh_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        _adopt_layout_item(list_layout, pending_item, 1)
        _adopt_layout_item(list_layout, history_item, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)
        detail_top = CardFrame()
        detail_top.set_tone("default")
        detail_top_layout = QHBoxLayout(detail_top)
        detail_top_layout.setContentsMargins(14, 10, 14, 10)
        detail_top_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a encomendas")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_pending_list)
        focus_block = QVBoxLayout()
        focus_block.setSpacing(2)
        self.pending_focus_label = QLabel("Sem encomenda selecionada")
        self.pending_focus_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.pending_meta_label = QLabel("Seleciona uma encomenda para preparar e emitir a guia.")
        self.pending_meta_label.setProperty("role", "muted")
        self.pending_state_chip = QLabel("-")
        _apply_state_chip(self.pending_state_chip, "-")
        focus_block.addWidget(self.pending_focus_label)
        focus_block.addWidget(self.pending_meta_label)
        detail_top_layout.addWidget(back_btn)
        detail_top_layout.addLayout(focus_block, 1)
        detail_top_layout.addWidget(self.pending_state_chip, 0, Qt.AlignTop)
        detail_layout.addWidget(detail_top)
        _adopt_layout_item(detail_layout, actions_item)
        top_split = QSplitter(Qt.Horizontal)
        top_split.setChildrenCollapsible(False)
        top_left = QWidget()
        top_left_layout = QVBoxLayout(top_left)
        top_left_layout.setContentsMargins(0, 0, 0, 0)
        top_left_layout.setSpacing(10)
        _adopt_layout_item(top_left_layout, pieces_item, 1)
        top_right = QWidget()
        top_right_layout = QVBoxLayout(top_right)
        top_right_layout.setContentsMargins(0, 0, 0, 0)
        top_right_layout.setSpacing(10)
        _adopt_layout_item(top_right_layout, draft_item, 1)
        top_split.addWidget(top_left)
        top_split.addWidget(top_right)
        top_split.setSizes([1120, 780])

        bottom_split = QSplitter(Qt.Horizontal)
        bottom_split.setChildrenCollapsible(False)
        bottom_right = QWidget()
        bottom_right_layout = QVBoxLayout(bottom_right)
        bottom_right_layout.setContentsMargins(0, 0, 0, 0)
        bottom_right_layout.setSpacing(10)
        _adopt_layout_item(bottom_right_layout, detail_item, 1)
        bottom_split.addWidget(bottom_right)
        bottom_split.setSizes([1900])

        vertical_split = QSplitter(Qt.Vertical)
        vertical_split.setChildrenCollapsible(False)
        vertical_split.addWidget(top_split)
        vertical_split.addWidget(bottom_split)
        vertical_split.setSizes([430, 340])
        detail_layout.addWidget(vertical_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.pending_table.itemSelectionChanged.connect(self._sync_pending_focus)
        self.pending_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_pending())
        self._show_pending_list()
        self._sync_pending_focus()

    def refresh(self) -> None:
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(self.current_pending_num)
        ExpeditionPage.refresh(self)
        self._sync_pending_focus()
        if self.pending_table.rowCount() == 0:
            self._show_pending_list()
        elif keep_detail:
            self._show_pending_detail()
        else:
            self._show_pending_list()

    def _show_pending_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_pending_focus()

    def _show_pending_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)
        self._sync_pending_focus()

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _sync_pending_focus(self) -> None:
        row = self._current_pending_row()
        has_pending = bool(row)
        self.open_pending_btn.setEnabled(has_pending)
        if not has_pending:
            self.pending_focus_label.setText("Sem encomenda selecionada")
            self.pending_meta_label.setText("Seleciona uma encomenda para preparar e emitir a guia.")
            _apply_state_chip(self.pending_state_chip, "-")
            return
        estado = str(row.get("estado_expedicao", row.get("estado", "-")) or "-").strip() or "-"
        self.pending_focus_label.setText(f"{row.get('numero', '-')} | {row.get('cliente', '-')}")
        self.pending_meta_label.setText(
            f"Estado producao {row.get('estado', '-')} | Expedição {row.get('estado_expedicao', '-')}"
            f" | Disponivel {float(row.get('disponivel', 0) or 0):.1f} | Entrega {row.get('data_entrega', '-') or '-'}"
        )
        _apply_state_chip(self.pending_state_chip, estado)

    def _on_pending_selected(self) -> None:
        ExpeditionPage._on_pending_selected(self)
        self._sync_pending_focus()

    def _open_selected_pending(self) -> None:
        if not self._current_pending_row():
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda.")
            return
        self._on_pending_selected()
        self._show_pending_detail()


class TransportsPage(QWidget):
    page_title = "Transportes"
    page_subtitle = "Planeamento de viagens e paragens para encomendas a nosso cargo, ligado a Expedição e Encomendas."
    uses_backend_reload = True
    allow_auto_timer_refresh = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.pending_rows: list[dict] = []
        self.trip_rows: list[dict] = []
        self.current_detail: dict[str, Any] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        actions = CardFrame()
        actions.set_tone("info")
        actions_layout = QVBoxLayout(actions)
        actions_layout.setContentsMargins(14, 10, 14, 10)
        actions_layout.setSpacing(8)
        self.new_trip_btn = QPushButton("Nova viagem")
        self.new_trip_btn.clicked.connect(self._new_trip)
        self.assign_btn = QPushButton("Agendar seleção")
        self.assign_btn.clicked.connect(self._assign_selected_orders)
        self.edit_trip_btn = QPushButton("Editar viagem")
        self.edit_trip_btn.setProperty("variant", "secondary")
        self.edit_trip_btn.clicked.connect(self._edit_trip)
        self.remove_trip_btn = QPushButton("Apagar viagem")
        self.remove_trip_btn.setProperty("variant", "danger")
        self.remove_trip_btn.clicked.connect(self._remove_trip)
        self.request_btn = QPushButton("Requisitar transporte")
        self.request_btn.setProperty("variant", "secondary")
        self.request_btn.clicked.connect(self._request_transport)
        self.tariff_btn = QPushButton("Tarifário")
        self.tariff_btn.setProperty("variant", "secondary")
        self.tariff_btn.clicked.connect(self._manage_tariffs)
        self.apply_cost_btn = QPushButton("Aplicar custo")
        self.apply_cost_btn.setProperty("variant", "secondary")
        self.apply_cost_btn.clicked.connect(self._apply_suggested_costs)
        self.trip_status_combo = QComboBox()
        self.trip_status_combo.addItems(["Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        self.trip_status_btn = QPushButton("Estado viagem")
        self.trip_status_btn.setProperty("variant", "secondary")
        self.trip_status_btn.clicked.connect(self._apply_trip_status)
        self.stop_status_combo = QComboBox()
        self.stop_status_combo.addItems(["Planeada", "Carregada", "Entregue", "Incidente"])
        self.stop_status_btn = QPushButton("Estado paragem")
        self.stop_status_btn.setProperty("variant", "secondary")
        self.stop_status_btn.clicked.connect(self._apply_stop_status)
        self.edit_stop_btn = QPushButton("Editar paragem / guia")
        self.edit_stop_btn.setProperty("variant", "secondary")
        self.edit_stop_btn.clicked.connect(self._edit_stop)
        self.stop_up_btn = QPushButton("Subir")
        self.stop_up_btn.setProperty("variant", "secondary")
        self.stop_up_btn.clicked.connect(lambda: self._move_stop(-1))
        self.stop_down_btn = QPushButton("Descer")
        self.stop_down_btn.setProperty("variant", "secondary")
        self.stop_down_btn.clicked.connect(lambda: self._move_stop(1))
        self.remove_stop_btn = QPushButton("Remover paragem")
        self.remove_stop_btn.setProperty("variant", "secondary")
        self.remove_stop_btn.clicked.connect(self._remove_stop)
        self.pdf_btn = QPushButton("Folha de rota PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_trip_pdf)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        for button in (
            self.new_trip_btn,
            self.assign_btn,
            self.edit_trip_btn,
            self.remove_trip_btn,
            self.request_btn,
            self.tariff_btn,
            self.apply_cost_btn,
            self.trip_status_btn,
            self.stop_status_btn,
            self.edit_stop_btn,
            self.stop_up_btn,
            self.stop_down_btn,
            self.remove_stop_btn,
            self.pdf_btn,
            self.refresh_btn,
        ):
            button.setMinimumWidth(146)
        self.trip_status_combo.setMinimumWidth(136)
        self.stop_status_combo.setMinimumWidth(136)

        actions_hint = QLabel("Organiza as viagens em 3 passos: criar, gerir estado da viagem e afinar as paragens.")
        actions_hint.setProperty("role", "muted")
        actions_layout.addWidget(actions_hint)

        action_grid = QGridLayout()
        action_grid.setHorizontalSpacing(8)
        action_grid.setVerticalSpacing(8)
        action_grid.setColumnMinimumWidth(0, 96)
        action_grid.setColumnStretch(8, 1)

        planning_label = QLabel("Planeamento")
        planning_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(planning_label, 0, 0)
        for col, widget in enumerate(
            (
                self.new_trip_btn,
                self.assign_btn,
                self.edit_trip_btn,
                self.remove_trip_btn,
                self.request_btn,
                self.tariff_btn,
                self.apply_cost_btn,
                self.refresh_btn,
            ),
            start=1,
        ):
            action_grid.addWidget(widget, 0, col)

        trip_label = QLabel("Viagem")
        trip_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(trip_label, 1, 0)
        action_grid.addWidget(self.trip_status_btn, 1, 1)
        action_grid.addWidget(self.trip_status_combo, 1, 2)

        stop_label = QLabel("Paragens")
        stop_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(stop_label, 2, 0)
        for col, widget in enumerate(
            (
                self.stop_status_btn,
                self.stop_status_combo,
                self.edit_stop_btn,
                self.stop_up_btn,
                self.stop_down_btn,
                self.remove_stop_btn,
                self.pdf_btn,
            ),
            start=1,
        ):
            action_grid.addWidget(widget, 2, col)

        actions_layout.addLayout(action_grid)
        root.addWidget(actions)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por viagem, encomenda, cliente, matrícula ou motorista")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Todas", "Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.state_combo)
        root.addWidget(filters)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        pending_card = CardFrame()
        pending_card.set_tone("warning")
        pending_layout = QVBoxLayout(pending_card)
        pending_layout.setContentsMargins(14, 12, 14, 12)
        pending_layout.setSpacing(8)
        pending_title = QLabel("Encomendas prontas para transporte ou subcontrato")
        pending_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        pending_hint = QLabel("Seleciona aqui as encomendas disponíveis para criar ou completar uma viagem.")
        pending_hint.setProperty("role", "muted")
        pending_hint.setWordWrap(True)
        self.pending_table = QTableWidget(0, 10)
        self.pending_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Entrega", "Zona", "Tipo", "Pal", "Peso kg", "Descarga", "Guia", "Disponível"])
        self.pending_table.verticalHeader().setVisible(False)
        self.pending_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pending_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pending_table.setSelectionMode(QAbstractItemView.MultiSelection)
        _configure_table(self.pending_table, stretch=(1, 3), contents=())
        _set_table_columns(
            self.pending_table,
            [
                (0, "interactive", 132),
                (1, "stretch", 190),
                (2, "interactive", 96),
                (3, "interactive", 140),
                (4, "interactive", 150),
                (5, "interactive", 64),
                (6, "interactive", 80),
                (7, "stretch", 220),
                (8, "interactive", 112),
                (9, "interactive", 86),
            ],
        )
        self.pending_table.verticalHeader().setDefaultSectionSize(30)
        self.pending_table.itemSelectionChanged.connect(self._sync_actions)
        self.pending_empty = QLabel("Sem encomendas a nosso cargo prontas para agendar neste momento.")
        self.pending_empty.setProperty("role", "muted")
        self.pending_empty.setVisible(False)
        pending_layout.addWidget(pending_title)
        pending_layout.addWidget(pending_hint)
        pending_layout.addWidget(self.pending_table)
        pending_layout.addWidget(self.pending_empty)
        splitter.addWidget(pending_card)

        trips_card = CardFrame()
        trips_card.set_tone("default")
        trips_layout = QVBoxLayout(trips_card)
        trips_layout.setContentsMargins(14, 12, 14, 12)
        trips_layout.setSpacing(8)
        trips_title = QLabel("Viagens agendadas")
        trips_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        trips_hint = QLabel("Abre, revê e atualiza o estado das viagens já criadas.")
        trips_hint.setProperty("role", "muted")
        trips_hint.setWordWrap(True)
        self.trip_table = QTableWidget(0, 9)
        self.trip_table.setHorizontalHeaderLabels(["Viagem", "Data", "Saída", "Tipo", "Estado", "Pedido", "Parceiro / Viatura", "Pal", "Paragens"])
        self.trip_table.verticalHeader().setVisible(False)
        self.trip_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.trip_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.trip_table, stretch=(4, 5), contents=())
        _set_table_columns(
            self.trip_table,
            [
                (0, "interactive", 120),
                (1, "interactive", 96),
                (2, "interactive", 86),
                (3, "interactive", 128),
                (4, "interactive", 110),
                (5, "interactive", 118),
                (6, "stretch", 210),
                (7, "interactive", 64),
                (8, "interactive", 82),
            ],
        )
        self.trip_table.verticalHeader().setDefaultSectionSize(30)
        self.trip_table.itemSelectionChanged.connect(self._show_trip_detail)
        self.trip_table.itemSelectionChanged.connect(self._sync_actions)
        self.trip_empty = QLabel("Sem viagens registadas para o filtro atual.")
        self.trip_empty.setProperty("role", "muted")
        self.trip_empty.setVisible(False)
        trips_layout.addWidget(trips_title)
        trips_layout.addWidget(trips_hint)
        trips_layout.addWidget(self.trip_table)
        trips_layout.addWidget(self.trip_empty)
        splitter.addWidget(trips_card)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        self.detail_card = CardFrame()
        self.detail_card.set_tone("info")
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(14, 12, 14, 12)
        detail_layout.setSpacing(8)
        header = QHBoxLayout()
        self.detail_title = QLabel("Seleciona uma viagem")
        self.detail_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.detail_state_chip = QLabel("-")
        _apply_state_chip(self.detail_state_chip, "-")
        header.addWidget(self.detail_title, 1)
        header.addWidget(self.detail_state_chip)
        detail_layout.addLayout(header)
        self.detail_meta = QLabel("Cria uma viagem e afeta encomendas a nosso cargo.")
        self.detail_meta.setWordWrap(True)
        self.detail_meta.setProperty("role", "muted")
        detail_layout.addWidget(self.detail_meta)
        self.detail_note = QLabel("-")
        self.detail_note.setWordWrap(True)
        detail_layout.addWidget(self.detail_note)
        self.stops_table = QTableWidget(0, 13)
        self.stops_table.setHorizontalHeaderLabels(["Ord", "Encomenda", "Cliente", "Zona", "Pal", "Peso", "Vol", "Descarga", "Planeado", "Guia", "Checklist", "POD", "Estado"])
        self.stops_table.verticalHeader().setVisible(False)
        self.stops_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.stops_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.stops_table, stretch=(2, 3), contents=())
        _set_table_columns(
            self.stops_table,
            [
                (0, "interactive", 54),
                (1, "interactive", 126),
                (2, "stretch", 170),
                (3, "interactive", 120),
                (4, "interactive", 54),
                (5, "interactive", 78),
                (6, "interactive", 70),
                (7, "stretch", 220),
                (8, "interactive", 126),
                (9, "interactive", 96),
                (10, "interactive", 86),
                (11, "interactive", 96),
                (12, "interactive", 104),
            ],
        )
        self.stops_table.verticalHeader().setDefaultSectionSize(30)
        self.stops_table.itemSelectionChanged.connect(self._sync_actions)
        self.stops_table.itemDoubleClicked.connect(lambda *_args: self._edit_stop())
        detail_layout.addWidget(self.stops_table, 1)
        right_layout.addWidget(self.detail_card, 1)

        splitter.addWidget(right_panel)
        splitter.setSizes([520, 420, 880])
        self._clear_trip_detail()
        self._sync_actions()

    def refresh(self) -> None:
        query = self.filter_edit.currentText().strip()
        previous_trip = str(self.current_detail.get("numero", "") or "").strip()
        self.pending_rows = self.backend.transport_pending_orders(query)
        self.trip_rows = self.backend.transport_rows(query, self.state_combo.currentText())
        self._refresh_filter_options(query)
        _fill_table(
            self.pending_table,
            [
                [
                    row.get("numero", "-"),
                    row.get("cliente", "-"),
                    row.get("data_entrega", "-"),
                    row.get("zona_transporte", "-"),
                    row.get("nota_transporte", "-"),
                    f"{float(row.get('paletes', 0) or 0):.2f}",
                    f"{float(row.get('peso_bruto_kg', 0) or 0):.1f}",
                    row.get("local_descarga", "-"),
                    row.get("guia_numero", "-"),
                    f"{float(row.get('disponivel', 0) or 0):.1f}",
                ]
                for row in self.pending_rows
            ],
            align_center_from=2,
        )
        self.pending_empty.setVisible(self.pending_table.rowCount() == 0)
        _fill_table(
            self.trip_table,
            [
                [
                    row.get("numero", "-"),
                    row.get("data_planeada", "-"),
                    row.get("hora_saida", "-"),
                    row.get("tipo_responsavel", "-"),
                    row.get("estado", "-"),
                    row.get("pedido_transporte_estado", "-"),
                    row.get("transportadora_nome", "-") if "subcontrat" in str(row.get("tipo_responsavel", "")).lower() else row.get("viatura", "-"),
                    f"{float(row.get('paletes', 0) or 0):.2f}",
                    row.get("paragens", 0),
                ]
                for row in self.trip_rows
            ],
            align_center_from=6,
        )
        for row_index, row in enumerate(self.trip_rows):
            _paint_table_row(self.trip_table, row_index, str(row.get("estado", "")))
        self.trip_empty.setVisible(self.trip_table.rowCount() == 0)
        self._restore_trip_selection(previous_trip)
        if self.trip_table.rowCount() == 0:
            self._clear_trip_detail()
        else:
            self._show_trip_detail()
        self._sync_actions()

    def _refresh_filter_options(self, current_text: str) -> None:
        values: list[str] = []
        for row in list(self.pending_rows) + list(self.trip_rows):
            for key in ("numero", "cliente", "local_descarga", "zona_transporte", "viatura", "motorista", "matricula", "transportadora_nome"):
                value = str(row.get(key, "") or "").strip()
                if value and value not in values:
                    values.append(value)
        block = self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.addItem("")
        for value in values:
            self.filter_edit.addItem(value)
        self.filter_edit.setCurrentText(current_text)
        self.filter_edit.blockSignals(block)

    def _restore_trip_selection(self, numero: str) -> None:
        if self.trip_table.rowCount() == 0:
            return
        row_index = 0
        if numero:
            for index, row in enumerate(self.trip_rows):
                if str(row.get("numero", "") or "").strip() == numero:
                    row_index = index
                    break
        self.trip_table.selectRow(row_index)

    def open_trip_numero(self, numero: str) -> None:
        target = str(numero or "").strip()
        if not target:
            return
        self.refresh()
        for row_index, row in enumerate(self.trip_rows):
            if str(row.get("numero", "") or "").strip() != target:
                continue
            self.trip_table.selectRow(row_index)
            self._show_trip_detail()
            self._sync_actions()
            return

    def _selected_pending_rows(self) -> list[dict]:
        indexes = sorted({item.row() for item in self.pending_table.selectedItems()})
        return [self.pending_rows[index] for index in indexes if 0 <= index < len(self.pending_rows)]

    def _current_trip_row(self) -> dict:
        current = self.trip_table.currentItem()
        if current is None or current.row() >= len(self.trip_rows):
            return {}
        return self.trip_rows[current.row()]

    def _current_stop_row(self) -> dict:
        current = self.stops_table.currentItem()
        stops = list(self.current_detail.get("paragens", []) or [])
        if current is None or current.row() >= len(stops):
            return {}
        return stops[current.row()]

    def _clear_trip_detail(self) -> None:
        self.current_detail = {}
        self.detail_title.setText("Seleciona uma viagem")
        self.detail_meta.setText("Cria uma viagem e afeta encomendas a nosso cargo.")
        self.detail_note.setText("-")
        _apply_state_chip(self.detail_state_chip, "-")
        self.stops_table.setRowCount(0)

    def _show_trip_detail(self) -> None:
        row = self._current_trip_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            self._clear_trip_detail()
            self._sync_actions()
            return
        try:
            detail = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.current_detail = detail
        self.detail_title.setText(f"Viagem {detail.get('numero', '-')}")
        _apply_state_chip(self.detail_state_chip, str(detail.get("estado", "") or "-"))
        self.trip_status_combo.setCurrentText(str(detail.get("estado", "Planeado") or "Planeado"))
        transportadora_txt = str(detail.get("transportadora_nome", "") or "-").strip() or "-"
        pedido_state = str(detail.get("pedido_transporte_estado", "Nao pedido") or "Nao pedido").strip() or "Nao pedido"
        pedido_meta = pedido_state
        if detail.get("pedido_confirmado_at"):
            pedido_meta += f" confirmado em {detail.get('pedido_confirmado_at', '-')}"
        elif detail.get("pedido_recusado_at"):
            pedido_meta += f" recusado em {detail.get('pedido_recusado_at', '-')}"
        elif detail.get("pedido_transporte_at"):
            pedido_meta += f" em {detail.get('pedido_transporte_at', '-')}"
        if detail.get("pedido_transporte_by"):
            pedido_meta += f" por {detail.get('pedido_transporte_by', '-')}"
        self.detail_meta.setText(
            f"Planeado {detail.get('data_planeada', '-') or '-'} às {detail.get('hora_saida', '-') or '-'} | "
            f"Tipo {detail.get('tipo_responsavel', '-') or '-'}\n"
            f"Viatura {detail.get('viatura', '-') or '-'} | Matrícula {detail.get('matricula', '-') or '-'} | "
            f"Motorista {detail.get('motorista', '-') or '-'} | Telefone {detail.get('telefone_motorista', '-') or '-'}"
        )
        carga_txt = (
            f"{float(detail.get('paletes', 0) or 0):.2f} pal | "
            f"{float(detail.get('peso_bruto_kg', 0) or 0):.1f} kg | "
            f"{float(detail.get('volume_m3', 0) or 0):.3f} m3"
        )
        if bool(detail.get("carga_manual")):
            carga_txt += (
                f" (manual; calc. {float(detail.get('paletes_calculadas', 0) or 0):.2f} pal / "
                f"{float(detail.get('peso_bruto_kg_calculado', 0) or 0):.1f} kg / "
                f"{float(detail.get('volume_m3_calculado', 0) or 0):.3f} m3)"
            )
        self.detail_note.setText(
            f"Origem: {detail.get('origem', '-') or '-'} | Transportadora / fornecedor: {transportadora_txt} | "
            f"Ref. externa: {detail.get('referencia_transporte', '-') or '-'}\n"
            f"Pedido transporte: {pedido_meta} | Ref. pedido: {detail.get('pedido_transporte_ref', '-') or '-'}\n"
            f"Checklist OK: {int(detail.get('checklist_ok', 0) or 0)} | POD recebidos: {int(detail.get('pod_recebidos', 0) or 0)}\n"
            f"Zonas: {', '.join(list(detail.get('zonas', []) or [])) or '-'} | "
            f"Custo sugerido {_fmt_eur(float(detail.get('custo_sugerido_total', 0) or 0))}\n"
            f"Carga {carga_txt} | Preço {_fmt_eur(float(detail.get('preco_total', 0) or 0))} | "
            f"Custo {_fmt_eur(float(detail.get('custo_total', 0) or 0))}"
        )
        _fill_table(
            self.stops_table,
            [
                [
                    stop.get("ordem", "-"),
                    stop.get("encomenda_numero", "-"),
                    stop.get("cliente_nome", "-"),
                    stop.get("zona_transporte", "-"),
                    f"{float(stop.get('paletes', 0) or 0):.2f}",
                    f"{float(stop.get('peso_bruto_kg', 0) or 0):.1f}",
                    f"{float(stop.get('volume_m3', 0) or 0):.3f}",
                    stop.get("local_descarga", "-"),
                    stop.get("data_planeada", "-"),
                    stop.get("guia_numero", "-"),
                    stop.get("checklist_estado", "-"),
                    stop.get("pod_estado", "-"),
                    stop.get("estado", "-"),
                ]
                for stop in list(detail.get("paragens", []) or [])
            ],
            align_center_from=0,
        )
        for row_index, stop in enumerate(list(detail.get("paragens", []) or [])):
            _paint_table_row(self.stops_table, row_index, str(stop.get("estado", "")))
        self._sync_actions()

    def _sync_actions(self) -> None:
        has_trip = bool(self._current_trip_row())
        has_pending = bool(self._selected_pending_rows())
        has_stop = bool(self._current_stop_row())
        self.assign_btn.setEnabled(has_trip and has_pending)
        self.edit_trip_btn.setEnabled(has_trip)
        self.request_btn.setEnabled(has_trip)
        self.tariff_btn.setEnabled(True)
        self.apply_cost_btn.setEnabled(has_trip)
        self.remove_trip_btn.setEnabled(has_trip)
        self.trip_status_combo.setEnabled(has_trip)
        self.trip_status_btn.setEnabled(has_trip)
        self.stop_status_combo.setEnabled(has_stop)
        self.stop_status_btn.setEnabled(has_stop)
        self.edit_stop_btn.setEnabled(has_stop)
        self.stop_up_btn.setEnabled(has_stop)
        self.stop_down_btn.setEnabled(has_stop)
        self.remove_stop_btn.setEnabled(has_stop)
        self.pdf_btn.setEnabled(has_trip)

    def _supplier_options(self, initial: dict | None = None) -> list[str]:
        options = [str(value or "").strip() for value in list((initial or {}).get("supplier_options", []) or []) if str(value or "").strip()]
        if options:
            return options
        try:
            return [
                f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -")
                for row in list(self.backend.ne_suppliers() or [])
                if str(row.get("id", "") or "").strip() or str(row.get("nome", "") or "").strip()
            ]
        except Exception:
            return []

    def _zone_options(self, initial: dict | None = None) -> list[str]:
        options = [str(value or "").strip() for value in list((initial or {}).get("zone_options", []) or []) if str(value or "").strip()]
        if options:
            return options
        try:
            return [str(value or "").strip() for value in list(self.backend.transport_zone_options() or []) if str(value or "").strip()]
        except Exception:
            return []

    def _trip_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or self.backend.transport_defaults() or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Viagem de transporte")
        dialog.setMinimumWidth(860)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        numero_label = QLabel(str(initial.get("numero", "(nova)") or "(nova)"))
        data_edit = QDateEdit()
        data_edit.setCalendarPopup(True)
        data_edit.setDisplayFormat("yyyy-MM-dd")
        raw_date = str(initial.get("data_planeada", "") or "").strip()
        qdate = QDate.fromString(raw_date, "yyyy-MM-dd") if raw_date else QDate.currentDate()
        if not qdate.isValid():
            qdate = QDate.currentDate()
        data_edit.setDate(qdate)
        hora_edit = QTimeEdit()
        hora_edit.setDisplayFormat("HH:mm")
        raw_time = str(initial.get("hora_saida", "") or "08:00").strip()
        qtime = QTime.fromString(raw_time, "HH:mm")
        if not qtime.isValid():
            qtime = QTime(8, 0)
        hora_edit.setTime(qtime)
        tipo_combo = QComboBox()
        tipo_combo.addItems(["Nosso Cargo", "Subcontratado"])
        tipo_combo.setCurrentText(str(initial.get("tipo_responsavel", "Nosso Cargo") or "Nosso Cargo"))
        estado_combo = QComboBox()
        estado_combo.addItems(["Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        estado_combo.setCurrentText(str(initial.get("estado", "Planeado") or "Planeado"))
        viatura_combo = QComboBox()
        viatura_combo.setEditable(True)
        viatura_combo.addItem("")
        for value in list(initial.get("vehicle_options", []) or []):
            viatura_combo.addItem(str(value))
        viatura_combo.setCurrentText(str(initial.get("viatura", "") or "").strip())
        matricula_edit = QLineEdit(str(initial.get("matricula", "") or "").strip())
        motorista_combo = QComboBox()
        motorista_combo.setEditable(True)
        motorista_combo.addItem("")
        for value in list(initial.get("driver_options", []) or []):
            motorista_combo.addItem(str(value))
        motorista_combo.setCurrentText(str(initial.get("motorista", "") or "").strip())
        telefone_edit = QLineEdit(str(initial.get("telefone_motorista", "") or "").strip())
        origem_edit = QLineEdit(str(initial.get("origem", "") or "").strip())
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )
        ref_edit = QLineEdit(str(initial.get("referencia_transporte", "") or "").strip())
        cost_spin = QDoubleSpinBox()
        cost_spin.setRange(0.0, 1000000.0)
        cost_spin.setDecimals(2)
        cost_spin.setPrefix("EUR ")
        cost_spin.setValue(float(initial.get("custo_previsto", 0) or 0))
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(96)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        fields = [
            ("Número", numero_label, 0, 0),
            ("Data", data_edit, 0, 2),
            ("Saída", hora_edit, 0, 4),
            ("Tipo", tipo_combo, 1, 0),
            ("Estado", estado_combo, 1, 2),
            ("Transportadora", carrier_combo, 1, 4),
            ("Viatura", viatura_combo, 2, 0),
            ("Matrícula", matricula_edit, 2, 2),
            ("Motorista", motorista_combo, 2, 4),
            ("Telefone", telefone_edit, 3, 0),
            ("Origem", origem_edit, 3, 2),
            ("Ref. externa", ref_edit, 4, 0),
            ("Custo previsto", cost_spin, 4, 2),
        ]
        for label_text, widget, row, col in fields:
            form.addWidget(QLabel(label_text), row, col)
            span = 1
            form.addWidget(widget, row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 5, 0)
        form.addWidget(obs_edit, 5, 1, 1, 5)
        layout.addLayout(form)
        buttons_row = QHBoxLayout()
        if str(initial.get("numero", "") or "").strip():
            remove_btn = QPushButton("Apagar viagem")
            remove_btn.setProperty("variant", "danger")

            def _confirm_remove() -> None:
                if (
                    QMessageBox.question(
                        dialog,
                        "Transportes",
                        f"Apagar a viagem {str(initial.get('numero', '') or '').strip()} e libertar as encomendas associadas?",
                    )
                    != QMessageBox.Yes
                ):
                    return
                dialog.done(2)

            remove_btn.clicked.connect(_confirm_remove)
            buttons_row.addWidget(remove_btn)
        else:
            buttons_row.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        buttons_row.addWidget(buttons)
        layout.addLayout(buttons_row)
        result = dialog.exec()
        if result == 2:
            return {"_delete_trip": True, "numero": str(initial.get("numero", "") or "").strip()}
        if result != QDialog.Accepted:
            return None
        return {
            "numero": str(initial.get("numero", "") or "").strip(),
            "tipo_responsavel": tipo_combo.currentText().strip(),
            "estado": estado_combo.currentText().strip(),
            "data_planeada": data_edit.date().toString("yyyy-MM-dd"),
            "hora_saida": hora_edit.time().toString("HH:mm"),
            "viatura": viatura_combo.currentText().strip(),
            "matricula": matricula_edit.text().strip(),
            "motorista": motorista_combo.currentText().strip(),
            "telefone_motorista": telefone_edit.text().strip(),
            "origem": origem_edit.text().strip(),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "referencia_transporte": ref_edit.text().strip(),
            "custo_previsto": cost_spin.value(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _request_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Requisitar transporte")
        dialog.setMinimumWidth(620)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        estado_combo = QComboBox()
        estado_combo.addItems(["Nao pedido", "Pedido enviado", "Confirmado", "Recusado"])
        estado_combo.setCurrentText(str(initial.get("pedido_transporte_estado", "Pedido enviado") or "Pedido enviado"))
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )
        ref_edit = QLineEdit(str(initial.get("pedido_transporte_ref", "") or "").strip())
        paletes_spin = QDoubleSpinBox()
        paletes_spin.setRange(0.0, 9999.0)
        paletes_spin.setDecimals(2)
        paletes_spin.setSuffix(" pal")
        paletes_spin.setValue(float(initial.get("paletes_total_manual", 0) or 0))
        peso_spin = QDoubleSpinBox()
        peso_spin.setRange(0.0, 100000.0)
        peso_spin.setDecimals(2)
        peso_spin.setSuffix(" kg")
        peso_spin.setValue(float(initial.get("peso_total_manual_kg", 0) or 0))
        volume_spin = QDoubleSpinBox()
        volume_spin.setRange(0.0, 10000.0)
        volume_spin.setDecimals(3)
        volume_spin.setSuffix(" m3")
        volume_spin.setValue(float(initial.get("volume_total_manual_m3", 0) or 0))
        cost_spin = QDoubleSpinBox()
        cost_spin.setRange(0.0, 1000000.0)
        cost_spin.setDecimals(2)
        cost_spin.setPrefix("EUR ")
        cost_spin.setValue(float(initial.get("custo_previsto", 0) or 0))
        note_edit = QTextEdit()
        note_edit.setFixedHeight(90)
        note_edit.setPlainText(str(initial.get("pedido_transporte_obs", "") or "").strip())
        response_edit = QTextEdit()
        response_edit.setFixedHeight(72)
        response_edit.setPlainText(str(initial.get("pedido_resposta_obs", "") or "").strip())
        form.addRow("Estado pedido", estado_combo)
        form.addRow("Transportadora", carrier_combo)
        form.addRow("Ref. pedido", ref_edit)
        form.addRow("Paletes carga", paletes_spin)
        form.addRow("Peso total", peso_spin)
        form.addRow("Volume total", volume_spin)
        form.addRow("Custo previsto", cost_spin)
        form.addRow("Observações pedido", note_edit)
        form.addRow("Resposta parceiro", response_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "pedido_transporte_estado": estado_combo.currentText().strip(),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "pedido_transporte_ref": ref_edit.text().strip(),
            "paletes_total_manual": paletes_spin.value(),
            "peso_total_manual_kg": peso_spin.value(),
            "volume_total_manual_m3": volume_spin.value(),
            "custo_previsto": cost_spin.value(),
            "pedido_transporte_obs": note_edit.toPlainText().strip(),
            "pedido_resposta_obs": response_edit.toPlainText().strip(),
        }

    def _stop_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Paragem / guia de transporte")
        dialog.setMinimumWidth(640)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        guide_combo = QComboBox()
        guide_combo.setEditable(False)
        guide_combo.addItem("", "")
        for row in list(initial.get("guide_options", []) or []):
            label = str(row.get("label", "") or row.get("numero", "") or "").strip()
            guide_combo.addItem(label, str(row.get("numero", "") or "").strip())
        current_guide = str(initial.get("guia_numero", "") or "").strip()
        current_index = 0
        for idx in range(guide_combo.count()):
            if str(guide_combo.itemData(idx) or "").strip() == current_guide:
                current_index = idx
                break
        guide_combo.setCurrentIndex(current_index)
        local_edit = QLineEdit(str(initial.get("local_descarga", "") or "").strip())
        zone_combo = QComboBox()
        zone_combo.setEditable(True)
        zone_combo.addItem("")
        for value in self._zone_options(initial):
            zone_combo.addItem(str(value))
        zone_combo.setCurrentText(str(initial.get("zona_transporte", "") or "").strip())
        contacto_edit = QLineEdit(str(initial.get("contacto", "") or "").strip())
        telefone_edit = QLineEdit(str(initial.get("telefone", "") or "").strip())
        raw_dt = str(initial.get("data_planeada", "") or "").strip().replace(" ", "T")
        raw_date = raw_dt.split("T", 1)[0] if raw_dt else ""
        raw_time = raw_dt.split("T", 1)[1] if "T" in raw_dt else ""
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("yyyy-MM-dd")
        qdate = QDate.fromString(raw_date, "yyyy-MM-dd") if raw_date else QDate.currentDate()
        if not qdate.isValid():
            qdate = QDate.currentDate()
        date_edit.setDate(qdate)
        time_edit = QTimeEdit()
        time_edit.setDisplayFormat("HH:mm")
        qtime = QTime.fromString(raw_time[:5], "HH:mm") if raw_time else QTime(8, 0)
        if not qtime.isValid():
            qtime = QTime(8, 0)
        time_edit.setTime(qtime)
        carga_box = QCheckBox("Carga conferida")
        carga_box.setChecked(bool(initial.get("check_carga_ok")))
        docs_box = QCheckBox("Documentos conferidos")
        docs_box.setChecked(bool(initial.get("check_docs_ok")))
        paletes_box = QCheckBox("Paletes conferidas")
        paletes_box.setChecked(bool(initial.get("check_paletes_ok")))
        pod_state_combo = QComboBox()
        pod_state_combo.addItems(["", "Pendente", "Recebido", "Incidente"])
        pod_state_combo.setCurrentText(str(initial.get("pod_estado", "") or "").strip())
        pod_name_edit = QLineEdit(str(initial.get("pod_recebido_nome", "") or "").strip())
        raw_pod_dt = str(initial.get("pod_recebido_at", "") or "").strip().replace(" ", "T")
        raw_pod_date = raw_pod_dt.split("T", 1)[0] if raw_pod_dt else ""
        raw_pod_time = raw_pod_dt.split("T", 1)[1] if "T" in raw_pod_dt else ""
        pod_date_edit = QDateEdit()
        pod_date_edit.setCalendarPopup(True)
        pod_date_edit.setDisplayFormat("yyyy-MM-dd")
        pod_qdate = QDate.fromString(raw_pod_date, "yyyy-MM-dd") if raw_pod_date else QDate.currentDate()
        if not pod_qdate.isValid():
            pod_qdate = QDate.currentDate()
        pod_date_edit.setDate(pod_qdate)
        pod_time_edit = QTimeEdit()
        pod_time_edit.setDisplayFormat("HH:mm")
        pod_qtime = QTime.fromString(raw_pod_time[:5], "HH:mm") if raw_pod_time else QTime.currentTime()
        if not pod_qtime.isValid():
            pod_qtime = QTime.currentTime()
        pod_time_edit.setTime(pod_qtime)
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(88)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        pod_obs_edit = QTextEdit()
        pod_obs_edit.setFixedHeight(72)
        pod_obs_edit.setPlainText(str(initial.get("pod_obs", "") or "").strip())
        form.addRow("Guia associada", guide_combo)
        form.addRow("Local descarga", local_edit)
        form.addRow("Zona", zone_combo)
        form.addRow("Contacto", contacto_edit)
        form.addRow("Telefone", telefone_edit)
        form.addRow("Data", date_edit)
        form.addRow("Hora", time_edit)
        form.addRow("Checklist", carga_box)
        form.addRow("", docs_box)
        form.addRow("", paletes_box)
        form.addRow("POD estado", pod_state_combo)
        form.addRow("Recebido por", pod_name_edit)
        form.addRow("Data POD", pod_date_edit)
        form.addRow("Hora POD", pod_time_edit)
        form.addRow("Obs. POD", pod_obs_edit)
        form.addRow("Observações", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected_guide = str(guide_combo.currentData() or "").strip()
        return {
            "expedicao_numero": selected_guide,
            "local_descarga": local_edit.text().strip(),
            "zona_transporte": zone_combo.currentText().strip(),
            "contacto": contacto_edit.text().strip(),
            "telefone": telefone_edit.text().strip(),
            "data_planeada": f"{date_edit.date().toString('yyyy-MM-dd')}T{time_edit.time().toString('HH:mm')}:00",
            "check_carga_ok": carga_box.isChecked(),
            "check_docs_ok": docs_box.isChecked(),
            "check_paletes_ok": paletes_box.isChecked(),
            "pod_estado": pod_state_combo.currentText().strip(),
            "pod_recebido_nome": pod_name_edit.text().strip(),
            "pod_recebido_at": f"{pod_date_edit.date().toString('yyyy-MM-dd')}T{pod_time_edit.time().toString('HH:mm')}:00" if pod_state_combo.currentText().strip() else "",
            "pod_obs": pod_obs_edit.toPlainText().strip(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _text_prompt(self, title: str, label: str) -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        edit = QLineEdit()
        form.addRow(label, edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return edit.text().strip()

    def _tariff_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or self.backend.transport_tariff_defaults() or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Tarifário de transporte")
        dialog.setMinimumWidth(620)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )
        zone_combo = QComboBox()
        zone_combo.setEditable(True)
        zone_combo.addItem("")
        for value in self._zone_options(initial):
            zone_combo.addItem(str(value))
        zone_combo.setCurrentText(str(initial.get("zona", "") or "").strip())
        base_spin = QDoubleSpinBox()
        base_spin.setRange(0.0, 1000000.0)
        base_spin.setDecimals(2)
        base_spin.setPrefix("EUR ")
        base_spin.setValue(float(initial.get("valor_base", 0) or 0))
        palette_spin = QDoubleSpinBox()
        palette_spin.setRange(0.0, 1000000.0)
        palette_spin.setDecimals(2)
        palette_spin.setPrefix("EUR ")
        palette_spin.setValue(float(initial.get("valor_por_palete", 0) or 0))
        kg_spin = QDoubleSpinBox()
        kg_spin.setRange(0.0, 1000.0)
        kg_spin.setDecimals(4)
        kg_spin.setPrefix("EUR ")
        kg_spin.setValue(float(initial.get("valor_por_kg", 0) or 0))
        volume_spin = QDoubleSpinBox()
        volume_spin.setRange(0.0, 1000000.0)
        volume_spin.setDecimals(2)
        volume_spin.setPrefix("EUR ")
        volume_spin.setValue(float(initial.get("valor_por_m3", 0) or 0))
        minimum_spin = QDoubleSpinBox()
        minimum_spin.setRange(0.0, 1000000.0)
        minimum_spin.setDecimals(2)
        minimum_spin.setPrefix("EUR ")
        minimum_spin.setValue(float(initial.get("custo_minimo", 0) or 0))
        active_box = QCheckBox("Ativo")
        active_box.setChecked(bool(initial.get("ativo", True)))
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(90)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        form.addRow("Transportadora", carrier_combo)
        form.addRow("Zona", zone_combo)
        form.addRow("Valor base", base_spin)
        form.addRow("Valor / palete", palette_spin)
        form.addRow("Valor / kg", kg_spin)
        form.addRow("Valor / m3", volume_spin)
        form.addRow("Custo mínimo", minimum_spin)
        form.addRow("", active_box)
        form.addRow("Observações", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "id": initial.get("id"),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "zona": zone_combo.currentText().strip(),
            "valor_base": base_spin.value(),
            "valor_por_palete": palette_spin.value(),
            "valor_por_kg": kg_spin.value(),
            "valor_por_m3": volume_spin.value(),
            "custo_minimo": minimum_spin.value(),
            "ativo": active_box.isChecked(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _manage_tariffs(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Tarifário por transportadora / zona")
        dialog.setMinimumSize(920, 520)
        layout = QVBoxLayout(dialog)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Pesquisa"))
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("Filtrar por transportadora, zona ou observações")
        filter_row.addWidget(filter_edit, 1)
        layout.addLayout(filter_row)
        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["ID", "Transportadora", "Zona", "Base", "Palete", "Kg", "M3", "Mínimo"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(1, 2), contents=(0, 3, 4, 5, 6, 7))
        _set_table_columns(
            table,
            [
                (0, "interactive", 60),
                (1, "stretch", 230),
                (2, "stretch", 170),
                (3, "interactive", 90),
                (4, "interactive", 90),
                (5, "interactive", 90),
                (6, "interactive", 90),
                (7, "interactive", 90),
            ],
        )
        layout.addWidget(table, 1)
        buttons_row = QHBoxLayout()
        new_btn = QPushButton("Novo")
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        buttons_row.addStretch(1)
        buttons_row.addWidget(new_btn)
        buttons_row.addWidget(edit_btn)
        buttons_row.addWidget(remove_btn)
        buttons_row.addWidget(close_btn)
        layout.addLayout(buttons_row)
        state: dict[str, list[dict[str, Any]]] = {"rows": []}

        def current_row() -> dict[str, Any]:
            item = table.currentItem()
            if item is None or item.row() >= len(state["rows"]):
                return {}
            return state["rows"][item.row()]

        def render() -> None:
            rows = list(self.backend.transport_tariff_rows(filter_edit.text().strip()) or [])
            state["rows"] = rows
            _fill_table(
                table,
                [
                    [
                        row.get("id", "-"),
                        row.get("transportadora_nome", "Sem transportadora"),
                        row.get("zona", "-"),
                        _fmt_eur(float(row.get("valor_base", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_palete", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_kg", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_m3", 0) or 0)),
                        _fmt_eur(float(row.get("custo_minimo", 0) or 0)),
                    ]
                    for row in rows
                ],
                align_center_from=0,
            )
            for row_index, row in enumerate(rows):
                _paint_table_row(table, row_index, "Concluído" if bool(row.get("ativo", True)) else "Anulado")
            edit_btn.setEnabled(bool(rows))
            remove_btn.setEnabled(bool(rows))

        def handle_new() -> None:
            payload = self._tariff_dialog()
            if payload is None:
                return
            try:
                self.backend.transport_tariff_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        def handle_edit() -> None:
            row = current_row()
            if not row:
                QMessageBox.warning(dialog, "Tarifário", "Seleciona um tarifário.")
                return
            payload = self._tariff_dialog(row)
            if payload is None:
                return
            try:
                self.backend.transport_tariff_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        def handle_remove() -> None:
            row = current_row()
            if not row:
                QMessageBox.warning(dialog, "Tarifário", "Seleciona um tarifário.")
                return
            if QMessageBox.question(dialog, "Tarifário", f"Remover o tarifário da zona {row.get('zona', '-') or '-'}?") != QMessageBox.Yes:
                return
            try:
                self.backend.transport_tariff_remove(row.get("id"))
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        filter_edit.textChanged.connect(lambda _text: render())
        new_btn.clicked.connect(handle_new)
        edit_btn.clicked.connect(handle_edit)
        remove_btn.clicked.connect(handle_remove)
        close_btn.clicked.connect(dialog.accept)
        render()
        dialog.exec()

    def _apply_suggested_costs(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            detail = self.backend.transport_apply_suggested_cost(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()
        QMessageBox.information(
            self,
            "Transportes",
            (
                f"Custo sugerido aplicado à viagem {numero}.\n"
                f"Novo custo previsto: {_fmt_eur(float(detail.get('custo_previsto', 0) or 0))}"
            ),
        )

    def _new_trip(self) -> None:
        payload = self._trip_dialog(self.backend.transport_defaults())
        if payload is None:
            return
        try:
            detail = self.backend.transport_create_or_update(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(str(detail.get("numero", "") or "").strip())
        self._show_trip_detail()

    def _edit_trip(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            initial = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._trip_dialog(initial)
        if payload is None:
            return
        if payload.get("_delete_trip"):
            try:
                self.backend.transport_remove_trip(numero)
            except Exception as exc:
                QMessageBox.critical(self, "Transportes", str(exc))
                return
            self.refresh()
            self._clear_trip_detail()
            return
        try:
            self.backend.transport_create_or_update(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _remove_trip(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        if QMessageBox.question(self, "Transportes", f"Apagar a viagem {numero} e libertar as encomendas associadas?") != QMessageBox.Yes:
            return
        try:
            self.backend.transport_remove_trip(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._clear_trip_detail()

    def _request_transport(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            initial = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._request_dialog(initial)
        if payload is None:
            return
        try:
            self.backend.transport_request_service(numero, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _edit_stop(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        try:
            guide_options = self.backend.transport_guide_options(enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._stop_dialog({**dict(stop), "guide_options": guide_options})
        if payload is None:
            return
        try:
            self.backend.transport_update_stop(numero, enc_num, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()
        for row_index, row in enumerate(list(self.current_detail.get("paragens", []) or [])):
            if str(row.get("encomenda_numero", "") or "").strip() == enc_num:
                self.stops_table.selectRow(row_index)
                break

    def _assign_selected_orders(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        rows = self._selected_pending_rows()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona primeiro uma viagem.")
            return
        if not rows:
            QMessageBox.warning(self, "Transportes", "Seleciona pelo menos uma encomenda.")
            return
        try:
            self.backend.transport_assign_orders(numero, [str(row.get("numero", "") or "").strip() for row in rows])
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _apply_trip_status(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            self.backend.transport_set_status(numero, self.trip_status_combo.currentText().strip())
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _apply_stop_status(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        note = ""
        if self.stop_status_combo.currentText().strip() == "Incidente":
            note = self._text_prompt("Incidente na paragem", "Motivo")
            if note is None:
                return
        try:
            self.backend.transport_set_stop_status(numero, enc_num, self.stop_status_combo.currentText().strip(), note)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _move_stop(self, direction: int) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        try:
            self.backend.transport_move_stop(numero, enc_num, direction)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _remove_stop(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        if QMessageBox.question(self, "Transportes", f"Remover a encomenda {enc_num} desta viagem?") != QMessageBox.Yes:
            return
        try:
            self.backend.transport_remove_stop(numero, enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _open_trip_pdf(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            path = self.backend.transport_route_sheet_open(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        QMessageBox.information(self, "Transportes", f"PDF aberto:\n{path}")


class PurchaseNotesPage(QWidget):
    page_title = "Notas Encomenda"
    page_subtitle = "Pedido de cotação, adjudicação por linha e notas de compra reais por fornecedor."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.line_rows: list[dict] = []
        self.supplier_rows: list[dict] = []
        self.current_number = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        filters = CardFrame()
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(16, 14, 16, 14)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por número, fornecedor ou estado")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Em edicao", "Aprovada", "Parcial", "Entregue", "Convertidas", "Todas"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.state_combo)
        root.addWidget(filters)

        notes_card = CardFrame()
        notes_card.set_tone("default")
        notes_layout = QVBoxLayout(notes_card)
        notes_layout.setContentsMargins(16, 14, 16, 14)
        notes_title = QLabel("Pedidos / NEs")
        notes_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.notes_table = QTableWidget(0, 6)
        self.notes_table.setHorizontalHeaderLabels(["Número", "Fornecedor", "Entrega", "Estado", "Total", "Linhas"])
        self.notes_table.verticalHeader().setVisible(False)
        self.notes_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.notes_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.notes_table, stretch=(1,), contents=(0, 2, 3, 4, 5))
        self.notes_table.itemSelectionChanged.connect(self._load_selected_note)
        notes_layout.addWidget(notes_title)
        notes_layout.addWidget(self.notes_table)
        root.addWidget(notes_card)

        self.note_form_card = CardFrame()
        form_layout = QVBoxLayout(self.note_form_card)
        form_layout.setContentsMargins(12, 8, 12, 10)
        form_layout.setSpacing(6)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(8)
        self.new_btn = QPushButton("Novo pedido")
        self.new_btn.clicked.connect(self._create_new_note)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_note)
        self.approve_btn = QPushButton("Aprovar")
        self.approve_btn.clicked.connect(self._approve_note)
        self.generate_btn = QPushButton("Gerar NEs")
        self.generate_btn.clicked.connect(self._generate_supplier_orders)
        self.remove_btn = QPushButton("Apagar")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_note)
        self.pdf_btn = QPushButton("Pre-visualizar NE")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(lambda: self._open_pdf(False))
        self.quote_btn = QPushButton("Pedir orçamento")
        self.quote_btn.clicked.connect(lambda: self._open_pdf(True))
        self.quote_btn.setStyleSheet(
            "QPushButton {background: #f4c542; color: #0f172a; border: 1px solid #caa12b; "
            "border-radius: 10px; padding: 8px 14px; font-weight: 800;}"
            "QPushButton:hover {background: #ffd65a;}"
            "QPushButton:disabled {background: #f5e9b4; color: #7c6a2b;}"
        )
        self.suppliers_btn = QPushButton("Fornecedores")
        self.suppliers_btn.setProperty("variant", "secondary")
        self.suppliers_btn.clicked.connect(self._manage_suppliers)
        self.delivery_btn = QPushButton("Entregar encomenda")
        self.delivery_btn.setProperty("variant", "secondary")
        self.delivery_btn.clicked.connect(self._register_delivery)
        self.attach_doc_btn = QPushButton("Associar Docs")
        self.attach_doc_btn.setProperty("variant", "secondary")
        self.attach_doc_btn.clicked.connect(self._attach_document)
        self.documents_btn = QPushButton("Registo Docs")
        self.documents_btn.setProperty("variant", "secondary")
        self.documents_btn.clicked.connect(self._show_documents)
        for button in (
            self.new_btn,
            self.save_btn,
            self.approve_btn,
            self.generate_btn,
            self.remove_btn,
            self.pdf_btn,
            self.suppliers_btn,
            self.delivery_btn,
            self.attach_doc_btn,
            self.documents_btn,
        ):
            actions_row.addWidget(button)
        actions_row.addStretch(1)
        actions_row.addWidget(self.quote_btn)
        form_layout.addLayout(actions_row)

        header = QHBoxLayout()
        self.number_label = QLabel("Novo pedido")
        self.number_label.setStyleSheet("font-size: 17px; font-weight: 800; color: #0f172a;")
        self.state_chip = QLabel("-")
        _apply_state_chip(self.state_chip, "-")
        header.addWidget(self.number_label, 1)
        header.addWidget(self.state_chip)
        form_layout.addLayout(header)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(12)
        supplier_card = CardFrame()
        supplier_card.set_tone("default")
        supplier_layout = QVBoxLayout(supplier_card)
        supplier_layout.setContentsMargins(12, 10, 12, 10)
        supplier_layout.setSpacing(6)
        self.supplier_combo = QComboBox()
        self._configure_supplier_selector(self.supplier_combo)
        self.supplier_combo.currentTextChanged.connect(self._sync_supplier_contact)
        self.contact_edit = QLineEdit()
        self.delivery_edit = QDateEdit()
        self.delivery_edit.setCalendarPopup(True)
        self.delivery_edit.setDisplayFormat("yyyy-MM-dd")
        self.delivery_edit.setMinimumDate(QDate(2000, 1, 1))
        self.delivery_edit.setDate(QDate.currentDate())
        self.location_edit = QComboBox()
        self.location_edit.setEditable(True)
        self.transport_edit = QComboBox()
        self.transport_edit.setEditable(True)
        self.transport_edit.addItems(["", "Nosso Cargo", "Vosso Cargo"])
        self.location_edit.addItems(["", "Vossas Instalações", "Nossas Instalações"])
        note_field_css = "font-size: 12px; padding: 4px 8px;"
        for field in (self.supplier_combo, self.contact_edit, self.delivery_edit, self.transport_edit, self.location_edit):
            field.setMinimumHeight(38)
            field.setStyleSheet(note_field_css)
        for combo in (self.supplier_combo, self.transport_edit, self.location_edit):
            try:
                combo.view().setStyleSheet("font-size: 12px;")
            except Exception:
                pass
        for widget, width, minimum in (
            (self.supplier_combo, 760, 700),
            (self.contact_edit, 290, 250),
            (self.delivery_edit, 164, 164),
            (self.transport_edit, 240, 230),
            (self.location_edit, 470, 420),
        ):
            _cap_width(widget, width)
            widget.setMinimumWidth(minimum)
        supplier_title = QLabel("Dados do documento")
        supplier_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        supplier_layout.addWidget(supplier_title)

        def _field_block(label_text: str, widget: QWidget) -> QWidget:
            host = QWidget()
            block = QVBoxLayout(host)
            block.setContentsMargins(0, 0, 0, 0)
            block.setSpacing(1)
            label = QLabel(label_text)
            label.setProperty("role", "muted")
            label.setStyleSheet("font-size: 11px; font-weight: 700; color: #415a77;")
            block.addWidget(label)
            block.addWidget(widget)
            return host

        supplier_grid = QGridLayout()
        supplier_grid.setContentsMargins(0, 0, 0, 0)
        supplier_grid.setHorizontalSpacing(8)
        supplier_grid.setVerticalSpacing(4)
        supplier_grid.addWidget(_field_block("Fornecedor base", self.supplier_combo), 0, 0)
        supplier_grid.addWidget(_field_block("Contacto", self.contact_edit), 0, 1)
        supplier_grid.addWidget(_field_block("Entrega", self.delivery_edit), 0, 2)
        supplier_grid.addWidget(_field_block("Transporte", self.transport_edit), 1, 0)
        supplier_grid.addWidget(_field_block("Local Descarga", self.location_edit), 1, 1, 1, 2)
        supplier_grid.setColumnStretch(0, 7)
        supplier_grid.setColumnStretch(1, 3)
        supplier_grid.setColumnStretch(2, 2)
        supplier_layout.addLayout(supplier_grid)
        supplier_card.setFixedHeight(170)
        supplier_card.setMinimumWidth(1320)
        supplier_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        summary_card = CardFrame()
        summary_card.set_tone("info")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(16, 12, 16, 12)
        summary_layout.setSpacing(5)
        summary_title = QLabel("Resumo")
        summary_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.summary_hint = QLabel("Pedido")
        self.summary_hint.setWordWrap(True)
        self.summary_hint.setProperty("role", "muted")
        self.summary_hint.setMaximumHeight(24)
        self.total_label = QLabel("0.00 EUR")
        self.total_label.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        self.total_label.setMinimumHeight(24)
        self.total_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_materials_label = QLabel("0,00 EUR")
        self.total_materials_label.setStyleSheet("font-size: 12px; font-weight: 800; color: #7c4a03;")
        self.total_materials_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_materials_label.setMinimumHeight(18)
        self.total_products_label = QLabel("0,00 EUR")
        self.total_products_label.setStyleSheet("font-size: 12px; font-weight: 800; color: #166534;")
        self.total_products_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_products_label.setMinimumHeight(18)
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(self.summary_hint)
        summary_grid = QGridLayout()
        summary_grid.setContentsMargins(0, 0, 0, 0)
        summary_grid.setHorizontalSpacing(10)
        summary_grid.setVerticalSpacing(4)
        mp_caption = QLabel("Matéria-Prima")
        mp_caption.setProperty("role", "muted")
        products_caption = QLabel("Produtos")
        products_caption.setProperty("role", "muted")
        summary_grid.addWidget(mp_caption, 0, 0)
        summary_grid.addWidget(self.total_materials_label, 0, 1)
        summary_grid.addWidget(products_caption, 1, 0)
        summary_grid.addWidget(self.total_products_label, 1, 1)
        summary_grid.setColumnStretch(0, 1)
        summary_grid.setColumnStretch(1, 1)
        summary_layout.addLayout(summary_grid)
        total_row = QHBoxLayout()
        total_row.setContentsMargins(0, 2, 0, 0)
        total_caption = QLabel("Total")
        total_caption.setProperty("role", "muted")
        total_row.addWidget(total_caption)
        total_row.addStretch(1)
        total_row.addWidget(self.total_label)
        summary_layout.addLayout(total_row)
        summary_card.setMinimumWidth(350)
        summary_card.setMaximumWidth(390)
        summary_card.setFixedHeight(170)
        summary_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        meta_row.addWidget(supplier_card, 10, Qt.AlignTop)
        meta_row.addWidget(summary_card, 3, Qt.AlignTop)
        form_layout.addLayout(meta_row)

        obs_card = CardFrame()
        obs_card.set_tone("default")
        obs_layout = QVBoxLayout(obs_card)
        obs_layout.setContentsMargins(10, 8, 10, 8)
        obs_layout.setSpacing(4)
        obs_title = QLabel("Observações")
        obs_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.obs_edit = QTextEdit()
        self.obs_edit.setPlaceholderText("Observações do documento")
        self.obs_edit.setMinimumHeight(38)
        self.obs_edit.setMaximumHeight(52)
        self.obs_edit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.obs_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        obs_layout.addWidget(obs_title)
        obs_layout.addWidget(self.obs_edit)
        obs_card.setMaximumHeight(82)
        form_layout.addWidget(obs_card)

        lines_card = CardFrame()
        lines_card.set_tone("default")
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(12, 10, 12, 12)
        lines_layout.setSpacing(6)
        line_header = QHBoxLayout()
        lines_title = QLabel("Linhas da nota")
        lines_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        line_header.addWidget(lines_title)
        line_header.addStretch(1)
        lines_layout.addLayout(line_header)

        line_actions = QHBoxLayout()
        self.add_material_btn = QPushButton("Adicionar Stock MP")
        self.add_material_btn.clicked.connect(self._add_material_line)
        self.add_material_btn.setToolTip("Adicionar matéria-prima de stock, incluindo chapa, perfil e tubo.")
        self.add_product_btn = QPushButton("Stock produtos")
        self.add_product_btn.setProperty("variant", "secondary")
        self.add_product_btn.clicked.connect(self._add_product_line)
        self.add_manual_btn = QPushButton("Adicionar Manual")
        self.add_manual_btn.setProperty("variant", "secondary")
        self.add_manual_btn.clicked.connect(self._add_manual_line)
        self.edit_line_btn = QPushButton("Editar Linha")
        self.edit_line_btn.setProperty("variant", "secondary")
        self.edit_line_btn.clicked.connect(self._edit_line)
        self.quick_supplier_btn = QPushButton("Fornecedor rápido")
        self.quick_supplier_btn.setProperty("variant", "secondary")
        self.quick_supplier_btn.clicked.connect(self._quick_assign_supplier)
        self.remove_line_btn = QPushButton("Remover Linha")
        self.remove_line_btn.setProperty("variant", "danger")
        self.remove_line_btn.clicked.connect(self._remove_line)
        for button in (self.add_material_btn, self.add_product_btn, self.add_manual_btn, self.edit_line_btn, self.quick_supplier_btn, self.remove_line_btn):
            line_actions.addWidget(button)
        line_actions.addStretch(1)
        lines_layout.addLayout(line_actions)

        self.lines_table = QTableWidget(0, 11)
        self.lines_table.setHorizontalHeaderLabels(["Código", "Descrição", "Origem", "Fornecedor", "Qtd", "Unid.", "P.Unit.", "Desc.%", "IVA%", "Total", "Entrega"])
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.lines_table)
        _set_table_columns(
            self.lines_table,
            [
                (0, "fixed", 150),
                (1, "stretch", 0),
                (2, "fixed", 110),
                (3, "stretch", 0),
                (4, "fixed", 72),
                (5, "fixed", 62),
                (6, "fixed", 96),
                (7, "fixed", 74),
                (8, "fixed", 68),
                (9, "fixed", 102),
                (10, "fixed", 118),
            ],
        )
        self.lines_table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QHeaderView::section { font-size: 11px; padding: 6px 8px; font-weight: 800; }"
        )
        self.lines_table.verticalHeader().setDefaultSectionSize(28)
        self.lines_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.lines_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.lines_table.setMinimumHeight(320)
        self.lines_table.setMaximumHeight(380)
        lines_layout.addWidget(self.lines_table)
        lines_card.setMinimumHeight(390)
        lines_card.setMaximumHeight(450)
        form_layout.addWidget(lines_card, 1)
        self.note_form_card.set_tone("default")
        root.addWidget(self.note_form_card, 1)
        self.current_documents: list[dict] = []

        self._new_note()

    def refresh(self) -> None:
        previous = self.current_number
        self.supplier_rows = self.backend.ne_suppliers()
        self._set_supplier_items()
        self.rows = self.backend.ne_rows(self.filter_edit.currentText().strip(), self.state_combo.currentText())
        _fill_table(
            self.notes_table,
            [[r.get("numero", "-"), r.get("fornecedor", "-"), r.get("data_entrega", "-"), r.get("estado", "-"), _fmt_eur(r.get("total", 0)), r.get("linhas", 0)] for r in self.rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.notes_table, row_index, str(row.get("estado", "")))
        if self.notes_table.rowCount() == 0:
            self._new_note(reset_number=False)
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "")).strip() == previous:
                    row_index = index
                    break
        self.notes_table.selectRow(row_index)
        self._load_selected_note()

    def _set_supplier_items(self) -> None:
        current = self.supplier_combo.currentText().strip()
        self.supplier_combo.blockSignals(True)
        self.supplier_combo.clear()
        self.supplier_combo.addItem("")
        for row in self.supplier_rows:
            label = f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -")
            self.supplier_combo.addItem(label)
        if current:
            self.supplier_combo.setCurrentText(current)
        self.supplier_combo.blockSignals(False)

    def _configure_supplier_selector(self, combo: QComboBox) -> None:
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMaxVisibleItems(18)
        completer = combo.completer()
        if completer is not None:
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            try:
                completer.setFilterMode(Qt.MatchContains)
            except Exception:
                pass

    def _selected_line_indexes(self) -> list[int]:
        indexes = sorted({idx.row() for idx in self.lines_table.selectionModel().selectedRows()}) if self.lines_table.selectionModel() else []
        return [idx for idx in indexes if 0 <= idx < len(self.line_rows)]

    def _state_text(self) -> str:
        return self.state_chip.text().strip() or "Em edicao"

    def _selected_row(self) -> dict:
        current = self.notes_table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def open_note_numero(self, numero: str) -> None:
        target = str(numero or "").strip()
        if not target:
            return
        self.refresh()
        for row_index, row in enumerate(self.rows):
            if str(row.get("numero", "") or "").strip() != target:
                continue
            self.notes_table.selectRow(row_index)
            self._load_selected_note()
            if hasattr(self, "_show_note_detail"):
                try:
                    self._show_note_detail()
                except Exception:
                    pass
            return

    def _selected_line_index(self) -> int:
        current = self.lines_table.currentItem()
        if current is None:
            return -1
        return current.row() if current.row() < len(self.line_rows) else -1

    def _supplier_lookup(self, text: str) -> dict:
        raw = str(text or "").strip()
        if not raw:
            return {}
        supplier_id = raw.split(" - ", 1)[0].strip()
        for row in self.supplier_rows:
            if supplier_id and supplier_id == str(row.get("id", "")).strip():
                return row
            if raw.lower() == f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -").lower():
                return row
            if raw.lower() == str(row.get("nome", "")).strip().lower():
                return row
        return {}

    def _sync_supplier_contact(self) -> None:
        supplier = self._supplier_lookup(self.supplier_combo.currentText())
        if supplier and not self.contact_edit.text().strip():
            self.contact_edit.setText(str(supplier.get("contacto", "") or "").strip())

    def _render_lines(self) -> None:
        total = 0.0
        materials_total = 0.0
        products_total = 0.0
        def _line_code(row: dict[str, Any]) -> str:
            ref = str(row.get("ref", "") or "").strip()
            if ref:
                return ref
            if self.backend.desktop_main.origem_is_materia(row.get("origem", "")):
                return "ID pendente"
            return "Manual"
        _fill_table(
            self.lines_table,
            [
                [
                    _line_code(row),
                    row.get("descricao", "-"),
                    row.get("origem", "-"),
                    row.get("fornecedor_linha", "-"),
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    row.get("unid", "-"),
                    f"{float(row.get('preco', 0) or 0):.4f}",
                    f"{float(row.get('desconto', 0) or 0):.2f}",
                    f"{float(row.get('iva', 0) or 0):.2f}",
                    f"{float(row.get('total', 0) or 0):.2f}",
                    row.get("entrega", "PENDENTE"),
                ]
                for row in self.line_rows
            ],
            align_center_from=3,
        )
        for row_index, row in enumerate(self.line_rows):
            row_total = float(row.get("total", 0) or 0)
            total += row_total
            origem = self.backend.desktop_main.norm_text(str(row.get("origem", "") or "").strip())
            if "mater" in origem:
                materials_total += row_total
            else:
                products_total += row_total
            entrega = str(row.get("entrega", "PENDENTE")).upper()
            tone = "Concluida" if "ENTREGUE" in entrega else "Em pausa" if "PARCIAL" in entrega else "Preparacao"
            _paint_table_row(self.lines_table, row_index, tone)
            code_item = self.lines_table.item(row_index, 0)
            desc_item = self.lines_table.item(row_index, 1)
            supplier_item = self.lines_table.item(row_index, 3)
            if code_item is not None:
                code_item.setToolTip(_line_code(row))
            if desc_item is not None:
                desc_item.setToolTip(str(row.get("descricao", "") or "").strip())
            if supplier_item is not None:
                supplier_item.setToolTip(str(row.get("fornecedor_linha", "") or "").strip())
        self.total_label.setText(_fmt_eur(total))
        self.total_materials_label.setText(_fmt_eur(materials_total))
        self.total_products_label.setText(_fmt_eur(products_total))

    def _set_header(self, numero: str, estado: str) -> None:
        self.current_number = str(numero or "").strip()
        self.number_label.setText(self.current_number or "Novo pedido")
        _apply_state_chip(self.state_chip, estado)
        _set_panel_tone(self.note_form_card, _state_tone(estado))

    def _delivery_text(self) -> str:
        selected = self.delivery_edit.date()
        if not selected.isValid() or selected.year() <= 2000:
            selected = QDate.currentDate()
            self.delivery_edit.setDate(selected)
        return selected.toString("yyyy-MM-dd").strip()

    def _set_delivery_text(self, raw: str) -> None:
        self.delivery_edit.setDate(_coerce_editor_qdate(raw, fallback_today=True))

    def _document_dialog(self, title: str, initial: dict | None = None, allow_line_apply: bool = True) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(860, 420)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        type_combo = QComboBox()
        for label, value in (
            ("Guia", "GUIA"),
            ("Fatura", "FATURA"),
            ("Guia + Fatura", "GUIA_FATURA"),
            ("Documento", "DOCUMENTO"),
        ):
            type_combo.addItem(label, value)
        initial_type = str(initial.get("tipo", "") or "").strip().upper() or "FATURA"
        for index in range(type_combo.count()):
            if str(type_combo.itemData(index) or "").strip().upper() == initial_type:
                type_combo.setCurrentIndex(index)
                break
        title_edit = QLineEdit(str(initial.get("titulo", "") or ""))
        guia_edit = QLineEdit(str(initial.get("guia", "") or ""))
        fatura_edit = QLineEdit(str(initial.get("fatura", "") or ""))
        path_edit = QLineEdit(str(initial.get("caminho", "") or ""))
        obs_edit = QLineEdit(str(initial.get("obs", "") or ""))
        entrega_edit = QDateEdit()
        entrega_edit.setCalendarPopup(True)
        entrega_edit.setDisplayFormat("yyyy-MM-dd")
        entrega_edit.setDate(_coerce_editor_qdate(str(initial.get("data_entrega", "") or ""), fallback_today=True))
        doc_edit = QDateEdit()
        doc_edit.setCalendarPopup(True)
        doc_edit.setDisplayFormat("yyyy-MM-dd")
        doc_edit.setDate(_coerce_editor_qdate(str(initial.get("data_documento", "") or ""), fallback_today=True))
        apply_lines_chk = QCheckBox("Aplicar aos registos de linhas já entregues")
        apply_lines_chk.setChecked(bool(initial.get("apply_to_lines", True)))
        apply_lines_chk.setVisible(allow_line_apply)

        def pick_path() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar documento",
                "",
                "Documentos (*.pdf *.png *.jpg *.jpeg *.bmp *.xlsx *.xls *.doc *.docx);;Todos (*.*)",
            )
            if path:
                path_edit.setText(path)

        browse_btn = QPushButton("Selecionar ficheiro")
        browse_btn.setProperty("variant", "secondary")
        browse_btn.clicked.connect(pick_path)

        form.addWidget(QLabel("Tipo"), 0, 0)
        form.addWidget(type_combo, 0, 1)
        form.addWidget(QLabel("Titulo"), 0, 2)
        form.addWidget(title_edit, 0, 3)
        form.addWidget(QLabel("Data entrega"), 1, 0)
        form.addWidget(entrega_edit, 1, 1)
        form.addWidget(QLabel("Data documento"), 1, 2)
        form.addWidget(doc_edit, 1, 3)
        form.addWidget(QLabel("Guia"), 2, 0)
        form.addWidget(guia_edit, 2, 1)
        form.addWidget(QLabel("Fatura"), 2, 2)
        form.addWidget(fatura_edit, 2, 3)
        form.addWidget(QLabel("Caminho"), 3, 0)
        form.addWidget(path_edit, 3, 1, 1, 2)
        form.addWidget(browse_btn, 3, 3)
        form.addWidget(QLabel("Obs."), 4, 0)
        form.addWidget(obs_edit, 4, 1, 1, 3)
        layout.addLayout(form)
        if allow_line_apply:
            layout.addWidget(apply_lines_chk)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = {
            "tipo": str(type_combo.currentData() or "DOCUMENTO").strip(),
            "titulo": title_edit.text().strip(),
            "guia": guia_edit.text().strip(),
            "fatura": fatura_edit.text().strip(),
            "caminho": path_edit.text().strip(),
            "data_entrega": entrega_edit.date().toString("yyyy-MM-dd"),
            "data_documento": doc_edit.date().toString("yyyy-MM-dd"),
            "obs": obs_edit.text().strip(),
            "apply_to_lines": bool(apply_lines_chk.isChecked()) if allow_line_apply else False,
        }
        if not any(str(payload.get(key, "") or "").strip() for key in ("titulo", "guia", "fatura", "caminho", "obs")):
            QMessageBox.warning(dialog, "Notas Encomenda", "Indica pelo menos titulo, guia, fatura, caminho ou observacao.")
            return None
        return payload

    def _attach_document(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            detail = self.backend.ne_detail(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        default_type = "FATURA" if str(detail.get("guia_ultima", "") or "").strip() and not str(detail.get("fatura_ultima", "") or "").strip() else "GUIA"
        payload = self._document_dialog(
            f"Associar documento {self.current_number}",
            {
                "tipo": default_type,
                "guia": str(detail.get("guia_ultima", "") or "").strip() if default_type == "GUIA" else "",
                "fatura": "",
                "caminho": str(detail.get("fatura_caminho_ultima", "") or "").strip() if default_type == "FATURA" else "",
                "data_entrega": str(detail.get("data_ultima_entrega", "") or detail.get("data_entrega", "") or "").strip(),
                "data_documento": str(detail.get("data_doc_ultima", "") or detail.get("data_entrega", "") or "").strip(),
                "apply_to_lines": True,
            },
            allow_line_apply=True,
        )
        if payload is None:
            return
        try:
            self.backend.ne_add_document(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Associar Docs", str(exc))
            return
        self.refresh()
        QMessageBox.information(self, "Associar Docs", f"Documento registado em {self.current_number}.")

    def _show_documents(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona primeiro uma nota.")
            return
        try:
            docs = [dict(row) for row in list(self.backend.ne_documents(self.current_number) or [])]
        except Exception as exc:
            QMessageBox.critical(self, "Registo Docs", str(exc))
            return
        self.current_documents = docs
        if not docs:
            QMessageBox.information(self, "Registo Docs", "Esta nota ainda não tem documentos associados.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Registo documental {self.current_number}")
        dialog.resize(1120, 520)
        layout = QVBoxLayout(dialog)
        info = QLabel(f"{len(docs)} documento(s) associado(s) a esta nota.")
        info.setProperty("role", "muted")
        layout.addWidget(info)
        table = QTableWidget(len(docs), 8)
        table.setHorizontalHeaderLabels(["Tipo", "Titulo", "Guia", "Fatura", "Data Doc", "Data Entrega", "Registo", "Caminho"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        _configure_table(table, stretch=(1, 7), contents=(0, 2, 3, 4, 5, 6))
        _fill_table(
            table,
            [
                [
                    str(doc.get("tipo_label", "") or doc.get("tipo", "") or "-"),
                    str(doc.get("titulo", "") or "-"),
                    str(doc.get("guia", "") or "-"),
                    str(doc.get("fatura", "") or "-"),
                    str(doc.get("data_documento", "") or "-"),
                    str(doc.get("data_entrega", "") or "-"),
                    str(doc.get("data_registo", "") or "-").replace("T", " ")[:19],
                    str(doc.get("caminho", "") or "-"),
                ]
                for doc in docs
            ],
            align_center_from=2,
        )
        layout.addWidget(table, 1)

        def open_selected() -> None:
            current = table.currentItem()
            if current is None:
                QMessageBox.warning(dialog, "Registo Docs", "Seleciona um registo.")
                return
            row_index = current.row()
            if row_index < 0 or row_index >= len(docs):
                return
            raw_path = str(docs[row_index].get("caminho", "") or "").strip()
            if not raw_path:
                QMessageBox.information(dialog, "Registo Docs", "Este registo não tem caminho associado.")
                return
            try:
                self.backend.open_file_reference(raw_path)
            except Exception as exc:
                QMessageBox.warning(dialog, "Registo Docs", str(exc))

        actions = QHBoxLayout()
        open_btn = QPushButton("Abrir ficheiro")
        open_btn.setProperty("variant", "secondary")
        open_btn.clicked.connect(open_selected)
        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept)
        actions.addWidget(open_btn)
        actions.addStretch(1)
        actions.addWidget(close_btn)
        layout.addLayout(actions)
        if table.rowCount() > 0:
            table.selectRow(0)
        dialog.exec()

    def _manage_suppliers(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Fornecedores")
        dialog.resize(1040, 680)
        layout = QVBoxLayout(dialog)
        page = SuppliersPage(self.backend, dialog)
        layout.addWidget(page)
        page.refresh()
        close_btn = QDialogButtonBox(QDialogButtonBox.Close)
        close_btn.rejected.connect(dialog.reject)
        close_btn.button(QDialogButtonBox.Close).clicked.connect(dialog.reject)
        layout.addWidget(close_btn)
        dialog.exec()
        self.refresh()

    def _product_line_dialog(self, title: str, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(920)
        dialog.setStyleSheet(
            "QDialog { font-size: 12px; }"
            " QLabel { font-size: 12px; }"
            " QLineEdit, QComboBox, QDoubleSpinBox { font-size: 12px; min-height: 30px; padding: 0 8px; }"
            " QPushButton { min-height: 34px; font-size: 12px; }"
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        intro = QLabel(
            "Seleciona um produto em stock e ajusta a linha comercial. "
            "A base do preço acompanha o tipo do produto para manter o custo coerente."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        product_rows = [dict(row or {}) for row in list(self.backend.ne_product_options("") or [])]
        current_ref = str(initial.get("ref", "") or "").strip()
        current_payload = next(
            (dict(row or {}) for row in product_rows if str(row.get("codigo", "") or "").strip() == current_ref),
            None,
        )
        if current_payload is None and current_ref:
            current_payload = {
                "codigo": current_ref,
                "descricao": str(initial.get("descricao", "") or "").strip(),
                "origem": "Produto",
                "stock": float(initial.get("stock", 0) or 0),
                "unid": str(initial.get("unid", "UN") or "UN").strip() or "UN",
                "preco": float(initial.get("preco", 0) or 0),
                "preco_unid": float(initial.get("preco", 0) or 0),
                "p_compra": float(initial.get("p_compra", initial.get("preco", 0)) or 0),
                "categoria": str(initial.get("categoria", "") or "").strip(),
                "tipo": str(initial.get("tipo", "") or "").strip(),
                "dimensoes": str(initial.get("dimensoes", "") or "").strip(),
                "peso_unid": float(initial.get("peso_unid", 0) or 0),
                "metros_unidade": float(initial.get("metros_unidade", initial.get("metros", 0)) or 0),
                "price_mode": str(initial.get("price_mode", "") or "").strip(),
            }
            product_rows.insert(0, current_payload)

        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.setInsertPolicy(QComboBox.NoInsert)
        product_combo.lineEdit().setPlaceholderText("Selecionar / pesquisar produto")
        product_combo.addItem("", None)

        def _product_label(row: dict[str, Any]) -> str:
            stock_value = float(row.get("stock", 0) or 0)
            return (
                f"{str(row.get('codigo', '') or '').strip()} | "
                f"{str(row.get('descricao', '') or '').strip()} | "
                f"{stock_value:.2f} {str(row.get('unid', 'UN') or 'UN').strip() or 'UN'}"
            ).strip(" |")

        for row in product_rows:
            product_combo.addItem(_product_label(row), row)

        selector_card = CardFrame()
        selector_layout = QVBoxLayout(selector_card)
        selector_layout.setContentsMargins(12, 10, 12, 10)
        selector_layout.setSpacing(8)
        selector_title = QLabel("Seleção de produto em stock")
        selector_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        selector_hint = QLabel("Pesquisa pelo código ou descrição. O stock disponível aparece no seletor e no cartão técnico.")
        selector_hint.setWordWrap(True)
        selector_hint.setProperty("role", "muted")
        selector_layout.addWidget(selector_title)
        selector_layout.addWidget(selector_hint)
        selector_layout.addWidget(product_combo)
        layout.addWidget(selector_card)

        info_card = CardFrame()
        info_card.set_tone("info")
        info_layout = QGridLayout(info_card)
        info_layout.setContentsMargins(12, 10, 12, 10)
        info_layout.setHorizontalSpacing(12)
        info_layout.setVerticalSpacing(8)
        code_display = QLineEdit()
        category_display = QLineEdit()
        dimensions_display = QLineEdit()
        stock_display = QLineEdit()
        for widget in (code_display, category_display, dimensions_display, stock_display):
            widget.setReadOnly(True)
        price_base_label = QLabel("Preço base (EUR/un)")
        price_base_spin = QDoubleSpinBox()
        price_base_spin.setRange(0.0, 1000000.0)
        price_base_spin.setDecimals(4)
        metric_label = QLabel("Métrica por unid.")
        metric_display = QLineEdit()
        metric_display.setReadOnly(True)
        info_layout.addWidget(QLabel("Código produto"), 0, 0)
        info_layout.addWidget(QLabel("Categoria / Tipo"), 0, 1)
        info_layout.addWidget(code_display, 1, 0)
        info_layout.addWidget(category_display, 1, 1)
        info_layout.addWidget(QLabel("Dimensões"), 2, 0)
        info_layout.addWidget(QLabel("Disponível"), 2, 1)
        info_layout.addWidget(dimensions_display, 3, 0)
        info_layout.addWidget(stock_display, 3, 1)
        info_layout.addWidget(price_base_label, 4, 0)
        info_layout.addWidget(metric_label, 4, 1)
        info_layout.addWidget(price_base_spin, 5, 0)
        info_layout.addWidget(metric_display, 5, 1)
        layout.addWidget(info_card)

        desc_edit = QLineEdit(str(initial.get("descricao", "") or ""))
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
        unid_edit = QLineEdit(str(initial.get("unid", "UN") or "UN"))
        preco_spin = QDoubleSpinBox()
        preco_spin.setRange(0.0, 1000000.0)
        preco_spin.setDecimals(4)
        preco_spin.setValue(float(initial.get("preco", 0) or 0))
        desconto_spin = QDoubleSpinBox()
        desconto_spin.setRange(0.0, 100.0)
        desconto_spin.setDecimals(2)
        desconto_spin.setValue(float(initial.get("desconto", 0) or 0))
        iva_spin = QDoubleSpinBox()
        iva_spin.setRange(0.0, 100.0)
        iva_spin.setDecimals(2)
        iva_spin.setValue(float(initial.get("iva", 23) or 23))

        commercial_card = CardFrame()
        commercial_layout = QFormLayout(commercial_card)
        commercial_layout.setContentsMargins(12, 10, 12, 10)
        commercial_layout.setHorizontalSpacing(10)
        commercial_layout.setVerticalSpacing(8)
        commercial_hint = QLabel("A descrição continua editável e podes ajustar o preço da linha sem perder a referência técnica do produto.")
        commercial_hint.setWordWrap(True)
        commercial_hint.setProperty("role", "muted")
        commercial_layout.addRow(commercial_hint)
        commercial_layout.addRow("Descrição da linha", desc_edit)
        commercial_layout.addRow("Fornecedor linha", supplier_combo)
        commercial_layout.addRow("Quantidade", qtd_spin)
        commercial_layout.addRow("Unid.", unid_edit)
        commercial_layout.addRow("Preço unid. (EUR)", preco_spin)
        commercial_layout.addRow("Desc. %", desconto_spin)
        commercial_layout.addRow("IVA %", iva_spin)
        layout.addWidget(commercial_card)

        sync = {"busy": False}
        desc_state = {"manual": False, "last_auto": ""}
        price_state = {"mode": "compra", "metric": 0.0}

        def _product_mode(payload: dict[str, Any]) -> tuple[str, float]:
            category = str(payload.get("categoria", "") or "").strip()
            prod_type = str(payload.get("tipo", "") or "").strip()
            mode = str(payload.get("price_mode", "") or self.backend.desktop_main.produto_modo_preco(category, prod_type) or "compra").strip()
            if mode == "peso":
                metric_value = float(payload.get("peso_unid", 0) or 0)
                return ("peso", metric_value) if metric_value > 0 else ("compra", 0.0)
            if mode == "metros":
                metric_value = float(payload.get("metros_unidade", payload.get("metros", 0)) or 0)
                return ("metros", metric_value) if metric_value > 0 else ("compra", 0.0)
            return "compra", 0.0

        def _base_label(mode: str) -> str:
            if mode == "peso":
                return "Preço kg (EUR/kg)"
            if mode == "metros":
                return "Preço metro (EUR/m)"
            return "Preço base (EUR/un)"

        def _metric_label_text(mode: str) -> str:
            if mode == "peso":
                return "Peso por unid. (kg)"
            if mode == "metros":
                return "Metros por unid. (m)"
            return "Métrica por unid."

        def _metric_text(mode: str, metric_value: float) -> str:
            if mode == "peso":
                return f"{self.backend._fmt(metric_value)} kg"
            if mode == "metros":
                return f"{self.backend._fmt(metric_value)} m"
            return "-"

        def _compute_unit_price(mode: str, metric_value: float, base_value: float) -> float:
            metric_value = max(0.0, float(metric_value or 0))
            if mode in {"peso", "metros"}:
                return round(base_value * metric_value, 4) if metric_value > 0 else 0.0
            return round(base_value, 4)

        def _compute_base_price(mode: str, metric_value: float, unit_value: float) -> float:
            metric_value = max(0.0, float(metric_value or 0))
            if mode in {"peso", "metros"}:
                return round(unit_value / metric_value, 6) if metric_value > 0 else round(unit_value, 6)
            return round(unit_value, 6)

        def _set_metric_visible(visible: bool) -> None:
            metric_label.setVisible(visible)
            metric_display.setVisible(visible)

        def _set_auto_description(text: str, force: bool = False) -> None:
            current_text = desc_edit.text().strip()
            if not force and desc_state["manual"] and current_text and current_text != desc_state["last_auto"]:
                return
            desc_edit.blockSignals(True)
            desc_edit.setText(text)
            desc_edit.blockSignals(False)
            desc_state["last_auto"] = text
            if force:
                desc_state["manual"] = False

        def _clear_product_state() -> None:
            sync["busy"] = True
            try:
                code_display.setText(current_ref if current_ref and current_payload is None else "")
                category_display.setText("-")
                dimensions_display.setText("-")
                stock_display.setText("-")
                price_state["mode"] = "compra"
                price_state["metric"] = 0.0
                price_base_label.setText(_base_label("compra"))
                if "p_compra" in initial:
                    price_base_spin.setValue(float(initial.get("p_compra", 0) or 0))
                else:
                    price_base_spin.setValue(float(initial.get("preco", 0) or 0))
                metric_label.setText(_metric_label_text("compra"))
                metric_display.setText("-")
                _set_metric_visible(False)
            finally:
                sync["busy"] = False

        def _apply_payload(payload: dict[str, Any], *, prefer_existing_prices: bool = False) -> None:
            if not isinstance(payload, dict):
                return
            mode, metric_value = _product_mode(payload)
            unit_value = (
                float(initial.get("preco", payload.get("preco_unid", payload.get("preco", 0))) or 0)
                if prefer_existing_prices
                else float(payload.get("preco_unid", payload.get("preco", 0)) or 0)
            )
            if prefer_existing_prices and "p_compra" in initial:
                base_value = float(initial.get("p_compra", 0) or 0)
            elif prefer_existing_prices:
                base_value = _compute_base_price(mode, metric_value, unit_value)
            else:
                base_value = float(payload.get("p_compra", payload.get("preco_unid", payload.get("preco", 0))) or 0)
            category = str(payload.get("categoria", "") or "").strip()
            prod_type = str(payload.get("tipo", "") or "").strip()
            category_type = " / ".join(part for part in (category, prod_type) if part) or "-"
            sync["busy"] = True
            try:
                code_display.setText(str(payload.get("codigo", "") or "").strip())
                category_display.setText(category_type)
                dimensions_display.setText(str(payload.get("dimensoes", "") or "-").strip() or "-")
                stock_display.setText(
                    f"{float(payload.get('stock', 0) or 0):.2f} {str(payload.get('unid', 'UN') or 'UN').strip() or 'UN'}"
                )
                unid_edit.setText(str(payload.get("unid", "UN") or "UN").strip() or "UN")
                price_state["mode"] = mode
                price_state["metric"] = metric_value
                price_base_label.setText(_base_label(mode))
                price_base_spin.setValue(base_value)
                metric_label.setText(_metric_label_text(mode))
                metric_display.setText(_metric_text(mode, metric_value))
                _set_metric_visible(mode in {"peso", "metros"})
                preco_spin.setValue(unit_value)
                auto_desc = str(payload.get("descricao", "") or "").strip()
                if prefer_existing_prices and str(initial.get("descricao", "") or "").strip():
                    _set_auto_description(str(initial.get("descricao", "") or "").strip(), force=True)
                else:
                    _set_auto_description(auto_desc, force=True)
            finally:
                sync["busy"] = False

        def _on_product_change() -> None:
            payload = product_combo.currentData()
            if isinstance(payload, dict):
                selected_code = str(payload.get("codigo", "") or "").strip()
                _apply_payload(payload, prefer_existing_prices=bool(current_ref and selected_code == current_ref))
            else:
                _clear_product_state()

        def _on_base_price_changed(value: float) -> None:
            if sync["busy"]:
                return
            sync["busy"] = True
            try:
                preco_spin.setValue(_compute_unit_price(str(price_state["mode"]), float(price_state["metric"]), value))
            finally:
                sync["busy"] = False

        def _on_unit_price_changed(value: float) -> None:
            if sync["busy"]:
                return
            sync["busy"] = True
            try:
                price_base_spin.setValue(_compute_base_price(str(price_state["mode"]), float(price_state["metric"]), value))
            finally:
                sync["busy"] = False

        def _on_description_edited(_text: str) -> None:
            if sync["busy"]:
                return
            desc_state["manual"] = True

        def _accept_dialog() -> None:
            if not isinstance(product_combo.currentData(), dict):
                QMessageBox.warning(dialog, "Notas Encomenda", "Seleciona um produto em stock.")
                return
            if not desc_edit.text().strip():
                QMessageBox.warning(dialog, "Notas Encomenda", "Indica a descrição da linha.")
                return
            dialog.accept()

        product_combo.currentIndexChanged.connect(lambda _i: _on_product_change())
        price_base_spin.valueChanged.connect(_on_base_price_changed)
        preco_spin.valueChanged.connect(_on_unit_price_changed)
        desc_edit.textEdited.connect(_on_description_edited)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(_accept_dialog)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if current_ref:
            for index in range(product_combo.count()):
                payload = product_combo.itemData(index)
                if isinstance(payload, dict) and str(payload.get("codigo", "") or "").strip() == current_ref:
                    product_combo.setCurrentIndex(index)
                    break
        else:
            product_combo.setCurrentIndex(0)
            _clear_product_state()
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = product_combo.currentData()
        if not isinstance(payload, dict):
            return None
        mode = str(price_state.get("mode", "compra") or "compra").strip() or "compra"
        metric_value = float(price_state.get("metric", 0) or 0)
        return {
            "ref": str(payload.get("codigo", "") or "").strip(),
            "descricao": desc_edit.text().strip(),
            "fornecedor_linha": supplier_combo.currentText().strip(),
            "origem": "Produto",
            "qtd": qtd_spin.value(),
            "unid": unid_edit.text().strip() or "UN",
            "preco": preco_spin.value(),
            "desconto": desconto_spin.value(),
            "iva": iva_spin.value(),
            "entrega": str(initial.get("entrega", "PENDENTE") or "PENDENTE"),
            "categoria": str(payload.get("categoria", "") or "").strip(),
            "tipo": str(payload.get("tipo", "") or "").strip(),
            "dimensoes": str(payload.get("dimensoes", "") or "").strip(),
            "p_compra": float(price_base_spin.value() or 0),
            "peso_unid": metric_value if mode == "peso" else 0.0,
            "metros_unidade": metric_value if mode == "metros" else 0.0,
            "price_basis": "kg" if mode == "peso" else ("meter" if mode == "metros" else "unit"),
        }

    def _delivery_dialog(self) -> dict | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Entrega {self.current_number}")
        dialog.setMinimumSize(900, 620)
        dialog.resize(980, 700)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setSizeGripEnabled(True)
        dialog_layout = QVBoxLayout(dialog)
        dialog_layout.setContentsMargins(12, 10, 12, 10)
        dialog_layout.setSpacing(10)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        content = QWidget()
        content.setMaximumWidth(940)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(10)
        scroll.setWidget(content)
        dialog_layout.addWidget(scroll, 1)
        entrega_edit = QDateEdit()
        entrega_edit.setCalendarPopup(True)
        entrega_edit.setDisplayFormat("yyyy-MM-dd")
        entrega_edit.setDate(QDate.currentDate())
        doc_edit = QDateEdit()
        doc_edit.setCalendarPopup(True)
        doc_edit.setDisplayFormat("yyyy-MM-dd")
        doc_edit.setDate(QDate.currentDate())
        title_edit = QLineEdit()
        guia_edit = QLineEdit()
        fatura_edit = QLineEdit()
        path_edit = QLineEdit()
        obs_edit = QLineEdit()
        for field in (entrega_edit, doc_edit, title_edit, guia_edit, fatura_edit, path_edit, obs_edit):
            field.setMinimumHeight(34)
        entrega_edit.setMinimumWidth(170)
        doc_edit.setMinimumWidth(170)
        _cap_width(entrega_edit, 190)
        _cap_width(doc_edit, 190)
        path_edit.setPlaceholderText("Opcional")
        title_edit.setPlaceholderText("Opcional")
        guia_edit.setPlaceholderText("Opcional")
        fatura_edit.setPlaceholderText("Opcional")
        obs_edit.setPlaceholderText("Observações da entrega (opcional)")

        def pick_path() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar documento de entrega",
                "",
                "Documentos (*.pdf *.png *.jpg *.jpeg *.bmp *.xlsx *.xls *.doc *.docx);;Todos (*.*)",
            )
            if path:
                path_edit.setText(path)

        browse_btn = QPushButton("Selecionar ficheiro")
        browse_btn.setProperty("variant", "secondary")
        browse_btn.clicked.connect(pick_path)
        _cap_width(browse_btn, 180)

        meta_card = CardFrame()
        meta_card.set_tone("info")
        meta_layout = QGridLayout(meta_card)
        meta_layout.setContentsMargins(14, 12, 14, 12)
        meta_layout.setHorizontalSpacing(12)
        meta_layout.setVerticalSpacing(8)
        meta_title = QLabel("Documento de entrega")
        meta_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        meta_hint = QLabel(
            "Preenche apenas o essencial. O lote da entrada e a localização são opcionais nas linhas abaixo."
        )
        meta_hint.setProperty("role", "muted")
        meta_hint.setWordWrap(True)
        path_row = QWidget()
        path_row_layout = QHBoxLayout(path_row)
        path_row_layout.setContentsMargins(0, 0, 0, 0)
        path_row_layout.setSpacing(8)
        path_row_layout.addWidget(path_edit, 1)
        path_row_layout.addWidget(browse_btn, 0)
        meta_layout.addWidget(meta_title, 0, 0, 1, 4)
        meta_layout.addWidget(meta_hint, 1, 0, 1, 4)
        label_entrega = QLabel("Data entrega")
        label_doc = QLabel("Data documento")
        label_titulo = QLabel("Título")
        label_guia = QLabel("Guia")
        label_fatura = QLabel("Fatura")
        label_path = QLabel("Documento")
        label_obs = QLabel("Obs.")
        for label in (label_entrega, label_doc, label_titulo, label_guia, label_fatura, label_path, label_obs):
            label.setProperty("role", "muted")
        meta_layout.addWidget(label_entrega, 2, 0)
        meta_layout.addWidget(label_doc, 2, 2)
        meta_layout.addWidget(entrega_edit, 3, 0, 1, 2)
        meta_layout.addWidget(doc_edit, 3, 2, 1, 2)
        meta_layout.addWidget(label_titulo, 4, 0)
        meta_layout.addWidget(label_guia, 4, 2)
        meta_layout.addWidget(title_edit, 5, 0, 1, 2)
        meta_layout.addWidget(guia_edit, 5, 2, 1, 2)
        meta_layout.addWidget(label_fatura, 6, 0)
        meta_layout.addWidget(label_path, 6, 2)
        meta_layout.addWidget(fatura_edit, 7, 0, 1, 2)
        meta_layout.addWidget(path_row, 7, 2, 1, 2)
        meta_layout.addWidget(label_obs, 8, 0)
        meta_layout.addWidget(obs_edit, 9, 0, 1, 4)
        for col in range(4):
            meta_layout.setColumnStretch(col, 1)
        layout.addWidget(meta_card)

        table = QTableWidget(len(self.line_rows), 6)
        table.setHorizontalHeaderLabels(["Ref", "Descrição", "Pendente", "Receber", "Lote (opc.)", "Localização (opc.)"])
        table.verticalHeader().setVisible(False)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setDefaultSectionSize(34)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        table.setColumnWidth(0, 110)
        table.setColumnWidth(2, 92)
        table.setColumnWidth(3, 98)
        table.setColumnWidth(4, 138)
        table.setColumnWidth(5, 162)
        table.setMinimumHeight(max(220, min(380, _table_visible_height(table, min(len(self.line_rows), 7), extra=12))))
        qty_inputs: list[QDoubleSpinBox] = []
        lote_inputs: list[QLineEdit] = []
        local_inputs: list[QComboBox] = []
        presets = self.backend.material_presets()
        location_options = [
            str(value or "").strip()
            for value in list(dict.fromkeys(list(presets.get("locais", []) or [])))
            if str(value or "").strip()
        ]
        for row_index, row in enumerate(self.line_rows):
            ref = str(row.get("ref", "") or "").strip()
            desc = str(row.get("descricao", "") or "").strip()
            entrega_txt = str(row.get("entrega", "PENDENTE") or "PENDENTE").upper()
            qtd_total = float(row.get("qtd", 0) or 0)
            pending = qtd_total
            if entrega_txt.startswith("PARCIAL"):
                try:
                    partial_txt = entrega_txt.split("(", 1)[1].split("/", 1)[0]
                    pending = max(0.0, qtd_total - float(str(partial_txt).replace(",", ".")))
                except Exception:
                    pending = qtd_total
            elif "ENTREGUE" in entrega_txt:
                pending = 0.0
            table.setItem(row_index, 0, QTableWidgetItem(ref))
            table.setItem(row_index, 1, QTableWidgetItem(desc))
            pending_item = QTableWidgetItem(f"{pending:.2f}")
            pending_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, 2, pending_item)
            qty_spin = QDoubleSpinBox()
            qty_spin.setRange(0.0, pending)
            qty_spin.setDecimals(2)
            qty_spin.setValue(pending if pending > 0 else 0.0)
            qty_spin.setMinimumWidth(112)
            qty_spin.setMinimumHeight(30)
            qty_spin.setAlignment(Qt.AlignCenter)
            qty_inputs.append(qty_spin)
            table.setCellWidget(row_index, 3, qty_spin)
            is_material_line = self.backend.desktop_main.origem_is_materia(row.get("origem", ""))
            needs_material_entry = (
                is_material_line
                and (
                    bool(row.get("_material_pending_create"))
                    or not ref
                    or self.backend.material_by_id(ref) is None
                )
            )
            lote_edit = QLineEdit(str(row.get("lote_fornecedor", "") or "").strip())
            lote_edit.setMinimumHeight(30)
            lote_edit.setPlaceholderText("Opcional nesta entrada")
            lote_edit.setEnabled(needs_material_entry)
            if not needs_material_entry and not lote_edit.text().strip():
                lote_edit.setText(str(row.get("lote_fornecedor", "") or "").strip())
            lote_inputs.append(lote_edit)
            table.setCellWidget(row_index, 4, lote_edit)
            local_combo = QComboBox()
            local_combo.setEditable(True)
            local_combo.setInsertPolicy(QComboBox.NoInsert)
            local_combo.addItem("")
            for option in location_options:
                local_combo.addItem(option)
            local_combo.setCurrentText(str(row.get("localizacao", "") or "").strip())
            local_combo.setMinimumHeight(30)
            local_combo.setEnabled(is_material_line)
            local_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            local_combo.setMinimumContentsLength(12)
            if local_combo.lineEdit() is not None:
                local_combo.lineEdit().setPlaceholderText("Selecionar / escrever")
            local_inputs.append(local_combo)
            table.setCellWidget(row_index, 5, local_combo)
        lines_card = CardFrame()
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(14, 12, 14, 12)
        lines_layout.setSpacing(8)
        lines_title = QLabel("Linhas a receber")
        lines_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        lines_hint = QLabel(
            "Nas linhas de matéria-prima, a entrada cria o registo de stock quando ele ainda não existe. "
            "O lote e a localização podem ser definidos agora ou ficar em branco."
        )
        lines_hint.setProperty("role", "muted")
        lines_hint.setWordWrap(True)
        lines_layout.addWidget(lines_title)
        lines_layout.addWidget(lines_hint)
        lines_layout.addWidget(table)
        layout.addWidget(lines_card, 1)
        layout.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn is not None:
            ok_btn.setText("Guardar")
        if cancel_btn is not None:
            cancel_btn.setText("Fechar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        dialog_layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        lines = []
        for row_index, row in enumerate(self.line_rows):
            value = qty_inputs[row_index].value()
            if value > 0:
                needs_material_entry = (
                    self.backend.desktop_main.origem_is_materia(row.get("origem", ""))
                    and (
                        bool(row.get("_material_pending_create"))
                        or not str(row.get("ref", "") or "").strip()
                        or self.backend.material_by_id(str(row.get("ref", "") or "").strip()) is None
                    )
                )
                lote_txt = lote_inputs[row_index].text().strip()
                lines.append(
                    {
                        "index": row_index,
                        "ref": str(row.get("ref", "") or "").strip(),
                        "qtd": value,
                        "lote_fornecedor": lote_txt,
                        "localizacao": local_inputs[row_index].currentText().strip(),
                    }
                )
        return {
            "data_entrega": entrega_edit.date().toString("yyyy-MM-dd"),
            "data_documento": doc_edit.date().toString("yyyy-MM-dd"),
            "titulo": title_edit.text().strip(),
            "guia": guia_edit.text().strip(),
            "fatura": fatura_edit.text().strip(),
            "caminho": path_edit.text().strip(),
            "obs": obs_edit.text().strip(),
            "lines": lines,
        }

    def _create_new_note(self) -> None:
        try:
            draft = self.backend.ne_create_draft()
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        self.current_number = str(draft.get("numero", "") or "").strip()
        self.refresh()
        if hasattr(self, "view_stack"):
            self._show_note_detail()

    def _new_note(self, reset_number: bool = True) -> None:
        self.notes_table.blockSignals(True)
        self.notes_table.clearSelection()
        self.notes_table.blockSignals(False)
        self.current_number = self.backend.ne_next_number() if reset_number else self.current_number or self.backend.ne_next_number()
        self.current_documents = []
        self.line_rows = []
        self._set_header(self.current_number, "Em edicao")
        self.summary_hint.setText("Pedido")
        self.generate_btn.setEnabled(True)
        self.delivery_btn.setEnabled(False)
        self.attach_doc_btn.setEnabled(False)
        self.documents_btn.setEnabled(False)
        self.pdf_btn.setEnabled(True)
        self.quote_btn.setEnabled(True)
        self.quote_btn.setToolTip("Gera o PDF de pedido de orçamento/cotação para a nota atual.")
        self.supplier_combo.setCurrentText("")
        self.contact_edit.clear()
        self._set_delivery_text("")
        self.location_edit.setCurrentText("")
        self.transport_edit.setCurrentText("")
        self.obs_edit.clear()
        self._render_lines()

    def _load_selected_note(self) -> None:
        row = self._selected_row()
        numero = str(row.get("numero", "")).strip()
        if not numero:
            return
        try:
            detail = self.backend.ne_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        self._set_header(detail.get("numero", ""), detail.get("estado", "Em edicao"))
        self.supplier_combo.setCurrentText(str(detail.get("fornecedor", "") or "").strip())
        self.contact_edit.setText(str(detail.get("contacto", "") or "").strip())
        self._set_delivery_text(str(detail.get("data_entrega", "") or "").strip())
        self.location_edit.setCurrentText(str(detail.get("local_descarga", "") or "").strip())
        self.transport_edit.setCurrentText(str(detail.get("meio_transporte", "") or "").strip())
        self.obs_edit.setPlainText(str(detail.get("obs", "") or "").strip())
        self.current_documents = [dict(row) for row in list(detail.get("documents", []) or [])]
        self.line_rows = [dict(line) for line in list(detail.get("lines", []) or [])]
        kind = str(detail.get("kind", "") or "").strip()
        generated = list(detail.get("ne_geradas", []) or [])
        doc_count = int(detail.get("document_count", len(self.current_documents)) or 0)
        if kind == "rfq":
            text = "Cotação"
            if generated:
                text = f"{text} | NEs {len(generated)}"
            if doc_count:
                text = f"{text} | Docs {doc_count}"
            self.summary_hint.setText(text)
            self.generate_btn.setEnabled(True)
            self.delivery_btn.setEnabled(False)
            self.attach_doc_btn.setEnabled(False)
            self.documents_btn.setEnabled(doc_count > 0)
            self.pdf_btn.setEnabled(False)
            self.quote_btn.setEnabled(True)
            self.quote_btn.setToolTip("Gera o PDF de pedido de orçamento/cotação para a nota atual.")
        elif kind == "supplier_order":
            summary = "NE gerada"
            if doc_count:
                summary = f"{summary} | Docs {doc_count}"
            self.summary_hint.setText(summary)
            self.generate_btn.setEnabled(False)
            self.delivery_btn.setEnabled(True)
            self.attach_doc_btn.setEnabled(True)
            self.documents_btn.setEnabled(True)
            self.pdf_btn.setEnabled(True)
            self.quote_btn.setEnabled(False)
            self.quote_btn.setToolTip("As NEs já geradas por fornecedor usam o botão 'Pre-visualizar NE'.")
        else:
            summary = "Compra direta / cotação simples"
            if doc_count:
                summary = f"{summary} | Docs {doc_count}"
            self.summary_hint.setText(summary)
            self.generate_btn.setEnabled(False)
            self.delivery_btn.setEnabled(True)
            self.attach_doc_btn.setEnabled(True)
            self.documents_btn.setEnabled(True)
            self.pdf_btn.setEnabled(True)
            self.quote_btn.setEnabled(True)
            self.quote_btn.setToolTip("Também podes emitir um pedido de orçamento para uma nota com fornecedor único.")
        self._render_lines()

    def _line_dialog(self, title: str, initial: dict | None = None, material_mode: bool = False) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(980 if material_mode else 760)
        dialog.resize(1080 if material_mode else 820, 760 if material_mode else 620)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setSizeGripEnabled(True)
        dialog.setStyleSheet(
            "QDialog { font-size: 12px; }"
            " QLabel { font-size: 12px; }"
            " QLineEdit, QComboBox, QDoubleSpinBox { font-size: 12px; min-height: 30px; padding: 0 8px; }"
            " QPushButton { min-height: 34px; font-size: 12px; }"
        )
        dialog_layout = QVBoxLayout(dialog)
        dialog_layout.setContentsMargins(12, 10, 12, 10)
        dialog_layout.setSpacing(10)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        content = QWidget()
        content.setMaximumWidth(1040 if material_mode else 780)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)
        scroll.setWidget(content)
        dialog_layout.addWidget(scroll, 1)
        ref_edit = QLineEdit(str(initial.get("ref", "") or ""))
        desc_edit = QLineEdit(str(initial.get("descricao", "") or ""))
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
        qtd_spin = QDoubleSpinBox()
        qtd_spin.setRange(0.0, 1000000.0)
        qtd_spin.setDecimals(2)
        qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
        unid_edit = QLineEdit(str(initial.get("unid", "UN") or "UN"))
        preco_spin = QDoubleSpinBox()
        preco_spin.setRange(0.0, 1000000.0)
        preco_spin.setDecimals(4)
        preco_spin.setValue(float(initial.get("preco", 0) or 0))
        desconto_spin = QDoubleSpinBox()
        desconto_spin.setRange(0.0, 100.0)
        desconto_spin.setDecimals(2)
        desconto_spin.setValue(float(initial.get("desconto", 0) or 0))
        iva_spin = QDoubleSpinBox()
        iva_spin.setRange(0.0, 100.0)
        iva_spin.setDecimals(2)
        iva_spin.setValue(float(initial.get("iva", 23) or 23))
        accept_handler = dialog.accept
        manual_price_controls = None
        if material_mode:
            intro = QLabel(
                "Escolhe stock existente pelos filtros ou cria uma linha manual. "
                "O código é sempre o ID do material; sem stock associado, o ID fica pendente até à entrada."
            )
            intro.setWordWrap(True)
            intro.setProperty("role", "muted")
            layout.addWidget(intro)
        if material_mode:
            material_rows = [dict(row or {}) for row in list(self.backend.ne_material_options("") or [])]
            current_ref = str(initial.get("ref", "") or "").strip()
            initial_has_manual_meta = any(
                str(initial.get(key, "") or "").strip()
                for key in ("material", "espessura", "formato", "lote_fornecedor", "localizacao", "material_familia", "secao_tipo")
            ) or any(float(initial.get(key, 0) or 0) > 0 for key in ("comprimento", "largura", "altura", "diametro", "metros", "kg_m", "peso_unid", "p_compra"))
            current_payload = next(
                (dict(row or {}) for row in material_rows if str(row.get("id", "") or "").strip() == current_ref),
                None,
            )
            if current_payload is None and (current_ref or initial_has_manual_meta):
                current_payload = {
                    "id": current_ref,
                    "descricao": str(initial.get("descricao", "") or "").strip(),
                    "material": str(initial.get("material", "") or "").strip(),
                    "material_familia": str(initial.get("material_familia", "") or "").strip(),
                    "espessura": str(initial.get("espessura", "") or "").strip(),
                    "formato": str(initial.get("formato", "") or "Chapa").strip() or "Chapa",
                    "preco": float(initial.get("preco", 0) or 0),
                    "preco_base": float(initial.get("p_compra", 0) or 0),
                    "preco_base_label": "EUR/m"
                    if str(initial.get("formato", "") or "").strip() == "Tubo"
                    else "EUR/kg",
                    "unid": str(initial.get("unid", "UN") or "UN"),
                    "lote": str(initial.get("lote_fornecedor", "") or "").strip(),
                    "localizacao": str(initial.get("localizacao", "") or "").strip(),
                    "comprimento": float(initial.get("comprimento", 0) or 0),
                    "largura": float(initial.get("largura", 0) or 0),
                    "altura": float(initial.get("altura", 0) or 0),
                    "diametro": float(initial.get("diametro", 0) or 0),
                    "metros": float(initial.get("metros", 0) or 0),
                    "kg_m": float(initial.get("kg_m", 0) or 0),
                    "peso_unid": float(initial.get("peso_unid", 0) or 0),
                    "secao_tipo": str(initial.get("secao_tipo", "") or "").strip(),
                }
                preview = dict(self.backend.material_price_preview(current_payload) or {})
                current_payload["dimensao"] = str(preview.get("dimension_label", "") or "").strip() or "-"
                current_payload["secao_tipo"] = str(preview.get("secao_tipo", current_payload.get("secao_tipo", "")) or "").strip()
                current_payload["kg_m"] = float(preview.get("kg_m", current_payload.get("kg_m", 0)) or 0)
                current_payload["peso_unid"] = float(preview.get("peso_unid", current_payload.get("peso_unid", 0)) or 0)
                material_rows.insert(0, current_payload)

            def _clean(value: Any, fallback: str = "") -> str:
                text = str(value or "").strip()
                return text or fallback

            def _float_text(value: Any, digits: int = 3) -> str:
                try:
                    number = float(value or 0)
                except Exception:
                    number = 0.0
                if digits <= 0:
                    text = f"{int(round(number))}"
                else:
                    text = f"{number:.{digits}f}".rstrip("0").rstrip(".")
                return text.replace(".", ",")

            def _material_family_options() -> list[dict[str, Any]]:
                try:
                    return [dict(row or {}) for row in list(self.backend.material_family_options() or [])]
                except Exception:
                    return [
                        {"key": "", "label": "Auto", "density": 0.0},
                        {"key": "steel", "label": "Aço / Ferro", "density": 7.85},
                        {"key": "stainless", "label": "Inox", "density": 7.93},
                        {"key": "aluminum", "label": "Alumínio", "density": 2.70},
                        {"key": "brass", "label": "Latão", "density": 8.50},
                        {"key": "copper", "label": "Cobre", "density": 8.96},
                    ]

            def _material_family_profile(material_txt: Any = "", family_txt: Any = "") -> dict[str, Any]:
                try:
                    return dict(self.backend.material_family_profile(material_txt, family_txt) or {})
                except Exception:
                    return {"key": "", "label": "Auto", "density": 7.85, "explicit": False}

            def _dimension_text(row: dict[str, Any]) -> str:
                try:
                    preview = dict(self.backend.material_geometry_preview(row) or {})
                except Exception:
                    preview = {}
                return _clean(preview.get("dimension_label", row.get("dimensao", "")), "-")

            def _norm_material_text(value: Any) -> str:
                helper = getattr(self.backend, "_norm_material_token", None)
                if callable(helper):
                    try:
                        return str(helper(value) or "")
                    except Exception:
                        pass
                return re.sub(r"\s+", "", str(value or "").strip().lower())

            def _norm_esp_text(value: Any) -> str:
                helper = getattr(self.backend, "_norm_esp_token", None)
                if callable(helper):
                    try:
                        return str(helper(value) or "")
                    except Exception:
                        pass
                txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
                txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
                return txt

            def _norm_dimension_text(formato_txt: str, dim_txt: Any) -> str:
                raw = str(dim_txt or "").strip().upper()
                if not raw:
                    return ""
                raw = raw.replace("×", "X").replace("MM", "").replace(",", ".")
                raw = re.sub(r"\s+", "", raw)
                raw = raw.replace("Ø", "D")
                tokens = re.findall(r"[A-Z]+|[-+]?\d+(?:\.\d+)?", raw)
                normalized: list[str] = []
                for token in tokens:
                    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", token):
                        try:
                            normalized.append(f"{float(token):.4f}".rstrip("0").rstrip("."))
                        except Exception:
                            normalized.append(token)
                    else:
                        normalized.append(token)
                return "|".join(normalized)

            def _matches(row: dict[str, Any], formato_txt: str, quality_txt: str, esp_txt: str, dim_txt: str) -> bool:
                row_formato = _clean(row.get("formato", ""), "Chapa")
                if formato_txt and row_formato.lower() != formato_txt.lower():
                    return False
                if quality_txt and _norm_material_text(row.get("material", "")) != _norm_material_text(quality_txt):
                    return False
                row_esp = _clean(row.get("espessura", ""), "-")
                if esp_txt and esp_txt != "-" and _norm_esp_text(row_esp) != _norm_esp_text(esp_txt):
                    return False
                if dim_txt and _norm_dimension_text(row_formato, _dimension_text(row)) != _norm_dimension_text(row_formato, dim_txt):
                    return False
                return True

            def _unique(rows: list[dict[str, Any]], key_fn) -> list[str]:
                values: list[str] = []
                for row in rows:
                    value = _clean(key_fn(row))
                    if not value:
                        continue
                    if value not in values:
                        values.append(value)
                return values

            selector_card = CardFrame()
            selector_layout = QGridLayout(selector_card)
            selector_layout.setContentsMargins(12, 10, 12, 10)
            selector_layout.setHorizontalSpacing(10)
            selector_layout.setVerticalSpacing(8)
            selector_title = QLabel("Seleção de matéria-prima")
            selector_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
            selector_hint = QLabel(
                "Começa vazio. Filtra por tipo, qualidade, espessura e dimensão. "
                "O stock é identificado pelo ID do material; o lote aparece apenas como referência adicional. "
                "Quando a qualidade é escrita à mão, escolhe a família para acertar a densidade e o peso."
            )
            selector_hint.setProperty("role", "muted")
            selector_hint.setWordWrap(True)
            common_sheet_dimensions = [
                "3000 x 1500 mm",
                "4000 x 2000 mm",
                "2000 x 1000 mm",
                "2500 x 1250 mm",
            ]
            formato_combo = QComboBox()
            qualidade_combo = QComboBox()
            familia_combo = QComboBox()
            espessura_combo = QComboBox()
            dimensao_combo = QComboBox()
            stock_combo = QComboBox()
            for combo, placeholder in (
                (formato_combo, "Selecionar / escrever tipo"),
                (qualidade_combo, "Selecionar / escrever qualidade"),
                (espessura_combo, "Selecionar / escrever espessura"),
                (dimensao_combo, "Selecionar / escrever dimensões"),
            ):
                combo.setEditable(True)
                combo.setInsertPolicy(QComboBox.NoInsert)
                combo.lineEdit().setPlaceholderText(placeholder)
            stock_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            stock_combo.setMinimumContentsLength(24)
            selector_layout.addWidget(selector_title, 0, 0, 1, 4)
            selector_layout.addWidget(selector_hint, 1, 0, 1, 4)
            selector_layout.addWidget(QLabel("Tipo material"), 2, 0)
            selector_layout.addWidget(QLabel("Qualidade"), 2, 1)
            selector_layout.addWidget(QLabel("Espessura"), 2, 2)
            selector_layout.addWidget(QLabel("Dimensões"), 2, 3)
            selector_layout.addWidget(formato_combo, 3, 0)
            selector_layout.addWidget(qualidade_combo, 3, 1)
            selector_layout.addWidget(espessura_combo, 3, 2)
            selector_layout.addWidget(dimensao_combo, 3, 3)
            selector_layout.addWidget(QLabel("Família material"), 4, 0)
            selector_layout.addWidget(QLabel("Densidade (g/cm3)"), 4, 1)
            selector_layout.addWidget(QLabel("Stock disponível (ID material)"), 4, 2, 1, 2)
            densidade_display = QLineEdit()
            densidade_display.setReadOnly(True)
            for option in _material_family_options():
                familia_combo.addItem(str(option.get("label", "") or ""), str(option.get("key", "") or ""))
            selector_layout.addWidget(familia_combo, 5, 0)
            selector_layout.addWidget(densidade_display, 5, 1)
            selector_layout.addWidget(stock_combo, 5, 2, 1, 2)
            for col in range(4):
                selector_layout.setColumnStretch(col, 1)
            layout.addWidget(selector_card)

            info_card = CardFrame()
            info_card.set_tone("info")
            info_layout = QGridLayout(info_card)
            info_layout.setContentsMargins(12, 10, 12, 10)
            info_layout.setHorizontalSpacing(12)
            info_layout.setVerticalSpacing(8)
            code_display = QLineEdit()
            code_display.setReadOnly(True)
            local_display = QLineEdit()
            lote_display = QLineEdit()
            desc_edit.setReadOnly(False)
            price_base_label = QLabel("Preço base")
            price_base_spin = QDoubleSpinBox()
            price_base_spin.setRange(0.0, 1000000.0)
            price_base_spin.setDecimals(4)
            metric_label = QLabel("Peso por unid. (kg)")
            metric_spin = QDoubleSpinBox()
            metric_spin.setRange(0.0, 1000000.0)
            metric_spin.setDecimals(4)
            metric_spin.setReadOnly(True)
            code_display.setPlaceholderText("Gerado na entrada")
            local_display.setPlaceholderText("Opcional")
            lote_display.setPlaceholderText("Opcional")
            info_layout.addWidget(QLabel("ID material"), 0, 0)
            info_layout.addWidget(QLabel("Localização"), 0, 1)
            info_layout.addWidget(QLabel("Lote fornecedor"), 0, 2)
            info_layout.addWidget(price_base_label, 0, 3)
            info_layout.addWidget(code_display, 1, 0)
            info_layout.addWidget(local_display, 1, 1)
            info_layout.addWidget(lote_display, 1, 2)
            info_layout.addWidget(price_base_spin, 1, 3)
            for col in range(4):
                info_layout.setColumnStretch(col, 1)
            layout.addWidget(info_card)

            tech_card = CardFrame()
            tech_layout = QGridLayout(tech_card)
            tech_layout.setContentsMargins(12, 10, 12, 10)
            tech_layout.setHorizontalSpacing(12)
            tech_layout.setVerticalSpacing(8)
            tech_title = QLabel("Dados técnicos")
            tech_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
            tech_hint = QLabel("Os mesmos campos técnicos da matéria-prima são usados aqui para garantir peso e stock coerentes.")
            tech_hint.setProperty("role", "muted")
            tech_hint.setWordWrap(True)
            secao_tipo_combo = QComboBox()
            secao_tipo_combo.setEditable(True)
            secao_tipo_combo.setInsertPolicy(QComboBox.NoInsert)
            secao_tipo_combo.lineEdit().setPlaceholderText("Selecionar tipo / série")
            comprimento_edit = QLineEdit()
            largura_edit = QLineEdit()
            altura_edit = QLineEdit()
            diametro_edit = QLineEdit()
            metros_edit = QLineEdit()
            kg_m_edit = QLineEdit()
            kg_m_edit.setPlaceholderText("Auto por tabela / fórmula")
            tech_layout.addWidget(tech_title, 0, 0, 1, 4)
            tech_layout.addWidget(tech_hint, 1, 0, 1, 4)
            secao_label = QLabel("Tipo secção")
            comp_label = QLabel("Comprimento (mm)")
            larg_label = QLabel("Largura (mm)")
            altura_label = QLabel("Altura / tamanho (mm)")
            diametro_label = QLabel("Diâmetro ext. (mm)")
            metros_label = QLabel("Comprimento barra (m)")
            kgm_label = QLabel("Peso por metro (kg/m)")
            for label in (secao_label, comp_label, larg_label, altura_label, diametro_label, metros_label, kgm_label):
                label.setProperty("role", "muted")
            tech_layout.addWidget(secao_label, 2, 0)
            tech_layout.addWidget(secao_tipo_combo, 2, 1)
            tech_layout.addWidget(comp_label, 2, 2)
            tech_layout.addWidget(comprimento_edit, 2, 3)
            tech_layout.addWidget(larg_label, 3, 0)
            tech_layout.addWidget(largura_edit, 3, 1)
            tech_layout.addWidget(altura_label, 3, 2)
            tech_layout.addWidget(altura_edit, 3, 3)
            tech_layout.addWidget(diametro_label, 4, 0)
            tech_layout.addWidget(diametro_edit, 4, 1)
            tech_layout.addWidget(metros_label, 4, 2)
            tech_layout.addWidget(metros_edit, 4, 3)
            tech_layout.addWidget(kgm_label, 5, 0)
            tech_layout.addWidget(kg_m_edit, 5, 1)
            tech_layout.addWidget(metric_label, 5, 2)
            tech_layout.addWidget(metric_spin, 5, 3)
            for col in range(4):
                tech_layout.setColumnStretch(col, 1)
            layout.addWidget(tech_card)
            commercial_card = CardFrame()
            commercial_layout = QGridLayout(commercial_card)
            commercial_layout.setContentsMargins(12, 10, 12, 10)
            commercial_layout.setHorizontalSpacing(10)
            commercial_layout.setVerticalSpacing(8)
            commercial_hint = QLabel("A linha fica simples: descrição, fornecedor, quantidade e preço. A ponte com o stock e os custos mantém-se por baixo.")
            commercial_hint.setProperty("role", "muted")
            commercial_hint.setWordWrap(True)
            commercial_layout.addWidget(commercial_hint, 0, 0, 1, 4)
            commercial_layout.addWidget(QLabel("Descrição da linha"), 1, 0)
            commercial_layout.addWidget(desc_edit, 1, 1, 1, 3)
            commercial_layout.addWidget(QLabel("Fornecedor linha"), 2, 0)
            commercial_layout.addWidget(supplier_combo, 2, 1)
            commercial_layout.addWidget(QLabel("Quantidade"), 2, 2)
            commercial_layout.addWidget(qtd_spin, 2, 3)
            commercial_layout.addWidget(QLabel("Unid."), 3, 0)
            commercial_layout.addWidget(unid_edit, 3, 1)
            commercial_layout.addWidget(QLabel("Preço unid. (EUR)"), 3, 2)
            commercial_layout.addWidget(preco_spin, 3, 3)
            commercial_layout.addWidget(QLabel("Desc. %"), 4, 0)
            commercial_layout.addWidget(desconto_spin, 4, 1)
            commercial_layout.addWidget(QLabel("IVA %"), 4, 2)
            commercial_layout.addWidget(iva_spin, 4, 3)
            for col in range(4):
                commercial_layout.setColumnStretch(col, 1)
            layout.addWidget(commercial_card)

            sync = {"busy": False}
            desc_state = {"manual": False, "last_auto": desc_edit.text().strip()}
            current_preview = {"data": {}}

            def _combo_text(combo: QComboBox) -> str:
                return combo.currentText().strip()

            def _current_family_key() -> str:
                return str(familia_combo.currentData() or "").strip()

            def _set_family_value(value: Any, *, material_txt: Any = "") -> None:
                if str(value or "").strip():
                    target_key = str(_material_family_profile(material_txt, value).get("key", "") or "").strip()
                else:
                    target_key = ""
                familia_combo.blockSignals(True)
                target_index = 0
                for idx in range(familia_combo.count()):
                    if str(familia_combo.itemData(idx) or "").strip() == target_key:
                        target_index = idx
                        break
                familia_combo.setCurrentIndex(target_index)
                familia_combo.blockSignals(False)

            def _set_section_options(formato_txt: str, current_value: Any = "") -> None:
                secao_tipo_combo.blockSignals(True)
                secao_tipo_combo.clear()
                secao_tipo_combo.addItem("", "")
                target_text = str(current_value or "").strip()
                target_index = 0
                for option in list(self.backend.material_section_options(formato_txt) or []):
                    label_txt = str(option.get("label", "") or "").strip()
                    key_txt = str(option.get("key", "") or "").strip()
                    if not label_txt:
                        continue
                    secao_tipo_combo.addItem(label_txt, key_txt)
                    if target_text and target_text in {label_txt, key_txt}:
                        target_index = secao_tipo_combo.count() - 1
                secao_tipo_combo.setCurrentIndex(target_index)
                if target_text and target_index == 0:
                    secao_tipo_combo.setEditText(target_text)
                secao_tipo_combo.blockSignals(False)

            def _section_value() -> str:
                return str(secao_tipo_combo.currentData() or secao_tipo_combo.currentText() or "").strip()

            def _price_base_label(formato_txt: str) -> str:
                return "Preço metro (EUR/m)" if formato_txt == "Tubo" else "Preço kg (EUR/kg)"

            def _metric_label_text(formato_txt: str) -> str:
                return "Peso por unid. (kg)"

            def _set_technical_visibility(formato_txt: str, secao_tipo: str, *, selected_stock: bool, profile_catalog: bool) -> None:
                tube_round = formato_txt == "Tubo" and secao_tipo == "redondo"
                secao_visible = formato_txt in {"Tubo", "Perfil"}
                comp_visible = formato_txt == "Chapa" or (formato_txt == "Tubo" and not tube_round)
                larg_visible = formato_txt == "Chapa" or (formato_txt == "Tubo" and not tube_round)
                altura_visible = formato_txt == "Perfil"
                diametro_visible = formato_txt == "Tubo" and tube_round
                metros_visible = formato_txt in {"Tubo", "Perfil"}
                kgm_visible = formato_txt in {"Tubo", "Perfil"}
                secao_label.setText("Tipo tubo" if formato_txt == "Tubo" else ("Tipo perfil / série" if formato_txt == "Perfil" else "Tipo secção"))
                comp_label.setText("Comprimento (mm)" if formato_txt == "Chapa" else "Lado A (mm)")
                larg_label.setText("Largura (mm)" if formato_txt == "Chapa" else "Lado B (mm)")
                altura_label.setText("Altura / tamanho (mm)")
                diametro_label.setText("Diâmetro ext. (mm)")
                metros_label.setText("Comprimento barra (m)")
                kgm_label.setText("Peso por metro (kg/m)")
                for widget in (secao_label, secao_tipo_combo):
                    widget.setVisible(secao_visible)
                for widget in (comp_label, comprimento_edit):
                    widget.setVisible(comp_visible)
                for widget in (larg_label, largura_edit):
                    widget.setVisible(larg_visible)
                for widget in (altura_label, altura_edit):
                    widget.setVisible(altura_visible)
                for widget in (diametro_label, diametro_edit):
                    widget.setVisible(diametro_visible)
                for widget in (metros_label, metros_edit):
                    widget.setVisible(metros_visible)
                for widget in (kgm_label, kg_m_edit):
                    widget.setVisible(kgm_visible)
                metric_label.setVisible(True)
                metric_spin.setVisible(True)
                secao_tipo_combo.setEnabled(not selected_stock and secao_visible)
                comprimento_edit.setReadOnly(selected_stock or not comp_visible)
                largura_edit.setReadOnly(selected_stock or not larg_visible)
                altura_edit.setReadOnly(selected_stock or not altura_visible)
                diametro_edit.setReadOnly(selected_stock or not diametro_visible)
                metros_edit.setReadOnly(selected_stock or not metros_visible)
                kg_m_edit.setReadOnly(formato_txt == "Tubo" or profile_catalog or selected_stock or not kgm_visible)

            def _current_geometry_payload() -> dict[str, Any]:
                return {
                    "formato": _combo_text(formato_combo) or _clean((current_payload or {}).get("formato", ""), "Chapa"),
                    "material": _combo_text(qualidade_combo),
                    "material_familia": _current_family_key(),
                    "espessura": _combo_text(espessura_combo),
                    "secao_tipo": _section_value(),
                    "comprimento": comprimento_edit.text().strip(),
                    "largura": largura_edit.text().strip(),
                    "altura": altura_edit.text().strip(),
                    "diametro": diametro_edit.text().strip(),
                    "metros": metros_edit.text().strip(),
                    "kg_m": kg_m_edit.text().strip(),
                    "peso_unid": float(metric_spin.value() or 0),
                }

            def _apply_preview_controls(preview: dict[str, Any], *, selected_stock: bool, hydrate_dimensions: bool) -> None:
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                secao_tipo = str(preview.get("secao_tipo", "") or "").strip()
                _set_section_options(formato_txt, secao_tipo)
                if secao_tipo:
                    secao_tipo_combo.setCurrentText(secao_tipo)
                profile_catalog = formato_txt == "Perfil" and bool(preview.get("usa_catalogo"))
                _set_technical_visibility(formato_txt, secao_tipo, selected_stock=selected_stock, profile_catalog=profile_catalog)
                if hydrate_dimensions:
                    comprimento_edit.setText(_float_text(preview.get("comprimento", 0), 3))
                    largura_edit.setText(_float_text(preview.get("largura", 0), 3))
                    altura_edit.setText(_float_text(preview.get("altura", 0), 3))
                    diametro_edit.setText(_float_text(preview.get("diametro", 0), 3))
                    metros_edit.setText(_float_text(preview.get("metros", 0), 4))
                kg_m_edit.setText(_float_text(preview.get("kg_m", 0), 4))
                metric_label.setText(_metric_label_text(formato_txt))
                metric_spin.setValue(float(preview.get("peso_unid", 0) or 0))

            def _parse_dimension_values(formato_txt: str, dim_txt: str) -> tuple[float, float]:
                if formato_txt not in {"Chapa", "Perfil"}:
                    return 0.0, 0.0
                raw_values = re.findall(r"[-+]?\d+(?:[.,]\d+)?", str(dim_txt or ""))
                if len(raw_values) < 2:
                    return 0.0, 0.0
                try:
                    comp_val = float(raw_values[0].replace(",", "."))
                    larg_val = float(raw_values[1].replace(",", "."))
                except Exception:
                    return 0.0, 0.0
                return round(comp_val, 3), round(larg_val, 3)

            def _current_rows() -> list[dict[str, Any]]:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                dim_txt = _combo_text(dimensao_combo)
                if not formato_txt:
                    return []
                return [row for row in material_rows if _matches(row, formato_txt, quality_txt, esp_txt, dim_txt)]

            def _set_combo_items(combo: QComboBox, values: list[str], current: str = "", allow_blank: bool = False) -> None:
                current_text = current if current is not None else _combo_text(combo)
                combo.blockSignals(True)
                combo.clear()
                if allow_blank:
                    combo.addItem("")
                for value in values:
                    combo.addItem(value)
                available = [combo.itemText(i) for i in range(combo.count())]
                if current_text and current_text in available:
                    combo.setCurrentText(current_text)
                elif current_text and combo.isEditable():
                    combo.setEditText(current_text)
                elif combo.count() > 0 and not allow_blank:
                    combo.setCurrentIndex(0)
                elif combo.count() > 0 and allow_blank:
                    combo.setCurrentIndex(0)
                combo.blockSignals(False)

            def _stock_label(row: dict[str, Any]) -> str:
                parts = [str(row.get("id", "") or "").strip()]
                lote_txt = str(row.get("lote", "") or "").strip()
                local_txt = str(row.get("localizacao", "") or "").strip()
                if lote_txt:
                    parts.append(f"Lote {lote_txt}")
                if local_txt:
                    parts.append(local_txt)
                return " | ".join(part for part in parts if part) or str(row.get("descricao", "") or "").strip()

            def _compute_unit_price(formato_txt: str, metric_value: float, base_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                return round(base_value * metric_value, 4) if metric_value > 0 else 0.0

            def _compute_base_price(formato_txt: str, metric_value: float, unit_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                return round(unit_value / metric_value, 6) if metric_value > 0 else 0.0

            def _auto_description() -> str:
                preview = dict(current_preview.get("data") or {})
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                dim_txt = str(preview.get("dimension_label", "") or "").strip() or _combo_text(dimensao_combo)
                parts = [part for part in (quality_txt,) if part]
                if esp_txt:
                    parts.append(f"{esp_txt} mm")
                if dim_txt and dim_txt != "-":
                    parts.append(dim_txt)
                if str(preview.get("formato", "") or "").strip() in {"Tubo", "Perfil"} and float(preview.get("metros", 0) or 0) > 0:
                    parts.append(f"{_float_text(preview.get('metros', 0), 3)} m")
                return " | ".join(parts).strip()

            def _set_auto_description(text: str, force: bool = False) -> None:
                current_text = desc_edit.text().strip()
                if not force and desc_state["manual"] and current_text and current_text != desc_state["last_auto"]:
                    return
                desc_edit.blockSignals(True)
                desc_edit.setText(text)
                desc_edit.blockSignals(False)
                desc_state["last_auto"] = text
                if force:
                    desc_state["manual"] = False

            def _selected_metric_value(payload: dict[str, Any] | None = None) -> float:
                if isinstance(payload, dict):
                    if str(payload.get("formato", "") or "").strip() == "Tubo":
                        return float(payload.get("metros", 0) or 0)
                    return float(payload.get("peso_unid", 0) or 0)
                return float(metric_spin.value() or 0)

            def _refresh_manual_state() -> None:
                if sync["busy"]:
                    return
                formato_txt = _combo_text(formato_combo) or _clean((current_payload or {}).get("formato", ""), "Chapa")
                sync["busy"] = True
                try:
                    price_base_label.setText(_price_base_label(formato_txt))
                    selected_stock = stock_combo.currentData()
                    has_selected_stock = isinstance(selected_stock, dict)
                    familia_combo.setEnabled(not has_selected_stock)
                    code_display.setText(
                        str((selected_stock or {}).get("id", "") or "").strip() if has_selected_stock else (current_ref if current_ref else "")
                    )
                    lote_display.setReadOnly(has_selected_stock)
                    local_display.setReadOnly(has_selected_stock)
                    quality_txt = _combo_text(qualidade_combo)
                    selected_family = str((selected_stock or {}).get("material_familia", (selected_stock or {}).get("material_familia_resolved", "")) or "").strip()
                    if has_selected_stock:
                        _set_family_value(selected_family, material_txt=(selected_stock or {}).get("material", ""))
                    family_profile = _material_family_profile(quality_txt or (selected_stock or {}).get("material", ""), _current_family_key() or selected_family)
                    densidade_display.setText(_float_text(family_profile.get("density", 0), 3))
                    manual_payload = dict(selected_stock or {})
                    if not has_selected_stock:
                        manual_payload = _current_geometry_payload()
                    preview = dict(self.backend.material_price_preview(manual_payload) or {})
                    current_preview["data"] = preview
                    _apply_preview_controls(preview, selected_stock=has_selected_stock, hydrate_dimensions=has_selected_stock)
                    if not has_selected_stock:
                        preco_spin.setValue(_compute_unit_price(formato_txt, _selected_metric_value(preview), float(price_base_spin.value() or 0)))
                    _set_auto_description(_auto_description())
                finally:
                    sync["busy"] = False

            def _apply_payload(row: dict[str, Any], *, prefer_existing_prices: bool = False) -> None:
                if not isinstance(row, dict):
                    return
                sync["busy"] = True
                try:
                    ref_edit.setText(str(row.get("id", "") or "").strip())
                    code_display.setText(str(row.get("id", "") or "").strip())
                    lote_display.setText(str(row.get("lote", "") or "").strip())
                    local_display.setText(str(row.get("localizacao", "") or "").strip())
                    unid_edit.setText(str(row.get("unid", "UN") or "UN").strip() or "UN")
                    formato_txt = str(row.get("formato", "") or "").strip() or "Chapa"
                    price_base_label.setText(_price_base_label(formato_txt))
                    _set_family_value(row.get("material_familia", row.get("material_familia_resolved", "")), material_txt=row.get("material", ""))
                    family_profile = _material_family_profile(row.get("material", ""), str(row.get("material_familia", row.get("material_familia_resolved", "")) or "").strip())
                    densidade_display.setText(_float_text(family_profile.get("density", 0), 3))
                    preview = dict(self.backend.material_price_preview(row) or {})
                    current_preview["data"] = preview
                    _apply_preview_controls(preview, selected_stock=True, hydrate_dimensions=True)
                    price_base_spin.setValue(
                        float(initial.get("p_compra", row.get("preco_base", 0)) or 0)
                        if prefer_existing_prices
                        else float(row.get("preco_base", 0) or 0)
                    )
                    preco_spin.setValue(
                        float(initial.get("preco", row.get("preco", 0)) or 0)
                        if prefer_existing_prices
                        else float(row.get("preco", 0) or 0)
                    )
                    lote_display.setReadOnly(True)
                    local_display.setReadOnly(True)
                    _set_auto_description(_auto_description(), force=True)
                finally:
                    sync["busy"] = False

            def _refresh_stock_combo(preserve_stock_id: str = "") -> None:
                rows = _current_rows()
                stock_combo.blockSignals(True)
                stock_combo.clear()
                stock_combo.addItem("", None)
                for row in rows:
                    stock_combo.addItem(_stock_label(row), row)
                stock_combo.blockSignals(False)
                target_id = preserve_stock_id or current_ref
                if target_id:
                    for idx in range(stock_combo.count()):
                        payload = stock_combo.itemData(idx)
                        if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == target_id:
                            stock_combo.setCurrentIndex(idx)
                            break
                elif rows:
                    stock_combo.setCurrentIndex(1)
                payload = stock_combo.currentData()
                if isinstance(payload, dict):
                    _apply_payload(payload, prefer_existing_prices=bool(current_ref and str(payload.get("id", "") or "").strip() == current_ref))
                else:
                    _refresh_manual_state()

            def _refresh_dimensions() -> None:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                if not formato_txt:
                    _set_combo_items(espessura_combo, [], _clean((current_payload or {}).get("espessura", "")), allow_blank=True)
                    _set_combo_items(dimensao_combo, [], _clean((current_payload or {}).get("dimensao", "")), allow_blank=True)
                    _refresh_stock_combo()
                    return
                candidate_rows = [row for row in material_rows if _matches(row, formato_txt, quality_txt, "", "")]
                esp_values = _unique(candidate_rows, lambda row: _clean(row.get("espessura", ""), "-"))
                current_esp = _combo_text(espessura_combo) or _clean((current_payload or {}).get("espessura", ""))
                _set_combo_items(espessura_combo, esp_values, current_esp, allow_blank=True)
                chosen_esp = _combo_text(espessura_combo)
                dim_rows = [row for row in candidate_rows if _matches(row, formato_txt, quality_txt, chosen_esp, "")]
                dim_values = _unique(dim_rows, _dimension_text)
                if formato_txt == "Chapa":
                    merged_dims: list[str] = []
                    for value in [*common_sheet_dimensions, *dim_values]:
                        if value and value not in merged_dims:
                            merged_dims.append(value)
                    dim_values = merged_dims
                current_dim = _combo_text(dimensao_combo) or (_dimension_text(current_payload or {}) if current_payload else "")
                _set_combo_items(dimensao_combo, dim_values, current_dim, allow_blank=True)
                _refresh_stock_combo()

            def _refresh_quality() -> None:
                formato_txt = _combo_text(formato_combo)
                if not formato_txt:
                    _set_combo_items(qualidade_combo, [], _clean((current_payload or {}).get("material", "")), allow_blank=True)
                    _refresh_dimensions()
                    return
                candidate_rows = [row for row in material_rows if _matches(row, formato_txt, "", "", "")]
                qualities = _unique(candidate_rows, lambda row: row.get("material", ""))
                current_quality = _combo_text(qualidade_combo) or _clean((current_payload or {}).get("material", ""))
                _set_combo_items(qualidade_combo, qualities, current_quality, allow_blank=True)
                _refresh_dimensions()

            def _refresh_formats() -> None:
                formats: list[str] = []
                for value in ["Chapa", "Perfil", "Tubo", *_unique(material_rows, lambda row: row.get("formato", ""))]:
                    if value and value not in formats:
                        formats.append(value)
                current_format = _clean((current_payload or {}).get("formato", ""))
                _set_combo_items(formato_combo, formats, current_format, allow_blank=True)
                _refresh_quality()

            def _on_stock_change() -> None:
                payload = stock_combo.currentData()
                if isinstance(payload, dict):
                    _apply_payload(payload)
                else:
                    _refresh_manual_state()

            def _on_base_price_changed(value: float) -> None:
                if sync["busy"]:
                    return
                preview = dict(current_preview.get("data") or {})
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                metric_value = _selected_metric_value(preview)
                sync["busy"] = True
                try:
                    preco_spin.setValue(_compute_unit_price(formato_txt, metric_value, value))
                finally:
                    sync["busy"] = False

            def _on_unit_price_changed(value: float) -> None:
                if sync["busy"]:
                    return
                preview = dict(current_preview.get("data") or {})
                formato_txt = str(preview.get("formato", "") or _combo_text(formato_combo) or "Chapa").strip() or "Chapa"
                metric_value = _selected_metric_value(preview)
                sync["busy"] = True
                try:
                    price_base_spin.setValue(_compute_base_price(formato_txt, metric_value, value))
                finally:
                    sync["busy"] = False

            def _on_description_edited(_text: str) -> None:
                if sync["busy"]:
                    return
                desc_state["manual"] = True

            formato_combo.currentTextChanged.connect(lambda _text: _refresh_quality())
            qualidade_combo.currentTextChanged.connect(lambda _text: _refresh_dimensions())
            familia_combo.currentIndexChanged.connect(lambda _i: _refresh_manual_state())
            espessura_combo.currentTextChanged.connect(lambda _text: _refresh_dimensions())
            dimensao_combo.currentTextChanged.connect(lambda _text: _refresh_stock_combo())
            secao_tipo_combo.currentTextChanged.connect(lambda _text: _refresh_manual_state())
            stock_combo.currentIndexChanged.connect(lambda _i: _on_stock_change())
            price_base_spin.valueChanged.connect(_on_base_price_changed)
            preco_spin.valueChanged.connect(_on_unit_price_changed)
            for edit in (comprimento_edit, largura_edit, altura_edit, diametro_edit, metros_edit, kg_m_edit):
                edit.textChanged.connect(lambda _text: _refresh_manual_state())
            desc_edit.textEdited.connect(_on_description_edited)
            supplier_combo.setCurrentText(str(initial.get("fornecedor_linha", "") or self.supplier_combo.currentText() or "").strip())
            local_display.setText(str(initial.get("localizacao", "") or "").strip())
            lote_display.setText(str(initial.get("lote_fornecedor", "") or "").strip())
            _set_family_value(str(initial.get("material_familia", (current_payload or {}).get("material_familia", "")) or "").strip(), material_txt=(current_payload or {}).get("material", ""))
            if current_payload is not None and not current_ref:
                preview = dict(self.backend.material_price_preview(current_payload) or {})
                current_preview["data"] = preview
                _apply_preview_controls(preview, selected_stock=False, hydrate_dimensions=True)
                price_base_spin.setValue(float(initial.get("p_compra", current_payload.get("preco_base", 0)) or 0))
                preco_spin.setValue(float(initial.get("preco", current_payload.get("preco", 0)) or 0))
            _refresh_formats()
        else:
            intro = QLabel(
                "Linha manual sem ligação direta ao stock. Define os dados comerciais e escolhe se o preço base é por unidade, kg ou metro."
            )
            intro.setWordWrap(True)
            intro.setProperty("role", "muted")
            layout.addWidget(intro)

            data_card = CardFrame()
            data_layout = QFormLayout(data_card)
            data_layout.setContentsMargins(12, 10, 12, 10)
            data_layout.setHorizontalSpacing(10)
            data_layout.setVerticalSpacing(8)
            data_hint = QLabel("A linha manual fica autónoma, mas com a mesma lógica de preço base da matéria-prima e dos produtos.")
            data_hint.setWordWrap(True)
            data_hint.setProperty("role", "muted")
            data_layout.addRow(data_hint)
            data_layout.addRow("Ref.", ref_edit)
            data_layout.addRow("Descrição", desc_edit)
            data_layout.addRow("Fornecedor linha", supplier_combo)
            data_layout.addRow("Quantidade", qtd_spin)
            data_layout.addRow("Unid.", unid_edit)
            layout.addWidget(data_card)

            pricing_card = CardFrame()
            pricing_card.set_tone("info")
            pricing_layout = QFormLayout(pricing_card)
            pricing_layout.setContentsMargins(12, 10, 12, 10)
            pricing_layout.setHorizontalSpacing(10)
            pricing_layout.setVerticalSpacing(8)
            pricing_hint = QLabel("Se a compra vier por kg ou metro, indica a métrica por unidade para a ponte automática com o preço unitário.")
            pricing_hint.setWordWrap(True)
            pricing_hint.setProperty("role", "muted")
            pricing_layout.addRow(pricing_hint)
            manual_basis_combo = QComboBox()
            manual_basis_combo.addItem("Unidade", "unit")
            manual_basis_combo.addItem("Kg", "kg")
            manual_basis_combo.addItem("Metro", "meter")
            manual_price_base_label = QLabel("Preço base (EUR/un)")
            manual_price_base_spin = QDoubleSpinBox()
            manual_price_base_spin.setRange(0.0, 1000000.0)
            manual_price_base_spin.setDecimals(4)
            manual_metric_label = QLabel("Métrica por unid.")
            manual_metric_spin = QDoubleSpinBox()
            manual_metric_spin.setRange(0.0, 1000000.0)
            manual_metric_spin.setDecimals(4)
            pricing_layout.addRow("Base de preço", manual_basis_combo)
            pricing_layout.addRow(manual_price_base_label, manual_price_base_spin)
            pricing_layout.addRow(manual_metric_label, manual_metric_spin)
            pricing_layout.addRow("Preço unid. (EUR)", preco_spin)
            pricing_layout.addRow("Desc. %", desconto_spin)
            pricing_layout.addRow("IVA %", iva_spin)
            layout.addWidget(pricing_card)

            manual_sync = {"busy": False}

            def _manual_basis() -> str:
                return str(manual_basis_combo.currentData() or "unit").strip() or "unit"

            def _manual_metric_value() -> float:
                return float(manual_metric_spin.value() or 0)

            def _manual_base_label_text(basis: str) -> str:
                if basis == "kg":
                    return "Preço kg (EUR/kg)"
                if basis == "meter":
                    return "Preço metro (EUR/m)"
                return "Preço base (EUR/un)"

            def _manual_metric_label_text(basis: str) -> str:
                if basis == "kg":
                    return "Kg por unid."
                if basis == "meter":
                    return "Metros por unid. (m)"
                return "Métrica por unid."

            def _manual_compute_unit_price(basis: str, metric_value: float, base_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                if basis in {"kg", "meter"}:
                    return round(base_value * metric_value, 4) if metric_value > 0 else 0.0
                return round(base_value, 4)

            def _manual_compute_base_price(basis: str, metric_value: float, unit_value: float) -> float:
                metric_value = max(0.0, float(metric_value or 0))
                if basis in {"kg", "meter"}:
                    return round(unit_value / metric_value, 6) if metric_value > 0 else round(unit_value, 6)
                return round(unit_value, 6)

            def _refresh_manual_price_state(*, recalc_from_unit: bool = False) -> None:
                basis = _manual_basis()
                show_metric = basis in {"kg", "meter"}
                manual_price_base_label.setText(_manual_base_label_text(basis))
                manual_metric_label.setText(_manual_metric_label_text(basis))
                manual_metric_label.setVisible(show_metric)
                manual_metric_spin.setVisible(show_metric)
                if manual_sync["busy"]:
                    return
                if basis == "unit":
                    manual_sync["busy"] = True
                    try:
                        manual_price_base_spin.setValue(float(preco_spin.value() or 0))
                    finally:
                        manual_sync["busy"] = False
                elif recalc_from_unit:
                    _on_manual_unit_price_changed(preco_spin.value())

            def _on_manual_base_price_changed(value: float) -> None:
                if manual_sync["busy"]:
                    return
                manual_sync["busy"] = True
                try:
                    preco_spin.setValue(_manual_compute_unit_price(_manual_basis(), _manual_metric_value(), value))
                finally:
                    manual_sync["busy"] = False

            def _on_manual_unit_price_changed(value: float) -> None:
                if manual_sync["busy"]:
                    return
                manual_sync["busy"] = True
                try:
                    manual_price_base_spin.setValue(_manual_compute_base_price(_manual_basis(), _manual_metric_value(), value))
                finally:
                    manual_sync["busy"] = False

            def _on_manual_metric_changed(_value: float) -> None:
                if manual_sync["busy"]:
                    return
                _refresh_manual_price_state(recalc_from_unit=True)

            def _on_manual_basis_changed(_index: int) -> None:
                if manual_sync["busy"]:
                    return
                _refresh_manual_price_state(recalc_from_unit=True)

            def _accept_manual() -> None:
                if not desc_edit.text().strip():
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a descrição da linha.")
                    return
                if _manual_basis() in {"kg", "meter"} and _manual_metric_value() <= 0:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a métrica por unidade para preço por kg ou metro.")
                    return
                dialog.accept()

            initial_basis = str(initial.get("price_basis", "") or "").strip().lower()
            if initial_basis in {"kg", "peso"}:
                initial_basis = "kg"
            elif initial_basis in {"meter", "metro", "metros"}:
                initial_basis = "meter"
            elif float(initial.get("peso_unid", 0) or 0) > 0:
                initial_basis = "kg"
            elif float(initial.get("metros_unidade", initial.get("metros", 0)) or 0) > 0:
                initial_basis = "meter"
            else:
                initial_basis = "unit"

            initial_metric = 0.0
            if initial_basis == "kg":
                initial_metric = float(initial.get("peso_unid", 0) or 0)
            elif initial_basis == "meter":
                initial_metric = float(initial.get("metros_unidade", initial.get("metros", 0)) or 0)

            initial_base_price = (
                float(initial.get("p_compra", 0) or 0)
                if "p_compra" in initial
                else _manual_compute_base_price(initial_basis, initial_metric, float(initial.get("preco", 0) or 0))
            )
            basis_index = {"unit": 0, "kg": 1, "meter": 2}.get(initial_basis, 0)
            manual_basis_combo.setCurrentIndex(basis_index)
            manual_metric_spin.setValue(initial_metric)
            manual_price_base_spin.setValue(initial_base_price)
            _refresh_manual_price_state(recalc_from_unit=False)

            manual_basis_combo.currentIndexChanged.connect(_on_manual_basis_changed)
            manual_price_base_spin.valueChanged.connect(_on_manual_base_price_changed)
            preco_spin.valueChanged.connect(_on_manual_unit_price_changed)
            manual_metric_spin.valueChanged.connect(_on_manual_metric_changed)
            accept_handler = _accept_manual
            manual_price_controls = {
                "basis": manual_basis_combo,
                "base_price": manual_price_base_spin,
                "metric": manual_metric_spin,
            }
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_button = buttons.button(QDialogButtonBox.Ok)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if ok_button is not None:
            ok_button.setText("Guardar")
        if cancel_button is not None:
            cancel_button.setText("Fechar")
        buttons.accepted.connect(accept_handler)
        buttons.rejected.connect(dialog.reject)
        dialog_layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        payload = {
            "ref": ref_edit.text().strip(),
            "descricao": desc_edit.text().strip(),
            "fornecedor_linha": supplier_combo.currentText().strip(),
            "origem": "Matéria-Prima" if material_mode else str(initial.get("origem", "Produto") or "Produto"),
            "qtd": qtd_spin.value(),
            "unid": unid_edit.text().strip() or "UN",
            "preco": preco_spin.value(),
            "desconto": desconto_spin.value(),
            "iva": iva_spin.value(),
            "entrega": str(initial.get("entrega", "PENDENTE") or "PENDENTE"),
        }
        if material_mode:
            selected_material = stock_combo.currentData()
            if isinstance(selected_material, dict):
                preview = dict(self.backend.material_price_preview(selected_material) or {})
                payload.update(
                    {
                        "formato": str(selected_material.get("formato", "") or "").strip(),
                        "material": str(selected_material.get("material", "") or "").strip(),
                        "material_familia": str(selected_material.get("material_familia", selected_material.get("material_familia_resolved", "")) or "").strip(),
                        "espessura": str(selected_material.get("espessura", "") or "").strip(),
                        "comprimento": float(preview.get("comprimento", selected_material.get("comprimento", 0)) or 0),
                        "largura": float(preview.get("largura", selected_material.get("largura", 0)) or 0),
                        "altura": float(preview.get("altura", selected_material.get("altura", 0)) or 0),
                        "diametro": float(preview.get("diametro", selected_material.get("diametro", 0)) or 0),
                        "metros": float(preview.get("metros", selected_material.get("metros", 0)) or 0),
                        "kg_m": float(preview.get("kg_m", selected_material.get("kg_m", 0)) or 0),
                        "peso_unid": float(preview.get("peso_unid", selected_material.get("peso_unid", 0)) or 0),
                        "secao_tipo": str(preview.get("secao_tipo", selected_material.get("secao_tipo", "")) or "").strip(),
                        "p_compra": float(price_base_spin.value() or 0),
                        "localizacao": str(selected_material.get("localizacao", "") or "").strip(),
                        "lote_fornecedor": str(selected_material.get("lote", "") or "").strip(),
                    }
                )
            else:
                formato_txt = _combo_text(formato_combo)
                quality_txt = _combo_text(qualidade_combo)
                esp_txt = _combo_text(espessura_combo)
                manual_raw = {
                    **_current_geometry_payload(),
                    "quantidade": 1,
                    "reservado": 0,
                    "p_compra": float(price_base_spin.value() or 0),
                    "local": local_display.text().strip(),
                    "lote_fornecedor": lote_display.text().strip(),
                }
                manual_desc = desc_edit.text().strip() or _auto_description()
                if not formato_txt:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica o tipo de material.")
                    return None
                if not quality_txt:
                    QMessageBox.warning(dialog, "Notas Encomenda", "Indica a qualidade da matéria-prima.")
                    return None
                try:
                    normalized_material = dict(self.backend._normalise_material_payload(manual_raw) or {})
                except Exception as exc:
                    QMessageBox.warning(dialog, "Notas Encomenda", str(exc))
                    return None
                payload.update(
                    {
                        "ref": "",
                        "descricao": manual_desc,
                        "formato": formato_txt,
                        "material": quality_txt,
                        "material_familia": str(normalized_material.get("material_familia", _current_family_key()) or "").strip(),
                        "espessura": str(normalized_material.get("espessura", esp_txt) or "").strip(),
                        "comprimento": float(normalized_material.get("comprimento", 0) or 0),
                        "largura": float(normalized_material.get("largura", 0) or 0),
                        "altura": float(normalized_material.get("altura", 0) or 0),
                        "diametro": float(normalized_material.get("diametro", 0) or 0),
                        "metros": float(normalized_material.get("metros", 0) or 0),
                        "kg_m": float(normalized_material.get("kg_m", 0) or 0),
                        "peso_unid": float(normalized_material.get("peso_unid", 0) or 0),
                        "secao_tipo": str(normalized_material.get("secao_tipo", "") or "").strip(),
                        "p_compra": float(price_base_spin.value() or 0),
                        "localizacao": local_display.text().strip(),
                        "lote_fornecedor": lote_display.text().strip(),
                        "_material_manual": True,
                        "_material_pending_create": True,
                    }
                )
        elif manual_price_controls is not None:
            basis = str(manual_price_controls["basis"].currentData() or "unit").strip() or "unit"
            metric_value = float(manual_price_controls["metric"].value() or 0)
            payload.update(
                {
                    "p_compra": float(manual_price_controls["base_price"].value() or 0),
                    "price_basis": basis,
                    "peso_unid": metric_value if basis == "kg" else 0.0,
                    "metros": metric_value if basis == "meter" else 0.0,
                    "metros_unidade": metric_value if basis == "meter" else 0.0,
                }
            )
        return payload

    def _quick_assign_supplier(self) -> None:
        if not self.line_rows:
            QMessageBox.warning(self, "Notas Encomenda", "Não existem linhas para adjudicar.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Fornecedor rápido")
        dialog.setMinimumWidth(460)
        layout = QVBoxLayout(dialog)
        info = QLabel("Aplica o fornecedor às linhas selecionadas. Se não houver seleção, aplica a todas as linhas.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        supplier_combo = QComboBox()
        self._configure_supplier_selector(supplier_combo)
        supplier_combo.addItem("")
        for supplier in self.supplier_rows:
            supplier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        supplier_combo.setCurrentText(self.supplier_combo.currentText().strip())
        layout.addWidget(supplier_combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        supplier_text = supplier_combo.currentText().strip()
        if not supplier_text:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona um fornecedor.")
            return
        target_indexes = self._selected_line_indexes() or list(range(len(self.line_rows)))
        for idx in target_indexes:
            self.line_rows[idx]["fornecedor_linha"] = supplier_text
        self._render_lines()

    def _add_material_line(self) -> None:
        payload = self._line_dialog("Adicionar linha de Matéria-Prima", {"origem": "Matéria-Prima", "iva": 23, "qtd": 1}, material_mode=True)
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _add_product_line(self) -> None:
        payload = self._product_line_dialog("Adicionar produto em stock", {"origem": "Produto", "iva": 23, "qtd": 1})
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _add_manual_line(self) -> None:
        payload = self._line_dialog("Adicionar linha manual", {"origem": "Produto", "iva": 23, "qtd": 1}, material_mode=False)
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_lines()

    def _edit_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma linha.")
            return
        current = dict(self.line_rows[index])
        is_material = self.backend.desktop_main.origem_is_materia(current.get("origem", ""))
        if str(current.get("origem", "") or "").strip().lower() == "produto":
            payload = self._product_line_dialog("Editar linha de produto", current)
        else:
            payload = self._line_dialog("Editar linha", current, material_mode=is_material)
        if payload is None:
            return
        payload["origem"] = current.get("origem", payload.get("origem", "Produto"))
        payload["entrega"] = current.get("entrega", "PENDENTE")
        self.line_rows[index] = payload
        self._render_lines()
        self.lines_table.selectRow(index)

    def _remove_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma linha.")
            return
        del self.line_rows[index]
        self._render_lines()

    def _note_payload(self) -> dict:
        supplier_text = self.supplier_combo.currentText().strip()
        supplier = self._supplier_lookup(supplier_text)
        payload = {
            "numero": self.current_number,
            "fornecedor": supplier_text or str(supplier.get("nome", "") or ""),
            "fornecedor_id": str(supplier.get("id", "") or ""),
            "contacto": self.contact_edit.text().strip(),
            "data_entrega": self._delivery_text(),
            "obs": self.obs_edit.toPlainText().strip(),
            "local_descarga": self.location_edit.currentText().strip(),
            "meio_transporte": self.transport_edit.currentText().strip(),
            "lines": self.line_rows,
        }
        return payload

    def _save_note(self) -> None:
        try:
            note = self.backend.ne_save(self._note_payload())
        except Exception as exc:
            QMessageBox.critical(self, "Guardar NE", str(exc))
            return
        self.current_number = str(note.get("numero", "")).strip()
        self.refresh()
        QMessageBox.information(self, "Guardar NE", f"Nota {self.current_number} guardada.")

    def _approve_note(self) -> None:
        try:
            self.backend.ne_save(self._note_payload())
            note = self.backend.ne_approve(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Aprovar NE", str(exc))
            return
        self.current_number = str(note.get("numero", "")).strip()
        self.refresh()
        QMessageBox.information(self, "Aprovar NE", f"Nota {self.current_number} aprovada.")

    def _remove_note(self) -> None:
        if not self.current_number:
            return
        if QMessageBox.question(self, "Apagar NE", f"Remover nota {self.current_number}?") != QMessageBox.Yes:
            return
        try:
            self.backend.ne_remove(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Apagar NE", str(exc))
            return
        self.current_number = ""
        self.refresh()

    def _open_pdf(self, quote: bool) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona ou guarda primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            path = self.backend.ne_open_pdf(self.current_number, quote=quote)
        except Exception as exc:
            QMessageBox.critical(self, "Notas Encomenda", str(exc))
            return
        QMessageBox.information(self, "Notas Encomenda", f"PDF aberto:\n{path}")

    def _register_delivery(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Guarda ou seleciona primeiro a nota.")
            return
        payload = self._delivery_dialog()
        if payload is None:
            return
        try:
            self.backend.ne_save(self._note_payload())
            self.backend.ne_register_delivery(self.current_number, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Entrega NE", str(exc))
            return
        self.refresh()
        QMessageBox.information(self, "Entrega NE", f"Entrega registada para {self.current_number}.")

    def _generate_supplier_orders(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Notas Encomenda", "Guarda ou seleciona primeiro a nota.")
            return
        try:
            self.backend.ne_save(self._note_payload())
            created = self.backend.ne_generate_supplier_orders(self.current_number)
        except Exception as exc:
            QMessageBox.critical(self, "Gerar NEs", str(exc))
            return
        self.refresh()
        summary = ", ".join(f"{row.get('numero')} ({row.get('fornecedor')})" for row in created)
        QMessageBox.information(self, "Gerar NEs", f"Notas geradas: {summary}")


class ClientsPage(QWidget):
    page_title = "Clientes"
    page_subtitle = "Cadastro comercial de clientes ligado ao backend atual."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_code = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        top = CardFrame()
        top.set_tone("info")
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(14, 10, 14, 10)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar codigo, nome, nif, contacto...")
        self.filter_edit.textChanged.connect(self.refresh)
        self.new_btn = QPushButton("Novo cliente")
        self.new_btn.clicked.connect(self._new_client)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_client)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_client)
        top_layout.addWidget(self.filter_edit, 1)
        top_layout.addWidget(self.new_btn)
        top_layout.addWidget(self.save_btn)
        top_layout.addWidget(self.remove_btn)
        root.addWidget(top)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_title = QLabel("Base de clientes")
        table_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Codigo", "Nome", "NIF", "Contacto", "Email"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.table, stretch=(1, 4), contents=(0, 2, 3))
        self.table.itemSelectionChanged.connect(self._load_selected_client)
        table_layout.addWidget(table_title)
        table_layout.addWidget(self.table)
        split.addWidget(table_card)

        form_card = CardFrame()
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)
        form_title = QLabel("Ficha do cliente")
        form_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        form = QFormLayout()
        self.client_code_edit = QLineEdit()
        self.client_name_edit = QLineEdit()
        self.client_nif_edit = QLineEdit()
        self.client_contact_edit = QLineEdit()
        self.client_email_edit = QLineEdit()
        self.client_address_edit = QTextEdit()
        self.client_address_edit.setMaximumHeight(88)
        self.client_terms_edit = QLineEdit()
        self.client_lead_edit = QLineEdit()
        self.client_notes_edit = QTextEdit()
        self.client_notes_edit.setMaximumHeight(110)
        for label, widget in (
            ("Codigo", self.client_code_edit),
            ("Nome", self.client_name_edit),
            ("NIF", self.client_nif_edit),
            ("Contacto", self.client_contact_edit),
            ("Email", self.client_email_edit),
            ("Morada", self.client_address_edit),
            ("Prazo entrega", self.client_lead_edit),
            ("Cond. pagamento", self.client_terms_edit),
            ("Observacoes", self.client_notes_edit),
        ):
            form.addRow(label, widget)
        form_layout.addWidget(form_title)
        form_layout.addLayout(form)
        form_layout.addStretch(1)
        split.addWidget(form_card)
        split.setSizes([760, 760])
        root.addWidget(split, 1)
        self._new_client()

    def refresh(self) -> None:
        previous = self.current_code
        self.rows = self.backend.client_rows(self.filter_edit.text().strip())
        _fill_table(
            self.table,
            [[r.get("codigo", "-"), r.get("nome", "-"), r.get("nif", "-"), r.get("contacto", "-"), r.get("email", "-")] for r in self.rows],
            align_center_from=2,
        )
        for idx, row in enumerate(self.rows):
            item = self.table.item(idx, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("codigo", "") or "").strip())
        if not self.rows:
            self._new_client()
            return
        target = 0
        if previous:
            for idx, row in enumerate(self.rows):
                if str(row.get("codigo", "") or "").strip() == previous:
                    target = idx
                    break
        self.table.selectRow(target)
        self._load_selected_client()

    def _selected_client(self) -> dict:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        item = self.table.item(row_index, 0)
        code = str(item.data(Qt.UserRole) or item.text() or "").strip()
        return next((row for row in self.rows if str(row.get("codigo", "") or "").strip() == code), {})

    def _new_client(self) -> None:
        self.current_code = ""
        self.client_code_edit.setText(self.backend.client_next_code())
        self.client_name_edit.clear()
        self.client_nif_edit.clear()
        self.client_contact_edit.clear()
        self.client_email_edit.clear()
        self.client_address_edit.clear()
        self.client_terms_edit.clear()
        self.client_lead_edit.clear()
        self.client_notes_edit.clear()

    def _load_selected_client(self) -> None:
        row = self._selected_client()
        if not row:
            return
        self.current_code = str(row.get("codigo", "") or "").strip()
        self.client_code_edit.setText(self.current_code)
        self.client_name_edit.setText(str(row.get("nome", "") or "").strip())
        self.client_nif_edit.setText(str(row.get("nif", "") or "").strip())
        self.client_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        self.client_email_edit.setText(str(row.get("email", "") or "").strip())
        self.client_address_edit.setPlainText(str(row.get("morada", "") or "").strip())
        self.client_lead_edit.setText(str(row.get("prazo_entrega", "") or "").strip())
        self.client_terms_edit.setText(str(row.get("cond_pagamento", "") or "").strip())
        self.client_notes_edit.setPlainText(str(row.get("observacoes", "") or "").strip())

    def _save_client(self) -> None:
        try:
            saved = self.backend.client_save(
                {
                    "codigo": self.client_code_edit.text().strip(),
                    "nome": self.client_name_edit.text().strip(),
                    "nif": self.client_nif_edit.text().strip(),
                    "contacto": self.client_contact_edit.text().strip(),
                    "email": self.client_email_edit.text().strip(),
                    "morada": self.client_address_edit.toPlainText().strip(),
                    "prazo_entrega": self.client_lead_edit.text().strip(),
                    "cond_pagamento": self.client_terms_edit.text().strip(),
                    "observacoes": self.client_notes_edit.toPlainText().strip(),
                }
            )
        except Exception as exc:
            QMessageBox.critical(self, "Clientes", str(exc))
            return
        self.current_code = str(saved.get("codigo", "") or "").strip()
        self.refresh()

    def _remove_client(self) -> None:
        code = self.client_code_edit.text().strip()
        if not code:
            return
        if QMessageBox.question(self, "Clientes", f"Remover cliente {code}?") != QMessageBox.Yes:
            return
        try:
            self.backend.client_remove(code)
        except Exception as exc:
            QMessageBox.critical(self, "Clientes", str(exc))
            return
        self._new_client()
        self.refresh()


class SuppliersPage(QWidget):
    page_title = "Fornecedores"
    page_subtitle = "Cadastro de fornecedores para notas de encomenda e compras."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_id = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        top = CardFrame()
        top.set_tone("info")
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(14, 10, 14, 10)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar fornecedor, nif, contacto...")
        self.filter_edit.textChanged.connect(self.refresh)
        self.new_btn = QPushButton("Novo fornecedor")
        self.new_btn.clicked.connect(self._new_supplier)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_supplier)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_supplier)
        top_layout.addWidget(self.filter_edit, 1)
        top_layout.addWidget(self.new_btn)
        top_layout.addWidget(self.save_btn)
        top_layout.addWidget(self.remove_btn)
        root.addWidget(top)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_title = QLabel("Fornecedores")
        table_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "Nome", "NIF", "Contacto", "Email"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.table, stretch=(1, 3, 4), contents=(2,))
        _set_table_columns(
            self.table,
            [
                (0, "interactive", 96),
                (1, "stretch", 220),
                (2, "interactive", 132),
                (3, "stretch", 156),
                (4, "stretch", 192),
            ],
        )
        self.table.itemSelectionChanged.connect(self._load_selected_supplier)
        table_layout.addWidget(table_title)
        table_layout.addWidget(self.table)
        split.addWidget(table_card)

        form_card = CardFrame()
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)
        form_title = QLabel("Ficha do fornecedor")
        form_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        form = QFormLayout()
        self.supplier_id_edit = QLineEdit()
        self.supplier_name_edit = QLineEdit()
        self.supplier_nif_edit = QLineEdit()
        self.supplier_contact_edit = QLineEdit()
        self.supplier_email_edit = QLineEdit()
        self.supplier_address_edit = QTextEdit()
        self.supplier_address_edit.setMaximumHeight(84)
        self.supplier_terms_edit = QLineEdit()
        self.supplier_lead_days_edit = QLineEdit()
        self.supplier_website_edit = QLineEdit()
        self.supplier_notes_edit = QTextEdit()
        self.supplier_notes_edit.setMaximumHeight(110)
        for label, widget in (
            ("ID", self.supplier_id_edit),
            ("Nome", self.supplier_name_edit),
            ("NIF", self.supplier_nif_edit),
            ("Contacto", self.supplier_contact_edit),
            ("Email", self.supplier_email_edit),
            ("Morada", self.supplier_address_edit),
            ("Cond. pagamento", self.supplier_terms_edit),
            ("Prazo entrega (dias)", self.supplier_lead_days_edit),
            ("Website", self.supplier_website_edit),
            ("Observacoes", self.supplier_notes_edit),
        ):
            form.addRow(label, widget)
        form_layout.addWidget(form_title)
        form_layout.addLayout(form)
        form_layout.addStretch(1)
        split.addWidget(form_card)
        split.setSizes([760, 760])
        root.addWidget(split, 1)
        self._new_supplier()

    def can_auto_refresh(self) -> bool:
        return False

    def refresh(self) -> None:
        previous = self.current_id
        self.rows = self.backend.ne_suppliers()
        query = self.filter_edit.text().strip().lower()
        if query:
            self.rows = [row for row in self.rows if any(query in str(value).lower() for value in row.values())]
        _fill_table(
            self.table,
            [[r.get("id", "-"), r.get("nome", "-"), r.get("nif", "-"), r.get("contacto", "-"), r.get("email", "-")] for r in self.rows],
            align_center_from=2,
        )
        for idx, row in enumerate(self.rows):
            item = self.table.item(idx, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("id", "") or "").strip())
        if not self.rows:
            self._new_supplier()
            return
        target = 0
        if previous:
            for idx, row in enumerate(self.rows):
                if str(row.get("id", "") or "").strip() == previous:
                    target = idx
                    break
        self.table.selectRow(target)
        self._load_selected_supplier()

    def _selected_supplier(self) -> dict:
        row_index = _selected_row_index(self.table)
        if row_index < 0:
            return {}
        item = self.table.item(row_index, 0)
        supplier_id = str(item.data(Qt.UserRole) or item.text() or "").strip()
        return next((row for row in self.rows if str(row.get("id", "") or "").strip() == supplier_id), {})

    def _new_supplier(self) -> None:
        self.current_id = ""
        self.supplier_id_edit.setText(self.backend.supplier_next_id())
        self.supplier_name_edit.clear()
        self.supplier_nif_edit.clear()
        self.supplier_contact_edit.clear()
        self.supplier_email_edit.clear()
        self.supplier_address_edit.clear()
        self.supplier_terms_edit.clear()
        self.supplier_lead_days_edit.clear()
        self.supplier_website_edit.clear()
        self.supplier_notes_edit.clear()

    def _load_selected_supplier(self) -> None:
        row = self._selected_supplier()
        if not row:
            return
        self.current_id = str(row.get("id", "") or "").strip()
        self.supplier_id_edit.setText(self.current_id)
        self.supplier_name_edit.setText(str(row.get("nome", "") or "").strip())
        self.supplier_nif_edit.setText(str(row.get("nif", "") or "").strip())
        self.supplier_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        self.supplier_email_edit.setText(str(row.get("email", "") or "").strip())
        self.supplier_address_edit.setPlainText(str(row.get("morada", "") or "").strip())
        self.supplier_terms_edit.setText(str(row.get("cond_pagamento", "") or "").strip())
        self.supplier_lead_days_edit.setText(str(row.get("prazo_entrega_dias", "") or "").strip())
        self.supplier_website_edit.setText(str(row.get("website", "") or "").strip())
        self.supplier_notes_edit.setPlainText(str(row.get("obs", "") or "").strip())

    def _save_supplier(self) -> None:
        try:
            saved = self.backend.supplier_save(
                {
                    "id": self.supplier_id_edit.text().strip(),
                    "nome": self.supplier_name_edit.text().strip(),
                    "nif": self.supplier_nif_edit.text().strip(),
                    "contacto": self.supplier_contact_edit.text().strip(),
                    "email": self.supplier_email_edit.text().strip(),
                    "morada": self.supplier_address_edit.toPlainText().strip(),
                    "cond_pagamento": self.supplier_terms_edit.text().strip(),
                    "prazo_entrega_dias": self.supplier_lead_days_edit.text().strip(),
                    "website": self.supplier_website_edit.text().strip(),
                    "obs": self.supplier_notes_edit.toPlainText().strip(),
                }
            )
        except Exception as exc:
            QMessageBox.critical(self, "Fornecedores", str(exc))
            return
        self.current_id = str(saved.get("id", "") or "").strip()
        self.refresh()

    def _remove_supplier(self) -> None:
        supplier_id = self.supplier_id_edit.text().strip()
        if not supplier_id:
            return
        if QMessageBox.question(self, "Fornecedores", f"Remover fornecedor {supplier_id}?") != QMessageBox.Yes:
            return
        try:
            self.backend.supplier_remove(supplier_id)
        except Exception as exc:
            QMessageBox.critical(self, "Fornecedores", str(exc))
            return
        self._new_supplier()
        self.refresh()


class OrdersPage(QWidget):
    page_title = "Encomendas"
    page_subtitle = "Encomendas com materiais, espessuras, peças e reservas organizadas por hierarquia."
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
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Todas", "Preparacao", "Montagem", "Em producao", "Concluida"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.client_combo = QComboBox()
        self.client_combo.currentTextChanged.connect(self.refresh)
        self.new_btn = QPushButton("Criar encomenda")
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
            ("Entrega", self.info_entrega, 1, 0),
            ("Nota cliente", self.info_nota, 1, 2),
            ("Cativar MP", self.info_cativar, 1, 4),
            ("Transporte", self.info_transporte, 2, 0),
            ("Descarga", self.info_descarga, 2, 2),
            ("Viagem", self.info_viagem, 2, 4),
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
        list_title = QLabel("Encomendas")
        list_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(["Numero", "Nota Cliente", "Cliente", "Data Entrega", "Tempo", "Estado", "Cativar", "Progresso"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setStyleSheet(
            f"QTableWidget {{ font-size: {LIST_TABLE_FONT_PX}px; }}"
            f" QHeaderView::section {{ font-size: {LIST_TABLE_FONT_PX}px; padding: 8px 10px; font-weight: 800; }}"
        )
        self.table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        _configure_table(self.table, stretch=(2,), contents=(3, 4, 5, 6, 7))
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 250),
                (1, "fixed", 220),
                (2, "stretch", 0),
                (3, "fixed", 132),
                (4, "fixed", 82),
                (5, "fixed", 126),
                (6, "fixed", 98),
                (7, "fixed", 102),
            ],
        )
        self.table.itemSelectionChanged.connect(self._on_order_selected)
        list_layout.addWidget(list_title)
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
        self.materials_table = QTableWidget(0, 2)
        self.materials_table.setHorizontalHeaderLabels(["Material", "Estado"])
        self.materials_table.verticalHeader().setVisible(False)
        self.materials_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.materials_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.materials_table, stretch=(0,), contents=(1,))
        _set_table_columns(
            self.materials_table,
            [
                (0, "stretch", 240),
                (1, "fixed", 130),
            ],
        )
        self.materials_table.verticalHeader().setDefaultSectionSize(30)
        self.materials_table.setStyleSheet(
            "QTableWidget { font-size: 11px; }"
            "QHeaderView::section { font-size: 11px; font-weight: 800; padding: 5px 8px; }"
        )
        self.materials_table.itemSelectionChanged.connect(self._on_material_selected)
        materials_btns = QHBoxLayout()
        self.add_material_btn = QPushButton("Adicionar")
        self.add_material_btn.clicked.connect(self._add_material)
        self.remove_material_btn = QPushButton("Remover")
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
        self.esp_table = QTableWidget(0, 4)
        self.esp_table.setHorizontalHeaderLabels(["Espessura", "Laser", "Outras ops", "Estado"])
        self.esp_table.verticalHeader().setVisible(False)
        self.esp_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.esp_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.esp_table, stretch=(2,), contents=(0, 1, 3))
        _set_table_columns(
            self.esp_table,
            [
                (0, "fixed", 88),
                (1, "fixed", 76),
                (2, "stretch", 186),
                (3, "fixed", 126),
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
        self.add_esp_btn = QPushButton("Adicionar")
        self.add_esp_btn.clicked.connect(self._add_espessura)
        self.remove_esp_btn = QPushButton("Remover")
        self.remove_esp_btn.setProperty("variant", "secondary")
        self.remove_esp_btn.clicked.connect(self._remove_espessura)
        self.edit_time_btn = QPushButton("Tempos")
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
        pieces_title = QLabel("Pecas")
        pieces_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.add_piece_btn = QPushButton("Adicionar peca")
        self.add_piece_btn.clicked.connect(self._add_piece)
        self.edit_piece_btn = QPushButton("Editar peca")
        self.edit_piece_btn.setProperty("variant", "secondary")
        self.edit_piece_btn.clicked.connect(self._edit_piece)
        self.remove_piece_btn = QPushButton("Remover peca")
        self.remove_piece_btn.setProperty("variant", "secondary")
        self.remove_piece_btn.clicked.connect(self._remove_piece)
        self.consume_montagem_btn = QPushButton("Consumir montagem")
        self.consume_montagem_btn.setProperty("variant", "success")
        self.consume_montagem_btn.clicked.connect(self._consume_montagem)
        self.open_piece_btn = QPushButton("Ver desenho")
        self.open_piece_btn.setProperty("variant", "secondary")
        self.open_piece_btn.clicked.connect(self._open_selected_piece_drawing)
        for button, width in (
            (self.add_piece_btn, 138),
            (self.edit_piece_btn, 128),
            (self.remove_piece_btn, 136),
            (self.consume_montagem_btn, 162),
            (self.open_piece_btn, 124),
        ):
            button.setMinimumWidth(width)
        pieces_header.addWidget(pieces_title, 1)
        for button in (self.add_piece_btn, self.edit_piece_btn, self.remove_piece_btn, self.consume_montagem_btn, self.open_piece_btn):
            pieces_header.addWidget(button)
        self.pieces_table = QTableWidget(0, 7)
        self.pieces_table.setHorizontalHeaderLabels(["Ref. Interna", "Ref. Externa", "Material", "Esp.", "Operacoes", "Qtd", "Estado"])
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.pieces_table, stretch=(1, 4), contents=(0, 2, 3, 5, 6))
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
        montagem_title = QLabel("Montagem / Componentes de stock")
        montagem_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.montagem_note_btn = QPushButton("Gerar NE montagem")
        self.montagem_note_btn.setProperty("variant", "secondary")
        self.montagem_note_btn.clicked.connect(self._create_montagem_purchase_note)
        self.montagem_note_btn.setMinimumWidth(154)
        montagem_header.addWidget(montagem_title, 1)
        montagem_header.addWidget(self.montagem_note_btn)
        self.montagem_meta = QLabel("Sem itens de montagem.")
        self.montagem_meta.setWordWrap(True)
        self.montagem_meta.setProperty("role", "muted")
        self.montagem_table = QTableWidget(0, 8)
        self.montagem_table.setHorizontalHeaderLabels(["Tipo", "Codigo", "Descricao", "Qtd plan.", "Consumida", "Pendente", "Falta", "Estado"])
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
                    r.get("nota_cliente", "-"),
                    r.get("cliente", "-"),
                    r.get("data_entrega", "-"),
                    r.get("tempo", "0"),
                    r.get("estado", "-"),
                    r.get("cativar", "NAO"),
                    f"{r.get('progress', 0):.1f}%",
                ]
                for r in self.rows
            ],
            align_center_from=4,
        )
        self.table.setSortingEnabled(False)
        for row_index, row in enumerate(self.rows):
            item = self.table.item(row_index, 0)
            if item is not None:
                item.setData(Qt.UserRole, str(row.get("numero", "") or "").strip())
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.table, row_index, str(row.get("estado", "")))
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
        self.edit_piece_btn.setEnabled(can_edit and has_piece)
        self.remove_piece_btn.setEnabled(can_edit and has_piece)
        self.open_piece_btn.setEnabled(has_piece)
        self.consume_montagem_btn.setEnabled(has_order and bool(self.current_detail.get("can_consume_montagem")))
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
            [[row.get("material", "-"), row.get("estado", "-")] for row in self.material_rows],
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
        header.resizeSection(3, 100)
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
                    row.get("material", "-"),
                    row.get("espessura", "-"),
                    row.get("operacoes", "-"),
                    row.get("qtd_plan", "0"),
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
            str(row.get("produto_codigo", "") or "").strip(): row
            for row in list(self.current_detail.get("montagem_shortages", []) or [])
            if str(row.get("produto_codigo", "") or "").strip()
        }
        _fill_table(
            self.montagem_table,
            [
                [
                    row.get("tipo_label", "-"),
                    row.get("produto_codigo", "-") or "-",
                    row.get("descricao", "-"),
                    f"{float(row.get('qtd_planeada', 0) or 0):.2f}",
                    f"{float(row.get('qtd_consumida', 0) or 0):.2f}",
                    f"{float(row.get('qtd_pendente', 0) or 0):.2f}",
                    f"{float((shortages_map.get(str(row.get('produto_codigo', '') or '').strip(), {}) or {}).get('qtd_em_falta', 0) or 0):.2f}",
                    row.get("estado", "-"),
                ]
                for row in self.detail_montagem
            ],
            align_center_from=3,
        )
        for row_index, row in enumerate(self.detail_montagem):
            _paint_table_row(self.montagem_table, row_index, str(row.get("estado", "")))
            code = str(row.get("produto_codigo", "") or "").strip()
            shortage = shortages_map.get(code, {})
            if not shortage:
                continue
            tip_parts = [
                f"Falta {float(shortage.get('qtd_em_falta', 0) or 0):.2f} {shortage.get('produto_unid', 'UN')}",
                f"Disponivel {float(shortage.get('qtd_disponivel', 0) or 0):.2f}",
            ]
            if str(shortage.get("fornecedor_sugerido", "") or "").strip():
                tip_parts.append(f"Sugestao {shortage.get('fornecedor_sugerido', '-')}")
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

    def _consume_montagem(self) -> None:
        numero = str(self.current_detail.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Montagem", "Seleciona uma encomenda.")
            return
        if QMessageBox.question(self, "Montagem", f"Consumir os componentes de stock pendentes da encomenda {numero}?") != QMessageBox.Yes:
            return
        try:
            detail = self.backend.order_consume_montagem(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Montagem", str(exc))
            return
        self.current_detail = detail
        self.refresh()
        self.open_order_numero(numero)
        QMessageBox.information(self, "Montagem", f"Montagem concluida para a encomenda {numero}.")

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
        dialog = QDialog(self)
        dialog.setWindowTitle("Cabecalho da encomenda")
        dialog.setMinimumWidth(640)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
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
        delivery_edit.setDisplayFormat("yyyy-MM-dd")
        delivery_raw = str(initial.get("data_entrega", "") or "").strip()
        delivery_date = QDate.fromString(delivery_raw, "yyyy-MM-dd") if delivery_raw else QDate.currentDate()
        if not delivery_date.isValid():
            delivery_date = QDate.currentDate()
        delivery_edit.setDate(delivery_date)
        posto_combo = QComboBox()
        posto_combo.setEditable(False)
        for posto in list(self.backend.quote_workcenter_options() or []):
            posto_combo.addItem(str(posto))
        posto_current = str(initial.get("posto_trabalho", "") or "").strip()
        if posto_current:
            posto_combo.setCurrentText(posto_current)
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
        form.addRow("Numero", numero_label)
        form.addRow("Cliente", client_combo)
        form.addRow("Posto trabalho", posto_combo)
        form.addRow("Transporte", transport_combo)
        form.addRow("Transportadora", carrier_combo)
        form.addRow("Zona transporte", zone_combo)
        form.addRow("Local descarga", local_descarga_edit)
        form.addRow("Preço transporte", transport_price_spin)
        form.addRow("Custo transporte", transport_cost_spin)
        form.addRow("Paletes", paletes_spin)
        form.addRow("Peso bruto", peso_spin)
        form.addRow("Volume", volume_spin)
        form.addRow("Ref. transporte", ref_transport_edit)
        form.addRow("Entrega", delivery_edit)
        form.addRow("Tempo (h)", tempo_spin)
        form.addRow("Nota cliente", note_edit)
        form.addRow("Observacoes", obs_edit)
        form.addRow("", cativar_box)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "numero": str(initial.get("numero", "") or "").strip(),
            "cliente": self._client_code_from_text(client_combo.currentText()),
            "posto_trabalho": posto_combo.currentText().strip(),
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
        keep_ref_box = QCheckBox("Guardar referencia na base")
        keep_ref_box.setChecked(bool(initial.get("guardar_ref", True)))

        def apply_reference(payload: dict | None) -> None:
            if not isinstance(payload, dict):
                return
            ref_ext_edit.setText(str(payload.get("ref_externa", "") or "").strip())
            ref_int_edit.setText(str(payload.get("ref_interna", "") or "").strip())
            desc_edit.setText(str(payload.get("descricao", "") or "").strip())
            material_combo.setCurrentText(str(payload.get("material", "") or "").strip())
            esp_combo.setCurrentText(str(payload.get("espessura", "") or "").strip())
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
            payload = _reference_catalog_dialog(self, references, "Histórico de referencias")
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

        ref_history.currentIndexChanged.connect(lambda _index: apply_reference(ref_history.currentData()))
        ref_buttons = QHBoxLayout()
        btn_generate = QPushButton("Gerar")
        btn_generate.setProperty("variant", "secondary")
        btn_generate.clicked.connect(generate_ref)
        btn_history = QPushButton("Histórico refs")
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
        draw_buttons.addWidget(btn_pick_draw)
        draw_buttons.addStretch(1)
        defaults_btn = QPushButton("Operacao padrao")
        defaults_btn.setProperty("variant", "secondary")
        defaults_btn.clicked.connect(lambda: apply_operations(str(self.presets.get("operacao_default", "Embalamento"))))

        form.addRow("Histórico", ref_history)
        form.addRow("Ref. interna", ref_int_edit)
        form.addRow("", ref_buttons)
        form.addRow("Ref. externa", ref_ext_edit)
        form.addRow("Descricao", desc_edit)
        form.addRow("Material", material_combo)
        form.addRow("Espessura", esp_combo)
        form.addRow("Postos", operation_selector)
        form.addRow("", defaults_btn)
        form.addRow("Quantidade", qtd_spin)
        form.addRow("Preco unit.", preco_spin)
        form.addRow("Desenho", drawing_edit)
        form.addRow("", draw_buttons)
        form.addRow("", keep_ref_box)
        layout.addLayout(form)
        if not ref_int_edit.text().strip():
            generate_ref()
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
            "material": material_combo.currentText().strip(),
            "espessura": esp_combo.currentText().strip(),
            "operacoes": operacoes_edit.text().strip(),
            "quantidade_pedida": qtd_spin.value(),
            "preco_unit": preco_spin.value(),
            "desenho": drawing_edit.text().strip(),
            "guardar_ref": keep_ref_box.isChecked(),
        }

    def _reserve_dialog(self, candidates: list[dict]) -> list[dict] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Cativar stock")
        dialog.setMinimumSize(940, 420)
        dialog.setStyleSheet(
            """
            QDialog { font-size: 12px; }
            QTableWidget { font-size: 12px; }
            QHeaderView::section { font-size: 12px; font-weight: 700; }
            QDoubleSpinBox { font-size: 12px; min-height: 28px; }
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
        table.verticalHeader().setDefaultSectionSize(30)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.resizeSection(4, 172)
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
            spin.setAlignment(Qt.AlignCenter)
            spin.setMinimumWidth(144)
            cell_host = QWidget()
            cell_layout = QHBoxLayout(cell_host)
            cell_layout.setContentsMargins(8, 3, 8, 3)
            cell_layout.addWidget(spin)
            table.setCellWidget(row_index, 4, cell_host)
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
        try:
            detail = self.backend.order_create_or_update(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Encomendas", str(exc))
            return
        self.refresh()
        self._select_order(str(detail.get("numero", "") or "").strip())

    def _edit_order_header(self) -> None:
        if not self.current_detail.get("numero"):
            QMessageBox.warning(self, "Encomendas", "Seleciona uma encomenda.")
            return
        payload = self._header_dialog(self.current_detail)
        if payload is None:
            return
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
            "Define só as operações desta espessura. O planeamento por operação usa estes minutos automaticamente."
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
            form.addRow(f"{op_name}", spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            payload = {}
            for op_name, spin in editors.items():
                minutes = int(spin.value())
                if minutes > 0:
                    payload[op_name] = str(minutes)
            self.backend.order_espessura_set_operation_times(numero, material, espessura, payload)
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


def _take_layout_items(layout) -> list:
    items = []
    while layout.count():
        items.append(layout.takeAt(0))
    return items


def _adopt_layout_item(target_layout, item, stretch: int = 0) -> None:
    if item is None:
        return
    widget = item.widget()
    child_layout = item.layout()
    spacer = item.spacerItem()
    if isinstance(target_layout, QSplitter):
        if widget is not None:
            target_layout.addWidget(widget)
            return
        if child_layout is not None:
            host = QWidget()
            host.setLayout(child_layout)
            target_layout.addWidget(host)
            return
        return
    if widget is not None:
        target_layout.addWidget(widget, stretch)
        return
    if child_layout is not None:
        target_layout.addLayout(child_layout, stretch)
        return
    if spacer is not None:
        target_layout.addItem(spacer)


def _clear_layout_widgets(layout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        child_layout = item.layout()
        if widget is not None:
            widget.deleteLater()
        elif child_layout is not None:
            _clear_layout_widgets(child_layout)


class LegacyPurchaseNotesPage(PurchaseNotesPage):
    page_subtitle = "Lista de notas com detalhe completo apenas ao abrir o registo."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        root = self.layout()
        sections = _take_layout_items(root)
        filters_item = sections[0] if len(sections) > 0 else None
        notes_item = sections[1] if len(sections) > 1 else None
        form_item = sections[2] if len(sections) > 2 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(14)
        _adopt_layout_item(list_layout, filters_item)
        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(16, 12, 16, 12)
        list_actions_layout.setSpacing(8)
        self.open_note_btn = QPushButton("Abrir nota")
        self.open_note_btn.clicked.connect(self._open_selected_note)
        create_note_btn = QPushButton("Criar nota")
        create_note_btn.clicked.connect(self._create_new_note)
        self.remove_note_list_btn = QPushButton("Apagar nota")
        self.remove_note_list_btn.setProperty("variant", "danger")
        self.remove_note_list_btn.clicked.connect(self._remove_note)
        refresh_note_btn = QPushButton("Atualizar")
        refresh_note_btn.setProperty("variant", "secondary")
        refresh_note_btn.clicked.connect(self.refresh)
        list_actions_layout.addWidget(create_note_btn)
        list_actions_layout.addWidget(self.open_note_btn)
        list_actions_layout.addWidget(self.remove_note_list_btn)
        list_actions_layout.addWidget(refresh_note_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        _adopt_layout_item(list_layout, notes_item, 1)

        self.detail_page = QWidget()
        detail_outer = QVBoxLayout(self.detail_page)
        detail_outer.setContentsMargins(0, 0, 0, 0)
        detail_outer.setSpacing(0)
        self.note_detail_scroll = QScrollArea()
        self.note_detail_scroll.setWidgetResizable(True)
        self.note_detail_scroll.setFrameShape(QFrame.NoFrame)
        self.note_detail_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.note_detail_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        detail_outer.addWidget(self.note_detail_scroll)
        self.note_detail_host = QWidget()
        self.note_detail_scroll.setWidget(self.note_detail_host)
        detail_layout = QVBoxLayout(self.note_detail_host)
        detail_layout.setContentsMargins(0, 0, 0, 8)
        detail_layout.setSpacing(14)
        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(16, 12, 16, 12)
        detail_actions_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a lista")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_note_list)
        detail_actions_layout.addWidget(back_btn)
        detail_actions_layout.addStretch(1)
        detail_layout.addWidget(detail_actions)
        _adopt_layout_item(detail_layout, form_item, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.notes_table.itemSelectionChanged.connect(self._sync_list_buttons)
        self.notes_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_note())
        self._show_note_list()
        self._sync_list_buttons()

    def refresh(self) -> None:
        previous = self.current_number
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(previous)
        self.supplier_rows = self.backend.ne_suppliers()
        self._set_supplier_items()
        self.rows = self.backend.ne_rows(self.filter_edit.currentText().strip(), self.state_combo.currentText())
        _fill_table(
            self.notes_table,
            [[r.get("numero", "-"), r.get("fornecedor", "-"), r.get("data_entrega", "-"), r.get("estado", "-"), _fmt_eur(r.get("total", 0)), r.get("linhas", 0)] for r in self.rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.notes_table, row_index, str(row.get("estado", "")))
        if self.notes_table.rowCount() == 0:
            PurchaseNotesPage._new_note(self, reset_number=False)
            self._show_note_list()
            self._sync_list_buttons()
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "")).strip() == previous:
                    row_index = index
                    break
        self.notes_table.selectRow(row_index)
        self._load_selected_note()
        if keep_detail:
            self._show_note_detail()
        else:
            self._show_note_list()
        self._sync_list_buttons()

    def _show_note_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_buttons()

    def _show_note_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _sync_list_buttons(self) -> None:
        has_selection = bool(self._selected_row())
        self.open_note_btn.setEnabled(has_selection)
        if hasattr(self, "remove_note_list_btn"):
            self.remove_note_list_btn.setEnabled(has_selection)

    def _open_selected_note(self) -> None:
        if not self._selected_row():
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma nota.")
            return
        self._load_selected_note()
        self._show_note_detail()

    def _new_note(self, reset_number: bool = True) -> None:
        PurchaseNotesPage._new_note(self, reset_number=reset_number)
        if hasattr(self, "view_stack"):
            self._show_note_detail()

    def _remove_note(self) -> None:
        PurchaseNotesPage._remove_note(self)
        self._show_note_list()
        self._sync_list_buttons()


class LegacyOrdersPage(OrdersPage):
    page_subtitle = "Lista de encomendas primeiro e detalhe completo apenas dentro da encomenda."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.materials_table.verticalHeader().setDefaultSectionSize(24)
        self.esp_table.verticalHeader().setDefaultSectionSize(24)
        self.pieces_table.verticalHeader().setDefaultSectionSize(24)
        root = self.layout()
        sections = _take_layout_items(root)
        filters_item = sections[0] if len(sections) > 0 else None
        info_item = sections[1] if len(sections) > 1 else None
        list_item = sections[2] if len(sections) > 2 else None
        mid_item = sections[3] if len(sections) > 3 else None
        pieces_item = sections[4] if len(sections) > 4 else None
        montagem_item = sections[5] if len(sections) > 5 else None
        if filters_item and filters_item.widget() is not None:
            filters_item.widget().hide()

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)

        filters_card = CardFrame()
        filters_card.set_tone("info")
        filters_layout = QGridLayout(filters_card)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setHorizontalSpacing(10)
        filters_layout.setVerticalSpacing(8)
        _cap_width(self.filter_edit, 420)
        _cap_width(self.state_combo, 150)
        _cap_width(self.year_combo, 110)
        _cap_width(self.client_combo, 360)
        for button, width in ((self.new_btn, 150), (self.edit_header_btn, 122), (self.remove_btn, 122)):
            button.setMinimumWidth(width)
        filters_layout.addWidget(QLabel("Pesquisa"), 0, 0)
        filters_layout.addWidget(self.filter_edit, 0, 1)
        filters_layout.addWidget(QLabel("Estado"), 0, 2)
        filters_layout.addWidget(self.state_combo, 0, 3)
        filters_layout.addWidget(QLabel("Ano"), 0, 4)
        filters_layout.addWidget(self.year_combo, 0, 5)
        filters_layout.addWidget(QLabel("Cliente"), 1, 0)
        filters_layout.addWidget(self.client_combo, 1, 1, 1, 3)
        filters_layout.addWidget(self.new_btn, 1, 3)
        filters_layout.addWidget(self.edit_header_btn, 1, 4)
        filters_layout.addWidget(self.remove_btn, 1, 5)
        list_layout.addWidget(filters_card)

        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(14, 10, 14, 10)
        list_actions_layout.setSpacing(8)
        self.open_order_btn = QPushButton("Abrir encomenda")
        self.open_order_btn.clicked.connect(self._open_selected_order)
        refresh_order_btn = QPushButton("Atualizar")
        refresh_order_btn.setProperty("variant", "secondary")
        refresh_order_btn.clicked.connect(self.refresh)
        self.open_order_btn.setMinimumWidth(146)
        refresh_order_btn.setMinimumWidth(104)
        list_actions_layout.addWidget(self.open_order_btn)
        list_actions_layout.addWidget(refresh_order_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        _adopt_layout_item(list_layout, list_item, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)
        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(14, 10, 14, 10)
        detail_actions_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a lista")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_order_list)
        edit_btn = QPushButton("Editar cabecalho")
        edit_btn.clicked.connect(self._edit_order_header)
        remove_btn = QPushButton("Remover encomenda")
        remove_btn.setProperty("variant", "danger")
        remove_btn.clicked.connect(self._remove_order)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.refresh)
        for button, width in ((back_btn, 126), (edit_btn, 144), (remove_btn, 170), (refresh_btn, 108)):
            button.setMinimumWidth(width)
            detail_actions_layout.addWidget(button)
        detail_actions_layout.addStretch(1)
        detail_layout.addWidget(detail_actions)

        content_split = QSplitter(Qt.Vertical)
        content_split.setChildrenCollapsible(False)
        top_host = QWidget()
        top_layout = QVBoxLayout(top_host)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)
        _adopt_layout_item(top_layout, info_item)
        bottom_split = QSplitter(Qt.Horizontal)
        bottom_split.setChildrenCollapsible(False)
        left_host = QWidget()
        left_layout = QVBoxLayout(left_host)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        _adopt_layout_item(left_layout, mid_item, 1)
        right_split = QSplitter(Qt.Vertical)
        right_split.setChildrenCollapsible(False)
        pieces_host = QWidget()
        pieces_layout = QVBoxLayout(pieces_host)
        pieces_layout.setContentsMargins(0, 0, 0, 0)
        pieces_layout.setSpacing(0)
        _adopt_layout_item(pieces_layout, pieces_item, 1)
        right_split.addWidget(pieces_host)
        if montagem_item is not None:
            montagem_host = QWidget()
            montagem_layout = QVBoxLayout(montagem_host)
            montagem_layout.setContentsMargins(0, 0, 0, 0)
            montagem_layout.setSpacing(0)
            _adopt_layout_item(montagem_layout, montagem_item, 1)
            right_split.addWidget(montagem_host)
            right_split.setSizes([520, 240])
        bottom_split.addWidget(left_host)
        bottom_split.addWidget(right_split)
        bottom_split.setSizes([560, 1340])
        content_split.addWidget(top_host)
        content_split.addWidget(bottom_split)
        content_split.setSizes([150, 690])
        detail_layout.addWidget(content_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.table.itemSelectionChanged.connect(self._sync_list_open_button)
        self.table.itemDoubleClicked.connect(lambda item: self._open_selected_order(item))
        self._show_order_list()
        self._sync_list_open_button()

    def refresh(self) -> None:
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(self.current_detail.get("numero"))
        OrdersPage.refresh(self)
        if self.table.rowCount() == 0:
            self._show_order_list()
        elif keep_detail:
            self._show_order_detail()
        else:
            self._show_order_list()
        self._sync_list_open_button()

    def _show_order_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_open_button()

    def _show_order_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

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
        self._show_order_detail()

    def _select_order(self, numero: str, piece_ref: str = "", material: str = "", espessura: str = "") -> None:
        OrdersPage._select_order(self, numero, piece_ref=piece_ref, material=material, espessura=espessura)
        self._show_order_detail()

    def _remove_order(self) -> None:
        OrdersPage._remove_order(self)
        self._show_order_list()
        self._sync_list_open_button()


class LegacyOperatorPage(OperatorPage):
    page_subtitle = "Encomendas ativas primeiro e detalhe operacional apenas dentro da encomenda."

    def __init__(self, runtime_service, backend, parent=None) -> None:
        super().__init__(runtime_service, backend, parent)
        self.order_rows: list[dict] = []
        self.order_rows_all: list[dict] = []
        self.selected_order_number = ""
        for card in self.cards:
            card.setMaximumHeight(112)
        self.global_progress.setMaximumHeight(18)
        self.groups_table.verticalHeader().setDefaultSectionSize(22)
        self.pieces_table.verticalHeader().setDefaultSectionSize(18)
        self.groups_table.setColumnCount(7)
        self.groups_table.setHorizontalHeaderLabels(["Enc.", "Cli.", "Estado", "Mat.", "Esp.", "Real", "%"])
        self.groups_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        group_header = self.groups_table.horizontalHeader()
        group_header.setStretchLastSection(True)
        for col, width in ((0, 132), (1, 84), (2, 106), (3, 94), (4, 56), (5, 64), (6, 74)):
            group_header.setSectionResizeMode(col, QHeaderView.Interactive)
            group_header.resizeSection(col, width)
        self.control_card.setMaximumHeight(92)
        self.context_card.setMaximumHeight(118)
        self.feedback_label.setMaximumHeight(24)
        for widget, width in ((self.operator_combo, 180), (self.posto_combo, 132), (self.operation_combo, 230)):
            widget.setProperty("compact", "true")
            widget.setMinimumWidth(width)
        for button in (self.start_btn, self.finish_btn, self.resume_btn, self.pause_btn, self.avaria_btn, self.close_avaria_btn, self.alert_btn, self.manual_consume_btn, self.drawing_btn, self.labels_btn, self.local_refresh_btn):
            button.setProperty("compact", "true")
            button.setMinimumWidth(0)
            button.setMinimumHeight(28)
        self.select_all_pieces_box.setProperty("compact", "true")
        self.piece_state_chip.setMinimumWidth(92)
        self.piece_state_chip.setAlignment(Qt.AlignCenter)
        self.piece_title_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.piece_meta_label.setStyleSheet("font-size: 11px;")
        self.pending_label.setStyleSheet("font-size: 10px; color: #5b6f86;")
        context_layout = self.context_card.layout()
        context_layout.setContentsMargins(10, 8, 10, 8)
        context_layout.setSpacing(4)
        self.issue_label.setMaximumHeight(18)
        self.pending_label.setMaximumHeight(22)
        self.operation_strip.setMaximumHeight(22)
        self.piece_progress.setMaximumHeight(14)
        self.operator_combo.setMinimumWidth(150)
        self.posto_combo.setMinimumWidth(110)
        self.operation_combo.setMinimumWidth(170)
        piece_header = self.pieces_table.horizontalHeader()
        for col, width in ((0, 34), (1, 140), (2, 320), (3, 96), (4, 126), (5, 104), (6, 74), (7, 66), (8, 66), (9, 126), (10, 360)):
            piece_header.setSectionResizeMode(col, QHeaderView.Interactive)
            piece_header.resizeSection(col, width)
        piece_header.setStretchLastSection(True)
        group_header = self.groups_table.horizontalHeader()
        for col, width in ((0, 138), (1, 116), (2, 108), (3, 138), (4, 60), (5, 62), (6, 62), (7, 70), (8, 86)):
            group_header.setSectionResizeMode(col, QHeaderView.Interactive)
            group_header.resizeSection(col, width)
        self.pieces_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pieces_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.pieces_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.groups_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.groups_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.groups_table.setStyleSheet("font-size: 11px;")
        self.pieces_table.setStyleSheet("font-size: 11px;")

        control_layout = self.control_card.layout()
        control_layout.setContentsMargins(8, 6, 8, 6)
        control_layout.setSpacing(4)
        control_items = _take_layout_items(control_layout)
        selectors_item = control_items[0] if len(control_items) > 0 else None
        feedback_item = control_items[2] if len(control_items) > 2 else None
        selectors_host = QWidget()
        selectors_grid = QGridLayout(selectors_host)
        selectors_grid.setContentsMargins(0, 0, 0, 0)
        selectors_grid.setHorizontalSpacing(6)
        selectors_grid.setVerticalSpacing(2)
        op_label = QLabel("Operador")
        posto_label = QLabel("Posto")
        oper_label = QLabel("Operacao")
        for label in (op_label, posto_label, oper_label):
            label.setStyleSheet("font-size: 11px; color: #334155;")
        selectors_grid.addWidget(op_label, 0, 0)
        selectors_grid.addWidget(posto_label, 0, 1)
        selectors_grid.addWidget(oper_label, 0, 2)
        selectors_grid.addWidget(self.operator_combo, 1, 0)
        selectors_grid.addWidget(self.posto_combo, 1, 1)
        selectors_grid.addWidget(self.operation_combo, 1, 2)
        selectors_grid.setColumnStretch(0, 3)
        selectors_grid.setColumnStretch(1, 2)
        selectors_grid.setColumnStretch(2, 3)
        control_layout.addWidget(selectors_host)
        _adopt_layout_item(control_layout, feedback_item)
        self.control_card.setMaximumHeight(98)
        self.feedback_label.setStyleSheet("font-size: 11px; color: #475467;")

        self.orders_table = QTableWidget(0, 8)
        self.orders_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Estado", "Grupos", "Pecas", "Em curso", "Avarias", "Progress"])
        self.orders_table.verticalHeader().setVisible(False)
        self.orders_table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        self.orders_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.orders_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.orders_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.orders_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.orders_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.orders_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        _configure_table(self.orders_table, stretch=(1,), contents=())
        orders_header = self.orders_table.horizontalHeader()
        for col, width in ((0, 192), (1, 430), (2, 126), (3, 66), (4, 64), (5, 82), (6, 78), (7, 84)):
            orders_header.setSectionResizeMode(col, QHeaderView.Interactive)
            orders_header.resizeSection(col, width)
        orders_header.setSectionResizeMode(1, QHeaderView.Stretch)
        self.orders_table.setStyleSheet(
            f"QTableWidget {{ font-size: {LIST_TABLE_FONT_PX}px; }}"
            f" QHeaderView::section {{ font-size: {LIST_TABLE_FONT_PX}px; padding: 8px 10px; font-weight: 800; }}"
        )
        self.orders_table.itemDoubleClicked.connect(lambda item: self._open_selected_order_from_item(item))

        root = self.layout()
        sections = _take_layout_items(root)
        stats_item = sections[0] if len(sections) > 0 else None
        progress_item = sections[1] if len(sections) > 1 else None
        control_item = sections[2] if len(sections) > 2 else None
        context_item = sections[3] if len(sections) > 3 else None
        groups_item = sections[4] if len(sections) > 4 else None
        pieces_item = sections[5] if len(sections) > 5 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)
        _adopt_layout_item(list_layout, stats_item)
        _adopt_layout_item(list_layout, progress_item)
        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(14, 10, 14, 10)
        list_actions_layout.setSpacing(8)
        self.open_order_btn = QPushButton("Abrir encomenda")
        self.open_order_btn.clicked.connect(self._open_selected_order)
        refresh_group_btn = QPushButton("Atualizar")
        refresh_group_btn.setProperty("variant", "secondary")
        refresh_group_btn.clicked.connect(self.refresh)
        self.open_order_btn.setMinimumWidth(152)
        refresh_group_btn.setMinimumWidth(108)
        list_actions_layout.addWidget(self.open_order_btn)
        list_actions_layout.addWidget(refresh_group_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        list_filters = CardFrame()
        list_filters.set_tone("default")
        list_filters_layout = QHBoxLayout(list_filters)
        list_filters_layout.setContentsMargins(12, 8, 12, 8)
        list_filters_layout.setSpacing(8)
        list_filters_layout.addWidget(QLabel("Estado"))
        self.orders_state_filter_combo = QComboBox()
        self.orders_state_filter_combo.setProperty("compact", "true")
        self.orders_state_filter_combo.addItems(["Todas", "Em producao", "Concluida", "Em pausa", "Avaria", "Preparacao"])
        self.orders_state_filter_combo.currentTextChanged.connect(self._refresh_orders_list_view)
        list_filters_layout.addWidget(self.orders_state_filter_combo)
        list_filters_layout.addWidget(QLabel("Pesquisa"))
        self.orders_search_edit = QLineEdit()
        self.orders_search_edit.setProperty("compact", "true")
        self.orders_search_edit.setPlaceholderText("Filtrar encomenda ou cliente...")
        self.orders_search_edit.textChanged.connect(self._refresh_orders_list_view)
        list_filters_layout.addWidget(self.orders_search_edit, 1)
        _cap_width(self.orders_state_filter_combo, 152)
        list_layout.addWidget(list_filters)
        self.orders_card = CardFrame()
        self.orders_card.set_tone("info")
        orders_layout = QVBoxLayout(self.orders_card)
        orders_layout.setContentsMargins(14, 12, 14, 12)
        orders_layout.setSpacing(8)
        orders_header = QHBoxLayout()
        orders_title = QLabel("Encomendas ativas")
        orders_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        orders_hint = QLabel("Entrar por encomenda e operar tudo dentro do detalhe.")
        orders_hint.setProperty("role", "muted")
        orders_header.addWidget(orders_title)
        orders_header.addStretch(1)
        orders_header.addWidget(orders_hint)
        orders_layout.addLayout(orders_header)
        orders_layout.addWidget(self.orders_table)
        list_layout.addWidget(self.orders_card, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(6)
        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(10, 6, 10, 6)
        detail_actions_layout.setSpacing(6)
        back_btn = QPushButton("Voltar a encomendas")
        back_btn.setProperty("variant", "secondary")
        back_btn.setMinimumWidth(166)
        back_btn.clicked.connect(self._show_order_list)
        focus_block = QVBoxLayout()
        focus_block.setSpacing(2)
        self.order_focus_label = QLabel("Sem encomenda selecionada")
        self.order_focus_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.order_meta_label = QLabel("Seleciona uma encomenda para entrar no detalhe.")
        self.order_meta_label.setProperty("role", "muted")
        self.order_meta_label.setStyleSheet("font-size: 11px;")
        self.order_state_chip = QLabel("-")
        _apply_state_chip(self.order_state_chip, "-")
        self.order_state_chip.setMinimumWidth(124)
        self.order_state_chip.setAlignment(Qt.AlignCenter)
        focus_block.addWidget(self.order_focus_label)
        focus_block.addWidget(self.order_meta_label)
        detail_actions_layout.addWidget(back_btn)
        detail_actions_layout.addLayout(focus_block, 1)
        right_header_actions = QHBoxLayout()
        right_header_actions.setSpacing(6)
        for button, width in ((self.avaria_btn, 118), (self.close_avaria_btn, 108), (self.alert_btn, 112)):
            button.setMinimumWidth(width)
            right_header_actions.addWidget(button)
        detail_actions_layout.addLayout(right_header_actions)
        detail_actions_layout.addWidget(self.order_state_chip, 0, Qt.AlignTop)
        detail_actions.setMaximumHeight(62)
        detail_layout.addWidget(detail_actions)

        self.detail_filter_card = CardFrame()
        self.detail_filter_card.set_tone("info")
        detail_filter_layout = QHBoxLayout(self.detail_filter_card)
        detail_filter_layout.setContentsMargins(10, 6, 10, 6)
        detail_filter_layout.setSpacing(5)
        self.detail_active_chip = QLabel("-")
        self.detail_late_chip = QLabel("-")
        self.detail_running_chip = QLabel("-")
        self.detail_groups_chip = QLabel("-")
        self.detail_total_chip = QLabel("-")
        for chip in (
            self.detail_active_chip,
            self.detail_late_chip,
            self.detail_running_chip,
            self.detail_groups_chip,
            self.detail_total_chip,
        ):
            chip.setMinimumWidth(74)
            chip.setAlignment(Qt.AlignCenter)
            detail_filter_layout.addWidget(chip)
        self.detail_state_filter_combo = QComboBox()
        self.detail_state_filter_combo.setProperty("compact", "true")
        self.detail_state_filter_combo.addItems(["Todas", "Em producao", "Concluida", "Em pausa", "Avaria"])
        self.detail_state_filter_combo.currentTextChanged.connect(self._handle_group_selection)
        self.detail_search_edit = QLineEdit()
        self.detail_search_edit.setProperty("compact", "true")
        self.detail_search_edit.setPlaceholderText("Pesquisar ref...")
        self.detail_search_edit.textChanged.connect(self._handle_group_selection)
        for widget, width in ((self.detail_state_filter_combo, 152), (self.detail_search_edit, 360)):
            _cap_width(widget, width)
        detail_filter_layout.addStretch(1)
        detail_filter_layout.addWidget(self.detail_state_filter_combo)
        detail_filter_layout.addWidget(self.detail_search_edit)
        self.detail_filter_card.setMaximumHeight(46)
        detail_layout.addWidget(self.detail_filter_card)

        control_host = QWidget()
        control_host.setMinimumWidth(390)
        control_host.setMaximumWidth(460)
        control_host_layout = QVBoxLayout(control_host)
        control_host_layout.setContentsMargins(0, 0, 0, 0)
        control_host_layout.setSpacing(6)
        _adopt_layout_item(control_host_layout, control_item)

        context_host = QWidget()
        context_host_layout = QVBoxLayout(context_host)
        context_host_layout.setContentsMargins(0, 0, 0, 0)
        context_host_layout.setSpacing(6)
        _adopt_layout_item(context_host_layout, context_item)

        workspace_split = QSplitter(Qt.Horizontal)
        workspace_split.setChildrenCollapsible(False)
        workspace_split.addWidget(control_host)
        workspace_split.addWidget(context_host)
        workspace_split.setSizes([430, 1570])
        workspace_split.setMaximumHeight(174)
        detail_layout.addWidget(workspace_split)

        self.groups_title_label.setText("Grupos (3 linhas)")
        self.groups_title_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.pieces_title_label.setText("Pecas do grupo")
        self.pieces_title_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.groups_table.setMinimumHeight(_table_visible_height(self.groups_table, 3, extra=8))
        self.groups_table.setMaximumHeight(_table_visible_height(self.groups_table, 3, extra=8))
        self.pieces_table.setMinimumHeight(900)
        self.pieces_table.setSizeAdjustPolicy(QAbstractItemView.AdjustIgnored)
        self.pieces_card.setMinimumHeight(950)
        pieces_layout = self.pieces_card.layout()
        if pieces_layout is not None and pieces_layout.count() > 0:
            header_item = pieces_layout.itemAt(0)
            header_layout = header_item.layout() if header_item is not None else None
            if header_layout is not None:
                header_layout.setSpacing(6)
                _take_layout_items(header_layout)
                for button, width in (
                    (self.start_btn, 92),
                    (self.finish_btn, 92),
                    (self.resume_btn, 92),
                    (self.pause_btn, 102),
                    (self.manual_consume_btn, 96),
                    (self.drawing_btn, 102),
                    (self.labels_btn, 98),
                    (self.local_refresh_btn, 96),
                ):
                    button.setMinimumWidth(width)
                self.pause_btn.setProperty("variant", "secondary")
                self.local_refresh_btn.setProperty("variant", "secondary")
                header_layout.addWidget(self.pieces_title_label)
                header_layout.addStretch(1)
                header_layout.addWidget(self.start_btn)
                header_layout.addWidget(self.finish_btn)
                header_layout.addWidget(self.resume_btn)
                header_layout.addWidget(self.pause_btn)
                header_layout.addWidget(self.manual_consume_btn)
                header_layout.addWidget(self.drawing_btn)
                header_layout.addWidget(self.labels_btn)
                header_layout.addWidget(self.local_refresh_btn)
                header_layout.addWidget(self.multi_count_chip)
                header_layout.addWidget(self.select_all_pieces_box)
        groups_widget = groups_item.widget()
        if groups_widget is not None:
            groups_widget.setMaximumHeight(_table_visible_height(self.groups_table, 3, extra=34))
        _adopt_layout_item(detail_layout, groups_item)
        _adopt_layout_item(detail_layout, pieces_item, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.orders_table.itemSelectionChanged.connect(self._sync_order_focus)
        self.orders_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_order())
        self.groups_table.itemSelectionChanged.connect(self._sync_group_focus)
        self.groups_table.itemDoubleClicked.connect(lambda *_args: self._show_order_detail())
        self._show_order_list()
        self._sync_order_focus()

    def _render_groups_table(self, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        self.groups_table.setSortingEnabled(False)
        self.groups_table.blockSignals(True)
        self.groups_table.setRowCount(len(self.items))
        for row_index, item in enumerate(self.items):
            group_piece_stats = [_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")) for piece in list(item.get("pieces", []) or [])]
            group_progress = round(sum(float(stat.get("progress_pct", 0) or 0) for stat in group_piece_stats) / len(group_piece_stats), 1) if group_piece_stats else float(item.get("progress_pct", 0) or 0)
            row_values = [
                item.get("encomenda", "-"),
                item.get("cliente", "-"),
                item.get("estado_espessura", item.get("estado", "-")),
                item.get("material", "-"),
                item.get("espessura", "-"),
                f"{item.get('tempo_real_min', 0):.1f}",
                f"{group_progress:.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, "|".join(self._group_key(item)))
                if col_index >= 4:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.groups_table.setItem(row_index, col_index, cell)
            _paint_table_row(self.groups_table, row_index, str(item.get("estado_espessura", "")))
        self.groups_table.blockSignals(False)
        self.groups_table.setSortingEnabled(False)
        target_row = 0
        if previous_group:
            for index, item in enumerate(self.items):
                if self._group_key(item) == previous_group:
                    target_row = index
                    break
        if self.groups_table.rowCount() > 0:
            self.groups_table.selectRow(target_row)
            self.selected_group_key = self._group_key(self.items[target_row]) if target_row < len(self.items) else None
            self.selected_piece_id = previous_piece
            self._handle_group_selection()
        else:
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self.pieces_title_label.setText("Pecas da encomenda")
            self._clear_piece_context()

    def refresh(self) -> None:
        show_detail = self.view_stack.currentWidget() is self.detail_page
        keep_detail = show_detail and bool(self.selected_order_number)
        previous_order = self.selected_order_number or str(self._current_order_row().get("encomenda", "") or "").strip()
        previous_group = self.selected_group_key
        previous_piece = self.selected_piece_id
        user = self.backend.user or {}
        getter = getattr(self.backend, "ui_options", None)
        if callable(getter):
            try:
                self.ui_options = dict(getter() or {})
            except Exception:
                self.ui_options = {}
        self._set_combo_items(
            self.operator_combo,
            self.backend.operator_names(),
            preferred=str(user.get("username", "") or "").strip(),
        )
        self._sync_operator_assignment()
        data = self.runtime_service.operator_board(username=str(user.get("username", "")), role=str(user.get("role", "")))
        summary = data.get("summary", {})
        all_items = self._hydrate_operator_items(list(data.get("items", [])))
        op_total = 0
        op_done = 0
        op_running = 0
        for board_item in all_items:
            for piece in list(board_item.get("pieces", []) or []):
                stats = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
                op_total += int(stats.get("total", 0) or 0)
                op_done += int(stats.get("done", 0) or 0)
                op_running += int(stats.get("running", 0) or 0)
        global_ops_progress = 0.0 if op_total <= 0 else round(((op_done + (0.5 * op_running)) / op_total) * 100.0, 1)
        self.cards[0].set_data(summary.get("encomendas_ativas", 0), f"Grupos {summary.get('grupos', 0)}")
        self.cards[1].set_data(summary.get("pecas_em_curso", 0), f"Em pausa {summary.get('pecas_em_pausa', 0)}")
        self.cards[2].set_data(summary.get("pecas_em_avaria", 0), f"Concluidas {summary.get('pecas_concluidas', 0)}")
        self.cards[3].set_data(f"{global_ops_progress:.1f}%", f"Ops {op_done}/{op_total} | Em curso {op_running}")
        self.global_progress.setValue(int(round(global_ops_progress)))
        self.all_items = all_items
        self.order_rows_all = self._build_order_rows()
        self._render_orders_table(previous_order)
        if not self.order_rows_all:
            self.selected_order_number = ""
            self.items = []
            self.current_pieces = []
            self.groups_table.setRowCount(0)
            self.pieces_table.setRowCount(0)
            self._clear_piece_context()
            self._show_order_list()
            return
        if show_detail and previous_order:
            self._apply_order_filter(previous_order, previous_group=previous_group, previous_piece=previous_piece)
            self._show_order_detail()
        else:
            self.items = []
            self.current_pieces = []
            self.selected_group_key = None
            self.selected_piece_id = ""
            self._show_order_list()
        self._sync_order_focus()

    def _show_order_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._refresh_orders_list_view()
        self._sync_order_focus()

    def _show_order_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _current_order_row(self) -> dict:
        row_index = _selected_row_index(self.orders_table)
        if row_index < 0:
            return {}
        row_item = self.orders_table.item(row_index, 0)
        numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if numero:
            for row in self.order_rows:
                if str(row.get("encomenda", "") or "").strip() == numero:
                    return row
        if row_index >= len(self.order_rows):
            return {}
        return self.order_rows[row_index]

    def _build_order_rows(self) -> list[dict]:
        rows_map: dict[str, dict] = {}
        for item in self.all_items:
            numero = str(item.get("encomenda", "") or "").strip()
            if not numero:
                continue
            row = rows_map.setdefault(
                numero,
                {
                    "encomenda": numero,
                    "cliente": str(item.get("cliente_label", item.get("cliente", "")) or "-"),
                    "cliente_codigo": str(item.get("cliente_codigo", "") or "").strip(),
                    "cliente_nome": str(item.get("cliente_nome", "") or "").strip(),
                    "grupos": 0,
                    "pecas": 0,
                    "em_curso": 0,
                    "avarias": 0,
                    "progress_samples": [],
                    "states": [],
                },
            )
            row["grupos"] += 1
            row["states"].append(str(item.get("estado_espessura", item.get("estado", "")) or ""))
            row["progress_samples"].append(float(item.get("progress_pct", 0) or 0))
            pieces = list(item.get("pieces", []) or [])
            row["pecas"] += len(pieces)
            for piece in pieces:
                state = str(piece.get("estado", "") or "").strip()
                row["states"].append(state)
                row["progress_samples"].append(float(_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")).get("progress_pct", item.get("progress_pct", 0)) or 0))
                lowered = state.lower()
                if "produc" in lowered or "curso" in lowered:
                    row["em_curso"] += 1
                if "avaria" in lowered:
                    row["avarias"] += 1
        rows: list[dict] = []
        for row in rows_map.values():
            samples = [float(value or 0) for value in row.get("progress_samples", [])]
            row["progress_pct"] = round(sum(samples) / len(samples), 1) if samples else 0.0
            row["estado"] = self._aggregate_order_state(list(row.get("states", []) or []))
            rows.append(row)
        rows.sort(key=lambda item: str(item.get("encomenda", "") or ""))
        return rows

    def _filtered_order_rows(self) -> list[dict]:
        rows = list(self.order_rows_all)
        state_filter = self.orders_state_filter_combo.currentText().strip().lower() if hasattr(self, "orders_state_filter_combo") else "todas"
        search = self.orders_search_edit.text().strip().lower() if hasattr(self, "orders_search_edit") else ""
        if state_filter and state_filter != "todas":
            filtered_by_state: list[dict] = []
            for row in rows:
                state = str(row.get("estado", "") or "").strip().lower()
                if state_filter == "em producao" and not ("produc" in state or "curso" in state):
                    continue
                if state_filter == "concluida" and "concl" not in state:
                    continue
                if state_filter == "em pausa" and not ("paus" in state or "interromp" in state):
                    continue
                if state_filter == "avaria" and "avaria" not in state:
                    continue
                if state_filter == "preparacao" and not ("prepar" in state or "pend" in state or "edicao" in state):
                    continue
                filtered_by_state.append(row)
            rows = filtered_by_state
        if search:
            rows = [row for row in rows if search in str(row.get("encomenda", "") or "").lower() or search in str(row.get("cliente", "") or "").lower() or search in str(row.get("estado", "") or "").lower()]
        return rows

    def _aggregate_order_state(self, states: list[str]) -> str:
        lowered = [str(state or "").strip().lower() for state in states if str(state or "").strip()]
        if any("avaria" in state for state in lowered):
            return "Avaria"
        if any("produc" in state or "curso" in state for state in lowered):
            return "Em producao"
        if any("paus" in state or "interromp" in state for state in lowered):
            return "Em pausa"
        if lowered and all("concl" in state for state in lowered):
            return "Concluida"
        if any("prepar" in state or "pend" in state or "edicao" in state for state in lowered):
            return "Preparacao"
        return states[0] if states else "-"

    def _render_orders_table(self, selected_order: str = "") -> None:
        self.order_rows = self._filtered_order_rows()
        self.orders_table.setSortingEnabled(False)
        self.orders_table.blockSignals(True)
        self.orders_table.setRowCount(len(self.order_rows))
        for row_index, row in enumerate(self.order_rows):
            row_values = [
                row.get("encomenda", "-"),
                _format_client_label(row.get("cliente", "-"), show_name=True),
                row.get("estado", "-"),
                row.get("grupos", 0),
                row.get("pecas", 0),
                row.get("em_curso", 0),
                row.get("avarias", 0),
                f"{float(row.get('progress_pct', 0) or 0):.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, str(row.get("encomenda", "") or "").strip())
                    font = cell.font()
                    font.setBold(True)
                    cell.setFont(font)
                if col_index >= 3:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.orders_table.setItem(row_index, col_index, cell)
        self.orders_table.blockSignals(False)
        self.orders_table.setSortingEnabled(False)
        for row_index, row in enumerate(self.order_rows):
            _paint_table_row(self.orders_table, row_index, str(row.get("estado", "")))
        if self.orders_table.rowCount() == 0:
            self.open_order_btn.setEnabled(False)
            self._set_order_header({})
            return
        target_row = 0
        if selected_order:
            for row_index, row in enumerate(self.order_rows):
                if str(row.get("encomenda", "") or "").strip() == selected_order:
                    target_row = row_index
                    break
        self.orders_table.selectRow(target_row)

    def _refresh_orders_list_view(self) -> None:
        selected_order = str(self._current_order_row().get("encomenda", "") or "").strip()
        self._render_orders_table(selected_order)
        self._sync_order_focus()

    def _refresh_detail_badges(self) -> None:
        pieces = [piece for item in self.items for piece in list(item.get("pieces", []) or [])]
        active_count = 0
        running_count = 0
        delayed_count = 0
        for item in self.items:
            if float(item.get("desvio_min", 0) or 0) > 0:
                delayed_count += 1
        for piece in pieces:
            state = str(piece.get("estado", "") or "").strip().lower()
            if "concl" not in state:
                active_count += 1
            if "produc" in state or "curso" in state:
                running_count += 1
        _apply_state_chip(self.detail_active_chip, "Preparacao", f"Ativas {active_count}")
        _apply_state_chip(self.detail_late_chip, "Em pausa", f"Atrasadas {delayed_count}")
        _apply_state_chip(self.detail_running_chip, "Em producao", f"Em curso {running_count}")
        _apply_state_chip(self.detail_groups_chip, "Concluida", f"{len(self.items)} grupos")
        _apply_state_chip(self.detail_total_chip, "-", f"Todas {len(pieces)}")

    def _set_order_header(self, row: dict) -> None:
        numero = str(row.get("encomenda", "") or "").strip()
        estado = str(row.get("estado", "-") or "-").strip() or "-"
        if not numero:
            self.order_focus_label.setText("Sem encomenda selecionada")
            self.order_meta_label.setText("Seleciona uma encomenda para entrar no detalhe.")
            _apply_state_chip(self.order_state_chip, "-")
            if hasattr(self, "detail_active_chip"):
                _apply_state_chip(self.detail_active_chip, "-", "Ativas 0")
                _apply_state_chip(self.detail_late_chip, "-", "Atrasadas 0")
                _apply_state_chip(self.detail_running_chip, "-", "Em curso 0")
                _apply_state_chip(self.detail_groups_chip, "-", "0 grupos")
                _apply_state_chip(self.detail_total_chip, "-", "Todas 0")
            return
        full_order_title = f"{numero} | {_format_client_label(row.get('cliente', '-'), show_name=True)}"
        self.order_focus_label.setText(_elide_middle(full_order_title, 72))
        self.order_focus_label.setToolTip(full_order_title)
        self.order_meta_label.setText(f"{int(row.get('grupos', 0) or 0)} grupos | {int(row.get('pecas', 0) or 0)} pecas | Em curso {int(row.get('em_curso', 0) or 0)} | Avarias {int(row.get('avarias', 0) or 0)} | Progresso {float(row.get('progress_pct', 0) or 0):.1f}%")
        _apply_state_chip(self.order_state_chip, estado)

    def _sync_order_focus(self) -> None:
        row = self._current_order_row()
        self.open_order_btn.setEnabled(bool(row))
        self._set_order_header(row)

    def _apply_order_filter(self, numero: str, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        numero_txt = str(numero or "").strip()
        self.selected_order_number = numero_txt
        self.items = [item for item in self.all_items if str(item.get("encomenda", "") or "").strip() == numero_txt]
        self._refresh_detail_badges()
        target_group = previous_group if previous_group and previous_group[0] == numero_txt else None
        self._render_groups_table(target_group, previous_piece if target_group else "")
        order_row = next((row for row in self.order_rows_all if str(row.get("encomenda", "") or "").strip() == numero_txt), {})
        self._set_order_header(order_row)
        if not self.items:
            self._clear_piece_context()

    def _sync_group_focus(self) -> None:
        group = self._current_group()
        row = next((item for item in self.order_rows_all if str(item.get("encomenda", "") or "").strip() == str(self.selected_order_number or "").strip()), {})
        if not row:
            return
        summary = (
                f"{int(row.get('grupos', 0) or 0)} grupos | {int(row.get('pecas', 0) or 0)} pecas | "
                f"Em curso {int(row.get('em_curso', 0) or 0)} | Avarias {int(row.get('avarias', 0) or 0)} | "
                f"Progresso {float(row.get('progress_pct', 0) or 0):.1f}%"
            )
        if group:
            summary = (
                f"{summary} | Cliente {_format_client_label(group.get('cliente', '-'), show_name=True)} | "
                f"Grupo {group.get('material', '-')} {group.get('espessura', '-')} mm | "
                f"Estado {group.get('estado_espessura', group.get('estado', '-'))}"
            )
        self.order_meta_label.setText(summary)

    def _handle_group_selection(self) -> None:
        OperatorPage._handle_group_selection(self)
        self._sync_group_focus()

    def _open_selected_order(self) -> None:
        row = self._current_order_row()
        numero = str(row.get("encomenda", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Operador", "Seleciona uma encomenda.")
            return
        self._apply_order_filter(numero)
        self._show_order_detail()

    def _open_selected_order_from_item(self, item: QTableWidgetItem) -> None:
        row_item = self.orders_table.item(item.row(), 0) or item
        numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if not numero:
            return
        self._apply_order_filter(numero)
        self._show_order_detail()


class LegacyPlanningPage(PlanningPage):
    page_subtitle = "Tabela semanal ao centro, encomendas pendentes à esquerda e bloqueios por semana."

    def __init__(self, runtime_service, backend=None, parent=None) -> None:
        super().__init__(runtime_service, backend, parent)
        root = self.layout()
        root.setSpacing(6)
        sections = _take_layout_items(root)
        top_item = sections[0] if len(sections) > 0 else None
        cards_item = sections[1] if len(sections) > 1 else None
        body_item = sections[2] if len(sections) > 2 else None

        top_widget = top_item.widget() if top_item and top_item.widget() is not None else None
        if top_widget is not None:
            top_widget.setMaximumHeight(56)
            top_layout = top_widget.layout()
            if top_layout is not None:
                top_layout.setContentsMargins(12, 8, 12, 8)
                top_layout.setSpacing(4)

        cards_widget = cards_item.widget() if cards_item and cards_item.widget() is not None else None
        if cards_widget is not None:
            cards_widget.setMaximumHeight(78)
            cards_layout = cards_widget.layout()
            if cards_layout is not None:
                cards_layout.setHorizontalSpacing(8)
                cards_layout.setVerticalSpacing(0)
        for card in getattr(self, "cards", []):
            card.setMaximumHeight(76)
            card.title_label.setStyleSheet("font-size: 9px;")
            card.value_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
            card.subtitle_label.setStyleSheet("font-size: 9px;")

        self.grid.verticalHeader().setDefaultSectionSize(17)
        self.grid.verticalHeader().setMinimumSectionSize(17)
        self.grid.horizontalHeader().setMinimumSectionSize(88)
        self.grid.horizontalHeader().setFixedHeight(30)
        self.grid.setStyleSheet("font-size: 8.5px;")
        self.grid.setMinimumHeight(_table_visible_height(self.grid, 20, extra=8))
        self.grid.setMaximumHeight(16777215)
        self.backlog_table.verticalHeader().setDefaultSectionSize(18)
        self.backlog_table.setStyleSheet("font-size: 10px;")
        self.backlog_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        backlog_header = self.backlog_table.horizontalHeader()
        for col, width in ((0, 150), (1, 126), (2, 96), (3, 50), (4, 54)):
            backlog_header.setSectionResizeMode(col, QHeaderView.Interactive)
            backlog_header.resizeSection(col, width)
        backlog_header.setStretchLastSection(False)

        _adopt_layout_item(root, top_item)
        _adopt_layout_item(root, cards_item)
        _adopt_layout_item(root, body_item, 1)
        body_widget = body_item.widget() if body_item and body_item.widget() is not None else None
        if body_widget is not None and isinstance(body_widget, QSplitter):
            body_widget.setSizes([440, 1120])
        QTimer.singleShot(0, self._fit_planning_grid)


class QuotesPage(QWidget):
    page_title = "Orçamentos"
    page_subtitle = "Lista de orçamentos primeiro e detalhe apenas quando abres o registo."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
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
        list_layout.setSpacing(14)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QGridLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setHorizontalSpacing(8)
        filters_layout.setVerticalSpacing(4)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por numero, cliente ou encomenda")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Ativas", "Todos", "Em edicao", "Enviado", "Aprovado", "Rejeitado", "Convertido"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        self.year_combo = QComboBox()
        self.year_combo.currentTextChanged.connect(self.refresh)
        self.new_quote_btn = QPushButton("Novo orcamento")
        self.new_quote_btn.clicked.connect(self._new_quote)
        self.open_quote_btn = QPushButton("Abrir orcamento")
        self.open_quote_btn.clicked.connect(self._open_selected_quote)
        self.remove_quote_btn = QPushButton("Remover")
        self.remove_quote_btn.setProperty("variant", "danger")
        self.remove_quote_btn.clicked.connect(self._remove_quote)
        filters_layout.addWidget(QLabel("Pesquisa"), 0, 0)
        filters_layout.addWidget(QLabel("Estado"), 0, 1)
        filters_layout.addWidget(QLabel("Ano"), 0, 2)
        filters_layout.addWidget(QLabel("Acoes"), 0, 3)
        filters_layout.addWidget(self.filter_edit, 1, 0)
        filters_layout.addWidget(self.state_combo, 1, 1)
        filters_layout.addWidget(self.year_combo, 1, 2)
        action_host = QWidget()
        action_layout = QHBoxLayout(action_host)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(6)
        for button, width in ((self.new_quote_btn, 140), (self.open_quote_btn, 142), (self.remove_quote_btn, 114)):
            button.setProperty("compact", "true")
            button.setMinimumWidth(width)
            action_layout.addWidget(button)
        action_layout.addStretch(1)
        filters_layout.addWidget(action_host, 1, 3)
        filters_layout.setColumnStretch(0, 4)
        filters_layout.setColumnStretch(1, 2)
        filters_layout.setColumnStretch(2, 2)
        filters_layout.setColumnStretch(3, 5)
        filters.setMaximumHeight(86)
        list_layout.addWidget(filters)

        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(16, 14, 16, 14)
        table_title = QLabel("Orçamentos")
        table_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Numero", "Cliente", "Estado", "Encomenda", "Total", "Data"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setStyleSheet(
            f"QTableWidget {{ font-size: {LIST_TABLE_FONT_PX}px; }}"
            f" QHeaderView::section {{ font-size: {LIST_TABLE_FONT_PX}px; padding: 8px 10px; font-weight: 800; }}"
        )
        self.table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        _configure_table(self.table, stretch=(1,), contents=(2, 4, 5))
        _set_table_columns(
            self.table,
            [
                (0, "fixed", 220),
                (1, "stretch", 0),
                (2, "fixed", 138),
                (3, "fixed", 210),
                (4, "fixed", 145),
                (5, "fixed", 135),
            ],
        )
        self.table.itemSelectionChanged.connect(self._sync_list_buttons)
        self.table.itemDoubleClicked.connect(lambda *_args: self._open_selected_quote())
        table_layout.addWidget(table_title)
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
        detail_outer.addWidget(self.quote_detail_scroll)
        self.quote_detail_host = QWidget()
        self.quote_detail_scroll.setWidget(self.quote_detail_host)
        detail_layout = QVBoxLayout(self.quote_detail_host)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)

        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(14, 10, 14, 10)
        detail_actions_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a lista")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_list)
        save_btn = QPushButton("Guardar")
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
        reject_btn.setProperty("variant", "danger")
        reject_btn.clicked.connect(lambda: self._set_quote_state("Rejeitado"))
        convert_btn = QPushButton("Converter em encomenda")
        convert_btn.clicked.connect(self._convert_quote)
        preview_btn = QPushButton("Pre-visualizar")
        preview_btn.setProperty("variant", "secondary")
        preview_btn.clicked.connect(self._preview_quote)
        pdf_btn = QPushButton("Guardar PDF")
        pdf_btn.setProperty("variant", "secondary")
        pdf_btn.clicked.connect(self._save_quote_pdf)
        for button, width in (
            (back_btn, 126),
            (save_btn, 100),
            (edit_btn, 116),
            (sent_btn, 104),
            (approve_btn, 110),
            (reject_btn, 112),
            (convert_btn, 188),
            (preview_btn, 126),
            (pdf_btn, 120),
        ):
            button.setMinimumWidth(width)
        for button in (back_btn, save_btn, edit_btn, sent_btn, approve_btn, reject_btn, convert_btn, preview_btn, pdf_btn):
            detail_actions_layout.addWidget(button)
        detail_actions_layout.addStretch(1)
        detail_layout.addWidget(detail_actions)

        self.client_combo = QComboBox()
        self.client_combo.setEditable(True)
        self.client_combo.currentTextChanged.connect(self._fill_client_from_combo)
        self.executed_combo = QComboBox()
        self.executed_combo.setEditable(True)
        self.workcenter_combo = QComboBox()
        self.workcenter_combo.setEditable(False)
        self.client_name_edit = QLineEdit()
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
        self.iva_spin = QDoubleSpinBox()
        self.iva_spin.setRange(0.0, 100.0)
        self.iva_spin.setDecimals(2)
        self.iva_spin.setValue(23.0)
        self.iva_spin.valueChanged.connect(lambda _value: self._render_quote_lines())
        self.notes_edit = QTextEdit()
        self.notes_edit.setMinimumHeight(62)
        self.notes_edit.setMaximumHeight(82)
        self.notes_edit.setPlaceholderText("Notas tecnicas e comerciais para o PDF do orcamento.")
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
        self.state_chip = QLabel("-")
        _apply_state_chip(self.state_chip, "-")
        header.addLayout(title_block, 1)
        header.addWidget(self.state_chip, 0, Qt.AlignTop)
        meta_layout.addLayout(header)
        detail_layout.addWidget(self.quote_header_card)

        self.quote_client_card = CardFrame()
        self.quote_client_card.set_tone("default")
        client_layout = QVBoxLayout(self.quote_client_card)
        client_layout.setContentsMargins(12, 10, 12, 10)
        client_layout.setSpacing(6)
        client_title = QLabel("Dados do Cliente")
        client_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        client_layout.addWidget(client_title)

        self.quote_exec_card = CardFrame()
        self.quote_exec_card.set_tone("default")
        exec_layout = QVBoxLayout(self.quote_exec_card)
        exec_layout.setContentsMargins(12, 10, 12, 10)
        exec_layout.setSpacing(6)
        exec_title = QLabel("Dados do Orcamentista")
        exec_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        exec_layout.addWidget(exec_title)

        meta_left_host = QWidget()
        meta_left_form = QFormLayout(meta_left_host)
        meta_left_form.setContentsMargins(0, 0, 0, 0)
        meta_left_form.setHorizontalSpacing(10)
        meta_left_form.setVerticalSpacing(4)
        meta_left_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_left_form.addRow("Cliente", self.client_combo)
        meta_left_form.addRow("Nome", self.client_name_edit)
        meta_left_form.addRow("NIF", self.client_nif_edit)
        meta_left_form.addRow("Email", self.client_email_edit)
        meta_right_host = QWidget()
        meta_right_form = QFormLayout(meta_right_host)
        meta_right_form.setContentsMargins(0, 0, 0, 0)
        meta_right_form.setHorizontalSpacing(10)
        meta_right_form.setVerticalSpacing(4)
        meta_right_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_right_form.addRow("Orcamentista", self.executed_combo)
        meta_right_form.addRow("Posto trab.", self.workcenter_combo)
        meta_right_form.addRow("Empresa", self.client_company_edit)
        meta_right_form.addRow("Contacto", self.client_contact_edit)
        meta_right_form.addRow("Morada", self.client_address_edit)
        client_layout.addWidget(meta_left_host, 1)
        exec_layout.addWidget(meta_right_host, 1)
        meta_note_form = QFormLayout()
        meta_note_form.setContentsMargins(0, 0, 0, 0)
        meta_note_form.setHorizontalSpacing(10)
        meta_note_form.setVerticalSpacing(4)
        meta_note_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        meta_note_form.addRow("Nota cliente", self.note_cliente_edit)
        client_layout.addLayout(meta_note_form)

        self.quote_notes_card = CardFrame()
        self.quote_notes_card.set_tone("warning")
        notes_layout = QVBoxLayout(self.quote_notes_card)
        notes_layout.setContentsMargins(10, 8, 10, 8)
        notes_layout.setSpacing(8)
        notes_header = QHBoxLayout()
        notes_title = QLabel("Notas do orcamento (PDF)")
        notes_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        notes_hint = QLabel("PDF / transporte")
        notes_hint.setProperty("role", "muted")
        notes_hint.setStyleSheet("font-size: 10px;")
        notes_header.addWidget(notes_title)
        notes_header.addStretch(1)
        notes_header.addWidget(notes_hint)
        notes_layout.addLayout(notes_header)

        self.notes_tabs = QTabWidget()
        self.notes_tabs.setDocumentMode(True)
        self.notes_tabs.setStyleSheet("QTabBar::tab { font-size: 11px; min-height: 27px; }")

        transport_page = QWidget()
        transport_page_layout = QVBoxLayout(transport_page)
        transport_page_layout.setContentsMargins(0, 0, 0, 0)
        transport_page_layout.setSpacing(7)
        transport_intro = QLabel("Define o transporte do orçamento, calcula a sugestão e aplica diretamente ao PDF e ao total final.")
        transport_intro.setProperty("role", "muted")
        transport_intro.setWordWrap(True)
        transport_intro.setStyleSheet("font-size: 10.5px;")
        transport_page_layout.addWidget(transport_intro)

        transport_grid = QGridLayout()
        transport_grid.setContentsMargins(0, 0, 0, 0)
        transport_grid.setHorizontalSpacing(8)
        transport_grid.setVerticalSpacing(8)

        transport_form_card = CardFrame()
        transport_form_card.set_tone("default")
        transport_form_layout = QGridLayout(transport_form_card)
        transport_form_layout.setContentsMargins(10, 9, 10, 9)
        transport_form_layout.setHorizontalSpacing(12)
        transport_form_layout.setVerticalSpacing(6)

        left_transport_form = QFormLayout()
        left_transport_form.setContentsMargins(0, 0, 0, 0)
        left_transport_form.setHorizontalSpacing(8)
        left_transport_form.setVerticalSpacing(6)
        left_transport_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        left_transport_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        left_transport_form.addRow("Modo", self.transport_combo)
        left_transport_form.addRow("Transportadora", self.transport_carrier_combo)
        left_transport_form.addRow("Zona", self.transport_zone_combo)
        left_transport_form.addRow("Preço manual", self.transport_price_spin)
        left_transport_form.addRow("IVA %", self.iva_spin)

        right_transport_form = QFormLayout()
        right_transport_form.setContentsMargins(0, 0, 0, 0)
        right_transport_form.setHorizontalSpacing(8)
        right_transport_form.setVerticalSpacing(6)
        right_transport_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        right_transport_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        right_transport_form.addRow("Distância", self.transport_km_spin)
        right_transport_form.addRow("Preço / km", self.transport_rate_spin)
        right_transport_form.addRow("Gasóleo", self.transport_diesel_spin)
        right_transport_form.addRow("Consumo", self.transport_consumption_spin)
        right_transport_form.addRow("Fator de viagem", self.transport_trip_factor_spin)

        transport_form_layout.addLayout(left_transport_form, 0, 0)
        transport_form_layout.addLayout(right_transport_form, 0, 1)
        transport_form_layout.setColumnStretch(0, 1)
        transport_form_layout.setColumnStretch(1, 1)

        transport_actions_card = CardFrame()
        transport_actions_card.set_tone("default")
        transport_actions_layout = QGridLayout(transport_actions_card)
        transport_actions_layout.setContentsMargins(10, 8, 10, 8)
        transport_actions_layout.setHorizontalSpacing(12)
        transport_actions_layout.setVerticalSpacing(6)
        transport_actions_title = QLabel("Cálculo e aplicação")
        transport_actions_title.setStyleSheet("font-size: 11px; font-weight: 800; color: #0f172a;")
        transport_actions_hint = QLabel("Usa os parâmetros do transporte para sugerir um valor coerente com a orçamentação.")
        transport_actions_hint.setProperty("role", "muted")
        transport_actions_hint.setWordWrap(True)
        transport_actions_hint.setStyleSheet("font-size: 10.5px;")
        transport_actions_text = QVBoxLayout()
        transport_actions_text.setContentsMargins(0, 0, 0, 0)
        transport_actions_text.setSpacing(4)
        transport_actions_text.addWidget(transport_actions_title)
        transport_actions_text.addWidget(transport_actions_hint)
        fill_notes_btn = QPushButton("Preencher notas PDF")
        fill_notes_btn.setProperty("variant", "secondary")
        fill_notes_btn.setProperty("compact", "true")
        fill_notes_btn.setToolTip("Preenche automaticamente as notas PDF com base no contexto do orçamento.")
        fill_notes_btn.clicked.connect(self._fill_pdf_notes_from_context)
        apply_transport_btn = QPushButton("Aplicar ao orçamento")
        apply_transport_btn.setProperty("variant", "secondary")
        apply_transport_btn.setProperty("compact", "true")
        apply_transport_btn.setToolTip("Aplica o cálculo do transporte ao orçamento atual.")
        apply_transport_btn.clicked.connect(self._apply_transport_calc)
        self.transport_suggest_label.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        self.transport_suggest_label.setAlignment(Qt.AlignCenter)
        transport_action_buttons = QVBoxLayout()
        transport_action_buttons.setContentsMargins(0, 0, 0, 0)
        transport_action_buttons.setSpacing(6)
        transport_action_buttons.addWidget(fill_notes_btn)
        transport_action_buttons.addWidget(apply_transport_btn)
        transport_actions_layout.addLayout(transport_actions_text, 0, 0)
        transport_actions_layout.addWidget(self.transport_suggest_label, 0, 1)
        transport_actions_layout.addLayout(transport_action_buttons, 0, 2)
        transport_actions_layout.setColumnStretch(0, 5)
        transport_actions_layout.setColumnStretch(1, 3)
        transport_actions_layout.setColumnStretch(2, 3)

        transport_grid.addWidget(transport_form_card, 0, 0, 1, 3)
        transport_grid.addWidget(transport_actions_card, 1, 0, 1, 3)
        transport_grid.setColumnStretch(0, 1)
        transport_grid.setColumnStretch(1, 1)
        transport_grid.setColumnStretch(2, 1)
        transport_page_layout.addLayout(transport_grid, 1)

        operations_page = QWidget()
        operations_layout = QVBoxLayout(operations_page)
        operations_layout.setContentsMargins(0, 0, 0, 0)
        operations_layout.setSpacing(8)
        operations_intro = QLabel("Insere rapidamente notas técnicas e comerciais no PDF conforme o processo considerado.")
        operations_intro.setProperty("role", "muted")
        operations_intro.setWordWrap(True)
        operations_intro.setStyleSheet("font-size: 10.5px;")
        operations_layout.addWidget(operations_intro)
        op_card = CardFrame()
        op_card.set_tone("default")
        op_card_layout = QVBoxLayout(op_card)
        op_card_layout.setContentsMargins(12, 10, 12, 10)
        op_card_layout.setSpacing(8)
        op_title = QLabel("Inserções rápidas por operação")
        op_title.setStyleSheet("font-size: 11px; font-weight: 800; color: #0f172a;")
        op_card_layout.addWidget(op_title)
        op_buttons = QGridLayout()
        op_buttons.setHorizontalSpacing(8)
        op_buttons.setVerticalSpacing(8)
        for index, op_text in enumerate(("Corte Laser", "Quinagem", "Roscagem", "Furo Manual", "Soldadura")):
            op_btn = QPushButton(op_text)
            op_btn.setProperty("variant", "secondary")
            op_btn.clicked.connect(lambda _checked=False, text=op_text: self._append_pdf_note(f"- Foi considerado: {text}."))
            op_btn.setMinimumHeight(34)
            op_buttons.addWidget(op_btn, index // 3, index % 3)
        op_card_layout.addLayout(op_buttons)
        operations_layout.addWidget(op_card, 1)

        notes_text_page = QWidget()
        notes_text_page_layout = QVBoxLayout(notes_text_page)
        notes_text_page_layout.setContentsMargins(0, 0, 0, 0)
        notes_text_page_layout.setSpacing(8)
        notes_text_intro = QLabel("Texto livre que segue diretamente para o PDF do orçamento.")
        notes_text_intro.setProperty("role", "muted")
        notes_text_intro.setWordWrap(True)
        notes_text_intro.setStyleSheet("font-size: 10.5px;")
        notes_text_page_layout.addWidget(notes_text_intro)
        notes_text_card = CardFrame()
        notes_text_card.set_tone("default")
        notes_text_layout = QVBoxLayout(notes_text_card)
        notes_text_layout.setContentsMargins(12, 10, 12, 10)
        notes_text_layout.setSpacing(8)
        notes_text_title = QLabel("Texto técnico e comercial")
        notes_text_title.setStyleSheet("font-size: 11px; font-weight: 800; color: #0f172a;")
        notes_text_layout.addWidget(notes_text_title)
        self.notes_edit.setProperty("compact", "false")
        self.notes_edit.setMinimumHeight(180)
        self.notes_edit.setMaximumHeight(260)
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
            widget.setProperty("compact", "true")
            widget.setMaximumWidth(16777215)
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            widget.setMinimumWidth(0)
            widget.setMinimumHeight(24)
            _repolish(widget)
        self.notes_tabs.setMinimumHeight(358)
        self.quote_summary_card = CardFrame()
        self.quote_summary_card.set_tone("info")
        self.quote_summary_card.setMaximumWidth(265)
        self.quote_summary_card.setMinimumWidth(236)
        summary_layout = QVBoxLayout(self.quote_summary_card)
        summary_layout.setContentsMargins(12, 10, 12, 10)
        summary_layout.setSpacing(8)
        summary_title = QLabel("Resumo financeiro")
        summary_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        summary_hint = QLabel("Preco final do orcamento com transporte e IVA.")
        summary_hint.setProperty("role", "muted")
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(summary_hint)
        summary_grid = QGridLayout()
        summary_grid.setHorizontalSpacing(8)
        summary_grid.setVerticalSpacing(7)
        self.lines_subtotal_label = QLabel("0,00 EUR")
        self.transport_total_label = QLabel("0,00 EUR")
        self.discount_value_label = QLabel("0,00 EUR")
        self.subtotal_without_iva_label = QLabel("0,00 EUR")
        self.iva_total_label = QLabel("0,00 EUR")
        self.total_label = QLabel("0,00 EUR")
        for widget in (
            self.lines_subtotal_label,
            self.transport_total_label,
            self.discount_value_label,
            self.subtotal_without_iva_label,
            self.iva_total_label,
        ):
            widget.setProperty("role", "field_value")
        self.total_label.setProperty("role", "field_value_strong")
        for row_index, (label_text, widget) in enumerate(
            (
                ("Linhas", self.lines_subtotal_label),
                ("Transporte", self.transport_total_label),
                ("Desconto", self.discount_value_label),
                ("Subtotal s/ IVA", self.subtotal_without_iva_label),
                ("IVA", self.iva_total_label),
                ("Total c/ IVA", self.total_label),
            )
        ):
            label = QLabel(label_text)
            label.setProperty("role", "field_label")
            summary_grid.addWidget(label, row_index, 0)
            summary_grid.addWidget(widget, row_index, 1)
        summary_layout.addLayout(summary_grid)
        discount_row = QHBoxLayout()
        discount_row.setContentsMargins(0, 2, 0, 0)
        discount_caption = QLabel("Desc. global")
        discount_caption.setProperty("role", "field_label")
        discount_row.addWidget(discount_caption)
        discount_row.addStretch(1)
        discount_row.addWidget(self.discount_spin)
        summary_layout.addLayout(discount_row)
        summary_layout.addStretch(1)
        top_split = QSplitter(Qt.Horizontal)
        top_split.setChildrenCollapsible(False)
        top_split.addWidget(self.quote_client_card)
        top_split.addWidget(self.quote_exec_card)
        top_split.addWidget(self.quote_summary_card)
        top_split.setSizes([500, 500, 250])
        detail_layout.addWidget(top_split)

        lines_card = CardFrame()
        lines_card.set_tone("default")
        lines_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.quote_lines_card = lines_card
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(16, 14, 16, 14)
        lines_layout.setSpacing(10)
        line_actions = QVBoxLayout()
        line_actions.setSpacing(6)
        line_title_row = QHBoxLayout()
        line_title_row.setSpacing(8)
        lines_title = QLabel("Referencias do orcamento")
        lines_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        add_line_btn = QPushButton("Adicionar linha")
        add_line_btn.clicked.connect(self._add_line)
        laser_line_btn = QPushButton("Peca Unit. DXF/DWG")
        laser_line_btn.clicked.connect(self._add_laser_line)
        laser_batch_btn = QPushButton("Lote DXF/DWG")
        laser_batch_btn.setProperty("variant", "secondary")
        laser_batch_btn.clicked.connect(self._add_laser_batch_lines)
        laser_nesting_btn = QPushButton("Plano chapa")
        laser_nesting_btn.setProperty("variant", "secondary")
        laser_nesting_btn.clicked.connect(self._open_laser_nesting)
        add_model_btn = QPushButton("Adicionar conjunto")
        add_model_btn.setProperty("variant", "secondary")
        add_model_btn.clicked.connect(self._add_assembly_model)
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
        check_weight_btn = QPushButton("Check Weight")
        check_weight_btn.setProperty("variant", "secondary")
        check_weight_btn.clicked.connect(self._check_selected_line_weight)
        remove_line_btn = QPushButton("Remover linha")
        remove_line_btn.setProperty("variant", "danger")
        remove_line_btn.clicked.connect(self._remove_line)
        open_draw_btn = QPushButton("Ver desenho")
        open_draw_btn.setProperty("variant", "secondary")
        open_draw_btn.clicked.connect(self._open_line_drawing)
        laser_nesting_btn.setToolTip("Abrir o plano de chapa / nesting das linhas laser do orcamento.")
        line_title_row.addWidget(lines_title)
        line_title_row.addStretch(1)
        line_actions.addLayout(line_title_row)
        self.nesting_bridge_label = QLabel(
            "Preco unitario da tabela = orcamentacao por peca. Plano chapa = validacao global de materia, stock/retalho e pecas realmente programadas."
        )
        self.nesting_bridge_label.setProperty("role", "muted")
        self.nesting_bridge_label.setWordWrap(True)
        self.nesting_bridge_label.setStyleSheet("font-size: 10.5px; line-height: 1.15;")
        line_actions.addWidget(self.nesting_bridge_label)
        line_buttons = QGridLayout()
        line_buttons.setHorizontalSpacing(8)
        line_buttons.setVerticalSpacing(6)
        action_rows = (
            (add_line_btn, laser_line_btn, laser_batch_btn, laser_nesting_btn),
            (laser_cfg_btn, operation_cfg_btn, manage_models_btn, add_model_btn),
            (edit_line_btn, open_draw_btn, remove_line_btn, check_weight_btn),
        )
        for row_index, button_row in enumerate(action_rows):
            for col_index, button in enumerate(button_row):
                if not isinstance(button, QPushButton):
                    continue
                button.setProperty("compact", "true")
                button.setMinimumWidth(0)
                button.setMinimumHeight(28)
                button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                line_buttons.addWidget(button, row_index, col_index)
        for col_index in range(4):
            line_buttons.setColumnStretch(col_index, 1)
        line_actions.addLayout(line_buttons)
        for button in (add_line_btn, laser_line_btn, laser_batch_btn, laser_nesting_btn, add_model_btn, manage_models_btn, laser_cfg_btn, operation_cfg_btn, edit_line_btn, remove_line_btn, open_draw_btn, check_weight_btn):
            button.setProperty("compact", "true")
        self.lines_table = QTableWidget(0, 12)
        self.lines_table.setHorizontalHeaderLabels(["Tipo", "Ref./Cod.", "Ref. Ext.", "Descricao", "Material/Produto", "Esp./Unid", "Operacao", "Tempo", "Qtd", "Preco", "Total", "Conjunto"])
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.lines_table.setAlternatingRowColors(True)
        self.lines_table.setStyleSheet(
            "QTableWidget { font-size: 11px; }"
            " QHeaderView::section { font-size: 11px; padding: 5px 6px; font-weight: 700; }"
            " QScrollBar:vertical { background: #e2e8f0; width: 14px; margin: 0; border-left: 1px solid #cbd5e1; }"
            " QScrollBar::handle:vertical { background: #1e3a8a; min-height: 28px; border-radius: 6px; margin: 2px; }"
            " QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }"
            " QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }"
        )
        self.lines_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.lines_table.verticalHeader().setDefaultSectionSize(28)
        self.lines_table.setMinimumHeight(380)
        self.lines_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        _set_table_columns(
            self.lines_table,
            [
                (0, "fixed", 108),
                (1, "fixed", 118),
                (2, "fixed", 148),
                (3, "stretch", 0),
                (4, "stretch", 0),
                (5, "fixed", 82),
                (6, "fixed", 118),
                (7, "fixed", 76),
                (8, "fixed", 58),
                (9, "fixed", 84),
                (10, "fixed", 92),
                (11, "fixed", 96),
            ],
        )
        lines_layout.addLayout(line_actions)
        lines_layout.addWidget(self.lines_table)
        self.quote_client_card.setMaximumHeight(212)
        self.quote_exec_card.setMaximumHeight(212)
        self.quote_notes_card.setMaximumWidth(660)
        self.quote_notes_card.setMinimumWidth(560)
        lines_card.setMinimumHeight(430)
        lines_card.setMaximumHeight(720)
        content_split = QSplitter(Qt.Horizontal)
        content_split.setChildrenCollapsible(False)
        content_split.addWidget(self.quote_notes_card)
        content_split.addWidget(lines_card)
        content_split.setSizes([620, 1270])
        detail_layout.addWidget(content_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        self._quote_save_feedback_timer = QTimer(self)
        self._quote_save_feedback_timer.setSingleShot(True)
        self._quote_save_feedback_timer.timeout.connect(self._reset_quote_save_button)
        self._clear_quote_detail()
        self._show_list()

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
            [[r.get("numero", "-"), r.get("cliente", "-"), r.get("estado", "-"), r.get("numero_encomenda", "-"), _fmt_eur(r.get("total", 0)), r.get("data", "-")] for r in self.rows],
            align_center_from=3,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.table, row_index, str(row.get("estado", "")))
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

    def _show_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_buttons()

    def _show_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _selected_quote_row(self) -> dict:
        current = self.table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def _selected_line_index(self) -> int:
        current = self.lines_table.currentItem()
        if current is None or current.row() >= len(self.line_rows):
            return -1
        return current.row()

    def _selected_line_indexes(self) -> list[int]:
        indexes = sorted(
            {
                item.row()
                for item in self.lines_table.selectedItems()
                if item is not None and 0 <= item.row() < len(self.line_rows)
            }
        )
        if indexes:
            return indexes
        single = self._selected_line_index()
        return [single] if single >= 0 else []

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
        current_workcenter = self.workcenter_combo.currentText().strip()
        self.workcenter_combo.blockSignals(True)
        self.workcenter_combo.clear()
        for value in list(self.backend.quote_workcenter_options() or []):
            self.workcenter_combo.addItem(str(value))
        if current_workcenter:
            self.workcenter_combo.setCurrentText(current_workcenter)
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

    def _quote_line_type(self, row: dict) -> str:
        return str(self.backend.desktop_main.normalize_orc_line_type((row or {}).get("tipo_item")) or self.backend.desktop_main.ORC_LINE_TYPE_PIECE)

    def _quote_line_type_label(self, row: dict) -> str:
        return str(self.backend.desktop_main.orc_line_type_label(self._quote_line_type(row)) or "-")

    def _quote_line_primary_ref(self, row: dict) -> str:
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        return str((row or {}).get("ref_interna", "") or "").strip() or "-"

    def _quote_line_material_display(self, row: dict) -> str:
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE:
            return "-"
        return str((row or {}).get("material", "") or "").strip() or "-"

    def _quote_line_unit_display(self, row: dict) -> str:
        line_type = self._quote_line_type(row)
        if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE:
            return str((row or {}).get("espessura", "") or "").strip() or "-"
        return str((row or {}).get("produto_unid", "") or "").strip() or "-"

    def _set_quote_header(self, numero: str, estado: str, encomenda: str = "") -> None:
        self.current_number = str(numero or "").strip()
        self.number_label.setText(self.current_number or "Novo orcamento")
        self.link_order_label.setText(f"Encomenda gerada: {encomenda}" if encomenda else "Sem encomenda gerada")
        _apply_state_chip(self.state_chip, estado)
        _set_panel_tone(self.quote_header_card, _state_tone(estado))
        _set_panel_tone(self.quote_summary_card, _state_tone(estado))

    def _clear_quote_detail(self) -> None:
        self.line_rows = []
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
        self.transport_km_spin.setValue(0.0)
        self.transport_rate_spin.setValue(0.65)
        self.transport_diesel_spin.setValue(1.65)
        self.transport_consumption_spin.setValue(8.5)
        self.transport_trip_factor_spin.setValue(2.0)
        self.executed_combo.setCurrentText("")
        self.workcenter_combo.setCurrentText("")
        self.notes_edit.clear()
        self.iva_spin.setValue(23.0)
        self._recalc_transport_calc()
        self._refresh_nesting_bridge()
        self._render_quote_lines()

    def _refresh_nesting_bridge(self) -> None:
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        if not bridge:
            self.nesting_bridge_label.setText(
                "Preco unitario da tabela = orcamentacao por peca. Plano chapa = validacao global de materia, stock/retalho e pecas realmente programadas."
            )
            self.nesting_bridge_label.setToolTip(
                "O preco unitario das linhas continua comercial e unitario. O plano de chapa serve para validar consumo real de materia, stock, compra e pecas efetivamente programadas."
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
            f"Plano chapa atual: {placed}/{requested} programadas | {sheets} chapa(s) | {method} | "
            f"perfil {profile_name} | materia real {material_cost} | compra {purchase_cost} | "
            f"comercial atual {quoted_total} | comercial ajustado ao rateio {rateio_total} | delta {rateio_delta}."
        )
        self.nesting_bridge_label.setToolTip(
            "Tabela do orcamento: preco comercial unitario por referencia.\n"
            "Plano chapa: custo global real de materia e validacao do lote completo.\n"
            "O rateio usa a area liquida colocada para repartir a materia real e a compra necessaria por referencia.\n"
            "Usa a tabela para vender/orcamentar por referencia e o plano de chapa para validar consumo real, compra e cobertura comercial."
        )

    def _render_quote_lines(self) -> None:
        subtotal = 0.0
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        bridge_rows = {
            str(row.get("ref_externa", "") or "").strip(): dict(row or {})
            for row in list(bridge.get("part_rows", []) or [])
            if str(row.get("ref_externa", "") or "").strip()
        }
        normalized_rows = []
        for row in self.line_rows:
            qtd = float(row.get("qtd", 0) or 0)
            preco_unit = float(row.get("preco_unit", 0) or 0)
            total = round(qtd * preco_unit, 2)
            row["total"] = total
            normalized_rows.append(row)
        self.line_rows = normalized_rows
        _fill_table(
            self.lines_table,
            [
                [
                    self._quote_line_type_label(row),
                    self._quote_line_primary_ref(row),
                    row.get("ref_externa", "-"),
                    row.get("descricao", "-"),
                    self._quote_line_material_display(row),
                    self._quote_line_unit_display(row),
                    row.get("operacao", "-"),
                    f"{float(row.get('tempo_peca_min', 0) or 0):.2f} min",
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    _fmt_eur(float(row.get("preco_unit", 0) or 0)),
                    _fmt_eur(float(row.get("total", 0) or 0)),
                    row.get("conjunto_nome", "-") or "-",
                ]
                for row in self.line_rows
            ],
            align_center_from=5,
        )
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
        for row_index, row in enumerate(self.line_rows):
            subtotal += float(row.get("total", 0) or 0)
            _paint_table_row(self.lines_table, row_index, "Preparacao")
            for col_index in (9, 10):
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
            ]
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
                    tooltip_lines.append(f"Plano chapa: {programmed}/{requested} peca(s) programada(s) no ultimo cenario.")
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
                current_total = float(row.get("total", 0) or 0.0)
                total_item = self.lines_table.item(row_index, 10)
                if total_item is not None and adjusted_total > 0:
                    if current_total + 0.009 < adjusted_total:
                        total_item.setBackground(QBrush(QColor("#fee2e2")))
                        total_item.setForeground(QBrush(QColor("#b42318")))
                    elif current_total > adjusted_total + 0.009:
                        total_item.setBackground(QBrush(QColor("#dcfce7")))
                        total_item.setForeground(QBrush(QColor("#166534")))
                    else:
                        total_item.setBackground(QBrush(QColor("#fef3c7")))
                        total_item.setForeground(QBrush(QColor("#92400e")))
        discount_pct = float(self.discount_spin.value() or 0)
        subtotal_bruto = subtotal + transport
        discount_value = round(subtotal_bruto * (discount_pct / 100.0), 2)
        subtotal_without_iva = max(0.0, subtotal_bruto - discount_value)
        iva_amount = subtotal_without_iva * (float(self.iva_spin.value()) / 100.0)
        total = subtotal_without_iva + iva_amount
        self.lines_subtotal_label.setText(_fmt_eur(subtotal))
        self.transport_total_label.setText(_fmt_eur(transport))
        self.discount_value_label.setText(_fmt_eur(discount_value))
        self.subtotal_without_iva_label.setText(_fmt_eur(subtotal_without_iva))
        self.iva_total_label.setText(_fmt_eur(iva_amount))
        self.total_label.setText(_fmt_eur(total))
        visible_rows = min(max(11, len(self.line_rows) + 2), 14)
        target_height = _table_visible_height(self.lines_table, visible_rows, extra=18)
        self.lines_table.setMinimumHeight(target_height)
        self.lines_table.setMaximumHeight(target_height)
        if hasattr(self, "quote_lines_card"):
            actions_height = 208
            self.quote_lines_card.setMinimumHeight(target_height + actions_height)

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
        self.executed_combo.setCurrentText(str(detail.get("executado_por", "") or "").strip())
        self.workcenter_combo.setCurrentText(str(detail.get("posto_trabalho", "") or "").strip())
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
        self.notes_edit.setPlainText(str(detail.get("notas_pdf", "") or "").strip())
        self.iva_spin.setValue(float(detail.get("iva_perc", 23) or 23))
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

    def _append_pdf_note(self, line: str) -> None:
        note = str(line or "").strip()
        if not note:
            return
        current = [item.strip() for item in self.notes_edit.toPlainText().splitlines() if item.strip()]
        if note not in current:
            current.append(note)
            self.notes_edit.setPlainText("\n".join(current))

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
            "posto_trabalho": self.workcenter_combo.currentText().strip(),
            "nota_transporte": self.transport_combo.currentText().strip(),
            "transportadora_nome": self.transport_carrier_combo.currentText().strip(),
            "zona_transporte": self.transport_zone_combo.currentText().strip(),
            "preco_transporte": self.transport_price_spin.value(),
            "desconto_perc": self.discount_spin.value(),
            "notas_pdf": self.notes_edit.toPlainText().strip(),
            "nota_cliente": self.note_cliente_edit.text().strip(),
            "iva_perc": self.iva_spin.value(),
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

    def _set_quote_save_button_state(self, state: str = "idle") -> None:
        button = getattr(self, "quote_save_btn", None)
        if not isinstance(button, QPushButton):
            return
        current_state = str(state or "idle").strip().lower()
        if current_state == "saving":
            button.setText("A guardar...")
            button.setEnabled(False)
            button.setProperty("variant", "secondary")
        elif current_state == "saved":
            button.setText("Guardado")
            button.setEnabled(True)
            button.setProperty("variant", "success")
        elif current_state == "error":
            button.setText("Falhou ao guardar")
            button.setEnabled(True)
            button.setProperty("variant", "danger")
        else:
            button.setText("Guardar")
            button.setEnabled(True)
            button.setProperty("variant", "")
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
                "style=\"max-width:220px; max-height:72px; display:block;\" />"
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
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#0f172a;\">30 dias, salvo rutura de stock.</td>"
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
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    powershell_script,
                ],
                check=True,
                env=env,
                timeout=30,
            )
        except Exception:
            mailto = (
                f"mailto:{quote(recipient)}"
                f"?subject={quote(subject)}"
                f"&body={quote(body_plain)}"
            )
            os.startfile(mailto)
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
        dialog.setMinimumWidth(820)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        client_code = self._client_code_from_text(self.client_combo.currentText())
        references = self.backend.order_reference_rows("", client_code)
        refs_by_ext = {str(row.get("ref_externa", "") or "").strip(): row for row in references if str(row.get("ref_externa", "") or "").strip()}
        refs_by_int = {str(row.get("ref_interna", "") or "").strip(): row for row in references if str(row.get("ref_interna", "") or "").strip()}
        presets = self.presets or self.backend.order_presets()
        product_rows = {str(row.get("codigo", "") or "").strip(): row for row in self.backend.ne_product_options("")}
        initial_type = str(
            self.backend.desktop_main.normalize_orc_line_type(
                initial.get("tipo_item", self.backend.desktop_main.ORC_LINE_TYPE_PIECE)
            )
        )
        ref_history = QComboBox()
        ref_history.setEditable(True)
        ref_history.addItem("")
        for row in references:
            ref_history.addItem(
                f"{row.get('ref_interna', '')} | {row.get('ref_externa', '')} | {row.get('descricao', '')}",
                row,
            )
        type_combo = QComboBox()
        type_combo.addItem("Peca fabricada", self.backend.desktop_main.ORC_LINE_TYPE_PIECE)
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
        product_unid_edit = QLineEdit(str(initial.get("produto_unid", "") or "").strip())
        product_unid_edit.setReadOnly(True)
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

        def current_line_type() -> str:
            return str(type_combo.currentData() or self.backend.desktop_main.ORC_LINE_TYPE_PIECE)

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
            if float(price_spin.value() or 0) <= 0:
                price_spin.setValue(float(row.get("preco", 0) or 0))
            if not ref_ext_edit.text().strip():
                ref_ext_edit.setText(code)

        def load_ref() -> None:
            selected = ref_history.currentData()
            if isinstance(selected, dict):
                apply_reference(selected)
                return
            key = ref_ext_edit.text().strip() or ref_history.currentText().strip()
            apply_reference(refs_by_ext.get(key) or refs_by_int.get(key))

        def browse_refs() -> None:
            payload = _reference_catalog_dialog(self, references, "Historico de referencias")
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
        btn_history = QPushButton("Historico refs")
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

        form.addRow("Tipo", type_combo)
        form.addRow("Historico", ref_history)
        form.addRow("Produto", product_combo)
        form.addRow("Ref. interna", ref_int_edit)
        form.addRow("Ref. externa", ref_ext_edit)
        form.addRow("Descricao", desc_edit)
        form.addRow("Codigo/unid", product_unid_edit)
        form.addRow("Material", material_combo)
        form.addRow("Espessura", esp_combo)
        form.addRow("Postos", operation_selector)
        form.addRow("Custeio op.", operation_cost_label)
        form.addRow("", operation_buttons)
        form.addRow("Tempo peca (min)", tempo_spin)
        form.addRow("Quantidade", qtd_spin)
        form.addRow("Preco unit.", price_spin)
        form.addRow("Desenho", drawing_edit)
        form.addRow("", ref_buttons)
        layout.addLayout(form)

        def sync_mode() -> None:
            line_type = current_line_type()
            is_piece = line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE
            is_product = line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT
            ref_history.setEnabled(is_piece)
            ref_int_edit.setEnabled(is_piece and not template_mode)
            material_combo.setEnabled(is_piece)
            esp_combo.setEnabled(is_piece)
            drawing_edit.setEnabled(is_piece)
            btn_generate.setEnabled(is_piece and not template_mode)
            btn_history.setEnabled(is_piece)
            btn_load.setEnabled(is_piece)
            btn_drawing.setEnabled(is_piece)
            btn_ops_detail.setEnabled(is_piece)
            btn_ops_profiles.setEnabled(is_piece)
            product_combo.setEnabled(is_product)
            product_unid_edit.setEnabled(is_product or line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE)
            if is_product:
                base_state["laser_base_enabled"] = False
                base_state["laser_base_time"] = 0.0
                base_state["laser_base_price"] = 0.0
                sync_product_fields()
                if not operation_edit.text().strip():
                    apply_operations("Montagem")
            elif line_type == self.backend.desktop_main.ORC_LINE_TYPE_SERVICE:
                base_state["laser_base_enabled"] = False
                base_state["laser_base_time"] = 0.0
                base_state["laser_base_price"] = 0.0
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
            sync_operation_meta_selection()
            sync_line_totals_from_meta()
            refresh_operation_cost_hint()

        type_combo.currentTextChanged.connect(lambda _value: sync_mode())
        sync_mode()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        sync_operation_meta_selection()
        line_type = current_line_type()
        final_time_unit, final_price_unit = current_composed_totals()
        return {
            "tipo_item": line_type,
            "ref_interna": "" if (template_mode or line_type != self.backend.desktop_main.ORC_LINE_TYPE_PIECE) else ref_int_edit.text().strip(),
            "ref_externa": ref_ext_edit.text().strip(),
            "descricao": desc_edit.text().strip(),
            "material": material_combo.currentText().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "material_family": str(initial.get("material_family", "") or (material_combo.currentText().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "")).strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "material_subtype": str(initial.get("material_subtype", "") or "").strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
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
            "operacao": " + ".join(full_route_operation_names()),
            "tempo_peca_min": final_time_unit,
            "qtd": qtd_spin.value(),
            "qtd_base": float(qtd_spin.value() if template_mode else initial.get("qtd_base", qtd_spin.value())),
            "preco_unit": final_price_unit,
            "desenho": drawing_edit.text().strip() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "laser_base_active": bool(line_has_laser_base()),
            "laser_base_tempo_unit": current_laser_base_totals()[0] if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else 0.0,
            "laser_base_preco_unit": current_laser_base_totals()[1] if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PIECE else 0.0,
            "produto_codigo": current_product_code() if line_type == self.backend.desktop_main.ORC_LINE_TYPE_PRODUCT else "",
            "produto_unid": product_unid_edit.text().strip() if line_type != self.backend.desktop_main.ORC_LINE_TYPE_PIECE else "",
            "conjunto_codigo": str(initial.get("conjunto_codigo", "") or "").strip(),
            "conjunto_nome": str(initial.get("conjunto_nome", "") or "").strip(),
            "grupo_uuid": str(initial.get("grupo_uuid", "") or "").strip(),
            "operacoes_lista": list(operation_meta.get("operacoes_lista", []) or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(operation_meta.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
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

        items: list[dict] = [dict(row or {}) for row in list(initial.get("itens", []) or [])]
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
                        str(self.backend.desktop_main.orc_line_type_label(item.get("tipo_item")) or "-"),
                        str(item.get("descricao", "") or "").strip() or "-",
                        str(item.get("produto_codigo", "") or item.get("ref_externa", "") or "").strip() or "-",
                        str(item.get("material", "") or "").strip() or "-",
                        str(item.get("espessura", "") or item.get("produto_unid", "") or "").strip() or "-",
                        f"{float(item.get('qtd', 0) or 0):.2f}",
                        _fmt_eur(float(item.get("preco_unit", 0) or 0)),
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
            payload = self._line_dialog({"qtd": 1}, template_mode=True)
            if payload is None:
                return
            items.append(payload)
            render_items()

        def edit_item() -> None:
            index = selected_item_index()
            if index < 0:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um item.")
                return
            payload = self._line_dialog(items[index], template_mode=True)
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
            "itens": items,
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

    def _add_assembly_model(self) -> None:
        rows = list(self.backend.assembly_model_rows() or [])
        if not rows:
            QMessageBox.information(self, "Conjuntos", "Ainda nao existem modelos. Cria primeiro um modelo de conjunto.")
            self._manage_assembly_models()
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Adicionar conjunto")
        layout = QVBoxLayout(dialog)
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
            self.line_rows.extend(self.backend.assembly_model_expand(str(combo.currentData() or "").strip(), qty_spin.value()))
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return
        self._render_quote_lines()

    def _configure_laser_profiles(self) -> None:
        dialog = LaserSettingsDialog(self.backend, self)
        dialog.exec()

    def _configure_operation_profiles(self) -> None:
        _open_operation_cost_profiles_dialog(self, self.backend)

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
        selected_indexes = sorted(
            {
                item.row()
                for item in self.lines_table.selectedItems()
                if item is not None and 0 <= item.row() < len(self.line_rows)
            }
        )
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
        bridge = dialog.quote_bridge_payload()
        if bridge:
            self.nesting_bridge_data = bridge
            self._refresh_nesting_bridge()
            self._render_quote_lines()

    def _add_line(self) -> None:
        payload = self._line_dialog({"qtd": 1})
        if payload is None:
            return
        self.line_rows.append(payload)
        self._render_quote_lines()

    def _edit_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        payload = self._line_dialog(self.line_rows[index])
        if payload is None:
            return
        self.line_rows[index] = payload
        self._render_quote_lines()
        self.lines_table.selectRow(index)

    def _remove_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        del self.line_rows[index]
        self._render_quote_lines()

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
            self.backend.orc_save(self._quote_payload())
            path = self.backend.orc_open_pdf(self.current_number)
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
            self.backend.orc_save(self._quote_payload())
            self.backend.orc_render_pdf(self.current_number, path)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF guardado em:\n{path}")

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





