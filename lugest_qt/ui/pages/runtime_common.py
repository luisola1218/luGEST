from __future__ import annotations

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHeaderView,
    QLabel,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)


def fill_table(table: QTableWidget, rows: list[list[str]], align_center_from: int = 0) -> None:
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


def state_visual(state: str) -> dict[str, str]:
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


def state_palette(state: str) -> tuple[str, str]:
    visual = state_visual(state)
    return visual["bg"], visual["fg"]


def state_tone(state: str) -> str:
    return state_visual(state)["tone"]


def repolish(widget: QWidget) -> None:
    style = widget.style()
    if style is not None:
        style.unpolish(widget)
        style.polish(widget)
    widget.update()


def set_panel_tone(frame: QWidget, tone: str = "default") -> None:
    frame.setProperty("tone", str(tone or "default"))
    repolish(frame)


def apply_state_chip(label: QLabel, state: str, text: str | None = None) -> None:
    visual = state_visual(state)
    label.setProperty("role", "state_chip")
    label.setText(str(text if text is not None else state or "-") or "-")
    label.setStyleSheet(
        "padding: 6px 12px; border-radius: 999px; font-weight: 800;"
        f" background: {visual['bg']}; color: {visual['fg']}; border: 1px solid {visual['border']};"
    )


def paint_table_row(table: QTableWidget, row_index: int, state: str) -> None:
    bg_hex, fg_hex = state_palette(state)
    bg = QBrush(QColor(bg_hex))
    fg = QBrush(QColor(fg_hex))
    for col_index in range(table.columnCount()):
        item = table.item(row_index, col_index)
        if item is None:
            continue
        item.setBackground(bg)
        item.setForeground(fg)


def configure_table(
    table: QTableWidget,
    *,
    stretch: tuple[int, ...] = (),
    contents: tuple[int, ...] = (),
    center_from: int | None = None,
) -> None:
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


def set_table_columns(table: QTableWidget, specs: list[tuple[int, str, int]]) -> None:
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


def cap_width(widget: QWidget, width: int) -> None:
    widget.setMaximumWidth(width)


def coerce_editor_qdate(raw: str | None = None, *, fallback_today: bool = True) -> QDate:
    text = str(raw or "").strip()[:10]
    parsed = QDate.fromString(text, "yyyy-MM-dd") if text else QDate()
    if parsed.isValid() and parsed.year() > 2000:
        return parsed
    return QDate.currentDate() if fallback_today else QDate(2000, 1, 1)


def elide_middle(text: str, max_chars: int = 48) -> str:
    raw = str(text or "").strip()
    if len(raw) <= max_chars:
        return raw
    keep = max(8, (max_chars - 3) // 2)
    return f"{raw[:keep]}...{raw[-keep:]}"


def fmt_eur(value: float) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


def selected_row_index(table: QTableWidget) -> int:
    selection_model = table.selectionModel()
    if selection_model is not None:
        selected_rows = selection_model.selectedRows()
        if selected_rows:
            return selected_rows[0].row()
    current_row = table.currentRow()
    return current_row if current_row >= 0 else -1


def table_visible_height(table: QTableWidget, rows: int, extra: int = 10) -> int:
    row_height = table.verticalHeader().defaultSectionSize() or 24
    header_height = table.horizontalHeader().height() or 26
    frame = table.frameWidth() * 2
    return header_height + frame + (max(0, int(rows)) * row_height) + extra


def apply_progress_style(bar: QProgressBar, *, compact: bool = False) -> None:
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
