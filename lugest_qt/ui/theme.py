from __future__ import annotations

import unicodedata

from PySide6.QtCore import QEvent, QObject, QSize
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QAbstractItemView, QApplication, QPushButton, QStyle, QTableView, QWidget


_DESTRUCTIVE_WORDS = ("apagar", "eliminar", "remover", "excluir", "delete")
_DESTRUCTIVE_CLEAR_TARGETS = ("lote", "guia", "semana", "visiveis", "dados", "ficheiros", "itens", "decisao")
_EXIT_WORDS = ("sair", "exit", "quit", "close", "fechar", "terminar", "encerrar")
_NEUTRAL_WORDS = ("cancelar", "voltar")
_DESTRUCTIVE_BUTTON_STYLE = """
QPushButton {
    background: #d92d20;
    color: #ffffff;
    border: 1px solid #8f1d14;
}
QPushButton:hover { background: #b42318; }
QPushButton:pressed { background: #8f1d14; }
QPushButton:disabled {
    background: #ead4d1;
    color: #9c7772;
    border-color: #d8bbb7;
}
"""
_UNIFORM_TABLE_SELECTION = """
QTableView::item:selected {
    background: #eaf7da;
    color: #26331d;
    border: 0;
    border-left: 3px solid #7ed321;
    border-bottom: 1px solid #d8e6c9;
}
"""


def _normalized_caption(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(text or ""))
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    return " ".join(normalized.casefold().replace("&", " ").split())


def _refresh_widget_style(widget: QWidget) -> None:
    style = widget.style()
    if style is not None:
        style.unpolish(widget)
        style.polish(widget)
    widget.update()


def polish_button_semantics(button: QPushButton) -> None:
    """Apply the shared visual language to buttons created in any dialog or page."""

    caption = _normalized_caption(button.text())
    if not caption:
        return
    if str(button.property("_lugest_semantic_caption") or "") == caption:
        return
    is_destructive = any(word in caption for word in _DESTRUCTIVE_WORDS) or (
        caption.startswith("limpar ") and any(target in caption for target in _DESTRUCTIVE_CLEAR_TARGETS)
    )
    is_exit = caption in _EXIT_WORDS or any(caption.startswith(f"{word} ") for word in _EXIT_WORDS)
    destructive_now = is_destructive or is_exit
    if bool(getattr(button, "_lugest_was_destructive", False)) and not destructive_now:
        button.setStyleSheet(str(getattr(button, "_lugest_original_style", "") or ""))
        button.setIcon(getattr(button, "_lugest_original_icon"))
        button.setProperty("variant", getattr(button, "_lugest_original_variant", None))
        button.setProperty("_lugest_destructive_style", False)
        button._lugest_was_destructive = False
        _refresh_widget_style(button)
    if destructive_now:
        button.setProperty("_lugest_semantic_caption", caption)
        if not hasattr(button, "_lugest_original_style"):
            button._lugest_original_style = str(button.styleSheet() or "")
            button._lugest_original_icon = button.icon()
            button._lugest_original_variant = button.property("variant")
        button._lugest_was_destructive = True
        if button.property("variant") != "rejected":
            button.setProperty("variant", "destructive")
        if not bool(button.property("_lugest_destructive_style")):
            button.setProperty("_lugest_destructive_style", True)
            current_style = str(button.styleSheet() or "").rstrip()
            button.setStyleSheet(f"{current_style}\n{_DESTRUCTIVE_BUTTON_STYLE}".strip())
        icon_type = QStyle.SP_DialogCloseButton if is_exit and not is_destructive else QStyle.SP_TrashIcon
        button.setIcon(button.style().standardIcon(icon_type))
        button.setIconSize(QSize(16, 16))
        _refresh_widget_style(button)
        return
    if caption in _NEUTRAL_WORDS and not button.property("variant"):
        button.setProperty("_lugest_semantic_caption", caption)
        button.setProperty("variant", "secondary")
        _refresh_widget_style(button)
        return
    button.setProperty("_lugest_semantic_caption", caption)


def polish_table_selection(table: QTableView) -> None:
    """Keep row selection and its color identical in every current and future table."""

    if bool(table.property("_lugest_uniform_selection")):
        return
    table.setProperty("_lugest_uniform_selection", True)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setAlternatingRowColors(True)
    header = table.horizontalHeader()
    if header is not None:
        header.setHighlightSections(False)
    current_style = str(table.styleSheet() or "").rstrip()
    table.setStyleSheet(f"{current_style}\n{_UNIFORM_TABLE_SELECTION}".strip())


def polish_widget_tree(root: QWidget) -> None:
    if isinstance(root, QPushButton):
        polish_button_semantics(root)
    elif isinstance(root, QTableView):
        polish_table_selection(root)
    for button in root.findChildren(QPushButton):
        polish_button_semantics(button)
    for table in root.findChildren(QTableView):
        polish_table_selection(table)


class _GlobalUiPolishFilter(QObject):
    """Polish lazy pages and modal windows when they first become visible."""

    def eventFilter(self, watched, event):  # type: ignore[override]
        if event.type() == QEvent.Show:
            if isinstance(watched, QPushButton):
                polish_button_semantics(watched)
            elif isinstance(watched, QTableView):
                polish_table_selection(watched)
            elif isinstance(watched, QWidget) and watched.isWindow():
                polish_widget_tree(watched)
        return False

def apply_theme(app: QApplication, branding: dict) -> None:
    primary = str(branding.get("primary_color") or "#0b1f66").strip() or "#0b1f66"
    primary_color = QColor(primary)
    primary_luminance = (primary_color.red() * 0.299) + (primary_color.green() * 0.587) + (primary_color.blue() * 0.114)
    primary_text = "#0b1f33" if primary_luminance >= 165 else "#ffffff"
    selection_fill = "#eaf7da"
    selection_fill_soft = "#eaf7da"
    selection_border = "#7ed321"
    selection_text = "#26331d"
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#f1f2f0"))
    palette.setColor(QPalette.WindowText, QColor("#30343b"))
    palette.setColor(QPalette.Base, QColor("#ffffff"))
    palette.setColor(QPalette.AlternateBase, QColor("#f7f7f5"))
    palette.setColor(QPalette.ToolTipBase, QColor("#ffffff"))
    palette.setColor(QPalette.ToolTipText, QColor("#30343b"))
    palette.setColor(QPalette.Text, QColor("#30343b"))
    palette.setColor(QPalette.Button, QColor("#ffffff"))
    palette.setColor(QPalette.ButtonText, QColor("#30343b"))
    palette.setColor(QPalette.Highlight, QColor(selection_fill))
    palette.setColor(QPalette.HighlightedText, QColor(selection_text))
    app.setPalette(palette)
    app.setStyle("Fusion")
    app.setStyleSheet(
        f"""
        QWidget {{
            font-family: Segoe UI;
            font-size: 13px;
            color: #30343b;
        }}
        QMainWindow, QDialog {{
            background: #f1f2f0;
        }}
        QFrame#Card, QFrame#Panel {{
            background: #ffffff;
            border: 1px solid #d5d8d4;
            border-radius: 5px;
        }}
        QFrame#Card[tone="info"], QFrame#Panel[tone="info"] {{
            background: #f5f6f4;
            border: 1px solid #cfd3cf;
        }}
        QFrame#Card[tone="success"], QFrame#Panel[tone="success"] {{
            background: #f0f8e7;
            border: 1px solid #b9d994;
        }}
        QFrame#Card[tone="warning"], QFrame#Panel[tone="warning"] {{
            background: #fff8eb;
            border: 1px solid #e4c37f;
        }}
        QFrame#Card[tone="danger"], QFrame#Panel[tone="danger"] {{
            background: #efefed;
            border: 1px solid #aeb2ae;
        }}
        QFrame#Card[tone="rejected"], QFrame#Panel[tone="rejected"] {{
            background: #fff4f3;
            border: 1px solid #e5aea7;
        }}
        QFrame#TopBar {{
            background: #ffffff;
            border: 1px solid #d5d8d4;
            border-radius: 5px;
        }}
        QFrame#NavBar {{
            background: #ffffff;
            border: 1px solid #d5d8d4;
            border-radius: 5px;
        }}
        QFrame#QualitySection {{
            background: #f7f7f5;
            border: 1px solid #d5d8d4;
            border-left: 4px solid #7ed321;
            border-radius: 5px;
        }}
        QFrame#QualityToolbar {{
            background: #ffffff;
            border: 1px solid #d5d8d4;
            border-radius: 5px;
        }}
        QFrame#LoginCard {{
            background: #ffffff;
            border: 1px solid #c6d2e0;
            border-radius: 24px;
        }}
        QLineEdit, QComboBox, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {{
            background: #ffffff;
            border: 1px solid #c9cdca;
            border-radius: 4px;
            padding: 8px 10px;
            selection-background-color: {selection_fill};
            selection-color: {selection_text};
        }}
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {{
            border: 2px solid #7ed321;
        }}
        QPushButton {{
            background: #444744;
            color: #ffffff;
            border: 1px solid #343734;
            border-radius: 18px;
            padding: 8px 15px;
            font-weight: 800;
        }}
        QPushButton:hover {{
            background: #303330;
        }}
        QPushButton:pressed {{
            background: #252825;
        }}
        QPushButton:disabled {{
            background: #e4e8ec;
            color: #929ca7;
            border-color: #d2d8df;
        }}
        QPushButton[variant="secondary"] {{
            background: #ffffff;
            color: #3c4140;
            border: 1px solid #c8cbc8;
        }}
        QPushButton[variant="secondary"]:hover {{
            background: #f1f2f0;
            border-color: #9ea39e;
        }}
        QPushButton[variant="danger"] {{
            border: 1px solid #343734;
            background: #555955;
            color: #ffffff;
        }}
        QPushButton[variant="danger"]:hover {{
            background: #3e423e;
        }}
        QPushButton[variant="destructive"] {{
            border: 1px solid #8f1d14;
            background: #d92d20;
            color: #ffffff;
        }}
        QPushButton[variant="destructive"]:hover {{
            background: #b42318;
        }}
        QPushButton[variant="destructive"]:pressed {{
            background: #8f1d14;
        }}
        QPushButton[variant="rejected"] {{
            border: 1px solid #8f1d14;
            background: #d92d20;
            color: #ffffff;
        }}
        QPushButton[variant="rejected"]:hover {{
            background: #b42318;
        }}
        QPushButton[variant="success"] {{
            border: 1px solid #60aa14;
            background: #70c51a;
            color: #ffffff;
        }}
        QPushButton[variant="success"]:hover {{
            background: #60ad13;
        }}
        QPushButton[variant="warning"] {{
            border: 1px solid #d97706;
            background: #f59e0b;
            color: #ffffff;
        }}
        QPushButton[variant="warning"]:hover {{
            background: #d97706;
        }}
        QLineEdit[compact="true"], QComboBox[compact="true"], QTextEdit[compact="true"], QPlainTextEdit[compact="true"], QSpinBox[compact="true"], QDoubleSpinBox[compact="true"] {{
            padding: 4px 6px;
            font-size: 11px;
        }}
        QPushButton[compact="true"] {{
            padding: 5px 8px;
            font-size: 11px;
        }}
        QPushButton[qualityAction="true"] {{
            padding: 8px 14px;
            font-size: 11px;
            font-weight: 800;
            border-radius: 17px;
            border-width: 1px;
            border-style: solid;
        }}
        QPushButton[toolbarAction="true"] {{
            padding: 8px 16px;
            font-size: 12px;
            font-weight: 800;
            border-radius: 18px;
            border: 1px solid #343734;
            background: #444744;
        }}
        QPushButton[toolbarAction="true"]:hover {{
            background: #303330;
        }}
        QPushButton[toolbarAction="true"][variant="secondary"] {{
            color: #3c4140;
            border: 1px solid #c8cbc8;
            background: #ffffff;
        }}
        QPushButton[toolbarAction="true"][variant="secondary"]:hover {{
            background: #f1f2f0;
            border-color: #9ea39e;
        }}
        QPushButton[toolbarAction="true"][headerControl] {{
            padding: 0;
            color: #3f4542;
            background: #f7f8f6;
            border: 1px solid #c9cdca;
        }}
        QPushButton[toolbarAction="true"][headerControl="refresh"]:hover {{
            background: #eef7e7;
            border-color: #8cbf58;
        }}
        QPushButton[toolbarAction="true"][headerControl="close"]:hover {{
            background: #fff0ef;
            border-color: #d98b83;
        }}
        QPushButton[toolbarAction="true"][variant="success"] {{
            border: 1px solid #60aa14;
            background: #70c51a;
        }}
        QPushButton[toolbarAction="true"][variant="danger"] {{
            border: 1px solid #343734;
            background: #555955;
        }}
        QPushButton[toolbarAction="true"][variant="destructive"] {{
            border: 1px solid #8f1d14;
            background: #d92d20;
        }}
        QPushButton[toolbarAction="true"][variant="destructive"]:hover {{
            background: #b42318;
        }}
        QPushButton[toolbarAction="true"][variant="rejected"] {{
            border: 1px solid #8f1d14;
            background: #d92d20;
        }}
        QPushButton[dashboardSegment="true"] {{
            background: #ffffff;
            color: #3c4140;
            font-weight: 700;
            border: 1px solid #c8cbc8;
            border-radius: 18px;
            padding: 8px 16px;
        }}
        QPushButton[dashboardSegment="true"]:hover {{
            background: #ffffff;
        }}
        QPushButton[dashboardSegment="true"]:checked {{
            background: #444744;
            color: #ffffff;
            font-weight: 800;
            border: 1px solid #343734;
        }}
        QToolButton[nav="true"] {{
            background: transparent;
            color: #3c4140;
            border: 1px solid transparent;
            border-radius: 6px;
            padding: 6px 8px;
            text-align: center;
            font-size: 12px;
            font-weight: 700;
        }}
        QToolButton[nav="true"]:hover {{
            background: #f0f2ee;
            border: 1px solid #c9cdca;
        }}
        QToolButton[nav="true"]:checked {{
            background: #eaf7da;
            border: 1px solid #7ed321;
            color: #2f4f18;
        }}
        QLabel[role="muted"] {{
            color: #6b706f;
        }}
        QLabel[role="section_title"] {{
            color: #30343b;
            font-size: 15px;
            font-weight: 900;
        }}
        QLabel[role="section_subtitle"] {{
            color: #6b706f;
            font-size: 11px;
        }}
        QLabel[role="field_label"] {{
            color: #6b706f;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }}
        QLabel[role="field_value"] {{
            color: #30343b;
            font-size: 13px;
            font-weight: 700;
        }}
        QLabel[role="field_value_strong"] {{
            color: #0f172a;
            font-size: 17px;
            font-weight: 900;
        }}
        QLabel[role="badge"] {{
            background: {primary};
            color: {primary_text};
            border-radius: 999px;
            padding: 5px 11px;
            font-weight: 700;
        }}
        QLabel[role="alert_text"] {{
            color: #b45f06;
            font-weight: 800;
        }}
        QLabel[role="state_chip"] {{
            border-radius: 999px;
            padding: 6px 12px;
            font-weight: 800;
            border: 1px solid #c9cdca;
            background: #f4f5f3;
            color: #3c4140;
        }}
        QTableWidget {{
            background: #ffffff;
            alternate-background-color: #f7f7f5;
            border: 1px solid #cfd3cf;
            border-radius: 4px;
            gridline-color: #e1e3e0;
            selection-background-color: {selection_fill_soft};
            selection-color: {selection_text};
        }}
        QTableWidget::item:selected {{
            color: {selection_text};
            border: 1px solid {selection_border};
        }}
        QTableWidget[workspaceTable="true"] {{
            background: #ffffff;
            alternate-background-color: #f7f7f5;
            border: 1px solid #cfd3cf;
            border-radius: 4px;
            gridline-color: transparent;
            outline: 0;
        }}
        QTableWidget[workspaceTable="true"]::item {{
            padding: 5px 7px;
            border: 0;
            border-bottom: 1px solid #e4e5e3;
        }}
        QTableWidget[workspaceTable="true"]::item:hover {{
            background: #f0f5ea;
        }}
        QTableWidget[workspaceTable="true"]::item:selected {{
            background: {selection_fill_soft};
            color: {selection_text};
            border: 0;
            border-left: 3px solid {selection_border};
            border-bottom: 1px solid #d8e6c9;
        }}
        QTableWidget[workspaceTable="true"] QHeaderView::section {{
            background: #444744;
            color: #ffffff;
            border: 0;
            border-right: 1px solid #5c605c;
            padding: 7px 8px;
            font-size: 10px;
            font-weight: 800;
        }}
        QTableWidget#QualityTable {{
            alternate-background-color: #f7f7f5;
            gridline-color: #d9dcd8;
        }}
        QTableWidget#StockTable {{
            background: #ffffff;
            alternate-background-color: #f7f7f5;
            gridline-color: #d9dcd8;
            border: 1px solid #cfd3cf;
            border-radius: 4px;
        }}
        QCalendarWidget QAbstractItemView:enabled {{
            selection-background-color: {selection_fill};
            selection-color: {selection_text};
        }}
        QHeaderView::section {{
            background: #444744;
            color: #ffffff;
            border: 0;
            padding: 8px 6px;
            font-weight: 700;
        }}
        QProgressBar {{
            border: 1px solid #b9d994;
            border-radius: 7px;
            background: #f1f8e8;
            text-align: center;
            font-weight: 800;
            color: #355716;
            min-height: 18px;
        }}
        QProgressBar::chunk {{
            border-radius: 6px;
            background: #7ed321;
        }}
        QTabWidget::pane {{
            border: 1px solid #cfd3cf;
            border-radius: 5px;
            background: #ffffff;
            top: -1px;
        }}
        QTabWidget#QualityTabs::pane {{
            background: #f7f7f5;
            border: 1px solid #cfd3cf;
            border-radius: 5px;
            top: -1px;
        }}
        QTabWidget#QualityTabs QTabBar::tab {{
            background: #f0f1ef;
            color: #3c4140;
            border: 1px solid #cfd3cf;
            padding: 11px 18px;
            min-height: 26px;
            min-width: 112px;
            margin-right: 3px;
            margin-top: 4px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
            font-size: 12px;
            font-weight: 800;
        }}
        QTabWidget#QualityTabs QTabBar::tab:selected {{
            background: #ffffff;
            color: #355716;
            border-color: #7ed321;
            border-bottom: 1px solid #ffffff;
            margin-top: 0px;
            padding-top: 11px;
        }}
        QTabWidget#QualityTabs QTabBar::tab:hover {{
            background: #ffffff;
            border-color: #9ea39e;
        }}
        QTabBar::tab {{
            background: #f0f1ef;
            color: #3c4140;
            border: 1px solid #cfd3cf;
            padding: 9px 16px;
            margin-right: 4px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
            font-weight: 800;
        }}
        QTabBar::tab:selected {{
            background: #444744;
            border-color: #343734;
            color: #ffffff;
        }}
        QTabBar::tab:hover {{
            background: #ffffff;
        }}
        QScrollBar:vertical, QScrollBar:horizontal {{
            background: #eceeeb;
            border-radius: 6px;
            margin: 2px;
        }}
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
            background: #b8bcb8;
            border-radius: 6px;
            min-height: 26px;
            min-width: 26px;
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            background: transparent;
            border: none;
            width: 0px;
            height: 0px;
        }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
            background: transparent;
        }}
        """
    )
    polish_filter = getattr(app, "_lugest_ui_polish_filter", None)
    if polish_filter is None:
        polish_filter = _GlobalUiPolishFilter(app)
        app.installEventFilter(polish_filter)
        app._lugest_ui_polish_filter = polish_filter
    for window in app.topLevelWidgets():
        if isinstance(window, QWidget):
            polish_widget_tree(window)
