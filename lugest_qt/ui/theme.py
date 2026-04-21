from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


def _shade(color_hex: str, factor: float) -> str:
    color = QColor(color_hex)
    if factor >= 1.0:
        color = color.lighter(int(100 * factor))
    else:
        color = color.darker(int(100 / max(0.01, factor)))
    return color.name()


def apply_theme(app: QApplication, branding: dict) -> None:
    primary = str(branding.get("primary_color") or "#0b1f66").strip() or "#0b1f66"
    primary_dark = _shade(primary, 0.82)
    primary_soft = _shade(primary, 1.65)
    primary_surface = _shade(primary, 1.9)
    selection_fill = "#efe4c8"
    selection_fill_soft = "rgba(198, 168, 108, 0.26)"
    selection_border = "#c9b387"
    selection_text = "#0f172a"
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#e3ebf4"))
    palette.setColor(QPalette.WindowText, QColor("#14243d"))
    palette.setColor(QPalette.Base, QColor("#ffffff"))
    palette.setColor(QPalette.AlternateBase, QColor("#f5f8fc"))
    palette.setColor(QPalette.ToolTipBase, QColor("#ffffff"))
    palette.setColor(QPalette.ToolTipText, QColor("#14243d"))
    palette.setColor(QPalette.Text, QColor("#14243d"))
    palette.setColor(QPalette.Button, QColor("#ffffff"))
    palette.setColor(QPalette.ButtonText, QColor("#14243d"))
    palette.setColor(QPalette.Highlight, QColor(selection_fill))
    palette.setColor(QPalette.HighlightedText, QColor(selection_text))
    app.setPalette(palette)
    app.setStyle("Fusion")
    app.setStyleSheet(
        f"""
        QWidget {{
            font-family: Segoe UI;
            font-size: 13px;
            color: #14243d;
        }}
        QMainWindow, QDialog {{
            background: #e3ebf4;
        }}
        QFrame#Card, QFrame#Panel {{
            background: #fbfdff;
            border: 1px solid #aebfd2;
            border-radius: 10px;
        }}
        QFrame#Card[tone="info"], QFrame#Panel[tone="info"] {{
            background: #f2f7fd;
            border: 1px solid #a8c3e2;
        }}
        QFrame#Card[tone="success"], QFrame#Panel[tone="success"] {{
            background: #f3fbf5;
            border: 1px solid #a9d2b8;
        }}
        QFrame#Card[tone="warning"], QFrame#Panel[tone="warning"] {{
            background: #fff8eb;
            border: 1px solid #e4c37f;
        }}
        QFrame#Card[tone="danger"], QFrame#Panel[tone="danger"] {{
            background: #fff4f3;
            border: 1px solid #e5aea7;
        }}
        QFrame#TopBar {{
            background: rgba(255, 255, 255, 0.95);
            border: 1px solid #afc0d4;
            border-radius: 10px;
        }}
        QFrame#NavBar {{
            background: #f9fbfe;
            border: 1px solid #afc0d4;
            border-radius: 10px;
        }}
        QFrame#LoginCard {{
            background: #ffffff;
            border: 1px solid #c6d2e0;
            border-radius: 24px;
        }}
        QLineEdit, QComboBox, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {{
            background: #ffffff;
            border: 1px solid #b4c4d6;
            border-radius: 6px;
            padding: 8px 10px;
            selection-background-color: {selection_fill};
            selection-color: {selection_text};
        }}
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {{
            border: 1px solid {primary};
        }}
        QPushButton {{
            background: {primary};
            color: #ffffff;
            border: 0;
            border-radius: 6px;
            padding: 9px 13px;
            font-weight: 700;
        }}
        QPushButton:hover {{
            background: {primary_dark};
        }}
        QPushButton[variant="secondary"] {{
            background: #eef4fa;
            color: #17314f;
            border: 1px solid #b7c7d8;
        }}
        QPushButton[variant="secondary"]:hover {{
            background: #dde8f4;
        }}
        QPushButton[variant="danger"] {{
            background: #b42318;
        }}
        QPushButton[variant="danger"]:hover {{
            background: #8f1d14;
        }}
        QPushButton[variant="success"] {{
            background: #0f8a6a;
            color: #ffffff;
        }}
        QPushButton[variant="success"]:hover {{
            background: #0b6f56;
        }}
        QPushButton[variant="warning"] {{
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
        QToolButton[nav="true"] {{
            background: transparent;
            color: #17314f;
            border: 1px solid transparent;
            border-radius: 6px;
            padding: 7px 10px;
            text-align: center;
            font-weight: 700;
        }}
        QToolButton[nav="true"]:hover {{
            background: #edf3fa;
            border: 1px solid #c7d5e4;
        }}
        QToolButton[nav="true"]:checked {{
            background: #dbe7f5;
            border: 1px solid #8faed0;
            color: #10253d;
        }}
        QLabel[role="muted"] {{
            color: #5b6f86;
        }}
        QLabel[role="field_label"] {{
            color: #5b6f86;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }}
        QLabel[role="field_value"] {{
            color: #14243d;
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
            color: #ffffff;
            border-radius: 999px;
            padding: 5px 11px;
            font-weight: 700;
        }}
        QLabel[role="alert_text"] {{
            color: #b42318;
            font-weight: 800;
        }}
        QLabel[role="state_chip"] {{
            border-radius: 999px;
            padding: 6px 12px;
            font-weight: 800;
            border: 1px solid #c6d2e0;
            background: #f4f7fb;
            color: #23364d;
        }}
        QTableWidget {{
            background: #ffffff;
            alternate-background-color: #f7fafd;
            border: 1px solid #afc0d4;
            border-radius: 8px;
            gridline-color: #d6e0ea;
            selection-background-color: {selection_fill_soft};
            selection-color: {selection_text};
        }}
        QTableWidget::item:selected {{
            color: {selection_text};
            border: 1px solid {selection_border};
        }}
        QCalendarWidget QAbstractItemView:enabled {{
            selection-background-color: {selection_fill};
            selection-color: {selection_text};
        }}
        QHeaderView::section {{
            background: {primary};
            color: #ffffff;
            border: 0;
            padding: 8px 6px;
            font-weight: 700;
        }}
        QProgressBar {{
            border: 1px solid #c6d2e0;
            border-radius: 7px;
            background: #edf3f9;
            text-align: center;
            font-weight: 800;
            color: #17314f;
            min-height: 18px;
        }}
        QProgressBar::chunk {{
            border-radius: 6px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {primary}, stop:1 {primary_dark});
        }}
        QTabWidget::pane {{
            border: 1px solid #c6d2e0;
            border-radius: 8px;
            background: #ffffff;
            top: -1px;
        }}
        QTabBar::tab {{
            background: #edf2f8;
            color: #24384e;
            border: 1px solid #c7d2df;
            border-bottom: 0;
            padding: 8px 14px;
            margin-right: 4px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            font-weight: 700;
        }}
        QTabBar::tab:selected {{
            background: #ffffff;
            color: {primary_dark};
        }}
        QTabBar::tab:hover {{
            background: #f4f7fb;
        }}
        QScrollBar:vertical, QScrollBar:horizontal {{
            background: #edf2f8;
            border-radius: 6px;
            margin: 2px;
        }}
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
            background: #b8c5d5;
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
