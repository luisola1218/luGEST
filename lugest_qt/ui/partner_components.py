from __future__ import annotations

import re
import unicodedata
from urllib.parse import quote_plus

from PySide6.QtCore import QUrl, Qt, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QScrollArea,
    QSplitter,
    QStyle,
    QTableWidget,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from lugest_qt.ui.widgets import CardFrame
from lugest_qt.ui.pages.runtime_common import (
    configure_table as _configure_table,
    fill_table as _fill_table,
    selected_row_index as _selected_row_index,
    set_table_columns as _set_table_columns,
)


PAYMENT_TERMS_OPTIONS = ["", "Pronto Pagamento", "30 dias", "60 dias", "90 dias", "180 dias"]

def _open_google_maps(parent: QWidget, address: str, latitude: str = "", longitude: str = "") -> None:
    lat = str(latitude or "").strip().replace(",", ".")
    lon = str(longitude or "").strip().replace(",", ".")
    query = f"{lat},{lon}" if lat and lon else str(address or "").strip()
    if not query:
        QMessageBox.information(parent, "Localização", "Preenche a morada ou as coordenadas.")
        return
    QDesktopServices.openUrl(QUrl(f"https://www.google.com/maps/search/?api=1&query={quote_plus(query)}"))


def _section_card(title: str, subtitle: str = "", tone: str = "default", minimum_height: int = 0) -> tuple[CardFrame, QFormLayout]:
    card = CardFrame()
    card.set_tone(tone)
    if minimum_height:
        card.setMinimumHeight(minimum_height)
    layout = QVBoxLayout(card)
    layout.setContentsMargins(15, 14, 15, 14)
    layout.setSpacing(10)
    title_label = QLabel(title)
    title_label.setStyleSheet("font-size: 13px; font-weight: 900; color: #29442d;")
    layout.addWidget(title_label)
    if subtitle:
        subtitle_label = QLabel(subtitle)
        subtitle_label.setProperty("role", "muted")
        subtitle_label.setWordWrap(True)
        subtitle_label.setStyleSheet("font-size: 10px; color: #68736b;")
        layout.addWidget(subtitle_label)
    form = QFormLayout()
    form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    form.setHorizontalSpacing(14)
    form.setVerticalSpacing(10)
    form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
    layout.addLayout(form)
    return card, form


def _scrollable_form_area(content_layout: QGridLayout) -> QScrollArea:
    content = QWidget()
    content.setLayout(content_layout)
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QScrollArea.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    scroll.setWidget(content)
    return scroll


def _prepare_partner_fields(*widgets: QWidget) -> None:
    for widget in widgets:
        if isinstance(widget, (QLineEdit, QComboBox)):
            widget.setMinimumHeight(max(36, widget.minimumHeight()))
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        elif isinstance(widget, QTextEdit):
            widget.setMinimumHeight(max(72, widget.minimumHeight()))
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)


def _search_box(edit: QLineEdit, object_name: str) -> QWidget:
    box = QWidget()
    box.setObjectName(object_name)
    box.setStyleSheet(
        f"QWidget#{object_name} {{ background: #ffffff; border: 1px solid #bdc9b9; border-radius: 8px; }}"
        "QLineEdit { border: none; background: transparent; padding: 7px 8px 7px 0; }"
    )
    layout = QHBoxLayout(box)
    layout.setContentsMargins(9, 0, 9, 0)
    layout.setSpacing(6)
    icon = QLabel("🔍")
    icon.setFixedWidth(20)
    icon.setAlignment(Qt.AlignCenter)
    icon.setStyleSheet("font-size: 15px; color: #4c6b43;")
    layout.addWidget(icon)
    layout.addWidget(edit, 1)
    return box


def _search_terms(value: str) -> list[str]:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-zA-Z0-9.]+", " ", text.casefold())
    return [part for part in text.split() if part]


def _row_matches_terms(row: dict, query: str) -> bool:
    terms = _search_terms(query)
    if not terms:
        return True
    text = unicodedata.normalize("NFKD", " ".join(str(value or "") for value in row.values()))
    haystack = re.sub(r"[^a-zA-Z0-9.]+", " ", "".join(ch for ch in text if not unicodedata.combining(ch)).casefold())
    return all(term in haystack for term in terms)


def _metric_chip(title: str, value: str = "-", tone: str = "default") -> QLabel:
    colors = {
        "default": ("#f0f3ef", "#4a5850"),
        "info": ("#edf3ea", "#3e5f35"),
        "success": ("#e9f4e3", "#315d2a"),
        "warning": ("#f4f2e8", "#6c6440"),
    }
    bg, fg = colors.get(tone, colors["default"])
    label = QLabel(f"{title}: {value}")
    label.setStyleSheet(
        f"background: {bg}; color: {fg}; border: 1px solid #c9d4c5; "
        "border-radius: 6px; padding: 4px 8px; font-size: 9px; font-weight: 800;"
    )
    return label


def _partner_detail_tabs(entries: tuple[tuple[str, CardFrame], ...]) -> QTabWidget:
    tabs = QTabWidget()
    tabs.setDocumentMode(True)
    tabs.setStyleSheet(
        """
        QTabWidget::pane {
            border: 1px solid #d0d8cd;
            border-radius: 7px;
            background: #f8faf7;
            top: -1px;
        }
        QTabBar::tab {
            min-width: 0;
            min-height: 34px;
            padding: 5px 7px;
            margin: 0;
            border: 1px solid #d0d8cd;
            background: #edf2eb;
            color: #4d5b51;
            font-size: 10px;
            font-weight: 800;
        }
        QTabBar::tab:selected {
            background: #ffffff;
            color: #355d2c;
            border-top: 2px solid #6f9f45;
            border-bottom-color: #ffffff;
        }
        QTabBar::tab:hover:!selected {
            background: #f4f7f2;
            color: #45673b;
        }
        """
    )
    tabs.tabBar().setExpanding(True)
    tabs.tabBar().setUsesScrollButtons(False)
    tabs.tabBar().setElideMode(Qt.ElideNone)
    for title, card in entries:
        card.setObjectName("PartnerTabSection")
        card.setStyleSheet("QFrame#PartnerTabSection { background: transparent; border: 0; border-radius: 0; }")
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(card)
        layout.addStretch(1)
        tabs.addTab(page, title)
    return tabs

