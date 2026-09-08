"""Presentation styles for quote panels; no application behavior."""

STYLE_OVERVIEW_BAND = """
            QFrame#QuoteOverviewBand {
                background: #ffffff;
                border: 1px solid #d8e0e8;
                border-radius: 8px;
            }
            QFrame#QuoteMetric {
                background: transparent;
                border: 0;
                border-left: 3px solid #7b91a8;
            }
            QFrame#QuoteMetric QLabel {
                background: transparent;
                border: 0;
            }
            """

STYLE_QUOTE_DETAIL_SCROLL = """QScrollArea { background: transparent; border: 0; }QScrollBar:vertical { width: 12px; background: #eef1ee; border: 0; margin: 2px; }QScrollBar::handle:vertical { min-height: 36px; background: #a6aea8; border-radius: 5px; }QScrollBar::handle:vertical:hover { background: #858e87; }QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }"""

STYLE_QUOTE_NOTES_CARD = """
            QFrame#Card {
                background: #fffaf0;
                border: 1.5px solid #e4bc6a;
                border-radius: 16px;
            }
            QLabel {
                color: #6f4a12;
            }
            QFrame#Card QLineEdit,
            QFrame#Card QComboBox,
            QFrame#Card QSpinBox,
            QFrame#Card QDoubleSpinBox {
                background: #ffffff;
                border: 1px solid #d4b170;
                border-radius: 8px;
                padding: 5px 8px;
            }
            QFrame#Card QComboBox::drop-down,
            QFrame#Card QSpinBox::down-button,
            QFrame#Card QDoubleSpinBox::down-button,
            QFrame#Card QSpinBox::up-button,
            QFrame#Card QDoubleSpinBox::up-button {
                width: 22px;
                border-left: 1px solid #e2c690;
                background: #fff6e4;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }
            """

STYLE_NOTES_TABS = """
            QTabWidget::pane {
                border: 1px solid #cfd8cb;
                border-radius: 7px;
                top: -1px;
                background: rgba(255, 255, 255, 0.92);
            }
            QTabBar::tab {
                font-size: 10px;
                min-height: 34px;
                min-width: 0;
                padding: 5px 7px;
                margin-right: 0;
                border: 1px solid #cbd7c5;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                background: #eef4e9;
                color: #42523f;
                font-weight: 700;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                border-color: #81a962;
                color: #2f432d;
                font-weight: 800;
            }
            QTabBar::tab:hover:!selected {
                background: #f5f8f2;
            }
            """

STYLE_TRANSPORT_FORM_CARD = """
            QFrame#Card {
                background: #f7f8f6;
                border: 1px solid #d5dbd3;
                border-radius: 8px;
            }
            QFrame#Card QLabel {
                font-size: 9.5px;
                color: #536057;
                font-weight: 700;
            }
            QFrame#Card QLineEdit,
            QFrame#Card QComboBox,
            QFrame#Card QAbstractSpinBox {
                background: #ffffff;
                border: 1px solid #c5d0c1;
                border-bottom: 1px solid #95aa8d;
                border-radius: 7px;
                font-size: 10px;
                min-height: 27px;
                padding: 1px 6px;
            }
            QFrame#Card QComboBox::drop-down,
            QFrame#Card QAbstractSpinBox::up-button,
            QFrame#Card QAbstractSpinBox::down-button {
                width: 20px;
                border-left: 1px solid #c5d0c1;
                border-bottom: 1px solid #95aa8d;
                background: #eef4ea;
                border-top-right-radius: 7px;
                border-bottom-right-radius: 7px;
            }
            QFrame#Card QComboBox::down-arrow {
                width: 9px;
                height: 9px;
            }
            QFrame#TransportFieldBox {
                background: #ffffff;
                border: 1px solid #bdcabb;
                border-bottom: 1px solid #859d7d;
                border-radius: 6px;
            }
            QFrame#TransportFieldBox QLineEdit,
            QFrame#TransportFieldBox QComboBox,
            QFrame#TransportFieldBox QAbstractSpinBox {
                background: transparent;
                border: 0;
                border-radius: 0;
                padding: 0 5px;
                min-height: 21px;
                font-size: 10px;
            }
            QFrame#TransportFieldBox QComboBox::drop-down,
            QFrame#TransportFieldBox QAbstractSpinBox::up-button,
            QFrame#TransportFieldBox QAbstractSpinBox::down-button {
                width: 18px;
                border-left: 1px solid #c9d4c5;
                border-bottom: 0;
                background: #f0f5ed;
                border-top-right-radius: 5px;
                border-bottom-right-radius: 5px;
            }
            """

STYLE_TRANSPORT_ACTIONS_CARD = """
            QFrame#Card {
                background: #f7f8f6;
                border: 1px solid #d5dbd3;
                border-radius: 8px;
            }
            QFrame#Card QLabel {
                font-size: 9.5px;
                color: #536057;
            }
            """

STYLE_TOTAL_PANEL = """
            QFrame#QuoteSummaryTotalPanel {
                background: #f5f8f3;
                border: 1px solid #c5d3c0;
                border-left: 4px solid #6f9f45;
                border-radius: 7px;
            }
            QFrame#QuoteSummaryTotalPanel QLabel {
                background: transparent;
                border: 0;
            }
            """

STYLE_SUMMARY_ROWS_HOST = """
            QFrame#QuoteFinancialBreakdown {
                background: #ffffff;
                border: 1px solid #d7ded5;
                border-radius: 7px;
            }
            QFrame#QuoteFinancialBreakdown QLabel {
                background: transparent;
                border: 0;
                padding: 0;
            }
            QFrame#QuoteSummaryDivider {
                background: #e6ebe4;
                border: 0;
            }
            """

STYLE_CONTROLS_PANEL = """
            QFrame#QuoteFinancialControls {
                background: #f5f7f4;
                border: 1px solid #d7ded5;
                border-radius: 7px;
            }
            QFrame#QuoteFinancialControls QLabel {
                background: transparent;
                border: 0;
            }
            """

STYLE_DISCOUNT_STATUS = """
            QFrame#QuoteDiscountStatus {
                background: #f7f9f6;
                border: 1px solid #d8e0d5;
                border-radius: 6px;
            }
            QFrame#QuoteDiscountStatus QLabel {
                background: transparent;
                border: 0;
            }
            """

STYLE_LINE_TOOLS_TABS = """
            QTabWidget::pane {
                border: 1px solid #d8e0e8;
                border-radius: 6px;
                background: #f8fafc;
                top: -1px;
            }
            QTabBar::tab {
                min-height: 22px;
                min-width: 116px;
                padding: 3px 12px;
                margin-right: 3px;
                border: 1px solid #d8e0e8;
                background: #eef3f7;
                color: #475467;
                font-size: 9px;
                font-weight: 800;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #132238;
                border-bottom-color: #ffffff;
            }
            """

STYLE_LINES_TABLE = """QTableWidget { font-family: 'Segoe UI'; font-size: 10px; } QTableWidget::indicator { width: 18px; height: 18px; } QHeaderView::section { font-family: 'Segoe UI Semibold'; font-size: 10px; padding: 5px 5px; font-weight: 600; } QScrollBar:vertical { background: #eceeeb; width: 14px; margin: 0; border-left: 1px solid #cfd3cf; } QScrollBar::handle:vertical { background: #8f9691; min-height: 28px; border-radius: 6px; margin: 2px; } QScrollBar::handle:vertical:hover { background: #747b76; } QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; } QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }"""

STYLE_SELECTED_LINE_FOOTER = """
            QFrame#QuoteSelectedLineFooter {
                background: #f1f3f1;
                border: 1px solid #d4d8d4;
                border-radius: 7px;
            }
            QLabel#QuoteSelectedLineCaption {
                color: #3f4943;
                font-size: 9px;
                font-weight: 900;
                background: transparent;
                border: 0;
            }
            """

STYLE_INSPECTOR_TABS = """
            QTabWidget::pane {
                border: 1px solid #d0d8cd;
                border-radius: 7px;
                background: #f7f9f6;
                top: -1px;
            }
            QTabBar::tab {
                min-height: 30px;
                min-width: 0;
                padding: 5px 4px;
                margin-right: 0;
                border: 1px solid #d0d8cd;
                background: #edf2eb;
                color: #4d5b51;
                font-size: 9px;
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

STYLE_SUMMARY_CONTROL = """QDoubleSpinBox, QComboBox, QPushButton { background: #ffffff; border: 1px solid #bdc9b9; border-bottom: 1px solid #8da086; border-radius: 6px; padding: 0 8px; font-size: 10px; font-weight: 700; color: #26342b;}QDoubleSpinBox::up-button, QDoubleSpinBox::down-button, QComboBox::drop-down { width: 20px; border-left: 1px solid #cbd5c8; background: #f2f5f0;}QPushButton:hover, QComboBox:hover, QDoubleSpinBox:hover { border-color: #6f9f45; background: #fbfdf9;}"""

STYLE_PANEL = """
                QFrame#QuoteInspectorSection {
                    background: transparent;
                    border: 0;
                    border-radius: 0;
                }
                """
