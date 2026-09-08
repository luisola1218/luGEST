"""Compatibility entry point for the desktop page registry."""
from lugest_modules.quotes.presentation.page import QuotePage
from lugest_qt.services.quote_page_composition import quote_page_services


class QuotesPage(QuotePage):
    def __init__(self, backend, parent=None):
        super().__init__(quote_page_services(backend), parent)
