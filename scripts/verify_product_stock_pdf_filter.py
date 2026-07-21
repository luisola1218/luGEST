from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pypdf import PdfReader
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QCheckBox, QDialog, QPushButton

from lugest_qt.services.legacy_backend import LegacyBackend
from lugest_qt.ui.pages.products_page import ProductsPage


def main() -> int:
    backend = LegacyBackend()
    options = backend.product_stock_filters()
    category = options["categories"][0]
    target = Path(tempfile.gettempdir()) / "lugest_product_stock_filter_test.pdf"
    backend.product_render_stock_pdf(target, categories=[category])
    text = "\n".join(page.extract_text() or "" for page in PdfReader(str(target)).pages)

    products = list(backend.ensure_data().get("produtos", []) or [])
    selected_codes = [
        str(row.get("codigo", "") or "")
        for row in products
        if str(row.get("categoria", "") or "Sem categoria").strip() == category
    ]
    other_codes = [
        str(row.get("codigo", "") or "")
        for row in products
        if str(row.get("categoria", "") or "Sem categoria").strip() != category
    ]
    assert selected_codes and selected_codes[0] in text
    assert not any(code and code in text for code in other_codes)
    assert f"Categorias: {category}" in text

    app = QApplication.instance() or QApplication(sys.argv)
    page = ProductsPage(backend)
    assert "Preview stock" in [button.text() for button in page.findChildren(QPushButton)]
    opened: dict[str, object] = {}
    backend.product_open_stock_pdf = lambda **kwargs: opened.update(kwargs) or target

    def choose_category() -> None:
        dialog = app.activeModalWidget()
        assert isinstance(dialog, QDialog)
        checks = {check.text(): check for check in dialog.findChildren(QCheckBox)}
        checks["Todas as categorias"].setChecked(False)
        checks[category].setChecked(True)
        dialog.accept()

    QTimer.singleShot(250, choose_category)
    page._open_stock_pdf()
    assert opened.get("categories") == [category]
    assert opened.get("types") is None
    print(
        f"product-stock-pdf-filter-ok categories={len(options['categories'])} "
        f"types={len(options['types'])} selected_rows={len(selected_codes)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
