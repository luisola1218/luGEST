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
from PySide6.QtWidgets import QApplication

from lugest_qt.services.legacy_backend import LegacyBackend
from lugest_qt.ui.pages.materials_page import MaterialsPage
from lugest_qt.ui.pages.products_page import ProductsPage


def _pdf_text(path: Path) -> str:
    return "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)


def main() -> int:
    backend = LegacyBackend()
    migration = backend.ensure_inventory_scan_codes(persist=True)
    data = backend.ensure_data()
    material = next(row for row in data.get("materiais", []) if str(row.get("id", "") or "").strip())
    product = next(row for row in data.get("produtos", []) if str(row.get("codigo", "") or "").strip())
    material_code = backend.inventory_scan_code("MAT", material["id"])
    product_code = backend.inventory_scan_code("PRD", product["codigo"])

    assert backend.inventory_scan_lookup(material_code, "MAT")["entity_id"] == str(material["id"]).upper()
    assert backend.inventory_scan_lookup(product_code, "PRD")["entity_id"] == str(product["codigo"]).upper()
    try:
        backend.inventory_scan_lookup(material_code, "PRD")
    except ValueError:
        pass
    else:
        raise AssertionError("A protecao entre material e produto nao foi aplicada.")

    output_dir = Path(tempfile.mkdtemp(prefix="lugest_inventory_scan_"))
    material_pdf = backend.material_identification_label_pdf(material["id"], output_dir / "material.pdf")
    product_pdf = backend.product_label_pdf(product["codigo"], output_dir / "product.pdf")
    assert material_code in _pdf_text(material_pdf)
    assert product_code in _pdf_text(product_pdf)

    app = QApplication.instance() or QApplication(sys.argv)
    materials_page = MaterialsPage(backend)
    materials_page.filter_edit.setText(material_code)
    materials_page.filter_edit.returnPressed.emit()
    assert materials_page.current_material_id == material["id"]
    products_page = ProductsPage(backend)
    products_page.filter_edit.setText(product_code)
    products_page.filter_edit.returnPressed.emit()
    assert products_page.current_code == product["codigo"]

    conn = backend.desktop_main._mysql_connect()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) AS total FROM inventory_scan_codes")
            table_count = int(cursor.fetchone()["total"])
    finally:
        conn.close()
    assert table_count == int(migration["entries"])
    print(f"inventory-scan-flow-ok material={material_code} product={product_code} db_rows={table_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
