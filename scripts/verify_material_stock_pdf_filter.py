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
from lugest_qt.ui.pages.materials_page import MaterialsPage


def main() -> int:
    backend = LegacyBackend()
    formats = backend.material_stock_formats()
    assert "Chapa" in formats

    target = Path(tempfile.gettempdir()) / "lugest_stock_chapa_filter_test.pdf"
    backend.material_render_stock_pdf(target, formats=["Chapa"])
    text = "\n".join(page.extract_text() or "" for page in PdfReader(str(target)).pages)
    assert "Formatos selecionados: Chapa" in text

    materials = list(backend.ensure_data().get("materiais", []) or [])
    searchable_material = next(
        (
            row
            for row in materials
            if str(row.get("formato", "") or "").strip().casefold() == "chapa"
            and str(row.get("material", "") or "").strip()
            and str(row.get("espessura", "") or "").strip()
        ),
        None,
    )
    assert searchable_material is not None
    smart_query = " ".join(
        (
            str(searchable_material.get("formato", "")).strip(),
            str(searchable_material.get("material", "")).strip(),
            f"{str(searchable_material.get('espessura', '')).strip()}mm",
        )
    )
    smart_ids = {
        str(payload.get("row", {}).get("id", "") or "").strip()
        for payload in backend.material_rows(smart_query)
    }
    assert str(searchable_material.get("id", "") or "").strip() in smart_ids

    selected_ids = [
        str(row.get("id", "") or "")
        for row in materials
        if str(row.get("formato", "") or "Chapa").strip() == "Chapa"
    ]
    other_ids = [
        str(row.get("id", "") or "")
        for row in materials
        if str(row.get("formato", "") or "Chapa").strip() != "Chapa"
    ]
    assert selected_ids and selected_ids[0] in text
    assert not any(material_id and material_id in text for material_id in other_ids)

    app = QApplication.instance() or QApplication(sys.argv)
    page = MaterialsPage(backend)
    page.filter_edit.setText(smart_query)
    page.refresh()
    assert page.table.rowCount() > 0
    expected_material = str(searchable_material.get("material", "") or "").strip().casefold()
    expected_thickness = str(searchable_material.get("espessura", "") or "").strip().replace(",", ".")
    for row_index in range(page.table.rowCount()):
        assert (page.table.item(row_index, 2).text() or "").strip().casefold() == expected_material
        assert (page.table.item(row_index, 5).text() or "").strip().replace(",", ".") == expected_thickness
        assert (page.table.item(row_index, 8).text() or "").strip().casefold() == "chapa"

    button_labels = [button.text() for button in page.findChildren(QPushButton)]
    assert button_labels.count("Preview PDF") == 1
    assert "PDF" not in button_labels

    opened: dict[str, object] = {}
    backend.material_open_stock_pdf = lambda **kwargs: opened.update(kwargs) or target

    def choose_sheet_only() -> None:
        dialog = app.activeModalWidget()
        assert isinstance(dialog, QDialog)
        checks = {check.text(): check for check in dialog.findChildren(QCheckBox)}
        checks["Todos os formatos"].setChecked(False)
        checks["Chapa"].setChecked(True)
        dialog.accept()

    QTimer.singleShot(250, choose_sheet_only)
    page.preview_pdf()
    assert opened.get("formats") == ["Chapa"]
    print(
        "material-stock-pdf-filter-ok "
        f"formats={len(formats)} chapa_rows={len(selected_ids)} smart_query={smart_query!r}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
