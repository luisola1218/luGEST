from __future__ import annotations

import os
import sys
from pathlib import Path


class _ReadOnlyClientBackend:
    def client_next_code(self) -> str:
        return "CL0001"


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

    from PySide6.QtWidgets import QApplication, QFrame, QLabel

    from lugest_qt.ui.pages.partners_pages import ClientsPage
    from lugest_qt.ui.theme import apply_theme

    app = QApplication.instance() or QApplication(["verify-partner-ui-layout"])
    apply_theme(app, {})
    page = ClientsPage(_ReadOnlyClientBackend())
    page.resize(1280, 720)
    page.show()
    app.processEvents()

    details_card = page.findChild(QFrame, "ClientDetailsCard")
    assert details_card is not None
    assert details_card.minimumWidth() >= 410
    assert "#6f9f45" in details_card.styleSheet()
    assert page.client_detail_tabs.minimumHeight() >= 410
    assert "#6f9f45" in page.client_detail_tabs.styleSheet()
    assert "#0aa6a6" not in page.client_detail_tabs.styleSheet()
    assert page.client_address_edit.minimumHeight() >= 105
    assert page.client_notes_edit.minimumHeight() >= 125
    labels = {label.text() for label in details_card.findChildren(QLabel)}
    assert {"Ficha do cliente", "Identificação", "Contacto e localização", "Condições comerciais"}.issubset(labels)

    page.close()
    print("partner-ui-layout-ok palette=green-gray tabs=readable details=responsive")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
