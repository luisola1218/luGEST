from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication

from lugest_core.materials import profile_entry, profile_sizes
from lugest_qt.services.legacy_backend import LegacyBackend
from lugest_qt.ui.pages.materials_page import _MaterialEditorDialog


def main() -> int:
    expected = {
        ("IPE", "220"): (220.0, 110.0, 5.9, 9.2, 26.2),
        ("IPN", "300"): (300.0, 125.0, 10.8, 16.2, 54.2),
        ("UPN", "200"): (200.0, 75.0, 8.5, 6.0, 25.3),
        ("HEA", "300"): (290.0, 300.0, 8.5, 14.0, 88.3),
        ("HEB", "300"): (300.0, 300.0, 11.0, 19.0, 117.0),
    }
    for (series, size), values in expected.items():
        row = profile_entry(series, size)
        actual = tuple(float(row[key]) for key in ("h", "b", "tw", "tf", "kg_m"))
        assert actual == values, (series, size, actual, values)
        assert size in profile_sizes(series)

    backend = LegacyBackend()
    command = "cria stock de 10 vigas Perfil IPE 220 S275JR com 6 metros de comprimento"
    candidate = backend._local_material_command_candidate(command)
    assert candidate["formato"] == "Perfil"
    assert candidate["secao_tipo"] == "IPE"
    assert candidate["altura"] == "220"
    assert candidate["espessura"] == ""
    assert candidate["metros"] == "6"
    assert candidate["quantidade"] == "10"

    natural_command = "Cria em stock 10 PERFIL IPE s275jr tamanho 220 com 6metros"
    natural_candidate = backend._local_material_command_candidate(natural_command)
    assert natural_candidate["formato"] == "Perfil"
    assert natural_candidate["secao_tipo"] == "IPE"
    assert natural_candidate["altura"] == "220"
    assert natural_candidate["metros"] == "6"
    assert natural_candidate["quantidade"] == "10"

    preview = backend.material_geometry_preview(candidate)
    assert float(preview["kg_m"]) == 26.2
    assert float(preview["peso_unid"]) == 157.2

    app = QApplication.instance() or QApplication([])
    dialog = _MaterialEditorDialog(
        backend,
        record={**candidate, "_ai_confidence": 0.95, "missing_fields": []},
        mode="ai",
    )
    assert dialog._current_profile_series() == "IPE"
    assert dialog.profile_size_combo.currentText() == "220"
    assert dialog.espessura_combo.isHidden()
    assert float(dialog.kg_m_edit.text().replace(",", ".")) == 26.2
    assert float(dialog.peso_edit.text().replace(",", ".")) == 157.2
    dialog.close()
    app.processEvents()
    print("Perfis estruturais: catálogo, Copiloto, cálculo e formulário verificados.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
