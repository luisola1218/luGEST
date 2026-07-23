from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from lugest_core.laser.quote_engine import default_laser_quote_settings
from lugest_qt.services.laser_nesting import nest_parts
from lugest_qt.ui.pages.laser_nesting_dialog import LaserNestingDialog, SheetProfilesDialog


TEST_DXF = """0
SECTION
2
ENTITIES
0
LWPOLYLINE
8
CUT
90
4
70
1
10
0
20
0
10
300
20
0
10
300
20
180
10
0
20
180
0
ENDSEC
0
EOF
"""


def _complex_dxf() -> str:
    entities = [
        "0", "SECTION", "2", "ENTITIES",
        "0", "LWPOLYLINE", "8", "CUT", "90", "4", "70", "1",
        "10", "0", "20", "0", "10", "600", "20", "0",
        "10", "600", "20", "300", "10", "0", "20", "300",
    ]
    for index in range(60):
        col = index % 12
        row = index // 12
        entities.extend(
            [
                "0", "CIRCLE", "8", "CUT",
                "10", str(25 + (col * 50)),
                "20", str(30 + (row * 55)),
                "40", "6",
            ]
        )
    entities.extend(["0", "ENDSEC", "0", "EOF"])
    return "\n".join(entities) + "\n"


def _assert_placements_within_sheet(result: dict) -> None:
    for sheet in list(result.get("sheets", []) or []):
        width = float(sheet.get("sheet_width_mm", 0) or 0)
        height = float(sheet.get("sheet_height_mm", 0) or 0)
        for placement in list(sheet.get("placements", []) or []):
            assert float(placement.get("x_mm", 0) or 0) >= 0.0
            assert float(placement.get("y_mm", 0) or 0) >= 0.0
            assert float(placement.get("x_mm", 0) or 0) + float(placement.get("width_mm", 0) or 0) <= width + 1e-6
            assert float(placement.get("y_mm", 0) or 0) + float(placement.get("height_mm", 0) or 0) <= height + 1e-6


class _NestingBackend:
    def __init__(self) -> None:
        self.settings = default_laser_quote_settings()
        self.studies: dict[str, dict] = {}

    def laser_quote_settings(self) -> dict:
        return self.settings

    def laser_quote_save_settings(self, settings: dict) -> dict:
        self.settings = dict(settings or {})
        return self.settings

    def laser_sheet_stock_candidates(self, _material: str, _thickness: str) -> list[dict]:
        return [
            {
                "name": "Stock MAT-TEST",
                "source_kind": "stock",
                "source_label": "Stock MAT-TEST",
                "material_id": "MAT-TEST",
                "lote": "LOTE-TEST",
                "local": "RACK A",
                "width_mm": 3000,
                "height_mm": 1500,
                "quantity_available": 3,
                "is_retalho": False,
            }
        ]

    def orc_nesting_studies(self, _number: str) -> dict[str, dict]:
        return self.studies

    def orc_save_nesting_study(self, _number: str, payload: dict) -> dict:
        stored = {**dict(payload or {}), "updated_at": "2026-07-22T10:00:00"}
        self.studies[str(stored.get("group_key", ""))] = stored
        return stored

    def _resolve_file_reference(self, path: str) -> Path:
        return Path(path)


def _wait_for_nesting(app: QApplication, dialog: LaserNestingDialog, timeout_seconds: float = 12.0) -> None:
    deadline = time.monotonic() + timeout_seconds
    while dialog._nesting_thread is not None and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(0.02)
    app.processEvents()
    assert dialog._nesting_thread is None, "A thread de nesting nao terminou dentro do tempo esperado."


def main() -> int:
    app = QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory() as temp_dir:
        drawing_path = Path(temp_dir) / "PECA-A.dxf"
        drawing_path.write_text(TEST_DXF, encoding="ascii")
        rows = [
            {
                "operacao": "Corte Laser",
                "desenho": str(drawing_path),
                "ref_externa": "PECA-A",
                "descricao": "Suporte industrial",
                "qtd": 8,
                "material": "S235JR",
                "espessura": "3",
                "preco": 12.5,
                "total": 100.0,
            }
        ]
        backend = _NestingBackend()
        dialog = LaserNestingDialog(backend, rows, quote_number="ORC-TEST")
        dialog._window_fitted_to_screen = True
        dialog.resize(1280, 720)
        dialog.show()
        app.processEvents()

        flags = dialog.windowFlags()
        assert flags & Qt.WindowMinimizeButtonHint
        assert flags & Qt.WindowMaximizeButtonHint
        assert flags & Qt.WindowCloseButtonHint
        assert dialog.body_scroll.horizontalScrollBar().maximum() == 0
        assert dialog.parts_table.rowCount() == 1
        assert dialog.stock_table.rowCount() == 1
        dialog._select_smallest_fitting_sheet([(3300.0, 380.0)])
        assert float(dict(dialog.sheet_combo.currentData() or {}).get("width_mm", 0) or 0) == 4000.0

        dialog._apply_study_controls(
            {"options": {"allow_mirror": False, "free_angle_rotation": True}}
        )
        assert not dialog.mirror_check.isChecked()
        assert dialog.free_angle_check.isChecked()
        dialog.free_angle_check.setChecked(False)

        for step in range(6):
            dialog._set_wizard_step(step)
            assert dialog._current_wizard_step_index() == step

        dialog._set_wizard_step(0)
        dialog._analyze()
        _wait_for_nesting(app, dialog)
        summary = dict(dialog.result_data.get("summary", {}) or {})
        assert int(summary.get("part_count_requested", 0) or 0) == 8
        assert int(summary.get("part_count_placed", 0) or 0) == 8
        assert backend.studies

        profiles = SheetProfilesDialog([], dialog)
        profile_flags = profiles.windowFlags()
        assert profile_flags & Qt.WindowMinimizeButtonHint
        assert profile_flags & Qt.WindowMaximizeButtonHint
        profiles.close()
        dialog.close()
        app.processEvents()

        oversized = nest_parts(
            rows,
            sheet_width_mm=200.0,
            sheet_height_mm=100.0,
            part_spacing_mm=5.0,
            edge_margin_mm=5.0,
            allow_rotate=True,
            laser_settings=backend.settings,
            shape_aware=False,
        )
        assert int(dict(oversized.get("summary", {}) or {}).get("part_count_placed", 0) or 0) == 0
        assert int(dict(oversized.get("summary", {}) or {}).get("part_count_unplaced", 0) or 0) == 8

        complex_path = Path(temp_dir) / "PECA-COMPLEXA.dxf"
        complex_path.write_text(_complex_dxf(), encoding="ascii")
        complex_rows = [
            {
                "operacao": "Corte Laser",
                "desenho": str(complex_path),
                "ref_externa": "PECA-COMPLEXA",
                "descricao": "Placa perfurada",
                "qtd": 30,
                "material": "S235JR",
                "espessura": "8",
            }
        ]
        started = time.monotonic()
        complex_result = nest_parts(
            complex_rows,
            sheet_width_mm=2000.0,
            sheet_height_mm=1000.0,
            part_spacing_mm=7.0,
            edge_margin_mm=7.0,
            allow_rotate=True,
            allow_mirror=True,
            free_angle_rotation=True,
            laser_settings=backend.settings,
            shape_aware=True,
        )
        elapsed = time.monotonic() - started
        complex_summary = dict(complex_result.get("summary", {}) or {})
        assert elapsed < 20.0, f"Nesting complexo demorou {elapsed:.2f}s."
        assert complex_summary.get("engine_used") == "shape"
        assert int(complex_summary.get("part_count_placed", 0) or 0) == 30
        assert int(complex_summary.get("geometry_solid_overlap_pair_count", 0) or 0) == 0
        assert int(complex_summary.get("geometry_spacing_violation_pair_count", 0) or 0) == 0
        assert int(complex_summary.get("geometry_edge_violation_count", 0) or 0) == 0
        assert float(complex_summary.get("geometry_min_part_distance_mm", 0) or 0) >= 7.0 - 1e-3
        assert float(complex_summary.get("geometry_min_edge_distance_mm", 0) or 0) >= 7.0 - 1e-3
        _assert_placements_within_sheet(complex_result)

    print("laser-nesting-flow-ok pieces=8/8 complex=30 geos<20s margins=yes bounds=yes responsive=1280 native-window-controls=yes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
