from __future__ import annotations

import sys
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_core.laser.quote_engine import (
    analyze_dxf_geometry,
    default_laser_quote_settings,
    estimate_laser_quote,
    estimate_profile_laser_quote,
)


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
100
20
0
10
100
20
50
10
0
20
50
0
CIRCLE
8
CUT
10
25
20
25
40
5
0
LINE
8
MARK_TEXT
10
10
20
10
11
70
21
10
0
ENDSEC
0
EOF
"""

MIXED_COLOR_DXF = """0
SECTION
2
ENTITIES
0
LWPOLYLINE
8
0
62
7
90
4
70
1
10
0
20
0
10
100
20
0
10
100
20
100
10
0
20
100
0
CIRCLE
8
0
62
7
10
20
20
20
40
5
0
LINE
8
0
62
8
10
10
20
10
11
90
21
10
0
LINE
8
0
62
8
10
90
20
10
11
90
21
90
0
LINE
8
0
62
8
10
90
20
90
11
10
21
90
0
LINE
8
0
62
8
10
10
20
90
11
10
21
10
0
ENDSEC
0
EOF
"""

DASHED_FOLD_DXF = MIXED_COLOR_DXF.replace(
    "62\n8",
    "62\n7\n6\nCENTERX2",
)


def main() -> int:
    with TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "laser_test_piece.dxf"
        path.write_text(TEST_DXF, encoding="utf-8")
        geometry = analyze_dxf_geometry(path, default_laser_quote_settings().get("layer_rules", {}))
        metrics = dict(geometry.get("metrics", {}) or {})
        assert abs(float(metrics.get("cut_length_m", 0)) - 0.3314) < 0.01, metrics
        assert abs(float(metrics.get("mark_length_m", 0)) - 0.06) < 0.01, metrics
        assert int(metrics.get("pierce_count", 0) or 0) == 2, metrics
        quote = estimate_laser_quote(
            {
                "path": str(path),
                "material": "Aco carbono",
                "material_subtype": "S235JR",
                "gas": "Oxigenio",
                "thickness_mm": 8,
                "qtd": 2,
                "include_marking": True,
            }
        )
        pricing = dict(quote.get("pricing", {}) or {})
        line = dict(quote.get("line_suggestion", {}) or {})
        assert float(pricing.get("unit_price", 0) or 0) > 0, pricing
        assert float(pricing.get("total_price", 0) or 0) >= float(pricing.get("unit_price", 0) or 0), pricing
        assert str(line.get("operacao", "") or "").strip().startswith("Corte Laser"), line
        mixed_path = Path(tmp_dir) / "laser_mixed_color_piece.dxf"
        mixed_path.write_text(MIXED_COLOR_DXF, encoding="utf-8")
        mixed_geometry = analyze_dxf_geometry(mixed_path, default_laser_quote_settings().get("layer_rules", {}))
        mixed_metrics = dict(mixed_geometry.get("metrics", {}) or {})
        assert abs(float(mixed_metrics.get("net_area_mm2", 0)) - (10000.0 - (3.141592653589793 * 25.0))) < 1.0, mixed_metrics
        assert abs(float(mixed_metrics.get("mark_length_mm", 0)) - 320.0) < 0.1, mixed_metrics
        assert int(mixed_metrics.get("inner_contours", 0) or 0) == 1, mixed_metrics
        assert any("cor ACI" in str(row) for row in list(mixed_geometry.get("warnings", []) or [])), mixed_geometry
        dashed_path = Path(tmp_dir) / "laser_dashed_fold_piece.dxf"
        dashed_path.write_text(DASHED_FOLD_DXF, encoding="utf-8")
        dashed_geometry = analyze_dxf_geometry(dashed_path, default_laser_quote_settings().get("layer_rules", {}))
        dashed_metrics = dict(dashed_geometry.get("metrics", {}) or {})
        assert abs(float(dashed_metrics.get("net_area_mm2", 0)) - (10000.0 - (3.141592653589793 * 25.0))) < 1.0, dashed_metrics
        assert abs(float(dashed_metrics.get("mark_length_mm", 0)) - 320.0) < 0.1, dashed_metrics
        assert int(dashed_metrics.get("inner_contours", 0) or 0) == 1, dashed_metrics
        assert any("tracejada" in str(row) for row in list(dashed_geometry.get("warnings", []) or [])), dashed_geometry
        profile_payload = {
            "path": str(Path(tmp_dir) / "cantoneira.step"),
            "material": "Aco carbono",
            "material_subtype": "S235JR",
            "gas": "Oxigenio",
            "qtd": 1,
            "profile_family": "Cantoneira",
            "section": "80x80x3",
            "cuts": 4,
            "holes": 0,
            "slots": 2,
            "outer_cuts": 2,
            "cut_length_m_override": 0.72,
            "material_supplied_by_client": True,
            "include_external_profile_cuts": False,
        }
        profile_quote = estimate_profile_laser_quote({**profile_payload, "thickness_mm": 3})
        thick_profile_quote = estimate_profile_laser_quote({**profile_payload, "thickness_mm": 8})
        profile_metrics = dict(profile_quote.get("metrics", {}) or {})
        profile_pricing = dict(profile_quote.get("pricing", {}) or {})
        thick_profile_pricing = dict(thick_profile_quote.get("pricing", {}) or {})
        assert int(profile_metrics.get("cut_event_count", 0) or 0) == 2, profile_metrics
        assert abs(float(profile_metrics.get("cut_length_m", 0) or 0) - 0.72) < 0.001, profile_metrics
        assert float(profile_pricing.get("cut_cost_meter_unit", 0) or 0) > 0, profile_pricing
        assert float(profile_pricing.get("unit_price", 0) or 0) > 0, profile_pricing
        assert float(thick_profile_pricing.get("unit_price", 0) or 0) > float(profile_pricing.get("unit_price", 0) or 0), (
            profile_pricing,
            thick_profile_pricing,
        )
        print(
            "laser-quote-engine-ok",
            metrics.get("cut_length_m"),
            pricing.get("unit_price"),
            pricing.get("total_price"),
            profile_metrics.get("cut_event_count"),
            profile_metrics.get("cut_length_m"),
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
