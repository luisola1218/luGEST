from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from lugest_qt.services.laser_quote_engine import analyze_dxf_geometry, default_laser_quote_settings, estimate_laser_quote


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
        print("laser-quote-engine-ok", metrics.get("cut_length_m"), pricing.get("unit_price"), pricing.get("total_price"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
