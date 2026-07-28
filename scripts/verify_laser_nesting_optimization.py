from __future__ import annotations

import sys
import tempfile
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.laser_nesting import nest_parts


def _rectangle_dxf(width: float, height: float) -> str:
    return (
        "0\nSECTION\n2\nENTITIES\n0\nLWPOLYLINE\n8\nCUT\n90\n4\n70\n1\n"
        f"10\n0\n20\n0\n10\n{width}\n20\n0\n10\n{width}\n20\n{height}\n10\n0\n20\n{height}\n"
        "0\nENDSEC\n0\nEOF\n"
    )


def _run(rows: list[dict], level: str, seconds: float, cancel_check=None) -> tuple[dict, float]:
    started = time.monotonic()
    result = nest_parts(
        rows,
        sheet_width_mm=4000,
        sheet_height_mm=2000,
        part_spacing_mm=7,
        edge_margin_mm=7,
        allow_rotate=True,
        shape_aware=False,
        optimization_level=level,
        time_limit_s=seconds,
        cancel_check=cancel_check,
    )
    return result, time.monotonic() - started


def main() -> int:
    parts = [
        ("A", 3300, 380, 4),
        ("B", 3300, 350, 4),
        ("C", 450, 1230, 8),
        ("D", 400, 975, 8),
        ("E", 500, 400, 16),
        ("F", 211, 148, 16),
    ]
    with tempfile.TemporaryDirectory() as temp_dir:
        rows: list[dict] = []
        for reference, width, height, quantity in parts:
            path = Path(temp_dir) / f"{reference}.dxf"
            path.write_text(_rectangle_dxf(width, height), encoding="ascii")
            rows.append(
                {
                    "operacao": "Corte Laser",
                    "desenho": str(path),
                    "ref_externa": reference,
                    "descricao": reference,
                    "qtd": quantity,
                    "material": "S235JR",
                    "espessura": "3",
                }
            )

        tap1, _ = _run(rows, "tap1", 0.5)
        tap2, tap2_elapsed = _run(rows, "tap2", 1.2)
        tap1_utils = [float(sheet.get("utilization_net_pct", 0) or 0) for sheet in tap1["sheets"]]
        tap2_utils = [float(sheet.get("utilization_net_pct", 0) or 0) for sheet in tap2["sheets"]]
        assert tap2_utils == sorted(tap2_utils, reverse=True)
        assert "prioritária" in str(tap2["sheets"][0].get("opening_reason", ""))
        assert all("complementar" in str(sheet.get("opening_reason", "")) for sheet in tap2["sheets"][1:])
        assert int(tap2["summary"]["sheet_count"]) <= int(tap1["summary"]["sheet_count"])
        assert tuple(tap2_utils) >= tuple(tap1_utils)
        assert 1.0 <= tap2_elapsed <= 2.5

        cancel_started = time.monotonic()
        cancelled, cancel_elapsed = _run(
            rows,
            "tap3",
            10.0,
            cancel_check=lambda: time.monotonic() - cancel_started >= 0.35,
        )
        assert cancel_elapsed < 2.0
        assert int(cancelled["summary"]["part_count_placed"]) == sum(part[3] for part in parts)
        print(
            "laser-nesting-optimization-ok",
            f"tap1={tap1_utils}",
            f"tap2={tap2_utils}",
            f"timed={tap2_elapsed:.2f}s",
            f"cancel={cancel_elapsed:.2f}s",
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
