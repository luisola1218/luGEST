from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def main() -> int:
    parser = argparse.ArgumentParser(description="Exporta SAF-T(PT) de faturação a partir do LuGEST interno.")
    parser.add_argument("--start-date", default="", help="Data inicial YYYY-MM-DD")
    parser.add_argument("--end-date", default="", help="Data final YYYY-MM-DD")
    parser.add_argument("--output", default="", help="Caminho do ficheiro XML de saída")
    args = parser.parse_args()

    backend = LegacyBackend()
    rendered = backend.billing_export_saft_pt(
        start_date=str(args.start_date or "").strip(),
        end_date=str(args.end_date or "").strip(),
        output_path=str(args.output or "").strip(),
    )
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
