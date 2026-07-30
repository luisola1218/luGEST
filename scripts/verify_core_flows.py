from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

SCRIPTS = [
    "scripts/verify_mysql_schema.py",
    "scripts/verify_data_integrity.py",
    "scripts/verify_purchase_flow.py",
    "scripts/verify_conjuntos_montagem_flow.py",
    "scripts/verify_fabrication_order_flow.py",
    "scripts/verify_planning_flow.py",
    "scripts/verify_operator_expedition_flow.py",
    "scripts/verify_shipping_flow.py",
    "scripts/verify_shipping_edge_cases.py",
    "scripts/verify_pulse_flow.py",
    "scripts/verify_management_reports.py",
    "scripts/verify_opp_dashboard_flow.py",
    "scripts/verify_opp_client_portfolio.py",
    "scripts/verify_inventory_bulk_delete.py",
    "scripts/verify_material_stock_pdf_filter.py",
    "scripts/verify_product_stock_pdf_filter.py",
    "scripts/verify_laser_nesting_flow.py",
    "scripts/verify_operator_material_session_flow.py",
]


def main() -> int:
    started = time.perf_counter()
    for index, script in enumerate(SCRIPTS, start=1):
        script_started = time.perf_counter()
        print(f"[{index:02d}/{len(SCRIPTS):02d}] {script}", flush=True)
        proc = subprocess.run([sys.executable, script], cwd=ROOT, text=True, capture_output=True)
        if proc.returncode != 0:
            sys.stdout.write(proc.stdout)
            sys.stderr.write(proc.stderr)
            raise SystemExit(proc.returncode)
        if proc.stdout.strip():
            print(proc.stdout.strip())
        elapsed = time.perf_counter() - script_started
        print(f"  OK ({elapsed:.1f}s)", flush=True)
    total = time.perf_counter() - started
    print(f"core-flows-ok ({total:.1f}s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
