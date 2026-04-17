from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

SCRIPTS = [
    "scripts/verify_mysql_schema.py",
    "scripts/verify_data_integrity.py",
    "scripts/verify_purchase_flow.py",
    "scripts/verify_conjuntos_montagem_flow.py",
    "scripts/verify_planning_flow.py",
    "scripts/verify_operator_expedition_flow.py",
    "scripts/verify_shipping_flow.py",
    "scripts/verify_shipping_edge_cases.py",
    "scripts/verify_pulse_flow.py",
    "scripts/verify_opp_dashboard_flow.py",
]


def main() -> int:
    for script in SCRIPTS:
        proc = subprocess.run([sys.executable, script], cwd=ROOT, text=True, capture_output=True)
        if proc.returncode != 0:
            sys.stdout.write(proc.stdout)
            sys.stderr.write(proc.stderr)
            raise SystemExit(proc.returncode)
        if proc.stdout.strip():
            print(proc.stdout.strip())
    print("core-flows-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
