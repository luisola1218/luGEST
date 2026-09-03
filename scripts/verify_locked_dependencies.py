from __future__ import annotations

import re
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "requirements-qt.lock.txt"
PIN = re.compile(r"^([A-Za-z0-9_.-]+)==([^\s;]+)$")


def main() -> int:
    expected: dict[str, str] = {}
    for raw in LOCK_PATH.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        match = PIN.fullmatch(line)
        if not match:
            raise RuntimeError(f"Linha não suportada no lock: {line}")
        expected[match.group(1)] = match.group(2)

    mismatches: list[str] = []
    for package, wanted in expected.items():
        try:
            installed = version(package)
        except PackageNotFoundError:
            mismatches.append(f"{package}: em falta (esperado {wanted})")
            continue
        if installed != wanted:
            mismatches.append(f"{package}: {installed} (esperado {wanted})")
    if mismatches:
        print("locked-dependencies-failed")
        for mismatch in mismatches:
            print(f"- {mismatch}")
        return 1
    print(f"locked-dependencies-ok packages={len(expected)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
