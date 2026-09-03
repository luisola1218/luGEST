from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

BOUNDARIES = {
    "lugest_core": ("lugest_qt", "lugest_desktop", "lugest_infra"),
    "lugest_infra": ("lugest_qt", "lugest_desktop"),
}


def _imports(path: Path) -> list[tuple[int, str]]:
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    result: list[tuple[int, str]] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            result.extend((node.lineno, alias.name) for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            result.append((node.lineno, node.module))
    return result


def main() -> int:
    violations: list[str] = []
    counts: dict[str, tuple[int, int]] = {}
    for package, forbidden_prefixes in BOUNDARIES.items():
        package_dir = ROOT / package
        files = sorted(package_dir.rglob("*.py"))
        line_count = 0
        for path in files:
            line_count += len(path.read_text(encoding="utf-8-sig").splitlines())
            for line, imported in _imports(path):
                if any(imported == prefix or imported.startswith(f"{prefix}.") for prefix in forbidden_prefixes):
                    violations.append(f"{path.relative_to(ROOT)}:{line}: import proibido de {imported}")
        counts[package] = (len(files), line_count)

    if violations:
        print("architecture-boundaries-failed")
        for violation in violations:
            print(f"- {violation}")
        return 1

    details = " ".join(f"{name}={files}files/{lines}lines" for name, (files, lines) in counts.items())
    print(f"architecture-boundaries-ok {details}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
