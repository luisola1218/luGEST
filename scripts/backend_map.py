"""Locate a backend method and its dependencies without starting the ERP.

Examples:
    python scripts/backend_map.py order_create_or_update
    python scripts/backend_map.py --area materials
    python scripts/backend_map.py --write-index
"""
from __future__ import annotations

import argparse
import ast
import inspect
from pathlib import Path
import sys
import textwrap

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def backend_methods():
    from lugest_qt.services.legacy_backend import LegacyBackend
    methods = {}
    for cls in LegacyBackend.__mro__:
        for name, value in vars(cls).items():
            if name in methods:
                continue
            fn = value.fget if isinstance(value, property) else value
            if isinstance(fn, (staticmethod, classmethod)):
                fn = fn.__func__
            if inspect.isfunction(fn):
                path = Path(inspect.getsourcefile(fn)).resolve()
                methods[name] = (fn, path, inspect.getsourcelines(fn)[1])
    return methods


def write_index(methods):
    target = ROOT / "docs/architecture/BACKEND_METHOD_INDEX.md"
    rows = ["# Indice dos metodos do backend", "",
            "Gerado por `python scripts/backend_map.py --write-index`. Nao editar manualmente.", "",
            "Mostra a implementacao efetiva segundo a ordem de heranca de `LegacyBackend`.", "",
            "| Metodo | Implementacao |", "| --- | --- |"]
    for name, (_fn, path, line) in sorted(methods.items()):
        relative = path.relative_to(ROOT).as_posix()
        rows.append(f"| `{name}` | [{path.stem}:{line}](../../{relative}#L{line}) |")
    target.write_text("\n".join(rows) + "\n", encoding="utf-8")
    print(target)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("method", nargs="?", help="Exact method name or part of its name")
    parser.add_argument("--area", help="Module filename, e.g. materials or users")
    parser.add_argument("--write-index", action="store_true")
    args = parser.parse_args()
    methods = backend_methods()
    if args.write_index:
        write_index(methods)
        return 0
    selected = {name: row for name, row in methods.items()
                if (not args.method or args.method in name) and (not args.area or row[1].stem == args.area)}
    if args.method in selected:
        selected = {args.method: selected[args.method]}
    for name, (fn, path, line) in sorted(selected.items()):
        print(f"{name}{inspect.signature(fn)}\n  {path.relative_to(ROOT)}:{line}")
        if len(selected) == 1:
            tree = ast.parse(textwrap.dedent(inspect.getsource(fn)))
            attributes = sorted({n.attr for n in ast.walk(tree) if isinstance(n, ast.Attribute)
                                 and isinstance(n.value, ast.Name) and n.value.id == "self"})
            for attr in attributes:
                if attr in methods:
                    _other, other_path, other_line = methods[attr]
                    print(f"  -> {attr}: {other_path.relative_to(ROOT)}:{other_line}")
                else:
                    print(f"  -> state/dependency: {attr}")
    return 0 if selected else 1


if __name__ == "__main__":
    raise SystemExit(main())
