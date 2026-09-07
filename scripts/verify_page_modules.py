"""Validate page imports and composition without opening the database."""
from pathlib import Path
import ast
import builtins
import importlib
import os
import symtable
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def global_references(table):
    names = {symbol.get_name() for symbol in table.get_symbols()
             if symbol.is_global() and symbol.is_referenced()}
    for child in table.get_children():
        names.update(global_references(child))
    return names


def main() -> int:
    from lugest_qt.ui.page_registry import PAGE_DEFINITIONS, build_page_factories
    backend, runtime = object(), object()
    factories = build_page_factories(backend, runtime)
    assert not any(name.startswith("lugest_qt.ui.pages.") for name in sys.modules)
    import lugest_qt.ui.main_window
    assert not any(name.startswith("lugest_qt.ui.pages.") for name in sys.modules), "Shell eagerly loads pages"
    from lugest_qt.ui.pages import runtime_pages
    assert "lugest_qt.ui.pages.quotes_page" not in sys.modules
    from PySide6.QtWidgets import QWidget
    for key, (module_name, class_name, dependencies) in PAGE_DEFINITIONS.items():
        module = importlib.import_module(f"lugest_qt.ui.pages.{module_name}")
        page_class = getattr(module, class_name)
        assert issubclass(page_class, QWidget), key
        factory = factories[key]
        assert factory.args[2] == tuple(backend if name == "backend" else runtime for name in dependencies)
        try:
            setattr(module, class_name, lambda *args: args)
            assert factory() == factory.args[2], f"{key}: wrong factory dispatch"
        finally:
            setattr(module, class_name, page_class)
    for name in runtime_pages._EXPORTS:
        assert getattr(runtime_pages, name) is getattr(runtime_pages, name)
    assert runtime_pages.QuotesPage is importlib.import_module("lugest_qt.ui.pages.quotes_page").QuotesPage
    # Catch lost imports in rarely visited callbacks as well as import-time errors.
    modules = set(runtime_pages._EXPORTS.values())
    for module_name in modules:
        module = importlib.import_module(f"lugest_qt.ui.pages.{module_name}")
        source = Path(module.__file__).read_text(encoding="utf-8-sig")
        missing = global_references(symtable.symtable(source, module.__file__, "exec"))
        missing -= set(vars(module)) | set(vars(builtins)) | {"__class__"}
        assert not missing, f"{module_name}: unresolved globals {sorted(missing)}"
        tree = ast.parse(source)
        assert not any(isinstance(node, ast.ImportFrom) and node.module == "runtime_pages" for node in ast.walk(tree))
    print(f"page-modules-ok pages={len(factories)} lazy=yes compatibility=yes callback-globals=yes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
