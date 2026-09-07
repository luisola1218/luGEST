"""Verify the existing API, extraction dependencies and offline composition."""
from pathlib import Path
import ast
import builtins
from dataclasses import fields
import inspect
import json
import os
import symtable
import sys
import tempfile
import textwrap
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def global_references(table):
    names = {s.get_name() for s in table.get_symbols() if s.is_global() and s.is_referenced()}
    for child in table.get_children():
        names |= global_references(child)
    return names


def main():
    from lugest_qt.services.main_bridge import LegacyBackend
    from lugest_qt.services.legacy_runtime import LegacyRuntime
    contract = json.loads((ROOT / "scripts/contracts/backend_api.json").read_text(encoding="utf-8"))
    for name, expected in contract.items():
        descriptor = inspect.getattr_static(LegacyBackend, name)
        fn = descriptor.fget if isinstance(descriptor, property) else descriptor
        if isinstance(fn, (staticmethod, classmethod)):
            fn = fn.__func__
        node = ast.parse(textwrap.dedent(inspect.getsource(fn))).body[0]
        if name == "__init__":
            # Only an optional keyword was added; the existing no-arg call works.
            inspect.signature(fn).bind(object())
        else:
            assert ast.dump(node.args) == expected["arguments"], name
            assert (ast.dump(node.returns) if node.returns else None) == expected["returns"], name
            assert [ast.dump(d) for d in node.decorator_list] == expected["decorators"], name
    # Includes references inside callbacks and methods not visited by UI smoke tests.
    modules = {inspect.getmodule(base) for base in LegacyBackend.__mro__ if base is not object}
    for module in modules:
        source = Path(module.__file__).read_text(encoding="utf-8-sig")
        missing = global_references(symtable.symtable(source, module.__file__, "exec"))
        missing -= set(vars(module)) | set(vars(builtins)) | {"__class__"}
        assert not missing, f"{module.__name__}: missing globals {sorted(missing)}"
        tree = ast.parse(source)
        if module.__name__.endswith("main_bridge"):
            cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "LegacyBackend")
            assert {n.name for n in cls.body if isinstance(n, ast.FunctionDef)} == {"__init__", "_env_float"}, "Business logic returned to the composition root"
        else:
            assert not any(isinstance(n, ast.ImportFrom) and (n.module or "").endswith("main_bridge") for n in ast.walk(tree)), "Adapter imports the composition root"
    with tempfile.TemporaryDirectory(prefix="lugest-backend-") as temporary:
        values = {field.name: SimpleNamespace() for field in fields(LegacyRuntime)}
        values["desktop_main"] = SimpleNamespace(BASE_DIR=Path(temporary))
        backend = LegacyBackend(runtime=LegacyRuntime(**values))
        assert backend.data is None
        assert backend.base_dir == Path(temporary)
        assert backend.data_cache_generation() == 0
        assert "main" not in sys.modules, "Injected construction loaded the legacy application"
    print(f"backend-modules-ok contracts={len(contract)} callbacks=yes injected-runtime=yes")


if __name__ == "__main__":
    main()
