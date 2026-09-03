from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_infra.diagnostics import RuntimeDiagnostics


def main() -> int:
    original_excepthook = sys.excepthook
    original_unraisablehook = getattr(sys, "unraisablehook", None)
    with tempfile.TemporaryDirectory(prefix="lugest-diagnostics-") as tmp:
        log_path = Path(tmp) / "runtime.log"
        diagnostics = RuntimeDiagnostics(path=log_path, max_bytes=64_000)
        assert diagnostics.install() == log_path
        assert diagnostics.install() == log_path
        diagnostics.write_event("TEST_EVENT", "detail=ok")
        try:
            raise RuntimeError("diagnostic-test")
        except RuntimeError as exc:
            diagnostics.write_exception("TEST_EXCEPTION", type(exc), exc, exc.__traceback__)
        diagnostics.close()
        text = log_path.read_text(encoding="utf-8")
        assert "TEST_EVENT | detail=ok" in text
        assert "TEST_EXCEPTION | RuntimeError: diagnostic-test" in text
        assert "Traceback (most recent call last)" in text
        assert sys.excepthook is original_excepthook
        if original_unraisablehook is not None:
            assert sys.unraisablehook is original_unraisablehook

    print("runtime-diagnostics-ok events=yes exceptions=yes hooks=restored")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
