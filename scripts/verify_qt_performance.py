from __future__ import annotations

import os
import sys
import time
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

    from PySide6.QtWidgets import QApplication

    from lugest_qt.services.legacy_backend import LegacyBackend
    from lugest_qt.services.runtime_service import RuntimeService
    from lugest_qt.ui.main_window import MainWindow

    app = QApplication(["verify-qt-performance"])
    backend = LegacyBackend()

    start = time.perf_counter()
    data = backend.ensure_data()
    initial_load_sec = time.perf_counter() - start

    backend.user = {
        **next(
            (row for row in data.get("users", []) if isinstance(row, dict)),
            {"username": "verify", "role": "admin"},
        ),
        "owner_session": True,
    }
    window = MainWindow(backend, RuntimeService())

    page_keys = [
        "home",
        "stock_dashboard",
        "materials",
        "products",
        "clients",
        "suppliers",
        "orders",
        "quotes",
        "purchase_notes",
        "planning",
        "operator",
        "shipping",
        "billing",
        "quality",
        "material_assistant",
        "transportes",
        "opp",
    ]
    page_keys = [key for key in page_keys if key in window.page_factories]

    cold_timings: dict[str, float] = {}
    for key in page_keys:
        start = time.perf_counter()
        window.show_page(key)
        app.processEvents()
        window.refresh_current_page(force=False, background=False)
        app.processEvents()
        cold_timings[key] = time.perf_counter() - start

    warm_timings: dict[str, float] = {}
    for key in page_keys:
        start = time.perf_counter()
        window.show_page(key)
        app.processEvents()
        window.refresh_current_page(force=False, background=False)
        app.processEvents()
        warm_timings[key] = time.perf_counter() - start

    window.close()

    max_cold = max(cold_timings.values() or [0.0])
    max_warm = max(warm_timings.values() or [0.0])
    slow_warm = {key: value for key, value in warm_timings.items() if value > 0.15}
    very_slow_cold = {key: value for key, value in cold_timings.items() if value > 2.50}

    print(f"qt-performance-ok initial_load={initial_load_sec:.3f}s max_cold={max_cold:.3f}s max_warm={max_warm:.3f}s")
    for key, value in sorted(warm_timings.items(), key=lambda item: item[1], reverse=True)[:8]:
        print(f"warm {key} {value:.4f}s")

    if initial_load_sec > 2.0:
        raise RuntimeError(f"Arranque de dados demasiado lento: {initial_load_sec:.3f}s")
    if very_slow_cold:
        print(f"qt-performance-warning cold_pages={very_slow_cold}")
    if slow_warm:
        raise RuntimeError(f"Navegacao quente demasiado lenta: {slow_warm}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
