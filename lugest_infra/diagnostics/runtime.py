from __future__ import annotations

import faulthandler
import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from types import TracebackType
from typing import TextIO


class RuntimeDiagnostics:
    """Own the application crash log without coupling it to the Qt interface.

    Installation is idempotent.  The log is rotated once it grows beyond the
    configured limit and every write is serialized, which also makes events
    emitted by background workers readable.
    """

    def __init__(
        self,
        *,
        application_dir: str = "LuisGEST",
        filename: str = "lugest_runtime.log",
        max_bytes: int = 2_000_000,
        path: Path | str | None = None,
    ) -> None:
        self.application_dir = str(application_dir or "LuisGEST").strip() or "LuisGEST"
        self.filename = str(filename or "lugest_runtime.log").strip() or "lugest_runtime.log"
        self.max_bytes = max(64_000, int(max_bytes or 2_000_000))
        self._configured_path = Path(path) if path is not None else None
        self.path: Path | None = None
        self._handle: TextIO | None = None
        self._lock = threading.RLock()
        self._installed = False
        self._faulthandler_enabled = False
        self._original_excepthook = None
        self._original_unraisablehook = None
        self._exception_hook = None
        self._unraisable_hook = None

    def log_path(self) -> Path:
        if self._configured_path is not None:
            return self._configured_path
        base_dir = str(os.environ.get("LOCALAPPDATA", "") or "").strip()
        if not base_dir:
            base_dir = str(Path.home() / "AppData" / "Local")
        return Path(base_dir) / self.application_dir / "logs" / self.filename

    def install(self) -> Path | None:
        with self._lock:
            if self._installed:
                return self.path
            try:
                path = self.log_path()
                path.parent.mkdir(parents=True, exist_ok=True)
                self._rotate(path)
                self._handle = path.open("a", encoding="utf-8", buffering=1)
                self.path = path
                self._install_hooks()
                self._installed = True
                return path
            except Exception:
                self._close_handle()
                self.path = None
                return None

    def write_event(self, event: str, detail: str = "") -> None:
        handle = self._handle
        if handle is None:
            return
        timestamp = datetime.now().isoformat(timespec="seconds")
        message = f"[{timestamp}] {str(event or '').strip()}"
        if str(detail or "").strip():
            message += f" | {str(detail or '').strip()}"
        with self._lock:
            try:
                handle.write(message + "\n")
                handle.flush()
            except Exception:
                pass

    def write_exception(
        self,
        event: str,
        exc_type: type[BaseException],
        exc_value: BaseException,
        exc_traceback: TracebackType | None,
    ) -> None:
        self.write_event(event, f"{getattr(exc_type, '__name__', exc_type)}: {exc_value}")
        handle = self._handle
        if handle is None:
            return
        with self._lock:
            try:
                traceback.print_exception(exc_type, exc_value, exc_traceback, file=handle)
                handle.flush()
            except Exception:
                pass

    def close(self) -> None:
        with self._lock:
            if self._exception_hook is not None and sys.excepthook is self._exception_hook:
                sys.excepthook = self._original_excepthook or sys.__excepthook__
            if (
                hasattr(sys, "unraisablehook")
                and self._unraisable_hook is not None
                and sys.unraisablehook is self._unraisable_hook
                and self._original_unraisablehook is not None
            ):
                sys.unraisablehook = self._original_unraisablehook
            if self._faulthandler_enabled:
                try:
                    faulthandler.disable()
                except Exception:
                    pass
            self._faulthandler_enabled = False
            self._exception_hook = None
            self._unraisable_hook = None
            self._installed = False
            self._close_handle()

    def _rotate(self, path: Path) -> None:
        if not path.exists() or path.stat().st_size <= self.max_bytes:
            return
        previous_path = path.with_suffix(".previous.log")
        try:
            previous_path.unlink(missing_ok=True)
            path.replace(previous_path)
        except OSError:
            path.write_text("", encoding="utf-8")

    def _install_hooks(self) -> None:
        handle = self._handle
        if handle is None:
            return
        try:
            faulthandler.enable(file=handle, all_threads=True)
            self._faulthandler_enabled = True
        except Exception:
            pass

        original_excepthook = sys.excepthook
        self._original_excepthook = original_excepthook

        def exception_hook(exc_type, exc_value, exc_traceback) -> None:
            self.write_exception("UNHANDLED_EXCEPTION", exc_type, exc_value, exc_traceback)
            original_excepthook(exc_type, exc_value, exc_traceback)

        self._exception_hook = exception_hook
        sys.excepthook = exception_hook
        if hasattr(sys, "unraisablehook"):
            original_unraisablehook = sys.unraisablehook
            self._original_unraisablehook = original_unraisablehook

            def unraisable_hook(unraisable) -> None:
                self.write_exception(
                    "UNRAISABLE_EXCEPTION",
                    type(unraisable.exc_value),
                    unraisable.exc_value,
                    unraisable.exc_traceback,
                )
                original_unraisablehook(unraisable)

            self._unraisable_hook = unraisable_hook
            sys.unraisablehook = unraisable_hook

    def _close_handle(self) -> None:
        handle = self._handle
        self._handle = None
        if handle is None:
            return
        try:
            handle.flush()
            handle.close()
        except Exception:
            pass
