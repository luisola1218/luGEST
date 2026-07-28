from __future__ import annotations

import argparse
import faulthandler
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QEvent, QObject
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QLineEdit, QMessageBox

from .services.legacy_backend import LegacyBackend
from .services.runtime_service import RuntimeService
from .ui.login_dialog import LoginDialog
from .ui.main_window import MainWindow
from .ui.theme import apply_theme


_DIAGNOSTIC_HANDLE = None
_DIAGNOSTIC_PATH: Path | None = None


class _FullSurfaceComboFilter(QObject):
    """Open selectors from the whole control, including editable combo text areas."""

    def eventFilter(self, watched, event):  # type: ignore[override]
        if event.type() != QEvent.MouseButtonPress:
            return False
        if isinstance(watched, QComboBox):
            if watched.isEnabled():
                watched.setFocus()
                watched.showPopup()
                return True
            return False
        if isinstance(watched, QLineEdit):
            combo = watched.parentWidget()
            if isinstance(combo, QComboBox) and combo.isEnabled():
                combo.showPopup()
                return bool(watched.isReadOnly())
        return False


def _diagnostic_log_path() -> Path:
    base_dir = str(os.environ.get("LOCALAPPDATA", "") or "").strip()
    if not base_dir:
        base_dir = str(Path.home() / "AppData" / "Local")
    return Path(base_dir) / "LuisGEST" / "logs" / "lugest_runtime.log"


def _write_diagnostic_event(event: str, detail: str = "") -> None:
    handle = _DIAGNOSTIC_HANDLE
    if handle is None:
        return
    timestamp = datetime.now().isoformat(timespec="seconds")
    message = f"[{timestamp}] {str(event or '').strip()}"
    if str(detail or "").strip():
        message += f" | {str(detail or '').strip()}"
    try:
        handle.write(message + "\n")
        handle.flush()
    except Exception:
        pass


def _write_diagnostic_exception(event: str, exc_type, exc_value, exc_traceback) -> None:
    _write_diagnostic_event(event, f"{getattr(exc_type, '__name__', exc_type)}: {exc_value}")
    handle = _DIAGNOSTIC_HANDLE
    if handle is None:
        return
    try:
        traceback.print_exception(exc_type, exc_value, exc_traceback, file=handle)
        handle.flush()
    except Exception:
        pass


def _install_runtime_diagnostics() -> Path | None:
    global _DIAGNOSTIC_HANDLE, _DIAGNOSTIC_PATH
    try:
        path = _diagnostic_log_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and path.stat().st_size > 2_000_000:
            previous_path = path.with_suffix(".previous.log")
            try:
                previous_path.unlink(missing_ok=True)
                path.replace(previous_path)
            except OSError:
                path.write_text("", encoding="utf-8")
        _DIAGNOSTIC_HANDLE = path.open("a", encoding="utf-8", buffering=1)
        _DIAGNOSTIC_PATH = path
        try:
            faulthandler.enable(file=_DIAGNOSTIC_HANDLE, all_threads=True)
        except Exception:
            pass

        original_excepthook = sys.excepthook

        def exception_hook(exc_type, exc_value, exc_traceback) -> None:
            _write_diagnostic_exception("UNHANDLED_EXCEPTION", exc_type, exc_value, exc_traceback)
            original_excepthook(exc_type, exc_value, exc_traceback)

        sys.excepthook = exception_hook
        if hasattr(sys, "unraisablehook"):
            original_unraisablehook = sys.unraisablehook

            def unraisable_hook(unraisable) -> None:
                _write_diagnostic_exception(
                    "UNRAISABLE_EXCEPTION",
                    type(unraisable.exc_value),
                    unraisable.exc_value,
                    unraisable.exc_traceback,
                )
                original_unraisablehook(unraisable)

            sys.unraisablehook = unraisable_hook
        return path
    except Exception:
        _DIAGNOSTIC_HANDLE = None
        _DIAGNOSTIC_PATH = None
        return None


def _parse_args(argv: list[str]) -> tuple[argparse.Namespace, list[str]]:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--auto-login-user", default="")
    parser.add_argument("--auto-login-password", default="")
    return parser.parse_known_args(argv[1:])


def _smoke_test_fallback_user(backend: LegacyBackend) -> dict | None:
    data = backend.ensure_data()
    profiles = backend._user_profiles()
    fallback_user = None
    for row in list(data.get("users", []) or []):
        if not isinstance(row, dict):
            continue
        row_username = str(row.get("username", "") or "").strip()
        if not row_username:
            continue
        profile = dict(profiles.get(row_username.lower(), {}) or {})
        if profile and not bool(profile.get("active", True)):
            continue
        merged = dict(row)
        for key in ("posto", "posto_trabalho", "work_center"):
            if str(profile.get("posto", "") or "").strip():
                merged[key] = str(profile.get("posto", "") or "").strip()
        merged["active"] = bool(profile.get("active", True))
        merged["menu_permissions"] = dict(profile.get("menu_permissions", {}) or {})
        fallback_user = merged
        if str(merged.get("role", "") or "").strip().lower() == "admin":
            break
    return fallback_user if isinstance(fallback_user, dict) else None


def _run_smoke_test(app: QApplication, backend: LegacyBackend, runtime_service: RuntimeService, username: str, password: str) -> int:
    cli_credentials_supplied = bool(str(username or "").strip() or str(password or ""))
    login_user = str(username or "").strip()
    login_password = str(password or "")
    if not login_user:
        owner_user_fn = getattr(backend.desktop_main, "trial_owner_username", None)
        if callable(owner_user_fn):
            login_user = str(owner_user_fn() or "").strip()
    if not login_password:
        owner_pass_fn = getattr(backend.desktop_main, "trial_owner_password", None)
        if callable(owner_pass_fn):
            login_password = str(owner_pass_fn() or "")
    if login_user and login_password:
        try:
            backend.authenticate(login_user, login_password)
        except Exception:
            if cli_credentials_supplied:
                raise
            fallback_user = _smoke_test_fallback_user(backend)
            if not isinstance(fallback_user, dict):
                raise
            backend.user = fallback_user
    else:
        fallback_user = _smoke_test_fallback_user(backend)
        if not isinstance(fallback_user, dict):
            raise RuntimeError("Smoke test sem credenciais validas e sem utilizador local ativo para bypass controlado.")
        backend.user = fallback_user
    window = MainWindow(backend, runtime_service)
    for key in ("home", "stock_dashboard", "pulse", "operator", "planning", "avarias", "materials", "products", "clients", "suppliers", "orders", "quotes", "purchase_notes"):
        window.show_page(key)
        app.processEvents()
    original_update_settings = backend.update_settings
    original_update_check = backend.update_check
    try:
        backend.update_settings = lambda: {"auto_check": True}
        backend.update_check = lambda: {}
        for _ in range(5):
            window._auto_check_updates()
            deadline = time.monotonic() + 3.0
            while window._update_check_thread is not None and time.monotonic() < deadline:
                app.processEvents()
                time.sleep(0.01)
            if window._update_check_thread is not None:
                raise RuntimeError("A thread de atualizacao automatica nao terminou durante o smoke test.")
    finally:
        backend.update_settings = original_update_settings
        backend.update_check = original_update_check
    print("qt-smoke-ok")
    window.close()
    return 0


def _show_startup_error(exc: Exception) -> None:
    text = str(exc or "").strip() or "Erro desconhecido no arranque."
    title = "Ligação MySQL" if "mysql" in text.lower() else "Erro de arranque"
    QMessageBox.critical(None, title, text)


def main(argv: list[str] | None = None) -> int:
    _install_runtime_diagnostics()
    args = list(argv if argv is not None else sys.argv)
    cli, qt_args = _parse_args(args)
    _write_diagnostic_event(
        "SESSION_START",
        f"pid={os.getpid()} frozen={bool(getattr(sys, 'frozen', False))} smoke={bool(cli.smoke_test)}",
    )
    if cli.smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication([args[0], *qt_args])
    combo_click_filter = _FullSurfaceComboFilter(app)
    app.installEventFilter(combo_click_filter)
    app._full_surface_combo_filter = combo_click_filter
    _write_diagnostic_event("QAPPLICATION_READY")
    try:
        backend = LegacyBackend()
        _write_diagnostic_event("BACKEND_READY")
        backend.ensure_inventory_scan_codes()
        backend.ensure_pdf_light_theme()
        logo_path = backend.logo_path
        if logo_path is not None:
            try:
                backend.ensure_branding_logo(str(logo_path))
            except Exception:
                pass
        runtime_service = RuntimeService()
        app.setApplicationName("luGEST Qt")
        app.setOrganizationName("luGEST")
        if backend.window_icon_path is not None:
            app.setWindowIcon(QIcon(str(backend.window_icon_path)))
        apply_theme(app, backend.branding)
        app.aboutToQuit.connect(lambda: backend.stop_async_save_worker(timeout_sec=1.0))

        if cli.smoke_test:
            result = _run_smoke_test(app, backend, runtime_service, cli.auto_login_user, cli.auto_login_password)
            _write_diagnostic_event("SMOKE_TEST_END", f"exit_code={result}")
            return result

        login = LoginDialog(backend)
        if login.exec() != QDialog.Accepted:
            _write_diagnostic_event("LOGIN_CANCELLED")
            return 0

        _write_diagnostic_event("LOGIN_ACCEPTED", f"user={str(dict(backend.user or {}).get('username', '') or '')}")
        window = MainWindow(backend, runtime_service)
        window.showMaximized()
        _write_diagnostic_event("MAIN_WINDOW_READY")
        exit_code = app.exec()
        _write_diagnostic_event("SESSION_END", f"exit_code={exit_code}")
        return exit_code
    except Exception as exc:
        _write_diagnostic_exception("STARTUP_EXCEPTION", type(exc), exc, exc.__traceback__)
        if cli.smoke_test:
            raise
        _show_startup_error(exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
