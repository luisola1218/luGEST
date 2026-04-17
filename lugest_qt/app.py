from __future__ import annotations

import argparse
import os
import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox

from .services.main_bridge import LegacyBackend
from .services.runtime_service import RuntimeService
from .ui.login_dialog import LoginDialog
from .ui.main_window import MainWindow
from .ui.theme import apply_theme


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
    print("qt-smoke-ok")
    window.close()
    return 0


def _show_startup_error(exc: Exception) -> None:
    text = str(exc or "").strip() or "Erro desconhecido no arranque."
    title = "Ligação MySQL" if "mysql" in text.lower() else "Erro de arranque"
    QMessageBox.critical(None, title, text)


def main(argv: list[str] | None = None) -> int:
    args = list(argv if argv is not None else sys.argv)
    cli, qt_args = _parse_args(args)
    if cli.smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication([args[0], *qt_args])
    try:
        backend = LegacyBackend()
        try:
            backend.ensure_branding_logo(str(backend.base_dir / "Logos" / "image (1).jpg"))
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
            return _run_smoke_test(app, backend, runtime_service, cli.auto_login_user, cli.auto_login_password)

        login = LoginDialog(backend)
        if login.exec() != QDialog.Accepted:
            return 0

        window = MainWindow(backend, runtime_service)
        window.showMaximized()
        return app.exec()
    except Exception as exc:
        if cli.smoke_test:
            raise
        _show_startup_error(exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
