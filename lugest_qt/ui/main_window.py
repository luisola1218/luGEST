from __future__ import annotations

from functools import partial
import os
import sys
import time

from PySide6.QtCore import QPoint, QRect, QSize, QProcess, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from .pages.home_page import HomePage
from .pages.billing_page import BillingPage
from .pages.material_assistant_page import MaterialAssistantPage
from .pages.materials_page import MaterialsPage
from .pages.products_page import ProductsPage
from .pages.quality_page import QualityPage
from .pages.runtime_pages import (
    AvariasPage,
    ClientsPage,
    LegacyExpeditionPage as ExpeditionPage,
    LegacyOperatorPage as OperatorPage,
    OppPage,
    LegacyOrdersPage as OrdersPage,
    LegacyPlanningPage as PlanningPage,
    LegacyPurchaseNotesPage as PurchaseNotesPage,
    PulsePage,
    QuotesPage,
    SuppliersPage,
    TransportsPage,
)
from .pages.stock_dashboard_page import StockDashboardPage


class MainWindow(QMainWindow):
    def __init__(self, backend, runtime_service, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.runtime_service = runtime_service
        self._logout_in_progress = False
        self.pages: dict[str, QWidget] = {}
        self.nav_buttons: dict[str, QToolButton] = {}
        self._trial_block_dialog_open = False
        self._save_warning_open = False
        self._closing = False
        self._screen_fitted = False
        self._alerts_cache: dict | None = None
        self._alerts_loaded_at = 0.0
        self._alerts_cache_ttl_sec = 10.0
        self.page_factories = {
            "home": lambda: HomePage(self.backend),
            "stock_dashboard": lambda: StockDashboardPage(self.backend),
            "pulse": lambda: PulsePage(self.runtime_service),
            "operator": lambda: OperatorPage(self.runtime_service, self.backend),
            "planning": lambda: PlanningPage(self.runtime_service, self.backend),
            "avarias": lambda: AvariasPage(self.runtime_service),
            "purchase_notes": lambda: PurchaseNotesPage(self.backend),
            "shipping": lambda: ExpeditionPage(self.backend),
            "billing": lambda: BillingPage(self.backend),
            "materials": lambda: MaterialsPage(self.backend),
            "products": lambda: ProductsPage(self.backend),
            "clients": lambda: ClientsPage(self.backend),
            "suppliers": lambda: SuppliersPage(self.backend),
            "orders": lambda: OrdersPage(self.backend),
            "quotes": lambda: QuotesPage(self.backend),
            "opp": lambda: OppPage(self.backend),
            "material_assistant": lambda: MaterialAssistantPage(self.backend),
            "transportes": lambda: TransportsPage(self.backend),
            "quality": lambda: QualityPage(self.backend),
        }

        self.setWindowTitle("luGEST Qt")
        self.setMinimumSize(1180, 760)
        self.resize(1600, 960)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        shell = QFrame()
        shell.setObjectName("TopBar")
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(16, 10, 16, 10)
        shell_layout.setSpacing(10)

        brand_col = QVBoxLayout()
        brand_col.setContentsMargins(0, 0, 0, 0)
        brand_col.setSpacing(2)
        brand = QLabel("luGEST")
        brand.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f172a;")
        brand_sub = QLabel("ERP industrial")
        brand_sub.setProperty("role", "muted")
        brand_col.addWidget(brand)
        brand_col.addWidget(brand_sub)
        shell_layout.addLayout(brand_col)

        shell_layout.addSpacing(18)

        page_col = QVBoxLayout()
        page_col.setContentsMargins(0, 0, 0, 0)
        page_col.setSpacing(2)
        self.title_label = QLabel("Resumo")
        self.title_label.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f172a;")
        self.subtitle_label = QLabel("Base de trabalho pronta para testes.")
        self.subtitle_label.setProperty("role", "muted")
        page_col.addWidget(self.title_label)
        page_col.addWidget(self.subtitle_label)
        shell_layout.addLayout(page_col, 1)

        right_col = QHBoxLayout()
        right_col.setContentsMargins(0, 0, 0, 0)
        right_col.setSpacing(10)
        self.status_label = QLabel("Pronto")
        self.status_label.setProperty("role", "muted")
        right_col.addWidget(self.status_label, 0, Qt.AlignVCenter)
        user = self.backend.user or {}
        self.user_chip = QLabel(f"{user.get('username', '-')} | {user.get('role', '-')}")
        self.user_chip.setProperty("role", "badge")
        right_col.addWidget(self.user_chip, 0, Qt.AlignVCenter)
        is_admin = str(user.get("role", "") or "").strip().lower() == "admin"
        if is_admin:
            extras_btn = QPushButton("Extras")
            extras_btn.setProperty("variant", "secondary")
            extras_btn.setProperty("toolbarAction", "true")
            extras_btn.setMinimumHeight(38)
            extras_btn.clicked.connect(self._open_admin_extras)
            right_col.addWidget(extras_btn)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("toolbarAction", "true")
        refresh_btn.setMinimumHeight(38)
        refresh_btn.clicked.connect(lambda: self.refresh_current_page(force=True, background=False))
        right_col.addWidget(refresh_btn)
        logout_btn = QPushButton("Logout")
        logout_btn.setProperty("variant", "logout")
        logout_btn.setProperty("toolbarAction", "true")
        logout_btn.setMinimumHeight(38)
        logout_btn.clicked.connect(self._logout)
        right_col.addWidget(logout_btn)
        shell_layout.addLayout(right_col)
        root.addWidget(shell)

        nav_scroll = QScrollArea()
        nav_scroll.setObjectName("NavBar")
        nav_scroll.setWidgetResizable(True)
        nav_scroll.setFrameShape(QFrame.NoFrame)
        nav_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        nav_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        nav_scroll.setFixedHeight(58)
        nav_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        nav_bar = QFrame()
        nav_bar.setObjectName("NavBar")
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(8, 5, 8, 5)
        nav_layout.setSpacing(4)
        nav_content_width = 16
        for key, label in (
            ("stock_dashboard", "Dashboard"),
            ("materials", "Matéria-Prima"),
            ("products", "Produtos"),
            ("clients", "Clientes"),
            ("suppliers", "Fornecedores"),
            ("orders", "Encomendas"),
            ("quotes", "Orçamentos"),
            ("planning", "Planeamento"),
            ("transportes", "Transportes"),
            ("material_assistant", "Assistente MP"),
            ("operator", "Operador"),
            ("opp", "OPP"),
            ("shipping", "Expedição"),
            ("billing", "Faturação"),
            ("purchase_notes", "Notas Encomenda"),
            ("quality", "Qualidade"),
            ("pulse", "Pulse"),
            ("avarias", "Avarias"),
            ("home", "Resumo"),
        ):
            button = QToolButton()
            button.setText(label)
            button.setToolTip(label)
            button.setCheckable(True)
            button.setProperty("nav", "true")
            button.setToolButtonStyle(Qt.ToolButtonTextOnly)
            text_width = button.fontMetrics().horizontalAdvance(label)
            button_width = max(58, text_width + 26)
            button.setMinimumWidth(button_width)
            button.setFixedHeight(34)
            button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            button.clicked.connect(partial(self.show_page, key))
            nav_layout.addWidget(button)
            self.nav_buttons[key] = button
            nav_content_width += button_width + nav_layout.spacing()
        nav_bar.setMinimumSize(nav_content_width + 10, 46)
        nav_bar.resize(nav_content_width + 10, 46)
        nav_scroll.setWidget(nav_bar)
        root.addWidget(nav_scroll)

        self.alert_bar = QFrame()
        self.alert_bar.setObjectName("Card")
        self.alert_bar.hide()
        alert_layout = QHBoxLayout(self.alert_bar)
        alert_layout.setContentsMargins(16, 10, 16, 10)
        alert_layout.setSpacing(10)
        self.alert_label = QLabel("")
        self.alert_label.setWordWrap(True)
        self.alert_label.setProperty("role", "alert_text")
        alert_layout.addWidget(self.alert_label, 1)
        open_alerts_btn = QPushButton("Ver Avarias")
        open_alerts_btn.setProperty("variant", "secondary")
        open_alerts_btn.setProperty("toolbarAction", "true")
        open_alerts_btn.setMinimumHeight(36)
        open_alerts_btn.clicked.connect(lambda: self.show_page("avarias"))
        alert_layout.addWidget(open_alerts_btn)
        root.addWidget(self.alert_bar)

        self.stack = QStackedWidget()
        root.addWidget(self.stack, 1)

        self.auto_refresh = QTimer(self)
        self.auto_refresh.setInterval(45000)
        self.auto_refresh.timeout.connect(lambda: self.refresh_current_page(force=False, background=True))
        self.auto_refresh.start()

        self.save_monitor = QTimer(self)
        self.save_monitor.setInterval(1200)
        self.save_monitor.timeout.connect(self._poll_save_runtime_state)
        self.save_monitor.start()

        self._apply_navigation_permissions()
        landing_key = self._default_page_key()
        if landing_key:
            self.show_page(landing_key)
        else:
            self.status_label.setText("Sem menus disponiveis")
        QTimer.singleShot(3500, self._auto_check_updates)

    def _fit_to_available_screen(self) -> None:
        screen = self.screen()
        if screen is None and self.windowHandle() is not None:
            screen = self.windowHandle().screen()
        if screen is None:
            return
        safe = screen.availableGeometry().adjusted(8, 8, -8, -8)
        if safe.width() <= 0 or safe.height() <= 0:
            safe = screen.availableGeometry()
        target_width = min(max(1180, int(safe.width() * 0.985)), safe.width())
        target_height = min(max(760, int(safe.height() * 0.985)), safe.height())
        pos_x = safe.x() + max(0, int((safe.width() - target_width) / 2))
        pos_y = safe.y() + max(0, int((safe.height() - target_height) / 2))
        self.setGeometry(QRect(QPoint(pos_x, pos_y), QSize(target_width, target_height)))

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        if not self._screen_fitted:
            self._screen_fitted = True
            QTimer.singleShot(0, self._fit_to_available_screen)

    def _auto_check_updates(self) -> None:
        try:
            settings = dict(self.backend.update_settings() or {})
            if not bool(settings.get("auto_check", False)):
                return
            result = dict(self.backend.update_check() or {})
        except Exception:
            return
        if not bool(result.get("update_available", False)):
            return
        self.status_label.setText(f"Atualizacao disponivel: {result.get('latest_version', '-')}")
        if QMessageBox.question(
            self,
            "Atualizacoes",
            f"Existe uma nova versao disponivel: {result.get('latest_version', '-')}\n\nQueres abrir o painel de atualizacoes?",
        ) == QMessageBox.Yes:
            self._open_updates_dialog()

    def _ensure_page(self, key: str) -> QWidget:
        if key in self.pages:
            return self.pages[key]
        page = self.page_factories[key]()
        self.pages[key] = page
        self.stack.addWidget(page)
        return page

    def show_page(self, key: str) -> None:
        if key not in self._allowed_page_keys():
            fallback = self._default_page_key()
            if not fallback:
                return
            key = fallback
        page = self._ensure_page(key)
        self.stack.setCurrentWidget(page)
        for nav_key, button in self.nav_buttons.items():
            button.setChecked(nav_key == key)
        self.title_label.setText(str(getattr(page, "page_title", "luGEST Qt")))
        self.subtitle_label.setText(str(getattr(page, "page_subtitle", "")))
        self.refresh_current_page(force=False, background=False)

    def refresh_current_page(self, force: bool = False, background: bool = False) -> None:
        if not self._ensure_trial_runtime_access():
            return
        current = self.stack.currentWidget()
        if current is None:
            return
        if force:
            invalidator = getattr(self.runtime_service, "invalidate_cache", None)
            if callable(invalidator):
                invalidator()
        if background and not bool(getattr(current, "allow_auto_timer_refresh", False)):
            return
        can_auto_refresh = getattr(current, "can_auto_refresh", None)
        if background and callable(can_auto_refresh) and not bool(can_auto_refresh()):
            self.status_label.setText("Edicao ativa")
            return
        if bool(getattr(current, "uses_backend_reload", False)):
            if not self._prepare_backend_reload(force=force):
                return
            self.backend.reload(force=force)
        refresh = getattr(current, "refresh", None)
        if callable(refresh):
            refresh()
        self._refresh_global_alerts(force=force)
        self._poll_save_runtime_state()
        self.status_label.setText("Atualizado agora")

    def _save_runtime_state(self) -> dict:
        getter = getattr(self.backend, "save_runtime_state", None)
        if not callable(getter):
            return {}
        try:
            payload = dict(getter() or {})
        except Exception:
            payload = {}
        payload.setdefault("pending", False)
        payload.setdefault("in_progress", False)
        payload.setdefault("async_enabled", False)
        payload.setdefault("last_error", "")
        return payload

    def _consume_save_error(self) -> str:
        getter = getattr(self.backend, "consume_async_save_error", None)
        if not callable(getter):
            return ""
        try:
            return str(getter() or "").strip()
        except Exception:
            return ""

    def _show_save_error(self, err: str, *, context: str = "") -> None:
        message = str(err or "").strip()
        if not message:
            return
        self.status_label.setText("Erro a gravar")
        if self._save_warning_open:
            return
        self._save_warning_open = True
        try:
            prefix = f"Falha ao guardar {context}:\n" if str(context or "").strip() else "Falha ao guardar na base de dados:\n"
            QMessageBox.warning(self, "Aviso MySQL", f"{prefix}{message}")
        finally:
            self._save_warning_open = False

    def _poll_save_runtime_state(self) -> None:
        err = self._consume_save_error()
        if err:
            self._show_save_error(err)
            return
        state = self._save_runtime_state()
        if bool(state.get("in_progress", False)):
            self.status_label.setText("A guardar...")
        elif bool(state.get("pending", False)):
            self.status_label.setText("Gravacao pendente")

    def _finalize_save_pipeline(self, *, context: str, timeout_sec: float, interactive: bool) -> bool:
        state_before = self._save_runtime_state()
        had_pending = bool(state_before.get("pending", False) or state_before.get("in_progress", False))

        flusher = getattr(self.backend, "flush_pending_save", None)
        flushed = False
        if callable(flusher):
            try:
                flushed = bool(flusher(force=True))
            except Exception as exc:
                if interactive:
                    QMessageBox.warning(self, "Aviso MySQL", f"Falha ao preparar a gravacao {context}:\n{exc}")
                self.status_label.setText("Erro a gravar")
                return False

        state_after_flush = self._save_runtime_state()
        needs_wait = had_pending or flushed or bool(state_after_flush.get("pending", False) or state_after_flush.get("in_progress", False))
        if needs_wait:
            self.status_label.setText("A concluir gravacao...")
            drainer = getattr(self.backend, "drain_async_saves", None)
            drained = True
            if callable(drainer):
                try:
                    drained = bool(drainer(timeout_sec=timeout_sec))
                except Exception:
                    drained = False
            if not drained:
                self.status_label.setText("Gravacao pendente")
                if interactive:
                    QMessageBox.warning(
                        self,
                        "Aviso MySQL",
                        f"A aplicacao ainda estava a gravar {context}. Espera alguns segundos e tenta novamente.",
                    )
                return False

        err = self._consume_save_error()
        if err:
            if interactive:
                self._show_save_error(err, context=context)
            else:
                self.status_label.setText("Erro a gravar")
            return False
        return True

    def _prepare_backend_reload(self, force: bool) -> bool:
        ok = self._finalize_save_pipeline(
            context="antes de atualizar os dados",
            timeout_sec=12.0 if force else 4.0,
            interactive=force,
        )
        if not ok and not force:
            self.status_label.setText("A guardar...")
        return ok

    def _ensure_trial_runtime_access(self) -> bool:
        getter = getattr(self.backend, "trial_status", None)
        if not callable(getter):
            return True
        try:
            status = dict(getter(force=False) or {})
        except Exception:
            return True
        if not bool(status.get("blocking", False)):
            self._trial_block_dialog_open = False
            return True
        if bool(dict(self.backend.user or {}).get("owner_session", False)):
            self._trial_block_dialog_open = False
            return True
        if not self._trial_block_dialog_open:
            self._trial_block_dialog_open = True
            QMessageBox.critical(
                self,
                "Licenciamento",
                str(status.get("message", "") or "Licenca bloqueada."),
            )
        self.close()
        return False

    def _allowed_page_keys(self) -> set[str]:
        return set(self._allowed_page_sequence())

    def _allowed_page_sequence(self) -> list[str]:
        navigation_order = list(self.nav_buttons.keys())
        getter = getattr(self.backend, "allowed_pages_for_user", None)
        if callable(getter):
            allowed = {
                str(key or "").strip()
                for key in getter()
                if str(key or "").strip()
            }
            return [key for key in navigation_order if key in allowed]
        return navigation_order

    def _default_page_key(self) -> str:
        allowed = self._allowed_page_sequence()
        return allowed[0] if allowed else ""

    def _apply_navigation_permissions(self) -> None:
        allowed = self._allowed_page_keys()
        for key, button in self.nav_buttons.items():
            visible = key in allowed
            button.setVisible(visible)
            button.setEnabled(visible)
        current = self.stack.currentWidget()
        if current is not None:
            current_key = next((key for key, page in self.pages.items() if page is current), "")
            if current_key and current_key not in allowed:
                fallback = self._default_page_key()
                if fallback:
                    self.show_page(fallback)

    def _refresh_global_alerts(self, *, force: bool = False) -> None:
        if (
            not force
            and isinstance(self._alerts_cache, dict)
            and self._alerts_loaded_at > 0
            and (time.time() - self._alerts_loaded_at) <= self._alerts_cache_ttl_sec
        ):
            payload = dict(self._alerts_cache)
        else:
            try:
                payload = self.runtime_service.alerts()
            except Exception:
                self.alert_bar.hide()
                return
            self._alerts_cache = dict(payload or {})
            self._alerts_loaded_at = time.time()
        items = list(payload.get("items", []) or [])
        if not items:
            self.alert_bar.hide()
            return
        banner = str(payload.get("banner", "") or "").strip()
        if not banner:
            banner = " | ".join([str(item.get("text", "") or "").strip() for item in items[:2] if str(item.get("text", "") or "").strip()])
        count = int(payload.get("count", 0) or len(items))
        self.alert_label.setText(f"{count} alerta(s) ativos | {banner}")
        self.alert_bar.setProperty("tone", "danger")
        style = self.alert_bar.style()
        if style is not None:
            style.unpolish(self.alert_bar)
            style.polish(self.alert_bar)
        self.alert_bar.show()

    def closeEvent(self, event) -> None:
        if self._closing:
            super().closeEvent(event)
            return
        self._closing = True
        try:
            self.auto_refresh.stop()
            self.save_monitor.stop()
            if not self._finalize_save_pipeline(context="antes de fechar a aplicacao", timeout_sec=20.0, interactive=True):
                self.auto_refresh.start()
                self.save_monitor.start()
                self._closing = False
                event.ignore()
                return
            stopper = getattr(self.backend, "stop_async_save_worker", None)
            if callable(stopper):
                stopper(timeout_sec=1.0)
            super().closeEvent(event)
        except Exception as exc:
            self.auto_refresh.start()
            self.save_monitor.start()
            self._closing = False
            QMessageBox.warning(self, "Aviso MySQL", f"Falha ao fechar a aplicacao em seguranca:\n{exc}")
            event.ignore()

    def _logout(self) -> None:
        if self._logout_in_progress:
            return
        if QMessageBox.question(self, "Logout", "Terminar a sessão atual e voltar ao login?") != QMessageBox.Yes:
            return
        self._logout_in_progress = True
        try:
            self.auto_refresh.stop()
            self.save_monitor.stop()
            if not self._finalize_save_pipeline(context="antes de terminar a sessao", timeout_sec=15.0, interactive=True):
                self.auto_refresh.start()
                self.save_monitor.start()
                self._logout_in_progress = False
                return
            stopper = getattr(self.backend, "stop_async_save_worker", None)
            if callable(stopper):
                stopper(timeout_sec=1.0)
            self.backend.user = {}
            if getattr(sys, "frozen", False):
                app_dir = os.path.dirname(os.path.abspath(sys.executable))
                launcher_vbs = os.path.join(app_dir, "Arrancar LuisGEST Desktop.vbs")
                launcher_bat = os.path.join(app_dir, "Arrancar LuisGEST Desktop.bat")
                launcher_ps1 = os.path.join(app_dir, "Arrancar LuisGEST Desktop.ps1")
                if os.path.exists(launcher_vbs):
                    restart_cmd = (
                        f'ping 127.0.0.1 -n 3 >nul & '
                        f'start "" wscript.exe "{launcher_vbs}" "{app_dir}"'
                    )
                    started = QProcess.startDetached("cmd.exe", ["/c", restart_cmd], app_dir)
                elif os.path.exists(launcher_bat):
                    restart_cmd = f'ping 127.0.0.1 -n 3 >nul & start "" "{launcher_bat}"'
                    started = QProcess.startDetached("cmd.exe", ["/c", restart_cmd], app_dir)
                elif os.path.exists(launcher_ps1):
                    restart_cmd = (
                        'ping 127.0.0.1 -n 3 >nul & '
                        f'start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{launcher_ps1}" -InstallDir "{app_dir}"'
                    )
                    started = QProcess.startDetached("cmd.exe", ["/c", restart_cmd], app_dir)
                else:
                    restart_cmd = f'ping 127.0.0.1 -n 3 >nul & start "" "{sys.executable}"'
                    started = QProcess.startDetached("cmd.exe", ["/c", restart_cmd], app_dir)
            else:
                started = QProcess.startDetached(sys.executable, list(sys.argv), os.getcwd())
            if not started:
                raise RuntimeError("Não foi possível reabrir o ecrã de login.")
            self.close()
        except Exception as exc:
            self.auto_refresh.start()
            self.save_monitor.start()
            self._logout_in_progress = False
            QMessageBox.warning(self, "Logout", str(exc))

    def _open_updates_dialog(self) -> None:
        try:
            settings = dict(self.backend.update_settings() or {})
        except Exception as exc:
            QMessageBox.critical(self, "Atualizacoes", str(exc))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Atualizacoes LuisGEST")
        dialog.resize(720, 460)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Atualizacoes do software")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(title)
        hint = QLabel("Usa um manifest latest.json em pasta partilhada, GitHub privado ou servidor proprio. O instalador faz backup antes de trocar os ficheiros.")
        hint.setWordWrap(True)
        hint.setProperty("role", "muted")
        layout.addWidget(hint)

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        version_label = QLabel(str(settings.get("current_version", "-") or "-"))
        manifest_edit = QLineEdit(str(settings.get("manifest_url", "") or ""))
        manifest_edit.setPlaceholderText(r"..\Atualizacoes\latest.json ou https://servidor/latest.json")
        token_edit = QLineEdit(str(settings.get("github_token", "") or ""))
        token_edit.setEchoMode(QLineEdit.Password)
        token_edit.setPlaceholderText("Opcional para GitHub privado")
        auto_check_box = QCheckBox("Verificar automaticamente no arranque")
        auto_check_box.setChecked(bool(settings.get("auto_check", False)))
        form.addRow("Versao instalada", version_label)
        form.addRow("Manifest", manifest_edit)
        form.addRow("Token GitHub", token_edit)
        form.addRow("", auto_check_box)
        layout.addLayout(form)

        status = QTextEdit()
        status.setReadOnly(True)
        status.setMinimumHeight(150)
        status.setPlainText("Pronto para verificar atualizacoes.")
        layout.addWidget(status, 1)

        def save_settings() -> dict:
            payload = {
                "manifest_url": manifest_edit.text().strip(),
                "github_token": token_edit.text().strip(),
                "auto_check": auto_check_box.isChecked(),
            }
            saved = dict(self.backend.update_save_settings(payload) or {})
            version_label.setText(str(saved.get("current_version", "-") or "-"))
            return saved

        def check_updates() -> dict | None:
            try:
                save_settings()
                result = dict(self.backend.update_check() or {})
            except Exception as exc:
                status.setPlainText(f"Erro ao verificar:\n{exc}")
                return None
            lines = [
                f"Versao instalada: {result.get('current_version', '-')}",
                f"Versao disponivel: {result.get('latest_version', '-')}",
                f"Canal: {result.get('channel', '-')}",
                f"Atualizacao disponivel: {'Sim' if result.get('update_available') else 'Nao'}",
                "",
                f"Pacote: {result.get('package_url', '-')}",
                f"SHA256: {result.get('sha256', '-') or '-'}",
                "",
                str(result.get("notes", "") or ""),
            ]
            status.setPlainText("\n".join(lines).strip())
            return result

        def start_update() -> None:
            result = check_updates()
            if not result:
                return
            if not bool(result.get("update_available", False)):
                QMessageBox.information(dialog, "Atualizacoes", "Nao existem atualizacoes novas.")
                return
            if QMessageBox.question(
                dialog,
                "Atualizacoes",
                "O atualizador vai abrir numa janela externa. Fecha o LuisGEST quando ele pedir para poder substituir os ficheiros.\n\nContinuar?",
            ) != QMessageBox.Yes:
                return
            try:
                self.backend.update_start_installer()
            except Exception as exc:
                QMessageBox.critical(dialog, "Atualizacoes", str(exc))
                return
            QMessageBox.information(dialog, "Atualizacoes", "Atualizador iniciado. Segue as instrucoes da janela externa.")

        buttons = QHBoxLayout()
        save_btn = QPushButton("Guardar")
        check_btn = QPushButton("Verificar")
        update_btn = QPushButton("Atualizar agora")
        update_btn.setProperty("variant", "primary")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        buttons.addStretch(1)
        for button in (save_btn, check_btn, update_btn, close_btn):
            buttons.addWidget(button)
        layout.addLayout(buttons)
        save_btn.clicked.connect(lambda: (save_settings(), status.setPlainText("Configuracao guardada.")))
        check_btn.clicked.connect(check_updates)
        update_btn.clicked.connect(start_update)
        close_btn.clicked.connect(dialog.accept)
        dialog.exec()

    def _open_admin_extras(self) -> None:
        user = dict(self.backend.user or {})
        if str(user.get("role", "") or "").strip().lower() != "admin":
            QMessageBox.information(self, "Extras", "Apenas o admin pode abrir os extras.")
            return
        getter = getattr(self.backend, "ui_options", None)
        setter = getattr(self.backend, "set_ui_option", None)
        options = dict(getter() or {}) if callable(getter) else {}
        supervisor_password_set = bool(options.get("operator_supervisor_password_set", False))
        dialog = QDialog(self)
        dialog.setWindowTitle("Extras Admin")
        dialog.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setWindowFlag(Qt.WindowSystemMenuHint, True)
        dialog.setSizeGripEnabled(True)
        available = self.screen().availableGeometry() if self.screen() is not None else QRect()
        min_w = 900
        min_h = 640
        target_w = 1080
        target_h = 760
        if not available.isNull():
            target_w = max(min_w, min(1080, available.width() - 48))
            target_h = max(min_h, min(760, available.height() - 48))
        dialog.setMinimumSize(min_w, min_h)
        dialog.resize(target_w, target_h)
        dialog.setStyleSheet(
            """
            QDialog { font-size: 12px; }
            QLabel { font-size: 12px; }
            QLineEdit, QComboBox, QTableWidget { font-size: 12px; }
            QCheckBox { font-size: 12px; }
            QPushButton { font-size: 12px; min-height: 34px; }
            """
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)
        intro = QLabel("Configuracoes adicionais do desktop Qt e autorizacoes operacionais.")
        intro.setWordWrap(True)
        layout.addWidget(intro)
        tools_row = QHBoxLayout()
        company_btn = QPushButton("Empresa / PDFs")
        company_btn.setProperty("variant", "secondary")
        tools_row.addWidget(company_btn)
        workcenters_btn = QPushButton("Postos Trabalho")
        workcenters_btn.setProperty("variant", "secondary")
        tools_row.addWidget(workcenters_btn)
        operations_btn = QPushButton("Operacoes")
        operations_btn.setProperty("variant", "secondary")
        tools_row.addWidget(operations_btn)
        updates_btn = QPushButton("Atualizacoes")
        updates_btn.setProperty("variant", "secondary")
        tools_row.addWidget(updates_btn)
        trial_manage_allowed = bool(getattr(self.backend, "is_owner_session", lambda: False)())
        trial_btn = None
        if trial_manage_allowed:
            trial_btn = QPushButton("Trial / Licenca")
            trial_btn.setProperty("variant", "secondary")
            trial_btn.setToolTip("Gestao de trial/licenca autorizada pela sessao OWNER.")
            tools_row.addWidget(trial_btn)
        tools_row.addStretch(1)
        layout.addLayout(tools_row)
        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        supervisor_edit = QLineEdit()
        supervisor_edit.setEchoMode(QLineEdit.Password)
        supervisor_edit.setPlaceholderText(
            "Deixa vazio para manter a password atual"
            if supervisor_password_set
            else "Define uma password forte para Dar Baixa"
        )
        show_client_box = QCheckBox("Mostrar nome do cliente no menu Operador")
        show_client_box.setChecked(bool(options.get("operator_show_client_name", True)))
        form.addRow("Password supervisor", supervisor_edit)
        form.addRow("", show_client_box)
        layout.addLayout(form)

        users_title = QLabel("Utilizadores e permissoes")
        users_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(users_title)
        users_host = QWidget()
        users_layout = QHBoxLayout(users_host)
        users_layout.setContentsMargins(0, 0, 0, 0)
        users_layout.setSpacing(12)

        users_table = QTableWidget(0, 4)
        users_table.setHorizontalHeaderLabels(["Utilizador", "Role", "Posto", "Ativo"])
        users_table.verticalHeader().setVisible(False)
        users_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        users_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        users_table.setSelectionMode(QAbstractItemView.SingleSelection)
        users_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        users_table.setWordWrap(False)
        users_table.horizontalHeader().setStretchLastSection(False)
        users_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        users_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        users_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        users_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        users_table.setAlternatingRowColors(True)
        users_layout.addWidget(users_table, 5)

        form_host = QWidget()
        form_host_layout = QVBoxLayout(form_host)
        form_host_layout.setContentsMargins(0, 0, 0, 0)
        form_host_layout.setSpacing(10)
        user_form = QFormLayout()
        user_form.setHorizontalSpacing(12)
        user_form.setVerticalSpacing(8)
        username_edit = QLineEdit()
        password_user_edit = QLineEdit()
        password_user_edit.setEchoMode(QLineEdit.Password)
        password_user_edit.setPlaceholderText("Obrigatoria para novo utilizador")
        password_toggle = QPushButton("Mostrar")
        password_toggle.setProperty("variant", "secondary")
        password_toggle.setCheckable(True)
        password_row = QWidget()
        password_row_layout = QHBoxLayout(password_row)
        password_row_layout.setContentsMargins(0, 0, 0, 0)
        password_row_layout.setSpacing(8)
        password_row_layout.addWidget(password_user_edit, 1)
        password_row_layout.addWidget(password_toggle, 0)
        role_combo = QComboBox()
        role_combo.addItems(list(getattr(self.backend, "available_roles", lambda: ["Admin", "Producao", "Qualidade", "Planeamento", "Orcamentista", "Operador"])() or []))
        posto_combo = QComboBox()
        posto_combo.setEditable(True)
        posto_combo.addItem("")
        for posto in list(getattr(self.backend, "available_postos", lambda: ["Geral"])() or []):
            if posto_combo.findText(str(posto)) < 0:
                posto_combo.addItem(str(posto))
        active_box = QCheckBox("Utilizador ativo")
        user_form.addRow("Utilizador", username_edit)
        user_form.addRow("Password", password_row)
        user_form.addRow("Role", role_combo)
        user_form.addRow("Posto", posto_combo)
        user_form.addRow("", active_box)
        form_host_layout.addLayout(user_form)
        password_note = QLabel("Deixa a password em branco para manter a atual quando estiveres a editar.")
        password_note.setProperty("role", "muted")
        password_note.setWordWrap(True)
        form_host_layout.addWidget(password_note)
        session_passwords: dict[str, str] = {}
        current_session_user = str((self.backend.user or {}).get("username", "") or "").strip().lower()
        current_session_password = str((self.backend.user or {}).get("_session_password", "") or "").strip()
        if current_session_user and current_session_password:
            session_passwords[current_session_user] = current_session_password
        password_preview_edit = QLineEdit()
        password_preview_edit.setReadOnly(True)
        password_preview_edit.setEchoMode(QLineEdit.Password)
        password_preview_edit.setPlaceholderText("Só aparece aqui a última password definida nesta sessão.")
        preview_toggle = QPushButton("Ver")
        preview_toggle.setProperty("variant", "secondary")
        preview_toggle.setCheckable(True)
        preview_row = QWidget()
        preview_row_layout = QHBoxLayout(preview_row)
        preview_row_layout.setContentsMargins(0, 0, 0, 0)
        preview_row_layout.setSpacing(8)
        preview_row_layout.addWidget(password_preview_edit, 1)
        preview_row_layout.addWidget(preview_toggle, 0)
        preview_note = QLabel(
            "As passwords existentes não podem ser reveladas depois de gravadas, porque ficam guardadas em hash. "
            "Aqui só aparece a última password que definiste nesta sessão de administração."
        )
        preview_note.setProperty("role", "muted")
        preview_note.setWordWrap(True)
        form_host_layout.addWidget(preview_row)
        form_host_layout.addWidget(preview_note)

        perms_label = QLabel("Menus permitidos")
        perms_label.setStyleSheet("font-size: 14px; font-weight: 700; color: #0f172a;")
        form_host_layout.addWidget(perms_label)
        perms_grid = QGridLayout()
        perms_grid.setContentsMargins(0, 0, 0, 0)
        perms_grid.setHorizontalSpacing(16)
        perms_grid.setVerticalSpacing(8)
        permission_checks: dict[str, QCheckBox] = {}
        menu_defs = list(getattr(self.backend, "available_menu_pages", lambda: [])() or [])
        for index, item in enumerate(menu_defs):
            key = str(item.get("key", "") or "").strip()
            label = str(item.get("label", key) or key).strip()
            if not key:
                continue
            check = QCheckBox(label)
            permission_checks[key] = check
            perms_grid.addWidget(check, index // 2, index % 2)
        form_host_layout.addLayout(perms_grid)

        user_actions = QHBoxLayout()
        new_user_btn = QPushButton("Novo")
        new_user_btn.setProperty("variant", "secondary")
        save_user_btn = QPushButton("Guardar utilizador")
        remove_user_btn = QPushButton("Remover utilizador")
        remove_user_btn.setProperty("variant", "danger")
        user_actions.addWidget(new_user_btn)
        user_actions.addWidget(save_user_btn)
        user_actions.addWidget(remove_user_btn)
        user_actions.addStretch(1)
        form_host_layout.addLayout(user_actions)
        users_layout.addWidget(form_host, 6)
        layout.addWidget(users_host, 1)

        current_username = {"value": ""}

        def sync_password_echo() -> None:
            visible = bool(password_toggle.isChecked())
            password_user_edit.setEchoMode(QLineEdit.Normal if visible else QLineEdit.Password)
            password_toggle.setText("Ocultar" if visible else "Mostrar")

        def sync_preview_echo() -> None:
            visible = bool(preview_toggle.isChecked())
            password_preview_edit.setEchoMode(QLineEdit.Normal if visible else QLineEdit.Password)
            preview_toggle.setText("Ocultar" if visible else "Ver")

        def update_password_preview(username: str = "") -> None:
            user_key = str(username or "").strip().lower()
            preview_value = str(session_passwords.get(user_key, "") or "")
            password_preview_edit.setText(preview_value)
            preview_toggle.setEnabled(bool(preview_value))
            if not preview_value:
                preview_toggle.setChecked(False)
            sync_preview_echo()

        def refresh_posto_options(selected_text: str = "") -> None:
            current_text = str(selected_text or posto_combo.currentText() or "").strip()
            posto_combo.blockSignals(True)
            posto_combo.clear()
            posto_combo.addItem("")
            for posto in list(getattr(self.backend, "available_postos", lambda: ["Geral"])() or []):
                value = str(posto or "").strip()
                if value and posto_combo.findText(value) < 0:
                    posto_combo.addItem(value)
            if current_text:
                if posto_combo.findText(current_text) < 0:
                    posto_combo.addItem(current_text)
                posto_combo.setCurrentText(current_text)
            else:
                posto_combo.setCurrentIndex(0)
            posto_combo.blockSignals(False)

        def open_company_dialog() -> None:
            getter_brand = getattr(self.backend, "branding_settings", None)
            saver_brand = getattr(self.backend, "save_branding_settings", None)
            if not callable(getter_brand) or not callable(saver_brand):
                QMessageBox.warning(dialog, "Empresa / PDFs", "Editor de branding indisponivel.")
                return
            data = dict(getter_brand() or {})
            emit = dict(data.get("guia_emitente", {}) or {})
            brand_dialog = QDialog(dialog)
            brand_dialog.setWindowTitle("Empresa / PDFs")
            brand_dialog.setMinimumSize(860, 620)
            brand_layout = QVBoxLayout(brand_dialog)
            brand_layout.setContentsMargins(14, 14, 14, 14)
            brand_layout.setSpacing(12)

            brand_form = QFormLayout()
            brand_form.setHorizontalSpacing(12)
            brand_form.setVerticalSpacing(8)
            logo_edit = QLineEdit(str(data.get("logo_path", "") or "").strip())
            use_default_logo_btn = QPushButton("Usar image (1)")
            use_default_logo_btn.setProperty("variant", "secondary")
            browse_logo_btn = QPushButton("Procurar")
            browse_logo_btn.setProperty("variant", "secondary")
            logo_row = QWidget()
            logo_row_layout = QHBoxLayout(logo_row)
            logo_row_layout.setContentsMargins(0, 0, 0, 0)
            logo_row_layout.setSpacing(8)
            logo_row_layout.addWidget(logo_edit, 1)
            logo_row_layout.addWidget(use_default_logo_btn)
            logo_row_layout.addWidget(browse_logo_btn)
            primary_edit = QLineEdit(str(data.get("primary_color", "#000040") or "#000040"))
            emit_name_edit = QLineEdit(str(emit.get("nome", "") or "").strip())
            emit_nif_edit = QLineEdit(str(emit.get("nif", "") or "").strip())
            emit_address_edit = QLineEdit(str(emit.get("morada", "") or "").strip())
            emit_carga_edit = QLineEdit(str(emit.get("local_carga", "") or "").strip())
            guia_serie_edit = QLineEdit(str(data.get("guia_serie_id", "") or "").strip())
            guia_validation_edit = QLineEdit(str(data.get("guia_validation_code", "") or "").strip())
            brand_form.addRow("Logo", logo_row)
            brand_form.addRow("Cor principal", primary_edit)
            brand_form.addRow("Nome emitente", emit_name_edit)
            brand_form.addRow("NIF", emit_nif_edit)
            brand_form.addRow("Morada", emit_address_edit)
            brand_form.addRow("Local carga", emit_carga_edit)
            brand_form.addRow("Série guia", guia_serie_edit)
            brand_form.addRow("Código AT/ATCUD", guia_validation_edit)
            brand_layout.addLayout(brand_form)

            rodape_title = QLabel("Rodape da empresa nos PDFs")
            rodape_title.setStyleSheet("font-size: 14px; font-weight: 700; color: #0f172a;")
            brand_layout.addWidget(rodape_title)
            rodape_edit = QTextEdit()
            rodape_edit.setPlaceholderText("Uma linha por entrada do rodape.")
            rodape_edit.setPlainText("\n".join(str(v).strip() for v in list(data.get("empresa_info_rodape", []) or []) if str(v).strip()))
            rodape_edit.setMinimumHeight(120)
            brand_layout.addWidget(rodape_edit)

            extra_title = QLabel("Informacao extra da guia / PDF")
            extra_title.setStyleSheet("font-size: 14px; font-weight: 700; color: #0f172a;")
            brand_layout.addWidget(extra_title)
            extra_edit = QTextEdit()
            extra_edit.setPlaceholderText("Linhas adicionais para guias e PDFs.")
            extra_edit.setPlainText("\n".join(str(v).strip() for v in list(data.get("guia_info_extra", []) or []) if str(v).strip()))
            extra_edit.setMinimumHeight(110)
            brand_layout.addWidget(extra_edit)

            def use_default_logo() -> None:
                logo_edit.setText(str(self.backend.base_dir / "Logos" / "image (1).jpg"))

            def browse_logo() -> None:
                path, _selected = QFileDialog.getOpenFileName(
                    brand_dialog,
                    "Selecionar logotipo",
                    str(self.backend.base_dir),
                    "Imagens (*.png *.jpg *.jpeg *.bmp)",
                )
                if path:
                    logo_edit.setText(path)

            use_default_logo_btn.clicked.connect(use_default_logo)
            browse_logo_btn.clicked.connect(browse_logo)

            buttons_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            buttons_box.button(QDialogButtonBox.Ok).setText("Guardar")
            buttons_box.button(QDialogButtonBox.Cancel).setText("Cancelar")
            buttons_box.accepted.connect(brand_dialog.accept)
            buttons_box.rejected.connect(brand_dialog.reject)
            brand_layout.addWidget(buttons_box)

            if brand_dialog.exec() != QDialog.Accepted:
                return
            try:
                saver_brand(
                    {
                        "logo_path": logo_edit.text().strip(),
                        "primary_color": primary_edit.text().strip(),
                        "empresa_info_rodape": rodape_edit.toPlainText(),
                        "guia_emitente": {
                            "nome": emit_name_edit.text().strip(),
                            "nif": emit_nif_edit.text().strip(),
                            "morada": emit_address_edit.text().strip(),
                            "local_carga": emit_carga_edit.text().strip(),
                        },
                        "guia_serie_id": guia_serie_edit.text().strip(),
                        "guia_validation_code": guia_validation_edit.text().strip(),
                        "guia_info_extra": extra_edit.toPlainText(),
                    }
                )
            except Exception as exc:
                QMessageBox.critical(brand_dialog, "Empresa / PDFs", str(exc))
                return
            QMessageBox.information(brand_dialog, "Empresa / PDFs", "Informacao da empresa atualizada.")

        def open_trial_dialog() -> None:
            if not trial_manage_allowed:
                QMessageBox.warning(dialog, "Trial / Licenca", "So o login OWNER pode alterar trial/licenca.")
                return
            getter_trial = getattr(self.backend, "trial_status", None)
            activate_trial = getattr(self.backend, "activate_trial_license", None)
            extend_trial = getattr(self.backend, "extend_trial_license", None)
            disable_trial = getattr(self.backend, "disable_trial_license", None)
            if not callable(getter_trial) or not callable(activate_trial) or not callable(extend_trial) or not callable(disable_trial):
                QMessageBox.warning(dialog, "Trial / Licenca", "Gestao de trial indisponivel.")
                return

            trial_dialog = QDialog(dialog)
            trial_dialog.setWindowTitle("Trial / Licenca")
            trial_dialog.setMinimumSize(860, 640)
            trial_dialog.resize(920, 680)
            trial_layout = QVBoxLayout(trial_dialog)
            trial_layout.setContentsMargins(16, 16, 16, 16)
            trial_layout.setSpacing(12)

            intro = QLabel(
                "Ativa um trial por empresa/equipamento. Quando o prazo terminar, "
                "o sistema bloqueia novos acessos e apenas o login do proprietario "
                "configurado no lugest.env pode voltar a autorizar."
            )
            intro.setWordWrap(True)
            trial_layout.addWidget(intro)

            status_card = QLabel("")
            status_card.setWordWrap(True)
            status_card.setStyleSheet(
                """
                QLabel {
                    border-radius: 16px;
                    padding: 12px 14px;
                    background: #eef5fc;
                    border: 1px solid #cdd9ea;
                    color: #36506d;
                    font-weight: 700;
                }
                """
            )
            trial_layout.addWidget(status_card)

            trial_form = QFormLayout()
            trial_form.setHorizontalSpacing(12)
            trial_form.setVerticalSpacing(8)
            company_edit = QLineEdit()
            duration_spin = QSpinBox()
            duration_spin.setRange(1, 3650)
            duration_spin.setSuffix(" dias")
            extend_spin = QSpinBox()
            extend_spin.setRange(1, 3650)
            extend_spin.setValue(30)
            extend_spin.setSuffix(" dias")
            notes_edit = QTextEdit()
            notes_edit.setPlaceholderText("Observacoes internas do trial.")
            notes_edit.setMinimumHeight(96)
            state_value = QLabel("-")
            owner_value = QLabel("-")
            owner_value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            device_value = QLabel("-")
            device_value.setWordWrap(True)
            device_value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            started_value = QLabel("-")
            expires_value = QLabel("-")
            remaining_value = QLabel("-")
            last_success_value = QLabel("-")
            last_owner_value = QLabel("-")
            trial_form.addRow("Empresa", company_edit)
            trial_form.addRow("Duracao inicial", duration_spin)
            trial_form.addRow("Prolongar", extend_spin)
            trial_form.addRow("Estado", state_value)
            trial_form.addRow("Login proprietario", owner_value)
            trial_form.addRow("Equipamento atual", device_value)
            trial_form.addRow("Inicio", started_value)
            trial_form.addRow("Expira em", expires_value)
            trial_form.addRow("Dias restantes", remaining_value)
            trial_form.addRow("Ultimo acesso", last_success_value)
            trial_form.addRow("Ultima autorizacao", last_owner_value)
            trial_form.addRow("Notas", notes_edit)
            trial_layout.addLayout(trial_form)

            actions_row = QHBoxLayout()
            actions_row.setSpacing(10)
            activate_btn = QPushButton("Ativar / Reiniciar")
            activate_btn.setProperty("variant", "secondary")
            extend_btn = QPushButton("Prolongar trial")
            extend_btn.setProperty("variant", "secondary")
            disable_btn = QPushButton("Desativar")
            disable_btn.setProperty("variant", "danger")
            actions_row.addWidget(activate_btn)
            actions_row.addWidget(extend_btn)
            actions_row.addWidget(disable_btn)
            actions_row.addStretch(1)
            trial_layout.addLayout(actions_row)

            buttons_box = QDialogButtonBox(QDialogButtonBox.Close)
            buttons_box.rejected.connect(trial_dialog.reject)
            buttons_box.button(QDialogButtonBox.Close).setText("Fechar")
            trial_layout.addWidget(buttons_box)

            def format_dt(raw_value: str) -> str:
                value = str(raw_value or "").strip()
                if not value:
                    return "-"
                try:
                    from datetime import datetime

                    return datetime.fromisoformat(value.replace("Z", "+00:00")).strftime("%d/%m/%Y %H:%M")
                except Exception:
                    return value

            def describe_state(status: dict) -> str:
                state = str(status.get("state", "") or "").strip().lower()
                mapping = {
                    "active": "Ativo",
                    "expired": "Expirado",
                    "disabled": "Desativado",
                    "invalid": "Invalido",
                    "device_mismatch": "Equipamento diferente",
                }
                label = mapping.get(state, state or "-")
                if bool(status.get("blocking", False)):
                    return f"{label} | bloqueado"
                return label

            def refresh_trial_status(sync_inputs: bool = True) -> dict:
                status = dict(getter_trial() or {})
                blocking = bool(status.get("blocking", False))
                status_card.setText(str(status.get("message", "") or "Sem informacao de licenciamento.").strip())
                status_card.setStyleSheet(
                    """
                    QLabel {
                        border-radius: 16px;
                        padding: 12px 14px;
                        font-weight: 700;
                        background: %s;
                        border: 1px solid %s;
                        color: %s;
                    }
                    """
                    % (
                        "#fff1f2" if blocking else "#eef5fc",
                        "#fecaca" if blocking else "#cdd9ea",
                        "#9f1239" if blocking else "#36506d",
                    )
                )
                state_value.setText(describe_state(status))
                owner_name = str(status.get("owner_username", "") or "").strip()
                if not owner_name:
                    owner_name = "-"
                if not bool(status.get("owner_configured", False)):
                    owner_name = f"{owner_name} | nao configurado".strip()
                owner_value.setText(owner_name)
                device_value.setText(str(status.get("current_device_fingerprint", "") or "-").strip() or "-")
                started_value.setText(format_dt(str(status.get("started_at", "") or "")))
                expires_value.setText(format_dt(str(status.get("expires_at", "") or "")))
                days_remaining = status.get("days_remaining")
                remaining_value.setText("-" if days_remaining is None else str(days_remaining))
                last_success_value.setText(
                    f"{format_dt(str(status.get('last_success_at', '') or ''))} | {str(status.get('last_success_user', '') or '-').strip() or '-'}"
                )
                last_owner_value.setText(
                    f"{format_dt(str(status.get('last_owner_auth_at', '') or ''))} | {str(status.get('last_owner_auth_user', '') or '-').strip() or '-'}"
                )
                if sync_inputs:
                    company_edit.setText(str(status.get("company_name", "") or "").strip())
                    duration_spin.setValue(max(1, int(status.get("duration_days", 60) or 60)))
                    notes_edit.setPlainText(str(status.get("notes", "") or "").strip())
                return status

            def on_activate_trial() -> None:
                company_name = company_edit.text().strip()
                if not company_name:
                    QMessageBox.warning(trial_dialog, "Trial / Licenca", "Indica o nome da empresa.")
                    return
                if QMessageBox.question(
                    trial_dialog,
                    "Ativar trial",
                    f"Ativar ou reiniciar o trial para {company_name} neste equipamento?",
                ) != QMessageBox.Yes:
                    return
                try:
                    activate_trial(
                        company_name=company_name,
                        duration_days=int(duration_spin.value()),
                        notes=notes_edit.toPlainText().strip(),
                    )
                except Exception as exc:
                    QMessageBox.critical(trial_dialog, "Trial / Licenca", str(exc))
                    return
                refresh_trial_status(sync_inputs=True)
                QMessageBox.information(trial_dialog, "Trial / Licenca", "Trial ativado com sucesso.")

            def on_extend_trial() -> None:
                try:
                    extend_trial(extra_days=int(extend_spin.value()))
                except Exception as exc:
                    QMessageBox.critical(trial_dialog, "Trial / Licenca", str(exc))
                    return
                refresh_trial_status(sync_inputs=True)
                QMessageBox.information(trial_dialog, "Trial / Licenca", "Trial prolongado com sucesso.")

            def on_disable_trial() -> None:
                if QMessageBox.question(
                    trial_dialog,
                    "Desativar trial",
                    "Desativar o controlo de trial neste equipamento?",
                ) != QMessageBox.Yes:
                    return
                try:
                    disable_trial()
                except Exception as exc:
                    QMessageBox.critical(trial_dialog, "Trial / Licenca", str(exc))
                    return
                refresh_trial_status(sync_inputs=True)
                QMessageBox.information(trial_dialog, "Trial / Licenca", "Trial desativado.")

            activate_btn.clicked.connect(on_activate_trial)
            extend_btn.clicked.connect(on_extend_trial)
            disable_btn.clicked.connect(on_disable_trial)
            refresh_trial_status(sync_inputs=True)
            trial_dialog.exec()

        def open_operations_dialog() -> None:
            rows_getter = getattr(self.backend, "operation_catalog_rows", None)
            saver = getattr(self.backend, "save_operation_catalog_row", None)
            remover = getattr(self.backend, "remove_operation_catalog_row", None)
            if not callable(rows_getter) or not callable(saver) or not callable(remover):
                QMessageBox.warning(dialog, "Operacoes", "Gestao de operacoes indisponivel.")
                return
            op_dialog = QDialog(dialog)
            op_dialog.setWindowTitle("Operacoes")
            op_dialog.setMinimumSize(760, 520)
            op_layout = QHBoxLayout(op_dialog)
            op_layout.setContentsMargins(14, 14, 14, 14)
            op_layout.setSpacing(12)
            table = QTableWidget(0, 3)
            table.setHorizontalHeaderLabels(["Nome", "Ativa", "Planeavel"])
            table.verticalHeader().setVisible(False)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectRows)
            table.setSelectionMode(QAbstractItemView.SingleSelection)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            op_layout.addWidget(table, 3)

            side = QWidget()
            side_layout = QVBoxLayout(side)
            side_layout.setContentsMargins(0, 0, 0, 0)
            form = QFormLayout()
            name_edit = QLineEdit()
            name_edit.setPlaceholderText("Ex.: Laser, Quinagem, Desenho")
            active_box = QCheckBox("Ativa")
            active_box.setChecked(True)
            planeavel_box = QCheckBox("Entra no planeamento")
            form.addRow("Nome", name_edit)
            form.addRow("", active_box)
            form.addRow("", planeavel_box)
            side_layout.addLayout(form)
            new_btn = QPushButton("Nova")
            save_btn = QPushButton("Guardar")
            remove_btn = QPushButton("Remover")
            remove_btn.setProperty("variant", "danger")
            side_layout.addWidget(new_btn)
            side_layout.addWidget(save_btn)
            side_layout.addWidget(remove_btn)
            side_layout.addStretch(1)
            close_buttons = QDialogButtonBox(QDialogButtonBox.Close)
            close_buttons.button(QDialogButtonBox.Close).setText("Fechar")
            close_buttons.rejected.connect(op_dialog.reject)
            side_layout.addWidget(close_buttons)
            op_layout.addWidget(side, 2)
            current_name = {"value": ""}

            def refresh_operations(select_name: str = "") -> None:
                rows = list(rows_getter() or [])
                table.setRowCount(len(rows))
                target = -1
                for row_index, row in enumerate(rows):
                    values = [
                        str(row.get("name", "") or ""),
                        "Sim" if bool(row.get("active", True)) else "Nao",
                        "Sim" if bool(row.get("planeavel", False)) else "Nao",
                    ]
                    for col_index, value in enumerate(values):
                        item = QTableWidgetItem(value)
                        if col_index == 0:
                            item.setData(Qt.UserRole, dict(row))
                        table.setItem(row_index, col_index, item)
                    if select_name and values[0].casefold() == select_name.casefold():
                        target = row_index
                if target >= 0:
                    table.selectRow(target)
                elif rows:
                    table.selectRow(0)

            def load_selected_operation() -> None:
                item = table.currentItem()
                if item is None:
                    return
                row = dict(table.item(item.row(), 0).data(Qt.UserRole) or {})
                current_name["value"] = str(row.get("name", "") or "").strip()
                name_edit.setText(current_name["value"])
                active_box.setChecked(bool(row.get("active", True)))
                planeavel_box.setChecked(bool(row.get("planeavel", False)))

            def clear_operation_form() -> None:
                current_name["value"] = ""
                name_edit.setText("")
                active_box.setChecked(True)
                planeavel_box.setChecked(False)
                name_edit.setFocus()

            def save_operation() -> None:
                try:
                    row = saver(
                        name_edit.text().strip(),
                        current_name=current_name["value"],
                        active=active_box.isChecked(),
                        planeavel=planeavel_box.isChecked(),
                    )
                except Exception as exc:
                    QMessageBox.critical(op_dialog, "Operacoes", str(exc))
                    return
                saved_name = str(row.get("name", "") or name_edit.text()).strip()
                refresh_operations(saved_name)
                QMessageBox.information(op_dialog, "Operacoes", "Operacao guardada com sucesso.")

            def remove_operation() -> None:
                item = table.currentItem()
                if item is None:
                    QMessageBox.warning(op_dialog, "Operacoes", "Seleciona a operação que pretendes remover.")
                    return
                row = dict(table.item(item.row(), 0).data(Qt.UserRole) or {})
                target_name = str(row.get("name", "") or "").strip()
                if not target_name:
                    QMessageBox.warning(op_dialog, "Operacoes", "Operação inválida.")
                    return
                if QMessageBox.question(op_dialog, "Remover operação", f"Remover a operação '{target_name}'?") != QMessageBox.Yes:
                    return
                try:
                    remover(target_name)
                except Exception as exc:
                    QMessageBox.critical(op_dialog, "Operacoes", str(exc))
                    return
                refresh_operations("")
                clear_operation_form()
                QMessageBox.information(op_dialog, "Operacoes", "Operação removida com sucesso.")

            table.itemSelectionChanged.connect(load_selected_operation)
            new_btn.clicked.connect(clear_operation_form)
            save_btn.clicked.connect(save_operation)
            remove_btn.clicked.connect(remove_operation)
            refresh_operations("")
            op_dialog.exec()

        def open_workcenters_dialog() -> None:
            getter_rows = getattr(self.backend, "workcenter_rows", None)
            group_options_getter = getattr(self.backend, "workcenter_group_options", None)
            operation_options_getter = getattr(self.backend, "operation_catalog_options", None)
            operation_rows_getter = getattr(self.backend, "operation_catalog_rows", None)
            save_group = getattr(self.backend, "save_workcenter_group", None)
            remove_group = getattr(self.backend, "remove_workcenter_group", None)
            save_machine = getattr(self.backend, "save_workcenter_machine", None)
            remove_machine = getattr(self.backend, "remove_workcenter_machine", None)
            if (
                not callable(getter_rows)
                or not callable(group_options_getter)
                or not callable(operation_options_getter)
                or not callable(operation_rows_getter)
                or not callable(save_group)
                or not callable(remove_group)
                or not callable(save_machine)
                or not callable(remove_machine)
            ):
                QMessageBox.warning(dialog, "Postos de Trabalho", "Gestao de postos indisponivel.")
                return

            wc_dialog = QDialog(dialog)
            wc_dialog.setWindowTitle("Postos de Trabalho")
            wc_dialog.setMinimumSize(1120, 620)
            wc_layout = QVBoxLayout(wc_dialog)
            wc_layout.setContentsMargins(14, 14, 14, 14)
            wc_layout.setSpacing(12)

            intro = QLabel(
                "Organiza a estrutura da empresa por posto de trabalho e por maquina/recurso. "
                "Exemplo: posto 'Corte Laser' com maquinas 'Maquina 3030', 'Maquina 5030' e 'Maquina 5040'."
            )
            intro.setWordWrap(True)
            wc_layout.addWidget(intro)

            host = QWidget()
            host_layout = QHBoxLayout(host)
            host_layout.setContentsMargins(0, 0, 0, 0)
            host_layout.setSpacing(12)

            operation_panel = QWidget()
            operation_panel_layout = QVBoxLayout(operation_panel)
            operation_panel_layout.setContentsMargins(0, 0, 0, 0)
            operation_panel_layout.setSpacing(8)
            operation_title = QLabel("Operacoes")
            operation_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
            operation_panel_layout.addWidget(operation_title)
            operation_hint = QLabel("Seleciona uma operacao para ver apenas os postos e maquinas associados.")
            operation_hint.setProperty("role", "muted")
            operation_hint.setWordWrap(True)
            operation_panel_layout.addWidget(operation_hint)
            operations_filter_table = QTableWidget(0, 3)
            operations_filter_table.setHorizontalHeaderLabels(["Operacao", "Ativa", "Plan."])
            operations_filter_table.verticalHeader().setVisible(False)
            operations_filter_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            operations_filter_table.setSelectionBehavior(QAbstractItemView.SelectRows)
            operations_filter_table.setSelectionMode(QAbstractItemView.SingleSelection)
            operations_filter_table.setAlternatingRowColors(True)
            operations_filter_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            operations_filter_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            operations_filter_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            operation_panel_layout.addWidget(operations_filter_table, 1)
            host_layout.addWidget(operation_panel, 2)

            table_panel = QWidget()
            table_panel_layout = QVBoxLayout(table_panel)
            table_panel_layout.setContentsMargins(0, 0, 0, 0)
            table_panel_layout.setSpacing(8)

            table_tools = QHBoxLayout()
            table_tools.setContentsMargins(0, 0, 0, 0)
            table_tools.setSpacing(8)
            table_hint = QLabel("Seleciona uma linha, edita à direita e guarda. Também podes eliminar a linha selecionada.")
            table_hint.setProperty("role", "muted")
            table_tools.addWidget(table_hint, 1)
            edit_selected_btn = QPushButton("Editar linha selecionada")
            edit_selected_btn.setProperty("variant", "secondary")
            table_tools.addWidget(edit_selected_btn, 0)
            remove_selected_btn = QPushButton("Eliminar linha selecionada")
            remove_selected_btn.setProperty("variant", "danger")
            table_tools.addWidget(remove_selected_btn, 0)
            table_panel_layout.addLayout(table_tools)

            workcenters_table = QTableWidget(0, 9)
            workcenters_table.setHorizontalHeaderLabels(["Tipo", "Nome", "Grupo", "Operacao", "Ativo", "Utiliz.", "Orc.", "Enc.", "Plan."])
            workcenters_table.verticalHeader().setVisible(False)
            workcenters_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            workcenters_table.setSelectionBehavior(QAbstractItemView.SelectRows)
            workcenters_table.setSelectionMode(QAbstractItemView.SingleSelection)
            workcenters_table.setAlternatingRowColors(True)
            workcenters_table.setWordWrap(False)
            workcenters_table.horizontalHeader().setStretchLastSection(False)
            workcenters_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            workcenters_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            workcenters_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            workcenters_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
            for column in range(4, 9):
                workcenters_table.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeToContents)
            table_panel_layout.addWidget(workcenters_table, 1)
            host_layout.addWidget(table_panel, 5)

            side = QWidget()
            side_layout = QVBoxLayout(side)
            side_layout.setContentsMargins(0, 0, 0, 0)
            side_layout.setSpacing(10)

            side_title = QLabel("Gestao da estrutura")
            side_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
            side_layout.addWidget(side_title)

            side_note = QLabel(
                "Primeiro defines o posto principal. Depois, se fizer sentido, adicionas as maquinas ou bancadas que trabalham dentro desse posto."
            )
            side_note.setProperty("role", "muted")
            side_note.setWordWrap(True)
            side_layout.addWidget(side_note)

            group_header = QLabel("Posto de trabalho")
            group_header.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
            side_layout.addWidget(group_header)

            group_form = QFormLayout()
            group_form.setHorizontalSpacing(12)
            group_form.setVerticalSpacing(8)
            group_name_edit = QLineEdit()
            group_name_edit.setPlaceholderText("Ex.: Corte Laser, Quinagem, Serralharia")
            group_operation_combo = QComboBox()
            for op_name in list(operation_options_getter() or []):
                group_operation_combo.addItem(str(op_name))
            group_active_box = QCheckBox("Posto ativo")
            group_active_box.setChecked(True)
            group_form.addRow("Nome do posto", group_name_edit)
            group_form.addRow("Operacao base", group_operation_combo)
            group_form.addRow("", group_active_box)
            side_layout.addLayout(group_form)

            group_usage_label = QLabel("Novo posto sem utilizacao.")
            group_usage_label.setWordWrap(True)
            group_usage_label.setStyleSheet(
                """
                QLabel {
                    border-radius: 14px;
                    padding: 10px 12px;
                    background: #eef5fc;
                    border: 1px solid #cdd9ea;
                    color: #36506d;
                    font-weight: 600;
                }
                """
            )
            side_layout.addWidget(group_usage_label)

            group_note = QLabel(
                "Um posto pode existir sozinho. Se precisares de detalhe operacional, adicionas depois as maquinas ou bancadas por baixo."
            )
            group_note.setProperty("role", "muted")
            group_note.setWordWrap(True)
            side_layout.addWidget(group_note)

            group_actions = QHBoxLayout()
            group_actions.setSpacing(8)
            new_group_btn = QPushButton("Novo posto")
            new_group_btn.setProperty("variant", "secondary")
            save_group_btn = QPushButton("Guardar posto")
            remove_group_btn = QPushButton("Remover posto")
            remove_group_btn.setProperty("variant", "danger")
            group_actions.addWidget(new_group_btn)
            group_actions.addWidget(save_group_btn)
            group_actions.addWidget(remove_group_btn)
            side_layout.addLayout(group_actions)

            machine_header = QLabel("Maquina / recurso")
            machine_header.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a; margin-top: 8px;")
            side_layout.addWidget(machine_header)

            machine_form = QFormLayout()
            machine_form.setHorizontalSpacing(12)
            machine_form.setVerticalSpacing(8)
            machine_group_combo = QComboBox()
            machine_name_edit = QLineEdit()
            machine_name_edit.setPlaceholderText("Ex.: Maquina 3030, Bancada 1, Serra 2")
            machine_active_box = QCheckBox("Recurso ativo")
            machine_active_box.setChecked(True)
            machine_form.addRow("Posto pai", machine_group_combo)
            machine_form.addRow("Nome da maquina", machine_name_edit)
            machine_form.addRow("", machine_active_box)
            side_layout.addLayout(machine_form)

            machine_usage_label = QLabel("Nova maquina sem utilizacao.")
            machine_usage_label.setWordWrap(True)
            machine_usage_label.setStyleSheet(
                """
                QLabel {
                    border-radius: 14px;
                    padding: 10px 12px;
                    background: #f8fafc;
                    border: 1px solid #d7e2ef;
                    color: #475467;
                    font-weight: 600;
                }
                """
            )
            side_layout.addWidget(machine_usage_label)

            machine_note = QLabel(
                "A maquina fica ligada ao posto principal. Isto permite ter varias maquinas a trabalhar dentro do mesmo posto."
            )
            machine_note.setProperty("role", "muted")
            machine_note.setWordWrap(True)
            side_layout.addWidget(machine_note)

            machine_actions = QHBoxLayout()
            machine_actions.setSpacing(8)
            new_machine_btn = QPushButton("Nova maquina")
            new_machine_btn.setProperty("variant", "secondary")
            save_machine_btn = QPushButton("Guardar maquina")
            remove_machine_btn = QPushButton("Remover maquina")
            remove_machine_btn.setProperty("variant", "danger")
            machine_actions.addWidget(new_machine_btn)
            machine_actions.addWidget(save_machine_btn)
            machine_actions.addWidget(remove_machine_btn)
            side_layout.addLayout(machine_actions)
            side_layout.addStretch(1)
            host_layout.addWidget(side, 4)

            wc_layout.addWidget(host, 1)

            wc_buttons = QDialogButtonBox(QDialogButtonBox.Close)
            wc_buttons.button(QDialogButtonBox.Close).setText("Fechar")
            wc_buttons.rejected.connect(wc_dialog.reject)
            wc_layout.addWidget(wc_buttons)

            current_group = {"value": ""}
            current_machine = {"value": ""}
            current_operation_filter = {"value": ""}

            def describe_usage(row: dict | None = None, empty_text: str = "Sem utilizacao.") -> str:
                if not isinstance(row, dict) or not str(row.get("name", "") or "").strip():
                    return empty_text
                return (
                    f"Em uso por {int(row.get('users', 0) or 0)} utilizador(es), "
                    f"{int(row.get('quotes', 0) or 0)} orçamento(s), "
                    f"{int(row.get('orders', 0) or 0)} encomenda(s) e "
                    f"{int(row.get('planning', 0) or 0)} registo(s) de planeamento."
                )

            def refresh_group_options(selected_group: str = "") -> None:
                target = str(selected_group or machine_group_combo.currentText() or "").strip()
                options = list(group_options_getter() or [])
                machine_group_combo.blockSignals(True)
                machine_group_combo.clear()
                for group_name in options:
                    machine_group_combo.addItem(str(group_name))
                if target:
                    machine_group_combo.setCurrentText(target)
                elif options:
                    machine_group_combo.setCurrentIndex(0)
                machine_group_combo.blockSignals(False)

            def selected_operation_name() -> str:
                item = operations_filter_table.currentItem()
                if item is None:
                    return ""
                return str(operations_filter_table.item(item.row(), 0).data(Qt.UserRole) or "").strip()

            def refresh_operation_filter(select_name: str = "") -> None:
                rows = [{"name": "", "active": True, "planeavel": False}]
                rows.extend(list(operation_rows_getter() or []))
                target = str(select_name or current_operation_filter["value"] or "").strip()
                operations_filter_table.blockSignals(True)
                operations_filter_table.setRowCount(len(rows))
                target_row = 0
                for row_index, row in enumerate(rows):
                    name = str(row.get("name", "") or "").strip()
                    values = [
                        name or "Todas",
                        "Sim" if bool(row.get("active", True)) else "Nao",
                        "Sim" if bool(row.get("planeavel", False)) else "Nao",
                    ]
                    for col_index, value in enumerate(values):
                        item = QTableWidgetItem(value)
                        if col_index == 0:
                            item.setData(Qt.UserRole, name)
                        operations_filter_table.setItem(row_index, col_index, item)
                    if target and name.casefold() == target.casefold():
                        target_row = row_index
                operations_filter_table.selectRow(target_row)
                current_operation_filter["value"] = selected_operation_name()
                operations_filter_table.blockSignals(False)

            def clear_group_form() -> None:
                current_group["value"] = ""
                group_name_edit.setText("")
                selected_op = str(current_operation_filter["value"] or "").strip()
                if selected_op and group_operation_combo.findText(selected_op) >= 0:
                    group_operation_combo.setCurrentText(selected_op)
                elif group_operation_combo.count() > 0:
                    group_operation_combo.setCurrentIndex(0)
                group_active_box.setChecked(True)
                group_usage_label.setText(describe_usage(None, "Novo posto sem utilizacao."))

            def clear_machine_form() -> None:
                current_machine["value"] = ""
                machine_name_edit.setText("")
                machine_active_box.setChecked(True)
                machine_usage_label.setText(describe_usage(None, "Nova maquina sem utilizacao."))

            def clear_workcenter_form() -> None:
                clear_group_form()
                clear_machine_form()
                workcenters_table.clearSelection()
                edit_selected_btn.setEnabled(False)
                remove_selected_btn.setEnabled(False)

            def selected_workcenter_row() -> dict:
                current_item = workcenters_table.currentItem()
                if current_item is None:
                    return {}
                payload = dict(
                    (
                        workcenters_table.item(current_item.row(), 0).data(Qt.UserRole)
                        if workcenters_table.item(current_item.row(), 0) is not None
                        else {}
                    )
                    or {}
                )
                row_name = str(payload.get("name", "") or "").strip().lower()
                row_type = str(payload.get("entry_type", "") or "").strip()
                return next(
                    (
                        row
                        for row in list(getter_rows() or [])
                        if str(row.get("entry_type", "") or "").strip() == row_type
                        and str(row.get("name", "") or "").strip().lower() == row_name
                    ),
                    {},
                )

            def load_workcenter_form(row: dict) -> None:
                edit_selected_btn.setEnabled(True)
                remove_selected_btn.setEnabled(True)
                entry_type = str(row.get("entry_type", "") or "").strip()
                if entry_type == "group":
                    current_group["value"] = str(row.get("name", "") or "").strip()
                    group_name_edit.setText(current_group["value"])
                    group_operation_combo.setCurrentText(str(row.get("operation", "") or "").strip())
                    group_active_box.setChecked(bool(row.get("active", True)))
                    group_usage_label.setText(describe_usage(row, "Novo posto sem utilizacao."))
                    machine_group_combo.setCurrentText(current_group["value"])
                    clear_machine_form()
                    return
                current_machine["value"] = str(row.get("name", "") or "").strip()
                machine_name_edit.setText(current_machine["value"])
                machine_active_box.setChecked(bool(row.get("active", True)))
                machine_group_combo.setCurrentText(str(row.get("group", "") or "").strip())
                machine_usage_label.setText(describe_usage(row, "Nova maquina sem utilizacao."))
                parent_group = str(row.get("group", "") or "").strip()
                current_group["value"] = parent_group
                group_name_edit.setText(parent_group)
                group_operation_combo.setCurrentText(str(row.get("operation", "") or "").strip())
                parent_row = next(
                    (
                        candidate
                        for candidate in list(getter_rows() or [])
                        if str(candidate.get("entry_type", "") or "").strip() == "group"
                        and str(candidate.get("name", "") or "").strip().lower() == parent_group.lower()
                    ),
                    {},
                )
                group_active_box.setChecked(bool(parent_row.get("active", True)))
                group_usage_label.setText(describe_usage(parent_row, "Novo posto sem utilizacao."))

            def refresh_workcenters(select_name: str = "", select_type: str = "") -> None:
                rows = list(getter_rows() or [])
                selected_op = str(current_operation_filter["value"] or "").strip()
                if selected_op:
                    rows = [
                        row
                        for row in rows
                        if str(row.get("operation", "") or "").strip().casefold() == selected_op.casefold()
                    ]
                refresh_group_options(machine_group_combo.currentText().strip())
                workcenters_table.setRowCount(len(rows))
                edit_selected_btn.setEnabled(bool(rows))
                remove_selected_btn.setEnabled(bool(rows))
                target_row = -1
                for row_index, row in enumerate(rows):
                    values = [
                        str(row.get("kind", "") or "").strip() or "-",
                        str(row.get("name", "") or "").strip(),
                        str(row.get("group", "") or "").strip() or "-",
                        str(row.get("operation", "") or "").strip() or "-",
                        "Sim" if bool(row.get("active", True)) else "Nao",
                        str(int(row.get("users", 0) or 0)),
                        str(int(row.get("quotes", 0) or 0)),
                        str(int(row.get("orders", 0) or 0)),
                        str(int(row.get("planning", 0) or 0)),
                    ]
                    for col_index, value in enumerate(values):
                        item = QTableWidgetItem(value)
                        if col_index == 0:
                            item.setData(
                                Qt.UserRole,
                                {
                                    "name": str(row.get("name", "") or "").strip(),
                                    "entry_type": str(row.get("entry_type", "") or "").strip(),
                                },
                            )
                        workcenters_table.setItem(row_index, col_index, item)
                    if (
                        select_name
                        and str(row.get("name", "") or "").strip().lower() == select_name.lower()
                        and (not select_type or str(row.get("entry_type", "") or "").strip() == select_type)
                    ):
                        target_row = row_index
                if target_row >= 0:
                    workcenters_table.selectRow(target_row)
                elif rows:
                    workcenters_table.selectRow(0)
                else:
                    clear_workcenter_form()

            def on_workcenter_selected() -> None:
                row = selected_workcenter_row()
                if row:
                    load_workcenter_form(row)
                else:
                    edit_selected_btn.setEnabled(False)
                    remove_selected_btn.setEnabled(False)

            def on_operation_filter_selected() -> None:
                current_operation_filter["value"] = selected_operation_name()
                if current_operation_filter["value"]:
                    group_operation_combo.setCurrentText(current_operation_filter["value"])
                refresh_workcenters("")

            def edit_selected_workcenter_row() -> None:
                row = selected_workcenter_row()
                if not row:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Seleciona a linha que pretendes editar.")
                    return
                load_workcenter_form(row)
                if str(row.get("entry_type", "") or "").strip() == "group":
                    group_name_edit.setFocus()
                    group_name_edit.selectAll()
                else:
                    machine_name_edit.setFocus()
                    machine_name_edit.selectAll()

            def remove_selected_workcenter_row() -> None:
                row = selected_workcenter_row()
                entry_type = str(row.get("entry_type", "") or "").strip()
                target_name = str(row.get("name", "") or "").strip()
                if not entry_type or not target_name:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Seleciona a linha que pretendes eliminar.")
                    return
                if entry_type == "group":
                    title = "Remover posto"
                    prompt = f"Remover o posto '{target_name}'?"
                else:
                    title = "Remover maquina"
                    prompt = f"Remover a maquina '{target_name}'?"
                if QMessageBox.question(wc_dialog, title, prompt) != QMessageBox.Yes:
                    return
                try:
                    if entry_type == "group":
                        remove_group(target_name)
                    else:
                        remove_machine(target_name)
                except Exception as exc:
                    QMessageBox.critical(wc_dialog, "Postos de Trabalho", str(exc))
                    return
                refresh_posto_options(posto_combo.currentText().strip())
                refresh_users(current_username["value"])
                refresh_workcenters("")
                clear_workcenter_form()
                QMessageBox.information(
                    wc_dialog,
                    "Postos de Trabalho",
                    "Maquina removida com sucesso." if entry_type == "machine" else "Posto removido com sucesso.",
                )

            def on_save_group() -> None:
                group_name = group_name_edit.text().strip()
                if not group_name:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Indica o nome do posto.")
                    return
                try:
                    result = save_group(
                        name=group_name,
                        operation=group_operation_combo.currentText().strip(),
                        current_name=current_group["value"],
                        active=group_active_box.isChecked(),
                    )
                except Exception as exc:
                    QMessageBox.critical(wc_dialog, "Postos de Trabalho", str(exc))
                    return
                saved_name = str(result.get("name", "") or group_name).strip()
                refresh_posto_options(posto_combo.currentText().strip())
                refresh_users(current_username["value"])
                refresh_workcenters(saved_name, "group")
                QMessageBox.information(wc_dialog, "Postos de Trabalho", "Posto guardado com sucesso.")

            def on_remove_group() -> None:
                row = selected_workcenter_row()
                if str(row.get("entry_type", "") or "").strip() == "group":
                    remove_selected_workcenter_row()
                    return
                target_name = str(current_group["value"] or "").strip()
                if not target_name:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Seleciona um posto de trabalho.")
                    return
                if QMessageBox.question(
                    wc_dialog,
                    "Remover posto",
                    f"Remover o posto '{target_name}'?",
                ) != QMessageBox.Yes:
                    return
                try:
                    remove_group(target_name)
                except Exception as exc:
                    QMessageBox.critical(wc_dialog, "Postos de Trabalho", str(exc))
                    return
                refresh_posto_options(posto_combo.currentText().strip())
                refresh_users(current_username["value"])
                refresh_workcenters("")
                clear_workcenter_form()
                QMessageBox.information(wc_dialog, "Postos de Trabalho", "Posto removido com sucesso.")

            def on_save_machine() -> None:
                machine_name = machine_name_edit.text().strip()
                parent_group = machine_group_combo.currentText().strip() or current_group["value"]
                if not parent_group:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Seleciona primeiro o posto pai da maquina.")
                    return
                if not machine_name:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Indica o nome da maquina.")
                    return
                try:
                    result = save_machine(
                        group_name=parent_group,
                        machine_name=machine_name,
                        current_name=current_machine["value"],
                        active=machine_active_box.isChecked(),
                    )
                except Exception as exc:
                    QMessageBox.critical(wc_dialog, "Postos de Trabalho", str(exc))
                    return
                saved_name = str(result.get("name", "") or machine_name).strip()
                refresh_posto_options(posto_combo.currentText().strip())
                refresh_users(current_username["value"])
                refresh_workcenters(saved_name, "machine")
                QMessageBox.information(wc_dialog, "Postos de Trabalho", "Maquina guardada com sucesso.")

            def on_remove_machine() -> None:
                row = selected_workcenter_row()
                if str(row.get("entry_type", "") or "").strip() == "machine":
                    remove_selected_workcenter_row()
                    return
                target_name = str(current_machine["value"] or "").strip()
                if not target_name:
                    QMessageBox.warning(wc_dialog, "Postos de Trabalho", "Seleciona uma maquina.")
                    return
                if QMessageBox.question(
                    wc_dialog,
                    "Remover maquina",
                    f"Remover a maquina '{target_name}'?",
                ) != QMessageBox.Yes:
                    return
                try:
                    remove_machine(target_name)
                except Exception as exc:
                    QMessageBox.critical(wc_dialog, "Postos de Trabalho", str(exc))
                    return
                refresh_posto_options(posto_combo.currentText().strip())
                refresh_users(current_username["value"])
                refresh_workcenters("")
                clear_machine_form()
                QMessageBox.information(wc_dialog, "Postos de Trabalho", "Maquina removida com sucesso.")

            workcenters_table.itemSelectionChanged.connect(on_workcenter_selected)
            workcenters_table.itemDoubleClicked.connect(lambda *_: edit_selected_workcenter_row())
            operations_filter_table.itemSelectionChanged.connect(on_operation_filter_selected)
            new_group_btn.clicked.connect(clear_group_form)
            new_machine_btn.clicked.connect(clear_machine_form)
            save_group_btn.clicked.connect(on_save_group)
            remove_group_btn.clicked.connect(on_remove_group)
            save_machine_btn.clicked.connect(on_save_machine)
            remove_machine_btn.clicked.connect(on_remove_machine)
            edit_selected_btn.clicked.connect(edit_selected_workcenter_row)
            remove_selected_btn.clicked.connect(remove_selected_workcenter_row)
            refresh_operation_filter("")
            refresh_workcenters("")
            wc_dialog.exec()

        def clear_user_form() -> None:
            current_username["value"] = ""
            username_edit.setText("")
            password_user_edit.setText("")
            password_toggle.setChecked(False)
            sync_password_echo()
            password_user_edit.setPlaceholderText("Obrigatoria para novo utilizador")
            role_combo.setCurrentText("Operador" if role_combo.findText("Operador") >= 0 else role_combo.currentText())
            posto_combo.setCurrentText("")
            active_box.setChecked(True)
            for key, check in permission_checks.items():
                check.setChecked(key == "home")
            update_password_preview("")

        def load_user_form(row: dict) -> None:
            current_username["value"] = str(row.get("username", "") or "").strip()
            username_edit.setText(current_username["value"])
            password_user_edit.setText("")
            password_toggle.setChecked(False)
            sync_password_echo()
            password_user_edit.setPlaceholderText(
                "Deixa vazio para manter a password atual"
                if bool(row.get("password_set", False))
                else "Obrigatoria para definir password"
            )
            role_combo.setCurrentText(str(row.get("role", "Operador") or "Operador"))
            posto_combo.setCurrentText(str(row.get("posto", "") or "").strip())
            active_box.setChecked(bool(row.get("active", True)))
            perms = dict(row.get("menu_permissions", {}) or {})
            if perms:
                for key, check in permission_checks.items():
                    check.setChecked(bool(perms.get(key, False)))
            else:
                for check in permission_checks.values():
                    check.setChecked(True)
            update_password_preview(current_username["value"])

        def refresh_users(select_username: str = "") -> None:
            rows = list(getattr(self.backend, "user_rows", lambda: [])() or [])
            users_table.setRowCount(len(rows))
            target_row = -1
            for row_index, row in enumerate(rows):
                username = str(row.get("username", "") or "").strip()
                values = [
                    username,
                    str(row.get("role", "") or "").strip(),
                    str(row.get("posto", "") or "").strip() or "-",
                    "Sim" if bool(row.get("active", True)) else "Nao",
                ]
                for col_index, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    if col_index == 0:
                        item.setData(Qt.UserRole, username)
                    users_table.setItem(row_index, col_index, item)
                if select_username and username.lower() == select_username.lower():
                    target_row = row_index
            if target_row >= 0:
                users_table.selectRow(target_row)
            elif rows:
                users_table.selectRow(0)
            else:
                clear_user_form()

        def selected_user_row() -> dict:
            current_item = users_table.currentItem()
            if current_item is None:
                return {}
            username = str((users_table.item(current_item.row(), 0).data(Qt.UserRole) if users_table.item(current_item.row(), 0) is not None else "") or "").strip()
            return next((row for row in getattr(self.backend, "user_rows", lambda: [])() if str(row.get("username", "") or "").strip() == username), {})

        def on_user_selected() -> None:
            row = selected_user_row()
            if row:
                load_user_form(row)

        def on_save_user() -> None:
            perms = {key: check.isChecked() for key, check in permission_checks.items()}
            if not any(perms.values()):
                perms["home"] = True
            raw_password = password_user_edit.text().strip()
            payload = {
                "username": username_edit.text().strip(),
                "password": raw_password,
                "role": role_combo.currentText().strip(),
                "posto": posto_combo.currentText().strip(),
                "active": active_box.isChecked(),
                "menu_permissions": perms,
            }
            try:
                result = self.backend.save_user(payload, current_username=current_username["value"])
            except Exception as exc:
                QMessageBox.critical(dialog, "Utilizadores", str(exc))
                return
            current_username["value"] = str(result.get("username", "") or "").strip()
            if raw_password:
                session_passwords[current_username["value"].lower()] = raw_password
            refresh_users(current_username["value"])
            update_password_preview(current_username["value"])
            current_user = dict(self.backend.user or {})
            if str(current_user.get("username", "") or "").strip().lower() == current_username["value"].lower():
                self.user_chip.setText(f"{self.backend.user.get('username', '-')} | {self.backend.user.get('role', '-')}")
            self._apply_navigation_permissions()

        def on_remove_user() -> None:
            row = selected_user_row()
            username = str(row.get("username", "") or "").strip()
            if not username:
                QMessageBox.warning(dialog, "Utilizadores", "Seleciona um utilizador.")
                return
            if QMessageBox.question(dialog, "Remover utilizador", f"Remover utilizador {username}?") != QMessageBox.Yes:
                return
            try:
                self.backend.remove_user(username)
            except Exception as exc:
                QMessageBox.critical(dialog, "Utilizadores", str(exc))
                return
            refresh_users("")
            self._apply_navigation_permissions()

        users_table.itemSelectionChanged.connect(on_user_selected)
        company_btn.clicked.connect(open_company_dialog)
        operations_btn.clicked.connect(open_operations_dialog)
        workcenters_btn.clicked.connect(open_workcenters_dialog)
        updates_btn.clicked.connect(self._open_updates_dialog)
        if trial_btn is not None:
            trial_btn.clicked.connect(open_trial_dialog)
        password_toggle.toggled.connect(sync_password_echo)
        preview_toggle.toggled.connect(sync_preview_echo)
        new_user_btn.clicked.connect(clear_user_form)
        save_user_btn.clicked.connect(on_save_user)
        remove_user_btn.clicked.connect(on_remove_user)
        sync_password_echo()
        sync_preview_echo()
        refresh_posto_options("")
        refresh_users("")

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        supervisor_password = supervisor_edit.text().strip()
        should_update_supervisor = bool(supervisor_password) or not supervisor_password_set
        weak_supervisor = should_update_supervisor and (
            len(supervisor_password) < 8
            or supervisor_password.lower() in {"", "admin", "1234", "123456", "password"}
        )
        if weak_supervisor:
            answer = QMessageBox.question(
                self,
                "Seguranca",
                "A password de supervisor esta vazia ou fraca. Isto reduz a protecao do menu Operador. Guardar assim mesmo?",
            )
            if answer != QMessageBox.Yes:
                return
        if callable(setter):
            if should_update_supervisor:
                setter("operator_supervisor_password", supervisor_password)
            setter("operator_show_client_name", bool(show_client_box.isChecked()))
        if self.backend.user:
            username = str(self.backend.user.get("username", "") or "").strip()
            password = str(self.backend.user.get("_session_password", "") or self.backend.user.get("password", "") or "").strip()
            if username and password:
                try:
                    self.backend.authenticate(username, password)
                except Exception:
                    pass
            self.user_chip.setText(f"{self.backend.user.get('username', '-')} | {self.backend.user.get('role', '-')}")
        self.refresh_current_page(force=True)
        self._apply_navigation_permissions()

