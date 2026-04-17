from __future__ import annotations

from functools import partial
import time

from PySide6.QtCore import Qt, QTimer
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
        self.pages: dict[str, QWidget] = {}
        self.nav_buttons: dict[str, QToolButton] = {}
        self._trial_block_dialog_open = False
        self._save_warning_open = False
        self._closing = False
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
        }

        self.setWindowTitle("luGEST Qt")
        self.setMinimumSize(1440, 900)
        self.resize(1600, 960)
        self.setWindowState(self.windowState() | Qt.WindowMaximized)

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
            extras_btn.clicked.connect(self._open_admin_extras)
            right_col.addWidget(extras_btn)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.clicked.connect(lambda: self.refresh_current_page(force=True, background=False))
        right_col.addWidget(refresh_btn)
        shell_layout.addLayout(right_col)
        root.addWidget(shell)

        nav_bar = QFrame()
        nav_bar.setObjectName("NavBar")
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(8, 5, 8, 5)
        nav_layout.setSpacing(5)
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
            ("pulse", "Pulse"),
            ("avarias", "Avarias"),
            ("home", "Resumo"),
        ):
            button = QToolButton()
            button.setText(label)
            button.setCheckable(True)
            button.setProperty("nav", "true")
            button.setToolButtonStyle(Qt.ToolButtonTextOnly)
            button.clicked.connect(partial(self.show_page, key))
            nav_layout.addWidget(button)
            self.nav_buttons[key] = button
        nav_layout.addStretch(1)
        root.addWidget(nav_bar)

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
        dialog.setMinimumSize(980, 720)
        dialog.resize(1080, 760)
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
        user_form.addRow("Password", password_user_edit)
        user_form.addRow("Role", role_combo)
        user_form.addRow("Posto", posto_combo)
        user_form.addRow("", active_box)
        form_host_layout.addLayout(user_form)
        password_note = QLabel("Deixa a password em branco para manter a atual quando estiveres a editar.")
        password_note.setProperty("role", "muted")
        password_note.setWordWrap(True)
        form_host_layout.addWidget(password_note)

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

        def clear_user_form() -> None:
            current_username["value"] = ""
            username_edit.setText("")
            password_user_edit.setText("")
            password_user_edit.setPlaceholderText("Obrigatoria para novo utilizador")
            role_combo.setCurrentText("Operador" if role_combo.findText("Operador") >= 0 else role_combo.currentText())
            posto_combo.setCurrentText("")
            active_box.setChecked(True)
            for key, check in permission_checks.items():
                check.setChecked(key == "home")

        def load_user_form(row: dict) -> None:
            current_username["value"] = str(row.get("username", "") or "").strip()
            username_edit.setText(current_username["value"])
            password_user_edit.setText("")
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
            payload = {
                "username": username_edit.text().strip(),
                "password": password_user_edit.text().strip(),
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
            refresh_users(current_username["value"])
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
        if trial_btn is not None:
            trial_btn.clicked.connect(open_trial_dialog)
        new_user_btn.clicked.connect(clear_user_form)
        save_user_btn.clicked.connect(on_save_user)
        remove_user_btn.clicked.connect(on_remove_user)
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

