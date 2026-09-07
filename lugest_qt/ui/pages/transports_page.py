from __future__ import annotations
from PySide6.QtCore import QDate, QTime, QTimer, QUrl, Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)
from urllib.parse import quote
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    configure_table as _configure_table,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    set_table_columns as _set_table_columns,
)
from .runtime_support import _fmt_eur
from ..widgets import (
    CardFrame,
    ClickableDateEdit as QDateEdit,
    FlexibleDecimalSpinBox as QDoubleSpinBox,
    StatCard,
)
try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
except ImportError:  # pragma: no cover - fallback for reduced workstation builds
    QWebEngineView = None  # type: ignore[assignment,misc]


class TransportsPage(QWidget):
    page_title = "Transportes"
    page_subtitle = "Planeamento de viagens e paragens para encomendas a nosso cargo, ligado a Expedição e Encomendas."
    uses_backend_reload = True
    allow_auto_timer_refresh = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.pending_rows: list[dict] = []
        self.trip_rows: list[dict] = []
        self.current_detail: dict[str, Any] = {}
        self._filter_refresh_timer = QTimer(self)
        self._filter_refresh_timer.setSingleShot(True)
        self._filter_refresh_timer.setInterval(180)
        self._filter_refresh_timer.timeout.connect(self.refresh)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        actions = CardFrame()
        actions.set_tone("info")
        actions_layout = QVBoxLayout(actions)
        actions_layout.setContentsMargins(14, 10, 14, 10)
        actions_layout.setSpacing(8)
        self.new_trip_btn = QPushButton("1 · Criar viagem")
        self.new_trip_btn.setProperty("variant", "success")
        self.new_trip_btn.clicked.connect(self._new_trip)
        self.assign_btn = QPushButton("2 · Adicionar destinos")
        self.assign_btn.clicked.connect(self._assign_selected_orders)
        self.edit_trip_btn = QPushButton("Editar viagem")
        self.edit_trip_btn.setProperty("variant", "secondary")
        self.edit_trip_btn.clicked.connect(self._edit_trip)
        self.remove_trip_btn = QPushButton("🗑 Apagar viagem")
        self.remove_trip_btn.setProperty("variant", "destructive")
        self.remove_trip_btn.clicked.connect(self._remove_trip)
        self.request_btn = QPushButton("Pedido à transportadora")
        self.request_btn.setProperty("variant", "secondary")
        self.request_btn.clicked.connect(self._request_transport)
        self.tariff_btn = QPushButton("Tarifário")
        self.tariff_btn.setProperty("variant", "secondary")
        self.tariff_btn.clicked.connect(self._manage_tariffs)
        self.apply_cost_btn = QPushButton("Aplicar custo")
        self.apply_cost_btn.setProperty("variant", "secondary")
        self.apply_cost_btn.clicked.connect(self._apply_suggested_costs)
        self.trip_status_combo = QComboBox()
        self.trip_status_combo.addItems(["Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        self.trip_status_btn = QPushButton("Aplicar à viagem")
        self.trip_status_btn.setProperty("variant", "secondary")
        self.trip_status_btn.clicked.connect(self._apply_trip_status)
        self.stop_status_combo = QComboBox()
        self.stop_status_combo.addItems(["Planeada", "Carregada", "Entregue", "Incidente"])
        self.stop_status_btn = QPushButton("Registar no destino")
        self.stop_status_btn.setProperty("variant", "secondary")
        self.stop_status_btn.clicked.connect(self._apply_stop_status)
        self.edit_stop_btn = QPushButton("Editar destino / guia")
        self.edit_stop_btn.setProperty("variant", "secondary")
        self.edit_stop_btn.clicked.connect(self._edit_stop)
        self.stop_map_btn = QPushButton("Ver destino no mapa")
        self.stop_map_btn.setProperty("variant", "secondary")
        self.stop_map_btn.clicked.connect(self._open_selected_stop_map)
        self.stop_up_btn = QPushButton("Subir")
        self.stop_up_btn.setProperty("variant", "secondary")
        self.stop_up_btn.clicked.connect(lambda: self._move_stop(-1))
        self.stop_down_btn = QPushButton("Descer")
        self.stop_down_btn.setProperty("variant", "secondary")
        self.stop_down_btn.clicked.connect(lambda: self._move_stop(1))
        self.remove_stop_btn = QPushButton("Remover paragem")
        self.remove_stop_btn.setProperty("variant", "destructive")
        self.remove_stop_btn.clicked.connect(self._remove_stop)
        self.map_route_btn = QPushButton("3 · Abrir rota")
        self.map_route_btn.setProperty("variant", "success")
        self.map_route_btn.clicked.connect(self._open_trip_route)
        self.pdf_btn = QPushButton("Documento da viagem")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_trip_pdf)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        for button in (
            self.new_trip_btn,
            self.assign_btn,
            self.edit_trip_btn,
            self.remove_trip_btn,
            self.request_btn,
            self.tariff_btn,
            self.apply_cost_btn,
            self.trip_status_btn,
            self.stop_status_btn,
            self.edit_stop_btn,
            self.stop_map_btn,
            self.stop_up_btn,
            self.stop_down_btn,
            self.remove_stop_btn,
            self.map_route_btn,
            self.pdf_btn,
            self.refresh_btn,
        ):
            button.setMinimumWidth(126)
        self.trip_status_combo.setMinimumWidth(136)
        self.stop_status_combo.setMinimumWidth(136)

        actions_hint = QLabel(
            "Uma viagem agrupa um ou vários destinos. Segue os quatro passos abaixo; "
            "o painel indica sempre o próximo passo da viagem selecionada."
        )
        actions_hint.setProperty("role", "muted")
        actions_layout.addWidget(actions_hint)

        workflow_strip = QFrame()
        workflow_strip.setObjectName("TransportWorkflow")
        workflow_strip.setStyleSheet(
            "QFrame#TransportWorkflow { background: #eef0ed; border: 1px solid #cbd0ca; border-radius: 7px; }"
        )
        workflow_strip_layout = QHBoxLayout(workflow_strip)
        workflow_strip_layout.setContentsMargins(6, 6, 6, 6)
        workflow_strip_layout.setSpacing(6)
        self.workflow_steps: list[QLabel] = []
        for number, title in (
            ("1", "Criar viagem"),
            ("2", "Adicionar destinos"),
            ("3", "Preparar e seguir rota"),
            ("4", "Confirmar entregas"),
        ):
            step = QLabel(f"{number}  {title}")
            step.setAlignment(Qt.AlignCenter)
            step.setMinimumHeight(34)
            step.setStyleSheet(
                "background: #ffffff; color: #59615a; border: 1px solid #d2d7d1; "
                "border-radius: 5px; padding: 5px 8px; font-size: 10px; font-weight: 800;"
            )
            workflow_strip_layout.addWidget(step, 1)
            self.workflow_steps.append(step)
        actions_layout.addWidget(workflow_strip)

        self.active_trip_label = QLabel("VIAGEM ATIVA · nenhuma selecionada")
        self.active_trip_label.setStyleSheet(
            "padding: 6px 10px; border-radius: 5px; background: #eaf2fb; "
            "color: #174a7c; font-size: 11px; font-weight: 800;"
        )
        actions_layout.addWidget(self.active_trip_label)
        self.workflow_status = QLabel("Começa por criar ou selecionar uma viagem.")
        self.workflow_status.setWordWrap(True)
        self.workflow_status.setStyleSheet(
            "padding: 7px 10px; border-left: 4px solid #7ed321; background: #f7f8f6; "
            "color: #30343b; font-size: 10px; font-weight: 700;"
        )
        actions_layout.addWidget(self.workflow_status)

        action_grid = QGridLayout()
        action_grid.setHorizontalSpacing(8)
        action_grid.setVerticalSpacing(8)
        action_grid.setColumnMinimumWidth(0, 96)
        action_grid.setColumnStretch(7, 1)

        planning_label = QLabel("VIAGEM")
        planning_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(planning_label, 0, 0)
        for col, widget in enumerate(
            (
                self.new_trip_btn,
                self.assign_btn,
                self.edit_trip_btn,
                self.request_btn,
                self.map_route_btn,
                self.pdf_btn,
            ),
            start=1,
        ):
            action_grid.addWidget(widget, 0, col)
        action_grid.addWidget(self.tariff_btn, 0, 8)
        action_grid.addWidget(self.refresh_btn, 0, 9)

        trip_label = QLabel("ESTADOS")
        trip_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(trip_label, 1, 0)
        action_grid.addWidget(self.trip_status_combo, 1, 1)
        action_grid.addWidget(self.trip_status_btn, 1, 2)
        action_grid.addWidget(self.stop_status_combo, 1, 3)
        action_grid.addWidget(self.stop_status_btn, 1, 4)
        action_grid.addWidget(self.apply_cost_btn, 1, 5)
        action_grid.addWidget(self.remove_trip_btn, 1, 9)

        stop_label = QLabel("DESTINO")
        stop_label.setStyleSheet("font-size: 11px; font-weight: 800; color: #475569; text-transform: uppercase;")
        action_grid.addWidget(stop_label, 2, 0)
        for col, widget in enumerate(
            (
                self.edit_stop_btn,
                self.stop_map_btn,
                self.stop_up_btn,
                self.stop_down_btn,
                self.remove_stop_btn,
            ),
            start=1,
        ):
            action_grid.addWidget(widget, 2, col)

        actions_layout.addLayout(action_grid)
        root.addWidget(actions)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por viagem, encomenda, cliente, matrícula ou motorista")
        self.filter_edit.lineEdit().textChanged.connect(lambda _text: self._filter_refresh_timer.start())
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Todas", "Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.state_combo)
        root.addWidget(filters)

        stats_host = QWidget()
        stats_layout = QHBoxLayout(stats_host)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setSpacing(8)
        self.transport_stats = [
            StatCard("Viagens visiveis"),
            StatCard("Em curso"),
            StatCard("Incidentes"),
            StatCard("Paragens"),
            StatCard("Carga planeada"),
            StatCard("Entregas concluidas"),
        ]
        for index, card in enumerate(self.transport_stats):
            card.set_tone(("info", "warning", "danger", "default", "info", "success")[index])
            card.setMinimumWidth(150)
            stats_layout.addWidget(card, 1)
        root.addWidget(stats_host)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)
        self.transport_tabs = QTabWidget()
        self.transport_tabs.setDocumentMode(True)

        pending_card = CardFrame()
        pending_card.set_tone("warning")
        pending_layout = QVBoxLayout(pending_card)
        pending_layout.setContentsMargins(14, 12, 14, 12)
        pending_layout.setSpacing(8)
        pending_title = QLabel("Destinos por planear")
        pending_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        pending_hint = QLabel("Encomendas prontas, com destino, carga e responsabilidade de transporte identificados.")
        pending_hint.setProperty("role", "muted")
        pending_hint.setWordWrap(True)
        self.pending_table = QTableWidget(0, 11)
        self.pending_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Entrega", "Zona", "Tipo", "Pal", "Peso kg", "Descarga", "GPS", "Guia", "Disponível"])
        self.pending_table.verticalHeader().setVisible(False)
        self.pending_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pending_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pending_table.setSelectionMode(QAbstractItemView.MultiSelection)
        _configure_table(self.pending_table, stretch=(1, 3), contents=())
        _set_table_columns(
            self.pending_table,
            [
                (0, "interactive", 132),
                (1, "stretch", 190),
                (2, "interactive", 96),
                (3, "interactive", 140),
                (4, "interactive", 150),
                (5, "interactive", 64),
                (6, "interactive", 80),
                (7, "stretch", 220),
                (8, "interactive", 92),
                (9, "interactive", 112),
                (10, "interactive", 86),
            ],
        )
        self.pending_table.verticalHeader().setDefaultSectionSize(30)
        self.pending_table.itemSelectionChanged.connect(self._sync_actions)
        self.pending_empty = QLabel("Sem encomendas a nosso cargo prontas para agendar neste momento.")
        self.pending_empty.setProperty("role", "muted")
        self.pending_empty.setVisible(False)
        pending_layout.addWidget(pending_title)
        pending_layout.addWidget(pending_hint)
        pending_layout.addWidget(self.pending_table)
        pending_layout.addWidget(self.pending_empty)
        self.transport_tabs.addTab(pending_card, "Por agendar")

        trips_card = CardFrame()
        trips_card.set_tone("default")
        trips_layout = QVBoxLayout(trips_card)
        trips_layout.setContentsMargins(14, 12, 14, 12)
        trips_layout.setSpacing(8)
        trips_title = QLabel("Plano de transportes")
        trips_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        trips_hint = QLabel("Viagens próprias e serviços de transportadora, acompanhados até à prova de entrega.")
        trips_hint.setProperty("role", "muted")
        trips_hint.setWordWrap(True)
        self.trip_table = QTableWidget(0, 9)
        self.trip_table.setHorizontalHeaderLabels(["Viagem", "Data", "Saída", "Tipo", "Estado", "Pedido", "Parceiro / Viatura", "Pal", "Paragens"])
        self.trip_table.verticalHeader().setVisible(False)
        self.trip_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.trip_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.trip_table, stretch=(4, 5), contents=())
        _set_table_columns(
            self.trip_table,
            [
                (0, "interactive", 120),
                (1, "interactive", 96),
                (2, "interactive", 86),
                (3, "interactive", 128),
                (4, "interactive", 110),
                (5, "interactive", 118),
                (6, "stretch", 210),
                (7, "interactive", 64),
                (8, "interactive", 82),
            ],
        )
        self.trip_table.verticalHeader().setDefaultSectionSize(30)
        self.trip_table.itemSelectionChanged.connect(self._show_trip_detail)
        self.trip_table.itemSelectionChanged.connect(self._sync_actions)
        self.trip_empty = QLabel("Sem viagens registadas para o filtro atual.")
        self.trip_empty.setProperty("role", "muted")
        self.trip_empty.setVisible(False)
        trips_layout.addWidget(trips_title)
        trips_layout.addWidget(trips_hint)
        trips_layout.addWidget(self.trip_table)
        trips_layout.addWidget(self.trip_empty)
        self.transport_tabs.addTab(trips_card, "Viagens")
        splitter.addWidget(self.transport_tabs)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        self.detail_card = CardFrame()
        self.detail_card.set_tone("info")
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(14, 12, 14, 12)
        detail_layout.setSpacing(8)
        header = QHBoxLayout()
        self.detail_title = QLabel("Seleciona uma viagem")
        self.detail_title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.detail_state_chip = QLabel("-")
        _apply_state_chip(self.detail_state_chip, "-")
        header.addWidget(self.detail_title, 1)
        header.addWidget(self.detail_state_chip)
        detail_layout.addLayout(header)
        self.detail_meta = QLabel("Cria uma viagem e afeta encomendas a nosso cargo.")
        self.detail_meta.setWordWrap(True)
        self.detail_meta.setProperty("role", "muted")
        detail_layout.addWidget(self.detail_meta)
        self.detail_note = QLabel("-")
        self.detail_note.setWordWrap(True)
        detail_layout.addWidget(self.detail_note)
        self.stops_table = QTableWidget(0, 14)
        self.stops_table.setHorizontalHeaderLabels(["Ord", "Encomenda", "Cliente", "Zona", "Pal", "Peso", "Vol", "Descarga", "GPS", "Planeado", "Guia", "Checklist", "POD", "Estado"])
        self.stops_table.verticalHeader().setVisible(False)
        self.stops_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.stops_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.stops_table, stretch=(2, 3), contents=())
        _set_table_columns(
            self.stops_table,
            [
                (0, "interactive", 54),
                (1, "interactive", 126),
                (2, "stretch", 170),
                (3, "interactive", 120),
                (4, "interactive", 54),
                (5, "interactive", 78),
                (6, "interactive", 70),
                (7, "stretch", 220),
                (8, "interactive", 88),
                (9, "interactive", 126),
                (10, "interactive", 96),
                (11, "interactive", 86),
                (12, "interactive", 96),
                (13, "interactive", 104),
            ],
        )
        self.stops_table.verticalHeader().setDefaultSectionSize(30)
        self.stops_table.itemSelectionChanged.connect(self._sync_actions)
        self.stops_table.itemSelectionChanged.connect(self._refresh_embedded_map_if_visible)
        self.stops_table.itemDoubleClicked.connect(lambda *_args: self._edit_stop())
        detail_layout.addWidget(self.stops_table, 1)

        self.route_tabs = QTabWidget()
        self.route_tabs.setDocumentMode(True)
        self.route_tabs.addTab(self.detail_card, "Rota e entregas")

        map_page = QWidget()
        map_layout = QVBoxLayout(map_page)
        map_layout.setContentsMargins(10, 10, 10, 10)
        map_layout.setSpacing(8)
        map_toolbar = QHBoxLayout()
        map_identity = QVBoxLayout()
        map_identity.setSpacing(0)
        map_title = QLabel("Mapa da viagem")
        map_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #20251f;")
        self.map_context_label = QLabel("Seleciona uma viagem para visualizar a rota.")
        self.map_context_label.setProperty("role", "muted")
        self.map_context_label.setWordWrap(True)
        map_identity.addWidget(map_title)
        map_identity.addWidget(self.map_context_label)
        self.map_external_btn = QPushButton("Abrir no Google Maps")
        self.map_external_btn.setProperty("variant", "success")
        self.map_external_btn.clicked.connect(self._open_trip_route)
        map_toolbar.addLayout(map_identity, 1)
        map_toolbar.addWidget(self.map_external_btn)
        map_layout.addLayout(map_toolbar)
        self.map_host = QFrame()
        self.map_host.setObjectName("TransportMapHost")
        self.map_host.setStyleSheet(
            "QFrame#TransportMapHost { background: #e7e9e5; border: 1px solid #bdc3bc; border-radius: 7px; }"
        )
        self.map_host_layout = QVBoxLayout(self.map_host)
        self.map_host_layout.setContentsMargins(1, 1, 1, 1)
        self.map_placeholder = QLabel(
            "O mapa é carregado apenas quando abres este separador.\n"
            "É necessária ligação à Internet."
        )
        self.map_placeholder.setAlignment(Qt.AlignCenter)
        self.map_placeholder.setWordWrap(True)
        self.map_placeholder.setStyleSheet("color: #59615a; font-size: 11px; font-weight: 700;")
        self.map_host_layout.addWidget(self.map_placeholder)
        self.map_view = None
        map_layout.addWidget(self.map_host, 1)
        self.route_tabs.addTab(map_page, "Mapa integrado")
        self.route_tabs.currentChanged.connect(self._on_route_tab_changed)
        right_layout.addWidget(self.route_tabs, 1)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 7)
        splitter.setSizes([720, 1040])
        self._clear_trip_detail()
        self._sync_actions()

    def refresh(self) -> None:
        query = self.filter_edit.currentText().strip()
        previous_trip = str(self.current_detail.get("numero", "") or "").strip()
        self.pending_rows = self.backend.transport_pending_orders(query)
        pending_overview = dict(self.backend.transport_pending_overview() or {})
        self.trip_rows = self.backend.transport_rows(query, self.state_combo.currentText())
        self._refresh_filter_options(query)
        _fill_table(
            self.pending_table,
            [
                [
                    row.get("numero", "-"),
                    row.get("cliente", "-"),
                    row.get("data_entrega", "-"),
                    row.get("zona_transporte", "-"),
                    row.get("nota_transporte", "-"),
                    f"{float(row.get('paletes', 0) or 0):.2f}",
                    f"{float(row.get('peso_bruto_kg', 0) or 0):.1f}",
                    row.get("local_descarga", "-"),
                    "Definido" if str(row.get("latitude", "") or "").strip() and str(row.get("longitude", "") or "").strip() else "Morada",
                    row.get("guia_numero", "-"),
                    f"{float(row.get('disponivel', 0) or 0):.1f}",
                ]
                for row in self.pending_rows
            ],
            align_center_from=2,
        )
        self.pending_empty.setVisible(self.pending_table.rowCount() == 0)
        if not self.pending_rows:
            waiting = int(pending_overview.get("waiting_stock_or_guide", 0) or 0)
            customer = int(pending_overview.get("customer_transport", 0) or 0)
            assigned = int(pending_overview.get("already_assigned", 0) or 0)
            if query:
                empty_text = "Nenhum destino corresponde à pesquisa atual."
            elif waiting:
                empty_text = (
                    f"Sem destinos disponíveis. {waiting} encomenda(s) a nosso cargo ainda não têm "
                    "peças disponíveis para expedição nem guia emitida."
                )
            else:
                empty_text = (
                    "Sem destinos disponíveis para planear. "
                    f"{customer} encomenda(s) estão a cargo do cliente e {assigned} já estão noutra viagem."
                )
            self.pending_empty.setText(empty_text)
        _fill_table(
            self.trip_table,
            [
                [
                    row.get("numero", "-"),
                    row.get("data_planeada", "-"),
                    row.get("hora_saida", "-"),
                    row.get("tipo_responsavel", "-"),
                    row.get("estado", "-"),
                    row.get("pedido_transporte_estado", "-"),
                    row.get("transportadora_nome", "-") if "subcontrat" in str(row.get("tipo_responsavel", "")).lower() else row.get("viatura", "-"),
                    f"{float(row.get('paletes', 0) or 0):.2f}",
                    row.get("paragens", 0),
                ]
                for row in self.trip_rows
            ],
            align_center_from=6,
        )
        for row_index, row in enumerate(self.trip_rows):
            _paint_table_row(self.trip_table, row_index, str(row.get("estado", "")))
        in_progress = sum(
            1
            for row in self.trip_rows
            if any(token in self.backend.desktop_main.norm_text(str(row.get("estado", "") or "")) for token in ("carga", "transito"))
        )
        incidents = sum(1 for row in self.trip_rows if "inciden" in self.backend.desktop_main.norm_text(str(row.get("estado", "") or "")))
        stops_total = sum(int(row.get("paragens", 0) or 0) for row in self.trip_rows)
        delivered_total = sum(int(row.get("entregues", 0) or 0) for row in self.trip_rows)
        pallets_total = sum(float(row.get("paletes", 0) or 0) for row in self.trip_rows)
        weight_total = sum(float(row.get("peso_bruto_kg", 0) or 0) for row in self.trip_rows)
        stat_values = [
            (str(len(self.trip_rows)), f"{len(self.pending_rows)} encomendas por agendar"),
            (str(in_progress), "em carga ou em transito"),
            (str(incidents), "exigem acompanhamento"),
            (str(stops_total), "destinos planeados"),
            (f"{pallets_total:.1f} pal", f"{weight_total:.0f} kg"),
            (str(delivered_total), f"de {stops_total} paragens"),
        ]
        for card, (value, subtitle) in zip(self.transport_stats, stat_values):
            card.set_data(value, subtitle)
        self.trip_empty.setVisible(self.trip_table.rowCount() == 0)
        self.transport_tabs.setTabText(0, f"Por agendar ({len(self.pending_rows)})")
        self.transport_tabs.setTabText(1, f"Viagens ({len(self.trip_rows)})")
        self._restore_trip_selection(previous_trip)
        if self.trip_table.rowCount() == 0:
            self._clear_trip_detail()
        else:
            self._show_trip_detail()
        self._sync_actions()

    def _refresh_filter_options(self, current_text: str) -> None:
        values: list[str] = []
        for row in list(self.pending_rows) + list(self.trip_rows):
            for key in ("numero", "cliente", "local_descarga", "zona_transporte", "viatura", "motorista", "matricula", "transportadora_nome"):
                value = str(row.get(key, "") or "").strip()
                if value and value not in values:
                    values.append(value)
        block = self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.addItem("")
        for value in values:
            self.filter_edit.addItem(value)
        self.filter_edit.setCurrentText(current_text)
        self.filter_edit.blockSignals(block)

    def _restore_trip_selection(self, numero: str) -> None:
        if self.trip_table.rowCount() == 0:
            return
        row_index = 0
        if numero:
            for index, row in enumerate(self.trip_rows):
                if str(row.get("numero", "") or "").strip() == numero:
                    row_index = index
                    break
        self.trip_table.selectRow(row_index)

    def open_trip_numero(self, numero: str) -> None:
        target = str(numero or "").strip()
        if not target:
            return
        self.refresh()
        for row_index, row in enumerate(self.trip_rows):
            if str(row.get("numero", "") or "").strip() != target:
                continue
            self.trip_table.selectRow(row_index)
            self._show_trip_detail()
            self._sync_actions()
            return

    def _selected_pending_rows(self) -> list[dict]:
        indexes = sorted({item.row() for item in self.pending_table.selectedItems()})
        return [self.pending_rows[index] for index in indexes if 0 <= index < len(self.pending_rows)]

    def _current_trip_row(self) -> dict:
        current = self.trip_table.currentItem()
        if current is None or current.row() >= len(self.trip_rows):
            return {}
        return self.trip_rows[current.row()]

    def _current_stop_row(self) -> dict:
        current = self.stops_table.currentItem()
        stops = list(self.current_detail.get("paragens", []) or [])
        if current is None or current.row() >= len(stops):
            return {}
        return stops[current.row()]

    def _clear_trip_detail(self) -> None:
        self.current_detail = {}
        self.active_trip_label.setText("VIAGEM ATIVA · nenhuma selecionada")
        self.detail_title.setText("Seleciona uma viagem")
        self.detail_meta.setText("Cria uma viagem e afeta encomendas a nosso cargo.")
        self.detail_note.setText("-")
        _apply_state_chip(self.detail_state_chip, "-")
        self.stops_table.setRowCount(0)
        if hasattr(self, "route_tabs"):
            self._refresh_embedded_map_if_visible()

    def _show_trip_detail(self) -> None:
        row = self._current_trip_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            self._clear_trip_detail()
            self._sync_actions()
            return
        try:
            detail = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.current_detail = detail
        self.active_trip_label.setText(f"VIAGEM ATIVA · {detail.get('numero', '-')} · {detail.get('estado', '-')}")
        self.detail_title.setText(f"Viagem {detail.get('numero', '-')}")
        _apply_state_chip(self.detail_state_chip, str(detail.get("estado", "") or "-"))
        self.trip_status_combo.setCurrentText(str(detail.get("estado", "Planeado") or "Planeado"))
        transportadora_txt = str(detail.get("transportadora_nome", "") or "-").strip() or "-"
        pedido_state = str(detail.get("pedido_transporte_estado", "Nao pedido") or "Nao pedido").strip() or "Nao pedido"
        pedido_meta = pedido_state
        if detail.get("pedido_confirmado_at"):
            pedido_meta += f" confirmado em {detail.get('pedido_confirmado_at', '-')}"
        elif detail.get("pedido_recusado_at"):
            pedido_meta += f" recusado em {detail.get('pedido_recusado_at', '-')}"
        elif detail.get("pedido_transporte_at"):
            pedido_meta += f" em {detail.get('pedido_transporte_at', '-')}"
        if detail.get("pedido_transporte_by"):
            pedido_meta += f" por {detail.get('pedido_transporte_by', '-')}"
        self.detail_meta.setText(
            f"Planeado {detail.get('data_planeada', '-') or '-'} às {detail.get('hora_saida', '-') or '-'} | "
            f"Tipo {detail.get('tipo_responsavel', '-') or '-'}\n"
            f"Viatura {detail.get('viatura', '-') or '-'} | Matrícula {detail.get('matricula', '-') or '-'} | "
            f"Motorista {detail.get('motorista', '-') or '-'} | Telefone {detail.get('telefone_motorista', '-') or '-'}"
        )
        carga_txt = (
            f"{float(detail.get('paletes', 0) or 0):.2f} pal | "
            f"{float(detail.get('peso_bruto_kg', 0) or 0):.1f} kg | "
            f"{float(detail.get('volume_m3', 0) or 0):.3f} m3"
        )
        if bool(detail.get("carga_manual")):
            carga_txt += (
                f" (manual; calc. {float(detail.get('paletes_calculadas', 0) or 0):.2f} pal / "
                f"{float(detail.get('peso_bruto_kg_calculado', 0) or 0):.1f} kg / "
                f"{float(detail.get('volume_m3_calculado', 0) or 0):.3f} m3)"
            )
        self.detail_note.setText(
            f"Origem: {detail.get('origem', '-') or '-'} | Transportadora / fornecedor: {transportadora_txt} | "
            f"Ref. externa: {detail.get('referencia_transporte', '-') or '-'}\n"
            f"Pedido transporte: {pedido_meta} | Ref. pedido: {detail.get('pedido_transporte_ref', '-') or '-'}\n"
            f"Checklist OK: {int(detail.get('checklist_ok', 0) or 0)} | POD recebidos: {int(detail.get('pod_recebidos', 0) or 0)}\n"
            f"Zonas: {', '.join(list(detail.get('zonas', []) or [])) or '-'} | "
            f"Custo sugerido {_fmt_eur(float(detail.get('custo_sugerido_total', 0) or 0))}\n"
            f"Carga {carga_txt} | Preço {_fmt_eur(float(detail.get('preco_total', 0) or 0))} | "
            f"Custo {_fmt_eur(float(detail.get('custo_total', 0) or 0))}"
        )
        _fill_table(
            self.stops_table,
            [
                [
                    stop.get("ordem", "-"),
                    stop.get("encomenda_numero", "-"),
                    stop.get("cliente_nome", "-"),
                    stop.get("zona_transporte", "-"),
                    f"{float(stop.get('paletes', 0) or 0):.2f}",
                    f"{float(stop.get('peso_bruto_kg', 0) or 0):.1f}",
                    f"{float(stop.get('volume_m3', 0) or 0):.3f}",
                    stop.get("local_descarga", "-"),
                    "Definido" if str(stop.get("latitude", "") or "").strip() and str(stop.get("longitude", "") or "").strip() else "Morada",
                    stop.get("data_planeada", "-"),
                    stop.get("guia_numero", "-"),
                    stop.get("checklist_estado", "-"),
                    stop.get("pod_estado", "-"),
                    stop.get("estado", "-"),
                ]
                for stop in list(detail.get("paragens", []) or [])
            ],
            align_center_from=0,
        )
        for row_index, stop in enumerate(list(detail.get("paragens", []) or [])):
            _paint_table_row(self.stops_table, row_index, str(stop.get("estado", "")))
        self._refresh_embedded_map_if_visible()
        self._sync_actions()

    def _sync_actions(self) -> None:
        has_trip = bool(self._current_trip_row())
        has_pending = bool(self._selected_pending_rows())
        has_stop = bool(self._current_stop_row())
        self.assign_btn.setEnabled(has_trip and has_pending)
        self.edit_trip_btn.setEnabled(has_trip)
        self.request_btn.setEnabled(has_trip)
        self.tariff_btn.setEnabled(True)
        self.apply_cost_btn.setEnabled(has_trip)
        self.remove_trip_btn.setEnabled(has_trip)
        self.trip_status_combo.setEnabled(has_trip)
        self.trip_status_btn.setEnabled(has_trip)
        self.stop_status_combo.setEnabled(has_stop)
        self.stop_status_btn.setEnabled(has_stop)
        self.edit_stop_btn.setEnabled(has_stop)
        self.stop_map_btn.setEnabled(has_stop)
        self.stop_up_btn.setEnabled(has_stop)
        self.stop_down_btn.setEnabled(has_stop)
        self.remove_stop_btn.setEnabled(has_stop)
        self.pdf_btn.setEnabled(has_trip)
        self.map_route_btn.setEnabled(has_trip and bool(list(self.current_detail.get("paragens", []) or [])))
        self.map_external_btn.setEnabled(has_trip and bool(list(self.current_detail.get("paragens", []) or [])))
        self._sync_workflow()

    def _sync_workflow(self) -> None:
        detail = dict(self.current_detail or {})
        stops = [dict(row or {}) for row in list(detail.get("paragens", []) or [])]
        state = self.backend.desktop_main.norm_text(str(detail.get("estado", "") or ""))
        selected_pending = len(self._selected_pending_rows())
        if not detail:
            active_step = 0
            status = "Passo 1: cria uma viagem própria ou subcontratada. Depois adiciona os destinos."
        elif not stops:
            active_step = 1
            if self.pending_rows:
                selection = (
                    f" Tens {selected_pending} destino(s) selecionado(s)."
                    if selected_pending
                    else " Seleciona uma ou mais linhas em «Por agendar»."
                )
                status = f"Passo 2: esta viagem está vazia.{selection} Depois carrega em «Adicionar destinos»."
            else:
                status = (
                    "Passo 2 bloqueado: não existem encomendas prontas para transporte. "
                    "A encomenda tem de estar a nosso cargo e ter peças disponíveis ou uma guia emitida."
                )
        elif "conclu" in state:
            active_step = 3
            status = "Fluxo concluído: todos os destinos foram entregues e possuem prova de entrega."
        elif any(token in state for token in ("transito", "inciden")):
            active_step = 3
            pending_delivery = sum(
                1
                for stop in stops
                if "entreg" not in self.backend.desktop_main.norm_text(str(stop.get("estado", "") or ""))
            )
            status = (
                f"Passo 4: regista cada entrega e o respetivo POD. "
                f"Faltam {pending_delivery} de {len(stops)} destino(s)."
            )
        else:
            active_step = 2
            checklist_ok = sum(1 for stop in stops if str(stop.get("checklist_estado", "") or "") == "OK")
            status = (
                f"Passo 3: confirma guia, carga, documentos e paletes em cada destino "
                f"({checklist_ok}/{len(stops)} preparados); ordena a rota e inicia a viagem."
            )
        completed_until = max(0, active_step)
        for index, step in enumerate(self.workflow_steps):
            if index < completed_until:
                style = (
                    "background: #e8f4dc; color: #31511d; border: 1px solid #a8c88b; "
                    "border-radius: 5px; padding: 5px 8px; font-size: 10px; font-weight: 800;"
                )
            elif index == active_step:
                style = (
                    "background: #454945; color: #ffffff; border: 1px solid #454945; "
                    "border-radius: 5px; padding: 5px 8px; font-size: 10px; font-weight: 900;"
                )
            else:
                style = (
                    "background: #ffffff; color: #7a817a; border: 1px solid #d2d7d1; "
                    "border-radius: 5px; padding: 5px 8px; font-size: 10px; font-weight: 800;"
                )
            step.setStyleSheet(style)
        self.workflow_status.setText(status)

    @staticmethod
    def _transport_map_location(row: dict, address_key: str, lat_key: str, lon_key: str) -> str:
        latitude = str(row.get(lat_key, "") or "").strip().replace(",", ".")
        longitude = str(row.get(lon_key, "") or "").strip().replace(",", ".")
        try:
            if latitude and longitude:
                float(latitude)
                float(longitude)
                return f"{latitude},{longitude}"
        except ValueError:
            pass
        return str(row.get(address_key, "") or "").strip()

    def _trip_route_url(self, detail: dict | None = None) -> str:
        current = dict(detail or self.current_detail or {})
        stops = [dict(row or {}) for row in list(current.get("paragens", []) or [])]
        destinations = [
            self._transport_map_location(stop, "local_descarga", "latitude", "longitude")
            for stop in stops
        ]
        destinations = [value for value in destinations if value]
        if not destinations:
            return ""
        origin = self._transport_map_location(
            current,
            "origem",
            "origem_latitude",
            "origem_longitude",
        )
        params = ["api=1", f"destination={quote(destinations[-1])}", "travelmode=driving"]
        if origin:
            params.insert(1, f"origin={quote(origin)}")
        if len(destinations) > 1:
            params.append(f"waypoints={quote('|'.join(destinations[:-1]))}")
        return "https://www.google.com/maps/dir/?" + "&".join(params)

    def _ensure_map_view(self) -> None:
        if self.map_view is not None or QWebEngineView is None:
            return
        self.map_placeholder.hide()
        self.map_view = QWebEngineView(self.map_host)
        self.map_view.setContextMenuPolicy(Qt.DefaultContextMenu)
        self.map_host_layout.addWidget(self.map_view, 1)

    def _on_route_tab_changed(self, index: int) -> None:
        if index != 1:
            return
        self._ensure_map_view()
        self._refresh_embedded_map()

    def _refresh_embedded_map_if_visible(self) -> None:
        if self.route_tabs.currentIndex() == 1:
            self._refresh_embedded_map()

    def _refresh_embedded_map(self) -> None:
        detail = dict(self.current_detail or {})
        stops = [dict(row or {}) for row in list(detail.get("paragens", []) or [])]
        selected_stop = dict(self._current_stop_row() or {})
        if selected_stop:
            location = self._transport_map_location(
                selected_stop,
                "local_descarga",
                "latitude",
                "longitude",
            )
            map_url = (
                f"https://www.google.com/maps/search/?api=1&query={quote(location)}"
                if location
                else ""
            )
            self.map_context_label.setText(
                f"Destino selecionado: {selected_stop.get('cliente_nome', '-') or '-'} · "
                f"{selected_stop.get('local_descarga', '-') or '-'}"
            )
        else:
            map_url = self._trip_route_url(detail)
            self.map_context_label.setText(
                f"Rota completa de {detail.get('numero', '-') or '-'} · {len(stops)} destino(s)"
                if detail
                else "Seleciona uma viagem para visualizar a rota."
            )
        self.map_external_btn.setEnabled(bool(map_url))
        if QWebEngineView is None:
            self.map_placeholder.setText(
                "O componente de mapa não está disponível neste posto.\n"
                "Usa «Abrir no Google Maps» para consultar a rota no navegador."
            )
            self.map_placeholder.show()
            return
        self._ensure_map_view()
        if self.map_view is None:
            return
        if map_url:
            self.map_view.setUrl(QUrl(map_url))
        else:
            self.map_view.setHtml(
                "<html><body style='margin:0;background:#e7e9e5;color:#59615a;"
                "font-family:Segoe UI;display:flex;align-items:center;justify-content:center;"
                "text-align:center'><div><h3>Mapa sem destinos</h3>"
                "<p>Adiciona uma encomenda à viagem e confirma a morada ou as coordenadas.</p>"
                "</div></body></html>"
            )

    def _supplier_options(self, initial: dict | None = None) -> list[str]:
        options = [str(value or "").strip() for value in list((initial or {}).get("supplier_options", []) or []) if str(value or "").strip()]
        if options:
            return options
        try:
            return [
                f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -")
                for row in list(self.backend.ne_suppliers() or [])
                if str(row.get("id", "") or "").strip() or str(row.get("nome", "") or "").strip()
            ]
        except Exception:
            return []

    def _zone_options(self, initial: dict | None = None) -> list[str]:
        options = [str(value or "").strip() for value in list((initial or {}).get("zone_options", []) or []) if str(value or "").strip()]
        if options:
            return options
        try:
            return [str(value or "").strip() for value in list(self.backend.transport_zone_options() or []) if str(value or "").strip()]
        except Exception:
            return []

    def _trip_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or self.backend.transport_defaults() or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Viagem de transporte")
        dialog.setMinimumWidth(860)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        numero_label = QLabel(str(initial.get("numero", "(nova)") or "(nova)"))
        data_edit = QDateEdit()
        data_edit.setCalendarPopup(True)
        data_edit.setDisplayFormat("dd/MM/yyyy")
        raw_date = str(initial.get("data_planeada", "") or "").strip()
        qdate = QDate.fromString(raw_date, "yyyy-MM-dd") if raw_date else QDate.currentDate()
        if not qdate.isValid():
            qdate = QDate.currentDate()
        data_edit.setDate(qdate)
        hora_edit = QTimeEdit()
        hora_edit.setDisplayFormat("HH:mm")
        raw_time = str(initial.get("hora_saida", "") or "08:00").strip()
        qtime = QTime.fromString(raw_time, "HH:mm")
        if not qtime.isValid():
            qtime = QTime(8, 0)
        hora_edit.setTime(qtime)
        tipo_combo = QComboBox()
        tipo_combo.addItems(["Nosso Cargo", "Subcontratado"])
        tipo_combo.setCurrentText(str(initial.get("tipo_responsavel", "Nosso Cargo") or "Nosso Cargo"))
        estado_combo = QComboBox()
        estado_combo.addItems(["Planeado", "Em carga", "Em trânsito", "Concluído", "Incidente", "Anulado"])
        estado_combo.setCurrentText(str(initial.get("estado", "Planeado") or "Planeado"))
        viatura_combo = QComboBox()
        viatura_combo.setEditable(True)
        viatura_combo.addItem("")
        for value in list(initial.get("vehicle_options", []) or []):
            viatura_combo.addItem(str(value))
        viatura_combo.setCurrentText(str(initial.get("viatura", "") or "").strip())
        matricula_edit = QLineEdit(str(initial.get("matricula", "") or "").strip())
        motorista_combo = QComboBox()
        motorista_combo.setEditable(True)
        motorista_combo.addItem("")
        for value in list(initial.get("driver_options", []) or []):
            motorista_combo.addItem(str(value))
        motorista_combo.setCurrentText(str(initial.get("motorista", "") or "").strip())
        telefone_edit = QLineEdit(str(initial.get("telefone_motorista", "") or "").strip())
        origem_edit = QLineEdit(str(initial.get("origem", "") or "").strip())
        origem_latitude_edit = QLineEdit(str(initial.get("origem_latitude", "") or "").strip())
        origem_latitude_edit.setPlaceholderText("ex.: 41.1579")
        origem_longitude_edit = QLineEdit(str(initial.get("origem_longitude", "") or "").strip())
        origem_longitude_edit.setPlaceholderText("ex.: -8.6291")
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )

        def _sync_transport_mode() -> None:
            outsourced = "subcontrat" in tipo_combo.currentText().strip().lower()
            carrier_combo.setEnabled(outsourced)
            ref_edit.setEnabled(outsourced)
            viatura_combo.setEnabled(not outsourced)
            matricula_edit.setEnabled(not outsourced)
            motorista_combo.setEnabled(not outsourced)
            telefone_edit.setEnabled(not outsourced)
            carrier_combo.setToolTip(
                "Transportadora responsável pelo serviço externo."
                if outsourced
                else "Disponível quando o transporte é subcontratado."
            )

        tipo_combo.currentTextChanged.connect(lambda _text: _sync_transport_mode())
        ref_edit = QLineEdit(str(initial.get("referencia_transporte", "") or "").strip())
        cost_spin = QDoubleSpinBox()
        cost_spin.setRange(0.0, 1000000.0)
        cost_spin.setDecimals(2)
        cost_spin.setPrefix("EUR ")
        cost_spin.setValue(float(initial.get("custo_previsto", 0) or 0))
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(96)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        _sync_transport_mode()
        fields = [
            ("Número", numero_label, 0, 0),
            ("Data", data_edit, 0, 2),
            ("Saída", hora_edit, 0, 4),
            ("Tipo", tipo_combo, 1, 0),
            ("Estado", estado_combo, 1, 2),
            ("Transportadora", carrier_combo, 1, 4),
            ("Viatura", viatura_combo, 2, 0),
            ("Matrícula", matricula_edit, 2, 2),
            ("Motorista", motorista_combo, 2, 4),
            ("Telefone", telefone_edit, 3, 0),
            ("Origem", origem_edit, 3, 2),
            ("Latitude origem", origem_latitude_edit, 3, 4),
            ("Longitude origem", origem_longitude_edit, 4, 4),
            ("Ref. externa", ref_edit, 4, 0),
            ("Custo previsto", cost_spin, 4, 2),
        ]
        for label_text, widget, row, col in fields:
            form.addWidget(QLabel(label_text), row, col)
            span = 1
            form.addWidget(widget, row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 5, 0)
        form.addWidget(obs_edit, 5, 1, 1, 5)
        layout.addLayout(form)
        buttons_row = QHBoxLayout()
        if str(initial.get("numero", "") or "").strip():
            remove_btn = QPushButton("Apagar viagem")
            remove_btn.setProperty("variant", "destructive")

            def _confirm_remove() -> None:
                if (
                    QMessageBox.question(
                        dialog,
                        "Transportes",
                        f"Apagar a viagem {str(initial.get('numero', '') or '').strip()} e libertar as encomendas associadas?",
                    )
                    != QMessageBox.Yes
                ):
                    return
                dialog.done(2)

            remove_btn.clicked.connect(_confirm_remove)
            buttons_row.addWidget(remove_btn)
        else:
            buttons_row.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        buttons_row.addWidget(buttons)
        layout.addLayout(buttons_row)
        result = dialog.exec()
        if result == 2:
            return {"_delete_trip": True, "numero": str(initial.get("numero", "") or "").strip()}
        if result != QDialog.Accepted:
            return None
        return {
            "numero": str(initial.get("numero", "") or "").strip(),
            "tipo_responsavel": tipo_combo.currentText().strip(),
            "estado": estado_combo.currentText().strip(),
            "data_planeada": data_edit.date().toString("yyyy-MM-dd"),
            "hora_saida": hora_edit.time().toString("HH:mm"),
            "viatura": viatura_combo.currentText().strip(),
            "matricula": matricula_edit.text().strip(),
            "motorista": motorista_combo.currentText().strip(),
            "telefone_motorista": telefone_edit.text().strip(),
            "origem": origem_edit.text().strip(),
            "origem_latitude": origem_latitude_edit.text().strip(),
            "origem_longitude": origem_longitude_edit.text().strip(),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "referencia_transporte": ref_edit.text().strip(),
            "custo_previsto": cost_spin.value(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _request_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Requisitar transporte")
        dialog.setMinimumWidth(620)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        estado_combo = QComboBox()
        estado_combo.addItems(["Nao pedido", "Pedido enviado", "Confirmado", "Recusado"])
        estado_combo.setCurrentText(str(initial.get("pedido_transporte_estado", "Pedido enviado") or "Pedido enviado"))
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )
        ref_edit = QLineEdit(str(initial.get("pedido_transporte_ref", "") or "").strip())
        paletes_spin = QDoubleSpinBox()
        paletes_spin.setRange(0.0, 9999.0)
        paletes_spin.setDecimals(2)
        paletes_spin.setSuffix(" pal")
        paletes_spin.setValue(float(initial.get("paletes_total_manual", 0) or 0))
        peso_spin = QDoubleSpinBox()
        peso_spin.setRange(0.0, 100000.0)
        peso_spin.setDecimals(2)
        peso_spin.setSuffix(" kg")
        peso_spin.setValue(float(initial.get("peso_total_manual_kg", 0) or 0))
        volume_spin = QDoubleSpinBox()
        volume_spin.setRange(0.0, 10000.0)
        volume_spin.setDecimals(3)
        volume_spin.setSuffix(" m3")
        volume_spin.setValue(float(initial.get("volume_total_manual_m3", 0) or 0))
        cost_spin = QDoubleSpinBox()
        cost_spin.setRange(0.0, 1000000.0)
        cost_spin.setDecimals(2)
        cost_spin.setPrefix("EUR ")
        cost_spin.setValue(float(initial.get("custo_previsto", 0) or 0))
        note_edit = QTextEdit()
        note_edit.setFixedHeight(90)
        note_edit.setPlainText(str(initial.get("pedido_transporte_obs", "") or "").strip())
        response_edit = QTextEdit()
        response_edit.setFixedHeight(72)
        response_edit.setPlainText(str(initial.get("pedido_resposta_obs", "") or "").strip())
        form.addRow("Estado pedido", estado_combo)
        form.addRow("Transportadora", carrier_combo)
        form.addRow("Ref. pedido", ref_edit)
        form.addRow("Paletes carga", paletes_spin)
        form.addRow("Peso total", peso_spin)
        form.addRow("Volume total", volume_spin)
        form.addRow("Custo previsto", cost_spin)
        form.addRow("Observações pedido", note_edit)
        form.addRow("Resposta parceiro", response_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "pedido_transporte_estado": estado_combo.currentText().strip(),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "pedido_transporte_ref": ref_edit.text().strip(),
            "paletes_total_manual": paletes_spin.value(),
            "peso_total_manual_kg": peso_spin.value(),
            "volume_total_manual_m3": volume_spin.value(),
            "custo_previsto": cost_spin.value(),
            "pedido_transporte_obs": note_edit.toPlainText().strip(),
            "pedido_resposta_obs": response_edit.toPlainText().strip(),
        }

    def _stop_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Paragem / guia de transporte")
        dialog.setMinimumWidth(640)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        guide_combo = QComboBox()
        guide_combo.setEditable(False)
        guide_combo.addItem("", "")
        for row in list(initial.get("guide_options", []) or []):
            label = str(row.get("label", "") or row.get("numero", "") or "").strip()
            guide_combo.addItem(label, str(row.get("numero", "") or "").strip())
        current_guide = str(initial.get("guia_numero", "") or "").strip()
        current_index = 0
        for idx in range(guide_combo.count()):
            if str(guide_combo.itemData(idx) or "").strip() == current_guide:
                current_index = idx
                break
        guide_combo.setCurrentIndex(current_index)
        local_edit = QLineEdit(str(initial.get("local_descarga", "") or "").strip())
        latitude_edit = QLineEdit(str(initial.get("latitude", "") or "").strip())
        latitude_edit.setPlaceholderText("ex.: 41.1579")
        longitude_edit = QLineEdit(str(initial.get("longitude", "") or "").strip())
        longitude_edit.setPlaceholderText("ex.: -8.6291")
        zone_combo = QComboBox()
        zone_combo.setEditable(True)
        zone_combo.addItem("")
        for value in self._zone_options(initial):
            zone_combo.addItem(str(value))
        zone_combo.setCurrentText(str(initial.get("zona_transporte", "") or "").strip())
        contacto_edit = QLineEdit(str(initial.get("contacto", "") or "").strip())
        telefone_edit = QLineEdit(str(initial.get("telefone", "") or "").strip())
        raw_dt = str(initial.get("data_planeada", "") or "").strip().replace(" ", "T")
        raw_date = raw_dt.split("T", 1)[0] if raw_dt else ""
        raw_time = raw_dt.split("T", 1)[1] if "T" in raw_dt else ""
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("dd/MM/yyyy")
        qdate = QDate.fromString(raw_date, "yyyy-MM-dd") if raw_date else QDate.currentDate()
        if not qdate.isValid():
            qdate = QDate.currentDate()
        date_edit.setDate(qdate)
        time_edit = QTimeEdit()
        time_edit.setDisplayFormat("HH:mm")
        qtime = QTime.fromString(raw_time[:5], "HH:mm") if raw_time else QTime(8, 0)
        if not qtime.isValid():
            qtime = QTime(8, 0)
        time_edit.setTime(qtime)
        carga_box = QCheckBox("Carga conferida")
        carga_box.setChecked(bool(initial.get("check_carga_ok")))
        docs_box = QCheckBox("Documentos conferidos")
        docs_box.setChecked(bool(initial.get("check_docs_ok")))
        paletes_box = QCheckBox("Paletes conferidas")
        paletes_box.setChecked(bool(initial.get("check_paletes_ok")))
        pod_state_combo = QComboBox()
        pod_state_combo.addItems(["", "Pendente", "Recebido", "Incidente"])
        pod_state_combo.setCurrentText(str(initial.get("pod_estado", "") or "").strip())
        pod_name_edit = QLineEdit(str(initial.get("pod_recebido_nome", "") or "").strip())
        raw_pod_dt = str(initial.get("pod_recebido_at", "") or "").strip().replace(" ", "T")
        raw_pod_date = raw_pod_dt.split("T", 1)[0] if raw_pod_dt else ""
        raw_pod_time = raw_pod_dt.split("T", 1)[1] if "T" in raw_pod_dt else ""
        pod_date_edit = QDateEdit()
        pod_date_edit.setCalendarPopup(True)
        pod_date_edit.setDisplayFormat("dd/MM/yyyy")
        pod_qdate = QDate.fromString(raw_pod_date, "yyyy-MM-dd") if raw_pod_date else QDate.currentDate()
        if not pod_qdate.isValid():
            pod_qdate = QDate.currentDate()
        pod_date_edit.setDate(pod_qdate)
        pod_time_edit = QTimeEdit()
        pod_time_edit.setDisplayFormat("HH:mm")
        pod_qtime = QTime.fromString(raw_pod_time[:5], "HH:mm") if raw_pod_time else QTime.currentTime()
        if not pod_qtime.isValid():
            pod_qtime = QTime.currentTime()
        pod_time_edit.setTime(pod_qtime)
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(88)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        pod_obs_edit = QTextEdit()
        pod_obs_edit.setFixedHeight(72)
        pod_obs_edit.setPlainText(str(initial.get("pod_obs", "") or "").strip())
        form.addRow("Guia associada", guide_combo)
        form.addRow("Local descarga", local_edit)
        form.addRow("Latitude", latitude_edit)
        form.addRow("Longitude", longitude_edit)
        form.addRow("Zona", zone_combo)
        form.addRow("Contacto", contacto_edit)
        form.addRow("Telefone", telefone_edit)
        form.addRow("Data", date_edit)
        form.addRow("Hora", time_edit)
        form.addRow("Checklist", carga_box)
        form.addRow("", docs_box)
        form.addRow("", paletes_box)
        form.addRow("POD estado", pod_state_combo)
        form.addRow("Recebido por", pod_name_edit)
        form.addRow("Data POD", pod_date_edit)
        form.addRow("Hora POD", pod_time_edit)
        form.addRow("Obs. POD", pod_obs_edit)
        form.addRow("Observações", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected_guide = str(guide_combo.currentData() or "").strip()
        return {
            "expedicao_numero": selected_guide,
            "local_descarga": local_edit.text().strip(),
            "latitude": latitude_edit.text().strip(),
            "longitude": longitude_edit.text().strip(),
            "zona_transporte": zone_combo.currentText().strip(),
            "contacto": contacto_edit.text().strip(),
            "telefone": telefone_edit.text().strip(),
            "data_planeada": f"{date_edit.date().toString('yyyy-MM-dd')}T{time_edit.time().toString('HH:mm')}:00",
            "check_carga_ok": carga_box.isChecked(),
            "check_docs_ok": docs_box.isChecked(),
            "check_paletes_ok": paletes_box.isChecked(),
            "pod_estado": pod_state_combo.currentText().strip(),
            "pod_recebido_nome": pod_name_edit.text().strip(),
            "pod_recebido_at": f"{pod_date_edit.date().toString('yyyy-MM-dd')}T{pod_time_edit.time().toString('HH:mm')}:00" if pod_state_combo.currentText().strip() else "",
            "pod_obs": pod_obs_edit.toPlainText().strip(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _text_prompt(self, title: str, label: str) -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        edit = QLineEdit()
        form.addRow(label, edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return edit.text().strip()

    def _tariff_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or self.backend.transport_tariff_defaults() or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Tarifário de transporte")
        dialog.setMinimumWidth(620)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        carrier_combo = QComboBox()
        carrier_combo.setEditable(True)
        carrier_combo.addItem("")
        for value in self._supplier_options(initial):
            carrier_combo.addItem(str(value))
        carrier_combo.setCurrentText(
            " - ".join(
                [
                    part
                    for part in [
                        str(initial.get("transportadora_id", "") or "").strip(),
                        str(initial.get("transportadora_nome", "") or "").strip(),
                    ]
                    if part
                ]
            ).strip(" -")
        )
        zone_combo = QComboBox()
        zone_combo.setEditable(True)
        zone_combo.addItem("")
        for value in self._zone_options(initial):
            zone_combo.addItem(str(value))
        zone_combo.setCurrentText(str(initial.get("zona", "") or "").strip())
        zone_edit = zone_combo.lineEdit()
        if zone_edit is not None:
            zone_edit.setPlaceholderText("Escreva uma zona ou selecione uma existente")
            zone_edit.setClearButtonEnabled(True)
        base_spin = QDoubleSpinBox()
        base_spin.setRange(0.0, 1000000.0)
        base_spin.setDecimals(2)
        base_spin.setPrefix("EUR ")
        base_spin.setValue(float(initial.get("valor_base", 0) or 0))
        palette_spin = QDoubleSpinBox()
        palette_spin.setRange(0.0, 1000000.0)
        palette_spin.setDecimals(2)
        palette_spin.setPrefix("EUR ")
        palette_spin.setValue(float(initial.get("valor_por_palete", 0) or 0))
        kg_spin = QDoubleSpinBox()
        kg_spin.setRange(0.0, 1000.0)
        kg_spin.setDecimals(4)
        kg_spin.setPrefix("EUR ")
        kg_spin.setValue(float(initial.get("valor_por_kg", 0) or 0))
        volume_spin = QDoubleSpinBox()
        volume_spin.setRange(0.0, 1000000.0)
        volume_spin.setDecimals(2)
        volume_spin.setPrefix("EUR ")
        volume_spin.setValue(float(initial.get("valor_por_m3", 0) or 0))
        minimum_spin = QDoubleSpinBox()
        minimum_spin.setRange(0.0, 1000000.0)
        minimum_spin.setDecimals(2)
        minimum_spin.setPrefix("EUR ")
        minimum_spin.setValue(float(initial.get("custo_minimo", 0) or 0))
        active_box = QCheckBox("Ativo")
        active_box.setChecked(bool(initial.get("ativo", True)))
        obs_edit = QTextEdit()
        obs_edit.setFixedHeight(90)
        obs_edit.setPlainText(str(initial.get("observacoes", "") or "").strip())
        form.addRow("Transportadora", carrier_combo)
        form.addRow("Zona", zone_combo)
        form.addRow("Valor base", base_spin)
        form.addRow("Valor / palete", palette_spin)
        form.addRow("Valor / kg", kg_spin)
        form.addRow("Valor / m3", volume_spin)
        form.addRow("Custo mínimo", minimum_spin)
        form.addRow("", active_box)
        form.addRow("Observações", obs_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)

        def accept_tariff() -> None:
            if not zone_combo.currentText().strip():
                QMessageBox.warning(
                    dialog,
                    "Tarifário de transporte",
                    "Indique a zona do tarifário. Pode escrever uma zona nova ou selecionar uma existente.",
                )
                if zone_edit is not None:
                    zone_edit.setFocus()
                else:
                    zone_combo.setFocus()
                return
            dialog.accept()

        buttons.accepted.connect(accept_tariff)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "id": initial.get("id"),
            "transportadora_nome": carrier_combo.currentText().strip(),
            "zona": zone_combo.currentText().strip(),
            "valor_base": base_spin.value(),
            "valor_por_palete": palette_spin.value(),
            "valor_por_kg": kg_spin.value(),
            "valor_por_m3": volume_spin.value(),
            "custo_minimo": minimum_spin.value(),
            "ativo": active_box.isChecked(),
            "observacoes": obs_edit.toPlainText().strip(),
        }

    def _manage_tariffs(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Tarifário por transportadora / zona")
        dialog.setMinimumSize(920, 520)
        layout = QVBoxLayout(dialog)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Pesquisa"))
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("Filtrar por transportadora, zona ou observações")
        filter_row.addWidget(filter_edit, 1)
        layout.addLayout(filter_row)
        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["ID", "Transportadora", "Zona", "Base", "Palete", "Kg", "M3", "Mínimo"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(1, 2), contents=(0, 3, 4, 5, 6, 7))
        _set_table_columns(
            table,
            [
                (0, "interactive", 60),
                (1, "stretch", 230),
                (2, "stretch", 170),
                (3, "interactive", 90),
                (4, "interactive", 90),
                (5, "interactive", 90),
                (6, "interactive", 90),
                (7, "interactive", 90),
            ],
        )
        layout.addWidget(table, 1)
        buttons_row = QHBoxLayout()
        new_btn = QPushButton("Novo")
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        buttons_row.addStretch(1)
        buttons_row.addWidget(new_btn)
        buttons_row.addWidget(edit_btn)
        buttons_row.addWidget(remove_btn)
        buttons_row.addWidget(close_btn)
        layout.addLayout(buttons_row)
        state: dict[str, list[dict[str, Any]]] = {"rows": []}

        def current_row() -> dict[str, Any]:
            item = table.currentItem()
            if item is None or item.row() >= len(state["rows"]):
                return {}
            return state["rows"][item.row()]

        def render() -> None:
            rows = list(self.backend.transport_tariff_rows(filter_edit.text().strip()) or [])
            state["rows"] = rows
            _fill_table(
                table,
                [
                    [
                        row.get("id", "-"),
                        row.get("transportadora_nome", "Sem transportadora"),
                        row.get("zona", "-"),
                        _fmt_eur(float(row.get("valor_base", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_palete", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_kg", 0) or 0)),
                        _fmt_eur(float(row.get("valor_por_m3", 0) or 0)),
                        _fmt_eur(float(row.get("custo_minimo", 0) or 0)),
                    ]
                    for row in rows
                ],
                align_center_from=0,
            )
            for row_index, row in enumerate(rows):
                _paint_table_row(table, row_index, "Concluído" if bool(row.get("ativo", True)) else "Anulado")
            edit_btn.setEnabled(bool(rows))
            remove_btn.setEnabled(bool(rows))

        def handle_new() -> None:
            payload = self._tariff_dialog()
            if payload is None:
                return
            try:
                self.backend.transport_tariff_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        def handle_edit() -> None:
            row = current_row()
            if not row:
                QMessageBox.warning(dialog, "Tarifário", "Seleciona um tarifário.")
                return
            payload = self._tariff_dialog(row)
            if payload is None:
                return
            try:
                self.backend.transport_tariff_save(payload)
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        def handle_remove() -> None:
            row = current_row()
            if not row:
                QMessageBox.warning(dialog, "Tarifário", "Seleciona um tarifário.")
                return
            if QMessageBox.question(dialog, "Tarifário", f"Remover o tarifário da zona {row.get('zona', '-') or '-'}?") != QMessageBox.Yes:
                return
            try:
                self.backend.transport_tariff_remove(row.get("id"))
            except Exception as exc:
                QMessageBox.critical(dialog, "Tarifário", str(exc))
                return
            render()

        filter_edit.textChanged.connect(lambda _text: render())
        new_btn.clicked.connect(handle_new)
        edit_btn.clicked.connect(handle_edit)
        remove_btn.clicked.connect(handle_remove)
        close_btn.clicked.connect(dialog.accept)
        render()
        dialog.exec()

    def _apply_suggested_costs(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            detail = self.backend.transport_apply_suggested_cost(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()
        QMessageBox.information(
            self,
            "Transportes",
            (
                f"Custo sugerido aplicado à viagem {numero}.\n"
                f"Novo custo previsto: {_fmt_eur(float(detail.get('custo_previsto', 0) or 0))}"
            ),
        )

    def _new_trip(self) -> None:
        selected_orders = [
            str(row.get("numero", "") or "").strip()
            for row in self._selected_pending_rows()
            if str(row.get("numero", "") or "").strip()
        ]
        payload = self._trip_dialog(self.backend.transport_defaults())
        if payload is None:
            return
        try:
            detail = self.backend.transport_create_or_update(payload)
            numero = str(detail.get("numero", "") or "").strip()
            if selected_orders:
                detail = self.backend.transport_assign_orders(numero, selected_orders)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(str(detail.get("numero", "") or "").strip())
        self._show_trip_detail()
        self.transport_tabs.setCurrentIndex(1 if selected_orders else 0)

    def _edit_trip(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            initial = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._trip_dialog(initial)
        if payload is None:
            return
        if payload.get("_delete_trip"):
            try:
                self.backend.transport_remove_trip(numero)
            except Exception as exc:
                QMessageBox.critical(self, "Transportes", str(exc))
                return
            self.refresh()
            self._clear_trip_detail()
            return
        try:
            self.backend.transport_create_or_update(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _remove_trip(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        if QMessageBox.question(self, "Transportes", f"Apagar a viagem {numero} e libertar as encomendas associadas?") != QMessageBox.Yes:
            return
        try:
            self.backend.transport_remove_trip(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._clear_trip_detail()

    def _request_transport(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            initial = self.backend.transport_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._request_dialog(initial)
        if payload is None:
            return
        try:
            self.backend.transport_request_service(numero, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _edit_stop(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        try:
            guide_options = self.backend.transport_guide_options(enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        payload = self._stop_dialog({**dict(stop), "guide_options": guide_options})
        if payload is None:
            return
        try:
            self.backend.transport_update_stop(numero, enc_num, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()
        for row_index, row in enumerate(list(self.current_detail.get("paragens", []) or [])):
            if str(row.get("encomenda_numero", "") or "").strip() == enc_num:
                self.stops_table.selectRow(row_index)
                break

    def _assign_selected_orders(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        rows = self._selected_pending_rows()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona primeiro uma viagem.")
            return
        if not rows:
            QMessageBox.warning(self, "Transportes", "Seleciona pelo menos uma encomenda.")
            return
        try:
            self.backend.transport_assign_orders(numero, [str(row.get("numero", "") or "").strip() for row in rows])
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()
        self.transport_tabs.setCurrentIndex(1)

    def _apply_trip_status(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            self.backend.transport_set_status(numero, self.trip_status_combo.currentText().strip())
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _apply_stop_status(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        note = ""
        if self.stop_status_combo.currentText().strip() == "Incidente":
            note = self._text_prompt("Incidente na paragem", "Motivo")
            if note is None:
                return
        try:
            self.backend.transport_set_stop_status(numero, enc_num, self.stop_status_combo.currentText().strip(), note)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _move_stop(self, direction: int) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        try:
            self.backend.transport_move_stop(numero, enc_num, direction)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _remove_stop(self) -> None:
        current = self._current_trip_row()
        stop = self._current_stop_row()
        numero = str(current.get("numero", "") or "").strip()
        enc_num = str(stop.get("encomenda_numero", "") or "").strip()
        if not numero or not enc_num:
            QMessageBox.warning(self, "Transportes", "Seleciona uma paragem.")
            return
        if QMessageBox.question(self, "Transportes", f"Remover a encomenda {enc_num} desta viagem?") != QMessageBox.Yes:
            return
        try:
            self.backend.transport_remove_stop(numero, enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        self.refresh()
        self._restore_trip_selection(numero)
        self._show_trip_detail()

    def _open_trip_pdf(self) -> None:
        current = self._current_trip_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Transportes", "Seleciona uma viagem.")
            return
        try:
            path = self.backend.transport_route_sheet_open(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Transportes", str(exc))
            return
        QMessageBox.information(self, "Transportes", f"PDF aberto:\n{path}")

    def _open_trip_route(self) -> None:
        detail = dict(self.current_detail or {})
        if not detail:
            row = self._current_trip_row()
            numero = str(row.get("numero", "") or "").strip()
            if numero:
                try:
                    detail = dict(self.backend.transport_detail(numero) or {})
                except Exception as exc:
                    QMessageBox.critical(self, "Transportes", str(exc))
                    return
        map_url = self._trip_route_url(detail)
        if not list(detail.get("paragens", []) or []):
            QMessageBox.information(self, "Itinerário", "Esta viagem ainda não tem destinos.")
            return
        if not map_url:
            QMessageBox.information(
                self,
                "Itinerário",
                "Preenche a morada ou as coordenadas de pelo menos um destino.",
            )
            return
        QDesktopServices.openUrl(QUrl(map_url))

    def _open_selected_stop_map(self) -> None:
        stop = dict(self._current_stop_row() or {})
        if not stop:
            QMessageBox.information(self, "Google Maps", "Seleciona primeiro um destino.")
            return
        location = self._transport_map_location(
            stop,
            "local_descarga",
            "latitude",
            "longitude",
        )
        if not location:
            QMessageBox.information(
                self,
                "Google Maps",
                "Preenche a morada ou as coordenadas deste destino.",
            )
            return
        QDesktopServices.openUrl(
            QUrl(f"https://www.google.com/maps/search/?api=1&query={quote(location)}")
        )
