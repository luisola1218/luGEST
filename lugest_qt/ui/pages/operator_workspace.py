from __future__ import annotations
import os
import re
from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from pathlib import Path
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    cap_width as _cap_width,
    configure_table as _configure_table,
    elide_middle as _elide_middle,
    paint_table_row as _paint_table_row,
    selected_row_index as _selected_row_index,
    set_panel_tone as _set_panel_tone,
    set_table_columns as _set_table_columns,
    state_tone as _state_tone,
    table_visible_height as _table_visible_height,
)
from .runtime_support import (
    LIST_TABLE_FONT_PX,
    LIST_TABLE_ROW_PX,
    _adopt_layout_item,
    _apply_progress_style,
    _clear_layout_widgets,
    _format_client_label,
    _make_inline_progress,
    _operations_for_posto,
    _piece_ops_progress,
    _split_client_label,
    _take_layout_items,
)
from ..widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox, StatCard


class OperatorPage(QWidget):
    page_title = "Operador"
    page_subtitle = "Resumo operacional dos grupos ativos, peças e progresso por encomenda."
    uses_backend_reload = True

    def __init__(self, runtime_service, backend, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        self.backend = backend
        self.all_items: list[dict] = []
        self.items: list[dict] = []
        self.current_pieces: list[dict] = []
        self.checked_piece_ids: set[str] = set()
        self._syncing_piece_checks = False
        self.selected_group_key: tuple[str, str, str] | None = None
        self.selected_piece_id = ""
        self.ui_options: dict[str, bool] = {}
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        self.cards = [StatCard(title) for title in ("Encomendas ativas", "Pecas em curso", "Pecas em avaria", "Progresso")]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, index)
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("success")
        self.cards[2].set_tone("danger")
        self.cards[3].set_tone("warning")
        root.addWidget(cards_host)

        self.global_progress = QProgressBar()
        self.global_progress.setRange(0, 100)
        self.global_progress.setFormat("%p%")
        _apply_progress_style(self.global_progress)
        root.addWidget(self.global_progress)

        self.control_card = CardFrame()
        self.control_card.set_tone("info")
        control_layout = QVBoxLayout(self.control_card)
        control_layout.setContentsMargins(16, 14, 16, 14)
        control_layout.setSpacing(10)

        selectors_row = QHBoxLayout()
        selectors_row.setSpacing(10)
        self.operator_combo = QComboBox()
        self.operator_combo.setEditable(True)
        self.operator_combo.currentTextChanged.connect(self._sync_operator_assignment)
        self.posto_combo = QComboBox()
        self.posto_combo.setEditable(True)
        self.posto_combo.setCurrentText("Geral")
        self.posto_combo.currentTextChanged.connect(self._update_piece_context)
        self.operation_combo = QComboBox()
        self.operation_combo.setEditable(False)
        self.scan_edit = QLineEdit()
        self.scan_edit.setPlaceholderText("Picar OF/OPP")
        self.scan_edit.returnPressed.connect(self._handle_scan_code)
        selectors_row.addWidget(QLabel("Operador"))
        selectors_row.addWidget(self.operator_combo, 1)
        selectors_row.addWidget(QLabel("Posto"))
        selectors_row.addWidget(self.posto_combo)
        selectors_row.addWidget(QLabel("Operacao"))
        selectors_row.addWidget(self.operation_combo, 1)
        selectors_row.addWidget(QLabel("Scanner"))
        selectors_row.addWidget(self.scan_edit, 1)
        control_layout.addLayout(selectors_row)
        self.options_btn = QPushButton("Opcoes")
        self.options_btn.setProperty("variant", "secondary")
        self.options_btn.clicked.connect(self._open_options_dialog)

        actions_row = QVBoxLayout()
        actions_row.setSpacing(8)
        actions_row_top = QHBoxLayout()
        actions_row_top.setSpacing(8)
        actions_row_bottom = QHBoxLayout()
        actions_row_bottom.setSpacing(8)
        self.start_btn = QPushButton("Iniciar")
        self.start_btn.setProperty("variant", "success")
        self.start_btn.clicked.connect(self._start_piece)
        self.finish_btn = QPushButton("Finalizar")
        self.finish_btn.clicked.connect(self._finish_piece)
        self.resume_btn = QPushButton("Retomar")
        self.resume_btn.clicked.connect(self._resume_piece)
        self.pause_btn = QPushButton("Interromper")
        self.pause_btn.setProperty("variant", "secondary")
        self.pause_btn.clicked.connect(self._pause_piece)
        self.avaria_btn = QPushButton("Registar Avaria")
        self.avaria_btn.setProperty("variant", "destructive")
        self.avaria_btn.clicked.connect(self._register_avaria)
        self.close_avaria_btn = QPushButton("Fim Avaria")
        self.close_avaria_btn.setProperty("variant", "secondary")
        self.close_avaria_btn.clicked.connect(self._close_avaria)
        self.alert_btn = QPushButton("Alertar Chefia")
        self.alert_btn.clicked.connect(self._alert_chefia)
        self.manual_consume_btn = QPushButton("Dar Baixa")
        self.manual_consume_btn.setProperty("variant", "secondary")
        self.manual_consume_btn.clicked.connect(self._manual_consume_material)
        self.consume_components_btn = QPushButton("Consumir comp.")
        self.consume_components_btn.setProperty("variant", "success")
        self.consume_components_btn.clicked.connect(self._consume_montagem_components)
        self.partial_consume_btn = QPushButton("Baixa Parcial")
        self.partial_consume_btn.setProperty("variant", "warning")
        self.partial_consume_btn.clicked.connect(self._partial_consume_reserved_material)
        self.drawing_btn = QPushButton("Ver desenho")
        self.drawing_btn.setProperty("variant", "secondary")
        self.drawing_btn.clicked.connect(self._open_drawing)
        self.labels_btn = QPushButton("Etiquetas")
        self.labels_btn.setProperty("variant", "secondary")
        self.labels_btn.clicked.connect(self._open_labels_dialog)
        self.local_refresh_btn = QPushButton("Atualizar")
        self.local_refresh_btn.setProperty("variant", "secondary")
        self.local_refresh_btn.clicked.connect(self.refresh)
        for button in (
            self.start_btn,
            self.finish_btn,
            self.resume_btn,
            self.pause_btn,
            self.avaria_btn,
            self.close_avaria_btn,
        ):
            actions_row_top.addWidget(button)
        for button in (
            self.alert_btn,
            self.manual_consume_btn,
            self.consume_components_btn,
            self.partial_consume_btn,
            self.drawing_btn,
            self.labels_btn,
            self.local_refresh_btn,
        ):
            actions_row_bottom.addWidget(button)
        actions_row_top.addStretch(1)
        actions_row_bottom.addStretch(1)
        actions_row.addLayout(actions_row_top, 1)
        actions_row.addLayout(actions_row_bottom, 1)
        control_layout.addLayout(actions_row)

        self.feedback_label = QLabel("Seleciona uma peca para operar.")
        self.feedback_label.setWordWrap(True)
        self.feedback_label.setProperty("role", "muted")
        control_layout.addWidget(self.feedback_label)
        root.addWidget(self.control_card)

        self.context_card = CardFrame()
        context_layout = QVBoxLayout(self.context_card)
        context_layout.setContentsMargins(16, 14, 16, 14)
        context_layout.setSpacing(8)
        context_header = QHBoxLayout()
        self.piece_title_label = QLabel("Sem peca selecionada")
        self.piece_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.piece_state_chip = QLabel("-")
        _apply_state_chip(self.piece_state_chip, "-")
        context_header.addWidget(self.piece_title_label, 1)
        context_header.addWidget(self.options_btn, 0, Qt.AlignRight)
        context_header.addWidget(self.piece_state_chip, 0, Qt.AlignRight)
        context_layout.addLayout(context_header)
        self.piece_meta_label = QLabel("Seleciona um grupo e uma peca para ver contexto, pendencias e bloqueios.")
        self.piece_meta_label.setWordWrap(True)
        self.piece_meta_label.setProperty("role", "muted")
        context_layout.addWidget(self.piece_meta_label)
        self.issue_label = QLabel("-")
        self.issue_label.setWordWrap(True)
        context_layout.addWidget(self.issue_label)
        self.pending_label = QLabel("-")
        self.pending_label.setWordWrap(True)
        self.pending_label.setProperty("role", "muted")
        context_layout.addWidget(self.pending_label)
        self.operation_strip = QWidget()
        self.operation_strip_layout = QHBoxLayout(self.operation_strip)
        self.operation_strip_layout.setContentsMargins(0, 0, 0, 0)
        self.operation_strip_layout.setSpacing(6)
        context_layout.addWidget(self.operation_strip)
        self.piece_progress = QProgressBar()
        self.piece_progress.setRange(0, 100)
        self.piece_progress.setFormat("%p%")
        _apply_progress_style(self.piece_progress, compact=True)
        context_layout.addWidget(self.piece_progress)
        self.context_card.set_tone("default")
        root.addWidget(self.context_card)

        self.groups_table = QTableWidget(0, 9)
        self.groups_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Estado", "Material", "Esp.", "Plan", "Real", "Desvio", "Progress"])
        self.groups_table.verticalHeader().setVisible(False)
        self.groups_table.verticalHeader().setDefaultSectionSize(30)
        self.groups_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.groups_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.groups_table, stretch=(1, 3), contents=(2, 4, 5, 6, 7, 8))
        _set_table_columns(
            self.groups_table,
            [
                (0, "interactive", 190),
                (1, "stretch", 260),
                (2, "interactive", 138),
                (3, "stretch", 190),
                (4, "interactive", 72),
                (5, "interactive", 70),
                (6, "interactive", 70),
                (7, "interactive", 74),
                (8, "interactive", 112),
            ],
        )
        self.groups_table.itemSelectionChanged.connect(self._handle_group_selection)
        self.pieces_table = QTableWidget(0, 11)
        self.pieces_table.setHorizontalHeaderLabels(["Sel.", "Ref. Int.", "Ref. Ext.", "Estado", "Operacao", "Operador", "Prod. final", "Tempo", "Plan", "Progress", "Pendentes"])
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.verticalHeader().setDefaultSectionSize(28)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pieces_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        _configure_table(self.pieces_table, stretch=(2, 10), contents=(0, 1, 3, 4, 5, 6, 7, 8))
        _set_table_columns(
            self.pieces_table,
            [
                (0, "fixed", 42),
                (1, "interactive", 170),
                (2, "stretch", 240),
                (3, "interactive", 122),
                (4, "interactive", 132),
                (5, "interactive", 124),
                (6, "interactive", 86),
                (7, "interactive", 76),
                (8, "interactive", 72),
                (9, "interactive", 118),
                (10, "stretch", 200),
            ],
        )
        self.pieces_table.itemSelectionChanged.connect(self._update_piece_context)
        self.pieces_table.itemChanged.connect(self._handle_piece_check_changed)

        self.groups_card = CardFrame()
        self.groups_card.set_tone("info")
        groups_layout = QVBoxLayout(self.groups_card)
        groups_layout.setContentsMargins(16, 14, 16, 14)
        groups_layout.setSpacing(8)
        self.groups_title_label = QLabel("Encomendas ativas no operador")
        self.groups_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        groups_layout.addWidget(self.groups_title_label)
        groups_layout.addWidget(self.groups_table)
        root.addWidget(self.groups_card)

        self.pieces_card = CardFrame()
        self.pieces_card.set_tone("default")
        pieces_layout = QVBoxLayout(self.pieces_card)
        pieces_layout.setContentsMargins(16, 14, 16, 14)
        pieces_layout.setSpacing(8)
        pieces_header = QHBoxLayout()
        pieces_header.setSpacing(8)
        self.pieces_title_label = QLabel("Pecas da encomenda")
        self.pieces_title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        self.select_all_pieces_box = QCheckBox("Selecionar todas")
        self.select_all_pieces_box.toggled.connect(self._toggle_all_piece_checks)
        self.multi_count_chip = QLabel("0 selecionadas")
        _apply_state_chip(self.multi_count_chip, "-", "0 selecionadas")
        pieces_header.addWidget(self.pieces_title_label)
        pieces_header.addStretch(1)
        pieces_header.addWidget(self.multi_count_chip)
        pieces_header.addWidget(self.select_all_pieces_box)
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.pieces_table, 1)
        root.addWidget(self.pieces_card, 1)

    def refresh(self) -> None:
        user = self.backend.user or {}
        getter = getattr(self.backend, "ui_options", None)
        if callable(getter):
            try:
                self.ui_options = dict(getter() or {})
            except Exception:
                self.ui_options = {}
        posto_options_getter = getattr(self.backend, "operator_posto_options", None)
        if callable(posto_options_getter):
            try:
                posto_options = list(posto_options_getter() or ["Geral"])
            except Exception:
                posto_options = ["Geral"]
        else:
            posto_options = ["Geral", "Corte Laser", "Quinagem", "Roscagem", "Embalamento", "Montagem", "Soldadura"]
        self._set_combo_items(
            self.posto_combo,
            posto_options,
            preferred=self._current_posto(),
        )
        self._set_combo_items(
            self.operator_combo,
            self.backend.operator_names(),
            preferred=str(user.get("username", "") or "").strip(),
        )
        self._sync_operator_assignment()
        previous_group = self.selected_group_key
        previous_piece = self.selected_piece_id
        data = self.runtime_service.operator_board(username=str(user.get("username", "")), role=str(user.get("role", "")))
        summary = data.get("summary", {})
        all_items = self._hydrate_operator_items(list(data.get("items", [])))
        all_items = self._append_montagem_stock_groups(all_items)
        op_total = 0
        op_done = 0
        op_running = 0
        for board_item in all_items:
            for piece in list(board_item.get("pieces", []) or []):
                stats = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
                op_total += int(stats.get("total", 0) or 0)
                op_done += int(stats.get("done", 0) or 0)
                op_running += int(stats.get("running", 0) or 0)
        global_ops_progress = 0.0 if op_total <= 0 else round(((op_done + (0.5 * op_running)) / op_total) * 100.0, 1)
        self.cards[0].set_data(summary.get("encomendas_ativas", 0), f"Grupos {summary.get('grupos', 0)}")
        self.cards[1].set_data(summary.get("pecas_em_curso", 0), f"Em pausa {summary.get('pecas_em_pausa', 0)}")
        self.cards[2].set_data(summary.get("pecas_em_avaria", 0), f"Concluidas {summary.get('pecas_concluidas', 0)}")
        self.cards[3].set_data(f"{global_ops_progress:.1f}%", f"Ops {op_done}/{op_total} | Em curso {op_running}")
        self.global_progress.setValue(int(round(global_ops_progress)))
        self.all_items = all_items
        self.items = list(self.all_items)
        self._render_groups_table(previous_group, previous_piece)

    def _render_groups_table(self, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        self.groups_table.setSortingEnabled(False)
        self.groups_table.blockSignals(True)
        self.groups_table.setRowCount(len(self.items))
        for row_index, item in enumerate(self.items):
            group_piece_stats = [_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")) for piece in list(item.get("pieces", []) or [])]
            group_progress = round(sum(float(stat.get("progress_pct", 0) or 0) for stat in group_piece_stats) / len(group_piece_stats), 1) if group_piece_stats else float(item.get("progress_pct", 0) or 0)
            row_values = [
                item.get("encomenda", "-"),
                _format_client_label(item.get("cliente", "-"), show_name=self._show_client_name()),
                item.get("estado_espessura", item.get("estado", "-")),
                item.get("material", "-"),
                item.get("espessura", "-"),
                f"{item.get('tempo_plan_min', 0):.1f}",
                f"{item.get('tempo_real_min', 0):.1f}",
                f"{item.get('desvio_min', 0):.1f}",
                f"{group_progress:.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, "|".join(self._group_key(item)))
                if col_index >= 4:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.groups_table.setItem(row_index, col_index, cell)
            _paint_table_row(self.groups_table, row_index, str(item.get("estado_espessura", "")))
        self.groups_table.blockSignals(False)
        self.groups_table.setSortingEnabled(True)
        target_row = 0
        if previous_group:
            target_key = "|".join(previous_group)
            for index in range(self.groups_table.rowCount()):
                cell = self.groups_table.item(index, 0)
                if str((cell.data(Qt.UserRole) if cell is not None else "") or "").strip() == target_key:
                    target_row = index
                    break
        if self.groups_table.rowCount() > 0:
            self.groups_table.selectRow(target_row)
            self.selected_group_key = previous_group if previous_group else self._group_key(self._current_group())
            self.selected_piece_id = previous_piece
            self._handle_group_selection()
        else:
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self.pieces_title_label.setText("Pecas da encomenda")
            self._clear_piece_context()

    def _group_key(self, item: dict) -> tuple[str, str, str]:
        return (
            str(item.get("encomenda", "") or "").strip(),
            str(item.get("material", "") or "").strip(),
            str(item.get("espessura", "") or "").strip(),
        )

    def _set_combo_items(self, combo: QComboBox, values: list[str], preferred: str = "") -> None:
        current = combo.currentText().strip()
        target = current or str(preferred or "").strip()
        combo.blockSignals(True)
        combo.clear()
        seen: set[str] = set()
        ordered: list[str] = []
        for raw in list(values or []):
            text = str(raw or "").strip()
            key = text.lower()
            if text and key not in seen:
                seen.add(key)
                ordered.append(text)
        combo.addItems(ordered)
        if target:
            combo.setCurrentText(target)
        elif ordered:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _client_name_map(self) -> dict[str, str]:
        getter = getattr(self.backend, "ensure_data", None)
        if not callable(getter):
            return {}
        try:
            data = getter() or {}
        except Exception:
            return {}
        return {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list((data or {}).get("clientes", []) or [])
            if isinstance(row, dict) and str(row.get("codigo", "") or "").strip()
        }

    def _hydrate_operator_items(self, rows: list[dict]) -> list[dict]:
        clients = self._client_name_map()
        hydrated: list[dict] = []
        for raw in list(rows or []):
            item = dict(raw or {})
            client_raw = str(item.get("cliente", "") or "").strip()
            client_code, client_name = _split_client_label(client_raw)
            if client_code and not client_name:
                client_name = clients.get(client_code, "")
            client_label = _format_client_label(
                f"{client_code} - {client_name}".strip(" -") if (client_code or client_name) else client_raw,
                show_name=True,
            )
            item["cliente"] = client_label or client_raw or "-"
            item["cliente_codigo"] = client_code or str(item.get("cliente_codigo", "") or "").strip()
            item["cliente_nome"] = client_name or str(item.get("cliente_nome", "") or "").strip()
            item["cliente_label"] = client_label or client_raw or "-"
            hydrated.append(item)
        return hydrated

    def _append_montagem_stock_groups(self, rows: list[dict]) -> list[dict]:
        """Add virtual operator groups per order with stock components/raw materials to consume."""
        getter = getattr(self.backend, "operator_montagem_stock_group", None)
        groups_getter = getattr(self.backend, "operator_montagem_stock_groups", None)
        if not callable(getter) and not callable(groups_getter):
            return rows
        result = list(rows or [])
        seen_orders: set[str] = set()
        for item in list(rows or []):
            enc_num = str(item.get("encomenda", "") or "").strip()
            if enc_num:
                seen_orders.add(enc_num)
        data_getter = getattr(self.backend, "ensure_data", None)
        if callable(data_getter):
            try:
                for enc in list((data_getter() or {}).get("encomendas", []) or []):
                    enc_num = str((enc or {}).get("numero", "") or "").strip()
                    estado = str((enc or {}).get("estado", "") or "").strip().lower()
                    is_closed = any(token in estado for token in ("conclu", "cancel", "expedid", "fechad"))
                    if enc_num and not is_closed and list((enc or {}).get("montagem_itens", []) or []):
                        seen_orders.add(enc_num)
            except Exception:
                pass
        existing_keys = {"|".join(self._group_key(item)) for item in result}
        for enc_num in sorted(seen_orders):
            try:
                if callable(groups_getter):
                    groups = [dict(group or {}) for group in list(groups_getter(enc_num) or []) if isinstance(group, dict)]
                else:
                    groups = [dict(getter(enc_num) or {})]
            except Exception:
                continue
            for group in groups:
                if not group or not group.get("is_montagem_stock_group"):
                    continue
                key = "|".join(self._group_key(group))
                if key in existing_keys:
                    continue
                result.append(group)
                existing_keys.add(key)
        return result

    def _is_montagem_stock_group(self, item: dict | None = None) -> bool:
        group = item if item is not None else self._current_group()
        return bool(isinstance(group, dict) and group.get("is_montagem_stock_group"))

    def _current_group(self) -> dict:
        row_index = _selected_row_index(self.groups_table)
        if row_index < 0:
            return {}
        row_item = self.groups_table.item(row_index, 0)
        group_key = str(row_item.data(Qt.UserRole) or "").strip()
        if group_key:
            for item in self.items:
                if "|".join(self._group_key(item)) == group_key:
                    return item
        if row_index >= len(self.items):
            return {}
        return self.items[row_index]

    def _current_piece(self) -> dict:
        row_index = _selected_row_index(self.pieces_table)
        if row_index >= 0:
            row_item = self.pieces_table.item(row_index, 0)
            piece_id = str((row_item.data(Qt.UserRole) if row_item is not None else "") or "").strip()
            if piece_id:
                for piece in self.current_pieces:
                    if str(piece.get("id", "") or "").strip() == piece_id:
                        return piece
            if row_index < len(self.current_pieces):
                return self.current_pieces[row_index]
        target_piece_id = str(getattr(self, "selected_piece_id", "") or "").strip()
        if target_piece_id:
            for piece in self.current_pieces:
                if str(piece.get("id", "") or "").strip() == target_piece_id:
                    return piece
        return {}

    def _select_operator_target(self, enc_num: str, piece_id: str = "") -> bool:
        enc_txt = str(enc_num or "").strip()
        piece_txt = str(piece_id or "").strip()
        if not enc_txt:
            return False
        if hasattr(self, "_apply_order_filter") and callable(getattr(self, "_apply_order_filter")):
            target_group = None
            if piece_txt:
                for item in list(getattr(self, "all_items", []) or []):
                    if str(item.get("encomenda", "") or "").strip() != enc_txt:
                        continue
                    if any(str(piece.get("id", "") or "").strip() == piece_txt for piece in list(item.get("pieces", []) or [])):
                        target_group = self._group_key(item)
                        break
            self._apply_order_filter(enc_txt, previous_group=target_group, previous_piece=piece_txt)
            self._show_order_detail()
            if target_group:
                self._show_group_detail()
            return bool(getattr(self, "items", []))
        target_group_index = -1
        for index, item in enumerate(self.items):
            if str(item.get("encomenda", "") or "").strip() != enc_txt:
                continue
            if piece_txt and not any(str(piece.get("id", "") or "").strip() == piece_txt for piece in list(item.get("pieces", []) or [])):
                continue
            target_group_index = index
            break
        if target_group_index < 0:
            for index, item in enumerate(self.items):
                if str(item.get("encomenda", "") or "").strip() == enc_txt:
                    target_group_index = index
                    break
        if target_group_index < 0 or target_group_index >= self.groups_table.rowCount():
            return False
        self.groups_table.selectRow(target_group_index)
        self.selected_piece_id = piece_txt
        self._handle_group_selection()
        if piece_txt:
            for row_index, piece in enumerate(self.current_pieces):
                if str(piece.get("id", "") or "").strip() == piece_txt:
                    self.pieces_table.selectRow(row_index)
                    self._update_piece_context()
                    break
        return True

    def open_montagem_stock_group(self, enc_num: str) -> bool:
        enc_txt = str(enc_num or "").strip()
        if not enc_txt:
            return False
        self.refresh()
        target_group = None
        for item in list(getattr(self, "all_items", []) or []):
            if str(item.get("encomenda", "") or "").strip() == enc_txt and self._is_montagem_stock_group(item):
                target_group = self._group_key(item)
                break
        if target_group and hasattr(self, "_apply_order_filter"):
            self._apply_order_filter(enc_txt, previous_group=target_group)
            self._show_order_detail()
            self._show_group_detail()
            self._set_feedback(f"Componentes de montagem da encomenda {enc_txt} abertos.")
            return True
        if self._select_operator_target(enc_txt):
            for row_index, item in enumerate(list(getattr(self, "items", []) or [])):
                if str(item.get("encomenda", "") or "").strip() == enc_txt and self._is_montagem_stock_group(item):
                    self.groups_table.selectRow(row_index)
                    self._handle_group_selection()
                    self._set_feedback(f"Componentes de montagem da encomenda {enc_txt} abertos.")
                    return True
        self._set_feedback(f"Componentes de montagem da encomenda {enc_txt} nao estao visiveis no Operador.", error=True)
        return False

    def _handle_scan_code(self, source_edit: QLineEdit | None = None, expected_type: str = "") -> None:
        active_edit = source_edit or self.scan_edit
        code = (
            active_edit.text()
            .replace("\r", "")
            .replace("\n", "")
            .replace("\t", "")
            .strip()
        )
        if not code:
            return
        active_edit.clear()
        scanner = getattr(self.backend, "operator_scan_code", None)
        if not callable(scanner):
            self._set_feedback("Leitura por scanner indisponivel.", error=True)
            return
        expected = str(expected_type or "").strip().upper()
        if not expected:
            expected = "OPP" if self._is_group_detail_active() else ("GRP" if self._is_order_detail_active() else "OF")
        try:
            result = dict(scanner(code, current_posto=self._current_posto()) or {})
        except Exception as exc:
            if expected in {"GRP", "ESP", "ESPESSURA"} and self._open_group_from_thickness_scan(code):
                return
            self._set_feedback(str(exc), error=True)
            self._focus_active_scan_field()
            return
        scan_type = str(result.get("tipo", "") or "").upper()
        detail_mode = self._is_order_detail_active()
        current_order = str(self.selected_order_number or self._current_order_row().get("encomenda", "") or "").strip()
        if scan_type == "OF":
            if expected and expected != "OF":
                self._set_feedback("Nesta fase pica a espessura ou a OPP, nao a OF.", error=True)
                self._focus_active_scan_field()
                return
            enc = dict(result.get("encomenda", {}) or {})
            enc_num = str(enc.get("numero", "") or "").strip()
            if detail_mode:
                if current_order and enc_num == current_order:
                    self._set_feedback(f"OF {code} confirmada. Nesta pagina usa a etiqueta da espessura ou a OPP para continuar.")
                else:
                    self._set_feedback("Ja estas dentro de uma encomenda. Aqui o scanner serve para abrir espessura ou OPP.", error=True)
                return
            if self._select_operator_target(enc_num):
                self._set_feedback(f"OF {code} carregada.")
            else:
                self.refresh()
                if self._select_operator_target(enc_num):
                    self._set_feedback(f"OF {code} carregada.")
                else:
                    self._set_feedback(f"OF {code} encontrada, mas nao esta visivel neste posto/filtro.", error=True)
            return
        if scan_type == "GRP":
            if expected and expected not in {"GRP", "ESP", "ESPESSURA"}:
                self._set_feedback("Nesta fase pica a OPP da peca.", error=True)
                self._focus_active_scan_field()
                return
            enc_num = str(result.get("encomenda_numero", "") or "").strip()
            material = str(result.get("material", "") or "").strip()
            espessura = str(result.get("espessura", "") or "").strip()
            if detail_mode and current_order and enc_num != current_order:
                self._set_feedback(
                    f"A espessura {material} {espessura} pertence a outra encomenda ({enc_num}). Nesta pagina trabalha-se dentro da encomenda atual.",
                    error=True,
                )
                return
            target_group = None
            material_norm = str(material or "").strip().casefold()
            espessura_norm = self._normalize_thickness_scan(espessura)
            for item in list(getattr(self, "all_items", []) or []):
                if str(item.get("encomenda", "") or "").strip() != enc_num:
                    continue
                item_material_norm = str(item.get("material", "") or "").strip().casefold()
                item_espessura_norm = self._normalize_thickness_scan(str(item.get("espessura", "") or "").strip())
                if item_material_norm == material_norm and item_espessura_norm == espessura_norm:
                    target_group = self._group_key(item)
                    break
            if target_group and hasattr(self, "_apply_order_filter"):
                self._apply_order_filter(enc_num, previous_group=target_group)
                self._show_order_detail()
                self._show_group_detail()
                self._set_feedback(f"Espessura {material} {espessura} aberta.")
            else:
                self._set_feedback(f"Espessura {material} {espessura} nao esta visivel neste posto/filtro.", error=True)
            return
        if expected in {"GRP", "ESP", "ESPESSURA"}:
            if self._open_group_from_thickness_scan(code):
                return
            self._focus_active_scan_field()
            return
        if scan_type == "CPI":
            if expected and expected not in {"OPP", "CPI"}:
                self._set_feedback("Nesta fase pica a espessura.", error=True)
                self._focus_active_scan_field()
                return
            enc_num = str(result.get("encomenda_numero", "") or "").strip()
            item_id = str(result.get("item_id", "") or "").strip()
            codigo = str(result.get("codigo", "") or "").strip()
            espessura = str(result.get("espessura", "") or "").strip()
            if detail_mode and current_order and enc_num != current_order:
                self._set_feedback(
                    f"O componente {codigo or item_id} pertence a outra encomenda ({enc_num}). Nesta pagina trabalha-se dentro da encomenda atual.",
                    error=True,
                )
                return
            target_group = None
            for item in list(getattr(self, "all_items", []) or []):
                if str(item.get("encomenda", "") or "").strip() == enc_num and self._is_montagem_stock_group(item):
                    target_group = self._group_key(item)
                    break
            if target_group and hasattr(self, "_apply_order_filter"):
                self._apply_order_filter(enc_num, previous_group=target_group)
                self._show_order_detail()
                self._show_group_detail()
                for row_index, piece in enumerate(list(getattr(self, "current_pieces", []) or [])):
                    if str(piece.get("id", "") or "").strip() == item_id:
                        self.pieces_table.selectRow(row_index)
                        self._update_piece_context()
                        break
                self._set_feedback(f"Componente {codigo or item_id} aberto em {espessura or 'stock'}.")
            else:
                self._set_feedback(f"Componente {codigo or item_id} nao esta visivel neste posto/filtro.", error=True)
            return
        if scan_type == "COMP":
            enc_num = str(result.get("encomenda_numero", "") or "").strip()
            if detail_mode and current_order and enc_num != current_order:
                self._set_feedback(
                    f"Os componentes pertencem a outra encomenda ({enc_num}). Nesta pagina trabalha-se dentro da encomenda atual.",
                    error=True,
                )
                return
            target_group = None
            for item in list(getattr(self, "all_items", []) or []):
                if str(item.get("encomenda", "") or "").strip() == enc_num and self._is_montagem_stock_group(item):
                    target_group = self._group_key(item)
                    break
            if target_group and hasattr(self, "_apply_order_filter"):
                self._apply_order_filter(enc_num, previous_group=target_group)
                self._show_order_detail()
                self._show_group_detail()
                self._set_feedback(f"Componentes de montagem da encomenda {enc_num} abertos.")
            elif self._select_operator_target(enc_num):
                for row_index, item in enumerate(list(getattr(self, "items", []) or [])):
                    if str(item.get("encomenda", "") or "").strip() == enc_num and self._is_montagem_stock_group(item):
                        self.groups_table.selectRow(row_index)
                        self._handle_group_selection()
                        self._set_feedback(f"Componentes de montagem da encomenda {enc_num} abertos.")
                        return
                self._set_feedback(f"Componentes da encomenda {enc_num} nao estao visiveis neste posto/filtro.", error=True)
            else:
                self._set_feedback(f"Componentes da encomenda {enc_num} nao estao visiveis neste posto/filtro.", error=True)
            return
        is_operation_scan = scan_type == "OPR"
        if expected and expected != "OPP":
            self._set_feedback("Nesta fase pica a espessura, nao a OPP.", error=True)
            self._focus_active_scan_field()
            return
        enc_num = str(result.get("encomenda_numero", "") or "").strip()
        piece_id = str(result.get("piece_id", "") or "").strip()
        operation = str(result.get("operacao", "") or "").strip()
        if detail_mode and current_order and enc_num != current_order:
            self._set_feedback(
                f"A OPP {code} pertence a outra encomenda ({enc_num}). Nesta pagina usa a OPP da encomenda atual.",
                error=True,
            )
            return
        if expected == "OPP" and self._is_group_detail_active():
            current_group = self._current_group()
            group_key = self._group_key(current_group) if current_group else None
            piece_group_key = None
            if piece_id:
                for item in list(getattr(self, "all_items", []) or []):
                    if str(item.get("encomenda", "") or "").strip() != enc_num:
                        continue
                    if any(str(piece.get("id", "") or "").strip() == piece_id for piece in list(item.get("pieces", []) or [])):
                        piece_group_key = self._group_key(item)
                        break
            if group_key and piece_group_key and piece_group_key != group_key:
                self._set_feedback("OPP não pertence à espessura selecionada.", error=True)
                self._focus_active_scan_field()
                return
        if piece_id and is_operation_scan:
            self.checked_piece_ids.clear()
            self.checked_piece_ids.add(piece_id)
        if not self._select_operator_target(enc_num, piece_id):
            self.refresh()
            if not self._select_operator_target(enc_num, piece_id):
                self._set_feedback(f"OPP {code} encontrada, mas nao esta visivel neste posto/filtro.", error=True)
                return
        if piece_id:
            visible_scan_ids = [piece_id]
            if not is_operation_scan:
                visible_scan_ids = self._visible_piece_ids_for_scan(result, piece_id)
            already_checked = bool(visible_scan_ids) and all(value in self.checked_piece_ids for value in visible_scan_ids)
            if not is_operation_scan:
                self.checked_piece_ids.update(visible_scan_ids)
                self.selected_piece_id = visible_scan_ids[0] if visible_scan_ids else piece_id
            target_row_to_select = -1
            for row_index in range(self.pieces_table.rowCount()):
                item = self.pieces_table.item(row_index, 0)
                if item is None:
                    continue
                if str(item.data(Qt.UserRole) or "").strip() in set(visible_scan_ids):
                    item.setCheckState(Qt.Checked)
                    if target_row_to_select < 0:
                        target_row_to_select = row_index
            if target_row_to_select >= 0:
                self.pieces_table.selectRow(target_row_to_select)
                item = self.pieces_table.item(target_row_to_select, 0)
                if item is not None:
                    self.pieces_table.scrollToItem(item)
            self._sync_piece_multi_state()
        if operation:
            self.operation_combo.setCurrentText(operation)
        if is_operation_scan:
            operator_name = self._current_operator()
            if not operator_name:
                self._set_feedback("Seleciona o operador antes de iniciar a operação.", error=True)
                return
            if not piece_id or not operation:
                self._set_feedback("Código de operação incompleto.", error=True)
                return
            try:
                ctx = dict(self.backend.operator_piece_context(enc_num, piece_id) or {})
                active_ops = {
                    str(op or "").strip()
                    for op in list(ctx.get("active_pending_ops", []) or [])
                    if str(op or "").strip()
                }
                operation_norm = str(self.backend.desktop_main.normalize_operacao_nome(operation) or operation).strip()
                if operation_norm in active_ops:
                    op_limits = dict(ctx.get("operation_limits", {}) or {})
                    op_done = dict(ctx.get("operation_done", {}) or {})
                    limit = float(op_limits.get(operation_norm, ctx.get("current_operation_limit", 0)) or 0)
                    done = float(op_done.get(operation_norm, ctx.get("current_operation_done", 0)) or 0)
                    ok_qty = round(max(0.0, limit - done), 4)
                    if ok_qty <= 0:
                        ok_qty = float(ctx.get("default_ok", 1) or 1)
                    finish_result = self.backend.operator_finish_piece(
                        enc_num,
                        piece_id,
                        operator_name,
                        ok_qty,
                        0,
                        0,
                        operation=operation_norm,
                        posto=self._current_posto(),
                    )
                    if "laser" in operation_norm.lower():
                        finished_piece = dict((finish_result or {}).get("piece") or {})
                        self._maybe_close_material_session(
                            enc_num,
                            str(finished_piece.get("material", "") or "").strip(),
                            str(finished_piece.get("espessura", "") or "").strip(),
                            operator_name,
                            "conclusao_codigo_barras",
                        )
                    self.refresh()
                    self._set_feedback(f"Operação {operation_norm} concluída por código de barras.")
                    return
                pending_ops = {
                    str(op or "").strip()
                    for op in list(ctx.get("pending_ops", []) or [])
                    if str(op or "").strip()
                }
                if operation_norm not in pending_ops:
                    self._set_feedback(f"Operação {operation_norm} já está concluída nesta OPP.", error=True)
                    self._focus_active_scan_field()
                    return
                self.backend.operator_start_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    operation=operation_norm,
                    posto=self._current_posto(),
                )
            except Exception as exc:
                self._set_feedback(str(exc), error=True)
                return
            self.refresh()
            self._set_feedback(f"Operação {operation_norm} iniciada por código de barras.")
            return
        if piece_id:
            visible_ids = [str(piece.get("id", "") or "").strip() for piece in self.current_pieces if str(piece.get("id", "") or "").strip()]
            selected_count = sum(1 for value in visible_ids if value in self.checked_piece_ids)
            state_txt = "já estava selecionada" if already_checked else "adicionada à seleção"
            self._set_feedback(f"OPP {code} {state_txt}. {selected_count} selecionada(s).")
            self._focus_active_scan_field()
            return
        self._set_feedback(f"OPP {code} aberta" + (f" | {operation}" if operation else ""))

    def _visible_piece_ids_for_scan(self, scan_result: dict, base_piece_id: str) -> list[str]:
        base_txt = str(base_piece_id or "").strip()
        context_piece = dict(dict(scan_result.get("context", {}) or {}).get("piece", {}) or {})
        ref_int = str(context_piece.get("ref_interna", "") or "").strip()
        opp_txt = str(scan_result.get("opp", "") or context_piece.get("opp", "") or "").strip()
        matches: list[str] = []
        for piece in list(getattr(self, "current_pieces", []) or []):
            visible_id = str(piece.get("id", "") or "").strip()
            if not visible_id:
                continue
            piece_ref = str(piece.get("ref_interna", "") or "").strip()
            piece_opp = str(piece.get("opp", "") or "").strip()
            if visible_id == base_txt or (base_txt and visible_id.startswith(base_txt + "-")):
                matches.append(visible_id)
            elif ref_int and piece_ref == ref_int:
                matches.append(visible_id)
            elif opp_txt and piece_opp == opp_txt:
                matches.append(visible_id)
        if matches:
            return matches
        return [base_txt] if base_txt else []

    def _is_order_detail_active(self) -> bool:
        return bool(hasattr(self, "view_stack") and self.view_stack.currentWidget() is getattr(self, "detail_page", None))

    def _is_group_detail_active(self) -> bool:
        return bool(
            self._is_order_detail_active()
            and hasattr(self, "group_overview_stack")
            and self.group_overview_stack.currentIndex() == 1
        )

    def _normalize_thickness_scan(self, code: str) -> str:
        text = str(code or "").strip().upper()
        text = text.replace(",", ".")
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)", text)
        if match:
            numeric = match.group(1).rstrip("0").rstrip(".")
            return numeric or "0"
        return re.sub(r"\s+", "", text)

    def _open_group_from_thickness_scan(self, code: str) -> bool:
        current_order = str(self.selected_order_number or self._current_order_row().get("encomenda", "") or "").strip()
        if not current_order:
            self._set_feedback("Abre primeiro a OF antes de picar a espessura.", error=True)
            return False
        scanned = self._normalize_thickness_scan(code)
        if not scanned:
            self._set_feedback("Espessura vazia.", error=True)
            return False
        scanned_norm = self._normalize_thickness_scan(scanned)
        matches: list[dict] = []
        for item in list(getattr(self, "items", []) or []):
            if str(item.get("encomenda", "") or "").strip() != current_order:
                continue
            esp = str(item.get("espessura", "") or "").strip()
            esp_norm = self._normalize_thickness_scan(esp)
            candidates = {
                esp.upper(),
                esp_norm,
                f"S{esp_norm}".upper(),
            }
            if scanned.upper() in candidates or scanned_norm in candidates:
                matches.append(item)
        if not matches:
            self._set_feedback("Espessura não encontrada nesta encomenda.", error=True)
            return False
        group = matches[0]
        self._apply_order_filter(current_order, previous_group=self._group_key(group))
        self._show_order_detail()
        self._show_group_detail()
        self._set_feedback(f"Espessura {group.get('material', '-')} {group.get('espessura', '-')} aberta.")
        return True

    def _focus_active_scan_field(self) -> None:
        target = None
        if not self._is_order_detail_active() and hasattr(self, "list_scan_edit"):
            target = self.list_scan_edit
        elif self._is_group_detail_active() and hasattr(self, "opp_scan_edit"):
            target = self.opp_scan_edit
        elif hasattr(self, "thickness_scan_edit"):
            target = self.thickness_scan_edit
        elif hasattr(self, "scan_edit"):
            target = self.scan_edit
        if target is not None:
            QTimer.singleShot(0, target.setFocus)

    def _sync_scan_placeholder(self) -> None:
        if not hasattr(self, "scan_edit"):
            return
        placeholder = "Picar OPP" if self._is_group_detail_active() else ("Picar Espessura" if self._is_order_detail_active() else "Picar OF")
        self.scan_edit.setPlaceholderText(placeholder)
        if hasattr(self, "list_scan_edit"):
            self.list_scan_edit.setPlaceholderText("Picar OF")
        if hasattr(self, "thickness_scan_edit"):
            self.thickness_scan_edit.setPlaceholderText("Picar Espessura")
        if hasattr(self, "opp_scan_edit"):
            self.opp_scan_edit.setPlaceholderText("Picar OPP")
        self._focus_active_scan_field()

    def _show_client_name(self) -> bool:
        return bool(self.ui_options.get("operator_show_client_name", True))

    def _open_options_dialog(self) -> None:
        user = dict(self.backend.user or {})
        if str(user.get("role", "") or "").strip().lower() != "admin":
            QMessageBox.information(self, "Opcoes", "Apenas o admin pode alterar estas opcoes.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Opcoes do Operador")
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        info = QLabel("Opcoes visuais e operacionais do menu Operador.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)
        show_client_box = QCheckBox("Mostrar nome do cliente no operador")
        show_client_box.setChecked(self._show_client_name())
        layout.addWidget(show_client_box)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        setter = getattr(self.backend, "set_ui_option", None)
        if callable(setter):
            setter("operator_show_client_name", bool(show_client_box.isChecked()))
        self.refresh()

    def _selected_piece_ids(self) -> list[str]:
        ids: list[str] = []
        seen: set[str] = set()
        for piece in self.current_pieces:
            piece_id = str(piece.get("id", "") or "").strip()
            if piece_id and piece_id in self.checked_piece_ids and piece_id not in seen:
                ids.append(piece_id)
                seen.add(piece_id)
        if ids:
            return ids
        selection_model = self.pieces_table.selectionModel()
        if selection_model is not None:
            for model_index in selection_model.selectedRows():
                row = model_index.row()
                item = self.pieces_table.item(row, 0)
                piece_id = str((item.data(Qt.UserRole) if item is not None else "") or "").strip()
                if piece_id and piece_id not in seen:
                    ids.append(piece_id)
                    seen.add(piece_id)
        if ids:
            return ids
        piece = self._current_piece()
        piece_id = str(piece.get("id", "") or "").strip()
        return [piece_id] if piece_id else []

    def _order_number_for_piece_id(self, piece_id: str, fallback: str = "") -> str:
        piece_txt = str(piece_id or "").strip()
        if not piece_txt:
            return str(fallback or "").strip()
        sources: list[dict] = []
        current_group = self._current_group()
        if current_group:
            sources.append(current_group)
        sources.extend(list(getattr(self, "items", []) or []))
        sources.extend(list(getattr(self, "all_items", []) or []))
        seen_groups: set[tuple[str, str, str]] = set()
        for item in sources:
            if not isinstance(item, dict):
                continue
            key = self._group_key(item)
            if key in seen_groups:
                continue
            seen_groups.add(key)
            for piece in list(item.get("pieces", []) or []):
                if str(piece.get("id", "") or "").strip() == piece_txt:
                    enc_num = str(item.get("encomenda", "") or "").strip()
                    if enc_num:
                        return enc_num
        return str(fallback or "").strip()

    def _visible_piece_for_id(self, piece_id: str) -> dict:
        piece_txt = str(piece_id or "").strip()
        if not piece_txt:
            return {}
        for piece in list(getattr(self, "current_pieces", []) or []):
            if str(piece.get("id", "") or "").strip() == piece_txt:
                return dict(piece or {})
        for item in list(getattr(self, "items", []) or []) + list(getattr(self, "all_items", []) or []):
            for piece in list((item or {}).get("pieces", []) or []):
                if str(piece.get("id", "") or "").strip() == piece_txt:
                    return dict(piece or {})
        return {}

    def _backend_piece_id_for_selection(self, enc_num: str, selected_piece_id: str) -> str:
        enc_txt = str(enc_num or "").strip()
        selected_txt = str(selected_piece_id or "").strip()
        if not enc_txt or not selected_txt:
            return selected_txt
        visible_piece = self._visible_piece_for_id(selected_txt)
        ref_int = str(visible_piece.get("ref_interna", "") or "").strip()
        detail_getter = getattr(self.backend, "order_detail", None)
        if not callable(detail_getter):
            return selected_txt
        try:
            detail = dict(detail_getter(enc_txt) or {})
        except Exception:
            return selected_txt
        for piece in list(detail.get("pieces", []) or []):
            raw_id = str((piece or {}).get("id", "") or "").strip()
            if raw_id and raw_id == selected_txt:
                return raw_id
            if raw_id and selected_txt.startswith(raw_id + "-"):
                return raw_id
            if ref_int and str((piece or {}).get("ref_interna", "") or "").strip() == ref_int:
                return raw_id or selected_txt
        return selected_txt

    def _selected_piece_refs(self, *, allow_multiple: bool = True) -> list[tuple[str, str]] | None:
        group = self._current_group()
        enc_num = str(group.get("encomenda", "") or "").strip()
        piece_ids = self._selected_piece_ids()
        if not piece_ids:
            QMessageBox.warning(self, "Operador", "Seleciona pelo menos uma peca.")
            return None
        refs: list[tuple[str, str]] = []
        for piece_id in piece_ids:
            if not piece_id:
                continue
            piece_enc_num = self._order_number_for_piece_id(piece_id, enc_num)
            if piece_enc_num:
                refs.append((piece_enc_num, self._backend_piece_id_for_selection(piece_enc_num, piece_id)))
        if not refs:
            QMessageBox.warning(self, "Operador", "Seleciona pelo menos uma peca.")
            return None
        if not allow_multiple and len(refs) > 1:
            QMessageBox.warning(self, "Operador", "Esta acao exige apenas uma peca selecionada.")
            return None
        return refs

    def _selected_refs(self) -> tuple[str, str] | None:
        refs = self._selected_piece_refs(allow_multiple=False)
        return refs[0] if refs else None

    def _handle_piece_check_changed(self, item: QTableWidgetItem) -> None:
        if self._syncing_piece_checks or item.column() != 0:
            return
        piece_id = str(item.data(Qt.UserRole) or "").strip()
        if not piece_id:
            return
        if item.checkState() == Qt.Checked:
            self.checked_piece_ids.add(piece_id)
            self.selected_piece_id = piece_id
            self.pieces_table.selectRow(item.row())
            self.pieces_table.scrollToItem(item)
        else:
            self.checked_piece_ids.discard(piece_id)
            if str(getattr(self, "selected_piece_id", "") or "").strip() == piece_id:
                self.selected_piece_id = ""
                for row_index in range(self.pieces_table.rowCount()):
                    row_item = self.pieces_table.item(row_index, 0)
                    row_piece_id = str((row_item.data(Qt.UserRole) if row_item is not None else "") or "").strip()
                    if row_piece_id and row_piece_id in self.checked_piece_ids:
                        self.selected_piece_id = row_piece_id
                        self.pieces_table.selectRow(row_index)
                        break
        self._sync_piece_multi_state()
        self._update_piece_context()

    def _toggle_all_piece_checks(self, checked: bool) -> None:
        if self._syncing_piece_checks:
            return
        self._syncing_piece_checks = True
        try:
            visible_ids = {str(piece.get("id", "") or "").strip() for piece in self.current_pieces if str(piece.get("id", "") or "").strip()}
            if checked:
                self.checked_piece_ids.update(visible_ids)
            else:
                self.checked_piece_ids.difference_update(visible_ids)
            for row_index in range(self.pieces_table.rowCount()):
                item = self.pieces_table.item(row_index, 0)
                if item is not None:
                    item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        finally:
            self._syncing_piece_checks = False
        self._sync_piece_multi_state()

    def _sync_piece_multi_state(self) -> None:
        visible_ids = [str(piece.get("id", "") or "").strip() for piece in self.current_pieces if str(piece.get("id", "") or "").strip()]
        selected_count = sum(1 for piece_id in visible_ids if piece_id in self.checked_piece_ids)
        total_count = len(visible_ids)
        self._syncing_piece_checks = True
        try:
            self.select_all_pieces_box.setChecked(bool(total_count) and selected_count == total_count)
        finally:
            self._syncing_piece_checks = False
        _apply_state_chip(
            self.multi_count_chip,
            "Em producao" if selected_count else "-",
            f"{selected_count} selecionadas",
        )

    def _current_operator(self) -> str:
        return self.operator_combo.currentText().strip()

    def _current_posto(self) -> str:
        return self.posto_combo.currentText().strip() or "Geral"

    def _current_operation(self) -> str:
        return self.operation_combo.currentText().strip()

    def _default_posto_for_operator(self) -> str:
        resolver = getattr(self.backend, "operator_default_posto", None)
        if callable(resolver):
            try:
                return str(resolver(self._current_operator()) or "").strip() or "Geral"
            except Exception:
                return "Geral"
        return "Geral"

    def _sync_operator_assignment(self) -> None:
        default_posto = self._default_posto_for_operator()
        has_assignment = False
        checker = getattr(self.backend, "operator_has_posto_assignment", None)
        if callable(checker):
            try:
                has_assignment = bool(checker(self._current_operator()))
            except Exception:
                has_assignment = False
        current_posto = self._current_posto()
        if default_posto and default_posto != current_posto:
            self.posto_combo.blockSignals(True)
            self.posto_combo.setCurrentText(default_posto)
            self.posto_combo.blockSignals(False)
        self.posto_combo.setEnabled(not has_assignment)
        self._update_piece_context()

    def _set_feedback(self, text: str, error: bool = False) -> None:
        raw = str(text or "").strip() or "-"
        self.feedback_label.setText(_elide_middle(raw, 110))
        self.feedback_label.setToolTip(raw if len(raw) > 110 else "")
        if error:
            self.feedback_label.setStyleSheet("color: #b45f06; font-weight: 700;")
            _set_panel_tone(self.control_card, "danger")
        else:
            self.feedback_label.setStyleSheet("color: #475467;")
            _set_panel_tone(self.control_card, "info")

    def _clear_piece_context(self) -> None:
        self.piece_title_label.setText("Sem peca selecionada")
        self.piece_meta_label.setText("Seleciona um grupo e uma peca para ver contexto, pendencias e bloqueios.")
        self.issue_label.setText("-")
        self.pending_label.setText("-")
        _clear_layout_widgets(self.operation_strip_layout)
        _apply_state_chip(self.piece_state_chip, "-")
        self.piece_progress.setValue(0)
        self.operation_combo.clear()
        _set_panel_tone(self.context_card, "default")
        self._sync_piece_multi_state()
        self._set_button_states(False, False, False, False, False, False)

    def _set_button_states(
        self,
        has_piece: bool,
        has_operator: bool,
        has_pending: bool,
        has_open_avaria: bool,
        can_resume: bool,
        can_finish: bool | None = None,
        *,
        component_group: bool = False,
        can_consume_components: bool = False,
    ) -> None:
        component_only_buttons = (
            self.start_btn,
            self.finish_btn,
            self.resume_btn,
            self.pause_btn,
            self.avaria_btn,
            self.close_avaria_btn,
            self.alert_btn,
            self.manual_consume_btn,
            self.partial_consume_btn,
            self.drawing_btn,
            self.labels_btn,
        )
        for button in component_only_buttons:
            button.setVisible(not component_group)
        self.consume_components_btn.setVisible(True)
        self.local_refresh_btn.setVisible(True)
        if component_group:
            self.start_btn.setEnabled(False)
            self.finish_btn.setEnabled(False)
            self.resume_btn.setEnabled(False)
            self.pause_btn.setEnabled(False)
            self.avaria_btn.setEnabled(False)
            self.close_avaria_btn.setEnabled(False)
            self.alert_btn.setEnabled(False)
            self.drawing_btn.setEnabled(False)
            self.manual_consume_btn.setEnabled(False)
            self.partial_consume_btn.setEnabled(False)
            self.labels_btn.setEnabled(False)
            self.consume_components_btn.setEnabled(bool(can_consume_components))
            self.local_refresh_btn.setEnabled(True)
            return
        finish_enabled = has_piece and has_operator and (can_finish if can_finish is not None else has_pending) and not has_open_avaria
        self.start_btn.setEnabled(has_piece and has_operator and has_pending and not has_open_avaria)
        self.finish_btn.setEnabled(finish_enabled)
        self.resume_btn.setEnabled(has_piece and has_operator and can_resume and not has_open_avaria)
        self.pause_btn.setEnabled(has_piece and has_operator and not has_open_avaria)
        self.avaria_btn.setEnabled(has_piece and has_operator and not has_open_avaria)
        self.close_avaria_btn.setEnabled(has_piece and has_operator and has_open_avaria)
        self.alert_btn.setEnabled(has_piece and has_operator)
        self.drawing_btn.setEnabled(has_piece)
        self.manual_consume_btn.setEnabled(has_piece)
        self.consume_components_btn.setEnabled(False)
        self.partial_consume_btn.setEnabled(has_piece)
        self.labels_btn.setEnabled(has_piece)

    def _handle_group_selection(self) -> None:
        item = self._current_group()
        if not item:
            self.pieces_title_label.setText("Pecas da encomenda")
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self._clear_piece_context()
            return
        self.selected_group_key = self._group_key(item)
        client_label = _format_client_label(item.get("cliente", "-"), show_name=self._show_client_name())
        material = str(item.get("material", "-") or "-").strip() or "-"
        espessura = str(item.get("espessura", "-") or "-").strip() or "-"
        if self._is_montagem_stock_group(item):
            self._render_montagem_stock_group(item, client_label)
            return
        self.pieces_title_label.setText(
            f"Pecas da encomenda {item.get('encomenda', '-')} | Cliente {client_label} | {material} {espessura} mm"
        )
        target_piece_id = self.selected_piece_id
        all_pieces = list(item.get("pieces", []) or [])
        state_filter = self.detail_state_filter_combo.currentText().strip().lower() if hasattr(self, "detail_state_filter_combo") else "todas"
        query = self.detail_search_edit.text().strip().lower() if hasattr(self, "detail_search_edit") else ""
        filtered_pieces = []
        for piece in all_pieces:
            state = str(piece.get("estado", "") or "").strip().lower()
            if state_filter and state_filter != "todas":
                if state_filter == "em producao" and not ("produc" in state or "curso" in state):
                    continue
                if state_filter == "concluida" and "concl" not in state:
                    continue
                if state_filter == "em pausa" and not ("paus" in state or "interromp" in state):
                    continue
                if state_filter == "avaria" and "avaria" not in state:
                    continue
            if query:
                haystack = " ".join(
                    [
                        str(piece.get("ref_interna", "") or ""),
                        str(piece.get("ref_externa", "") or ""),
                        str(piece.get("operacao_atual", "") or ""),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            filtered_pieces.append(piece)
        self.current_pieces = filtered_pieces
        self.checked_piece_ids.intersection_update({str(piece.get("id", "") or "").strip() for piece in self.current_pieces})
        self.pieces_table.setSortingEnabled(False)
        self.pieces_table.blockSignals(True)
        self.pieces_table.setRowCount(len(self.current_pieces))
        for row_index, piece in enumerate(self.current_pieces):
            produced = f"{piece.get('produzido', 0):.1f}"
            piece_id = str(piece.get("id", "") or "").strip()
            ops_progress = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked if piece_id in self.checked_piece_ids else Qt.Unchecked)
            check_item.setData(Qt.UserRole, piece_id)
            check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            self.pieces_table.setItem(row_index, 0, check_item)
            row_values = [
                piece.get("ref_interna", "-"),
                piece.get("ref_externa", "-"),
                piece.get("estado", "-"),
                piece.get("operacao_atual", "-"),
                piece.get("operador", "-"),
                produced,
                f"{piece.get('tempo_min', 0):.1f}",
                f"{piece.get('planeado', 0):.1f}",
                "",
                ", ".join(piece.get("pendentes", []) or []),
            ]
            for offset, value in enumerate(row_values, start=1):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if offset in (6, 7, 8):
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.pieces_table.setItem(row_index, offset, cell)
            self.pieces_table.setCellWidget(row_index, 9, _make_inline_progress(float(ops_progress.get("progress_pct", 0) or 0)))
            _paint_table_row(self.pieces_table, row_index, str(piece.get("estado", "")))
        self.pieces_table.blockSignals(False)
        self.pieces_table.setSortingEnabled(False)
        if self.pieces_table.rowCount() == 0:
            self.selected_piece_id = ""
            self._clear_piece_context()
            return
        target_row = 0
        if target_piece_id:
            for index, piece in enumerate(self.current_pieces):
                if str(piece.get("id", "") or "").strip() == target_piece_id:
                    target_row = index
                    break
        self.pieces_table.selectRow(target_row)
        self._sync_piece_multi_state()
        self._update_piece_context()

    def _render_montagem_stock_group(self, item: dict, client_label: str) -> None:
        enc_num = str(item.get("encomenda", "") or "").strip()
        rows = list(item.get("montagem_items", []) or [])
        comp_count = int(item.get("componentes_count", 0) or sum(1 for row in rows if str(row.get("grupo_operador", "") or "") != "matéria-prima"))
        raw_count = int(item.get("materia_prima_count", 0) or sum(1 for row in rows if str(row.get("grupo_operador", "") or "") == "matéria-prima"))
        group_label = str(item.get("montagem_group_label", "") or "").strip()
        if not group_label:
            group_label = f"Componentes {comp_count} | Matéria-prima {raw_count}"
        self.pieces_title_label.setText(
            f"{group_label} | Encomenda {enc_num or '-'} | Cliente {client_label}"
        )
        self.current_pieces = rows
        self.checked_piece_ids.intersection_update({str(row.get("id", "") or "").strip() for row in rows})
        self.pieces_table.setSortingEnabled(False)
        self.pieces_table.blockSignals(True)
        self.pieces_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            row_id = str(row.get("id", "") or "").strip() or f"COMP::{enc_num}::{row_index}"
            plan = float(row.get("qtd_planeada", 0) or 0)
            done = float(row.get("qtd_consumida", 0) or 0)
            pending = max(0.0, plan - done)
            progress = 0.0 if plan <= 0 else min(100.0, (done / plan) * 100.0)
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked if row_id in self.checked_piece_ids else Qt.Unchecked)
            check_item.setData(Qt.UserRole, row_id)
            check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            self.pieces_table.setItem(row_index, 0, check_item)
            row_values = [
                row.get("codigo", "-") or "-",
                row.get("descricao", "-") or "-",
                row.get("estado", "-") or "-",
                f"{'Matéria-prima' if str(row.get('grupo_operador', '') or '') == 'matéria-prima' else 'Componentes'} / {row.get('tipo_label', 'Stock') or 'Stock'}",
                row.get("unidade", "-") or "-",
                f"{done:.1f}",
                "-",
                f"{plan:.1f}",
                "",
                f"Pendente {pending:.1f} | Falta {float(row.get('falta', 0) or 0):.1f}",
            ]
            for offset, value in enumerate(row_values, start=1):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if offset in (6, 7, 8):
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.pieces_table.setItem(row_index, offset, cell)
            self.pieces_table.setCellWidget(row_index, 9, _make_inline_progress(progress))
            _paint_table_row(self.pieces_table, row_index, str(row.get("estado", "")))
        self.pieces_table.blockSignals(False)
        self.pieces_table.setSortingEnabled(False)
        if self.pieces_table.rowCount() == 0:
            self.selected_piece_id = ""
            self._clear_piece_context()
            return
        self.pieces_table.selectRow(0)
        self._sync_piece_multi_state()
        self._update_piece_context()

    def _update_piece_context(self) -> None:
        group = self._current_group()
        piece = self._current_piece()
        if not group or not piece:
            self.selected_piece_id = ""
            self._clear_piece_context()
            return
        enc_num = str(group.get("encomenda", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        self.selected_piece_id = piece_id
        if self._is_montagem_stock_group(group):
            self._update_montagem_stock_context(group, piece)
            return
        try:
            ctx = self.backend.operator_piece_context(enc_num, piece_id)
        except Exception as exc:
            self._clear_piece_context()
            self._set_feedback(str(exc), error=True)
            return
        live_piece = dict(ctx.get("piece") or {})
        state = str(live_piece.get("estado", "") or piece.get("estado", "-")).strip() or "-"
        pending = list(ctx.get("pending_ops", []) or [])
        done = list(ctx.get("done_ops", []) or [])
        ops_progress = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
        issue_text = ""
        if ctx.get("has_open_avaria"):
            issue_text = (
                f"Avaria em aberto: {ctx.get('avaria_motivo', '-') or '-'}"
                f" | Aberta {float(ctx.get('avaria_open_min', 0) or 0):.1f} min"
                f" | Total nesta referencia {float(ctx.get('avaria_total_min', 0) or 0):.1f} min"
            )
        else:
            motivo = str(live_piece.get("interrupcao_peca_motivo", "") or "").strip()
            if motivo:
                issue_text = f"Interrupcao registada: {motivo}"
        full_piece_title = f"{piece.get('ref_interna', '-')} | {piece.get('ref_externa', '-')}"
        self.piece_title_label.setText(_elide_middle(full_piece_title, 64))
        self.piece_title_label.setToolTip(full_piece_title)
        _apply_state_chip(self.piece_state_chip, state)
        current_op_txt = str(ctx.get("current_operation", "") or piece.get("operacao_atual", "") or "-").strip() or "-"
        op_done_map = dict(ctx.get("operation_done", {}) or {})
        op_limit_map = dict(ctx.get("operation_limits", {}) or {})
        current_op_done = float(ctx.get("current_operation_done", 0) or 0)
        current_op_limit = float(ctx.get("current_operation_limit", 0) or 0)
        self.piece_meta_label.setText(
            f"Enc. {enc_num} | Cliente {_format_client_label(group.get('cliente', '-'), show_name=self._show_client_name())}\n"
            f"Material {group.get('material', '-')} | Esp. {group.get('espessura', '-')} mm | Op. atual {current_op_txt}\n"
            f"Oper. {piece.get('operador', '-') or '-'} | Qtd op. {current_op_done:.1f}/{current_op_limit:.1f} | Tempo {float(ctx.get('current_operation_elapsed_min', 0) or 0):.1f} min | Fluxo {int(ops_progress.get('done', 0))}/{int(ops_progress.get('total', 0))}"
        )
        self.issue_label.setText(issue_text)
        self.issue_label.setVisible(bool(issue_text))
        self.issue_label.setStyleSheet("font-size: 8.4px; color: #b45f06; font-weight: 700;" if ctx.get("has_open_avaria") else "font-size: 8.3px; color: #475467;")
        pend_txt = ", ".join(pending[:3]) if pending else "-"
        done_txt = ", ".join(done[:3]) if done else "-"
        if len(pending) > 3:
            pend_txt = f"{pend_txt}, +{len(pending) - 3}"
        if len(done) > 3:
            done_txt = f"{done_txt}, +{len(done) - 3}"
        self.pending_label.setText(f"Pendentes: {pend_txt} | Feitas: {done_txt}")
        _clear_layout_widgets(self.operation_strip_layout)
        for op in list(piece.get("ops", []) or []):
            chip = QLabel()
            op_name = str(op.get("nome", "") or "-")
            op_done = float(op_done_map.get(op_name, 0) or 0)
            op_limit = float(op_limit_map.get(op_name, 0) or 0)
            op_state = "Concluida" if op_limit > 0 and op_done >= op_limit - 1e-9 else str(op.get("estado", "") or "-")
            _apply_state_chip(chip, op_state, _elide_middle(f"{op_name} {op_done:.1f}/{op_limit:.1f}", 22))
            chip.setStyleSheet(
                chip.styleSheet()
                + " min-height: 21px; max-height: 21px; padding: 3px 8px; font-size: 8.4px; margin-top: 6px;"
            )
            self.operation_strip_layout.addWidget(chip)
        self.operation_strip_layout.addStretch(1)
        self.piece_progress.setValue(int(round(float(ops_progress.get("progress_pct", 0) or 0))))
        _set_panel_tone(self.context_card, _state_tone(state))
        self._sync_piece_multi_state()
        filtered_pending = _operations_for_posto(self._current_posto(), pending or list(piece.get("pendentes", []) or []))
        active_pending = list(ctx.get("active_pending_ops", []) or [])
        filtered_active_pending = _operations_for_posto(self._current_posto(), active_pending)
        preferred_op = (filtered_active_pending[0] if filtered_active_pending else "") or self._current_operation() or str(ctx.get("current_operation", "") or str(piece.get("operacao_atual", "") or "").split(" + ")[0])
        self._set_combo_items(
            self.operation_combo,
            filtered_pending,
            preferred=preferred_op,
        )
        state_norm = state.lower()
        can_resume = ("paus" in state_norm) or ("interromp" in state_norm)
        self._set_button_states(
            True,
            bool(self._current_operator()),
            self.operation_combo.count() > 0,
            bool(ctx.get("has_open_avaria")),
            can_resume,
            can_finish=bool(filtered_active_pending),
        )

    def _update_montagem_stock_context(self, group: dict, row: dict) -> None:
        enc_num = str(group.get("encomenda", "") or "").strip()
        code = str(row.get("codigo", "") or "-").strip() or "-"
        desc = str(row.get("descricao", "") or "-").strip() or "-"
        plan = float(row.get("qtd_planeada", 0) or 0)
        done = float(row.get("qtd_consumida", 0) or 0)
        pending = max(0.0, plan - done)
        falta = float(row.get("falta", 0) or 0)
        progress = 0 if plan <= 0 else int(round(min(100.0, (done / plan) * 100.0)))
        state = str(row.get("estado", "") or ("Consumido" if pending <= 0 else "Pendente")).strip()
        title = f"{code} | {desc}"
        self.piece_title_label.setText(_elide_middle(title, 72))
        self.piece_title_label.setToolTip(title)
        _apply_state_chip(self.piece_state_chip, state)
        self.piece_meta_label.setText(
            f"Enc. {enc_num} | Cliente {_format_client_label(group.get('cliente', '-'), show_name=self._show_client_name())}\n"
            f"Grupo: {'Matéria-prima' if str(row.get('grupo_operador', '') or '') == 'matéria-prima' else 'Componentes'} | Tipo {row.get('tipo_label', 'Stock') or 'Stock'} | Unidade {row.get('unidade', '-') or '-'}\n"
            f"Planeado {plan:.1f} | Consumido {done:.1f} | Pendente {pending:.1f} | Falta stock {falta:.1f}"
        )
        issue = ""
        if falta > 0:
            issue = f"Stock insuficiente neste componente: faltam {falta:.1f}."
        self.issue_label.setText(issue)
        self.issue_label.setVisible(bool(issue))
        self.issue_label.setStyleSheet("font-size: 8.4px; color: #b45f06; font-weight: 700;")
        self.pending_label.setText("Este grupo não inicia OPP. Usa Consumir comp. para dar baixa aos produtos e matéria-prima pendentes.")
        _clear_layout_widgets(self.operation_strip_layout)
        chip = QLabel()
        _apply_state_chip(chip, state, f"Stock {done:.1f}/{plan:.1f}")
        chip.setStyleSheet(chip.styleSheet() + " min-height: 21px; max-height: 21px; padding: 3px 8px; font-size: 8.4px; margin-top: 6px;")
        self.operation_strip_layout.addWidget(chip)
        self.operation_strip_layout.addStretch(1)
        self.piece_progress.setValue(progress)
        self.operation_combo.clear()
        _set_panel_tone(self.context_card, "warning" if pending > 0 else "success")
        self._sync_piece_multi_state()
        self._set_button_states(
            True,
            bool(self._current_operator()),
            False,
            False,
            False,
            False,
            component_group=True,
            can_consume_components=bool(group.get("can_consume_montagem")),
        )

    def _prompt_reason(self, title: str, label: str, options: list[str], default_value: str = "") -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        prompt = QLabel(label)
        prompt.setWordWrap(True)
        combo = QComboBox()
        combo.setEditable(True)
        self._set_combo_items(combo, options, preferred=default_value)
        layout.addWidget(prompt)
        layout.addWidget(combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return combo.currentText().strip() or None

    def _collect_start_operation_targets(self, refs_list: list[tuple[str, str]]) -> dict[str, object]:
        ordered_ops: list[str] = []
        targets: dict[str, list[tuple[str, str]]] = {}
        skipped_without_ops = 0
        errors: list[str] = []
        posto = self._current_posto()
        for enc_num, piece_id in list(refs_list or []):
            try:
                ctx = self.backend.operator_piece_context(enc_num, piece_id)
            except Exception as exc:
                errors.append(str(exc))
                continue
            pending = _operations_for_posto(posto, list(ctx.get("pending_ops", []) or []))
            if not pending:
                skipped_without_ops += 1
                continue
            for op_name in pending:
                op_key = str(op_name or "").strip()
                if not op_key:
                    continue
                if op_key not in targets:
                    targets[op_key] = []
                    ordered_ops.append(op_key)
                if (enc_num, piece_id) not in targets[op_key]:
                    targets[op_key].append((enc_num, piece_id))
        return {
            "ordered_ops": ordered_ops,
            "targets": targets,
            "skipped_without_ops": skipped_without_ops,
            "errors": errors,
        }

    def _prompt_start_operation(
        self,
        refs_list: list[tuple[str, str]],
        ordered_ops: list[str],
        targets: dict[str, list[tuple[str, str]]],
        skipped_without_ops: int = 0,
    ) -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Iniciar operacao")
        dialog.setMinimumSize(560, 380)
        dialog.setStyleSheet(
            """
            QDialog { background: #eaf2fb; }
            QLabel#startTitle { color: #06133f; font-size: 19px; font-weight: 900; }
            QLabel#startMeta { color: #475467; font-size: 11px; }
            QListWidget {
                background: #f8fbff;
                border: 1px solid #9fb7d4;
                border-radius: 12px;
                padding: 8px;
                outline: 0;
            }
            QListWidget::item {
                min-height: 42px;
                border: 1px solid #c8d7ea;
                border-radius: 10px;
                padding: 8px 12px;
                margin: 4px;
                color: #06133f;
                font-weight: 800;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #e8f0fa);
            }
            QListWidget::item:selected {
                color: #ffffff;
                border: 1px solid #020348;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #182a70, stop:1 #020348);
            }
            QDialogButtonBox QPushButton {
                min-width: 92px;
                min-height: 34px;
                border-radius: 18px;
                border: 1px solid #020348;
                color: white;
                font-weight: 900;
                background: #101a64;
            }
            QDialogButtonBox QPushButton:hover { background: #142174; }
            """
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        title_label = QLabel("Iniciar operação")
        title_label.setObjectName("startTitle")
        layout.addWidget(title_label)
        intro = QLabel(
            f"Selecionaste {len(refs_list)} peca(s). Escolhe a operacao que queres iniciar. "
            "So aparecem operacoes realmente pendentes nas pecas selecionadas."
        )
        intro.setObjectName("startMeta")
        intro.setWordWrap(True)
        layout.addWidget(intro)
        posto = str(self._current_posto() or "").strip() or "Geral"
        posto_label = QLabel(f"Posto atual: {posto}")
        posto_label.setObjectName("startMeta")
        posto_label.setProperty("role", "muted")
        layout.addWidget(posto_label)
        if skipped_without_ops:
            skipped_label = QLabel(
                f"{skipped_without_ops} peca(s) ficaram fora porque nao têm operacoes pendentes neste posto."
            )
            skipped_label.setWordWrap(True)
            skipped_label.setProperty("role", "muted")
            layout.addWidget(skipped_label)
        list_widget = QListWidget()
        list_widget.setAlternatingRowColors(False)
        for op_name in list(ordered_ops or []):
            compatible = len(list(targets.get(op_name, []) or []))
            item = QListWidgetItem(f"{op_name} ({compatible} peca{'s' if compatible != 1 else ''})")
            item.setData(Qt.UserRole, op_name)
            list_widget.addItem(item)
        preferred = self._current_operation()
        preferred_row = 0
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            op_name = str((item.data(Qt.UserRole) if item is not None else "") or "").strip()
            if preferred and op_name == preferred:
                preferred_row = index
                break
        if list_widget.count() > 0:
            list_widget.setCurrentRow(preferred_row)
        list_widget.itemDoubleClicked.connect(lambda _item: dialog.accept())
        layout.addWidget(list_widget)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        current_item = list_widget.currentItem()
        if current_item is None:
            QMessageBox.warning(self, "Operador", "Seleciona uma operacao para iniciar.")
            return None
        return str(current_item.data(Qt.UserRole) or "").strip() or None

    def _prompt_finish(self, ctx: dict, *, batch_idx: int = 0, batch_total: int = 0, preferred_operation: str = "") -> dict[str, float | str] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Finalizar operacao ({batch_idx}/{batch_total})" if batch_idx and batch_total else "Finalizar operacao")
        dialog.setMinimumWidth(540)
        dialog.setStyleSheet(
            """
            QDialog { background: #eaf2fb; }
            QLabel#finishTitle { color: #06133f; font-size: 17px; font-weight: 900; }
            QLabel#finishMeta { color: #475467; font-size: 11px; }
            QComboBox, QDoubleSpinBox {
                min-height: 32px;
                border: 1px solid #9fb7d4;
                border-radius: 8px;
                padding: 4px 8px;
                background: #ffffff;
                color: #06133f;
                font-weight: 700;
            }
            QDialogButtonBox QPushButton {
                min-width: 92px;
                min-height: 34px;
                border-radius: 18px;
                border: 1px solid #020348;
                color: white;
                font-weight: 900;
                background: #101a64;
            }
            QDialogButtonBox QPushButton:hover { background: #142174; }
            """
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        piece = dict(ctx.get("piece") or {})
        ref_int = str(piece.get("ref_interna", "") or "-").strip() or "-"
        ref_ext = str(piece.get("ref_externa", "") or "-").strip() or "-"
        title = QLabel(f"{ref_int} | {ref_ext}")
        title.setObjectName("finishTitle")
        meta = QLabel(
            f"Qtd {float(ctx.get('produzido_ok', 0) or 0):.1f}/{float(ctx.get('quantidade_pedida', 0) or 0):.1f}"
            f" | Pendentes: {', '.join(list(ctx.get('pending_ops', []) or [])[:3]) or '-'}"
        )
        meta.setObjectName("finishMeta")
        meta.setProperty("role", "muted")
        meta.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(meta)
        form = QFormLayout()
        op_combo = QComboBox()
        raw_finish_ops = list(ctx.get("active_pending_ops", []) or []) or list(ctx.get("pending_ops", []) or [])
        posto_finish_ops = _operations_for_posto(self._current_posto(), raw_finish_ops)
        finish_ops = posto_finish_ops or raw_finish_ops
        self._set_combo_items(
            op_combo,
            finish_ops,
            preferred=preferred_operation or self._current_operation(),
        )
        op_limits = dict(ctx.get("operation_limits", {}) or {})
        op_done = dict(ctx.get("operation_done", {}) or {})
        ok_spin = QDoubleSpinBox()
        nok_spin = QDoubleSpinBox()
        qual_spin = QDoubleSpinBox()
        for spin in (ok_spin, nok_spin, qual_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(1)
            spin.setSingleStep(1.0)

        def _suggest_remaining() -> None:
            op_name = op_combo.currentText().strip()
            limit = float(op_limits.get(op_name, ctx.get("current_operation_limit", 0)) or 0)
            done = float(op_done.get(op_name, ctx.get("current_operation_done", 0)) or 0)
            remaining = max(0.0, round(limit - done, 1))
            ok_spin.setValue(remaining)

        _suggest_remaining()
        op_combo.currentTextChanged.connect(lambda _text: _suggest_remaining())
        form.addRow("Operacao", op_combo)
        form.addRow("OK", ok_spin)
        form.addRow("NOK", nok_spin)
        form.addRow("Qualidade", qual_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return {
            "operation": op_combo.currentText().strip(),
            "ok": ok_spin.value(),
            "nok": nok_spin.value(),
            "qual": qual_spin.value(),
        }

    def _prompt_laser_stock_resolution(self, stock_state: dict) -> dict[str, float | str | bool] | None:
        total_qty = float(stock_state.get("total_qty", 0) or 0)
        reserved_qty = float(stock_state.get("reserved_qty", 0) or 0)
        remaining_qty = float(stock_state.get("remaining_qty", 0) or 0)
        manual_stock_required = bool(stock_state.get("manual_stock_required", False))
        reserved_sources = list(stock_state.get("reserved_sources", []) or [])
        has_reserved = reserved_qty > 0
        if not manual_stock_required and not has_reserved:
            return {"material_id": "", "quantity": 0.0, "allow_without_stock": False, "retalho": {}, "source_material_id": ""}
        candidates = reserved_sources if has_reserved else list(stock_state.get("candidates", []) or [])
        if not candidates and not has_reserved:
            answer = QMessageBox.question(
                self,
                "Baixa material Laser",
                (
                    f"Laser concluido para {stock_state.get('material', '-')} {stock_state.get('espessura', '-')} mm.\n\n"
                    f"Peças produzidas: {total_qty:.1f}\n"
                    f"Baixa automatica por cativacao: {reserved_qty:.1f}\n"
                    "Baixa manual de stock: obrigatoria\n\n"
                    "Nao existe stock disponivel correspondente.\n"
                    "Pretende concluir sem baixa adicional?"
                ),
                QMessageBox.Yes | QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return None
            return {"material_id": "", "quantity": 0.0, "allow_without_stock": True, "retalho": {}, "source_material_id": ""}

        dialog = QDialog(self)
        session_close = bool(stock_state.get("session_close"))
        dialog.setWindowTitle("Baixa do material cativado" if session_close else "Baixa material Laser")
        dialog.setMinimumWidth(720)
        layout = QVBoxLayout(dialog)
        info_lines = [
            (
                f"Fim da sessão de trabalho em {stock_state.get('material', '-')} {stock_state.get('espessura', '-')} mm"
                if session_close
                else f"Laser concluido para {stock_state.get('material', '-')} {stock_state.get('espessura', '-')} mm"
            ),
            f"Peças produzidas: {total_qty:.1f} | Material cativado disponível: {reserved_qty:.1f}",
        ]
        if has_reserved:
            info_lines.append(
                "Seleciona o lote e a quantidade realmente consumida. O restante continuará cativado para esta encomenda."
            )
        else:
            info_lines.append("Sem material cativado. Indica a quantidade de stock realmente consumida e o lote utilizado.")
        info = QLabel("\n".join(info_lines))
        info.setWordWrap(True)
        layout.addWidget(info)
        form = QFormLayout()
        material_combo = QComboBox()
        for row in candidates:
            label = (
                f"{row.get('material_id', '-') } | {row.get('dimensao', '-') } | "
                f"Disp. {float(row.get('disponivel', 0) or 0):.1f} | {row.get('local', '-') } | {row.get('lote', '-') }"
            )
            material_combo.addItem(label, row)
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.0, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setSingleStep(1.0)
        qty_spin.setValue(0.0)
        skip_box = QCheckBox("Concluir sem baixa adicional")
        skip_box.toggled.connect(lambda checked: (material_combo.setEnabled(not checked), qty_spin.setEnabled(not checked)))
        form.addRow("Material cativado" if has_reserved else "Stock", material_combo)
        form.addRow("Qtd cativada a baixar" if has_reserved else "Qtd stock consumido", qty_spin)
        form.addRow("", skip_box)
        if has_reserved:
            skip_box.hide()

            def _sync_reserved_quantity() -> None:
                selected = dict(material_combo.currentData() or {})
                maximum = max(
                    0.0,
                    float(selected.get("quantidade", selected.get("disponivel", 0)) or 0),
                )
                qty_spin.setRange(0.01 if maximum > 0 else 0.0, maximum if maximum > 0 else 0.0)
                qty_spin.setValue(min(1.0, maximum) if maximum > 0 else 0.0)
                qty_spin.setSuffix(f" / {maximum:.2f} cativadas" if maximum > 0 else "")

            material_combo.currentIndexChanged.connect(_sync_reserved_quantity)
            _sync_reserved_quantity()
        layout.addLayout(form)

        retalho_card = CardFrame()
        retalho_card.set_tone("warning")
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(12, 10, 12, 10)
        retalho_layout.setHorizontalSpacing(8)
        retalho_layout.setVerticalSpacing(6)
        retalho_layout.addWidget(QLabel("Retalho comprimento"), 0, 0)
        retalho_layout.addWidget(QLabel("Retalho largura"), 0, 1)
        retalho_layout.addWidget(QLabel("Qtd retalho"), 0, 2)
        retalho_layout.addWidget(QLabel("Metros"), 0, 3)
        comp_spin = QDoubleSpinBox()
        larg_spin = QDoubleSpinBox()
        qtd_retalho_spin = QDoubleSpinBox()
        metros_spin = QDoubleSpinBox()
        for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(2)
            spin.setAlignment(Qt.AlignCenter)
        retalho_layout.addWidget(comp_spin, 1, 0)
        retalho_layout.addWidget(larg_spin, 1, 1)
        retalho_layout.addWidget(qtd_retalho_spin, 1, 2)
        retalho_layout.addWidget(metros_spin, 1, 3)
        retalho_layout.addWidget(QLabel("Lote origem do retalho"), 2, 0)
        source_combo = QComboBox()
        source_candidates = reserved_sources if reserved_sources else candidates
        for row in source_candidates:
            source_combo.addItem(
                f"{row.get('material_id', '-') } | {row.get('lote', '-') } | {row.get('dimensao', '-')}",
                str(row.get("material_id", "") or "").strip(),
            )
        retalho_layout.addWidget(source_combo, 2, 1, 1, 3)
        layout.addWidget(retalho_card)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        retalho = {
            "comprimento": comp_spin.value(),
            "largura": larg_spin.value(),
            "quantidade": qtd_retalho_spin.value(),
            "metros": metros_spin.value(),
        }
        has_retalho = any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
        if skip_box.isChecked():
            return {
                "material_id": "",
                "quantity": 0.0,
                "allow_without_stock": True,
                "retalho": retalho if has_retalho else {},
                "source_material_id": str(source_combo.currentData() or "").strip() if has_retalho else "",
            }
        current = dict(material_combo.currentData() or {})
        return {
            "material_id": str(current.get("material_id", "") or "").strip(),
            "quantity": qty_spin.value(),
            "allow_without_stock": False,
            "retalho": retalho if has_retalho else {},
            "source_material_id": str(source_combo.currentData() or "").strip() if has_retalho else "",
        }

    def _prompt_material_session_decision(self, stock_state: dict) -> str:
        dialog = QDialog(self)
        dialog.setWindowTitle("Terminar operação de material")
        dialog.setMinimumWidth(650)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)
        title = QLabel("Não existem mais peças em execução desta espessura.")
        title.setStyleSheet("font-size: 15px; font-weight: 900; color: #10253d;")
        layout.addWidget(title)
        material = str(stock_state.get("material", "-") or "-")
        espessura = str(stock_state.get("espessura", "-") or "-")
        reserved = float(stock_state.get("reserved_qty", 0) or 0)
        context = QLabel(
            f"Material {material} | Espessura {espessura} mm | Cativado {reserved:.1f}\n"
            "Como pretende terminar esta operação?"
        )
        context.setWordWrap(True)
        layout.addWidget(context)
        decision = {"value": "cancel"}
        actions = QHBoxLayout()
        consume_btn = QPushButton("Dar baixa do material cativado")
        consume_btn.setProperty("variant", "primary")
        keep_btn = QPushButton("Manter material cativado")
        keep_btn.setProperty("variant", "secondary")
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setProperty("variant", "secondary")

        def choose(value: str) -> None:
            decision["value"] = value
            dialog.accept()

        consume_btn.clicked.connect(lambda: choose("consume"))
        keep_btn.clicked.connect(lambda: choose("keep"))
        cancel_btn.clicked.connect(dialog.reject)
        actions.addWidget(consume_btn)
        actions.addWidget(keep_btn)
        actions.addStretch(1)
        actions.addWidget(cancel_btn)
        layout.addLayout(actions)
        if dialog.exec() != QDialog.Accepted:
            return "cancel"
        return str(decision["value"])

    def _prompt_supervisor_password(self) -> bool:
        dialog = QDialog(self)
        dialog.setWindowTitle("Autorizacao superior")
        dialog.setMinimumWidth(360)
        layout = QVBoxLayout(dialog)
        info = QLabel("Introduz a password de supervisor para autorizar a baixa manual.")
        info.setWordWrap(True)
        layout.addWidget(info)
        password_edit = QLineEdit()
        password_edit.setEchoMode(QLineEdit.Password)
        password_edit.setPlaceholderText("Password supervisor")
        layout.addWidget(password_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return False
        verify = getattr(self.backend, "verify_supervisor_password", None)
        if not callable(verify) or not verify(password_edit.text()):
            QMessageBox.critical(self, "Dar Baixa", "Password de supervisor invalida.")
            return False
        return True

    def _prompt_manual_material_consumption(self, material: str, espessura: str) -> dict[str, Any] | None:
        candidates_fn = getattr(self.backend, "material_candidates", None)
        if not callable(candidates_fn):
            return None
        candidates = list(candidates_fn(material, espessura) or [])
        if not candidates:
            QMessageBox.information(self, "Dar Baixa", f"Sem stock disponivel para {material} {espessura} mm.")
            return None
        dialog = QDialog(self)
        dialog.setWindowTitle("Dar Baixa")
        dialog.resize(860, 560)
        layout = QVBoxLayout(dialog)
        info = QLabel(
            f"Baixa manual para {material} {espessura} mm.\n"
            "Seleciona as quantidades por lote. Se criares retalho, associa-o ao lote certo."
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        table = QTableWidget(len(candidates), 6)
        table.setHorizontalHeaderLabels(["Dimensao", "Disponivel", "Local", "Lote", "Peso/Un.", "Baixar"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(26)
        table.setStyleSheet(
            "QTableWidget { font-size: 11px; } "
            "QHeaderView::section { font-size: 11px; font-weight: 700; } "
            "QDoubleSpinBox { font-size: 11px; } "
            "QDoubleSpinBox#consumeCellSpin { border: none; background: transparent; padding: 0 8px; }"
        )
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.resizeSection(5, 122)
        spinners: list[tuple[dict[str, Any], QDoubleSpinBox]] = []
        for row_index, row in enumerate(candidates):
            table.setItem(row_index, 0, QTableWidgetItem(str(row.get("dimensao", "-"))))
            table.setItem(row_index, 1, QTableWidgetItem(f"{float(row.get('disponivel', 0) or 0):.2f}"))
            table.setItem(row_index, 2, QTableWidgetItem(str(row.get("local", "-"))))
            table.setItem(row_index, 3, QTableWidgetItem(str(row.get("lote", "-"))))
            table.setItem(row_index, 4, QTableWidgetItem(f"{float(row.get('peso_unid', 0) or 0):.3f} kg"))
            spin = QDoubleSpinBox()
            spin.setRange(0.0, float(row.get("disponivel", 0) or 0))
            spin.setDecimals(2)
            spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
            spin.setAlignment(Qt.AlignCenter)
            spin.setFrame(False)
            spin.setObjectName("consumeCellSpin")
            spin.setMinimumWidth(104)
            table.setCellWidget(row_index, 5, spin)
            spinners.append((row, spin))
        table.setMinimumHeight(_table_visible_height(table, max(6, len(candidates)), extra=22))
        layout.addWidget(table, 1)

        retalho_card = CardFrame()
        retalho_card.set_tone("warning")
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(12, 10, 12, 10)
        retalho_layout.setHorizontalSpacing(8)
        retalho_layout.setVerticalSpacing(6)
        retalho_layout.addWidget(QLabel("Retalho comprimento"), 0, 0)
        retalho_layout.addWidget(QLabel("Retalho largura"), 0, 1)
        retalho_layout.addWidget(QLabel("Qtd retalho"), 0, 2)
        retalho_layout.addWidget(QLabel("Metros"), 0, 3)
        comp_spin = QDoubleSpinBox()
        larg_spin = QDoubleSpinBox()
        qtd_retalho_spin = QDoubleSpinBox()
        metros_spin = QDoubleSpinBox()
        for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(2)
            spin.setAlignment(Qt.AlignCenter)
        retalho_layout.addWidget(comp_spin, 1, 0)
        retalho_layout.addWidget(larg_spin, 1, 1)
        retalho_layout.addWidget(qtd_retalho_spin, 1, 2)
        retalho_layout.addWidget(metros_spin, 1, 3)
        retalho_layout.addWidget(QLabel("Lote origem do retalho"), 2, 0)
        source_combo = QComboBox()
        for row in candidates:
            source_combo.addItem(
                f"{row.get('material_id', '-') } | {row.get('lote', '-') } | {row.get('dimensao', '-')}",
                str(row.get("material_id", "") or "").strip(),
            )
        retalho_layout.addWidget(source_combo, 2, 1, 1, 3)
        layout.addWidget(retalho_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        allocations: list[dict[str, Any]] = []
        selected_ids: set[str] = set()
        for row, spin in spinners:
            value = spin.value()
            if value <= 0:
                continue
            material_id = str(row.get("material_id", "") or "").strip()
            allocations.append({"material_id": material_id, "quantidade": value})
            selected_ids.add(material_id)
        if not allocations:
            QMessageBox.warning(self, "Dar Baixa", "Define pelo menos uma quantidade para baixa.")
            return None
        retalho = {
            "comprimento": comp_spin.value(),
            "largura": larg_spin.value(),
            "quantidade": qtd_retalho_spin.value(),
            "metros": metros_spin.value(),
        }
        has_retalho = any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
        source_material_id = str(source_combo.currentData() or "").strip()
        if has_retalho and len(selected_ids) > 1 and source_material_id not in selected_ids:
            QMessageBox.warning(self, "Dar Baixa", "Seleciona como origem do retalho um dos lotes efetivamente baixados.")
            return None
        return {
            "allocations": allocations,
            "retalho": retalho if has_retalho else {},
            "source_material_id": source_material_id if has_retalho else "",
        }

    def _montagem_pending_component_rows(self, group: dict) -> list[dict]:
        rows = []
        for index, row in enumerate(list((group or {}).get("montagem_items", []) or [])):
            plan = float(row.get("qtd_planeada", 0) or 0)
            done = float(row.get("qtd_consumida", 0) or 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            item = dict(row)
            item["id"] = str(item.get("id", "") or f"COMP::{str((group or {}).get('encomenda', '') or '').strip()}::{index}").strip()
            item["qtd_pendente"] = pending
            rows.append(item)
        return rows

    def _checked_montagem_component_ids(self, group: dict) -> list[str]:
        pending_ids = {str(row.get("id", "") or "").strip() for row in self._montagem_pending_component_rows(group)}
        return [item_id for item_id in self.checked_piece_ids if item_id in pending_ids]

    def _prompt_montagem_component_selection(self, group: dict) -> list[str] | None:
        enc_num = str((group or {}).get("encomenda", "") or "").strip()
        rows = self._montagem_pending_component_rows(group)
        if not rows:
            return []
        dialog = QDialog(self)
        dialog.setWindowTitle("Consumir componentes")
        dialog.resize(980, 560)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        title = QLabel(f"Seleciona componentes/matéria-prima a consumir da encomenda {enc_num or '-'}")
        title.setStyleSheet("font-size: 16px; font-weight: 900; color: #10253d;")
        layout.addWidget(title)
        hint = QLabel("As linhas marcadas serão baixadas do stock. Se quiseres o atalho, marca linhas na grelha anterior e carrega em Consumir comp.")
        hint.setWordWrap(True)
        hint.setProperty("role", "muted")
        layout.addWidget(hint)
        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["Sel.", "Codigo", "Descricao", "Tipo", "Unid.", "Pendente", "Falta", "Estado"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(30)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setWordWrap(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        table.setColumnWidth(0, 46)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        table.setColumnWidth(1, 130)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        for col, width in ((3, 130), (4, 70), (5, 86), (6, 76), (7, 112)):
            table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Interactive)
            table.setColumnWidth(col, width)
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            item_id = str(row.get("id", "") or "").strip()
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked)
            check_item.setData(Qt.UserRole, item_id)
            check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, 0, check_item)
            values = [
                row.get("codigo", "-") or "-",
                row.get("descricao", "-") or "-",
                f"{'Matéria-prima' if str(row.get('grupo_operador', '') or '') == 'matéria-prima' else 'Componentes'} / {row.get('tipo_label', 'Stock') or 'Stock'}",
                row.get("unidade", "-") or "-",
                f"{float(row.get('qtd_pendente', 0) or 0):.2f}",
                f"{float(row.get('falta', 0) or 0):.2f}",
                row.get("estado", "-") or "-",
            ]
            for offset, value in enumerate(values, start=1):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if offset in (4, 5, 6):
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                table.setItem(row_index, offset, cell)
            _paint_table_row(table, row_index, str(row.get("estado", "")))
        layout.addWidget(table, 1)

        select_all = QCheckBox("Selecionar todas")
        select_all.setChecked(True)

        def _toggle_all(checked: bool) -> None:
            for row_index in range(table.rowCount()):
                item = table.item(row_index, 0)
                if item is not None:
                    item.setCheckState(Qt.Checked if checked else Qt.Unchecked)

        select_all.toggled.connect(_toggle_all)
        footer = QHBoxLayout()
        footer.addWidget(select_all)
        footer.addStretch(1)
        layout.addLayout(footer)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Consumir selecionados")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected_ids = []
        for row_index in range(table.rowCount()):
            item = table.item(row_index, 0)
            if item is not None and item.checkState() == Qt.Checked:
                selected_ids.append(str(item.data(Qt.UserRole) or "").strip())
        return [item_id for item_id in selected_ids if item_id]

    def _required_meters_from_raw_row(self, row: dict) -> float:
        text = " ".join(
            str(row.get(key, "") or "")
            for key in ("descricao", "dimensao", "dimensoes")
        ).replace(",", ".")
        matches = re.findall(r"x\s*([0-9]+(?:\.[0-9]+)?)\s*m\b", text, flags=re.IGNORECASE)
        if matches:
            try:
                return float(matches[-1])
            except Exception:
                return 0.0
        match = re.search(r"\b([0-9]+(?:\.[0-9]+)?)\s*m\b", text, flags=re.IGNORECASE)
        if not match:
            return 0.0
        try:
            return float(match.group(1))
        except Exception:
            return 0.0

    def _prompt_montagem_raw_material_allocations(self, group: dict, selected_ids: list[str]) -> dict[str, Any] | None:
        enc_num = str((group or {}).get("encomenda", "") or "").strip()
        rows_by_id = {
            str(row.get("id", "") or "").strip(): dict(row)
            for row in self._montagem_pending_component_rows(group)
            if str(row.get("grupo_operador", "") or "") == "matéria-prima"
        }
        selected_raw_ids = [item_id for item_id in selected_ids if item_id in rows_by_id]
        if not selected_raw_ids:
            return {}
        options_fn = getattr(self.backend, "operator_montagem_stock_options", None)
        if not callable(options_fn):
            QMessageBox.critical(self, "Matéria-prima", "Backend sem suporte para seleção física de matéria-prima.")
            return None

        payloads: list[dict[str, Any]] = []
        try:
            for item_id in selected_raw_ids:
                payload = dict(options_fn(enc_num, item_id) or {})
                if not list(payload.get("options", []) or []):
                    QMessageBox.warning(self, "Matéria-prima", f"Sem unidades físicas disponíveis para:\n{rows_by_id[item_id].get('descricao', item_id)}")
                    return None
                payloads.append(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Matéria-prima", str(exc))
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar matéria-prima física")
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        dialog.resize(1180, 660)
        dialog.setMinimumSize(980, 560)
        dialog.setStyleSheet(
            "QDialog { background: #eaf2fb; }"
            "QLabel#physicalTitle { font-size: 18px; font-weight: 950; color: #071b3a; }"
            "QLabel#physicalHint { color: #31547d; font-size: 12px; }"
            "QLabel#physicalCardTitle { font-size: 13px; font-weight: 950; color: #071b3a; }"
            "QLabel[role='pill'] { border: 1px solid #b8cbea; border-radius: 4px; padding: 4px 8px; background: #f7fbff; color: #12345f; font-weight: 800; }"
            "QTableWidget { background: #ffffff; border: 1px solid #b8cbea; gridline-color: #d7e2f1; selection-background-color: #eaf7da; selection-color: #26331d; }"
            "QHeaderView::section { background: #050348; color: white; font-weight: 900; padding: 7px 8px; border: none; }"
            "QDoubleSpinBox { min-height: 26px; border: 1px solid #a9bedc; border-radius: 4px; background: #ffffff; padding: 2px 6px; }"
            "QComboBox { min-height: 28px; border: 1px solid #a9bedc; border-radius: 4px; background: #ffffff; padding: 2px 8px; }"
            "QCheckBox { font-weight: 800; color: #10253d; }"
        )
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header_card = CardFrame()
        header_card.set_tone("default")
        header_layout = QGridLayout(header_card)
        header_layout.setContentsMargins(14, 10, 14, 10)
        title = QLabel(f"Baixa física de matéria-prima | Encomenda {enc_num or '-'}")
        title.setObjectName("physicalTitle")
        hint = QLabel("Escolhe a unidade física real utilizada. A baixa fica ligada ao lote e o remanescente/retalho fica rastreável para reutilização.")
        hint.setObjectName("physicalHint")
        hint.setWordWrap(True)
        count_pill = QLabel(f"{len(payloads)} linha(s)")
        count_pill.setProperty("role", "pill")
        header_layout.addWidget(title, 0, 0)
        header_layout.addWidget(count_pill, 0, 1, alignment=Qt.AlignRight)
        header_layout.addWidget(hint, 1, 0, 1, 2)
        layout.addWidget(header_card)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(10)
        editors: dict[str, dict[str, Any]] = {}

        for payload in payloads:
            item_id = str(payload.get("item_id", "") or "").strip()
            row = rows_by_id.get(item_id, {})
            pending = float(payload.get("qtd_pendente", row.get("qtd_pendente", 0)) or 0)
            card = CardFrame()
            card.set_tone("warning")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(14, 12, 14, 12)
            card_layout.setSpacing(8)
            card_header = QGridLayout()
            card_header.setHorizontalSpacing(8)
            header = QLabel(str(row.get("descricao", payload.get("descricao", "-")) or "-"))
            header.setObjectName("physicalCardTitle")
            header.setWordWrap(True)
            pending_pill = QLabel(f"Pendente {pending:.2f}")
            pending_pill.setProperty("role", "pill")
            dim_pill = QLabel(str(payload.get("dimensao", row.get("dimensao", "-")) or "-"))
            dim_pill.setProperty("role", "pill")
            card_header.addWidget(header, 0, 0)
            card_header.addWidget(pending_pill, 0, 1, alignment=Qt.AlignRight)
            card_header.addWidget(dim_pill, 0, 2, alignment=Qt.AlignRight)
            card_layout.addLayout(card_header)

            options = [dict(item or {}) for item in list(payload.get("options", []) or [])]
            table = QTableWidget(len(options), 7)
            table.setHorizontalHeaderLabels(["Unidade física", "Dim. A/B", "Disp.", "Comp./un", "Local", "Tipo", "Baixar"])
            table.verticalHeader().setVisible(False)
            table.verticalHeader().setDefaultSectionSize(32)
            table.setAlternatingRowColors(True)
            table.setSelectionBehavior(QTableWidget.SelectRows)
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            for col, width in ((2, 76), (3, 88), (4, 110), (5, 82), (6, 96)):
                table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Fixed)
                table.setColumnWidth(col, width)
            spinners: list[tuple[dict[str, Any], QDoubleSpinBox]] = []
            default_done = False
            for row_index, option in enumerate(options):
                is_retalho = bool(option.get("is_retalho"))
                values = [
                    option.get("lote", option.get("lote_interno", option.get("material_id", "-"))) or "-",
                    option.get("dimensao", "-") or "-",
                    f"{float(option.get('disponivel', 0) or 0):.2f}",
                    f"{float(option.get('metros', 0) or 0):.2f}" if float(option.get("metros", 0) or 0) > 0 else "-",
                    option.get("local", "-") or "-",
                    "Retalho" if is_retalho else "Stock",
                ]
                for col, value in enumerate(values):
                    cell = QTableWidgetItem(str(value))
                    cell.setToolTip(str(value))
                    if col in (2, 3, 5):
                        cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    table.setItem(row_index, col, cell)
                spin = QDoubleSpinBox()
                max_value = float(option.get("disponivel", 0) or 0)
                spin.setRange(0.0, max_value)
                spin.setDecimals(2)
                spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
                spin.setAlignment(Qt.AlignCenter)
                if not default_done and max_value >= pending:
                    spin.setValue(pending)
                    default_done = True
                table.setCellWidget(row_index, 6, spin)
                spinners.append((option, spin))
            table.setMinimumHeight(_table_visible_height(table, max(2, min(6, len(options))), extra=26))
            card_layout.addWidget(table)

            retalho_box = QCheckBox("Registar remanescente / retalho")
            retalho_layout = QGridLayout()
            retalho_layout.setHorizontalSpacing(8)
            retalho_layout.setVerticalSpacing(6)
            comp_spin = QDoubleSpinBox()
            larg_spin = QDoubleSpinBox()
            qtd_ret_spin = QDoubleSpinBox()
            metros_spin = QDoubleSpinBox()
            for spin in (comp_spin, larg_spin, qtd_ret_spin, metros_spin):
                spin.setRange(0.0, 1000000.0)
                spin.setDecimals(2)
                spin.setAlignment(Qt.AlignCenter)
            source_combo = QComboBox()
            for option in options:
                source_combo.addItem(
                    f"{option.get('material_id', '-')} | {option.get('lote', '-')} | {option.get('dimensao', '-')}",
                    str(option.get("material_id", "") or "").strip(),
                )
            required_m = self._required_meters_from_raw_row(row)
            first_selected = next((option for option, spin in spinners if float(spin.value() or 0) > 0), options[0] if options else {})
            stock_m = float((first_selected or {}).get("metros", 0) or 0)
            if pending == 1 and required_m > 0 and stock_m > required_m:
                retalho_box.setChecked(True)
                metros_spin.setValue(round(stock_m - required_m, 2))
                qtd_ret_spin.setValue(1.0)
                source_combo.setCurrentIndex(0)
            retalho_layout.addWidget(QLabel("Dim. A"), 0, 0)
            retalho_layout.addWidget(QLabel("Dim. B"), 0, 1)
            retalho_layout.addWidget(QLabel("Qtd rem."), 0, 2)
            retalho_layout.addWidget(QLabel("Metros"), 0, 3)
            retalho_layout.addWidget(QLabel("Origem"), 0, 4)
            retalho_layout.addWidget(comp_spin, 1, 0)
            retalho_layout.addWidget(larg_spin, 1, 1)
            retalho_layout.addWidget(qtd_ret_spin, 1, 2)
            retalho_layout.addWidget(metros_spin, 1, 3)
            retalho_layout.addWidget(source_combo, 1, 4)
            card_layout.addWidget(retalho_box)
            card_layout.addLayout(retalho_layout)
            editors[item_id] = {
                "pending": pending,
                "spinners": spinners,
                "retalho_box": retalho_box,
                "comp": comp_spin,
                "larg": larg_spin,
                "qtd_ret": qtd_ret_spin,
                "metros": metros_spin,
                "source": source_combo,
            }
            content_layout.addWidget(card)

        content_layout.addStretch(1)
        scroll.setWidget(content)
        layout.addWidget(scroll, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Confirmar baixa física")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None

        result: dict[str, Any] = {}
        for item_id, editor in editors.items():
            allocations = []
            selected_ids: set[str] = set()
            for option, spin in list(editor.get("spinners", []) or []):
                qty = float(spin.value() or 0)
                if qty <= 0:
                    continue
                material_id = str(option.get("material_id", "") or "").strip()
                allocations.append({"material_id": material_id, "quantidade": qty})
                selected_ids.add(material_id)
            total = round(sum(float(row.get("quantidade", 0) or 0) for row in allocations), 4)
            pending = float(editor.get("pending", 0) or 0)
            if abs(total - pending) > 0.0001:
                QMessageBox.warning(dialog, "Matéria-prima", f"A linha {item_id} tem de totalizar {pending:.2f}; selecionado {total:.2f}.")
                return None
            retalho = {}
            source_id = ""
            if bool(editor["retalho_box"].isChecked()):
                retalho = {
                    "comprimento": float(editor["comp"].value() or 0),
                    "largura": float(editor["larg"].value() or 0),
                    "quantidade": float(editor["qtd_ret"].value() or 0),
                    "metros": float(editor["metros"].value() or 0),
                }
                has_retalho = any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
                if has_retalho:
                    source_id = str(editor["source"].currentData() or "").strip()
                    if source_id not in selected_ids:
                        QMessageBox.warning(dialog, "Matéria-prima", "A origem do retalho tem de ser uma unidade física baixada nessa linha.")
                        return None
                else:
                    retalho = {}
            result[item_id] = {"allocations": allocations, "retalho": retalho, "source_material_id": source_id}
        return result

    def _consume_montagem_components(self) -> None:
        group = self._current_group()
        if not self._is_montagem_stock_group(group):
            QMessageBox.information(self, "Componentes", "Seleciona primeiro o grupo Componentes Montagem/Stock.")
            return
        enc_num = str(group.get("encomenda", "") or "").strip()
        if not enc_num:
            return
        pending_rows = self._montagem_pending_component_rows(group)
        if not pending_rows:
            QMessageBox.information(self, "Componentes", "Nao existem componentes pendentes para consumir.")
            return
        selected_ids = self._checked_montagem_component_ids(group)
        shortcut = bool(selected_ids)
        if not selected_ids:
            selected_ids = self._prompt_montagem_component_selection(group) or []
        if not selected_ids:
            return
        material_allocations = {}
        if str(group.get("montagem_group_kind", "") or "") == "materia_prima":
            allocation_payload = self._prompt_montagem_raw_material_allocations(group, selected_ids)
            if allocation_payload is None:
                return
            material_allocations = allocation_payload
        consume_fn = getattr(self.backend, "operator_consume_montagem_stock", None)
        if not callable(consume_fn):
            QMessageBox.critical(self, "Componentes", "Backend sem suporte para consumo de componentes no operador.")
            return
        try:
            consume_fn(enc_num, operador=self._current_operator(), item_ids=selected_ids, material_allocations=material_allocations)
        except Exception as exc:
            QMessageBox.critical(self, "Componentes", str(exc))
            return
        self.checked_piece_ids.difference_update(set(selected_ids))
        mode_txt = "atalho" if shortcut else "seleção"
        self._set_feedback(f"{len(selected_ids)} componente(s) consumido(s) na encomenda {enc_num} por {mode_txt}.", error=False)
        self.refresh()

    def _manual_consume_material(self) -> None:
        group = self._current_group()
        if self._is_montagem_stock_group(group):
            self._consume_montagem_components()
            return
        material = str(group.get("material", "") or "").strip()
        espessura = str(group.get("espessura", "") or "").strip()
        if not material or not espessura:
            QMessageBox.warning(self, "Dar Baixa", "Seleciona primeiro um grupo com material e espessura.")
            return
        if not self._prompt_supervisor_password():
            return
        payload = self._prompt_manual_material_consumption(material, espessura)
        if payload is None:
            return
        consume_fn = getattr(self.backend, "consume_material_allocations", None)
        if not callable(consume_fn):
            QMessageBox.critical(self, "Dar Baixa", "Backend sem suporte para baixa manual.")
            return
        try:
            result = dict(
                consume_fn(
                    payload.get("allocations", []),
                    retalho=payload.get("retalho", {}),
                    source_material_id=str(payload.get("source_material_id", "") or "").strip(),
                    reason=f"operador_{str(group.get('encomenda', '') or '').strip()}_{material}_{espessura}",
                )
                or {}
            )
        except Exception as exc:
            QMessageBox.critical(self, "Dar Baixa", str(exc))
            return
        message = f"Baixa registada: {float(result.get('consumed_total', 0) or 0):.2f}"
        if str(result.get("retalho_id", "") or "").strip():
            message = f"{message} | Retalho {result.get('retalho_id')}"
        self._set_feedback(message, error=False)
        self.refresh()

    def _prompt_partial_reserved_consumption(self, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Baixa Parcial")
        dialog.resize(980, 620)
        layout = QVBoxLayout(dialog)
        info = QLabel(
            "Seleciona o material cativado que foi parcialmente consumido. "
            "Esta baixa nao fecha o grupo de laser; apenas atualiza o stock e permite registar retalho."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        table = QTableWidget(len(rows), 7)
        table.setHorizontalHeaderLabels(["Sel.", "Material", "Esp.", "Qtd cativada", "Lote", "Dimensao", "Baixar"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(27)
        table.setStyleSheet(
            "QTableWidget { font-size: 11px; } "
            "QHeaderView::section { font-size: 11px; font-weight: 700; } "
            "QDoubleSpinBox { font-size: 11px; } "
            "QDoubleSpinBox#partialConsumeCellSpin { border: none; background: transparent; padding: 0 8px; }"
        )
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.resizeSection(0, 44)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 148)
        checks: list[tuple[dict[str, Any], QCheckBox, QDoubleSpinBox]] = []
        for row_index, row in enumerate(rows):
            check = QCheckBox()
            check.setStyleSheet("margin-left: 10px;")
            table.setCellWidget(row_index, 0, check)
            table.setItem(row_index, 1, QTableWidgetItem(str(row.get("material", "-"))))
            table.setItem(row_index, 2, QTableWidgetItem(str(row.get("espessura", "-"))))
            table.setItem(row_index, 3, QTableWidgetItem(f"{float(row.get('quantidade', 0) or 0):.2f}"))
            table.setItem(row_index, 4, QTableWidgetItem(str(row.get("lote", "-"))))
            table.setItem(row_index, 5, QTableWidgetItem(str(row.get("dimensao", "-"))))
            spin = QDoubleSpinBox()
            spin.setRange(0.0, float(row.get("quantidade", 0) or 0))
            spin.setDecimals(2)
            spin.setSingleStep(1.0)
            spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
            spin.setAlignment(Qt.AlignCenter)
            spin.setMinimumWidth(126)
            spin.setFrame(False)
            spin.setObjectName("partialConsumeCellSpin")
            spin.setValue(float(row.get("quantidade", 0) or 0))
            table.setCellWidget(row_index, 6, spin)
            checks.append((row, check, spin))
        if checks:
            checks[0][1].setChecked(True)

        def _single_select(changed: QCheckBox) -> None:
            if not changed.isChecked():
                return
            for _row, check, _spin in checks:
                if check is not changed:
                    check.setChecked(False)

        for _row, check, _spin in checks:
            check.toggled.connect(lambda _checked, box=check: _single_select(box))
        table.setMinimumHeight(_table_visible_height(table, max(6, len(rows)), extra=24))
        layout.addWidget(table, 1)

        retalho_box = QCheckBox("Sobrou retalho e quero inserir em stock")
        retalho_box.setChecked(False)
        layout.addWidget(retalho_box)
        retalho_card = CardFrame()
        retalho_card.set_tone("warning")
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(12, 10, 12, 10)
        retalho_layout.setHorizontalSpacing(8)
        retalho_layout.setVerticalSpacing(6)
        retalho_layout.addWidget(QLabel("Retalho comprimento"), 0, 0)
        retalho_layout.addWidget(QLabel("Retalho largura"), 0, 1)
        retalho_layout.addWidget(QLabel("Qtd retalho"), 0, 2)
        retalho_layout.addWidget(QLabel("Metros"), 0, 3)
        comp_spin = QDoubleSpinBox()
        larg_spin = QDoubleSpinBox()
        qtd_retalho_spin = QDoubleSpinBox()
        metros_spin = QDoubleSpinBox()
        for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
            spin.setRange(0.0, 1000000.0)
            spin.setDecimals(2)
            spin.setAlignment(Qt.AlignCenter)
            spin.setEnabled(False)
        retalho_layout.addWidget(comp_spin, 1, 0)
        retalho_layout.addWidget(larg_spin, 1, 1)
        retalho_layout.addWidget(qtd_retalho_spin, 1, 2)
        retalho_layout.addWidget(metros_spin, 1, 3)
        layout.addWidget(retalho_card)

        def _toggle_retalho(enabled: bool) -> None:
            for spin in (comp_spin, larg_spin, qtd_retalho_spin, metros_spin):
                spin.setEnabled(enabled)

        retalho_box.toggled.connect(_toggle_retalho)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Confirmar baixa parcial")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        selected = next(((row, spin) for row, check, spin in checks if check.isChecked()), None)
        if selected is None:
            QMessageBox.warning(self, "Baixa Parcial", "Seleciona um material cativado.")
            return None
        row, qty_spin = selected
        quantity = float(qty_spin.value() or 0)
        if quantity <= 0:
            QMessageBox.warning(self, "Baixa Parcial", "Indica uma quantidade para dar baixa.")
            return None
        retalho = {
            "comprimento": comp_spin.value(),
            "largura": larg_spin.value(),
            "quantidade": qtd_retalho_spin.value(),
            "metros": metros_spin.value(),
        }
        has_retalho = bool(retalho_box.isChecked()) and any(float(retalho[key] or 0) > 0 for key in ("comprimento", "largura", "quantidade", "metros"))
        return {
            "material_id": str(row.get("material_id", "") or "").strip(),
            "material": str(row.get("material", "") or "").strip(),
            "espessura": str(row.get("espessura", "") or "").strip(),
            "quantity": quantity,
            "retalho": retalho if has_retalho else {},
            "source_material_id": str(row.get("material_id", "") or "").strip() if has_retalho else "",
        }

    def _partial_consume_reserved_material(self) -> None:
        group = self._current_group()
        enc_num = str(group.get("encomenda", "") or "").strip()
        material = str(group.get("material", "") or "").strip()
        espessura = str(group.get("espessura", "") or "").strip()
        if not enc_num:
            QMessageBox.warning(self, "Baixa Parcial", "Seleciona primeiro uma encomenda/grupo.")
            return
        if not self._prompt_supervisor_password():
            return
        rows_fn = getattr(self.backend, "operator_reserved_materials", None)
        consume_fn = getattr(self.backend, "operator_partial_reserved_material_consumption", None)
        if not callable(rows_fn) or not callable(consume_fn):
            QMessageBox.critical(self, "Baixa Parcial", "Backend sem suporte para baixa parcial.")
            return
        try:
            rows = list(rows_fn(enc_num, material, espessura) or [])
        except Exception as exc:
            QMessageBox.critical(self, "Baixa Parcial", str(exc))
            return
        if not rows:
            QMessageBox.information(self, "Baixa Parcial", "Esta encomenda/grupo nao tem material cativado para baixar parcialmente.")
            return
        payload = self._prompt_partial_reserved_consumption(rows)
        if payload is None:
            return
        try:
            result = dict(
                consume_fn(
                    enc_num,
                    material_id=str(payload.get("material_id", "") or "").strip(),
                    quantidade=float(payload.get("quantity", 0) or 0),
                    material=str(payload.get("material", "") or "").strip(),
                    espessura=str(payload.get("espessura", "") or "").strip(),
                    retalho=payload.get("retalho", {}),
                    source_material_id=str(payload.get("source_material_id", "") or "").strip(),
                )
                or {}
            )
        except Exception as exc:
            QMessageBox.critical(self, "Baixa Parcial", str(exc))
            return
        message = f"Baixa parcial registada: {float(result.get('consumed_total', 0) or 0):.2f}"
        if str(result.get("retalho_id", "") or "").strip():
            message = f"{message} | Retalho {result.get('retalho_id')}"
        self._set_feedback(message, error=False)
        self.refresh()

    def _maybe_close_material_session(
        self,
        enc_num: str,
        material: str,
        espessura: str,
        operator_name: str,
        trigger: str,
    ) -> bool:
        status_fn = getattr(self.backend, "operator_material_session_state", None)
        resolve_fn = getattr(self.backend, "operator_resolve_laser_stock", None)
        if not callable(status_fn) or not callable(resolve_fn):
            return True
        try:
            stock_state = dict(status_fn(enc_num, material, espessura, operator_name) or {})
        except Exception as exc:
            self._set_feedback(str(exc), error=True)
            QMessageBox.critical(self, "Baixa material", str(exc))
            return False
        if not stock_state or not stock_state.get("should_prompt"):
            return True
        decision = self._prompt_material_session_decision(stock_state)
        if decision == "cancel":
            self._set_feedback("Operação fechada; decisão sobre o material cativado pendente.", error=False)
            return True
        if decision == "keep":
            record_fn = getattr(self.backend, "operator_record_material_session_decision", None)
            if callable(record_fn):
                try:
                    record_fn(enc_num, material, espessura, operator_name, "manter_cativado", trigger)
                except Exception as exc:
                    self._set_feedback(str(exc), error=True)
                    QMessageBox.critical(self, "Material cativado", str(exc))
                    return False
            self._set_feedback("Material mantido cativado para continuação posterior.", error=False)
            return True
        stock_state["session_close"] = True
        payload = self._prompt_laser_stock_resolution(stock_state)
        if payload is None:
            self._set_feedback("Operação fechada; baixa do material cativado não registada.", error=False)
            return False
        try:
            result = dict(
                resolve_fn(
                    enc_num,
                    material,
                    espessura,
                    material_id=str(payload.get("material_id", "") or "").strip(),
                    quantidade=float(payload.get("quantity", 0) or 0),
                    allow_without_stock=bool(payload.get("allow_without_stock")),
                    retalho=payload.get("retalho", {}),
                    source_material_id=str(payload.get("source_material_id", "") or "").strip(),
                    session_close=True,
                    operator_name=operator_name,
                )
                or {}
            )
        except Exception as exc:
            self._set_feedback(str(exc), error=True)
            QMessageBox.critical(self, "Baixa material", str(exc))
            return False
        consumed = float(result.get("consumed_total", 0) or 0)
        remaining = float(result.get("remaining_qty", 0) or 0)
        if remaining > 0 and bool(result.get("allow_without_stock")):
            self._set_feedback(f"Laser concluido. Baixa pendente assumida: {remaining:.1f}.", error=False)
        elif consumed > 0:
            message = f"Baixa material registada: {consumed:.1f}."
            remaining_reserved = float(result.get("remaining_reserved", 0) or 0)
            if remaining_reserved > 0:
                message = f"{message} Permanecem {remaining_reserved:.1f} cativadas."
            if str(result.get("retalho_id", "") or "").strip():
                message = f"{message} Retalho {result.get('retalho_id')}."
            self._set_feedback(message, error=False)
        elif str(result.get("retalho_id", "") or "").strip():
            self._set_feedback(f"Retalho registado: {result.get('retalho_id')}.", error=False)
        return True

    def _run_action(self, action_fn, success_text: str, *, allow_multiple: bool = True) -> None:
        refs_list = self._selected_piece_refs(allow_multiple=allow_multiple)
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        for enc_num, piece_id in refs_list:
            try:
                action_fn(enc_num, piece_id)
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if applied > 1:
            self._set_feedback(f"{success_text} ({applied} pecas)", error=False)
        else:
            self._set_feedback(success_text, error=False)

    def _start_piece(self) -> None:
        operator_name = self._current_operator()
        if not operator_name:
            QMessageBox.warning(self, "Operador", "Seleciona o operador antes de iniciar.")
            return
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        collected = self._collect_start_operation_targets(refs_list)
        ordered_ops = list(collected.get("ordered_ops", []) or [])
        targets = dict(collected.get("targets", {}) or {})
        skipped_without_ops = int(collected.get("skipped_without_ops", 0) or 0)
        context_errors = [str(err) for err in list(collected.get("errors", []) or []) if str(err).strip()]
        if not ordered_ops:
            message = context_errors[0] if context_errors else "As pecas selecionadas nao têm operacoes pendentes para iniciar."
            self._set_feedback(message, error=True)
            QMessageBox.warning(self, "Operador", message)
            return
        operation = self._prompt_start_operation(refs_list, ordered_ops, targets, skipped_without_ops=skipped_without_ops)
        if not operation:
            return
        self.operation_combo.setCurrentText(operation)
        compatible_refs = list(targets.get(operation, []) or [])
        skipped_for_choice = max(0, len(refs_list) - len(compatible_refs))
        if not compatible_refs:
            QMessageBox.warning(self, "Operador", "Nenhuma das pecas selecionadas tem essa operacao pendente.")
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = compatible_refs[0][1]
        errors = list(context_errors)
        applied = 0
        for enc_num, piece_id in compatible_refs:
            try:
                self.backend.operator_start_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    operation=operation,
                    posto=self._current_posto(),
                )
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            if applied:
                message = f"Operacao iniciada em {applied} peca(s), mas houve falhas:\n{message}"
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        success = f"Operacao iniciada: {operation}"
        if applied > 1:
            success = f"{success} ({applied} pecas)"
        if skipped_for_choice:
            success = f"{success} | Ignoradas {skipped_for_choice} sem esta operacao"
        self._set_feedback(success, error=False)

    def _finish_piece(self) -> None:
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        operator_name = self._current_operator()
        self.selected_group_key = self._group_key(self._current_group())
        errors: list[str] = []
        completed = 0
        finished_ops: list[str] = []
        batch_total = len(refs_list)
        current_group = self._current_group()
        group_material = str(current_group.get("material", "") or "").strip()
        group_esp = str(current_group.get("espessura", "") or "").strip()
        group_enc = str(current_group.get("encomenda", "") or "").strip()
        for batch_idx, (enc_num, piece_id) in enumerate(refs_list, start=1):
            try:
                ctx = self.backend.operator_piece_context(enc_num, piece_id)
            except Exception as exc:
                errors.append(str(exc))
                continue
            payload = self._prompt_finish(
                ctx,
                batch_idx=batch_idx if batch_total > 1 else 0,
                batch_total=batch_total if batch_total > 1 else 0,
                preferred_operation=self._current_operation(),
            )
            if payload is None:
                break
            self.selected_piece_id = piece_id
            try:
                result = self.backend.operator_finish_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    payload["ok"],
                    payload["nok"],
                    payload["qual"],
                    operation=str(payload["operation"] or ""),
                    posto=self._current_posto(),
                )
                completed += 1
                op_name = str((result or {}).get("operation", "") or payload.get("operation", "") or "").strip()
                if op_name:
                    finished_ops.append(op_name)
            except Exception as exc:
                errors.append(str(exc))
        if completed and group_enc and any("laser" in str(op or "").lower() for op in finished_ops):
            self._maybe_close_material_session(
                group_enc,
                group_material,
                group_esp,
                operator_name,
                "conclusao",
            )
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if completed > 1:
            self._set_feedback(f"Operacoes concluidas com sucesso ({completed} pecas).", error=False)
        elif completed == 1:
            self._set_feedback("Operacao concluida com sucesso.", error=False)

    def _resume_piece(self) -> None:
        operator_name = self._current_operator()
        self._run_action(
            lambda enc_num, piece_id: self.backend.operator_resume_piece(enc_num, piece_id, operator_name, posto=self._current_posto()),
            "Peca retomada.",
        )

    def _pause_piece(self) -> None:
        reason = self._prompt_reason(
            "Interromper peca",
            "Seleciona ou escreve o motivo da interrupcao.",
            self.backend.operator_interruption_options(),
        )
        if reason is None:
            return
        operator_name = self._current_operator()
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        group = self._current_group()
        group_material = str(group.get("material", "") or "").strip()
        group_esp = str(group.get("espessura", "") or "").strip()
        group_enc = str(group.get("encomenda", "") or "").strip()
        laser_context = False
        errors: list[str] = []
        applied = 0
        for enc_num, piece_id in refs_list:
            try:
                ctx = dict(self.backend.operator_piece_context(enc_num, piece_id) or {})
                active_ops = list(ctx.get("active_pending_ops", []) or [])
                laser_context = laser_context or any("laser" in str(op or "").lower() for op in active_ops)
                self.backend.operator_pause_piece(
                    enc_num,
                    piece_id,
                    operator_name,
                    reason,
                    posto=self._current_posto(),
                )
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        if applied and laser_context and group_enc:
            self._maybe_close_material_session(
                group_enc,
                group_material,
                group_esp,
                operator_name,
                "interrupcao",
            )
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        suffix = f" ({applied} pecas)" if applied > 1 else ""
        self._set_feedback(f"Interrupcao registada: {reason}{suffix}", error=False)

    def _register_avaria(self) -> None:
        reason = self._prompt_reason(
            "Registar avaria",
            "Seleciona ou escreve a causa da avaria.",
            self.backend.operator_avaria_options(),
        )
        if reason is None:
            return
        operator_name = self._current_operator()
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        shared_group_id = ""
        shared_started_at = ""
        for enc_num, piece_id in refs_list:
            try:
                result = self.backend.operator_register_avaria(
                    enc_num,
                    piece_id,
                    operator_name,
                    reason,
                    posto=self._current_posto(),
                    group_id=shared_group_id,
                    ts_now=shared_started_at,
                )
                if not shared_group_id:
                    shared_group_id = str((result or {}).get("avaria_group_key", "") or "").strip()
                if not shared_started_at:
                    shared_started_at = str((result or {}).get("avaria_started_at", "") or "").strip()
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        if applied > 1:
            self._set_feedback(f"Avaria aberta: {reason} ({applied} pecas)", error=False)
        else:
            self._set_feedback(f"Avaria aberta: {reason}", error=False)

    def _close_avaria(self) -> None:
        operator_name = self._current_operator()
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        self.selected_group_key = self._group_key(self._current_group())
        self.selected_piece_id = refs_list[0][1]
        errors: list[str] = []
        applied = 0
        group_minutes: dict[str, float] = {}
        for enc_num, piece_id in refs_list:
            try:
                result = self.backend.operator_close_avaria(enc_num, piece_id, operator_name, posto=self._current_posto())
                group_key = str((result or {}).get("avaria_group_key", "") or "").strip() or piece_id
                group_minutes[group_key] = max(
                    group_minutes.get(group_key, 0.0),
                    float((result or {}).get("duracao_avaria_min", 0) or 0),
                )
                applied += 1
            except Exception as exc:
                errors.append(str(exc))
        self.refresh()
        if errors:
            message = errors[0] if len(errors) == 1 else "\n".join(errors[:5])
            self._set_feedback(message, error=True)
            QMessageBox.critical(self, "Operador", message)
            return
        total_minutes = sum(group_minutes.values())
        if applied > 1:
            self._set_feedback(f"Avaria fechada ({applied} pecas) | Tempo de paragem {total_minutes:.1f} min", error=False)
        else:
            self._set_feedback(f"Avaria fechada. Tempo {total_minutes:.1f} min", error=False)

    def _alert_chefia(self) -> None:
        operator_name = self._current_operator()
        self._run_action(
            lambda enc_num, piece_id: self.backend.operator_alert_chefia(enc_num, piece_id, operator_name, posto=self._current_posto()),
            "Poke enviado para a chefia.",
        )

    def _open_drawing(self) -> None:
        refs_list = self._selected_piece_refs()
        if refs_list is None:
            return
        opened = 0
        last_drawing = ""
        for enc_num, piece_id in refs_list:
            try:
                last_drawing = str(self.backend.operator_open_drawing(enc_num, piece_id))
                opened += 1
            except Exception as exc:
                self._set_feedback(str(exc), error=True)
                QMessageBox.critical(self, "Ver desenho", str(exc))
                return
        if opened > 1:
            self._set_feedback(f"Desenhos abertos: {opened}", error=False)
        else:
            drawing_name = Path(str(last_drawing or "")).name or str(last_drawing or "")
            self._set_feedback(f"Desenho aberto: {drawing_name}", error=False)
            self.feedback_label.setToolTip(str(last_drawing or ""))

    def _open_labels_dialog(self) -> None:
        group = self._current_group()
        enc_num = str(group.get("encomenda", "") or "").strip()
        if not enc_num:
            QMessageBox.warning(self, "Etiquetas", "Seleciona primeiro uma encomenda com pecas.")
            return
        dialog = _OperatorLabelsDialog(
            self.backend,
            enc_num,
            current_posto=self._current_posto(),
            preselected_ids=self._selected_piece_ids(),
            parent=self,
        )
        dialog.exec()


class _OperatorLabelsDialog(QDialog):
    def __init__(self, backend, order_number: str, current_posto: str = "Geral", preselected_ids: list[str] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.order_number = str(order_number or "").strip()
        self.selected_ids: set[str] = {str(value or "").strip() for value in list(preselected_ids or []) if str(value or "").strip()}
        self.rows: list[dict] = []
        self.filtered_rows: list[dict] = []
        self._syncing_checks = False

        self.setWindowTitle(f"Etiquetas | {self.order_number or '-'}")
        self.resize(1180, 760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro_card = CardFrame()
        intro_card.set_tone("info")
        intro_layout = QVBoxLayout(intro_card)
        intro_layout.setContentsMargins(14, 12, 14, 12)
        intro_layout.setSpacing(4)
        title = QLabel("Etiquetas do Operador")
        title.setStyleSheet("font-size: 17px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel(
            "Seleciona as referencias da encomenda para imprimir etiqueta por unidade 110x50 ou etiqueta de palete A4. "
            "A etiqueta de palete agrupa automaticamente por proximo posto."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        intro_layout.addWidget(title)
        intro_layout.addWidget(subtitle)
        layout.addWidget(intro_card)

        filters_card = CardFrame()
        filters_card.set_tone("default")
        filters_layout = QGridLayout(filters_card)
        filters_layout.setContentsMargins(12, 10, 12, 10)
        filters_layout.setHorizontalSpacing(8)
        filters_layout.setVerticalSpacing(6)
        filters_layout.addWidget(QLabel("Posto de origem"), 0, 0)
        self.source_posto_combo = QComboBox()
        self.source_posto_combo.setProperty("compact", "true")
        posto_source_getter = getattr(self.backend, "operator_posto_options", None)
        if callable(posto_source_getter):
            try:
                self.source_posto_combo.addItems(list(posto_source_getter() or ["Geral"]))
            except Exception:
                self.source_posto_combo.addItems(["Geral"])
        else:
            self.source_posto_combo.addItems(list(self.backend.available_postos() or ["Geral"]))
        self.source_posto_combo.setCurrentText(str(current_posto or "").strip() or "Geral")
        self.source_posto_combo.currentTextChanged.connect(self._reload_rows)
        filters_layout.addWidget(self.source_posto_combo, 1, 0)
        filters_layout.addWidget(QLabel("Filtrar"), 0, 1)
        self.search_edit = QLineEdit()
        self.search_edit.setProperty("compact", "true")
        self.search_edit.setPlaceholderText("Pesquisar ref., OPP, descricao ou posto...")
        self.search_edit.textChanged.connect(self._render_rows)
        filters_layout.addWidget(self.search_edit, 1, 1)
        filters_layout.addWidget(QLabel("Proximo posto"), 0, 2)
        self.destination_combo = QComboBox()
        self.destination_combo.setProperty("compact", "true")
        self.destination_combo.currentTextChanged.connect(self._render_rows)
        filters_layout.addWidget(self.destination_combo, 1, 2)
        _cap_width(self.source_posto_combo, 170)
        _cap_width(self.destination_combo, 180)
        layout.addWidget(filters_card)

        selection_row = QHBoxLayout()
        selection_row.setSpacing(8)
        self.select_visible_btn = QPushButton("Selecionar visiveis")
        self.select_visible_btn.setProperty("variant", "secondary")
        self.select_visible_btn.clicked.connect(lambda: self._apply_visible_selection(True))
        self.clear_visible_btn = QPushButton("Limpar visiveis")
        self.clear_visible_btn.setProperty("variant", "secondary")
        self.clear_visible_btn.clicked.connect(lambda: self._apply_visible_selection(False))
        self.selection_chip = QLabel("0 selecionadas")
        _apply_state_chip(self.selection_chip, "-", "0 selecionadas")
        selection_row.addWidget(self.select_visible_btn)
        selection_row.addWidget(self.clear_visible_btn)
        selection_row.addStretch(1)
        selection_row.addWidget(self.selection_chip)
        layout.addLayout(selection_row)

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(["Sel.", "Ref. Int.", "Ref. Ext.", "Descricao", "OPP", "Qtd", "Estado", "Origem", "Proximo posto"])
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        _configure_table(self.table, stretch=(3, 8), contents=(0, 1, 4, 5, 6, 7))
        header = self.table.horizontalHeader()
        for col, width in ((0, 34), (1, 128), (2, 118), (3, 320), (4, 112), (5, 62), (6, 96), (7, 108), (8, 150)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        self.table.itemChanged.connect(self._handle_check_change)
        layout.addWidget(self.table, 1)

        actions_host = QWidget()
        actions_layout = QVBoxLayout(actions_host)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)

        unit_card = CardFrame()
        unit_card.set_tone("info")
        unit_layout = QHBoxLayout(unit_card)
        unit_layout.setContentsMargins(12, 10, 12, 10)
        unit_layout.setSpacing(8)
        unit_info = QLabel("Etiqueta por unidade 110x50")
        unit_info.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
        unit_hint = QLabel("Uma etiqueta por OPP selecionada.")
        unit_hint.setProperty("role", "muted")
        unit_block = QVBoxLayout()
        unit_block.setSpacing(2)
        unit_block.addWidget(unit_info)
        unit_block.addWidget(unit_hint)
        self.preview_unit_btn = QPushButton("Pre-visualizar")
        self.print_unit_btn = QPushButton("Imprimir")
        self.save_unit_btn = QPushButton("Guardar PDF")
        self.preview_unit_btn.setProperty("variant", "secondary")
        self.print_unit_btn.setProperty("variant", "secondary")
        self.save_unit_btn.setProperty("variant", "secondary")
        self.preview_unit_btn.clicked.connect(self._preview_unit_labels)
        self.print_unit_btn.clicked.connect(self._print_unit_labels)
        self.save_unit_btn.clicked.connect(self._save_unit_labels)
        unit_layout.addLayout(unit_block, 1)
        unit_layout.addWidget(self.preview_unit_btn)
        unit_layout.addWidget(self.print_unit_btn)
        unit_layout.addWidget(self.save_unit_btn)
        actions_layout.addWidget(unit_card)

        pallet_card = CardFrame()
        pallet_card.set_tone("default")
        pallet_layout = QHBoxLayout(pallet_card)
        pallet_layout.setContentsMargins(12, 10, 12, 10)
        pallet_layout.setSpacing(8)
        pallet_info = QLabel("Etiqueta de palete A4")
        pallet_info.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
        pallet_hint = QLabel("Agrupa as referencias selecionadas por proximo posto no mesmo PDF.")
        pallet_hint.setProperty("role", "muted")
        pallet_block = QVBoxLayout()
        pallet_block.setSpacing(2)
        pallet_block.addWidget(pallet_info)
        pallet_block.addWidget(pallet_hint)
        self.preview_pallet_btn = QPushButton("Pre-visualizar")
        self.print_pallet_btn = QPushButton("Imprimir")
        self.save_pallet_btn = QPushButton("Guardar PDF")
        self.preview_pallet_btn.setProperty("variant", "secondary")
        self.print_pallet_btn.setProperty("variant", "secondary")
        self.save_pallet_btn.setProperty("variant", "secondary")
        self.preview_pallet_btn.clicked.connect(self._preview_pallet_labels)
        self.print_pallet_btn.clicked.connect(self._print_pallet_labels)
        self.save_pallet_btn.clicked.connect(self._save_pallet_labels)
        pallet_layout.addLayout(pallet_block, 1)
        pallet_layout.addWidget(self.preview_pallet_btn)
        pallet_layout.addWidget(self.print_pallet_btn)
        pallet_layout.addWidget(self.save_pallet_btn)
        actions_layout.addWidget(pallet_card)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        actions_layout.addLayout(close_row)
        layout.addWidget(actions_host)

        self._reload_rows()

    def _selected_piece_ids(self) -> list[str]:
        ordered_ids: list[str] = []
        seen: set[str] = set()
        for row in self.rows:
            piece_id = str(row.get("piece_id", "") or "").strip()
            if piece_id and piece_id in self.selected_ids and piece_id not in seen:
                ordered_ids.append(piece_id)
                seen.add(piece_id)
        return ordered_ids

    def _reload_rows(self) -> None:
        try:
            payload = dict(self.backend.operator_label_rows(self.order_number, source_posto=self.source_posto_combo.currentText().strip()) or {})
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))
            self.rows = []
            self.filtered_rows = []
            self.table.setRowCount(0)
            self._sync_selection_state()
            return
        self.rows = list(payload.get("rows", []) or [])
        valid_ids = {str(row.get("piece_id", "") or "").strip() for row in self.rows}
        self.selected_ids.intersection_update(valid_ids)
        current_filter = self.destination_combo.currentText().strip()
        destinations = ["Todos"] + sorted({str(row.get("proximo_posto", "") or "-").strip() or "-" for row in self.rows})
        self.destination_combo.blockSignals(True)
        self.destination_combo.clear()
        self.destination_combo.addItems(destinations)
        if current_filter and current_filter in destinations:
            self.destination_combo.setCurrentText(current_filter)
        self.destination_combo.blockSignals(False)
        self._render_rows()

    def _render_rows(self) -> None:
        query = self.search_edit.text().strip().lower()
        destination = self.destination_combo.currentText().strip()
        rows = []
        for row in self.rows:
            if destination and destination != "Todos" and str(row.get("proximo_posto", "") or "-").strip() != destination:
                continue
            if query:
                haystack = " ".join(
                    [
                        str(row.get("ref_interna", "") or ""),
                        str(row.get("ref_externa", "") or ""),
                        str(row.get("descricao", "") or ""),
                        str(row.get("opp", "") or ""),
                        str(row.get("posto_origem", "") or ""),
                        str(row.get("proximo_posto", "") or ""),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            rows.append(row)
        self.filtered_rows = rows
        self._syncing_checks = True
        try:
            self.table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                piece_id = str(row.get("piece_id", "") or "").strip()
                check_item = QTableWidgetItem("")
                check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                check_item.setCheckState(Qt.Checked if piece_id in self.selected_ids else Qt.Unchecked)
                check_item.setData(Qt.UserRole, piece_id)
                check_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, 0, check_item)
                values = [
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("descricao", "") or "-",
                    row.get("opp", "-"),
                    row.get("quantidade_txt", "0"),
                    row.get("estado", "-"),
                    row.get("posto_origem", "-"),
                    row.get("proximo_posto", "-"),
                ]
                for column, value in enumerate(values, start=1):
                    item = QTableWidgetItem(str(value))
                    if column in (5, 6, 7, 8):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    item.setToolTip(str(value))
                    self.table.setItem(row_index, column, item)
                _paint_table_row(self.table, row_index, str(row.get("estado", "") or ""))
        finally:
            self._syncing_checks = False
        self._sync_selection_state()

    def _handle_check_change(self, item: QTableWidgetItem) -> None:
        if self._syncing_checks or item.column() != 0:
            return
        piece_id = str(item.data(Qt.UserRole) or "").strip()
        if not piece_id:
            return
        if item.checkState() == Qt.Checked:
            self.selected_ids.add(piece_id)
        else:
            self.selected_ids.discard(piece_id)
        self._sync_selection_state()

    def _apply_visible_selection(self, selected: bool) -> None:
        self._syncing_checks = True
        try:
            for row_index, row in enumerate(self.filtered_rows):
                piece_id = str(row.get("piece_id", "") or "").strip()
                if not piece_id:
                    continue
                if selected:
                    self.selected_ids.add(piece_id)
                else:
                    self.selected_ids.discard(piece_id)
                item = self.table.item(row_index, 0)
                if item is not None:
                    item.setCheckState(Qt.Checked if selected else Qt.Unchecked)
        finally:
            self._syncing_checks = False
        self._sync_selection_state()

    def _sync_selection_state(self) -> None:
        selected_count = len(self._selected_piece_ids())
        _apply_state_chip(self.selection_chip, "Em producao" if selected_count else "-", f"{selected_count} selecionadas")
        enabled = selected_count > 0
        for button in (
            self.preview_unit_btn,
            self.print_unit_btn,
            self.save_unit_btn,
            self.preview_pallet_btn,
            self.print_pallet_btn,
            self.save_pallet_btn,
        ):
            button.setEnabled(enabled)

    def _build_pdf(self, kind: str, output_path: str | None = None):
        piece_ids = self._selected_piece_ids()
        if not piece_ids:
            raise ValueError("Seleciona pelo menos uma referencia.")
        source_posto = self.source_posto_combo.currentText().strip() or "Geral"
        if kind == "unit":
            return self.backend.operator_unit_labels_pdf(self.order_number, piece_ids, source_posto=source_posto, output_path=output_path)
        return self.backend.operator_pallet_labels_pdf(self.order_number, piece_ids, source_posto=source_posto, output_path=output_path)

    def _preview_pdf(self, kind: str) -> None:
        try:
            path = self._build_pdf(kind)
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))

    def _print_pdf(self, kind: str) -> None:
        try:
            path = self._build_pdf(kind)
            try:
                os.startfile(str(path), "print")
            except Exception:
                os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiquetas", str(exc))

    def _save_pdf(self, kind: str) -> None:
        default_name = f"etiquetas_{self.order_number}_{'unit' if kind == 'unit' else 'palete'}.pdf"
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", default_name, "PDF (*.pdf)")
        if not path:
            return
        try:
            self._build_pdf(kind, output_path=path)
        except Exception as exc:
            QMessageBox.critical(self, "Guardar PDF", str(exc))
            return
        QMessageBox.information(self, "Guardar PDF", f"PDF guardado em:\n{path}")

    def _preview_unit_labels(self) -> None:
        self._preview_pdf("unit")

    def _print_unit_labels(self) -> None:
        self._print_pdf("unit")

    def _save_unit_labels(self) -> None:
        self._save_pdf("unit")

    def _preview_pallet_labels(self) -> None:
        self._preview_pdf("pallet")

    def _print_pallet_labels(self) -> None:
        self._print_pdf("pallet")

    def _save_pallet_labels(self) -> None:
        self._save_pdf("pallet")


class LegacyOperatorPage(OperatorPage):
    page_subtitle = "Encomendas ativas primeiro e detalhe operacional apenas dentro da encomenda."

    def __init__(self, runtime_service, backend, parent=None) -> None:
        super().__init__(runtime_service, backend, parent)
        self.order_rows: list[dict] = []
        self.order_rows_all: list[dict] = []
        self.selected_order_number = ""
        for card in self.cards:
            card.setMinimumHeight(70)
            card.setMaximumHeight(78)
            if card.layout() is not None:
                card.layout().setContentsMargins(12, 9, 12, 9)
                card.layout().setSpacing(4)
            card.title_label.setWordWrap(True)
            card.subtitle_label.setWordWrap(True)
            card.title_label.setStyleSheet("font-size: 9px;")
            card.value_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
            card.subtitle_label.setStyleSheet("font-size: 9px;")
        self.global_progress.setMaximumHeight(20)
        self.groups_table.verticalHeader().setDefaultSectionSize(24)
        self.pieces_table.verticalHeader().setDefaultSectionSize(22)
        self.groups_table.setColumnCount(7)
        self.groups_table.setHorizontalHeaderLabels(["Enc.", "Cli.", "Estado", "Mat.", "Esp.", "Real", "%"])
        self.groups_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.groups_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.pieces_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        group_header = self.groups_table.horizontalHeader()
        group_header.setStretchLastSection(True)
        for col, width in ((0, 132), (1, 84), (2, 106), (3, 94), (4, 56), (5, 64), (6, 74)):
            group_header.setSectionResizeMode(col, QHeaderView.Interactive)
            group_header.resizeSection(col, width)
        self.control_card.setMinimumHeight(178)
        self.control_card.setMaximumHeight(178)
        self.context_card.setMinimumHeight(178)
        self.context_card.setMaximumHeight(178)
        self.feedback_label.setMaximumHeight(54)
        for widget, width in ((self.operator_combo, 176), (self.posto_combo, 128), (self.operation_combo, 292)):
            widget.setProperty("compact", "true")
            widget.setMinimumWidth(width)
        for button in (self.start_btn, self.finish_btn, self.resume_btn, self.pause_btn, self.avaria_btn, self.close_avaria_btn, self.alert_btn, self.manual_consume_btn, self.consume_components_btn, self.partial_consume_btn, self.drawing_btn, self.labels_btn, self.local_refresh_btn):
            button.setProperty("compact", "true")
            button.setMinimumWidth(0)
            button.setMinimumHeight(28)
        self.select_all_pieces_box.setProperty("compact", "true")
        self.piece_state_chip.setMinimumWidth(118)
        self.piece_state_chip.setAlignment(Qt.AlignCenter)
        self.piece_title_label.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        self.piece_title_label.setMinimumHeight(22)
        self.piece_meta_label.setStyleSheet("font-size: 9.2px; color: #334155;")
        self.pending_label.setStyleSheet("font-size: 8.8px; color: #5b6f86;")
        context_layout = self.context_card.layout()
        context_layout.setContentsMargins(16, 12, 16, 12)
        context_layout.setSpacing(5)
        self.issue_label.setMaximumHeight(28)
        self.issue_label.setStyleSheet("font-size: 8.3px; color: #475467;")
        self.piece_meta_label.setWordWrap(True)
        self.pending_label.setWordWrap(True)
        self.pending_label.setMaximumHeight(34)
        self.operation_strip.setMaximumHeight(40)
        self.piece_progress.setMaximumHeight(18)
        self.operator_combo.setMinimumWidth(112)
        self.posto_combo.setMinimumWidth(82)
        self.operation_combo.setMinimumWidth(258)
        for combo in (self.operator_combo, self.posto_combo):
            combo.setStyleSheet(
                "font-size: 9.4px; padding-right: 20px;"
                " QComboBox::drop-down { width: 22px; subcontrol-origin: padding; subcontrol-position: top right; }"
            )
        self.operation_combo.setStyleSheet(
            "font-size: 8.4px; padding-right: 44px;"
            " QComboBox::drop-down { width: 36px; subcontrol-origin: padding; subcontrol-position: center right; border-left: 1px solid #d5dfea; background: #f8fbff; }"
            " QComboBox::down-arrow { image: none; width: 0px; height: 0px; border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 6px solid #475467; margin-right: 10px; }"
        )
        piece_header = self.pieces_table.horizontalHeader()
        for col, width in ((0, 34), (1, 140), (2, 320), (3, 96), (4, 126), (5, 104), (6, 74), (7, 66), (8, 66), (9, 126), (10, 360)):
            piece_header.setSectionResizeMode(col, QHeaderView.Interactive)
            piece_header.resizeSection(col, width)
        piece_header.setStretchLastSection(True)
        group_header = self.groups_table.horizontalHeader()
        for col, width in ((0, 138), (1, 116), (2, 108), (3, 138), (4, 60), (5, 62), (6, 62), (7, 70), (8, 86)):
            group_header.setSectionResizeMode(col, QHeaderView.Interactive)
            group_header.resizeSection(col, width)
        self.pieces_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pieces_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.pieces_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.groups_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.groups_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.groups_table.setStyleSheet("font-size: 11.5px;")
        self.pieces_table.setStyleSheet("font-size: 11.5px;")

        control_layout = self.control_card.layout()
        control_layout.setContentsMargins(15, 11, 15, 11)
        control_layout.setSpacing(7)
        control_items = _take_layout_items(control_layout)
        selectors_item = control_items[0] if len(control_items) > 0 else None
        feedback_item = control_items[2] if len(control_items) > 2 else None
        selectors_host = QWidget()
        selectors_grid = QGridLayout(selectors_host)
        selectors_grid.setContentsMargins(0, 0, 0, 0)
        selectors_grid.setHorizontalSpacing(7)
        selectors_grid.setVerticalSpacing(4)
        op_label = QLabel("Operador")
        posto_label = QLabel("Posto")
        oper_label = QLabel("Operacao")
        for label in (op_label, posto_label, oper_label):
            label.setStyleSheet("font-size: 9.4px; color: #334155;")
        selectors_grid.addWidget(op_label, 0, 0)
        selectors_grid.addWidget(posto_label, 0, 1)
        selectors_grid.addWidget(self.operator_combo, 1, 0)
        selectors_grid.addWidget(self.posto_combo, 1, 1)
        selectors_grid.addWidget(oper_label, 2, 0, 1, 2)
        selectors_grid.addWidget(self.operation_combo, 3, 0, 1, 2)
        scan_label = QLabel("Scanner")
        scan_label.setStyleSheet("font-size: 9.4px; color: #334155;")
        selectors_grid.addWidget(scan_label, 4, 0, 1, 2)
        selectors_grid.addWidget(self.scan_edit, 5, 0, 1, 2)
        selectors_grid.setColumnStretch(0, 1)
        selectors_grid.setColumnStretch(1, 1)
        control_layout.addWidget(selectors_host)
        _adopt_layout_item(control_layout, feedback_item)
        self.control_card.setMinimumHeight(190)
        self.control_card.setMaximumHeight(190)
        self.feedback_label.setStyleSheet("font-size: 9px; color: #475467;")
        self.feedback_label.setWordWrap(True)

        self.orders_table = QTableWidget(0, 9)
        self.orders_table.setHorizontalHeaderLabels(["Encomenda", "OF", "Cliente", "Estado", "Grupos", "Pecas", "Em curso", "Avarias", "Progress"])
        self.orders_table.verticalHeader().setVisible(False)
        self.orders_table.verticalHeader().setDefaultSectionSize(LIST_TABLE_ROW_PX)
        self.orders_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.orders_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.orders_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.orders_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.orders_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.orders_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        _configure_table(self.orders_table, stretch=(1,), contents=())
        orders_header = self.orders_table.horizontalHeader()
        for col, width in ((0, 174), (1, 132), (2, 300), (3, 116), (4, 66), (5, 64), (6, 82), (7, 78), (8, 84)):
            orders_header.setSectionResizeMode(col, QHeaderView.Interactive)
            orders_header.resizeSection(col, width)
        orders_header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.orders_table.setStyleSheet(
            f"QTableWidget {{ font-size: {max(10, LIST_TABLE_FONT_PX - 1)}px; }}"
            f" QHeaderView::section {{ font-size: {max(10, LIST_TABLE_FONT_PX - 1)}px; padding: 7px 8px; font-weight: 800; }}"
        )
        self.orders_table.itemDoubleClicked.connect(lambda item: self._open_selected_order_from_item(item))

        root = self.layout()
        sections = _take_layout_items(root)
        stats_item = sections[0] if len(sections) > 0 else None
        progress_item = sections[1] if len(sections) > 1 else None
        control_item = sections[2] if len(sections) > 2 else None
        context_item = sections[3] if len(sections) > 3 else None
        groups_item = sections[4] if len(sections) > 4 else None
        pieces_item = sections[5] if len(sections) > 5 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(8)
        _adopt_layout_item(list_layout, stats_item)
        _adopt_layout_item(list_layout, progress_item)
        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(14, 10, 14, 10)
        list_actions_layout.setSpacing(8)
        self.open_order_btn = QPushButton("Abrir encomenda")
        self.open_order_btn.setProperty("variant", "success")
        self.open_order_btn.clicked.connect(self._open_selected_order)
        refresh_group_btn = QPushButton("Atualizar")
        refresh_group_btn.setProperty("variant", "secondary")
        refresh_group_btn.clicked.connect(self.refresh)
        self.open_order_btn.setMinimumWidth(152)
        refresh_group_btn.setMinimumWidth(108)
        list_actions_layout.addWidget(self.open_order_btn)
        list_actions_layout.addWidget(refresh_group_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        list_filters = CardFrame()
        list_filters.set_tone("default")
        list_filters_layout = QHBoxLayout(list_filters)
        list_filters_layout.setContentsMargins(12, 8, 12, 8)
        list_filters_layout.setSpacing(8)
        list_filters_layout.addWidget(QLabel("Estado"))
        self.orders_state_filter_combo = QComboBox()
        self.orders_state_filter_combo.setProperty("compact", "true")
        self.orders_state_filter_combo.addItems(["Todas", "Em producao", "Concluida", "Em pausa", "Avaria", "Preparacao"])
        self.orders_state_filter_combo.currentTextChanged.connect(self._refresh_orders_list_view)
        list_filters_layout.addWidget(self.orders_state_filter_combo)
        list_filters_layout.addWidget(QLabel("Pesquisa"))
        self.orders_search_edit = QLineEdit()
        self.orders_search_edit.setProperty("compact", "true")
        self.orders_search_edit.setPlaceholderText("Filtrar encomenda ou cliente...")
        self.orders_search_edit.textChanged.connect(self._refresh_orders_list_view)
        list_filters_layout.addWidget(self.orders_search_edit, 1)
        _cap_width(self.orders_state_filter_combo, 152)
        list_layout.addWidget(list_filters)
        list_scan_card = CardFrame()
        list_scan_card.set_tone("info")
        list_scan_layout = QHBoxLayout(list_scan_card)
        list_scan_layout.setContentsMargins(12, 8, 12, 8)
        list_scan_layout.setSpacing(8)
        list_scan_title = QLabel("Scanner OF")
        list_scan_title.setStyleSheet("font-size: 12px; font-weight: 800; color: #0f172a;")
        list_scan_hint = QLabel("Pica a OF para abrir diretamente a encomenda certa.")
        list_scan_hint.setProperty("role", "muted")
        self.list_scan_edit = QLineEdit()
        self.list_scan_edit.setPlaceholderText("Picar OF")
        self.list_scan_edit.returnPressed.connect(lambda: self._handle_scan_code(self.list_scan_edit, "OF"))
        self.list_scan_edit.editingFinished.connect(lambda: self._handle_scan_code(self.list_scan_edit, "OF"))
        list_scan_layout.addWidget(list_scan_title)
        list_scan_layout.addWidget(list_scan_hint, 1)
        list_scan_layout.addWidget(self.list_scan_edit, 0)
        list_layout.addWidget(list_scan_card)
        self.orders_card = CardFrame()
        self.orders_card.set_tone("info")
        orders_layout = QVBoxLayout(self.orders_card)
        orders_layout.setContentsMargins(14, 12, 14, 12)
        orders_layout.setSpacing(8)
        orders_header = QHBoxLayout()
        orders_title = QLabel("Encomendas ativas")
        orders_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        orders_hint = QLabel("Entrar por encomenda e operar tudo dentro do detalhe.")
        orders_hint.setProperty("role", "muted")
        orders_header.addWidget(orders_title)
        orders_header.addStretch(1)
        orders_header.addWidget(orders_hint)
        orders_layout.addLayout(orders_header)
        orders_layout.addWidget(self.orders_table)
        list_layout.addWidget(self.orders_card, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(6)
        detail_actions = CardFrame()
        detail_actions.set_tone("default")
        detail_actions_layout = QHBoxLayout(detail_actions)
        detail_actions_layout.setContentsMargins(14, 8, 14, 8)
        detail_actions_layout.setSpacing(10)
        back_btn = QPushButton("Voltar a encomendas")
        back_btn.setProperty("variant", "secondary")
        back_btn.setMinimumWidth(166)
        back_btn.setToolTip("Regressar à lista de encomendas do Operador.")
        back_btn.clicked.connect(self._show_order_list)
        focus_block = QVBoxLayout()
        focus_block.setSpacing(2)
        self.order_focus_label = QLabel("Sem encomenda selecionada")
        self.order_focus_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.order_meta_label = QLabel("Seleciona uma encomenda para entrar no detalhe.")
        self.order_meta_label.setProperty("role", "muted")
        self.order_meta_label.setStyleSheet("font-size: 10.5px; color: #486581;")
        self.order_meta_label.setWordWrap(True)
        self.order_state_chip = QLabel("-")
        _apply_state_chip(self.order_state_chip, "-")
        self.order_state_chip.setMinimumWidth(132)
        self.order_state_chip.setAlignment(Qt.AlignCenter)
        focus_block.addWidget(self.order_focus_label)
        focus_block.addWidget(self.order_meta_label)
        detail_actions_layout.addWidget(back_btn)
        detail_actions_layout.addLayout(focus_block, 1)
        right_header_actions = QHBoxLayout()
        right_header_actions.setSpacing(6)
        for button, width in ((self.avaria_btn, 118), (self.close_avaria_btn, 108), (self.alert_btn, 112)):
            button.setMinimumWidth(width)
            right_header_actions.addWidget(button)
        detail_actions_layout.addLayout(right_header_actions)
        detail_actions_layout.addWidget(self.order_state_chip, 0, Qt.AlignTop)
        detail_actions.setMaximumHeight(72)
        detail_layout.addWidget(detail_actions)

        self.detail_filter_card = CardFrame()
        self.detail_filter_card.set_tone("info")
        detail_filter_layout = QHBoxLayout(self.detail_filter_card)
        detail_filter_layout.setContentsMargins(12, 8, 12, 8)
        detail_filter_layout.setSpacing(8)
        self.detail_active_chip = QLabel("-")
        self.detail_late_chip = QLabel("-")
        self.detail_running_chip = QLabel("-")
        self.detail_groups_chip = QLabel("-")
        self.detail_total_chip = QLabel("-")
        for chip in (
            self.detail_active_chip,
            self.detail_late_chip,
            self.detail_running_chip,
            self.detail_groups_chip,
            self.detail_total_chip,
        ):
            chip.setMinimumWidth(86)
            chip.setAlignment(Qt.AlignCenter)
            detail_filter_layout.addWidget(chip)
        self.detail_state_filter_combo = QComboBox()
        self.detail_state_filter_combo.setProperty("compact", "true")
        self.detail_state_filter_combo.addItems(["Todas", "Em producao", "Concluida", "Em pausa", "Avaria"])
        self.detail_state_filter_combo.currentTextChanged.connect(self._handle_group_selection)
        self.detail_search_edit = QLineEdit()
        self.detail_search_edit.setProperty("compact", "true")
        self.detail_search_edit.setPlaceholderText("Pesquisar ref...")
        self.detail_search_edit.textChanged.connect(self._handle_group_selection)
        for widget, width in ((self.detail_state_filter_combo, 152), (self.detail_search_edit, 360)):
            _cap_width(widget, width)
        detail_filter_layout.addStretch(1)
        detail_filter_layout.addWidget(self.detail_state_filter_combo)
        detail_filter_layout.addWidget(self.detail_search_edit)
        self.detail_filter_card.setMaximumHeight(50)

        control_host = QWidget()
        control_host.setMinimumWidth(558)
        control_host.setMaximumWidth(624)
        control_host_layout = QVBoxLayout(control_host)
        control_host_layout.setContentsMargins(0, 0, 0, 0)
        control_host_layout.setSpacing(4)
        _adopt_layout_item(control_host_layout, control_item)

        context_host = QWidget()
        context_host_layout = QVBoxLayout(context_host)
        context_host_layout.setContentsMargins(0, 0, 0, 0)
        context_host_layout.setSpacing(6)
        _adopt_layout_item(context_host_layout, context_item)

        workspace_split = QSplitter(Qt.Horizontal)
        workspace_split.setChildrenCollapsible(False)
        workspace_split.addWidget(control_host)
        workspace_split.addWidget(context_host)
        workspace_split.setSizes([580, 1530])
        workspace_split.setMaximumHeight(258)
        detail_layout.addWidget(workspace_split)

        self.groups_title_label.setText("Grupos ativos")
        self.groups_title_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        self.pieces_title_label.setText("Pecas do grupo")
        self.pieces_title_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
        for table, row_height, font_size in (
            (self.groups_table, 38, 10),
            (self.pieces_table, 40, 10),
        ):
            table.verticalHeader().setDefaultSectionSize(row_height)
            table.verticalHeader().setMinimumSectionSize(row_height)
            table.setStyleSheet(f"QTableWidget {{ font-size: {font_size}px; }} QTableWidget::item {{ padding: 6px 7px; }}")
        groups_visible_height = _table_visible_height(self.groups_table, 8, extra=18)
        self.groups_table.setMinimumHeight(groups_visible_height)
        self.groups_table.setMaximumHeight(16777215)
        self.pieces_table.setMinimumHeight(260)
        self.pieces_table.setSizeAdjustPolicy(QAbstractItemView.AdjustIgnored)
        self.pieces_card.setMinimumHeight(300)
        pieces_layout = self.pieces_card.layout()
        if pieces_layout is not None and pieces_layout.count() > 0:
            header_item = pieces_layout.itemAt(0)
            header_layout = header_item.layout() if header_item is not None else None
            if header_layout is not None:
                header_layout.setSpacing(6)
                _take_layout_items(header_layout)
                for button, width in (
                    (self.start_btn, 92),
                    (self.finish_btn, 92),
                    (self.resume_btn, 92),
                    (self.pause_btn, 102),
                    (self.manual_consume_btn, 96),
                    (self.consume_components_btn, 118),
                    (self.partial_consume_btn, 112),
                    (self.drawing_btn, 102),
                    (self.labels_btn, 98),
                    (self.local_refresh_btn, 96),
                ):
                    button.setMinimumWidth(width)
                self.pause_btn.setProperty("variant", "secondary")
                self.partial_consume_btn.setProperty("variant", "warning")
                self.local_refresh_btn.setProperty("variant", "secondary")
                header_layout.addWidget(self.pieces_title_label)
                header_layout.addStretch(1)
                header_layout.addWidget(self.start_btn)
                header_layout.addWidget(self.finish_btn)
                header_layout.addWidget(self.resume_btn)
                header_layout.addWidget(self.pause_btn)
                header_layout.addWidget(self.manual_consume_btn)
                header_layout.addWidget(self.consume_components_btn)
                header_layout.addWidget(self.partial_consume_btn)
                header_layout.addWidget(self.drawing_btn)
                header_layout.addWidget(self.labels_btn)
                header_layout.addWidget(self.local_refresh_btn)
                header_layout.addWidget(self.multi_count_chip)
                header_layout.addWidget(self.select_all_pieces_box)
        groups_widget = groups_item.widget()
        pieces_widget = pieces_item.widget() if pieces_item is not None and pieces_item.widget() is not None else None

        self.open_group_btn = QPushButton("Abrir espessura")
        self.open_group_btn.clicked.connect(self._open_selected_group_detail)
        self.open_group_btn.setMinimumWidth(146)
        self.group_overview_stack = QStackedWidget()

        overview_page = QWidget()
        overview_layout = QVBoxLayout(overview_page)
        overview_layout.setContentsMargins(0, 0, 0, 0)
        overview_layout.setSpacing(8)
        overview_actions = CardFrame()
        overview_actions.set_tone("info")
        overview_actions_layout = QHBoxLayout(overview_actions)
        overview_actions_layout.setContentsMargins(12, 8, 12, 8)
        overview_actions_layout.setSpacing(8)
        overview_title = QLabel("Espessuras da encomenda")
        overview_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        overview_hint = QLabel("Seleciona uma espessura para abrir as referências e o nesting associado.")
        overview_hint.setProperty("role", "muted")
        overview_hint.setWordWrap(True)
        overview_actions_layout.addWidget(overview_title)
        overview_actions_layout.addWidget(overview_hint, 1)
        self.thickness_scan_edit = QLineEdit()
        self.thickness_scan_edit.setPlaceholderText("Picar Espessura")
        self.thickness_scan_edit.setMinimumWidth(180)
        self.thickness_scan_edit.returnPressed.connect(lambda: self._handle_scan_code(self.thickness_scan_edit, "GRP"))
        self.thickness_scan_edit.editingFinished.connect(lambda: self._handle_scan_code(self.thickness_scan_edit, "GRP"))
        overview_actions_layout.addWidget(self.thickness_scan_edit)
        overview_actions_layout.addWidget(self.open_group_btn)
        overview_layout.addWidget(overview_actions)
        if groups_widget is not None:
            groups_widget.setMinimumHeight(groups_visible_height + 96)
            groups_widget.setMaximumHeight(16777215)
        _adopt_layout_item(overview_layout, groups_item, 1)

        detail_group_page = QWidget()
        detail_group_layout = QVBoxLayout(detail_group_page)
        detail_group_layout.setContentsMargins(0, 0, 0, 0)
        detail_group_layout.setSpacing(8)
        self.group_detail_header = CardFrame()
        self.group_detail_header.set_tone("info")
        group_detail_header_layout = QHBoxLayout(self.group_detail_header)
        group_detail_header_layout.setContentsMargins(14, 10, 14, 10)
        group_detail_header_layout.setSpacing(10)
        self.back_to_groups_btn = QPushButton("Voltar às espessuras")
        self.back_to_groups_btn.setProperty("variant", "success")
        self.back_to_groups_btn.setMinimumWidth(164)
        self.back_to_groups_btn.setToolTip("Regressar à seleção de material e espessura desta encomenda.")
        self.back_to_groups_btn.clicked.connect(self._show_group_overview)
        self.group_detail_label = QLabel("Sem espessura selecionada")
        self.group_detail_label.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.group_detail_meta = QLabel("Seleciona uma espessura para abrir o detalhe produtivo.")
        self.group_detail_meta.setProperty("role", "muted")
        self.group_detail_meta.setStyleSheet("font-size: 10.5px; color: #486581;")
        self.group_detail_meta.setWordWrap(True)
        group_detail_text = QVBoxLayout()
        group_detail_text.setSpacing(2)
        group_detail_text.addWidget(self.group_detail_label)
        group_detail_text.addWidget(self.group_detail_meta)
        group_detail_header_layout.addWidget(self.back_to_groups_btn)
        group_detail_header_layout.addLayout(group_detail_text, 1)
        self.opp_scan_edit = QLineEdit()
        self.opp_scan_edit.setPlaceholderText("Picar OPP")
        self.opp_scan_edit.setMinimumWidth(210)
        self.opp_scan_edit.returnPressed.connect(lambda: self._handle_scan_code(self.opp_scan_edit, "OPP"))
        self.opp_scan_edit.editingFinished.connect(lambda: self._handle_scan_code(self.opp_scan_edit, "OPP"))
        group_detail_header_layout.addWidget(self.opp_scan_edit)
        detail_group_layout.addWidget(self.group_detail_header)
        detail_group_layout.addWidget(self.detail_filter_card)

        self.nesting_info_card = None
        self.nesting_info_hint = None
        self.nesting_info_table = None
        if pieces_widget is not None:
            pieces_widget.setMinimumHeight(300)
            pieces_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        _adopt_layout_item(detail_group_layout, pieces_item, 1)

        self.group_overview_stack.addWidget(overview_page)
        self.group_overview_stack.addWidget(detail_group_page)
        detail_layout.addWidget(self.group_overview_stack, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.orders_table.itemSelectionChanged.connect(self._sync_order_focus)
        self.orders_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_order())
        self.groups_table.itemSelectionChanged.connect(self._sync_group_focus)
        self.groups_table.itemSelectionChanged.connect(self._sync_group_open_button)
        self.groups_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_group_detail())
        self._show_order_list()
        self._sync_order_focus()
        self._sync_group_open_button()

    def _render_groups_table(self, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        self.groups_table.setSortingEnabled(False)
        self.groups_table.blockSignals(True)
        self.groups_table.setRowCount(len(self.items))
        for row_index, item in enumerate(self.items):
            group_piece_stats = [_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")) for piece in list(item.get("pieces", []) or [])]
            group_progress = round(sum(float(stat.get("progress_pct", 0) or 0) for stat in group_piece_stats) / len(group_piece_stats), 1) if group_piece_stats else float(item.get("progress_pct", 0) or 0)
            row_values = [
                item.get("encomenda", "-"),
                item.get("cliente", "-"),
                item.get("estado_espessura", item.get("estado", "-")),
                item.get("material", "-"),
                item.get("espessura", "-"),
                f"{item.get('tempo_real_min', 0):.1f}",
                f"{group_progress:.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, "|".join(self._group_key(item)))
                if col_index >= 4:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.groups_table.setItem(row_index, col_index, cell)
            _paint_table_row(self.groups_table, row_index, str(item.get("estado_espessura", "")))
        self.groups_table.blockSignals(False)
        self.groups_table.setSortingEnabled(False)
        target_row = 0
        if previous_group:
            for index, item in enumerate(self.items):
                if self._group_key(item) == previous_group:
                    target_row = index
                    break
        if self.groups_table.rowCount() > 0:
            self.groups_table.selectRow(target_row)
            self.selected_group_key = self._group_key(self.items[target_row]) if target_row < len(self.items) else None
            self.selected_piece_id = previous_piece
            self._handle_group_selection()
        else:
            self.current_pieces = []
            self.pieces_table.setRowCount(0)
            self.pieces_title_label.setText("Pecas da encomenda")
            self._clear_piece_context()

    def refresh(self) -> None:
        show_detail = self.view_stack.currentWidget() is self.detail_page
        keep_group_detail = show_detail and hasattr(self, "group_overview_stack") and self.group_overview_stack.currentIndex() == 1
        keep_detail = show_detail and bool(self.selected_order_number)
        previous_order = self.selected_order_number or str(self._current_order_row().get("encomenda", "") or "").strip()
        previous_group = self.selected_group_key
        previous_piece = self.selected_piece_id
        user = self.backend.user or {}
        getter = getattr(self.backend, "ui_options", None)
        if callable(getter):
            try:
                self.ui_options = dict(getter() or {})
            except Exception:
                self.ui_options = {}
        self._set_combo_items(
            self.operator_combo,
            self.backend.operator_names(),
            preferred=str(user.get("username", "") or "").strip(),
        )
        self._sync_operator_assignment()
        data = self.runtime_service.operator_board(username=str(user.get("username", "")), role=str(user.get("role", "")))
        summary = data.get("summary", {})
        all_items = self._hydrate_operator_items(list(data.get("items", [])))
        all_items = self._append_montagem_stock_groups(all_items)
        op_total = 0
        op_done = 0
        op_running = 0
        for board_item in all_items:
            for piece in list(board_item.get("pieces", []) or []):
                stats = _piece_ops_progress(piece, str(piece.get("operacao_atual", "") or ""))
                op_total += int(stats.get("total", 0) or 0)
                op_done += int(stats.get("done", 0) or 0)
                op_running += int(stats.get("running", 0) or 0)
        global_ops_progress = 0.0 if op_total <= 0 else round(((op_done + (0.5 * op_running)) / op_total) * 100.0, 1)
        self.cards[0].set_data(summary.get("encomendas_ativas", 0), f"Grupos {summary.get('grupos', 0)}")
        self.cards[1].set_data(summary.get("pecas_em_curso", 0), f"Em pausa {summary.get('pecas_em_pausa', 0)}")
        self.cards[2].set_data(summary.get("pecas_em_avaria", 0), f"Concluidas {summary.get('pecas_concluidas', 0)}")
        self.cards[3].set_data(f"{global_ops_progress:.1f}%", f"Ops {op_done}/{op_total} | Em curso {op_running}")
        self.global_progress.setValue(int(round(global_ops_progress)))
        self.all_items = all_items
        self.order_rows_all = self._build_order_rows()
        self._render_orders_table(previous_order)
        if not self.order_rows_all:
            self.selected_order_number = ""
            self.items = []
            self.current_pieces = []
            self.groups_table.setRowCount(0)
            self.pieces_table.setRowCount(0)
            self._clear_piece_context()
            self._show_order_list()
            return
        if show_detail and previous_order:
            self._apply_order_filter(previous_order, previous_group=previous_group, previous_piece=previous_piece)
            if keep_group_detail and self.selected_group_key:
                self.view_stack.setCurrentWidget(self.detail_page)
                self._show_group_detail()
            else:
                self._show_order_detail()
        else:
            self.items = []
            self.current_pieces = []
            self.selected_group_key = None
            self.selected_piece_id = ""
            self._show_order_list()
        self._sync_order_focus()

    def _show_order_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_scan_placeholder()
        self._refresh_orders_list_view()
        self._sync_order_focus()

    def _show_order_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)
        self._sync_scan_placeholder()
        self._show_group_overview()

    def _show_group_overview(self) -> None:
        if hasattr(self, "group_overview_stack"):
            self.group_overview_stack.setCurrentIndex(0)
        self._sync_group_open_button()
        self._sync_scan_placeholder()

    def _show_group_detail(self) -> None:
        if hasattr(self, "group_overview_stack"):
            self.group_overview_stack.setCurrentIndex(1)
        self._refresh_group_support_panel()
        self._sync_group_open_button()
        self._sync_scan_placeholder()

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _current_order_row(self) -> dict:
        row_index = _selected_row_index(self.orders_table)
        if row_index < 0:
            return {}
        row_item = self.orders_table.item(row_index, 0)
        numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if numero:
            for row in self.order_rows:
                if str(row.get("encomenda", "") or "").strip() == numero:
                    return row
        if row_index >= len(self.order_rows):
            return {}
        return self.order_rows[row_index]

    def _build_order_rows(self) -> list[dict]:
        rows_map: dict[str, dict] = {}
        for item in self.all_items:
            numero = str(item.get("encomenda", "") or "").strip()
            if not numero:
                continue
            row = rows_map.setdefault(
                numero,
                {
                    "encomenda": numero,
                    "cliente": str(item.get("cliente_label", item.get("cliente", "")) or "-"),
                    "cliente_codigo": str(item.get("cliente_codigo", "") or "").strip(),
                    "cliente_nome": str(item.get("cliente_nome", "") or "").strip(),
                    "of_codigo": str(item.get("of", "") or item.get("of_codigo", "") or "").strip(),
                    "grupos": 0,
                    "pecas": 0,
                    "em_curso": 0,
                    "avarias": 0,
                    "progress_samples": [],
                    "states": [],
                },
            )
            row["grupos"] += 1
            pieces = list(item.get("pieces", []) or [])
            if not str(row.get("of_codigo", "") or "").strip():
                for piece in pieces:
                    of_txt = str(piece.get("of", "") or piece.get("of_codigo", "") or "").strip()
                    if of_txt:
                        row["of_codigo"] = of_txt
                        break
            row["states"].append(str(item.get("estado_espessura", item.get("estado", "")) or ""))
            row["progress_samples"].append(float(item.get("progress_pct", 0) or 0))
            row["pecas"] += len(pieces)
            for piece in pieces:
                state = str(piece.get("estado", "") or "").strip()
                row["states"].append(state)
                row["progress_samples"].append(float(_piece_ops_progress(piece, str(piece.get("operacao_atual", "") or "")).get("progress_pct", item.get("progress_pct", 0)) or 0))
                lowered = state.lower()
                if "produc" in lowered or "curso" in lowered:
                    row["em_curso"] += 1
                if "avaria" in lowered:
                    row["avarias"] += 1
        rows: list[dict] = []
        for row in rows_map.values():
            samples = [float(value or 0) for value in row.get("progress_samples", [])]
            row["progress_pct"] = round(sum(samples) / len(samples), 1) if samples else 0.0
            row["estado"] = self._aggregate_order_state(list(row.get("states", []) or []))
            rows.append(row)
        rows.sort(key=lambda item: str(item.get("encomenda", "") or ""))
        return rows

    def _filtered_order_rows(self) -> list[dict]:
        rows = list(self.order_rows_all)
        state_filter = self.orders_state_filter_combo.currentText().strip().lower() if hasattr(self, "orders_state_filter_combo") else "todas"
        search = self.orders_search_edit.text().strip().lower() if hasattr(self, "orders_search_edit") else ""
        if state_filter and state_filter != "todas":
            filtered_by_state: list[dict] = []
            for row in rows:
                state = str(row.get("estado", "") or "").strip().lower()
                if state_filter == "em producao" and not ("produc" in state or "curso" in state):
                    continue
                if state_filter == "concluida" and "concl" not in state:
                    continue
                if state_filter == "em pausa" and not ("paus" in state or "interromp" in state):
                    continue
                if state_filter == "avaria" and "avaria" not in state:
                    continue
                if state_filter == "preparacao" and not ("prepar" in state or "pend" in state or "edicao" in state):
                    continue
                filtered_by_state.append(row)
            rows = filtered_by_state
        if search:
            rows = [row for row in rows if search in str(row.get("encomenda", "") or "").lower() or search in str(row.get("cliente", "") or "").lower() or search in str(row.get("estado", "") or "").lower()]
        return rows

    def _aggregate_order_state(self, states: list[str]) -> str:
        lowered = [str(state or "").strip().lower() for state in states if str(state or "").strip()]
        if any("avaria" in state for state in lowered):
            return "Avaria"
        if any("produc" in state or "curso" in state for state in lowered):
            return "Em producao"
        if any("paus" in state or "interromp" in state for state in lowered):
            return "Em pausa"
        if lowered and all("concl" in state for state in lowered):
            return "Concluida"
        if any("prepar" in state or "pend" in state or "edicao" in state for state in lowered):
            return "Preparacao"
        return states[0] if states else "-"

    def _render_orders_table(self, selected_order: str = "") -> None:
        self.order_rows = self._filtered_order_rows()
        self.orders_table.setSortingEnabled(False)
        self.orders_table.blockSignals(True)
        self.orders_table.setRowCount(len(self.order_rows))
        for row_index, row in enumerate(self.order_rows):
            row_values = [
                row.get("encomenda", "-"),
                row.get("of_codigo", "-") or "-",
                _format_client_label(row.get("cliente", "-"), show_name=True),
                row.get("estado", "-"),
                row.get("grupos", 0),
                row.get("pecas", 0),
                row.get("em_curso", 0),
                row.get("avarias", 0),
                f"{float(row.get('progress_pct', 0) or 0):.1f}%",
            ]
            for col_index, value in enumerate(row_values):
                cell = QTableWidgetItem(str(value))
                cell.setToolTip(str(value))
                if col_index == 0:
                    cell.setData(Qt.UserRole, str(row.get("encomenda", "") or "").strip())
                    font = cell.font()
                    font.setBold(True)
                    cell.setFont(font)
                if col_index >= 4:
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.orders_table.setItem(row_index, col_index, cell)
        self.orders_table.blockSignals(False)
        self.orders_table.setSortingEnabled(False)
        for row_index, row in enumerate(self.order_rows):
            _paint_table_row(self.orders_table, row_index, str(row.get("estado", "")))
        if self.orders_table.rowCount() == 0:
            self.open_order_btn.setEnabled(False)
            self._set_order_header({})
            return
        target_row = 0
        if selected_order:
            for row_index, row in enumerate(self.order_rows):
                if str(row.get("encomenda", "") or "").strip() == selected_order:
                    target_row = row_index
                    break
        self.orders_table.selectRow(target_row)

    def _refresh_orders_list_view(self) -> None:
        selected_order = str(self._current_order_row().get("encomenda", "") or "").strip()
        self._render_orders_table(selected_order)
        self._sync_order_focus()

    def _refresh_detail_badges(self) -> None:
        pieces = [piece for item in self.items for piece in list(item.get("pieces", []) or [])]
        active_count = 0
        running_count = 0
        delayed_count = 0
        for item in self.items:
            if float(item.get("desvio_min", 0) or 0) > 0:
                delayed_count += 1
        for piece in pieces:
            state = str(piece.get("estado", "") or "").strip().lower()
            if "concl" not in state:
                active_count += 1
            if "produc" in state or "curso" in state:
                running_count += 1
        _apply_state_chip(self.detail_active_chip, "Preparacao", f"Ativas {active_count}")
        _apply_state_chip(self.detail_late_chip, "Em pausa", f"Atrasadas {delayed_count}")
        _apply_state_chip(self.detail_running_chip, "Em producao", f"Em curso {running_count}")
        _apply_state_chip(self.detail_groups_chip, "Concluida", f"{len(self.items)} grupos")
        _apply_state_chip(self.detail_total_chip, "-", f"Todas {len(pieces)}")

    def _set_order_header(self, row: dict) -> None:
        numero = str(row.get("encomenda", "") or "").strip()
        estado = str(row.get("estado", "-") or "-").strip() or "-"
        if not numero:
            self.order_focus_label.setText("Sem encomenda selecionada")
            self.order_meta_label.setText("Seleciona uma encomenda para entrar no detalhe.")
            _apply_state_chip(self.order_state_chip, "-")
            if hasattr(self, "detail_active_chip"):
                _apply_state_chip(self.detail_active_chip, "-", "Ativas 0")
                _apply_state_chip(self.detail_late_chip, "-", "Atrasadas 0")
                _apply_state_chip(self.detail_running_chip, "-", "Em curso 0")
                _apply_state_chip(self.detail_groups_chip, "-", "0 grupos")
                _apply_state_chip(self.detail_total_chip, "-", "Todas 0")
            return
        of_txt = str(row.get("of_codigo", "") or "-").strip() or "-"
        full_order_title = f"{numero} | OF {of_txt} | {_format_client_label(row.get('cliente', '-'), show_name=True)}"
        self.order_focus_label.setText(_elide_middle(full_order_title, 72))
        self.order_focus_label.setToolTip(full_order_title)
        self.order_meta_label.setText(f"{int(row.get('grupos', 0) or 0)} grupos | {int(row.get('pecas', 0) or 0)} pecas | Em curso {int(row.get('em_curso', 0) or 0)} | Avarias {int(row.get('avarias', 0) or 0)} | Progresso {float(row.get('progress_pct', 0) or 0):.1f}%")
        _apply_state_chip(self.order_state_chip, estado)

    def _sync_order_focus(self) -> None:
        row = self._current_order_row()
        self.open_order_btn.setEnabled(bool(row))
        self._set_order_header(row)

    def _sync_group_open_button(self) -> None:
        has_group = bool(self._current_group())
        if hasattr(self, "open_group_btn"):
            self.open_group_btn.setEnabled(has_group)

    def _apply_order_filter(self, numero: str, previous_group: tuple[str, str, str] | None = None, previous_piece: str = "") -> None:
        numero_txt = str(numero or "").strip()
        self.selected_order_number = numero_txt
        self.items = [item for item in self.all_items if str(item.get("encomenda", "") or "").strip() == numero_txt]
        self._refresh_detail_badges()
        target_group = previous_group if previous_group and previous_group[0] == numero_txt else None
        self._render_groups_table(target_group, previous_piece if target_group else "")
        order_row = next((row for row in self.order_rows_all if str(row.get("encomenda", "") or "").strip() == numero_txt), {})
        self._set_order_header(order_row)
        if not self.items:
            self._clear_piece_context()
            self._refresh_group_support_panel()

    def _sync_group_focus(self) -> None:
        group = self._current_group()
        row = next((item for item in self.order_rows_all if str(item.get("encomenda", "") or "").strip() == str(self.selected_order_number or "").strip()), {})
        if not row:
            return
        summary = (
                f"{int(row.get('grupos', 0) or 0)} grupos | {int(row.get('pecas', 0) or 0)} pecas | "
                f"Em curso {int(row.get('em_curso', 0) or 0)} | Avarias {int(row.get('avarias', 0) or 0)} | "
                f"Progresso {float(row.get('progress_pct', 0) or 0):.1f}%"
            )
        if group:
            summary = (
                f"{summary} | Cliente {_format_client_label(group.get('cliente', '-'), show_name=True)} | "
                f"Grupo {group.get('material', '-')} {group.get('espessura', '-')} mm | "
                f"Estado {group.get('estado_espessura', group.get('estado', '-'))}"
            )
        self.order_meta_label.setText(summary)
        if hasattr(self, "group_detail_label"):
            if group:
                self.group_detail_label.setText(
                    f"{group.get('material', '-')} {group.get('espessura', '-')} mm | {_format_client_label(group.get('cliente', '-'), show_name=True)}"
                )
                self.group_detail_meta.setText(
                    f"{int(len(list(group.get('pieces', []) or [])))} referência(s) | "
                    f"Estado {group.get('estado_espessura', group.get('estado', '-'))} | "
                    f"Tempo planeado {float(group.get('tempo_plan_min', 0) or 0):.1f} min | "
                    f"Tempo real {float(group.get('tempo_real_min', 0) or 0):.1f} min"
                )
            else:
                self.group_detail_label.setText("Sem espessura selecionada")
                self.group_detail_meta.setText("Seleciona uma espessura para abrir o detalhe produtivo.")
        self._refresh_group_support_panel()

    def _handle_group_selection(self) -> None:
        OperatorPage._handle_group_selection(self)
        self._sync_group_focus()

    def _refresh_group_support_panel(self) -> None:
        if not getattr(self, "nesting_info_table", None):
            return
        group = self._current_group()
        rows: list[tuple[str, str, str]] = []
        if group:
            order_num = str(group.get("encomenda", "") or "").strip()
            material_txt = str(group.get("material", "") or "").strip()
            thickness_txt = str(group.get("espessura", "") or "").strip()
            chapa_txt = str(group.get("chapa", "") or "").strip()
            if chapa_txt and chapa_txt != "-":
                rows.append(("Chapa", chapa_txt, "Chapa / lote atualmente associado à espessura."))
            drawings_seen: set[str] = set()
            refs_seen: set[str] = set()
            op_seen: set[str] = set()
            quote_snapshots: list[dict[str, Any]] = []
            for piece in list(group.get("pieces", []) or []):
                ref_txt = str(piece.get("ref_externa", "") or piece.get("ref_interna", "") or "").strip()
                if ref_txt and ref_txt not in refs_seen:
                    refs_seen.add(ref_txt)
                    rows.append(("Referência", ref_txt, str(piece.get("descricao", "") or "-").strip() or "-"))
                drawing_txt = str(piece.get("desenho", "") or "").strip()
                if drawing_txt and drawing_txt not in drawings_seen:
                    drawings_seen.add(drawing_txt)
                    rows.append(("Desenho", Path(drawing_txt).name, drawing_txt))
                snapshot = dict(piece.get("quote_cost_snapshot", {}) or {})
                if snapshot:
                    quote_snapshots.append(snapshot)
                for op_name in list(piece.get("pendentes", []) or []):
                    op_txt = str(op_name or "").strip()
                    if op_txt and op_txt not in op_seen:
                        op_seen.add(op_txt)
            if op_seen:
                rows.append(("Fluxo", "Operações", ", ".join(sorted(op_seen))))
            if quote_snapshots:
                quote_number = next((str(row.get("quote_number", "") or "").strip() for row in quote_snapshots if str(row.get("quote_number", "") or "").strip()), "")
                quote_state = next((str(row.get("quote_state", "") or "").strip() for row in quote_snapshots if str(row.get("quote_state", "") or "").strip()), "")
                avg_price = 0.0
                avg_time = 0.0
                valid_prices = [float(row.get("preco_unit_total_eur", 0) or 0) for row in quote_snapshots]
                valid_times = [float(row.get("tempo_total_peca_min", 0) or 0) for row in quote_snapshots]
                if valid_prices:
                    avg_price = sum(valid_prices) / max(1, len(valid_prices))
                if valid_times:
                    avg_time = sum(valid_times) / max(1, len(valid_times))
                if quote_number:
                    rows.append(("Orçamento", quote_number, quote_state or "Origem comercial da espessura."))
                if avg_price > 0 or avg_time > 0:
                    rows.append(
                        (
                            "Estimativa",
                            f"{avg_price:.2f} EUR/pc",
                            f"Tempo médio {avg_time:.2f} min/pc | snapshots {len(quote_snapshots)}",
                        )
                    )
            quote_number = ""
            order_detail_getter = getattr(self.backend, "order_detail", None)
            if callable(order_detail_getter) and order_num:
                try:
                    order_detail = dict(order_detail_getter(order_num) or {})
                except Exception:
                    order_detail = {}
                quote_number = str(order_detail.get("numero_orcamento", "") or "").strip()
                if quote_number and not any(row[0] == "Orçamento" for row in rows):
                    rows.append(("Orçamento", quote_number, "Ligação da encomenda ao orçamento de origem."))
            studies_getter = getattr(self.backend, "orc_nesting_studies", None)
            if callable(studies_getter) and quote_number:
                try:
                    studies = dict(studies_getter(quote_number) or {})
                except Exception:
                    studies = {}
                matched_studies: list[dict[str, Any]] = []
                refs_norm = {str(value or "").strip().lower() for value in refs_seen if str(value or "").strip()}
                material_norm = str(material_txt or "").strip().lower()
                thickness_norm = str(thickness_txt or "").strip().lower()
                for study_key, study_value in studies.items():
                    study = dict(study_value or {})
                    result_data = dict(study.get("result_data", {}) or {})
                    summary = dict(study.get("summary", result_data.get("summary", {})) or {})
                    bridge = dict(study.get("quote_bridge", {}) or {})
                    group_label = str(study.get("group_label", "") or study_key).strip()
                    label_norm = group_label.lower()
                    bridge_refs = {
                        str(row.get("ref_externa", "") or "").strip().lower()
                        for row in list(bridge.get("part_rows", []) or [])
                        if isinstance(row, dict) and str(row.get("ref_externa", "") or "").strip()
                    }
                    material_match = bool(material_norm) and material_norm in label_norm
                    thickness_match = bool(thickness_norm) and thickness_norm in label_norm
                    refs_match = bool(refs_norm and bridge_refs and refs_norm.intersection(bridge_refs))
                    if refs_match or (material_match and thickness_match):
                        matched_studies.append(
                            {
                                "key": str(study_key or "").strip(),
                                "label": group_label or "-",
                                "updated_at": str(study.get("updated_at", "") or "").strip(),
                                "summary": summary,
                            }
                        )
                if matched_studies:
                    matched_studies.sort(key=lambda row: str(row.get("updated_at", "") or ""), reverse=True)
                    best = matched_studies[0]
                    best_summary = dict(best.get("summary", {}) or {})
                    rows.append(
                        (
                            "Nesting",
                            best.get("label", "-"),
                            f"{float(best_summary.get('utilization_net_pct', 0) or 0):.1f}% real | "
                            f"{int(best_summary.get('sheet_count', 0) or 0)} chapa(s) | "
                            f"{int(best_summary.get('part_count_placed', 0) or 0)}/{int(best_summary.get('part_count_requested', 0) or 0)} colocadas",
                        )
                    )
                    if len(matched_studies) > 1:
                        rows.append(("Estudos", str(len(matched_studies)), "Foram encontrados vários estudos compatíveis com esta espessura."))
        self.nesting_info_table.setRowCount(len(rows))
        for row_index, values in enumerate(rows):
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if col_index == 0:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                self.nesting_info_table.setItem(row_index, col_index, item)
        if hasattr(self, "nesting_info_hint"):
            self.nesting_info_hint.setText(
                "Sem informação técnica adicional para esta espessura."
                if not rows
                else f"{len(rows)} registo(s) técnicos ligados à espessura selecionada."
            )

    def _open_selected_group_detail(self) -> None:
        if not self._current_group():
            QMessageBox.warning(self, "Operador", "Seleciona uma espessura.")
            return
        self._show_group_detail()

    def _open_selected_order(self) -> None:
        row = self._current_order_row()
        numero = str(row.get("encomenda", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Operador", "Seleciona uma encomenda.")
            return
        self._apply_order_filter(numero)
        self._show_order_detail()

    def _open_selected_order_from_item(self, item: QTableWidgetItem) -> None:
        row_item = self.orders_table.item(item.row(), 0) or item
        numero = str(row_item.data(Qt.UserRole) or row_item.text() or "").strip()
        if not numero:
            return
        self._apply_order_filter(numero)
        self._show_order_detail()
