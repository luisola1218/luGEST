from __future__ import annotations
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
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
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from datetime import datetime
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    configure_table as _configure_table,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
    set_panel_tone as _set_panel_tone,
    set_table_columns as _set_table_columns,
    state_tone as _state_tone,
)
from .runtime_support import _adopt_layout_item, _take_layout_items
from ..widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


class ExpeditionPage(QWidget):
    page_title = "Expedição"
    page_subtitle = "Expedição operacional com peças disponíveis, guia em preparação e histórico de ações."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.pending_rows: list[dict] = []
        self.piece_rows: list[dict] = []
        self.history_rows: list[dict] = []
        self.current_guide_lines: list[dict] = []
        self.draft_rows: list[dict] = []
        self.current_pending_num = ""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        actions = CardFrame()
        actions.set_tone("info")
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(14, 10, 14, 10)
        actions_layout.setSpacing(8)
        self.emit_off_btn = QPushButton("Emitir Guia OFF")
        self.emit_off_btn.clicked.connect(self._emit_off_guide)
        self.manual_guide_btn = QPushButton("Criar Guia Manual")
        self.manual_guide_btn.setProperty("variant", "secondary")
        self.manual_guide_btn.clicked.connect(self._create_manual_guide)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.edit_guide_btn = QPushButton("Editar Guia")
        self.edit_guide_btn.setProperty("variant", "secondary")
        self.edit_guide_btn.clicked.connect(self._edit_guide)
        self.cancel_guide_btn = QPushButton("Anular Guia")
        self.cancel_guide_btn.setProperty("variant", "danger")
        self.cancel_guide_btn.clicked.connect(self._cancel_guide)
        self.save_pdf_btn = QPushButton("Guardar PDF")
        self.save_pdf_btn.setProperty("variant", "secondary")
        self.save_pdf_btn.clicked.connect(self._save_guide_pdf)
        self.history_dialog_btn = QPushButton("Histórico Guias")
        self.history_dialog_btn.setProperty("variant", "secondary")
        self.history_dialog_btn.clicked.connect(self._show_history_dialog)
        for button in (self.emit_off_btn, self.manual_guide_btn, self.refresh_btn, self.edit_guide_btn, self.cancel_guide_btn, self.preview_pdf_btn if hasattr(self, "preview_pdf_btn") else None, self.save_pdf_btn, self.history_dialog_btn):
            if button is not None:
                actions_layout.addWidget(button)
        actions_layout.addStretch(1)
        root.addWidget(actions)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(10)
        self.filter_edit = QComboBox()
        self.filter_edit.setEditable(True)
        self.filter_edit.setInsertPolicy(QComboBox.NoInsert)
        self.filter_edit.lineEdit().setPlaceholderText("Filtrar por encomenda, cliente, guia ou matrícula")
        self.filter_edit.lineEdit().textChanged.connect(self.refresh)
        self.estado_combo = QComboBox()
        self.estado_combo.addItems(["Todas", "Não expedida", "Parcialmente expedida", "Totalmente expedida"])
        self.estado_combo.currentTextChanged.connect(self.refresh)
        filters_layout.addWidget(QLabel("Pesquisa"))
        filters_layout.addWidget(self.filter_edit, 1)
        filters_layout.addWidget(QLabel("Estado"))
        filters_layout.addWidget(self.estado_combo)
        root.addWidget(filters)

        pending_card = CardFrame()
        pending_card.set_tone("warning")
        pending_layout = QVBoxLayout(pending_card)
        pending_layout.setContentsMargins(14, 12, 14, 12)
        pending_title = QLabel("Encomendas com material para expedir")
        pending_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.pending_table = QTableWidget(0, 6)
        self.pending_table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Estado Prod.", "Estado Exp.", "Disponível", "Entrega"])
        self.pending_table.verticalHeader().setVisible(False)
        self.pending_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pending_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.pending_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pending_table.verticalHeader().setDefaultSectionSize(24)
        self.pending_table.itemSelectionChanged.connect(self._on_pending_selected)
        pending_layout.addWidget(pending_title)
        pending_layout.addWidget(self.pending_table)
        self.pending_empty_label = QLabel("Sem quantidades prontas para expedição neste momento.")
        self.pending_empty_label.setProperty("role", "muted")
        self.pending_empty_label.setVisible(False)
        pending_layout.addWidget(self.pending_empty_label)
        root.addWidget(pending_card)

        pieces_card = CardFrame()
        pieces_card.set_tone("success")
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(14, 12, 14, 12)
        pieces_header = QHBoxLayout()
        pieces_title = QLabel("Peças disponíveis para expedição")
        pieces_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.open_piece_drawing_btn = QPushButton("Ver desenho da peça")
        self.open_piece_drawing_btn.setProperty("variant", "secondary")
        self.open_piece_drawing_btn.clicked.connect(self._open_piece_drawing)
        pieces_header.addWidget(pieces_title, 1)
        pieces_header.addWidget(self.open_piece_drawing_btn)
        self.pieces_table = QTableWidget(0, 8)
        self.pieces_table.setHorizontalHeaderLabels(["ID", "Ref. Int.", "Ref. Ext.", "Estado", "Pronta", "Expedida", "Disponível", "Desenho"])
        self.pieces_table.verticalHeader().setVisible(False)
        self.pieces_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pieces_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.pieces_table, stretch=(2,), contents=())
        _set_table_columns(
            self.pieces_table,
            [
                (0, "interactive", 120),
                (1, "interactive", 190),
                (2, "stretch", 320),
                (3, "interactive", 136),
                (4, "interactive", 82),
                (5, "interactive", 88),
                (6, "interactive", 96),
                (7, "interactive", 82),
            ],
        )
        self.pieces_table.verticalHeader().setDefaultSectionSize(24)
        self.pieces_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        pieces_layout.addLayout(pieces_header)
        pieces_layout.addWidget(self.pieces_table)
        root.addWidget(pieces_card)

        draft_card = CardFrame()
        draft_card.set_tone("info")
        draft_layout = QVBoxLayout(draft_card)
        draft_layout.setContentsMargins(14, 12, 14, 12)
        draft_layout.setSpacing(8)
        draft_header = QHBoxLayout()
        draft_title = QLabel("Guia em preparação")
        draft_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.draft_context = QLabel("Seleciona uma encomenda e uma peça para começar a guia.")
        self.draft_context.setProperty("role", "muted")
        draft_header.addWidget(draft_title)
        draft_header.addStretch(1)
        draft_header.addWidget(self.draft_context)
        draft_layout.addLayout(draft_header)
        draft_actions = QHBoxLayout()
        self.exp_qty_spin = QDoubleSpinBox()
        self.exp_qty_spin.setRange(0.01, 1000000.0)
        self.exp_qty_spin.setDecimals(2)
        self.exp_qty_spin.setValue(1.0)
        self.add_draft_btn = QPushButton("Adicionar linha")
        self.add_draft_btn.clicked.connect(self._add_draft_line)
        self.remove_draft_btn = QPushButton("Remover linha")
        self.remove_draft_btn.setProperty("variant", "secondary")
        self.remove_draft_btn.clicked.connect(self._remove_draft_line)
        self.clear_draft_btn = QPushButton("Limpar guia")
        self.clear_draft_btn.setProperty("variant", "secondary")
        self.clear_draft_btn.clicked.connect(self._clear_draft)
        draft_actions.addWidget(QLabel("Qtd"))
        draft_actions.addWidget(self.exp_qty_spin)
        draft_actions.addWidget(self.add_draft_btn)
        draft_actions.addWidget(self.remove_draft_btn)
        draft_actions.addWidget(self.clear_draft_btn)
        draft_actions.addStretch(1)
        draft_layout.addLayout(draft_actions)
        self.draft_table = QTableWidget(0, 5)
        self.draft_table.setHorizontalHeaderLabels(["Peça", "Ref. Int.", "Ref. Ext.", "Descrição", "Qtd"])
        self.draft_table.verticalHeader().setVisible(False)
        self.draft_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.draft_table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(self.draft_table, stretch=(2, 3), contents=())
        _set_table_columns(
            self.draft_table,
            [
                (0, "interactive", 96),
                (1, "interactive", 160),
                (2, "stretch", 210),
                (3, "stretch", 240),
                (4, "interactive", 82),
            ],
        )
        self.draft_table.verticalHeader().setDefaultSectionSize(24)
        self.draft_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        draft_layout.addWidget(self.draft_table)
        root.addWidget(draft_card)

        history_card = CardFrame()
        history_card.set_tone("default")
        history_layout = QVBoxLayout(history_card)
        history_layout.setContentsMargins(14, 12, 14, 12)
        history_header = QHBoxLayout()
        history_title = QLabel("Histórico de guias")
        history_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.preview_pdf_btn = QPushButton("Abrir PDF da guia")
        self.preview_pdf_btn.clicked.connect(self._open_guide_pdf)
        history_header.addWidget(history_title, 1)
        history_header.addWidget(self.preview_pdf_btn)
        self.history_table = QTableWidget(0, 7)
        self.history_table.setHorizontalHeaderLabels(["Guia", "Tipo", "Encomenda", "Cliente", "Emissão", "Estado", "Linhas"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table.verticalHeader().setDefaultSectionSize(24)
        self.history_table.itemSelectionChanged.connect(self._update_guide_detail)
        self.history_table.itemSelectionChanged.connect(self._sync_expedicao_actions)
        history_layout.addLayout(history_header)
        history_layout.addWidget(self.history_table)
        self.history_empty_label = QLabel("Sem guias emitidas neste momento.")
        self.history_empty_label.setProperty("role", "muted")
        self.history_empty_label.setVisible(False)
        history_layout.addWidget(self.history_empty_label)
        root.addWidget(history_card)

        self.detail_card = CardFrame()
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(14, 12, 14, 12)
        detail_layout.setSpacing(8)
        detail_header = QHBoxLayout()
        self.detail_title = QLabel("Seleciona uma guia")
        self.detail_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.detail_state_chip = QLabel("-")
        _apply_state_chip(self.detail_state_chip, "-")
        detail_header.addWidget(self.detail_title, 1)
        detail_header.addWidget(self.detail_state_chip, 0, Qt.AlignRight)
        detail_layout.addLayout(detail_header)
        self.detail_meta = QLabel("-")
        self.detail_meta.setProperty("role", "muted")
        self.detail_meta.setWordWrap(True)
        detail_layout.addWidget(self.detail_meta)
        self.detail_note = QLabel("-")
        self.detail_note.setWordWrap(True)
        detail_layout.addWidget(self.detail_note)
        self.lines_table = QTableWidget(0, 7)
        self.lines_table.setHorizontalHeaderLabels(["Ref. Int.", "Ref. Ext.", "Descrição", "Qtd", "Peso", "Manual", "Encomenda"])
        self.lines_table.verticalHeader().setVisible(False)
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.lines_table, stretch=(1, 2, 6), contents=())
        _set_table_columns(
            self.lines_table,
            [
                (0, "interactive", 170),
                (1, "stretch", 250),
                (2, "stretch", 280),
                (3, "interactive", 82),
                (4, "interactive", 82),
                (5, "interactive", 82),
                (6, "stretch", 170),
            ],
        )
        self.lines_table.verticalHeader().setDefaultSectionSize(24)
        detail_layout.addWidget(self.lines_table)
        self.detail_card.set_tone("default")
        root.addWidget(self.detail_card, 1)
        self._render_draft_table()
        self._sync_expedicao_actions()

    def refresh(self) -> None:
        query = self.filter_edit.currentText().strip()
        previous_pending = self.current_pending_num or self._current_pending_row().get("numero", "")
        previous_guide = self._current_history_row().get("numero", "")
        self.pending_rows = self.backend.expedicao_pending_orders(query, self.estado_combo.currentText())
        self.history_rows = self.backend.expedicao_rows(query)
        self._refresh_filter_options(query)
        _fill_table(
            self.pending_table,
            [[r.get("numero", "-"), r.get("cliente", "-"), r.get("estado", "-"), r.get("estado_expedicao", "-"), f"{r.get('disponivel', 0):.1f}", r.get("data_entrega", "-")] for r in self.pending_rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.pending_rows):
            _paint_table_row(self.pending_table, row_index, str(row.get("estado_expedicao", "")))
        self.pending_empty_label.setVisible(self.pending_table.rowCount() == 0)
        self.pending_empty_label.setText(
            "Sem quantidades prontas para expedição neste momento."
            if not query
            else "Nenhuma encomenda disponível para expedição com o filtro atual."
        )
        _fill_table(
            self.history_table,
            [[r.get("numero", "-"), r.get("tipo", "-"), r.get("encomenda", "-"), r.get("cliente", "-"), r.get("data_emissao", "-"), r.get("estado", "-"), r.get("linhas", 0)] for r in self.history_rows],
            align_center_from=6,
        )
        self.history_empty_label.setVisible(self.history_table.rowCount() == 0)
        self.history_empty_label.setText(
            "Sem guias emitidas neste momento."
            if not query
            else "Nenhuma guia encontrada com o filtro atual."
        )
        for row_index, row in enumerate(self.history_rows):
            _paint_table_row(self.history_table, row_index, str(row.get("estado", "")))
        self._restore_selection(self.pending_table, self.pending_rows, previous_pending)
        self._restore_selection(self.history_table, self.history_rows, previous_guide)
        if self.pending_table.rowCount() == 0:
            self.current_pending_num = ""
            self.piece_rows = []
            self.pieces_table.setRowCount(0)
            self.draft_context.setText("Seleciona uma encomenda e uma peça para começar a guia.")
        else:
            self._on_pending_selected()
        if self.history_table.rowCount() == 0:
            self._clear_guide_detail()
        else:
            self._update_guide_detail()
        self._sync_expedicao_actions()

    def _refresh_filter_options(self, current_text: str) -> None:
        suggestions: list[str] = []
        for row in list(self.pending_rows) + list(self.history_rows):
            for key in ("numero", "cliente", "encomenda", "matricula"):
                value = str(row.get(key, "") or "").strip()
                if value and value not in suggestions:
                    suggestions.append(value)
        block = self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.addItem("")
        for value in suggestions:
            self.filter_edit.addItem(value)
        self.filter_edit.setCurrentText(current_text)
        self.filter_edit.blockSignals(block)

    def _restore_selection(self, table: QTableWidget, rows: list[dict], target_num: str) -> None:
        if table.rowCount() == 0:
            return
        row_index = 0
        if target_num:
            for index, row in enumerate(rows):
                if str(row.get("numero", "")).strip() == str(target_num).strip():
                    row_index = index
                    break
        table.selectRow(row_index)

    def _current_pending_row(self) -> dict:
        current = self.pending_table.currentItem()
        if current is None or current.row() >= len(self.pending_rows):
            return {}
        return self.pending_rows[current.row()]

    def _current_piece_row(self) -> dict:
        current = self.pieces_table.currentItem()
        if current is None or current.row() >= len(self.piece_rows):
            return {}
        return self.piece_rows[current.row()]

    def _current_history_row(self) -> dict:
        current = self.history_table.currentItem()
        if current is None or current.row() >= len(self.history_rows):
            return {}
        return self.history_rows[current.row()]

    def _current_draft_row(self) -> dict:
        current = self.draft_table.currentItem()
        if current is None or current.row() >= len(self.draft_rows):
            return {}
        return self.draft_rows[current.row()]

    def _sync_expedicao_actions(self) -> None:
        has_pending = bool(self._current_pending_row())
        has_piece = bool(self._current_piece_row())
        has_history = bool(self._current_history_row())
        history_state = str(self._current_history_row().get("estado", "") or "").strip().lower()
        has_draft = bool(self.draft_rows)
        self.emit_off_btn.setEnabled(has_pending and has_draft)
        self.manual_guide_btn.setEnabled(True)
        self.edit_guide_btn.setEnabled(has_history and "anulad" not in history_state)
        self.cancel_guide_btn.setEnabled(has_history and "anulad" not in history_state)
        self.preview_pdf_btn.setEnabled(has_history)
        self.save_pdf_btn.setEnabled(has_history)
        self.open_piece_drawing_btn.setEnabled(has_pending and has_piece)
        self.add_draft_btn.setEnabled(has_pending and has_piece)
        self.remove_draft_btn.setEnabled(bool(self._current_draft_row()))
        self.clear_draft_btn.setEnabled(has_draft)

    def _on_pending_selected(self) -> None:
        current = self._current_pending_row()
        enc_num = str(current.get("numero", "") or "").strip()
        if not enc_num:
            self.current_pending_num = ""
            self.piece_rows = []
            self.pieces_table.setRowCount(0)
            self.draft_context.setText("Seleciona uma encomenda e uma peça para começar a guia.")
            self._sync_expedicao_actions()
            return
        if self.current_pending_num and self.current_pending_num != enc_num and self.draft_rows:
            self.draft_rows = []
            self._render_draft_table()
        self.current_pending_num = enc_num
        self.draft_context.setText(f"Encomenda ativa: {enc_num}")
        self.piece_rows = self.backend.expedicao_available_pieces(enc_num) if enc_num else []
        _fill_table(
            self.pieces_table,
            [[r.get("id", "-"), r.get("ref_interna", "-"), r.get("ref_externa", "-"), r.get("estado", "-"), r.get("pronta_expedicao", "0"), r.get("qtd_expedida", "0"), r.get("disponivel", "0"), "SIM" if r.get("desenho") else "NAO"] for r in self.piece_rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.piece_rows):
            for col_index in range(self.pieces_table.columnCount()):
                item = self.pieces_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
            _paint_table_row(self.pieces_table, row_index, str(row.get("estado", "")))
        if self.pieces_table.rowCount() > 0:
            self.pieces_table.selectRow(0)
        self._sync_expedicao_actions()

    def _clear_guide_detail(self) -> None:
        self.detail_title.setText("Seleciona uma guia")
        _apply_state_chip(self.detail_state_chip, "-")
        self.detail_meta.setText("-")
        self.detail_note.setText("Sem detalhe disponível para o filtro atual.")
        self.current_guide_lines = []
        self.lines_table.setRowCount(0)
        _set_panel_tone(self.detail_card, "default")
        self._sync_expedicao_actions()

    def _update_guide_detail(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            self._clear_guide_detail()
            return
        try:
            detail = self.backend.expedicao_detail(numero)
        except Exception as exc:
            self._clear_guide_detail()
            self.detail_note.setText(str(exc))
            return
        state = str(detail.get("estado", "") or "-").strip() or "-"
        self.detail_title.setText(f"Guia {detail.get('numero', '-')}")
        _apply_state_chip(self.detail_state_chip, state)
        self.detail_meta.setText(
            f"Enc {detail.get('encomenda', '-')} | Cliente {detail.get('cliente', '-')}"
            f" | Emissão {detail.get('data_emissao', '-') or '-'} | Transporte {detail.get('data_transporte', '-') or '-'}"
        )
        self.detail_note.setText(
            f"Destino: {detail.get('destinatario', '-') or '-'} | Descarga: {detail.get('local_descarga', '-') or '-'}"
            f" | Transportador: {detail.get('transportador', '-') or '-'} | Matrícula: {detail.get('matricula', '-') or '-'}"
            f" | Obs: {detail.get('observacoes', '-') or '-'}"
        )
        self.current_guide_lines = list(detail.get("lines", []) or [])
        _fill_table(
            self.lines_table,
            [[r.get("ref_interna", "-"), r.get("ref_externa", "-"), r.get("descricao", "-"), r.get("qtd", "0"), r.get("peso", "0"), "SIM" if r.get("manual") else "NAO", r.get("encomenda", "-")] for r in self.current_guide_lines],
            align_center_from=3,
        )
        for row_index in range(self.lines_table.rowCount()):
            for col_index in range(self.lines_table.columnCount()):
                item = self.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        _set_panel_tone(self.detail_card, _state_tone(state))
        self._sync_expedicao_actions()

    def _render_draft_table(self) -> None:
        _fill_table(
            self.draft_table,
            [
                [
                    row.get("peca_id", "-"),
                    row.get("ref_interna", "-"),
                    row.get("ref_externa", "-"),
                    row.get("descricao", "-"),
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                ]
                for row in self.draft_rows
            ],
            align_center_from=4,
        )
        for row_index in range(self.draft_table.rowCount()):
            for col_index in range(self.draft_table.columnCount()):
                item = self.draft_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        self._sync_expedicao_actions()

    def _select_history(self, numero: str) -> None:
        numero = str(numero or "").strip()
        if not numero:
            return
        for index, row in enumerate(self.history_rows):
            if str(row.get("numero", "") or "").strip() == numero:
                self.history_table.selectRow(index)
                break

    def _add_draft_line(self) -> None:
        order = self._current_pending_row()
        piece = self._current_piece_row()
        enc_num = str(order.get("numero", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        if not enc_num or not piece_id:
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda e uma peça disponível.")
            return
        qty = float(self.exp_qty_spin.value() or 0)
        if qty <= 0:
            QMessageBox.warning(self, "Expedição", "Quantidade inválida.")
            return
        used = sum(float(row.get("qtd", 0) or 0) for row in self.draft_rows if str(row.get("peca_id", "") or "").strip() == piece_id)
        available = float(piece.get("disponivel_num", 0) or 0)
        if qty > available - used + 1e-9:
            QMessageBox.warning(self, "Expedição", f"Quantidade superior ao disponível ({max(0.0, available - used):.2f}).")
            return
        merged = False
        for row in self.draft_rows:
            if str(row.get("peca_id", "") or "").strip() == piece_id:
                row["qtd"] = float(row.get("qtd", 0) or 0) + qty
                merged = True
                break
        if not merged:
            self.draft_rows.append(
                {
                    "encomenda": enc_num,
                    "peca_id": piece_id,
                    "ref_interna": piece.get("ref_interna", ""),
                    "ref_externa": piece.get("ref_externa", ""),
                    "descricao": piece.get("descricao", "") or piece.get("ref_externa", "") or piece.get("ref_interna", ""),
                    "qtd": qty,
                    "unid": "UN",
                    "peso": 0.0,
                    "manual": False,
                }
            )
        self._render_draft_table()
        self._on_pending_selected()

    def _remove_draft_line(self) -> None:
        current = self.draft_table.currentItem()
        if current is None or current.row() >= len(self.draft_rows):
            return
        del self.draft_rows[current.row()]
        self._render_draft_table()
        self._on_pending_selected()

    def _clear_draft(self) -> None:
        if not self.draft_rows:
            return
        if QMessageBox.question(self, "Expedição", "Limpar todas as linhas da guia em preparação?") != QMessageBox.Yes:
            return
        self.draft_rows = []
        self._render_draft_table()
        self._on_pending_selected()

    def _normalize_datetime_input(self, raw: str) -> str:
        value = str(raw or "").strip()
        if not value:
            return value
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%dT%H:%M:%S")
            except Exception:
                continue
        return value

    def _text_prompt(self, title: str, label: str, initial: str = "") -> str | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        edit = QLineEdit(str(initial or ""))
        form.addRow(label, edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        return edit.text().strip()

    def _guide_dialog(self, title: str, initial: dict, lines: list[dict]) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(900)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        fields = {
            "codigo_at": QLineEdit(str(initial.get("codigo_at", "") or "").strip()),
            "data_transporte": QLineEdit(str(initial.get("data_transporte", "") or "").replace("T", " ").strip()),
            "emitente_nome": QLineEdit(str(initial.get("emitente_nome", "") or "").strip()),
            "emitente_nif": QLineEdit(str(initial.get("emitente_nif", "") or "").strip()),
            "emitente_morada": QLineEdit(str(initial.get("emitente_morada", "") or "").strip()),
            "destinatario": QLineEdit(str(initial.get("destinatario", "") or "").strip()),
            "dest_nif": QLineEdit(str(initial.get("dest_nif", "") or "").strip()),
            "dest_morada": QLineEdit(str(initial.get("dest_morada", "") or "").strip()),
            "local_carga": QLineEdit(str(initial.get("local_carga", "") or "").strip()),
            "local_descarga": QLineEdit(str(initial.get("local_descarga", "") or "").strip()),
            "transportador": QLineEdit(str(initial.get("transportador", "") or "").strip()),
            "matricula": QLineEdit(str(initial.get("matricula", "") or "").strip()),
        }
        observations = QTextEdit()
        observations.setPlainText(str(initial.get("observacoes", "") or "").strip())
        observations.setFixedHeight(84)
        labels = [
            ("Cod. validação AT", "codigo_at", 0, 0),
            ("Início transporte", "data_transporte", 0, 2),
            ("Emitente", "emitente_nome", 1, 0),
            ("NIF emitente", "emitente_nif", 1, 2),
            ("Morada emitente", "emitente_morada", 2, 0),
            ("Destinatário", "destinatario", 3, 0),
            ("NIF destino", "dest_nif", 3, 2),
            ("Morada destino", "dest_morada", 4, 0),
            ("Local carga", "local_carga", 5, 0),
            ("Local descarga", "local_descarga", 5, 2),
            ("Transportador", "transportador", 6, 0),
            ("Matrícula", "matricula", 6, 2),
        ]
        for label_text, key, row, col in labels:
            form.addWidget(QLabel(label_text), row, col)
            span = 3 if key in {"emitente_morada", "dest_morada"} else 1
            form.addWidget(fields[key], row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 7, 0)
        form.addWidget(observations, 7, 1, 1, 3)
        layout.addLayout(form)
        preview_title = QLabel("Linhas da guia")
        preview_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        preview_table = QTableWidget(0, 5)
        preview_table.setHorizontalHeaderLabels(["Ref. Int.", "Ref. Ext.", "Descrição", "Qtd", "Unid"])
        preview_table.verticalHeader().setVisible(False)
        preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(preview_table, stretch=(1, 2), contents=())
        _set_table_columns(
            preview_table,
            [
                (0, "interactive", 170),
                (1, "stretch", 240),
                (2, "stretch", 300),
                (3, "interactive", 82),
                (4, "interactive", 76),
            ],
        )
        _fill_table(
            preview_table,
            [
                [row.get("ref_interna", "-"), row.get("ref_externa", "-"), row.get("descricao", "-"), row.get("qtd", "0"), row.get("unid", "UN")]
                for row in list(lines or [])
            ],
            align_center_from=3,
        )
        for row_index in range(preview_table.rowCount()):
            for col_index in range(preview_table.columnCount()):
                item = preview_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(str(item.text() or "").strip())
        layout.addWidget(preview_title)
        layout.addWidget(preview_table)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        if not fields["destinatario"].text().strip():
            QMessageBox.warning(self, "Expedição", "Indica o destinatário.")
            return None
        return {
            "codigo_at": fields["codigo_at"].text().strip(),
            "tipo_via": "Original",
            "emitente_nome": fields["emitente_nome"].text().strip(),
            "emitente_nif": fields["emitente_nif"].text().strip(),
            "emitente_morada": fields["emitente_morada"].text().strip(),
            "destinatario": fields["destinatario"].text().strip(),
            "dest_nif": fields["dest_nif"].text().strip(),
            "dest_morada": fields["dest_morada"].text().strip(),
            "local_carga": fields["local_carga"].text().strip(),
            "local_descarga": fields["local_descarga"].text().strip(),
            "data_transporte": self._normalize_datetime_input(fields["data_transporte"].text()),
            "transportador": fields["transportador"].text().strip(),
            "matricula": fields["matricula"].text().strip(),
            "observacoes": observations.toPlainText().strip(),
        }

    def _manual_guide_dialog(self) -> dict | None:
        initial = self.backend.expedicao_manual_defaults()
        dialog = QDialog(self)
        dialog.setWindowTitle("Criar Guia Manual")
        dialog.setMinimumWidth(980)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        fields = {
            "codigo_at": QLineEdit(str(initial.get("codigo_at", "") or "").strip()),
            "emitente_nome": QLineEdit(str(initial.get("emitente_nome", "") or "").strip()),
            "destinatario": QLineEdit(""),
            "dest_nif": QLineEdit(""),
            "dest_morada": QLineEdit(""),
            "local_carga": QLineEdit(str(initial.get("local_carga", "") or "").strip()),
            "local_descarga": QLineEdit(""),
            "transportador": QLineEdit(""),
            "matricula": QLineEdit(""),
        }
        observations = QTextEdit()
        observations.setFixedHeight(84)
        meta_labels = [
            ("Cod. validação AT", "codigo_at", 0, 0),
            ("Emitente", "emitente_nome", 0, 2),
            ("Destinatário", "destinatario", 1, 0),
            ("NIF", "dest_nif", 1, 2),
            ("Morada", "dest_morada", 2, 0),
            ("Local carga", "local_carga", 3, 0),
            ("Local descarga", "local_descarga", 3, 2),
            ("Transportador", "transportador", 4, 0),
            ("Matrícula", "matricula", 4, 2),
        ]
        for label_text, key, row, col in meta_labels:
            form.addWidget(QLabel(label_text), row, col)
            span = 3 if key == "dest_morada" else 1
            form.addWidget(fields[key], row, col + 1, 1, span)
        form.addWidget(QLabel("Observações"), 5, 0)
        form.addWidget(observations, 5, 1, 1, 3)
        layout.addLayout(form)

        product_rows = self.backend.expedicao_product_options("")
        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.addItem("", {})
        for row in product_rows:
            product_combo.addItem(f"{row.get('codigo', '')} - {row.get('descricao', '')}", row)
        desc_edit = QLineEdit()
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(1.0)
        unid_edit = QLineEdit("UN")
        line_items: list[dict] = []
        lines_table = QTableWidget(0, 4)
        lines_table.setHorizontalHeaderLabels(["Código", "Descrição", "Qtd", "Unid"])
        lines_table.verticalHeader().setVisible(False)
        lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        lines_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        def render_manual_lines() -> None:
            _fill_table(
                lines_table,
                [[row.get("produto_codigo", "-"), row.get("descricao", "-"), f"{float(row.get('qtd', 0) or 0):.2f}", row.get("unid", "UN")] for row in line_items],
                align_center_from=2,
            )

        def on_product_change() -> None:
            payload = product_combo.currentData()
            if isinstance(payload, dict):
                desc_edit.setText(str(payload.get("descricao", "") or "").strip())
                unid_edit.setText(str(payload.get("unid", "UN") or "UN").strip() or "UN")

        def add_line() -> None:
            payload = product_combo.currentData()
            product_code = ""
            description = desc_edit.text().strip()
            if isinstance(payload, dict):
                product_code = str(payload.get("codigo", "") or "").strip()
                if not description:
                    description = str(payload.get("descricao", "") or "").strip()
            if not description and not product_code:
                QMessageBox.warning(dialog, "Expedição", "Seleciona um produto ou indica uma descrição.")
                return
            line_items.append(
                {
                    "produto_codigo": product_code,
                    "descricao": description,
                    "qtd": qty_spin.value(),
                    "unid": unid_edit.text().strip() or "UN",
                }
            )
            render_manual_lines()
            qty_spin.setValue(1.0)
            desc_edit.clear()

        def remove_line() -> None:
            current = lines_table.currentItem()
            if current is None or current.row() >= len(line_items):
                return
            del line_items[current.row()]
            render_manual_lines()

        product_combo.currentIndexChanged.connect(lambda _index: on_product_change())
        line_controls = QHBoxLayout()
        line_controls.addWidget(QLabel("Produto"))
        line_controls.addWidget(product_combo, 1)
        line_controls.addWidget(QLabel("Descrição"))
        line_controls.addWidget(desc_edit, 1)
        line_controls.addWidget(QLabel("Qtd"))
        line_controls.addWidget(qty_spin)
        line_controls.addWidget(QLabel("Unid"))
        line_controls.addWidget(unid_edit)
        add_line_btn = QPushButton("Adicionar linha")
        add_line_btn.clicked.connect(add_line)
        remove_line_btn = QPushButton("Remover linha")
        remove_line_btn.setProperty("variant", "secondary")
        remove_line_btn.clicked.connect(remove_line)
        line_controls.addWidget(add_line_btn)
        line_controls.addWidget(remove_line_btn)
        layout.addLayout(line_controls)
        layout.addWidget(lines_table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None
        if not fields["destinatario"].text().strip():
            QMessageBox.warning(self, "Expedição", "Indica o destinatário.")
            return None
        if not line_items:
            QMessageBox.warning(self, "Expedição", "Adiciona pelo menos uma linha.")
            return None
        return {
            "guide": {
                "codigo_at": fields["codigo_at"].text().strip(),
                "tipo_via": "Original",
                "emitente_nome": fields["emitente_nome"].text().strip(),
                "emitente_nif": str(initial.get("emitente_nif", "") or "").strip(),
                "emitente_morada": str(initial.get("emitente_morada", "") or "").strip(),
                "destinatario": fields["destinatario"].text().strip(),
                "dest_nif": fields["dest_nif"].text().strip(),
                "dest_morada": fields["dest_morada"].text().strip(),
                "local_carga": fields["local_carga"].text().strip(),
                "local_descarga": fields["local_descarga"].text().strip(),
                "data_transporte": self._normalize_datetime_input(str(initial.get("data_transporte", ""))),
                "transportador": fields["transportador"].text().strip(),
                "matricula": fields["matricula"].text().strip(),
                "observacoes": observations.toPlainText().strip(),
            },
            "lines": line_items,
        }

    def _emit_off_guide(self) -> None:
        order = self._current_pending_row()
        enc_num = str(order.get("numero", "") or "").strip()
        if not enc_num:
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda.")
            return
        if not self.draft_rows:
            QMessageBox.warning(self, "Expedição", "Sem linhas na guia.")
            return
        try:
            initial = self.backend.expedicao_defaults_for_order(enc_num)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        payload = self._guide_dialog("Confirmar Dados da Guia OFF", initial, self.draft_rows)
        if payload is None:
            return
        try:
            detail = self.backend.expedicao_emit_off(enc_num, self.draft_rows, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Emitir guia", str(exc))
            return
        self.draft_rows = []
        self.refresh()
        self._select_history(detail.get("numero", ""))
        if QMessageBox.question(self, "Expedição", f"Guia emitida: {detail.get('numero', '-')}\n\nAbrir PDF agora?") == QMessageBox.Yes:
            self._open_guide_pdf()

    def _create_manual_guide(self) -> None:
        payload = self._manual_guide_dialog()
        if payload is None:
            return
        try:
            detail = self.backend.expedicao_emit_manual(payload.get("guide", {}), payload.get("lines", []))
        except Exception as exc:
            QMessageBox.critical(self, "Guia manual", str(exc))
            return
        self.refresh()
        self._select_history(detail.get("numero", ""))
        if QMessageBox.question(self, "Expedição", f"Guia manual emitida: {detail.get('numero', '-')}\n\nAbrir PDF agora?") == QMessageBox.Yes:
            self._open_guide_pdf()

    def _edit_guide(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        try:
            detail = self.backend.expedicao_detail(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        payload = self._guide_dialog(f"Editar Guia {numero}", detail, list(detail.get("lines", []) or []))
        if payload is None:
            return
        try:
            self.backend.expedicao_update(numero, payload)
        except Exception as exc:
            QMessageBox.critical(self, "Editar guia", str(exc))
            return
        self.refresh()
        self._select_history(numero)

    def _cancel_guide(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        if QMessageBox.question(self, "Anular Guia", f"Anular guia {numero}?") != QMessageBox.Yes:
            return
        motivo = self._text_prompt("Anular Guia", "Justificação")
        if motivo is None:
            return
        try:
            self.backend.expedicao_cancel(numero, motivo)
        except Exception as exc:
            QMessageBox.critical(self, "Anular guia", str(exc))
            return
        self.refresh()
        self._select_history(numero)

    def _open_guide_pdf(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        try:
            path = self.backend.expedicao_open_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Expedição", str(exc))
            return
        QMessageBox.information(self, "Expedição", f"PDF aberto:\n{path}")

    def _save_guide_pdf(self) -> None:
        current = self._current_history_row()
        numero = str(current.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Expedição", "Seleciona uma guia.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", f"guia_{numero}.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            self.backend.expedicao_render_pdf(numero, path)
        except Exception as exc:
            QMessageBox.critical(self, "Guardar PDF", str(exc))
            return
        QMessageBox.information(self, "Guardar PDF", f"PDF guardado em:\n{path}")

    def _show_history_dialog(self) -> None:
        self.refresh()
        dialog = QDialog(self)
        dialog.setWindowTitle("Histórico de guias")
        dialog.resize(1100, 520)
        layout = QVBoxLayout(dialog)
        title = QLabel("Histórico de guias emitidas")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(title)
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["Guia", "Tipo", "Encomenda", "Cliente", "Emissao", "Estado", "Linhas"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        _configure_table(table, stretch=(3,), contents=(0, 1, 2, 4, 5, 6))
        _fill_table(
            table,
            [[r.get("numero", "-"), r.get("tipo", "-"), r.get("encomenda", "-"), r.get("cliente", "-"), r.get("data_emissao", "-"), r.get("estado", "-"), r.get("linhas", 0)] for r in self.history_rows],
            align_center_from=6,
        )
        for row_index, row in enumerate(self.history_rows):
            _paint_table_row(table, row_index, str(row.get("estado", "")))
        layout.addWidget(table, 1)
        if not self.history_rows:
            empty = QLabel("Sem guias emitidas para o filtro atual.")
            empty.setProperty("role", "muted")
            layout.addWidget(empty)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _open_piece_drawing(self) -> None:
        order = self._current_pending_row()
        piece = self._current_piece_row()
        enc_num = str(order.get("numero", "") or "").strip()
        piece_id = str(piece.get("id", "") or "").strip()
        if not enc_num or not piece_id:
            QMessageBox.warning(self, "Expedição", "Seleciona uma peça disponível.")
            return
        try:
            self.backend.operator_open_drawing(enc_num, piece_id)
        except Exception as exc:
            QMessageBox.critical(self, "Ver desenho", str(exc))


class LegacyExpeditionPage(ExpeditionPage):
    page_subtitle = "Expedição por encomenda, com registo da guia apenas no detalhe da encomenda."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        self.pending_table.verticalHeader().setDefaultSectionSize(34)
        self.pending_table.horizontalHeader().setFixedHeight(34)
        _set_table_columns(
            self.pending_table,
            [
                (0, "fixed", 190),
                (1, "stretch", 0),
                (2, "fixed", 130),
                (3, "fixed", 128),
                (4, "fixed", 138),
                (5, "fixed", 76),
            ],
        )
        root = self.layout()
        sections = _take_layout_items(root)
        actions_item = sections[0] if len(sections) > 0 else None
        filters_item = sections[1] if len(sections) > 1 else None
        pending_item = sections[2] if len(sections) > 2 else None
        pieces_item = sections[3] if len(sections) > 3 else None
        draft_item = sections[4] if len(sections) > 4 else None
        history_item = sections[5] if len(sections) > 5 else None
        detail_item = sections[6] if len(sections) > 6 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(10)
        _adopt_layout_item(list_layout, filters_item)
        list_actions = CardFrame()
        list_actions.set_tone("info")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(14, 10, 14, 10)
        list_actions_layout.setSpacing(8)
        self.open_pending_btn = QPushButton("Abrir encomenda")
        self.open_pending_btn.clicked.connect(self._open_selected_pending)
        history_btn = QPushButton("Histórico Guias")
        history_btn.setProperty("variant", "secondary")
        history_btn.clicked.connect(self._show_history_dialog)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.refresh)
        list_actions_layout.addWidget(self.open_pending_btn)
        list_actions_layout.addWidget(history_btn)
        list_actions_layout.addWidget(refresh_btn)
        list_actions_layout.addStretch(1)
        list_layout.addWidget(list_actions)
        _adopt_layout_item(list_layout, pending_item, 1)
        _adopt_layout_item(list_layout, history_item, 1)

        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)
        detail_top = CardFrame()
        detail_top.set_tone("default")
        detail_top_layout = QHBoxLayout(detail_top)
        detail_top_layout.setContentsMargins(14, 10, 14, 10)
        detail_top_layout.setSpacing(8)
        back_btn = QPushButton("Voltar a encomendas")
        back_btn.setProperty("variant", "secondary")
        back_btn.clicked.connect(self._show_pending_list)
        focus_block = QVBoxLayout()
        focus_block.setSpacing(2)
        self.pending_focus_label = QLabel("Sem encomenda selecionada")
        self.pending_focus_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.pending_meta_label = QLabel("Seleciona uma encomenda para preparar e emitir a guia.")
        self.pending_meta_label.setProperty("role", "muted")
        self.pending_state_chip = QLabel("-")
        _apply_state_chip(self.pending_state_chip, "-")
        focus_block.addWidget(self.pending_focus_label)
        focus_block.addWidget(self.pending_meta_label)
        detail_top_layout.addWidget(back_btn)
        detail_top_layout.addLayout(focus_block, 1)
        detail_top_layout.addWidget(self.pending_state_chip, 0, Qt.AlignTop)
        detail_layout.addWidget(detail_top)
        _adopt_layout_item(detail_layout, actions_item)
        top_split = QSplitter(Qt.Horizontal)
        top_split.setChildrenCollapsible(False)
        top_left = QWidget()
        top_left_layout = QVBoxLayout(top_left)
        top_left_layout.setContentsMargins(0, 0, 0, 0)
        top_left_layout.setSpacing(10)
        _adopt_layout_item(top_left_layout, pieces_item, 1)
        top_right = QWidget()
        top_right_layout = QVBoxLayout(top_right)
        top_right_layout.setContentsMargins(0, 0, 0, 0)
        top_right_layout.setSpacing(10)
        _adopt_layout_item(top_right_layout, draft_item, 1)
        top_split.addWidget(top_left)
        top_split.addWidget(top_right)
        top_split.setSizes([1120, 780])

        bottom_split = QSplitter(Qt.Horizontal)
        bottom_split.setChildrenCollapsible(False)
        bottom_right = QWidget()
        bottom_right_layout = QVBoxLayout(bottom_right)
        bottom_right_layout.setContentsMargins(0, 0, 0, 0)
        bottom_right_layout.setSpacing(10)
        _adopt_layout_item(bottom_right_layout, detail_item, 1)
        bottom_split.addWidget(bottom_right)
        bottom_split.setSizes([1900])

        vertical_split = QSplitter(Qt.Vertical)
        vertical_split.setChildrenCollapsible(False)
        vertical_split.addWidget(top_split)
        vertical_split.addWidget(bottom_split)
        vertical_split.setSizes([430, 340])
        detail_layout.addWidget(vertical_split, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.pending_table.itemSelectionChanged.connect(self._sync_pending_focus)
        self.pending_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_pending())
        self._show_pending_list()
        self._sync_pending_focus()

    def refresh(self) -> None:
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(self.current_pending_num)
        ExpeditionPage.refresh(self)
        self._sync_pending_focus()
        if self.pending_table.rowCount() == 0:
            self._show_pending_list()
        elif keep_detail:
            self._show_pending_detail()
        else:
            self._show_pending_list()

    def _show_pending_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_pending_focus()

    def _show_pending_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)
        self._sync_pending_focus()

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _sync_pending_focus(self) -> None:
        row = self._current_pending_row()
        has_pending = bool(row)
        self.open_pending_btn.setEnabled(has_pending)
        if not has_pending:
            self.pending_focus_label.setText("Sem encomenda selecionada")
            self.pending_meta_label.setText("Seleciona uma encomenda para preparar e emitir a guia.")
            _apply_state_chip(self.pending_state_chip, "-")
            return
        estado = str(row.get("estado_expedicao", row.get("estado", "-")) or "-").strip() or "-"
        self.pending_focus_label.setText(f"{row.get('numero', '-')} | {row.get('cliente', '-')}")
        self.pending_meta_label.setText(
            f"Estado producao {row.get('estado', '-')} | Expedição {row.get('estado_expedicao', '-')}"
            f" | Disponivel {float(row.get('disponivel', 0) or 0):.1f} | Entrega {row.get('data_entrega', '-') or '-'}"
        )
        _apply_state_chip(self.pending_state_chip, estado)

    def _on_pending_selected(self) -> None:
        ExpeditionPage._on_pending_selected(self)
        self._sync_pending_focus()

    def _open_selected_pending(self) -> None:
        if not self._current_pending_row():
            QMessageBox.warning(self, "Expedição", "Seleciona uma encomenda.")
            return
        self._on_pending_selected()
        self._show_pending_detail()
