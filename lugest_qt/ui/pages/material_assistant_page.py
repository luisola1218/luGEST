from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard


def _fmt_number(value: object) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    if abs(number - round(number)) <= 1e-9:
        return str(int(round(number)))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def _table_selected_row(table: QTableWidget) -> int:
    selected = table.selectionModel().selectedRows() if table.selectionModel() is not None else []
    if selected:
        return selected[0].row()
    current = table.currentRow()
    return current if current >= 0 else -1


def _tone_colors(tone: str) -> tuple[str, str]:
    tone_txt = str(tone or "default").strip().lower()
    if tone_txt == "danger":
        return "#fff1f2", "#b42318"
    if tone_txt == "warning":
        return "#fff8e6", "#a16207"
    if tone_txt == "success":
        return "#ecfdf3", "#166534"
    if tone_txt == "info":
        return "#eef4ff", "#1d4ed8"
    return "#f8fafc", "#334155"


def _apply_chip(label: QLabel, text: str, tone: str = "default") -> None:
    bg, fg = _tone_colors(tone)
    label.setText(str(text or "-") or "-")
    label.setStyleSheet(
        "padding: 6px 12px; border-radius: 999px; font-weight: 800;"
        f"background: {bg}; color: {fg}; border: 1px solid #cbd5e1;"
    )


def _paint_row(table: QTableWidget, row: int, tone: str, status_key: str = "") -> None:
    bg_hex, fg_hex = _tone_colors("default" if status_key == "ignored" else tone)
    if status_key == "accepted":
        bg_hex, fg_hex = _tone_colors("success")
    bg = QBrush(QColor(bg_hex))
    fg = QBrush(QColor(fg_hex))
    for col in range(table.columnCount()):
        item = table.item(row, col)
        if item is None:
            continue
        item.setBackground(bg)
        item.setForeground(fg)


class MaterialAssistantPage(QWidget):
    page_title = "Assistente MP"
    page_subtitle = "Separacao inteligente de materia-prima em tempo real, ligada ao planeamento e ao stock."
    uses_backend_reload = True
    allow_auto_timer_refresh = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.snapshot: dict = {}
        self.suggestions: list[dict] = []
        self.needs: list[dict] = []
        self._selected_suggestion_id = ""
        self._selected_need_key = ""

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        self.cards = [StatCard(title) for title in ("Linhas separacao", "Trocas por urgencia", "Nao arrumar", "Sem stock")]
        for idx, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, idx)
        root.addWidget(cards_host)

        controls = CardFrame()
        controls.set_tone("info")
        controls_layout = QVBoxLayout(controls)
        controls_layout.setContentsMargins(16, 14, 16, 14)
        controls_layout.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        top_row.addWidget(QLabel("Horizonte"))
        self.horizon_combo = QComboBox()
        self.horizon_combo.addItem("Hoje", 1)
        self.horizon_combo.addItem("48 horas", 2)
        self.horizon_combo.addItem("5 dias", 5)
        self.horizon_combo.addItem("10 dias", 10)
        self.horizon_combo.setCurrentIndex(2)
        self.horizon_combo.currentTextChanged.connect(self.refresh)
        top_row.addWidget(self.horizon_combo)
        top_row.addStretch(1)

        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.accept_btn = QPushButton("Validar")
        self.accept_btn.clicked.connect(lambda: self._apply_feedback("accepted"))
        self.ignore_btn = QPushButton("Ignorar hoje")
        self.ignore_btn.setProperty("variant", "secondary")
        self.ignore_btn.clicked.connect(lambda: self._apply_feedback("ignored"))
        self.clear_btn = QPushButton("Limpar decisao")
        self.clear_btn.setProperty("variant", "secondary")
        self.clear_btn.clicked.connect(lambda: self._apply_feedback("clear"))
        self.alerts_btn = QPushButton("Sugestoes / alertas")
        self.alerts_btn.setProperty("variant", "secondary")
        self.alerts_btn.clicked.connect(self._open_alerts_dialog)
        self.separation_btn = QPushButton("Separacao - MP")
        self.separation_btn.setProperty("variant", "secondary")
        self.separation_btn.clicked.connect(self._open_separation_dialog)
        self.stock_btn = QPushButton("Ver stock")
        self.stock_btn.setProperty("variant", "secondary")
        self.stock_btn.clicked.connect(self._open_stock_dialog)
        self.order_btn = QPushButton("Ver encomenda")
        self.order_btn.setProperty("variant", "secondary")
        self.order_btn.clicked.connect(self._open_order_dialog)
        for widget in (
            self.refresh_btn,
            self.accept_btn,
            self.ignore_btn,
            self.clear_btn,
            self.alerts_btn,
            self.separation_btn,
            self.stock_btn,
            self.order_btn,
        ):
            top_row.addWidget(widget)
        controls_layout.addLayout(top_row)

        self.summary_label = QLabel("Sem dados para mostrar.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setProperty("role", "muted")
        controls_layout.addWidget(self.summary_label)
        root.addWidget(controls)

        main_split = QSplitter(Qt.Vertical)
        main_split.setChildrenCollapsible(False)

        suggestions_card = CardFrame()
        suggestions_card.set_tone("warning")
        suggestions_layout = QVBoxLayout(suggestions_card)
        suggestions_layout.setContentsMargins(14, 12, 14, 12)
        suggestions_layout.setSpacing(8)
        suggestions_title = QLabel("Sugestoes / alertas")
        suggestions_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        suggestions_layout.addWidget(suggestions_title)
        suggestions_subtitle = QLabel(
            "Mostra apenas trocas por urgencia, falta de stock e materiais a manter prontos."
        )
        suggestions_subtitle.setProperty("role", "muted")
        suggestions_subtitle.setWordWrap(True)
        suggestions_layout.addWidget(suggestions_subtitle)

        self.suggestions_table = QTableWidget(0, 8)
        self.suggestions_table.setHorizontalHeaderLabels(
            ["Prio", "Tema", "Encomenda", "Cliente", "Material", "Recomendacao", "Quando", "Estado"]
        )
        self.suggestions_table.verticalHeader().setVisible(False)
        self.suggestions_table.verticalHeader().setDefaultSectionSize(28)
        self.suggestions_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.suggestions_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.suggestions_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.suggestions_table.setAlternatingRowColors(True)
        self.suggestions_table.itemSelectionChanged.connect(self._sync_selection_detail)
        header = self.suggestions_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.suggestions_table.setColumnWidth(2, 150)
        self.suggestions_table.setColumnWidth(3, 220)
        self.suggestions_table.setColumnWidth(4, 150)
        suggestions_layout.addWidget(self.suggestions_table)
        main_split.addWidget(suggestions_card)

        lower_split = QSplitter(Qt.Horizontal)
        lower_split.setChildrenCollapsible(False)

        needs_card = CardFrame()
        needs_card.set_tone("info")
        needs_layout = QVBoxLayout(needs_card)
        needs_layout.setContentsMargins(14, 12, 14, 12)
        needs_layout.setSpacing(8)
        needs_title = QLabel("Separacao orientada pelo planeamento")
        needs_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        needs_layout.addWidget(needs_title)
        needs_subtitle = QLabel(
            "Leitura rapida para o separador: encomenda, espessura, lote, dimensao e quantidade."
        )
        needs_subtitle.setProperty("role", "muted")
        needs_subtitle.setWordWrap(True)
        needs_layout.addWidget(needs_subtitle)

        self.needs_table = QTableWidget(0, 10)
        self.needs_table.setHorizontalHeaderLabels(
            ["Encomenda", "Posto", "Cliente", "Material", "Esp.", "Lote", "Dimensao", "Qtd.", "Entrega", "Planeado"]
        )
        self.needs_table.verticalHeader().setVisible(False)
        self.needs_table.verticalHeader().setDefaultSectionSize(28)
        self.needs_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.needs_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.needs_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.needs_table.setAlternatingRowColors(True)
        self.needs_table.itemSelectionChanged.connect(self._sync_selection_detail)
        needs_header = self.needs_table.horizontalHeader()
        needs_header.setSectionResizeMode(0, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(1, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(2, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(3, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        needs_header.setSectionResizeMode(5, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(6, QHeaderView.Interactive)
        needs_header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        needs_header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        needs_header.setSectionResizeMode(9, QHeaderView.Stretch)
        self.needs_table.setColumnWidth(0, 180)
        self.needs_table.setColumnWidth(1, 120)
        self.needs_table.setColumnWidth(2, 220)
        self.needs_table.setColumnWidth(3, 160)
        self.needs_table.setColumnWidth(5, 120)
        self.needs_table.setColumnWidth(6, 130)
        needs_layout.addWidget(self.needs_table)
        lower_split.addWidget(needs_card)

        detail_card = CardFrame()
        detail_card.set_tone("default")
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(16, 14, 16, 14)
        detail_layout.setSpacing(10)
        detail_top = QHBoxLayout()
        detail_top.setSpacing(8)
        self.detail_title = QLabel("Seleciona uma linha")
        self.detail_title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        self.detail_status = QLabel("Nova")
        _apply_chip(self.detail_status, "Nova", "warning")
        detail_top.addWidget(self.detail_title, 1)
        detail_top.addWidget(self.detail_status, 0, Qt.AlignTop)
        detail_layout.addLayout(detail_top)

        self.detail_meta = QLabel("Sem detalhe ainda.")
        self.detail_meta.setProperty("role", "muted")
        self.detail_meta.setWordWrap(True)
        detail_layout.addWidget(self.detail_meta)

        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        self.detail_text.setFrameStyle(QFrame.NoFrame)
        self.detail_text.setMinimumHeight(220)
        detail_layout.addWidget(self.detail_text, 1)
        lower_split.addWidget(detail_card)
        lower_split.setSizes([1180, 620])

        main_split.addWidget(lower_split)
        main_split.setSizes([360, 520])
        root.addWidget(main_split, 1)

    def can_auto_refresh(self) -> bool:
        return True

    def _selected_suggestion(self) -> dict | None:
        row = _table_selected_row(self.suggestions_table)
        if row < 0:
            return None
        item = self.suggestions_table.item(row, 0)
        return dict(item.data(Qt.UserRole) or {}) if item is not None else None

    def _selected_need(self) -> dict | None:
        row = _table_selected_row(self.needs_table)
        if row < 0:
            return None
        item = self.needs_table.item(row, 0)
        return dict(item.data(Qt.UserRole) or {}) if item is not None else None

    def _selected_context(self) -> tuple[str, str, str, str]:
        suggestion = self._selected_suggestion()
        if suggestion:
            return (
                str(suggestion.get("numero", "") or "").strip(),
                str(suggestion.get("material", "") or "").strip(),
                str(suggestion.get("espessura", "") or "").strip(),
                str(suggestion.get("need_key", "") or "").strip(),
            )
        need = self._selected_need()
        if need:
            return (
                str(need.get("numero", "") or "").strip(),
                str(need.get("material", "") or "").strip(),
                str(need.get("espessura", "") or "").strip(),
                str(need.get("key", "") or need.get("need_key", "") or "").strip(),
            )
        return "", "", "", ""

    def refresh(self) -> None:
        horizon = int(self.horizon_combo.currentData() or 5)
        selected_suggestion_id = str(self._selected_suggestion_id or "")
        selected_need_key = str(self._selected_need_key or "")
        self.snapshot = dict(self.backend.material_assistant_snapshot(horizon_days=horizon) or {})
        self.suggestions = list(self.snapshot.get("suggestions", []) or [])
        self.needs = list(self.snapshot.get("needs", []) or [])

        for card, payload in zip(self.cards, list(self.snapshot.get("cards", []) or [])):
            card.title_label.setText(str(payload.get("title", "") or "-"))
            card.set_data(str(payload.get("value", "-") or "-"), str(payload.get("subtitle", "") or ""))
            card.set_tone(str(payload.get("tone", "default") or "default"))

        generated_at = str(self.snapshot.get("generated_at", "") or "").strip()
        generated_label = generated_at[:16].replace("T", " ") if generated_at else "-"
        self.summary_label.setText(
            f"Horizonte de {horizon} dias | {len(self.needs)} linhas de separacao | "
            f"{len(self.suggestions)} alertas | Atualizado {generated_label}"
        )

        self.suggestions_table.setRowCount(len(self.suggestions))
        restore_suggestion_row = 0
        for row_index, row in enumerate(self.suggestions):
            values = [
                str(row.get("priority_label", "") or "-"),
                self._kind_label(str(row.get("kind", "") or "")),
                str(row.get("numero", "") or "").strip(),
                str(row.get("cliente", "") or "").strip(),
                f"{row.get('material', '')} {row.get('espessura', '')}".strip(),
                str(row.get("recommendation", "") or "").strip(),
                str(row.get("when", "") or "-").strip(),
                str(row.get("status_label", "") or "Nova").strip(),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index in (0, 1, 2, 6, 7):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                if col_index == 0:
                    item.setData(Qt.UserRole, dict(row))
                self.suggestions_table.setItem(row_index, col_index, item)
            _paint_row(
                self.suggestions_table,
                row_index,
                str(row.get("priority_tone", "") or "default"),
                str(row.get("status_key", "") or ""),
            )
            if selected_suggestion_id and str(row.get("id", "") or "").strip() == selected_suggestion_id:
                restore_suggestion_row = row_index

        self.needs_table.setColumnCount(10)
        self.needs_table.setRowCount(len(self.needs))
        restore_need_row = 0
        for row_index, row in enumerate(self.needs):
            material_cativado = bool(row.get("material_cativado"))
            values = [
                str(row.get("numero", "") or "").strip(),
                str(row.get("posto_trabalho", "") or "-").strip(),
                str(row.get("cliente", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                (
                    str(row.get("preferred_lot", "") or row.get("current_lot", "") or "-").strip() or "-"
                    if material_cativado
                    else "Sem material cativado"
                ),
                (
                    str(row.get("preferred_dimensao", "") or "-").strip() or "-"
                    if material_cativado
                    else "-"
                ),
                _fmt_number(row.get("quantidade_preparar", row.get("piece_qty", 0))),
                str(row.get("data_entrega", "") or "-").strip() or "-",
                str(row.get("next_action_label", "") or "-").strip(),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index in (0, 1, 4, 7, 8):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                if col_index == 0:
                    item.setData(Qt.UserRole, dict(row))
                self.needs_table.setItem(row_index, col_index, item)
            tone = "danger" if not bool(row.get("stock_ready")) else "warning" if (not material_cativado or bool(row.get("lot_change_required"))) else "success"
            _paint_row(self.needs_table, row_index, tone)
            if selected_need_key and str(row.get("key", "") or "").strip() == selected_need_key:
                restore_need_row = row_index

        if self.suggestions:
            self.suggestions_table.selectRow(min(restore_suggestion_row, len(self.suggestions) - 1))
        elif self.needs:
            self.needs_table.selectRow(min(restore_need_row, len(self.needs) - 1))
        else:
            self._render_empty_detail()
        self._sync_action_buttons()

    def _render_empty_detail(self) -> None:
        self.detail_title.setText("Sem alertas operacionais neste momento")
        _apply_chip(self.detail_status, "Sem alertas", "success")
        self.detail_meta.setText("O planeamento e o stock nao pedem acao imediata.")
        self.detail_text.setHtml(
            "<b>Leitura rapida</b><br>"
            "Quando surgirem trocas de lote por urgencia, faltas de stock ou materiais a manter disponiveis, "
            "aparecem aqui."
        )

    def _kind_label(self, kind: str) -> str:
        mapping = {
            "shortage": "Sem stock",
            "uncativated": "Sem cativacao",
            "keep_ready": "Nao arrumar",
            "fito_lot": "Urgencia / FITO",
        }
        return mapping.get(str(kind or "").strip(), str(kind or "").strip() or "-")

    def _sync_action_buttons(self) -> None:
        has_suggestion = self._selected_suggestion() is not None
        has_context = bool(self._selected_context()[0])
        self.accept_btn.setEnabled(has_suggestion)
        self.ignore_btn.setEnabled(has_suggestion)
        self.clear_btn.setEnabled(has_suggestion)
        self.stock_btn.setEnabled(has_context)
        self.order_btn.setEnabled(has_context)

    def _sync_selection_detail(self) -> None:
        suggestion = self._selected_suggestion()
        if suggestion is not None:
            self._selected_suggestion_id = str(suggestion.get("id", "") or "").strip()
            self._selected_need_key = str(suggestion.get("need_key", "") or "").strip()
            self._render_suggestion_detail(suggestion)
            self._sync_action_buttons()
            return
        need = self._selected_need()
        if need is not None:
            self._selected_need_key = str(need.get("key", "") or "").strip()
            self._render_need_detail(need)
            self._sync_action_buttons()
            return
        self._render_empty_detail()
        self._sync_action_buttons()

    def _render_suggestion_detail(self, suggestion: dict) -> None:
        self.detail_title.setText(str(suggestion.get("headline", "") or "Sugestao"))
        _apply_chip(
            self.detail_status,
            str(suggestion.get("status_label", "") or "Nova"),
            str(suggestion.get("status_tone", "") or "default"),
        )
        self.detail_meta.setText(
            " | ".join(
                value
                for value in (
                    f"Encomenda {suggestion.get('numero', '-')}",
                    str(suggestion.get("cliente", "") or "").strip(),
                    f"Material {suggestion.get('material', '-')}",
                    f"{suggestion.get('espessura', '-')} mm",
                    f"Prioridade {suggestion.get('priority_label', '-')}",
                )
                if str(value).strip()
            )
        )
        detail_lines = [str(line or "").strip() for line in list(suggestion.get("detail_lines", []) or []) if str(line or "").strip()]
        html = [f"<b>Recomendacao</b><br>{suggestion.get('recommendation', '-') or '-'}"]
        if detail_lines:
            html.append("<br><br><b>Porque apareceu</b><ul>")
            html.extend(f"<li>{line}</li>" for line in detail_lines)
            html.append("</ul>")
        if suggestion.get("stock_state"):
            html.append(f"<br><b>Estado do stock:</b> {suggestion.get('stock_state')}")
        self.detail_text.setHtml("".join(html))

    def _render_need_detail(self, need: dict) -> None:
        self.detail_title.setText(
            f"{need.get('numero', '-')} | {need.get('material', '-')} {need.get('espessura', '-')} mm"
        )
        material_cativado = bool(need.get("material_cativado"))
        tone = "danger" if not bool(need.get("stock_ready")) else "warning" if (not material_cativado or bool(need.get("lot_change_required"))) else "success"
        label = (
            "Sem material cativado"
            if not material_cativado
            else "Troca por urgencia" if bool(need.get("lot_change_required")) else "Necessidade ativa"
        )
        _apply_chip(self.detail_status, label, tone)
        self.detail_meta.setText(
            " | ".join(
                value
                for value in (
                    str(need.get("cliente", "") or "").strip(),
                    f"Entrega {need.get('data_entrega', '-') or '-'}",
                    f"Planeado {need.get('next_action_label', '-') or '-'}",
                )
                if str(value).strip()
            )
        )
        html = [
            "<b>Leitura rapida</b><ul>",
            f"<li>Estado de cativação: {'Cativado' if material_cativado else 'Sem material cativado'}</li>",
            f"<li>Lote atual: {need.get('current_lot', '-') or '-'}</li>",
            f"<li>Lote sugerido: {need.get('preferred_lot', '-') or '-' if material_cativado else '-'}</li>",
            f"<li>Dimensao sugerida: {need.get('preferred_dimensao', '-') or '-' if material_cativado else '-'}</li>",
            f"<li>Quantidade a preparar: {_fmt_number(need.get('quantidade_preparar', need.get('piece_qty', 0)))}</li>",
            f"<li>Retalhos encontrados: {int(need.get('retalho_count', 0) or 0)}</li>",
            f"<li>Estado de stock: {need.get('stock_state', '-') or '-'}</li>",
            "</ul>",
        ]
        if not material_cativado:
            html.append(
                "<br><b>Regra operacional</b><br>"
                "A separação só pode indicar uma chapa/lote depois de existir cativação. "
                "Enquanto isso, esta linha é apenas uma necessidade planeada."
            )
        if bool(need.get("lot_change_required")):
            html.append(
                "<br><b>Regra de urgencia</b><br>"
                f"{need.get('lot_change_note', '-') or '-'}<br>"
                f"Encomenda em conflito: {need.get('lot_change_conflict_order', '-') or '-'}"
            )
        self.detail_text.setHtml("".join(html))

    def _apply_feedback(self, decision: str) -> None:
        suggestion = self._selected_suggestion()
        if suggestion is None:
            return
        try:
            self.backend.material_assistant_set_feedback(str(suggestion.get("id", "") or "").strip(), decision)
        except Exception as exc:
            QMessageBox.critical(self, "Assistente MP", str(exc))
            return
        self.refresh()

    def _open_stock_dialog(self) -> None:
        numero, material, espessura, _need_key = self._selected_context()
        if not numero or not material or not espessura:
            return
        rows = list(self.backend.material_candidates(material, espessura) or [])
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Stock sugerido | {numero}")
        dialog.resize(980, 520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)
        title = QLabel(f"{material} {espessura} mm")
        title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        layout.addWidget(title)
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["ID", "Tipo", "Lote", "Dimensao", "Local", "Disponivel", "Reservado"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = [
                str(row.get("material_id", "") or "").strip(),
                "Retalho" if bool(row.get("is_retalho")) else "Lote",
                str(row.get("lote", "") or "-").strip(),
                str(row.get("dimensao", "") or "-").strip(),
                str(row.get("local", "") or "-").strip(),
                _fmt_number(row.get("disponivel", 0)),
                _fmt_number(row.get("reservado", 0)),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index in (1, 5, 6):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                table.setItem(row_index, col_index, item)
            _paint_row(table, row_index, "success" if bool(row.get("is_retalho")) else "info")
        layout.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _open_alerts_dialog(self) -> None:
        horizon = int(self.horizon_combo.currentData() or 5)
        rows = list(self.backend.material_assistant_alert_rows(horizon_days=horizon) or [])
        dialog = QDialog(self)
        dialog.setWindowTitle("Sugestoes / Alertas")
        dialog.resize(1420, 760)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        top = QHBoxLayout()
        top.setSpacing(8)
        title_col = QVBoxLayout()
        title_col.setContentsMargins(0, 0, 0, 0)
        title_col.setSpacing(2)
        title = QLabel("Sugestoes / Alertas")
        title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel(
            f"Trocas por urgencia, nao arrumar e sem stock | Horizonte {horizon} dias | {len(rows)} linhas"
        )
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        top.addLayout(title_col, 1)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        top.addWidget(refresh_btn)
        layout.addLayout(top)

        info = QLabel(
            "Aqui ficam apenas os alertas inteligentes. A separacao operacional fica na janela 'Separacao - MP'."
        )
        info.setProperty("role", "muted")
        info.setWordWrap(True)
        layout.addWidget(info)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)

        table_card = CardFrame()
        table_card.set_tone("warning")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(12, 12, 12, 12)
        table_layout.setSpacing(8)

        table = QTableWidget(0, 10)
        table.setHorizontalHeaderLabels(
            ["Prio", "Tema", "Dia", "Turno", "Posto", "Encomenda", "Cliente", "Material", "Recomendacao", "Estado"]
        )
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setAlternatingRowColors(True)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        header.setSectionResizeMode(5, QHeaderView.Interactive)
        header.setSectionResizeMode(6, QHeaderView.Interactive)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        table.setColumnWidth(4, 160)
        table.setColumnWidth(5, 220)
        table.setColumnWidth(6, 160)

        detail = QTextEdit()
        detail.setReadOnly(True)
        detail.setFrameStyle(QFrame.NoFrame)
        detail.setMinimumHeight(180)

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        accept_btn = QPushButton("Validar")
        ignore_btn = QPushButton("Ignorar hoje")
        ignore_btn.setProperty("variant", "secondary")
        clear_btn = QPushButton("Limpar decisao")
        clear_btn.setProperty("variant", "secondary")
        action_row.addWidget(accept_btn)
        action_row.addWidget(ignore_btn)
        action_row.addWidget(clear_btn)
        action_row.addStretch(1)

        def render(current_rows: list[dict]) -> None:
            table.setRowCount(len(current_rows))
            for row_index, row in enumerate(current_rows):
                values = [
                    str(row.get("priority_label", "") or "-"),
                    self._kind_label(str(row.get("kind", "") or "")),
                    str(row.get("planeamento_dia", "") or "-"),
                    str(row.get("planeamento_turno", "") or "-"),
                    str(row.get("posto_trabalho", "") or "-"),
                    str(row.get("numero", "") or "-"),
                    str(row.get("cliente", "") or "-"),
                    f"{row.get('material', '')} {row.get('espessura', '')}".strip(),
                    str(row.get("recommendation", "") or "-"),
                    str(row.get("status_label", "") or "-"),
                ]
                for col_index, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    if col_index in (0, 1, 2, 3, 4, 9):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    if col_index == 0:
                        item.setData(Qt.UserRole, dict(row))
                    table.setItem(row_index, col_index, item)
                _paint_row(
                    table,
                    row_index,
                    str(row.get("priority_tone", "") or "default"),
                    str(row.get("status_key", "") or ""),
                )
            if current_rows:
                table.selectRow(0)
            else:
                detail.setHtml("<b>Sem alertas</b><br>Nao existem sugestoes ativas no horizonte atual.")

        def selected_row() -> dict:
            row_index = _table_selected_row(table)
            if row_index < 0:
                return {}
            item = table.item(row_index, 0)
            return dict(item.data(Qt.UserRole) or {}) if item is not None else {}

        def sync_detail() -> None:
            row = selected_row()
            if not row:
                detail.clear()
                return
            lines = [str(line or "").strip() for line in list(row.get("detail_lines", []) or []) if str(line or "").strip()]
            html = [
                f"<b>{row.get('headline', 'Alerta')}</b><br>",
                f"{row.get('recommendation', '-') or '-'}<br><br>",
                f"<b>Janela:</b> {row.get('planeamento_dia', '-')} | {row.get('planeamento_turno', '-')} | {row.get('planeamento_hora', '-')}<br>",
                f"<b>Posto:</b> {row.get('posto_trabalho', '-')}<br>",
                f"<b>Encomenda:</b> {row.get('numero', '-')}<br>",
                f"<b>Cliente:</b> {row.get('cliente', '-')}<br>",
                f"<b>Material:</b> {row.get('material', '-')} {row.get('espessura', '-')} mm<br>",
                f"<b>Estado:</b> {row.get('status_label', '-')}<br>",
            ]
            if lines:
                html.append("<br><b>Contexto</b><ul>")
                html.extend(f"<li>{line}</li>" for line in lines)
                html.append("</ul>")
            detail.setHtml("".join(html))

        def refresh_dialog() -> None:
            fresh_rows = list(self.backend.material_assistant_alert_rows(horizon_days=horizon) or [])
            subtitle.setText(
                f"Trocas por urgencia, nao arrumar e sem stock | Horizonte {horizon} dias | {len(fresh_rows)} linhas"
            )
            render(fresh_rows)

        def apply_feedback(decision: str) -> None:
            row = selected_row()
            if not row:
                return
            try:
                self.backend.material_assistant_set_feedback(str(row.get("id", "") or "").strip(), decision)
            except Exception as exc:
                QMessageBox.critical(dialog, "Assistente MP", str(exc))
                return
            refresh_dialog()
            self.refresh()

        table.itemSelectionChanged.connect(sync_detail)
        refresh_btn.clicked.connect(refresh_dialog)
        accept_btn.clicked.connect(lambda: apply_feedback("accepted"))
        ignore_btn.clicked.connect(lambda: apply_feedback("ignored"))
        clear_btn.clicked.connect(lambda: apply_feedback("clear"))
        render(rows)
        table_layout.addWidget(table, 1)
        table_layout.addLayout(action_row)
        split.addWidget(table_card)

        detail_card = CardFrame()
        detail_card.set_tone("default")
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(12, 12, 12, 12)
        detail_layout.addWidget(detail)
        split.addWidget(detail_card)
        split.setSizes([980, 400])

        layout.addWidget(split, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _open_separation_dialog(self) -> None:
        horizon = int(self.horizon_combo.currentData() or 5)
        rows = list(self.backend.material_assistant_separation_rows(horizon_days=horizon) or [])
        dialog = QDialog(self)
        dialog.setWindowTitle("Separacao - Materia-Prima")
        dialog.resize(1480, 760)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        top = QHBoxLayout()
        top.setSpacing(8)
        title_col = QVBoxLayout()
        title_col.setContentsMargins(0, 0, 0, 0)
        title_col.setSpacing(2)
        title = QLabel("Separacao - Materia-Prima")
        title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel(
            f"Lista operacional ligada ao planeamento | Horizonte {horizon} dias | {len(rows)} linhas"
        )
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        top.addLayout(title_col, 1)
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        open_pdf_btn = QPushButton("Abrir PDF")
        open_pdf_btn.clicked.connect(lambda: self.backend.material_assistant_open_separation_pdf(horizon_days=horizon))
        top.addWidget(refresh_btn)
        top.addWidget(open_pdf_btn)
        layout.addLayout(top)

        info = QLabel(
            "Folha operacional ordenada por dia, turno, materia-prima e formato quando existirem varios formatos na mesma espessura."
        )
        info.setProperty("role", "muted")
        info.setWordWrap(True)
        layout.addWidget(info)

        card = CardFrame()
        card.set_tone("info")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 12, 12, 12)
        card_layout.setSpacing(8)

        table = QTableWidget(0, 15)
        table.setHorizontalHeaderLabels(
            [
                "Prio",
                "Dia",
                "Turno",
                "Posto",
                "Encomenda",
                "Material",
                "Esp.",
                "Lote",
                "Dimensao",
                "Qtd.",
                "Entrega",
                "Planeado",
                "Acao sugerida",
                "Visto sep.",
                "Visto conf.",
            ]
        )
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setAlternatingRowColors(True)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.Interactive)
        header.setSectionResizeMode(8, QHeaderView.Interactive)
        header.setSectionResizeMode(9, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(10, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(11, QHeaderView.Interactive)
        header.setSectionResizeMode(12, QHeaderView.Stretch)
        header.setSectionResizeMode(13, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(14, QHeaderView.ResizeToContents)
        table.setColumnWidth(2, 90)
        table.setColumnWidth(3, 120)
        table.setColumnWidth(4, 190)
        table.setColumnWidth(7, 120)
        table.setColumnWidth(8, 130)
        table.setColumnWidth(11, 130)

        detail = QTextEdit()
        detail.setReadOnly(True)
        detail.setFrameStyle(QFrame.NoFrame)
        detail.setMinimumHeight(140)
        updating_checks = {"active": False}
        display_rows: list[dict] = []

        def render_rows(current_rows: list[dict]) -> None:
            updating_checks["active"] = True
            display_rows.clear()
            last_group_key = ""
            for row in current_rows:
                group_key = str(row.get("group_key", "") or "").strip()
                if group_key and group_key != last_group_key:
                    display_rows.append(
                        {
                            "_group_header": True,
                            "group_key": group_key,
                            "group_label": str(row.get("group_label", "") or "-").strip() or "-",
                        }
                    )
                    last_group_key = group_key
                display_rows.append(dict(row))

            table.clearSpans()
            table.setRowCount(len(display_rows))
            for row_index, row in enumerate(display_rows):
                if bool(row.get("_group_header")):
                    group_item = QTableWidgetItem(str(row.get("group_label", "") or "-"))
                    group_item.setFlags(Qt.ItemIsEnabled)
                    group_item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
                    table.setItem(row_index, 0, group_item)
                    table.setSpan(row_index, 0, 1, table.columnCount())
                    table.setRowHeight(row_index, 24)
                    for col_index in range(1, table.columnCount()):
                        empty_item = QTableWidgetItem("")
                        empty_item.setFlags(Qt.NoItemFlags)
                        table.setItem(row_index, col_index, empty_item)
                    group_item.setBackground(QColor("#e8eefc"))
                    group_item.setForeground(QColor("#0f172a"))
                    continue

                values = [
                    str(row.get("priority_label", "") or "-"),
                    str(row.get("planeamento_dia", "") or "-"),
                    str(row.get("planeamento_turno", "") or "-"),
                    str(row.get("posto_trabalho", "") or "-"),
                    str(row.get("numero", "") or "-"),
                    str(row.get("material", "") or "-"),
                    str(row.get("espessura", "") or "-"),
                    str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-"),
                    str(row.get("dimensao", "") or "-"),
                    str(row.get("quantidade_label", "") or _fmt_number(row.get("quantidade", 0))),
                    str(row.get("data_entrega", "") or "-"),
                    f"{row.get('planeamento_hora', '-') or '-'} | {row.get('proxima_acao', '-') or '-'}",
                    str(row.get("acao_sugerida", "") or "-"),
                    str(row.get("visto_sep", "[ ]") or "[ ]"),
                    str(row.get("visto_conf", "[ ]") or "[ ]"),
                ]
                for col_index, value in enumerate(values):
                    if col_index in (13, 14):
                        item = QTableWidgetItem("")
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                        is_checked = bool(row.get("visto_sep_checked")) if col_index == 13 else bool(row.get("visto_conf_checked"))
                        item.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
                    else:
                        item = QTableWidgetItem(value)
                    if col_index in (0, 1, 2, 3, 6, 9, 10):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    if col_index == 0:
                        item.setData(Qt.UserRole, dict(row))
                    table.setItem(row_index, col_index, item)
                table.setRowHeight(row_index, 28)
                _paint_row(
                    table,
                    row_index,
                    str(row.get("priority_tone", "") or "default"),
                    str(row.get("status_key", "") or ""),
                )
            first_data_row = next((idx for idx, row in enumerate(display_rows) if not bool(row.get("_group_header"))), -1)
            if first_data_row >= 0:
                table.selectRow(first_data_row)
            else:
                detail.setHtml("<b>Sem linhas de separacao</b><br>Nao existem encomendas dentro do horizonte escolhido.")
            updating_checks["active"] = False

        def sync_detail() -> None:
            row_index = _table_selected_row(table)
            if row_index < 0:
                detail.clear()
                return
            item = table.item(row_index, 0)
            row = dict(item.data(Qt.UserRole) or {}) if item is not None else {}
            if not row:
                detail.clear()
                return
            detail.setHtml(
                "<b>Instrucao de separacao</b><br>"
                f"Preparar <b>{row.get('numero', '-')}</b> para {row.get('cliente', '-')}<br><br>"
                f"<b>Janela:</b> {row.get('planeamento_dia', '-')} | {row.get('planeamento_turno', '-')} | {row.get('planeamento_hora', '-')}<br>"
                f"<b>Agrupamento:</b> {row.get('group_format_label', '-') if row.get('group_format_label') else 'Sem segregacao adicional'}<br>"
                f"<b>Posto:</b> {row.get('posto_trabalho', '-')}<br>"
                f"<b>Material:</b> {row.get('material', '-')} {row.get('espessura', '-')} mm<br>"
                f"<b>Dimensao:</b> {row.get('dimensao', '-')}<br>"
                f"<b>Quantidade a separar:</b> {row.get('quantidade_label', '-') or '-'}<br>"
                f"<b>Necessidade:</b> {row.get('necessidade_label', _fmt_number(row.get('quantidade_necessaria', row.get('quantidade', 0))))}<br>"
                f"<b>Lote sugerido:</b> {row.get('lote_sugerido', row.get('lote_atual', '-'))}<br>"
                f"<b>Opcoes MP:</b> {row.get('opcoes_mp', '-') or '-'}<br>"
                f"<b>Entrega:</b> {row.get('data_entrega', '-')}<br>"
                f"<b>Planeado:</b> {row.get('proxima_acao', '-')}<br>"
                f"<b>Acao sugerida:</b> {row.get('acao_sugerida', '-')}<br>"
                f"<b>Alerta adicional:</b> {row.get('alerta_texto', '-') if row.get('alerta_retalho') else 'Sem alerta extra'}<br>"
                f"<b>Vistos:</b> {row.get('visto_sep', '[ ]')} Separado | {row.get('visto_conf', '[ ]')} Conferido"
            )

        def refresh_dialog() -> None:
            fresh_rows = list(self.backend.material_assistant_separation_rows(horizon_days=horizon) or [])
            subtitle.setText(
                f"Lista operacional ligada ao planeamento | Horizonte {horizon} dias | {len(fresh_rows)} linhas"
            )
            render_rows(fresh_rows)

        def on_item_changed(item: QTableWidgetItem) -> None:
            if updating_checks["active"]:
                return
            row_index = item.row()
            if row_index < 0:
                return
            row_item = table.item(row_index, 0)
            row = dict(row_item.data(Qt.UserRole) or {}) if row_item is not None else {}
            if not row:
                return
            target_key = str(row.get("check_key", "") or row.get("need_key", "") or "").strip()
            if item.column() == 13:
                self.backend.material_assistant_set_check(target_key, "sep", item.checkState() == Qt.Checked)
            elif item.column() == 14:
                self.backend.material_assistant_set_check(target_key, "conf", item.checkState() == Qt.Checked)
            else:
                return
            sync_detail()

        table.itemSelectionChanged.connect(sync_detail)
        table.itemChanged.connect(on_item_changed)
        refresh_btn.clicked.connect(refresh_dialog)
        render_rows(rows)
        card_layout.addWidget(table, 1)
        card_layout.addWidget(detail)
        layout.addWidget(card, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _open_order_dialog(self) -> None:
        numero, _material, _espessura, _need_key = self._selected_context()
        if not numero:
            return
        try:
            detail = dict(self.backend.order_detail(numero) or {})
        except Exception as exc:
            QMessageBox.critical(self, "Assistente MP", str(exc))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Encomenda {numero}")
        dialog.resize(1080, 680)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel(f"{numero} | {detail.get('cliente', '-')} - {detail.get('cliente_nome', '')}")
        title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        layout.addWidget(title)
        meta = QLabel(
            " | ".join(
                value
                for value in (
                    f"Estado {detail.get('estado', '-')}",
                    f"Entrega {detail.get('data_entrega', '-') or '-'}",
                    f"Montagem {detail.get('montagem_estado', '-')}",
                    f"Reservas {len(list(detail.get('reservas', []) or []))}",
                )
                if str(value).strip()
            )
        )
        meta.setProperty("role", "muted")
        meta.setWordWrap(True)
        layout.addWidget(meta)

        split = QSplitter(Qt.Vertical)
        split.setChildrenCollapsible(False)

        pieces_card = CardFrame()
        pieces_card.set_tone("info")
        pieces_layout = QVBoxLayout(pieces_card)
        pieces_layout.setContentsMargins(12, 12, 12, 12)
        pieces_layout.setSpacing(8)
        pieces_layout.addWidget(QLabel("Pecas"))
        pieces_table = QTableWidget(0, 6)
        pieces_table.setHorizontalHeaderLabels(["Ref. Int.", "Ref. Ext.", "Descricao", "Material", "Qtd plan.", "Estado"])
        pieces_table.verticalHeader().setVisible(False)
        pieces_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        pieces_header = pieces_table.horizontalHeader()
        pieces_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        pieces_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        pieces_header.setSectionResizeMode(2, QHeaderView.Stretch)
        pieces_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        pieces_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        pieces_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        piece_rows = list(detail.get("pieces", []) or [])
        pieces_table.setRowCount(len(piece_rows))
        for row_index, row in enumerate(piece_rows):
            values = [
                str(row.get("ref_interna", "") or "").strip(),
                str(row.get("ref_externa", "") or "").strip(),
                str(row.get("descricao", "") or "").strip(),
                f"{row.get('material', '')} {row.get('espessura', '')}".strip(),
                str(row.get("qtd_plan", "") or "-"),
                str(row.get("estado", "") or "-"),
            ]
            for col_index, value in enumerate(values):
                pieces_table.setItem(row_index, col_index, QTableWidgetItem(value))
        pieces_layout.addWidget(pieces_table)
        split.addWidget(pieces_card)

        reserve_card = CardFrame()
        reserve_card.set_tone("warning")
        reserve_layout = QVBoxLayout(reserve_card)
        reserve_layout.setContentsMargins(12, 12, 12, 12)
        reserve_layout.setSpacing(8)
        reserve_layout.addWidget(QLabel("Reservas / materia-prima"))
        reserve_table = QTableWidget(0, 4)
        reserve_table.setHorizontalHeaderLabels(["Material", "Esp.", "Quantidade", "Material ID"])
        reserve_table.verticalHeader().setVisible(False)
        reserve_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        reserve_header = reserve_table.horizontalHeader()
        reserve_header.setSectionResizeMode(0, QHeaderView.Stretch)
        reserve_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        reserve_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        reserve_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        reserve_rows = list(detail.get("reservas", []) or [])
        reserve_table.setRowCount(len(reserve_rows))
        for row_index, row in enumerate(reserve_rows):
            values = [
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                str(row.get("quantidade", "") or "-"),
                str(row.get("material_id", "") or "").strip(),
            ]
            for col_index, value in enumerate(values):
                reserve_table.setItem(row_index, col_index, QTableWidgetItem(value))
        reserve_layout.addWidget(reserve_table)
        split.addWidget(reserve_card)
        split.setSizes([420, 220])

        layout.addWidget(split, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()
