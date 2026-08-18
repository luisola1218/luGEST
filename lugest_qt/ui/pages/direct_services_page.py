from __future__ import annotations

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, ClickableDateEdit as QDateEdit, StatCard
from .runtime_common import configure_table as _configure_table, fmt_eur as _fmt_eur


_KIND_LABELS = {
    "set": "Conjunto",
    "product": "Produto",
    "material": "Matéria-prima",
    "service": "Serviço",
}


class _CatalogDialog(QDialog):
    def __init__(self, backend, kind: str, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.kind = kind
        self.selected: dict = {}
        self.setWindowTitle(f"Adicionar {_KIND_LABELS.get(kind, 'artigo').lower()}")
        self.resize(1160 if kind == "material" else 920, 640 if kind == "material" else 600)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        title = QLabel(f"Selecionar {_KIND_LABELS.get(kind, 'artigo')}")
        title.setStyleSheet("font-size: 19px; font-weight: 900; color: #172b3f;")
        hint = QLabel("Pesquisa no catálogo existente; o stock só é consumido quando o serviço for confirmado.")
        hint.setProperty("role", "muted")
        self.search = QLineEdit()
        self.search.setPlaceholderText("Pesquisar referência ou descrição...")
        self.search.textChanged.connect(self._refresh)
        if kind == "material":
            self.table = QTableWidget(0, 8)
            self.table.setHorizontalHeaderLabels(
                ["Referência", "Material", "Formato", "Dimensões", "Esp./Perfil", "Lote", "Preço", "Disponível"]
            )
        else:
            self.table = QTableWidget(0, 5)
            self.table.setHorizontalHeaderLabels(["Referência", "Descrição", "Unidade", "Preço", "Disponível"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setMinimumWidth(0)
        self.table.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        if kind == "material":
            _configure_table(self.table, stretch=(1,), contents=(0, 2, 3, 4, 5, 6, 7))
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            for column, width in ((0, 108), (2, 92), (3, 154), (4, 108), (5, 148), (6, 96), (7, 88)):
                header.setSectionResizeMode(column, QHeaderView.Interactive)
                header.resizeSection(column, width)
        else:
            _configure_table(self.table, stretch=(1,), contents=(0, 2, 3, 4))
            self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.itemDoubleClicked.connect(lambda *_: self._accept())
        buttons = QDialogButtonBox(QDialogButtonBox.Cancel)
        add_btn = buttons.addButton("Adicionar ao serviço", QDialogButtonBox.AcceptRole)
        add_btn.setProperty("variant", "success")
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addWidget(self.search)
        layout.addWidget(self.table, 1)
        layout.addWidget(buttons)
        self._rows: list[dict] = []
        self._refresh()

    def _refresh(self) -> None:
        try:
            self._rows = list(self.backend.direct_service_catalog(self.kind, self.search.text()) or [])
        except Exception as exc:
            QMessageBox.warning(self, "Catálogo", str(exc))
            self._rows = []
        self.table.setRowCount(len(self._rows))
        for row_index, row in enumerate(self._rows):
            if self.kind == "material":
                values = (
                    row.get("ref", ""), row.get("material", "-"), row.get("format", "-"),
                    row.get("dimensions", "-"), row.get("specification", "-"), row.get("lot", "-"),
                    _fmt_eur(row.get("unit_price", 0)), row.get("available", "-"),
                )
                numeric_columns = (6, 7)
            else:
                values = (
                    row.get("ref", ""), row.get("description", ""), row.get("unit", "UN"),
                    _fmt_eur(row.get("unit_price", 0)), row.get("available", "-"),
                )
                numeric_columns = (3, 4)
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if col in numeric_columns:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_index, col, item)
        if self._rows:
            self.table.selectRow(0)

    def _accept(self) -> None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._rows):
            QMessageBox.information(self, "Selecionar", "Seleciona uma linha do catálogo.")
            return
        self.selected = dict(self._rows[row])
        self.accept()


class DirectServicesPage(QWidget):
    page_title = "Serviços"
    page_subtitle = "Venda direta de conjuntos, artigos, matéria-prima e mão de obra, sem passar pela produção."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current: dict = {}
        self._loading = False

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.stack = QStackedWidget()
        self.stack.setMinimumWidth(0)
        self.stack.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        root.addWidget(self.stack)
        self._build_list_page()
        self._build_detail_page()
        self.stack.addWidget(self.list_page)
        self.stack.addWidget(self.detail_page)

    def _build_list_page(self) -> None:
        self.list_page = QWidget()
        self.list_page.setMinimumWidth(0)
        self.list_page.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        layout = QVBoxLayout(self.list_page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        intro = CardFrame()
        intro.setProperty("tone", "info")
        intro_layout = QGridLayout(intro)
        intro_layout.setContentsMargins(16, 12, 16, 12)
        intro_layout.setHorizontalSpacing(10)
        title = QLabel("Balcão de serviços")
        title.setStyleSheet("font-size: 20px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel("Cliente → artigos/conjuntos/mão de obra → confirmação do stock → faturação.")
        subtitle.setProperty("role", "muted")
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar serviço, cliente, local ou faturação...")
        self.filter_edit.textChanged.connect(self.refresh)
        self.state_combo = QComboBox()
        self.state_combo.addItems(["Todos", "Rascunho", "Confirmado", "Faturado", "Anulado"])
        self.state_combo.currentTextChanged.connect(self.refresh)
        new_btn = QPushButton("Novo serviço")
        new_btn.setProperty("variant", "success")
        new_btn.clicked.connect(self._new_service)
        open_btn = QPushButton("Abrir")
        open_btn.clicked.connect(self._open_selected)
        intro_layout.addWidget(title, 0, 0, 1, 3)
        intro_layout.addWidget(subtitle, 1, 0, 1, 3)
        intro_layout.addWidget(self.filter_edit, 2, 0)
        intro_layout.addWidget(self.state_combo, 2, 1)
        actions = QWidget()
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)
        actions_layout.addWidget(new_btn)
        actions_layout.addWidget(open_btn)
        intro_layout.addWidget(actions, 2, 2)
        intro_layout.setColumnStretch(0, 5)
        intro_layout.setColumnStretch(1, 2)
        intro_layout.setColumnStretch(2, 2)
        layout.addWidget(intro)

        cards_host = QWidget()
        cards_layout = QHBoxLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setSpacing(10)
        self.cards = [StatCard(title) for title in ("Rascunhos", "Prontos a faturar", "Faturados", "Valor ativo")]
        for card in self.cards:
            cards_layout.addWidget(card)
        layout.addWidget(cards_host)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(16, 14, 16, 14)
        table_head = QHBoxLayout()
        table_head.setContentsMargins(0, 0, 0, 0)
        table_title = QLabel("Serviços diretos")
        table_title.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        self.list_count_label = QLabel("0 serviços")
        self.list_count_label.setProperty("role", "muted")
        table_head.addWidget(table_title)
        table_head.addStretch(1)
        table_head.addWidget(self.list_count_label)
        self.list_table = QTableWidget(0, 7)
        self.list_table.setHorizontalHeaderLabels(["Número", "Data", "Cliente", "Estado", "Linhas", "Total", "Faturação"])
        self.list_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.list_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.list_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_table.setMinimumWidth(0)
        self.list_table.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.list_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        _configure_table(self.list_table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6))
        self.list_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.list_table.itemDoubleClicked.connect(lambda *_: self._open_selected())
        table_layout.addLayout(table_head)
        table_layout.addWidget(self.list_table, 1)
        layout.addWidget(table_card, 1)

    def _build_detail_page(self) -> None:
        self.detail_page = QWidget()
        self.detail_page.setMinimumWidth(0)
        self.detail_page.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        layout = QVBoxLayout(self.detail_page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        head = CardFrame()
        head_layout = QHBoxLayout(head)
        head_layout.setContentsMargins(14, 10, 14, 10)
        head_layout.setSpacing(10)
        back = QPushButton("Voltar à lista")
        back.setProperty("variant", "secondary")
        back.clicked.connect(self._show_list)
        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        self.number_label = QLabel("Novo serviço")
        self.number_label.setStyleSheet("font-size: 19px; font-weight: 900; color: #0f172a;")
        self.context_label = QLabel("Venda direta sem passagem pela produção")
        self.context_label.setProperty("role", "muted")
        title_col.addWidget(self.number_label)
        title_col.addWidget(self.context_label)
        self.state_label = QLabel("Rascunho")
        self.state_label.setProperty("role", "state_chip")
        self.save_btn = QPushButton("Guardar rascunho")
        self.save_btn.setProperty("variant", "success")
        self.save_btn.clicked.connect(self._save)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove)
        head_layout.addWidget(back)
        head_layout.addLayout(title_col, 1)
        head_layout.addWidget(self.state_label)
        head_layout.addWidget(self.save_btn)
        head_layout.addWidget(self.remove_btn)
        layout.addWidget(head)

        flow = CardFrame()
        flow_layout = QHBoxLayout(flow)
        flow_layout.setContentsMargins(14, 8, 14, 8)
        flow_layout.setSpacing(6)
        self.flow_labels: list[QLabel] = []
        for text in ("1  Preparar", "2  Confirmar e consumir stock", "3  Faturar"):
            label = QLabel(text)
            label.setAlignment(Qt.AlignCenter)
            label.setMinimumHeight(36)
            label.setStyleSheet("border: 1px solid #ccd3da; padding: 7px; font-weight: 800; background: #ffffff;")
            flow_layout.addWidget(label, 1)
            self.flow_labels.append(label)
        layout.addWidget(flow)

        info = CardFrame()
        info_layout = QGridLayout(info)
        info_layout.setContentsMargins(16, 12, 16, 12)
        info_layout.setHorizontalSpacing(10)
        info_layout.setVerticalSpacing(5)
        info_title = QLabel("Dados do serviço")
        info_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        info_hint = QLabel("Identifica quem pediu o trabalho, onde será realizado e quando deve ser cobrado.")
        info_hint.setProperty("role", "muted")
        self.client_combo = QComboBox()
        self.client_combo.setEditable(True)
        self.client_combo.setInsertPolicy(QComboBox.NoInsert)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.due_edit = QDateEdit()
        self.due_edit.setCalendarPopup(True)
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("Morada ou local do serviço")
        self.responsible_edit = QLineEdit()
        self.obs_edit = QTextEdit()
        self.obs_edit.setMaximumHeight(62)
        info_layout.addWidget(info_title, 0, 0, 1, 4)
        info_layout.addWidget(info_hint, 1, 0, 1, 4)
        for col, text in enumerate(("Cliente", "Data do serviço", "Vencimento", "Responsável")):
            info_layout.addWidget(QLabel(text), 2, col)
        info_layout.addWidget(self.client_combo, 3, 0)
        info_layout.addWidget(self.date_edit, 3, 1)
        info_layout.addWidget(self.due_edit, 3, 2)
        info_layout.addWidget(self.responsible_edit, 3, 3)
        info_layout.addWidget(QLabel("Local do serviço"), 4, 0)
        info_layout.addWidget(QLabel("Observações"), 4, 2)
        info_layout.addWidget(self.location_edit, 5, 0, 1, 2)
        info_layout.addWidget(self.obs_edit, 5, 2, 1, 2)
        for col in range(4):
            info_layout.setColumnStretch(col, 1)
        layout.addWidget(info)

        lines_card = CardFrame()
        lines_layout = QVBoxLayout(lines_card)
        lines_layout.setContentsMargins(14, 12, 14, 12)
        lines_layout.setSpacing(8)
        toolbar = QVBoxLayout()
        toolbar.setSpacing(8)
        toolbar_head = QHBoxLayout()
        line_title = QLabel("Trabalho e artigos")
        line_title.setStyleSheet("font-size: 17px; font-weight: 900;")
        self.line_count_label = QLabel("0 linhas")
        self.line_count_label.setProperty("role", "muted")
        toolbar_head.addWidget(line_title)
        toolbar_head.addStretch(1)
        toolbar_head.addWidget(self.line_count_label)
        toolbar.addLayout(toolbar_head)
        toolbar_actions = QHBoxLayout()
        toolbar_actions.setSpacing(6)
        self.add_buttons: list[QPushButton] = []
        for text, kind in (("Conjunto", "set"), ("Produto", "product"), ("Matéria-prima", "material")):
            button = QPushButton(f"+ {text}")
            button.clicked.connect(lambda _checked=False, current_kind=kind: self._add_catalog_line(current_kind))
            toolbar_actions.addWidget(button)
            self.add_buttons.append(button)
        manual = QPushButton("+ Mão de obra / serviço")
        manual.clicked.connect(self._add_manual_line)
        toolbar_actions.addWidget(manual)
        self.add_buttons.append(manual)
        toolbar_actions.addStretch(1)
        self.remove_line_btn = QPushButton("Remover linha")
        self.remove_line_btn.setProperty("variant", "danger")
        self.remove_line_btn.clicked.connect(self._remove_line)
        toolbar_actions.addWidget(self.remove_line_btn)
        toolbar.addLayout(toolbar_actions)
        self.lines_table = QTableWidget(0, 9)
        self.lines_table.setHorizontalHeaderLabels(["Tipo", "Referência", "Descrição", "Qtd.", "Unid.", "Preço unit.", "IVA %", "Subtotal", "Total"])
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.lines_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.lines_table.setMinimumWidth(0)
        self.lines_table.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.lines_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        _configure_table(self.lines_table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6, 7, 8))
        self.lines_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.lines_table.itemChanged.connect(self._line_changed)
        lines_layout.addLayout(toolbar)
        lines_layout.addWidget(self.lines_table, 1)
        layout.addWidget(lines_card, 1)

        footer = CardFrame()
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(16, 10, 16, 10)
        footer_layout.setSpacing(12)
        totals_host = QWidget()
        totals_layout = QHBoxLayout(totals_host)
        totals_layout.setContentsMargins(0, 0, 0, 0)
        totals_layout.setSpacing(8)
        self.summary_values: dict[str, QLabel] = {}
        for key, caption in (("subtotal", "Base sem IVA"), ("tax", "IVA"), ("total", "Total do serviço")):
            total_card = CardFrame()
            total_card.setProperty("tone", "success" if key == "total" else "default")
            total_card_layout = QVBoxLayout(total_card)
            total_card_layout.setContentsMargins(12, 7, 12, 7)
            total_card_layout.setSpacing(1)
            caption_label = QLabel(caption)
            caption_label.setProperty("role", "muted")
            value_label = QLabel("0,00 EUR")
            value_label.setStyleSheet(
                "font-size: 17px; font-weight: 900; color: #0f5132;" if key == "total"
                else "font-size: 15px; font-weight: 800; color: #172b3f;"
            )
            total_card_layout.addWidget(caption_label)
            total_card_layout.addWidget(value_label)
            totals_layout.addWidget(total_card, 1)
            self.summary_values[key] = value_label
        self.confirm_btn = QPushButton("Confirmar e consumir stock")
        self.confirm_btn.setProperty("variant", "success")
        self.confirm_btn.clicked.connect(self._confirm)
        self.billing_btn = QPushButton("Enviar para Faturação")
        self.billing_btn.setProperty("variant", "success")
        self.billing_btn.clicked.connect(self._send_to_billing)
        self.cancel_btn = QPushButton("Anular serviço")
        self.cancel_btn.setProperty("variant", "danger")
        self.cancel_btn.clicked.connect(self._cancel)
        footer_actions = QHBoxLayout()
        footer_actions.setSpacing(6)
        footer_actions.addWidget(self.confirm_btn)
        footer_actions.addWidget(self.billing_btn)
        footer_actions.addWidget(self.cancel_btn)
        footer_layout.addWidget(totals_host, 1)
        footer_layout.addLayout(footer_actions)
        layout.addWidget(footer)

    def refresh(self, *_args, **_kwargs) -> None:
        try:
            dashboard = dict(self.backend.direct_service_dashboard() or {})
            self.rows = list(self.backend.direct_service_rows(self.filter_edit.text(), self.state_combo.currentText()) or [])
        except Exception as exc:
            QMessageBox.warning(self, "Serviços", str(exc))
            return
        values = (
            dashboard.get("draft_count", 0), dashboard.get("pending_billing_count", 0),
            dashboard.get("billed_count", 0), _fmt_eur(dashboard.get("active_total", 0)),
        )
        for card, value in zip(self.cards, values):
            card.set_data(str(value))
        for card, tone in zip(self.cards, ("default", "warning", "success", "info")):
            card.set_tone(tone)
        self.list_count_label.setText(f"{len(self.rows)} serviço(s)")
        self.list_table.setRowCount(len(self.rows))
        for row_index, row in enumerate(self.rows):
            values = (
                row.get("numero", ""), row.get("data_servico", ""), row.get("cliente", ""),
                row.get("estado", ""), row.get("linhas", 0), _fmt_eur(row.get("total", 0)),
                row.get("faturacao_numero", "") or "-",
            )
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col in (4, 5):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.list_table.setItem(row_index, col, item)

    def _load_clients(self, selected: str = "") -> None:
        self.client_combo.clear()
        self.client_combo.addItem("Selecionar cliente", "")
        for row in list(self.backend.client_rows("") or []):
            code = str(row.get("codigo", "") or "")
            self.client_combo.addItem(f"{code} - {row.get('nome', '')}", code)
        index = self.client_combo.findData(selected)
        self.client_combo.setCurrentIndex(index if index >= 0 else 0)

    def _new_service(self) -> None:
        try:
            detail = self.backend.direct_service_create()
        except Exception as exc:
            QMessageBox.warning(self, "Novo serviço", str(exc))
            return
        self._load_detail(detail)

    def _open_selected(self) -> None:
        row = self.list_table.currentRow()
        if row < 0 or row >= len(self.rows):
            QMessageBox.information(self, "Abrir", "Seleciona um serviço.")
            return
        try:
            self._load_detail(self.backend.direct_service_detail(self.rows[row]["numero"]))
        except Exception as exc:
            QMessageBox.warning(self, "Abrir", str(exc))

    def _load_detail(self, detail: dict) -> None:
        self._loading = True
        self.current = dict(detail or {})
        self.number_label.setText(f"{self.current.get('numero', '')}  ·  {self.current.get('cliente_nome', '') or 'Cliente por selecionar'}")
        billing_number = str(self.current.get("faturacao_numero", "") or "").strip()
        self.context_label.setText(
            f"Registo de faturação {billing_number}" if billing_number
            else "Venda direta sem passagem pela produção"
        )
        self.state_label.setText(str(self.current.get("estado", "Rascunho")))
        self._load_clients(str(self.current.get("cliente_codigo", "") or ""))
        service_date = QDate.fromString(str(self.current.get("data_servico", "")), "yyyy-MM-dd")
        due_date = QDate.fromString(str(self.current.get("data_vencimento", "")), "yyyy-MM-dd")
        self.date_edit.setDate(service_date if service_date.isValid() else QDate.currentDate())
        self.due_edit.setDate(due_date if due_date.isValid() else QDate.currentDate())
        self.location_edit.setText(str(self.current.get("local_servico", "") or ""))
        self.responsible_edit.setText(str(self.current.get("responsavel", "") or ""))
        self.obs_edit.setPlainText(str(self.current.get("obs", "") or ""))
        self._fill_lines(list(self.current.get("linhas", []) or []))
        self._loading = False
        self._update_detail_state()
        self.stack.setCurrentWidget(self.detail_page)

    def _fill_lines(self, lines: list[dict]) -> None:
        self.lines_table.setRowCount(len(lines))
        self.line_count_label.setText(f"{len(lines)} linha(s)")
        for row_index, row in enumerate(lines):
            values = (
                _KIND_LABELS.get(str(row.get("kind", "")), "Serviço"), row.get("ref", ""),
                row.get("description", ""), row.get("qty", 1), row.get("unit", "UN"),
                row.get("unit_price", 0), row.get("iva_perc", 23), row.get("subtotal", 0), row.get("total", 0),
            )
            for col, value in enumerate(values):
                text = _fmt_eur(value) if col in (5, 7, 8) else str(value)
                item = QTableWidgetItem(text)
                item.setData(Qt.UserRole, dict(row))
                if col not in (2, 3, 4, 5, 6):
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.lines_table.setItem(row_index, col, item)

    def _collect_lines(self) -> list[dict]:
        result: list[dict] = []
        for row in range(self.lines_table.rowCount()):
            base = dict((self.lines_table.item(row, 0) or QTableWidgetItem()).data(Qt.UserRole) or {})
            def txt(col: int) -> str:
                return str((self.lines_table.item(row, col) or QTableWidgetItem()).text() or "").strip()
            def number(col: int, default: float = 0.0) -> float:
                raw = txt(col).replace("EUR", "").replace(" ", "")
                if "," in raw:
                    raw = raw.replace(".", "").replace(",", ".")
                try:
                    return float(raw)
                except Exception:
                    return default
            base.update({"description": txt(2), "qty": number(3, 1), "unit": txt(4), "unit_price": number(5), "iva_perc": number(6, 23)})
            result.append(base)
        return result

    def _payload(self) -> dict:
        return {
            "numero": self.current.get("numero", ""),
            "cliente_codigo": self.client_combo.currentData() or "",
            "data_servico": self.date_edit.date().toString("yyyy-MM-dd"),
            "data_vencimento": self.due_edit.date().toString("yyyy-MM-dd"),
            "local_servico": self.location_edit.text(),
            "responsavel": self.responsible_edit.text(),
            "obs": self.obs_edit.toPlainText(),
            "linhas": self._collect_lines(),
        }

    def _save(self, quiet: bool = False) -> bool:
        try:
            self.current = dict(self.backend.direct_service_save(self._payload()) or {})
        except Exception as exc:
            QMessageBox.warning(self, "Guardar serviço", str(exc))
            return False
        self._load_detail(self.current)
        if not quiet:
            QMessageBox.information(self, "Serviço", "Rascunho guardado.")
        return True

    def _add_catalog_line(self, kind: str) -> None:
        dialog = _CatalogDialog(self.backend, kind, self)
        if dialog.exec() != QDialog.Accepted:
            return
        row = dict(dialog.selected)
        row.update({"qty": 1, "iva_perc": 23})
        self._append_line(row)

    def _add_manual_line(self) -> None:
        description, ok = QInputDialog.getText(self, "Mão de obra / serviço", "Descrição do trabalho:")
        if not ok or not str(description or "").strip():
            return
        price, ok = QInputDialog.getDouble(self, "Preço", "Preço unitário sem IVA:", 0, 0, 9999999, 2)
        if not ok:
            return
        self._append_line({"kind": "service", "ref": "", "description": description, "qty": 1, "unit": "SV", "unit_price": price, "iva_perc": 23})

    def _append_line(self, row: dict) -> None:
        lines = self._collect_lines()
        lines.append(dict(row))
        self._loading = True
        self._fill_lines(lines)
        self._loading = False
        self._recalculate_preview()

    def _remove_line(self) -> None:
        row = self.lines_table.currentRow()
        if row >= 0:
            self.lines_table.removeRow(row)
            self._recalculate_preview()

    def _line_changed(self, *_args) -> None:
        if not self._loading:
            self._recalculate_preview()

    def _recalculate_preview(self) -> None:
        subtotal = iva = 0.0
        for row in self._collect_lines():
            value = float(row.get("qty", 0) or 0) * float(row.get("unit_price", 0) or 0)
            subtotal += value
            iva += value * float(row.get("iva_perc", 0) or 0) / 100.0
        self.summary_values["subtotal"].setText(_fmt_eur(subtotal))
        self.summary_values["tax"].setText(_fmt_eur(iva))
        self.summary_values["total"].setText(_fmt_eur(subtotal + iva))

    def _confirm(self) -> None:
        if not self._save(quiet=True):
            return
        answer = QMessageBox.question(self, "Confirmar serviço", "Confirmar o serviço e consumir agora o stock associado? Esta operação fecha o documento.")
        if answer != QMessageBox.Yes:
            return
        try:
            self._load_detail(self.backend.direct_service_confirm(self.current["numero"]))
        except Exception as exc:
            QMessageBox.warning(self, "Confirmar serviço", str(exc))

    def _send_to_billing(self) -> None:
        try:
            result = dict(self.backend.direct_service_send_to_billing(self.current["numero"]) or {})
            self._load_detail(dict(result.get("service", {}) or {}))
            number = str((result.get("billing", {}) or {}).get("numero", "") or "")
            QMessageBox.information(self, "Faturação", f"Serviço enviado para Faturação no registo {number}.")
        except Exception as exc:
            QMessageBox.warning(self, "Faturação", str(exc))

    def _cancel(self) -> None:
        state = str(self.current.get("estado", "") or "").strip().casefold()
        if state == "confirmado":
            warning = (
                "O stock deste serviço já foi consumido e não será reposto automaticamente. "
                "A anulação fecha o documento, mas qualquer devolução deve ser regularizada "
                "pelos movimentos de stock. Pretende continuar?"
            )
            if QMessageBox.warning(
                self,
                "Anular serviço confirmado",
                warning,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            ) != QMessageBox.Yes:
                return
        reason, ok = QInputDialog.getText(self, "Anular serviço", "Motivo da anulação:")
        if not ok:
            return
        reason = reason.strip()
        if not reason:
            QMessageBox.information(self, "Anular serviço", "Indique o motivo da anulação para manter o histórico auditável.")
            return
        try:
            self._load_detail(self.backend.direct_service_cancel(self.current["numero"], reason))
        except Exception as exc:
            QMessageBox.warning(self, "Anular serviço", str(exc))

    def _remove(self) -> None:
        if QMessageBox.question(self, "Remover", "Remover definitivamente este rascunho?") != QMessageBox.Yes:
            return
        try:
            self.backend.direct_service_remove(self.current["numero"])
        except Exception as exc:
            QMessageBox.warning(self, "Remover", str(exc))
            return
        self._show_list()

    def _update_detail_state(self) -> None:
        state = str(self.current.get("estado", "Rascunho") or "Rascunho")
        editable = state == "Rascunho"
        for widget in (self.client_combo, self.date_edit, self.due_edit, self.location_edit, self.responsible_edit, self.obs_edit, self.lines_table):
            widget.setEnabled(editable)
        for button in self.add_buttons + [self.remove_line_btn]:
            button.setEnabled(editable)
        self.save_btn.setEnabled(editable)
        self.remove_btn.setEnabled(editable)
        self.confirm_btn.setEnabled(editable)
        self.billing_btn.setEnabled(state == "Confirmado" or (state == "Faturado" and not self.current.get("faturacao_numero")))
        self.cancel_btn.setEnabled(state in {"Rascunho", "Confirmado"})
        active_index = 0 if state == "Rascunho" else (1 if state == "Confirmado" else 2)
        for index, label in enumerate(self.flow_labels):
            if index <= active_index:
                label.setStyleSheet("border: 1px solid #72c900; padding: 7px; font-weight: 900; background: #eef9df; color: #305c00;")
            else:
                label.setStyleSheet("border: 1px solid #ccd3da; padding: 7px; font-weight: 800; background: #ffffff;")
        self._recalculate_preview()

    def _show_list(self) -> None:
        self.current = {}
        self.stack.setCurrentWidget(self.list_page)
        self.refresh()
