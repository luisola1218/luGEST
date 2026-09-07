from __future__ import annotations
import unicodedata
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)
from .purchase_notes_page import PurchaseNotesPage
from .runtime_common import (
    apply_state_chip as _apply_state_chip,
    fill_table as _fill_table,
    paint_table_row as _paint_table_row,
)
from .runtime_support import _adopt_layout_item, _fmt_eur, _take_layout_items
from ..widgets import CardFrame


class LegacyPurchaseNotesPage(PurchaseNotesPage):
    page_subtitle = "Carteira de aprovisionamento, cotações, encomendas a fornecedor e receções."

    def __init__(self, backend, parent=None) -> None:
        super().__init__(backend, parent)
        root = self.layout()
        sections = _take_layout_items(root)
        filters_item = sections[0] if len(sections) > 0 else None
        notes_item = sections[1] if len(sections) > 1 else None
        form_item = sections[2] if len(sections) > 2 else None

        self.view_stack = QStackedWidget()
        self.list_page = QWidget()
        list_layout = QVBoxLayout(self.list_page)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(14)
        _adopt_layout_item(list_layout, filters_item)
        list_actions = CardFrame()
        list_actions.set_tone("default")
        list_actions_layout = QHBoxLayout(list_actions)
        list_actions_layout.setContentsMargins(10, 7, 10, 7)
        list_actions_layout.setSpacing(6)
        self.open_note_btn = QPushButton("Abrir nota")
        self.open_note_btn.clicked.connect(self._open_selected_note)
        create_note_btn = QPushButton("Criar nota")
        create_note_btn.clicked.connect(self._create_new_note)
        self.remove_note_list_btn = QPushButton("Apagar nota")
        self.remove_note_list_btn.setProperty("variant", "destructive")
        self.remove_note_list_btn.clicked.connect(self._remove_note)
        refresh_note_btn = QPushButton("Atualizar")
        refresh_note_btn.setProperty("variant", "secondary")
        refresh_note_btn.clicked.connect(self.refresh)
        for button, width in (
            (create_note_btn, 92),
            (self.open_note_btn, 92),
            (self.remove_note_list_btn, 92),
            (refresh_note_btn, 82),
        ):
            button.setProperty("compact", "true")
            button.setFixedWidth(width)
            button.setMinimumHeight(29)
            button.setMaximumHeight(31)
            button.setStyleSheet("font-family: 'Segoe UI'; font-size: 10px; font-weight: 700;")
        list_actions_layout.addWidget(create_note_btn)
        list_actions_layout.addWidget(self.open_note_btn)
        list_actions_layout.addWidget(self.remove_note_list_btn)
        list_actions_layout.addWidget(refresh_note_btn)
        list_actions_layout.addStretch(1)
        list_actions.setMaximumHeight(48)
        list_layout.addWidget(list_actions)

        notes_widget = notes_item.widget() if notes_item is not None else None
        self.note_list_inspector = CardFrame()
        self.note_list_inspector.set_tone("default")
        self.note_list_inspector.setMinimumWidth(360)
        self.note_list_inspector.setMaximumWidth(470)
        self.note_list_inspector.setStyleSheet(
            "QFrame#PurchaseListSummary { background: #f6f9fc; border: 1px solid #d7e2ee; }"
            "QLabel#PurchaseListEyebrow { color: #5b7088; font-size: 8px; font-weight: 700; }"
            "QLabel#PurchaseListTitle { color: #0f172a; font-size: 16px; font-weight: 800; }"
            "QLabel#PurchaseListMetricLabel { color: #5b7088; font-size: 8px; font-weight: 700; }"
            "QLabel#PurchaseListMetricValue { color: #10253d; font-size: 13px; font-weight: 800; }"
        )
        inspector_layout = QVBoxLayout(self.note_list_inspector)
        inspector_layout.setContentsMargins(12, 12, 12, 12)
        inspector_layout.setSpacing(9)
        inspector_header = QHBoxLayout()
        inspector_heading = QVBoxLayout()
        inspector_heading.setContentsMargins(0, 0, 0, 0)
        inspector_heading.setSpacing(1)
        inspector_eyebrow = QLabel("DOCUMENTO SELECIONADO")
        inspector_eyebrow.setObjectName("PurchaseListEyebrow")
        self.note_inspector_number = QLabel("Sem seleção")
        self.note_inspector_number.setObjectName("PurchaseListTitle")
        self.note_inspector_number.setWordWrap(True)
        inspector_heading.addWidget(inspector_eyebrow)
        inspector_heading.addWidget(self.note_inspector_number)
        self.note_inspector_state = QLabel("-")
        _apply_state_chip(self.note_inspector_state, "-")
        inspector_header.addLayout(inspector_heading, 1)
        inspector_header.addWidget(self.note_inspector_state, 0, Qt.AlignTop)
        inspector_layout.addLayout(inspector_header)

        self.note_inspector_supplier = QLabel("Escolhe um documento para consultar o seu contexto.")
        self.note_inspector_supplier.setProperty("role", "muted")
        self.note_inspector_supplier.setWordWrap(True)
        self.note_inspector_delivery = QLabel("Entrega: -")
        self.note_inspector_delivery.setProperty("role", "muted")
        inspector_layout.addWidget(self.note_inspector_supplier)
        inspector_layout.addWidget(self.note_inspector_delivery)

        summary_strip = QFrame()
        summary_strip.setObjectName("PurchaseListSummary")
        summary_layout = QHBoxLayout(summary_strip)
        summary_layout.setContentsMargins(10, 8, 10, 8)
        summary_layout.setSpacing(8)

        def _list_metric(label_text: str) -> tuple[QWidget, QLabel]:
            host = QWidget()
            metric_layout = QVBoxLayout(host)
            metric_layout.setContentsMargins(0, 0, 0, 0)
            metric_layout.setSpacing(0)
            label = QLabel(label_text.upper())
            label.setObjectName("PurchaseListMetricLabel")
            value = QLabel("-")
            value.setObjectName("PurchaseListMetricValue")
            metric_layout.addWidget(label)
            metric_layout.addWidget(value)
            return host, value

        value_metric, self.note_inspector_value = _list_metric("Valor")
        lines_metric, self.note_inspector_lines = _list_metric("Linhas")
        summary_layout.addWidget(value_metric, 1)
        summary_layout.addWidget(lines_metric, 1)
        inspector_layout.addWidget(summary_strip)

        flow_title = QLabel("Fluxo de aprovisionamento")
        flow_title.setStyleSheet("font-size: 12px; font-weight: 800; color: #0f172a;")
        inspector_layout.addWidget(flow_title)
        self.note_inspector_flow = QLabel(
            "Cotação, adjudicação, encomenda, receção e documentação ficam ligados ao mesmo registo."
        )
        self.note_inspector_flow.setProperty("role", "muted")
        self.note_inspector_flow.setWordWrap(True)
        inspector_layout.addWidget(self.note_inspector_flow)
        inspector_layout.addStretch(1)

        list_workspace = QSplitter(Qt.Horizontal)
        list_workspace.setChildrenCollapsible(False)
        if notes_widget is not None:
            list_workspace.addWidget(notes_widget)
        list_workspace.addWidget(self.note_list_inspector)
        list_workspace.setStretchFactor(0, 1)
        list_workspace.setStretchFactor(1, 0)
        list_workspace.setSizes([1380, 420])
        list_layout.addWidget(list_workspace, 1)

        self.detail_page = QWidget()
        detail_outer = QVBoxLayout(self.detail_page)
        detail_outer.setContentsMargins(0, 0, 0, 0)
        detail_outer.setSpacing(0)
        self.note_detail_scroll = QScrollArea()
        self.note_detail_scroll.setWidgetResizable(True)
        self.note_detail_scroll.setFrameShape(QFrame.NoFrame)
        self.note_detail_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.note_detail_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        detail_outer.addWidget(self.note_detail_scroll)
        self.note_detail_host = QWidget()
        self.note_detail_scroll.setWidget(self.note_detail_host)
        detail_layout = QVBoxLayout(self.note_detail_host)
        detail_layout.setContentsMargins(0, 0, 0, 8)
        detail_layout.setSpacing(14)
        back_btn = QPushButton("Voltar")
        back_btn.setProperty("variant", "secondary")
        back_btn.setProperty("compact", "true")
        back_btn.setFixedWidth(72)
        back_btn.setMinimumHeight(29)
        back_btn.setMaximumHeight(31)
        back_btn.setStyleSheet("font-family: 'Segoe UI'; font-size: 10px; font-weight: 700;")
        back_btn.clicked.connect(self._show_note_list)
        form_widget = form_item.widget() if form_item is not None else None
        form_widget_layout = form_widget.layout() if isinstance(form_widget, QWidget) else None
        base_actions_item = form_widget_layout.itemAt(0) if form_widget_layout is not None and form_widget_layout.count() else None
        base_actions_layout = base_actions_item.layout() if base_actions_item is not None else None
        if isinstance(base_actions_layout, QHBoxLayout):
            base_actions_layout.insertWidget(0, back_btn)
        _adopt_layout_item(detail_layout, form_item, 1)

        self.view_stack.addWidget(self.list_page)
        self.view_stack.addWidget(self.detail_page)
        root.addWidget(self.view_stack, 1)

        self.notes_table.itemSelectionChanged.connect(self._sync_list_buttons)
        self.notes_table.itemDoubleClicked.connect(lambda *_args: self._open_selected_note())
        self._show_note_list()
        self._sync_list_buttons()

    def refresh(self) -> None:
        previous = self.current_number
        keep_detail = self.view_stack.currentWidget() is self.detail_page and bool(previous)
        self.supplier_rows = self.backend.ne_suppliers()
        self._set_supplier_items()
        self.rows = self.backend.ne_rows(self.filter_edit.currentText().strip(), self.state_combo.currentText())
        self._refresh_note_metrics()
        _fill_table(
            self.notes_table,
            [[r.get("numero", "-"), r.get("fornecedor", "-"), r.get("data_entrega", "-"), r.get("estado", "-"), _fmt_eur(r.get("total", 0)), r.get("linhas", 0)] for r in self.rows],
            align_center_from=4,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.notes_table, row_index, str(row.get("estado", "")))
        if self.notes_table.rowCount() == 0:
            PurchaseNotesPage._new_note(self, reset_number=False)
            self._show_note_list()
            self._sync_list_buttons()
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "")).strip() == previous:
                    row_index = index
                    break
        self.notes_table.selectRow(row_index)
        self._load_selected_note()
        if keep_detail:
            self._show_note_detail()
        else:
            self._show_note_list()
        self._sync_list_buttons()

    def _show_note_list(self) -> None:
        self.view_stack.setCurrentWidget(self.list_page)
        self._sync_list_buttons()

    def _show_note_detail(self) -> None:
        self.view_stack.setCurrentWidget(self.detail_page)

    def can_auto_refresh(self) -> bool:
        return self.view_stack.currentWidget() is self.list_page

    def _sync_list_buttons(self) -> None:
        row = self._selected_row()
        has_selection = bool(row)
        self.open_note_btn.setEnabled(has_selection)
        if hasattr(self, "remove_note_list_btn"):
            self.remove_note_list_btn.setEnabled(has_selection)
        if not hasattr(self, "note_inspector_number"):
            return
        if not row:
            self.note_inspector_number.setText("Sem seleção")
            self.note_inspector_supplier.setText("Escolhe um documento para consultar o seu contexto.")
            self.note_inspector_delivery.setText("Entrega: -")
            self.note_inspector_value.setText("-")
            self.note_inspector_lines.setText("-")
            self.note_inspector_flow.setText(
                "Cotação, adjudicação, encomenda, receção e documentação ficam ligados ao mesmo registo."
            )
            _apply_state_chip(self.note_inspector_state, "-")
            return
        numero = str(row.get("numero", "") or "").strip() or "-"
        fornecedor = str(row.get("fornecedor", "") or "").strip() or "Fornecedor por definir"
        entrega = str(row.get("data_entrega", "") or "").strip() or "Por definir"
        estado = str(row.get("estado", "") or "").strip() or "-"
        self.note_inspector_number.setText(numero)
        self.note_inspector_supplier.setText(fornecedor)
        self.note_inspector_delivery.setText(f"Entrega: {entrega}")
        self.note_inspector_value.setText(_fmt_eur(row.get("total", 0)))
        self.note_inspector_lines.setText(str(int(float(row.get("linhas", 0) or 0))))
        _apply_state_chip(self.note_inspector_state, estado)
        state_key = unicodedata.normalize("NFKD", estado).encode("ascii", "ignore").decode().casefold()
        if "entreg" in state_key:
            flow_text = "Receção concluída. O documento mantém o histórico de linhas, custos e anexos."
        elif "parcial" in state_key:
            flow_text = "Receção parcial. Consulta o detalhe para registar as quantidades ainda pendentes."
        elif "aprov" in state_key or "enviad" in state_key:
            flow_text = "Documento adjudicado. O detalhe reúne envio, receção e documentação do fornecedor."
        else:
            flow_text = "Documento em preparação. Abre-o para rever linhas, fornecedores e condições."
        self.note_inspector_flow.setText(flow_text)

    def _open_selected_note(self) -> None:
        if not self._selected_row():
            QMessageBox.warning(self, "Notas Encomenda", "Seleciona uma nota.")
            return
        self._load_selected_note()
        self._show_note_detail()

    def _new_note(self, reset_number: bool = True) -> None:
        PurchaseNotesPage._new_note(self, reset_number=reset_number)
        if hasattr(self, "view_stack"):
            self._show_note_detail()

    def _remove_note(self) -> None:
        PurchaseNotesPage._remove_note(self)
        self._show_note_list()
        self._sync_list_buttons()
