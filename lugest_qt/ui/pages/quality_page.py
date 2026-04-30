from __future__ import annotations

import os
import webbrowser
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard


def _table_item(value: Any, *, center: bool = False) -> QTableWidgetItem:
    item = QTableWidgetItem(str(value or ""))
    if center:
        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
    return item


def _fill_table(table: QTableWidget, rows: list[list[Any]], center_from: int = 0) -> None:
    sorting = table.isSortingEnabled()
    table.setSortingEnabled(False)
    table.setRowCount(len(rows))
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            table.setItem(row_index, col_index, _table_item(value, center=col_index >= center_from))
    table.setSortingEnabled(sorting)


class QualityPage(QWidget):
    page_title = "Qualidade"
    page_subtitle = "Rastreabilidade, evidencias e nao conformidades para apoiar ISO 9001."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.link_options: dict[str, list[dict[str, str]]] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(10)

        stats = QGridLayout()
        stats.setHorizontalSpacing(12)
        stats.setVerticalSpacing(10)
        self.open_nc_card = StatCard("NC abertas")
        self.overdue_nc_card = StatCard("NC fora prazo")
        self.supplier_nc_card = StatCard("NC fornecedores")
        self.blocked_materials_card = StatCard("Material bloqueado")
        for index, card in enumerate((self.open_nc_card, self.overdue_nc_card, self.supplier_nc_card, self.blocked_materials_card)):
            card.setMinimumHeight(92)
            stats.addWidget(card, 0, index)
        root.addLayout(stats)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("QualityTabs")
        self.tabs.setDocumentMode(True)
        root.addWidget(self.tabs, 1)

        self.reception_rows: list[dict[str, Any]] = []
        self.reception_filter = QLineEdit()
        self.reception_filter.setPlaceholderText("Filtrar receções a inspecionar")
        self.reception_filter.setProperty("qualityFilter", "true")
        self.reception_state = QComboBox()
        self.reception_state.addItems(["Pendentes", "Aprovados", "Rejeitados", "Devolução", "Averiguação", "Todos"])
        self.reception_table = QTableWidget(0, 13)
        self.reception_table.setHorizontalHeaderLabels(["Tipo", "Ref", "Estado Q.", "Logística", "Qtd qual.", "Fornecedor", "Lote", "NE", "Material", "Esp.", "Descrição", "NC"])
        self._configure_table(self.reception_table)
        self.reception_table.setHorizontalHeaderLabels(["Tipo", "Ref", "Estado Q.", "Logística", "Qtd pend.", "Fornecedor", "Lote", "NE", "Material", "Esp.", "Descrição", "NC", "Mov."])
        self.reception_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        for col, width in ((0, 86), (1, 120), (2, 116), (3, 92), (4, 74), (5, 132), (6, 118), (7, 126), (8, 150), (9, 62), (10, 240), (11, 90)):
            self.reception_table.setColumnWidth(col, width)
        self.reception_table.horizontalHeader().setSectionResizeMode(10, QHeaderView.Stretch)
        self.reception_table.horizontalHeader().setStretchLastSection(True)
        self.reception_table.setColumnHidden(12, True)
        self.tabs.addTab(self._reception_tab(), "Receção / inspeção")

        self.nc_filter = QLineEdit()
        self.nc_filter.setPlaceholderText("Filtrar nao conformidades")
        self.nc_filter.setProperty("qualityFilter", "true")
        self.nc_state = QComboBox()
        self.nc_state.addItems(["Ativas", "Abertas", "Em tratamento", "Fechadas", "Todos"])
        self.nc_table = QTableWidget(0, 11)
        self.nc_table.setHorizontalHeaderLabels(["ID", "Estado", "Grav.", "Qtd rej.", "Entidade", "Referencia", "Origem", "Tipo", "Responsavel", "Prazo", "Descricao"])
        self._configure_table(self.nc_table)
        self.nc_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        for col, width in ((0, 92), (1, 92), (2, 78), (3, 74), (4, 170), (5, 130), (6, 140), (7, 100), (8, 110), (9, 90), (10, 260)):
            self.nc_table.setColumnWidth(col, width)
        self.nc_table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self._nc_tab(), "Nao conformidades")

        self.doc_filter = QLineEdit()
        self.doc_filter.setPlaceholderText("Filtrar documentos/evidencias")
        self.doc_filter.setProperty("qualityFilter", "true")
        self.doc_table = QTableWidget(0, 8)
        self.doc_table.setHorizontalHeaderLabels(["ID", "Titulo", "Tipo", "Entidade", "Referencia", "Versao", "Estado", "Atualizado"])
        self._configure_table(self.doc_table)
        self.tabs.addTab(self._documents_tab(), "Documentos")

        self.audit_filter = QLineEdit()
        self.audit_filter.setPlaceholderText("Filtrar auditoria")
        self.audit_filter.setProperty("qualityFilter", "true")
        self.audit_table = QTableWidget(0, 6)
        self.audit_table.setHorizontalHeaderLabels(["Data", "Utilizador", "Acao", "Entidade", "Ref.", "Resumo"])
        self._configure_table(self.audit_table)
        self.tabs.addTab(self._audit_tab(), "Auditoria")

        self.check_table = QTableWidget(0, 3)
        self.check_table.setHorizontalHeaderLabels(["Area", "Estado", "Evidencia"])
        self._configure_table(self.check_table)
        self.tabs.addTab(self._checklist_tab(), "Checklist ISO")

        self.nc_filter.textChanged.connect(self._refresh_ncs)
        self.nc_state.currentTextChanged.connect(lambda _text: self._refresh_ncs())
        self.reception_filter.textChanged.connect(self._refresh_reception)
        self.reception_state.currentTextChanged.connect(lambda _text: self._refresh_reception())
        self.doc_filter.textChanged.connect(self._refresh_documents)
        self.audit_filter.textChanged.connect(self._refresh_audit)

    def _configure_table(self, table: QTableWidget) -> None:
        table.setObjectName("QualityTable")
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setSortingEnabled(True)

    def _section_header(self, title: str, subtitle: str = "") -> CardFrame:
        card = CardFrame()
        card.setObjectName("QualitySection")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(12)
        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)
        title_label = QLabel(title)
        title_label.setProperty("role", "section_title")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setProperty("role", "section_subtitle")
        subtitle_label.setWordWrap(True)
        text_col.addWidget(title_label)
        if subtitle:
            text_col.addWidget(subtitle_label)
        layout.addLayout(text_col, 1)
        return card

    def _reception_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(self._section_header("Rececao e inspecao", "Entrada fisica separada da libertacao de stock. Aprova, rejeita ou mantem em analise depois de validar a rececao."))
        toolbar_card = CardFrame()
        toolbar_card.setObjectName("QualityToolbar")
        toolbar = QHBoxLayout(toolbar_card)
        toolbar.setContentsMargins(12, 10, 12, 10)
        toolbar.setSpacing(8)
        toolbar.addWidget(self.reception_filter, 1)
        toolbar.addWidget(self.reception_state)
        evaluate_btn = QPushButton("Avaliar")
        evaluate_btn.clicked.connect(self._evaluate_selected_reception)
        approve_btn = QPushButton("Aprovar")
        approve_btn.setProperty("variant", "success")
        approve_btn.clicked.connect(lambda: self._evaluate_selected_reception(default_status="APROVADO"))
        reject_btn = QPushButton("Rejeitar / NC")
        reject_btn.setProperty("variant", "danger")
        reject_btn.clicked.connect(lambda: self._evaluate_selected_reception(default_status="REJEITADO"))
        toolbar.addWidget(evaluate_btn)
        toolbar.addWidget(approve_btn)
        toolbar.addWidget(reject_btn)
        layout.addWidget(toolbar_card)
        note = QLabel("A receção apenas dá entrada física. Nesta lista a qualidade liberta, mantém em inspeção ou rejeita o material/produto recebido.")
        note.setProperty("role", "muted")
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addWidget(self.reception_table, 1)
        return page

    def _selected_id(self, table: QTableWidget) -> str:
        row = table.currentRow()
        if row < 0:
            return ""
        item = table.item(row, 0)
        return str(item.text() if item else "").strip()

    def _nc_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(self._section_header("Nao conformidades", "Registo de desvios, reclamacoes a fornecedores, devolucoes e acoes corretivas com rastreabilidade por entidade."))
        toolbar_card = CardFrame()
        toolbar_card.setObjectName("QualityToolbar")
        toolbar = QHBoxLayout(toolbar_card)
        toolbar.setContentsMargins(12, 10, 12, 10)
        toolbar.setSpacing(8)
        toolbar.addWidget(self.nc_filter, 1)
        toolbar.addWidget(self.nc_state)
        add_btn = QPushButton("Nova NC")
        add_btn.clicked.connect(lambda: self._edit_nc({}))
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        edit_btn.clicked.connect(self._edit_selected_nc)
        close_btn = QPushButton("Fechar NC")
        close_btn.setProperty("variant", "secondary")
        close_btn.clicked.connect(self._close_selected_nc)
        release_btn = QPushButton("Libertar material")
        release_btn.setProperty("variant", "secondary")
        release_btn.clicked.connect(self._release_selected_nc_material)
        pdf_btn = QPushButton("PDF NC")
        pdf_btn.setProperty("variant", "secondary")
        pdf_btn.clicked.connect(self._pdf_selected_nc)
        remove_btn = QPushButton("Apagar")
        remove_btn.setProperty("variant", "danger")
        remove_btn.clicked.connect(self._remove_selected_nc)
        for button in (add_btn, edit_btn, close_btn, release_btn, pdf_btn, remove_btn):
            toolbar.addWidget(button)
        layout.addWidget(toolbar_card)
        layout.addWidget(self.nc_table, 1)
        return page

    def _documents_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(self._section_header("Documentos e evidencias", "Certificados, fotografias, relatorios e documentos de suporte ligados a material, fornecedor, OPP ou cliente."))
        toolbar_card = CardFrame()
        toolbar_card.setObjectName("QualityToolbar")
        toolbar = QHBoxLayout(toolbar_card)
        toolbar.setContentsMargins(12, 10, 12, 10)
        toolbar.setSpacing(8)
        toolbar.addWidget(self.doc_filter, 1)
        add_btn = QPushButton("Novo documento")
        add_btn.clicked.connect(lambda: self._edit_document({}))
        edit_btn = QPushButton("Editar")
        edit_btn.setProperty("variant", "secondary")
        edit_btn.clicked.connect(self._edit_selected_document)
        open_btn = QPushButton("Abrir ficheiro")
        open_btn.setProperty("variant", "secondary")
        open_btn.clicked.connect(self._open_selected_document)
        remove_btn = QPushButton("Apagar")
        remove_btn.setProperty("variant", "danger")
        remove_btn.clicked.connect(self._remove_selected_document)
        for button in (add_btn, edit_btn, open_btn, remove_btn):
            toolbar.addWidget(button)
        layout.addWidget(toolbar_card)
        layout.addWidget(self.doc_table, 1)
        return page

    def _audit_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(self._section_header("Auditoria", "Historico de eventos criticos para perceber quem decidiu, quando decidiu e com que referencia."))
        toolbar_card = CardFrame()
        toolbar_card.setObjectName("QualityToolbar")
        toolbar = QHBoxLayout(toolbar_card)
        toolbar.setContentsMargins(12, 10, 12, 10)
        toolbar.addWidget(self.audit_filter, 1)
        layout.addWidget(toolbar_card)
        layout.addWidget(self.audit_table, 1)
        return page

    def _checklist_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(self._section_header("Checklist ISO", "Pontos de controlo praticos para auditoria interna e acompanhamento do sistema de gestao."))
        toolbar_card = CardFrame()
        toolbar_card.setObjectName("QualityToolbar")
        actions = QHBoxLayout(toolbar_card)
        actions.setContentsMargins(12, 10, 12, 10)
        dossier_btn = QPushButton("Gerar dossier PDF")
        dossier_btn.clicked.connect(self._open_dossier_pdf)
        actions.addStretch(1)
        actions.addWidget(dossier_btn)
        note = QLabel("Checklist operacional para apoiar auditorias. Nao substitui a interpretacao formal da norma.")
        note.setProperty("role", "muted")
        layout.addWidget(toolbar_card)
        layout.addWidget(note)
        layout.addWidget(self.check_table, 1)
        return page

    def refresh(self) -> None:
        summary = dict(self.backend.quality_summary() or {})
        self.open_nc_card.set_data(str(summary.get("open_nc", 0)), "Nao conformidades ainda ativas")
        self.open_nc_card.set_tone("warning" if int(summary.get("open_nc", 0) or 0) else "success")
        self.overdue_nc_card.set_data(str(summary.get("overdue_nc", 0)), "Prazo ultrapassado")
        self.overdue_nc_card.set_tone("danger" if int(summary.get("overdue_nc", 0) or 0) else "success")
        self.supplier_nc_card.set_data(str(summary.get("supplier_nc", 0)), "Reclamacoes/NC ligadas a fornecedores")
        self.supplier_nc_card.set_tone("warning" if int(summary.get("supplier_nc", 0) or 0) else "success")
        self.blocked_materials_card.set_data(str(summary.get("blocked_materials", 0)), "Lotes em inspecao, bloqueados ou reclamados")
        self.blocked_materials_card.set_tone("danger" if int(summary.get("blocked_materials", 0) or 0) else "success")
        self._refresh_reception()
        self._refresh_ncs()
        self._refresh_documents()
        self._refresh_audit()
        self._refresh_checklist()

    def _refresh_reception(self) -> None:
        try:
            self.reception_rows = list(self.backend.quality_reception_rows(self.reception_filter.text(), self.reception_state.currentText()) or [])
        except Exception as exc:
            self.reception_rows = []
            QMessageBox.warning(self, "Qualidade", str(exc))
        _fill_table(
            self.reception_table,
            [
                [
                    row.get("tipo", ""),
                    row.get("ref", ""),
                    row.get("quality_status", ""),
                    row.get("logistic_status", ""),
                    row.get("qtd", ""),
                    row.get("fornecedor", ""),
                    row.get("lote", ""),
                    row.get("referencia", ""),
                    row.get("material", ""),
                    row.get("espessura", ""),
                    row.get("descricao", ""),
                    row.get("nc_id", ""),
                    row.get("movement_id", ""),
                ]
                for row in self.reception_rows
            ],
            center_from=2,
        )
        for row_index in range(self.reception_table.rowCount()):
            movement_item = self.reception_table.item(row_index, 12)
            movement_id = str(movement_item.text() if movement_item else "").strip()
            ref_item = self.reception_table.item(row_index, 1)
            type_item = self.reception_table.item(row_index, 0)
            ref = str(ref_item.text() if ref_item else "").strip()
            kind = str(type_item.text() if type_item else "").strip()
            payload = next(
                (
                    dict(row)
                    for row in self.reception_rows
                    if (
                        movement_id
                        and str(row.get("movement_id", "") or "").strip() == movement_id
                    )
                    or (
                        not movement_id
                        and str(row.get("ref", "") or "").strip() == ref
                        and str(row.get("tipo", "") or "").strip() == kind
                    )
                ),
                {},
            )
            first_item = self.reception_table.item(row_index, 0)
            if first_item is not None:
                first_item.setData(Qt.UserRole, payload)

    def _selected_reception(self) -> dict[str, Any]:
        row_index = self.reception_table.currentRow()
        if row_index < 0:
            return {}
        first_item = self.reception_table.item(row_index, 0)
        if first_item is not None:
            payload = first_item.data(Qt.UserRole)
            if isinstance(payload, dict) and payload:
                return dict(payload)
        movement_item = self.reception_table.item(row_index, 12)
        movement_id = str(movement_item.text() if movement_item else "").strip()
        if movement_id:
            for row in self.reception_rows:
                if str(row.get("movement_id", "") or "").strip() == movement_id:
                    return dict(row)
        ref_item = self.reception_table.item(row_index, 1)
        type_item = self.reception_table.item(row_index, 0)
        ref = str(ref_item.text() if ref_item else "").strip()
        kind = str(type_item.text() if type_item else "").strip()
        for row in self.reception_rows:
            if str(row.get("ref", "") or "").strip() == ref and str(row.get("tipo", "") or "").strip() == kind:
                return dict(row)
        return {}

    def _evaluate_selected_reception(self, default_status: str = "") -> None:
        current = self._selected_reception()
        if not current:
            QMessageBox.warning(self, "Qualidade", "Seleciona uma linha recebida.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Inspeção de receção")
        dialog.resize(760, 520)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        entity = QLineEdit(" | ".join(part for part in (str(current.get("tipo", "") or ""), str(current.get("ref", "") or ""), str(current.get("descricao", "") or "")) if part))
        entity.setReadOnly(True)
        pending_qty = float(current.get("qtd", 0) or 0)
        received_qty = float(current.get("qtd_recebida", pending_qty) or pending_qty)
        received = QLineEdit(f"{received_qty:.4f}".rstrip("0").rstrip("."))
        received.setReadOnly(True)
        approved_qty = QDoubleSpinBox()
        approved_qty.setDecimals(4)
        approved_qty.setRange(0.0, max(0.0, pending_qty))
        approved_qty.setSingleStep(1.0)
        rejected_qty = QDoubleSpinBox()
        rejected_qty.setDecimals(4)
        rejected_qty.setRange(0.0, max(0.0, pending_qty))
        rejected_qty.setSingleStep(1.0)
        status = QComboBox()
        status.addItems(["EM_INSPECAO", "EM_AVERIGUACAO", "APROVADO", "REJEITADO", "DEVOLVER_FORNECEDOR"])
        status.setCurrentText(default_status or str(current.get("quality_status", "") or "EM_INSPECAO"))
        if status.currentText() == "APROVADO":
            approved_qty.setValue(pending_qty)
        elif status.currentText() in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
            rejected_qty.setValue(pending_qty)
        defect = QComboBox()
        defect.setEditable(True)
        defect.setInsertPolicy(QComboBox.NoInsert)
        if str(current.get("tipo", "") or "").casefold().startswith("material"):
            defect.addItems(["", "Empenado", "Oxidação", "Riscos/danos", "Dimensão errada", "Material errado", "Certificado em falta", "Quantidade errada", "Sem rastreabilidade/lote", "Outro"])
        else:
            defect.addItems(["", "Produto danificado", "Embalagem danificada", "Referência errada", "Quantidade errada", "Faltam acessórios", "Sem certificado/declaração", "Funcionamento não conforme", "Prazo/validade incorreto", "Outro"])
        defect.setCurrentText(str(current.get("defeito", "") or ""))
        decision = QComboBox()
        decision.setEditable(True)
        decision.setInsertPolicy(QComboBox.NoInsert)
        decision.addItems(["", "Libertar para stock", "Manter em inspeção", "Aceitar condicionado", "Bloquear stock", "Abrir reclamação fornecedor", "Rejeitar entrada", "Devolver ao fornecedor", "Corrigir internamente"])
        decision.setCurrentText(str(current.get("decisao", "") or ""))
        notes = QTextEdit()
        notes.setPlaceholderText("Observação complementar opcional")

        def sync_decision() -> None:
            if decision.currentText().strip():
                return
            value = status.currentText().strip()
            if value == "APROVADO":
                decision.setCurrentText("Libertar para stock")
            elif value == "REJEITADO":
                decision.setCurrentText("Abrir reclamação fornecedor")
            elif value == "DEVOLVER_FORNECEDOR":
                decision.setCurrentText("Devolver ao fornecedor")
            elif value == "EM_AVERIGUACAO":
                decision.setCurrentText("Manter em averiguação")
            else:
                decision.setCurrentText("Manter em inspeção")

        def sync_quantities_from_status() -> None:
            value = status.currentText().strip()
            if value == "APROVADO":
                approved_qty.setValue(pending_qty)
                rejected_qty.setValue(0.0)
            elif value in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                approved_qty.setValue(0.0)
                rejected_qty.setValue(pending_qty)
            elif approved_qty.value() + rejected_qty.value() > pending_qty:
                approved_qty.setValue(0.0)
                rejected_qty.setValue(0.0)

        def sync_clean_quality_fields(force: bool = False) -> None:
            value = status.currentText().strip()
            decision_text = decision.currentText().casefold()
            if value == "APROVADO":
                approved_qty.setValue(pending_qty)
                rejected_qty.setValue(0.0)
                if force or not decision_text or "devol" in decision_text or "reclam" in decision_text:
                    decision.setCurrentText("Libertar para stock")
                if force:
                    defect.setCurrentText("")
            elif value == "REJEITADO":
                approved_qty.setValue(0.0)
                rejected_qty.setValue(pending_qty)
                if force or not decision_text or "libert" in decision_text:
                    decision.setCurrentText("Abrir reclamacao fornecedor")
            elif value == "DEVOLVER_FORNECEDOR":
                approved_qty.setValue(0.0)
                rejected_qty.setValue(pending_qty)
                if force or not decision_text or "libert" in decision_text:
                    decision.setCurrentText("Devolver ao fornecedor")
            elif value == "EM_AVERIGUACAO":
                if approved_qty.value() + rejected_qty.value() > pending_qty:
                    approved_qty.setValue(0.0)
                    rejected_qty.setValue(0.0)
                if force or not decision_text or "libert" in decision_text:
                    decision.setCurrentText("Manter em averiguacao")
            else:
                if approved_qty.value() + rejected_qty.value() > pending_qty:
                    approved_qty.setValue(0.0)
                    rejected_qty.setValue(0.0)
                if force or not decision_text or "libert" in decision_text:
                    decision.setCurrentText("Manter em inspecao")
            defect.setEnabled(value != "APROVADO" or rejected_qty.value() > 0)

        def sync_partial_quality_fields() -> None:
            if approved_qty.value() + rejected_qty.value() > pending_qty:
                rejected_qty.setValue(max(0.0, pending_qty - approved_qty.value()))
            defect.setEnabled(status.currentText().strip() != "APROVADO" or rejected_qty.value() > 0)
            if rejected_qty.value() > 0 and status.currentText().strip() == "APROVADO":
                decision.setCurrentText("Aceitar condicionado")

        status.currentTextChanged.connect(lambda _text: (sync_decision(), sync_quantities_from_status(), sync_clean_quality_fields(True)))
        approved_qty.valueChanged.connect(lambda _value: sync_partial_quality_fields())
        rejected_qty.valueChanged.connect(lambda _value: sync_partial_quality_fields())
        sync_decision()
        sync_clean_quality_fields(bool(default_status))
        for label, widget in (
            ("Linha", entity),
            ("Qtd recebida", received),
            ("Qtd boa/aprovada", approved_qty),
            ("Qtd rejeitada", rejected_qty),
            ("Inspeção", status),
            ("Defeito/obs.", defect),
            ("Decisão", decision),
            ("Notas", notes),
        ):
            form.addRow(label, widget)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Guardar")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        defect_txt = defect.currentText().strip()
        notes_txt = notes.toPlainText().strip()
        if notes_txt:
            defect_txt = f"{defect_txt} | {notes_txt}".strip(" |")
        if approved_qty.value() + rejected_qty.value() > pending_qty + 1e-9:
            QMessageBox.warning(self, "Qualidade", "A quantidade boa + rejeitada não pode ultrapassar a quantidade pendente.")
            return
        if status.currentText() == "APROVADO" and rejected_qty.value() <= 0:
            defect_txt = ""
            decision.setCurrentText("Libertar para stock")
        payload = {
            "tipo": current.get("tipo", ""),
            "id": current.get("id", current.get("ref", "")),
            "movement_id": current.get("movement_id", ""),
            "referencia": current.get("referencia", ""),
            "quality_status": status.currentText(),
            "qtd_aprovada": approved_qty.value(),
            "qtd_rejeitada": rejected_qty.value(),
            "defeito": defect_txt,
            "decisao": decision.currentText(),
            "create_nc": rejected_qty.value() > 0 or status.currentText() in {"REJEITADO", "DEVOLVER_FORNECEDOR"} or bool(defect_txt),
        }
        try:
            self.backend.quality_reception_save(payload)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _refresh_ncs(self) -> None:
        rows = list(self.backend.quality_nc_rows(self.nc_filter.text(), self.nc_state.currentText()) or [])
        _fill_table(
            self.nc_table,
            [
                [
                    row.get("id", ""),
                    row.get("estado", ""),
                    row.get("gravidade", ""),
                    row.get("qtd_rejeitada", ""),
                    row.get("entidade_label", "") or row.get("entidade_id", ""),
                    row.get("referencia", ""),
                    row.get("origem", ""),
                    row.get("tipo", ""),
                    row.get("responsavel", ""),
                    row.get("prazo", ""),
                    row.get("descricao", ""),
                ]
                for row in rows
            ],
            center_from=1,
        )

    def _refresh_documents(self) -> None:
        rows = list(self.backend.quality_document_rows(self.doc_filter.text()) or [])
        _fill_table(
            self.doc_table,
            [
                [
                    row.get("id", ""),
                    row.get("titulo", ""),
                    row.get("tipo", ""),
                    row.get("entidade", ""),
                    row.get("referencia", ""),
                    row.get("versao", ""),
                    row.get("estado", ""),
                    row.get("updated_at", ""),
                ]
                for row in rows
            ],
            center_from=5,
        )

    def _refresh_audit(self) -> None:
        rows = list(self.backend.audit_rows(self.audit_filter.text()) or [])
        _fill_table(
            self.audit_table,
            [
                [
                    row.get("created_at", ""),
                    row.get("user", ""),
                    row.get("action", ""),
                    row.get("entity_type", ""),
                    row.get("entity_id", ""),
                    row.get("summary", ""),
                ]
                for row in rows
            ],
            center_from=0,
        )

    def _refresh_checklist(self) -> None:
        rows = list(self.backend.quality_iso_checklist() or [])
        _fill_table(self.check_table, [[row.get("area", ""), row.get("estado", ""), row.get("evidencia", "")] for row in rows], center_from=1)

    def _find_nc(self, nc_id: str) -> dict[str, Any]:
        for row in list(self.backend.quality_nc_rows("", "Todos") or []):
            if str(row.get("id", "") or "").strip() == nc_id:
                return dict(row)
        return {}

    def _edit_selected_nc(self) -> None:
        nc_id = self._selected_id(self.nc_table)
        if not nc_id:
            QMessageBox.warning(self, "Qualidade", "Seleciona uma nao conformidade.")
            return
        self._edit_nc(self._find_nc(nc_id))

    def _edit_nc(self, current: dict[str, Any]) -> None:
        try:
            self.link_options = dict(self.backend.quality_link_options() or {})
        except Exception:
            self.link_options = {}
        dialog = QDialog(self)
        dialog.setWindowTitle("Nao conformidade")
        dialog.resize(760, 620)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        nc_id = QLineEdit(str(current.get("id", "") or ""))
        nc_id.setReadOnly(True)
        origem = QComboBox()
        origem.addItems(["Cliente", "Fornecedor", "Stock", "Producao", "Expedicao", "Sistema", "Outro"])
        origem.setCurrentText(str(current.get("origem", "") or "Producao"))
        referencia = QLineEdit(str(current.get("referencia", "") or ""))
        tipo = QComboBox()
        tipo.addItems(["Produto", "Processo", "Fornecedor", "Documento", "Cliente", "Sistema"])
        tipo.setCurrentText(str(current.get("tipo", "") or "Processo"))
        entidade_tipo = QComboBox()
        entity_types = list(self.link_options.keys()) or ["Livre", "OPP", "Encomenda", "Material", "Fornecedor", "Cliente", "Documento"]
        entidade_tipo.addItems(entity_types)
        entidade_tipo.setCurrentText(str(current.get("entidade_tipo", "") or "Livre"))
        entidade_id = QComboBox()
        entidade_id.setEditable(True)
        gravidade = QComboBox()
        gravidade.addItems(["Baixa", "Media", "Alta", "Critica"])
        gravidade.setCurrentText(str(current.get("gravidade", "") or "Media"))
        estado = QComboBox()
        estado.addItems(["Aberta", "Em tratamento", "Fechada", "Cancelada"])
        estado.setCurrentText(str(current.get("estado", "") or "Aberta"))
        responsavel = QLineEdit(str(current.get("responsavel", "") or ""))
        prazo = QLineEdit(str(current.get("prazo", "") or ""))
        descricao = QTextEdit(str(current.get("descricao", "") or ""))
        causa = QTextEdit(str(current.get("causa", "") or ""))
        acao = QTextEdit(str(current.get("acao", "") or ""))
        eficacia = QTextEdit(str(current.get("eficacia", "") or ""))

        def refresh_entity_choices() -> None:
            selected_type = entidade_tipo.currentText().strip() or "Livre"
            current_id = str(current.get("entidade_id", "") or entidade_id.currentText() or "").strip()
            entidade_id.blockSignals(True)
            entidade_id.clear()
            for row in list(self.link_options.get(selected_type, []) or []):
                label = str(row.get("label", "") or row.get("id", "") or "").strip()
                value = str(row.get("id", "") or "").strip()
                entidade_id.addItem(label, value)
            if current_id:
                found = False
                for index in range(entidade_id.count()):
                    if str(entidade_id.itemData(index) or "").strip() == current_id:
                        entidade_id.setCurrentIndex(index)
                        found = True
                        break
                if not found:
                    entidade_id.setCurrentText(current_id)
            entidade_id.blockSignals(False)

        def sync_reference_from_entity() -> None:
            value = str(entidade_id.currentData() or entidade_id.currentText() or "").strip()
            if value and not referencia.text().strip():
                referencia.setText(value)

        entidade_tipo.currentTextChanged.connect(lambda _text: refresh_entity_choices())
        entidade_id.currentTextChanged.connect(lambda _text: sync_reference_from_entity())
        refresh_entity_choices()
        for label, widget in (
            ("ID", nc_id),
            ("Origem", origem),
            ("Referencia", referencia),
            ("Entidade tipo", entidade_tipo),
            ("Entidade ligada", entidade_id),
            ("Tipo", tipo),
            ("Gravidade", gravidade),
            ("Estado", estado),
            ("Responsavel", responsavel),
            ("Prazo", prazo),
            ("Descricao", descricao),
            ("Causa", causa),
            ("Acao corretiva", acao),
            ("Eficacia", eficacia),
        ):
            form.addRow(label, widget)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Guardar")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        layout.addWidget(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        if dialog.exec() != QDialog.Accepted:
            return
        payload = {
            "id": nc_id.text(),
            "origem": origem.currentText(),
            "referencia": referencia.text(),
            "entidade_tipo": entidade_tipo.currentText(),
            "entidade_id": str(entidade_id.currentData() or entidade_id.currentText() or "").strip(),
            "entidade_label": entidade_id.currentText(),
            "tipo": tipo.currentText(),
            "gravidade": gravidade.currentText(),
            "estado": estado.currentText(),
            "responsavel": responsavel.text(),
            "prazo": prazo.text(),
            "descricao": descricao.toPlainText(),
            "causa": causa.toPlainText(),
            "acao": acao.toPlainText(),
            "eficacia": eficacia.toPlainText(),
        }
        try:
            self.backend.quality_nc_save(payload)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _open_pdf_path(self, path: Any) -> None:
        path_txt = str(path or "").strip()
        if not path_txt:
            return
        try:
            if os.name == "nt":
                os.startfile(path_txt)  # type: ignore[attr-defined]
            else:
                webbrowser.open(Path(path_txt).as_uri())
        except Exception as exc:
            QMessageBox.warning(self, "Qualidade", str(exc))

    def _pdf_selected_nc(self) -> None:
        nc_id = self._selected_id(self.nc_table)
        if not nc_id:
            QMessageBox.warning(self, "Qualidade", "Seleciona uma nao conformidade.")
            return
        try:
            self._open_pdf_path(self.backend.quality_nc_pdf(nc_id))
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _open_dossier_pdf(self) -> None:
        try:
            self._open_pdf_path(self.backend.quality_dossier_pdf())
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _close_selected_nc(self) -> None:
        nc_id = self._selected_id(self.nc_table)
        if not nc_id:
            QMessageBox.warning(self, "Qualidade", "Seleciona uma nao conformidade.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Fechar nao conformidade")
        layout = QVBoxLayout(dialog)
        eficacia = QTextEdit()
        eficacia.setPlaceholderText("Valida a eficacia da acao corretiva")
        layout.addWidget(eficacia)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Fechar NC")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        layout.addWidget(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            self.backend.quality_nc_close(nc_id, eficacia.toPlainText())
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _release_selected_nc_material(self) -> None:
        nc_id = self._selected_id(self.nc_table)
        if not nc_id:
            QMessageBox.warning(self, "Qualidade", "Seleciona uma nao conformidade.")
            return
        if QMessageBox.question(
            self,
            "Qualidade",
            f"Libertar o material ligado a {nc_id} para stock disponivel?",
        ) != QMessageBox.Yes:
            return
        try:
            self.backend.quality_nc_release_material(nc_id)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _remove_selected_nc(self) -> None:
        nc_id = self._selected_id(self.nc_table)
        if not nc_id:
            return
        if QMessageBox.question(self, "Qualidade", f"Apagar {nc_id}?") != QMessageBox.Yes:
            return
        try:
            self.backend.quality_nc_remove(nc_id)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _find_document(self, doc_id: str) -> dict[str, Any]:
        for row in list(self.backend.quality_document_rows("") or []):
            if str(row.get("id", "") or "").strip() == doc_id:
                return dict(row)
        return {}

    def _edit_selected_document(self) -> None:
        doc_id = self._selected_id(self.doc_table)
        if not doc_id:
            QMessageBox.warning(self, "Qualidade", "Seleciona um documento.")
            return
        self._edit_document(self._find_document(doc_id))

    def _edit_document(self, current: dict[str, Any]) -> None:
        try:
            self.link_options = dict(self.backend.quality_link_options() or {})
        except Exception:
            self.link_options = {}
        dialog = QDialog(self)
        dialog.setWindowTitle("Documento / evidencia")
        dialog.resize(720, 420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        doc_id = QLineEdit(str(current.get("id", "") or ""))
        doc_id.setReadOnly(True)
        titulo = QLineEdit(str(current.get("titulo", "") or ""))
        tipo = QComboBox()
        tipo.addItems(["Procedimento", "Instrucao", "Certificado", "Relatorio", "Evidencia", "Registo"])
        tipo.setCurrentText(str(current.get("tipo", "") or "Evidencia"))
        entidade = QComboBox()
        entity_types = list(self.link_options.keys()) or ["Livre", "OPP", "Encomenda", "Material", "Fornecedor", "Cliente", "Documento"]
        entidade.addItems(entity_types)
        entidade.setCurrentText(str(current.get("entidade_tipo", "") or current.get("entidade", "") or "Livre"))
        referencia = QComboBox()
        referencia.setEditable(True)
        versao = QLineEdit(str(current.get("versao", "") or "1"))
        estado = QComboBox()
        estado.addItems(["Ativo", "Obsoleto", "Em revisao"])
        estado.setCurrentText(str(current.get("estado", "") or "Ativo"))
        responsavel = QLineEdit(str(current.get("responsavel", "") or ""))
        caminho = QLineEdit(str(current.get("caminho", "") or ""))
        browse_btn = QPushButton("Selecionar ficheiro")
        browse_btn.setProperty("variant", "secondary")
        browse_btn.clicked.connect(lambda: self._browse_file_into(caminho))
        path_row = QWidget()
        path_layout = QHBoxLayout(path_row)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.addWidget(caminho, 1)
        path_layout.addWidget(browse_btn)
        obs = QTextEdit(str(current.get("obs", "") or ""))

        def refresh_doc_refs() -> None:
            selected_type = entidade.currentText().strip() or "Livre"
            current_ref = str(current.get("entidade_id", "") or current.get("referencia", "") or referencia.currentText() or "").strip()
            referencia.blockSignals(True)
            referencia.clear()
            for row in list(self.link_options.get(selected_type, []) or []):
                referencia.addItem(str(row.get("label", "") or row.get("id", "") or "").strip(), str(row.get("id", "") or "").strip())
            if current_ref:
                found = False
                for index in range(referencia.count()):
                    if str(referencia.itemData(index) or "").strip() == current_ref:
                        referencia.setCurrentIndex(index)
                        found = True
                        break
                if not found:
                    referencia.setCurrentText(current_ref)
            referencia.blockSignals(False)

        entidade.currentTextChanged.connect(lambda _text: refresh_doc_refs())
        refresh_doc_refs()
        for label, widget in (
            ("ID", doc_id),
            ("Titulo", titulo),
            ("Tipo", tipo),
            ("Entidade", entidade),
            ("Referencia", referencia),
            ("Versao", versao),
            ("Estado", estado),
            ("Responsavel", responsavel),
            ("Ficheiro", path_row),
            ("Obs.", obs),
        ):
            form.addRow(label, widget)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Guardar")
        buttons.button(QDialogButtonBox.Cancel).setText("Cancelar")
        layout.addWidget(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            self.backend.quality_document_save(
                {
                    "id": doc_id.text(),
                    "titulo": titulo.text(),
                    "tipo": tipo.currentText(),
                    "entidade": entidade.currentText(),
                    "referencia": str(referencia.currentData() or referencia.currentText() or "").strip(),
                    "entidade_tipo": entidade.currentText(),
                    "entidade_id": str(referencia.currentData() or referencia.currentText() or "").strip(),
                    "versao": versao.text(),
                    "estado": estado.currentText(),
                    "responsavel": responsavel.text(),
                    "caminho": caminho.text(),
                    "obs": obs.toPlainText(),
                }
            )
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))

    def _browse_file_into(self, line_edit: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar ficheiro", "", "Todos (*.*)")
        if path:
            line_edit.setText(path)

    def _open_selected_document(self) -> None:
        doc_id = self._selected_id(self.doc_table)
        row = self._find_document(doc_id)
        path = str(row.get("caminho", "") or "").strip()
        if not path:
            QMessageBox.information(self, "Qualidade", "Este documento nao tem ficheiro associado.")
            return
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                webbrowser.open(Path(path).as_uri())
        except Exception as exc:
            QMessageBox.warning(self, "Qualidade", str(exc))

    def _remove_selected_document(self) -> None:
        doc_id = self._selected_id(self.doc_table)
        if not doc_id:
            return
        if QMessageBox.question(self, "Qualidade", f"Apagar {doc_id}?") != QMessageBox.Yes:
            return
        try:
            self.backend.quality_document_remove(doc_id)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Qualidade", str(exc))
