from __future__ import annotations

import re
import unicodedata

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QFormLayout,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStyle,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame


class _ProductAILookupThread(QThread):
    completed = Signal(dict)
    failed = Signal(str)

    def __init__(self, lookup, description: str, parent=None) -> None:
        super().__init__(parent)
        self._lookup = lookup
        self._description = description

    def run(self) -> None:
        try:
            self.completed.emit(dict(self._lookup(self._description) or {}))
        except Exception as exc:
            self.failed.emit(str(exc))


class _ProductAIResultDialog(QDialog):
    """Human-in-the-loop review for a product classification proposed by AI."""

    def __init__(
        self,
        result: dict,
        current_description: str,
        refine_callback=None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.action = "cancel"
        self.result = dict(result or {})
        self.current_description = str(current_description or "").strip()
        self.refine_callback = refine_callback
        self._refine_thread = None
        self._pending_question = ""
        self.setWindowTitle("Copiloto LuGEST · Análise do produto")
        self.setModal(True)
        self.resize(820, 720)
        self.setMinimumSize(760, 660)

        candidate = dict(result.get("candidate", {}) or {})
        sources = list(result.get("sources", []) or [])
        engine = str(result.get("engine", "") or "").strip()
        confidence = int(round(float(candidate.get("confidence", 0) or 0) * 100))
        category = str(candidate.get("categoria", "") or "").strip()
        subcategory = str(candidate.get("subcat", "") or "").strip()
        product_type = str(candidate.get("tipo", "") or "").strip()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title = QLabel("Análise do Copiloto")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #172b3f;")
        self.confidence_label = QLabel(f"CONFIANÇA {confidence}%")
        self.confidence_label.setStyleSheet(
            "background: #edf8df; color: #47780c; border: 1px solid #aedb75;"
            "padding: 6px 10px; font-size: 10px; font-weight: 800;"
        )
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(self.confidence_label)
        layout.addLayout(title_row)

        self.subtitle = QLabel(
            f"Motor: {engine.removeprefix('ollama-') if engine.startswith('ollama-') else engine or 'IA'}"
            + (" · análise local e privada" if engine.startswith("ollama-") else "")
        )
        self.subtitle.setStyleSheet("color: #60758a; font-size: 10px;")
        layout.addWidget(self.subtitle)

        answer_parts = [
            str(candidate.get("resumo", "") or "").strip()
            or f"Interpretei “{current_description}” como {str(candidate.get('descricao_normalizada', '') or current_description)}.",
        ]
        justification = str(candidate.get("justificacao", "") or "").strip()
        recommendation = str(candidate.get("recomendacao", "") or "").strip()
        if justification:
            answer_parts.extend(["", f"Porque proponho esta classificação:\n{justification}"])
        if recommendation:
            answer_parts.extend(["", f"O que deves confirmar:\n{recommendation}"])
        if sources:
            answer_parts.extend(
                ["", "Fontes:"]
                + [f"• {row.get('title', '')}: {row.get('url', '')}" for row in sources[:5]]
            )
        elif engine.startswith("ollama-"):
            answer_parts.extend(
                ["", "Esta análise foi feita localmente e não incluiu pesquisa na Internet."]
            )
        self.answer = QPlainTextEdit()
        self.answer.setReadOnly(True)
        self.answer.setPlainText("\n".join(answer_parts))
        self.answer.setMinimumHeight(175)
        self.answer.setStyleSheet(
            "QPlainTextEdit { background: #f7f9f8; color: #203448; border: 1px solid #cbd4cf;"
            "padding: 10px; font-size: 11px; }"
        )
        layout.addWidget(self.answer)

        proposal = QFrame()
        proposal.setStyleSheet(
            "QFrame { background: white; border: 1px solid #cbd4cf; }"
            "QLabel { border: none; color: #203448; }"
        )
        proposal_layout = QGridLayout(proposal)
        proposal_layout.setContentsMargins(14, 12, 14, 12)
        proposal_layout.setHorizontalSpacing(16)
        proposal_layout.setVerticalSpacing(8)
        fields = [
            ("Descrição proposta", candidate.get("descricao_normalizada") or current_description),
            ("Categoria", category or "-"),
            ("Subcategoria", subcategory or "-"),
            ("Tipo", product_type or "-"),
            ("Fabricante", candidate.get("fabricante") or "-"),
            ("Modelo", candidate.get("modelo") or "-"),
            ("Dimensões", candidate.get("dimensoes") or "-"),
        ]
        self.field_value_labels: dict[str, QLabel] = {}
        for index, (label_text, value) in enumerate(fields):
            row, column = divmod(index, 2)
            host = QWidget()
            host_layout = QVBoxLayout(host)
            host_layout.setContentsMargins(0, 0, 0, 0)
            host_layout.setSpacing(2)
            label = QLabel(label_text.upper())
            label.setStyleSheet("color: #6c7e8f; font-size: 8px; font-weight: 800;")
            content = QLabel(str(value or "-"))
            content.setWordWrap(True)
            content.setStyleSheet("font-size: 11px; font-weight: 700;")
            self.field_value_labels[label_text] = content
            host_layout.addWidget(label)
            host_layout.addWidget(content)
            proposal_layout.addWidget(host, row, column)
        proposal_layout.setColumnStretch(0, 1)
        proposal_layout.setColumnStretch(1, 1)
        layout.addWidget(proposal)

        conversation_label = QLabel("Queres corrigir ou aprofundar esta análise?")
        conversation_label.setStyleSheet("font-size: 10px; font-weight: 800; color: #203448;")
        layout.addWidget(conversation_label)
        question_row = QHBoxLayout()
        self.question_edit = QPlainTextEdit()
        self.question_edit.setPlaceholderText(
            "Ex.: Isto não deve ser classificado como periférico de computador?"
        )
        self.question_edit.setMaximumHeight(66)
        self.question_edit.setStyleSheet(
            "QPlainTextEdit { background: white; border: 1px solid #aebbb4;"
            "padding: 7px; font-size: 10px; }"
        )
        self.ask_button = QPushButton("Perguntar à IA")
        self.ask_button.setMinimumSize(138, 42)
        self.ask_button.setStyleSheet(
            "background: #454a46; color: white; border: none; padding: 8px 12px; font-weight: 800;"
        )
        self.ask_button.clicked.connect(self._ask_ai)
        question_row.addWidget(self.question_edit, 1)
        question_row.addWidget(self.ask_button)
        layout.addLayout(question_row)

        note = QLabel(
            "Nada é guardado automaticamente. Aplicar apenas preenche a ficha; "
            "Ensinar também memoriza esta associação para produtos semelhantes."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #60758a; font-size: 10px;")
        layout.addWidget(note)

        actions = QHBoxLayout()
        self.cancel_button = QPushButton("Cancelar")
        self.teach_button = QPushButton("Criar estrutura / ensinar")
        self.apply_button = QPushButton("Aplicar proposta")
        for button in (self.cancel_button, self.teach_button, self.apply_button):
            button.setMinimumHeight(38)
            button.setMinimumWidth(145)
        self.teach_button.setStyleSheet(
            "background: #454a46; color: white; border: none; padding: 8px 14px; font-weight: 800;"
        )
        self.apply_button.setStyleSheet(
            "background: #6dcc12; color: white; border: none; padding: 8px 14px; font-weight: 800;"
        )
        self.cancel_button.clicked.connect(self.reject)
        self.teach_button.clicked.connect(lambda: self._finish("teach"))
        self.apply_button.clicked.connect(lambda: self._finish("apply"))
        actions.addStretch(1)
        actions.addWidget(self.cancel_button)
        actions.addWidget(self.teach_button)
        actions.addWidget(self.apply_button)
        layout.addLayout(actions)
        self._render_result()

    def _render_result(self) -> None:
        candidate = dict(self.result.get("candidate", {}) or {})
        sources = list(self.result.get("sources", []) or [])
        engine = str(self.result.get("engine", "") or "").strip()
        confidence = int(round(float(candidate.get("confidence", 0) or 0) * 100))
        self.confidence_label.setText(f"CONFIANÇA {confidence}%")
        self.subtitle.setText(
            f"Motor: {engine.removeprefix('ollama-') if engine.startswith('ollama-') else engine or 'IA'}"
            + (" · análise local e privada" if engine.startswith("ollama-") else "")
        )
        answer_parts = [
            str(candidate.get("resumo", "") or "").strip()
            or f"Interpretei “{self.current_description}” como "
            f"{str(candidate.get('descricao_normalizada', '') or self.current_description)}.",
        ]
        justification = str(candidate.get("justificacao", "") or "").strip()
        recommendation = str(candidate.get("recomendacao", "") or "").strip()
        if justification:
            answer_parts.extend(["", f"Porque proponho esta classificação:\n{justification}"])
        if recommendation:
            answer_parts.extend(["", f"O que deves confirmar:\n{recommendation}"])
        if sources:
            answer_parts.extend(
                ["", "Fontes:"]
                + [f"• {row.get('title', '')}: {row.get('url', '')}" for row in sources[:5]]
            )
        elif engine.startswith("ollama-"):
            answer_parts.extend(
                ["", "Esta análise foi feita localmente e não incluiu pesquisa na Internet."]
            )
        self.answer.setPlainText("\n".join(answer_parts))
        values = {
            "Descrição proposta": candidate.get("descricao_normalizada") or self.current_description,
            "Categoria": candidate.get("categoria") or "-",
            "Subcategoria": candidate.get("subcat") or "-",
            "Tipo": candidate.get("tipo") or "-",
            "Fabricante": candidate.get("fabricante") or "-",
            "Modelo": candidate.get("modelo") or "-",
            "Dimensões": candidate.get("dimensoes") or "-",
        }
        for label_text, value in values.items():
            label = self.field_value_labels.get(label_text)
            if label is not None:
                label.setText(str(value or "-"))

    def _ask_ai(self) -> None:
        question = self.question_edit.toPlainText().strip()
        if not question:
            QMessageBox.information(self, "Copiloto LuGEST", "Escreve primeiro a tua pergunta.")
            return
        if not callable(self.refine_callback) or self._refine_thread is not None:
            return
        self.ask_button.setEnabled(False)
        self.ask_button.setText("A responder…")
        self.question_edit.setEnabled(False)
        self.cancel_button.setEnabled(False)
        self.teach_button.setEnabled(False)
        self.apply_button.setEnabled(False)
        self._pending_question = question
        worker = _ProductAILookupThread(
            lambda _value: self.refine_callback(question, self.result),
            question,
            self,
        )
        self._refine_thread = worker
        worker.completed.connect(self._refinement_completed)
        worker.failed.connect(self._refinement_failed)
        worker.finished.connect(self._refinement_finished)
        worker.start()

    def _refinement_completed(self, result: dict) -> None:
        previous_conversation = self.answer.toPlainText().strip()
        self.result = dict(result or {})
        self.question_edit.clear()
        self._render_result()
        latest_answer = self.answer.toPlainText().strip()
        self.answer.setPlainText(
            previous_conversation
            + "\n\n────────────────────────────────\n"
            + f"Tu:\n{self._pending_question}\n\nCopiloto:\n{latest_answer}"
        )
        scrollbar = self.answer.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _refinement_failed(self, message: str) -> None:
        QMessageBox.warning(self, "Conversa com IA", message or "Não foi possível responder.")

    def _refinement_finished(self) -> None:
        self.ask_button.setEnabled(True)
        self.ask_button.setText("Perguntar à IA")
        self.question_edit.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.teach_button.setEnabled(True)
        self.apply_button.setEnabled(True)
        worker = self._refine_thread
        self._refine_thread = None
        self._pending_question = ""
        if worker is not None:
            worker.deleteLater()

    def _finish(self, action: str) -> None:
        self.action = action
        self.accept()

    def reject(self) -> None:
        if self._refine_thread is not None:
            QMessageBox.information(
                self,
                "Copiloto LuGEST",
                "Aguarda pela resposta da IA antes de fechar esta janela.",
            )
            return
        super().reject()


def _product_grid_search_normalize(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.casefold()
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^a-z0-9./-]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _product_grid_numeric_key(value: object) -> str:
    text = _product_grid_search_normalize(value).strip().strip("./-")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return ""
    try:
        number = float(text)
    except Exception:
        return ""
    return f"{number:.4f}".rstrip("0").rstrip(".")


def _product_grid_search_matches(values: list[object], query: object) -> bool:
    terms = [term for term in _product_grid_search_normalize(query).split() if term]
    if not terms:
        return True
    normalized = _product_grid_search_normalize(" ".join(str(value or "") for value in values))
    tokens = set(normalized.split())
    numeric_tokens = {_product_grid_numeric_key(token) for token in tokens}
    numeric_tokens.discard("")
    for token in list(tokens):
        if _product_grid_numeric_key(token):
            continue
        for part in re.findall(r"\d+(?:\.\d+)?", token):
            part_numeric = _product_grid_numeric_key(part)
            if part_numeric:
                numeric_tokens.add(part_numeric)
                tokens.add(part)
    for term in terms:
        numeric = _product_grid_numeric_key(term)
        if numeric:
            if numeric not in numeric_tokens and term not in tokens:
                return False
            continue
        if term not in normalized:
            return False
    return True


class _ConsumeDialog(QDialog):
    def __init__(self, codigo: str, operators: list[str] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Baixa de produto {codigo}")
        self.setModal(True)
        self.resize(470, 250)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.qty_edit = QLineEdit("1")
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Baixa de stock", "Entrega a operador"])
        self.operator_combo = QComboBox()
        self.operator_combo.setEditable(True)
        self.operator_combo.addItem("")
        seen: set[str] = set()
        for raw_name in list(operators or []):
            name = str(raw_name or "").strip()
            key = name.lower()
            if not name or key in seen:
                continue
            seen.add(key)
            self.operator_combo.addItem(name)
        self.operator_label = QLabel("Responsável")
        self.help_label = QLabel("")
        self.help_label.setWordWrap(True)
        self.help_label.setStyleSheet("color: #5b6f86; font-size: 11px;")
        self.obs_edit = QLineEdit()
        form.addRow("Quantidade", self.qty_edit)
        form.addRow("Destino", self.mode_combo)
        form.addRow(self.operator_label, self.operator_combo)
        form.addRow("Observacoes", self.obs_edit)
        layout.addLayout(form)
        layout.addWidget(self.help_label)
        self.mode_combo.currentIndexChanged.connect(self._sync_state)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._sync_state()

    def _sync_state(self) -> None:
        issue_to_operator = self.mode_combo.currentIndex() == 1
        if issue_to_operator:
            self.operator_label.setText("Operador destino")
            self.help_label.setText("Seleciona quem recebe o produto. Nesta opção o operador é obrigatório.")
        else:
            self.operator_label.setText("Responsável")
            self.help_label.setText("Opcional: podes indicar quem fez a baixa de stock para o movimento ficar identificado.")
        self.operator_combo.setEnabled(True)
        line_edit = self.operator_combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText(
                "Seleciona quem recebe o produto" if issue_to_operator else "Opcional para registar o responsável"
            )

    def values(self) -> dict[str, str]:
        return {
            "qtd": self.qty_edit.text().strip(),
            "obs": self.obs_edit.text().strip(),
            "mode": "operator" if self.mode_combo.currentIndex() == 1 else "stock",
            "operator": self.operator_combo.currentText().strip(),
        }


class _AssembliesCatalogDialog(QDialog):
    TECHNICAL_FIELDS = (
        ("familia_produto", "Família"),
        ("aplicacao", "Aplicação"),
        ("modelo_versao", "Modelo / versão"),
        ("configuracao", "Configuração"),
        ("dimensoes_gerais", "Dimensões"),
        ("materiais_acabamentos", "Materiais / acabamentos"),
        ("caracteristicas", "Características"),
        ("requisitos_instalacao", "Instalação"),
        ("normas_conformidade", "Normas"),
        ("controlo_qualidade", "Controlo de qualidade"),
    )

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows: list[dict] = []
        self.current_code = ""
        self.setWindowTitle("Catálogo de conjuntos parametrizados")
        self.setWindowFlag(Qt.WindowMinMaxButtonsHint, True)
        self.resize(1380, 760)
        self.setMinimumSize(1080, 640)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        header = CardFrame()
        header.set_tone("info")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 9, 12, 9)
        header_layout.setSpacing(8)
        heading = QVBoxLayout()
        heading.setContentsMargins(0, 0, 0, 0)
        heading.setSpacing(1)
        title = QLabel("Conjuntos como produto")
        title.setStyleSheet("font-family: 'Segoe UI'; font-size: 15px; font-weight: 800; color: #102a43;")
        subtitle = QLabel("Catálogo ligado aos componentes, à parametrização e aos preços atuais do stock.")
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        subtitle.setMinimumWidth(0)
        subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        heading.addWidget(title)
        heading.addWidget(subtitle)
        header_layout.addLayout(heading, 1)
        self.count_label = QLabel("0 conjuntos")
        self.count_label.setProperty("role", "state_chip")
        self.linked_label = QLabel("0 ligações")
        self.linked_label.setProperty("role", "state_chip")
        self.value_label = QLabel("0,00 EUR")
        self.value_label.setProperty("role", "state_chip")
        header_layout.addWidget(self.count_label)
        header_layout.addWidget(self.linked_label)
        header_layout.addWidget(self.value_label)
        root.addWidget(header)

        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(6)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Pesquisar código, parametrização ou descrição")
        self.search_edit.textChanged.connect(self.refresh)
        refresh_btn = QPushButton("Atualizar preços")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self._refresh_prices)
        create_btn = QPushButton("Criar Parametrização")
        create_btn.setProperty("variant", "secondary")
        create_btn.clicked.connect(self._create_parameterization)
        self.pdf_btn = QPushButton("Ficha técnica PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_pdf)
        toolbar.addWidget(self.search_edit, 1)
        toolbar.addWidget(create_btn)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(self.pdf_btn)
        root.addLayout(toolbar)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(7)

        list_card = CardFrame()
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(10, 9, 10, 10)
        list_layout.setSpacing(6)
        list_title = QLabel("Portefólio")
        list_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #102a43;")
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["Param.", "Código", "Descrição", "Itens", "Ligados", "Margem", "Preço final"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(2, QHeaderView.Stretch)
        for column in (3, 4, 5, 6):
            header_view.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        self.table.itemSelectionChanged.connect(self._load_selected)
        self.table.itemDoubleClicked.connect(lambda *_args: self._open_pdf())
        list_layout.addWidget(list_title)
        list_layout.addWidget(self.table)
        splitter.addWidget(list_card)

        detail_card = CardFrame()
        detail_card.set_tone("default")
        detail_card.setMinimumWidth(560)
        detail_card.setMaximumWidth(16777215)
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(10, 9, 10, 10)
        detail_layout.setSpacing(7)
        detail_heading = QHBoxLayout()
        detail_text = QVBoxLayout()
        detail_text.setSpacing(1)
        self.detail_title = QLabel("Seleciona um conjunto")
        self.detail_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #102a43;")
        self.detail_title.setWordWrap(True)
        self.detail_title.setMinimumWidth(0)
        self.detail_title.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.detail_subtitle = QLabel("A parametrização e os componentes aparecem aqui.")
        self.detail_subtitle.setProperty("role", "muted")
        self.detail_subtitle.setWordWrap(True)
        self.detail_subtitle.setMinimumWidth(0)
        self.detail_subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        detail_text.addWidget(self.detail_title)
        detail_text.addWidget(self.detail_subtitle)
        self.param_label = QLabel("PARAM. ----")
        self.param_label.setAlignment(Qt.AlignCenter)
        self.param_label.setStyleSheet(
            "background: #e9f7f5; border: 1px solid #9fd4cd; border-radius: 6px;"
            " color: #176b63; padding: 5px 8px; font-size: 9px; font-weight: 900;"
        )
        detail_heading.addLayout(detail_text, 1)
        detail_heading.addWidget(self.param_label, 0, Qt.AlignTop)
        detail_layout.addLayout(detail_heading)

        metrics = QGridLayout()
        metrics.setContentsMargins(0, 0, 0, 0)
        metrics.setHorizontalSpacing(6)
        metrics.setVerticalSpacing(6)
        self.cost_label = QLabel("0,00 EUR")
        self.final_label = QLabel("0,00 EUR")
        self.margin_label = QLabel("0,00 %")
        for index, (caption, value, accent) in enumerate(
            (
                ("CUSTO ATUAL", self.cost_label, "#64748b"),
                ("PREÇO FINAL", self.final_label, "#0aa6a6"),
                ("MARGEM", self.margin_label, "#c58a20"),
            )
        ):
            metric = QFrame()
            metric.setStyleSheet(
                f"QFrame {{ background: #ffffff; border: 1px solid #d6e0e8;"
                f" border-top: 3px solid {accent}; border-radius: 6px; }}"
                " QFrame QLabel { border: 0; background: transparent; }"
            )
            metric_layout = QVBoxLayout(metric)
            metric_layout.setContentsMargins(7, 5, 7, 6)
            metric_layout.setSpacing(1)
            label = QLabel(caption)
            label.setStyleSheet("font-size: 8px; font-weight: 800; color: #667085;")
            value.setStyleSheet("font-size: 10px; font-weight: 800; color: #102a43;")
            metric_layout.addWidget(label)
            metric_layout.addWidget(value)
            metrics.addWidget(metric, 0, index)
            metrics.setColumnStretch(index, 1)
        detail_layout.addLayout(metrics)

        self.detail_tabs = QTabWidget()
        self.detail_tabs.setDocumentMode(True)
        self.detail_tabs.setStyleSheet(
            """
            QTabWidget::pane { border: 1px solid #cbd8e5; border-radius: 7px; background: #ffffff; top: -1px; }
            QTabBar::tab { min-width: 0; min-height: 28px; padding: 5px 5px; margin: 0;
                border: 1px solid #cbd8e5; background: #edf3f7; color: #475467;
                font-size: 9px; font-weight: 800; }
            QTabBar::tab:selected { background: #ffffff; color: #0b6868;
                border-top: 2px solid #0aa6a6; border-bottom-color: #ffffff; }
            """
        )
        self.detail_tabs.tabBar().setExpanding(True)
        self.detail_tabs.tabBar().setUsesScrollButtons(False)

        bom_page = QWidget()
        bom_layout = QVBoxLayout(bom_page)
        bom_layout.setContentsMargins(7, 7, 7, 7)
        self.bom_table = QTableWidget(0, 6)
        self.bom_table.setHorizontalHeaderLabels(["Tipo", "Referência", "Descrição", "Qtd", "Preço", "Origem"])
        self.bom_table.verticalHeader().setVisible(False)
        self.bom_table.verticalHeader().setDefaultSectionSize(25)
        self.bom_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bom_table.setSelectionBehavior(QTableWidget.SelectRows)
        bom_header = self.bom_table.horizontalHeader()
        bom_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        bom_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        bom_header.setSectionResizeMode(2, QHeaderView.Stretch)
        for column in (3, 4, 5):
            bom_header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        bom_layout.addWidget(self.bom_table)

        technical_page = QWidget()
        technical_layout = QVBoxLayout(technical_page)
        technical_layout.setContentsMargins(7, 7, 7, 7)
        self.technical_table = QTableWidget(0, 2)
        self.technical_table.setHorizontalHeaderLabels(["Característica", "Especificação"])
        self.technical_table.verticalHeader().setVisible(False)
        self.technical_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.technical_table.setSelectionMode(QTableWidget.NoSelection)
        self.technical_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.technical_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        technical_layout.addWidget(self.technical_table)

        self.detail_tabs.addTab(bom_page, "Componentes")
        self.detail_tabs.addTab(technical_page, "Ficha técnica")
        detail_layout.addWidget(self.detail_tabs, 1)
        splitter.addWidget(detail_card)
        splitter.setSizes([720, 660])
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 5)
        root.addWidget(splitter, 1)

        close_buttons = QDialogButtonBox(QDialogButtonBox.Close)
        close_buttons.button(QDialogButtonBox.Close).setText("Fechar")
        close_buttons.rejected.connect(self.reject)
        root.addWidget(close_buttons)
        self.refresh()

    def _fmt_eur(self, value: object) -> str:
        try:
            number = float(value or 0)
        except Exception:
            number = 0.0
        return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")

    def refresh(self) -> None:
        previous = self.current_code
        try:
            self.rows = list(self.backend.conjunto_rows(self.search_edit.text().strip()) or [])
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return
        self.table.setRowCount(len(self.rows))
        for row_index, row in enumerate(self.rows):
            values = (
                row.get("param_codigo", "-"),
                row.get("codigo", "-"),
                row.get("descricao", "-"),
                str(int(row.get("itens", 0) or 0)),
                str(int(row.get("itens_ligados", 0) or 0)),
                f"{float(row.get('margem_perc', 0) or 0):.2f} %",
                self._fmt_eur(row.get("total_final", 0)),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or "-"))
                if column != 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, column, item)
        linked = sum(int(row.get("itens_ligados", 0) or 0) for row in self.rows)
        total = sum(float(row.get("total_final", 0) or 0) for row in self.rows)
        self.count_label.setText(f"{len(self.rows)} conjuntos")
        self.linked_label.setText(f"{linked} ligações")
        self.value_label.setText(self._fmt_eur(total))
        if not self.rows:
            self.current_code = ""
            self._render_detail({})
            return
        target = next(
            (index for index, row in enumerate(self.rows) if str(row.get("codigo", "") or "") == previous),
            0,
        )
        self.table.selectRow(target)
        self._load_selected()

    def _selected_row(self) -> dict:
        current = self.table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def _load_selected(self) -> None:
        row = self._selected_row()
        code = str(row.get("codigo", "") or "").strip()
        if not code:
            self._render_detail({})
            return
        try:
            detail = dict(self.backend.conjunto_detail(code) or {})
        except Exception as exc:
            QMessageBox.warning(self, "Conjuntos", str(exc))
            return
        self.current_code = code
        self._render_detail(detail)

    def _item_type(self, item: dict) -> str:
        if str(item.get("produto_codigo", "") or "").strip():
            return "Produto"
        if str(item.get("stock_material_id", "") or "").strip():
            return "Matéria-prima"
        if str(item.get("operacao", "") or "").strip():
            return "Peça fabricada"
        return "Serviço"

    def _render_detail(self, detail: dict) -> None:
        code = str(detail.get("codigo", "") or "").strip()
        description = str(detail.get("descricao", "") or "").strip()
        param = str(detail.get("param_codigo", "") or "").strip()
        self.detail_title.setText(description or "Seleciona um conjunto")
        self.detail_subtitle.setText(code or "A parametrização e os componentes aparecem aqui.")
        self.param_label.setText(f"PARAM. {param or '----'}")
        self.cost_label.setText(self._fmt_eur(detail.get("total_custo", 0)))
        self.final_label.setText(self._fmt_eur(detail.get("total_final", 0)))
        self.margin_label.setText(f"{float(detail.get('margem_perc', 0) or 0):.2f} %")
        items = list(detail.get("itens", []) or [])
        self.bom_table.setRowCount(len(items))
        for row_index, item in enumerate(items):
            reference = (
                str(item.get("produto_codigo", "") or "").strip()
                or str(item.get("stock_material_id", "") or "").strip()
                or str(item.get("ref_externa", "") or "").strip()
                or "-"
            )
            values = (
                self._item_type(item),
                reference,
                str(item.get("descricao", "") or "-"),
                f"{float(item.get('qtd', 0) or 0):.2f}",
                self._fmt_eur(item.get("preco_unit", 0)),
                str(item.get("pricing_source_label", "") or "Valor manual"),
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if column in (0, 1, 3, 4):
                    cell.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.bom_table.setItem(row_index, column, cell)
        sheet = dict(detail.get("ficha_tecnica", {}) or {})
        technical_rows = [
            (label, str(sheet.get(key, "") or "").strip())
            for key, label in self.TECHNICAL_FIELDS
            if str(sheet.get(key, "") or "").strip()
        ]
        self.technical_table.setRowCount(max(1, len(technical_rows)))
        if technical_rows:
            for row_index, (label, value) in enumerate(technical_rows):
                self.technical_table.setItem(row_index, 0, QTableWidgetItem(label))
                self.technical_table.setItem(row_index, 1, QTableWidgetItem(value))
        else:
            self.technical_table.setItem(0, 0, QTableWidgetItem("Ficha técnica"))
            self.technical_table.setItem(0, 1, QTableWidgetItem("Sem especificações registadas."))
        self.pdf_btn.setEnabled(bool(code))

    def _refresh_prices(self) -> None:
        try:
            self.backend.conjunto_refresh_prices(self.current_code)
        except Exception as exc:
            QMessageBox.warning(self, "Conjuntos", str(exc))
            return
        self.refresh()

    def _create_parameterization(self) -> None:
        owner = self.parent()
        main_window = owner.window() if isinstance(owner, QWidget) else None
        quote_page = getattr(main_window, "pages", {}).get("quotes") if main_window is not None else None
        if quote_page is None and main_window is not None:
            ensure_page = getattr(main_window, "_ensure_page", None)
            if callable(ensure_page):
                try:
                    quote_page = ensure_page("quotes")
                except Exception as exc:
                    QMessageBox.warning(self, "Criar Parametrização", str(exc))
                    return
        builder = getattr(quote_page, "_calculated_assembly_builder_dialog", None)
        if not callable(builder):
            QMessageBox.warning(
                self,
                "Criar Parametrização",
                "O editor completo de parametrizações não está disponível nesta sessão.",
            )
            return
        payload = builder()
        if not payload:
            return
        self.current_code = str(payload.get("assembly_code", "") or payload.get("codigo", "") or "").strip()
        self.refresh()

    def _open_pdf(self) -> None:
        if not self.current_code:
            return
        try:
            self.backend.conjunto_open_sheet_pdf(self.current_code)
        except Exception as exc:
            QMessageBox.warning(self, "Ficha técnica", str(exc))


class ProductsPage(QWidget):
    page_title = "Produtos"
    page_subtitle = "Portefólio, disponibilidade, preços e rastreabilidade do produto acabado."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.current_code = ""
        self._moves_years: list[str] = []
        self._columns_compact: bool | None = None
        self._auto_catalog_values: dict[str, str] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(7)

        top_card = CardFrame()
        top_card.set_tone("info")
        top_card.setObjectName("ProductPortfolio")
        top_card.setStyleSheet(
            "QFrame#ProductMetric { background: #ffffff; border: 1px solid #c6d5e5; }"
            "QFrame#ProductValueMetric { background: #e8f6f4; border: 1px solid #9fd8d3; }"
            "QLabel#ProductMetricLabel { color: #5b7088; font-size: 8px; font-weight: 700; }"
            "QLabel#ProductMetricValue { color: #10253d; font-size: 13px; font-weight: 800; }"
            "QLabel#ProductValueMetricLabel { color: #087f83; font-size: 8px; font-weight: 700; }"
            "QLabel#ProductValueMetricValue { color: #087f83; font-size: 13px; font-weight: 800; }"
        )
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 10, 14, 10)
        top_layout.setSpacing(9)

        portfolio_row = QHBoxLayout()
        portfolio_row.setSpacing(10)
        title_wrap = QVBoxLayout()
        title_wrap.setContentsMargins(0, 0, 0, 0)
        title_wrap.setSpacing(3)
        title = QLabel("Portefólio de Produtos")
        title.setMinimumHeight(23)
        title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        title.setStyleSheet("font-size: 18px; font-weight: 850; color: #10253d;")
        subtitle = QLabel("Stock, preços, disponibilidade e rastreabilidade numa única área de trabalho.")
        subtitle.setProperty("role", "muted")
        subtitle.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        subtitle.setStyleSheet("font-size: 11px; color: #5b6675;")
        subtitle.setWordWrap(True)
        subtitle.setMinimumWidth(0)
        subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        title_wrap.addWidget(title)
        title_wrap.addWidget(subtitle)
        portfolio_row.addLayout(title_wrap, 1)
        portfolio_row.setAlignment(title_wrap, Qt.AlignVCenter)

        def metric_widget(label_text: str) -> tuple[QFrame, QLabel]:
            frame = QFrame()
            frame.setObjectName("ProductMetric")
            frame.setFixedSize(142, 52)
            layout = QVBoxLayout(frame)
            layout.setContentsMargins(10, 6, 10, 6)
            layout.setSpacing(0)
            label = QLabel(label_text.upper())
            label.setObjectName("ProductMetricLabel")
            value = QLabel("-")
            value.setObjectName("ProductMetricValue")
            value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            layout.addWidget(label)
            layout.addWidget(value)
            return frame, value

        total_metric, self.total_products_label = metric_widget("Produtos")
        stock_metric, self.products_in_stock_label = metric_widget("Com stock")
        value_metric, self.portfolio_value_label = metric_widget("Valor em stock")
        value_metric.setObjectName("ProductValueMetric")
        value_metric.findChildren(QLabel)[0].setObjectName("ProductValueMetricLabel")
        self.portfolio_value_label.setObjectName("ProductValueMetricValue")
        portfolio_row.addWidget(total_metric)
        portfolio_row.addWidget(stock_metric)
        portfolio_row.addWidget(value_metric)
        top_layout.addLayout(portfolio_row)

        self._product_filter_timer = QTimer(self)
        self._product_filter_timer.setSingleShot(True)
        self._product_filter_timer.timeout.connect(self.refresh)
        filter_row = QHBoxLayout()
        filter_row.setSpacing(7)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar código, descrição ou picar etiqueta...")
        self.filter_edit.setMinimumWidth(250)
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(lambda _text: self._product_filter_timer.start(180))
        self.filter_edit.returnPressed.connect(self._select_scanned_product)
        filter_row.addWidget(self.filter_edit, 2)
        self.filter_category_combo = QComboBox()
        self.filter_category_combo.setEditable(False)
        self.filter_category_combo.setProperty("allLabel", "Categoria: Todas")
        self.filter_category_combo.setMinimumWidth(138)
        self.filter_category_combo.setToolTip("Filtrar por categoria")
        filter_row.addWidget(self.filter_category_combo)
        self.filter_subcat_combo = QComboBox()
        self.filter_subcat_combo.setEditable(False)
        self.filter_subcat_combo.setProperty("allLabel", "Subcat.: Todas")
        self.filter_subcat_combo.setMinimumWidth(138)
        self.filter_subcat_combo.setToolTip("Filtrar por subcategoria")
        filter_row.addWidget(self.filter_subcat_combo)
        self.filter_type_combo = QComboBox()
        self.filter_type_combo.setEditable(False)
        self.filter_type_combo.setProperty("allLabel", "Tipo: Todos")
        self.filter_type_combo.setMinimumWidth(128)
        self.filter_type_combo.setToolTip("Filtrar por tipo")
        filter_row.addWidget(self.filter_type_combo)
        self.filter_state_combo = QComboBox()
        self.filter_state_combo.setEditable(False)
        self.filter_state_combo.addItems(["Estado: Todos", "Disponivel", "Stock baixo", "Sem stock", "Qualidade"])
        self.filter_state_combo.setMinimumWidth(128)
        self.filter_state_combo.setToolTip("Filtrar por estado")
        filter_row.addWidget(self.filter_state_combo)
        self.only_stock_check = QCheckBox("Apenas com stock")
        self.only_stock_check.setChecked(False)
        self.only_stock_check.toggled.connect(self.refresh)
        filter_row.addWidget(self.only_stock_check)
        clear_filters_btn = QPushButton("Limpar")
        clear_filters_btn.setProperty("compact", "true")
        clear_filters_btn.setProperty("variant", "secondary")
        clear_filters_btn.setFixedWidth(72)
        clear_filters_btn.clicked.connect(self._clear_product_filters)
        filter_row.addWidget(clear_filters_btn)
        top_layout.addLayout(filter_row)
        root.addWidget(top_card)

        actions_card = CardFrame()
        actions_card.set_tone("default")
        actions_card.setMaximumHeight(50)
        actions = QHBoxLayout(actions_card)
        actions.setContentsMargins(10, 6, 10, 6)
        actions.setSpacing(6)
        self.new_btn = QPushButton("Novo")
        self.new_btn.setProperty("variant", "success")
        self.new_btn.clicked.connect(self._new_product)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.setProperty("variant", "success")
        self.save_btn.clicked.connect(self._save_product)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "destructive")
        self.remove_btn.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.remove_btn.clicked.connect(self._remove_product)
        self.consume_btn = QPushButton("Baixa")
        self.consume_btn.setProperty("variant", "secondary")
        self.consume_btn.clicked.connect(self._consume_product)
        self.pdf_btn = QPushButton("Documentos  ▾")
        self.pdf_btn.setProperty("variant", "secondary")
        self.stock_pdf_btn = QPushButton("Preview stock")
        self.stock_pdf_btn.setProperty("variant", "secondary")
        self.stock_pdf_btn.clicked.connect(self._open_stock_pdf)
        self.label_btn = QPushButton("Etiqueta")
        self.label_btn.setProperty("variant", "secondary")
        self.label_btn.clicked.connect(self._open_label_pdf)
        documents_menu = QMenu(self.pdf_btn)
        documents_menu.addAction("Ficha técnica PDF", self._open_pdf)
        documents_menu.addAction("Mapa de stock PDF", self._open_stock_pdf)
        documents_menu.addAction("Etiqueta do produto", self._open_label_pdf)
        self.pdf_btn.setMenu(documents_menu)
        self.stock_pdf_btn.hide()
        self.label_btn.hide()
        self.form_mode_btn = QPushButton("Ficha")
        self.form_mode_btn.setProperty("variant", "secondary")
        self.form_mode_btn.clicked.connect(self._show_form_page)
        self.moves_mode_btn = QPushButton("Movs.")
        self.moves_mode_btn.setProperty("variant", "secondary")
        self.moves_mode_btn.clicked.connect(self._show_moves_page)
        self.full_grid_btn = QPushButton("▦  Grelha")
        self.full_grid_btn.setProperty("variant", "secondary")
        self.full_grid_btn.clicked.connect(self.open_full_grid)
        self.assemblies_btn = QPushButton("Conjuntos")
        self.assemblies_btn.setProperty("variant", "secondary")
        self.assemblies_btn.setToolTip("Abrir o catálogo de conjuntos parametrizados ligados aos preços atuais.")
        self.assemblies_btn.clicked.connect(self._open_assemblies_catalog)
        for button, width in (
            (self.new_btn, 74),
            (self.save_btn, 86),
            (self.remove_btn, 88),
            (self.consume_btn, 92),
            (self.pdf_btn, 100),
            (self.stock_pdf_btn, 98),
            (self.label_btn, 86),
            (self.full_grid_btn, 82),
            (self.assemblies_btn, 100),
        ):
            button.setStyleSheet("font-size: 10px; font-weight: 700;")
            button.setFixedWidth(width)
        for button in (
            self.new_btn,
            self.consume_btn,
            self.pdf_btn,
            self.assemblies_btn,
        ):
            actions.addWidget(button)
        self.form_mode_btn.hide()
        self.moves_mode_btn.hide()
        actions.addStretch(1)
        actions.addWidget(self.full_grid_btn)
        actions.addWidget(self.save_btn)
        actions.addWidget(self.remove_btn)
        root.addWidget(actions_card)

        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_header = QHBoxLayout()
        table_title = QLabel("Catálogo de produtos")
        table_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #10253d;")
        self.table_count_label = QLabel("-")
        self.table_count_label.setProperty("role", "muted")
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(self.table_count_label)
        self.table = QTableWidget(0, 11)
        self.table.setObjectName("StockTable")
        self.table.setStyleSheet(
            "QTableWidget { gridline-color: #d7e1ec; selection-background-color: #eaf7da; selection-color: #26331d; }"
            "QTableWidget::item:selected {"
            " background: #eaf7da;"
            " color: #26331d;"
            " border-left: 3px solid #7ed321;"
            " border-bottom: 1px solid #d8e6c9;"
            "}"
        )
        self.table.setHorizontalHeaderLabels(
            [
                "Codigo",
                "Descricao",
                "Categoria",
                "Tipo",
                "Qtd",
                "Disponivel",
                "Alerta",
                "Preco/Unid.",
                "Valor Stock",
                "Atualizado",
                "Estado",
            ]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setWordWrap(False)
        header = self.table.horizontalHeader()
        header.setFixedHeight(34)
        # The description is the information users need to scan; keep Estado
        # compact instead of allowing the last column to consume the spare room.
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.resizeSection(0, 144)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col, width in ((2, 116), (3, 104), (4, 68), (5, 78), (6, 68), (7, 102), (8, 108), (9, 132), (10, 88)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        self.table.setColumnHidden(3, True)
        self.table.setColumnHidden(6, True)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        self.table.itemDoubleClicked.connect(lambda *_args: self._show_form_page())
        table_layout.addLayout(table_header)
        table_layout.addWidget(self.table)

        detail_host = CardFrame()
        detail_host.setObjectName("ProductInspector")
        detail_host.set_tone("default")
        detail_host.setMinimumWidth(480)
        detail_host.setMaximumWidth(680)
        detail_host_layout = QVBoxLayout(detail_host)
        detail_host_layout.setContentsMargins(12, 10, 12, 12)
        detail_host_layout.setSpacing(9)
        inspector_header = QHBoxLayout()
        inspector_title_wrap = QVBoxLayout()
        inspector_title_wrap.setSpacing(1)
        inspector_eyebrow = QLabel("PRODUTO SELECIONADO")
        inspector_eyebrow.setStyleSheet("color: #5b7088; font-size: 9px; font-weight: 700;")
        self.current_product_label = QLabel("Sem produto selecionado")
        self.current_product_label.setWordWrap(True)
        self.current_product_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #10253d;")
        self.inspector_meta_label = QLabel("Selecione uma linha para consultar ou editar.")
        self.inspector_meta_label.setProperty("role", "muted")
        self.inspector_meta_label.setWordWrap(True)
        self.inspector_meta_label.setStyleSheet("font-size: 11px; color: #5b7088;")
        inspector_title_wrap.addWidget(inspector_eyebrow)
        inspector_title_wrap.addWidget(self.current_product_label)
        inspector_title_wrap.addWidget(self.inspector_meta_label)
        self.inspector_state_label = QLabel("-")
        self.inspector_state_label.setAlignment(Qt.AlignCenter)
        self.inspector_state_label.setMinimumWidth(92)
        inspector_header.addLayout(inspector_title_wrap, 1)
        inspector_header.addWidget(self.inspector_state_label, 0, Qt.AlignTop)
        detail_host_layout.addLayout(inspector_header)

        summary_strip = QFrame()
        summary_strip.setObjectName("ProductSummaryStrip")
        summary_strip.setFixedHeight(58)
        summary_strip.setStyleSheet(
            "QFrame#ProductSummaryStrip { background: #f1f6fa; border: none; }"
            "QLabel#ProductSummaryLabel { color: #60758d; font-size: 9px; font-weight: 700; border: none; background: transparent; }"
            "QLabel#ProductSummaryValue { color: #10253d; font-size: 12px; font-weight: 800; border: none; background: transparent; }"
        )
        inspector_metrics = QHBoxLayout(summary_strip)
        inspector_metrics.setContentsMargins(10, 6, 10, 6)
        inspector_metrics.setSpacing(12)

        def inspector_metric(label_text: str) -> tuple[QWidget, QLabel]:
            item = QWidget()
            layout = QVBoxLayout(item)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(0)
            label = QLabel(label_text)
            label.setObjectName("ProductSummaryLabel")
            value = QLabel("-")
            value.setObjectName("ProductSummaryValue")
            layout.addWidget(label)
            layout.addWidget(value)
            return item, value

        qty_metric, self.inspector_qty_label = inspector_metric("Stock físico")
        available_metric, self.inspector_available_label = inspector_metric("Disponível")
        price_metric, self.price_unit_label = inspector_metric("Preço/unidade")
        stock_value_metric, self.stock_value_label = inspector_metric("Valor em stock")
        inspector_metrics.addWidget(qty_metric, 1)
        inspector_metrics.addWidget(available_metric, 1)
        inspector_metrics.addWidget(price_metric, 1)
        inspector_metrics.addWidget(stock_value_metric, 1)
        detail_host_layout.addWidget(summary_strip)

        self.detail_mode_label = QLabel("Ficha do produto")
        self.detail_mode_label.hide()
        self.detail_stack = QTabWidget()
        self.detail_stack.setDocumentMode(True)
        self.detail_stack.setUsesScrollButtons(False)
        self.detail_stack.setStyleSheet(
            "QTabWidget::pane { border: 1px solid #cbd8e5; background: #ffffff; top: -1px; }"
            "QTabBar::tab { min-width: 104px; min-height: 32px; padding: 0 9px; color: #4a6179; font-size: 10px; font-weight: 700; }"
            "QTabBar::tab:selected { background: #ffffff; color: #087f83; border-bottom: 2px solid #08a6a6; }"
        )
        detail_host_layout.addWidget(self.detail_stack, 1)

        self.form_page = QWidget()
        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.NoFrame)
        form_content = QWidget()
        form_scroll.setWidget(form_content)
        form_page_layout = QVBoxLayout(self.form_page)
        form_page_layout.setContentsMargins(0, 0, 0, 0)
        form_page_layout.addWidget(form_scroll)
        form_content_layout = QVBoxLayout(form_content)
        form_content_layout.setContentsMargins(10, 10, 10, 10)
        form_content_layout.setSpacing(8)
        form_grid = QGridLayout()
        form_grid.setHorizontalSpacing(8)
        form_grid.setVerticalSpacing(5)
        self.code_edit = QLineEdit()
        self.desc_edit = QLineEdit()
        self._catalog_suggestion_timer = QTimer(self)
        self._catalog_suggestion_timer.setSingleShot(True)
        self._catalog_suggestion_timer.setInterval(320)
        self._catalog_suggestion_timer.timeout.connect(self._apply_catalog_suggestion)
        self.desc_edit.textEdited.connect(lambda _text: self._catalog_suggestion_timer.start())
        self.desc_edit.editingFinished.connect(self._apply_catalog_suggestion)
        self.category_combo = self._make_combo()
        self.subcat_combo = self._make_combo()
        self.type_combo = self._make_combo()
        self.unit_combo = self._make_combo()
        self.dim_edit = QLineEdit()
        self.meters_edit = QLineEdit()
        self.weight_edit = QLineEdit()
        self.qty_edit = QLineEdit()
        self.alert_edit = QLineEdit()
        self.buy_price_edit = QLineEdit()
        self.pvp1_edit = QLineEdit()
        self.pvp2_edit = QLineEdit()
        self.maker_edit = QLineEdit()
        self.model_edit = QLineEdit()
        self.obs_edit = QLineEdit()

        def add_field(grid: QGridLayout, label_text: str, widget: QWidget, row: int, col: int, span: int = 1) -> None:
            label = QLabel(label_text)
            label.setStyleSheet("color: #5b7088; font-size: 9px; font-weight: 700;")
            grid.addWidget(label, row * 2, col, 1, span)
            grid.addWidget(widget, row * 2 + 1, col, 1, span)

        add_field(form_grid, "Código", self.code_edit, 0, 0)
        add_field(form_grid, "Unidade", self.unit_combo, 0, 1)
        add_field(form_grid, "Descrição", self.desc_edit, 1, 0, 2)
        add_field(form_grid, "Categoria", self.category_combo, 2, 0)
        add_field(form_grid, "Subcategoria", self.subcat_combo, 2, 1)
        add_field(form_grid, "Tipo", self.type_combo, 3, 0)
        add_field(form_grid, "Dimensões", self.dim_edit, 3, 1)
        add_field(form_grid, "Fabricante", self.maker_edit, 4, 0)
        add_field(form_grid, "Modelo", self.model_edit, 4, 1)
        add_field(form_grid, "Observações", self.obs_edit, 5, 0, 2)
        form_grid.setColumnStretch(0, 1)
        form_grid.setColumnStretch(1, 1)
        form_content_layout.addLayout(form_grid)
        self.catalog_suggestion_label = QLabel("")
        self.catalog_suggestion_label.setWordWrap(True)
        self.catalog_suggestion_label.setStyleSheet(
            "background: #f1f8e8; color: #456a13; border: 1px solid #c9e6a3;"
            "padding: 6px 8px; font-size: 10px; font-weight: 700;"
        )
        self.catalog_suggestion_label.hide()
        self.normalize_description_button = QPushButton("Normalizar")
        self.normalize_description_button.setMinimumWidth(112)
        self.normalize_description_button.setToolTip(
            "Aplica a designação limpa sugerida pelo Copiloto. A alteração só é gravada quando clicares em Guardar."
        )
        self.normalize_description_button.clicked.connect(self._apply_normalized_description)
        self.normalize_description_button.hide()
        self.teach_catalog_button = QPushButton("Criar estrutura / ensinar")
        self.teach_catalog_button.setMinimumWidth(148)
        self.teach_catalog_button.setToolTip(
            "Guarda a classificação atual e cria automaticamente a categoria, subcategoria "
            "ou tipo que ainda não existam."
        )
        self.teach_catalog_button.clicked.connect(self._teach_catalog_classification)
        self.teach_catalog_button.hide()
        self.product_ai_button = QPushButton("Pesquisar com IA")
        self.product_ai_button.setMinimumWidth(148)
        self.product_ai_button.setToolTip(
            "Pesquisa fabricante, referência e características através do serviço seguro LuGEST. "
            "Nenhuma sugestão é gravada sem confirmação."
        )
        self.product_ai_button.clicked.connect(self._lookup_product_with_ai)
        self.product_ai_button.hide()
        copilot_actions = QHBoxLayout()
        copilot_actions.setContentsMargins(0, 0, 0, 0)
        copilot_actions.setSpacing(6)
        copilot_actions.addStretch(1)
        copilot_actions.addWidget(self.normalize_description_button)
        copilot_actions.addWidget(self.teach_catalog_button)
        copilot_actions.addWidget(self.product_ai_button)
        form_content_layout.addWidget(self.catalog_suggestion_label)
        form_content_layout.addLayout(copilot_actions)
        form_content_layout.addStretch(1)

        self.stock_page = QWidget()
        stock_scroll = QScrollArea()
        stock_scroll.setWidgetResizable(True)
        stock_scroll.setFrameShape(QFrame.NoFrame)
        stock_content = QWidget()
        stock_scroll.setWidget(stock_content)
        stock_page_layout = QVBoxLayout(self.stock_page)
        stock_page_layout.setContentsMargins(0, 0, 0, 0)
        stock_page_layout.addWidget(stock_scroll)
        stock_content_layout = QVBoxLayout(stock_content)
        stock_content_layout.setContentsMargins(10, 10, 10, 10)
        stock_grid = QGridLayout()
        stock_grid.setHorizontalSpacing(8)
        stock_grid.setVerticalSpacing(5)
        add_field(stock_grid, "Quantidade física", self.qty_edit, 0, 0)
        add_field(stock_grid, "Alerta mínimo", self.alert_edit, 0, 1)
        add_field(stock_grid, "Metros/unidade", self.meters_edit, 1, 0)
        add_field(stock_grid, "Peso/unidade", self.weight_edit, 1, 1)
        add_field(stock_grid, "Preço de compra", self.buy_price_edit, 2, 0)
        add_field(stock_grid, "PVP 1", self.pvp1_edit, 2, 1)
        add_field(stock_grid, "PVP 2", self.pvp2_edit, 3, 0)
        stock_grid.setColumnStretch(0, 1)
        stock_grid.setColumnStretch(1, 1)
        stock_content_layout.addLayout(stock_grid)
        stock_content_layout.addStretch(1)

        self.moves_page = QWidget()
        moves_layout = QVBoxLayout(self.moves_page)
        moves_layout.setContentsMargins(10, 10, 10, 10)
        moves_layout.setSpacing(8)
        moves_filters = QHBoxLayout()
        moves_filters.setSpacing(6)
        self.moves_operator_combo = QComboBox()
        self.moves_operator_combo.setEditable(False)
        self.moves_operator_combo.currentIndexChanged.connect(self._refresh_moves_view)
        self.moves_year_combo = QComboBox()
        self.moves_year_combo.setEditable(False)
        self.moves_year_combo.currentIndexChanged.connect(self._refresh_moves_view)
        self.moves_summary_label = QLabel("Sem movimentos de operador no período selecionado.")
        self.moves_summary_label.setProperty("role", "muted")
        moves_filters.addWidget(QLabel("Operador"))
        moves_filters.addWidget(self.moves_operator_combo)
        moves_filters.addWidget(QLabel("Ano"))
        moves_filters.addWidget(self.moves_year_combo)
        self.moves_table = QTableWidget(0, 7)
        self.moves_table.setHorizontalHeaderLabels(["Data", "Tipo", "Operador", "Qtd", "Antes", "Depois", "Observacoes"])
        self.moves_table.verticalHeader().setVisible(False)
        self.moves_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.moves_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.moves_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.moves_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.moves_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.moves_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        for col in (3, 4, 5):
            self.moves_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        moves_layout.addLayout(moves_filters)
        moves_layout.addWidget(self.moves_summary_label)
        moves_layout.addWidget(self.moves_table)

        self.detail_stack.addTab(self.form_page, "Identificação")
        self.detail_stack.addTab(self.stock_page, "Stock e preços")
        self.detail_stack.addTab(self.moves_page, "Movimentos")

        workspace = QSplitter(Qt.Horizontal)
        workspace.setObjectName("ProductsWorkspace")
        workspace.setChildrenCollapsible(False)
        workspace.addWidget(table_card)
        workspace.addWidget(detail_host)
        workspace.setStretchFactor(0, 3)
        workspace.setStretchFactor(1, 1)
        workspace.setSizes([1320, 540])
        workspace.splitterMoved.connect(lambda *_args: self._resize_product_columns())
        root.addWidget(workspace, 1)
        QTimer.singleShot(0, self._resize_product_columns)

        for edit in (self.meters_edit, self.weight_edit, self.qty_edit, self.buy_price_edit):
            edit.textChanged.connect(self._refresh_price_labels)
        self.category_combo.currentTextChanged.connect(self._sync_form_catalog)
        self.subcat_combo.currentTextChanged.connect(self._sync_form_catalog)
        self.filter_category_combo.currentTextChanged.connect(self._sync_filter_catalog)
        self.filter_subcat_combo.currentTextChanged.connect(self._sync_filter_catalog)
        self.filter_type_combo.currentTextChanged.connect(lambda _text: self.refresh())
        self.filter_state_combo.currentTextChanged.connect(lambda _text: self.refresh())

        self._load_presets()
        self._new_product()
        self._show_form_page()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        QTimer.singleShot(0, self._resize_product_columns)

    def _resize_product_columns(self) -> None:
        if not hasattr(self, "table"):
            return
        viewport_width = self.table.viewport().width()
        compact = viewport_width < 1000
        if self._columns_compact is compact:
            return
        self._columns_compact = compact
        widths = (
            ((0, 108), (2, 88), (3, 82), (4, 52), (5, 64), (6, 56), (7, 82), (8, 92), (9, 108), (10, 78))
            if compact
            else ((0, 144), (2, 116), (3, 104), (4, 68), (5, 78), (6, 68), (7, 102), (8, 108), (9, 132), (10, 88))
        )
        header = self.table.horizontalHeader()
        for column, width in widths:
            header.resizeSection(column, width)

    def _make_combo(self) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        return combo

    def _set_combo_values(self, combo: QComboBox, values: list[str], current: str = "") -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(values)
        combo.setCurrentText(current)
        combo.blockSignals(False)

    def _set_filter_combo_values(self, combo: QComboBox, values: list[str], current: str = "") -> None:
        combo.blockSignals(True)
        combo.clear()
        all_label = str(combo.property("allLabel") or "Todas")
        combo.addItem(all_label)
        combo.addItems([value for value in values if str(value or "").strip()])
        available = [combo.itemText(index) for index in range(combo.count())]
        combo.setCurrentText(current if current in available else all_label)
        combo.blockSignals(False)

    def _clear_product_filters(self) -> None:
        self.filter_edit.clear()
        for combo in (
            self.filter_category_combo,
            self.filter_subcat_combo,
            self.filter_type_combo,
            self.filter_state_combo,
        ):
            combo.blockSignals(True)
            combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self.only_stock_check.blockSignals(True)
        self.only_stock_check.setChecked(False)
        self.only_stock_check.blockSignals(False)
        self.refresh()

    def _select_scanned_product(self) -> None:
        value = self.filter_edit.text().strip()
        if not value:
            return
        try:
            result = self.backend.inventory_scan_lookup(value, expected_type="PRD")
        except Exception as exc:
            if "|" in value or re.fullmatch(r"PRD[-_A-Z0-9]+", value, flags=re.IGNORECASE):
                QMessageBox.warning(self, "Picagem", str(exc))
            return
        self.current_code = str(result.get("entity_id", "") or "").strip()
        self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.blockSignals(False)
        for combo in (
            self.filter_category_combo,
            self.filter_subcat_combo,
            self.filter_type_combo,
            self.filter_state_combo,
        ):
            combo.blockSignals(True)
            combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self.only_stock_check.blockSignals(True)
        self.only_stock_check.setChecked(False)
        self.only_stock_check.blockSignals(False)
        self.refresh()
        self.table.setFocus()

    def _filter_combo_text(self, combo: QComboBox) -> str:
        if combo.currentIndex() <= 0:
            return ""
        text = str(combo.currentText() or "").strip()
        return "" if text.casefold() in {"", "todos", "todas", "all"} else text

    def _set_mode_buttons(self, moves: bool) -> None:
        self.form_mode_btn.setEnabled(moves)
        self.moves_mode_btn.setEnabled(not moves)
        self.detail_mode_label.setText("Movimentos" if moves else "Ficha do produto")

    def _set_inspector_state(self, text: str) -> None:
        normalized = _product_grid_search_normalize(text)
        background = "#e8f6f4"
        foreground = "#087f83"
        border = "#9fd8d3"
        if "rejeit" in normalized:
            background, foreground, border = "#fff0f0", "#b42318", "#f1b6b2"
        elif "sem stock" in normalized or "bloque" in normalized:
            background, foreground, border = "#fff8eb", "#b45f06", "#e4c37f"
        elif "baixo" in normalized or "pend" in normalized or "inspec" in normalized:
            background, foreground, border = "#fff7e6", "#9a5b00", "#efd29b"
        elif normalized in {"novo", "-"}:
            background, foreground, border = "#eef4fb", "#33516f", "#c4d4e5"
        self.inspector_state_label.setText(text or "-")
        self.inspector_state_label.setStyleSheet(
            f"background: {background}; color: {foreground}; border: 1px solid {border};"
            "padding: 5px 8px; font-size: 9px; font-weight: 800;"
        )

    def _show_form_page(self) -> None:
        self.detail_stack.setCurrentWidget(self.form_page)
        self._set_mode_buttons(False)

    def _show_moves_page(self) -> None:
        self.detail_stack.setCurrentWidget(self.moves_page)
        self._set_mode_buttons(True)

    def _load_presets(self) -> None:
        presets = self.backend.product_catalog_options(self.category_combo.currentText().strip(), self.subcat_combo.currentText().strip())
        self._set_combo_values(self.category_combo, presets.get("categorias", []), self.category_combo.currentText())
        self._set_combo_values(self.subcat_combo, presets.get("subcats", []), self.subcat_combo.currentText())
        self._set_combo_values(self.type_combo, presets.get("tipos", []), self.type_combo.currentText())
        self._set_combo_values(self.unit_combo, presets.get("unidades", []), self.unit_combo.currentText() or "UN")
        filter_presets = self.backend.product_catalog_options(
            self._filter_combo_text(self.filter_category_combo),
            self._filter_combo_text(self.filter_subcat_combo),
        )
        self._set_filter_combo_values(self.filter_category_combo, self.backend.product_catalog_options().get("categorias", []), self.filter_category_combo.currentText())
        self._set_filter_combo_values(self.filter_subcat_combo, filter_presets.get("subcats", []), self.filter_subcat_combo.currentText())
        self._set_filter_combo_values(self.filter_type_combo, filter_presets.get("tipos", []), self.filter_type_combo.currentText())

    def _sync_form_catalog(self) -> None:
        current_category = self.category_combo.currentText().strip()
        current_subcat = self.subcat_combo.currentText().strip()
        current_type = self.type_combo.currentText().strip()
        presets = self.backend.product_catalog_options(current_category, current_subcat)
        self._set_combo_values(self.subcat_combo, presets.get("subcats", []), current_subcat)
        selected_subcat = self.subcat_combo.currentText().strip()
        type_presets = self.backend.product_catalog_options(current_category, selected_subcat)
        self._set_combo_values(self.type_combo, type_presets.get("tipos", []), current_type)

    def _sync_filter_catalog(self) -> None:
        current_category = self._filter_combo_text(self.filter_category_combo)
        current_subcat = self._filter_combo_text(self.filter_subcat_combo)
        current_type = self.filter_type_combo.currentText()
        presets = self.backend.product_catalog_options(current_category, current_subcat)
        self._set_filter_combo_values(self.filter_subcat_combo, presets.get("subcats", []), self.filter_subcat_combo.currentText())
        selected_subcat = self._filter_combo_text(self.filter_subcat_combo)
        type_presets = self.backend.product_catalog_options(current_category, selected_subcat)
        self._set_filter_combo_values(self.filter_type_combo, type_presets.get("tipos", []), current_type)
        self.refresh()

    def _apply_catalog_suggestion(self) -> None:
        description = self.desc_edit.text().strip()
        suggester = getattr(self.backend, "product_copilot_analysis", None)
        if not callable(suggester):
            suggester = getattr(self.backend, "product_catalog_suggestion", None)
        if not description or not callable(suggester):
            self.catalog_suggestion_label.hide()
            self.normalize_description_button.hide()
            self.teach_catalog_button.hide()
            self.product_ai_button.hide()
            return
        try:
            suggestion = dict(suggester(description, self.code_edit.text().strip()) or {})
        except TypeError:
            suggestion = dict(suggester(description) or {})
        category = str(suggestion.get("categoria", "") or "").strip()
        subcategory = str(suggestion.get("subcat", "") or "").strip()
        product_type = str(suggestion.get("tipo", "") or "").strip()
        dimensions = str(suggestion.get("dimensoes", "") or "").strip()
        manufacturer = str(suggestion.get("fabricante", "") or "").strip()
        model = str(suggestion.get("modelo", "") or "").strip()
        if not category:
            self.catalog_suggestion_label.hide()
            self.normalize_description_button.hide()
            self.teach_catalog_button.hide()
            self.product_ai_button.setVisible(bool(description))
            return

        previous = dict(self._auto_catalog_values)
        current_category = self.category_combo.currentText().strip()
        may_set_category = not current_category or current_category == previous.get("categoria", "")
        if may_set_category:
            self.category_combo.setCurrentText(category)
            self._sync_form_catalog()
            current_subcategory = self.subcat_combo.currentText().strip()
            if subcategory and (not current_subcategory or current_subcategory == previous.get("subcat", "")):
                self.subcat_combo.setCurrentText(subcategory)
                self._sync_form_catalog()
            current_type = self.type_combo.currentText().strip()
            if product_type and (not current_type or current_type == previous.get("tipo", "")):
                self.type_combo.setCurrentText(product_type)
            current_dimensions = self.dim_edit.text().strip()
            if dimensions and (not current_dimensions or current_dimensions == previous.get("dimensoes", "")):
                self.dim_edit.setText(dimensions)
            current_manufacturer = self.maker_edit.text().strip()
            if manufacturer and (
                not current_manufacturer or current_manufacturer == previous.get("fabricante", "")
            ):
                self.maker_edit.setText(manufacturer)
            current_model = self.model_edit.text().strip()
            if model and (not current_model or current_model == previous.get("modelo", "")):
                self.model_edit.setText(model)
            self._auto_catalog_values = {
                "categoria": category,
                "subcat": subcategory,
                "tipo": product_type,
                "dimensoes": dimensions,
                "fabricante": manufacturer,
                "modelo": model,
            }

        applied_category = self.category_combo.currentText().strip()
        applied_subcategory = self.subcat_combo.currentText().strip()
        applied_type = self.type_combo.currentText().strip()
        confidence = int(suggestion.get("confidence_percent", 0) or 0)
        path = "  ›  ".join(value for value in (category, applied_subcategory, applied_type) if value)
        if applied_category == category:
            heading = f"Copiloto LuGEST{f' · {confidence}%' if confidence else ''}: {path}"
        else:
            heading = (
                f"Sugestão disponível{f' · {confidence}%' if confidence else ''}: {category}  ›  {subcategory}"
                + (f"  ›  {product_type}" if product_type else "")
                + ". A seleção manual foi mantida."
            )
        details: list[str] = [heading]
        attributes = dict(suggestion.get("atributos", {}) or {})
        if attributes:
            details.append(
                "Atributos: "
                + " · ".join(f"{str(key).replace('_', ' ').title()}: {value}" for key, value in attributes.items())
            )
        normalized = str(suggestion.get("descricao_normalizada", "") or "").strip()
        self._suggested_normalized_description = normalized
        if normalized and normalized.casefold() != description.casefold():
            details.append(f"Designação normalizada: {normalized}")
            self.normalize_description_button.show()
        else:
            self.normalize_description_button.hide()
        similar = list(suggestion.get("semelhantes", []) or [])
        if similar:
            details.append(
                "Possíveis semelhantes: "
                + " · ".join(
                    f"{row.get('codigo', '')} ({int(row.get('percent', 0) or 0)}%)"
                    for row in similar
                )
            )
        duplicate_warning = bool(suggestion.get("duplicate_warning"))
        needs_learning = bool(suggestion.get("needs_learning"))
        self.teach_catalog_button.setVisible(needs_learning)
        self.product_ai_button.setVisible(bool(description))
        if needs_learning:
            details.append(
                "Termo ainda não conhecido. Corrige a classificação, se necessário, e clica “Ensinar ao LuGEST”."
            )
        if duplicate_warning:
            self.catalog_suggestion_label.setStyleSheet(
                "background: #fff7e8; color: #8a5400; border: 1px solid #efc36f;"
                "padding: 6px 8px; font-size: 10px; font-weight: 700;"
            )
        else:
            self.catalog_suggestion_label.setStyleSheet(
                "background: #f1f8e8; color: #456a13; border: 1px solid #c9e6a3;"
                "padding: 6px 8px; font-size: 10px; font-weight: 700;"
            )
        self.catalog_suggestion_label.setText("\n".join(details))
        self.catalog_suggestion_label.show()

    def _lookup_product_with_ai(self) -> None:
        description = self.desc_edit.text().strip()
        lookup = getattr(self.backend, "product_ai_lookup", None)
        if not description:
            QMessageBox.information(self, "Copiloto LuGEST", "Escreve primeiro a descrição do produto.")
            return
        if not callable(lookup):
            QMessageBox.warning(self, "Copiloto LuGEST", "A pesquisa externa não está disponível nesta versão.")
            return
        if getattr(self, "_product_ai_thread", None) is not None:
            return
        self.product_ai_button.setEnabled(False)
        self.product_ai_button.setText("A pesquisar…")
        worker = _ProductAILookupThread(lookup, description, self)
        self._product_ai_thread = worker
        worker.completed.connect(self._product_ai_lookup_completed)
        worker.failed.connect(self._product_ai_lookup_failed)
        worker.finished.connect(self._product_ai_lookup_finished)
        worker.start()

    def _product_ai_lookup_finished(self) -> None:
        self.product_ai_button.setEnabled(True)
        self.product_ai_button.setText("Pesquisar com IA")
        worker = getattr(self, "_product_ai_thread", None)
        self._product_ai_thread = None
        if worker is not None:
            worker.deleteLater()

    def _product_ai_lookup_failed(self, message: str) -> None:
        QMessageBox.warning(
            self,
            "Pesquisa com IA",
            message or "Não foi possível concluir a pesquisa externa.",
        )

    def _product_ai_lookup_completed(self, result: dict) -> None:
        original_description = self.desc_edit.text().strip()

        def refine(question: str, previous_result: dict) -> dict:
            lookup = getattr(self.backend, "product_ai_lookup", None)
            if not callable(lookup):
                raise RuntimeError("A conversa com IA não está disponível nesta versão.")
            return dict(
                lookup(
                    original_description,
                    user_instruction=question,
                    previous_candidate=dict(previous_result.get("candidate", {}) or {}),
                )
                or {}
            )

        dialog = _ProductAIResultDialog(result, original_description, refine, self)
        if dialog.exec() != QDialog.Accepted or dialog.action == "cancel":
            return
        result = dict(dialog.result or {})
        candidate = dict(result.get("candidate", {}) or {})
        sources = list(result.get("sources", []) or [])
        category = str(candidate.get("categoria", "") or "").strip()
        subcategory = str(candidate.get("subcat", "") or "").strip()
        product_type = str(candidate.get("tipo", "") or "").strip()
        normalized = str(candidate.get("descricao_normalizada", "") or "").strip()
        manufacturer = str(candidate.get("fabricante", "") or "").strip()
        model = str(candidate.get("modelo", "") or "").strip()
        dimensions = str(candidate.get("dimensoes", "") or "").strip()
        confidence = int(round(float(candidate.get("confidence", 0) or 0) * 100))
        if category:
            self.category_combo.setCurrentText(category)
            self._sync_form_catalog()
        if subcategory:
            self.subcat_combo.setCurrentText(subcategory)
            self._sync_form_catalog()
        if product_type:
            self.type_combo.setCurrentText(product_type)
        if normalized:
            self.desc_edit.setText(normalized)
        if manufacturer:
            self.maker_edit.setText(manufacturer)
        if model:
            self.model_edit.setText(model)
        if dimensions:
            self.dim_edit.setText(dimensions)
        self._auto_catalog_values = {
            "categoria": category,
            "subcat": subcategory,
            "tipo": product_type,
            "dimensoes": dimensions,
            "fabricante": manufacturer,
            "modelo": model,
        }
        source_names = ", ".join(str(row.get("title", "") or "") for row in sources[:3] if row.get("title"))
        self.catalog_suggestion_label.setText(
            f"Proposta da IA aplicada · {confidence}%"
            + (f"\nFontes: {source_names}" if source_names else "\nAnálise local; confirma os dados antes de guardar.")
        )
        self.catalog_suggestion_label.show()
        if dialog.action == "teach":
            self._teach_catalog_classification(already_confirmed=True)

    def _apply_normalized_description(self) -> None:
        normalized = str(getattr(self, "_suggested_normalized_description", "") or "").strip()
        if not normalized:
            return
        self.desc_edit.setText(normalized)
        self._apply_catalog_suggestion()

    def _teach_catalog_classification(self, already_confirmed: bool = False) -> None:
        teacher = getattr(self.backend, "product_catalog_teach", None)
        if not callable(teacher):
            QMessageBox.warning(self, "Copiloto LuGEST", "A aprendizagem local não está disponível.")
            return
        description = self.desc_edit.text().strip()
        category = self.category_combo.currentText().strip()
        subcategory = self.subcat_combo.currentText().strip()
        product_type = self.type_combo.currentText().strip()
        if not description or not category or not subcategory:
            QMessageBox.warning(
                self,
                "Copiloto LuGEST",
                "Preenche a descrição, a categoria e a subcategoria antes de ensinar.",
            )
            return
        if not already_confirmed:
            answer = QMessageBox.question(
                self,
                "Ensinar ao LuGEST",
                "Guardar esta associação para os próximos produtos?\n\n"
                f"{description}\n"
                f"→ {category} › {subcategory}"
                + (f" › {product_type}" if product_type else ""),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes:
                return
        try:
            learned = dict(teacher(description, category, subcategory, product_type) or {})
        except Exception as exc:
            QMessageBox.critical(self, "Copiloto LuGEST", str(exc))
            return
        self._load_presets()
        self._apply_catalog_suggestion()
        created_parts = []
        if learned.get("created_category"):
            created_parts.append(f"categoria “{category}”")
        if learned.get("created_subcategory"):
            created_parts.append(f"subcategoria “{subcategory}”")
        if learned.get("created_type"):
            created_parts.append(f"tipo “{product_type}”")
        creation_message = (
            "\n\nEstrutura criada: " + ", ".join(created_parts) + "."
            if created_parts
            else ""
        )
        QMessageBox.information(
            self,
            "Copiloto LuGEST",
            f"Aprendido. O termo “{learned.get('keyword', '')}” será classificado automaticamente."
            + creation_message,
        )

    def _fmt_eur(self, value) -> str:
        try:
            number = float(value or 0)
        except Exception:
            number = 0.0
        return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")

    def _payload(self) -> dict[str, str]:
        return {
            "codigo": self.code_edit.text().strip(),
            "descricao": self.desc_edit.text().strip(),
            "categoria": self.category_combo.currentText().strip(),
            "subcat": self.subcat_combo.currentText().strip(),
            "tipo": self.type_combo.currentText().strip(),
            "unid": self.unit_combo.currentText().strip() or "UN",
            "dimensoes": self.dim_edit.text().strip(),
            "metros_unidade": self.meters_edit.text().strip(),
            "peso_unid": self.weight_edit.text().strip(),
            "qty": self.qty_edit.text().strip(),
            "alerta": self.alert_edit.text().strip(),
            "p_compra": self.buy_price_edit.text().strip(),
            "pvp1": self.pvp1_edit.text().strip(),
            "pvp2": self.pvp2_edit.text().strip(),
            "fabricante": self.maker_edit.text().strip(),
            "modelo": self.model_edit.text().strip(),
            "obs": self.obs_edit.text().strip(),
        }

    def _new_product(self) -> None:
        self.current_code = ""
        self._auto_catalog_values = {}
        if hasattr(self, "catalog_suggestion_label"):
            self.catalog_suggestion_label.hide()
            self.normalize_description_button.hide()
            self.teach_catalog_button.hide()
        self._suggested_normalized_description = ""
        self.current_product_label.setText("Novo produto")
        self.inspector_meta_label.setText("Preencha a identificação e os dados de stock.")
        self._set_inspector_state("Novo")
        self.inspector_available_label.setText("0,00 UN")
        self.code_edit.setText(self.backend.product_next_code())
        for widget in (
            self.desc_edit,
            self.dim_edit,
            self.meters_edit,
            self.weight_edit,
            self.qty_edit,
            self.alert_edit,
            self.buy_price_edit,
            self.pvp1_edit,
            self.pvp2_edit,
            self.maker_edit,
            self.model_edit,
            self.obs_edit,
        ):
            widget.clear()
        self.unit_combo.setCurrentText("UN")
        self.category_combo.setCurrentText("")
        self.subcat_combo.setCurrentText("")
        self.type_combo.setCurrentText("")
        self.moves_table.setRowCount(0)
        self._refresh_price_labels()
        self._show_form_page()

    def _fill_form(self, detail: dict) -> None:
        self._auto_catalog_values = {}
        self.catalog_suggestion_label.hide()
        self.normalize_description_button.hide()
        self.teach_catalog_button.hide()
        self._suggested_normalized_description = ""
        self.current_code = str(detail.get("codigo", "") or "").strip()
        description = str(detail.get("descricao", "") or "").strip() or "-"
        category = str(detail.get("categoria", "") or "").strip() or "Sem categoria"
        product_type = str(detail.get("tipo", "") or "").strip() or "Sem tipo"
        unit = str(detail.get("unid", "UN") or "UN").strip() or "UN"
        self.current_product_label.setText(self.current_code or "Produto")
        self.inspector_meta_label.setText(f"{description}\n{category} · {product_type}")
        self.code_edit.setText(self.current_code)
        self.desc_edit.setText(description if description != "-" else "")
        self.category_combo.setCurrentText(str(detail.get("categoria", "") or "").strip())
        self._sync_form_catalog()
        self.subcat_combo.setCurrentText(str(detail.get("subcat", "") or "").strip())
        self._sync_form_catalog()
        self.type_combo.setCurrentText(str(detail.get("tipo", "") or "").strip())
        self.unit_combo.setCurrentText(str(detail.get("unid", "UN") or "UN").strip() or "UN")
        self.dim_edit.setText(str(detail.get("dimensoes", "") or "").strip())
        self.meters_edit.setText(str(detail.get("metros_unidade", detail.get("metros", 0)) or 0))
        self.weight_edit.setText(str(detail.get("peso_unid", 0) or 0))
        self.qty_edit.setText(str(detail.get("qty", 0) or 0))
        self.alert_edit.setText(str(detail.get("alerta", 0) or 0))
        self.buy_price_edit.setText(str(detail.get("p_compra", 0) or 0))
        self.pvp1_edit.setText(str(detail.get("pvp1", 0) or 0))
        self.pvp2_edit.setText(str(detail.get("pvp2", 0) or 0))
        self.maker_edit.setText(str(detail.get("fabricante", "") or "").strip())
        self.model_edit.setText(str(detail.get("modelo", "") or "").strip())
        self.obs_edit.setText(str(detail.get("obs", "") or "").strip())
        self._refresh_price_labels()
        available = float(detail.get("available_qty", detail.get("qty", 0)) or 0)
        self.inspector_available_label.setText(f"{available:,.2f} {unit}".replace(",", "X").replace(".", ",").replace("X", "."))
        self._set_inspector_state(self._product_state_label(detail))
        self._refresh_moves_filters(detail)
        self._refresh_moves_view()

    def _refresh_price_labels(self) -> None:
        try:
            detail = self.backend._product_normalize_payload(self._payload())
            preco_unid = float(detail.get("preco_unid", 0) or 0)
            qty = float(detail.get("qty", 0) or 0)
            unit = str(detail.get("unid", "UN") or "UN").strip() or "UN"
            self.price_unit_label.setText(self._fmt_eur(preco_unid))
            self.stock_value_label.setText(self._fmt_eur(preco_unid * qty))
            self.inspector_qty_label.setText(f"{qty:,.2f} {unit}".replace(",", "X").replace(".", ",").replace("X", "."))
        except Exception:
            self.price_unit_label.setText(self._fmt_eur(0))
            self.stock_value_label.setText(self._fmt_eur(0))
            self.inspector_qty_label.setText("0,00 UN")

    def _apply_row_colors(self, row_index: int, severity: str, band: str) -> None:
        background = QColor("#ffffff" if band == "even" else "#f6f9fd")
        foreground = QColor("#0f172a")
        for col in range(self.table.columnCount()):
            item = self.table.item(row_index, col)
            if item is None:
                continue
            item.setBackground(QBrush(background))
            item.setForeground(QBrush(foreground))
        qty_item = self.table.item(row_index, 5)
        status_item = self.table.item(row_index, 10)
        if severity == "warning":
            for item in (qty_item, status_item):
                if item is not None:
                    item.setBackground(QBrush(QColor("#fff7ed")))
                    item.setForeground(QBrush(QColor("#9a3412")))

    def _product_state_label(self, row: dict) -> str:
        status = str(row.get("quality_status", "") or "").strip()
        pending = float(row.get("quality_pending_qty", 0) or 0)
        if pending > 0 and status:
            return status
        qty = float(row.get("qty", 0) or 0)
        alerta = float(row.get("alerta", 0) or 0)
        if qty <= 0:
            return "Sem stock"
        if alerta > 0 and qty <= alerta:
            return "Stock baixo"
        return "Disponível"

    def _fill_moves(self, rows: list[dict]) -> None:
        self.moves_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = [
                str(row.get("data", "") or ""),
                str(row.get("tipo", "") or ""),
                str(row.get("operador", "") or ""),
                f"{float(row.get('qtd', 0) or 0):.2f}",
                f"{float(row.get('antes', 0) or 0):.2f}",
                f"{float(row.get('depois', 0) or 0):.2f}",
                str(row.get("obs", "") or ""),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index < 6:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.moves_table.setItem(row_index, col_index, item)

    def _refresh_moves_filters(self, detail: dict) -> None:
        operators = ["Todos"] + self.backend.operator_names()
        years = ["Todos"] + self.backend.product_movement_years(str(detail.get("codigo", "") or "").strip())
        current_operator = self.moves_operator_combo.currentText() or "Todos"
        current_year = self.moves_year_combo.currentText() or "Todos"
        self.moves_operator_combo.blockSignals(True)
        self.moves_year_combo.blockSignals(True)
        self.moves_operator_combo.clear()
        self.moves_operator_combo.addItems(operators)
        self.moves_year_combo.clear()
        self.moves_year_combo.addItems(years)
        self.moves_operator_combo.setCurrentText(current_operator if current_operator in operators else "Todos")
        self.moves_year_combo.setCurrentText(current_year if current_year in years else "Todos")
        self.moves_operator_combo.blockSignals(False)
        self.moves_year_combo.blockSignals(False)

    def _refresh_moves_view(self) -> None:
        code = str(self.code_edit.text().strip() or self.current_code or "").strip()
        if not code:
            self.moves_table.setRowCount(0)
            self.moves_summary_label.setText("Sem produto selecionado.")
            return
        operator_name = self.moves_operator_combo.currentText().strip()
        year = self.moves_year_combo.currentText().strip()
        if operator_name == "Todos":
            operator_name = ""
        if year == "Todos":
            year = ""
        rows = self.backend.product_movements(code, limit=240, operator_name=operator_name, year=year)
        self._fill_moves(rows)
        summary = self.backend.product_issue_summary(operator_name=operator_name, year=year, codigo=code)
        if summary["linhas"] > 0:
            self.moves_summary_label.setText(
                f"Entregas a operador: {summary['linhas']} | Qtd {summary['qtd_total']:.2f} | Valor {self._fmt_eur(summary['valor_total'])}"
            )
        else:
            self.moves_summary_label.setText("Sem entregas a operador no período selecionado.")

    def _filtered_product_rows(self, query: str = "") -> list[dict]:
        rows = self.backend.product_rows("", in_stock_only=self.only_stock_check.isChecked())
        category_filter = self._filter_combo_text(self.filter_category_combo)
        subcat_filter = self._filter_combo_text(self.filter_subcat_combo)
        type_filter = self._filter_combo_text(self.filter_type_combo)
        state_filter = self._filter_combo_text(self.filter_state_combo).casefold()
        if category_filter:
            rows = [row for row in rows if str(row.get("categoria", "") or "").strip().casefold() == category_filter.casefold()]
        if subcat_filter:
            rows = [row for row in rows if str(row.get("subcat", "") or "").strip().casefold() == subcat_filter.casefold()]
        if type_filter:
            rows = [row for row in rows if str(row.get("tipo", "") or "").strip().casefold() == type_filter.casefold()]
        if state_filter:
            filtered: list[dict] = []
            for row in rows:
                state_label = _product_grid_search_normalize(self._product_state_label(row))
                if state_filter == "disponivel" and not state_label.startswith("dispon"):
                    continue
                if state_filter == "stock baixo" and "baixo" not in state_label:
                    continue
                if state_filter == "sem stock" and "sem stock" not in state_label:
                    continue
                if state_filter == "qualidade" and not any(token in state_label for token in ("qual", "inspec", "pend", "bloque", "rejeit")):
                    continue
                filtered.append(row)
            rows = filtered
        query = str(query or "").strip()
        if query:
            rows = [row for row in rows if _product_grid_search_matches(list(row.values()), query)]
        return rows

    def _open_assemblies_catalog(self) -> None:
        dialog = _AssembliesCatalogDialog(self.backend, self)
        dialog.exec()
        self.refresh()

    def open_full_grid(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Grelha de produtos")
        dialog.resize(1500, 820)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("Stock de produtos")
        title.setStyleSheet("font-size: 16px; font-weight: 900; color: #10253d;")
        info = QLabel(self.table_count_label.text())
        info.setProperty("role", "muted")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(info)
        layout.addLayout(header)

        search_box = QWidget()
        search_box.setObjectName("FullProductSearchBox")
        search_box.setStyleSheet(
            "QWidget#FullProductSearchBox { background: #ffffff; border: 1px solid #b8c9df; border-radius: 8px; }"
            "QLineEdit { border: none; background: transparent; padding: 7px 8px 7px 0; }"
        )
        search_layout = QHBoxLayout(search_box)
        search_layout.setContentsMargins(9, 0, 9, 0)
        search_layout.setSpacing(6)
        search_icon = QLabel("🔍")
        search_icon.setFixedWidth(20)
        search_icon.setAlignment(Qt.AlignCenter)
        search_icon.setStyleSheet("font-size: 15px; color: #33516f;")
        search_edit = QLineEdit()
        search_edit.setPlaceholderText("Pesquisar ou picar PRD|... por codigo, descricao, categoria, tipo, dimensao...")
        search_layout.addWidget(search_icon)
        search_layout.addWidget(search_edit, 1)
        layout.addWidget(search_box)

        table = QTableWidget(0, 22)
        table.setHorizontalHeaderLabels(
            [
                "Codigo",
                "Descricao",
                "Categoria",
                "Subcat.",
                "Tipo",
                "Dimensoes",
                "Unid.",
                "Qtd fisica",
                "Disponivel",
                "Pendente Q.",
                "Alerta",
                "Compra",
                "Preco/Unid.",
                "PVP1",
                "PVP2",
                "Valor stock",
                "Metros/Un.",
                "Peso/Un.",
                "Fabricante",
                "Modelo",
                "Atualizado",
                "Estado",
            ]
        )
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setWordWrap(False)
        table.setStyleSheet(self.table.styleSheet())
        header_view = table.horizontalHeader()
        header_view.setMinimumSectionSize(48)
        widths = [128, 300, 140, 140, 150, 150, 64, 84, 86, 96, 78, 94, 104, 94, 94, 112, 92, 92, 140, 140, 150, 118]
        for column, width in enumerate(widths):
            header_view.setSectionResizeMode(column, QHeaderView.Interactive)
            table.setColumnWidth(column, width)
        table.setSortingEnabled(True)

        render_timer = QTimer(dialog)
        render_timer.setSingleShot(True)

        def row_columns(row: dict) -> list[object]:
            state_label = self._product_state_label(row)
            return [
                row.get("codigo", "-"),
                row.get("descricao", "-"),
                row.get("category_display", row.get("categoria", "-")),
                row.get("subcat", "-"),
                row.get("tipo", "-"),
                row.get("dimensoes", "-"),
                row.get("unid", "UN"),
                f"{float(row.get('qty', 0) or 0):.2f}",
                f"{float(row.get('available_qty', 0) or 0):.2f}",
                f"{float(row.get('quality_pending_qty', 0) or 0):.2f}",
                f"{float(row.get('alerta', 0) or 0):.2f}",
                self._fmt_eur(row.get("p_compra", 0)),
                self._fmt_eur(row.get("preco_unid", 0)),
                self._fmt_eur(row.get("pvp1", 0)),
                self._fmt_eur(row.get("pvp2", 0)),
                self._fmt_eur(row.get("valor_stock", 0)),
                f"{float(row.get('metros_unidade', 0) or 0):.2f}",
                f"{float(row.get('peso_unid', 0) or 0):.2f}",
                row.get("fabricante", ""),
                row.get("modelo", ""),
                str(row.get("updated_at", "") or "").replace("T", " ")[:19],
                state_label,
            ]

        def render_full_grid() -> None:
            selected_code = ""
            selection = table.selectionModel()
            if selection is not None and selection.selectedRows():
                selected_item = table.item(selection.selectedRows()[0].row(), 0)
                selected_code = selected_item.text().strip() if selected_item is not None else ""
            selected_code = selected_code or self.current_code
            sort_col = header_view.sortIndicatorSection()
            sort_order = header_view.sortIndicatorOrder()
            table.setSortingEnabled(False)
            rows = self._filtered_product_rows(search_edit.text().strip())
            info.setText(f"{len(rows)} registos")
            table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                for col_index, value in enumerate(row_columns(row)):
                    item = QTableWidgetItem(str(value))
                    item.setToolTip(str(value))
                    if col_index not in (1, 2, 3, 4, 5, 18, 19, 20):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    table.setItem(row_index, col_index, item)
            table.setSortingEnabled(True)
            table.sortItems(sort_col, sort_order)
            if selected_code:
                for row_index in range(table.rowCount()):
                    item = table.item(row_index, 0)
                    if item is not None and item.text().strip() == selected_code:
                        table.selectRow(row_index)
                        break
            elif table.rowCount() > 0:
                table.selectRow(0)

        def select_scanned_product() -> None:
            value = search_edit.text().strip()
            try:
                result = self.backend.inventory_scan_lookup(value, expected_type="PRD")
            except ValueError as exc:
                if "|" in value or re.fullmatch(r"PRD[-_A-Z0-9]+", value, flags=re.IGNORECASE):
                    QMessageBox.warning(dialog, "Codigo de picagem", str(exc))
                return
            code = str(result.get("entity_id", "") or "").strip()
            self.current_code = code
            search_edit.setText(code)
            render_full_grid()
            for row_index in range(table.rowCount()):
                item = table.item(row_index, 0)
                if item is not None and item.text().strip().upper() == code.upper():
                    table.selectRow(row_index)
                    table.scrollToItem(item)
                    table.setFocus()
                    break

        render_timer.timeout.connect(render_full_grid)
        search_edit.textChanged.connect(lambda _text: render_timer.start(180))
        search_edit.returnPressed.connect(select_scanned_product)
        render_full_grid()
        search_edit.setFocus()
        table.itemDoubleClicked.connect(lambda *_args: dialog.accept())
        layout.addWidget(table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Selecionar")
        delete_btn = QPushButton("Apagar selecionados")
        delete_btn.setProperty("variant", "danger")
        delete_btn.setEnabled(False)
        buttons.addButton(delete_btn, QDialogButtonBox.ActionRole)

        def selected_product_codes() -> list[str]:
            selection = table.selectionModel()
            if selection is None:
                return []
            result: list[str] = []
            for index in selection.selectedRows():
                item = table.item(index.row(), 0)
                code = item.text().strip() if item is not None else ""
                if code and code not in result:
                    result.append(code)
            return result

        def update_delete_state() -> None:
            delete_btn.setEnabled(bool(selected_product_codes()))

        def delete_selected_products() -> None:
            codes = selected_product_codes()
            if not codes:
                QMessageBox.warning(dialog, "Apagar produtos", "Seleciona uma ou várias linhas.")
                return
            preview = ", ".join(codes[:6])
            if len(codes) > 6:
                preview += f" e mais {len(codes) - 6}"
            message = f"Apagar definitivamente {len(codes)} produto(s)?\n\n{preview}"
            if QMessageBox.question(dialog, "Apagar produtos", message) != QMessageBox.Yes:
                return
            try:
                removed = self.backend.product_remove_many(codes)
            except Exception as exc:
                QMessageBox.critical(dialog, "Apagar produtos", str(exc))
                return
            if self.current_code in codes:
                self._new_product()
            self.refresh()
            render_full_grid()
            QMessageBox.information(dialog, "Produtos", f"{removed} produto(s) apagado(s).")

        table.itemSelectionChanged.connect(update_delete_state)
        delete_btn.clicked.connect(delete_selected_products)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        selection = table.selectionModel()
        if selection is None or not selection.selectedRows():
            return
        item = table.item(selection.selectedRows()[0].row(), 0)
        if item is None:
            return
        self.current_code = item.text().strip()
        try:
            self._fill_form(self.backend.product_detail(self.current_code))
        except Exception:
            self.refresh()

    def refresh(self) -> None:
        self._load_presets()
        portfolio_rows = self.backend.product_rows("", in_stock_only=False)
        rows = self._filtered_product_rows(self.filter_edit.text().strip())
        products_in_stock = sum(1 for row in portfolio_rows if float(row.get("qty", 0) or 0) > 0)
        portfolio_value = sum(float(row.get("valor_stock", 0) or 0) for row in portfolio_rows)
        self.total_products_label.setText(str(len(portfolio_rows)))
        self.products_in_stock_label.setText(str(products_in_stock))
        self.portfolio_value_label.setText(self._fmt_eur(portfolio_value))
        self.table_count_label.setText(f"{len(rows)} de {len(portfolio_rows)} produtos")
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            state_label = self._product_state_label(row)
            values = [
                row.get("codigo", "-"),
                row.get("descricao", "-"),
                row.get("category_display", row.get("categoria", "-")),
                row.get("type_display", row.get("tipo", "-")),
                f"{float(row.get('qty', 0) or 0):.2f}",
                f"{float(row.get('available_qty', row.get('qty', 0)) or 0):.2f}",
                f"{float(row.get('alerta', 0) or 0):.2f}",
                self._fmt_eur(row.get("preco_unid", 0)),
                self._fmt_eur(row.get("valor_stock", 0)),
                str(row.get("updated_at", "") or "").replace("T", " ")[:19],
                state_label,
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if col_index >= 4:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, col_index, item)
            self._apply_row_colors(row_index, str(row.get("severity", "ok")), str(row.get("band", "even")))
        if self.table.rowCount() > 0:
            target = 0
            for row_index in range(self.table.rowCount()):
                if str(self.table.item(row_index, 0).text()) == self.current_code:
                    target = row_index
                    break
            self.table.selectRow(target)
            self._on_selection_changed()
        else:
            self._new_product()

    def _on_selection_changed(self) -> None:
        current = self.table.currentItem()
        if current is None:
            return
        code = str(self.table.item(current.row(), 0).text())
        if not code:
            return
        try:
            detail = self.backend.product_detail(code)
        except Exception:
            return
        self._fill_form(detail)

    def _save_product(self) -> None:
        payload = self._payload()
        analyzer = getattr(self.backend, "product_copilot_analysis", None)
        if callable(analyzer):
            analysis = dict(analyzer(payload.get("descricao", ""), payload.get("codigo", "")) or {})
            similar = list(analysis.get("semelhantes", []) or [])
            if analysis.get("duplicate_warning") and similar:
                first = dict(similar[0] or {})
                answer = QMessageBox.question(
                    self,
                    "Possível produto duplicado",
                    "O Copiloto encontrou um produto muito semelhante:\n\n"
                    f"{first.get('codigo', '')} · {first.get('descricao', '')}\n"
                    f"Semelhança: {int(first.get('percent', 0) or 0)}%\n\n"
                    "Queres guardar este registo mesmo assim?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return
        try:
            detail = self.backend.product_save(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
            return
        self.current_code = str(detail.get("codigo", "") or "").strip()
        self.refresh()

    def _remove_product(self) -> None:
        code = self.code_edit.text().strip() or self.current_code
        if not code:
            QMessageBox.warning(self, "Produtos", "Seleciona um produto.")
            return
        if QMessageBox.question(self, "Remover produto", f"Remover produto {code}?") != QMessageBox.Yes:
            return
        try:
            self.backend.product_remove(code)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
            return
        self._new_product()
        self.refresh()

    def _consume_product(self) -> None:
        code = self.code_edit.text().strip() or self.current_code
        if not code:
            QMessageBox.warning(self, "Produtos", "Seleciona um produto.")
            return
        dialog = _ConsumeDialog(code, self.backend.operator_names(), self)
        if dialog.exec() != QDialog.Accepted:
            return
        payload = dialog.values()
        try:
            detail = self.backend.product_consume(
                code,
                payload.get("qtd", 0),
                payload.get("obs", ""),
                target_operator=payload.get("operator", ""),
                issue_mode=payload.get("mode", "stock"),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
            return
        self.current_code = str(detail.get("codigo", "") or "").strip()
        self.refresh()
        self._show_moves_page()

    def _open_pdf(self) -> None:
        code = self.code_edit.text().strip() or self.current_code
        if not code:
            QMessageBox.warning(self, "Produtos", "Seleciona um produto.")
            return
        try:
            self.backend.product_open_sheet_pdf(code)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))

    def _select_stock_pdf_filters(self) -> dict[str, object] | None:
        options = dict(self.backend.product_stock_filters() or {})
        categories = list(options.get("categories", []) or [])
        types = list(options.get("types", []) or [])
        if not categories and not types:
            QMessageBox.warning(self, "Preview stock", "Não existem produtos para apresentar.")
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle("Conteúdo do relatório de produtos")
        dialog.setMinimumWidth(680)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        title = QLabel("Filtros incluídos no relatório")
        title.setStyleSheet("font-size: 15px; font-weight: 800; color: #10253d;")
        layout.addWidget(title)

        def add_filter_group(label: str, all_label: str, values: list[str]) -> tuple[QCheckBox, list[QCheckBox]]:
            heading = QLabel(label)
            heading.setStyleSheet("font-weight: 800; color: #10253d;")
            layout.addWidget(heading)
            all_check = QCheckBox(all_label)
            all_check.setChecked(True)
            all_check.setStyleSheet("font-weight: 700;")
            layout.addWidget(all_check)
            grid = QGridLayout()
            grid.setHorizontalSpacing(14)
            grid.setVerticalSpacing(6)
            checks: list[QCheckBox] = []
            for index, value in enumerate(values):
                check = QCheckBox(value)
                check.setChecked(True)
                checks.append(check)
                grid.addWidget(check, index // 4, index % 4)
            layout.addLayout(grid)

            def set_all(checked: bool) -> None:
                for check in checks:
                    check.setChecked(checked)

            def update_all() -> None:
                all_check.blockSignals(True)
                all_check.setChecked(all(check.isChecked() for check in checks))
                all_check.blockSignals(False)

            all_check.toggled.connect(set_all)
            for check in checks:
                check.toggled.connect(update_all)
            return all_check, checks

        category_all, category_checks = add_filter_group("Categorias", "Todas as categorias", categories)
        type_all, type_checks = add_filter_group("Tipos", "Todos os tipos", types)

        stock_only_check = QCheckBox("Apenas produtos com stock disponível")
        stock_only_check.setChecked(self.only_stock_check.isChecked())
        layout.addWidget(stock_only_check)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Pré-visualizar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return None

        selected_categories = [value for value, check in zip(categories, category_checks) if check.isChecked()]
        selected_types = [value for value, check in zip(types, type_checks) if check.isChecked()]
        if not selected_categories or not selected_types:
            QMessageBox.warning(self, "Preview stock", "Seleciona pelo menos uma categoria e um tipo.")
            return None
        self.only_stock_check.setChecked(stock_only_check.isChecked())
        return {
            "categories": None if category_all.isChecked() else selected_categories,
            "types": None if type_all.isChecked() else selected_types,
            "in_stock_only": stock_only_check.isChecked(),
        }

    def _open_stock_pdf(self) -> None:
        filters = self._select_stock_pdf_filters()
        if filters is None:
            return
        try:
            self.backend.product_open_stock_pdf(**filters)
        except Exception as exc:
            QMessageBox.critical(self, "Preview stock", str(exc))

    def _open_label_pdf(self) -> None:
        code = self.code_edit.text().strip() or self.current_code
        if not code:
            QMessageBox.warning(self, "Produtos", "Seleciona um produto.")
            return
        try:
            self.backend.product_open_label_pdf(code)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
