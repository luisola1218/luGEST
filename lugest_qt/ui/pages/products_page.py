from __future__ import annotations

import re
import unicodedata

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame


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


class ProductsPage(QWidget):
    page_title = "Produtos"
    page_subtitle = "Cadastro, stock, preco e movimentos do produto acabado no desktop Qt."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.current_code = ""
        self._moves_years: list[str] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(6)

        top_card = CardFrame()
        top_card.set_tone("info")
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 10, 14, 10)
        top_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        title_wrap = QVBoxLayout()
        title_wrap.setContentsMargins(0, 0, 0, 0)
        title_wrap.setSpacing(2)
        title = QLabel("Stock de Produtos")
        title.setStyleSheet("font-size: 17px; font-weight: 900; color: #10253d;")
        subtitle = QLabel("Cadastro, disponibilidade, preços e movimentos com leitura rápida para operação.")
        subtitle.setProperty("role", "muted")
        title_wrap.addWidget(title)
        title_wrap.addWidget(subtitle)
        top_bar.addLayout(title_wrap, 1)
        self._product_filter_timer = QTimer(self)
        self._product_filter_timer.setSingleShot(True)
        self._product_filter_timer.timeout.connect(self.refresh)
        search_box = QWidget()
        search_box.setObjectName("ProductSearchBox")
        search_box.setMaximumWidth(360)
        search_box.setStyleSheet(
            "QWidget#ProductSearchBox { background: #ffffff; border: 1px solid #b8c9df; border-radius: 8px; }"
            "QLineEdit { border: none; background: transparent; padding: 6px 8px 6px 0; }"
        )
        search_layout = QHBoxLayout(search_box)
        search_layout.setContentsMargins(8, 0, 8, 0)
        search_layout.setSpacing(6)
        search_icon = QLabel("🔍")
        search_icon.setFixedWidth(20)
        search_icon.setAlignment(Qt.AlignCenter)
        search_icon.setStyleSheet("font-size: 15px; color: #33516f;")
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar ou picar PRD|...")
        self.filter_edit.textChanged.connect(lambda _text: self._product_filter_timer.start(180))
        self.filter_edit.returnPressed.connect(self._select_scanned_product)
        search_layout.addWidget(search_icon)
        search_layout.addWidget(self.filter_edit, 1)
        top_bar.addWidget(search_box)
        self.only_stock_check = QCheckBox("Mostrar apenas com stock")
        self.only_stock_check.setChecked(False)
        self.only_stock_check.toggled.connect(self.refresh)
        top_bar.addWidget(self.only_stock_check)
        top_layout.addLayout(top_bar)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        filter_row.addWidget(QLabel("Categoria"))
        self.filter_category_combo = QComboBox()
        self.filter_category_combo.setEditable(False)
        self.filter_category_combo.setMinimumWidth(150)
        filter_row.addWidget(self.filter_category_combo)
        filter_row.addWidget(QLabel("Subcat."))
        self.filter_subcat_combo = QComboBox()
        self.filter_subcat_combo.setEditable(False)
        self.filter_subcat_combo.setMinimumWidth(150)
        filter_row.addWidget(self.filter_subcat_combo)
        filter_row.addWidget(QLabel("Tipo"))
        self.filter_type_combo = QComboBox()
        self.filter_type_combo.setEditable(False)
        self.filter_type_combo.setMinimumWidth(150)
        filter_row.addWidget(self.filter_type_combo)
        filter_row.addWidget(QLabel("Estado"))
        self.filter_state_combo = QComboBox()
        self.filter_state_combo.setEditable(False)
        self.filter_state_combo.addItems(["Todos", "Disponivel", "Stock baixo", "Sem stock", "Qualidade"])
        self.filter_state_combo.setMinimumWidth(132)
        filter_row.addWidget(self.filter_state_combo)
        clear_filters_btn = QPushButton("Limpar")
        clear_filters_btn.setProperty("compact", "true")
        clear_filters_btn.setProperty("variant", "secondary")
        clear_filters_btn.clicked.connect(self._clear_product_filters)
        filter_row.addWidget(clear_filters_btn)
        filter_row.addStretch(1)
        top_layout.addLayout(filter_row)

        summary = QHBoxLayout()
        summary.setSpacing(14)
        self.current_product_label = QLabel("Sem produto selecionado")
        self.current_product_label.setProperty("role", "field_value")
        self.price_unit_label = QLabel("0,00 EUR")
        self.price_unit_label.setProperty("role", "field_value")
        self.stock_value_label = QLabel("0,00 EUR")
        self.stock_value_label.setProperty("role", "field_value")
        summary.addWidget(self.current_product_label, 1)
        summary.addWidget(QLabel("Preco/Unid."))
        summary.addWidget(self.price_unit_label)
        summary.addWidget(QLabel("Valor stock"))
        summary.addWidget(self.stock_value_label)
        top_layout.addLayout(summary)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.new_btn = QPushButton("Novo")
        self.new_btn.clicked.connect(self._new_product)
        self.save_btn = QPushButton("Guardar")
        self.save_btn.clicked.connect(self._save_product)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self._remove_product)
        self.consume_btn = QPushButton("Baixa")
        self.consume_btn.setProperty("variant", "secondary")
        self.consume_btn.clicked.connect(self._consume_product)
        self.pdf_btn = QPushButton("Ficha PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_pdf)
        self.stock_pdf_btn = QPushButton("Preview stock")
        self.stock_pdf_btn.setProperty("variant", "secondary")
        self.stock_pdf_btn.clicked.connect(self._open_stock_pdf)
        self.label_btn = QPushButton("Etiqueta")
        self.label_btn.setProperty("variant", "secondary")
        self.label_btn.clicked.connect(self._open_label_pdf)
        self.form_mode_btn = QPushButton("Ficha")
        self.form_mode_btn.setProperty("variant", "secondary")
        self.form_mode_btn.clicked.connect(self._show_form_page)
        self.moves_mode_btn = QPushButton("Movs.")
        self.moves_mode_btn.setProperty("variant", "secondary")
        self.moves_mode_btn.clicked.connect(self._show_moves_page)
        self.full_grid_btn = QPushButton("Grelha")
        self.full_grid_btn.setProperty("variant", "secondary")
        self.full_grid_btn.clicked.connect(self.open_full_grid)
        for button, width in (
            (self.new_btn, 74),
            (self.save_btn, 86),
            (self.remove_btn, 88),
            (self.consume_btn, 78),
            (self.pdf_btn, 96),
            (self.stock_pdf_btn, 112),
            (self.label_btn, 82),
            (self.form_mode_btn, 72),
            (self.moves_mode_btn, 72),
            (self.full_grid_btn, 74),
        ):
            button.setStyleSheet("font-weight: 500;")
            button.setMaximumWidth(width)
            actions.addWidget(button)
        actions.addStretch(1)
        top_layout.addLayout(actions)
        root.addWidget(top_card)

        table_card = CardFrame()
        table_card.set_tone("default")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_header = QHBoxLayout()
        table_title = QLabel("Produtos em stock")
        table_title.setStyleSheet("font-size: 15px; font-weight: 900; color: #10253d;")
        self.table_count_label = QLabel("-")
        self.table_count_label.setProperty("role", "muted")
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(self.table_count_label)
        open_grid_btn = QPushButton("Janela inteira")
        open_grid_btn.setProperty("compact", "true")
        open_grid_btn.setProperty("variant", "secondary")
        open_grid_btn.clicked.connect(self.open_full_grid)
        table_header.addWidget(open_grid_btn)
        self.table = QTableWidget(0, 11)
        self.table.setObjectName("StockTable")
        self.table.setStyleSheet(
            "QTableWidget {"
            " selection-background-color: #fff3bf;"
            " selection-color: #0f172a;"
            "}"
            "QTableWidget::item:selected {"
            " background: #fff3bf;"
            " color: #0f172a;"
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
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.resizeSection(0, 156)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col, width in ((2, 112), (3, 104), (4, 76), (5, 86), (6, 76), (7, 112), (8, 112), (9, 142), (10, 118)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        table_layout.addLayout(table_header)
        table_layout.addWidget(self.table)
        root.addWidget(table_card, 5)

        detail_host = CardFrame()
        detail_host.set_tone("default")
        detail_host_layout = QVBoxLayout(detail_host)
        detail_host_layout.setContentsMargins(14, 12, 14, 12)
        detail_host_layout.setSpacing(8)
        self.detail_mode_label = QLabel("Ficha do produto")
        self.detail_mode_label.setStyleSheet("font-size: 16px; color: #0f172a;")
        detail_host_layout.addWidget(self.detail_mode_label)
        self.detail_stack = QStackedWidget()
        detail_host_layout.addWidget(self.detail_stack)
        detail_host.setMaximumHeight(282)
        root.addWidget(detail_host, 2)

        self.form_page = QWidget()
        form_page_layout = QVBoxLayout(self.form_page)
        form_page_layout.setContentsMargins(0, 0, 0, 0)
        form_page_layout.setSpacing(10)
        form_grid = QGridLayout()
        form_grid.setHorizontalSpacing(10)
        form_grid.setVerticalSpacing(8)
        self.code_edit = QLineEdit()
        self.desc_edit = QLineEdit()
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
        fields = [
            ("Codigo", self.code_edit),
            ("Descricao", self.desc_edit),
            ("Categoria", self.category_combo),
            ("Subcat.", self.subcat_combo),
            ("Tipo", self.type_combo),
            ("Unid.", self.unit_combo),
            ("Dimensoes", self.dim_edit),
            ("Metros/Unid.", self.meters_edit),
            ("Peso/Unid.", self.weight_edit),
            ("Quantidade", self.qty_edit),
            ("Alerta", self.alert_edit),
            ("Compra (EUR)", self.buy_price_edit),
            ("PVP1", self.pvp1_edit),
            ("PVP2", self.pvp2_edit),
            ("Fabricante", self.maker_edit),
            ("Modelo", self.model_edit),
            ("Observacoes", self.obs_edit),
        ]
        for index, (label_text, widget) in enumerate(fields):
            row = index // 3
            col = (index % 3) * 2
            label = QLabel(label_text)
            label.setProperty("role", "muted")
            form_grid.addWidget(label, row, col)
            form_grid.addWidget(widget, row, col + 1)
        form_page_layout.addLayout(form_grid)
        form_page_layout.addStretch(1)

        self.moves_page = QWidget()
        moves_layout = QVBoxLayout(self.moves_page)
        moves_layout.setContentsMargins(0, 0, 0, 0)
        moves_layout.setSpacing(10)
        moves_filters = QHBoxLayout()
        moves_filters.setSpacing(8)
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
        moves_filters.addStretch(1)
        moves_filters.addWidget(self.moves_summary_label)
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
        moves_layout.addWidget(self.moves_table)

        self.detail_stack.addWidget(self.form_page)
        self.detail_stack.addWidget(self.moves_page)

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
        combo.addItem("Todas")
        combo.addItems([value for value in values if str(value or "").strip()])
        combo.setCurrentText(current if current in [combo.itemText(index) for index in range(combo.count())] else "Todas")
        combo.blockSignals(False)

    def _clear_product_filters(self) -> None:
        self.filter_edit.clear()
        for combo, value in (
            (self.filter_category_combo, "Todas"),
            (self.filter_subcat_combo, "Todas"),
            (self.filter_type_combo, "Todas"),
            (self.filter_state_combo, "Todos"),
        ):
            combo.blockSignals(True)
            combo.setCurrentText(value)
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
        for combo, value in (
            (self.filter_category_combo, "Todas"),
            (self.filter_subcat_combo, "Todas"),
            (self.filter_type_combo, "Todas"),
            (self.filter_state_combo, "Todos"),
        ):
            combo.blockSignals(True)
            combo.setCurrentText(value)
            combo.blockSignals(False)
        self.only_stock_check.blockSignals(True)
        self.only_stock_check.setChecked(False)
        self.only_stock_check.blockSignals(False)
        self.refresh()
        self.table.setFocus()

    def _filter_combo_text(self, combo: QComboBox) -> str:
        text = str(combo.currentText() or "").strip()
        return "" if text.casefold() in {"", "todos", "todas", "all"} else text

    def _set_mode_buttons(self, moves: bool) -> None:
        self.form_mode_btn.setEnabled(moves)
        self.moves_mode_btn.setEnabled(not moves)
        self.detail_mode_label.setText("Movimentos" if moves else "Ficha do produto")

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
            "" if self.filter_category_combo.currentText() == "Todas" else self.filter_category_combo.currentText(),
            "" if self.filter_subcat_combo.currentText() == "Todas" else self.filter_subcat_combo.currentText(),
        )
        self._set_filter_combo_values(self.filter_category_combo, self.backend.product_catalog_options().get("categorias", []), self.filter_category_combo.currentText() or "Todas")
        self._set_filter_combo_values(self.filter_subcat_combo, filter_presets.get("subcats", []), self.filter_subcat_combo.currentText() or "Todas")
        self._set_filter_combo_values(self.filter_type_combo, filter_presets.get("tipos", []), self.filter_type_combo.currentText() or "Todas")

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
        current_category = "" if self.filter_category_combo.currentText() == "Todas" else self.filter_category_combo.currentText().strip()
        current_subcat = "" if self.filter_subcat_combo.currentText() == "Todas" else self.filter_subcat_combo.currentText().strip()
        current_type = self.filter_type_combo.currentText() or "Todas"
        presets = self.backend.product_catalog_options(current_category, current_subcat)
        self._set_filter_combo_values(self.filter_subcat_combo, presets.get("subcats", []), self.filter_subcat_combo.currentText() or "Todas")
        selected_subcat = "" if self.filter_subcat_combo.currentText() == "Todas" else self.filter_subcat_combo.currentText().strip()
        type_presets = self.backend.product_catalog_options(current_category, selected_subcat)
        self._set_filter_combo_values(self.filter_type_combo, type_presets.get("tipos", []), current_type)
        self.refresh()

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
        self.current_product_label.setText("Novo produto")
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
        self.current_code = str(detail.get("codigo", "") or "").strip()
        self.current_product_label.setText(f"{self.current_code} | {detail.get('descricao', '-') or '-'}")
        self.code_edit.setText(self.current_code)
        self.desc_edit.setText(str(detail.get("descricao", "") or "").strip())
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
        self._refresh_moves_filters(detail)
        self._refresh_moves_view()

    def _refresh_price_labels(self) -> None:
        try:
            detail = self.backend._product_normalize_payload(self._payload())
            preco_unid = float(detail.get("preco_unid", 0) or 0)
            qty = float(detail.get("qty", 0) or 0)
            self.price_unit_label.setText(self._fmt_eur(preco_unid))
            self.stock_value_label.setText(self._fmt_eur(preco_unid * qty))
        except Exception:
            self.price_unit_label.setText(self._fmt_eur(0))
            self.stock_value_label.setText(self._fmt_eur(0))

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
        rows = self._filtered_product_rows(self.filter_edit.text().strip())
        self.table_count_label.setText(f"{len(rows)} registos")
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
        try:
            detail = self.backend.product_save(self._payload())
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
