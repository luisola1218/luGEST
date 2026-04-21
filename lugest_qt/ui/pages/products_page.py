from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
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
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(14)

        top_card = CardFrame()
        top_card.set_tone("info")
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 10, 14, 10)
        top_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        title_wrap = QVBoxLayout()
        title_wrap.setContentsMargins(0, 0, 0, 0)
        title_wrap.setSpacing(2)
        title = QLabel("Gestao de Produtos")
        title.setStyleSheet("font-size: 17px; color: #0f172a;")
        subtitle = QLabel("Fluxo antigo: lista principal, ficha compacta e movimentos em sub-menu.")
        subtitle.setProperty("role", "muted")
        title_wrap.addWidget(title)
        title_wrap.addWidget(subtitle)
        top_bar.addLayout(title_wrap, 1)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar codigo, descricao, categoria, tipo...")
        self.filter_edit.textChanged.connect(self.refresh)
        self.filter_edit.setMaximumWidth(420)
        top_bar.addWidget(self.filter_edit)
        top_layout.addLayout(top_bar)

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
        self.pdf_btn = QPushButton("Pre-visualizar PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_pdf)
        self.label_btn = QPushButton("Etiqueta")
        self.label_btn.setProperty("variant", "secondary")
        self.label_btn.clicked.connect(self._open_label_pdf)
        self.form_mode_btn = QPushButton("Ficha produto")
        self.form_mode_btn.setProperty("variant", "secondary")
        self.form_mode_btn.clicked.connect(self._show_form_page)
        self.moves_mode_btn = QPushButton("Movimentos")
        self.moves_mode_btn.setProperty("variant", "secondary")
        self.moves_mode_btn.clicked.connect(self._show_moves_page)
        for button, width in (
            (self.new_btn, 90),
            (self.save_btn, 96),
            (self.remove_btn, 96),
            (self.consume_btn, 88),
            (self.pdf_btn, 146),
            (self.label_btn, 92),
            (self.form_mode_btn, 118),
            (self.moves_mode_btn, 108),
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
        table_title = QLabel("Produtos")
        table_title.setStyleSheet("font-size: 17px; color: #0f172a;")
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            ["Codigo", "Descricao", "Categoria", "Tipo", "Qtd", "Alerta", "Preco/Unid.", "Valor Stock", "Atualizado"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.resizeSection(0, 156)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col, width in ((2, 112), (3, 104), (4, 76), (5, 76), (6, 112), (7, 112), (8, 148)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        table_layout.addWidget(table_title)
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
        detail_host.setMaximumHeight(308)
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
        form_note = QLabel("A tabela principal fica com mais area e os movimentos passam para o sub-menu proprio.")
        form_note.setProperty("role", "muted")
        form_page_layout.addWidget(form_note)
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
        presets = self.backend.product_presets()
        self._set_combo_values(self.category_combo, presets.get("categorias", []), self.category_combo.currentText())
        self._set_combo_values(self.subcat_combo, presets.get("subcats", []), self.subcat_combo.currentText())
        self._set_combo_values(self.type_combo, presets.get("tipos", []), self.type_combo.currentText())
        self._set_combo_values(self.unit_combo, presets.get("unidades", []), self.unit_combo.currentText() or "UN")

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
        self.moves_table.setRowCount(0)
        self._refresh_price_labels()
        self._show_form_page()

    def _fill_form(self, detail: dict) -> None:
        self.current_code = str(detail.get("codigo", "") or "").strip()
        self.current_product_label.setText(f"{self.current_code} | {detail.get('descricao', '-') or '-'}")
        self.code_edit.setText(self.current_code)
        self.desc_edit.setText(str(detail.get("descricao", "") or "").strip())
        self.category_combo.setCurrentText(str(detail.get("categoria", "") or "").strip())
        self.subcat_combo.setCurrentText(str(detail.get("subcat", "") or "").strip())
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
        background = QColor("#eef2f8" if band == "even" else "#e6ecf5")
        foreground = QColor("#0f172a")
        if severity == "warning":
            foreground = QColor("#b45309")
        for col in range(self.table.columnCount()):
            item = self.table.item(row_index, col)
            if item is None:
                continue
            item.setBackground(QBrush(background))
            item.setForeground(QBrush(foreground))

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

    def refresh(self) -> None:
        self._load_presets()
        rows = self.backend.product_rows(self.filter_edit.text().strip())
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = [
                row.get("codigo", "-"),
                row.get("descricao", "-"),
                row.get("categoria", "-"),
                row.get("tipo", "-"),
                f"{float(row.get('qty', 0) or 0):.2f}",
                f"{float(row.get('alerta', 0) or 0):.2f}",
                self._fmt_eur(row.get("preco_unid", 0)),
                self._fmt_eur(row.get("valor_stock", 0)),
                str(row.get("updated_at", "") or "").replace("T", " ")[:19],
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
        try:
            self.backend.product_open_stock_pdf()
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))

    def _open_label_pdf(self) -> None:
        code = self.code_edit.text().strip() or self.current_code
        if not code:
            QMessageBox.warning(self, "Produtos", "Seleciona um produto.")
            return
        try:
            self.backend.product_open_label_pdf(code)
        except Exception as exc:
            QMessageBox.critical(self, "Produtos", str(exc))
