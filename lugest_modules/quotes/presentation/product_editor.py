from __future__ import annotations
import unicodedata
from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from lugest_modules.quotes.domain.lines import product_line as build_product_line
def edit_product(owner: QWidget, backend: QuoteEditorPorts, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
    wrapped_initial = dict(initial or {})
    line_initial = dict(wrapped_initial.get("line") or {})
    if line_initial:
        for key, value in wrapped_initial.items():
            if key == "line":
                continue
            if key not in line_initial and value not in (None, ""):
                line_initial[key] = value
        initial = line_initial
    else:
        initial = wrapped_initial
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Produto de stock")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    search_box = QWidget()
    search_layout = QHBoxLayout(search_box)
    search_layout.setContentsMargins(0, 0, 0, 0)
    search_layout.setSpacing(6)
    search_icon = QLabel("🔍")
    search_icon.setFixedWidth(24)
    search_icon.setAlignment(Qt.AlignCenter)
    search_icon.setStyleSheet("font-size: 16px; color: #10253d;")
    search_edit = QLineEdit()
    search_edit.setPlaceholderText("Pesquisar por codigo, descricao, medida, tipo...")
    search_layout.addWidget(search_icon)
    search_layout.addWidget(search_edit, 1)
    combo = QComboBox()
    combo.setEditable(True)
    combo.setInsertPolicy(QComboBox.NoInsert)
    combo.setMaxVisibleItems(18)
    combo.addItem("")
    product_rows_all = [dict(row) for row in list(backend.ne_product_options("") or []) if isinstance(row, dict)]
    product_rows = list(product_rows_all)
    by_code = {}

    def _product_search_text(value: str) -> str:
        text = unicodedata.normalize("NFKD", str(value or ""))
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        return text.casefold().strip()

    def _product_row_haystack(row: dict) -> str:
        return _product_search_text(
            " ".join(
                str(row.get(key, "") or "")
                for key in (
                    "codigo",
                    "descricao",
                    "categoria",
                    "subcat",
                    "tipo",
                    "marca",
                    "modelo",
                    "dimensoes",
                    "obs",
                )
            )
        )

    def _filter_product_rows(query: str) -> list[dict]:
        needle = _product_search_text(query)
        if not needle:
            return list(product_rows_all[:120])
        terms = [term for term in needle.split() if term]
        matches = []
        for row in product_rows_all:
            haystack = _product_row_haystack(row)
            if all(term in haystack for term in terms):
                matches.append(row)
            if len(matches) >= 120:
                break
        return matches

    def _reload_product_options(select_code: str = "", search_text: str = "") -> None:
        nonlocal product_rows, by_code
        product_rows = [dict(row) for row in _filter_product_rows(search_text)]
        selected_code = str(select_code or "").strip()
        if selected_code and not any(str(row.get("codigo", "") or "").strip() == selected_code for row in product_rows):
            selected_row = next(
                (dict(row) for row in product_rows_all if str(row.get("codigo", "") or "").strip() == selected_code),
                None,
            )
            if selected_row:
                product_rows.insert(0, selected_row)
        by_code = {}
        current_text = combo.currentText().strip()
        keep_text = "" if str(search_text or "").strip() else str(current_text or "").strip()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("")
        for row in product_rows:
            code = str(row.get("codigo", "") or "").strip()
            if not code:
                continue
            by_code[code] = row
            combo.addItem(f"{code} - {str(row.get('descricao', '') or '').strip()}".strip(" -"), code)
        if select_code:
            for index in range(combo.count()):
                if str(combo.itemData(index) or "").strip() == select_code:
                    combo.setCurrentIndex(index)
                    break
        elif keep_text:
            combo.setEditText(keep_text)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    product_search_timer = QTimer(dialog)
    product_search_timer.setSingleShot(True)

    def _apply_product_search() -> None:
        _reload_product_options(search_text=search_edit.text())
        search_edit.setFocus()
        _refresh()

    def _on_product_search(_text: str) -> None:
        product_search_timer.start(180)

    def _new_product_dialog() -> dict | None:
        product_dialog = QDialog(dialog)
        product_dialog.setWindowTitle("Novo produto")
        product_layout = QVBoxLayout(product_dialog)
        product_form = QFormLayout()
        presets = dict(backend.product_catalog_options() or {})
        code_edit = QLineEdit(str(backend.product_next_code() or "").strip())
        desc_edit = QLineEdit()
        category_combo = QComboBox()
        category_combo.setEditable(True)
        category_combo.addItems([str(value or "").strip() for value in list(presets.get("categorias", []) or []) if str(value or "").strip()])
        subcat_combo = QComboBox()
        subcat_combo.setEditable(True)
        subcat_combo.addItems([str(value or "").strip() for value in list(presets.get("subcats", []) or []) if str(value or "").strip()])
        type_combo = QComboBox()
        type_combo.setEditable(True)
        type_combo.addItems([str(value or "").strip() for value in list(presets.get("tipos", []) or []) if str(value or "").strip()])
        unit_combo = QComboBox()
        unit_combo.setEditable(True)
        unit_combo.addItems([str(value or "").strip() for value in list(presets.get("unidades", []) or []) if str(value or "").strip()])
        unit_combo.setCurrentText("UN")
        dim_edit = QLineEdit()
        meters_spin = QDoubleSpinBox()
        meters_spin.setRange(0.0, 1000000.0)
        meters_spin.setDecimals(4)
        weight_spin = QDoubleSpinBox()
        weight_spin.setRange(0.0, 1000000.0)
        weight_spin.setDecimals(4)
        qty_spin_new = QDoubleSpinBox()
        qty_spin_new.setRange(0.0, 1000000.0)
        qty_spin_new.setDecimals(2)
        qty_spin_new.setValue(0.0)
        alert_spin = QDoubleSpinBox()
        alert_spin.setRange(0.0, 1000000.0)
        alert_spin.setDecimals(2)
        buy_spin = QDoubleSpinBox()
        buy_spin.setRange(0.0, 1000000.0)
        buy_spin.setDecimals(4)
        buy_total_spin = QDoubleSpinBox()
        buy_total_spin.setRange(0.0, 1000000000.0)
        buy_total_spin.setDecimals(4)
        pvp1_spin = QDoubleSpinBox()
        pvp1_spin.setRange(0.0, 1000000.0)
        pvp1_spin.setDecimals(4)
        pvp2_spin = QDoubleSpinBox()
        pvp2_spin.setRange(0.0, 1000000.0)
        pvp2_spin.setDecimals(4)
        obs_edit = QLineEdit()

        def _set_combo_items(combo: QComboBox, values: list[str], current_text: str = "") -> None:
            combo.blockSignals(True)
            combo.clear()
            combo.addItems([str(value or "").strip() for value in list(values or []) if str(value or "").strip()])
            combo.setCurrentText(current_text)
            combo.blockSignals(False)

        def _sync_catalog_combos() -> None:
            current_category = category_combo.currentText().strip()
            current_subcat = subcat_combo.currentText().strip()
            current_type = type_combo.currentText().strip()
            subcat_presets = dict(backend.product_catalog_options(current_category, current_subcat) or {})
            _set_combo_items(subcat_combo, subcat_presets.get("subcats", []), current_subcat)
            current_subcat = subcat_combo.currentText().strip()
            type_presets = dict(backend.product_catalog_options(current_category, current_subcat) or {})
            _set_combo_items(type_combo, type_presets.get("tipos", []), current_type)

        product_form.addRow("Codigo", code_edit)
        product_form.addRow("Descricao", desc_edit)
        product_form.addRow("Categoria", category_combo)
        product_form.addRow("Subcat.", subcat_combo)
        product_form.addRow("Tipo", type_combo)
        product_form.addRow("Unid.", unit_combo)
        product_form.addRow("Dimensoes", dim_edit)
        product_form.addRow("Metros/Unid.", meters_spin)
        product_form.addRow("Peso/Unid.", weight_spin)
        product_form.addRow("Quantidade", qty_spin_new)
        product_form.addRow("Alerta", alert_spin)
        product_form.addRow("Compra/Unid. (EUR)", buy_spin)
        product_form.addRow("Compra total stock", buy_total_spin)
        product_form.addRow("PVP1", pvp1_spin)
        product_form.addRow("PVP2", pvp2_spin)
        product_form.addRow("Observacoes", obs_edit)
        product_layout.addLayout(product_form)
        price_note = QLabel("")
        price_note.setWordWrap(True)
        price_note.setProperty("role", "muted")
        product_layout.addWidget(price_note)
        note = QLabel("Se o produto ainda nao existir em stock, podes cria-lo aqui com stock inicial zero e reutiliza-lo logo no conjunto.")
        note.setWordWrap(True)
        note.setProperty("role", "muted")
        product_layout.addWidget(note)
        product_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        product_buttons.accepted.connect(product_dialog.accept)
        product_buttons.rejected.connect(product_dialog.reject)
        product_layout.addWidget(product_buttons)
        price_sync = {"busy": False, "last_source": "unit"}

        def _refresh_product_price_note() -> None:
            qty_value = float(qty_spin_new.value() or 0.0)
            unit_txt = unit_combo.currentText().strip() or "UN"
            unit_price = float(buy_spin.value() or 0.0)
            total_price = float(buy_total_spin.value() or 0.0)
            if qty_value > 0:
                price_note.setText(
                    f"Compra: {_fmt_eur(unit_price)}/{unit_txt} | total stock {_fmt_eur(total_price)} "
                    f"| {qty_value:.2f} {unit_txt}."
                )
            else:
                price_note.setText(f"Compra: {_fmt_eur(unit_price)}/{unit_txt}. Define quantidade para calcular o total.")

        def _sync_buy_total_from_unit() -> None:
            if price_sync["busy"]:
                return
            price_sync["busy"] = True
            try:
                price_sync["last_source"] = "unit"
                buy_total_spin.setValue(round(float(qty_spin_new.value() or 0.0) * float(buy_spin.value() or 0.0), 4))
                _refresh_product_price_note()
            finally:
                price_sync["busy"] = False

        def _sync_buy_unit_from_total() -> None:
            if price_sync["busy"]:
                return
            price_sync["busy"] = True
            try:
                price_sync["last_source"] = "total"
                qty_value = float(qty_spin_new.value() or 0.0)
                if qty_value > 0:
                    buy_spin.setValue(round(float(buy_total_spin.value() or 0.0) / qty_value, 4))
                _refresh_product_price_note()
            finally:
                price_sync["busy"] = False

        def _sync_buy_after_qty_change() -> None:
            if price_sync.get("last_source") == "total":
                _sync_buy_unit_from_total()
            else:
                _sync_buy_total_from_unit()

        category_combo.currentTextChanged.connect(lambda _text: _sync_catalog_combos())
        subcat_combo.currentTextChanged.connect(lambda _text: _sync_catalog_combos())
        unit_combo.currentTextChanged.connect(lambda _text: _refresh_product_price_note())
        qty_spin_new.valueChanged.connect(lambda _value: _sync_buy_after_qty_change())
        buy_spin.valueChanged.connect(lambda _value: _sync_buy_total_from_unit())
        buy_total_spin.valueChanged.connect(lambda _value: _sync_buy_unit_from_total())
        _sync_catalog_combos()
        _refresh_product_price_note()
        if product_dialog.exec() != QDialog.Accepted:
            return None
        try:
            return dict(
                backend.product_save(
                    {
                        "codigo": code_edit.text().strip(),
                        "descricao": desc_edit.text().strip(),
                        "categoria": category_combo.currentText().strip(),
                        "subcat": subcat_combo.currentText().strip(),
                        "tipo": type_combo.currentText().strip() or "Montagem",
                        "unid": unit_combo.currentText().strip() or "UN",
                        "dimensoes": dim_edit.text().strip(),
                        "metros_unidade": float(meters_spin.value() or 0.0),
                        "peso_unid": float(weight_spin.value() or 0.0),
                        "qty": float(qty_spin_new.value() or 0.0),
                        "alerta": float(alert_spin.value() or 0.0),
                        "p_compra": float(buy_spin.value() or 0.0),
                        "pvp1": float(pvp1_spin.value() or 0.0),
                        "pvp2": float(pvp2_spin.value() or 0.0),
                        "obs": obs_edit.text().strip(),
                    }
                )
                or {}
            )
        except Exception as exc:
            QMessageBox.critical(product_dialog, "Produtos", str(exc))
            return None

    wanted_code = str(initial.get("produto_codigo", "") or initial.get("ref_externa", "") or "").strip()
    wanted_desc = str(initial.get("descricao", "") or "").strip()
    _reload_product_options(wanted_code)
    qty_spin = QDoubleSpinBox()
    qty_spin.setRange(0.01, 1000000.0)
    qty_spin.setDecimals(2)
    qty_spin.setValue(float(initial.get("quantity_units", initial.get("qtd", 1)) or 1))
    manual_unit_combo = QComboBox()
    manual_unit_combo.setEditable(True)
    manual_unit_combo.addItems(["UN", "MT", "KG", "L", "CJ"])
    manual_unit_combo.setCurrentText(str(initial.get("produto_unid", "") or "UN").strip() or "UN")
    manual_price_spin = QDoubleSpinBox()
    manual_price_spin.setRange(0.0, 1000000.0)
    manual_price_spin.setDecimals(4)
    manual_price_spin.setValue(float(initial.get("preco_unit", 0) or 0.0))
    manual_state = {"unit_dirty": False, "price_dirty": False}
    form.addRow("Pesquisar", search_box)
    form.addRow("Produto", combo)
    form.addRow("Quantidade", qty_spin)
    form.addRow("Unid. manual", manual_unit_combo)
    form.addRow("Preco manual", manual_price_spin)
    layout.addLayout(form)
    info_label = QLabel("")
    info_label.setWordWrap(True)
    layout.addWidget(info_label)
    actions = QHBoxLayout()
    new_product_btn = QPushButton("Novo artigo")
    new_product_btn.setProperty("variant", "secondary")
    actions.addWidget(new_product_btn)
    actions.addStretch(1)
    layout.addLayout(actions)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    def _current_product() -> dict | None:
        code = str(combo.currentData() or "").strip()
        if not code:
            text = combo.currentText().strip()
            code = text.split(" - ", 1)[0].strip()
        return by_code.get(code)

    if not _current_product():
        fallback_text = ""
        if wanted_code and wanted_desc:
            fallback_text = f"{wanted_code} - {wanted_desc}"
        elif wanted_code:
            fallback_text = wanted_code
        elif wanted_desc:
            fallback_text = wanted_desc
        if fallback_text:
            combo.blockSignals(True)
            combo.setCurrentText(fallback_text)
            combo.blockSignals(False)

    def _refresh() -> None:
        row = _current_product() or {}
        if row:
            sale_price = float(row.get("preco_venda", row.get("pvp1", row.get("preco_unid", row.get("preco", 0)))) or 0)
            total = float(qty_spin.value() or 0) * sale_price
            info_label.setText(
                f"Unid.: {str(row.get('unid', '-') or '-')} | PVP1: {_fmt_eur(sale_price)} | "
                f"Total: {_fmt_eur(total)}"
            )
        else:
            desc = combo.currentText().strip()
            total = float(qty_spin.value() or 0) * float(manual_price_spin.value() or 0)
            if desc:
                info_label.setText(f"Produto novo para compra/cotacao | Total: {_fmt_eur(total)}")
            else:
                info_label.setText("Seleciona um produto existente, cria um novo artigo ou escreve uma descricao para cotacao.")

    combo.currentTextChanged.connect(lambda _t: _refresh())
    qty_spin.valueChanged.connect(lambda _v: _refresh())
    manual_unit_combo.currentTextChanged.connect(lambda _t: manual_state.__setitem__("unit_dirty", True))
    manual_price_spin.valueChanged.connect(lambda _v: (manual_state.__setitem__("price_dirty", True), _refresh()))
    def _create_and_select_product() -> None:
        detail = _new_product_dialog()
        if not isinstance(detail, dict) or not str(detail.get("codigo", "") or "").strip():
            return
        product_rows_all[:] = [dict(row) for row in list(backend.ne_product_options("") or []) if isinstance(row, dict)]
        _reload_product_options(str(detail.get("codigo", "") or "").strip())
        _refresh()
    new_product_btn.clicked.connect(_create_and_select_product)
    product_search_timer.timeout.connect(_apply_product_search)
    search_edit.textChanged.connect(_on_product_search)
    _refresh()
    if dialog.exec() != QDialog.Accepted:
        return None
    product = _current_product()
    if isinstance(product, dict):
        line = partial(build_product_line, line_type=backend.ORC_LINE_TYPE_PRODUCT)(product, float(qty_spin.value() or 0))
        product_code = str(product.get("codigo", "") or "").strip()
        preserve_existing_line = bool(wanted_code and product_code == wanted_code)
        if isinstance(line, dict):
            manual_price = round(float(manual_price_spin.value() or 0.0), 4)
            manual_unit = manual_unit_combo.currentText().strip()
            if manual_price > 0 and (manual_state.get("price_dirty") or preserve_existing_line):
                line["preco_unit"] = manual_price
            if manual_unit and (manual_state.get("unit_dirty") or preserve_existing_line):
                line["produto_unid"] = manual_unit
    else:
        manual_desc = combo.currentText().strip()
        if not manual_desc:
            return None
        manual_code = ""
        if " - " in manual_desc:
            manual_code, manual_desc = [chunk.strip() for chunk in manual_desc.split(" - ", 1)]
        line = {
            "tipo_item": backend.ORC_LINE_TYPE_PRODUCT,
            "stock_item_kind": "product",
            "produto_codigo": manual_code,
            "ref_externa": manual_code,
            "descricao": manual_desc,
            "produto_unid": manual_unit_combo.currentText().strip() or "UN",
            "qtd": round(float(qty_spin.value() or 0), 2),
            "tempo_peca_min": 0.0,
            "preco_unit": round(float(manual_price_spin.value() or 0.0), 4),
            "operacao": "Montagem",
            "_product_pending_create": True,
        }
    if not isinstance(line, dict):
        return None
    return {
        "kind": "product",
        "quantity_units": round(float(qty_spin.value() or 0), 2),
        "total_cost": round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2),
        "line": line,
    }
