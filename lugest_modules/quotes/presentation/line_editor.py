from __future__ import annotations
import os
from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QAbstractItemView, QAbstractSpinBox, QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
from pathlib import Path
from lugest_qt.ui.pages.materials_page import _MaterialEditorDialog
from lugest_qt.ui.pages.runtime_support import _build_operation_selector, _fmt_eur, _open_operation_cost_profiles_dialog, _open_quote_operation_detail_dialog, _operation_tokens, _reference_catalog_dialog
from lugest_qt.ui.widgets import FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from typing import Any
from lugest_modules.quotes.presentation.material_prices import edit_material_prices, sync_stock_price
def edit_line(owner: QWidget, backend: QuoteEditorPorts, initial: dict | None = None, *, template_mode: bool = False, presets: dict, client_code: str, current_number: str, line_rows: list[dict]) -> dict | None:
    initial = dict(initial or {})
    dialog = QDialog(owner)
    dialog.setWindowTitle("Item do conjunto" if template_mode else "Linha de orcamento")
    dialog.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
    dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
    dialog.setMinimumSize(860, 560)
    try:
        screen_geo = QApplication.primaryScreen().availableGeometry()
        dialog.resize(min(980, max(860, screen_geo.width() - 120)), min(760, max(560, screen_geo.height() - 110)))
    except Exception:
        dialog.resize(960, 720)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)
    form_host = QWidget()
    form = QFormLayout(form_host)
    form.setContentsMargins(0, 0, 0, 0)
    form.setVerticalSpacing(7)
    client_code = client_code
    references = backend.order_reference_rows("", client_code)
    refs_by_ext = {str(row.get("ref_externa", "") or "").strip(): row for row in references if str(row.get("ref_externa", "") or "").strip()}
    refs_by_int = {str(row.get("ref_interna", "") or "").strip(): row for row in references if str(row.get("ref_interna", "") or "").strip()}
    presets = presets or backend.order_presets()
    product_rows = {str(row.get("codigo", "") or "").strip(): row for row in backend.ne_product_options("")}
    stock_material_rows: list[tuple[dict[str, Any], dict[str, Any]]] = []

    def reload_stock_material_rows() -> None:
        stock_material_rows.clear()
        for item in list(backend.material_rows("") or []):
            if not isinstance(item, dict) or not isinstance(item.get("record"), dict):
                continue
            record = dict(item.get("record") or {})
            preview = dict(item.get("preview") or backend.material_price_preview(record) or {})
            stock_material_rows.append((record, preview))

    reload_stock_material_rows()
    initial_type = str(
        backend.normalize_orc_line_type(
            initial.get("tipo_item", backend.ORC_LINE_TYPE_PIECE)
        )
    )
    if (
        (
            str(initial.get("stock_material_id", "") or "").strip()
            or str(initial.get("stock_item_kind", "") or "").strip() == "raw_material"
        )
        and initial_type == backend.ORC_LINE_TYPE_PIECE
    ):
        initial_type = "stock_mp"
    ref_history = QComboBox()
    ref_history.setEditable(True)
    ref_history.addItem("")
    for row in references:
        ref_history.addItem(
            f"{row.get('ref_interna', '')} | {row.get('ref_externa', '')} | {row.get('descricao', '')}",
            row,
        )
    type_combo = QComboBox()
    type_combo.addItem("Componente fabricado", backend.ORC_LINE_TYPE_PIECE)
    type_combo.addItem("Matéria Prima - Stock", "stock_mp")
    type_combo.addItem("Produto stock", backend.ORC_LINE_TYPE_PRODUCT)
    type_combo.addItem("Servico montagem", backend.ORC_LINE_TYPE_SERVICE)
    for index in range(type_combo.count()):
        if str(type_combo.itemData(index) or "") == initial_type:
            type_combo.setCurrentIndex(index)
            break
    product_combo = QComboBox()
    product_combo.setEditable(True)
    product_combo.addItem("")
    for code, row in sorted(product_rows.items()):
        product_combo.addItem(f"{code} - {row.get('descricao', '')}", code)
    stock_material_combo = QComboBox()
    stock_material_combo.setEditable(True)
    stock_kind_combo = QComboBox()
    stock_kind_combo.addItem("Chapa", "Chapa")
    stock_kind_combo.addItem("Viga / Perfil", "Perfil")
    stock_kind_combo.addItem("Tubo", "Tubo")
    stock_kind_combo.addItem("Cantoneira", "Cantoneira")
    stock_kind_combo.addItem("Barra", "Barra")
    stock_kind_combo.addItem("Ferro nervurado", "Ferro nervurado")
    stock_kind_combo.addItem("Todos", "")

    def _stock_kind(record: dict[str, Any], preview: dict[str, Any]) -> str:
        formato = str(preview.get("formato", record.get("formato", "")) or "").strip().title()
        secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
        material_norm = backend.norm_text(str(record.get("material", "") or ""))
        if formato == "Perfil" and secao == "L":
            return "Cantoneira"
        if formato == "Chapa" and any(token in material_norm for token in ("barra", "chata", "plat", "flat")):
            return "Barra"
        if formato in {"Chapa", "Tubo", "Perfil", "Cantoneira", "Barra", "Ferro Nervurado"}:
            return "Ferro nervurado" if formato == "Ferro Nervurado" else formato
        detected = str(backend.detect_materia_formato(record) or "").strip().title()
        if detected == "Ferro Nervurado":
            return "Ferro nervurado"
        return detected

    def _stock_label(record: dict[str, Any], preview: dict[str, Any]) -> str:
        material_txt = str(record.get("material", "") or "-").strip()
        dim_txt = str(preview.get("dimension_label", "") or "-").strip()
        esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
        qty_txt = backend._fmt(record.get("quantidade", 0))
        price_txt = _fmt_eur(float(preview.get("preco_unid", record.get("preco_unid", 0)) or 0))
        parts = [str(record.get("id", "") or "").strip(), material_txt, dim_txt]
        if esp_txt:
            parts.append(f"{esp_txt} mm")
        if qty_txt:
            parts.append(f"Qtd {qty_txt}")
        if price_txt:
            parts.append(f"Base {price_txt}")
        return " | ".join(part for part in parts if part and part != "-")

    def refresh_stock_material_combo(preferred_id: str = "") -> None:
        current_id = str(preferred_id or "").strip()
        if not current_id:
            payload = stock_material_combo.currentData()
            if isinstance(payload, dict):
                current_id = str(payload.get("id", "") or "").strip()
        wanted_kind = str(stock_kind_combo.currentData() or "").strip()
        stock_material_combo.blockSignals(True)
        stock_material_combo.clear()
        stock_material_combo.addItem("", None)
        for record, preview in stock_material_rows:
            if wanted_kind and _stock_kind(record, preview) != wanted_kind:
                continue
            stock_material_combo.addItem(_stock_label(record, preview), dict(record))
        if current_id:
            for index in range(stock_material_combo.count()):
                payload = stock_material_combo.itemData(index)
                if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == current_id:
                    stock_material_combo.setCurrentIndex(index)
                    break
        stock_material_combo.blockSignals(False)
    initial_product_code = str(initial.get("produto_codigo", "") or "").strip()
    if initial_product_code:
        for index in range(product_combo.count()):
            if str(product_combo.itemData(index) or "").strip() == initial_product_code:
                product_combo.setCurrentIndex(index)
                break
        else:
            product_combo.setCurrentText(initial_product_code)
    ref_int_edit = QLineEdit("" if template_mode else str(initial.get("ref_interna", "") or "").strip())
    ref_ext_edit = QLineEdit(str(initial.get("ref_externa", "") or "").strip())
    desc_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
    dimension_edit = QLineEdit(str(initial.get("dimensao", initial.get("dimensoes", "")) or "").strip())
    product_unid_edit = QLineEdit(str(initial.get("produto_unid", "") or "").strip())
    product_unid_edit.setReadOnly(True)
    mp_metric_spin = QDoubleSpinBox()
    mp_metric_spin.setRange(0.0, 1000000.0)
    mp_metric_spin.setDecimals(4)
    mp_metric_spin.setSuffix(" kg")
    mp_metric_spin.setReadOnly(True)
    mp_metric_spin.setButtonSymbols(QAbstractSpinBox.NoButtons)
    mp_price_kg_spin = QDoubleSpinBox()
    mp_price_kg_spin.setRange(0.0, 1000000.0)
    mp_price_kg_spin.setDecimals(4)
    mp_price_kg_spin.setPrefix("EUR ")
    mp_price_kg_spin.setSuffix("/kg")
    mp_margin_spin = QDoubleSpinBox()
    mp_margin_spin.setRange(-100.0, 1000000.0)
    mp_margin_spin.setDecimals(2)
    mp_margin_spin.setSuffix(" %")
    mp_margin_spin.setValue(float(initial.get("price_markup_pct", initial.get("markup_pct", 0)) or 0.0))
    mp_price_info = QLabel("")
    mp_price_info.setWordWrap(True)
    mp_price_info.setProperty("role", "muted")
    mp_price_table_btn = QPushButton("Tabela preços matéria-prima")
    mp_price_table_btn.setProperty("variant", "secondary")
    stock_create_btn = QPushButton("Criar material")
    stock_create_btn.setProperty("variant", "secondary")
    stock_tools_layout = QHBoxLayout()
    stock_tools_layout.setContentsMargins(0, 0, 0, 0)
    stock_tools_layout.setSpacing(8)
    stock_tools_layout.addWidget(stock_create_btn)
    stock_tools_layout.addWidget(mp_price_table_btn)
    stock_tools_layout.addStretch(1)
    stock_tools_host = QWidget()
    stock_tools_host.setLayout(stock_tools_layout)
    stock_price_state = {
        "base_label": "EUR/kg",
        "metric_label": "Kg por unid.",
        "metric_suffix": " kg",
        "metric_value": 0.0,
        "base_value": 0.0,
        "price_kg_equiv": 0.0,
        "busy": False,
        "price_user_touched": False,
    }
    stock_line_meta = {
        "stock_material_id": str(initial.get("stock_material_id", "") or "").strip(),
        "material_family": str(initial.get("material_family", "") or "").strip(),
        "material_subtype": str(initial.get("material_subtype", "") or "").strip(),
    }
    if stock_line_meta["stock_material_id"]:
        for record, preview in stock_material_rows:
            if str(record.get("id", "") or "").strip() == stock_line_meta["stock_material_id"]:
                kind = _stock_kind(record, preview)
                for index in range(stock_kind_combo.count()):
                    if str(stock_kind_combo.itemData(index) or "").strip() == kind:
                        stock_kind_combo.setCurrentIndex(index)
                        break
                break
    refresh_stock_material_combo(stock_line_meta["stock_material_id"])
    material_combo = QComboBox()
    material_combo.setEditable(True)
    for value in list(presets.get("materiais", []) or []):
        material_combo.addItem(str(value))
    material_combo.setCurrentText(str(initial.get("material", "") or "").strip())
    esp_combo = QComboBox()
    esp_combo.setEditable(True)
    for value in list(presets.get("espessuras", []) or []):
        esp_combo.addItem(str(value))
    esp_combo.setCurrentText(str(initial.get("espessura", "") or "").strip())
    def _normalize_ops_from_any(value: Any) -> list[str]:
        items = backend.quote_parse_operacoes_lista(value)
        ordered: list[str] = []
        for raw_name in list(items or []):
            normalized = str(backend.normalize_operacao_nome(raw_name) or raw_name or "").strip()
            if normalized and normalized not in ordered:
                ordered.append(normalized)
        return ordered

    def _display_operation_text(value: Any, *, has_laser_base: bool) -> str:
        ops = _normalize_ops_from_any(value)
        if has_laser_base:
            ops = [op_name for op_name in ops if op_name != "Corte Laser"]
        return " + ".join(ops)

    operation_change_state = {"user_initiated": False, "last_prompt_signature": ""}
    operation_selector, operation_edit, apply_operations = _build_operation_selector(
        list(presets.get("operacoes", []) or []),
        "",
        on_change=lambda _text, user_initiated: operation_change_state.__setitem__("user_initiated", bool(user_initiated)),
    )
    tempo_spin = QDoubleSpinBox()
    tempo_spin.setRange(0.0, 1000000.0)
    tempo_spin.setDecimals(2)
    tempo_spin.setValue(float(initial.get("tempo_peca_min", initial.get("tempo_pecas_min", 0)) or 0))
    qtd_spin = QDoubleSpinBox()
    qtd_spin.setRange(0.0, 1000000.0)
    qtd_spin.setDecimals(2)
    qtd_spin.setValue(float(initial.get("qtd", 1) or 1))
    price_spin = QDoubleSpinBox()
    price_spin.setRange(0.0, 1000000.0)
    price_spin.setDecimals(4)
    price_spin.setValue(float(initial.get("preco_unit", 0) or 0))
    drawing_edit = QLineEdit(str(initial.get("desenho", "") or "").strip())
    pdf_docs: list[str] = []
    for raw_doc in [initial.get("desenho_pdf", ""), *list(initial.get("desenhos_pdf", []) or [])]:
        doc_txt = str(raw_doc or "").strip()
        if doc_txt and doc_txt.lower().endswith(".pdf") and doc_txt not in pdf_docs:
            pdf_docs.append(doc_txt)
    for raw_doc in list(initial.get("ficheiros", []) or []):
        doc_txt = str(raw_doc or "").strip()
        if doc_txt and doc_txt.lower().endswith(".pdf") and doc_txt not in pdf_docs:
            pdf_docs.append(doc_txt)
    initial_ref_int = str(initial.get("ref_interna", "") or "").strip()
    operation_meta = {
        "operacoes_lista": list(initial.get("operacoes_lista", []) or []),
        "operacoes_fluxo": [dict(item or {}) for item in list(initial.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
        "operacoes_detalhe": [dict(item or {}) for item in list(initial.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
        "tempos_operacao": dict(initial.get("tempos_operacao", {}) or {}),
        "custos_operacao": dict(initial.get("custos_operacao", {}) or {}),
        "quote_cost_snapshot": dict(initial.get("quote_cost_snapshot", {}) or {}),
    }
    operation_cost_label = QLabel("")
    operation_cost_label.setWordWrap(True)
    operation_cost_label.setProperty("role", "muted")

    def current_type_token() -> str:
        return str(type_combo.currentData() or backend.ORC_LINE_TYPE_PIECE)

    def current_line_type() -> str:
        token = current_type_token()
        if token == "stock_mp":
            return backend.ORC_LINE_TYPE_PIECE
        return token

    def _match_stock_material_payload() -> dict[str, Any] | None:
        payload = stock_material_combo.currentData()
        if isinstance(payload, dict):
            return dict(payload)
        probe = stock_material_combo.currentText().strip().lower()
        if not probe:
            return None
        for index in range(1, stock_material_combo.count()):
            candidate = stock_material_combo.itemData(index)
            if not isinstance(candidate, dict):
                continue
            label = stock_material_combo.itemText(index).strip().lower()
            material_id = str(candidate.get("id", "") or "").strip().lower()
            if probe == label or (material_id and (probe == material_id or probe.startswith(material_id))):
                stock_material_combo.setCurrentIndex(index)
                return dict(candidate)
        return None

    def stock_material_detail() -> tuple[dict[str, Any], dict[str, Any]]:
        payload = _match_stock_material_payload()
        if not isinstance(payload, dict):
            return {}, {}
        preview = dict(backend.material_price_preview(payload) or {})
        return dict(payload), preview

    def stock_pricing_context(record: dict[str, Any] | None = None, preview: dict[str, Any] | None = None) -> dict[str, Any]:
        record = dict(record or {})
        preview = dict(preview or {})
        formato = str(preview.get("formato", record.get("formato", "")) or "").strip() or "Chapa"
        base_label = str(preview.get("base_label", "EUR/m" if formato == "Tubo" else "EUR/kg") or "EUR/kg").strip()
        metric_label = "Metros por unid." if base_label == "EUR/m" else "Kg por unid."
        metric_suffix = " m" if base_label == "EUR/m" else " kg"
        metric_value = float(
            (
                preview.get("metros", record.get("metros", 0))
                if base_label == "EUR/m"
                else preview.get("peso_unid", record.get("peso_unid", 0))
            )
            or 0.0
        )
        kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
        base_value = float(record.get("p_compra", 0) or 0.0)
        price_kg_equiv = round((base_value / kg_m), 4) if base_label == "EUR/m" and kg_m > 0 else round(base_value, 4)
        return {
            "formato": formato,
            "base_label": base_label,
            "metric_label": metric_label,
            "metric_suffix": metric_suffix,
            "metric_value": round(metric_value, 4),
            "base_value": round(base_value, 4),
            "price_kg_equiv": price_kg_equiv,
            "kg_m": round(kg_m, 4),
        }

    def set_form_row_visible(field: QWidget, visible: bool, label_text: str | None = None) -> None:
        field.setVisible(visible)
        label_widget = form.labelForField(field)
        if label_widget is not None:
            label_widget.setVisible(visible)
            if label_text is not None:
                label_widget.setText(label_text)

    def current_piece_metric_kg() -> float:
        weight_total = float(initial.get("weight_total", 0) or 0.0)
        quantity_units = float(initial.get("quantity_units", initial.get("qtd", 1)) or 1.0)
        if weight_total > 0 and quantity_units > 0:
            return round(weight_total / quantity_units, 4)
        kg_m = float(initial.get("kg_per_m", 0) or 0.0)
        meters = float(initial.get("meters_per_unit", 0) or 0.0)
        if kg_m > 0 and meters > 0:
            return round(kg_m * meters, 4)
        stock_record = None
        stock_preview = None
        if str(stock_line_meta.get("stock_material_id", "") or "").strip():
            stock_record = backend.material_by_id(str(stock_line_meta.get("stock_material_id", "") or "").strip())
            stock_preview = dict(backend.material_price_preview(stock_record) or {}) if isinstance(stock_record, dict) else {}
        if isinstance(stock_preview, dict):
            return round(float(stock_preview.get("peso_unid", stock_record.get("peso_unid", 0) if isinstance(stock_record, dict) else 0) or 0.0), 4)
        return 0.0

    def _payload_has_laser_base(payload_row: dict[str, Any] | None) -> bool:
        source = dict(payload_row or {})
        if bool(source.get("laser_base_active", False)):
            return True
        if current_line_type() != backend.ORC_LINE_TYPE_PIECE:
            return False
        if not str(source.get("desenho", "") or "").strip():
            return False
        if backend._parse_float(source.get("tempo_peca_min", source.get("tempo_pecas_min", 0)), 0) <= 0 and backend._parse_float(source.get("preco_unit", 0), 0) <= 0:
            return False
        return "Corte Laser" in _normalize_ops_from_any(source.get("operacao", source.get("operacoes", source.get("operacoes_lista", []))))

    base_state = {
        "laser_base_enabled": _payload_has_laser_base(initial),
        "laser_base_time": round(float(initial.get("laser_base_tempo_unit", initial.get("tempo_peca_min", initial.get("tempo_pecas_min", 0))) or 0), 4),
        "laser_base_price": round(float(initial.get("laser_base_preco_unit", initial.get("preco_unit", 0)) or 0), 4),
    }
    apply_operations(_display_operation_text(initial.get("operacao", initial.get("operacoes_lista", [])), has_laser_base=bool(base_state["laser_base_enabled"])))

    def current_line_refs() -> list[str]:
        refs: list[str] = []
        skipped_current = False
        for row in list(line_rows):
            ref_txt = str((row or {}).get("ref_interna", "") or "").strip()
            if not ref_txt:
                continue
            if initial_ref_int and not skipped_current and ref_txt == initial_ref_int:
                skipped_current = True
                continue
            refs.append(ref_txt)
        return refs

    def current_product_code() -> str:
        raw_code = str(product_combo.currentData() or "").strip()
        if raw_code:
            return raw_code
        raw_text = product_combo.currentText().strip()
        candidate = raw_text.split(" - ", 1)[0].strip()
        return candidate if candidate in product_rows else ""

    def apply_reference(payload: dict | None) -> None:
        if not isinstance(payload, dict):
            return
        ref_ext_edit.setText(str(payload.get("ref_externa", "") or "").strip())
        if not template_mode:
            ref_int_edit.setText(str(payload.get("ref_interna", "") or "").strip())
        desc_edit.setText(str(payload.get("descricao", "") or "").strip())
        material_combo.setCurrentText(str(payload.get("material", "") or "").strip())
        esp_combo.setCurrentText(str(payload.get("espessura", "") or "").strip())
        base_state["laser_base_enabled"] = _payload_has_laser_base(payload)
        base_state["laser_base_time"] = round(
            float(payload.get("laser_base_tempo_unit", payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0))) or 0),
            4,
        )
        base_state["laser_base_price"] = round(float(payload.get("laser_base_preco_unit", payload.get("preco_unit", payload.get("preco", 0))) or 0), 4)
        operation_meta["operacoes_lista"] = list(payload.get("operacoes_lista", []) or [])
        operation_meta["operacoes_fluxo"] = [dict(item or {}) for item in list(payload.get("operacoes_fluxo", []) or []) if isinstance(item, dict)]
        operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        operation_meta["tempos_operacao"] = dict(payload.get("tempos_operacao", {}) or {})
        operation_meta["custos_operacao"] = dict(payload.get("custos_operacao", {}) or {})
        operation_meta["quote_cost_snapshot"] = dict(payload.get("quote_cost_snapshot", {}) or {})
        apply_operations(_display_operation_text(payload.get("operacoes", payload.get("operacao", payload.get("operacoes_lista", []))), has_laser_base=bool(base_state["laser_base_enabled"])))
        tempo_spin.setValue(float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)) or 0))
        price_spin.setValue(float(payload.get("preco_unit", payload.get("preco", 0)) or 0))
        if not drawing_edit.text().strip():
            drawing_edit.setText(str(payload.get("desenho", "") or "").strip())

    def sync_product_fields() -> None:
        code = current_product_code()
        row = product_rows.get(code)
        if row is None:
            return
        product_unid_edit.setText(str(row.get("unid", "") or "UN").strip())
        if not desc_edit.text().strip() or current_line_type() == backend.ORC_LINE_TYPE_PRODUCT:
            desc_edit.setText(str(row.get("descricao", "") or "").strip())
        sale_price = float(row.get("preco_venda", row.get("pvp1", row.get("preco", 0))) or 0)
        if current_line_type() == backend.ORC_LINE_TYPE_PRODUCT or float(price_spin.value() or 0) <= 0:
            price_spin.setValue(sale_price)

    def sync_line_price_from_mp() -> None:
        stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
        if not stock_id:
            return
        if stock_price_state["busy"]:
            return
        metric_value = float(mp_metric_spin.value() or stock_price_state.get("metric_value", 0.0) or 0.0)
        base_value = float(mp_price_kg_spin.value() or 0.0)
        if metric_value <= 0 or base_value <= 0:
            return
        sale_unit = round(metric_value * base_value * (1.0 + (float(mp_margin_spin.value() or 0.0) / 100.0)), 4)
        stock_price_state["busy"] = True
        try:
            price_spin.setValue(sale_unit)
        finally:
            stock_price_state["busy"] = False

    def sync_margin_from_line_price() -> None:
        stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
        if not stock_id or stock_price_state["busy"]:
            return
        metric_value = float(mp_metric_spin.value() or stock_price_state.get("metric_value", 0.0) or 0.0)
        base_value = float(mp_price_kg_spin.value() or 0.0)
        cost_unit = metric_value * base_value
        if cost_unit <= 0:
            return
        stock_price_state["busy"] = True
        try:
            mp_margin_spin.setValue(round(((float(price_spin.value() or 0.0) / cost_unit) - 1.0) * 100.0, 2))
        finally:
            stock_price_state["busy"] = False

    def sync_stock_material_fields() -> None:
        record, preview = stock_material_detail()
        if not record:
            stock_line_meta["stock_material_id"] = ""
            mp_metric_spin.setValue(0.0)
            mp_price_kg_spin.setValue(0.0)
            mp_margin_spin.setValue(0.0)
            mp_price_info.setText("")
            return
        previous_stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
        current_stock_id = str(record.get("id", "") or "").strip()
        if current_stock_id and current_stock_id != previous_stock_id:
            stock_price_state["price_user_touched"] = False
        desc_txt = str(record.get("material", "") or "").strip()
        dim_txt = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
        esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
        if not esp_txt:
            esp_txt = dim_txt
        if dim_txt:
            desc_txt = f"{desc_txt} {dim_txt}".strip()
            dimension_edit.setText(dim_txt)
        if esp_txt and esp_txt not in desc_txt:
            desc_txt = f"{desc_txt} {esp_txt} mm".strip()
        material_combo.setCurrentText(str(record.get("material", "") or "").strip())
        if esp_txt:
            esp_combo.setCurrentText(esp_txt)
        if not desc_edit.text().strip() or current_type_token() == "stock_mp":
            desc_edit.setText(desc_txt)
        if not ref_ext_edit.text().strip() or current_type_token() == "stock_mp":
            ref_ext_edit.setText(str(record.get("id", "") or "").strip() or desc_txt)
        stock_line_meta["stock_material_id"] = current_stock_id
        stock_line_meta["material_family"] = str(record.get("material", "") or "").strip()
        stock_line_meta["material_subtype"] = str(preview.get("formato", record.get("formato", "")) or "Stock MP").strip()
        if float(qtd_spin.value() or 0) <= 0:
            qtd_spin.setValue(1.0)
        pricing = stock_pricing_context(record, preview)
        stock_price_state.update(pricing)
        metric_label = form.labelForField(mp_metric_spin)
        if metric_label is not None:
            metric_label.setText(pricing["metric_label"])
        price_label = form.labelForField(mp_price_kg_spin)
        if price_label is not None:
            price_label.setText(f"Preco compra ({pricing['base_label']})")
        mp_metric_spin.setSuffix(pricing["metric_suffix"])
        mp_metric_spin.setValue(float(pricing["metric_value"] or 0.0))
        mp_price_kg_spin.setSuffix("/m" if pricing["base_label"] == "EUR/m" else "/kg")
        if not bool(stock_price_state.get("price_user_touched", False)):
            mp_price_kg_spin.setValue(float(pricing["base_value"] or 0.0))
        if float(initial.get("preco_unit", 0) or 0.0) > 0 and float(qtd_spin.value() or 0.0) > 0:
            cost_unit = float(pricing["metric_value"] or 0.0) * float(pricing["base_value"] or 0.0)
            if cost_unit > 0 and float(initial.get("price_markup_pct", 0) or 0.0) == 0.0:
                mp_margin_spin.setValue(round(((float(initial.get("preco_unit", 0) or 0.0) / cost_unit) - 1.0) * 100.0, 2))
        sync_line_price_from_mp()
        mp_price_info.setText(
            f"Matéria-prima ligada: {stock_line_meta['stock_material_id']} | "
            f"Base {pricing['base_value']:.4f} {pricing['base_label']} | "
            f"Equiv. {float(pricing['price_kg_equiv'] or 0.0):.4f} EUR/kg"
        )

    def refresh_mp_price_panel() -> None:
        stock_id = str(stock_line_meta.get("stock_material_id", "") or "").strip()
        visible = bool(stock_id) and current_type_token() == "stock_mp"
        tools_visible = current_type_token() == "stock_mp"
        if visible:
            record = backend.material_by_id(stock_id)
            preview = dict(backend.material_price_preview(record) or {}) if isinstance(record, dict) else {}
            pricing = stock_pricing_context(record, preview)
            stock_price_state.update(pricing)
            metric_label = form.labelForField(mp_metric_spin)
            if metric_label is not None:
                metric_label.setText(pricing["metric_label"])
            price_label = form.labelForField(mp_price_kg_spin)
            if price_label is not None:
                price_label.setText(f"Preco compra ({pricing['base_label']})")
            mp_metric_spin.setSuffix(pricing["metric_suffix"])
            mp_price_kg_spin.setSuffix("/m" if pricing["base_label"] == "EUR/m" else "/kg")
            mp_metric_spin.setValue(float(pricing["metric_value"] or 0.0))
            if bool(stock_price_state.get("price_user_touched", False)):
                pass
            elif float(initial.get("price_base_value", 0) or 0.0) > 0:
                mp_price_kg_spin.setValue(float(initial.get("price_base_value", 0) or 0.0))
            elif pricing["base_value"] > 0:
                mp_price_kg_spin.setValue(float(pricing["base_value"] or 0.0))
            mp_price_info.setText(
                f"Matéria-prima ligada: {stock_id} | {str(preview.get('formato', record.get('formato', '-')) if isinstance(record, dict) else '-')}"
                f" | Base {float(pricing['base_value'] or 0.0):.4f} {pricing['base_label']} | "
                f"Equiv. {float(pricing['price_kg_equiv'] or 0.0):.4f} EUR/kg"
            )
            sync_line_price_from_mp()
        else:
            mp_metric_spin.setValue(0.0)
            mp_price_kg_spin.setValue(0.0)
            mp_margin_spin.setValue(0.0)
            mp_price_info.setText("")
        for widget in (mp_metric_spin, mp_price_kg_spin, mp_margin_spin, mp_price_info):
            widget.setVisible(visible)
        stock_tools_host.setVisible(tools_visible)
        for label_widget in (
            form.labelForField(mp_metric_spin),
            form.labelForField(mp_price_kg_spin),
            form.labelForField(mp_margin_spin),
        ):
            if label_widget is not None:
                label_widget.setVisible(visible)

    def load_ref() -> None:
        selected = ref_history.currentData()
        if isinstance(selected, dict):
            apply_reference(selected)
            return
        key = ref_ext_edit.text().strip() or ref_history.currentText().strip()
        apply_reference(refs_by_ext.get(key) or refs_by_int.get(key))

    def browse_refs() -> None:
        payload = _reference_catalog_dialog(
            owner,
            backend.order_reference_rows("", ""),
            "Historico de referencias",
            backend=backend,
            current_client=client_code,
        )
        apply_reference(payload)

    def generate_ref() -> None:
        if template_mode or current_line_type() != backend.ORC_LINE_TYPE_PIECE:
            return
        ref_int_edit.setText(
            backend.orc_suggest_ref_interna(
                client_code,
                existing_refs=current_line_refs(),
                numero=current_number,
            )
        )

    def pick_drawing() -> None:
        path, _ = QFileDialog.getOpenFileName(
            owner,
            "Selecionar desenho",
            "",
            "Desenhos (*.pdf *.dwg *.dxf *.step *.stp *.iges *.igs *.png *.jpg *.jpeg *.bmp);;Todos (*.*)",
        )
        if path:
            drawing_edit.setText(path)

    def create_stock_material() -> None:
        dialog_editor = _MaterialEditorDialog(backend, dialog)
        current_kind = str(stock_kind_combo.currentData() or "").strip()
        if current_kind:
            dialog_editor.formato_combo.setCurrentText("Perfil" if current_kind == "Perfil" else current_kind)
        if dialog_editor.exec() != QDialog.Accepted:
            return
        try:
            record = backend.add_material(dialog_editor.payload())
        except Exception as exc:
            QMessageBox.critical(dialog, "Matéria Prima - Stock", str(exc))
            return
        reload_stock_material_rows()
        created_id = str(record.get("id", "") or "").strip()
        if created_id:
            preview = dict(backend.material_price_preview(record) or {})
            kind = _stock_kind(dict(record), preview)
            for index in range(stock_kind_combo.count()):
                if str(stock_kind_combo.itemData(index) or "").strip() == kind:
                    stock_kind_combo.setCurrentIndex(index)
                    break
        refresh_stock_material_combo(created_id)
        sync_stock_material_fields()

    docs_label = QLabel("")
    docs_label.setProperty("role", "muted")
    docs_table = QTableWidget(0, 3)
    docs_table.setHorizontalHeaderLabels(["PDF", "Estado", "Caminho"])
    docs_table.verticalHeader().setVisible(False)
    docs_table.setEditTriggers(QTableWidget.NoEditTriggers)
    docs_table.setSelectionBehavior(QTableWidget.SelectRows)
    docs_table.setSelectionMode(QAbstractItemView.SingleSelection)
    docs_table.setMinimumHeight(118)
    docs_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
    docs_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
    docs_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)

    def _resolve_pdf_doc_path(raw: str) -> Path:
        resolver = getattr(backend, "_resolve_file_reference", None)
        if callable(resolver):
            try:
                resolved = resolver(raw)
                if resolved is not None:
                    return Path(resolved)
            except Exception:
                pass
        return Path(raw)

    def _selected_doc_index() -> int:
        current = docs_table.currentItem()
        if current is None or current.row() >= len(pdf_docs):
            return -1
        return current.row()

    def _render_pdf_docs() -> None:
        docs_table.setRowCount(len(pdf_docs))
        for row_index, pdf_path in enumerate(pdf_docs):
            path_obj = _resolve_pdf_doc_path(pdf_path)
            values = [path_obj.name or pdf_path, "OK" if path_obj.exists() else "Em falta", pdf_path]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if col_index == 1:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    item.setForeground(QBrush(QColor("#4f7f1f" if path_obj.exists() else "#b45f06")))
                docs_table.setItem(row_index, col_index, item)
        docs_label.setText(
            f"{len(pdf_docs)} PDF(s) associado(s)." if pdf_docs else "Sem PDFs técnicos associados a esta peça."
        )

    def _add_pdf_docs() -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            dialog,
            "Adicionar PDFs técnicos",
            "",
            "PDF (*.pdf);;Todos (*.*)",
        )
        for path in paths:
            clean = str(path or "").strip()
            if clean and clean.lower().endswith(".pdf") and clean not in pdf_docs:
                pdf_docs.append(clean)
        _render_pdf_docs()

    def _open_pdf_doc() -> None:
        index = _selected_doc_index()
        if index < 0:
            QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
            return
        path = _resolve_pdf_doc_path(pdf_docs[index])
        if not path.exists():
            QMessageBox.critical(dialog, "PDFs associados", f"PDF não encontrado:\n{path}")
            return
        os.startfile(str(path))

    def _download_pdf_doc() -> None:
        index = _selected_doc_index()
        if index < 0:
            QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
            return
        source = _resolve_pdf_doc_path(pdf_docs[index])
        if not source.exists():
            QMessageBox.critical(dialog, "PDFs associados", f"PDF não encontrado:\n{source}")
            return
        target, _ = QFileDialog.getSaveFileName(dialog, "Guardar cópia do PDF", source.name, "PDF (*.pdf)")
        if not target:
            return
        try:
            import shutil

            shutil.copy2(source, target)
        except Exception as exc:
            QMessageBox.critical(dialog, "PDFs associados", str(exc))

    def _replace_pdf_doc() -> None:
        index = _selected_doc_index()
        if index < 0:
            QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
            return
        path, _ = QFileDialog.getOpenFileName(dialog, "Substituir PDF", "", "PDF (*.pdf);;Todos (*.*)")
        clean = str(path or "").strip()
        if clean:
            pdf_docs[index] = clean
            _render_pdf_docs()
            docs_table.selectRow(index)

    def _remove_pdf_doc() -> None:
        index = _selected_doc_index()
        if index < 0:
            QMessageBox.information(dialog, "PDFs associados", "Seleciona um PDF.")
            return
        del pdf_docs[index]
        _render_pdf_docs()

    docs_buttons = QHBoxLayout()
    docs_buttons.setSpacing(8)
    btn_doc_open = QPushButton("Visualizar")
    btn_doc_open.setProperty("variant", "secondary")
    btn_doc_download = QPushButton("Download")
    btn_doc_download.setProperty("variant", "secondary")
    btn_doc_add = QPushButton("Adicionar PDF")
    btn_doc_replace = QPushButton("Substituir")
    btn_doc_replace.setProperty("variant", "secondary")
    btn_doc_remove = QPushButton("Apagar")
    btn_doc_remove.setProperty("variant", "danger")
    for button in (btn_doc_open, btn_doc_download, btn_doc_add, btn_doc_replace, btn_doc_remove):
        docs_buttons.addWidget(button)
    docs_buttons.addStretch(1)
    docs_buttons_host = QWidget()
    docs_buttons_host.setLayout(docs_buttons)
    docs_host = QWidget()
    docs_layout = QVBoxLayout(docs_host)
    docs_layout.setContentsMargins(0, 0, 0, 0)
    docs_layout.setSpacing(6)
    docs_layout.addWidget(docs_label)
    docs_layout.addWidget(docs_table)
    docs_layout.addWidget(docs_buttons_host)
    btn_doc_open.clicked.connect(_open_pdf_doc)
    btn_doc_download.clicked.connect(_download_pdf_doc)
    btn_doc_add.clicked.connect(_add_pdf_docs)
    btn_doc_replace.clicked.connect(_replace_pdf_doc)
    btn_doc_remove.clicked.connect(_remove_pdf_doc)
    _render_pdf_docs()

    def selected_operation_names() -> list[str]:
        selected_ops: list[str] = []
        for token in _operation_tokens(operation_edit.text().strip()):
            normalized = str(backend.normalize_operacao_nome(token) or token or "").strip()
            if normalized and normalized not in selected_ops:
                selected_ops.append(normalized)
        return selected_ops

    def _normalized_op_key(value: str) -> str:
        return str(backend.normalize_operacao_nome(value) or value or "").strip()

    def line_has_laser_base() -> bool:
        if current_line_type() != backend.ORC_LINE_TYPE_PIECE:
            return False
        if not bool(base_state.get("laser_base_enabled", False)):
            return False
        if not drawing_edit.text().strip():
            return False
        return float(tempo_spin.value() or 0) > 0 or float(price_spin.value() or 0) > 0

    def full_route_operation_names() -> list[str]:
        route = list(selected_operation_names())
        if line_has_laser_base() and "Corte Laser" not in route:
            route.insert(0, "Corte Laser")
        return route

    def costing_operation_names() -> list[str]:
        selected_ops = full_route_operation_names()
        if line_has_laser_base():
            return [op_name for op_name in selected_ops if op_name != "Corte Laser"]
        return list(selected_ops)

    def current_extra_totals() -> tuple[float, float]:
        extra_ops = {_normalized_op_key(op_name) for op_name in costing_operation_names() if _normalized_op_key(op_name)}
        tempo_extra = 0.0
        custo_extra = 0.0
        for op_name, raw_value in dict(operation_meta.get("tempos_operacao", {}) or {}).items():
            normalized = _normalized_op_key(str(op_name or ""))
            if normalized in extra_ops:
                tempo_extra += float(raw_value or 0)
        for op_name, raw_value in dict(operation_meta.get("custos_operacao", {}) or {}).items():
            normalized = _normalized_op_key(str(op_name or ""))
            if normalized in extra_ops:
                custo_extra += float(raw_value or 0)
        return round(tempo_extra, 4), round(custo_extra, 4)

    def current_laser_base_totals() -> tuple[float, float]:
        if not line_has_laser_base():
            return float(tempo_spin.value() or 0), float(price_spin.value() or 0)
        return (
            round(float(base_state.get("laser_base_time", 0) or 0), 4),
            round(float(base_state.get("laser_base_price", 0) or 0), 4),
        )

    def current_composed_totals() -> tuple[float, float]:
        if not line_has_laser_base():
            return round(float(tempo_spin.value() or 0), 4), round(float(price_spin.value() or 0), 4)
        base_time, base_cost = current_laser_base_totals()
        extra_time, extra_cost = current_extra_totals()
        return round(base_time + extra_time, 4), round(base_cost + extra_cost, 4)

    def sync_line_totals_from_meta() -> None:
        if not line_has_laser_base():
            return
        composed_time, composed_cost = current_composed_totals()
        previous_tempo_state = tempo_spin.blockSignals(True)
        previous_price_state = price_spin.blockSignals(True)
        try:
            tempo_spin.setValue(float(composed_time or 0))
            price_spin.setValue(float(composed_cost or 0))
        finally:
            tempo_spin.blockSignals(previous_tempo_state)
            price_spin.blockSignals(previous_price_state)

    def sync_operation_meta_selection() -> None:
        selected_ops = full_route_operation_names()
        selected_keys = {_normalized_op_key(op_name) for op_name in selected_ops if _normalized_op_key(op_name)}
        previous_keys = {_normalized_op_key(op_name) for op_name in list(operation_meta.get("operacoes_lista", []) or []) if _normalized_op_key(op_name)}
        operation_meta["operacoes_lista"] = list(selected_ops)
        operation_meta["operacoes_fluxo"] = backend.build_operacoes_fluxo(
            selected_ops,
            operation_meta.get("operacoes_fluxo") if isinstance(operation_meta.get("operacoes_fluxo"), list) else None,
        )
        operation_meta["operacoes_detalhe"] = [
            dict(item or {})
            for item in list(operation_meta.get("operacoes_detalhe", []) or [])
            if _normalized_op_key(str((item or {}).get("nome", "") or "")) in selected_keys
        ]
        operation_meta["tempos_operacao"] = {
            _normalized_op_key(str(op_name or "")): float(raw_value or 0)
            for op_name, raw_value in dict(operation_meta.get("tempos_operacao", {}) or {}).items()
            if _normalized_op_key(str(op_name or "")) in selected_keys
        }
        operation_meta["custos_operacao"] = {
            _normalized_op_key(str(op_name or "")): float(raw_value or 0)
            for op_name, raw_value in dict(operation_meta.get("custos_operacao", {}) or {}).items()
            if _normalized_op_key(str(op_name or "")) in selected_keys
        }
        if previous_keys != selected_keys:
            operation_meta["quote_cost_snapshot"] = {}
            operation_change_state["last_prompt_signature"] = ""

    def current_operation_cost_payload() -> dict[str, Any]:
        base_blend = line_has_laser_base()
        base_time, base_cost = current_laser_base_totals()
        return {
            "operacao": operation_edit.text().strip(),
            "operacoes_lista": list(selected_operation_names()),
            "costing_operations": list(costing_operation_names()),
            "qtd": qtd_spin.value(),
            "tempo_peca_min": tempo_spin.value(),
            "preco_unit": price_spin.value(),
            "area_m2": float(initial.get("area_m2", initial.get("net_area_m2", 0)) or 0),
            "blend_with_current_line": base_blend,
            "base_tempo_unit_min": base_time if base_blend else 0.0,
            "base_preco_unit_eur": base_cost if base_blend else 0.0,
            "base_operation_label": "Laser base",
            "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
        }

    def pending_operation_inputs(estimate: dict[str, Any]) -> list[dict[str, Any]]:
        pending_rows: list[dict[str, Any]] = []
        for item in list(estimate.get("operations", []) or []):
            if not isinstance(item, dict):
                continue
            if bool(item.get("missing_driver_input")):
                pending_rows.append(dict(item))
        return pending_rows

    def refresh_operation_cost_hint() -> None:
        ops_txt = operation_edit.text().strip()
        if not ops_txt:
            operation_cost_label.setText("Sem postos selecionados nesta linha.")
            return
        estimate = dict(backend.operation_cost_estimate(current_operation_cost_payload()) or {})
        summary = dict(estimate.get("summary", {}) or {})
        pending_rows = pending_operation_inputs(estimate)
        base_blend = bool(current_operation_cost_payload().get("blend_with_current_line"))
        extra_ops = list(costing_operation_names())
        state_txt = {
            "detailed": "detalhe completo",
            "partial_detail": "detalhe parcial",
            "aggregate_pending": "preco ainda agregado",
            "single_operation_total": "operacao simples",
        }.get(str(summary.get("costing_mode", "") or ""), "sem detalhe")
        if base_blend and not extra_ops:
            base_time, base_cost = current_laser_base_totals()
            operation_cost_label.setText(
                f"Base laser carregada na linha: {_fmt_eur(base_cost)}/un | {base_time:.3f} min/un. "
                "Seleciona as operacoes seguintes para agregar custo."
            )
            return
        if pending_rows:
            missing_txt = ", ".join(
                f"{str(item.get('nome', '') or '').strip()} ({str(item.get('driver_label', '') or 'Qtd./peca').strip()})"
                for item in pending_rows
                if str(item.get("nome", "") or "").strip()
            )
            operation_cost_label.setText(
                f"Falta quantificar: {missing_txt}. Abre a quantificacao para fechar o custo e somar ao valor do laser."
            )
            return
        if base_blend:
            base_time, base_cost = current_laser_base_totals()
            extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
            extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
            total_time = base_time + extra_time
            total_cost = base_cost + extra_cost
            operation_cost_label.setText(
                f"Perfis operacao: {str(estimate.get('active_profile', '') or 'Base')} | "
                f"base atual {_fmt_eur(base_cost)}/un + extras {_fmt_eur(extra_cost)}/un = {_fmt_eur(total_cost)}/un | "
                f"tempo {base_time:.3f} + {extra_time:.3f} = {total_time:.3f} min/un"
            )
            return
        operation_cost_label.setText(
            f"Perfis operacao: {str(estimate.get('active_profile', '') or 'Base')} | {state_txt} | "
            f"sugestao {float(summary.get('tempo_unit_total_min', 0) or 0):.3f} min/un | "
            f"{_fmt_eur(float(summary.get('custo_unit_total_eur', 0) or 0))}/un"
        )

    def edit_operation_costs(*, auto_prompt: bool = False) -> bool:
        if line_has_laser_base() and not costing_operation_names():
            if not auto_prompt:
                QMessageBox.information(
                    owner,
                    "Operacoes",
                    "Esta linha ja tem o laser como base. Seleciona as operacoes seguintes para quantificar os acrescimos.",
                )
            return False
        result = _open_quote_operation_detail_dialog(owner, backend, current_operation_cost_payload())
        if not isinstance(result, dict):
            return False
        operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(result.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        operation_meta["tempos_operacao"] = dict(result.get("tempos_operacao", {}) or {})
        operation_meta["custos_operacao"] = dict(result.get("custos_operacao", {}) or {})
        operation_meta["quote_cost_snapshot"] = dict(result.get("quote_cost_snapshot", {}) or {})
        if line_has_laser_base():
            sync_line_totals_from_meta()
        elif bool(result.get("apply_totals")):
            tempo_spin.setValue(float(result.get("suggested_tempo_unit_min", tempo_spin.value()) or 0))
            price_spin.setValue(float(result.get("suggested_preco_unit_eur", price_spin.value()) or 0))
        if auto_prompt:
            operation_change_state["last_prompt_signature"] = operation_edit.text().strip()
        refresh_operation_cost_hint()
        return True

    def maybe_prompt_operation_breakdown() -> None:
        if current_line_type() != backend.ORC_LINE_TYPE_PIECE:
            return
        estimate = dict(backend.operation_cost_estimate(current_operation_cost_payload()) or {})
        pending_rows = pending_operation_inputs(estimate)
        if not pending_rows:
            return
        signature = "|".join(
            [" + ".join(costing_operation_names())]
            + [str(item.get("nome", "") or "").strip() for item in pending_rows if str(item.get("nome", "") or "").strip()]
        )
        if signature == str(operation_change_state.get("last_prompt_signature", "") or ""):
            return
        operation_change_state["last_prompt_signature"] = signature
        edit_operation_costs(auto_prompt=True)

    def handle_operation_text_changed() -> None:
        user_initiated = bool(operation_change_state.get("user_initiated", False))
        operation_change_state["user_initiated"] = False
        sync_operation_meta_selection()
        sync_line_totals_from_meta()
        refresh_operation_cost_hint()
        if user_initiated:
            maybe_prompt_operation_breakdown()

    ref_history.currentIndexChanged.connect(lambda _index: apply_reference(ref_history.currentData()))
    product_combo.currentTextChanged.connect(lambda _value: sync_product_fields())
    operation_edit.textChanged.connect(lambda _value: handle_operation_text_changed())
    qtd_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())
    tempo_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())
    price_spin.valueChanged.connect(lambda _value: refresh_operation_cost_hint())

    ref_buttons = QHBoxLayout()
    btn_generate = QPushButton("Gerar interna")
    btn_generate.setProperty("variant", "secondary")
    btn_generate.clicked.connect(generate_ref)
    btn_history = QPushButton("Referencias criadas")
    btn_history.setProperty("variant", "secondary")
    btn_history.clicked.connect(browse_refs)
    btn_load = QPushButton("Carregar referencia")
    btn_load.setProperty("variant", "secondary")
    btn_load.clicked.connect(load_ref)
    btn_drawing = QPushButton("Selecionar desenho")
    btn_drawing.setProperty("variant", "secondary")
    btn_drawing.clicked.connect(pick_drawing)
    ref_buttons.addWidget(btn_generate)
    ref_buttons.addWidget(btn_history)
    ref_buttons.addWidget(btn_load)
    ref_buttons.addWidget(btn_drawing)
    ref_buttons.addStretch(1)
    ref_buttons_host = QWidget()
    ref_buttons_host.setLayout(ref_buttons)

    operation_buttons = QHBoxLayout()
    operation_buttons.setSpacing(8)
    btn_ops_detail = QPushButton("Quantificar operacoes seguintes")
    btn_ops_detail.setProperty("variant", "secondary")
    btn_ops_detail.clicked.connect(lambda: edit_operation_costs(auto_prompt=False))
    btn_ops_profiles = QPushButton("Perfis operacoes")
    btn_ops_profiles.setProperty("variant", "secondary")
    def configure_operation_profiles_from_line() -> None:
        if _open_operation_cost_profiles_dialog(owner, backend):
            refresh_operation_cost_hint()
    btn_ops_profiles.clicked.connect(configure_operation_profiles_from_line)
    operation_buttons.addWidget(btn_ops_detail)
    operation_buttons.addWidget(btn_ops_profiles)
    operation_buttons.addStretch(1)
    operation_buttons_host = QWidget()
    operation_buttons_host.setLayout(operation_buttons)

    form.addRow("Tipo", type_combo)
    form.addRow("Historico", ref_history)
    form.addRow("Produto", product_combo)
    form.addRow("Ref. interna", ref_int_edit)
    form.addRow("Ref. externa", ref_ext_edit)
    form.addRow("Descricao", desc_edit)
    form.addRow("Dimensao", dimension_edit)
    form.addRow("Codigo/unid", product_unid_edit)
    form.addRow("Tipo matéria-prima", stock_kind_combo)
    form.addRow("Matéria Prima - Stock", stock_material_combo)
    form.addRow("Kg por unid.", mp_metric_spin)
    form.addRow("Preço matéria-prima", mp_price_kg_spin)
    form.addRow("Margem", mp_margin_spin)
    form.addRow("", stock_tools_host)
    form.addRow("", mp_price_info)
    form.addRow("Material", material_combo)
    form.addRow("Espessura", esp_combo)
    form.addRow("Postos", operation_selector)
    form.addRow("Custeio op.", operation_cost_label)
    form.addRow("", operation_buttons_host)
    form.addRow("Tempo peca (min)", tempo_spin)
    form.addRow("Quantidade", qtd_spin)
    form.addRow("Preco unit.", price_spin)
    form.addRow("Desenho", drawing_edit)
    form.addRow("", ref_buttons_host)
    form.addRow("PDFs associados", docs_host)
    form_scroll = QScrollArea()
    form_scroll.setWidgetResizable(True)
    form_scroll.setFrameShape(QFrame.NoFrame)
    form_scroll.setWidget(form_host)
    layout.addWidget(form_scroll, 1)
    mode_state = {"initialized": False, "last_token": initial_type}

    def sync_mode() -> None:
        token = current_type_token()
        line_type = current_line_type()
        is_stock_mp = token == "stock_mp"
        entered_stock_mp = bool(mode_state["initialized"] and is_stock_mp and mode_state.get("last_token") != "stock_mp")
        is_piece = line_type == backend.ORC_LINE_TYPE_PIECE
        is_product = line_type == backend.ORC_LINE_TYPE_PRODUCT
        ref_history.setEnabled(is_piece)
        ref_int_edit.setEnabled(is_piece and not template_mode and not is_stock_mp)
        ref_ext_edit.setEnabled(is_piece and not is_stock_mp)
        material_combo.setEnabled(is_piece)
        esp_combo.setEnabled(is_piece)
        dimension_edit.setEnabled(is_piece)
        drawing_edit.setEnabled(is_piece and not is_stock_mp)
        btn_generate.setEnabled(is_piece and not template_mode and not is_stock_mp)
        btn_history.setEnabled(is_piece and not is_stock_mp)
        btn_load.setEnabled(is_piece and not is_stock_mp)
        btn_drawing.setEnabled(is_piece and not is_stock_mp)
        btn_ops_detail.setEnabled(is_piece and not is_stock_mp)
        btn_ops_profiles.setEnabled(is_piece and not is_stock_mp)
        product_combo.setEnabled(is_product)
        stock_kind_combo.setEnabled(is_stock_mp)
        stock_material_combo.setEnabled(is_stock_mp)
        set_form_row_visible(stock_kind_combo, is_stock_mp, "Tipo matéria-prima")
        set_form_row_visible(stock_material_combo, is_stock_mp, "Matéria Prima - Stock")
        product_unid_edit.setEnabled(is_product or line_type == backend.ORC_LINE_TYPE_SERVICE)
        set_form_row_visible(ref_history, is_piece and not is_stock_mp, "Historico")
        set_form_row_visible(ref_int_edit, is_piece and not is_stock_mp and not template_mode, "Ref. interna")
        set_form_row_visible(ref_ext_edit, is_piece and not is_stock_mp, "Ref. externa")
        set_form_row_visible(product_combo, is_product, "Produto")
        set_form_row_visible(product_unid_edit, is_product or line_type == backend.ORC_LINE_TYPE_SERVICE, "Codigo/unid")
        set_form_row_visible(material_combo, is_piece, "Material")
        set_form_row_visible(esp_combo, is_piece, "Espessura")
        set_form_row_visible(dimension_edit, is_stock_mp, "Dimensao")
        show_stock_commercial = is_stock_mp
        show_operations = is_piece and not is_stock_mp
        set_form_row_visible(operation_selector, show_operations, "Postos")
        set_form_row_visible(operation_cost_label, show_operations, "Custeio op.")
        operation_buttons_host.setVisible(show_operations)
        set_form_row_visible(tempo_spin, show_operations, "Tempo peca (min)")
        set_form_row_visible(drawing_edit, show_operations, "Desenho")
        set_form_row_visible(docs_host, show_operations, "PDFs associados")
        ref_buttons_host.setVisible(show_operations)
        set_form_row_visible(mp_metric_spin, show_stock_commercial, stock_price_state.get("metric_label", "Kg por unid."))
        set_form_row_visible(mp_price_kg_spin, show_stock_commercial, f"Preco compra ({stock_price_state.get('base_label', 'EUR/kg')})")
        set_form_row_visible(mp_margin_spin, show_stock_commercial, "Margem")
        if is_product:
            base_state["laser_base_enabled"] = False
            base_state["laser_base_time"] = 0.0
            base_state["laser_base_price"] = 0.0
            ref_int_edit.clear()
            ref_ext_edit.clear()
            material_combo.setCurrentText("")
            esp_combo.setCurrentText("")
            dimension_edit.clear()
            drawing_edit.clear()
            apply_operations("")
            tempo_spin.setValue(0.0)
            stock_line_meta["stock_material_id"] = ""
            stock_line_meta["material_family"] = ""
            stock_line_meta["material_subtype"] = ""
            sync_product_fields()
        elif is_stock_mp:
            base_state["laser_base_enabled"] = False
            base_state["laser_base_time"] = 0.0
            base_state["laser_base_price"] = 0.0
            ref_int_edit.clear()
            drawing_edit.clear()
            if entered_stock_mp:
                desc_edit.clear()
                ref_ext_edit.clear()
                dimension_edit.clear()
                material_combo.setCurrentText("")
                esp_combo.setCurrentText("")
                stock_line_meta["stock_material_id"] = ""
                stock_line_meta["material_family"] = ""
                stock_line_meta["material_subtype"] = ""
                stock_material_combo.setCurrentIndex(0)
                refresh_stock_material_combo("")
            apply_operations("")
            tempo_spin.setValue(0.0)
            sync_stock_material_fields()
        elif line_type == backend.ORC_LINE_TYPE_SERVICE:
            base_state["laser_base_enabled"] = False
            base_state["laser_base_time"] = 0.0
            base_state["laser_base_price"] = 0.0
            ref_int_edit.clear()
            ref_ext_edit.clear()
            material_combo.setCurrentText("")
            esp_combo.setCurrentText("")
            dimension_edit.clear()
            drawing_edit.clear()
            product_unid_edit.setText("SV")
            if not operation_edit.text().strip():
                apply_operations("Montagem")
        else:
            product_unid_edit.clear()
            base_state["laser_base_enabled"] = bool(base_state.get("laser_base_enabled", False) or _payload_has_laser_base(initial))
            if base_state["laser_base_enabled"] and base_state["laser_base_time"] <= 0 and base_state["laser_base_price"] <= 0:
                base_state["laser_base_time"] = round(float(tempo_spin.value() or 0), 4)
                base_state["laser_base_price"] = round(float(price_spin.value() or 0), 4)
            if not template_mode and not ref_int_edit.text().strip():
                generate_ref()
        if not is_stock_mp and line_type == backend.ORC_LINE_TYPE_PIECE and str(initial.get("stock_material_id", "") or "").strip():
            stock_line_meta["stock_material_id"] = str(initial.get("stock_material_id", "") or "").strip()
        sync_operation_meta_selection()
        sync_line_totals_from_meta()
        refresh_operation_cost_hint()
        refresh_mp_price_panel()
        mode_state["initialized"] = True
        mode_state["last_token"] = token

    type_combo.currentTextChanged.connect(lambda _value: sync_mode())
    stock_kind_combo.currentTextChanged.connect(lambda _value: (refresh_stock_material_combo(), sync_stock_material_fields()))
    stock_material_combo.currentTextChanged.connect(lambda _value: sync_stock_material_fields())
    mp_price_kg_spin.valueChanged.connect(lambda _value: sync_line_price_from_mp())
    mp_margin_spin.valueChanged.connect(lambda _value: sync_line_price_from_mp())
    price_spin.valueChanged.connect(lambda _value: sync_margin_from_line_price())
    mp_price_kg_spin.lineEdit().textEdited.connect(lambda _text: stock_price_state.__setitem__("price_user_touched", True))
    mp_price_table_btn.clicked.connect(
        lambda: (
            partial(edit_material_prices, owner, backend)(
                str(stock_kind_combo.currentData() or "").strip(),
                str(stock_line_meta.get("stock_material_id", "") or "").strip(),
                parent=dialog,
            ),
            reload_stock_material_rows(),
            refresh_stock_material_combo(str(stock_line_meta.get("stock_material_id", "") or "").strip()),
            refresh_mp_price_panel(),
        )
    )
    stock_create_btn.clicked.connect(create_stock_material)
    sync_mode()

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    def accept_line_dialog() -> None:
        try:
            if current_type_token() == "stock_mp":
                record, preview = stock_material_detail()
                if not record and not material_combo.currentText().strip():
                    QMessageBox.warning(dialog, "Matéria Prima - Stock", "Seleciona stock de matéria-prima ou indica o material manualmente.")
                    return
                if not desc_edit.text().strip():
                    if record:
                        sync_stock_material_fields()
                    else:
                        desc_edit.setText(" ".join(part for part in (material_combo.currentText().strip(), dimension_edit.text().strip()) if part).strip())
                if not material_combo.currentText().strip():
                    material_combo.setCurrentText(str(record.get("material", "") or "").strip())
                if record:
                    esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                    if esp_txt and not esp_combo.currentText().strip():
                        esp_combo.setCurrentText(esp_txt)
            dialog.accept()
        except Exception as exc:
            QMessageBox.critical(dialog, "Linha de orcamento", str(exc))
    buttons.accepted.connect(accept_line_dialog)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return None
    sync_operation_meta_selection()
    line_type = current_line_type()
    token = current_type_token()
    stock_row = None
    if str(stock_line_meta.get("stock_material_id", "") or "").strip():
        stock_row = next(
            (
                row
                for row in list(backend.material_price_rows() or [])
                if str(row.get("id", "") or "").strip() == str(stock_line_meta.get("stock_material_id", "") or "").strip()
            ),
            None,
        )
    stock_update = None
    stock_pricing_label = str(stock_price_state.get("base_label", "EUR/kg") or "EUR/kg").strip() or "EUR/kg"
    stock_price_base_value = float(mp_price_kg_spin.value() or 0.0)
    stock_price_kg_value = stock_price_base_value
    if stock_pricing_label == "EUR/m":
        kg_m_value = float((stock_row or {}).get("kg_m", initial.get("kg_per_m", 0)) or 0.0)
        stock_price_kg_value = round((stock_price_base_value / kg_m_value), 4) if kg_m_value > 0 else 0.0
    if line_type == backend.ORC_LINE_TYPE_PIECE and str(stock_line_meta.get("stock_material_id", "") or "").strip():
        if isinstance(stock_row, dict):
            stock_update = partial(sync_stock_price, owner, backend)(
                str(stock_line_meta.get("stock_material_id", "") or "").strip(),
                float(stock_price_kg_value or 0.0),
                float(stock_row.get("price_kg", 0) or 0.0),
                parent=dialog,
            )
    stock_update_payload = dict(stock_update or {}) if isinstance(stock_update, dict) else {}
    final_time_unit, final_price_unit = current_composed_totals()
    if float(stock_update_payload.get("price_kg", 0) or 0.0) > 0:
        if stock_pricing_label == "EUR/m":
            base_unit_value = round(float(stock_update_payload.get("price_kg", 0) or 0.0) * float((stock_row or {}).get("kg_m", initial.get("kg_per_m", 0)) or 0.0), 4)
        else:
            base_unit_value = float(stock_update_payload.get("price_kg", 0) or 0.0)
        final_price_unit = round(
            float(mp_metric_spin.value() or current_piece_metric_kg() or 0.0)
            * base_unit_value
            * (1.0 + (float(mp_margin_spin.value() or 0.0) / 100.0)),
            4,
        )
    elif token == "stock_mp":
        final_time_unit = 0.0
        final_price_unit = round(float(price_spin.value() or 0.0), 4)
    is_raw_stock_line = token == "stock_mp"
    is_product_line = line_type == backend.ORC_LINE_TYPE_PRODUCT
    route_operations = full_route_operation_names()
    raw_has_technical_route = bool(
        is_raw_stock_line
        and (
            route_operations
            or list(operation_meta.get("operacoes_detalhe", []) or [])
            or dict(operation_meta.get("tempos_operacao", {}) or {})
            or dict(operation_meta.get("custos_operacao", {}) or {})
        )
    )
    allow_technical_docs = bool(
        line_type == backend.ORC_LINE_TYPE_PIECE
        and (not is_raw_stock_line or raw_has_technical_route)
    )
    final_ref_externa = ""
    if is_raw_stock_line:
        record, _preview = stock_material_detail()
        final_ref_externa = str((record or {}).get("id", "") or stock_line_meta.get("stock_material_id", "") or "").strip()
    elif line_type == backend.ORC_LINE_TYPE_PIECE:
        final_ref_externa = ref_ext_edit.text().strip()
    final_operation = ""
    if not is_product_line and (not is_raw_stock_line or raw_has_technical_route):
        final_operation = " + ".join(route_operations)
    keep_operation_meta = not is_product_line and (not is_raw_stock_line or raw_has_technical_route)
    return {
        "tipo_item": line_type,
        "stock_item_kind": "raw_material" if is_raw_stock_line else ("product" if is_product_line else str(initial.get("stock_item_kind", "") or "").strip()),
        "ref_interna": "" if (template_mode or line_type != backend.ORC_LINE_TYPE_PIECE or is_raw_stock_line) else ref_int_edit.text().strip(),
        "ref_externa": final_ref_externa,
        "descricao": desc_edit.text().strip(),
        "dimensao": dimension_edit.text().strip(),
        "dimensoes": dimension_edit.text().strip(),
        "material": material_combo.currentText().strip() if line_type == backend.ORC_LINE_TYPE_PIECE else "",
        "material_family": str(stock_line_meta.get("material_family", "") or initial.get("material_family", "") or (material_combo.currentText().strip() if line_type == backend.ORC_LINE_TYPE_PIECE else "")).strip() if line_type == backend.ORC_LINE_TYPE_PIECE else "",
        "material_subtype": str(
            stock_line_meta.get("material_subtype", "")
            or initial.get("material_subtype", "")
            or (str(stock_kind_combo.currentData() or "").strip() if token == "stock_mp" else "")
        ).strip() if line_type == backend.ORC_LINE_TYPE_PIECE else "",
        "material_supplied_by_client": bool(initial.get("material_supplied_by_client", False) or initial.get("material_fornecido_cliente", False)) if line_type == backend.ORC_LINE_TYPE_PIECE else False,
        "material_fornecido_cliente": bool(initial.get("material_fornecido_cliente", False) or initial.get("material_supplied_by_client", False)) if line_type == backend.ORC_LINE_TYPE_PIECE else False,
        "material_cost_included": (
            (
                bool(initial.get("material_cost_included", True))
                if "material_cost_included" in initial
                else not bool(initial.get("material_supplied_by_client", False) or initial.get("material_fornecido_cliente", False))
            )
            if line_type == backend.ORC_LINE_TYPE_PIECE
            else False
        ),
        "espessura": esp_combo.currentText().strip() if line_type == backend.ORC_LINE_TYPE_PIECE else "",
        "operacao": final_operation,
        "tempo_peca_min": final_time_unit,
        "qtd": qtd_spin.value(),
        "qtd_base": float(qtd_spin.value() if template_mode else initial.get("qtd_base", qtd_spin.value())),
        "preco_unit": final_price_unit,
        "desenho": drawing_edit.text().strip() if allow_technical_docs else "",
        "desenho_pdf": pdf_docs[0] if allow_technical_docs and pdf_docs else "",
        "desenhos_pdf": list(pdf_docs) if allow_technical_docs else [],
        "ficheiros": [
            item
            for item in [drawing_edit.text().strip(), *pdf_docs]
            if item
        ] if allow_technical_docs else [],
        "stock_material_id": str(stock_line_meta.get("stock_material_id", "") or "").strip() if (token == "stock_mp" and stock_material_detail()[0]) or str(initial.get("stock_material_id", "") or "").strip() else "",
        "price_per_kg": float(stock_update_payload.get("price_kg", stock_price_kg_value) or 0.0) if line_type == backend.ORC_LINE_TYPE_PIECE else 0.0,
        "price_base_value": float(stock_price_base_value or 0.0) if token == "stock_mp" else float(initial.get("price_base_value", 0) or 0.0),
        "price_base_label": stock_pricing_label if token == "stock_mp" else str(initial.get("price_base_label", "") or "").strip(),
        "price_markup_pct": float(mp_margin_spin.value() or 0.0) if token == "stock_mp" else float(initial.get("price_markup_pct", 0) or 0.0),
        "stock_metric_value": float(mp_metric_spin.value() or 0.0) if token == "stock_mp" else float(initial.get("stock_metric_value", 0) or 0.0),
        "kg_per_m": float(initial.get("kg_per_m", stock_row.get("kg_m", 0) if isinstance(stock_row, dict) else 0) or 0.0),
        "laser_base_active": bool(line_has_laser_base()),
        "laser_base_tempo_unit": current_laser_base_totals()[0] if line_type == backend.ORC_LINE_TYPE_PIECE else 0.0,
        "laser_base_preco_unit": current_laser_base_totals()[1] if line_type == backend.ORC_LINE_TYPE_PIECE else 0.0,
        "produto_codigo": current_product_code() if line_type == backend.ORC_LINE_TYPE_PRODUCT else "",
        "produto_unid": product_unid_edit.text().strip() if line_type != backend.ORC_LINE_TYPE_PIECE else "",
        "_product_pending_create": bool(line_type == backend.ORC_LINE_TYPE_PRODUCT and not current_product_code()),
        "conjunto_codigo": str(initial.get("conjunto_codigo", "") or "").strip(),
        "conjunto_nome": str(initial.get("conjunto_nome", "") or "").strip(),
        "grupo_uuid": str(initial.get("grupo_uuid", "") or "").strip(),
        "operacoes_lista": list(operation_meta.get("operacoes_lista", []) or []) if keep_operation_meta else [],
        "operacoes_fluxo": [dict(item or {}) for item in list(operation_meta.get("operacoes_fluxo", []) or []) if isinstance(item, dict)] if keep_operation_meta else [],
        "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)] if keep_operation_meta else [],
        "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}) if keep_operation_meta else {},
        "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}) if keep_operation_meta else {},
        "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}) if keep_operation_meta else {},
    }
