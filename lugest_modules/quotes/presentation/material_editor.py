from __future__ import annotations
import re
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QFrame, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QSizePolicy, QStackedWidget, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_support import _build_operation_selector, _fmt_eur, _open_operation_cost_profiles_dialog, _open_quote_operation_detail_dialog, _operation_tokens
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from typing import Any
from lugest_modules.quotes.presentation.material_prices import edit_material_prices, sync_stock_price
def edit_material(owner: QWidget, backend: QuoteEditorPorts, initial: dict | None = None, parent: QWidget | None = None, *, presets: dict) -> dict | None:
    initial = dict(initial or {})
    def _initial_float(value: Any, default: float = 0.0) -> float:
        try:
            return float(str(value).strip().replace(" ", "").replace(",", "."))
        except Exception:
            return float(default)

    if isinstance(initial.get("line"), dict):
        wrapper = dict(initial)
        line_payload = dict(wrapper.get("line") or {})
        merged = dict(line_payload)
        for key, value in wrapper.items():
            if key == "line":
                continue
            if merged.get(key) in (None, "", [], {}):
                merged[key] = value
        initial = merged

    allowed_modes = {"Perfil", "Tubo", "Chapa", "Cantoneira", "Barra", "Ferro nervurado", "Manual"}
    initial_desc_text = str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip()
    if not str(initial.get("descricao_base", "") or "").strip() and " | " in initial_desc_text:
        initial["descricao_base"] = initial_desc_text.split(" | ", 1)[0].strip()
    initial_mode = str(initial.get("calc_mode", "") or initial.get("material_subtype", "") or "").strip()
    initial_probe = backend.norm_text(
        " ".join(
            str(initial.get(key, "") or "")
            for key in ("descricao_base", "descricao", "material", "material_family", "material_subtype", "calc_mode")
        )
    )
    if initial_mode not in allowed_modes:
        initial_mode = ""
    if not initial_mode or (initial_mode == "Perfil" and any(token in initial_probe for token in ("barra", "chata", "plat", "flat"))):
        if any(token in initial_probe for token in ("barra", "chata", "plat", "flat")):
            initial_mode = "Barra"
        elif "cantoneira" in initial_probe or re.search(r"\bl\s*\d+\s*x\s*\d+", initial_probe):
            initial_mode = "Cantoneira"
        elif "tubo" in initial_probe:
            initial_mode = "Tubo"
        elif "chapa" in initial_probe:
            initial_mode = "Chapa"
        elif "ferro nervurado" in initial_probe or "varao nervurado" in initial_probe:
            initial_mode = "Ferro nervurado"
    initial["calc_mode"] = initial_mode or "Perfil"

    desc_for_metrics = str(initial.get("descricao", initial_desc_text) or "")
    if "quantity_units" not in initial and initial.get("qtd") not in (None, ""):
        initial["quantity_units"] = initial.get("qtd")
    if _initial_float(initial.get("meters_per_unit", 0), 0.0) <= 0:
        match_meters = re.search(r"x\s*(\d+(?:[.,]\d+)?)\s*m\b", desc_for_metrics, flags=re.IGNORECASE)
        if match_meters:
            initial["meters_per_unit"] = float(match_meters.group(1).replace(",", "."))
    if _initial_float(initial.get("kg_per_m", 0), 0.0) <= 0:
        match_kg_m = re.search(r"(\d+(?:[.,]\d+)?)\s*kg\s*/\s*m", desc_for_metrics, flags=re.IGNORECASE)
        if match_kg_m:
            initial["kg_per_m"] = float(match_kg_m.group(1).replace(",", "."))
    if _initial_float(initial.get("price_per_kg", 0), 0.0) <= 0:
        match_price_kg = re.search(r"(\d+(?:[.,]\d+)?)\s*eur\s*/\s*kg", desc_for_metrics, flags=re.IGNORECASE)
        if match_price_kg:
            initial["price_per_kg"] = float(match_price_kg.group(1).replace(",", "."))
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Item material")
    dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)
    dialog.setSizeGripEnabled(True)
    try:
        screen = QApplication.screenAt(dialog.parentWidget().mapToGlobal(dialog.parentWidget().rect().center())) if dialog.parentWidget() is not None else QApplication.primaryScreen()
        available = screen.availableGeometry() if screen is not None else None
        if available is not None:
            dialog.resize(min(980, max(820, available.width() - 100)), min(760, max(560, available.height() - 100)))
            dialog.setMaximumHeight(max(520, available.height() - 60))
        else:
            dialog.resize(940, 720)
    except Exception:
        dialog.resize(940, 720)
    dialog.setMinimumSize(760, 520)
    root_layout = QVBoxLayout(dialog)
    root_layout.setContentsMargins(12, 12, 12, 12)
    root_layout.setSpacing(8)
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    body = QWidget()
    layout = QVBoxLayout(body)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(8)
    scroll.setWidget(body)
    root_layout.addWidget(scroll, 1)

    top_form = QFormLayout()
    top_form.setContentsMargins(0, 0, 0, 0)
    top_form.setHorizontalSpacing(10)
    top_form.setVerticalSpacing(8)

    mode_combo = QComboBox()
    mode_combo.addItems(["Perfil", "Tubo", "Chapa", "Cantoneira", "Barra", "Ferro nervurado", "Manual"])
    mode_combo.setCurrentText(str(initial.get("calc_mode", "Perfil") or "Perfil"))
    desc_edit = QLineEdit(str(initial.get("descricao_base", initial.get("descricao", "")) or "").strip())
    ref_edit = QLineEdit(str(initial.get("ref_externa", "") or "").strip())
    qty_spin = QDoubleSpinBox()
    qty_spin.setRange(0.01, 1000000.0)
    qty_spin.setDecimals(2)
    qty_spin.setValue(float(initial.get("quantity_units", initial.get("qtd", 1)) or 1))
    top_form.addRow("Modo", mode_combo)
    top_form.addRow("Descricao", desc_edit)
    top_form.addRow("Codigo/Ref.", ref_edit)
    top_form.addRow("Quantidade", qty_spin)
    layout.addLayout(top_form)

    detail_intro = QLabel(
        "Define o material do conjunto com origem em catálogo, stock de matéria-prima ou criação manual. "
        "A descrição e a referência passam a acompanhar a dimensão, a qualidade e o formato escolhido."
    )
    detail_intro.setWordWrap(True)
    detail_intro.setProperty("role", "muted")
    layout.addWidget(detail_intro)

    stack = QStackedWidget()
    stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
    stack.setMaximumHeight(360)
    layout.addWidget(stack)

    summary_card = CardFrame()
    summary_card.set_tone("info")
    summary_layout = QVBoxLayout(summary_card)
    summary_layout.setContentsMargins(10, 8, 10, 8)
    summary_layout.setSpacing(4)
    summary_weight = QLabel("")
    summary_cost = QLabel("")
    summary_desc = QLabel("")
    summary_desc.setWordWrap(True)
    summary_layout.addWidget(summary_weight)
    summary_layout.addWidget(summary_cost)
    summary_layout.addWidget(summary_desc)
    layout.addWidget(summary_card)

    tool_row = QHBoxLayout()
    price_table_btn = QPushButton("Tabela precos MP")
    price_table_btn.setProperty("variant", "secondary")
    tool_row.addWidget(price_table_btn)
    tool_row.addStretch(1)
    layout.addLayout(tool_row)

    operation_meta = {
        "operacoes_lista": list(initial.get("operacoes_lista", []) or []),
        "operacoes_fluxo": [dict(item or {}) for item in list(initial.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
        "operacoes_detalhe": [dict(item or {}) for item in list(initial.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
        "tempos_operacao": dict(initial.get("tempos_operacao", {}) or {}),
        "custos_operacao": dict(initial.get("custos_operacao", {}) or {}),
        "quote_cost_snapshot": dict(initial.get("quote_cost_snapshot", {}) or {}),
    }
    operation_selector, operation_edit, apply_operations = _build_operation_selector(
        list(presets.get("operacoes", []) or []),
        " + ".join(list(operation_meta.get("operacoes_lista", []) or [])) or str(initial.get("operacao", "") or ""),
    )
    operation_cost_label = QLabel("")
    operation_cost_label.setWordWrap(True)
    operation_cost_label.setProperty("role", "muted")
    operation_buttons = QHBoxLayout()
    operation_buttons.setContentsMargins(0, 0, 0, 0)
    operation_buttons.setSpacing(8)
    op_detail_btn = QPushButton("Quantificar operações")
    op_detail_btn.setProperty("variant", "secondary")
    op_profiles_btn = QPushButton("Perfis operações")
    op_profiles_btn.setProperty("variant", "secondary")
    operation_buttons.addWidget(op_detail_btn)
    operation_buttons.addWidget(op_profiles_btn)
    operation_buttons.addStretch(1)
    operation_buttons_host = QWidget()
    operation_buttons_host.setLayout(operation_buttons)
    op_form = QFormLayout()
    op_form.setContentsMargins(0, 0, 0, 0)
    op_form.setHorizontalSpacing(10)
    op_form.setVerticalSpacing(6)
    op_form.addRow("Operações seguintes", operation_selector)
    op_form.addRow("Custeio op.", operation_cost_label)
    op_form.addRow("", operation_buttons_host)
    root_layout.addLayout(op_form)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    root_layout.addWidget(buttons)

    desc_state = {"manual": bool(desc_edit.text().strip()), "last_auto": ""}
    ref_state = {"manual": bool(ref_edit.text().strip()), "last_auto": ""}
    sync = {"busy": False}

    profile_options = [dict(row or {}) for row in list(backend.material_section_options("Perfil") or [])]
    tube_options = [dict(row or {}) for row in list(backend.material_section_options("Tubo") or [])]
    stock_rows = [dict(item.get("record") or {}) for item in list(backend.material_rows("") or []) if isinstance(item, dict) and isinstance(item.get("record"), dict)]
    family_options = [
        dict(row or {})
        for row in list(backend.material_family_options() or [])
        if str((row or {}).get("key", "") or "").strip()
    ]
    if not family_options:
        family_options = [
            {"key": "steel", "label": "Ferro / Aço", "density": 7.85},
            {"key": "stainless", "label": "Inox", "density": 7.93},
            {"key": "aluminum", "label": "Alumínio", "density": 2.7},
            {"key": "brass", "label": "Latão", "density": 8.5},
            {"key": "copper", "label": "Cobre", "density": 8.96},
        ]
    family_quality_presets = {
        "steel": ["S235JR", "S275JR", "S355JR", "DC01", "DD11"],
        "stainless": ["Inox 304L", "Inox 316L", "Inox 430"],
        "aluminum": ["Alumínio 5754", "Alumínio 5083", "Alumínio 6082", "Alumínio 1050"],
        "brass": ["Latão"],
        "copper": ["Cobre"],
    }

    equal_angle_presets: dict[str, list[str]] = {
        "30": ["3", "4"],
        "35": ["3", "4"],
        "40": ["3", "4", "5"],
        "45": ["4", "5"],
        "50": ["4", "5", "6"],
        "60": ["5", "6", "8"],
        "65": ["5", "6", "8"],
        "70": ["6", "7", "8"],
        "75": ["6", "8"],
        "80": ["6", "8", "10"],
        "90": ["8", "9", "10"],
        "100": ["8", "10", "12"],
        "120": ["10", "12"],
    }
    flat_bar_presets = [
        "20x3",
        "25x3",
        "25x5",
        "30x3",
        "30x5",
        "40x5",
        "40x6",
        "50x5",
        "50x6",
        "60x6",
        "60x8",
        "70x8",
        "70x10",
        "70x20",
        "80x8",
        "80x10",
        "100x8",
        "100x10",
        "100x12",
        "120x10",
        "120x12",
    ]
    initial_bar_size = str(initial.get("bar_size", "") or "").strip()
    if not initial_bar_size:
        desc_for_bar = str(initial.get("descricao_base", initial.get("descricao", "")) or "")
        match_bar = re.search(r"(\d+(?:[.,]\d+)?)\s*[xX]\s*(\d+(?:[.,]\d+)?)", desc_for_bar)
        if match_bar:
            initial_bar_size = f"{match_bar.group(1)}x{match_bar.group(2)}".replace(",", ".")
    if initial_bar_size and initial_bar_size not in flat_bar_presets:
        flat_bar_presets.append(initial_bar_size)
    tube_model_presets: dict[str, list[str]] = {
        "redondo": ["Ø33.7x2.6", "Ø42.4x3.2", "Ø48.3x3.2", "Ø60.3x3.2", "Ø76.1x3.2", "Ø88.9x4.0", "Ø114.3x4.0"],
        "quadrado": ["30x30x2.0", "40x40x2.0", "40x40x3.0", "50x50x2.5", "60x60x3.0", "80x80x4.0", "100x100x5.0"],
        "retangular": ["40x20x2.0", "50x30x2.5", "60x40x3.0", "80x40x3.0", "100x50x4.0", "120x60x4.0"],
    }
    sheet_model_presets = [
        "2000x1000x2",
        "Chapa gota 2000x1000x3/5",
        "Chapa gota 2000x1000x4/6",
        "2000x1000x5/7",
        "2000x1000x6/8",
        "2000x1000x8/10",
        "Chapa gota 2500x1250x3/5",
        "Chapa gota 2500x1250x4/6",
        "Chapa gota 2500x1250x5/7",
        "Chapa gota 2500x1250x6/8",
        "Chapa gota 2500x1250x8/10",
        "2500x1250x3",
        "2500x1250x5",
        "3000x1500x3",
        "Chapa gota 3000x1500x3/5",
        "Chapa gota 3000x1500x4/6",
        "Chapa gota 3000x1500x5/7",
        "Chapa gota 3000x1500x6/8",
        "Chapa gota 3000x1500x8/10",
        "Chapa gota 3000x1500x10/12",
        "3000x1500x5",
        "3000x1500x10",
        "3000x1500x15",
        "4000x2000x5",
        "4000x2000x8",
        "4000x2000x10",
    ]

    def _parse_size_value(value: object) -> tuple[str, float]:
        text = str(value or "").strip()
        if not text:
            return "", 0.0
        match = re.search(r"(\d+(?:[.,]\d+)?)", text)
        if not match:
            return text, 0.0
        try:
            return text, float(match.group(1).replace(",", "."))
        except Exception:
            return text, 0.0

    def _parse_pair(value: object) -> tuple[float, float]:
        text = str(value or "").lower().replace(",", ".").replace("mm", "").strip()
        match = re.search(r"(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)", text)
        if not match:
            return 0.0, 0.0
        try:
            return float(match.group(1)), float(match.group(2))
        except Exception:
            return 0.0, 0.0

    def _sheet_model_info(value: object) -> dict[str, Any]:
        raw = str(value or "").strip()
        text = raw.lower().replace(",", ".").replace("mm", "")
        is_checker = str(sheet_kind_combo.currentData() or "").strip() == "checker" or any(
            token in backend.norm_text(text) for token in ("gota", "xadrez", "antiderrap")
        )
        slash_match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
        numbers = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", text)]
        length = width = thickness = 0.0
        thickness_label = ""
        if len(numbers) >= 3:
            length = numbers[0]
            width = numbers[1]
            thickness = numbers[2]
        elif is_checker and slash_match:
            length = float(length_mm_spin.value() or 3000.0)
            width = float(width_mm_spin.value() or 1500.0)
            thickness = float(slash_match.group(1))
        if slash_match:
            left_raw = slash_match.group(1)
            right_raw = slash_match.group(2)
            left = left_raw.rstrip("0").rstrip(".") if "." in left_raw else left_raw
            right = right_raw.rstrip("0").rstrip(".") if "." in right_raw else right_raw
            thickness_label = f"{left}/{right}"
            thickness = float(slash_match.group(1))
            is_checker = True
        elif thickness > 0:
            thickness_label = _float_text(thickness, 2).replace(",00", "").replace(",0", "")
        return {
            "raw": raw,
            "length": length,
            "width": width,
            "thickness": thickness,
            "thickness_label": thickness_label,
            "is_checker": is_checker,
        }

    def _sort_dimension_labels(values: list[str]) -> list[str]:
        def _sort_key(value: str) -> tuple[float, float, float, str]:
            text = str(value or "").lower().replace("ø", "").replace("mm", "").replace(",", ".").strip()
            parts = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", text)]
            while len(parts) < 3:
                parts.append(0.0)
            return (parts[0], parts[1], parts[2], text)

        return sorted({str(value or "").strip() for value in values if str(value or "").strip()}, key=_sort_key)

    def _add_combo_item_if_missing(combo: QComboBox, value: object) -> None:
        text = str(value or "").strip()
        if not text:
            return
        for idx in range(combo.count()):
            if str(combo.itemText(idx) or "").strip().lower() == text.lower():
                return
        combo.addItem(text)

    def _set_combo_text_preserving_custom(combo: QComboBox, value: object) -> None:
        text = str(value or "").strip()
        if not text:
            return
        _add_combo_item_if_missing(combo, text)
        combo.setCurrentText(text)

    def _float_text(value: float, digits: int = 2) -> str:
        return f"{float(value or 0):.{digits}f}".replace(".", ",")

    def _set_edit_if_auto(edit: QLineEdit, state: dict, value: str) -> None:
        current = edit.text().strip()
        if state["manual"] and current and current != state["last_auto"]:
            return
        edit.blockSignals(True)
        edit.setText(value)
        edit.blockSignals(False)
        state["last_auto"] = value

    def _stock_price_kg(record: dict, preview: dict) -> float:
        base_value = float(record.get("p_compra", 0) or 0.0)
        if base_value <= 0:
            base_value = float(preview.get("p_compra", 0) or 0.0)
        if base_value <= 0:
            total_unit = float(preview.get("preco_unid", record.get("preco_unid", 0)) or 0.0)
            total_weight = float(preview.get("peso_unid", record.get("peso_unid", 0)) or 0.0)
            if total_unit > 0 and total_weight > 0:
                return round(total_unit / total_weight, 4)
            return 0.0
        if str(preview.get("base_label", "") or "").strip().upper() == "EUR/M":
            kg_m = float(preview.get("kg_m", 0) or 0.0)
            if kg_m > 0:
                return round(base_value / kg_m, 4)
        return round(base_value, 4)

    def _stock_price_base(record: dict, preview: dict) -> tuple[str, float]:
        base_label = str(preview.get("base_label", "") or "").strip()
        if not base_label:
            base_label = "EUR/m" if str(preview.get("formato", record.get("formato", "")) or "").strip() == "Tubo" else "EUR/kg"
        base_value = float(record.get("p_compra", 0) or preview.get("p_compra", 0) or 0.0)
        if base_value <= 0:
            kg_price = _stock_price_kg(record, preview)
            if base_label.strip().upper() == "EUR/M":
                kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
                return base_label, round(kg_price * kg_m, 4) if kg_m > 0 else 0.0
            return base_label, kg_price
        return base_label, round(base_value, 4)

    def _stock_mode(record: dict, preview: dict) -> str:
        formato_txt = str(preview.get("formato", record.get("formato", "")) or "").strip().title()
        secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
        material_txt = backend.norm_text(str(record.get("material", "") or ""))
        if formato_txt == "Cantoneira":
            return "Cantoneira"
        if formato_txt == "Barra":
            return "Barra"
        if formato_txt == "Perfil" and secao == "L":
            return "Cantoneira"
        if formato_txt == "Perfil":
            return "Perfil"
        if formato_txt == "Tubo":
            return "Tubo"
        if formato_txt == "Chapa":
            if any(token in material_txt for token in ("barra", "chata", "plat", "flat")):
                return "Barra"
            return "Chapa"
        return ""

    def _stock_label(record: dict, preview: dict) -> str:
        material_txt = str(record.get("material", "") or "").strip()
        dim_txt = str(preview.get("dimension_label", "") or "-").strip()
        esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
        lote_txt = str(record.get("id", record.get("lote_fornecedor", "")) or "").strip()
        qty_txt = backend._fmt(record.get("quantidade", 0))
        price_txt = _fmt_eur(_stock_price_kg(record, preview))
        parts = [part for part in (lote_txt, material_txt, dim_txt) if part and part != "-"]
        if esp_txt:
            parts.append(f"{esp_txt} mm")
        if qty_txt:
            parts.append(f"Qtd {qty_txt}")
        if price_txt:
            base_label, base_value = _stock_price_base(record, preview)
            parts.append(f"Base {_fmt_eur(base_value)}/{base_label.split('/')[-1]}")
        return " | ".join(parts) or material_txt or lote_txt or "Stock"

    stock_cache: dict[str, list[tuple[dict, dict]]] = {}

    def _stock_options(mode: str) -> list[tuple[dict, dict]]:
        key = str(mode or "").strip()
        if key in stock_cache:
            return list(stock_cache[key])
        rows: list[tuple[dict, dict]] = []
        for record in stock_rows:
            preview = dict(backend.material_price_preview(record) or {})
            if _stock_mode(record, preview) == key:
                rows.append((record, preview))
        stock_cache[key] = rows
        return list(rows)

    def _fill_stock_combo(combo: QComboBox, mode: str, preferred_id: str = "") -> None:
        current_id = preferred_id or str(combo.currentData() or {}).strip() if isinstance(combo.currentData(), str) else preferred_id
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)
        for record, preview in _stock_options(mode):
            combo.addItem(_stock_label(record, preview), dict(record))
        if preferred_id:
            for idx in range(combo.count()):
                payload = combo.itemData(idx)
                if isinstance(payload, dict) and str(payload.get("id", "") or "").strip() == preferred_id:
                    combo.setCurrentIndex(idx)
                    break
        combo.blockSignals(False)

    def _current_stock(combo: QComboBox) -> tuple[dict, dict]:
        payload = combo.currentData()
        if not isinstance(payload, dict):
            return {}, {}
        preview = dict(backend.material_price_preview(payload) or {})
        return dict(payload), preview

    def _set_combo_by_data(combo: QComboBox, wanted: str) -> None:
        wanted_txt = str(wanted or "").strip().lower()
        if not wanted_txt:
            return
        for idx in range(combo.count()):
            current_txt = str(combo.itemData(idx) or "").strip().lower()
            if current_txt == wanted_txt:
                combo.setCurrentIndex(idx)
                return

    def _profile_series_and_size(record: dict, preview: dict) -> tuple[str, str]:
        secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
        dimension = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
        if secao and dimension.upper().startswith(f"{secao} "):
            dimension = dimension[len(secao) :].strip()
        if not dimension:
            dimension = str(record.get("altura", "") or record.get("espessura", "") or "").strip()
        return secao, dimension

    def _geometry_weight_kg_m(area_mm2: float, density: float) -> float:
        return round((max(0.0, float(area_mm2 or 0.0)) * max(0.0, float(density or 0.0))) / 1000.0, 4)

    def _quality_from_record(record: dict) -> str:
        material_txt = str(record.get("material", "") or "").strip()
        familia_txt = str(record.get("material_familia", "") or "").strip()
        if material_txt:
            return material_txt
        if familia_txt:
            return familia_txt
        return "S235JR"

    def _stock_material_identity(mode: str, record: dict, preview: dict) -> dict[str, str]:
        mode_txt = str(mode or "").strip()
        quality = _quality_from_record(record)
        dimension = str(preview.get("dimension_label", "") or "").strip()
        secao = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
        if mode_txt == "Perfil":
            normalized_dimension = dimension
            if secao and dimension.upper().startswith(f"{secao} "):
                normalized_dimension = dimension[len(secao) :].strip()
            ref_txt = str(record.get("id", "") or "").strip() or f"{secao} {normalized_dimension}".strip()
            desc_txt = f"Perfil {secao} {normalized_dimension} {quality}".replace("  ", " ").strip()
            return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
        if mode_txt == "Tubo":
            ref_txt = str(record.get("id", "") or "").strip() or f"TUB-{dimension}".replace(" ", "")
            desc_txt = f"Tubo {quality} {dimension}".replace("  ", " ").strip()
            return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
        if mode_txt == "Chapa":
            esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
            size_txt = str(dimension or "").replace(" mm", "").strip()
            ref_txt = str(record.get("id", "") or "").strip() or f"CH-{size_txt}x{esp_txt}".replace(" ", "")
            desc_txt = f"Chapa {quality} {size_txt}x{esp_txt} mm".replace("  ", " ").strip()
            return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
        if mode_txt == "Cantoneira":
            ref_txt = str(record.get("id", "") or "").strip() or f"L {dimension}".strip()
            desc_txt = f"Cantoneira abas iguais {quality} {dimension}".replace("  ", " ").strip()
            return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
        if mode_txt == "Barra":
            ref_txt = str(record.get("id", "") or "").strip() or f"BAR {dimension}".strip()
            desc_txt = f"Barra chata {quality} {dimension}".replace("  ", " ").strip()
            return {"ref": ref_txt, "desc": desc_txt, "quality": quality}
        return {"ref": str(record.get("id", "") or "").strip(), "desc": quality, "quality": quality}

    stock_tube_dimensions: dict[str, list[str]] = {"redondo": [], "quadrado": [], "retangular": []}
    stock_sheet_dimensions: list[str] = []
    stock_bar_dimensions: list[str] = []
    for record in stock_rows:
        preview = dict(backend.material_price_preview(record) or {})
        mode_key = _stock_mode(record, preview)
        dimension = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
        if mode_key == "Tubo":
            tube_key = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().lower()
            if tube_key in stock_tube_dimensions and dimension:
                stock_tube_dimensions[tube_key].append(dimension)
        elif mode_key == "Chapa":
            size_txt = str(preview.get("dimension_label", "") or "").strip().replace(" mm", "")
            esp_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
            if size_txt and esp_txt:
                stock_sheet_dimensions.append(f"{size_txt}x{esp_txt}")
        elif mode_key == "Barra":
            if dimension:
                stock_bar_dimensions.append(dimension)
        elif mode_key == "Cantoneira":
            a_val, b_val = _parse_pair(dimension)
            if a_val > 0:
                leg_key = str(int(round(a_val)))
                thickness_txt = str(preview.get("espessura", record.get("espessura", "")) or "").strip()
                if leg_key and thickness_txt:
                    equal_angle_presets.setdefault(leg_key, [])
                    if thickness_txt not in equal_angle_presets[leg_key]:
                        equal_angle_presets[leg_key].append(thickness_txt)

    for key, values in list(stock_tube_dimensions.items()):
        tube_model_presets[key] = _sort_dimension_labels(list(tube_model_presets.get(key, [])) + list(values))
    sheet_model_presets = _sort_dimension_labels(sheet_model_presets + stock_sheet_dimensions)
    flat_bar_presets = _sort_dimension_labels(flat_bar_presets + stock_bar_dimensions)
    equal_angle_presets = {
        leg: _sort_dimension_labels(values)
        for leg, values in sorted(equal_angle_presets.items(), key=lambda item: float(item[0]))
    }

    mode_to_index = {
        "Perfil": 0,
        "Tubo": 1,
        "Chapa": 2,
        "Cantoneira": 3,
        "Barra": 4,
        "Ferro nervurado": 5,
        "Manual": 6,
    }

    profile_source_combo = QComboBox()
    profile_source_combo.addItems(["Catálogo", "Stock MP"])
    profile_stock_combo = QComboBox()
    profile_series_combo = QComboBox()
    for row in profile_options:
        profile_series_combo.addItem(str(row.get("label", row.get("key", "")) or "").strip(), str(row.get("key", "") or "").strip())
    profile_size_combo = QComboBox()
    profile_size_combo.setEditable(True)
    profile_meters_spin = QDoubleSpinBox()
    profile_meters_spin.setRange(0.0, 1000.0)
    profile_meters_spin.setDecimals(3)
    profile_meters_spin.setSuffix(" m")
    profile_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
    profile_kg_per_m_spin = QDoubleSpinBox()
    profile_kg_per_m_spin.setRange(0.0, 1000.0)
    profile_kg_per_m_spin.setDecimals(4)
    profile_kg_per_m_spin.setSuffix(" kg/m")
    profile_price_kg_spin = QDoubleSpinBox()
    profile_price_kg_spin.setRange(0.0, 1000000.0)
    profile_price_kg_spin.setDecimals(4)
    profile_price_kg_spin.setPrefix("EUR ")
    profile_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
    profile_hint = QLabel("")
    profile_hint.setWordWrap(True)
    profile_hint.setProperty("role", "muted")

    page_profile = QWidget()
    profile_form = QFormLayout(page_profile)
    profile_form.setContentsMargins(0, 0, 0, 0)
    profile_form.addRow("Origem", profile_source_combo)
    profile_form.addRow("Stock MP", profile_stock_combo)
    profile_form.addRow("Série", profile_series_combo)
    profile_form.addRow("Tamanho", profile_size_combo)
    profile_form.addRow("Metros / unidade", profile_meters_spin)
    profile_form.addRow("Kg / metro", profile_kg_per_m_spin)
    profile_form.addRow("Preço / kg", profile_price_kg_spin)
    profile_form.addRow("", profile_hint)
    stack.addWidget(page_profile)

    tube_source_combo = QComboBox()
    tube_source_combo.addItems(["Novo modelo", "Stock MP"])
    tube_stock_combo = QComboBox()
    tube_quality_edit = QLineEdit(str(initial.get("quality", initial.get("material", "S235JR")) or "S235JR").strip())
    tube_section_combo = QComboBox()
    for row in tube_options:
        tube_section_combo.addItem(str(row.get("label", row.get("key", "")) or "").strip(), str(row.get("key", "") or "").strip())
    tube_model_combo = QComboBox()
    tube_model_combo.setEditable(True)
    tube_side_a_spin = QDoubleSpinBox()
    tube_side_a_spin.setRange(0.0, 5000.0)
    tube_side_a_spin.setDecimals(1)
    tube_side_a_spin.setSuffix(" mm")
    tube_side_b_spin = QDoubleSpinBox()
    tube_side_b_spin.setRange(0.0, 5000.0)
    tube_side_b_spin.setDecimals(1)
    tube_side_b_spin.setSuffix(" mm")
    tube_diameter_spin = QDoubleSpinBox()
    tube_diameter_spin.setRange(0.0, 5000.0)
    tube_diameter_spin.setDecimals(1)
    tube_diameter_spin.setSuffix(" mm")
    tube_thickness_spin = QDoubleSpinBox()
    tube_thickness_spin.setRange(0.0, 200.0)
    tube_thickness_spin.setDecimals(2)
    tube_thickness_spin.setSuffix(" mm")
    tube_meters_spin = QDoubleSpinBox()
    tube_meters_spin.setRange(0.0, 1000.0)
    tube_meters_spin.setDecimals(3)
    tube_meters_spin.setSuffix(" m")
    tube_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
    tube_kg_per_m_spin = QDoubleSpinBox()
    tube_kg_per_m_spin.setRange(0.0, 1000.0)
    tube_kg_per_m_spin.setDecimals(4)
    tube_kg_per_m_spin.setSuffix(" kg/m")
    tube_price_kg_spin = QDoubleSpinBox()
    tube_price_kg_spin.setRange(0.0, 1000000.0)
    tube_price_kg_spin.setDecimals(4)
    tube_price_kg_spin.setPrefix("EUR ")
    tube_price_kg_spin.setValue(float(initial.get("price_base_value", initial.get("price_per_kg", 3.15)) or 3.15))
    tube_hint = QLabel("")
    tube_hint.setWordWrap(True)
    tube_hint.setProperty("role", "muted")

    page_tube = QWidget()
    tube_form = QFormLayout(page_tube)
    tube_form.setContentsMargins(0, 0, 0, 0)
    tube_form.addRow("Origem", tube_source_combo)
    tube_form.addRow("Stock MP", tube_stock_combo)
    tube_form.addRow("Qualidade", tube_quality_edit)
    tube_form.addRow("Tipo", tube_section_combo)
    tube_form.addRow("Modelo / histórico", tube_model_combo)
    tube_form.addRow("Lado A / lado", tube_side_a_spin)
    tube_form.addRow("Lado B", tube_side_b_spin)
    tube_form.addRow("Diâmetro", tube_diameter_spin)
    tube_form.addRow("Espessura", tube_thickness_spin)
    tube_form.addRow("Metros / unidade", tube_meters_spin)
    tube_form.addRow("Kg / metro", tube_kg_per_m_spin)
    tube_form.addRow("Preço / m", tube_price_kg_spin)
    tube_form.addRow("", tube_hint)
    stack.addWidget(page_tube)

    sheet_source_combo = QComboBox()
    sheet_source_combo.addItems(["Novo modelo", "Stock MP"])
    sheet_stock_combo = QComboBox()
    sheet_family_combo = QComboBox()
    for option in family_options:
        sheet_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
    initial_quality_text = str(initial.get("quality", initial.get("material", "S235JR")) or "S235JR").strip()
    initial_family_key = str(initial.get("material_family_key", "") or initial.get("material_familia", "") or "").strip()
    if not initial_family_key:
        initial_family_key = str(backend.material_family_profile(initial_quality_text, "").get("key", "steel") or "steel").strip()
    _set_combo_by_data(sheet_family_combo, initial_family_key)
    sheet_quality_combo = QComboBox()
    sheet_quality_combo.setEditable(True)
    sheet_kind_combo = QComboBox()
    sheet_kind_combo.addItem("Chapa lisa", "plain")
    sheet_kind_combo.addItem("Chapa gota / xadrez", "checker")
    sheet_model_combo = QComboBox()
    sheet_model_combo.setEditable(True)
    length_mm_spin = QDoubleSpinBox()
    length_mm_spin.setRange(0.0, 100000.0)
    length_mm_spin.setDecimals(1)
    length_mm_spin.setSuffix(" mm")
    length_mm_spin.setValue(float(initial.get("length_mm", 3000) or 3000))
    width_mm_spin = QDoubleSpinBox()
    width_mm_spin.setRange(0.0, 100000.0)
    width_mm_spin.setDecimals(1)
    width_mm_spin.setSuffix(" mm")
    width_mm_spin.setValue(float(initial.get("width_mm", 1500) or 1500))
    thickness_mm_spin = QDoubleSpinBox()
    thickness_mm_spin.setRange(0.0, 1000.0)
    thickness_mm_spin.setDecimals(2)
    thickness_mm_spin.setSuffix(" mm")
    thickness_mm_spin.setValue(float(initial.get("thickness_mm", 15) or 15))
    density_spin = QDoubleSpinBox()
    density_spin.setRange(0.0, 50000.0)
    density_spin.setDecimals(1)
    density_spin.setSuffix(" kg/m3")
    density_spin.setValue(float(initial.get("density", 7850) or 7850))
    sheet_price_kg_spin = QDoubleSpinBox()
    sheet_price_kg_spin.setRange(0.0, 1000000.0)
    sheet_price_kg_spin.setDecimals(4)
    sheet_price_kg_spin.setPrefix("EUR ")
    sheet_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
    sheet_hint = QLabel("")
    sheet_hint.setWordWrap(True)
    sheet_hint.setProperty("role", "muted")
    initial_sheet_probe = backend.norm_text(
        " ".join(
            str(initial.get(key, "") or "")
            for key in ("sheet_kind", "sheet_model", "descricao_base", "descricao", "material")
        )
    )
    if any(token in initial_sheet_probe for token in ("gota", "xadrez", "antiderrap")):
        sheet_kind_combo.setCurrentIndex(1)

    page_sheet = QWidget()
    sheet_form = QFormLayout(page_sheet)
    sheet_form.setContentsMargins(0, 0, 0, 0)
    sheet_form.addRow("Origem", sheet_source_combo)
    sheet_form.addRow("Stock MP", sheet_stock_combo)
    sheet_form.addRow("Tipo material", sheet_family_combo)
    sheet_form.addRow("Qualidade", sheet_quality_combo)
    sheet_form.addRow("Tipo de chapa", sheet_kind_combo)
    sheet_form.addRow("Modelo / dimensão", sheet_model_combo)
    sheet_form.addRow("Comprimento", length_mm_spin)
    sheet_form.addRow("Largura", width_mm_spin)
    sheet_form.addRow("Espessura", thickness_mm_spin)
    sheet_form.addRow("Densidade", density_spin)
    sheet_form.addRow("Preço / kg", sheet_price_kg_spin)
    sheet_form.addRow("", sheet_hint)
    stack.addWidget(page_sheet)

    def _sheet_family_key() -> str:
        return str(sheet_family_combo.currentData() or "").strip() or "steel"

    def _sheet_family_density_kg_m3() -> float:
        profile = dict(backend.material_family_profile("", _sheet_family_key()) or {})
        density = float(profile.get("density", 7.85) or 7.85)
        return round(density * 1000.0, 1)

    def _refresh_sheet_quality_options(*, preserve_current: bool = True) -> None:
        current = sheet_quality_combo.currentText().strip() if preserve_current else ""
        if preserve_current and not current:
            current = initial_quality_text
        key = _sheet_family_key()
        values = list(family_quality_presets.get(key, []) or [])
        if current and current not in values and preserve_current:
            values.insert(0, current)
        if not current or current not in values:
            current = values[0] if values else current
        sheet_quality_combo.blockSignals(True)
        sheet_quality_combo.clear()
        for value in values:
            sheet_quality_combo.addItem(value)
        sheet_quality_combo.setCurrentText(current or (values[0] if values else ""))
        sheet_quality_combo.blockSignals(False)
        density_spin.blockSignals(True)
        density_spin.setValue(_sheet_family_density_kg_m3())
        density_spin.blockSignals(False)

    _refresh_sheet_quality_options()

    angle_source_combo = QComboBox()
    angle_source_combo.addItems(["Catálogo / histórico", "Stock MP"])
    angle_stock_combo = QComboBox()
    angle_family_combo = QComboBox()
    for option in family_options:
        angle_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
    _set_combo_by_data(angle_family_combo, initial_family_key)
    angle_quality_combo = QComboBox()
    angle_quality_combo.setEditable(True)
    angle_leg_combo = QComboBox()
    angle_leg_combo.addItems(list(equal_angle_presets.keys()))
    angle_leg_combo.setCurrentText(str(initial.get("angle_leg", "60") or "60"))
    angle_thickness_combo = QComboBox()
    angle_meters_spin = QDoubleSpinBox()
    angle_meters_spin.setRange(0.0, 1000.0)
    angle_meters_spin.setDecimals(3)
    angle_meters_spin.setSuffix(" m")
    angle_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
    angle_kg_per_m_spin = QDoubleSpinBox()
    angle_kg_per_m_spin.setRange(0.0, 1000.0)
    angle_kg_per_m_spin.setDecimals(4)
    angle_kg_per_m_spin.setSuffix(" kg/m")
    angle_price_kg_spin = QDoubleSpinBox()
    angle_price_kg_spin.setRange(0.0, 1000000.0)
    angle_price_kg_spin.setDecimals(4)
    angle_price_kg_spin.setPrefix("EUR ")
    angle_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
    angle_hint = QLabel("")
    angle_hint.setWordWrap(True)
    angle_hint.setProperty("role", "muted")

    page_angle = QWidget()
    angle_form = QFormLayout(page_angle)
    angle_form.setContentsMargins(0, 0, 0, 0)
    angle_form.addRow("Origem", angle_source_combo)
    angle_form.addRow("Stock MP", angle_stock_combo)
    angle_form.addRow("Tipo material", angle_family_combo)
    angle_form.addRow("Qualidade", angle_quality_combo)
    angle_form.addRow("Abas iguais", angle_leg_combo)
    angle_form.addRow("Espessura", angle_thickness_combo)
    angle_form.addRow("Metros / unidade", angle_meters_spin)
    angle_form.addRow("Kg / metro", angle_kg_per_m_spin)
    angle_form.addRow("Preço / kg", angle_price_kg_spin)
    angle_form.addRow("", angle_hint)
    stack.addWidget(page_angle)

    bar_source_combo = QComboBox()
    bar_source_combo.addItems(["Modelo / histórico", "Stock MP"])
    bar_stock_combo = QComboBox()
    bar_family_combo = QComboBox()
    for option in family_options:
        bar_family_combo.addItem(str(option.get("label", "") or "").strip(), str(option.get("key", "") or "").strip())
    _set_combo_by_data(bar_family_combo, initial_family_key)
    bar_quality_combo = QComboBox()
    bar_quality_combo.setEditable(True)
    bar_size_combo = QComboBox()
    bar_size_combo.setEditable(True)
    for value in flat_bar_presets:
        bar_size_combo.addItem(value)
    bar_size_combo.setCurrentText(initial_bar_size or "70x20")
    bar_meters_spin = QDoubleSpinBox()
    bar_meters_spin.setRange(0.0, 1000.0)
    bar_meters_spin.setDecimals(3)
    bar_meters_spin.setSuffix(" m")
    bar_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
    bar_kg_per_m_spin = QDoubleSpinBox()
    bar_kg_per_m_spin.setRange(0.0, 1000.0)
    bar_kg_per_m_spin.setDecimals(4)
    bar_kg_per_m_spin.setSuffix(" kg/m")
    bar_price_kg_spin = QDoubleSpinBox()
    bar_price_kg_spin.setRange(0.0, 1000000.0)
    bar_price_kg_spin.setDecimals(4)
    bar_price_kg_spin.setPrefix("EUR ")
    bar_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))
    bar_hint = QLabel("")
    bar_hint.setWordWrap(True)
    bar_hint.setProperty("role", "muted")

    page_bar = QWidget()
    bar_form = QFormLayout(page_bar)
    bar_form.setContentsMargins(0, 0, 0, 0)
    bar_form.addRow("Origem", bar_source_combo)
    bar_form.addRow("Stock MP", bar_stock_combo)
    bar_form.addRow("Tipo material", bar_family_combo)
    bar_form.addRow("Qualidade", bar_quality_combo)
    bar_form.addRow("Barra", bar_size_combo)
    bar_form.addRow("Metros / unidade", bar_meters_spin)
    bar_form.addRow("Kg / metro", bar_kg_per_m_spin)
    bar_form.addRow("Preço / kg", bar_price_kg_spin)
    bar_form.addRow("", bar_hint)
    stack.addWidget(page_bar)

    def _family_key(combo: QComboBox) -> str:
        return str(combo.currentData() or "").strip() or "steel"

    def _family_density_g_cm3(combo: QComboBox) -> float:
        profile = dict(backend.material_family_profile("", _family_key(combo)) or {})
        return float(profile.get("density", 7.85) or 7.85)

    def _refresh_quality_combo(family_combo: QComboBox, quality_combo: QComboBox, *, preserve_current: bool = True) -> None:
        current = quality_combo.currentText().strip() if preserve_current else ""
        if preserve_current and not current:
            current = initial_quality_text
        values = list(family_quality_presets.get(_family_key(family_combo), []) or [])
        if current and current not in values and preserve_current:
            values.insert(0, current)
        if not current or current not in values:
            current = values[0] if values else current
        quality_combo.blockSignals(True)
        quality_combo.clear()
        for value in values:
            quality_combo.addItem(value)
        quality_combo.setCurrentText(current or (values[0] if values else ""))
        quality_combo.blockSignals(False)

    _refresh_quality_combo(angle_family_combo, angle_quality_combo)
    _refresh_quality_combo(bar_family_combo, bar_quality_combo)

    rebar_meters_spin = QDoubleSpinBox()
    rebar_meters_spin.setRange(0.0, 1000.0)
    rebar_meters_spin.setDecimals(3)
    rebar_meters_spin.setSuffix(" m")
    rebar_meters_spin.setValue(float(initial.get("meters_per_unit", 6.0) or 6.0))
    diameter_mm_spin = QDoubleSpinBox()
    diameter_mm_spin.setRange(0.0, 500.0)
    diameter_mm_spin.setDecimals(1)
    diameter_mm_spin.setSuffix(" mm")
    diameter_mm_spin.setValue(float(initial.get("diameter_mm", 12) or 12))
    rebar_price_kg_spin = QDoubleSpinBox()
    rebar_price_kg_spin.setRange(0.0, 1000000.0)
    rebar_price_kg_spin.setDecimals(4)
    rebar_price_kg_spin.setPrefix("EUR ")
    rebar_price_kg_spin.setValue(float(initial.get("price_per_kg", 3.15) or 3.15))

    page_rebar = QWidget()
    rebar_form = QFormLayout(page_rebar)
    rebar_form.setContentsMargins(0, 0, 0, 0)
    rebar_form.addRow("Metros / unidade", rebar_meters_spin)
    rebar_form.addRow("Diâmetro", diameter_mm_spin)
    rebar_form.addRow("Preço / kg", rebar_price_kg_spin)
    stack.addWidget(page_rebar)

    unit_combo = QComboBox()
    unit_combo.addItems(["kg", "m", "un"])
    unit_combo.setCurrentText(str(initial.get("produto_unid", "un") or "un"))
    manual_unit_price_spin = QDoubleSpinBox()
    manual_unit_price_spin.setRange(0.0, 1000000.0)
    manual_unit_price_spin.setDecimals(4)
    manual_unit_price_spin.setPrefix("EUR ")
    manual_unit_price_spin.setValue(float(initial.get("manual_unit_price", initial.get("preco_unit", 0)) or 0))

    page_manual = QWidget()
    manual_form = QFormLayout(page_manual)
    manual_form.setContentsMargins(0, 0, 0, 0)
    manual_form.addRow("Unidade", unit_combo)
    manual_form.addRow("Preço unitário", manual_unit_price_spin)
    stack.addWidget(page_manual)

    def _refresh_profile_size_options() -> None:
        current_key = str(profile_series_combo.currentData() or "").strip()
        wanted = profile_size_combo.currentText().strip() or str(initial.get("profile_size", initial.get("perfil_tamanho", "")) or "").strip()
        options = [str(value or "").strip() for value in list(backend.material_profile_size_options(current_key) or []) if str(value or "").strip()]
        profile_size_combo.blockSignals(True)
        profile_size_combo.clear()
        for value in options:
            profile_size_combo.addItem(value)
        _set_combo_text_preserving_custom(profile_size_combo, wanted)
        profile_size_combo.blockSignals(False)

    def _refresh_angle_thickness_options() -> None:
        leg = str(angle_leg_combo.currentText() or "").strip()
        wanted = str(initial.get("angle_thickness", "6") or "6")
        angle_thickness_combo.blockSignals(True)
        angle_thickness_combo.clear()
        for value in equal_angle_presets.get(leg, ["6"]):
            angle_thickness_combo.addItem(value)
        angle_thickness_combo.setCurrentText(wanted)
        angle_thickness_combo.blockSignals(False)

    def _refresh_tube_model_options() -> None:
        secao = str(tube_section_combo.currentData() or "").strip().lower() or "quadrado"
        wanted = str(initial.get("tube_model", "") or "").strip()
        options = list(tube_model_presets.get(secao, []))
        if secao == "redondo" and not wanted and float(tube_diameter_spin.value() or 0) > 0 and float(tube_thickness_spin.value() or 0) > 0:
            wanted = f"Ø{tube_diameter_spin.value():.1f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
        elif not wanted and float(tube_side_a_spin.value() or 0) > 0 and float(tube_thickness_spin.value() or 0) > 0:
            if secao == "retangular" and float(tube_side_b_spin.value() or 0) > 0:
                wanted = f"{tube_side_a_spin.value():.0f}x{tube_side_b_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
            else:
                wanted = f"{tube_side_a_spin.value():.0f}x{tube_side_a_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
        tube_model_combo.blockSignals(True)
        tube_model_combo.clear()
        for value in options:
            tube_model_combo.addItem(value)
        _set_combo_text_preserving_custom(tube_model_combo, wanted)
        tube_model_combo.blockSignals(False)

    def _apply_tube_model_selection() -> None:
        if tube_source_combo.currentText().strip() == "Stock MP":
            return
        model_txt = tube_model_combo.currentText().strip().lower().replace(",", ".").replace("mm", "")
        if not model_txt:
            return
        numbers = [float(chunk) for chunk in re.findall(r"\d+(?:\.\d+)?", model_txt)]
        secao = str(tube_section_combo.currentData() or "").strip().lower()
        if secao == "redondo":
            if len(numbers) >= 2:
                tube_diameter_spin.setValue(numbers[0])
                tube_thickness_spin.setValue(numbers[1])
        elif secao == "retangular":
            if len(numbers) >= 3:
                tube_side_a_spin.setValue(numbers[0])
                tube_side_b_spin.setValue(numbers[1])
                tube_thickness_spin.setValue(numbers[2])
        else:
            if len(numbers) >= 3:
                tube_side_a_spin.setValue(numbers[0])
                tube_side_b_spin.setValue(numbers[0])
                tube_thickness_spin.setValue(numbers[2])
            elif len(numbers) >= 2:
                tube_side_a_spin.setValue(numbers[0])
                tube_side_b_spin.setValue(numbers[0])
                tube_thickness_spin.setValue(numbers[1])

    def _refresh_sheet_model_options() -> None:
        wanted = str(initial.get("sheet_model", "") or "").strip()
        if not wanted and float(length_mm_spin.value() or 0) > 0 and float(width_mm_spin.value() or 0) > 0 and float(thickness_mm_spin.value() or 0) > 0:
            wanted = f"{length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{thickness_mm_spin.value():.1f}".replace(".0", "")
        checker_mode = str(sheet_kind_combo.currentData() or "").strip() == "checker"
        if checker_mode and (not wanted or "/" not in wanted):
            wanted = "Chapa gota 3000x1500x6/8"
        options = []
        for value in sheet_model_presets:
            normalized_value = backend.norm_text(value)
            value_is_checker = "/" in str(value or "") or any(token in normalized_value for token in ("gota", "xadrez", "antiderrap"))
            if checker_mode == value_is_checker:
                options.append(value)
        sheet_model_combo.blockSignals(True)
        sheet_model_combo.clear()
        for value in options:
            sheet_model_combo.addItem(value)
        _set_combo_text_preserving_custom(sheet_model_combo, wanted)
        sheet_model_combo.blockSignals(False)

    def _apply_sheet_model_selection() -> None:
        if sheet_source_combo.currentText().strip() == "Stock MP":
            return
        model = _sheet_model_info(sheet_model_combo.currentText())
        if not str(model.get("raw", "") or "").strip():
            return
        if float(model.get("length", 0) or 0) > 0:
            length_mm_spin.setValue(float(model.get("length", 0) or 0))
        if float(model.get("width", 0) or 0) > 0:
            width_mm_spin.setValue(float(model.get("width", 0) or 0))
        if float(model.get("thickness", 0) or 0) > 0:
            thickness_mm_spin.setValue(float(model.get("thickness", 0) or 0))

    _fill_stock_combo(profile_stock_combo, "Perfil", str(initial.get("stock_material_id", "") or "").strip())
    _fill_stock_combo(tube_stock_combo, "Tubo", str(initial.get("stock_material_id", "") or "").strip())
    _fill_stock_combo(sheet_stock_combo, "Chapa", str(initial.get("stock_material_id", "") or "").strip())
    _fill_stock_combo(angle_stock_combo, "Cantoneira", str(initial.get("stock_material_id", "") or "").strip())
    _fill_stock_combo(bar_stock_combo, "Barra", str(initial.get("stock_material_id", "") or "").strip())
    _refresh_profile_size_options()
    _refresh_angle_thickness_options()
    _refresh_tube_model_options()
    _refresh_sheet_model_options()
    if str(initial.get("stock_material_id", "") or "").strip():
        current_mode = str(initial.get("calc_mode", "Perfil") or "Perfil").strip()
        if current_mode == "Perfil":
            profile_source_combo.setCurrentText("Stock MP")
        elif current_mode == "Tubo":
            tube_source_combo.setCurrentText("Stock MP")
        elif current_mode == "Chapa":
            sheet_source_combo.setCurrentText("Stock MP")
        elif current_mode == "Cantoneira":
            angle_source_combo.setCurrentText("Stock MP")
        elif current_mode == "Barra":
            bar_source_combo.setCurrentText("Stock MP")

    stock_sync_state = {
        "profile": "",
        "tube": "",
        "sheet": "",
        "angle": "",
        "bar": "",
    }

    default_price_map = {
        "Perfil": float(backend.material_default_price_kg("Perfil") or 3.15),
        "Tubo": float(backend.material_default_price_kg("Tubo") or 3.15),
        "Chapa": float(backend.material_default_price_kg("Chapa") or 3.15),
        "Cantoneira": float(backend.material_default_price_kg("Cantoneira") or backend.material_default_price_kg("Perfil") or 3.15),
        "Barra": float(backend.material_default_price_kg("Barra") or backend.material_default_price_kg("Chapa") or 3.15),
        "Ferro nervurado": float(backend.material_default_price_kg("Barra", "ferro") or backend.material_default_price_kg("Barra") or 3.15),
    }
    if float(initial.get("price_per_kg", 0) or 0) <= 0:
        profile_price_kg_spin.setValue(default_price_map["Perfil"])
        tube_price_kg_spin.setValue(default_price_map["Tubo"])
        sheet_price_kg_spin.setValue(default_price_map["Chapa"])
        angle_price_kg_spin.setValue(default_price_map["Cantoneira"])
        bar_price_kg_spin.setValue(default_price_map["Barra"])
        rebar_price_kg_spin.setValue(default_price_map["Ferro nervurado"])

    def _profile_payload() -> dict:
        if profile_source_combo.currentText().strip() == "Stock MP":
            record, preview = _current_stock(profile_stock_combo)
            if record:
                identity = _stock_material_identity("Perfil", record, preview)
                series, size_text = _profile_series_and_size(record, preview)
                meters = float(profile_meters_spin.value() or record.get("metros", 0) or 0.0)
                kg_m = float(profile_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                return {
                    "stock_material_id": str(record.get("id", "") or "").strip(),
                    "ref": identity["ref"],
                    "desc": identity["desc"],
                    "meters": meters,
                    "kg_m": kg_m,
                    "price_kg": float(profile_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                    "weight_each": round(meters * kg_m, 4),
                    "hint": str(preview.get("calc_hint", "") or "").strip(),
                    "quality": identity["quality"],
                    "series": series,
                    "size": size_text,
                }
        series = str(profile_series_combo.currentData() or "").strip()
        size_text, size_mm = _parse_size_value(profile_size_combo.currentText())
        preview = dict(
            backend.material_geometry_preview(
                {
                    "formato": "Perfil",
                    "material": f"{series} {size_text}".strip(),
                    "secao_tipo": series,
                    "altura": size_mm,
                    "perfil_tamanho": size_text,
                    "metros": float(profile_meters_spin.value() or 0.0),
                }
            )
            or {}
        )
        desc = f"Perfil {series} {size_text}".strip()
        return {
            "stock_material_id": "",
            "ref": f"{series} {size_text}".strip(),
            "desc": desc,
            "meters": float(profile_meters_spin.value() or 0.0),
            "kg_m": float(preview.get("kg_m", 0) or 0.0),
            "price_kg": float(profile_price_kg_spin.value() or 0.0),
            "weight_each": float(preview.get("peso_unid", 0) or 0.0),
            "hint": str(preview.get("calc_hint", "") or "").strip(),
            "series": series,
            "size": size_text,
        }

    def _tube_payload() -> dict:
        if tube_source_combo.currentText().strip() == "Stock MP":
            record, preview = _current_stock(tube_stock_combo)
            if record:
                identity = _stock_material_identity("Tubo", record, preview)
                meters = float(tube_meters_spin.value() or record.get("metros", 0) or 0.0)
                kg_m = float(tube_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                base_label, base_value = _stock_price_base(record, preview)
                if base_label.strip().upper() != "EUR/M":
                    base_label = "EUR/m"
                price_m = float(tube_price_kg_spin.value() or base_value or 0.0)
                price_kg_equiv = round(price_m / kg_m, 4) if kg_m > 0 else 0.0
                return {
                    "stock_material_id": str(record.get("id", "") or "").strip(),
                    "ref": identity["ref"],
                    "desc": identity["desc"],
                    "meters": meters,
                    "kg_m": kg_m,
                    "price_kg": price_kg_equiv,
                    "price_base_label": "EUR/m",
                    "price_base_value": price_m,
                    "weight_each": round(meters * kg_m, 4),
                    "hint": str(preview.get("calc_hint", "") or "").strip(),
                    "quality": identity["quality"],
                }
        secao = str(tube_section_combo.currentData() or "").strip()
        quality = tube_quality_edit.text().strip() or "S235JR"
        payload = {
            "formato": "Tubo",
            "material": quality,
            "secao_tipo": secao,
            "espessura": _float_text(tube_thickness_spin.value(), 2).replace(",", "."),
            "metros": float(tube_meters_spin.value() or 0.0),
            "comprimento": float(tube_side_a_spin.value() or 0.0),
            "largura": float(tube_side_b_spin.value() or 0.0),
            "diametro": float(tube_diameter_spin.value() or 0.0),
        }
        preview = dict(backend.material_geometry_preview(payload) or {})
        meters = float(tube_meters_spin.value() or 0.0)
        kg_m = float(preview.get("kg_m", 0) or 0.0)
        price_m = float(tube_price_kg_spin.value() or 0.0)
        price_kg_equiv = round(price_m / kg_m, 4) if kg_m > 0 else 0.0
        if secao == "redondo":
            desc = f"Tubo {quality} Ø{tube_diameter_spin.value():.1f}x{tube_thickness_spin.value():.1f} mm".replace(".0", "")
            ref = f"TUB-RED-{tube_diameter_spin.value():.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
        else:
            a_val = tube_side_a_spin.value()
            b_val = tube_side_b_spin.value() if secao == "retangular" else tube_side_a_spin.value()
            dim_txt = f"{a_val:.0f}x{b_val:.0f}x{tube_thickness_spin.value():.1f}".replace(".0", "")
            desc = f"Tubo {quality} {dim_txt} mm"
            ref = f"TUB-{dim_txt}"
        return {
            "stock_material_id": "",
            "ref": ref,
            "desc": desc,
            "meters": meters,
            "kg_m": kg_m,
            "price_kg": price_kg_equiv,
            "price_base_label": "EUR/m",
            "price_base_value": price_m,
            "weight_each": float(preview.get("peso_unid", 0) or 0.0),
            "hint": str(preview.get("calc_hint", "") or "").strip(),
            "quality": quality,
        }

    def _sheet_payload() -> dict:
        if sheet_source_combo.currentText().strip() == "Stock MP":
            record, preview = _current_stock(sheet_stock_combo)
            if record:
                identity = _stock_material_identity("Chapa", record, preview)
                return {
                    "stock_material_id": str(record.get("id", "") or "").strip(),
                    "ref": identity["ref"],
                    "desc": identity["desc"],
                    "meters": 0.0,
                    "kg_m": 0.0,
                    "price_kg": float(sheet_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                    "weight_each": float(preview.get("peso_unid", 0) or 0.0),
                    "hint": str(preview.get("calc_hint", "") or "").strip(),
                    "quality": identity["quality"],
                }
        quality = sheet_quality_combo.currentText().strip() or "S235JR"
        family_key = _sheet_family_key()
        model = _sheet_model_info(sheet_model_combo.currentText())
        is_checker = bool(model.get("is_checker"))
        thickness_label = str(model.get("thickness_label", "") or "").strip()
        display_thickness = thickness_label or _float_text(thickness_mm_spin.value(), 2).replace(",00", "").replace(",0", "")
        sheet_kind = "Chapa gota" if is_checker else "Chapa"
        density_g_cm3 = round(float(density_spin.value() or 7850.0) / 1000.0, 4)
        preview = dict(
            backend.material_geometry_preview(
                {
                    "formato": "Chapa",
                    "material": quality,
                    "comprimento": float(length_mm_spin.value() or 0.0),
                    "largura": float(width_mm_spin.value() or 0.0),
                    "espessura": display_thickness.replace(",", "."),
                    "material_familia": family_key,
                    "densidade": density_g_cm3,
                    "descricao": f"{sheet_kind} {quality}",
                }
            )
            or {}
        )
        desc = (
            f"{sheet_kind} {quality} {length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{display_thickness} mm"
        ).replace(".0", "")
        ref_prefix = "CH-GOTA" if is_checker else "CH"
        ref = f"{ref_prefix}-{length_mm_spin.value():.0f}x{width_mm_spin.value():.0f}x{display_thickness}".replace(".0", "")
        return {
            "stock_material_id": "",
            "ref": ref,
            "desc": desc,
            "meters": 0.0,
            "kg_m": 0.0,
            "price_kg": float(sheet_price_kg_spin.value() or 0.0),
            "weight_each": float(preview.get("peso_unid", 0) or 0.0),
            "hint": str(preview.get("calc_hint", "") or "").strip(),
            "quality": quality,
            "material_family_key": family_key,
            "esp_label": display_thickness,
            "sheet_kind": sheet_kind,
        }

    def _angle_payload() -> dict:
        if angle_source_combo.currentText().strip() == "Stock MP":
            record, preview = _current_stock(angle_stock_combo)
            if record:
                identity = _stock_material_identity("Cantoneira", record, preview)
                meters = float(angle_meters_spin.value() or record.get("metros", 0) or 0.0)
                kg_m = float(angle_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                return {
                    "stock_material_id": str(record.get("id", "") or "").strip(),
                    "ref": identity["ref"],
                    "desc": identity["desc"],
                    "meters": meters,
                    "kg_m": kg_m,
                    "price_kg": float(angle_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                    "weight_each": round(meters * kg_m, 4),
                    "hint": str(preview.get("calc_hint", "") or "").strip(),
                    "quality": identity["quality"],
                    "material_family_key": str(record.get("material_familia", preview.get("material_familia_resolved", "")) or "").strip(),
                }
        leg = float(angle_leg_combo.currentText() or 0.0)
        thickness = float(angle_thickness_combo.currentText() or 0.0)
        density = _family_density_g_cm3(angle_family_combo)
        kg_m = _geometry_weight_kg_m((thickness * ((2.0 * leg) - thickness)), density)
        meters = float(angle_meters_spin.value() or 0.0)
        quality = angle_quality_combo.currentText().strip() or "S235JR"
        desc = f"Cantoneira abas iguais {quality} {leg:.0f}x{leg:.0f}x{thickness:.0f} mm"
        ref = f"L {leg:.0f}x{leg:.0f}x{thickness:.0f}"
        return {
            "stock_material_id": "",
            "ref": ref,
            "desc": desc,
            "meters": meters,
            "kg_m": kg_m,
            "price_kg": float(angle_price_kg_spin.value() or 0.0),
            "weight_each": round(kg_m * meters, 4),
            "hint": f"Cantoneira: área aproximada t x (2a - t) x densidade {density:.3f} g/cm3.",
            "quality": quality,
            "material_family_key": _family_key(angle_family_combo),
        }

    def _bar_payload() -> dict:
        if bar_source_combo.currentText().strip() == "Stock MP":
            record, preview = _current_stock(bar_stock_combo)
            if record:
                identity = _stock_material_identity("Barra", record, preview)
                meters = float(bar_meters_spin.value() or record.get("metros", 0) or 0.0)
                kg_m = float(bar_kg_per_m_spin.value() or preview.get("kg_m", 0) or record.get("kg_m", 0) or 0.0)
                return {
                    "stock_material_id": str(record.get("id", "") or "").strip(),
                    "ref": identity["ref"],
                    "desc": identity["desc"],
                    "meters": meters,
                    "kg_m": kg_m,
                    "price_kg": float(bar_price_kg_spin.value() or _stock_price_kg(record, preview) or 0.0),
                    "weight_each": round(meters * kg_m, 4),
                    "hint": str(preview.get("calc_hint", "") or "").strip(),
                    "quality": identity["quality"],
                    "material_family_key": str(record.get("material_familia", preview.get("material_familia_resolved", "")) or "").strip(),
                }
        width_mm, thickness_mm = _parse_pair(bar_size_combo.currentText())
        density = _family_density_g_cm3(bar_family_combo)
        kg_m = _geometry_weight_kg_m(width_mm * thickness_mm, density)
        meters = float(bar_meters_spin.value() or 0.0)
        quality = bar_quality_combo.currentText().strip() or "S235JR"
        desc = f"Barra chata {quality} {width_mm:.0f}x{thickness_mm:.0f} mm"
        ref = f"BAR {width_mm:.0f}x{thickness_mm:.0f}"
        return {
            "stock_material_id": "",
            "ref": ref,
            "desc": desc,
            "meters": meters,
            "kg_m": kg_m,
            "price_kg": float(bar_price_kg_spin.value() or 0.0),
            "weight_each": round(kg_m * meters, 4),
            "hint": f"Barra chata: largura x espessura x densidade {density:.3f} g/cm3.",
            "quality": quality,
            "material_family_key": _family_key(bar_family_combo),
        }

    def _rebar_payload() -> dict:
        meters = float(rebar_meters_spin.value() or 0.0)
        diameter = float(diameter_mm_spin.value() or 0.0)
        kg_m = round(0.006165 * (diameter ** 2), 4)
        desc = f"Ferro nervurado Ø{diameter:.0f} mm".replace(".0", "")
        ref = f"REBAR {diameter:.0f}".replace(".0", "")
        return {
            "stock_material_id": "",
            "ref": ref,
            "desc": desc,
            "meters": meters,
            "kg_m": kg_m,
            "price_kg": float(rebar_price_kg_spin.value() or 0.0),
            "weight_each": round(kg_m * meters, 4),
            "hint": "Ferro nervurado: 0.006165 x Ø².",
        }

    def _manual_payload() -> dict:
        unit_txt = unit_combo.currentText().strip() or "un"
        return {
            "stock_material_id": "",
            "ref": ref_edit.text().strip(),
            "desc": desc_edit.text().strip() or "Material manual",
            "meters": 0.0,
            "kg_m": 0.0,
            "price_kg": 0.0,
            "weight_each": 0.0,
            "hint": "",
            "unit": unit_txt,
            "unit_price": float(manual_unit_price_spin.value() or 0.0),
        }

    def _mode_payload() -> dict:
        mode = mode_combo.currentText().strip()
        if mode == "Perfil":
            return _profile_payload()
        if mode == "Tubo":
            return _tube_payload()
        if mode == "Chapa":
            return _sheet_payload()
        if mode == "Cantoneira":
            return _angle_payload()
        if mode == "Barra":
            return _bar_payload()
        if mode == "Ferro nervurado":
            return _rebar_payload()
        return _manual_payload()

    def _selected_operation_names() -> list[str]:
        selected: list[str] = []
        for token in _operation_tokens(operation_edit.text().strip()):
            normalized = str(backend.normalize_operacao_nome(token) or token or "").strip()
            if normalized and normalized not in selected:
                selected.append(normalized)
        return selected

    def _sync_operation_meta_selection() -> None:
        selected_ops = _selected_operation_names()
        selected_keys = {
            str(backend.normalize_operacao_nome(op) or op or "").strip()
            for op in selected_ops
            if str(op or "").strip()
        }
        operation_meta["operacoes_lista"] = list(selected_ops)
        operation_meta["operacoes_fluxo"] = backend.build_operacoes_fluxo(
            selected_ops,
            operation_meta.get("operacoes_fluxo") if isinstance(operation_meta.get("operacoes_fluxo"), list) else None,
        )
        operation_meta["operacoes_detalhe"] = [
            dict(item or {})
            for item in list(operation_meta.get("operacoes_detalhe", []) or [])
            if str(backend.normalize_operacao_nome((item or {}).get("nome", "")) or (item or {}).get("nome", "") or "").strip() in selected_keys
        ]
        operation_meta["tempos_operacao"] = {
            str(backend.normalize_operacao_nome(op_name) or op_name or "").strip(): float(value or 0)
            for op_name, value in dict(operation_meta.get("tempos_operacao", {}) or {}).items()
            if str(backend.normalize_operacao_nome(op_name) or op_name or "").strip() in selected_keys
        }
        operation_meta["custos_operacao"] = {
            str(backend.normalize_operacao_nome(op_name) or op_name or "").strip(): float(value or 0)
            for op_name, value in dict(operation_meta.get("custos_operacao", {}) or {}).items()
            if str(backend.normalize_operacao_nome(op_name) or op_name or "").strip() in selected_keys
        }

    def _operation_totals() -> tuple[float, float]:
        selected_keys = {
            str(backend.normalize_operacao_nome(op) or op or "").strip()
            for op in _selected_operation_names()
            if str(op or "").strip()
        }
        time_total = 0.0
        cost_total = 0.0
        for op_name, value in dict(operation_meta.get("tempos_operacao", {}) or {}).items():
            normalized = str(backend.normalize_operacao_nome(op_name) or op_name or "").strip()
            if normalized in selected_keys:
                time_total += float(value or 0)
        for op_name, value in dict(operation_meta.get("custos_operacao", {}) or {}).items():
            normalized = str(backend.normalize_operacao_nome(op_name) or op_name or "").strip()
            if normalized in selected_keys:
                cost_total += float(value or 0)
        return round(time_total, 4), round(cost_total, 4)

    def _operation_cost_payload(base_unit_cost: float = 0.0, qty_value: float | None = None) -> dict:
        return {
            "operacao": operation_edit.text().strip(),
            "operacoes_lista": list(_selected_operation_names()),
            "costing_operations": list(_selected_operation_names()),
            "qtd": float(qty_value if qty_value is not None else qty_spin.value() or 0),
            "tempo_peca_min": float(_operation_totals()[0] or 0),
            "preco_unit": float(base_unit_cost or 0),
            "blend_with_current_line": True,
            "base_tempo_unit_min": 0.0,
            "base_preco_unit_eur": float(base_unit_cost or 0),
            "base_operation_label": "Material base",
            "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
        }

    def _refresh_operation_cost_hint(base_unit_cost: float = 0.0) -> None:
        _sync_operation_meta_selection()
        ops = _selected_operation_names()
        if not ops:
            operation_cost_label.setText("Sem operações seguintes. O item fica apenas como matéria-prima/stock.")
            return
        estimate = dict(backend.operation_cost_estimate(_operation_cost_payload(base_unit_cost)) or {})
        summary = dict(estimate.get("summary", {}) or {})
        pending = [
            dict(item)
            for item in list(estimate.get("operations", []) or [])
            if isinstance(item, dict) and bool(item.get("missing_driver_input"))
        ]
        if pending:
            missing_txt = ", ".join(str(item.get("nome", "") or "").strip() for item in pending if str(item.get("nome", "") or "").strip())
            operation_cost_label.setText(f"Falta quantificar: {missing_txt}.")
            return
        extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
        extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
        operation_cost_label.setText(
            f"Operações: {extra_time:.3f} min/un | {_fmt_eur(extra_cost)}/un | "
            f"total com material {_fmt_eur(float(base_unit_cost or 0) + extra_cost)}/un"
        )

    def _edit_operation_costs() -> bool:
        _sync_operation_meta_selection()
        if not _selected_operation_names():
            QMessageBox.information(dialog, "Operações", "Seleciona primeiro as operações seguintes deste material.")
            return False
        current = _compute()
        base_cost = float((current.get("line") or {}).get("_material_base_preco_unit", (current.get("line") or {}).get("preco_unit", 0)) or 0)
        result = _open_quote_operation_detail_dialog(dialog, backend, _operation_cost_payload(base_cost))
        if not isinstance(result, dict):
            return False
        operation_meta["operacoes_detalhe"] = [dict(item or {}) for item in list(result.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        operation_meta["tempos_operacao"] = dict(result.get("tempos_operacao", {}) or {})
        operation_meta["custos_operacao"] = dict(result.get("custos_operacao", {}) or {})
        operation_meta["quote_cost_snapshot"] = dict(result.get("quote_cost_snapshot", {}) or {})
        _refresh()
        return True

    def _apply_auto_texts(payload: dict) -> None:
        if sync["busy"]:
            return
        auto_desc = str(payload.get("desc", "") or "").strip()
        auto_ref = str(payload.get("ref", "") or "").strip()
        if auto_desc:
            _set_edit_if_auto(desc_edit, desc_state, auto_desc)
        if auto_ref:
            _set_edit_if_auto(ref_edit, ref_state, auto_ref)

    def _toggle_controls() -> None:
        mode = mode_combo.currentText().strip()
        stack.setCurrentIndex(mode_to_index.get(mode, 0))
        profile_is_stock = profile_source_combo.currentText().strip() == "Stock MP"
        tube_is_stock = tube_source_combo.currentText().strip() == "Stock MP"
        sheet_is_stock = sheet_source_combo.currentText().strip() == "Stock MP"
        angle_is_stock = angle_source_combo.currentText().strip() == "Stock MP"
        bar_is_stock = bar_source_combo.currentText().strip() == "Stock MP"

        profile_stock_combo.setVisible(profile_is_stock)
        profile_form.labelForField(profile_stock_combo).setVisible(profile_is_stock)
        for widget in (profile_series_combo, profile_size_combo):
            widget.setEnabled(not profile_is_stock)

        tube_stock_combo.setVisible(tube_is_stock)
        tube_form.labelForField(tube_stock_combo).setVisible(tube_is_stock)
        for widget in (tube_quality_edit, tube_section_combo, tube_model_combo, tube_side_a_spin, tube_side_b_spin, tube_diameter_spin, tube_thickness_spin):
            widget.setEnabled(not tube_is_stock)
        tube_model_combo.setVisible(not tube_is_stock)
        tube_form.labelForField(tube_model_combo).setVisible(not tube_is_stock)
        tube_is_round = str(tube_section_combo.currentData() or "").strip() == "redondo"
        tube_side_a_spin.setVisible(not tube_is_round)
        tube_form.labelForField(tube_side_a_spin).setVisible(not tube_is_round)
        tube_side_b_spin.setVisible(not tube_is_round and str(tube_section_combo.currentData() or "").strip() == "retangular")
        tube_form.labelForField(tube_side_b_spin).setVisible(not tube_is_round and str(tube_section_combo.currentData() or "").strip() == "retangular")
        tube_diameter_spin.setVisible(tube_is_round)
        tube_form.labelForField(tube_diameter_spin).setVisible(tube_is_round)

        sheet_stock_combo.setVisible(sheet_is_stock)
        sheet_form.labelForField(sheet_stock_combo).setVisible(sheet_is_stock)
        for widget in (sheet_family_combo, sheet_quality_combo, sheet_kind_combo, sheet_model_combo, length_mm_spin, width_mm_spin, thickness_mm_spin, density_spin):
            widget.setEnabled(not sheet_is_stock)
        sheet_model_combo.setVisible(not sheet_is_stock)
        sheet_form.labelForField(sheet_model_combo).setVisible(not sheet_is_stock)

        angle_stock_combo.setVisible(angle_is_stock)
        angle_form.labelForField(angle_stock_combo).setVisible(angle_is_stock)
        for widget in (angle_family_combo, angle_quality_combo, angle_leg_combo, angle_thickness_combo):
            widget.setEnabled(not angle_is_stock)

        bar_stock_combo.setVisible(bar_is_stock)
        bar_form.labelForField(bar_stock_combo).setVisible(bar_is_stock)
        for widget in (bar_family_combo, bar_quality_combo, bar_size_combo):
            widget.setEnabled(not bar_is_stock)

    def _compute() -> dict:
        mode = mode_combo.currentText().strip()
        qty_units = float(qty_spin.value() or 0.0)
        payload = dict(_mode_payload() or {})
        description = desc_edit.text().strip() or str(payload.get("desc", "") or mode).strip() or mode
        reference = ref_edit.text().strip() or str(payload.get("ref", "") or "").strip()
        total_weight = 0.0
        unit_cost = 0.0
        total_cost = 0.0
        qty_line = qty_units
        unit_label = "un"
        hint = str(payload.get("hint", "") or "").strip()
        weight_each = float(payload.get("weight_each", 0) or 0.0)
        meters_per_unit = float(payload.get("meters", 0) or 0.0)
        kg_per_m = float(payload.get("kg_m", 0) or 0.0)
        price_base_label = str(payload.get("price_base_label", "EUR/kg" if mode != "Tubo" else "EUR/m") or "").strip() or ("EUR/m" if mode == "Tubo" else "EUR/kg")
        price_base_value = float(payload.get("price_base_value", payload.get("price_kg", 0)) or 0.0)
        if mode == "Manual":
            unit_label = str(payload.get("unit", unit_combo.currentText()) or "un").strip() or "un"
            unit_cost = round(float(payload.get("unit_price", manual_unit_price_spin.value()) or 0.0), 4)
            total_cost = round(qty_units * unit_cost, 2)
            detail = f"{description} | {qty_units:.2f} {unit_label}"
        else:
            if weight_each <= 0 and meters_per_unit > 0 and kg_per_m > 0:
                weight_each = round(meters_per_unit * kg_per_m, 4)
            total_weight = round(weight_each * qty_units, 3)
            if mode == "Tubo" and price_base_label.upper() == "EUR/M":
                unit_cost = round(meters_per_unit * price_base_value, 4) if meters_per_unit > 0 else 0.0
            else:
                unit_cost = round(weight_each * float(payload.get("price_kg", price_base_value) or 0.0), 4) if weight_each > 0 else 0.0
            total_cost = round(unit_cost * qty_units, 2)
            if meters_per_unit > 0:
                base_txt = f" | {price_base_value:.4f} {price_base_label}" if price_base_value > 0 else ""
                detail = f"{description} | {qty_units:.2f} un x {meters_per_unit:.2f} m | {kg_per_m:.3f} kg/m | {total_weight:.2f} kg{base_txt}"
            elif total_weight > 0:
                detail = f"{description} | {qty_units:.2f} un | {total_weight:.2f} kg"
            else:
                detail = f"{description} | {qty_units:.2f} un"
        material_base_cost = round(unit_cost, 4)
        operation_time_unit, operation_cost_unit = _operation_totals()
        unit_cost = round(unit_cost + operation_cost_unit, 4)
        total_cost = round(unit_cost * qty_units, 2)
        thickness_value = 0.0
        if mode == "Chapa":
            thickness_value = float(thickness_mm_spin.value() or 0.0)
        elif mode == "Tubo":
            thickness_value = float(tube_thickness_spin.value() or 0.0)
        elif mode == "Cantoneira":
            thickness_value = float(angle_thickness_combo.currentText() or 0.0)
        elif mode == "Barra":
            _bar_w, thickness_value = _parse_pair(bar_size_combo.currentText())
        elif mode == "Ferro nervurado":
            thickness_value = float(diameter_mm_spin.value() or 0.0)

        quality_value = str(payload.get("quality", "") or "").strip()
        material_label = quality_value
        if mode != "Manual":
            material_label = f"{mode} {quality_value}".strip() if quality_value else mode
        if mode == "Chapa" and str(payload.get("sheet_kind", "") or "").strip():
            sheet_kind = str(payload.get("sheet_kind", "") or "").strip()
            material_label = f"{sheet_kind} {quality_value}".strip() if quality_value else sheet_kind
        esp_txt = ""
        if mode == "Chapa" and str(payload.get("esp_label", "") or "").strip():
            esp_txt = str(payload.get("esp_label", "") or "").strip()
        elif thickness_value > 0:
            esp_txt = _float_text(thickness_value, 2).replace(",00", "").replace(",0", "")
        elif mode == "Perfil":
            esp_txt = str(payload.get("size", "") or profile_size_combo.currentText() or "").strip()
        elif mode == "Manual":
            esp_txt = str(unit_label).strip()
        line = {
            "tipo_item": backend.ORC_LINE_TYPE_PIECE,
            "stock_item_kind": "raw_material",
            "descricao": detail,
            "ref_externa": reference,
            "material": material_label or mode,
            "material_family": quality_value or material_label or mode,
            "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
            "material_subtype": mode,
            "espessura": esp_txt,
            "operacao": " + ".join(_selected_operation_names()),
            "qtd": round(qty_line, 2),
            "preco_unit": round(unit_cost, 4),
            "desenho": "",
            "tempo_peca_min": round(operation_time_unit, 4),
            "calc_mode": mode,
            "descricao_base": description,
            "weight_total": round(total_weight, 3),
            "total_cost": round(total_cost, 2),
            "quantity_units": round(qty_units, 2),
            "price_per_kg": round(float(payload.get("price_kg", 0) or 0.0), 4),
            "price_base_label": price_base_label,
            "price_base_value": round(price_base_value, 4),
            "stock_metric_value": round(meters_per_unit if price_base_label.upper() == "EUR/M" else weight_each, 4),
            "meters_per_unit": round(meters_per_unit, 3),
            "kg_per_m": round(kg_per_m, 4),
            "length_mm": round(float(length_mm_spin.value() or 0.0), 1),
            "width_mm": round(float(width_mm_spin.value() or 0.0), 1),
            "thickness_mm": round(float(thickness_mm_spin.value() or tube_thickness_spin.value() or 0.0), 2),
            "density": round(float(density_spin.value() or 0.0), 1),
            "diameter_mm": round(float(tube_diameter_spin.value() or diameter_mm_spin.value() or 0.0), 1),
            "manual_unit_price": round(float(manual_unit_price_spin.value() or 0.0), 4),
            "profile_section": str(payload.get("series", "") or profile_series_combo.currentData() or "").strip(),
            "profile_size": str(payload.get("size", "") or profile_size_combo.currentText() or "").strip(),
            "tube_section": str(tube_section_combo.currentData() or "").strip(),
            "tube_model": tube_model_combo.currentText().strip(),
            "sheet_model": sheet_model_combo.currentText().strip(),
            "sheet_kind": str(payload.get("sheet_kind", "") or "").strip(),
            "sheet_esp_label": str(payload.get("esp_label", "") or "").strip(),
            "bar_size": bar_size_combo.currentText().strip(),
            "quality": str(payload.get("quality", "") or "").strip(),
            "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
            "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
            "hint": hint,
            "_material_base_preco_unit": material_base_cost,
            "operacoes_lista": list(operation_meta.get("operacoes_lista", []) or []),
            "operacoes_fluxo": [dict(item or {}) for item in list(operation_meta.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
            "operacoes_detalhe": [dict(item or {}) for item in list(operation_meta.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
            "tempos_operacao": dict(operation_meta.get("tempos_operacao", {}) or {}),
            "custos_operacao": dict(operation_meta.get("custos_operacao", {}) or {}),
            "quote_cost_snapshot": dict(operation_meta.get("quote_cost_snapshot", {}) or {}),
        }
        return {
            "kind": "material",
            "stock_item_kind": "raw_material",
            "calc_mode": mode,
            "descricao_base": description,
            "weight_total": round(total_weight, 3),
            "total_cost": round(total_cost, 2),
            "quantity_units": round(qty_units, 2),
            "produto_unid": unit_label,
            "price_per_kg": round(float(payload.get("price_kg", 0) or 0.0), 4),
            "price_base_label": price_base_label,
            "price_base_value": round(price_base_value, 4),
            "stock_metric_value": round(meters_per_unit if price_base_label.upper() == "EUR/M" else weight_each, 4),
            "meters_per_unit": round(meters_per_unit, 3),
            "kg_per_m": round(kg_per_m, 4),
            "length_mm": round(float(length_mm_spin.value() or 0.0), 1),
            "width_mm": round(float(width_mm_spin.value() or 0.0), 1),
            "thickness_mm": round(float(thickness_mm_spin.value() or tube_thickness_spin.value() or 0.0), 2),
            "density": round(float(density_spin.value() or 0.0), 1),
            "diameter_mm": round(float(tube_diameter_spin.value() or diameter_mm_spin.value() or 0.0), 1),
            "manual_unit_price": round(float(manual_unit_price_spin.value() or 0.0), 4),
            "profile_section": str(profile_series_combo.currentData() or "").strip(),
            "profile_size": profile_size_combo.currentText().strip(),
            "tube_section": str(tube_section_combo.currentData() or "").strip(),
            "tube_model": tube_model_combo.currentText().strip(),
            "sheet_model": sheet_model_combo.currentText().strip(),
            "sheet_kind": str(payload.get("sheet_kind", "") or "").strip(),
            "sheet_esp_label": str(payload.get("esp_label", "") or "").strip(),
            "bar_size": bar_size_combo.currentText().strip(),
            "quality": str(payload.get("quality", "") or "").strip(),
            "material_family_key": str(payload.get("material_family_key", "") or "").strip(),
            "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
            "hint": hint,
            "operation_cost_unit": round(operation_cost_unit, 4),
            "operation_time_unit": round(operation_time_unit, 4),
            "line": line,
        }

    def _refresh() -> None:
        if sync["busy"]:
            return
        sync["busy"] = True
        try:
            _toggle_controls()
            payload = dict(_mode_payload() or {})
            if mode_combo.currentText().strip() == "Perfil":
                if profile_source_combo.currentText().strip() == "Stock MP":
                    record, preview = _current_stock(profile_stock_combo)
                    if record:
                        current_stock_id = str(record.get("id", "") or "").strip()
                        if stock_sync_state["profile"] != current_stock_id:
                            series_txt, size_txt = _profile_series_and_size(record, preview)
                            stock_sync_state["profile"] = current_stock_id
                            profile_meters_spin.setValue(float(record.get("metros", profile_meters_spin.value()) or profile_meters_spin.value() or 0.0))
                            _set_combo_by_data(profile_series_combo, series_txt)
                            _refresh_profile_size_options()
                            if size_txt:
                                profile_size_combo.setCurrentText(size_txt)
                            profile_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                            profile_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                        identity = _stock_material_identity("Perfil", record, preview)
                        profile_hint.setText(str(payload.get("hint", "") or "").strip())
                        _set_edit_if_auto(desc_edit, desc_state, identity["desc"])
                        _set_edit_if_auto(ref_edit, ref_state, identity["ref"])
                    else:
                        stock_sync_state["profile"] = ""
                else:
                    stock_sync_state["profile"] = ""
                if profile_source_combo.currentText().strip() != "Stock MP":
                    profile_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0 and profile_source_combo.currentText().strip() != "Stock MP":
                    profile_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                profile_hint.setText(str(payload.get("hint", "") or "").strip())
            elif mode_combo.currentText().strip() == "Tubo":
                if tube_source_combo.currentText().strip() == "Stock MP":
                    record, preview = _current_stock(tube_stock_combo)
                    if record:
                        current_stock_id = str(record.get("id", "") or "").strip()
                        if stock_sync_state["tube"] != current_stock_id:
                            stock_sync_state["tube"] = current_stock_id
                            tube_quality_edit.setText(_quality_from_record(record))
                            tube_meters_spin.setValue(float(record.get("metros", tube_meters_spin.value()) or tube_meters_spin.value() or 0.0))
                            secao_txt = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().lower()
                            for idx in range(tube_section_combo.count()):
                                if str(tube_section_combo.itemData(idx) or "").strip().lower() == secao_txt:
                                    tube_section_combo.setCurrentIndex(idx)
                                    break
                            tube_thickness_spin.setValue(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0))
                            if secao_txt == "redondo":
                                tube_diameter_spin.setValue(float(preview.get("diametro", record.get("diametro", 0)) or 0.0))
                            else:
                                tube_side_a_spin.setValue(float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                                width_value = float(preview.get("largura", record.get("largura", 0)) or 0.0)
                                tube_side_b_spin.setValue(width_value if width_value > 0 else float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                            tube_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                            tube_price_kg_spin.setValue(float(_stock_price_base(record, preview)[1] or 0.0))
                else:
                    stock_sync_state["tube"] = ""
                if tube_source_combo.currentText().strip() != "Stock MP":
                    tube_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                    if float(initial.get("price_base_value", initial.get("price_per_kg", 0)) or 0.0) <= 0 and float(payload.get("price_base_value", 0) or 0.0) > 0:
                        tube_price_kg_spin.setValue(float(payload.get("price_base_value", 0) or 0.0))
                tube_hint.setText(str(payload.get("hint", "") or "").strip())
            elif mode_combo.currentText().strip() == "Chapa":
                if sheet_source_combo.currentText().strip() == "Stock MP":
                    record, preview = _current_stock(sheet_stock_combo)
                    if record:
                        current_stock_id = str(record.get("id", "") or "").strip()
                        if stock_sync_state["sheet"] != current_stock_id:
                            stock_sync_state["sheet"] = current_stock_id
                            family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                            if family_key:
                                _set_combo_by_data(sheet_family_combo, family_key)
                                _refresh_sheet_quality_options()
                            sheet_quality_combo.setCurrentText(_quality_from_record(record))
                            is_checker_stock = bool(str(preview.get("espessura", record.get("espessura", "")) or "").strip().find("/") >= 0) or any(
                                token in backend.norm_text(
                                    " ".join(str(record.get(key, "") or "") for key in ("material", "descricao", "obs", "formato"))
                                )
                                for token in ("gota", "xadrez", "antiderrap")
                            )
                            sheet_kind_combo.setCurrentIndex(1 if is_checker_stock else 0)
                            length_mm_spin.setValue(float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0))
                            width_mm_spin.setValue(float(preview.get("largura", record.get("largura", 0)) or 0.0))
                            thickness_mm_spin.setValue(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0))
                            density_spin.setValue(float(preview.get("densidade", density_spin.value()) or density_spin.value() or 7850.0))
                            sheet_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                else:
                    stock_sync_state["sheet"] = ""
                if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0 and sheet_source_combo.currentText().strip() != "Stock MP":
                    sheet_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                sheet_hint.setText(str(payload.get("hint", "") or "").strip())
            elif mode_combo.currentText().strip() == "Cantoneira":
                if angle_source_combo.currentText().strip() == "Stock MP":
                    record, preview = _current_stock(angle_stock_combo)
                    if record:
                        current_stock_id = str(record.get("id", "") or "").strip()
                        if stock_sync_state["angle"] != current_stock_id:
                            stock_sync_state["angle"] = current_stock_id
                            family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                            if family_key:
                                _set_combo_by_data(angle_family_combo, family_key)
                                _refresh_quality_combo(angle_family_combo, angle_quality_combo)
                            angle_quality_combo.setCurrentText(_quality_from_record(record))
                            angle_meters_spin.setValue(float(record.get("metros", angle_meters_spin.value()) or angle_meters_spin.value() or 0.0))
                            a_val = float(preview.get("comprimento", record.get("comprimento", 0)) or 0.0)
                            if a_val > 0:
                                angle_leg_combo.setCurrentText(str(int(round(a_val))))
                            thickness_txt = _float_text(float(preview.get("espessura_mm", record.get("espessura", 0)) or 0.0), 0)
                            if thickness_txt:
                                angle_thickness_combo.setCurrentText(thickness_txt)
                            angle_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                            angle_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                else:
                    stock_sync_state["angle"] = ""
                if angle_source_combo.currentText().strip() != "Stock MP":
                    angle_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                    if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0:
                        angle_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                angle_hint.setText(str(payload.get("hint", "") or "").strip())
            elif mode_combo.currentText().strip() == "Barra":
                if bar_source_combo.currentText().strip() == "Stock MP":
                    record, preview = _current_stock(bar_stock_combo)
                    if record:
                        current_stock_id = str(record.get("id", "") or "").strip()
                        if stock_sync_state["bar"] != current_stock_id:
                            stock_sync_state["bar"] = current_stock_id
                            family_key = str(record.get("material_familia", "") or preview.get("material_familia_resolved", "") or "").strip()
                            if family_key:
                                _set_combo_by_data(bar_family_combo, family_key)
                                _refresh_quality_combo(bar_family_combo, bar_quality_combo)
                            bar_quality_combo.setCurrentText(_quality_from_record(record))
                            bar_meters_spin.setValue(float(record.get("metros", bar_meters_spin.value()) or bar_meters_spin.value() or 0.0))
                            pair_txt = str(preview.get("dimension_label", "") or "").replace(" mm", "").strip()
                            if pair_txt:
                                bar_size_combo.setCurrentText(pair_txt)
                            bar_kg_per_m_spin.setValue(float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0))
                            bar_price_kg_spin.setValue(float(_stock_price_kg(record, preview) or 0.0))
                else:
                    stock_sync_state["bar"] = ""
                if bar_source_combo.currentText().strip() != "Stock MP":
                    bar_kg_per_m_spin.setValue(float(payload.get("kg_m", 0) or 0.0))
                    if float(initial.get("price_per_kg", 0) or 0.0) <= 0 and float(payload.get("price_kg", 0) or 0.0) > 0:
                        bar_price_kg_spin.setValue(float(payload.get("price_kg", 0) or 0.0))
                bar_hint.setText(str(payload.get("hint", "") or "").strip())
            _apply_auto_texts(payload)
            result = _compute()
            summary_weight.setText(f"Peso total: {float(result.get('weight_total', 0) or 0):.2f} kg")
            summary_cost.setText(f"Custo total: {_fmt_eur(float(result.get('total_cost', 0) or 0))}")
            summary_desc.setText(str((result.get("line") or {}).get("descricao", "") or "").strip())
            _refresh_operation_cost_hint(float((result.get("line") or {}).get("_material_base_preco_unit", 0) or 0))
        finally:
            sync["busy"] = False

    for widget in (
        mode_combo,
        desc_edit,
        ref_edit,
        qty_spin,
        profile_source_combo,
        profile_stock_combo,
        profile_series_combo,
        profile_size_combo,
        profile_meters_spin,
        profile_kg_per_m_spin,
        profile_price_kg_spin,
        tube_source_combo,
        tube_stock_combo,
        tube_quality_edit,
        tube_section_combo,
        tube_model_combo,
        tube_side_a_spin,
        tube_side_b_spin,
        tube_diameter_spin,
        tube_thickness_spin,
        tube_meters_spin,
        tube_kg_per_m_spin,
        tube_price_kg_spin,
        sheet_source_combo,
        sheet_stock_combo,
        sheet_family_combo,
        sheet_quality_combo,
        sheet_kind_combo,
        sheet_model_combo,
        length_mm_spin,
        width_mm_spin,
        thickness_mm_spin,
        density_spin,
        sheet_price_kg_spin,
        angle_source_combo,
        angle_stock_combo,
        angle_family_combo,
        angle_quality_combo,
        angle_leg_combo,
        angle_thickness_combo,
        angle_meters_spin,
        angle_kg_per_m_spin,
        angle_price_kg_spin,
        bar_source_combo,
        bar_stock_combo,
        bar_family_combo,
        bar_quality_combo,
        bar_size_combo,
        bar_meters_spin,
        bar_kg_per_m_spin,
        bar_price_kg_spin,
        rebar_meters_spin,
        diameter_mm_spin,
        rebar_price_kg_spin,
        unit_combo,
        manual_unit_price_spin,
    ):
        if isinstance(widget, QLineEdit):
            widget.textChanged.connect(_refresh)
        elif isinstance(widget, QComboBox):
            widget.currentTextChanged.connect(_refresh)
        else:
            widget.valueChanged.connect(lambda _value: _refresh())

    desc_edit.textEdited.connect(lambda _text: desc_state.__setitem__("manual", True))
    ref_edit.textEdited.connect(lambda _text: ref_state.__setitem__("manual", True))
    profile_series_combo.currentIndexChanged.connect(lambda _idx: _refresh_profile_size_options())
    tube_section_combo.currentIndexChanged.connect(lambda _idx: _refresh_tube_model_options())
    tube_model_combo.currentTextChanged.connect(lambda _text: _apply_tube_model_selection())
    angle_leg_combo.currentTextChanged.connect(lambda _text: _refresh_angle_thickness_options())
    sheet_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_sheet_quality_options(preserve_current=False), _refresh()))
    sheet_kind_combo.currentIndexChanged.connect(lambda _idx: (_refresh_sheet_model_options(), _apply_sheet_model_selection(), _refresh()))
    sheet_model_combo.currentTextChanged.connect(lambda _text: _apply_sheet_model_selection())
    angle_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_quality_combo(angle_family_combo, angle_quality_combo, preserve_current=False), _refresh()))
    bar_family_combo.currentIndexChanged.connect(lambda _idx: (_refresh_quality_combo(bar_family_combo, bar_quality_combo, preserve_current=False), _refresh()))

    def _open_price_table() -> None:
        mode = mode_combo.currentText().strip()
        preferred_id = ""
        if mode == "Perfil":
            preferred_id = str((_current_stock(profile_stock_combo)[0] or {}).get("id", "") or "").strip()
        elif mode == "Tubo":
            preferred_id = str((_current_stock(tube_stock_combo)[0] or {}).get("id", "") or "").strip()
        elif mode == "Chapa":
            preferred_id = str((_current_stock(sheet_stock_combo)[0] or {}).get("id", "") or "").strip()
        elif mode == "Cantoneira":
            preferred_id = str((_current_stock(angle_stock_combo)[0] or {}).get("id", "") or "").strip()
        elif mode == "Barra":
            preferred_id = str((_current_stock(bar_stock_combo)[0] or {}).get("id", "") or "").strip()
        partial(edit_material_prices, owner, backend)(mode, preferred_id, parent=dialog)
        _fill_stock_combo(profile_stock_combo, "Perfil", preferred_id if mode == "Perfil" else "")
        _fill_stock_combo(tube_stock_combo, "Tubo", preferred_id if mode == "Tubo" else "")
        _fill_stock_combo(sheet_stock_combo, "Chapa", preferred_id if mode == "Chapa" else "")
        _fill_stock_combo(angle_stock_combo, "Cantoneira", preferred_id if mode == "Cantoneira" else "")
        _fill_stock_combo(bar_stock_combo, "Barra", preferred_id if mode == "Barra" else "")
        _refresh()

    price_table_btn.clicked.connect(_open_price_table)
    op_detail_btn.clicked.connect(_edit_operation_costs)
    op_profiles_btn.clicked.connect(lambda: (_open_operation_cost_profiles_dialog(dialog, backend), _refresh()))
    operation_edit.textChanged.connect(lambda _text: _refresh())

    _refresh()
    if dialog.exec() != QDialog.Accepted:
        return None
    result = _compute()
    if not isinstance(result.get("line"), dict):
        return None
    mode = mode_combo.currentText().strip()
    if mode == "Perfil":
        stock_record, stock_preview = _current_stock(profile_stock_combo)
        if profile_source_combo.currentText().strip() == "Stock MP" and stock_record:
            updated = partial(sync_stock_price, owner, backend)(
                str(stock_record.get("id", "") or "").strip(),
                float(result.get("price_per_kg", 0) or 0.0),
                _stock_price_kg(stock_record, stock_preview),
                parent=dialog,
            )
            if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["price_base_value"] = float(updated.get("p_compra", result.get("price_base_value", 0)) or 0.0)
                result["line"]["price_base_value"] = float(updated.get("p_compra", result["line"].get("price_base_value", 0)) or 0.0)
                result["price_base_label"] = str(updated.get("base_label", "EUR/kg") or "EUR/kg")
                result["line"]["price_base_label"] = str(updated.get("base_label", "EUR/kg") or "EUR/kg")
    elif mode == "Tubo":
        stock_record, stock_preview = _current_stock(tube_stock_combo)
        if tube_source_combo.currentText().strip() == "Stock MP" and stock_record:
            updated = partial(sync_stock_price, owner, backend)(
                str(stock_record.get("id", "") or "").strip(),
                float(result.get("price_per_kg", 0) or 0.0),
                _stock_price_kg(stock_record, stock_preview),
                parent=dialog,
            )
            if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["price_base_value"] = float(updated.get("p_compra", result.get("price_base_value", 0)) or 0.0)
                result["line"]["price_base_value"] = float(updated.get("p_compra", result["line"].get("price_base_value", 0)) or 0.0)
                result["price_base_label"] = str(updated.get("base_label", "EUR/m") or "EUR/m")
                result["line"]["price_base_label"] = str(updated.get("base_label", "EUR/m") or "EUR/m")
    elif mode == "Chapa":
        stock_record, stock_preview = _current_stock(sheet_stock_combo)
        if sheet_source_combo.currentText().strip() == "Stock MP" and stock_record:
            updated = partial(sync_stock_price, owner, backend)(
                str(stock_record.get("id", "") or "").strip(),
                float(result.get("price_per_kg", 0) or 0.0),
                _stock_price_kg(stock_record, stock_preview),
                parent=dialog,
            )
            if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
    elif mode == "Cantoneira":
        stock_record, stock_preview = _current_stock(angle_stock_combo)
        if angle_source_combo.currentText().strip() == "Stock MP" and stock_record:
            updated = partial(sync_stock_price, owner, backend)(
                str(stock_record.get("id", "") or "").strip(),
                float(result.get("price_per_kg", 0) or 0.0),
                _stock_price_kg(stock_record, stock_preview),
                parent=dialog,
            )
            if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
    elif mode == "Barra":
        stock_record, stock_preview = _current_stock(bar_stock_combo)
        if bar_source_combo.currentText().strip() == "Stock MP" and stock_record:
            updated = partial(sync_stock_price, owner, backend)(
                str(stock_record.get("id", "") or "").strip(),
                float(result.get("price_per_kg", 0) or 0.0),
                _stock_price_kg(stock_record, stock_preview),
                parent=dialog,
            )
            if isinstance(updated, dict) and float(updated.get("price_kg", 0) or 0) > 0:
                result["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
                result["line"]["price_per_kg"] = float(updated.get("price_kg", 0) or 0.0)
    return result
