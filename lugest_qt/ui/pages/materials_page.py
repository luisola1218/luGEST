from __future__ import annotations

import os
from pathlib import Path
import re
import unicodedata

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


_PROFILE_MASS_KG_M: dict[str, dict[str, float]] = {
    "IPE": {
        "80": 6.0,
        "100": 8.1,
        "120": 10.4,
        "140": 12.9,
        "160": 15.8,
        "180": 18.8,
        "200": 22.4,
        "220": 26.2,
        "240": 30.7,
        "270": 36.1,
        "300": 42.2,
    },
    "HEA": {
        "100": 16.7,
        "120": 19.9,
        "140": 24.7,
        "160": 30.4,
        "180": 35.5,
        "200": 42.3,
        "220": 50.5,
        "240": 60.3,
        "260": 68.2,
        "280": 76.4,
        "300": 88.3,
    },
    "HEB": {
        "100": 20.4,
        "120": 26.7,
        "140": 33.7,
        "160": 42.6,
        "180": 51.2,
        "200": 61.3,
        "220": 71.5,
        "240": 83.2,
        "260": 93.0,
        "280": 103.0,
        "300": 117.0,
    },
    "UPN": {
        "80": 8.64,
        "100": 10.6,
        "120": 13.4,
        "140": 16.0,
        "160": 18.8,
        "180": 22.0,
        "200": 25.3,
        "220": 29.4,
        "240": 33.2,
        "260": 37.9,
        "280": 41.8,
        "300": 46.2,
    },
    "UPE": {
        "80": 7.93,
        "100": 9.82,
        "120": 12.1,
        "140": 14.5,
        "160": 17.0,
        "180": 19.7,
        "200": 22.8,
        "220": 26.6,
        "240": 30.2,
        "270": 35.2,
        "300": 40.5,
    },
}


def _grid_search_normalize(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.casefold()
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^a-z0-9./]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _grid_numeric_key(value: object) -> str:
    text = _grid_search_normalize(value).strip().strip("./")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return ""
    try:
        number = float(text)
    except Exception:
        return ""
    return f"{number:.4f}".rstrip("0").rstrip(".")


def _grid_search_terms(value: object) -> list[str]:
    return [term for term in _grid_search_normalize(value).split() if term]


def _grid_search_bucket(values: list[object]) -> tuple[str, set[str]]:
    normalized = _grid_search_normalize(" ".join(str(value or "") for value in values))
    tokens = set(normalized.split())
    expanded = set(tokens)
    for token in tokens:
        numeric = _grid_numeric_key(token)
        if numeric:
            expanded.add(numeric)
        if "/" in token:
            for part in token.split("/"):
                part = part.strip()
                if part:
                    expanded.add(part)
                    part_numeric = _grid_numeric_key(part)
                    if part_numeric:
                        expanded.add(part_numeric)
    return normalized, expanded


def _grid_search_matches(values: list[object], query: object) -> bool:
    terms = _grid_search_terms(query)
    if not terms:
        return True
    normalized, tokens = _grid_search_bucket(values)
    for term in terms:
        numeric = _grid_numeric_key(term)
        if numeric:
            if numeric not in tokens and term not in tokens:
                return False
            continue
        if "/" in term:
            term_parts = [part for part in term.split("/") if part]
            if term_parts and all((_grid_numeric_key(part) or part) in tokens for part in term_parts):
                continue
        if term not in normalized:
            return False
    return True


def _detect_profile_section(text: str) -> tuple[str, str]:
    raw = str(text or "").upper()
    match = re.search(r"\b(IPE|HEA|HEB|UPN|UPE)\s*[- ]?(\d{2,3})\b", raw)
    if not match:
        return "", ""
    return match.group(1), match.group(2)


def _material_family_options(backend) -> list[dict]:
    try:
        return [dict(row or {}) for row in list(backend.material_family_options() or [])]
    except Exception:
        return [
            {"key": "", "label": "Auto", "density": 0.0},
            {"key": "steel", "label": "Aço / Ferro", "density": 7.85},
            {"key": "stainless", "label": "Inox", "density": 7.93},
            {"key": "aluminum", "label": "Alumínio", "density": 2.70},
            {"key": "brass", "label": "Latão", "density": 8.50},
            {"key": "copper", "label": "Cobre", "density": 8.96},
        ]


def _normalise_material_family_key(backend, material: str = "", family: str = "") -> str:
    if not str(family or "").strip():
        return ""
    try:
        return str((backend.material_family_profile(material, family) or {}).get("key", "") or "").strip()
    except Exception:
        return str(family or "").strip()


def _set_material_family_combo(backend, combo: QComboBox, current_value: str = "", *, material: str = "") -> None:
    options = _material_family_options(backend)
    target_key = _normalise_material_family_key(backend, material, current_value)
    combo.blockSignals(True)
    combo.clear()
    for option in options:
        combo.addItem(str(option.get("label", "") or ""), str(option.get("key", "") or ""))
    target_index = 0
    for index in range(combo.count()):
        if str(combo.itemData(index) or "").strip() == target_key:
            target_index = index
            break
    combo.setCurrentIndex(target_index)
    combo.blockSignals(False)


class _SimpleFormDialog(QDialog):
    def __init__(self, title: str, fields: list[tuple[str, str]], defaults: dict[str, str] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(460, 260)
        self.edits: dict[str, QLineEdit] = {}
        layout = QVBoxLayout(self)
        form = QFormLayout()
        defaults = defaults or {}
        for key, label in fields:
            edit = QLineEdit(str(defaults.get(key, "")))
            form.addRow(label, edit)
            self.edits[key] = edit
        layout.addLayout(form)
        actions = QHBoxLayout()
        ok_btn = QPushButton("Guardar")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setProperty("variant", "secondary")
        cancel_btn.clicked.connect(self.reject)
        actions.addWidget(ok_btn)
        actions.addWidget(cancel_btn)
        layout.addLayout(actions)

    def values(self) -> dict[str, str]:
        return {key: edit.text().strip() for key, edit in self.edits.items()}


class _WeightCalculatorDialog(QDialog):
    def __init__(self, defaults: dict[str, float] | None = None, parent=None) -> None:
        super().__init__(parent)
        defaults = defaults or {}
        self._building = True
        self.setWindowTitle("Calculadora de peso")
        self.resize(560, 420)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Chapa", "Tubo", "Perfil", "Cantoneira", "Barra"])
        self.mode_combo.setCurrentText(str(defaults.get("formato", "Chapa") or "Chapa"))
        top_form = QFormLayout()
        top_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        top_form.setFormAlignment(Qt.AlignTop)
        top_form.addRow("Modo", self.mode_combo)
        layout.addLayout(top_form)

        self.form_layout = QFormLayout()
        self.form_layout.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.form_layout.setFormAlignment(Qt.AlignTop)
        self.form_layout.setHorizontalSpacing(16)
        self.form_layout.setVerticalSpacing(8)
        layout.addLayout(self.form_layout, 1)

        self._rows: dict[str, tuple[QLabel, QWidget]] = {}
        self._build_form(defaults)

        self.summary_hint = QLabel("")
        self.summary_hint.setWordWrap(True)
        self.summary_hint.setProperty("role", "muted")
        layout.addWidget(self.summary_hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.mode_combo.currentTextChanged.connect(self._apply_mode)
        self._building = False
        self._apply_mode(self.mode_combo.currentText())
        self._recalc()

    def _new_dim_spin(self, *, decimals: int = 1, max_value: float = 100000.0) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(0.0, max_value)
        spin.setDecimals(decimals)
        spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
        spin.valueChanged.connect(self._recalc)
        return spin

    def _new_value_label(self, text: str, strong: bool = False) -> QLabel:
        label = QLabel(text)
        label.setProperty("role", "field_value_strong" if strong else "field_value")
        return label

    def _add_row(self, key: str, label_text: str, widget: QWidget) -> None:
        label = QLabel(label_text)
        self.form_layout.addRow(label, widget)
        self._rows[key] = (label, widget)

    def _set_row_visible(self, key: str, visible: bool) -> None:
        label, widget = self._rows[key]
        label.setVisible(visible)
        widget.setVisible(visible)

    def _build_form(self, defaults: dict[str, float]) -> None:
        self.length_spin = self._new_dim_spin()
        self.width_spin = self._new_dim_spin()
        self.thickness_spin = self._new_dim_spin()
        self.density_spin = self._new_dim_spin(decimals=3, max_value=20.0)
        self.length_spin.setValue(float(defaults.get("comprimento", 0) or 0))
        self.width_spin.setValue(float(defaults.get("largura", 0) or 0))
        self.thickness_spin.setValue(float(defaults.get("espessura", 0) or 0))
        self.density_spin.setValue(float(defaults.get("densidade", 7.85) or 7.85))
        self.weight_label = self._new_value_label("0,000 kg", strong=True)
        self.area_label = self._new_value_label("0,000 m2")
        self._add_row("sheet_length", "Comprimento (mm)", self.length_spin)
        self._add_row("sheet_width", "Largura (mm)", self.width_spin)
        self._add_row("sheet_thickness", "Espessura (mm)", self.thickness_spin)
        self._add_row("sheet_density", "Densidade (g/cm3)", self.density_spin)
        self._add_row("sheet_weight", "Peso", self.weight_label)
        self._add_row("sheet_area", "Area", self.area_label)

        self.tube_shape_combo = QComboBox()
        self.tube_shape_combo.addItems(["Quadrado/Retangular", "Redondo"])
        tube_text = f"{defaults.get('material_name', '')} {defaults.get('lote', '')}".lower()
        if "redond" in tube_text or "ø" in tube_text or "diam" in tube_text:
            self.tube_shape_combo.setCurrentText("Redondo")
        self.tube_shape_combo.currentTextChanged.connect(self._apply_tube_shape)
        self.tube_length_m_spin = self._new_dim_spin(decimals=3, max_value=1000.0)
        self.tube_length_m_spin.setValue(float(defaults.get("metros", 0) or 0) or 6.0)
        self.tube_width_spin = self._new_dim_spin()
        self.tube_height_spin = self._new_dim_spin()
        self.tube_diameter_spin = self._new_dim_spin()
        self.tube_thickness_spin = self._new_dim_spin()
        self.tube_density_spin = self._new_dim_spin(decimals=3, max_value=20.0)
        self.tube_density_spin.setValue(float(defaults.get("densidade", 7.85) or 7.85))
        self.tube_width_spin.setValue(float(defaults.get("comprimento", 0) or 0))
        self.tube_height_spin.setValue(float(defaults.get("largura", 0) or 0))
        self.tube_diameter_spin.setValue(float(defaults.get("diametro", 0) or 0))
        self.tube_thickness_spin.setValue(float(defaults.get("espessura", 0) or 0))
        self.tube_section_label = self._new_value_label("0,00 mm2")
        self.tube_kgm_label = self._new_value_label("0,000 kg/m")
        self.tube_weight_label = self._new_value_label("0,000 kg", strong=True)
        self._add_row("tube_shape", "Tipo tubo", self.tube_shape_combo)
        self._add_row("tube_length", "Comprimento barra (m)", self.tube_length_m_spin)
        self._add_row("tube_width", "Lado A (mm)", self.tube_width_spin)
        self._add_row("tube_height", "Lado B (mm)", self.tube_height_spin)
        self._add_row("tube_diameter", "Diâmetro ext. (mm)", self.tube_diameter_spin)
        self._add_row("tube_thickness", "Espessura (mm)", self.tube_thickness_spin)
        self._add_row("tube_density", "Densidade (g/cm3)", self.tube_density_spin)
        self._add_row("tube_section", "Secção", self.tube_section_label)
        self._add_row("tube_kgm", "Peso por metro", self.tube_kgm_label)
        self._add_row("tube_total", "Peso total", self.tube_weight_label)

        self.profile_series_combo = QComboBox()
        self.profile_series_combo.addItems(sorted(_PROFILE_MASS_KG_M.keys()))
        self.profile_series_combo.currentTextChanged.connect(self._update_profile_sizes)
        self.profile_size_combo = QComboBox()
        self.profile_size_combo.currentTextChanged.connect(self._sync_profile_kgm_from_lookup)
        self.profile_length_m_spin = self._new_dim_spin(decimals=3, max_value=1000.0)
        guessed_length = float(defaults.get("metros", 0) or 0)
        if guessed_length <= 0:
            guessed_length = float(defaults.get("comprimento", 0) or 0)
            if guessed_length > 100.0:
                guessed_length = guessed_length / 1000.0
        self.profile_length_m_spin.setValue(guessed_length or 6.0)
        self.profile_kgm_spin = self._new_dim_spin(decimals=3, max_value=1000.0)
        self.profile_manual_check = QCheckBox("Permitir kg/m manual")
        self.profile_manual_check.toggled.connect(self._toggle_profile_manual)
        self.profile_weight_m_label = self._new_value_label("0,000 kg/m")
        self.profile_weight_label = self._new_value_label("0,000 kg", strong=True)
        combined_text = f"{defaults.get('material_name', '')} {defaults.get('lote', '')}"
        guessed_series, guessed_size = _detect_profile_section(combined_text)
        if guessed_series and guessed_series in _PROFILE_MASS_KG_M:
            self.profile_series_combo.setCurrentText(guessed_series)
        self._update_profile_sizes(self.profile_series_combo.currentText())
        if guessed_size:
            self.profile_size_combo.setCurrentText(guessed_size)
        else:
            self._sync_profile_kgm_from_lookup()
        self._add_row("profile_series", "Série", self.profile_series_combo)
        self._add_row("profile_size", "Tamanho", self.profile_size_combo)
        self._add_row("profile_length", "Comprimento barra (m)", self.profile_length_m_spin)
        self._add_row("profile_kgm", "Peso por metro (kg/m)", self.profile_kgm_spin)
        self._add_row("profile_manual", "", self.profile_manual_check)
        self._add_row("profile_weight_m", "Peso por metro", self.profile_weight_m_label)
        self._add_row("profile_total", "Peso total", self.profile_weight_label)
        self._toggle_profile_manual(False)

        self.angle_leg_a_spin = self._new_dim_spin()
        self.angle_leg_b_spin = self._new_dim_spin()
        self.angle_thickness_spin = self._new_dim_spin()
        self.angle_length_m_spin = self._new_dim_spin(decimals=3, max_value=1000.0)
        self.angle_density_spin = self._new_dim_spin(decimals=3, max_value=20.0)
        self.angle_density_spin.setValue(float(defaults.get("densidade", 7.85) or 7.85))
        self.angle_length_m_spin.setValue(float(defaults.get("metros", 0) or 0) or 6.0)
        self.angle_leg_a_spin.setValue(float(defaults.get("comprimento", 0) or 0))
        self.angle_leg_b_spin.setValue(float(defaults.get("largura", 0) or 0) or float(defaults.get("comprimento", 0) or 0))
        self.angle_thickness_spin.setValue(float(defaults.get("espessura", 0) or 0))
        self.angle_kgm_label = self._new_value_label("0,000 kg/m")
        self.angle_weight_label = self._new_value_label("0,000 kg", strong=True)
        self._add_row("angle_leg_a", "Aba A (mm)", self.angle_leg_a_spin)
        self._add_row("angle_leg_b", "Aba B (mm)", self.angle_leg_b_spin)
        self._add_row("angle_thickness", "Espessura (mm)", self.angle_thickness_spin)
        self._add_row("angle_length", "Comprimento barra (m)", self.angle_length_m_spin)
        self._add_row("angle_density", "Densidade (g/cm3)", self.angle_density_spin)
        self._add_row("angle_kgm", "Peso por metro", self.angle_kgm_label)
        self._add_row("angle_total", "Peso total", self.angle_weight_label)

        self.bar_side_a_spin = self._new_dim_spin()
        self.bar_side_b_spin = self._new_dim_spin()
        self.bar_length_m_spin = self._new_dim_spin(decimals=3, max_value=1000.0)
        self.bar_density_spin = self._new_dim_spin(decimals=3, max_value=20.0)
        self.bar_density_spin.setValue(float(defaults.get("densidade", 7.85) or 7.85))
        self.bar_length_m_spin.setValue(float(defaults.get("metros", 0) or 0) or 6.0)
        self.bar_side_a_spin.setValue(float(defaults.get("comprimento", 0) or 0))
        self.bar_side_b_spin.setValue(float(defaults.get("largura", 0) or defaults.get("espessura", 0) or 0))
        self.bar_kgm_label = self._new_value_label("0,000 kg/m")
        self.bar_weight_label = self._new_value_label("0,000 kg", strong=True)
        self._add_row("bar_side_a", "Lado A (mm)", self.bar_side_a_spin)
        self._add_row("bar_side_b", "Lado B (mm)", self.bar_side_b_spin)
        self._add_row("bar_length", "Comprimento barra (m)", self.bar_length_m_spin)
        self._add_row("bar_density", "Densidade (g/cm3)", self.bar_density_spin)
        self._add_row("bar_kgm", "Peso por metro", self.bar_kgm_label)
        self._add_row("bar_total", "Peso total", self.bar_weight_label)

    def _apply_mode(self, mode: str) -> None:
        mode = str(mode or "Chapa").strip()
        show_sheet = mode == "Chapa"
        show_tube = mode == "Tubo"
        show_profile = mode == "Perfil"
        show_angle = mode == "Cantoneira"
        show_bar = mode == "Barra"
        for key in ("sheet_length", "sheet_width", "sheet_thickness", "sheet_density", "sheet_weight", "sheet_area"):
            self._set_row_visible(key, show_sheet)
        for key in ("tube_shape", "tube_length", "tube_thickness", "tube_density", "tube_section", "tube_kgm", "tube_total"):
            self._set_row_visible(key, show_tube)
        for key in ("profile_series", "profile_size", "profile_length", "profile_kgm", "profile_manual", "profile_weight_m", "profile_total"):
            self._set_row_visible(key, show_profile)
        for key in ("angle_leg_a", "angle_leg_b", "angle_thickness", "angle_length", "angle_density", "angle_kgm", "angle_total"):
            self._set_row_visible(key, show_angle)
        for key in ("bar_side_a", "bar_side_b", "bar_length", "bar_density", "bar_kgm", "bar_total"):
            self._set_row_visible(key, show_bar)
        self._apply_tube_shape(self.tube_shape_combo.currentText())
        self._recalc()

    def _apply_tube_shape(self, shape: str) -> None:
        mode = self.mode_combo.currentText().strip()
        is_round = str(shape or "").startswith("Redondo")
        self._set_row_visible("tube_width", mode == "Tubo" and not is_round)
        self._set_row_visible("tube_height", mode == "Tubo" and not is_round)
        self._set_row_visible("tube_diameter", mode == "Tubo" and is_round)
        self._recalc()

    def _update_profile_sizes(self, series: str) -> None:
        self.profile_size_combo.blockSignals(True)
        self.profile_size_combo.clear()
        self.profile_size_combo.addItems(sorted(_PROFILE_MASS_KG_M.get(str(series or ""), {}).keys(), key=lambda x: int(x)))
        self.profile_size_combo.blockSignals(False)
        self._sync_profile_kgm_from_lookup()

    def _sync_profile_kgm_from_lookup(self) -> None:
        if self.profile_manual_check.isChecked():
            self._recalc()
            return
        series = self.profile_series_combo.currentText().strip()
        size = self.profile_size_combo.currentText().strip()
        value = float(_PROFILE_MASS_KG_M.get(series, {}).get(size, 0.0) or 0.0)
        self.profile_kgm_spin.blockSignals(True)
        self.profile_kgm_spin.setValue(value)
        self.profile_kgm_spin.blockSignals(False)
        self._recalc()

    def _toggle_profile_manual(self, enabled: bool) -> None:
        self.profile_kgm_spin.setReadOnly(not enabled)
        self.profile_kgm_spin.setButtonSymbols(QDoubleSpinBox.UpDownArrows if enabled else QDoubleSpinBox.NoButtons)
        if not enabled:
            self._sync_profile_kgm_from_lookup()

    def _recalc(self) -> None:
        if getattr(self, "_building", False):
            return
        mode = self.mode_combo.currentText().strip()
        if mode == "Tubo":
            density = float(self.tube_density_spin.value() or 0)
            thickness = float(self.tube_thickness_spin.value() or 0)
            length_m = float(self.tube_length_m_spin.value() or 0)
            if self.tube_shape_combo.currentText().startswith("Redondo"):
                diameter = float(self.tube_diameter_spin.value() or 0)
                inner = max(0.0, diameter - (2.0 * thickness))
                area_mm2 = max(0.0, 3.141592653589793 * ((diameter ** 2) - (inner ** 2)) / 4.0)
            else:
                width = float(self.tube_width_spin.value() or 0)
                height = float(self.tube_height_spin.value() or 0)
                inner_w = max(0.0, width - (2.0 * thickness))
                inner_h = max(0.0, height - (2.0 * thickness))
                area_mm2 = max(0.0, (width * height) - (inner_w * inner_h))
            kg_m = area_mm2 * density / 1000.0
            total = kg_m * length_m
            self.tube_section_label.setText(f"{area_mm2:,.2f} mm2".replace(",", "X").replace(".", ",").replace("X", "."))
            self.tube_kgm_label.setText(f"{kg_m:,.3f} kg/m".replace(",", "X").replace(".", ",").replace("X", "."))
            self.tube_weight_label.setText(f"{total:,.3f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
            self.summary_hint.setText("Tubo por geometria: secção metálica × densidade × comprimento.")
            return
        if mode == "Perfil":
            length_m = float(self.profile_length_m_spin.value() or 0)
            kg_m = float(self.profile_kgm_spin.value() or 0)
            total = kg_m * length_m
            self.profile_weight_m_label.setText(f"{kg_m:,.3f} kg/m".replace(",", "X").replace(".", ",").replace("X", "."))
            self.profile_weight_label.setText(f"{total:,.3f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
            self.summary_hint.setText("Perfil por tabela standard: kg/m da série escolhida × comprimento da barra.")
            return
        if mode == "Cantoneira":
            a_mm = float(self.angle_leg_a_spin.value() or 0)
            b_mm = float(self.angle_leg_b_spin.value() or 0)
            thickness = float(self.angle_thickness_spin.value() or 0)
            density = float(self.angle_density_spin.value() or 0)
            length_m = float(self.angle_length_m_spin.value() or 0)
            area_mm2 = max(0.0, thickness * ((a_mm + b_mm) - thickness))
            kg_m = area_mm2 * density / 1000.0
            total = kg_m * length_m
            self.angle_kgm_label.setText(f"{kg_m:,.3f} kg/m".replace(",", "X").replace(".", ",").replace("X", "."))
            self.angle_weight_label.setText(f"{total:,.3f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
            self.summary_hint.setText("Cantoneira: área aproximada t × (a + b - t) × densidade × comprimento.")
            return
        if mode == "Barra":
            side_a = float(self.bar_side_a_spin.value() or 0)
            side_b = float(self.bar_side_b_spin.value() or 0)
            density = float(self.bar_density_spin.value() or 0)
            length_m = float(self.bar_length_m_spin.value() or 0)
            area_mm2 = max(0.0, side_a * side_b)
            kg_m = area_mm2 * density / 1000.0
            total = kg_m * length_m
            self.bar_kgm_label.setText(f"{kg_m:,.3f} kg/m".replace(",", "X").replace(".", ",").replace("X", "."))
            self.bar_weight_label.setText(f"{total:,.3f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
            self.summary_hint.setText("Barra maciça: lado A × lado B × densidade × comprimento.")
            return
        length = float(self.length_spin.value() or 0)
        width = float(self.width_spin.value() or 0)
        thickness = float(self.thickness_spin.value() or 0)
        density = float(self.density_spin.value() or 0)
        weight = (length * width * thickness * density) / 1000000.0
        area = (length * width) / 1000000.0
        self.weight_label.setText(f"{weight:.3f} kg".replace(".", ","))
        self.area_label.setText(f"{area:.3f} m2".replace(".", ","))
        self.summary_hint.setText("Chapa por área: comprimento × largura × espessura × densidade.")

    def values(self) -> dict[str, float]:
        mode = self.mode_combo.currentText().strip()
        if mode == "Tubo":
            density = float(self.tube_density_spin.value() or 0)
            thickness = float(self.tube_thickness_spin.value() or 0)
            length_m = float(self.tube_length_m_spin.value() or 0)
            if self.tube_shape_combo.currentText().startswith("Redondo"):
                diameter = float(self.tube_diameter_spin.value() or 0)
                inner = max(0.0, diameter - (2.0 * thickness))
                area_mm2 = max(0.0, 3.141592653589793 * ((diameter ** 2) - (inner ** 2)) / 4.0)
            else:
                width = float(self.tube_width_spin.value() or 0)
                height = float(self.tube_height_spin.value() or 0)
                inner_w = max(0.0, width - (2.0 * thickness))
                inner_h = max(0.0, height - (2.0 * thickness))
                area_mm2 = max(0.0, (width * height) - (inner_w * inner_h))
            kg_m = area_mm2 * density / 1000.0
            return {"mode": mode, "peso_unid": round(kg_m * length_m, 4), "metros": round(length_m, 4), "kg_m": round(kg_m, 4)}
        if mode == "Perfil":
            length_m = float(self.profile_length_m_spin.value() or 0)
            kg_m = float(self.profile_kgm_spin.value() or 0)
            return {"mode": mode, "peso_unid": round(kg_m * length_m, 4), "metros": round(length_m, 4), "kg_m": round(kg_m, 4)}
        if mode == "Cantoneira":
            a_mm = float(self.angle_leg_a_spin.value() or 0)
            b_mm = float(self.angle_leg_b_spin.value() or 0)
            thickness = float(self.angle_thickness_spin.value() or 0)
            density = float(self.angle_density_spin.value() or 0)
            length_m = float(self.angle_length_m_spin.value() or 0)
            kg_m = max(0.0, thickness * ((a_mm + b_mm) - thickness)) * density / 1000.0
            return {"mode": mode, "peso_unid": round(kg_m * length_m, 4), "metros": round(length_m, 4), "kg_m": round(kg_m, 4)}
        if mode == "Barra":
            side_a = float(self.bar_side_a_spin.value() or 0)
            side_b = float(self.bar_side_b_spin.value() or 0)
            density = float(self.bar_density_spin.value() or 0)
            length_m = float(self.bar_length_m_spin.value() or 0)
            kg_m = max(0.0, side_a * side_b) * density / 1000.0
            return {"mode": mode, "peso_unid": round(kg_m * length_m, 4), "metros": round(length_m, 4), "kg_m": round(kg_m, 4)}
        length = float(self.length_spin.value() or 0)
        width = float(self.width_spin.value() or 0)
        thickness = float(self.thickness_spin.value() or 0)
        density = float(self.density_spin.value() or 0)
        return {"mode": mode, "peso_unid": round((length * width * thickness * density) / 1000000.0, 4), "metros": round((length * width) / 1000000.0, 4), "kg_m": 0.0}


class _HistoryDialog(QDialog):
    def __init__(self, title: str, rows: list[dict[str, str]], backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.dialog_title = title
        self.all_rows = list(rows or [])
        self.filtered_rows = list(self.all_rows)
        self.setWindowTitle(title)
        self.resize(1380, 720)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        intro = CardFrame()
        intro.set_tone("info")
        intro_layout = QVBoxLayout(intro)
        intro_layout.setContentsMargins(14, 12, 14, 12)
        intro_layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel("Consulta detalhada dos movimentos do stock com leitura limpa, filtro rápido e saída para impressão.")
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        intro_layout.addWidget(title_label)
        intro_layout.addWidget(subtitle)
        layout.addWidget(intro)

        toolbar_card = CardFrame()
        toolbar_card.set_tone("default")
        toolbar = QHBoxLayout(toolbar_card)
        toolbar.setContentsMargins(12, 10, 12, 10)
        toolbar.setSpacing(8)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filtrar por ação, operador, matéria-prima, espessura, lote ou detalhe...")
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.count_label = QLabel("0 registos")
        self.count_label.setProperty("role", "muted")
        self.pdf_btn = QPushButton("Abrir PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_pdf)
        self.print_btn = QPushButton("Imprimir")
        self.print_btn.setProperty("variant", "secondary")
        self.print_btn.clicked.connect(self._print_pdf)
        toolbar.addWidget(self.filter_edit, 1)
        toolbar.addWidget(self.count_label)
        toolbar.addWidget(self.pdf_btn)
        toolbar.addWidget(self.print_btn)
        layout.addWidget(toolbar_card)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(["Data", "Ação", "Operador", "Matéria-prima", "Esp.", "Dim.", "Lote", "Qtd", "Reserv.", "Detalhes"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        header = self.table.horizontalHeader()
        for col, width in ((0, 138), (1, 92), (2, 102), (3, 220), (4, 60), (5, 96), (6, 150), (7, 64), (8, 72), (9, 440)):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            header.resizeSection(col, width)
        header.setSectionResizeMode(9, QHeaderView.Stretch)
        layout.addWidget(self.table)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)
        self._apply_filter()

    def _apply_filter(self) -> None:
        query = self.filter_edit.text().strip().lower()
        if query:
            self.filtered_rows = [
                row for row in self.all_rows if query in " ".join(str(value or "") for value in row.values()).lower()
            ]
        else:
            self.filtered_rows = list(self.all_rows)
        self._set_rows(self.filtered_rows)

    def _set_rows(self, rows: list[dict[str, str]]) -> None:
        self.count_label.setText(f"{len(rows)} registos")
        self.table.setRowCount(len(rows))
        keys = ("data", "acao", "operador", "material", "espessura", "dimensao", "lote", "qtd", "reservado", "detalhes")
        for row_index, row in enumerate(rows):
            for col_index, key in enumerate(keys):
                item = QTableWidgetItem(str(row.get(key, "") or ""))
                if col_index in (0, 1, 2, 4, 7, 8):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                item.setToolTip(item.text())
                self.table.setItem(row_index, col_index, item)

    def _render_pdf(self) -> Path:
        target = Path(os.path.join(Path.home(), "AppData", "Local", "Temp", "lugest_material_history_dialog.pdf"))
        self.backend.material_render_history_pdf(self.filtered_rows, self.dialog_title, target)
        return target

    def _open_pdf(self) -> None:
        try:
            path = self._render_pdf()
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Histórico", str(exc))

    def _print_pdf(self) -> None:
        try:
            path = self._render_pdf()
            try:
                os.startfile(str(path), "print")
            except Exception:
                os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Histórico", str(exc))


class _MaterialEditorDialog(QDialog):
    def __init__(self, backend, parent=None, record: dict | None = None, mode: str = "add") -> None:
        super().__init__(parent)
        self.backend = backend
        self._record = dict(record or {})
        self._mode = str(mode or "add").strip().lower()
        editing = self._mode == "edit"
        self.setWindowTitle("Editar material" if editing else "Adicionar material")
        self.setModal(True)
        self.resize(980, 360)
        self.setStyleSheet(
            "QDialog { font-size: 12px; }"
            " QLabel { font-size: 12px; }"
            " QLineEdit, QComboBox { min-height: 30px; padding: 0 8px; }"
            " QPushButton { min-height: 34px; }"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        intro = QLabel(
            "Confirma os dados do registo selecionado e guarda apenas no fim."
            if editing
            else "Novo registo de matéria-prima. O formulário abre sempre limpo para "
            "evitar herdar o material atualmente selecionado."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)

        form_card = CardFrame()
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(16, 14, 16, 14)
        form_layout.setSpacing(12)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.formato_combo = self._make_combo()
        self.material_combo = self._make_combo()
        self.material_family_combo = QComboBox()
        self.secao_tipo_combo = self._make_combo()
        self.espessura_combo = self._make_combo()
        self.local_combo = self._make_combo()
        self.lote_interno_edit = QLineEdit()
        self.lote_edit = QLineEdit()
        self.comprimento_edit = QLineEdit()
        self.largura_edit = QLineEdit()
        self.altura_edit = QLineEdit()
        self.diametro_edit = QLineEdit()
        self.contorno_edit = QLineEdit()
        self.metros_edit = QLineEdit()
        self.kg_m_edit = QLineEdit()
        self.peso_edit = QLineEdit()
        self.preco_compra_edit = QLineEdit()
        self.preco_unit_edit = QLineEdit()
        self.quantidade_edit = QLineEdit()
        self.reservado_edit = QLineEdit()
        self.preco_unit_edit.setReadOnly(True)
        self.preco_unit_edit.setFocusPolicy(Qt.NoFocus)
        self.preco_unit_edit.setPlaceholderText("0,00 EUR")
        self.peso_edit.setReadOnly(True)
        self.lote_interno_edit.setReadOnly(True)
        self.lote_interno_edit.setPlaceholderText("Gerado automaticamente")
        self.preco_compra_edit.setPlaceholderText("EUR/kg ou EUR/m")
        self.contorno_edit.setPlaceholderText("Opcional: 0,0; 1000,0; 900,400; 0,400")
        self.kg_m_edit.setPlaceholderText("Auto por tabela / fórmula")

        for combo, placeholder in (
            (self.formato_combo, "Selecionar formato"),
            (self.material_combo, "Selecionar / escrever material"),
            (self.secao_tipo_combo, "Selecionar tipo / série"),
            (self.espessura_combo, "Selecionar / escrever espessura"),
            (self.local_combo, "Selecionar / escrever local"),
        ):
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.setPlaceholderText(placeholder)

        fields = [
            ("Formato", self.formato_combo),
            ("Material", self.material_combo),
            ("Família", self.material_family_combo),
            ("Tipo secção", self.secao_tipo_combo),
            ("Espessura", self.espessura_combo),
            ("Lote interno", self.lote_interno_edit),
            ("Lote fornecedor", self.lote_edit),
            ("Comprimento", self.comprimento_edit),
            ("Largura", self.largura_edit),
            ("Altura", self.altura_edit),
            ("Diâmetro", self.diametro_edit),
            ("Contorno retalho", self.contorno_edit),
            ("Metros", self.metros_edit),
            ("Kg/m", self.kg_m_edit),
            ("Peso/Un.", self.peso_edit),
            ("Compra (EUR/kg|m)", self.preco_compra_edit),
            ("Preço/Unid.", self.preco_unit_edit),
            ("Quantidade", self.quantidade_edit),
            ("Reservado", self.reservado_edit),
            ("Localização", self.local_combo),
        ]
        self._field_labels: dict[str, QLabel] = {}
        self._field_widgets: dict[str, QWidget] = {}
        for index, (label_text, widget) in enumerate(fields):
            row = index // 4
            col = (index % 4) * 2
            label = QLabel(label_text)
            label.setProperty("role", "muted")
            self._field_labels[label_text] = label
            self._field_widgets[label_text] = widget
            grid.addWidget(label, row, col)
            grid.addWidget(widget, row, col + 1)
        form_layout.addLayout(grid)
        layout.addWidget(form_card)

        actions = QHBoxLayout()
        calc_btn = QPushButton("Calc. peso")
        calc_btn.setProperty("variant", "secondary")
        calc_btn.clicked.connect(self._open_weight_calculator)
        save_btn = QPushButton("Guardar alterações" if editing else "Adicionar")
        save_btn.clicked.connect(self._accept_if_valid)
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setProperty("variant", "secondary")
        cancel_btn.clicked.connect(self.reject)
        actions.addWidget(calc_btn)
        actions.addStretch(1)
        actions.addWidget(save_btn)
        actions.addWidget(cancel_btn)
        layout.addLayout(actions)

        for combo in (self.formato_combo, self.material_combo, self.secao_tipo_combo, self.espessura_combo, self.local_combo):
            combo.currentTextChanged.connect(self._on_form_value_changed)
        self.material_family_combo.currentIndexChanged.connect(self._on_form_value_changed)
        for edit in (
            self.lote_edit,
            self.lote_interno_edit,
            self.comprimento_edit,
            self.largura_edit,
            self.altura_edit,
            self.diametro_edit,
            self.contorno_edit,
            self.metros_edit,
            self.kg_m_edit,
            self.peso_edit,
            self.preco_compra_edit,
            self.quantidade_edit,
            self.reservado_edit,
        ):
            edit.textChanged.connect(self._on_form_value_changed)

        self._load_presets()
        self._set_form_defaults()
        if self._record:
            self._load_record(self._record)

    def _make_combo(self) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMinimumContentsLength(10)
        return combo

    def _set_combo_values(self, combo: QComboBox, values: list[str], current_text: str = "") -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(values)
        combo.setCurrentText(current_text)
        combo.blockSignals(False)

    def _set_section_options(self, formato: str, current_value: str = "") -> None:
        options = [
            str(row.get("label", "") or "").strip()
            for row in list(self.backend.material_section_options(formato) or [])
            if str(row.get("label", "") or "").strip()
        ]
        self._set_combo_values(self.secao_tipo_combo, options, current_value)

    def _set_field_visible(self, key: str, visible: bool) -> None:
        label = self._field_labels.get(key)
        widget = self._field_widgets.get(key)
        if label is not None:
            label.setVisible(visible)
        if widget is not None:
            widget.setVisible(visible)

    def _set_section_options(self, formato: str, current_value: str = "") -> None:
        options = [
            str(row.get("label", "") or "").strip()
            for row in list(self.backend.material_section_options(formato) or [])
            if str(row.get("label", "") or "").strip()
        ]
        self._set_combo_values(self.secao_tipo_combo, options, current_value)

    def _set_field_visible(self, key: str, visible: bool) -> None:
        label = self._field_labels.get(key)
        widget = self._field_widgets.get(key)
        if label is not None:
            label.setVisible(visible)
        if widget is not None:
            widget.setVisible(visible)

    def _set_section_options(self, formato: str, current_value: str = "") -> None:
        options = [str(row.get("label", "") or "").strip() for row in list(self.backend.material_section_options(formato) or []) if str(row.get("label", "") or "").strip()]
        self._set_combo_values(self.secao_tipo_combo, options, current_value)

    def _set_field_visible(self, key: str, visible: bool) -> None:
        label = self._field_labels.get(key)
        widget = self._field_widgets.get(key)
        if label is not None:
            label.setVisible(visible)
        if widget is not None:
            widget.setVisible(visible)

    def _load_presets(self) -> None:
        presets = self.backend.material_presets()
        current_values = {
            "formato": self.formato_combo.currentText(),
            "material": self.material_combo.currentText(),
            "material_familia": str(self.material_family_combo.currentData() or "").strip(),
            "secao_tipo": self.secao_tipo_combo.currentText(),
            "espessura": self.espessura_combo.currentText(),
            "local": self.local_combo.currentText(),
        }
        self._set_combo_values(self.formato_combo, presets["formatos"], current_values["formato"] or "Chapa")
        self._set_combo_values(self.material_combo, presets["materiais"], current_values["material"])
        _set_material_family_combo(self.backend, self.material_family_combo, current_values["material_familia"], material=current_values["material"])
        self._set_section_options(current_values["formato"] or "Chapa", current_values["secao_tipo"])
        self._set_combo_values(self.espessura_combo, presets["espessuras"], current_values["espessura"])
        self._set_combo_values(self.local_combo, presets["locais"], current_values["local"])

    def _set_form_defaults(self) -> None:
        self.formato_combo.setCurrentText("Chapa")
        self.material_combo.setCurrentText("")
        _set_material_family_combo(self.backend, self.material_family_combo, "")
        self._set_section_options("Chapa", "")
        self.espessura_combo.setCurrentText("")
        self.local_combo.setCurrentText("")
        for edit in (
            self.lote_edit,
            self.lote_interno_edit,
            self.comprimento_edit,
            self.largura_edit,
            self.altura_edit,
            self.diametro_edit,
            self.contorno_edit,
            self.metros_edit,
            self.kg_m_edit,
            self.peso_edit,
            self.preco_compra_edit,
            self.quantidade_edit,
            self.reservado_edit,
        ):
            edit.clear()
        self.preco_unit_edit.setText("0,00 EUR")
        self.reservado_edit.setText("0")
        self._refresh_form_state()
        self._refresh_price_preview()

    def _load_record(self, record: dict[str, object]) -> None:
        preview = self.backend.material_geometry_preview(record)
        formato = str(record.get("formato", "Chapa") or "Chapa")
        material = str(record.get("material", "") or "").strip()
        familia = str(record.get("material_familia", "") or "").strip()
        self.formato_combo.setCurrentText(formato)
        self.material_combo.setCurrentText(material)
        _set_material_family_combo(self.backend, self.material_family_combo, familia, material=material)
        self._set_section_options(formato, str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip())
        self.secao_tipo_combo.setCurrentText(str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip())
        self.espessura_combo.setCurrentText(str(record.get("espessura", "") or ""))
        self.lote_edit.setText(str(record.get("lote_fornecedor", "") or ""))
        self.lote_interno_edit.setText(str(record.get("lote_interno", "") or ""))
        self.comprimento_edit.setText(self.backend._fmt(preview.get("comprimento", record.get("comprimento", 0))))
        self.largura_edit.setText(self.backend._fmt(preview.get("largura", record.get("largura", 0))))
        self.altura_edit.setText(self.backend._fmt(preview.get("altura", record.get("altura", 0))))
        self.diametro_edit.setText(self.backend._fmt(preview.get("diametro", record.get("diametro", 0))))
        self.contorno_edit.setText(self.backend.format_material_contour_points(record.get("contorno_points", record.get("shape_points", []))))
        self.metros_edit.setText(self.backend._fmt(preview.get("metros", record.get("metros", 0))))
        self.kg_m_edit.setText(self.backend._fmt(preview.get("kg_m", record.get("kg_m", 0))))
        self.peso_edit.setText(self.backend._fmt(preview.get("peso_unid", record.get("peso_unid", 0))))
        self.preco_compra_edit.setText(self.backend._fmt(record.get("p_compra", 0)))
        self.quantidade_edit.setText(self.backend._fmt(record.get("quantidade", 0)))
        self.reservado_edit.setText(self.backend._fmt(record.get("reservado", 0)))
        self.local_combo.setCurrentText(self.backend._localizacao(record))
        self._refresh_form_state()
        self._refresh_price_preview()

    def payload(self) -> dict[str, str]:
        return {
            "formato": self.formato_combo.currentText().strip(),
            "material": self.material_combo.currentText().strip(),
            "material_familia": str(self.material_family_combo.currentData() or "").strip(),
            "secao_tipo": self.secao_tipo_combo.currentText().strip(),
            "espessura": self.espessura_combo.currentText().strip(),
            "comprimento": self.comprimento_edit.text().strip(),
            "largura": self.largura_edit.text().strip(),
            "altura": self.altura_edit.text().strip(),
            "diametro": self.diametro_edit.text().strip(),
            "contorno_points": self.contorno_edit.text().strip(),
            "metros": self.metros_edit.text().strip(),
            "kg_m": self.kg_m_edit.text().strip(),
            "peso_unid": self.peso_edit.text().strip(),
            "p_compra": self.preco_compra_edit.text().strip(),
            "quantidade": self.quantidade_edit.text().strip(),
            "reservado": self.reservado_edit.text().strip(),
            "local": self.local_combo.currentText().strip(),
            "lote_interno": self.lote_interno_edit.text().strip(),
            "lote_fornecedor": self.lote_edit.text().strip(),
        }

    def _on_form_value_changed(self, *_args) -> None:
        self._refresh_form_state()
        self._refresh_price_preview()

    def _refresh_form_state(self) -> None:
        formato = (self.formato_combo.currentText().strip() or "Chapa").title()
        self._set_section_options(formato, self.secao_tipo_combo.currentText().strip())
        preview = self.backend.material_geometry_preview(self.payload())
        secao_tipo = str(preview.get("secao_tipo", "") or "").strip()
        tube_round = formato == "Tubo" and secao_tipo == "redondo"
        profile_catalog = formato == "Perfil" and bool(preview.get("usa_catalogo"))
        espessura_required = formato in {"Chapa", "Tubo", "Cantoneira"}
        esp_label = self._field_labels.get("Espessura")
        if esp_label is not None:
            esp_label.setText("Espessura" if espessura_required else "Espessura (opc.)")
        secao_label = self._field_labels.get("Tipo secção")
        if secao_label is not None:
            if formato == "Tubo":
                secao_label.setText("Tipo tubo")
            elif formato == "Perfil":
                secao_label.setText("Tipo perfil / série")
            elif formato == "Cantoneira":
                secao_label.setText("Tipo cantoneira")
            elif formato == "Barra":
                secao_label.setText("Tipo barra")
            else:
                secao_label.setText("Tipo secção")
        comp_label = self._field_labels.get("Comprimento")
        if comp_label is not None:
            if formato == "Chapa":
                comp_label.setText("Comprimento (mm)")
            elif formato == "Tubo":
                comp_label.setText("Lado A (mm)")
            elif formato == "Cantoneira":
                comp_label.setText("Aba A (mm)")
            elif formato == "Barra":
                comp_label.setText("Lado A (mm)")
            else:
                comp_label.setText("Comprimento")
        larg_label = self._field_labels.get("Largura")
        if larg_label is not None:
            if formato == "Chapa":
                larg_label.setText("Largura (mm)")
            elif formato == "Cantoneira":
                larg_label.setText("Aba B (mm)")
            else:
                larg_label.setText("Lado B (mm)")
        altura_label = self._field_labels.get("Altura")
        if altura_label is not None:
            altura_label.setText("Altura / tamanho (mm)")
        diametro_label = self._field_labels.get("Diâmetro")
        if diametro_label is not None:
            diametro_label.setText("Diâmetro ext. (mm)")
        kgm_label = self._field_labels.get("Kg/m")
        if kgm_label is not None:
            kgm_label.setText("Peso por metro (kg/m)")
        compra_label = self._field_labels.get("Compra (EUR/kg|m)")
        if compra_label is not None:
            compra_label.setText("Compra (EUR/m)" if formato == "Tubo" else "Compra (EUR/kg)")
        esp_line = self.espessura_combo.lineEdit()
        if esp_line is not None:
            esp_line.setPlaceholderText("Obrigatória" if espessura_required else "Opcional / derivada")
        self.espessura_combo.setToolTip("Obrigatória para chapa, tubo e cantoneira." if espessura_required else "Opcional quando a secção já define a espessura.")
        self.preco_compra_edit.setPlaceholderText("EUR/m" if formato == "Tubo" else "EUR/kg")
        self.preco_compra_edit.setToolTip("Preço de compra base por metro." if formato == "Tubo" else "Preço de compra base por kg.")
        self.metros_edit.setPlaceholderText("Comprimento barra (m)" if formato in {"Tubo", "Perfil", "Cantoneira", "Barra"} else "")
        self.peso_edit.setPlaceholderText("Peso calculado automaticamente")
        self.kg_m_edit.setReadOnly(formato == "Tubo" or profile_catalog)
        self._set_field_visible("Tipo secção", formato in {"Tubo", "Perfil", "Cantoneira", "Barra"})
        self._set_field_visible("Comprimento", formato in {"Chapa", "Cantoneira", "Barra"} or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Largura", formato in {"Chapa", "Cantoneira", "Barra"} or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Altura", formato == "Perfil")
        self._set_field_visible("Diâmetro", formato == "Tubo" and tube_round)
        self._set_field_visible("Contorno retalho", formato == "Chapa")
        self._set_field_visible("Metros", formato != "Chapa")
        self._set_field_visible("Kg/m", formato in {"Tubo", "Perfil", "Cantoneira", "Barra"})

    def _refresh_price_preview(self) -> None:
        try:
            preview = self.backend.material_price_preview(self.payload())
        except Exception:
            self.preco_unit_edit.setText("0,00 EUR")
            self.preco_unit_edit.setToolTip("")
            return
        self.peso_edit.blockSignals(True)
        self.peso_edit.setText(self.backend._fmt(preview.get("peso_unid", 0)))
        self.peso_edit.blockSignals(False)
        if str(preview.get("formato", "") or "") in {"Tubo", "Perfil", "Cantoneira", "Barra"}:
            self.kg_m_edit.blockSignals(True)
            self.kg_m_edit.setText(self.backend._fmt(preview.get("kg_m", 0)))
            self.kg_m_edit.blockSignals(False)
        preco_unid = float(preview.get("preco_unid", 0.0) or 0.0)
        self.preco_unit_edit.setText(f"{preco_unid:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."))
        tooltip_bits = [
            f"Formato: {preview.get('formato', '-')}",
            f"Secção: {preview.get('secao_label', '-')}",
            f"Dimensão: {preview.get('dimension_label', '-')}",
            f"Família: {preview.get('material_familia_label', 'Aço / Ferro')}",
            f"Densidade: {self.backend._fmt(preview.get('densidade', 0))} g/cm3",
            f"Base: {float(preview.get('p_compra', 0.0) or 0.0):.4f} {preview.get('base_label', 'EUR/kg')}",
        ]
        if str(preview.get("formato", "") or "") == "Tubo":
            tooltip_bits.append(f"Metros/unid.: {self.backend._fmt(preview.get('metros', 0))}")
        else:
            tooltip_bits.append(f"Peso/unid.: {self.backend._fmt(preview.get('peso_unid', 0))} kg")
        if float(preview.get("kg_m", 0) or 0) > 0:
            tooltip_bits.append(f"Kg/m: {self.backend._fmt(preview.get('kg_m', 0))}")
        if str(preview.get("calc_hint", "") or "").strip():
            tooltip_bits.append(str(preview.get("calc_hint", "") or "").strip())
        self.preco_unit_edit.setToolTip(" | ".join(tooltip_bits))

    def _open_weight_calculator(self) -> None:
        formato = self.formato_combo.currentText().strip() or "Chapa"
        profile = self.backend.material_family_profile(
            self.material_combo.currentText().strip(),
            str(self.material_family_combo.currentData() or "").strip(),
        )
        dialog = _WeightCalculatorDialog(
            {
                "formato": formato,
                "comprimento": float(self.backend._parse_float(self.comprimento_edit.text(), 0)),
                "largura": float(self.backend._parse_float(self.largura_edit.text(), 0)),
                "espessura": float(self.backend._parse_float(self.espessura_combo.currentText(), 0)),
                "metros": float(self.backend._parse_float(self.metros_edit.text(), 0)),
                "diametro": float(self.backend._parse_float(self.diametro_edit.text(), 0)),
                "densidade": float(profile.get("density", 7.85) or 7.85),
                "material_name": self.material_combo.currentText().strip(),
                "lote": self.lote_edit.text().strip(),
            },
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        self.peso_edit.setText(self.backend._fmt(values.get("peso_unid", 0)))
        if str(values.get("mode", "") or "") in {"Tubo", "Perfil"}:
            self.metros_edit.setText(self.backend._fmt(values.get("metros", 0)))
        self._refresh_price_preview()

    def _accept_if_valid(self) -> None:
        try:
            self.backend._normalise_material_payload(self.payload())
        except Exception as exc:
            QMessageBox.warning(self, "Adicionar material", str(exc))
            return
        self.accept()


class MaterialsPage(QWidget):
    page_title = "Matéria-Prima"
    page_subtitle = "Portefólio de lotes, formatos, reservas, valorização e rastreabilidade."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.current_material_id = ""
        self._combo_keys = ("formato", "material", "material_familia", "espessura", "local")

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(6)

        form_card = CardFrame()
        form_card.set_tone("info")
        form_card.setObjectName("MaterialPortfolio")
        form_card.setStyleSheet(
            "QLineEdit:disabled, QComboBox:disabled {"
            " background: #f8fbff;"
            " color: #0f172a;"
            " border: 1px solid #d6e3f3;"
            " border-radius: 8px;"
            "}"
            "QFrame#MaterialMetric { background: #ffffff; border: 1px solid #c6d5e5; }"
            "QFrame#MaterialReserveMetric { background: #fff7e6; border: 1px solid #efd29b; }"
            "QFrame#MaterialValueMetric { background: #e8f6f4; border: 1px solid #9fd8d3; }"
            "QLabel#MaterialMetricLabel { color: #5b7088; font-size: 8px; font-weight: 700; border: none; background: transparent; }"
            "QLabel#MaterialMetricValue { color: #10253d; font-size: 13px; font-weight: 800; border: none; background: transparent; }"
            "QLabel#MaterialReserveValue { color: #9a5b00; font-size: 13px; font-weight: 800; border: none; background: transparent; }"
            "QLabel#MaterialValueValue { color: #087f83; font-size: 13px; font-weight: 800; border: none; background: transparent; }"
        )
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(6)

        top = QHBoxLayout()
        title = QLabel("Portefólio de Matéria-Prima")
        title.setStyleSheet("font-size: 17px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel("Lotes, formatos, disponibilidade, reservas e valorização numa única área de trabalho.")
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        subtitle.setMinimumWidth(0)
        subtitle.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        ttl_wrap = QVBoxLayout()
        ttl_wrap.setContentsMargins(0, 0, 0, 0)
        ttl_wrap.setSpacing(2)
        ttl_wrap.addWidget(title)
        ttl_wrap.addWidget(subtitle)
        top.addLayout(ttl_wrap, 1)

        def _portfolio_metric(label_text: str, object_name: str, value_name: str) -> tuple[QFrame, QLabel]:
            frame = QFrame()
            frame.setObjectName(object_name)
            frame.setFixedSize(142, 52)
            layout = QVBoxLayout(frame)
            layout.setContentsMargins(10, 6, 10, 6)
            layout.setSpacing(0)
            label = QLabel(label_text.upper())
            label.setObjectName("MaterialMetricLabel")
            value = QLabel("-")
            value.setObjectName(value_name)
            layout.addWidget(label)
            layout.addWidget(value)
            return frame, value

        count_metric, self.material_count_metric = _portfolio_metric("Registos", "MaterialMetric", "MaterialMetricValue")
        reserve_metric, self.material_available_metric = _portfolio_metric("Disponível", "MaterialReserveMetric", "MaterialReserveValue")
        value_metric, self.material_value_metric = _portfolio_metric("Valor em stock", "MaterialValueMetric", "MaterialValueValue")
        top.addWidget(count_metric)
        top.addWidget(reserve_metric)
        top.addWidget(value_metric)
        self._stock_filter_timer = QTimer(self)
        self._stock_filter_timer.setSingleShot(True)
        self._stock_filter_timer.timeout.connect(self.refresh)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar material, lote, localização ou picar etiqueta...")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(lambda _text: self._stock_filter_timer.start(180))
        self.filter_edit.returnPressed.connect(self._select_scanned_material)
        self.only_stock_check = QCheckBox("Mostrar apenas com stock")
        self.only_stock_check.setChecked(False)
        self.only_stock_check.toggled.connect(self.refresh)
        form_layout.addLayout(top)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        filter_row.addWidget(self.filter_edit, 2)
        self.format_filter_combo = QComboBox()
        self.material_filter_combo = QComboBox()
        self.thickness_filter_combo = QComboBox()
        self.local_filter_combo = QComboBox()
        self.state_filter_combo = QComboBox()
        self.format_filter_combo.setProperty("allLabel", "Formato: Todos")
        self.material_filter_combo.setProperty("allLabel", "Material: Todos")
        self.thickness_filter_combo.setProperty("allLabel", "Esp.: Todas")
        self.local_filter_combo.setProperty("allLabel", "Local: Todos")
        self.state_filter_combo.addItems(["Estado: Todos", "Disponível", "Baixo", "Crítico", "Última unid.", "Retalhos", "Bloqueado/Qualidade"])
        for combo, width in (
            (self.format_filter_combo, 118),
            (self.material_filter_combo, 142),
            (self.thickness_filter_combo, 102),
            (self.local_filter_combo, 124),
            (self.state_filter_combo, 124),
        ):
            combo.setMinimumWidth(width)
            combo.setMaximumWidth(width + 36)
            combo.currentTextChanged.connect(lambda _text: self._stock_filter_timer.start(120))
            filter_row.addWidget(combo)
        filter_row.addWidget(self.only_stock_check)
        clear_filters_btn = QPushButton("Limpar")
        clear_filters_btn.setProperty("compact", "true")
        clear_filters_btn.setProperty("variant", "secondary")
        clear_filters_btn.setFixedWidth(68)
        clear_filters_btn.clicked.connect(self._clear_stock_filters)
        filter_row.addWidget(clear_filters_btn)
        form_layout.addLayout(filter_row)

        info_row = QHBoxLayout()
        info_row.setSpacing(10)
        self.selection_hint = QLabel("Seleciona uma linha da tabela para ver o detalhe completo aqui.")
        self.selection_hint.setStyleSheet("font-size: 12px; color: #1d4ed8; font-weight: 600;")
        self.stock_hint = QLabel("Stock baixo deixa de pintar a linha toda e passa a ser assinalado só em Disponível.")
        self.stock_hint.setProperty("role", "muted")
        info_row.addWidget(self.selection_hint, 1)
        info_row.addWidget(self.stock_hint)

        detail_card = CardFrame()
        detail_card.setStyleSheet("QFrame#Card { background: #f8fbff; border-color: #d6e3f3; }")
        detail_layout = QHBoxLayout(detail_card)
        detail_layout.setContentsMargins(12, 10, 12, 10)
        detail_layout.setSpacing(12)
        detail_text = QVBoxLayout()
        detail_text.setContentsMargins(0, 0, 0, 0)
        detail_text.setSpacing(2)
        self.detail_title = QLabel("Nenhum registo selecionado")
        self.detail_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.detail_meta = QLabel("Escolhe uma linha para ver o resumo técnico e de stock.")
        self.detail_meta.setProperty("role", "muted")
        detail_text.addWidget(self.detail_title)
        detail_text.addWidget(self.detail_meta)
        detail_layout.addLayout(detail_text, 1)
        self.detail_status = QLabel("Sem seleção")
        self.detail_status.setAlignment(Qt.AlignCenter)
        self.detail_status.setMinimumWidth(116)
        self.detail_status.setStyleSheet(
            "background: #eef2f8; color: #334155; border: 1px solid #d6e3f3; "
            "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
        )
        detail_layout.addWidget(self.detail_status)
        self.detail_available = QLabel("Disponível: -")
        self.detail_available.setAlignment(Qt.AlignCenter)
        self.detail_available.setMinimumWidth(140)
        self.detail_available.setStyleSheet(
            "background: white; color: #0f172a; border: 1px solid #d6e3f3; "
            "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
        )
        detail_layout.addWidget(self.detail_available)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        self.formato_combo = self._make_combo()
        self.material_combo = self._make_combo()
        self.material_family_combo = QComboBox()
        self.secao_tipo_combo = self._make_combo()
        self.espessura_combo = self._make_combo()
        self.local_combo = self._make_combo()
        self.lote_interno_edit = QLineEdit()
        self.lote_edit = QLineEdit()
        self.comprimento_edit = QLineEdit()
        self.largura_edit = QLineEdit()
        self.altura_edit = QLineEdit()
        self.diametro_edit = QLineEdit()
        self.contorno_edit = QLineEdit()
        self.metros_edit = QLineEdit()
        self.kg_m_edit = QLineEdit()
        self.peso_edit = QLineEdit()
        self.preco_compra_edit = QLineEdit()
        self.preco_unit_edit = QLineEdit()
        self.quantidade_edit = QLineEdit()
        self.reservado_edit = QLineEdit()
        self.preco_unit_edit.setReadOnly(True)
        self.preco_unit_edit.setFocusPolicy(Qt.NoFocus)
        self.preco_unit_edit.setPlaceholderText("0,00 EUR")
        self.peso_edit.setReadOnly(True)
        self.lote_interno_edit.setReadOnly(True)
        self.lote_interno_edit.setPlaceholderText("Gerado automaticamente")
        self.preco_compra_edit.setPlaceholderText("EUR/kg ou EUR/m")
        self.contorno_edit.setPlaceholderText("Opcional: 0,0; 1000,0; 900,400; 0,400")
        self.kg_m_edit.setPlaceholderText("Auto por tabela / fórmula")
        for combo, placeholder in (
            (self.formato_combo, "Formato"),
            (self.material_combo, "Material"),
            (self.secao_tipo_combo, "Tipo / série"),
            (self.espessura_combo, "Espessura"),
            (self.local_combo, "Localização"),
        ):
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.setPlaceholderText(placeholder)

        fields = [
            ("Formato", self.formato_combo),
            ("Material", self.material_combo),
            ("Família", self.material_family_combo),
            ("Tipo secção", self.secao_tipo_combo),
            ("Espessura", self.espessura_combo),
            ("Lote interno", self.lote_interno_edit),
            ("Lote fornecedor", self.lote_edit),
            ("Comprimento", self.comprimento_edit),
            ("Largura", self.largura_edit),
            ("Altura", self.altura_edit),
            ("Diâmetro", self.diametro_edit),
            ("Contorno retalho", self.contorno_edit),
            ("Metros", self.metros_edit),
            ("Kg/m", self.kg_m_edit),
            ("Peso/Un.", self.peso_edit),
            ("Compra (EUR/kg|m)", self.preco_compra_edit),
            ("Preço/Unid.", self.preco_unit_edit),
            ("Quantidade", self.quantidade_edit),
            ("Reservado", self.reservado_edit),
            ("Localização", self.local_combo),
        ]
        self._field_labels: dict[str, QLabel] = {}
        self._field_widgets: dict[str, QWidget] = {}
        for index, (label_text, widget) in enumerate(fields):
            row = index // 5
            col = (index % 5) * 2
            label = QLabel(label_text)
            label.setProperty("role", "muted")
            self._field_labels[label_text] = label
            self._field_widgets[label_text] = widget
            grid.addWidget(label, row, col)
            grid.addWidget(widget, row, col + 1)
        actions_card = CardFrame()
        actions_card.set_tone("default")
        actions_card.setMaximumHeight(50)
        actions_primary = QHBoxLayout(actions_card)
        actions_primary.setContentsMargins(10, 6, 10, 6)
        actions_primary.setSpacing(6)
        self.add_btn = QPushButton("Adicionar")
        self.add_btn.clicked.connect(self.add_material)
        self.edit_btn = QPushButton("Editar")
        self.edit_btn.clicked.connect(self.edit_material)
        self.remove_btn = QPushButton("Remover")
        self.remove_btn.setProperty("variant", "danger")
        self.remove_btn.clicked.connect(self.remove_material)
        self.baixa_btn = QPushButton("Dar baixa")
        self.baixa_btn.setProperty("variant", "secondary")
        self.baixa_btn.clicked.connect(self.consume_material)
        self.correct_btn = QPushButton("Corrigir")
        self.correct_btn.setProperty("variant", "secondary")
        self.correct_btn.clicked.connect(self.correct_material)
        self.history_btn = QPushButton("Histórico")
        self.history_btn.setProperty("variant", "secondary")
        self.history_btn.clicked.connect(self.show_history)
        self.label_btn = QPushButton("Etiqueta")
        self.label_btn.setProperty("variant", "secondary")
        self.label_btn.clicked.connect(self.preview_label)
        self.label_print_btn = QPushButton("Imprimir ID")
        self.label_print_btn.setProperty("variant", "secondary")
        self.label_print_btn.clicked.connect(self.print_label)
        self.label_save_btn = QPushButton("Guardar ID")
        self.label_save_btn.setProperty("variant", "secondary")
        self.label_save_btn.clicked.connect(self.save_label)
        label_menu = QMenu(self.label_btn)
        preview_label_action = label_menu.addAction("Pré-visualizar")
        print_label_action = label_menu.addAction("Imprimir")
        save_label_action = label_menu.addAction("Guardar PDF")
        preview_label_action.triggered.connect(self.preview_label)
        print_label_action.triggered.connect(self.print_label)
        save_label_action.triggered.connect(self.save_label)
        self.label_btn.setMenu(label_menu)
        self.label_print_btn.hide()
        self.label_save_btn.hide()
        self.pdf_btn = QPushButton("Preview PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self.preview_pdf)
        self.calc_btn = QPushButton("Calc. peso")
        self.calc_btn.setProperty("variant", "secondary")
        self.calc_btn.clicked.connect(self._open_weight_calculator)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.export_btn = QPushButton("CSV")
        self.export_btn.setProperty("variant", "secondary")
        self.export_btn.clicked.connect(self.export_csv)
        self.full_grid_btn = QPushButton("Grelha")
        self.full_grid_btn.setProperty("variant", "secondary")
        self.full_grid_btn.clicked.connect(self.open_full_grid)
        for button, width in (
            (self.add_btn, 78),
            (self.edit_btn, 72),
            (self.remove_btn, 80),
            (self.baixa_btn, 88),
            (self.correct_btn, 78),
            (self.history_btn, 82),
            (self.label_btn, 88),
            (self.pdf_btn, 90),
            (self.calc_btn, 88),
            (self.refresh_btn, 82),
            (self.export_btn, 58),
            (self.full_grid_btn, 74),
        ):
            button.setProperty("compact", "true")
            button.setMinimumHeight(30)
            button.setFixedWidth(width)
            button.setStyleSheet("font-size: 10px; font-weight: 700;")
            actions_primary.addWidget(button)
        actions_primary.addStretch(1)
        root.addWidget(form_card)
        root.addWidget(actions_card)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(14, 12, 14, 12)
        table_layout.setSpacing(8)
        table_header = QHBoxLayout()
        table_title = QLabel("Catálogo de matéria-prima")
        table_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #10253d;")
        self.table_count_label = QLabel("-")
        self.table_count_label.setProperty("role", "muted")
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(self.table_count_label)
        table_layout.addLayout(table_header)
        self.table = QTableWidget(0, 18)
        self.table.setObjectName("StockTable")
        self.table.setStyleSheet(
            "QTableWidget {"
            " gridline-color: #d8e3f2;"
            " selection-background-color: #dff5f3;"
            " selection-color: #0f172a;"
            "}"
            "QTableWidget::item:selected {"
            " background: #dff5f3;"
            " color: #0f172a;"
            " border-top: 1px solid #08a6a6;"
            " border-bottom: 1px solid #08a6a6;"
            "}"
            "QHeaderView::section {"
            " background: #0b0f5c;"
            " color: white;"
            " padding: 8px 6px;"
            " border: 0;"
            " font-weight: 700;"
            "}"
        )
        self.table.setHorizontalHeaderLabels(
            [
                "Lote interno",
                "Lote fornecedor",
                "Material",
                "Dimensões (mm)",
                "Dim. B",
                "Esp.",
                "Qtd.",
                "Reserva",
                "Formato",
                "Metros (m)",
                "Peso/Un. (kg)",
                "Compra (EUR)",
                "Preço/Unid.",
                "Disponível",
                "Tipo",
                "Localização",
                "ID",
                "Estado",
            ]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.setWordWrap(False)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.itemDoubleClicked.connect(lambda *_args: self.edit_material())
        header = self.table.horizontalHeader()
        header.setFixedHeight(34)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(48)
        column_specs = [
            (0, QHeaderView.Interactive, 134),  # Lote interno
            (1, QHeaderView.Interactive, 130),  # Lote fornecedor
            (2, QHeaderView.Interactive, 260),  # Material
            (3, QHeaderView.Fixed, 72),         # Dim. A
            (4, QHeaderView.Fixed, 72),         # Dim. B
            (5, QHeaderView.Fixed, 78),         # Espessura
            (6, QHeaderView.Fixed, 86),         # Quantidade
            (7, QHeaderView.Fixed, 72),         # Reserva
            (8, QHeaderView.Fixed, 92),         # Formato
            (9, QHeaderView.Fixed, 92),         # Metros
            (10, QHeaderView.Interactive, 112), # Peso/Un.
            (11, QHeaderView.Fixed, 104),       # Compra
            (12, QHeaderView.Fixed, 104),       # Preço
            (13, QHeaderView.Fixed, 104),       # Disponível
            (14, QHeaderView.Interactive, 210), # Tipo
            (15, QHeaderView.Interactive, 132), # Localização
            (16, QHeaderView.Fixed, 86),        # ID
            (17, QHeaderView.Interactive, 118), # Estado
        ]
        for column, mode, width in column_specs:
            header.setSectionResizeMode(column, mode)
            if mode != QHeaderView.Stretch:
                header.resizeSection(column, width)
        visible_columns = (16, 2, 3, 8, 0, 5, 6, 13, 17)
        for column in range(self.table.columnCount()):
            self.table.setColumnHidden(column, column not in visible_columns)
        for visual_target, logical_index in enumerate(visible_columns):
            current_visual = header.visualIndex(logical_index)
            if current_visual != visual_target:
                header.moveSection(current_visual, visual_target)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        for column, width in ((16, 76), (3, 118), (8, 76), (0, 100), (5, 54), (6, 58), (13, 68), (17, 94)):
            header.setSectionResizeMode(column, QHeaderView.Interactive)
            header.resizeSection(column, width)
        table_layout.addWidget(self.table)

        inspector = CardFrame()
        inspector.set_tone("default")
        inspector.setMinimumWidth(390)
        inspector.setMaximumWidth(560)
        inspector.setStyleSheet(
            "QLineEdit:disabled, QComboBox:disabled { background: #f8fbff; color: #0f172a; border: 1px solid #d6e3f3; }"
        )
        inspector_layout = QVBoxLayout(inspector)
        inspector_layout.setContentsMargins(12, 10, 12, 12)
        inspector_layout.setSpacing(9)
        inspector_header = QHBoxLayout()
        inspector_heading = QVBoxLayout()
        inspector_heading.setSpacing(1)
        inspector_eyebrow = QLabel("MATÉRIA-PRIMA SELECIONADA")
        inspector_eyebrow.setStyleSheet("color: #5b7088; font-size: 8px; font-weight: 700;")
        self.detail_title.setWordWrap(True)
        self.detail_meta.setWordWrap(True)
        inspector_heading.addWidget(inspector_eyebrow)
        inspector_heading.addWidget(self.detail_title)
        inspector_heading.addWidget(self.detail_meta)
        inspector_header.addLayout(inspector_heading, 1)
        inspector_header.addWidget(self.detail_status, 0, Qt.AlignTop)
        inspector_layout.addLayout(inspector_header)
        self.selection_hint.setParent(inspector)
        self.selection_hint.hide()
        self.detail_available.setParent(inspector)
        self.detail_available.hide()

        summary_strip = QFrame()
        summary_strip.setObjectName("MaterialSummaryStrip")
        summary_strip.setFixedHeight(58)
        summary_strip.setStyleSheet(
            "QFrame#MaterialSummaryStrip { background: #f1f6fa; border: none; }"
            "QLabel#MaterialSummaryLabel { color: #60758d; font-size: 8px; font-weight: 700; border: none; background: transparent; }"
            "QLabel#MaterialSummaryValue { color: #10253d; font-size: 11px; font-weight: 800; border: none; background: transparent; }"
        )
        summary_layout = QHBoxLayout(summary_strip)
        summary_layout.setContentsMargins(10, 6, 10, 6)
        summary_layout.setSpacing(10)

        def _summary_metric(label_text: str) -> tuple[QWidget, QLabel]:
            host = QWidget()
            layout = QVBoxLayout(host)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(0)
            label = QLabel(label_text)
            label.setObjectName("MaterialSummaryLabel")
            value = QLabel("-")
            value.setObjectName("MaterialSummaryValue")
            layout.addWidget(label)
            layout.addWidget(value)
            return host, value

        qty_summary, self.inspector_qty_label = _summary_metric("Stock físico")
        reserved_summary, self.inspector_reserved_label = _summary_metric("Reservado")
        available_summary, self.inspector_available_label = _summary_metric("Disponível")
        price_summary, self.inspector_price_label = _summary_metric("Preço/unid.")
        value_summary, self.inspector_value_label = _summary_metric("Valor stock")
        for host in (qty_summary, reserved_summary, available_summary, price_summary, value_summary):
            summary_layout.addWidget(host, 1)
        inspector_layout.addWidget(summary_strip)

        self.detail_tabs = QTabWidget()
        self.detail_tabs.setDocumentMode(True)
        self.detail_tabs.setUsesScrollButtons(False)
        self.detail_tabs.setStyleSheet(
            "QTabWidget::pane { border: 1px solid #cbd8e5; background: #ffffff; top: -1px; }"
            "QTabBar::tab { min-width: 92px; min-height: 30px; padding: 0 8px; color: #4a6179; font-size: 9px; font-weight: 700; }"
            "QTabBar::tab:selected { background: #ffffff; color: #087f83; border-bottom: 2px solid #08a6a6; }"
        )
        inspector_layout.addWidget(self.detail_tabs, 1)

        def _tab_page(fields_to_add: list[tuple[str, QWidget]]) -> QWidget:
            page = QWidget()
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            content = QWidget()
            scroll.setWidget(content)
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(0, 0, 0, 0)
            page_layout.addWidget(scroll)
            content_layout = QVBoxLayout(content)
            content_layout.setContentsMargins(10, 10, 10, 10)
            field_grid = QGridLayout()
            field_grid.setHorizontalSpacing(8)
            field_grid.setVerticalSpacing(5)
            for index, (label_text, widget) in enumerate(fields_to_add):
                row = index // 2
                column = index % 2
                label = QLabel(label_text)
                label.setStyleSheet("color: #5b7088; font-size: 8px; font-weight: 700;")
                self._field_labels[label_text] = label
                field_grid.addWidget(label, row * 2, column)
                field_grid.addWidget(widget, row * 2 + 1, column)
            field_grid.setColumnStretch(0, 1)
            field_grid.setColumnStretch(1, 1)
            content_layout.addLayout(field_grid)
            content_layout.addStretch(1)
            return page

        identification_page = _tab_page(
            [
                ("Formato", self.formato_combo),
                ("Material", self.material_combo),
                ("Família", self.material_family_combo),
                ("Tipo secção", self.secao_tipo_combo),
                ("Espessura", self.espessura_combo),
                ("Lote interno", self.lote_interno_edit),
                ("Lote fornecedor", self.lote_edit),
                ("Localização", self.local_combo),
            ]
        )
        geometry_page = _tab_page(
            [
                ("Comprimento", self.comprimento_edit),
                ("Largura", self.largura_edit),
                ("Altura", self.altura_edit),
                ("Diâmetro", self.diametro_edit),
                ("Metros", self.metros_edit),
                ("Kg/m", self.kg_m_edit),
                ("Peso/Un.", self.peso_edit),
                ("Contorno retalho", self.contorno_edit),
            ]
        )
        stock_page = _tab_page(
            [
                ("Quantidade", self.quantidade_edit),
                ("Reservado", self.reservado_edit),
                ("Compra (EUR/kg|m)", self.preco_compra_edit),
                ("Preço/Unid.", self.preco_unit_edit),
            ]
        )
        self.detail_tabs.addTab(identification_page, "Identificação")
        self.detail_tabs.addTab(geometry_page, "Geometria")
        self.detail_tabs.addTab(stock_page, "Stock e valor")

        workspace = QSplitter(Qt.Horizontal)
        workspace.setChildrenCollapsible(False)
        workspace.addWidget(table_card)
        workspace.addWidget(inspector)
        workspace.setStretchFactor(0, 1)
        workspace.setStretchFactor(1, 0)
        workspace.setSizes([1400, 440])
        root.addWidget(workspace, 1)

        for combo in (self.formato_combo, self.material_combo, self.secao_tipo_combo, self.espessura_combo, self.local_combo):
            combo.currentTextChanged.connect(self._on_form_value_changed)
        self.material_family_combo.currentIndexChanged.connect(self._on_form_value_changed)
        for edit in (
            self.lote_edit,
            self.lote_interno_edit,
            self.comprimento_edit,
            self.largura_edit,
            self.altura_edit,
            self.diametro_edit,
            self.contorno_edit,
            self.metros_edit,
            self.kg_m_edit,
            self.peso_edit,
            self.preco_compra_edit,
            self.quantidade_edit,
            self.reservado_edit,
        ):
            edit.textChanged.connect(self._on_form_value_changed)

        self._set_preview_fields_enabled(False)
        self._set_form_defaults()

    def _make_combo(self) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMinimumContentsLength(10)
        return combo

    def _set_combo_values(self, combo: QComboBox, values: list[str], current_text: str = "") -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(values)
        combo.setCurrentText(current_text)
        combo.blockSignals(False)

    def _set_section_options(self, formato: str, current_value: str = "") -> None:
        options = [
            str(row.get("label", "") or "").strip()
            for row in list(self.backend.material_section_options(formato) or [])
            if str(row.get("label", "") or "").strip()
        ]
        self._set_combo_values(self.secao_tipo_combo, options, current_value)

    def _set_field_visible(self, key: str, visible: bool) -> None:
        label = self._field_labels.get(key)
        widget = self._field_widgets.get(key)
        if label is not None:
            label.setVisible(visible)
        if widget is not None:
            widget.setVisible(visible)

    def _set_form_defaults(self) -> None:
        self.formato_combo.setCurrentText("Chapa")
        self.material_combo.setCurrentText("")
        _set_material_family_combo(self.backend, self.material_family_combo, "")
        self._set_section_options("Chapa", "")
        self.espessura_combo.setCurrentText("")
        self.local_combo.setCurrentText("")
        for edit in (
            self.lote_edit,
            self.lote_interno_edit,
            self.comprimento_edit,
            self.largura_edit,
            self.altura_edit,
            self.diametro_edit,
            self.contorno_edit,
            self.metros_edit,
            self.kg_m_edit,
            self.peso_edit,
            self.preco_compra_edit,
            self.quantidade_edit,
            self.reservado_edit,
        ):
            edit.clear()
        self.preco_unit_edit.setText("0,00 EUR")
        self.reservado_edit.setText("0")
        self.current_material_id = ""
        selection_model = self.table.selectionModel()
        if selection_model is not None:
            self.table.blockSignals(True)
            self.table.clearSelection()
            try:
                selection_model.clearCurrentIndex()
            except Exception:
                pass
            self.table.blockSignals(False)
        self.selection_hint.setText("Seleciona uma linha da tabela para ver o detalhe completo aqui.")
        self._set_detail_summary(None)
        self._refresh_form_state()
        self._refresh_price_preview()

    def _severity_for_record(self, record: dict[str, object] | None) -> str:
        if not isinstance(record, dict):
            return "ok"
        quantidade = self.backend._parse_float(record.get("quantidade", 0), 0)
        reservado = self.backend._parse_float(record.get("reservado", 0), 0)
        disponivel = quantidade - reservado
        if quantidade == 1:
            return "one"
        if disponivel <= float(self.backend.desktop_main.STOCK_VERMELHO):
            return "critical"
        if disponivel <= float(self.backend.desktop_main.STOCK_AMARELO):
            return "warning"
        return "ok"

    def _set_detail_summary(self, record: dict[str, object] | None) -> None:
        if not isinstance(record, dict):
            self.detail_title.setText("Nenhum registo selecionado")
            self.detail_meta.setText("Escolhe uma linha para ver o resumo técnico e de stock.")
            self.detail_status.setText("Sem seleção")
            self.detail_status.setStyleSheet(
                "background: #eef2f8; color: #334155; border: 1px solid #d6e3f3; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
            )
            self.detail_available.setText("Disponível: -")
            self.detail_available.setStyleSheet(
                "background: white; color: #0f172a; border: 1px solid #d6e3f3; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
            )
            for label in (
                self.inspector_qty_label,
                self.inspector_reserved_label,
                self.inspector_available_label,
                self.inspector_price_label,
                self.inspector_value_label,
            ):
                label.setText("-")
            return
        material_id = str(record.get("id", "") or "").strip()
        material = str(record.get("material", "") or "").strip() or "Sem material"
        formato = str(record.get("formato", "") or "").strip() or "-"
        local = self.backend._localizacao(record) or "Sem localização"
        lote = str(record.get("lote_interno", "") or "").strip() or "Sem lote interno"
        lote_fornecedor = str(record.get("lote_fornecedor", "") or "").strip()
        quantidade = self.backend._parse_float(record.get("quantidade", 0), 0)
        reservado = self.backend._parse_float(record.get("reservado", 0), 0)
        disponivel = quantidade - reservado
        preco_unid = self.backend._parse_float(record.get("preco_unid", record.get("preco_unit", 0)), 0)
        valor_stock = quantidade * preco_unid
        self.detail_title.setText(f"{material} | {material_id}")
        fornecedor_txt = f" | Fornecedor: {lote_fornecedor}" if lote_fornecedor else ""
        self.detail_meta.setText(f"{formato} | Lote interno: {lote}{fornecedor_txt} | Localização: {local}")
        severity = self._severity_for_record(record)
        if severity == "critical":
            self.detail_status.setText("Stock crítico")
            self.detail_status.setStyleSheet(
                "background: #fee4e2; color: #b42318; border: 1px solid #f3b7b3; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 800;"
            )
        elif severity == "warning":
            self.detail_status.setText("Stock baixo")
            self.detail_status.setStyleSheet(
                "background: #fff4e5; color: #b54708; border: 1px solid #efcf98; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 800;"
            )
        elif severity == "one":
            self.detail_status.setText("Última unidade")
            self.detail_status.setStyleSheet(
                "background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 800;"
            )
        else:
            self.detail_status.setText("Stock normal")
            self.detail_status.setStyleSheet(
                "background: #ecfdf3; color: #027a48; border: 1px solid #abefc6; "
                "border-radius: 10px; padding: 6px 10px; font-weight: 800;"
            )
        self.detail_available.setText(
            f"Qtd: {self.backend._fmt(quantidade)}  |  Disponível: {self.backend._fmt(disponivel)}"
        )
        self.detail_available.setStyleSheet(
            "background: white; color: #0f172a; border: 1px solid #d6e3f3; "
            "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
        )
        self.inspector_qty_label.setText(self.backend._fmt(quantidade))
        self.inspector_reserved_label.setText(self.backend._fmt(reservado))
        self.inspector_available_label.setText(self.backend._fmt(disponivel))
        self.inspector_price_label.setText(
            f"{preco_unid:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        self.inspector_value_label.setText(
            f"{valor_stock:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    def _set_preview_fields_enabled(self, enabled: bool) -> None:
        preview_only = (
            self.formato_combo,
            self.material_combo,
            self.material_family_combo,
            self.secao_tipo_combo,
            self.espessura_combo,
            self.local_combo,
            self.lote_interno_edit,
            self.lote_edit,
            self.comprimento_edit,
            self.largura_edit,
            self.altura_edit,
            self.diametro_edit,
            self.contorno_edit,
            self.metros_edit,
            self.kg_m_edit,
            self.peso_edit,
            self.preco_compra_edit,
            self.preco_unit_edit,
            self.quantidade_edit,
            self.reservado_edit,
        )
        for widget in preview_only:
            widget.setEnabled(enabled)

    def _payload(self) -> dict[str, str]:
        return {
            "formato": self.formato_combo.currentText().strip(),
            "material": self.material_combo.currentText().strip(),
            "material_familia": str(self.material_family_combo.currentData() or "").strip(),
            "secao_tipo": self.secao_tipo_combo.currentText().strip(),
            "espessura": self.espessura_combo.currentText().strip(),
            "comprimento": self.comprimento_edit.text().strip(),
            "largura": self.largura_edit.text().strip(),
            "altura": self.altura_edit.text().strip(),
            "diametro": self.diametro_edit.text().strip(),
            "contorno_points": self.contorno_edit.text().strip(),
            "metros": self.metros_edit.text().strip(),
            "kg_m": self.kg_m_edit.text().strip(),
            "peso_unid": self.peso_edit.text().strip(),
            "p_compra": self.preco_compra_edit.text().strip(),
            "quantidade": self.quantidade_edit.text().strip(),
            "reservado": self.reservado_edit.text().strip(),
            "local": self.local_combo.currentText().strip(),
            "lote_interno": self.lote_interno_edit.text().strip(),
            "lote_fornecedor": self.lote_edit.text().strip(),
        }

    def _on_form_value_changed(self, *_args) -> None:
        self._refresh_form_state()
        self._refresh_price_preview()

    def _refresh_form_state(self) -> None:
        formato = (self.formato_combo.currentText().strip() or "Chapa").title()
        self._set_section_options(formato, self.secao_tipo_combo.currentText().strip())
        preview = self.backend.material_geometry_preview(self._payload())
        secao_tipo = str(preview.get("secao_tipo", "") or "").strip()
        tube_round = formato == "Tubo" and secao_tipo == "redondo"
        profile_catalog = formato == "Perfil" and bool(preview.get("usa_catalogo"))
        espessura_required = formato in {"Chapa", "Tubo", "Cantoneira"}
        esp_label = self._field_labels.get("Espessura")
        if esp_label is not None:
            esp_label.setText("Espessura" if espessura_required else "Espessura (opc.)")
        secao_label = self._field_labels.get("Tipo secção")
        if secao_label is not None:
            if formato == "Tubo":
                secao_label.setText("Tipo tubo")
            elif formato == "Perfil":
                secao_label.setText("Tipo perfil / série")
            elif formato == "Cantoneira":
                secao_label.setText("Tipo cantoneira")
            elif formato == "Barra":
                secao_label.setText("Tipo barra")
            else:
                secao_label.setText("Tipo secção")
        comp_label = self._field_labels.get("Comprimento")
        if comp_label is not None:
            if formato == "Chapa":
                comp_label.setText("Comprimento (mm)")
            elif formato == "Tubo":
                comp_label.setText("Lado A (mm)")
            elif formato == "Cantoneira":
                comp_label.setText("Aba A (mm)")
            elif formato == "Barra":
                comp_label.setText("Lado A (mm)")
            else:
                comp_label.setText("Comprimento")
        larg_label = self._field_labels.get("Largura")
        if larg_label is not None:
            if formato == "Chapa":
                larg_label.setText("Largura (mm)")
            elif formato == "Cantoneira":
                larg_label.setText("Aba B (mm)")
            else:
                larg_label.setText("Lado B (mm)")
        altura_label = self._field_labels.get("Altura")
        if altura_label is not None:
            altura_label.setText("Altura / tamanho (mm)")
        diametro_label = self._field_labels.get("Diâmetro")
        if diametro_label is not None:
            diametro_label.setText("Diâmetro ext. (mm)")
        kgm_label = self._field_labels.get("Kg/m")
        if kgm_label is not None:
            kgm_label.setText("Peso por metro (kg/m)")
        compra_label = self._field_labels.get("Compra (EUR/kg|m)")
        if compra_label is not None:
            compra_label.setText("Compra (EUR/m)" if formato == "Tubo" else "Compra (EUR/kg)")
        esp_line = self.espessura_combo.lineEdit()
        if esp_line is not None:
            esp_line.setPlaceholderText("Obrigatória" if espessura_required else "Opcional / derivada")
        self.espessura_combo.setToolTip("Obrigatória para chapa, tubo e cantoneira." if espessura_required else "Opcional quando a secção já define a espessura.")
        self.preco_compra_edit.setPlaceholderText("EUR/m" if formato == "Tubo" else "EUR/kg")
        self.preco_compra_edit.setToolTip("Preço de compra base por metro." if formato == "Tubo" else "Preço de compra base por kg.")
        self.metros_edit.setPlaceholderText("Comprimento barra (m)" if formato in {"Tubo", "Perfil", "Cantoneira", "Barra"} else "")
        self.peso_edit.setPlaceholderText("Peso calculado automaticamente")
        self.kg_m_edit.setReadOnly(formato == "Tubo" or profile_catalog)
        self._set_field_visible("Tipo secção", formato in {"Tubo", "Perfil", "Cantoneira", "Barra"})
        self._set_field_visible("Comprimento", formato in {"Chapa", "Cantoneira", "Barra"} or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Largura", formato in {"Chapa", "Cantoneira", "Barra"} or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Altura", formato == "Perfil")
        self._set_field_visible("Diâmetro", formato == "Tubo" and tube_round)
        self._set_field_visible("Contorno retalho", formato == "Chapa")
        self._set_field_visible("Metros", formato != "Chapa")
        self._set_field_visible("Kg/m", formato in {"Tubo", "Perfil", "Cantoneira", "Barra"})

    def _refresh_price_preview(self) -> None:
        try:
            preview = self.backend.material_price_preview(self._payload())
        except Exception:
            self.preco_unit_edit.setText("0,00 EUR")
            self.preco_unit_edit.setToolTip("")
            return
        self.peso_edit.blockSignals(True)
        self.peso_edit.setText(self.backend._fmt(preview.get("peso_unid", 0)))
        self.peso_edit.blockSignals(False)
        if str(preview.get("formato", "") or "") in {"Tubo", "Perfil", "Cantoneira", "Barra"}:
            self.kg_m_edit.blockSignals(True)
            self.kg_m_edit.setText(self.backend._fmt(preview.get("kg_m", 0)))
            self.kg_m_edit.blockSignals(False)
        preco_unid = float(preview.get("preco_unid", 0.0) or 0.0)
        self.preco_unit_edit.setText(f"{preco_unid:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."))
        if self.current_material_id:
            quantidade = self.backend._parse_float(self.quantidade_edit.text(), 0)
            self.inspector_price_label.setText(
                f"{preco_unid:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            self.inspector_value_label.setText(
                f"{quantidade * preco_unid:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")
            )
        tooltip_bits = [
            f"Formato: {preview.get('formato', '-')}",
            f"Secção: {preview.get('secao_label', '-')}",
            f"Dimensão: {preview.get('dimension_label', '-')}",
            f"Família: {preview.get('material_familia_label', 'Aço / Ferro')}",
            f"Densidade: {self.backend._fmt(preview.get('densidade', 0))} g/cm3",
            f"Base: {float(preview.get('p_compra', 0.0) or 0.0):.4f} {preview.get('base_label', 'EUR/kg')}",
        ]
        if str(preview.get("formato", "") or "") == "Tubo":
            tooltip_bits.append(f"Metros/unid.: {self.backend._fmt(preview.get('metros', 0))}")
        else:
            tooltip_bits.append(f"Peso/unid.: {self.backend._fmt(preview.get('peso_unid', 0))} kg")
        if float(preview.get("kg_m", 0) or 0) > 0:
            tooltip_bits.append(f"Kg/m: {self.backend._fmt(preview.get('kg_m', 0))}")
        if str(preview.get("calc_hint", "") or "").strip():
            tooltip_bits.append(str(preview.get("calc_hint", "") or "").strip())
        self.preco_unit_edit.setToolTip(" | ".join(tooltip_bits))

    def _apply_row_colors(self, row_index: int, severity: str, band: str) -> None:
        background = QColor("#ffffff" if band == "even" else "#f6f9fd")
        foreground = QColor("#0f172a")
        for col in range(self.table.columnCount()):
            item = self.table.item(row_index, col)
            if item is None:
                continue
            item.setBackground(QBrush(background))
            item.setForeground(QBrush(foreground))
        available_item = self.table.item(row_index, 13)
        if available_item is None:
            return
        if severity == "critical":
            available_item.setBackground(QBrush(QColor("#fee4e2")))
            available_item.setForeground(QBrush(QColor("#b42318")))
            available_item.setToolTip("Stock crítico: disponível abaixo do limite vermelho.")
        elif severity == "warning":
            available_item.setBackground(QBrush(QColor("#fff4e5")))
            available_item.setForeground(QBrush(QColor("#b54708")))
            available_item.setToolTip("Stock baixo: disponível abaixo do limite amarelo.")
        elif severity == "one":
            available_item.setBackground(QBrush(QColor("#eff6ff")))
            available_item.setForeground(QBrush(QColor("#1d4ed8")))
            available_item.setToolTip("Última unidade em stock.")
        else:
            available_item.setToolTip("")

    def _stock_state_label(self, record: dict[str, object] | None, severity: str) -> str:
        if isinstance(record, dict):
            q_status = str(record.get("quality_status", record.get("inspection_status", "")) or "").strip()
            if q_status and q_status != "APROVADO":
                return q_status
        if severity == "critical":
            return "Crítico"
        if severity == "warning":
            return "Baixo"
        if severity == "one":
            return "Última unid."
        return "Disponível"

    def _apply_state_cell_style(self, row_index: int, severity: str, state_label: str) -> None:
        item = self.table.item(row_index, 17)
        if item is None:
            return
        state_norm = state_label.casefold()
        if any(token in state_norm for token in ("rejeit", "devol", "bloque")):
            bg, fg = "#fee4e2", "#b42318"
        elif any(token in state_norm for token in ("inspec", "averig", "baixo")) or severity == "warning":
            bg, fg = "#fff7ed", "#9a3412"
        elif severity == "critical":
            bg, fg = "#fee4e2", "#b42318"
        else:
            bg, fg = "#ecfdf3", "#027a48"
        item.setBackground(QBrush(QColor(bg)))
        item.setForeground(QBrush(QColor(fg)))

    def _set_filter_combo_values(self, combo: QComboBox, values: list[str], current_text: str = "Todos") -> None:
        all_label = str(combo.property("allLabel") or "Todos")
        current = str(current_text or combo.currentText() or all_label).strip() or all_label
        ordered = [all_label]
        for value in values:
            text = str(value or "").strip()
            if text and text not in ordered:
                ordered.append(text)
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(ordered)
        combo.setCurrentText(current if current in ordered else all_label)
        combo.blockSignals(False)

    def _clear_stock_filters(self) -> None:
        self.filter_edit.clear()
        for combo in (
            self.format_filter_combo,
            self.material_filter_combo,
            self.thickness_filter_combo,
            self.local_filter_combo,
            self.state_filter_combo,
        ):
            combo.blockSignals(True)
            combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self.refresh()

    def _select_scanned_material(self) -> None:
        value = self.filter_edit.text().strip()
        if not value:
            return
        try:
            result = self.backend.inventory_scan_lookup(value, expected_type="MAT")
        except Exception as exc:
            if "|" in value or re.fullmatch(r"MAT\d+", value, flags=re.IGNORECASE):
                QMessageBox.warning(self, "Picagem", str(exc))
            return
        self.current_material_id = str(result.get("entity_id", "") or "").strip()
        self.filter_edit.blockSignals(True)
        self.filter_edit.clear()
        self.filter_edit.blockSignals(False)
        for combo in (
            self.format_filter_combo,
            self.material_filter_combo,
            self.thickness_filter_combo,
            self.local_filter_combo,
            self.state_filter_combo,
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
        return "" if text.lower() in {"", "todos", "todas", "all"} else text

    def _filter_material_payload_rows(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        formato = self._filter_combo_text(self.format_filter_combo).casefold()
        material = self._filter_combo_text(self.material_filter_combo).casefold()
        espessura = self._filter_combo_text(self.thickness_filter_combo).casefold()
        local = self._filter_combo_text(self.local_filter_combo).casefold()
        estado = self._filter_combo_text(self.state_filter_combo).casefold()
        filtered: list[dict[str, object]] = []
        for payload in list(rows or []):
            values = dict(payload.get("row", {}) or {})
            record = dict(payload.get("record", {}) or {})
            severity = str(payload.get("severity", "ok") or "ok")
            state_label = self._stock_state_label(record, severity)
            if formato and formato != str(values.get("formato", "") or "").casefold():
                continue
            if material and material != str(values.get("material", "") or "").casefold():
                continue
            if espessura and espessura != str(values.get("espessura", "") or "").casefold():
                continue
            if local and local != str(values.get("local", "") or "").casefold():
                continue
            if estado:
                state_norm = state_label.casefold()
                is_retalho = "retalho" in str(values.get("tipo", "") or "").casefold() or bool(record.get("is_sobra"))
                if estado == "disponível" and state_norm != "disponível":
                    continue
                if estado == "baixo" and "baixo" not in state_norm:
                    continue
                if estado == "crítico" and "cr" not in state_norm:
                    continue
                if estado == "última unid." and "última" not in state_norm:
                    continue
                if estado == "retalhos" and not is_retalho:
                    continue
                if estado == "bloqueado/qualidade" and not any(token in state_norm for token in ("rejeit", "devol", "bloque", "inspec", "averig")):
                    continue
            filtered.append(payload)
        return filtered

    def _selected_material_id(self) -> str:
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return ""
        rows = selection_model.selectedRows()
        if not rows:
            return ""
        item = self.table.item(rows[0].row(), 16)
        return item.text().strip() if item else ""

    def _selected_material_record(self) -> dict | None:
        material_id = self.current_material_id or self._selected_material_id()
        if not material_id:
            return None
        return self.backend.material_by_id(material_id)

    def refresh(self) -> None:
        presets = self.backend.material_presets()
        current_values = {
            "formato": self.formato_combo.currentText(),
            "material": self.material_combo.currentText(),
            "material_familia": str(self.material_family_combo.currentData() or "").strip(),
            "secao_tipo": self.secao_tipo_combo.currentText(),
            "espessura": self.espessura_combo.currentText(),
            "local": self.local_combo.currentText(),
        }
        self._set_combo_values(self.formato_combo, presets["formatos"], current_values["formato"] or "Chapa")
        self._set_combo_values(self.material_combo, presets["materiais"], current_values["material"])
        _set_material_family_combo(self.backend, self.material_family_combo, current_values["material_familia"], material=current_values["material"])
        self._set_section_options(current_values["formato"] or "Chapa", current_values["secao_tipo"])
        self._set_combo_values(self.espessura_combo, presets["espessuras"], current_values["espessura"])
        self._set_combo_values(self.local_combo, presets["locais"], current_values["local"])
        self._set_filter_combo_values(self.format_filter_combo, presets["formatos"], self.format_filter_combo.currentText())
        self._set_filter_combo_values(self.material_filter_combo, presets["materiais"], self.material_filter_combo.currentText())
        self._set_filter_combo_values(self.thickness_filter_combo, presets["espessuras"], self.thickness_filter_combo.currentText())
        self._set_filter_combo_values(self.local_filter_combo, presets["locais"], self.local_filter_combo.currentText())
        self._refresh_form_state()
        self._refresh_price_preview()

        portfolio_rows = self.backend.material_rows("", in_stock_only=False)
        available_total = 0.0
        portfolio_value = 0.0
        for payload in portfolio_rows:
            record = dict(payload.get("record", {}) or {})
            values = dict(payload.get("row", {}) or {})
            quantity = self.backend._parse_float(record.get("quantidade", values.get("quantidade", 0)), 0)
            reserved = self.backend._parse_float(record.get("reservado", values.get("reservado", 0)), 0)
            unit_price = self.backend._parse_float(
                record.get("preco_unid", record.get("preco_unit", values.get("preco_unid", 0))),
                0,
            )
            available_total += max(0.0, quantity - reserved)
            portfolio_value += max(0.0, quantity) * max(0.0, unit_price)
        self.material_count_metric.setText(str(len(portfolio_rows)))
        self.material_available_metric.setText(self.backend._fmt(available_total))
        self.material_value_metric.setText(
            f"{portfolio_value:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        all_rows = self.backend.material_rows(self.filter_edit.text(), in_stock_only=self.only_stock_check.isChecked())
        rows = self._filter_material_payload_rows(all_rows)
        selected_id = self.current_material_id or self._selected_material_id()
        self.table_count_label.setText(f"{len(rows)} de {len(all_rows)} registos")
        self.table.setSortingEnabled(False)
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(rows))
        for row_index, payload in enumerate(rows):
            values = payload["row"]
            state_label = self._stock_state_label(payload.get("record"), str(payload.get("severity", "ok")))
            comprimento = str(values.get("comprimento", "") or "").strip()
            largura = str(values.get("largura", "") or "").strip()
            dimensoes = " x ".join(value for value in (comprimento, largura) if value and value != "0") or "-"
            columns = [
                values["lote"],
                values.get("lote_fornecedor", ""),
                values["material"],
                dimensoes,
                values["largura"],
                values["espessura"],
                values["quantidade"],
                values["reservado"],
                values["formato"],
                values["metros"],
                values["peso_unid"],
                values["p_compra"],
                values["preco_unid"],
                values["disponivel"],
                values["tipo"],
                values["local"],
                values["id"],
                state_label,
            ]
            for col_index, value in enumerate(columns):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if col_index not in (0, 1, 2, 14, 15, 16):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, col_index, item)
            self._apply_row_colors(row_index, payload["severity"], payload["band"])
            self._apply_state_cell_style(row_index, str(payload.get("severity", "ok")), state_label)
            try:
                available_value = self.backend._parse_float(values.get("disponivel", 0), 0)
            except Exception:
                available_value = 0.0
            if available_value <= 0:
                zero_item = self.table.item(row_index, 13)
                if zero_item is not None:
                    zero_item.setForeground(QBrush(QColor("#b42318")))
        self.table.setUpdatesEnabled(True)
        self.table.setSortingEnabled(True)

        if selected_id:
            for row_index in range(self.table.rowCount()):
                item = self.table.item(row_index, 16)
                if item and item.text().strip() == selected_id:
                    self.table.selectRow(row_index)
                    self.current_material_id = selected_id
                    break
        elif self.table.rowCount() > 0:
            self.table.selectRow(0)
            self.on_selection_changed()
        else:
            self._set_detail_summary(None)
            self._refresh_form_state()
            self._refresh_price_preview()

    def on_selection_changed(self) -> None:
        material_id = self._selected_material_id()
        if not material_id:
            self.current_material_id = ""
            self.selection_hint.setText("Seleciona uma linha da tabela para ver o detalhe completo aqui.")
            self._set_detail_summary(None)
            return
        record = self.backend.material_by_id(material_id)
        if record is None:
            return
        self.current_material_id = material_id
        self.selection_hint.setText(
            f"Registo selecionado: {material_id} | {str(record.get('material', '') or '').strip() or 'Sem material'}"
        )
        self._set_detail_summary(record)
        self.formato_combo.setCurrentText(str(record.get("formato", "Chapa") or "Chapa"))
        self.material_combo.setCurrentText(str(record.get("material", "")))
        _set_material_family_combo(
            self.backend,
            self.material_family_combo,
            str(record.get("material_familia", "") or "").strip(),
            material=str(record.get("material", "") or "").strip(),
        )
        preview = self.backend.material_geometry_preview(record)
        self._set_section_options(str(record.get("formato", "Chapa") or "Chapa"), str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip())
        self.secao_tipo_combo.setCurrentText(str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip())
        self.espessura_combo.setCurrentText(str(record.get("espessura", "")))
        self.comprimento_edit.setText(self.backend._fmt(preview.get("comprimento", record.get("comprimento", 0))))
        self.largura_edit.setText(self.backend._fmt(preview.get("largura", record.get("largura", 0))))
        self.altura_edit.setText(self.backend._fmt(preview.get("altura", record.get("altura", 0))))
        self.diametro_edit.setText(self.backend._fmt(preview.get("diametro", record.get("diametro", 0))))
        self.contorno_edit.setText(self.backend.format_material_contour_points(record.get("contorno_points", record.get("shape_points", []))))
        self.metros_edit.setText(self.backend._fmt(preview.get("metros", record.get("metros", 0))))
        self.kg_m_edit.setText(self.backend._fmt(preview.get("kg_m", record.get("kg_m", 0))))
        self.peso_edit.setText(self.backend._fmt(preview.get("peso_unid", record.get("peso_unid", 0))))
        self.preco_compra_edit.setText(self.backend._fmt(record.get("p_compra", 0)))
        self.quantidade_edit.setText(self.backend._fmt(record.get("quantidade", 0)))
        self.reservado_edit.setText(self.backend._fmt(record.get("reservado", 0)))
        self.local_combo.setCurrentText(self.backend._localizacao(record))
        self.lote_interno_edit.setText(str(record.get("lote_interno", "")))
        self.lote_edit.setText(str(record.get("lote_fornecedor", "")))
        self._refresh_form_state()
        self._refresh_price_preview()

    def add_material(self) -> None:
        dialog = _MaterialEditorDialog(self.backend, self)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            record = self.backend.add_material(dialog.payload())
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.current_material_id = str(record.get("id", ""))
        self.refresh()

    def edit_material(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        if not material_id:
            QMessageBox.warning(self, "Aviso", "Seleciona um material primeiro.")
            return
        record = self.backend.material_by_id(material_id)
        if record is None:
            QMessageBox.warning(self, "Aviso", "O material selecionado já não existe.")
            return
        dialog = _MaterialEditorDialog(self.backend, self, record=record, mode="edit")
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            self.backend.update_material(material_id, dialog.payload())
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.current_material_id = material_id
        self.refresh()

    def remove_material(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        if not material_id:
            QMessageBox.warning(self, "Aviso", "Seleciona um material primeiro.")
            return
        if QMessageBox.question(self, "Confirmar", f"Remover {material_id}?") != QMessageBox.Yes:
            return
        try:
            self.backend.remove_material(material_id)
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self._set_form_defaults()
        self.refresh()

    def correct_material(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        record = self.backend.material_by_id(material_id)
        if record is None:
            QMessageBox.warning(self, "Aviso", "Seleciona um material primeiro.")
            return
        dlg = _SimpleFormDialog(
            "Corrigir Stock",
            [("quantidade", "Quantidade"), ("reservado", "Reservado"), ("metros", "Metros (m)")],
            {
                "quantidade": self.backend._fmt(record.get("quantidade", 0)),
                "reservado": self.backend._fmt(record.get("reservado", 0)),
                "metros": self.backend._fmt(record.get("metros", 0)),
            },
            self,
        )
        if dlg.exec() != QDialog.Accepted:
            return
        values = dlg.values()
        try:
            self.backend.correct_material_stock(material_id, values["quantidade"], values["reservado"], values["metros"])
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.current_material_id = material_id
        self.refresh()

    def _retalho_payload_has_data(self, payload: dict[str, str]) -> bool:
        return any(str(payload.get(key, "") or "").strip() for key in ("comprimento", "largura", "contorno_points", "quantidade", "metros"))

    def _consume_material_dialog(self, record: dict, *, include_baixa: bool = True) -> tuple[str, dict[str, str]] | None:
        preview = dict(self.backend.material_geometry_preview(record) or {})
        formato = str(preview.get("formato", record.get("formato", "")) or "").strip()
        is_chapa = formato == "Chapa"
        total_qty = max(0.0, self.backend._parse_float(record.get("quantidade", 0), 0))
        reserved_qty = max(0.0, self.backend._parse_float(record.get("reservado", 0), 0))
        available_qty = max(0.0, total_qty - reserved_qty)
        dialog = QDialog(self)
        dialog.setWindowTitle("Baixa de Material")
        dialog.setWindowFlags(
            dialog.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )
        dialog.setMinimumSize(680, 500 if is_chapa else 450)
        dialog.resize(760, 560 if is_chapa else 500)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(12)

        header = CardFrame()
        header.set_tone("info")
        header_layout = QGridLayout(header)
        header_layout.setContentsMargins(16, 14, 16, 14)
        header_layout.setHorizontalSpacing(18)
        header_layout.setVerticalSpacing(4)
        material_id = str(record.get("id", "") or "").strip()
        material_name = str(record.get("material", "") or "").strip()
        title = QLabel(f"{material_id} · {material_name}")
        title.setStyleSheet("font-size: 17px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel("Regista apenas o material efetivamente retirado. O retalho útil regressa ao stock com novo ID e lote.")
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        header_layout.addWidget(title, 0, 0, 1, 3)
        header_layout.addWidget(subtitle, 1, 0, 1, 3)
        for column, (label, value) in enumerate(
            (
                ("Formato", formato or "-"),
                ("Disponível", self.backend._fmt(available_qty)),
                ("Reservado", self.backend._fmt(reserved_qty)),
            )
        ):
            stat = QLabel(f"{label}\n{value}")
            stat.setStyleSheet(
                "margin-top: 6px; padding: 8px 10px; background: #ffffff;"
                "border: 1px solid #cbd8e6; border-radius: 6px; color: #334155; font-weight: 700;"
            )
            header_layout.addWidget(stat, 2, column)
        layout.addWidget(header)

        baixa_card = CardFrame()
        baixa_layout = QGridLayout(baixa_card)
        baixa_layout.setContentsMargins(16, 12, 16, 12)
        baixa_layout.setHorizontalSpacing(12)
        baixa_title = QLabel("Movimento de stock" if include_baixa else "Adicionar outro retalho")
        baixa_title.setStyleSheet("font-size: 14px; font-weight: 900; color: #0f172a;")
        baixa_layout.addWidget(baixa_title, 0, 0, 1, 2)
        qty_edit = QDoubleSpinBox()
        qty_edit.setRange(0.0, available_qty)
        qty_edit.setDecimals(2)
        qty_edit.setSingleStep(1.0)
        qty_edit.setValue(min(1.0, available_qty) if include_baixa else 0.0)
        qty_edit.setSuffix(" un")
        if include_baixa:
            baixa_layout.addWidget(QLabel("Quantidade a baixar"), 1, 0)
            baixa_layout.addWidget(qty_edit, 1, 1)
        else:
            qty_edit.hide()
        if include_baixa and available_qty <= 0:
            qty_edit.setEnabled(False)
            no_stock = QLabel("Não existe quantidade livre. O material está integralmente reservado.")
            no_stock.setProperty("role", "muted")
            baixa_layout.addWidget(no_stock, 2, 0, 1, 2)
        layout.addWidget(baixa_card)

        retalho_card = CardFrame()
        retalho_layout = QGridLayout(retalho_card)
        retalho_layout.setContentsMargins(16, 12, 16, 14)
        retalho_layout.setHorizontalSpacing(12)
        retalho_layout.setVerticalSpacing(9)
        retalho_box = QCheckBox("Registar retalho aproveitável")
        retalho_box.setChecked(not include_baixa)
        retalho_box.setStyleSheet("font-size: 14px; font-weight: 900; color: #0f172a;")
        retalho_layout.addWidget(retalho_box, 0, 0, 1, 4)
        retalho_hint = QLabel("Preenche as dimensões úteis que regressam ao stock.")
        retalho_hint.setProperty("role", "muted")
        retalho_layout.addWidget(retalho_hint, 1, 0, 1, 4)

        comp_edit = QDoubleSpinBox()
        comp_edit.setRange(0.0, 100000.0)
        comp_edit.setDecimals(1)
        comp_edit.setSuffix(" mm")
        comp_edit.setSpecialValueText("-")
        larg_edit = QDoubleSpinBox()
        larg_edit.setRange(0.0, 100000.0)
        larg_edit.setDecimals(1)
        larg_edit.setSuffix(" mm")
        larg_edit.setSpecialValueText("-")
        contorno_edit = QLineEdit()
        contorno_edit.setPlaceholderText("Opcional: 0,0; 1000,0; 900,400; 0,400")
        retalho_qty_edit = QDoubleSpinBox()
        retalho_qty_edit.setRange(0.0, 100000.0)
        retalho_qty_edit.setDecimals(2)
        retalho_qty_edit.setValue(1.0)
        retalho_qty_edit.setSuffix(" un")
        metros_edit = QDoubleSpinBox()
        metros_edit.setRange(0.0, 100000.0)
        metros_edit.setDecimals(3)
        metros_edit.setSuffix(" m")
        metros_edit.setSpecialValueText("-")
        if is_chapa:
            retalho_layout.addWidget(QLabel("Comprimento útil"), 2, 0)
            retalho_layout.addWidget(comp_edit, 2, 1)
            retalho_layout.addWidget(QLabel("Largura útil"), 2, 2)
            retalho_layout.addWidget(larg_edit, 2, 3)
            retalho_layout.addWidget(QLabel("Contorno irregular"), 3, 0)
            retalho_layout.addWidget(contorno_edit, 3, 1, 1, 3)
            retalho_layout.addWidget(QLabel("Quantidade"), 4, 0)
            retalho_layout.addWidget(retalho_qty_edit, 4, 1)
        else:
            retalho_layout.addWidget(QLabel("Comprimento útil"), 2, 0)
            retalho_layout.addWidget(metros_edit, 2, 1)
            retalho_layout.addWidget(QLabel("Quantidade"), 2, 2)
            retalho_layout.addWidget(retalho_qty_edit, 2, 3)

        retalho_fields = [comp_edit, larg_edit, contorno_edit, retalho_qty_edit, metros_edit]

        def sync_retalho_state() -> None:
            for widget in retalho_fields:
                widget.setEnabled(retalho_box.isChecked())

        retalho_box.toggled.connect(sync_retalho_state)
        sync_retalho_state()
        layout.addWidget(retalho_card)

        result_label = QLabel("")
        result_label.setStyleSheet(
            "padding: 9px 12px; background: #f8fafc; border: 1px solid #d6e0eb;"
            "border-radius: 6px; color: #334155; font-weight: 700;"
        )

        def update_result() -> None:
            remaining = max(0.0, total_qty - (qty_edit.value() if include_baixa else 0.0))
            text = f"Stock após movimento: {self.backend._fmt(remaining)} un"
            if retalho_box.isChecked() and is_chapa and comp_edit.value() > 0 and larg_edit.value() > 0:
                text += f"  |  Retalho: {self.backend._fmt(comp_edit.value())} × {self.backend._fmt(larg_edit.value())} mm"
            elif retalho_box.isChecked() and not is_chapa and metros_edit.value() > 0:
                text += f"  |  Retalho: {self.backend._fmt(metros_edit.value())} m"
            result_label.setText(text)

        for spin in (qty_edit, comp_edit, larg_edit, retalho_qty_edit, metros_edit):
            spin.valueChanged.connect(update_result)
        retalho_box.toggled.connect(update_result)
        update_result()
        layout.addWidget(result_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Confirmar movimento")

        def accept_dialog() -> None:
            if include_baixa and qty_edit.value() <= 0:
                QMessageBox.warning(dialog, "Baixa de material", "Indica a quantidade a baixar.")
                return
            if retalho_box.isChecked():
                if retalho_qty_edit.value() <= 0:
                    QMessageBox.warning(dialog, "Retalho", "Indica a quantidade do retalho.")
                    return
                if is_chapa and not contorno_edit.text().strip() and (comp_edit.value() <= 0 or larg_edit.value() <= 0):
                    QMessageBox.warning(dialog, "Retalho", "Indica comprimento e largura úteis ou define um contorno irregular.")
                    return
                if not is_chapa and metros_edit.value() <= 0:
                    QMessageBox.warning(dialog, "Retalho", "Indica o comprimento útil do retalho.")
                    return
            dialog.accept()

        buttons.accepted.connect(accept_dialog)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return None
        retalho = {}
        if retalho_box.isChecked():
            retalho = {
                "comprimento": self.backend._fmt(comp_edit.value()) if is_chapa else "",
                "largura": self.backend._fmt(larg_edit.value()) if is_chapa else "",
                "contorno_points": contorno_edit.text().strip() if is_chapa else "",
                "quantidade": self.backend._fmt(retalho_qty_edit.value()),
                "metros": "" if is_chapa else self.backend._fmt(metros_edit.value()),
            }
        return (self.backend._fmt(qty_edit.value()) if include_baixa else "0", retalho)

    def consume_material(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        if not material_id:
            QMessageBox.warning(self, "Aviso", "Seleciona um material primeiro.")
            return
        added_any = False
        include_baixa = True
        while True:
            record = self.backend.material_by_id(material_id)
            if record is None:
                QMessageBox.warning(self, "Aviso", "O material selecionado já não existe.")
                return
            result = self._consume_material_dialog(record, include_baixa=include_baixa)
            if result is None:
                break
            quantidade, retalho = result
            has_retalho = self._retalho_payload_has_data(retalho)
            try:
                self.backend.consume_material(material_id, quantidade, retalho if has_retalho else {})
            except Exception as exc:
                QMessageBox.critical(self, "Erro", str(exc))
                return
            added_any = True
            if not has_retalho:
                break
            if QMessageBox.question(self, "Retalhos", "Adicionar outro retalho sobrante deste material?") != QMessageBox.Yes:
                break
            include_baixa = False
        self.current_material_id = material_id
        if added_any:
            self.refresh()

    def open_full_grid(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Grelha de matéria-prima")
        dialog.resize(1500, 820)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        header = QHBoxLayout()
        title = QLabel("Stock de matéria-prima")
        title.setStyleSheet("font-size: 16px; font-weight: 900; color: #10253d;")
        info = QLabel(self.table_count_label.text())
        info.setProperty("role", "muted")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(info)
        layout.addLayout(header)

        search_box = QWidget()
        search_box.setObjectName("FullMaterialSearchBox")
        search_box.setStyleSheet(
            "QWidget#FullMaterialSearchBox { background: #ffffff; border: 1px solid #b8c9df; border-radius: 8px; }"
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
        search_edit.setPlaceholderText("Pesquisar ou picar MAT|... por formato, material, espessura, lote, dimensão...")
        search_layout.addWidget(search_icon)
        search_layout.addWidget(search_edit, 1)
        layout.addWidget(search_box)

        table = QTableWidget(0, self.table.columnCount())
        table.setHorizontalHeaderLabels(
            [
                "Formato",
                "Lote interno",
                "Lote fornecedor",
                "Material",
                "Dim. A",
                "Dim. B",
                "Espessura",
                "Quantidade",
                "Reserva",
                "Metros (m)",
                "Peso/Un. (kg)",
                "Compra (EUR)",
                "Preço/Unid.",
                "Disponível",
                "Tipo",
                "Localização",
                "ID",
                "Estado",
            ]
        )
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(False)
        table.setWordWrap(False)
        table.setStyleSheet(self.table.styleSheet())
        header_view = table.horizontalHeader()
        header_view.setMinimumSectionSize(48)
        full_grid_widths = [98, 134, 130, 260, 72, 72, 78, 86, 72, 92, 112, 104, 104, 104, 210, 132, 86, 118]
        for column, width in enumerate(full_grid_widths):
            header_view.setSectionResizeMode(column, QHeaderView.Interactive)
            table.setColumnWidth(column, width)
        table.setSortingEnabled(True)

        render_timer = QTimer(dialog)
        render_timer.setSingleShot(True)

        def row_columns(payload: dict) -> list[object]:
            values = payload["row"]
            state_label = self._stock_state_label(payload.get("record"), str(payload.get("severity", "ok")))
            return [
                values["formato"],
                values["lote"],
                values.get("lote_fornecedor", ""),
                values["material"],
                values["comprimento"],
                values["largura"],
                values["espessura"],
                values["quantidade"],
                values["reservado"],
                values["metros"],
                values["peso_unid"],
                values["p_compra"],
                values["preco_unid"],
                values["disponivel"],
                values["tipo"],
                values["local"],
                values["id"],
                state_label,
            ]

        def searchable_values(payload: dict) -> list[object]:
            record = dict(payload.get("record") or {})
            values = dict(payload.get("row") or {})
            raw_values = [
                record.get("formato", ""),
                record.get("material", ""),
                record.get("espessura", ""),
                record.get("comprimento", ""),
                record.get("largura", ""),
                record.get("diametro", ""),
                record.get("metros", ""),
                record.get("quantidade", ""),
                record.get("reservado", ""),
                record.get("peso_unid", ""),
                record.get("preco_unid", ""),
                record.get("p_compra", ""),
                record.get("lote_interno", ""),
                record.get("lote_fornecedor", ""),
                record.get("Localizacao", ""),
                record.get("Localização", ""),
                record.get("tipo", ""),
                record.get("categoria", ""),
                record.get("descricao", ""),
                record.get("observacoes", ""),
                record.get("obs", ""),
                record.get("referencia", ""),
                record.get("id", ""),
            ]
            raw_values.extend(row_columns(payload))
            raw_values.extend(values.values())
            return raw_values

        def render_full_grid() -> None:
            selected_id = ""
            selection = table.selectionModel()
            if selection is not None and selection.selectedRows():
                selected_item = table.item(selection.selectedRows()[0].row(), 16)
                selected_id = selected_item.text().strip() if selected_item is not None else ""
            selected_id = selected_id or self.current_material_id or self._selected_material_id()
            sort_col = header_view.sortIndicatorSection()
            sort_order = header_view.sortIndicatorOrder()
            table.setSortingEnabled(False)
            query = search_edit.text().strip()
            all_rows = self.backend.material_rows("", in_stock_only=self.only_stock_check.isChecked())
            rows = [payload for payload in all_rows if _grid_search_matches(searchable_values(payload), query)]
            info.setText(f"{len(rows)} registos")
            table.setRowCount(len(rows))
            selected_row = -1
            for row_index, payload in enumerate(rows):
                values = payload["row"]
                columns = row_columns(payload)
                for col_index, value in enumerate(columns):
                    item = QTableWidgetItem(str(value))
                    item.setToolTip(str(value))
                    if col_index not in (1, 2, 3, 14, 15, 16):
                        item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    table.setItem(row_index, col_index, item)
                if str(values.get("id", "") or "").strip() == selected_id:
                    selected_row = row_index
            table.setSortingEnabled(True)
            table.sortItems(sort_col, sort_order)
            if selected_id:
                for row_index in range(table.rowCount()):
                    item = table.item(row_index, 16)
                    if item is not None and item.text().strip() == selected_id:
                        table.selectRow(row_index)
                        break
            elif table.rowCount() > 0:
                table.selectRow(0)

        def select_scanned_material() -> None:
            value = search_edit.text().strip()
            try:
                result = self.backend.inventory_scan_lookup(value, expected_type="MAT")
            except ValueError as exc:
                if "|" in value or re.fullmatch(r"MAT[-_A-Z0-9]+", value, flags=re.IGNORECASE):
                    QMessageBox.warning(dialog, "Codigo de picagem", str(exc))
                return
            material_id = str(result.get("entity_id", "") or "").strip()
            self.current_material_id = material_id
            search_edit.setText(material_id)
            render_full_grid()
            for row_index in range(table.rowCount()):
                item = table.item(row_index, 16)
                if item is not None and item.text().strip().upper() == material_id.upper():
                    table.selectRow(row_index)
                    table.scrollToItem(item)
                    table.setFocus()
                    break

        render_timer.timeout.connect(render_full_grid)
        search_edit.textChanged.connect(lambda _text: render_timer.start(180))
        search_edit.returnPressed.connect(select_scanned_material)
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

        def selected_material_ids() -> list[str]:
            selection = table.selectionModel()
            if selection is None:
                return []
            result: list[str] = []
            for index in selection.selectedRows():
                item = table.item(index.row(), 16)
                material_id = item.text().strip() if item is not None else ""
                if material_id and material_id not in result:
                    result.append(material_id)
            return result

        def update_delete_state() -> None:
            delete_btn.setEnabled(bool(selected_material_ids()))

        def delete_selected_materials() -> None:
            material_ids = selected_material_ids()
            if not material_ids:
                QMessageBox.warning(dialog, "Apagar materiais", "Seleciona uma ou várias linhas.")
                return
            preview = ", ".join(material_ids[:6])
            if len(material_ids) > 6:
                preview += f" e mais {len(material_ids) - 6}"
            message = f"Apagar definitivamente {len(material_ids)} material(is)?\n\n{preview}"
            if QMessageBox.question(dialog, "Apagar materiais", message) != QMessageBox.Yes:
                return
            try:
                removed = self.backend.remove_materials(material_ids)
            except Exception as exc:
                QMessageBox.critical(dialog, "Apagar materiais", str(exc))
                return
            if self.current_material_id in material_ids:
                self.current_material_id = ""
                self._set_form_defaults()
            self.refresh()
            render_full_grid()
            QMessageBox.information(dialog, "Materiais", f"{removed} material(is) apagado(s).")

        table.itemSelectionChanged.connect(update_delete_state)
        delete_btn.clicked.connect(delete_selected_materials)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        selection = table.selectionModel()
        if selection is None or not selection.selectedRows():
            return
        row = selection.selectedRows()[0].row()
        item = table.item(row, 16)
        if item is None:
            return
        self.current_material_id = item.text().strip()
        self.refresh()

    def _open_weight_calculator(self) -> None:
        formato = self.formato_combo.currentText().strip() or "Chapa"
        profile = self.backend.material_family_profile(
            self.material_combo.currentText().strip(),
            str(self.material_family_combo.currentData() or "").strip(),
        )
        dialog = _WeightCalculatorDialog(
            {
                "formato": formato,
                "comprimento": float(self.backend._parse_float(self.comprimento_edit.text(), 0)),
                "largura": float(self.backend._parse_float(self.largura_edit.text(), 0)),
                "diametro": float(self.backend._parse_float(self.diametro_edit.text(), 0)),
                "espessura": float(self.backend._parse_float(self.espessura_combo.currentText(), 0)),
                "metros": float(self.backend._parse_float(self.metros_edit.text(), 0)),
                "densidade": float(profile.get("density", 7.85) or 7.85),
                "material_name": self.material_combo.currentText().strip(),
                "lote": self.lote_edit.text().strip(),
            },
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        self.peso_edit.setText(self.backend._fmt(values.get("peso_unid", 0)))
        if str(values.get("mode", "") or "") in {"Tubo", "Perfil"}:
            self.metros_edit.setText(self.backend._fmt(values.get("metros", 0)))
            self.kg_m_edit.setText(self.backend._fmt(values.get("kg_m", 0)))
        self._refresh_price_preview()

    def export_csv(self) -> None:
        default_path = str((Path.cwd() / "materiais_qt.csv").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Exportar CSV", default_path, "CSV (*.csv)")
        if not path:
            return
        try:
            target = self.backend.export_materials_csv(path, self.filter_edit.text())
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        QMessageBox.information(self, "CSV", f"Exportado para:\n{target}")

    def _select_pdf_formats(self) -> list[str] | None:
        formats = list(self.backend.material_stock_formats() or [])
        if not formats:
            QMessageBox.warning(self, "Preview PDF", "Não existem formatos de matéria-prima para apresentar.")
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle("Conteúdo do relatório de stock")
        dialog.setMinimumWidth(460)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        title = QLabel("Formatos incluídos no relatório")
        title.setStyleSheet("font-size: 15px; font-weight: 800; color: #10253d;")
        layout.addWidget(title)

        all_check = QCheckBox("Todos os formatos")
        all_check.setChecked(True)
        all_check.setStyleSheet("font-weight: 700;")
        layout.addWidget(all_check)

        format_grid = QGridLayout()
        format_grid.setHorizontalSpacing(16)
        format_grid.setVerticalSpacing(8)
        checks: list[QCheckBox] = []
        for index, value in enumerate(formats):
            check = QCheckBox(value)
            check.setChecked(True)
            checks.append(check)
            format_grid.addWidget(check, index // 3, index % 3)
        layout.addLayout(format_grid)

        stock_only_check = QCheckBox("Apenas linhas com stock disponível")
        stock_only_check.setChecked(self.only_stock_check.isChecked())
        layout.addWidget(stock_only_check)

        def set_all_formats(checked: bool) -> None:
            for check in checks:
                check.setChecked(checked)

        def update_all_state() -> None:
            all_selected = all(check.isChecked() for check in checks)
            all_check.blockSignals(True)
            all_check.setChecked(all_selected)
            all_check.blockSignals(False)

        all_check.toggled.connect(set_all_formats)
        for check in checks:
            check.toggled.connect(update_all_state)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Pré-visualizar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return None
        selected = [value for value, check in zip(formats, checks) if check.isChecked()]
        if not selected:
            QMessageBox.warning(self, "Preview PDF", "Seleciona pelo menos um formato.")
            return None
        self.only_stock_check.setChecked(stock_only_check.isChecked())
        return selected

    def preview_pdf(self) -> None:
        formats = self._select_pdf_formats()
        if formats is None:
            return
        try:
            self.backend.material_open_stock_pdf(
                in_stock_only=self.only_stock_check.isChecked(),
                formats=formats,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))

    def _build_material_label(self, output_path: str | None = None):
        record = self._selected_material_record()
        if not record:
            raise ValueError("Seleciona um material primeiro.")
        return self.backend.material_identification_label_pdf(str(record.get("id", "") or "").strip(), output_path=output_path)

    def preview_label(self) -> None:
        try:
            path = self._build_material_label()
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiqueta", str(exc))

    def print_label(self) -> None:
        try:
            path = self._build_material_label()
            try:
                os.startfile(str(path), "print")
            except Exception:
                os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "Etiqueta", str(exc))

    def save_label(self) -> None:
        record = self._selected_material_record()
        if not record:
            QMessageBox.warning(self, "Etiqueta", "Seleciona um material primeiro.")
            return
        material_id = str(record.get("id", "") or "").strip() or "material"
        path, _ = QFileDialog.getSaveFileName(self, "Guardar etiqueta", f"etiqueta_{material_id}.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            self._build_material_label(path)
        except Exception as exc:
            QMessageBox.critical(self, "Etiqueta", str(exc))
            return
        QMessageBox.information(self, "Etiqueta", f"PDF guardado em:\n{path}")

    def show_history(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        title = "Histórico de stock"
        if material_id:
            title = f"Histórico de stock | {material_id}"
        rows = self.backend.material_history_rows(material_id, limit=320)
        dlg = _HistoryDialog(title, rows, self.backend, self)
        dlg.exec()

