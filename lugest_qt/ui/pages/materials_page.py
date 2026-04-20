from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path
import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QCheckBox,
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
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame


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
        self.mode_combo.addItems(["Chapa", "Tubo", "Perfil"])
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

    def _apply_mode(self, mode: str) -> None:
        mode = str(mode or "Chapa").strip()
        show_sheet = mode == "Chapa"
        show_tube = mode == "Tubo"
        show_profile = mode == "Perfil"
        for key in ("sheet_length", "sheet_width", "sheet_thickness", "sheet_density", "sheet_weight", "sheet_area"):
            self._set_row_visible(key, show_sheet)
        for key in ("tube_shape", "tube_length", "tube_thickness", "tube_density", "tube_section", "tube_kgm", "tube_total"):
            self._set_row_visible(key, show_tube)
        for key in ("profile_series", "profile_size", "profile_length", "profile_kgm", "profile_manual", "profile_weight_m", "profile_total"):
            self._set_row_visible(key, show_profile)
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
        length = float(self.length_spin.value() or 0)
        width = float(self.width_spin.value() or 0)
        thickness = float(self.thickness_spin.value() or 0)
        density = float(self.density_spin.value() or 0)
        return {"mode": mode, "peso_unid": round((length * width * thickness * density) / 1000000.0, 4), "metros": round((length * width) / 1000000.0, 4), "kg_m": 0.0}


class _HistoryDialog(QDialog):
    def __init__(self, title: str, rows: list[dict[str, str]], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(980, 560)
        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Data", "Acao", "Detalhes"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.table)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)
        self.set_rows(rows)

    def set_rows(self, rows: list[dict[str, str]]) -> None:
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, key in enumerate(("data", "acao", "detalhes")):
                item = QTableWidgetItem(str(row.get(key, "") or ""))
                if col_index < 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, col_index, item)


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
            ("Lote", self.lote_edit),
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
        espessura_required = formato in {"Chapa", "Tubo"}
        esp_label = self._field_labels.get("Espessura")
        if esp_label is not None:
            esp_label.setText("Espessura" if espessura_required else "Espessura (opc.)")
        secao_label = self._field_labels.get("Tipo secção")
        if secao_label is not None:
            secao_label.setText("Tipo tubo" if formato == "Tubo" else ("Tipo perfil / série" if formato == "Perfil" else "Tipo secção"))
        comp_label = self._field_labels.get("Comprimento")
        if comp_label is not None:
            if formato == "Chapa":
                comp_label.setText("Comprimento (mm)")
            elif formato == "Tubo":
                comp_label.setText("Lado A (mm)")
            else:
                comp_label.setText("Comprimento")
        larg_label = self._field_labels.get("Largura")
        if larg_label is not None:
            larg_label.setText("Largura (mm)" if formato == "Chapa" else "Lado B (mm)")
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
            esp_line.setPlaceholderText("Obrigatória" if espessura_required else "Opcional para perfil")
        self.espessura_combo.setToolTip("Obrigatória para chapa e tubo." if espessura_required else "Opcional em perfis.")
        self.preco_compra_edit.setPlaceholderText("EUR/m" if formato == "Tubo" else "EUR/kg")
        self.preco_compra_edit.setToolTip("Preço de compra base por metro." if formato == "Tubo" else "Preço de compra base por kg.")
        self.metros_edit.setPlaceholderText("Comprimento barra (m)" if formato in {"Tubo", "Perfil"} else "")
        self.peso_edit.setPlaceholderText("Peso calculado automaticamente")
        self.kg_m_edit.setReadOnly(formato == "Tubo" or profile_catalog)
        self._set_field_visible("Tipo secção", formato in {"Tubo", "Perfil"})
        self._set_field_visible("Comprimento", formato == "Chapa" or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Largura", formato == "Chapa" or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Altura", formato == "Perfil")
        self._set_field_visible("Diâmetro", formato == "Tubo" and tube_round)
        self._set_field_visible("Contorno retalho", formato == "Chapa")
        self._set_field_visible("Metros", formato != "Chapa")
        self._set_field_visible("Kg/m", formato in {"Tubo", "Perfil"})

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
        if str(preview.get("formato", "") or "") in {"Tubo", "Perfil"}:
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
    page_subtitle = "Stock, retalhos, preços e reservas com atualização direta sobre a base atual."
    uses_backend_reload = True

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.current_material_id = ""
        self._combo_keys = ("formato", "material", "material_familia", "espessura", "local")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        form_card = CardFrame()
        form_card.setStyleSheet(
            "QLineEdit:disabled, QComboBox:disabled {"
            " background: #f8fbff;"
            " color: #0f172a;"
            " border: 1px solid #d6e3f3;"
            " border-radius: 8px;"
            "}"
        )
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 12)
        form_layout.setSpacing(8)

        top = QHBoxLayout()
        title = QLabel("Gestão de Stock")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel("Consulta rápida do stock. Adicionar e editar abrem um quadro próprio, separado da lista.")
        subtitle.setProperty("role", "muted")
        ttl_wrap = QVBoxLayout()
        ttl_wrap.setContentsMargins(0, 0, 0, 0)
        ttl_wrap.setSpacing(2)
        ttl_wrap.addWidget(title)
        ttl_wrap.addWidget(subtitle)
        top.addLayout(ttl_wrap, 1)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Pesquisar material, lote, dimensão, local...")
        self.filter_edit.setMaximumWidth(280)
        self.filter_edit.textChanged.connect(self.refresh)
        top.addWidget(self.filter_edit)
        form_layout.addLayout(top)

        info_row = QHBoxLayout()
        info_row.setSpacing(10)
        self.selection_hint = QLabel("Seleciona uma linha da tabela para ver o detalhe completo aqui.")
        self.selection_hint.setStyleSheet("font-size: 12px; color: #1d4ed8; font-weight: 600;")
        self.stock_hint = QLabel("Stock baixo deixa de pintar a linha toda e passa a ser assinalado só em Disponível.")
        self.stock_hint.setProperty("role", "muted")
        info_row.addWidget(self.selection_hint, 1)
        info_row.addWidget(self.stock_hint)
        form_layout.addLayout(info_row)

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
        form_layout.addWidget(detail_card)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        self.formato_combo = self._make_combo()
        self.material_combo = self._make_combo()
        self.material_family_combo = QComboBox()
        self.secao_tipo_combo = self._make_combo()
        self.espessura_combo = self._make_combo()
        self.local_combo = self._make_combo()
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
            ("Lote", self.lote_edit),
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
        form_layout.addLayout(grid)

        actions_primary = QHBoxLayout()
        actions_primary.setSpacing(8)
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
        self.correct_btn = QPushButton("Corrigir stock")
        self.correct_btn.setProperty("variant", "secondary")
        self.correct_btn.clicked.connect(self.correct_material)
        self.history_btn = QPushButton("Histórico")
        self.history_btn.setProperty("variant", "secondary")
        self.history_btn.clicked.connect(self.show_history)
        self.label_btn = QPushButton("Etiqueta ID")
        self.label_btn.setProperty("variant", "secondary")
        self.label_btn.clicked.connect(self.preview_label)
        self.label_print_btn = QPushButton("Imprimir etiqueta")
        self.label_print_btn.setProperty("variant", "secondary")
        self.label_print_btn.clicked.connect(self.print_label)
        self.label_save_btn = QPushButton("Guardar etiqueta")
        self.label_save_btn.setProperty("variant", "secondary")
        self.label_save_btn.clicked.connect(self.save_label)
        self.pdf_btn = QPushButton("Pre-visualizar PDF")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self.preview_pdf)
        self.pdf_save_btn = QPushButton("Guardar PDF")
        self.pdf_save_btn.setProperty("variant", "secondary")
        self.pdf_save_btn.clicked.connect(self.save_pdf)
        self.calc_btn = QPushButton("Calc. peso")
        self.calc_btn.setProperty("variant", "secondary")
        self.calc_btn.clicked.connect(self._open_weight_calculator)
        self.refresh_btn = QPushButton("Atualizar")
        self.refresh_btn.setProperty("variant", "secondary")
        self.refresh_btn.clicked.connect(self.refresh)
        self.export_btn = QPushButton("Exportar CSV")
        self.export_btn.setProperty("variant", "secondary")
        self.export_btn.clicked.connect(self.export_csv)
        for button in (
            self.add_btn,
            self.edit_btn,
            self.remove_btn,
            self.baixa_btn,
            self.correct_btn,
        ):
            actions_primary.addWidget(button)
        actions_primary.addStretch(1)
        form_layout.addLayout(actions_primary)

        actions_secondary = QHBoxLayout()
        actions_secondary.setSpacing(8)
        for button in (
            self.history_btn,
            self.label_btn,
            self.label_print_btn,
            self.label_save_btn,
            self.pdf_btn,
            self.pdf_save_btn,
            self.calc_btn,
            self.refresh_btn,
            self.export_btn,
        ):
            actions_secondary.addWidget(button)
        actions_secondary.addStretch(1)
        form_layout.addLayout(actions_secondary)
        root.addWidget(form_card)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(16, 14, 16, 14)
        table_layout.setSpacing(10)
        self.table = QTableWidget(0, 16)
        self.table.setStyleSheet(
            "QTableWidget {"
            " gridline-color: #d8e3f2;"
            " selection-background-color: #dbeafe;"
            " selection-color: #0f172a;"
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
                "Lote",
                "Material",
                "Dim. A",
                "Dim. B",
                "Espessura",
                "Quantidade",
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
            ]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.setWordWrap(False)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.itemDoubleClicked.connect(lambda *_args: self.edit_material())
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in range(2, 16):
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        table_layout.addWidget(self.table)
        root.addWidget(table_card, 1)

        for combo in (self.formato_combo, self.material_combo, self.secao_tipo_combo, self.espessura_combo, self.local_combo):
            combo.currentTextChanged.connect(self._on_form_value_changed)
        self.material_family_combo.currentIndexChanged.connect(self._on_form_value_changed)
        for edit in (
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
            return
        material_id = str(record.get("id", "") or "").strip()
        material = str(record.get("material", "") or "").strip() or "Sem material"
        formato = str(record.get("formato", "") or "").strip() or "-"
        local = self.backend._localizacao(record) or "Sem localização"
        lote = str(record.get("lote_fornecedor", "") or "").strip() or "Sem lote"
        quantidade = self.backend._parse_float(record.get("quantidade", 0), 0)
        reservado = self.backend._parse_float(record.get("reservado", 0), 0)
        disponivel = quantidade - reservado
        self.detail_title.setText(f"{material} | {material_id}")
        self.detail_meta.setText(f"{formato} | Lote: {lote} | Localização: {local}")
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
            f"Disponível: {self.backend._fmt(disponivel)}  |  Qtd: {self.backend._fmt(quantidade)}"
        )
        self.detail_available.setStyleSheet(
            "background: white; color: #0f172a; border: 1px solid #d6e3f3; "
            "border-radius: 10px; padding: 6px 10px; font-weight: 700;"
        )

    def _set_preview_fields_enabled(self, enabled: bool) -> None:
        preview_only = (
            self.formato_combo,
            self.material_combo,
            self.material_family_combo,
            self.secao_tipo_combo,
            self.espessura_combo,
            self.local_combo,
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
        espessura_required = formato in {"Chapa", "Tubo"}
        esp_label = self._field_labels.get("Espessura")
        if esp_label is not None:
            esp_label.setText("Espessura" if espessura_required else "Espessura (opc.)")
        secao_label = self._field_labels.get("Tipo secção")
        if secao_label is not None:
            secao_label.setText("Tipo tubo" if formato == "Tubo" else ("Tipo perfil / série" if formato == "Perfil" else "Tipo secção"))
        comp_label = self._field_labels.get("Comprimento")
        if comp_label is not None:
            if formato == "Chapa":
                comp_label.setText("Comprimento (mm)")
            elif formato == "Tubo":
                comp_label.setText("Lado A (mm)")
            else:
                comp_label.setText("Comprimento")
        larg_label = self._field_labels.get("Largura")
        if larg_label is not None:
            larg_label.setText("Largura (mm)" if formato == "Chapa" else "Lado B (mm)")
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
            esp_line.setPlaceholderText("Obrigatória" if espessura_required else "Opcional para perfil")
        self.espessura_combo.setToolTip("Obrigatória para chapa e tubo." if espessura_required else "Opcional em perfis.")
        self.preco_compra_edit.setPlaceholderText("EUR/m" if formato == "Tubo" else "EUR/kg")
        self.preco_compra_edit.setToolTip("Preço de compra base por metro." if formato == "Tubo" else "Preço de compra base por kg.")
        self.metros_edit.setPlaceholderText("Comprimento barra (m)" if formato in {"Tubo", "Perfil"} else "")
        self.peso_edit.setPlaceholderText("Peso calculado automaticamente")
        self.kg_m_edit.setReadOnly(formato == "Tubo" or profile_catalog)
        self._set_field_visible("Tipo secção", formato in {"Tubo", "Perfil"})
        self._set_field_visible("Comprimento", formato == "Chapa" or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Largura", formato == "Chapa" or (formato == "Tubo" and not tube_round))
        self._set_field_visible("Altura", formato == "Perfil")
        self._set_field_visible("Diâmetro", formato == "Tubo" and tube_round)
        self._set_field_visible("Contorno retalho", formato == "Chapa")
        self._set_field_visible("Metros", formato != "Chapa")
        self._set_field_visible("Kg/m", formato in {"Tubo", "Perfil"})

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
        if str(preview.get("formato", "") or "") in {"Tubo", "Perfil"}:
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

    def _apply_row_colors(self, row_index: int, severity: str, band: str) -> None:
        background = QColor("#eef2f8" if band == "even" else "#e6ecf5")
        foreground = QColor("#0f172a")
        for col in range(self.table.columnCount()):
            item = self.table.item(row_index, col)
            if item is None:
                continue
            item.setBackground(QBrush(background))
            item.setForeground(QBrush(foreground))
        available_item = self.table.item(row_index, 12)
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

    def _selected_material_id(self) -> str:
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return ""
        rows = selection_model.selectedRows()
        if not rows:
            return ""
        item = self.table.item(rows[0].row(), 15)
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
        self._refresh_form_state()
        self._refresh_price_preview()

        rows = self.backend.material_rows(self.filter_edit.text())
        selected_id = self.current_material_id or self._selected_material_id()
        self.table.setSortingEnabled(False)
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(rows))
        for row_index, payload in enumerate(rows):
            values = payload["row"]
            columns = [
                values["lote"],
                values["material"],
                values["comprimento"],
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
            ]
            for col_index, value in enumerate(columns):
                item = QTableWidgetItem(str(value))
                if col_index not in (0, 1, 13, 14, 15):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.table.setItem(row_index, col_index, item)
            self._apply_row_colors(row_index, payload["severity"], payload["band"])
            try:
                available_value = float(values.get("disponivel", 0) or 0)
            except Exception:
                available_value = 0.0
            if available_value <= 0:
                zero_item = self.table.item(row_index, 12)
                if zero_item is not None:
                    zero_item.setForeground(QBrush(QColor("#b42318")))
        self.table.setUpdatesEnabled(True)
        self.table.setSortingEnabled(True)

        if selected_id:
            for row_index in range(self.table.rowCount()):
                item = self.table.item(row_index, 15)
                if item and item.text().strip() == selected_id:
                    self.table.selectRow(row_index)
                    self.current_material_id = selected_id
                    break
        else:
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

    def consume_material(self) -> None:
        material_id = self.current_material_id or self._selected_material_id()
        if not material_id:
            QMessageBox.warning(self, "Aviso", "Seleciona um material primeiro.")
            return
        dlg = _SimpleFormDialog(
            "Baixa de Material",
            [
                ("quantidade", "Quantidade a baixar"),
                ("comprimento", "Retalho comprimento"),
                ("largura", "Retalho largura"),
                ("contorno_points", "Retalho contorno"),
                ("quantidade_retalho", "Retalho quantidade"),
                ("metros", "Retalho metros"),
            ],
            {"quantidade": "", "comprimento": "", "largura": "", "contorno_points": "", "quantidade_retalho": "", "metros": ""},
            self,
        )
        if dlg.exec() != QDialog.Accepted:
            return
        values = dlg.values()
        retalho = {
            "comprimento": values["comprimento"],
            "largura": values["largura"],
            "contorno_points": values["contorno_points"],
            "quantidade": values["quantidade_retalho"],
            "metros": values["metros"],
        }
        try:
            self.backend.consume_material(material_id, values["quantidade"], retalho)
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.current_material_id = material_id
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

    def preview_pdf(self) -> None:
        try:
            self.backend.material_open_stock_pdf()
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))

    def save_pdf(self) -> None:
        default_path = str((Path.cwd() / f"stock_materia_prima_{datetime.now().strftime('%Y%m%d')}.pdf").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", default_path, "PDF (*.pdf)")
        if not path:
            return
        try:
            target = self.backend.material_render_stock_pdf(path)
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        QMessageBox.information(self, "PDF", f"PDF guardado em:\n{target}")

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
        dlg = _HistoryDialog(title, rows, self)
        dlg.exec()

