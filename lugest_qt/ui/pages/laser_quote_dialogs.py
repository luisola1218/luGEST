from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
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
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame
from lugest_core.laser.quote_engine import default_laser_quote_settings

MATERIAL_FAMILY_ALIASES = {
    "FERRO": "Aco carbono",
    "ACO": "Aco carbono",
    "ACOCARBONO": "Aco carbono",
    "CARBONSTEEL": "Aco carbono",
    "INOX": "Aco inox",
    "ACOINOX": "Aco inox",
    "STAINLESS": "Aco inox",
    "ALUMINIO": "Aluminio",
    "ALUMINUM": "Aluminio",
    "ALU": "Aluminio",
}
MATERIAL_FAMILY_LABELS = {
    "Aco carbono": "Ferro",
    "Aco inox": "INOX",
    "Aluminio": "Aluminio",
}


def _fmt_eur(value: Any) -> str:
    try:
        amount = float(value or 0)
    except Exception:
        amount = 0.0
    return f"{amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_num(value: Any, digits: int = 2, suffix: str = "") -> str:
    try:
        amount = float(value or 0)
    except Exception:
        amount = 0.0
    text = f"{amount:.{digits}f}".rstrip("0").rstrip(".")
    text = text.replace(".", ",")
    return f"{text}{suffix}"


def _combo_items(combo: QComboBox) -> list[str]:
    return [combo.itemText(index) for index in range(combo.count())]


def _set_combo_values(combo: QComboBox, values: list[str], current: str = "") -> None:
    combo.blockSignals(True)
    combo.clear()
    for value in values:
        combo.addItem(str(value))
    if current:
        combo.setCurrentText(current)
    combo.blockSignals(False)


def _table_number_item(value: Any, digits: int = 2) -> QTableWidgetItem:
    item = QTableWidgetItem(_fmt_num(value, digits))
    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
    return item


def _table_text_item(value: Any) -> QTableWidgetItem:
    item = QTableWidgetItem(str(value or "").strip())
    item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
    return item


def _spin(decimals: int, min_value: float, max_value: float, value: float, step: float = 1.0) -> QDoubleSpinBox:
    widget = QDoubleSpinBox()
    widget.setDecimals(decimals)
    widget.setRange(min_value, max_value)
    widget.setSingleStep(step)
    widget.setValue(value)
    return widget


def _densify_editor(widget: QWidget, *, height: int = 28) -> QWidget:
    try:
        widget.setProperty("compact", "true")
    except Exception:
        pass
    try:
        widget.setMinimumHeight(height)
        widget.setMaximumHeight(max(height, widget.maximumHeight()))
    except Exception:
        pass
    return widget


def _settings_material_names(settings: dict[str, Any], machine_name: str) -> list[str]:
    machines = dict(settings.get("machine_profiles", {}) or {})
    machine = dict(machines.get(machine_name, {}) or {})
    materials = dict(machine.get("materials", {}) or {})
    values: list[str] = []
    for key in materials.keys():
        display = _display_material_family(key)
        if display and display not in values:
            values.append(display)
    return values or ["Ferro"]


def _settings_gas_names(settings: dict[str, Any], machine_name: str, material_name: str) -> list[str]:
    machines = dict(settings.get("machine_profiles", {}) or {})
    machine = dict(machines.get(machine_name, {}) or {})
    materials = dict(machine.get("materials", {}) or {})
    material = _resolve_material_mapping(materials, material_name)
    gases = dict(material.get("gases", {}) or {})
    forced = _forced_gas_for_material(material_name)
    if forced and forced in gases:
        return [forced]
    return list(gases.keys()) or [forced or "Oxigenio"]


def _norm_material_token(value: Any) -> str:
    return "".join(char for char in str(value or "").upper() if char.isalnum())


def _guess_material_family(value: Any) -> str:
    token = _norm_material_token(value)
    if not token:
        return ""
    inox_markers = ("INOX", "AISI", "14301", "14307", "14404", "14571", "304", "304L", "316", "316L", "430")
    aluminio_markers = ("ALUMINIO", "ALUMINUM", "ALU", "5083", "5754", "6082")
    carbono_markers = ("ACO", "ACOCARBONO", "S235", "S275", "S355", "CORTEN", "HARDOX", "DX51", "GALV")
    if any(marker in token for marker in inox_markers):
        return "INOX"
    if any(marker in token for marker in aluminio_markers):
        return "Aluminio"
    if any(marker in token for marker in carbono_markers):
        return "Ferro"
    return ""


def _canonical_material_family(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    token = _norm_material_token(raw)
    return MATERIAL_FAMILY_ALIASES.get(token, raw)


def _display_material_family(value: Any) -> str:
    canonical = _canonical_material_family(value)
    if not canonical:
        return ""
    return MATERIAL_FAMILY_LABELS.get(canonical, canonical)


def _forced_gas_for_material(material_name: Any) -> str:
    canonical = _canonical_material_family(material_name)
    if canonical == "Aco inox":
        return "Nitrogenio"
    return ""


def _preferred_material_gas(settings: dict[str, Any], machine_name: str, material_name: str, fallback: str = "") -> str:
    forced = _forced_gas_for_material(material_name)
    values = _settings_gas_names(settings, machine_name, material_name)
    if forced and forced in values:
        return forced
    if fallback and fallback in values:
        return fallback
    machines = dict(settings.get("machine_profiles", {}) or {})
    machine = dict(machines.get(machine_name, {}) or {})
    materials = dict(machine.get("materials", {}) or {})
    material = _resolve_material_mapping(materials, material_name)
    default_gas = str(material.get("default_gas", "") or "").strip()
    if default_gas and default_gas in values:
        return default_gas
    return values[0] if values else (forced or "Oxigenio")


def _material_catalog_metrics(family_catalog: dict[str, Any]) -> tuple[int, int]:
    subtype_count = 0
    rule_count = 0
    for subtype, payload in dict(family_catalog or {}).items():
        if not str(subtype or "").strip():
            continue
        subtype_count += 1
        rule_count += 1
        for row in list(dict(payload or {}).get("thickness_overrides", []) or []):
            if isinstance(row, dict):
                rule_count += 1
    return subtype_count, rule_count


def _cut_rows_summary(rows: list[dict[str, Any]]) -> str:
    valid = [
        dict(row or {})
        for row in list(rows or [])
        if float(str((row or {}).get("thickness_mm", 0) or 0).replace(",", ".")) > 0.0
    ]
    if not valid:
        return "Sem linhas de corte configuradas."
    values = sorted(float(row.get("thickness_mm", 0) or 0) for row in valid)
    return f"{len(valid)} linhas | espessuras {_fmt_num(values[0], 1)} a {_fmt_num(values[-1], 1)} mm"


class CutTableDialog(QDialog):
    def __init__(self, material_name: str, gas_name: str, rows: list[dict[str, Any]], parent=None) -> None:
        super().__init__(parent)
        self.material_name = str(material_name or "").strip() or "Material"
        self.gas_name = str(gas_name or "").strip() or "Gas"
        self.setWindowTitle(f"Tabela de corte - {self.material_name} / {self.gas_name}")
        self.resize(1180, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        header_card = CardFrame()
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(6)
        title = QLabel(f"Tabela de corte para {self.material_name} / {self.gas_name}")
        title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        helper = QLabel(
            "Edita aqui a tabela completa por espessura. Cada linha representa um set de parametros de corte."
        )
        helper.setProperty("role", "muted")
        self.summary_label = QLabel("")
        self.summary_label.setProperty("role", "muted")
        header_layout.addWidget(title)
        header_layout.addWidget(helper)
        header_layout.addWidget(self.summary_label)
        root.addWidget(header_card)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(12, 10, 12, 10)
        table_layout.setSpacing(8)
        self.table = QTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels(["Esp.", "Vel min", "Vel max", "Bico dist.", "Gas min", "Gas max", "Foco", "Bico", "Duty %", "Freq Hz", "Potencia"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table_layout.addWidget(self.table, 1)
        actions = QHBoxLayout()
        self.add_btn = QPushButton("Adicionar linha")
        self.remove_btn = QPushButton("Remover linha")
        self.remove_btn.setProperty("variant", "danger")
        actions.addWidget(self.add_btn)
        actions.addWidget(self.remove_btn)
        actions.addStretch(1)
        table_layout.addLayout(actions)
        root.addWidget(table_card, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self.add_btn.clicked.connect(self._add_row)
        self.remove_btn.clicked.connect(self._remove_row)
        self.table.itemChanged.connect(lambda *_args: self._update_summary())
        self._load_rows(rows)

    def _load_rows(self, rows: list[dict[str, Any]]) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            self.table.setItem(row_index, 0, _table_number_item(row.get("thickness_mm", 0), 1))
            self.table.setItem(row_index, 1, _table_number_item(row.get("speed_min_m_min", 0), 2))
            self.table.setItem(row_index, 2, _table_number_item(row.get("speed_max_m_min", 0), 2))
            self.table.setItem(row_index, 3, _table_number_item(row.get("nozzle_distance_mm", 0), 2))
            self.table.setItem(row_index, 4, _table_number_item(row.get("gas_pressure_bar_min", 0), 2))
            self.table.setItem(row_index, 5, _table_number_item(row.get("gas_pressure_bar_max", 0), 2))
            self.table.setItem(row_index, 6, _table_number_item(row.get("focus_mm", 0), 2))
            self.table.setItem(row_index, 7, _table_text_item(row.get("nozzle", "")))
            self.table.setItem(row_index, 8, _table_number_item(row.get("duty_pct", 0), 1))
            self.table.setItem(row_index, 9, _table_number_item(row.get("frequency_hz", 0), 0))
            self.table.setItem(row_index, 10, _table_number_item(row.get("power_w", row.get("power_pct", 0)), 0))
        self.table.blockSignals(False)
        self._update_summary()

    def _update_summary(self) -> None:
        self.summary_label.setText(_cut_rows_summary(self.rows_payload()))

    def _add_row(self) -> None:
        row_index = self.table.rowCount()
        self.table.insertRow(row_index)
        for col_index in range(self.table.columnCount()):
            if col_index == 7:
                self.table.setItem(row_index, col_index, _table_text_item(""))
            else:
                self.table.setItem(row_index, col_index, _table_number_item(0, 2 if col_index else 1))
        self._update_summary()

    def _remove_row(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = self.table.rowCount() - 1
        if row >= 0:
            self.table.removeRow(row)
        self._update_summary()

    def rows_payload(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for row_index in range(self.table.rowCount()):
            numeric_values: list[float] = []
            for col_index in (0, 1, 2, 3, 4, 5, 6, 8, 9, 10):
                item = self.table.item(row_index, col_index)
                text = str(item.text() if item is not None else "").strip().replace(",", ".")
                numeric_values.append(float(text or 0.0))
            if numeric_values[0] <= 0.0:
                continue
            nozzle_item = self.table.item(row_index, 7)
            rows.append(
                {
                    "thickness_mm": numeric_values[0],
                    "speed_min_m_min": numeric_values[1],
                    "speed_max_m_min": numeric_values[2],
                    "nozzle_distance_mm": numeric_values[3],
                    "gas_pressure_bar_min": numeric_values[4],
                    "gas_pressure_bar_max": numeric_values[5],
                    "focus_mm": numeric_values[6],
                    "nozzle": str(nozzle_item.text() if nozzle_item is not None else "").strip(),
                    "duty_pct": numeric_values[7],
                    "frequency_hz": numeric_values[8],
                    "power_w": numeric_values[9],
                }
            )
        rows.sort(key=lambda row: float(row.get("thickness_mm", 0) or 0))
        return rows


def _resolve_material_mapping(materials: dict[str, Any], material_name: str) -> dict[str, Any]:
    family = _canonical_material_family(material_name) or str(material_name or "").strip()
    if family in materials:
        return dict(materials.get(family, {}) or {})
    family_token = _norm_material_token(family)
    for key, value in dict(materials or {}).items():
        if _norm_material_token(key) == family_token:
            return dict(value or {})
    return {}


def _settings_material_subtypes(settings: dict[str, Any], material_name: str, extra_values: list[str] | None = None) -> list[str]:
    family = _canonical_material_family(material_name) or str(material_name or "").strip()
    configured = list(dict(settings.get("material_subtypes", {}) or {}).get(family, []) or [])
    catalog_values: list[str] = []
    for profile in dict(settings.get("commercial_profiles", {}) or {}).values():
        if not isinstance(profile, dict):
            continue
        material_catalog = dict(profile.get("material_catalog", {}) or {})
        family_catalog = dict(material_catalog.get(family, {}) or {})
        for subtype in family_catalog.keys():
            clean = str(subtype or "").strip()
            if clean and clean not in catalog_values:
                catalog_values.append(clean)
    values = configured + catalog_values + list(extra_values or [])
    out: list[str] = []
    for value in values:
        clean = str(value or "").strip()
        if clean in out:
            continue
        if not clean and out:
            continue
        out.append(clean)
    return out


class MaterialSubtypeCatalogDialog(QDialog):
    def __init__(self, material_name: str, family_catalog: dict[str, Any], fallback_price_per_kg: float, fallback_scrap_credit_per_kg: float, parent=None) -> None:
        super().__init__(parent)
        self.material_name = str(material_name or "").strip() or "Material"
        self.fallback_price_per_kg = float(fallback_price_per_kg or 0.0)
        self.fallback_scrap_credit_per_kg = float(fallback_scrap_credit_per_kg or 0.0)
        self.setWindowTitle(f"Tabela de subtipos - {self.material_name}")
        self.resize(980, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        header_card = CardFrame()
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(6)
        title = QLabel(f"Subtipos / categorias de {self.material_name}")
        title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        note = QLabel(
            f"Preco base da familia: {_fmt_num(self.fallback_price_per_kg, 3)} EUR/kg | "
            f"Sucata base: {_fmt_num(self.fallback_scrap_credit_per_kg, 3)} EUR/kg"
        )
        note.setProperty("role", "muted")
        helper = QLabel("Cada linha sobrepoe o preco base da familia quando esse subtipo for escolhido no orcamento.")
        helper.setProperty("role", "muted")
        header_layout.addWidget(title)
        header_layout.addWidget(note)
        header_layout.addWidget(helper)
        root.addWidget(header_card)

        filter_card = CardFrame()
        filter_layout = QHBoxLayout(filter_card)
        filter_layout.setContentsMargins(12, 8, 12, 8)
        filter_layout.setSpacing(8)
        filter_layout.addWidget(QLabel("Procurar"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Filtrar por subtipo ou espessura")
        self.count_label = QLabel("")
        self.count_label.setProperty("role", "muted")
        filter_layout.addWidget(self.search_edit, 1)
        filter_layout.addWidget(self.count_label)
        root.addWidget(filter_card)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(12, 10, 12, 10)
        table_layout.setSpacing(8)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Subtipo", "Esp. min (mm)", "Esp. max (mm)", "Preco material (EUR/kg)", "Densidade (kg/m3)", "Valor sucata (EUR/kg)"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for section in (1, 2, 3, 4, 5):
            header.setSectionResizeMode(section, QHeaderView.ResizeToContents)
        table_layout.addWidget(self.table, 1)
        table_actions = QHBoxLayout()
        self.add_btn = QPushButton("Adicionar regra")
        self.remove_btn = QPushButton("Remover regra")
        self.remove_btn.setProperty("variant", "secondary")
        table_actions.addWidget(self.add_btn)
        table_actions.addWidget(self.remove_btn)
        table_actions.addStretch(1)
        table_layout.addLayout(table_actions)
        root.addWidget(table_card, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self.add_btn.clicked.connect(self._add_row)
        self.remove_btn.clicked.connect(self._remove_row)
        self.search_edit.textChanged.connect(self._apply_filter)
        self._load_rows(family_catalog)

    def _load_rows(self, family_catalog: dict[str, Any]) -> None:
        rows: list[tuple[str, dict[str, Any], bool]] = []
        for subtype, payload in sorted(
            [
                (str(subtype or "").strip(), dict(payload or {}))
                for subtype, payload in dict(family_catalog or {}).items()
                if str(subtype or "").strip()
            ],
            key=lambda item: item[0].upper(),
        ):
            rows.append((subtype, payload, True))
            for override in list(payload.get("thickness_overrides", []) or []):
                if isinstance(override, dict):
                    rows.append((subtype, dict(override), False))
        self.table.setRowCount(len(rows))
        for row_index, (subtype, payload, is_base) in enumerate(rows):
            self.table.setItem(row_index, 0, _table_text_item(subtype))
            self.table.setItem(row_index, 1, _table_text_item("" if is_base else _fmt_num(payload.get("thickness_min_mm", 0.0), 1)))
            self.table.setItem(row_index, 2, _table_text_item("" if is_base else _fmt_num(payload.get("thickness_max_mm", 0.0), 1)))
            self.table.setItem(row_index, 3, _table_number_item(payload.get("price_per_kg", self.fallback_price_per_kg), 3))
            self.table.setItem(row_index, 4, _table_number_item(payload.get("density_kg_m3", 0.0), 1))
            self.table.setItem(row_index, 5, _table_number_item(payload.get("scrap_credit_per_kg", self.fallback_scrap_credit_per_kg), 3))
        self._apply_filter()

    def _add_row(self) -> None:
        row_index = self.table.rowCount()
        self.table.insertRow(row_index)
        self.table.setItem(row_index, 0, _table_text_item(""))
        self.table.setItem(row_index, 1, _table_text_item(""))
        self.table.setItem(row_index, 2, _table_text_item(""))
        self.table.setItem(row_index, 3, _table_number_item(self.fallback_price_per_kg, 3))
        self.table.setItem(row_index, 4, _table_number_item(0.0, 1))
        self.table.setItem(row_index, 5, _table_number_item(self.fallback_scrap_credit_per_kg, 3))
        self.table.setCurrentCell(row_index, 0)
        item = self.table.item(row_index, 0)
        if item is not None:
            self.table.editItem(item)
        self._apply_filter()

    def _remove_row(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = self.table.rowCount() - 1
        if row >= 0:
            subtype_item = self.table.item(row, 0)
            subtype = str(subtype_item.text() if subtype_item is not None else "").strip()
            min_item = self.table.item(row, 1)
            max_item = self.table.item(row, 2)
            is_base_row = not str(min_item.text() if min_item is not None else "").strip() and not str(max_item.text() if max_item is not None else "").strip()
            if subtype and is_base_row:
                rows_to_remove = []
                for row_index in range(self.table.rowCount()):
                    other_item = self.table.item(row_index, 0)
                    other_subtype = str(other_item.text() if other_item is not None else "").strip()
                    if other_subtype == subtype:
                        rows_to_remove.append(row_index)
                for row_index in reversed(rows_to_remove):
                    self.table.removeRow(row_index)
            else:
                self.table.removeRow(row)
        self._apply_filter()

    def _apply_filter(self) -> None:
        token = str(self.search_edit.text() or "").strip().lower()
        visible = 0
        for row_index in range(self.table.rowCount()):
            texts = []
            for col_index in range(self.table.columnCount()):
                item = self.table.item(row_index, col_index)
                texts.append(str(item.text() if item is not None else "").strip().lower())
            matches = not token or any(token in text for text in texts if text)
            self.table.setRowHidden(row_index, not matches)
            if matches:
                visible += 1
        self.count_label.setText(f"{visible} regras visiveis")

    def catalog_payload(self) -> dict[str, Any]:
        catalog: dict[str, Any] = {}
        for row_index in range(self.table.rowCount()):
            subtype_item = self.table.item(row_index, 0)
            subtype = str(subtype_item.text() if subtype_item is not None else "").strip()
            if not subtype:
                continue
            min_item = self.table.item(row_index, 1)
            max_item = self.table.item(row_index, 2)
            price_item = self.table.item(row_index, 3)
            density_item = self.table.item(row_index, 4)
            scrap_item = self.table.item(row_index, 5)
            min_txt = str(min_item.text() if min_item is not None else "").strip().replace(",", ".")
            max_txt = str(max_item.text() if max_item is not None else "").strip().replace(",", ".")
            try:
                price_per_kg = float(str(price_item.text() if price_item is not None else "0").strip().replace(",", ".") or 0.0)
            except Exception:
                price_per_kg = 0.0
            try:
                density_kg_m3 = float(str(density_item.text() if density_item is not None else "0").strip().replace(",", ".") or 0.0)
            except Exception:
                density_kg_m3 = 0.0
            try:
                scrap_credit_per_kg = float(str(scrap_item.text() if scrap_item is not None else "0").strip().replace(",", ".") or 0.0)
            except Exception:
                scrap_credit_per_kg = 0.0
            payload = {
                "price_per_kg": max(0.0, price_per_kg),
                "scrap_credit_per_kg": max(0.0, scrap_credit_per_kg),
            }
            if density_kg_m3 > 0.0:
                payload["density_kg_m3"] = density_kg_m3
            if min_txt or max_txt:
                try:
                    thickness_min_mm = max(0.0, float(min_txt or 0.0))
                except Exception:
                    thickness_min_mm = 0.0
                try:
                    thickness_max_mm = max(0.0, float(max_txt or 0.0))
                except Exception:
                    thickness_max_mm = 0.0
                if thickness_min_mm > 0.0 and thickness_max_mm > 0.0 and thickness_max_mm < thickness_min_mm:
                    thickness_min_mm, thickness_max_mm = thickness_max_mm, thickness_min_mm
                override = dict(payload)
                if thickness_min_mm > 0.0:
                    override["thickness_min_mm"] = thickness_min_mm
                if thickness_max_mm > 0.0:
                    override["thickness_max_mm"] = thickness_max_mm
                base = dict(catalog.get(subtype, {}) or {})
                overrides = [dict(item or {}) for item in list(base.get("thickness_overrides", []) or [])]
                overrides.append(override)
                base["thickness_overrides"] = overrides
                catalog[subtype] = base
            else:
                base = dict(catalog.get(subtype, {}) or {})
                overrides = [dict(item or {}) for item in list(base.get("thickness_overrides", []) or [])]
                base.update(payload)
                if overrides:
                    base["thickness_overrides"] = overrides
                catalog[subtype] = base
        for subtype, payload in list(catalog.items()):
            overrides = [dict(item or {}) for item in list(payload.get("thickness_overrides", []) or []) if isinstance(item, dict)]
            if overrides:
                overrides.sort(
                    key=lambda item: (
                        float(item.get("thickness_min_mm", 0) or 0),
                        float(item.get("thickness_max_mm", 0) or 0),
                    )
                )
                payload["thickness_overrides"] = overrides
        return catalog


class SeriesPricingDialog(QDialog):
    def __init__(self, tiers: list[dict[str, Any]], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Series de quantidade")
        self.resize(860, 520)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        header_card = CardFrame()
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(6)
        title = QLabel("Regras de preco por serie")
        title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        helper = QLabel(
            "Cada linha define uma faixa de quantidade. "
            "O ajuste de margem soma ao valor base e o multiplicador de setup reforca ou reduz o peso do setup."
        )
        helper.setProperty("role", "muted")
        header_layout.addWidget(title)
        header_layout.addWidget(helper)
        root.addWidget(header_card)

        table_card = CardFrame()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(12, 10, 12, 10)
        table_layout.setSpacing(8)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Serie", "Qtd min", "Qtd max", "Ajuste margem %", "Mult. setup"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for section in (1, 2, 3, 4):
            header.setSectionResizeMode(section, QHeaderView.ResizeToContents)
        table_layout.addWidget(self.table, 1)
        actions = QHBoxLayout()
        self.add_btn = QPushButton("Adicionar serie")
        self.remove_btn = QPushButton("Remover serie")
        self.remove_btn.setProperty("variant", "secondary")
        actions.addWidget(self.add_btn)
        actions.addWidget(self.remove_btn)
        actions.addStretch(1)
        table_layout.addLayout(actions)
        root.addWidget(table_card, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self.add_btn.clicked.connect(self._add_row)
        self.remove_btn.clicked.connect(self._remove_row)
        self._load_rows(tiers)

    def _load_rows(self, tiers: list[dict[str, Any]]) -> None:
        normalized = sorted(
            [dict(tier or {}) for tier in list(tiers or [])],
            key=lambda row: (int(float(row.get("qty_min", 1) or 1)), int(float(row.get("qty_max", 1) or 1))),
        )
        self.table.setRowCount(len(normalized))
        for row_index, tier in enumerate(normalized):
            self.table.setItem(row_index, 0, _table_text_item(tier.get("label", "")))
            self.table.setItem(row_index, 1, _table_number_item(tier.get("qty_min", 1), 0))
            self.table.setItem(row_index, 2, _table_number_item(tier.get("qty_max", 1), 0))
            self.table.setItem(row_index, 3, _table_number_item(tier.get("margin_delta_pct", 0), 2))
            self.table.setItem(row_index, 4, _table_number_item(tier.get("setup_multiplier", 1), 2))

    def _add_row(self) -> None:
        row_index = self.table.rowCount()
        self.table.insertRow(row_index)
        defaults = [
            _table_text_item(f"Serie {row_index + 1}"),
            _table_number_item(1, 0),
            _table_number_item(1, 0),
            _table_number_item(0, 2),
            _table_number_item(1, 2),
        ]
        for col_index, item in enumerate(defaults):
            self.table.setItem(row_index, col_index, item)
        self.table.setCurrentCell(row_index, 0)
        item = self.table.item(row_index, 0)
        if item is not None:
            self.table.editItem(item)

    def _remove_row(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = self.table.rowCount() - 1
        if row >= 0:
            self.table.removeRow(row)

    def tiers_payload(self) -> list[dict[str, Any]]:
        tiers: list[dict[str, Any]] = []
        for row_index in range(self.table.rowCount()):
            label_item = self.table.item(row_index, 0)
            label = str(label_item.text() if label_item is not None else "").strip()
            if not label:
                continue
            qty_min_item = self.table.item(row_index, 1)
            qty_max_item = self.table.item(row_index, 2)
            margin_item = self.table.item(row_index, 3)
            setup_item = self.table.item(row_index, 4)
            try:
                qty_min = max(1, int(float(str(qty_min_item.text() if qty_min_item is not None else "1").strip().replace(",", ".") or 1)))
            except Exception:
                qty_min = 1
            try:
                qty_max = max(qty_min, int(float(str(qty_max_item.text() if qty_max_item is not None else str(qty_min)).strip().replace(",", ".") or qty_min)))
            except Exception:
                qty_max = qty_min
            try:
                margin_delta = float(str(margin_item.text() if margin_item is not None else "0").strip().replace(",", ".") or 0.0)
            except Exception:
                margin_delta = 0.0
            try:
                setup_multiplier = max(0.05, float(str(setup_item.text() if setup_item is not None else "1").strip().replace(",", ".") or 1.0))
            except Exception:
                setup_multiplier = 1.0
            tiers.append(
                {
                    "key": f"tier_{row_index + 1}",
                    "label": label,
                    "qty_min": qty_min,
                    "qty_max": qty_max,
                    "margin_delta_pct": margin_delta,
                    "setup_multiplier": setup_multiplier,
                }
            )
        tiers.sort(key=lambda row: (int(row.get("qty_min", 1)), int(row.get("qty_max", 1))))
        return tiers


class LaserSettingsDialog(QDialog):
    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self._material_catalog_cache: dict[str, Any] = {}
        self._series_tiers_cache: list[dict[str, Any]] = []
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        self.setWindowTitle("Configuracao de Orcamento Laser")
        self.setMinimumSize(920, 620)
        self.resize(1040, 690)
        self.setStyleSheet(
            """
            QLabel {
                font-size: 11px;
            }
            QLineEdit,
            QComboBox,
            QDoubleSpinBox,
            QPushButton {
                font-size: 11px;
            }
            QTabBar::tab {
                min-height: 28px;
                padding: 6px 14px;
            }
            """
        )
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFrameShape(QScrollArea.NoFrame)
        outer.addWidget(scroll, 1)

        content = QWidget()
        scroll.setWidget(content)
        self._content_widget = content
        self._scroll_area = scroll
        self._apply_screen_bounds()

        root = QVBoxLayout(content)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        top_card = CardFrame()
        top_layout = QGridLayout(top_card)
        top_layout.setContentsMargins(10, 8, 10, 8)
        top_layout.setHorizontalSpacing(8)
        top_layout.setVerticalSpacing(4)

        self.machine_combo = QComboBox()
        self.machine_combo.setEditable(False)
        self.machine_combo.addItems(list(dict(self.settings.get("machine_profiles", {}) or {}).keys()))
        self.machine_combo.setCurrentText(str(self.settings.get("active_machine", "") or self.machine_combo.currentText()))

        self.commercial_combo = QComboBox()
        self.commercial_combo.setEditable(False)
        self.commercial_combo.addItems(list(dict(self.settings.get("commercial_profiles", {}) or {}).keys()))
        self.commercial_combo.setCurrentText(str(self.settings.get("active_commercial", "") or self.commercial_combo.currentText()))

        self.material_combo = QComboBox()
        self.material_combo.setEditable(True)
        self.subtype_combo = QComboBox()
        self.subtype_combo.setEditable(True)
        self.gas_combo = QComboBox()
        self.gas_combo.setEditable(True)
        self.mark_patterns_edit = QLineEdit()
        self.ignore_patterns_edit = QLineEdit()
        for widget in (
            self.machine_combo,
            self.commercial_combo,
            self.material_combo,
            self.subtype_combo,
            self.gas_combo,
            self.mark_patterns_edit,
            self.ignore_patterns_edit,
        ):
            _densify_editor(widget)

        top_layout.addWidget(QLabel("Maquina"), 0, 0)
        top_layout.addWidget(QLabel("Perfil comercial"), 0, 1)
        top_layout.addWidget(QLabel("Material"), 0, 2)
        top_layout.addWidget(QLabel("Subtipo"), 0, 3)
        top_layout.addWidget(QLabel("Gas"), 0, 4)
        top_layout.addWidget(self.machine_combo, 1, 0)
        top_layout.addWidget(self.commercial_combo, 1, 1)
        top_layout.addWidget(self.material_combo, 1, 2)
        top_layout.addWidget(self.subtype_combo, 1, 3)
        top_layout.addWidget(self.gas_combo, 1, 4)
        top_layout.addWidget(QLabel("Padroes marcacao"), 2, 0)
        top_layout.addWidget(QLabel("Padroes ignorar"), 2, 2)
        top_layout.addWidget(self.mark_patterns_edit, 3, 0, 1, 2)
        top_layout.addWidget(self.ignore_patterns_edit, 3, 2, 1, 2)
        root.addWidget(top_card)

        manage_card = CardFrame()
        manage_layout = QGridLayout(manage_card)
        manage_layout.setContentsMargins(10, 8, 10, 8)
        manage_layout.setHorizontalSpacing(8)
        manage_layout.setVerticalSpacing(4)
        manage_title = QLabel("Tabelas e regras")
        manage_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        manage_layout.addWidget(manage_title, 0, 0, 1, 3)
        self.manage_cut_table_btn = QPushButton("Tabela de corte")
        self.manage_cut_table_btn.setProperty("variant", "secondary")
        self.manage_material_catalog_btn = QPushButton("Tabela subtipos")
        self.manage_material_catalog_btn.setProperty("variant", "secondary")
        self.manage_series_btn = QPushButton("Series")
        self.manage_series_btn.setProperty("variant", "secondary")
        for button in (self.manage_cut_table_btn, self.manage_material_catalog_btn, self.manage_series_btn):
            button.setMinimumHeight(30)
            button.setCursor(Qt.PointingHandCursor)
        self.cut_table_summary_label = QLabel("")
        self.material_catalog_summary_label = QLabel("")
        self.series_summary_label = QLabel("")
        for label in (self.cut_table_summary_label, self.material_catalog_summary_label, self.series_summary_label):
            label.setProperty("role", "muted")
            label.setWordWrap(True)
        manage_layout.addWidget(self.manage_cut_table_btn, 1, 0)
        manage_layout.addWidget(self.manage_material_catalog_btn, 1, 1)
        manage_layout.addWidget(self.manage_series_btn, 1, 2)
        manage_layout.addWidget(self.cut_table_summary_label, 2, 0)
        manage_layout.addWidget(self.material_catalog_summary_label, 2, 1)
        manage_layout.addWidget(self.series_summary_label, 2, 2)
        root.addWidget(manage_card)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        root.addWidget(self.tabs, 1)

        self.machine_tab = QWidget()
        self.machine_tab_layout = QVBoxLayout(self.machine_tab)
        self.machine_tab_layout.setContentsMargins(4, 4, 4, 4)
        self.machine_tab_layout.setSpacing(8)

        motion_card = CardFrame()
        motion_shell = QVBoxLayout(motion_card)
        motion_shell.setContentsMargins(10, 8, 10, 8)
        motion_shell.setSpacing(6)
        motion_grid = QGridLayout()
        motion_grid.setHorizontalSpacing(12)
        motion_grid.setVerticalSpacing(0)
        motion_left = QFormLayout()
        motion_left.setHorizontalSpacing(10)
        motion_left.setVerticalSpacing(6)
        motion_right = QFormLayout()
        motion_right.setHorizontalSpacing(10)
        motion_right.setVerticalSpacing(6)
        self.rapid_speed_spin = _spin(1, 1.0, 2000.0, 200.0, 10.0)
        self.mark_speed_spin = _spin(2, 0.1, 200.0, 18.0, 0.5)
        self.speed_factor_spin = _spin(1, 1.0, 200.0, 92.0, 1.0)
        self.lead_in_spin = _spin(2, 0.0, 100.0, 2.0, 0.5)
        self.lead_out_spin = _spin(2, 0.0, 100.0, 2.0, 0.5)
        self.lead_speed_spin = _spin(2, 0.1, 100.0, 3.0, 0.5)
        self.pierce_base_spin = _spin(0, 0.0, 10000.0, 400.0, 50.0)
        self.pierce_per_mm_spin = _spin(1, 0.0, 1000.0, 35.0, 1.0)
        self.first_gas_delay_spin = _spin(0, 0.0, 10000.0, 200.0, 50.0)
        self.gas_delay_spin = _spin(0, 0.0, 10000.0, 0.0, 50.0)
        self.overhead_spin = _spin(1, 0.0, 200.0, 4.0, 0.5)
        for widget in (
            self.rapid_speed_spin,
            self.mark_speed_spin,
            self.speed_factor_spin,
            self.lead_in_spin,
            self.lead_out_spin,
            self.lead_speed_spin,
            self.pierce_base_spin,
            self.pierce_per_mm_spin,
            self.first_gas_delay_spin,
            self.gas_delay_spin,
            self.overhead_spin,
        ):
            _densify_editor(widget)
        motion_left.addRow("Rapidos (mm/s)", self.rapid_speed_spin)
        motion_left.addRow("Marcacao (m/min)", self.mark_speed_spin)
        motion_left.addRow("Fator velocidade %", self.speed_factor_spin)
        motion_left.addRow("Lead-in (mm)", self.lead_in_spin)
        motion_left.addRow("Lead-out (mm)", self.lead_out_spin)
        motion_left.addRow("Velocidade lead (mm/s)", self.lead_speed_spin)
        motion_right.addRow("Pierce base (ms)", self.pierce_base_spin)
        motion_right.addRow("Pierce por mm (ms)", self.pierce_per_mm_spin)
        motion_right.addRow("1o atraso gas (ms)", self.first_gas_delay_spin)
        motion_right.addRow("Atraso gas (ms)", self.gas_delay_spin)
        motion_right.addRow("Overhead movimento %", self.overhead_spin)
        motion_grid.addLayout(motion_left, 0, 0)
        motion_grid.addLayout(motion_right, 0, 1)
        motion_grid.setColumnStretch(0, 1)
        motion_grid.setColumnStretch(1, 1)
        motion_shell.addLayout(motion_grid)
        self.machine_tab_layout.addWidget(motion_card)

        machine_note_card = CardFrame()
        machine_note_layout = QVBoxLayout(machine_note_card)
        machine_note_layout.setContentsMargins(12, 10, 12, 10)
        machine_note_layout.setSpacing(6)
        machine_note_title = QLabel("Tabela de corte separada")
        machine_note_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
        machine_note_text = QLabel(
            "Os parametros por espessura ficam numa janela propria para poderes trabalhar com mais detalhe "
            "sem apertar esta configuracao principal."
        )
        machine_note_text.setWordWrap(True)
        machine_note_text.setProperty("role", "muted")
        machine_note_layout.addWidget(machine_note_title)
        machine_note_layout.addWidget(machine_note_text)
        self.machine_tab_layout.addWidget(machine_note_card, 1)
        self.tabs.addTab(self.machine_tab, "Maquina")

        self.cost_tab = QWidget()
        self.cost_tab_layout = QVBoxLayout(self.cost_tab)
        self.cost_tab_layout.setContentsMargins(4, 4, 4, 4)
        self.cost_tab_layout.setSpacing(8)

        self.tabs.addTab(self.cost_tab, "Custos")

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        try:
            self.subtype_combo.lineEdit().setPlaceholderText("Opcional / grau")
        except Exception:
            pass
        self.machine_combo.currentTextChanged.connect(self._refresh_materials)
        self.material_combo.currentTextChanged.connect(self._refresh_subtypes)
        self._build_cost_tab()
        self.gas_combo.currentTextChanged.connect(self._load_current_values)
        self.commercial_combo.currentTextChanged.connect(self._on_commercial_changed)
        self.manage_cut_table_btn.clicked.connect(self._open_cut_table_dialog)
        self.manage_material_catalog_btn.clicked.connect(self._open_material_catalog_dialog)
        self.manage_series_btn.clicked.connect(self._open_series_dialog)
        self._reload_material_catalog_cache_from_current_profile()
        self._reload_series_cache_from_current_profile()
        self._refresh_materials()

    def _apply_screen_bounds(self) -> None:
        screen = self.screen()
        if screen is None:
            screen = QApplication.primaryScreen()
        if screen is None:
            return
        available = screen.availableGeometry()
        max_width = max(760, available.width() - 24)
        max_height = max(560, available.height() - 40)
        self.setMaximumSize(max_width, max_height)
        self.resize(min(1040, max_width), min(690, max_height))
        current = self.frameGeometry()
        x = max(available.left(), min(current.x(), available.right() - current.width() + 1))
        y = max(available.top(), min(current.y(), available.bottom() - current.height() + 1))
        self.move(x, y)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._apply_screen_bounds()

    def _build_cost_tab(self) -> None:
        cost_card = CardFrame()
        cost_shell = QVBoxLayout(cost_card)
        cost_shell.setContentsMargins(10, 8, 10, 8)
        cost_shell.setSpacing(6)
        cost_grid = QGridLayout()
        cost_grid.setHorizontalSpacing(12)
        cost_grid.setVerticalSpacing(0)
        cost_left = QFormLayout()
        cost_left.setHorizontalSpacing(10)
        cost_left.setVerticalSpacing(6)
        cost_right = QFormLayout()
        cost_right.setHorizontalSpacing(10)
        cost_right.setVerticalSpacing(6)
        self.density_spin = _spin(1, 1.0, 50000.0, 7800.0, 50.0)
        self.material_price_spin = _spin(3, 0.0, 1000.0, 1.30, 0.05)
        self.scrap_credit_spin = _spin(3, 0.0, 1000.0, 1.20, 0.05)
        self.cut_rate_spin = _spin(3, 0.0, 1000.0, 1.50, 0.05)
        self.mark_rate_spin = _spin(3, 0.0, 1000.0, 0.66, 0.05)
        self.defilm_rate_spin = _spin(3, 0.0, 1000.0, 0.50, 0.05)
        self.pierce_rate_spin = _spin(3, 0.0, 1000.0, 0.15, 0.01)
        self.machine_hour_spin = _spin(2, 0.0, 10000.0, 90.0, 1.0)
        self.setup_time_spin = _spin(2, 0.0, 600.0, 3.0, 0.5)
        self.utilization_spin = _spin(1, 1.0, 100.0, 82.0, 1.0)
        self.fill_factor_spin = _spin(1, 1.0, 100.0, 72.0, 1.0)
        self.minimum_spin = _spin(2, 0.0, 10000.0, 0.0, 1.0)
        self.margin_spin = _spin(2, 0.0, 500.0, 18.0, 1.0)
        self.handling_spin = _spin(2, 0.0, 10000.0, 0.0, 0.5)
        self.cost_mode_combo = QComboBox()
        self.cost_mode_combo.addItems(["hybrid_max", "per_meter", "machine_time"])
        self.scrap_credit_check = QCheckBox("Subtrair valor da sucata")
        self.scrap_credit_check.setChecked(True)
        for widget in (
            self.density_spin,
            self.material_price_spin,
            self.scrap_credit_spin,
            self.cut_rate_spin,
            self.mark_rate_spin,
            self.defilm_rate_spin,
            self.pierce_rate_spin,
            self.machine_hour_spin,
            self.setup_time_spin,
            self.utilization_spin,
            self.fill_factor_spin,
            self.margin_spin,
            self.handling_spin,
            self.cost_mode_combo,
        ):
            _densify_editor(widget)
        cost_left.addRow("Densidade (kg/m3)", self.density_spin)
        cost_left.addRow("Preco material (EUR/kg)", self.material_price_spin)
        cost_left.addRow("Valor sucata (EUR/kg)", self.scrap_credit_spin)
        cost_left.addRow("Corte (EUR/m)", self.cut_rate_spin)
        cost_left.addRow("Marcacao (EUR/m)", self.mark_rate_spin)
        cost_left.addRow("Defilm (EUR/m)", self.defilm_rate_spin)
        cost_right.addRow("Pierce (EUR/pc)", self.pierce_rate_spin)
        cost_right.addRow("Hora maquina (EUR/h)", self.machine_hour_spin)
        cost_right.addRow("Setup (min)", self.setup_time_spin)
        cost_right.addRow("Aproveitamento %", self.utilization_spin)
        cost_right.addRow("Fallback area %", self.fill_factor_spin)
        cost_right.addRow("Margem %", self.margin_spin)
        cost_right.addRow("Manuseamento (EUR)", self.handling_spin)
        cost_right.addRow("Modo custo", self.cost_mode_combo)
        cost_right.addRow("", self.scrap_credit_check)
        cost_grid.addLayout(cost_left, 0, 0)
        cost_grid.addLayout(cost_right, 0, 1)
        cost_grid.setColumnStretch(0, 1)
        cost_grid.setColumnStretch(1, 1)
        cost_shell.addLayout(cost_grid)
        self.cost_tab_layout.addWidget(cost_card)

    def _machine_profile(self) -> dict[str, Any]:
        profiles = dict(self.settings.get("machine_profiles", {}) or {})
        name = self.machine_combo.currentText().strip()
        return dict(profiles.get(name, {}) or {})

    def _commercial_profile(self) -> dict[str, Any]:
        profiles = dict(self.settings.get("commercial_profiles", {}) or {})
        name = self.commercial_combo.currentText().strip()
        return dict(profiles.get(name, {}) or {})

    def _reload_material_catalog_cache_from_current_profile(self) -> None:
        commercial = self._commercial_profile()
        self._material_catalog_cache = {
            str(family or "").strip(): dict(payload or {})
            for family, payload in dict(commercial.get("material_catalog", {}) or {}).items()
            if str(family or "").strip()
        }

    def _reload_series_cache_from_current_profile(self) -> None:
        commercial = self._commercial_profile()
        series_cfg = dict(commercial.get("series_pricing", {}) or {})
        self._series_tiers_cache = [dict(item or {}) for item in list(series_cfg.get("tiers", []) or [])]

    def _current_material_catalog(self) -> dict[str, Any]:
        family = _canonical_material_family(self.material_combo.currentText().strip()) or self.material_combo.currentText().strip()
        return dict(self._material_catalog_cache.get(family, {}) or {})

    def _on_commercial_changed(self) -> None:
        self._reload_material_catalog_cache_from_current_profile()
        self._reload_series_cache_from_current_profile()
        self._load_current_values()

    def _refresh_materials(self) -> None:
        current_material = _display_material_family(self.material_combo.currentText().strip()) or self.material_combo.currentText().strip()
        values = _settings_material_names(self.settings, self.machine_combo.currentText().strip())
        _set_combo_values(self.material_combo, values, current_material if current_material else values[0])
        self._refresh_subtypes()

    def _material_subtype_candidates(self, material_name: str) -> list[str]:
        extras: list[str] = []
        try:
            presets = dict(self.backend.material_presets() or {})
            family = _guess_material_family(material_name) or str(material_name or "").strip()
            for value in list(presets.get("materiais", []) or []):
                clean = str(value or "").strip()
                if not clean or clean == material_name:
                    continue
                if _guess_material_family(clean) == family and clean not in extras:
                    extras.append(clean)
        except Exception:
            pass
        return _settings_material_subtypes(self.settings, material_name, extras)

    def _refresh_subtypes(self) -> None:
        current_subtype = self.subtype_combo.currentText().strip()
        values = self._material_subtype_candidates(self.material_combo.currentText().strip())
        _set_combo_values(self.subtype_combo, values, current_subtype)
        if self.subtype_combo.count() and not self.subtype_combo.currentText().strip():
            self.subtype_combo.setCurrentIndex(0)
        self._refresh_gases()

    def _refresh_gases(self) -> None:
        current_gas = self.gas_combo.currentText().strip()
        values = _settings_gas_names(self.settings, self.machine_combo.currentText().strip(), self.material_combo.currentText().strip())
        preferred = _preferred_material_gas(self.settings, self.machine_combo.currentText().strip(), self.material_combo.currentText().strip(), current_gas)
        _set_combo_values(self.gas_combo, values, preferred)
        self._load_current_values()

    def _load_current_values(self) -> None:
        machine = self._machine_profile()
        motion = dict(machine.get("motion", {}) or {})
        self.mark_patterns_edit.setText(",".join(list(dict(self.settings.get("layer_rules", {}) or {}).get("mark_patterns", []) or [])))
        self.ignore_patterns_edit.setText(",".join(list(dict(self.settings.get("layer_rules", {}) or {}).get("ignore_patterns", []) or [])))
        self.rapid_speed_spin.setValue(float(motion.get("rapid_speed_mm_s", 200.0) or 200.0))
        self.mark_speed_spin.setValue(float(motion.get("mark_speed_m_min", 18.0) or 18.0))
        self.speed_factor_spin.setValue(float(motion.get("effective_speed_factor_pct", 92.0) or 92.0))
        self.lead_in_spin.setValue(float(motion.get("lead_in_mm", 2.0) or 2.0))
        self.lead_out_spin.setValue(float(motion.get("lead_out_mm", 2.0) or 2.0))
        self.lead_speed_spin.setValue(float(motion.get("lead_move_speed_mm_s", 3.0) or 3.0))
        self.pierce_base_spin.setValue(float(motion.get("pierce_base_ms", 400.0) or 400.0))
        self.pierce_per_mm_spin.setValue(float(motion.get("pierce_per_mm_ms", 35.0) or 35.0))
        self.first_gas_delay_spin.setValue(float(motion.get("first_gas_delay_ms", 200.0) or 200.0))
        self.gas_delay_spin.setValue(float(motion.get("gas_delay_ms", 0.0) or 0.0))
        self.overhead_spin.setValue(float(motion.get("motion_overhead_pct", 4.0) or 4.0))

        materials = dict(machine.get("materials", {}) or {})
        material = _resolve_material_mapping(materials, self.material_combo.currentText().strip())
        gases = dict(material.get("gases", {}) or {})
        gas = dict(gases.get(self.gas_combo.currentText().strip(), {}) or {})
        rows = list(gas.get("rows", []) or [])
        self.cut_table_summary_label.setText(
            f"{self.material_combo.currentText().strip() or 'Material'} / {self.gas_combo.currentText().strip() or 'Gas'}: {_cut_rows_summary(rows)}"
        )

        commercial = self._commercial_profile()
        commercial_materials = dict(commercial.get("materials", {}) or {})
        commercial_material = _resolve_material_mapping(commercial_materials, self.material_combo.currentText().strip())
        rates = dict(commercial.get("rates", {}) or {})
        self.density_spin.setValue(float(commercial_material.get("density_kg_m3", material.get("density_kg_m3", 7800.0)) or 7800.0))
        self.material_price_spin.setValue(float(commercial_material.get("price_per_kg", 1.3) or 1.3))
        self.scrap_credit_spin.setValue(float(commercial_material.get("scrap_credit_per_kg", 1.2) or 1.2))
        self.cut_rate_spin.setValue(float(rates.get("cut_per_m_eur", 1.5) or 1.5))
        self.mark_rate_spin.setValue(float(rates.get("marking_per_m_eur", 0.66) or 0.66))
        self.defilm_rate_spin.setValue(float(rates.get("defilm_per_m_eur", 0.5) or 0.5))
        self.pierce_rate_spin.setValue(float(rates.get("pierce_eur", 0.15) or 0.15))
        self.machine_hour_spin.setValue(float(rates.get("machine_hour_eur", 90.0) or 90.0))
        self.setup_time_spin.setValue(float(commercial.get("setup_time_min", 3.0) or 3.0))
        self.utilization_spin.setValue(float(commercial.get("material_utilization_pct", 82.0) or 82.0))
        self.fill_factor_spin.setValue(float(commercial.get("fallback_fill_pct", 72.0) or 72.0))
        self.minimum_spin.setValue(0.0)
        self.margin_spin.setValue(float(commercial.get("margin_pct", 18.0) or 18.0))
        self.handling_spin.setValue(float(commercial.get("handling_eur", 0.0) or 0.0))
        self.cost_mode_combo.setCurrentText(str(commercial.get("cost_mode", "hybrid_max") or "hybrid_max"))
        self.scrap_credit_check.setChecked(bool(commercial.get("use_scrap_credit", True)))
        self._update_material_catalog_summary()
        self._update_series_summary()

    def _update_material_catalog_summary(self) -> None:
        catalog = self._current_material_catalog()
        current_family = self.material_combo.currentText().strip() or _display_material_family(self.material_combo.currentText().strip()) or "material"
        rows = sorted([str(subtype or "").strip() for subtype in dict(catalog or {}).keys() if str(subtype or "").strip()], key=str.upper)
        sample = ", ".join(rows[:4])
        if len(rows) > 4:
            sample = f"{sample}, ..."
        if rows:
            summary = (
                f"{len(rows)} subtipos configurados para {current_family}: {sample}. "
                "Abre a janela para editar os precos por categoria."
            )
            subtype_count, rule_count = _material_catalog_metrics(catalog)
            self.manage_material_catalog_btn.setText(f"Tabela subtipos ({subtype_count})")
            self.material_catalog_summary_label.setText(f"{rule_count} regras comerciais configuradas.")
        else:
            summary = (
                f"Sem subtipos configurados para {current_family}. "
                "Abre a janela para criar categorias com preco proprio."
            )
            self.manage_material_catalog_btn.setText("Tabela subtipos")
            self.material_catalog_summary_label.setText("Sem regras comerciais para este material.")
        self.manage_material_catalog_btn.setToolTip(summary)
        self.manage_material_catalog_btn.setStatusTip(summary)

    def _update_series_summary(self) -> None:
        tiers = [dict(item or {}) for item in list(self._series_tiers_cache or [])]
        if not tiers:
            summary = "Sem series configuradas. O sistema usa a margem e o setup base em todas as quantidades."
            self.manage_series_btn.setText("Series")
            self.manage_series_btn.setToolTip(summary)
            self.manage_series_btn.setStatusTip(summary)
            self.series_summary_label.setText(summary)
            return
        parts: list[str] = []
        for tier in tiers[:4]:
            label = str(tier.get("label", "") or "").strip() or "Serie"
            qty_min = int(float(tier.get("qty_min", 1) or 1))
            qty_max = int(float(tier.get("qty_max", qty_min) or qty_min))
            margin_delta = float(tier.get("margin_delta_pct", 0.0) or 0.0)
            setup_multiplier = float(tier.get("setup_multiplier", 1.0) or 1.0)
            parts.append(f"{label}: {qty_min}-{qty_max} | margem {margin_delta:+.1f}% | setup x{setup_multiplier:.2f}")
        summary = " ; ".join(parts)
        if len(tiers) > 4:
            summary = f"{summary} ; ..."
        self.manage_series_btn.setText(f"Series ({len(tiers)})")
        self.manage_series_btn.setToolTip(summary)
        self.manage_series_btn.setStatusTip(summary)
        self.series_summary_label.setText(summary)

    def _open_cut_table_dialog(self) -> None:
        machine_name = self.machine_combo.currentText().strip()
        material_name = _canonical_material_family(self.material_combo.currentText().strip()) or "Aco carbono"
        gas_name = _preferred_material_gas(self.settings, machine_name, material_name, self.gas_combo.currentText().strip())
        machine_profiles = dict(self.settings.get("machine_profiles", {}) or {})
        machine = dict(machine_profiles.get(machine_name, {}) or {})
        materials = dict(machine.get("materials", {}) or {})
        material = dict(materials.get(material_name, {}) or {})
        gases = dict(material.get("gases", {}) or {})
        rows = list(dict(gases.get(gas_name, {}) or {}).get("rows", []) or [])
        dialog = CutTableDialog(_display_material_family(material_name), gas_name, rows, self)
        if dialog.exec() != QDialog.Accepted:
            return
        gases[gas_name] = {"rows": dialog.rows_payload()}
        material["gases"] = gases
        if _forced_gas_for_material(material_name):
            material["default_gas"] = _forced_gas_for_material(material_name)
        else:
            material["default_gas"] = gas_name
        materials[material_name] = material
        machine["materials"] = materials
        machine_profiles[machine_name] = machine
        self.settings["machine_profiles"] = machine_profiles
        self.gas_combo.setCurrentText(material.get("default_gas", gas_name))
        self._load_current_values()
        self._persist_settings(close_dialog=False)

    def _open_material_catalog_dialog(self) -> None:
        current_family = _canonical_material_family(self.material_combo.currentText().strip()) or "Aco carbono"
        dialog = MaterialSubtypeCatalogDialog(
            _display_material_family(current_family),
            self._current_material_catalog(),
            self.material_price_spin.value(),
            self.scrap_credit_spin.value(),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        self._material_catalog_cache[current_family] = dialog.catalog_payload()
        self._update_material_catalog_summary()
        self._refresh_subtypes()
        self._persist_settings(close_dialog=False)

    def _open_series_dialog(self) -> None:
        dialog = SeriesPricingDialog(self._series_tiers_cache, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._series_tiers_cache = dialog.tiers_payload()
        self._update_series_summary()
        self._persist_settings(close_dialog=False)

    def _persist_settings(self, *, close_dialog: bool) -> bool:
        settings = dict(self.settings or {})
        defaults = default_laser_quote_settings()
        settings["active_machine"] = self.machine_combo.currentText().strip()
        settings["active_commercial"] = self.commercial_combo.currentText().strip()
        settings["layer_rules"] = {
            "mark_patterns": [item.strip() for item in self.mark_patterns_edit.text().split(",") if item.strip()],
            "ignore_patterns": [item.strip() for item in self.ignore_patterns_edit.text().split(",") if item.strip()],
        }
        machine_profiles = dict(settings.get("machine_profiles", {}) or {})
        machine = dict(machine_profiles.get(settings["active_machine"], {}) or {})
        motion = dict(machine.get("motion", {}) or {})
        motion.update(
            {
                "rapid_speed_mm_s": self.rapid_speed_spin.value(),
                "mark_speed_m_min": self.mark_speed_spin.value(),
                "effective_speed_factor_pct": self.speed_factor_spin.value(),
                "lead_in_mm": self.lead_in_spin.value(),
                "lead_out_mm": self.lead_out_spin.value(),
                "lead_move_speed_mm_s": self.lead_speed_spin.value(),
                "pierce_base_ms": self.pierce_base_spin.value(),
                "pierce_per_mm_ms": self.pierce_per_mm_spin.value(),
                "first_gas_delay_ms": self.first_gas_delay_spin.value(),
                "gas_delay_ms": self.gas_delay_spin.value(),
                "motion_overhead_pct": self.overhead_spin.value(),
            }
        )
        machine["motion"] = motion
        materials = dict(machine.get("materials", {}) or {})
        material_name = _canonical_material_family(self.material_combo.currentText().strip()) or "Aco carbono"
        material = dict(materials.get(material_name, {}) or {})
        material["density_kg_m3"] = self.density_spin.value()
        forced_gas = _forced_gas_for_material(material_name)
        material["default_gas"] = forced_gas or self.gas_combo.currentText().strip() or "Oxigenio"
        materials[material_name] = material
        machine["materials"] = materials
        machine_profiles[settings["active_machine"]] = machine
        settings["machine_profiles"] = machine_profiles

        commercial_profiles = dict(settings.get("commercial_profiles", {}) or {})
        commercial = dict(commercial_profiles.get(settings["active_commercial"], {}) or {})
        commercial["cost_mode"] = self.cost_mode_combo.currentText().strip() or "hybrid_max"
        commercial["minimum_line_eur"] = 0.0
        commercial["margin_pct"] = self.margin_spin.value()
        commercial["setup_time_min"] = self.setup_time_spin.value()
        commercial["handling_eur"] = self.handling_spin.value()
        commercial["material_utilization_pct"] = self.utilization_spin.value()
        commercial["fallback_fill_pct"] = self.fill_factor_spin.value()
        commercial["use_scrap_credit"] = bool(self.scrap_credit_check.isChecked())
        commercial["series_pricing"] = {"tiers": [dict(item or {}) for item in list(self._series_tiers_cache or [])]}
        rates = dict(commercial.get("rates", {}) or {})
        rates.update(
            {
                "cut_per_m_eur": self.cut_rate_spin.value(),
                "marking_per_m_eur": self.mark_rate_spin.value(),
                "defilm_per_m_eur": self.defilm_rate_spin.value(),
                "pierce_eur": self.pierce_rate_spin.value(),
                "machine_hour_eur": self.machine_hour_spin.value(),
            }
        )
        commercial["rates"] = rates
        commercial_materials = dict(commercial.get("materials", {}) or {})
        commercial_materials[material_name] = {
            "density_kg_m3": self.density_spin.value(),
            "price_per_kg": self.material_price_spin.value(),
            "scrap_credit_per_kg": self.scrap_credit_spin.value(),
        }
        commercial["materials"] = commercial_materials
        material_catalog = {
            str(family or "").strip(): dict(payload or {})
            for family, payload in dict(self._material_catalog_cache or {}).items()
            if str(family or "").strip()
        }
        commercial["material_catalog"] = material_catalog
        default_profiles = dict(defaults.get("commercial_profiles", {}) or {})
        default_profile = dict(default_profiles.get(settings["active_commercial"], {}) or {})
        default_catalog = dict(default_profile.get("material_catalog", {}) or {})
        material_catalog_hidden: dict[str, list[str]] = {}
        for family in list(default_catalog.keys()) + [key for key in material_catalog.keys() if key not in default_catalog]:
            default_family = {
                str(key or "").strip()
                for key in dict(default_catalog.get(family, {}) or {}).keys()
                if str(key or "").strip()
            }
            current_family = {
                str(key or "").strip()
                for key in dict(material_catalog.get(family, {}) or {}).keys()
                if str(key or "").strip()
            }
            hidden = sorted(default_family - current_family, key=str.upper)
            if hidden:
                material_catalog_hidden[family] = hidden
        commercial["material_catalog_hidden"] = material_catalog_hidden
        commercial_profiles[settings["active_commercial"]] = commercial
        settings["commercial_profiles"] = commercial_profiles
        material_subtypes: dict[str, list[str]] = {}
        families: list[str] = []
        for profile in dict(commercial_profiles or {}).values():
            if not isinstance(profile, dict):
                continue
            for family in dict(profile.get("material_catalog", {}) or {}).keys():
                clean = str(family or "").strip()
                if clean and clean not in families:
                    families.append(clean)
        if material_name and material_name not in families:
            families.append(material_name)
        for family in families:
            values: list[str] = []
            for profile in dict(commercial_profiles or {}).values():
                if not isinstance(profile, dict):
                    continue
                for subtype in dict(dict(profile.get("material_catalog", {}) or {}).get(family, {}) or {}).keys():
                    clean = str(subtype or "").strip()
                    if clean and clean not in values:
                        values.append(clean)
            material_subtypes[family] = values
        default_subtypes = dict(defaults.get("material_subtypes", {}) or {})
        material_subtypes_hidden: dict[str, list[str]] = {}
        for family in list(default_subtypes.keys()) + [key for key in material_subtypes.keys() if key not in default_subtypes]:
            default_values = {
                str(value or "").strip()
                for value in list(default_subtypes.get(family, []) or [])
                if str(value or "").strip()
            }
            current_values = {
                str(value or "").strip()
                for value in list(material_subtypes.get(family, []) or [])
                if str(value or "").strip()
            }
            hidden = sorted(default_values - current_values, key=str.upper)
            if hidden:
                material_subtypes_hidden[family] = hidden
        settings["material_subtypes"] = material_subtypes
        settings["material_subtypes_hidden"] = material_subtypes_hidden

        try:
            self.backend.laser_quote_save_settings(settings)
        except Exception as exc:
            QMessageBox.critical(self, "Configuracao Laser", str(exc))
            return False
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self._reload_material_catalog_cache_from_current_profile()
        self._reload_series_cache_from_current_profile()
        self._load_current_values()
        if close_dialog:
            self.accept()
        return True

    def _save(self) -> None:
        self._persist_settings(close_dialog=True)


class LaserQuoteDialog(QDialog):
    def __init__(self, backend, parent=None, *, default_machine: str = "") -> None:
        super().__init__(parent)
        self.backend = backend
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self.analysis: dict[str, Any] = {}
        self.line_payload: dict[str, Any] = {}
        self.setWindowTitle("Peca Unit. DXF/DWG")
        self.resize(1120, 860)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        top_card = CardFrame()
        top_layout = QGridLayout(top_card)
        top_layout.setContentsMargins(12, 10, 12, 10)
        top_layout.setHorizontalSpacing(8)
        top_layout.setVerticalSpacing(6)
        self.file_edit = QLineEdit()
        self.description_edit = QLineEdit()
        self.ref_ext_edit = QLineEdit()
        self.machine_combo = QComboBox()
        self.machine_combo.addItems(list(dict(self.settings.get("machine_profiles", {}) or {}).keys()))
        self.machine_combo.setCurrentText(default_machine if default_machine in _combo_items(self.machine_combo) else str(self.settings.get("active_machine", "") or self.machine_combo.currentText()))
        self.commercial_combo = QComboBox()
        self.commercial_combo.addItems(list(dict(self.settings.get("commercial_profiles", {}) or {}).keys()))
        self.commercial_combo.setCurrentText(str(self.settings.get("active_commercial", "") or self.commercial_combo.currentText()))
        self.material_combo = QComboBox()
        self.material_combo.setEditable(True)
        self.subtype_combo = QComboBox()
        self.subtype_combo.setEditable(True)
        self.gas_combo = QComboBox()
        self.gas_combo.setEditable(True)
        self.thickness_spin = _spin(1, 0.1, 500.0, 8.0, 0.5)
        self.quantity_spin = _spin(0, 1.0, 100000.0, 1.0, 1.0)
        self.marking_check = QCheckBox("Contar marcacao")
        self.defilm_check = QCheckBox("Contar defilm")
        self.customer_material_check = QCheckBox("Material do cliente (sem materia-prima)")
        self.customer_material_check.setToolTip("Mantem corte/tempo/processo, mas retira o custo de materia-prima do orcamento.")
        browse_btn = QPushButton("Selecionar DXF/DWG")
        browse_btn.clicked.connect(self._pick_dxf)
        config_btn = QPushButton("Configurar perfis")
        config_btn.setProperty("variant", "secondary")
        config_btn.clicked.connect(self._configure_profiles)
        top_layout.addWidget(QLabel("DXF/DWG"), 0, 0)
        top_layout.addWidget(self.file_edit, 1, 0, 1, 3)
        top_layout.addWidget(browse_btn, 1, 3)
        top_layout.addWidget(QLabel("Descricao"), 2, 0)
        top_layout.addWidget(QLabel("Ref. externa"), 2, 2)
        top_layout.addWidget(self.description_edit, 3, 0, 1, 2)
        top_layout.addWidget(self.ref_ext_edit, 3, 2, 1, 2)
        top_layout.addWidget(QLabel("Maquina"), 4, 0)
        top_layout.addWidget(QLabel("Perfil comercial"), 4, 1)
        top_layout.addWidget(QLabel("Material"), 4, 2)
        top_layout.addWidget(QLabel("Subtipo"), 4, 3)
        top_layout.addWidget(QLabel("Gas"), 4, 4)
        top_layout.addWidget(self.machine_combo, 5, 0)
        top_layout.addWidget(self.commercial_combo, 5, 1)
        top_layout.addWidget(self.material_combo, 5, 2)
        top_layout.addWidget(self.subtype_combo, 5, 3)
        top_layout.addWidget(self.gas_combo, 5, 4)
        top_layout.addWidget(QLabel("Espessura (mm)"), 6, 0)
        top_layout.addWidget(QLabel("Quantidade"), 6, 1)
        top_layout.addWidget(self.marking_check, 6, 2)
        top_layout.addWidget(self.defilm_check, 6, 3)
        top_layout.addWidget(self.customer_material_check, 6, 4)
        top_layout.addWidget(self.thickness_spin, 7, 0)
        top_layout.addWidget(self.quantity_spin, 7, 1)
        top_layout.addWidget(config_btn, 7, 3, 1, 2)
        root.addWidget(top_card)

        middle = QHBoxLayout()
        middle.setSpacing(10)
        root.addLayout(middle, 1)

        left_host = QVBoxLayout()
        left_host.setSpacing(10)
        middle.addLayout(left_host, 3)

        overrides_card = CardFrame()
        overrides_layout = QFormLayout(overrides_card)
        overrides_layout.setContentsMargins(12, 10, 12, 10)
        overrides_layout.setHorizontalSpacing(10)
        overrides_layout.setVerticalSpacing(8)
        self.cut_override_spin = _spin(4, 0.0, 100000.0, 0.0, 0.05)
        self.mark_override_spin = _spin(4, 0.0, 100000.0, 0.0, 0.05)
        self.area_override_spin = _spin(6, 0.0, 1000000.0, 0.0, 0.001)
        self.pierce_override_spin = _spin(0, 0.0, 100000.0, 0.0, 1.0)
        overrides_layout.addRow("Corte m", self.cut_override_spin)
        overrides_layout.addRow("Marcacao m", self.mark_override_spin)
        overrides_layout.addRow("Area m2", self.area_override_spin)
        overrides_layout.addRow("Pierces", self.pierce_override_spin)
        left_host.addWidget(overrides_card)

        metrics_card = CardFrame()
        metrics_layout = QGridLayout(metrics_card)
        metrics_layout.setContentsMargins(12, 10, 12, 10)
        metrics_layout.setHorizontalSpacing(10)
        metrics_layout.setVerticalSpacing(6)
        self.metric_labels: dict[str, QLabel] = {}
        metric_specs = [
            ("bbox", "Caixa"),
            ("cut", "Corte"),
            ("mark", "Marcacao"),
            ("pierce", "Pierces"),
            ("area", "Area"),
            ("weight", "Peso"),
            ("speed", "Velocidade"),
            ("time", "Tempo"),
        ]
        for row_index, (key, title) in enumerate(metric_specs):
            label_title = QLabel(title)
            label_title.setProperty("role", "field_label")
            label_value = QLabel("-")
            label_value.setProperty("role", "field_value")
            metrics_layout.addWidget(label_title, row_index, 0)
            metrics_layout.addWidget(label_value, row_index, 1)
            self.metric_labels[key] = label_value
        left_host.addWidget(metrics_card)

        right_host = QVBoxLayout()
        right_host.setSpacing(10)
        middle.addLayout(right_host, 2)

        pricing_card = CardFrame()
        pricing_layout = QGridLayout(pricing_card)
        pricing_layout.setContentsMargins(12, 10, 12, 10)
        pricing_layout.setHorizontalSpacing(10)
        pricing_layout.setVerticalSpacing(6)
        self.price_labels: dict[str, QLabel] = {}
        price_specs = [
            ("material", "Material"),
            ("cut", "Corte usado"),
            ("machine_ref", "Tempo maquina"),
            ("mark", "Marcacao"),
            ("defilm", "Defilm"),
            ("pierce", "Pierce"),
            ("subtotal", "Subtotal tecnico"),
            ("unit", "Preco unit."),
            ("total", "Preco total"),
        ]
        for row_index, (key, title) in enumerate(price_specs):
            label_title = QLabel(title)
            label_title.setProperty("role", "field_label")
            label_value = QLabel("-")
            label_value.setProperty("role", "field_value")
            if key in ("unit", "total"):
                label_value.setProperty("role", "field_value_strong")
            pricing_layout.addWidget(label_title, row_index, 0)
            pricing_layout.addWidget(label_value, row_index, 1)
            self.price_labels[key] = label_value
        right_host.addWidget(pricing_card)

        self.warning_edit = QTextEdit()
        self.warning_edit.setReadOnly(True)
        self.warning_edit.setMinimumHeight(160)
        warning_card = CardFrame()
        warning_layout = QVBoxLayout(warning_card)
        warning_layout.setContentsMargins(12, 10, 12, 10)
        warning_layout.setSpacing(6)
        warning_title = QLabel("Alertas e observacoes")
        warning_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        warning_layout.addWidget(warning_title)
        warning_layout.addWidget(self.warning_edit, 1)
        right_host.addWidget(warning_card, 1)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.analyze_btn = QPushButton("Analisar e calcular")
        self.analyze_btn.clicked.connect(self._analyze)
        actions.addWidget(self.analyze_btn)
        actions.addStretch(1)
        root.addLayout(actions)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._accept_payload)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        try:
            self.subtype_combo.lineEdit().setPlaceholderText("Opcional / grau")
        except Exception:
            pass
        self.machine_combo.currentTextChanged.connect(self._refresh_materials)
        self.material_combo.currentTextChanged.connect(self._refresh_subtypes)
        self.file_edit.textChanged.connect(self._prefill_from_file)
        self._refresh_materials()

    def _refresh_materials(self) -> None:
        current_material = _display_material_family(self.material_combo.currentText().strip()) or self.material_combo.currentText().strip()
        values = _settings_material_names(self.settings, self.machine_combo.currentText().strip())
        _set_combo_values(self.material_combo, values, current_material if current_material else values[0])
        self._refresh_subtypes()

    def _material_subtype_candidates(self, material_name: str) -> list[str]:
        extras: list[str] = []
        try:
            presets = dict(self.backend.material_presets() or {})
            family = _guess_material_family(material_name) or str(material_name or "").strip()
            for value in list(presets.get("materiais", []) or []):
                clean = str(value or "").strip()
                if not clean or clean == material_name:
                    continue
                if _guess_material_family(clean) == family and clean not in extras:
                    extras.append(clean)
        except Exception:
            pass
        return _settings_material_subtypes(self.settings, material_name, extras)

    def _refresh_subtypes(self) -> None:
        current_subtype = self.subtype_combo.currentText().strip()
        values = self._material_subtype_candidates(self.material_combo.currentText().strip())
        _set_combo_values(self.subtype_combo, values, current_subtype)
        if self.subtype_combo.count() and not self.subtype_combo.currentText().strip():
            self.subtype_combo.setCurrentIndex(0)
        self._refresh_gases()

    def _refresh_gases(self) -> None:
        current_gas = self.gas_combo.currentText().strip()
        values = _settings_gas_names(self.settings, self.machine_combo.currentText().strip(), self.material_combo.currentText().strip())
        _set_combo_values(self.gas_combo, values, current_gas if current_gas else values[0])

    def _pick_dxf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar DXF/DWG",
            "",
            "Desenhos CAD (*.dxf *.dwg);;DXF (*.dxf);;DWG (*.dwg);;Todos (*.*)",
        )
        if path:
            self.file_edit.setText(path)

    def _prefill_from_file(self) -> None:
        text = self.file_edit.text().strip()
        if not text:
            return
        stem = text.replace("\\", "/").split("/")[-1].rsplit(".", 1)[0]
        if stem and not self.description_edit.text().strip():
            self.description_edit.setText(stem.replace("_", " "))
        if stem and not self.ref_ext_edit.text().strip():
            ref = []
            for char in stem.upper():
                ref.append(char if char.isalnum() or char in ("-", "_") else "-")
            self.ref_ext_edit.setText("".join(ref).strip("-_") or "PECA-LASER")

    def _configure_profiles(self) -> None:
        dialog = LaserSettingsDialog(self.backend, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.settings = dict(self.backend.laser_quote_settings() or {})
        current_machine = self.machine_combo.currentText().strip()
        current_commercial = self.commercial_combo.currentText().strip()
        _set_combo_values(self.machine_combo, list(dict(self.settings.get("machine_profiles", {}) or {}).keys()), current_machine)
        _set_combo_values(self.commercial_combo, list(dict(self.settings.get("commercial_profiles", {}) or {}).keys()), current_commercial)
        self._refresh_materials()

    def _payload(self) -> dict[str, Any]:
        return {
            "path": self.file_edit.text().strip(),
            "description": self.description_edit.text().strip(),
            "ref_externa": self.ref_ext_edit.text().strip(),
            "machine_name": self.machine_combo.currentText().strip(),
            "commercial_name": self.commercial_combo.currentText().strip(),
            "material": _canonical_material_family(self.material_combo.currentText().strip()) or self.material_combo.currentText().strip(),
            "material_subtype": self.subtype_combo.currentText().strip(),
            "gas": self.gas_combo.currentText().strip(),
            "thickness_mm": self.thickness_spin.value(),
            "qtd": int(self.quantity_spin.value()),
            "include_marking": bool(self.marking_check.isChecked()),
            "include_defilm": bool(self.defilm_check.isChecked()),
            "material_supplied_by_client": bool(self.customer_material_check.isChecked()),
            "cut_length_m_override": self.cut_override_spin.value() if self.cut_override_spin.value() > 0 else None,
            "mark_length_m_override": self.mark_override_spin.value() if self.mark_override_spin.value() > 0 else None,
            "net_area_m2_override": self.area_override_spin.value() if self.area_override_spin.value() > 0 else None,
            "pierce_count_override": int(self.pierce_override_spin.value()) if self.pierce_override_spin.value() > 0 else None,
        }

    def _apply_analysis(self, result: dict[str, Any]) -> None:
        self.analysis = dict(result or {})
        self.line_payload = dict(self.analysis.get("line_suggestion", {}) or {})
        metrics = dict(self.analysis.get("metrics", {}) or {})
        bbox = dict(dict(self.analysis.get("geometry", {}) or {}).get("bbox_mm", {}) or {})
        times = dict(self.analysis.get("times", {}) or {})
        pricing = dict(self.analysis.get("pricing", {}) or {})
        machine = dict(self.analysis.get("machine", {}) or {})
        commercial = dict(self.analysis.get("commercial", {}) or {})
        self.metric_labels["bbox"].setText(f"{_fmt_num(bbox.get('width', 0), 1)} x {_fmt_num(bbox.get('height', 0), 1)} mm")
        self.metric_labels["cut"].setText(f"{_fmt_num(metrics.get('cut_length_m', 0), 3)} m")
        self.metric_labels["mark"].setText(f"{_fmt_num(metrics.get('mark_length_m', 0), 3)} m")
        self.metric_labels["pierce"].setText(str(int(metrics.get("pierce_count", 0) or 0)))
        self.metric_labels["area"].setText(f"{_fmt_num(metrics.get('net_area_m2', 0), 4)} m2")
        self.metric_labels["weight"].setText(f"{_fmt_num(metrics.get('gross_mass_kg', 0), 3)} kg")
        self.metric_labels["speed"].setText(f"{_fmt_num(machine.get('effective_cut_speed_m_min', 0), 2)} m/min")
        self.metric_labels["time"].setText(f"{_fmt_num(times.get('machine_total_min', 0), 2)} min")
        self.price_labels["material"].setText(_fmt_eur(pricing.get("material_cost_unit", 0)))
        self.price_labels["cut"].setText(_fmt_eur(pricing.get("effective_cutting_cost_unit", 0)) + f" ({commercial.get('effective_cutting_label', '-')})")
        self.price_labels["machine_ref"].setText(_fmt_eur(pricing.get("machine_runtime_cost_unit", 0)))
        self.price_labels["mark"].setText(_fmt_eur(pricing.get("mark_cost_unit", 0)))
        self.price_labels["defilm"].setText(_fmt_eur(pricing.get("defilm_cost_unit", 0)))
        self.price_labels["pierce"].setText(_fmt_eur(pricing.get("pierce_cost_unit", 0)))
        self.price_labels["subtotal"].setText(_fmt_eur(pricing.get("subtotal_cost_unit", 0)))
        self.price_labels["unit"].setText(_fmt_eur(pricing.get("unit_price", 0)))
        self.price_labels["total"].setText(_fmt_eur(pricing.get("total_price", 0)))
        warnings = list(self.analysis.get("warnings", []) or [])
        if not warnings:
            warnings = ["Sem alertas relevantes."]
        self.warning_edit.setPlainText("\n".join(f"- {row}" for row in warnings))
        if not self.description_edit.text().strip():
            self.description_edit.setText(str(self.line_payload.get("descricao", "") or "").strip())
        if not self.ref_ext_edit.text().strip():
            self.ref_ext_edit.setText(str(self.line_payload.get("ref_externa", "") or "").strip())

    def _analyze(self) -> bool:
        if not self.file_edit.text().strip():
            QMessageBox.warning(self, "Peca Unit. DXF/DWG", "Seleciona primeiro um ficheiro DXF ou DWG.")
            return False
        try:
            result = self.backend.laser_quote_build_line(self._payload())
        except Exception as exc:
            QMessageBox.critical(self, "Peca Unit. DXF/DWG", str(exc))
            return False
        self._apply_analysis(dict(result.get("analysis", {}) or {}))
        return True

    def _accept_payload(self) -> None:
        if not self.analysis and not self._analyze():
            return
        if not self.line_payload:
            QMessageBox.warning(self, "Peca Unit. DXF/DWG", "Nao foi possivel gerar a linha de orcamento.")
            return
        self.line_payload["descricao"] = self.description_edit.text().strip() or str(self.line_payload.get("descricao", "") or "").strip()
        self.line_payload["ref_externa"] = self.ref_ext_edit.text().strip() or str(self.line_payload.get("ref_externa", "") or "").strip()
        self.accept()

    def result_payload(self) -> dict[str, Any]:
        return {
            "analysis": dict(self.analysis or {}),
            "line": dict(self.line_payload or {}),
        }
