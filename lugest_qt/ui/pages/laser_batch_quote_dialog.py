from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
)

from ..widgets import CardFrame
from .laser_quote_dialogs import (
    LaserSettingsDialog,
    _canonical_material_family,
    _combo_items,
    _display_material_family,
    _fmt_eur,
    _fmt_num,
    _guess_material_family,
    _set_combo_values,
    _settings_gas_names,
    _settings_material_names,
    _settings_material_subtypes,
    _spin,
)


class LaserBatchQuoteDialog(QDialog):
    def __init__(self, backend, parent=None, *, default_machine: str = "") -> None:
        super().__init__(parent)
        self.backend = backend
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self.line_payloads: list[dict[str, Any]] = []
        self.summary: dict[str, Any] = {}
        self.setWindowTitle("Lote DXF/DWG")
        self.resize(1180, 860)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        top_card = CardFrame()
        top_layout = QGridLayout(top_card)
        top_layout.setContentsMargins(12, 10, 12, 10)
        top_layout.setHorizontalSpacing(8)
        top_layout.setVerticalSpacing(6)
        self.machine_combo = QComboBox()
        self.machine_combo.addItems(list(dict(self.settings.get("machine_profiles", {}) or {}).keys()))
        self.machine_combo.setCurrentText(str(default_machine or self.settings.get("active_machine", "") or self.machine_combo.currentText()))
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
        self.marking_check = QCheckBox("Contar marcacao")
        self.defilm_check = QCheckBox("Contar defilm")
        self.customer_material_check = QCheckBox("Material do cliente (sem materia-prima)")
        self.customer_material_check.setToolTip("Mantem corte/tempo/processo, mas retira o custo de materia-prima do lote orcamentado.")
        config_btn = QPushButton("Configurar perfis")
        config_btn.setProperty("variant", "secondary")
        config_btn.clicked.connect(self._configure_profiles)
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
        top_layout.addWidget(QLabel("Espessura (mm)"), 2, 0)
        top_layout.addWidget(self.marking_check, 2, 2)
        top_layout.addWidget(self.defilm_check, 2, 3)
        top_layout.addWidget(self.customer_material_check, 2, 4)
        top_layout.addWidget(self.thickness_spin, 3, 0)
        top_layout.addWidget(config_btn, 3, 2, 1, 2)
        root.addWidget(top_card)

        batch_card = CardFrame()
        batch_layout = QVBoxLayout(batch_card)
        batch_layout.setContentsMargins(12, 10, 12, 10)
        batch_layout.setSpacing(8)
        batch_header = QHBoxLayout()
        batch_title = QLabel("Pecas do lote")
        batch_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.batch_info_label = QLabel("Seleciona varios DXF/DWG da mesma espessura e preenche as quantidades.")
        self.batch_info_label.setProperty("role", "muted")
        batch_header.addWidget(batch_title)
        batch_header.addStretch(1)
        batch_header.addWidget(self.batch_info_label)
        batch_layout.addLayout(batch_header)
        self.batch_table = QTableWidget(0, 7)
        self.batch_table.setHorizontalHeaderLabels(["Ficheiro", "Descricao", "Ref. externa", "Qtd", "Tempo", "Preco unit.", "Total"])
        self.batch_table.verticalHeader().setVisible(False)
        self.batch_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.batch_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.SelectedClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
        )
        batch_header_view = self.batch_table.horizontalHeader()
        batch_header_view.setSectionResizeMode(0, QHeaderView.Stretch)
        batch_header_view.setSectionResizeMode(1, QHeaderView.Stretch)
        batch_header_view.setSectionResizeMode(2, QHeaderView.Stretch)
        for col_index in (3, 4, 5, 6):
            batch_header_view.setSectionResizeMode(col_index, QHeaderView.ResizeToContents)
        batch_layout.addWidget(self.batch_table, 1)
        batch_actions = QHBoxLayout()
        self.add_files_btn = QPushButton("Selecionar DXF/DWG")
        self.remove_files_btn = QPushButton("Remover selecionada")
        self.remove_files_btn.setProperty("variant", "secondary")
        self.clear_files_btn = QPushButton("Limpar lote")
        self.clear_files_btn.setProperty("variant", "danger")
        batch_actions.addWidget(self.add_files_btn)
        batch_actions.addWidget(self.remove_files_btn)
        batch_actions.addWidget(self.clear_files_btn)
        batch_actions.addStretch(1)
        batch_layout.addLayout(batch_actions)
        root.addWidget(batch_card, 1)

        bottom = QHBoxLayout()
        bottom.setSpacing(10)
        root.addLayout(bottom)

        summary_card = CardFrame()
        summary_layout = QGridLayout(summary_card)
        summary_layout.setContentsMargins(12, 10, 12, 10)
        summary_layout.setHorizontalSpacing(10)
        summary_layout.setVerticalSpacing(6)
        self.summary_labels: dict[str, QLabel] = {}
        for row_index, (key, title) in enumerate(
            (
                ("files", "Ficheiros"),
                ("pieces", "Pecas"),
                ("time", "Tempo total"),
                ("unit", "Preco medio"),
                ("total", "Preco total"),
            )
        ):
            label_title = QLabel(title)
            label_title.setProperty("role", "field_label")
            label_value = QLabel("-")
            label_value.setProperty("role", "field_value_strong" if key in ("unit", "total") else "field_value")
            summary_layout.addWidget(label_title, row_index, 0)
            summary_layout.addWidget(label_value, row_index, 1)
            self.summary_labels[key] = label_value
        bottom.addWidget(summary_card, 2)

        warning_card = CardFrame()
        warning_layout = QVBoxLayout(warning_card)
        warning_layout.setContentsMargins(12, 10, 12, 10)
        warning_layout.setSpacing(6)
        warning_title = QLabel("Alertas e observacoes")
        warning_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.warning_edit = QTextEdit()
        self.warning_edit.setReadOnly(True)
        warning_layout.addWidget(warning_title)
        warning_layout.addWidget(self.warning_edit, 1)
        bottom.addWidget(warning_card, 3)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.analyze_btn = QPushButton("Orcamentar lote")
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
        self.batch_table.itemSelectionChanged.connect(self._sync_buttons)
        self.add_files_btn.clicked.connect(self._pick_files)
        self.remove_files_btn.clicked.connect(self._remove_selected_rows)
        self.clear_files_btn.clicked.connect(self._clear_rows)
        self._refresh_materials()
        self._sync_buttons()

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
        if current_subtype and current_subtype not in _combo_items(self.subtype_combo):
            self.subtype_combo.setCurrentText(current_subtype)
        self._refresh_gases()

    def _refresh_gases(self) -> None:
        current_gas = self.gas_combo.currentText().strip()
        values = _settings_gas_names(self.settings, self.machine_combo.currentText().strip(), self.material_combo.currentText().strip())
        _set_combo_values(self.gas_combo, values, current_gas if current_gas else values[0])

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

    def _suggest_from_path(self, path: str) -> tuple[str, str, str]:
        normalized = str(path or "").replace("\\", "/")
        file_name = normalized.split("/")[-1]
        stem = file_name.rsplit(".", 1)[0]
        desc = stem.replace("_", " ").strip()
        ref = []
        for char in stem.upper():
            ref.append(char if char.isalnum() or char in ("-", "_") else "-")
        return file_name, desc, "".join(ref).strip("-_") or "PECA-LASER"

    def _make_item(self, value: Any, *, center: bool = False, editable: bool = True) -> QTableWidgetItem:
        item = QTableWidgetItem(str(value or "").strip())
        item.setTextAlignment(int((Qt.AlignCenter if center else Qt.AlignLeft) | Qt.AlignVCenter))
        if not editable:
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        return item

    def _set_result_cell(self, row_index: int, col_index: int, text: str) -> None:
        item = self.batch_table.item(row_index, col_index)
        if item is None:
            item = self._make_item(text, center=True, editable=False)
            self.batch_table.setItem(row_index, col_index, item)
        else:
            item.setText(str(text))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))

    def _pick_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecionar varios DXF/DWG",
            "",
            "Desenhos CAD (*.dxf *.dwg);;DXF (*.dxf);;DWG (*.dwg);;Todos (*.*)",
        )
        if not paths:
            return
        existing = {self._row_path(row_index) for row_index in range(self.batch_table.rowCount())}
        for path in paths:
            clean_path = str(path or "").strip()
            if not clean_path or clean_path in existing:
                continue
            file_name, desc, ref = self._suggest_from_path(clean_path)
            row_index = self.batch_table.rowCount()
            self.batch_table.insertRow(row_index)
            file_item = self._make_item(file_name, editable=False)
            file_item.setData(Qt.UserRole, clean_path)
            file_item.setToolTip(clean_path)
            self.batch_table.setItem(row_index, 0, file_item)
            self.batch_table.setItem(row_index, 1, self._make_item(desc))
            self.batch_table.setItem(row_index, 2, self._make_item(ref))
            self.batch_table.setItem(row_index, 3, self._make_item("1", center=True))
            self._set_result_cell(row_index, 4, "-")
            self._set_result_cell(row_index, 5, "-")
            self._set_result_cell(row_index, 6, "-")
            existing.add(clean_path)
        self.line_payloads = []
        self.summary = {}
        self._clear_summary()
        self._sync_buttons()

    def _row_path(self, row_index: int) -> str:
        item = self.batch_table.item(row_index, 0)
        return str((item.data(Qt.UserRole) if item is not None else "") or "").strip()

    def _row_text(self, row_index: int, col_index: int) -> str:
        item = self.batch_table.item(row_index, col_index)
        return str(item.text() if item is not None else "").strip()

    def _row_quantity(self, row_index: int) -> int:
        text = self._row_text(row_index, 3).replace(",", ".")
        try:
            value = int(round(float(text or 1)))
        except Exception:
            value = 1
        return max(1, value)

    def _remove_selected_rows(self) -> None:
        rows = sorted({item.row() for item in self.batch_table.selectedItems()}, reverse=True)
        if not rows:
            return
        for row_index in rows:
            self.batch_table.removeRow(row_index)
        self.line_payloads = []
        self.summary = {}
        self._clear_summary()
        self._sync_buttons()

    def _clear_rows(self) -> None:
        self.batch_table.setRowCount(0)
        self.line_payloads = []
        self.summary = {}
        self._clear_summary()
        self._sync_buttons()

    def _sync_buttons(self) -> None:
        has_selection = bool(self.batch_table.selectedItems())
        self.remove_files_btn.setEnabled(has_selection)
        self.clear_files_btn.setEnabled(self.batch_table.rowCount() > 0)
        total_pieces = sum(self._row_quantity(row_index) for row_index in range(self.batch_table.rowCount()))
        if self.batch_table.rowCount() > 0:
            self.batch_info_label.setText(f"{self.batch_table.rowCount()} ficheiros | {total_pieces} pecas totais | espessura global {self.thickness_spin.value():g} mm")
        else:
            self.batch_info_label.setText("Seleciona varios DXF/DWG da mesma espessura e preenche as quantidades.")

    def _clear_summary(self) -> None:
        for label in self.summary_labels.values():
            label.setText("-")
        self.warning_edit.clear()

    def _base_payload(self) -> dict[str, Any]:
        return {
            "machine_name": self.machine_combo.currentText().strip(),
            "commercial_name": self.commercial_combo.currentText().strip(),
            "material": _canonical_material_family(self.material_combo.currentText().strip()) or self.material_combo.currentText().strip(),
            "material_subtype": self.subtype_combo.currentText().strip(),
            "gas": self.gas_combo.currentText().strip(),
            "thickness_mm": self.thickness_spin.value(),
            "include_marking": bool(self.marking_check.isChecked()),
            "include_defilm": bool(self.defilm_check.isChecked()),
            "material_supplied_by_client": bool(self.customer_material_check.isChecked()),
        }

    def _row_payload(self, row_index: int) -> dict[str, Any]:
        payload = self._base_payload()
        payload.update(
            {
                "path": self._row_path(row_index),
                "description": self._row_text(row_index, 1),
                "ref_externa": self._row_text(row_index, 2),
                "qtd": self._row_quantity(row_index),
            }
        )
        return payload

    def _analyze(self) -> bool:
        if self.batch_table.rowCount() == 0:
            QMessageBox.warning(self, "Lote DXF/DWG", "Seleciona primeiro os ficheiros DXF ou DWG do lote.")
            return False
        results: list[dict[str, Any]] = []
        warnings: list[str] = []
        total_parts = 0
        total_time = 0.0
        total_price = 0.0
        for row_index in range(self.batch_table.rowCount()):
            path = self._row_path(row_index)
            if not path:
                QMessageBox.warning(self, "Lote DXF/DWG", f"A linha {row_index + 1} nao tem ficheiro associado.")
                return False
            try:
                result = self.backend.laser_quote_build_line(self._row_payload(row_index))
            except Exception as exc:
                QMessageBox.critical(self, "Lote DXF/DWG", f"Erro ao analisar {path}:\n{exc}")
                return False
            analysis = dict(result.get("analysis", {}) or {})
            line = dict(result.get("line", {}) or {})
            pricing = dict(analysis.get("pricing", {}) or {})
            times = dict(analysis.get("times", {}) or {})
            geometry = dict(analysis.get("geometry", {}) or {})
            qty = int(pricing.get("quantity", 1) or 1)
            total_parts += qty
            total_time += float(times.get("machine_total_min", 0) or 0) * qty
            total_price += float(pricing.get("total_price", 0) or 0)
            self._set_result_cell(row_index, 4, f"{_fmt_num((times.get('machine_total_min', 0) or 0) * qty, 2)} min")
            self._set_result_cell(row_index, 5, _fmt_eur(pricing.get("unit_price", 0)))
            self._set_result_cell(row_index, 6, _fmt_eur(pricing.get("total_price", 0)))
            for warn in [str(item or "").strip() for item in list(analysis.get("warnings", []) or []) if str(item or "").strip()]:
                warnings.append(f"{str(geometry.get('file_name', '') or line.get('ref_externa', '') or 'DXF').strip()}: {warn}")
            results.append({"analysis": analysis, "line": line})
        self.line_payloads = [dict(item.get("line", {}) or {}) for item in results if dict(item.get("line", {}) or {})]
        average_price = (total_price / total_parts) if total_parts else 0.0
        first_analysis = dict(dict(results[0].get("analysis", {}) or {}) if results else {})
        self.summary = {
            "machine": dict(first_analysis.get("machine", {}) or {}),
            "commercial": dict(first_analysis.get("commercial", {}) or {}),
            "warnings": warnings,
            "pricing": {
                "quantity": total_parts,
                "unit_price": round(average_price, 4),
                "total_price": round(total_price, 2),
            },
            "times": {
                "machine_total_min": round(total_time, 3),
            },
            "batch_count": len(results),
        }
        self.summary_labels["files"].setText(str(len(results)))
        self.summary_labels["pieces"].setText(str(total_parts))
        self.summary_labels["time"].setText(f"{_fmt_num(total_time, 2)} min")
        self.summary_labels["unit"].setText(_fmt_eur(average_price))
        self.summary_labels["total"].setText(_fmt_eur(total_price))
        if not warnings:
            warnings = ["Sem alertas relevantes."]
        self.warning_edit.setPlainText("\n".join(f"- {row}" for row in warnings))
        return bool(self.line_payloads)

    def _accept_payload(self) -> None:
        if not self._analyze():
            return
        self.accept()

    def result_payload(self) -> dict[str, Any]:
        return {
            "analysis": dict(self.summary or {}),
            "lines": [dict(row or {}) for row in list(self.line_payloads or [])],
        }
