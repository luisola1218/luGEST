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
    QWidget,
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
        self.resize(1340, 860)
        self.setMinimumWidth(1240)

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
        self.batch_table = QTableWidget(0, 10)
        self.batch_table.setHorizontalHeaderLabels(
            ["Ficheiro", "Descricao", "Ref. externa", "Qtd", "Operacoes", "Tempo ops", "Preco ops", "Tempo", "Preco unit.", "Total"]
        )
        self.batch_table.setStyleSheet(
            "QTableWidget { font-size: 11px; }"
            "QHeaderView::section { font-size: 11px; padding: 6px 6px; }"
            "QPushButton { font-size: 9px; padding: 3px 5px; min-height: 24px; }"
        )
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
        batch_header_view.setSectionResizeMode(3, QHeaderView.Fixed)
        batch_header_view.resizeSection(3, 48)
        batch_header_view.setSectionResizeMode(4, QHeaderView.Fixed)
        batch_header_view.resizeSection(4, 190)
        for col_index, width in ((5, 92), (6, 98), (7, 88), (8, 96), (9, 86)):
            batch_header_view.setSectionResizeMode(col_index, QHeaderView.Fixed)
            batch_header_view.resizeSection(col_index, width)
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
        self.machine_combo.currentTextChanged.connect(lambda _text: self._invalidate_analysis())
        self.material_combo.currentTextChanged.connect(self._refresh_subtypes)
        self.batch_table.itemSelectionChanged.connect(self._sync_buttons)
        self.add_files_btn.clicked.connect(self._pick_files)
        self.remove_files_btn.clicked.connect(self._remove_selected_rows)
        self.clear_files_btn.clicked.connect(self._clear_rows)
        self.batch_table.itemChanged.connect(self._handle_table_item_changed)
        self.commercial_combo.currentTextChanged.connect(lambda _text: self._invalidate_analysis())
        self.material_combo.currentTextChanged.connect(lambda _text: self._invalidate_analysis())
        self.subtype_combo.currentTextChanged.connect(lambda _text: self._invalidate_analysis())
        self.gas_combo.currentTextChanged.connect(lambda _text: self._invalidate_analysis())
        self.thickness_spin.valueChanged.connect(lambda _value: self._invalidate_analysis())
        self.marking_check.toggled.connect(lambda _checked: self._invalidate_analysis())
        self.defilm_check.toggled.connect(lambda _checked: self._invalidate_analysis())
        self.customer_material_check.toggled.connect(lambda _checked: self._invalidate_analysis())
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

    def _row_operation_meta(self, row_index: int) -> dict[str, Any]:
        item = self.batch_table.item(row_index, 0)
        data = item.data(Qt.UserRole + 1) if item is not None else {}
        return dict(data or {}) if isinstance(data, dict) else {}

    def _edit_sender_operations(self) -> None:
        sender = self.sender()
        for row_index in range(self.batch_table.rowCount()):
            if self.batch_table.cellWidget(row_index, 4) is sender:
                self._edit_row_operations(row_index)
                return

    def _set_row_operation_meta(self, row_index: int, meta: dict[str, Any]) -> None:
        item = self.batch_table.item(row_index, 0)
        if item is not None:
            item.setData(Qt.UserRole + 1, dict(meta or {}))
        extra_names = [
            str(row.get("nome", "") or "").strip()
            for row in list(dict(meta or {}).get("operacoes_detalhe", []) or [])
            if str(row.get("nome", "") or "").strip()
        ]
        names = ["Corte Laser"] + [name for name in extra_names if name != "Corte Laser"]
        label = " + ".join(names)
        button = self.batch_table.cellWidget(row_index, 4)
        if isinstance(button, QPushButton):
            button.setText(label if len(label) <= 36 else label[:33] + "...")
            button.setToolTip(label)
        self._set_result_cell(row_index, 5, f"{_fmt_num(float(dict(meta or {}).get('tempo_ops_unit', 0.0) or 0.0), 3)} min")
        self._set_result_cell(row_index, 6, _fmt_eur(float(dict(meta or {}).get("preco_ops_unit", 0.0) or 0.0)))

    def _invalidate_analysis(self) -> None:
        self.line_payloads = []
        self.summary = {}
        self._clear_summary()
        self._sync_buttons()

    def _recalculate_row_operation_meta(self, row_index: int) -> None:
        meta = self._row_operation_meta(row_index)
        detail_rows = [dict(row or {}) for row in list(meta.get("operacoes_detalhe", []) or []) if isinstance(row, dict)]
        if not detail_rows:
            self._set_row_operation_meta(row_index, {})
            return
        estimate = dict(
            self.backend.operation_cost_estimate(
                {
                    "qtd": self._row_quantity(row_index),
                    "costing_operations": [str(row.get("nome", "") or "").strip() for row in detail_rows],
                    "operacoes_detalhe": detail_rows,
                }
            )
            or {}
        )
        summary = dict(estimate.get("summary", {}) or {})
        final_rows = [dict(row or {}) for row in list(estimate.get("operations", []) or []) if isinstance(row, dict)]
        self._set_row_operation_meta(
            row_index,
            {
                "operacoes_detalhe": final_rows,
                "tempos_operacao": {
                    str(row.get("nome", "") or "").strip(): float(row.get("tempo_unit_min", 0.0) or 0.0)
                    for row in final_rows
                    if str(row.get("nome", "") or "").strip() and row.get("tempo_unit_min") not in (None, "")
                },
                "custos_operacao": {
                    str(row.get("nome", "") or "").strip(): float(row.get("custo_unit_eur", 0.0) or 0.0)
                    for row in final_rows
                    if str(row.get("nome", "") or "").strip() and row.get("custo_unit_eur") not in (None, "")
                },
                "tempo_ops_unit": round(float(summary.get("tempo_unit_total_min", 0.0) or 0.0), 4),
                "preco_ops_unit": round(float(summary.get("custo_unit_total_eur", 0.0) or 0.0), 4),
                "quote_cost_snapshot": {
                    "costing_mode": str(summary.get("costing_mode", "") or ""),
                    "tempo_total_peca_min": round(float(summary.get("tempo_unit_total_min", 0.0) or 0.0), 4),
                    "preco_unit_total_eur": round(float(summary.get("custo_unit_total_eur", 0.0) or 0.0), 4),
                    "qtd": self._row_quantity(row_index),
                },
            },
        )

    def _operation_names(self) -> list[str]:
        try:
            settings = dict(self.backend.operation_cost_settings() or {})
            active = str(settings.get("active_profile", "Base") or "Base").strip() or "Base"
            profile = dict(dict(settings.get("profiles", {}) or {}).get(active, {}) or {})
            names = [str(name or "").strip() for name in profile.keys() if str(name or "").strip()]
        except Exception:
            names = []
        defaults = ["Quinagem", "Roscagem", "Serralharia", "Maquinacao", "Soldadura", "Pintura", "Lacagem", "Montagem", "Embalamento"]
        out: list[str] = []
        for name in defaults + names:
            clean = str(name or "").strip()
            if clean and clean != "Corte Laser" and clean not in out:
                out.append(clean)
        return out

    def _edit_row_operations(self, row_index: int) -> None:
        if row_index < 0 or row_index >= self.batch_table.rowCount():
            return
        meta = self._row_operation_meta(row_index)
        selected_map = {str(row.get("nome", "") or "").strip(): dict(row or {}) for row in list(meta.get("operacoes_detalhe", []) or [])}
        dialog = QDialog(self)
        dialog.setWindowTitle("Operacoes da peca")
        dialog.resize(1120, 520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        info = QLabel("Define aqui as operacoes extra desta linha. Corte Laser continua a ser calculado pelo DXF; estas linhas somam tempo e preco por unidade.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        layout.addWidget(info)

        table = QTableWidget(0, 9)
        table.setHorizontalHeaderLabels(["Usar", "Operacao", "Modo", "Tipo qtd.", "Qtd/peca", "Setup", "Tempo base", "EUR/h", "Fixo/manual"])
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        for col in (4, 5, 6, 7, 8):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)

        controls: list[dict[str, Any]] = []
        modes = [("per_feature", "Por quantidade"), ("per_piece", "Por peca"), ("per_area_m2", "Por area"), ("manual", "Manual")]
        for op_name in self._operation_names():
            estimate = dict(
                self.backend.operation_cost_estimate(
                    {
                        "qtd": self._row_quantity(row_index),
                        "costing_operations": [op_name],
                        "operacoes_detalhe": [selected_map.get(op_name, {})] if op_name in selected_map else [],
                    }
                )
                or {}
            )
            op_row = dict((list(estimate.get("operations", []) or [{}]) or [{}])[0] or {})
            current = dict(selected_map.get(op_name, op_row) or {})
            row = table.rowCount()
            table.insertRow(row)
            check = QCheckBox()
            check.setChecked(op_name in selected_map)
            mode_combo = QComboBox()
            for key, label in modes:
                mode_combo.addItem(label, key)
            target_mode = str(current.get("pricing_mode", op_row.get("pricing_mode", "per_feature")) or "per_feature")
            for idx in range(mode_combo.count()):
                if str(mode_combo.itemData(idx) or "") == target_mode:
                    mode_combo.setCurrentIndex(idx)
                    break
            driver_edit = self._make_item(str(current.get("driver_label", op_row.get("driver_label", "Qtd./peca")) or "Qtd./peca"), editable=False)
            driver_spin = _spin(4, 0.0, 1000000.0, float(current.get("driver_units", op_row.get("driver_units", 1.0)) or 0.0), 1.0)
            setup_spin = _spin(4, 0.0, 1000000.0, float(current.get("setup_min", op_row.get("setup_min", 0.0)) or 0.0), 0.25)
            time_spin = _spin(4, 0.0, 1000000.0, float(current.get("unit_time_base_min", current.get("tempo_unit_min", op_row.get("unit_time_base_min", 0.0))) or 0.0), 0.05)
            hour_spin = _spin(4, 0.0, 1000000.0, float(current.get("hour_rate_eur", op_row.get("hour_rate_eur", 0.0)) or 0.0), 1.0)
            fixed_spin = _spin(4, 0.0, 1000000.0, float(current.get("fixed_unit_eur", current.get("custo_unit_eur", op_row.get("fixed_unit_eur", 0.0))) or 0.0), 0.1)
            table.setCellWidget(row, 0, check)
            table.setItem(row, 1, self._make_item(op_name, editable=False))
            table.setCellWidget(row, 2, mode_combo)
            table.setItem(row, 3, driver_edit)
            table.setCellWidget(row, 4, driver_spin)
            table.setCellWidget(row, 5, setup_spin)
            table.setCellWidget(row, 6, time_spin)
            table.setCellWidget(row, 7, hour_spin)
            table.setCellWidget(row, 8, fixed_spin)
            table.setRowHeight(row, 36)
            controls.append(
                {
                    "nome": op_name,
                    "check": check,
                    "mode": mode_combo,
                    "driver_label": driver_edit.text(),
                    "driver_units": driver_spin,
                    "setup": setup_spin,
                    "time": time_spin,
                    "hour": hour_spin,
                    "fixed": fixed_spin,
                }
            )

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        detail_rows: list[dict[str, Any]] = []
        for row in controls:
            if not row["check"].isChecked():
                continue
            mode = str(row["mode"].currentData() or "manual")
            detail_rows.append(
                {
                    "nome": row["nome"],
                    "pricing_mode": mode,
                    "driver_label": str(row["driver_label"] or "Qtd./peca"),
                    "driver_units": row["driver_units"].value(),
                    "driver_units_confirmed": True,
                    "setup_min": row["setup"].value(),
                    "unit_time_base_min": row["time"].value(),
                    "hour_rate_eur": row["hour"].value(),
                    "fixed_unit_eur": row["fixed"].value(),
                    "manual_values_confirmed": mode == "manual",
                    "tempo_unit_min": row["time"].value() if mode == "manual" else None,
                    "custo_unit_eur": row["fixed"].value() if mode == "manual" else None,
                }
            )
        estimate = dict(
            self.backend.operation_cost_estimate(
                {
                    "qtd": self._row_quantity(row_index),
                    "costing_operations": [str(row.get("nome", "") or "").strip() for row in detail_rows],
                    "operacoes_detalhe": detail_rows,
                }
            )
            or {}
        )
        summary = dict(estimate.get("summary", {}) or {})
        final_rows = [dict(row or {}) for row in list(estimate.get("operations", []) or []) if isinstance(row, dict)]
        self._set_row_operation_meta(
            row_index,
            {
                "operacoes_detalhe": final_rows,
                "tempos_operacao": {
                    str(row.get("nome", "") or "").strip(): float(row.get("tempo_unit_min", 0.0) or 0.0)
                    for row in final_rows
                    if str(row.get("nome", "") or "").strip() and row.get("tempo_unit_min") not in (None, "")
                },
                "custos_operacao": {
                    str(row.get("nome", "") or "").strip(): float(row.get("custo_unit_eur", 0.0) or 0.0)
                    for row in final_rows
                    if str(row.get("nome", "") or "").strip() and row.get("custo_unit_eur") not in (None, "")
                },
                "tempo_ops_unit": round(float(summary.get("tempo_unit_total_min", 0.0) or 0.0), 4),
                "preco_ops_unit": round(float(summary.get("custo_unit_total_eur", 0.0) or 0.0), 4),
                "quote_cost_snapshot": {
                    "costing_mode": str(summary.get("costing_mode", "") or ""),
                    "tempo_total_peca_min": round(float(summary.get("tempo_unit_total_min", 0.0) or 0.0), 4),
                    "preco_unit_total_eur": round(float(summary.get("custo_unit_total_eur", 0.0) or 0.0), 4),
                    "qtd": self._row_quantity(row_index),
                },
            },
        )
        self._invalidate_analysis()

    def _apply_row_operations_to_line(self, line: dict[str, Any], row_index: int) -> dict[str, Any]:
        meta = self._row_operation_meta(row_index)
        detail_rows = [dict(row or {}) for row in list(meta.get("operacoes_detalhe", []) or []) if isinstance(row, dict)]
        if not detail_rows:
            return dict(line or {})
        row = dict(line or {})
        base_operation = str(row.get("operacao", "") or "Corte Laser").strip() or "Corte Laser"
        operations = [part.strip() for part in base_operation.split("+") if part.strip()]
        extra_time = float(meta.get("tempo_ops_unit", 0.0) or 0.0)
        extra_cost = float(meta.get("preco_ops_unit", 0.0) or 0.0)
        tempos = dict(row.get("tempos_operacao", {}) or {})
        custos = dict(row.get("custos_operacao", {}) or {})
        details = [dict(item or {}) for item in list(row.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        for operation in detail_rows:
            name = str(operation.get("nome", "") or "").strip()
            if not name:
                continue
            if name not in operations:
                operations.append(name)
            time_value = round(float(operation.get("tempo_unit_min", 0.0) or 0.0), 3)
            cost_value = round(float(operation.get("custo_unit_eur", 0.0) or 0.0), 4)
            if time_value > 0:
                tempos[name] = time_value
            if cost_value > 0:
                custos[name] = cost_value
            details.append(dict(operation))
        row["operacao"] = " + ".join(operations)
        row["operacoes_lista"] = operations
        row["operacoes_detalhe"] = details
        row["tempos_operacao"] = tempos
        row["custos_operacao"] = custos
        row["tempo_peca_min"] = round(float(row.get("tempo_peca_min", 0.0) or 0.0) + extra_time, 3)
        row["preco_unit"] = round(float(row.get("preco_unit", 0.0) or 0.0) + extra_cost, 4)
        row["total"] = round(float(row.get("qtd", 0.0) or 0.0) * float(row.get("preco_unit", 0.0) or 0.0), 2)
        return row

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
            ops_btn = QPushButton("Corte Laser")
            ops_btn.setProperty("variant", "secondary")
            ops_btn.clicked.connect(self._edit_sender_operations)
            self.batch_table.setCellWidget(row_index, 4, ops_btn)
            self._set_row_operation_meta(row_index, {})
            self._set_result_cell(row_index, 7, "-")
            self._set_result_cell(row_index, 8, "-")
            self._set_result_cell(row_index, 9, "-")
            existing.add(clean_path)
        self.line_payloads = []
        self.summary = {}
        self._clear_summary()
        self._sync_buttons()

    def _handle_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item is None:
            return
        row_index = item.row()
        if row_index < 0:
            return
        if item.column() == 3:
            self._recalculate_row_operation_meta(row_index)
        self._invalidate_analysis()

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
                self._recalculate_row_operation_meta(row_index)
                result = self.backend.laser_quote_build_line(self._row_payload(row_index))
            except Exception as exc:
                QMessageBox.critical(self, "Lote DXF/DWG", f"Erro ao analisar {path}:\n{exc}")
                return False
            analysis = dict(result.get("analysis", {}) or {})
            line = self._apply_row_operations_to_line(dict(result.get("line", {}) or {}), row_index)
            pricing = dict(analysis.get("pricing", {}) or {})
            times = dict(analysis.get("times", {}) or {})
            geometry = dict(analysis.get("geometry", {}) or {})
            qty = int(line.get("qtd", pricing.get("quantity", 1)) or 1)
            unit_price = float(line.get("preco_unit", pricing.get("unit_price", 0)) or 0.0)
            line_total = round(unit_price * qty, 2)
            unit_time = float(line.get("tempo_peca_min", times.get("machine_total_min", 0)) or 0.0)
            total_parts += qty
            total_time += unit_time * qty
            total_price += line_total
            self._set_result_cell(row_index, 7, f"{_fmt_num(unit_time * qty, 2)} min")
            self._set_result_cell(row_index, 8, _fmt_eur(unit_price))
            self._set_result_cell(row_index, 9, _fmt_eur(line_total))
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
