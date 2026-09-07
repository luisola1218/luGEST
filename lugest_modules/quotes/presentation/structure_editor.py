from __future__ import annotations
import math
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QFrame, QGridLayout, QLabel, QLineEdit, QMessageBox, QScrollArea, QVBoxLayout, QWidget
from datetime import datetime
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from functools import partial
from lugest_modules.quotes.domain.lines import service_line as build_service_line, product_line as build_product_line
def edit_structure(owner: QWidget, backend: QuoteEditorPorts, *, workcenter: str) -> dict | None:
    dialog = QDialog(owner)
    dialog.setWindowTitle("Orcamento Estruturas Metalicas")
    dialog.resize(1060, 780)
    dialog.setMinimumSize(920, 700)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(8)

    family_presets = {
        "Leve": {"profile_kg_m2": 18.0, "tube_kg_m2": 5.5, "frame_spacing_m": 6.0, "fab_h_m2": 0.16, "assembly_h_m2": 0.11, "eng_h_m2": 0.018},
        "Standard": {"profile_kg_m2": 24.0, "tube_kg_m2": 7.5, "frame_spacing_m": 5.0, "fab_h_m2": 0.24, "assembly_h_m2": 0.16, "eng_h_m2": 0.028},
        "Reforcada": {"profile_kg_m2": 31.0, "tube_kg_m2": 9.0, "frame_spacing_m": 4.5, "fab_h_m2": 0.31, "assembly_h_m2": 0.21, "eng_h_m2": 0.04},
    }
    cladding_presets = {
        "Sem revestimento": {"roof_price": 0.0, "facade_price": 0.0, "roof_factor": 0.0, "facade_factor": 0.0, "name": "Sem revestimento"},
        "Chapa simples": {"roof_price": 19.5, "facade_price": 16.5, "roof_factor": 1.0, "facade_factor": 1.0, "name": "Chapa simples"},
        "Sandwich cobertura": {"roof_price": 32.0, "facade_price": 0.0, "roof_factor": 1.0, "facade_factor": 0.0, "name": "Sandwich cobertura"},
        "Sandwich completo": {"roof_price": 34.0, "facade_price": 29.0, "roof_factor": 1.0, "facade_factor": 1.0, "name": "Sandwich completo"},
    }
    finish_presets = {
        "Sem acabamento": {"price_m2": 0.0, "paint_factor": 0.0, "operation": "Serralharia"},
        "Pintura": {"price_m2": 8.5, "paint_factor": 1.15, "operation": "Pintura"},
        "Galvanizacao": {"price_m2": 11.75, "paint_factor": 1.05, "operation": "Lacagem"},
        "Metalizacao + pintura": {"price_m2": 16.2, "paint_factor": 1.25, "operation": "Lacagem"},
    }
    product_rows = [dict(row) for row in list(backend.ne_product_options("") or []) if isinstance(row, dict)]
    product_rows_by_code = {
        str(row.get("codigo", "") or "").strip(): row
        for row in product_rows
        if str(row.get("codigo", "") or "").strip()
    }

    intro = QLabel(
        "Versao 2 do modelo de estruturas metalicas. "
        "Gera uma proposta mais tecnica para pavilhoes, coberturas e estruturas especiais, "
        "com familia estrutural, pórticos, revestimento, acabamento, acessorios e produtos reais de stock."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    intro.setStyleSheet("font-size: 11px;")
    layout.addWidget(intro)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    scroll_host = QWidget()
    scroll.setWidget(scroll_host)
    scroll_layout = QVBoxLayout(scroll_host)
    scroll_layout.setContentsMargins(0, 0, 0, 0)
    scroll_layout.setSpacing(8)
    layout.addWidget(scroll, 1)

    top_grid = QGridLayout()
    top_grid.setHorizontalSpacing(10)
    top_grid.setVerticalSpacing(8)
    scroll_layout.addLayout(top_grid)

    project_card = CardFrame()
    project_card.set_tone("default")
    project_form = QFormLayout(project_card)
    project_form.setContentsMargins(10, 10, 10, 10)
    project_form.setHorizontalSpacing(10)
    project_form.setVerticalSpacing(6)
    project_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

    structure_name_edit = QLineEdit()
    structure_name_edit.setPlaceholderText("Ex.: Pavilhao logistico cliente")
    structure_type_combo = QComboBox()
    structure_type_combo.addItems(["Pavilhao industrial", "Cobertura metalica", "Mezanino", "Estrutura especial"])
    family_combo = QComboBox()
    family_combo.addItems(list(family_presets.keys()))
    qty_spin = QDoubleSpinBox()
    qty_spin.setRange(1.0, 100.0)
    qty_spin.setDecimals(0)
    qty_spin.setValue(1.0)
    length_spin = QDoubleSpinBox()
    length_spin.setRange(1.0, 500.0)
    length_spin.setDecimals(2)
    length_spin.setSuffix(" m")
    length_spin.setValue(30.0)
    width_spin = QDoubleSpinBox()
    width_spin.setRange(1.0, 200.0)
    width_spin.setDecimals(2)
    width_spin.setSuffix(" m")
    width_spin.setValue(18.0)
    height_spin = QDoubleSpinBox()
    height_spin.setRange(1.0, 50.0)
    height_spin.setDecimals(2)
    height_spin.setSuffix(" m")
    height_spin.setValue(6.0)
    roof_slope_spin = QDoubleSpinBox()
    roof_slope_spin.setRange(0.0, 100.0)
    roof_slope_spin.setDecimals(1)
    roof_slope_spin.setSuffix(" %")
    roof_slope_spin.setValue(12.0)
    frame_spacing_spin = QDoubleSpinBox()
    frame_spacing_spin.setRange(2.0, 12.0)
    frame_spacing_spin.setDecimals(2)
    frame_spacing_spin.setSuffix(" m")
    frame_spacing_spin.setValue(family_presets["Standard"]["frame_spacing_m"])
    project_form.addRow("Designacao", structure_name_edit)
    project_form.addRow("Tipologia", structure_type_combo)
    project_form.addRow("Familia estrutural", family_combo)
    project_form.addRow("Qtd. estruturas", qty_spin)
    project_form.addRow("Comprimento", length_spin)
    project_form.addRow("Largura", width_spin)
    project_form.addRow("Altura util", height_spin)
    project_form.addRow("Inclinacao cobertura", roof_slope_spin)
    project_form.addRow("Espacamento pórticos", frame_spacing_spin)

    cost_card = CardFrame()
    cost_card.set_tone("default")
    cost_form = QFormLayout(cost_card)
    cost_form.setContentsMargins(10, 10, 10, 10)
    cost_form.setHorizontalSpacing(10)
    cost_form.setVerticalSpacing(6)
    cost_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

    profile_kg_spin = QDoubleSpinBox()
    profile_kg_spin.setRange(0.0, 500.0)
    profile_kg_spin.setDecimals(2)
    profile_kg_spin.setSuffix(" kg/m2")
    profile_kg_spin.setValue(24.0)
    profile_price_spin = QDoubleSpinBox()
    profile_price_spin.setRange(0.0, 1000.0)
    profile_price_spin.setDecimals(3)
    profile_price_spin.setPrefix("EUR ")
    profile_price_spin.setValue(3.15)
    tube_kg_spin = QDoubleSpinBox()
    tube_kg_spin.setRange(0.0, 500.0)
    tube_kg_spin.setDecimals(2)
    tube_kg_spin.setSuffix(" kg/m2")
    tube_kg_spin.setValue(7.5)
    tube_price_spin = QDoubleSpinBox()
    tube_price_spin.setRange(0.0, 1000.0)
    tube_price_spin.setDecimals(3)
    tube_price_spin.setPrefix("EUR ")
    tube_price_spin.setValue(3.05)
    cladding_combo = QComboBox()
    cladding_combo.addItems(list(cladding_presets.keys()))
    roof_price_spin = QDoubleSpinBox()
    roof_price_spin.setRange(0.0, 1000.0)
    roof_price_spin.setDecimals(2)
    roof_price_spin.setPrefix("EUR ")
    roof_price_spin.setSuffix("/m2")
    roof_price_spin.setValue(cladding_presets["Sandwich completo"]["roof_price"])
    facade_price_spin = QDoubleSpinBox()
    facade_price_spin.setRange(0.0, 1000.0)
    facade_price_spin.setDecimals(2)
    facade_price_spin.setPrefix("EUR ")
    facade_price_spin.setSuffix("/m2")
    facade_price_spin.setValue(cladding_presets["Sandwich completo"]["facade_price"])
    finish_combo = QComboBox()
    finish_combo.addItems(list(finish_presets.keys()))
    paint_factor_spin = QDoubleSpinBox()
    paint_factor_spin.setRange(0.0, 5.0)
    paint_factor_spin.setDecimals(2)
    paint_factor_spin.setValue(finish_presets["Pintura"]["paint_factor"])
    paint_price_spin = QDoubleSpinBox()
    paint_price_spin.setRange(0.0, 1000.0)
    paint_price_spin.setDecimals(2)
    paint_price_spin.setPrefix("EUR ")
    paint_price_spin.setSuffix("/m2")
    paint_price_spin.setValue(finish_presets["Pintura"]["price_m2"])
    cost_form.addRow("Perfis principais", profile_kg_spin)
    cost_form.addRow("Preco perfis", profile_price_spin)
    cost_form.addRow("Tubos / travamentos", tube_kg_spin)
    cost_form.addRow("Preco tubos", tube_price_spin)
    cost_form.addRow("Revestimento", cladding_combo)
    cost_form.addRow("Cobertura / remates", roof_price_spin)
    cost_form.addRow("Fachadas / fechamentos", facade_price_spin)
    cost_form.addRow("Acabamento", finish_combo)
    cost_form.addRow("Fator acabamento", paint_factor_spin)
    cost_form.addRow("Preco acabamento", paint_price_spin)

    labour_card = CardFrame()
    labour_card.set_tone("default")
    labour_form = QFormLayout(labour_card)
    labour_form.setContentsMargins(10, 10, 10, 10)
    labour_form.setHorizontalSpacing(10)
    labour_form.setVerticalSpacing(6)
    labour_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

    fab_hours_spin = QDoubleSpinBox()
    fab_hours_spin.setRange(0.0, 5000.0)
    fab_hours_spin.setDecimals(1)
    fab_hours_spin.setSuffix(" h")
    fab_hours_spin.setValue(140.0)
    fab_rate_spin = QDoubleSpinBox()
    fab_rate_spin.setRange(0.0, 1000.0)
    fab_rate_spin.setDecimals(2)
    fab_rate_spin.setPrefix("EUR ")
    fab_rate_spin.setSuffix("/h")
    fab_rate_spin.setValue(28.0)
    assembly_hours_spin = QDoubleSpinBox()
    assembly_hours_spin.setRange(0.0, 5000.0)
    assembly_hours_spin.setDecimals(1)
    assembly_hours_spin.setSuffix(" h")
    assembly_hours_spin.setValue(90.0)
    assembly_rate_spin = QDoubleSpinBox()
    assembly_rate_spin.setRange(0.0, 1000.0)
    assembly_rate_spin.setDecimals(2)
    assembly_rate_spin.setPrefix("EUR ")
    assembly_rate_spin.setSuffix("/h")
    assembly_rate_spin.setValue(32.0)
    engineering_hours_spin = QDoubleSpinBox()
    engineering_hours_spin.setRange(0.0, 1000.0)
    engineering_hours_spin.setDecimals(1)
    engineering_hours_spin.setSuffix(" h")
    engineering_hours_spin.setValue(24.0)
    engineering_rate_spin = QDoubleSpinBox()
    engineering_rate_spin.setRange(0.0, 1000.0)
    engineering_rate_spin.setDecimals(2)
    engineering_rate_spin.setPrefix("EUR ")
    engineering_rate_spin.setSuffix("/h")
    engineering_rate_spin.setValue(35.0)
    crane_days_spin = QDoubleSpinBox()
    crane_days_spin.setRange(0.0, 365.0)
    crane_days_spin.setDecimals(1)
    crane_days_spin.setSuffix(" dias")
    crane_days_spin.setValue(2.0)
    crane_day_rate_spin = QDoubleSpinBox()
    crane_day_rate_spin.setRange(0.0, 1000000.0)
    crane_day_rate_spin.setDecimals(2)
    crane_day_rate_spin.setPrefix("EUR ")
    crane_day_rate_spin.setSuffix("/dia")
    crane_day_rate_spin.setValue(450.0)
    extras_desc_edit = QLineEdit()
    extras_desc_edit.setPlaceholderText("Ex.: portas, caleiras, platibandas, acessorios")
    extras_value_spin = QDoubleSpinBox()
    extras_value_spin.setRange(0.0, 1000000.0)
    extras_value_spin.setDecimals(2)
    extras_value_spin.setPrefix("EUR ")
    labour_form.addRow("Horas de fabrico", fab_hours_spin)
    labour_form.addRow("Preco fabrico", fab_rate_spin)
    labour_form.addRow("Horas de montagem", assembly_hours_spin)
    labour_form.addRow("Preco montagem", assembly_rate_spin)
    labour_form.addRow("Horas engenharia", engineering_hours_spin)
    labour_form.addRow("Preco engenharia", engineering_rate_spin)
    labour_form.addRow("Grua / elevacao", crane_days_spin)
    labour_form.addRow("Preco grua", crane_day_rate_spin)
    labour_form.addRow("Extras", extras_desc_edit)
    labour_form.addRow("Valor extras", extras_value_spin)

    accessory_card = CardFrame()
    accessory_card.set_tone("default")
    accessory_layout = QVBoxLayout(accessory_card)
    accessory_layout.setContentsMargins(10, 10, 10, 10)
    accessory_layout.setSpacing(6)
    accessory_title = QLabel("Acessorios / produtos stock")
    accessory_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
    accessory_hint = QLabel("Seleciona produtos reais de stock para incluir na proposta, como parafusaria, componentes ou apoio de montagem.")
    accessory_hint.setProperty("role", "muted")
    accessory_hint.setWordWrap(True)
    accessory_layout.addWidget(accessory_title)
    accessory_layout.addWidget(accessory_hint)
    accessory_grid = QGridLayout()
    accessory_grid.setHorizontalSpacing(8)
    accessory_grid.setVerticalSpacing(6)
    accessory_layout.addLayout(accessory_grid)
    accessory_controls: list[dict[str, object]] = []
    for idx in range(3):
        product_combo = QComboBox()
        product_combo.setEditable(True)
        product_combo.addItem("")
        for row in product_rows:
            code = str(row.get("codigo", "") or "").strip()
            label = f"{code} - {str(row.get('descricao', '') or '').strip()}".strip(" -")
            product_combo.addItem(label, code)
        qty_combo_spin = QDoubleSpinBox()
        qty_combo_spin.setRange(0.0, 1000000.0)
        qty_combo_spin.setDecimals(2)
        qty_combo_spin.setValue(0.0)
        accessory_grid.addWidget(QLabel(f"Produto {idx + 1}"), idx, 0)
        accessory_grid.addWidget(product_combo, idx, 1)
        accessory_grid.addWidget(QLabel("Qtd."), idx, 2)
        accessory_grid.addWidget(qty_combo_spin, idx, 3)
        accessory_controls.append({"combo": product_combo, "qty": qty_combo_spin})

    top_grid.addWidget(project_card, 0, 0)
    top_grid.addWidget(cost_card, 0, 1)
    top_grid.addWidget(labour_card, 1, 0, 1, 2)
    top_grid.addWidget(accessory_card, 2, 0, 1, 2)
    top_grid.setColumnStretch(0, 1)
    top_grid.setColumnStretch(1, 1)

    summary_card = CardFrame()
    summary_card.set_tone("info")
    summary_layout = QVBoxLayout(summary_card)
    summary_layout.setContentsMargins(10, 8, 10, 8)
    summary_layout.setSpacing(5)
    summary_title = QLabel("Resumo tecnico e comercial")
    summary_title.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
    summary_text = QLabel("")
    summary_text.setWordWrap(True)
    summary_text.setTextInteractionFlags(Qt.TextSelectableByMouse)
    summary_layout.addWidget(summary_title)
    summary_layout.addWidget(summary_text)
    scroll_layout.addWidget(summary_card)

    buttons = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
    ok_button = buttons.button(QDialogButtonBox.Ok)
    if ok_button is not None:
        ok_button.setText("Gerar orcamento")
    buttons.rejected.connect(dialog.reject)
    buttons.accepted.connect(dialog.accept)
    layout.addWidget(buttons)

    def _apply_family_preset() -> None:
        preset = dict(family_presets.get(family_combo.currentText().strip(), family_presets["Standard"]))
        profile_kg_spin.setValue(float(preset.get("profile_kg_m2", profile_kg_spin.value()) or 0))
        tube_kg_spin.setValue(float(preset.get("tube_kg_m2", tube_kg_spin.value()) or 0))
        frame_spacing_spin.setValue(float(preset.get("frame_spacing_m", frame_spacing_spin.value()) or 0))

    def _apply_cladding_preset() -> None:
        preset = dict(cladding_presets.get(cladding_combo.currentText().strip(), cladding_presets["Sem revestimento"]))
        roof_price_spin.setValue(float(preset.get("roof_price", roof_price_spin.value()) or 0))
        facade_price_spin.setValue(float(preset.get("facade_price", facade_price_spin.value()) or 0))

    def _apply_finish_preset() -> None:
        preset = dict(finish_presets.get(finish_combo.currentText().strip(), finish_presets["Sem acabamento"]))
        paint_factor_spin.setValue(float(preset.get("paint_factor", paint_factor_spin.value()) or 0))
        paint_price_spin.setValue(float(preset.get("price_m2", paint_price_spin.value()) or 0))

    def _compute_payload() -> dict:
        structure_name = structure_name_edit.text().strip() or structure_type_combo.currentText().strip() or "Estrutura metalica"
        assembly_code = f"EST-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        quantity = float(qty_spin.value() or 1.0)
        length = float(length_spin.value() or 0.0)
        width = float(width_spin.value() or 0.0)
        height = float(height_spin.value() or 0.0)
        slope_pct = float(roof_slope_spin.value() or 0.0)
        frame_spacing = max(2.0, float(frame_spacing_spin.value() or 0.0))
        family_name = family_combo.currentText().strip() or "Standard"
        family_preset = dict(family_presets.get(family_name, family_presets["Standard"]))
        cladding_preset = dict(cladding_presets.get(cladding_combo.currentText().strip(), cladding_presets["Sem revestimento"]))
        finish_name = finish_combo.currentText().strip() or "Sem acabamento"
        finish_preset = dict(finish_presets.get(finish_name, finish_presets["Sem acabamento"]))
        footprint_area = round(length * width * quantity, 2)
        roof_factor = math.sqrt(1.0 + ((slope_pct / 100.0) ** 2))
        roof_area = round(length * width * roof_factor * quantity * float(cladding_preset.get("roof_factor", 1.0) or 0.0), 2)
        perimeter = round((2.0 * (length + width)) * quantity, 2)
        facade_area = round(perimeter * height * float(cladding_preset.get("facade_factor", 1.0) or 0.0), 2)
        frame_count_single = max(2, int(math.ceil(length / frame_spacing)) + 1)
        portal_count = int(frame_count_single * quantity)
        column_count = int(portal_count * 2)
        steel_profiles_kg = round(footprint_area * float(profile_kg_spin.value() or 0.0), 2)
        steel_tubes_kg = round(footprint_area * float(tube_kg_spin.value() or 0.0), 2)
        finish_area = round((roof_area + facade_area) * float(paint_factor_spin.value() or 0.0), 2)
        fab_hours = round(max(float(fab_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("fab_h_m2", 0.0) or 0.0)), 1)
        assembly_hours = round(max(float(assembly_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("assembly_h_m2", 0.0) or 0.0)), 1)
        engineering_hours = round(max(float(engineering_hours_spin.value() or 0.0), footprint_area * float(family_preset.get("eng_h_m2", 0.0) or 0.0)), 1)

        lines: list[dict] = []
        for row in (
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Perfis estruturais {family_name} | {structure_name}", steel_profiles_kg, "kg", float(profile_price_spin.value() or 0.0), "Serralharia"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Tubos e travamentos | {structure_name}", steel_tubes_kg, "kg", float(tube_price_spin.value() or 0.0), "Serralharia"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Cobertura e remates | {structure_name}", roof_area, "m2", float(roof_price_spin.value() or 0.0), "Montagem"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Fachadas e fechamentos | {structure_name}", facade_area, "m2", float(facade_price_spin.value() or 0.0), "Montagem"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"{finish_name} / protecao estrutural | {structure_name}", finish_area, "m2", float(paint_price_spin.value() or 0.0), str(finish_preset.get("operation", "Pintura") or "Pintura")),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Fabrico e soldadura | {structure_name}", fab_hours, "h", float(fab_rate_spin.value() or 0.0), "Serralharia"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Montagem em obra | {structure_name}", assembly_hours, "h", float(assembly_rate_spin.value() or 0.0), "Montagem"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Engenharia e preparacao | {structure_name}", engineering_hours, "h", float(engineering_rate_spin.value() or 0.0), "Serralharia"),
            partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"Grua / elevacao | {structure_name}", float(crane_days_spin.value() or 0.0), "dia", float(crane_day_rate_spin.value() or 0.0), "Montagem"),
        ):
            if isinstance(row, dict):
                lines.append(row)

        for control in accessory_controls:
            combo = control.get("combo")
            qty_widget = control.get("qty")
            if not isinstance(combo, QComboBox) or not isinstance(qty_widget, QDoubleSpinBox):
                continue
            code = str(combo.currentData() or "").strip()
            if not code:
                text = combo.currentText().strip()
                code = text.split(" - ", 1)[0].strip()
            product_line = partial(build_product_line, line_type=backend.ORC_LINE_TYPE_PRODUCT)(
                product_rows_by_code.get(code),
                float(qty_widget.value() or 0.0),
                descricao_extra=structure_name,
            )
            if isinstance(product_line, dict):
                lines.append(product_line)

        extras_value = float(extras_value_spin.value() or 0.0)
        extras_desc = extras_desc_edit.text().strip() or "Extras de estrutura"
        extra_line = partial(build_service_line, line_type=backend.ORC_LINE_TYPE_SERVICE)(f"{extras_desc} | {structure_name}", 1.0, "SV", extras_value, "Montagem")
        if isinstance(extra_line, dict):
            lines.append(extra_line)

        estimated_total = round(sum(float(row.get("total", 0) or 0.0) for row in lines), 2)
        note_cliente = structure_name
        note_lines = [
            f"Modelo estrutura: {structure_type_combo.currentText().strip()}",
            f"Familia: {family_name} | acabamento: {finish_name} | revestimento: {cladding_preset.get('name', cladding_combo.currentText().strip())}",
            f"Dimensoes base: {length:.2f} x {width:.2f} x {height:.2f} m | quantidade {quantity:.0f}",
            f"Porticos: {portal_count} | colunas: {column_count} | espacamento medio: {frame_spacing:.2f} m",
            f"Area implantacao: {footprint_area:.2f} m2 | cobertura: {roof_area:.2f} m2 | fachadas: {facade_area:.2f} m2",
            f"Perfis: {steel_profiles_kg:.2f} kg | tubos: {steel_tubes_kg:.2f} kg | acabamento: {finish_area:.2f} m2",
        ]
        return {
            "name": structure_name,
            "type": structure_type_combo.currentText().strip(),
            "assembly_code": assembly_code,
            "assembly_name": structure_name,
            "workcenter": workcenter,
            "note_cliente": note_cliente,
            "notes_pdf": "\n".join(note_lines),
            "summary_html": (
                f"{structure_name} | {structure_type_combo.currentText().strip()} | familia {family_name} | "
                f"porticos {portal_count} | implantacao {footprint_area:.2f} m2 | cobertura {roof_area:.2f} m2 | "
                f"fachadas {facade_area:.2f} m2 | perfis {steel_profiles_kg:.2f} kg | tubos {steel_tubes_kg:.2f} kg | "
                f"acabamento {finish_name} | total base {_fmt_eur(estimated_total)}"
            ),
            "lines": lines,
        }

    def _refresh_summary() -> None:
        payload = _compute_payload()
        lines = list(payload.get("lines", []) or [])
        if not lines:
            summary_text.setText("Preenche valores comerciais para gerar uma base de orcamento.")
            return
        summary_rows = [
            str(payload.get("summary_html", "") or "").strip(),
            f"Linhas geradas: {len(lines)}",
            "Conjunto gerado como mini-projeto: materiais, mao de obra, consumiveis/produtos stock e extras.",
            "Categorias: perfis, tubos, cobertura, fachadas, acabamento, fabrico, montagem, engenharia, grua, extras e produtos de stock.",
            "Depois de gerar, continuas com acesso total ao orcamento normal para acrescentar pecas, produtos ou linhas manuais.",
        ]
        summary_text.setText("\n".join([row for row in summary_rows if row]))

    family_combo.currentTextChanged.connect(lambda _text: _apply_family_preset())
    cladding_combo.currentTextChanged.connect(lambda _text: _apply_cladding_preset())
    finish_combo.currentTextChanged.connect(lambda _text: _apply_finish_preset())
    for widget in (
        structure_name_edit,
        structure_type_combo,
        family_combo,
        qty_spin,
        length_spin,
        width_spin,
        height_spin,
        roof_slope_spin,
        frame_spacing_spin,
        profile_kg_spin,
        profile_price_spin,
        tube_kg_spin,
        tube_price_spin,
        cladding_combo,
        roof_price_spin,
        facade_price_spin,
        finish_combo,
        paint_factor_spin,
        paint_price_spin,
        fab_hours_spin,
        fab_rate_spin,
        assembly_hours_spin,
        assembly_rate_spin,
        engineering_hours_spin,
        engineering_rate_spin,
        crane_days_spin,
        crane_day_rate_spin,
        extras_desc_edit,
        extras_value_spin,
    ):
        if isinstance(widget, (QLineEdit, QComboBox)):
            signal = widget.textChanged if isinstance(widget, QLineEdit) else widget.currentTextChanged
            signal.connect(_refresh_summary)
        elif isinstance(widget, QDoubleSpinBox):
            widget.valueChanged.connect(lambda _value: _refresh_summary())
    for control in accessory_controls:
        combo = control.get("combo")
        qty_widget = control.get("qty")
        if isinstance(combo, QComboBox):
            combo.currentTextChanged.connect(lambda _text: _refresh_summary())
        if isinstance(qty_widget, QDoubleSpinBox):
            qty_widget.valueChanged.connect(lambda _value: _refresh_summary())
    _apply_family_preset()
    _apply_cladding_preset()
    _apply_finish_preset()
    _refresh_summary()

    if dialog.exec() != QDialog.Accepted:
        return None
    payload = _compute_payload()
    if not list(payload.get("lines", []) or []):
        QMessageBox.warning(owner, "Orcamento Estruturas", "Nao foi gerada nenhuma linha para o modelo de estruturas.")
        return None
    return payload
