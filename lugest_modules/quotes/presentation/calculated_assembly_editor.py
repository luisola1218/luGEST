"""Standalone calculated assembly editor with explicit actions and initial context."""
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.pages.runtime_common import configure_table as _configure_table, fill_table as _fill_table

@dataclass(frozen=True)
class CalculatedAssemblyPorts:
    quote_number: str
    workcenter: str
    assembly_model_save: Callable
    conjunto_next_param_codigo: Callable
    conjunto_save: Callable
    laser_batch_dialog: Callable
    norm_text: Callable
    orc_line_is_piece: Callable
    _assembly_item_kind_from_line: Callable
    _consumable_assembly_item_dialog: Callable
    _edit_laser_batch_lines: Callable
    _labor_assembly_item_dialog: Callable
    _material_assembly_item_dialog: Callable
    _open_profile_step_igs_quote_builder: Callable
    _product_assembly_item_dialog: Callable
    _quote_line_is_laser_2d: Callable
    _quote_pick_workcenter: Callable
    _structure_quote_dialog: Callable
    _wrap_assembly_item: Callable


def edit_calculated_assembly(owner: QWidget, ports: CalculatedAssemblyPorts, initial: dict | None = None) -> dict | None:
    initial = dict(initial or {})
    dialog = QDialog(owner)
    dialog.setWindowTitle("Conjunto calculado")
    dialog.resize(1040, 840)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(8)

    intro = QLabel(
        "Cria um conjunto como mini-projeto: materiais, mao de obra, consumiveis e produtos. "
        "Cada conjunto fica guardado no catalogo. A opcao abaixo permite marca-lo como modelo reutilizavel."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(intro)

    header_tabs = QTabWidget()
    identity_tab = QWidget()
    header_form = QFormLayout(identity_tab)
    header_form.setContentsMargins(10, 10, 10, 10)
    header_form.setHorizontalSpacing(10)
    header_form.setVerticalSpacing(6)
    code_edit = QLineEdit(str(initial.get("codigo", "") or f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}").strip())
    try:
        suggested_param = str(ports.conjunto_next_param_codigo() or "").strip()
    except Exception:
        suggested_param = ""
    param_edit = QLineEdit(str(initial.get("param_codigo", "") or suggested_param or "0001").strip())
    param_edit.setReadOnly(True)
    param_edit.setToolTip("Codigo sequencial e permanente da parametrizacao do conjunto")
    name_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
    name_edit.setPlaceholderText("Ex.: Construcao de Bascula")
    margin_spin = QDoubleSpinBox()
    margin_spin.setRange(0.0, 500.0)
    margin_spin.setDecimals(2)
    margin_spin.setSuffix(" %")
    margin_spin.setValue(float(initial.get("margem_perc", 30.0) or 30.0))
    save_template_check = QCheckBox("Marcar tambem como template de conjunto")
    save_template_check.setChecked(bool(initial.get("template", False)))
    notes_edit = QTextEdit()
    notes_edit.setMaximumHeight(70)
    notes_edit.setPlainText(str(initial.get("notas", "") or "").strip())
    header_form.addRow("Codigo conjunto", code_edit)
    header_form.addRow("Codigo parametrizacao", param_edit)
    header_form.addRow("Descricao", name_edit)
    header_form.addRow("Margem", margin_spin)
    header_form.addRow("Notas", notes_edit)
    header_form.addRow("", save_template_check)
    header_tabs.addTab(identity_tab, "Identificacao e valor")

    technical = dict(initial.get("ficha_tecnica", {}) or {})
    technical_tab = QWidget()
    technical_grid = QGridLayout(technical_tab)
    technical_grid.setContentsMargins(10, 10, 10, 10)
    technical_grid.setHorizontalSpacing(10)
    technical_grid.setVerticalSpacing(6)

    family_combo = QComboBox()
    family_combo.setEditable(True)
    family_combo.addItems([
        "Equipamento de pesagem",
        "Quiosque multimedia",
        "Caixilharia",
        "Estrutura metalica",
        "Maquina / equipamento",
        "Mobiliario tecnico",
        "Outro produto fabricado",
    ])
    family_value = str(technical.get("familia_produto", "") or "").strip()
    if family_value:
        family_combo.setCurrentText(family_value)
    application_edit = QLineEdit(str(technical.get("aplicacao", "") or "").strip())
    application_edit.setPlaceholderText("Ex.: pesagem industrial, atendimento publico, fachada exterior")
    model_edit = QLineEdit(str(technical.get("modelo_versao", "") or "").strip())
    model_edit.setPlaceholderText("Modelo, variante ou revisao")
    configuration_edit = QLineEdit(str(technical.get("configuracao", "") or "").strip())
    configuration_edit.setPlaceholderText("Ex.: 1500 kg / visor remoto / 2 folhas / RAL 7016")
    dimensions_edit = QLineEdit(str(technical.get("dimensoes_gerais", "") or "").strip())
    dimensions_edit.setPlaceholderText("C x L x A, vao, capacidade ou formato relevante")
    finishes_edit = QLineEdit(str(technical.get("materiais_acabamentos", "") or "").strip())
    finishes_edit.setPlaceholderText("Materiais principais, cor e acabamento")

    characteristics_edit = QTextEdit()
    characteristics_edit.setMaximumHeight(66)
    characteristics_edit.setPlaceholderText("Funcoes, desempenho, opcoes e caracteristicas essenciais")
    characteristics_edit.setPlainText(str(technical.get("caracteristicas", "") or "").strip())
    installation_edit = QTextEdit()
    installation_edit.setMaximumHeight(66)
    installation_edit.setPlaceholderText("Alimentacao, fixacao, ligacoes, ambiente e pre-requisitos")
    installation_edit.setPlainText(str(technical.get("requisitos_instalacao", "") or "").strip())
    standards_edit = QTextEdit()
    standards_edit.setMaximumHeight(58)
    standards_edit.setPlaceholderText("Normas, diretivas, classe, IP, CE ou requisitos do cliente")
    standards_edit.setPlainText(str(technical.get("normas_conformidade", "") or "").strip())
    quality_edit = QTextEdit()
    quality_edit.setMaximumHeight(58)
    quality_edit.setPlaceholderText("Inspecoes, ensaios, tolerancias e criterios de aceitacao")
    quality_edit.setPlainText(str(technical.get("controlo_qualidade", "") or "").strip())

    technical_grid.addWidget(QLabel("Familia de produto"), 0, 0)
    technical_grid.addWidget(family_combo, 0, 1)
    technical_grid.addWidget(QLabel("Aplicacao / destino"), 0, 2)
    technical_grid.addWidget(application_edit, 0, 3)
    technical_grid.addWidget(QLabel("Modelo / versao"), 1, 0)
    technical_grid.addWidget(model_edit, 1, 1)
    technical_grid.addWidget(QLabel("Configuracao principal"), 1, 2)
    technical_grid.addWidget(configuration_edit, 1, 3)
    technical_grid.addWidget(QLabel("Dimensoes / capacidade"), 2, 0)
    technical_grid.addWidget(dimensions_edit, 2, 1)
    technical_grid.addWidget(QLabel("Materiais / acabamentos"), 2, 2)
    technical_grid.addWidget(finishes_edit, 2, 3)
    technical_grid.addWidget(QLabel("Caracteristicas"), 3, 0)
    technical_grid.addWidget(characteristics_edit, 3, 1)
    technical_grid.addWidget(QLabel("Instalacao"), 3, 2)
    technical_grid.addWidget(installation_edit, 3, 3)
    technical_grid.addWidget(QLabel("Normas / conformidade"), 4, 0)
    technical_grid.addWidget(standards_edit, 4, 1)
    technical_grid.addWidget(QLabel("Controlo de qualidade"), 4, 2)
    technical_grid.addWidget(quality_edit, 4, 3)
    technical_grid.setColumnStretch(1, 1)
    technical_grid.setColumnStretch(3, 1)
    header_tabs.addTab(technical_tab, "Ficha tecnica do produto")
    layout.addWidget(header_tabs)

    items: list[dict] = [ports._wrap_assembly_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
    table = QTableWidget(0, 8)
    table.setHorizontalHeaderLabels(["Categoria", "Descricao", "Ref./Cod.", "Qtd", "Unid", "Peso", "Preco", "Total"])
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QTableWidget.NoEditTriggers)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    _configure_table(table, stretch=(1,), contents=(0, 2, 3, 4, 5, 6, 7))
    layout.addWidget(table, 1)

    actions = QGridLayout()
    actions.setHorizontalSpacing(8)
    actions.setVerticalSpacing(6)
    add_material_btn = QPushButton("Material")
    add_labor_btn = QPushButton("Mao de obra")
    add_consumable_btn = QPushButton("Consumivel")
    add_product_btn = QPushButton("Produto stock")
    add_laser_batch_btn = QPushButton("Lote DXF/DWG")
    add_step_igs_btn = QPushButton("STEP/IGS")
    add_structure_btn = QPushButton("Estrutura metalica")
    edit_btn = QPushButton("Editar item")
    edit_btn.setProperty("variant", "secondary")
    remove_btn = QPushButton("Remover item")
    remove_btn.setProperty("variant", "danger")
    for idx, button in enumerate((
        add_material_btn,
        add_labor_btn,
        add_consumable_btn,
        add_product_btn,
        add_laser_batch_btn,
        add_step_igs_btn,
        add_structure_btn,
        edit_btn,
        remove_btn,
    )):
        button.setProperty("compact", "true")
        actions.addWidget(button, idx // 3, idx % 3)
    layout.addLayout(actions)

    summary_card = CardFrame()
    summary_card.set_tone("info")
    summary_layout = QGridLayout(summary_card)
    summary_layout.setContentsMargins(10, 8, 10, 8)
    summary_layout.setHorizontalSpacing(12)
    summary_layout.setVerticalSpacing(6)
    material_total_label = QLabel("0,00 EUR")
    labor_total_label = QLabel("0,00 EUR")
    consumable_total_label = QLabel("0,00 EUR")
    product_total_label = QLabel("0,00 EUR")
    subtotal_label = QLabel("0,00 EUR")
    final_total_label = QLabel("0,00 EUR")
    for text, widget, row, col in (
        ("Materiais", material_total_label, 0, 0),
        ("Mao de obra", labor_total_label, 0, 2),
        ("Consumiveis", consumable_total_label, 1, 0),
        ("Produtos", product_total_label, 1, 2),
        ("Subtotal custo", subtotal_label, 2, 0),
        ("Preco final c/margem", final_total_label, 2, 2),
    ):
        summary_layout.addWidget(QLabel(text), row, col)
        summary_layout.addWidget(widget, row, col + 1)
    layout.addWidget(summary_card)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    def _selected_index() -> int:
        current = table.currentItem()
        if current is None or current.row() >= len(items):
            return -1
        return current.row()

    def _render_items() -> None:
        _fill_table(
            table,
            [
                [
                    {"material": "Material", "labor": "Mao de obra", "consumable": "Consumivel", "product": "Produto"}.get(str(item.get("kind", "")), "-"),
                    str(((item.get("line") or {}).get("descricao", "") or "-")).strip() or "-",
                    str(((item.get("line") or {}).get("produto_codigo", "") or (item.get("line") or {}).get("ref_externa", "") or "-")).strip() or "-",
                    f"{float(((item.get('line') or {}).get('qtd', 0) or 0)):.2f}",
                    str(((item.get("line") or {}).get("produto_unid", "") or "-")).strip() or "-",
                    f"{float(item.get('weight_total', 0) or 0):.2f} kg" if float(item.get("weight_total", 0) or 0) > 0 else "-",
                    _fmt_eur(float(((item.get("line") or {}).get("preco_unit", 0) or 0))),
                    _fmt_eur(float(item.get("total_cost", 0) or 0)),
                ]
                for item in items
            ],
            align_center_from=3,
        )
        totals = {"material": 0.0, "labor": 0.0, "consumable": 0.0, "product": 0.0}
        for item in items:
            kind = str(item.get("kind", "") or "")
            totals[kind] = totals.get(kind, 0.0) + float(item.get("total_cost", 0) or 0.0)
        subtotal = round(sum(totals.values()), 2)
        final_total = round(subtotal * (1.0 + (float(margin_spin.value() or 0) / 100.0)), 2)
        material_total_label.setText(_fmt_eur(totals.get("material", 0.0)))
        labor_total_label.setText(_fmt_eur(totals.get("labor", 0.0)))
        consumable_total_label.setText(_fmt_eur(totals.get("consumable", 0.0)))
        product_total_label.setText(_fmt_eur(totals.get("product", 0.0)))
        subtotal_label.setText(_fmt_eur(subtotal))
        final_total_label.setText(_fmt_eur(final_total))

    def _open_editor_for(kind: str, current: dict | None = None) -> dict | None:
        if kind == "material":
            return ports._material_assembly_item_dialog(current, parent=dialog)
        if kind == "labor":
            return ports._labor_assembly_item_dialog(current, parent=dialog)
        if kind == "consumable":
            return ports._consumable_assembly_item_dialog(current, parent=dialog)
        if kind == "product":
            return ports._product_assembly_item_dialog(current, parent=dialog)
        return None

    def _add_item(kind: str) -> None:
        try:
            payload = _open_editor_for(kind)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjunto", str(exc))
            return
        if payload is None:
            return
        items.append(payload)
        _render_items()

    def _line_kind_for_shortcut(line: dict) -> str:
        operation = str(line.get("operacao", "") or "").strip().lower()
        unit_txt = str(line.get("produto_unid", "") or "").strip().lower()
        if unit_txt == "h" and "corte laser" not in operation:
            return "labor"
        return ports._assembly_item_kind_from_line(line)

    def _add_lines_from_shortcut(lines: list[dict], *, laser_batch_id: str = "") -> None:
        added = 0
        for row in list(lines or []):
            if not isinstance(row, dict) or not row:
                continue
            line = dict(row)
            if laser_batch_id:
                line["laser_source_mode"] = "batch"
                line["laser_batch_id"] = laser_batch_id
            kind = _line_kind_for_shortcut(line)
            total_cost = round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2)
            items.append(
                {
                    "kind": kind,
                    "quantity_units": round(float(line.get("qtd", 0) or 0), 3),
                    "total_cost": total_cost,
                    "weight_total": 0.0,
                    "line": line,
                }
            )
            added += 1
        if added:
            _render_items()

    def _add_laser_batch_to_assembly() -> None:
        batch_dialog = ports.laser_batch_dialog(dialog,
            default_machine=ports.workcenter,
        )
        if batch_dialog.exec() != QDialog.Accepted:
            return
        result = dict(batch_dialog.result_payload() or {})
        _add_lines_from_shortcut(
            [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})],
            laser_batch_id=str(batch_dialog.batch_id or "").strip(),
        )

    def _add_step_igs_to_assembly() -> None:
        lines = ports._open_profile_step_igs_quote_builder(return_lines=True, parent=dialog)
        _add_lines_from_shortcut([dict(row or {}) for row in list(lines or []) if dict(row or {})])

    def _add_structure_to_assembly() -> None:
        payload = ports._structure_quote_dialog()
        if not payload:
            return
        _add_lines_from_shortcut([dict(row or {}) for row in list(payload.get("lines", []) or []) if dict(row or {})])

    def _edit_item() -> None:
        index = _selected_index()
        if index < 0:
            QMessageBox.warning(dialog, "Conjunto", "Seleciona um item.")
            return
        current = dict(items[index] or {})
        try:
            current_line = dict(current.get("line") or current)
            if ports._quote_line_is_laser_2d(current_line):
                batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
                batch_indexes = [
                    row_index
                    for row_index, candidate in enumerate(items)
                    if batch_id
                    and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
                ] or [index]
                source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
                edited_lines = ports._edit_laser_batch_lines(source_lines, parent=dialog)
                if edited_lines is None:
                    return
                insert_at = min(batch_indexes)
                for row_index in sorted(batch_indexes, reverse=True):
                    del items[row_index]
                for offset, edited_line in enumerate(edited_lines):
                    items.insert(insert_at + offset, ports._wrap_assembly_item(edited_line))
                _render_items()
                if edited_lines:
                    table.selectRow(insert_at)
                return
            payload = _open_editor_for(str(current.get("kind", "") or ""), current)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjunto", str(exc))
            return
        if payload is None:
            return
        items[index] = payload
        _render_items()
        table.selectRow(index)

    def _remove_item() -> None:
        index = _selected_index()
        if index < 0:
            QMessageBox.warning(dialog, "Conjunto", "Seleciona um item.")
            return
        del items[index]
        _render_items()

    add_material_btn.clicked.connect(lambda: _add_item("material"))
    add_labor_btn.clicked.connect(lambda: _add_item("labor"))
    add_consumable_btn.clicked.connect(lambda: _add_item("consumable"))
    add_product_btn.clicked.connect(lambda: _add_item("product"))
    add_laser_batch_btn.clicked.connect(_add_laser_batch_to_assembly)
    add_step_igs_btn.clicked.connect(_add_step_igs_to_assembly)
    add_structure_btn.clicked.connect(_add_structure_to_assembly)
    edit_btn.clicked.connect(_edit_item)
    remove_btn.clicked.connect(_remove_item)
    margin_spin.valueChanged.connect(lambda _v: _render_items())
    _render_items()

    if dialog.exec() != QDialog.Accepted:
        return None
    if not items:
        QMessageBox.warning(owner, "Conjunto", "O conjunto precisa de pelo menos um item.")
        return None
    assembly_code = code_edit.text().strip() or f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    assembly_name = name_edit.text().strip() or assembly_code
    lines = [dict(item.get("line") or {}) for item in items if isinstance(item.get("line"), dict)]
    for line in lines:
        operation_norm = ports.norm_text(str(line.get("operacao", "") or ""))
        if ports.orc_line_is_piece(line) and "laser" in operation_norm:
            line["source_quote_number"] = str(ports.quote_number or "").strip()
            line["source_ref_externa"] = str(line.get("ref_externa", "") or "").strip()
            line["pricing_source"] = "quote_laser"
    totals = {"material": 0.0, "labor": 0.0, "consumable": 0.0, "product": 0.0}
    for item in items:
        kind = str(item.get("kind", "") or "")
        totals[kind] = totals.get(kind, 0.0) + float(item.get("total_cost", 0) or 0.0)
    subtotal = round(sum(totals.values()), 2)
    final_total = round(subtotal * (1.0 + (float(margin_spin.value() or 0) / 100.0)), 2)
    notes_lines = [
        f"Conjunto: {assembly_name}",
        f"Materiais: {_fmt_eur(totals.get('material', 0.0))}",
        f"Mao de obra: {_fmt_eur(totals.get('labor', 0.0))}",
        f"Consumiveis: {_fmt_eur(totals.get('consumable', 0.0))}",
        f"Produtos: {_fmt_eur(totals.get('product', 0.0))}",
        f"Margem aplicada: {float(margin_spin.value() or 0):.2f}%",
        f"Preco final conjunto: {_fmt_eur(final_total)}",
    ]
    if notes_edit.toPlainText().strip():
        notes_lines.append(notes_edit.toPlainText().strip())
    technical_sheet = {
        "familia_produto": family_combo.currentText().strip(),
        "aplicacao": application_edit.text().strip(),
        "modelo_versao": model_edit.text().strip(),
        "configuracao": configuration_edit.text().strip(),
        "dimensoes_gerais": dimensions_edit.text().strip(),
        "materiais_acabamentos": finishes_edit.text().strip(),
        "caracteristicas": characteristics_edit.toPlainText().strip(),
        "requisitos_instalacao": installation_edit.toPlainText().strip(),
        "normas_conformidade": standards_edit.toPlainText().strip(),
        "controlo_qualidade": quality_edit.toPlainText().strip(),
    }
    try:
        ports.assembly_model_save(
            {
                "codigo": assembly_code,
                "param_codigo": param_edit.text().strip(),
                "descricao": assembly_name,
                "notas": "\n".join(notes_lines),
                "itens": lines,
                "template": bool(save_template_check.isChecked()),
                "origem": "orcamento_conjunto_calculado",
                "created_at": str(initial.get("created_at", "") or "").strip(),
                "ficha_tecnica": technical_sheet,
            }
        )
        ports.conjunto_save(
            {
                "codigo": assembly_code,
                "param_codigo": param_edit.text().strip(),
                "descricao": assembly_name,
                "notas": "\n".join(notes_lines),
                "itens": lines,
                "template": bool(save_template_check.isChecked()),
                "origem": "orcamento_conjunto_calculado",
                "margem_perc": float(margin_spin.value() or 0.0),
                "total_custo": subtotal,
                "total_final": final_total,
                "created_at": str(initial.get("created_at", "") or "").strip(),
                "ficha_tecnica": technical_sheet,
            }
        )
    except Exception as exc:
        QMessageBox.critical(owner, "Conjunto", str(exc))
        return None
    return {
        "assembly_code": assembly_code,
        "assembly_name": assembly_name,
        "name": assembly_name,
        "note_cliente": assembly_name,
        "notes_pdf": "\n".join(notes_lines),
        "summary_html": (
            f"{assembly_name} | materiais {_fmt_eur(totals.get('material', 0.0))} | "
            f"mao de obra {_fmt_eur(totals.get('labor', 0.0))} | "
            f"consumiveis {_fmt_eur(totals.get('consumable', 0.0))} | "
            f"produtos {_fmt_eur(totals.get('product', 0.0))} | final {_fmt_eur(final_total)}"
        ),
        "workcenter": ports._quote_pick_workcenter("Serralharia", "Montagem"),
        "lines": lines,
    }
