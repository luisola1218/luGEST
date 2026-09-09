"""Save a selection as an assembly using draft values and explicit catalog ports."""
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget, QTextEdit, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_common import configure_table as _configure_table, fill_table as _fill_table, table_visible_height as _table_visible_height
from lugest_qt.ui.pages.runtime_support import _fmt_eur

@dataclass(frozen=True)
class GroupEditorPorts:
    assembly_model_detail: Callable
    assembly_model_rows: Callable
    assembly_model_save: Callable
    conjunto_detail: Callable
    conjunto_rows: Callable
    conjunto_save: Callable
    norm_text: Callable
    orc_line_is_piece: Callable
    orc_line_is_product: Callable
    save_assembly_pair: Callable
    line_type_label: Callable

@dataclass(frozen=True)
class GroupResult:
    lines: list[dict]
    code: str
    description: str
    overwrite: bool

def save_group(owner: QWidget, ports: GroupEditorPorts, quote_lines: list[dict],
               selected_indexes: list[int], quote_number: str = "", note: str = "") -> GroupResult | None:
    lines = deepcopy(quote_lines)
    if not selected_indexes:
        QMessageBox.warning(owner, "Conjuntos", "Seleciona pelo menos uma linha para guardar no conjunto/modelo.")
        return

    selected_rows = [dict(lines[index] or {}) for index in selected_indexes]
    selected_codes = {
        str(row.get("conjunto_codigo", "") or "").strip()
        for row in selected_rows
        if str(row.get("conjunto_codigo", "") or "").strip()
    }
    selected_names = {
        str(row.get("conjunto_nome", "") or "").strip()
        for row in selected_rows
        if str(row.get("conjunto_nome", "") or "").strip()
    }
    prefill_code = next(iter(selected_codes), "") if len(selected_codes) == 1 else ""
    prefill_name = next(iter(selected_names), "") if len(selected_names) == 1 else ""
    if not prefill_name:
        prefill_name = str(note or quote_number or "").strip()
    if not prefill_code:
        prefill_code = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    dialog = QDialog(owner)
    dialog.setWindowTitle("Guardar linhas como conjunto/modelo")
    dialog.resize(760, 520)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(10)

    intro = QLabel(
        "Seleciona como queres guardar as linhas atuais: num conjunto montado, num modelo reutilizável, "
        "ou nos dois ao mesmo tempo. Se escolheres sobrepor, o sistema junta as linhas selecionadas às linhas "
        "desse conjunto que já estejam no orçamento atual, para não perder o que já montaste."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(intro)

    form = QFormLayout()
    form.setHorizontalSpacing(10)
    form.setVerticalSpacing(8)

    destination_combo = QComboBox()
    destination_combo.addItem("Conjunto", "conjunto")
    destination_combo.addItem("Modelo", "modelo")
    destination_combo.addItem("Conjunto + Modelo", "both")

    save_mode_combo = QComboBox()
    save_mode_combo.addItem("Novo registo", "new")
    save_mode_combo.addItem("Sobrepor existente", "overwrite")

    existing_combo = QComboBox()
    existing_combo.setEnabled(False)

    code_edit = QLineEdit(prefill_code)
    name_edit = QLineEdit(prefill_name)
    notes_edit = QTextEdit()
    notes_edit.setMaximumHeight(100)
    notes_edit.setPlainText(
        f"Guardado a partir do orçamento {str(quote_number or '').strip()} com {len(selected_rows)} linha(s) selecionada(s)."
    )
    template_check = QCheckBox("Marcar também como template reutilizável")

    info_label = QLabel("")
    info_label.setWordWrap(True)
    info_label.setProperty("role", "muted")

    form.addRow("Guardar em", destination_combo)
    form.addRow("Modo", save_mode_combo)
    form.addRow("Registo existente", existing_combo)
    form.addRow("Código", code_edit)
    form.addRow("Descrição", name_edit)
    form.addRow("Notas", notes_edit)
    form.addRow("", template_check)
    layout.addLayout(form)
    layout.addWidget(info_label)

    preview_table = QTableWidget(0, 5)
    preview_table.setHorizontalHeaderLabels(["Tipo", "Ref. Ext.", "Descrição", "Qtd", "Total"])
    preview_table.verticalHeader().setVisible(False)
    preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
    preview_table.setSelectionMode(QTableWidget.NoSelection)
    preview_table.setAlternatingRowColors(True)
    _configure_table(preview_table, stretch=(2,), contents=(0, 1, 3, 4))
    _fill_table(
        preview_table,
        [
            [
                ports.line_type_label(row),
                str(row.get("ref_externa", "") or "-").strip() or "-",
                str(row.get("descricao", "") or "-").strip() or "-",
                f"{float(row.get('qtd', 0) or 0):.2f}",
                _fmt_eur(float(row.get("total", float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0)) or 0)),
            ]
            for row in selected_rows
        ],
        align_center_from=3,
    )
    preview_table.setMinimumHeight(max(170, min(290, _table_visible_height(preview_table, min(max(len(selected_rows), 3), 6), extra=14))))
    layout.addWidget(preview_table, 1)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    ok_btn = buttons.button(QDialogButtonBox.Ok)

    def _list_destination_rows(destination: str) -> list[dict]:
        dest = str(destination or "").strip().lower()
        if dest == "conjunto":
            return [dict(row or {}) for row in list(ports.conjunto_rows() or [])]
        if dest == "modelo":
            return [dict(row or {}) for row in list(ports.assembly_model_rows() or [])]
        combined: dict[str, dict] = {}
        for source_row in list(ports.conjunto_rows() or []):
            row = dict(source_row or {})
            code = str(row.get("codigo", "") or "").strip()
            if code:
                combined[code] = {
                    "codigo": code,
                    "descricao": str(row.get("descricao", "") or "").strip(),
                    "source": "both",
                }
        for source_row in list(ports.assembly_model_rows() or []):
            row = dict(source_row or {})
            code = str(row.get("codigo", "") or "").strip()
            if not code:
                continue
            combined.setdefault(
                code,
                {
                    "codigo": code,
                    "descricao": str(row.get("descricao", "") or "").strip(),
                    "source": "both",
                },
            )
        return [combined[key] for key in sorted(combined)]

    def _existing_codes(destination: str) -> set[str]:
        return {
            str(row.get("codigo", "") or "").strip()
            for row in _list_destination_rows(destination)
            if str(row.get("codigo", "") or "").strip()
        }

    def _target_detail_lines(destination: str, code: str) -> list[dict]:
        target_code = str(code or "").strip()
        if not target_code:
            return []
        if destination in {"conjunto", "both"}:
            try:
                detail = dict(ports.conjunto_detail(target_code) or {})
            except Exception:
                detail = {}
            rows = [dict(item or {}) for item in list(detail.get("itens", []) or []) if isinstance(item, dict)]
            if rows:
                return rows
        if destination in {"modelo", "both"}:
            try:
                detail = dict(ports.assembly_model_detail(target_code) or {})
            except Exception:
                detail = {}
            return [dict(item or {}) for item in list(detail.get("itens", []) or []) if isinstance(item, dict)]
        return []

    def _line_signature(row: dict) -> tuple:
        return (
            str(row.get("tipo_item", "") or "").strip(),
            str(row.get("produto_codigo", "") or "").strip(),
            str(row.get("ref_externa", "") or "").strip(),
            str(row.get("descricao", "") or "").strip(),
            str(row.get("material", "") or "").strip(),
            str(row.get("espessura", "") or "").strip(),
            str(row.get("operacao", "") or "").strip(),
            round(float(row.get("qtd", 0) or 0), 4),
            round(float(row.get("preco_unit", 0) or 0), 4),
            str(row.get("desenho", "") or "").strip(),
        )

    def _sanitize_group_line(raw_row: dict) -> dict:
        row = dict(raw_row or {})
        row.pop("conjunto_codigo", None)
        row.pop("conjunto_nome", None)
        row.pop("grupo_uuid", None)
        row.pop("total", None)
        if ports.orc_line_is_product(row) and not str(row.get("ref_externa", "") or "").strip():
            row["ref_externa"] = str(row.get("produto_codigo", "") or "").strip()
        operation_norm = ports.norm_text(str(row.get("operacao", "") or ""))
        if ports.orc_line_is_piece(row) and "laser" in operation_norm:
            row["source_quote_number"] = str(quote_number or "").strip()
            row["source_ref_externa"] = str(row.get("ref_externa", "") or "").strip()
            row["pricing_source"] = "quote_laser"
        return row

    def _compose_rows_for_save(destination: str, target_code: str, overwrite: bool) -> list[dict]:
        rows = []
        if overwrite:
            quote_rows = [
                dict(row or {})
                for row in list(lines)
                if str((row or {}).get("conjunto_codigo", "") or "").strip() == target_code
            ]
            if quote_rows:
                rows.extend(quote_rows)
            else:
                rows.extend(_target_detail_lines(destination, target_code))
        rows.extend(selected_rows)
        merged: list[dict] = []
        seen: set[tuple] = set()
        for raw_row in rows:
            clean_row = _sanitize_group_line(raw_row)
            signature = _line_signature(clean_row)
            if signature in seen:
                continue
            seen.add(signature)
            merged.append(clean_row)
        return merged

    def _refresh_overwrite_options() -> None:
        destination = str(destination_combo.currentData() or "conjunto").strip()
        overwrite = str(save_mode_combo.currentData() or "new").strip() == "overwrite"
        rows = _list_destination_rows(destination)
        existing_combo.blockSignals(True)
        existing_combo.clear()
        for row in rows:
            code = str(row.get("codigo", "") or "").strip()
            desc = str(row.get("descricao", "") or "").strip()
            existing_combo.addItem(f"{code} - {desc}", code)
        existing_combo.blockSignals(False)
        existing_combo.setVisible(overwrite)
        existing_combo.setEnabled(overwrite and bool(rows))
        code_edit.setReadOnly(overwrite)
        if overwrite and rows:
            wanted_code = str(prefill_code or "").strip()
            wanted_index = 0
            if wanted_code:
                for idx, row in enumerate(rows):
                    if str(row.get("codigo", "") or "").strip() == wanted_code:
                        wanted_index = idx
                        break
            existing_combo.setCurrentIndex(wanted_index)
            code_edit.setText(str(existing_combo.currentData() or "").strip())
            if not name_edit.text().strip():
                row = rows[wanted_index]
                name_edit.setText(str(row.get("descricao", "") or "").strip())
        elif overwrite:
            code_edit.clear()
        elif not code_edit.text().strip():
            code_edit.setText(prefill_code)
        if overwrite and not rows:
            info_label.setText("Nao existem registos desse tipo para sobrepor. Escolhe 'Novo registo' ou cria primeiro um conjunto/modelo.")
            if isinstance(ok_btn, QPushButton):
                ok_btn.setEnabled(False)
            return
        info_label.setText(
            "As linhas selecionadas vao ficar ligadas ao conjunto/modelo escolhido no orçamento atual."
            if overwrite
            else "Vai ser criado um novo conjunto/modelo a partir das linhas selecionadas."
        )
        if isinstance(ok_btn, QPushButton):
            ok_btn.setEnabled(True)

    def _sync_selected_existing() -> None:
        code = str(existing_combo.currentData() or "").strip()
        if not code:
            return
        destination = str(destination_combo.currentData() or "conjunto").strip()
        rows = _list_destination_rows(destination)
        detail_row = next((row for row in rows if str(row.get("codigo", "") or "").strip() == code), {})
        code_edit.setText(code)
        if detail_row:
            name_edit.setText(str(detail_row.get("descricao", "") or "").strip())

    destination_combo.currentIndexChanged.connect(_refresh_overwrite_options)
    save_mode_combo.currentIndexChanged.connect(_refresh_overwrite_options)
    existing_combo.currentIndexChanged.connect(_sync_selected_existing)
    _refresh_overwrite_options()

    if dialog.exec() != QDialog.Accepted:
        return

    destination = str(destination_combo.currentData() or "conjunto").strip()
    overwrite = str(save_mode_combo.currentData() or "new").strip() == "overwrite"
    code = (
        str(existing_combo.currentData() or "").strip()
        if overwrite
        else str(code_edit.text() or "").strip()
    )
    if not code:
        QMessageBox.warning(owner, "Conjuntos", "Indica um código para o conjunto/modelo.")
        return
    description = str(name_edit.text() or "").strip() or code
    notes = str(notes_edit.toPlainText() or "").strip()
    if not overwrite and code in _existing_codes(destination):
        QMessageBox.warning(
            owner,
            "Conjuntos",
            f"Ja existe um registo com o codigo {code}. Se queres atualizar esse registo, usa o modo 'Sobrepor existente'.",
        )
        return

    lines_for_save = _compose_rows_for_save(destination, code, overwrite)
    if not lines_for_save:
        QMessageBox.warning(owner, "Conjuntos", "Nao ha linhas validas para guardar no conjunto/modelo.")
        return

    total_value = round(
        sum(
            float(row.get("total", float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0)) or 0)
            for row in lines_for_save
        ),
        2,
    )
    payload = {
        "codigo": code,
        "descricao": description,
        "notas": notes,
        "itens": lines_for_save,
        "template": bool(template_check.isChecked()),
        "origem": "orcamento_linhas_guardadas",
        "margem_perc": 0.0,
        "total_custo": total_value,
        "total_final": total_value,
    }

    try:
        if destination == "both":
            ports.save_assembly_pair(payload, payload)
        elif destination == "conjunto":
            ports.conjunto_save(payload)
        else:
            ports.assembly_model_save(payload)
    except Exception as exc:
        QMessageBox.critical(owner, "Conjuntos", str(exc))
        return

    existing_group_rows = [
        dict(row or {})
        for row in list(lines)
        if str((row or {}).get("conjunto_codigo", "") or "").strip() == code
        and str((row or {}).get("grupo_uuid", "") or "").strip()
    ]
    group_uuid = (
        str(existing_group_rows[0].get("grupo_uuid", "") or "").strip()
        if existing_group_rows
        else f"{code}-01"
    )
    affected_indexes = set(selected_indexes)
    if overwrite:
        affected_indexes.update(
            index
            for index, row in enumerate(lines)
            if str((row or {}).get("conjunto_codigo", "") or "").strip() == code
        )
    for index in sorted(affected_indexes):
        if not (0 <= index < len(lines)):
            continue
        row = dict(lines[index] or {})
        row["conjunto_codigo"] = code
        row["conjunto_nome"] = description
        row["grupo_uuid"] = group_uuid
        lines[index] = row

    return GroupResult(lines, code, description, overwrite)
