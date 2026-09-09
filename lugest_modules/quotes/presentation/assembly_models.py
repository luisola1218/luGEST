"""Assembly template editor and catalog manager with explicit UI capabilities."""
from dataclasses import dataclass
from typing import Callable
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget, QTextEdit, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_common import configure_table as _configure_table, fill_table as _fill_table
from lugest_qt.ui.pages.runtime_support import _fmt_eur

@dataclass(frozen=True)
class AssemblyModelEditorPorts:
    wrap_item: Callable
    item_kind: Callable
    kind_label: Callable
    pick_kind: Callable
    edit_item: Callable
    is_laser: Callable
    edit_laser_batch: Callable

@dataclass(frozen=True)
class AssemblyModelManagerPorts:
    rows: Callable
    detail: Callable
    save: Callable
    remove: Callable
    edit: Callable

def edit_assembly_model(owner: QWidget, ports: AssemblyModelEditorPorts, initial: dict | None = None) -> dict | None:
    initial = dict(initial or {})
    dialog = QDialog(owner)
    dialog.setWindowTitle("Modelo de conjunto")
    dialog.resize(920, 640)
    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    code_edit = QLineEdit(str(initial.get("codigo", "") or "").strip())
    desc_edit = QLineEdit(str(initial.get("descricao", "") or "").strip())
    notes_edit = QTextEdit()
    notes_edit.setMinimumHeight(76)
    notes_edit.setPlainText(str(initial.get("notas", "") or "").strip())
    form.addRow("Codigo", code_edit)
    form.addRow("Descricao", desc_edit)
    form.addRow("Notas", notes_edit)
    layout.addLayout(form)

    intro = QLabel(
        "Modelo/Conjunto tecnico separado das linhas DXF/DWG. "
        "Aqui trabalhas sempre com itens de conjunto: material, mao de obra, consumiveis e produtos."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(intro)

    items: list[dict] = [ports.wrap_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
    items_table = QTableWidget(0, 7)
    items_table.setHorizontalHeaderLabels(["Tipo", "Descricao", "Codigo/Ref", "Material", "Esp./Unid", "Qtd", "Preco"])
    items_table.verticalHeader().setVisible(False)
    items_table.setEditTriggers(QTableWidget.NoEditTriggers)
    items_table.setSelectionBehavior(QTableWidget.SelectRows)
    _configure_table(items_table, stretch=(1, 3), contents=(0, 2, 4, 5, 6))

    def render_items() -> None:
        _fill_table(
            items_table,
            [
                [
                    ports.kind_label(ports.item_kind(item)),
                    str(((item.get("line") or {}).get("descricao", "") or "-")).strip() or "-",
                    str(((item.get("line") or {}).get("produto_codigo", "") or (item.get("line") or {}).get("ref_externa", "") or "-")).strip() or "-",
                    str(((item.get("line") or {}).get("material", "") or "-")).strip() or "-",
                    str(((item.get("line") or {}).get("espessura", "") or (item.get("line") or {}).get("produto_unid", "") or "-")).strip() or "-",
                    f"{float(((item.get('line') or {}).get('qtd', 0) or 0)):.2f}",
                    _fmt_eur(float(((item.get("line") or {}).get("preco_unit", 0) or 0))),
                ]
                for item in items
            ],
            align_center_from=4,
        )

    def selected_item_index() -> int:
        current = items_table.currentItem()
        if current is None or current.row() >= len(items):
            return -1
        return current.row()

    actions = QHBoxLayout()
    add_btn = QPushButton("Adicionar item")
    edit_btn = QPushButton("Editar item")
    edit_btn.setProperty("variant", "secondary")
    remove_btn = QPushButton("Remover item")
    remove_btn.setProperty("variant", "danger")
    actions.addWidget(add_btn)
    actions.addWidget(edit_btn)
    actions.addWidget(remove_btn)
    actions.addStretch(1)
    layout.addLayout(actions)
    layout.addWidget(items_table, 1)

    def add_item() -> None:
        kind = ports.pick_kind(dialog)
        if not kind:
            return
        payload = ports.edit_item(kind, parent=dialog)
        if payload is None:
            return
        items.append(payload)
        render_items()

    def edit_item() -> None:
        index = selected_item_index()
        if index < 0:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um item.")
            return
        current = dict(items[index] or {})
        current_line = dict(current.get("line") or current)
        if ports.is_laser(current_line):
            batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
            batch_indexes = [
                row_index
                for row_index, candidate in enumerate(items)
                if batch_id
                and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
            ] or [index]
            source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
            edited_lines = ports.edit_laser_batch(source_lines, parent=dialog)
            if edited_lines is None:
                return
            insert_at = min(batch_indexes)
            for row_index in sorted(batch_indexes, reverse=True):
                del items[row_index]
            for offset, edited_line in enumerate(edited_lines):
                items.insert(insert_at + offset, ports.wrap_item(edited_line))
            render_items()
            if edited_lines:
                items_table.selectRow(insert_at)
            return
        else:
            payload = ports.edit_item(ports.item_kind(current), current, parent=dialog)
        if payload is None:
            return
        items[index] = payload
        render_items()
        items_table.selectRow(index)

    def remove_item() -> None:
        index = selected_item_index()
        if index < 0:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um item.")
            return
        del items[index]
        render_items()

    add_btn.clicked.connect(add_item)
    edit_btn.clicked.connect(edit_item)
    remove_btn.clicked.connect(remove_item)
    render_items()

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return None
    return {
        "codigo": code_edit.text().strip(),
        "descricao": desc_edit.text().strip(),
        "notas": notes_edit.toPlainText().strip(),
        "itens": [dict(item.get("line") or {}) for item in items if isinstance(item.get("line"), dict)],
    }


def manage_assembly_models(owner: QWidget, ports: AssemblyModelManagerPorts) -> None:
    dialog = QDialog(owner)
    dialog.setWindowTitle("Modelos de conjuntos")
    dialog.resize(940, 620)
    layout = QVBoxLayout(dialog)
    table = QTableWidget(0, 6)
    table.setHorizontalHeaderLabels(["Codigo", "Descricao", "Itens", "Pecas", "Produtos", "Total base"])
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QTableWidget.NoEditTriggers)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    _configure_table(table, stretch=(1,), contents=(0, 2, 3, 4, 5))
    layout.addWidget(table, 1)

    def current_code() -> str:
        current = table.currentItem()
        if current is None:
            return ""
        row_item = table.item(current.row(), 0)
        return str(row_item.text() or "").strip() if row_item is not None else ""

    def refresh_models() -> None:
        rows = list(ports.rows() or [])
        _fill_table(
            table,
            [
                [
                    row.get("codigo", "-"),
                    row.get("descricao", "-"),
                    row.get("itens", 0),
                    row.get("pecas", 0),
                    row.get("produtos", 0),
                    _fmt_eur(float(row.get("total_base", 0) or 0)),
                ]
                for row in rows
            ],
            align_center_from=2,
        )
        if table.rowCount() > 0:
            table.selectRow(0)

    actions = QHBoxLayout()
    new_btn = QPushButton("Novo")
    edit_btn = QPushButton("Editar")
    edit_btn.setProperty("variant", "secondary")
    remove_btn = QPushButton("Remover")
    remove_btn.setProperty("variant", "danger")
    close_btn = QPushButton("Fechar")
    close_btn.setProperty("variant", "secondary")
    actions.addWidget(new_btn)
    actions.addWidget(edit_btn)
    actions.addWidget(remove_btn)
    actions.addStretch(1)
    actions.addWidget(close_btn)
    layout.addLayout(actions)

    def create_model() -> None:
        payload = ports.edit()
        if payload is None:
            return
        try:
            ports.save(payload)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        refresh_models()

    def edit_model() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um modelo.")
            return
        try:
            detail = ports.detail(code)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        payload = ports.edit(detail)
        if payload is None:
            return
        payload["codigo"] = code
        try:
            ports.save(payload)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        refresh_models()

    def remove_model() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um modelo.")
            return
        if QMessageBox.question(dialog, "Conjuntos", f"Remover o modelo {code}?") != QMessageBox.Yes:
            return
        try:
            ports.remove(code)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        refresh_models()

    new_btn.clicked.connect(create_model)
    edit_btn.clicked.connect(edit_model)
    remove_btn.clicked.connect(remove_model)
    close_btn.clicked.connect(dialog.reject)
    refresh_models()
    dialog.exec()
