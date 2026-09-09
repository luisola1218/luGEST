"""Saved assembly catalog UI; quote line ownership stays with the caller."""
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from PySide6.QtWidgets import QDialog, QHBoxLayout, QInputDialog, QLabel, QMessageBox, QPushButton, QTableWidget, QVBoxLayout, QWidget
from lugest_qt.ui.pages.runtime_common import configure_table as _configure_table, fill_table as _fill_table
from lugest_qt.ui.pages.runtime_support import _fmt_eur

@dataclass(frozen=True)
class SavedAssemblyPorts:
    rows: Callable
    detail: Callable
    expand: Callable
    remove: Callable
    open_pdf: Callable
    edit: Callable
    append_lines: Callable

def manage_saved_assemblies(owner: QWidget, ports: SavedAssemblyPorts) -> None:
    dialog = QDialog(owner)
    dialog.setWindowTitle("Conjuntos guardados")
    dialog.resize(1080, 640)
    layout = QVBoxLayout(dialog)
    intro = QLabel(
        "Os conjuntos guardados funcionam como produto montado: materiais, mao de obra, consumiveis e produtos. "
        "Daqui podes criar, editar, duplicar, remover e aplicar ao orçamento atual."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(intro)

    table = QTableWidget(0, 9)
    table.setHorizontalHeaderLabels(["Codigo", "Param.", "Descricao", "Itens", "Ligados", "Template", "Margem", "Custo atual", "Final atual"])
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QTableWidget.NoEditTriggers)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    _configure_table(table, stretch=(2,), contents=(0, 1, 3, 4, 5, 6, 7, 8))
    layout.addWidget(table, 1)

    def current_code() -> str:
        current = table.currentItem()
        if current is None:
            return ""
        row_item = table.item(current.row(), 0)
        return str(row_item.text() or "").strip() if row_item is not None else ""

    def refresh_rows(select_code: str = "") -> None:
        rows = list(ports.rows() or [])
        _fill_table(
            table,
            [
                [
                    row.get("codigo", "-"),
                    row.get("param_codigo", "-"),
                    row.get("descricao", "-"),
                    row.get("itens", 0),
                    f"{int(row.get('itens_ligados', 0) or 0)}/{int(row.get('itens', 0) or 0)}",
                    "Sim" if bool(row.get("template", False)) else "Nao",
                    f"{float(row.get('margem_perc', 0) or 0):.2f} %",
                    _fmt_eur(float(row.get("total_custo", 0) or 0)),
                    _fmt_eur(float(row.get("total_final", 0) or 0)),
                ]
                for row in rows
            ],
            align_center_from=3,
        )
        if table.rowCount() <= 0:
            return
        wanted = str(select_code or "").strip()
        row_index = 0
        if wanted:
            for index, row in enumerate(rows):
                if str(row.get("codigo", "") or "").strip() == wanted:
                    row_index = index
                    break
        table.selectRow(row_index)

    actions = QHBoxLayout()
    new_btn = QPushButton("Novo conjunto")
    edit_btn = QPushButton("Editar")
    edit_btn.setProperty("variant", "secondary")
    duplicate_btn = QPushButton("Duplicar")
    duplicate_btn.setProperty("variant", "secondary")
    apply_btn = QPushButton("Adicionar ao orçamento")
    apply_btn.setProperty("variant", "secondary")
    preview_btn = QPushButton("Previsualizar PDF")
    preview_btn.setProperty("variant", "secondary")
    remove_btn = QPushButton("Remover")
    remove_btn.setProperty("variant", "danger")
    close_btn = QPushButton("Fechar")
    close_btn.setProperty("variant", "secondary")
    for button in (new_btn, edit_btn, duplicate_btn, apply_btn, preview_btn, remove_btn):
        actions.addWidget(button)
    actions.addStretch(1)
    actions.addWidget(close_btn)
    layout.addLayout(actions)

    def create_conjunto() -> None:
        payload = ports.edit()
        if not payload:
            return
        refresh_rows(str(payload.get("assembly_code", "") or "").strip())

    def edit_conjunto() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
            return
        try:
            detail = dict(ports.detail(code) or {})
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        payload = ports.edit(detail)
        if not payload:
            return
        refresh_rows(str(payload.get("assembly_code", code) or code).strip())

    def duplicate_conjunto() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
            return
        try:
            detail = dict(ports.detail(code) or {})
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        detail["codigo"] = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        detail.pop("param_codigo", None)
        detail["descricao"] = f"{str(detail.get('descricao', code) or code).strip()} (Copia)"
        detail.pop("created_at", None)
        detail.pop("updated_at", None)
        payload = ports.edit(detail)
        if not payload:
            return
        refresh_rows(str(payload.get("assembly_code", "") or "").strip())

    def apply_conjunto() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
            return
        qty, ok = QInputDialog.getDouble(dialog, "Adicionar conjunto", "Quantidade de conjuntos", 1.0, 0.01, 1000000.0, 2)
        if not ok:
            return
        try:
            ports.append_lines(ports.expand(code, qty))
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        QMessageBox.information(dialog, "Conjuntos", f"O conjunto {code} foi adicionado ao orçamento.")

    def preview_conjunto() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
            return
        try:
            path = ports.open_pdf(code)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        QMessageBox.information(dialog, "Conjuntos", f"Ficha PDF aberta:\n{path}")

    def remove_conjunto() -> None:
        code = current_code()
        if not code:
            QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
            return
        if QMessageBox.question(dialog, "Conjuntos", f"Remover o conjunto {code}?") != QMessageBox.Yes:
            return
        try:
            ports.remove(code)
        except Exception as exc:
            QMessageBox.critical(dialog, "Conjuntos", str(exc))
            return
        refresh_rows()

    new_btn.clicked.connect(create_conjunto)
    edit_btn.clicked.connect(edit_conjunto)
    duplicate_btn.clicked.connect(duplicate_conjunto)
    apply_btn.clicked.connect(apply_conjunto)
    preview_btn.clicked.connect(preview_conjunto)
    remove_btn.clicked.connect(remove_conjunto)
    close_btn.clicked.connect(dialog.reject)
    refresh_rows()
    dialog.exec()
