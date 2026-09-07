from __future__ import annotations
import unicodedata
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from .runtime_common import (
    configure_table as _configure_table,
    selected_row_index as _selected_row_index,
    set_table_columns as _set_table_columns,
    smart_sort_key as _smart_sort_key,
    table_visible_height as _table_visible_height,
)
from ..widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


LIST_TABLE_FONT_PX = 15


LIST_TABLE_ROW_PX = 42


def _is_dark(hex_color: str) -> bool:
    color = QColor(str(hex_color or "#ffffff"))
    return (color.red() * 0.299 + color.green() * 0.587 + color.blue() * 0.114) < 160


def _reference_catalog_dialog(
    parent: QWidget,
    references: list[dict],
    title: str = "Histórico de referencias",
    *,
    backend: object | None = None,
    current_client: str = "",
) -> dict | None:
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.resize(1280, 680)
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(10)

    header = QHBoxLayout()
    title_lbl = QLabel(title)
    title_lbl.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
    client_filter = QComboBox()
    client_filter.addItem("Todos os clientes", "")
    client_codes = sorted(
        {
            str(row.get("cliente_codigo", "") or "").strip().upper()
            for row in list(references or [])
            if str(row.get("cliente_codigo", "") or "").strip()
        }
    )
    for code in client_codes:
        client_filter.addItem(code, code)
    current_client_code = str(current_client or "").strip().upper()
    if current_client_code:
        for index in range(client_filter.count()):
            if str(client_filter.itemData(index) or "").strip().upper() == current_client_code:
                client_filter.setCurrentIndex(index)
                break
    search_edit = QLineEdit()
    search_edit.setPlaceholderText("Pesquisar por cliente, ref. interna, externa, descricao, material, qualidade ou espessura")
    header.addWidget(title_lbl)
    header.addStretch(1)
    header.addWidget(client_filter)
    header.addWidget(search_edit, 1)
    layout.addLayout(header)

    card = CardFrame()
    card.set_tone("default")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(12, 12, 12, 12)
    card_layout.setSpacing(8)
    table = QTableWidget(0, 10)
    table.setHorizontalHeaderLabels(["Cliente", "Ref. interna", "Ref. externa", "Descricao", "Material", "Qualidade MP", "Esp.", "Tempo", "Preco", "Origem"])
    table.verticalHeader().setVisible(False)
    table.setSelectionBehavior(QTableWidget.SelectRows)
    table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked | QTableWidget.EditKeyPressed)
    _configure_table(table, stretch=(3,), contents=(0, 1, 2, 4, 5, 6, 7, 8, 9))
    card_layout.addWidget(table)
    layout.addWidget(card, 1)

    result: dict | None = None
    visible_rows: list[dict] = []

    def _editable_item(value: object, *, editable: bool = True, center: bool = False) -> QTableWidgetItem:
        item = QTableWidgetItem(str(value or "").strip())
        item.setTextAlignment(int((Qt.AlignCenter if center else Qt.AlignLeft) | Qt.AlignVCenter))
        if not editable:
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        return item

    def selected_payload_from_table() -> dict | None:
        row_index = _selected_row_index(table)
        if row_index < 0:
            return None
        item = table.item(row_index, 0)
        base = dict(item.data(Qt.UserRole) or {}) if item is not None else {}
        if not base:
            return None
        base.update(
            {
                "cliente_codigo": table.item(row_index, 0).text().strip() if table.item(row_index, 0) else "",
                "ref_interna": table.item(row_index, 1).text().strip() if table.item(row_index, 1) else "",
                "ref_externa": table.item(row_index, 2).text().strip() if table.item(row_index, 2) else "",
                "descricao": table.item(row_index, 3).text().strip() if table.item(row_index, 3) else "",
                "material": table.item(row_index, 4).text().strip() if table.item(row_index, 4) else "",
                "material_subtype": table.item(row_index, 5).text().strip() if table.item(row_index, 5) else "",
                "espessura": table.item(row_index, 6).text().strip() if table.item(row_index, 6) else "",
                "tempo_peca_min": table.item(row_index, 7).text().strip().replace(",", ".") if table.item(row_index, 7) else "0",
                "preco_unit": table.item(row_index, 8).text().strip().replace(",", ".") if table.item(row_index, 8) else "0",
                "origem_doc": table.item(row_index, 9).text().strip() if table.item(row_index, 9) else "",
            }
        )
        return base

    def render() -> None:
        nonlocal visible_rows
        query = search_edit.text().strip().lower()
        selected_client = str(client_filter.currentData() or "").strip().upper()
        filtered = []
        for row in references:
            row_client = str(row.get("cliente_codigo", "") or "").strip().upper()
            if selected_client and row_client != selected_client:
                continue
            hay = " | ".join(
                [
                    row_client,
                    str(row.get("ref_interna", "") or ""),
                    str(row.get("ref_externa", "") or ""),
                    str(row.get("descricao", "") or ""),
                    str(row.get("material", "") or ""),
                    str(row.get("material_subtype", "") or ""),
                    str(row.get("espessura", "") or ""),
                ]
            ).lower()
            if query and query not in hay:
                continue
            filtered.append(row)
        filtered.sort(
            key=lambda row: (
                str(row.get("cliente_codigo", "") or ""),
                str(row.get("ref_interna", "") or ""),
                str(row.get("ref_externa", "") or ""),
            )
        )
        visible_rows = [dict(row or {}) for row in filtered]
        table.setRowCount(len(filtered))
        for row_index, row in enumerate(filtered):
            values = [
                str(row.get("cliente_codigo", "") or "").strip(),
                str(row.get("ref_interna", "") or "").strip(),
                str(row.get("ref_externa", "") or "").strip(),
                str(row.get("descricao", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("material_subtype", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                f"{float(row.get('tempo_peca_min', row.get('tempo_pecas_min', 0)) or 0):.2f}",
                f"{float(row.get('preco_unit', row.get('preco', 0)) or 0):.4f}",
                str(row.get("origem_doc", "") or row.get("origem_tipo", "") or "").strip(),
            ]
            for col_index, value in enumerate(values):
                editable = col_index in (3, 4, 5, 6, 7, 8)
                item = _editable_item(value, editable=editable, center=col_index in (0, 6, 7, 8))
                if col_index == 0:
                    item.setData(Qt.UserRole, dict(row))
                table.setItem(row_index, col_index, item)
        if filtered:
            table.selectRow(0)

    def refresh_from_backend() -> None:
        nonlocal references
        if backend is not None and hasattr(backend, "order_reference_rows"):
            try:
                references = list(backend.order_reference_rows("", "") or [])  # type: ignore[attr-defined]
            except Exception:
                references = list(references or [])
        render()

    def save_selected() -> None:
        if backend is None or not hasattr(backend, "orc_reference_update"):
            QMessageBox.information(dialog, "Historico de referencias", "Este ambiente nao permite guardar alteracoes ao catalogo.")
            return
        payload = selected_payload_from_table()
        if not payload:
            return
        ref_ext = str(payload.get("ref_externa", "") or "").strip()
        if not ref_ext:
            QMessageBox.warning(dialog, "Historico de referencias", "A referencia externa e obrigatoria para guardar.")
            return
        try:
            updated = backend.orc_reference_update(ref_ext, payload)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.critical(dialog, "Historico de referencias", str(exc))
            return
        QMessageBox.information(dialog, "Historico de referencias", "Referencia atualizada no catalogo.")
        refresh_from_backend()
        for row_index, row in enumerate(visible_rows):
            if str(row.get("ref_externa", "") or "").strip() == str(updated.get("ref_externa", "") or "").strip():
                table.selectRow(row_index)
                break

    def remove_selected() -> None:
        if backend is None or not hasattr(backend, "orc_reference_remove"):
            QMessageBox.information(dialog, "Historico de referencias", "Este ambiente nao permite remover referencias do catalogo.")
            return
        payload = selected_payload_from_table()
        if not payload:
            return
        ref_ext = str(payload.get("ref_externa", "") or "").strip()
        if not ref_ext:
            return
        if QMessageBox.question(
            dialog,
            "Remover referencia",
            f"Remover '{ref_ext}' do catalogo de referencias?\n\nOrcamentos e encomendas existentes mantem os seus dados.",
        ) != QMessageBox.Yes:
            return
        try:
            backend.orc_reference_remove(ref_ext)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.critical(dialog, "Historico de referencias", str(exc))
            return
        refresh_from_backend()

    def accept_selected() -> None:
        nonlocal result
        result = selected_payload_from_table()
        if result:
            dialog.accept()

    actions = QHBoxLayout()
    save_btn = QPushButton("Guardar alteracoes")
    save_btn.setProperty("variant", "secondary")
    save_btn.clicked.connect(save_selected)
    remove_btn = QPushButton("Remover do catalogo")
    remove_btn.setProperty("variant", "danger")
    remove_btn.clicked.connect(remove_selected)
    actions.addWidget(save_btn)
    actions.addWidget(remove_btn)
    actions.addStretch(1)
    layout.addLayout(actions)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(accept_selected)
    buttons.rejected.connect(dialog.reject)
    buttons.button(QDialogButtonBox.Ok).setText("Carregar referencia")
    layout.addWidget(buttons)

    search_edit.textChanged.connect(render)
    client_filter.currentIndexChanged.connect(lambda _index: render())
    table.itemDoubleClicked.connect(lambda *_args: accept_selected())
    render()
    if dialog.exec() != QDialog.Accepted:
        return None
    return result


def _split_client_label(text: str) -> tuple[str, str]:
    raw = str(text or "").strip()
    if not raw:
        return "", ""
    if " - " in raw:
        left, right = raw.split(" - ", 1)
        return left.strip(), right.strip()
    parts = raw.split(None, 1)
    if len(parts) == 2 and parts[0].upper().startswith("CL"):
        return parts[0].strip(), parts[1].strip()
    return raw, ""


def _format_client_label(text: str, *, show_name: bool = True) -> str:
    code, name = _split_client_label(text)
    if show_name and name:
        return f"{code} - {name}".strip(" -")
    return code or name or "-"


def _piece_ops_progress(piece: dict, current_operation: str = "") -> dict[str, float | int]:
    ops = list(piece.get("ops", []) or [])
    if not ops:
        pending = list(piece.get("pendentes", []) or [])
        total = len(pending)
        return {"total": total, "done": 0, "running": 0, "pending": total, "progress_pct": 0.0}
    current_tokens = {str(token).strip().lower() for token in str(current_operation or "").split("+") if str(token).strip()}
    total = len(ops)
    done = 0
    running = 0
    pending = 0
    weighted_done = 0.0
    planned_qty = float(piece.get("planeado", piece.get("quantidade_pedida", 0)) or 0)
    previous_output = planned_qty
    for index, op in enumerate(ops):
        name = str(op.get("nome", "") or "").strip().lower()
        state = str(op.get("estado", "") or "").strip().lower()
        done_qty = float(op.get("qtd_ok", 0) or 0) + float(op.get("qtd_nok", 0) or 0) + float(op.get("qtd_qual", 0) or 0)
        capacity = max(0.0, planned_qty if index == 0 else previous_output)
        if done_qty <= 0 and "concl" in state and capacity > 0:
            done_qty = capacity
        completion = 1.0 if capacity <= 0 else max(0.0, min(1.0, done_qty / capacity))
        previous_output = done_qty if done_qty > 0 else previous_output
        weighted_done += completion
        if completion >= 1.0 or "concl" in state:
            done += 1
        elif completion > 0 or "curso" in state or "produc" in state or (name and name in current_tokens):
            running += 1
        else:
            pending += 1
    progress = 0.0 if total <= 0 else round((weighted_done / total) * 100.0, 1)
    if done >= total and total > 0:
        progress = 100.0
    return {"total": total, "done": done, "running": running, "pending": pending, "progress_pct": progress}


def _normalize_operation_text(raw: str) -> str:
    txt = unicodedata.normalize("NFKD", str(raw or "").strip().lower())
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    if "laser" in txt:
        return "Corte Laser"
    if "quin" in txt:
        return "Quinagem"
    if "rosc" in txt:
        return "Roscagem"
    if "embal" in txt:
        return "Embalamento"
    if "mont" in txt:
        return "Montagem"
    if "sold" in txt:
        return "Soldadura"
    if "furo" in txt:
        return "Furo Manual"
    return str(raw or "").strip()


def _operations_for_posto(posto: str, operations: list[str]) -> list[str]:
    posto_norm = unicodedata.normalize("NFKD", str(posto or "").strip().lower())
    posto_norm = "".join(ch for ch in posto_norm if not unicodedata.combining(ch))
    normalized = [_normalize_operation_text(op) for op in list(operations or []) if str(op or "").strip()]
    if not posto_norm or posto_norm == "geral":
        seen: set[str] = set()
        out: list[str] = []
        for op in normalized:
            key = op.lower()
            if key not in seen:
                seen.add(key)
                out.append(op)
        return out
    keyword_map = {
        "laser": "laser",
        "quinagem": "quin",
        "quinadora": "quin",
        "quinadeira": "quin",
        "quinadeiras": "quin",
        "quinad": "quin",
        "roscagem": "rosc",
        "roscadora": "rosc",
        "embalamento": "embal",
        "embalar": "embal",
        "montagem": "mont",
        "soldadura": "sold",
        "serralharia": "serral",
        "furo manual": "furo",
    }
    keyword = next((token for key, token in keyword_map.items() if key in posto_norm), "")
    if not keyword:
        return normalized
    filtered = [op for op in normalized if keyword in op.lower()]
    return filtered


def _fmt_eur(value: float) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


def _make_inline_progress(value: float) -> QProgressBar:
    pct = int(round(float(value or 0)))
    bar = QProgressBar()
    bar.setRange(0, 100)
    bar.setValue(max(0, min(100, pct)))
    bar.setFormat(f"{float(value or 0):.1f}%")
    bar.setTextVisible(True)
    bar.setMaximumHeight(14)
    bar.setStyleSheet(
        "QProgressBar {"
        " background: #edf2f7;"
        " border: 1px solid #c6d2e0;"
        " border-radius: 7px;"
        " color: #10253d;"
        " font-size: 10px;"
        " text-align: center;"
        "}"
        "QProgressBar::chunk {"
        " background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #f59e0b, stop:1 #ea580c);"
        " border-radius: 6px;"
        "}"
    )
    return bar


def _apply_progress_style(bar: QProgressBar, *, compact: bool = False) -> None:
    height = 14 if compact else 18
    radius = 6 if compact else 7
    bar.setStyleSheet(
        "QProgressBar {"
        " background: #edf2f7;"
        " border: 1px solid #c6d2e0;"
        f" border-radius: {radius}px;"
        " color: #10253d;"
        f" font-size: {'10px' if compact else '11px'};"
        " text-align: center;"
        " font-weight: 700;"
        f" min-height: {height}px;"
        "}"
        "QProgressBar::chunk {"
        " background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #f59e0b, stop:1 #ea580c);"
        f" border-radius: {max(4, radius - 1)}px;"
        "}"
    )


def _operation_tokens(raw: str) -> list[str]:
    return [part.strip() for part in str(raw or "").split("+") if part.strip()]


def _build_operation_selector(
    values: list[str], initial_text: str = "", on_change: callable | None = None
) -> tuple[QWidget, QLineEdit, callable]:
    host = QWidget()
    host_layout = QVBoxLayout(host)
    host_layout.setContentsMargins(0, 0, 0, 0)
    host_layout.setSpacing(8)
    summary = QLineEdit()
    summary.setReadOnly(True)
    summary.setPlaceholderText("Seleciona os postos de trabalho")
    host_layout.addWidget(summary)
    grid = QGridLayout()
    grid.setContentsMargins(0, 0, 0, 0)
    grid.setHorizontalSpacing(10)
    grid.setVerticalSpacing(6)
    host_layout.addLayout(grid)
    checks: list[tuple[str, QCheckBox]] = []
    syncing = {"value": False}

    def sync_summary(*, user_initiated: bool = False) -> None:
        selected = [label for label, checkbox in checks if checkbox.isChecked()]
        summary.setText(" + ".join(selected))
        if callable(on_change) and not bool(syncing["value"]):
            on_change(summary.text().strip(), user_initiated)

    for index, value in enumerate([str(item or "").strip() for item in values if str(item or "").strip()]):
        checkbox = QCheckBox(value)
        checkbox.toggled.connect(lambda _checked, _sync=sync_summary: _sync(user_initiated=True))
        checks.append((value, checkbox))
        grid.addWidget(checkbox, index // 3, index % 3)

    def apply_text(text: str) -> None:
        selected = {token.lower() for token in _operation_tokens(text)}
        changed = False
        syncing["value"] = True
        for label, checkbox in checks:
            state = label.lower() in selected if selected else False
            if checkbox.isChecked() != state:
                checkbox.setChecked(state)
                changed = True
        if not selected and checks and not any(checkbox.isChecked() for _label, checkbox in checks):
            checks[0][1].setChecked(True)
            changed = True
        syncing["value"] = False
        if not changed:
            sync_summary(user_initiated=False)
        else:
            sync_summary(user_initiated=False)

    apply_text(initial_text)
    return host, summary, apply_text


_OPERATION_PRICING_MODE_ITEMS: list[tuple[str, str]] = [
    ("manual", "Manual"),
    ("per_piece", "Por peca"),
    ("per_feature", "Por quantidade"),
    ("per_area_m2", "Por m2"),
]


def _operation_pricing_mode_label(value: str) -> str:
    raw = str(value or "").strip().lower()
    for key, label in _OPERATION_PRICING_MODE_ITEMS:
        if raw == key:
            return label
    return "Manual"


def _open_operation_cost_profiles_dialog(parent: QWidget, backend) -> bool:
    settings = dict(backend.operation_cost_settings() or {})
    active_profile = str(settings.get("active_profile", "Base") or "Base").strip() or "Base"
    profiles = dict(settings.get("profiles", {}) or {})
    active_map = dict(profiles.get(active_profile, {}) or {})
    operations = [str(op or "").strip() for op in list(backend.desktop_main.OFF_OPERACOES_DISPONIVEIS) if str(op or "").strip()]

    dialog = QDialog(parent)
    dialog.setWindowTitle("Perfis de custo por operação")
    dialog.resize(1460, 600)
    dialog.setStyleSheet(
        "QLabel { font-size: 10px; }"
        "QLineEdit, QComboBox, QDoubleSpinBox { font-size: 10px; min-height: 24px; padding: 0 6px; }"
        "QTableWidget { font-size: 10px; gridline-color: #d3dde9; alternate-background-color: #f8fbff; }"
        "QTableWidget::item { padding: 0 8px; }"
        "QHeaderView::section { font-size: 10px; padding: 6px 8px; }"
    )
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)
    info = QLabel(
        "Define a lógica base de custo de cada posto. Estes valores alimentam o orçamento e podem ser ajustados "
        "individualmente em cada linha."
    )
    info.setWordWrap(True)
    info.setProperty("role", "muted")
    info.setStyleSheet("font-size: 10px;")
    layout.addWidget(info)

    profile_row = QHBoxLayout()
    profile_row.setContentsMargins(2, 0, 2, 0)
    profile_row.setSpacing(8)
    profile_row.addWidget(QLabel("Perfil ativo"))
    profile_name_edit = QLineEdit(active_profile)
    profile_name_edit.setPlaceholderText("Base")
    profile_name_edit.setMaximumHeight(28)
    profile_row.addWidget(profile_name_edit, 1)
    layout.addLayout(profile_row)

    sort_hint = QLabel("Clique num cabeçalho para ordenar. Um segundo clique inverte a ordem.")
    sort_hint.setProperty("role", "muted")
    sort_hint.setStyleSheet("font-size: 9.5px; color: #52657d;")
    layout.addWidget(sort_hint)

    table = QTableWidget(0, 9)
    table.setHorizontalHeaderLabels(
        ["Operação", "Modo de cálculo", "Unidade técnica", "Qtd. padrão", "Setup (min)", "Tempo base", "EUR/h", "Fixo/un", "Mín./un"]
    )
    table.verticalHeader().setVisible(False)
    table.setAlternatingRowColors(True)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setSelectionMode(QAbstractItemView.SingleSelection)
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignCenter)
    _configure_table(table, stretch=(), contents=())
    table.setShowGrid(True)
    header.setSectionResizeMode(0, QHeaderView.Interactive)
    header.setSectionResizeMode(1, QHeaderView.Interactive)
    header.setSectionResizeMode(2, QHeaderView.Interactive)
    for col_index in range(3, 9):
        header.setSectionResizeMode(col_index, QHeaderView.Interactive)
    table.setColumnWidth(0, 190)
    table.setColumnWidth(1, 150)
    header.setSectionResizeMode(2, QHeaderView.Stretch)
    for col_index in range(3, 9):
        table.setColumnWidth(col_index, 118)
    table.setCornerButtonEnabled(False)
    header_items = (
        "Posto ou operação produtiva",
        "Forma de cálculo aplicada",
        "Grandeza usada no cálculo por peça",
        "Quantidade técnica assumida por defeito",
        "Tempo de preparação do posto",
        "Tempo produtivo base por unidade técnica",
        "Custo horário do posto",
        "Custo fixo aplicado por unidade",
        "Preço mínimo aplicado por unidade",
    )
    for column, tooltip in enumerate(header_items):
        item = table.horizontalHeaderItem(column)
        if item is not None:
            item.setToolTip(f"{tooltip}. Clique para ordenar.")
    layout.addWidget(table, 1)

    def _cell_host(widget: QWidget, *, left: int = 6, right: int = 6) -> QWidget:
        host = QWidget()
        host_layout = QHBoxLayout(host)
        host_layout.setContentsMargins(left, 3, right, 3)
        host_layout.setSpacing(0)
        host_layout.addWidget(widget)
        return host

    row_controls: list[dict[str, Any]] = []
    for op_name in operations:
        profile = dict(active_map.get(op_name, {}) or {})
        mode_combo = QComboBox()
        for key, label in _OPERATION_PRICING_MODE_ITEMS:
            mode_combo.addItem(label, key)
        current_mode = str(profile.get("pricing_mode", "manual") or "manual").strip() or "manual"
        for combo_index in range(mode_combo.count()):
            if str(mode_combo.itemData(combo_index) or "") == current_mode:
                mode_combo.setCurrentIndex(combo_index)
                break
        driver_edit = QLineEdit(str(profile.get("driver_label", "") or "Qtd./peca").strip())
        driver_units_spin = QDoubleSpinBox()
        driver_units_spin.setRange(0.0, 1000000.0)
        driver_units_spin.setDecimals(4)
        driver_units_spin.setValue(float(profile.get("default_units", 1) or 1))
        setup_spin = QDoubleSpinBox()
        setup_spin.setRange(0.0, 1000000.0)
        setup_spin.setDecimals(3)
        setup_spin.setValue(float(profile.get("setup_min", 0) or 0))
        time_spin = QDoubleSpinBox()
        time_spin.setRange(0.0, 1000000.0)
        time_spin.setDecimals(4)
        time_spin.setValue(float(profile.get("unit_time_min", 0) or 0))
        hour_spin = QDoubleSpinBox()
        hour_spin.setRange(0.0, 1000000.0)
        hour_spin.setDecimals(4)
        hour_spin.setValue(float(profile.get("hour_rate_eur", 0) or 0))
        fixed_spin = QDoubleSpinBox()
        fixed_spin.setRange(0.0, 1000000.0)
        fixed_spin.setDecimals(4)
        fixed_spin.setValue(float(profile.get("fixed_unit_eur", 0) or 0))
        min_spin = QDoubleSpinBox()
        min_spin.setRange(0.0, 1000000.0)
        min_spin.setDecimals(4)
        min_spin.setValue(float(profile.get("min_unit_eur", 0) or 0))
        for widget in (mode_combo, driver_edit, driver_units_spin, setup_spin, time_spin, hour_spin, fixed_spin, min_spin):
            widget.setMaximumHeight(28)
            widget.setStyleSheet("font-size: 10px;")

        hosts = [
            _cell_host(mode_combo),
            _cell_host(driver_edit),
            _cell_host(driver_units_spin),
            _cell_host(setup_spin),
            _cell_host(time_spin),
            _cell_host(hour_spin),
            _cell_host(fixed_spin),
            _cell_host(min_spin),
        ]
        row_controls.append(
            {
                "op_name": op_name,
                "mode_combo": mode_combo,
                "driver_edit": driver_edit,
                "driver_units_spin": driver_units_spin,
                "setup_spin": setup_spin,
                "time_spin": time_spin,
                "hour_spin": hour_spin,
                "fixed_spin": fixed_spin,
                "min_spin": min_spin,
                "hosts": hosts,
            }
        )

    operation_sort_state = {"section": -1, "order": Qt.AscendingOrder}

    def _profile_sort_value(controls: dict[str, Any], section: int) -> tuple[int, object]:
        values: tuple[object, ...] = (
            controls["op_name"],
            controls["mode_combo"].currentText(),
            controls["driver_edit"].text(),
            controls["driver_units_spin"].value(),
            controls["setup_spin"].value(),
            controls["time_spin"].value(),
            controls["hour_spin"].value(),
            controls["fixed_spin"].value(),
            controls["min_spin"].value(),
        )
        return _smart_sort_key(values[max(0, min(section, len(values) - 1))])

    def _render_profile_rows() -> None:
        for controls in row_controls:
            for host in controls["hosts"]:
                host.setParent(None)
        table.clearContents()
        table.setRowCount(0)
        for row_index, controls in enumerate(row_controls):
            table.insertRow(row_index)
            table.setRowHeight(row_index, 36)
            name_item = QTableWidgetItem(str(controls["op_name"]))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            name_item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
            table.setItem(row_index, 0, name_item)
            for column, host in enumerate(controls["hosts"], start=1):
                table.setCellWidget(row_index, column, host)

    def _sort_profile_rows(section: int) -> None:
        if operation_sort_state["section"] == section:
            operation_sort_state["order"] = (
                Qt.DescendingOrder
                if operation_sort_state["order"] == Qt.AscendingOrder
                else Qt.AscendingOrder
            )
        else:
            operation_sort_state["section"] = section
            operation_sort_state["order"] = Qt.AscendingOrder
        reverse = operation_sort_state["order"] == Qt.DescendingOrder
        row_controls.sort(key=lambda controls: _profile_sort_value(controls, section), reverse=reverse)
        header.setSortIndicator(section, operation_sort_state["order"])
        header.setSortIndicatorShown(True)
        _render_profile_rows()

    _render_profile_rows()
    header.setSectionsClickable(True)
    header.setSortIndicatorShown(False)
    header.sectionClicked.connect(_sort_profile_rows)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.setStyleSheet("QPushButton { font-size: 10.5px; min-height: 28px; padding: 4px 10px; }")
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return False

    profile_name = profile_name_edit.text().strip() or "Base"
    profile_payload: dict[str, Any] = {}
    for controls in row_controls:
        mode_combo = controls["mode_combo"]
        profile_payload[str(controls["op_name"])] = {
            "pricing_mode": str(mode_combo.currentData() or "manual"),
            "driver_label": controls["driver_edit"].text().strip() or "Qtd./peca",
            "default_units": controls["driver_units_spin"].value(),
            "setup_min": controls["setup_spin"].value(),
            "unit_time_min": controls["time_spin"].value(),
            "hour_rate_eur": controls["hour_spin"].value(),
            "fixed_unit_eur": controls["fixed_spin"].value(),
            "min_unit_eur": controls["min_spin"].value(),
        }
    merged_profiles = dict(profiles)
    merged_profiles[profile_name] = profile_payload
    backend.operation_cost_save_settings({"active_profile": profile_name, "profiles": merged_profiles})
    return True


def _open_quote_operation_detail_dialog(parent: QWidget, backend, payload: dict[str, Any]) -> dict[str, Any] | None:
    row = dict(payload or {})
    if not str(row.get("operacao", "") or "").strip():
        QMessageBox.warning(parent, "Operacoes", "Seleciona primeiro os postos de trabalho desta linha.")
        return None
    dialog = QDialog(parent)
    dialog.setWindowTitle("Detalhe de custo por operacao")
    dialog.resize(1320, 470)
    dialog.setStyleSheet(
        "QLabel { font-size: 10px; }"
        "QLineEdit, QComboBox, QDoubleSpinBox { font-size: 10px; min-height: 24px; padding: 0 6px; }"
        "QTableWidget { font-size: 10px; gridline-color: #d3dde9; alternate-background-color: #f8fbff; }"
        "QTableWidget::item { padding: 0 8px; }"
        "QHeaderView::section { font-size: 10px; padding: 6px 8px; }"
        "QPushButton { font-size: 10.5px; min-height: 28px; padding: 4px 10px; }"
    )
    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)
    intro = QLabel(
        "Detalha o custo/tempo por posto. Em 'Qtd/peca' introduzes a quantidade tecnica da operacao: "
        "Quinagem = n. de dobras por peca, Roscagem = n. de roscas por peca, Maquinacao = n. de operacoes por peca."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    intro.setStyleSheet("font-size: 10px;")
    layout.addWidget(intro)

    summary_label = QLabel("")
    summary_label.setWordWrap(True)
    summary_label.setProperty("role", "field_value")
    summary_label.setStyleSheet("font-size: 10px; font-weight: 700;")
    layout.addWidget(summary_label)

    table = QTableWidget(0, 11)
    table.setHorizontalHeaderLabels(
        ["Operacao", "Modo", "Tipo qtd.", "Qtd/peca", "Setup", "Tempo base/manual", "EUR/h", "Fixo/manual", "Min/un", "Tempo final/un", "Custo final/un"]
    )
    table.verticalHeader().setVisible(False)
    _configure_table(table, stretch=(), contents=())
    table.setShowGrid(True)
    header = table.horizontalHeader()
    header.setDefaultAlignment(Qt.AlignCenter)
    _set_table_columns(
        table,
        [
            (0, "fixed", 170),
            (1, "fixed", 155),
            (2, "fixed", 190),
            (3, "fixed", 118),
            (4, "fixed", 118),
            (5, "fixed", 150),
            (6, "fixed", 112),
            (7, "fixed", 138),
            (8, "fixed", 112),
            (9, "fixed", 130),
            (10, "fixed", 138),
        ],
    )
    layout.addWidget(table)

    controls_by_op: dict[str, dict[str, Any]] = {}
    last_estimate: dict[str, Any] = {}

    def _cell_host(widget: QWidget, *, left: int = 6, right: int = 6) -> QWidget:
        host = QWidget()
        host_layout = QHBoxLayout(host)
        host_layout.setContentsMargins(left, 3, right, 3)
        host_layout.setSpacing(0)
        host_layout.addWidget(widget)
        return host

    def _driver_tooltip(op_name: str, driver_label: str) -> str:
        normalized = str(op_name or "").strip().lower()
        if "quin" in normalized:
            return "Quinagem: indica aqui o numero de dobras por peca."
        if "rosc" in normalized:
            return "Roscagem: indica aqui o numero de roscas por peca."
        if "maquin" in normalized:
            return "Maquinacao: indica aqui o numero de operacoes de maquinação por peca."
        if "sold" in normalized:
            return "Soldadura: indica aqui o numero de pontos/cordoes por peca."
        if "serralh" in normalized:
            return "Serralharia: indica aqui o numero de operacoes por peca."
        if "laca" in normalized:
            return "Lacagem: indica aqui a area em m2 por peca."
        if "montag" in normalized:
            return "Montagem: indica aqui o numero de operacoes por peca."
        if "embal" in normalized:
            return "Embalamento: indica aqui o numero de volumes por peca."
        return f"Indica aqui o valor tecnico de '{driver_label or 'Qtd/peca'}'."

    def _refresh_table_height() -> None:
        visible_rows = min(max(4, table.rowCount()), 8)
        target_height = _table_visible_height(table, visible_rows, extra=24)
        table.setMinimumHeight(target_height)
        table.setMaximumHeight(target_height)

    def _set_combo_value(combo: QComboBox, value: str) -> None:
        for combo_index in range(combo.count()):
            if str(combo.itemData(combo_index) or "") == str(value or ""):
                combo.setCurrentIndex(combo_index)
                return

    def _build_payload_from_widgets() -> dict[str, Any]:
        detail_rows: list[dict[str, Any]] = []
        for op_name, controls in controls_by_op.items():
            mode = str(controls["mode_combo"].currentData() or "manual")
            manual_time_value = controls["time_spin"].value()
            manual_cost_value = controls["fixed_spin"].value()
            manual_confirmed = bool(controls.get("manual_confirmed", False))
            if mode == "manual" and (abs(float(manual_time_value or 0)) > 0.000001 or abs(float(manual_cost_value or 0)) > 0.000001):
                manual_confirmed = True
            detail_rows.append(
                {
                    "nome": op_name,
                    "pricing_mode": mode,
                    "driver_label": controls["driver_edit"].text().strip() or "Qtd./peca",
                    "driver_units": controls["driver_units_spin"].value(),
                    "driver_units_confirmed": mode != "manual",
                    "manual_values_confirmed": manual_confirmed,
                    "setup_min": controls["setup_spin"].value(),
                    "unit_time_base_min": controls["time_spin"].value(),
                    "hour_rate_eur": controls["hour_spin"].value(),
                    "fixed_unit_eur": controls["fixed_spin"].value(),
                    "min_unit_eur": controls["min_spin"].value(),
                    "tempo_unit_min": controls["time_spin"].value() if mode == "manual" and manual_confirmed else None,
                    "custo_unit_eur": controls["fixed_spin"].value() if mode == "manual" and manual_confirmed else None,
                }
            )
        return {**row, "operacoes_detalhe": detail_rows}

    def _apply_estimate(estimate: dict[str, Any], *, overwrite_inputs: bool) -> None:
        operations = list(estimate.get("operations", []) or [])
        summary = dict(estimate.get("summary", {}) or {})
        profile_name = str(estimate.get("active_profile", "") or "").strip() or "Base"
        if overwrite_inputs:
            table.setRowCount(len(operations))
            for current_row in range(len(operations)):
                table.setRowHeight(current_row, 34)
        for row_index, op_row in enumerate(operations):
            op_name = str(op_row.get("nome", "") or "").strip()
            if not op_name:
                continue
            controls = controls_by_op.get(op_name)
            if controls is None:
                op_item = QTableWidgetItem(op_name)
                op_item.setFlags(op_item.flags() & ~Qt.ItemIsEditable)
                op_item.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
                table.setItem(row_index, 0, op_item)
                mode_combo = QComboBox()
                for key, label in _OPERATION_PRICING_MODE_ITEMS:
                    mode_combo.addItem(label, key)
                driver_edit = QLineEdit()
                driver_units_spin = QDoubleSpinBox()
                driver_units_spin.setRange(0.0, 1000000.0)
                driver_units_spin.setDecimals(4)
                setup_spin = QDoubleSpinBox()
                setup_spin.setRange(0.0, 1000000.0)
                setup_spin.setDecimals(4)
                time_spin = QDoubleSpinBox()
                time_spin.setRange(0.0, 1000000.0)
                time_spin.setDecimals(4)
                hour_spin = QDoubleSpinBox()
                hour_spin.setRange(0.0, 1000000.0)
                hour_spin.setDecimals(4)
                fixed_spin = QDoubleSpinBox()
                fixed_spin.setRange(0.0, 1000000.0)
                fixed_spin.setDecimals(4)
                min_spin = QDoubleSpinBox()
                min_spin.setRange(0.0, 1000000.0)
                min_spin.setDecimals(4)
                for widget in (mode_combo, driver_edit, driver_units_spin, setup_spin, time_spin, hour_spin, fixed_spin, min_spin):
                    widget.setMaximumHeight(28)
                    widget.setStyleSheet("font-size: 10px;")
                table.setCellWidget(row_index, 1, _cell_host(mode_combo))
                table.setCellWidget(row_index, 2, _cell_host(driver_edit))
                table.setCellWidget(row_index, 3, _cell_host(driver_units_spin))
                table.setCellWidget(row_index, 4, _cell_host(setup_spin))
                table.setCellWidget(row_index, 5, _cell_host(time_spin))
                table.setCellWidget(row_index, 6, _cell_host(hour_spin))
                table.setCellWidget(row_index, 7, _cell_host(fixed_spin))
                table.setCellWidget(row_index, 8, _cell_host(min_spin))
                controls = {
                    "mode_combo": mode_combo,
                    "driver_edit": driver_edit,
                    "driver_units_spin": driver_units_spin,
                    "setup_spin": setup_spin,
                    "time_spin": time_spin,
                    "hour_spin": hour_spin,
                    "fixed_spin": fixed_spin,
                    "min_spin": min_spin,
                    "manual_confirmed": False,
                }
                controls_by_op[op_name] = controls
            if overwrite_inputs:
                _set_combo_value(controls["mode_combo"], str(op_row.get("pricing_mode", "manual") or "manual"))
                driver_label_txt = str(op_row.get("driver_label", "") or "Qtd./peca").strip()
                controls["driver_edit"].setText(driver_label_txt)
                controls["driver_units_spin"].setValue(float(op_row.get("driver_units", 0) or 0))
                controls["setup_spin"].setValue(float(op_row.get("setup_min", 0) or 0))
                controls["time_spin"].setValue(float(op_row.get("unit_time_base_min", op_row.get("tempo_unit_min", 0)) or 0))
                controls["hour_spin"].setValue(float(op_row.get("hour_rate_eur", 0) or 0))
                controls["fixed_spin"].setValue(float(op_row.get("fixed_unit_eur", op_row.get("custo_unit_eur", 0)) or 0))
                controls["min_spin"].setValue(float(op_row.get("min_unit_eur", 0) or 0))
                controls["manual_confirmed"] = bool(op_row.get("manual_values_confirmed", False))
                tip = _driver_tooltip(op_name, driver_label_txt)
                controls["driver_edit"].setToolTip(tip)
                controls["driver_units_spin"].setToolTip(tip)
            time_result = QTableWidgetItem("-" if op_row.get("tempo_unit_min") in (None, "") else f"{float(op_row.get('tempo_unit_min', 0) or 0):.3f}")
            cost_result = QTableWidgetItem("-" if op_row.get("custo_unit_eur") in (None, "") else _fmt_eur(float(op_row.get("custo_unit_eur", 0) or 0)))
            time_result.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            cost_result.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
            table.setItem(row_index, 9, time_result)
            table.setItem(row_index, 10, cost_result)
        mode_txt = str(summary.get("costing_mode", "") or "").strip() or "aggregate_pending"
        pending_rows = [dict(item) for item in operations if isinstance(item, dict) and bool(item.get("missing_driver_input"))]
        blend_with_current_line = bool(row.get("blend_with_current_line", False))
        base_time = float(row.get("base_tempo_unit_min", row.get("tempo_peca_min", 0)) or 0)
        base_cost = float(row.get("base_preco_unit_eur", row.get("preco_unit", 0)) or 0)
        state_txt = {
            "detailed": "Detalhe completo",
            "partial_detail": "Detalhe parcial",
            "aggregate_pending": "Ainda agregado",
            "single_operation_total": "Linha simples",
        }.get(mode_txt, mode_txt)
        if pending_rows:
            missing_txt = ", ".join(
                f"{str(item.get('nome', '') or '').strip()} ({str(item.get('driver_label', '') or 'Qtd./peca').strip()})"
                for item in pending_rows
                if str(item.get("nome", "") or "").strip()
            )
            summary_label.setText(f"Perfil {profile_name} | falta quantificar: {missing_txt}")
        elif blend_with_current_line and (base_time > 0 or base_cost > 0):
            extra_time = float(summary.get("tempo_unit_total_min", 0) or 0)
            extra_cost = float(summary.get("custo_unit_total_eur", 0) or 0)
            summary_label.setText(
                f"Perfil {profile_name} | base atual {_fmt_eur(base_cost)}/un + extras {_fmt_eur(extra_cost)}/un = "
                f"{_fmt_eur(base_cost + extra_cost)}/un | tempo {base_time:.3f} + {extra_time:.3f} = {base_time + extra_time:.3f} min/un"
            )
        else:
            summary_label.setText(
                f"Perfil {profile_name} | {state_txt} | sugestao {float(summary.get('tempo_unit_total_min', 0) or 0):.3f} min/un "
                f"| {_fmt_eur(float(summary.get('custo_unit_total_eur', 0) or 0))}/un"
            )
        _refresh_table_height()

    def recompute() -> None:
        nonlocal last_estimate
        last_estimate = dict(backend.operation_cost_estimate(_build_payload_from_widgets()) or {})
        _apply_estimate(last_estimate, overwrite_inputs=False)

    initial_estimate = dict(backend.operation_cost_estimate(row) or {})
    _apply_estimate(initial_estimate, overwrite_inputs=True)
    last_estimate = dict(initial_estimate)

    for controls in controls_by_op.values():
        controls["mode_combo"].currentIndexChanged.connect(lambda _idx: recompute())
        controls["driver_edit"].textChanged.connect(lambda _txt: recompute())
        controls["driver_units_spin"].valueChanged.connect(lambda _val: recompute())
        controls["setup_spin"].valueChanged.connect(lambda _val: recompute())
        controls["time_spin"].valueChanged.connect(lambda _val: recompute())
        controls["hour_spin"].valueChanged.connect(lambda _val: recompute())
        controls["fixed_spin"].valueChanged.connect(lambda _val: recompute())
        controls["min_spin"].valueChanged.connect(lambda _val: recompute())

    action_row = QHBoxLayout()
    action_row.setContentsMargins(2, 0, 2, 0)
    action_row.setSpacing(8)
    apply_profiles_btn = QPushButton("Aplicar perfis")
    apply_profiles_btn.setProperty("variant", "secondary")
    config_profiles_btn = QPushButton("Configurar perfis")
    config_profiles_btn.setProperty("variant", "secondary")
    action_row.addWidget(apply_profiles_btn)
    action_row.addWidget(config_profiles_btn)
    action_row.addStretch(1)
    layout.addLayout(action_row)

    def apply_profiles_defaults() -> None:
        nonlocal last_estimate
        clean_payload = dict(row)
        clean_payload["operacoes_detalhe"] = []
        last_estimate = dict(backend.operation_cost_estimate(clean_payload) or {})
        _apply_estimate(last_estimate, overwrite_inputs=True)
        recompute()

    apply_profiles_btn.clicked.connect(apply_profiles_defaults)

    def configure_profiles() -> None:
        if not _open_operation_cost_profiles_dialog(dialog, backend):
            return
        apply_profiles_defaults()

    config_profiles_btn.clicked.connect(configure_profiles)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.Accepted:
        return None

    final_estimate = dict(backend.operation_cost_estimate(_build_payload_from_widgets()) or {})
    final_rows = [dict(item or {}) for item in list(final_estimate.get("operations", []) or []) if isinstance(item, dict)]
    tempos_operacao = {
        str(item.get("nome", "") or "").strip(): float(item.get("tempo_unit_min", 0) or 0)
        for item in final_rows
        if str(item.get("nome", "") or "").strip() and item.get("tempo_unit_min") not in (None, "")
    }
    custos_operacao = {
        str(item.get("nome", "") or "").strip(): float(item.get("custo_unit_eur", 0) or 0)
        for item in final_rows
        if str(item.get("nome", "") or "").strip() and item.get("custo_unit_eur") not in (None, "")
    }
    summary = dict(final_estimate.get("summary", {}) or {})
    base_time_unit = round(float(row.get("base_tempo_unit_min", row.get("tempo_peca_min", 0)) or 0), 4)
    base_price_unit = round(float(row.get("base_preco_unit_eur", row.get("preco_unit", 0)) or 0), 4)
    blend_with_current_line = bool(row.get("blend_with_current_line", False))
    detailed_extras_time = round(float(summary.get("tempo_unit_total_min", 0) or 0), 4)
    detailed_extras_price = round(float(summary.get("custo_unit_total_eur", 0) or 0), 4)
    suggested_time_unit = round(base_time_unit + detailed_extras_time, 4) if blend_with_current_line else detailed_extras_time
    suggested_price_unit = round(base_price_unit + detailed_extras_price, 4) if blend_with_current_line else detailed_extras_price
    return {
        "operacoes_detalhe": final_rows,
        "tempos_operacao": tempos_operacao,
        "custos_operacao": custos_operacao,
        "quote_cost_snapshot": {
            "costing_mode": str(summary.get("costing_mode", "") or ""),
            "tempo_total_peca_min": round(float(summary.get("tempo_unit_total_min", 0) or 0), 4),
            "preco_unit_total_eur": round(float(summary.get("custo_unit_total_eur", 0) or 0), 4),
            "qtd": round(float(summary.get("qtd", row.get("qtd", 0)) or 0), 2),
            "cost_profile": str(final_estimate.get("active_profile", "") or "").strip(),
        },
        "suggested_tempo_unit_min": suggested_time_unit,
        "suggested_preco_unit_eur": suggested_price_unit,
        "apply_totals": bool(summary.get("complete", False) or blend_with_current_line),
        "blended_with_current_line": blend_with_current_line,
        "summary": summary,
    }


def _take_layout_items(layout) -> list:
    items = []
    while layout.count():
        items.append(layout.takeAt(0))
    return items


def _adopt_layout_item(target_layout, item, stretch: int = 0) -> None:
    if item is None:
        return
    widget = item.widget()
    child_layout = item.layout()
    spacer = item.spacerItem()
    if isinstance(target_layout, QSplitter):
        if widget is not None:
            target_layout.addWidget(widget)
            return
        if child_layout is not None:
            host = QWidget()
            host.setLayout(child_layout)
            target_layout.addWidget(host)
            return
        return
    if widget is not None:
        target_layout.addWidget(widget, stretch)
        return
    if child_layout is not None:
        target_layout.addLayout(child_layout, stretch)
        return
    if spacer is not None:
        target_layout.addItem(spacer)


def _clear_layout_widgets(layout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        child_layout = item.layout()
        if widget is not None:
            widget.deleteLater()
        elif child_layout is not None:
            _clear_layout_widgets(child_layout)
