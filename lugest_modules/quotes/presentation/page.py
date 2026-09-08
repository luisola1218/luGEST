from __future__ import annotations
from lugest_modules.quotes.presentation.page_services import QuotePageServices
from lugest_modules.quotes.presentation.workspace import QuoteWorkspace, QuoteWorkspaceActions
from lugest_modules.quotes.presentation.profile_editor import edit_profile
from lugest_modules.quotes.domain.lines import service_line, product_line
from lugest_modules.quotes.presentation.material_prices import edit_material_prices
from lugest_modules.quotes.presentation.material_prices import sync_stock_price
from lugest_modules.quotes.presentation.material_editor import edit_material
from lugest_modules.quotes.presentation.labor_editor import edit_labor
from lugest_modules.quotes.presentation.consumable_editor import edit_consumable
from lugest_modules.quotes.presentation.product_editor import edit_product
from lugest_modules.quotes.presentation.structure_editor import edit_structure
from lugest_modules.quotes.presentation.line_editor import edit_line
import html
import os
import tempfile
import unicodedata
from PySide6.QtCore import QDate, QEvent, Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QGridLayout, QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QPushButton, QSplitter, QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from lugest_qt.ui.pages.runtime_common import apply_state_chip as _apply_state_chip, configure_table as _configure_table, fill_table as _fill_table, paint_table_row as _paint_table_row, repolish as _repolish, run_process_async as _run_process_async, set_panel_tone as _set_panel_tone, set_table_columns as _set_table_columns, smart_sort_key as _smart_sort_key, state_tone as _state_tone, table_visible_height as _table_visible_height
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


class QuotePage(QWidget):
    page_title = "Orçamentos"
    page_subtitle = "Lista de orçamentos primeiro e detalhe apenas quando abres o registo."
    uses_backend_reload = True

    LINE_COL_MARK = 0
    LINE_COL_TYPE = 1
    LINE_COL_REFERENCE = 2
    LINE_COL_EXTERNAL_REFERENCE = 3
    LINE_COL_DESCRIPTION = 4
    LINE_COL_MATERIAL = 5
    LINE_COL_UNIT = 6
    LINE_COL_OPERATION = 7
    LINE_COL_TIME = 8
    LINE_COL_QUANTITY = 9
    LINE_COL_PRICE = 10
    LINE_COL_DISCOUNTED_PRICE = 11
    LINE_COL_TOTAL = 12
    LINE_COL_ASSEMBLY = 13

    def __init__(self, services: QuotePageServices, parent=None) -> None:
        super().__init__(parent)
        self.services = services
        self.rows: list[dict] = []
        self.client_rows: list[dict] = []
        self.line_rows: list[dict] = []
        self.presets: dict = {}
        self.current_number = ""
        self._quote_lines_sort_section = -1
        self._quote_lines_sort_order = Qt.AscendingOrder
        self.view = QuoteWorkspace(self)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(self.view)
        actions = QuoteWorkspaceActions(
            _add_laser_batch_lines=self._add_laser_batch_lines,
            _add_line=self._add_line,
            _append_pdf_note=self._append_pdf_note,
            _apply_default_delivery_deadline=self._apply_default_delivery_deadline,
            _apply_transport_calc=self._apply_transport_calc,
            _check_selected_line_weight=self._check_selected_line_weight,
            _clear_transport=self._clear_transport,
            _configure_laser_profiles=self._configure_laser_profiles,
            _configure_operation_profiles=self._configure_operation_profiles,
            _convert_quote=self._convert_quote,
            _create_quote_purchase_note=self._create_quote_purchase_note,
            _edit_line=self._edit_line,
            _fill_client_from_combo=self._fill_client_from_combo,
            _fill_pdf_notes_from_context=self._fill_pdf_notes_from_context,
            _handle_quote_line_item_changed=self._handle_quote_line_item_changed,
            _handle_quote_lines_sort=self._handle_quote_lines_sort,
            _manage_assembly_models=self._manage_assembly_models,
            _manage_saved_conjuntos=self._manage_saved_conjuntos,
            _new_quote=self._new_quote,
            _new_structure_quote=self._new_structure_quote,
            _open_calculated_assembly_builder=self._open_calculated_assembly_builder,
            _open_laser_nesting=self._open_laser_nesting,
            _open_line_drawing=self._open_line_drawing,
            _open_profile_step_igs_quote_builder=self._open_profile_step_igs_quote_builder,
            _open_selected_quote=self._open_selected_quote,
            _open_structure_quote_builder=self._open_structure_quote_builder,
            _pick_quote_discount_groups=self._pick_quote_discount_groups,
            _prepare_selected_line_for_production=self._prepare_selected_line_for_production,
            _preview_quote=self._preview_quote,
            _print_quote_pdf=self._print_quote_pdf,
            _recalc_transport_calc=self._recalc_transport_calc,
            _refresh_quote_identity_label=self._refresh_quote_identity_label,
            _remove_line=self._remove_line,
            _remove_quote=self._remove_quote,
            _render_quote_lines=self._render_quote_lines,
            _reset_quote_save_button=self._reset_quote_save_button,
            _save_quote=self._save_quote,
            _save_quote_pdf=self._save_quote_pdf,
            _save_selected_lines_as_group=self._save_selected_lines_as_group,
            _set_quote_state=self._set_quote_state,
            _show_list=self._show_list,
            _sync_list_buttons=self._sync_list_buttons,
            _sync_quote_selected_line_actions=self._sync_quote_selected_line_actions,
            _toggle_quote_inspector=self._toggle_quote_inspector,
            refresh=self.refresh,
        )
        self.view.build(actions)
        self._clear_quote_detail()
        self._show_list()

    def eventFilter(self, watched, event):  # type: ignore[override]
        combo = self.view._combo_click_targets.get(watched)
        if combo is not None and event.type() == QEvent.MouseButtonPress:
            combo.showPopup()
            return True
        return super().eventFilter(watched, event)

    def refresh(self) -> None:
        previous = self.current_number
        keep_detail = self.view.view_stack.currentWidget() is self.view.detail_page and bool(previous)
        current_filter = self.view.filter_edit.currentText().strip()
        self.client_rows = self.services.orc_clients()
        self.presets = self.services.order_presets()
        self._set_client_items()
        current_carrier = self.view.transport_carrier_combo.currentText().strip()
        self.view.transport_carrier_combo.blockSignals(True)
        self.view.transport_carrier_combo.clear()
        self.view.transport_carrier_combo.addItem("")
        for supplier in list(self.services.ne_suppliers() or []):
            self.view.transport_carrier_combo.addItem(f"{supplier.get('id', '')} - {supplier.get('nome', '')}".strip(" -"))
        self.view.transport_carrier_combo.setCurrentText(current_carrier)
        self.view.transport_carrier_combo.blockSignals(False)
        current_zone = self.view.transport_zone_combo.currentText().strip()
        self.view.transport_zone_combo.blockSignals(True)
        self.view.transport_zone_combo.clear()
        self.view.transport_zone_combo.addItem("")
        for value in list(self.services.transport_zone_options() or []):
            self.view.transport_zone_combo.addItem(str(value))
        self.view.transport_zone_combo.setCurrentText(current_zone)
        self.view.transport_zone_combo.blockSignals(False)
        current_year = self.view.year_combo.currentText().strip() or "Todos"
        year_values = ["Todos"] + list(self.services.orc_available_years())
        self.view.year_combo.blockSignals(True)
        self.view.year_combo.clear()
        self.view.year_combo.addItems(year_values)
        self.view.year_combo.setCurrentText(current_year if current_year in year_values else year_values[0])
        self.view.year_combo.blockSignals(False)
        self.rows = self.services.orc_rows(current_filter, self.view.state_combo.currentText(), self.view.year_combo.currentText() or "Todos")
        self._refresh_quote_overview()
        self.view.filter_edit.blockSignals(True)
        if self.view.filter_edit.count() == 0:
            self.view.filter_edit.addItem("")
        known_values = {self.view.filter_edit.itemText(i) for i in range(self.view.filter_edit.count())}
        for row in self.rows:
            numero = str(row.get("numero", "") or "")
            if numero and numero not in known_values:
                self.view.filter_edit.addItem(numero)
                known_values.add(numero)
        self.view.filter_edit.setCurrentText(current_filter)
        self.view.filter_edit.blockSignals(False)
        _fill_table(
            self.view.table,
            [
                [
                    r.get("numero", "-"),
                    r.get("cliente", "-"),
                    r.get("data", "-"),
                    r.get("estado", "-"),
                    _fmt_eur(r.get("total", 0)),
                    r.get("numero_encomenda", "-"),
                ]
                for r in self.rows
            ],
            align_center_from=2,
        )
        for row_index, row in enumerate(self.rows):
            _paint_table_row(self.view.table, row_index, str(row.get("estado", "")))
            total_item = self.view.table.item(row_index, 4)
            if total_item is not None:
                total_item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
        if self.view.table.rowCount() == 0:
            self._clear_quote_detail()
            self._show_list()
            self._sync_list_buttons()
            return
        row_index = 0
        if previous:
            for index, row in enumerate(self.rows):
                if str(row.get("numero", "") or "").strip() == previous:
                    row_index = index
                    break
        self.view.table.selectRow(row_index)
        if keep_detail:
            self._load_quote(str(self.rows[row_index].get("numero", "") or "").strip())
            self._show_detail()
        else:
            self._show_list()
        self._sync_list_buttons()

    def _refresh_quote_overview(self) -> None:
        rows = list(self.rows or [])
        state_counts = {"editing": 0, "sent": 0, "approved": 0}
        total_value = 0.0
        for row in rows:
            state = str(row.get("estado", "") or "").strip().casefold()
            total_value += float(row.get("total", 0) or 0)
            if "edi" in state:
                state_counts["editing"] += 1
            elif "enviado" in state:
                state_counts["sent"] += 1
            elif "aprovado" in state:
                state_counts["approved"] += 1
        labels = getattr(self.view, "quote_metric_labels", {})
        if isinstance(labels, dict):
            if isinstance(labels.get("value"), QLabel):
                labels["value"].setText(_fmt_eur(total_value))
            for key in ("editing", "sent", "approved"):
                if isinstance(labels.get(key), QLabel):
                    labels[key].setText(str(state_counts[key]))
        count = len(rows)
        self.view.quote_list_count_label.setText(f"{count} orçamento" if count == 1 else f"{count} orçamentos")

    def _show_list(self) -> None:
        self.view.view_stack.setCurrentWidget(self.view.list_page)
        self._sync_list_buttons()

    def _show_detail(self) -> None:
        self.view.view_stack.setCurrentWidget(self.view.detail_page)

    def _toggle_quote_inspector(self) -> None:
        inspector = getattr(self.view, "quote_inspector_tabs", None)
        splitter = getattr(self.view, "quote_workspace_split", None)
        button = getattr(self.view, "quote_inspector_btn", None)
        if not isinstance(inspector, QWidget) or not isinstance(splitter, QSplitter):
            return
        show_inspector = not inspector.isVisible()
        inspector.setVisible(show_inspector)
        if show_inspector:
            splitter.setSizes([max(760, splitter.width() - 420), 420])
        if isinstance(button, QPushButton):
            button.setText("Ocultar painel" if show_inspector else "Mostrar painel")

    def can_auto_refresh(self) -> bool:
        return self.view.view_stack.currentWidget() is self.view.list_page

    def _selected_quote_row(self) -> dict:
        current = self.view.table.currentItem()
        if current is None or current.row() >= len(self.rows):
            return {}
        return self.rows[current.row()]

    def _selected_line_index(self) -> int:
        current = self.view.lines_table.currentItem()
        if current is None:
            return -1
        source_index = current.data(Qt.UserRole)
        if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
            return source_index
        return current.row() if 0 <= current.row() < len(self.line_rows) else -1

    def _selected_line_indexes(self) -> list[int]:
        indexes: set[int] = set()
        for item in self.view.lines_table.selectedItems():
            if item is None:
                continue
            source_index = item.data(Qt.UserRole)
            if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
                indexes.add(source_index)
            elif 0 <= item.row() < len(self.line_rows):
                indexes.add(item.row())
        indexes = sorted(indexes)
        if indexes:
            return indexes
        single = self._selected_line_index()
        return [single] if single >= 0 else []

    def _checked_line_indexes(self) -> list[int]:
        indexes: set[int] = set()
        for visual_row in range(self.view.lines_table.rowCount()):
            item = self.view.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is None or item.checkState() != Qt.Checked:
                continue
            source_index = item.data(Qt.UserRole)
            if isinstance(source_index, int) and 0 <= source_index < len(self.line_rows):
                indexes.add(source_index)
        return sorted(indexes)

    def _handle_quote_line_item_changed(self, item: QTableWidgetItem) -> None:
        if bool(getattr(self, "_rendering_quote_lines", False)):
            return
        if item is not None and item.column() == self.LINE_COL_MARK:
            self._sync_quote_selected_line_actions()

    def _select_quote_line_source_index(self, source_index: int) -> None:
        for visual_row in range(self.view.lines_table.rowCount()):
            item = self.view.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is not None and item.data(Qt.UserRole) == source_index:
                self.view.lines_table.selectRow(visual_row)
                return

    def _sync_quote_selected_line_actions(self) -> None:
        indexes = self._selected_line_indexes() if hasattr(self.view, "lines_table") else []
        checked_indexes = self._checked_line_indexes() if hasattr(self.view, "lines_table") else []
        has_selection = bool(indexes)
        for button in getattr(self.view, "quote_selected_line_buttons", ()):
            button.setEnabled(has_selection)
        remove_button = getattr(self.view, "remove_quote_lines_btn", None)
        removal_count = len(checked_indexes or indexes)
        if isinstance(remove_button, QPushButton):
            remove_button.setEnabled(removal_count > 0)
            remove_button.setText("Remover linha" if removal_count <= 1 else f"Remover {removal_count} linhas")
        caption = getattr(self.view, "quote_selected_line_caption", None)
        if not isinstance(caption, QLabel):
            return
        if checked_indexes:
            caption.setText(
                "1 LINHA MARCADA PARA REMOVER"
                if len(checked_indexes) == 1
                else f"{len(checked_indexes)} LINHAS MARCADAS PARA REMOVER"
            )
            return
        if not indexes:
            caption.setText("SELECIONA UMA LINHA")
            return
        if len(indexes) > 1:
            caption.setText(f"{len(indexes)} LINHAS SELECIONADAS")
            return
        row = dict(self.line_rows[indexes[0]] or {}) if indexes[0] < len(self.line_rows) else {}
        reference = self._quote_line_primary_ref(row)
        caption.setText(f"LINHA · {reference}" if reference and reference != "-" else "LINHA SELECIONADA")

    def _quote_line_sort_value(self, row: dict, section: int) -> tuple[int, object]:
        values: tuple[object, ...] = (
            self._quote_line_type_label(row),
            self._quote_line_primary_ref(row),
            row.get("ref_externa", ""),
            row.get("descricao", ""),
            self._quote_line_material_display(row),
            row.get("espessura", "") or self._quote_line_unit_display(row),
            row.get("operacao", ""),
            row.get("tempo_peca_min", 0),
            row.get("qtd", 0),
            row.get("preco_unit_incrementado", row.get("preco_unit", 0)),
            row.get("preco_unit_desconto", row.get("preco_unit", 0)),
            row.get("total_desconto", row.get("total", 0)),
            row.get("conjunto_nome", ""),
        )
        data_section = max(0, int(section) - 1)
        return _smart_sort_key(values[max(0, min(data_section, len(values) - 1))])

    def _handle_quote_lines_sort(self, section: int) -> None:
        if section == self.LINE_COL_MARK:
            real_items = [
                self.view.lines_table.item(row, self.LINE_COL_MARK)
                for row in range(self.view.lines_table.rowCount())
                if isinstance(self.view.lines_table.item(row, self.LINE_COL_MARK), QTableWidgetItem)
                and isinstance(self.view.lines_table.item(row, self.LINE_COL_MARK).data(Qt.UserRole), int)
            ]
            mark_all = bool(real_items) and not all(item.checkState() == Qt.Checked for item in real_items)
            self._rendering_quote_lines = True
            for item in real_items:
                item.setCheckState(Qt.Checked if mark_all else Qt.Unchecked)
            self._rendering_quote_lines = False
            self._sync_quote_selected_line_actions()
            return
        if self._quote_lines_sort_section == section:
            self._quote_lines_sort_order = (
                Qt.DescendingOrder
                if self._quote_lines_sort_order == Qt.AscendingOrder
                else Qt.AscendingOrder
            )
        else:
            self._quote_lines_sort_section = section
            self._quote_lines_sort_order = Qt.AscendingOrder
        header = self.view.lines_table.horizontalHeader()
        header.setSortIndicator(section, self._quote_lines_sort_order)
        header.setSortIndicatorShown(True)
        self._render_quote_lines()

    def _sync_list_buttons(self) -> None:
        has_row = bool(self._selected_quote_row())
        self.view.open_quote_btn.setEnabled(has_row)
        self.view.remove_quote_btn.setEnabled(has_row)

    def _set_client_items(self) -> None:
        current = self.view.client_combo.currentText().strip()
        self.view.client_combo.blockSignals(True)
        self.view.client_combo.clear()
        for row in self.client_rows:
            label = str(row.get("label", "") or f"{row.get('codigo', '')} - {row.get('nome', '')}").strip(" -")
            self.view.client_combo.addItem(label, row)
        self.view.client_combo.setCurrentText(current)
        self.view.client_combo.blockSignals(False)
        current_exec = self.view.executed_combo.currentText().strip()
        self.view.executed_combo.clear()
        for value in list(self.services.quote_authors() or []):
            self.view.executed_combo.addItem(str(value))
        self.view.executed_combo.setCurrentText(current_exec)
        self.view.workcenter_combo.blockSignals(True)
        self.view.workcenter_combo.clear()
        self.view.workcenter_combo.blockSignals(False)

    def _client_lookup(self, text: str) -> dict:
        raw = str(text or "").strip()
        if not raw:
            return {}
        code = raw.split(" - ", 1)[0].strip()
        for row in self.client_rows:
            if code and code == str(row.get("codigo", "") or "").strip():
                return row
            label = str(row.get("label", "") or f"{row.get('codigo', '')} - {row.get('nome', '')}").strip(" -")
            if raw.lower() == label.lower():
                return row
        return {}

    def _client_code_from_text(self, text: str) -> str:
        return str(self._client_lookup(text).get("codigo", "") or text.split(" - ", 1)[0].strip())

    def _fill_client_from_combo(self) -> None:
        row = self._client_lookup(self.view.client_combo.currentText())
        if not row:
            self._refresh_quote_identity_label()
            return
        if not self.view.client_name_edit.text().strip():
            self.view.client_name_edit.setText(str(row.get("nome", "") or "").strip())
        if not self.view.client_company_edit.text().strip():
            self.view.client_company_edit.setText(str(row.get("nome", "") or "").strip())
        if not self.view.client_nif_edit.text().strip():
            self.view.client_nif_edit.setText(str(row.get("nif", "") or "").strip())
        if not self.view.client_address_edit.text().strip():
            self.view.client_address_edit.setText(str(row.get("morada", "") or "").strip())
        if not self.view.client_contact_edit.text().strip():
            self.view.client_contact_edit.setText(str(row.get("contacto", "") or "").strip())
        if not self.view.client_email_edit.text().strip():
            self.view.client_email_edit.setText(str(row.get("email", "") or "").strip())
        self._refresh_quote_identity_label()

    def _quote_line_type(self, row: dict) -> str:
        return str(self.services.normalize_orc_line_type((row or {}).get("tipo_item")) or self.services.ORC_LINE_TYPE_PIECE)

    def _quote_line_is_raw_material_ui(self, row: dict) -> bool:
        data = dict(row or {})
        if self._quote_line_type(data) != self.services.ORC_LINE_TYPE_PIECE:
            return False
        if str(data.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(data.get("stock_material_id", "") or "").strip():
            return True
        ref_ext = str(data.get("ref_externa", "") or "").strip().upper()
        if ref_ext.startswith("MAT") and ref_ext[3:].isdigit():
            if str(data.get("desenho", "") or "").strip():
                return False
            try:
                tempo = float(data.get("tempo_peca_min", data.get("tempo_pecas_min", 0)) or 0)
            except Exception:
                tempo = 0.0
            if tempo > 0:
                return False
            try:
                op_norm = self.services.norm_text(str(data.get("operacao", "") or "").strip())
            except Exception:
                op_norm = str(data.get("operacao", "") or "").strip().lower()
            return op_norm in {"", "-", "stockmp", "materia prima", "materia-prima"}
        subtype = str(data.get("material_subtype", "") or data.get("calc_mode", "") or "").strip()
        try:
            return self.services.norm_text(subtype) == "stockmp"
        except Exception:
            return subtype.lower() == "stockmp"

    def _quote_line_should_be_raw_material_ui(self, row: dict) -> bool:
        if self._quote_line_is_raw_material_ui(row):
            return True
        data = dict(row or {})
        if self._quote_line_type(data) != self.services.ORC_LINE_TYPE_PIECE:
            return False
        if str(data.get("desenho", "") or "").strip():
            return False
        try:
            tempo = float(data.get("tempo_peca_min", data.get("tempo_pecas_min", 0)) or 0)
        except Exception:
            tempo = 0.0
        if tempo > 0:
            return False
        operacao = str(data.get("operacao", "") or "").strip()
        if operacao:
            try:
                op_norm = self.services.norm_text(operacao)
            except Exception:
                op_norm = operacao.lower()
            if op_norm not in {"stockmp", "materia prima", "materia-prima"}:
                return False
        has_raw_origin = any(
            str(data.get(key, "") or "").strip()
            for key in (
                "conjunto_codigo",
                "conjunto_nome",
                "calc_mode",
                "material_subtype",
                "quantity_units",
                "weight_total",
                "stock_material_id",
            )
        )
        if not has_raw_origin:
            return False
        marker_text = " ".join(
            str(data.get(key, "") or "")
            for key in ("ref_externa", "material", "material_family", "material_subtype", "calc_mode", "descricao")
        )
        try:
            marker_norm = self.services.norm_text(marker_text)
        except Exception:
            marker_norm = marker_text.lower()
        raw_tokens = (
            "perfil",
            "ipe",
            "ipn",
            "hea",
            "heb",
            "upn",
            "rebar",
            "ferro nervurado",
            "barra",
            "chapa",
            "tubo",
            "cantoneira",
            "varao",
            "mat",
            "stock mp",
            "stockmp",
        )
        return any(token in marker_norm for token in raw_tokens)

    def _coerce_quote_line_raw_material_ui(self, row: dict) -> dict:
        data = dict(row or {})
        data["tipo_item"] = self.services.ORC_LINE_TYPE_PIECE
        data["stock_item_kind"] = "raw_material"
        data["ref_interna"] = ""
        data["operacao"] = ""
        data["tempo_peca_min"] = 0.0
        data["desenho"] = ""
        data["laser_base_active"] = False
        data["laser_base_tempo_unit"] = 0.0
        data["laser_base_preco_unit"] = 0.0
        for key in ("operacoes_lista", "operacoes_fluxo", "operacoes_detalhe"):
            data[key] = []
        for key in ("tempos_operacao", "custos_operacao", "quote_cost_snapshot"):
            data[key] = {}
        if not str(data.get("material_subtype", "") or "").strip():
            data["material_subtype"] = str(data.get("calc_mode", "") or "Stock MP").strip()
        return data

    def _quote_line_type_label(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            return "Matéria-prima"
        return str(self.services.orc_line_type_label(self._quote_line_type(row)) or "-")

    def _quote_line_primary_ref(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            stock_id = str((row or {}).get("stock_material_id", "") or "").strip()
            return stock_id or str((row or {}).get("ref_externa", "") or "").strip() or "-"
        line_type = self._quote_line_type(row)
        if line_type == self.services.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        return str((row or {}).get("ref_interna", "") or "").strip() or "-"

    def _quote_line_material_display(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            return (
                str((row or {}).get("material_family", "") or "").strip()
                or str((row or {}).get("material", "") or "").strip()
                or "-"
            )
        line_type = self._quote_line_type(row)
        if line_type == self.services.ORC_LINE_TYPE_PRODUCT:
            return str((row or {}).get("produto_codigo", "") or "").strip() or "-"
        if line_type == self.services.ORC_LINE_TYPE_SERVICE:
            return "-"
        return str((row or {}).get("material", "") or "").strip() or "-"

    def _quote_line_unit_display(self, row: dict) -> str:
        if self._quote_line_is_raw_material_ui(row):
            subtype = str((row or {}).get("material_subtype", "") or "").strip()
            esp = str((row or {}).get("espessura", "") or "").strip()
            dimensao = str((row or {}).get("dimensao", (row or {}).get("dimensoes", "")) or "").strip()
            if dimensao:
                return dimensao
            if subtype and esp:
                return f"{subtype} | {esp}"
            return esp or subtype or "-"
        line_type = self._quote_line_type(row)
        if line_type == self.services.ORC_LINE_TYPE_PIECE:
            return str((row or {}).get("espessura", "") or "").strip() or "-"
        return str((row or {}).get("produto_unid", "") or "").strip() or "-"

    def _refresh_quote_identity_label(self) -> None:
        if not hasattr(self.view, "number_label"):
            return
        number = str(getattr(self, "current_number", "") or "").strip() or "Novo orçamento"
        client_name = ""
        if hasattr(self.view, "client_name_edit"):
            client_name = str(self.view.client_name_edit.text() or "").strip()
        if not client_name and hasattr(self.view, "client_combo"):
            client_row = self._client_lookup(self.view.client_combo.currentText())
            client_name = str(client_row.get("nome", "") or "").strip()
        self.view.number_label.setText(f"{number}  ·  {client_name}" if client_name else number)

    def _set_quote_header(self, numero: str, estado: str, encomenda: str = "") -> None:
        self.current_number = str(numero or "").strip()
        self._refresh_quote_identity_label()
        self.view.link_order_label.setText(f"Encomenda gerada: {encomenda}" if encomenda else "Sem encomenda gerada")
        _apply_state_chip(self.view.state_chip, estado)
        _set_panel_tone(self.view.quote_header_card, _state_tone(estado))
        _set_panel_tone(self.view.quote_summary_card, _state_tone(estado))

    def _clear_quote_detail(self) -> None:
        self.line_rows = []
        self.discount_group_keys = []
        self.nesting_bridge_data = {}
        self._set_quote_header(self.services.orc_next_number(), "Em edicao", "")
        self.view.client_combo.setCurrentText("")
        self.view.client_name_edit.clear()
        self.view.client_company_edit.clear()
        self.view.client_nif_edit.clear()
        self.view.client_address_edit.clear()
        self.view.client_contact_edit.clear()
        self.view.client_email_edit.clear()
        self.view.note_cliente_edit.clear()
        self.view.transport_combo.setCurrentText("")
        self.view.transport_carrier_combo.setCurrentText("")
        self.view.transport_zone_combo.setCurrentText("")
        self.view.transport_price_spin.setValue(0.0)
        self.view.discount_spin.setValue(0.0)
        self.view.price_increment_spin.setValue(0.0)
        self.view.discount_mode_combo.setCurrentIndex(0)
        self.view.transport_km_spin.setValue(0.0)
        self.view.transport_rate_spin.setValue(0.65)
        self.view.transport_diesel_spin.setValue(1.65)
        self.view.transport_consumption_spin.setValue(8.5)
        self.view.transport_trip_factor_spin.setValue(2.0)
        user = dict(self.services.current_user() or {})
        role = str(user.get("role", "") or "").strip().casefold()
        username = str(user.get("username", "") or "").strip()
        self.view.executed_combo.setCurrentText(username if username and role in {"orcamentista", "orçamentista"} else "")
        self.view.workcenter_combo.setCurrentText("")
        self.view.notes_edit.clear()
        self.view.delivery_date_min = QDate.currentDate()
        self.view.delivery_date_edit.setMinimumDate(self.view.delivery_date_min)
        self.view.delivery_date_edit.setDate(self.view.delivery_date_min)
        self.view.iva_spin.setValue(23.0)
        self._recalc_transport_calc()
        self._refresh_nesting_bridge()
        self._render_quote_lines()

    def _refresh_nesting_bridge(self) -> None:
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        if not bridge:
            self.view.nesting_bridge_label.setText(
                "Preco unitario da tabela = orcamentacao por peca. Nesting = validacao global de materia, stock/retalho e pecas realmente programadas."
            )
            self.view.nesting_bridge_label.setToolTip(
                "O preco unitario das linhas continua comercial e unitario. O nesting serve para validar consumo real de materia, stock, compra e pecas efetivamente programadas."
            )
            return
        placed = int(bridge.get("part_count_placed", 0) or 0)
        requested = int(bridge.get("part_count_requested", 0) or 0)
        sheets = int(bridge.get("sheet_count", 0) or 0)
        material_cost = _fmt_eur(float(bridge.get("material_net_cost_eur", 0) or 0))
        purchase_cost = _fmt_eur(float(bridge.get("material_purchase_requirement_eur", 0) or 0))
        quoted_total = _fmt_eur(float(bridge.get("quoted_total_eur", 0) or 0))
        rateio_total = _fmt_eur(float(bridge.get("rateio_adjusted_quote_total_eur", 0) or 0))
        rateio_delta = _fmt_eur(float(bridge.get("rateio_delta_eur", 0) or 0))
        method = str(bridge.get("analysis_method", "-") or "-").strip() or "-"
        profile_name = str(bridge.get("selected_profile_name", "") or "").strip() or "Apenas stock"
        self.view.nesting_bridge_label.setText(
            f"Nesting atual: {placed}/{requested} programadas | {sheets} chapa(s) | {method} | "
            f"perfil {profile_name} | materia real {material_cost} | compra {purchase_cost} | "
            f"comercial atual {quoted_total} | comercial ajustado ao rateio {rateio_total} | delta {rateio_delta}."
        )
        self.view.nesting_bridge_label.setToolTip(
            "Tabela do orcamento: preco comercial unitario por referencia.\n"
            "Nesting: custo global real de materia e validacao do lote completo.\n"
            "O rateio usa a area liquida colocada para repartir a materia real e a compra necessaria por referencia.\n"
            "Usa a tabela para vender/orcamentar por referencia e o plano de chapa para validar consumo real, compra e cobertura comercial."
        )

    def _quote_discount_mode(self) -> str:
        mode = str(self.view.discount_mode_combo.currentData() or "").strip().lower()
        return mode if mode in {"total", "lotes_espessura"} else "total"

    def _quote_discount_groups(self) -> list[dict]:
        groups: list[dict] = []
        by_key: dict[str, dict] = {}
        for row_index, row in enumerate(list(self.line_rows or [])):
            line_total = round(float(row.get("total", 0) or (float(row.get("qtd", 0) or 0) * float(row.get("preco_unit", 0) or 0))), 2)
            if line_total <= 0:
                continue
            key, label = self._quote_discount_group_label(row)
            bucket = by_key.setdefault(key, {"key": key, "label": label, "base": 0.0, "rows": []})
            bucket["base"] = round(float(bucket.get("base", 0) or 0) + line_total, 2)
            bucket["rows"].append(row_index)
        groups.extend(by_key.values())
        return groups

    def _quote_selected_discount_group_keys(self) -> set[str]:
        groups = self._quote_discount_groups()
        available = {str(group.get("key", "") or "").strip() for group in groups if str(group.get("key", "") or "").strip()}
        mode = self._quote_discount_mode()
        selected = {
            str(key or "").strip()
            for key in list(getattr(self, "discount_group_keys", []) or [])
            if str(key or "").strip() in available
        }
        if mode == "lotes_espessura":
            return selected or available
        return available

    def _quote_effective_discount_group_keys(self) -> list[str]:
        return sorted(self._quote_selected_discount_group_keys())

    def _sync_discount_targets_label(self) -> None:
        mode = self._quote_discount_mode()
        self.view.discount_groups_btn.setEnabled(mode == "lotes_espessura")
        if mode != "lotes_espessura":
            self.view.discount_target_label.setText("Desconto aplicado a todas as linhas do orçamento. O transporte fica fora do desconto.")
            return
        groups = self._quote_discount_groups()
        selected = self._quote_selected_discount_group_keys()
        if not groups:
            self.view.discount_target_label.setText("Sem lotes elegíveis para desconto neste orçamento.")
            return
        if not selected:
            self.view.discount_target_label.setText("Sem lotes selecionados manualmente. A pré-visualização aplica a todos os lotes elegíveis.")
            return
        labels = [str(group.get("label", "") or "").strip() for group in groups if str(group.get("key", "") or "").strip() in selected]
        preview = ", ".join(labels[:3])
        if len(labels) > 3:
            preview = f"{preview}, ..."
        self.view.discount_target_label.setText(f"Desconto aplicado a: {preview}")

    def _pick_quote_discount_groups(self) -> None:
        groups = self._quote_discount_groups()
        if not groups:
            QMessageBox.information(self, "Desconto por lote", "Este orçamento ainda não tem linhas elegíveis para selecionar lotes.")
            return
        selected = self._quote_selected_discount_group_keys()
        dialog = QDialog(self)
        dialog.setWindowTitle("Selecionar lotes / espessuras para desconto")
        dialog.resize(560, 420)
        root = QVBoxLayout(dialog)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)
        helper = QLabel("Marca os lotes/espessuras onde queres aplicar o desconto global.")
        helper.setWordWrap(True)
        root.addWidget(helper)
        list_widget = QListWidget()
        for group in groups:
            label = f"{group.get('label', '-')}: {_fmt_eur(float(group.get('base', 0) or 0))}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, str(group.get("key", "") or "").strip())
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if str(group.get("key", "") or "").strip() in selected else Qt.Unchecked)
            list_widget.addItem(item)
        root.addWidget(list_widget, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Aplicar")
        buttons.button(QDialogButtonBox.Cancel).setText("Fechar")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        root.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        chosen: list[str] = []
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item is not None and item.checkState() == Qt.Checked:
                key = str(item.data(Qt.UserRole) or "").strip()
                if key:
                    chosen.append(key)
        self.discount_group_keys = chosen
        self._render_quote_lines()

    def _quote_discount_group_label(self, row: dict) -> tuple[str, str]:
        if self._quote_line_is_raw_material_ui(row) or self._quote_line_type(row) == self.services.ORC_LINE_TYPE_PIECE:
            family = str(row.get("material_family", "") or row.get("material", "") or "Material").strip() or "Material"
            subtype = str(row.get("material_subtype", "") or "").strip()
            thickness = str(row.get("espessura", "") or "").strip()
            label = family
            if subtype:
                label = f"{label} / {subtype}"
            if thickness:
                label = f"{label} / {thickness} mm"
            key = f"piece|{family}|{subtype}|{thickness}"
            return key, label
        line_type = self._quote_line_type(row)
        if line_type == self.services.ORC_LINE_TYPE_PRODUCT:
            return "product", "Produtos / outros artigos"
        return "service", "Servicos / outras operacoes"

    def _quote_discount_breakdown(self, selected_keys: set[str], discount_pct: float) -> list[dict]:
        groups: list[dict] = []
        if discount_pct <= 0:
            return groups
        for group in self._quote_discount_groups():
            key = str(group.get("key", "") or "").strip()
            base_value = round(float(group.get("base", 0) or 0), 2)
            if not key or base_value <= 0:
                continue
            discount_value = round(base_value * (discount_pct / 100.0), 2) if key in selected_keys else 0.0
            groups.append(
                {
                    "key": key,
                    "label": str(group.get("label", "") or "").strip() or "-",
                    "base": base_value,
                    "discount": discount_value,
                    "subtotal_after_discount": round(max(0.0, base_value - discount_value), 2),
                }
            )
        return groups

    def _render_quote_lines(self) -> None:
        checked_source_indexes = (
            set(self._checked_line_indexes())
            if hasattr(self.view, "lines_table") and not bool(getattr(self, "_rendering_quote_lines", False))
            else set()
        )
        subtotal = 0.0
        subtotal_base = 0.0
        bridge = dict(getattr(self, "nesting_bridge_data", {}) or {})
        bridge_rows = {
            str(row.get("ref_externa", "") or "").strip(): dict(row or {})
            for row in list(bridge.get("part_rows", []) or [])
            if str(row.get("ref_externa", "") or "").strip()
        }
        discount_pct = float(self.view.discount_spin.value() or 0)
        increment_pct = float(self.view.price_increment_spin.value() or 0)
        selected_discount_keys = self._quote_selected_discount_group_keys()
        normalized_rows = []
        for row in self.line_rows:
            if self._quote_line_should_be_raw_material_ui(row):
                row = self._coerce_quote_line_raw_material_ui(row)
            qtd = float(row.get("qtd", 0) or 0)
            preco_unit = float(row.get("preco_unit", 0) or 0)
            subtotal_base = round(subtotal_base + (qtd * preco_unit), 2)
            preco_unit_incrementado = round(max(0.0, preco_unit * (1.0 + (increment_pct / 100.0))), 4)
            total = round(qtd * preco_unit_incrementado, 2)
            row["preco_unit_incrementado"] = preco_unit_incrementado
            row["total"] = total
            group_key, _group_label = self._quote_discount_group_label(row)
            apply_discount = discount_pct > 0 and group_key in selected_discount_keys
            discounted_unit = round(preco_unit_incrementado * (1.0 - (discount_pct / 100.0)), 4) if apply_discount else round(preco_unit_incrementado, 4)
            discounted_total = round(qtd * discounted_unit, 2)
            row["discount_group_key"] = group_key
            row["preco_unit_desconto"] = discounted_unit
            row["total_desconto"] = discounted_total
            row["desconto_aplicado"] = round(max(0.0, total - discounted_total), 2)
            normalized_rows.append(row)
        self.line_rows = normalized_rows
        display_rows = list(enumerate(self.line_rows))
        sort_section = int(getattr(self, "_quote_lines_sort_section", -1))
        if sort_section >= 0:
            reverse = getattr(self, "_quote_lines_sort_order", Qt.AscendingOrder) == Qt.DescendingOrder
            display_rows.sort(
                key=lambda indexed_row: self._quote_line_sort_value(indexed_row[1], sort_section),
                reverse=reverse,
            )
        self._rendering_quote_lines = True
        previous_signal_state = self.view.lines_table.blockSignals(True)
        updates_were_enabled = self.view.lines_table.updatesEnabled()
        self.view.lines_table.setUpdatesEnabled(False)
        _fill_table(
            self.view.lines_table,
            [
                [
                    "",
                    self._quote_line_type_label(row),
                    self._quote_line_primary_ref(row),
                    row.get("ref_externa", "-") or "-",
                    row.get("descricao", "-") or "-",
                    self._quote_line_material_display(row),
                    self._quote_line_unit_display(row),
                    "-" if (self._quote_line_is_raw_material_ui(row) and not str(row.get("operacao", "") or "").strip()) or self._quote_line_type(row) == self.services.ORC_LINE_TYPE_PRODUCT else (row.get("operacao", "-") or "-"),
                    "-" if (self._quote_line_is_raw_material_ui(row) and not str(row.get("operacao", "") or "").strip()) or self._quote_line_type(row) == self.services.ORC_LINE_TYPE_PRODUCT else f"{float(row.get('tempo_peca_min', 0) or 0):.2f} min",
                    f"{float(row.get('qtd', 0) or 0):.2f}",
                    _fmt_eur(float(row.get("preco_unit_incrementado", row.get("preco_unit", 0)) or 0)),
                    _fmt_eur(float(row.get("preco_unit_desconto", row.get("preco_unit", 0)) or 0)),
                    _fmt_eur(float(row.get("total_desconto", row.get("total", 0)) or 0)),
                    row.get("conjunto_nome", "-") or "-",
                ]
                for _source_index, row in display_rows
            ],
            align_center_from=self.LINE_COL_UNIT,
        )
        for visual_row, (source_index, _row) in enumerate(display_rows):
            for col_index in range(self.view.lines_table.columnCount()):
                item = self.view.lines_table.item(visual_row, col_index)
                if item is not None:
                    item.setData(Qt.UserRole, source_index)
            mark_item = self.view.lines_table.item(visual_row, self.LINE_COL_MARK)
            if mark_item is not None:
                mark_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                mark_item.setCheckState(Qt.Checked if source_index in checked_source_indexes else Qt.Unchecked)
                mark_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                mark_item.setToolTip("Marca esta linha para a remover do orçamento.")
        for _ in range(2):
            spacer_row = self.view.lines_table.rowCount()
            self.view.lines_table.insertRow(spacer_row)
            self.view.lines_table.setRowHeight(spacer_row, 24)
            for col_index in range(self.view.lines_table.columnCount()):
                spacer_item = QTableWidgetItem("")
                spacer_item.setFlags(Qt.ItemIsEnabled)
                spacer_item.setBackground(QBrush(QColor("#f8fafc")))
                self.view.lines_table.setItem(spacer_row, col_index, spacer_item)
        transport = float(self.view.transport_price_spin.value() or 0)
        for row_index, (_source_index, row) in enumerate(display_rows):
            subtotal += float(row.get("total", 0) or 0)
            _paint_table_row(self.view.lines_table, row_index, "Preparacao")
            for col_index in (self.LINE_COL_PRICE, self.LINE_COL_DISCOUNTED_PRICE, self.LINE_COL_TOTAL):
                item = self.view.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            line_type = self._quote_line_type(row)
            ref_externa = str(row.get("ref_externa", "") or "").strip()
            bridge_row = bridge_rows.get(ref_externa, {})
            tooltip_lines = [
                f"Ref. externa: {ref_externa or '-'}",
                f"Descricao: {str(row.get('descricao', '') or '-').strip() or '-'}",
                f"Operacao: {str(row.get('operacao', '') or '-').strip() or '-'}",
                f"Quantidade na linha: {float(row.get('qtd', 0) or 0):.2f}",
                f"Preco unitario comercial: {_fmt_eur(float(row.get('preco_unit', 0) or 0))}",
                f"Preco unitario com desconto: {_fmt_eur(float(row.get('preco_unit_desconto', row.get('preco_unit', 0)) or 0))}",
            ]
            pdf_docs = [
                str(item or "").strip()
                for item in [row.get("desenho_pdf", ""), *list(row.get("desenhos_pdf", []) or [])]
                if str(item or "").strip()
            ]
            if pdf_docs:
                tooltip_lines.append(f"Documentacao PDF: {len(dict.fromkeys(pdf_docs))} ficheiro(s) associado(s).")
            else:
                tooltip_lines.append("Documentacao PDF: sem PDF associado.")
            if bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False)):
                tooltip_lines.append("Materia-prima: fornecida pelo cliente. A linha considera apenas transformacao/processo.")
            quote_snapshot = dict(row.get("quote_cost_snapshot", {}) or {})
            if line_type == self.services.ORC_LINE_TYPE_PIECE and str(quote_snapshot.get("costing_mode", "") or "").strip() == "aggregate_pending":
                tooltip_lines.append("Rota multi-operacao: o preco atual da linha ainda esta agregado e nao repartido por posto.")
            elif line_type == self.services.ORC_LINE_TYPE_PIECE and str(quote_snapshot.get("costing_mode", "") or "").strip() in {"detailed", "partial_detail"}:
                for op_row in list(row.get("operacoes_detalhe", []) or []):
                    if not isinstance(op_row, dict):
                        continue
                    op_name = str(op_row.get("nome", "") or "").strip()
                    if not op_name:
                        continue
                    tempo_txt = "-" if op_row.get("tempo_unit_min") in (None, "") else f"{float(op_row.get('tempo_unit_min', 0) or 0):.3f} min/un"
                    custo_txt = "-" if op_row.get("custo_unit_eur") in (None, "") else _fmt_eur(float(op_row.get("custo_unit_eur", 0) or 0))
                    tooltip_lines.append(f"{op_name}: {tempo_txt} | {custo_txt}")
            if line_type == self.services.ORC_LINE_TYPE_PIECE and str(row.get("desenho", "") or "").strip() and "corte laser" in str(row.get("operacao", "") or "").strip().lower():
                if bridge:
                    programmed = int(bridge_row.get("qty", 0) or 0)
                    requested = int(round(float(row.get("qtd", 0) or 0)))
                    tooltip_lines.append(f"Nesting: {programmed}/{requested} peca(s) programada(s) no ultimo cenario.")
                    if bridge_row:
                        tooltip_lines.append(f"Valor comercial colocado no plano: {_fmt_eur(float(bridge_row.get('quoted_total_eur', 0) or 0))}")
                        tooltip_lines.append(
                            f"Rateio matéria real: {_fmt_eur(float(bridge_row.get('allocated_material_total_eur', 0) or 0))} "
                            f"({_fmt_eur(float(bridge_row.get('allocated_material_unit_eur', 0) or 0))}/un)"
                        )
                        tooltip_lines.append(
                            f"Rateio compra: {_fmt_eur(float(bridge_row.get('allocated_purchase_total_eur', 0) or 0))} "
                            f"({_fmt_eur(float(bridge_row.get('allocated_purchase_unit_eur', 0) or 0))}/un)"
                        )
                        tooltip_lines.append(
                            f"Comercial atual: {_fmt_eur(float(bridge_row.get('current_quote_total_eur', 0) or 0))} | "
                            f"ajustado ao plano: {_fmt_eur(float(bridge_row.get('adjusted_quote_total_eur', 0) or 0))}"
                        )
                        tooltip_lines.append(
                            f"Peso no plano: {float(bridge_row.get('share_pct', 0) or 0):.2f}% | "
                            f"delta comercial: {_fmt_eur(float(bridge_row.get('quote_delta_eur', 0) or 0))}"
                        )
                    else:
                        tooltip_lines.append("Esta referencia nao ficou colocada no ultimo plano analisado.")
                else:
                    tooltip_lines.append("Sem plano de chapa aplicado a este orcamento nesta sessao.")
            tooltip = "\n".join(tooltip_lines)
            for col_index in range(self.view.lines_table.columnCount()):
                item = self.view.lines_table.item(row_index, col_index)
                if item is not None:
                    item.setToolTip(tooltip)
            if bridge_row:
                adjusted_total = float(bridge_row.get("adjusted_quote_total_eur", 0) or 0.0)
                current_total = float(row.get("total_desconto", row.get("total", 0)) or 0.0)
                total_item = self.view.lines_table.item(row_index, self.LINE_COL_TOTAL)
                if total_item is not None and adjusted_total > 0:
                    if current_total + 0.009 < adjusted_total:
                        total_item.setBackground(QBrush(QColor("#fff8eb")))
                        total_item.setForeground(QBrush(QColor("#b45f06")))
                    elif current_total > adjusted_total + 0.009:
                        total_item.setBackground(QBrush(QColor("#dcfce7")))
                        total_item.setForeground(QBrush(QColor("#166534")))
                    else:
                        total_item.setBackground(QBrush(QColor("#fef3c7")))
                        total_item.setForeground(QBrush(QColor("#92400e")))
        self.view.lines_table.blockSignals(previous_signal_state)
        self.view.lines_table.setUpdatesEnabled(updates_were_enabled)
        self._rendering_quote_lines = False
        discount_mode = self._quote_discount_mode()
        discount_value = round(sum(float(row.get("desconto_aplicado", 0) or 0) for row in self.line_rows), 2)
        subtotal_discounted = round(sum(float(row.get("total_desconto", row.get("total", 0)) or 0) for row in self.line_rows), 2)
        subtotal_without_iva = max(0.0, subtotal_discounted + transport)
        iva_rate = 23.0
        if abs(float(self.view.iva_spin.value() or 0) - iva_rate) > 0.0001:
            self.view.iva_spin.blockSignals(True)
            self.view.iva_spin.setValue(iva_rate)
            self.view.iva_spin.blockSignals(False)
        iva_amount = round(subtotal_without_iva * (iva_rate / 100.0), 2)
        total = round(subtotal_without_iva + iva_amount, 2)
        if hasattr(self.view, "line_count_label"):
            line_count = len(self.line_rows)
            self.view.line_count_label.setText(f"{line_count} linha" if line_count == 1 else f"{line_count} linhas")
        if hasattr(self.view, "quote_finance_status_label"):
            if not self.line_rows:
                finance_status = "AGUARDA LINHAS"
            elif discount_pct > 0 and increment_pct > 0:
                finance_status = "PREÇOS AJUSTADOS"
            elif discount_pct > 0:
                finance_status = f"DESCONTO {discount_pct:.1f}%"
            elif increment_pct != 0:
                finance_status = f"INCREMENTO {increment_pct:.1f}%"
            else:
                finance_status = "SEM AJUSTES"
            self.view.quote_finance_status_label.setText(finance_status.replace(".", ","))
        if hasattr(self.view, "total_caption_label"):
            self.view.total_caption_label.setText("Total C/IVA 23%")
        if getattr(self.view, "iva_summary_row_label", None) is not None:
            self.view.iva_summary_row_label.setText("IVA 23%")
        self.view.lines_subtotal_label.setText(_fmt_eur(subtotal))
        self.view.increment_value_label.setText(_fmt_eur(max(0.0, round(subtotal - subtotal_base, 2))))
        self.view.transport_total_label.setText(_fmt_eur(transport))
        self.view.discount_value_label.setText(_fmt_eur(discount_value))
        self.view.subtotal_without_iva_label.setText(_fmt_eur(subtotal_without_iva))
        self.view.iva_total_label.setText(_fmt_eur(iva_amount))
        self.view.total_label.setText(_fmt_eur(total))
        if hasattr(self.view, "header_total_label"):
            self.view.header_total_label.setText(_fmt_eur(total))
        self._sync_discount_targets_label()
        discount_breakdown = self._quote_discount_breakdown(selected_discount_keys, discount_pct)
        if discount_mode == "lotes_espessura" and discount_breakdown:
            breakdown_lines = [
                f"{str(group.get('label', '-') or '-').strip()}: -{_fmt_eur(float(group.get('discount', 0) or 0))}"
                for group in discount_breakdown
                if float(group.get("discount", 0) or 0) > 0
            ]
            if breakdown_lines:
                self.view.discount_breakdown_label.setText("Desconto aplicado por lote:\n" + "\n".join(breakdown_lines))
            else:
                self.view.discount_breakdown_label.setText("Nenhum dos lotes selecionados tem desconto aplicado.")
        elif discount_pct > 0:
            self.view.discount_breakdown_label.setText("Desconto aplicado a todas as linhas do orçamento. O transporte mantém-se fora do desconto.")
        else:
            self.view.discount_breakdown_label.setText("Sem desconto global aplicado.")
        # A tabela ocupa apenas o espaço disponível no cartão e faz o seu próprio
        # scroll. Assim, o rodapé de ações e o limite inferior do cartão ficam
        # sempre visíveis, independentemente do número de referências.
        self.view.lines_table.setMinimumHeight(280)
        self.view.lines_table.setMaximumHeight(16777215)
        if hasattr(self.view, "quote_lines_card"):
            self.view.quote_lines_card.setMinimumHeight(500)
        self._sync_quote_selected_line_actions()

    def _load_quote(self, numero: str) -> None:
        detail = self.services.orc_detail(numero)
        client = dict(detail.get("cliente", {}) or {})
        self._set_quote_header(detail.get("numero", ""), detail.get("estado", "Em edicao"), detail.get("numero_encomenda", ""))
        client_label = f"{client.get('codigo', '')} - {client.get('nome', '')}".strip(" -")
        self.view.client_combo.setCurrentText(client_label or client.get("nome", ""))
        self.view.client_name_edit.setText(str(client.get("nome", "") or "").strip())
        self.view.client_company_edit.setText(str(client.get("empresa", "") or client.get("nome", "") or "").strip())
        self.view.client_nif_edit.setText(str(client.get("nif", "") or "").strip())
        self.view.client_address_edit.setText(str(client.get("morada", "") or "").strip())
        self.view.client_contact_edit.setText(str(client.get("contacto", "") or "").strip())
        self.view.client_email_edit.setText(str(client.get("email", "") or "").strip())
        self._refresh_quote_identity_label()
        self.view.executed_combo.setCurrentText(str(detail.get("executado_por", "") or "").strip())
        self.view.workcenter_combo.setCurrentText("")
        self.view.note_cliente_edit.setText(str(detail.get("nota_cliente", "") or "").strip())
        self.view.transport_combo.setCurrentText(str(detail.get("nota_transporte", "") or "").strip())
        carrier_txt = " - ".join(
            [
                part
                for part in [
                    str(detail.get("transportadora_id", "") or "").strip(),
                    str(detail.get("transportadora_nome", "") or "").strip(),
                ]
                if part
            ]
        ).strip(" -")
        self.view.transport_carrier_combo.setCurrentText(carrier_txt)
        self.view.transport_zone_combo.setCurrentText(str(detail.get("zona_transporte", "") or "").strip())
        self.view.transport_price_spin.setValue(float(detail.get("preco_transporte", 0) or 0))
        self.view.discount_spin.setValue(float(detail.get("desconto_perc", 0) or 0))
        self.view.price_increment_spin.setValue(float(detail.get("incremento_preco_perc", 0) or 0))
        discount_mode = str(detail.get("desconto_modo", "total") or "total").strip().lower()
        self.view.discount_mode_combo.setCurrentIndex(1 if discount_mode == "lotes_espessura" else 0)
        self.discount_group_keys = [
            str(key or "").strip()
            for key in list(detail.get("desconto_grupos", []) or [])
            if str(key or "").strip()
        ]
        self.view.notes_edit.setPlainText(str(detail.get("notas_pdf", "") or "").strip())
        delivery_date = str(detail.get("prazo_entrega_data", "") or "").strip()[:10]
        qdate = QDate.fromString(delivery_date, "yyyy-MM-dd") if delivery_date else QDate()
        self.view.delivery_date_min = QDate.currentDate()
        self.view.delivery_date_edit.setMinimumDate(self.view.delivery_date_min)
        self.view.delivery_date_edit.setDate(qdate if qdate.isValid() and qdate >= self.view.delivery_date_min else self.view.delivery_date_min)
        self.view.iva_spin.setValue(23.0)
        self.nesting_bridge_data = dict(detail.get("nesting_bridge", {}) or {})
        self._refresh_nesting_bridge()
        self.line_rows = [dict(row) for row in list(detail.get("linhas", []) or [])]
        self._recalc_transport_calc()
        self._render_quote_lines()

    def _open_selected_quote(self) -> None:
        row = self._selected_quote_row()
        numero = str(row.get("numero", "") or "").strip()
        if not numero:
            QMessageBox.warning(self, "Orçamentos", "Seleciona um orcamento.")
            return
        try:
            self._load_quote(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._show_detail()

    def _new_quote(self) -> None:
        self._clear_quote_detail()
        self._show_detail()

    def _quote_pick_workcenter(self, *preferred_values: str) -> str:
        options = [str(self.view.workcenter_combo.itemText(index) or "").strip() for index in range(self.view.workcenter_combo.count())]
        normalized = {str(option).strip().lower(): option for option in options if str(option).strip()}
        for value in preferred_values:
            key = str(value or "").strip().lower()
            if key and key in normalized:
                return normalized[key]
        return str(options[0] if options else "").strip()

    def _structure_line(self, descricao: str, qtd: float, unid: str, preco_unit: float, operacao: str) -> dict | None:
        return service_line(descricao, qtd, unid, preco_unit, operacao, line_type=self.services.ORC_LINE_TYPE_SERVICE)

    def _structure_product_line(self, product: dict | None, qtd: float, *, descricao_extra: str = "") -> dict | None:
        return product_line(product, qtd, descricao_extra=descricao_extra, line_type=self.services.ORC_LINE_TYPE_PRODUCT)

    def _structure_quote_dialog(self) -> dict | None:
        return edit_structure(self, self.services.editor_ports, workcenter=self._quote_pick_workcenter("Serralharia", "Montagem"))

    def _apply_structure_quote_payload(self, payload: dict, *, replace_existing: bool) -> None:
        self._apply_group_payload(payload, replace_existing=replace_existing)

    def _apply_group_payload(self, payload: dict, *, replace_existing: bool) -> None:
        if not isinstance(payload, dict):
            return
        assembly_code = str(payload.get("assembly_code", "") or "").strip()
        assembly_name = str(payload.get("assembly_name", payload.get("name", "")) or "").strip()
        group_uuid = f"{assembly_code}-01" if assembly_code else f"EST-GRP-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        lines = []
        for raw_row in list(payload.get("lines", []) or []):
            if not isinstance(raw_row, dict):
                continue
            row = dict(raw_row)
            if assembly_code:
                row["conjunto_codigo"] = assembly_code
            if assembly_name:
                row["conjunto_nome"] = assembly_name
            row["grupo_uuid"] = group_uuid
            lines.append(row)
        if not lines:
            return
        if replace_existing:
            self.line_rows = lines
        else:
            self.line_rows.extend(lines)
        note_cliente = str(payload.get("note_cliente", "") or "").strip()
        if note_cliente and (replace_existing or not self.view.note_cliente_edit.text().strip()):
            self.view.note_cliente_edit.setText(note_cliente)
        notes_pdf = str(payload.get("notes_pdf", "") or "").strip()
        if notes_pdf:
            current_notes = [] if replace_existing else [row.strip() for row in self.view.notes_edit.toPlainText().splitlines() if row.strip()]
            for row in notes_pdf.splitlines():
                row_txt = str(row or "").strip()
                if row_txt and row_txt not in current_notes:
                    current_notes.append(row_txt)
            self.view.notes_edit.setPlainText("\n".join(current_notes))
        workcenter = str(payload.get("workcenter", "") or "").strip()
        if workcenter:
            self.view.workcenter_combo.setCurrentText(workcenter)
        self._render_quote_lines()
        self._show_detail()

    def _material_price_manager_dialog(self, formato_filter: str = "", preferred_id: str = "", parent: QWidget | None = None) -> dict | None:
        return edit_material_prices(self, self.services.editor_ports, formato_filter, preferred_id, parent)

    def _sync_stock_price_from_context(self, material_id: str, price_kg: float, current_price_kg: float, parent: QWidget | None = None) -> dict[str, Any] | None:
        return sync_stock_price(self, self.services.editor_ports, material_id, price_kg, current_price_kg, parent)

    def _material_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        return edit_material(self, self.services.editor_ports, initial, parent, presets=self.presets)

    def _labor_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        return edit_labor(self, self.services.editor_ports, initial, parent, presets=self.presets)

    def _consumable_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        return edit_consumable(self, self.services.editor_ports, initial, parent)

    def _product_assembly_item_dialog(self, initial: dict | None = None, parent: QWidget | None = None) -> dict | None:
        return edit_product(self, self.services.editor_ports, initial, parent)

    def _assembly_item_kind_label(self, kind: str) -> str:
        return {
            "material": "Material",
            "labor": "Mao de obra",
            "consumable": "Consumivel",
            "product": "Produto stock",
        }.get(str(kind or "").strip(), "Item")

    def _assembly_item_kind_from_line(self, row: dict | None = None) -> str:
        payload = dict(row or {})
        explicit = str(payload.get("kind", "") or "").strip()
        if explicit in {"material", "labor", "consumable", "product"}:
            return explicit
        line = dict(payload.get("line") or payload)
        if self.services.orc_line_is_product(line):
            return "product"
        if self.services.orc_line_is_piece(line):
            return "material"
        unit_txt = str(line.get("produto_unid", "") or "").strip().lower()
        if unit_txt == "h":
            return "labor"
        return "consumable"

    def _wrap_assembly_item(self, row: dict | None = None) -> dict:
        payload = dict(row or {})
        if isinstance(payload.get("line"), dict):
            line = dict(payload.get("line") or {})
            total_cost = float(payload.get("total_cost", line.get("qtd", 0)) or 0) * 1.0
            if float(payload.get("total_cost", 0) or 0) <= 0:
                total_cost = round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2)
            payload["kind"] = self._assembly_item_kind_from_line(payload)
            payload["total_cost"] = round(total_cost, 2)
            return payload
        line = payload
        return {
            "kind": self._assembly_item_kind_from_line(line),
            "descricao_base": str(line.get("descricao", "") or "").strip(),
            "total_cost": round(float(line.get("qtd", 0) or 0) * float(line.get("preco_unit", 0) or 0), 2),
            "weight_total": 0.0,
            "line": dict(line),
        }

    def _pick_assembly_item_kind(self, parent: QWidget | None = None) -> str:
        labels = ["Material", "Mao de obra", "Consumivel", "Produto stock"]
        selected, ok = QInputDialog.getItem(
            parent if isinstance(parent, QWidget) else self,
            "Adicionar item ao conjunto",
            "Tipo de item",
            labels,
            0,
            False,
        )
        if not ok:
            return ""
        return {
            "Material": "material",
            "Mao de obra": "labor",
            "Consumivel": "consumable",
            "Produto stock": "product",
        }.get(str(selected or "").strip(), "")

    def _open_assembly_item_editor(self, kind: str, current: dict | None = None, parent: QWidget | None = None) -> dict | None:
        if kind == "material":
            return self._material_assembly_item_dialog(current, parent=parent)
        if kind == "labor":
            return self._labor_assembly_item_dialog(current, parent=parent)
        if kind == "consumable":
            return self._consumable_assembly_item_dialog(current, parent=parent)
        if kind == "product":
            return self._product_assembly_item_dialog(current, parent=parent)
        return None

    def _calculated_assembly_builder_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Conjunto calculado")
        dialog.resize(1040, 840)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        intro = QLabel(
            "Cria um conjunto como mini-projeto: materiais, mao de obra, consumiveis e produtos. "
            "Cada conjunto criado aqui fica guardado no backend; se ativares a opcao abaixo, fica tambem marcado como template reutilizavel."
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
            suggested_param = str(self.services.conjunto_next_param_codigo() or "").strip()
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

        items: list[dict] = [self._wrap_assembly_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
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
                return self._material_assembly_item_dialog(current, parent=dialog)
            if kind == "labor":
                return self._labor_assembly_item_dialog(current, parent=dialog)
            if kind == "consumable":
                return self._consumable_assembly_item_dialog(current, parent=dialog)
            if kind == "product":
                return self._product_assembly_item_dialog(current, parent=dialog)
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
            return self._assembly_item_kind_from_line(line)

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
            batch_dialog = self.services.laser_batch_dialog(dialog,
                default_machine=self.view.workcenter_combo.currentText().strip(),
            )
            if batch_dialog.exec() != QDialog.Accepted:
                return
            result = dict(batch_dialog.result_payload() or {})
            _add_lines_from_shortcut(
                [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})],
                laser_batch_id=str(batch_dialog.batch_id or "").strip(),
            )

        def _add_step_igs_to_assembly() -> None:
            lines = self._open_profile_step_igs_quote_builder(return_lines=True, parent=dialog)
            _add_lines_from_shortcut([dict(row or {}) for row in list(lines or []) if dict(row or {})])

        def _add_structure_to_assembly() -> None:
            payload = self._structure_quote_dialog()
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
                if self._quote_line_is_laser_2d(current_line):
                    batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
                    batch_indexes = [
                        row_index
                        for row_index, candidate in enumerate(items)
                        if batch_id
                        and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
                    ] or [index]
                    source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
                    edited_lines = self._edit_laser_batch_lines(source_lines, parent=dialog)
                    if edited_lines is None:
                        return
                    insert_at = min(batch_indexes)
                    for row_index in sorted(batch_indexes, reverse=True):
                        del items[row_index]
                    for offset, edited_line in enumerate(edited_lines):
                        items.insert(insert_at + offset, self._wrap_assembly_item(edited_line))
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
            QMessageBox.warning(self, "Conjunto", "O conjunto precisa de pelo menos um item.")
            return None
        assembly_code = code_edit.text().strip() or f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        assembly_name = name_edit.text().strip() or assembly_code
        lines = [dict(item.get("line") or {}) for item in items if isinstance(item.get("line"), dict)]
        for line in lines:
            operation_norm = self.services.norm_text(str(line.get("operacao", "") or ""))
            if self.services.orc_line_is_piece(line) and "laser" in operation_norm:
                line["source_quote_number"] = str(self.current_number or "").strip()
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
            self.services.assembly_model_save(
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
            self.services.conjunto_save(
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
            QMessageBox.critical(self, "Conjunto", str(exc))
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
            "workcenter": self._quote_pick_workcenter("Serralharia", "Montagem"),
            "lines": lines,
        }

    def _open_calculated_assembly_builder(self) -> None:
        payload = self._calculated_assembly_builder_dialog()
        if not payload:
            return
        self._apply_group_payload(payload, replace_existing=False)

    def _open_structure_quote_builder(self) -> None:
        if self.line_rows:
            if QMessageBox.question(
                self,
                "Orcamento Estruturas",
                "Substituir as linhas atuais pelo modelo de estruturas metalicas?",
            ) != QMessageBox.Yes:
                return
            replace_existing = True
        else:
            replace_existing = True
        payload = self._structure_quote_dialog()
        if not payload:
            return
        self._apply_structure_quote_payload(payload, replace_existing=replace_existing)

    def _new_structure_quote(self) -> None:
        self._clear_quote_detail()
        payload = self._structure_quote_dialog()
        if not payload:
            self._show_list()
            return
        self._apply_structure_quote_payload(payload, replace_existing=True)

    def _transport_calc_value(self) -> float:
        kms = float(self.view.transport_km_spin.value() or 0)
        factor = float(self.view.transport_trip_factor_spin.value() or 1)
        rate = float(self.view.transport_rate_spin.value() or 0)
        diesel_price = float(self.view.transport_diesel_spin.value() or 0)
        consumption = float(self.view.transport_consumption_spin.value() or 0)
        route_cost = kms * factor * rate
        diesel_cost = kms * factor * (consumption / 100.0) * diesel_price
        return round(route_cost + diesel_cost, 2)

    def _recalc_transport_calc(self) -> None:
        self.view.transport_suggest_label.setText(f"Transporte sugerido: {_fmt_eur(self._transport_calc_value())}")

    def _apply_transport_calc(self) -> None:
        self.view.transport_price_spin.setValue(self._transport_calc_value())
        self._render_quote_lines()

    def _clear_transport(self) -> None:
        self.view.transport_combo.setCurrentText("")
        self.view.transport_carrier_combo.setCurrentText("")
        self.view.transport_zone_combo.setCurrentText("")
        self.view.transport_price_spin.setValue(0.0)
        self.view.transport_km_spin.setValue(0.0)
        self.view.transport_rate_spin.setValue(0.0)
        self.view.transport_diesel_spin.setValue(0.0)
        self.view.transport_consumption_spin.setValue(0.0)
        self.view.transport_trip_factor_spin.setValue(1.0)
        self._recalc_transport_calc()
        self._render_quote_lines()

    def _append_pdf_note(self, line: str) -> None:
        note = str(line or "").strip()
        if not note:
            return
        current = [item.strip() for item in self.view.notes_edit.toPlainText().splitlines() if item.strip()]
        if note not in current:
            current.append(note)
            self.view.notes_edit.setPlainText("\n".join(current))

    def _delivery_date_text(self) -> str:
        if self.view.delivery_date_edit.date() <= self.view.delivery_date_min:
            return ""
        return self.view.delivery_date_edit.date().toString("yyyy-MM-dd")

    def _delivery_text(self) -> str:
        return self.view.delivery_default_text

    def _apply_default_delivery_deadline(self) -> None:
        self.view.delivery_date_min = QDate.currentDate()
        self.view.delivery_date_edit.setMinimumDate(self.view.delivery_date_min)
        self.view.delivery_date_edit.setDate(self.view.delivery_date_min)
        self._append_pdf_note(f"- Prazo de entrega: {self.view.delivery_default_text}")

    def _fill_pdf_notes_from_context(self) -> None:
        try:
            text = self.services.orc_suggest_notes(self._quote_payload())
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        lines = [item.strip() for item in str(text or "").splitlines() if item.strip()]
        current = [item.strip() for item in self.view.notes_edit.toPlainText().splitlines() if item.strip()]
        keep = [item for item in current if "foi considerado" not in item.lower()]
        merged = []
        for item in lines + keep:
            if item and item not in merged:
                merged.append(item)
        self.view.notes_edit.setPlainText("\n".join(merged))

    def _quote_payload(self) -> dict:
        self._render_quote_lines()
        return {
            "numero": self.current_number,
            "estado": self.view.state_chip.text().strip() or "Em edição",
            "cliente": {
                "codigo": self._client_code_from_text(self.view.client_combo.currentText()),
                "nome": self.view.client_name_edit.text().strip(),
                "empresa": self.view.client_company_edit.text().strip(),
                "nif": self.view.client_nif_edit.text().strip(),
                "morada": self.view.client_address_edit.text().strip(),
                "contacto": self.view.client_contact_edit.text().strip(),
                "email": self.view.client_email_edit.text().strip(),
            },
            "executado_por": self.view.executed_combo.currentText().strip(),
            "nota_transporte": self.view.transport_combo.currentText().strip(),
            "prazo_entrega_texto": self._delivery_text(),
            "prazo_entrega_data": self._delivery_date_text(),
            "transportadora_nome": self.view.transport_carrier_combo.currentText().strip(),
            "zona_transporte": self.view.transport_zone_combo.currentText().strip(),
            "preco_transporte": self.view.transport_price_spin.value(),
            "incremento_preco_perc": self.view.price_increment_spin.value(),
            "desconto_perc": self.view.discount_spin.value(),
            "desconto_modo": self._quote_discount_mode(),
            "desconto_grupos": self._quote_effective_discount_group_keys(),
            "notas_pdf": self.view.notes_edit.toPlainText().strip(),
            "nota_cliente": self.view.note_cliente_edit.text().strip(),
            "iva_perc": 23.0,
            "linhas": self.line_rows,
        }

    def _save_quote(self) -> None:
        self._set_quote_save_button_state("saving")
        QApplication.processEvents()
        try:
            detail = self.services.orc_save(self._quote_payload())
        except Exception as exc:
            self._set_quote_save_button_state("error")
            QMessageBox.critical(self, "Guardar Orcamento", str(exc))
            self.view._quote_save_feedback_timer.start(2200)
            return
        self._load_quote(str(detail.get("numero", "") or "").strip())
        self.refresh()
        self._show_detail()
        self._set_quote_save_button_state("saved")
        self.view._quote_save_feedback_timer.start(1800)

    def _create_quote_purchase_note(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Nota de encomenda", "Guarda primeiro o orcamento.")
            return
        try:
            detail = self.services.orc_save(self._quote_payload())
            result = self.services.orc_create_purchase_quote(str(detail.get("numero", "") or self.current_number).strip())
        except Exception as exc:
            QMessageBox.critical(self, "Nota de encomenda", str(exc))
            return
        note_number = str(result.get("numero", "") or "").strip()
        line_count = int(result.get("line_count", 0) or 0)
        QMessageBox.information(self, "Nota de encomenda", f"Nota {note_number} criada com {line_count} linha(s) para pedido de cotacao.")
        main_window = self.window()
        if hasattr(main_window, "show_page"):
            try:
                main_window.show_page("purchase_notes")
                page = getattr(main_window, "pages", {}).get("purchase_notes")
                if page is not None and hasattr(page, "refresh"):
                    page.refresh()
                if page is not None and hasattr(page, "_load_note"):
                    page._load_note(note_number)
            except Exception:
                pass

    def _set_quote_save_button_state(self, state: str = "idle") -> None:
        buttons = [
            button
            for button in (
                getattr(self.view, "quote_save_btn", None),
                getattr(self.view, "quote_inspector_save_btn", None),
            )
            if isinstance(button, QPushButton)
        ]
        if not buttons:
            return
        current_state = str(state or "idle").strip().lower()
        for button in buttons:
            if current_state == "saving":
                button.setText("A guardar...")
                button.setEnabled(False)
                button.setProperty("variant", "secondary")
            elif current_state == "saved":
                button.setText("Guardado")
                button.setEnabled(True)
                button.setProperty("variant", "success")
            elif current_state == "error":
                button.setText("Falhou")
                button.setEnabled(True)
                button.setProperty("variant", "danger")
            else:
                button.setText("Guardar")
                button.setEnabled(True)
                button.setProperty("variant", "warning")
            _repolish(button)

    def _reset_quote_save_button(self) -> None:
        self._set_quote_save_button_state("idle")

    def _quote_email_subject(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.view.note_cliente_edit.text().strip()).strip()
        client = dict(payload.get("cliente", {}) or {})
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        subject = f"Proposta {client_name} [{numero}]" if numero else "Proposta luGEST"
        if rfq_ref:
            subject += f" | Pedido de Cotação {rfq_ref}"
        return subject

    def _quote_email_body(self, detail: dict[str, object] | None = None) -> str:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.view.note_cliente_edit.text().strip()).strip()
        lines = [
            f"Exmos. Senhores {client_name},",
            "",
            f"Segue em anexo o orçamento {numero}." if numero else "Segue em anexo o orçamento solicitado.",
        ]
        if rfq_ref:
            lines.append(f"Referência do pedido de cotação: {rfq_ref}.")
        lines.extend(
            [
                "",
                "Ficamos ao dispor para qualquer esclarecimento.",
                "",
                "Cumprimentos,",
                "luGEST",
            ]
        )
        return "\n".join(lines)

    def _quote_email_html_body(self, detail: dict[str, object] | None = None, *, logo_cid: str = "") -> str:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        branding = dict(self.services.branding() or {})
        company_name = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
        client_name = (
            str(client.get("empresa", "") or "").strip()
            or str(client.get("nome", "") or "").strip()
            or "Cliente"
        )
        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        rfq_ref = str(payload.get("nota_cliente", "") or self.view.note_cliente_edit.text().strip()).strip()
        subtotal = _fmt_eur(float(payload.get("subtotal", 0) or 0))
        client_code = str(client.get("codigo", "") or "").strip()
        payment_terms = "Conforme acordado"
        try:
            client_rows = list(self.services.client_rows("") or [])
            match = next(
                (
                    row
                    for row in client_rows
                    if (
                        client_code
                        and str(row.get("codigo", "") or "").strip() == client_code
                    )
                    or (
                        not client_code
                        and str(row.get("email", "") or "").strip().lower() == str(client.get("email", "") or "").strip().lower()
                    )
                ),
                None,
            )
            if isinstance(match, dict):
                payment_terms = str(match.get("cond_pagamento", "") or "").strip() or payment_terms
        except Exception:
            pass
        header_logo = ""
        if logo_cid:
            header_logo = (
                f"<img src=\"cid:{html.escape(logo_cid)}\" alt=\"{html.escape(company_name)}\" "
                "style=\"max-width:110px; max-height:36px; display:block;\" />"
            )
        else:
            header_logo = (
                f"<div style=\"font-size:34px; font-weight:800; letter-spacing:-1px; color:#ffffff;\">"
                f"{html.escape(company_name)}</div>"
            )
        reference_block = html.escape(rfq_ref or numero or "-")
        return (
            "<html><body style=\"margin:0; padding:0; background:#eef2f7; font-family:Segoe UI, Arial, sans-serif; color:#334155;\">"
            "<div style=\"padding:28px 0;\">"
            "<div style=\"width:620px; margin:0 auto; background:#ffffff; border-radius:18px; overflow:hidden; box-shadow:0 16px 40px rgba(15,23,42,0.10);\">"
            "<div style=\"background:#1f2933; padding:28px 34px;\">"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;\">"
            "<tr>"
            f"<td style=\"vertical-align:middle;\">{header_logo}</td>"
            "<td style=\"vertical-align:top; text-align:right; color:#cbd5e1; font-size:11px; letter-spacing:0.6px;\">"
            "REFERÊNCIA<br>"
            f"<span style=\"display:inline-block; margin-top:8px; font-size:23px; font-weight:800; color:#ffffff;\">{html.escape(reference_block)}</span>"
            "</td>"
            "</tr>"
            "</table>"
            "</div>"
            "<div style=\"padding:34px 36px 26px 36px;\">"
            f"<p style=\"margin:0 0 18px 0; font-size:22px; font-weight:800; color:#0f172a;\">Exmo(a). {html.escape(client_name)},</p>"
            "<p style=\"margin:0 0 16px 0; font-size:16px; line-height:1.7;\">Boa tarde,</p>"
            "<p style=\"margin:0 0 24px 0; font-size:16px; line-height:1.7;\">"
            "Conforme solicitado, segue o orçamento em anexo."
            "</p>"
            "<div style=\"margin:28px 0 12px 0; font-size:13px; font-weight:800; color:#94a3b8; letter-spacing:0.6px; text-transform:uppercase;\">"
            "Condições da proposta"
            "</div>"
            "<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse; border:1px solid #e2e8f0;\">"
            "<tr style=\"background:#1f2933; color:#ffffff; font-size:12px; font-weight:800; text-transform:uppercase;\">"
            "<td style=\"padding:12px 16px;\">Campo</td>"
            "<td style=\"padding:12px 16px;\">Valor</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Valor Total (s/ IVA)</td>"
            f"<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:30px; font-weight:800; color:#16a34a;\">{html.escape(subtotal)}</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Validade da proposta</td>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#0f172a;\">5 dias úteis, salvo rutura de stock.</td>"
            "</tr>"
            "<tr>"
            "<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#475569;\">Condições de pagamento</td>"
            f"<td style=\"padding:14px 16px; border-top:1px solid #e2e8f0; font-size:15px; color:#0f172a;\">{html.escape(payment_terms)}</td>"
            "</tr>"
            "</table>"
            "<p style=\"margin:28px 0 0 0; font-size:15px; line-height:1.7;\">"
            "Ficamos ao dispor para qualquer esclarecimento."
            "</p>"
            "</div>"
            f"<div style=\"padding:16px 36px; background:#f8fafc; border-top:1px solid #e2e8f0; font-size:12px; color:#94a3b8; text-align:center;\">© {datetime.now().year} {html.escape(company_name)}</div>"
            "</div>"
            "</div>"
            "</body></html>"
        )

    def _open_quote_email_draft(self, detail: dict[str, object] | None = None) -> None:
        payload = dict(detail or {})
        client = dict(payload.get("cliente", {}) or {})
        recipient = str(client.get("email", "") or self.view.client_email_edit.text().strip()).strip()
        if not recipient:
            QMessageBox.warning(self, "Orçamentos", "O cliente não tem email definido para preparar o envio.")
            return

        numero = str(payload.get("numero", "") or self.current_number or "").strip()
        if not numero:
            raise ValueError("Guarda primeiro o orçamento antes de preparar o email.")

        rfq_ref = str(payload.get("nota_cliente", "") or self.view.note_cliente_edit.text().strip()).strip()
        safe_ref = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in (rfq_ref or numero))[:48]
        attachment_name = f"Proposta_{safe_ref}_{numero}.pdf" if numero else f"Proposta_{safe_ref}.pdf"
        attachment_path: Path | None = Path(tempfile.gettempdir()) / attachment_name
        attachment_issue = ""
        try:
            self.services.orc_render_pdf(numero, attachment_path)
        except Exception as exc:
            attachment_issue = str(exc)
            attachment_path = None

        subject = self._quote_email_subject(payload)
        body_plain = self._quote_email_body(payload)
        logo_path = getattr(self.services, "logo_path", None)
        logo_file = Path(logo_path) if isinstance(logo_path, Path) and logo_path.exists() else None
        logo_cid = "lugest-mail-logo" if logo_file is not None else ""
        body_html = self._quote_email_html_body(payload, logo_cid=logo_cid)

        env = os.environ.copy()
        env["LUGEST_MAIL_TO"] = recipient
        env["LUGEST_MAIL_SUBJECT"] = subject
        env["LUGEST_MAIL_BODY"] = body_html
        env["LUGEST_MAIL_ATTACHMENT"] = str(attachment_path) if attachment_path is not None else ""
        env["LUGEST_MAIL_LOGO"] = str(logo_file) if logo_file is not None else ""
        env["LUGEST_MAIL_LOGO_CID"] = logo_cid
        powershell_script = (
            "$ErrorActionPreference='Stop'; "
            "$outlook = New-Object -ComObject Outlook.Application; "
            "$mail = $outlook.CreateItem(0); "
            "$mail.To = $env:LUGEST_MAIL_TO; "
            "$mail.Subject = $env:LUGEST_MAIL_SUBJECT; "
            "if ($env:LUGEST_MAIL_LOGO -and (Test-Path $env:LUGEST_MAIL_LOGO)) "
            "{ "
            "  $logo = $mail.Attachments.Add($env:LUGEST_MAIL_LOGO); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', $env:LUGEST_MAIL_LOGO_CID); "
            "  $logo.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true) "
            "}; "
            "if ($env:LUGEST_MAIL_ATTACHMENT -and (Test-Path $env:LUGEST_MAIL_ATTACHMENT)) "
            "{ $null = $mail.Attachments.Add($env:LUGEST_MAIL_ATTACHMENT) }; "
            "$mail.HTMLBody = $env:LUGEST_MAIL_BODY; "
            "$mail.Display()"
        )
        def on_email_ready(ok: bool, _error: str) -> None:
            if not ok:
                mailto = f"mailto:{quote(recipient)}?subject={quote(subject)}&body={quote(body_plain)}"
                try:
                    os.startfile(mailto)
                except Exception as exc:
                    QMessageBox.warning(self, "Orçamentos", f"Não foi possível abrir o cliente de email:\n{exc}")
                    return
                fallback_message = "Outlook indisponível. Foi aberto o cliente de email por defeito."
                if attachment_issue:
                    fallback_message += f"\n\nTambém não foi possível gerar o PDF em anexo:\n{attachment_issue}"
                else:
                    fallback_message += "\n\nNota: o anexo PDF terá de ser adicionado manualmente neste modo."
                QMessageBox.information(self, "Orçamentos", fallback_message)
                return
            if attachment_issue:
                QMessageBox.information(
                    self,
                    "Orçamentos",
                    f"Rascunho aberto no Outlook, mas o PDF não foi anexado automaticamente:\n{attachment_issue}",
                )

        _run_process_async(
            self,
            "powershell",
            ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
            environment=env,
            timeout_ms=30000,
            finished=on_email_ready,
        )

    def _set_quote_state(self, estado: str) -> None:
        try:
            self.services.orc_save(self._quote_payload())
            detail = self.services.orc_set_state(self.current_number, estado)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._load_quote(str(detail.get("numero", "") or "").strip())
        self.refresh()
        self._show_detail()
        if str(estado or "").strip().lower() == "enviado":
            try:
                self._open_quote_email_draft(detail)
            except Exception as exc:
                QMessageBox.warning(
                    self,
                    "Orçamentos",
                    "O orçamento foi marcado como Enviado, mas não foi possível abrir o email:\n"
                    f"{exc}",
                )

    def _remove_quote(self) -> None:
        row = self._selected_quote_row()
        numero = str((row or {}).get("numero", "") or self.current_number).strip()
        if not numero:
            QMessageBox.warning(self, "Orçamentos", "Seleciona um orcamento.")
            return
        if QMessageBox.question(self, "Apagar orcamento", f"Remover orcamento {numero}?") != QMessageBox.Yes:
            return
        try:
            self.services.orc_remove(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        self._show_list()
        self.refresh()

    def _line_dialog(self, initial: dict | None = None, *, template_mode: bool = False) -> dict | None:
        return edit_line(self, self.services.editor_ports, initial, template_mode=template_mode, presets=self.presets, client_code=self._client_code_from_text(self.view.client_combo.currentText()), current_number=self.current_number, line_rows=self.line_rows)

    def _assembly_model_editor_dialog(self, initial: dict | None = None) -> dict | None:
        initial = dict(initial or {})
        dialog = QDialog(self)
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

        items: list[dict] = [self._wrap_assembly_item(dict(row or {})) for row in list(initial.get("itens", []) or [])]
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
                        self._assembly_item_kind_label(self._assembly_item_kind_from_line(item)),
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
            kind = self._pick_assembly_item_kind(dialog)
            if not kind:
                return
            payload = self._open_assembly_item_editor(kind, parent=dialog)
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
            if self._quote_line_is_laser_2d(current_line):
                batch_id = str(current_line.get("laser_batch_id", "") or "").strip()
                batch_indexes = [
                    row_index
                    for row_index, candidate in enumerate(items)
                    if batch_id
                    and str(dict(candidate.get("line") or candidate).get("laser_batch_id", "") or "").strip() == batch_id
                ] or [index]
                source_lines = [dict(items[row_index].get("line") or items[row_index]) for row_index in batch_indexes]
                edited_lines = self._edit_laser_batch_lines(source_lines, parent=dialog)
                if edited_lines is None:
                    return
                insert_at = min(batch_indexes)
                for row_index in sorted(batch_indexes, reverse=True):
                    del items[row_index]
                for offset, edited_line in enumerate(edited_lines):
                    items.insert(insert_at + offset, self._wrap_assembly_item(edited_line))
                render_items()
                if edited_lines:
                    items_table.selectRow(insert_at)
                return
            else:
                payload = self._open_assembly_item_editor(self._assembly_item_kind_from_line(current), current, parent=dialog)
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

    def _manage_assembly_models(self) -> None:
        dialog = QDialog(self)
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
            rows = list(self.services.assembly_model_rows() or [])
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
            payload = self._assembly_model_editor_dialog()
            if payload is None:
                return
            try:
                self.services.assembly_model_save(payload)
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
                detail = self.services.assembly_model_detail(code)
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            payload = self._assembly_model_editor_dialog(detail)
            if payload is None:
                return
            payload["codigo"] = code
            try:
                self.services.assembly_model_save(payload)
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
                self.services.assembly_model_remove(code)
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

    def _manage_saved_conjuntos(self) -> None:
        dialog = QDialog(self)
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
            rows = list(self.services.conjunto_rows() or [])
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
            payload = self._calculated_assembly_builder_dialog()
            if not payload:
                return
            refresh_rows(str(payload.get("assembly_code", "") or "").strip())

        def edit_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                detail = dict(self.services.conjunto_detail(code) or {})
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            payload = self._calculated_assembly_builder_dialog(detail)
            if not payload:
                return
            refresh_rows(str(payload.get("assembly_code", code) or code).strip())

        def duplicate_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                detail = dict(self.services.conjunto_detail(code) or {})
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            detail["codigo"] = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            detail.pop("param_codigo", None)
            detail["descricao"] = f"{str(detail.get('descricao', code) or code).strip()} (Copia)"
            detail.pop("created_at", None)
            detail.pop("updated_at", None)
            payload = self._calculated_assembly_builder_dialog(detail)
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
                self.line_rows.extend(self.services.conjunto_expand(code, qty))
            except Exception as exc:
                QMessageBox.critical(dialog, "Conjuntos", str(exc))
                return
            self._render_quote_lines()
            QMessageBox.information(dialog, "Conjuntos", f"O conjunto {code} foi adicionado ao orçamento.")

        def preview_conjunto() -> None:
            code = current_code()
            if not code:
                QMessageBox.warning(dialog, "Conjuntos", "Seleciona um conjunto.")
                return
            try:
                path = self.services.conjunto_open_sheet_pdf(code)
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
                self.services.conjunto_remove(code)
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

    def _save_selected_lines_as_group(self) -> None:
        selected_indexes = self._selected_line_indexes()
        if not selected_indexes:
            QMessageBox.warning(self, "Conjuntos", "Seleciona pelo menos uma linha para guardar no conjunto/modelo.")
            return

        selected_rows = [dict(self.line_rows[index] or {}) for index in selected_indexes]
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
            prefill_name = str(self.view.note_cliente_edit.text() or self.current_number or "").strip()
        if not prefill_code:
            prefill_code = f"CJ-{datetime.now().strftime('%Y%m%d%H%M%S')}"

        dialog = QDialog(self)
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
            f"Guardado a partir do orçamento {str(self.current_number or '').strip()} com {len(selected_rows)} linha(s) selecionada(s)."
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
                    self._quote_line_type_label(row),
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
                return [dict(row or {}) for row in list(self.services.conjunto_rows() or [])]
            if dest == "modelo":
                return [dict(row or {}) for row in list(self.services.assembly_model_rows() or [])]
            combined: dict[str, dict] = {}
            for source_row in list(self.services.conjunto_rows() or []):
                row = dict(source_row or {})
                code = str(row.get("codigo", "") or "").strip()
                if code:
                    combined[code] = {
                        "codigo": code,
                        "descricao": str(row.get("descricao", "") or "").strip(),
                        "source": "both",
                    }
            for source_row in list(self.services.assembly_model_rows() or []):
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
                    detail = dict(self.services.conjunto_detail(target_code) or {})
                except Exception:
                    detail = {}
                rows = [dict(item or {}) for item in list(detail.get("itens", []) or []) if isinstance(item, dict)]
                if rows:
                    return rows
            if destination in {"modelo", "both"}:
                try:
                    detail = dict(self.services.assembly_model_detail(target_code) or {})
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
            if self.services.orc_line_is_product(row) and not str(row.get("ref_externa", "") or "").strip():
                row["ref_externa"] = str(row.get("produto_codigo", "") or "").strip()
            operation_norm = self.services.norm_text(str(row.get("operacao", "") or ""))
            if self.services.orc_line_is_piece(row) and "laser" in operation_norm:
                row["source_quote_number"] = str(self.current_number or "").strip()
                row["source_ref_externa"] = str(row.get("ref_externa", "") or "").strip()
                row["pricing_source"] = "quote_laser"
            return row

        def _compose_rows_for_save(destination: str, target_code: str, overwrite: bool) -> list[dict]:
            rows = []
            if overwrite:
                quote_rows = [
                    dict(row or {})
                    for row in list(self.line_rows)
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
            QMessageBox.warning(self, "Conjuntos", "Indica um código para o conjunto/modelo.")
            return
        description = str(name_edit.text() or "").strip() or code
        notes = str(notes_edit.toPlainText() or "").strip()
        if not overwrite and code in _existing_codes(destination):
            QMessageBox.warning(
                self,
                "Conjuntos",
                f"Ja existe um registo com o codigo {code}. Se queres atualizar esse registo, usa o modo 'Sobrepor existente'.",
            )
            return

        lines_for_save = _compose_rows_for_save(destination, code, overwrite)
        if not lines_for_save:
            QMessageBox.warning(self, "Conjuntos", "Nao ha linhas validas para guardar no conjunto/modelo.")
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
            if destination in {"conjunto", "both"}:
                self.services.conjunto_save(payload)
            if destination in {"modelo", "both"}:
                self.services.assembly_model_save(payload)
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return

        existing_group_rows = [
            dict(row or {})
            for row in list(self.line_rows)
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
                for index, row in enumerate(self.line_rows)
                if str((row or {}).get("conjunto_codigo", "") or "").strip() == code
            )
        for index in sorted(affected_indexes):
            if not (0 <= index < len(self.line_rows)):
                continue
            row = dict(self.line_rows[index] or {})
            row["conjunto_codigo"] = code
            row["conjunto_nome"] = description
            row["grupo_uuid"] = group_uuid
            self.line_rows[index] = row

        self._render_quote_lines()
        if selected_indexes:
            self._select_quote_line_source_index(selected_indexes[0])
        QMessageBox.information(
            self,
            "Conjuntos",
            (
                f"As linhas selecionadas foram guardadas em {description} ({code})."
                if not overwrite
                else f"O registo {code} foi atualizado e as linhas ficaram ligadas ao conjunto/modelo."
            ),
        )

    def _add_assembly_model(self) -> None:
        rows = list(self.services.conjunto_rows() or [])
        expand_fn = getattr(self.services, "conjunto_expand", None)
        empty_message = "Ainda nao existem conjuntos guardados. Cria primeiro um conjunto calculado."
        if not rows:
            rows = list(self.services.assembly_model_rows() or [])
            expand_fn = getattr(self.services, "assembly_model_expand", None)
            empty_message = "Ainda nao existem modelos. Cria primeiro um modelo de conjunto."
        if not rows:
            QMessageBox.information(self, "Conjuntos", empty_message)
            self._manage_saved_conjuntos()
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Adicionar conjunto guardado")
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Seleciona um conjunto guardado para o expandir no orçamento atual. "
            "Os conjuntos calculados usam uma lógica separada das linhas DXF/DWG."
        )
        intro.setWordWrap(True)
        intro.setProperty("role", "muted")
        layout.addWidget(intro)
        form = QFormLayout()
        combo = QComboBox()
        for row in rows:
            combo.addItem(f"{row.get('codigo', '')} - {row.get('descricao', '')}", str(row.get("codigo", "") or "").strip())
        qty_spin = QDoubleSpinBox()
        qty_spin.setRange(0.01, 1000000.0)
        qty_spin.setDecimals(2)
        qty_spin.setValue(1.0)
        form.addRow("Modelo", combo)
        form.addRow("Quantidade conjuntos", qty_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            if not callable(expand_fn):
                raise ValueError("Funcao de expansao do conjunto indisponivel.")
            self.line_rows.extend(expand_fn(str(combo.currentData() or "").strip(), qty_spin.value()))
        except Exception as exc:
            QMessageBox.critical(self, "Conjuntos", str(exc))
            return
        self._render_quote_lines()

    def _configure_laser_profiles(self) -> None:
        dialog = self.services.laser_settings_dialog(self)
        dialog.exec()

    def _configure_operation_profiles(self) -> None:
        self.services.open_operation_profiles(self)

    def _resolve_laser_edit_source(self, row: dict) -> dict:
        source = dict(row or {})
        drawing = str(source.get("desenho", "") or "").strip()
        if Path(drawing).suffix.casefold() in {".dxf", ".dwg"}:
            return source
        finder = getattr(self.services, "_conjunto_find_quote_source", None)
        if not callable(finder):
            return source
        try:
            quote_line, quote_number = finder(source, str(source.get("conjunto_codigo", "") or "").strip())
        except Exception:
            return source
        if not isinstance(quote_line, dict):
            return source
        resolved = dict(source)
        for key, value in dict(quote_line or {}).items():
            if key not in resolved or resolved.get(key) in (None, "", [], {}):
                resolved[key] = value
        if quote_number and not str(resolved.get("source_quote_number", "") or "").strip():
            resolved["source_quote_number"] = str(quote_number).strip()
        return resolved

    def _quote_line_is_laser_2d(self, row: dict | None) -> bool:
        line = dict(row or {})
        if not self.services.orc_line_is_piece(line):
            return False
        operation = str(line.get("operacao", "") or "").casefold()
        has_laser = bool(line.get("laser_base_active", False) or "laser" in operation)
        if not has_laser:
            return False
        if str(line.get("laser_source_mode", "") or "").strip().casefold() == "batch":
            return True
        if str(line.get("laser_batch_id", "") or "").strip():
            return True
        drawing = str(line.get("desenho", "") or "").strip()
        suffix = Path(drawing).suffix.casefold()
        return suffix in {".dxf", ".dwg"}

    def _edit_laser_batch_lines(
        self,
        rows: list[dict],
        *,
        parent: QWidget | None = None,
    ) -> list[dict] | None:
        source_lines = [self._resolve_laser_edit_source(dict(row or {})) for row in rows if isinstance(row, dict)]
        if not source_lines:
            return None
        batch_dialog = self.services.laser_batch_dialog(parent if isinstance(parent, QWidget) else self,
            default_machine=(
                str(source_lines[0].get("laser_machine", source_lines[0].get("machine", "")) or "").strip()
                or self.view.workcenter_combo.currentText().strip()
            ),
            initial_lines=source_lines,
        )
        if batch_dialog.exec() != QDialog.Accepted:
            return None
        result = dict(batch_dialog.result_payload() or {})
        edited = [dict(row or {}) for row in list(result.get("lines", []) or []) if isinstance(row, dict) and row]
        if not edited:
            return None

        def identity(line: dict) -> tuple[str, str]:
            return (
                str(line.get("desenho", "") or "").strip().replace("\\", "/").casefold(),
                str(line.get("ref_externa", "") or "").strip().casefold(),
            )

        by_identity = {identity(source): source for source in source_lines}
        batch_id = str(batch_dialog.batch_id or "").strip()
        merged_lines: list[dict] = []
        preserved_keys = (
            "conjunto_codigo",
            "conjunto_nome",
            "conjunto_param_codigo",
            "grupo_uuid",
            "ficha_tecnica",
            "source_quote_number",
            "source_ref_externa",
            "pricing_source",
            "pricing_source_ref",
        )
        for line in edited:
            source = by_identity.get(identity(line), {})
            merged = {**source, **line}
            for key in preserved_keys:
                if key not in line and key in source:
                    merged[key] = source[key]
            merged["laser_source_mode"] = "batch"
            merged["laser_batch_id"] = batch_id
            merged_lines.append(merged)
        return merged_lines

    def _edit_laser_2d_line(self, row: dict, *, parent: QWidget | None = None) -> dict | None:
        source = self._resolve_laser_edit_source(dict(row or {}))
        dialog = self.services.laser_dialog(parent if isinstance(parent, QWidget) else self,
            default_machine=(
                str(source.get("laser_machine", source.get("machine", "")) or "").strip()
                or self.view.workcenter_combo.currentText().strip()
            ),
            initial_line=source,
        )
        if dialog.exec() != QDialog.Accepted:
            return None
        result = dict(dialog.result_payload() or {})
        laser_line = dict(result.get("line", {}) or {})
        if not laser_line:
            return None

        merged = {**source, **laser_line}
        try:
            source_operations = list(
                self.services.quote_parse_operacoes_lista(
                    source.get("operacoes_lista", source.get("operacao", ""))
                )
                or []
            )
            laser_operations = list(
                self.services.quote_parse_operacoes_lista(laser_line.get("operacao", "Corte Laser"))
                or []
            )
        except Exception:
            source_operations = [part.strip() for part in str(source.get("operacao", "") or "").split("+") if part.strip()]
            laser_operations = [part.strip() for part in str(laser_line.get("operacao", "Corte Laser") or "").split("+") if part.strip()]

        def is_laser_component(name: object) -> bool:
            normalized = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode().casefold()
            return any(token in normalized for token in ("laser", "marcacao", "defilm"))

        extra_operations = [name for name in source_operations if not is_laser_component(name)]
        combined_operations: list[str] = []
        for name in [*laser_operations, *extra_operations]:
            clean = str(name or "").strip()
            if clean and clean.casefold() not in {item.casefold() for item in combined_operations}:
                combined_operations.append(clean)

        extra_times = {
            str(key): float(value or 0)
            for key, value in dict(source.get("tempos_operacao", {}) or {}).items()
            if not is_laser_component(key)
        }
        extra_costs = {
            str(key): float(value or 0)
            for key, value in dict(source.get("custos_operacao", {}) or {}).items()
            if not is_laser_component(key)
        }
        laser_time = float(laser_line.get("tempo_peca_min", 0) or 0)
        laser_price = float(laser_line.get("preco_unit", 0) or 0)
        merged["operacao"] = " + ".join(combined_operations or ["Corte Laser"])
        merged["operacoes_lista"] = combined_operations or ["Corte Laser"]
        merged["tempo_peca_min"] = round(laser_time + sum(extra_times.values()), 4)
        merged["preco_unit"] = round(laser_price + sum(extra_costs.values()), 4)
        merged["total"] = round(float(merged.get("qtd", 0) or 0) * float(merged["preco_unit"]), 2)
        merged["laser_base_active"] = True
        merged["laser_base_tempo_unit"] = round(laser_time, 4)
        merged["laser_base_preco_unit"] = round(laser_price, 4)
        merged["tempos_operacao"] = extra_times
        merged["custos_operacao"] = extra_costs
        for key in ("desenho_pdf", "desenhos_pdf", "ficheiros", "conjunto_codigo", "conjunto_nome", "grupo_uuid"):
            if key not in laser_line and key in source:
                merged[key] = source[key]
        return merged

    def _add_laser_line(self) -> None:
        dialog = self.services.laser_dialog(self,
            default_machine=self.view.workcenter_combo.currentText().strip(),
        )
        if dialog.exec() != QDialog.Accepted:
            return
        result = dict(dialog.result_payload() or {})
        lines = [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})]
        line = dict(result.get("line", {}) or {})
        analysis = dict(result.get("analysis", {}) or {})
        if not lines and line:
            lines = [line]
        if not lines:
            QMessageBox.warning(self, "Peca Unit. DXF/DWG", "Nao foi possivel gerar a linha de orcamento.")
            return
        self.line_rows.extend(lines)
        machine_name = str(dict(analysis.get("machine", {}) or {}).get("name", "") or "").strip()
        if machine_name and not self.view.workcenter_combo.currentText().strip():
            self.view.workcenter_combo.setCurrentText(machine_name)
        self._render_quote_lines()

    def _add_laser_batch_lines(self) -> None:
        dialog = self.services.laser_batch_dialog(self,
            default_machine=self.view.workcenter_combo.currentText().strip(),
        )
        if dialog.exec() != QDialog.Accepted:
            return
        result = dict(dialog.result_payload() or {})
        lines = [dict(row or {}) for row in list(result.get("lines", []) or []) if dict(row or {})]
        analysis = dict(result.get("analysis", {}) or {})
        if not lines:
            QMessageBox.warning(self, "Lote DXF/DWG", "Nao foi possivel gerar linhas de orcamento para o lote.")
            return
        self.line_rows.extend(lines)
        machine_name = str(dict(analysis.get("machine", {}) or {}).get("name", "") or "").strip()
        if machine_name and not self.view.workcenter_combo.currentText().strip():
            self.view.workcenter_combo.setCurrentText(machine_name)
        self._render_quote_lines()

    def _open_laser_nesting(self) -> None:
        selected_indexes = self._selected_line_indexes()
        candidate_rows = [dict(self.line_rows[index] or {}) for index in selected_indexes] if selected_indexes else [dict(row or {}) for row in self.line_rows]
        laser_rows = [
            row
            for row in candidate_rows
            if str(row.get("desenho", "") or "").strip()
            and "corte laser" in str(row.get("operacao", "") or "").strip().lower()
        ]
        if not laser_rows:
            QMessageBox.information(self, "Nesting Laser", "Seleciona primeiro linhas laser com desenho associado.")
            return
        dialog = self.services.nesting_dialog(laser_rows, self, quote_number=self.current_number)
        dialog.exec()

    def _open_profile_step_igs_quote_builder(
        self,
        _checked: bool = False,
        *,
        return_lines: bool = False,
        parent: QWidget | None = None,
    ) -> list[dict] | None:
        lines = edit_profile(self, self.services.profile_editor_ports, parent=parent)
        if return_lines:
            return lines or []
        if lines is None:
            return None
        self.line_rows.extend(lines)
        if not self.view.workcenter_combo.currentText().strip():
            self.view.workcenter_combo.setCurrentText("Laser")
        self._render_quote_lines()
        return None

    def _add_line(self) -> None:
        try:
            payload = self._line_dialog({"qtd": 1})
            if payload is None:
                return
            self.line_rows.append(payload)
            self._render_quote_lines()
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))

    def _edit_line(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        try:
            current = dict(self.line_rows[index] or {})
            batch_id = str(current.get("laser_batch_id", "") or "").strip()
            is_saved_batch = (
                str(current.get("laser_source_mode", "") or "").strip().casefold() == "batch"
                or bool(batch_id)
            )
            if self._quote_line_is_laser_2d(current) and is_saved_batch:
                batch_indexes = [
                    row_index
                    for row_index, line in enumerate(self.line_rows)
                    if batch_id
                    and str(dict(line or {}).get("laser_batch_id", "") or "").strip() == batch_id
                ] or [index]
                edited_lines = self._edit_laser_batch_lines(
                    [dict(self.line_rows[row_index] or {}) for row_index in batch_indexes]
                )
                if edited_lines is None:
                    return
                insert_at = min(batch_indexes)
                for row_index in sorted(batch_indexes, reverse=True):
                    del self.line_rows[row_index]
                for offset, edited_line in enumerate(edited_lines):
                    self.line_rows.insert(insert_at + offset, edited_line)
                self._render_quote_lines()
                if edited_lines:
                    self._select_quote_line_source_index(insert_at)
                return
            payload = self._edit_laser_2d_line(current) if self._quote_line_is_laser_2d(current) else self._line_dialog(current)
            if payload is None:
                return
            self.line_rows[index] = payload
            self._render_quote_lines()
            self._select_quote_line_source_index(index)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))

    def _prepare_selected_line_for_production(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        current = dict(self.line_rows[index] or {})
        if not self.services.orc_line_is_piece(current):
            QMessageBox.information(
                self,
                "Preparar producao",
                "Esta acao so se aplica a pecas fabricadas. Produtos de stock continuam no fluxo de consumo de montagem.",
            )
            return
        payload = self._line_dialog(current)
        if payload is None:
            return
        self.line_rows[index] = payload
        self._render_quote_lines()
        self._select_quote_line_source_index(index)
        drawing_ready = bool(str(payload.get("desenho", "") or "").strip())
        ops_ready = bool(str(payload.get("operacao", "") or "").strip())
        if drawing_ready and ops_ready:
            QMessageBox.information(
                self,
                "Preparar producao",
                "A linha ficou preparada para fluxo de operador. Na conversão para encomenda vai seguir para produção.",
            )
        else:
            QMessageBox.information(
                self,
                "Preparar producao",
                "A linha foi atualizada, mas so entra em producao quando tiver desenho tecnico e operacoes definidas.",
            )

    def _remove_line(self) -> None:
        indexes = self._checked_line_indexes() or self._selected_line_indexes()
        if not indexes:
            QMessageBox.warning(self, "Orçamentos", "Marca ou seleciona pelo menos uma linha.")
            return
        references = [
            self._quote_line_primary_ref(dict(self.line_rows[index] or {}))
            for index in indexes
            if 0 <= index < len(self.line_rows)
        ]
        count = len(indexes)
        detail = ", ".join(reference for reference in references[:3] if reference and reference != "-")
        if len(references) > 3:
            detail += f" e mais {len(references) - 3}"
        answer = QMessageBox.question(
            self,
            "Remover linhas do orçamento",
            (
                f"Remover {count} linha{'s' if count != 1 else ''} do orçamento atual?"
                + (f"\n\n{detail}" if detail else "")
                + "\n\nEsta alteração só fica definitiva quando guardares o orçamento."
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self._rendering_quote_lines = True
        for visual_row in range(self.view.lines_table.rowCount()):
            item = self.view.lines_table.item(visual_row, self.LINE_COL_MARK)
            if item is not None and item.flags() & Qt.ItemIsUserCheckable:
                item.setCheckState(Qt.Unchecked)
        self._rendering_quote_lines = False
        next_source_index = min(indexes)
        for index in sorted(indexes, reverse=True):
            if 0 <= index < len(self.line_rows):
                del self.line_rows[index]
        self._render_quote_lines()
        if self.line_rows:
            self._select_quote_line_source_index(min(next_source_index, len(self.line_rows) - 1))

    def _open_line_drawing(self) -> None:
        index = self._selected_line_index()
        if index < 0:
            QMessageBox.warning(self, "Orçamentos", "Seleciona uma linha.")
            return
        path = str(self.line_rows[index].get("desenho", "") or "").strip()
        if not path:
            QMessageBox.information(self, "Orçamentos", "A linha nao tem desenho associado.")
            return
        if not Path(path).exists():
            QMessageBox.critical(self, "Orçamentos", f"Ficheiro nao encontrado:\n{path}")
            return
        os.startfile(path)

    def _check_selected_line_weight(self) -> None:
        indexes = self._selected_line_indexes()
        if not indexes:
            QMessageBox.warning(self, "Check Weight", "Seleciona pelo menos uma linha.")
            return
        results: list[dict[str, Any]] = []
        issues: list[str] = []
        for index in indexes:
            row = dict(self.line_rows[index] or {})
            desenho = str(row.get("desenho", "") or "").strip()
            if not desenho:
                issues.append(f"Linha {index + 1}: sem desenho associado.")
                continue
            if not Path(desenho).exists():
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: desenho nao encontrado.")
                continue
            esp_raw = row.get("espessura", row.get("espessura_mm", row.get("esp", "")))
            try:
                thickness_mm = float(str(esp_raw).replace(",", ".").strip() or 0)
            except Exception:
                thickness_mm = 0.0
            if thickness_mm <= 0:
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: espessura invalida.")
                continue
            payload = {
                "path": desenho,
                "machine": self.view.workcenter_combo.currentText().strip(),
                "commercial_profile": "",
                "material": str(row.get("material_family", "") or row.get("material", "") or "").strip(),
                "material_subtype": str(row.get("material_subtype", "") or "").strip(),
                "gas": "",
                "thickness_mm": thickness_mm,
                "quantity": max(1, int(float(row.get("qtd", 1) or 1))),
                "material_supplied_by_client": bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False)),
            }
            try:
                analysis = dict(self.services.laser_quote_analyze(payload) or {})
            except Exception as exc:
                issues.append(f"{str(row.get('ref_externa', '') or row.get('descricao', '-') or '-')}: {exc}")
                continue
            metrics = dict(analysis.get("metrics", {}) or {})
            mass_unit = float(metrics.get("net_mass_kg", 0.0) or 0.0)
            qty = max(1, int(float(row.get("qtd", 1) or 1)))
            results.append(
                {
                    "row_index": index,
                    "ref": str(row.get("ref_externa", "") or row.get("descricao", "-") or "-").strip(),
                    "material": str((analysis.get("machine", {}) or {}).get("material", row.get("material", "-")) or "-").strip(),
                    "thickness_mm": thickness_mm,
                    "qty": qty,
                    "mass_unit": mass_unit,
                    "mass_total": mass_unit * qty,
                }
            )
        if not results and issues:
            QMessageBox.warning(self, "Check Weight", "\n".join(issues))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Check Weight")
        dialog.resize(760, 360)
        layout = QVBoxLayout(dialog)
        title = QLabel("Peso por unidade das linhas selecionadas")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        layout.addWidget(title)
        help_lbl = QLabel("Informacao de apoio. O peso e calculado a partir do desenho, espessura e densidade configurada para o material.")
        help_lbl.setProperty("role", "muted")
        help_lbl.setWordWrap(True)
        layout.addWidget(help_lbl)
        table = QTableWidget(len(results), 6)
        table.setHorizontalHeaderLabels(["Ref. Ext.", "Material", "Esp. (mm)", "Qtd", "Peso/un (kg)", "Peso linha (kg)"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setAlternatingRowColors(True)
        table.setStyleSheet(
            "QTableWidget { font-size: 12px; }"
            " QHeaderView::section { font-size: 11px; padding: 6px 8px; font-weight: 800; }"
        )
        _set_table_columns(
            table,
            [
                (0, "stretch", 0),
                (1, "fixed", 130),
                (2, "fixed", 80),
                (3, "fixed", 60),
                (4, "fixed", 110),
                (5, "fixed", 118),
            ],
        )
        for row_index, result in enumerate(results):
            values = [
                str(result.get("ref", "-")),
                str(result.get("material", "-")),
                f"{float(result.get('thickness_mm', 0.0) or 0.0):.3f}",
                str(int(result.get("qty", 1) or 1)),
                f"{float(result.get('mass_unit', 0.0) or 0.0):.3f}",
                f"{float(result.get('mass_total', 0.0) or 0.0):.3f}",
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col_index >= 2:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                table.setItem(row_index, col_index, item)
        table.setMinimumHeight(max(140, min(320, 42 + (len(results) * 30))))
        layout.addWidget(table)
        if issues:
            issues_lbl = QLabel("Linhas ignoradas:\n" + "\n".join(issues))
            issues_lbl.setWordWrap(True)
            issues_lbl.setStyleSheet("color: #9a3412; font-size: 11px;")
            layout.addWidget(issues_lbl)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _preview_quote(self) -> None:
        try:
            detail = self.services.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            path = self.services.orc_open_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF aberto:\n{path}")

    def _save_quote_pdf(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Orçamentos", "Guarda primeiro o orcamento.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", f"orcamento_{self.current_number}.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            detail = self.services.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            self.services.orc_render_pdf(numero, path)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF guardado em:\n{path}")

    def _print_quote_pdf(self) -> None:
        try:
            detail = self.services.orc_save(self._quote_payload())
            numero = str(detail.get("numero", "") or self.current_number).strip()
            self.current_number = numero
            path = self.services.orc_print_pdf(numero)
        except Exception as exc:
            QMessageBox.critical(self, "Orçamentos", str(exc))
            return
        QMessageBox.information(self, "Orçamentos", f"PDF enviado para impressão:\n{path}")

    def _convert_quote(self) -> None:
        if not self.current_number:
            QMessageBox.warning(self, "Orçamentos", "Guarda primeiro o orcamento.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Converter em encomenda")
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        note_edit = QLineEdit(self.view.note_cliente_edit.text().strip())
        form.addRow("Nota cliente", note_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            self.services.orc_save(self._quote_payload())
            result = self.services.orc_convert_to_order(self.current_number, note_edit.text().strip())
        except Exception as exc:
            QMessageBox.critical(self, "Converter em encomenda", str(exc))
            return
        self.refresh()
        self._load_quote(self.current_number)
        self._show_detail()
        enc_num = str(((result or {}).get("encomenda", {}) or {}).get("numero", "") or "").strip()
        QMessageBox.information(self, "Converter em encomenda", f"Encomenda criada: {enc_num or '-'}")
