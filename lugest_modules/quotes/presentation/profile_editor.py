from __future__ import annotations
import re
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
from lugest_core.cad.profile_analysis import analyze_profile_cut_features, render_step_preview_image
from pathlib import Path
from lugest_qt.ui.pages.laser_quote_dialogs import LaserSettingsDialog, _canonical_material_family as _laser_canonical_material_family, _display_material_family as _laser_display_material_family, _guess_material_family as _laser_guess_material_family, _set_combo_values as _laser_set_combo_values, _settings_gas_names as _laser_settings_gas_names, _settings_material_names as _laser_settings_material_names, _settings_material_subtypes as _laser_settings_material_subtypes
from lugest_qt.ui.pages.runtime_common import selected_row_index as _selected_row_index, set_table_columns as _set_table_columns
from lugest_qt.ui.pages.runtime_support import _fmt_eur
from lugest_qt.ui.widgets import CardFrame, FlexibleDecimalSpinBox as QDoubleSpinBox


from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class ProfileEditorPorts:
    ORC_LINE_TYPE_SERVICE: str
    laser_quote_settings: Callable[[], dict]
    laser_quote_save_settings: Callable[[dict], Any]
    material_presets: Callable[[], dict]
    profile_laser_quote_analyze: Callable[[dict], dict]
    profile_laser_quote_build_line: Callable[[dict], dict]

def edit_profile(owner: QWidget, backend: ProfileEditorPorts, *, parent: QWidget | None = None) -> list[dict] | None:
    dialog = QDialog(parent if isinstance(parent, QWidget) else owner)
    dialog.setWindowTitle("Corte Laser STEP/IGS")
    dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)
    dialog.setSizeGripEnabled(True)
    dialog.setMinimumSize(860, 520)
    try:
        screen = dialog.screen() or QApplication.primaryScreen()
        available = screen.availableGeometry() if screen is not None else None
        if available is not None:
            width = min(1180, max(860, available.width() - 80))
            height = min(860, max(520, available.height() - 96))
            dialog.resize(width, height)
            dialog.move(
                available.x() + max(0, (available.width() - width) // 2),
                available.y() + max(0, (available.height() - height) // 2),
            )
    except Exception:
        dialog.resize(1120, 780)
    outer_layout = QVBoxLayout(dialog)
    outer_layout.setContentsMargins(14, 12, 14, 12)
    outer_layout.setSpacing(10)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll_content = QWidget()
    layout = QVBoxLayout(scroll_content)
    layout.setContentsMargins(2, 2, 2, 2)
    layout.setSpacing(10)

    title = QLabel("Corte laser de perfis / tubos / cantoneiras")
    title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
    intro = QLabel(
        "Fluxo separado do nesting plano. Aqui orçamentas apenas cortes, furos e rasgos a partir de ficheiros STEP/IGS, "
        "sem assumir o custo do material base do perfil."
    )
    intro.setWordWrap(True)
    intro.setProperty("role", "muted")
    layout.addWidget(title)
    layout.addWidget(intro)
    auto_hint = QLabel(
        "Leitura automatica: Eventos = furos + rasgos + outros cortes internos cobrados. "
        "Cortes terminais do perfil ficam separados e ignorados por defeito; o preco usa os metros reais, "
        "a espessura e a tabela da maquina selecionada."
    )
    auto_hint.setWordWrap(True)
    auto_hint.setProperty("role", "muted")
    layout.addWidget(auto_hint)

    laser_settings = dict(backend.laser_quote_settings() or {})

    toolbar = QHBoxLayout()
    toolbar.setSpacing(8)
    add_files_btn = QPushButton("Adicionar STEP/IGS")
    add_files_btn.setProperty("variant", "secondary")
    remove_file_btn = QPushButton("Remover selecionado")
    remove_file_btn.setProperty("variant", "secondary")
    configure_btn = QPushButton("Perfis laser")
    configure_btn.setProperty("variant", "secondary")
    toolbar.addWidget(add_files_btn)
    toolbar.addWidget(remove_file_btn)
    toolbar.addWidget(configure_btn)
    toolbar.addStretch(1)
    layout.addLayout(toolbar)

    files_table = QTableWidget(0, 10)
    files_table.setHorizontalHeaderLabels(["Ficheiro", "Tipo", "Secao", "Qtd", "Eventos", "Furos", "Rasgos", "m corte", "Comp. m", "kg/m"])
    files_table.verticalHeader().setVisible(False)
    files_table.verticalHeader().setDefaultSectionSize(42)
    files_table.verticalHeader().setMinimumSectionSize(38)
    files_table.setSelectionBehavior(QTableWidget.SelectRows)
    files_table.setEditTriggers(QTableWidget.NoEditTriggers)
    files_table.setAlternatingRowColors(True)
    files_table.setMinimumHeight(260)
    files_table.setStyleSheet(
        "QTableWidget { font-size: 12px; }"
        " QTableWidget::item { padding: 7px 6px; }"
        " QHeaderView::section { padding: 8px 6px; font-weight: 800; }"
    )
    files_table.horizontalHeader().setStretchLastSection(False)
    _set_table_columns(
        files_table,
        [
            (0, "stretch", 0),
            (1, "fixed", 120),
            (2, "fixed", 170),
            (3, "fixed", 72),
            (4, "fixed", 78),
            (5, "fixed", 72),
            (6, "fixed", 78),
            (7, "fixed", 86),
            (8, "fixed", 86),
            (9, "fixed", 86),
        ],
    )
    layout.addWidget(files_table)

    preview_card = CardFrame()
    preview_card.set_tone("default")
    preview_layout = QVBoxLayout(preview_card)
    preview_layout.setContentsMargins(14, 12, 14, 12)
    preview_layout.setSpacing(8)
    preview_title = QLabel("Preview do STEP/IGS")
    preview_title.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
    preview_image_label = QLabel("Seleciona um ficheiro para gerar preview.")
    preview_image_label.setAlignment(Qt.AlignCenter)
    preview_image_label.setMinimumHeight(240)
    preview_image_label.setStyleSheet("border: 1px solid #d0d5dd; border-radius: 10px; background: #f8fafc; color: #475467;")
    preview_info_label = QLabel("O preview usa FreeCAD quando estiver instalado neste posto.")
    preview_info_label.setWordWrap(True)
    preview_info_label.setProperty("role", "muted")
    preview_layout.addWidget(preview_title)
    preview_layout.addWidget(preview_image_label)
    preview_layout.addWidget(preview_info_label)
    layout.addWidget(preview_card)

    pricing_card = CardFrame()
    pricing_card.set_tone("default")
    pricing_layout = QGridLayout(pricing_card)
    pricing_layout.setContentsMargins(14, 12, 14, 12)
    pricing_layout.setHorizontalSpacing(12)
    pricing_layout.setVerticalSpacing(8)

    machine_combo = QComboBox()
    commercial_combo = QComboBox()
    material_combo = QComboBox()
    subtype_combo = QComboBox()
    subtype_combo.setEditable(True)
    gas_combo = QComboBox()
    thickness_spin = QDoubleSpinBox()
    material_price_spin = QDoubleSpinBox()
    material_price_spin.setRange(0.0, 1000000.0)
    material_price_spin.setDecimals(4)
    material_price_spin.setSingleStep(0.1)
    material_price_unit_combo = QComboBox()
    material_price_unit_combo.addItems(["EUR/kg", "EUR/m", "EUR/ton"])
    thickness_spin.setRange(0.1, 200.0)
    thickness_spin.setDecimals(2)
    thickness_spin.setSingleStep(0.5)
    thickness_spin.setValue(3.0)
    customer_material_check = QCheckBox("Perfil/tubo fornecido pelo cliente")
    customer_material_check.setChecked(True)
    customer_material_check.setToolTip("No fluxo STEP/IGS o material base do perfil fica fora do calculo por defeito.")
    total_label = QLabel("Total estimado: 0,00 EUR")
    total_label.setStyleSheet("font-size: 15px; font-weight: 800; color: #0f172a;")
    status_label = QLabel("Usa as tabelas do laser com contagem de eventos STEP/IGS, sem duplicar furos.")
    status_label.setWordWrap(True)
    status_label.setProperty("role", "muted")
    pricing_layout.addWidget(QLabel("Maquina"), 0, 0)
    pricing_layout.addWidget(machine_combo, 0, 1)
    pricing_layout.addWidget(QLabel("Perfil comercial"), 0, 2)
    pricing_layout.addWidget(commercial_combo, 0, 3)
    pricing_layout.addWidget(QLabel("Familia material"), 1, 0)
    pricing_layout.addWidget(material_combo, 1, 1)
    pricing_layout.addWidget(QLabel("Subtipo / qualidade"), 1, 2)
    pricing_layout.addWidget(subtype_combo, 1, 3)
    pricing_layout.addWidget(QLabel("Gas"), 2, 0)
    pricing_layout.addWidget(gas_combo, 2, 1)
    pricing_layout.addWidget(QLabel("Espessura (mm)"), 2, 2)
    pricing_layout.addWidget(thickness_spin, 2, 3)
    pricing_layout.addWidget(QLabel("Preco material"), 3, 0)
    pricing_layout.addWidget(material_price_spin, 3, 1)
    pricing_layout.addWidget(QLabel("Unid. preco material"), 3, 2)
    pricing_layout.addWidget(material_price_unit_combo, 3, 3)
    pricing_layout.addWidget(customer_material_check, 4, 0, 1, 2)
    pricing_layout.addWidget(total_label, 4, 2, 1, 2)
    pricing_layout.addWidget(status_label, 5, 0, 1, 4)
    layout.addWidget(pricing_card)
    scroll.setWidget(scroll_content)
    outer_layout.addWidget(scroll, 1)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
    buttons.button(QDialogButtonBox.Cancel).setText("Fechar")
    outer_layout.addWidget(buttons)

    result_lines: list[dict] = []
    families = ["Perfil", "Tubo", "Cantoneira", "Barra"]
    material_price_internal_update = False
    material_price_user_touched = False
    material_price_auto_filled = False

    def _infer_family(path_txt: str) -> str:
        probe = Path(path_txt).stem.lower()
        if "tubo" in probe:
            return "Tubo"
        if "cant" in probe:
            return "Cantoneira"
        if "barra" in probe or "chata" in probe:
            return "Barra"
        if any(token in probe for token in ("ipe", "ipn", "hea", "heb", "upn", "rhs", "shs", "perfil")):
            return "Perfil"
        return "Perfil"

    def _family_from_analysis(cad_analysis: dict[str, Any], fallback: str) -> str:
        family = str(cad_analysis.get("family_guess", "") or "").strip()
        if family in {"Tubo", "Cantoneira", "Barra", "Perfil"}:
            return family
        return fallback

    def _infer_section(path_txt: str) -> str:
        stem = Path(path_txt).stem.upper().replace("_", " ").replace("-", " ")
        profile_match = re.search(r"\b(IPE|IPN|HEA|HEB|UPN|RHS|SHS)\s*(\d+)\b", stem)
        if profile_match:
            return f"{profile_match.group(1)} {profile_match.group(2)}"
        size_match = re.search(r"(\d+\s*[Xx]\s*\d+(?:\s*[Xx]\s*\d+(?:[.,]\d+)?)?)", stem)
        if size_match:
            return re.sub(r"\s*[Xx]\s*", "x", size_match.group(1)).replace(",", ".")
        return ""

    def _infer_profile_length_m(path_txt: str) -> float:
        stem = Path(path_txt).stem.lower().replace("_", " ").replace("-", " ")
        unit_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(mm|m)\b", stem)
        if unit_match:
            value = float(unit_match.group(1).replace(",", "."))
            return round(value / 1000.0 if unit_match.group(2) == "mm" else value, 4)
        numbers = [float(item.replace(",", ".")) for item in re.findall(r"\d+(?:[.,]\d+)?", stem)]
        if len(numbers) == 1 and numbers[0] > 20:
            return round(numbers[0] / 1000.0, 4)
        return 0.0

    def _infer_material_from_profile(path_txt: str, text: str = "") -> tuple[str, str]:
        probe = f"{Path(path_txt).stem} {str(text or '')[:120000]}".upper()
        subtype_patterns = [
            r"\bS235(?:JR)?\b",
            r"\bS275(?:JR)?\b",
            r"\bS355(?:JR|J2\+N|MC|JOW)?\b",
            r"\bS420MC\b",
            r"\bDX5[13]D(?:\+Z)?\b",
            r"\bDD11\b",
            r"\bDC01\b",
            r"\bCORTEN\b",
            r"\bHARDOX(?:\s*4[05]0)?\b",
            r"\bINOX\s*3(?:04|16)L?\b",
            r"\bAISI\s*3(?:04|16)L?\b",
            r"\b1\.4(?:301|307|401|404|016)\b",
        ]
        matched_subtype = ""
        for pattern in subtype_patterns:
            match = re.search(pattern, probe)
            if match:
                matched_subtype = re.sub(r"\s+", " ", match.group(0).strip())
                break
        guessed = _laser_guess_material_family(matched_subtype)
        if guessed:
            return _laser_canonical_material_family(guessed) or guessed, matched_subtype
        if re.search(r"\b(INOX|STAINLESS|AISI\s*3(?:04|16)L?|1\.4(?:301|307|401|404|016))\b", probe):
            return "Aco inox", matched_subtype
        if re.search(r"\b(FERRO|ACO|AÇO|S235|S275|S355|S420|CORTEN|HARDOX|DX5[13]D)\b", probe):
            return "Aco carbono", matched_subtype
        return "", matched_subtype

    def _make_int_spin(value: int) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(0.0, 1000000.0)
        spin.setDecimals(0)
        spin.setSingleStep(1.0)
        spin.setValue(float(value))
        spin.setMinimumHeight(32)
        spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
        return spin

    def _read_geometry_preview(path_txt: str) -> str:
        try:
            raw = Path(path_txt).read_bytes()
        except Exception:
            return ""
        if not raw:
            return ""
        sample = raw[:1_200_000]
        for encoding in ("utf-8", "latin-1", "cp1252"):
            try:
                return sample.decode(encoding, errors="ignore")
            except Exception:
                continue
        return sample.decode("latin-1", errors="ignore")

    def _clear_step_preview(message: str, info: str = "") -> None:
        preview_image_label.clear()
        preview_image_label.setPixmap(QPixmap())
        preview_image_label.setText(message)
        preview_info_label.setText(info or "O preview usa FreeCAD quando estiver instalado neste posto.")

    def _update_step_preview() -> None:
        row_index = _selected_row_index(files_table)
        if row_index < 0:
            _clear_step_preview("Seleciona um ficheiro para gerar preview.")
            return
        file_item = files_table.item(row_index, 0)
        path_txt = str(file_item.data(Qt.UserRole) if isinstance(file_item, QTableWidgetItem) else "").strip()
        if not path_txt:
            _clear_step_preview("Sem ficheiro associado.")
            return
        cached_preview = dict(file_item.data(Qt.UserRole + 2) if isinstance(file_item, QTableWidgetItem) else {} or {})
        if not cached_preview:
            _clear_step_preview("A gerar preview FreeCAD...", Path(path_txt).name)
            QApplication.processEvents()
            try:
                cached_preview = dict(render_step_preview_image(path_txt) or {})
            except Exception as exc:
                cached_preview = {"available": False, "note": str(exc)}
            if isinstance(file_item, QTableWidgetItem):
                file_item.setData(Qt.UserRole + 2, dict(cached_preview))
        image_path = str(cached_preview.get("image_path", "") or "").strip()
        if cached_preview.get("available") and image_path and Path(image_path).exists():
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                preview_image_label.setText("")
                preview_image_label.setPixmap(
                    pixmap.scaled(760, 260, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                )
                preview_info_label.setText(
                    f"{Path(path_txt).name} | Preview gerado por {str(cached_preview.get('engine', 'FreeCAD') or 'FreeCAD')}"
                )
                return
        note_txt = str(cached_preview.get("note", "") or "").strip()
        _clear_step_preview(
            "Preview indisponivel neste posto.",
            note_txt or f"{Path(path_txt).name} | Instala/configura o FreeCAD para gerar imagem do STEP.",
        )

    def _refresh_machine_and_commercial() -> None:
        _laser_set_combo_values(machine_combo, list(dict(laser_settings.get("machine_profiles", {}) or {}).keys()))
        _laser_set_combo_values(commercial_combo, list(dict(laser_settings.get("commercial_profiles", {}) or {}).keys()))

    def _material_subtype_candidates(material_name: str) -> list[str]:
        extras: list[str] = []
        try:
            presets = dict(backend.material_presets() or {})
            family = _laser_guess_material_family(material_name) or str(material_name or "").strip()
            for value in list(presets.get("materiais", []) or []):
                clean = str(value or "").strip()
                if not clean or clean == material_name:
                    continue
                if _laser_guess_material_family(clean) == family and clean not in extras:
                    extras.append(clean)
        except Exception:
            pass
        return _laser_settings_material_subtypes(laser_settings, material_name, extras)

    def _refresh_materials() -> None:
        current_material = _laser_display_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip()
        values = _laser_settings_material_names(laser_settings, machine_combo.currentText().strip())
        _laser_set_combo_values(material_combo, values, current_material if current_material else (values[0] if values else "Ferro"))
        _refresh_subtypes()

    def _refresh_subtypes() -> None:
        current_subtype = subtype_combo.currentText().strip()
        values = _material_subtype_candidates(material_combo.currentText().strip())
        _laser_set_combo_values(subtype_combo, values, current_subtype)
        if current_subtype and current_subtype not in [subtype_combo.itemText(index) for index in range(subtype_combo.count())]:
            subtype_combo.setCurrentText(current_subtype)
        _refresh_gases()

    def _refresh_gases() -> None:
        current_gas = gas_combo.currentText().strip()
        values = _laser_settings_gas_names(laser_settings, machine_combo.currentText().strip(), material_combo.currentText().strip())
        _laser_set_combo_values(gas_combo, values, current_gas if current_gas else (values[0] if values else "Oxigenio"))

    def _configure_profiles() -> None:
        nonlocal laser_settings
        cfg_dialog = LaserSettingsDialog(backend, owner)
        if cfg_dialog.exec() != QDialog.Accepted:
            return
        current_machine = machine_combo.currentText().strip()
        current_commercial = commercial_combo.currentText().strip()
        current_material = material_combo.currentText().strip()
        current_gas = gas_combo.currentText().strip()
        laser_settings = dict(backend.laser_quote_settings() or {})
        _laser_set_combo_values(machine_combo, list(dict(laser_settings.get("machine_profiles", {}) or {}).keys()), current_machine)
        _laser_set_combo_values(commercial_combo, list(dict(laser_settings.get("commercial_profiles", {}) or {}).keys()), current_commercial)
        _refresh_materials()
        if current_material:
            material_combo.setCurrentText(current_material)
            _refresh_subtypes()
        if current_gas:
            gas_combo.setCurrentText(current_gas)
        _recalc_total()

    def _selected_profile_family() -> str:
        row_index = _selected_row_index(files_table)
        if row_index < 0 and files_table.rowCount() > 0:
            row_index = 0
        if row_index < 0:
            return ""
        family_widget = files_table.cellWidget(row_index, 1)
        if isinstance(family_widget, QComboBox):
            return family_widget.currentText().strip()
        return ""

    def _sync_material_unit_for_family() -> None:
        nonlocal material_price_internal_update
        if _selected_profile_family() != "Tubo":
            return
        material_price_internal_update = True
        try:
            material_price_unit_combo.setCurrentText("EUR/m")
        finally:
            material_price_internal_update = False

    def _sync_material_cost_controls() -> None:
        enabled = not bool(customer_material_check.isChecked())
        material_price_spin.setEnabled(enabled)
        material_price_unit_combo.setEnabled(enabled)

    def _refresh_material_price_default() -> None:
        nonlocal material_price_internal_update, material_price_auto_filled
        if material_price_user_touched:
            return
        if _selected_profile_family() == "Tubo":
            _sync_material_unit_for_family()
            return
        if material_price_spin.value() > 0.0:
            return
        try:
            profiles = dict(laser_settings.get("commercial_profiles", {}) or {})
            profile = dict(profiles.get(commercial_combo.currentText().strip(), {}) or {})
            family = _laser_canonical_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip()
            subtype = subtype_combo.currentText().strip()
            material = dict(dict(profile.get("materials", {}) or {}).get(family, {}) or {})
            catalog = dict(dict(profile.get("material_catalog", {}) or {}).get(family, {}) or {})
            if subtype and subtype in catalog:
                material.update(dict(catalog.get(subtype, {}) or {}))
            price = float(material.get("price_per_kg", 0.0) or 0.0)
        except Exception:
            price = 0.0
        if price > 0.0:
            material_price_internal_update = True
            try:
                material_price_unit_combo.setCurrentText("EUR/kg")
                material_price_spin.setValue(price)
                material_price_auto_filled = True
            finally:
                material_price_internal_update = False

    def _normalize_counts(cuts_value: float, holes_value: float, slots_value: float) -> tuple[int, int, int, int]:
        holes_count = max(0, int(round(float(holes_value or 0.0))))
        slots_count = max(0, int(round(float(slots_value or 0.0))))
        total_cut_count = max(0, int(round(float(cuts_value or 0.0))))
        if total_cut_count < (holes_count + slots_count):
            total_cut_count = holes_count + slots_count
        outer_cut_count = max(0, total_cut_count - holes_count - slots_count)
        return total_cut_count, holes_count, slots_count, outer_cut_count

    def _estimate_profile_operations(path_txt: str) -> dict[str, Any]:
        family_txt = _infer_family(path_txt)
        section_txt = _infer_section(path_txt)
        section_norm = section_txt.upper().replace(" ", "")
        stem_lower = Path(path_txt).stem.lower()
        suffix = Path(path_txt).suffix.lower()
        try:
            cad_analysis = dict(analyze_profile_cut_features(path_txt) or {})
        except Exception as exc:
            cad_analysis = {"note": str(exc)}
        family_txt = _family_from_analysis(cad_analysis, family_txt)
        if not section_txt:
            section_txt = str(cad_analysis.get("section_label", "") or "").strip()
        text = _read_geometry_preview(path_txt)
        cuts = int(cad_analysis.get("cuts", 0) or 0)
        holes = int(cad_analysis.get("holes", 0) or 0)
        slots = int(cad_analysis.get("slots", 0) or 0)
        outer_cuts = int(cad_analysis.get("outer_cuts", max(0, cuts - holes - slots)) or 0)
        generic_cuts = int(cad_analysis.get("generic_cuts", 0) or 0)
        end_cut_count = int(cad_analysis.get("end_cut_count", 0) or 0)
        cut_length_m = float(cad_analysis.get("cut_length_m", 0.0) or 0.0)
        internal_cut_length_m = float(cad_analysis.get("feature_cut_length_m", 0.0) or 0.0)
        if internal_cut_length_m <= 0.0:
            internal_cut_length_m = (
                float(cad_analysis.get("hole_cut_length_mm", 0.0) or 0.0)
                + float(cad_analysis.get("slot_cut_length_mm", 0.0) or 0.0)
                + float(cad_analysis.get("generic_cut_length_mm", 0.0) or 0.0)
            ) / 1000.0
        if internal_cut_length_m > 0.0:
            cut_length_m = internal_cut_length_m
        if cad_analysis:
            cuts = int(max(0, holes + slots + generic_cuts))
            outer_cuts = int(max(0, generic_cuts))
        notes: list[str] = []
        inferred_material_family, inferred_material_subtype = _infer_material_from_profile(path_txt, text)
        cad_note = str(cad_analysis.get("note", "") or "").strip()
        if cad_note:
            notes.append(cad_note)
        if inferred_material_family:
            material_label = _laser_display_material_family(inferred_material_family) or inferred_material_family
            notes.append(
                f"material sugerido: {material_label}{f' / {inferred_material_subtype}' if inferred_material_subtype else ''}"
            )
        complex_tokens = ("mitra", "chanfro", "angulo", "bisel", "45")
        if any(token in stem_lower for token in complex_tokens) and not cad_analysis:
            notes.append("nome sugere cortes angulados; confirma cortes internos")
        if text and not cad_analysis:
            plane_count = len(re.findall(r"\bPLANE\b", text, re.IGNORECASE))
            cylindrical_count = len(re.findall(r"\bCYLINDRICAL_SURFACE\b", text, re.IGNORECASE))
            if suffix in {".step", ".stp"}:
                circle_count = len(re.findall(r"\bCIRCLE\s*\(", text, re.IGNORECASE))
                ellipse_count = len(re.findall(r"\bELLIPSE\s*\(", text, re.IGNORECASE))
                spline_count = len(re.findall(r"\bB_SPLINE_CURVE(?:_WITH_KNOTS)?\b", text, re.IGNORECASE))
                notes.append("leitura STEP textual aplicada")
            else:
                circle_count = len(re.findall(r"(^|,)\s*100\s*,", text, re.MULTILINE))
                ellipse_count = len(re.findall(r"(^|,)\s*104\s*,", text, re.MULTILINE))
                spline_count = len(re.findall(r"(^|,)\s*126\s*,", text, re.MULTILINE))
                notes.append("leitura IGES textual aplicada")
            if holes <= 0 and 0 < circle_count <= 12:
                holes = int(circle_count)
            elif any(token in stem_lower for token in ("furo", "furos", "hole", "holes")):
                holes = 1
            if slots <= 0 and 0 < ellipse_count <= 8:
                slots = int(ellipse_count)
            elif any(token in stem_lower for token in ("rasgo", "rasgos", "slot", "slots", "oblongo")):
                slots = 1
            if slots == 0 and 0 < spline_count <= 4 and any(token in stem_lower for token in ("slot", "rasgo", "oblongo")):
                slots = int(spline_count)
            if any(token in section_norm for token in ("IPE", "IPN", "HEA", "HEB", "HEM", "UPN", "UNP")):
                notes.append("perfil estrutural identificado; cortes terminais nao sao cobrados neste criterio")
            elif family_txt == "Tubo":
                if any(token in section_norm for token in ("RHS", "SHS")) or ("X" in section_norm):
                    notes.append("tubo/perfil retangular detetado; cobrados apenas cortes internos")
                elif re.search(r"(CHS|ROUND|REDONDO|DN\d+|D\d+|Ø)", section_norm) or cylindrical_count >= 4:
                    notes.append("tubo redondo identificado; cobrados apenas cortes internos")
            elif family_txt in {"Cantoneira", "Perfil", "Barra"}:
                notes.append("cobrados apenas cortes internos; extremidades ignoradas")
        elif not cad_analysis:
            notes.append("sem leitura textual disponivel; sem fallback de cortes exteriores")
        cuts, holes, slots, outer_cuts = _normalize_counts(cuts, holes, slots)
        if generic_cuts > 0 or end_cut_count > 0:
            notes.append(
                f"leitura atual: {int(cuts)} cortes internos cobrados; {int(end_cut_count)} cortes terminais ignorados"
            )
        else:
            notes.append(
                f"leitura atual: {int(cuts)} cortes internos = {int(holes)} furos + {int(slots)} rasgos + {int(outer_cuts)} outros internos"
            )
        if cut_length_m > 0:
            notes.append(f"comprimento medido: {cut_length_m:.3f} m")
        else:
            notes.append("comprimento nao medido automaticamente; valor editavel")
        return {
            "family": family_txt,
            "section": section_txt,
            "cuts": int(max(0, cuts)),
            "holes": int(max(0, holes)),
            "slots": int(max(0, slots)),
            "outer_cuts": int(max(0, outer_cuts)),
            "cut_length_m": round(max(0.0, cut_length_m), 4),
            "profile_length_m": float(cad_analysis.get("profile_length_m", 0.0) or 0.0) or _infer_profile_length_m(path_txt),
            "profile_kg_m": float(cad_analysis.get("profile_kg_m", 0.0) or 0.0),
            "thickness_mm": float(cad_analysis.get("thickness_mm_guess", 0.0) or 0.0),
            "material_family": inferred_material_family,
            "material_subtype": inferred_material_subtype,
            "cad_analysis": dict(cad_analysis or {}),
            "note": ". ".join(part for part in notes if part).strip(),
        }

    def _row_payload(row_index: int) -> dict[str, Any]:
        file_item = files_table.item(row_index, 0)
        path_txt = str(file_item.data(Qt.UserRole) if isinstance(file_item, QTableWidgetItem) else "").strip()
        cached_estimate = dict(file_item.data(Qt.UserRole + 1) if isinstance(file_item, QTableWidgetItem) else {} or {})
        family_combo = files_table.cellWidget(row_index, 1)
        section_edit = files_table.cellWidget(row_index, 2)
        qty_spin = files_table.cellWidget(row_index, 3)
        cuts_spin = files_table.cellWidget(row_index, 4)
        holes_spin = files_table.cellWidget(row_index, 5)
        slots_spin = files_table.cellWidget(row_index, 6)
        cut_length_spin = files_table.cellWidget(row_index, 7)
        profile_length_spin = files_table.cellWidget(row_index, 8)
        profile_kg_m_spin = files_table.cellWidget(row_index, 9)
        family_txt = family_combo.currentText().strip() if isinstance(family_combo, QComboBox) else str(cached_estimate.get("family", "Perfil") or "Perfil")
        section_txt = section_edit.text().strip() if isinstance(section_edit, QLineEdit) else str(cached_estimate.get("section", "") or "")
        qty = int(round(float(qty_spin.value() if isinstance(qty_spin, QDoubleSpinBox) else 0.0)))
        cut_length_m = float(cut_length_spin.value() if isinstance(cut_length_spin, QDoubleSpinBox) else float(cached_estimate.get("cut_length_m", 0.0) or 0.0))
        profile_length_m = float(profile_length_spin.value() if isinstance(profile_length_spin, QDoubleSpinBox) else float(cached_estimate.get("profile_length_m", 0.0) or 0.0))
        profile_kg_m = float(profile_kg_m_spin.value() if isinstance(profile_kg_m_spin, QDoubleSpinBox) else float(cached_estimate.get("profile_kg_m", 0.0) or 0.0))
        cuts, holes, slots, outer_cuts = _normalize_counts(
            cuts_spin.value() if isinstance(cuts_spin, QDoubleSpinBox) else 0.0,
            holes_spin.value() if isinstance(holes_spin, QDoubleSpinBox) else 0.0,
            slots_spin.value() if isinstance(slots_spin, QDoubleSpinBox) else 0.0,
        )
        material_family = _laser_canonical_material_family(material_combo.currentText().strip()) or material_combo.currentText().strip() or "Aco carbono"
        subtype_txt = subtype_combo.currentText().strip()
        material_price_unit = material_price_unit_combo.currentText().strip()
        if family_txt == "Tubo":
            material_price_unit = "EUR/m"
        return {
            "path": path_txt,
            "machine_name": machine_combo.currentText().strip(),
            "commercial_name": commercial_combo.currentText().strip(),
            "material": material_family,
            "material_subtype": subtype_txt,
            "gas": gas_combo.currentText().strip(),
            "thickness_mm": float(thickness_spin.value() or 0.0),
            "qtd": max(1, qty),
            "material_supplied_by_client": bool(customer_material_check.isChecked()),
            "material_fornecido_cliente": bool(customer_material_check.isChecked()),
            "profile_material_price": float(material_price_spin.value() or 0.0),
            "profile_material_price_unit": material_price_unit,
            "profile_length_m": max(0.0, profile_length_m),
            "profile_kg_m": max(0.0, profile_kg_m),
            "profile_family": family_txt or "Perfil",
            "section": section_txt,
            "cuts": cuts,
            "holes": holes,
            "slots": slots,
            "outer_cuts": outer_cuts,
            "cut_length_m_override": max(0.0, cut_length_m),
            "include_external_profile_cuts": False,
            "include_profile_event_rates": False,
            "profile_metrics": dict(cached_estimate.get("cad_analysis", {}) or {}),
        }

    def _recalc_total() -> None:
        total = 0.0
        status_parts: list[str] = []
        for row_index in range(files_table.rowCount()):
            try:
                analysis = dict(backend.profile_laser_quote_analyze(_row_payload(row_index)) or {})
            except Exception as exc:
                total_label.setText("Total estimado: -")
                status_label.setText(str(exc))
                return
            pricing = dict(analysis.get("pricing", {}) or {})
            metrics = dict(analysis.get("metrics", {}) or {})
            total += float(pricing.get("total_price", 0.0) or 0.0)
            end_count = int(metrics.get("raw_end_cut_count", metrics.get("end_cut_count", 0)) or 0)
            holes_count = int(metrics.get("hole_count", 0) or 0)
            slots_count = int(metrics.get("slot_count", 0) or 0)
            other_count = int(metrics.get("generic_cut_count", 0) or 0)
            cut_length_m = float(metrics.get("cut_length_m", 0.0) or 0.0)
            thickness_factor = float(metrics.get("thickness_rate_factor", 1.0) or 1.0)
            density_value = float(metrics.get("density_kg_m3", 0.0) or 0.0)
            material_cost = float(pricing.get("material_cost_unit", 0.0) or 0.0)
            material_price_value = float(pricing.get("profile_material_price", 0.0) or 0.0)
            material_price_unit = str(pricing.get("profile_material_price_unit", "") or "").strip().upper()
            cut_length_widget = files_table.cellWidget(row_index, 7)
            if isinstance(cut_length_widget, QDoubleSpinBox) and cut_length_widget.value() <= 0.0 and cut_length_m > 0.0:
                cut_length_widget.blockSignals(True)
                cut_length_widget.setValue(cut_length_m)
                cut_length_widget.blockSignals(False)
            status_parts.append(
                f"{int(metrics.get('cut_event_count', 0) or 0)} eventos cobrados = "
                f"{holes_count} furos + {slots_count} rasgos + {other_count} outros; "
                f"{end_count} terminais ignorados | {cut_length_m:.3f} m | "
                f"fator esp. x{thickness_factor:.2f} | dens. {density_value:.0f} kg/m3 | "
                f"MP {_fmt_eur(material_cost)} @ {material_price_value:.4f} {material_price_unit}"
            )
        total_label.setText(f"Total estimado: {_fmt_eur(total)}")
        status_label.setText(status_parts[0] if status_parts else "Usa as tabelas do laser com contagem de eventos STEP/IGS, sem duplicar furos.")

    def _add_file_row(path_txt: str) -> None:
        estimate = _estimate_profile_operations(path_txt)
        row_index = files_table.rowCount()
        files_table.insertRow(row_index)
        estimated_thickness = float(estimate.get("thickness_mm", 0.0) or 0.0)
        if row_index == 0 and estimated_thickness > 0.0:
            thickness_spin.blockSignals(True)
            thickness_spin.setValue(estimated_thickness)
            thickness_spin.blockSignals(False)
        if row_index == 0 and str(estimate.get("material_family", "") or "").strip():
            material_display = _laser_display_material_family(str(estimate.get("material_family", "") or "").strip())
            if material_display:
                material_combo.setCurrentText(material_display)
                _refresh_subtypes()
            subtype_txt = str(estimate.get("material_subtype", "") or "").strip()
            if subtype_txt:
                subtype_combo.setCurrentText(subtype_txt)
                _refresh_gases()
        file_item = QTableWidgetItem(Path(path_txt).name)
        file_item.setData(Qt.UserRole, str(path_txt))
        file_item.setData(Qt.UserRole + 1, dict(estimate))
        file_item.setData(Qt.UserRole + 2, {})
        file_item.setToolTip(f"{path_txt}\n{str(estimate.get('note', '') or '').strip()}")
        files_table.setItem(row_index, 0, file_item)

        family_combo = QComboBox()
        family_combo.addItems(families)
        family_combo.setCurrentText(str(estimate.get("family", "") or _infer_family(path_txt)))
        section_edit = QLineEdit(str(estimate.get("section", "") or _infer_section(path_txt)))
        qty_spin = _make_int_spin(1)
        cuts_spin = _make_int_spin(int(estimate.get("cuts", 2) or 2))
        holes_spin = _make_int_spin(int(estimate.get("holes", 0) or 0))
        slots_spin = _make_int_spin(int(estimate.get("slots", 0) or 0))
        cut_length_spin = QDoubleSpinBox()
        cut_length_spin.setRange(0.0, 1000000.0)
        cut_length_spin.setDecimals(3)
        cut_length_spin.setSingleStep(0.1)
        cut_length_spin.setValue(float(estimate.get("cut_length_m", 0.0) or 0.0))
        cut_length_spin.setMinimumHeight(32)
        cut_length_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
        profile_length_spin = QDoubleSpinBox()
        profile_length_spin.setRange(0.0, 1000000.0)
        profile_length_spin.setDecimals(3)
        profile_length_spin.setSingleStep(0.1)
        profile_length_spin.setValue(float(estimate.get("profile_length_m", 0.0) or 0.0))
        profile_length_spin.setMinimumHeight(32)
        profile_length_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
        profile_kg_m_spin = QDoubleSpinBox()
        profile_kg_m_spin.setRange(0.0, 1000000.0)
        profile_kg_m_spin.setDecimals(4)
        profile_kg_m_spin.setSingleStep(0.1)
        profile_kg_m_spin.setValue(float(estimate.get("profile_kg_m", 0.0) or 0.0))
        profile_kg_m_spin.setMinimumHeight(32)
        profile_kg_m_spin.setStyleSheet("QDoubleSpinBox { padding: 4px 8px; font-size: 12px; }")
        family_combo.setMinimumHeight(32)
        section_edit.setMinimumHeight(32)
        files_table.setCellWidget(row_index, 1, family_combo)
        files_table.setCellWidget(row_index, 2, section_edit)
        files_table.setCellWidget(row_index, 3, qty_spin)
        files_table.setCellWidget(row_index, 4, cuts_spin)
        files_table.setCellWidget(row_index, 5, holes_spin)
        files_table.setCellWidget(row_index, 6, slots_spin)
        files_table.setCellWidget(row_index, 7, cut_length_spin)
        files_table.setCellWidget(row_index, 8, profile_length_spin)
        files_table.setCellWidget(row_index, 9, profile_kg_m_spin)
        files_table.setRowHeight(row_index, 42)
        analysis_note = str(estimate.get("note", "") or "").strip()
        for widget in (family_combo, section_edit, qty_spin, cuts_spin, holes_spin, slots_spin, cut_length_spin, profile_length_spin, profile_kg_m_spin):
            widget.setToolTip(analysis_note)
        for widget in (qty_spin, cuts_spin, holes_spin, slots_spin, cut_length_spin, profile_length_spin, profile_kg_m_spin):
            widget.valueChanged.connect(_recalc_total)
        family_combo.currentTextChanged.connect(_sync_material_unit_for_family)
        family_combo.currentTextChanged.connect(_recalc_total)
        section_edit.textChanged.connect(_recalc_total)
        files_table.selectRow(row_index)
        _sync_material_unit_for_family()
        _update_step_preview()

    def _pick_files() -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            dialog,
            "Selecionar ficheiros STEP/IGS",
            "",
            "Modelos 3D (*.step *.stp *.igs *.iges);;Todos (*.*)",
        )
        for path_txt in paths:
            if path_txt:
                _add_file_row(path_txt)
        _recalc_total()

    def _remove_selected() -> None:
        selected_rows = sorted({item.row() for item in files_table.selectedItems() if item is not None}, reverse=True)
        for row_index in selected_rows:
            files_table.removeRow(row_index)
        _recalc_total()
        _update_step_preview()

    def _on_material_price_changed(_value: float) -> None:
        nonlocal material_price_user_touched, material_price_auto_filled
        if not material_price_internal_update:
            material_price_user_touched = True
            material_price_auto_filled = False
        _recalc_total()

    def _on_material_unit_changed(_text: str) -> None:
        if not material_price_internal_update and _selected_profile_family() == "Tubo":
            _sync_material_unit_for_family()
        _recalc_total()

    def _accept() -> None:
        nonlocal result_lines
        if files_table.rowCount() == 0:
            QMessageBox.warning(dialog, "STEP/IGS", "Adiciona pelo menos um ficheiro STEP/IGS.")
            return
        lines: list[dict[str, Any]] = []
        for row_index in range(files_table.rowCount()):
            payload = _row_payload(row_index)
            if int(payload.get("qtd", 0) or 0) <= 0:
                continue
            try:
                result = dict(backend.profile_laser_quote_build_line(payload) or {})
            except Exception as exc:
                QMessageBox.critical(dialog, "STEP/IGS", str(exc))
                return
            analysis = dict(result.get("analysis", {}) or {})
            line = dict(result.get("line", {}) or {})
            if not line:
                continue
            metrics = dict(analysis.get("metrics", {}) or {})
            line["tipo_item"] = backend.ORC_LINE_TYPE_SERVICE
            line["produto_unid"] = "UN"
            line["line_origin"] = "step_igs_profile_laser"
            line["summary_html"] = (
                f"{str(payload.get('profile_family', 'Perfil') or 'Perfil')} | "
                f"cortes internos {int(metrics.get('cut_event_count', 0) or 0)} | "
                f"m corte {float(metrics.get('cut_length_m', 0.0) or 0.0):.3f} | "
                f"furos {int(metrics.get('hole_count', 0) or 0)} | "
                f"rasgos {int(metrics.get('slot_count', 0) or 0)} | "
                f"terminais ignorados {int(metrics.get('raw_end_cut_count', 0) or 0)}"
            )
            lines.append(line)
        if not lines:
            QMessageBox.warning(dialog, "STEP/IGS", "Nao foi possivel gerar linhas com os dados atuais.")
            return
        result_lines = [dict(row or {}) for row in lines]
        dialog.accept()

    add_files_btn.clicked.connect(_pick_files)
    remove_file_btn.clicked.connect(_remove_selected)
    configure_btn.clicked.connect(_configure_profiles)
    buttons.accepted.connect(_accept)
    buttons.rejected.connect(dialog.reject)
    files_table.itemSelectionChanged.connect(_update_step_preview)
    files_table.itemSelectionChanged.connect(_sync_material_unit_for_family)
    machine_combo.currentTextChanged.connect(_refresh_materials)
    machine_combo.currentTextChanged.connect(_recalc_total)
    commercial_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
    commercial_combo.currentTextChanged.connect(_recalc_total)
    material_combo.currentTextChanged.connect(_refresh_subtypes)
    material_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
    material_combo.currentTextChanged.connect(_recalc_total)
    subtype_combo.currentTextChanged.connect(lambda _txt: _refresh_material_price_default())
    subtype_combo.currentTextChanged.connect(_recalc_total)
    gas_combo.currentTextChanged.connect(_recalc_total)
    thickness_spin.valueChanged.connect(_recalc_total)
    material_price_spin.valueChanged.connect(_on_material_price_changed)
    material_price_unit_combo.currentTextChanged.connect(_on_material_unit_changed)
    customer_material_check.toggled.connect(lambda _checked: _sync_material_cost_controls())
    customer_material_check.toggled.connect(_recalc_total)
    _refresh_machine_and_commercial()
    _refresh_materials()
    _refresh_material_price_default()
    _sync_material_cost_controls()
    _recalc_total()
    _update_step_preview()
    if dialog.exec() != QDialog.Accepted:
        return None
    return [dict(row or {}) for row in result_lines]
