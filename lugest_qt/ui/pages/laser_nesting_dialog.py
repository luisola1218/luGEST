from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, QPoint, QPointF, QRect, QRectF, QSize, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QGuiApplication, QColor, QBrush, QIcon, QPainter, QPainterPath, QPen, QPixmap, QPolygonF
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QFormLayout,
    QGraphicsScene,
    QGraphicsView,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QToolButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from ...services.laser_nesting import _sheet_overlap_diagnostics, default_nesting_options, default_sheet_profiles, grouped_laser_rows, nest_parts
from lugest_core.laser.quote_engine import estimate_laser_quote
from ..widgets import CardFrame
from .laser_quote_dialogs import LaserSettingsDialog, _fmt_num, _spin


STRATEGY_LABELS = {
    "longest-side": "Maior lado",
    "area": "Maior area",
    "height-first": "Altura primeiro",
    "width-first": "Largura primeiro",
    "compact": "Compactação",
    "shape-longest-side": "Contorno + maior lado",
    "shape-area": "Contorno + maior area",
    "shape-height-first": "Contorno + altura primeiro",
    "shape-width-first": "Contorno + largura primeiro",
    "shape-compact": "Contorno + compactação",
    "shape": "Contorno",
}

SOURCE_LABELS = {
    "purchase": "Compra",
    "stock": "Stock",
    "retalho": "Retalho",
}

PRIORITY_LABELS = {
    -1: "Baixa",
    0: "Normal",
    1: "Alta",
    2: "Crítica",
}

PREVIEW_PALETTE = ["#dbeafe", "#dcfce7", "#fef3c7", "#ffe4e6", "#ede9fe", "#cffafe", "#e2e8f0"]
WIZARD_STEPS = [
    (0, 0, "1. Definir", "Regras e formatos"),
    (0, 1, "1. Definir", "Peças e stock"),
    (1, 0, "2. Nest", "Cenários"),
    (1, 1, "2. Nest", "Layouts / chapas"),
    (2, 0, "3. Corte / orçamento", "Pendentes / avisos"),
    (2, 1, "3. Corte / orçamento", "Custos / decisão"),
]
SECTION_DESCRIPTIONS = {
    0: "Ajusta rapidamente formatos, stock e parâmetros essenciais do estudo.",
    1: "Compara cenários e valida visualmente o melhor aproveitamento de chapa.",
    2: "Fecha o estudo com pendentes, custos e decisão final.",
}


class NestingWorker(QObject):
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(self, payload: dict[str, Any]) -> None:
        super().__init__()
        self.payload = dict(payload or {})

    def run(self) -> None:
        try:
            result = nest_parts(**self.payload)
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.finished.emit(dict(result or {}))


# Normaliza a linguagem do wizard e dos estados para evitar texto com
# encoding partido, mantendo a estrutura já existente do diálogo.
STRATEGY_LABELS = {
    "longest-side": "Maior lado",
    "area": "Maior area",
    "height-first": "Altura primeiro",
    "width-first": "Largura primeiro",
    "compact": "Compactacao",
    "shape-longest-side": "Contorno + maior lado",
    "shape-area": "Contorno + maior area",
    "shape-height-first": "Contorno + altura primeiro",
    "shape-width-first": "Contorno + largura primeiro",
    "shape-compact": "Contorno + compactacao",
    "shape": "Contorno",
}
PRIORITY_LABELS = {
    -1: "Baixa",
    0: "Normal",
    1: "Alta",
    2: "Critica",
}
WIZARD_STEPS = [
    (0, 0, "1. Definir", "Parametros e chapa"),
    (0, 1, "1. Definir", "Pecas e stock"),
    (1, 0, "2. Nest", "Cenarios"),
    (1, 1, "2. Nest", "Layouts e chapas"),
    (2, 0, "3. Corte / Orcamento", "Pendentes e avisos"),
    (2, 1, "3. Corte / Orcamento", "Custos e decisao"),
]
SECTION_DESCRIPTIONS = {
    0: "Define chapa, stock, retalhos e parametros base do estudo.",
    1: "Compara cenarios e valida visualmente o melhor aproveitamento de chapa.",
    2: "Fecha o estudo com pendentes, custos tecnicos e decisao final.",
}


def _repair_mojibake_text(text: str) -> str:
    fixed = str(text or "")
    replacements = {
        "Ã§": "c",
        "Ã£": "a",
        "Ã¡": "a",
        "Ã ": "a",
        "Ã©": "e",
        "Ãª": "e",
        "Ã­": "i",
        "Ã³": "o",
        "Ã´": "o",
        "Ãº": "u",
        "Ã‰": "E",
        "Ã‡": "C",
        "Â°": "°",
        "ÃƒÂ¢": "a",
        "ÃƒÂ§": "c",
        "ÃƒÂ£": "a",
        "ÃƒÂ¡": "a",
        "ÃƒÂ©": "e",
        "ÃƒÂ³": "o",
        "ÃƒÂµ": "o",
        "ÃƒÂ­": "i",
        "ÃƒÂº": "u",
        "nÃ£o": "nao",
        "NÃ£o": "Nao",
        "orÃ§": "orc",
        "OrÃ§": "Orc",
        "anÃ¡": "ana",
        "cenÃ¡": "cena",
        "geometrÃ": "geometri",
        "matÃ©": "mate",
        "MÃ©": "Me",
        "EstratÃ©gia": "Estrategia",
        "CompactaÃ§Ã£o": "Compactacao",
        "CompactaÃ§Ã£o": "Compactacao",
        "UtilizaÃ§Ã£o": "Utilizacao",
        "OcupaÃ§Ã£o": "Ocupacao",
        "ComparaÃ§Ã£o": "Comparacao",
        "RecomendaÃ§Ã£o": "Recomendacao",
        "DiagnÃ³stico": "Diagnostico",
        "observaÃ§Ãµes": "observacoes",
        "PeÃ§as": "Pecas",
        "peÃ§as": "pecas",
        "CenÃ¡rios": "Cenarios",
        "decisÃ£o": "decisao",
        "ResoluÃ§Ã£o": "Resolucao",
        "TolerÃ¢ncia": "Tolerancia",
        "ReduÃ§Ã£o": "Reducao",
        "rotaÃ§Ã£o": "rotacao",
        "RotaÃ§Ã£o": "Rotacao",
        "prÃ©-visualizar": "pre-visualizar",
        "colisÃ£o(Ãµes)": "colisao(oes)",
        "vÃ¡lido(s)": "valido(s)",
        "otimizaÃ§Ã£o": "otimizacao",
        "automÃ¡tica": "automatica",
        "mÃ¡quina": "maquina",
        "MÃ¡quina": "Maquina",
        "Ã¡rea": "area",
        "Ãrea": "Area",
        "matÃ©ria": "materia",
        "Ãºteis": "uteis",
        "invÃ¡lida": "invalida",
        "CÃ³pia": "Copia",
    }
    for old, new in replacements.items():
        fixed = fixed.replace(old, new)
    return fixed

def _fmt_eur(value: Any) -> str:
    try:
        amount = float(value or 0.0)
    except Exception:
        amount = 0.0
    return f"{amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


class NestPreviewView(QGraphicsView):
    def __init__(self, scene: QGraphicsScene, parent=None) -> None:
        super().__init__(scene, parent)
        self._zoom_level = 0
        self._fit_pending = False
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setBackgroundBrush(QColor("#f8fafc"))
        self.setAlignment(Qt.AlignCenter)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

    def request_fit(self) -> None:
        self._fit_pending = True
        QTimer.singleShot(0, self._apply_pending_fit)

    def _apply_pending_fit(self) -> None:
        scene = self.scene()
        viewport = self.viewport()
        if scene is None or scene.sceneRect().isNull() or viewport is None:
            return
        if not self.isVisible() or viewport.width() < 24 or viewport.height() < 24:
            self._fit_pending = True
            return
        self.resetTransform()
        self._zoom_level = 0
        self.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)
        self.centerOn(scene.sceneRect().center())
        self._fit_pending = False

    def fit_scene(self) -> None:
        self.request_fit()

    def wheelEvent(self, event) -> None:  # type: ignore[override]
        delta = int(event.angleDelta().y())
        if delta == 0:
            super().wheelEvent(event)
            return
        step = 1 if delta > 0 else -1
        next_level = self._zoom_level + step
        if next_level < -6 or next_level > 20:
            event.accept()
            return
        factor = 1.15 if step > 0 else 1.0 / 1.15
        self._zoom_level = next_level
        self.scale(factor, factor)
        event.accept()

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        self._apply_pending_fit()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._apply_pending_fit()

    def mouseDoubleClickEvent(self, event) -> None:  # type: ignore[override]
        self.request_fit()
        event.accept()


class SheetProfilesDialog(QDialog):
    def __init__(self, profiles: list[dict[str, Any]], parent=None) -> None:
        super().__init__(parent)
        self._profiles = [dict(row or {}) for row in list(profiles or [])]
        self.setWindowTitle("Formatos de chapa")
        self.resize(720, 420)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        info = QLabel("Define os formatos de chapa standard usados como fallback no nesting manual e automático.")
        info.setWordWrap(True)
        info.setProperty("role", "muted")
        root.addWidget(info)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Nome", "Largura (mm)", "Altura (mm)"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        root.addWidget(self.table, 1)

        actions = QHBoxLayout()
        add_btn = QPushButton("Adicionar")
        add_btn.setProperty("variant", "secondary")
        add_btn.clicked.connect(self._add_row)
        remove_btn = QPushButton("Remover")
        remove_btn.setProperty("variant", "secondary")
        remove_btn.clicked.connect(self._remove_selected)
        reset_btn = QPushButton("Repor standard")
        reset_btn.setProperty("variant", "secondary")
        reset_btn.clicked.connect(self._reset_defaults)
        actions.addWidget(add_btn)
        actions.addWidget(remove_btn)
        actions.addWidget(reset_btn)
        actions.addStretch(1)
        root.addLayout(actions)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._load_profiles(self._profiles or default_sheet_profiles())

    def _load_profiles(self, profiles: list[dict[str, Any]]) -> None:
        self.table.setRowCount(0)
        for row in list(profiles or []):
            self._add_row(dict(row or {}))

    def _add_row(self, profile: dict[str, Any] | None = None) -> None:
        row = dict(profile or {})
        row_index = self.table.rowCount()
        self.table.insertRow(row_index)
        values = [
            str(row.get("name", "") or "").strip(),
            _fmt_num(row.get("width_mm", 0), 3) if row else "",
            _fmt_num(row.get("height_mm", 0), 3) if row else "",
        ]
        for col_index, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setTextAlignment(int((Qt.AlignLeft if col_index == 0 else Qt.AlignCenter) | Qt.AlignVCenter))
            self.table.setItem(row_index, col_index, item)
        if not profile:
            self.table.setCurrentCell(row_index, 0)

    def _remove_selected(self) -> None:
        selected = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        if not selected and self.table.rowCount() > 0:
            selected = [self.table.rowCount() - 1]
        for row_index in selected:
            self.table.removeRow(row_index)

    def _reset_defaults(self) -> None:
        self._load_profiles(default_sheet_profiles())

    def _read_profiles(self) -> list[dict[str, Any]]:
        profiles: list[dict[str, Any]] = []
        for row_index in range(self.table.rowCount()):
            name_item = self.table.item(row_index, 0)
            width_item = self.table.item(row_index, 1)
            height_item = self.table.item(row_index, 2)
            name = str(name_item.text() if name_item is not None else "").strip()
            width_txt = str(width_item.text() if width_item is not None else "").strip().replace(",", ".")
            height_txt = str(height_item.text() if height_item is not None else "").strip().replace(",", ".")
            if not name and not width_txt and not height_txt:
                continue
            try:
                width_mm = float(width_txt or 0.0)
                height_mm = float(height_txt or 0.0)
            except Exception:
                raise ValueError(f"Linha {row_index + 1}: largura e altura devem ser numéricas.")
            if width_mm <= 0.0 or height_mm <= 0.0:
                raise ValueError(f"Linha {row_index + 1}: indica largura e altura acima de zero.")
            profiles.append(
                {
                    "name": name or f"{width_mm:g} x {height_mm:g}",
                    "width_mm": round(width_mm, 3),
                    "height_mm": round(height_mm, 3),
                }
            )
        if not profiles:
            raise ValueError("Define pelo menos um formato de chapa.")
        return profiles

    def _accept(self) -> None:
        try:
            self._profiles = self._read_profiles()
        except Exception as exc:
            QMessageBox.warning(self, "Formatos de chapa", str(exc))
            return
        self.accept()

    def result_profiles(self) -> list[dict[str, Any]]:
        return [dict(row or {}) for row in list(self._profiles or [])]


class LaserNestingDialog(QDialog):
    def __init__(self, backend, rows: list[dict[str, Any]], parent=None, quote_number: str = "") -> None:
        super().__init__(parent)
        self.backend = backend
        self.rows = [dict(row or {}) for row in list(rows or [])]
        self.quote_number = str(quote_number or "").strip()
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self.nesting_options = default_nesting_options(self.settings)
        self.groups = grouped_laser_rows(self.rows)
        self.result_data: dict[str, Any] = {}
        self.current_group_rows: list[dict[str, Any]] = []
        self.sheet_profiles = list(self.nesting_options.get("sheet_profiles", []) or [])
        self.stock_candidates: list[dict[str, Any]] = []
        self.part_geometry_cache: dict[str, dict[str, Any]] = {}
        self.part_rotation_overrides: dict[str, str] = {}
        self.part_priority_overrides: dict[str, int] = {}
        self.saved_studies: dict[str, dict[str, Any]] = {}
        self._window_fitted_to_screen = False
        self._stored_window_geometry = QRect()
        self.setWindowTitle("Nesting Laser")
        self.setMinimumSize(980, 620)
        self.resize(1540, 980)
        screen = self.screen()
        if screen is not None:
            available = screen.availableGeometry()
            target_width = min(max(1280, int(available.width() * 0.94)), max(1280, available.width() - 24))
            target_height = min(max(820, int(available.height() * 0.90)), max(820, available.height() - 24))
            self.resize(target_width, target_height)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        self.header_card = CardFrame()
        header_card = self.header_card
        header_card.set_tone("default")
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(12)
        header_text = QVBoxLayout()
        header_text.setSpacing(4)
        header_title = QLabel("Nesting Laser")
        header_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        header_subtitle = QLabel("")
        header_subtitle.setProperty("role", "muted")
        header_subtitle.setWordWrap(True)
        self.study_status_label = QLabel("")
        self.study_status_label.setWordWrap(True)
        self.header_title_label = header_title
        self.header_subtitle_label = header_subtitle
        header_text.addWidget(header_title)
        header_text.addWidget(header_subtitle)
        header_text.addWidget(self.study_status_label)
        header_layout.addLayout(header_text, 1)
        self.window_toggle_btn = QToolButton()
        self.window_toggle_btn.setCursor(Qt.PointingHandCursor)
        self.window_toggle_btn.setAutoRaise(False)
        self.window_toggle_btn.clicked.connect(self._toggle_window_mode)
        self.window_toggle_btn.setStyleSheet(
            "QToolButton {"
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #e8eef5);"
            "color: #24384f; border: 1px solid #b8c7d8; border-bottom: 3px solid #9fb0c3;"
            "border-radius: 18px; padding: 6px 14px; font-size: 12px; font-weight: 800;"
            "}"
            "QToolButton:hover {background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #dde7f0);}"
        )
        header_layout.addWidget(self.window_toggle_btn, 0, Qt.AlignTop)
        root.addWidget(header_card)

        self.stepper_card = CardFrame()
        stepper_card = self.stepper_card
        stepper_card.setStyleSheet(
            """
            QFrame {
                background: #f8fbff;
                border: 1px solid #bfd0e1;
                border-radius: 12px;
            }
            """
        )
        stepper_layout = QVBoxLayout(stepper_card)
        stepper_layout.setContentsMargins(12, 8, 12, 8)
        stepper_layout.setSpacing(6)
        stepper_top = QHBoxLayout()
        stepper_top.setSpacing(10)
        self.group_badge_label = QLabel("Grupo não definido")
        self.group_badge_label.setStyleSheet(
            "padding: 6px 12px; border-radius: 999px; background: #eef4ff; color: #274c77; "
            "border: 1px solid #bfd2ea; font-size: 12px; font-weight: 900;"
        )
        stepper_top.addWidget(self.group_badge_label, 0)
        stepper_top.addStretch(1)
        self.major_step_buttons: list[QPushButton] = []
        for section_index, title in enumerate(("1  Definir", "2  Nest", "3  Corte / orçamento")):
            step_btn = QPushButton(title)
            step_btn.setCheckable(True)
            step_btn.setMinimumHeight(34)
            step_btn.setCursor(Qt.PointingHandCursor)
            step_btn.clicked.connect(lambda _checked=False, idx=section_index: self._jump_to_major_step(idx))
            self.major_step_buttons.append(step_btn)
            stepper_top.addWidget(step_btn)
        stepper_layout.addLayout(stepper_top)
        self.page_title_label = QLabel("Regras e formatos")
        self.page_title_label.setStyleSheet("font-size: 17px; font-weight: 900; color: #0f172a;")
        self.page_subtitle_label = QLabel("")
        self.page_subtitle_label.setProperty("role", "muted")
        self.page_subtitle_label.setWordWrap(True)
        self.page_subtitle_label.setStyleSheet("font-size: 11px; color: #486581;")
        stepper_layout.addWidget(self.page_title_label)
        self.page_subtitle_label.hide()
        root.addWidget(stepper_card)

        self.group_combo = QComboBox()
        for group in self.groups:
            self.group_combo.addItem(str(group.get("label", "")), dict(group))

        self.sheet_combo = QComboBox()
        self.spacing_spin = _spin(1, 0.0, 100.0, float(self.nesting_options.get("default_part_spacing_mm", 8.0) or 8.0), 0.5)
        self.edge_spin = _spin(1, 0.0, 200.0, float(self.nesting_options.get("default_edge_margin_mm", 8.0) or 8.0), 0.5)
        self.rotate_check = QCheckBox("Permitir rotação automática")
        self.rotate_check.setChecked(bool(self.nesting_options.get("allow_rotate", True)))
        self.rotate_check.setToolTip("Afeta apenas peças em modo Automático. Peças fixadas a 0° ou 90° respeitam essa escolha.")
        self.mirror_check = QCheckBox("Permitir espelhar peças")
        self.mirror_check.setChecked(bool(self.nesting_options.get("allow_mirror", True)))
        self.mirror_check.setToolTip("Testa também a geometria espelhada no nesting por contorno para melhorar aproveitamento de chapa.")
        self.free_angle_check = QCheckBox("Rotações livres")
        self.free_angle_check.setChecked(bool(self.nesting_options.get("free_angle_rotation", True)))
        self.free_angle_check.setToolTip("Testa ângulos intermédios no nesting por contorno. É mais lento, mas pode encaixar muito melhor.")
        self.auto_sheet_check = QCheckBox("Escolher melhor formato automaticamente")
        self.auto_sheet_check.setChecked(bool(self.nesting_options.get("auto_select_sheet", False)))
        self.use_stock_check = QCheckBox("Usar stock e retalhos primeiro")
        self.use_stock_check.setChecked(bool(self.nesting_options.get("use_stock_first", False)))
        self.allow_purchase_check = QCheckBox("Permitir compra complementar")
        self.allow_purchase_check.setChecked(bool(self.nesting_options.get("allow_purchase_fallback", True)))
        self.shape_check = QCheckBox("Nesting por contorno")
        self.shape_check.setChecked(bool(self.nesting_options.get("shape_aware", True)))
        self.shape_strict_check = QCheckBox("Forçar só contorno real")
        self.shape_strict_check.setChecked(bool(self.nesting_options.get("shape_strict", False)))
        self.shape_grid_spin = _spin(1, 2.0, 100.0, float(self.nesting_options.get("shape_grid_mm", 10.0) or 10.0), 0.5)
        self.common_line_check = QCheckBox("Estimar common-line")
        self.common_line_check.setChecked(bool(self.nesting_options.get("common_line_estimate", True)))
        self.common_line_tol_spin = _spin(1, 0.0, 10.0, float(self.nesting_options.get("common_line_tolerance_mm", 1.0) or 1.0), 0.1)
        self.lead_opt_check = QCheckBox("Otimizar lead-ins / lead-outs")
        self.lead_opt_check.setChecked(bool(self.nesting_options.get("lead_optimization", True)))
        self.lead_opt_pct_spin = _spin(1, 0.0, 50.0, float(self.nesting_options.get("lead_optimization_pct", 8.0) or 8.0), 0.5)

        sheet_btn = QPushButton("Biblioteca de chapas")
        sheet_btn.setProperty("variant", "secondary")
        sheet_btn.clicked.connect(self._configure_sheet_profiles)
        self.sheet_library_btn = sheet_btn
        config_btn = QPushButton("Parâmetros DXF e custos")
        config_btn.setProperty("variant", "secondary")
        config_btn.clicked.connect(self._configure_laser_settings)
        self.dxf_settings_btn = config_btn

        self.section_stack = QTabWidget()
        self.section_stack.setDocumentMode(True)
        self.section_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.section_stack.setStyleSheet(
            """
            QTabWidget::pane {
                border: 1px solid #9fb6cd;
                background: #f7fbff;
                border-radius: 10px;
                top: -1px;
            }
            QTabBar::tab {
                background: #dbe8f5;
                border: 1px solid #9fb6cd;
                padding: 8px 16px;
                min-width: 122px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                color: #18324d;
                font-weight: 800;
            }
            QTabBar::tab:hover {
                background: #edf4fb;
            }
            QTabBar::tab:selected {
                background: #0f4c81;
                color: #ffffff;
                border-color: #0f4c81;
            }
            """
        )
        self.section_stack.tabBar().hide()
        self.body_scroll = QScrollArea()
        self.body_scroll.setWidgetResizable(True)
        self.body_scroll.setFrameShape(QFrame.NoFrame)
        self.body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.body_scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        body_host = QWidget()
        self.body_layout = QVBoxLayout(body_host)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(10)
        self.body_scroll.setWidget(body_host)
        root.addWidget(self.body_scroll, 1)

        config_page = QWidget()
        config_page_layout = QVBoxLayout(config_page)
        config_page_layout.setContentsMargins(0, 0, 0, 0)
        config_page_layout.setSpacing(10)
        top_card = CardFrame()
        top_layout = QGridLayout(top_card)
        top_layout.setContentsMargins(14, 12, 14, 12)
        top_layout.setHorizontalSpacing(14)
        top_layout.setVerticalSpacing(10)
        top_layout.addWidget(QLabel("Grupo de material"), 0, 0)
        top_layout.addWidget(QLabel("Formato standard / compra"), 0, 1)
        top_layout.addWidget(QLabel("Margem entre peças (mm)"), 0, 2)
        top_layout.addWidget(QLabel("Margem à borda (mm)"), 0, 3)
        top_layout.addWidget(self.group_combo, 1, 0)
        top_layout.addWidget(self.sheet_combo, 1, 1)
        top_layout.addWidget(self.spacing_spin, 1, 2)
        top_layout.addWidget(self.edge_spin, 1, 3)
        top_layout.addWidget(self.rotate_check, 2, 0)
        top_layout.addWidget(self.mirror_check, 2, 1)
        top_layout.addWidget(self.auto_sheet_check, 2, 2)
        top_layout.addWidget(self.use_stock_check, 2, 3)
        top_layout.addWidget(self.shape_check, 3, 0)
        top_layout.addWidget(self.free_angle_check, 3, 1)
        top_layout.addWidget(QLabel("Resolução da grelha (mm)"), 3, 2)
        top_layout.addWidget(self.shape_grid_spin, 3, 3)
        top_layout.addWidget(self.allow_purchase_check, 4, 0)
        top_layout.addWidget(self.shape_strict_check, 4, 1)
        top_layout.addWidget(sheet_btn, 4, 2)
        top_layout.addWidget(config_btn, 4, 3)
        top_layout.addWidget(self.common_line_check, 5, 0)
        top_layout.addWidget(QLabel("TolerÃ¢ncia common-line (mm)"), 5, 1)
        top_layout.addWidget(self.common_line_tol_spin, 5, 2)
        top_layout.addWidget(self.lead_opt_check, 6, 0)
        top_layout.addWidget(QLabel("ReduÃ§Ã£o lead-ins %"), 6, 1)
        top_layout.addWidget(self.lead_opt_pct_spin, 6, 2)
        top_layout.setColumnStretch(0, 3)
        top_layout.setColumnStretch(1, 3)
        top_layout.setColumnStretch(2, 3)
        top_layout.setColumnStretch(3, 3)
        config_page_layout.addWidget(top_card)

        parts_card = CardFrame()
        parts_card.setStyleSheet(
            "CardFrame { background: #fbfdff; border: 1px solid #c8d7e6; border-radius: 16px; }"
        )
        parts_layout = QVBoxLayout(parts_card)
        parts_layout.setContentsMargins(12, 10, 12, 10)
        parts_layout.setSpacing(8)
        parts_header = QHBoxLayout()
        parts_header.setSpacing(8)
        parts_title = QLabel("Peças do grupo selecionado")
        parts_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.parts_status_label = QLabel("Sem otimização ainda.")
        self.parts_status_label.setProperty("role", "muted")
        self.parts_status_label.setStyleSheet("font-size: 11px;")
        parts_header.addWidget(parts_title)
        parts_header.addStretch(1)
        parts_header.addWidget(self.parts_status_label)
        self.parts_table = QTableWidget(0, 9)
        self.parts_table.setHorizontalHeaderLabels(["Vista", "Ref. externa", "Descricao", "Prog.", "Qtd", "Prioridade", "Caixa", "Area real", "Rotação"])
        self.parts_table.verticalHeader().setVisible(False)
        self.parts_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.parts_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.parts_table.setAlternatingRowColors(True)
        self.parts_table.setWordWrap(False)
        self.parts_table.setStyleSheet(
            """
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d8e2ec;
                border-radius: 12px;
                alternate-background-color: #f8fbff;
                gridline-color: #d7e3ee;
                font-size: 12px;
            }
            QHeaderView::section {
                background: #08104f;
                color: #ffffff;
                padding: 10px 8px;
                border: none;
                font-size: 12px;
                font-weight: 900;
            }
            QTableWidget::item {
                padding: 6px 8px;
            }
            """
        )
        header = self.parts_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        for col in (3, 4, 5, 6, 7, 8):
            header.setSectionResizeMode(col, QHeaderView.Fixed)
        self.parts_table.setColumnWidth(0, 188)
        self.parts_table.setColumnWidth(3, 74)
        self.parts_table.setColumnWidth(4, 54)
        self.parts_table.setColumnWidth(5, 150)
        self.parts_table.setColumnWidth(6, 130)
        self.parts_table.setColumnWidth(7, 88)
        self.parts_table.setColumnWidth(8, 168)
        self.parts_table.verticalHeader().setDefaultSectionSize(96)
        parts_card.setMinimumHeight(0)
        self.parts_table.setMinimumHeight(190)
        parts_layout.addLayout(parts_header)
        parts_layout.addWidget(self.parts_table, 1)

        self.stock_card = CardFrame()
        stock_layout = QVBoxLayout(self.stock_card)
        stock_layout.setContentsMargins(12, 10, 12, 10)
        stock_layout.setSpacing(8)
        stock_title = QLabel("Stock e retalhos elegíveis")
        stock_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.stock_hint = QLabel("Sem stock carregado.")
        self.stock_hint.setProperty("role", "muted")
        self.stock_empty_label = QLabel("Sem stock de chapa disponível para este material/espessura.")
        self.stock_empty_label.setAlignment(Qt.AlignCenter)
        self.stock_empty_label.setWordWrap(True)
        self.stock_empty_label.setStyleSheet(
            "padding: 12px; border: 1px dashed #bfd0e1; border-radius: 10px; "
            "background: #f8fbff; color: #486581; font-size: 12px; font-weight: 700;"
        )
        self.stock_table = QTableWidget(0, 6)
        self.stock_table.setHorizontalHeaderLabels(["Origem", "Lote/ID", "Dimensao", "Forma", "Disponivel", "Local"])
        self.stock_table.verticalHeader().setVisible(False)
        self.stock_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.stock_table.setSelectionBehavior(QTableWidget.SelectRows)
        stock_header = self.stock_table.horizontalHeader()
        stock_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        stock_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        stock_header.setSectionResizeMode(2, QHeaderView.Stretch)
        stock_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        stock_header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.stock_card.setMinimumHeight(118)
        self.stock_card.setMaximumHeight(220)
        self.stock_table.setMinimumHeight(72)
        stock_layout.addWidget(stock_title)
        stock_layout.addWidget(self.stock_hint)
        stock_layout.addWidget(self.stock_empty_label)
        stock_layout.addWidget(self.stock_table, 1)
        config_page_layout.addWidget(self.stock_card)
        config_page_layout.addStretch(1)

        materials_page = QWidget()
        materials_page_layout = QVBoxLayout(materials_page)
        materials_page_layout.setContentsMargins(0, 0, 0, 0)
        materials_page_layout.setSpacing(10)
        define_summary_card = CardFrame()
        define_summary_card.set_tone("default")
        define_summary_layout = QHBoxLayout(define_summary_card)
        define_summary_layout.setContentsMargins(12, 8, 12, 8)
        define_summary_layout.setSpacing(14)
        self.define_metric_group = QLabel("-")
        self.define_metric_forms = QLabel("-")
        self.define_metric_qty = QLabel("-")
        self.define_metric_prog = QLabel("-")
        self.define_metric_stock = QLabel("-")
        for title, label in (
            ("Grupo", self.define_metric_group),
            ("Formas", self.define_metric_forms),
            ("Peças pedidas", self.define_metric_qty),
            ("Programadas", self.define_metric_prog),
            ("Stock elegível", self.define_metric_stock),
        ):
            block = QVBoxLayout()
            block.setSpacing(1)
            title_label = QLabel(title)
            title_label.setStyleSheet("font-size: 9px; font-weight: 800; color: #486581; text-transform: uppercase;")
            label.setStyleSheet("font-size: 13px; font-weight: 900; color: #0f172a;")
            block.addWidget(title_label)
            block.addWidget(label)
            define_summary_layout.addLayout(block, 1)
        materials_page_layout.addWidget(define_summary_card)
        materials_page_layout.addWidget(parts_card, 1)

        recommendation_card = CardFrame()
        recommendation_card.set_tone("info")
        recommendation_layout = QVBoxLayout(recommendation_card)
        recommendation_layout.setContentsMargins(12, 10, 12, 10)
        recommendation_layout.setSpacing(6)
        recommendation_title = QLabel("Recomendação do motor")
        recommendation_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        recommendation_hint = QLabel("Resumo industrial do cenário mais interessante, com foco em compactação, compra adicional e aproveitamento real.")
        recommendation_hint.setProperty("role", "muted")
        recommendation_hint.setWordWrap(True)
        recommendation_metrics = QHBoxLayout()
        recommendation_metrics.setSpacing(10)
        self.recommendation_primary = QLabel("-")
        self.recommendation_secondary = QLabel("-")
        self.recommendation_delta = QLabel("-")
        for label in (self.recommendation_primary, self.recommendation_secondary, self.recommendation_delta):
            label.setStyleSheet(
                "padding: 8px 12px; border-radius: 10px; background: #ffffff; color: #0f172a; "
                "border: 1px solid #cbd5e1; font-size: 12px; font-weight: 800;"
            )
            recommendation_metrics.addWidget(label, 1)
        self.recommendation_note = QLabel("Corre a otimização para receber uma recomendação automática do estudo.")
        self.recommendation_note.setWordWrap(True)
        self.recommendation_note.setStyleSheet("font-size: 11px; color: #365b7c;")
        recommendation_layout.addWidget(recommendation_title)
        recommendation_layout.addWidget(recommendation_hint)
        recommendation_layout.addLayout(recommendation_metrics)
        recommendation_layout.addWidget(self.recommendation_note)

        summary_card = CardFrame()
        summary_layout = QFormLayout(summary_card)
        summary_layout.setContentsMargins(12, 10, 12, 10)
        summary_layout.setHorizontalSpacing(10)
        summary_layout.setVerticalSpacing(8)
        self.summary_labels: dict[str, QLabel] = {}
        for key, title in (
            ("profile", "Perfil escolhido"),
            ("method", "Método"),
            ("strategy", "Estratégia"),
            ("sheets", "Chapas no plano"),
            ("stock_used", "Stock aplicado"),
            ("buy_count", "Chapas a comprar"),
            ("parts", "Peças colocadas"),
            ("unplaced", "Peças fora"),
            ("util_net", "Utilização real %"),
            ("util_bbox", "Ocupação geométrica %"),
            ("compactness", "Compactação layout %"),
            ("purchase", "Compra adicional"),
            ("area", "Área total de chapa"),
            ("mass", "Massa estimada"),
            ("material_cost", "Custo estimado"),
        ):
            label = QLabel("-")
            label.setProperty("role", "field_value_strong" if key in ("util_net", "purchase", "material_cost") else "field_value")
            summary_layout.addRow(title, label)
            self.summary_labels[key] = label

        candidates_card = CardFrame()
        candidates_layout = QVBoxLayout(candidates_card)
        candidates_layout.setContentsMargins(12, 10, 12, 10)
        candidates_layout.setSpacing(8)
        candidates_title = QLabel("Comparação de cenários")
        candidates_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.candidate_table = QTableWidget(0, 7)
        self.candidate_table.setHorizontalHeaderLabels(["Cenário", "Método", "Chapas", "Compact. %", "Compra m2", "Total m2", "Fora"])
        self.candidate_table.verticalHeader().setVisible(False)
        self.candidate_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.candidate_table.setSelectionBehavior(QTableWidget.SelectRows)
        candidate_header = self.candidate_table.horizontalHeader()
        candidate_header.setSectionResizeMode(0, QHeaderView.Stretch)
        candidate_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        candidate_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        candidate_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        candidate_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        candidate_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        candidate_header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        candidates_layout.addWidget(candidates_title)
        candidates_layout.addWidget(self.candidate_table, 1)

        preview_card = CardFrame()
        preview_card.setStyleSheet(
            "CardFrame { background: #fbfdff; border: 1px solid #c8d7e6; border-radius: 16px; }"
        )
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(12, 10, 12, 14)
        preview_layout.setSpacing(8)
        preview_header = QHBoxLayout()
        preview_title = QLabel("Nesting")
        preview_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        fit_preview_btn = QPushButton("Ajustar vista")
        fit_preview_btn.setProperty("variant", "secondary")
        self.sheet_view_combo = QComboBox()
        self.sheet_view_combo.currentIndexChanged.connect(self._render_sheet_preview)
        preview_header.addWidget(preview_title)
        preview_header.addStretch(1)
        preview_header.addWidget(fit_preview_btn)
        preview_header.addWidget(self.sheet_view_combo)
        preview_layout.addLayout(preview_header)
        self.preview_hint = QLabel("Seleciona uma chapa analisada para ver o plano, a origem e o aproveitamento. Usa a roda do rato para zoom e duplo clique para enquadrar.")
        self.preview_hint.setProperty("role", "muted")
        self.preview_hint.setWordWrap(True)
        preview_layout.addWidget(self.preview_hint)
        self.preview_scene = QGraphicsScene(self)
        self.preview_view = NestPreviewView(self.preview_scene, self)
        self.preview_view.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing | QPainter.SmoothPixmapTransform)
        self.preview_view.setStyleSheet(
            "background: #ffffff; border: 1px solid #b9c8d8; border-radius: 10px;"
        )
        self.preview_view.setAlignment(Qt.AlignCenter)
        self.preview_view.setFrameShape(QFrame.NoFrame)
        self.preview_view.setLineWidth(0)
        self.preview_view.setMidLineWidth(0)
        self.preview_view.setContentsMargins(0, 0, 0, 0)
        self.preview_view.setMinimumHeight(268)
        fit_preview_btn.clicked.connect(self.preview_view.fit_scene)
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet(
            "QFrame { background: #f4f8fc; border: 2px solid #9cb2c8; border-radius: 14px; "
            "border-bottom: 3px solid #7f97ae; }"
        )
        preview_frame_layout = QVBoxLayout(self.preview_frame)
        preview_frame_layout.setContentsMargins(8, 8, 8, 14)
        preview_frame_layout.setSpacing(0)
        preview_frame_layout.addWidget(self.preview_view, 1)
        preview_layout.addWidget(self.preview_frame, 1)

        layouts_card = CardFrame()
        layouts_card.setStyleSheet(
            "CardFrame { background: #fbfdff; border: 1px solid #c8d7e6; border-radius: 16px; }"
        )
        layouts_layout = QVBoxLayout(layouts_card)
        layouts_layout.setContentsMargins(12, 10, 12, 10)
        layouts_layout.setSpacing(8)
        layouts_title = QLabel("Galeria de layouts")
        layouts_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        layouts_hint = QLabel("Vista rápida de todas as chapas calculadas. Seleciona um layout para atualizar o plano detalhado e a tabela.")
        layouts_hint.setProperty("role", "muted")
        layouts_hint.setWordWrap(True)
        self.layouts_hint_label = layouts_hint
        self.sheet_gallery = QListWidget()
        self.sheet_gallery.setViewMode(QListWidget.IconMode)
        self.sheet_gallery.setFlow(QListWidget.LeftToRight)
        self.sheet_gallery.setWrapping(True)
        self.sheet_gallery.setResizeMode(QListWidget.Adjust)
        self.sheet_gallery.setMovement(QListWidget.Static)
        self.sheet_gallery.setUniformItemSizes(False)
        self.sheet_gallery.setSpacing(10)
        self.sheet_gallery.setIconSize(QSize(250, 160))
        self.sheet_gallery.setGridSize(QSize(286, 244))
        self.sheet_gallery.setMaximumHeight(286)
        self.sheet_gallery.setStyleSheet(
            """
            QListWidget {
                background: #ffffff;
                border: 1px solid #d3dde8;
                border-radius: 14px;
                padding: 8px;
            }
            QListWidget::item {
                background: #fffdf7;
                border: 1px solid #dde5ee;
                border-radius: 12px;
                margin: 6px;
                padding: 8px;
            }
            QListWidget::item:selected {
                background: #fff7df;
                border: 1px solid #d3b25f;
            }
            """
        )
        self.sheet_gallery.currentRowChanged.connect(self._sync_preview_from_gallery)
        layouts_layout.addWidget(layouts_title)
        layouts_layout.addWidget(layouts_hint)
        layouts_layout.addWidget(self.sheet_gallery, 1)

        sheet_plan_card = CardFrame()
        sheet_plan_card.setStyleSheet(
            "CardFrame { background: #fbfdff; border: 1px solid #c8d7e6; border-radius: 16px; }"
        )
        sheet_plan_layout = QVBoxLayout(sheet_plan_card)
        sheet_plan_layout.setContentsMargins(12, 10, 12, 10)
        sheet_plan_layout.setSpacing(8)
        sheet_plan_title = QLabel("Mapa de chapas")
        sheet_plan_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        sheet_plan_hint = QLabel("Cada linha mostra a origem da chapa, o formato utilizado, peças colocadas e aproveitamento.")
        sheet_plan_hint.setProperty("role", "muted")
        sheet_plan_hint.setWordWrap(True)
        self.sheet_plan_hint_label = sheet_plan_hint
        self.sheet_plan_table = QTableWidget(0, 7)
        self.sheet_plan_table.setHorizontalHeaderLabels(["Chapa", "Origem", "Perfil", "Peças", "Util. real %", "Área m2", "Compra"])
        self.sheet_plan_table.verticalHeader().setVisible(False)
        self.sheet_plan_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.sheet_plan_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.sheet_plan_table.itemSelectionChanged.connect(self._sync_preview_from_sheet_table)
        self.sheet_plan_table.setStyleSheet(
            """
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d8e2ec;
                border-radius: 12px;
                gridline-color: #d7e3ee;
                font-size: 11.5px;
            }
            QHeaderView::section {
                background: #08104f;
                color: #ffffff;
                padding: 9px 8px;
                border: none;
                font-size: 11.5px;
                font-weight: 900;
            }
            """
        )
        sheet_plan_header = self.sheet_plan_table.horizontalHeader()
        sheet_plan_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        sheet_plan_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        sheet_plan_header.setSectionResizeMode(2, QHeaderView.Stretch)
        sheet_plan_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        sheet_plan_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        sheet_plan_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        sheet_plan_header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        sheet_plan_layout.addWidget(sheet_plan_title)
        sheet_plan_layout.addWidget(sheet_plan_hint)
        sheet_plan_layout.addWidget(self.sheet_plan_table, 1)

        warnings_card = CardFrame()
        warnings_layout = QVBoxLayout(warnings_card)
        warnings_layout.setContentsMargins(12, 10, 12, 10)
        warnings_layout.setSpacing(6)
        warnings_title = QLabel("Diagnóstico e observações")
        warnings_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        self.warning_edit = QTextEdit()
        self.warning_edit.setReadOnly(True)
        warnings_layout.addWidget(warnings_title)
        warnings_layout.addWidget(self.warning_edit, 1)

        unplaced_card = CardFrame()
        unplaced_layout = QVBoxLayout(unplaced_card)
        unplaced_layout.setContentsMargins(12, 10, 12, 10)
        unplaced_layout.setSpacing(8)
        unplaced_title = QLabel("Peças fora do plano")
        unplaced_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        unplaced_hint = QLabel("Peças não colocadas nesta análise, úteis para perceber falta de stock, formato curto ou geometria inválida.")
        unplaced_hint.setProperty("role", "muted")
        unplaced_hint.setWordWrap(True)
        self.unplaced_hint_label = unplaced_hint
        self.unplaced_table = QTableWidget(0, 4)
        self.unplaced_table.setHorizontalHeaderLabels(["Ref. externa", "Descricao", "Desenho", "Cópia"])
        self.unplaced_table.verticalHeader().setVisible(False)
        self.unplaced_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.unplaced_table.setSelectionBehavior(QTableWidget.SelectRows)
        unplaced_header = self.unplaced_table.horizontalHeader()
        unplaced_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        unplaced_header.setSectionResizeMode(1, QHeaderView.Stretch)
        unplaced_header.setSectionResizeMode(2, QHeaderView.Stretch)
        unplaced_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        unplaced_layout.addWidget(unplaced_title)
        unplaced_layout.addWidget(unplaced_hint)
        unplaced_layout.addWidget(self.unplaced_table, 1)

        summary_tab = QWidget()
        summary_tab_layout = QVBoxLayout(summary_tab)
        summary_tab_layout.setContentsMargins(0, 0, 0, 0)
        summary_tab_layout.setSpacing(10)
        summary_tab_layout.addWidget(recommendation_card)
        results_top_split = QSplitter(Qt.Horizontal)
        results_top_split.setChildrenCollapsible(False)
        results_top_split.addWidget(summary_card)
        results_top_split.addWidget(candidates_card)
        results_top_split.setSizes([380, 780])
        summary_tab_layout.addWidget(results_top_split, 1)

        sheets_tab = QWidget()
        sheets_tab_layout = QVBoxLayout(sheets_tab)
        sheets_tab_layout.setContentsMargins(0, 0, 0, 0)
        sheets_tab_layout.setSpacing(10)
        sheets_split = QSplitter(Qt.Horizontal)
        sheets_split.setChildrenCollapsible(False)
        sheets_split.addWidget(sheet_plan_card)
        sheets_split.addWidget(preview_card)
        sheets_split.setSizes([540, 760])
        sheets_tab_layout.addWidget(layouts_card)
        sheets_tab_layout.addWidget(sheets_split, 1)

        diagnostics_tab = QWidget()
        diagnostics_tab_layout = QVBoxLayout(diagnostics_tab)
        diagnostics_tab_layout.setContentsMargins(0, 0, 0, 0)
        diagnostics_tab_layout.setSpacing(10)
        diagnostics_split = QSplitter(Qt.Horizontal)
        diagnostics_split.setChildrenCollapsible(False)
        diagnostics_split.addWidget(unplaced_card)
        diagnostics_split.addWidget(warnings_card)
        diagnostics_split.setSizes([640, 640])
        diagnostics_tab_layout.addWidget(diagnostics_split, 1)

        cost_card = CardFrame()
        cost_layout = QVBoxLayout(cost_card)
        cost_layout.setContentsMargins(12, 10, 12, 10)
        cost_layout.setSpacing(8)
        cost_title = QLabel("Quadro de custos")
        cost_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        cost_hint = QLabel("Combina o custo real de matéria pelo plano de chapa com uma estimativa de processo baseada no motor de orçamentação laser.")
        cost_hint.setProperty("role", "muted")
        cost_hint.setWordWrap(True)
        self.cost_hint_label = cost_hint
        self.cost_table = QTableWidget(0, 4)
        self.cost_table.setHorizontalHeaderLabels(["Indicador", "Valor", "Base", "Notas"])
        self.cost_table.verticalHeader().setVisible(False)
        self.cost_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.cost_table.setSelectionBehavior(QTableWidget.SelectRows)
        cost_header = self.cost_table.horizontalHeader()
        cost_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        cost_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        cost_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        cost_header.setSectionResizeMode(3, QHeaderView.Stretch)
        cost_layout.addWidget(cost_title)
        cost_layout.addWidget(cost_hint)
        cost_layout.addWidget(self.cost_table, 1)

        process_card = CardFrame()
        process_layout = QVBoxLayout(process_card)
        process_layout.setContentsMargins(12, 10, 12, 10)
        process_layout.setSpacing(8)
        process_title = QLabel("Resumo por peça")
        process_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        process_hint = QLabel("Mostra as peças colocadas, quantidade, tempo máquina, comprimento de corte, pierces e valor comercial estimado.")
        process_hint.setProperty("role", "muted")
        process_hint.setWordWrap(True)
        self.process_hint_label = process_hint
        self.process_table = QTableWidget(0, 10)
        self.process_table.setHorizontalHeaderLabels(["Ref.", "Descricao", "Qtd", "Tempo min", "Corte m", "Pierces", "MP rateada", "Compra rateada", "Comercial atual", "Ajustado plano"])
        self.process_table.verticalHeader().setVisible(False)
        self.process_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.process_table.setSelectionBehavior(QTableWidget.SelectRows)
        process_header = self.process_table.horizontalHeader()
        process_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(1, QHeaderView.Stretch)
        process_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        process_header.setSectionResizeMode(9, QHeaderView.ResizeToContents)
        process_layout.addWidget(process_title)
        process_layout.addWidget(process_hint)
        process_layout.addWidget(self.process_table, 1)

        self.decision_edit = QTextEdit()
        self.decision_edit.setReadOnly(True)
        decision_card = CardFrame()
        decision_layout = QVBoxLayout(decision_card)
        decision_layout.setContentsMargins(12, 10, 12, 10)
        decision_layout.setSpacing(8)
        decision_title = QLabel("Log de decisão")
        decision_title.setStyleSheet("font-size: 14px; font-weight: 800; color: #0f172a;")
        decision_hint = QLabel("Explica porque o cenário foi escolhido e quais os trade-offs face às alternativas.")
        decision_hint.setProperty("role", "muted")
        decision_hint.setWordWrap(True)
        self.decision_hint_label = decision_hint
        decision_layout.addWidget(decision_title)
        decision_layout.addWidget(decision_hint)
        decision_layout.addWidget(self.decision_edit, 1)

        costs_tab = QWidget()
        costs_tab_layout = QVBoxLayout(costs_tab)
        costs_tab_layout.setContentsMargins(0, 0, 0, 0)
        costs_tab_layout.setSpacing(10)
        costs_top_split = QSplitter(Qt.Horizontal)
        costs_top_split.setChildrenCollapsible(False)
        costs_top_split.addWidget(cost_card)
        costs_top_split.addWidget(decision_card)
        costs_top_split.setSizes([780, 500])
        costs_tab_layout.addWidget(costs_top_split)
        costs_tab_layout.addWidget(process_card, 1)

        define_page = QWidget()
        define_page_layout = QVBoxLayout(define_page)
        define_page_layout.setContentsMargins(0, 0, 0, 0)
        define_page_layout.setSpacing(10)
        self.define_tabs = QTabWidget()
        self.define_tabs.setDocumentMode(True)
        self.define_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.define_tabs.setStyleSheet(
            """
            QTabBar::tab {
                background: #edf3f9;
                border: 1px solid #bfd0e1;
                padding: 6px 12px;
                border-top-left-radius: 7px;
                border-top-right-radius: 7px;
                color: #23415e;
                font-weight: 700;
            }
            QTabBar::tab:hover {
                background: #f7fbff;
            }
            QTabBar::tab:selected {
                background: #1d4ed8;
                color: #ffffff;
                border-color: #1d4ed8;
            }
            """
        )
        self.define_tabs.addTab(config_page, "Regras e formatos")
        self.define_tabs.addTab(materials_page, "Peças e stock")
        self.define_tabs.tabBar().hide()
        define_page_layout.addWidget(self.define_tabs, 1)

        nest_page = QWidget()
        nest_page_layout = QVBoxLayout(nest_page)
        nest_page_layout.setContentsMargins(0, 0, 0, 0)
        nest_page_layout.setSpacing(10)
        self.nest_tabs = QTabWidget()
        self.nest_tabs.setDocumentMode(True)
        self.nest_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.nest_tabs.setStyleSheet(self.define_tabs.styleSheet())
        self.nest_tabs.addTab(summary_tab, "Cenários")
        self.nest_tabs.addTab(sheets_tab, "Layouts / chapas")
        self.nest_tabs.tabBar().hide()
        nest_page_layout.addWidget(self.nest_tabs, 1)

        results_page = QWidget()
        results_page_layout = QVBoxLayout(results_page)
        results_page_layout.setContentsMargins(0, 0, 0, 0)
        results_page_layout.setSpacing(10)
        self.results_tabs = QTabWidget()
        self.results_tabs.setDocumentMode(True)
        self.results_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.results_tabs.setStyleSheet(self.define_tabs.styleSheet())
        self.results_tabs.addTab(diagnostics_tab, "Pendentes / avisos")
        self.results_tabs.addTab(costs_tab, "Custos / decisão")
        self.results_tabs.tabBar().hide()
        results_page_layout.addWidget(self.results_tabs, 1)

        self.section_stack.addTab(define_page, "1. Definir")
        self.section_stack.addTab(nest_page, "2. Nest")
        self.section_stack.addTab(results_page, "3. Corte / orçamento")
        self.body_layout.addWidget(self.section_stack, 1)
        self.wizard_path_label = QLabel("")
        self.wizard_path_label.setProperty("role", "muted")
        self.wizard_path_label.setStyleSheet("font-size: 11px; font-weight: 700; color: #365b7c;")
        self.body_layout.addWidget(self.wizard_path_label)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.prev_section_btn = QPushButton("Anterior")
        self.prev_section_btn.setProperty("variant", "secondary")
        self.prev_section_btn.clicked.connect(self._go_previous_section)
        self.next_section_btn = QPushButton("Avançar")
        self.next_section_btn.setProperty("variant", "secondary")
        self.next_section_btn.clicked.connect(self._go_next_section)
        self.analyze_btn = QPushButton("Iniciar otimização")
        self.analyze_btn.clicked.connect(self._analyze)
        actions.addWidget(self.prev_section_btn)
        actions.addWidget(self.next_section_btn)
        actions.addWidget(self.analyze_btn)
        self.nesting_progress = QProgressBar()
        self.nesting_progress.setRange(0, 100)
        self.nesting_progress.setValue(0)
        self.nesting_progress.setTextVisible(True)
        self.nesting_progress.setFormat("A otimizar nesting... %p%")
        self.nesting_progress.setMinimumWidth(220)
        self.nesting_progress.setStyleSheet(
            "QProgressBar {"
            "border:1px solid #d58a22;"
            "border-radius:9px;"
            "background:#fff6e8;"
            "color:#3a2205;"
            "font-weight:800;"
            "text-align:center;"
            "height:22px;"
            "}"
            "QProgressBar::chunk {"
            "border-radius:8px;"
            "background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #ffb347, stop:1 #f47c20);"
            "}"
        )
        self.nesting_progress.hide()
        actions.addWidget(self.nesting_progress)
        actions.addStretch(1)
        self.pdf_btn = QPushButton("PDF do estudo")
        self.pdf_btn.setProperty("variant", "secondary")
        self.pdf_btn.clicked.connect(self._open_study_pdf)
        actions.addWidget(self.pdf_btn)
        root.addLayout(actions)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Close)
        buttons.accepted.connect(self._accept_dialog)
        buttons.rejected.connect(self.reject)
        buttons.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
        buttons.button(QDialogButtonBox.Close).setText("Fechar")
        root.addWidget(buttons)
        self.dialog_buttons = buttons
        self._nesting_thread: QThread | None = None
        self._nesting_worker: NestingWorker | None = None
        self._nesting_progress_value = 0
        self._nesting_progress_timer = QTimer(self)
        self._nesting_progress_timer.setInterval(420)
        self._nesting_progress_timer.timeout.connect(self._tick_nesting_progress)

        self.group_combo.currentIndexChanged.connect(self._load_group_rows)
        self.auto_sheet_check.toggled.connect(self._sync_sheet_mode)
        self.shape_check.toggled.connect(self._sync_shape_mode)
        self.common_line_check.toggled.connect(self._sync_process_mode)
        self.lead_opt_check.toggled.connect(self._sync_process_mode)
        self.section_stack.currentChanged.connect(self._handle_section_changed)
        self.define_tabs.currentChanged.connect(self._sync_wizard_controls)
        self.nest_tabs.currentChanged.connect(self._handle_nest_tab_changed)
        self.results_tabs.currentChanged.connect(self._sync_wizard_controls)
        self._refresh_sheet_profiles()
        self._sync_sheet_mode()
        self._sync_shape_mode()
        self._sync_process_mode()
        self.section_stack.setCurrentIndex(0)
        self._set_wizard_step(0)
        self._load_saved_studies()
        self._load_group_rows()
        self._repair_dialog_copy()
        self._apply_premium_layout()
        self._sync_window_toggle_label()

    def _resolve_target_screen(self):
        window_handle = self.windowHandle()
        if window_handle is not None and window_handle.screen() is not None:
            return window_handle.screen()
        parent = self.parentWidget()
        if parent is not None:
            parent_handle = parent.windowHandle()
            if parent_handle is not None and parent_handle.screen() is not None:
                return parent_handle.screen()
            parent_screen = parent.screen()
            if parent_screen is not None:
                return parent_screen
            parent_frame = parent.frameGeometry()
            if parent_frame.isValid():
                located = QGuiApplication.screenAt(parent_frame.center())
                if located is not None:
                    return located
        screen = self.screen()
        if screen is not None:
            return screen
        return QGuiApplication.primaryScreen()

    def _frame_extra_size(self) -> QSize:
        window_handle = self.windowHandle()
        if window_handle is not None:
            try:
                margins = window_handle.frameMargins()
                extra_width = max(0, int(margins.left()) + int(margins.right()))
                extra_height = max(0, int(margins.top()) + int(margins.bottom()))
                if extra_width > 0 or extra_height > 0:
                    return QSize(extra_width, extra_height)
            except Exception:
                pass
        frame = self.frameGeometry()
        body = self.geometry()
        return QSize(max(16, frame.width() - body.width()), max(48, frame.height() - body.height()))

    def _fit_to_available_screen(self) -> None:
        screen = self._resolve_target_screen()
        if screen is None:
            return
        if not self._stored_window_geometry.isValid():
            self._stored_window_geometry = self.geometry()
        safe = screen.availableGeometry().adjusted(8, 8, -8, -8)
        if safe.width() <= 0 or safe.height() <= 0:
            safe = screen.availableGeometry()
        frame_extra = self._frame_extra_size()
        max_width = max(980, safe.width() - frame_extra.width())
        max_height = max(620, safe.height() - frame_extra.height())
        target_width = min(max(1180, int(max_width * 0.985)), max_width)
        target_height = min(max(720, int(max_height * 0.965)), max_height)
        pos_x = safe.x() + max(0, int((safe.width() - target_width) / 2))
        pos_y = safe.y() + max(0, int((safe.height() - target_height) / 2))
        self.resize(target_width, target_height)
        self.move(pos_x, pos_y)
        self._window_fitted_to_screen = True
        layout = self.layout()
        if layout is not None:
            layout.activate()
        self._sync_window_toggle_label()

    def _toggle_window_mode(self) -> None:
        if self._window_fitted_to_screen:
            self.showNormal()
            self.setMaximumSize(16777215, 16777215)
            if self._stored_window_geometry.isValid():
                self.setGeometry(self._stored_window_geometry)
            self._window_fitted_to_screen = False
        else:
            self._stored_window_geometry = self.geometry()
            self.showNormal()
            self._fit_to_available_screen()
            self._window_fitted_to_screen = True
        QTimer.singleShot(0, self._sync_window_toggle_label)

    def _sync_window_toggle_label(self) -> None:
        if self._window_fitted_to_screen:
            self.window_toggle_btn.setText("Restaurar")
            self.window_toggle_btn.setToolTip("Volta ao tamanho anterior da janela.")
        else:
            self.window_toggle_btn.setText("Maximizar")
            self.window_toggle_btn.setToolTip("Ajusta a janela a area util do ecra, sem sobrepor a barra do Windows.")

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        if not self._window_fitted_to_screen:
            QTimer.singleShot(0, self._fit_to_available_screen)
            QTimer.singleShot(80, self._fit_to_available_screen)
        self._sync_window_toggle_label()

    def changeEvent(self, event) -> None:  # type: ignore[override]
        super().changeEvent(event)
        QTimer.singleShot(0, self._sync_window_toggle_label)

    def _repair_dialog_copy(self) -> None:
        self.setWindowTitle("Nesting Laser")
        self.group_badge_label.setText("Grupo não definido")
        self.header_title_label.setText("Nesting Laser")
        self.header_subtitle_label.setText("Fluxo orçamental com stock, retalhos, formatos standard e análise automática de cenários.")
        self.rotate_check.setText("Permitir rotação automática")
        self.mirror_check.setText("Permitir espelhar peças")
        self.free_angle_check.setText("Rotações livres")
        self.auto_sheet_check.setText("Escolher melhor formato automaticamente")
        self.use_stock_check.setText("Usar stock e retalhos primeiro")
        self.allow_purchase_check.setText("Permitir compra complementar")
        self.shape_check.setText("Nesting por contorno")
        self.common_line_check.setText("Estimar common-line")
        self.lead_opt_check.setText("Otimizar lead-ins / lead-outs")
        self.prev_section_btn.setText("Anterior")
        self.next_section_btn.setText("Seguinte")
        self.analyze_btn.setText("Iniciar otimização")
        self.pdf_btn.setText("PDF do estudo")
        self.group_badge_label.setText("Grupo não definido")
        self.header_subtitle_label.setText("")
        self.rotate_check.setText("Permitir rotação automática")
        self.analyze_btn.setText("Iniciar otimização")
        labels = self.findChildren(QLabel)
        for widget in labels:
            widget.setText(_repair_mojibake_text(widget.text()))
        for widget in self.findChildren(QPushButton):
            widget.setText(_repair_mojibake_text(widget.text()))
            if "Biblioteca de chapas" in widget.text():
                widget.setText("Biblioteca de chapas")
            elif "Parametros DXF" in widget.text() or "Parâmetros DXF" in widget.text() or "ParÃ¢metros DXF" in widget.text():
                widget.setText("Parâmetros DXF e custos")
        for widget in self.findChildren(QCheckBox):
            widget.setText(_repair_mojibake_text(widget.text()))
        self.sheet_library_btn.setText("Biblioteca de chapas")
        self.dxf_settings_btn.setText("Parâmetros DXF e custos")
        self.define_tabs.setTabText(0, "Parâmetros e chapa")
        self.define_tabs.setTabText(1, "Peças e stock")
        self.nest_tabs.setTabText(0, "Cenários")
        self.nest_tabs.setTabText(1, "Layouts e chapas")
        self.results_tabs.setTabText(0, "Pendentes e avisos")
        self.results_tabs.setTabText(1, "Custos e decisão")
        self.section_stack.setTabText(0, "1. Definir")
        self.section_stack.setTabText(1, "2. Nest")
        self.dxf_settings_btn.setText("Parâmetros DXF e custos")
        self.define_tabs.setTabText(0, "Parâmetros e chapa")
        self.define_tabs.setTabText(1, "Peças e stock")
        self.nest_tabs.setTabText(0, "Cenários")
        self.results_tabs.setTabText(1, "Custos e decisão")
        self.section_stack.setTabText(2, "3. Corte / orçamento")
        for tabs in (self.define_tabs, self.nest_tabs, self.results_tabs, self.section_stack):
            for index in range(tabs.count()):
                tabs.setTabText(index, _repair_mojibake_text(tabs.tabText(index)))
        ok_btn = self.findChild(QDialogButtonBox)
        if ok_btn is not None:
            if ok_btn.button(QDialogButtonBox.Ok):
                ok_btn.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
            if ok_btn.button(QDialogButtonBox.Ok):
                ok_btn.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
            if ok_btn.button(QDialogButtonBox.Ok):
                ok_btn.button(QDialogButtonBox.Ok).setText("Aplicar ao orçamento")
            if ok_btn.button(QDialogButtonBox.Close):
                ok_btn.button(QDialogButtonBox.Close).setText("Fechar")

    def _apply_premium_layout(self) -> None:
        root = self.layout()
        if isinstance(root, QVBoxLayout):
            root.setContentsMargins(18, 18, 18, 18)
            root.setSpacing(14)
        self.setStyleSheet(
            """
            QDialog {
                background: #edf3f8;
            }
            QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox, QTextEdit, QPlainTextEdit {
                min-height: 34px;
                padding: 4px 10px;
                background: #ffffff;
                border: 1px solid #c8d4e2;
                border-radius: 10px;
                selection-background-color: #16344f;
            }
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus, QTextEdit:focus {
                border: 1px solid #5f87a8;
            }
            QCheckBox {
                spacing: 8px;
                font-size: 12px;
                color: #16324d;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d3dde8;
                border-radius: 12px;
                gridline-color: #dde6ef;
            }
            QListWidget {
                background: #ffffff;
                border: 1px solid #d3dde8;
                border-radius: 14px;
                padding: 6px;
            }
            """
        )
        self.header_card.setStyleSheet(
            "QFrame {background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #ffffff, stop:1 #f6f9fc); border: 1px solid #d4dde7; border-radius: 20px;}"
        )
        self.stepper_card.setStyleSheet(
            "QFrame {background: #ffffff; border: 1px solid #d4dde7; border-radius: 20px;}"
        )
        self.header_title_label.setStyleSheet("font-size: 30px; font-weight: 900; color: #102a43;")
        self.header_subtitle_label.setText("")
        self.header_subtitle_label.hide()
        self.header_subtitle_label.setMinimumHeight(0)
        self.header_subtitle_label.setMaximumHeight(0)
        self.group_badge_label.setStyleSheet(
            "padding: 10px 18px; border-radius: 999px; background: #edf4fb; color: #274c77; "
            "border: 1px solid #c7d7e8; font-size: 12px; font-weight: 900;"
        )
        self.section_stack.setStyleSheet(
            """
            QTabWidget::pane {
                border: 1px solid #d5dee9;
                background: #f7fafc;
                border-radius: 18px;
                top: -1px;
            }
            QTabBar::tab {
                background: #f0f4f8;
                border: 1px solid #d6e0ea;
                padding: 10px 20px;
                min-width: 136px;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                color: #16324d;
                font-weight: 800;
            }
            QTabBar::tab:selected {
                background: #16344f;
                color: #ffffff;
                border-color: #16344f;
            }
            """
        )
        self.page_title_label.setStyleSheet("font-size: 24px; font-weight: 900; color: #102a43;")
        self.page_subtitle_label.setText("")
        self.page_subtitle_label.hide()
        self.page_subtitle_label.setMinimumHeight(0)
        self.page_subtitle_label.setMaximumHeight(0)
        for button in self.major_step_buttons:
            button.setMinimumHeight(44)
            button.setMinimumWidth(168)
            button.setCursor(Qt.PointingHandCursor)
        for button in (self.sheet_library_btn, self.dxf_settings_btn, self.prev_section_btn, self.next_section_btn, self.pdf_btn):
            button.setMinimumHeight(38)
            button.setCursor(Qt.PointingHandCursor)
        self.analyze_btn.setMinimumHeight(40)
        self.analyze_btn.setMinimumWidth(168)
        self.analyze_btn.setCursor(Qt.PointingHandCursor)
        self.group_combo.setMinimumHeight(36)
        self.sheet_combo.setMinimumHeight(36)
        self.sheet_view_combo.setMinimumHeight(34)
        self.preview_view.setMinimumHeight(360)
        self.sheet_gallery.setMinimumHeight(244)
        self.sheet_gallery.setMaximumHeight(286)
        self.stock_card.setMaximumHeight(186)
        self.stock_empty_label.setMinimumHeight(46)
        self.sheet_library_btn.setMinimumWidth(188)
        self.dxf_settings_btn.setMinimumWidth(206)
        for table in (
            self.stock_table,
            self.candidate_table,
            self.sheet_plan_table,
            self.unplaced_table,
            self.cost_table,
            self.process_table,
        ):
            table.verticalHeader().setDefaultSectionSize(32)
            table.setAlternatingRowColors(True)
        self.parts_table.verticalHeader().setDefaultSectionSize(96)
        self.parts_table.setAlternatingRowColors(True)
        self.parts_table.setMinimumHeight(190)
        self.preview_hint.setVisible(False)
        self.preview_hint.setText("")
        self.preview_hint.setMinimumHeight(0)
        self.preview_hint.setMaximumHeight(0)
        for hint_label in (
            getattr(self, "layouts_hint_label", None),
            getattr(self, "sheet_plan_hint_label", None),
            getattr(self, "unplaced_hint_label", None),
            getattr(self, "cost_hint_label", None),
            getattr(self, "process_hint_label", None),
            getattr(self, "decision_hint_label", None),
        ):
            if hint_label is not None:
                hint_label.setText("")
                hint_label.setVisible(False)
                hint_label.setMinimumHeight(0)
                hint_label.setMaximumHeight(0)
        self.stock_hint.setVisible(False)
        self.stock_hint.setText("")
        self.stock_hint.setMinimumHeight(0)
        self.stock_hint.setMaximumHeight(0)
        self.parts_status_label.setStyleSheet("font-size: 11px; font-weight: 700; color: #486581;")
        self.wizard_path_label.setText("")
        self.wizard_path_label.hide()
        self.wizard_path_label.setMinimumHeight(0)
        self.wizard_path_label.setMaximumHeight(0)

    def quote_bridge_payload(self) -> dict[str, Any]:
        if not self.result_data:
            return {}
        summary = dict(self.result_data.get("summary", {}) or {})
        report = self._build_cost_report(summary)
        totals = dict(report.get("totals", {}) or {})
        rateio = dict(report.get("rateio", {}) or {})
        return {
            "analysis_method": self._analysis_method_label(summary),
            "selected_profile_name": str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or "Apenas stock").strip() or "Apenas stock",
            "part_count_requested": int(summary.get("part_count_requested", 0) or 0),
            "part_count_placed": int(summary.get("part_count_placed", 0) or 0),
            "part_count_unplaced": int(summary.get("part_count_unplaced", 0) or 0),
            "sheet_count": int(summary.get("sheet_count", 0) or 0),
            "material_net_cost_eur": float(summary.get("material_net_cost_eur", 0) or 0),
            "material_purchase_requirement_eur": float(summary.get("material_purchase_requirement_eur", 0) or 0),
            "quoted_total_eur": float(totals.get("quoted_total_eur", 0) or 0),
            "rateio_adjusted_quote_total_eur": float(rateio.get("adjusted_quote_total_eur", 0) or 0),
            "rateio_current_quote_total_eur": float(rateio.get("current_quote_total_eur", 0) or 0),
            "rateio_delta_eur": float(rateio.get("delta_eur", 0) or 0),
            "rateio_base": str(rateio.get("base", "") or "").strip(),
            "part_rows": [dict(row or {}) for row in list(report.get("part_rows", []) or [])],
        }

    def _set_study_status(self, text: str, tone: str = "muted") -> None:
        palette = {
            "muted": ("#eef4ff", "#365b7c", "#bfd2ea"),
            "success": ("#edf4fb", "#1f3b57", "#bfd0e1"),
            "warning": ("#fff8eb", "#b54708", "#f1d39a"),
            "danger": ("#fff1f2", "#b42318", "#f0c1bc"),
        }
        bg, fg, border = palette.get(str(tone or "muted"), palette["muted"])
        self.study_status_label.setText(str(text or "").strip())
        self.study_status_label.setStyleSheet(
            "padding: 6px 10px; border-radius: 10px; font-size: 11px; font-weight: 700;"
            f" background: {bg}; color: {fg}; border: 1px solid {border};"
        )
        can_export = bool(self.quote_number and (self.result_data or self.saved_studies.get(self._current_group_key(), {})))
        if hasattr(self, "pdf_btn"):
            self.pdf_btn.setEnabled(can_export)

    def _current_group_key(self) -> str:
        group = self._current_group()
        return str(group.get("key", "") or self.group_combo.currentText() or "").strip()

    def _group_rows_signature(self, rows: list[dict[str, Any]] | None = None) -> str:
        signature_rows: list[dict[str, Any]] = []
        for row in list(rows or self.current_group_rows or []):
            payload = dict(row or {})
            signature_rows.append(
                {
                    "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
                    "descricao": str(payload.get("descricao", "") or "").strip(),
                    "desenho": str(payload.get("desenho", "") or "").strip(),
                    "material": str(payload.get("material", "") or "").strip(),
                    "espessura": str(payload.get("espessura", "") or "").strip(),
                    "operacao": str(payload.get("operacao", "") or "").strip(),
                    "qtd": round(float(payload.get("qtd", 0) or 0), 4),
                }
            )
        try:
            return __import__("json").dumps(signature_rows, ensure_ascii=False, sort_keys=True, default=str)
        except Exception:
            return repr(signature_rows)

    def _current_options_snapshot(self) -> dict[str, Any]:
        return {
            "sheet_profile_name": self.sheet_combo.currentText().strip(),
            "part_spacing_mm": float(self.spacing_spin.value()),
            "edge_margin_mm": float(self.edge_spin.value()),
            "allow_rotate": bool(self.rotate_check.isChecked()),
            "allow_mirror": bool(self.mirror_check.isChecked()),
            "free_angle_rotation": bool(self.free_angle_check.isChecked()),
            "auto_select_sheet": bool(self.auto_sheet_check.isChecked()),
            "use_stock_first": bool(self.use_stock_check.isChecked()),
            "allow_purchase_fallback": bool(self.allow_purchase_check.isChecked()),
            "shape_aware": bool(self.shape_check.isChecked()),
            "shape_strict": bool(self.shape_strict_check.isChecked()),
            "shape_grid_mm": float(self.shape_grid_spin.value()),
            "common_line_estimate": bool(self.common_line_check.isChecked()),
            "common_line_tolerance_mm": float(self.common_line_tol_spin.value()),
            "lead_optimization": bool(self.lead_opt_check.isChecked()),
            "lead_optimization_pct": float(self.lead_opt_pct_spin.value()),
        }

    def _current_overrides_snapshot(self) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        for row in list(self.current_group_rows or []):
            row_key = self._part_override_key(row)
            payload[row_key] = {
                "rotation_policy": self._part_rotation_override(row),
                "priority": self._part_priority_override(row),
                "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                "descricao": str(row.get("descricao", "") or "").strip(),
            }
        return payload

    def _load_saved_studies(self) -> None:
        self.saved_studies = {}
        if self.quote_number and hasattr(self.backend, "orc_nesting_studies"):
            try:
                raw = self.backend.orc_nesting_studies(self.quote_number) or {}
                self.saved_studies = {str(key): dict(value or {}) for key, value in dict(raw).items() if str(key).strip()}
            except Exception:
                self.saved_studies = {}
        if not self.quote_number:
            self._set_study_status("Guarda primeiro o orçamento para manter o estudo de nesting em base de dados e gerar PDF oficial.", "warning")
        elif self.saved_studies:
            self._set_study_status("Estudos anteriores encontrados. Ao trocar de grupo, o plano guardado será reaberto automaticamente.", "success")
        else:
            self._set_study_status("Ainda não existe estudo guardado para este orçamento. O primeiro cálculo ficará gravado automaticamente.", "muted")

    def _apply_study_controls(self, study: dict[str, Any]) -> None:
        options = dict(study.get("options", {}) or {})
        part_overrides = dict(study.get("part_overrides", {}) or {})
        self.part_rotation_overrides = {
            str(key): str(dict(value or {}).get("rotation_policy", "auto") or "auto")
            for key, value in part_overrides.items()
            if str(key).strip()
        }
        self.part_priority_overrides = {
            str(key): int(dict(value or {}).get("priority", 0) or 0)
            for key, value in part_overrides.items()
            if str(key).strip()
        }
        if options:
            self.spacing_spin.setValue(float(options.get("part_spacing_mm", self.spacing_spin.value()) or self.spacing_spin.value()))
            self.edge_spin.setValue(float(options.get("edge_margin_mm", self.edge_spin.value()) or self.edge_spin.value()))
            self.rotate_check.setChecked(bool(options.get("allow_rotate", self.rotate_check.isChecked())))
            self.auto_sheet_check.setChecked(bool(options.get("auto_select_sheet", self.auto_sheet_check.isChecked())))
            self.use_stock_check.setChecked(bool(options.get("use_stock_first", self.use_stock_check.isChecked())))
            self.allow_purchase_check.setChecked(bool(options.get("allow_purchase_fallback", self.allow_purchase_check.isChecked())))
            self.shape_check.setChecked(bool(options.get("shape_aware", self.shape_check.isChecked())))
            self.shape_strict_check.setChecked(bool(options.get("shape_strict", self.shape_strict_check.isChecked())))
            self.shape_grid_spin.setValue(float(options.get("shape_grid_mm", self.shape_grid_spin.value()) or self.shape_grid_spin.value()))
            self.common_line_check.setChecked(bool(options.get("common_line_estimate", self.common_line_check.isChecked())))
            self.common_line_tol_spin.setValue(float(options.get("common_line_tolerance_mm", self.common_line_tol_spin.value()) or self.common_line_tol_spin.value()))
            self.lead_opt_check.setChecked(bool(options.get("lead_optimization", self.lead_opt_check.isChecked())))
            self.lead_opt_pct_spin.setValue(float(options.get("lead_optimization_pct", self.lead_opt_pct_spin.value()) or self.lead_opt_pct_spin.value()))
            selected_sheet_name = str(options.get("sheet_profile_name", "") or "").strip()
            if selected_sheet_name:
                self._refresh_sheet_profiles(selected_sheet_name)
        self._sync_sheet_mode()
        self._sync_shape_mode()
        self._sync_process_mode()

    def _study_payload(self) -> dict[str, Any]:
        summary = dict(self.result_data.get("summary", {}) or {})
        return {
            "quote_number": self.quote_number,
            "group_key": self._current_group_key(),
            "group_label": str(self._current_group().get("label", "") or self.group_combo.currentText() or "").strip(),
            "row_signature": self._group_rows_signature(),
            "options": self._current_options_snapshot(),
            "part_overrides": self._current_overrides_snapshot(),
            "summary": summary,
            "result_data": dict(self.result_data or {}),
            "quote_bridge": self.quote_bridge_payload(),
            "cost_report": self._build_cost_report(summary),
        }

    def _save_current_study(self, *, silent: bool = False) -> bool:
        if not self.quote_number or not self.result_data or not hasattr(self.backend, "orc_save_nesting_study"):
            return False
        try:
            stored = self.backend.orc_save_nesting_study(self.quote_number, self._study_payload()) or {}
        except Exception as exc:
            if not silent:
                self._set_study_status(str(exc), "danger")
            return False
        group_key = str(stored.get("group_key", "") or self._current_group_key()).strip()
        if group_key:
            self.saved_studies[group_key] = dict(stored or {})
        saved_at = str(stored.get("updated_at", "") or "").strip()
        message = "Estudo guardado automaticamente na base de dados."
        if saved_at:
            message += f" Última atualização: {saved_at[:16].replace('T', ' ')}."
        self._set_study_status(message, "success")
        return True

    def _restore_saved_group_result(self, study: dict[str, Any]) -> None:
        if not study:
            if self.quote_number:
                self._set_study_status("Este grupo ainda não tem estudo guardado. Corre a otimização para gravar o primeiro cenário.", "muted")
            return
        saved_signature = str(study.get("row_signature", "") or "").strip()
        current_signature = self._group_rows_signature()
        if saved_signature and saved_signature != current_signature:
            self._set_study_status("Existe um estudo guardado para este grupo, mas as linhas do orçamento mudaram. Recalcula para atualizar o plano.", "warning")
            return
        restored = dict(study.get("result_data", {}) or {})
        if not restored:
            self._set_study_status("Foram recuperadas regras guardadas para este grupo, mas sem resultado final associado.", "warning")
            return
        self._apply_result_data(restored)
        saved_at = str(study.get("updated_at", "") or "").strip()
        message = "Estudo anterior reaberto automaticamente."
        if saved_at:
            message += f" Guardado em {saved_at[:16].replace('T', ' ')}."
        self._set_study_status(message, "success")

    def _open_study_pdf(self) -> None:
        if not self.quote_number:
            QMessageBox.information(self, "Nesting", "Guarda primeiro o orçamento para gerar o PDF oficial do estudo.")
            return
        if self.result_data:
            self._save_current_study(silent=True)
        if not hasattr(self.backend, "orc_open_nesting_study_pdf"):
            QMessageBox.warning(self, "Nesting", "O backend atual não suporta exportação PDF do estudo.")
            return
        try:
            self.backend.orc_open_nesting_study_pdf(self.quote_number, self._current_group_key())
            self._set_study_status("PDF do estudo gerado com sucesso a partir do cenário guardado.", "success")
        except Exception as exc:
            QMessageBox.critical(self, "Nesting", str(exc))

    def _accept_dialog(self) -> None:
        if self.result_data:
            self._save_current_study(silent=True)
        self.accept()

    def _reload_nesting_settings(self, selected_sheet_name: str = "") -> None:
        self.settings = dict(self.backend.laser_quote_settings() or {})
        self.nesting_options = default_nesting_options(self.settings)
        self.sheet_profiles = list(self.nesting_options.get("sheet_profiles", []) or [])
        self._refresh_sheet_profiles(selected_sheet_name)
        self.auto_sheet_check.setChecked(bool(self.nesting_options.get("auto_select_sheet", False)))
        self.mirror_check.setChecked(bool(self.nesting_options.get("allow_mirror", True)))
        self.free_angle_check.setChecked(bool(self.nesting_options.get("free_angle_rotation", True)))
        self.use_stock_check.setChecked(bool(self.nesting_options.get("use_stock_first", False)))
        self.allow_purchase_check.setChecked(bool(self.nesting_options.get("allow_purchase_fallback", True)))
        self.shape_check.setChecked(bool(self.nesting_options.get("shape_aware", True)))
        self.shape_strict_check.setChecked(bool(self.nesting_options.get("shape_strict", False)))
        self.shape_grid_spin.setValue(float(self.nesting_options.get("shape_grid_mm", 10.0) or 10.0))
        self.common_line_check.setChecked(bool(self.nesting_options.get("common_line_estimate", True)))
        self.common_line_tol_spin.setValue(float(self.nesting_options.get("common_line_tolerance_mm", 1.0) or 1.0))
        self.lead_opt_check.setChecked(bool(self.nesting_options.get("lead_optimization", True)))
        self.lead_opt_pct_spin.setValue(float(self.nesting_options.get("lead_optimization_pct", 8.0) or 8.0))
        self._sync_sheet_mode()
        self._sync_shape_mode()
        self._sync_process_mode()

    def _persist_nesting_preferences(self) -> None:
        settings = dict(self.settings or {})
        nesting = dict(settings.get("nesting", {}) or {})
        nesting["sheet_profiles"] = [dict(row or {}) for row in list(self.sheet_profiles or [])]
        nesting["default_part_spacing_mm"] = float(self.spacing_spin.value())
        nesting["default_edge_margin_mm"] = float(self.edge_spin.value())
        nesting["allow_rotate"] = bool(self.rotate_check.isChecked())
        nesting["allow_mirror"] = bool(self.mirror_check.isChecked())
        nesting["free_angle_rotation"] = bool(self.free_angle_check.isChecked())
        nesting["auto_select_sheet"] = bool(self.auto_sheet_check.isChecked())
        nesting["use_stock_first"] = bool(self.use_stock_check.isChecked())
        nesting["allow_purchase_fallback"] = bool(self.allow_purchase_check.isChecked())
        nesting["shape_aware"] = bool(self.shape_check.isChecked())
        nesting["shape_strict"] = bool(self.shape_strict_check.isChecked())
        nesting["shape_grid_mm"] = float(self.shape_grid_spin.value())
        nesting["common_line_estimate"] = bool(self.common_line_check.isChecked())
        nesting["common_line_tolerance_mm"] = float(self.common_line_tol_spin.value())
        nesting["lead_optimization"] = bool(self.lead_opt_check.isChecked())
        nesting["lead_optimization_pct"] = float(self.lead_opt_pct_spin.value())
        settings["nesting"] = nesting
        self.settings = dict(self.backend.laser_quote_save_settings(settings) or {})
        self.nesting_options = default_nesting_options(self.settings)
        self.sheet_profiles = list(self.nesting_options.get("sheet_profiles", []) or [])

    def _refresh_sheet_profiles(self, selected_name: str = "") -> None:
        current_name = selected_name or self.sheet_combo.currentText().strip()
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        for profile in list(self.sheet_profiles or []):
            self.sheet_combo.addItem(str(profile.get("name", "")), dict(profile))
        if self.sheet_combo.count() > 0:
            target_name = current_name or str(self.sheet_profiles[0].get("name", "") or "")
            index = self.sheet_combo.findText(target_name)
            self.sheet_combo.setCurrentIndex(index if index >= 0 else 0)
        self.sheet_combo.blockSignals(False)

    def _sync_sheet_mode(self) -> None:
        self.sheet_combo.setEnabled(not bool(self.auto_sheet_check.isChecked()))

    def _sync_shape_mode(self) -> None:
        enabled = bool(self.shape_check.isChecked())
        self.shape_grid_spin.setEnabled(enabled)
        self.shape_strict_check.setEnabled(enabled)
        self.mirror_check.setEnabled(enabled)
        self.free_angle_check.setEnabled(enabled)
        if not enabled:
            self.shape_strict_check.setChecked(False)

    def _sync_process_mode(self) -> None:
        self.common_line_tol_spin.setEnabled(bool(self.common_line_check.isChecked()))
        self.lead_opt_pct_spin.setEnabled(bool(self.lead_opt_check.isChecked()))

    def _configure_sheet_profiles(self) -> None:
        dialog = SheetProfilesDialog(self.sheet_profiles, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.sheet_profiles = dialog.result_profiles()
        self._persist_nesting_preferences()
        self._refresh_sheet_profiles()
        self._sync_sheet_mode()

    def _configure_laser_settings(self) -> None:
        dialog = LaserSettingsDialog(self.backend, self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected_name = self.sheet_combo.currentText().strip()
        self._reload_nesting_settings(selected_name)

    def _current_group(self) -> dict[str, Any]:
        return dict(self.group_combo.currentData() or {})

    def _selected_sheet_profile(self) -> dict[str, Any]:
        return dict(self.sheet_combo.currentData() or {})

    def _stock_candidates_for_group(self, group: dict[str, Any]) -> list[dict[str, Any]]:
        material = str(group.get("material", "") or "").strip()
        thickness = str(group.get("thickness_mm", "") or "").strip()
        if not material or not thickness:
            return []
        if hasattr(self.backend, "laser_sheet_stock_candidates"):
            return [dict(row or {}) for row in list(self.backend.laser_sheet_stock_candidates(material, thickness) or [])]
        candidates: list[dict[str, Any]] = []
        if hasattr(self.backend, "material_candidates"):
            for row in list(self.backend.material_candidates(material, thickness) or []):
                width_mm = float(row.get("largura", 0) or 0)
                height_mm = float(row.get("comprimento", 0) or 0)
                disponivel = int(float(row.get("disponivel", 0) or 0))
                if width_mm <= 0 or height_mm <= 0 or disponivel <= 0:
                    continue
                is_retalho = bool(row.get("is_retalho"))
                candidates.append(
                    {
                        "name": f"{'Retalho' if is_retalho else 'Stock'} {row.get('material_id', '-')}",
                        "source_kind": "retalho" if is_retalho else "stock",
                        "source_label": f"{'Retalho' if is_retalho else 'Stock'} {row.get('material_id', '-')}",
                        "material_id": str(row.get("material_id", "") or "").strip(),
                        "lote": str(row.get("lote", "") or "").strip(),
                        "local": str(row.get("local", "") or "").strip(),
                        "width_mm": round(width_mm, 3),
                        "height_mm": round(height_mm, 3),
                        "quantity_available": disponivel,
                        "is_retalho": is_retalho,
                    }
                )
        return candidates

    def _part_override_key(self, row: dict[str, Any]) -> str:
        return "|".join(
            [
                str(row.get("ref_externa", "") or "").strip(),
                str(row.get("desenho", "") or "").strip(),
                str(row.get("descricao", "") or "").strip(),
            ]
        )

    def _part_rotation_override(self, row: dict[str, Any]) -> str:
        key = self._part_override_key(row)
        current = str(self.part_rotation_overrides.get(key, str(row.get("nest_rotation_policy", "auto") or "auto")) or "auto").strip().lower()
        return current if current in {"auto", "0", "90"} else "auto"

    def _set_part_rotation_override(self, row_key: str, value: str) -> None:
        normalized = str(value or "auto").strip().lower()
        self.part_rotation_overrides[row_key] = normalized if normalized in {"auto", "0", "90"} else "auto"

    def _part_priority_override(self, row: dict[str, Any]) -> int:
        key = self._part_override_key(row)
        raw = self.part_priority_overrides.get(key, row.get("nest_priority", row.get("priority", row.get("prioridade", 0))))
        try:
            value = int(raw)
        except Exception:
            value = 0
        return max(-1, min(2, value))

    def _set_part_priority_override(self, row_key: str, value: int) -> None:
        try:
            normalized = int(value)
        except Exception:
            normalized = 0
        self.part_priority_overrides[row_key] = max(-1, min(2, normalized))

    def _programmed_counts_by_part(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for sheet in list(self.result_data.get("sheets", []) or []):
            for placement in list(dict(sheet or {}).get("placements", []) or []):
                key = "|".join(
                    [
                        str(placement.get("ref_externa", "") or "").strip(),
                        str(placement.get("path", "") or "").strip(),
                        str(placement.get("description", "") or "").strip(),
                    ]
                )
                counts[key] = counts.get(key, 0) + 1
        return counts

    def _part_programming_status(self, row: dict[str, Any]) -> tuple[int, int]:
        requested = max(0, int(round(float(row.get("qtd", 0) or 0))))
        programmed = int(self._programmed_counts_by_part().get(self._part_override_key(row), 0) or 0)
        return programmed, requested

    def _resolve_drawing_path(self, raw: str) -> str:
        path_txt = str(raw or "").strip()
        if not path_txt:
            return ""
        try:
            resolver = getattr(self.backend, "_resolve_file_reference", None)
            if callable(resolver):
                resolved = resolver(path_txt)
                if resolved is not None:
                    return str(resolved)
        except Exception:
            pass
        return path_txt

    def _part_preview_payload(self, path: str, geometry: dict[str, Any] | None = None) -> dict[str, Any]:
        resolved_path = self._resolve_drawing_path(path)
        cache_key = str(resolved_path or path or "").strip()
        if geometry is None and cache_key and cache_key in self.part_geometry_cache:
            return dict(self.part_geometry_cache[cache_key])
        payload: dict[str, Any] = {"width_mm": 0.0, "height_mm": 0.0, "outer": [], "holes": [], "all": [], "paths": [], "preview_mode": "solid"}
        raw_geometry = dict(geometry or {})
        bbox = dict(raw_geometry.get("bbox_mm", {}) or {})
        nesting_shape = dict(raw_geometry.get("nesting_shape", {}) or {})
        preview_paths = dict(raw_geometry.get("preview_paths", {}) or {})
        metrics = dict(raw_geometry.get("metrics", {}) or {})
        min_x = float(bbox.get("min_x", 0) or 0)
        min_y = float(bbox.get("min_y", 0) or 0)
        outer_polygons: list[list[QPointF]] = []
        hole_polygons: list[list[QPointF]] = []
        cut_paths: list[list[QPointF]] = []
        for raw_polygon in list(nesting_shape.get("outer_polygons", []) or []):
            polygon = [QPointF(float(point[0]) - min_x, float(point[1]) - min_y) for point in list(raw_polygon or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
            if len(polygon) >= 3:
                outer_polygons.append(polygon)
        for raw_polygon in list(nesting_shape.get("hole_polygons", []) or []):
            polygon = [QPointF(float(point[0]) - min_x, float(point[1]) - min_y) for point in list(raw_polygon or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
            if len(polygon) >= 3:
                hole_polygons.append(polygon)
        for raw_path in list(preview_paths.get("cut_paths", []) or []):
            path_points = [QPointF(float(point[0]) - min_x, float(point[1]) - min_y) for point in list(raw_path or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
            if len(path_points) >= 2:
                cut_paths.append(path_points)
        width_mm = float(bbox.get("width", 0) or 0)
        height_mm = float(bbox.get("height", 0) or 0)
        if not outer_polygons and not cut_paths and width_mm > 0 and height_mm > 0:
            outer_polygons = [[QPointF(0.0, 0.0), QPointF(width_mm, 0.0), QPointF(width_mm, height_mm), QPointF(0.0, height_mm)]]
        all_polygons = list(outer_polygons) + list(hole_polygons)
        preview_mode = "linework" if cut_paths and not outer_polygons else "multi_contour" if int(metrics.get("outer_contours", 0) or 0) > 1 else "solid"
        payload.update(
            {
                "width_mm": width_mm,
                "height_mm": height_mm,
                "outer": outer_polygons,
                "holes": hole_polygons,
                "all": all_polygons,
                "paths": cut_paths,
                "preview_mode": preview_mode,
            }
        )
        if cache_key:
            self.part_geometry_cache[cache_key] = dict(payload)
        return payload

    def _build_part_preview_pixmap(self, preview: dict[str, Any], programmed: int, requested: int, *, width: int = 92, height: int = 40) -> QPixmap:
        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        bg_rect = QRectF(0.5, 0.5, width - 1.0, height - 1.0)
        painter.setPen(QPen(QColor("#c6d2e0"), 1))
        painter.setBrush(QColor("#f8fbff"))
        painter.drawRoundedRect(bg_rect, 8, 8)

        outer_polygons = list(preview.get("outer", []) or [])
        hole_polygons = list(preview.get("holes", []) or [])
        all_polygons = list(preview.get("all", []) or [])
        preview_paths = list(preview.get("paths", []) or [])
        preview_mode = str(preview.get("preview_mode", "solid") or "solid").strip().lower()
        if not outer_polygons and not preview_paths:
            painter.setPen(QPen(QColor("#94a3b8"), 1, Qt.DashLine))
            painter.setBrush(Qt.NoBrush)
            painter.drawRoundedRect(QRectF(8.0, 8.0, width - 16.0, height - 16.0), 6, 6)
            painter.setPen(QColor("#64748b"))
            painter.drawText(QRectF(8.0, 8.0, width - 16.0, height - 16.0), Qt.AlignCenter, "Sem\ndesenho")
            painter.end()
            return pixmap
        shape_width = max(1.0, float(preview.get("width_mm", 0.0) or 0.0))
        shape_height = max(1.0, float(preview.get("height_mm", 0.0) or 0.0))
        display_outer = outer_polygons
        display_holes = hole_polygons
        display_all = all_polygons
        display_paths = preview_paths
        display_width = shape_width
        display_height = shape_height
        if shape_height > (shape_width * 1.6):
            display_width = shape_height
            display_height = shape_width
            display_outer = [[QPointF(point.y(), shape_width - point.x()) for point in polygon_points] for polygon_points in outer_polygons]
            display_holes = [[QPointF(point.y(), shape_width - point.x()) for point in polygon_points] for polygon_points in hole_polygons]
            display_all = [[QPointF(point.y(), shape_width - point.x()) for point in polygon_points] for polygon_points in all_polygons]
            display_paths = [[QPointF(point.y(), shape_width - point.x()) for point in path_points] for path_points in preview_paths]
        scale = min((width - 14.0) / display_width, (height - 14.0) / display_height)
        offset_x = (width - (display_width * scale)) / 2.0
        offset_y = (height - (display_height * scale)) / 2.0

        fill_color = QColor("#3fbf7f" if requested and programmed >= requested else "#f59e0b" if programmed > 0 else "#94a3b8")
        if preview_mode == "multi_contour" and display_all:
            painter.setPen(QPen(QColor("#1e3a5f"), 1.2))
            painter.setBrush(QColor(fill_color.red(), fill_color.green(), fill_color.blue(), 90))
            for polygon_points in display_all:
                polygon = QPolygonF([QPointF(offset_x + (point.x() * scale), offset_y + (point.y() * scale)) for point in polygon_points])
                painter.drawPolygon(polygon)
        else:
            painter.setPen(QPen(QColor("#1e3a5f"), 1))
            painter.setBrush(fill_color)
            for polygon_points in display_outer:
                polygon = QPolygonF([QPointF(offset_x + (point.x() * scale), offset_y + (point.y() * scale)) for point in polygon_points])
                painter.drawPolygon(polygon)
            if display_holes:
                painter.setPen(QPen(QColor("#f8fbff"), 1))
                painter.setBrush(QColor("#f8fbff"))
                for polygon_points in display_holes:
                    polygon = QPolygonF([QPointF(offset_x + (point.x() * scale), offset_y + (point.y() * scale)) for point in polygon_points])
                    painter.drawPolygon(polygon)
        if display_paths:
            painter.setPen(QPen(QColor("#475569"), 0.9))
            painter.setBrush(Qt.NoBrush)
            for path_points in display_paths:
                polyline = QPolygonF([QPointF(offset_x + (point.x() * scale), offset_y + (point.y() * scale)) for point in path_points])
                painter.drawPolyline(polyline)
        painter.end()
        return pixmap

    def _populate_stock_table(self) -> None:
        self.stock_table.setRowCount(len(self.stock_candidates))
        if not self.stock_candidates:
            self.stock_hint.setText("Sem stock de chapa disponível para este material/espessura.")
            self.stock_empty_label.show()
            self.stock_table.hide()
        else:
            self.stock_hint.setText(
                f"{len(self.stock_candidates)} formato(s) disponíveis em stock para este grupo."
            )
            self.stock_empty_label.hide()
            self.stock_table.show()
        for row_index, row in enumerate(self.stock_candidates):
            values = [
                SOURCE_LABELS.get(str(row.get("source_kind", "") or "").strip(), "-"),
                str(row.get("lote", "") or row.get("material_id", "") or "-").strip() or "-",
                self._sheet_dimension_label(float(row.get("width_mm", 0) or 0), float(row.get("height_mm", 0) or 0)),
                "Irregular" if list(row.get("outer_polygons", []) or []) else "Retangular",
                str(int(row.get("quantity_available", 0) or 0)),
                str(row.get("local", "") or "-").strip() or "-",
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(int((Qt.AlignCenter if col_index in (0, 1, 3, 4) else Qt.AlignLeft) | Qt.AlignVCenter))
                self.stock_table.setItem(row_index, col_index, item)

    def _load_group_rows(self) -> None:
        group = self._current_group()
        rows = [dict(row or {}) for row in list(group.get("rows", []) or [])]
        study = dict(self.saved_studies.get(str(group.get("key", "") or "").strip(), {}) or {})
        if study:
            self._apply_study_controls(study)
        self.current_group_rows = rows
        requested_total = sum(max(0, int(round(float(row.get("qtd", 0) or 0)))) for row in rows)
        group_label = str(group.get("label", "") or "Grupo sem identificação").strip() or "Grupo sem identificação"
        self.group_badge_label.setText(group_label)
        self.define_metric_group.setText(group_label)
        self.define_metric_forms.setText(str(len(rows)))
        self.define_metric_qty.setText(str(requested_total))
        default_spacing = max(0.0, float(self.nesting_options.get("default_part_spacing_mm", 8.0) or 8.0))
        default_edge = max(0.0, float(self.nesting_options.get("default_edge_margin_mm", 8.0) or 8.0))
        self._clear_result_view()
        if not study:
            self.spacing_spin.setValue(default_spacing)
            self.edge_spin.setValue(default_edge)
            self.part_rotation_overrides = {}
            self.part_priority_overrides = {}
        self.parts_table.setRowCount(len(rows))
        layer_rules = dict(self.settings.get("layer_rules", {}) or {})
        for row_index, row in enumerate(rows):
            path = str(row.get("desenho", "") or "").strip()
            resolved_path = self._resolve_drawing_path(path)
            bbox_txt = "-"
            bbox_exact_txt = "-"
            area_txt = "-"
            preview_payload = self._part_preview_payload(resolved_path or path)
            try:
                from lugest_core.laser.quote_engine import analyze_dxf_geometry

                geometry = analyze_dxf_geometry(resolved_path or path, layer_rules)
                bbox = dict(geometry.get("bbox_mm", {}) or {})
                metrics = dict(geometry.get("metrics", {}) or {})
                bbox_txt = f"{_fmt_num(bbox.get('width', 0), 1)} x {_fmt_num(bbox.get('height', 0), 1)} mm"
                bbox_exact_txt = f"{float(bbox.get('width', 0) or 0):.3f} x {float(bbox.get('height', 0) or 0):.3f} mm"
                area_txt = f"{_fmt_num((metrics.get('net_area_m2', 0) or 0), 4)} m2"
                preview_payload = self._part_preview_payload(resolved_path or path, geometry)
            except Exception:
                pass
            path_name = Path(resolved_path or path).name if (resolved_path or path) else "-"
            programmed, requested = self._part_programming_status(row)
            preview_label = QLabel()
            preview_label.setAlignment(Qt.AlignCenter)
            preview_label.setPixmap(self._build_part_preview_pixmap(preview_payload, programmed, requested, width=168, height=82))
            preview_label.setToolTip(f"{path_name}\n{bbox_txt}\nProgramadas: {programmed} / {requested}")
            self.parts_table.setCellWidget(row_index, 0, preview_label)

            ref_item = QTableWidgetItem(str(row.get("ref_externa", "") or "").strip())
            ref_item.setToolTip(path or ref_item.text() or "")
            desc_item = QTableWidgetItem(str(row.get("descricao", "") or "").strip())
            desc_item.setToolTip(path or desc_item.text() or "")
            prog_item = QTableWidgetItem(f"{programmed} / {requested}")
            qty_item = QTableWidgetItem(f"{float(row.get('qtd', 0) or 0):.2f}")
            bbox_item = QTableWidgetItem(bbox_txt)
            bbox_item.setToolTip(f"{bbox_txt}\nExato: {bbox_exact_txt}")
            area_item = QTableWidgetItem(area_txt)
            area_item.setToolTip(f"{path_name}\nÁrea real: {area_txt}")
            for target_col, item, align_center in (
                (1, ref_item, False),
                (2, desc_item, False),
                (3, prog_item, True),
                (4, qty_item, True),
                (6, bbox_item, True),
                (7, area_item, True),
            ):
                item.setTextAlignment(int((Qt.AlignCenter if align_center else Qt.AlignLeft) | Qt.AlignVCenter))
                self.parts_table.setItem(row_index, target_col, item)
            priority_combo = QComboBox()
            priority_combo.addItem("Baixa", -1)
            priority_combo.addItem("Normal", 0)
            priority_combo.addItem("Alta", 1)
            priority_combo.addItem("Crítica", 2)
            priority_combo.setProperty("compact", "true")
            priority_combo.setMinimumWidth(144)
            priority_combo.setMaximumWidth(170)
            priority_combo.setMinimumHeight(40)
            row_key = self._part_override_key(row)
            current_priority = self._part_priority_override(row)
            priority_index = max(0, priority_combo.findData(current_priority))
            priority_combo.setCurrentIndex(priority_index)
            priority_combo.currentIndexChanged.connect(
                lambda _index, combo=priority_combo, key=row_key: self._set_part_priority_override(str(key), int(combo.currentData() or 0))
            )
            self.parts_table.setCellWidget(row_index, 5, priority_combo)
            rotation_combo = QComboBox()
            rotation_combo.addItem("Automático", "auto")
            rotation_combo.addItem("Fixar 0°", "0")
            rotation_combo.addItem("Fixar 90°", "90")
            rotation_combo.setProperty("compact", "true")
            rotation_combo.setMinimumWidth(164)
            rotation_combo.setMaximumWidth(188)
            rotation_combo.setMinimumHeight(40)
            current_policy = self._part_rotation_override(row)
            combo_index = max(0, rotation_combo.findData(current_policy))
            rotation_combo.setCurrentIndex(combo_index)
            rotation_combo.currentIndexChanged.connect(
                lambda _index, combo=rotation_combo, key=row_key: self._set_part_rotation_override(str(key), str(combo.currentData() or "auto"))
            )
            self.parts_table.setCellWidget(row_index, 8, rotation_combo)
        self.stock_candidates = self._stock_candidates_for_group(group)
        self.define_metric_stock.setText(str(len(self.stock_candidates)))
        self._populate_stock_table()
        self._refresh_parts_programming_status()
        self._restore_saved_group_result(study)

    def _refresh_parts_programming_status(self) -> None:
        counts = self._programmed_counts_by_part()
        programmed_total = 0
        requested_total = 0
        for row_index, row in enumerate(list(self.current_group_rows or [])):
            requested = max(0, int(round(float(row.get("qtd", 0) or 0))))
            programmed = int(counts.get(self._part_override_key(row), 0) or 0)
            programmed_total += programmed
            requested_total += requested
            progress_item = self.parts_table.item(row_index, 3)
            if progress_item is None:
                progress_item = QTableWidgetItem()
                progress_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.parts_table.setItem(row_index, 3, progress_item)
            progress_item.setText(f"{programmed} / {requested}")
            if requested <= 0:
                progress_item.setBackground(QBrush(QColor("#e2e8f0")))
                progress_item.setForeground(QBrush(QColor("#334155")))
            elif programmed >= requested:
                progress_item.setBackground(QBrush(QColor("#dcfce7")))
                progress_item.setForeground(QBrush(QColor("#166534")))
            elif programmed > 0:
                progress_item.setBackground(QBrush(QColor("#fef3c7")))
                progress_item.setForeground(QBrush(QColor("#92400e")))
            else:
                progress_item.setBackground(QBrush(QColor("#fee2e2")))
                progress_item.setForeground(QBrush(QColor("#b42318")))
            preview_widget = self.parts_table.cellWidget(row_index, 0)
            if isinstance(preview_widget, QLabel):
                path = str(row.get("desenho", "") or "").strip()
                resolved_path = self._resolve_drawing_path(path)
                preview_widget.setPixmap(self._build_part_preview_pixmap(self._part_preview_payload(resolved_path or path), programmed, requested, width=168, height=82))
                preview_widget.setToolTip(f"{Path(resolved_path or path).name if (resolved_path or path) else '-'}\nProgramadas: {programmed} / {requested}")
        if requested_total > 0:
            self.parts_status_label.setText(f"Programadas {programmed_total} / {requested_total} peça(s) neste grupo.")
        else:
            self.parts_status_label.setText("Sem peças válidas neste grupo.")
        self.define_metric_prog.setText(f"{programmed_total} / {requested_total}")

    def _render_sheet_preview(self) -> None:
        self.preview_scene.clear()
        if not self.result_data:
            self.preview_hint.setText("Ainda não existe um plano calculado para pré-visualizar.")
            return
        sheets = list(self.result_data.get("sheets", []) or [])
        if not sheets:
            self.preview_hint.setText("O resultado atual não tem chapas colocadas.")
            return
        sheet_index = max(0, self.sheet_view_combo.currentIndex())
        if sheet_index >= len(sheets):
            sheet_index = 0
        sheet = dict(sheets[sheet_index] or {})
        display = self._sheet_display_context(sheet)
        sheet_width = float(display.get("display_width_mm", 0) or 0)
        sheet_height = float(display.get("display_height_mm", 0) or 0)
        if sheet_width <= 0 or sheet_height <= 0:
            summary = dict(self.result_data.get("summary", {}) or {})
            fallback_display = self._sheet_display_context(summary)
            sheet_width = float(fallback_display.get("display_width_mm", 0) or 0)
            sheet_height = float(fallback_display.get("display_height_mm", 0) or 0)
        if sheet_width <= 0 or sheet_height <= 0:
            self.preview_hint.setText("Não foi possível determinar as dimensões da chapa deste cenário.")
            return
        source_label = self._sheet_combo_label(sheet, str(dict(self.result_data.get("summary", {}) or {}).get("selected_sheet_profile", {}).get("name", "") or ""))
        geometry_validation = dict(sheet.get("geometry_validation", {}) or {})
        if not geometry_validation and list(sheet.get("placements", []) or []):
            geometry_validation = _sheet_overlap_diagnostics({"placements": list(sheet.get("placements", []) or [])})
        part_in_part_pairs = int(geometry_validation.get("part_in_part_pair_count", 0) or 0)
        solid_overlap_pairs = int(geometry_validation.get("solid_overlap_pair_count", 0) or 0)
        hint_text = (
            f"{source_label} | {self._sheet_dimension_label(sheet_width, sheet_height)} | "
            f"{int(sheet.get('part_count', 0) or 0)} peça(s) | "
            f"real {_fmt_num(sheet.get('utilization_net_pct', 0), 1)}%"
        )
        if solid_overlap_pairs > 0:
            hint_text += f" | ALERTA: {solid_overlap_pairs} colisão(ões) reais"
        elif part_in_part_pairs > 0:
            hint_text += f" | {part_in_part_pairs} encaixe(s) interno(s) válido(s)"
        self.preview_hint.setText(hint_text)
        self.preview_scene.setSceneRect(0, 0, sheet_width, sheet_height)
        sheet_outer_polygons = list(display.get("display_outer_polygons", []) or [])
        sheet_hole_polygons = list(display.get("display_hole_polygons", []) or [])
        if sheet_outer_polygons:
            for polygon_points in sheet_outer_polygons:
                polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                self.preview_scene.addPolygon(polygon, QPen(QColor("#17314f"), 2), QBrush(QColor("#f8fafc")))
            for polygon_points in sheet_hole_polygons:
                polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                self.preview_scene.addPolygon(polygon, QPen(QColor("#17314f"), 2), QBrush(QColor("#ffffff")))
        else:
            self.preview_scene.addRect(0, 0, sheet_width, sheet_height, QPen(QColor("#17314f"), 2))
        for index, placement in enumerate(list(display.get("placements", []) or [])):
            color = self._sheet_preview_color(index)
            tooltip = (
                f"{placement.get('ref_externa', '-')}\n"
                f"{placement.get('display_width_mm', 0)} x {placement.get('display_height_mm', 0)} mm\n"
                f"Rotacao: {'sim' if placement.get('rotated') else 'nao'}"
            )
            outer_polygons = list(placement.get("display_outer_polygons", []) or [])
            hole_polygons = list(placement.get("display_hole_polygons", []) or [])
            preview_paths = list(placement.get("display_preview_paths", []) or [])
            if outer_polygons:
                for polygon_points in outer_polygons:
                    polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                    polygon_item = self.preview_scene.addPolygon(polygon, QPen(QColor("#274c77"), 1), QBrush(color))
                    polygon_item.setToolTip(tooltip)
                for polygon_points in hole_polygons:
                    polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                    self.preview_scene.addPolygon(polygon, QPen(QColor("#274c77"), 1), QBrush(QColor("#ffffff")))
            else:
                rect = self.preview_scene.addRect(
                    float(placement.get("display_x_mm", 0) or 0),
                    float(placement.get("display_y_mm", 0) or 0),
                    float(placement.get("display_width_mm", 0) or 0),
                    float(placement.get("display_height_mm", 0) or 0),
                    QPen(QColor("#274c77"), 1),
                    QBrush(color),
                )
                rect.setToolTip(tooltip)
            if preview_paths:
                path_pen = QPen(QColor("#475569"), 0.9)
                for path_points in preview_paths:
                    if len(list(path_points or [])) < 2:
                        continue
                    painter_path = QPainterPath(QPointF(float(path_points[0][0]), float(path_points[0][1])))
                    for point_x, point_y in list(path_points or [])[1:]:
                        painter_path.lineTo(float(point_x), float(point_y))
                    path_item = self.preview_scene.addPath(painter_path, path_pen)
                    path_item.setToolTip(tooltip)
            if float(placement.get("display_width_mm", 0) or 0) > 120 and float(placement.get("display_height_mm", 0) or 0) > 40:
                text = self.preview_scene.addText(str(placement.get("ref_externa", "") or "").strip())
                text.setDefaultTextColor(QColor("#0f172a"))
                text.setPos(float(placement.get("display_x_mm", 0) or 0) + 4, float(placement.get("display_y_mm", 0) or 0) + 4)
        self.preview_view.request_fit()
        if self.sheet_plan_table.rowCount() > sheet_index:
            self.sheet_plan_table.blockSignals(True)
            self.sheet_plan_table.clearSelection()
            self.sheet_plan_table.selectRow(sheet_index)
            self.sheet_plan_table.blockSignals(False)
        if self.sheet_gallery.count() > sheet_index:
            self.sheet_gallery.blockSignals(True)
            self.sheet_gallery.setCurrentRow(sheet_index)
            self.sheet_gallery.blockSignals(False)

    def _default_notes(self, summary: dict[str, Any], warnings: list[str]) -> list[str]:
        notes: list[str] = []
        engine_requested = str(summary.get("engine_requested", "") or "").strip().lower()
        engine_used = str(summary.get("engine_used", "") or "").strip().lower()
        tested_modes = {str(mode or "").strip().lower() for mode in list(summary.get("engine_modes_tested", []) or [])}
        if "shape" in tested_modes and "bbox" in tested_modes:
            if engine_requested == "shape" and engine_used == "bbox":
                notes.append("O contorno real foi pedido, mas este cenário caiu para bounding box por limitação geométrica do estudo atual.")
            elif engine_used == "shape":
                notes.append("O contorno real foi mantido e o nesting foi calculado sem recorrer à bounding box.")
        if bool(summary.get("shape_aware")):
            notes.append(f"Modo por contorno ativo | grelha {_fmt_num(summary.get('shape_grid_mm', 0), 2)} mm.")
        priority_rows = [
            row
            for row in list(self._current_group().get("rows", []) or [])
            if self._part_priority_override(dict(row or {})) != 0
        ]
        if priority_rows:
            notes.append(
                "Prioridades ativas por peca: "
                + ", ".join(
                    f"{str(row.get('ref_externa', '-') or '-').strip() or '-'}={PRIORITY_LABELS.get(self._part_priority_override(dict(row or {})), 'Normal')}"
                    for row in priority_rows
                )
                + "."
            )
        partial_rows = []
        for row in list(self.current_group_rows or []):
            programmed, requested = self._part_programming_status(dict(row or {}))
            if requested > 0 and programmed < requested:
                partial_rows.append(f"{str(row.get('ref_externa', '-') or '-').strip() or '-'} {programmed}/{requested}")
        if partial_rows:
            notes.append("Peças ainda não totalmente programadas: " + ", ".join(partial_rows) + ".")
        profile = dict(summary.get("selected_sheet_profile", {}) or {})
        profile_name = str(profile.get("name", "") or "").strip()
        if bool(self.auto_sheet_check.isChecked()) and list(self.result_data.get("sheet_candidates", []) or []):
            notes.append(f"Cenário escolhido: {profile_name or '-'}")
        if warnings:
            notes.extend(list(warnings))
        return notes

    def _sheet_combo_label(self, sheet_row: dict[str, Any], fallback_name: str) -> str:
        source_kind = str(sheet_row.get("source_kind", "") or "").strip().lower()
        source_label = str(sheet_row.get("source_label", "") or "").strip() or fallback_name or "-"
        return f"{SOURCE_LABELS.get(source_kind, 'Chapa')} | {source_label}"

    def _sheet_dimension_label(self, width_mm: float, height_mm: float) -> str:
        major = max(0.0, float(width_mm or 0.0), float(height_mm or 0.0))
        minor = min(max(0.0, float(width_mm or 0.0)), max(0.0, float(height_mm or 0.0)))
        return f"{_fmt_num(major, 1)} x {_fmt_num(minor, 1)} mm"

    def _sheet_display_context(self, sheet_row: dict[str, Any]) -> dict[str, Any]:
        raw_width = max(0.0, float(sheet_row.get("sheet_width_mm", 0) or 0))
        raw_height = max(0.0, float(sheet_row.get("sheet_height_mm", 0) or 0))
        sheet_outer_polygons = [list(points or []) for points in list(sheet_row.get("sheet_outer_polygons", []) or [])]
        sheet_hole_polygons = [list(points or []) for points in list(sheet_row.get("sheet_hole_polygons", []) or [])]
        transpose_axes = (not sheet_outer_polygons) and raw_height > raw_width

        def _map_polygon(polygon_points: list[tuple[float, float]] | list[list[float]]) -> list[tuple[float, float]]:
            mapped: list[tuple[float, float]] = []
            for point in list(polygon_points or []):
                if not isinstance(point, (list, tuple)) or len(point) < 2:
                    continue
                x = float(point[0] or 0.0)
                y = float(point[1] or 0.0)
                mapped.append((round(y, 3), round(x, 3)) if transpose_axes else (round(x, 3), round(y, 3)))
            return mapped

        def _map_path(path_points: list[tuple[float, float]] | list[list[float]]) -> list[tuple[float, float]]:
            mapped: list[tuple[float, float]] = []
            for point in list(path_points or []):
                if not isinstance(point, (list, tuple)) or len(point) < 2:
                    continue
                x = float(point[0] or 0.0)
                y = float(point[1] or 0.0)
                mapped.append((round(y, 3), round(x, 3)) if transpose_axes else (round(x, 3), round(y, 3)))
            return mapped

        def _map_rect(x_mm: float, y_mm: float, width_mm: float, height_mm: float) -> tuple[float, float, float, float]:
            if transpose_axes:
                return (round(y_mm, 3), round(x_mm, 3), round(height_mm, 3), round(width_mm, 3))
            return (round(x_mm, 3), round(y_mm, 3), round(width_mm, 3), round(height_mm, 3))

        placements: list[dict[str, Any]] = []
        for placement in list(sheet_row.get("placements", []) or []):
            payload = dict(placement or {})
            mapped_x, mapped_y, mapped_width, mapped_height = _map_rect(
                float(payload.get("x_mm", 0) or 0),
                float(payload.get("y_mm", 0) or 0),
                float(payload.get("width_mm", 0) or 0),
                float(payload.get("height_mm", 0) or 0),
            )
            payload["display_x_mm"] = mapped_x
            payload["display_y_mm"] = mapped_y
            payload["display_width_mm"] = mapped_width
            payload["display_height_mm"] = mapped_height
            payload["display_outer_polygons"] = [_map_polygon(list(points or [])) for points in list(payload.get("shape_outer_polygons", []) or [])]
            payload["display_hole_polygons"] = [_map_polygon(list(points or [])) for points in list(payload.get("shape_hole_polygons", []) or [])]
            payload["display_preview_paths"] = [_map_path(list(points or [])) for points in list(payload.get("preview_paths", []) or [])]
            placements.append(payload)

        return {
            "raw_width_mm": raw_width,
            "raw_height_mm": raw_height,
            "display_width_mm": raw_height if transpose_axes else raw_width,
            "display_height_mm": raw_width if transpose_axes else raw_height,
            "display_outer_polygons": [_map_polygon(list(points or [])) for points in sheet_outer_polygons],
            "display_hole_polygons": [_map_polygon(list(points or [])) for points in sheet_hole_polygons],
            "placements": placements,
            "transposed": transpose_axes,
        }

    def _sheet_preview_color(self, index: int) -> QColor:
        return QColor(PREVIEW_PALETTE[index % len(PREVIEW_PALETTE)])

    def _build_sheet_thumbnail(self, sheet_row: dict[str, Any]) -> QPixmap:
        thumb_width = 220
        thumb_height = 130
        pixmap = QPixmap(thumb_width, thumb_height)
        pixmap.fill(QColor("#ffffff"))
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)

        canvas = QRectF(8.0, 8.0, float(thumb_width - 16), float(thumb_height - 16))
        display = self._sheet_display_context(sheet_row)
        sheet_width = max(0.0, float(display.get("display_width_mm", 0) or 0))
        sheet_height = max(0.0, float(display.get("display_height_mm", 0) or 0))
        if sheet_width <= 0.0 or sheet_height <= 0.0:
            painter.setPen(QColor("#475569"))
            painter.drawText(canvas, Qt.AlignCenter, "Sem geometria")
            painter.end()
            return pixmap

        scale = min(canvas.width() / sheet_width, canvas.height() / sheet_height)
        draw_width = sheet_width * scale
        draw_height = sheet_height * scale
        offset_x = canvas.x() + ((canvas.width() - draw_width) / 2.0)
        offset_y = canvas.y() + ((canvas.height() - draw_height) / 2.0)

        def _map_point(point_x: float, point_y: float) -> QPointF:
            return QPointF(offset_x + (point_x * scale), offset_y + (point_y * scale))

        sheet_pen = QPen(QColor("#17314f"), 1.4)
        placement_pen = QPen(QColor("#274c77"), 1.0)
        sheet_outer_polygons = list(display.get("display_outer_polygons", []) or [])
        sheet_hole_polygons = list(display.get("display_hole_polygons", []) or [])
        if sheet_outer_polygons:
            for polygon_points in sheet_outer_polygons:
                polygon = QPolygonF([_map_point(float(x), float(y)) for x, y in list(polygon_points or [])])
                painter.setPen(sheet_pen)
                painter.setBrush(QColor("#f8fafc"))
                painter.drawPolygon(polygon)
            for polygon_points in sheet_hole_polygons:
                polygon = QPolygonF([_map_point(float(x), float(y)) for x, y in list(polygon_points or [])])
                painter.setPen(sheet_pen)
                painter.setBrush(QColor("#ffffff"))
                painter.drawPolygon(polygon)
        else:
            painter.setPen(sheet_pen)
            painter.setBrush(QColor("#f8fafc"))
            painter.drawRect(QRectF(offset_x, offset_y, draw_width, draw_height))

        for index, placement in enumerate(list(display.get("placements", []) or [])):
            outer_polygons = list(placement.get("display_outer_polygons", []) or [])
            hole_polygons = list(placement.get("display_hole_polygons", []) or [])
            preview_paths = list(placement.get("display_preview_paths", []) or [])
            painter.setPen(placement_pen)
            painter.setBrush(self._sheet_preview_color(index))
            if outer_polygons:
                for polygon_points in outer_polygons:
                    polygon = QPolygonF([_map_point(float(x), float(y)) for x, y in list(polygon_points or [])])
                    painter.drawPolygon(polygon)
                painter.setBrush(QColor("#ffffff"))
                for polygon_points in hole_polygons:
                    polygon = QPolygonF([_map_point(float(x), float(y)) for x, y in list(polygon_points or [])])
                    painter.drawPolygon(polygon)
            else:
                painter.drawRect(
                    QRectF(
                        offset_x + (float(placement.get("display_x_mm", 0) or 0) * scale),
                        offset_y + (float(placement.get("display_y_mm", 0) or 0) * scale),
                        float(placement.get("display_width_mm", 0) or 0) * scale,
                        float(placement.get("display_height_mm", 0) or 0) * scale,
                    )
                )
            if preview_paths:
                painter.setPen(QPen(QColor("#475569"), 0.8))
                painter.setBrush(Qt.NoBrush)
                for path_points in preview_paths:
                    polyline = QPolygonF([_map_point(float(x), float(y)) for x, y in list(path_points or [])])
                    painter.drawPolyline(polyline)
        painter.end()
        return pixmap

    def _populate_sheet_gallery(self, fallback_name: str) -> None:
        sheets = [dict(row or {}) for row in list(self.result_data.get("sheets", []) or [])]
        self.sheet_gallery.blockSignals(True)
        self.sheet_gallery.clear()
        for row_index, sheet_row in enumerate(sheets):
            label = (
                f"Chapa {int(sheet_row.get('index', row_index + 1) or (row_index + 1))}\n"
                f"{self._sheet_combo_label(sheet_row, fallback_name)}\n"
                f"{int(sheet_row.get('part_count', 0) or 0)} peça(s) | real {_fmt_num(sheet_row.get('utilization_net_pct', 0), 1)}%"
            )
            item = QListWidgetItem(QIcon(self._build_sheet_thumbnail(sheet_row)), label)
            item.setTextAlignment(int(Qt.AlignHCenter | Qt.AlignTop))
            item.setData(Qt.UserRole, row_index)
            self.sheet_gallery.addItem(item)
        self.sheet_gallery.blockSignals(False)

    def _clear_result_view(self) -> None:
        self.result_data = {}
        for label in self.summary_labels.values():
            label.setText("-")
        self.warning_edit.clear()
        self.candidate_table.setRowCount(0)
        self.sheet_plan_table.setRowCount(0)
        self.unplaced_table.setRowCount(0)
        self.cost_table.setRowCount(0)
        self.process_table.setRowCount(0)
        self.decision_edit.clear()
        self.sheet_view_combo.blockSignals(True)
        self.sheet_view_combo.clear()
        self.sheet_view_combo.blockSignals(False)
        self.sheet_gallery.blockSignals(True)
        self.sheet_gallery.clear()
        self.sheet_gallery.blockSignals(False)
        self.preview_hint.setText("Seleciona uma chapa analisada para ver o plano, a origem e o aproveitamento. Usa a roda do rato para zoom e duplo clique para enquadrar.")
        self.preview_scene.clear()
        self.nest_tabs.setCurrentIndex(0)
        self.results_tabs.setCurrentIndex(0)

    def _current_wizard_step_index(self) -> int:
        current = (
            int(self.section_stack.currentIndex()),
            int(self.define_tabs.currentIndex()) if self.section_stack.currentIndex() == 0 else int(self.nest_tabs.currentIndex()) if self.section_stack.currentIndex() == 1 else int(self.results_tabs.currentIndex()),
        )
        for index, (section_index, sub_index, _section_label, _page_label) in enumerate(WIZARD_STEPS):
            if current == (section_index, sub_index):
                return index
        return 0

    def _set_wizard_step(self, step_index: int) -> None:
        target = max(0, min(int(step_index), len(WIZARD_STEPS) - 1))
        section_index, sub_index, _section_label, _page_label = WIZARD_STEPS[target]
        self.section_stack.blockSignals(True)
        self.define_tabs.blockSignals(True)
        self.nest_tabs.blockSignals(True)
        self.results_tabs.blockSignals(True)
        self.section_stack.setCurrentIndex(section_index)
        if section_index == 0:
            self.define_tabs.setCurrentIndex(sub_index)
        elif section_index == 1:
            self.nest_tabs.setCurrentIndex(sub_index)
        else:
            self.results_tabs.setCurrentIndex(sub_index)
        self.section_stack.blockSignals(False)
        self.define_tabs.blockSignals(False)
        self.nest_tabs.blockSignals(False)
        self.results_tabs.blockSignals(False)
        self._sync_wizard_controls()

    def _jump_to_major_step(self, section_index: int) -> None:
        current = int(self.section_stack.currentIndex())
        if int(section_index) == current:
            return
        if int(section_index) > 0 and not self.result_data:
            QMessageBox.information(self, "Planeador de Chapa", "Gera primeiro o nesting para avançar para as fases seguintes.")
            return
        target_step = {0: 0, 1: 2, 2: 4}.get(int(section_index), 0)
        self._set_wizard_step(target_step)

    def _handle_section_changed(self, _index: int) -> None:
        self._sync_wizard_controls()

    def _handle_nest_tab_changed(self, index: int) -> None:
        self._sync_wizard_controls()
        if int(index) == 1 and self.result_data:
            QTimer.singleShot(0, self._render_sheet_preview)

    def _sync_wizard_controls(self) -> None:
        step_index = self._current_wizard_step_index()
        section_label, page_label = WIZARD_STEPS[step_index][2], WIZARD_STEPS[step_index][3]
        self.wizard_path_label.setText("")
        current_section = int(self.section_stack.currentIndex())
        self.page_title_label.setText(page_label)
        self.page_subtitle_label.setText("")
        enabled_sections = {0}
        if self.result_data or current_section >= 1:
            enabled_sections.add(1)
        if self.result_data or current_section >= 2:
            enabled_sections.add(2)
        for section_index, button in enumerate(self.major_step_buttons):
            active = section_index == current_section
            button.setChecked(active)
            button.setEnabled(section_index in enabled_sections)
            if active:
                button.setStyleSheet(
                    "QPushButton {background: #7ed321; color: #ffffff; border: 1px solid #6ab619; "
                    "border-radius: 22px; padding: 10px 18px; font-size: 15px; font-weight: 900;}"
                )
            elif section_index in enabled_sections:
                button.setStyleSheet(
                    "QPushButton {background: #ffffff; color: #2f4b1d; border: 1px solid #cfe4a4; "
                    "border-radius: 22px; padding: 10px 18px; font-size: 15px; font-weight: 800;}"
                    "QPushButton:hover {background: #f6fce8;}"
                )
            else:
                button.setStyleSheet(
                    "QPushButton {background: #f3f4f6; color: #94a3b8; border: 1px solid #d7dce2; "
                    "border-radius: 22px; padding: 10px 18px; font-size: 15px; font-weight: 800;}"
                )
        self.prev_section_btn.setEnabled(step_index > 0)
        self.next_section_btn.setEnabled(step_index < (len(WIZARD_STEPS) - 1))
        if step_index == 0:
            self.next_section_btn.setText("Seguinte: Peças e stock")
        elif step_index == 1:
            self.next_section_btn.setText("Seguinte: Gerar nesting")
        elif step_index == 2:
            self.next_section_btn.setText("Seguinte: Layouts / chapas")
        elif step_index == 3:
            self.next_section_btn.setText("Seguinte: Pendentes / avisos")
        elif step_index == 4:
            self.next_section_btn.setText("Seguinte: Custos / decisão")
        else:
            self.next_section_btn.setText("Concluir")
        self.analyze_btn.setText("Atualizar nesting" if step_index >= 2 else "Iniciar otimização")
        if self.section_stack.currentIndex() == 1 and self.nest_tabs.currentIndex() == 1 and self.result_data:
            QTimer.singleShot(0, self._render_sheet_preview)

    def _go_previous_section(self) -> None:
        self._set_wizard_step(self._current_wizard_step_index() - 1)

    def _go_next_section(self) -> None:
        step_index = self._current_wizard_step_index()
        if step_index == 1:
            self._analyze()
            return
        if step_index < (len(WIZARD_STEPS) - 1):
            self._set_wizard_step(step_index + 1)

    def _set_nesting_busy(self, busy: bool) -> None:
        for widget in (
            self.prev_section_btn,
            self.next_section_btn,
            self.analyze_btn,
            self.pdf_btn,
            self.sheet_library_btn,
            self.dxf_settings_btn,
            self.group_combo,
            self.sheet_combo,
        ):
            widget.setEnabled(not busy)
        if hasattr(self, "dialog_buttons"):
            self.dialog_buttons.setEnabled(not busy)
        self.nesting_progress.setVisible(busy)
        if busy:
            self._nesting_progress_value = 0
            self.nesting_progress.setValue(0)
            self.nesting_progress.setFormat("A otimizar nesting... %p%")
            self._nesting_progress_timer.start()
        else:
            self._nesting_progress_timer.stop()
            self.nesting_progress.setValue(0)

    def _tick_nesting_progress(self) -> None:
        if self._nesting_thread is None:
            return
        self._nesting_progress_value = min(95, int(self._nesting_progress_value or 0) + 1)
        self.nesting_progress.setValue(self._nesting_progress_value)

    def _cleanup_nesting_thread(self) -> None:
        self._nesting_worker = None
        self._nesting_thread = None
        self._set_nesting_busy(False)

    def _handle_nesting_success(self, result: dict) -> None:
        self._nesting_progress_value = 100
        self.nesting_progress.setValue(100)
        self._apply_result_data(dict(result or {}))
        self._save_current_study(silent=True)

    def _handle_nesting_error(self, message: str) -> None:
        QMessageBox.critical(self, "Nesting Laser", str(message or "Falha ao calcular nesting."))

    def _geometry_engine_label(self, summary: dict[str, Any]) -> str:
        requested = str(summary.get("engine_requested", "") or "").strip().lower()
        used = str(summary.get("engine_used", "") or "").strip().lower()
        if used not in {"shape", "bbox"}:
            used = "shape" if bool(summary.get("shape_aware")) else "bbox"
        label = "Contorno DXF" if used == "shape" else "Caixa DXF"
        if requested == "shape" and used == "bbox":
            label += " (fallback)"
        return label

    def _analysis_method_label(self, summary: dict[str, Any]) -> str:
        selection_mode = str(summary.get("selection_mode", "") or "").strip().lower()
        purchased_count = int(summary.get("purchased_sheet_count", 0) or 0)
        geometry_label = self._geometry_engine_label(summary)
        if selection_mode == "manual":
            flow_label = "Formato manual"
        elif selection_mode == "auto":
            flow_label = "Comparacao automatica"
        elif selection_mode == "manual_stock":
            flow_label = "Stock + compra" if purchased_count > 0 else "So stock/retalho"
        elif selection_mode == "auto_stock":
            flow_label = "Stock + compra automatica" if purchased_count > 0 else "So stock/retalho"
        else:
            flow_label = "Analise direta"
        return f"{geometry_label} | {flow_label}"

    def _populate_candidate_table(self, summary: dict[str, Any]) -> None:
        selected_profile = dict(summary.get("selected_sheet_profile", {}) or {})
        selected_name = str(selected_profile.get("name", "") or "").strip() or "Apenas stock"
        candidates = [dict(row or {}) for row in list(self.result_data.get("sheet_candidates", []) or [])]
        if not candidates:
            candidates = [
                {
                    "name": selected_name,
                    "method": self._analysis_method_label(summary),
                    "sheet_count": int(summary.get("sheet_count", 0) or 0),
                    "stock_sheet_count": int(summary.get("stock_sheet_count", 0) or 0),
                    "purchased_sheet_count": int(summary.get("purchased_sheet_count", 0) or 0),
                    "part_count_unplaced": int(summary.get("part_count_unplaced", 0) or 0),
                    "layout_compactness_pct": float(summary.get("layout_compactness_pct", 0) or 0),
                    "purchase_sheet_area_mm2": float(summary.get("purchase_sheet_area_mm2", 0) or 0),
                    "total_sheet_area_mm2": float(summary.get("total_sheet_area_mm2", 0) or 0),
                }
            ]

        shape_label = self._analysis_method_label(summary)
        self.candidate_table.setRowCount(len(candidates))
        for row_index, row in enumerate(candidates):
            candidate = dict(row or {})
            purchased_count = int(candidate.get("purchased_sheet_count", 0) or 0)
            if "method" in candidate:
                method_label = str(candidate.get("method", "") or "").strip() or self._analysis_method_label(summary)
            elif str(candidate.get("name", "") or "").strip().lower() == "apenas stock":
                method_label = f"{shape_label} | So stock/retalho"
            elif int(candidate.get("stock_sheet_count", 0) or 0) > 0:
                method_label = f"{shape_label} | {'Stock + compra' if purchased_count > 0 else 'Stock/retalho'}"
            elif bool(self.auto_sheet_check.isChecked()):
                method_label = f"{shape_label} | Comparacao automatica"
            else:
                method_label = f"{shape_label} | Formato manual"

            values = [
                str(candidate.get("name", "") or "-").strip() or "-",
                method_label,
                str(int(candidate.get("sheet_count", 0) or 0)),
                f"{_fmt_num(candidate.get('layout_compactness_pct', 0), 2)} %",
                f"{_fmt_num((candidate.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2",
                f"{_fmt_num((candidate.get('total_sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2",
                str(int(candidate.get("part_count_unplaced", 0) or 0)),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(int((Qt.AlignLeft if col_index in (0, 1) else Qt.AlignCenter) | Qt.AlignVCenter))
                self.candidate_table.setItem(row_index, col_index, item)

    def _populate_sheet_plan_table(self, selected_name: str) -> None:
        sheets = [dict(row or {}) for row in list(self.result_data.get("sheets", []) or [])]
        self.sheet_plan_table.setRowCount(len(sheets))
        for row_index, row in enumerate(sheets):
            source_kind = str(row.get("source_kind", "") or "").strip().lower()
            source_name = SOURCE_LABELS.get(source_kind, "Chapa")
            source_label = str(row.get("source_label", "") or "").strip() or selected_name or "-"
            profile_label = self._sheet_dimension_label(float(row.get("sheet_width_mm", 0) or 0), float(row.get("sheet_height_mm", 0) or 0))
            values = [
                str(int(row.get("index", row_index + 1) or (row_index + 1))),
                source_name,
                f"{source_label} | {profile_label}",
                str(int(row.get("part_count", 0) or 0)),
                f"{_fmt_num(row.get('utilization_net_pct', 0), 2)} %",
                f"{_fmt_num((row.get('sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2",
                "Sim" if source_kind == "purchase" else "Não",
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(int((Qt.AlignLeft if col_index == 2 else Qt.AlignCenter) | Qt.AlignVCenter))
                self.sheet_plan_table.setItem(row_index, col_index, item)

    def _populate_unplaced_table(self) -> None:
        unplaced_rows = [dict(row or {}) for row in list(self.result_data.get("unplaced", []) or [])]
        self.unplaced_table.setRowCount(len(unplaced_rows))
        for row_index, row in enumerate(unplaced_rows):
            values = [
                str(row.get("ref_externa", "") or "-").strip() or "-",
                str(row.get("description", "") or "-").strip() or "-",
                str(row.get("file_name", "") or "-").strip() or "-",
                str(int(row.get("copy_index", 0) or 0)),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(int((Qt.AlignCenter if col_index == 3 else Qt.AlignLeft) | Qt.AlignVCenter))
                self.unplaced_table.setItem(row_index, col_index, item)

    def _sync_preview_from_sheet_table(self) -> None:
        selected_rows = sorted({item.row() for item in self.sheet_plan_table.selectedItems() if item is not None})
        if not selected_rows:
            return
        target_row = selected_rows[0]
        if self.sheet_view_combo.currentIndex() == target_row:
            if self.sheet_gallery.currentRow() != target_row:
                self.sheet_gallery.blockSignals(True)
                self.sheet_gallery.setCurrentRow(target_row)
                self.sheet_gallery.blockSignals(False)
            return
        self.sheet_view_combo.setCurrentIndex(target_row)

    def _sync_preview_from_gallery(self, target_row: int) -> None:
        if target_row < 0:
            return
        if self.sheet_plan_table.rowCount() > target_row:
            self.sheet_plan_table.blockSignals(True)
            self.sheet_plan_table.clearSelection()
            self.sheet_plan_table.selectRow(target_row)
            self.sheet_plan_table.blockSignals(False)
        if self.sheet_view_combo.currentIndex() != target_row:
            self.sheet_view_combo.setCurrentIndex(target_row)

    def _candidate_score(self, row: dict[str, Any]) -> tuple[float, ...]:
        return (
            int(row.get("part_count_unplaced", 0) or 0),
            float(row.get("purchase_sheet_area_mm2", 0) or 0),
            float(row.get("total_sheet_area_mm2", 0) or 0),
            int(row.get("sheet_count", 0) or 0),
            -float(row.get("utilization_net_pct", 0) or 0),
            -float(row.get("utilization_bbox_pct", 0) or 0),
        )

    def _estimate_common_line_metrics(self) -> dict[str, Any]:
        tolerance_mm = float(self.common_line_tol_spin.value())
        if not bool(self.common_line_check.isChecked()):
            return {
                "enabled": False,
                "shared_length_mm": 0.0,
                "shared_edge_count": 0,
                "blocked_candidates": 0,
                "tolerance_mm": tolerance_mm,
            }

        shared_length_mm = 0.0
        shared_edge_count = 0
        blocked_candidates = 0
        spacing_hint = max(0.0, float(self.spacing_spin.value()))
        for sheet in list(self.result_data.get("sheets", []) or []):
            placements = [dict(row or {}) for row in list(dict(sheet or {}).get("placements", []) or [])]
            for index, left in enumerate(placements):
                lx = float(left.get("x_mm", 0) or 0)
                ly = float(left.get("y_mm", 0) or 0)
                lw = float(left.get("width_mm", 0) or 0)
                lh = float(left.get("height_mm", 0) or 0)
                for right in placements[index + 1 :]:
                    rx = float(right.get("x_mm", 0) or 0)
                    ry = float(right.get("y_mm", 0) or 0)
                    rw = float(right.get("width_mm", 0) or 0)
                    rh = float(right.get("height_mm", 0) or 0)

                    vertical_overlap = max(0.0, min(ly + lh, ry + rh) - max(ly, ry))
                    horizontal_overlap = max(0.0, min(lx + lw, rx + rw) - max(lx, rx))
                    if vertical_overlap > tolerance_mm:
                        gap_a = abs((lx + lw) - rx)
                        gap_b = abs((rx + rw) - lx)
                        gap = min(gap_a, gap_b)
                        if gap <= tolerance_mm:
                            shared_length_mm += vertical_overlap
                            shared_edge_count += 1
                        elif gap <= spacing_hint + tolerance_mm:
                            blocked_candidates += 1
                    if horizontal_overlap > tolerance_mm:
                        gap_a = abs((ly + lh) - ry)
                        gap_b = abs((ry + rh) - ly)
                        gap = min(gap_a, gap_b)
                        if gap <= tolerance_mm:
                            shared_length_mm += horizontal_overlap
                            shared_edge_count += 1
                        elif gap <= spacing_hint + tolerance_mm:
                            blocked_candidates += 1
        return {
            "enabled": True,
            "shared_length_mm": round(shared_length_mm, 3),
            "shared_edge_count": int(shared_edge_count),
            "blocked_candidates": int(blocked_candidates),
            "tolerance_mm": tolerance_mm,
        }

    def _build_cost_report(self, summary: dict[str, Any]) -> dict[str, Any]:
        rows = list(self._current_group().get("rows", []) or [])
        source_by_path: dict[str, dict[str, Any]] = {}
        for row in rows:
            path = str(row.get("desenho", "") or "").strip()
            if path and path not in source_by_path:
                source_by_path[path] = dict(row or {})

        placed_map: dict[tuple[str, str, str, str, float], dict[str, Any]] = {}
        for sheet in list(self.result_data.get("sheets", []) or []):
            for placement in list(dict(sheet or {}).get("placements", []) or []):
                payload = dict(placement or {})
                key = (
                    str(payload.get("path", "") or "").strip(),
                    str(payload.get("ref_externa", "") or "").strip(),
                    str(payload.get("description", "") or "").strip(),
                    str(payload.get("material", "") or "").strip(),
                    float(payload.get("thickness_mm", 0) or 0),
                )
                current = placed_map.setdefault(
                    key,
                    {
                        "path": key[0],
                        "ref_externa": key[1] or "-",
                        "description": key[2] or "-",
                        "material": key[3] or "-",
                        "thickness_mm": key[4],
                        "qty": 0,
                        "net_area_mm2": 0.0,
                    },
                )
                current["qty"] = int(current.get("qty", 0) or 0) + 1
                current["net_area_mm2"] = float(current.get("net_area_mm2", 0.0) or 0.0) + float(payload.get("net_area_mm2", 0.0) or 0.0)

        machine_name = str(self.settings.get("active_machine", "") or "").strip() or "-"
        commercial_name = str(self.settings.get("active_commercial", "") or "").strip() or "-"
        totals = {
            "quoted_total_eur": 0.0,
            "technical_process_eur": 0.0,
            "machine_total_min": 0.0,
            "cut_length_m": 0.0,
            "pierce_count": 0,
            "effective_cutting_eur": 0.0,
            "pierce_cost_eur": 0.0,
            "machine_runtime_eur": 0.0,
        }
        part_rows: list[dict[str, Any]] = []
        errors: list[str] = []
        reference_estimate: dict[str, Any] | None = None

        for row in sorted(placed_map.values(), key=lambda item: (str(item.get("ref_externa", "")), str(item.get("description", "")))):
            path = str(row.get("path", "") or "").strip()
            if not path:
                continue
            source_row = dict(source_by_path.get(path, {}) or {})
            material_supplied_by_client = bool(
                source_row.get("material_supplied_by_client", False)
                or source_row.get("material_fornecido_cliente", False)
                or not bool(source_row.get("material_cost_included", True))
            )
            payload = {
                "path": path,
                "material": str(row.get("material", "") or source_row.get("material", "") or "").strip(),
                "thickness_mm": float(row.get("thickness_mm", 0) or source_row.get("espessura", 0) or 0),
                "qtd": int(row.get("qty", 0) or 0),
                "quantity": int(row.get("qty", 0) or 0),
                "machine_name": machine_name,
                "commercial_name": commercial_name,
                "description": str(row.get("description", "") or source_row.get("descricao", "") or "").strip(),
                "ref_externa": str(row.get("ref_externa", "") or source_row.get("ref_externa", "") or "").strip(),
                "gas": str(source_row.get("gas", "") or "").strip(),
                "material_supplied_by_client": material_supplied_by_client,
            }
            try:
                estimate = estimate_laser_quote(payload, self.settings)
            except Exception as exc:
                errors.append(f"{row.get('ref_externa', '-')}: {exc}")
                continue
            if reference_estimate is None:
                reference_estimate = dict(estimate or {})
            metrics = dict(estimate.get("metrics", {}) or {})
            times = dict(estimate.get("times", {}) or {})
            pricing = dict(estimate.get("pricing", {}) or {})
            commercial = dict(estimate.get("commercial", {}) or {})
            quantity = int(row.get("qty", 0) or 0)
            process_eur = max(
                0.0,
                (float(pricing.get("subtotal_cost_unit", 0) or 0) - float(pricing.get("material_cost_unit", 0) or 0)) * quantity,
            )
            quoted_total = float(pricing.get("total_price", 0) or 0)
            estimated_material_total = float(pricing.get("material_cost_unit", 0) or 0) * quantity
            non_material_subtotal_total = max(
                0.0,
                (float(pricing.get("subtotal_cost_unit", 0) or 0) - float(pricing.get("material_cost_unit", 0) or 0)) * quantity,
            )
            current_quote_unit = float(source_row.get("preco_unit", pricing.get("unit_price", 0)) or 0)
            current_quote_total = current_quote_unit * quantity
            machine_total_min = float(times.get("machine_total_min", 0) or 0) * quantity
            cut_length_total = float(metrics.get("cut_length_m", 0) or 0) * quantity
            pierce_total = int(metrics.get("pierce_count", 0) or 0) * quantity
            totals["quoted_total_eur"] += quoted_total
            totals["technical_process_eur"] += process_eur
            totals["machine_total_min"] += machine_total_min
            totals["cut_length_m"] += cut_length_total
            totals["pierce_count"] += pierce_total
            totals["effective_cutting_eur"] += float(pricing.get("effective_cutting_cost_unit", 0) or 0) * quantity
            totals["pierce_cost_eur"] += float(pricing.get("pierce_cost_unit", 0) or 0) * quantity
            totals["machine_runtime_eur"] += float(pricing.get("machine_runtime_cost_unit", 0) or 0) * quantity
            part_rows.append(
                {
                    "ref_externa": str(row.get("ref_externa", "") or "-").strip() or "-",
                    "description": str(row.get("description", "") or "-").strip() or "-",
                    "qty": quantity,
                    "placed_net_area_m2": round(float(row.get("net_area_mm2", 0.0) or 0.0) / 1_000_000.0, 6),
                    "machine_total_min": round(machine_total_min, 3),
                    "cut_length_m": round(cut_length_total, 4),
                    "pierce_count": pierce_total,
                    "quoted_total_eur": round(quoted_total, 2),
                    "estimated_material_total_eur": round(estimated_material_total, 2),
                    "estimated_material_unit_eur": round(float(pricing.get("material_cost_unit", 0) or 0), 4),
                    "non_material_subtotal_total_eur": round(non_material_subtotal_total, 2),
                    "non_material_subtotal_unit_eur": round(max(0.0, non_material_subtotal_total / max(1, quantity)), 4),
                    "current_quote_total_eur": round(current_quote_total, 2),
                    "current_quote_unit_eur": round(current_quote_unit, 4),
                    "material_supplied_by_client": material_supplied_by_client,
                    "material_cost_included": not material_supplied_by_client,
                    "margin_pct": round(float(commercial.get("margin_pct", 0) or 0), 3),
                    "cost_mode": str(commercial.get("effective_cutting_label", "") or "-").strip() or "-",
                }
            )

        common_line = self._estimate_common_line_metrics()
        reference_machine = dict((reference_estimate or {}).get("machine", {}) or {})
        reference_commercial = dict((reference_estimate or {}).get("commercial", {}) or {})
        machine_profiles = dict(self.settings.get("machine_profiles", {}) or {})
        machine_profile = dict(machine_profiles.get(machine_name, {}) or {})
        if not machine_profile and machine_profiles:
            machine_profile = dict(next(iter(machine_profiles.values())) or {})
        motion = dict(machine_profile.get("motion", {}) or {})
        commercial_profiles = dict(self.settings.get("commercial_profiles", {}) or {})
        commercial_profile = dict(commercial_profiles.get(commercial_name, {}) or {})
        if not commercial_profile and commercial_profiles:
            commercial_profile = dict(next(iter(commercial_profiles.values())) or {})
        rates = dict(commercial_profile.get("rates", {}) or {})
        cost_mode = str(reference_commercial.get("cost_mode", commercial_profile.get("cost_mode", "hybrid_max")) or "hybrid_max").strip().lower()
        effective_cut_speed_m_min = max(0.1, float(reference_machine.get("effective_cut_speed_m_min", 0) or 0) or 0.1)
        lead_in_mm = max(0.0, float(motion.get("lead_in_mm", 2.0) or 2.0))
        lead_out_mm = max(0.0, float(motion.get("lead_out_mm", 2.0) or 2.0))
        lead_move_speed_mm_s = max(0.1, float(motion.get("lead_move_speed_mm_s", 3.0) or 3.0))
        pierce_sec_each = max(0.0, (float(motion.get("pierce_base_ms", 400.0) or 400.0) + (float(self._current_group().get("thickness_mm", 0) or 0) * float(motion.get("pierce_per_mm_ms", 35.0) or 35.0))) / 1000.0)
        first_gas_delay_sec = max(0.0, float(motion.get("first_gas_delay_ms", 200.0) or 200.0) / 1000.0)
        gas_delay_sec = max(0.0, float(motion.get("gas_delay_ms", 0.0) or 0.0) / 1000.0)
        motion_overhead_factor = 1.0 + (max(0.0, float(motion.get("motion_overhead_pct", 4.0) or 4.0)) / 100.0)
        cut_rate_per_m = max(0.0, float(rates.get("cut_per_m_eur", 0.0) or 0.0))
        pierce_rate = max(0.0, float(rates.get("pierce_eur", 0.0) or 0.0))
        machine_hour_eur = max(0.0, float(rates.get("machine_hour_eur", 0.0) or 0.0))
        nominal_cut_length_m = max(0.0, float(totals.get("cut_length_m", 0) or 0))
        nominal_pierce_count = max(0, int(totals.get("pierce_count", 0) or 0))
        common_line_saving_m = max(0.0, float(common_line.get("shared_length_mm", 0.0) or 0.0) / 1000.0)
        common_line_pierce_saving = min(nominal_pierce_count, int(common_line.get("shared_edge_count", 0) or 0))
        adjusted_cut_length_m = max(0.0, nominal_cut_length_m - common_line_saving_m)
        adjusted_pierce_count = max(0, nominal_pierce_count - common_line_pierce_saving)
        lead_optimization_factor = 1.0 - ((float(self.lead_opt_pct_spin.value()) / 100.0) if bool(self.lead_opt_check.isChecked()) else 0.0)
        lead_optimization_factor = max(0.5, min(1.0, lead_optimization_factor))
        nominal_lead_time_sec = nominal_pierce_count * ((lead_in_mm + lead_out_mm) / lead_move_speed_mm_s)
        adjusted_lead_time_sec = adjusted_pierce_count * ((lead_in_mm + lead_out_mm) / lead_move_speed_mm_s) * lead_optimization_factor
        nominal_pierce_time_sec = nominal_pierce_count * pierce_sec_each
        adjusted_pierce_time_sec = adjusted_pierce_count * pierce_sec_each
        nominal_gas_time_sec = (first_gas_delay_sec if nominal_pierce_count > 0 else 0.0) + (max(0, nominal_pierce_count - 1) * gas_delay_sec)
        adjusted_gas_time_sec = (first_gas_delay_sec if adjusted_pierce_count > 0 else 0.0) + (max(0, adjusted_pierce_count - 1) * gas_delay_sec)
        nominal_cut_time_sec = nominal_cut_length_m / max(0.0001, effective_cut_speed_m_min / 60.0)
        adjusted_cut_time_sec = adjusted_cut_length_m / max(0.0001, effective_cut_speed_m_min / 60.0)
        runtime_delta_sec = max(
            0.0,
            (nominal_cut_time_sec - adjusted_cut_time_sec)
            + (nominal_lead_time_sec - adjusted_lead_time_sec)
            + (nominal_pierce_time_sec - adjusted_pierce_time_sec)
            + (nominal_gas_time_sec - adjusted_gas_time_sec),
        ) * motion_overhead_factor
        adjusted_machine_total_min = max(0.0, float(totals.get("machine_total_min", 0) or 0) - (runtime_delta_sec / 60.0))
        adjusted_machine_runtime_eur = (adjusted_machine_total_min / 60.0) * machine_hour_eur
        adjusted_cut_eur_by_meter = adjusted_cut_length_m * cut_rate_per_m
        if cost_mode == "per_meter":
            adjusted_effective_cutting_eur = adjusted_cut_eur_by_meter
        elif cost_mode == "machine_time":
            adjusted_effective_cutting_eur = adjusted_machine_runtime_eur
        else:
            adjusted_effective_cutting_eur = max(adjusted_cut_eur_by_meter, adjusted_machine_runtime_eur)
        adjusted_pierce_cost_eur = adjusted_pierce_count * pierce_rate
        static_process_eur = max(
            0.0,
            float(totals.get("technical_process_eur", 0) or 0) - float(totals.get("effective_cutting_eur", 0) or 0) - float(totals.get("pierce_cost_eur", 0) or 0),
        )
        adjusted_process_eur = static_process_eur + adjusted_effective_cutting_eur + adjusted_pierce_cost_eur

        candidate_rows = [dict(row or {}) for row in list(self.result_data.get("sheet_candidates", []) or [])]
        if not candidate_rows:
            candidate_rows = [
                {
                    "name": str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or "Apenas stock").strip() or "Apenas stock",
                    "sheet_count": int(summary.get("sheet_count", 0) or 0),
                    "part_count_unplaced": int(summary.get("part_count_unplaced", 0) or 0),
                    "purchase_sheet_area_mm2": float(summary.get("purchase_sheet_area_mm2", 0) or 0),
                    "total_sheet_area_mm2": float(summary.get("total_sheet_area_mm2", 0) or 0),
                    "utilization_net_pct": float(summary.get("utilization_net_pct", 0) or 0),
                    "utilization_bbox_pct": float(summary.get("utilization_bbox_pct", 0) or 0),
                }
            ]
        ordered_candidates = sorted(candidate_rows, key=self._candidate_score)
        selected_name = str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or "Apenas stock").strip() or "Apenas stock"
        selected_candidate = next((row for row in ordered_candidates if str(row.get("name", "") or "").strip() == selected_name), ordered_candidates[0] if ordered_candidates else {})
        next_candidate = ordered_candidates[1] if len(ordered_candidates) > 1 else {}

        decision_lines = [
            f"Cenário escolhido: {selected_name} | {self._analysis_method_label(summary)} | estratégia {STRATEGY_LABELS.get(str(summary.get('strategy_name', '') or '').strip(), str(summary.get('strategy_name', '') or '-').strip() or '-')}.",
            "Critérios de decisão do motor: menos peças fora, menos compra adicional, menos área total, menos chapas e maior utilização.",
            f"Resultado final: {int(summary.get('part_count_placed', 0) or 0)} peça(s) colocada(s) em {int(summary.get('sheet_count', 0) or 0)} chapa(s), com {_fmt_num(summary.get('utilization_net_pct', 0), 2)}% de utilização real.",
            f"Stock usado: {int(summary.get('stock_sheet_count', 0) or 0)} chapa(s), incluindo {int(summary.get('remnant_sheet_count', 0) or 0)} retalho(s). Compra complementar: {int(summary.get('purchased_sheet_count', 0) or 0)} chapa(s).",
            f"Estimativa comercial das peças colocadas com {machine_name} / {commercial_name}: {_fmt_eur(totals['quoted_total_eur'])}.",
        ]
        if bool(common_line.get("enabled")):
            if common_line_saving_m > 0.0:
                decision_lines.append(
                    f"Common-line estimado: {_fmt_num(common_line_saving_m, 3)} m de corte evitado em {int(common_line.get('shared_edge_count', 0) or 0)} bordo(s) partilhado(s)."
                )
            else:
                decision_lines.append(
                    f"Common-line sem ganho direto neste plano. TolerÃ¢ncia usada: {_fmt_num(common_line.get('tolerance_mm', 0), 2)} mm."
                )
            if int(common_line.get("blocked_candidates", 0) or 0) > 0:
                decision_lines.append(
                    f"Foram encontradas {int(common_line.get('blocked_candidates', 0) or 0)} aproximaÃ§Ã£o(Ãµes) com potencial, mas a folga entre peÃ§as impediu partilha de corte."
                )
        if bool(self.lead_opt_check.isChecked()):
            decision_lines.append(
                f"OtimizaÃ§Ã£o de lead-ins aplicada com reduÃ§Ã£o de {_fmt_num(self.lead_opt_pct_spin.value(), 1)}% sobre o percurso remanescente de entrada/saÃ­da."
            )
        decision_lines.append(
            f"Processo ajustado do plano: corte {_fmt_num(nominal_cut_length_m, 3)} -> {_fmt_num(adjusted_cut_length_m, 3)} m, pierces {nominal_pierce_count} -> {adjusted_pierce_count}, tempo mÃ¡quina {_fmt_num(totals.get('machine_total_min', 0), 2)} -> {_fmt_num(adjusted_machine_total_min, 2)} min."
        )
        if next_candidate:
            purchase_delta_m2 = ((float(next_candidate.get("purchase_sheet_area_mm2", 0) or 0) - float(selected_candidate.get("purchase_sheet_area_mm2", 0) or 0)) / 1_000_000.0)
            total_delta_m2 = ((float(next_candidate.get("total_sheet_area_mm2", 0) or 0) - float(selected_candidate.get("total_sheet_area_mm2", 0) or 0)) / 1_000_000.0)
            sheets_delta = int(next_candidate.get("sheet_count", 0) or 0) - int(selected_candidate.get("sheet_count", 0) or 0)
            outside_delta = int(next_candidate.get("part_count_unplaced", 0) or 0) - int(selected_candidate.get("part_count_unplaced", 0) or 0)
            decision_lines.append(
                f"Face à alternativa seguinte ({str(next_candidate.get('name', '-') or '-').strip()}), o cenário escolhido varia {purchase_delta_m2:+.4f} m2 de compra, {total_delta_m2:+.4f} m2 de área total, {sheets_delta:+d} chapa(s) e {outside_delta:+d} peça(s) fora."
            )
        if errors:
            decision_lines.append("Não foi possível estimar o processo para: " + "; ".join(errors))

        total_placed_net_area_m2 = sum(max(0.0, float(row.get("placed_net_area_m2", 0) or 0.0)) for row in part_rows)
        total_placed_material_area_m2 = sum(
            max(0.0, float(row.get("placed_net_area_m2", 0) or 0.0))
            for row in part_rows
            if not bool(row.get("material_supplied_by_client", False))
        )
        allocated_material_total = max(0.0, float(summary.get("material_net_cost_eur", 0) or 0))
        allocated_purchase_total = max(0.0, float(summary.get("material_purchase_requirement_eur", 0) or 0))
        adjusted_total_sum = 0.0
        current_total_sum = 0.0
        for row in part_rows:
            qty = max(1, int(row.get("qty", 0) or 0))
            area_share = max(0.0, float(row.get("placed_net_area_m2", 0) or 0.0))
            material_supplied_by_client = bool(row.get("material_supplied_by_client", False))
            if material_supplied_by_client:
                share_ratio = 0.0
            elif total_placed_material_area_m2 > 0:
                share_ratio = area_share / total_placed_material_area_m2
            else:
                material_rows = [candidate for candidate in part_rows if not bool(candidate.get("material_supplied_by_client", False))]
                share_ratio = (1.0 / max(1, len(material_rows))) if material_rows else 0.0
            rateio_material_total = 0.0 if material_supplied_by_client else (allocated_material_total * share_ratio)
            rateio_purchase_total = 0.0 if material_supplied_by_client else (allocated_purchase_total * share_ratio)
            rateio_material_unit = rateio_material_total / qty
            rateio_purchase_unit = rateio_purchase_total / qty
            non_material_subtotal_unit = max(0.0, float(row.get("non_material_subtotal_unit_eur", 0) or 0.0))
            margin_pct = max(0.0, float(row.get("margin_pct", 0) or 0.0))
            adjusted_subtotal_unit = rateio_material_unit + non_material_subtotal_unit
            adjusted_total = (adjusted_subtotal_unit * (1.0 + (margin_pct / 100.0))) * qty
            adjusted_unit = adjusted_total / qty
            current_quote_total = float(row.get("current_quote_total_eur", 0) or 0.0)
            current_total_sum += current_quote_total
            adjusted_total_sum += adjusted_total
            row["share_pct"] = round(share_ratio * 100.0, 2)
            row["allocated_material_total_eur"] = round(rateio_material_total, 2)
            row["allocated_material_unit_eur"] = round(rateio_material_unit, 4)
            row["allocated_purchase_total_eur"] = round(rateio_purchase_total, 2)
            row["allocated_purchase_unit_eur"] = round(rateio_purchase_unit, 4)
            row["adjusted_quote_total_eur"] = round(adjusted_total, 2)
            row["adjusted_quote_unit_eur"] = round(adjusted_unit, 4)
            row["quote_delta_eur"] = round(current_quote_total - adjusted_total, 2)
            row["quote_delta_pct"] = round(((current_quote_total - adjusted_total) / adjusted_total * 100.0) if adjusted_total > 0 else 0.0, 2)

        if part_rows:
            if any(bool(row.get("material_supplied_by_client", False)) for row in part_rows):
                decision_lines.append(
                    "Algumas referencias estao marcadas como materia-prima fornecida pelo cliente; o rateio do plano ignora materia-prima nessas linhas e considera apenas transformacao/processo."
                )
            decision_lines.append(
                f"Rateio de matéria real por área líquida colocada: {_fmt_eur(allocated_material_total)} distribuídos por {_fmt_num(total_placed_net_area_m2, 4)} m2."
            )
            if allocated_purchase_total > 0:
                decision_lines.append(
                    f"Compra necessária rateada pelas referências colocadas: {_fmt_eur(allocated_purchase_total)}."
                )
            decision_lines.append(
                f"Comparação entre orçamento atual das linhas e cenário ajustado pelo plano: {_fmt_eur(current_total_sum)} vs {_fmt_eur(adjusted_total_sum)}."
            )

        return {
            "machine_name": machine_name,
            "commercial_name": commercial_name,
            "totals": totals,
            "part_rows": part_rows,
            "rateio": {
                "base": "area_liquida_colocada",
                "total_placed_net_area_m2": round(total_placed_net_area_m2, 6),
                "total_placed_material_area_m2": round(total_placed_material_area_m2, 6),
                "allocated_material_total_eur": round(allocated_material_total, 2),
                "allocated_purchase_total_eur": round(allocated_purchase_total, 2),
                "adjusted_quote_total_eur": round(adjusted_total_sum, 2),
                "current_quote_total_eur": round(current_total_sum, 2),
                "delta_eur": round(current_total_sum - adjusted_total_sum, 2),
            },
            "decision_lines": decision_lines,
            "errors": errors,
            "common_line": common_line,
            "nominal_cut_length_m": round(nominal_cut_length_m, 4),
            "adjusted_cut_length_m": round(adjusted_cut_length_m, 4),
            "nominal_pierce_count": int(nominal_pierce_count),
            "adjusted_pierce_count": int(adjusted_pierce_count),
            "adjusted_machine_total_min": round(adjusted_machine_total_min, 3),
            "adjusted_process_eur": round(adjusted_process_eur, 2),
        }

    def _populate_cost_report(self, summary: dict[str, Any]) -> None:
        report = self._build_cost_report(summary)
        totals = dict(report.get("totals", {}) or {})
        rateio = dict(report.get("rateio", {}) or {})
        common_line = dict(report.get("common_line", {}) or {})
        cost_rows = [
            ("Materia real do plano", _fmt_eur(summary.get("material_net_cost_eur", 0)), "Nesting", "Custo liquido da materia considerando area real de chapa e credito de sucata."),
            ("Compra necessaria", _fmt_eur(summary.get("material_purchase_requirement_eur", 0)), "Nesting", "Custo apenas da chapa que tera de ser comprada para completar o plano."),
            ("Valor comercial colocado", _fmt_eur(totals.get("quoted_total_eur", 0)), str(report.get("commercial_name", "-") or "-"), "Estimativa comercial das pecas colocadas com o perfil atual de orcamentacao."),
            ("Processo tecnico", _fmt_eur(totals.get("technical_process_eur", 0)), str(report.get("machine_name", "-") or "-"), "Estimativa nominal de corte, pierces, handling e tempos."),
            ("Processo ajustado", _fmt_eur(report.get("adjusted_process_eur", 0)), "CAM estimado", "Reflete a estimativa de common-line e otimizacao de lead-ins/lead-outs no plano atual."),
            ("Comercial atual linhas", _fmt_eur(rateio.get("current_quote_total_eur", 0)), "Orçamento", "Soma atual das linhas colocadas no grupo, usando o preço comercial atualmente gravado no orçamento."),
            ("Comercial ajustado plano", _fmt_eur(rateio.get("adjusted_quote_total_eur", 0)), "Rateio", "Valor comercial reestimado ao substituir o custo de matéria teórico pelo custo real rateado do plano de chapa."),
            ("Delta comercial", _fmt_eur(rateio.get("delta_eur", 0)), "Atual - ajustado", "Positivo = orçamento atual acima do plano rateado. Negativo = orçamento atual abaixo do cenário ajustado."),
            ("Tempo maquina", f"{_fmt_num(totals.get('machine_total_min', 0), 2)} min", str(report.get("machine_name", "-") or "-"), "Tempo total nominal estimado para as pecas colocadas."),
            ("Tempo maquina ajustado", f"{_fmt_num(report.get('adjusted_machine_total_min', 0), 2)} min", "CAM estimado", "Tempo apos ajuste de common-line e entradas/saidas."),
            ("Corte total", f"{_fmt_num(report.get('nominal_cut_length_m', totals.get('cut_length_m', 0)), 3)} m", "DXF", "Soma nominal dos comprimentos de corte das pecas colocadas."),
            ("Corte ajustado", f"{_fmt_num(report.get('adjusted_cut_length_m', 0), 3)} m", "Common-line", "Comprimento estimado apos partilha de bordo entre pecas."),
            ("Pierces totais", str(int(report.get("nominal_pierce_count", totals.get("pierce_count", 0)) or 0)), "DXF", "Numero nominal estimado de pierces das pecas colocadas."),
            ("Pierces ajustados", str(int(report.get("adjusted_pierce_count", 0) or 0)), "Common-line", "Pierces estimados apos consolidacao de bordos/entradas."),
            ("Bordos common-line", str(int(common_line.get("shared_edge_count", 0) or 0)), f"{_fmt_num(common_line.get('tolerance_mm', 0), 2)} mm", "Bordos partilhados detetados no plano atual."),
            ("Massa bruta de chapa", f"{_fmt_num(summary.get('gross_sheet_mass_kg', 0), 3)} kg", "Nesting", "Massa da chapa usada no plano atual."),
        ]
        self.cost_table.setRowCount(len(cost_rows))
        for row_index, values in enumerate(cost_rows):
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(int((Qt.AlignLeft if col_index in (0, 3) else Qt.AlignCenter) | Qt.AlignVCenter))
                self.cost_table.setItem(row_index, col_index, item)

        part_rows = [dict(row or {}) for row in list(report.get("part_rows", []) or [])]
        self.process_table.setRowCount(len(part_rows))
        for row_index, row in enumerate(part_rows):
            values = [
                str(row.get("ref_externa", "") or "-").strip() or "-",
                str(row.get("description", "") or "-").strip() or "-",
                str(int(row.get("qty", 0) or 0)),
                f"{_fmt_num(row.get('machine_total_min', 0), 2)}",
                f"{_fmt_num(row.get('cut_length_m', 0), 3)}",
                str(int(row.get("pierce_count", 0) or 0)),
                _fmt_eur(row.get("allocated_material_total_eur", 0)),
                _fmt_eur(row.get("allocated_purchase_total_eur", 0)),
                _fmt_eur(row.get("current_quote_total_eur", row.get("quoted_total_eur", 0))),
                _fmt_eur(row.get("adjusted_quote_total_eur", 0)),
            ]
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(int((Qt.AlignLeft if col_index in (0, 1) else Qt.AlignCenter) | Qt.AlignVCenter))
                self.process_table.setItem(row_index, col_index, item)

        self.decision_edit.setPlainText("\n".join(f"- {line}" for line in list(report.get("decision_lines", []) or [])))

    def _apply_result_data(self, result_data: dict[str, Any]) -> None:
        self.result_data = dict(result_data or {})
        summary = dict(self.result_data.get("summary", {}) or {})
        selected_profile = dict(summary.get("selected_sheet_profile", {}) or {})
        selected_name = str(selected_profile.get("name", "") or "").strip()
        strategy_name = STRATEGY_LABELS.get(str(summary.get("strategy_name", "") or "").strip(), str(summary.get("strategy_name", "") or "").strip() or "-")
        requested = int(summary.get("part_count_requested", 0) or 0)
        placed = int(summary.get("part_count_placed", 0) or 0)
        unplaced = int(summary.get("part_count_unplaced", 0) or 0)

        stock_used_txt = (
            f"{int(summary.get('stock_sheet_count', 0) or 0)} total | "
            f"{int(summary.get('remnant_sheet_count', 0) or 0)} retalho(s)"
        )
        buy_txt = (
            f"{int(summary.get('purchased_sheet_count', 0) or 0)} chapa(s) | "
            f"{_fmt_num((summary.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2"
        )

        self.summary_labels["profile"].setText(selected_name or "Apenas stock")
        self.summary_labels["method"].setText(self._analysis_method_label(summary))
        self.summary_labels["strategy"].setText(strategy_name)
        self.summary_labels["sheets"].setText(str(int(summary.get("sheet_count", 0) or 0)))
        self.summary_labels["stock_used"].setText(stock_used_txt)
        self.summary_labels["buy_count"].setText(str(int(summary.get("purchased_sheet_count", 0) or 0)))
        self.summary_labels["parts"].setText(f"{placed} / {requested}")
        self.summary_labels["unplaced"].setText(str(unplaced))
        self.summary_labels["util_net"].setText(f"{_fmt_num(summary.get('utilization_net_pct', 0), 2)} %")
        self.summary_labels["util_bbox"].setText(f"{_fmt_num(summary.get('utilization_bbox_pct', 0), 2)} %")
        self.summary_labels["compactness"].setText(f"{_fmt_num(summary.get('layout_compactness_pct', 0), 2)} %")
        self.summary_labels["purchase"].setText(buy_txt)
        self.summary_labels["area"].setText(f"{_fmt_num((summary.get('total_sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2")
        self.summary_labels["mass"].setText(f"{_fmt_num(summary.get('gross_sheet_mass_kg', 0), 3)} kg")
        self.summary_labels["material_cost"].setText(_fmt_eur(summary.get("material_net_cost_eur", 0)))

        candidates = [dict(row or {}) for row in list(self.result_data.get("sheet_candidates", []) or []) if isinstance(row, dict)]
        if candidates:
            best_compact = max(candidates, key=lambda row: float(row.get("layout_compactness_pct", 0) or 0))
            best_purchase = min(candidates, key=lambda row: float(row.get("purchase_sheet_area_mm2", 0) or 0))
            selected_name_txt = str(summary.get("selected_candidate_name", "") or candidates[0].get("name", "") or "-").strip() or "-"
            compact_txt = f"{_fmt_num(best_compact.get('layout_compactness_pct', 0), 1)}%"
            purchase_txt = f"{_fmt_num((best_purchase.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0, 4)} m2"
            self.recommendation_primary.setText(f"Selecionado: {selected_name_txt}")
            self.recommendation_secondary.setText(f"Melhor compactação: {compact_txt}")
            self.recommendation_delta.setText(f"Menor compra: {purchase_txt}")
            self.recommendation_note.setText(
                f"O motor comparou {len(candidates)} cenário(s). O cenário ativo favorece "
                f"{self._analysis_method_label(summary).lower()}, com {_fmt_num(summary.get('utilization_net_pct', 0), 1)}% de aproveitamento real."
            )
        else:
            self.recommendation_primary.setText("Sem cenários")
            self.recommendation_secondary.setText("Compactação: -")
            self.recommendation_delta.setText("Compra: -")
            self.recommendation_note.setText("Corre a otimização para receber uma recomendação automática do estudo.")

        warnings = list(self.result_data.get("warnings", []) or [])
        notes = self._default_notes(summary, warnings)
        self.warning_edit.setPlainText("\n".join(f"- {row}" for row in notes))
        self._refresh_parts_programming_status()
        self._populate_candidate_table(summary)
        self._populate_sheet_plan_table(selected_name or "Apenas stock")
        self._populate_sheet_gallery(selected_name or "Apenas stock")
        self._populate_unplaced_table()
        try:
            self._populate_cost_report(summary)
        except Exception as exc:
            self.cost_table.setRowCount(1)
            error_row = [
                "Erro no relatório de custo",
                str(exc),
                "Diagnóstico",
                "O backend do nesting calculou o plano, mas a apresentação do relatório falhou. Revê a fórmula ou o payload do relatório.",
            ]
            for col_index, value in enumerate(error_row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(int((Qt.AlignLeft if col_index in (0, 1, 3) else Qt.AlignCenter) | Qt.AlignVCenter))
                self.cost_table.setItem(0, col_index, item)
            self.process_table.setRowCount(0)
            self.decision_edit.setPlainText(f"- Falha ao montar o relatório de custo: {exc}")

        if selected_name and selected_profile.get("width_mm"):
            self._refresh_sheet_profiles(selected_name)
        self.sheet_view_combo.blockSignals(True)
        self.sheet_view_combo.clear()
        for sheet_row in list(self.result_data.get("sheets", []) or []):
            self.sheet_view_combo.addItem(
                f"Chapa {sheet_row.get('index', 0)} | {self._sheet_combo_label(sheet_row, selected_name)} | "
                f"{int(sheet_row.get('part_count', 0) or 0)} pecas | "
                f"real {_fmt_num(sheet_row.get('utilization_net_pct', 0), 1)}% | "
                f"nesting {_fmt_num(sheet_row.get('utilization_bbox_pct', 0), 1)}%"
            )
        self.sheet_view_combo.blockSignals(False)
        if self.sheet_view_combo.count() > 0:
            self.sheet_view_combo.setCurrentIndex(0)
        if self.sheet_plan_table.rowCount() > 0:
            self.sheet_plan_table.selectRow(0)
        if self.sheet_gallery.count() > 0:
            self.sheet_gallery.setCurrentRow(0)
        self._render_sheet_preview()
        self.nest_tabs.setCurrentIndex(0)
        self.results_tabs.setCurrentIndex(0)
        self._set_wizard_step(2 if self.sheet_plan_table.rowCount() > 0 else 4)

    def _analyze(self) -> None:
        if self._nesting_thread is not None:
            return
        group = self._current_group()
        rows = [dict(row or {}) for row in list(group.get("rows", []) or [])]
        for row in rows:
            row["nest_rotation_policy"] = self._part_rotation_override(row)
            row["nest_priority"] = self._part_priority_override(row)
        if not rows:
            QMessageBox.warning(self, "Nesting Laser", "Nao existem linhas laser validas para este grupo.")
            return

        self._persist_nesting_preferences()
        auto_mode = bool(self.auto_sheet_check.isChecked())
        use_stock = bool(self.use_stock_check.isChecked())
        allow_purchase = bool(self.allow_purchase_check.isChecked())
        shape_aware = bool(self.shape_check.isChecked())
        needs_purchase_profile = (not use_stock) or allow_purchase
        if use_stock and not self.stock_candidates and not allow_purchase:
            QMessageBox.warning(self, "Nesting Laser", "Nao existe stock ou retalho disponivel para este grupo.")
            return
        if needs_purchase_profile and not self.sheet_profiles:
            QMessageBox.warning(self, "Nesting Laser", "Define pelo menos um formato de chapa para compra ou complemento.")
            return
        sheet = self._selected_sheet_profile()
        if needs_purchase_profile and not auto_mode and not sheet:
            QMessageBox.warning(self, "Nesting Laser", "Seleciona um formato de chapa valido.")
            return
        payload = {
            "rows": rows,
            "sheet_width_mm": float(sheet.get("width_mm", 0) or 0) if not auto_mode else None,
            "sheet_height_mm": float(sheet.get("height_mm", 0) or 0) if not auto_mode else None,
            "part_spacing_mm": self.spacing_spin.value(),
            "edge_margin_mm": self.edge_spin.value(),
            "allow_rotate": bool(self.rotate_check.isChecked()),
            "allow_mirror": bool(self.mirror_check.isChecked()),
            "free_angle_rotation": bool(self.free_angle_check.isChecked()),
            "laser_settings": self.settings,
            "sheet_name": str(sheet.get("name", "") or ""),
            "sheet_profiles": self.sheet_profiles,
            "auto_select_sheet": auto_mode,
            "stock_sheet_candidates": self.stock_candidates,
            "use_stock_first": use_stock,
            "allow_purchase_fallback": allow_purchase,
            "shape_aware": shape_aware,
            "strict_shape_only": bool(self.shape_strict_check.isChecked()),
            "shape_grid_mm": float(self.shape_grid_spin.value()),
        }
        self._set_nesting_busy(True)
        thread = QThread(self)
        worker = NestingWorker(payload)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._handle_nesting_success)
        worker.failed.connect(self._handle_nesting_error)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._cleanup_nesting_thread)
        self._nesting_thread = thread
        self._nesting_worker = worker
        thread.start()


def _laser_nesting_render_sheet_preview_v2(self: LaserNestingDialog) -> None:
    self.preview_scene.clear()
    if not self.result_data:
        self.preview_hint.setText("Ainda nao existe um plano calculado para pre-visualizar.")
        return

    sheets = list(self.result_data.get("sheets", []) or [])
    if not sheets:
        self.preview_hint.setText("O resultado atual nao tem chapas colocadas.")
        return

    sheet_index = max(0, self.sheet_view_combo.currentIndex())
    if sheet_index >= len(sheets):
        sheet_index = 0

    sheet = dict(sheets[sheet_index] or {})
    summary = dict(self.result_data.get("summary", {}) or {})
    display = self._sheet_display_context(sheet)
    sheet_width = float(display.get("display_width_mm", 0) or 0)
    sheet_height = float(display.get("display_height_mm", 0) or 0)
    if sheet_width <= 0 or sheet_height <= 0:
        fallback_display = self._sheet_display_context(summary)
        sheet_width = float(fallback_display.get("display_width_mm", 0) or 0)
        sheet_height = float(fallback_display.get("display_height_mm", 0) or 0)
    if sheet_width <= 0 or sheet_height <= 0:
        self.preview_hint.setText("Nao foi possivel determinar as dimensoes da chapa deste cenario.")
        return

    source_label = self._sheet_combo_label(
        sheet,
        str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or ""),
    )
    geometry_validation = dict(sheet.get("geometry_validation", {}) or {})
    if not geometry_validation and list(sheet.get("placements", []) or []):
        geometry_validation = _sheet_overlap_diagnostics({"placements": list(sheet.get("placements", []) or [])})
    part_in_part_pairs = int(geometry_validation.get("part_in_part_pair_count", 0) or 0)
    solid_overlap_pairs = int(geometry_validation.get("solid_overlap_pair_count", 0) or 0)

    hint_parts = [
        source_label,
        self._sheet_dimension_label(sheet_width, sheet_height),
        f"{int(sheet.get('part_count', 0) or 0)} peca(s)",
        f"real {_fmt_num(sheet.get('utilization_net_pct', 0), 1)}%",
        f"nesting {_fmt_num(sheet.get('utilization_bbox_pct', 0), 1)}%",
        f"compactacao {_fmt_num(sheet.get('layout_compactness_pct', summary.get('layout_compactness_pct', 0)), 1)}%",
    ]
    if solid_overlap_pairs > 0:
        hint_parts.append(f"ALERTA: {solid_overlap_pairs} colisao(oes) reais")
    elif part_in_part_pairs > 0:
        hint_parts.append(f"{part_in_part_pairs} encaixe(s) interno(s) valido(s)")
    self.preview_hint.setText(" | ".join(hint_parts))

    edge_margin_mm = max(0.0, float(self.edge_spin.value()))
    preview_pad = max(10.0, min(18.0, min(sheet_width, sheet_height) * 0.018))
    bottom_preview_pad = max(preview_pad * 2.2, 34.0)
    base_scene_rect = QRectF(
        -preview_pad,
        -preview_pad,
        sheet_width + (preview_pad * 2.0),
        sheet_height + preview_pad + bottom_preview_pad,
    )
    self.preview_scene.setSceneRect(base_scene_rect)
    sheet_outer_polygons = list(display.get("display_outer_polygons", []) or [])
    sheet_hole_polygons = list(display.get("display_hole_polygons", []) or [])
    if sheet_outer_polygons:
        for polygon_points in sheet_outer_polygons:
            polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
            self.preview_scene.addPolygon(polygon, QPen(QColor("#17314f"), 2), QBrush(QColor("#f8fafc")))
        for polygon_points in sheet_hole_polygons:
            polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
            self.preview_scene.addPolygon(polygon, QPen(QColor("#17314f"), 2), QBrush(QColor("#ffffff")))
    else:
        self.preview_scene.addRect(
            0,
            0,
            sheet_width,
            sheet_height,
            QPen(QColor("#17314f"), 2),
            QBrush(QColor("#f8fafc")),
        )

    if edge_margin_mm > 0.0 and (sheet_width - (2.0 * edge_margin_mm)) > 0 and (sheet_height - (2.0 * edge_margin_mm)) > 0:
        usable_pen = QPen(QColor("#7f97ae"), 1.1, Qt.DashLine)
        usable_pen.setDashPattern([6, 4])
        usable_fill = QBrush(QColor(125, 151, 174, 18))
        self.preview_scene.addRect(
            edge_margin_mm,
            edge_margin_mm,
            sheet_width - (2.0 * edge_margin_mm),
            sheet_height - (2.0 * edge_margin_mm),
            usable_pen,
            usable_fill,
        )
        margin_note = self.preview_scene.addText(f"Margem {edge_margin_mm:.1f} mm")
        margin_note.setDefaultTextColor(QColor("#486581"))
        margin_note.setPos(edge_margin_mm + 6, max(2.0, edge_margin_mm - 18.0))

    for index, placement in enumerate(list(display.get("placements", []) or []), start=1):
        color = self._sheet_preview_color(index)
        tooltip = (
            f"{placement.get('ref_externa', '-')}\n"
            f"{placement.get('display_width_mm', 0)} x {placement.get('display_height_mm', 0)} mm\n"
            f"Rotação: {_fmt_num(placement.get('rotation_deg', 90 if placement.get('rotated') else 0), 1)}°"
            f"{' | espelhada' if placement.get('mirrored') else ''}"
        )
        outer_polygons = list(placement.get("display_outer_polygons", []) or [])
        hole_polygons = list(placement.get("display_hole_polygons", []) or [])
        preview_paths = list(placement.get("display_preview_paths", []) or [])
        if outer_polygons:
            for polygon_points in outer_polygons:
                polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                polygon_item = self.preview_scene.addPolygon(polygon, QPen(QColor("#274c77"), 1), QBrush(color))
                polygon_item.setToolTip(tooltip)
            for polygon_points in hole_polygons:
                polygon = QPolygonF([QPointF(float(x), float(y)) for x, y in list(polygon_points or [])])
                self.preview_scene.addPolygon(polygon, QPen(QColor("#274c77"), 1), QBrush(QColor("#ffffff")))
        else:
            rect = self.preview_scene.addRect(
                float(placement.get("display_x_mm", 0) or 0),
                float(placement.get("display_y_mm", 0) or 0),
                float(placement.get("display_width_mm", 0) or 0),
                float(placement.get("display_height_mm", 0) or 0),
                QPen(QColor("#274c77"), 1),
                QBrush(color),
            )
            rect.setToolTip(tooltip)
        if preview_paths:
            path_pen = QPen(QColor("#475569"), 0.9)
            for path_points in preview_paths:
                if len(list(path_points or [])) < 2:
                    continue
                painter_path = QPainterPath(QPointF(float(path_points[0][0]), float(path_points[0][1])))
                for point_x, point_y in list(path_points or [])[1:]:
                    painter_path.lineTo(float(point_x), float(point_y))
                path_item = self.preview_scene.addPath(painter_path, path_pen)
                path_item.setToolTip(tooltip)

        badge_x = float(placement.get("display_x_mm", 0) or 0) + 4
        badge_y = float(placement.get("display_y_mm", 0) or 0) + 4
        badge_item = self.preview_scene.addEllipse(
            badge_x,
            badge_y,
            16,
            16,
            QPen(QColor("#0f172a"), 0.8),
            QBrush(QColor("#ffffff")),
        )
        badge_item.setToolTip(tooltip)
        badge_text = self.preview_scene.addText(str(index))
        badge_text.setDefaultTextColor(QColor("#0f172a"))
        badge_text.setPos(badge_x + 4, badge_y - 1)
        badge_text.setToolTip(tooltip)

        if float(placement.get("display_width_mm", 0) or 0) > 150 and float(placement.get("display_height_mm", 0) or 0) > 34:
            ref_text = str(placement.get("ref_externa", "") or "").strip()
            if ref_text:
                ref_item = self.preview_scene.addText(ref_text[:22])
                ref_item.setDefaultTextColor(QColor("#0f172a"))
                ref_item.setPos(float(placement.get("display_x_mm", 0) or 0) + 24, float(placement.get("display_y_mm", 0) or 0) + 2)
                ref_item.setToolTip(tooltip)

    self.preview_scene.setSceneRect(base_scene_rect)
    self.preview_view.request_fit()
    if self.sheet_plan_table.rowCount() > sheet_index:
        self.sheet_plan_table.blockSignals(True)
        self.sheet_plan_table.clearSelection()
        self.sheet_plan_table.selectRow(sheet_index)
        self.sheet_plan_table.blockSignals(False)
    if self.sheet_gallery.count() > sheet_index:
        self.sheet_gallery.blockSignals(True)
        self.sheet_gallery.setCurrentRow(sheet_index)
        self.sheet_gallery.blockSignals(False)


LaserNestingDialog._render_sheet_preview = _laser_nesting_render_sheet_preview_v2


def _laser_nesting_sync_wizard_controls_v2(self: LaserNestingDialog) -> None:
    step_index = self._current_wizard_step_index()
    section_label, page_label = WIZARD_STEPS[step_index][2], WIZARD_STEPS[step_index][3]
    self.wizard_path_label.setText("")
    current_section = int(self.section_stack.currentIndex())
    self.page_title_label.setText(page_label)
    self.page_subtitle_label.setText("")
    enabled_sections = {0}
    if self.result_data or current_section >= 1:
        enabled_sections.add(1)
    if self.result_data or current_section >= 2:
        enabled_sections.add(2)
    for section_index, button in enumerate(self.major_step_buttons):
        active = section_index == current_section
        button.setChecked(active)
        button.setEnabled(section_index in enabled_sections)
        if active:
            button.setStyleSheet(
                "QPushButton {"
                "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4b6078, stop:1 #24384f);"
                "color: #ffffff; border: 1px solid #1c2d41; border-bottom: 3px solid #152334;"
                "border-radius: 22px; padding: 10px 22px; font-size: 15px; font-weight: 900;"
                "}"
            )
        elif section_index in enabled_sections:
            button.setStyleSheet(
                "QPushButton {"
                "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #e8eef5);"
                "color: #24384f; border: 1px solid #b8c7d8; border-bottom: 3px solid #9fb0c3;"
                "border-radius: 22px; padding: 10px 22px; font-size: 15px; font-weight: 800;"
                "}"
                "QPushButton:hover {background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #dde7f0);}"
            )
        else:
            button.setStyleSheet(
                "QPushButton {"
                "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f4f6f8, stop:1 #e6ebf1);"
                "color: #8b98a7; border: 1px solid #d2d9e1; border-bottom: 3px solid #c4ccd6;"
                "border-radius: 22px; padding: 10px 22px; font-size: 15px; font-weight: 800;"
                "}"
            )
    self.prev_section_btn.setEnabled(step_index > 0)
    self.next_section_btn.setEnabled(step_index < (len(WIZARD_STEPS) - 1))
    next_labels = {
        0: "Seguinte: Pecas e stock",
        1: "Seguinte: Gerar nesting",
        2: "Seguinte: Layouts e chapas",
        3: "Seguinte: Pendentes e avisos",
        4: "Seguinte: Custos e decisao",
    }
    self.next_section_btn.setText(next_labels.get(step_index, "Concluir"))
    self.analyze_btn.setText("Atualizar nesting" if step_index >= 2 else "Iniciar otimizacao")
    if self.section_stack.currentIndex() == 1 and self.nest_tabs.currentIndex() == 1 and self.result_data:
        QTimer.singleShot(0, self._render_sheet_preview)


LaserNestingDialog._sync_wizard_controls = _laser_nesting_sync_wizard_controls_v2
