from __future__ import annotations

from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard
from .runtime_common import (
    apply_progress_style as _apply_progress_style,
    cap_width as _cap_width,
    configure_table as _configure_table,
    fill_table as _fill_table,
    repolish as _repolish,
    set_panel_tone as _set_panel_tone,
)


class PulsePage(QWidget):
    page_title = "Pulse"
    page_subtitle = "OEE, desvios, paragens e peças em curso com o mesmo backend do mobile."
    allow_auto_timer_refresh = True

    def __init__(self, runtime_service, parent=None) -> None:
        super().__init__(parent)
        self.runtime_service = runtime_service
        self.last_pulse_data: dict = {}
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        filters = CardFrame()
        filters.set_tone("info")
        filters_layout = QHBoxLayout(filters)
        filters_layout.setContentsMargins(14, 10, 14, 10)
        filters_layout.setSpacing(8)
        self.period_combo = QComboBox()
        self.period_combo.addItems(["Hoje", "7 dias", "30 dias", "Tudo"])
        self.period_combo.setCurrentText("7 dias")
        self.year_combo = QComboBox()
        current_year = str(datetime.now().year)
        self.year_combo.addItems([current_year, "Todos"])
        self.year_combo.setCurrentText(current_year)
        self.origin_combo = QComboBox()
        self.origin_combo.addItems(["Ambos", "Em curso", "Histórico"])
        self.view_combo = QComboBox()
        self.view_combo.addItems(["Todas", "So desvio"])
        self.graphs_btn = QPushButton("Graficos")
        self.graphs_btn.setProperty("variant", "secondary")
        self.graphs_btn.clicked.connect(self._show_graphs_dialog)
        self.plan_delay_btn = QPushButton("Atrasos planeamento")
        self.plan_delay_btn.setProperty("variant", "secondary")
        self.plan_delay_btn.clicked.connect(self._show_plan_delay_dialog)
        for widget in (self.period_combo, self.year_combo, self.origin_combo, self.view_combo):
            widget.currentTextChanged.connect(self.refresh)
        for widget, width in ((self.period_combo, 140), (self.year_combo, 110), (self.origin_combo, 130), (self.view_combo, 130)):
            _cap_width(widget, width)
        filters_layout.addWidget(QLabel("Periodo"))
        filters_layout.addWidget(self.period_combo)
        filters_layout.addWidget(QLabel("Ano"))
        filters_layout.addWidget(self.year_combo)
        filters_layout.addWidget(QLabel("Origem"))
        filters_layout.addWidget(self.origin_combo)
        filters_layout.addWidget(QLabel("Visao"))
        filters_layout.addWidget(self.view_combo)
        filters_layout.addStretch(1)
        filters_layout.addWidget(self.plan_delay_btn)
        filters_layout.addWidget(self.graphs_btn)
        root.addWidget(filters)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(10)
        cards_layout.setVerticalSpacing(10)
        self.cards = [StatCard(title) for title in ("OEE", "Disponibilidade", "Performance", "Paragens", "Desvio max.")]
        for index, card in enumerate(self.cards):
            card.setMaximumHeight(112)
            cards_layout.addWidget(card, 0, index)
        self.cards[0].set_tone("info")
        self.cards[1].set_tone("success")
        self.cards[2].set_tone("warning")
        self.cards[3].set_tone("danger")
        self.cards[4].set_tone("warning")
        root.addWidget(cards_host)

        self.alert_card = CardFrame()
        alert_layout = QVBoxLayout(self.alert_card)
        alert_layout.setContentsMargins(14, 12, 14, 12)
        alert_layout.setSpacing(8)
        alert_title = QLabel("Alertas")
        alert_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        self.alert_label = QLabel("-")
        self.alert_label.setWordWrap(True)
        alert_layout.addWidget(alert_title)
        alert_layout.addWidget(self.alert_label)
        self.alert_card.setMaximumHeight(120)
        self.alert_card.set_tone("default")

        self.running_table = QTableWidget(0, 6)
        self.running_table.setHorizontalHeaderLabels(["Encomenda", "Peca", "Operacao", "Operador", "Tempo", "Desvio"])
        self.running_table.verticalHeader().setVisible(False)
        self.running_table.verticalHeader().setDefaultSectionSize(24)
        self.running_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.running_table, stretch=(1, 2), contents=(0, 3, 4, 5))

        self.stops_table = QTableWidget(0, 5)
        self.stops_table.setHorizontalHeaderLabels(["Causa", "Encomenda", "Operador", "Ocorrencias", "Minutos"])
        self.stops_table.verticalHeader().setVisible(False)
        self.stops_table.verticalHeader().setDefaultSectionSize(24)
        self.stops_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.stops_table, stretch=(0, 1), contents=(2, 3, 4))

        self.history_table = QTableWidget(0, 5)
        self.history_table.setHorizontalHeaderLabels(["Encomenda", "Ops", "Tempo", "Planeado", "Desvio"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.verticalHeader().setDefaultSectionSize(24)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        _configure_table(self.history_table, stretch=(0,), contents=(1, 2, 3, 4))

        running_card = CardFrame()
        running_card.set_tone("success")
        running_layout = QVBoxLayout(running_card)
        running_layout.setContentsMargins(14, 12, 14, 12)
        running_layout.setSpacing(8)
        running_title = QLabel("Pecas em curso")
        running_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        running_layout.addWidget(running_title)
        running_layout.addWidget(self.running_table)

        stops_card = CardFrame()
        stops_card.set_tone("danger")
        stops_layout = QVBoxLayout(stops_card)
        stops_layout.setContentsMargins(14, 12, 14, 12)
        stops_layout.setSpacing(8)
        stops_title = QLabel("Top causas de paragem")
        stops_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        stops_layout.addWidget(stops_title)
        stops_layout.addWidget(self.stops_table)

        history_card = CardFrame()
        history_card.set_tone("info")
        history_layout = QVBoxLayout(history_card)
        history_layout.setContentsMargins(14, 12, 14, 12)
        history_layout.setSpacing(8)
        history_title = QLabel("Histórico consolidado")
        history_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a;")
        history_layout.addWidget(history_title)
        history_layout.addWidget(self.history_table)

        right_stack = QWidget()
        right_stack_layout = QVBoxLayout(right_stack)
        right_stack_layout.setContentsMargins(0, 0, 0, 0)
        right_stack_layout.setSpacing(10)
        right_stack_layout.addWidget(self.alert_card)
        right_stack_layout.addWidget(stops_card, 1)

        top_split = QSplitter(Qt.Horizontal)
        top_split.setChildrenCollapsible(False)
        top_split.addWidget(running_card)
        top_split.addWidget(right_stack)
        top_split.setSizes([1180, 720])
        root.addWidget(top_split, 1)
        root.addWidget(history_card, 1)

    def refresh(self) -> None:
        year = self.year_combo.currentText()
        data = self.runtime_service.dashboard(
            period=self.period_combo.currentText(),
            year=None if year == "Todos" else year,
            visao=self.view_combo.currentText(),
            origem=self.origin_combo.currentText(),
        )
        self.last_pulse_data = dict(data or {})
        summary = data.get("summary", {})
        self.cards[0].set_data(f"{summary.get('oee', 0):.1f}%", f"Atualizado {data.get('updated_at', '-')}")
        self.cards[1].set_data(f"{summary.get('disponibilidade', 0):.1f}%", f"Qualidade {summary.get('qualidade', 0):.1f}%")
        perf_plan = float(summary.get("perf_plan_total", 0) or 0)
        perf_real = float(summary.get("perf_real_total", 0) or 0)
        self.cards[2].set_data(f"{summary.get('performance', 0):.1f}%", f"Plano {perf_plan:.1f} | Real {perf_real:.1f}")
        self.cards[3].set_data(f"{summary.get('paragens_min', 0):.1f} min", f"Fora do tempo {summary.get('pecas_fora_tempo', 0)}")
        self.cards[4].set_data(f"{summary.get('desvio_max_min', 0):.1f} min", f"Em curso {summary.get('pecas_em_curso', 0)}")
        perf_value = float(summary.get("performance", 0) or 0)
        self.cards[2].set_tone("success" if perf_value >= 100.0 else ("warning" if perf_value >= 80.0 else "danger"))
        self.alert_label.setText(str(summary.get("alerts", "-")))
        _set_panel_tone(self.alert_card, "danger" if str(summary.get("alerts", "") or "-").strip() not in {"", "-"} else "default")
        plan_delay = dict(data.get("plan_delay", {}) or {})
        plan_delay_open = int(plan_delay.get("open_count", 0) or 0)
        plan_delay_ack = int(plan_delay.get("acknowledged_count", 0) or 0)
        self.plan_delay_btn.setText(f"Atrasos planeamento ({plan_delay_open})" if plan_delay_open > 0 else "Atrasos planeamento")
        self.plan_delay_btn.setToolTip(
            f"{plan_delay_open} pendente(s) | {plan_delay_ack} justificado(s)"
            if (plan_delay_open or plan_delay_ack)
            else "Sem grupos fora do horário planeado neste momento."
        )
        self.plan_delay_btn.setProperty(
            "variant",
            "danger" if plan_delay_open > 0 else ("secondary" if plan_delay_ack <= 0 else "warning"),
        )
        _repolish(self.plan_delay_btn)
        _fill_table(
            self.running_table,
            [[r.get("encomenda", "-"), r.get("peca", "-"), r.get("operacao", "-"), r.get("operador", "-"), f"{r.get('elapsed_min', 0):.1f} min", f"{r.get('delta_min', 0):.1f}"] for r in data.get("running", [])],
            align_center_from=4,
        )
        _fill_table(
            self.stops_table,
            [[r.get("causa", "-"), r.get("encomenda", "-"), r.get("operador", "-"), r.get("ocorrencias", 0), f"{r.get('minutos', 0):.1f}"] for r in data.get("top_stops", [])],
            align_center_from=3,
        )
        _fill_table(
            self.history_table,
            [[r.get("encomenda", "-"), r.get("ops", 0), f"{r.get('elapsed_min', 0):.1f}", f"{r.get('plan_min', 0):.1f}", f"{r.get('delta_min', 0):.1f}"] for r in data.get("history", [])],
            align_center_from=1,
        )

    def _prompt_plan_delay_reason(self, current_reason: str = "") -> str:
        options = [
            "Mudança de prioridade / urgência",
            "Matéria-prima ainda não disponível",
            "Aguardar posto / máquina ocupada",
            "Avaria / paragem no posto",
            "Decisão da chefia / replaneado manualmente",
        ]
        default_idx = 0
        current_txt = str(current_reason or "").strip()
        if current_txt:
            for idx, option in enumerate(options):
                if option.lower() == current_txt.lower():
                    default_idx = idx
                    break
        reason_txt, ok = QInputDialog.getItem(
            self,
            "Motivo do atraso ao planeamento",
            "Indica o motivo para retirar este grupo do aviso ativo:",
            options,
            default_idx,
            True,
        )
        if not ok:
            return ""
        return str(reason_txt or "").strip()

    def _show_plan_delay_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Atrasos Face Ao Planeamento")
        dialog.setMinimumWidth(980)
        dialog.setMinimumHeight(520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Grupos de corte laser fora do horário planeado")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        info_label = QLabel("")
        info_label.setWordWrap(True)
        info_label.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(info_label)

        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels(["Encomenda", "Cliente", "Material", "Esp.", "Planeado", "Posto", "Estado", "Motivo"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(26)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        _configure_table(table, stretch=(1, 4, 7), contents=(0, 2, 3, 5, 6))
        layout.addWidget(table, 1)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        justify_btn = QPushButton("Sinalizar motivo")
        justify_btn.setProperty("variant", "warning")
        reactivate_btn = QPushButton("Reativar aviso")
        reactivate_btn.setProperty("variant", "secondary")
        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("variant", "secondary")
        close_btn = QPushButton("Fechar")
        close_btn.setProperty("variant", "secondary")
        actions.addWidget(justify_btn)
        actions.addWidget(reactivate_btn)
        actions.addStretch(1)
        actions.addWidget(refresh_btn)
        actions.addWidget(close_btn)
        layout.addLayout(actions)

        row_map: dict[int, dict] = {}

        def _reload_dashboard() -> None:
            self.refresh()
            plan_delay = dict((self.last_pulse_data or {}).get("plan_delay", {}) or {})
            items = list(plan_delay.get("items", []) or [])
            info_label.setText(
                f"{int(plan_delay.get('open_count', 0) or 0)} pendente(s) | "
                f"{int(plan_delay.get('acknowledged_count', 0) or 0)} justificado(s). "
                "Este aviso é só face ao planeamento atual; não significa atraso final ao cliente."
            )
            row_map.clear()
            table.setRowCount(len(items))
            for row_idx, item in enumerate(items):
                values = [
                    str(item.get("numero", "") or "-"),
                    str(item.get("cliente", "") or "-"),
                    str(item.get("material", "") or "-"),
                    str(item.get("espessura", "") or "-"),
                    str(item.get("planned_end_txt", "") or item.get("planned_start_txt", "") or "-"),
                    str(item.get("posto", "") or "-"),
                    str(item.get("status_label", "") or "-"),
                    str(item.get("reason", "") or "-"),
                ]
                for col_idx, value in enumerate(values):
                    table_item = QTableWidgetItem(value)
                    if col_idx in {0, 2, 3, 4, 5, 6}:
                        table_item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                    table.setItem(row_idx, col_idx, table_item)
                status_open = not bool(item.get("acknowledged"))
                bg = QColor("#fff4e5" if status_open else "#ecfdf3")
                fg = QColor("#9a3412" if status_open else "#166534")
                for col_idx in range(table.columnCount()):
                    cell = table.item(row_idx, col_idx)
                    if cell is not None:
                        cell.setBackground(QBrush(bg))
                        if col_idx == 6:
                            cell.setForeground(QBrush(fg))
                row_map[row_idx] = dict(item)
            if items:
                table.selectRow(0)

        def _selected_item() -> dict | None:
            current_row = table.currentRow()
            return dict(row_map.get(current_row, {}) or {}) if current_row >= 0 else None

        def _justify_selected() -> None:
            item = _selected_item()
            if not item:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", "Seleciona uma linha primeiro.")
                return
            reason_txt = self._prompt_plan_delay_reason(str(item.get("reason", "") or "").strip())
            if not reason_txt:
                return
            try:
                self.runtime_service.pulse_plan_delay_set_reason(str(item.get("item_key", "") or "").strip(), reason_txt)
            except Exception as exc:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", str(exc))
                return
            _reload_dashboard()

        def _reactivate_selected() -> None:
            item = _selected_item()
            if not item:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", "Seleciona uma linha primeiro.")
                return
            try:
                self.runtime_service.pulse_plan_delay_clear_reason(str(item.get("item_key", "") or "").strip())
            except Exception as exc:
                QMessageBox.warning(dialog, "Atrasos Face Ao Planeamento", str(exc))
                return
            _reload_dashboard()

        justify_btn.clicked.connect(_justify_selected)
        reactivate_btn.clicked.connect(_reactivate_selected)
        refresh_btn.clicked.connect(_reload_dashboard)
        close_btn.clicked.connect(dialog.accept)
        _reload_dashboard()
        dialog.exec()

    def _show_graphs_dialog(self) -> None:
        summary = dict((self.last_pulse_data or {}).get("summary", {}) or {})
        dialog = QDialog(self)
        dialog.setWindowTitle("Gráficos Pulse")
        dialog.setMinimumWidth(720)
        dialog.setMinimumHeight(480)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        title = QLabel("Performance e eficiencia operacional")
        title.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        subtitle = QLabel(
            f"Periodo {self.period_combo.currentText()} | Ano {self.year_combo.currentText()} | Origem {self.origin_combo.currentText()} | "
            f"Plano {float(summary.get('perf_plan_total', 0) or 0):.1f} min | Real {float(summary.get('perf_real_total', 0) or 0):.1f} min"
        )
        subtitle.setProperty("role", "muted")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        metrics_card = CardFrame()
        metrics_card.set_tone("info")
        metrics_layout = QVBoxLayout(metrics_card)
        metrics_layout.setContentsMargins(12, 12, 12, 12)
        metrics_layout.setSpacing(10)
        no_data_scope = (
            not list((self.last_pulse_data or {}).get("running", []) or [])
            and not list((self.last_pulse_data or {}).get("history", []) or [])
            and float(summary.get("perf_plan_total", 0) or 0) <= 0
            and float(summary.get("perf_real_total", 0) or 0) <= 0
        )
        for label_text, value, tone in (
            ("OEE", 0.0 if no_data_scope else float(summary.get("oee", 0) or 0), "info"),
            ("Disponibilidade", 0.0 if no_data_scope else float(summary.get("disponibilidade", 0) or 0), "success"),
            ("Performance", 0.0 if no_data_scope else float(summary.get("performance", 0) or 0), "warning"),
            ("Qualidade", 0.0 if no_data_scope else float(summary.get("qualidade", 0) or 0), "success"),
        ):
            row_host = QWidget()
            row_layout = QVBoxLayout(row_host)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(4)
            row_head = QHBoxLayout()
            row_head.setContentsMargins(0, 0, 0, 0)
            name = QLabel(label_text)
            name.setStyleSheet("font-size: 13px; font-weight: 700; color: #0f172a;")
            value_lbl = QLabel(f"{value:.1f}%")
            value_lbl.setStyleSheet("font-size: 13px; font-weight: 800; color: #0f172a;")
            row_head.addWidget(name)
            row_head.addStretch(1)
            row_head.addWidget(value_lbl)
            bar = QProgressBar()
            max_range = 250 if label_text == "Performance" else 100
            bar.setRange(0, max_range)
            bar.setValue(max(0, min(max_range, int(round(value)))))
            bar.setTextVisible(False)
            _apply_progress_style(bar, compact=True)
            if tone == "warning":
                bar.setStyleSheet(
                    "QProgressBar {background: #fff7ed; border: 1px solid #f3d6aa; border-radius: 8px; min-height: 14px; text-align: center; color: #9a3412; font-weight: 700;}"
                    "QProgressBar::chunk {background: #ea580c; border-radius: 7px;}"
                )
            row_layout.addLayout(row_head)
            row_layout.addWidget(bar)
            metrics_layout.addWidget(row_host)
        layout.addWidget(metrics_card)

        andon = dict(summary.get("andon", {}) or {})
        andon_card = CardFrame()
        andon_card.set_tone("default")
        andon_layout = QGridLayout(andon_card)
        andon_layout.setContentsMargins(12, 12, 12, 12)
        andon_layout.setHorizontalSpacing(10)
        andon_layout.setVerticalSpacing(8)
        for idx, (label_text, key) in enumerate((("Produção", "prod"), ("Setup", "setup"), ("Espera", "espera"), ("Parado", "stop"))):
            box = CardFrame()
            box.set_tone("warning" if key in {"setup", "espera"} else ("danger" if key == "stop" else "success"))
            box_layout = QVBoxLayout(box)
            box_layout.setContentsMargins(10, 10, 10, 10)
            lbl = QLabel(label_text)
            lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #0f172a;")
            val = QLabel(str(int(andon.get(key, 0) or 0)))
            val.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
            box_layout.addWidget(lbl)
            box_layout.addWidget(val)
            andon_layout.addWidget(box, 0, idx)
        layout.addWidget(andon_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        buttons.button(QDialogButtonBox.Close).setText("Fechar")
        layout.addWidget(buttons)
        dialog.exec()
