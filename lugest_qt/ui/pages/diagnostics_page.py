from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..widgets import CardFrame, StatCard


class DiagnosticsPage(QWidget):
    page_title = "Diagnóstico"
    page_subtitle = "Saúde dos dados, stock, gravações e referências críticas."
    uses_backend_reload = True

    AREA_LABELS = {
        "stock": "Stock",
        "retalhos": "Retalhos",
        "notas": "Notas encomenda",
        "plano_orfao": "Planeamento",
        "expedicoes": "Expedição",
        "ref_interna_reutilizada": "Refs. reutilizadas",
        "ref_interna_prefixo_errado": "Prefixos",
        "ref_externa_duplicada": "Refs. externas",
        "ref_interna_duplicada": "Refs. internas",
        "opp_duplicada": "OPP",
        "of_duplicada": "OF",
    }

    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        self._last_report: dict[str, Any] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        header = CardFrame()
        header.set_tone("info")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 12, 16, 12)
        header_layout.setSpacing(10)

        title_col = QVBoxLayout()
        title_col.setContentsMargins(0, 0, 0, 0)
        title_col.setSpacing(2)
        title = QLabel("Painel de diagnóstico")
        title.setStyleSheet("font-size: 19px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel("Valida incoerências que podem afetar rapidez, stock, documentos e fluxo operacional.")
        subtitle.setProperty("role", "muted")
        subtitle.setWordWrap(True)
        self.updated_label = QLabel("Sem leitura carregada.")
        self.updated_label.setStyleSheet("font-size: 11px; color: #17324d;")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        title_col.addWidget(self.updated_label)
        header_layout.addLayout(title_col, 1)

        refresh_btn = QPushButton("Atualizar")
        refresh_btn.setProperty("toolbarAction", "true")
        refresh_btn.setMinimumHeight(34)
        refresh_btn.clicked.connect(self.refresh)
        header_layout.addWidget(refresh_btn)

        self.fix_btn = QPushButton("Corrigir seguro")
        self.fix_btn.setProperty("variant", "success")
        self.fix_btn.setProperty("toolbarAction", "true")
        self.fix_btn.setMinimumHeight(34)
        self.fix_btn.clicked.connect(self._fix_safe)
        header_layout.addWidget(self.fix_btn)
        root.addWidget(header)

        cards_host = QWidget()
        cards_layout = QGridLayout(cards_host)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(10)
        cards_layout.setVerticalSpacing(10)
        self.status_card = StatCard("Estado")
        self.critical_card = StatCard("Críticos")
        self.warning_card = StatCard("Avisos")
        self.runtime_card = StatCard("Gravação")
        self.cards = [self.status_card, self.critical_card, self.warning_card, self.runtime_card]
        for index, card in enumerate(self.cards):
            cards_layout.addWidget(card, 0, index)
        root.addWidget(cards_host)

        overview_card = CardFrame()
        overview_layout = QVBoxLayout(overview_card)
        overview_layout.setContentsMargins(14, 12, 14, 12)
        overview_layout.setSpacing(8)
        overview_title = QLabel("Resumo por área")
        overview_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        overview_layout.addWidget(overview_title)
        self.summary_table = QTableWidget(0, 4)
        self.summary_table.setHorizontalHeaderLabels(["Área", "Estado", "Problemas", "Ação"])
        self.summary_table.verticalHeader().setVisible(False)
        self.summary_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.summary_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.summary_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        overview_layout.addWidget(self.summary_table)
        root.addWidget(overview_card, 2)

        details_card = CardFrame()
        details_layout = QVBoxLayout(details_card)
        details_layout.setContentsMargins(14, 12, 14, 12)
        details_layout.setSpacing(8)
        details_title = QLabel("Detalhe dos alertas")
        details_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #0f172a;")
        details_layout.addWidget(details_title)
        self.details_table = QTableWidget(0, 4)
        self.details_table.setHorizontalHeaderLabels(["Nível", "Área", "Item", "Detalhe"])
        self.details_table.verticalHeader().setVisible(False)
        self.details_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.details_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.details_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.details_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.details_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.details_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        details_layout.addWidget(self.details_table)
        root.addWidget(details_card, 3)

    def refresh(self) -> None:
        getter = getattr(self.backend, "system_diagnostics_report", None)
        if not callable(getter):
            return
        self._last_report = dict(getter() or {})
        self._render()

    def _fix_safe(self) -> None:
        if QMessageBox.question(
            self,
            "Diagnóstico",
            "Aplicar apenas correções seguras?\n\nIsto pode recalcular totais, completar retalhos com origem conhecida e remover planeamento órfão.",
        ) != QMessageBox.Yes:
            return
        fixer = getattr(self.backend, "system_diagnostics_fix_safe", None)
        if not callable(fixer):
            return
        try:
            self._last_report = dict(fixer() or {})
        except Exception as exc:
            QMessageBox.critical(self, "Diagnóstico", str(exc))
            return
        fixes = dict(self._last_report.get("safe_fixes", {}) or {})
        QMessageBox.information(
            self,
            "Diagnóstico",
            (
                "Correção segura concluída.\n\n"
                f"Retalhos atualizados: {int(fixes.get('retalhos_rehidratados', 0) or 0)}\n"
                f"Notas recalculadas: {int(fixes.get('notas_total_recalculado', 0) or 0)}\n"
                f"Planeamento removido: {int(fixes.get('blocos_orfaos_removidos', 0) or 0)}"
            ),
        )
        self._render()

    def _render(self) -> None:
        report = dict(self._last_report or {})
        status = str(report.get("status", "ok") or "ok")
        status_text = {"ok": "Saudável", "warning": "Atenção", "critical": "Crítico"}.get(status, "Atenção")
        tone = {"ok": "success", "warning": "warning", "critical": "danger"}.get(status, "warning")
        generated_at = str(report.get("generated_at", "") or "").replace("T", " ")[:19]
        self.updated_label.setText(f"Última leitura: {generated_at or '-'}")

        counts = dict(report.get("counts", {}) or {})
        runtime = dict(report.get("runtime", {}) or {})
        self.status_card.set_tone(tone)
        self.status_card.set_data(status_text, f"{int(counts.get('encomendas', 0) or 0)} encomendas | {int(counts.get('materiais', 0) or 0)} materiais")
        self.critical_card.set_tone("danger" if int(report.get("critical_count", 0) or 0) else "success")
        self.critical_card.set_data(str(int(report.get("critical_count", 0) or 0)), "stock/gravação")
        self.warning_card.set_tone("warning" if int(report.get("warning_count", 0) or 0) else "success")
        self.warning_card.set_data(str(int(report.get("warning_count", 0) or 0)), "dados a rever")
        pending = bool(runtime.get("pending")) or bool(runtime.get("in_progress"))
        last_error = str(runtime.get("last_error", "") or "").strip()
        self.runtime_card.set_tone("danger" if last_error else ("warning" if pending else "success"))
        self.runtime_card.set_data("Pendente" if pending else "OK", last_error or f"cache {runtime.get('cache_age_sec', 0)}s")

        rows = self._summary_rows(report)
        self.summary_table.setSortingEnabled(False)
        self.summary_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = (row["area"], row["state"], str(row["count"]), row["action"])
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col_index in (1, 2):
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.summary_table.setItem(row_index, col_index, item)
        self.summary_table.setSortingEnabled(True)

        details = self._detail_rows(report)
        self.details_table.setSortingEnabled(False)
        self.details_table.setRowCount(len(details))
        for row_index, row in enumerate(details):
            for col_index, value in enumerate((row["level"], row["area"], row["item"], row["detail"])):
                item = QTableWidgetItem(str(value))
                if col_index == 0:
                    item.setTextAlignment(int(Qt.AlignCenter | Qt.AlignVCenter))
                self.details_table.setItem(row_index, col_index, item)
        self.details_table.setSortingEnabled(True)
        self.fix_btn.setEnabled(any(row["safe"] and row["count"] > 0 for row in rows))

    def _summary_rows(self, report: dict[str, Any]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        safe = dict(report.get("issues_safe", {}) or {})
        risky = dict(report.get("issues_risky", {}) or {})
        actions = {
            "stock": "Corrigir manualmente antes de consumir.",
            "retalhos": "Pode completar origem/peso quando existir base.",
            "notas": "Pode recalcular totais.",
            "plano_orfao": "Pode remover blocos sem encomenda.",
            "expedicoes": "Rever histórico antes de apagar.",
            "ref_interna_reutilizada": "Normal se for o mesmo artigo; rever quando necessário.",
            "ref_interna_prefixo_errado": "Rever cliente/ref. interna.",
            "ref_externa_duplicada": "Confirmar se é repetição legítima.",
            "ref_interna_duplicada": "Rever antes de corrigir.",
            "opp_duplicada": "Rever geração/duplicação.",
            "of_duplicada": "Rever ordens de fabrico.",
        }
        for key, values in safe.items():
            count = len(list(values or []))
            rows.append(
                {
                    "area": self.AREA_LABELS.get(key, key),
                    "state": "OK" if count == 0 else ("Crítico" if key == "stock" else "Aviso"),
                    "count": count,
                    "action": actions.get(key, "Rever."),
                    "safe": key in {"retalhos", "notas", "plano_orfao"},
                }
            )
        for key, values in risky.items():
            count = len(list(values or []))
            rows.append(
                {
                    "area": self.AREA_LABELS.get(key, key),
                    "state": "OK" if count == 0 else "Rever",
                    "count": count,
                    "action": actions.get(key, "Rever manualmente."),
                    "safe": False,
                }
            )
        rows.sort(key=lambda row: (0 if row["count"] else 1, row["area"]))
        return rows

    def _detail_rows(self, report: dict[str, Any]) -> list[dict[str, str]]:
        details: list[dict[str, str]] = []
        for bucket, level in (("issues_safe", "Aviso"), ("issues_risky", "Rever")):
            for key, values in dict(report.get(bucket, {}) or {}).items():
                area = self.AREA_LABELS.get(key, key)
                for value in list(values or [])[:200]:
                    item, detail = self._format_issue(key, value)
                    row_level = "Crítico" if key == "stock" else level
                    details.append({"level": row_level, "area": area, "item": item, "detail": detail})
        runtime = dict(report.get("runtime", {}) or {})
        last_error = str(runtime.get("last_error", "") or "").strip()
        if last_error:
            details.insert(0, {"level": "Crítico", "area": "Gravação", "item": "save", "detail": last_error})
        if not details:
            details.append({"level": "OK", "area": "Sistema", "item": "-", "detail": "Sem alertas ativos."})
        return details

    def _format_issue(self, key: str, value: Any) -> tuple[str, str]:
        if isinstance(value, dict):
            if key == "stock":
                return str(value.get("id", "-") or "-"), f"{value.get('material', '')} {value.get('espessura', '')} | qtd {value.get('quantidade')} | reservado {value.get('reservado')}"
            if key == "retalhos":
                return str(value.get("id", "-") or "-"), f"{value.get('material', '')} {value.get('espessura', '')} | lote {value.get('lote') or '-'} | peso {value.get('peso_unid')}"
            if key == "notas":
                return str(value.get("numero", "-") or "-"), f"guardado {value.get('guardado')} | calculado {value.get('calculado')} | linhas {value.get('linhas')}"
            if key == "plano_orfao":
                return str(value.get("id", "-") or "-"), f"encomenda inexistente {value.get('encomenda', '')}"
            if key == "expedicoes":
                return str(value.get("numero", "-") or "-"), f"encomenda inexistente {value.get('encomenda', '')}"
            if key in {"ref_interna_reutilizada", "ref_interna_duplicada"}:
                return str(value.get("ref_interna", "-") or "-"), f"{value.get('linhas', 0)} linhas | {value.get('encomendas', '')}"
            if key == "ref_externa_duplicada":
                return str(value.get("ref_externa", "-") or "-"), f"{value.get('encomenda', '')} | {value.get('ref_interna', '')}"
            if key == "ref_interna_prefixo_errado":
                return str(value.get("ref_interna", "-") or "-"), f"{value.get('encomenda', '')} | cliente {value.get('cliente', '')}"
            first_key = next(iter(value.keys()), "")
            return str(value.get(first_key, "-") or "-"), " | ".join(f"{k}: {v}" for k, v in value.items())
        return str(value or "-"), ""
