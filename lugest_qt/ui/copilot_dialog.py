from __future__ import annotations

import html
from collections.abc import Callable

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from .theme import polish_widget_tree


class _CallWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, callback: Callable[[], object]) -> None:
        super().__init__()
        self.callback = callback

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(self.callback())
        except Exception as exc:
            self.failed.emit(str(exc))


class ERPCopilotDialog(QDialog):
    navigate_requested = Signal(str)
    idle = Signal()

    PAGE_PROMPTS = {
        "materials": ("Analisar stock crítico", "Que materiais exigem atenção?"),
        "products": ("Analisar stock de produtos", "Que produtos estão sem stock?"),
        "orders": ("Mostrar encomendas atrasadas", "Resume as encomendas ativas"),
        "quotes": ("Resumir orçamentos em aberto", "Que propostas exigem atenção?"),
        "planning": ("Resumir planeamento", "Analisa a carga registada"),
        "transportes": ("Resumir transportes", "Que viagens estão ativas?"),
        "purchase_notes": ("Resumir compras pendentes", "Que notas exigem acompanhamento?"),
        "quality": ("Resumir qualidade", "Que situações continuam abertas?"),
        "avarias": ("Resumir avarias", "Que avarias continuam ativas?"),
    }

    def __init__(self, backend, *, current_page: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.backend = backend
        self.current_page = str(current_page or "")
        self._threads: set[QThread] = set()
        self._workers: set[_CallWorker] = set()
        self._last_target = ""
        self._last_intent = "consult"
        self._pending_action: dict[str, object] = {}
        self._conversation: list[dict[str, str]] = []
        self._pending_question = ""
        self.setWindowTitle("Copiloto LuGEST")
        self.setMinimumSize(780, 620)
        self.resize(900, 700)
        self.setModal(False)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(12)

        header = QHBoxLayout()
        title_col = QVBoxLayout()
        title = QLabel("Copiloto LuGEST")
        title.setStyleSheet("font-size: 22px; font-weight: 900; color: #0f172a;")
        subtitle = QLabel("Conhece os módulos e dados do luGEST, mantém o contexto e prepara ações com segurança.")
        subtitle.setProperty("role", "muted")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header.addLayout(title_col, 1)
        self.status_chip = QLabel("A verificar motor local…")
        self.status_chip.setProperty("role", "state_chip")
        header.addWidget(self.status_chip, 0, Qt.AlignTop)
        root.addLayout(header)

        safety = QFrame()
        safety.setObjectName("Card")
        safety_layout = QHBoxLayout(safety)
        safety_layout.setContentsMargins(12, 8, 12, 8)
        safety_text = QLabel(
            "Consulta segura · sem credenciais ou dados fiscais · alterações exigem revisão e confirmação."
        )
        safety_text.setStyleSheet("color: #3f5f27; font-weight: 700;")
        safety_layout.addWidget(safety_text)
        root.addWidget(safety)

        self.transcript = QTextBrowser()
        self.transcript.setOpenExternalLinks(False)
        self.transcript.setStyleSheet(
            "QTextBrowser { background: #ffffff; border: 1px solid #cfd3cf; "
            "border-radius: 6px; padding: 10px; font-size: 13px; }"
        )
        root.addWidget(self.transcript, 1)
        self._append_assistant(
            "Olá. Conheço os módulos e os fluxos do luGEST. Posso explicar como trabalhar, "
            "analisar os dados permitidos e encaminhar-te para a ação certa. Também mantenho "
            "o contexto, por isso podes fazer perguntas de seguimento."
        )

        prompt_bar = QHBoxLayout()
        prompts = list(self.PAGE_PROMPTS.get(self.current_page, ()))
        prompts.extend(["Analisar stock crítico", "Mostrar encomendas atrasadas", "Resumir planeamento"])
        seen: set[str] = set()
        for prompt in prompts:
            if prompt in seen or len(seen) >= 4:
                continue
            seen.add(prompt)
            button = QPushButton(prompt)
            button.setProperty("variant", "secondary")
            button.setMinimumHeight(34)
            button.clicked.connect(lambda _checked=False, value=prompt: self._use_prompt(value))
            prompt_bar.addWidget(button)
        prompt_bar.addStretch(1)
        root.addLayout(prompt_bar)

        input_row = QHBoxLayout()
        self.question = QPlainTextEdit()
        self.question.setPlaceholderText(
            "Ex.: Quais são as prioridades de hoje?  (Ctrl+Enter para enviar)"
        )
        self.question.setMaximumHeight(82)
        self.question.setMinimumHeight(70)
        input_row.addWidget(self.question, 1)
        self.ask_button = QPushButton("Perguntar")
        self.ask_button.setProperty("variant", "primary")
        self.ask_button.setMinimumSize(128, 70)
        self.ask_button.clicked.connect(self.ask)
        input_row.addWidget(self.ask_button)
        root.addLayout(input_row)

        footer = QHBoxLayout()
        self.engine_label = QLabel("Motor: a verificar")
        self.engine_label.setProperty("role", "muted")
        footer.addWidget(self.engine_label, 1)
        self.open_menu_button = QPushButton("Abrir menu relacionado")
        self.open_menu_button.setProperty("variant", "secondary")
        self.open_menu_button.setVisible(False)
        self.open_menu_button.clicked.connect(self._navigate)
        footer.addWidget(self.open_menu_button)
        close_button = QPushButton("Fechar")
        close_button.setProperty("variant", "secondary")
        close_button.clicked.connect(self.close)
        footer.addWidget(close_button)
        root.addLayout(footer)

        polish_widget_tree(self)
        self._start_call(
            lambda: self.backend.erp_copilot_status(),
            self._status_ready,
            self._status_failed,
        )

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        if (
            event.key() in (Qt.Key_Return, Qt.Key_Enter)
            and bool(event.modifiers() & Qt.ControlModifier)
        ):
            self.ask()
            return
        super().keyPressEvent(event)

    def _start_call(
        self,
        callback: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[str], None],
    ) -> None:
        thread = QThread(self)
        worker = _CallWorker(callback)
        worker.moveToThread(thread)
        self._threads.add(thread)
        self._workers.add(worker)
        thread.started.connect(worker.run)
        worker.finished.connect(on_success)
        worker.finished.connect(lambda _value: self._workers.discard(worker))
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(on_failure)
        worker.failed.connect(lambda _message: self._workers.discard(worker))
        worker.failed.connect(thread.quit)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(lambda: self._thread_finished(thread))
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _thread_finished(self, thread: QThread) -> None:
        self._threads.discard(thread)
        if not self._threads:
            self.idle.emit()

    def _status_ready(self, value: object) -> None:
        status = dict(value or {}) if isinstance(value, dict) else {}
        available = bool(status.get("available", False))
        provider = str(status.get("provider", "") or "").casefold()
        self.status_chip.setText(
            "IA OpenAI disponível"
            if available and provider == "openai"
            else ("IA local disponível" if available else "Modo sem IA")
        )
        self.status_chip.setStyleSheet(
            "border: 1px solid #7ed321; background: #eef8df; color: #31591b; "
            "border-radius: 12px; padding: 6px 12px; font-weight: 800;"
            if available
            else "border: 1px solid #c9cdca; background: #f4f5f3; color: #555b58; "
                 "border-radius: 12px; padding: 6px 12px; font-weight: 800;"
        )
        self.status_chip.setToolTip(str(status.get("message", "") or ""))
        self.engine_label.setText(
            f"Motor: {status.get('model', 'regras locais')}" if available
            else "Motor: regras locais (funcional sem IA)"
        )

    def _status_failed(self, _message: str) -> None:
        self._status_ready({})

    def _use_prompt(self, prompt: str) -> None:
        self.question.setPlainText(prompt)
        self.ask()

    def ask(self) -> None:
        question = self.question.toPlainText().strip()
        if not question or not self.ask_button.isEnabled():
            return
        normalized = " ".join(question.casefold().strip().split())
        if self._pending_action and normalized in {
            "sim", "s", "ok", "okay", "confirma", "confirmar", "pode ser",
            "avanca", "avançar", "faz", "faz la", "faz lá", "segue", "siga",
            "entao", "então", "entao?", "então?", "e entao", "e então",
            "continua", "continuar", "procede", "proceder", "faz isso",
            "podes continuar", "pode continuar", "forca", "força",
        }:
            self._append_user(question)
            self.question.clear()
            self._execute_pending_action()
            return
        if self._pending_action and normalized in {
            "nao", "não", "n", "cancelar", "cancela", "esquece",
        }:
            self._append_user(question)
            self.question.clear()
            self._pending_action = {}
            self._append_assistant("A ação foi cancelada. Nada foi alterado.", title="Ação cancelada")
            self.open_menu_button.setVisible(False)
            return
        if self._pending_action:
            self._append_user(question)
            history = [dict(turn) for turn in self._conversation[-10:]]
            self._conversation.append({"role": "user", "content": question})
            self.question.clear()
            self.ask_button.setEnabled(False)
            self.open_menu_button.setEnabled(False)
            self.ask_button.setText("A interpretar…")
            self.engine_label.setText("Motor: Ollama a interpretar a resposta…")
            action = dict(self._pending_action)
            self._start_call(
                lambda: self.backend.erp_copilot_resolve_action_followup(
                    question,
                    action,
                    history,
                ),
                self._pending_followup_ready,
                self._action_failed,
            )
            return
        self._append_user(question)
        history = [dict(turn) for turn in self._conversation[-10:]]
        self._pending_question = question
        self._conversation.append({"role": "user", "content": question})
        self.question.clear()
        self.ask_button.setEnabled(False)
        self.ask_button.setText("A analisar…")
        self.engine_label.setText("A preparar leitura operacional…")
        self._start_call(
            lambda: self.backend.erp_copilot_ask(question, self.current_page, history),
            self._answer_ready,
            self._answer_failed,
        )

    def _answer_ready(self, value: object) -> None:
        result = dict(value or {}) if isinstance(value, dict) else {}
        answer = str(result.get("answer", "") or "Não foi possível produzir uma leitura.")
        self._append_assistant(answer, title=str(result.get("title", "") or "Copiloto"))
        self._conversation.append({"role": "assistant", "content": answer})
        self._conversation = self._conversation[-12:]
        self._pending_question = ""
        self.engine_label.setText(f"Motor: {result.get('engine', 'regras-locais')}")
        self._last_target = str(result.get("navigation_target", "") or "")
        self._last_intent = str(result.get("intent", "consult") or "consult")
        proposed = result.get("proposed_action", {})
        self._pending_action = dict(proposed) if isinstance(proposed, dict) else {}
        self.open_menu_button.setText(
            "Rever e confirmar"
            if self._pending_action.get("type") == "create_material_stock"
            else "Preparar no menu relacionado"
            if self._last_intent == "prepare_action"
            else "Abrir menu relacionado"
        )
        self.open_menu_button.setVisible(bool(self._last_target))
        self.ask_button.setEnabled(True)
        self.ask_button.setText("Perguntar")

    def _answer_failed(self, message: str) -> None:
        self._append_assistant(
            "Não consegui concluir esta análise. Os restantes menus continuam disponíveis.\n"
            + str(message or "")
        )
        self._pending_question = ""
        self.ask_button.setEnabled(True)
        self.ask_button.setText("Perguntar")
        self.engine_label.setText("Motor: indisponível nesta consulta")

    def _pending_followup_ready(self, value: object) -> None:
        result = dict(value or {}) if isinstance(value, dict) else {}
        decision = str(result.get("decision", "clarify") or "clarify").casefold()
        self.ask_button.setEnabled(True)
        self.open_menu_button.setEnabled(True)
        self.ask_button.setText("Perguntar")
        if decision == "execute":
            self.engine_label.setText("Motor: confirmação compreendida pelo Ollama")
            self._execute_pending_action()
            return
        if decision == "cancel":
            self._pending_action = {}
            self._append_assistant(
                str(result.get("answer", "") or "A ação foi cancelada. Nada foi alterado."),
                title="Ação cancelada",
            )
            self.open_menu_button.setVisible(False)
            self.engine_label.setText("Motor: ação cancelada")
            return
        if decision == "update":
            correction = str(result.get("correction", "") or "").strip()
            if correction:
                original = str(self._pending_action.get("command", "") or "").strip()
                self._pending_action["command"] = (
                    f"{original}. Correção posterior do utilizador: {correction}"
                    if original else correction
                )
            answer = str(
                result.get("answer", "")
                or "A correção foi associada à proposta. Confirma quando quiseres avançar."
            )
            self._append_assistant(answer, title="Proposta atualizada")
            self._conversation.append({"role": "assistant", "content": answer})
            self._conversation = self._conversation[-12:]
            self.engine_label.setText("Motor: Ollama · proposta atualizada")
            return
        answer = str(
            result.get("answer", "")
            or "A ação continua preparada. Indica a correção pretendida ou confirma para avançar."
        )
        self._append_assistant(answer, title="Confirmação necessária")
        self._conversation.append({"role": "assistant", "content": answer})
        self._conversation = self._conversation[-12:]
        self.engine_label.setText("Motor: Ollama · ação ainda pendente")

    def _append_user(self, text: str) -> None:
        self.transcript.append(
            '<div style="margin:10px 0 4px 70px; padding:10px 12px; '
            'background:#eaf7da; border:1px solid #a8d96d; border-radius:6px;">'
            '<b>Tu</b><br>' + html.escape(text).replace("\n", "<br>") + "</div>"
        )

    def _append_assistant(self, text: str, *, title: str = "Copiloto") -> None:
        self.transcript.append(
            '<div style="margin:4px 70px 10px 0; padding:10px 12px; '
            'background:#f4f5f3; border:1px solid #cfd3cf; border-radius:6px;">'
            f"<b>{html.escape(title)}</b><br>"
            + html.escape(text).replace("\n", "<br>")
            + "</div>"
        )
        bar = self.transcript.verticalScrollBar()
        bar.setValue(bar.maximum())

    def _navigate(self) -> None:
        if self._pending_action.get("type") == "create_material_stock":
            self._execute_pending_action()
            return
        if self._last_target:
            self.navigate_requested.emit(self._last_target)

    def _execute_pending_action(self) -> None:
        action = dict(self._pending_action or {})
        action_type = str(action.get("type", "") or "")
        if action_type == "create_material_stock":
            command = str(action.get("command", "") or "").strip()
            if not command:
                self._append_assistant("Falta a descrição do stock a criar.")
                return
            self.ask_button.setEnabled(False)
            self.open_menu_button.setEnabled(False)
            self.engine_label.setText("A interpretar o stock pedido…")
            self._start_call(
                lambda: self.backend.erp_copilot_prepare_action(action),
                self._material_candidate_ready,
                self._action_failed,
            )
            return
        target = str(action.get("target", "") or self._last_target)
        self._pending_action = {}
        if target:
            self.navigate_requested.emit(target)

    def _material_candidate_ready(self, value: object) -> None:
        self.ask_button.setEnabled(True)
        self.open_menu_button.setEnabled(True)
        prepared = dict(value or {}) if isinstance(value, dict) else {}
        candidate = dict(prepared.get("candidate", {}) or {})
        try:
            from .pages.materials_page import _MaterialEditorDialog

            dialog = _MaterialEditorDialog(
                self.backend,
                self,
                record=candidate,
                mode="ai",
            )
            if dialog.exec() != QDialog.Accepted:
                self._pending_action = {}
                self._append_assistant("A proposta foi cancelada. Nada foi alterado.", title="Ação cancelada")
                return
            executed = dict(
                self.backend.erp_copilot_execute_action(
                    self._pending_action,
                    dialog.payload(),
                )
                or {}
            )
            record = dict(executed.get("record", {}) or {})
        except Exception as exc:
            self._action_failed(str(exc))
            return
        self._pending_action = {}
        material_id = str(record.get("id", "") or "")
        self._append_assistant(
            f"Stock criado com sucesso{': ' + material_id if material_id else ''}.",
            title="Ação concluída",
        )
        self.engine_label.setText("Motor: ação luGEST confirmada")
        self.open_menu_button.setText("Abrir Matéria-Prima")
        self.open_menu_button.setVisible(True)
        self._last_target = "materials"
        self._last_intent = "navigate"
        self.navigate_requested.emit("materials")

    def _action_failed(self, message: str) -> None:
        self.ask_button.setEnabled(True)
        self.open_menu_button.setEnabled(True)
        self.engine_label.setText("Motor: não foi possível preparar a ação")
        QMessageBox.warning(
            self,
            "Copiloto luGEST",
            str(message or "Não foi possível preparar esta ação."),
        )
        self._append_assistant(
            str(message or "Não foi possível preparar esta ação."),
            title="Ação não concluída",
        )

    def shutdown(self, timeout_ms: int = 1200) -> None:
        """Finish background workers when the whole application is closing."""

        threads = list(self._threads)
        if not threads:
            return
        share = max(100, int(timeout_ms / max(1, len(threads))))
        for thread in threads:
            if not thread.isRunning():
                continue
            thread.requestInterruption()
            thread.quit()
            if not thread.wait(share):
                # urlopen cannot be interrupted while the socket timeout is active.
                # At application shutdown there is no UI state left to preserve.
                thread.terminate()
                thread.wait(500)
