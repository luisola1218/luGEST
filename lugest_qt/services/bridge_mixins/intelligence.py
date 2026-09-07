from __future__ import annotations

import copy
import os
import re
from lugest_core.intelligence import (
    ERPCopilot as _ERPCopilot,
    MODULE_KNOWLEDGE as _ERP_MODULE_KNOWLEDGE,
    RemoteProductAIClient as _RemoteProductAIClient,
    build_erp_snapshot as _build_erp_snapshot,
    contextual_question as _contextual_copilot_question,
)
from lugest_core.materials import profile_entry as _profile_entry
from lugest_qt.services.bridge_helpers import _detect_profile_catalog_from_text
from typing import Any


class IntelligenceBackendMixin:
    """Legacy adapter for intelligence; see BACKEND_GUIDE.md."""

    def _local_material_command_candidate(self, command: str) -> dict[str, Any]:
        raw = str(command or "").strip()
        text = self._product_catalog_text(raw)
        candidate: dict[str, Any] = {
            "formato": "",
            "material": "",
            "material_familia": "",
            "secao_tipo": "",
            "espessura": "",
            "comprimento": "",
            "largura": "",
            "altura": "",
            "diametro": "",
            "metros": "",
            "kg_m": "",
            "quantidade": "",
            "reservado": "0",
            "local": "",
            "lote_fornecedor": "",
            "p_compra": "",
        }
        format_rules = (
            ("Chapa", r"\bchapas?\b"),
            ("Tubo", r"\btubos?\b"),
            ("Perfil", r"\b(?:perfil|perfis)\b"),
            ("Cantoneira", r"\bcantoneiras?\b"),
            ("Barra", r"\bbarras?\b"),
            ("Varão nervurado", r"\bvar(?:a|ã)o\s+nervurado\b"),
        )
        for label, pattern in format_rules:
            if re.search(pattern, text):
                candidate["formato"] = label
                break
        profile_series, profile_size = _detect_profile_catalog_from_text(text)
        if profile_series:
            candidate["formato"] = "Perfil"
            candidate["secao_tipo"] = profile_series
            candidate["altura"] = profile_size

        dimension_match = re.search(
            r"\b(?:formato|dimens(?:ao|oes))?\s*:?\s*(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*(?:mm)?\b",
            text,
        )
        text_without_dimensions = text
        if dimension_match:
            candidate["comprimento"] = dimension_match.group(1).replace(",", ".")
            candidate["largura"] = dimension_match.group(2).replace(",", ".")
            text_without_dimensions = text[: dimension_match.start()] + " " + text[dimension_match.end() :]

        thickness_match = re.search(
            r"\b(?:espessura\s*(?:de)?\s*)?(\d+(?:[.,]\d+)?)\s*mm\b",
            text_without_dimensions,
        )
        if thickness_match and not profile_series:
            candidate["espessura"] = thickness_match.group(1).replace(",", ".")

        quantity_match = re.search(
            r"\b(\d+(?:[.,]\d+)?)\s*(?:unidades?|unid\.?|uds?|chapas?|barras?|vigas?|perfil|perfis)\b",
            text,
        )
        if quantity_match:
            candidate["quantidade"] = quantity_match.group(1).replace(",", ".")
        length_match = re.search(
            r"\b(?:com\s+)?(\d+(?:[.,]\d+)?)\s*(?:m|metros?)\s*(?:de\s+comprimento)?\b",
            text,
        )
        if length_match and candidate["formato"] in {"Tubo", "Perfil", "Cantoneira", "Barra", "Varão nervurado"}:
            candidate["metros"] = length_match.group(1).replace(",", ".")

        lot_match = re.search(
            r"\blote\s*(?:externo|fornecedor|do\s+fornecedor)?\s*[:#=-]?\s*([a-z0-9][a-z0-9._/-]*)",
            text,
        )
        if lot_match:
            candidate["lote_fornecedor"] = lot_match.group(1).upper()

        location_match = re.search(
            r"\b(?:local|localizacao)\s*[:#=-]?\s*([a-z0-9][a-z0-9._/-]*)",
            text,
        )
        if location_match:
            candidate["local"] = location_match.group(1).upper()

        presets = self.material_presets()
        known_materials = sorted(
            [str(value or "").strip() for value in presets.get("materiais", []) if str(value or "").strip()],
            key=len,
            reverse=True,
        )
        normalized_material_lookup = {
            self._product_catalog_text(value): value
            for value in known_materials
        }
        for normalized, original in normalized_material_lookup.items():
            if normalized and re.search(rf"\b{re.escape(normalized)}\b", text):
                candidate["material"] = original
                break
        if not candidate["material"]:
            grade_match = re.search(
                r"\b(s\d{3,4}j[a-z0-9]*|dc0?1|dx5[1-9]d|aisi\s*\d{3,4}[a-z]?|inox\s*\d{3,4}[a-z]?|aluminio\s*\d{4})\b",
                text,
            )
            if grade_match:
                candidate["material"] = grade_match.group(1).upper().replace(" ", "")
        if not candidate["material"] and candidate["formato"]:
            material_match = re.search(
                rf"\b{re.escape(candidate['formato'].casefold().replace('ã', 'a'))}s?\s+(.+?)\s+\d+(?:[.,]\d+)?\s*mm\b",
                text,
            )
            if material_match:
                candidate["material"] = material_match.group(1).strip().upper()
        if candidate["material"]:
            candidate["material_familia"] = str(
                self.material_family_profile(candidate["material"]).get("key", "") or ""
            ).strip()
        return candidate

    def material_ai_command(self, command: str) -> dict[str, Any]:
        """Prepare, but never persist, a material record from natural language."""

        clean_command = str(command or "").strip()
        if not clean_command:
            raise ValueError("Escreve primeiro o stock de matéria-prima que pretendes criar.")
        local_candidate = self._local_material_command_candidate(clean_command)
        cfg = self._load_qt_config()
        try:
            timeout_seconds = float(
                os.getenv("LUGEST_AI_TIMEOUT_SECONDS")
                or cfg.get("product_ai_timeout_seconds")
                or 45
            )
        except (TypeError, ValueError):
            timeout_seconds = 45.0
        client = _RemoteProductAIClient(
            str(os.getenv("LUGEST_AI_ENDPOINT") or cfg.get("product_ai_endpoint") or "").strip(),
            str(os.getenv("LUGEST_AI_ACCESS_TOKEN") or cfg.get("product_ai_access_token") or "").strip(),
            ollama_url=str(
                os.getenv("LUGEST_OLLAMA_URL")
                or cfg.get("product_ai_ollama_url")
                or "http://127.0.0.1:11434"
            ).strip(),
            ollama_model=str(
                os.getenv("LUGEST_OLLAMA_MODEL")
                or cfg.get("product_ai_ollama_model")
                or "qwen3:4b"
            ).strip(),
            timeout_seconds=timeout_seconds,
        )
        remote_result: dict[str, Any] = {}
        remote_error = ""
        try:
            remote_result = dict(
                client.material_stock_command(
                    clean_command,
                    presets={
                        **self.material_presets(),
                        "familias": self.material_family_options(),
                    },
                    locale="pt-PT",
                )
                or {}
            )
        except Exception as exc:
            remote_error = str(exc)
        candidate = dict(remote_result.get("candidate", {}) or {})
        for key, value in local_candidate.items():
            if str(value or "").strip():
                candidate[key] = value
        detected_series, detected_size = _detect_profile_catalog_from_text(
            " ".join(
                str(value or "")
                for value in (
                    clean_command,
                    candidate.get("secao_tipo"),
                    candidate.get("material"),
                    candidate.get("altura"),
                )
            )
        )
        if detected_series:
            entry = _profile_entry(detected_series, detected_size)
            candidate["formato"] = "Perfil"
            candidate["secao_tipo"] = detected_series
            candidate["altura"] = detected_size
            candidate["espessura"] = ""
            candidate["kg_m"] = self._fmt(entry.get("kg_m", 0))
            metros = self._parse_float(candidate.get("metros", 0), 0)
            if metros > 0:
                candidate["peso_unid"] = self._fmt(float(entry.get("kg_m", 0) or 0) * metros)
            candidate["resumo"] = (
                f"{entry.get('designation', detected_series)} · "
                f"{entry.get('kg_m', 0):g} kg/m · {entry.get('standard', 'EN 10365')}"
            )
        candidate.setdefault("reservado", "0")
        required = ["formato", "material", "quantidade"]
        if str(candidate.get("formato", "") or "").strip().title() == "Chapa":
            required.extend(["espessura", "comprimento", "largura"])
        elif str(candidate.get("formato", "") or "").strip().title() == "Perfil":
            required.extend(["secao_tipo", "altura", "metros"])
        missing = [
            key
            for key in required
            if not str(candidate.get(key, "") or "").strip()
        ]
        candidate["missing_fields"] = missing
        candidate["_ai_summary"] = str(candidate.get("resumo", "") or "").strip()
        candidate["_ai_confidence"] = float(candidate.get("confidence", 0) or 0)
        candidate["_ai_engine"] = str(remote_result.get("engine", "") or "regras-locais")
        candidate["_ai_warning"] = remote_error
        if missing and remote_error:
            raise ValueError(
                remote_error
                + "\n\nNão foi possível completar automaticamente: "
                + ", ".join(missing)
                + "."
            )
        return candidate

    def _erp_copilot_client(self) -> _ERPCopilot:
        cfg = self._load_qt_config()
        try:
            timeout_seconds = float(
                os.getenv("LUGEST_AI_TIMEOUT_SECONDS")
                or cfg.get("product_ai_timeout_seconds")
                or 45
            )
        except (TypeError, ValueError):
            timeout_seconds = 45.0
        return _ERPCopilot(
            ollama_url=str(
                os.getenv("LUGEST_OLLAMA_URL")
                or cfg.get("product_ai_ollama_url")
                or "http://127.0.0.1:11434"
            ).strip(),
            ollama_model=str(
                os.getenv("LUGEST_OLLAMA_MODEL")
                or cfg.get("product_ai_ollama_model")
                or "qwen3:4b"
            ).strip(),
            timeout_seconds=timeout_seconds,
        )

    def erp_copilot_status(self) -> dict[str, Any]:
        """Report optional AI availability without making it a runtime dependency."""

        cfg = self._load_qt_config()
        enabled_text = str(
            os.getenv("LUGEST_AI_ENABLED")
            or cfg.get("erp_copilot_enabled", "1")
            or "1"
        ).strip().casefold()
        enabled = enabled_text not in {"0", "false", "no", "off", "nao", "não"}
        if not enabled:
            return {
                "enabled": False,
                "available": False,
                "service_online": False,
                "message": "Copiloto desativado na configuração deste posto.",
            }
        return {"enabled": True, **self._erp_copilot_client().status()}

    def erp_copilot_snapshot(self) -> dict[str, Any]:
        user = dict(self.user or {})
        permissions = dict(user.get("menu_permissions", {}) or {})
        data = self.ensure_data()
        snapshot = _build_erp_snapshot(data, permissions=permissions)
        snapshot["modulos_lugest"] = {
            key: copy.deepcopy(value)
            for key, value in _ERP_MODULE_KNOWLEDGE.items()
            if not permissions or bool(permissions.get(key, True))
        }
        if not permissions or bool(permissions.get("stock_dashboard", True)):
            try:
                dashboard = dict(self.finance_dashboard("Todos") or {})
                summary = dict(dashboard.get("executive_summary", {}) or {})
                allowed_finance_keys = (
                    "stock_total",
                    "stock_disponivel",
                    "stock_reservado",
                    "stock_materias",
                    "stock_produtos",
                    "compras_total",
                    "compras_materias",
                    "compras_produtos",
                    "compromissos_total",
                    "ne_aprovadas",
                    "vendido_total",
                    "faturado_total",
                    "recebido_total",
                    "saldo_clientes",
                    "referencias_sem_preco",
                    "periodo",
                )
                snapshot["financeiro"] = {
                    key: summary.get(key)
                    for key in allowed_finance_keys
                    if key in summary
                }
            except Exception:
                # The Copilot remains usable if an optional dashboard calculation fails.
                pass
        snapshot["cadastros"] = {
            "clientes": len(list(data.get("clientes", []) or []))
            if not permissions or bool(permissions.get("customers", True))
            else None,
            "fornecedores": len(list(data.get("fornecedores", []) or []))
            if not permissions or bool(permissions.get("suppliers", True))
            else None,
        }
        return snapshot

    def erp_copilot_ask(
        self,
        question: str,
        current_page: str = "",
        conversation: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Answer from a privacy-bounded snapshot; this method never mutates ERP data."""

        cfg = self._load_qt_config()
        enabled_text = str(
            os.getenv("LUGEST_AI_ENABLED")
            or cfg.get("erp_copilot_enabled", "1")
            or "1"
        ).strip().casefold()
        enabled = enabled_text not in {"0", "false", "no", "off", "nao", "não"}
        snapshot = self.erp_copilot_snapshot()
        if enabled:
            return self._erp_copilot_client().ask(
                question,
                snapshot,
                current_page=current_page,
                conversation=conversation,
            )
        from lugest_core.intelligence import deterministic_answer

        result = deterministic_answer(
            _contextual_copilot_question(question, conversation),
            snapshot,
            current_page=current_page,
        )
        result["engine"] = "regras-locais (IA desativada)"
        return result

    def erp_copilot_prepare_action(self, action: dict[str, Any]) -> dict[str, Any]:
        """Prepare an agent action without persisting it and enforce menu permissions."""

        payload = dict(action or {})
        action_type = str(payload.get("type", "") or "").strip()
        target = str(payload.get("target", "") or "").strip()
        allowed_pages = set(self.allowed_pages_for_user())
        if target and target not in allowed_pages:
            raise PermissionError("O utilizador atual não tem permissão para este módulo.")
        if action_type == "create_material_stock":
            command = str(payload.get("command", "") or "").strip()
            return {
                "type": action_type,
                "target": "materials",
                "candidate": self.material_ai_command(command),
            }
        if action_type == "open_workflow":
            return {
                "type": action_type,
                "target": target,
            }
        raise ValueError("Esta ação ainda não possui um executor seguro no Copiloto.")

    def erp_copilot_resolve_action_followup(
        self,
        question: str,
        action: dict[str, Any],
        conversation: list[dict[str, Any]] | None = None,
    ) -> dict[str, str]:
        """Let the local model understand natural confirmation/correction language."""

        return self._erp_copilot_client().resolve_action_followup(
            question,
            dict(action or {}),
            conversation=list(conversation or []),
        )

    def erp_copilot_execute_action(
        self,
        action: dict[str, Any],
        confirmed_payload: dict[str, Any],
    ) -> dict[str, Any]:
        """Execute a previously reviewed action. UI confirmation is mandatory."""

        payload = dict(action or {})
        action_type = str(payload.get("type", "") or "").strip()
        target = str(payload.get("target", "") or "").strip()
        allowed_pages = set(self.allowed_pages_for_user())
        if target and target not in allowed_pages:
            raise PermissionError("O utilizador atual não tem permissão para este módulo.")
        if action_type == "create_material_stock":
            record = dict(self.add_material(dict(confirmed_payload or {})) or {})
            return {
                "ok": True,
                "type": action_type,
                "target": "materials",
                "record_id": str(record.get("id", "") or ""),
                "record": record,
            }
        raise ValueError("Esta ação não pode ser executada automaticamente.")
